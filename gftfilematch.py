import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import inch
from datetime import datetime
import time
import os
import subprocess
import pandas as pd

# Variables globales para almacenar la información necesaria
file_1_lines = []
file_2_lines = []
file_1_content = ""
file_2_content = ""
file_1_name = ""
file_2_name = ""
comparison_status = ""
comparison_time = 0
differences = None
difference_count = 0

def select_file_1():
    file_path = filedialog.askopenfilename()
    if file_path:
        entry_file_1.config(state=ttk.NORMAL)
        entry_file_1.delete(0, ttk.END)
        entry_file_1.insert(0, file_path)
        entry_file_1.config(state=ttk.DISABLED)

def select_file_2():
    file_path = filedialog.askopenfilename()
    if file_path:
        entry_file_2.config(state=ttk.NORMAL)
        entry_file_2.delete(0, ttk.END)
        entry_file_2.insert(0, file_path)
        entry_file_2.config(state=ttk.DISABLED)

def get_file_details(file_path):
    try:
        file_size = os.path.getsize(file_path)
        creation_time = datetime.fromtimestamp(os.path.getctime(file_path)).strftime('%Y-%m-%d %H:%M:%S')
        modification_time = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S')
        return file_size, creation_time, modification_time
    except Exception as e:
        messagebox.showerror("Error", f"Could not retrieve properties for file {file_path}: {e}")
        return None, None, None

def read_file_content(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            return file.read()
    except UnicodeDecodeError:
        try:
            with open(file_path, 'r', encoding='ISO-8859-1') as file:
                return file.read()
        except Exception as e:
            messagebox.showerror("Error", f"Could not read file {file_path}: {e}")
            return None

def compare_files():
    global comparison_result, differences, comparison_time, comparison_status, file_1_name, file_2_name, file_1_lines, file_2_lines, file_1_content, file_2_content, difference_count
    file_1_path = entry_file_1.get()
    file_2_path = entry_file_2.get()

    if not file_1_path or not file_2_path:
        messagebox.showerror("Error", "Both files must be selected.")
        return

    file_1_name = os.path.basename(file_1_path)
    file_2_name = os.path.basename(file_2_path)

    start_time = time.time()

    file_1_content = read_file_content(file_1_path)
    file_2_content = read_file_content(file_2_path)

    if file_1_content is None or file_2_content is None:
        return

    file_1_size, file_1_creation, file_1_modification = get_file_details(file_1_path)
    file_2_size, file_2_creation, file_2_modification = get_file_details(file_2_path)

    try:
        file_1_lines = file_1_content.splitlines()
        file_2_lines = file_2_content.splitlines()

        comparison_time = round(time.time() - start_time, 2)

        if file_1_content == file_2_content:
            comparison_status = "Successful"
            differences = None
            difference_count = 0
        else:
            comparison_status = "Failed"
            differences = compare_lines(file_1_lines, file_2_lines)
            difference_count = len(differences)

        details_file_1 = (f"{file_1_name}:\n"
                          f"Number of lines: {len(file_1_lines)}\n"
                          f"Total length: {len(file_1_content)} characters\n"
                          f"Size: {file_1_size} bytes\n"
                          f"Creation date: {file_1_creation}\n"
                          f"Last modification date: {file_1_modification}")

        details_file_2 = (f"{file_2_name}:\n"
                          f"Number of lines: {len(file_2_lines)}\n"
                          f"Total length: {len(file_2_content)} characters\n"
                          f"Size: {file_2_size} bytes\n"
                          f"Creation date: {file_2_creation}\n"
                          f"Last modification date: {file_2_modification}")

        comparison_info = f"Processing time: {comparison_time} seconds"
        differences_info = f"Lines with differences: {difference_count}"

        comparison_result = f"Comparison Status: {comparison_status}\n\n"
        comparison_result += f"{details_file_1}\n\n"
        comparison_result += f"{details_file_2}\n\n"
        comparison_result += f"{comparison_info}\n\n"
        comparison_result += f"{differences_info}\n\n"

        if differences:
            max_differences = int(entry_max_differences.get())
            comparison_result += f"Comparison details (differences only, showing first {max_differences}):\n"
            comparison_result += "\n".join([f"Line {diff[0]}:\n\t{file_1_name}: {diff[1]}\n\t{file_2_name}: {diff[2]}" for diff in differences[:max_differences]])

        text_comparison_result.config(state=ttk.NORMAL)
        text_comparison_result.delete(1.0, ttk.END)
        text_comparison_result.insert(ttk.END, comparison_result)
        text_comparison_result.config(state=ttk.DISABLED)

        button_export_pdf.config(state=ttk.NORMAL)
        button_export_excel.config(state=ttk.NORMAL)

    except Exception as e:
        messagebox.showerror("Error", f"Could not compare the files: {e}")

def compare_lines(lines1, lines2):
    differences = []
    max_len = max(len(lines1), len(lines2))
    for i in range(max_len):
        line1 = lines1[i] if i < len(lines1) else ""
        line2 = lines2[i] if i < len(lines2) else ""
        
        if line1.strip() == "" and line2.strip() == "":
            continue
        
        if line1 != line2:
            differences.append((i + 1, line1, line2))
    return differences

def add_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont('Helvetica', 10)
    canvas.drawString(inch, 0.75 * inch, "GFT FileMatch")
    canvas.drawString(6.5 * inch, 0.75 * inch, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    canvas.restoreState()

def export_to_pdf():
    global differences, difference_count
    if not comparison_result:
        messagebox.showerror("Error", "There are no comparison results to export.")
        return

    max_differences = int(entry_max_differences.get())
    if difference_count > max_differences:
        messagebox.showwarning("Warning", f"Only the first {max_differences} differences will be exported.")

    file_path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile="GFT FileMatch", filetypes=[("PDF files", "*.pdf")])
    if not file_path:
        return

    try:
        pdf = SimpleDocTemplate(file_path, pagesize=letter)
        elements = []

        styles = getSampleStyleSheet()
        title_style = styles['Title']
        title_style.alignment = TA_CENTER

        title = Paragraph("GFT FileMatch", title_style)
        elements.append(title)
        elements.append(Spacer(1, 12))

        elements.append(Paragraph(f"Comparison Status: {comparison_status}", styles['Normal']))
        elements.append(Spacer(1, 12))

        details_file_1 = (f"{file_1_name}:\n"
                          f"Number of lines: {len(file_1_lines)}\n"
                          f"Total length: {len(file_1_content)} characters\n"
                          f"Size: {os.path.getsize(entry_file_1.get())} bytes\n"
                          f"Creation date: {datetime.fromtimestamp(os.path.getctime(entry_file_1.get())).strftime('%Y-%m-%d %H:%M:%S')}\n"
                          f"Last modification date: {datetime.fromtimestamp(os.path.getmtime(entry_file_1.get())).strftime('%Y-%m-%d %H:%M:%S')}")
        elements.append(Paragraph(details_file_1, styles['Normal']))
        elements.append(Spacer(1, 12))

        details_file_2 = (f"{file_2_name}:\n"
                          f"Number of lines: {len(file_2_lines)}\n"
                          f"Total length: {len(file_2_content)} characters\n"
                          f"Size: {os.path.getsize(entry_file_2.get())} bytes\n"
                          f"Creation date: {datetime.fromtimestamp(os.path.getctime(entry_file_2.get())).strftime('%Y-%m-%d %H:%M:%S')}\n"
                          f"Last modification date: {datetime.fromtimestamp(os.path.getmtime(entry_file_2.get())).strftime('%Y-%m-%d %H:%M:%S')}")
        elements.append(Paragraph(details_file_2, styles['Normal']))
        elements.append(Spacer(1, 12))

        elements.append(Paragraph(f"Processing time: {comparison_time} seconds", styles['Normal']))
        elements.append(Spacer(1, 12))

        elements.append(Paragraph(f"Lines with differences: {difference_count}", styles['Normal']))
        elements.append(Spacer(1, 12))

        if differences:
            elements.append(Paragraph(f"Comparison details (differences only, showing first {max_differences}):", styles['Normal']))
            elements.append(Spacer(1, 12))

            table_data = [["Line", file_1_name, file_2_name]]
            for diff in differences[:max_differences]:  # Limitar a max_differences
                table_data.append([str(diff[0]), Paragraph(diff[1], styles['BodyText']), Paragraph(diff[2], styles['BodyText'])])

            table = Table(table_data, colWidths=[50, 250, 250])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            elements.append(table)

        pdf.build(elements, onFirstPage=add_footer, onLaterPages=add_footer)

        messagebox.showinfo("Success", "The results have been successfully exported to PDF.")
        subprocess.Popen([file_path], shell=True)

    except Exception as e:
        messagebox.showerror("Error", f"Could not export the results: {e}")

def export_to_excel():
    global differences, difference_count
    if not comparison_result:
        messagebox.showerror("Error", "There are no comparison results to export.")
        return

    max_differences = int(entry_max_differences.get())
    if difference_count > max_differences:
        messagebox.showwarning("Warning", f"Only the first {max_differences} differences will be exported.")

    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="GFT FileMatch", filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return

    try:
        data = {
            "File 1": [file_1_name],
            "File 2": [file_2_name],
            "Comparison Status": [comparison_status],
            "Processing Time (s)": [comparison_time],
            "Lines with Differences": [difference_count]
        }

        df = pd.DataFrame(data)

        if differences:
            diff_data = {
                "Line Number": [diff[0] for diff in differences[:max_differences]],  # Limitar a max_differences
                f"{file_1_name}": [diff[1] for diff in differences[:max_differences]],
                f"{file_2_name}": [diff[2] for diff in differences[:max_differences]]
            }
            diff_df = pd.DataFrame(diff_data)
            with pd.ExcelWriter(file_path) as writer:
                df.to_excel(writer, sheet_name="Summary", index=False)
                diff_df.to_excel(writer, sheet_name="Differences", index=False)
        else:
            df.to_excel(file_path, index=False)

        messagebox.showinfo("Success", "The results have been successfully exported to Excel.")
        subprocess.Popen([file_path], shell=True)

    except Exception as e:
        messagebox.showerror("Error", f"Could not export the results to Excel: {e}")

# Crear la ventana principal usando ttkbootstrap con un tema moderno
root = ttk.Window(themename="superhero")  # Cambiado el tema a 'superhero'
root.title("GFT FileMatch")

# Ajustar el tamaño de la ventana para adaptarse a la pantalla
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = int(screen_width * 0.8)
window_height = int(screen_height * 0.8)

root.geometry(f"{window_width}x{window_height}")
root.resizable(True, True)

# Espacio superior añadido
ttk.Label(root, text=" ").grid(row=0, column=0, pady=10)

# Crear entradas y botones
file_frame_1 = ttk.Frame(root)
file_frame_1.grid(row=1, column=0, columnspan=3, padx=10, pady=5)

label_file_1 = ttk.Label(file_frame_1, text="File 1:")
label_file_1.pack(side=LEFT, padx=(0, 5))

entry_file_1 = ttk.Entry(file_frame_1, width=80, state=ttk.DISABLED)
entry_file_1.pack(side=LEFT, padx=(0, 5))

button_browse_1 = ttk.Button(file_frame_1, text="Browse", command=select_file_1, bootstyle="primary-outline")
button_browse_1.pack(side=LEFT)

file_frame_2 = ttk.Frame(root)
file_frame_2.grid(row=2, column=0, columnspan=3, padx=10, pady=5)

label_file_2 = ttk.Label(file_frame_2, text="File 2:")
label_file_2.pack(side=LEFT, padx=(0, 5))

entry_file_2 = ttk.Entry(file_frame_2, width=80, state=ttk.DISABLED)
entry_file_2.pack(side=LEFT, padx=(0, 5))

button_browse_2 = ttk.Button(file_frame_2, text="Browse", command=select_file_2, bootstyle="primary-outline")
button_browse_2.pack(side=LEFT)

button_compare = ttk.Button(root, text="Compare", command=compare_files, bootstyle="success-outline")
button_compare.grid(row=3, column=0, columnspan=3, pady=10)

# Crear un frame para el número máximo de diferencias a exportar y los botones de exportación
export_frame = ttk.Frame(root)
export_frame.grid(row=4, column=0, columnspan=3, pady=10)

label_max_differences = ttk.Label(export_frame, text="Maximum number of records to export:")
label_max_differences.pack(side=LEFT, padx=(0, 10))

entry_max_differences = ttk.Entry(export_frame, width=10)
entry_max_differences.insert(0, "10000")  # Valor predeterminado
entry_max_differences.pack(side=LEFT, padx=(0, 20))

button_export_pdf = ttk.Button(export_frame, text="Export to PDF", command=export_to_pdf, state=DISABLED, bootstyle="info-outline")
button_export_pdf.pack(side=LEFT, padx=(10, 10))

button_export_excel = ttk.Button(export_frame, text="Export to Excel", command=export_to_excel, state=DISABLED, bootstyle="info-outline")
button_export_excel.pack(side=LEFT)

# Agregar cuadro de texto con barra de desplazamiento para mostrar el resultado de la comparación
label_comparison_result = ttk.Label(root, text="Comparison Result:")
label_comparison_result.grid(row=5, column=0, padx=10, pady=5, sticky="w")

frame_comparison_result = ttk.Frame(root)
frame_comparison_result.grid(row=6, column=0, columnspan=3, padx=10, pady=5, sticky="nsew")

scrollbar = ttk.Scrollbar(frame_comparison_result, bootstyle="round")
scrollbar.pack(side=RIGHT, fill=Y)

text_comparison_result = ttk.Text(frame_comparison_result, width=110, height=20, state=DISABLED, yscrollcommand=scrollbar.set)
text_comparison_result.pack(side=LEFT, fill=BOTH, expand=True)

scrollbar.config(command=text_comparison_result.yview)

# Añadir un label en la parte inferior derecha
label_credit = ttk.Label(root, text="Developed with Python by Jaime Londoño - Property of GFT", font=("Arial", 8))
label_credit.grid(row=7, column=2, padx=10, pady=10, sticky="se")

# Asegurar que la ventana principal se ajuste a los cambios
root.grid_rowconfigure(6, weight=1)
root.grid_columnconfigure(1, weight=1)

# Ejecutar el bucle principal de Tkinter
root.mainloop()