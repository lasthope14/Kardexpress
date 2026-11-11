# -*- coding: utf-8 -*-
"""
Reconstrucción del código fuente de la versión 1 de Kardexpress (kardex.exe, ago-2023).

Obtenido a partir del ejecutable PyInstaller (Python 3.11):
  - REA/Ghidra identificó el bootloader de PyInstaller en kardex.exe
    (evidencia ev_243c191c57f5b3f26ad3ae509534306ea22da9825c4ebb554b039c8f674327a8).
  - Se extrajo el script principal 'kardex' del archivo CArchive y se desensambló
    el bytecode 3.11 con xdis. Los números de línea coinciden con el original.

Es una reconstrucción manual a partir del bytecode: los nombres, constantes y el
flujo de control son exactos; los comentarios son añadidos. Se conserva SOLO como
referencia de comportamiento — la versión nueva vive en el paquete `kardexpress`.
"""
import openpyxl
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
from tkinter import PhotoImage
from PIL import Image, ImageTk
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import threading
import os


def load_file_1():                                                   # línea 12
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    entry_path_1.delete(0, tk.END)
    entry_path_1.insert(0, file_path)


def load_file_2():                                                   # línea 17
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    entry_path_2.delete(0, tk.END)
    entry_path_2.insert(0, file_path)


def process_data():                                                  # línea 22
    run_button.config(state=tk.DISABLED)
    progress_bar.config(mode="determinate", maximum=100)
    progress_bar.start(10)                     # animación falsa, no es progreso real
    thread = threading.Thread(target=run_script)
    thread.start()


def update_progress(value):                                          # línea 30 (nunca se usa)
    progress_bar.stop()
    progress_bar["value"] = value


def stop_progress():                                                 # línea 34
    progress_bar.stop()
    progress_bar["value"] = 0
    progress_bar.config(mode="indeterminate")
    run_button.config(state=tk.NORMAL)
    messagebox.showinfo("Proceso Finalizado", "El proceso ha finalizado con éxito.")


def show_error_message_and_stop_progress(message):                   # línea 41 (nunca se usa)
    progress_bar.stop()
    progress_bar["value"] = 0
    progress_bar.config(mode="indeterminate")
    run_button.config(state=tk.NORMAL)
    messagebox.showerror("Error", message)


def check_process_finished():                                        # línea 48
    if not process_finished_event.is_set():
        root.after(100, check_process_finished)
        return
    stop_progress()                            # deja de sondear para siempre tras la 1ª ejecución


def run_script():                                                    # línea 55 (hilo secundario)
    file_path_1 = entry_path_1.get()           # acceso a Tk desde otro hilo
    file_path_2 = entry_path_2.get()

    if file_path_1 and file_path_2:            # si falta uno: retorna y la UI queda bloqueada
        # --- 1) Archivo de REQUISICIONES: col B = "REQUI-xxxx", fecha en fila+1, col W
        workbook = openpyxl.load_workbook(file_path_1)
        sheet = workbook.active

        result_data = []
        row_index = 1
        while row_index <= sheet.max_row:
            cell_value = sheet.cell(row=row_index, column=2).value
            if cell_value and cell_value.startswith("REQUI-") and len(cell_value) == 10:
                req_number = cell_value
                creation_date = sheet.cell(row=row_index + 1, column=23).value
                if creation_date:
                    creation_date = creation_date.strftime("%m/%d/%Y")
                    result_data.append((req_number, creation_date))
                row_index += 2
            else:
                row_index += 1

        result_data_dict = {req: date for req, date in result_data}  # línea 80

        # --- 2) Archivo KARDEX: col E = requisición -> escribir fecha en col S
        other_workbook = openpyxl.load_workbook(file_path_2)
        other_sheet = other_workbook.active
        lastRow = other_sheet.max_row

        for i in range(2, lastRow + 1):
            req_number = other_sheet.cell(row=i, column=5).value
            creation_date = result_data_dict.get(req_number)
            if creation_date:
                other_sheet.cell(row=i, column=19).value = creation_date

        # Rellenar hacia abajo col S (arrastra la fecha de la fila anterior,
        # incluso a través de órdenes distintas)
        for i in range(3, lastRow + 1):
            if not other_sheet.cell(row=i, column=19).value:
                other_sheet.cell(row=i, column=19).value = other_sheet.cell(row=i - 1, column=19).value

        # Filas de encabezado de cada grupo
        for i in range(1, lastRow + 1):
            cell_value = other_sheet.cell(row=i, column=2).value
            if cell_value and cell_value.startswith("Consecutivo General"):
                other_sheet.cell(row=i, column=19).value = None
                continue
            if cell_value and cell_value == "Código":
                cell = other_sheet.cell(row=i, column=19)
                cell.value = "Fecha"
                cell.font = Font(name="Calibri", size=11, bold=True)

        # --- 3) Formato de la columna S desde la fila 7
        font = Font(name="Courier New", size=8)
        alignment = Alignment(horizontal="center", vertical="center")
        border = Border(left=Side(border_style="thin", color="000000"),
                        right=Side(border_style="thin", color="000000"),
                        top=Side(border_style="thin", color="000000"),
                        bottom=Side(border_style="thin", color="000000"))
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

        for i in range(7, lastRow + 1):
            cell = other_sheet.cell(row=i, column=19)
            cell.font = font
            cell.alignment = alignment
            cell.border = border

        # --- 4) Semáforo por fila desde la fila 8 (M=13 comprada, N=14 entrada, R=18 pendiente)
        for i in range(8, lastRow + 1):                     # AMARILLO: sin orden de compra
            cell_value = other_sheet.cell(row=i, column=13).value
            if cell_value == 0:
                for j in range(1, other_sheet.max_column + 1):
                    other_sheet.cell(row=i, column=j).fill = yellow_fill

        for i in range(8, lastRow + 1):                     # VERDE: sin entrada/remisión
            cant_comprada = other_sheet.cell(row=i, column=13).value
            cant_entrada = other_sheet.cell(row=i, column=14).value
            if cant_comprada != 0 and cant_entrada == 0:
                for j in range(1, other_sheet.max_column + 1):
                    other_sheet.cell(row=i, column=j).fill = green_fill

        for i in range(8, lastRow + 1):                     # ROJO: pendiente proveedor
            cant_entrada = other_sheet.cell(row=i, column=14).value
            pendiente = other_sheet.cell(row=i, column=18).value
            if (cant_entrada != 0 and cant_entrada is not None and not isinstance(cant_entrada, str)
                    and pendiente != 0 and pendiente is not None and not isinstance(pendiente, str)):
                for j in range(1, other_sheet.max_column + 1):
                    other_sheet.cell(row=i, column=j).fill = red_fill

        # --- 5) Leyenda en T3:T5
        other_sheet.cell(row=3, column=20).value = "SIN ORDEN DE COMPRA"
        other_sheet.cell(row=4, column=20).value = "SIN ENTRADA/REMISION"
        other_sheet.cell(row=5, column=20).value = "PENDIENTE PROVEEDOR"
        for i in range(3, 6):
            cell = other_sheet.cell(row=i, column=20)
            cell.font = Font(name="Calibri", size=11, bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = yellow_fill if i == 3 else green_fill if i == 4 else red_fill
        other_sheet.column_dimensions["T"].width = 25

        other_workbook.save(file_path_2)       # SOBRESCRIBE el kardex original, sin respaldo
        other_workbook.close()
        process_finished_event.set()           # si hubo excepción antes, nunca se llega aquí


ruta_actual = os.path.dirname(__file__)
ruta_imagen_fondo = os.path.join(ruta_actual, "assets", "kardex.png")

root = tk.Tk()
root.title("Kardexpress")
root.geometry("1000x800")
root.resizable(True, False)

bg_image = PhotoImage(file=ruta_imagen_fondo)
bg_label = tk.Label(root, image=bg_image)
bg_label.place(relx=0.5, rely=0, anchor="n")

style = ttk.Style()
style.configure("TButton", font=("Roboto", 14, "bold"), padding=10)

element_width = 20

entry_path_1 = tk.Entry(root, font=("Roboto", 12), width=63)
load_button_1 = ttk.Button(root, text="Cargar", command=load_file_1, width=element_width)
entry_path_2 = tk.Entry(root, font=("Roboto", 12), width=56)
load_button_2 = ttk.Button(root, text="Cargar", command=load_file_2, width=element_width)

entry_path_1.place(relx=0.5, rely=0.47, anchor="center")
load_button_1.place(relx=0.5, rely=0.54, anchor="center")
entry_path_2.place(relx=0.5, rely=0.74, anchor="center")
load_button_2.place(relx=0.5, rely=0.81, anchor="center")

run_button = ttk.Button(root, text="Realizar Kardex", command=process_data, width=element_width)
run_button.place(relx=0.5, rely=0.89, anchor="center")

progress_bar = ttk.Progressbar(root, mode="determinate", length=700)
progress_bar.place(relx=0.5, rely=0.95, anchor="center")

process_finished_event = threading.Event()

root.after(100, check_process_finished)
root.mainloop()
