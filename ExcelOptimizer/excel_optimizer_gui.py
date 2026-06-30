import os
import zipfile
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

class ExcelOptimizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Ghost Range Optimizer v1.1")
        self.root.geometry("600x480")
        
        style = ttk.Style()
        style.configure("TButton", padding=6)
        style.configure("TLabel", padding=4)
        
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Optimizador de Archivos Excel", font=("Helvetica", 16, "bold")).pack(pady=10)
        ttk.Label(main_frame, text="Elimina filas fantasma y ajusta tablas/filtros.", font=("Helvetica", 10)).pack(pady=5)
        
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=20)
        
        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, state='readonly').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Button(file_frame, text="Buscar Archivo", command=self.browse_file).pack(side=tk.RIGHT)
        
        options_frame = ttk.LabelFrame(main_frame, text="Opciones de Limpieza", padding="10")
        options_frame.pack(fill=tk.X, pady=10)
        
        self.aggro_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Limpieza agresiva (ignorar filas con solo ceros al final)", variable=self.aggro_var).pack(anchor=tk.W)
        
        self.calc_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Eliminar cadena de cálculos (Recalcular al abrir)", variable=self.calc_var).pack(anchor=tk.W)

        self.table_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Ajustar definiciones de Tablas y AutoFiltros", variable=self.table_var).pack(anchor=tk.W)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=20)
        
        self.status_var = tk.StringVar(value="Listo")
        ttk.Label(main_frame, textvariable=self.status_var).pack()
        
        self.run_btn = ttk.Button(main_frame, text="OPTIMIZAR AHORA", command=self.start_optimization)
        self.run_btn.pack(pady=10)
        
        self.log_text = tk.Text(main_frame, height=6, state='disabled', font=("Consolas", 8))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        self.last_detected_row = 1

    def log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"> {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update_idletasks()

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if filename:
            self.file_path_var.set(filename)
            self.log(f"Archivo seleccionado: {os.path.basename(filename)}")

    def start_optimization(self):
        input_file = self.file_path_var.get()
        if not input_file:
            messagebox.showwarning("Advertencia", "Por favor seleccione un archivo .xlsx")
            return
        
        self.run_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.optimize_process, args=(input_file,), daemon=True).start()

    def optimize_process(self, input_file):
        try:
            output_file = input_file.replace(".xlsx", "_FIXED.xlsx")
            self.status_var.set("Optimizando...")
            self.progress_var.set(0)
            
            with zipfile.ZipFile(input_file, 'r') as zin:
                with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zout:
                    items = zin.infolist()
                    total_items = len(items)
                    
                    # Primero procesamos las hojas para saber el límite real
                    for i, item in enumerate(items):
                        if item.filename.startswith('xl/worksheets/sheet') and item.filename.endswith('.xml'):
                            self.log(f"Optimizando hoja: {item.filename}")
                            self.process_sheet(zin, zout, item)
                        elif item.filename.startswith('xl/tables/table') and self.table_var.get():
                            # Procesaremos las tablas al final o después de las hojas
                            continue
                        elif item.filename == 'xl/calcChain.xml' and self.calc_var.get():
                            continue
                        else:
                            zout.writestr(item, zin.read(item.filename))

                    # Ahora procesamos las tablas con el límite encontrado
                    if self.table_var.get():
                        for item in items:
                            if item.filename.startswith('xl/tables/table'):
                                self.log(f"Ajustando tabla: {item.filename}")
                                self.process_table(zin, zout, item)

            self.progress_var.set(100)
            self.status_var.set("¡Completado!")
            self.log(f"Nuevo tamaño: {os.path.getsize(output_file) / 1024:.2f} KB")
            messagebox.showinfo("Éxito", f"Archivo optimizado y corregido para Power Query.\nGuardado como: {os.path.basename(output_file)}")
            
        except Exception as e:
            self.log(f"ERROR: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")
        finally:
            self.run_btn.config(state=tk.NORMAL)

    def process_sheet(self, zin, zout, item):
        with zin.open(item.filename) as f_in:
            content = f_in.read()
            row_pattern = re.compile(br'<row r="(\d+)"')
            val_pattern = re.compile(br'<v>[^0<][^<]*</v>|<t>|<is>') if self.aggro_var.get() else re.compile(br'<v>|<t>|<is>')
            
            row_matches = list(re.finditer(br'<row r="(\d+)"[^>]*>(.*?)</row>', content, re.DOTALL))
            last_row = 1
            for m in reversed(row_matches):
                if val_pattern.search(m.group(2)):
                    last_row = int(m.group(1))
                    break
            
            self.last_detected_row = last_row
            self.log(f"Límite detectado: Fila {last_row}")
            
            header_end = content.find(b'<sheetData>')
            head = content[:header_end].decode('utf-8', errors='ignore')
            head = re.sub(r'dimension ref="([A-Z0-9]+):[A-Z0-9]+"', rf'dimension ref="\1:CY{last_row}"', head)
            
            new_xml = head.encode('utf-8') + b'<sheetData>'
            for m in row_matches:
                if int(m.group(1)) <= last_row:
                    new_xml += m.group(0)
                else: break
            
            footer_start = content.find(b'</sheetData>')
            new_xml += content[footer_start:]
            zout.writestr(item.filename, new_xml)

    def process_table(self, zin, zout, item):
        content = zin.read(item.filename).decode('utf-8', errors='ignore')
        # Ajustar ref="A5:CU1048576" -> ref="A5:CU98"
        content = re.sub(r'ref="([A-Z]+[0-9]+):[A-Z]+[0-9]+"', rf'ref="\1:CU{self.last_detected_row}"', content)
        # Ajustar autoFilter ref
        content = re.sub(r'<autoFilter ref="([A-Z]+[0-9]+):[A-Z]+[0-9]+"', rf'<autoFilter ref="\1:CU{self.last_detected_row}"', content)
        zout.writestr(item.filename, content.encode('utf-8'))

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelOptimizerApp(root)
    root.mainloop()
