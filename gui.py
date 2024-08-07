import tkinter as tk
from tkinter import filedialog
from functions import load_excel
import pandas as pd
import re

class DuplicatesRemoverApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Eliminador de Duplicados en Excel")
        self.root.geometry("900x800")

        self.filename = None
        self.output_folder = None
        self.columns = []
        self.verify_selection = []
        self.keep_selection = []

        # Crear marcos (frames) para la disposición
        self.create_frames()

        # Componentes de la columna izquierda superior
        self.create_left_column()

        # Componentes de la columna derecha superior
        self.create_right_column()

        # Componentes de la columna inferior
        self.create_bottom_column()

    def create_frames(self):
        self.left_frame = tk.Frame(self.root)
        self.left_frame.grid(row=0, column=0, padx=10, pady=10, sticky="n")

        self.right_frame = tk.Frame(self.root)
        self.right_frame.grid(row=0, column=1, padx=10, pady=10, sticky="n")

        self.bottom_frame = tk.Frame(self.root)
        self.bottom_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

    def create_left_column(self):
        tk.Label(self.left_frame, text="Selecciona el archivo EXCEL que quieras limpiar de emails duplicados", font=("Helvetica", 12, "bold"), fg="#FCB12D").grid(row=0, column=0, sticky="w")

        self.file_button = tk.Button(self.left_frame, text="Selecciona el archivo", command=self.select_file)
        self.file_button.grid(row=1, column=0, pady=5, sticky="w")

        self.filepath_label = tk.Label(self.left_frame, text="Ruta del archivo seleccionado: Ninguno")
        self.filepath_label.grid(row=2, column=0, sticky="w")

        tk.Label(self.left_frame, text="Nombre del archivo limpio de duplicados", font=("Helvetica", 12, "bold"), fg="#FCB12D").grid(row=3, column=0, pady=15, sticky="w")

        self.output_entry = tk.Entry(self.left_frame, width=30)
        self.output_entry.grid(row=4, column=0, pady=5, sticky="w")

        self.folder_button = tk.Button(self.left_frame, text="Selecciona la carpeta para guardar el archivo generado", command=self.select_folder)
        self.folder_button.grid(row=5, column=0, pady=15, sticky="w")

        self.folderpath_label = tk.Label(self.left_frame, text="Ruta de la carpeta seleccionada: Ninguna")
        self.folderpath_label.grid(row=6, column=0, sticky="w")

    def create_right_column(self):
        tk.Label(self.right_frame, text="Selecciona la columna para verificar los duplicados", font=("Helvetica", 12, "bold"), fg="#FCB12D").grid(row=0, column=0, sticky="w")

        self.verify_listbox = tk.Listbox(self.right_frame, selectmode=tk.MULTIPLE, exportselection=0)
        self.verify_listbox.grid(row=1, column=0, pady=5, sticky="w")

        tk.Label(self.right_frame, text="Selecciona las columnas que quieres mantener", font=("Helvetica", 12, "bold"), fg="#FCB12D").grid(row=2, column=0, pady=15, sticky="w")

        self.keep_listbox = tk.Listbox(self.right_frame, selectmode=tk.MULTIPLE, exportselection=0)
        self.keep_listbox.grid(row=3, column=0, pady=5, sticky="w")

    def create_bottom_column(self):
        self.status_label = tk.Label(self.bottom_frame, text="")
        self.status_label.grid(row=0, column=0, sticky="w")

        self.process_button = tk.Button(self.bottom_frame, text="Eliminar Duplicados", command=self.process_file)
        self.process_button.grid(row=1, column=0, pady=10, sticky="w")

        self.result_text = tk.Text(self.bottom_frame, height=10, width=110)
        self.result_text.grid(row=2, column=0, pady=10, sticky="w")

        # Footer
        self.footer_label = tk.Label(self.bottom_frame, text="Programa realizado por Francesc Xavier Escandell | Escandell.cat", fg="#FCB12D")
        self.footer_label.grid(row=3, column=0, pady=5, sticky="ew")
        self.footer_label.bind("<Button-1>", lambda e: self.open_url("https://escandell.cat"))

    def select_file(self):
        try:
            self.filename = filedialog.askopenfilename()
            if self.filename:
                df = load_excel(self.filename)
                if df is not None:
                    self.columns = df.columns.tolist()
                    print(f"Columnas encontradas: {self.columns}")
                    self.update_listboxes()
                    self.filepath_label.config(text=f"Ruta del archivo seleccionado:\n{self.filename}")
                else:
                    print("Error al cargar el archivo. Asegúrate de que el archivo tenga datos y esté en el formato correcto.")
            else:
                print("No se seleccionó ningún archivo.")
        except Exception as e:
            print(f"Error al seleccionar el archivo: {e}")

    def update_listboxes(self):
        self.verify_listbox.delete(0, tk.END)
        self.keep_listbox.delete(0, tk.END)
        for column in self.columns:
            self.verify_listbox.insert(tk.END, column)
            self.keep_listbox.insert(tk.END, column)

        for index in self.verify_selection:
            self.verify_listbox.select_set(index)
        for index in self.keep_selection:
            self.keep_listbox.select_set(index)

    def select_folder(self):
        try:
            self.output_folder = filedialog.askdirectory()
            self.folderpath_label.config(text=f"Ruta de la carpeta seleccionada:\n{self.output_folder}")
        except Exception as e:
            print(f"Error al seleccionar la carpeta: {e}")

    def process_file(self):
        if not self.filename or not self.output_folder:
            self.status_label.config(text="Error: Archivo de entrada o carpeta de salida no seleccionados.")
            return

        output_filename = self.output_entry.get()
        if not output_filename:
            self.status_label.config(text="Error: Nombre de archivo de salida no proporcionado.")
            return

        self.verify_selection = self.verify_listbox.curselection()
        self.keep_selection = self.keep_listbox.curselection()
        selected_verify_columns = [self.verify_listbox.get(i) for i in self.verify_selection]
        selected_keep_columns = [self.keep_listbox.get(i) for i in self.keep_selection]

        if not selected_verify_columns:
            self.status_label.config(text="Error: Seleccione al menos una columna para verificación.")
            return

        if not selected_keep_columns:
            self.status_label.config(text="Error: Seleccione al menos una columna para mantener.")
            return

        self.status_label.config(text="En proceso...")
        self.root.update_idletasks()

        df = load_excel(self.filename)
        if df is not None:
            try:
                data = []
                duplicates = set()
                invalid_emails = set()
                unique_emails_set = set()

                email_pattern = re.compile(r'^[\w\.-]+@[\w\.-]+\.\w+$')

                for _, row in df.iterrows():
                    emails = set()
                    for col in selected_verify_columns:
                        if pd.notna(row[col]):
                            if email_pattern.match(row[col]):
                                if row[col] in unique_emails_set:
                                    duplicates.add(row[col])
                                else:
                                    unique_emails_set.add(row[col])
                                    emails.add(row[col])
                            else:
                                invalid_emails.add(row[col])

                    for email in emails:
                        row_data = {col: row[col] for col in selected_keep_columns}
                        row_data['email'] = email
                        data.append(row_data)

                df_unique_emails = pd.DataFrame(data)

                df_unique_emails.drop_duplicates(subset=['email'], inplace=True)

                output_path = f"{self.output_folder}/{output_filename}.xlsx"
                df_unique_emails.to_excel(output_path, index=False)
                self.status_label.config(text="Completado")

                duplicates_count = len(duplicates)
                invalid_emails_count = len(invalid_emails)
                total_valid_emails = df_unique_emails.shape[0]

                self.result_text.delete(1.0, tk.END)
                self.result_text.insert(tk.END, f"Emails duplicados encontrados y eliminados: {duplicates_count}\n")
                self.result_text.insert(tk.END, f"Emails con formato incorrecto eliminados: {invalid_emails_count}\n")
                self.result_text.insert(tk.END, f"Total de emails válidos: {total_valid_emails}\n")
                self.result_text.insert(tk.END, "\nEmails duplicados:\n")
                self.result_text.insert(tk.END, ", ".join(duplicates) + "\n")
                self.result_text.insert(tk.END, "\nEmails con formato incorrecto:\n")
                self.result_text.insert(tk.END, ", ".join(invalid_emails) + "\n")

            except Exception as e:
                self.status_label.config(text="Error: No se pudo guardar el archivo.")
        else:
            self.status_label.config(text="Error al procesar el archivo.")

    def open_url(self, url):
        import webbrowser
        webbrowser.open_new(url)

if __name__ == "__main__":
    root = tk.Tk()
    app = DuplicatesRemoverApp(root)
    root.mainloop()
