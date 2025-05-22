import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import PyPDF2
import re
import os
import win32com.client

class AccountManager(tk.Toplevel):
    def __init__(self, parent, main_app):
        super().__init__(parent)
        self.main_app = main_app
        self.title("Gestión de Cuentas de Correo")
        self.geometry("600x400")
        
        self.accounts = []
        self.create_widgets()
        self.load_accounts()
    
    def create_widgets(self):
        main_frame = ttk.Frame(self)
        main_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Cuentas disponibles en Outlook:").pack(anchor=tk.W)
        
        self.accounts_listbox = tk.Listbox(main_frame, width=70, height=10)
        self.accounts_listbox.pack(fill=tk.BOTH, expand=True)
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Vincular", command=self.link_account).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="Desvincular", command=self.unlink_account).grid(row=0, column=1, padx=5)
        ttk.Button(btn_frame, text="Actualizar", command=self.load_accounts).grid(row=0, column=2, padx=5)
        ttk.Button(btn_frame, text="Cerrar", command=self.destroy).grid(row=0, column=3, padx=5)
        
        self.selected_label = ttk.Label(main_frame, text="Cuenta activa actual: Ninguna")
        self.selected_label.pack(pady=5)
    
    def load_accounts(self):
        try:
            self.accounts_listbox.delete(0, tk.END)
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            self.accounts = namespace.Accounts
            
            for account in self.accounts:
                self.accounts_listbox.insert(tk.END, f"{account.SmtpAddress} - {account.DisplayName}")
            
            if self.main_app.selected_account:
                self.selected_label.config(
                    text=f"Cuenta activa actual: {self.main_app.selected_account.SmtpAddress}"
                )
        
        except Exception as e:
            messagebox.showerror("Error", f"Error cargando cuentas: {str(e)}")
    
    def link_account(self):
        selection = self.accounts_listbox.curselection()
        if selection:
            self.main_app.selected_account = self.accounts[selection[0]]
            self.main_app.update_account_display()
            self.selected_label.config(
                text=f"Cuenta activa actual: {self.main_app.selected_account.SmtpAddress}"
            )
            messagebox.showinfo("Cuenta Vinculada", 
                              f"Cuenta vinculada exitosamente:\n{self.main_app.selected_account.SmtpAddress}")
    
    def unlink_account(self):
        self.main_app.selected_account = None
        self.main_app.update_account_display()
        self.selected_label.config(text="Cuenta activa actual: Ninguna")
        messagebox.showinfo("Desvinculación", "Cuenta desvinculada exitosamente")

class FacturaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Disparador Masivo de Facturas")
        self.root.geometry("1100x700")
        
        
        self.pdf_paths = []
        self.excel_path = tk.StringVar()
        self.selected_account = None
        
        self.create_widgets()
        self.create_footer() 
        self.update_account_display()
    
    def create_widgets(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        # Sección de cuentas
        account_frame = ttk.LabelFrame(main_frame, text="Cuenta de Envío")
        account_frame.pack(fill=tk.X, pady=5)
        
        self.account_status = ttk.Label(account_frame, text="Cuenta actual: Ninguna seleccionada")
        self.account_status.pack(side=tk.LEFT, padx=5)
        
        btn_container = ttk.Frame(account_frame)
        btn_container.pack(side=tk.RIGHT)
        
        ttk.Button(btn_container, text="Gestionar Cuentas", command=self.manage_accounts).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_container, text="Desvincular", command=self.unlink_current_account).pack(side=tk.LEFT, padx=2)
        
        # Sección PDFs
        pdf_frame = ttk.LabelFrame(main_frame, text="Documentos PDF")
        pdf_frame.pack(fill=tk.BOTH, pady=5, expand=True)
        
        self.pdf_listbox = tk.Listbox(pdf_frame, width=100, height=10)
        self.pdf_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        btn_frame = ttk.Frame(pdf_frame)
        btn_frame.pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(btn_frame, text="Agregar PDFs", command=self.select_pdfs).pack(pady=2)
        ttk.Button(btn_frame, text="Limpiar Lista", command=self.clear_pdfs).pack(pady=2)
        ttk.Button(btn_frame, text="Eliminar Selección", command=self.remove_selected_pdf).pack(pady=2)
        
        # Sección Excel
        excel_frame = ttk.Frame(main_frame)
        excel_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(excel_frame, text="Archivo Excel:").pack(side=tk.LEFT)
        ttk.Entry(excel_frame, textvariable=self.excel_path, width=80).pack(side=tk.LEFT, padx=5)
        ttk.Button(excel_frame, text="Examinar", command=self.select_excel).pack(side=tk.LEFT)
        
        # Botón de proceso
        ttk.Button(main_frame, text="INICIAR ENVÍO MASIVO", command=self.process_all).pack(pady=15)
        
        # Log
        log_frame = ttk.LabelFrame(main_frame, text="Registro de Actividad")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log = tk.Text(log_frame, height=15, width=120, wrap=tk.WORD)
        self.log.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)
        
        # Estilos
        style = ttk.Style()
        style.configure("Accent.TButton", foreground="white", background="#0078D4", font=('Helvetica', 10, 'bold'))

    def create_footer(self):
        # Crear el frame del footer
        footer_frame = tk.Frame(self.root, bg="#f0f0f0")
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        
        # Texto del desarrollador
        texto_developer = tk.Label(
            footer_frame, 
            text="Developed by Tobias Gallo", 
            bg="#f0f0f0", 
            fg="#666666",
            font=('Helvetica', 9)
        )
        texto_developer.pack(side=tk.LEFT, padx=10)
    
    def manage_accounts(self):
        AccountManager(self.root, self)
    
    def unlink_current_account(self):
        self.selected_account = None
        self.update_account_display()
        messagebox.showinfo("Cuenta Desvinculada", "Cuenta desvinculada exitosamente")
    
    def update_account_display(self):
        if self.selected_account:
            text = f"Cuenta actual: {self.selected_account.SmtpAddress}"
        else:
            text = "Cuenta actual: Ninguna seleccionada"
        self.account_status.config(text=text)
    
    def select_pdfs(self):
        paths = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if paths:
            self.pdf_paths.extend(paths)
            self.update_pdf_listbox()
    
    def clear_pdfs(self):
        self.pdf_paths = []
        self.update_pdf_listbox()
    
    def remove_selected_pdf(self):
        selection = self.pdf_listbox.curselection()
        if selection:
            del self.pdf_paths[selection[0]]
            self.update_pdf_listbox()
    
    def update_pdf_listbox(self):
        self.pdf_listbox.delete(0, tk.END)
        for path in self.pdf_paths:
            self.pdf_listbox.insert(tk.END, os.path.basename(path))
    
    def select_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.excel_path.set(path)
    
    def extract_data(self, pdf_path):
        try:
            with open(pdf_path, "rb") as f:
                reader = PyPDF2.PdfReader(f)
                text = ""
                for page in reader.pages:
                    text += page.extract_text()
                
                # Patrones de búsqueda actualizados según el código simple
                dni = re.search(r'DNI:\s*([\d\.]+)', text)
                nombre = re.search(r'Sr\. \(es\):\s*(.+?)\n', text)
                
                return (
                    dni.group(1).replace(".", "") if dni else None,
                    nombre.group(1).strip() if nombre else None
                )
        except Exception as e:
            self.log_insert(f"Error leyendo PDF: {str(e)}")
            return None, None
    
    def find_email(self, dni, nombre):
        try:
            df = pd.read_excel(self.excel_path.get())
            cleaned_dni = dni.replace(".", "")
            
            # Búsqueda actualizada según el código simple
            # Primero por CUIT/DNI
            mask_cuit = df['CUIT'].astype(str).str.replace(".", "") == cleaned_dni
            if not df[mask_cuit].empty:
                return df[mask_cuit].iloc[0]['Email']
            
            # Luego por nombre
            mask_nombre = df['Nombre'].str.strip().str.upper() == nombre.upper()
            if not df[mask_nombre].empty:
                return df[mask_nombre].iloc[0]['Email']
            
            return None
        except Exception as e:
            self.log_insert(f"Error en Excel: {str(e)}")
            return None
    
    def create_email(self, to_email, pdf_path):
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            
            if self.selected_account:
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, self.selected_account))
                cuenta = self.selected_account.SmtpAddress
            else:
                cuenta = "Cuenta predeterminada"
            
            mail.To = to_email
            mail.Subject = "Factura adjunta - Por favor revisar"
            mail.Body = "Estimado cliente,\n\nAdjuntamos su factura correspondiente.\n\nSaludos cordiales"
            mail.Attachments.Add(pdf_path)
            mail.Send()
            
            return cuenta
        except Exception as e:
            self.log_insert(f"Error enviando correo: {str(e)}")
            return None
    
    def log_insert(self, message):
        self.log.insert(tk.END, message + "\n")
        self.log.see(tk.END)
        self.root.update_idletasks()
    
    def process_all(self):
        self.log.delete(1.0, tk.END)
        try:
            if not self.pdf_paths:
                raise ValueError("Debe seleccionar al menos un PDF")
            if not self.excel_path.get():
                raise ValueError("Debe seleccionar el archivo Excel")
            if not self.selected_account:
                raise ValueError("Debe seleccionar una cuenta de correo")
            
            total = len(self.pdf_paths)
            exitosos = 0
            
            for idx, pdf_path in enumerate(self.pdf_paths, 1):
                try:
                    self.log_insert(f"\n[{idx}/{total}] Procesando: {os.path.basename(pdf_path)}")
                    
                    dni, nombre = self.extract_data(pdf_path)
                    if not dni or not nombre:
                        raise ValueError("Datos incompletos en el PDF")
                    
                    self.log_insert(f"Datos encontrados:\n- DNI: {dni}\n- Nombre: {nombre}")
                    
                    email = self.find_email(dni, nombre)
                    if not email:
                        raise ValueError("Email no encontrado en Excel")
                    
                    self.log_insert(f"Email encontrado: {email}")
                    
                    cuenta_usada = self.create_email(email, pdf_path)
                    if cuenta_usada:
                        exitosos += 1
                        self.log_insert(f"✅ Enviado exitosamente a: {email}")
                        self.log_insert(f"   Cuenta utilizada: {cuenta_usada}")
                    else:
                        self.log_insert("❌ Falló el envío")
                
                except Exception as e:
                    self.log_insert(f"⚠️ Error: {str(e)}")
            
            resumen = (
                f"\n{'='*50}\n"
                f"PROCESO COMPLETADO\n"
                f"Total archivos procesados: {total}\n"
                f"Envíos exitosos: {exitosos}\n"
                f"Envíos fallidos: {total - exitosos}\n"
                f"Cuenta utilizada: {self.selected_account.SmtpAddress}\n"
                f"{'='*50}"
            )
            self.log_insert(resumen)
            messagebox.showinfo("Resumen Final", resumen)
        
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = FacturaApp(root)
    root.mainloop()