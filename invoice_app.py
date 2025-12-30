import os
import re
import sys
import time
import fitz  # PyMuPDF
import pdfrw
import json
import urllib.request
import urllib.parse
import ssl
from datetime import datetime
from pdf2image import convert_from_path
from fpdf import FPDF
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit,
    QTextEdit, QListWidget, QProgressBar, QFileDialog, QInputDialog, QMessageBox,
    QLabel, QCheckBox
)
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtCore import QUrl

# Configurar SSL para evitar problemas de certificados en macOS
try:
    ssl._create_default_https_context = ssl._create_unverified_context
except:
    pass

# Configuración de email
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587
EMAIL_ADDRESS = "fernando@boogiemanmedia.com"
RECIPIENT_EMAIL = "streamingap@netflix.com"

# Contraseña en texto plano SOLO para test. Borra esto para producción y usá el .env.
EMAIL_PASSWORD = "Fr33d0mF|ghter@.,."  # ← Cambiá o borrá esta línea para producción

# Configuración del servidor remoto de invoice numbers
REMOTE_COUNTER_URL = "https://boogiemanmedia.com/Invoice_Number/invoice_counter.php"
REMOTE_COUNTER_TOKEN = "36c0718096aea52c66d0935dcaddd997e37fafb5a24ec7da91aa7c0a5942c1c0"
REMOTE_TIMEOUT = 15  # segundos

script_dir = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(script_dir, "INVOICE_BASE_DynamicConvert_FormCamposOK.pdf")
LAST_INVOICE_FILE = os.path.expanduser("~/last_invoice_backup.txt")  # Backup local por si falla la conexión

def _remote_call(params):
    """Realiza llamadas al servidor remoto del contador de facturas."""
    try:
        params['token'] = REMOTE_COUNTER_TOKEN
        
        # Construir URL con parámetros GET
        query_string = urllib.parse.urlencode(params)
        url = f"{REMOTE_COUNTER_URL}?{query_string}"
        
        # Debug: mostrar qué se está enviando
        print(f"DEBUG: URL completa: {url}")
        
        # Crear request con headers apropiados
        req = urllib.request.Request(url)
        req.add_header('User-Agent', 'Mozilla/5.0 (compatible; InvoiceApp/1.0)')
        
        # Realizar la petición GET
        with urllib.request.urlopen(req, timeout=REMOTE_TIMEOUT) as resp:
            response_text = resp.read().decode('utf-8')
            print(f"DEBUG: Response raw: {response_text}")
            
            # Intentar parsear JSON
            try:
                response_json = json.loads(response_text)
                print(f"DEBUG: Response JSON: {response_json}")
                return response_json
            except json.JSONDecodeError as je:
                print(f"ERROR: No se pudo parsear JSON: {je}")
                print(f"ERROR: Respuesta recibida: {response_text[:500]}")
                
                # Si la respuesta es solo un número, intentar procesarlo
                if response_text.strip().isdigit():
                    return {"ok": True, "last": int(response_text.strip())}
                    
                return None
                
    except urllib.error.HTTPError as e:
        print(f"ERROR HTTP {e.code}: {e.reason}")
        try:
            error_body = e.read().decode('utf-8')
            print(f"ERROR Body: {error_body}")
        except:
            pass
        return None
    except urllib.error.URLError as e:
        print(f"ERROR URL: {e.reason}")
        return None
    except Exception as e:
        print(f"ERROR General: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        return None

def get_last_invoice_number():
    """Obtiene el último número de invoice desde el servidor web."""
    # Intentar obtener desde el servidor remoto
    try:
        res = _remote_call({"action": "current"})
        if res and res.get("ok") and "last" in res:
            remote_num = int(res["last"])
            print(f"Número de invoice obtenido del servidor: {remote_num}")
            # Guardar backup local
            save_last_invoice_number_local(remote_num)
            return remote_num
    except Exception as e:
        print(f"Error obteniendo número desde el servidor: {e}")
    
    # Si falla el servidor, usar backup local
    print("Usando backup local del último número de invoice")
    return get_last_invoice_number_local()

def get_last_invoice_number_local():
    """Obtiene el último número de invoice desde archivo local (backup)."""
    try:
        with open(LAST_INVOICE_FILE, "r") as f:
            return int(f.read().strip())
    except Exception:
        return 1

def save_last_invoice_number_local(num):
    """Guarda el último número de invoice en archivo local (backup)."""
    try:
        with open(LAST_INVOICE_FILE, "w") as f:
            f.write(str(num))
    except Exception:
        pass

def save_last_invoice_number(num):
    """Actualiza el número de invoice en el servidor web y localmente."""
    try:
        print(f"📤 Actualizando servidor con número: {num}")
        
        # Usar action=set que es lo que funciona en tu servidor
        params = {"action": "set", "value": str(num)}
        params['token'] = REMOTE_COUNTER_TOKEN
        
        # Construir URL con parámetros GET
        query_string = urllib.parse.urlencode(params)
        url = f"{REMOTE_COUNTER_URL}?{query_string}"
        
        print(f"DEBUG: URL de actualización: {url[:100]}...")
        
        req = urllib.request.Request(url)
        req.add_header('User-Agent', 'Mozilla/5.0 (compatible; InvoiceApp/1.0)')
        
        with urllib.request.urlopen(req, timeout=REMOTE_TIMEOUT) as resp:
            response_text = resp.read().decode('utf-8')
            print(f"DEBUG: Respuesta de actualización: {response_text}")
            
            try:
                res = json.loads(response_text)
                if res and res.get("ok"):
                    print(f"✅ Número de invoice actualizado en el servidor: {num}")
                    if "last" in res:
                        print(f"   Confirmado: Servidor ahora en #{res['last']}")
                else:
                    print(f"⚠️ Error del servidor: {res}")
            except json.JSONDecodeError:
                print(f"⚠️ Respuesta inesperada: {response_text}")
                    
    except urllib.error.HTTPError as e:
        print(f"❌ Error HTTP actualizando: {e.code} - {e.reason}")
        try:
            error_body = e.read().decode('utf-8')
            print(f"   Detalles: {error_body}")
        except:
            pass
    except Exception as e:
        print(f"❌ Error actualizando número en el servidor: {e}")
    
    # Siempre guardar backup local
    save_last_invoice_number_local(num)
    print(f"💾 Backup local guardado: {num}")

def reserve_next_invoice_remote():
    """Reserva el siguiente número de invoice en el servidor (incrementa y retorna)."""
    try:
        print("📤 Reservando siguiente número en el servidor...")
        
        # Primero obtener el número actual
        current = get_remote_current()
        if current:
            next_num = current + 1
            
            # Usar action=set para establecer el nuevo número
            params = {"action": "set", "value": str(next_num)}
            params['token'] = REMOTE_COUNTER_TOKEN
            
            query_string = urllib.parse.urlencode(params)
            url = f"{REMOTE_COUNTER_URL}?{query_string}"
            
            req = urllib.request.Request(url)
            req.add_header('User-Agent', 'Mozilla/5.0 (compatible; InvoiceApp/1.0)')
            
            with urllib.request.urlopen(req, timeout=REMOTE_TIMEOUT) as resp:
                response_text = resp.read().decode('utf-8')
                print(f"DEBUG: Respuesta de reserva: {response_text}")
                
                try:
                    res = json.loads(response_text)
                    if res and res.get("ok"):
                        print(f"✅ Número reservado: {next_num}")
                        save_last_invoice_number_local(next_num)
                        return next_num
                except json.JSONDecodeError:
                    if response_text.strip().isdigit():
                        reserved = int(response_text.strip())
                        print(f"✅ Número reservado (directo): {reserved}")
                        save_last_invoice_number_local(reserved)
                        return reserved
                    
    except Exception as e:
        print(f"❌ Error reservando número: {e}")
    
    # Si falla, usar método local
    local_num = get_last_invoice_number_local() + 1
    save_last_invoice_number_local(local_num)
    print(f"⚠️ Usando numeración local: {local_num}")
    return local_num

def _remote_call_post(params):
    """Realiza llamadas POST al servidor (para acciones que modifican estado)."""
    try:
        params['token'] = REMOTE_COUNTER_TOKEN
        data = urllib.parse.urlencode(params).encode('utf-8')
        
        req = urllib.request.Request(REMOTE_COUNTER_URL, data=data)
        req.add_header('Content-Type', 'application/x-www-form-urlencoded')
        req.add_header('User-Agent', 'Mozilla/5.0 (compatible; InvoiceApp/1.0)')
        
        with urllib.request.urlopen(req, timeout=REMOTE_TIMEOUT) as resp:
            response_text = resp.read().decode('utf-8')
            
            try:
                return json.loads(response_text)
            except json.JSONDecodeError:
                if response_text.strip().isdigit():
                    return {"ok": True, "reserved": int(response_text.strip())}
                return None
                
    except Exception as e:
        print(f"Error en POST: {e}")
        return None

def get_remote_current():
    """Obtiene el número actual del servidor sin incrementar."""
    try:
        res = _remote_call({"action": "current"})
        if res and res.get("ok") and "last" in res:
            return int(res["last"])
    except Exception:
        pass
    return None

def extract_information_from_first_page(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        page = doc.load_page(0)
        text = page.get_text()
        doc.close()
        
        grand_arcade_match = re.search(r"GrandArcade: (.+)", text)
        grand_arcade = grand_arcade_match.group(1).strip() if grand_arcade_match else ""
        
        po_date_match = re.search(r"Purchase Order Date\s*(.+)", text)
        purchase_order_date = po_date_match.group(1).strip() if po_date_match else ""
        
        total_po_amount_match = re.search(r"USD\s*([0-9,]+\.[0-9]{2})", text)
        total_po_amount = "$" + total_po_amount_match.group(1).strip() if total_po_amount_match else ""
        
        po_number_match = re.search(r"Purchase Order Number\s*(PO-\d+)", text)
        purchase_order_number = po_number_match.group(1).strip() if po_number_match else ""
        
        return {
            "GrandArcade": grand_arcade,
            "Purchase Order Date": purchase_order_date,
            "Total PO Amount": total_po_amount,
            "Purchase Order Number": purchase_order_number,
        }
    except Exception as e:
        print(f"Error processing {pdf_path}: {e}")
        return {}

def fill_pdf(template_path, output_path, data, invoice_number):
    template_pdf = pdfrw.PdfReader(template_path)
    for page in template_pdf.pages:
        annots = page['/Annots']
        if annots is None:
            continue
        for annotation in annots:
            if annotation['/Subtype'] == '/Widget' and annotation['/T']:
                field_name = annotation['/T'][1:-1]
                if field_name == 'Detail':
                    annotation.update(pdfrw.PdfDict(V=data.get("GrandArcade", "")))
                elif field_name in ['UnitPrice', 'TotalPrice', 'SubTotal', 'Total']:
                    annotation.update(pdfrw.PdfDict(V=data.get("Total PO Amount", "")))
                elif field_name == 'Date':
                    annotation.update(pdfrw.PdfDict(V=datetime.now().strftime('%m/%d/%Y')))
                elif field_name == 'InvoiceNumber':
                    annotation.update(pdfrw.PdfDict(V=f'Invoice #{invoice_number}'))
                elif field_name == 'PONumber':
                    annotation.update(pdfrw.PdfDict(V=data.get("Purchase Order Number", "")))
                elif field_name == 'Cantidad':
                    annotation.update(pdfrw.PdfDict(V='1'))
    pdfrw.PdfWriter(output_path, trailer=template_pdf).write()

def convert_pdf_to_images(pdf_path, output_folder, dpi=300):
    images = convert_from_path(pdf_path, dpi=dpi)
    paths = []
    for i, image in enumerate(images):
        image_name = f"{os.path.basename(pdf_path).replace('.pdf','')}_page_{i+1}.png"
        image_path = os.path.join(output_folder, image_name)
        image.save(image_path, "PNG")
        paths.append(image_path)
    return paths

def create_pdf_from_images(image_paths, final_pdf_path):
    pdf = FPDF()
    for image_path in image_paths:
        pdf.add_page()
        pdf.image(image_path, x=0, y=0, w=210, h=297)
    pdf.output(final_pdf_path, "F")

def remove_temp_files(paths):
    for p in paths:
        if os.path.exists(p):
            os.remove(p)

def export_to_excel(data_list, output_file):
    try:
        df = pd.DataFrame(data_list)
        df.to_excel(output_file, index=False)
    except Exception:
        pass

def process_pdf_files(input_folder, output_folder, invoice_start, log=None):
    extracted = []
    excel_data = []
    
    # Buscar archivos PDF
    pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]
    if not pdf_files:
        if log:
            log("❌ No se encontraron archivos PDF en la carpeta seleccionada")
        return []
    
    if log:
        log(f"📋 Encontrados {len(pdf_files)} archivos PDF")
    
    for name in pdf_files:
        path = os.path.join(input_folder, name)
        info = extract_information_from_first_page(path)
        if info and info.get("Purchase Order Number"):
            extracted.append((info, name))
        elif log:
            log(f"⚠️ No se pudo extraer información de: {name}")
    
    if not extracted:
        if log:
            log("❌ No se pudo extraer información de ningún archivo PDF")
        return []
    
    # Ordenar por número de PO y fecha
    def po_num(po):
        m = re.search(r"PO-(\d+)", po)
        return int(m.group(1)) if m else float('inf')
    
    try:
        extracted.sort(key=lambda x: (po_num(x[0]["Purchase Order Number"]),
                                     datetime.strptime(x[0]["Purchase Order Date"], "%m/%d/%Y")))
    except:
        # Si hay error en las fechas, ordenar solo por PO
        extracted.sort(key=lambda x: po_num(x[0]["Purchase Order Number"]))
    
    invoice_number = invoice_start
    generated = []
    
    for info, original_name in extracted:
        out_name = f"Invoice_{invoice_number}_Netflix_{info['Purchase Order Number']}.pdf"
        out_path = os.path.join(output_folder, out_name)
        
        if log:
            log(f"Generating {out_name}")
        
        try:
            fill_pdf(TEMPLATE_PATH, out_path, info, invoice_number)
            imgs = convert_pdf_to_images(out_path, output_folder)
            create_pdf_from_images(imgs, out_path)
            remove_temp_files(imgs)
            
            info_excel = info.copy()
            info_excel["Invoice Number"] = invoice_number
            excel_data.append(info_excel)
            generated.append(out_path)
            
        except Exception as e:
            if log:
                log(f"❌ Error generando {out_name}: {e}")
        
        invoice_number += 1
    
    # Actualizar el último número en el servidor
    save_last_invoice_number(invoice_number - 1)
    
    # Exportar a Excel
    if excel_data:
        excel_path = os.path.join(output_folder, "PO_Data.xlsx")
        export_to_excel(excel_data, excel_path)
        if log:
            log(f"📊 Datos exportados a: {os.path.basename(excel_path)}")
    
    return generated

def send_email(subject, body, attachment_path, recipient_email, log=None):
    try:
        msg = MIMEMultipart()
        msg["From"] = EMAIL_ADDRESS
        msg["To"] = recipient_email
        msg["Subject"] = subject

        msg.attach(MIMEText(body, "plain"))

        with open(attachment_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename={os.path.basename(attachment_path)}",
        )
        msg.attach(part)

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.set_debuglevel(0)
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.sendmail(EMAIL_ADDRESS, recipient_email, msg.as_string())

        return True

    except smtplib.SMTPAuthenticationError as e:
        if log:
            log(f"❌ Error de autenticación SMTP: {e}")
        return False
    except Exception as e:
        if log:
            log(f"❌ Error enviando email: {e}")
        return False

class InvoiceApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PO2Invoice - Generador y Enviador de Facturas")
        self.setGeometry(100, 100, 800, 600)
        self.invoices = []
        self.default_po_folder = "/Users/fernandorajlevsky/Desktop/PAGAR/PO-FACTURAR"
        self.setup_ui()
        self.check_server_connection()
        self.load_default_folder()

    def check_server_connection(self):
        """Verifica la conexión con el servidor al iniciar."""
        self.log("🔍 Verificando conexión con el servidor...")
        
        # Primero intentar una conexión simple de prueba
        try:
            import ssl
            import socket
            
            # Verificar si podemos alcanzar el servidor
            socket.create_connection(("boogiemanmedia.com", 443), timeout=5)
            self.log("✅ Servidor alcanzable")
        except Exception as e:
            self.log(f"⚠️ No se puede alcanzar el servidor: {e}")
        
        # Ahora intentar obtener el número
        remote_num = get_remote_current()
        if remote_num is not None:
            self.log(f"✅ Conectado al servidor. Último invoice: #{remote_num}")
            self.connection_label.setText(f"🔡 Conectado - Último invoice: #{remote_num}")
            self.connection_label.setStyleSheet("color: green; margin: 5px;")
        else:
            self.log("⚠️ No se pudo conectar al servidor. Usando modo offline.")
            self.connection_label.setText("🔡 Sin conexión - Modo offline")
            self.connection_label.setStyleSheet("color: orange; margin: 5px;")

    def setup_ui(self):
        layout = QVBoxLayout()
        
        # Título
        title = QLabel("🧾 Generador de Facturas desde Purchase Orders")
        title.setStyleSheet("font-size: 16px; font-weight: bold; margin: 10px;")
        layout.addWidget(title)
        
        # Indicador de conexión
        self.connection_label = QLabel("🔡 Estado del servidor: Verificando...")
        self.connection_label.setStyleSheet("color: gray; margin: 5px;")
        layout.addWidget(self.connection_label)
        
        # Carpeta de entrada (POs)
        inp_layout = QHBoxLayout()
        inp_label = QLabel("📁 Carpeta POs:")
        inp_label.setMinimumWidth(120)
        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText(f"Predeterminada: {self.default_po_folder}")
        inp_btn = QPushButton("📂 Seleccionar POs")
        inp_btn.clicked.connect(self.select_input)
        inp_layout.addWidget(inp_label)
        inp_layout.addWidget(self.input_edit)
        inp_layout.addWidget(inp_btn)
        layout.addLayout(inp_layout)
        
        # Carpeta de salida (Facturas)
        out_layout = QHBoxLayout()
        out_label = QLabel("💾 Carpeta Salida:")
        out_label.setMinimumWidth(120)
        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText("(Opcional) Carpeta donde guardar facturas...")
        out_btn = QPushButton("📂 Seleccionar Salida")
        out_btn.clicked.connect(self.select_output)
        out_layout.addWidget(out_label)
        out_layout.addWidget(self.output_edit)
        out_layout.addWidget(out_btn)
        layout.addLayout(out_layout)
        
        # Email de destino
        email_layout = QHBoxLayout()
        email_label = QLabel("📧 Email Destino:")
        email_label.setMinimumWidth(120)
        self.email_edit = QLineEdit(RECIPIENT_EMAIL)
        email_layout.addWidget(email_label)
        email_layout.addWidget(self.email_edit)
        layout.addLayout(email_layout)
        
        # Checkbox para envío automático
        self.auto_send_checkbox = QCheckBox("📤 Enviar emails automáticamente después de generar")
        self.auto_send_checkbox.setChecked(False)
        layout.addWidget(self.auto_send_checkbox)
        
        # Botones principales
        btn_layout = QHBoxLayout()
        self.gen_btn = QPushButton("🔧 Generar Facturas")
        self.gen_btn.clicked.connect(self.generate)
        self.gen_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; font-weight: bold; padding: 8px; }")
        
        # NUEVO: Botón para cargar facturas existentes
        self.load_btn = QPushButton("📥 Cargar Facturas")
        self.load_btn.clicked.connect(self.load_existing_invoices)
        self.load_btn.setStyleSheet("QPushButton { background-color: #FF9800; color: white; font-weight: bold; padding: 8px; }")
        self.load_btn.setToolTip("Cargar facturas ya generadas desde una carpeta para enviar por email")
        
        self.send_btn = QPushButton("📤 Enviar Emails")
        self.send_btn.clicked.connect(self.send_emails)
        self.send_btn.setEnabled(False)
        self.send_btn.setStyleSheet("QPushButton { background-color: #2196F3; color: white; font-weight: bold; padding: 8px; }")
        
        self.open_btn = QPushButton("📂 Abrir Carpeta")
        self.open_btn.clicked.connect(self.open_output_folder)
        self.open_btn.setEnabled(False)
        
        self.sync_btn = QPushButton("🔄 Sincronizar")
        self.sync_btn.clicked.connect(self.sync_with_server)
        self.sync_btn.setToolTip("Sincronizar con el servidor de números de invoice")
        
        btn_layout.addWidget(self.gen_btn)
        btn_layout.addWidget(self.load_btn)  # NUEVO
        btn_layout.addWidget(self.send_btn)
        btn_layout.addWidget(self.open_btn)
        btn_layout.addWidget(self.sync_btn)
        layout.addLayout(btn_layout)
        
        # Lista de facturas generadas
        list_label = QLabel("📋 Facturas Cargadas:")
        layout.addWidget(list_label)
        self.list_widget = QListWidget()
        self.list_widget.itemDoubleClicked.connect(self.preview_invoice)
        layout.addWidget(self.list_widget)
        
        # Barra de progreso
        self.progress = QProgressBar()
        layout.addWidget(self.progress)
        
        # Log de actividad
        log_label = QLabel("📝 Log de Actividad:")
        layout.addWidget(log_label)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        layout.addWidget(self.log_text)
        
        self.setLayout(layout)

    def log(self, text):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {text}")
        QApplication.processEvents()

    def sync_with_server(self):
        """Sincroniza con el servidor y actualiza el estado."""
        self.log("🔄 Sincronizando con el servidor...")
        remote_num = get_remote_current()
        if remote_num is not None:
            self.connection_label.setText(f"🔡 Conectado - Último invoice: #{remote_num}")
            self.connection_label.setStyleSheet("color: green; margin: 5px;")
            self.log(f"✅ Sincronizado. Último invoice en servidor: #{remote_num}")
        else:
            self.connection_label.setText("🔡 Sin conexión - Modo offline")
            self.connection_label.setStyleSheet("color: orange; margin: 5px;")
            self.log("⚠️ No se pudo conectar al servidor")

    def load_default_folder(self):
        """Carga la carpeta predeterminada si existe."""
        if os.path.exists(self.default_po_folder):
            self.input_edit.setText(self.default_po_folder)
            self.output_edit.setText(self.default_po_folder)
            self.log(f"📁 Carpeta predeterminada cargada: {self.default_po_folder}")
            
            # Contar archivos PDF en la carpeta
            try:
                pdf_files = [f for f in os.listdir(self.default_po_folder) if f.lower().endswith('.pdf')]
                if pdf_files:
                    self.log(f"📋 {len(pdf_files)} archivos PDF encontrados en la carpeta")
            except:
                pass
        else:
            self.log(f"ℹ️ Carpeta predeterminada no encontrada: {self.default_po_folder}")

    def select_input(self):
        # Usar la carpeta predeterminada como punto de inicio si existe
        start_folder = self.default_po_folder if os.path.exists(self.default_po_folder) else ""
        folder = QFileDialog.getExistingDirectory(
            self, 
            "Seleccionar Carpeta con POs",
            start_folder
        )
        if folder:
            self.input_edit.setText(folder)
            self.log(f"📁 Carpeta POs seleccionada: {folder}")
            # Si no hay carpeta de salida, usar la misma
            if not self.output_edit.text():
                self.output_edit.setText(folder)

    def select_output(self):
        # Usar la carpeta de entrada actual como punto de inicio
        start_folder = self.input_edit.text() or self.default_po_folder
        if not os.path.exists(start_folder):
            start_folder = ""
            
        folder = QFileDialog.getExistingDirectory(
            self, 
            "Seleccionar Carpeta de Salida",
            start_folder
        )
        if folder:
            self.output_edit.setText(folder)
            self.log(f"💾 Carpeta de salida seleccionada: {folder}")

    def open_output_folder(self):
        folder = self.output_edit.text() or self.input_edit.text()
        if folder and os.path.exists(folder):
            QDesktopServices.openUrl(QUrl.fromLocalFile(folder))
        else:
            QMessageBox.warning(self, "Error", "Carpeta no encontrada")

    def preview_invoice(self, item):
        folder = self.output_edit.text() or self.input_edit.text()
        if folder:
            path = os.path.join(folder, item.text())
            if os.path.exists(path):
                QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    # NUEVA FUNCIÓN: Cargar facturas existentes desde una carpeta
    def load_existing_invoices(self):
        """Carga facturas ya generadas desde una carpeta para enviarlas por email."""
        # Seleccionar carpeta con facturas
        start_folder = self.output_edit.text() or self.input_edit.text() or self.default_po_folder
        if not os.path.exists(start_folder):
            start_folder = ""
        
        folder = QFileDialog.getExistingDirectory(
            self, 
            "Seleccionar Carpeta con Facturas Existentes",
            start_folder
        )
        
        if not folder:
            return
        
        # Buscar archivos PDF en la carpeta
        try:
            pdf_files = [f for f in os.listdir(folder) 
                        if f.lower().endswith('.pdf') and 'invoice' in f.lower()]
            
            if not pdf_files:
                QMessageBox.warning(
                    self, 
                    "Sin Facturas", 
                    "No se encontraron archivos de facturas (PDF con 'invoice' en el nombre) en la carpeta seleccionada."
                )
                return
            
            # Limpiar lista anterior
            self.list_widget.clear()
            self.invoices = []
            
            # Cargar facturas encontradas
            for pdf_file in sorted(pdf_files):
                full_path = os.path.join(folder, pdf_file)
                self.invoices.append(full_path)
                self.list_widget.addItem(pdf_file)
            
            # Actualizar carpeta de salida
            self.output_edit.setText(folder)
            
            # Habilitar botones
            self.send_btn.setEnabled(True)
            self.open_btn.setEnabled(True)
            
            # Log
            self.log(f"📥 Cargadas {len(self.invoices)} facturas desde: {folder}")
            
            QMessageBox.information(
                self, 
                "Facturas Cargadas", 
                f"Se cargaron {len(self.invoices)} facturas correctamente.\n\n"
                f"Ahora puedes enviarlas por email usando el botón '📤 Enviar Emails'."
            )
            
        except Exception as e:
            self.log(f"❌ Error cargando facturas: {e}")
            QMessageBox.critical(self, "Error", f"Error cargando facturas: {e}")

    def generate(self):
        # Validaciones
        inp = self.input_edit.text().strip()
        if not inp:
            QMessageBox.warning(self, "Error", "Por favor selecciona la carpeta con los archivos PO")
            return
        
        if not os.path.exists(inp):
            QMessageBox.warning(self, "Error", f"La carpeta no existe: {inp}")
            return
        
        if not os.path.exists(TEMPLATE_PATH):
            QMessageBox.critical(self, "Error", f"Template de factura no encontrado: {TEMPLATE_PATH}")
            return
        
        # Carpeta de salida
        out = self.output_edit.text().strip() or inp
        if not os.path.exists(out):
            try:
                os.makedirs(out)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"No se pudo crear la carpeta de salida: {e}")
                return
        
        # Obtener número de factura inicial desde el servidor
        default_num = get_last_invoice_number() + 1
        
        # Mostrar diálogo con el número sugerido desde el servidor
        num, ok = QInputDialog.getInt(
            self, 
            "Número de Factura", 
            f"Número inicial de factura (desde servidor: {default_num}):", 
            default_num, 1, 999999, 1
        )
        if not ok:
            return
        
        # Limpiar interfaz
        self.log_text.clear()
        self.list_widget.clear()
        self.progress.setValue(0)
        
        self.log("🚀 Iniciando generación de facturas...")
        self.log(f"📁 Carpeta POs: {inp}")
        self.log(f"💾 Carpeta salida: {out}")
        self.log(f"🔢 Número inicial: {num}")
        
        # Si el usuario cambió el número, actualizar en el servidor ANTES de generar
        if num != default_num:
            self.log(f"📡 Actualizando servidor con nuevo número inicial: {num - 1}")
            save_last_invoice_number(num - 1)
        
        # Generar facturas
        try:
            self.invoices = process_pdf_files(inp, out, num, log=self.log)
            
            # Actualizar lista
            for inv in self.invoices:
                self.list_widget.addItem(os.path.basename(inv))
            
            if self.invoices:
                self.log(f"✅ Generación completada. {len(self.invoices)} facturas creadas.")
                self.log(f"📡 Actualizando servidor con último número generado...")
                
                # Forzar actualización del servidor con el último número usado
                ultimo_numero = num + len(self.invoices) - 1
                save_last_invoice_number(ultimo_numero)
                
                self.send_btn.setEnabled(True)
                self.open_btn.setEnabled(True)
                
                # Sincronizar con el servidor después de generar para verificar
                self.sync_with_server()
                
                QMessageBox.information(self, "Éxito", f"Se generaron {len(self.invoices)} facturas correctamente.\nÚltimo número: {ultimo_numero}")
                
                # Envío automático si está habilitado
                if self.auto_send_checkbox.isChecked():
                    self.log("📤 Iniciando envío automático...")
                    self.send_emails()
            else:
                self.log("❌ No se pudieron generar facturas. Revisa los archivos PO.")
                QMessageBox.warning(self, "Sin Resultados", "No se pudieron generar facturas. Revisa que los archivos PO sean válidos.")
                
        except Exception as e:
            self.log(f"❌ Error durante la generación: {e}")
            QMessageBox.critical(self, "Error", f"Error durante la generación: {e}")

    def send_emails(self):
        if not self.invoices:
            QMessageBox.information(self, "Sin Facturas", "No hay facturas para enviar. Genera o carga las facturas primero.")
            return
        
        if not EMAIL_PASSWORD:
            QMessageBox.critical(self, "Error de Configuración", "Contraseña de email no configurada.")
            return
        
        recipient = self.email_edit.text().strip() or RECIPIENT_EMAIL
        if not recipient:
            QMessageBox.warning(self, "Error", "Especifica un email de destino.")
            return
        
        # Confirmar envío
        reply = QMessageBox.question(
            self, 
            "Confirmar Envío", 
            f"¿Enviar {len(self.invoices)} facturas a {recipient}?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        self.log(f"📧 Iniciando envío de {len(self.invoices)} emails a {recipient}")
        
        self.progress.setValue(0)
        self.progress.setMaximum(len(self.invoices))
        
        sent = 0
        failed = 0
        
        for i, inv in enumerate(self.invoices):
            try:
                subject = os.path.basename(inv)
                body = f"Hi!\n\nAttaching \"{subject}\"\n\nThanks,\nBest,\nFernando."
                
                if send_email(subject, body, inv, recipient, log=self.log):
                    sent += 1
                    self.log(f"✅ Enviado: {subject}")
                else:
                    failed += 1
                    self.log(f"❌ Falló: {subject}")
                
                self.progress.setValue(i + 1)
                QApplication.processEvents()
                time.sleep(2)  # Pausa entre envíos
                
            except Exception as e:
                failed += 1
                self.log(f"❌ Error enviando {os.path.basename(inv)}: {e}")
        
        # Resumen final
        self.log(f"📊 Proceso completado: {sent} enviados, {failed} fallaron")
        
        if sent > 0:
            QMessageBox.information(self, "Envío Completado", f"Se enviaron {sent} de {len(self.invoices)} facturas correctamente.")
        else:
            QMessageBox.warning(self, "Envío Fallido", "No se pudo enviar ninguna factura. Revisa la configuración de email.")

def test_server_connection():
    """Función de prueba para verificar la conexión con el servidor."""
    print("=" * 60)
    print("PRUEBA DE CONEXIÓN AL SERVIDOR")
    print("=" * 60)
    
    print(f"\n🔍 URL: {REMOTE_COUNTER_URL}")
    print(f"🔑 Token: {REMOTE_COUNTER_TOKEN[:10]}...")
    
    # Prueba 1: Verificar conectividad básica
    print("\n1️⃣ Verificando conectividad con el servidor...")
    try:
        import socket
        socket.create_connection(("boogiemanmedia.com", 443), timeout=5)
        print("✅ Servidor alcanzable")
    except Exception as e:
        print(f"❌ Error de conectividad: {e}")
        return
    
    # Prueba 2: Intentar obtener el número actual con GET
    print("\n2️⃣ Obteniendo número actual del servidor (método GET)...")
    result = _remote_call({"action": "current"})
    
    if result:
        print(f"✅ Respuesta exitosa: {result}")
        if result.get("ok") and "last" in result:
            print(f"📊 Último número de invoice: {result['last']}")
        else:
            print("⚠️ Respuesta inesperada del servidor")
    else:
        print("❌ No se pudo obtener respuesta del servidor")
    
    # Prueba 3: Probar directamente con GET en la URL
    print("\n3️⃣ Probando acceso directo con GET...")
    try:
        test_url = f"{REMOTE_COUNTER_URL}?action=current&token={REMOTE_COUNTER_TOKEN}"
        print(f"URL completa: {test_url[:80]}...")
        
        req = urllib.request.Request(test_url)
        req.add_header('User-Agent', 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36')
        
        with urllib.request.urlopen(req, timeout=10) as response:
            data = response.read().decode('utf-8')
            print(f"✅ Respuesta obtenida: {data}")
            
            try:
                json_data = json.loads(data)
                if json_data.get("ok"):
                    print(f"✅ JSON válido - Último invoice: {json_data.get('last', 'N/A')}")
            except:
                if data.strip().isdigit():
                    print(f"✅ Número directo obtenido: {data.strip()}")
                    
    except urllib.error.HTTPError as e:
        print(f"❌ Error HTTP {e.code}: {e.reason}")
        try:
            error_body = e.read().decode('utf-8')
            print(f"   Respuesta del servidor: {error_body}")
        except:
            pass
    except Exception as e:
        print(f"❌ Error: {e}")
    
    print("\n" + "=" * 60)

def main():
    # Opción para ejecutar prueba de conexión
    if len(sys.argv) > 1 and sys.argv[1] == "--test":
        test_server_connection()
        return
    
    app = QApplication(sys.argv)
    
    # Verificar dependencias
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(script_dir, "INVOICE_BASE_DynamicConvert_FormCamposOK.pdf")
    
    if not os.path.exists(template_path):
        QMessageBox.critical(None, "Error", f"Template de factura no encontrado:\n{template_path}")
        sys.exit(1)
    
    win = InvoiceApp()
    win.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
