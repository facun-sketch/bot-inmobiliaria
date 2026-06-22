
from flask import Flask, request, jsonify, send_from_directory, send_file, render_template, session, redirect, url_for
from flask_cors import CORS
from openai import OpenAI
from twilio.twiml.messaging_response import MessagingResponse
from twilio.rest import Client
import csv
from datetime import datetime, timedelta
import os
import json
import sqlite3
import io
from werkzeug.security import check_password_hash, generate_password_hash

print("APP IMPORTADA")

app = Flask(__name__)
app.secret_key = os.getenv('CRM_SECRET_KEY', 'crm-local-secret-change-in-production')
app.permanent_session_lifetime = timedelta(minutes=int(os.getenv('CRM_SESSION_TIMEOUT_MINUTES', '60')))
CORS(app)

print("FLASK LISTO")

# =========================
# PAGINAS
# =========================

@app.route("/")
@app.route("/index.html")
def home():
    return send_from_directory(".", "index.html")

@app.route("/demo")
def demo():
    return render_template("demo.html")

@app.route("/health")
def health():
    return "OK"

@app.route("/detalle-villa-aurelia.html")
def villa_aurelia():
    return send_from_directory(".", "detalle-villa-aurelia.html")

@app.route("/detalle-bernardino.html")
def bernardino():
    return send_from_directory(".", "detalle-bernardino.html")

@app.route("/detalle-edificio-ventura-ykua-sati.html")
def ykua_sati():
    return send_from_directory(".", "detalle-edificio-ventura-ykua-sati.html")

@app.route("/detalle-ventura-hassler.html")
def hassler():
    return send_from_directory(".", "detalle-ventura-hassler.html")

@app.route('/detalle-duplex-vya-raity.html')
def duplex_vya_raity():
    return send_file('detalle-duplex-vya-raity.html')

@app.route("/detalle-altea-laguna.html")
def altea_laguna():
    return send_from_directory(".", "detalle-altea-laguna.html")

@app.route("/detalle-altea-sol.html")
def altea_sol():
    return send_from_directory(".", "detalle-altea-sol.html")

@app.route("/detalle-aqua-village.html")
def aqua_village():
    return send_from_directory(".", "detalle-aqua-village.html")

@app.route("/detalle-altea-hub.html")
def altea_hub():
    return send_from_directory(".", "detalle-altea-hub.html")

@app.route("/detalle-duplex-shopping-sol.html")
def detalle_duplex_shopping_sol():
    return send_from_directory(".", "detalle-duplex-shopping-sol.html")

@app.route("/detalle-residencia-barrio-las-mercedes.html")
def detalle_residencia_barrio_las_mercedes():
    return send_from_directory(".", "detalle-residencia-barrio-las-mercedes.html")

@app.route("/detalle-departamento-torre-augusta.html")
def detalle_departamento_torre_augusta():
    return send_from_directory(".", "detalle-departamento-torre-augusta.html")

@app.route("/detalle-inversion-luque-aregua.html")
def detalle_inversion_luque_aregua():
    return send_from_directory(".", "detalle-inversion-luque-aregua.html")

@app.route("/detalle-terreno-zona-terminal.html")
def detalle_terreno_zona_terminal():
    return send_from_directory(".", "detalle-terreno-zona-terminal.html")

@app.route("/detalle-terreno-barrio-mburucuya.html")
def detalle_terreno_barrio_mburucuya():
    return send_from_directory(".", "detalle-terreno-barrio-mburucuya.html")


# =========================
# ARCHIVOS
# =========================

@app.route('/assets/<path:filename>')
def assets_files(filename):
    return send_from_directory('assets', filename)

@app.route('/imagenes/<path:filename>')
def imagenes_files(filename):
    return send_from_directory('imagenes', filename)

# =========================
# OPENAI
# =========================

openai_api_key = os.getenv("OPENAI_API_KEY")
cliente = OpenAI(api_key=openai_api_key) if openai_api_key else None

print("OPENAI CONFIGURADO")

# =========================
# TWILIO
# =========================

account_sid = os.getenv("TWILIO_ACCOUNT_SID")
auth_token = os.getenv("TWILIO_AUTH_TOKEN")

twilio_client = Client(account_sid, auth_token) if account_sid and auth_token else None

print("TWILIO CONFIGURADO")

# =========================
# GUARDAR CLIENTES
# =========================

def guardar_en_excel(numero_cliente, mensaje, respuesta, estado):

    with open(
        'clientes.csv',
        mode='a',
        newline='',
        encoding='utf-8'
    ) as file:

        writer = csv.writer(file)

        writer.writerow([
            datetime.now(),
            numero_cliente,
            mensaje,
            respuesta,
            estado
        ])

# =========================
# CLIENTES CALIENTES
# =========================

def guardar_cliente_caliente(numero_cliente, mensaje, respuesta):

    with open(
        'clientes_calientes.csv',
        mode='a',
        newline='',
        encoding='utf-8'
    ) as file:

        writer = csv.writer(file)

        writer.writerow([
            datetime.now(),
            numero_cliente,
            mensaje,
            respuesta,
            "CALIENTE 🔥"
        ])

# =========================
# INTERESADOS
# =========================

INTERESADOS_HEADERS = [
    'fecha',
    'nombre',
    'telefono',
    'email',
    'propiedad',
    'mensaje',
    'estado',
    'fecha_seguimiento',
    'fuente'
]

ESTADOS_INTERESADO = [
    'Nuevo',
    'Contactado',
    'Visita agendada',
    'Cerrado'
]

ESTADOS_LEAD = [
    'Nuevo',
    'Contactado',
    'Visita agendada',
    'Negociación',
    'Cerrado'
]

FUENTES_LEAD = [
    'Web',
    'InfoCasas',
    'Marketplace',
    'WhatsApp',
    'Manual'
]

def columna_a_letra(numero):

    letras = ''

    while numero:
        numero, resto = divmod(numero - 1, 26)
        letras = chr(65 + resto) + letras

    return letras

def asegurar_columnas_interesados(worksheet, total_columnas):

    columnas_actuales = getattr(worksheet, 'col_count', 0) or 0

    if columnas_actuales < total_columnas:
        worksheet.resize(cols=total_columnas)

def asegurar_headers_interesados(worksheet):

    headers_actuales = worksheet.row_values(1)

    if not headers_actuales:
        asegurar_columnas_interesados(worksheet, len(INTERESADOS_HEADERS))
        worksheet.append_row(INTERESADOS_HEADERS)
        return INTERESADOS_HEADERS[:]

    headers_actualizados = headers_actuales[:]

    for header in INTERESADOS_HEADERS:
        if header not in headers_actualizados:
            headers_actualizados.append(header)

    if headers_actualizados != headers_actuales:
        asegurar_columnas_interesados(worksheet, len(headers_actualizados))
        ultima_columna = columna_a_letra(len(headers_actualizados))
        worksheet.update(f'A1:{ultima_columna}1', [headers_actualizados])

    return headers_actualizados

def lead_desde_fila(headers, row):

    lead = {
        header: row[pos] if pos < len(row) else ''
        for pos, header in enumerate(headers)
    }

    if 'fuente' not in lead or not lead.get('fuente'):
        lead['fuente'] = 'Web'

    if 'fecha_seguimiento' not in lead:
        lead['fecha_seguimiento'] = ''

    for header in INTERESADOS_HEADERS:
        if header not in lead:
            lead[header] = ''

    return lead

def google_sheets_configurado():

    return bool(
        os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON')
        and os.getenv('GOOGLE_SHEET_ID')
    )

def normalizar_fila_interesado(headers, row):

    lead = lead_desde_fila(headers, row)

    return [
        lead.get(header, '')
        for header in INTERESADOS_HEADERS
    ]

def leer_interesados_csv():

    archivo = 'interesados.csv'

    if not os.path.exists(archivo) or os.path.getsize(archivo) == 0:
        return INTERESADOS_HEADERS[:], []

    with open(archivo, mode='r', newline='', encoding='utf-8') as file:
        reader = csv.reader(file)
        contenido = list(reader)

    if not contenido:
        return INTERESADOS_HEADERS[:], []

    headers = contenido[0] or INTERESADOS_HEADERS[:]
    filas = [
        normalizar_fila_interesado(headers, row)
        for row in contenido[1:]
        if any(str(cell).strip() for cell in row)
    ]

    return INTERESADOS_HEADERS[:], filas

def escribir_interesados_csv(rows):

    with open('interesados.csv', mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(INTERESADOS_HEADERS)
        writer.writerows(rows)

def guardar_interesado_csv(row):

    headers, rows = leer_interesados_csv()
    rows.append(normalizar_fila_interesado(INTERESADOS_HEADERS, row))
    escribir_interesados_csv(rows)
    print("Interesado guardado en interesados.csv")

def obtener_interesados_datos():

    if not google_sheets_configurado():
        print("Google Sheets no configurado. Usando interesados.csv")
        return leer_interesados_csv()

    try:
        worksheet, gspread = obtener_interesados_worksheet()
        headers = asegurar_headers_interesados(worksheet)
        values = worksheet.get_all_values()
        rows = values[1:] if len(values) > 1 else []
        return headers, rows
    except Exception as e:
        print(f"Error Google Sheets. Usando interesados.csv: {e}")
        return leer_interesados_csv()

def actualizar_estado_interesado_storage(row_number, nuevo_estado):

    fecha_seguimiento = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    if google_sheets_configurado():
        try:
            worksheet, gspread = obtener_interesados_worksheet()
            headers = asegurar_headers_interesados(worksheet)

            estado_col = headers.index('estado') + 1
            seguimiento_col = headers.index('fecha_seguimiento') + 1

            worksheet.update_cell(row_number, estado_col, nuevo_estado)
            worksheet.update_cell(row_number, seguimiento_col, fecha_seguimiento)
            return
        except Exception as e:
            print(f"Error actualizando Google Sheets. Usando interesados.csv: {e}")

    headers, rows = leer_interesados_csv()
    row_index = row_number - 2

    if row_index < 0 or row_index >= len(rows):
        raise RuntimeError("Fila inválida")

    lead = lead_desde_fila(headers, rows[row_index])
    lead['estado'] = nuevo_estado
    lead['fecha_seguimiento'] = fecha_seguimiento
    rows[row_index] = [
        lead.get(header, '')
        for header in INTERESADOS_HEADERS
    ]
    escribir_interesados_csv(rows)

def guardar_interesado_backup(row):

    archivo = 'interesados_backup.csv'
    crear_encabezado = not os.path.exists(archivo) or os.path.getsize(archivo) == 0

    with open(archivo, mode='a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)

        if crear_encabezado:
            writer.writerow(INTERESADOS_HEADERS)

        writer.writerow(row)

def guardar_interesado_google_sheets(row):

    if not google_sheets_configurado():
        print("Google Sheets no configurado. Guardando interesado en interesados.csv")
        guardar_interesado_csv(row)
        return

    try:
        print("Intentando guardar interesado en Google Sheets")

        worksheet, gspread = obtener_interesados_worksheet()

        worksheet.append_row(row)

        print("Interesado guardado en Google Sheets")
    except Exception as e:
        print(f"Error Google Sheets. Guardando interesado en interesados.csv: {e}")
        guardar_interesado_csv(row)

def obtener_interesados_worksheet():

    import json
    import gspread
    from google.oauth2.service_account import Credentials

    service_account_json = os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON')
    sheet_id = os.getenv('GOOGLE_SHEET_ID')

    if not service_account_json or not sheet_id:
        return None, None

    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]

    service_account_info = json.loads(service_account_json)
    credentials = Credentials.from_service_account_info(
        service_account_info,
        scopes=scopes
    )

    client = gspread.authorize(credentials)
    worksheet_name = os.getenv('GOOGLE_SHEET_TAB', 'Interesados')
    spreadsheet = client.open_by_key(sheet_id)

    try:
        worksheet = spreadsheet.worksheet(worksheet_name)
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title=worksheet_name,
            rows=1000,
            cols=len(INTERESADOS_HEADERS)
        )

    asegurar_headers_interesados(worksheet)

    return worksheet, gspread

def enviar_email_interesado(fecha, nombre, telefono, email, propiedad, mensaje):

    import smtplib
    from email.message import EmailMessage

    smtp_host = os.getenv('SMTP_HOST')
    smtp_port = int(os.getenv('SMTP_PORT', 587))
    smtp_user = os.getenv('SMTP_USER')
    smtp_password = os.getenv('SMTP_PASSWORD')
    email_from = os.getenv('EMAIL_FROM', smtp_user)
    email_to = os.getenv('EMAIL_TO')

    if not smtp_host or not smtp_user or not smtp_password or not email_to:
        raise RuntimeError("Faltan variables SMTP para enviar email")

    body = f"""
Nuevo interesado recibido.

Fecha: {fecha}
Nombre: {nombre}
Teléfono: {telefono}
Email: {email}
Propiedad: {propiedad}
Mensaje: {mensaje}
"""

    msg = EmailMessage()
    msg['Subject'] = f'Nuevo interesado - {propiedad}'
    msg['From'] = email_from
    msg['To'] = email_to
    msg.set_content(body)

    with smtplib.SMTP(smtp_host, smtp_port, timeout=15) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)

# =========================
# INFOCASAS HOTMAIL SYNC
# =========================

INFOCASAS_IMPORTADOS_CSV = 'infocasas_importados.csv'

def microsoft_graph_configurado():

    return all([
        os.getenv('MS_CLIENT_ID'),
        os.getenv('MS_CLIENT_SECRET'),
        os.getenv('MS_TENANT_ID', 'consumers'),
        os.getenv('MS_REDIRECT_URI'),
        os.getenv('MS_REFRESH_TOKEN')
    ])

def obtener_ms_access_token():

    import requests

    tenant_id = os.getenv('MS_TENANT_ID', 'consumers')
    token_url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'

    data = {
        'client_id': os.getenv('MS_CLIENT_ID'),
        'client_secret': os.getenv('MS_CLIENT_SECRET'),
        'redirect_uri': os.getenv('MS_REDIRECT_URI'),
        'refresh_token': os.getenv('MS_REFRESH_TOKEN'),
        'grant_type': 'refresh_token',
        'scope': 'https://graph.microsoft.com/Mail.Read offline_access'
    }

    response = requests.post(token_url, data=data, timeout=20)
    response.raise_for_status()

    token_data = response.json()
    return token_data['access_token']

def leer_ultimos_correos_graph(limit=20):

    import requests

    access_token = obtener_ms_access_token()
    url = 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages'
    params = {
        '$top': limit,
        '$orderby': 'receivedDateTime desc',
        '$select': 'id,internetMessageId,subject,from,receivedDateTime,bodyPreview,body'
    }
    headers = {
        'Authorization': f'Bearer {access_token}'
    }

    response = requests.get(url, headers=headers, params=params, timeout=20)
    response.raise_for_status()

    mensajes = response.json().get('value', [])
    print(f"Correos leídos: {len(mensajes)}")
    return mensajes

def limpiar_html(texto):

    import re
    import html

    texto = texto or ''
    texto = re.sub(r'<br\s*/?>', '\n', texto, flags=re.I)
    texto = re.sub(r'</p\s*>', '\n', texto, flags=re.I)
    texto = re.sub(r'<[^>]+>', ' ', texto)
    texto = html.unescape(texto)
    texto = re.sub(r'[ \t]+', ' ', texto)
    texto = re.sub(r'\n\s+', '\n', texto)
    return texto.strip()

def obtener_texto_correo(correo):

    body = correo.get('body') or {}
    body_content = body.get('content') or ''
    body_preview = correo.get('bodyPreview') or ''
    return limpiar_html(f'{body_preview}\n{body_content}')

def es_correo_infocasas(correo):

    remitente = correo.get('from') or {}
    email_info = remitente.get('emailAddress') or {}
    sender = f"{email_info.get('name', '')} {email_info.get('address', '')}".lower()
    subject = (correo.get('subject') or '').lower()
    body = obtener_texto_correo(correo).lower()

    return (
        'infocasas' in sender
        and 'te han consultado por tu propiedad' in subject
        and '¡tienes una nueva consulta!' in body
        and 'consultó por tu propiedad' in body
    )

def extraer_linea(texto, etiquetas):

    import re

    for etiqueta in etiquetas:
        patron = rf'{etiqueta}\s*[:\-]\s*(.+)'
        match = re.search(patron, texto, flags=re.I)

        if match:
            return match.group(1).strip()

    return ''

def extraer_lead_infocasas(correo):

    import re

    texto = obtener_texto_correo(correo)

    nombre = extraer_linea(texto, [
        'nombre',
        'contacto',
        'interesado',
        'cliente'
    ])

    telefono = extraer_linea(texto, [
        'teléfono',
        'telefono',
        'celular',
        'móvil',
        'movil',
        'whatsapp'
    ])

    email = extraer_linea(texto, [
        'email',
        'e-mail',
        'correo'
    ])

    propiedad = ''
    propiedad_match = re.search(
        r'consultó por tu propiedad\s*:\s*(.+?)(?:\n|$)',
        texto,
        flags=re.I
    )

    if propiedad_match:
        propiedad = propiedad_match.group(1).strip()

    mensaje = extraer_linea(texto, [
        'mensaje',
        'consulta',
        'comentario'
    ])

    if not telefono:
        telefono_match = re.search(r'(\+?595[\s\-]?\d{6,10}|0\d{8,10})', texto)
        telefono = telefono_match.group(1).strip() if telefono_match else ''

    if not email:
        email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', texto)
        email = email_match.group(0).strip() if email_match else ''

    if not nombre:
        nombre = 'Contacto InfoCasas'

    if not mensaje:
        mensaje = texto

    return {
        'message_id': correo.get('internetMessageId') or correo.get('id') or '',
        'fecha': correo.get('receivedDateTime') or datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'nombre': nombre,
        'telefono': telefono,
        'email': email,
        'propiedad': propiedad,
        'mensaje': mensaje,
        'fuente': 'InfoCasas',
        'estado': 'Nuevo'
    }

def leer_infocasas_importados():

    if not os.path.exists(INFOCASAS_IMPORTADOS_CSV):
        return set()

    with open(INFOCASAS_IMPORTADOS_CSV, mode='r', newline='', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        return {
            row.get('message_id', '')
            for row in reader
            if row.get('message_id')
        }

def guardar_infocasas_importado(message_id):

    crear_encabezado = (
        not os.path.exists(INFOCASAS_IMPORTADOS_CSV)
        or os.path.getsize(INFOCASAS_IMPORTADOS_CSV) == 0
    )

    with open(INFOCASAS_IMPORTADOS_CSV, mode='a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)

        if crear_encabezado:
            writer.writerow(['message_id', 'fecha_importacion'])

        writer.writerow([
            message_id,
            datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        ])

def obtener_leads_infocasas_preview(limit=20):

    correos = leer_ultimos_correos_graph(limit)
    procesados = leer_infocasas_importados()
    leads = []
    duplicados = 0
    detectados = 0

    for correo in correos:
        if not es_correo_infocasas(correo):
            continue

        detectados += 1
        lead = extraer_lead_infocasas(correo)

        if lead['message_id'] in procesados:
            lead['duplicado'] = True
            duplicados += 1
        else:
            lead['duplicado'] = False

        leads.append(lead)

    print(f"Correos detectados como InfoCasas: {detectados}")
    print(f"Duplicados omitidos: {duplicados}")
    return leads

def importar_leads_infocasas(limit=20):

    leads = obtener_leads_infocasas_preview(limit)
    importados = 0
    duplicados = 0

    for lead in leads:
        if lead.get('duplicado'):
            duplicados += 1
            continue

        row = [
            datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            lead.get('nombre', ''),
            lead.get('telefono', ''),
            lead.get('email', ''),
            lead.get('propiedad', ''),
            lead.get('mensaje', ''),
            'Nuevo',
            '',
            'InfoCasas'
        ]

        guardar_interesado_google_sheets(row)
        guardar_infocasas_importado(lead.get('message_id', ''))
        importados += 1

    print(f"Leads importados: {importados}")
    print(f"Duplicados omitidos: {duplicados}")
    return leads, importados, duplicados

def respuesta_interesado(titulo, mensaje):

    return f'''
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{titulo}</title>
        <style>
            body{{
                margin:0;
                font-family:Arial;
                background:#f4f4f1;
                color:#111;
                display:flex;
                min-height:100vh;
                align-items:center;
                justify-content:center;
                padding:24px;
            }}
            .box{{
                background:white;
                max-width:520px;
                padding:40px;
                border-radius:24px;
                box-shadow:0 10px 30px rgba(0,0,0,.08);
                text-align:center;
            }}
            .btn{{
                display:inline-block;
                padding:14px 28px;
                border-radius:999px;
                background:#111;
                color:white;
                text-decoration:none;
                margin-top:20px;
            }}
        </style>
    </head>
    <body>
        <div class="box">
            <h1>{titulo}</h1>
            <p>{mensaje}</p>
            <a class="btn" href="/">Volver al inicio</a>
        </div>
    </body>
    </html>
    '''

@app.route('/guardar-interesado', methods=['POST'])
def guardar_interesado():

    fecha = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    nombre = request.form.get('nombre', '').strip()
    telefono = request.form.get('telefono', '').strip()
    email = request.form.get('email', '').strip()
    propiedad = request.form.get('propiedad', '').strip()
    mensaje = request.form.get('mensaje', '').strip()
    estado = 'Nuevo'
    fecha_seguimiento = ''
    fuente = 'Web'

    row = [
        fecha,
        nombre,
        telefono,
        email,
        propiedad,
        mensaje,
        estado,
        fecha_seguimiento,
        fuente
    ]

    try:
        guardar_interesado_google_sheets(row)

        try:
            enviar_email_interesado(
                fecha,
                nombre,
                telefono,
                email,
                propiedad,
                mensaje
            )
            print("Email enviado correctamente")
        except Exception as email_error:
            print(f"Error enviando email: {email_error}")

        return respuesta_interesado(
            "Gracias.",
            "Recibimos tu consulta y nos pondremos en contacto."
        )

    except Exception as e:
        print(f"Error Google Sheets: {e}")

        try:
            guardar_interesado_backup(row)
        except Exception as backup_error:
            print(f"Error backup interesados: {backup_error}")

        return respuesta_interesado(
            "No pudimos guardar tu consulta.",
            "Por favor escribinos por WhatsApp."
        ), 500

@app.route('/admin/interesados')
def admin_interesados():

    try:
        import html

        headers, rows = obtener_interesados_datos()

        table_rows = ""

        for index, row in enumerate(rows, start=2):
            interesado = lead_desde_fila(headers, row)

            nombre = html.escape(str(interesado.get('nombre', '')))
            telefono = html.escape(str(interesado.get('telefono', '')))
            email = html.escape(str(interesado.get('email', '')))
            propiedad = html.escape(str(interesado.get('propiedad', '')))
            estado_actual = str(interesado.get('estado', '')) or 'Nuevo'
            estado = html.escape(estado_actual)
            fecha = html.escape(str(interesado.get('fecha', '')))

            botones_estado = ""

            for estado_opcion in ESTADOS_INTERESADO:
                estado_opcion_html = html.escape(estado_opcion)
                activo = ' active' if estado_opcion == estado_actual else ''

                botones_estado += f"""
                <form action="/admin/interesados/estado" method="POST">
                    <input type="hidden" name="row" value="{index}">
                    <input type="hidden" name="estado" value="{estado_opcion_html}">
                    <button class="state-btn{activo}" type="submit">{estado_opcion_html}</button>
                </form>
                """

            table_rows += f"""
            <tr>
                <td>{nombre}</td>
                <td>{telefono}</td>
                <td>{email}</td>
                <td>{propiedad}</td>
                <td>
                    <strong class="status">{estado}</strong>
                    <div class="state-actions">
                        {botones_estado}
                    </div>
                </td>
                <td>{fecha}</td>
            </tr>
            """

        return f'''
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Interesados</title>
            <style>
                body{{
                    margin:0;
                    font-family:Arial;
                    background:#f4f4f1;
                    color:#111;
                    padding:32px;
                }}
                .page{{
                    max-width:1200px;
                    margin:0 auto;
                }}
                .top{{
                    display:flex;
                    justify-content:space-between;
                    align-items:center;
                    gap:16px;
                    margin-bottom:24px;
                }}
                .top p{{
                    margin:6px 0 0;
                    color:#555;
                }}
                .table-wrap{{
                    overflow-x:auto;
                    background:white;
                    border-radius:16px;
                    box-shadow:0 10px 30px rgba(0,0,0,.08);
                }}
                table{{
                    width:100%;
                    border-collapse:collapse;
                    min-width:900px;
                }}
                th,
                td{{
                    padding:14px;
                    border-bottom:1px solid #ddd;
                    text-align:left;
                    vertical-align:top;
                }}
                th{{
                    background:#111;
                    color:white;
                    font-size:14px;
                }}
                .status{{
                    display:inline-block;
                    margin-bottom:10px;
                }}
                .state-actions{{
                    display:flex;
                    flex-wrap:wrap;
                    gap:8px;
                }}
                .state-actions form{{
                    margin:0;
                }}
                .state-btn{{
                    border:1px solid #ddd;
                    background:#f7f7f7;
                    color:#111;
                    border-radius:999px;
                    padding:8px 12px;
                    cursor:pointer;
                    font-family:Arial;
                    font-size:13px;
                }}
                .state-btn.active{{
                    background:#111;
                    border-color:#111;
                    color:white;
                }}
                @media(max-width:700px){{
                    body{{
                        padding:18px;
                    }}
                    .top{{
                        display:block;
                    }}
                }}
            </style>
        </head>
        <body>
            <main class="page">
                <div class="top">
                    <div>
                        <h1>Interesados</h1>
                        <p>Seguimiento básico de consultas recibidas.</p>
                    </div>
                </div>

                <div class="table-wrap">
                    <table>
                        <thead>
                            <tr>
                                <th>Nombre</th>
                                <th>Teléfono</th>
                                <th>Email</th>
                                <th>Propiedad</th>
                                <th>Estado</th>
                                <th>Fecha</th>
                            </tr>
                        </thead>
                        <tbody>
                            {table_rows}
                        </tbody>
                    </table>
                </div>
            </main>
        </body>
        </html>
        '''

    except Exception as e:
        print(f"Error admin interesados: {e}")
        return "No se pudieron cargar los interesados.", 500

@app.route('/admin/interesados/estado', methods=['POST'])
def actualizar_estado_interesado():

    try:
        row_number = int(request.form.get('row', '0'))
        nuevo_estado = request.form.get('estado', '').strip()

        if row_number < 2:
            raise RuntimeError("Fila inválida")

        if nuevo_estado not in ESTADOS_INTERESADO:
            raise RuntimeError("Estado inválido")

        actualizar_estado_interesado_storage(row_number, nuevo_estado)

        return '''
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <meta http-equiv="refresh" content="0; url=/admin/interesados">
            <title>Estado actualizado</title>
        </head>
        <body>
            <a href="/admin/interesados">Volver</a>
        </body>
        </html>
        '''

    except Exception as e:
        print(f"Error actualizando estado interesado: {e}")
        return "No se pudo actualizar el estado.", 500

@app.route('/admin/leads')
def admin_leads():

    try:
        import html

        headers, rows = obtener_interesados_datos()

        table_rows = ""

        for index, row in enumerate(rows, start=2):
            lead = lead_desde_fila(headers, row)

            fecha = html.escape(str(lead.get('fecha', '')))
            nombre = html.escape(str(lead.get('nombre', '')))
            telefono = html.escape(str(lead.get('telefono', '')))
            email = html.escape(str(lead.get('email', '')))
            propiedad = html.escape(str(lead.get('propiedad', '')))
            mensaje = html.escape(str(lead.get('mensaje', '')))
            fuente = html.escape(str(lead.get('fuente', '') or 'Web'))
            estado_actual = str(lead.get('estado', '')) or 'Nuevo'
            estado = html.escape(estado_actual)

            botones_estado = ""

            for estado_opcion in ESTADOS_LEAD:
                estado_opcion_html = html.escape(estado_opcion)
                activo = ' active' if estado_opcion == estado_actual else ''

                botones_estado += f"""
                <form action="/admin/leads/estado" method="POST">
                    <input type="hidden" name="row" value="{index}">
                    <input type="hidden" name="estado" value="{estado_opcion_html}">
                    <button class="state-btn{activo}" type="submit">{estado_opcion_html}</button>
                </form>
                """

            table_rows += f"""
            <tr>
                <td>{fecha}</td>
                <td>{nombre}</td>
                <td>{telefono}</td>
                <td>{email}</td>
                <td>{propiedad}</td>
                <td class="message-cell">{mensaje}</td>
                <td><span class="source">{fuente}</span></td>
                <td>
                    <strong class="status">{estado}</strong>
                    <div class="state-actions">
                        {botones_estado}
                    </div>
                </td>
            </tr>
            """

        return f'''
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>CRM Leads</title>
            <style>
                body{{
                    margin:0;
                    font-family:Arial;
                    background:#f4f4f1;
                    color:#111;
                    padding:32px;
                }}
                .page{{
                    max-width:1280px;
                    margin:0 auto;
                }}
                .top{{
                    display:flex;
                    justify-content:space-between;
                    align-items:center;
                    gap:16px;
                    margin-bottom:24px;
                }}
                .top p{{
                    margin:6px 0 0;
                    color:#555;
                }}
                .btn{{
                    display:inline-block;
                    border:0;
                    background:#111;
                    color:white;
                    border-radius:999px;
                    padding:12px 18px;
                    text-decoration:none;
                    font-size:14px;
                }}
                .table-wrap{{
                    overflow-x:auto;
                    background:white;
                    border-radius:16px;
                    box-shadow:0 10px 30px rgba(0,0,0,.08);
                }}
                table{{
                    width:100%;
                    border-collapse:collapse;
                    min-width:1100px;
                }}
                th,
                td{{
                    padding:14px;
                    border-bottom:1px solid #ddd;
                    text-align:left;
                    vertical-align:top;
                    font-size:14px;
                }}
                th{{
                    background:#111;
                    color:white;
                }}
                .message-cell{{
                    max-width:260px;
                    line-height:1.4;
                }}
                .source{{
                    display:inline-block;
                    background:#f0f0f0;
                    border-radius:999px;
                    padding:6px 10px;
                }}
                .status{{
                    display:inline-block;
                    margin-bottom:10px;
                }}
                .state-actions{{
                    display:flex;
                    flex-wrap:wrap;
                    gap:8px;
                }}
                .state-actions form{{
                    margin:0;
                }}
                .state-btn{{
                    border:1px solid #ddd;
                    background:#f7f7f7;
                    color:#111;
                    border-radius:999px;
                    padding:8px 12px;
                    cursor:pointer;
                    font-family:Arial;
                    font-size:13px;
                }}
                .state-btn.active{{
                    background:#111;
                    border-color:#111;
                    color:white;
                }}
                @media(max-width:700px){{
                    body{{
                        padding:18px;
                    }}
                    .top{{
                        display:block;
                    }}
                    .btn{{
                        margin-top:14px;
                    }}
                }}
            </style>
        </head>
        <body>
            <main class="page">
                <div class="top">
                    <div>
                        <h1>CRM Leads</h1>
                        <p>Leads unificados por fuente.</p>
                    </div>
                    <a class="btn" href="/admin/leads/nuevo">Cargar lead manual</a>
                </div>

                <div class="table-wrap">
                    <table>
                        <thead>
                            <tr>
                                <th>Fecha</th>
                                <th>Nombre</th>
                                <th>Teléfono</th>
                                <th>Email</th>
                                <th>Propiedad</th>
                                <th>Mensaje</th>
                                <th>Fuente</th>
                                <th>Estado</th>
                            </tr>
                        </thead>
                        <tbody>
                            {table_rows}
                        </tbody>
                    </table>
                </div>
            </main>
        </body>
        </html>
        '''

    except Exception as e:
        import html
        import traceback

        print(f"Error admin leads: {repr(e)}")
        traceback.print_exc()

        error = html.escape(str(e))

        return f'''
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Error CRM Leads</title>
            <style>
                body{{
                    margin:0;
                    font-family:Arial;
                    background:#f4f4f1;
                    color:#111;
                    padding:32px;
                }}
                .box{{
                    max-width:760px;
                    background:white;
                    border-radius:16px;
                    box-shadow:0 10px 30px rgba(0,0,0,.08);
                    padding:24px;
                }}
                code{{
                    display:block;
                    white-space:pre-wrap;
                    background:#f0f0f0;
                    border-radius:10px;
                    padding:14px;
                    margin-top:12px;
                }}
                a{{
                    display:inline-block;
                    margin-top:18px;
                    color:#111;
                }}
            </style>
        </head>
        <body>
            <div class="box">
                <h1>No se pudieron cargar los leads.</h1>
                <p>Revisá la consola de Flask/Railway. El error real también es:</p>
                <code>{error}</code>
                <a href="/admin/leads/nuevo">Cargar lead manual</a>
            </div>
        </body>
        </html>
        ''', 500

@app.route('/admin/leads/nuevo', methods=['GET', 'POST'])
def admin_leads_nuevo():

    import html

    if request.method == 'POST':

        try:
            fecha = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            nombre = request.form.get('nombre', '').strip()
            telefono = request.form.get('telefono', '').strip()
            email = request.form.get('email', '').strip()
            propiedad = request.form.get('propiedad', '').strip()
            mensaje = request.form.get('mensaje', '').strip()
            fuente = request.form.get('fuente', '').strip()
            estado = request.form.get('estado', '').strip()
            fecha_seguimiento = ''

            if fuente not in FUENTES_LEAD:
                raise RuntimeError("Fuente inválida")

            if estado not in ESTADOS_LEAD:
                raise RuntimeError("Estado inválido")

            row = [
                fecha,
                nombre,
                telefono,
                email,
                propiedad,
                mensaje,
                estado,
                fecha_seguimiento,
                fuente
            ]

            guardar_interesado_google_sheets(row)

            return '''
            <!DOCTYPE html>
            <html lang="es">
            <head>
                <meta charset="UTF-8">
                <meta http-equiv="refresh" content="0; url=/admin/leads">
                <title>Lead guardado</title>
            </head>
            <body>
                <a href="/admin/leads">Volver</a>
            </body>
            </html>
            '''

        except Exception as e:
            import traceback

            print(f"Error creando lead manual: {repr(e)}")
            traceback.print_exc()
            return "No se pudo guardar el lead manual.", 500

    opciones_fuente = ""
    for fuente in FUENTES_LEAD:
        fuente_html = html.escape(fuente)
        selected = ' selected' if fuente == 'InfoCasas' else ''
        opciones_fuente += f'<option value="{fuente_html}"{selected}>{fuente_html}</option>'

    opciones_estado = ""
    for estado in ESTADOS_LEAD:
        estado_html = html.escape(estado)
        selected = ' selected' if estado == 'Nuevo' else ''
        opciones_estado += f'<option value="{estado_html}"{selected}>{estado_html}</option>'

    return f'''
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Nuevo Lead</title>
        <style>
            body{{
                margin:0;
                font-family:Arial;
                background:#f4f4f1;
                color:#111;
                padding:32px;
            }}
            .page{{
                max-width:760px;
                margin:0 auto;
            }}
            .top{{
                margin-bottom:24px;
            }}
            .top p{{
                margin:6px 0 0;
                color:#555;
            }}
            form{{
                background:white;
                border-radius:16px;
                box-shadow:0 10px 30px rgba(0,0,0,.08);
                padding:24px;
            }}
            label{{
                display:block;
                font-weight:bold;
                margin:16px 0 8px;
            }}
            input,
            select,
            textarea{{
                width:100%;
                box-sizing:border-box;
                border:1px solid #ddd;
                border-radius:10px;
                padding:12px;
                font-family:Arial;
                font-size:15px;
            }}
            textarea{{
                min-height:120px;
                resize:vertical;
            }}
            .actions{{
                display:flex;
                flex-wrap:wrap;
                gap:12px;
                margin-top:22px;
            }}
            .btn{{
                border:0;
                background:#111;
                color:white;
                border-radius:999px;
                padding:12px 18px;
                text-decoration:none;
                cursor:pointer;
                font-size:14px;
            }}
            .btn-light{{
                background:#eee;
                color:#111;
            }}
            @media(max-width:700px){{
                body{{
                    padding:18px;
                }}
                form{{
                    padding:18px;
                }}
            }}
        </style>
    </head>
    <body>
        <main class="page">
            <div class="top">
                <h1>Nuevo lead</h1>
                <p>Carga manual para InfoCasas, Marketplace, WhatsApp o contactos directos.</p>
            </div>

            <form action="/admin/leads/nuevo" method="POST">
                <label for="nombre">Nombre</label>
                <input id="nombre" name="nombre" type="text" required>

                <label for="telefono">Teléfono</label>
                <input id="telefono" name="telefono" type="text">

                <label for="email">Email</label>
                <input id="email" name="email" type="email">

                <label for="propiedad">Propiedad</label>
                <input id="propiedad" name="propiedad" type="text">

                <label for="mensaje">Mensaje</label>
                <textarea id="mensaje" name="mensaje"></textarea>

                <label for="fuente">Fuente</label>
                <select id="fuente" name="fuente">
                    {opciones_fuente}
                </select>

                <label for="estado">Estado</label>
                <select id="estado" name="estado">
                    {opciones_estado}
                </select>

                <div class="actions">
                    <button class="btn" type="submit">Guardar lead</button>
                    <a class="btn btn-light" href="/admin/leads">Volver</a>
                </div>
            </form>
        </main>
    </body>
    </html>
    '''

@app.route('/admin/leads/estado', methods=['POST'])
def actualizar_estado_lead():

    try:
        row_number = int(request.form.get('row', '0'))
        nuevo_estado = request.form.get('estado', '').strip()

        if row_number < 2:
            raise RuntimeError("Fila inválida")

        if nuevo_estado not in ESTADOS_LEAD:
            raise RuntimeError("Estado inválido")

        actualizar_estado_interesado_storage(row_number, nuevo_estado)

        return '''
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <meta http-equiv="refresh" content="0; url=/admin/leads">
            <title>Estado actualizado</title>
        </head>
        <body>
            <a href="/admin/leads">Volver</a>
        </body>
        </html>
        '''

    except Exception as e:
        print(f"Error actualizando estado lead: {e}")
        return "No se pudo actualizar el estado del lead.", 500

@app.route('/admin/sync/infocasas', methods=['GET', 'POST'])
def admin_sync_infocasas():

    import html

    if not microsoft_graph_configurado():
        return '''
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Sync InfoCasas</title>
            <style>
                body{
                    margin:0;
                    font-family:Arial;
                    background:#f4f4f1;
                    color:#111;
                    padding:32px;
                }
                .box{
                    max-width:760px;
                    background:white;
                    border-radius:16px;
                    box-shadow:0 10px 30px rgba(0,0,0,.08);
                    padding:24px;
                }
                code{
                    display:block;
                    background:#f0f0f0;
                    border-radius:10px;
                    padding:12px;
                    margin:12px 0;
                    white-space:pre-wrap;
                }
                a{
                    color:#111;
                }
            </style>
        </head>
        <body>
            <main class="box">
                <h1>Integración Hotmail no configurada</h1>
                <p>Para sincronizar InfoCasas desde Hotmail/Outlook faltan estas variables:</p>
                <code>MS_CLIENT_ID
MS_CLIENT_SECRET
MS_TENANT_ID=consumers
MS_REDIRECT_URI
MS_REFRESH_TOKEN</code>
                <a href="/admin/leads">Volver al CRM</a>
            </main>
        </body>
        </html>
        '''

    mensaje_resultado = ''

    try:
        if request.method == 'POST':
            leads, importados, duplicados = importar_leads_infocasas(20)
            mensaje_resultado = (
                f'{importados} lead(s) importado(s). '
                f'{duplicados} duplicado(s) omitido(s).'
            )
        else:
            leads = obtener_leads_infocasas_preview(20)

    except Exception as e:
        import traceback

        print(f"Error conexión InfoCasas/Hotmail: {repr(e)}")
        traceback.print_exc()

        error = html.escape(str(e))
        return f'''
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Error Sync InfoCasas</title>
            <style>
                body{{
                    margin:0;
                    font-family:Arial;
                    background:#f4f4f1;
                    color:#111;
                    padding:32px;
                }}
                .box{{
                    max-width:760px;
                    background:white;
                    border-radius:16px;
                    box-shadow:0 10px 30px rgba(0,0,0,.08);
                    padding:24px;
                }}
                code{{
                    display:block;
                    background:#f0f0f0;
                    border-radius:10px;
                    padding:12px;
                    margin-top:12px;
                    white-space:pre-wrap;
                }}
            </style>
        </head>
        <body>
            <main class="box">
                <h1>Error de conexión</h1>
                <p>No se pudieron leer los correos de Hotmail/Outlook.</p>
                <code>{error}</code>
                <a href="/admin/leads">Volver al CRM</a>
            </main>
        </body>
        </html>
        ''', 500

    filas = ''

    for lead in leads:
        nombre = html.escape(str(lead.get('nombre', '')))
        telefono = html.escape(str(lead.get('telefono', '')))
        email = html.escape(str(lead.get('email', '')))
        propiedad = html.escape(str(lead.get('propiedad', '')))
        mensaje = html.escape(str(lead.get('mensaje', '')))
        fecha = html.escape(str(lead.get('fecha', '')))
        estado_importacion = 'Duplicado' if lead.get('duplicado') else 'Nuevo'
        estado_clase = 'duplicate' if lead.get('duplicado') else 'new'

        filas += f'''
        <tr>
            <td>{fecha}</td>
            <td>{nombre}</td>
            <td>{telefono}</td>
            <td>{email}</td>
            <td>{propiedad}</td>
            <td class="message-cell">{mensaje}</td>
            <td><span class="pill {estado_clase}">{estado_importacion}</span></td>
        </tr>
        '''

    if not filas:
        filas = '''
        <tr>
            <td colspan="7">No se detectaron correos de InfoCasas en los últimos 20 correos.</td>
        </tr>
        '''

    resultado_html = ''

    if mensaje_resultado:
        resultado_html = f'<div class="notice">{html.escape(mensaje_resultado)}</div>'

    return f'''
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Sync InfoCasas</title>
        <style>
            body{{
                margin:0;
                font-family:Arial;
                background:#f4f4f1;
                color:#111;
                padding:32px;
            }}
            .page{{
                max-width:1280px;
                margin:0 auto;
            }}
            .top{{
                display:flex;
                justify-content:space-between;
                align-items:center;
                gap:16px;
                margin-bottom:24px;
            }}
            .top p{{
                margin:6px 0 0;
                color:#555;
            }}
            .actions{{
                display:flex;
                gap:10px;
                flex-wrap:wrap;
            }}
            .btn{{
                border:0;
                background:#111;
                color:white;
                border-radius:999px;
                padding:12px 18px;
                text-decoration:none;
                cursor:pointer;
                font-size:14px;
            }}
            .btn-light{{
                background:#eee;
                color:#111;
            }}
            .notice{{
                background:white;
                border-radius:12px;
                padding:14px 16px;
                margin-bottom:18px;
                box-shadow:0 8px 20px rgba(0,0,0,.06);
            }}
            .table-wrap{{
                overflow-x:auto;
                background:white;
                border-radius:16px;
                box-shadow:0 10px 30px rgba(0,0,0,.08);
            }}
            table{{
                width:100%;
                border-collapse:collapse;
                min-width:1100px;
            }}
            th,
            td{{
                padding:14px;
                border-bottom:1px solid #ddd;
                text-align:left;
                vertical-align:top;
                font-size:14px;
            }}
            th{{
                background:#111;
                color:white;
            }}
            .message-cell{{
                max-width:320px;
                line-height:1.4;
            }}
            .pill{{
                display:inline-block;
                border-radius:999px;
                padding:6px 10px;
                background:#eee;
            }}
            .pill.new{{
                background:#e8f4ed;
            }}
            .pill.duplicate{{
                background:#f7e8e8;
            }}
            @media(max-width:700px){{
                body{{
                    padding:18px;
                }}
                .top{{
                    display:block;
                }}
                .actions{{
                    margin-top:14px;
                }}
            }}
        </style>
    </head>
    <body>
        <main class="page">
            <div class="top">
                <div>
                    <h1>Sync InfoCasas</h1>
                    <p>Preview de los últimos 20 correos detectados desde Hotmail/Outlook.</p>
                </div>
                <div class="actions">
                    <form action="/admin/sync/infocasas" method="POST">
                        <button class="btn" type="submit">Importar al CRM</button>
                    </form>
                    <a class="btn btn-light" href="/admin/leads">Volver al CRM</a>
                </div>
            </div>

            {resultado_html}

            <div class="table-wrap">
                <table>
                    <thead>
                        <tr>
                            <th>Fecha</th>
                            <th>Nombre</th>
                            <th>Teléfono</th>
                            <th>Email</th>
                            <th>Propiedad</th>
                            <th>Mensaje</th>
                            <th>Importación</th>
                        </tr>
                    </thead>
                    <tbody>
                        {filas}
                    </tbody>
                </table>
            </div>
        </main>
    </body>
    </html>
    '''

# =========================
# STORAGE CRM PRODUCCION
# =========================

CRM_DATABASE_URL = os.getenv('DATABASE_URL', '').strip()
CRM_SQLITE_PATH = os.getenv('CRM_SQLITE_PATH', 'crm_storage.db')
CRM_STORAGE_TABLE = 'crm_storage_records'


def crm_storage_es_postgres():
    return bool(CRM_DATABASE_URL)


def crm_storage_placeholder():
    return '%s' if crm_storage_es_postgres() else '?'


def crm_storage_connect():
    if crm_storage_es_postgres():
        import psycopg2
        database_url = CRM_DATABASE_URL
        if database_url.startswith('postgres://'):
            database_url = database_url.replace('postgres://', 'postgresql://', 1)
        return psycopg2.connect(database_url)

    return sqlite3.connect(CRM_SQLITE_PATH)


def crm_storage_init():
    conn = crm_storage_connect()
    try:
        cur = conn.cursor()
        cur.execute(f"""
            CREATE TABLE IF NOT EXISTS {CRM_STORAGE_TABLE} (
                collection TEXT NOT NULL,
                position INTEGER NOT NULL,
                data TEXT NOT NULL,
                PRIMARY KEY (collection, position)
            )
        """)
        conn.commit()
    finally:
        conn.close()


def crm_storage_has_rows(collection):
    crm_storage_init()
    conn = crm_storage_connect()
    ph = crm_storage_placeholder()
    try:
        cur = conn.cursor()
        cur.execute(
            f'SELECT COUNT(*) FROM {CRM_STORAGE_TABLE} WHERE collection = {ph}',
            (collection,)
        )
        return int(cur.fetchone()[0]) > 0
    finally:
        conn.close()


def crm_storage_read(collection, headers):
    crm_storage_init()
    conn = crm_storage_connect()
    ph = crm_storage_placeholder()
    try:
        cur = conn.cursor()
        cur.execute(
            f'SELECT data FROM {CRM_STORAGE_TABLE} WHERE collection = {ph} ORDER BY position',
            (collection,)
        )
        rows = []
        for (data,) in cur.fetchall():
            try:
                row = json.loads(data)
            except (TypeError, json.JSONDecodeError):
                row = {}
            rows.append({header: row.get(header, '') for header in headers})
        return rows
    finally:
        conn.close()


def crm_storage_write(collection, headers, rows):
    crm_storage_init()
    conn = crm_storage_connect()
    ph = crm_storage_placeholder()
    try:
        cur = conn.cursor()
        cur.execute(
            f'DELETE FROM {CRM_STORAGE_TABLE} WHERE collection = {ph}',
            (collection,)
        )
        insert_sql = (
            f'INSERT INTO {CRM_STORAGE_TABLE} (collection, position, data) '
            f'VALUES ({ph}, {ph}, {ph})'
        )
        clean_rows = []
        for index, row in enumerate(rows, start=1):
            clean_row = {header: row.get(header, '') for header in headers}
            clean_rows.append((
                collection,
                index,
                json.dumps(clean_row, ensure_ascii=False)
            ))
        if clean_rows:
            cur.executemany(insert_sql, clean_rows)
        conn.commit()
    finally:
        conn.close()


def crm_importar_csv_si_existe(collection, archivo, headers):
    if crm_storage_has_rows(collection):
        return

    if not os.path.exists(archivo) or os.path.getsize(archivo) == 0:
        return

    with open(archivo, mode='r', newline='', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        rows = []
        for row in reader:
            rows.append({header: row.get(header, '') for header in headers})

    if rows:
        crm_storage_write(collection, headers, rows)
        print(f'Migracion CRM: {archivo} importado a almacenamiento persistente')

# =========================
# ACCESO PRIVADO CRM
# =========================

CRM_AUTH_CSV = 'crm_auth.csv'
CRM_AUTH_COLLECTION = 'crm_auth'
CRM_AUTH_HEADERS = ['username', 'password_hash', 'email_recordatorios', 'recordatorios_activos', 'updated_at']
CRM_DEFAULT_USERNAME = os.getenv('CRM_DEFAULT_USERNAME', 'silviamar23!')
CRM_DEFAULT_PASSWORD_HASH = os.getenv(
    'CRM_DEFAULT_PASSWORD_HASH',
    'scrypt:32768:8:1$QhCP0Eyj6q01FCUF$e60ed4d5a252b992b1e63259c69bec238d5ea3970a56d949414712b108e88c03c54aab442e9e0bd32038fb729a330c12f1c486cf0998e1e03b857824e9d3a5cb'
)
CRM_PUBLIC_PATHS = ['/crm/login']


def crm_auth_crear_default():
    auth = {
        'username': CRM_DEFAULT_USERNAME,
        'password_hash': CRM_DEFAULT_PASSWORD_HASH,
        'email_recordatorios': os.getenv('EMAIL_RECORDATORIOS', ''),
        'recordatorios_activos': os.getenv('CRM_RECORDATORIOS_ACTIVOS', 'No'),
        'updated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }

    crm_storage_write(CRM_AUTH_COLLECTION, CRM_AUTH_HEADERS, [auth])

    return auth


def crm_auth_leer():
    crm_importar_csv_si_existe(CRM_AUTH_COLLECTION, CRM_AUTH_CSV, CRM_AUTH_HEADERS)
    rows = crm_storage_read(CRM_AUTH_COLLECTION, CRM_AUTH_HEADERS)

    for row in rows:
        if row.get('username') and row.get('password_hash'):
            row.setdefault('email_recordatorios', os.getenv('EMAIL_RECORDATORIOS', ''))
            row.setdefault('recordatorios_activos', os.getenv('CRM_RECORDATORIOS_ACTIVOS', 'No'))
            return row

    return crm_auth_crear_default()


def crm_auth_guardar(username, password_hash, email_recordatorios='', recordatorios_activos='No'):
    auth = {
        'username': username,
        'password_hash': password_hash,
        'email_recordatorios': email_recordatorios,
        'recordatorios_activos': recordatorios_activos,
        'updated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }

    crm_storage_write(CRM_AUTH_COLLECTION, CRM_AUTH_HEADERS, [auth])

    return auth


def crm_sesion_expirada():
    ultima_actividad = session.get('crm_last_activity')

    if not ultima_actividad:
        return False

    try:
        ultima = datetime.fromisoformat(ultima_actividad)
    except ValueError:
        return True

    return datetime.now() - ultima > app.permanent_session_lifetime


def crm_verificar_recordatorios_globales():

    ahora = datetime.now()
    ultimo = app.config.get('CRM_RECORDATORIOS_ULTIMA_REVISION')

    if ultimo and (ahora - ultimo).total_seconds() < 60:
        return

    app.config['CRM_RECORDATORIOS_ULTIMA_REVISION'] = ahora

    try:
        tareas = crm_leer_csv(CRM_TAREAS_CSV, CRM_TAREA_HEADERS)
        if crm_actualizar_recordatorios_tareas(tareas):
            crm_escribir_csv(CRM_TAREAS_CSV, CRM_TAREA_HEADERS, tareas)
    except Exception as e:
        print(f'Error verificando recordatorios CRM: {e}')


@app.before_request
def revisar_recordatorios_crm():
    ruta = request.path or '/'

    if ruta.startswith('/static'):
        return None

    crm_verificar_recordatorios_globales()
    return None

@app.before_request
def proteger_crm_privado():
    ruta = request.path.rstrip('/') or '/'

    if not ruta.startswith('/crm'):
        return None

    if ruta in CRM_PUBLIC_PATHS:
        return None

    if not session.get('crm_authenticated'):
        return redirect(url_for('crm_login', next=request.full_path.rstrip('?')))

    if crm_sesion_expirada():
        session.clear()
        return redirect(url_for('crm_login', expirado='1'))

    session.permanent = True
    session['crm_last_activity'] = datetime.now().isoformat()
    return None


@app.route('/crm/login', methods=['GET', 'POST'])
def crm_login():
    error = ''
    expirado = request.args.get('expirado') == '1'

    if session.get('crm_authenticated') and not crm_sesion_expirada():
        return redirect(url_for('crm_dashboard'))

    if request.method == 'POST':
        usuario = request.form.get('usuario', '').strip()
        password = request.form.get('password', '')
        auth = crm_auth_leer()

        if usuario == auth.get('username') and check_password_hash(auth.get('password_hash', ''), password):
            session.clear()
            session.permanent = True
            session['crm_authenticated'] = True
            session['crm_user'] = usuario
            session['crm_last_activity'] = datetime.now().isoformat()
            destino = request.form.get('next') or url_for('crm_dashboard')
            return redirect(destino)

        error = 'Usuario o contraseña incorrectos.'

    return render_template(
        'crm/login.html',
        title='Acceso CRM',
        error=error,
        expirado=expirado,
        next_url=request.args.get('next', '')
    )


@app.route('/crm/logout')
def crm_logout():
    session.clear()
    return redirect(url_for('crm_login'))


@app.route('/crm/cuenta', methods=['GET', 'POST'])
def crm_cuenta():
    auth = crm_auth_leer()
    error = ''
    success = ''

    if request.method == 'POST':
        usuario = request.form.get('usuario', '').strip()
        password_actual = request.form.get('password_actual', '')
        password_nueva = request.form.get('password_nueva', '')
        password_confirmar = request.form.get('password_confirmar', '')
        email_recordatorios = request.form.get('email_recordatorios', '').strip()
        recordatorios_activos = 'Si' if request.form.get('recordatorios_activos') == 'Si' else 'No'

        if not check_password_hash(auth.get('password_hash', ''), password_actual):
            error = 'La contraseña actual no es correcta.'
        elif not usuario:
            error = 'El usuario no puede quedar vacío.'
        elif password_nueva and password_nueva != password_confirmar:
            error = 'La nueva contraseña y la confirmación no coinciden.'
        else:
            password_hash = auth.get('password_hash', '')
            if password_nueva:
                password_hash = generate_password_hash(password_nueva)

            auth = crm_auth_guardar(usuario, password_hash, email_recordatorios, recordatorios_activos)
            session['crm_user'] = usuario
            success = 'Cuenta actualizada correctamente.'

    return render_template(
        'crm/cuenta.html',
        active='cuenta',
        title='Mi Cuenta',
        usuario=auth.get('username', ''),
        email_recordatorios=auth.get('email_recordatorios', ''),
        recordatorios_activos=auth.get('recordatorios_activos', 'No'),
        error=error,
        success=success
    )

# =========================
# CRM INMOBILIARIO
# =========================

CRM_OPORTUNIDADES_CSV = 'crm_oportunidades.csv'
CRM_TAREAS_CSV = 'crm_tareas.csv'
CRM_CLIENTES_CSV = 'crm_clientes.csv'
CRM_AGENDA_CSV = 'crm_agenda.csv'
CRM_LEAD_NOTAS_CSV = 'crm_lead_notas.csv'

CRM_STORAGE_COLLECTIONS = {
    CRM_OPORTUNIDADES_CSV: 'crm_oportunidades',
    CRM_TAREAS_CSV: 'crm_tareas',
    CRM_CLIENTES_CSV: 'crm_clientes',
    CRM_AGENDA_CSV: 'crm_agenda',
    CRM_LEAD_NOTAS_CSV: 'crm_lead_notas',
}

CRM_OPORTUNIDAD_HEADERS = [
    'id',
    'nombre',
    'telefono',
    'email',
    'propiedad',
    'fuente',
    'etapa',
    'valor',
    'fecha_creacion',
    'notas',
    'historial'
]

CRM_TAREA_HEADERS = [
    'id',
    'titulo',
    'cliente',
    'propiedad',
    'fecha_limite',
    'hora_limite',
    'prioridad',
    'estado',
    'recordatorio',
    'notas',
    'fecha_creacion',
    'recordatorio_enviado'
]

CRM_CLIENTE_HEADERS = [
    'id',
    'nombre',
    'telefono',
    'email',
    'tipo',
    'propiedad',
    'fuente',
    'historial',
    'notas',
    'fecha_creacion'
]

CRM_AGENDA_HEADERS = [
    'id',
    'fecha',
    'hora',
    'cliente',
    'telefono',
    'propiedad',
    'estado',
    'notas',
    'fecha_creacion'
]

CRM_LEAD_NOTA_HEADERS = [
    'id',
    'lead_row',
    'tipo',
    'contenido',
    'fecha'
]

CRM_ETAPAS_OPORTUNIDAD = [
    'Cultivar',
    'Cita',
    'Activo',
    'Oferta',
    'Bajo contrato',
    'Cerrado'
]

CRM_PRIORIDADES = [
    'Alta',
    'Media',
    'Baja'
]

CRM_ESTADOS_TAREA = [
    'Pendiente',
    'En curso',
    'Completada',
    'Vencida'
]

CRM_ESTADOS_VISITA = [
    'Agendada',
    'Confirmada',
    'Realizada',
    'Cancelada'
]

def crm_now():

    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

def crm_generar_id(prefix):

    return f"{prefix}-{datetime.now().strftime('%Y%m%d%H%M%S%f')}"

def crm_storage_collection(archivo):

    return CRM_STORAGE_COLLECTIONS.get(archivo, archivo.replace('.csv', ''))

def crm_leer_csv(archivo, headers):

    collection = crm_storage_collection(archivo)
    crm_importar_csv_si_existe(collection, archivo, headers)
    return crm_storage_read(collection, headers)

def crm_escribir_csv(archivo, headers, rows):

    collection = crm_storage_collection(archivo)
    crm_storage_write(collection, headers, rows)

def crm_agregar_csv(archivo, headers, row):

    rows = crm_leer_csv(archivo, headers)
    rows.append({
        header: row.get(header, '')
        for header in headers
    })
    crm_escribir_csv(archivo, headers, rows)

def crm_parse_fecha(valor):

    if not valor:
        return None

    formatos = [
        '%Y-%m-%d',
        '%Y-%m-%d %H:%M:%S',
        '%Y-%m-%dT%H:%M:%SZ'
    ]

    for formato in formatos:
        try:
            return datetime.strptime(valor[:19], formato)
        except ValueError:
            continue

    return None

def crm_parse_fecha_hora_tarea(tarea):

    fecha = tarea.get('fecha_limite', '').strip()
    hora = tarea.get('hora_limite', '').strip()

    if not fecha:
        return None

    if hora:
        try:
            return datetime.strptime(f'{fecha} {hora}', '%Y-%m-%d %H:%M')
        except ValueError:
            pass

    fecha_dt = crm_parse_fecha(fecha)

    if fecha_dt:
        return fecha_dt.replace(hour=23, minute=59, second=0, microsecond=0)

    return None

def enviar_email_recordatorio_tarea(tarea):

    auth = crm_auth_leer()

    if auth.get('recordatorios_activos') != 'Si':
        return False

    email_to = auth.get('email_recordatorios') or os.getenv('EMAIL_RECORDATORIOS')

    if not email_to:
        return False

    import smtplib
    from email.message import EmailMessage

    smtp_host = os.getenv('SMTP_HOST', '').strip()
    smtp_port_raw = os.getenv('SMTP_PORT', '587').strip() or '587'
    smtp_user = os.getenv('SMTP_USER', '').strip()
    smtp_password = os.getenv('SMTP_PASSWORD', '').strip()
    email_from = (os.getenv('EMAIL_FROM') or smtp_user).strip()

    try:
        smtp_port = int(smtp_port_raw)
    except (TypeError, ValueError):
        print(f'Advertencia CRM: SMTP_PORT inválido ({smtp_port_raw}). No se envió email de recordatorio.')
        return False

    if not smtp_host or not smtp_user or not smtp_password:
        print('Advertencia CRM: faltan variables SMTP. No se envió email de recordatorio.')
        return False

    body = f"""
Recordatorio de tarea CRM.

Nombre de la tarea: {tarea.get('titulo', '')}
Cliente: {tarea.get('cliente', '')}
Propiedad: {tarea.get('propiedad', '')}
Fecha: {tarea.get('fecha_limite', '')}
Hora: {tarea.get('hora_limite', '')}
"""

    msg = EmailMessage()
    msg['Subject'] = 'Recordatorio CRM - Tarea pendiente'
    msg['From'] = email_from
    msg['To'] = email_to
    msg.set_content(body)

    try:
        with smtplib.SMTP(smtp_host, smtp_port, timeout=15) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.send_message(msg)
        print('Email de recordatorio enviado correctamente')
        return True
    except Exception as e:
        print(f'Error enviando recordatorio CRM: {e}')
        return False
def crm_actualizar_recordatorios_tareas(tareas):

    ahora = datetime.now()
    cambios = False

    for tarea in tareas:
        if tarea.get('estado') == 'Completada':
            continue

        vencimiento = crm_parse_fecha_hora_tarea(tarea)

        if not vencimiento:
            continue

        if vencimiento <= ahora:
            if tarea.get('recordatorio_enviado') != 'Si':
                if enviar_email_recordatorio_tarea(tarea):
                    tarea['recordatorio_enviado'] = 'Si'
                    cambios = True

            if tarea.get('estado') != 'Vencida':
                tarea['estado'] = 'Vencida'
                cambios = True

    return cambios

def crm_tareas_por_vencimiento(tareas):

    ahora = datetime.now()
    en_una_hora = ahora + timedelta(hours=1)
    hoy = ahora.date()
    grupos = {
        'vence_ahora': [],
        'proxima_hora': [],
        'hoy': [],
        'vencidas': []
    }

    for tarea in tareas:
        if tarea.get('estado') == 'Completada':
            continue

        vencimiento = crm_parse_fecha_hora_tarea(tarea)

        if not vencimiento:
            continue

        if vencimiento <= ahora and vencimiento >= ahora - timedelta(minutes=15):
            grupos['vence_ahora'].append(tarea)
        elif vencimiento > ahora and vencimiento <= en_una_hora:
            grupos['proxima_hora'].append(tarea)
        elif vencimiento.date() == hoy and vencimiento > ahora:
            grupos['hoy'].append(tarea)
        elif vencimiento < ahora:
            grupos['vencidas'].append(tarea)

    return grupos

def crm_obtener_leads():

    headers, rows = obtener_interesados_datos()
    leads = []

    for index, row in enumerate(rows, start=2):
        lead = lead_desde_fila(headers, row)
        lead['row_number'] = index
        estado = lead.get('estado') or 'Nuevo'
        fuente = lead.get('fuente') or 'Web'
        lead['prioridad'] = 'Alta' if estado == 'Nuevo' else 'Media'
        lead['etiquetas'] = [fuente, estado]
        leads.append(lead)

    leads.sort(
        key=lambda item: item.get('fecha', ''),
        reverse=True
    )
    return leads

def crm_leads_sin_contactar(leads):

    return [
        lead
        for lead in leads
        if (lead.get('estado') or 'Nuevo') == 'Nuevo'
    ]

def crm_seguimientos_atrasados(tareas):

    ahora = datetime.now()
    atrasadas = []

    for tarea in tareas:
        fecha = crm_parse_fecha_hora_tarea(tarea)

        if (
            fecha
            and fecha < ahora
            and tarea.get('estado') != 'Completada'
        ):
            atrasadas.append(tarea)

    return atrasadas

def crm_tareas_hoy(tareas):

    hoy = datetime.now().date()
    return [
        tarea
        for tarea in tareas
        if (
            crm_parse_fecha_hora_tarea(tarea)
            and crm_parse_fecha_hora_tarea(tarea).date() == hoy
            and tarea.get('estado') != 'Completada'
        )
    ]

def crm_visitas_hoy(agenda):

    hoy = datetime.now().date()
    return [
        visita
        for visita in agenda
        if (
            crm_parse_fecha(visita.get('fecha'))
            and crm_parse_fecha(visita.get('fecha')).date() == hoy
            and visita.get('estado') in ['', 'Agendada', 'Confirmada']
        )
    ]

def crm_visitas_proximas(agenda):

    hoy = datetime.now().date()
    visitas = []

    for visita in agenda:
        fecha = crm_parse_fecha(visita.get('fecha'))

        if (
            fecha
            and fecha.date() >= hoy
            and visita.get('estado') in ['', 'Agendada', 'Confirmada']
        ):
            visitas.append(visita)

    return visitas

def crm_lead_por_row(row_number):

    leads = crm_obtener_leads()

    for lead in leads:
        if str(lead.get('row_number')) == str(row_number):
            return lead

    return None

def crm_notas_lead(row_number):

    notas = crm_leer_csv(CRM_LEAD_NOTAS_CSV, CRM_LEAD_NOTA_HEADERS)
    return [
        nota
        for nota in notas
        if str(nota.get('lead_row')) == str(row_number)
    ]

def crm_tareas_asociadas_lead(lead):

    tareas = crm_leer_csv(CRM_TAREAS_CSV, CRM_TAREA_HEADERS)
    nombre = (lead.get('nombre') or '').lower()
    propiedad = (lead.get('propiedad') or '').lower()

    return [
        tarea
        for tarea in tareas
        if (
            nombre
            and nombre in (tarea.get('cliente') or '').lower()
        ) or (
            propiedad
            and propiedad in (tarea.get('propiedad') or '').lower()
        )
    ]

def crm_oportunidades_asociadas(lead):

    oportunidades = crm_leer_csv(CRM_OPORTUNIDADES_CSV, CRM_OPORTUNIDAD_HEADERS)
    nombre = (lead.get('nombre') or '').lower()
    propiedad = (lead.get('propiedad') or '').lower()

    return [
        oportunidad
        for oportunidad in oportunidades
        if (
            nombre
            and nombre in (oportunidad.get('nombre') or '').lower()
        ) or (
            propiedad
            and propiedad in (oportunidad.get('propiedad') or '').lower()
        )
    ]

def crm_cliente_por_id(cliente_id):

    clientes = crm_leer_csv(CRM_CLIENTES_CSV, CRM_CLIENTE_HEADERS)

    for cliente in clientes:
        if cliente.get('id') == cliente_id:
            return cliente

    return None

def crm_tareas_asociadas_cliente(cliente):

    tareas = crm_leer_csv(CRM_TAREAS_CSV, CRM_TAREA_HEADERS)
    nombre = (cliente.get('nombre') or '').lower()
    propiedad = (cliente.get('propiedad') or '').lower()

    return [
        tarea
        for tarea in tareas
        if (
            nombre
            and nombre in (tarea.get('cliente') or '').lower()
        ) or (
            propiedad
            and propiedad in (tarea.get('propiedad') or '').lower()
        )
    ]

def crm_tareas_segmentadas(tareas):

    hoy = datetime.now().date()
    grupos = {
        'hoy': [],
        'proximas': [],
        'vencidas': [],
        'completadas': []
    }

    for tarea in tareas:
        estado = tarea.get('estado') or 'Pendiente'
        fecha = crm_parse_fecha_hora_tarea(tarea)

        if estado == 'Completada':
            grupos['completadas'].append(tarea)
        elif estado == 'Vencida':
            grupos['vencidas'].append(tarea)
        elif fecha and fecha.date() < hoy:
            grupos['vencidas'].append(tarea)
        elif fecha and fecha.date() == hoy:
            grupos['hoy'].append(tarea)
        else:
            grupos['proximas'].append(tarea)

    return grupos

def crm_agenda_por_fecha(agenda):

    grupos = {}

    for visita in agenda:
        fecha = visita.get('fecha') or 'Sin fecha'

        if fecha not in grupos:
            grupos[fecha] = []

        grupos[fecha].append(visita)

    return grupos

def crm_tareas_proximas(tareas, limite=5):

    ahora = datetime.now()
    pendientes = []

    for tarea in tareas:
        if tarea.get('estado') == 'Completada':
            continue

        fecha = crm_parse_fecha_hora_tarea(tarea)

        if fecha and fecha >= ahora:
            pendientes.append((fecha, tarea))

    pendientes.sort(key=lambda item: item[0])
    return [tarea for _, tarea in pendientes[:limite]]


def crm_cliente_ultimo_contacto(cliente, tareas):

    fechas = []
    creado = crm_parse_fecha(cliente.get('fecha_creacion'))

    if creado:
        fechas.append(creado)

    nombre = (cliente.get('nombre') or '').lower()
    propiedad = (cliente.get('propiedad') or '').lower()

    for tarea in tareas:
        coincide = (
            nombre and nombre in (tarea.get('cliente') or '').lower()
        ) or (
            propiedad and propiedad in (tarea.get('propiedad') or '').lower()
        )

        if coincide:
            fecha = crm_parse_fecha(tarea.get('fecha_creacion')) or crm_parse_fecha_hora_tarea(tarea)
            if fecha:
                fechas.append(fecha)

    return max(fechas) if fechas else None


def crm_clientes_sin_contacto(clientes, tareas, dias=7):

    limite = datetime.now() - timedelta(days=dias)
    sin_contacto = []

    for cliente in clientes:
        ultimo = crm_cliente_ultimo_contacto(cliente, tareas)

        if ultimo and ultimo < limite:
            sin_contacto.append(cliente)

    return sin_contacto


def crm_leads_sin_seguimiento(leads, tareas):

    sin_seguimiento = []

    for lead in leads:
        nombre = (lead.get('nombre') or '').lower()
        propiedad = (lead.get('propiedad') or '').lower()
        tiene_tarea = any(
            (nombre and nombre in (tarea.get('cliente') or '').lower())
            or (propiedad and propiedad in (tarea.get('propiedad') or '').lower())
            for tarea in tareas
        )

        if (lead.get('estado') or 'Nuevo') != 'Cerrado' and not tiene_tarea:
            sin_seguimiento.append(lead)

    return sin_seguimiento


def crm_visitas_pendientes(agenda):

    hoy = datetime.now().date()
    pendientes = []

    for visita in agenda:
        fecha = crm_parse_fecha(visita.get('fecha'))
        estado = visita.get('estado') or 'Agendada'

        if fecha and fecha.date() >= hoy and estado in ['', 'Agendada', 'Confirmada']:
            pendientes.append(visita)

    return pendientes


def crm_tareas_por_prioridad_notificacion(tareas):

    grupos = {'Alta': [], 'Media': [], 'Baja': []}

    for tarea in tareas:
        if tarea.get('estado') == 'Completada':
            continue

        prioridad = tarea.get('prioridad') or 'Media'
        if prioridad not in grupos:
            prioridad = 'Media'

        grupos[prioridad].append(tarea)

    for prioridad in grupos:
        grupos[prioridad] = sorted(
            grupos[prioridad],
            key=lambda tarea: crm_parse_fecha_hora_tarea(tarea) or datetime.max
        )

    return grupos

def crm_dashboard_context():

    leads = crm_obtener_leads()
    oportunidades = crm_leer_csv(CRM_OPORTUNIDADES_CSV, CRM_OPORTUNIDAD_HEADERS)
    tareas = crm_leer_csv(CRM_TAREAS_CSV, CRM_TAREA_HEADERS)
    clientes = crm_leer_csv(CRM_CLIENTES_CSV, CRM_CLIENTE_HEADERS)
    agenda = crm_leer_csv(CRM_AGENDA_CSV, CRM_AGENDA_HEADERS)

    leads_nuevos = [
        lead
        for lead in leads
        if (lead.get('estado') or 'Nuevo') == 'Nuevo'
    ]
    visitas_proximas = crm_visitas_proximas(agenda)
    if crm_actualizar_recordatorios_tareas(tareas):
        crm_escribir_csv(CRM_TAREAS_CSV, CRM_TAREA_HEADERS, tareas)

    tareas_atrasadas = crm_seguimientos_atrasados(tareas)
    recordatorios_tareas = crm_tareas_por_vencimiento(tareas)
    oportunidades_activas = [
        oportunidad
        for oportunidad in oportunidades
        if oportunidad.get('etapa') != 'Cerrado'
    ]
    proximas_tareas = crm_tareas_proximas(tareas)
    clientes_sin_contacto = crm_clientes_sin_contacto(clientes, tareas)
    leads_sin_seguimiento = crm_leads_sin_seguimiento(leads, tareas)
    visitas_pendientes = crm_visitas_pendientes(agenda)
    tareas_por_prioridad = crm_tareas_por_prioridad_notificacion(tareas)

    fuentes = {}
    for lead in leads:
        fuente = lead.get('fuente') or 'Web'
        fuentes[fuente] = fuentes.get(fuente, 0) + 1

    etapas = {}
    for etapa in CRM_ETAPAS_OPORTUNIDAD:
        etapas[etapa] = 0

    for oportunidad in oportunidades:
        etapa = oportunidad.get('etapa') or 'Cultivar'
        etapas[etapa] = etapas.get(etapa, 0) + 1

    return {
        'leads': leads,
        'leads_nuevos': leads_nuevos,
        'leads_sin_contactar': crm_leads_sin_contactar(leads),
        'visitas_proximas': visitas_proximas,
        'visitas_hoy': crm_visitas_hoy(agenda),
        'tareas_hoy': crm_tareas_hoy(tareas),
        'recordatorios_tareas': recordatorios_tareas,
        'tareas_vence_ahora': recordatorios_tareas['vence_ahora'],
        'tareas_proxima_hora': recordatorios_tareas['proxima_hora'],
        'proximas_tareas': proximas_tareas,
        'clientes_sin_contacto': clientes_sin_contacto,
        'leads_sin_seguimiento': leads_sin_seguimiento,
        'visitas_pendientes': visitas_pendientes,
        'tareas_por_prioridad': tareas_por_prioridad,
        'seguimientos_atrasados': tareas_atrasadas,
        'oportunidades': oportunidades,
        'oportunidades_activas': oportunidades_activas,
        'tareas': tareas,
        'clientes': clientes,
        'agenda': agenda,
        'fuentes': fuentes,
        'etapas': etapas
    }

def crm_notificaciones_context():

    context = crm_dashboard_context()
    notificaciones = []

    for lead in context['leads_nuevos'][:8]:
        notificaciones.append({
            'tipo': 'Lead nuevo',
            'titulo': lead.get('nombre') or 'Nuevo contacto',
            'detalle': lead.get('propiedad') or lead.get('fuente') or 'Sin propiedad',
            'fecha': lead.get('fecha', ''),
            'nivel': 'info'
        })

    for tarea in context['tareas_hoy'][:8]:
        notificaciones.append({
            'tipo': 'Tarea vence hoy',
            'titulo': tarea.get('titulo') or 'Seguimiento pendiente',
            'detalle': tarea.get('cliente') or tarea.get('propiedad') or 'Sin cliente',
            'fecha': tarea.get('fecha_limite', ''),
            'nivel': 'warning'
        })

    for tarea in context['seguimientos_atrasados'][:8]:
        notificaciones.append({
            'tipo': 'Tarea vencida',
            'titulo': tarea.get('titulo') or 'Seguimiento pendiente',
            'detalle': tarea.get('cliente') or tarea.get('propiedad') or 'Sin cliente',
            'fecha': tarea.get('fecha_limite', ''),
            'nivel': 'danger'
        })

    for visita in context['visitas_hoy'][:8]:
        notificaciones.append({
            'tipo': 'Visita hoy',
            'titulo': visita.get('cliente') or 'Visita agendada',
            'detalle': f"{visita.get('propiedad', '')} {visita.get('hora', '')}".strip(),
            'fecha': visita.get('fecha', ''),
            'nivel': 'success'
        })

    for visita in context['visitas_proximas'][:8]:
        notificaciones.append({
            'tipo': 'Visita próxima',
            'titulo': visita.get('cliente') or 'Visita agendada',
            'detalle': f"{visita.get('propiedad', '')} {visita.get('hora', '')}".strip(),
            'fecha': visita.get('fecha', ''),
            'nivel': 'success'
        })

    return notificaciones

@app.route('/crm')
def crm_home():

    return crm_dashboard()

@app.route('/crm/dashboard')
def crm_dashboard():

    context = crm_dashboard_context()
    return render_template(
        'crm/dashboard.html',
        active='dashboard',
        title='Dashboard',
        **context
    )

@app.route('/crm/leads')
def crm_leads():

    leads = crm_obtener_leads()
    fuente = request.args.get('fuente', '').strip()
    estado = request.args.get('estado', '').strip()
    busqueda = request.args.get('q', '').strip().lower()
    prioridad = request.args.get('prioridad', '').strip()
    orden = request.args.get('orden', '').strip() or 'recientes'

    if fuente:
        leads = [
            lead
            for lead in leads
            if (lead.get('fuente') or 'Web') == fuente
        ]

    if estado:
        leads = [
            lead
            for lead in leads
            if (lead.get('estado') or 'Nuevo') == estado
        ]

    if prioridad:
        leads = [
            lead
            for lead in leads
            if lead.get('prioridad') == prioridad
        ]

    if busqueda:
        leads = [
            lead
            for lead in leads
            if busqueda in ' '.join([
                lead.get('nombre', ''),
                lead.get('telefono', ''),
                lead.get('email', ''),
                lead.get('propiedad', ''),
                lead.get('mensaje', '')
            ]).lower()
        ]

    if orden == 'nombre':
        leads.sort(key=lambda lead: (lead.get('nombre') or '').lower())
    elif orden == 'fuente':
        leads.sort(key=lambda lead: (lead.get('fuente') or '').lower())
    else:
        leads.sort(key=lambda lead: lead.get('fecha', ''), reverse=True)

    return render_template(
        'crm/leads.html',
        active='leads',
        title='Leads',
        leads=leads,
        fuentes=FUENTES_LEAD,
        estados=ESTADOS_LEAD,
        filtro_fuente=fuente,
        filtro_estado=estado,
        filtro_prioridad=prioridad,
        orden=orden,
        busqueda=busqueda
    )

@app.route('/crm/lead/<int:row_number>', methods=['GET', 'POST'])
def crm_lead_perfil(row_number):

    lead = crm_lead_por_row(row_number)

    if not lead:
        return "Lead no encontrado.", 404

    if request.method == 'POST':
        action = request.form.get('action', '').strip() or 'crear_nota'

        if action == 'crear_tarea':
            tarea = {
                'id': crm_generar_id('task'),
                'titulo': request.form.get('titulo', '').strip() or f"Seguimiento: {lead.get('nombre') or 'Lead'}",
                'cliente': lead.get('nombre', ''),
                'propiedad': lead.get('propiedad', ''),
                'fecha_limite': request.form.get('fecha_limite', '').strip(),
                'hora_limite': request.form.get('hora_limite', '').strip(),
                'prioridad': request.form.get('prioridad', '').strip() or 'Media',
                'estado': 'Pendiente',
                'recordatorio': request.form.get('recordatorio', '').strip(),
                'notas': request.form.get('notas', '').strip(),
                'fecha_creacion': crm_now(),
                'recordatorio_enviado': ''
            }
            crm_agregar_csv(CRM_TAREAS_CSV, CRM_TAREA_HEADERS, tarea)
            crm_agregar_csv(CRM_LEAD_NOTAS_CSV, CRM_LEAD_NOTA_HEADERS, {
                'id': crm_generar_id('note'),
                'lead_row': str(row_number),
                'tipo': 'Tarea',
                'contenido': f"Tarea creada: {tarea['titulo']} para {tarea['fecha_limite']}",
                'fecha': crm_now()
            })

        elif action == 'agendar_visita':
            visita = {
                'id': crm_generar_id('visit'),
                'fecha': request.form.get('fecha', '').strip(),
                'hora': request.form.get('hora', '').strip(),
                'cliente': lead.get('nombre', ''),
                'telefono': lead.get('telefono', ''),
                'propiedad': lead.get('propiedad', ''),
                'estado': 'Agendada',
                'notas': request.form.get('notas', '').strip(),
                'fecha_creacion': crm_now()
            }
            crm_agregar_csv(CRM_AGENDA_CSV, CRM_AGENDA_HEADERS, visita)
            crm_agregar_csv(CRM_LEAD_NOTAS_CSV, CRM_LEAD_NOTA_HEADERS, {
                'id': crm_generar_id('note'),
                'lead_row': str(row_number),
                'tipo': 'Visita',
                'contenido': f"Visita agendada: {visita['fecha']} {visita['hora']}",
                'fecha': crm_now()
            })

        elif action == 'registrar_llamada':
            resultado = request.form.get('resultado', '').strip() or 'Sin resultado'
            observacion = request.form.get('observacion', '').strip()
            crm_agregar_csv(CRM_LEAD_NOTAS_CSV, CRM_LEAD_NOTA_HEADERS, {
                'id': crm_generar_id('note'),
                'lead_row': str(row_number),
                'tipo': 'Llamada',
                'contenido': f"Resultado: {resultado}. {observacion}".strip(),
                'fecha': crm_now()
            })

        elif action == 'convertir_oportunidad':
            notas = crm_notas_lead(row_number)
            historial = ' | '.join([
                f"{nota.get('fecha', '')} {nota.get('tipo', '')}: {nota.get('contenido', '')}"
                for nota in notas
            ])
            oportunidad = {
                'id': crm_generar_id('opp'),
                'nombre': lead.get('nombre', ''),
                'telefono': lead.get('telefono', ''),
                'email': lead.get('email', ''),
                'propiedad': lead.get('propiedad', ''),
                'fuente': lead.get('fuente', '') or 'Web',
                'etapa': request.form.get('etapa', '').strip() or 'Cultivar',
                'valor': request.form.get('valor', '').strip(),
                'fecha_creacion': crm_now(),
                'notas': request.form.get('notas', '').strip() or lead.get('mensaje', ''),
                'historial': historial
            }
            crm_agregar_csv(CRM_OPORTUNIDADES_CSV, CRM_OPORTUNIDAD_HEADERS, oportunidad)
            actualizar_estado_interesado_storage(row_number, 'Negociación')
            crm_agregar_csv(CRM_LEAD_NOTAS_CSV, CRM_LEAD_NOTA_HEADERS, {
                'id': crm_generar_id('note'),
                'lead_row': str(row_number),
                'tipo': 'Oportunidad',
                'contenido': f"Lead convertido a oportunidad en etapa {oportunidad['etapa']}",
                'fecha': crm_now()
            })

        else:
            nota = {
                'id': crm_generar_id('note'),
                'lead_row': str(row_number),
                'tipo': request.form.get('tipo', '').strip() or 'Nota',
                'contenido': request.form.get('contenido', '').strip(),
                'fecha': crm_now()
            }

            if nota['contenido']:
                crm_agregar_csv(CRM_LEAD_NOTAS_CSV, CRM_LEAD_NOTA_HEADERS, nota)

        return redirect(url_for('crm_lead_perfil', row_number=row_number))

    return render_template(
        'crm/lead_perfil.html',
        active='leads',
        title='Perfil de lead',
        lead=lead,
        notas=crm_notas_lead(row_number),
        tareas=crm_tareas_asociadas_lead(lead),
        oportunidades=crm_oportunidades_asociadas(lead),
        etapas=CRM_ETAPAS_OPORTUNIDAD,
        prioridades=CRM_PRIORIDADES
    )

@app.route('/crm/oportunidades', methods=['GET', 'POST'])
def crm_oportunidades():

    if request.method == 'POST':
        action = request.form.get('action', '').strip()
        oportunidades = crm_leer_csv(CRM_OPORTUNIDADES_CSV, CRM_OPORTUNIDAD_HEADERS)

        if action == 'update_etapa':
            oportunidad_id = request.form.get('id', '').strip()
            etapa = request.form.get('etapa', '').strip()

            if etapa in CRM_ETAPAS_OPORTUNIDAD:
                for oportunidad in oportunidades:
                    if oportunidad.get('id') == oportunidad_id:
                        oportunidad['etapa'] = etapa
                        break

                crm_escribir_csv(
                    CRM_OPORTUNIDADES_CSV,
                    CRM_OPORTUNIDAD_HEADERS,
                    oportunidades
                )

            return jsonify({'ok': True})

        oportunidad = {
            'id': crm_generar_id('opp'),
            'nombre': request.form.get('nombre', '').strip(),
            'telefono': request.form.get('telefono', '').strip(),
            'email': request.form.get('email', '').strip(),
            'propiedad': request.form.get('propiedad', '').strip(),
            'fuente': request.form.get('fuente', '').strip() or 'Manual',
            'etapa': request.form.get('etapa', '').strip() or 'Cultivar',
            'valor': request.form.get('valor', '').strip(),
            'fecha_creacion': crm_now(),
            'notas': request.form.get('notas', '').strip()
        }
        crm_agregar_csv(
            CRM_OPORTUNIDADES_CSV,
            CRM_OPORTUNIDAD_HEADERS,
            oportunidad
        )

    oportunidades = crm_leer_csv(CRM_OPORTUNIDADES_CSV, CRM_OPORTUNIDAD_HEADERS)
    tablero = {
        etapa: []
        for etapa in CRM_ETAPAS_OPORTUNIDAD
    }

    for oportunidad in oportunidades:
        etapa = oportunidad.get('etapa') or 'Cultivar'

        if etapa not in tablero:
            tablero[etapa] = []

        tablero[etapa].append(oportunidad)

    return render_template(
        'crm/oportunidades.html',
        active='oportunidades',
        title='Oportunidades',
        etapas=CRM_ETAPAS_OPORTUNIDAD,
        tablero=tablero,
        fuentes=FUENTES_LEAD
    )

@app.route('/crm/tareas', methods=['GET', 'POST'])
def crm_tareas():

    if request.method == 'POST':
        action = request.form.get('action', '').strip() or 'create'
        tareas = crm_leer_csv(CRM_TAREAS_CSV, CRM_TAREA_HEADERS)
        tarea_id = request.form.get('id', '').strip()

        if action == 'delete':
            tareas = [tarea for tarea in tareas if tarea.get('id') != tarea_id]
            crm_escribir_csv(CRM_TAREAS_CSV, CRM_TAREA_HEADERS, tareas)
            return redirect(url_for('crm_tareas'))

        if action in ['complete', 'update']:
            for tarea in tareas:
                if tarea.get('id') == tarea_id:
                    if action == 'complete':
                        tarea['estado'] = 'Completada'
                    else:
                        tarea['titulo'] = request.form.get('titulo', '').strip()
                        tarea['cliente'] = request.form.get('cliente', '').strip()
                        tarea['propiedad'] = request.form.get('propiedad', '').strip()
                        tarea['fecha_limite'] = request.form.get('fecha_limite', '').strip()
                        tarea['hora_limite'] = request.form.get('hora_limite', '').strip()
                        tarea['prioridad'] = request.form.get('prioridad', '').strip() or 'Media'
                        tarea['estado'] = request.form.get('estado', '').strip() or 'Pendiente'
                        tarea['recordatorio'] = request.form.get('recordatorio', '').strip()
                        tarea['notas'] = request.form.get('notas', '').strip()
                    break
            crm_escribir_csv(CRM_TAREAS_CSV, CRM_TAREA_HEADERS, tareas)
            return redirect(url_for('crm_tareas'))

        tarea = {
            'id': crm_generar_id('task'),
            'titulo': request.form.get('titulo', '').strip(),
            'cliente': request.form.get('cliente', '').strip(),
            'propiedad': request.form.get('propiedad', '').strip(),
            'fecha_limite': request.form.get('fecha_limite', '').strip(),
            'hora_limite': request.form.get('hora_limite', '').strip(),
            'prioridad': request.form.get('prioridad', '').strip() or 'Media',
            'estado': request.form.get('estado', '').strip() or 'Pendiente',
            'recordatorio': request.form.get('recordatorio', '').strip(),
            'notas': request.form.get('notas', '').strip(),
            'fecha_creacion': crm_now(),
            'recordatorio_enviado': ''
        }
        crm_agregar_csv(CRM_TAREAS_CSV, CRM_TAREA_HEADERS, tarea)
        return redirect(url_for('crm_tareas'))

    tareas = crm_leer_csv(CRM_TAREAS_CSV, CRM_TAREA_HEADERS)
    if crm_actualizar_recordatorios_tareas(tareas):
        crm_escribir_csv(CRM_TAREAS_CSV, CRM_TAREA_HEADERS, tareas)
    atrasadas = {
        tarea.get('id')
        for tarea in crm_seguimientos_atrasados(tareas)
    }
    grupos = crm_tareas_segmentadas(tareas)

    return render_template(
        'crm/tareas.html',
        active='tareas',
        title='Tareas',
        tareas=tareas,
        grupos=grupos,
        atrasadas=atrasadas,
        prioridades=CRM_PRIORIDADES,
        estados=CRM_ESTADOS_TAREA
    )

def crm_normalizar_contacto_valor(valor):

    return ''.join(ch for ch in (valor or '') if ch.isdigit() or ch == '+').strip()

def crm_contacto_duplicado(contacto, clientes):

    telefono = crm_normalizar_contacto_valor(contacto.get('telefono'))
    email = (contacto.get('email') or '').strip().lower()

    for cliente in clientes:
        cliente_telefono = crm_normalizar_contacto_valor(cliente.get('telefono'))
        cliente_email = (cliente.get('email') or '').strip().lower()

        if telefono and cliente_telefono and telefono == cliente_telefono:
            return True

        if email and cliente_email and email == cliente_email:
            return True

    return False

def crm_extraer_contactos_csv(contenido):

    contactos = []
    reader = csv.DictReader(io.StringIO(contenido))

    for row in reader:
        normal = {str(k or '').strip().lower(): (v or '').strip() for k, v in row.items()}
        nombre = (
            normal.get('nombre')
            or normal.get('name')
            or normal.get('full name')
            or normal.get('full_name')
            or ' '.join(filter(None, [normal.get('first name'), normal.get('last name')]))
            or ' '.join(filter(None, [normal.get('nombre de pila'), normal.get('apellidos')]))
        ).strip()
        telefono = (
            normal.get('telefono')
            or normal.get('teléfono')
            or normal.get('phone')
            or normal.get('mobile')
            or normal.get('celular')
            or normal.get('phone 1 - value')
            or normal.get('phone value')
        ).strip()
        email = (
            normal.get('email')
            or normal.get('e-mail')
            or normal.get('correo')
            or normal.get('email 1 - value')
            or normal.get('e-mail 1 - value')
        ).strip()

        if nombre or telefono or email:
            contactos.append({
                'nombre': nombre or telefono or email,
                'telefono': telefono,
                'email': email
            })

    return contactos

def crm_vcf_unescape(valor):

    return (
        (valor or '')
        .replace('\\n', ' ')
        .replace('\\,', ',')
        .replace('\\;', ';')
        .strip()
    )

def crm_extraer_contactos_vcf(contenido):

    contactos = []
    actual = {}
    lineas = []

    for linea in contenido.splitlines():
        if linea.startswith((' ', '\t')) and lineas:
            lineas[-1] += linea[1:]
        else:
            lineas.append(linea)

    for linea in lineas:
        limpia = linea.strip()
        upper = limpia.upper()

        if upper == 'BEGIN:VCARD':
            actual = {}
        elif upper == 'END:VCARD':
            if actual.get('nombre') or actual.get('telefono') or actual.get('email'):
                contactos.append({
                    'nombre': actual.get('nombre') or actual.get('telefono') or actual.get('email'),
                    'telefono': actual.get('telefono', ''),
                    'email': actual.get('email', '')
                })
            actual = {}
        elif ':' in limpia:
            key, value = limpia.split(':', 1)
            key_upper = key.upper()
            value = crm_vcf_unescape(value)

            if key_upper.startswith('FN') and not actual.get('nombre'):
                actual['nombre'] = value
            elif key_upper.startswith('N') and not actual.get('nombre'):
                partes = [parte for parte in value.split(';') if parte]
                actual['nombre'] = ' '.join(reversed(partes[:2])).strip() or value
            elif key_upper.startswith('TEL') and not actual.get('telefono'):
                actual['telefono'] = value
            elif key_upper.startswith('EMAIL') and not actual.get('email'):
                actual['email'] = value

    return contactos

def crm_extraer_contactos_archivo(archivo):

    nombre_archivo = (archivo.filename or '').lower()
    contenido = archivo.read().decode('utf-8-sig', errors='ignore')

    if nombre_archivo.endswith('.vcf'):
        return crm_extraer_contactos_vcf(contenido)

    return crm_extraer_contactos_csv(contenido)

@app.route('/crm/clientes/importar', methods=['GET', 'POST'])
def crm_clientes_importar():

    preview = []
    resultado = None
    error = ''

    if request.method == 'POST':
        action = request.form.get('action', '').strip()

        if action == 'confirmar':
            preview = session.get('crm_contactos_preview', [])
            clientes = crm_leer_csv(CRM_CLIENTES_CSV, CRM_CLIENTE_HEADERS)
            importados = 0
            duplicados = 0
            errores = 0

            for contacto in preview:
                try:
                    if crm_contacto_duplicado(contacto, clientes):
                        duplicados += 1
                        continue

                    cliente = {
                        'id': crm_generar_id('client'),
                        'nombre': contacto.get('nombre', '').strip(),
                        'telefono': contacto.get('telefono', '').strip(),
                        'email': contacto.get('email', '').strip(),
                        'tipo': 'Comprador',
                        'propiedad': '',
                        'fuente': 'Teléfono',
                        'historial': 'Importado desde contactos del teléfono',
                        'notas': '',
                        'fecha_creacion': crm_now()
                    }

                    clientes.append(cliente)
                    importados += 1
                except Exception:
                    errores += 1

            crm_escribir_csv(CRM_CLIENTES_CSV, CRM_CLIENTE_HEADERS, clientes)
            session.pop('crm_contactos_preview', None)
            resultado = {
                'importados': importados,
                'duplicados': duplicados,
                'errores': errores
            }
            preview = []
        else:
            archivo = request.files.get('archivo')

            if not archivo or not archivo.filename:
                error = 'Seleccioná un archivo CSV o VCF.'
            elif not archivo.filename.lower().endswith(('.csv', '.vcf')):
                error = 'Formato no permitido. Subí un archivo .csv o .vcf.'
            else:
                try:
                    preview = crm_extraer_contactos_archivo(archivo)
                    session['crm_contactos_preview'] = preview
                except Exception as e:
                    error = f'No se pudo leer el archivo: {e}'

    return render_template(
        'crm/importar_contactos.html',
        active='clientes',
        title='Importar contactos',
        preview=preview,
        resultado=resultado,
        error=error
    )

@app.route('/crm/clientes', methods=['GET', 'POST'])
def crm_clientes():

    if request.method == 'POST':
        cliente = {
            'id': crm_generar_id('client'),
            'nombre': request.form.get('nombre', '').strip(),
            'telefono': request.form.get('telefono', '').strip(),
            'email': request.form.get('email', '').strip(),
            'tipo': request.form.get('tipo', '').strip() or 'Comprador',
            'propiedad': request.form.get('propiedad', '').strip(),
            'fuente': request.form.get('fuente', '').strip() or 'Manual',
            'historial': request.form.get('historial', '').strip(),
            'notas': request.form.get('notas', '').strip(),
            'fecha_creacion': crm_now()
        }
        crm_agregar_csv(CRM_CLIENTES_CSV, CRM_CLIENTE_HEADERS, cliente)

    clientes = crm_leer_csv(CRM_CLIENTES_CSV, CRM_CLIENTE_HEADERS)
    busqueda = request.args.get('q', '').strip().lower()
    tipo = request.args.get('tipo', '').strip()

    if tipo:
        clientes = [
            cliente
            for cliente in clientes
            if cliente.get('tipo') == tipo
        ]

    if busqueda:
        clientes = [
            cliente
            for cliente in clientes
            if busqueda in ' '.join([
                cliente.get('nombre', ''),
                cliente.get('telefono', ''),
                cliente.get('email', ''),
                cliente.get('propiedad', ''),
                cliente.get('notas', '')
            ]).lower()
        ]

    compradores = [
        cliente
        for cliente in clientes
        if cliente.get('tipo') == 'Comprador'
    ]
    vendedores = [
        cliente
        for cliente in clientes
        if cliente.get('tipo') == 'Vendedor'
    ]

    return render_template(
        'crm/clientes.html',
        active='clientes',
        title='Clientes',
        clientes=clientes,
        compradores=compradores,
        vendedores=vendedores,
        fuentes=FUENTES_LEAD,
        busqueda=busqueda,
        filtro_tipo=tipo
    )

@app.route('/crm/cliente/<cliente_id>')
def crm_cliente_perfil(cliente_id):

    cliente = crm_cliente_por_id(cliente_id)

    if not cliente:
        return "Cliente no encontrado.", 404

    return render_template(
        'crm/cliente_perfil.html',
        active='clientes',
        title='Perfil de cliente',
        cliente=cliente,
        tareas=crm_tareas_asociadas_cliente(cliente)
    )

@app.route('/crm/agenda', methods=['GET', 'POST'])
def crm_agenda():

    if request.method == 'POST':
        action = request.form.get('action', '').strip() or 'create'
        agenda = crm_leer_csv(CRM_AGENDA_CSV, CRM_AGENDA_HEADERS)
        visita_id = request.form.get('id', '').strip()

        if action in ['realizada', 'cancelar', 'update']:
            for visita in agenda:
                if visita.get('id') == visita_id:
                    if action == 'realizada':
                        visita['estado'] = 'Realizada'
                    elif action == 'cancelar':
                        visita['estado'] = 'Cancelada'
                    else:
                        visita['fecha'] = request.form.get('fecha', '').strip()
                        visita['hora'] = request.form.get('hora', '').strip()
                        visita['cliente'] = request.form.get('cliente', '').strip()
                        visita['telefono'] = request.form.get('telefono', '').strip()
                        visita['propiedad'] = request.form.get('propiedad', '').strip()
                        visita['estado'] = request.form.get('estado', '').strip() or 'Agendada'
                        visita['notas'] = request.form.get('notas', '').strip()
                    break
            crm_escribir_csv(CRM_AGENDA_CSV, CRM_AGENDA_HEADERS, agenda)
            return redirect(url_for('crm_agenda'))

        visita = {
            'id': crm_generar_id('visit'),
            'fecha': request.form.get('fecha', '').strip(),
            'hora': request.form.get('hora', '').strip(),
            'cliente': request.form.get('cliente', '').strip(),
            'telefono': request.form.get('telefono', '').strip(),
            'propiedad': request.form.get('propiedad', '').strip(),
            'estado': request.form.get('estado', '').strip() or 'Agendada',
            'notas': request.form.get('notas', '').strip(),
            'fecha_creacion': crm_now()
        }
        crm_agregar_csv(CRM_AGENDA_CSV, CRM_AGENDA_HEADERS, visita)
        return redirect(url_for('crm_agenda'))

    agenda = crm_leer_csv(CRM_AGENDA_CSV, CRM_AGENDA_HEADERS)
    agenda.sort(
        key=lambda visita: f"{visita.get('fecha', '')} {visita.get('hora', '')}"
    )

    return render_template(
        'crm/agenda.html',
        active='agenda',
        title='Agenda',
        agenda=agenda,
        agenda_por_fecha=crm_agenda_por_fecha(agenda),
        visitas_hoy=crm_visitas_hoy(agenda),
        estados=CRM_ESTADOS_VISITA
    )

@app.route('/crm/notificaciones')
def crm_notificaciones():

    context = crm_dashboard_context()
    return render_template(
        'crm/notificaciones.html',
        active='notificaciones',
        title='Notificaciones',
        notificaciones=crm_notificaciones_context(),
        recordatorios_tareas=context['recordatorios_tareas'],
        tareas_vence_ahora=context['tareas_vence_ahora'],
        tareas_proxima_hora=context['tareas_proxima_hora'],
        tareas_por_prioridad=context['tareas_por_prioridad'],
        tareas_hoy=context['tareas_hoy'],
        seguimientos_atrasados=context['seguimientos_atrasados']
    )

# =========================
# IA WEB
# =========================

@app.route('/chat', methods=['POST'])
def chat_web():

    try:

        data = request.get_json()

        mensaje = data.get("mensaje", "")

        if cliente is None:
            raise RuntimeError("OPENAI_API_KEY no configurada")

        respuesta_ia = cliente.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": """
Sos un asesor inmobiliario profesional de Paraguay.
Respondé natural, corto y amable.
"""
                },
                {
                    "role": "user",
                    "content": mensaje
                }
            ]
        )

        respuesta = respuesta_ia.choices[0].message.content

        return jsonify({
            "reply": respuesta
        })

    except Exception as e:

        print("ERROR IA:", e)

        return jsonify({
            "reply": "Ahora mismo el asistente no está disponible."
        })

# =========================
# WHATSAPP IA
# =========================

@app.route('/mensaje', methods=['POST'])
def whatsapp_bot():

    mensaje = request.form.get('Body')
    numero_cliente = request.form.get('From')

    propiedades = """

    PROPIEDADES DISPONIBLES:

    1) Dúplex en Villa Aurelia
    - Precio: Gs. 1.397.000.000
    - 3 dormitorios
    - 2 baños
    - Patio
    - Quincho
    - Crédito bancario

    2) Edificio Bernardino
    - Precio: Gs. 1.680.000.000
    - 2 dormitorios
    - Piscina
    - Gimnasio
    - Quinchos
    - Playground
    - Portería 24h

    3) Ventura Hassler
    - Desde 137.173 USD

    4) Edificio Ventura Ykua Sati
    - Desde 100.000 USD
    
     DUPLEX SHOPPING DEL SOL

Precio: USD 220.000

175 m² construidos

Ubicado a 5 minutos del Shopping del Sol.

Planta Baja:
- Garage para 2 vehículos
- Sala comedor integrada
- Cocina moderna amoblada
- Baño social
- Área de servicio
- Lavadero
- Tragaluz

Planta Alta:
- Suite con vestidor
- 2 habitaciones
- Baño familiar
- Quincho climatizado con parrilla

Ubicación:
https://maps.app.goo.gl/2rFB5deRSUYXeXov8

    """

    if cliente is None:
        resp = MessagingResponse()
        resp.message("Ahora mismo el asistente no está disponible.")
        return str(resp)

    respuesta_ia = cliente.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {
                "role": "system",
                "content": f"""
                Sos un asesor inmobiliario profesional.

                Catálogo:
                {propiedades}
                """
            },
            {
                "role": "user",
                "content": mensaje
            }
        ]
    )

    respuesta = respuesta_ia.choices[0].message.content

    respuesta = f"""
     Hola 👋

     Soy Silvia Espínola, agente inmobiliaria asociada a Keller Williams Paraguay.

     Gracias por comunicarte 😊
     Te responderé en la brevedad posible.

     Mientras tanto, podés ver algunas propiedades disponibles aquí:

     https://silviaespinolakw.up.railway.app

     --------------------------------

      {respuesta}
      """

    palabras_calientes = [
        "precio",
        "comprar",
        "visita",
        "agendar",
        "usd",
        "casa",
        "departamento"
    ]

    cliente_caliente = any(
        palabra in mensaje.lower()
        for palabra in palabras_calientes
    )

    if cliente_caliente:

        estado = "CALIENTE 🔥"

        guardar_cliente_caliente(
            numero_cliente,
            mensaje,
            respuesta
        )

    else:

        estado = "FRIO"

    guardar_en_excel(
        numero_cliente,
        mensaje,
        respuesta,
        estado
    )

    resp = MessagingResponse()
    resp.message(respuesta)

    return str(resp)

# =========================
# START
# =========================

if __name__ == '__main__':

    app.run(
        host="0.0.0.0",
        port=int(os.getenv("PORT", 5000)),
        debug=True
    )






