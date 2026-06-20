
from flask import Flask, request, jsonify, send_from_directory, send_file, render_template
from flask_cors import CORS
from openai import OpenAI
from twilio.twiml.messaging_response import MessagingResponse
from twilio.rest import Client
import csv
from datetime import datetime
import os

print("APP IMPORTADA")

app = Flask(__name__)
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

    contenido = f'{sender} {subject} {body}'
    return 'infocasas' in contenido

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
    subject = correo.get('subject') or ''
    remitente = correo.get('from') or {}
    email_info = remitente.get('emailAddress') or {}

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

    propiedad = extraer_linea(texto, [
        'propiedad',
        'inmueble',
        'publicación',
        'publicacion',
        'aviso'
    ])

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
        nombre = email_info.get('name', '').strip()

    if not propiedad:
        propiedad = subject.strip()

    if not mensaje:
        mensaje = texto[:1200]

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
