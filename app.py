
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
    'fecha_seguimiento'
]

ESTADOS_INTERESADO = [
    'Nuevo',
    'Contactado',
    'Visita agendada',
    'Cerrado'
]

def guardar_interesado_backup(row):

    archivo = 'interesados_backup.csv'
    crear_encabezado = not os.path.exists(archivo) or os.path.getsize(archivo) == 0

    with open(archivo, mode='a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)

        if crear_encabezado:
            writer.writerow(INTERESADOS_HEADERS)

        writer.writerow(row)

def guardar_interesado_google_sheets(row):

    print("Intentando guardar interesado en Google Sheets")

    worksheet, gspread = obtener_interesados_worksheet()

    worksheet.append_row(row)

    print("Interesado guardado en Google Sheets")

def obtener_interesados_worksheet():

    import json
    import gspread
    from google.oauth2.service_account import Credentials

    service_account_json = os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON')
    sheet_id = os.getenv('GOOGLE_SHEET_ID')

    if not service_account_json:
        raise RuntimeError("Falta GOOGLE_SERVICE_ACCOUNT_JSON")

    if not sheet_id:
        raise RuntimeError("Falta GOOGLE_SHEET_ID")

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

    headers_actuales = worksheet.row_values(1)

    if not headers_actuales:
        worksheet.append_row(INTERESADOS_HEADERS)
    elif headers_actuales != INTERESADOS_HEADERS:
        headers_actualizados = headers_actuales[:]

        for header in INTERESADOS_HEADERS:
            if header not in headers_actualizados:
                headers_actualizados.append(header)

        worksheet.update('1:1', [headers_actualizados])

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

    row = [
        fecha,
        nombre,
        telefono,
        email,
        propiedad,
        mensaje,
        estado,
        fecha_seguimiento
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

        worksheet, gspread = obtener_interesados_worksheet()
        values = worksheet.get_all_values()
        headers = values[0] if values else INTERESADOS_HEADERS
        rows = values[1:] if len(values) > 1 else []

        table_rows = ""

        for index, row in enumerate(rows, start=2):
            interesado = {
                header: row[pos] if pos < len(row) else ''
                for pos, header in enumerate(headers)
            }

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

        worksheet, gspread = obtener_interesados_worksheet()
        headers = worksheet.row_values(1)

        estado_col = headers.index('estado') + 1
        seguimiento_col = headers.index('fecha_seguimiento') + 1
        fecha_seguimiento = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        worksheet.update_cell(row_number, estado_col, nuevo_estado)
        worksheet.update_cell(row_number, seguimiento_col, fecha_seguimiento)

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
