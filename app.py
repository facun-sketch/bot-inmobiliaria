
from flask import Flask, request, jsonify, send_from_directory, send_file
from flask_cors import CORS
from openai import OpenAI
from twilio.twiml.messaging_response import MessagingResponse
from twilio.rest import Client
import csv
from datetime import datetime
import os
import json

try:
    import gspread
    from google.oauth2.service_account import Credentials
except ImportError:
    gspread = None
    Credentials = None

app = Flask(__name__)
CORS(app)

# =========================
# PAGINAS
# =========================

@app.route("/")
@app.route("/index.html")
def home():
    return send_from_directory(".", "index.html")

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

# =========================
# TWILIO
# =========================

account_sid = os.getenv("TWILIO_ACCOUNT_SID")
auth_token = os.getenv("TWILIO_AUTH_TOKEN")

twilio_client = Client(account_sid, auth_token) if account_sid and auth_token else None

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
    'estado'
]

def obtener_google_sheets_client():

    if gspread is None or Credentials is None:
        raise RuntimeError("Faltan las dependencias de Google Sheets")

    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]

    service_account_json = os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON')

    if not service_account_json:
        raise RuntimeError("Falta GOOGLE_SERVICE_ACCOUNT_JSON")

    service_account_info = json.loads(service_account_json)
    credentials = Credentials.from_service_account_info(
        service_account_info,
        scopes=scopes
    )

    return gspread.authorize(credentials)

def guardar_interesado_google_sheets(fecha, nombre, telefono, email, propiedad, mensaje, estado):

    sheet_id = os.getenv('GOOGLE_SHEET_ID')

    if not sheet_id:
        raise RuntimeError("Falta GOOGLE_SHEET_ID")

    client = obtener_google_sheets_client()

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

    if not worksheet.row_values(1):
        worksheet.append_row(INTERESADOS_HEADERS)

    worksheet.append_row([
        fecha,
        nombre,
        telefono,
        email,
        propiedad,
        mensaje,
        estado
    ])

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

    try:
        guardar_interesado_google_sheets(
            fecha,
            nombre,
            telefono,
            email,
            propiedad,
            mensaje,
            estado
        )

        return respuesta_interesado(
            "Gracias.",
            "Recibimos tu consulta y nos pondremos en contacto."
        )

    except Exception as e:
        print("ERROR GOOGLE SHEETS:", e)

        return respuesta_interesado(
            "No pudimos guardar tu consulta.",
            "Por favor escribinos por WhatsApp o intentá nuevamente en unos minutos."
        ), 500

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
