
from flask import Flask, request, jsonify, send_from_directory, send_file
from flask_cors import CORS
from openai import OpenAI
from twilio.twiml.messaging_response import MessagingResponse
from twilio.rest import Client
import csv
from datetime import datetime
import os

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

cliente = OpenAI(
    api_key=os.getenv("OPENAI_API_KEY")
)

# =========================
# TWILIO
# =========================

account_sid = os.getenv("TWILIO_ACCOUNT_SID")
auth_token = os.getenv("TWILIO_AUTH_TOKEN")

twilio_client = Client(account_sid, auth_token)

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
# IA WEB
# =========================

@app.route('/chat', methods=['POST'])
def chat_web():

    try:

        data = request.get_json()

        mensaje = data.get("mensaje", "")

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

    """

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
        port=5000,
        debug=True
    )
