from flask import Flask, request, jsonify
from flask_cors import CORS
from openai import OpenAI
from flask import Flask, request, send_from_directory
from twilio.twiml.messaging_response import MessagingResponse
from twilio.rest import Client
import csv
from datetime import datetime
import os

app = Flask(__name__)
CORS(app)

@app.route("/")
def home():
    return send_from_directory(".", "index.html")

# 🔑 API KEY OPENAI
cliente = OpenAI(
    api_key=os.getenv("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")
)
account_sid = os.getenv("AC05e4e8d4cdca2fd34a1d688b6fb09f8f")
auth_token = os.getenv("c9e703d17d240d6f4b80059dddfca9f8")

twilio_client = Client(account_sid, auth_token)
# 📁 GUARDAR CLIENTES
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

# 📁 CLIENTES CALIENTES
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

# 🚀 CHAT IA
@app.route('/mensaje', methods=['POST'])
def chat():

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
    - 7 pisos
    - 56 departamentos
    - 62 cocheras
    - Piscina
    - Gimnasio
    - Financiación

    4) Edificio Ventura Ykua Sati
    - Desde 100.000 USD
    - 2 torres
    - 1, 2 y 3 dormitorios
    - 1028 m² patio interno
    - Financiación propia
    - Entrega 2027 y 2028

    """

    respuesta_ia = cliente.chat.completions.create(
        model="gpt-4o-mini",

        messages=[

            {
                "role": "system",
                "content": f"""
                Sos un asesor inmobiliario profesional de Paraguay.

                Respondé como humano, natural y amable.

                Tu objetivo es:
                - ayudar al cliente
                - detectar interés
                - recomendar propiedades
                - llevar al cliente a una visita

                Siempre preguntá:
                - presupuesto
                - zona
                - tipo de propiedad

                Detectá automáticamente clientes interesados.

                Catálogo actual:
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

    # 📸 MENSAJES AUTOMÁTICOS

    if "villa aurelia" in mensaje.lower():

        twilio_client.messages.create(
            from_='whatsapp:+14155238886',
            to=numero_cliente,
            body='📸 Te puedo mostrar más fotos y detalles del Dúplex en Villa Aurelia. Tiene patio, quincho y acepta crédito bancario 😎'
        )

    if "bernardino" in mensaje.lower():

        twilio_client.messages.create(
            from_='whatsapp:+14155238886',
            to=numero_cliente,
            body='🏢 El Edificio Bernardino tiene piscina, gimnasio, quinchos y portería 24h 🔥'
        )

    if "ventura" in mensaje.lower():

        twilio_client.messages.create(
            from_='whatsapp:+14155238886',
            to=numero_cliente,
            body='✨ Ventura Ykua Sati tiene financiación y departamentos desde 100.000 USD.'
        )

    # 🔥 CLIENTE CALIENTE

    palabras_calientes = [
        "precio",
        "presupuesto",
        "quiero",
        "comprar",
        "visita",
        "agendar",
        "usd",
        "dolares",
        "departamento",
        "casa",
        "zona"
    ]

    cliente_caliente = any(
        palabra in mensaje.lower()
        for palabra in palabras_calientes
    )

    if cliente_caliente:

        estado = "CALIENTE 🔥"

        respuesta += """

🔥 Tengo opciones que te pueden encajar muy bien.

Si querés podemos coordinar una visita o seguir por WhatsApp 👇
"""

        guardar_cliente_caliente(
            numero_cliente,
            mensaje,
            respuesta
        )

        print("🔥 CLIENTE CALIENTE:", mensaje)

    else:

        estado = "FRIO"

    # 💾 GUARDAR

    guardar_en_excel(
        numero_cliente,
        mensaje,
        respuesta,
        estado
    )

    # 📲 RESPUESTA WHATSAPP

    resp = MessagingResponse()
    resp.message(respuesta)

    return str(resp)
if __name__ == '__main__':

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )  