import json
import os
import random
import time

import paho.mqtt.client as mqtt
from dotenv import load_dotenv

load_dotenv()

TOPIC_ESTADO = "proyectoiot/jutiapa/estado"
TOPIC_COMANDO = "proyectoiot/jutiapa/comando"
LIMITE = 2.5

estado = "NORMAL"


def al_conectar(client, userdata, flags, reason_code, properties):
    print("Simulador conectado a", os.environ["MQTT_HOST"], "->", reason_code)
    client.subscribe(TOPIC_COMANDO)


def al_recibir(client, userdata, msg):
    global estado
    try:
        comando = json.loads(msg.payload.decode())
    except json.JSONDecodeError:
        return
    if isinstance(comando, dict) and comando.get("accion") == "REARMAR" and estado == "CORTE":
        estado = "NORMAL"
        print(">>> Relé rearmado desde la app")


cliente = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2)
cliente.username_pw_set(os.environ["MQTT_USER"], os.environ["MQTT_PASS"])
cliente.tls_set()
cliente.on_connect = al_conectar
cliente.on_message = al_recibir
cliente.connect(os.environ["MQTT_HOST"], 8883)
cliente.loop_start()

while True:
    if estado == "NORMAL":
        # Casi siempre consumo normal; de vez en cuando, un pico
        if random.random() < 0.08:
            corriente = round(random.uniform(2.6, 3.0), 2)
        else:
            corriente = round(random.uniform(0.4, 2.2), 2)
        if corriente > LIMITE:
            estado = "CORTE"
            print(">>> Sobrecorriente: relé abierto")
    else:
        corriente = 0.0   # con el relé abierto no pasa corriente

    dato = {
        "dispositivo": "esp32-jutiapa-01",
        "corriente": corriente,
        "limite": LIMITE,
        "estado": estado,
    }
    cliente.publish(TOPIC_ESTADO, json.dumps(dato))
    print("Enviado:", dato)
    time.sleep(2)