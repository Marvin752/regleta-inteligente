import paho.mqtt.client as mqtt

BROKER = "broker.hivemq.com"
TOPIC = "proyectoiot/jutiapa/#"   # el # significa "todo lo que cuelgue de aquí"

def al_conectar(client, userdata, flags, reason_code, properties):
    print("Conectado al broker:", reason_code)
    client.subscribe(TOPIC)

def al_recibir(client, userdata, msg):
    print(msg.topic, "->", msg.payload.decode())

cliente = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2)
cliente.on_connect = al_conectar
cliente.on_message = al_recibir
cliente.connect(BROKER, 1883)
cliente.loop_forever()