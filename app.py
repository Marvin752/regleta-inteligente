import hmac
import json
import os
import sqlite3
from datetime import datetime
from functools import wraps

import paho.mqtt.client as mqtt
from dotenv import load_dotenv
from flask import (Flask, jsonify, redirect, render_template, request,
                   session, url_for)

load_dotenv()

MQTT_HOST = os.environ["MQTT_HOST"]
MQTT_USER = os.environ["MQTT_USER"]
MQTT_PASS = os.environ["MQTT_PASS"]
MQTT_PORT = 8883
TOPIC_ESTADO = "proyectoiot/jutiapa/estado"
TOPIC_COMANDO = "proyectoiot/jutiapa/comando"
DB = "lecturas.db"

APP_USER = os.environ["APP_USER"]
APP_PASS = os.environ["APP_PASS"]

app = Flask(__name__)
app.secret_key = os.environ["SECRET_KEY"]
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
cliente_mqtt = None


# ---------- Base de datos ----------

def crear_tabla():
    con = sqlite3.connect(DB)
    con.execute("""
        CREATE TABLE IF NOT EXISTS lecturas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT NOT NULL,
            dispositivo TEXT NOT NULL,
            corriente REAL NOT NULL,
            limite REAL NOT NULL,
            estado TEXT NOT NULL
        )
    """)
    con.commit()
    con.close()


def guardar(dato):
    con = sqlite3.connect(DB)
    con.execute(
        "INSERT INTO lecturas (fecha, dispositivo, corriente, limite, estado) "
        "VALUES (?, ?, ?, ?, ?)",
        (
            datetime.now().isoformat(timespec="seconds"),
            dato["dispositivo"],
            dato["corriente"],
            dato["limite"],
            dato["estado"],
        ),
    )
    con.commit()
    con.close()


def consultar(sql, params=()):
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    filas = [dict(f) for f in con.execute(sql, params).fetchall()]
    con.close()
    return filas


# ---------- MQTT ----------

def al_conectar(client, userdata, flags, reason_code, properties):
    print("Conectado al broker", MQTT_HOST, "->", reason_code)
    client.subscribe(TOPIC_ESTADO)


def al_recibir(client, userdata, msg):
    try:
        dato = json.loads(msg.payload.decode())
        guardar(dato)
        print("Guardado:", dato)
    except (json.JSONDecodeError, KeyError, TypeError) as error:
        print("Mensaje ignorado:", error)


def crear_cliente_mqtt():
    cliente = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2)
    cliente.username_pw_set(MQTT_USER, MQTT_PASS)
    cliente.tls_set()
    cliente.on_connect = al_conectar
    cliente.on_message = al_recibir
    cliente.connect(MQTT_HOST, MQTT_PORT)
    return cliente


# ---------- Sesión ----------

def requiere_login(vista):
    @wraps(vista)
    def envoltura(*args, **kwargs):
        if not session.get("usuario"):
            if request.path.startswith("/api/"):
                return jsonify({"ok": False, "mensaje": "Sesión no iniciada."}), 401
            return redirect(url_for("login"))
        return vista(*args, **kwargs)
    return envoltura


def credenciales_validas(usuario, clave):
    usuario_ok = hmac.compare_digest(usuario.encode(), APP_USER.encode())
    clave_ok = hmac.compare_digest(clave.encode(), APP_PASS.encode())
    return usuario_ok and clave_ok


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        usuario = request.form.get("usuario", "")
        clave = request.form.get("clave", "")
        if credenciales_validas(usuario, clave):
            session["usuario"] = usuario
            return redirect(url_for("inicio"))
        error = "Usuario o contraseña incorrectos."
    return render_template("login.html", error=error)


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return redirect(url_for("login"))


# ---------- Rutas web ----------

@app.route("/")
@requiere_login
def inicio():
    return render_template("index.html", usuario=session["usuario"])


@app.route("/api/lecturas")
@requiere_login
def api_lecturas():
    return jsonify(consultar("SELECT * FROM lecturas ORDER BY id DESC LIMIT 20"))


@app.route("/api/resumen")
@requiere_login
def api_resumen():
    hoy = datetime.now().date().isoformat()
    metricas = consultar(
        "SELECT COUNT(*) AS total, MAX(corriente) AS maximo, AVG(corriente) AS promedio "
        "FROM lecturas WHERE fecha >= ?",
        (hoy,),
    )[0]
    cortes = consultar(
        "SELECT COUNT(*) AS cortes FROM ("
        "  SELECT estado, LAG(estado) OVER (ORDER BY id) AS anterior"
        "  FROM lecturas WHERE fecha >= ?"
        ") WHERE estado = 'CORTE' AND (anterior IS NULL OR anterior <> 'CORTE')",
        (hoy,),
    )[0]
    ultima = consultar("SELECT * FROM lecturas ORDER BY id DESC LIMIT 1")
    return jsonify({
        "ultima": ultima[0] if ultima else None,
        "maximo": metricas["maximo"],
        "promedio": metricas["promedio"],
        "total": metricas["total"],
        "cortes": cortes["cortes"],
    })


@app.route("/api/rearmar", methods=["POST"])
@requiere_login
def api_rearmar():
    ultima = consultar("SELECT estado FROM lecturas ORDER BY id DESC LIMIT 1")
    if not ultima or ultima[0]["estado"] != "CORTE":
        return jsonify({"ok": False, "mensaje": "El relé no está en corte."}), 409
    cliente_mqtt.publish(TOPIC_COMANDO, json.dumps({"accion": "REARMAR"}), qos=1)
    print(">>> Orden de rearme enviada por", session["usuario"])
    return jsonify({"ok": True, "mensaje": "Orden enviada al dispositivo."})


# ---------- Arranque ----------

if __name__ == "__main__":
    crear_tabla()
    cliente_mqtt = crear_cliente_mqtt()
    cliente_mqtt.loop_start()
    app.run(port=5000)