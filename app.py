import hmac
import json
import os
import sqlite3
from collections import deque
from datetime import date, datetime, timedelta
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

# Lo último que mandó el dispositivo, se haya guardado o no.
# De aquí sale la vista en vivo del dashboard.
recientes = deque(maxlen=20)


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
    con.execute("CREATE INDEX IF NOT EXISTS idx_lecturas_fecha ON lecturas (fecha)")
    con.commit()
    con.close()


def guardar(lectura):
    con = sqlite3.connect(DB)
    con.execute(
        "INSERT INTO lecturas (fecha, dispositivo, corriente, limite, estado) "
        "VALUES (?, ?, ?, ?, ?)",
        (
            lectura["fecha"],
            lectura["dispositivo"],
            lectura["corriente"],
            lectura["limite"],
            lectura["estado"],
        ),
    )
    con.commit()
    con.close()


def es_repetida(lectura, anterior):
    """Una lectura en 0 A es repetida si la anterior también estaba en 0 A
    y en el mismo estado. Esas no aportan nada al historial."""
    return (
        anterior is not None
        and lectura["corriente"] == 0
        and anterior["corriente"] == 0
        and lectura["estado"] == anterior["estado"]
    )


def consultar(sql, params=()):
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    filas = [dict(f) for f in con.execute(sql, params).fetchall()]
    con.close()
    return filas


def ultima_lectura():
    """Lo más reciente que se sabe del dispositivo: primero lo que llegó
    en vivo; si la app acaba de arrancar, lo último guardado."""
    if recientes:
        return recientes[-1]
    guardada = consultar("SELECT * FROM lecturas ORDER BY id DESC LIMIT 1")
    return guardada[0] if guardada else None


def estado_actual():
    """'NORMAL', 'CORTE' o None si todavía no hay datos."""
    ultima = ultima_lectura()
    return ultima["estado"] if ultima else None


# ---------- MQTT ----------

def al_conectar(client, userdata, flags, reason_code, properties):
    print("Conectado al broker", MQTT_HOST, "->", reason_code)
    client.subscribe(TOPIC_ESTADO)


def al_recibir(client, userdata, msg):
    try:
        dato = json.loads(msg.payload.decode())
        lectura = {
            "fecha": datetime.now().isoformat(timespec="seconds"),
            "dispositivo": str(dato["dispositivo"]),
            "corriente": float(dato["corriente"]),
            "limite": float(dato["limite"]),
            "estado": str(dato["estado"]),
        }
        anterior = ultima_lectura()
        recientes.append(lectura)          # la vista en vivo recibe todo
        if es_repetida(lectura, anterior):
            print("En vivo (sin guardar):", dato)
        else:
            guardar(lectura)
            print("Guardado:", dato)
    except (json.JSONDecodeError, KeyError, TypeError, ValueError) as error:
        print("Mensaje ignorado:", error)


def crear_cliente_mqtt():
    cliente = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2)
    cliente.username_pw_set(MQTT_USER, MQTT_PASS)
    cliente.tls_set()
    cliente.on_connect = al_conectar
    cliente.on_message = al_recibir
    cliente.connect(MQTT_HOST, MQTT_PORT)
    return cliente


def enviar_comando(accion):
    """Publica una orden para el dispositivo. El dispositivo decide si la ejecuta."""
    cliente_mqtt.publish(TOPIC_COMANDO, json.dumps({"accion": accion}), qos=1)
    print(">>> Orden", accion, "enviada por", session["usuario"])


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


# ---------- Páginas ----------

@app.route("/")
@requiere_login
def inicio():
    return render_template("index.html", usuario=session["usuario"], pagina="dashboard")


@app.route("/reportes")
@requiere_login
def reportes():
    return render_template("reportes.html", usuario=session["usuario"], pagina="reportes")


# ---------- API del dashboard ----------

@app.route("/api/lecturas")
@requiere_login
def api_lecturas():
    # La más nueva primero. Si la app acaba de arrancar, se muestra lo guardado.
    if recientes:
        return jsonify(list(recientes)[::-1])
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
    return jsonify({
        "ultima": ultima_lectura(),
        "maximo": metricas["maximo"],
        "promedio": metricas["promedio"],
        "total": metricas["total"],
        "cortes": cortes["cortes"],
    })


@app.route("/api/rearmar", methods=["POST"])
@requiere_login
def api_rearmar():
    if estado_actual() != "CORTE":
        return jsonify({"ok": False, "mensaje": "El relé no está en corte."}), 409
    enviar_comando("REARMAR")
    return jsonify({"ok": True, "mensaje": "Orden de reactivación enviada."})


@app.route("/api/cortar", methods=["POST"])
@requiere_login
def api_cortar():
    if estado_actual() != "NORMAL":
        return jsonify({"ok": False, "mensaje": "El relé ya está abierto o no hay datos."}), 409
    enviar_comando("CORTAR")
    return jsonify({"ok": True, "mensaje": "Orden de corte enviada."})


# ---------- API de reportes ----------

PERIODOS = ("dia", "semana", "mes", "anio")


def rango_periodo(periodo, dia):
    """Devuelve (inicio, fin) del período que contiene a 'dia'. El fin no se incluye."""
    if periodo == "semana":
        inicio = dia - timedelta(days=dia.weekday())          # lunes
        fin = inicio + timedelta(days=7)
    elif periodo == "mes":
        inicio = dia.replace(day=1)
        fin = (inicio + timedelta(days=32)).replace(day=1)    # día 1 del mes siguiente
    elif periodo == "anio":
        inicio = dia.replace(month=1, day=1)
        fin = inicio.replace(year=inicio.year + 1)
    else:
        inicio = dia
        fin = dia + timedelta(days=1)
    return inicio, fin


@app.route("/api/reporte")
@requiere_login
def api_reporte():
    periodo = request.args.get("periodo", "dia")
    if periodo not in PERIODOS:
        periodo = "dia"
    try:
        dia = date.fromisoformat(request.args.get("fecha", ""))
    except ValueError:
        dia = date.today()

    inicio, fin = rango_periodo(periodo, dia)
    rango = (inicio.isoformat(), fin.isoformat())

    resumen = consultar(
        "SELECT COUNT(*) AS total, MAX(corriente) AS maximo, "
        "AVG(corriente) AS promedio, MAX(limite) AS limite "
        "FROM lecturas WHERE fecha >= ? AND fecha < ?",
        rango,
    )[0]

    # Día -> por hora; semana y mes -> por día; año -> por mes.
    # 'grupo' sale de valores fijos, nunca de lo que manda el navegador.
    if periodo == "dia":
        grupo = "substr(fecha, 12, 2)"
    elif periodo == "anio":
        grupo = "substr(fecha, 1, 7)"
    else:
        grupo = "substr(fecha, 1, 10)"

    serie = consultar(
        f"SELECT {grupo} AS etiqueta, COUNT(*) AS lecturas, "
        "AVG(corriente) AS promedio, MAX(corriente) AS maximo "
        "FROM lecturas WHERE fecha >= ? AND fecha < ? "
        "GROUP BY etiqueta ORDER BY etiqueta",
        rango,
    )

    cortes = consultar(
        "SELECT fecha, corriente FROM ("
        "  SELECT fecha, corriente, estado, LAG(estado) OVER (ORDER BY id) AS anterior"
        "  FROM lecturas WHERE fecha >= ? AND fecha < ?"
        ") WHERE estado = 'CORTE' AND (anterior IS NULL OR anterior <> 'CORTE') "
        "ORDER BY fecha DESC",
        rango,
    )

    resumen["cortes"] = len(cortes)
    return jsonify({
        "periodo": periodo,
        "desde": inicio.isoformat(),
        "hasta": (fin - timedelta(days=1)).isoformat(),
        "resumen": resumen,
        "serie": serie,
        "cortes": cortes,
    })


# ---------- Arranque ----------

if __name__ == "__main__":
    crear_tabla()
    cliente_mqtt = crear_cliente_mqtt()
    cliente_mqtt.loop_start()
    app.run(port=5000)