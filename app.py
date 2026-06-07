"""
Flask-Web-App für die Druckerei-Agentic-AI.
Startet mit: python app.py
"""

import json
import os
from flask import Flask, jsonify, request, send_from_directory
from agent import PrintShopAgent, FIELD_LABELS, REQUIRED_FIELDS, _load_orders

app = Flask(__name__, static_folder="static", template_folder="templates")

# Aktive Sitzungen im Speicher (für Produktion: Redis o.Ä. verwenden)
sessions: dict[str, PrintShopAgent] = {}


@app.route("/")
def index():
    return send_from_directory("templates", "index.html")


@app.route("/api/session", methods=["POST"])
def create_session():
    """Erstellt eine neue Bestellsitzung."""
    agent = PrintShopAgent()
    sessions[agent.session_id] = agent

    greeting = (
        "Willkommen bei Ihrer Druckerei! Ich bin Ihr persönlicher Auftragsassistent. "
        "Bitte schildern Sie mir Ihren Druckbedarf — was möchten Sie drucken, "
        "in welcher Menge und bis wann? Je mehr Details Sie angeben, desto schneller "
        "kann ich Ihren Auftrag aufnehmen."
    )

    return jsonify({
        "session_id": agent.session_id,
        "message": greeting,
    })


@app.route("/api/chat", methods=["POST"])
def chat():
    """Verarbeitet eine Nachricht in einer Sitzung."""
    data = request.get_json()
    session_id = data.get("session_id")
    message = data.get("message", "").strip()

    if not session_id or session_id not in sessions:
        return jsonify({"error": "Ungültige oder abgelaufene Sitzung."}), 400

    if not message:
        return jsonify({"error": "Nachricht darf nicht leer sein."}), 400

    agent = sessions[session_id]

    if agent.order_complete:
        return jsonify({
            "message": "Ihr Auftrag ist bereits vollständig aufgenommen. "
                        "Starten Sie eine neue Sitzung für einen weiteren Auftrag.",
            "order_complete": True,
            "status": agent.get_order_status(),
        })

    reply = agent.chat(message)
    status = agent.get_order_status()

    return jsonify({
        "message": reply,
        "order_complete": agent.order_complete,
        "status": status,
    })


@app.route("/api/status/<session_id>", methods=["GET"])
def session_status(session_id: str):
    """Gibt den Bestellstatus einer Sitzung zurück."""
    if session_id not in sessions:
        return jsonify({"error": "Sitzung nicht gefunden."}), 404

    agent = sessions[session_id]
    status = agent.get_order_status()

    # Feldbezeichnungen hinzufügen
    status["field_labels"] = {
        f: FIELD_LABELS[f] for f in REQUIRED_FIELDS
    }
    return jsonify(status)


@app.route("/api/orders", methods=["GET"])
def list_orders():
    """Listet alle abgeschlossenen Aufträge auf."""
    orders = _load_orders()
    return jsonify({"orders": orders, "total": len(orders)})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    print(f"Druckerei-Agent läuft auf http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=debug)
