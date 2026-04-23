"""
Lokaler KI-Proxy für Verwaltungsdokument-Automatisierer.
Flask-Server auf localhost:5000 als Proxy zu Ollama.
Kein Internetzugang erforderlich nach Einrichtung.
"""

import json
import logging
import os
from datetime import datetime
from flask import Flask, request, jsonify

# Konfiguration
OLLAMA_BASE_URL = os.environ.get("OLLAMA_BASE_URL", "http://localhost:11434")
OLLAMA_MODEL = os.environ.get("OLLAMA_MODEL", "llama3")
LOG_DATEI = os.environ.get("LOG_DATEI", "proxy_log.csv")
PROMPTS_DIR = os.path.join(os.path.dirname(__file__), "..", "prompts")

# Logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

app = Flask(__name__)


def load_prompt(prompt_type):
    """Lädt den Prompt aus der Datei."""
    prompt_file = os.path.join(PROMPTS_DIR, f"{prompt_type}.txt")
    try:
        with open(prompt_file, "r", encoding="utf-8") as f:
            return f.read()
    except FileNotFoundError:
        logger.warning(f"Prompt-Datei nicht gefunden: {prompt_file}")
        return "Verbessere folgenden Text im formellen Behördenstil:"


def call_ollama(prompt, text):
    """Ruft Ollama lokal auf."""
    import urllib.request
    import urllib.error

    url = f"{OLLAMA_BASE_URL}/api/generate"

    payload = json.dumps({
        "model": OLLAMA_MODEL,
        "prompt": f"{prompt}\n\n=== TEXT ===\n{text}\n\n=== ANTWORT ===",
        "stream": False,
        "options": {
            "temperature": 0.3,
            "top_p": 0.9
        }
    }).encode("utf-8")

    req = urllib.request.Request(
        url,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST"
    )

    try:
        with urllib.request.urlopen(req, timeout=60) as response:
            result = json.loads(response.read().decode("utf-8"))
            return result.get("response", "")
    except urllib.error.URLError as e:
        logger.error(f"Ollama nicht erreichbar: {e}")
        raise ConnectionError(f"Ollama nicht erreichbar unter {OLLAMA_BASE_URL}")
    except Exception as e:
        logger.error(f"Ollama-Fehler: {e}")
        raise RuntimeError(f"KI-Anfrage fehlgeschlagen: {e}")


def log_request(prompt_type, text_length, success, error=""):
    """Protokolliert Anfragen."""
    try:
        with open(LOG_DATEI, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().isoformat()};{prompt_type};{text_length};{success};{error}\n")
    except Exception as e:
        logger.error(f"Logging fehlgeschlagen: {e}")


@app.route("/health", methods=["GET"])
def health():
    """Health-Check-Endpunkt."""
    return jsonify({"status": "ok", "timestamp": datetime.now().isoformat()})


@app.route("/api/optimize", methods=["POST"])
def optimize():
    """Optimiert einen Text mit KI."""
    try:
        data = request.get_json()

        if not data or "text" not in data:
            return jsonify({"error": "Fehlender Parameter 'text'"}), 400

        text = data.get("text", "")
        prompt_type = data.get("prompt_type", "bescheid_optimierung")

        if not text.strip():
            return jsonify({"error": "Leerer Text"}), 400

        # Prompt laden
        prompt = load_prompt(prompt_type)

        # KI aufrufen
        result = call_ollama(prompt, text)

        # Loggen
        log_request(prompt_type, len(text), True)

        return jsonify({
            "result": result,
            "model": OLLAMA_MODEL,
            "timestamp": datetime.now().isoformat()
        })

    except ConnectionError as e:
        log_request("unknown", 0, False, str(e))
        return jsonify({"error": str(e)}), 503
    except Exception as e:
        log_request("unknown", 0, False, str(e))
        logger.error(f"Fehler bei /api/optimize: {e}")
        return jsonify({"error": str(e)}), 500


@app.route("/api/status", methods=["GET"])
def status():
    """Status-Informationen."""
    ollama_reachable = False
    try:
        import urllib.request
        req = urllib.request.Request(f"{OLLAMA_BASE_URL}/api/tags", method="GET")
        with urllib.request.urlopen(req, timeout=3) as resp:
            ollama_reachable = resp.status == 200
    except Exception:
        pass

    return jsonify({
        "status": "ok",
        "ollama_url": OLLAMA_BASE_URL,
        "ollama_model": OLLAMA_MODEL,
        "ollama_reachable": ollama_reachable,
        "timestamp": datetime.now().isoformat()
    })


if __name__ == "__main__":
    logger.info("KI-Proxy wird gestartet...")
    logger.info(f"Ollama URL: {OLLAMA_BASE_URL}")
    logger.info(f"Ollama Modell: {OLLAMA_MODEL}")
    logger.info("Port: 5000")

    app.run(host="127.0.0.1", port=5000, debug=False)
