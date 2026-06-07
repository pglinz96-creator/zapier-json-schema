"""
Druckerei-Agentic-AI — Kernlogik
Nimmt Aufträge entgegen, extrahiert strukturierte Daten und stellt gezielte
Rückfragen für fehlende Pflichtangaben.
"""

import json
import os
import uuid
from datetime import datetime
from pathlib import Path

import anthropic
from dotenv import load_dotenv

load_dotenv()

_api_key = os.environ.get("ANTHROPIC_API_KEY")
client = anthropic.Anthropic(api_key=_api_key) if _api_key else None
MODEL = "claude-sonnet-4-6"

ORDERS_FILE = Path("auftraege.json")

FIELD_LABELS = {
    "produkt": "Produkt",
    "auflage": "Auflage (Stückzahl)",
    "format": "Format",
    "farbigkeit": "Farbigkeit",
    "material": "Material / Papier",
    "veredelung": "Veredelung",
    "liefertermin": "Liefertermin",
    "besondere_anforderungen": "Besondere Anforderungen",
}

FIELD_EXAMPLES = {
    "produkt": "z.B. Flyer, Visitenkarten, Plakat, Broschüre, Banner",
    "auflage": "z.B. 500, 1000, 5000 Stück",
    "format": "z.B. A4, A5, DIN lang (99×210 mm), 90×50 mm",
    "farbigkeit": "z.B. 4/4 (beidseitig farbig), 4/0 (einseitig farbig), 1/0 (schwarz)",
    "material": "z.B. 135g/m² Bilderdruck, 300g/m² Karton, 80g/m² Offsetpapier",
    "veredelung": "z.B. Keine, Glanzlaminat, Mattlaminat, UV-Lack, Prägung",
    "liefertermin": "z.B. 15.06.2026 oder 'so schnell wie möglich'",
    "besondere_anforderungen": "z.B. Pantone-Farben, Schneidmarken, Nummerierung",
}

REQUIRED_FIELDS = [
    "produkt", "auflage", "format", "farbigkeit",
    "material", "veredelung", "liefertermin",
]

EXTRACT_TOOL = {
    "name": "auftrags_extraktion",
    "description": (
        "Extrahiert strukturierte Druckauftragsdaten aus dem Kundentext. "
        "Felder, die im Text nicht erwähnt werden, bleiben leer (leerer String). "
        "Liefert außerdem eine Liste der fehlenden Pflichtfelder."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "produkt": {"type": "string", "description": "Art des Druckprodukts"},
            "auflage": {"type": "string", "description": "Stückzahl / Auflage"},
            "format": {"type": "string", "description": "Druckformat / Größe"},
            "farbigkeit": {"type": "string", "description": "Farbigkeit (z.B. 4/4, 1/0)"},
            "material": {"type": "string", "description": "Papier oder Material"},
            "veredelung": {"type": "string", "description": "Veredelung (Laminat, UV usw.)"},
            "liefertermin": {"type": "string", "description": "Gewünschter Liefertermin"},
            "besondere_anforderungen": {
                "type": "string",
                "description": "Sonstige besondere Anforderungen",
            },
            "fehlende_angaben": {
                "type": "array",
                "items": {"type": "string"},
                "description": "Liste der Felder, die im Text fehlen und nachgefragt werden müssen",
            },
        },
        "required": [
            "produkt", "auflage", "format", "farbigkeit",
            "material", "veredelung", "liefertermin",
            "besondere_anforderungen", "fehlende_angaben",
        ],
    },
}

SYSTEM_PROMPT = """Du bist ein erfahrener Kundenberater einer professionellen Druckerei.
Du hilfst Kunden dabei, Druckaufträge präzise aufzunehmen.

Aufgaben:
1. Extrahiere alle Auftragsinformationen aus dem Kundenttext.
2. Erkenne, welche Pflichtangaben fehlen.
3. Formuliere freundliche, konkrete Rückfragen für fehlende Informationen.
4. Bestätige vollständige Aufträge professionell.

Sprache: Deutsch. Ton: freundlich, professionell, präzise.
Wenn Informationen unvollständig sind, frage gezielt nach — niemals alle
fehlenden Felder auf einmal, sondern priorisiert (max. 2-3 Fragen pro Runde).
"""


def _load_orders() -> list:
    if ORDERS_FILE.exists():
        return json.loads(ORDERS_FILE.read_text(encoding="utf-8"))
    return []


def _save_order(order: dict) -> None:
    orders = _load_orders()
    orders.append(order)
    ORDERS_FILE.write_text(
        json.dumps(orders, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _get_client() -> anthropic.Anthropic:
    if client is None:
        raise RuntimeError(
            "ANTHROPIC_API_KEY nicht gesetzt. "
            "Bitte .env Datei anlegen oder Umgebungsvariable setzen."
        )
    return client


def extract_order_data(conversation_history: list) -> dict:
    """Ruft Claude mit Tool Use auf, um Auftragsdaten zu extrahieren."""
    response = _get_client().messages.create(
        model=MODEL,
        max_tokens=1024,
        system=SYSTEM_PROMPT,
        tools=[EXTRACT_TOOL],
        tool_choice={"type": "auto"},
        messages=conversation_history,
    )

    for block in response.content:
        if block.type == "tool_use" and block.name == "auftrags_extraktion":
            return block.input

    return {}


def generate_followup(conversation_history: list, missing_fields: list) -> str:
    """Generiert eine natürliche Rückfrage für die fehlenden Felder."""
    fields_info = "\n".join(
        f"- {FIELD_LABELS[f]} ({FIELD_EXAMPLES[f]})"
        for f in missing_fields
        if f in FIELD_LABELS
    )

    prompt_addition = (
        f"\n\nFolgende Pflichtangaben fehlen noch:\n{fields_info}\n\n"
        "Formuliere eine freundliche Rückfrage. Priorisiere die wichtigsten "
        "fehlenden Angaben (max. 2-3 auf einmal). Erkläre kurz, warum diese "
        "Informationen für die Kalkulation wichtig sind."
    )

    messages = conversation_history + [
        {
            "role": "user",
            "content": f"[System: Generiere jetzt die Rückfrage.{prompt_addition}]",
        }
    ]

    response = _get_client().messages.create(
        model=MODEL,
        max_tokens=512,
        system=SYSTEM_PROMPT,
        messages=messages,
    )

    return response.content[0].text


def generate_confirmation(order_data: dict) -> str:
    """Generiert eine professionelle Auftragsbestätigung."""
    order_summary = "\n".join(
        f"- {FIELD_LABELS.get(k, k)}: {v}"
        for k, v in order_data.items()
        if k not in ("fehlende_angaben", "rueckfrage", "id", "erstellt_am")
        and v
    )

    response = _get_client().messages.create(
        model=MODEL,
        max_tokens=512,
        system=SYSTEM_PROMPT,
        messages=[
            {
                "role": "user",
                "content": (
                    f"Bestätige diesen vollständigen Druckauftrag freundlich und professionell:\n\n"
                    f"{order_summary}\n\n"
                    "Erstelle eine strukturierte Auftragsbestätigung mit allen Details "
                    "und gib dem Kunden eine Auftragsnummer."
                ),
            }
        ],
    )

    return response.content[0].text


class PrintShopAgent:
    """Zustandsbasierter Agent für eine Druckerei-Bestellsitzung."""

    def __init__(self):
        self.session_id = str(uuid.uuid4())[:8]
        self.conversation: list = []
        self.current_order: dict = {}
        self.order_complete = False

    def chat(self, user_message: str) -> str:
        """Verarbeitet eine Nutzernachricht und gibt die Agentenantwort zurück."""
        self.conversation.append({"role": "user", "content": user_message})

        # Extrahiere Auftragsdaten aus der gesamten Konversation
        extracted = extract_order_data(self.conversation)

        # Merge mit bestehendem Auftrag (neue Daten überschreiben leere Felder)
        for key, value in extracted.items():
            if key == "fehlende_angaben":
                continue
            if value and (key not in self.current_order or not self.current_order[key]):
                self.current_order[key] = value

        # Prüfe fehlende Pflichtfelder
        missing = [
            f for f in REQUIRED_FIELDS
            if not self.current_order.get(f)
        ]

        if missing:
            reply = generate_followup(self.conversation, missing)
            self.current_order["fehlende_angaben"] = missing
        else:
            # Alle Pflichtfelder vorhanden — Auftrag abschließen
            self.current_order["id"] = f"DA-{self.session_id.upper()}"
            self.current_order["erstellt_am"] = datetime.now().isoformat()
            self.current_order["fehlende_angaben"] = []
            _save_order(self.current_order)
            self.order_complete = True
            reply = generate_confirmation(self.current_order)

        self.conversation.append({"role": "assistant", "content": reply})
        return reply

    def get_order_status(self) -> dict:
        """Gibt den aktuellen Bestellstatus zurück."""
        filled = {
            k: v for k, v in self.current_order.items()
            if k in REQUIRED_FIELDS and v
        }
        missing = [f for f in REQUIRED_FIELDS if not self.current_order.get(f)]
        return {
            "session_id": self.session_id,
            "filled_fields": filled,
            "missing_fields": missing,
            "complete": self.order_complete,
        }
