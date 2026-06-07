"""Unit-Tests für den Druckerei-Agenten (mit Mock-API)."""

import json
import os
import sys
import types
import unittest
from unittest.mock import MagicMock, patch

# Sicherstellen, dass kein echter API-Key benötigt wird
os.environ.setdefault("ANTHROPIC_API_KEY", "test-key")

import agent as agent_module
from agent import (
    FIELD_LABELS,
    REQUIRED_FIELDS,
    PrintShopAgent,
    _load_orders,
    _save_order,
)


def make_tool_response(extracted: dict) -> MagicMock:
    """Baut eine Mock-Claude-Antwort mit Tool-Use-Block."""
    block = MagicMock()
    block.type = "tool_use"
    block.name = "auftrags_extraktion"
    block.input = extracted

    response = MagicMock()
    response.content = [block]
    return response


def make_text_response(text: str) -> MagicMock:
    block = MagicMock()
    block.type = "text"
    block.text = text

    response = MagicMock()
    response.content = [block]
    return response


class TestFieldConstants(unittest.TestCase):
    def test_required_fields_subset_of_labels(self):
        for f in REQUIRED_FIELDS:
            self.assertIn(f, FIELD_LABELS)

    def test_required_fields_not_empty(self):
        self.assertGreater(len(REQUIRED_FIELDS), 0)


class TestPrintShopAgent(unittest.TestCase):
    def setUp(self):
        # Patch den Client im agent-Modul
        self.mock_client = MagicMock()
        agent_module.client = self.mock_client

    def tearDown(self):
        # Testdateien aufräumen
        from pathlib import Path
        Path("auftraege.json").unlink(missing_ok=True)

    def _make_agent(self) -> PrintShopAgent:
        return PrintShopAgent()

    def test_session_id_generated(self):
        a = self._make_agent()
        self.assertTrue(len(a.session_id) > 0)

    def test_initial_state(self):
        a = self._make_agent()
        self.assertFalse(a.order_complete)
        self.assertEqual(a.current_order, {})

    def test_partial_order_triggers_followup(self):
        """Unvollständiger Auftrag → Rückfrage."""
        a = self._make_agent()

        # Extraktion: nur Produkt bekannt
        self.mock_client.messages.create.side_effect = [
            make_tool_response({
                "produkt": "Flyer",
                "auflage": "",
                "format": "",
                "farbigkeit": "",
                "material": "",
                "veredelung": "",
                "liefertermin": "",
                "besondere_anforderungen": "",
                "fehlende_angaben": [
                    "auflage", "format", "farbigkeit", "material",
                    "veredelung", "liefertermin",
                ],
            }),
            make_text_response("Wie viele Stück benötigen Sie?"),
        ]

        reply = a.chat("Ich möchte Flyer drucken.")
        self.assertFalse(a.order_complete)
        self.assertEqual(a.current_order.get("produkt"), "Flyer")
        self.assertIn("auflage", a.current_order.get("fehlende_angaben", []))

    def test_complete_order_saves_and_confirms(self):
        """Vollständiger Auftrag → Bestätigung und Speicherung."""
        a = self._make_agent()

        self.mock_client.messages.create.side_effect = [
            make_tool_response({
                "produkt": "Visitenkarten",
                "auflage": "1000",
                "format": "90x50mm",
                "farbigkeit": "4/4",
                "material": "350g Karton",
                "veredelung": "Mattlaminat",
                "liefertermin": "20.06.2026",
                "besondere_anforderungen": "",
                "fehlende_angaben": [],
            }),
            make_text_response("Ihr Auftrag DA-XXXX ist bestätigt!"),
        ]

        reply = a.chat(
            "1000 Visitenkarten, 90x50mm, 4/4 farbig, 350g Karton, "
            "Mattlaminat, Lieferung bis 20.06.2026."
        )

        self.assertTrue(a.order_complete)
        self.assertIn("id", a.current_order)
        self.assertIn("erstellt_am", a.current_order)

        # Auftrag muss gespeichert sein
        saved = _load_orders()
        self.assertEqual(len(saved), 1)
        self.assertEqual(saved[0]["produkt"], "Visitenkarten")

    def test_field_merge_across_turns(self):
        """Felder aus mehreren Nachrichten werden korrekt zusammengeführt."""
        a = self._make_agent()

        # Turn 1: nur Produkt + Auflage
        self.mock_client.messages.create.side_effect = [
            make_tool_response({
                "produkt": "Plakat",
                "auflage": "200",
                "format": "",
                "farbigkeit": "",
                "material": "",
                "veredelung": "",
                "liefertermin": "",
                "besondere_anforderungen": "",
                "fehlende_angaben": ["format", "farbigkeit", "material", "veredelung", "liefertermin"],
            }),
            make_text_response("Welches Format?"),
        ]
        a.chat("200 Plakate bitte.")

        # Turn 2: restliche Felder
        self.mock_client.messages.create.side_effect = [
            make_tool_response({
                "produkt": "Plakat",
                "auflage": "200",
                "format": "A1",
                "farbigkeit": "4/0",
                "material": "170g Bilderdruck",
                "veredelung": "Keine",
                "liefertermin": "nächste Woche",
                "besondere_anforderungen": "",
                "fehlende_angaben": [],
            }),
            make_text_response("Bestätigung!"),
        ]
        a.chat("A1, 4/0, 170g Bilderdruck, keine Veredelung, nächste Woche.")

        self.assertTrue(a.order_complete)
        self.assertEqual(a.current_order["produkt"], "Plakat")
        self.assertEqual(a.current_order["format"], "A1")

    def test_status_returns_correct_structure(self):
        a = self._make_agent()
        status = a.get_order_status()
        self.assertIn("session_id", status)
        self.assertIn("filled_fields", status)
        self.assertIn("missing_fields", status)
        self.assertIn("complete", status)
        self.assertFalse(status["complete"])


class TestOrderPersistence(unittest.TestCase):
    def tearDown(self):
        from pathlib import Path
        Path("auftraege.json").unlink(missing_ok=True)

    def test_save_and_load(self):
        order = {"produkt": "Flyer", "auflage": "500", "id": "DA-TEST"}
        _save_order(order)
        loaded = _load_orders()
        self.assertEqual(len(loaded), 1)
        self.assertEqual(loaded[0]["id"], "DA-TEST")

    def test_multiple_saves(self):
        _save_order({"id": "DA-001"})
        _save_order({"id": "DA-002"})
        loaded = _load_orders()
        self.assertEqual(len(loaded), 2)


if __name__ == "__main__":
    unittest.main(verbosity=2)
