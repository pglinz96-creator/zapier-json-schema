"""
CLI-Modus für die Druckerei-Agentic-AI.
Startet mit: python cli.py
"""

import os
import sys
from agent import PrintShopAgent, FIELD_LABELS, REQUIRED_FIELDS

RESET  = "\033[0m"
BOLD   = "\033[1m"
BLUE   = "\033[34m"
GREEN  = "\033[32m"
YELLOW = "\033[33m"
CYAN   = "\033[36m"
RED    = "\033[31m"
DIM    = "\033[2m"


def print_banner():
    print(f"""
{BLUE}{BOLD}╔══════════════════════════════════════════╗
║     🖨️  Druckerei Auftragsassistent     ║
║         Powered by Claude AI             ║
╚══════════════════════════════════════════╝{RESET}
{DIM}Tippen Sie Ihren Druckauftrag ein.
Beenden mit 'exit' oder Strg+C.{RESET}
""")


def print_agent(text: str):
    print(f"\n{BLUE}🖨️  Agent:{RESET}")
    print(f"  {text.replace(chr(10), chr(10) + '  ')}\n")


def print_status(agent: PrintShopAgent):
    status = agent.get_order_status()
    filled = status["filled_fields"]
    missing = status["missing_fields"]

    print(f"\n{DIM}{'─' * 44}")
    print(f"  Auftragsfortschritt: {len(filled)}/{len(REQUIRED_FIELDS)} Felder")
    for f in REQUIRED_FIELDS:
        if f in filled:
            print(f"  {GREEN}✓{RESET} {FIELD_LABELS[f]}: {DIM}{filled[f]}{RESET}")
        else:
            print(f"  {YELLOW}○{RESET} {FIELD_LABELS[f]}")
    print(f"{'─' * 44}{RESET}\n")


def main():
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print(f"{RED}Fehler: ANTHROPIC_API_KEY nicht gesetzt.{RESET}")
        print("Bitte setzen Sie die Umgebungsvariable und versuchen Sie es erneut.")
        sys.exit(1)

    print_banner()
    agent = PrintShopAgent()

    greeting = (
        "Willkommen bei Ihrer Druckerei! Ich bin Ihr persönlicher Auftragsassistent.\n"
        "Bitte schildern Sie mir Ihren Druckbedarf — was möchten Sie drucken, "
        "in welcher Menge und bis wann?\n"
        "Je mehr Details Sie angeben, desto schneller kann ich Ihren Auftrag aufnehmen."
    )
    print_agent(greeting)

    while True:
        try:
            user_input = input(f"{CYAN}Sie:{RESET} ").strip()
        except (KeyboardInterrupt, EOFError):
            print(f"\n{DIM}Sitzung beendet.{RESET}")
            break

        if not user_input:
            continue

        if user_input.lower() in ("exit", "quit", "beenden"):
            print(f"\n{DIM}Auf Wiedersehen!{RESET}")
            break

        if user_input.lower() == "status":
            print_status(agent)
            continue

        print(f"\n{DIM}Agent analysiert...{RESET}", end="\r")
        reply = agent.chat(user_input)
        print_agent(reply)
        print_status(agent)

        if agent.order_complete:
            print(f"{GREEN}{BOLD}✅ Auftrag erfolgreich aufgenommen und gespeichert!{RESET}\n")
            break


if __name__ == "__main__":
    main()
