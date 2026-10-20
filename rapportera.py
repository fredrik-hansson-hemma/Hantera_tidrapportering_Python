"""
Hela tidrapporteringsflödet i ett fönster:
  1. Genererar Excel-underlagen från Grindstone
  2. Startar Chrome och väntar på Antura-inloggning vid behov
  3. Fyller i Antura vecka för vecka

Skickar aldrig in något – det gör du själv i Antura.
"""

import asyncio
import sys

import antura_filler
import Hantera_tidrapportering


def main():
    print("\n[1/2] Genererar Excel-underlagen\n")
    Hantera_tidrapportering.main()

    print("\n[2/2] Fyller i Antura\n")
    asyncio.run(antura_filler.main())


if __name__ == "__main__":
    # Säkerställ att åäö och symboler kan skrivas ut oavsett konsolens teckentabell
    sys.stdout.reconfigure(encoding="utf-8")
    main()
