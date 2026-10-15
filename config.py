"""
Gemensam konfiguration för tidrapporteringsskripten.
Ändra sökvägar och adresser här.
"""

from pathlib import Path

PROJECT_DIR = Path(__file__).parent
DATA_DIR = PROJECT_DIR / "data"

# Indata
GRINDSTONE_FILE = Path(
    r"C:\Users\freha296\OneDrive - Uppsala universitet\Documents\My Grindstone Database.gsjdb"
)

# Utdata
ANTURA_EXCEL_FILE = DATA_DIR / "Tidrapport_Antura.xlsx"
FLEX_EXCEL_FILE = DATA_DIR / "Tidrapport_Flex.xlsx"
CATEGORY_HTML_FILE = DATA_DIR / "Tidrapport_Kategoriserad.html"

# Antura och Chrome (start_chrome.bat har egna kopior av URL och port)
ANTURA_URL = "https://ucrantura.ucr.uu.se/Views/MyPlace/TimeReporting/Reporting.aspx"
CDP_PORT = 9222
CDP_URL = f"http://localhost:{CDP_PORT}"
