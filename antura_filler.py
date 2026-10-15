"""
antura_filler.py
================
Fyller automatiskt i Antura tidrapportering från Tidrapport_Antura.xlsx.

Flöde:
  1. Kör start_chrome.bat  → Chrome startar med fjärrfelsökning aktiverat
  2. Logga in på Antura manuellt (behövs bara första gången)
  3. Kör detta skript: poetry run python antura_filler.py

Skriptet:
  - Hittar senast skickade dag och bearbetar alla veckor från och med den veckan
  - Matchar Excel-rader mot Antura-rader med fuzzy matching
  - Fyller ALDRIG i låsta (redan skickade) celler
  - Trycker ALDRIG på "Skicka" – det gör du manuellt
  - Markerar Excel-filen med gult (osäker matchning) eller rött (ingen matchning)
"""

import asyncio
import re
import sys
from datetime import date, timedelta
from pathlib import Path
from functools import partial

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill

try:
    from rapidfuzz import fuzz
    from rapidfuzz import process as rfprocess
    HAS_RAPIDFUZZ = True
except ImportError:
    import difflib
    HAS_RAPIDFUZZ = False

from playwright.async_api import async_playwright

# ============================================================================
# KONFIGURATION – ändra här om något behöver justeras
# ============================================================================

ANTURA_URL   = "https://ucrantura.ucr.uu.se/Views/MyPlace/TimeReporting/Reporting.aspx"
EXCEL_FILE   = Path(__file__).parent / "data" / "Tidrapport_Antura.xlsx"
CDP_URL      = "http://localhost:9222"

# Fuzzy-matchningsgränser (0–100)
THRESHOLD_OK        = 80   # ≥ 80 → grön  → fylls i automatiskt
THRESHOLD_UNCERTAIN = 55   # 55–79 → gul  → fylls i men flaggas i Excel
#                          # < 55  → röd  → fylls EJ i, flaggas i Excel

# Excel-markeringsfärger
FILL_GREEN  = PatternFill("solid", fgColor="C6EFCE")  # grön  – OK
FILL_YELLOW = PatternFill("solid", fgColor="FFEB9C")  # gul   – osäker
FILL_RED    = PatternFill("solid", fgColor="FFC7CE")  # röd   – ej matchad

# Dagordning i Excel och Antura (index 0–6 = Mån–Sön)
EXCEL_DAY_KEYS = ["Mån", "Tis", "Ons", "Tor", "Fre", "Lör", "Sön"]
ANTURA_DAY_NAMES = ["Måndag", "Tisdag", "Onsdag", "Torsdag", "Fredag", "Lördag", "Söndag"]

# Svenska månadsnamn för datumtolkning
SV_MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "maj": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "okt": 10, "nov": 11, "dec": 12,
}

# Manuella alias: Excel-namn → exakt Antura-etikett (projekt - aktivitet).
# Lägg till rader här för work items som inte matchas automatiskt.
# Exempel:
#   "Svedem": "SOReg - 8.x. Svedem-aktivitet",
MANUAL_ALIASES: dict[str, str] = {
    # "Svedem":            "...",
    # "Swedevox":          "...",
    # "soreg - Plattformslyft": "...",
}


# ============================================================================
# HJÄLPFUNKTIONER
# ============================================================================

def _extract_numbers(s: str) -> list[str]:
    """Plockar ut alla tal ur en sträng, t.ex. ['6', '11'] ur '6.11. Agil'."""
    return re.findall(r"\d+", s)


def fuzzy_score(a: str, b: str) -> float:
    """
    Returnerar likhetsscore 0–100.
    Kombinerar token-sort-ratio med en nummerjustering: om ett tal finns i
    frågesträngen men INTE i kandidaten straffas poängen (-15 per sådant tal).
    Det förhindrar t.ex. att '6.11' matchar '6.12'.
    """
    if HAS_RAPIDFUZZ:
        base = fuzz.token_sort_ratio(a, b)
    else:
        base = difflib.SequenceMatcher(None, a.lower(), b.lower()).ratio() * 100

    # Nummerjustering
    nums_a = _extract_numbers(a)
    nums_b = _extract_numbers(b)
    penalty = sum(15 for n in nums_a if n not in nums_b)
    return max(0.0, base - penalty)


def best_fuzzy_match(query: str, choices: list[str]) -> tuple[str, float]:
    """Returnerar (bästa_match, score) bland valbara strängar."""
    if not choices:
        return ("", 0.0)
    scored = [(c, fuzzy_score(query, c)) for c in choices]
    return max(scored, key=lambda x: x[1])


def parse_sv_date(text: str) -> date | None:
    """
    Tolkar 'Senast skickade dag: 31 Mar, 2026' → date(2026, 3, 31).
    Returnerar None om inget datum hittades.
    """
    m = re.search(r"(\d{1,2})\s+([A-Za-zÅÄÖåäö]+),?\s+(\d{4})", text)
    if not m:
        return None
    try:
        day   = int(m.group(1))
        month = SV_MONTHS.get(m.group(2).lower()[:3])
        year  = int(m.group(3))
        if not month:
            return None
        return date(year, month, day)
    except ValueError:
        return None


def iso_week(d: date) -> int:
    return d.isocalendar()[1]


async def wait_for_enter(prompt: str):
    """Skriver ut en prompt och väntar på Enter utan att blockera event-loopen."""
    loop = asyncio.get_event_loop()
    await loop.run_in_executor(None, partial(input, prompt))


# ============================================================================
# EXCEL – läsa in data
# ============================================================================

def load_excel(
    path: Path,
) -> tuple[dict[int, list[dict]], dict[int, list[dict]], openpyxl.Workbook]:
    """
    Läser alla veckor ur Excel-filen.

    Returnerar:
        weeks  – {veckonr: [{"name": str, "days": {"Mån": float, ...}}, ...]}
        notes  – {veckonr: [{"date": date, "name": str, "note": str, "hours": float}, ...]}
        wb     – öppen Workbook (för att kunna spara markeringar)
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    weeks: dict[int, list[dict]] = {}
    notes: dict[int, list[dict]] = {}

    for sheet_name in wb.sheetnames:
        m = re.match(r"Vecka\s+(\d+)", sheet_name, re.IGNORECASE)
        if not m:
            continue
        week_num = int(m.group(1))
        ws = wb[sheet_name]

        items = []
        in_data = False
        for row in ws.iter_rows(values_only=True):
            if row[0] == "Work Item":
                in_data = True
                continue
            if not in_data:
                continue
            # Stoppa vid Dagsumma eller Anteckningar-sektionen
            if not row[0]:
                continue
            if row[0] in ("Dagsumma", "Anteckningar", "Datum"):
                break

            days = {}
            for i, key in enumerate(EXCEL_DAY_KEYS):
                raw = row[i + 1] if (i + 1) < len(row) else 0
                try:
                    days[key] = round(float(raw or 0), 2)
                except (ValueError, TypeError):
                    days[key] = 0.0

            # Hoppa rader där ALLA dagar är 0 (eller nära 0)
            if all(v < 0.001 for v in days.values()):
                continue

            items.append({"name": str(row[0]).strip(), "days": days})

        if items:
            weeks[week_num] = items

        week_notes = _read_notes(ws)
        if week_notes:
            notes[week_num] = week_notes

    return weeks, notes, wb


def _read_notes(ws) -> list[dict]:
    """
    Läser Anteckningar-sektionen i ett veckoblad.
    Kolumnerna hittas via rubrikraden (Datum, Work Item, Anteckning, Timmar),
    så äldre filer utan Work Item-kolumn fungerar också.
    """
    result = []
    header: dict[str, int] | None = None
    in_notes = False

    for row in ws.iter_rows(values_only=True):
        if row[0] == "Anteckningar":
            in_notes = True
            continue
        if not in_notes:
            continue
        if header is None:
            if row[0] == "Datum":
                header = {str(v): i for i, v in enumerate(row) if v}
            continue
        if not row[0]:
            continue

        def col(name: str):
            idx = header.get(name)
            return row[idx] if idx is not None and idx < len(row) else None

        raw_date = col("Datum")
        result.append({
            "date":  raw_date.date() if hasattr(raw_date, "date") else raw_date,
            "name":  str(col("Work Item") or "").strip(),
            "note":  str(col("Anteckning") or "").strip(),
            "hours": float(col("Timmar") or 0),
        })

    return result


# ============================================================================
# EXCEL – markera matchningsresultat
# ============================================================================

def mark_excel(wb: openpyxl.Workbook, week_num: int, results: list[dict]):
    """
    Markerar rader i Excel-bladet med färg och kommentar:
      Grön  = matchad med hög säkerhet (≥ THRESHOLD_OK %)
      Gul   = osäker matchning, kontrollera
      Röd   = ingen matchning, INTE ifyllt i Antura
    """
    sheet_name = f"Vecka {week_num}"
    if sheet_name not in wb.sheetnames:
        return
    ws = wb[sheet_name]

    # Bygg namn→rad-mappning
    name_to_row: dict[str, tuple] = {}
    for row in ws.iter_rows():
        if row[0].value and str(row[0].value) not in ("Work Item", "Dagsumma"):
            name_to_row[str(row[0].value).strip()] = row

    for res in results:
        target_row = name_to_row.get(res["excel_name"])
        if not target_row:
            continue

        cell = target_row[0]
        status = res["status"]

        if status == "ok":
            cell.fill = FILL_GREEN
            cell.comment = Comment(
                f"✅ Matchad mot Antura-rad:\n\"{res['antura_name']}\"\n(likhet: {res['score']:.0f}%)",
                "antura_filler.py",
            )
        elif status == "uncertain":
            cell.fill = FILL_YELLOW
            cell.comment = Comment(
                f"⚠️ Osäker matchning – kontrollera!\n"
                f"Excel:  \"{res['excel_name']}\"\n"
                f"Antura: \"{res['antura_name']}\"\n"
                f"(likhet: {res['score']:.0f}%)",
                "antura_filler.py",
            )
        else:  # no_match
            cell.fill = FILL_RED
            cell.comment = Comment(
                f"❌ Ingen Antura-rad matchade tillräckligt bra.\n"
                f"Bästa förslag: \"{res['antura_name']}\" ({res['score']:.0f}%)\n"
                f"Fyll i manuellt!",
                "antura_filler.py",
            )


# ============================================================================
# PLAYWRIGHT – läsa Antura-sidan
# ============================================================================

async def get_page_text(page) -> str:
    """Returnerar synlig text från hela sidan."""
    try:
        return await page.inner_text("body") or ""
    except Exception:
        return ""


async def get_last_submitted_date(page) -> date | None:
    """Läser 'Senast skickade dag: X Mmm, YYYY' från sidan."""
    text = await get_page_text(page)
    if "Senast skickade dag" not in text:
        return None
    m = re.search(r"Senast skickade dag[:\s]+(.{5,25})", text)
    if not m:
        return None
    return parse_sv_date(m.group(1))


WEEK_DROPDOWN   = '[name="_toolbar$ctl01$_weekDropDownList"]'
FILTER_DROPDOWN = '[name="_toolbar$ctl02"]'
FILTER_ALL      = "All"

async def get_current_week_num(page) -> int | None:
    """Läser vald vecka direkt från Anturas vecko-dropdown."""
    try:
        return await page.evaluate(f"""
            () => {{
                const sel = document.querySelector('{WEEK_DROPDOWN}');
                if (!sel) return null;
                const opt = sel.options[sel.selectedIndex];
                if (!opt) return null;
                const m = opt.text.match(/^v\\.(\\d{{1,2}}),/);
                return m ? parseInt(m[1]) : null;
            }}
        """)
    except Exception:
        return None


async def get_week_option_value(page, target_week: int) -> str | None:
    """Returnerar dropdown-värdet (ett datum-string) för given vecka, eller None."""
    try:
        return await page.evaluate(f"""
            () => {{
                const sel = document.querySelector('{WEEK_DROPDOWN}');
                if (!sel) return null;
                for (const opt of sel.options) {{
                    const m = opt.text.match(/^v\\.(\\d{{1,2}}),/);
                    if (m && parseInt(m[1]) === {target_week}) return opt.value;
                }}
                return null;
            }}
        """)
    except Exception:
        return None



async def read_antura_rows(page) -> list[dict]:
    """
    Läser alla aktivitetsrader från Antura via ett enda JavaScript-anrop.

    Antura renderar tidsfälten som:
        <div class="reportingField" id="...FieldsBlock_N">
            <input type="text" ...>   ← redigerbara fält
            eller
            <span ...>               ← summerings-/projektraderna
        </div>
    där N är dagindex (0=Mån, 1=Tis, … 6=Sön).

    Rader vars .reportingField-divar saknar redigerbara inputs
    (ej disabled, ej readonly) behandlas som projekthuvuden.
    Det filtrerar automatiskt bort projektsammanfattningsrader
    som har .reportingField men inga redigerbara fält.
    """
    return await page.evaluate("""
        () => {
            const SKIP = new Set(['Projekt/Aktivitet', 'Total', 'Arbetstid', 'Dagsumma']);
            const rows_data = [];
            let current_project = '';

            for (const row of document.querySelectorAll('tr')) {
                const cells = row.querySelectorAll('td');
                if (!cells.length) continue;

                const first_text = (cells[0].innerText || '').trim();
                if (!first_text || SKIP.has(first_text)) continue;

                const reporting_fields = row.querySelectorAll('.reportingField');
                if (!reporting_fields.length) {
                    current_project = first_text;
                    continue;
                }

                const input_keys = new Array(7).fill(null);
                let has_any_input = false;

                for (const div of reporting_fields) {
                    const div_id = div.id || '';
                    if (!div_id.includes('FieldsBlock_')) continue;
                    const suffix = div_id.split('FieldsBlock_').pop();
                    const day_idx = parseInt(suffix, 10);
                    if (isNaN(day_idx) || day_idx < 0 || day_idx > 6) continue;

                    // Bara redigerbara inputs räknas – ej disabled, ej readonly
                    const inp = div.querySelector('input:not([disabled]):not([readonly])');
                    if (!inp) continue;

                    input_keys[day_idx] = div_id;
                    has_any_input = true;
                }

                if (!has_any_input) {
                    // Inga redigerbara fält → behandla som projektrubrik
                    current_project = first_text;
                    continue;
                }

                const combined = current_project
                    ? `${current_project} - ${first_text}`
                    : first_text;

                rows_data.push({
                    label:      combined,
                    project:    current_project,
                    activity:   first_text,
                    input_keys: input_keys,
                });
            }

            return rows_data;
        }
    """)


# ============================================================================
# PLAYWRIGHT – navigation
# ============================================================================

async def navigate_to_week(page, target_week: int) -> bool:
    """
    Navigerar till rätt vecka via Anturas vecko-dropdown.
    Väntar på att AJAX-uppdateringen är helt klar (networkidle, max 10 s)
    innan den returnerar – annars kan DOM:en vara halvrenderad.
    Returnerar True om lyckat.
    """
    opt_value = await get_week_option_value(page, target_week)
    if not opt_value:
        print(f"    ❌ Vecka {target_week} saknas i dropdown (finns den i Excel men inte i Antura?)")
        return False

    current = await get_current_week_num(page)
    if current == target_week:
        print(f"    Redan på vecka {target_week}")
        return True

    print(f"    Väljer vecka {target_week} i dropdown (value={opt_value})")

    # Välj rätt option – Playwright triggar change-eventet automatiskt
    await page.select_option(WEEK_DROPDOWN, value=opt_value)

    # Vänta tills nätverkstrafiken har lagt sig (AJAX klar), max 10 s
    try:
        await page.wait_for_load_state("networkidle", timeout=10_000)
    except Exception:
        print("    ⚠️  networkidle-timeout – fortsätter ändå")

    # Bekräfta att dropdown faktiskt visar rätt vecka
    try:
        await page.wait_for_function(
            f"""
            () => {{
                const sel = document.querySelector('{WEEK_DROPDOWN}');
                if (!sel) return false;
                const opt = sel.options[sel.selectedIndex];
                if (!opt) return false;
                const m = opt.text.match(/^v\\.(\\d{{1,2}}),/);
                return m && parseInt(m[1]) === {target_week};
            }}
            """,
            timeout=5_000,
        )
    except Exception:
        print(f"    ⚠️  Dropdown visar inte vecka {target_week} – fortsätter ändå")

    return True


async def ensure_filter_all(page):
    """
    Kontrollerar att visningsfiltret är satt till 'Alla'.
    Väljer om det om det inte redan är valt och väntar på att AJAX är klar.
    """
    try:
        current = await page.evaluate(f"""
            () => {{
                const sel = document.querySelector('{FILTER_DROPDOWN}');
                return sel ? sel.value : null;
            }}
        """)
        if current == FILTER_ALL:
            return  # Redan rätt, inget att göra
        print("    Filter är inte 'Alla' – ändrar...")
        await page.select_option(FILTER_DROPDOWN, value=FILTER_ALL)
        try:
            await page.wait_for_load_state("networkidle", timeout=10_000)
        except Exception:
            await page.wait_for_timeout(2_000)
    except Exception as e:
        print(f"    ⚠️  Kunde inte sätta filter till 'Alla': {e}")


# ============================================================================
# PLAYWRIGHT – fylla i celler
# ============================================================================

async def fill_input(page, input_key: str, value: float):
    """
    Sätter värdet i ett Antura-inmatningsfält via JavaScript.

    input_key är ID:t på den omslutande <div class="reportingField">.
    Inputs i Antura saknar name/id, så vi hittar dem via sin föräldra-div:
        document.getElementById(div_id).querySelector('input')

    Vi använder JS direkt (inp.value = ...) istället för Playwright's fill()
    + Tab, eftersom Tab triggar onchange → __doPostBack → AJAX-uppdatering
    som nollställer alla värden vi just fyllt i.

    ASP.NET hämtar ändå input-värdena direkt från DOM:en när användaren
    klickar "Skicka vecka", så direktsättning fungerar korrekt.
    """
    if value < 0.001:
        return  # Lämna 0-fält orörda

    val_str = f"{value:.2f}".replace(".", ",")

    found = await page.evaluate(f"""
        () => {{
            const div = document.getElementById('{input_key}');
            if (!div) return false;
            const inp = div.querySelector('input');
            if (!inp) return false;
            inp.value = '{val_str}';
            inp.dispatchEvent(new Event('input',  {{ bubbles: true }}));
            inp.dispatchEvent(new Event('change', {{ bubbles: true }}));
            return true;
        }}
    """)

    if not found:
        print(f"    ⚠️  Fält '{input_key}' hittades inte i DOM:en")


# ============================================================================
# MATCHNING
# ============================================================================

def match_items(
    excel_items: list[dict], antura_labels: list[str]
) -> list[dict]:
    """
    Matchar varje Excel-rad mot bästa Antura-etiketten.

    Prioritetsordning:
      0. Manuellt alias i MANUAL_ALIASES → alltid 100 %, status "ok"
      1. Fuzzy på hela strängen
      2. Fuzzy på bara aktivitetsdelen (efter sista " - ")
      3. Fuzzy på bara projektdelen (före första " - ")
    """
    label_set = set(antura_labels)
    results = []

    for item in excel_items:
        name = item["name"]

        # 0. Manuellt alias
        if name in MANUAL_ALIASES:
            alias = MANUAL_ALIASES[name]
            if alias in label_set:
                results.append({
                    "excel_name":  name,
                    "antura_name": alias,
                    "score":       100.0,
                    "days":        item["days"],
                    "status":      "ok",
                })
                continue
            print(f"  ⚠️  Alias för '{name}' pekar på okänt Antura-namn '{alias}'")

        # 1. Hela strängen
        best_label, best_score = best_fuzzy_match(name, antura_labels)

        # 2. Aktivitetsdelen
        if " - " in name:
            act = name.rsplit(" - ", 1)[1]
            lbl, scr = best_fuzzy_match(act, antura_labels)
            if scr > best_score:
                best_label, best_score = lbl, scr

        # 3. Projektdelen
        if " - " in name:
            proj = name.split(" - ", 1)[0]
            lbl, scr = best_fuzzy_match(proj, antura_labels)
            if scr > best_score:
                best_label, best_score = lbl, scr

        if best_score >= THRESHOLD_OK:
            status = "ok"
        elif best_score >= THRESHOLD_UNCERTAIN:
            status = "uncertain"
        else:
            status = "no_match"

        results.append({
            "excel_name":  name,
            "antura_name": best_label,
            "score":       best_score,
            "days":        item["days"],
            "status":      status,
        })

    return results


# ============================================================================
# BEARBETA EN VECKA
# ============================================================================

async def process_week(
    page,
    week_num: int,
    excel_items: list[dict],
) -> list[dict]:
    """
    Navigerar till veckan, väntar på klartecken från användaren,
    scannar sidan och fyller i alla matchade, olåsta fält.
    """

    print(f"\n  → Navigerar till vecka {week_num}…")
    ok = await navigate_to_week(page, week_num)
    if not ok:
        print(f"  ❌ Kunde inte navigera till vecka {week_num}")
        return []

    # Säkerställ att alla projekt visas
    await ensure_filter_all(page)

    # ── Vänta på att användaren bekräftar att sidan är klar ──────────────────
    print()
    await wait_for_enter(
        f"  ⏸  Vecka {week_num} vald. Vänta tills sidan är klar i Chrome,\n"
        f"     tryck sedan ENTER för att börja fylla i… "
    )
    print()

    # Läs Antura-rader (hittar inputs via .reportingField-divars FieldsBlock_N-ID)
    antura_rows = await read_antura_rows(page)
    print(f"  Aktivitetsrader: {len(antura_rows)} st")
    if not antura_rows:
        print("  ❌ Inga aktivitetsrader hittades")
        return []

    antura_labels = [r["label"] for r in antura_rows]
    label_to_row  = {r["label"]: r for r in antura_rows}

    # Matcha
    results = match_items(excel_items, antura_labels)

    print(f"  Matchningsresultat ({len(results)} Excel-rader):")
    for r in results:
        icon = {"ok": "✅", "uncertain": "⚠️ ", "no_match": "❌"}[r["status"]]
        print(
            f"    {icon} {r['excel_name'][:45]:<45} → "
            f"{r['antura_name'][:45]:<45} ({r['score']:.0f}%)"
        )

    # Fyll i
    filled = 0
    skipped = 0
    for res in results:
        if res["status"] == "no_match":
            skipped += 1
            continue

        antura_row = label_to_row.get(res["antura_name"])
        if not antura_row:
            print(f"  ⚠️  '{res['antura_name'][:50]}' saknas i label_to_row")
            skipped += 1
            continue

        for day_idx, day_key in enumerate(EXCEL_DAY_KEYS):
            value = res["days"].get(day_key, 0.0)
            if value < 0.001:
                continue

            input_key = antura_row["input_keys"][day_idx]
            if input_key is None:
                print(f"  ⚠️  {res['antura_name'][:35]} / {day_key}: input_key=None (låst?)")
                continue

            await fill_input(page, input_key, value)
            filled += 1
            print(f"  Fyll {day_key} '{input_key[-30:]}' = {value:.2f}")

    print(f"\n  → {filled} celler ifyllda, {skipped} rader ej matchade/ifyllda")
    return results


def print_week_notes(week_notes: list[dict], results: list[dict]):
    """
    Listar veckans Grindstone-kommentarer så att de kan läggas in manuellt
    i Antura. Visar vilken Antura-rad work itemet matchades mot.
    """
    if not week_notes:
        print("\n  📝 Inga kommentarer i Grindstone denna vecka.")
        return

    excel_to_antura = {
        r["excel_name"]: r["antura_name"] for r in results if r["status"] != "no_match"
    }

    print(f"\n  📝 Kommentarer att lägga in manuellt i Antura ({len(week_notes)} st):")
    for n in week_notes:
        d = n["date"]
        day = f"{EXCEL_DAY_KEYS[d.weekday()]} {d.isoformat()}" if isinstance(d, date) else str(d)
        antura_name = excel_to_antura.get(n["name"])
        target = f"→ {antura_name}" if antura_name else "→ (ingen matchad Antura-rad)"
        print(f"    {day}  {n['name'] or '(okänt work item)'}")
        print(f"      {target}")
        # Flerradiga kommentarer indenteras så att de hålls ihop visuellt
        note_text = "\n       ".join(n["note"].splitlines())
        hours = f"{n['hours']:.2f}".replace(".", ",")
        print(f"      \"{note_text}\"  ({hours} h)")


# ============================================================================
# HUVUDFUNKTION
# ============================================================================

async def main():
    print("=" * 65)
    print("  Antura Tidrapporterings-fyllare")
    print("=" * 65)

    if not HAS_RAPIDFUZZ:
        print("ℹ️  rapidfuzz saknas – använder difflib (lägre precision).")
        print("   Installera: poetry add rapidfuzz\n")

    # ── Ladda Excel ──────────────────────────────────────────────────────────
    if not EXCEL_FILE.exists():
        print(f"❌ Hittade inte {EXCEL_FILE}")
        sys.exit(1)

    print(f"\nLäser {EXCEL_FILE.name}…")
    excel_weeks, excel_notes, wb = load_excel(EXCEL_FILE)
    if not excel_weeks:
        print("❌ Inga veckor med data hittades i Excel-filen.")
        sys.exit(1)
    print(f"  Veckor med data: {sorted(excel_weeks)}")

    # ── Anslut till Chrome ───────────────────────────────────────────────────
    print(f"\nAnsluter till Chrome ({CDP_URL})…")
    print("  (Se till att start_chrome.bat är igång och Antura är inloggad)")

    async with async_playwright() as p:
        try:
            browser = await p.chromium.connect_over_cdp(CDP_URL)
        except Exception as exc:
            print(f"\n❌ Kunde inte ansluta: {exc}")
            print("   Kontrollera att start_chrome.bat körs och att porten är 9222.")
            sys.exit(1)

        # Hitta eller öppna Antura-fliken
        page = None
        for ctx in browser.contexts:
            for pg in ctx.pages:
                if "TimeReporting" in pg.url or "antura" in pg.url.lower():
                    page = pg
                    break
            if page:
                break

        if not page:
            ctx  = browser.contexts[0] if browser.contexts else await browser.new_context()
            page = await ctx.new_page()
            print("  Öppnar Antura…")
            await page.goto(ANTURA_URL)
            await page.wait_for_load_state("networkidle", timeout=30_000)

        await page.bring_to_front()
        print("  ✅ Ansluten")

        if "TimeReporting" not in page.url:
            await page.goto(ANTURA_URL)
            await page.wait_for_load_state("networkidle", timeout=30_000)

        await page.wait_for_timeout(1_000)

        # ── Hitta senast skickade dag ─────────────────────────────────────────
        last_sub = await get_last_submitted_date(page)
        if last_sub:
            print(f"\nSenast skickade dag: {last_sub.strftime('%d %b %Y')} (vecka {iso_week(last_sub)})")
        else:
            print("\n⚠️  Kunde inte läsa 'Senast skickade dag' – bearbetar alla veckor i Excel")

        # ── Välj veckor att bearbeta ──────────────────────────────────────────
        if last_sub:
            # Börja från veckan som innehåller dagen EFTER senast skickade dag.
            # Om senaste skickade dag är söndag (sista dagen i veckan) hamnar
            # nästa dag i nästkommande vecka – och den veckan är den första att fylla.
            first_unfilled = last_sub + timedelta(days=1)
            first_week = iso_week(first_unfilled)
            weeks_to_do = sorted(w for w in excel_weeks if w >= first_week)
        else:
            weeks_to_do = sorted(excel_weeks)

        if not weeks_to_do:
            print("✅ Inga veckor att bearbeta – allt är redan klart!")
            return

        print(f"Veckor att bearbeta: {weeks_to_do}")

        # ── Bearbeta varje vecka ──────────────────────────────────────────────
        all_results: dict[int, list[dict]] = {}
        for i, wk in enumerate(weeks_to_do):
            sep = "─" * 65
            print(f"\n{sep}")
            print(f"  Vecka {wk}  ({i + 1} av {len(weeks_to_do)})")
            print(sep)
            res = await process_week(page, wk, excel_weeks[wk])
            all_results[wk] = res
            print_week_notes(excel_notes.get(wk, []), res)

            # Spara Excel-markeringar efter varje vecka (ger delresultat vid avbrott)
            if res:
                mark_excel(wb, wk, res)
                wb.save(EXCEL_FILE)

            # Paus mellan veckor (inte efter sista)
            if i < len(weeks_to_do) - 1:
                next_wk = weeks_to_do[i + 1]
                print()
                await wait_for_enter(
                    f"  ✅ Vecka {wk} är ifylld. Granska i Antura och skicka in om allt ser bra ut.\n"
                    f"     Tryck ENTER när du är redo att gå vidare till vecka {next_wk}… "
                )

        # ── Klar ─────────────────────────────────────────────────────────────
        print(f"\n{'=' * 65}")
        print(f"✅ {EXCEL_FILE.name} sparad med markeringar")
        print()
        print("Klart! Granska gula/röda rader i Excel.")
        print("Tryck INTE på 'Skicka' i Antura – det gör du manuellt.")


if __name__ == "__main__":
    asyncio.run(main())
