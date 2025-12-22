"""
Tidrapportering - Python-konvertering från SAS
Hanterar tidrapporter från TSV-fil och skapar rapporter för Antura och flextid
"""

import polars as pl
from datetime import datetime, timedelta
import time
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import sys
from weekly_category_report import create_category_report

# ============================================================================
# KONFIGURATION
# ============================================================================

INPUT_FILE = r"C:\Users\freha296\Documents\Administrativt\All_my_time.tsv"
ANTURA_OUTPUT_FILE = "Tidrapport_Antura.xlsx"
FLEX_OUTPUT_FILE = "Tidrapport_Flex.xlsx"
CATEGORY_OUTPUT_FILE = "Tidrapport_Kategoriserad.html"

# ============================================================================
# FUNKTIONER
# ============================================================================


def load_time_data(filepath: str) -> pl.DataFrame:
    """
    Läser in tidrapportdata från TSV-fil
    Motsvarar DATA WORK.All_my_time i SAS
    """
    # Läs filen rad för rad för att hantera kommatecken i texten korrekt
    import csv

    rows = []
    with open(filepath, "r", encoding="utf-8") as f:
        # TSV-filen använder tab som separator, men SAS-koden visar DLM='2c'x vilket är komma
        # Låt oss testa med komma först, sedan tab om det inte fungerar
        f.seek(0)
        first_line = f.readline()

        # Räkna antal tabs vs komman i första raden
        if "\t" in first_line:
            delimiter = "\t"
        else:
            delimiter = ","

        f.seek(0)
        reader = csv.reader(f, delimiter=delimiter, quotechar='"')
        headers = next(reader)  # Läs headers

        for row in reader:
            if len(row) >= 4:  # Minst Work Item, Start, End, Duration
                rows.append(row)

    # Skapa DataFrame
    df = pl.DataFrame(
        {
            "Work Item": [row[0] if len(row) > 0 else "" for row in rows],
            "Start": [row[1] if len(row) > 1 else "" for row in rows],
            "End": [row[2] if len(row) > 2 else "" for row in rows],
            "Duration": [row[3] if len(row) > 3 else "0:00:00" for row in rows],
            "Notes": [row[4] if len(row) > 4 else "" for row in rows],
        }
    )

    # Konvertera datum och tid
    df = df.with_columns(
        [
            pl.col("Start")
            .str.strptime(pl.Datetime, format="%Y-%m-%d %H:%M:%S", strict=False)
            .alias("Start"),
            pl.col("End")
            .str.strptime(pl.Datetime, format="%Y-%m-%d %H:%M:%S", strict=False)
            .alias("End"),
            # Duration är i format HH:MM:SS, konvertera till sekunder
            pl.col("Duration")
            .str.split(":")
            .map_elements(
                lambda x: int(x[0]) * 3600 + int(x[1]) * 60 + int(x[2])
                if len(x) == 3
                else 0,
                return_dtype=pl.Int64,
            )
            .alias("Duration"),
        ]
    )

    return df


def add_derived_columns(df: pl.DataFrame) -> pl.DataFrame:
    """
    Lägger till härledda kolumner
    Motsvarar WORK.DERIVED_COLUMNS i SAS
    """
    df = df.with_columns(
        [
            pl.col("Start").dt.date().alias("startdatum"),
            pl.col("Start").dt.month().alias("manad"),
        ]
    )

    return df


def validate_single_month(df: pl.DataFrame) -> None:
    """
    Kontrollerar att varje månadsnummer (1-12) endast förekommer EN gång
    Dvs maj 2024 eller maj 2025 går bra, men inte båda samtidigt
    Motsvarar varna_om_antal_manader makrot
    """
    # Räkna hur många gånger varje månad (oavsett år) förekommer
    month_counts = (
        df.select(
            [
                pl.col("startdatum").dt.year().alias("ar"),
                pl.col("startdatum").dt.month().alias("manad"),
            ]
        )
        .unique()
        .group_by("manad")
        .agg(pl.col("ar").alias("ar_lista"), pl.len().alias("antal_ar"))
    )

    # Hitta månader som förekommer i flera år
    duplicates = month_counts.filter(pl.col("antal_ar") > 1)

    if len(duplicates) > 0:
        print("ERROR: Följande månader förekommer i flera år:")
        for row in duplicates.iter_rows(named=True):
            manad_namn = [
                "",
                "Jan",
                "Feb",
                "Mar",
                "Apr",
                "Maj",
                "Jun",
                "Jul",
                "Aug",
                "Sep",
                "Okt",
                "Nov",
                "Dec",
            ][row["manad"]]
            ar_lista = row["ar_lista"]
            print(f"  - {manad_namn} (månad {row['manad']}): finns i år {ar_lista}")
        print("\nERROR: Varje månad får bara förekomma ett år")
        print("ERROR: Sessionen avbryts")
        sys.exit(1)

    # Visa vilka månader som finns
    all_months = (
        df.select(
            [
                pl.col("startdatum").dt.year().alias("ar"),
                pl.col("startdatum").dt.month().alias("manad"),
            ]
        )
        .unique()
        .sort(["ar", "manad"])
    )

    print("Data validerad - funna månader:")
    for row in all_months.iter_rows(named=True):
        print(f"  - {row['ar']}-{row['manad']:02d}")


def create_all_days(df: pl.DataFrame) -> pl.DataFrame:
    """
    Skapar upp alla dagar för att fylla ut dataset (för lediga dagar)
    Motsvarar data alla_dagar i SAS
    """
    min_datum = df.select(pl.col("startdatum").min()).item()
    max_datum = df.select(pl.col("startdatum").max()).item()

    # Första och sista dagen i månaden
    from dateutil.relativedelta import relativedelta

    forsta = min_datum.replace(day=1)
    sista = (max_datum.replace(day=1) + relativedelta(months=1)) - timedelta(days=1)

    # Skapa alla datum
    date_range = pl.date_range(forsta, sista, interval="1d", eager=True)

    all_days = pl.DataFrame(
        {"startdatum": date_range, "Duration": [0] * len(date_range)}
    )

    return all_days


def create_main_query(df: pl.DataFrame, all_days: pl.DataFrame) -> pl.DataFrame:
    """
    Skapar huvuddataset med alla beräknade kolumner
    Motsvarar WORK.QUERY_FOR_THISMONTH_DERIVED_COLU i SAS
    """
    # Full outer join mellan tiddata och alla dagar
    merged = df.join(all_days, on="startdatum", how="full", suffix="_all")

    merged = merged.with_columns(
        [
            # Duration - använd från tiddata om den finns, annars från all_days
            pl.when(pl.col("Duration").is_not_null())
            .then(pl.col("Duration"))
            .otherwise(pl.col("Duration_all"))
            .alias("Duration"),
            # DurationTimmar - konvertera sekunder till timmar
            (pl.coalesce(["Duration", "Duration_all"]) / 3600.0).alias(
                "DurationTimmar"
            ),
            # startdatum - använd från tiddata om den finns
            pl.coalesce(["startdatum", "startdatum_all"]).alias("startdatum"),
        ]
    )

    # Lägg till vecka, veckodag, år etc
    merged = merged.with_columns(
        [
            pl.col("startdatum").dt.week().alias("Vecka"),
            pl.col("startdatum").dt.strftime("%a").alias("veckodag"),
            pl.col("startdatum").dt.month().alias("manad"),
            pl.col("startdatum").dt.year().alias("ar"),
        ]
    )

    # Välj relevanta kolumner
    cols = [
        "Work Item",
        "Start",
        "End",
        "Duration",
        "DurationTimmar",
        "startdatum",
        "Vecka",
        "veckodag",
        "manad",
        "ar",
        "Notes",
    ]

    return merged.select([c for c in cols if c in merged.columns])


def add_time_type(df: pl.DataFrame) -> pl.DataFrame:
    """
    Lägger till kategori för typ av tid
    Motsvarar WORK.typ_av_tid i SAS
    """
    twelve_months_ago = datetime.now().date() - timedelta(days=365)

    df = df.filter(pl.col("startdatum") > twelve_months_ago)

    df = df.with_columns(
        [
            pl.when(~pl.col("Work Item").str.starts_with("Från"))
            .then(pl.lit("Arbetstid"))
            .otherwise(pl.col("Work Item"))
            .alias("Typ_av_tid")
        ]
    )

    return df


def create_antura_report(df: pl.DataFrame, output_file: str):
    """
    Skapar Antura-rapport i Excel
    Motsvarar PROC TABULATE för Antura i SAS
    """
    # Filtrera till förra månaden
    from dateutil.relativedelta import relativedelta

    today = datetime.now().date()
    first_day_last_month = today.replace(day=1) - relativedelta(months=1)

    df_filtered = df.filter(pl.col("startdatum") >= first_day_last_month)

    # Skapa Excel-arbetsbok
    wb = openpyxl.Workbook()

    # Ta bort standard-sheet
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    # Hämta unika veckor
    weeks = df_filtered.select(["ar", "Vecka"]).unique().sort(["ar", "Vecka"])

    for row in weeks.iter_rows():
        ar, vecka = row

        # Skapa nytt sheet för varje vecka
        ws = wb.create_sheet(title=f"Vecka {vecka}")

        # Filtrera data för denna vecka
        week_data = df_filtered.filter(
            (pl.col("ar") == ar) & (pl.col("Vecka") == vecka)
        )

        # Skapa pivot-tabell med Polars SQL
        pivot_df = week_data.sql("""
            SELECT 
                "Work Item",
                SUM(CASE WHEN veckodag = 'Mon' THEN DurationTimmar ELSE 0 END) as Mån,
                SUM(CASE WHEN veckodag = 'Tue' THEN DurationTimmar ELSE 0 END) as Tis,
                SUM(CASE WHEN veckodag = 'Wed' THEN DurationTimmar ELSE 0 END) as Ons,
                SUM(CASE WHEN veckodag = 'Thu' THEN DurationTimmar ELSE 0 END) as Tor,
                SUM(CASE WHEN veckodag = 'Fri' THEN DurationTimmar ELSE 0 END) as Fre,
                SUM(CASE WHEN veckodag = 'Sat' THEN DurationTimmar ELSE 0 END) as Lör,
                SUM(CASE WHEN veckodag = 'Sun' THEN DurationTimmar ELSE 0 END) as Sön,
                SUM(DurationTimmar) as Veckosumma
            FROM self
            WHERE "Work Item" IS NOT NULL
            GROUP BY "Work Item"
            ORDER BY "Work Item"
        """)

        # Konvertera till dict för enklare hantering
        pivot_dict = pivot_df.to_dicts()

        # Lägg till dagsummor
        day_columns = ["Mån", "Tis", "Ons", "Tor", "Fre", "Lör", "Sön", "Veckosumma"]
        day_totals = ["Dagsumma"] + [
            sum(row.get(col, 0) or 0 for row in pivot_dict) for col in day_columns
        ]

        # Skriv till Excel
        ws.append(["Vecka " + str(vecka)])
        ws.append([])

        # Headers
        ws.append(["Work Item"] + day_columns)

        # Data
        for row_dict in pivot_dict:
            ws.append(
                [row_dict.get("Work Item", "")]
                + [row_dict.get(col, 0) for col in day_columns]
            )

        # Dagsummor
        ws.append(day_totals)

        # Formatera
        for row in ws.iter_rows(min_row=3, max_row=3):
            for cell in row:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="CCCCCC", fill_type="solid")

        # Formatera siffror
        for row in ws.iter_rows(min_row=4, min_col=2):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0.00"

        # Anteckningar
        notes_data = week_data.filter(
            pl.col("Notes").is_not_null() & (pl.col("Notes") != "")
        )

        if len(notes_data) > 0:
            ws.append([])
            ws.append(["Anteckningar"])
            ws.append(["Datum", "Anteckning", "Timmar"])

            for row in notes_data.select(
                ["startdatum", "Notes", "DurationTimmar"]
            ).iter_rows():
                ws.append(list(row))

    # Spara
    wb.save(output_file)
    print(f"Antura-rapport sparad: {output_file}")


def create_flex_report(df: pl.DataFrame, output_file: str):
    """
    Skapar flextidsrapport i Excel med tid i HH:MM format
    Motsvarar PROC TABULATE för flex i SAS
    """

    print(f"Flex rapport - antal rader: {len(df)}")
    print(f"Kolumner: {df.columns}")
    # print(f"Första raden: {df.head(1)}")  # Kommenterad pga encoding-problem

    # Skapa Excel-arbetsbok
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Flextid"

    # Gruppera data per månad
    months = df.select(["manad", "ar"]).unique().sort(["ar", "manad"])

    row_num = 1

    for month_row in months.iter_rows(named=True):
        ar = month_row["ar"]
        manad = month_row["manad"]

        # Filtrera data för denna månad
        month_data = df.filter((pl.col("ar") == ar) & (pl.col("manad") == manad))

        # Skapa pivot med Polars SQL
        pivot_df = month_data.sql("""
            SELECT 
                startdatum,
                veckodag,
                MAX(CASE WHEN Typ_av_tid = 'Arbetstid' THEN Duration ELSE NULL END) as Arbetstid,
                MAX(CASE WHEN Typ_av_tid LIKE 'Från%' THEN Duration ELSE NULL END) as Frånvaro,
                SUM(Duration) as Total
            FROM self
            GROUP BY startdatum, veckodag
            ORDER BY startdatum
        """)

        # Konvertera sekunder till HH:MM (avrunda till närmaste minut)
        def seconds_to_hhmm(seconds):
            if seconds is None or seconds == 0:
                return ""
            # Avrunda till närmaste minut (lägg till 30 sekunder innan konvertering)
            rounded_seconds = int(seconds) + 30
            hours = rounded_seconds // 3600
            minutes = (rounded_seconds % 3600) // 60
            return f"{hours}:{minutes:02d}"

        # Månadstitel
        manad_namn = [
            "",
            "Januari",
            "Februari",
            "Mars",
            "April",
            "Maj",
            "Juni",
            "Juli",
            "Augusti",
            "September",
            "Oktober",
            "November",
            "December",
        ][manad]
        ws.cell(row_num, 1, f"{manad_namn} {ar}")
        ws.cell(row_num, 1).font = Font(bold=True, size=14)
        row_num += 2

        # Headers
        headers = ["Datum", "Veckodag", "Arbetstid", "Frånvaro", "Total"]
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row_num, col_num, header)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="CCCCCC", fill_type="solid")
        row_num += 1

        # Data rader
        total_arbetstid = 0
        total_franvaro = 0
        total_total = 0

        for data_row in pivot_df.iter_rows(named=True):
            ws.cell(row_num, 1, data_row["startdatum"])
            ws.cell(row_num, 2, data_row["veckodag"])
            ws.cell(row_num, 3, seconds_to_hhmm(data_row["Arbetstid"]))
            ws.cell(row_num, 4, seconds_to_hhmm(data_row["Frånvaro"]))
            ws.cell(row_num, 5, seconds_to_hhmm(data_row["Total"]))

            total_arbetstid += data_row["Arbetstid"] or 0
            total_franvaro += data_row["Frånvaro"] or 0
            total_total += data_row["Total"] or 0

            row_num += 1

        # Total rad
        ws.cell(row_num, 1, "Total (ALL)")
        ws.cell(row_num, 1).font = Font(bold=True)
        ws.cell(row_num, 3, seconds_to_hhmm(total_arbetstid))
        ws.cell(row_num, 4, seconds_to_hhmm(total_franvaro))
        ws.cell(row_num, 5, seconds_to_hhmm(total_total))

        for col in range(3, 6):
            ws.cell(row_num, col).font = Font(bold=True)

        row_num += 3  # Mellanrum till nästa månad

    # Justera kolumnbredd
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 12

    # Spara
    wb.save(output_file)
    print(f"Flextidsrapport sparad: {output_file}")


# ============================================================================
# HUVUDPROGRAM
# ============================================================================


def main():
    print("=== Tidrapportering - Python version ===\n")

    # 1. Läs in data
    print("Läser in data...")
    df = load_time_data(INPUT_FILE)
    print(f"Läste {len(df)} rader")

    # 2. Lägg till härledda kolumner
    print("\nLägger till härledda kolumner...")
    df = add_derived_columns(df)

    # 3. Validera att endast en månad finns
    print("\nValiderar data...")
    validate_single_month(df)

    # 4. Skapa alla dagar
    print("\nSkapar alla dagar i månaden...")
    all_days = create_all_days(df)

    # 5. Skapa huvuddataset
    print("\nSkapar huvuddataset...")
    main_df = create_main_query(df, all_days)

    # 6. Lägg till typ av tid
    print("\nLägger till tidskategorier...")
    type_df = add_time_type(main_df)

    # 7. Skapa Antura-rapport
    print("\nSkapar Antura-rapport...")
    create_antura_report(main_df, ANTURA_OUTPUT_FILE)

    # 8. Skapa flextidsrapport
    print("\nSkapar flextidsrapport...")
    create_flex_report(type_df, FLEX_OUTPUT_FILE)

    # 9. Skapa kategoriserad veckerapport
    print("\nSkapar kategoriserad veckerapport (20 veckor)...")
    create_category_report(main_df, CATEGORY_OUTPUT_FILE, weeks_back=20)

    print("\n=== Klart! ===")

    time.sleep(5)


if __name__ == "__main__":
    main()
