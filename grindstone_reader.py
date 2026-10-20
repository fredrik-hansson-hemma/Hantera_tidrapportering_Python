"""
Grindstone Database Reader
Läser tidsposter från en Grindstone .gsjdb-fil (JSON-format)
och returnerar en Polars DataFrame.
"""

import json
import shutil
import tempfile
from datetime import datetime, timezone
from dateutil.relativedelta import relativedelta
from pathlib import Path

import polars as pl


def load_grindstone_data(filepath: str) -> pl.DataFrame:
    """
    Läser tidsposter från en Grindstone .gsjdb-databasfil.

    Filen är JSON med strukturen:
      root['f']['r'] = records (tidsposter)
      root['f']['t'] = tasks (uppgifter med namn)

    Returnerar en DataFrame med kolumner:
      Work Item (str), Start (Datetime), End (Datetime),
      Duration (Int64, sekunder), Notes (str)
    """
    # Kopiera till temporär fil för att undvika OneDrive/Grindstone-låsning
    temp_path = _copy_to_temp(filepath)

    try:
        data = _read_json(temp_path)
    finally:
        Path(temp_path).unlink(missing_ok=True)

    # Bygg task-ID → task-namn lookup
    tasks = {task["i"]: task["n"] for task in data["f"]["t"]}

    # Beräkna cutoff-datum: 11 månader tillbaka från den 1:a i nuvarande månad,
    # dvs. 12 kalendermånader inklusive innevarande
    cutoff = (datetime.now(timezone.utc).replace(day=1) - relativedelta(months=11)).replace(
        hour=0, minute=0, second=0, microsecond=0
    )

    # Extrahera poster
    records = data["f"]["r"]

    work_items = []
    starts = []
    ends = []
    durations = []
    notes = []

    for rec in records:
        # Parsa tider
        start_utc = datetime.fromisoformat(rec["s"])
        end_utc = datetime.fromisoformat(rec["e"])

        # Filtrera bort gamla poster
        if start_utc < cutoff:
            continue

        # Task-namn via lookup
        task_id = rec.get("t")
        work_items.append(tasks.get(task_id, ""))

        # Konvertera UTC till naiv lokal tid
        start_local = _utc_to_local_naive(start_utc)
        end_local = _utc_to_local_naive(end_utc)

        starts.append(start_local)
        ends.append(end_local)

        # Beräkna duration i sekunder
        durations.append(int((end_utc - start_utc).total_seconds()))

        # Anteckningar (valfritt fält)
        notes.append(rec.get("n", ""))

    df = pl.DataFrame(
        {
            "Work Item": work_items,
            "Start": starts,
            "End": ends,
            "Duration": durations,
            "Notes": notes,
        }
    )

    df = df.with_columns(
        [
            pl.col("Start").cast(pl.Datetime),
            pl.col("End").cast(pl.Datetime),
            pl.col("Duration").cast(pl.Int64),
        ]
    )

    return df


def _copy_to_temp(filepath: str) -> str:
    """Kopierar databasfilen till en temporär plats för att undvika låsning."""
    temp_dir = tempfile.gettempdir()
    temp_path = Path(temp_dir) / "grindstone_temp.gsjdb"
    shutil.copy2(filepath, temp_path)
    return str(temp_path)


def _read_json(filepath: str) -> dict:
    """Läser och parsar JSON-databasfilen."""
    with open(filepath, "r", encoding="utf-8") as f:
        return json.load(f)


def _utc_to_local_naive(dt: datetime) -> datetime:
    """
    Konverterar en UTC-medveten datetime till en naiv lokal datetime.
    Resten av flödet räknar med naiva lokala tidsstämplar.
    """
    if dt.tzinfo is not None:
        local_dt = dt.astimezone()
    else:
        local_dt = dt.replace(tzinfo=timezone.utc).astimezone()

    return local_dt.replace(tzinfo=None)