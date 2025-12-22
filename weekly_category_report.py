"""
Veckokategoriserad tidrapport - Modul
Skapar en HTML-rapport med stapeldiagram och tabeller för de senaste 20 veckorna
Kategoriserar tidsposter i tre grupper: Övrigt, Python/Soreg, och Övrig tid
"""

import polars as pl
from datetime import datetime, timedelta
import plotly.graph_objects as go


def categorize_work_items(df: pl.DataFrame) -> pl.DataFrame:
    """
    Kategoriserar Work Items i tre kategorier:
    1. 'Övrig tid' - börjar med "Övrig tid"
    2. 'Python/Soreg' - innehåller "Python" eller "soreg" (case-insensitive)
    3. 'Team Sol' - alla andra poster
    """
    df = df.with_columns([
        pl.when(pl.col("Work Item").str.starts_with("Övrig tid"))
        .then(pl.lit("Övrig tid"))
        .when(
            pl.col("Work Item").str.to_lowercase().str.contains("python") |
            pl.col("Work Item").str.to_lowercase().str.contains("soreg")
        )
        .then(pl.lit("Python/Soreg"))
        .otherwise(pl.lit("Team Sol"))
        .alias("Kategori")
    ])

    return df


def get_week_data(df: pl.DataFrame, weeks_back: int = 20) -> list:
    """
    Hämtar data för de senaste X veckorna
    Returnerar en lista med tupler (år, vecka, week_df)
    """
    # Beräkna datum för X veckor sedan
    today = datetime.now().date()
    start_date = today - timedelta(weeks=weeks_back)

    # Filtrera data
    df_filtered = df.filter(pl.col("startdatum") >= start_date)

    # Lägg till kategori
    df_filtered = categorize_work_items(df_filtered)

    # Hämta unika veckor sorterade i omvänd ordning (nyaste först)
    weeks = (
        df_filtered
        .select(["ar", "Vecka"])
        .unique()
        .sort(["ar", "Vecka"], descending=[True, True])
    )

    week_data = []
    for row in weeks.iter_rows(named=True):
        ar = row["ar"]
        vecka = row["Vecka"]

        # Filtrera data för denna vecka
        week_df = df_filtered.filter(
            (pl.col("ar") == ar) & (pl.col("Vecka") == vecka)
        )

        week_data.append((ar, vecka, week_df))

    return week_data


def create_bar_chart(category_hours: dict, year: int, week: int) -> str:
    """
    Skapar ett interaktivt stapeldiagram för en vecka med Plotly
    Returnerar HTML-sträng med Plotly-diagram
    """
    categories = ["Team Sol", "Python/Soreg", "Övrig tid"]
    hours = [category_hours.get(cat, 0) for cat in categories]
    colors = ['#3498db', '#2ecc71', '#e74c3c']  # Blå, Grön, Röd

    # Skapa stapeldiagram med Plotly
    fig = go.Figure(data=[
        go.Bar(
            x=categories,
            y=hours,
            marker=dict(
                color=colors,
                line=dict(color='#000000', width=1.5)
            ),
            text=[f'{h:.1f}h' for h in hours],
            textposition='outside',
            textfont=dict(size=14, color='#2c3e50', family='Segoe UI'),
            hovertemplate='<b>%{x}</b><br>Timmar: %{y:.2f}h<extra></extra>'
        )
    ])

    # Uppdatera layout
    fig.update_layout(
        title=dict(
            text=f'Tidskategorisering - Vecka {week}, {year}',
            font=dict(size=18, color='#2c3e50', family='Segoe UI'),
            x=0.5,
            xanchor='center'
        ),
        yaxis=dict(
            title=dict(
                text='Timmar',
                font=dict(size=14, color='#2c3e50')
            ),
            gridcolor='rgba(0,0,0,0.1)',
            zeroline=True,
            zerolinecolor='rgba(0,0,0,0.2)'
        ),
        xaxis=dict(
            title=dict(
                font=dict(size=14, color='#2c3e50')
            )
        ),
        plot_bgcolor='white',
        paper_bgcolor='white',
        height=450,
        margin=dict(t=80, b=60, l=60, r=40),
        hovermode='x unified',
        showlegend=False
    )

    # Returnera HTML-div med interaktivt diagram
    return fig.to_html(include_plotlyjs='cdn', div_id=f"chart_week_{week}_{year}", config={'displayModeBar': True, 'displaylogo': False})


def create_detail_table(week_df: pl.DataFrame) -> str:
    """
    Skapar en HTML-tabell med detaljer per kategori
    """
    # Gruppera efter kategori och Work Item
    summary = (
        week_df
        .group_by(["Kategori", "Work Item"])
        .agg([
            pl.col("DurationTimmar").sum().alias("Timmar"),
            pl.col("startdatum").count().alias("Antal_poster")
        ])
        .sort(["Kategori", "Timmar"], descending=[False, True])
    )

    html = '<table class="detail-table">'
    html += '<thead><tr><th>Kategori</th><th>Work Item</th><th>Timmar</th><th>Antal poster</th></tr></thead>'
    html += '<tbody>'

    current_category = None
    for row in summary.iter_rows(named=True):
        kategori = row["Kategori"]
        work_item = row["Work Item"]
        timmar = row["Timmar"]
        antal = row["Antal_poster"]

        # Färgkod för kategori
        if kategori != current_category:
            current_category = kategori
            if kategori == "Team Sol":
                row_class = "cat-ovrigt"
            elif kategori == "Python/Soreg":
                row_class = "cat-python"
            else:
                row_class = "cat-ovrig-tid"

        html += f'<tr class="{row_class}">'
        html += f'<td><strong>{kategori}</strong></td>'
        html += f'<td>{work_item if work_item else "(Ingen specifikation)"}</td>'
        html += f'<td>{timmar:.2f}</td>'
        html += f'<td>{antal}</td>'
        html += '</tr>'

    html += '</tbody></table>'
    return html


def create_summary_table(week_df: pl.DataFrame) -> dict:
    """
    Skapar en sammanfattning per kategori
    Returnerar en dict med kategori -> timmar
    """
    summary = (
        week_df
        .group_by("Kategori")
        .agg(pl.col("DurationTimmar").sum().alias("Timmar"))
    )

    result = {}
    for row in summary.iter_rows(named=True):
        result[row["Kategori"]] = row["Timmar"]

    return result


def generate_html_report(week_data: list, output_file: str = "Tidrapport_Kategoriserad.html"):
    """
    Genererar en komplett HTML-rapport med diagram och tabeller
    """
    html = """
<!DOCTYPE html>
<html lang="sv">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Veckokategoriserad Tidrapport</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f5f5f5;
        }
        h1 {
            text-align: center;
            color: #2c3e50;
            margin-bottom: 10px;
        }
        .subtitle {
            text-align: center;
            color: #7f8c8d;
            margin-bottom: 30px;
            font-size: 14px;
        }
        .week-section {
            background-color: white;
            margin-bottom: 40px;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .week-title {
            font-size: 24px;
            color: #2c3e50;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 3px solid #3498db;
        }
        .chart-container {
            margin-bottom: 30px;
        }
        .detail-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        .detail-table thead {
            background-color: #34495e;
            color: white;
        }
        .detail-table th {
            padding: 12px;
            text-align: left;
            font-weight: bold;
        }
        .detail-table td {
            padding: 10px;
            border-bottom: 1px solid #ecf0f1;
        }
        .detail-table tr:hover {
            background-color: #f8f9fa;
        }
        .cat-ovrigt {
            background-color: #ebf5fb;
        }
        .cat-python {
            background-color: #eafaf1;
        }
        .cat-ovrig-tid {
            background-color: #fadbd8;
        }
        .summary {
            display: flex;
            justify-content: space-around;
            margin-bottom: 20px;
            padding: 15px;
            background-color: #ecf0f1;
            border-radius: 5px;
        }
        .summary-item {
            text-align: center;
        }
        .summary-item .label {
            font-size: 12px;
            color: #7f8c8d;
            text-transform: uppercase;
            margin-bottom: 5px;
        }
        .summary-item .value {
            font-size: 24px;
            font-weight: bold;
        }
        .summary-item.ovrigt .value { color: #3498db; }
        .summary-item.python .value { color: #2ecc71; }
        .summary-item.ovrig-tid .value { color: #e74c3c; }
    </style>
</head>
<body>
    <h1>Veckokategoriserad Tidrapport</h1>
    <p class="subtitle">Genererad: """ + datetime.now().strftime("%Y-%m-%d %H:%M") + """</p>
"""

    # Lägg till varje vecka
    for ar, vecka, week_df in week_data:
        # Beräkna sammanfattning
        category_hours = create_summary_table(week_df)

        # Skapa diagram (Plotly returnerar färdig HTML)
        chart_html = create_bar_chart(category_hours, ar, vecka)

        # Skapa detaljerad tabell
        detail_table = create_detail_table(week_df)

        # Beräkna total tid för veckan
        total_hours = sum(category_hours.values())

        html += f"""
    <div class="week-section">
        <h2 class="week-title">Vecka {vecka}, {ar}</h2>

        <div class="summary">
            <div class="summary-item ovrigt">
                <div class="label">Övrigt</div>
                <div class="value">{category_hours.get("Team Sol", 0):.1f}h</div>
            </div>
            <div class="summary-item python">
                <div class="label">Python/Soreg</div>
                <div class="value">{category_hours.get("Python/Soreg", 0):.1f}h</div>
            </div>
            <div class="summary-item ovrig-tid">
                <div class="label">Övrig tid</div>
                <div class="value">{category_hours.get("Övrig tid", 0):.1f}h</div>
            </div>
        </div>

        <div class="chart-container">
            {chart_html}
        </div>

        <h3>Detaljerad uppdelning</h3>
        {detail_table}

        <p style="text-align: right; margin-top: 15px; color: #7f8c8d; font-weight: bold;">
            Total tid denna vecka: {total_hours:.2f} timmar
        </p>
    </div>
"""

    html += """
</body>
</html>
"""

    # Skriv till fil
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"Kategoriserad veckerapport sparad: {output_file}")


def create_category_report(df: pl.DataFrame, output_file: str = "Tidrapport_Kategoriserad.html", weeks_back: int = 20):
    """
    Huvudfunktion för att skapa kategoriserad veckerapport

    Args:
        df: Polars DataFrame med tidrapportdata (måste ha kolumnerna Work Item, DurationTimmar, startdatum, Vecka, ar)
        output_file: Namn på HTML-filen som ska skapas
        weeks_back: Antal veckor bakåt att visa (default 20)
    """
    print(f"\n=== Skapar kategoriserad veckerapport ({weeks_back} veckor) ===")

    # Hämta data för de senaste veckorna
    week_data = get_week_data(df, weeks_back)

    print(f"Hittade {len(week_data)} veckor med data")

    # Generera HTML-rapport
    generate_html_report(week_data, output_file)

    print("=== Rapport klar! ===\n")
