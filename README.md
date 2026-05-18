# Tidrapportering – lathund och beskrivning

## Snabbguide

```
1. Kör rapportera.bat
2. Logga in i Antura om det behövs (webbläsaren öppnas automatiskt)
3. Kontrollera och skicka manuellt i Antura
```

Det är allt. Inloggning behövs normalt bara första gången – sessionen sparas sedan.

---

## Vad som händer i varje steg

`rapportera.bat` kör tre steg i ordning och väntar på en knapptryckning mellan varje.

### Steg 1 – Generera Excel-underlagen

Kör `Hantera_tidrapportering.py`, som läser din tidrapportering från
Grindstone (`.gsjdb`-databasen) och OneDrive-filen `All_my_time.tsv`.

Skriptet skapar eller uppdaterar tre filer i mappen `data\`:

- **`Tidrapport_Antura.xlsx`** – tiddata uppdelad per vecka och aktivitet,
  formaterad för att kunna matchas mot Antura. Det är den här filen som
  `antura_filler.py` läser från.
- **`Tidrapport_Flex.xlsx`** – flextidsrapport med tider i HH:MM-format.
- **`Tidrapport_Kategoriserad.html`** – kategoriserad sammanfattning.

Kör det här direkt efter att du avslutat veckan i Grindstone, innan du
rapporterar i Antura.

### Steg 2 – Starta Chrome med fjärrfelsökning

Startar en separat Chrome-instans med fjärrfelsökning aktiverat på port 9222.
Chrome öppnas direkt mot Antura-sidan och använder en dedikerad profil
sparad i mappen `chrome_profile\` i projektmappen.

Första gången: logga in med ditt UCR-konto. Sessionen sparas i profilen
och återanvänds vid nästa körning – du behöver alltså normalt inte logga in igen.

Lämna Chrome-fönstret öppet och inloggat medan steg 3 körs.

### Steg 3 – Fyll i Antura automatiskt

Skriptet ansluter till Chrome-instansen via CDP (Chrome DevTools Protocol)
och läser av Antura-sidan utan att du behöver lämna ut ditt lösenord till skriptet.

Det här händer i tur och ordning:

1. **Läser Excel** – hämtar alla veckor med data från `data\Tidrapport_Antura.xlsx`.

2. **Hittar startpunkten** – läser texten "Senast skickade dag: X Mmm, YYYY"
   på Antura-sidan för att veta var ifyllningen ska börja.

3. **Navigerar vecka för vecka** – väljer rätt vecka i Anturas dropdown och
   bearbetar sedan alla veckor framåt som har data i Excel.

4. **Matchar aktiviteter** – jämför varje rad i Excel mot raderna i Antura
   med fuzzy matching (ungefärlig textlikhet). Tre strategier prövas i tur och ordning:
   hela strängen, bara aktivitetsdelen (efter `–`) och bara projektdelen (före `–`).
   Den matchning som ger högst poäng används.

   | Poäng | Åtgärd |
   |---|---|
   | ≥ 80 % | Fylls i, markeras grönt i Excel |
   | 55–79 % | Fylls i, markeras **gult** i Excel – kontrollera! |
   | < 55 % | Fylls **inte** i, markeras **rött** i Excel |

5. **Fyller i formuläret** – skriver in värdena i varje redigerbar cell.
   Låsta celler (redan skickade dagar) hoppas alltid över.
   Celler med värdet 0 lämnas orörda.

6. **Sparar markeringar** – `data\Tidrapport_Antura.xlsx` sparas med färgmarkeringar
   och kommentarer på de rader som var osäkra eller inte matchade.
   Granska gula och röda rader innan du skickar.

### Kontrollera och skicka manuellt

Skriptet trycker aldrig på "Skicka vecka" eller "Skicka del av vecka".
Det är avsiktligt – du granskar resultatet och skickar in själv.

Titta särskilt på:
- **Gula rader i Excel** – skriptet matchade, men med lägre säkerhet.
  Kontrollera att rätt Antura-rad fick rätt värde.
- **Röda rader i Excel** – ingen matchning hittades. Fyll i de raderna manuellt i Antura.

---

## Förutsättningar (första gången)

```
poetry install
poetry run playwright install chromium
```

Python 3.13+ och Poetry måste vara installerade sedan tidigare.

---

## Projektstruktur

| Fil / mapp | Beskrivning |
|---|---|
| `rapportera.bat` | Huvudskript – kör hela flödet i ett |
| `Hantera_tidrapportering.py` | Läser Grindstone/TSV och genererar Excel-filerna |
| `antura_filler.py` | Fyller i Antura automatiskt via Playwright |
| `grindstone_reader.py` | Hjälpmodul för att läsa Grindstone-databasen |
| `weekly_category_report.py` | Hjälpmodul för kategorirapporten |
| `run.bat` | Kör `Hantera_tidrapportering.py` (används av `rapportera.bat`) |
| `run_antura.bat` | Kör `antura_filler.py` (används av `rapportera.bat`) |
| `start_chrome.bat` | Startar Chrome med CDP-stöd (används av `rapportera.bat`) |
| `data\` | Genererade rapportfiler (gitignorerade) |
| `chrome_profile\` | Sparad Chrome-session (skapas automatiskt, gitignorerad) |
| `arkiv\` | Inaktuella filer som inte längre används aktivt |
