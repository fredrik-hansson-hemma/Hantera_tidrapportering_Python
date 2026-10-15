# Tidrapportering – lathund och beskrivning

## Snabbguide

```
1. Kör rapportera.bat
2. Logga in i Antura om det behövs (Chrome öppnas automatiskt)
3. Kontrollera och skicka manuellt i Antura, tryck Enter för nästa vecka
```

Det är allt. Inloggning behövs normalt bara första gången – sessionen sparas sedan.

---

## Vad som händer i varje steg

`rapportera.bat` kör `rapportera.py`, som gör allt i samma fönster. Du behöver
bara växla till Chrome för att logga in (om det behövs) och för att granska och
skicka in varje vecka.

### Steg 1 – Generera Excel-underlagen

Kör `Hantera_tidrapportering.py`, som läser din tidrapportering från
Grindstone (`.gsjdb`-databasen). Sökvägen anges i `config.py`.

Skriptet skapar eller uppdaterar tre filer i mappen `data\`:

- **`Tidrapport_Antura.xlsx`** – tiddata uppdelad per vecka och aktivitet,
  formaterad för att kunna matchas mot Antura. Det är den här filen som
  `antura_filler.py` läser från.
- **`Tidrapport_Flex.xlsx`** – flextidsrapport med tider i HH:MM-format.
- **`Tidrapport_Kategoriserad.html`** – kategoriserad sammanfattning.

Kör det här direkt efter att du avslutat veckan i Grindstone, innan du
rapporterar i Antura.

### Steg 2 – Starta Chrome och logga in

Om ingen Chrome med fjärrfelsökning redan kör startas en separat Chrome-instans
i bakgrunden (port 9222, se `config.py`). Chrome öppnas direkt mot Antura och
använder en dedikerad profil sparad i mappen `chrome_profile\` i projektmappen.
Chrome-fönstret ligger kvar när skriptet är klart.

Om du inte är inloggad visas "👉 Logga in på Antura i Chrome-fönstret".
Logga in med ditt UCR-konto – skriptet fortsätter av sig självt när
tidrapporteringssidan visas. Sessionen sparas i profilen, så normalt behöver
du bara logga in första gången.

### Steg 3 – Fyll i Antura automatiskt

Skriptet ansluter till Chrome-instansen via CDP (Chrome DevTools Protocol)
och läser av Antura-sidan utan att du behöver lämna ut ditt lösenord till skriptet.

Det här händer i tur och ordning:

1. **Läser Excel** – hämtar alla veckor med data från `data\Tidrapport_Antura.xlsx`.

2. **Hittar startpunkten** – läser texten "Senast skickade dag: X Mmm, YYYY"
   på Antura-sidan för att veta var ifyllningen ska börja.

3. **Navigerar vecka för vecka** – väljer rätt vecka i Anturas dropdown och
   bearbetar sedan alla veckor framåt som har data i Excel. Skriptet väntar
   själv tills sidan har laddat klart innan det fyller i.

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

6. **Listar veckans kommentarer** – kommentarer som loggats i Grindstone skrivs
   ut med dag, work item och vilken Antura-rad de hör till. Lägg in dem manuellt
   i Antura.

7. **Sparar markeringar** – `data\Tidrapport_Antura.xlsx` sparas med färgmarkeringar
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
| `rapportera.bat` | Huvudskript – dubbelklicka för att köra hela flödet |
| `rapportera.py` | Kör hela flödet: Excel-underlag, Chrome och Antura-ifyllnad |
| `Hantera_tidrapportering.py` | Läser Grindstone och genererar Excel-filerna |
| `antura_filler.py` | Fyller i Antura automatiskt via Playwright |
| `grindstone_reader.py` | Hjälpmodul för att läsa Grindstone-databasen |
| `weekly_category_report.py` | Hjälpmodul för kategorirapporten |
| `config.py` | Sökvägar och adresser (Grindstone-fil, utdatafiler, Antura-URL) |
| `filhjalp.py` | Gemensam hjälpfunktion för att spara filer som kan vara öppna i Excel |
| `chrome_starter.py` | Startar Chrome med CDP-stöd om den inte redan kör |
| `run.bat` | Kör bara `Hantera_tidrapportering.py` |
| `run_antura.bat` | Kör bara `antura_filler.py` |
| `start_chrome.bat` | Används inte längre – ersatt av `chrome_starter.py` |
| `data\` | Genererade rapportfiler (gitignorerade) |
| `chrome_profile\` | Sparad Chrome-session (skapas automatiskt, gitignorerad) |
| `arkiv\` | Inaktuella filer som inte längre används aktivt |
