# Instruktioner för tidrapportering i Antura

## Översikt
Hjälp Fredrik fylla i tidrapport i Antura från Excel-filen `Tidrapport_Antura.xlsx`.
Skicka ALDRIG veckan – det gör Fredrik själv.

## Steg-för-steg

1. **Läs Excel-filen** – flik som matchar aktuell vecka (t.ex. "Vecka 17")
2. **Öppna Antura i Edge** med Claude-tillägget inloggat
   - URL: `https://ucrantura.ucr.uu.se/Views/MyPlace/TimeReporting/Reporting.aspx`
3. **Kontrollera att rätt vecka är vald** i Antura
4. **Ändra "Visa"-dropdown till "Alla"** – annars syns inte alla projekt (t.ex. ThoR).
   ⚠️ OBS: Detta triggar en sidomladdning! Vänta tills sidan laddats klart INNAN du fyller i värden.
5. **Fyll i timmar via JavaScript** med `setVal(baseId, dagIndex, värde)`
6. **Verifiera** att alla värden stämmer med en kontrollkörning

## Aktivitetsmappning: Excel → Antura

| Excel-namn | Antura-projekt | Antura-aktivitet |
|---|---|---|
| soreg - Plattformslyft | SOReg registerlyft | 6. Utdata |
| Swedevox | SWEDEVOX - Oxygenbehandling och respiratorbehandling i hemmet | 9.4. Rapportutveckling |
| Thor | ThoR - Allmän Thoraxkirurgi | 9.4. Rapportutveckling |
| Underhåll Uppsala-register - 10. SPAHR - Uppgradering | Underhåll Uppsala-register | 10. SPAHR - Uppgradering |
| Övergripande Kvalitetsregister - 6.11. Agil process - inom team | Övergripande kvalitetsregister | 6.11. Agil process – inom team |
| Övrig tid - 1.1 Avdelningstid | Övrig, icke projektrelaterad arbetstid | 1.1. Övrig tid |
| Övrig tid - 1.4. Kompetensutveckling | Övrig, icke projektrelaterad arbetstid | 1.4. Kompetensutveckling |
| Övrig tid - Friskvårdstimme | Övrig, icke projektrelaterad arbetstid | 1.6. Friskvårdstimme |

## Dagindex i JavaScript
- 0 = Måndag
- 1 = Tisdag
- 2 = Onsdag
- 3 = Torsdag
- 4 = Fredag
- 5 = Lördag
- 6 = Söndag

## Timvärden
Runda till 2 decimaler. Lämna tomt (eller hoppa över) om värdet är 0.

## JavaScript-funktion för att sätta värden
OBS: Varje fält är en DIV med klassen `reportingField`. Input-elementet finns INUTI div:en.
Sätt alltid värdet på input:en inuti div:en, inte på div:en själv.

```javascript
function setVal(baseId, dayIdx, val) {
  const div = document.getElementById(baseId + '_FieldsBlock_' + dayIdx);
  if (!div) return 'DIV NOT FOUND: ' + baseId + '_FieldsBlock_' + dayIdx;
  const input = div.querySelector('input[type="text"]');
  if (!input) return 'INPUT NOT FOUND inside ' + baseId + '_FieldsBlock_' + dayIdx;
  input.value = val.toString();
  input.dispatchEvent(new Event('input', {bubbles: true}));
  input.dispatchEvent(new Event('change', {bubbles: true}));
  input.dispatchEvent(new Event('blur', {bubbles: true}));
  return 'OK';
}
```

## OBS
- Aktivitets-ID:n (baseId) kan ändras om Antura uppdateras – hämta dem via JavaScript vid behov:
```javascript
const rows = document.querySelectorAll('table tr');
let currentProject = '';
const all = [];
rows.forEach((row) => {
  const cells = row.querySelectorAll('td');
  const inputs = row.querySelectorAll('input[type="text"]');
  const projectLink = cells[0]?.querySelector('a[href*="ProjectPlace"]');
  if (projectLink) currentProject = cells[0]?.innerText?.trim().replace(/\s+/g, ' ');
  if (inputs.length > 0) {
    let el = inputs[0];
    while (el && !el.id?.match(/^r\d+/)) el = el.parentElement;
    const baseId = el?.id?.replace(/_FieldsBlock_\d+$/, '') || '?';
    all.push(`${currentProject} | ${cells[0]?.innerText?.trim()} | ${baseId}`);
  }
});
all.join('\n')
```
