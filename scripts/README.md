# Generatoren für PowerPoint- und PDF-Anhänge

Erzeugen aus denselben Digest-Daten sowohl ein PowerPoint-Deck (`gen_deck.js`,
via `pptxgenjs`) als auch ein PDF (`gen_pdf.js`, via `pdfkit`) — ein Deck/PDF
pro versendeter Mail (planmäßig, manuell oder Korrektur).

## Verwendung

```bash
npm install          # einmalig, installiert pptxgenjs + pdfkit
node gen_deck.js data.json    # schreibt data.outFile (.pptx)
node gen_pdf.js data.json     # schreibt denselben Namen mit .pdf-Endung
```

## Daten-Schema (`data.json`)

```jsonc
{
  "outFile": "KI-Update-2026-07-29.pptx",   // .pptx-Name; gen_pdf.js ersetzt die Endung selbst
  "dateLabel": "29. JULI 2026",              // Kopfzeile rechts, Großbuchstaben-Stil
  "subtitle": "Der tägliche KI-Digest ...",  // Kopfzeile, Untertitel
  "counts": { "praxis": 2, "tipps": null, "reg": 1, "forschung": 1 }, // null = "—", leere Kategorie
  "warnbox": "Teilausfall: ...",             // optional, nur bei echter technischer Auffälligkeit
  "briefing": [                              // optional, nur Montag (Quellencheck + Wochenrückblick)
    { "label": "Quellencheck", "text": "..." },
    { "label": "Woche im Rückblick", "text": "..." }
  ],
  "praxis": [ { "title": "...", "summary": "...", "source": "heise.de", "url": "https://..." } ],
  "tipps": [ /* gleiche Form */ ],
  "reg": [ /* gleiche Form */ ],
  "forschung": [ /* gleiche Form */ ],
  "forschungNote": "Für den Zeitraum ... (Teilausfall).", // nur wenn forschung: [] und ein Hinweis nötig ist
  "tippsNote": "Keine Themen im 24h-Fenster gefunden.",   // analog für tipps, wird aber i. d. R. nicht als Slide/Absatz gerendert (siehe unten)
  "omitted": [ { "title": "...", "source": "heise.de", "url": "https://..." } ]
}
```

Wichtig:
- Leere Kategorie (kein Teilausfall, einfach nichts Berichtenswertes): Array leer lassen,
  KEIN `*Note`-Feld setzen — die Kategorie wird dann komplett weggelassen (Slide/Abschnitt),
  genau wie in der Mail selbst. Nur bei einem echten Teilausfall (Warnbox-Fall) das `*Note`-Feld
  füllen, dann erscheint ein grau-kursiver Hinweis statt der Kategorie.
- `gen_pdf.js` ersetzt automatisch Zeichen, die Helvetica (WinAnsiEncoding) nicht darstellen kann
  (`→` → `->`, `⚠` wird entfernt) — Umlaute, Gedankenstriche und deutsche Anführungszeichen („…“)
  sind unproblematisch und werden korrekt dargestellt.
- Bei eigenen Kopien der JSON-Datei (z. B. zum Ändern von `outFile`) **niemals** über Python
  `json.load`/`json.dump` ohne `encoding="utf-8"` gehen — Python öffnet Dateien unter Windows sonst
  mit der System-Codepage, was Umlaute zu Mojibake doppelt-encodiert. Node
  (`fs.readFileSync(path, "utf-8")` / `fs.writeFileSync(path, str, "utf-8")`) ist unkritisch.
