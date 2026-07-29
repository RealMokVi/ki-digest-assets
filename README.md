# ki-digest-assets

Öffentlich gehostetes Anhang-Verzeichnis für die Claude Code Routine `ki-digest`
(KI-Update Digest per Mail). Existiert ausschließlich, damit generierte Dateien
per `raw.githubusercontent.com`-URL als Mail-Anhang über den Resend-Connector
versendet werden können — Resend ruft Anhänge über die `url`-Option selbst ab,
was den teuren Umweg über Base64-Inhalte im Modell-Kontext vermeidet (~150.000
Tokens für eine ~100 KB-Datei).

Seit 2026-07-29 fester Bestandteil der Routine (Schritt 5a in `routine-prompt.md`, Repo
`RealMokVi/claude` unter `llm-wiki/newsletter/` — privat, daher kein Link von hier aus),
nicht mehr nur ein einmaliger Test.

## Inhalt

- `scripts/` — die zwei Node-Generatoren (`gen_deck.js` für PowerPoint via `pptxgenjs`,
  `gen_pdf.js` für PDF via `pdfkit`) plus Datenschema-Doku (`scripts/README.md`). Die Routine
  ruft diese bei jedem Lauf auf, statt den Code jeden Tag neu zu schreiben.
- `pptx/` — generierte PowerPoint-Decks, ein Deck pro versendeter Mail (Dateiname
  `KI-Update-{Datum}.pptx`, bei Korrektur-/Test-Läufen mit `-Korrektur`/`-Test`-Suffix).
- `pdf/` — dieselben Digests als PDF, gleiches Namensschema.

Kein Bestandteil der eigentlichen Spezifikation selbst (die steht im Haupt-Repo unter
`llm-wiki/newsletter/`) und kein privates Datenrepo (siehe `ki-digest-log` für log.jsonl,
sources.md, model-releases.md, errors.log). Nur öffentlich unbedenkliche Inhalte hier
ablegen — alles hier ist für jeden mit dem Link einsehbar.
