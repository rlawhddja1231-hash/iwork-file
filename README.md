<p align="center"><img src="assets/icon-256.png" width="120" alt="iWork Files plugin icon"></p>

# iWork Files — open and edit Pages, Numbers & Keynote without a Mac

**A free Claude plugin that reads Apple Pages (.pages), Numbers (.numbers) and Keynote (.key) files, converts them to Word, Excel and PowerPoint, and edits them — on Windows, Linux, or anywhere Claude runs.**

Someone sends you a `.pages` résumé, a `.numbers` budget or a `.key` deck, and your PC can't open it. Claude can't read these files on its own either. Install this plugin and just upload the file.

## What you can ask

| Say this | You get |
|---|---|
| "What's in this Pages file? Summarize it." | The full text with headings, lists, tables, text boxes, headers/footers and footnotes |
| "Convert this Pages file to Word" | A .docx with headings, bullet/numbered lists and tables |
| "Turn this Numbers file into Excel" | An .xlsx — one sheet per table, **working formulas**, merged cells, bold, colors, column widths, US dates |
| "Make this Keynote a PowerPoint" | A .pptx — every slide with its title, bullets and speaker notes, same 4:3 or 16:9 shape |
| "Change the client name in this Keynote / Pages file" | The same file with the text replaced and the formatting kept |
| "Change B3 to $1,200 in this Numbers file" | The same .numbers file with that cell changed — every other cell checked unchanged |
| "Replace ACME with Globex everywhere" | Find-and-replace across Numbers tables, Pages text or Keynote slides and notes (shows matches first) |
| "Show me the slides" / "Get the photos out" | The slide pictures and images saved inside the file |
| "Make this Excel file a Numbers file" | A new .numbers file from .xlsx (all sheets) or .csv |
| "It has a password" | Password-protected Numbers files open with the password (the saved hint is shown) |

Works with iWork files from 2013 to today, old iWork '09 files, and files saved as Mac folder packages.

## Install

In Claude Code:

```
/plugin marketplace add rlawhddja1231-hash/iwork-file
/plugin install iwork-file@iwork-file
```

In the Claude app (claude.ai): zip the `skills/iwork-file` folder and upload it under **Settings → Capabilities → Skills**.

On first use Claude installs the Python packages it needs:

```
pip install numbers-parser keynote-parser openpyxl python-docx python-pptx
```

You can also run the tool yourself:

```
python3 skills/iwork-file/scripts/iwork.py read    Report.pages
python3 skills/iwork-file/scripts/iwork.py convert Budget.numbers Budget.xlsx
python3 skills/iwork-file/scripts/iwork.py convert Deck.key Deck.pptx
python3 skills/iwork-file/scripts/iwork.py replace Deck.key "ACME" "Globex" --out Deck_new.key
python3 skills/iwork-file/scripts/iwork.py set     Budget.numbers B3=1200 --out Budget_new.numbers
```

## How it was tested

- **130 public sample files** from the numbers-parser, Apache Tika and LibreOffice test suites — Pages, Numbers and Keynote from iWork '09 to 2018, including encrypted, damaged and folder-package files. 966 runs (read, info, every conversion, pictures, previews): every file worked or stopped with a clear message. No crashes.
- **Formulas**: 7,109 Numbers formulas converted to live Excel formulas and recalculated by a separate spreadsheet engine (LibreOffice). Of the 6,952 it could calculate, 92% gave exactly the same result as Numbers; the rest were deliberate edge cases in test files (text-to-number tricks, 1904 dates, wildcard lookups). The other 157 use newer functions such as XLOOKUP that Excel has but the checking engine doesn't.
- **Pages/Keynote editing**: every 2013+ sample edited with length-changing replacements; all text came back exactly as expected, no formatting marker out of place, and the untouched parts byte-for-byte identical. LibreOffice, an independent iWork reader, opens the edited files and shows the new text.

Details: [`skills/iwork-file/references/log.md`](skills/iwork-file/references/log.md).

## Limits (honest list)

- Replace changes text only; it can't add line breaks, slides, pages or pictures.
- Numbers formulas can't be written into .numbers files (values only).
- Pictures, charts and exact layout are not copied into Word/PowerPoint.
- Password-protected Pages/Keynote files can't be opened — remove the password on the Mac first.
- Edited files haven't been opened in Apple's own apps by us yet. If one doesn't open, please [open an issue](https://github.com/rlawhddja1231-hash/iwork-file/issues) — that report helps a lot.

## Credits

- [numbers-parser](https://github.com/masaccio/numbers-parser) by Jon Connell — MIT
- [keynote-parser](https://github.com/psobot/keynote-parser) by Peter Sobot — MIT
- [openpyxl](https://openpyxl.readthedocs.io/), [python-docx](https://github.com/python-openxml/python-docx), [python-pptx](https://github.com/scanny/python-pptx) — MIT

Apple, iWork, Pages, Numbers and Keynote are trademarks of Apple Inc. This project is independent and is not affiliated with or endorsed by Apple.

License: MIT · Privacy: [PRIVACY.md](PRIVACY.md)
