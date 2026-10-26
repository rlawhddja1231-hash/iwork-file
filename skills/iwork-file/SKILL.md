---
name: iwork-file
description: "Apple iWork files without a Mac: read Pages (.pages), Numbers (.numbers) and Keynote (.key) documents; convert them to Word (.docx), Excel (.xlsx with working formulas, .csv), PowerPoint (.pptx) or Markdown; edit Numbers cells; find-and-replace text in Numbers, Pages and Keynote; build a new .numbers file from Excel/CSV; save the pictures and slide thumbnails inside. Use this whenever the user uploads or mentions a .pages, .numbers or .key file (or a folder ending in .pages/.numbers/.key), or says things like \"I got a Pages file and can't open it\", \"open this Numbers spreadsheet\", \"convert Keynote to PowerPoint\", \"Pages to Word\", \"Numbers to Excel\", \"Mac file on Windows\", \"iWork\", \"change the name in this Keynote\", \"make it a Numbers file\". Use it even when the user does not name the skill. Not for .docx/.xlsx/.pptx files that are already Office files."
---

# iWork files (Pages · Numbers · Keynote)

Apple's iWork files only open on a Mac, iPhone or iPad. This skill lets Claude open and change them anywhere. The tool is `scripts/iwork.py`. It is built on numbers-parser (MIT) and keynote-parser (MIT), plus its own decoders for Pages text, for tables inside Pages and Keynote, and for 2013–2014 Numbers files.

Evidence tags: **[tested]** run on 130 public sample files (2026-10-27) / **[confirmed]** checked in the real Mac app or with a real user file (add the date) / **[guess]** not checked yet.

---

## 0. How to talk

- Answer in the user's language, in plain words. Don't show command names, flags or paths unless asked: say "I opened your Pages file", not "I ran iwork.py read".
- Use the user's conventions. Default to US style for English speakers: dates as 10/27/2026, `$1,200.50`, Letter paper. The tool accepts `10/27/2026`, `2026-10-27`, `$1,200`, `1,200.50` and `12%` as cell values and keeps ZIP codes like `02134` as text.
- **Before changing content (set / replace), show what will change and ask.** For replace, run `--dry` first and report how many places match, with a few examples.
- Never overwrite the original. Save under a new name and send the whole file (SendUserFile).
- When something can't be done, say so in one sentence and offer the nearest thing that can.

---

## 1. What works and what doesn't

**Works** [tested]
- **Read all three types** — 2013+ files and old iWork '09 files, single files or Mac folder packages. Pages: paragraphs, headings, bullet/numbered lists, tables as real rows and columns, text boxes; headers/footers and footnotes (2013+ files). Keynote: every slide in order, skipped slides marked, speaker notes, bullet levels, tables. Numbers: every sheet and table, dates, pop-up menu choices, formulas (`read --formulas` shows them).
- **Convert**
  - Numbers → .xlsx: one sheet per table, merged cells, bold, fill colors, column widths, US date format. **Formulas stay live** when Excel has the same functions (SUM, IF, VLOOKUP, XLOOKUP, cross-table references …). Numbers-only formulas (DUR2HOURS, PLAINTEXT, header-name references …) become their value with the original formula in a cell note. `--values-only` writes plain values. Also .csv and .md.
  - Pages → .docx (headings, lists, tables), .md, .txt. Keynote → .pptx (one slide each with title, bullets, speaker notes, same 4:3/16:9 shape), .docx, .md, .txt.
  - Old iWork '09 files that carry a full PDF inside → .pdf.
- **Edit Numbers**: set cells (`B3=1200`, text, numbers, dates, `true/false`), find-and-replace text in every table. After saving, the tool reopens the file and checks that **every other cell is unchanged**.
- **Edit Pages and Keynote (2013+)**: find-and-replace text in the body, text boxes, headers/footers, table cells and speaker notes. Formatting is kept: every style, paragraph and list marker after the change is moved to match. Parts of the file that weren't touched stay byte-for-byte the same. Master-slide (theme) text in Keynote is left alone.
- **New Numbers file** from .xlsx (every sheet) or .csv.
- **Password-protected Numbers**: open, read, convert and edit with `--password`; the stored hint is shown if the password is missing or wrong.
- **Pictures**: `images` saves the photos and media the document actually uses (`--all` adds theme pictures). `preview` saves the cover and every Keynote slide thumbnail — Read them to *see* the slides.

**Doesn't work (yet)**
- Adding or removing line breaks with replace; restyling text; adding slides, pages or pictures.
- Writing formulas into Numbers files (values only). Excel formulas become their last calculated value.
- Pictures, charts, shapes and exact layout are not copied into Word/PowerPoint — text, lists and tables are.
- Password-protected Pages/Keynote files; editing Pages/Keynote saved before 2013 or Numbers saved before 2015 (reading them works).
- The small preview picture inside an edited file still shows the old text until the file is saved again on a Mac.
- Edited files have been checked with our own reader and with LibreOffice (an independent iWork reader), but not yet in Apple's apps [guess: they open]. LibreOffice shows edited **Numbers** files empty — it does the same for any numbers-parser save, so this is a LibreOffice limit.

---

## 2. Steps

1. **Set up** (once per chat)
   ```
   pip install -q --break-system-packages numbers-parser keynote-parser openpyxl python-docx python-pptx
   W=$(find / -name iwork.py -path '*iwork-file*' 2>/dev/null | head -1); echo $W
   ```
2. **Get the file**: an upload or a file from a connected folder. Work on a copy.
3. **Look first**
   ```
   python3 $W info FILE                 # type, age, sheets/slides, password, what can be edited
   python3 $W read FILE --out FILE.md
   ```
   Read the Markdown yourself and tell the user in a line or two what the file is.
4. **Do the job**

   | User says | Command |
   |---|---|
   | "What's in this file / summarize it" | `read`, then summarize |
   | "Show me the slides" | `preview FILE --out pics`, then Read the pictures |
   | "Get the photos out" | `images FILE --out pics` |
   | "Convert to Word / Excel / PowerPoint / PDF" | `convert FILE out.docx` · `out.xlsx` · `out.pptx` · `out.pdf` (old files only) |
   | "Change B3 to 1,200" | `set FILE.numbers B3=1200 [--sheet "Sheet 1"] [--table "Table 1"] --out FILE_edited.numbers` |
   | "Replace ACME with Globex everywhere" (any of the three) | `replace FILE ACME Globex --dry` → ask → `replace FILE ACME Globex --out FILE_edited.ext` |
   | "Make this spreadsheet a Numbers file" | `new data.xlsx --out data.numbers` |
   | "It asks for a password" | add `--password PW` (Numbers only) |

   Several replacements at once: `replace FILE old1 new1 old2 new2`. `set` without `--sheet/--table` uses the first table of the first sheet; use the names `info` prints.
5. **Check before handing over**: `read` the result and compare with the request. For .docx/.pptx/.xlsx you can render or open them with the docx/pptx/xlsx skills. If the tool printed `WARNING` or `skipped`, tell the user.
6. **Hand over**: one or two lines on what changed, then the file. For edited iWork files add: "Please open it on your Mac and let me know if anything looks off."
7. After real use, add what you learned to `references/log.md`.

---

## 3. Common messages

| Message | Meaning | What to do |
|---|---|---|
| "password-protected — run again with --password" | locked Numbers file | ask for the password (the hint may be shown) |
| "password-protected or uses packing this tool cannot open" | locked Pages/Keynote | ask them to remove the password on the Mac (File › Set Password) or export to Word/PowerPoint/PDF |
| "damaged" / "not a valid zip package" | broken or not an iWork file | ask for the file again; if it came by email, ask them to zip it first |
| "Older Numbers files (before 2015) can be read and converted but not edited" | 2013–2014 Numbers | convert to .xlsx, edit, then `new` back to .numbers |
| "editing works on Pages/Keynote files saved in 2013 or later" | iWork '09 file | convert to .docx/.pptx instead |
| "line breaks can't be added or removed" | replace text contains a new line | replace one line at a time |
| "skipped … formatting changes inside the word" | part of the word is styled differently | tell the user that spot was left as is |
| "part(s) of this file are damaged and were skipped" | some inner parts unreadable | give what was read and say what's missing |
| "no text — picture or shapes only" on a slide | slide is pictures | use `preview` and look at the thumbnail |
| "writing formulas … is not supported" | `=...` given to set | write the number, or convert to Excel |

---

## 4. Field log

→ `references/log.md`. Read it before working; add to it after.
