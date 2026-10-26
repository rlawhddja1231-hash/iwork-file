# iWork files — field log

Newest entries on top. Tag each line [tested] / [confirmed: date] / [guess].

## 2026-10-27 v0.1.0 — first release tests [tested]

**Sample set**: 130 public files — the numbers-parser test set (103 Numbers files, including 8 Mac folder packages, encrypted, damaged and pre-2015 files), Apache Tika test documents (Pages, Keynote and Numbers from iWork '09 to 2018) and LibreOffice/libetonyek test files (Keynote 1 to 6, Pages 4 and 5).

**Full run** — for every file: `info`, `read`, `images`, `preview`, and every conversion for its type. 966 runs, 896 succeeded. The 70 that stopped were files with nothing to give (no preview inside, no pictures) or the 10 files built to fail (encrypted without password, broken zip, missing parts, password-protected Pages). All stopped with a plain message; no crash.

**Pages/Keynote replace** — every 2013+ sample edited with `e → éé` and `t → TT` (changes text length everywhere):
- every text storage and table cell came back exactly as expected; 0 style/paragraph markers out of range;
- parts not touched stayed byte-for-byte identical;
- LibreOffice (an independent iWork reader) opened the edited Pages and Keynote files and showed the new text, including Korean characters;
- a Mac folder-package Pages file was edited and saved as a normal single file.
- keynote-parser's encoder was checked first: 174 inner files decoded → re-encoded → decoded with identical content.

**Numbers → Excel formulas** — 7,109 formulas translated, then recalculated by LibreOffice:
- 6,952 it could calculate: 6,386 identical (92%). Differences were edge cases in test files: text-to-number coercion inside arrays, date strings like "15/05/2001", wildcard lookups, 1904-based times, DOLLAR() currency symbol.
- 157 used XLOOKUP etc., which Excel has but this LibreOffice lacks (#NAME?).
- Numbers-only functions (DUR2*, PLAINTEXT, DURATION, header-name references) are written as values with the formula in a cell note.

**Fixed while testing**
1. numbers-parser saved numbers through float maths: 27 → 26.999999999999996, -500000 → -499999.99999999994. Replaced its decimal128 packing with exact Decimal maths; 13 tricky values (0.1, 1/3, 2^52, 7.5e20 …) now come back exactly.
2. Keynote 6 (2013) keeps slides under a hidden root node — slide order now follows it.
3. 2013–2014 Numbers files ("pre-BNC" storage) that numbers-parser refuses are read through our own table decoder.
4. Tables inside Pages/Keynote came out as a jumbled word list → now decoded cell by cell.
5. Non-English style names (e.g. Czech "Odrážka" = bullet): bullets/numbers now come from the list's label type, headings from the built-in English style id.
6. Old iWork '09 Numbers rebuilt as grids (dates, formula results, pop-up menu choice).
7. 16:9 Keynote → PowerPoint: template boxes stretched to the wide page.
8. Keynote 2018 export with the document inside an extra folder and an inner Index.zip: reading and writing both fixed.
9. US conventions: cell input accepts 10/27/2026, $1,200.50, 12%; ZIP codes like 02134 stay text; Excel dates shown mm/dd/yyyy; CSV read as UTF-8 then Windows-1252.

**Still to confirm on a real Mac** [guess]: edited .numbers / .pages / .key files open in Apple's apps. LibreOffice shows edited Numbers files empty, but it does the same for an untouched file simply re-saved by numbers-parser, so that is a LibreOffice limit.
