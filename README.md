# District Performance Reports — PM SHRI DBT Verification

A local dashboard that turns a district export into a ranked, official PDF report:

**Upload Excel → confirm data and district names → select columns → calculate → rank and add the stakeholder message → preview → save PDF.**

## Privacy note for the hosted version

The web page only contains the tool. When someone opens it and uploads an Excel file, the file is read inside their own browser and never sent to GitHub or anywhere else. Do not commit real district data, district lists or generated reports to this repository (`.gitignore` blocks Excel, CSV, PDF and JSON files).

## Run it

1. Open `index.html` (or the GitHub Pages link for this repository) in **Google Chrome or Microsoft Edge** (double-click it, or drag it into a browser window).
2. That's all. There is nothing to install, no server and no internet connection needed. SheetJS and the Mukta font are bundled inside the file.

Your data never leaves the computer: the file is read by the browser itself and nothing is uploaded.

## Saving the PDF

Press **Save as PDF** in step 6. The browser's print window opens; choose **Save as PDF** as the destination (the file name is filled in from the report title and period).

This uses the browser's own print engine on purpose. It is the only option that shapes Hindi conjuncts (क्ष, त्र, श्चि) correctly. Python's ReportLab and the common JavaScript PDF libraries break them. In Chrome and Edge the footer shows "Page X of Y", the generation date and your footer text. Firefox prints the report correctly but without those footer items.

Table headings repeat on every page. Wide tables switch to landscape automatically, and the font scales down as columns are added.

## What the dashboard found in the reference workbook

- **The headings are three rows, not two.** Row 1 is the group (Enrollment Status, Gaurdian Adhar), row 2 the sub-group (Given By ICDS), and row 3 the leaf (Total, Verify, Pending; 3-4 Years). Merged cells are expanded and turned into names like `Given By ICDS - Total` and `Balvatika 1 (3-4 Years)`.
- **Row 4 is the state total.** Every one of its 13 values equals the sum of rows 5–79 exactly, so it is excluded from ranking and can be shown as a "State total" line.
- **There are no district names.** Rows 5–79 (75 rows, which matches the number of districts in UP) contain only numbers. Ranking stays locked until you paste the district list **in the same row order as the portal export**, and you confirm that order again for every new file. Do not assume alphabetical order.
- **Guardian Aadhaar does not add up.** Pending + Not Verify + Verify differs from Total Enrollment in 62 of 75 districts (and in the state total). Use the "Consistency checks" section of step 2 before picking a denominator for Aadhaar percentages.
- **Small denominators distort percentage ranks.** One district has 16 enrolled children; 4 verifications give it 25% and rank 1. Use "Only rank districts where … is at least …" in step 5 for percentage rankings.

## The six steps

1. **Upload** `.xlsx`, `.xls` or `.csv`. Choose the worksheet if there are several. Use "Replace file" for each new reporting period.
2. **Confirm data.** See the detected heading rows, total row and column types, with warnings for blanks, invalid numbers, duplicate or missing names. Rename any column; renames are remembered by heading. Identify districts by a column in the file, by pasting a list in row order (or loading it from a file), or by matching a district code column. A searchable preview of every row is at the bottom.
3. **Select columns.** Tick, search, select all or clear, reorder by drag or arrows, and set the printed heading.
4. **Calculations.** Choose from percentage, ratio, division with a multiplier, difference, percentage-point difference, percentage change, product, sum, average, share of the all-district total, or a custom formula such as `([A] + [B]) ÷ [C] × 100`. You choose every numerator and denominator. Each calculation shows its formula and a preview before you apply it. Division by zero and missing inputs give a blank (—), never a number, and results above 100% are flagged.
5. **Rank and message.** Pick any original or calculated indicator and whether **higher or lower is better** (nothing is assumed). Options include best-first or weakest-first order, a tie-break indicator, three tie rules, top/bottom N, value range, minimum base, and include/exclude districts. Write the stakeholder message in Hindi or English with subject, period, signatory, designation, department and date; place it on the first page, a cover page, or just before the table; save and reuse message templates. A district-specific message supports `{district} {rank} {total_districts} {indicator} {value} {state_value} {period}`.
6. **Preview and download.** There are four formats:
   - **Ranked summary:** rank, district, ranking indicator(s) and calculated indicators.
   - **Detailed report:** every selected column.
   - **Top and bottom performers:** two short tables.
   - **One page per district:** each page carries the district-specific message.
   
   You can also set the title, programme, organisation, logo and footer, and download the ranked table as `.xlsx` or `.csv` for checking.

## Message library (your sample messages)

Step 5 keeps a library of messages you have sent before, so you don't retype them for each report.

- **Add samples:** click "Upload samples" and choose one or more `.txt` or `.docx` files (one message per file). You can also paste a message in. A first line starting with `विषय:` or `Subject:` becomes the subject.
- **Exact text:** "Use" puts the sample into the report word for word.
- **Fill in figures:** write fields in curly brackets anywhere in a sample, for example `{worst_district} is last with {worst_value}`. They are replaced with this report's figures, and recalculated automatically when you upload the next period's file. Click a field button to insert it. Available fields: `{period} {date_today} {indicator} {state_value} {average} {total_districts} {best_district} {best_value} {worst_district} {worst_value} {top_list} {bottom_list} {below_avg_count} {below_avg_list}`.
- **Write a draft from the data:** keeps the first and last paragraphs of the current message (your salutation and your instruction), and writes the middle from the ranking in Hindi, English or both: state figure, average, best and weakest districts, and those below average. You can then edit it freely.
- District pages can also use `{district} {rank} {value}`.

The draft is assembled from the ranking by fixed sentence patterns. It is not written by an AI, so it works offline, and the same data always gives the same wording.

## Saved setups (for recurring reviews)

- Your current work is saved automatically in this browser.
- **Save setup / Save as…** stores named setups: columns, renames, calculations, ranking, report layout, message and district list.
- When you upload a new file, the setup is matched to the new file **by heading**. If any heading changed, you map it to a new column or drop it before anything runs; a formula is never silently applied to a guessed column. The district list is kept, but you must confirm the row order for the new file.
- **Export / Import** saves a setup (with message templates) as a `.json` file. Use it as a backup, because clearing browser data deletes saved setups, and to share a setup with a colleague.

## Source and rebuilding

```
src/template.html   layout and styles
src/core.js         header parsing, district mapping, formula engine, ranking
src/ui.js           screens, events, report builder, saved setups
build.py            inlines SheetJS 0.18.5 and Mukta into index.html
tests/e2e_test.py   end-to-end test in headless Chromium (40 checks)
tests/test_messages.py  message library tests (17 checks; run make_samples.py first)
```

To rebuild: `npm pack xlsx@0.18.5 @fontsource/mukta@5.0.17` into `../deps`, unpack both, then run `python build.py`.

To test: `pip install playwright openpyxl && python -m playwright install chromium`, then `python tests/make_fixtures.py` and `python tests/e2e_test.py` (needs `pdftotext` from poppler).

The test uses placeholder district names (`Test District 01`, `परीक्षण जनपद 02` …). They are not real district assignments.

Third-party: SheetJS Community Edition (Apache 2.0); Mukta by Ek Type (SIL Open Font License).
