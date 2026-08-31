# Audit Correction Studio

A zero-backend web tool that applies audit-findings comments to an Excel workbook and
produces the ready-to-submit corrected version — entirely in the browser.

Built for the *Overview of Transaction Fees and Incentives (ECAG)* review cycle, but
generic: it works for any workbook pair that follows the same pattern (a data sheet +
an audit copy of it that carries comment/response columns or Excel cell comments).

## How it works

1. **Upload the workbook to be corrected** (first worksheet is the one processed).
2. **Upload the audit-findings workbook.** The tool auto-detects:
   - the header row (looks for *Product Group / Products*),
   - the comment columns (headers containing *comment / question / finding*),
   - the response columns (headers containing *response / answer / info*),
   - any extra "late note" columns to the right of the headed table,
   - legacy cell notes and threaded comments, and
   - the row offset between the two files (audit copies often carry an inserted title row).
3. **Review the proposals.** Every commented row is classified by a rules engine:
   | Signal in the response / note | Proposed action |
   |---|---|
   | "no longer exists", "discontinued", "delisted" (late note wins over earlier answers) | delete row |
   | "can be deleted", "will remove", "line is outdated", "to be removed" | delete row |
   | "will change to valid to DD.MM.YYYY" | replace an imprecise *Until …* validity with the exact date |
   | "should be referred to *X* PSS" | fill the relevant incentive cell with `PSS - X` (flagged for review — validity must be verified) |
   | rename hints ("apply PSS's version", "unify", "only *CODE* remained", "*CODE* does not exist anymore") | rewrite the product title (PSS title quoted in the comment is used; ® marks and tickers are carried over; flagged for review) |
   | "No, …", clarifying answers | no change |
   | anything else | left for a human decision |
   Rows **without** any comment are never modified.
4. **Generate.** The corrected `.xlsx` plus a change-log `.csv` are downloaded.

## Fidelity guarantee

The output is produced by *surgical edits to the original file's XML* (JSZip +
DOMParser) — not by re-writing the workbook through a spreadsheet library. Only the
first worksheet's XML is touched. Styles, merged cells, row heights, hyperlinks,
auto-filter, freeze panes, print settings, and **all other worksheets stay
byte-identical**. Row deletions renumber subsequent rows and correctly shrink/shift
merged ranges, hyperlink anchors, the auto-filter range and the sheet dimension.

## Privacy

Everything runs client-side. No server, no upload, no analytics. The page works
offline once loaded (all assets are vendored). Suitable for internal documents.

## Repository layout

```
index.html                        the complete app (self-contained single file:
                                  UI + OOXML engine + vendored JSZip 3.10.1, MIT/GPLv3)
apply_corrections_2026-09-01.py   the audited Python reference implementation used
                                  for the 2026-09-01 correction run (openpyxl),
                                  kept for provenance

```

## Hosting

Any static host works. For GitHub Pages: push this folder as a repository,
then *Settings → Pages → Deploy from branch → main / root*. `.nojekyll` is included.

## Limitations

- `.xlsx` only (not `.xls` / `.xlsb`).
- Free-text audit comments are interpreted by rules; anything ambiguous is flagged
  for a human decision in the review table rather than guessed silently.
- If a proposal's wording needs polish (e.g. a product title after removing a
  discontinued code), edit it directly in the review table before generating.
- Formula cells are renumbered with their rows but formula *references* are not
  rewritten (the ECAG overview sheet contains no formulas).
