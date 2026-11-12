# Doc to Markdown

A browser-based converter between Word, Excel or PDF documents and Markdown. It
was designed for dense technical reports, such as Environmental Impact
Assessment (EIA) reports, whose Markdown output is intended for AI tools.

- **No installation and no cost.** The app is a static web page, hosted free on
  GitHub Pages. It can also be installed as a Progressive Web App (PWA) or used
  as a single downloaded HTML file.
- **Local processing.** Documents are read and converted inside the browser and
  are never uploaded. A Content Security Policy technically prevents the page
  from sending data to any other server.
- **Offline use.** After the first visit, the installed app works with no
  internet connection.

## Functions

| Direction | Input | Output |
| --- | --- | --- |
| Document → Markdown | `.pdf`, `.docx`, `.xlsx` (also `.docm`, `.xlsm`) | `.md` plus a verification report |
| Markdown → Document | `.md` | Word `.docx`, Excel `.xlsx`, PDF, standalone `.html` |

### Conversion to Markdown

- **Text structure.** Headings (including automatic numbering such as "6.1.1.5"),
  paragraphs, bulleted and numbered lists (nested), bold and italic text,
  hyperlinks, footnotes, sub- and superscripts.
- **Tables with merged cells.** Row and column spans are preserved. Two output
  styles are available:
  - *HTML tables* (recommended): a `<table>` with `rowspan`/`colspan` is
    embedded in the Markdown. The structure is exact, and both AI tools and
    Markdown viewers read it correctly. Tables without merged cells remain
    ordinary Markdown tables.
  - *Pure Markdown*: merged values are repeated in every cell they span.
- **Word.** The document's internal XML is read directly. Merged cells, list
  numbering, tracked changes, text boxes, footnotes and endnotes, comments,
  headers and footers, the table of contents and cover pages are all
  recognised. Single-column layout boxes are output as ordinary text.
- **Excel.** Every visible sheet becomes a section. Cells are shown as Excel
  displays them (number formats, percentages and dates as `yyyy-mm-dd`), and
  merged cells are kept. Blocks separated by empty rows become separate tables.
  Hidden sheets, rows and columns are optional.
- **PDF (digitally generated).** Text is extracted rather than recognised.
  Tables are rebuilt from the drawn border lines, which recovers merged cells.
  Tagged PDFs, such as Word's "Save as PDF", also provide headings, lists,
  figures and header/footer marks. Tables without borders are rebuilt from those
  tags. Tables that continue across pages are joined, with the repeated header
  row removed. Optional page markers (`<!-- Page 5 -->`) support traceability.
  Scanned (image-only) PDFs are not supported.

### Removal before conversion

Before each conversion, the app asks which items to remove: images and charts,
headers and footers, page numbers, the table of contents (and lists of tables
and figures), footnotes, comments and the cover page. It also asks whether to
accept tracked changes. The choices are remembered in the browser. Every removed
item is listed in the verification report.

### Verification report

The app checks each conversion against an **independent inventory** of the
source file and produces a report with an overall result: *Complete*, *Review
recommended* or *Incomplete*.

1. **Text inventory check.** Every word in the source is looked for in the
   output. The result is a coverage percentage plus a list of any passages not
   found, with their location (section, page or cell).
2. **Numerical data check.** Every number (concentrations, limits, dates,
   coordinates and so on) is matched individually.
3. **Structural check.** Headings, list items, tables and sheets are counted in
   the source and the output and compared. For Word and Excel, every table is
   also compared **cell by cell**, including merged-cell positions.
4. **Deliberate exclusions.** Items removed on purpose are listed separately, so
   that they are not mistaken for losses.
5. **Duplication check.** Words that occur more often in the output than in the
   source are flagged.

**Items to review.** The top of each report lists only the items that need
attention, with colour coding: 🔴 missing or mismatched, 🟠 check
recommended (for example, a table rebuilt without border lines), 🟢 OK. The
same symbols are used throughout the downloaded report, so they remain visible
in any text editor or AI tool. Inside the app, the checklist provides:
- a **Show** button that opens the side-by-side review at the item;
- a **Reviewed** tick box (ticks are saved in the browser only);
- a filter that hides items already reviewed.

**Side-by-side review (PDF).** Each original page is shown next to its
converted output.
- Flagged tables are outlined in amber on the page and in the output.
- Text missing from the output is outlined in red on the page.
- The item selected from the checklist is outlined in blue.
- **Flagged page** buttons jump between pages that contain flagged items.

**Limitation.** The text check proves that every value is *present* in the
output. For PDFs, it cannot prove that every value sits in the *correct table
cell*, because a PDF stores no table structure. Critical PDF tables should
therefore be inspected visually in the side-by-side review. For Word and Excel,
cell positions are compared exactly.

## Use

### In the browser
Open the GitHub Pages address (for example `https://avl-io.github.io/doc-to-md/`).
Then:
1. Drag files onto the drop zone.
2. Click **Convert to Markdown**.
3. Confirm the options.

The two directions are shown as two cards at the top. Each has its own colour theme (navy for Document → Markdown, slate for Markdown → Document), with matching dark-mode versions, a three-step guide and its own drop zone. Files are routed automatically: a Markdown file dropped on Document → Markdown goes to Markdown → Document (and the reverse), and the card shows how many files are waiting. Messages appear as in-app notices rather than browser pop-ups. The `.md` files are downloaded automatically. Each tab (Document → Markdown,
Markdown → Document) keeps its own results list. **Clear results** empties that
list; files already downloaded are not affected.

### Installing as an app (PWA)
1. Open the address in Microsoft Edge or Google Chrome.
2. Click the **Install** icon in the address bar, or **Install app** in the page
   header.

The app then appears in the Start menu, opens in its own window and works
offline. Installation is per user and needs no administrator rights. Updates
are offered automatically when you are online. Some organisations disable PWA
installation; in that case, use the address directly or the offline file.

### Single-file offline version
Use the link at the bottom of the page to download `doc-to-md-offline.html`
(about 6 MB). It can be opened directly from disk, without hosting or internet
access. It does not update automatically.

## Development

Requirements: Node.js 20 or later.

```bash
npm install
```

```bash
npm run build
```

`npm run build` writes the app to `dist/`, including `dist/doc-to-md-offline.html`.

```bash
npm run serve
```

This serves `dist/` at http://localhost:8080/.

```bash
npm test
```

`npm test` builds the app, then converts every test file in a headless browser
with the real security policy. It writes the results to `test-output/`.

- `tests/fixtures/` contains synthetic test documents created by
  `tests/make-fixtures.py`. They hold no real data, so they can be committed.
  Expected results are listed in `tests/expectations.json`.
- `tests/private/` is excluded from git. Real documents placed there are
  converted by the test run but are never committed.
- Playwright is used only for tests. It is loaded from the global npm folder if
  it is not installed in the project.

### Source layout

| Path | Purpose |
| --- | --- |
| `src/app.js`, `src/index.html`, `src/styles.css` | User interface |
| `src/sw.js`, `src/manifest.webmanifest`, `src/icons/` | PWA (offline cache, install) |
| `src/lib/docx-reader.js`, `xlsx-reader.js`, `pdf-reader.js` | Readers that build the shared document model (`ast.js`) |
| `src/lib/md-writer.js` | Document model → Markdown |
| `src/lib/md-reader.js`, `docx-writer.js`, `xlsx-writer.js`, `pdf-writer.js`, `html-writer.js` | Markdown → Word / Excel / PDF / HTML |
| `src/lib/verify.js` | Independent inventory and verification report |
| `build.mjs` | Build of the PWA and the single-file version |

### Deployment
`.github/workflows/pages.yml` builds the app and publishes it to GitHub Pages.
It runs on every push to `main`, and it can also be started manually from the
**Actions** tab. One-time setting: **Settings → Pages → Build and deployment →
Source: GitHub Actions**.

## Third-party components

All of the following are bundled into the app. None is loaded from the
internet at run time.

| Component | Use | Licence |
| --- | --- | --- |
| [pdf.js](https://github.com/mozilla/pdf.js) (`pdfjs-dist`) | Reading PDF files | Apache-2.0 |
| [JSZip](https://github.com/Stuk/jszip) | Reading and writing Office packages | MIT |
| [docx](https://github.com/dolanmiu/docx) | Writing Word files | MIT |
| [marked](https://github.com/markedjs/marked) | Reading Markdown | MIT |
| [pdfmake](https://github.com/bpampuch/pdfmake) (with Roboto font) | Writing PDF files | MIT (font: Apache-2.0) |
