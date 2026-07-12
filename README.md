# Newspaper Infographic Web App

Static browser app for newspaper writers who need to turn election-poll Excel files into publishable infographic PNGs.

The app runs entirely in the browser. It reads an `.xlsx` workbook, extracts party names and mandate values, lets the user edit the chart content and typography, previews the result on a canvas, and exports a transparent PNG.

## Project Structure

- `index.html` - Hebrew RTL app shell and form controls.
- `styles.css` - responsive layout and editor styling.
- `app.js` - XLSX parsing, data extraction, chart rendering, and PNG export.
- `average-5.html` - separate page for the last-five-polls grouped horizontal chart.
- `average-5.js` - XLSX extraction and canvas rendering for the last-five-polls chart.
- `סקר הסקרים - 27.4.xlsx` - sample election-poll workbook.
- `EXAMPLE.png` - reference infographic example.
- `election-polls-27.4 (6).png` - exported output example.
- `average_5_part_*.png` - reference images for the grouped horizontal chart.

There is no build step and no package install. The app does not use a backend.

## Running Locally

For normal manual use, open `index.html` in a current Chrome or Edge browser and choose an `.xlsx` file with `בחר XLSX`.

To use the bundled `טען דוגמה` sample button reliably, serve the folder through a local static server because browsers usually block `fetch()` from `file://` pages:

```powershell
python -m http.server 8000
```

Then open:

```text
http://localhost:8000/
```

The last-five-polls chart is available at:

```text
http://localhost:8000/average-5.html
```

## Writer Workflow

1. Load the sample workbook with `טען דוגמה`, or import a new workbook with `בחר XLSX`.
2. Select the relevant sheet.
3. Review and edit the title, subtitle, date, party names, and mandate values.
4. Adjust PNG dimensions, sorting, transparent background, font sizes, headline spacing, bold settings, and party-label layout.
5. Check the canvas preview.
6. Export with `ייצוא PNG`.

Use the header button for `5 הסקרים האחרונים` to open the grouped horizontal chart. It reads the `כל הסקרים` sheet and renders the rightmost five numeric poll columns. Each bar is labeled with its mandate value and publishing media outlet. The average chart page includes basic title/date controls plus advanced controls for workbook loading, layout, font sizes, poll labels, and party values.

The exported filename follows this pattern:

```text
election-polls-{date}.png
```

## XLSX Input Format

The parser is intentionally permissive so newsroom spreadsheets can stay simple.

Recommended workbook layout:

- Row 1: infographic title, for example `נתוני 6 הסקרים האחרונים`.
- Row 2: subtitle, for example `שפורסמו בכלי התקשורת`.
- Row 3: date, for example `27.4`.
- Header row: must contain a party-name column named `מפלגה` or `party`.
- Data rows: party names and mandate numbers.

Mandate values can be provided in either form:

- A single `ממוצע`, `average`, or `avg` column.
- Multiple numeric poll columns. In this case, the app averages the numeric cells per party.

The sample workbook includes both patterns:

- `כל הסקרים` - individual poll columns from several media outlets.
- `ממוצע` - precomputed average column.

When a workbook has a sheet named `ממוצע`, `average`, or `avg`, the app prefers it by default.

Rows whose names look like summaries are skipped, including coalition, opposition, total, and Hebrew equivalents such as `קואליציה`, `אופוזיציה`, and `סה"כ`.

## Implementation Notes

- XLSX files are read client-side in `app.js` by parsing the ZIP container and internal XML files directly.
- Deflated XLSX entries use the browser `DecompressionStream` API.
- Chart output is rendered with the HTML canvas API.
- The last-five-polls chart keeps a fixed 550px canvas width, calculates canvas height from the party count, and exports with a transparent background.
- Hebrew labels are split by words and repositioned to reduce overlap, with an optional Excel-style diagonal mode for dense party names.
- Diagonal party labels can be tuned by angle, alignment, anchor spacing, horizontal and vertical offsets, stagger, and reserved label area.
- Diagonal party labels can optionally split each word onto its own rotated line to reduce the reserved label height for long names.
- Each party row has an individual tag-distance value for separating one diagonal label from its neighbors without changing the full layout.
- Mandate numbers are rounded for display in the infographic.
- The transparent-background toggle affects the exported canvas itself; the checkerboard pattern is only a preview background.
- The canvas uses `FbPractica` for titles and `FbPracticaNarrow` for graph text. If those fonts are not installed, add licensed files under `fonts/` using the names documented in `fonts/README.md`.
- The app checks canvas font availability at runtime and shows a warning if it falls back to another font.

## Browser Requirements

Use a current Chromium-based browser such as Chrome or Edge.

Required browser APIs include:

- `File`
- `Blob`
- `DOMParser`
- `TextDecoder`
- `DecompressionStream`
- `HTMLCanvasElement.toBlob`

## Manual Test Checklist

Before handing off changes, verify:

- `index.html` opens in Chrome or Edge.
- `טען דוגמה` loads the sample workbook when served over `http://localhost`.
- `בחר XLSX` imports `סקר הסקרים - 27.4.xlsx`.
- Switching between `כל הסקרים` and `ממוצע` updates the chart data.
- Editing title, subtitle, date, party rows, and mandate values updates the preview.
- Sorting, transparent background, font size, bold, and party-label layout controls update the preview.
- PNG export creates a usable image comparable to the included PNG examples.
