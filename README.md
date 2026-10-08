# excel-viewer

A small web page that displays an Excel (`.xlsx`) file as a table. Everything runs in the browser, so the file you choose is never uploaded anywhere.

## Features

- Choose a file or drag and drop it onto the page
- Shows the first sheet as a table
- Sheet tabs appear when the workbook has more than one sheet
- Clear error messages for files that are not valid `.xlsx`

## Usage

No install or build step is needed. Open `index.html` in a browser, then choose an `.xlsx` file.

## Files

| File | Purpose |
| --- | --- |
| `index.html` | Page layout |
| `script.js` | Reads the file and renders the table |
| `style.css` | Styling |

## Built with

- Plain HTML, CSS and JavaScript
- [SheetJS (xlsx)](https://sheetjs.com/) 0.18.5, loaded from cdnjs, for reading Excel files (an internet connection is needed to load it)

## Limitations

- Only `.xlsx` files are supported (not `.xls` or `.csv`)
