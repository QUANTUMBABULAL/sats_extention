# SATS Chrome Extension

This is a Manifest V3 Chrome extension that reads Excel data and fills dropdowns in a table on the active webpage.

## What It Does

- Loads an `.xlsx` file from the popup
- Parses the first sheet into JSON rows
- Sends the row data to `content.js`
- Fills dropdowns inside the table with id `example`
- Updates one row at a time with a short delay so the page stays responsive

## Requirements

- Google Chrome
- Node.js and npm
- A webpage that contains:
	- `#example tbody tr`
	- dropdowns with names like:
		- `registrationCustomDTO.result.firstl_grade`
		- `registrationCustomDTO.result.secondl_grade`
		- `registrationCustomDTO.result.sub0_grade`
		- `registrationCustomDTO.result.sub1_grade`

## Project Files

- `manifest.json` - Chrome extension manifest
- `popup.html` - extension popup UI
- `popup.js` - Excel file loading and data sending logic
- `content.js` - page automation logic

## Setup

1. Open the project folder in VS Code.
2. Install dependencies:

```bash
npm install
```

3. Make sure the `node_modules` folder is present.

## How To Use

1. Open Chrome and go to `chrome://extensions`.
2. Turn on **Developer mode**.
3. Click **Load unpacked**.
4. Select this project folder.
5. Open the target website that contains the table.
6. Click the extension icon.
7. Click **Load Excel** and choose your `.xlsx` file.
8. Click **Start Automation**.

## Excel Format

The first sheet of the Excel file should contain rows like this:

```json
[
	{
		"Lan1": "A",
		"Lan2": "B+",
		"Sub1": "A+",
		"Sub2": "B+"
	}
]
```

The extension converts these values into the format used by the dropdown options, for example:

- `A` -> `1,A`
- `A+` -> `1,A+`
- `B+` -> `1,B+`

## Notes

- The extension processes rows in order.
- It waits about 300 ms between row updates to avoid freezing the UI.
- Missing rows or missing dropdowns are skipped safely.
- If the page uses a different table structure, update the selectors in `content.js`.

## Troubleshooting

- If nothing happens after clicking **Start Automation**, confirm the target page has already loaded the table.
- If the extension cannot read the file, make sure it is an `.xlsx` workbook.
- If dropdown values do not update visually, check whether the page uses custom JavaScript on the select elements.

## License

No license has been specified for this project.
