# Recurring Entries Review

A browser-based monthly review for a recurring entries schedule, a GL account list and the accrual-basis ResMan Journal Entry Register. The Python review engine runs in WebAssembly in the browser. Vercel serves static files. Workbook contents and result ZIPs stay in the browser.

## Deploy with GitHub and Vercel

1. Push the files in this directory to the root of a GitHub repository. Keep accounting exports and result ZIPs out of the repository.
2. In Vercel, open **Add New → Project** and import the GitHub repository.
3. Set the project root to the directory containing `package.json`. The preset is **Other**, the build command is `npm run build`, and the output directory is `dist`. The included configuration sets these values. No environment variables are needed.
4. Deploy and use the deployment URL Vercel returns. Subsequent pushes to the configured production branch trigger a new deployment.

For a local build, use Node.js 24, run `npm ci` and `npm run build`, and serve the `dist` directory over HTTP (for example, `python -m http.server --directory dist 8000`).

## Run a review

Choose the three Excel files and set an inclusive start and end date. The app reviews each calendar month in the range and recognizes dated schedule tabs or dated rows in a shared tab. Select a tab manually to resolve an ambiguous single month or use a shared schedule tab. A missing schedule is reported for that month. If enabled, the prior-month option uses the immediately preceding schedule when current data is missing. A manual tab plus this option in a single-month run selects the prior schedule explicitly.

Export the accrual-basis register for all properties, covering the selected dates and the preceding calendar month. Current journal activity is filtered to the selected dates. The prior comparison uses the preceding calendar month. Drafts require a complete calendar month within the selected range, a register printed after month-end, and matching prior journal support. The ZIP includes review rows, journal activity, monthly coverage, and draft lines where supported.

The GL account list supplies account names and codes. Trailing P&L and mortgage statement validation require separate source documents.

## Dependencies

The build bundles Pyodide from the pinned npm package and the pinned openpyxl and et_xmlfile wheels in `vendor/`. The wheels include their license notices. Pyodide is distributed under MPL-2.0 and its corresponding source is available from [Pyodide v0.29.5](https://github.com/pyodide/pyodide/tree/0.29.5). Browser runtime assets are served from the same deployment.
