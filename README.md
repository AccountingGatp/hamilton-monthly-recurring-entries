# Recurring Entries Review

A browser-based monthly review for a recurring entries schedule, a GL account list and the accrual-basis ResMan Journal Entry Register. The Python review engine runs in WebAssembly in the browser. Vercel serves static files. Workbook contents and result ZIPs stay in the browser.

## Deploy with GitHub and Vercel

1. Push the files in this directory to the root of a GitHub repository. Keep accounting exports and result ZIPs out of the repository.
2. In Vercel, open **Add New → Project** and import the GitHub repository.
3. Set the project root to the directory containing `package.json`. The preset is **Other**, the build command is `npm run build`, and the output directory is `dist`. The included configuration sets these values. No environment variables are needed.
4. Deploy and use the deployment URL Vercel returns. Subsequent pushes to the configured production branch trigger a new deployment.

For a local build, use Node.js 24, run `npm ci` and `npm run build`, and serve the `dist` directory over HTTP (for example, `python -m http.server --directory dist 8000`).

## Run a review

Choose the three Excel files and any review month. Choose **Automatic by month** or select a specific schedule tab. For an accrual-basis register, export all properties and include the review month and the preceding month. Drafts require a full review month in the register, an export printed after month-end, and supporting prior journal activity. The download contains a summary, the schedule review, journal activity and a draft workpaper if supported.

The GL account list supplies account names and codes. Trailing P&L and mortgage statement validation require separate source documents.

## Dependencies

The build bundles Pyodide from the pinned npm package and the pinned openpyxl and et_xmlfile wheels in `vendor/`. The wheels include their license notices. Pyodide is distributed under MPL-2.0 and its corresponding source is available from [Pyodide v0.29.5](https://github.com/pyodide/pyodide/tree/0.29.5). Browser runtime assets are served from the same deployment.
