# Duplicate Tooling Tracking – V7

Static GitHub Pages application.

## Files

- `index.html`
- `style.css`
- `app.js`

## Deploy

Copy these three files into the root of the GitHub repository and replace the old files.

Delete old `css/` and `js/` folders if they exist, because V7 only uses the three files above.

Then in GitHub Desktop:

1. Commit to `main`
2. Push origin
3. Wait for GitHub Pages deployment
4. Open the website and press `Ctrl + F5`

## Gantt logic

- Timeline begins from the selected Report Start date. Default: `2026-06-01` (WK23).
- Original row uses Target dates as the baseline.
- Actual row uses Actual dates when available.
- When Actual is blank, Estimate is used.
- Delay from an earlier stage is carried forward to later stages without Actual dates.
- Bars are rectangular and fill the entire row.
- Timeline includes Year, Week number, and week start date.


## V7.1 fix

- Correctly infers the year for MM/DD values.
- Example: `12/28` followed by `01/27` becomes Dec 2026 followed by Jan 2027.
- Prevents the forecast from being incorrectly pushed into 2028.
