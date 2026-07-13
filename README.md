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


## V7.2 date normalization fix

- Ignores hidden Excel years when cells are displayed as MM/DD.
- Infers year from Report Start and the left-to-right stage sequence.
- `12/28` followed by `01/27` becomes 2026 then 2027.
- Final dates in May remain in May 2027 instead of being shifted to 2028.
- Timeline ends two weeks after the latest milestone.


## V7.3 — MM/DD/YY standard

- Excel date cells should contain a real date and display as `MM/DD/YY`.
- Text dates are accepted only as `MM/DD/YY` or `MM/DD/YYYY`.
- The web no longer guesses years.
- Dates such as `12/28/26`, `01/27/27`, and `05/24/27` are read exactly.
- The timeline ends two weeks after the latest milestone.


## V8 — Milestone interval Gantt

- Original and Actual are milestone timelines.
- Each stage color fills from its milestone date to the next stage milestone.
- The final SER stage fills for one week only.
- Actual uses Actual date when available; otherwise Estimate / Forecast.
- Forecast delay is carried only into stages without Actual dates.
- The timeline stops one week after the last milestone.
- Excel dates must be real dates displayed as `MM/DD/YY`.
