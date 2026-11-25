# Walkthrough script (30 minutes)

| Min | Show | Say |
|---|---|---|
| 0–3 | Web app home | "The problem: prices by hand, CIQ copy-paste, manual checks. The solution keeps the Sheet as the model and adds an engine plus a dashboard. Google sign-in is the gate; it runs as whoever is signed in, so access = Sheet sharing." |
| 3–8 | Step 2: drop the 20 CIQ files → Read & validate | "Each file is converted, matched by ticker and read **by label**, not by cell. Layouts differ: BlackRock says 'Current Capitalization'; Robinhood and Tradeweb use CIQ's financial template with no EV bridge and no EBITDA." |
| 8–13 | Review queue | "Every proposed value shows its source cell. Before anything is written, market cap and EV are rebuilt and reconciled to CIQ's own figures; 18 of 18 reconcile within 0.5%. Blocked values never reach the model." |
| 13–16 | Apply → Comps tab in the Sheet | "Only blue cells change. Structure and formatting are untouched. Formulas are engine-installed, NM when the denominator ≤ 0, medians ignore NM." |
| 16–19 | Refresh prices | "GOOGLEFINANCE, cross-checked against Yahoo. Prior price is looked up from saved history, so a second run in the same week is safe. A failed quote never overwrites a good one." |
| 19–24 | Exceptions | Walk the judgement items: BlackRock press-release LTM (Sep vs Jun peers); Coincheck JPY translated at spot; Robinhood broker funding in debt (like the excluded brokers) and implied EBITDA; Tradeweb implied vs listed shares; PENN negative NCI. "The engine flags; the banker decides and signs off, with a comment that persists." |
| 24–26 | Overwrite a formula in the Sheet → Run checks → Repair formulas | "Controls catch hardcodes and recalc mismatches independently." |
| 26–28 | Audit Trail, Run Log, Validation Report | "Every value has a trail back to the file, tab and cell." |
| 28–30 | Config tab + limitations | "Bankers change the schedule, thresholds, tickers and access without code. Limits: Google Finance delay, unofficial cross-check, implied EBITDA is a proxy. Next step: AI-drafted exception commentary, kept outside the numbers." |
