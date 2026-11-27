# Areta Comparables Automation: banker guide (1 page)

**Where:** open the [Comparables Automation dashboard](https://alycz.github.io/Areta-Comparables-Automation/), or the Sheet *Areta Trading Comps* and then **Comparables Automation › Open dashboard**.

## Every week: nothing to do
Prices refresh automatically on Monday morning, London time. The model recalculates, and last week's close becomes the prior price. Need fresh numbers before a client call? Click **Refresh prices**.

## When a company reports
1. In Capital IQ, export the company's **Financials** in the usual format (millions, US$).
2. Drop the file(s) on the dashboard (step 01), then click **Read & validate**.
3. Open the **Review queue** tab. Each row shows the current value, the new value, the % change and the exact CIQ cell it came from.
   - 🟢 **Ready**: checks passed.
   - 🟡 **Review**: will be written, but needs your judgement (see Flags).
   - 🔴 **Blocked**: will **not** be written; the current value stays.
4. Untick anything you disagree with, then click **Apply changes**.

Companies you didn't upload keep their current numbers.

## Before you send anything
- Open the **Exceptions** tab. Deal with every 🔴 and 🟡 item, then tick it and click **Mark reviewed**, adding a comment, e.g. "Use listed shares".
- Click **Check & build report** to get a one-page Validation Report tab you can attach or forward.

## Things the engine will never do
- Overwrite a good price with a failed one. (A suspicious price is written, and flagged for you to review.)
- Write a market cap or EV bridge that does not reconcile to CIQ's own figures (where CIQ shows one).
- Overwrite your own notes in the Notes column. Engine notes start with ⚑.
- Make a judgement call for you, e.g. FX convention, implied EBITDA, a press-release period, or broker debt.

## Changing things (no code)
Use the **Config** tab: schedule, alert thresholds, who can use the web app, and Google Finance ticker overrides. To add a company, add a row on Comps with its CIQ ticker and run **Comparables Automation › Tools › Set up / reinstall schedule**.
