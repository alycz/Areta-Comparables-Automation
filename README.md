# Areta Comparables Automation

Automates Areta's trading comparables model in **Google Sheets**. It does three things:
- refreshes share prices on a schedule;
- loads new Capital IQ exports into the model without copy-paste;
- checks every number, so a banker only looks at the items that need judgement.

The engine is Google Apps Script bound to the trading comparables Sheet. Bankers use it from an Areta-branded dashboard, which can be the **web app** or a **sidebar inside the Sheet**. The Sheet is always the source of truth.

| What bankers wanted | What the engine does |
|---|---|
| Update share prices by hand | Prices refresh weekly on a schedule, or on demand. They come from Google Finance, are cross-checked against Yahoo Finance and saved to Price History. Prior price is the saved close from a week earlier. |
| Pull financials out of Capital IQ | Drop the CIQ "Financials" exports in. The engine finds the company by ticker and extracts shares, net debt, NCI, preferred, LTM revenue, LTM EBITDA and growth by row label. It proposes changes with their source cells, and writes only the blue input cells once a banker approves. |
| Check the calculations | Market cap and EV are rebuilt and reconciled to CIQ on every import. Every formula is recomputed independently. The engine flags hardcodes and #errors, and logs every change with its source. |

On the supplied data, all 20 exports parse and 200 fields are extracted. Market cap and EV reconcile to CIQ within 0.5% for all 18 companies that have a CIQ EV bridge. There are 0 blocked values, 14 items for banker judgement and 28 for information. See [docs/VALIDATION_REPORT.md](docs/VALIDATION_REPORT.md).

---

## How a banker uses it

The one-page version is [docs/BANKER_GUIDE.md](docs/BANKER_GUIDE.md). The Sheet opens on a **Start Here** tab, which shows the live status, these steps and the links. The **Comparables Automation** menu is numbered in run order:

| Menu | What it does | When |
|---|---|---|
| **1. Import CIQ exports** | Reads every export in the CIQ Inbox folder, matches it by ticker, extracts the fields, reconciles them to CIQ and fills the **Review Queue**: **Ready** (passed every check), **Review** (flagged for judgement) or **Blocked** (never written). | When companies report |
| **2. Apply financial changes** | Writes the ticked queue rows into the blue input cells only, and moves the files to Processed. | After a quick look at the queue |
| **3. Refresh share prices** | Pulls Google Finance prices, cross-checks them against Yahoo Finance, saves them to Price History and sets the prior price. | Automatic every Monday 07:00 London; by hand any time |
| **4. Run checks** | Recomputes every formula independently, and flags hardcodes, #errors, missing inputs and outliers on the **Exceptions** tab. | Runs after steps 2 and 3 anyway |
| **5. Build validation report** | Rebuilds the **Validation Report** tab and saves a dated PDF copy to Drive › Reports. | Before sending numbers out |
| **▶ Run all steps** | Runs 1 → 5. Ready and Review rows are applied, and Review items stay open on Exceptions for sign-off. Blocked rows are never written. | One click for the weekly routine |

Sign off judgement items on the **Exceptions** tab: set Status to Reviewed and add a comment. Sign-offs persist between runs. The same actions are available in the branded dashboard (**Open dashboard** in the menu) and in the web app.

Companies without a new export keep their existing figures. Bankers' own notes in the Notes column are never overwritten; engine flags start with ⚑.

**Who can use it:** anyone the Sheet is shared with as an editor, each with their own Google account. Every run and change is logged under that person's email. Scheduled refreshes run under the account that set them up. The `allowedUsers` setting on the Config tab can limit the web app to named people.

## What's in the Sheet

| Tab | Purpose |
|---|---|
| **Comps** | The model. Same structure and formatting as the template. Yellow = weekly inputs (engine), blue = quarterly inputs (engine, after approval), black = formulas (engine-installed). |
| **Exceptions** | Open ERROR / REVIEW / INFO items with source and sign-off status. |
| **Review Queue** | Proposed financial changes awaiting Apply. |
| **Start Here** | Landing page: live status, the steps in order, links, access. |
| **Validation Report** | Summary of updates, exceptions and items for manual review. One tab, rebuilt each run; dated PDFs are kept in Drive › Reports. |
| **Read Me** | Areta's conventions, plus an "Automation" section added by the engine. |
| **Config** | Schedule, thresholds, access list and ticker map. Bankers can edit it; no code needed. |
| **Price History** | Every price snapshot (date, ticker, price, source, cross-check). |
| **Audit Trail** | Every value written: old, new, user, run, and source file / tab / cell. |
| **Run Log** | Every run (scheduled or manual) with status and summary. |

---

## Architecture

```
CIQ exports (.xls) ──► Drive "CIQ Inbox" ──► Import.js ──► Core.parseCiqWorkbook ─┐
                          ▲   (convert to temp Sheet)                              │
            web app upload┘                                                        ▼
GOOGLEFINANCE ─► Prices.js ─► Core.validatePrice ─┐                    Core.validateFinancials
Yahoo Finance ──┘ (cross-check, first-run prior)  │                                │
                                                  ▼                                ▼
                       Comps tab (yellow/blue inputs only) ◄── Review Queue ◄── Apply (banker)
                                  │
                 Checks.js: recompute formulas in JS, hardcodes, #errors, outliers
                                  │
                Exceptions · Audit Trail · Run Log · Validation Report
                                  │
                 App.js: menu, sidebar, web app (ui/Index.html), triggers
```

| File | Responsibility |
|---|---|
| `appsscript/Core.js` | **Pure logic, no Google APIs**: CIQ parser, validation rules, model mapping, calculations. Unit-tested in Node against all 20 exports. |
| `appsscript/Prices.js` | GOOGLEFINANCE feed, Yahoo cross-check, history, prior-price lookup, weekly / daily schedule entry points. |
| `appsscript/Import.js` | Drive inbox, xls → Sheet conversion, staging, Review Queue, Apply, archiving. |
| `appsscript/Checks.js` | Model integrity checks and formula repair. |
| `appsscript/Model.js` | Locates columns by header text and companies by ticker; installs formulas. |
| `appsscript/Log.js` | Audit Trail, Run Log, Exceptions (with persistent sign-off). |
| `appsscript/Report.js` | Validation Report tab. |
| `appsscript/Config.js` | Config tab, settings and ticker map. |
| `appsscript/App.js` | Menu, one-click setup, triggers, web app and its API. |
| `appsscript/ui/` | Areta-branded dashboard (Instrument Sans / Benne, Areta blues, logo). |
| `tests/`, `tools/` | Node test suite, fixture converter, simulator, report generator. |

### Key technical decisions

- **Google Sheets stays the operating environment.** The brief names it, and bankers already work there. The web app is a front end to the Sheet, not a second copy of the data. The model can always be opened, audited and edited directly.
- **The web dashboard has its own password, and the Sheet stays behind Google sharing.** Google's sign-in hop breaks when a browser has several Google accounts signed in, so the dashboard (GitHub Pages) calls the Apps Script API with a password. The password is stored only as a salted hash; login returns a signed token that expires after 12 hours. The script's own web address serves no page, only the API, so nothing can be run without the password. The optional `allowedUsers` setting on the Config tab limits which names or emails can sign in.
- **Parse by label, not cell address.** CIQ layouts differ between companies. BlackRock uses "Current Capitalization", and Robinhood and Tradeweb use the financial-company template with no EV bridge or EBITDA. The parser finds rows by label and columns by period header, so a layout shift doesn't break it.
- **Prove the mapping.** On each import the engine rebuilds market cap (CIQ price × shares) and EV (market cap + net debt + NCI + preferred) and compares them with CIQ's own figures. If they don't reconcile, the value is blocked.
- **Never overwrite with worse data.** A failed price, a missing field, the wrong units or the wrong currency leave the existing value in place, with an ERROR item.
- **Judgement is flagged, not assumed.** Examples: press-release periods, FX translation, implied EBITDA, broker funding inside debt, implied shares, negative NCI and period mismatches. Each becomes a REVIEW item with the numbers behind it.
- **Prior price comes from history.** It is looked up in the saved history, never shifted cell to cell, so re-running in the same week is safe and the weekly move is always computed against a dated close.
- **Pure core, thin adapters.** All business rules live in `Core.js`, which runs unchanged in Apps Script and Node, so they are tested against the real exports.
- **AI.** The product was built with Claude Code (Anthropic), which did the analysis, code, tests and docs. The engine itself makes no AI calls: financial numbers come only from deterministic, reconciled rules. A natural next step is LLM-drafted commentary on the exceptions, kept outside the numbers.

### Data conventions

The figures follow Areta's Read Me:
- **EV** = market cap + net debt + NCI + preferred, as reported by CIQ.
- **Net debt** = total debt − cash & ST investments − LT marketable securities. This is CIQ's TEV bridge.
- **Revenue** = CIQ standardised Total Revenue (the as-reported difference is flagged above 2%). **EBITDA** = CIQ standardised EBITDA.
- **Growth** = CIQ LTM revenue growth over the prior year.
- **Multiples** show **NM** when the denominator is ≤ 0. **Medians** ignore NM.

Where CIQ shows no EBITDA (financial template), the engine uses **implied EBITDA = operating income + D&A**, labels it in the EBITDA Basis column and flags it.

---

## Setup

### Using the deployed copy

- The Sheet is *Areta Trading Comps*, in the Drive folder *Areta Comps — Areta Trading Comps*.
- Open it, then run **Comparables Automation › Set up / reinstall schedule** once and approve the permissions. Setup installs formulas, the Config tab, the log tabs, the Drive inbox and the weekly trigger. It is idempotent.

### Independent test by Areta (a fresh copy)

1. **File › Make a copy** of the Sheet. The script is copied with it.
2. In the copy: **Comparables Automation › Tools › Set up / reinstall schedule**, then approve the permissions. Google shows "unverified app" for any private script: click *Advanced › Go to Areta Comparables Automation*.
3. **Comparables Automation › Tools › Open CIQ inbox folder** and drop in the 20 CIQ exports.
4. **Comparables Automation › Open dashboard**, then click **Read & validate**, **Apply changes**, **Refresh prices** and **Check & build report** in turn (or **Run all steps**).
5. *(Optional, for the web dashboard)* **Extensions › Apps Script › Deploy › New deployment › Web app**. Set *Execute as: Me* and *Who has access: Anyone* (the API checks its own password), then **Comparables Automation › Tools › Set web app password**.

### For a developer (from this repo)

```bash
npm i -g @google/clasp && clasp login
npm test                     # 17 tests against the real CIQ exports (fixtures needed, see below)
clasp push -f                # push appsscript/ to the bound script (.clasp.json)
npm run deploy               # push + update the web app deployment
node tools/report.js         # regenerate docs/VALIDATION_REPORT.md
```

`tests/fixtures/` is generated from the CIQ exports with `npm run fixtures`, using Python with `xlrd`. The CIQ exports and the Areta template are not committed, because they are licensed and confidential case materials. Put them in `ciq-exports/` and `template/` locally.

## Operation and scheduling

- **Weekly refresh**: `scheduledWeeklyPrices`. Day and hour come from the Config tab, and the default is Monday 07:00 in the Sheet's time zone (Europe/London). Re-run **Set up** after changing the schedule.
- **Daily snapshot** (optional, `dailySnapshot = YES`): weekday 22:00. It writes Price History only, so the Comps tab still rolls weekly.
- Runs take a lock, so two updates can't collide. Every run is logged on Run Log.

## Maintenance (no code needed)

| Task | How |
|---|---|
| Add a company | Insert a row in the right peer group on the Comps tab with name and CIQ ticker (e.g. `NasdaqGS:XYZ`), then **Set up**. Formulas and the medians are re-installed, and the ticker appears in the Config ticker map. |
| Remove a company | Delete its row, then **Set up**. |
| New peer group | Add a group label row, company rows and a "Median" row, then **Set up**. |
| Google Finance can't find a ticker | Edit its *Google Finance ticker* on the Config tab (e.g. `NYSEAMERICAN:EXOD`). |
| Change thresholds / schedule / access | Edit the Config tab. Re-run **Set up** for schedule changes. |
| A formula was overwritten | **Run checks** flags it, and **Repair formulas** restores it. |
| Different column order (existing Areta model) | The engine finds columns by header text, so headers need to match the template's (e.g. `EV / EBITDA`). |

## Limitations

- **Google Finance** quotes can be delayed by up to 20 minutes and are not for trading. Apps Script cannot read Google Finance *historical* data, so the first run seeds the prior price from Yahoo Finance daily closes. After that it uses the engine's own history.
- The **Yahoo Finance cross-check** uses a public, unofficial endpoint. If it is unavailable the engine says so (INFO) and carries on with Google Finance alone.
- **CIQ exports** must be the standard *Financials* export in millions. Key Stats must be in US$; the reporting currency can differ, and that is flagged. Other templates are read where the labels match. Anything not found is blocked, and the existing value is kept.
- **Implied EBITDA** for financial-template companies is a proxy, so it is flagged for review.
- **Robinhood** has a balance-sheet-heavy debt stack, like the brokers the Read Me excludes. The engine flags this rather than adjusting it.
- **Stale LTM**: all exports except BlackRock end Jun-30-2026, while Sep-30-2026 results are already published. Each one is flagged (INFO), and a refreshed export updates it.
- Apps Script quotas: a run is limited to 6 minutes. A full 20-file import takes about 1–2 minutes.
