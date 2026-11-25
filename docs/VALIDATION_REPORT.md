# Validation report: supplied CIQ exports

Generated 2026-11-25 by `tools/report.js`, which runs the same `Core.js` as the Apps Script engine. The live Sheet also builds this report on its **Validation Report** tab, with prices included.

## Summary

| | |
|---|---|
| Exports read | 20 of 20 |
| Companies matched to the model | 20 |
| Fields extracted | 200 |
| Market cap and EV rebuilt and reconciled to CIQ (±0.5%) | 18 of 18 with a CIQ EV bridge |
| ERROR items (value blocked) | 0 |
| REVIEW items (banker judgement) | 11 |
| INFO items | 28 |
| Peer LTM period | Jun-30-2026 |

## Items requiring banker judgement

| Company | Issue | Detail |
|---|---|---|
| BlackRock, Inc. | PRESS_RELEASE | LTM to Sep-30-2026 is built from a press release (not yet in a 10-Q/10-K). Figures may change on filing, and EBITDA may be less complete. |
| BlackRock, Inc. | PERIOD_MISMATCH | LTM ends Sep-30-2026 while most peers end Jun-30-2026. Multiples are not on a like-for-like period. |
| BlackRock, Inc. | IMPLIED_SHARES | CIQ also reports 162.2mm implied shares (incl. exchangeable units) vs 154.6mm listed shares (4.9%). Model uses listed shares per CIQ convention; banker to confirm. |
| Coincheck Group N.V. | FX_TRANSLATION | Company reports in JPY. CIQ translated LTM figures to US$ at today's spot rate (setting: "Today's Spot Rate"), not period-average rates. Confirm the FX convention. |
| PENN Entertainment, Inc. | NEGATIVE_NCI | NCI is negative ($-7.5mm), which reduces EV. Confirm the treatment. |
| Robinhood Markets, Inc. | NO_EV_BRIDGE | CIQ shows no enterprise value bridge for this company (financial-company template). Net debt and NCI taken from Capital Structure Summary, latest quarter. Confirm debt is corporate debt, not operating/client funding. |
| Robinhood Markets, Inc. | BROKER_FUNDING_IN_DEBT | Short-term borrowings are 90% of total debt ($20,536mm of $22,925mm). Likely broker / securities-lending funding. The Read Me excludes balance-sheet-intensive brokers for this reason; banker to decide treatment. |
| Robinhood Markets, Inc. | IMPLIED_EBITDA | CIQ does not report EBITDA for this company (financial template). Implied LTM EBITDA = operating income $2,270.0mm + D&A $91.0mm (cash flow D&A total + other amortization) = $2,361.0mm. Banker to confirm this is a meaningful metric for the business. |
| Tradeweb Markets Inc. | NO_EV_BRIDGE | CIQ shows no enterprise value bridge for this company (financial-company template). Net debt and NCI taken from Capital Structure Summary, latest quarter. Confirm debt is corporate debt, not operating/client funding. |
| Tradeweb Markets Inc. | IMPLIED_EBITDA | CIQ does not report EBITDA for this company (financial template). Implied LTM EBITDA = operating income $961.8mm + D&A $246.6mm (income statement) = $1,208.5mm. Banker to confirm this is a meaningful metric for the business. |
| Tradeweb Markets Inc. | IMPLIED_SHARES | CIQ also reports 234.0mm implied shares (incl. exchangeable units) vs 211.0mm listed shares (10.9%). Model uses listed shares per CIQ convention; banker to confirm. |

## Extracted values (US$ mm)

| Company | Ticker | Shares (mm) | Net debt | NCI | Pref. | LTM revenue | LTM EBITDA | Growth | Period | EV source | Reconciles to CIQ |
|---|---|--:|--:|--:|--:|--:|--:|--:|---|---|---|
| BlackRock, Inc. | NYSE:BLK | 154.6 | 1,125 | 6,927 | 0 | 28,542 | 12,006 | 24.7% | LTM Sep-30-2026 (press release) | CIQ TEV bridge | Yes |
| Bullish | NYSE:BLSH | 151.6 | 205 | 45 | 0 | 301 | -795 | 80.5% | LTM Jun-30-2026 | CIQ TEV bridge | Yes |
| Cboe Global Markets, Inc. | BATS:CBOE | 104.4 | -774 | 0 | 0 | 5,061 | 1,931 | 11.7% | LTM Jun-30-2026 | CIQ TEV bridge | Yes |
| Circle Internet Group, Inc. | NYSE:CRCL | 255.1 | -1,714 | 1 | 0 | 2,905 | 229 | 37.2% | LTM Jun-30-2026 | CIQ TEV bridge | Yes |
| CME Group Inc. | NasdaqGS:CME | 359.6 | 1,593 | 0 | 0 | 6,757 | 4,771 | 5.0% | LTM Jun-30-2026 | CIQ TEV bridge | Yes |
| Coinbase Global, Inc. | NasdaqGS:COIN | 263.8 | -2,121 | 0 | 0 | 6,044 | 700 | -9.2% | LTM Jun-30-2026 | CIQ TEV bridge | Yes |
| Coincheck Group N.V. | NasdaqGM:CNCK | 191.8 | -83 | 0 | 0 | 3,229 | -10 | 30.2% | LTM Jun-30-2026 | CIQ TEV bridge | Yes |
| CoinShares PLC | NasdaqGS:CSHR | 131.8 | 129 | 0 | 0 | 137 | -2,409 | -24.8% | LTM Jun-30-2026 | CIQ TEV bridge | Yes |
| DraftKings Inc. | NasdaqGS:DKNG | 496.5 | 932 | 0 | 0 | 6,223 | 123 | 15.0% | LTM Jun-30-2026 | CIQ TEV bridge | Yes |
| Exodus Movement, Inc. | NYSEAM:EXOD | 30.2 | -21 | 0 | 0 | 109 | -24 | -14.2% | LTM Jun-30-2026 | CIQ TEV bridge | Yes |
| Flutter Entertainment plc | NYSE:FLUT | 173.5 | 10,964 | 605 | 0 | 17,161 | 1,587 | 15.2% | LTM Jun-30-2026 | CIQ TEV bridge | Yes |
| Gemini Space Station, Inc. | NasdaqGS:GEMI | 129.7 | 311 | 0 | 0 | 207 | -365 | 51.5% | LTM Jun-30-2026 | CIQ TEV bridge | Yes |
| Intercontinental Exchange, Inc. | NYSE:ICE | 561.4 | 19,460 | 101 | 0 | 10,558 | 6,714 | 9.0% | LTM Jun-30-2026 | CIQ TEV bridge | Yes |
| Invesco Ltd. | NYSE:IVZ | 441.5 | 709 | 998 | 2,511 | 6,903 | 1,445 | 12.2% | LTM Jun-30-2026 | CIQ TEV bridge | Yes |
| MarketAxess Holdings Inc. | NasdaqGS:MKTX | 35.2 | -167 | 14 | 0 | 870 | 435 | 3.9% | LTM Jun-30-2026 | CIQ TEV bridge | Yes |
| Nasdaq, Inc. | NasdaqGS:NDAQ | 559.0 | 8,762 | 5 | 0 | 5,614 | 3,353 | 14.2% | LTM Jun-30-2026 | CIQ TEV bridge | Yes |
| PENN Entertainment, Inc. | NasdaqGS:PENN | 127.3 | 10,175 | -8 | 0 | 7,160 | 873 | 6.1% | LTM Jun-30-2026 | CIQ TEV bridge | Yes |
| Robinhood Markets, Inc. | NasdaqGS:HOOD | 899.1 | -980 | 61 | 0 | 4,932 | 2,361 | 38.3% | LTM Jun-30-2026 | Capital Structure Summary | n/a (no CIQ TEV) |
| T. Rowe Price Group, Inc. | NasdaqGS:TROW | 213.3 | -2,784 | 1,170 | 0 | 7,592 | 2,964 | 7.0% | LTM Jun-30-2026 | CIQ TEV bridge | Yes |
| Tradeweb Markets Inc. | NasdaqGS:TW | 211.0 | -1,908 | 704 | 0 | 2,204 | 1,208 | 14.1% | LTM Jun-30-2026 | Capital Structure Summary | n/a (no CIQ TEV) |

## For information

| Company | Issue | Detail |
|---|---|---|
| Bullish | STALE_PERIOD | Latest LTM is Jun-30-2026; results to Sep-30-2026 should be published by now. Request a refreshed CIQ export. |
| Bullish | NEGATIVE_EBITDA | LTM EBITDA is negative ($-794.5mm): EV / EBITDA shows NM. |
| Cboe Global Markets, Inc. | STALE_PERIOD | Latest LTM is Jun-30-2026; results to Sep-30-2026 should be published by now. Request a refreshed CIQ export. |
| Circle Internet Group, Inc. | STALE_PERIOD | Latest LTM is Jun-30-2026; results to Sep-30-2026 should be published by now. Request a refreshed CIQ export. |
| CME Group Inc. | STALE_PERIOD | Latest LTM is Jun-30-2026; results to Sep-30-2026 should be published by now. Request a refreshed CIQ export. |
| Coinbase Global, Inc. | STALE_PERIOD | Latest LTM is Jun-30-2026; results to Sep-30-2026 should be published by now. Request a refreshed CIQ export. |
| Coinbase Global, Inc. | REVENUE_DEFINITION | CIQ standardised revenue ($6,044mm) differs from as-reported revenue ($6,283mm) by 4.0%. Model uses CIQ standardised for comparability. |
| Coincheck Group N.V. | STALE_PERIOD | Latest LTM is Jun-30-2026; results to Sep-30-2026 should be published by now. Request a refreshed CIQ export. |
| Coincheck Group N.V. | NEGATIVE_EBITDA | LTM EBITDA is negative ($-10.3mm): EV / EBITDA shows NM. |
| CoinShares PLC | STALE_PERIOD | Latest LTM is Jun-30-2026; results to Sep-30-2026 should be published by now. Request a refreshed CIQ export. |
| CoinShares PLC | NEGATIVE_EBITDA | LTM EBITDA is negative ($-2,408.6mm): EV / EBITDA shows NM. |
| DraftKings Inc. | STALE_PERIOD | Latest LTM is Jun-30-2026; results to Sep-30-2026 should be published by now. Request a refreshed CIQ export. |
| Exodus Movement, Inc. | STALE_PERIOD | Latest LTM is Jun-30-2026; results to Sep-30-2026 should be published by now. Request a refreshed CIQ export. |
| Exodus Movement, Inc. | NEGATIVE_EBITDA | LTM EBITDA is negative ($-23.8mm): EV / EBITDA shows NM. |
| Flutter Entertainment plc | STALE_PERIOD | Latest LTM is Jun-30-2026; results to Sep-30-2026 should be published by now. Request a refreshed CIQ export. |
| Gemini Space Station, Inc. | STALE_PERIOD | Latest LTM is Jun-30-2026; results to Sep-30-2026 should be published by now. Request a refreshed CIQ export. |
| Gemini Space Station, Inc. | NEGATIVE_EBITDA | LTM EBITDA is negative ($-364.8mm): EV / EBITDA shows NM. |
| Intercontinental Exchange, Inc. | STALE_PERIOD | Latest LTM is Jun-30-2026; results to Sep-30-2026 should be published by now. Request a refreshed CIQ export. |
| Invesco Ltd. | STALE_PERIOD | Latest LTM is Jun-30-2026; results to Sep-30-2026 should be published by now. Request a refreshed CIQ export. |
| Invesco Ltd. | PREFERRED | Preferred equity of $2,511mm included in EV. |
| Invesco Ltd. | MATERIAL_NCI | NCI of $998mm (7.9% of market cap) included in EV. |
| MarketAxess Holdings Inc. | STALE_PERIOD | Latest LTM is Jun-30-2026; results to Sep-30-2026 should be published by now. Request a refreshed CIQ export. |
| Nasdaq, Inc. | STALE_PERIOD | Latest LTM is Jun-30-2026; results to Sep-30-2026 should be published by now. Request a refreshed CIQ export. |
| PENN Entertainment, Inc. | STALE_PERIOD | Latest LTM is Jun-30-2026; results to Sep-30-2026 should be published by now. Request a refreshed CIQ export. |
| Robinhood Markets, Inc. | STALE_PERIOD | Latest LTM is Jun-30-2026; results to Sep-30-2026 should be published by now. Request a refreshed CIQ export. |
| T. Rowe Price Group, Inc. | STALE_PERIOD | Latest LTM is Jun-30-2026; results to Sep-30-2026 should be published by now. Request a refreshed CIQ export. |
| T. Rowe Price Group, Inc. | MATERIAL_NCI | NCI of $1,170mm (5.4% of market cap) included in EV. |
| Tradeweb Markets Inc. | STALE_PERIOD | Latest LTM is Jun-30-2026; results to Sep-30-2026 should be published by now. Request a refreshed CIQ export. |

## Source traceability (example: Coinbase)

| Field | Tab | Cell | Label | Raw |
|---|---|---|---|---|
| LTM revenue | Key Stats | F16 | Total Revenue | 6043.753 |
| Revenue growth | Key Stats | F17 | Total Revenue growth over prior year | -0.092388 |
| LTM EBITDA | Key Stats | F22 | EBITDA | 699.549 |
| Shares outstanding | Key Stats | B46 | Shares Out. | 263.836923 |
| Net debt | Key Stats | B56 − B55 | Total Debt − Cash & ST Investments | 6668.29 − 8788.843 |
| Preferred equity | Key Stats | B57 | + Pref. Equity | - |
| NCI | Key Stats | B58 | + Total Minority Interest | - |
| Financial period | Key Stats | F13 | Key Financials period header | LTM² 12 months Jun-30-2026A |
| Revenue basis | Key Stats |  | Derived from CIQ export settings | Trading Currency |
| EBITDA basis | Key Stats |  | Derived from CIQ export settings | Trading Currency |

Every value the engine writes is logged like this on the Sheet's **Audit Trail** tab.
