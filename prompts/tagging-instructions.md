# Ticket tagging instructions (template)

> Edit the product description, THEME list and product areas to match your business.
> Keep the THEME codes identical to `taxonomy.json`.

You are tagging support tickets for **<YOUR COMPANY>**, which makes **<SHORT PRODUCT DESCRIPTION>**.
Teams: **<TEAM A>** (commercial: accounts, licensing, billing, renewals) and **<TEAM B>** (technical).

Each ticket in your batch file starts with a header line:
`### T<ticket_id> | group=... | agent=... | cat=... | req=... | tier=... | status=... | msgs=N`
followed by SUBJECT and messages tagged [CUST] (customer), [AGENT] (public agent reply) and [NOTE] (internal note).
Texts are shortened and may be in several languages. Tag in English.

Tag EVERY ticket. Output ONE JSON object per line (JSONL), with no extra text, using exactly these keys:

| key | values |
|---|---|
| `ticket_id` | number (from the header, without the T) |
| `theme` | ONE code from the THEME list below |
| `sub_theme` | short, specific label of 2-6 words, e.g. "Password reset link fails" |
| `product_area` | ONE of the product areas in `taxonomy.json`, or `None` |
| `component` | ONE of the components in `taxonomy.json` if product-related, else `` |
| `root_cause` | ONE of: `Product defect/bug`, `Product usability/UX gap`, `Missing feature`, `Documentation/KB gap`, `Licensing/commercial process`, `Portal/self-service gap`, `Customer environment/config`, `Customer how-to question`, `Internal process/routing`, `Noise` |
| `preventable_by` | ONE of: `Product`, `Self-service/KB`, `Process/Automation`, `Not preventable` |
| `product_fix` | if the product could have prevented this: one concrete sentence on what it could do; else `` |
| `summary` | at most 20 words: what the customer needed and what happened |
| `resolved` | `Yes`, `No`, `Unclear` |
| `first_reply_resolved` | `Yes` if the first agent reply solved it with no further customer follow-up; `No`; `N/A` if there was no agent reply |
| `clarity` | 1-5 agent communication clarity; `0` if there was no agent reply |
| `empathy` | 1-5 acknowledges impact, polite, human tone; `0` if there was no agent reply |
| `ownership` | 1-5 proactive, sets next steps, no ping-pong; `0` if there was no agent reply |
| `sentiment_end` | customer tone at the end: `Positive`, `Neutral`, `Negative`, `Unknown` (use Unknown when the customer never wrote back) |
| `escalated` | `Yes` if escalated to engineering or another team, else `No` |
| `notable` | optional, at most 25 words: something worth noting for coaching; else `` |

## THEME codes
Commercial: `ACCOUNT_ACCESS`, `LICENSE_ACTIVATION`, `RENEWAL_CANCEL`, `BILLING_INVOICE`, `VENDOR_COMPLIANCE`, `ACCOUNT_DATA`, `SALES_QUOTE_TRIAL`, `CALLBACK_CHASE`
Technical: `INSTALL_UPGRADE`, `MIGRATION`, `CONFIG_HOWTO`, `BUG_ERROR`, `PERFORMANCE_STABILITY`, `INTEGRATION_API`, `NOTIFICATIONS`, `USERS_SSO`, `SECURITY`, `REPORTING_DATA`, `FEATURE_REQUEST`
Noise: `NOISE`: spam, marketing, job applications, auto-replies, bounces, test tickets, internal notifications with no customer need

## Rules
- Judge on the actual conversation, not only the subject or category fields (those are often blank or wrong).
- Calibrate quality scores: 3 = acceptable, 4 = good, 5 = excellent (reserve for genuinely great), 1-2 = poor. Score only public [AGENT] replies.
- Output valid JSON: escape quotes and keep each value on one line.
