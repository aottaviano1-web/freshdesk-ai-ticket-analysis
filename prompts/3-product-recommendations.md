# Prompt: turn product-related tickets into ranked recommendations

First create the input file:

```powershell
Import-Csv .\run\output\tickets_tagged.csv | ? { $_.theme -ne 'NOISE' -and ($_.product_fix -or $_.preventable_by -eq 'Product' -or $_.root_cause -in 'Product defect/bug','Product usability/UX gap','Missing feature') } |
  Sort-Object product_area, sub_theme |
  % { "$($_.ticket_id)|$($_.group)|$($_.product_area)|$($_.component)|$($_.root_cause)|$($_.sub_theme)|$($_.product_fix)|res_h=$($_.resolution_hours)|msgs=$($_.message_count)" } |
  Set-Content .\run\product_fixes.txt -Encoding UTF8
```

---

You are a senior product analyst writing the "For the Product team" section of a support-ticket analysis for <YOUR COMPANY / PRODUCT>. Local files only.

Input: `<RUN_DIR>\product_fixes.txt`, one line per ticket in the format
`ticket_id|group|product_area|component|root_cause|sub_theme|suggested product_fix|res_h=<resolution hours>|msgs=<message count>`.
Read the whole file. For extra context, you may Grep a few representative tickets in `<RUN_DIR>\batches\batch_*.txt` (each ticket starts with `### T<ticket_id>`).

Produce `<RUN_DIR>\product_recommendations.json`: a JSON array of 8-14 recommendations, ranked by impact (ticket count × effort per ticket). Each object has these fields:

```json
{
  "title": "short imperative title",
  "product_area": "...",
  "tickets": 0,
  "ticket_ids": [],
  "problem": "2-3 sentences, in the customer's terms; include versions or error messages if common",
  "evidence": "1-2 sentences with numbers (median resolution, messages per ticket, split by team)",
  "recommendation": ["2-4 concrete product changes, specific enough to write a ticket"],
  "type": "Bug fix | UX improvement | New feature | Self-service"
}
```

`ticket_ids` holds up to 6 representative IDs.

Group similar tickets, don't double-count, and keep the ticket counts honest. Also write `product_recommendations.md` with the same content in readable form.
Final reply: the list of titles with ticket counts.
