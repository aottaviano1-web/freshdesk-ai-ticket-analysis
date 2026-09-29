# Lessons learned

Notes from running this on about a month of real support tickets across two teams (commercial and technical).
No real data is included here; the patterns are what generalises.

## 1. Give every AI agent the same fixed tag list
Running 10 agents in parallel only works if they all pick from **one shared instructions file** with fixed codes
(`prompts/tagging-instructions.md`). Otherwise "Login issue", "Portal access" and "Can't sign in" become three
categories and your totals stop adding up. Keep one free-text field (`sub_theme`) for detail. That's where specifics live.

## 2. Free-text labels multiply, so plan a grouping pass
Even with fixed themes, the free-text issue labels came out almost one-per-ticket. A second pass
(`prompts/2-cluster-issues.md`) groups them into 3–10 named clusters per theme, such as "Password reset link fails".
That list is what people actually act on.

## 3. Ask every agent to check its own work
Each tagging agent counts the tickets in its batch against the lines it wrote. One agent missed a single ticket
and caught it in that check. Writing output in parts (every 50–75 tickets) also protects against an agent
running out of context halfway through.

## 4. Automatic closure timers distort resolution time
Freshdesk automations that follow up with the customer and then close the ticket after roughly 10 days of silence
create a spike in resolution time right at the timer length. In a technical queue this can be
**a third or more of all resolutions**. Detect these tickets from the follow-up note text (configurable regex) and
report resolution **with and without** them. The honest story is usually "most open time is spent waiting on the customer",
and you can measure that directly (next point).

## 5. Measure who the ticket is waiting on
Split each ticket's open time by who spoke last. After an agent reply, the clock is on the customer; after a
customer message, it's on the team. For technical queues the customer's share of open time was typically large;
for quick commercial queues it was small. Report both rather than one blended number.

## 6. Decide metric definitions before you look at the results
Small definition choices move numbers by 20+ points:
- "Solved on first reply" (zero customer follow-up) vs. **FCR ≤ 2 customer replies**
- Calendar hours vs. **weekday hours** for first response
- Sentiment over all tickets vs. only tickets where the customer wrote back

All of these are legitimate choices. Pick the one that matches how your team already reports, **label it on the chart**,
and state the alternative in the method notes. Don't keep changing definitions until a target number appears;
readers who check the numbers against the source system will lose trust in the whole report.

## 7. Sentiment has a large "Unknown" group
Many customers never write back after the agent's answer, especially when a ticket is resolved. Report the base
explicitly: "X% of customers whose tone was visible ended positive or neutral". A real CSAT survey beats inferred tone.

## 8. AI quality scores are useful for coaching, not for precise rankings
Clarity, empathy and ownership scores cluster around 3–4 out of 5. Use them to:
- find **practices worth spreading**, from the per-ticket `notable` field
- see which dimension is weakest team-wide (for us, empathy was consistently below clarity)

Compare agents only **within their own team**, with a minimum ticket count, and blend quality with speed and FCR.

## 9. Two editions: remove sensitive data from the file, don't just hide it
The shareable HTML is built from metrics JSON with the per-agent array removed. Hiding a table with CSS would leave
the data readable in page source. The shareable workbook drops the Agents sheet, agent names and coaching notes.

## 10. Freshdesk API notes
- The list endpoint returns tickets updated in the last 30 days by default. Use `updated_since` and filter by `created_at` yourself.
- `include=requester,stats,description` saves several calls per ticket, although each include costs API credits.
- The account-wide rate limit gives predictable `429` responses. Sleep for `Retry-After` and continue.
- `stats.resolved_at` holds only the **latest** resolution, so a reopened ticket shows the second resolution time.
- Custom "source" channels come back as numbers that don't match the documented list. Label them in `config.json`.
- Ticket category fields are often blank. Reading the conversation beats relying on them.

## 11. Windows PowerShell 5.1 gotchas
- **Variable names ignore upper/lower case**: a `$Groups` parameter and a `$groups` lookup table are the same variable.
- `powershell -File script.ps1 -List "A","B"` passes **one** string, `"A,B"`. Split on commas inside the script.
- `R` is a built-in alias (`Invoke-History`), and aliases beat functions. Avoid one-letter function names.
- Relative paths: .NET methods (`[IO.File]::WriteAllText`) resolve against a different current directory than PowerShell does. Call `Resolve-Path` first.
- `@($list)` on a generic `List[object]` taken from a hashtable can throw "Argument types do not match". Use `.ToArray()`.
- `ConvertTo-Json` escapes `<`, `>` and `'` as `<` and so on, so plain-text find-and-replace on saved JSON fails. Edit it as an object instead.
