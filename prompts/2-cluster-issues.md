# Prompt: group free-text issue labels into clusters

Even with a fixed theme list, the free-text `sub_theme` labels multiply. In our run the labels came out almost one per ticket. This second pass turns them into a list you can act on.

First create the input file (one line per non-noise ticket):

```powershell
Import-Csv .\run\output\tickets_tagged.csv | ? theme -ne 'NOISE' | Sort-Object theme, sub_theme |
  % { "$($_.ticket_id)|$($_.group)|$($_.theme)|$($_.sub_theme)" } | Set-Content .\run\subthemes.txt -Encoding UTF8
```

---

You are consolidating free-text support-ticket issue labels into clean clusters for a report. Local files only.

Input: `<RUN_DIR>\subthemes.txt`, with lines in the format `ticket_id|group|THEME|sub_theme`, sorted by THEME. Read the whole file.

Task: within EACH THEME, group the sub_theme labels into 3-10 "issue clusters". Give them clear, specific, business-readable names (e.g. "Password reset link fails", "License still bound to old server"). Every ticket gets exactly one cluster. Use a catch-all such as "Other <theme> issues" only for true one-offs, and keep it small. Mention product versions or error messages when many tickets share them.

Output 1: `<RUN_DIR>\issue_clusters.csv`, with header `ticket_id,theme,issue_cluster` and one row per ticket. Put the cluster name in double quotes. Verify that the row count and the set of ticket IDs match the input.

Output 2: `<RUN_DIR>\issue_clusters_summary.md`, listing for each THEME every cluster with its ticket count, split by group, and a one-line description.

Final reply: a brief confirmation, the row count, and the 15 largest clusters. Don't paste the full CSV.
