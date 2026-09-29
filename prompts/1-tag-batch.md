# Prompt: tag one batch of tickets

Run one agent per batch file, all in parallel (about 200 tickets per agent works well).
A mid-size model is fine for this; replace `NN` and `<RUN_DIR>`.

---

You are tagging customer support tickets for an analysis. Work only with local files; no web access is needed.

1. Read the instructions: `<KIT_DIR>\prompts\tagging-instructions.md` and the allowed values in `<KIT_DIR>\taxonomy.json`.
2. Read your ENTIRE batch file: `<RUN_DIR>\batches\batch_NN.txt` (about 200 tickets). It will need several Read calls with offset/limit, e.g. 500 lines at a time. Make sure you reach the end of the file.
3. Tag EVERY ticket according to the instructions. Write the JSONL output in parts as you go, so no work is lost: `<RUN_DIR>\tags\batch_NN_part1.jsonl`, `_part2.jsonl`, and so on (about 50-75 tickets per part). One JSON object per line and nothing else in the files.
4. Before finishing, verify: count the `### T` headers in your batch file and make sure the total number of lines across your part files equals it, with no ticket missing or duplicated. Fix any gaps.

Final reply: only the number of tickets tagged, the part files written, and any tickets you could not tag. Do not paste the tags into your reply.

---

**Why it's written this way**
- *Writing in parts* protects against an agent running out of context halfway through.
- *The self-check* matters: in our run, one agent initially missed one ticket and caught it in the verification step.
- *A shared instructions file* keeps 10 parallel agents on the same labels, so totals add up across batches.
