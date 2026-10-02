# AGENTS.md — Agent execution rules

## Adaptive token and context efficiency (mandatory)

<!-- TOKEN_EFFICIENCY_POLICY_V1 -->

Optimize redundant context, not reasoning quality. Token/cost reduction must never weaken correctness, safety, verification, or evidence.

- Use progressive context: start with the smallest useful evidence, then automatically widen it when confidence is insufficient: `search/rg -> targeted range -> wider range -> whole file/log/DOM if needed`.
- Treat output limits as **soft**. Never hide evidence needed to diagnose a defect. If a summary or tail is insufficient, read more.
- Avoid replay-amplifying output: do not print full JSON payloads, database dumps, recursive directory listings, huge DOM snapshots, or long logs when counts, keys, aggregates, anomalies, samples, or a relevant tail answer the question.
- Save large stdout/stderr or reports as artifacts/files when practical. Keep model-visible output to command, exit code, key result/error, and a relevant excerpt; expand only when necessary.
- Do not repeat an unchanged read/check without a new reason. After edits, prefer a targeted diff before rereading the whole file. Re-read when dependencies, runtime state, hypotheses, or verification needs changed.
- Keep repository searches scoped. Avoid repository-wide `find`/`grep`/recursive scans unless the task genuinely requires them.
- Prefer targeted tests during iteration. Run broader/full regression when blast radius, milestone, pre-deploy verification, or unresolved uncertainty requires it. Never skip necessary regression merely to save tokens.
- Batch independent checks when safe, and avoid tool-call ping-pong that can be answered by one bounded command.
- Keep long work in compact structured state: goal, scope, verified facts, changed files, current error/blocker, tests, runtime/deploy state, next step, critical constraints.
- At natural phase boundaries (inspect / implement / verify / deploy), refresh the compact checkpoint. For very long sessions/turns, prepare a concise handoff and continue in a fresh session when practical instead of indefinitely replaying a large history.
- Do not reset context in the middle of an active diagnosis if doing so would discard useful working evidence.
- Use subagents for broad reading/research only when they preserve quality and their context isolation is known or measured; do not assume they are cheaper.
- Do not lower reasoning effort merely to reduce cost. First reduce redundant history, tool output, duplicate reads, and unnecessary steps; evaluate reasoning-effort changes separately with quality metrics.
- Optimize for **tokens/finished task and steps/finished task**, not minimum tokens per individual call. A cheaper call that causes more retries is not an optimization.

Quality guardrail: when compact evidence is not enough for a confident conclusion, automatically spend more context and verification. Correctness has priority over token savings.

