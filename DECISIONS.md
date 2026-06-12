# Decisions

## 2026-06-12: Adopt File-Based Project Memory

- Decision: Store durable project context in `PROJECT_CONTEXT.md`, `DECISIONS.md`, and `ROADMAP.md`.
- Reason: Claude and Codex sessions do not reliably share complete conversation history.
- Consequence: Important conclusions from project chats must be summarized into these files.
- Safety: Do not copy credentials, personal data, confidential raw data, or full chat transcripts.
