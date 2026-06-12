# Existing Chat Handoff Prompt

Send the text below once in the existing Claude Code or Codex chat for this project.

```text
Summarize the durable information from this project's existing conversation so future Claude and Codex sessions can continue the work.

Target project: 市場作り / アフィリンク

Read these files first:
- PROJECT_CONTEXT.md
- DECISIONS.md
- ROADMAP.md

Using only facts and decisions established in this conversation, update:
- PROJECT_CONTEXT.md: purpose, users, current architecture, operations, and important constraints
- DECISIONS.md: important decisions, reasons, and consequences, with dates
- ROADMAP.md: completed work, current priorities, next work, and blockers

Rules:
- Save concise durable conclusions, not the full conversation transcript
- Do not record API keys, credentials, personal data, customer data, or confidential raw data
- Do not turn uncertain information into a fact; mark it as unknown
- Do not implement code changes or deploy anything
- Preserve existing content and explicitly note contradictions

Finish by reporting the files updated and the items that still require confirmation.
```
