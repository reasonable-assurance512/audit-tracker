# Audit Resource Tracker

Personal project: a Streamlit-based tool for audit resource planning.

## Notice to readers

This tool is a personal project. It contains no client, employer, or agency data, and does not connect to any workplace system.

It does not use, call, or transmit data to any artificial intelligence or large-language-model service. This is a deliberate design choice documented in the project's living specification.

All resource names are role-based (for example, PM, Auditor 1). No personal names, agency identifiers, or confidential information appear in the code or interface.

## Status

Sprint 0 (project scaffolding) — in progress.

Full sprint plan, architecture decisions, and design rationale are documented in `docs/AuditTracker_Living_Document_v0.4.docx`.

## Environment

- Python 3.11+
- Streamlit 1.56+
- openpyxl 3.1+
- pytest 9.0+

See `requirements.txt` for full dependency versions.

## Browser support

Designed and tested for Microsoft Edge in a Microsoft 365 environment. Other browsers may work but are not validated.

## Cold start

When hosted on Streamlit Community Cloud free tier, the application may take 10–30 seconds to wake after a period of inactivity.

## License

MIT — see `LICENSE`.
