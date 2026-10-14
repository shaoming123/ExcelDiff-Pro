# Changelog / 更新记录

## v1.1.0 — 2026-10-12

- Fixed a CSS loading bug that printed raw stylesheet declarations above the page.
- Switched style injection to Streamlit `st.html(Path("assets/theme.css"))` and ensured the asset contains raw CSS only.
- Made **English** the first-load UI language; Simplified Chinese remains available.
- Added a bilingual customization / freelance service CTA with clickable `mailto:greatfoodprefer@gmail.com` link in both the sidebar and main page.
- Updated English + Chinese README, deployment and user guides and added tests that prevent regression of the raw CSS issue.

## v1.0.0 — 2026-10-12

- Introduced database-free Excel / CSV comparison web app (Streamlit).
- Added Chinese / English interface with Excel letter-based column selection.
- Added numeric, text and presence-only comparisons, composite matching keys.
- Added duplicate detection, explicit grouping option, empty key and invalid number handling.
- Added pre-flight diagnostics, preview records and original row tracking.
- Added styled XLSX reports, CSV exceptions, mapping JSON profiles and CLI tool.
- Added sample workbooks, user documentation, Docker instructions and automated tests.

### Known limitations / 已知限制

- UI supports one compared-value pair per run, plus one or two matching keys.
- Only regular XLSX / CSV tables are supported; complex merged/pivoted workbooks require preprocessing.
- Production hosting, security review, resource limits and access control must be handled by the deployer.
