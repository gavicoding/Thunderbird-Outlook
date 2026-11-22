# Contributing to Thunderbird ⇄ Outlook

Thank you for your interest in improving **Thunderbird ⇄ Outlook**!

## Development Workflow

1. Fork the repository and create a feature branch from `main`:
   ```bash
   git checkout -b feat/your-change
   ```
2. Make focused changes and add/update unit tests in `tests/` whenever touching parsing or conversion logic (such as `vcard_util.py`).
3. Verify locally before opening a Pull Request:
   ```bash
   python -m compileall -q .
   python -m unittest discover -s tests -v
   ```
4. Open a Pull Request describing:
   - What problem your change solves
   - Whether it touches Thunderbird profiles, Outlook COM automation, or portable formats (`.eml` / `.vcf`)
   - How you verified it (unit tests or manual profile test)

## Safety Guidelines

- Never store or commit real `.eml`, `.msg`, `.vcf`, or profile `prefs.js` / `abook.sqlite` files in the repository.
- Preserve non-destructive behavior: migrations should always target dedicated output folders rather than overwriting existing mail or contact folders.
