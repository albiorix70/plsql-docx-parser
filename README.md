Project ebook-crush

To create the Word file, run:
```bash
pip install -r scripts/requirements.txt
python scripts/generate_test_docx.py
```

Want me to run the script here (I can't run arbitrary Python in your environment), or should I adjust the document contents (more sections, different table, etc.)?

CI Integration
--
This repository now includes a GitHub Actions workflow that compiles the `docx_parser` package using Oracle SQLcl. Workflow: `.github/workflows/compile.yml`.

Usage notes:
- Provide a direct download URL for SQLcl in the secret `SQLCL_ZIP_URL` (Oracle requires acceptance; use a private hosted copy if needed).
- Set DB connection secrets: `DB_USER`, `DB_PASSWORD`, `DB_HOST`, `DB_PORT`, `DB_SERVICE`.
- The workflow will download SQLcl, set up Java 11, and run `ci/run_compile.sh` which invokes `run_compile.sql`.

I can update the workflow to target another CI platform or to use a pre-built SQLcl container if you prefer.
