# Packaging Instruction Enrichment & Decision Maker

Streamlit app for reviewing and enriching packaging-instruction cases from Excel files.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Or on Windows:

```bat
start_enrichment_maker.bat
```

## Inputs

- `Chybné instrukce` Excel file
- `Produkty a vlastnosti` Excel file

Optional:

- `Baliace pravidlá` Excel file

## Output

The app builds cleaned report/product tables, case enrichment, decision-maker outputs, and exports a single workbook for review.

## Persistent storage

The app stores the latest processed batch and case review state in a local SQLite database at `.data/instruction_validator.sqlite3`.
The database is created automatically on first run and is excluded from Git. To use another location, set the `INSTRUCTION_VALIDATOR_DB` environment variable before starting Streamlit.

When the app is opened without a new upload, it loads the latest saved batch. Review status, validator notes, and execution metadata are saved per case and survive a page refresh or app restart.

## Admin access and backups

The upload screen is visible only after admin login. Configure the admin password outside Git with `INSTRUCTION_VALIDATOR_ADMIN_PASSWORD`.
The app creates a SQLite backup after a new batch upload, an explicit review save, and confirmation that a change was executed. Backups are stored in `.data/backups/` and the newest 20 files are retained.
To use another backup directory, set `INSTRUCTION_VALIDATOR_BACKUP_DIR`.

For local Windows testing, configure the admin password in the process environment before starting the app:

```powershell
$env:INSTRUCTION_VALIDATOR_ADMIN_PASSWORD = "set-this-outside-the-repository"
streamlit run app.py
```

For shared use, configure this variable in the hosting service or Windows service account rather than in a committed file.

## Local network test

For a functional test with multiple users, run `start_enrichment_maker_network.bat` on the host laptop.
Users connected to the same network can then open `http://HOST_IPV4_ADDRESS:8501`, where `HOST_IPV4_ADDRESS` is the IPv4 address printed by the batch file.
The host laptop must remain powered on and connected to the same network. The Windows Firewall may need an inbound rule for TCP port `8501`.
