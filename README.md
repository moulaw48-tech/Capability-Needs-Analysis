# CNAMIS local application — Office of the Public Prosecutor

This is a local, single-agency web application based on the supplied OPP workbooks. It seeds a local SQLite database on first launch from:

- `Office of Public Presecutors (Responses).xlsx`
- `Form 3.17 - Occupancy Report updated.xlsx`
- `Capability Needs Analysis Hub.xlsx`

The source workbooks remain unchanged. After the first launch, plan and monitoring edits are stored in `cnamis.db`; editing an Excel workbook later does not silently overwrite those user edits.

## Source-control and data handling

Agency workbooks (`*.xlsx`), corporate-plan PDFs, extracted response text and the local SQLite database are intentionally excluded from source control because they contain agency and personnel records. To run a code-only checkout, place authorized copies of the three named source workbooks in this folder before starting CNAMIS. Do not commit generated CNA reports or database backups.

## Run on Windows

1. Install Python 3.10 or later.
2. From this folder, install the one external dependency:

   ```powershell
   python -m venv .venv
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
   ```

3. Start CNAMIS:

   ```powershell
   .\.venv\Scripts\python.exe app.py
   ```

4. Open <http://127.0.0.1:8766>.
5. Press `Ctrl+C` in the server terminal to stop it.

You can also run `run-cnamis.bat` after installing the dependency.

Set `CNAMIS_PORT` or `CNAMIS_DB_PATH` before starting if you need a different local port or database file.

## What is available

- Agency dashboard and Form 3.17/CNA reconciliation, including a manual verification workflow for ambiguous names.
- The supplied Papua New Guinea national emblem displayed beside CNAMIS, with a responsive navy, gold, green and light-blue theme.
- CNA item and dimension summaries using 1–3 Poor/Low, 4–8 Average, and 9–10 High/Excellent; gap is `10 − score`.
- Qualification review prompts, reported competencies, separate SPA 1–5 data, individual response cards, and seven phrase-matched E7 interest categories.
- The three planning options and the 35 editable 2027–2031 programme activities from the CNAMIS workbook.
- Editable timing, 10:20:70 learning approaches, budget estimates, provider-quote notes, approval status, indicators, and outcome notes.
- Monitoring of participants, expenditure, baselines, targets and results.
- Downloadable, detailed Capability Needs Analysis Report in PDF, generated from current OPP findings and reconciled participation evidence.
- Excel report export and a local evidence workspace that answers supported questions using the OPP records.

## Evidence and limitations

- The seed includes OPP only. No other agency records or existing CNAMIS backend/database were available.
- The evidence workspace is deterministic and local; it does not call an external AI model.
- The corporate-plan file supplied covers 2025–2030 and is image-only. Confirm the approved KRA references and successor priorities for 2031 before approval.
- Provider quotes, agency training eligibility rules, expected competency standards, attendance, targets and actual training outcomes were not supplied. Related fields remain blank or explicitly marked for validation.
- SPA 4–5 is not used as an automatic entitlement or eligibility rule.
- The server binds to loopback (`127.0.0.1`) and has no login or multi-user approval controls. Do not expose it to a network or treat local status changes as formal approval.
- PDF CNA and Excel report exports are available. Word export is not implemented.

Back up `cnamis.db` before moving the app or making source-data changes. To seed a fresh database, configure a new `CNAMIS_DB_PATH`; do not remove a database containing plan or monitoring edits.

## Tests

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```
