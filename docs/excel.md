# Excel import contract

Use an `.xlsx` workbook with headers on row 1 of the **active worksheet**. To import another sheet, make it active in Excel and save the workbook first. The app supports up to 200 columns and streams contact rows instead of loading the full workbook into memory. The preview shows at most five data rows; its row count comes from the worksheet metadata and may include formatted empty rows.

## Mapping

Header matching ignores spaces, punctuation and case. For example, `First Name`, `first_name` and `FIRSTNAME` all map to `First Name`.

| Destination | Accepted headers / aliases |
|---|---|
| First Name | First Name, nome |
| Last Name | Last Name, cognome |
| Organization | Organization, azienda, company |
| Title | Title, ruolo |
| Email | Email |
| Mobile | Mobile, cell, cellphone, cellulare |
| Switchboard | Switchboard, centralino |
| Direct Office | Direct Office, directofficephone, office |
| Address | Address, indirizzo |
| LinkedIn | LinkedIn |

Unknown nonempty headers become custom fields, preserving the previous import behavior. Review these suggestions: `Titolo` in the supplied sample is custom, while `Ruolo` maps to the job title. Choose **Ignore** to exclude a column. Duplicate destinations are rejected rather than silently selecting one value. Custom names must normalize to unique letters A–Z, digits and hyphens; rename unsupported headers in Excel or ignore those columns.

## Cell handling

- Phone numbers should be **text**, including `+`, country codes and leading zeros. Numeric cells trigger a report warning; their missing zeros cannot be inferred from the displayed formatting.
- Formula and error cells in mapped columns reject that row. The importer does not execute formulas or rely on cached results. Formulas in ignored columns are not imported.
- Completely empty mapped rows are skipped. A populated row needs a name, organization, email or phone number; an address or custom field alone is not sufficient.
- Dates and other literal values are converted to text. Email, phone, URL and custom-field rules still apply.
- Multiline text values are escaped in the vCard, so a line break cannot introduce another property.
- The workbook is read again during export; changed headers require reloading the preview. Contact values are read from the file at export time.

## Output and cancellation

Every run creates a separate folder under the selected destination. Names look like `0001_Alex_Morgan.png`; the prefix identifies the contact's data-row number. Names are sanitized and truncated for safe filesystem use. Duplicate contacts still produce distinct files.

`report.jsonl` uses the original Excel row number, including the header. For example:

```json
{"row": 2, "status": "saved", "file": "0001_Alex_Morgan.png", "warnings": []}
{"row": 3, "status": "failed", "error": "Enter an email address such as name@example.com."}
{"row": 4, "status": "skipped"}
```

The report is flushed after each row. A contact error does not stop later contacts. Failure to read the workbook or write the report stops the job and is shown in the UI. Completed exports remain on disk. Cancellation is cooperative between rows; opening a workbook or completing one PNG may take time before the cancellation request is observed. Closing the app during a job requests cancellation and waits for the worker to finish safely.

Appearance settings are captured at the start of a run. Editing them during conversion changes the next run, not the running one. Batch work never modifies the single-contact preview.
