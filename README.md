# JobQueueCheck

`JobQueueCheck.ps1` is a PowerShell script that checks the **Job Queue Entry** table in one or more Business Central databases. It collects the entries with status `0`, builds a small HTML report and can optionally send it by email.

## Prerequisites
* PowerShell (Windows PowerShell or PowerShell 7)
* Access to the SQL server(s) hosting your Business Central databases
* Basic familiarity with PowerShell and SQL

## Configuration
Edit `JobQueueCheck.ps1` and adjust these variables:

- `$SqlServer`, `$SqlUser`, `$SqlPassword` – connection details for the SQL server
- `$TableGuid` – table GUID used for BC15 and BC23 (adjust if necessary)
- `$selectColumns` – columns from `Job Queue Entry` that should appear in the output
- `$linkedServers` – list of servers/databases/companies to query

Example entry for `$linkedServers`:

```powershell
$linkedServers = @(
    @{Server='sql-01'; DB='Live'; Company='Meine GmbH & Co_ KG'; Version=1 ; Country='DE'}
    #@{Server='sql-02'; DB='BC__ES'; Company='MyCompany_ES'; Version=1 ; Country='ES'},
)
```

## Running the script
Run the script in PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File JobQueueCheck.ps1
```

For each configured server the script performs the following steps:

1. Query `Job Queue Entry` for entries with status `0`.
2. Collect the rows into an ordered list.
3. Build an HTML table (including optional country flags).
4. Open the report in a browser for review.

If mail settings are supplied (`$smtpServer`, `$smtpUser`, etc.), the HTML report is also sent via `Send-MailMessage`.

## Customisation
* Modify `$selectColumns` to display additional columns.
* Adjust the CSS block inside the script to change the look of the report.
* Update the `Get-FlagFromCountry` or `Get-FlagImageFromCountry` helpers if you prefer different country indicators.

This repository only contains the script and this README. Use the script as a starting point for your own monitoring or automation tasks.
