#Start-Transcript -Path "c:\temp\log.txt" -Append -NoClobber -Force
################################

# SQL connection details
$SqlServer = "1"  # e.g. "localhost"
$SqlUser = "jobqueuecheck"
$SqlPassword = "jobqueuecheck"
# GUID for BC15 and BC23 (adjust if necessary)
$TableGuid = "437dbf0e-84ff-417a-965d-ed2bb9650972"


# Desired columns
$selectColumns = @(    
    "Object ID to Run",
    "Status",
    "Description",
    "Job Queue Category Code",
    "Earliest Start Date_Time"
)

# Initialize result list
$allRows = New-Object System.Collections.Generic.List[Object]


# Your server list
# if you only want to query one database, leave a single entry; multiple linked servers are allowed
# Version: 0 = pre BC15 (without table GUID), 1 = BC15 - BC21, 2 = BC21+ (new extension model)
$linkedServers = @(
    @{Server='sql-01';      DB='Live';    Company='Meine GmbH & Co_ KG'; Version=1 ; Country='DE'}
    #@{Server='sql-02';    DB='BC__ES';    Company='MyCompany_ES';        Version=1 ; Country='ES'},
    
)

$script:failedCompanies = @()

# Function for a single query
function Query-JobQueue {
    param ($srv)

    $tableName = switch ($srv.Version) {
        0 { "[$($srv.Company)`$Job Queue Entry]" }
        default { "[$($srv.Company)`$Job Queue Entry`$$TableGuid]" }
    }

    $columnString = ($selectColumns | ForEach-Object { "[$_]" }) -join ", "

    $sql = @"
SELECT 
    '$($srv.Server)' AS LinkedServer,
    '$($srv.Company)' AS Company,
    $columnString
FROM [$($srv.Server)].[$($srv.DB)].dbo.$tableName
WHERE 
    (
        [Status] = 0
        
    )
    
"@

    $connectionString = "Server=$SqlServer;Database=master;User Id=$SqlUser;Password=$SqlPassword;TrustServerCertificate=True"
    $conn = New-Object System.Data.SqlClient.SqlConnection $connectionString
    $cmd = $conn.CreateCommand()
    $cmd.CommandTimeout = 180
    $cmd.CommandText = $sql

    try {
        $conn.Open()
        $reader = $cmd.ExecuteReader()
        while ($reader.Read()) {
            $row = [ordered]@{
                #LinkedServer = $srv.Server
                Company      = $srv.Company
                Country = $srv.Country
            }
            foreach ($col in $selectColumns) {
                $row[$col] = $reader[$col]
            }
            $allRows.Add([PSCustomObject]$row)

        }
        Write-Host "✓ Erfolgreich: $($srv.Server)" -ForegroundColor Green
    } catch {
        Write-Warning "✗ Fehler bei $($srv.Server): $_"
        $script:failedCompanies += [PSCustomObject]@{
            Company = $srv.Company
            Country = $srv.Country
            Error   = $_.Exception.Message
        }

    } finally {
        $conn.Close()
    }
}

# Iterate over all servers
foreach ($srv in $linkedServers) {
    Query-JobQueue -srv $srv
}

# Display result
$allRows | Format-Table -AutoSize

# 🔧 Columns that should appear in the table
#$selectColumns = @(
#    "ID",
#    "User ID",
#    "Earliest Start Date_Time",
#    "Object Type to Run",
#    "Object ID to Run",
#    "Status",
#    "Description",
#    "Error Message"
#)

# 🏁 Function: ISO code → Flag
function Get-FlagFromCountry {
    param([string]$isoCode)
    if ($isoCode.Length -ne 2) { return "" }
    $flag = ""
    $isoCode.ToUpper().ToCharArray() | ForEach-Object {
        $flag += [char]::ConvertFromUtf32([int][char]$_ + 127397)
    }
    return $flag
}

# 🎨 CSS style definition
$style = @"
<style>
    body {
        font-family: Segoe UI, Arial, sans-serif;
        font-size: 13px;
        color: #333;
    }
    h1 {
        color: #0078d7;
        border-bottom: 2px solid #ccc;
    }
    h2 {
        margin-top: 40px;
        color: #0078d7;
    }
    table {
        border-collapse: collapse;
        width: 100%;
        margin-bottom: 30px;
    }
    th {
        background-color: #0078d7;
        color: white;
        padding: 8px;
        border: 1px solid #ccc;
        text-align: left;
    }
    td {
        padding: 8px;
        border: 1px solid #ccc;
    }
    tr:nth-child(even) {
        background-color: #f2f2f2;
    }
    .not-reachable {
        color: red;
        font-weight: bold;
    }
</style>
"@

# 🏁 Helper function: country code → flag (as image)
function Get-FlagImageFromCountry {
    param([string]$isoCode)
    if ([string]::IsNullOrWhiteSpace($isoCode)) { return "" }
    return "<img src='https://flagcdn.com/24x18/$($isoCode.ToLower()).png' alt='$isoCode' style='vertical-align:middle;margin-right:6px;' />"
}

# Column list without "Company" and "Country"
$columns = $allRows[0].PSObject.Properties.Name | Where-Object { $_ -notin @("Company", "Country") }

# 📦 Successful data groups
$groupedHtml = foreach ($group in $allRows | Group-Object Company) {
    $company = $group.Name
    $entries = $group.Group

    $country = $entries[0].Country
    $flag = Get-FlagImageFromCountry -isoCode $country

    $header = "<tr>" + ($columns | ForEach-Object { "<th>$_</th>" }) -join "" + "</tr>"

    $rows = foreach ($entry in $entries) {
        $cells = foreach ($col in $columns) {
            $value = $entry.$col
            if ($col -eq "Description" -and $entry.Status -eq 0) {
                "<td style='color:green; font-weight:bold;'>$value</td>"
            } else {
                "<td>$value</td>"
            }
        }
        "<tr>" + ($cells -join "") + "</tr>"
    }

    @"
<h2>$flag $company</h2>
<table>
$header
$($rows -join "`n")
</table>
"@
}

# 🚨 Unreachable servers/countries
$failedHtml = foreach ($entry in $failedCompanies) {
    $flag = Get-FlagImageFromCountry -isoCode $entry.Country
    $company = $entry.Company

    @"
<h2>$flag $company</h2>
<p class='not-reachable'>Server nicht erreichbar. Bitte manuell pruefen.</p>
"@
}

# 📾 Complete HTML document
$html = @"
<html>
<head>
<meta charset='utf-8'>
$style
</head>
<body>
<h1>Fehlerhafte Job Queue Entries $(Get-Date -Format "dd.MM.yyyy HH:mm")</h1>
$($groupedHtml -join "`n")
$($failedHtml -join "`n")
</body>
</html>
"@

# Open preview locally (optional)
# Comment out if the HTML file should not be checked beforehand
$tempFile = "$env:TEMP\jobqueue_report_land.html"
$html | Out-File -Encoding utf8 $tempFile
Start-Process $tempFile

$smtpServer   = "deinSmptServer"
$smtpPort     = 25                  # e.g. 25 or 465 depending on your provider
$smtpUser     = "deinsuer"
$smtpPassword = "deinpasswort"
$from         = "jobqueuecheck@company.com"
$to           = @("empfaenger1@mail.de", "empfaenger2@mail.de", "empfaenger3@mail.de") # any number of recipients separated by commas
$subject      = "Job Queue Overview $(Get-Date -Format 'dd.MM.yyyy HH:mm')"


# Prepare HTML content (from the previous step)
# -> If you already have $html from the previous step, this section is unnecessary.
# -> Example:
# $html = "<html><body><h1>Test</h1><p>Dies ist ein Test</p></body></html>"

# Convert password to SecureString
$securePassword = ConvertTo-SecureString $smtpPassword -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential($smtpUser, $securePassword)

# Send email
Send-MailMessage -From $from `
                 -To $to `
                 -Subject $subject `
                 -Body $html `
                 -BodyAsHtml `
                 -SmtpServer $smtpServer `
                 -Port $smtpPort `
                 -Credential $cred `
                 -UseSsl



#Stop-Transcript
