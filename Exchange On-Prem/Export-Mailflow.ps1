#Requires -Version 5.1
<#
.SYNOPSIS
    Exports mail flow data from Exchange 2016 Frontend SMTP Receive logs.

.DESCRIPTION
    This script reads all SMTP Receive logs from the specified path for a given number of past months.
    It parses the logs to extract the sending IP address, sender email (MailFrom), and recipient email (MailTo).
    The data is then exported to a single CSV file.

.PARAMETER LogPath
    The full path to the root directory of the SMTP Receive logs.

.PARAMETER OutputPath
    The full path for the output CSV file.

.PARAMETER Months
    The number of past months of logs to analyze.

.EXAMPLE
    .\Export-Mailflow.ps1
    (Uses default values for all parameters)

.EXAMPLE
    .\Export-Mailflow.ps1 -LogPath "D:\ExchangeLogs" -OutputPath "C:\Temp\Export.csv" -Months 3
    (Uses custom paths and a 3-month timeframe)
#>
param (
    [string[]]$Servers = @('STGEXCHSRV01', 'STGEXCHSRV02'),
    [string]$OutputPath = "$PSScriptRoot\MailflowExport.csv"
)

# --- Script Execution ---

Write-Host "Gathering all available logs."

# --- Collect all log files first --- 
$allLogFiles = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
Write-Host "Discovering log files on all servers..."

foreach ($server in $Servers) {
    $logPath = "\\$($server)\C$\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive"
    if (-not (Test-Path -Path $logPath -PathType Container)) {
        Write-Warning "Log path not found for server '$server': '$logPath'. Skipping."
        continue
    }

    $serverFiles = Get-ChildItem -Path $logPath -Filter "*.log" -Recurse
    if ($null -ne $serverFiles) {
        $count = 0
        $serverFiles | ForEach-Object { 
            $allLogFiles.Add($_)
            $count++
        }
        Write-Host "Found $count log files on $server."
    }
}

if ($allLogFiles.Count -eq 0) {
    Write-Warning "No log files found on any of the specified servers."
    return
}

Write-Host "`nFound a total of $($allLogFiles.Count) log files to process. This may take some time..."

# --- Process all collected log files with a single progress bar ---
$sessions = @{}
$results = [System.Collections.Generic.List[pscustomobject]]::new()
$progress = 0

foreach ($file in $allLogFiles) {
    $progress++
    Write-Progress -Activity "Processing Log Files" -Status "File $progress of $($allLogFiles.Count): $($file.FullName)" -PercentComplete (($progress / $allLogFiles.Count) * 100)

    try {
        # Read the raw content of the log file
        $fileContent = Get-Content -Path $file.FullName -Raw

        # Find the header line and clean it for use with Import-Csv
        $headerLine = ($fileContent -split '\r?\n' | Where-Object { $_ -like '#Fields:*' })[0]
        if (-not $headerLine) { continue } # Skip file if no header found
        $headers = ($headerLine -replace '#Fields: ', '').Split(',')

        # Filter out comment lines and import the data as CSV
        $logData = $fileContent | Select-String -Pattern "^#" -NotMatch | ConvertFrom-Csv -Header $headers

        foreach ($row in $logData) {
            $sessionId = $row.'session-id'
            $data = $row.data
            $event = $row.event
            $context = $row.context

            # If it's a new session, initialize it
            if (-not $sessions.ContainsKey($sessionId)) {
                $clientIp = ($row.'remote-endpoint' -split ':')[0]
                $sessions[$sessionId] = [pscustomobject]@{
                    ClientIP = $clientIp
                    MailFrom = $null
                    Recipients = [System.Collections.Generic.List[string]]::new()
                }
            }

            # Capture MailFrom and Recipients
            if ($data -match 'MAIL FROM:<([^>]+)>') {
                $sessions[$sessionId].MailFrom = $Matches[1]
            } elseif ($data -match 'RCPT TO:<([^>]+)>') {
                $sessions[$sessionId].Recipients.Add($Matches[1])
            }

            # When the session ends definitively, process the collected data
            if ($event -eq '-' -and ($context -eq 'QUIT' -or $context -like 'Remote*')) {
                if ($sessions.ContainsKey($sessionId)) {
                    $sessionData = $sessions[$sessionId]
                    if ($sessionData.MailFrom -and $sessionData.Recipients.Count -gt 0) {
                        foreach ($recipient in $sessionData.Recipients) {
                            # Extract server name from the file path
                            $serverName = ($file.Directory.Root.FullName -split '\\')[2]
                            $results.Add([pscustomobject]@{
                                Server = $serverName
                                SendingIP = $sessionData.ClientIP
                                MailFrom = $sessionData.MailFrom
                                MailTo = $recipient
                            })
                        }
                    }
                    # Remove the completed session from memory
                    $sessions.Remove($sessionId)
                }
            }
        }
    } catch {
        Write-Warning "Error processing file $($file.FullName): $_"
    }
}

Write-Progress -Activity "Processing Log Files" -Completed

if ($results.Count -eq 0) {
    Write-Warning "Processing complete, but no mail flow data was extracted. The logs might be empty or in an unexpected format."
    return
}

# Export the results to a CSV file with Unicode encoding
Write-Host "Exporting $($results.Count) records to '$OutputPath'..."
$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding Unicode

Write-Host "Export complete!"
