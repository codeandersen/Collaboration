# PowerShell script to fetch hard bounces from Brevo API

function Get-HardBounces {
    param (
        [string]$ApiKey,
        [datetime]$StartDate, # Start of the date range
        [datetime]$EndDate,   # End of the date range
        [int]$Limit = 50
    )

    $url = "https://api.brevo.com/v3/smtp/statistics/events"
    $headers = @{
        "accept" = "application/json"
        "api-key" = $ApiKey
    }

    $hardBounces = @()
    $offset = 0

    do {
        $params = @{ 
            "limit" = $Limit
            "offset" = $offset
            "startDate" = $StartDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
            "endDate" = $EndDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
            "event" = "hardBounce"
        }

        try {
            $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get -Body ($params | ConvertTo-Json -Depth 10)
        }
        catch {
            if ($_.Exception.Response.StatusCode -eq 429) {
                # Get rate limit reset time from headers
                $resetTime = $_.Exception.Response.Headers["x-sib-ratelimit-reset"]
                if ($resetTime) {
                    $waitSeconds = [int]$resetTime
                    Write-Warning "Rate limit hit. Waiting $waitSeconds seconds before retrying..."
                    Start-Sleep -Seconds $waitSeconds
                    # Retry the same request
                    $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get -Body ($params | ConvertTo-Json -Depth 10)
                }
                else {
                    # If we can't get the reset time, wait 60 seconds as a fallback
                    Write-Warning "Rate limit hit. Waiting 60 seconds before retrying..."
                    Start-Sleep -Seconds 60
                    $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get -Body ($params | ConvertTo-Json -Depth 10)
                }
            }
            else {
                throw
            }
        }

        if ($response -eq $null -or !$response.events) {
            Write-Output "No more events found."
            break
        }

        $hardBounces += $response.events

        if ($response.events.Count -lt $Limit) {
            break
        }

        $offset += $Limit

        # Add a small delay between successful requests to avoid hitting rate limits
        Start-Sleep -Seconds 1
    } while ($true)

    return $hardBounces
}

# Example usage
$APIKey = "" # Replace with your Brevo API key
$StartDate = (Get-Date).AddDays(-1).Date # Start of the range: yesterday
$EndDate = (Get-Date).Date               # End of the range: today

try {
    $hardBounces = Get-HardBounces -ApiKey $APIKey -StartDate $StartDate -EndDate $EndDate

    # Extract sender and recipient addresses
    $output = $hardBounces | Select-Object @{Name="Sender";Expression={$_."sender"}}, @{Name="Recipient";Expression={$_."recipient"}}

    # Save the results to a CSV file
    $outputFile = "hard_bounces.csv"
    $output | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8

    Write-Host "Fetched $($hardBounces.Count) hard bounces. Saved to $outputFile."
} catch {
    Write-Error "An error occurred: $_"
}
