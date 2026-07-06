param(
    [Parameter(Mandatory=$true)][string]$Sha
)

$repo = "fanatics-hue/sakarya-dashboard"
$headers = @{ "User-Agent" = "AGGIORNA_DASHBOARD"; "Accept" = "application/vnd.github+json" }
$maxTries = 12
$finalState = $null
$logUrl = $null

for ($i = 0; $i -lt $maxTries; $i++) {
    Start-Sleep -Seconds 5
    try {
        $deployments = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/deployments?per_page=5" -Headers $headers -TimeoutSec 10
    } catch {
        continue
    }

    $dep = $deployments | Where-Object { $_.sha -eq $Sha } | Select-Object -First 1
    if (-not $dep) { continue }

    try {
        $statuses = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/deployments/$($dep.id)/statuses" -Headers $headers -TimeoutSec 10
    } catch {
        continue
    }
    if ($statuses.Count -eq 0) { continue }

    $latest = $statuses[0]
    if ($latest.state -in @("success", "failure", "error")) {
        $finalState = $latest.state
        $logUrl = $latest.log_url
        break
    }
}

if (-not $finalState) {
    Write-Output "STATO:TIMEOUT"
} elseif ($finalState -eq "success") {
    Write-Output "STATO:SUCCESS"
} else {
    Write-Output "STATO:FAILURE"
    Write-Output "LOG:$logUrl"
}
