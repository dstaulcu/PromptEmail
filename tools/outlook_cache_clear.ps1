# Kill all Outlook processes
Get-Process OUTLOOK -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

# Clear Office add-in cache (for Office 2016/2019/365)
$cachePath = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
if (Test-Path $cachePath) {
    $maxRetries = 5
    $retryDelay = 2 # seconds
    $attempt = 0
    $success = $false
    while (-not $success -and $attempt -lt $maxRetries) {
        try {
            Remove-Item "$cachePath\*" -Recurse -Force -ErrorAction Stop
            $success = $true
        } catch {
            Write-Host "Attempt $($attempt + 1) to clear cache failed: $($_.Exception.Message)" -ForegroundColor Yellow
            Start-Sleep -Seconds $retryDelay
            $attempt++
        }
    }
    if (-not $success) {
        Write-Host "Some cache files could not be deleted after $maxRetries attempts. You may need to close other processes or run as administrator." -ForegroundColor Red
    }
    Write-Host "Office add-in cache cleared: $cachePath"
} else {
    Write-Host "Cache path not found: $cachePath"
}

Write-Host "Outlook closed and cache cleared. Launching Outlook..."
Start-Process OUTLOOK.EXE
