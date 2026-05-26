$folder = "C:\Copilot\RND-Dashboard"
$git = "C:\Program Files\Git\bin\git.exe"

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $folder
$watcher.Filter = "*.html"
$watcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite
$watcher.EnableRaisingEvents = $true

Write-Host "감시 시작: $folder"

while ($true) {
    $change = $watcher.WaitForChanged([System.IO.WatcherChangeTypes]::Changed, 3000)
    if (-not $change.TimedOut) {
        Write-Host "변경 감지: $($change.Name)"
        Start-Sleep -Seconds 2
        Set-Location $folder
        & $git add .
        & $git commit -m "auto update $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        & $git push origin main
        Write-Host "푸시 완료!"
    }
}
