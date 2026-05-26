# auto_push.ps1 — every 30s: commit + push any changes to GitHub
$repoPath = "C:\Users\MLAMP-8\Automation"
Set-Location $repoPath

Write-Host "Auto-push aktywny. Sprawdzam co 30 sekund. Ctrl+C aby zatrzymac."

while ($true) {
    Start-Sleep -Seconds 30
    $status = & git status --porcelain 2>$null
    if ($status) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        & git add -A
        & git commit -m "auto: $timestamp"
        & git push
        Write-Host "[$timestamp] Wyslano do GitHub"
    }
}
