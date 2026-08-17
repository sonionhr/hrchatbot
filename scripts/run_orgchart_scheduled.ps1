# Regenerates html/orgchart.html from raw_docs/Orgchart.csv on a schedule
# (Windows Task Scheduler). Does NOT commit to git — review html/orgchart.html
# and the log file this run produces, then commit manually per the STEP 3/4
# spot-check-then-commit workflow in prompt-orgchart-update.md.

$RepoRoot = Split-Path -Parent $PSScriptRoot
$LogDir = Join-Path $PSScriptRoot "logs"
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }

$Timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
$LogFile = Join-Path $LogDir "orgchart_$Timestamp.log"

Set-Location $RepoRoot
& python "scripts\generate_orgchart.py" *>&1 | Tee-Object -FilePath $LogFile

# Keep only the 30 most recent run logs so this directory doesn't grow forever.
Get-ChildItem $LogDir -Filter "orgchart_*.log" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -Skip 30 |
    Remove-Item
