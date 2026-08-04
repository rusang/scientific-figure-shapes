param(
    [Parameter(Mandatory = $true)]
    [string]$VbaFile,
    [string]$MacroName = "ReconstructFromImage"
)

function Write-FigureJson {
    param(
        [string]$Status,
        [hashtable]$Fields = @{},
        [int]$ExitCode = 0
    )
    $payload = @{ status = $Status }
    foreach ($key in $Fields.Keys) {
        $payload[$key] = $Fields[$key]
    }
    $payload | ConvertTo-Json -Depth 8
    exit $ExitCode
}

$resolved = Resolve-Path -LiteralPath $VbaFile -ErrorAction SilentlyContinue
if (-not $resolved) {
    Write-FigureJson "vba_file_not_found" @{ path = $VbaFile } 2
}

try {
    $powerpoint = New-Object -ComObject PowerPoint.Application
} catch {
    Write-FigureJson "powerpoint_com_unavailable" @{ error = $_.Exception.Message } 2
}

$attempts = New-Object System.Collections.Generic.List[object]
try {
    $powerpoint.Visible = $true
    $presentation = $powerpoint.Presentations.Add()
    $module = $presentation.VBProject.VBComponents.Import($resolved.Path)
    $attempts.Add(@{ mode = "import_module"; component = $module.Name })
    $powerpoint.Run($MacroName)
    Write-FigureJson "ran" @{ vba_file = $resolved.Path; macro = $MacroName; attempts = $attempts }
} catch {
    $attempts.Add(@{ mode = "automation"; error = $_.Exception.Message })
    Write-FigureJson "automation_failed" @{
        vba_file = $resolved.Path
        macro = $MacroName
        attempts = $attempts
        fallback = "PowerPoint may be blocking programmatic VBA access. Import the .bas manually or enable trusted VBA project access."
    } 1
}
