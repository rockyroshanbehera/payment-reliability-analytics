<#
    Windows convenience wrapper around the pipeline.

        .\run_pipeline.ps1                 # full run
        .\run_pipeline.ps1 -Rows 20000     # quick run on a smaller dataset
        .\run_pipeline.ps1 -SkipPreview    # skip the Excel rendering step

    Identical to `python src/run_pipeline.py`; it exists so the project can be
    reproduced from a fresh PowerShell window without remembering the flags.
#>
param(
    [int]$Rows = 0,
    [switch]$SkipGenerate,
    [switch]$SkipPreview,
    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$arguments = @("src/run_pipeline.py")
if ($Rows -gt 0) { $arguments += @("--rows", $Rows) }
if ($SkipGenerate) { $arguments += "--skip-generate" }
if ($SkipPreview) { $arguments += "--skip-preview" }
if ($SkipTests) { $arguments += "--skip-tests" }

python @arguments
exit $LASTEXITCODE
