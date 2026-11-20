<#
    Excel quality-assurance pass for the PayPulse workbook (Windows + Excel only).

    Opens the workbook, forces a full recalculation, scans every sheet for
    visible formula errors, records the recalculated value of the cells the
    documentation quotes, and exports a PNG of each sheet so the layout can be
    inspected. The workbook is never saved: the committed file stays exactly as
    the Python builder produced it.

    Usage:
      powershell -File scripts/excel_qa.ps1 -Workbook outputs/Payment_Reliability_Analytics_Dashboard.xlsx
                 -RenderDir work/qa -PreviewPath outputs/dashboard_preview.png
#>
param(
    [string]$Workbook = "outputs/Payment_Reliability_Analytics_Dashboard.xlsx",
    [string]$RenderDir = "work/qa",
    [string]$PreviewPath = "outputs/dashboard_preview.png",
    [string]$ReportPath = "work/qa/excel_qa_report.json",
    [int]$PreviewZoom = 130
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$wbPath = if ([System.IO.Path]::IsPathRooted($Workbook)) { $Workbook } else { Join-Path $root $Workbook }
$renderDir = if ([System.IO.Path]::IsPathRooted($RenderDir)) { $RenderDir } else { Join-Path $root $RenderDir }
$previewPath = if ([System.IO.Path]::IsPathRooted($PreviewPath)) { $PreviewPath } else { Join-Path $root $PreviewPath }
$reportPath = if ([System.IO.Path]::IsPathRooted($ReportPath)) { $ReportPath } else { Join-Path $root $ReportPath }

if (-not (Test-Path $wbPath)) { throw "Workbook not found: $wbPath" }
New-Item -ItemType Directory -Force -Path $renderDir | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $previewPath) | Out-Null

$xlScreen = 1; $xlBitmap = 2; $xlCellTypeFormulas = -4123; $xlCellTypeConstants = 2; $xlErrors = 16

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $true
$report = [ordered]@{ workbook = $wbPath; sheets = @(); errors = @(); values = [ordered]@{} }

try {
    $wb = $excel.Workbooks.Open($wbPath, 0, $false)
    $excel.CalculateFullRebuild()

    foreach ($ws in $wb.Worksheets) {
        $name = $ws.Name
        $ws.Activate()
        $errorCells = @()
        foreach ($cellType in @($xlCellTypeFormulas, $xlCellTypeConstants)) {
            try {
                $found = $ws.Cells.SpecialCells($cellType, $xlErrors)
                if ($null -ne $found) {
                    foreach ($cell in $found) {
                        $errorCells += ("{0}!{1} = {2}" -f $name, $cell.Address(0, 0), $cell.Text)
                    }
                }
            } catch { }   # SpecialCells throws when nothing matches, which is the healthy case
        }
        $used = $ws.UsedRange
        $report.sheets += [ordered]@{
            name        = $name
            index       = $ws.Index
            rows        = $used.Rows.Count
            columns     = $used.Columns.Count
            charts      = $ws.ChartObjects().Count
            error_cells = $errorCells.Count
        }
        if ($errorCells.Count -gt 0) { $report.errors += $errorCells }
        Write-Host ("[{0}] {1,-28} used {2}x{3}, charts {4}, error cells {5}" -f $ws.Index, $name, $used.Rows.Count, $used.Columns.Count, $ws.ChartObjects().Count, $errorCells.Count)
    }

    # Recalculated values of the figures quoted in the documentation.
    $dash = $wb.Worksheets.Item("Executive Dashboard")
    $sim = $wb.Worksheets.Item("Impact Simulator")
    $report.values["dashboard_transactions"] = $dash.Range("A5").Value2
    $report.values["dashboard_payment_value"] = $dash.Range("D5").Value2
    $report.values["dashboard_success_rate"] = $dash.Range("G5").Value2
    $report.values["dashboard_failure_rate"] = $dash.Range("J5").Value2
    $report.values["dashboard_p95"] = $dash.Range("M5").Value2
    $report.values["dashboard_retry_recovery"] = $dash.Range("P5").Value2
    $report.values["simulator_baseline_success_rate"] = $sim.Range("B8").Value2
    foreach ($r in 15..30) {
        $label = $sim.Cells.Item($r, 1).Value2
        if ($label -in @("Conservative", "Expected", "Ambitious")) {
            $report.values["scenario_${label}_recovered"] = $sim.Cells.Item($r, 5).Value2
            $report.values["scenario_${label}_value"] = $sim.Cells.Item($r, 6).Value2
            $report.values["scenario_${label}_uplift_pp"] = $sim.Cells.Item($r, 7).Value2
            $report.values["scenario_${label}_projected"] = $sim.Cells.Item($r, 8).Value2
        }
    }

    # --- Render every sheet -------------------------------------------------
    # Let Excel settle after the full rebuild: the first CopyPicture call
    # otherwise tends to fail on a still-busy clipboard.
    Start-Sleep -Milliseconds 800
    foreach ($ws in $wb.Worksheets) {
        $name = $ws.Name
        $safe = ($name -replace '[^A-Za-z0-9]+', '_')
        $ws.Activate()
        $used = $ws.UsedRange
        $lastRow = $used.Rows.Count + $used.Row - 1
        $lastCol = $used.Columns.Count + $used.Column - 1
        # Floating charts do not extend the used range, so grow the capture
        # window until it covers every chart: this is how chart clipping and
        # chart-over-table overlaps become visible in the QA renders.
        foreach ($co in $ws.ChartObjects()) {
            $right = $co.Left + $co.Width
            $bottom = $co.Top + $co.Height
            $c = $lastCol
            while ($c -lt 40 -and ($ws.Cells.Item(1, $c).Left + $ws.Cells.Item(1, $c).Width) -lt $right) { $c++ }
            if ($c -gt $lastCol) { $lastCol = $c }
            $r = $lastRow
            while ($r -lt 400 -and ($ws.Cells.Item($r, 1).Top + $ws.Cells.Item($r, 1).Height) -lt $bottom) { $r++ }
            if ($r -gt $lastRow) { $lastRow = $r }
        }
        $lastRow = [Math]::Min($lastRow, 75)
        $lastCol = [Math]::Min($lastCol, 24)
        $isDashboard = $name -eq "Executive Dashboard"
        $excel.ActiveWindow.Zoom = if ($isDashboard) { $PreviewZoom } else { 100 }
        $rng = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($lastRow, $lastCol))
        $target = if ($isDashboard) { $previewPath } else { Join-Path $renderDir ("sheet_{0:d2}_{1}.png" -f $ws.Index, $safe) }
        if (Test-Path $target) { Remove-Item $target -Force }

        $exported = $false
        for ($attempt = 1; $attempt -le 3 -and -not $exported; $attempt++) {
            try {
                $rng.CopyPicture($xlScreen, $xlBitmap) | Out-Null
                Start-Sleep -Milliseconds 350
                $chartObject = $ws.ChartObjects().Add(0, 0, $rng.Width + 8, $rng.Height + 8)
                $chart = $chartObject.Chart
                $chart.ChartArea.Border.LineStyle = -4142   # xlNone
                $chart.Paste()
                Start-Sleep -Milliseconds 250
                $chart.Export($target, "PNG") | Out-Null
                $chartObject.Delete()
                $exported = Test-Path $target
            } catch {
                Write-Host ("  render attempt {0} failed for {1}: {2}" -f $attempt, $name, $_.Exception.Message)
                Start-Sleep -Milliseconds 700
            }
        }
        if ($exported) {
            $size = (Get-Item $target).Length
            Write-Host ("  rendered {0,-28} -> {1} ({2:N0} bytes)" -f $name, (Split-Path -Leaf $target), $size)
        } else {
            $report.errors += "Failed to render sheet: $name"
            Write-Host ("  RENDER FAILED for {0}" -f $name)
        }
        $excel.ActiveWindow.Zoom = 100
    }

    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
}

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $reportPath) | Out-Null
$report | ConvertTo-Json -Depth 6 | Out-File -FilePath $reportPath -Encoding utf8
Write-Host ""
Write-Host ("QA report written to {0}" -f $reportPath)
if ($report.errors.Count -gt 0) {
    Write-Host "ERRORS DETECTED:"
    $report.errors | ForEach-Object { Write-Host ("  " + $_) }
    exit 1
}
Write-Host "No formula errors found in any sheet."
