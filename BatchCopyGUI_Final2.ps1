# BatchCopyGUI_Distribute_ETA_Final_ASCII.ps1
# Even distribution across batches + progress bar + ETA
# PowerShell 5.1 compatible (ASCII-only, no ternary operators)

# Relaunch in STA if needed and keep console open
if ($host.Runspace.ApartmentState -ne 'STA') {
  $argsList = "-NoProfile -ExecutionPolicy Bypass -NoExit -STA -File `"$PSCommandPath`""
  Start-Process -FilePath "powershell.exe" -ArgumentList $argsList
  exit
}

$ErrorActionPreference = 'Stop'
trap {
  try {
    [System.Windows.Forms.MessageBox]::Show(
      ($_.Exception.Message + "`r`n`r`n" + $_.InvocationInfo.PositionMessage),
      "Error", 0, 16
    ) | Out-Null
  } catch {}
  break
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------- Helpers ----------
function Ensure-Dir([string]$d) {
  if (-not (Test-Path -LiteralPath $d)) {
    [void][System.IO.Directory]::CreateDirectory($d)
  }
}

function RelPath($root, $path) {
  try {
    $r = (Resolve-Path $root).Path
    if (-not $r.EndsWith('\')) { $r += '\' }
    $u1 = [uri]$r
    $u2 = [uri](Resolve-Path $path).Path
    $rel = $u1.MakeRelativeUri($u2).ToString().Replace('/', '\')
    return [System.Uri]::UnescapeDataString($rel)
  } catch {
    $rn = [System.IO.Path]::GetFullPath((Resolve-Path $root).Path).TrimEnd('\')
    $pn = [System.IO.Path]::GetFullPath((Resolve-Path $path).Path)
    if ($pn.StartsWith($rn, [System.StringComparison]::InvariantCultureIgnoreCase)) {
      return $pn.Substring($rn.Length).TrimStart('\')
    }
    return (Split-Path $pn -Leaf)
  }
}

# PS 5.1: safe move with overwrite behavior
function Move-FileCompat {
  param(
    [Parameter(Mandatory)][string]$Source,
    [Parameter(Mandatory)][string]$Destination
  )
  if (Test-Path -LiteralPath $Destination) {
    Remove-Item -LiteralPath $Destination -Force
  }
  try {
    [System.IO.File]::Move($Source, $Destination)
  } catch {
    [System.IO.File]::Copy($Source, $Destination, $true)
    Remove-Item -LiteralPath $Source -Force
  }
}

function BatchIndexFromRelPath([string]$rel, [int]$batches) {
  $sha = [System.Security.Cryptography.SHA1]::Create()
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($rel)
  $h = $sha.ComputeHash($bytes)
  $num = [BitConverter]::ToUInt32($h, 0)
  return ($num % $batches) + 1
}

function BatchRootAt([string]$destBase, [string]$prefix, [int]$index) {
  $name = ("{0}{1:D2}" -f $prefix, $index)
  return (Join-Path $destBase $name)
}

function Format-ETA([TimeSpan]$ts) {
  if ($ts.TotalHours -ge 1) { return ("{0:hh\:mm\:ss}" -f $ts) }
  else { return ("{0:mm\:ss}" -f $ts) }
}

function Get-TotalCount($root, $patterns) {
  $count = 0
  foreach ($p in $patterns) {
    foreach ($f in [System.IO.Directory]::EnumerateFiles($root, $p, [System.IO.SearchOption]::AllDirectories)) {
      $count++
      if ($count % 5000 -eq 0) { [System.Windows.Forms.Application]::DoEvents() }
    }
  }
  return $count
}

# ---------- UI ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Distribute Copy/Move - balanced batches (progress + ETA)"
$form.StartPosition = 'CenterScreen'
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
$form.ClientSize = New-Object System.Drawing.Size(980, 420)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

function Add-Label($t, $x, $y) { $l = New-Object System.Windows.Forms.Label; $l.Text = $t; $l.Location = New-Object System.Drawing.Point($x, $y); $l.AutoSize = $true; $form.Controls.Add($l); $l }
function Add-TextBox($x, $y, $w) { $t = New-Object System.Windows.Forms.TextBox; $t.Location = New-Object System.Drawing.Point($x, $y); $t.Size = New-Object System.Drawing.Size($w, 24); $form.Controls.Add($t); $t }
function Add-Button($txt, $x, $y, $w) { $b = New-Object System.Windows.Forms.Button; $b.Text = $txt; $b.Location = New-Object System.Drawing.Point($x, $y); $b.Size = New-Object System.Drawing.Size($w, 30); $form.Controls.Add($b); $b }
function Add-NUD($x, $y, $min, $max, $val) { $n = New-Object System.Windows.Forms.NumericUpDown; $n.Minimum = $min; $n.Maximum = $max; $n.Value = $val; $n.Location = New-Object System.Drawing.Point($x, $y); $n.Size = New-Object System.Drawing.Size(120, 24); $form.Controls.Add($n); $n }

$lblSrc = Add-Label "Source root:" 15 20
$txtSrc = Add-TextBox 140 18 740
$btnSrc = Add-Button "Browse..." 890 17 80
$btnSrc.Add_Click({ $d = New-Object System.Windows.Forms.FolderBrowserDialog; if ($d.ShowDialog() -eq 'OK') { $txtSrc.Text = $d.SelectedPath } })

$lblDst = Add-Label "Destination BASE:" 15 55
$txtDst = Add-TextBox 140 53 740
$btnDst = Add-Button "Browse..." 890 52 80
$btnDst.Add_Click({ $d = New-Object System.Windows.Forms.FolderBrowserDialog; if ($d.ShowDialog() -eq 'OK') { $txtDst.Text = $d.SelectedPath } })

$grpAct = New-Object System.Windows.Forms.GroupBox
$grpAct.Text = "Action"; $grpAct.Location = New-Object System.Drawing.Point(18, 90); $grpAct.Size = New-Object System.Drawing.Size(260, 64)
$optCopy = New-Object System.Windows.Forms.RadioButton; $optCopy.Text = "Copy"; $optCopy.AutoSize = $true; $optCopy.Location = New-Object System.Drawing.Point(15, 25); $optCopy.Checked = $true
$optMove = New-Object System.Windows.Forms.RadioButton; $optMove.Text = "Move"; $optMove.AutoSize = $true; $optMove.Location = New-Object System.Drawing.Point(85, 25)
$grpAct.Controls.AddRange(@($optCopy, $optMove)); $form.Controls.Add($grpAct)

$grpMode = New-Object System.Windows.Forms.GroupBox
$grpMode.Text = "Distribution mode (choose ONE)"; $grpMode.Location = New-Object System.Drawing.Point(290, 90); $grpMode.Size = New-Object System.Drawing.Size(680, 64)
$optByBatches = New-Object System.Windows.Forms.RadioButton; $optByBatches.Text = "By number of batches"; $optByBatches.AutoSize = $true; $optByBatches.Location = New-Object System.Drawing.Point(15, 25); $optByBatches.Checked = $true
$lblBatches  = Add-Label "Batches:" 400 110
$nudBatches  = Add-NUD 460 107 1 100000 5
$optBySize   = New-Object System.Windows.Forms.RadioButton; $optBySize.Text = "By files per batch (auto-calc batches)"; $optBySize.AutoSize = $true; $optBySize.Location = New-Object System.Drawing.Point(180, 25)
$lblFilesPer = Add-Label "Files per batch:" 600 110
$nudFilesPer = Add-NUD 700 107 1 100000000 20000
$grpMode.Controls.AddRange(@($optByBatches, $optBySize)); $form.Controls.Add($grpMode)

$action = {
  if ($optByBatches.Checked) {
    $nudBatches.Enabled = $true;  $lblBatches.Enabled  = $true
    $nudFilesPer.Enabled = $false; $lblFilesPer.Enabled = $false
  } else {
    $nudBatches.Enabled = $false; $lblBatches.Enabled  = $false
    $nudFilesPer.Enabled = $true;  $lblFilesPer.Enabled = $true
  }
}
$optByBatches.Add_CheckedChanged($action)
$optBySize.Add_CheckedChanged($action)
$null = $action.Invoke()   # <-- FIXED: just invoke; do not use & with .Invoke()

$lblPrefix = Add-Label "Batch folder prefix:" 15 160
$txtPrefix = Add-TextBox 140 157 160; $txtPrefix.Text = "batch_"

$chkDry = New-Object System.Windows.Forms.CheckBox
$chkDry.Text = "Dry run (no changes)"; $chkDry.AutoSize = $true; $chkDry.Location = New-Object System.Drawing.Point(320, 158)
$form.Controls.Add($chkDry)

$lblFilter = Add-Label "Optional file filter (e.g. *.jpg;*.png;*.pdf;*.docx):" 15 195
$txtFilter = Add-TextBox 320 193 650; $txtFilter.Text = "*"

$lblLog = Add-Label "CSV log (optional):" 15 230
$txtLog = Add-TextBox 140 228 740
$btnLog = Add-Button "Choose..." 890 227 80
$btnLog.Add_Click({
  $dlg = New-Object System.Windows.Forms.SaveFileDialog
  $dlg.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
  if ($dlg.ShowDialog() -eq 'OK') { $txtLog.Text = $dlg.FileName }
})

# Progress UI
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(18, 275)
$progress.Size     = New-Object System.Drawing.Size(952, 26)
$progress.Style    = 'Continuous'
$form.Controls.Add($progress)

$lblStatus = Add-Label "Status: idle" 18 310
$lblETA    = Add-Label "ETA: --:--"   18 335
$btnStart  = Add-Button "Start" 830 305 120

# ---------- Core ----------
function Validate-Roots {
  $s = $txtSrc.Text.Trim()
  $d = $txtDst.Text.Trim()
  if (-not (Test-Path -LiteralPath $s)) { [System.Windows.Forms.MessageBox]::Show("Source root not found."); return $null }
  if (-not (Test-Path -LiteralPath $d)) { Ensure-Dir $d }
  return @{ Source = $s; DestBase = $d }
}

$btnStart.Add_Click({
  $btnStart.Enabled = $false

  $r = Validate-Roots; if (-not $r) { $btnStart.Enabled = $true; return }
  $prefix = $txtPrefix.Text.Trim(); if ([string]::IsNullOrWhiteSpace($prefix)) { $prefix = "batch_" }
  $patterns = ($txtFilter.Text.Trim() -split ';' | ForEach-Object { if ([string]::IsNullOrWhiteSpace($_)) { '*' } else { $_ } })
  $move = $optMove.Checked
  $dry  = $chkDry.Checked
  $logp = $txtLog.Text.Trim()

  # Count files
  $total = Get-TotalCount $r.Source $patterns
  if ($total -le 0) { [System.Windows.Forms.MessageBox]::Show("No matching files found."); $btnStart.Enabled = $true; return }

  # Determine number of batches
  if ($optBySize.Checked) {
    $filesPer = [int]$nudFilesPer.Value
    $batches  = [int][Math]::Ceiling($total / [double]$filesPer)
  } else {
    $batches  = [int]$nudBatches.Value
  }
  if ($batches -lt 1) { $batches = 1 }

  # Prepare batch roots
  for ($i = 1; $i -le $batches; $i++) {
    Ensure-Dir (BatchRootAt $r.DestBase $prefix $i)
  }

  # Init progress
  $progress.Minimum = 0; $progress.Maximum = $total; $progress.Value = 0
  $lblStatus.Text = "Status: starting..."
  $startTime = Get-Date
  $done = 0; $errs = 0

  if ([string]::IsNullOrWhiteSpace($logp)) {
    $rows = $null
  } else {
    $rows = New-Object System.Collections.Generic.List[object]
  }

  foreach ($p in $patterns) {
    foreach ($f in [System.IO.Directory]::EnumerateFiles($r.Source, $p, [System.IO.SearchOption]::AllDirectories)) {
      if (-not (Test-Path -LiteralPath $f)) { continue }

      $rel = RelPath $r.Source $f
      $bi  = BatchIndexFromRelPath $rel $batches
      $root = BatchRootAt $r.DestBase $prefix $bi
      $tgt = Join-Path $root $rel
      Ensure-Dir (Split-Path $tgt -Parent)

      $act = if ($move) { "Move" } else { "Copy" }

      try {
        if (-not $dry) {
          if ($move) { Move-FileCompat -Source $f -Destination $tgt }
          else       { [System.IO.File]::Copy($f, $tgt, $true) }
        }
        if ($rows) {
          $rows.Add([pscustomobject]@{
            Timestamp = Get-Date
            Action    = $act
            Status    = 'Done'
            Source    = $f
            Target    = $tgt
            Batch     = $bi
            Message   = 'OK'
          }) | Out-Null
        }
      } catch {
        $errs++
        if ($rows) {
          $rows.Add([pscustomobject]@{
            Timestamp = Get-Date
            Action    = $act
            Status    = 'Error'
            Source    = $f
            Target    = $tgt
            Batch     = $bi
            Message   = $_.Exception.Message
          }) | Out-Null
        }
      }

      $done++
      if ($done -le $progress.Maximum) { $progress.Value = $done }

      if ($done % 200 -eq 0 -or $done -eq $total) {
        $elapsed = (Get-Date) - $startTime
        if ($elapsed.TotalSeconds -gt 0) { $rate = $done / $elapsed.TotalSeconds } else { $rate = 0 }
        $remain = $total - $done
        if ($rate -gt 0) { $etaTs = [TimeSpan]::FromSeconds($remain / $rate) } else { $etaTs = [TimeSpan]::Zero }
        $lblETA.Text = "ETA: " + (Format-ETA $etaTs)
        $lblStatus.Text = "Status: " + $done + " / " + $total + " processed  |  Errors: " + $errs + "  |  Batches: " + $batches
        [System.Windows.Forms.Application]::DoEvents()
      }
    }
  }

  if ($rows -ne $null) {
    try {
      $rows | Export-Csv -Path $logp -NoTypeInformation -Force
    } catch {
      [System.Windows.Forms.MessageBox]::Show("Log write failed: " + $_.Exception.Message)
    }
  }

  $lblStatus.Text = "Status: COMPLETE - " + $done + " files, " + $errs + " errors, " + $batches + " batches"
  $lblETA.Text = "ETA: 00:00"
  [System.Windows.Forms.MessageBox]::Show("Complete: " + $done + " files, " + $errs + " errors across " + $batches + " batches.", "Done", 0, 64) | Out-Null
  $btnStart.Enabled = $true
})

[void]$form.ShowDialog()
