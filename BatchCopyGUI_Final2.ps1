# BatchCopyGUI_Final_Features.ps1
# Copy/Move with balanced batch distribution, progress + ETA, Estimate, Cancel
# Auto-delete empty source folders after MOVE (only if not cancelled)
# PowerShell 5.1 compatible (ASCII only; no typed foreach; no ternary)

# ---------- Bootstrap ----------
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
  if (-not (Test-Path -LiteralPath $d)) { [void][System.IO.Directory]::CreateDirectory($d) }
}
function RelPath($root, $path) {
  try {
    $r = (Resolve-Path $root).Path; if (-not $r.EndsWith('\')) { $r += '\' }
    $u1 = [uri]$r; $u2 = [uri](Resolve-Path $path).Path
    $rel = $u1.MakeRelativeUri($u2).ToString().Replace('/', '\')
    return [System.Uri]::UnescapeDataString($rel)
  } catch {
    $rn = [System.IO.Path]::GetFullPath((Resolve-Path $root).Path).TrimEnd('\')
    $pn = [System.IO.Path]::GetFullPath((Resolve-Path $path).Path)
    if ($pn.StartsWith($rn, [System.StringComparison]::InvariantCultureIgnoreCase)) { return $pn.Substring($rn.Length).TrimStart('\') }
    return Split-Path $pn -Leaf
  }
}
function Move-FileCompat {
  param([Parameter(Mandatory)][string]$Source,[Parameter(Mandatory)][string]$Destination)
  if (Test-Path -LiteralPath $Destination) { Remove-Item -LiteralPath $Destination -Force }
  try { [System.IO.File]::Move($Source, $Destination) }
  catch { [System.IO.File]::Copy($Source, $Destination, $true); Remove-Item -LiteralPath $Source -Force }
}
function BatchIndexFromRelPath([string]$rel, [int]$batches) {
  $sha = [System.Security.Cryptography.SHA1]::Create()
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($rel)
  $h = $sha.ComputeHash($bytes)
  $num = [BitConverter]::ToUInt32($h, 0)
  return ($num % $batches) + 1
}
function BatchRootAt([string]$destBase,[string]$prefix,[int]$index) {
  $name = ("{0}{1:D2}" -f $prefix, $index)
  return (Join-Path $destBase $name)
}
function Format-ETA([TimeSpan]$ts) {
  if ($ts.TotalHours -ge 1) { return ("{0:hh\:mm\:ss}" -f $ts) }
  else { return ("{0:mm\:ss}" -f $ts) }
}
function Get-TotalCount($root,$patterns) {
  $count = 0
  foreach ($p in $patterns) {
    foreach ($f in [System.IO.Directory]::EnumerateFiles($root, $p, [System.IO.SearchOption]::AllDirectories)) {
      $count++
      if ($count % 5000 -eq 0) { [System.Windows.Forms.Application]::DoEvents() }
    }
  }
  return $count
}
function AnchorLR($ctrl){ $ctrl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right }
function AnchorBR($ctrl){ $ctrl.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right }

# ---------- UI ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Distribute Copy/Move - balanced batches (progress + ETA)"
$form.StartPosition = 'CenterScreen'
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
$form.ClientSize = New-Object System.Drawing.Size(1180, 540)
$form.MinimumSize = New-Object System.Drawing.Size(1000, 540)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

# Row 1: Source
$lblSrc = New-Object System.Windows.Forms.Label
$lblSrc.Text = "Source root:"
$lblSrc.Location = '16,18'; $lblSrc.AutoSize = $true; $form.Controls.Add($lblSrc)

$txtSrc = New-Object System.Windows.Forms.TextBox
$txtSrc.Location = '120,15'; $txtSrc.Size = New-Object System.Drawing.Size(960,24); AnchorLR $txtSrc; $form.Controls.Add($txtSrc)

$btnSrc = New-Object System.Windows.Forms.Button
$btnSrc.Text = "Browse..."; $btnSrc.Location = '1090,14'; $btnSrc.Size = New-Object System.Drawing.Size(80,26); $form.Controls.Add($btnSrc)
$btnSrc.Add_Click({ $d=New-Object System.Windows.Forms.FolderBrowserDialog; if($d.ShowDialog() -eq 'OK'){ $txtSrc.Text=$d.SelectedPath } })

# Row 2: Destination
$lblDst = New-Object System.Windows.Forms.Label
$lblDst.Text = "Destination BASE:"
$lblDst.Location = '16,52'; $lblDst.AutoSize = $true; $form.Controls.Add($lblDst)

$txtDst = New-Object System.Windows.Forms.TextBox
$txtDst.Location = '120,49'; $txtDst.Size = New-Object System.Drawing.Size(960,24); AnchorLR $txtDst; $form.Controls.Add($txtDst)

$btnDst = New-Object System.Windows.Forms.Button
$btnDst.Text = "Browse..."; $btnDst.Location = '1090,48'; $btnDst.Size = New-Object System.Drawing.Size(80,26); $form.Controls.Add($btnDst)
$btnDst.Add_Click({ $d=New-Object System.Windows.Forms.FolderBrowserDialog; if($d.ShowDialog() -eq 'OK'){ $txtDst.Text=$d.SelectedPath } })

# Row 3: Action + Distribution
$grpAct = New-Object System.Windows.Forms.GroupBox
$grpAct.Text="Action"; $grpAct.Location='16,90'; $grpAct.Size=New-Object System.Drawing.Size(240,80); $form.Controls.Add($grpAct)

$optCopy = New-Object System.Windows.Forms.RadioButton
$optCopy.Text="Copy"; $optCopy.Location='14,32'; $optCopy.AutoSize=$true; $optCopy.Checked=$true; $grpAct.Controls.Add($optCopy)

$optMove = New-Object System.Windows.Forms.RadioButton
$optMove.Text="Move"; $optMove.Location='90,32'; $optMove.AutoSize=$true; $grpAct.Controls.Add($optMove)

$grpMode = New-Object System.Windows.Forms.GroupBox
$grpMode.Text="Distribution mode"; $grpMode.Location='270,90'; $grpMode.Size=New-Object System.Drawing.Size(900,80); AnchorLR $grpMode; $form.Controls.Add($grpMode)

$optByBatches = New-Object System.Windows.Forms.RadioButton
$optByBatches.Text="By number of batches"; $optByBatches.Location='16,32'; $optByBatches.AutoSize=$true; $optByBatches.Checked=$true; $grpMode.Controls.Add($optByBatches)

$lblBatches = New-Object System.Windows.Forms.Label
$lblBatches.Text="Batches:"; $lblBatches.Location='190,34'; $lblBatches.AutoSize=$true; $grpMode.Controls.Add($lblBatches)

$nudBatches = New-Object System.Windows.Forms.NumericUpDown
$nudBatches.Minimum=1; $nudBatches.Maximum=100000; $nudBatches.Value=5; $nudBatches.Location='250,30'; $nudBatches.Size=New-Object System.Drawing.Size(110,24); $grpMode.Controls.Add($nudBatches)

$optBySize = New-Object System.Windows.Forms.RadioButton
$optBySize.Text="By files per batch (auto)"; $optBySize.Location='400,32'; $optBySize.AutoSize=$true; $grpMode.Controls.Add($optBySize)

$lblFilesPer = New-Object System.Windows.Forms.Label
$lblFilesPer.Text="Files per batch:"; $lblFilesPer.Location='580,34'; $lblFilesPer.AutoSize=$true; $grpMode.Controls.Add($lblFilesPer)

$nudFilesPer = New-Object System.Windows.Forms.NumericUpDown
$nudFilesPer.Minimum=1; $nudFilesPer.Maximum=100000000; $nudFilesPer.Value=20000; $nudFilesPer.Location='680,30'; $nudFilesPer.Size=New-Object System.Drawing.Size(120,24); $grpMode.Controls.Add($nudFilesPer)

$action = {
  if ($optByBatches.Checked) {
    $nudBatches.Enabled=$true; $lblBatches.Enabled=$true
    $nudFilesPer.Enabled=$false; $lblFilesPer.Enabled=$false
  } else {
    $nudBatches.Enabled=$false; $lblBatches.Enabled=$false
    $nudFilesPer.Enabled=$true; $lblFilesPer.Enabled=$true
  }
}
$optByBatches.Add_CheckedChanged($action)
$optBySize.Add_CheckedChanged($action)
$null = $action.Invoke()

# Row 4: Prefix
$lblPrefix = New-Object System.Windows.Forms.Label
$lblPrefix.Text="Batch folder prefix:"; $lblPrefix.Location='16,188'; $lblPrefix.AutoSize=$true; $form.Controls.Add($lblPrefix)

$txtPrefix = New-Object System.Windows.Forms.TextBox
$txtPrefix.Location='140,185'; $txtPrefix.Size=New-Object System.Drawing.Size(180,24); $txtPrefix.Text='batch_'; $form.Controls.Add($txtPrefix)

# Row 5: Filter
$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text="Optional file filter (e.g. *.jpg;*.png;*.pdf;*.docx):"; $lblFilter.Location='16,222'; $lblFilter.AutoSize=$true; $form.Controls.Add($lblFilter)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Location='320,219'; $txtFilter.Size=New-Object System.Drawing.Size(850,24); $txtFilter.Text='*'; AnchorLR $txtFilter; $form.Controls.Add($txtFilter)

# Row 6: CSV log
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text="CSV log (optional):"; $lblLog.Location='16,256'; $lblLog.AutoSize=$true; $form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location='140,253'; $txtLog.Size=New-Object System.Drawing.Size(940,24); AnchorLR $txtLog; $form.Controls.Add($txtLog)

$btnLog = New-Object System.Windows.Forms.Button
$btnLog.Text="Choose..."; $btnLog.Location='1090,252'; $btnLog.Size=New-Object System.Drawing.Size(80,26); $form.Controls.Add($btnLog)
$btnLog.Add_Click({ $dlg=New-Object System.Windows.Forms.SaveFileDialog; $dlg.Filter="CSV files (*.csv)|*.csv|All files (*.*)|*.*"; if($dlg.ShowDialog() -eq 'OK'){ $txtLog.Text=$dlg.FileName } })

# Row 7: Progress + controls
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location='16,300'; $progress.Size=New-Object System.Drawing.Size(1154,28); $progress.Style='Continuous'; AnchorLR $progress; $form.Controls.Add($progress)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text="Status: idle"; $lblStatus.Location='16,340'; $lblStatus.AutoSize=$true; $form.Controls.Add($lblStatus)

$lblETA = New-Object System.Windows.Forms.Label
$lblETA.Text="ETA: --:--"; $lblETA.Location='16,364'; $lblETA.AutoSize=$true; $form.Controls.Add($lblETA)

$btnEstimate = New-Object System.Windows.Forms.Button
$btnEstimate.Text="Estimate count"; $btnEstimate.Location='796,410'; $btnEstimate.Size=New-Object System.Drawing.Size(120,36); AnchorBR $btnEstimate; $form.Controls.Add($btnEstimate)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text="Cancel"; $btnCancel.Location='922,410'; $btnCancel.Size=New-Object System.Drawing.Size(120,36); $btnCancel.Enabled=$false; AnchorBR $btnCancel; $form.Controls.Add($btnCancel)

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text="Start"; $btnStart.Location='1048,410'; $btnStart.Size=New-Object System.Drawing.Size(120,36); AnchorBR $btnStart; $form.Controls.Add($btnStart)

# Cancel flag
$script:cancelRequested = $false

# ---------- Core ----------
function Validate-Roots {
  $s=$txtSrc.Text.Trim(); $d=$txtDst.Text.Trim()
  if(-not (Test-Path -LiteralPath $s)) { [System.Windows.Forms.MessageBox]::Show("Source root not found."); return $null }
  if(-not (Test-Path -LiteralPath $d)) { Ensure-Dir $d }
  @{ Source=$s; DestBase=$d }
}

$btnEstimate.Add_Click({
  $r=Validate-Roots; if(-not $r){ return }
  $patterns=($txtFilter.Text.Trim() -split ';' | ForEach-Object { if([string]::IsNullOrWhiteSpace($_)){'*'} else { $_ } })
  $total = Get-TotalCount $r.Source $patterns
  if ($total -le 0) { $lblStatus.Text="Status: no matching files"; return }

  if($optBySize.Checked){
    $filesPer=[int]$nudFilesPer.Value
    $batches=[int][Math]::Ceiling($total / [double]$filesPer)
  } else { $batches=[int]$nudBatches.Value }

  $approxPer = [int][Math]::Ceiling($total / [double]$batches)
  $lblStatus.Text="Status: estimated $total files; planned $batches batches (~$approxPer files per batch)"
  $progress.Minimum=0; $progress.Maximum=$total; $progress.Value=0
  $lblETA.Text="ETA: --:--"
})

$btnCancel.Add_Click({
  $script:cancelRequested = $true
  $btnCancel.Enabled = $false
  $lblStatus.Text = "Status: cancelling... finishing current file"
})

$btnStart.Add_Click({
  $btnStart.Enabled=$false
  $btnEstimate.Enabled=$false
  $btnCancel.Enabled=$true
  $script:cancelRequested = $false
  $cancelled = $false

  $r=Validate-Roots; if(-not $r){ $btnStart.Enabled=$true; $btnEstimate.Enabled=$true; $btnCancel.Enabled=$false; return }

  $prefix=$txtPrefix.Text.Trim(); if([string]::IsNullOrWhiteSpace($prefix)){ $prefix='batch_' }
  $patterns=($txtFilter.Text.Trim() -split ';' | ForEach-Object { if([string]::IsNullOrWhiteSpace($_)){'*'} else { $_ } })
  $move=$optMove.Checked
  $logp=$txtLog.Text.Trim()

  $total = Get-TotalCount $r.Source $patterns
  if($total -le 0){ [System.Windows.Forms.MessageBox]::Show("No matching files found."); $btnStart.Enabled=$true; $btnEstimate.Enabled=$true; $btnCancel.Enabled=$false; return }

  if($optBySize.Checked){
    $filesPer=[int]$nudFilesPer.Value
    $batches=[int][Math]::Ceiling($total / [double]$filesPer)
  } else {
    $batches=[int]$nudBatches.Value
  }
  if($batches -lt 1){ $batches=1 }

  for($i=1;$i -le $batches;$i++){ Ensure-Dir (BatchRootAt $r.DestBase $prefix $i) }

  $progress.Minimum=0; $progress.Maximum=$total; $progress.Value=0
  $lblStatus.Text="Status: starting..."
  $startTime=Get-Date; $done=0; $errs=0
  $rows = if([string]::IsNullOrWhiteSpace($logp)){ $null } else { New-Object System.Collections.Generic.List[object] }

  foreach ($p in $patterns) {
    if ($script:cancelRequested) { $cancelled = $true; break }
    foreach ($f in [System.IO.Directory]::EnumerateFiles($r.Source, $p, [System.IO.SearchOption]::AllDirectories)) {
      if ($script:cancelRequested) { $cancelled = $true; break }
      if(-not (Test-Path -LiteralPath $f)) { continue }

      $rel = RelPath $r.Source $f
      $bi  = BatchIndexFromRelPath $rel $batches
      $tgt = Join-Path (BatchRootAt $r.DestBase $prefix $bi) $rel
      Ensure-Dir (Split-Path $tgt -Parent)

      $act = if($move){"Move"} else {"Copy"}
      try{
        if($move){ Move-FileCompat -Source $f -Destination $tgt }
        else     { [System.IO.File]::Copy($f,$tgt,$true) }

        if($rows){
          $rows.Add([pscustomobject]@{Timestamp=Get-Date;Action=$act;Status='Done';Source=$f;Target=$tgt;Batch=$bi;Message='OK'}) | Out-Null
        }
      }catch{
        $errs++
        if($rows){
          $rows.Add([pscustomobject]@{Timestamp=Get-Date;Action=$act;Status='Error';Source=$f;Target=$tgt;Batch=$bi;Message=$_.Exception.Message}) | Out-Null
        }
      }

      $done++; if($done -le $progress.Maximum){ $progress.Value=$done }
      if($done % 200 -eq 0 -or $done -eq $total){
        $elapsed=(Get-Date)-$startTime
        $rate = if($elapsed.TotalSeconds -gt 0){ $done / $elapsed.TotalSeconds } else { 0 }
        $etaTs = if($rate -gt 0){ [TimeSpan]::FromSeconds( ($total-$done)/$rate ) } else { [TimeSpan]::Zero }
        $lblETA.Text="ETA: " + (Format-ETA $etaTs)
        $lblStatus.Text="Status: $done / $total processed | Errors: $errs | Batches: $batches"
        [System.Windows.Forms.Application]::DoEvents()
      }
    }
    if ($cancelled) { break }
  }

  if($rows){
    try{ $rows | Export-Csv -Path $logp -NoTypeInformation -Force }
    catch{ [System.Windows.Forms.MessageBox]::Show("Log write failed: " + $_.Exception.Message) }
  }

  if (-not $cancelled -and $move) {
    $lblStatus.Text="Status: cleaning up empty source folders..."
    [System.Windows.Forms.Application]::DoEvents()
    $dirs = [System.IO.Directory]::EnumerateDirectories($r.Source,'*',[System.IO.SearchOption]::AllDirectories) |
            Sort-Object { $_.Length } -Descending
    foreach($d in $dirs){
      try{
        $hasItems = [System.IO.Directory]::EnumerateFileSystemEntries($d) | Select-Object -First 1
        if(-not $hasItems){ [System.IO.Directory]::Delete($d, $false) }
      } catch { }
    }
    try{
      $rootHas = [System.IO.Directory]::EnumerateFileSystemEntries($r.Source) | Select-Object -First 1
      if(-not $rootHas){ [System.IO.Directory]::Delete($r.Source, $false) }
    } catch { }
  }

  if ($cancelled) {
    $lblStatus.Text="Status: CANCELLED - $done files processed, $errs errors"
    [System.Windows.Forms.MessageBox]::Show("Cancelled: $done files processed, $errs errors.","Cancelled",0,48) | Out-Null
  } else {
    $lblStatus.Text="Status: COMPLETE - $done files, $errs errors, $batches batches"
    $lblETA.Text="ETA: 00:00"
    [System.Windows.Forms.MessageBox]::Show("Complete: $done files, $errs errors across $batches batches.","Done",0,64) | Out-Null
  }

  $btnStart.Enabled=$true
  $btnEstimate.Enabled=$true
  $btnCancel.Enabled=$false
})

[void]$form.ShowDialog()
