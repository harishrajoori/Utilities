# BatchCopyGUI_Sharded.ps1 — Split source into multiple batch roots with file-count caps
# PowerShell 5.1 compatible

# Relaunch in STA and keep console visible if started via "Run with PowerShell"
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
function Ensure-Dir([string]$d){
  if(-not(Test-Path -LiteralPath $d)){ [void][System.IO.Directory]::CreateDirectory($d) }
}
function RelPath($root,$path){
  try{
    $r=(Resolve-Path $root).Path; if(-not $r.EndsWith('\')){$r+='\'}
    $u1=[uri]$r; $u2=[uri](Resolve-Path $path).Path
    $rel=$u1.MakeRelativeUri($u2).ToString().Replace('/','\')
    return [System.Uri]::UnescapeDataString($rel)   # decode %20 etc.
  }catch{
    $rn=[System.IO.Path]::GetFullPath((Resolve-Path $root).Path).TrimEnd('\')
    $pn=[System.IO.Path]::GetFullPath((Resolve-Path $path).Path)
    if($pn.StartsWith($rn,[System.StringComparison]::InvariantCultureIgnoreCase)){ return $pn.Substring($rn.Length).TrimStart('\') }
    return Split-Path $pn -Leaf
  }
}
# PS 5.1: emulate overwrite-safe move
function Move-FileCompat {
  param([Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Destination)
  if (Test-Path -LiteralPath $Destination) { Remove-Item -LiteralPath $Destination -Force }
  try   { [System.IO.File]::Move($Source, $Destination) }   # 2-arg move in .NET Framework
  catch { [System.IO.File]::Copy($Source, $Destination, $true); Remove-Item -LiteralPath $Source -Force }
}
function Log($m){ $txtOut.AppendText($m+[Environment]::NewLine); [System.Windows.Forms.Application]::DoEvents() }

# ---------- UI ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Batch Copy/Move (preserve structure, sharded batches)"
$form.StartPosition = 'CenterScreen'
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
$form.ClientSize = New-Object System.Drawing.Size(980, 640)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

function Add-Label($t,$x,$y){$l=New-Object System.Windows.Forms.Label;$l.Text=$t;$l.Location=New-Object System.Drawing.Point($x,$y);$l.AutoSize=$true;$form.Controls.Add($l);$l}
function Add-TextBox($x,$y,$w){$t=New-Object System.Windows.Forms.TextBox;$t.Location=New-Object System.Drawing.Point($x,$y);$t.Size=New-Object System.Drawing.Size($w,24);$form.Controls.Add($t);$t}
function Add-Button($txt,$x,$y,$w){$b=New-Object System.Windows.Forms.Button;$b.Text=$txt;$b.Location=New-Object System.Drawing.Point($x,$y);$b.Size=New-Object System.Drawing.Size($w,28);$form.Controls.Add($b);$b}
function Add-NUD($x,$y,$min,$max,$val){$n=New-Object System.Windows.Forms.NumericUpDown;$n.Minimum=$min;$n.Maximum=$max;$n.Value=$val;$n.Location=New-Object System.Drawing.Point($x,$y);$n.Size=New-Object System.Drawing.Size(120,24);$form.Controls.Add($n);$n}

$lblSrc = Add-Label "Source root:" 15 20
$txtSrc = Add-TextBox 140 18 730
$btnSrc = Add-Button "Browse..." 880 17 80
$btnSrc.Add_Click({ $d=New-Object System.Windows.Forms.FolderBrowserDialog; if($d.ShowDialog() -eq 'OK'){ $txtSrc.Text=$d.SelectedPath; Reset-Pipeline } })

$lblDst = Add-Label "Destination BASE:" 15 55
$txtDst = Add-TextBox 140 53 730
$btnDst = Add-Button "Browse..." 880 52 80
$btnDst.Add_Click({ $d=New-Object System.Windows.Forms.FolderBrowserDialog; if($d.ShowDialog() -eq 'OK'){ $txtDst.Text=$d.SelectedPath } })

$grp = New-Object System.Windows.Forms.GroupBox
$grp.Text="Action"; $grp.Location=New-Object System.Drawing.Point(18, 90); $grp.Size=New-Object System.Drawing.Size(320, 64)
$optCopy = New-Object System.Windows.Forms.RadioButton; $optCopy.Text="Copy"; $optCopy.AutoSize=$true; $optCopy.Location=New-Object System.Drawing.Point(15, 25); $optCopy.Checked=$true
$optMove = New-Object System.Windows.Forms.RadioButton; $optMove.Text="Move"; $optMove.AutoSize=$true; $optMove.Location=New-Object System.Drawing.Point(85, 25)
$grp.Controls.AddRange(@($optCopy,$optMove)); $form.Controls.Add($grp)

$lblFilesPer = Add-Label "Files per batch (cap):" 360 110
$nudFilesPer = Add-NUD 510 107 1 100000000 20000

$lblNumBatches = Add-Label "Number of batches:" 660 110
$nudNumBatches = Add-NUD 790 107 1 100000 5

$lblPrefix = Add-Label "Batch folder prefix:" 15 160
$txtPrefix = Add-TextBox 140 157 160; $txtPrefix.Text = "batch_"

$chkDry = New-Object System.Windows.Forms.CheckBox
$chkDry.Text="Dry run (no changes)"; $chkDry.AutoSize=$true; $chkDry.Location=New-Object System.Drawing.Point(320, 158)
$form.Controls.Add($chkDry)

$lblFilter = Add-Label "Optional file filter (e.g. *.pdf;*.jpg;*.docx):" 15 195
$txtFilter = Add-TextBox 300 193 660; $txtFilter.Text='*'
$txtFilter.Add_TextChanged({ Reset-Pipeline })

$lblLog = Add-Label "CSV log (optional):" 15 230
$txtLog = Add-TextBox 140 228 730
$btnLog = Add-Button "Choose..." 880 227 80
$btnLog.Add_Click({ $dlg=New-Object System.Windows.Forms.SaveFileDialog; $dlg.Filter="CSV files (*.csv)|*.csv|All files (*.*)|*.*"; if($dlg.ShowDialog() -eq 'OK'){ $txtLog.Text=$dlg.FileName } })

$btnEstimate = Add-Button "Estimate count"      18  265 120
$btnStart    = Add-Button "Start (fill batch 1)" 148 265 160
$btnNext     = Add-Button "Process NEXT batch"  313 265 180
$btnReset    = Add-Button "Reset pipeline"      498 265 140

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location=New-Object System.Drawing.Point(18, 305)
$txtOut.Multiline=$true; $txtOut.ScrollBars='Vertical'; $txtOut.ReadOnly=$true
$txtOut.Font=New-Object System.Drawing.Font('Consolas', 9)
$txtOut.Size=New-Object System.Drawing.Size(942, 300)
$form.Controls.Add($txtOut)

# ---------- Sharding state: single, flat, lazy enumerator with peek ----------
$global:fileEnum   = $null   # IEnumerator[string]
$global:hasPeek    = $false
$global:peekValue  = $null
$global:currentBatchIndex = 0
$global:filesInCurrentBatch = 0

function Reset-Pipeline {
  $global:fileEnum  = $null
  $global:hasPeek   = $false
  $global:peekValue = $null
  $global:currentBatchIndex = 0
  $global:filesInCurrentBatch = 0
  ToggleButtons -state 'idle'
}

function Build-Enumerator {
  $src=$txtSrc.Text.Trim()
  $patterns = ($txtFilter.Text.Trim() -split ';' | ForEach-Object { if([string]::IsNullOrWhiteSpace($_)){'*'} else { $_ } })

  $sequence = foreach($p in $patterns) {
    foreach($f in [System.IO.Directory]::EnumerateFiles($src,$p,[System.IO.SearchOption]::AllDirectories)) { $f }
  }

  $global:fileEnum = $sequence.GetEnumerator()
  $global:hasPeek  = $global:fileEnum.MoveNext()
  if($global:hasPeek){ $global:peekValue = $global:fileEnum.Current }
}

function Has-NextItem {
  if(-not $global:fileEnum){ Build-Enumerator }
  return $global:hasPeek
}
function Next-Item {
  if(-not $global:hasPeek){ return $null }
  $c=$global:peekValue
  $global:hasPeek = $global:fileEnum.MoveNext()
  if($global:hasPeek){ $global:peekValue = $global:fileEnum.Current } else { $global:peekValue = $null }
  return $c
}

function Validate-Roots {
  $s=$txtSrc.Text.Trim(); $d=$txtDst.Text.Trim()
  if(-not (Test-Path -LiteralPath $s)) { [System.Windows.Forms.MessageBox]::Show("Source root not found."); return $null }
  if(-not (Test-Path -LiteralPath $d)) { Ensure-Dir $d }
  @{ Source=$s; DestBase=$d }
}

function ToggleButtons([string]$state) {
  switch ($state) {
    'idle'   { $btnStart.Enabled=$true;  $btnNext.Enabled=$false }
    'started'{ $btnStart.Enabled=$false; $btnNext.Enabled=$true  }
    'done'   { $btnStart.Enabled=$false; $btnNext.Enabled=$false }
  }
}

function Get-CurrentBatchRoot {
  param([string]$destBase, [string]$prefix, [int]$batchIndex)
  $name = ("{0}{1:D2}" -f $prefix, $batchIndex)
  return (Join-Path $destBase $name)
}

# ---------- Actions ----------
$btnEstimate.Add_Click({
  $r=Validate-Roots; if(-not $r){ return }
  $patterns=$txtFilter.Text.Trim() -split ';'
  $count=0
  foreach($p in $patterns){
    foreach($f in [System.IO.Directory]::EnumerateFiles($r.Source,$p,[System.IO.SearchOption]::AllDirectories)){
      $count++; if($count%5000 -eq 0){ [System.Windows.Forms.Application]::DoEvents() }
    }
  }
  $cap=[int]$nudFilesPer.Value; $nb=[int]$nudNumBatches.Value
  $maxCap = $cap * $nb
  Log ("Estimated files matching filter: {0} | Planned capacity: {1} ({2} x {3})" -f $count,$maxCap,$nb,$cap)
})

$btnReset.Add_Click({ Reset-Pipeline; Log "Pipeline reset. Next batch will scan from the beginning." })

function Run-Batch {
  param([switch]$IsFirst)

  $r=Validate-Roots; if(-not $r){ return }
  $filesPerBatch = [int]$nudFilesPer.Value
  $numBatches    = [int]$nudNumBatches.Value
  $prefix        = $txtPrefix.Text.Trim()
  if([string]::IsNullOrWhiteSpace($prefix)){ $prefix = "batch_" }

  if($global:currentBatchIndex -eq 0){
    $global:currentBatchIndex = 1
    $global:filesInCurrentBatch = 0
  }

  $move=$optMove.Checked; $dry=$chkDry.Checked
  $logp=$txtLog.Text.Trim()
  $rows=New-Object System.Collections.Generic.List[object]
  $placed=0; $errors=0

  # Ensure base and current batch dir
  $batchRoot = Get-CurrentBatchRoot -destBase $r.DestBase -prefix $prefix -batchIndex $global:currentBatchIndex
  Ensure-Dir $batchRoot

  while( (Has-NextItem) -and ($global:currentBatchIndex -le $numBatches) ){
    if($global:filesInCurrentBatch -ge $filesPerBatch){
      # Move to next batch root
      $global:currentBatchIndex++
      if($global:currentBatchIndex -gt $numBatches){ break }
      $global:filesInCurrentBatch = 0
      $batchRoot = Get-CurrentBatchRoot -destBase $r.DestBase -prefix $prefix -batchIndex $global:currentBatchIndex
      Ensure-Dir $batchRoot
      if($IsFirst){ ToggleButtons -state 'started' }
      Log ("--- Switched to {0}" -f $batchRoot)
    }

    $f=Next-Item; if(-not $f){ break }
    if(-not (Test-Path -LiteralPath $f)) { continue }

    $rel = RelPath $r.Source $f
    $tgt = Join-Path $batchRoot $rel
    Ensure-Dir (Split-Path $tgt -Parent)

    $act = if($move){"Move"}else{"Copy"}
    try{
      if(-not $dry){
        if($move){ Move-FileCompat -Source $f -Destination $tgt }
        else     { [System.IO.File]::Copy($f, $tgt, $true) }
      }
      Log "[$act] $f -> $tgt"
      $rows.Add([pscustomobject]@{Timestamp=Get-Date;Action=$act;Status='Done';Source=$f;Target=$tgt;Batch=$global:currentBatchIndex;Message='OK'}) | Out-Null
      $global:filesInCurrentBatch++
      $placed++
    } catch {
      $errors++
      $rows.Add([pscustomobject]@{Timestamp=Get-Date;Action=$act;Status='Error';Source=$f;Target=$tgt;Batch=$global:currentBatchIndex;Message=$_.Exception.Message}) | Out-Null
      Log "Error: $($_.Exception.Message)"
    }

    # If this call was the initial Start, we still only fill ONE batch per call.
    if ($IsFirst -and $global:filesInCurrentBatch -ge $filesPerBatch) { break }
  }

  if(-not (Has-NextItem) -or $global:currentBatchIndex -gt $numBatches){
    Log ("No more files to process or batches filled. Last batch #{0} had {1} files this run." -f $global:currentBatchIndex,$global:filesInCurrentBatch)
    ToggleButtons -state 'done'
  } else {
    Log ("Batch #{0} status: {1}/{2} files | This run placed: {3}, Errors: {4}" -f $global:currentBatchIndex, $global:filesInCurrentBatch, $filesPerBatch, $placed, $errors)
    if ($IsFirst) { ToggleButtons -state 'started' }
  }

  if ($logp) {
    try { $rows | Export-Csv -Path $logp -NoTypeInformation -Append:$(Test-Path $logp) -Force }
    catch { Log "Log write failed: $($_.Exception.Message)" }
  }
}

$btnStart.Add_Click({ Run-Batch -IsFirst })   # fills batch #1 (up to cap)
$btnNext.Add_Click({  Run-Batch })            # fills next batch, and the next…

# Estimate count (no batching changes)
$btnEstimate.Add_Click({
  $r=Validate-Roots; if(-not $r){ return }
  $patterns=$txtFilter.Text.Trim() -split ';'
  $count=0
  foreach($p in $patterns){
    foreach($f in [System.IO.Directory]::EnumerateFiles($r.Source,$p,[System.IO.SearchOption]::AllDirectories)){
      $count++; if($count%5000 -eq 0){ [System.Windows.Forms.Application]::DoEvents() }
    }
  }
  $cap=[int]$nudFilesPer.Value; $nb=[int]$nudNumBatches.Value
  $maxCap = $cap * $nb
  Log ("Estimated files matching filter: {0} | Planned capacity: {1} ({2} x {3})" -f $count,$maxCap,$nb,$cap)
})

# initial state
Reset-Pipeline
[void]$form.ShowDialog()
