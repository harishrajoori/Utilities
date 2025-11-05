# BatchCopyGUI_Final3.ps1  — PowerShell 5.1 compatible

# --- Relaunch in STA and keep console open if started via "Run with PowerShell" ---
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
      "Error",
      0, 16
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
function RelPath($root,$path){
  try{
    $r=(Resolve-Path $root).Path
    if(-not $r.EndsWith('\')){$r+='\'}
    $u1=[uri]$r; $u2=[uri](Resolve-Path $path).Path
    $rel=$u1.MakeRelativeUri($u2).ToString().Replace('/','\')
    return [System.Uri]::UnescapeDataString($rel)   # fix %20 etc.
  }catch{
    $rn=[System.IO.Path]::GetFullPath((Resolve-Path $root).Path).TrimEnd('\')
    $pn=[System.IO.Path]::GetFullPath((Resolve-Path $path).Path)
    if($pn.StartsWith($rn,[System.StringComparison]::InvariantCultureIgnoreCase)){
      return $pn.Substring($rn.Length).TrimStart('\')
    }
    return Split-Path $pn -Leaf
  }
}
# PS 5.1/.NET Framework: File.Move has only 2 args, so emulate "overwrite-safe" move
function Move-FileCompat {
  param([Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Destination)

  if (Test-Path -LiteralPath $Destination) {
    Remove-Item -LiteralPath $Destination -Force
  }
  try {
    [System.IO.File]::Move($Source, $Destination)   # 2-arg move on .NET Framework
  } catch {
    # Fallback: copy-then-delete (handles cross-volume or sharing quirks)
    [System.IO.File]::Copy($Source, $Destination, $true)
    Remove-Item -LiteralPath $Source -Force
  }
}

# ---------- UI ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Batch Copy/Move (preserve structure)"
$form.StartPosition = 'CenterScreen'
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
$form.ClientSize = New-Object System.Drawing.Size(900, 600)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

function Add-Label($t,$x,$y){$l=New-Object System.Windows.Forms.Label;$l.Text=$t;$l.Location=New-Object System.Drawing.Point($x,$y);$l.AutoSize=$true;$form.Controls.Add($l);$l}
function Add-TextBox($x,$y,$w){$t=New-Object System.Windows.Forms.TextBox;$t.Location=New-Object System.Drawing.Point($x,$y);$t.Size=New-Object System.Drawing.Size($w,24);$form.Controls.Add($t);$t}
function Add-Button($txt,$x,$y,$w){$b=New-Object System.Windows.Forms.Button;$b.Text=$txt;$b.Location=New-Object System.Drawing.Point($x,$y);$b.Size=New-Object System.Drawing.Size($w,28);$form.Controls.Add($b);$b}

$lblSrc = Add-Label "Source root:" 15 20
$txtSrc = Add-TextBox 130 18 660
$btnSrc = Add-Button "Browse..." 805 17 80
$btnSrc.Add_Click({
  $d=New-Object System.Windows.Forms.FolderBrowserDialog
  if($d.ShowDialog() -eq 'OK'){ $txtSrc.Text=$d.SelectedPath; $global:enumerators=$null; ToggleButtons -state 'idle' }
})

$lblDst = Add-Label "Destination root:" 15 55
$txtDst = Add-TextBox 130 53 660
$btnDst = Add-Button "Browse..." 805 52 80
$btnDst.Add_Click({
  $d=New-Object System.Windows.Forms.FolderBrowserDialog
  if($d.ShowDialog() -eq 'OK'){ $txtDst.Text=$d.SelectedPath }
})

$grp = New-Object System.Windows.Forms.GroupBox
$grp.Text="Action"; $grp.Location=New-Object System.Drawing.Point(18, 90); $grp.Size=New-Object System.Drawing.Size(320, 64)
$optCopy = New-Object System.Windows.Forms.RadioButton; $optCopy.Text="Copy"; $optCopy.AutoSize=$true; $optCopy.Location=New-Object System.Drawing.Point(15, 25); $optCopy.Checked=$true
$optMove = New-Object System.Windows.Forms.RadioButton; $optMove.Text="Move"; $optMove.AutoSize=$true; $optMove.Location=New-Object System.Drawing.Point(85, 25)
$grp.Controls.AddRange(@($optCopy,$optMove)); $form.Controls.Add($grp)

$lblBatch = Add-Label "Files per batch:" 360 110
$nudBatch = New-Object System.Windows.Forms.NumericUpDown
$nudBatch.Minimum=1; $nudBatch.Maximum=1000000; $nudBatch.Value=5000
$nudBatch.Location=New-Object System.Drawing.Point(455, 107); $nudBatch.Size=New-Object System.Drawing.Size(100,24)
$form.Controls.Add($nudBatch)

$chkDry = New-Object System.Windows.Forms.CheckBox
$chkDry.Text="Dry run (no changes)"; $chkDry.AutoSize=$true; $chkDry.Location=New-Object System.Drawing.Point(575, 108)
$form.Controls.Add($chkDry)

$lblFilter = Add-Label "Optional file filter (e.g. *.pdf;*.docx):" 15 165
$txtFilter = Add-TextBox 275 163 530; $txtFilter.Text='*'
$txtFilter.Add_TextChanged({ $global:enumerators=$null; ToggleButtons -state 'idle' })

$lblLog = Add-Label "CSV log (optional):" 15 200
$txtLog = Add-TextBox 130 198 660
$btnLog = Add-Button "Choose..." 805 197 80
$btnLog.Add_Click({
  $dlg=New-Object System.Windows.Forms.SaveFileDialog
  $dlg.Filter="CSV files (*.csv)|*.csv|All files (*.*)|*.*"
  if($dlg.ShowDialog() -eq 'OK'){ $txtLog.Text=$dlg.FileName }
})

$btnEstimate = Add-Button "Estimate count"      18  240 120
$btnStart    = Add-Button "Start"              148  240 100
$btnNext     = Add-Button "Process NEXT batch" 255  240 180
$btnReset    = Add-Button "Reset pipeline"     442  240 120

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location=New-Object System.Drawing.Point(18, 285)
$txtOut.Multiline=$true; $txtOut.ScrollBars='Vertical'; $txtOut.ReadOnly=$true
$txtOut.Font=New-Object System.Drawing.Font('Consolas', 9)
$txtOut.Size=New-Object System.Drawing.Size(867, 300)
$form.Controls.Add($txtOut)

function Log($m){ $txtOut.AppendText($m+[Environment]::NewLine); [System.Windows.Forms.Application]::DoEvents() }

# ---------- Batching state ----------
$global:enumerators = $null   # array of IEnumerator[string]
$global:activeIndex = 0

function Build-Enumerators {
  $src=$txtSrc.Text.Trim()
  $patterns=($txtFilter.Text.Trim() -split ';' | ForEach-Object { if([string]::IsNullOrWhiteSpace($_)){'*'} else { $_ } })
  $list = New-Object System.Collections.Generic.List[System.Collections.Generic.IEnumerator[string]]
  foreach($p in $patterns){
    $e = [System.IO.Directory]::EnumerateFiles($src,$p,[System.IO.SearchOption]::AllDirectories).GetEnumerator()
    $null = $e.MoveNext()    # prime
    $list.Add($e) | Out-Null
  }
  $global:enumerators = $list.ToArray()
  $global:activeIndex = 0
}

function Has-NextItem {
  if(-not $global:enumerators -or $global:enumerators.Count -eq 0){ return $false }
  for($i=$global:activeIndex; $i -lt $global:enumerators.Count; $i++){
    $e = $global:enumerators[$i]
    if($e -and $e.Current){ $global:activeIndex=$i; return $true }
    while($e -and $e.MoveNext()){ if($e.Current){ $global:activeIndex=$i; return $true } }
  }
  return $false
}
function Next-Item {
  $i=$global:activeIndex; $e=$global:enumerators[$i]
  $c=$e.Current
  if(-not $e.MoveNext()){ $global:activeIndex=$i+1 }
  return $c
}

function Validate-Roots {
  $s=$txtSrc.Text.Trim(); $d=$txtDst.Text.Trim()
  if(-not (Test-Path -LiteralPath $s)) { [System.Windows.Forms.MessageBox]::Show("Source root not found."); return $null }
  if(-not (Test-Path -LiteralPath $d)) { Ensure-Dir $d }
  @{ Source=$s; Dest=$d }
}

function ToggleButtons([string]$state) {
  switch ($state) {
    'idle'   { $btnStart.Enabled=$true;  $btnNext.Enabled=$false }
    'started'{ $btnStart.Enabled=$false; $btnNext.Enabled=$true  }
    'done'   { $btnStart.Enabled=$false; $btnNext.Enabled=$false }
  }
}

# ---------- Actions ----------
$btnEstimate.Add_Click({
  $r=Validate-Roots; if(-not $r){ return }
  $patterns=$txtFilter.Text.Trim() -split ';'
  $count=0
  foreach($p in $patterns){
    foreach($f in [System.IO.Directory]::EnumerateFiles($r.Source,$p,[System.IO.SearchOption]::AllDirectories)){
      $count++; if($count%2000 -eq 0){ [System.Windows.Forms.Application]::DoEvents() }
    }
  }
  Log "Estimated files matching filter: $count"
})

$btnReset.Add_Click({
  $global:enumerators=$null
  ToggleButtons -state 'idle'
  Log "Pipeline reset. Next batch will scan from the beginning."
})

function Run-Batch {
  param([switch]$IsFirst)
  $r=Validate-Roots; if(-not $r){ return }
  if(-not $global:enumerators){ Build-Enumerators }

  $batch=[int]$nudBatch.Value; $move=$optMove.Checked; $dry=$chkDry.Checked
  $logp=$txtLog.Text.Trim()
  $rows=New-Object System.Collections.Generic.List[object]
  $processed=0; $errors=0

  while($processed -lt $batch -and (Has-NextItem)) {
    $f=Next-Item; if(-not $f){ break }
    if(-not (Test-Path -LiteralPath $f)) { continue }

    $rel = RelPath $r.Source $f
    $tgt = Join-Path $r.Dest $rel
    Ensure-Dir (Split-Path $tgt -Parent)

    $act = if($move){"Move"}else{"Copy"}
    try{
      if(-not $dry){
        if($move){ Move-FileCompat -Source $f -Destination $tgt }
        else     { [System.IO.File]::Copy($f, $tgt, $true) }
      }
      Log "[$act] $f -> $tgt"
      $rows.Add([pscustomobject]@{Timestamp=Get-Date;Action=$act;Status='Done';Source=$f;Target=$tgt;Message='OK'}) | Out-Null
      $processed++
    } catch {
      $errors++
      $rows.Add([pscustomobject]@{Timestamp=Get-Date;Action=$act;Status='Error';Source=$f;Target=$tgt;Message=$_.Exception.Message}) | Out-Null
      Log "Error: $($_.Exception.Message)"
    }
  }

  if ($processed -eq 0 -and -not (Has-NextItem)) {
    Log "No more files to process."
    ToggleButtons -state 'done'
  } else {
    Log "Batch complete. Files: $processed, Errors: $errors"
    if ($IsFirst) { ToggleButtons -state 'started' }
  }

  if ($logp) {
    try { $rows | Export-Csv -Path $logp -NoTypeInformation -Append:$(Test-Path $logp) -Force }
    catch { Log "Log write failed: $($_.Exception.Message)" }
  }
}

$btnStart.Add_Click({ Run-Batch -IsFirst })
$btnNext.Add_Click({ Run-Batch })

# initial state
ToggleButtons -state 'idle'

[void]$form.ShowDialog()
