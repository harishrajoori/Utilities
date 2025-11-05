Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Ensure-Dir([string]$d){ if(-not(Test-Path -LiteralPath $d)){[void][IO.Directory]::CreateDirectory($d)} }
function RelPath($root,$path){
  try{
    $r=(Resolve-Path $root).Path; if(-not $r.EndsWith('\')){$r+='\'}
    $u1=[uri]$r; $u2=[uri](Resolve-Path $path).Path
    return $u1.MakeRelativeUri($u2).ToString().Replace('/','\')
  }catch{
    $rn=[IO.Path]::GetFullPath((Resolve-Path $root).Path).TrimEnd('\')
    $pn=[IO.Path]::GetFullPath((Resolve-Path $path).Path)
    if($pn.StartsWith($rn,[StringComparison]::InvariantCultureIgnoreCase)){return $pn.Substring($rn.Length).TrimStart('\')}
    return Split-Path $pn -Leaf
  }
}

$form = New-Object Windows.Forms.Form
$form.Text="Batch Copy/Move (preserve structure)"
$form.StartPosition='CenterScreen'
$form.AutoScaleMode=[Windows.Forms.AutoScaleMode]::Font
$form.ClientSize=New-Object Drawing.Size(880,560)
$form.Font=New-Object Drawing.Font('Segoe UI',9)

function Add-Label($t,$x,$y){$l=New-Object Windows.Forms.Label;$l.Text=$t;$l.Location=New-Object Drawing.Point($x,$y);$l.AutoSize=$true;$form.Controls.Add($l);return $l}
function Add-TextBox($x,$y,$w){$t=New-Object Windows.Forms.TextBox;$t.Location=New-Object Drawing.Point($x,$y);$t.Size=New-Object Drawing.Size($w,24);$form.Controls.Add($t);return $t}
function Add-Button($txt,$x,$y,$w){$b=New-Object Windows.Forms.Button;$b.Text=$txt;$b.Location=New-Object Drawing.Point($x,$y);$b.Size=New-Object Drawing.Size($w,28);$form.Controls.Add($b);return $b}

$lblSrc=Add-Label "Source root:" 15 20
$txtSrc=Add-TextBox 120 18 640
$btnSrc=Add-Button "Browse..." 770 17 80
$btnSrc.Add_Click({$fbd=New-Object Windows.Forms.FolderBrowserDialog;if($fbd.ShowDialog() -eq 'OK'){$txtSrc.Text=$fbd.SelectedPath}})

$lblDst=Add-Label "Destination root:" 15 55
$txtDst=Add-TextBox 120 53 640
$btnDst=Add-Button "Browse..." 770 52 80
$btnDst.Add_Click({$fbd=New-Object Windows.Forms.FolderBrowserDialog;if($fbd.ShowDialog() -eq 'OK'){$txtDst.Text=$fbd.SelectedPath}})

$grp=New-Object Windows.Forms.GroupBox;$grp.Text="Action";$grp.Location=New-Object Drawing.Point(18,90);$grp.Size=New-Object Drawing.Size(250,60)
$optCopy=New-Object Windows.Forms.RadioButton;$optCopy.Text="Copy";$optCopy.Location=New-Object Drawing.Point(15,25);$optCopy.Checked=$true
$optMove=New-Object Windows.Forms.RadioButton;$optMove.Text="Move (free source space)";$optMove.Location=New-Object Drawing.Point(75,25)
$grp.Controls.AddRange(@($optCopy,$optMove));$form.Controls.Add($grp)

$lblBatch=Add-Label "Files per batch:" 285 110
$nudBatch=New-Object Windows.Forms.NumericUpDown;$nudBatch.Minimum=1;$nudBatch.Maximum=1000000;$nudBatch.Value=5000;$nudBatch.Location=New-Object Drawing.Point(380,107);$nudBatch.Size=New-Object Drawing.Size(100,24);$form.Controls.Add($nudBatch)
$chkDry=New-Object Windows.Forms.CheckBox;$chkDry.Text="Dry run (no changes)";$chkDry.Location=New-Object Drawing.Point(500,108);$chkDry.AutoSize=$true;$form.Controls.Add($chkDry)

$lblFilter=Add-Label "Optional file filter (e.g. *.pdf;*.docx):" 15 160
$txtFilter=Add-TextBox 260 157 505;$txtFilter.Text='*'

$lblLog=Add-Label "CSV log (optional):" 15 195
$txtLog=Add-TextBox 120 192 640
$btnLog=Add-Button "Choose..." 770 191 80
$btnLog.Add_Click({$dlg=New-Object Windows.Forms.SaveFileDialog;$dlg.Filter="CSV files (*.csv)|*.csv|All files (*.*)|*.*";if($dlg.ShowDialog() -eq 'OK'){$txtLog.Text=$dlg.FileName}})

$btnEst=Add-Button "Estimate count" 18 232 120
$btnRun=Add-Button "Process NEXT batch" 150 232 160
$btnCancel=Add-Button "Cancel" 320 232 120;$btnCancel.Enabled=$false

$txtOut=New-Object Windows.Forms.TextBox
$txtOut.Location=New-Object Drawing.Point(18,275)
$txtOut.Multiline=$true;$txtOut.ScrollBars='Vertical';$txtOut.ReadOnly=$true
$txtOut.Font=New-Object Drawing.Font('Consolas',9)
$txtOut.Size=New-Object Drawing.Size(840,260)
$form.Controls.Add($txtOut)
function Log($m){$txtOut.AppendText($m+[Environment]::NewLine);[Windows.Forms.Application]::DoEvents()}

function Validate-Roots{
  $s=$txtSrc.Text.Trim();$d=$txtDst.Text.Trim()
  if(-not(Test-Path -LiteralPath $s)){[Windows.Forms.MessageBox]::Show("Source root not found.");return $null}
  if(-not(Test-Path -LiteralPath $d)){Ensure-Dir $d}
  return @{Source=$s;Dest=$d}
}

$btnEst.Add_Click({
  $r=Validate-Roots;if(-not $r){return}
  $patterns=$txtFilter.Text.Trim() -split ';'
  $count=0
  foreach($p in $patterns){foreach($f in [IO.Directory]::EnumerateFiles($r.Source,$p,[IO.SearchOption]::AllDirectories)){$count++;if($count%1000-eq0){[Windows.Forms.Application]::DoEvents()}}}
  Log "Estimated files matching filter: $count"
})

$btnRun.Add_Click({
  $r=Validate-Roots;if(-not $r){return}
  $batch=[int]$nudBatch.Value;$move=$optMove.Checked;$dry=$chkDry.Checked
  $patterns=$txtFilter.Text.Trim() -split ';';$logp=$txtLog.Text.Trim()
  $rows=New-Object System.Collections.Generic.List[object];$processed=0;$errors=0
  $btnRun.Enabled=$false;$btnCancel.Enabled=$true
  foreach($p in $patterns){
    foreach($f in [IO.Directory]::EnumerateFiles($r.Source,$p,[IO.SearchOption]::AllDirectories)){
      if($processed -ge $batch){break}
      $rel=RelPath $r.Source $f
      $target=Join-Path $r.Dest $rel
      Ensure-Dir (Split-Path $target -Parent)
      $act=if($move){"Move"}else{"Copy"}
      try{
        if(-not $dry){
          if($move){[IO.File]::Move($f,$target,$false)}else{[IO.File]::Copy($f,$target,$true)}
        }
        $rows.Add([pscustomobject]@{Timestamp=Get-Date;Action=$act;Status='Done';Source=$f;Target=$target;Message='OK'})|Out-Null
        Log "[$act] $f -> $target"
        $processed++
      }catch{
        $errors++
        $rows.Add([pscustomobject]@{Timestamp=Get-Date;Action=$act;Status='Error';Source=$f;Target=$target;Message=$_.Exception.Message})|Out-Null
        Log "Error: $($_.Exception.Message)"
      }
      if($processed -ge $batch){break}
    }
    if($processed -ge $batch){break}
  }
  if($logp){try{$rows|Export-Csv -Path $logp -NoTypeInformation -Force}catch{Log "Log write failed: $($_.Exception.Message)"}}
  Log "Batch complete. Files: $processed, Errors: $errors"
  $btnRun.Enabled=$true;$btnCancel.Enabled=$false
})

[void]$form.ShowDialog()
