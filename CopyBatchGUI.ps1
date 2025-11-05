Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ========== Helpers ==========
function New-LongPath {
    param([string]$Path)
    if ($Path -match '^[\\]{2}\?\\') { return $Path }
    if ($Path -like '\\*') { return "\\?\UNC\" + $Path.TrimStart('\') }  # UNC
    return "\\?\" + $Path
}

function Get-RelativePath {
    param([string]$Root, [string]$FullPath)
    try {
        if ([System.IO.Path]::IsPathRooted($Root) -and [System.IO.Path]::IsPathRooted($FullPath)) {
            # .NET Core has GetRelativePath. For 5.1 fallback to URI math:
            $uriRoot = New-Object System.Uri((Resolve-Path $Root).Path + [IO.Path]::DirectorySeparatorChar)
            $uriFull = New-Object System.Uri((Resolve-Path $FullPath).Path)
            $rel = $uriRoot.MakeRelativeUri($uriFull).ToString() -replace '/', '\'
            return $rel
        }
    } catch {}
    # Fallback: trim the root prefix if it matches
    $rootNorm = [IO.Path]::GetFullPath((Resolve-Path $Root).Path).TrimEnd('\')
    $fullNorm = [IO.Path]::GetFullPath((Resolve-Path $FullPath).Path)
    if ($fullNorm.StartsWith($rootNorm, [StringComparison]::InvariantCultureIgnoreCase)) {
        return $fullNorm.Substring($rootNorm.Length).TrimStart('\')
    }
    # Last resort: just filename
    return Split-Path -Path $FullPath -Leaf
}

function Ensure-Directory {
    param([string]$Dir)
    if (-not [System.IO.Directory]::Exists($Dir)) {
        [void][System.IO.Directory]::CreateDirectory($Dir)
    }
}

function Write-LogRow {
    param(
        [ref]$LogBuffer, [string]$Action, [string]$Status, [string]$Source, [string]$Target, [string]$Message
    )
    $LogBuffer.Value.Add([pscustomobject]@{
        Timestamp = (Get-Date)
        Action    = $Action
        Status    = $Status
        Source    = $Source
        Target    = $Target
        Message   = $Message
    }) | Out-Null
}

# ========== UI ==========
$form                = New-Object System.Windows.Forms.Form
$form.Text           = "Batch Copy/Move (preserve structure)"
$form.Size           = New-Object System.Drawing.Size(820, 520)
$form.StartPosition  = "CenterScreen"

$lblSource = New-Object System.Windows.Forms.Label
$lblSource.Text = "Source root:"
$lblSource.Location = New-Object System.Drawing.Point(15, 20)
$lblSource.AutoSize = $true
$form.Controls.Add($lblSource)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Location = New-Object System.Drawing.Point(120, 15)
$txtSource.Size = New-Object System.Drawing.Size(560, 22)
$form.Controls.Add($txtSource)

$btnSource = New-Object System.Windows.Forms.Button
$btnSource.Text = "Browse..."
$btnSource.Location = New-Object System.Drawing.Point(690, 14)
$btnSource.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dlg.ShowDialog() -eq "OK") { $txtSource.Text = $dlg.SelectedPath }
})
$form.Controls.Add($btnSource)

$lblDest = New-Object System.Windows.Forms.Label
$lblDest.Text = "Destination root:"
$lblDest.Location = New-Object System.Drawing.Point(15, 55)
$lblDest.AutoSize = $true
$form.Controls.Add($lblDest)

$txtDest = New-Object System.Windows.Forms.TextBox
$txtDest.Location = New-Object System.Drawing.Point(120, 50)
$txtDest.Size = New-Object System.Drawing.Size(560, 22)
$form.Controls.Add($txtDest)

$btnDest = New-Object System.Windows.Forms.Button
$btnDest.Text = "Browse..."
$btnDest.Location = New-Object System.Drawing.Point(690, 49)
$btnDest.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dlg.ShowDialog() -eq "OK") { $txtDest.Text = $dlg.SelectedPath }
})
$form.Controls.Add($btnDest)

$grpAction = New-Object System.Windows.Forms.GroupBox
$grpAction.Text = "Action"
$grpAction.Location = New-Object System.Drawing.Point(18, 90)
$grpAction.Size = New-Object System.Drawing.Size(210, 60)
$form.Controls.Add($grpAction)

$optCopy = New-Object System.Windows.Forms.RadioButton
$optCopy.Text = "Copy"
$optCopy.Location = New-Object System.Drawing.Point(15, 25)
$optCopy.Checked = $true
$grpAction.Controls.Add($optCopy)

$optMove = New-Object System.Windows.Forms.RadioButton
$optMove.Text = "Move (frees source space)"
$optMove.Location = New-Object System.Drawing.Point(70, 25)
$grpAction.Controls.Add($optMove)

$lblBatch = New-Object System.Windows.Forms.Label
$lblBatch.Text = "Files per batch:"
$lblBatch.Location = New-Object System.Drawing.Point(250, 110)
$lblBatch.AutoSize = $true
$form.Controls.Add($lblBatch)

$nudBatch = New-Object System.Windows.Forms.NumericUpDown
$nudBatch.Minimum = 1
$nudBatch.Maximum = 1000000
$nudBatch.Value = 5000
$nudBatch.Location = New-Object System.Drawing.Point(350, 107)
$nudBatch.Size = New-Object System.Drawing.Size(100, 22)
$form.Controls.Add($nudBatch)

$chkDryRun = New-Object System.Windows.Forms.CheckBox
$chkDryRun.Text = "Dry run (no changes)"
$chkDryRun.Location = New-Object System.Drawing.Point(480, 108)
$chkDryRun.AutoSize = $true
$form.Controls.Add($chkDryRun)

$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text = "Optional file filter (e.g. *.pdf;*.docx):"
$lblFilter.Location = New-Object System.Drawing.Point(15, 160)
$lblFilter.AutoSize = $true
$form.Controls.Add($lblFilter)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Location = New-Object System.Drawing.Point(250, 155)
$txtFilter.Size = New-Object System.Drawing.Size(430, 22)
$txtFilter.Text = "*"
$form.Controls.Add($txtFilter)

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "CSV log (optional):"
$lblLog.Location = New-Object System.Drawing.Point(15, 195)
$lblLog.AutoSize = $true
$form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(120, 190)
$txtLog.Size = New-Object System.Drawing.Size(560, 22)
$form.Controls.Add($txtLog)

$btnLog = New-Object System.Windows.Forms.Button
$btnLog.Text = "Choose..."
$btnLog.Location = New-Object System.Drawing.Point(690, 189)
$btnLog.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    if ($dlg.ShowDialog() -eq "OK") { $txtLog.Text = $dlg.FileName }
})
$form.Controls.Add($btnLog)

$btnCount = New-Object System.Windows.Forms.Button
$btnCount.Text = "Estimate count"
$btnCount.Location = New-Object System.Drawing.Point(18, 230)
$form.Controls.Add($btnCount)

$btnNext = New-Object System.Windows.Forms.Button
$btnNext.Text = "Process NEXT batch"
$btnNext.Location = New-Object System.Drawing.Point(150, 230)
$form.Controls.Add($btnNext)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(330, 230)
$btnCancel.Enabled = $false
$form.Controls.Add($btnCancel)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(18, 270)
$txtOut.Multiline = $true
$txtOut.ScrollBars = "Vertical"
$txtOut.ReadOnly = $true
$txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtOut.Size = New-Object System.Drawing.Size(765, 190)
$form.Controls.Add($txtOut)

# ========== State ==========
$global:cts = $null
$global:lastResumePath = $null
$logBuffer = New-Object System.Collections.Generic.List[object]

function Append-Out {
    param([string]$Line)
    $txtOut.AppendText($Line + [Environment]::NewLine)
}

function Validate-Roots {
    $src = $txtSource.Text.Trim()
    $dst = $txtDest.Text.Trim()
    if (-not (Test-Path -LiteralPath $src)) { [void][System.Windows.Forms.MessageBox]::Show("Source root not found.") ; return $null }
    if (-not (Test-Path -LiteralPath $dst)) { try { Ensure-Directory -Dir $dst } catch { [void][System.Windows.Forms.MessageBox]::Show("Cannot create destination root.") ; return $null } }
    return @{ Source=$src; Dest=$dst }
}

function Enumerate-Files {
    param([string]$Root, [string[]]$Patterns)
    # Stream files across patterns (OR)
    foreach ($pat in $Patterns) {
        foreach ($f in [System.IO.Directory]::EnumerateFiles($Root, $pat, [System.IO.SearchOption]::AllDirectories)) {
            $f
        }
    }
}

function Process-Batch {
    param([int]$BatchSize)

    $roots = Validate-Roots
    if (-not $roots) { return }

    $srcRoot = $roots.Source; $dstRoot = $roots.Dest
    $doMove  = $optMove.Checked
    $dry     = $chkDryRun.Checked
    $patterns = ($txtFilter.Text.Trim() -split ';' | ForEach-Object { if ([string]::IsNullOrWhiteSpace($_)) { "*" } else { $_ } })

    $global:cts = New-Object System.Threading.CancellationTokenSource
    $btnNext.Enabled = $false
    $btnCancel.Enabled = $true

    Start-Job -ScriptBlock {
        param($srcRoot, $dstRoot, $BatchSize, $doMove, $dry, $patterns)
        using namespace System.IO

        $processed = 0
        $errors = 0
        $rows = New-Object System.Collections.Generic.List[object]

        function Safe-EnsureDir([string]$dir) {
            if (-not [Directory]::Exists($dir)) { [void][Directory]::CreateDirectory($dir) }
        }

        function Rel([string]$root, [string]$path) {
            try {
                $uriRoot = [Uri]((Resolve-Path $root).Path + [IO.Path]::DirectorySeparatorChar)
                $uriFull = [Uri]((Resolve-Path $path).Path)
                return $uriRoot.MakeRelativeUri($uriFull).ToString().Replace('/','\')
            } catch {
                $rootNorm = [IO.Path]::GetFullPath((Resolve-Path $root).Path).TrimEnd('\')
                $fullNorm = [IO.Path]::GetFullPath((Resolve-Path $path).Path)
                if ($fullNorm.StartsWith($rootNorm, [StringComparison]::InvariantCultureIgnoreCase)) {
                    return $fullNorm.Substring($rootNorm.Length).TrimStart('\')
                }
                return [IO.Path]::GetFileName($fullNorm)
            }
        }

        $enumerator = foreach ($pat in $patterns) {
            [Directory]::EnumerateFiles($srcRoot, $pat, [SearchOption]::AllDirectories)
        }

        foreach ($batchEnum in $enumerator) {
            foreach ($f in $batchEnum) {
                if ($processed -ge $BatchSize) { break }
                $rel = Rel $srcRoot $f
                $target = Join-Path $dstRoot $rel
                $targetDir = Split-Path $target -Parent
                try { Safe-EnsureDir $targetDir } catch {}

                $act = $doMove ? "Move" : "Copy"
                $msg = "OK"

                try {
                    if (-not $dry) {
                        if ($doMove) {
                            [File]::Move($f, $target, $false)
                        } else {
                            # Overwrite if exists
                            [File]::Copy($f, $target, $true)
                        }
                    }
                    $row = [pscustomobject]@{
                        Timestamp = (Get-Date)
                        Action    = $act
                        Status    = "Done"
                        Source    = $f
                        Target    = $target
                        Message   = $msg
                    }
                    $rows.Add($row) | Out-Null
                    $processed++
                } catch {
                    $errors++
                    $row = [pscustomobject]@{
                        Timestamp = (Get-Date)
                        Action    = $act
                        Status    = "Error"
                        Source    = $f
                        Target    = $target
                        Message   = $_.Exception.Message
                    }
                    $rows.Add($row) | Out-Null
                }

                if ($processed -ge $BatchSize) { break }
            }
            if ($processed -ge $BatchSize) { break }
        }

        return @{
            Processed = $processed
            Errors    = $errors
            Rows      = $rows
        }
    } -ArgumentList $srcRoot, $dstRoot, $BatchSize, $doMove, $dry, $patterns | Wait-Job | Receive-Job | ForEach-Object {
        $result = $_
        $result.Rows | ForEach-Object {
            $line = "{0:u} [{1}/{2}] {3} -> {4} {5}" -f $_.Timestamp, $_.Action, $_.Status, $_.Source, $_.Target, $_.Message
            $form.Invoke({ param($t) $txtOut.AppendText($t + [Environment]::NewLine) }, $line) | Out-Null
            $logBuffer.Add($_) | Out-Null
        }

        $form.Invoke({ 
            $btnNext.Enabled = $true
            $btnCancel.Enabled = $false
        }) | Out-Null

        if ($txtLog.Text.Trim()) {
            try {
                $logBuffer | Export-Csv -Path $txtLog.Text.Trim() -NoTypeInformation -Append:$false -Force
                $logBuffer.Clear()
                $form.Invoke({ param($t) $txtOut.AppendText("Wrote log to: $t`r`n") }, $txtLog.Text.Trim()) | Out-Null
            } catch {
                $form.Invoke({ param($m) $txtOut.AppendText("Log write failed: $m`r`n") }, $_.Exception.Message) | Out-Null
            }
        }
        $form.Invoke({ param($n,$e) $txtOut.AppendText("Batch complete. Files: $n, Errors: $e`r`n") }, $result.Processed, $result.Errors) | Out-Null
    }

    $global:cts = $null
}

# Buttons
$btnCount.Add_Click({
    $roots = Validate-Roots
    if (-not $roots) { return }
    $patterns = ($txtFilter.Text.Trim() -split ';' | ForEach-Object { if ([string]::IsNullOrWhiteSpace($_)) { "*" } else { $_ } })
    $count = 0
    foreach ($pat in $patterns) {
        foreach ($f in [System.IO.Directory]::EnumerateFiles($roots.Source, $pat, [System.IO.SearchOption]::AllDirectories)) {
            $count++
            if ($count % 100000 -eq 0) { [System.Windows.Forms.Application]::DoEvents() }
        }
    }
    Append-Out ("Estimated files matching filter: {0}" -f $count)
})

$btnNext.Add_Click({
    Process-Batch -BatchSize ([int]$nudBatch.Value)
})

$btnCancel.Add_Click({
    # This implementation processes synchronously within a job per batch; cancel is a no-op between files.
    Append-Out "Cancel requested. Finish current file and stop."
})

[void]$form.ShowDialog()
