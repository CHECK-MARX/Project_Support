#requires -version 5.1

param([string]$BasePath)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = 'Stop'

$script:themePalettes = @{
    Light = @{
        FormBack      = [System.Drawing.Color]::FromArgb(245, 246, 250)
        Surface       = [System.Drawing.Color]::FromArgb(234, 237, 244)
        Control       = [System.Drawing.Color]::White
        ControlFore   = [System.Drawing.Color]::FromArgb(33, 37, 41)
        MutedFore     = [System.Drawing.Color]::FromArgb(120, 124, 131)
        Accent        = [System.Drawing.Color]::FromArgb(0, 120, 215)
        AccentText    = [System.Drawing.Color]::White
        DisabledBack  = [System.Drawing.Color]::FromArgb(215, 218, 224)
        DisabledFore  = [System.Drawing.Color]::FromArgb(150, 154, 162)
        Border        = [System.Drawing.Color]::FromArgb(200, 204, 214)
    }
    Dark = @{
        FormBack      = [System.Drawing.Color]::FromArgb(32, 34, 38)
        Surface       = [System.Drawing.Color]::FromArgb(44, 46, 51)
        Control       = [System.Drawing.Color]::FromArgb(56, 58, 64)
        ControlFore   = [System.Drawing.Color]::FromArgb(230, 233, 240)
        MutedFore     = [System.Drawing.Color]::FromArgb(170, 175, 187)
        Accent        = [System.Drawing.Color]::FromArgb(10, 132, 255)
        AccentText    = [System.Drawing.Color]::White
        DisabledBack  = [System.Drawing.Color]::FromArgb(70, 72, 78)
        DisabledFore  = [System.Drawing.Color]::FromArgb(135, 138, 147)
        Border        = [System.Drawing.Color]::FromArgb(85, 88, 96)
    }
}

$script:isDarkTheme = $false
$script:mainForm = $null

function Get-ControlTree {
    param([System.Windows.Forms.Control]$Root)
    $list = New-Object 'System.Collections.Generic.List[System.Windows.Forms.Control]'
    if (-not $Root) { return $list }
    $queue = New-Object System.Collections.Generic.Queue[System.Windows.Forms.Control]
    $queue.Enqueue($Root)
    while ($queue.Count -gt 0) {
        $node = $queue.Dequeue()
        foreach ($child in $node.Controls) {
            $list.Add($child)
            if ($child.Controls.Count -gt 0) {
                $queue.Enqueue($child)
            }
        }
    }
    return $list
}

function Apply-Theme {
    param([bool]$DarkMode)
    $script:isDarkTheme = [bool]$DarkMode
    if (-not $script:mainForm) { return }

    $palette = if ($DarkMode) { $script:themePalettes.Dark } else { $script:themePalettes.Light }

    $script:mainForm.BackColor = $palette.FormBack
    $script:mainForm.ForeColor = $palette.ControlFore

    foreach ($ctrl in (Get-ControlTree -Root $script:mainForm)) {
        if (-not $ctrl) { continue }
        switch ($ctrl.GetType().Name) {
            'Label' {
                $ctrl.BackColor = [System.Drawing.Color]::Transparent
                if ($ctrl.ForeColor.ToArgb() -eq [System.Drawing.Color]::Gray.ToArgb()) {
                    $ctrl.ForeColor = $palette.MutedFore
                }
                else {
                    $ctrl.ForeColor = $palette.ControlFore
                }
            }
            'TextBox' {
                $ctrl.BackColor = $palette.Control
                $ctrl.ForeColor = $palette.ControlFore
                $ctrl.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
            }
            'ComboBox' {
                $ctrl.BackColor = $palette.Control
                $ctrl.ForeColor = $palette.ControlFore
                $ctrl.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            }
            'DateTimePicker' {
                $ctrl.CalendarForeColor = $palette.ControlFore
                $ctrl.CalendarMonthBackground = $palette.Control
                $ctrl.CalendarTitleBackColor = $palette.Surface
                $ctrl.CalendarTitleForeColor = $palette.ControlFore
                $ctrl.BackColor = $palette.Control
                $ctrl.ForeColor = $palette.ControlFore
            }
            'Panel' {
                $ctrl.BackColor = $palette.Surface
                $ctrl.ForeColor = $palette.ControlFore
            }
            'FlowLayoutPanel' {
                $ctrl.BackColor = $palette.Surface
                $ctrl.ForeColor = $palette.ControlFore
            }
            'CheckBox' {
                $ctrl.BackColor = $palette.Surface
                $ctrl.ForeColor = $palette.ControlFore
            }
            'Button' {
                $ctrl.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $ctrl.FlatAppearance.BorderSize = 1
                $ctrl.FlatAppearance.BorderColor = $palette.Border
                $ctrl.FlatAppearance.MouseOverBackColor = $palette.Accent
                $ctrl.FlatAppearance.MouseDownBackColor = $palette.Accent
                $ctrl.UseVisualStyleBackColor = $false
                if ($ctrl.Enabled) {
                    $ctrl.BackColor = $palette.Control
                    $ctrl.ForeColor = $palette.ControlFore
                }
                else {
                    $ctrl.BackColor = $palette.DisabledBack
                    $ctrl.ForeColor = $palette.DisabledFore
                }
            }
            default {
                if ($ctrl.BackColor.IsEmpty -or $ctrl.BackColor.ToArgb() -eq [System.Drawing.SystemColors]::Control.ToArgb()) {
                    $ctrl.BackColor = $palette.Surface
                }
                $ctrl.ForeColor = $palette.ControlFore
            }
        }
    }
    $script:mainForm.Refresh()
}

function Resolve-ExistingPath {
    param(
        [string[]]$Candidates,
        [switch]$AllowDirectory
    )
    if (-not $Candidates) { return $null }
    foreach ($candidate in $Candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        try {
            if (Test-Path -LiteralPath $candidate) {
                $resolved = (Resolve-Path -LiteralPath $candidate -ErrorAction Stop).Path
                if ($AllowDirectory) { return $resolved }
                try {
                    $item = Get-Item -LiteralPath $resolved -ErrorAction Stop
                    if (-not $item.PSIsContainer) { return $resolved }
                }
                catch {
                    if (Test-Path -LiteralPath $resolved -PathType Leaf) { return $resolved }
                }
            }
        }
        catch {
            try {
                if (Test-Path -LiteralPath $candidate) {
                    if ($AllowDirectory) { return $candidate }
                    if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
                }
            }
            catch { }
        }
    }
    return $null
}

function U {
    param([string]$Literal)
    return [System.Text.RegularExpressions.Regex]::Unescape($Literal)
}

function Write-Log {
    param([string]$Message)
    if (-not $script:LogFile) { return }
    $line = "[{0}] {1}" -f (Get-Date).ToString('u'), $Message
    Add-Content -Path $script:LogFile -Value $line -Encoding UTF8
}

function Remove-InvalidPathChars {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $builder = New-Object System.Text.StringBuilder
    foreach ($ch in $Text.ToCharArray()) {
        if ($invalid -notcontains $ch) {
            [void]$builder.Append($ch)
        }
    }
    return $builder.ToString().Trim()
}

function Normalize-SupportNumber {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $trim = $Text.Trim()
    if ($trim -match '^\d+$') {
        $normalized = $trim.TrimStart('0')
        if ([string]::IsNullOrWhiteSpace($normalized)) { return '0' }
        return $normalized
    }
    return $trim.ToUpperInvariant()
}

function Ensure-NormalizedSupport {
    param($Case)
    if (-not $Case) { return $Case }

    if ($Case.PSObject.Properties['SupportNumber']) {
        $Case.SupportNumber = if ([string]::IsNullOrWhiteSpace($Case.SupportNumber)) { '' } else { $Case.SupportNumber.Trim() }
    }

    $normalized = Normalize-SupportNumber ($Case.SupportNumber)
    if ($Case.PSObject.Properties['NormalizedSupport']) {
        $Case.NormalizedSupport = $normalized
    }
    else {
        $Case | Add-Member -NotePropertyName NormalizedSupport -NotePropertyValue $normalized -Force
    }
    return $Case
}

function Split-StatusAndStamp {
    param([string]$Text)
    $status = if ($Text) { $Text.Trim() } else { '' }
    $stamp = ''
    if ([string]::IsNullOrWhiteSpace($status)) {
        return [PSCustomObject]@{ Status = ''; Stamp = '' }
    }
    while ($status -match "^(.*?)[_\-\s](\d{8})$") {
        $statusCandidate = $matches[1].TrimEnd('_',' ','-')
        $stampCandidate = $matches[2]
        $status = $statusCandidate
        if ([string]::IsNullOrWhiteSpace($stamp)) { $stamp = $stampCandidate }
    }
    return [PSCustomObject]@{ Status = $status.Trim(); Stamp = $stamp }
}

function Normalize-Status {
    param([string]$Text)
    return (Split-StatusAndStamp -Text $Text).Status
}

function Get-CaseFolderName {
    param(
        [string]$Date,
        [string]$Company,
        [string]$Support,
        [string]$Status,
        [string]$Updated
    )

    $safeDate = $Date
    if ([string]::IsNullOrWhiteSpace($safeDate)) { return '' }

    $cleanStatus = Normalize-Status $Status

    $safeCompany = Remove-InvalidPathChars -Text $Company
    if ([string]::IsNullOrWhiteSpace($safeCompany)) { return '' }

    $safeSupport = Remove-InvalidPathChars -Text $Support
    $safeStatus = Remove-InvalidPathChars -Text $cleanStatus
    if ([string]::IsNullOrWhiteSpace($safeStatus)) { return '' }

    $inner = $safeCompany
    if (-not [string]::IsNullOrWhiteSpace($safeSupport)) {
        $inner = '{0}_{1}' -f $safeCompany, $safeSupport
    }

    $updatedStamp = Remove-InvalidPathChars -Text $Updated
    if ([string]::IsNullOrWhiteSpace($updatedStamp)) {
        return '{0}({1}){2}' -f $safeDate, $inner, $safeStatus
    }
    else {
        return '{0}({1}){2}_{3}' -f $safeDate, $inner, $safeStatus, $updatedStamp
    }
}

function Get-UpdateStampFromIso {
    param([string]$Iso)
    if ([string]::IsNullOrWhiteSpace($Iso)) { return '' }
    try {
        return ([DateTime]::Parse($Iso)).ToString('yyyyMMdd')
    }
    catch {
        return ''
    }
}

function Get-NoteFileName {
    param(
        [string]$BaseName,
        [string]$SupportNumber
    )

    $base = Remove-InvalidPathChars -Text $BaseName
    if ([string]::IsNullOrWhiteSpace($base)) { $base = 'note' }

    $support = Remove-InvalidPathChars -Text $SupportNumber
    if ([string]::IsNullOrWhiteSpace($support)) {
        return '{0}.txt' -f $base
    }
    else {
        return '{0}_{1}.txt' -f $base, $support
    }
}

function Sanitize-FileName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    $clean = Remove-InvalidPathChars -Text $Name
    if ([string]::IsNullOrWhiteSpace($clean)) { return '' }
    if (-not $clean.EndsWith('.txt', [System.StringComparison]::OrdinalIgnoreCase)) {
        $clean += '.txt'
    }
    return $clean
}

function Backup-CaseNoteFile {
    param(
        [string]$FolderPath,
        [string]$FilePath
    )

    if ([string]::IsNullOrWhiteSpace($FolderPath) -or [string]::IsNullOrWhiteSpace($FilePath)) { return $false }
    if (-not (Test-Path -LiteralPath $FilePath)) { return $false }

    $backupDir = Join-Path -Path $FolderPath -ChildPath '_bak_'
    try {
        if (-not (Test-Path -LiteralPath $backupDir)) {
            New-Item -Path $backupDir -ItemType Directory -ErrorAction Stop | Out-Null
        }

        $leaf = Split-Path -Path $FilePath -Leaf
        $backupPath = Join-Path -Path $backupDir -ChildPath $leaf
        if (Test-Path -LiteralPath $backupPath) {
            $name = [System.IO.Path]::GetFileNameWithoutExtension($leaf)
            $ext = [System.IO.Path]::GetExtension($leaf)
            $suffix = (Get-Date).ToString('yyyyMMddHHmmss')
            $leaf = '{0}_{1}{2}' -f $name, $suffix, $ext
            $backupPath = Join-Path -Path $backupDir -ChildPath $leaf
        }

        Move-Item -LiteralPath $FilePath -Destination $backupPath -ErrorAction Stop
        Write-Log ("Backed up note file {0} -> {1}" -f $FilePath, $backupPath)
        return $true
    }
    catch {
        Write-Log ("Backup failed for note file {0}: {1}" -f $FilePath, $_.Exception.Message)
        return $false
    }
}

function Copy-CaseNoteFileToBackup {
    param(
        [string]$FolderPath,
        [string]$SourcePath
    )

    if ([string]::IsNullOrWhiteSpace($FolderPath) -or [string]::IsNullOrWhiteSpace($SourcePath)) { return $null }
    if (-not (Test-Path -LiteralPath $SourcePath)) { return $null }

    $backupDir = Join-Path -Path $FolderPath -ChildPath '_bak_'
    try {
        if (-not (Test-Path -LiteralPath $backupDir)) {
            New-Item -Path $backupDir -ItemType Directory -ErrorAction Stop | Out-Null
        }
        $leaf = Split-Path -Path $SourcePath -Leaf
        $backupPath = Join-Path -Path $backupDir -ChildPath $leaf

        while (Test-Path -LiteralPath $backupPath) {
            $name = [System.IO.Path]::GetFileNameWithoutExtension($leaf)
            $ext = [System.IO.Path]::GetExtension($leaf)
            $suffix = (Get-Date).ToString('yyyyMMddHHmmss')
            $leaf = '{0}_{1}{2}' -f $name, $suffix, $ext
            $backupPath = Join-Path -Path $backupDir -ChildPath $leaf
        }

        Copy-Item -LiteralPath $SourcePath -Destination $backupPath -ErrorAction Stop
        Write-Log ("Copied note file to backup {0} -> {1}" -f $SourcePath, $backupPath)
        return $backupPath
    }
    catch {
        Write-Log ("Copy to backup failed for note file {0}: {1}" -f $SourcePath, $_.Exception.Message)
        return $null
    }
}

function Get-NoteEditorTargetPath {
    param(
        $Case,
        $Editor
    )
    if (-not $Case -or -not $Editor) { return $null }
    if ([string]::IsNullOrWhiteSpace($Case.FolderPath) -or -not (Test-Path -LiteralPath $Case.FolderPath)) { return $null }

    $name = ''
    if ($Editor.FileBox -and $Editor.FileBox.Text) {
        $name = Sanitize-FileName $Editor.FileBox.Text
    }
    if ([string]::IsNullOrWhiteSpace($name)) {
        $name = Get-NoteFileName -BaseName $Editor.BaseName -Support $Case.SupportNumber
    }
    if ([string]::IsNullOrWhiteSpace($name)) { return $null }
    return Join-Path -Path $Case.FolderPath -ChildPath $name
}

function Start-NoteEditorProcess {
    param(
        $Editor,
        [string]$Path,
        [switch]$Silent
    )

    if (-not $Editor) { return $false }
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }

    $resolvedPath = $Path
    try {
        if (Test-Path -LiteralPath $Path) {
            $resolvedPath = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
        }
        else {
            if (-not $Silent) {
                [System.Windows.Forms.MessageBox]::Show(
                    ((U "\u30d5\u30a1\u30a4\u30eb\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093: {0}") -f $Path),
                    $(U "\u8b66\u544a"),
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
            }
            return $false
        }
    }
    catch {
        $resolvedPath = $Path
    }

    if ($Editor.PSObject.Properties['OpenProcess'] -and $Editor.OpenProcess) {
        try {
            if (-not $Editor.OpenProcess.HasExited) {
                if (-not $Silent) {
                    [System.Windows.Forms.MessageBox]::Show(
                        $(U "\u3053\u306e\u30d5\u30a1\u30a4\u30eb\u306f\u65e2\u306b\u958b\u3044\u3066\u3044\u307e\u3059\u3002"),
                        $(U "\u60c5\u5831"),
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information
                    ) | Out-Null
                }
                return $true
            }
            $Editor.OpenProcess.Dispose()
        }
        catch { }
        $Editor.OpenProcess = $null
    }

    try {
        $proc = Start-Process -FilePath 'notepad.exe' -ArgumentList ("`"{0}`"" -f $resolvedPath) -PassThru -ErrorAction Stop
    }
    catch {
        if (-not $Silent) {
            [System.Windows.Forms.MessageBox]::Show(
                ((U "\u30d5\u30a1\u30a4\u30eb\u3092\u958b\u3051\u307e\u305b\u3093\u3067\u3057\u305f\u3002\u8a73\u7d30: {0}") -f $_.Exception.Message),
                $(U "\u30a8\u30e9\u30fc"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }
        Write-Log ("Start-NoteEditorProcess failed: {0}" -f $_.Exception.Message)
        return $false
    }

    $Editor.OpenProcess = $proc
    return $true
}

function Close-NoteEditorProcess {
    param(
        $Editor,
        [bool]$Silent = $false
    )
    if (-not $Editor) { return }
    if (-not $Editor.PSObject.Properties['OpenProcess']) { return }
    $proc = $Editor.OpenProcess
    if (-not $proc) { return }

    if (-not $proc.HasExited) {
        try {
            if (-not $proc.CloseMainWindow()) {
                if (-not $Silent) {
                    [System.Windows.Forms.MessageBox]::Show(
                        $(U "\u30d5\u30a1\u30a4\u30eb\u3092\u624b\u52d5\u3067\u9589\u3058\u3066\u304f\u3060\u3055\u3044\u3002"),
                        $(U "\u8b66\u544a"),
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Warning
                    ) | Out-Null
                }
                return
            }
            if (-not $proc.WaitForExit(5000)) {
                if (-not $Silent) {
                    [System.Windows.Forms.MessageBox]::Show(
                        $(U "\u30a8\u30c7\u30a3\u30bf\u304c\u7d42\u4e86\u3057\u307e\u305b\u3093\u3002\u624b\u52d5\u3067\u78ba\u8a8d\u3057\u3066\u304f\u3060\u3055\u3044\u3002"),
                        $(U "\u8b66\u544a"),
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Warning
                    ) | Out-Null
                }
                return
            }
        }
        catch {
            if (-not $Silent) {
                [System.Windows.Forms.MessageBox]::Show(
                    ((U "\u30d5\u30a1\u30a4\u30eb\u9589\u3058\u8a66\u884c\u306b\u5931\u6557\u3057\u307e\u3057\u305f\u3002\u8a73\u7d30: {0}") -f $_.Exception.Message),
                    $(U "\u30a8\u30e9\u30fc"),
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                ) | Out-Null
            }
            return
        }
    }

    try { $proc.Dispose() } catch { }
    $Editor.OpenProcess = $null
    Update-NoteEditorButtons -HasCase:($script:currentCase -ne $null)
}

function Update-NoteEditorButtons {
    param([bool]$HasCase)
    foreach ($editor in $noteEditors) {
        $activeProcess = $false
        if ($editor.PSObject.Properties['OpenProcess'] -and $editor.OpenProcess) {
            try {
                if ($editor.OpenProcess.HasExited) {
                    try { $editor.OpenProcess.Dispose() } catch { }
                    $editor.OpenProcess = $null
                }
                else {
                    $activeProcess = $true
                }
            }
            catch {
                $editor.OpenProcess = $null
            }
        }

        if ($editor.OpenButton) {
            $editor.OpenButton.Enabled = $HasCase
        }
        if ($editor.CloseButton) {
            $editor.CloseButton.Enabled = $HasCase
        }
    }
}

function Ensure-CurrentCaseContext {
    param([switch]$Silent)

    if ($script:currentCase -and $script:currentCase.FolderPath -and (Test-Path -LiteralPath $script:currentCase.FolderPath)) {
        return $true
    }

    $support = $textSupportNo.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($support)) {
        $support = $textSearchSupport.Text.Trim()
    }

    if (-not [string]::IsNullOrWhiteSpace($support)) {
        $case = Load-CaseFromSupportNumber -SupportNumber $support -Silent:$Silent
        if ($case -and $case.FolderPath -and (Test-Path -LiteralPath $case.FolderPath)) {
            return $true
        }
    }

    if ($comboHistory -and $comboHistory.SelectedItem -and $comboHistory.SelectedItem.Data) {
        $historyItem = $comboHistory.SelectedItem.Data
        if ($historyItem -and $historyItem.FolderPath -and (Test-Path -LiteralPath $historyItem.FolderPath)) {
            Set-CurrentCase -Case $historyItem
            if ($script:currentCase) { return $true }
        }
    }

    if ($script:currentCase -and $script:currentCase.FolderPath -and (Test-Path -LiteralPath $script:currentCase.FolderPath)) {
        return $true
    }

    $previewFolder = $null
    if (-not [string]::IsNullOrWhiteSpace($textPreview.Text) -and -not $textPreview.Text.StartsWith('(')) {
        $base = $textBasePath.Text
        if (-not [string]::IsNullOrWhiteSpace($base)) {
            $previewFolder = Join-Path -Path $base -ChildPath $textPreview.Text
        }
    }

    if ($previewFolder -and (Test-Path -LiteralPath $previewFolder)) {
        $tempCase = [PSCustomObject]@{
            SupportNumber = $textSupportNo.Text.Trim()
            Company       = $textCompany.Text.Trim()
            Status        = Normalize-Status $comboStatus.Text
            CreatedOn     = $pickerDate.Value.ToString('yyyyMMdd')
            FolderName    = Split-Path -Path $previewFolder -Leaf
            FolderPath    = $previewFolder
            LastUpdated   = (Get-Date).ToString('o')
        }
        $script:currentCase = Ensure-NoteFiles -Case $tempCase
        Update-NoteEditorButtons -HasCase:([bool]$script:currentCase)
        if ($script:currentCase -and (Test-Path -LiteralPath $script:currentCase.FolderPath)) {
            return $true
        }
    }

    if (-not $Silent) {
        [System.Windows.Forms.MessageBox]::Show(
            $(U "\u5c65\u6b74\u304b\u3089\u6848\u4ef6\u3092\u9078\u629e\u3057\u3066\u304f\u3060\u3055\u3044\u3002"),
            $(U "\u8b66\u544a"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
    }
    return $false
}

function Update-NoteLabels {
    param($Case)
    if (-not $Case) {
        foreach ($editor in $noteEditors) {
            if ($editor.FileBox) { $editor.FileBox.Text = '' }
            if ($editor.PSObject.Properties['SelectedPath']) { $editor.SelectedPath = $null }
            if ($editor.FileBox) { $editor.FileBox.Tag = $null }
            if ($editor.SelectedLabel) { $editor.SelectedLabel.Text = '' }
            if ($editor.PSObject.Properties['SelectedSourceName']) { $editor.SelectedSourceName = '' }
            if ($editor.PSObject.Properties['LastSourcePath']) { $editor.LastSourcePath = $null }
        }
        return
    }

    foreach ($editor in $noteEditors) {
        if ($editor.PSObject.Properties['SelectedPath']) { $editor.SelectedPath = $null }
        if ($editor.FileBox) {
            $editor.FileBox.Tag = $null
            $editor.FileBox.Text = if ([string]::IsNullOrWhiteSpace($editor.FileBox.Text)) {
                Get-NoteFileName -BaseName $editor.BaseName -Support $Case.SupportNumber
            }
            else {
                Sanitize-FileName $editor.FileBox.Text
            }
        }
        if ($editor.SelectedLabel) { $editor.SelectedLabel.Text = '' }
        if ($editor.PSObject.Properties['SelectedSourceName']) { $editor.SelectedSourceName = '' }
    }
    Update-NoteEditorButtons -HasCase:([bool]$Case)
}

function Close-ExplorerWindow {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    $normalized = $Path.TrimEnd('\')
    try {
        $shell = New-Object -ComObject Shell.Application
        foreach ($window in $shell.Windows()) {
            try {
                $locationUrl = $window.LocationURL
                if ([string]::IsNullOrWhiteSpace($locationUrl)) { continue }
                if ($locationUrl.StartsWith('file:///')) {
                    $localPath = [System.Uri]::UnescapeDataString($locationUrl.Substring(8)).Replace('/', '\')
                }
                else {
                    $localPath = [System.Uri]::UnescapeDataString($locationUrl).Replace('/', '\')
                }
                if ([string]::IsNullOrWhiteSpace($localPath)) { continue }
                if ([string]::Equals($localPath.TrimEnd('\'), $normalized, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $window.Quit()
                }
            }
            catch { continue }
        }
    }
    catch { }
}

function Test-FolderWindowOpen {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    $normalized = $Path.TrimEnd('\')
    try {
        $shell = New-Object -ComObject Shell.Application
        foreach ($window in $shell.Windows()) {
            try {
                $locationUrl = $window.LocationURL
                if ([string]::IsNullOrWhiteSpace($locationUrl)) { continue }
                if ($locationUrl.StartsWith('file:///')) {
                    $localPath = [System.Uri]::UnescapeDataString($locationUrl.Substring(8)).Replace('/', '\')
                }
                else {
                    $localPath = [System.Uri]::UnescapeDataString($locationUrl).Replace('/', '\')
                }
                if ([string]::IsNullOrWhiteSpace($localPath)) { continue }
                if ([string]::Equals($localPath.TrimEnd('\'), $normalized, [System.StringComparison]::OrdinalIgnoreCase)) {
                    return $true
                }
            }
            catch { continue }
        }
    }
    catch { }
    return $false
}

function Close-CurrentCaseFolder {
    param([bool]$ShowMessage = $false)
    if (-not $script:currentCase -or [string]::IsNullOrWhiteSpace($script:currentCase.FolderPath)) {
        if ($ShowMessage) {
            [System.Windows.Forms.MessageBox]::Show(
                $(U "\u5c65\u6b74\u304b\u3089\u6848\u4ef6\u3092\u9078\u629e\u3057\u3066\u304f\u3060\u3055\u3044\u3002"),
                $(U "\u8b66\u544a"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
        }
        return
    }

    $path = $script:currentCase.FolderPath
    if ($script:openedFolders.Contains($path)) {
        Close-ExplorerWindow -Path $path
        $script:openedFolders.Remove($path) | Out-Null
        if ($ShowMessage) {
            [System.Windows.Forms.MessageBox]::Show(
                ((U "\u30d5\u30a9\u30eb\u30c0\u3092\u9589\u3058\u307e\u3057\u305f: {0}") -f $path),
                $(U "\u5b8c\u4e86"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        }
    }
    elseif ($ShowMessage) {
        [System.Windows.Forms.MessageBox]::Show(
            $(U "\u5bfe\u8c61\u30d5\u30a9\u30eb\u30c0\u306f\u958b\u3044\u3066\u3044\u307e\u305b\u3093\u3002"),
            $(U "\u60c5\u5831"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
}

function Move-FileToBackup {
    param(
        [string]$FilePath,
        [string]$BackupFolder
    )
    if (-not (Test-Path -LiteralPath $FilePath)) { return }
    if (-not (Test-Path -LiteralPath $BackupFolder)) {
        try {
            New-Item -Path $BackupFolder -ItemType Directory -ErrorAction Stop | Out-Null
        }
        catch { return }
    }

    $name = Split-Path -Path $FilePath -Leaf
    $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
    $ext = [System.IO.Path]::GetExtension($name)
    $timestamp = (Get-Date).ToString('yyyyMMddHHmmssfff')
    if ([string]::IsNullOrWhiteSpace($base)) { $base = 'note' }
    $targetName = '{0}_{1}{2}' -f $base, $timestamp, $ext
    $targetPath = Join-Path -Path $BackupFolder -ChildPath $targetName
    try {
        Move-Item -LiteralPath $FilePath -Destination $targetPath -ErrorAction Stop
    }
    catch {
        Write-Log ("Backup move failed: {0}" -f $_.Exception.Message)
    }
}

function Get-IndexFilePath {
    param([string]$BasePath)
    return (Join-Path -Path $BasePath -ChildPath 'cases-index.json')
}

function Load-CaseIndex {
    param([string]$BasePath)
    $indexFile = Get-IndexFilePath -BasePath $BasePath
    if (-not (Test-Path -Path $indexFile)) { return @() }
    try {
        $json = Get-Content -Path $indexFile -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($json)) { return @() }
        $data = $json | ConvertFrom-Json
        if ($null -eq $data) { return @() }
        if ($data -isnot [System.Collections.IEnumerable]) { return @($data) }
        return @($data)
    }
    catch {
        Write-Log ("Load-CaseIndex failed: {0}" -f $_.Exception.Message)
        return @()
    }
}

function Save-CaseIndex {
    param(
        [string]$BasePath,
        [System.Collections.IEnumerable]$Index
    )
    $indexFile = Get-IndexFilePath -BasePath $BasePath
    try {
        $Index | ConvertTo-Json -Depth 6 | Set-Content -Path $indexFile -Encoding UTF8
    }
    catch {
        Write-Log ("Save-CaseIndex failed: {0}" -f $_.Exception.Message)
    }
}

function Build-DisplayText {
    param($Case)
    $updateStamp = Get-UpdateStampFromIso $Case.LastUpdated
    if ([string]::IsNullOrWhiteSpace($updateStamp)) {
        return "{0} ({1} {2}) {3}" -f $Case.CreatedOn, $Case.Company, $Case.SupportNumber, $Case.Status
    }
    else {
        return "{0} ({1} {2}) {3} [{4}]" -f $Case.CreatedOn, $Case.Company, $Case.SupportNumber, $Case.Status, $updateStamp
    }
}

function Get-AllCasesFromFolders {
    param(
        [string]$BasePath
    )
    
    if (-not (Test-Path -Path $BasePath)) {
        return @()
    }
    
    $allCases = @()
    
    try {
        # 再帰的にフォルダを検索
        $directories = Get-ChildItem -Path $BasePath -Directory -Recurse -ErrorAction SilentlyContinue
        
        foreach ($dir in $directories) {
            # フォルダ名のパターンマッチング（すべての案件を含む）
            if ($dir.Name -match "^(?<date>\d{8})\((?<inner>.+)\)(?<status>.+)$") {
                $status = $matches['status'].Trim()
                $inner = $matches['inner'].Trim()
                $created = $matches['date']
                
                # 会社名とサポート番号を解析
                $company = $inner
                $supportNumber = ""
                
                $lastUnderscore = $inner.LastIndexOf('_')
                if ($lastUnderscore -gt 0 -and $lastUnderscore -lt ($inner.Length - 1)) {
                    $company = $inner.Substring(0, $lastUnderscore).Trim()
                    $supportNumber = $inner.Substring($lastUnderscore + 1).Trim()
                }
                
                # ステータスから更新日時を抽出
                $lastUpdated = $dir.LastWriteTime.ToString('o')
                if ($status -match "(.+)_(\d{8})$") {
                    $status = $matches[1].Trim()
                    $updateDate = $matches[2]
                    try {
                        $lastUpdated = [DateTime]::ParseExact($updateDate, 'yyyyMMdd', $null).ToString('o')
                    }
                    catch { }
                }
                
                $caseObj = [PSCustomObject]@{
                    SupportNumber = $supportNumber
                    Company = $company
                    Status = $status
                    CreatedOn = $created
                    FolderName = $dir.Name
                    FolderPath = $dir.FullName
                    LastUpdated = $lastUpdated
                    IsFromFolder = $true
                }
                $allCases += (Ensure-NormalizedSupport -Case $caseObj)
            }
        }
    }
    catch {
        Write-Log "Failed to scan folders for all cases: $($_.Exception.Message)"
    }
    
    return $allCases
}

function Merge-CaseHistories {
    param(
        [array]$IndexedCases,
        [array]$FolderCases
    )
    
    $mergedCases = @()
    $seenPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $seenSupports = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    
    # インデックスされた案件を追加
    foreach ($case in $IndexedCases) {
        $case = Ensure-NormalizedSupport -Case $case
        $supportKey = if ([string]::IsNullOrWhiteSpace($case.NormalizedSupport)) { $null } else { $case.NormalizedSupport }
        if ($supportKey -and -not $seenSupports.Add($supportKey)) { continue }
        if ($seenPaths.Add($case.FolderPath)) {
            $mergedCases += $case
        }
    }
    
    # フォルダから検出された案件を追加（重複を避ける）
    foreach ($case in $FolderCases) {
        $case = Ensure-NormalizedSupport -Case $case
        $supportKey = if ([string]::IsNullOrWhiteSpace($case.NormalizedSupport)) { $null } else { $case.NormalizedSupport }
        if ($supportKey -and -not $seenSupports.Add($supportKey)) { continue }
        if ($seenPaths.Add($case.FolderPath)) {
            $mergedCases += $case
        }
    }
    
    # 最終更新日時でソート
    return $mergedCases | Sort-Object -Property @{ Expression = { $_.LastUpdated }; Descending = $true }
}

function Try-ParseCaseFromFolder {
    param(
        [System.IO.DirectoryInfo]$Directory,
        [string]$SupportNumber
    )

    if (-not $Directory) { return $null }

    $name = $Directory.Name
    $created = $Directory.CreationTime.ToString('yyyyMMdd')
    $status = $(U "\u672a\u8a2d\u5b9a")
    $company = $name
    $supportNumber = ''
    $lastUpdated = $Directory.LastWriteTime.ToString('o')

    $pattern = "^(?<date>\d{8})\((?<inner>.+)\)(?<status>.+)$"
    $regex = [System.Text.RegularExpressions.Regex]::new($pattern)
    if ($regex.IsMatch($name)) {
        $match = $regex.Match($name)
        $created = $match.Groups['date'].Value
        $inner = $match.Groups['inner'].Value.Trim()
        $statusCandidate = $match.Groups['status'].Value.Trim()
        if (-not [string]::IsNullOrWhiteSpace($statusCandidate)) {
            $status = $statusCandidate
        }
        $lastUnderscore = $inner.LastIndexOf('_')
        if ($lastUnderscore -gt 0 -and $lastUnderscore -lt ($inner.Length - 1)) {
            $candidateCompany = $inner.Substring(0, $lastUnderscore).Trim()
            $candidateSupport = $inner.Substring($lastUnderscore + 1).Trim()
            if (-not [string]::IsNullOrWhiteSpace($candidateCompany)) {
                $company = $candidateCompany
            }
            if (-not [string]::IsNullOrWhiteSpace($candidateSupport)) {
                $supportNumber = $candidateSupport
            }
        }
        else {
            $company = $inner
        }

        $statusInfo = Split-StatusAndStamp -Text $status
        $status = if ([string]::IsNullOrWhiteSpace($statusInfo.Status)) { $status } else { $statusInfo.Status }
        if (-not [string]::IsNullOrWhiteSpace($statusInfo.Stamp)) {
            try {
                $stampDate = [DateTime]::ParseExact($statusInfo.Stamp, 'yyyyMMdd', $null)
                $lastUpdated = $stampDate.ToString('o')
            }
            catch { }
        }
    }

    if ([string]::IsNullOrWhiteSpace($supportNumber) -and $company -match '(?<digits>\d{3,})$') {
        $supportNumber = $matches['digits']
        if ($company.Length -gt $supportNumber.Length) {
            $company = $company.Substring(0, $company.Length - $supportNumber.Length).TrimEnd('_', ' ', '-')
            if ([string]::IsNullOrWhiteSpace($company)) { $company = $Directory.Name }
        }
    }
    elseif ($SupportNumber) {
        $supportNumber = $SupportNumber
    }

    $statusInfoFinal = Split-StatusAndStamp -Text $status
    if (-not [string]::IsNullOrWhiteSpace($statusInfoFinal.Stamp)) {
        try {
            $stampDate = [DateTime]::ParseExact($statusInfoFinal.Stamp, 'yyyyMMdd', $null)
            $lastUpdated = $stampDate.ToString('o')
        }
        catch { }
    }
    $status = if ([string]::IsNullOrWhiteSpace($statusInfoFinal.Status)) { $status } else { $statusInfoFinal.Status }

    $company = $company.Trim()
    if ([string]::IsNullOrWhiteSpace($company)) {
        $company = $Directory.Name
    }

    $status = Normalize-Status $status

    $caseObj = [PSCustomObject]@{
        SupportNumber = $supportNumber
        Company       = $company
        Status        = $status
        CreatedOn     = $created
        FolderName    = $name
        FolderPath    = $Directory.FullName
        LastUpdated   = $lastUpdated
    }
    return (Ensure-NormalizedSupport -Case $caseObj)
}

function Search-FolderBySupportNumber {
    param(
        [string]$SupportNumber,
        [string]$BasePath
    )

    if ([string]::IsNullOrWhiteSpace($SupportNumber) -or -not (Test-Path -Path $BasePath)) {
        return $null
    }

    $escapedSupport = [System.Management.Automation.WildcardPattern]::Escape($SupportNumber)
    $normalizedSupport = Normalize-SupportNumber $SupportNumber
    try {
        $directories = Get-ChildItem -Path $BasePath -Directory -Recurse -ErrorAction Stop
    }
    catch {
        Write-Log ("Search-FolderBySupportNumber enumeration failed: {0}" -f $_.Exception.Message)
        return $null
    }

    foreach ($dir in $directories) {
        if ([string]::IsNullOrWhiteSpace($normalizedSupport)) {
            if ($dir.Name -notlike "*$escapedSupport*") { continue }
        }
        elseif ($dir.Name -notlike "*$escapedSupport*" -and $dir.Name -notlike "*$normalizedSupport*") {
            continue
        }

        $indexed = $script:caseIndex |
            Where-Object { $_.FolderPath -eq $dir.FullName } |
            Select-Object -First 1

        if ($indexed) {
            $indexed = Ensure-NormalizedSupport -Case $indexed
            $matchNormalized = $indexed.NormalizedSupport
            if (-not [string]::IsNullOrWhiteSpace($normalizedSupport) -and
                $matchNormalized -eq $normalizedSupport) {
                return $indexed
            }
            if ((-not [string]::IsNullOrWhiteSpace($indexed.SupportNumber) -and
                 ($indexed.SupportNumber -like "*$escapedSupport*" -or $indexed.SupportNumber -like "*$normalizedSupport*")) -or
                ([string]::IsNullOrWhiteSpace($indexed.SupportNumber) -and
                 ($dir.Name -like "*$escapedSupport*" -or $dir.Name -like "*$normalizedSupport*"))) {
                return $indexed
            }
            else {
                continue
            }
        }

        $entry = Try-ParseCaseFromFolder -Directory $dir -SupportNumber $SupportNumber
        if ($entry) {
            $entryNormalized = $entry.NormalizedSupport
            if (-not [string]::IsNullOrWhiteSpace($normalizedSupport) -and
                $entryNormalized -eq $normalizedSupport) {
                return $entry
            }
            if (
                (-not [string]::IsNullOrWhiteSpace($entry.SupportNumber) -and
                 ($entry.SupportNumber -like "*$escapedSupport*" -or $entry.SupportNumber -like "*$normalizedSupport*")) -or
                ([string]::IsNullOrWhiteSpace($entry.SupportNumber) -and
                 ($entry.FolderName -like "*$escapedSupport*" -or $entry.FolderName -like "*$normalizedSupport*"))
            ) {
                return $entry
            }
        }
    }
    return $null
}

function Ensure-BasePathReady {
    if ([string]::IsNullOrWhiteSpace($textBasePath.Text) -or -not (Test-Path -Path $textBasePath.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            $(U "\u30d9\u30fc\u30b9\u30d5\u30a9\u30eb\u30c0\u304c\u5b58\u5728\u3057\u307e\u305b\u3093\u3002"),
            $(U "\u30a8\u30e9\u30fc"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return $false
    }
    return $true
}

# --- init paths ---
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition
if ([string]::IsNullOrWhiteSpace($scriptDirectory)) {
    $scriptDirectory = (Get-Location).Path
}

$defaultBasePath = $scriptDirectory
if ((Split-Path -Leaf $scriptDirectory) -ieq 'Tools') {
    $parent = Split-Path -Parent $scriptDirectory
    if ([string]::IsNullOrWhiteSpace($parent) -eq $false) {
        $defaultBasePath = $parent
    }
}

# --- NEW: コンテキストメニューから渡された BasePath を優先適用 ---
if ($PSBoundParameters.ContainsKey('BasePath') -and -not [string]::IsNullOrWhiteSpace($BasePath)) {
    $candidate = $BasePath.Trim('"')  # 引用符で渡っても安全
    if (Test-Path -LiteralPath $candidate) {
        try { $defaultBasePath = (Resolve-Path -LiteralPath $candidate -ErrorAction Stop).Path }
        catch { $defaultBasePath = $candidate }
    }
}


$script:openedFolders = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$script:LogFile = Join-Path -Path $scriptDirectory -ChildPath 'SupportCaseManager.log'
Remove-Item -Path $script:LogFile -ErrorAction SilentlyContinue
Write-Log "Script start"

$statuses = @(
    $(U "\u53d7\u4ed8"),
    $(U "\u8abf\u67fb\u4e2d"),
    $(U "\u304a\u5ba2\u69d8\u56de\u7b54\u5f85\u3061"),
    $(U "\u30e1\u30fc\u30ab\u30fc\u78ba\u8a8d\u4e2d"),
    $(U "\u30af\u30ed\u30fc\u30ba\u4e88\u5b9a"),
    $(U "\u30af\u30ed\u30fc\u30ba")
)

$form = New-Object System.Windows.Forms.Form
$script:mainForm = $form
$form.Text = $(U "\u30b5\u30dd\u30fc\u30c8\u53d7\u4ed8\u30c7\u30a3\u30ec\u30af\u30c8\u30ea\u4f5c\u6210\u30c4\u30fc\u30eb (改善版)")
$form.Size = [System.Drawing.Size]::new(1000, 760)
$form.MinimumSize = [System.Drawing.Size]::new(970, 740)
$form.StartPosition = 'CenterScreen'
$form.TopMost = $true
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.Font = New-Object System.Drawing.Font('Yu Gothic UI', 10.5)

$doubleBufferedProperty = $form.GetType().GetProperty('DoubleBuffered', [System.Reflection.BindingFlags]'NonPublic,Instance')
if ($doubleBufferedProperty) { $doubleBufferedProperty.SetValue($form, $true, $null) }

$labelBasePath = New-Object System.Windows.Forms.Label
$labelBasePath.Text = $(U "\u30d9\u30fc\u30b9\u30d5\u30a9\u30eb\u30c0:")
$labelBasePath.Location = [System.Drawing.Point]::new(20, 20)
$labelBasePath.AutoSize = $true
$form.Controls.Add($labelBasePath)

$textBasePath = New-Object System.Windows.Forms.TextBox
$textBasePath.Location = [System.Drawing.Point]::new(120, 16)
$textBasePath.Size = [System.Drawing.Size]::new(600, 24)
$textBasePath.Text = $defaultBasePath
$form.Controls.Add($textBasePath)

$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Text = $(U "\u53c2\u7167...")
$buttonBrowse.Location = [System.Drawing.Point]::new(730, 15)
$buttonBrowse.Size = [System.Drawing.Size]::new(120, 26)
$form.Controls.Add($buttonBrowse)

$labelSearch = New-Object System.Windows.Forms.Label
$labelSearch.Text = $(U "\u30b5\u30dd\u30fc\u30c8\u756a\u53f7\u691c\u7d22:")
$labelSearch.Location = [System.Drawing.Point]::new(20, 60)
$labelSearch.AutoSize = $true
$form.Controls.Add($labelSearch)

$textSearchSupport = New-Object System.Windows.Forms.TextBox
$textSearchSupport.Location = [System.Drawing.Point]::new(150, 56)
$textSearchSupport.Size = [System.Drawing.Size]::new(150, 24)
$form.Controls.Add($textSearchSupport)

$buttonSearch = New-Object System.Windows.Forms.Button
$buttonSearch.Text = $(U "\u691c\u7d22")
$buttonSearch.Location = [System.Drawing.Point]::new(310, 55)
$buttonSearch.Size = [System.Drawing.Size]::new(90, 26)
$form.Controls.Add($buttonSearch)

$labelHistory = New-Object System.Windows.Forms.Label
$labelHistory.Text = $(U "\u5c65\u6b74:")
$labelHistory.Location = [System.Drawing.Point]::new(420, 60)
$labelHistory.AutoSize = $true
$form.Controls.Add($labelHistory)

$comboHistory = New-Object System.Windows.Forms.ComboBox
$comboHistory.Location = [System.Drawing.Point]::new(520, 56)
$comboHistory.Size = [System.Drawing.Size]::new(340, 24)
$comboHistory.DropDownStyle = 'DropDownList'
$comboHistory.DropDownWidth = 600  # ドロップダウンの幅を広く設定
$form.Controls.Add($comboHistory)

$buttonReloadHistory = New-Object System.Windows.Forms.Button
$buttonReloadHistory.Text = $(U "\u5c65\u6b74\u66f4\u65b0")
$buttonReloadHistory.Location = [System.Drawing.Point]::new(520, 86)
$buttonReloadHistory.Size = [System.Drawing.Size]::new(120, 26)
$form.Controls.Add($buttonReloadHistory)

$buttonNew = New-Object System.Windows.Forms.Button
$buttonNew.Text = $(U "\u65b0\u898f")
$buttonNew.Location = [System.Drawing.Point]::new(690, 145)
$buttonNew.Size = [System.Drawing.Size]::new(70, 26)
$form.Controls.Add($buttonNew)

$buttonOpenFolder = New-Object System.Windows.Forms.Button
$buttonOpenFolder.Text = $(U "\u958b\u304f")
$buttonOpenFolder.Location = [System.Drawing.Point]::new(765, 145)
$buttonOpenFolder.Size = [System.Drawing.Size]::new(70, 26)
$buttonOpenFolder.Enabled = $false
$form.Controls.Add($buttonOpenFolder)

$labelDate = New-Object System.Windows.Forms.Label
$labelDate.Text = $(U "\u53d7\u4ed8\u65e5:")
$labelDate.Location = [System.Drawing.Point]::new(20, 110)
$labelDate.AutoSize = $true
$form.Controls.Add($labelDate)

$pickerDate = New-Object System.Windows.Forms.DateTimePicker
$pickerDate.Format = 'Short'
$pickerDate.Location = [System.Drawing.Point]::new(120, 106)
$pickerDate.Value = Get-Date
$form.Controls.Add($pickerDate)

$labelCompany = New-Object System.Windows.Forms.Label
$labelCompany.Text = $(U "\u4f1a\u793e\u540d:")
$labelCompany.Location = [System.Drawing.Point]::new(20, 150)
$labelCompany.AutoSize = $true
$form.Controls.Add($labelCompany)

$textCompany = New-Object System.Windows.Forms.TextBox
$textCompany.Location = [System.Drawing.Point]::new(120, 146)
$textCompany.Size = [System.Drawing.Size]::new(200, 24)
$form.Controls.Add($textCompany)

$labelSupportNo = New-Object System.Windows.Forms.Label
$labelSupportNo.Text = $(U "\u30b5\u30dd\u30fc\u30c8\u756a\u53f7:")
$labelSupportNo.Location = [System.Drawing.Point]::new(340, 150)
$labelSupportNo.AutoSize = $true
$form.Controls.Add($labelSupportNo)

$textSupportNo = New-Object System.Windows.Forms.TextBox
$textSupportNo.Location = [System.Drawing.Point]::new(440, 146)
$textSupportNo.Size = [System.Drawing.Size]::new(200, 24)
$form.Controls.Add($textSupportNo)

$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Text = $(U "\u30b9\u30c6\u30fc\u30bf\u30b9:")
$labelStatus.Location = [System.Drawing.Point]::new(20, 190)
$labelStatus.AutoSize = $true
$form.Controls.Add($labelStatus)

$comboStatus = New-Object System.Windows.Forms.ComboBox
$comboStatus.Location = [System.Drawing.Point]::new(120, 186)
$comboStatus.Size = [System.Drawing.Size]::new(200, 24)
$comboStatus.DropDownStyle = 'DropDown'
$comboStatus.Items.AddRange([object[]]$statuses)
$comboStatus.SelectedIndex = 0
$form.Controls.Add($comboStatus)

$labelCategory = New-Object System.Windows.Forms.Label
$labelCategory.Text = $(U "\u4fdd\u5b58\u5148\u30d5\u30a9\u30eb\u30c0:")
$labelCategory.Location = [System.Drawing.Point]::new(340, 190)
$labelCategory.AutoSize = $true
$form.Controls.Add($labelCategory)

$comboCategory = New-Object System.Windows.Forms.ComboBox
$comboCategory.Location = [System.Drawing.Point]::new(440, 186)
$comboCategory.Size = [System.Drawing.Size]::new(200, 24)
$comboCategory.DropDownStyle = 'DropDown'
$form.Controls.Add($comboCategory)

$buttonReloadCategory = New-Object System.Windows.Forms.Button
$buttonReloadCategory.Text = $(U "\u30d5\u30a9\u30eb\u30c0\u66f4\u65b0")
$buttonReloadCategory.Location = [System.Drawing.Point]::new(650, 185)
$buttonReloadCategory.Size = [System.Drawing.Size]::new(110, 26)
$form.Controls.Add($buttonReloadCategory)

$labelPreview = New-Object System.Windows.Forms.Label
$labelPreview.Text = $(U "\u4f5c\u6210\u30d5\u30a9\u30eb\u30c0\u540d (\u5fc5\u9808\u9805\u76ee\u3092\u5165\u529b\u3057\u3066\u304f\u3060\u3055\u3044)")
$labelPreview.Location = [System.Drawing.Point]::new(20, 230)
$labelPreview.AutoSize = $true
$form.Controls.Add($labelPreview)

$textPreview = New-Object System.Windows.Forms.TextBox
$textPreview.Location = [System.Drawing.Point]::new(180, 254)
$textPreview.Size = [System.Drawing.Size]::new(620, 22)
$textPreview.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$textPreview.ReadOnly = $true
$form.Controls.Add($textPreview)

$noteTemplates = @(
    @{ Label = $(U "\u304a\u5ba2\u69d8\u3054\u76f8\u8ac7\u5185\u5bb9"); BaseName = $(U "\u304a\u5ba2\u69d8\u3054\u76f8\u8ac7\u5185\u5bb9") },
    @{ Label = $(U "\u304a\u5ba2\u69d8\u3078\u306e\u8fd4\u4fe1\u6848"); BaseName = $(U "\u304a\u5ba2\u69d8\u3078\u306e\u8fd4\u4fe1\u6848") },
    @{ Label = $(U "\u30e1\u30fc\u30ab\u9023\u643a\u5185\u5bb9"); BaseName = $(U "\u30e1\u30fc\u30ab\u9023\u643a\u5185\u5bb9") }
)

$noteEditors = @()
$sectionBaseY = $textPreview.Location.Y + $textPreview.Height + 6
$sectionSpacing = 115
$index = 0
foreach ($tpl in $noteTemplates) {
    $sectionTop = $sectionBaseY + ($sectionSpacing * $index)
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $tpl.Label
    $label.Location = [System.Drawing.Point]::new(20, $sectionTop)
    $label.AutoSize = $true
    $form.Controls.Add($label)

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Multiline = $true
    $textbox.ScrollBars = 'Vertical'
    $textbox.Location = [System.Drawing.Point]::new(220, $sectionTop - 2)
    $textbox.Size = [System.Drawing.Size]::new(580, 64)
    $form.Controls.Add($textbox)

    $openButton = New-Object System.Windows.Forms.Button
    $openButton.Text = $(U "\u958b\u304f")
    $openButton.Location = [System.Drawing.Point]::new(20, $sectionTop + 26)
    $openButton.Size = [System.Drawing.Size]::new(90, 26)
    $openButton.Enabled = $false
    $form.Controls.Add($openButton)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = $(U "\u4fdd\u5b58")
    $closeButton.Location = [System.Drawing.Point]::new(20, $sectionTop + 50)
    $closeButton.Size = [System.Drawing.Size]::new(90, 26)
    $closeButton.Enabled = $false
    $form.Controls.Add($closeButton)

    $fileNameLabel = New-Object System.Windows.Forms.Label
    $fileNameLabel.Text = $(U "\u30d5\u30a1\u30a4\u30eb\u540d:")
    $fileNameLabel.Location = [System.Drawing.Point]::new(120, $sectionTop + 78)
    $fileNameLabel.AutoSize = $true
    $fileNameLabel.ForeColor = [System.Drawing.Color]::Gray
    $form.Controls.Add($fileNameLabel)

    $fileBox = New-Object System.Windows.Forms.TextBox
    $fileBox.Location = [System.Drawing.Point]::new(190, $sectionTop + 74)
    $fileBox.Size = [System.Drawing.Size]::new(320, 22)
    $fileBox.Text = ''
    $form.Controls.Add($fileBox)

    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Text = $(U "\u9078\u629e...")
    $browseButton.Location = [System.Drawing.Point]::new(530, $sectionTop + 70)
    $browseButton.Size = [System.Drawing.Size]::new(90, 26)
    $form.Controls.Add($browseButton)

    $selectedLabel = New-Object System.Windows.Forms.Label
    $selectedLabel.Text = ''
    $selectedLabel.Location = [System.Drawing.Point]::new(630, $sectionTop + 78)
    $selectedLabel.AutoSize = $true
    $selectedLabel.ForeColor = [System.Drawing.Color]::Gray
    $form.Controls.Add($selectedLabel)

    $noteEditor = [PSCustomObject]@{
        LabelText          = $tpl.Label
        BaseName           = $tpl.BaseName
        TextBox            = $textbox
        FileBox            = $fileBox
        BrowseButton       = $browseButton
        OpenButton         = $openButton
        CloseButton        = $closeButton
        SelectedLabel      = $selectedLabel
        SelectedPath       = $null
        SelectedSourceName = ''
        OpenProcess        = $null
        LastSourcePath     = $null
    }
    $noteEditors += $noteEditor

    $browseButton.Tag = $noteEditor
    $openButton.Tag = $noteEditor
    $closeButton.Tag = $noteEditor

    $browseHandler = {
        param($sender, $eventArgs)
        $editor = $sender.Tag
        if (-not $editor) { return }

        if (-not (Ensure-CurrentCaseContext -Silent)) {
            [System.Windows.Forms.MessageBox]::Show(
                $(U "\u5c65\u6b74\u304b\u3089\u6848\u4ef6\u3092\u9078\u629e\u3057\u3066\u304f\u3060\u3055\u3044\u3002"),
                $(U "\u8b66\u544a"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return
        }

        $initialDirectory = $null

        # --- 追加: サポート番号から案件を自動解決（現在のケースが無い/壊れている時だけ）---
        if (-not $script:currentCase -or
            -not $script:currentCase.FolderPath -or
            -not (Test-Path -LiteralPath $script:currentCase.FolderPath)) {
    
            $candidateSupport = $textSupportNo.Text.Trim()
            if (-not [string]::IsNullOrWhiteSpace($candidateSupport)) {
                $maybe = Load-CaseFromSupportNumber -SupportNumber $candidateSupport -Silent
                if ($maybe -and $maybe.FolderPath -and (Test-Path -LiteralPath $maybe.FolderPath)) {
                    $script:currentCase = $maybe
                }
            }
        }
    
        # --- 初期フォルダの優先順位 ---
        if ($script:currentCase -and $script:currentCase.FolderPath -and
            (Test-Path -LiteralPath $script:currentCase.FolderPath)) {
            $initialDirectory = $script:currentCase.FolderPath
        }
        elseif ($editor.PSObject.Properties['LastSourcePath'] -and
                $editor.LastSourcePath -and
                (Test-Path -LiteralPath $editor.LastSourcePath)) {
            $initialDirectory = Split-Path -Path $editor.LastSourcePath -Parent
        }
        elseif ($editor.FileBox -and $editor.FileBox.Tag -and
                (Test-Path -LiteralPath $editor.FileBox.Tag)) {
            $initialDirectory = Split-Path -Path $editor.FileBox.Tag -Parent
        }
        elseif (Test-Path -LiteralPath $textBasePath.Text) {
            $initialDirectory = $textBasePath.Text
        }
        else {
            try { $initialDirectory = [Environment]::GetFolderPath('MyDocuments') }
            catch { $initialDirectory = [Environment]::GetFolderPath('Desktop') }
        }
    
        if ($initialDirectory) {
            try { $initialDirectory = [System.IO.Path]::GetFullPath($initialDirectory) } catch { }
        }
    
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        if ($initialDirectory) { $dialog.InitialDirectory = $initialDirectory }
        $dialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        $dialog.RestoreDirectory = $true
        $dialog.CheckFileExists = $true

        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedPath = $dialog.FileName
            $selectedResolved = $selectedPath
            try { $selectedResolved = (Resolve-Path -LiteralPath $selectedPath -ErrorAction Stop).Path } catch { }
            $selectedLeaf = Split-Path -Path $selectedResolved -Leaf

            if ($script:currentCase -and (Test-Path -LiteralPath $script:currentCase.FolderPath)) {
                $caseFolder = $script:currentCase.FolderPath
                $targetName = Sanitize-FileName (Get-NoteFileName -BaseName $editor.BaseName -Support $textSupportNo.Text)
                if ([string]::IsNullOrWhiteSpace($targetName)) { $targetName = $selectedLeaf }
                $targetPath = Join-Path -Path $caseFolder -ChildPath $targetName

                $backupDir = Join-Path -Path $caseFolder -ChildPath '_bak_'
                if (-not (Test-Path -LiteralPath $backupDir)) {
                    try { New-Item -Path $backupDir -ItemType Directory -ErrorAction Stop | Out-Null } catch { }
                }

                $selectedParent = Split-Path -Path $selectedResolved -Parent
                $selectedInCase = [string]::Equals($selectedParent.TrimEnd('\'), $caseFolder.TrimEnd('\'), [System.StringComparison]::OrdinalIgnoreCase)

                $originalBackupPath = Join-Path -Path $backupDir -ChildPath $selectedLeaf
                while (Test-Path -LiteralPath $originalBackupPath) {
                    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($selectedLeaf)
                    $extName = [System.IO.Path]::GetExtension($selectedLeaf)
                    $suffix = (Get-Date).ToString('yyyyMMddHHmmss')
                    $originalBackupPath = Join-Path -Path $backupDir -ChildPath ("{0}_{1}{2}" -f $baseName, $suffix, $extName)
                }

                $sourceForTarget = $selectedResolved
                if ($selectedInCase) {
                    try {
                        Move-Item -LiteralPath $selectedResolved -Destination $originalBackupPath -Force -ErrorAction Stop
                        $sourceForTarget = $originalBackupPath
                    }
                    catch {
                        Write-Log ("Move original to backup failed: {0}" -f $_.Exception.Message)
                        $sourceForTarget = $selectedResolved
                    }
                }
                else {
                    try {
                        Copy-Item -LiteralPath $selectedResolved -Destination $originalBackupPath -Force -ErrorAction Stop
                    }
                    catch {
                        Write-Log ("Copy original to backup failed: {0}" -f $_.Exception.Message)
                    }
                }

                if (Test-Path -LiteralPath $targetPath) {
                    [void](Backup-CaseNoteFile -FolderPath $caseFolder -FilePath $targetPath)
                }

                try {
                    Copy-Item -LiteralPath $sourceForTarget -Destination $targetPath -Force -ErrorAction Stop
                    $editor.LastSourcePath = (Resolve-Path -LiteralPath $targetPath -ErrorAction Stop).Path
                }
                catch {
                    Write-Log ("Copy selected file to target failed: {0}" -f $_.Exception.Message)
                    $editor.LastSourcePath = $sourceForTarget
                }

                if ($editor.FileBox) {
                    $editor.FileBox.Text = Split-Path -Path $targetPath -Leaf
                    $editor.FileBox.Tag = $targetPath
                }
                if ($editor.SelectedLabel) {
                    $editor.SelectedLabel.Text = "{0} {1} {2} {3}" -f $(U "\u9078\u629e\u5143:"), $selectedLeaf, $(U "\u2192 \u9069\u7528\u5148:"), (Split-Path -Path $targetPath -Leaf)
                }
                if ($editor.PSObject.Properties['SelectedSourceName']) { $editor.SelectedSourceName = $selectedLeaf }
                if ($editor.PSObject.Properties['SelectedPath']) { $editor.SelectedPath = $targetPath }

                if ($script:currentCase.PSObject.Properties['LastUpdated']) {
                    $script:currentCase.LastUpdated = (Get-Date).ToString('o')
                }
                $script:currentCase = Ensure-NoteFiles -Case $script:currentCase
                Update-NoteEditorButtons -HasCase:$true
            }
            else {
                if ($editor.PSObject.Properties['SelectedSourceName']) { $editor.SelectedSourceName = $selectedLeaf }
                if ($editor.FileBox) {
                    $editor.FileBox.Tag = $selectedResolved
                    $suggested = Sanitize-FileName (Get-NoteFileName -BaseName $editor.BaseName -Support $textSupportNo.Text)
                    if ([string]::IsNullOrWhiteSpace($suggested)) { $suggested = $selectedLeaf }
                    $editor.FileBox.Text = $suggested
                }
                if ($editor.SelectedLabel) {
                    $editor.SelectedLabel.Text = "{0} {1}" -f $(U "\u9078\u629e\u5143:"), $selectedLeaf
                }
                if ($editor.PSObject.Properties['LastSourcePath']) { $editor.LastSourcePath = $selectedResolved }
            }
        }
    }.GetNewClosure()
    $browseButton.add_Click($browseHandler)

    $openHandler = {
        param($sender, $eventArgs)
        $editor = $sender.Tag
        if (-not $editor) { return }

        Write-Log ("Open handler invoked (BaseName={0})" -f $editor.BaseName)

        try {
            $fallbackCandidates = @()
            if ($editor.PSObject.Properties['LastSourcePath'] -and -not [string]::IsNullOrWhiteSpace($editor.LastSourcePath)) {
                $fallbackCandidates += $editor.LastSourcePath
            }
            if ($editor.FileBox -and $editor.FileBox.Tag) {
                $fallbackCandidates += $editor.FileBox.Tag
            }
            if ($editor.PSObject.Properties['SelectedPath'] -and $editor.SelectedPath) {
                $fallbackCandidates += $editor.SelectedPath
            }

            $existingPath = Resolve-ExistingPath $fallbackCandidates
            if ($existingPath -and (Test-Path -LiteralPath $existingPath -PathType Leaf)) {
                Write-Log ("Open handler attempting existing path {0}" -f $existingPath)
                try {
                    if (Start-NoteEditorProcess -Editor $editor -Path $existingPath) {
                        $leaf = Split-Path -Path $existingPath -Leaf
                        if ($editor.FileBox) {
                            $editor.FileBox.Text = $leaf
                            $editor.FileBox.Tag = $existingPath
                        }
                        if ($editor.PSObject.Properties['LastSourcePath']) { $editor.LastSourcePath = $existingPath }
                        if ($editor.SelectedLabel -and [string]::IsNullOrWhiteSpace($editor.SelectedLabel.Text)) {
                            $editor.SelectedLabel.Text = "{0} {1}" -f $(U "\u9069\u7528\u5148:"), $leaf
                        }
                        Update-NoteEditorButtons -HasCase:([bool]$script:currentCase)
                        return
                    }
                }
                catch {
                    Write-Log ("Open handler fallback launch failed: {0}" -f $_.Exception.Message)
                }
            }

            if (-not (Ensure-CurrentCaseContext -Silent)) {
                [System.Windows.Forms.MessageBox]::Show(
                    $(U "\u5c65\u6b74\u304b\u3089\u6848\u4ef6\u3092\u9078\u629e\u3057\u3066\u304f\u3060\u3055\u3044\u3002"),
                    $(U "\u8b66\u544a"),
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
                return
            }
            if (-not $script:currentCase) { return }
            if (-not (Test-Path -LiteralPath $script:currentCase.FolderPath)) {
                [System.Windows.Forms.MessageBox]::Show(
                    $(U "\u307e\u305a\u30d5\u30a9\u30eb\u30c0\u3092\u4f5c\u6210\u3057\u3066\u304f\u3060\u3055\u3044\u3002"),
                    $(U "\u8b66\u544a"),
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
                return
            }

            $script:currentCase = Ensure-NoteFiles -Case $script:currentCase -ApplyUserFilenames:$true
            $case = $script:currentCase
            $targetPath = Get-NoteEditorTargetPath -Case $case -Editor $editor
            if (-not $targetPath) {
                [System.Windows.Forms.MessageBox]::Show(
                    $(U "\u30d5\u30a1\u30a4\u30eb\u3092\u6c7a\u5b9a\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002"),
                    $(U "\u30a8\u30e9\u30fc"),
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                ) | Out-Null
                return
            }

            $openPath = $null
            if (-not [string]::IsNullOrWhiteSpace($targetPath)) {
                try {
                    if (Test-Path -LiteralPath $targetPath) {
                        $openPath = (Resolve-Path -LiteralPath $targetPath -ErrorAction Stop).Path
                    }
                }
                catch { $openPath = $targetPath }
            }

            if (-not $openPath) {
                $updatedFallback = $existingPath
                if ($editor.PSObject.Properties['LastSourcePath'] -and $editor.LastSourcePath) {
                    try {
                        if (Test-Path -LiteralPath $editor.LastSourcePath) {
                            $updatedFallback = (Resolve-Path -LiteralPath $editor.LastSourcePath -ErrorAction Stop).Path
                        }
                    }
                    catch { $updatedFallback = $editor.LastSourcePath }
                }
                if ($updatedFallback) {
                    $openPath = $updatedFallback
                }
                else {
                    [System.Windows.Forms.MessageBox]::Show(
                        $(U "\u958b\u3051\u308b\u30d5\u30a1\u30a4\u30eb\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002"),
                        $(U "\u30a8\u30e9\u30fc"),
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    ) | Out-Null
                    return
                }
            }

            if (-not (Start-NoteEditorProcess -Editor $editor -Path $openPath)) {
                return
            }
            if ($editor.SelectedLabel) {
                $leaf = Split-Path -Path $openPath -Leaf
                if ($editor.PSObject.Properties['SelectedSourceName'] -and -not [string]::IsNullOrWhiteSpace($editor.SelectedSourceName)) {
                    $editor.SelectedLabel.Text = "{0} {1} {2} {3}" -f $(U "\u9078\u629e\u5143:"), $editor.SelectedSourceName, $(U "\u2192 \u9069\u7528\u5148:"), $leaf
                }
                elseif ([string]::IsNullOrWhiteSpace($editor.SelectedLabel.Text)) {
                    $editor.SelectedLabel.Text = "{0} {1}" -f $(U "\u9069\u7528\u5148:"), $leaf
                }
            }
            if ($editor.PSObject.Properties['LastSourcePath']) {
                $editor.LastSourcePath = $openPath
            }
            Update-NoteEditorButtons -HasCase:$true
        }
        catch {
            Write-Log ("Open handler error (BaseName={0}): {1}" -f $editor.BaseName, $_.Exception.Message)
            [System.Windows.Forms.MessageBox]::Show(
                ((U "\u30d5\u30a1\u30a4\u30eb\u3092\u958b\u304f\u969b\u306b\u4e88\u671f\u305b\u306c\u30a8\u30e9\u30fc\u304c\u767a\u751f\u3057\u307e\u3057\u305f\u3002\u8a73\u7d30: {0}") -f $_.Exception.Message),
                $(U "\u30a8\u30e9\u30fc"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }
    }.GetNewClosure()
    $openButton.add_Click($openHandler)

    $closeHandler = {
        param($sender, $eventArgs)
        $editor = $sender.Tag
        if (-not $editor) { return }
        New-EditorAttachmentFolder -Editor $editor -ShowMessage
    }.GetNewClosure()
    $closeButton.add_Click($closeHandler)

    $index++
}

Update-NoteEditorButtons -HasCase:$false

$checkboxBaseY = $sectionBaseY + ($sectionSpacing * $noteTemplates.Count) + 8

$checkOpenAfter = New-Object System.Windows.Forms.CheckBox
$checkOpenAfter.Text = $(U "\u4f5c\u6210\u5f8c\u306b\u30d5\u30a9\u30eb\u30c0\u3092\u958b\u304f")
$checkOpenAfter.Location = [System.Drawing.Point]::new(20, $checkboxBaseY)
$checkOpenAfter.AutoSize = $true
$checkOpenAfter.Checked = $true
$form.Controls.Add($checkOpenAfter)

$checkDarkTheme = New-Object System.Windows.Forms.CheckBox
$checkDarkTheme.Text = $(U "\u30c0\u30fc\u30af\u30e2\u30fc\u30c9")
$checkDarkTheme.Location = [System.Drawing.Point]::new(200, $checkboxBaseY)
$checkDarkTheme.AutoSize = $true
$checkDarkTheme.Checked = $false
$checkDarkTheme.add_CheckedChanged({ Apply-Theme -DarkMode:$checkDarkTheme.Checked })
$form.Controls.Add($checkDarkTheme)

$buttonAppend = New-Object System.Windows.Forms.Button
$buttonAppend.Text = $(U "\u8ffd\u8a18\u4fdd\u5b58")
$buttonAppend.Location = [System.Drawing.Point]::new(340, $checkboxBaseY + 32)
$buttonAppend.Size = [System.Drawing.Size]::new(150, 34)
$buttonAppend.Enabled = $false
$form.Controls.Add($buttonAppend)

$buttonCloseFolder = New-Object System.Windows.Forms.Button
$buttonCloseFolder.Text = $(U "\u30d5\u30a9\u30eb\u30c0\u3092\u9589\u3058\u308b")
$buttonCloseFolder.Location = [System.Drawing.Point]::new(500, $checkboxBaseY + 32)
$buttonCloseFolder.Size = [System.Drawing.Size]::new(150, 34)
$buttonCloseFolder.Enabled = $false
$form.Controls.Add($buttonCloseFolder)

$buttonCreate = New-Object System.Windows.Forms.Button
$buttonCreate.Text = $(U "\u30d5\u30a9\u30eb\u30c0\u4f5c\u6210")
$buttonCreate.Location = [System.Drawing.Point]::new(660, $checkboxBaseY + 32)
$buttonCreate.Size = [System.Drawing.Size]::new(150, 34)
$form.Controls.Add($buttonCreate)

Apply-Theme -DarkMode:$checkDarkTheme.Checked

Write-Log "Controls initialized"







# --- load button-accent patch & wire once ---
try {
  . "$PSScriptRoot\src\button-accent.patch.ps1"
  Wire-ButtonAccent
} catch { }
# --------------------------------------------# --- load dark-mode persist (minimal) ---
try {
  . "$PSScriptRoot\src\persist-darkmode.min.ps1"
  Wire-PersistDarkMode-Min
} catch { }
# ----------------------------------------# --- external loader: openonly + persist-dark ---
try {
  . "$PSScriptRoot\src\openonly.patch.ps1"
  . "$PSScriptRoot\src\persist-darkmode.patch.ps1"
  if (Get-Command Wire-NoteOpenOnly -ErrorAction SilentlyContinue) { Wire-NoteOpenOnly }
  if (Get-Command Wire-PersistDarkMode -ErrorAction SilentlyContinue) { Wire-PersistDarkMode }
} catch { }
# -----------------------------------------------
# --- load persist-darkmode & wire ---
try {
  . "$PSScriptRoot\src\persist-darkmode.patch.ps1"
  if (Get-Command Wire-PersistDarkMode -ErrorAction SilentlyContinue) { Wire-PersistDarkMode }
} catch { }
# ------------------------------------
$script:caseIndex = @()
$script:currentCase = $null
$script:lastIndexBasePath = $null
$script:lastIndexTimestamp = $null

function Refresh-Categories {
    $comboCategory.Items.Clear()
    [void]$comboCategory.Items.Add($(U "(\u30d9\u30fc\u30b9\u76f4\u4e0b)"))
    if (Test-Path -Path $textBasePath.Text) {
        try {
            foreach ($dir in Get-ChildItem -Path $textBasePath.Text -Directory -ErrorAction Stop) {
                [void]$comboCategory.Items.Add($dir.Name)
            }
        }
        catch {
            Write-Log ("Refresh-Categories failed: {0}" -f $_.Exception.Message)
        }
    }
    $comboCategory.SelectedIndex = 0
}

function Refresh-History {
    $comboHistory.DataSource = $null
    $comboHistory.Items.Clear()

    if (-not $script:caseIndex -or $script:caseIndex.Count -eq 0) { return }

    $sorted = $script:caseIndex |
        Sort-Object -Property @{ Expression = { $_.LastUpdated }; Descending = $true },
                                   @{ Expression = { $_.CreatedOn }; Descending = $true }

    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $items = @()
    $closedStatus = $(U "\u30af\u30ed\u30fc\u30ba")
    foreach ($case in $sorted) {
        $case = Ensure-NormalizedSupport -Case $case
        $normalizedStatus = Normalize-Status $case.Status
        if ([string]::Equals($normalizedStatus, $closedStatus, [System.StringComparison]::Ordinal)) { continue }

        $key = if ([string]::IsNullOrWhiteSpace($case.NormalizedSupport)) {
            "{0}|{1}" -f $case.Company, $case.FolderPath
        } else {
            $case.NormalizedSupport
        }
        if (-not $seen.Add($key)) { continue }

        $items += [PSCustomObject]@{
            Display = Build-DisplayText -Case $case
            Data    = $case
        }
    }

    if ($items.Count -gt 0) {
        $array = [System.Array]::CreateInstance([object], $items.Count)
        for ($i = 0; $i -lt $items.Count; $i++) {
            $array.SetValue($items[$i], $i)
        }
        $comboHistory.DisplayMember = 'Display'
        $comboHistory.ValueMember = 'Data'
        $comboHistory.DataSource = $array
    }
}

function Load-CaseIndexForBasePath {
    if (-not (Ensure-BasePathReady)) { return }
    $basePath = $textBasePath.Text.Trim()
    $script:caseIndex = Load-CaseIndex -BasePath $basePath
    $script:caseIndex = @(
        $script:caseIndex |
        Where-Object { Test-Path -Path $_.FolderPath } |
        ForEach-Object { Ensure-CaseFolderName -Case $_ }
    )
    
    # フォルダからすべての案件を検出（クローズ済みも含む）
    $folderCases = Get-AllCasesFromFolders -BasePath $basePath
    $script:caseIndex = Merge-CaseHistories -IndexedCases $script:caseIndex -FolderCases $folderCases
    
    Save-CaseIndex -BasePath $basePath -Index $script:caseIndex
    $indexFile = Get-IndexFilePath -BasePath $basePath
    try {
        if (Test-Path -LiteralPath $indexFile) {
            $info = Get-Item -LiteralPath $indexFile -ErrorAction Stop
            $script:lastIndexTimestamp = $info.LastWriteTimeUtc
        }
        else {
            $script:lastIndexTimestamp = (Get-Date).ToUniversalTime()
        }
    }
    catch {
        $script:lastIndexTimestamp = (Get-Date).ToUniversalTime()
    }
    $script:lastIndexBasePath = $basePath
    Refresh-History
}

function Update-Preview {
    $date = $pickerDate.Value.ToString('yyyyMMdd')
    $updateStamp = if ($script:currentCase) {
        $stamp = Get-UpdateStampFromIso $script:currentCase.LastUpdated
        if ([string]::IsNullOrWhiteSpace($stamp)) { (Get-Date).ToString('yyyyMMdd') } else { $stamp }
    } else {
        (Get-Date).ToString('yyyyMMdd')
    }

    $statusText = Normalize-Status $comboStatus.Text

    $folderName = Get-CaseFolderName -Date $date -Company $textCompany.Text -Support $textSupportNo.Text -Status $statusText -Updated $updateStamp
    if ([string]::IsNullOrWhiteSpace($folderName)) {
        $textPreview.Text = $(U "(\u5fc5\u9808\u9805\u76ee\u3092\u5165\u529b\u3057\u3066\u304f\u3060\u3055\u3044)")
        return
    }
    $textPreview.Text = $folderName
}

function Clear-NoteEditors {
    foreach ($editor in $noteEditors) {
        $editor.TextBox.Clear()
        if ($editor.PSObject.Properties['SelectedPath']) { $editor.SelectedPath = $null }
        if ($editor.FileBox) { $editor.FileBox.Tag = $null }
        if ($editor.SelectedLabel) { $editor.SelectedLabel.Text = '' }
        if ($editor.PSObject.Properties['SelectedSourceName']) { $editor.SelectedSourceName = '' }
        if ($editor.PSObject.Properties['LastSourcePath']) { $editor.LastSourcePath = $null }
        if ($editor.PSObject.Properties['OpenProcess'] -and $editor.OpenProcess) {
            Close-NoteEditorProcess -Editor $editor -Silent:$true
        }
        if ($editor.OpenButton) { $editor.OpenButton.Enabled = $false }
        if ($editor.CloseButton) { $editor.CloseButton.Enabled = $false }
    }
    Update-NoteEditorButtons -HasCase:$false
}

function Set-CurrentCase {
    param($Case)

    if (-not $Case) {
        $script:currentCase = $null
        $textSearchSupport.Clear()
        $pickerDate.Value = Get-Date
        $textCompany.Clear()
        $textSupportNo.Clear()
        $comboStatus.SelectedIndex = 0
        Clear-NoteEditors
        $textPreview.Text = $(U "(\u5fc5\u9808\u9805\u76ee\u3092\u5165\u529b\u3057\u3066\u304f\u3060\u3055\u3044)")
        $buttonAppend.Enabled = $false
        $buttonOpenFolder.Enabled = $false
        $buttonCloseFolder.Enabled = $false
        Update-Preview
        Update-NoteLabels -Case $null
        return
    }

    $Case.Status = Normalize-Status $Case.Status
    $Case = Ensure-CaseFolderName -Case $Case
    $Case = Ensure-NoteFiles -Case $Case
    Update-CaseIndexEntry -Entry $Case
    $script:currentCase = $Case
    $textSearchSupport.Text = $Case.SupportNumber
    try {
        $pickerDate.Value = [DateTime]::ParseExact($Case.CreatedOn, 'yyyyMMdd', $null)
    }
    catch {
        $pickerDate.Value = Get-Date
    }
    $textCompany.Text = $Case.Company
    $textSupportNo.Text = $Case.SupportNumber
    $normalizedStatus = Normalize-Status $Case.Status
    if (-not [string]::IsNullOrWhiteSpace($normalizedStatus)) {
        if (-not $comboStatus.Items.Contains($normalizedStatus)) {
            [void]$comboStatus.Items.Add($normalizedStatus)
        }
        $comboStatus.SelectedItem = $normalizedStatus
    }
    $comboStatus.Text = $normalizedStatus
    $textPreview.Text = $Case.FolderName
    Clear-NoteEditors
    $buttonAppend.Enabled = $true
    $buttonOpenFolder.Enabled = $true
    $buttonCloseFolder.Enabled = $true
    Update-NoteLabels -Case $Case
    Update-NoteEditorButtons -HasCase:$true
}

function Update-CaseIndexEntry {
    param($Entry)
    if (-not $Entry) { return }

    $Entry = Ensure-NormalizedSupport -Case $Entry

    $script:caseIndex = @(
        $script:caseIndex |
        ForEach-Object {
            $existing = Ensure-NormalizedSupport -Case $_
            if ($existing.NormalizedSupport -ne $Entry.NormalizedSupport) {
                $existing
            }
        }
    )

    if ($Entry.Status -ne $(U "\u30af\u30ed\u30fc\u30ba")) {
        $script:caseIndex += $Entry
    }

    Save-CaseIndex -BasePath $textBasePath.Text -Index $script:caseIndex
    Refresh-History
}

function Ensure-CaseFolderName {
    param($Case)
    if (-not $Case) { return $Case }

    $Case.Status = Normalize-Status $Case.Status
    $Case = Ensure-NormalizedSupport -Case $Case

    if ([string]::IsNullOrWhiteSpace($Case.LastUpdated)) {
        try {
            $Case.LastUpdated = ([DateTime]::ParseExact($Case.CreatedOn, 'yyyyMMdd', $null)).ToString('o')
        }
        catch {
            $Case.LastUpdated = (Get-Date).ToString('o')
        }
    }

    $updateStamp = Get-UpdateStampFromIso $Case.LastUpdated
    $expected = Get-CaseFolderName -Date $Case.CreatedOn -Company $Case.Company -Support $Case.SupportNumber -Status $Case.Status -Updated $updateStamp
    if ([string]::IsNullOrWhiteSpace($expected)) { return $Case }
    if ([string]::Equals($Case.FolderName, $expected, [System.StringComparison]::OrdinalIgnoreCase)) { return $Case }

    $currentPath = $Case.FolderPath
    $parentPath = Split-Path -Parent $currentPath
    $targetPath = Join-Path -Path $parentPath -ChildPath $expected

    if ([string]::Equals($currentPath, $targetPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        $Case.FolderName = $expected
        $Case.FolderPath = $targetPath
        return $Case
    }

    if (Test-Path -Path $targetPath) {
        Write-Log ("Cannot rename {0} because {1} already exists." -f $currentPath, $targetPath)
        return $Case
    }

    try {
        if ($script:openedFolders.Contains($currentPath)) {
            Close-ExplorerWindow -Path $currentPath
            $script:openedFolders.Remove($currentPath) | Out-Null
        }
        Move-Item -LiteralPath $currentPath -Destination $targetPath -ErrorAction Stop
        $Case.FolderName = $expected
        $Case.FolderPath = $targetPath
        Write-Log ("Renamed folder {0} -> {1}" -f $currentPath, $targetPath)
    }
    catch {
        Write-Log ("Rename failed: {0}" -f $_.Exception.Message)
    }
    return $Case
}

function Ensure-NoteFiles {
    param(
        $Case,
        [bool]$ApplyUserFilenames = $false
    )
    if (-not $Case) { return $Case }

    if (-not (Test-Path -Path $Case.FolderPath)) { return $Case }

    foreach ($editor in $noteEditors) {
        $folderPath = $Case.FolderPath

        $boxText = if ($editor.FileBox) { $editor.FileBox.Text.Trim() } else { '' }
        $hasUserPreference = -not [string]::IsNullOrWhiteSpace($boxText)
        $desiredInput = $null
        if ($ApplyUserFilenames -and $hasUserPreference) {
            $desiredInput = Sanitize-FileName $boxText
        }

        $defaultName = Get-NoteFileName -BaseName $editor.BaseName -Support $Case.SupportNumber
        $defaultPath = Join-Path -Path $folderPath -ChildPath $defaultName

        $targetName = $null
        if ($desiredInput) {
            $targetName = $desiredInput
        }
        elseif ($ApplyUserFilenames -or -not $hasUserPreference) {
            $targetName = $defaultName
        }

        $targetPath = if ($targetName) { Join-Path -Path $folderPath -ChildPath $targetName } else { $null }

        $actualPath = $null
        if ($editor.PSObject.Properties['SelectedPath']) { $editor.SelectedPath = $null }
        if ($editor.FileBox) { $editor.FileBox.Tag = $null }

        if (-not $actualPath) {
            $currentBoxName = Sanitize-FileName $boxText

            $candidates = @()
            if ($desiredInput) { $candidates += Join-Path -Path $folderPath -ChildPath $desiredInput }
            if ($currentBoxName) { $candidates += Join-Path -Path $folderPath -ChildPath $currentBoxName }
            if ($targetPath) { $candidates += $targetPath } else { $candidates += $defaultPath }

            $basePattern = Remove-InvalidPathChars -Text $editor.BaseName
            if ([string]::IsNullOrWhiteSpace($basePattern)) { $basePattern = $editor.BaseName }
            $wildcards = Get-ChildItem -Path $folderPath -Filter ("{0}*.txt" -f $basePattern) -File -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending

            foreach ($path in ($candidates | Select-Object -Unique)) {
                if (Test-Path -LiteralPath $path) { $actualPath = (Resolve-Path -LiteralPath $path).Path; break }
            }
            if (-not $actualPath -and $wildcards) {
                $actualPath = $wildcards[0].FullName
            }
        }

        if (-not $actualPath) {
            $targetForCreate = if ($targetPath) { $targetPath } else { $defaultPath }
            try {
                Set-Content -LiteralPath $targetForCreate -Value '' -Encoding UTF8
                $actualPath = $targetForCreate
            }
            catch {
                Write-Log ("Note create failed: {0}" -f $_.Exception.Message)
                continue
            }
        }

        if ($targetPath -and $actualPath -and -not [string]::Equals($targetPath, $actualPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            if (Test-Path -LiteralPath $targetPath) {
                [void](Backup-CaseNoteFile -FolderPath $folderPath -FilePath $targetPath)
            }

            if (Test-Path -LiteralPath $actualPath) {
                try {
                    Move-Item -LiteralPath $actualPath -Destination $targetPath -ErrorAction Stop
                    $actualPath = $targetPath
                }
                catch {
                    Write-Log ("Note rename failed: {0}" -f $_.Exception.Message)
                }
            }
        }

        if ($editor.FileBox -and $actualPath) {
            $appliedLeaf = Split-Path -Path $actualPath -Leaf
            $editor.FileBox.Text = $appliedLeaf
            $editor.FileBox.Tag = $actualPath
            if ($editor.SelectedLabel) {
                $sourceLeaf = ''
                if ($editor.PSObject.Properties['SelectedSourceName']) {
                    $sourceLeaf = $editor.SelectedSourceName
                }
                if ([string]::IsNullOrWhiteSpace($sourceLeaf)) {
                    $editor.SelectedLabel.Text = "{0} {1}" -f $(U "\u9069\u7528\u5148:"), $appliedLeaf
                }
                else {
                    $editor.SelectedLabel.Text = "{0} {1} {2} {3}" -f $(U "\u9078\u629e\u5143:"), $sourceLeaf, $(U "\u2192 \u9069\u7528\u5148:"), $appliedLeaf
                }
            }
        }
        if ($editor.PSObject.Properties['LastSourcePath']) {
            if ($actualPath -and (Test-Path -LiteralPath $actualPath)) {
                try {
                    $editor.LastSourcePath = (Resolve-Path -LiteralPath $actualPath -ErrorAction Stop).Path
                }
                catch {
                    $editor.LastSourcePath = $actualPath
                }
            }
        }
        if ($editor.PSObject.Properties['SelectedSourceName']) { $editor.SelectedSourceName = '' }
    }
    Update-NoteEditorButtons -HasCase:$true
    return $Case
}

function Open-Folder {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -Path $Path)) { return }
    if ($script:openedFolders.Contains($Path)) {
        if (-not (Test-FolderWindowOpen -Path $Path)) {
            $script:openedFolders.Remove($Path) | Out-Null
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                ((U "\u3053\u306e\u6848\u4ef6\u306f\u65e2\u306b\u958b\u3044\u3066\u3044\u307e\u3059: {0}") -f $Path),
                $(U "\u60c5\u5831"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }
    }
    Start-Process explorer.exe -ArgumentList $Path | Out-Null
    [void]$script:openedFolders.Add($Path)
}

function Create-CaseFolder {
    if (-not (Ensure-BasePathReady)) { return }
    Update-Preview
    if ($textPreview.Text.StartsWith('(')) {
        [System.Windows.Forms.MessageBox]::Show(
            $(U "\u5fc5\u9808\u9805\u76ee\u3092\u78ba\u8a8d\u3057\u3066\u304f\u3060\u3055\u3044\u3002"),
            $(U "\u30a8\u30e9\u30fc"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $basePath = $textBasePath.Text
    $targetRoot = $basePath
    $category = $comboCategory.Text
    if ($category -and $category -ne $(U "(\u30d9\u30fc\u30b9\u76f4\u4e0b)")) {
        $targetRoot = Join-Path -Path $basePath -ChildPath $category
        if (-not (Test-Path -Path $targetRoot)) {
            try {
                New-Item -Path $targetRoot -ItemType Directory -ErrorAction Stop | Out-Null
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show(
                    ((U "\u4fdd\u5b58\u5148\u30d5\u30a9\u30eb\u30c0\u3092\u4f5c\u6210\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002\u8a73\u7d30: {0}") -f $_.Exception.Message),
                    $(U "\u30a8\u30e9\u30fc"),
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                ) | Out-Null
                return
            }
        }
    }

    $enteredSupport = $textSupportNo.Text.Trim()
    $normalizedSupport = Normalize-SupportNumber $enteredSupport
    if (-not [string]::IsNullOrWhiteSpace($normalizedSupport)) {
        $casesWithNormalized = @(
            $script:caseIndex |
            ForEach-Object { Ensure-NormalizedSupport -Case $_ }
        )
        $script:caseIndex = $casesWithNormalized
        $existing = $casesWithNormalized |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.NormalizedSupport) -and $_.NormalizedSupport -eq $normalizedSupport } |
            Sort-Object -Property @{ Expression = { $_.LastUpdated }; Descending = $true } |
            Select-Object -First 1

        if (-not $existing) {
            $existing = Search-FolderBySupportNumber -SupportNumber $enteredSupport -BasePath $basePath
            if (-not $existing -and $normalizedSupport -ne $enteredSupport) {
                $existing = Search-FolderBySupportNumber -SupportNumber $normalizedSupport -BasePath $basePath
            }
            if ($existing) {
                $existing = Ensure-CaseFolderName -Case $existing
                Update-CaseIndexEntry -Entry $existing
            }
        }

        if ($existing) {
            Set-CurrentCase -Case $existing
            [System.Windows.Forms.MessageBox]::Show(
                ((U "\u30b5\u30dd\u30fc\u30c8\u756a\u53f7 {0} (\u65e2\u5b58: {1}) \u306e\u6848\u4ef6\u306f\u65e2\u306b\u5b58\u5728\u3057\u307e\u3059\u3002\u5c65\u6b74\u306e\u6848\u4ef6\u3092\u518d\u5229\u7528\u3057\u3066\u304f\u3060\u3055\u3044\u3002") -f $enteredSupport, $existing.SupportNumber),
                $(U "\u60c5\u5831"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            Refresh-History
            return
        }
    }

    $folderPath = Join-Path -Path $targetRoot -ChildPath $textPreview.Text
    if (Test-Path -Path $folderPath) {
        [System.Windows.Forms.MessageBox]::Show(
            $(U "\u540c\u540d\u306e\u30d5\u30a9\u30eb\u30c0\u304c\u65e2\u306b\u5b58\u5728\u3057\u307e\u3059\u3002"),
            $(U "\u30a8\u30e9\u30fc"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    try {
        New-Item -Path $folderPath -ItemType Directory -ErrorAction Stop | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            ((U "\u30d5\u30a9\u30eb\u30c0\u3092\u4f5c\u6210\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002\u8a73\u7d30: {0}") -f $_.Exception.Message),
            $(U "\u30a8\u30e9\u30fc"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    foreach ($editor in $noteEditors) {
        $fileName = Get-NoteFileName -BaseName $editor.BaseName -Support $textSupportNo.Text
        $filePath = Join-Path -Path $folderPath -ChildPath $fileName
        try {
            [System.IO.File]::WriteAllText($filePath, $editor.TextBox.Text, [System.Text.Encoding]::UTF8)
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                ((U "\u30d5\u30a1\u30a4\u30eb {0} \u3092\u4f5c\u6210\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002\u8a73\u7d30: {1}") -f $fileName, $_.Exception.Message),
                $(U "\u30a8\u30e9\u30fc"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
            return
        }
    }

    $entry = [PSCustomObject]@{
        SupportNumber = $textSupportNo.Text.Trim()
        Company       = $textCompany.Text.Trim()
        Status        = Normalize-Status $comboStatus.Text
        CreatedOn     = $pickerDate.Value.ToString('yyyyMMdd')
        FolderName    = Split-Path -Path $folderPath -Leaf
        FolderPath    = $folderPath
        LastUpdated   = (Get-Date).ToString('o')
    }

    $entry = Ensure-NoteFiles -Case $entry -ApplyUserFilenames:$true
    $entry = Ensure-CaseFolderName -Case $entry
    Set-CurrentCase -Case $entry
    $entry = $script:currentCase

    [System.Windows.Forms.MessageBox]::Show(
        ((U "\u30d5\u30a9\u30eb\u30c0\u3092\u4f5c\u6210\u3057\u307e\u3057\u305f: {0}") -f $entry.FolderPath),
        $(U "\u5b8c\u4e86"),
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null

    if ($checkOpenAfter.Checked) {
        Open-Folder -Path $entry.FolderPath
    }
}

function New-EditorAttachmentFolder {
    param(
        $Editor,
        [switch]$ShowMessage
    )

    if (-not $Editor) { return }

    if (-not (Ensure-CurrentCaseContext -Silent)) {
        [System.Windows.Forms.MessageBox]::Show(
            $(U "\u5c65\u6b74\u304b\u3089\u6848\u4ef6\u3092\u9078\u629e\u3057\u3066\u304f\u3060\u3055\u3044\u3002"),
            $(U "\u8b66\u544a"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    if (-not $script:currentCase -or -not (Test-Path -LiteralPath $script:currentCase.FolderPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            $(U "\u5bfe\u8c61\u30d5\u30a9\u30eb\u30c0\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002"),
            $(U "\u30a8\u30e9\u30fc"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    $case = $script:currentCase
    $caseFolder = $case.FolderPath

    $rawBase = if ($Editor.LabelText) { $Editor.LabelText } elseif ($Editor.BaseName) { $Editor.BaseName } else { '' }
    if ([string]::IsNullOrWhiteSpace($rawBase)) { $rawBase = $(U "\u6dfb\u4ed8") }
    $baseName = Remove-InvalidPathChars -Text $rawBase
    if ([string]::IsNullOrWhiteSpace($baseName)) { $baseName = 'attachment' }

    $today = (Get-Date).ToString('yyyyMMdd')
    $prefix = '{0}_添付_{1}' -f $baseName, $today
    $searchPattern = "{0}_添付_{1}_*" -f $baseName, $today
    $escapedBase = [System.Text.RegularExpressions.Regex]::Escape($baseName)
    $regexPattern = '^' + $escapedBase + '_添付_' + $today + '_(?<order>\d+)$'
    $regex = [System.Text.RegularExpressions.Regex]::new($regexPattern)

    $nextIndex = 1
    try {
        $existing = Get-ChildItem -Path $caseFolder -Directory -Filter $searchPattern -ErrorAction SilentlyContinue
    }
    catch {
        $existing = @()
    }

    foreach ($dir in $existing) {
        $match = $regex.Match($dir.Name)
        if ($match.Success) {
            $orderValue = 0
            if ([int]::TryParse($match.Groups['order'].Value, [ref]$orderValue)) {
                if ($orderValue -ge $nextIndex) {
                    $nextIndex = $orderValue + 1
                }
            }
        }
    }

    $folderName = "{0}_{1}" -f $prefix, $nextIndex
    $targetPath = Join-Path -Path $caseFolder -ChildPath $folderName
    while (Test-Path -LiteralPath $targetPath) {
        $nextIndex++
        $folderName = "{0}_{1}" -f $prefix, $nextIndex
        $targetPath = Join-Path -Path $caseFolder -ChildPath $folderName
    }

    try {
        New-Item -Path $targetPath -ItemType Directory -ErrorAction Stop | Out-Null
        Write-Log ("Created attachment folder: {0}" -f $targetPath)
    }
    catch {
        Write-Log ("Attachment folder creation failed: {0}" -f $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show(
            ((U "\u30d5\u30a9\u30eb\u30c0\u3092\u4f5c\u6210\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002\u8a73\u7d30: {0}") -f $_.Exception.Message),
            $(U "\u30a8\u30e9\u30fc"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    if ($Editor.SelectedLabel) {
        $Editor.SelectedLabel.Text = "{0} {1}" -f $(U "\u4f5c\u6210\u5148:"), $folderName
    }
    if ($Editor.PSObject.Properties['SelectedPath']) {
        $Editor.SelectedPath = $targetPath
    }

    $script:currentCase.LastUpdated = (Get-Date).ToString('o')
    Update-CaseIndexEntry -Entry $script:currentCase
    Update-NoteEditorButtons -HasCase:$true

    if ($ShowMessage) {
        [System.Windows.Forms.MessageBox]::Show(
            ((U "\u30d5\u30a9\u30eb\u30c0\u3092\u4f5c\u6210\u3057\u307e\u3057\u305f\u3002`n\u4f5c\u6210\u5148: {0}") -f $targetPath),
            $(U "\u5b8c\u4e86"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
}

function Append-NoteText {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw [System.ArgumentException]::new('Path must not be empty.')
    }

    $fs = $null
    try {
        $fs = New-Object System.IO.FileStream(
            $Path,
            [System.IO.FileMode]::OpenOrCreate,
            [System.IO.FileAccess]::Write,
            [System.IO.FileShare]::ReadWrite
        )
        [void]$fs.Seek(0, [System.IO.SeekOrigin]::End)
        $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$Value)
        $fs.Write($bytes, 0, $bytes.Length)
    }
    finally {
        if ($fs) {
            $fs.Dispose()
        }
    }
}

function Append-Notes {
    if (-not $script:currentCase) {
        [System.Windows.Forms.MessageBox]::Show(
            $(U "\u5c65\u6b74\u304b\u3089\u6848\u4ef6\u3092\u9078\u629e\u3057\u3066\u304f\u3060\u3055\u3044\u3002"),
            $(U "\u8b66\u544a"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }
    if (-not (Test-Path -Path $script:currentCase.FolderPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            $(U "\u5bfe\u8c61\u30d5\u30a9\u30eb\u30c0\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002"),
            $(U "\u30a8\u30e9\u30fc"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    $folderPath = $script:currentCase.FolderPath
    $now = Get-Date
    $headerTimestamp = $now.ToString('yyyy/MM/dd HH:mm:ss')
    $statusForHeader = Normalize-Status $comboStatus.Text
    $appended = $false

    foreach ($editor in $noteEditors) {
        $content = $editor.TextBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($content)) { continue }

        $headerLine = if ([string]::IsNullOrWhiteSpace($statusForHeader)) {
            ((U "*****\u8ffd\u8a18\u90e8_{0}******") -f $headerTimestamp)
        }
        else {
            ((U "*****\u8ffd\u8a18\u90e8_{0}({1})******") -f $headerTimestamp, $statusForHeader)
        }
        $entryText = "`r`n{0}`r`n{1}`r`n" -f $headerLine, $content
        $fileName = Get-NoteFileName -BaseName $editor.BaseName -Support $textSupportNo.Text
        $filePath = Join-Path -Path $folderPath -ChildPath $fileName
        try {
            Append-NoteText -Path $filePath -Value $entryText
            $appended = $true
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                ((U "\u30d5\u30a1\u30a4\u30eb {0} \u3078\u306e\u8ffd\u8a18\u306b\u5931\u6557\u3057\u307e\u3057\u305f\u3002\u8a73\u7d30: {1}") -f $fileName, $_.Exception.Message),
                $(U "\u30a8\u30e9\u30fc"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
            return
        }
        finally {
            $editor.TextBox.Clear()
        }
    }

    $updatedEntry = [PSCustomObject]@{
        SupportNumber = $textSupportNo.Text.Trim()
        Company       = $textCompany.Text.Trim()
        Status        = Normalize-Status $comboStatus.Text
        CreatedOn     = $pickerDate.Value.ToString('yyyyMMdd')
        FolderName    = Split-Path -Path $folderPath -Leaf
        FolderPath    = $folderPath
        LastUpdated   = (Get-Date).ToString('o')
    }
    $updatedEntry = Ensure-NoteFiles -Case $updatedEntry -ApplyUserFilenames:$true
    $updatedEntry = Ensure-CaseFolderName -Case $updatedEntry
    Set-CurrentCase -Case $updatedEntry
    $updatedEntry = $script:currentCase

    if ($appended) {
        Close-CurrentCaseFolder -ShowMessage:$false
    }

    if ($updatedEntry.Status -eq $(U "\u30af\u30ed\u30fc\u30ba")) {
        [System.Windows.Forms.MessageBox]::Show(
            $(U "\u30af\u30ed\u30fc\u30ba\u306b\u66f4\u65b0\u3057\u307e\u3057\u305f\u3002\u5fc5\u8981\u306b\u5fdc\u3058\u3066\u30af\u30ed\u30fc\u30ba\u7528\u30d5\u30a9\u30eb\u30c0\u3078\u79fb\u52d5\u3057\u3066\u304f\u3060\u3055\u3044\u3002"),
            $(U "\u5b8c\u4e86"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        Set-CurrentCase -Case $null
        return
    }

    if ($appended) {
        [System.Windows.Forms.MessageBox]::Show(
            $(U "\u8ffd\u8a18\u3057\u307e\u3057\u305f\u3002"),
            $(U "\u5b8c\u4e86"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
}

function Load-CaseFromSupportNumber {
    param(
        [string]$SupportNumber = $null,
        [switch]$Silent
    )

    if (-not (Ensure-BasePathReady)) { return $null }
    $basePath = $textBasePath.Text.Trim()
    $input = if ($SupportNumber) { $SupportNumber.Trim() } else { $textSearchSupport.Text.Trim() }
    if ([string]::IsNullOrWhiteSpace($input)) {
        if (-not $Silent) {
            [System.Windows.Forms.MessageBox]::Show(
                $(U "\u30b5\u30dd\u30fc\u30c8\u756a\u53f7\u3092\u5165\u529b\u3057\u3066\u304f\u3060\u3055\u3044\u3002"),
                $(U "\u8b66\u544a"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
        }
        return $null
    }

    $needsReload = $false
    if (-not $script:caseIndex -or $script:caseIndex.Count -eq 0) {
        $needsReload = $true
    }
    elseif ($script:lastIndexBasePath -and -not [string]::Equals($script:lastIndexBasePath, $basePath, [System.StringComparison]::OrdinalIgnoreCase)) {
        $needsReload = $true
    }
    else {
        $indexFile = Get-IndexFilePath -BasePath $basePath
        if (Test-Path -LiteralPath $indexFile) {
            try {
                $lastWriteUtc = (Get-Item -LiteralPath $indexFile -ErrorAction Stop).LastWriteTimeUtc
                if (-not $script:lastIndexTimestamp -or $script:lastIndexTimestamp -lt $lastWriteUtc) {
                    $needsReload = $true
                }
            }
            catch {
                $needsReload = $true
            }
        }
        elseif (-not $script:lastIndexTimestamp) {
            $needsReload = $true
        }
    }

    if ($needsReload) {
        Load-CaseIndexForBasePath
    }

    $casesWithNormalized = @(
        $script:caseIndex |
        ForEach-Object { Ensure-NormalizedSupport -Case $_ }
    )
    $script:caseIndex = $casesWithNormalized
    $normalizedInput = Normalize-SupportNumber $input
    $escapedInput = [System.Management.Automation.WildcardPattern]::Escape($input)
    $escapedNormalized = if ([string]::IsNullOrWhiteSpace($normalizedInput)) { $null } else { [System.Management.Automation.WildcardPattern]::Escape($normalizedInput) }

    $match = $null
    if (-not [string]::IsNullOrWhiteSpace($normalizedInput)) {
        $match = $casesWithNormalized |
            Where-Object {
                $candidate = $null
                if ($_.PSObject.Properties['NormalizedSupport']) {
                    $candidate = $_.NormalizedSupport
                }
                if ([string]::IsNullOrWhiteSpace($candidate) -and $_.PSObject.Properties['SupportNumber']) {
                    $candidate = Normalize-SupportNumber $_.SupportNumber
                }
                -not [string]::IsNullOrWhiteSpace($candidate) -and $candidate -eq $normalizedInput
            } |
            Sort-Object -Property @{ Expression = { $_.LastUpdated }; Descending = $true } |
            Select-Object -First 1
    }

    if (-not $match -and $escapedNormalized) {
        $match = $casesWithNormalized |
            Where-Object {
                $candidate = $null
                if ($_.PSObject.Properties['NormalizedSupport']) {
                    $candidate = $_.NormalizedSupport
                }
                if ([string]::IsNullOrWhiteSpace($candidate) -and $_.PSObject.Properties['SupportNumber']) {
                    $candidate = Normalize-SupportNumber $_.SupportNumber
                }
                -not [string]::IsNullOrWhiteSpace($candidate) -and $candidate -like "*$escapedNormalized*"
            } |
            Sort-Object -Property @{ Expression = { $_.LastUpdated }; Descending = $true } |
            Select-Object -First 1
    }

    if (-not $match) {
        $match = $casesWithNormalized |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace($_.SupportNumber) -and
                $_.SupportNumber -like "*$escapedInput*"
            } |
            Sort-Object -Property @{ Expression = { $_.LastUpdated }; Descending = $true } |
            Select-Object -First 1
    }

    if (-not $match) {
        $match = $casesWithNormalized |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace($_.FolderName) -and
                $_.FolderName -like "*$escapedInput*"
            } |
            Sort-Object -Property @{ Expression = { $_.LastUpdated }; Descending = $true } |
            Select-Object -First 1
    }

    if ($match) {
        try {
            $dirInfo = Get-Item -LiteralPath $match.FolderPath -ErrorAction Stop
            if ($dirInfo -and $dirInfo.PSIsContainer) {
                $parsed = Try-ParseCaseFromFolder -Directory $dirInfo -SupportNumber $match.SupportNumber
                if ($parsed) {
                    $match = $parsed
                }
            }
        }
        catch {
            Write-Log ("Indexed folder not found for {0}: {1}" -f $input, $_.Exception.Message)
            $match = $null
        }
    }

    if (-not $match) {
        $match = Search-FolderBySupportNumber -SupportNumber $input -BasePath $basePath
        if ($match) {
            Update-CaseIndexEntry -Entry $match
        }
    }

    if (-not $match -and -not [string]::IsNullOrWhiteSpace($normalizedInput) -and $normalizedInput -ne $input) {
        $match = Search-FolderBySupportNumber -SupportNumber $normalizedInput -BasePath $basePath
        if ($match) {
            Update-CaseIndexEntry -Entry $match
        }
    }

    if (-not $match) {
        if (-not $Silent) {
            [System.Windows.Forms.MessageBox]::Show(
                ((U "\u30b5\u30dd\u30fc\u30c8\u756a\u53f7 {0} \u306e\u6848\u4ef6\u306f\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3067\u3057\u305f\u3002") -f $input),
                $(U "\u60c5\u5831"),
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        }
        return $null
    }

    $match = Ensure-CaseFolderName -Case $match
    Update-CaseIndexEntry -Entry $match
    Set-CurrentCase -Case $match
    if (-not $Silent) {
        Open-Folder -Path $match.FolderPath
    }
    return $script:currentCase
}

$buttonBrowse.add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.SelectedPath = $textBasePath.Text
    if ($dialog.ShowDialog() -eq 'OK') {
        $textBasePath.Text = $dialog.SelectedPath
        Refresh-Categories
        Load-CaseIndexForBasePath
        Set-CurrentCase -Case $null
    }
})

$buttonReloadCategory.add_Click({
    if (-not (Ensure-BasePathReady)) { return }

    if ($null -eq $script:currentCase) {
        Refresh-Categories
        [System.Windows.Forms.MessageBox]::Show(
            $(U "\u5c65\u6b74\u304b\u3089\u6848\u4ef6\u3092\u9078\u629e\u3057\u3066\u304f\u3060\u3055\u3044\u3002"),
            $(U "\u8b66\u544a"),
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $script:currentCase.Company = $textCompany.Text.Trim()
    $script:currentCase.SupportNumber = $textSupportNo.Text.Trim()
    $script:currentCase.Status = Normalize-Status $comboStatus.Text
    $script:currentCase.CreatedOn = $pickerDate.Value.ToString('yyyyMMdd')
    $script:currentCase.LastUpdated = (Get-Date).ToString('o')

    $script:currentCase = Ensure-NoteFiles -Case $script:currentCase -ApplyUserFilenames:$true
    Set-CurrentCase -Case $script:currentCase
    Refresh-History
    Refresh-Categories
    Update-Preview

    [System.Windows.Forms.MessageBox]::Show(
        $(U "\u30d5\u30a9\u30eb\u30c0\u540d\u3092\u66f4\u65b0\u3057\u307e\u3057\u305f\u3002"),
        $(U "\u5b8c\u4e86"),
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
})
$buttonSearch.add_Click({ [void](Load-CaseFromSupportNumber) })
$buttonNew.add_Click({ Set-CurrentCase -Case $null })
$buttonOpenFolder.add_Click({
    if ($script:currentCase) {
        Open-Folder -Path $script:currentCase.FolderPath
    }
})
$buttonCloseFolder.add_Click({ Close-CurrentCaseFolder -ShowMessage:$true })
$buttonCreate.add_Click({ Create-CaseFolder })
$buttonAppend.add_Click({ Append-Notes })

$textBasePath.add_TextChanged({ Update-Preview })
$pickerDate.add_ValueChanged({ Update-Preview })
$textCompany.add_TextChanged({ Update-Preview })
$textSupportNo.add_TextChanged({ Update-Preview })
    $comboStatus.add_TextChanged({ Update-Preview })

$buttonReloadHistory.add_Click({
    if (-not (Ensure-BasePathReady)) { return }
    Load-CaseIndexForBasePath
    Refresh-History
    [System.Windows.Forms.MessageBox]::Show(
        $(U "\u5c65\u6b74\u3092\u66f4\u65b0\u3057\u307e\u3057\u305f\u3002"),
        $(U "\u5b8c\u4e86"),
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
})

$comboHistory.add_SelectedIndexChanged({
    if ($comboHistory.SelectedIndex -lt 0) { return }
    $item = $comboHistory.SelectedItem
    if ($item -and $item.Data) {
        $case = Ensure-CaseFolderName -Case $item.Data
        Update-CaseIndexEntry -Entry $case
        Set-CurrentCase -Case $case
    }
})

$form.add_Shown({
    Write-Log "Before Application.Run"
    Refresh-Categories
    Load-CaseIndexForBasePath
    Update-Preview
    $form.TopMost = $false
    Write-Log "Form shown"
})

[System.Windows.Forms.Application]::EnableVisualStyles()
try {
    [System.Windows.Forms.Application]::Run($form)
    Write-Log "Application.Run returned"
}
catch {
    Write-Log ("Unhandled exception: {0}" -f $_.Exception.Message)
    [System.Windows.Forms.MessageBox]::Show(
        ((U "\u4e88\u671f\u305b\u306c\u30a8\u30e9\u30fc\u304c\u767a\u751f\u3057\u307e\u3057\u305f\u3002\u8a73\u7d30: {0}") -f $_.Exception.Message),
        $(U "\u30a8\u30e9\u30fc"),
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}



 
