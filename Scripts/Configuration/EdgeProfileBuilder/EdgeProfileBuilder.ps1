<#
.SYNOPSIS
    Visual WPF GUI builder for isolated Microsoft Edge profile shortcuts with auto .ico conversion.

.DESCRIPTION
    Displays a dark-themed visual thumbnail grid of bundled default colored Microsoft Edge icons
    and supports uploading or dragging-and-dropping custom images (.png, .jpg, .jpeg, .bmp, .gif, .ico).
    Automatically converts non-ICO images into crisp multi-size .ico files (256, 64, 48, 32, and 16px)
    while preserving aspect ratio and transparency, creates isolated Edge user-data-dir profiles
    (or standard --profile-directory flags), and stamps a unique System.AppUserModel.ID onto each .lnk
    shortcut so Windows 11 allows all selected profiles to be pinned independently to the taskbar.

.PARAMETER OutputPath
    Root folder where Shortcuts\, Profiles\, and ConvertedIcons\ are created.
    Defaults to %LOCALAPPDATA%\EdgeProfiles.

.PARAMETER IconFolder
    Optional initial icon folder to load instead of the bundled .\assets default icon pack.

.EXAMPLE
    .\EdgeProfileBuilder.ps1
    Launches the interactive WPF GUI with the bundled 9-color Microsoft Edge icon pack.

.EXAMPLE
    .\EdgeProfileBuilder.ps1 -OutputPath "C:\EdgeProfiles"
    Launches the GUI with a custom output directory for profiles, icons, and shortcuts.

.NOTES
    Script Name   : EdgeProfileBuilder.ps1
    Author        : JP
    Version       : 1.0.0
    Created       : 25/11/26
    Last Modified : 25/11/26
    Log Path      : C:\Logs\EdgeProfileBuilder\EdgeProfileBuilder.log
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = (Join-Path -Path $env:LOCALAPPDATA -ChildPath 'EdgeProfiles'),

    [Parameter(Mandatory = $false)]
    [string]$IconFolder = ''
)

$ErrorActionPreference = 'Stop'

$Script:RepoRawBaseUrl = 'https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Configuration/EdgeProfileBuilder'
$Script:CacheRoot = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'EdgeProfileBuilder\cache'

# ---------------------------------------------------------------------------
# 1. Enforce STA Threading for WPF (Supports Local & In-Memory irm | iex)
# ---------------------------------------------------------------------------
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $relaunchPath = if ($PSScriptRoot -and $MyInvocation.MyCommand.Name) {
        Join-Path -Path $PSScriptRoot -ChildPath $MyInvocation.MyCommand.Name
    } elseif ($PSCommandPath) {
        $PSCommandPath
    } else {
        $null
    }

    if (-not $relaunchPath -or -not (Test-Path -Path $relaunchPath -PathType Leaf)) {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
        if (-not (Test-Path -Path $Script:CacheRoot)) {
            New-Item -Path $Script:CacheRoot -ItemType Directory -Force | Out-Null
        }
        $relaunchPath = Join-Path -Path $Script:CacheRoot -ChildPath 'EdgeProfileBuilder.ps1'
        Invoke-WebRequest -Uri "$($Script:RepoRawBaseUrl)/EdgeProfileBuilder.ps1" -OutFile $relaunchPath -UseBasicParsing
    }

    $argList = New-Object System.Collections.Generic.List[string]
    $argList.Add('-STA')
    $argList.Add('-ExecutionPolicy')
    $argList.Add('Bypass')
    $argList.Add('-NoProfile')
    $argList.Add('-File')
    $argList.Add("`"$relaunchPath`"")
    if ($PSBoundParameters.ContainsKey('OutputPath')) {
        $argList.Add('-OutputPath')
        $argList.Add("`"$OutputPath`"")
    }
    if ($PSBoundParameters.ContainsKey('IconFolder') -and -not [string]::IsNullOrWhiteSpace($IconFolder)) {
        $argList.Add('-IconFolder')
        $argList.Add("`"$IconFolder`"")
    }
    Start-Process -FilePath 'powershell.exe' -ArgumentList ($argList -join ' ')
    exit
}

function Initialize-RemoteAsset {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$CacheDir
    )
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

    $assetsDir = Join-Path -Path $CacheDir -ChildPath 'assets'
    foreach ($dir in @($CacheDir, $assetsDir)) {
        if (-not (Test-Path -Path $dir)) {
            New-Item -Path $dir -ItemType Directory -Force | Out-Null
        }
    }

    $prevProgress = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'
    try {
        foreach ($coreFile in @('EdgeProfileEngine.psm1', 'EdgeProfileBuilder.xaml')) {
            $destFile = Join-Path -Path $CacheDir -ChildPath $coreFile
            $url = "$($Script:RepoRawBaseUrl)/$coreFile"
            try {
                Invoke-WebRequest -Uri $url -OutFile $destFile -UseBasicParsing -ErrorAction Stop
            } catch {
                if (-not (Test-Path -Path $destFile -PathType Leaf)) {
                    throw "Failed to download required component '$coreFile' from '$url': $($_.Exception.Message)"
                }
                Write-Warning "Using cached '$coreFile' because download failed: $($_.Exception.Message)"
            }
        }

        $defaultIcons = @(
            'Edge_CyberNeon.png',
            'Edge_Gold.png',
            'Edge_Lime.png',
            'Edge_Magenta.png',
            'Edge_Orange.png',
            'Edge_Purple.png',
            'Edge_Rainbow.png',
            'Edge_Silver.png',
            'Edge_Teal.png'
        )
        foreach ($iconName in $defaultIcons) {
            $iconDest = Join-Path -Path $assetsDir -ChildPath $iconName
            if (-not (Test-Path -Path $iconDest -PathType Leaf)) {
                $iconUrl = "$($Script:RepoRawBaseUrl)/assets/$iconName"
                try {
                    Invoke-WebRequest -Uri $iconUrl -OutFile $iconDest -UseBasicParsing -ErrorAction Stop
                } catch {
                    Write-Warning "Could not download default icon '$iconName': $($_.Exception.Message)"
                }
            }
        }
    } finally {
        $ProgressPreference = $prevProgress
    }
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Drawing, System.Windows.Forms, System.Xaml

$localCandidate = if ($PSScriptRoot) {
    $PSScriptRoot
} elseif ($PSCommandPath) {
    Split-Path -Path $PSCommandPath -Parent
} else {
    $null
}

$hasLocalFiles = $localCandidate -and
    (Test-Path -Path (Join-Path -Path $localCandidate -ChildPath 'EdgeProfileEngine.psm1') -PathType Leaf) -and
    (Test-Path -Path (Join-Path -Path $localCandidate -ChildPath 'EdgeProfileBuilder.xaml') -PathType Leaf)

if ($hasLocalFiles) {
    $Script:ScriptRoot = $localCandidate
} else {
    $Script:ScriptRoot = $Script:CacheRoot
    Initialize-RemoteAsset -CacheDir $Script:ScriptRoot
}

$Script:DefaultAssetsDir = Join-Path -Path $Script:ScriptRoot -ChildPath 'assets'
$Script:XamlPath = Join-Path -Path $Script:ScriptRoot -ChildPath 'EdgeProfileBuilder.xaml'
$Script:EnginePath = Join-Path -Path $Script:ScriptRoot -ChildPath 'EdgeProfileEngine.psm1'
$Script:SupportedExtensions = @('.png', '.ico', '.jpg', '.jpeg', '.bmp', '.gif')
$Script:LogTextBox = $null

Remove-Module -Name 'EdgeProfileEngine' -ErrorAction SilentlyContinue
$engineSource = Get-Content -Path $Script:EnginePath -Raw -Encoding UTF8
$engineModule = New-Module -Name 'EdgeProfileEngine' -ScriptBlock ([scriptblock]::Create($engineSource))
Import-Module -ModuleInfo $engineModule -Force -Global | Out-Null

# ---------------------------------------------------------------------------
# 2. Diagnostic Logging (UK Format dd/MM/yy HH:mm:ss)
# ---------------------------------------------------------------------------
function Invoke-UiRefresh {
    [CmdletBinding()]
    param()
    if ([System.Windows.Application]::Current -or $null -ne $Script:LogTextBox) {
        $frame = New-Object System.Windows.Threading.DispatcherFrame
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke(
            [System.Windows.Threading.DispatcherPriority]::Background,
            [Action[System.Windows.Threading.DispatcherFrame]]{ param($f) $f.Continue = $false },
            $frame
        ) | Out-Null
        [System.Windows.Threading.Dispatcher]::PushFrame($frame)
    }
}

function Write-ToolLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')][string]$Level = 'INFO'
    )
    $logDir = 'C:\Logs\EdgeProfileBuilder'
    if (-not (Test-Path -Path $logDir)) {
        try {
            New-Item -Path $logDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        } catch {
            $logDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'Logs\EdgeProfileBuilder'
            if (-not (Test-Path -Path $logDir)) {
                New-Item -Path $logDir -ItemType Directory -Force | Out-Null
            }
        }
    }

    $fileTimestamp = Get-Date -Format 'dd/MM/yy HH:mm:ss'
    $uiTimestamp = Get-Date -Format 'HH:mm:ss'
    $fileEntry = "[$fileTimestamp] [$Level] $Message"
    $uiEntry = "[$uiTimestamp] [$Level] $Message"

    try {
        Add-Content -Path (Join-Path -Path $logDir -ChildPath 'EdgeProfileBuilder.log') -Value $fileEntry -Encoding UTF8 -ErrorAction Stop
    } catch {
        Write-Verbose "Failed to append log entry: $($_.Exception.Message)"
    }

    if ($null -ne $Script:LogTextBox) {
        $Script:LogTextBox.AppendText("$uiEntry`r`n")
        $Script:LogTextBox.ScrollToEnd()
        Invoke-UiRefresh
    } else {
        Write-Verbose $fileEntry
    }
}

# ---------------------------------------------------------------------------
# 3. WPF GUI Controller
# ---------------------------------------------------------------------------
function Show-EdgeProfileBuilderGui {
    [CmdletBinding()]
    param()

    if (-not (Test-Path -Path $Script:XamlPath -PathType Leaf)) {
        throw "UI definition not found: '$($Script:XamlPath)'."
    }

    $edgeExe = Get-EdgeExecutablePath
    [xml]$xamlDoc = Get-Content -Path $Script:XamlPath -Raw -Encoding UTF8
    $xmlReader = New-Object System.Xml.XmlNodeReader $xamlDoc
    $window = [Windows.Markup.XamlReader]::Load($xmlReader)

    $workArea = [System.Windows.SystemParameters]::WorkArea
    $targetWidth = [Math]::Max(860, [Math]::Min(1440, $workArea.Width * 0.70))
    $targetHeight = [Math]::Max(640, [Math]::Min(900, $workArea.Height * 0.80))
    $window.Width = [Math]::Round([Math]::Min($targetWidth, $workArea.Width - 24))
    $window.Height = [Math]::Round([Math]::Min($targetHeight, $workArea.Height - 24))
    $window.Left = [Math]::Round($workArea.Left + (($workArea.Width - $window.Width) / 2))
    $window.Top = [Math]::Round($workArea.Top + (($workArea.Height - $window.Height) / 2))

    $txtSelectedBadge = $window.FindName('TxtSelectedBadge')
    $btnUploadIcons = $window.FindName('BtnUploadIcons')
    $btnLoadFolder = $window.FindName('BtnLoadFolder')
    $btnDefaultPack = $window.FindName('BtnDefaultPack')
    $btnSelectAll = $window.FindName('BtnSelectAll')
    $btnClearAll = $window.FindName('BtnClearAll')
    $dropZoneBorder = $window.FindName('DropZoneBorder')
    $icIcons = $window.FindName('IcIcons')
    $rbIsolated = $window.FindName('RbIsolated')
    $chkInPrivate = $window.FindName('ChkInPrivate')
    $chkDesktop = $window.FindName('ChkDesktop')
    $txtOutputRoot = $window.FindName('TxtOutputRoot')
    $btnBrowseOutput = $window.FindName('BtnBrowseOutput')
    $txtLog = $window.FindName('TxtLog')
    $btnOpenShortcuts = $window.FindName('BtnOpenShortcuts')
    $btnGenerate = $window.FindName('BtnGenerate')

    $Script:LogTextBox = $txtLog
    $txtOutputRoot.Text = $OutputPath

    $cards = New-Object 'System.Collections.ObjectModel.ObservableCollection[EdgeIconCardItem]'
    $icIcons.ItemsSource = $cards

    $updateSummary = {
        $selectedCount = @($cards | Where-Object { $_.IsSelected }).Count
        $txtSelectedBadge.Text = "$selectedCount Selected"
        if ($selectedCount -gt 0) {
            $plural = if ($selectedCount -eq 1) { 'Profile & Shortcut' } else { 'Profiles & Shortcuts' }
            $btnGenerate.Content = "Generate $selectedCount $plural"
            $btnGenerate.IsEnabled = $true
        } else {
            $btnGenerate.Content = 'Generate Shortcuts'
            $btnGenerate.IsEnabled = $false
        }
    }

    $addIconCard = {
        param(
            [string]$FilePath,
            [bool]$IsUploaded,
            [bool]$SelectOnAdd
        )
        if (-not (Test-Path -Path $FilePath -PathType Leaf)) { return $false }
        $ext = [System.IO.Path]::GetExtension($FilePath).ToLowerInvariant()
        if ($Script:SupportedExtensions -notcontains $ext) { return $false }

        $fullPath = [System.IO.Path]::GetFullPath($FilePath)
        $existing = $cards | Where-Object { $_.SourcePath -eq $fullPath } | Select-Object -First 1
        if ($null -ne $existing) {
            if ($SelectOnAdd) { $existing.IsSelected = $true }
            return $false
        }

        try {
            $thumb = New-ThumbnailBitmap -ImagePath $fullPath
            $fileName = [System.IO.Path]::GetFileName($fullPath)
            $card = New-Object EdgeIconCardItem
            $card.Id = [Guid]::NewGuid().ToString('N')
            $card.SourcePath = $fullPath
            $card.FileName = $fileName
            $card.ProfileName = Get-CleanProfileName -FileName $fileName
            $card.FormatBadge = if ($ext -eq '.ico') { 'ICO' } else { "$($ext.TrimStart('.').ToUpperInvariant()) -> ICO" }
            $card.SourceBadge = if ($IsUploaded) { 'UPLOADED' } else { 'DEFAULT PACK' }
            $card.IsUploaded = $IsUploaded
            $card.Thumbnail = $thumb
            $card.IsSelected = $SelectOnAdd
            $card.add_PropertyChanged({
                param($sender, $eventArgs)
                if ($eventArgs.PropertyName -eq 'IsSelected') {
                    & $updateSummary
                }
            })
            $cards.Add($card)
            return $true
        } catch {
            Write-ToolLog -Message "Skipped '$FilePath': $($_.Exception.Message)" -Level WARN
            return $false
        }
    }

    $loadFolderIntoGrid = {
        param(
            [string]$FolderPath,
            [bool]$AsUploaded,
            [bool]$KeepUploaded
        )
        if (-not (Test-Path -Path $FolderPath -PathType Container)) {
            Write-ToolLog -Message "Icon folder not found: '$FolderPath'." -Level WARN
            return
        }

        if ($KeepUploaded) {
            $toRemove = @($cards | Where-Object { -not $_.IsUploaded })
            foreach ($item in $toRemove) { [void]$cards.Remove($item) }
        } else {
            $cards.Clear()
        }

        $files = Get-ChildItem -Path $FolderPath -File |
            Where-Object { $Script:SupportedExtensions -contains $_.Extension.ToLowerInvariant() } |
            Sort-Object -Property @{ Expression = { if ($_.Extension.ToLowerInvariant() -eq '.ico') { 0 } else { 1 } } }, Name

        $added = 0
        foreach ($file in $files) {
            if (& $addIconCard -FilePath $file.FullName -IsUploaded $AsUploaded -SelectOnAdd $false) {
                $added++
            }
        }
        & $updateSummary
        Write-ToolLog -Message "Loaded $added icon(s) from '$FolderPath'." -Level INFO
    }

    $icIcons.AddHandler(
        [System.Windows.Controls.Primitives.ButtonBase]::ClickEvent,
        [System.Windows.RoutedEventHandler]{
            param($sender, $e)
            $btn = $e.OriginalSource -as [System.Windows.Controls.Button]
            if ($null -eq $btn -and ($e.OriginalSource -is [System.Windows.DependencyObject])) {
                $parent = [System.Windows.Media.VisualTreeHelper]::GetParent($e.OriginalSource)
                while ($null -ne $parent -and -not ($parent -is [System.Windows.Controls.Button])) {
                    $parent = [System.Windows.Media.VisualTreeHelper]::GetParent($parent)
                }
                $btn = $parent -as [System.Windows.Controls.Button]
            }
            if ($null -eq $btn -or [string]::IsNullOrWhiteSpace([string]$btn.Tag)) { return }

            $targetId = [string]$btn.Tag
            $targetCard = $cards | Where-Object { $_.Id -eq $targetId } | Select-Object -First 1
            if ($null -eq $targetCard) { return }

            if ($btn.Name -eq 'BtnRemoveCard') {
                [void]$cards.Remove($targetCard)
                & $updateSummary
                Write-ToolLog -Message "Removed uploaded icon '$($targetCard.ProfileName)'." -Level INFO
            } elseif ($btn.Name -eq 'BtnToggleCard') {
                $targetCard.IsSelected = -not $targetCard.IsSelected
            }
        }
    )

    $btnUploadIcons.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title = 'Select Custom Icon Images'
        $dlg.Multiselect = $true
        $dlg.Filter = 'Image & Icon Files|*.png;*.ico;*.jpg;*.jpeg;*.bmp;*.gif|All Files|*.*'
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $addedCount = 0
            foreach ($fileName in $dlg.FileNames) {
                if (& $addIconCard -FilePath $fileName -IsUploaded $true -SelectOnAdd $true) {
                    $addedCount++
                }
            }
            & $updateSummary
            Write-ToolLog -Message "Added and ticked $addedCount custom icon(s)." -Level SUCCESS
        }
    })

    $btnLoadFolder.Add_Click({
        $folderDlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDlg.Description = 'Select a folder containing .png, .jpg, or .ico files'
        $folderDlg.SelectedPath = if (Test-Path -Path $Script:DefaultAssetsDir) { $Script:DefaultAssetsDir } else { $Script:ScriptRoot }
        if ($folderDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            & $loadFolderIntoGrid -FolderPath $folderDlg.SelectedPath -AsUploaded $true -KeepUploaded $true
        }
    })

    $btnDefaultPack.Add_Click({
        & $loadFolderIntoGrid -FolderPath $Script:DefaultAssetsDir -AsUploaded $false -KeepUploaded $true
    })

    $btnSelectAll.Add_Click({
        foreach ($item in $cards) { $item.IsSelected = $true }
        & $updateSummary
    })

    $btnClearAll.Add_Click({
        foreach ($item in $cards) { $item.IsSelected = $false }
        & $updateSummary
    })

    $dropZoneBorder.Add_DragOver({
        param($sender, $e)
        if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
            $e.Effects = [System.Windows.DragDropEffects]::Copy
        } else {
            $e.Effects = [System.Windows.DragDropEffects]::None
        }
        $e.Handled = $true
    })

    $dropZoneBorder.Add_Drop({
        param($sender, $e)
        if (-not $e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { return }
        $dropped = $e.Data.GetData([System.Windows.DataFormats]::FileDrop)
        $addedCount = 0
        foreach ($pathItem in $dropped) {
            if (Test-Path -Path $pathItem -PathType Container) {
                $subFiles = Get-ChildItem -Path $pathItem -File | Where-Object { $Script:SupportedExtensions -contains $_.Extension.ToLowerInvariant() }
                foreach ($sf in $subFiles) {
                    if (& $addIconCard -FilePath $sf.FullName -IsUploaded $true -SelectOnAdd $true) { $addedCount++ }
                }
            } elseif (& $addIconCard -FilePath $pathItem -IsUploaded $true -SelectOnAdd $true) {
                $addedCount++
            }
        }
        & $updateSummary
        if ($addedCount -gt 0) {
            Write-ToolLog -Message "Dropped and ticked $addedCount icon(s)." -Level SUCCESS
        } else {
            Write-ToolLog -Message 'No new supported image files found in drop.' -Level WARN
        }
    })

    $btnBrowseOutput.Add_Click({
        $outDlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $outDlg.Description = 'Choose root folder for Shortcuts, Profiles, and ConvertedIcons'
        if (Test-Path -Path $txtOutputRoot.Text) {
            $outDlg.SelectedPath = $txtOutputRoot.Text
        }
        if ($outDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtOutputRoot.Text = $outDlg.SelectedPath
        }
    })

    $btnOpenShortcuts.Add_Click({
        $shortcutsDir = Join-Path -Path $txtOutputRoot.Text.Trim() -ChildPath 'Shortcuts'
        if (-not (Test-Path -Path $shortcutsDir)) {
            New-Item -Path $shortcutsDir -ItemType Directory -Force | Out-Null
        }
        Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$shortcutsDir`""
    })

    $btnGenerate.Add_Click({
        $selectedCards = @($cards | Where-Object { $_.IsSelected })
        if ($selectedCards.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                'Please tick at least one icon in the grid.',
                'No Icons Selected',
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            ) | Out-Null
            return
        }

        if (-not $edgeExe) {
            $edgeExe = Get-EdgeExecutablePath
            if (-not $edgeExe) {
                Write-ToolLog -Message 'Microsoft Edge (msedge.exe) was not found in Program Files.' -Level ERROR
                [System.Windows.MessageBox]::Show(
                    'Microsoft Edge (msedge.exe) could not be found on this computer.',
                    'Microsoft Edge Not Found',
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                ) | Out-Null
                return
            }
        }

        $outRoot = $txtOutputRoot.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($outRoot)) {
            $outRoot = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'EdgeProfiles'
            $txtOutputRoot.Text = $outRoot
        }

        $btnGenerate.IsEnabled = $false
        $usedSlugs = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $createdCount = 0
        $profileIdx = 1

        try {
            Write-ToolLog -Message "Starting generation for $($selectedCards.Count) selected icon(s) into '$outRoot'..." -Level INFO
            foreach ($card in $selectedCards) {
                $displayName = if ([string]::IsNullOrWhiteSpace($card.ProfileName)) { $card.FileName } else { $card.ProfileName.Trim() }
                $baseSlug = Get-SafeFileSlug -Name $displayName
                $finalSlug = $baseSlug
                $suffix = 2
                while ($usedSlugs.Contains($finalSlug)) {
                    $finalSlug = "${baseSlug}_$suffix"
                    $suffix++
                }
                [void]$usedSlugs.Add($finalSlug)

                if ($finalSlug -ne $baseSlug) {
                    $displayName = "$displayName ($($suffix - 1))"
                    Write-ToolLog -Message "Duplicate name adjusted to '$displayName' [$finalSlug]." -Level WARN
                }

                $result = New-EdgeProfileShortcut `
                    -EdgeExe $edgeExe `
                    -ProfileName $displayName `
                    -SafeSlug $finalSlug `
                    -SourceIconPath $card.SourcePath `
                    -OutputRoot $outRoot `
                    -IsolatedMode ([bool]$rbIsolated.IsChecked) `
                    -ProfileIndex $profileIdx `
                    -UseInPrivate ([bool]$chkInPrivate.IsChecked) `
                    -CopyToDesktop ([bool]$chkDesktop.IsChecked)

                if ($result.ConvertedFromImage) {
                    Write-ToolLog -Message "Converted '$($card.FileName)' -> '$($result.IconPath | Split-Path -Leaf)' (multi-size ICO)." -Level INFO
                }
                Write-ToolLog -Message "Created '$($result.ShortcutPath | Split-Path -Leaf)' [AppID: $($result.AppId)]" -Level SUCCESS
                $createdCount++
                $profileIdx++
            }

            $shortcutsDir = Join-Path -Path $outRoot -ChildPath 'Shortcuts'
            Write-ToolLog -Message "Completed $createdCount profile shortcut(s) in '$shortcutsDir'." -Level SUCCESS
            Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$shortcutsDir`""
        } catch {
            Write-ToolLog -Message "Error generating shortcuts: $($_.Exception.Message)" -Level ERROR
            [System.Windows.MessageBox]::Show(
                "Failed while generating shortcuts:`r`n$($_.Exception.Message)",
                'Generation Error',
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            ) | Out-Null
        } finally {
            & $updateSummary
        }
    })

    if ($edgeExe) {
        Write-ToolLog -Message "Detected Microsoft Edge at '$edgeExe'." -Level INFO
    } else {
        Write-ToolLog -Message 'Microsoft Edge executable not detected yet; checked before generating.' -Level WARN
    }

    $initialSource = if (-not [string]::IsNullOrWhiteSpace($IconFolder)) { $IconFolder } else { $Script:DefaultAssetsDir }
    & $loadFolderIntoGrid -FolderPath $initialSource -AsUploaded $false -KeepUploaded $false

    [void]$window.ShowDialog()
}

if ($MyInvocation.InvocationName -ne '.') {
    try {
        Show-EdgeProfileBuilderGui
    } catch {
        Write-ToolLog -Message "Unhandled GUI error: $($_.Exception.Message)" -Level ERROR
        throw
    }
}
