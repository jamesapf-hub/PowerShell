<#
.SYNOPSIS
    Core icon conversion, thumbnail decoding, and Edge profile shortcut engine for EdgeProfileBuilder.
#>

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Drawing, System.Xaml

if (-not ('PheneoxEdgeShortcutAppId' -as [type])) {
    $appIdCSharp = @'
using System;
using System.Runtime.InteropServices;

public class PheneoxEdgeShortcutAppId
{
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct PROPERTYKEY
    {
        public Guid fmtid;
        public uint pid;
    }

    [StructLayout(LayoutKind.Explicit)]
    public struct PROPVARIANT
    {
        [FieldOffset(0)] public ushort vt;
        [FieldOffset(8)] public IntPtr pwszVal;
    }

    [ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPropertyStore
    {
        int GetCount(out uint cProps);
        int GetAt(uint iProp, out PROPERTYKEY pkey);
        int GetValue([In] ref PROPERTYKEY key, out PROPVARIANT pv);
        int SetValue([In] ref PROPERTYKEY key, [In] ref PROPVARIANT pv);
        int Commit();
    }

    [DllImport("shell32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern int SHGetPropertyStoreFromParsingName(
        string pszPath,
        IntPtr pbc,
        int flags,
        [In] ref Guid iid,
        [Out, MarshalAs(UnmanagedType.Interface)] out IPropertyStore ppv);

    [DllImport("ole32.dll")]
    private static extern int PropVariantClear(ref PROPVARIANT pvar);

    public static void SetShortcutAppId(string shortcutPath, string appId)
    {
        Guid iid = new Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99");
        IPropertyStore store;
        int hr = SHGetPropertyStoreFromParsingName(shortcutPath, IntPtr.Zero, 2, ref iid, out store);
        if (hr != 0)
        {
            Marshal.ThrowExceptionForHR(hr);
        }

        PROPERTYKEY pkey = new PROPERTYKEY();
        pkey.fmtid = new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3");
        pkey.pid = 5;

        PROPVARIANT pv = new PROPVARIANT();
        pv.vt = 31; // VT_LPWSTR
        pv.pwszVal = Marshal.StringToCoTaskMemUni(appId);

        try
        {
            int setHr = store.SetValue(ref pkey, ref pv);
            if (setHr != 0)
            {
                Marshal.ThrowExceptionForHR(setHr);
            }
            int commitHr = store.Commit();
            if (commitHr != 0)
            {
                Marshal.ThrowExceptionForHR(commitHr);
            }
        }
        finally
        {
            PropVariantClear(ref pv);
            if (store != null)
            {
                Marshal.FinalReleaseComObject(store);
            }
        }
    }
}
'@
    Add-Type -TypeDefinition $appIdCSharp -Language CSharp
}

if (-not ('EdgeIconCardItem' -as [type])) {
    $cardCSharp = @'
using System;
using System.ComponentModel;
using System.Windows.Media.Imaging;

public class EdgeIconCardItem : INotifyPropertyChanged
{
    private string _profileName;
    private bool _isSelected;

    public string Id { get; set; }
    public string SourcePath { get; set; }
    public string FileName { get; set; }
    public string FormatBadge { get; set; }
    public string SourceBadge { get; set; }
    public bool IsUploaded { get; set; }
    public BitmapSource Thumbnail { get; set; }

    public string ProfileName
    {
        get { return _profileName; }
        set
        {
            if (_profileName != value)
            {
                _profileName = value;
                OnPropertyChanged("ProfileName");
            }
        }
    }

    public bool IsSelected
    {
        get { return _isSelected; }
        set
        {
            if (_isSelected != value)
            {
                _isSelected = value;
                OnPropertyChanged("IsSelected");
            }
        }
    }

    public event PropertyChangedEventHandler PropertyChanged;

    protected void OnPropertyChanged(string name)
    {
        PropertyChangedEventHandler handler = PropertyChanged;
        if (handler != null)
        {
            handler(this, new PropertyChangedEventArgs(name));
        }
    }
}
'@
    Add-Type -TypeDefinition $cardCSharp -Language CSharp -ReferencedAssemblies 'PresentationCore', 'WindowsBase', 'PresentationFramework', 'System.Xaml'
}

function Get-EdgeExecutablePath {
    [CmdletBinding()]
    param()
    $candidates = @(
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        (Join-Path -Path $env:LOCALAPPDATA -ChildPath 'Microsoft\Edge\Application\msedge.exe')
    )
    foreach ($candidate in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -Path $candidate -PathType Leaf)) {
            return $candidate
        }
    }
    return $null
}

function Get-CleanProfileName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$FileName
    )
    $base = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $base = $base -replace '^(?i)edge[\s_-]+', ''
    $base = $base -replace '_', ' '
    $base = ($base -replace '\s+', ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($base)) {
        return 'Profile'
    }
    return $base
}

function Get-SafeFileSlug {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name
    )
    $slug = $Name -replace '[^a-zA-Z0-9_-]', '_'
    $slug = $slug -replace '_+', '_'
    $slug = $slug.Trim('_')
    if ([string]::IsNullOrWhiteSpace($slug)) {
        return 'Profile'
    }
    return $slug
}

function New-ThumbnailBitmap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ImagePath
    )
    $bytes = [System.IO.File]::ReadAllBytes($ImagePath)
    $stream = New-Object System.IO.MemoryStream(, $bytes)
    try {
        $decoder = [System.Windows.Media.Imaging.BitmapDecoder]::Create(
            $stream,
            [System.Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat,
            [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        )
        $bestFrame = $decoder.Frames | Sort-Object -Property PixelWidth -Descending | Select-Object -First 1
        if ($null -eq $bestFrame) {
            throw "No image frames decoded from '$ImagePath'."
        }
        $frozenFrame = $bestFrame.Clone()
        if ($frozenFrame.CanFreeze) {
            $frozenFrame.Freeze()
        }
        return $frozenFrame
    } finally {
        $stream.Dispose()
    }
}

function ConvertTo-IconFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$DestinationIcoPath
    )
    $destDir = Split-Path -Path $DestinationIcoPath -Parent
    if (-not (Test-Path -Path $destDir)) {
        New-Item -Path $destDir -ItemType Directory -Force | Out-Null
    }

    $ext = [System.IO.Path]::GetExtension($SourcePath).ToLowerInvariant()
    if ($ext -eq '.ico') {
        $resolvedSrc = [System.IO.Path]::GetFullPath($SourcePath)
        $resolvedDst = [System.IO.Path]::GetFullPath($DestinationIcoPath)
        if ($resolvedSrc -ne $resolvedDst) {
            Copy-Item -Path $SourcePath -Destination $DestinationIcoPath -Force
        }
        return $false
    }

    $sizes = @(256, 64, 48, 32, 16)
    $srcBytes = [System.IO.File]::ReadAllBytes($SourcePath)
    $srcStream = New-Object System.IO.MemoryStream(, $srcBytes)
    $srcImage = $null
    $framePngList = New-Object System.Collections.Generic.List[byte[]]

    try {
        $srcImage = [System.Drawing.Image]::FromStream($srcStream)
        foreach ($size in $sizes) {
            $bitmap = New-Object System.Drawing.Bitmap($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            $frameMem = New-Object System.IO.MemoryStream
            try {
                $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
                $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
                $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
                $graphics.Clear([System.Drawing.Color]::Transparent)

                $scale = [Math]::Min(($size / [double]$srcImage.Width), ($size / [double]$srcImage.Height))
                $drawW = [Math]::Max(1, [int][Math]::Round($srcImage.Width * $scale))
                $drawH = [Math]::Max(1, [int][Math]::Round($srcImage.Height * $scale))
                $drawX = [int][Math]::Floor(($size - $drawW) / 2.0)
                $drawY = [int][Math]::Floor(($size - $drawH) / 2.0)

                $graphics.DrawImage($srcImage, $drawX, $drawY, $drawW, $drawH)
                $bitmap.Save($frameMem, [System.Drawing.Imaging.ImageFormat]::Png)
                $framePngList.Add($frameMem.ToArray())
            } finally {
                $frameMem.Dispose()
                $graphics.Dispose()
                $bitmap.Dispose()
            }
        }
    } finally {
        if ($null -ne $srcImage) { $srcImage.Dispose() }
        $srcStream.Dispose()
    }

    $fileStream = New-Object System.IO.FileStream($DestinationIcoPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
    $writer = New-Object System.IO.BinaryWriter($fileStream)
    try {
        $count = $sizes.Count
        $writer.Write([UInt16]0)
        $writer.Write([UInt16]1)
        $writer.Write([UInt16]$count)

        $offset = 6 + (16 * $count)
        for ($i = 0; $i -lt $count; $i++) {
            $dim = $sizes[$i]
            $dimByte = if ($dim -ge 256) { [byte]0 } else { [byte]$dim }
            $pngLen = [UInt32]$framePngList[$i].Length

            $writer.Write($dimByte)
            $writer.Write($dimByte)
            $writer.Write([byte]0)
            $writer.Write([byte]0)
            $writer.Write([UInt16]1)
            $writer.Write([UInt16]32)
            $writer.Write($pngLen)
            $writer.Write([UInt32]$offset)

            $offset += $pngLen
        }

        for ($i = 0; $i -lt $count; $i++) {
            $writer.Write($framePngList[$i])
        }
        $writer.Flush()
    } finally {
        $writer.Close()
        $fileStream.Close()
    }
    return $true
}

function New-EdgeProfileShortcut {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][string]$EdgeExe,
        [Parameter(Mandatory = $true)][string]$ProfileName,
        [Parameter(Mandatory = $true)][string]$SafeSlug,
        [Parameter(Mandatory = $true)][string]$SourceIconPath,
        [Parameter(Mandatory = $true)][string]$OutputRoot,
        [Parameter(Mandatory = $true)][bool]$IsolatedMode,
        [Parameter(Mandatory = $true)][int]$ProfileIndex,
        [Parameter(Mandatory = $true)][bool]$UseInPrivate,
        [Parameter(Mandatory = $true)][bool]$CopyToDesktop
    )

    $shortcutsDir = Join-Path -Path $OutputRoot -ChildPath 'Shortcuts'
    $profilesDir = Join-Path -Path $OutputRoot -ChildPath 'Profiles'
    $convertedDir = Join-Path -Path $OutputRoot -ChildPath 'ConvertedIcons'

    foreach ($dirPath in @($shortcutsDir, $convertedDir)) {
        if (-not (Test-Path -Path $dirPath)) {
            New-Item -Path $dirPath -ItemType Directory -Force | Out-Null
        }
    }

    $finalIcoPath = Join-Path -Path $convertedDir -ChildPath "$SafeSlug.ico"
    $convertedFromImage = ConvertTo-IconFile -SourcePath $SourceIconPath -DestinationIcoPath $finalIcoPath

    $inPrivateFlag = if ($UseInPrivate) { ' --inprivate' } else { '' }
    if ($IsolatedMode) {
        $instanceDir = Join-Path -Path $profilesDir -ChildPath "Edge_$SafeSlug"
        if (-not (Test-Path -Path $instanceDir)) {
            New-Item -Path $instanceDir -ItemType Directory -Force | Out-Null
        }
        $edgeArgs = "--user-data-dir=`"$instanceDir`"$inPrivateFlag --no-first-run --no-default-browser-check"
    } else {
        $edgeArgs = "--profile-directory=`"Profile $ProfileIndex`"$inPrivateFlag"
    }

    $invalidFileChars = [System.IO.Path]::GetInvalidFileNameChars() -join ''
    $cleanTitle = ($ProfileName -replace "[$([Regex]::Escape($invalidFileChars))]", '_').Trim()
    if ([string]::IsNullOrWhiteSpace($cleanTitle)) {
        $cleanTitle = $SafeSlug
    }

    $shortcutFileName = if ($UseInPrivate) { "Edge - $cleanTitle (InPrivate).lnk" } else { "Edge - $cleanTitle.lnk" }
    $shortcutPath = Join-Path -Path $shortcutsDir -ChildPath $shortcutFileName
    $modeTag = if ($UseInPrivate) { 'InPrivate' } else { 'Profile' }
    $uniqueAppId = "Pheneox.Edge.$modeTag.$SafeSlug"

    if ($PSCmdlet.ShouldProcess($shortcutPath, 'Create Edge profile shortcut')) {
        $wshShell = New-Object -ComObject WScript.Shell
        try {
            $targets = New-Object System.Collections.Generic.List[string]
            $targets.Add($shortcutPath)
            if ($CopyToDesktop) {
                $desktopDir = [Environment]::GetFolderPath('Desktop')
                if ($desktopDir -and (Test-Path -Path $desktopDir)) {
                    $targets.Add((Join-Path -Path $desktopDir -ChildPath $shortcutFileName))
                }
            }

            foreach ($lnkPath in $targets) {
                $lnk = $wshShell.CreateShortcut($lnkPath)
                $lnk.TargetPath = $EdgeExe
                $lnk.Arguments = $edgeArgs
                $lnk.IconLocation = "$finalIcoPath,0"
                $lnk.WorkingDirectory = Split-Path -Path $EdgeExe -Parent
                $lnk.Description = "Microsoft Edge $modeTag ($cleanTitle)"
                $lnk.Save()
                [PheneoxEdgeShortcutAppId]::SetShortcutAppId($lnkPath, $uniqueAppId)
            }
        } finally {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wshShell) | Out-Null
        }
    }

    return [PSCustomObject]@{
        ShortcutPath       = $shortcutPath
        IconPath           = $finalIcoPath
        AppId              = $uniqueAppId
        ConvertedFromImage = $convertedFromImage
    }
}

Export-ModuleMember -Function @(
    'Get-EdgeExecutablePath',
    'Get-CleanProfileName',
    'Get-SafeFileSlug',
    'New-ThumbnailBitmap',
    'ConvertTo-IconFile',
    'New-EdgeProfileShortcut'
)
