<#
.SYNOPSIS
    Interactive GUI tool to package TCP/IP printers as Microsoft Intune Win32 Apps (.intunewin).
.DESCRIPTION
    Provides a WPF graphical interface to validate printer drivers (.inf), configure TCP/IP printer ports,
    and generate a complete Microsoft Intune Win32 deployment bundle (.intunewin, custom detection script,
    macOS shell script, and step-by-step Intune upload instructions with C:\Logs\Printer_<Name>\ logging).
.NOTES
    Script Name   : AddPrinter.ps1
    Author        : JP
    Version       : 1.4.3
    Created       : 10/09/26
    Last Modified : 17/11/26
    Log Path      : C:\Logs\Printer_<PrinterName>\
#>

# Ensure Single-Threaded Apartment (STA) mode required for PresentationFramework (WPF)
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $relaunchPath = if ($PSScriptRoot) { Join-Path $PSScriptRoot $MyInvocation.MyCommand.Name } else { $PSCommandPath }
    if ($relaunchPath) {
        Start-Process powershell.exe -ArgumentList "-STA -ExecutionPolicy Bypass -NoProfile -File `"$relaunchPath`""
        exit
    }
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

function Initialize-StandardGuiWindow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Window]$Window,
        [switch]$PreserveSize
    )

    $fontFiles = @('Figtree-Regular.ttf', 'Figtree-SemiBold.ttf', 'Figtree-Bold.ttf')
    $resolvedFontDir = $null

    $searchRoot = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $PSScriptRoot
    } elseif (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
        Split-Path -Path $PSCommandPath -Parent
    } else {
        $null
    }

    if ($searchRoot) {
        $probe = $searchRoot
        for ($i = 0; $i -lt 5 -and -not [string]::IsNullOrWhiteSpace($probe); $i++) {
            $candidateDir = Join-Path -Path $probe -ChildPath 'assets\fonts'
            if (Test-Path -Path (Join-Path -Path $candidateDir -ChildPath 'Figtree-Regular.ttf') -PathType Leaf) {
                $resolvedFontDir = $candidateDir
                break
            }
            $probe = Split-Path -Path $probe -Parent
        }
    }

    if (-not $resolvedFontDir) {
        $cacheDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'PowerShellGui\fonts'
        try {
            if (-not (Test-Path -Path $cacheDir)) {
                New-Item -Path $cacheDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
            }
            $allCached = $true
            foreach ($ff in $fontFiles) {
                if (-not (Test-Path -Path (Join-Path -Path $cacheDir -ChildPath $ff) -PathType Leaf)) {
                    $allCached = $false
                    break
                }
            }
            if (-not $allCached) {
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
                $prevProgress = $ProgressPreference
                $ProgressPreference = 'SilentlyContinue'
                try {
                    foreach ($ff in $fontFiles) {
                        $dest = Join-Path -Path $cacheDir -ChildPath $ff
                        if (-not (Test-Path -Path $dest -PathType Leaf)) {
                            $url = "https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/assets/fonts/$ff"
                            Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop
                        }
                    }
                } finally {
                    $ProgressPreference = $prevProgress
                }
            }
            if (Test-Path -Path (Join-Path -Path $cacheDir -ChildPath 'Figtree-Regular.ttf') -PathType Leaf) {
                $resolvedFontDir = $cacheDir
            }
        } catch {
            Write-Verbose "Figtree font download skipped: $($_.Exception.Message)"
        }
    }

    if ($resolvedFontDir) {
        try {
            $fontUri = ([System.Uri]($resolvedFontDir.TrimEnd('\') + '\')).AbsoluteUri
            $Window.FontFamily = New-Object System.Windows.Media.FontFamily ("$($fontUri)#Figtree, Segoe UI Variable, Segoe UI")
        } catch {
            $Window.FontFamily = New-Object System.Windows.Media.FontFamily ('Segoe UI Variable, Segoe UI')
        }
    } else {
        $Window.FontFamily = New-Object System.Windows.Media.FontFamily ('Segoe UI Variable, Segoe UI')
    }

    if (-not $PreserveSize) {
        $workArea = [System.Windows.SystemParameters]::WorkArea
        $targetWidth  = [Math]::Round([Math]::Min($workArea.Width - 24,  [Math]::Max(860, [Math]::Min(1440, $workArea.Width * 0.70))))
        $targetHeight = [Math]::Round([Math]::Min($workArea.Height - 24, [Math]::Max(640, [Math]::Min(900,  $workArea.Height * 0.80))))

        $Window.WindowStartupLocation = 'Manual'
        $Window.Width  = $targetWidth
        $Window.Height = $targetHeight
        $Window.Left   = [Math]::Round($workArea.Left + (($workArea.Width - $targetWidth) / 2))
        $Window.Top    = [Math]::Round($workArea.Top + (($workArea.Height - $targetHeight) / 2))
        $Window.MinWidth  = [Math]::Min(860, [int]$targetWidth)
        $Window.MinHeight = [Math]::Min(640, [int]$targetHeight)
        $Window.ResizeMode = 'CanResize'
    }
}

# WPF XAML Definition
$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Intune Printer Packager v1.4.3" Height="740" Width="1060"
        MinHeight="640" MinWidth="860"
        ResizeMode="CanResize" WindowStartupLocation="Manual"
        Foreground="#F8FAFC">
    <Window.Background>
        <RadialGradientBrush Center="0.5,0.0" GradientOrigin="0.5,0.0" RadiusX="0.95" RadiusY="0.85">
            <GradientStop Color="#0F1F3D" Offset="0.0"/>
            <GradientStop Color="#091122" Offset="0.55"/>
            <GradientStop Color="#05070C" Offset="1.0"/>
        </RadialGradientBrush>
    </Window.Background>
    <Window.Resources>
        <!-- Primary Teal CTA Button Style -->
        <Style x:Key="PrimaryButton" TargetType="Button">
            <Setter Property="Background" Value="#2DD4BF"/>
            <Setter Property="Foreground" Value="#042F2E"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="12.5"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#5EEAD4"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#14B8A6"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Secondary Button Style -->
        <Style x:Key="SecondaryButton" TargetType="Button">
            <Setter Property="Background" Value="#14FFFFFF"/>
            <Setter Property="Foreground" Value="#E2E8F0"/>
            <Setter Property="BorderBrush" Value="#26354D"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="12.5"/>
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#24FFFFFF"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="#2DD4BF"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#0DFFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Build CTA Button Style -->
        <Style x:Key="BuildButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
            <Setter Property="FontSize" Value="13.5"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Padding" Value="22,9"/>
        </Style>

        <!-- Input Field Border Style -->
        <Style x:Key="InputBorderStyle" TargetType="Border">
            <Setter Property="Background" Value="#99070D1A"/>
            <Setter Property="BorderBrush" Value="#26354D"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="6"/>
            <Setter Property="Padding" Value="6,3"/>
            <Style.Triggers>
                <Trigger Property="IsKeyboardFocusWithin" Value="True">
                    <Setter Property="BorderBrush" Value="#2DD4BF"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header Bar -->
        <Border Grid.Row="0" Background="#B3070D19" BorderBrush="#1FFFFFFF" BorderThickness="0,0,0,1" Padding="24,14">
            <DockPanel>
                <Border DockPanel.Dock="Right" Background="#660D192E" BorderBrush="#26354D" BorderThickness="1"
                        CornerRadius="12" Padding="12,4" VerticalAlignment="Center">
                    <TextBlock Text="v1.4.3 | Win32 App" FontSize="11.5" FontWeight="SemiBold" Foreground="#2DD4BF"/>
                </Border>
                <StackPanel>
                    <TextBlock Text="Intune Printer Packager" FontSize="20" FontWeight="Bold" Foreground="#F8FAFC"/>
                    <TextBlock Text="Validate printer INF drivers, configure TCP/IP ports, and export production-ready .intunewin bundles"
                               FontSize="12.5" Foreground="#94A3B8" Margin="0,3,0,0"/>
                </StackPanel>
            </DockPanel>
        </Border>

        <!-- Flush 2-Column Workspace -->
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="1.15*"/>
                <ColumnDefinition Width="1"/>
                <ColumnDefinition Width="0.85*"/>
            </Grid.ColumnDefinitions>

            <!-- Left Column: Inputs -->
            <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto" Padding="24,18">
                <StackPanel>
                    <TextBlock Text="PRINTER &amp; DRIVER CONFIGURATION" FontSize="11" FontWeight="Bold" Foreground="#2DD4BF" Margin="0,0,0,14"/>

                    <!-- Printer Name Input -->
                    <TextBlock Text="Printer Name (e.g., Showroom | Upstairs | Sharp MX-4061)" Foreground="#CBD5E1" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <Border Style="{StaticResource InputBorderStyle}" Margin="0,0,0,14">
                        <TextBox Name="txtPrinterName" BorderThickness="0" Background="Transparent" Foreground="#F8FAFC" FontSize="13" Padding="4" VerticalContentAlignment="Center" CaretBrush="#2DD4BF"/>
                    </Border>

                    <!-- Driver Name Input -->
                    <TextBlock Text="Driver Name (e.g., SHARP MX-4061 PCL6)" Foreground="#CBD5E1" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <Border Style="{StaticResource InputBorderStyle}" Margin="0,0,0,14">
                        <TextBox Name="txtDriverName" BorderThickness="0" Background="Transparent" Foreground="#F8FAFC" FontSize="13" Padding="4" VerticalContentAlignment="Center" CaretBrush="#2DD4BF"/>
                    </Border>

                    <!-- IP Address Input -->
                    <TextBlock Text="Printer IPv4 Address (Port 9100, e.g., 192.168.0.20)" Foreground="#CBD5E1" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <Border Style="{StaticResource InputBorderStyle}" Margin="0,0,0,14">
                        <TextBox Name="txtPrinterIP" BorderThickness="0" Background="Transparent" Foreground="#F8FAFC" FontSize="13" Padding="4" VerticalContentAlignment="Center" CaretBrush="#2DD4BF"/>
                    </Border>

                    <!-- Driver Source Folder Browse -->
                    <TextBlock Text="Printer Driver Folder (Source containing .inf &amp; .cat files)" Foreground="#CBD5E1" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <Grid Margin="0,0,0,14">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <Border Grid.Column="0" Style="{StaticResource InputBorderStyle}" Margin="0,0,10,0">
                            <TextBox Name="txtDriverPath" IsReadOnly="True" BorderThickness="0" Background="Transparent" Foreground="#94A3B8" FontSize="12" Padding="4" VerticalContentAlignment="Center"/>
                        </Border>
                        <Button Name="btnBrowseDriver" Grid.Column="1" Content="Browse..." Style="{StaticResource SecondaryButton}"/>
                    </Grid>

                    <!-- Output Folder Browse -->
                    <TextBlock Text="Output Folder (Where .intunewin bundle &amp; guide will be saved)" Foreground="#CBD5E1" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <Grid Margin="0,0,0,10">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <Border Grid.Column="0" Style="{StaticResource InputBorderStyle}" Margin="0,0,10,0">
                            <TextBox Name="txtOutputPath" IsReadOnly="True" BorderThickness="0" Background="Transparent" Foreground="#94A3B8" FontSize="12" Padding="4" VerticalContentAlignment="Center"/>
                        </Border>
                        <Button Name="btnBrowseOutput" Grid.Column="1" Content="Browse..." Style="{StaticResource SecondaryButton}"/>
                    </Grid>
                </StackPanel>
            </ScrollViewer>

            <!-- 1px Crisp Hairline Divider -->
            <Border Grid.Column="1" Background="#1FFFFFFF"/>

            <!-- Right Column: Bundle Artifacts + Console -->
            <Grid Grid.Column="2" Margin="20,18,24,18">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <Border Grid.Row="0" Background="#660D192E" BorderBrush="#26354D" BorderThickness="1" CornerRadius="6" Padding="14,12" Margin="0,0,0,14">
                    <StackPanel>
                        <TextBlock Text="GENERATED BUNDLE ARTIFACTS" FontSize="11" FontWeight="Bold" Foreground="#2DD4BF" Margin="0,0,0,8"/>
                        <TextBlock Text="&#x2022; &lt;Printer&gt;.intunewin (SYSTEM installer + driver store staging)" Foreground="#CBD5E1" FontSize="12" Margin="0,0,0,3"/>
                        <TextBlock Text="&#x2022; &lt;Printer&gt;_Detection.ps1 (registry + Get-Printer verification)" Foreground="#CBD5E1" FontSize="12" Margin="0,0,0,3"/>
                        <TextBlock Text="&#x2022; &lt;Printer&gt;_macOS_Install.sh (lpadmin LPD queue script)" Foreground="#CBD5E1" FontSize="12" Margin="0,0,0,3"/>
                        <TextBlock Text="&#x2022; &lt;Printer&gt;_Instructions.txt (5-step Intune upload wizard guide)" Foreground="#CBD5E1" FontSize="12"/>
                    </StackPanel>
                </Border>

                <Grid Grid.Row="1">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <TextBlock Grid.Row="0" Text="BUILD LOG CONSOLE (SELECTABLE)" Foreground="#2DD4BF" FontSize="11" FontWeight="Bold" Margin="0,0,0,6"/>
                    <Border Grid.Row="1" BorderBrush="#26354D" BorderThickness="1" CornerRadius="6" Background="#CC050A14" Padding="10">
                        <ScrollViewer Name="scrollConsole" VerticalScrollBarVisibility="Auto">
                            <TextBox Name="txtStatus" Text="Ready to configure..." IsReadOnly="True" BorderThickness="0" Background="Transparent" Foreground="#34D399" FontFamily="Cascadia Mono, Consolas" FontSize="11.5" TextWrapping="Wrap" CaretBrush="#2DD4BF"/>
                        </ScrollViewer>
                    </Border>
                </Grid>
            </Grid>
        </Grid>

        <!-- Single-Row Footer -->
        <Border Grid.Row="2" Background="#B3070D19" BorderBrush="#1FFFFFFF" BorderThickness="0,1,0,0" Padding="24,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0"
                           Text="Endpoint logs: C:\Logs\Printer_&lt;Name&gt;\ (Install-Printer.log, Detect-Printer.log, Uninstall-Printer.log)"
                           Foreground="#94A3B8" FontSize="11.5" VerticalAlignment="Center"
                           TextTrimming="CharacterEllipsis" Margin="0,0,16,0"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button Name="btnOpenOutput" Content="Open Output Folder" Style="{StaticResource SecondaryButton}" Margin="0,0,10,0"/>
                    <Button Name="btnBuild" Content="Build IntuneWin Package" Style="{StaticResource BuildButton}"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

# Parse XAML
$reader = (New-Object System.Xml.XmlNodeReader ([xml]$xaml))
$Form = [Windows.Markup.XamlReader]::Load($reader)
Initialize-StandardGuiWindow -Window $Form

# Get UI element handles
$txtPrinterName  = $Form.FindName("txtPrinterName")
$txtDriverName   = $Form.FindName("txtDriverName")
$txtPrinterIP    = $Form.FindName("txtPrinterIP")
$txtDriverPath   = $Form.FindName("txtDriverPath")
$txtOutputPath   = $Form.FindName("txtOutputPath")
$btnBrowseDriver = $Form.FindName("btnBrowseDriver")
$btnBrowseOutput = $Form.FindName("btnBrowseOutput")
$btnOpenOutput   = $Form.FindName("btnOpenOutput")
$btnBuild        = $Form.FindName("btnBuild")
$txtStatus       = $Form.FindName("txtStatus")
$scrollConsole   = $Form.FindName("scrollConsole")

# Autofill default output folder to the script directory if it runs locally, or Downloads if running in-memory
$defaultOutput = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot) -and (Test-Path -Path $PSScriptRoot)) {
    $PSScriptRoot
} else {
    Join-Path ([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile)) "Downloads"
}
$txtOutputPath.Text = $defaultOutput

# Helper to update console log UI responsively
function Update-Log ($message, $color = "#34D399") {
    if ([string]::IsNullOrWhiteSpace($txtStatus.Text) -or $txtStatus.Text -eq "Ready to configure...") {
        $txtStatus.Text = $message
    } else {
        $txtStatus.Text = $txtStatus.Text + "`r`n" + $message
    }
    $converter = New-Object System.Windows.Media.BrushConverter
    $txtStatus.Foreground = $converter.ConvertFromString($color)
    $scrollConsole.ScrollToEnd()
    $null = [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [Action]{})
}

# Function to locate or automatically download IntuneWinAppUtil.exe
function Get-IntuneWinAppUtilPath {
    $candidatePaths = @()

    # 1. Check script directory or tools subfolder (local clone / batch launcher)
    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $candidatePaths += (Join-Path $PSScriptRoot "IntuneWinAppUtil.exe")
        $candidatePaths += (Join-Path $PSScriptRoot "tools\IntuneWinAppUtil.exe")
    }

    # 2. Check user's local application data and persistent tool cache
    $localAppDataDir = Join-Path $env:LOCALAPPDATA "Microsoft\IntuneWinAppUtil"
    $customAppDataDir = Join-Path $env:LOCALAPPDATA "IntunePrinterPackager"
    $candidatePaths += (Join-Path $localAppDataDir "IntuneWinAppUtil.exe")
    $candidatePaths += (Join-Path $customAppDataDir "IntuneWinAppUtil.exe")
    $candidatePaths += (Join-Path $env:TEMP "IntuneWinAppUtil.exe")

    # 3. Check user downloads
    $userDownloads = Join-Path ([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile)) "Downloads"
    $candidatePaths += (Join-Path $userDownloads "IntuneWinAppUtil.exe")

    # 4. Check known sibling repository directories if in dev environment
    $scriptDir = $PSScriptRoot
    if (-not [string]::IsNullOrWhiteSpace($scriptDir)) {
        $intuneRoot = Split-Path -Parent $scriptDir
        if (-not [string]::IsNullOrWhiteSpace($intuneRoot)) {
            $candidatePaths += (Join-Path $intuneRoot "DesktopOverlay\IntuneWinAppUtil.exe")
            $candidatePaths += (Join-Path $intuneRoot "TeamsBackground\tools\IntuneWinAppUtil.exe")
        }
    }

    # 5. Check system PATH
    $cmd = try { Get-Command "IntuneWinAppUtil.exe" -ErrorAction Stop } catch { $null }
    if ($cmd -and [string]::IsNullOrWhiteSpace($cmd.Source) -eq $false -and (Test-Path -Path $cmd.Source -PathType Leaf)) {
        return $cmd.Source
    }

    # 6. Check existing candidate paths
    foreach ($cand in $candidatePaths) {
        if (-not [string]::IsNullOrWhiteSpace($cand) -and (Test-Path -Path $cand -PathType Leaf)) {
            try {
                if ((Get-Item $cand).Length -gt 10000) {
                    Update-Log "Found Win32 Content Prep Tool at: $cand"
                    return $cand
                }
            } catch {}
        }
    }

    # 7. Tool not found locally; determine persistent target download directory
    $downloadTargetDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot) -and (Test-Path -Path $PSScriptRoot)) {
        $PSScriptRoot
    } else {
        $customAppDataDir
    }

    if (-not (Test-Path -Path $downloadTargetDir)) {
        try { New-Item -ItemType Directory -Path $downloadTargetDir -Force -ErrorAction Stop | Out-Null } catch {}
    }
    $targetExePath = Join-Path $downloadTargetDir "IntuneWinAppUtil.exe"

    # Prompt user to confirm auto-download
    $msg = "Microsoft Win32 Content Prep Tool (IntuneWinAppUtil.exe) was not found.`n`nWould you like to automatically download it from Microsoft's official GitHub repository?"
    $confirm = [System.Windows.MessageBox]::Show($msg, "Packer Tool Missing", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)

    if ($confirm -eq [System.Windows.MessageBoxResult]::Yes) {
        Update-Log "IntuneWinAppUtil.exe is missing. Attempting official download..."
        try {
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
            $downloadUrl = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe"
            Update-Log "Downloading from official Microsoft GitHub to: $targetExePath"

            [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
            $null = [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [Action]{})

            Invoke-WebRequest -Uri $downloadUrl -OutFile $targetExePath -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
            $sig = Get-AuthenticodeSignature -FilePath $targetExePath -ErrorAction Stop
            if ($sig.Status -ne 'Valid') {
                if (Test-Path -LiteralPath $targetExePath) { try { Remove-Item $targetExePath -Force -ErrorAction Stop } catch {} }
                throw "Authenticode signature check failed for IntuneWinAppUtil.exe ($($sig.Status))"
            }

            [System.Windows.Input.Mouse]::OverrideCursor = $null

            if ((Test-Path -Path $targetExePath -PathType Leaf) -and ((Get-Item $targetExePath).Length -gt 10000)) {
                Update-Log "Download complete! IntuneWinAppUtil.exe is now ready." "#34D399"
                return $targetExePath
            } else {
                throw "Downloaded file is invalid or zero bytes."
            }
        } catch {
            [System.Windows.Input.Mouse]::OverrideCursor = $null
            Update-Log "Failed to download packer tool: $_" "#F87171"
        }
    }

    # If auto-download declined or failed, offer manual file browser
    $browseMsg = "Would you like to browse and locate IntuneWinAppUtil.exe manually on your computer?"
    $browseConfirm = [System.Windows.MessageBox]::Show($browseMsg, "Locate IntuneWinAppUtil.exe", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Information)
    if ($browseConfirm -eq [System.Windows.MessageBoxResult]::Yes) {
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title = "Locate IntuneWinAppUtil.exe"
        $dlg.Filter = "Executable Files (*.exe)|*.exe"
        $dlg.FileName = "IntuneWinAppUtil.exe"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            if (-not [string]::IsNullOrWhiteSpace($dlg.FileName) -and (Test-Path -Path $dlg.FileName -PathType Leaf)) {
                Update-Log "Using IntuneWinAppUtil.exe: $($dlg.FileName)"
                return $dlg.FileName
            }
        }
    }

    return $null
}

# Open output folder
$btnOpenOutput.Add_Click({
    $dir = $txtOutputPath.Text.Trim()
    if (-not [string]::IsNullOrWhiteSpace($dir)) {
        if (-not (Test-Path -Path $dir)) {
            try { New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null } catch {}
        }
        if (Test-Path -Path $dir) {
            Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$dir`""
        }
    }
})

# Browser event for driver folder
$btnBrowseDriver.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select the folder containing the printer drivers (must contain a .inf file)"
    $dialog.ShowNewFolderButton = $false
    
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtDriverPath.Text = $dialog.SelectedPath
        Update-Log "Driver source folder selected: $($dialog.SelectedPath)"
    }
})

# Browser event for output folder
$btnBrowseOutput.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select output folder to save the .intunewin package"
    $dialog.ShowNewFolderButton = $true
    
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOutputPath.Text = $dialog.SelectedPath
        Update-Log "Output folder selected: $($dialog.SelectedPath)"
    }
})

# Package Build Handler
$btnBuild.Add_Click({
    $txtStatus.Text = ""
    # 1. Validation
    $printerName = $txtPrinterName.Text.Trim()
    $driverName  = $txtDriverName.Text.Trim()
    $printerIP   = $txtPrinterIP.Text.Trim()
    $driverPath  = $txtDriverPath.Text.Trim()
    $outputPath  = $txtOutputPath.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($printerName) -or 
        [string]::IsNullOrWhiteSpace($driverName) -or 
        [string]::IsNullOrWhiteSpace($printerIP)) {
        [System.Windows.MessageBox]::Show("Please fill out all printer parameters (Name, Driver Name, and IP Address).", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    # Verify IP Address format (ensures it is a mathematically valid IPv4 address)
    $parsedIp = $null
    if (-not [System.Net.IPAddress]::TryParse($printerIP, [ref]$parsedIp) -or $parsedIp.AddressFamily -ne 'InterNetwork') {
        [System.Windows.MessageBox]::Show("Please enter a valid IPv4 address (e.g., 192.168.0.20).", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    if ([string]::IsNullOrWhiteSpace($driverPath) -or (-not (Test-Path -Path $driverPath))) {
        [System.Windows.MessageBox]::Show("Please select a valid driver directory.", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    # Verify that there are INF files inside the driver directory
    $infFiles = try { Get-ChildItem -Path $driverPath -Filter "*.inf" -Recurse -ErrorAction Stop } catch { @() }
    if (-not $infFiles) {
        [System.Windows.MessageBox]::Show("The selected driver folder does not contain any driver definition (.inf) files. Please verify the directory.", "No Drivers Found", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    if ([string]::IsNullOrWhiteSpace($outputPath) -or (-not (Test-Path -Path $outputPath))) {
        [System.Windows.MessageBox]::Show("Please select a valid output directory.", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    # Check for IntuneWinAppUtil.exe
    $rawPacker = Get-IntuneWinAppUtilPath
    $packerExe = if ($rawPacker -is [System.Array]) {
        ($rawPacker | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Last 1)
    } else {
        $rawPacker
    }
    $packerExe = [string]$packerExe
    if ([string]::IsNullOrWhiteSpace($packerExe) -or (-not (Test-Path -Path $packerExe -PathType Leaf))) {
        Update-Log "Build aborted. Microsoft IntuneWinAppUtil.exe is required to compile .intunewin packages." "#F87171"
        return
    }

    # Verify that the input Driver Name matches one of the models defined in the INF files
    Update-Log "Verifying Driver Name against INF configuration..."
    $foundDriverNameInInf = $false
    # First pass: Quick scan for the exact driver name (case-insensitive match)
    foreach ($inf in $infFiles) {
        try {
            $content = Get-Content -Path $inf.FullName -Raw -ErrorAction Stop
            if ($content -and $content -match [regex]::Escape($driverName)) {
                $foundDriverNameInInf = $true
                break # Matched! Skip scanning the rest of the INF files.
            }
        } catch {}
    }

    # Second pass: Only extract candidates if the name was not found (mismatch fallback)
    $candidates = @()
    if (-not $foundDriverNameInInf) {
        foreach ($inf in $infFiles) {
            try {
                $lines = Get-Content -Path $inf.FullName -ErrorAction Stop
                $inStringsSection = $false
                foreach ($line in $lines) {
                    $trimmed = $line.Trim()
                    if ($trimmed -match '^\[strings\]') {
                        $inStringsSection = $true
                        continue
                    }
                    if ($trimmed -match '^\[.*\]' -and $trimmed -notmatch '^\[strings\]') {
                        $inStringsSection = $false
                    }
                    if ($inStringsSection) {
                        if ($trimmed -match '^[^=]+=\s*"([^"]+)"') {
                            $candidate = $Matches[1].Trim()
                            if ($candidate.Length -gt 5 -and $candidate -notin $candidates) {
                                $candidates += $candidate
                            }
                        }
                    }
                }
            } catch {}
        }
    }

    if (-not $foundDriverNameInInf) {
        $dialogXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Driver Name Mismatch" Height="460" Width="540"
        MinHeight="380" MinWidth="460"
        ResizeMode="CanResize" WindowStartupLocation="CenterOwner"
        Foreground="#F8FAFC">
    <Window.Background>
        <RadialGradientBrush Center="0.5,0.0" GradientOrigin="0.5,0.0" RadiusX="0.95" RadiusY="0.85">
            <GradientStop Color="#0F1F3D" Offset="0.0"/>
            <GradientStop Color="#091122" Offset="0.55"/>
            <GradientStop Color="#05070C" Offset="1.0"/>
        </RadialGradientBrush>
    </Window.Background>
    <Window.Resources>
        <Style TargetType="ListBoxItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#F8FAFC"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="Margin" Value="0,1"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBoxItem">
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#14FFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#242DD4BF"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Foreground" Value="#2DD4BF"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="DialogPrimaryButton" TargetType="Button">
            <Setter Property="Background" Value="#2DD4BF"/>
            <Setter Property="Foreground" Value="#042F2E"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="6" Padding="12,6">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#5EEAD4"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#14B8A6"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="DialogSecondaryButton" TargetType="Button">
            <Setter Property="Background" Value="#14FFFFFF"/>
            <Setter Property="Foreground" Value="#E2E8F0"/>
            <Setter Property="BorderBrush" Value="#26354D"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6" Padding="12,6">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#24FFFFFF"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="#2DD4BF"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="DialogDangerButton" TargetType="Button">
            <Setter Property="Background" Value="#26EF4444"/>
            <Setter Property="Foreground" Value="#FCA5A5"/>
            <Setter Property="BorderBrush" Value="#66EF4444"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6" Padding="12,6">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#45EF4444"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Foreground" Value="#FFFFFF"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="#B3070D19" BorderBrush="#1FFFFFFF" BorderThickness="0,0,0,1" Padding="20,14">
            <StackPanel>
                <TextBlock Text="Driver Name Mismatch" FontSize="16" FontWeight="Bold" Foreground="#FBBF24"/>
                <TextBlock Name="lblDescription" Text="The driver name entered was not found. Select a detected driver below to use, or choose to proceed anyway." FontSize="12" Foreground="#94A3B8" TextWrapping="Wrap" Margin="0,4,0,0"/>
            </StackPanel>
        </Border>

        <!-- Candidates ListBox -->
        <Border Grid.Row="1" BorderBrush="#26354D" BorderThickness="1" CornerRadius="6" Background="#CC050A14" Margin="20,14,20,14" Padding="4">
            <ListBox Name="lstCandidates" Background="Transparent" BorderThickness="0" Foreground="#F8FAFC" FontSize="12.5" ScrollViewer.VerticalScrollBarVisibility="Auto"/>
        </Border>

        <!-- Action Buttons -->
        <Border Grid.Row="2" Background="#B3070D19" BorderBrush="#1FFFFFFF" BorderThickness="0,1,0,0" Padding="20,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Button Name="btnCancel" Grid.Column="0" Content="Cancel Build" Style="{StaticResource DialogDangerButton}" HorizontalAlignment="Left" Height="34"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <Button Name="btnProceed" Content="Proceed Anyway" Style="{StaticResource DialogSecondaryButton}" Height="34" Margin="0,0,8,0"/>
                    <Button Name="btnUseSelected" Content="Use Selected" Style="{StaticResource DialogPrimaryButton}" Height="34"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

        $dialogReader = (New-Object System.Xml.XmlNodeReader ([xml]$dialogXaml))
        $DialogForm = [Windows.Markup.XamlReader]::Load($dialogReader)
        Initialize-StandardGuiWindow -Window $DialogForm -PreserveSize
        
        # Set Owner to parent window to center correctly
        $DialogForm.Owner = $Form

        # Get controls
        $lstCandidates  = $DialogForm.FindName("lstCandidates")
        $btnUseSelected = $DialogForm.FindName("btnUseSelected")
        $btnProceed     = $DialogForm.FindName("btnProceed")
        $btnCancel      = $DialogForm.FindName("btnCancel")
        $lblDescription = $DialogForm.FindName("lblDescription")

        $lblDescription.Text = "The driver name '$driverName' was not found in the INF files. Select a detected driver below to use, or choose to proceed anyway."

        # Load candidates sorted alphabetically
        $sortedCandidates = $candidates | Sort-Object
        if ($sortedCandidates) {
            foreach ($c in $sortedCandidates) {
                $lstCandidates.Items.Add($c) | Out-Null
            }
            $lstCandidates.SelectedIndex = 0
        } else {
            $lstCandidates.Items.Add("(No drivers detected in INF files)") | Out-Null
            $btnUseSelected.IsEnabled = $false
        }

        # Handle button click results using reference hashtable
        $dialogResult = @{
            Action = "cancel"
            Driver = $null
        }

        $btnUseSelected.Add_Click({
            if ($lstCandidates.SelectedItem) {
                $dialogResult.Driver = $lstCandidates.SelectedItem.ToString()
                $dialogResult.Action = "use"
                $DialogForm.Close()
            }
        })

        $btnProceed.Add_Click({
            $dialogResult.Action = "proceed"
            $DialogForm.Close()
        })

        $btnCancel.Add_Click({
            $dialogResult.Action = "cancel"
            $DialogForm.Close()
        })

        # Display window
        $DialogForm.ShowDialog() | Out-Null

        # Evaluate action
        if ($dialogResult.Action -eq "use") {
            # Update the form text box so they see the corrected driver name
            $txtDriverName.Text = $dialogResult.Driver
            $driverName = $dialogResult.Driver
            Update-Log "Updated driver name to matched candidate: $driverName"
        } elseif ($dialogResult.Action -eq "proceed") {
            Update-Log "Proceeding with original driver name: $driverName"
        } else {
            Update-Log "Build cancelled by user due to driver name mismatch." "#F87171"
            return
        }
    }

    # 2. Block UI and set cursor
    $btnBuild.IsEnabled = $false
    $btnBrowseDriver.IsEnabled = $false
    $btnBrowseOutput.IsEnabled = $false
    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
    # Temporary Staging Folder setup
    $tempFolder = Join-Path $env:TEMP "IntunePrinterStaging_$([Guid]::NewGuid().Guid)"
    
    try {
        # Sanitize printer name for folder names, replacing invalid characters with dashes
        $sanitized = $printerName -replace '[\\/:*?"<>|]', '-'
        $sanitized = $sanitized -replace '\s+', ' '
        $sanitized = $sanitized.Trim()

        # Remove all spaces and special characters from the .intunewin filename for Windows Sandbox compatibility
        $cleanFileName = $sanitized -replace '[^a-zA-Z0-9]', ''
        if ([string]::IsNullOrWhiteSpace($cleanFileName)) {
            $cleanFileName = "PrinterPackage"
        }
        $finalFileName = "$cleanFileName.intunewin"

        # Sanitize printer name for macOS printer queue naming (alphanumeric and dashes/underscores only)
        $macQueueName = $sanitized -replace '[^a-zA-Z0-9-]', '_'
        $macQueueName = $macQueueName -replace '_+', '_'
        $macQueueName = $macQueueName.Trim('_')

        # Create output directory dedicated to this printer
        $printerOutputDir = Join-Path $outputPath $sanitized
        Update-Log "Creating output folder: $printerOutputDir"
        New-Item -ItemType Directory -Path $printerOutputDir -Force | Out-Null

        Update-Log "Initializing temporary workspace folder..."
        New-Item -ItemType Directory -Path $tempFolder -Force | Out-Null

        # Staging embedded script templates
        Update-Log "Generating installer scripts..."
        
        $installTemplateText = @'
<#
.SYNOPSIS
    Deploys TCP/IP printers and their drivers globally for all users under the SYSTEM context.
.DESCRIPTION
    1. Pre-trusts Authenticode catalog (.cat) signing certificates in LocalMachine\TrustedPublisher so silent SYSTEM installs via pnputil do not fail with untrusted publisher errors.
    2. Scans for local driver .inf files (filtering out 32-bit and stub 'etc' INFs on 64-bit systems) and registers them in the Windows Driver Store.
    3. Registers the driver name in the Print Spooler.
    4. Creates the TCP/IP printer port machine-wide.
    5. Adds the printer machine-wide so it is visible to all users and configures default settings.
.NOTES
    Script Name   : Install-Printer.ps1
    Author        : JP
    Version       : 1.4.3
    Context       : SYSTEM (Intune Win32 App Install Behavior: System)
    Log Path      : C:\Logs\Printer_<PrinterName>\Install-Printer.log
.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File .\Install-Printer.ps1
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param()

# Read printers.csv as input
$Printers = Import-Csv '.\printers.csv'
$FirstPrinterName = $Printers[0].Name

# Sanitize the printer name for file system compatibility (replacing invalid characters like '|' with '-')
$SanitizedName = ($FirstPrinterName -replace '[\\/:*?"<>|]', '-' -replace '\s+', ' ').Trim()

# Create log folder if it doesn't exist
$LogFolder = "$env:SystemDrive\Logs\Printer_$SanitizedName"
if (-not (Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}
$LogFile = "$LogFolder\Install-Printer.log"

function Write-PrinterLog {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")][string]$Level = "INFO"
    )
    $TimeStamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
    $Entry = "[$TimeStamp] [$Level] $Message"
    Write-Output $Entry
    try { Add-Content -Path $LogFile -Value $Entry -Encoding UTF8 -ErrorAction Stop } catch {}
}

try { Start-Transcript -Path $LogFile -Append -ErrorAction Stop } catch {}
Write-PrinterLog -Message "=== Starting Printer Installation ($FirstPrinterName) ===" -Level "INFO"

$InstallFailed = $false

# Temporarily stop Program Compatibility Assistant Service (PcaSvc) to suppress driver warnings
$PcaService = try { Get-Service -Name "PcaSvc" -ErrorAction Stop } catch { $null }
$PcaServiceStarted = $false
if ($PcaService -and $PcaService.Status -eq 'Running') {
    if ($PSCmdlet.ShouldProcess("PcaSvc", "Temporarily stop Program Compatibility Assistant Service")) {
        Write-PrinterLog -Message "Temporarily stopping Program Compatibility Assistant Service (PcaSvc)..." -Level "INFO"
        try { Stop-Service -Name "PcaSvc" -Force -ErrorAction Stop | Out-Null } catch {}
        $PcaServiceStarted = $true
    }
}

try {
    # 1. Install & register the driver files into the Driver Store using pnputil
    Write-PrinterLog -Message "[1/4] Staging drivers to Windows Driver Store..." -Level "INFO"

    # Pre-trust Authenticode catalog (.cat) signer certificates in LocalMachine\TrustedPublisher
    # Prevents PnPUtil error: "The publisher of an Authenticode(tm) signed catalog has not yet been established as trusted."
    $catFiles = try { Get-ChildItem -Path . -Filter "*.cat" -Recurse -ErrorAction Stop } catch { @() }
    if ($catFiles) {
        try {
            $publisherStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("TrustedPublisher", "LocalMachine")
            $publisherStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
            $existingThumbprints = @($publisherStore.Certificates | ForEach-Object { $_.Thumbprint })

            foreach ($cat in $catFiles) {
                try {
                    $sig = Get-AuthenticodeSignature -FilePath $cat.FullName -ErrorAction Stop
                    if ($sig -and $null -ne $sig.SignerCertificate) {
                        $cert = $sig.SignerCertificate
                        if ($cert.Thumbprint -notin $existingThumbprints) {
                            Write-PrinterLog -Message "Trusting driver catalog publisher certificate: $($cert.Subject) (Thumbprint: $($cert.Thumbprint))" -Level "INFO"
                            $publisherStore.Add($cert)
                            $existingThumbprints += $cert.Thumbprint
                        }
                    }
                } catch {
                    Write-PrinterLog -Message "Could not inspect or trust catalog signature for $($cat.Name): $_" -Level "WARN"
                }
            }
            $publisherStore.Close()
        } catch {
            Write-PrinterLog -Message "Could not open LocalMachine\TrustedPublisher certificate store: $_" -Level "WARN"
        }
    }

    $allInfs = try { Get-ChildItem -Path . -Filter "*.inf" -Recurse -ErrorAction Stop } catch { @() }
    $targetDrivers = $Printers.DriverName | Select-Object -Unique
    $matchedInfs = @()

    foreach ($inf in $allInfs) {
        try {
            $content = Get-Content -Path $inf.FullName -Raw -ErrorAction Stop
            if ($content) {
                foreach ($driver in $targetDrivers) {
                    if ($content -match [regex]::Escape($driver)) {
                        $matchedInfs += $inf.FullName
                        break
                    }
                }
            }
        } catch {}
    }

    if ($matchedInfs.Count -eq 0) {
        Write-PrinterLog -Message "Could not match Driver Name to a specific INF file. Staging all INF files..." -Level "WARN"
        $infsToStage = @($allInfs | Select-Object -ExpandProperty FullName)
    } else {
        $infsToStage = @($matchedInfs)
    }

    # Filter out 32-bit driver folders when running on a 64-bit operating system
    if ([System.Environment]::Is64BitOperatingSystem -and $infsToStage.Count -gt 1) {
        $archFiltered = @($infsToStage | Where-Object { $_ -notmatch '[\\/](32bit|i386|x86|win32)[\\/]' })
        if ($archFiltered.Count -gt 0 -and $archFiltered.Count -lt $infsToStage.Count) {
            Write-PrinterLog -Message "Filtered out $($infsToStage.Count - $archFiltered.Count) 32-bit INF file(s) on 64-bit OS." -Level "INFO"
            $infsToStage = $archFiltered
        }
    }

    # Filter out supplementary 'etc' stub INFs (e.g., Canon \etc\CNP60MA64.INF) when main driver INFs exist
    if ($infsToStage.Count -gt 1) {
        $nonEtcInfs = @($infsToStage | Where-Object { $_ -notmatch '[\\/]etc[\\/]' })
        if ($nonEtcInfs.Count -gt 0 -and $nonEtcInfs.Count -lt $infsToStage.Count) {
            Write-PrinterLog -Message "Skipping $($infsToStage.Count - $nonEtcInfs.Count) supplementary 'etc' INF stub file(s)." -Level "INFO"
            $infsToStage = $nonEtcInfs
        }
    }

    Write-PrinterLog -Message "INF file(s) selected for staging:" -Level "INFO"
    $infsToStage | ForEach-Object { Write-PrinterLog -Message " - $_" -Level "INFO" }

    $totalInfs = $infsToStage.Count
    $currentInf = 1

    # Handle 32-bit execution context on 64-bit Windows (Intune Management Extension runs as 32-bit)
    $PnpUtilPath = "pnputil.exe"
    if (Test-Path "$env:windir\sysnative\pnputil.exe") {
        $PnpUtilPath = "$env:windir\sysnative\pnputil.exe"
    }

    foreach ($inf in $infsToStage) {
        if ($PSCmdlet.ShouldProcess($inf, "Stage INF driver in Windows Driver Store")) {
            Write-PrinterLog -Message "Staging INF ($currentInf/$totalInfs): $inf" -Level "INFO"
            try {
                $pnpResult = & $PnpUtilPath /add-driver $inf /install 2>&1
                $pnpExitCode = $LASTEXITCODE
                $pnpOutputText = ($pnpResult -join ' ')
                if ($pnpExitCode -eq 0 -or $pnpExitCode -eq 259 -or $pnpExitCode -eq 3010) {
                    Write-PrinterLog -Message "PnPUtil Result (Exit $pnpExitCode): $pnpOutputText" -Level "SUCCESS"
                } else {
                    Write-PrinterLog -Message "PnPUtil Result (Exit $pnpExitCode): $pnpOutputText" -Level "WARN"
                }
            } catch {
                Write-PrinterLog -Message "Failed to run PnPUtil for $($inf): $_" -Level "ERROR"
            }
        }
        $currentInf++
    }

    # 2. Add the drivers to the Print Spooler
    Write-PrinterLog -Message "[2/4] Registering drivers in Print Spooler..." -Level "INFO"
    foreach ($driver in $Printers.DriverName | Select-Object -Unique) {
        $existingDriver = try { Get-PrinterDriver -Name $driver -ErrorAction Stop } catch { $null }
        if (-not $existingDriver) {
            if ($PSCmdlet.ShouldProcess($driver, "Add Printer Driver to Print Spooler")) {
                Write-PrinterLog -Message "Adding printer driver to spooler: $driver" -Level "INFO"
                try {
                    Add-PrinterDriver -Name $driver -ErrorAction Stop
                    Write-PrinterLog -Message "Successfully registered driver in spooler: $driver" -Level "SUCCESS"
                } catch {
                    $InstallFailed = $true
                    Write-PrinterLog -Message "Failed to add printer driver $($driver) to spooler: $_" -Level "ERROR"
                    
                    # Diagnostic Logging
                    Write-Host "`n=================== DIAGNOSTICS FOR FAILURE ==================="
                    Write-Host "1. Currently registered spooler drivers on this machine:"
                    try {
                        Get-PrinterDriver | Select-Object Name, MajorVersion, Environment | Out-String | Write-Host
                    } catch {
                        Write-Host "Could not query Get-PrinterDriver: $_"
                    }
                    
                    Write-Host "2. Staged printer driver packages in Windows Driver Store (PnPUtil):"
                    try {
                        $drivers = & $PnpUtilPath /enum-drivers 2>&1
                        $printerDrivers = @()
                        $currentDriver = @{}
                        
                        foreach ($line in $drivers) {
                            $trimmed = $line.Trim()
                            if ($trimmed -match '^Published Name:\s*(.*)') {
                                if ($currentDriver.Keys.Count -gt 0) {
                                    $printerDrivers += [PSCustomObject]$currentDriver
                                }
                                $currentDriver = @{ 'PublishedName' = $Matches[1] }
                            }
                            elseif ($trimmed -match '^Original Name:\s*(.*)') { $currentDriver['OriginalName'] = $Matches[1] }
                            elseif ($trimmed -match '^Provider Name:\s*(.*)') { $currentDriver['Provider'] = $Matches[1] }
                            elseif ($trimmed -match '^Class Name:\s*(.*)') { $currentDriver['Class'] = $Matches[1] }
                            elseif ($trimmed -match '^Driver Date and Version:\s*(.*)') { $currentDriver['Version'] = $Matches[1] }
                        }
                        if ($currentDriver.Keys.Count -gt 0) {
                            $printerDrivers += [PSCustomObject]$currentDriver
                        }
                        
                        $printerDrivers | Where-Object { $_.Class -eq "Printers" -or $_.Class -eq "Printer" } | Format-Table -AutoSize | Out-String | Write-Host
                    } catch {
                        Write-Host "Could not query PnPUtil driver store: $_"
                    }
                    Write-Host "==============================================================`n"
                }
            }
        } else {
            Write-PrinterLog -Message "Driver '$driver' already registered in spooler." -Level "INFO"
        }
    }

    # 3. Create TCP/IP Ports and Add Printers
    Write-PrinterLog -Message "[3/4] Creating ports and installing printers..." -Level "INFO"
    foreach ($printer in $Printers) {
        $PortName = $printer.Name # Matching port name to printer name is standard practice
        
        # Create standard TCP/IP port
        $existingPort = try { Get-PrinterPort -Name $PortName -ErrorAction Stop } catch { $null }
        if (-not $existingPort) {
            if ($PSCmdlet.ShouldProcess("$PortName ($($printer.PortName))", "Create TCP/IP Printer Port")) {
                Write-PrinterLog -Message "Creating TCP/IP Port: $PortName ($($printer.PortName))" -Level "INFO"
                $PrinterPortOptions = @{
                    Name               = $PortName
                    PrinterHostAddress = $printer.PortName
                    PortNumber         = '9100'
                }
                try {
                    Add-PrinterPort @PrinterPortOptions -ErrorAction Stop | Out-Null
                    Write-PrinterLog -Message "Successfully created printer port: $PortName" -Level "SUCCESS"
                } catch {
                    $InstallFailed = $true
                    Write-PrinterLog -Message "Failed to create printer port $($PortName): $_" -Level "ERROR"
                }
            }
        } else {
            Write-PrinterLog -Message "Port $PortName already exists." -Level "INFO"
        }
        
        # Create the printer globally
        $existingPrinter = try { Get-Printer -Name $printer.Name -ErrorAction Stop } catch { $null }
        if (-not $existingPrinter) {
            if ($PSCmdlet.ShouldProcess($printer.Name, "Install Printer and apply default Print Configuration")) {
                Write-PrinterLog -Message "Adding printer: $($printer.Name)" -Level "INFO"
                $PrinterAddOptions = @{
                    Comment    = $printer.Comment
                    DriverName = $printer.DriverName
                    Location   = $printer.Location
                    Name       = $printer.Name
                    PortName   = $PortName
                }
                try {
                    Add-Printer @PrinterAddOptions -ErrorAction Stop | Out-Null
                    Write-PrinterLog -Message "Successfully added printer: $($printer.Name)" -Level "SUCCESS"
                    
                    # 4. Configure Printer Settings (e.g., Color default = 0 (Mono), Duplex = TwoSidedLongEdge)
                    Write-PrinterLog -Message "[4/4] Configuring settings for: $($printer.Name)" -Level "INFO"
                    $PrinterConfigOptions = @{
                        Color         = 0
                        DuplexingMode = 'TwoSidedLongEdge'
                        PrinterName   = $printer.Name
                    }
                    try {
                        Set-PrintConfiguration @PrinterConfigOptions -ErrorAction Stop
                        Write-PrinterLog -Message "Successfully configured print preferences." -Level "SUCCESS"
                    } catch {
                        Write-PrinterLog -Message "Printer added, but could not apply default print preferences (Duplex/Mono) for $($printer.Name): $_" -Level "WARN"
                    }
                } catch {
                    $InstallFailed = $true
                    Write-PrinterLog -Message "Failed to add printer $($printer.Name): $_" -Level "ERROR"
                }
            }
        } else {
            Write-PrinterLog -Message "Printer $($printer.Name) already installed." -Level "INFO"
        }
    }
} finally {
    # Restore Program Compatibility Assistant Service (PcaSvc) if it was running previously
    if ($PcaServiceStarted) {
        if ($PSCmdlet.ShouldProcess("PcaSvc", "Restore Program Compatibility Assistant Service")) {
            Write-PrinterLog -Message "Restoring Program Compatibility Assistant Service (PcaSvc)..." -Level "INFO"
            try { Start-Service -Name "PcaSvc" -ErrorAction Stop | Out-Null } catch {}
        }
    }
    Write-PrinterLog -Message "=== Printer Installation Script Finished ===" -Level "INFO"
}

try { Stop-Transcript -ErrorAction Stop } catch {}
if ($InstallFailed) {
    exit 1
}
'@

        $uninstallTemplateText = @'
<#
.SYNOPSIS
    Removes printer configurations, standard ports, and drivers globally from the system.
.DESCRIPTION
    Reads printers.csv, removes the associated printers, deletes standard TCP/IP printer ports, and removes the printer driver registration from the driver store.
.NOTES
    Script Name   : Uninstall-Printer.ps1
    Author        : JP
    Version       : 1.4.3
    Context       : SYSTEM (Intune Win32 App Uninstall Behavior: System)
    Log Path      : C:\Logs\Printer_<PrinterName>\Uninstall-Printer.log
.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File .\Uninstall-Printer.ps1
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param()

# Read printers.csv as input
$Printers = Import-Csv .\printers.csv
$FirstPrinterName = $Printers[0].Name

# Sanitize the printer name for file system compatibility (replacing invalid characters like '|' with '-')
$SanitizedName = ($FirstPrinterName -replace '[\\/:*?"<>|]', '-' -replace '\s+', ' ').Trim()

# Create log folder if it doesn't exist
$LogFolder = "$env:SystemDrive\Logs\Printer_$SanitizedName"
if (-not (Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}
$LogFile = "$LogFolder\Uninstall-Printer.log"

function Write-PrinterLog {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")][string]$Level = "INFO"
    )
    $TimeStamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
    $Entry = "[$TimeStamp] [$Level] $Message"
    Write-Output $Entry
    try { Add-Content -Path $LogFile -Value $Entry -Encoding UTF8 -ErrorAction Stop } catch {}
}

try { Start-Transcript -Path $LogFile -Append -ErrorAction Stop } catch {}
Write-PrinterLog -Message "=== Starting Printer Uninstallation ($FirstPrinterName) ===" -Level "INFO"

# Loop through all printers in the csv-file and remove the printers and their ports
foreach ($printer in $Printers) {
    # Remove printers
    $existingPrinter = try { Get-Printer -Name $printer.Name -ErrorAction Stop } catch { $null }
    if ($existingPrinter) {
        if ($PSCmdlet.ShouldProcess($printer.Name, "Remove Printer Queue")) {
            Write-PrinterLog -Message "Removing printer: $($printer.Name)" -Level "INFO"
            try {
                Remove-Printer -Name $printer.Name -Confirm:$false -ErrorAction Stop
                Write-PrinterLog -Message "Successfully removed printer: $($printer.Name)" -Level "SUCCESS"
            } catch {
                Write-PrinterLog -Message "Failed to remove printer $($printer.Name): $_" -Level "ERROR"
            }
        }
    } else {
        Write-PrinterLog -Message "Printer $($printer.Name) not found, skipping removal." -Level "INFO"
    }
    
    if (-not $WhatIfPreference) {
        Start-Sleep -Seconds 5
    }
    
    # Remove standard TCP/IP port
    $PortName = $printer.Name
    if (Get-PrinterPort -ComputerName $env:COMPUTERNAME | Where-Object Name -EQ $PortName) {
        if ($PSCmdlet.ShouldProcess($PortName, "Remove TCP/IP Printer Port")) {
            Write-PrinterLog -Message "Removing printer port: $PortName" -Level "INFO"
            try {
                Remove-PrinterPort -Name $PortName -ComputerName $env:COMPUTERNAME -Confirm:$false -ErrorAction Stop
                Write-PrinterLog -Message "Successfully removed printer port: $PortName" -Level "SUCCESS"
            } catch {
                Write-PrinterLog -Message "Failed to remove printer port $($PortName): $_" -Level "ERROR"
            }
        }
    } else {
        Write-PrinterLog -Message "Printer port $PortName not found, skipping removal." -Level "INFO"
    }
}

# Remove drivers from the spooler and driver store
foreach ($driver in $Printers.DriverName | Select-Object -Unique) {
    $existingDriver = try { Get-PrinterDriver -Name $driver -ErrorAction Stop } catch { $null }
    if ($existingDriver) {
        if ($PSCmdlet.ShouldProcess($driver, "Remove Printer Driver from Spooler and Driver Store")) {
            Write-PrinterLog -Message "Removing driver from Spooler: $driver" -Level "INFO"
            $PrinterDriverRemoveOptions = @{
                Confirm               = $false
                Computername          = $env:COMPUTERNAME
                Name                  = $driver
                RemoveFromDriverStore = $true
            }
            try {
                Remove-PrinterDriver @PrinterDriverRemoveOptions -ErrorAction Stop
                Write-PrinterLog -Message "Successfully removed driver $driver from store." -Level "SUCCESS"
            } catch {
                Write-PrinterLog -Message "Could not completely remove driver $driver from store: $_" -Level "WARN"
            }
        }
    } else {
        Write-PrinterLog -Message "Driver '$driver' not found in spooler, skipping removal." -Level "INFO"
    }
}

Write-PrinterLog -Message "=== Printer Uninstallation Script Finished ===" -Level "INFO"
try { Stop-Transcript -ErrorAction Stop } catch {}
'@

        # Write templates to staging
        $installTemplateText | Set-Content -Path (Join-Path $tempFolder "Install-Printer.ps1") -Encoding utf8 -Force
        $uninstallTemplateText | Set-Content -Path (Join-Path $tempFolder "Uninstall-Printer.ps1") -Encoding utf8 -Force

        # Dynamically compile printers.csv
        Update-Log "Compiling printers.csv configuration..."
        $escapedName = $printerName -replace '"', '""'
        $escapedDriver = $driverName -replace '"', '""'
        $escapedIP = $printerIP -replace '"', '""'
        $csvContent = @"
Name,DriverName,PortName,Comment,Location
`"$escapedName`",`"$escapedDriver`",`"$escapedIP`",`"Customer Service`",`"Customer Service`"
"@
        $csvContent | Set-Content -Path (Join-Path $tempFolder "printers.csv") -Encoding utf8 -Force

        # Dynamically customize and stage Detection.ps1
        Update-Log "Customizing Intune detection rules..."
        $detectionTemplateText = @'
<#
.SYNOPSIS
    Detection script for Intune to verify if the printer is installed globally on the machine.
.DESCRIPTION
    Checks the registry path first (instant & reliable under SYSTEM context), then falls back to Get-Printer with a short retry loop to handle print spooler latency.
.NOTES
    Script Name   : Detection.ps1
    Author        : JP
    Version       : 1.4.3
    Context       : SYSTEM (Intune Custom Detection Script)
    Log Path      : C:\Logs\Printer_<PrinterName>\Detect-Printer.log
#>
$printers = @(
    '__PRINTER_NAME__'
)

$SanitizedName = ($printers[0] -replace '[\\/:*?"<>|]', '-' -replace '\s+', ' ').Trim()
$LogFolder     = "$env:SystemDrive\Logs\Printer_$SanitizedName"
if (-not (Test-Path $LogFolder)) {
    try { New-Item -ItemType Directory -Path $LogFolder -Force -ErrorAction Stop | Out-Null } catch {}
}
$LogFile = "$LogFolder\Detect-Printer.log"

function Write-DetectLog ([string]$Message) {
    $TimeStamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
    try { Add-Content -Path $LogFile -Value "[$TimeStamp] [DETECTION] $Message" -Encoding UTF8 -ErrorAction Stop } catch {}
}

$numberofprintersfound = 0
foreach ($printer in $printers) {
    # 1. Try registry key check first (fastest and most reliable under SYSTEM account)
    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\$printer"
    if (Test-Path $regPath) {
        Write-DetectLog "Detected printer in registry: $printer"
        $numberofprintersfound++
        continue
    }

    # 2. Fallback to Get-Printer with a short retry loop to accommodate print spooler delays
    for ($i = 0; $i -lt 3; $i++) {
        try {
            if (Get-Printer -Name $printer -ErrorAction Stop) {
                Write-DetectLog "Detected printer via Get-Printer: $printer"
                $numberofprintersfound++
                break
            }
        }
        catch {
            Start-Sleep -Seconds 2
        }
    }
}

# If all printers are detected, exit 0
if ($numberofprintersfound -eq $printers.count) {
    Write-DetectLog "STATUS: INSTALLED (All $($printers.count) printer(s) found)."
    Write-Host "All printers were found"
    exit 0
}
else {
    Write-DetectLog "STATUS: NOT INSTALLED ($numberofprintersfound of $($printers.count) printer(s) found)."
    Write-Host "Not all $($printers.count) printers were found"
    exit 1
}
'@
        $detectionScript = $detectionTemplateText -replace '__PRINTER_NAME__', ($printerName -replace "'", "''")
        $detectionScript | Set-Content -Path (Join-Path $tempFolder "Detection.ps1") -Encoding utf8 -Force

        # Copy Driver Files
        Update-Log "Staging driver files (this may take a few seconds)..."
        $destDriverDir = Join-Path $tempFolder "Driver"
        New-Item -ItemType Directory -Path $destDriverDir -Force | Out-Null
        Copy-Item -Path "$driverPath\*" -Destination $destDriverDir -Recurse -Force

        # Run packaging utility
        Update-Log "Staging complete. Packaging with IntuneWinAppUtil.exe..."
        
        $processArgs = @(
            "-c", "`"$tempFolder`"",
            "-s", "Install-Printer.ps1",
            "-o", "`"$printerOutputDir`"",
            "-q"
        )
        
        # Create temp files for standard output and error redirection
        $stdoutFile = [System.IO.Path]::GetTempFileName()
        $stderrFile = [System.IO.Path]::GetTempFileName()

        if ($packerExe -is [System.Array]) {
            $packerExe = ($packerExe | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Last 1)
        }
        $packerExeStr = [string]$packerExe

        # Start the process redirecting to files
        $process = Start-Process -FilePath $packerExeStr -ArgumentList ($processArgs -join " ") -NoNewWindow -PassThru -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile

        # Access the handle to force cache population (known .NET process handle caching quirk)
        $null = $process.Handle

        $lastReadOffset = 0
        while (-not $process.HasExited) {
            # Read new lines from stdout file and print them to the GUI log console
            if (Test-Path $stdoutFile) {
                $content = try { Get-Content -Path $stdoutFile -ErrorAction Stop } catch { $null }
                if ($content) {
                    $newLines = $content[$lastReadOffset..($content.Count - 1)]
                    foreach ($line in $newLines) {
                        if ($line) { Update-Log $line }
                    }
                    $lastReadOffset = $content.Count
                }
            }

            # Run WPF dispatcher queue to keep GUI responsive and prevent freezing
            $null = [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [Action]{})
            Start-Sleep -Milliseconds 250
        }

        # Read final remaining stdout lines
        if (Test-Path $stdoutFile) {
            $content = try { Get-Content -Path $stdoutFile -ErrorAction Stop } catch { $null }
            if ($content -and $content.Count -gt $lastReadOffset) {
                $newLines = $content[$lastReadOffset..($content.Count - 1)]
                foreach ($line in $newLines) {
                    if ($line) { Update-Log $line }
                }
            }
        }

        $stderr = ""
        if (Test-Path $stderrFile) {
            $stderr = try { Get-Content -Path $stderrFile -Raw -ErrorAction Stop } catch { "" }
        }

        # Ensure the process is completely finished
        $process.WaitForExit()

        # Retrieve the exit code with fallback logic for the handle caching issue
        $exitCode = $process.ExitCode
        if ($null -eq $exitCode) {
            $defaultOutputFile = Join-Path $printerOutputDir "Install-Printer.intunewin"
            if (Test-Path $defaultOutputFile) {
                $exitCode = 0
            } else {
                $exitCode = 1
            }
        }

        # Cleanup temp files
        if (Test-Path -LiteralPath $stdoutFile) { try { Remove-Item -Path $stdoutFile -Force -ErrorAction Stop | Out-Null } catch {} }
        if (Test-Path -LiteralPath $stderrFile) { try { Remove-Item -Path $stderrFile -Force -ErrorAction Stop | Out-Null } catch {} }

        if ($exitCode -ne 0) {
            throw "IntuneWinAppUtil exited with code $exitCode. Error details: $stderr"
        }

        # Rename the default Install-Printer.intunewin output to the custom sanitized printer name
        Update-Log "Renaming package file..."
        
        $defaultOutputFile = Join-Path $printerOutputDir "Install-Printer.intunewin"
        $targetOutputFile = Join-Path $printerOutputDir $finalFileName

        if (Test-Path $defaultOutputFile) {
            if (Test-Path $targetOutputFile) {
                Remove-Item -Path $targetOutputFile -Force
            }
            Rename-Item -Path $defaultOutputFile -NewName $finalFileName -Force
        } elseif (-not (Test-Path $targetOutputFile)) {
            $generated = Get-ChildItem -Path $printerOutputDir -Filter "*.intunewin" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($generated) {
                Move-Item -Path $generated.FullName -Destination $targetOutputFile -Force
            }
        }

        if (Test-Path $targetOutputFile) {
            # Export Detection.ps1 to the output folder
            $exportedDetectionFile = Join-Path $printerOutputDir "$($sanitized)_Detection.ps1"
            if (Test-Path (Join-Path $tempFolder "Detection.ps1")) {
                if (Test-Path $exportedDetectionFile) {
                    Remove-Item -Path $exportedDetectionFile -Force
                }
                Copy-Item -Path (Join-Path $tempFolder "Detection.ps1") -Destination $exportedDetectionFile -Force
            }

            # Export macOS Install Shell Script
            $macScriptFile = Join-Path $printerOutputDir "$($sanitized)_macOS_Install.sh"
            $macOsTemplateText = @'
#!/bin/bash
# ======================================================================
# macOS Printer Installation Script
# Generated by Intune Printer Packager v1.4.3
# ======================================================================

# Variables
QUEUE_NAME="__QUEUE_NAME__"
DISPLAY_NAME="__DISPLAY_NAME__"
PRINTER_IP="__PRINTER_IP__"
DRIVER_NAME="__DRIVER_NAME__"
LOCATION="Customer Service"

echo "======================================================"
echo "Starting installation of printer: $DISPLAY_NAME"
echo "======================================================"

# 1. Search for the PPD file in the standard macOS PPD directories
echo "Searching for driver PPD file matching: $DRIVER_NAME..."
CLEAN_TERM=$(echo "$DRIVER_NAME" | sed -e 's/ PCL[0-9]*//g' -e 's/ KX//g' -e 's/ PS//g' -e 's/ PPD//g' -e 's/ KPDL//g')
PPD_PATH=""

SEARCH_PATHS=(
  "/Library/Printers/PPDs/Contents/Resources"
  "/Library/Printers/PPDs/Contents/Resources/en.lproj"
)

for path in "${SEARCH_PATHS[@]}"; do
  if [ -d "$path" ]; then
    MATCH=$(find "$path" -type f -iname "*${CLEAN_TERM}*" -print -quit 2>/dev/null)
    if [ -n "$MATCH" ]; then
      PPD_PATH="$MATCH"
      break
    fi
  fi
done

if [ -n "$PPD_PATH" ]; then
  echo "Found PPD driver file at: $PPD_PATH"
else
  echo "WARNING: Could not find PPD file for '$DRIVER_NAME' in standard folders."
  echo "Please ensure the manufacturer driver PKG is uploaded and installed first."
  echo "Falling back to generic PostScript driver..."
  PPD_PATH="/System/Library/Frameworks/ApplicationServices.framework/Versions/A/Frameworks/PrintCore.framework/Versions/A/Resources/Generic.ppd"
fi

# 2. Check if printer queue already exists and remove it if so (overwrites)
if lpstat -p "$QUEUE_NAME" &>/dev/null; then
  echo "Existing printer queue '$QUEUE_NAME' found. Removing it first..."
  lpadmin -x "$QUEUE_NAME"
fi

# 3. Add the printer using lpadmin
echo "Creating printer queue..."
lpadmin -p "$QUEUE_NAME" \
        -D "$DISPLAY_NAME" \
        -L "$LOCATION" \
        -E \
        -v "lpd://$PRINTER_IP" \
        -P "$PPD_PATH" \
        -o printer-is-shared=false

# 4. Apply standard settings (e.g. Duplex by default, Gray/Color if supported)
echo "Applying printer options..."
lpadmin -p "$QUEUE_NAME" -o Duplex=DuplexNoTumble 2>/dev/null
lpadmin -p "$QUEUE_NAME" -o ColorModel=Gray 2>/dev/null

# 5. Verify queue exists
if lpstat -p "$QUEUE_NAME" &>/dev/null; then
  echo "SUCCESS: Printer '$DISPLAY_NAME' has been successfully installed!"
  exit 0
else
  echo "ERROR: Failed to install printer queue."
  exit 1
fi
'@
            $macScriptContent = $macOsTemplateText -replace '__QUEUE_NAME__', $macQueueName
            $macScriptContent = $macScriptContent -replace '__DISPLAY_NAME__', ($printerName -replace "'", "''")
            $macScriptContent = $macScriptContent -replace '__PRINTER_IP__', $printerIP
            $macScriptContent = $macScriptContent -replace '__DRIVER_NAME__', $driverName
            
            # Make sure it's written with LF endings for Unix compatibility
            [System.IO.File]::WriteAllText($macScriptFile, ($macScriptContent -replace "`r`n", "`n"))

            # Generate copy-pasteable instruction txt file aligned with the Intune upload wizard
            $instructionsFile = Join-Path $printerOutputDir "$($sanitized)_Instructions.txt"
            $generatedTimestamp = Get-Date -Format "dd/MM/yy HH:mm:ss"
            $createdDateStr = Get-Date -Format "dd/MM/yy"
            $instructionsContent = @"
========================================================================
MICROSOFT INTUNE WIN32 APP DEPLOYMENT GUIDE
Generated : $generatedTimestamp
========================================================================

STEP 1: SELECT APP PACKAGE FILE
------------------------------------------------------------------------
App package file:
$finalFileName


STEP 2: APP INFORMATION
------------------------------------------------------------------------
Name:
Printer - $printerName

Description (Copy & Paste everything between the === lines into Intune):
======================================================================
Deploys TCP/IP printer '$printerName' with driver
'$driverName' on IP port $printerIP for all users.
All runtime logs are saved to C:\Logs\Printer_$sanitized\

PACKAGE DETAILS:
- Printer Name : $printerName
- Driver Name  : $driverName
- IP Address   : $printerIP
- Author       : JP
- App Version  : 1.4.3
- Created      : $createdDateStr

WORKSTATION LOG LOCATION:
C:\Logs\Printer_$sanitized\

LOG FILES IN THIS FOLDER:
- Install Log   : C:\Logs\Printer_$sanitized\Install-Printer.log
- Detection Log : C:\Logs\Printer_$sanitized\Detect-Printer.log
- Uninstall Log : C:\Logs\Printer_$sanitized\Uninstall-Printer.log
======================================================================

Publisher:
Pheneox

App Version:
1.4.3


STEP 3: PROGRAM
------------------------------------------------------------------------
Install command:
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File .\Install-Printer.ps1

Uninstall command:
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File .\Uninstall-Printer.ps1

Install behavior:
System

Device restart behavior:
No specific action


STEP 4: REQUIREMENTS
------------------------------------------------------------------------
Operating system architecture:
64-bit

Minimum operating system:
Windows 10 1909


STEP 5: DETECTION RULES
------------------------------------------------------------------------
Rules format:
Use a custom detection script

Script file:
$($sanitized)_Detection.ps1

Run script as 32-bit process on 64-bit clients:
No

Enforce script signature check and run script silently:
No


========================================================================
ADDITIONAL NOTES & macOS INTUNE DEPLOYMENT
========================================================================
1. NETWORK REQUIREMENTS:
   [IMPORTANT] Ensure you create a DHCP IP Reservation for the printer
   IP ($printerIP) on your network so it remains static!

2. macOS INTUNE DEPLOYMENT (SHELL SCRIPT METHOD):
   File: $($sanitized)_macOS_Install.sh

   Step A: Prepare & Deploy the Manufacturer Driver (PKG or DMG)
   - Drivers are downloaded from the vendor (Kyocera, Sharp, HP, etc.)
     as a .dmg (Disk Image) or .pkg (Package) installer.
   - If downloaded as a .dmg:
     * Mount the .dmg file on a Mac (or open it using 7-Zip on Windows).
     * Extract the underlying installer .pkg file (e.g., Kyocera OS X 10.9+.pkg).
   - In the Microsoft Intune admin center:
     * Go to Apps > macOS apps > Add.
     * Select "macOS app (PKG)" as the App type.
     * Upload the extracted .pkg driver installer.
     * Fill in App Details (Name, Description, Publisher).
     * Set Minimum Operating System (e.g., macOS 12.0 Monterey).
     * Assign as "Required" to your target macOS devices group.
   - This installs the drivers and copies the PPD files to /Library/Printers/PPDs/.

   Step B: Deploy the Printer Installation Script
   - In the Microsoft Intune admin center:
     * Go to Devices > macOS > Shell scripts > Add.
     * Upload the generated script: $($sanitized)_macOS_Install.sh
     * Configure settings:
       - Run script as signed-in user: No (Must run as root to use lpadmin)
       - Hide script notifications on devices: Yes
       - Script frequency: Run once
       - Max retries if script fails: 3
     * Assign as "Required" to the same macOS device group.
========================================================================
"@
            $instructionsContent | Set-Content -Path $instructionsFile -Encoding utf8 -Force

            $successMsg = @"
Success! IntuneWin package generated successfully.
File saved to: $targetOutputFile
Instructions saved to: $instructionsFile
Detection Script saved to: $exportedDetectionFile
macOS Script saved to: $macScriptFile

$instructionsContent
"@
            Update-Log $successMsg "#34D399"
            [System.Windows.MessageBox]::Show("Successfully generated IntuneWin package, instructions, detection, and macOS script!`n`nFile: $finalFileName`nLocation: $printerOutputDir", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        } else {
            throw "The expected packaging file '$finalFileName' was not found in the output directory."
        }

    } catch {
        Update-Log "An error occurred during build:`n$_" "#F87171"
        [System.Windows.MessageBox]::Show("An error occurred during the packaging process:`n`n$_", "Build Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
    } finally {
        # Cleanup temporary workspace
        if (Test-Path $tempFolder) {
            Update-Log "Cleaning up temporary staging directory..."
            try { Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction Stop } catch {}
        }

        # Restore UI controls
        $btnBuild.IsEnabled = $true
        $btnBrowseDriver.IsEnabled = $true
        $btnBrowseOutput.IsEnabled = $true
        [System.Windows.Input.Mouse]::OverrideCursor = $null
    }
})

# Show Form
$Form.ShowDialog() | Out-Null
