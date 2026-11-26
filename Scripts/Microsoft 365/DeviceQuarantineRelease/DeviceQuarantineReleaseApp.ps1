#Requires -Version 5.1
<#
.SYNOPSIS
    Exchange Online Mobile Device Quarantine Release GUI Tool
.DESCRIPTION
    A desktop GUI tool to connect to Exchange Online, search and select licensed mailbox users,
    check their mobile device quarantine status, and release devices to the Allowed list.
.NOTES
    Script Name   : DeviceQuarantineReleaseApp.ps1
    Author        : JP
    Version       : 1.1.0
    Log Path      : C:\Logs\DeviceQuarantineRelease\DeviceQuarantineRelease.log
#>

# Enforce STA Threading for WPF
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $scriptPath = if ($PSScriptRoot) { Join-Path $PSScriptRoot $MyInvocation.MyCommand.Name } else { $PSCommandPath }
    Start-Process powershell.exe -ArgumentList "-STA -ExecutionPolicy Bypass -NoProfile -File `"$scriptPath`""
    exit
}

# Load WPF assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

function Initialize-StandardGuiWindow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Window]$Window,
        [double]$MinWidth = 860,
        [double]$MinHeight = 640
    )

    $fontFiles = @('Figtree-Regular.ttf', 'Figtree-SemiBold.ttf', 'Figtree-Bold.ttf')
    $fontDir = $null

    if ($PSScriptRoot) {
        $probe = $PSScriptRoot
        for ($i = 0; $i -le 5; $i++) {
            if ([string]::IsNullOrWhiteSpace($probe)) { break }
            $candidateDir = Join-Path -Path $probe -ChildPath 'assets\fonts'
            $candidateRegular = Join-Path -Path $candidateDir -ChildPath 'Figtree-Regular.ttf'
            if (Test-Path -Path $candidateRegular -PathType Leaf) {
                $fontDir = $candidateDir
                break
            }
            $probe = Split-Path -Path $probe -Parent
        }
    }

    if (-not $fontDir) {
        $cacheFontDir = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'PowerShellGui\fonts'
        $cachedRegular = Join-Path -Path $cacheFontDir -ChildPath 'Figtree-Regular.ttf'
        if (-not (Test-Path -Path $cachedRegular -PathType Leaf)) {
            try {
                if (-not (Test-Path -Path $cacheFontDir)) {
                    New-Item -Path $cacheFontDir -ItemType Directory -Force | Out-Null
                }
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
                $prevProgress = $ProgressPreference
                $ProgressPreference = 'SilentlyContinue'
                try {
                    $baseUrl = 'https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/assets/fonts/'
                    foreach ($f in $fontFiles) {
                        $dest = Join-Path -Path $cacheFontDir -ChildPath $f
                        if (-not (Test-Path -Path $dest -PathType Leaf)) {
                            Invoke-WebRequest -Uri ($baseUrl + $f) -OutFile $dest -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
                        }
                    }
                } finally {
                    $ProgressPreference = $prevProgress
                }
            } catch {
                Write-Verbose "Figtree font download skipped (offline fallback): $($_.Exception.Message)"
            }
        }
        if (Test-Path -Path $cachedRegular -PathType Leaf) {
            $fontDir = $cacheFontDir
        }
    }

    if ($fontDir) {
        try {
            $resolvedFontDir = (Resolve-Path -Path $fontDir -ErrorAction Stop).Path
            $fontUri = ([System.Uri]::new((Join-Path -Path $resolvedFontDir -ChildPath ''))).AbsoluteUri
            $Window.FontFamily = New-Object System.Windows.Media.FontFamily ("$($fontUri)#Figtree, Segoe UI Variable, Segoe UI")
        } catch {
            $Window.FontFamily = New-Object System.Windows.Media.FontFamily ('Segoe UI Variable, Segoe UI')
        }
    } else {
        $Window.FontFamily = New-Object System.Windows.Media.FontFamily ('Segoe UI Variable, Segoe UI')
    }

    $workArea = [System.Windows.SystemParameters]::WorkArea
    $targetWidth  = [Math]::Round([Math]::Min($workArea.Width - 24,  [Math]::Max(860, [Math]::Min(1440, $workArea.Width * 0.70))))
    $targetHeight = [Math]::Round([Math]::Min($workArea.Height - 24, [Math]::Max(640, [Math]::Min(900,  $workArea.Height * 0.80))))

    $Window.WindowStartupLocation = 'Manual'
    $Window.Width  = $targetWidth
    $Window.Height = $targetHeight
    $Window.MinWidth  = [Math]::Min($MinWidth, $targetWidth)
    $Window.MinHeight = [Math]::Min($MinHeight, $targetHeight)
    $Window.Left = [Math]::Round($workArea.Left + (($workArea.Width - $targetWidth) / 2))
    $Window.Top  = [Math]::Round($workArea.Top + (($workArea.Height - $targetHeight) / 2))
    $Window.ResizeMode = 'CanResize'
}

# Resolve log directory (UK date format dd/MM/yy HH:mm:ss)
$script:LogDirectory = 'C:\Logs\DeviceQuarantineRelease'
try {
    if (-not (Test-Path -Path $script:LogDirectory)) {
        New-Item -ItemType Directory -Path $script:LogDirectory -Force -ErrorAction Stop | Out-Null
    }
} catch {
    $script:LogDirectory = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'Logs\DeviceQuarantineRelease'
    if (-not (Test-Path -Path $script:LogDirectory)) {
        try {
            New-Item -ItemType Directory -Path $script:LogDirectory -Force -ErrorAction Stop | Out-Null
        } catch {}
    }
}
$script:LogFilePath = Join-Path -Path $script:LogDirectory -ChildPath ("DeviceQuarantineRelease_{0}.log" -f (Get-Date -Format 'dd-MM-yy'))

# Define XAML layout (Obsidian Teal Standard)
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Exchange Online Quarantine Release Manager v1.1"
        Height="720" Width="1060" MinHeight="640" MinWidth="860"
        WindowStartupLocation="Manual" FontSize="13" Foreground="#F8FAFC">
    <Window.Background>
        <RadialGradientBrush Center="0.5,0.0" GradientOrigin="0.5,0.0" RadiusX="0.95" RadiusY="0.85">
            <GradientStop Color="#0F1F3D" Offset="0.0"/>
            <GradientStop Color="#091122" Offset="0.55"/>
            <GradientStop Color="#05070C" Offset="1.0"/>
        </RadialGradientBrush>
    </Window.Background>

    <Window.Resources>
        <Style x:Key="HeaderStyle" TargetType="TextBlock">
            <Setter Property="FontSize" Value="19"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Foreground" Value="#F8FAFC"/>
        </Style>
        <Style x:Key="SubHeaderStyle" TargetType="TextBlock">
            <Setter Property="FontSize" Value="12.5"/>
            <Setter Property="FontWeight" Value="Normal"/>
            <Setter Property="Foreground" Value="#94A3B8"/>
        </Style>

        <!-- Primary CTA Button (Obsidian Teal) -->
        <Style TargetType="Button">
            <Setter Property="Background" Value="#2DD4BF"/>
            <Setter Property="Foreground" Value="#042F2E"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Border" CornerRadius="6" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#5EEAD4"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#14B8A6"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Border" Property="Background" Value="#1A2B44"/>
                                <Setter Property="Foreground" Value="#64748B"/>
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
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Border" CornerRadius="6" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#24FFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#0DFFFFFF"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Border" Property="Opacity" Value="0.45"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Action CTA Button (Obsidian Teal) -->
        <Style x:Key="ActionButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}"/>

        <!-- Input TextBox Style -->
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#99070D1A"/>
            <Setter Property="Foreground" Value="#F8FAFC"/>
            <Setter Property="BorderBrush" Value="#26354D"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,7"/>
            <Setter Property="CaretBrush" Value="#2DD4BF"/>
            <Setter Property="SelectionBrush" Value="#2DD4BF"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="Border" CornerRadius="6" BorderThickness="{TemplateBinding BorderThickness}"
                                BorderBrush="{TemplateBinding BorderBrush}" Background="{TemplateBinding Background}">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="0"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter TargetName="Border" Property="BorderBrush" Value="#2DD4BF"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Border" Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Dark ScrollBar -->
        <Style TargetType="ScrollBar">
            <Setter Property="Background" Value="#050A14"/>
            <Setter Property="Width" Value="10"/>
            <Setter Property="MinWidth" Value="10"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollBar">
                        <Grid Background="{TemplateBinding Background}">
                            <Track x:Name="PART_Track" IsDirectionReversed="True">
                                <Track.Thumb>
                                    <Thumb>
                                        <Thumb.Template>
                                            <ControlTemplate TargetType="Thumb">
                                                <Border Background="#26354D" CornerRadius="4" Margin="2"/>
                                            </ControlTemplate>
                                        </Thumb.Template>
                                    </Thumb>
                                </Track.Thumb>
                            </Track>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Directory ListBox Style -->
        <Style TargetType="ListBox">
            <Setter Property="Background" Value="#091122"/>
            <Setter Property="Foreground" Value="#E2E8F0"/>
            <Setter Property="BorderBrush" Value="#26354D"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="4"/>
            <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
            <Setter Property="ItemContainerStyle">
                <Setter.Value>
                    <Style TargetType="ListBoxItem">
                        <Setter Property="Padding" Value="10,9"/>
                        <Setter Property="Margin" Value="0,1"/>
                        <Setter Property="Background" Value="Transparent"/>
                        <Setter Property="BorderThickness" Value="1"/>
                        <Setter Property="BorderBrush" Value="Transparent"/>
                        <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="ListBoxItem">
                                    <Border x:Name="Border" CornerRadius="5" Background="{TemplateBinding Background}"
                                            BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}"
                                            Padding="{TemplateBinding Padding}">
                                        <ContentPresenter/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True">
                                            <Setter TargetName="Border" Property="Background" Value="#14FFFFFF"/>
                                        </Trigger>
                                        <Trigger Property="IsSelected" Value="True">
                                            <Setter TargetName="Border" Property="Background" Value="#1E3A5F"/>
                                            <Setter TargetName="Border" Property="BorderBrush" Value="#2DD4BF"/>
                                            <Setter Property="Foreground" Value="#F8FAFC"/>
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                    </Style>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="130"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- HEADER BAR -->
        <Border Grid.Row="0" Background="#B3070D19" BorderBrush="#1FFFFFFF" BorderThickness="0,0,0,1" Padding="20,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel VerticalAlignment="Center">
                    <TextBlock Text="Exchange Online Quarantine Release Manager v1.1" Style="{StaticResource HeaderStyle}"/>
                    <TextBlock Text="Search Exchange Online mailboxes and release quarantined ActiveSync mobile devices." Style="{StaticResource SubHeaderStyle}" Margin="0,3,0,0"/>
                </StackPanel>

                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Border x:Name="StatusBadge" Background="#14FFFFFF" BorderBrush="#26354D" BorderThickness="1" CornerRadius="6" Padding="11,6" Margin="0,0,12,0" VerticalAlignment="Center">
                        <TextBlock x:Name="StatusText" Text="Disconnected" Foreground="#94A3B8" VerticalAlignment="Center" FontWeight="SemiBold" FontSize="12"/>
                    </Border>
                    <Button x:Name="BtnConnect" Content="Connect to Exchange Online" Height="36" MinWidth="200"/>
                    <Button x:Name="BtnDisconnect" Content="Disconnect" Height="36" MinWidth="105" Margin="10,0,0,0" Style="{StaticResource SecondaryButton}" Visibility="Collapsed"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- MAIN WORKSPACE -->
        <Grid Grid.Row="1" Margin="16,14,16,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="340" MinWidth="260"/>
                <ColumnDefinition Width="14"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- LEFT COLUMN: USER DIRECTORY -->
            <Border Grid.Column="0" Background="#660D192E" BorderBrush="#26354D" BorderThickness="1" CornerRadius="8" Padding="14">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- Local Filtering Search Box -->
                    <Grid Grid.Row="0" Margin="0,0,0,10">
                        <TextBox x:Name="TxtSearch" Height="36" VerticalContentAlignment="Center" IsEnabled="False"/>
                        <TextBlock IsHitTestVisible="False" Text="Filter mailbox users..." VerticalAlignment="Center" HorizontalAlignment="Left" Margin="12,0,0,0" Foreground="#94A3B8">
                            <TextBlock.Style>
                                <Style TargetType="{x:Type TextBlock}">
                                    <Setter Property="Visibility" Value="Collapsed"/>
                                    <Style.Triggers>
                                        <DataTrigger Binding="{Binding Text, ElementName=TxtSearch}" Value="">
                                            <Setter Property="Visibility" Value="Visible"/>
                                        </DataTrigger>
                                    </Style.Triggers>
                                </Style>
                            </TextBlock.Style>
                        </TextBlock>
                    </Grid>

                    <!-- Mailbox User List -->
                    <ListBox Grid.Row="1" x:Name="LstUsers" Margin="0,0,0,10" IsEnabled="False">
                        <ListBox.ItemTemplate>
                            <DataTemplate>
                                <StackPanel>
                                    <TextBlock Text="{Binding DisplayName}" FontWeight="SemiBold" Foreground="#F8FAFC" FontSize="13"/>
                                    <TextBlock Text="{Binding PrimarySmtpAddress}" Foreground="#94A3B8" FontSize="11.5" Margin="0,2,0,0"/>
                                </StackPanel>
                            </DataTemplate>
                        </ListBox.ItemTemplate>
                    </ListBox>

                    <!-- Fetch Directory Action -->
                    <Button Grid.Row="2" x:Name="BtnLoadUsers" Content="Load User Directory" Style="{StaticResource SecondaryButton}" Height="36" IsEnabled="False"/>
                </Grid>
            </Border>

            <!-- RIGHT COLUMN: SELECTION & DEVICE PANEL -->
            <Border Grid.Column="2" Background="#660D192E" BorderBrush="#26354D" BorderThickness="1" CornerRadius="8" Padding="16">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <!-- Selected Target Header -->
                    <StackPanel Grid.Row="0" Margin="0,0,0,14">
                        <TextBlock x:Name="TxtSelectedUserName" Text="No Mailbox Selected" FontSize="17" FontWeight="Bold" Foreground="#F8FAFC"/>
                        <TextBlock x:Name="TxtSelectedUserEmail" Text="Select a mailbox from the left panel to display quarantine status." FontSize="12.5" Foreground="#94A3B8" Margin="0,3,0,0"/>
                    </StackPanel>

                    <!-- Quarantined Devices Container -->
                    <Grid Grid.Row="1">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <TextBlock Text="Detected Quarantined Devices" FontWeight="SemiBold" Foreground="#CBD5E1" Margin="0,0,0,8"/>

                        <ListBox Grid.Row="1" x:Name="LstDevices">
                            <ListBox.ItemTemplate>
                                <DataTemplate>
                                    <Border BorderBrush="#1A2B44" BorderThickness="0,0,0,1" Padding="4,6">
                                        <Grid>
                                            <Grid.ColumnDefinitions>
                                                <ColumnDefinition Width="*"/>
                                                <ColumnDefinition Width="Auto"/>
                                            </Grid.ColumnDefinitions>

                                            <StackPanel>
                                                <TextBlock Text="{Binding DeviceModel}" FontWeight="SemiBold" Foreground="#F8FAFC" FontSize="13.5"/>
                                                <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
                                                    <TextBlock Text="Type: " Foreground="#94A3B8" FontSize="11.5"/>
                                                    <TextBlock Text="{Binding DeviceType}" Foreground="#CBD5E1" FontSize="11.5" Margin="0,0,14,0"/>
                                                    <TextBlock Text="ID: " Foreground="#94A3B8" FontSize="11.5"/>
                                                    <TextBlock Text="{Binding DeviceId}" Foreground="#2DD4BF" FontFamily="Cascadia Mono, Consolas" FontSize="11.5"/>
                                                </StackPanel>
                                            </StackPanel>

                                            <Border Grid.Column="1" Background="#26FBBF24" BorderBrush="#66FBBF24" BorderThickness="1" CornerRadius="4" Padding="8,4" VerticalAlignment="Center" Margin="12,0,0,0">
                                                <TextBlock Text="{Binding DeviceAccessState}" Foreground="#FBBF24" FontSize="11" FontWeight="Bold"/>
                                            </Border>
                                        </Grid>
                                    </Border>
                                </DataTemplate>
                            </ListBox.ItemTemplate>
                        </ListBox>

                        <!-- Empty/Feedback States overlaying the ListBox -->
                        <TextBlock Grid.Row="1" x:Name="TxtNoDevices" Text="No quarantined devices found for this user." HorizontalAlignment="Center" VerticalAlignment="Center" Foreground="#94A3B8" Visibility="Collapsed"/>
                        <TextBlock Grid.Row="1" x:Name="TxtLoadingDevices" Text="Querying Mobile Devices..." HorizontalAlignment="Center" VerticalAlignment="Center" Foreground="#2DD4BF" FontWeight="SemiBold" Visibility="Collapsed"/>
                    </Grid>
                </Grid>
            </Border>
        </Grid>

        <!-- LIVE ACTIVITY CONSOLE -->
        <Border Grid.Row="2" Background="#CC050A14" BorderBrush="#26354D" BorderThickness="1" CornerRadius="6" Margin="16,0,16,12" Padding="10,8">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <TextBlock Text="Activity Console" FontSize="11" FontWeight="SemiBold" Foreground="#94A3B8" Margin="0,0,0,4"/>
                <TextBox Grid.Row="1" x:Name="TxtConsole" Background="Transparent" BorderThickness="0" Padding="0"
                         FontFamily="Cascadia Mono, Consolas" FontSize="11.5" Foreground="#CBD5E1"
                         IsReadOnly="True" VerticalScrollBarVisibility="Auto" AcceptsReturn="True" TextWrapping="Wrap"/>
            </Grid>
        </Border>

        <!-- SINGLE-ROW FOOTER -->
        <Border Grid.Row="3" Background="#B3070D19" BorderBrush="#1FFFFFFF" BorderThickness="0,1,0,0" Padding="20,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBlock x:Name="TxtFooterHint" Grid.Column="0" Text="Log: C:\Logs\DeviceQuarantineRelease" Foreground="#94A3B8"
                           FontFamily="Cascadia Mono, Consolas" FontSize="11.5" VerticalAlignment="Center"
                           TextTrimming="CharacterEllipsis" Margin="0,0,16,0"/>

                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button x:Name="BtnCheckDevices" Content="Re-Scan Status" Style="{StaticResource SecondaryButton}" Height="36" MinWidth="140" Margin="0,0,10,0" IsEnabled="False"/>
                    <Button x:Name="BtnRelease" Content="Release Quarantined Devices" Style="{StaticResource ActionButton}" Height="36" MinWidth="220" IsEnabled="False"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Read XAML layout
$reader = (New-Object System.Xml.XmlNodeReader $XAML)
try {
    $Form = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    Write-Host "XAML parse error: $_" -ForegroundColor Red
    exit
}

Initialize-StandardGuiWindow -Window $Form

# Bind elements with Name/x:Name attributes to local PowerShell variables
$XAML.SelectNodes("//*[@*[local-name()='Name']]") | ForEach-Object {
    $attr = $_.Attributes | Where-Object { $_.LocalName -eq 'Name' }
    if ($attr) {
        $controlName = $attr.Value
        $control = $Form.FindName($controlName)
        if ($control) {
            Set-Variable -Name $controlName -Value $control -Scope Script
        }
    }
}

if ($TxtFooterHint -and $script:LogFilePath) {
    $TxtFooterHint.Text = "Log: $($script:LogFilePath)"
}

# Console/Log Writer helper
function Write-Log {
    param(
        [string]$Message,
        [string]$Type = "Info"
    )
    $uiTimestamp = (Get-Date).ToString("HH:mm:ss")
    $fileTimestamp = (Get-Date).ToString("dd/MM/yy HH:mm:ss")
    $prefix = "[$uiTimestamp] "
    switch ($Type) {
        "Error"   { $colorPrefix = "[ERROR] ";   $msg = "$prefix$colorPrefix$Message"; $fileMsg = "[$fileTimestamp] [ERROR] $Message" }
        "Success" { $colorPrefix = "[SUCCESS] "; $msg = "$prefix$colorPrefix$Message"; $fileMsg = "[$fileTimestamp] [SUCCESS] $Message" }
        "Warning" { $colorPrefix = "[WARNING] "; $msg = "$prefix$colorPrefix$Message"; $fileMsg = "[$fileTimestamp] [WARNING] $Message" }
        Default   { $colorPrefix = "[INFO] ";    $msg = "$prefix$colorPrefix$Message"; $fileMsg = "[$fileTimestamp] [INFO] $Message" }
    }

    if ($script:LogFilePath) {
        try {
            Add-Content -Path $script:LogFilePath -Value $fileMsg -Encoding UTF8 -ErrorAction Stop
        } catch {}
    }

    # Run on UI thread to prevent dispatcher exceptions
    $TxtConsole.Dispatcher.Invoke([Action]{
        $TxtConsole.AppendText("$msg`r`n")
        $TxtConsole.ScrollToEnd()
    })
}

# Reset GUI UI elements to initial state
function Reset-UIState {
    $bc = [System.Windows.Media.BrushConverter]::new()
    $StatusBadge.Background  = $bc.ConvertFromString("#14FFFFFF")
    $StatusBadge.BorderBrush = $bc.ConvertFromString("#26354D")
    $StatusText.Text = "Disconnected"
    $StatusText.Foreground = $bc.ConvertFromString("#94A3B8")

    $BtnDisconnect.Visibility = [System.Windows.Visibility]::Collapsed
    $BtnConnect.Visibility = [System.Windows.Visibility]::Visible
    $BtnConnect.IsEnabled = $true

    $TxtSearch.IsEnabled = $false
    $TxtSearch.Text = ""
    $LstUsers.IsEnabled = $false
    $LstUsers.ItemsSource = $null

    $BtnLoadUsers.IsEnabled = $false

    Reset-SelectedUserDetails
}

# Clear selections and secondary states
function Reset-SelectedUserDetails {
    $TxtSelectedUserName.Text = "No Mailbox Selected"
    $TxtSelectedUserEmail.Text = "Select a mailbox from the left panel to display quarantine status."
    $LstDevices.ItemsSource = $null
    $BtnCheckDevices.IsEnabled = $false
    $BtnRelease.IsEnabled = $false
    $TxtNoDevices.Visibility = [System.Windows.Visibility]::Collapsed
    $TxtLoadingDevices.Visibility = [System.Windows.Visibility]::Collapsed
}

# Fetch quarantined devices
function Check-QuarantinedDevices {
    param(
        [string]$UserEmail
    )

    if ([string]::IsNullOrWhiteSpace($UserEmail)) { return }

    Write-Log "Checking mobile devices for: $UserEmail"

    # Update UI to loading state
    $LstDevices.ItemsSource = $null
    $TxtNoDevices.Visibility = [System.Windows.Visibility]::Collapsed
    $TxtLoadingDevices.Visibility = [System.Windows.Visibility]::Visible
    $BtnCheckDevices.IsEnabled = $false
    $BtnRelease.IsEnabled = $false

    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait

    # Force UI refresh to display loading state
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [Action]{})

    try {
        $devices = Get-MobileDevice -Mailbox $UserEmail -ErrorAction Stop

        # Filter only Quarantined
        $quarantined = $devices | Where-Object { $_.DeviceAccessState -eq "Quarantined" }

        $TxtLoadingDevices.Visibility = [System.Windows.Visibility]::Collapsed

        if ($null -eq $quarantined -or $quarantined.Count -eq 0) {
            Write-Log "No quarantined devices found for $UserEmail." "Info"
            $TxtNoDevices.Visibility = [System.Windows.Visibility]::Visible
            $LstDevices.ItemsSource = $null
            $BtnRelease.IsEnabled = $false
        } else {
            $count = if ($quarantined -is [array]) { $quarantined.Count } else { 1 }
            Write-Log "Found $count quarantined device(s) for $UserEmail." "Warning"

            # Format collection for binding
            $mappedDevices = @()
            foreach ($dev in $quarantined) {
                $mappedDevices += [PSCustomObject]@{
                    DeviceModel       = if ($dev.DeviceModel) { $dev.DeviceModel } else { "Unknown Model" }
                    DeviceType        = if ($dev.DeviceType) { $dev.DeviceType } else { "Unknown Type" }
                    DeviceId          = $dev.DeviceId
                    DeviceAccessState = $dev.DeviceAccessState
                }
            }

            $LstDevices.ItemsSource = $mappedDevices
            $BtnRelease.IsEnabled = $true
        }
    } catch {
        $TxtLoadingDevices.Visibility = [System.Windows.Visibility]::Collapsed
        Write-Log "Failed to check devices: $_" "Error"
    } finally {
        [System.Windows.Input.Mouse]::OverrideCursor = $null
        $BtnCheckDevices.IsEnabled = $true
    }
}

# --- EVENT HANDLERS ---

# Connect to Exchange Online
$BtnConnect.Add_Click({
    Write-Log "Connecting to Exchange Online. Please authenticate via modern login dialog..."
    $BtnConnect.IsEnabled = $false

    # Run Connect
    try {
        # Check and load module if needed
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            Write-Log "ExchangeOnlineManagement module not found. Installing module..." "Warning"
            Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser -AllowClobber | Out-Null
            Write-Log "ExchangeOnlineManagement module installed successfully." "Success"
        }

        # Connect to Exchange
        Connect-ExchangeOnline -ErrorAction Stop

        # Resolve administrator identity
        $adminUser = "Exchange Online"
        try {
            $sessions = Get-ConnectionInfo
            if ($sessions -and $sessions.UserName) {
                $adminUser = $sessions.UserName
            }
        } catch {}

        Write-Log "Connection established successfully with $adminUser." "Success"

        # Update Connection Pane UI
        $bc = [System.Windows.Media.BrushConverter]::new()
        $StatusBadge.Background  = $bc.ConvertFromString("#1F34D399")
        $StatusBadge.BorderBrush = $bc.ConvertFromString("#34D399")
        $StatusText.Text = "Connected ($adminUser)"
        $StatusText.Foreground = $bc.ConvertFromString("#F8FAFC")

        $BtnDisconnect.Visibility = [System.Windows.Visibility]::Visible
        $BtnConnect.Visibility = [System.Windows.Visibility]::Collapsed
        $BtnLoadUsers.IsEnabled = $true

    } catch {
        Write-Log "Authentication failed: $_" "Error"
        $BtnConnect.IsEnabled = $true
    }
})

# Disconnect Session
$BtnDisconnect.Add_Click({
    Write-Log "Disconnecting active Exchange Online session..."
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Write-Log "Session disconnected successfully." "Success"
    } catch {
        Write-Log "Error occurred during session teardown: $_" "Warning"
    }
    Reset-UIState
})

# Load User Directory
$BtnLoadUsers.Add_Click({
    Write-Log "Downloading user directory list from tenant..."
    $BtnLoadUsers.IsEnabled = $false
    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait

    # Force UI refresh to display loading cursor
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [Action]{})

    try {
        Write-Log "Retrieving user mailboxes..."
        $script:allUsers = @()

        try {
            $script:allUsers = Get-EXOMailbox -ResultSize Unlimited -Properties DisplayName, PrimarySmtpAddress |
                               Select-Object DisplayName, PrimarySmtpAddress, UserPrincipalName
        } catch {
            Write-Log "Get-EXOMailbox failed. Querying with fallback Get-Mailbox..." "Warning"
            $script:allUsers = Get-Mailbox -ResultSize Unlimited |
                               Select-Object DisplayName, PrimarySmtpAddress, UserPrincipalName
        }

        if (-not $script:allUsers) {
            Write-Log "No mailbox users returned from active tenant." "Warning"
            $script:allUsers = @()
        }

        Write-Log "Loaded $($script:allUsers.Count) users successfully." "Success"

        # Bind to UI
        $LstUsers.ItemsSource = $script:allUsers
        $LstUsers.IsEnabled = $true
        $TxtSearch.IsEnabled = $true

    } catch {
        Write-Log "Directory query failed: $_" "Error"
    } finally {
        [System.Windows.Input.Mouse]::OverrideCursor = $null
        $BtnLoadUsers.IsEnabled = $true
    }
})

# Filter users directory locally
$TxtSearch.Add_TextChanged({
    if (-not $script:allUsers) { return }
    $filterText = $TxtSearch.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($filterText)) {
        $LstUsers.ItemsSource = $script:allUsers
    } else {
        # Filter collection
        $filtered = $script:allUsers | Where-Object {
            $_.DisplayName -like "*$filterText*" -or
            $_.PrimarySmtpAddress -like "*$filterText*" -or
            $_.UserPrincipalName -like "*$filterText*"
        }
        $LstUsers.ItemsSource = $filtered
    }
})

# Selection change event in list of users
$LstUsers.Add_SelectionChanged({
    $selectedUser = $LstUsers.SelectedItem
    if ($null -eq $selectedUser) {
        Reset-SelectedUserDetails
        return
    }

    $TxtSelectedUserName.Text = $selectedUser.DisplayName
    $email = if ($selectedUser.PrimarySmtpAddress) { $selectedUser.PrimarySmtpAddress } else { $selectedUser.UserPrincipalName }
    $TxtSelectedUserEmail.Text = $email

    $BtnCheckDevices.IsEnabled = $true

    # Automatically query mobile devices when user is clicked
    Check-QuarantinedDevices -UserEmail $email
})

# Re-Scan Status Button Click
$BtnCheckDevices.Add_Click({
    $selectedUser = $LstUsers.SelectedItem
    if ($selectedUser) {
        $email = if ($selectedUser.PrimarySmtpAddress) { $selectedUser.PrimarySmtpAddress } else { $selectedUser.UserPrincipalName }
        Check-QuarantinedDevices -UserEmail $email
    }
})

# Release Devices Button Click
$BtnRelease.Add_Click({
    $selectedUser = $LstUsers.SelectedItem
    if (-not $selectedUser) { return }

    $email = if ($selectedUser.PrimarySmtpAddress) { $selectedUser.PrimarySmtpAddress } else { $selectedUser.UserPrincipalName }
    $quarantinedDevices = $LstDevices.ItemsSource

    if (-not $quarantinedDevices) {
        Write-Log "No quarantined devices available to release." "Warning"
        return
    }

    $BtnRelease.IsEnabled = $false
    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait

    Write-Log "Executing quarantine release sequence for $email..."

    # Force UI refresh
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [Action]{})

    try {
        $IDsToRelease = @()
        foreach ($dev in $quarantinedDevices) {
            $IDsToRelease += $dev.DeviceId
        }

        Write-Log "Target Device ID(s): $($IDsToRelease -join ', ')"

        # Fetch existing allowed devices
        Write-Log "Reading current CASMailbox allowed devices..."
        $casMailbox = Get-CASMailbox -Identity $email -ErrorAction Stop
        $CurrentAllowed = $casMailbox.ActiveSyncAllowedDeviceIDs

        # Merge and de-duplicate list
        $UpdatedList = ($CurrentAllowed + $IDsToRelease) | Select-Object -Unique

        # Apply changes
        Write-Log "Writing updated ActiveSyncAllowedDeviceIDs in Exchange Online..."
        Set-CASMailbox -Identity $email -ActiveSyncAllowedDeviceIDs $UpdatedList -ErrorAction Stop

        Write-Log "Quarantine released successfully for $email!" "Success"

        # Recheck devices list to show they are cleared
        Check-QuarantinedDevices -UserEmail $email

    } catch {
        Write-Log "Release sequence failed: $_" "Error"
        $BtnRelease.IsEnabled = $true
    } finally {
        [System.Windows.Input.Mouse]::OverrideCursor = $null
    }
})

# Teardown Exchange Online session on Close
$Form.Add_Closing({
    Write-Log "Closing Application. Cleaning up..."
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
    } catch {}
})

# Splash Message on Startup
Write-Log "Application initialized. Click 'Connect to Exchange Online' to get started."

# Launch the WPF application
[void]$Form.ShowDialog()
