<#
.SYNOPSIS
    Interactive GUI tool to package TCP/IP printers as Microsoft Intune Win32 Apps (.intunewin).
.NOTES
    Requires Microsoft's IntuneWinAppUtil.exe in the same directory.
    Version: 1.3
    Last Updated: 2026-06-25
#>

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

# WPF XAML Definition
$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Intune Printer Packager Tool v1.3" Height="680" Width="520"
        ResizeMode="NoResize" WindowStartupLocation="CenterScreen"
        Background="#18181B">
    <Window.Resources>
        <!-- Modern Rounded Button Style (Primary Blue) -->
        <Style x:Key="PrimaryButton" TargetType="Button">
            <Setter Property="Background" Value="#3B82F6"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#2563EB"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#1D4ED8"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Background" Value="#3F3F46"/>
                                <Setter Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Secondary Gray Button Style -->
        <Style x:Key="SecondaryButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
            <Setter Property="Background" Value="#4B5563"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#374151"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#1F2937"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Background" Value="#3F3F46"/>
                                <Setter Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Build Package Large Button Style (Green) -->
        <Style x:Key="BuildButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
            <Setter Property="Background" Value="#10B981"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#059669"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#047857"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Background" Value="#3F3F46"/>
                                <Setter Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <!-- Header -->
            <RowDefinition Height="*"/>    <!-- Content Form -->
            <RowDefinition Height="Auto"/> <!-- Build Button -->
        </Grid.RowDefinitions>

        <!-- Header Block -->
        <StackPanel Grid.Row="0" Margin="0,0,0,15">
            <TextBlock Text="Intune Printer Packager v1.3" FontSize="24" FontWeight="Bold" Foreground="#3B82F6" FontFamily="Segoe UI"/>
            <TextBlock Text="Package TCP/IP printers as Intune Win32 Apps in seconds" FontSize="12" Foreground="#A1A1AA" FontFamily="Segoe UI" Margin="0,4,0,0"/>
            <Separator Height="1" Background="#27272A" Margin="0,10,0,0"/>
        </StackPanel>

        <!-- Main Form Container -->
        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
            <StackPanel Margin="0,0,10,0">
                
                <!-- Printer Name Input -->
                <TextBlock Text="Printer Name (e.g., Showroom | Upstairs | Sharp MX-4061)" Foreground="#A1A1AA" FontSize="11" FontWeight="SemiBold" Margin="0,0,0,5"/>
                <Border BorderBrush="#27272A" BorderThickness="1" CornerRadius="4" Background="#202023" Padding="2" Margin="0,0,0,15">
                    <TextBox Name="txtPrinterName" BorderThickness="0" Background="Transparent" Foreground="#FFFFFF" FontSize="13" Padding="4" VerticalContentAlignment="Center" CaretBrush="White"/>
                </Border>

                <!-- Driver Name Input -->
                <TextBlock Text="Driver Name (e.g., SHARP MX-4061 PCL6)" Foreground="#A1A1AA" FontSize="11" FontWeight="SemiBold" Margin="0,0,0,5"/>
                <Border BorderBrush="#27272A" BorderThickness="1" CornerRadius="4" Background="#202023" Padding="2" Margin="0,0,0,15">
                    <TextBox Name="txtDriverName" BorderThickness="0" Background="Transparent" Foreground="#FFFFFF" FontSize="13" Padding="4" VerticalContentAlignment="Center" CaretBrush="White"/>
                </Border>

                <!-- IP Address Input -->
                <TextBlock Text="Printer IP Address (e.g., 192.168.0.20)" Foreground="#A1A1AA" FontSize="11" FontWeight="SemiBold" Margin="0,0,0,5"/>
                <Border BorderBrush="#27272A" BorderThickness="1" CornerRadius="4" Background="#202023" Padding="2" Margin="0,0,0,15">
                    <TextBox Name="txtPrinterIP" BorderThickness="0" Background="Transparent" Foreground="#FFFFFF" FontSize="13" Padding="4" VerticalContentAlignment="Center" CaretBrush="White"/>
                </Border>

                <!-- Driver Source Folder Browse -->
                <TextBlock Text="Printer Driver Folder (Source)" Foreground="#A1A1AA" FontSize="11" FontWeight="SemiBold" Margin="0,0,0,5"/>
                <Grid Margin="0,0,0,15">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <Border Grid.Column="0" BorderBrush="#27272A" BorderThickness="1" CornerRadius="4" Background="#202023" Padding="2" Margin="0,0,10,0">
                        <TextBox Name="txtDriverPath" IsReadOnly="True" BorderThickness="0" Background="Transparent" Foreground="#A1A1AA" FontSize="12" Padding="4" VerticalContentAlignment="Center"/>
                    </Border>
                    <Button Name="btnBrowseDriver" Grid.Column="1" Content="Browse..." Style="{StaticResource SecondaryButton}" FontWeight="SemiBold"/>
                </Grid>

                <!-- Output Folder Browse -->
                <TextBlock Text="Output Folder (Where .intunewin will be saved)" Foreground="#A1A1AA" FontSize="11" FontWeight="SemiBold" Margin="0,0,0,5"/>
                <Grid Margin="0,0,0,20">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <Border Grid.Column="0" BorderBrush="#27272A" BorderThickness="1" CornerRadius="4" Background="#202023" Padding="2" Margin="0,0,10,0">
                        <TextBox Name="txtOutputPath" IsReadOnly="True" BorderThickness="0" Background="Transparent" Foreground="#A1A1AA" FontSize="12" Padding="4" VerticalContentAlignment="Center"/>
                    </Border>
                    <Button Name="btnBrowseOutput" Grid.Column="1" Content="Browse..." Style="{StaticResource SecondaryButton}" FontWeight="SemiBold"/>
                </Grid>

                <!-- Console Status Area -->
                <TextBlock Text="Build Log Console (Selectable)" Foreground="#A1A1AA" FontSize="11" FontWeight="SemiBold" Margin="0,0,0,5"/>
                <Border BorderBrush="#27272A" BorderThickness="1" CornerRadius="4" Background="#09090B" Height="140" Padding="8">
                    <ScrollViewer Name="scrollConsole" VerticalScrollBarVisibility="Auto">
                        <TextBox Name="txtStatus" Text="Ready to configure..." IsReadOnly="True" BorderThickness="0" Background="Transparent" Foreground="#10B981" FontFamily="Consolas" FontSize="11" TextWrapping="Wrap" CaretBrush="White"/>
                    </ScrollViewer>
                </Border>

            </StackPanel>
        </ScrollViewer>

        <!-- Build Button -->
        <Button Name="btnBuild" Grid.Row="2" Content="Build IntuneWin Package" Style="{StaticResource BuildButton}" FontSize="14" FontWeight="Bold" Height="40" Margin="0,15,0,0"/>

    </Grid>
</Window>
'@

# Parse XAML
$reader = (New-Object System.Xml.XmlNodeReader ([xml]$xaml))
$Form = [Windows.Markup.XamlReader]::Load($reader)

# Get UI element handles
$txtPrinterName  = $Form.FindName("txtPrinterName")
$txtDriverName   = $Form.FindName("txtDriverName")
$txtPrinterIP    = $Form.FindName("txtPrinterIP")
$txtDriverPath   = $Form.FindName("txtDriverPath")
$txtOutputPath   = $Form.FindName("txtOutputPath")
$btnBrowseDriver = $Form.FindName("btnBrowseDriver")
$btnBrowseOutput = $Form.FindName("btnBrowseOutput")
$btnBuild        = $Form.FindName("btnBuild")
$txtStatus       = $Form.FindName("txtStatus")
$scrollConsole   = $Form.FindName("scrollConsole")

# Autofill default output folder to the script directory if it runs locally
$txtOutputPath.Text = $PSScriptRoot

# Helper to update console log UI responsively
function Update-Log ($message, $color = "#10B981") {
    if ([string]::IsNullOrWhiteSpace($txtStatus.Text) -or $txtStatus.Text -eq "Ready to configure...") {
        $txtStatus.Text = $message
    } else {
        $txtStatus.Text = $txtStatus.Text + "`r`n" + $message
    }
    $converter = New-Object System.Windows.Media.BrushConverter
    $txtStatus.Foreground = $converter.ConvertFromString($color)
    $scrollConsole.ScrollToEnd()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [Action]{})
}

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

    if (-not (Test-Path $driverPath) -or [string]::IsNullOrWhiteSpace($driverPath)) {
        [System.Windows.MessageBox]::Show("Please select a valid driver directory.", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    # Verify that there are INF files inside the driver directory
    $infFiles = Get-ChildItem -Path $driverPath -Filter "*.inf" -Recurse -ErrorAction SilentlyContinue
    if (-not $infFiles) {
        [System.Windows.MessageBox]::Show("The selected driver folder does not contain any driver definition (.inf) files. Please verify the directory.", "No Drivers Found", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    if (-not (Test-Path $outputPath) -or [string]::IsNullOrWhiteSpace($outputPath)) {
        [System.Windows.MessageBox]::Show("Please select a valid output directory.", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    # Check for IntuneWinAppUtil.exe
    $packerExe = Join-Path $PSScriptRoot "IntuneWinAppUtil.exe"
    if (-not (Test-Path $packerExe)) {
        $msg = "Microsoft Win32 Content Prep Tool (IntuneWinAppUtil.exe) was not found in the script directory.`n`nWould you like to automatically download it from Microsoft's official GitHub repository?"
        $confirm = [System.Windows.MessageBox]::Show($msg, "Packer Tool Missing", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        
        if ($confirm -eq [System.Windows.MessageBoxResult]::Yes) {
            Update-Log "IntuneWinAppUtil.exe is missing. Attempting official download..."
            try {
                # Force TLS 1.2 which is required by GitHub
                [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
                
                $downloadUrl = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe"
                Update-Log "Downloading from official Microsoft GitHub..."
                
                # Show wait cursor during download
                [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
                [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [Action]{})
                
                Invoke-WebRequest -Uri $downloadUrl -OutFile $packerExe -UseBasicParsing -ErrorAction Stop
                
                [System.Windows.Input.Mouse]::OverrideCursor = $null
                Update-Log "Download complete! IntuneWinAppUtil.exe is now ready." "#10B981"
            } catch {
                [System.Windows.Input.Mouse]::OverrideCursor = $null
                Update-Log "Failed to download packer tool: $_" "#EF4444"
                [System.Windows.MessageBox]::Show("Failed to download IntuneWinAppUtil.exe:`n`n$_`n`nPlease download it manually and place it in: $PSScriptRoot", "Download Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
                return
            }
        } else {
            Update-Log "Build aborted. IntuneWinAppUtil.exe is required." "#EF4444"
            return
        }
    }

    # Verify that the input Driver Name matches one of the models defined in the INF files
    Update-Log "Verifying Driver Name against INF configuration..."
    $foundDriverNameInInf = $false
    # First pass: Quick scan for the exact driver name (case-insensitive match)
    foreach ($inf in $infFiles) {
        try {
            $content = Get-Content -Path $inf.FullName -Raw -ErrorAction SilentlyContinue
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
                $lines = Get-Content -Path $inf.FullName -ErrorAction SilentlyContinue
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
        Title="Driver Name Mismatch" Height="440" Width="480"
        ResizeMode="NoResize" WindowStartupLocation="CenterOwner"
        Background="#18181B">
    <Window.Resources>
        <!-- Custom ListBoxItem Style -->
        <Style TargetType="ListBoxItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="Margin" Value="0,1"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBoxItem">
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="2" Padding="{TemplateBinding Padding}">
                            <ContentPresenter />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#2D2D30"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#3B82F6"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Rounded Dialog Button Style -->
        <Style x:Key="DialogButton" TargetType="Button">
            <Setter Property="Background" Value="#3B82F6"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Opacity" Value="0.9"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Opacity" Value="0.8"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <!-- Header -->
            <RowDefinition Height="*"/>    <!-- Candidates List -->
            <RowDefinition Height="Auto"/> <!-- Action Buttons -->
        </Grid.RowDefinitions>

        <!-- Header -->
        <StackPanel Grid.Row="0" Margin="0,0,0,12">
            <TextBlock Text="Driver Name Mismatch" FontSize="16" FontWeight="Bold" Foreground="#EF4444" FontFamily="Segoe UI"/>
            <TextBlock Name="lblDescription" Text="The driver name entered was not found. Select a detected driver below to use, or choose to proceed anyway." FontSize="11" Foreground="#A1A1AA" FontFamily="Segoe UI" TextWrapping="Wrap" Margin="0,4,0,0"/>
        </StackPanel>

        <!-- Candidates ListBox -->
        <Border Grid.Row="1" BorderBrush="#27272A" BorderThickness="1" CornerRadius="4" Background="#09090B" Margin="0,0,0,15">
            <ListBox Name="lstCandidates" Background="Transparent" BorderThickness="0" Foreground="#FFFFFF" FontFamily="Segoe UI" FontSize="12" ScrollViewer.VerticalScrollBarVisibility="Auto"/>
        </Border>

        <!-- Action Buttons -->
        <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Button Name="btnUseSelected" Grid.Column="0" Content="Use Selected" Style="{StaticResource DialogButton}" Background="#10B981" Height="32" Margin="0,0,6,0"/>
            <Button Name="btnProceed" Grid.Column="1" Content="Proceed Anyway" Style="{StaticResource DialogButton}" Background="#4B5563" Height="32" Margin="3,0,3,0"/>
            <Button Name="btnCancel" Grid.Column="2" Content="Cancel Build" Style="{StaticResource DialogButton}" Background="#EF4444" Height="32" Margin="6,0,0,0"/>
        </Grid>
    </Grid>
</Window>
'@

        $dialogReader = (New-Object System.Xml.XmlNodeReader ([xml]$dialogXaml))
        $DialogForm = [Windows.Markup.XamlReader]::Load($dialogReader)
        
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
            Update-Log "Build cancelled by user due to driver name mismatch." "#EF4444"
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
    1. Scans for local driver .inf files and registers them in the Windows Driver Store.
    2. Registers the driver name in the Print Spooler.
    3. Creates the TCP/IP printer port machine-wide.
    4. Adds the printer machine-wide so it is visible to all users.
    5. Configures default printer settings (e.g., Duplex, Mono).
.NOTES
    Run under the SYSTEM context (Intune Win32 App Install Behavior: System).
    Version: 1.3
.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File .\Install-Printer.ps1
#>
# Read printers.csv as input
$Printers = Import-Csv '.\printers.csv'
$FirstPrinterName = $Printers[0].Name

# Sanitize the printer name for file system compatibility (replacing invalid characters like '|' with '-')
$SanitizedName = $FirstPrinterName -replace '[\\/:*?"<>|]', '-'

# Create log folder if it doesn't exist
$LogFolder = "$env:SystemDrive\Logs\AddPrinter\$SanitizedName"
if (-not (Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

Start-Transcript -Path "$LogFolder\Install-Printer.log" -Append

# Temporarily stop Program Compatibility Assistant Service (PcaSvc) to suppress driver warnings
$PcaService = Get-Service -Name "PcaSvc" -ErrorAction SilentlyContinue
$PcaServiceStarted = $false
if ($PcaService -and $PcaService.Status -eq 'Running') {
    Write-Host "Temporarily stopping Program Compatibility Assistant Service (PcaSvc)..." -ForegroundColor Yellow
    Stop-Service -Name "PcaSvc" -Force -ErrorAction SilentlyContinue | Out-Null
    $PcaServiceStarted = $true
}

try {
    # 1. Install & register the driver files into the Driver Store using pnputil
    Write-Host "[1/4] Staging drivers to Windows Driver Store..." -ForegroundColor Green
    $allInfs = Get-ChildItem -Path . -Filter "*.inf" -Recurse -ErrorAction SilentlyContinue
    $targetDrivers = $Printers.DriverName | Select-Object -Unique
    $matchedInfs = @()

    foreach ($inf in $allInfs) {
        try {
            $content = Get-Content -Path $inf.FullName -Raw -ErrorAction SilentlyContinue
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
        Write-Host "Could not match Driver Name to a specific INF file. Staging all INF files..." -ForegroundColor Yellow
        $infsToStage = $allInfs | Select-Object -ExpandProperty FullName
    } else {
        Write-Host "Matched Driver Name to specific INF file(s):" -ForegroundColor Green
        $matchedInfs | ForEach-Object { Write-Host " - $_" }
        $infsToStage = $matchedInfs
    }

    $totalInfs = $infsToStage.Count
    $currentInf = 1

    # Handle 32-bit execution context on 64-bit Windows (Intune Management Extension runs as 32-bit)
    $PnpUtilPath = "pnputil.exe"
    if (Test-Path "$env:windir\sysnative\pnputil.exe") {
        $PnpUtilPath = "$env:windir\sysnative\pnputil.exe"
    }

    foreach ($inf in $infsToStage) {
        Write-Host "Staging INF ($currentInf/$totalInfs): $inf" -ForegroundColor Green
        try {
        # Capture standard output and error output from pnputil
        $pnpResult = & $PnpUtilPath /add-driver $inf /install 2>&1
        Write-Host "PnPUtil Result: $pnpResult"
    } catch {
        Write-Error "Failed to run PnPUtil for $($inf): $_"
    }
    $currentInf++
}

# 2. Add the drivers to the Print Spooler
Write-Host "`n[2/4] Registering drivers in Print Spooler..." -ForegroundColor Green
foreach ($driver in $Printers.DriverName | Select-Object -Unique) {
    if (-not (Get-PrinterDriver -Name $driver -ErrorAction SilentlyContinue)) {
        Write-Host "Adding printer driver to spooler: $driver" -ForegroundColor Green
        try {
            Add-PrinterDriver -Name $driver -ErrorAction Stop
            Write-Host "Successfully registered driver in spooler." -ForegroundColor Green
        } catch {
            Write-Error "Failed to add printer driver $($driver) to spooler: $_"
            
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
                # Run pnputil to list staged driver packages and filter for printer class or provider info
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
                
                # Filter to only display Printer class driver store packages
                $printerDrivers | Where-Object { $_.Class -eq "Printers" -or $_.Class -eq "Printer" } | Format-Table -AutoSize | Out-String | Write-Host
            } catch {
                Write-Host "Could not query PnPUtil driver store: $_"
            }
            Write-Host "==============================================================`n"
        }
    } else {
        Write-Host "Driver '$driver' already registered in spooler."
    }
}

# 3. Create TCP/IP Ports and Add Printers
Write-Host "`n[3/4] Creating ports and installing printers..." -ForegroundColor Green
foreach ($printer in $Printers) {
    $PortName = $printer.Name # Matching port name to printer name is standard practice
    
    # Create standard TCP/IP port
    if (-not (Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue)) {
        Write-Host "Creating TCP/IP Port: $PortName ($($printer.PortName))" -ForegroundColor Green
        $PrinterPortOptions = @{
            Name               = $PortName
            PrinterHostAddress = $printer.PortName
            PortNumber         = '9100'
        }
        try {
            Add-PrinterPort @PrinterPortOptions -ErrorAction Stop | Out-Null
            Write-Host "Successfully created printer port: $PortName" -ForegroundColor Green
        } catch {
            Write-Error "Failed to create printer port $($PortName): $_"
        }
    } else {
        Write-Host "Port $PortName already exists."
    }
    
    # Create the printer globally
    if (-not (Get-Printer -Name $printer.Name -ErrorAction SilentlyContinue)) {
        Write-Host "Adding printer: $($printer.Name)" -ForegroundColor Green
        $PrinterAddOptions = @{
            Comment    = $printer.Comment
            DriverName = $printer.DriverName
            Location   = $printer.Location
            Name       = $printer.Name
            PortName   = $PortName
        }
        try {
            Add-Printer @PrinterAddOptions -ErrorAction Stop | Out-Null
            Write-Host "Successfully added printer: $($printer.Name)" -ForegroundColor Green
            
            # 4. Configure Printer Settings (e.g., Color default = 0 (Mono), Duplex = TwoSidedLongEdge)
            Write-Host "Configuring settings for: $($printer.Name)" -ForegroundColor Green
            $PrinterConfigOptions = @{
                Color         = 0
                DuplexingMode = 'TwoSidedLongEdge'
                PrinterName   = $printer.Name
            }
            Set-PrintConfiguration @PrinterConfigOptions -ErrorAction Stop
            Write-Host "Successfully configured print preferences." -ForegroundColor Green
        } catch {
            Write-Error "Failed to add or configure printer $($printer.Name): $_"
        }
    } else {
        Write-Host "Printer $($printer.Name) already installed."
    }
}
} finally {
    # Restore Program Compatibility Assistant Service (PcaSvc) if it was running previously
    if ($PcaServiceStarted) {
        Write-Host "Restoring Program Compatibility Assistant Service (PcaSvc)..." -ForegroundColor Yellow
        Start-Service -Name "PcaSvc" -ErrorAction SilentlyContinue | Out-Null
    }
}

Stop-Transcript
'@

        $uninstallTemplateText = @'
<#
.SYNOPSIS
    Removes printer configurations, standard ports, and drivers globally from the system.
.DESCRIPTION
    Reads printers.csv, removes the associated printers, deletes standard TCP/IP printer ports, and removes the printer driver registration from the driver store.
.NOTES
    This script must run in the SYSTEM (administrator) context (e.g., deployed as a System-level uninstall app in Intune).
    Version: 1.3
.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File .\Uninstall-Printer.ps1
#>
#Read printers.csv as input
$Printers = Import-Csv .\printers.csv
$FirstPrinterName = $Printers[0].Name

# Sanitize the printer name for file system compatibility (replacing invalid characters like '|' with '-')
$SanitizedName = $FirstPrinterName -replace '[\\/:*?"<>|]', '-'

# Create log folder if it doesn't exist
$LogFolder = "$env:SystemDrive\Logs\AddPrinter\$SanitizedName"
if (-not (Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

Start-Transcript -Path "$LogFolder\Uninstall-Printer.log" -Append

#Loop through all printers in the csv-file and remove the printers and their ports
foreach ($printer in $printers) {
    #Remove printers
    if (Get-Printer -Name $printer.Name -ErrorAction SilentlyContinue) {
        Write-Host "Removing printer: $($printer.Name)" -ForegroundColor Green
        try {
            Remove-Printer -Name $printer.Name -Confirm:$false -ErrorAction Stop
            Write-Host "Successfully removed printer: $($printer.Name)" -ForegroundColor Green
        } catch {
            Write-Error "Failed to remove printer $($printer.Name): $_"
        }
    } else {
        Write-Host "Printer $($printer.Name) not found, skipping removal."
    }
    
    Start-Sleep -Seconds 5
    
    #Remove standard TCP/IP port
    $PortName = $printer.Name
    if (Get-PrinterPort -ComputerName $env:COMPUTERNAME | Where-Object Name -EQ $PortName) {
        Write-Host "Removing printer port: $PortName" -ForegroundColor Green
        try {
            Remove-PrinterPort -Name $PortName -ComputerName $env:COMPUTERNAME -Confirm:$false -ErrorAction Stop
            Write-Host "Successfully removed printer port: $PortName" -ForegroundColor Green
        } catch {
            Write-Error "Failed to remove printer port $($PortName): $_"
        }
    } else {
        Write-Host "Printer port $PortName not found, skipping removal."
    }
}

#Remove drivers from the spooler and driver store
foreach ($driver in $printers.drivername | Select-Object -Unique) {
    if (Get-PrinterDriver -Name $driver -ErrorAction SilentlyContinue) {
        Write-Host "Removing driver from Spooler: $driver" -ForegroundColor Green
        $PrinterDriverRemoveOptions = @{
            Confirm               = $false
            Computername          = $env:COMPUTERNAME
            Name                  = $driver
            RemoveFromDriverStore = $true
        }
        try {
            Remove-PrinterDriver @PrinterDriverRemoveOptions -ErrorAction Stop
            Write-Host "Successfully removed driver $driver from store." -ForegroundColor Green
        } catch {
            Write-Warning "Could not completely remove driver $driver from store: $_"
        }
    } else {
        Write-Host "Driver '$driver' not found in spooler, skipping removal."
    }
}

Stop-Transcript
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
    Runs in SYSTEM context as an Intune custom detection script.
    Version: 1.3
#>
$printers = @(
    '__PRINTER_NAME__'
)

$numberofprintersfound = 0
foreach ($printer in $printers) {
    # 1. Try registry key check first (fastest and most reliable under SYSTEM account)
    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\$printer"
    if (Test-Path $regPath) {
        $numberofprintersfound++
        continue
    }

    # 2. Fallback to Get-Printer with a short retry loop to accommodate print spooler delays
    for ($i = 0; $i -lt 3; $i++) {
        try {
            if (Get-Printer -Name $printer -ErrorAction Stop) {
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
    Write-Host "All printers were found"
    exit 0
}
else {
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
            "-o", "`"$printerOutputDir`""
        )
        
        # Create temp files for standard output and error redirection
        $stdoutFile = [System.IO.Path]::GetTempFileName()
        $stderrFile = [System.IO.Path]::GetTempFileName()

        # Start the process redirecting to files
        $process = Start-Process -FilePath $packerExe -ArgumentList ($processArgs -join " ") -NoNewWindow -PassThru -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile

        # Access the handle to force cache population (known .NET process handle caching quirk)
        $null = $process.Handle

        $lastReadOffset = 0
        while (-not $process.HasExited) {
            # Read new lines from stdout file and print them to the GUI log console
            if (Test-Path $stdoutFile) {
                $content = Get-Content -Path $stdoutFile -ErrorAction SilentlyContinue
                if ($content) {
                    $newLines = $content[$lastReadOffset..($content.Count - 1)]
                    foreach ($line in $newLines) {
                        if ($line) { Update-Log $line }
                    }
                    $lastReadOffset = $content.Count
                }
            }

            # Run WPF dispatcher queue to keep GUI responsive and prevent freezing
            [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [Action]{})
            Start-Sleep -Milliseconds 250
        }

        # Read final remaining stdout lines
        if (Test-Path $stdoutFile) {
            $content = Get-Content -Path $stdoutFile -ErrorAction SilentlyContinue
            if ($content -and $content.Count -gt $lastReadOffset) {
                $newLines = $content[$lastReadOffset..($content.Count - 1)]
                foreach ($line in $newLines) {
                    if ($line) { Update-Log $line }
                }
            }
        }

        $stderr = ""
        if (Test-Path $stderrFile) {
            $stderr = Get-Content -Path $stderrFile -Raw -ErrorAction SilentlyContinue
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
        Remove-Item -Path $stdoutFile -Force -ErrorAction SilentlyContinue | Out-Null
        Remove-Item -Path $stderrFile -Force -ErrorAction SilentlyContinue | Out-Null

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
# Generated by Intune Printer Packager v1.3
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

            # Generate copy-pasteable instruction txt file
            $instructionsFile = Join-Path $printerOutputDir "$($sanitized)_Instructions.txt"
            $instructionsContent = @"
======================================================================
INTUNE CONFIGURATION INSTRUCTIONS FOR PRINTER:
$printerName
======================================================================

1. INSTALL PROGRAM SETTINGS:
   -------------------------
   Install Command:
   powershell.exe -ExecutionPolicy Bypass -File .\Install-Printer.ps1

   Uninstall Command:
   powershell.exe -ExecutionPolicy Bypass -File .\Uninstall-Printer.ps1

   Install Behavior: 
   System (Required for printer driver installation)

2. DETECTION RULES:
   ----------------
   Rules format: Use a custom detection script
   Script file: $($sanitized)_Detection.ps1

3. RECOMMENDED INTUNE DESCRIPTION:
   -------------------------------
======================================================================
INTUNE WIN32 APP CONFIGURATION & LOG REFERENCE
======================================================================
Printer Name:   $printerName
Driver Name:    $driverName
IP Address:     $printerIP
----------------------------------------------------------------------
WORKSTATION LOG DIRECTORY:
$env:SystemDrive\Logs\AddPrinter\$sanitized\

INSTALLATION LOG:
$env:SystemDrive\Logs\AddPrinter\$sanitized\Install-Printer.log

UNINSTALLATION LOG:
$env:SystemDrive\Logs\AddPrinter\$sanitized\Uninstall-Printer.log
======================================================================

4. NETWORK REQUIREMENTS:
   ---------------------
    [IMPORTANT] Ensure you create a DHCP IP Reservation for the printer
   IP ($printerIP) on your network so it remains static!

======================================================================
5. macOS INTUNE DEPLOYMENT (SHELL SCRIPT METHOD):
   ----------------------------------------------
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
======================================================================
Generated on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
======================================================================
"@
            $instructionsContent | Set-Content -Path $instructionsFile -Encoding utf8 -Force

            $successMsg = @"
Success! IntuneWin package generated successfully.
File saved to: $targetOutputFile
Instructions saved to: $instructionsFile
Detection Script saved to: $exportedDetectionFile
macOS Script saved to: $macScriptFile

=======================================================
INTUNE WIN32 APP CONFIGURATION INFO (Copy-Pasteable):
=======================================================
Install Command:
powershell.exe -ExecutionPolicy Bypass -File .\Install-Printer.ps1

Uninstall Command:
powershell.exe -ExecutionPolicy Bypass -File .\Uninstall-Printer.ps1

Detection Rule:
Select: 'Use a custom detection script'
Upload the script file: $($sanitized)_Detection.ps1

Recommended Description:
======================================================================
INTUNE WIN32 APP CONFIGURATION & LOG REFERENCE
======================================================================
Printer Name:   $printerName
Driver Name:    $driverName
IP Address:     $printerIP
----------------------------------------------------------------------
WORKSTATION LOG DIRECTORY:
$env:SystemDrive\Logs\AddPrinter\$sanitized\

INSTALLATION LOG:
$env:SystemDrive\Logs\AddPrinter\$sanitized\Install-Printer.log

UNINSTALLATION LOG:
$env:SystemDrive\Logs\AddPrinter\$sanitized\Uninstall-Printer.log
======================================================================

[DHCP RESERVATION REMINDER]
Make sure to reserve the IP address ($printerIP) in DHCP so the printer maintains a static IP!
===============================================================

=======================================================
macOS DEPLOYMENT INFO (Copy-Pasteable):
=======================================================
Step 1: Deploy Manufacturer Driver (PKG or DMG)
- Extracted PKG driver installer must be uploaded to Intune.
- Intune section: Apps > macOS apps > Add
- App type: macOS app (PKG)
- Assign as Required to macOS devices.

Step 2: Deploy Installation Script
- Script File: $($sanitized)_macOS_Install.sh
- Intune section: Devices > macOS > Shell scripts > Add
- Settings:
  * Run as signed-in user: No (runs as root)
  * Hide notifications: Yes
  * Script frequency: Run once
  * Max retries: 3
- Assign as Required to the same macOS devices.
=======================================================
"@
            Update-Log $successMsg "#10B981"
            [System.Windows.MessageBox]::Show("Successfully generated IntuneWin package, instructions, detection, and macOS script!`n`nFile: $finalFileName`nLocation: $printerOutputDir", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        } else {
            throw "The expected packaging file 'Install-Printer.intunewin' was not found in the output directory."
        }

    } catch {
        Update-Log "An error occurred during build:`n$_" "#EF4444"
        [System.Windows.MessageBox]::Show("An error occurred during the packaging process:`n`n$_", "Build Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
    } finally {
        # Cleanup temporary workspace
        if (Test-Path $tempFolder) {
            Update-Log "Cleaning up temporary staging directory..."
            Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
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
