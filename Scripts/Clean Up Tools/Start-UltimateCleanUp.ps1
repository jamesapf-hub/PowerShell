<#
.SYNOPSIS
    Main runner script for the Ultimate PC CleanUp Utility.
.DESCRIPTION
    This script can be run in two modes:
    1. Interactive GUI mode (default): Launches a modern dark-themed WPF dashboard.
    2. Command-Line mode (via switches): Performs cleanup directly in the terminal,
       making it suitable for RMM, scheduled tasks, or shell automation.
.PARAMETER NoGui
    Launches the tool in command-line mode without displaying the GUI window.
.PARAMETER RunAll
    Automatically runs all discovered tasks in command-line mode.
.PARAMETER RunTasks
    An array of task basenames/filenames to execute in command-line mode.
.PARAMETER ListTasks
    Lists all available tasks with descriptions in the terminal and exits.
.NOTES
    Author  : Antigravity Pair Programmer
    Version : 1.1
#>

[CmdletBinding()]
param (
    [switch]$NoGui,
    [switch]$RunAll,
    [string[]]$RunTasks,
    [switch]$ListTasks,
    [switch]$Gui
)

# Set base parameters
$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
if (-not $PSScriptRoot) { $PSScriptRoot = Get-Location }
$ScriptsFolder = Join-Path -Path $PSScriptRoot -ChildPath "Scripts"

# Ensure Scripts folder exists
if (-not (Test-Path $ScriptsFolder)) {
    New-Item -ItemType Directory -Path $ScriptsFolder | Out-Null
}

# Default behavior: If no parameters are specified, default to GUI mode
$noParamsSpecified = -not ($PSBoundParameters.ContainsKey('NoGui') -or 
                           $PSBoundParameters.ContainsKey('RunAll') -or 
                           $PSBoundParameters.ContainsKey('RunTasks') -or 
                           $PSBoundParameters.ContainsKey('ListTasks') -or 
                           $PSBoundParameters.ContainsKey('Gui'))

if ($noParamsSpecified) {
    $Gui = $true
    $NoGui = $false
}

# If any command-line execution parameter is passed, default to NoGui
if ($RunAll -or $RunTasks -or $ListTasks) {
    $NoGui = $true
}

# If GUI is explicitly requested, override NoGui
if ($Gui) {
    $NoGui = $false
}

# Check if running as administrator (required for system-level cleanup tasks)
$currentUser = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
$isAdmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    # Request UAC Elevation
    if ($PSCommandPath) {
        $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    } else {
        $arguments = "-NoProfile -ExecutionPolicy Bypass -Command `"[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; iex (irm 'https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Clean%20Up%20Tools/Start-UltimateCleanUp.ps1')`""
    }
    if ($NoGui) { $arguments += " -NoGui" }
    if ($RunAll) { $arguments += " -RunAll" }
    if ($ListTasks) { $arguments += " -ListTasks" }
    if ($Gui) { $arguments += " -Gui" }
    if ($RunTasks) {
        $arguments += " -RunTasks $($RunTasks -join ',')"
    }
    
    try {
        $currentDir = (Get-Location).Path
        if ($NoGui) {
            # In NoGui/CLI mode, we want to run in the current console and wait for exit
            Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -Verb RunAs -WorkingDirectory $currentDir -Wait
        } else {
            Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -Verb RunAs -WorkingDirectory $currentDir
        }
    } catch {
        Write-Error "This utility requires Administrator privileges."
    }
    exit
}

# Enforce STA ApartmentState for WPF UI rendering (only if GUI mode is active)
if (-not $NoGui -and [System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Verbose "Thread is not STA. Relaunching GUI in a new STA process..."
    if ($PSCommandPath) {
        $arguments = "-NoProfile -ExecutionPolicy Bypass -Sta -File `"$PSCommandPath`""
    } else {
        $arguments = "-NoProfile -ExecutionPolicy Bypass -Sta -Command `"[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; iex (irm 'https://raw.githubusercontent.com/jamesapf-hub/PowerShell/main/Scripts/Clean%20Up%20Tools/Start-UltimateCleanUp.ps1')`""
    }
    if ($Gui) { $arguments += " -Gui" }
    
    $currentDir = (Get-Location).Path
    Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -WorkingDirectory $currentDir -Wait
    return
}

# Reusable function to discover scripts and extract synopsis metadata
function Get-DiscoveredScripts {
    if (-not (Test-Path $ScriptsFolder)) {
        return @()
    }
    
    $files = Get-ChildItem -Path $ScriptsFolder -Filter *.ps1 -Recurse | Sort-Object Name
    $list = [System.Collections.Generic.List[PSCustomObject]]::new()
    
    foreach ($file in $files) {
        # Parse synopsis block if present
        $content = Get-Content $file.FullName -Head 20
        $synopsis = ""
        $inSynopsis = $false
        foreach ($line in $content) {
            if ($line -match '\.SYNOPSIS') {
                $inSynopsis = $true
                continue
            }
            if ($inSynopsis) {
                if ($line -match '^\s*\.\w+' -or $line -match '#>' -or $line -match '^\s*$') {
                    $inSynopsis = $false
                } else {
                    $synopsis += ($line.Trim('#').Trim() + " ")
                }
            }
        }
        
        $synopsis = $synopsis.Trim()
        if (-not $synopsis) {
            $synopsis = "No script synopsis details provided."
        }
        
        # Format friendly title (e.g. "01_ClearTempFiles.ps1" -> "Clear Temp Files")
        $title = $file.BaseName -replace '^\d+[-_]', '' -replace '[-_]', ' '
        $title = $title -creplace '([a-z])([A-Z])', '$1 $2'
        $title = (Get-Culture).TextInfo.ToTitleCase($title)
        
        $list.Add([PSCustomObject]@{
            Path        = $file.FullName
            FileName    = $file.Name
            BaseName    = $file.BaseName
            Title       = $title
            Description = $synopsis
        }) | Out-Null
    }
    return $list
}

# ==========================================
# COMMAND LINE / AUTOMATION MODE
# ==========================================
if ($NoGui) {
    $tasks = Get-DiscoveredScripts
    if ($tasks.Count -eq 0) {
        Write-Host "[-] No cleanup tasks found in '$ScriptsFolder'." -ForegroundColor Red
        exit 0
    }
    
    if ($ListTasks) {
        Write-Host "`r`n=== DISCOVERED CLEANUP TASKS ===" -ForegroundColor Cyan
        foreach ($task in $tasks) {
            Write-Host "- [$($task.BaseName)] $($task.Title)" -ForegroundColor Yellow
            Write-Host "  Description: $($task.Description)`r`n" -ForegroundColor DarkGray
        }
        exit 0
    }
    
    # Filter tasks to run
    $tasksToRun = @()
    if ($RunTasks) {
        foreach ($target in $RunTasks) {
            $match = $tasks | Where-Object { $_.BaseName -eq $target -or $_.Title -eq $target -or $_.FileName -eq $target }
            if ($match) {
                $tasksToRun += $match
            } else {
                Write-Host "[-] Warning: Task '$target' not found." -ForegroundColor Yellow
            }
        }
    } else {
        # Default to running all if RunAll or no parameter specified
        $tasksToRun = $tasks
    }
    
    if ($tasksToRun.Count -eq 0) {
        Write-Host "[-] No valid tasks to execute." -ForegroundColor Red
        exit 1
    }
    
    Write-Host "`r`n==================================================" -ForegroundColor Cyan
    Write-Host "STARTING COMMAND LINE PC CLEANUP" -ForegroundColor Cyan
    Write-Host "Executing $($tasksToRun.Count) tasks..." -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    
    $successCount = 0
    foreach ($task in $tasksToRun) {
        Write-Host "`r`n>>> Running task: $($task.Title)..." -ForegroundColor Yellow
        
        # Start background powershell instance to run the task
        $proc = Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$($task.Path)`" -Force" -NoNewWindow -Wait -PassThru
        
        if ($proc.ExitCode -eq 0) {
            Write-Host "[+] Task '$($task.Title)' completed successfully." -ForegroundColor Green
            $successCount++
        } else {
            Write-Host "[-] Task '$($task.Title)' failed with exit code $($proc.ExitCode)." -ForegroundColor Red
        }
    }
    
    Write-Host "`r`n==================================================" -ForegroundColor Cyan
    Write-Host "CLEANUP FINISHED: $successCount of $($tasksToRun.Count) tasks completed successfully." -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    exit 0
}

# ==========================================
# INTERACTIVE GUI MODE (WPF)
# ==========================================

# Import GUI assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# Define XAML for the Main Window
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Ultimate PC CleanUp Utility"
        Height="700" Width="1050"
        Background="#0e0e0e"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize">
    
    <Window.Resources>
        <!-- Primary Action Button Style (Seriun Orange Accent) -->
        <Style TargetType="Button" x:Key="PrimaryBtn">
            <Setter Property="Background" Value="#ff6900"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="13"/>
            <Style.Resources>
                <Style TargetType="Border">
                    <Setter Property="CornerRadius" Value="6"/>
                </Style>
            </Style.Resources>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#cc5400"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#662a00"/>
                    <Setter Property="Foreground" Value="#9ca3af"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Secondary Action Button Style (Outline Seriun Orange Accent) -->
        <Style TargetType="Button" x:Key="AccentBtn">
            <Setter Property="Background" Value="#131313"/>
            <Setter Property="Foreground" Value="#ff6900"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#ff6900"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="13"/>
            <Style.Resources>
                <Style TargetType="Border">
                    <Setter Property="CornerRadius" Value="6"/>
                </Style>
            </Style.Resources>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#1b1b1b"/>
                    <Setter Property="BorderBrush" Value="#ff8533"/>
                    <Setter Property="Foreground" Value="#ff8533"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#0e0e0e"/>
                    <Setter Property="BorderBrush" Value="#662a00"/>
                    <Setter Property="Foreground" Value="#662a00"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- General Button Style (Seriun Secondary Elements) -->
        <Style TargetType="Button" x:Key="DefaultBtn">
            <Setter Property="Background" Value="#1b1b1b"/>
            <Setter Property="Foreground" Value="#e4e4e7"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#393939"/>
            <Setter Property="FontWeight" Value="Medium"/>
            <Setter Property="FontSize" Value="11"/>
            <Style.Resources>
                <Style TargetType="Border">
                    <Setter Property="CornerRadius" Value="4"/>
                </Style>
            </Style.Resources>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#393939"/>
                    <Setter Property="BorderBrush" Value="#ff6900"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#0e0e0e"/>
                    <Setter Property="BorderBrush" Value="#1b1b1b"/>
                    <Setter Property="Foreground" Value="#71717a"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Script Card Item Style (Seriun Surface System) -->
        <Style TargetType="Border" x:Key="ScriptCardStyle">
            <Setter Property="Background" Value="#131313"/>
            <Setter Property="BorderBrush" Value="#1b1b1b"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="6"/>
            <Setter Property="Padding" Value="12"/>
            <Setter Property="Margin" Value="0,0,0,10"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#1b1b1b"/>
                    <Setter Property="BorderBrush" Value="#393939"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    
    <Grid Margin="25">
        <Grid.RowDefinitions>
            <!-- Title Bar/Header -->
            <RowDefinition Height="Auto"/>
            <!-- Main Content Area -->
            <RowDefinition Height="*"/>
            <!-- Status & Progress Footer -->
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header Panel -->
        <Border Grid.Row="0" BorderThickness="0,0,0,1" BorderBrush="#1b1b1b" Padding="0,0,0,18" Margin="0,0,0,20">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                    <TextBlock Text="Ultimate PC CleanUp" FontSize="28" FontWeight="Bold" Foreground="#F4F4F5"/>
                    <TextBlock Text="Run custom PC tidy-up scripts safely, monitor outputs, and keep Windows clean." FontSize="13" Foreground="#A1A1AA" Margin="0,5,0,0"/>
                </StackPanel>
                <Button Name="BtnRefresh" Grid.Column="1" Style="{StaticResource DefaultBtn}" Content="Scan Folder" Width="100" Height="32" Cursor="Hand" VerticalAlignment="Center"/>
            </Grid>
        </Border>
        
        <!-- Main Layout Split -->
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <!-- Task Checklist Selection (Left) -->
                <ColumnDefinition Width="4*"/>
                <!-- Vertical Divider -->
                <ColumnDefinition Width="20"/>
                <!-- Console Logging Console (Right) -->
                <ColumnDefinition Width="5*"/>
            </Grid.ColumnDefinitions>
            
            <!-- Left Panel: Available Scripts -->
            <Grid Grid.Column="0">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <TextBlock Text="DISCOVERED CLEANUP TASKS" FontSize="11" FontWeight="Bold" Foreground="#71717A" Margin="0,0,0,10"/>
                
                <!-- Scrollable Checklist -->
                <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                    <StackPanel Name="ScriptListContainer" Margin="0,0,5,0"/>
                </ScrollViewer>
                
                <!-- Bulk Selection controls -->
                <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,15,0,0">
                    <Button Name="BtnSelectAll" Style="{StaticResource DefaultBtn}" Content="Select All" Margin="0,0,8,0" Padding="14,6" Cursor="Hand"/>
                    <Button Name="BtnDeselectAll" Style="{StaticResource DefaultBtn}" Content="Clear Selection" Padding="14,6" Cursor="Hand"/>
                </StackPanel>
            </Grid>
            
            <!-- Right Panel: Logs & Execution Controls -->
            <Grid Grid.Column="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                
                <!-- Execution Buttons -->
                <Grid Grid.Row="0" Margin="0,0,0,12">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Button Name="BtnRunSelected" Style="{StaticResource AccentBtn}" Content="Run Selected Tasks" Margin="0,0,10,0" Height="42" Cursor="Hand"/>
                    <Button Name="BtnRunAll" Grid.Column="1" Style="{StaticResource PrimaryBtn}" Content="Run All Tasks" Height="42" Cursor="Hand"/>
                </Grid>
                
                <!-- Monospaced Log Console -->
                <Border Grid.Row="1" Background="#0e0e0e" CornerRadius="6" BorderThickness="1" BorderBrush="#1b1b1b">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        
                        <!-- Console Header -->
                        <Border Background="#131313" Padding="12,8" BorderThickness="0,0,0,1" BorderBrush="#1b1b1b">
                            <Grid>
                                <TextBlock Text="REAL-TIME EXECUTION LOG" FontSize="11" FontWeight="Bold" Foreground="#A1A1AA" VerticalAlignment="Center"/>
                                <Button Name="BtnClearLog" Content="Clear Console" Style="{StaticResource DefaultBtn}" HorizontalAlignment="Right" Padding="8,3" FontSize="10" Cursor="Hand"/>
                            </Grid>
                        </Border>
                        
                        <!-- Console Textbox -->
                        <TextBox Name="TxtConsole" Grid.Row="1" Background="Transparent" Foreground="#F4F4F5" FontFamily="Consolas" FontSize="12.5" BorderThickness="0" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" IsReadOnly="True" AcceptsReturn="True" TextWrapping="Wrap" Margin="12" Padding="2"/>
                    </Grid>
                </Border>
            </Grid>
        </Grid>
        
        <!-- Footer Panel -->
        <Border Grid.Row="2" BorderThickness="0,1,0,0" BorderBrush="#1b1b1b" Padding="0,18,0,0" Margin="0,20,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Name="TxtProgress" Text="Ready to begin cleanup" Foreground="#A1A1AA" VerticalAlignment="Center" FontSize="12"/>
                <ProgressBar Name="ProgressIndicator" Grid.Column="1" Margin="20,0" Height="8" Background="#1b1b1b" Foreground="#ff6900" Minimum="0" Maximum="100" Value="0" BorderThickness="0">
                    <ProgressBar.Template>
                        <ControlTemplate TargetType="ProgressBar">
                            <Border Background="{TemplateBinding Background}" CornerRadius="4" ClipToBounds="True">
                                <Grid>
                                    <Border Name="PART_Track" />
                                    <Border Name="PART_Indicator" Background="{TemplateBinding Foreground}" HorizontalAlignment="Left" CornerRadius="4"/>
                                </Grid>
                            </Border>
                        </ControlTemplate>
                    </ProgressBar.Template>
                </ProgressBar>
                <TextBlock Name="TxtStatusCount" Grid.Column="2" Text="0 / 0 Completed" Foreground="#A1A1AA" VerticalAlignment="Center" FontSize="12" FontWeight="SemiBold"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Load the XML XAML into WPF Window object
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Map XAML named controls to PowerShell script-scoped variables automatically
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    Set-Variable -Name $_.Name -Value $window.FindName($_.Name) -Scope Script
}

# Synchronized hashtable to share state with background execution thread safely
$syncHash = [hashtable]::Synchronized(@{})
$syncHash.Window = $window
$syncHash.TxtConsole = $TxtConsole
$syncHash.ProgressIndicator = $ProgressIndicator
$syncHash.TxtStatusCount = $TxtStatusCount
$syncHash.BtnRunAll = $BtnRunAll
$syncHash.BtnRunSelected = $BtnRunSelected
$syncHash.BtnRefresh = $BtnRefresh
$syncHash.TxtProgress = $TxtProgress

# Initialize a list to hold discovered script objects
$scriptList = [System.Collections.Generic.List[PSCustomObject]]::new()

# Create a background runspace pool / runspace to execute scripts asynchronously
$runspace = [RunspaceFactory]::CreateRunspace()
$runspace.ApartmentState = "STA"
$runspace.ThreadOptions = "ReuseThread"
$runspace.Open()
$runspace.SessionStateProxy.SetVariable("syncHash", $syncHash)

# Function to parse scripts inside the 'Scripts' folder
function Load-Scripts {
    $ScriptListContainer.Children.Clear()
    $scriptList.Clear()
    
    $discovered = Get-DiscoveredScripts
    foreach ($file in $discovered) {
        # Load XAML for the dynamic Card UI element
        $cardXaml = @"
<Border xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Style="{StaticResource ScriptCardStyle}">
    <Grid>
        <Grid.ColumnDefinitions>
            <Definition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        
        <CheckBox Name="ChkSelect" VerticalAlignment="Center" Margin="0,0,12,0" IsChecked="True" Cursor="Hand"/>
        
        <StackPanel Grid.Column="1" VerticalAlignment="Center">
            <TextBlock Name="TxtTitle" Text="$($file.Title)" FontWeight="SemiBold" FontSize="14" Foreground="#F4F4F5"/>
            <TextBlock Name="TxtDesc" Text="$($file.Description)" FontSize="11.5" Foreground="#A1A1AA" TextWrapping="Wrap" Margin="0,4,0,0"/>
        </StackPanel>
        
        <Border Grid.Column="2" Name="StatusBadge" CornerRadius="4" Padding="8,4" VerticalAlignment="Center" Background="#1b1b1b" MinWidth="75">
            <TextBlock Name="TxtStatus" Text="Pending" Foreground="#A1A1AA" FontWeight="Bold" FontSize="10.5" HorizontalAlignment="Center"/>
        </Border>
    </Grid>
</Border>
"@
        $card = [Windows.Markup.XamlReader]::Parse($cardXaml)
        $ScriptListContainer.Children.Add($card)
        
        # Extract dynamic card elements
        $scriptObj = [PSCustomObject]@{
            Path        = $file.Path
            FileName    = $file.FileName
            Title       = $file.Title
            Description = $file.Description
            Card        = $card
            Checkbox    = $card.FindName("ChkSelect")
            StatusText  = $card.FindName("TxtStatus")
            StatusBadge = $card.FindName("StatusBadge")
        }
        $scriptList.Add($scriptObj) | Out-Null
    }
    
    # Update Footer statuses
    $TxtStatusCount.Text = "0 / $($scriptList.Count) Completed"
    $ProgressIndicator.Value = 0
}

# Function to kick off execution (Runs on background Runspace thread)
function Start-Execution {
    param([bool]$onlySelected)
    
    # Identify scripts to execute
    $selectedScripts = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($script in $scriptList) {
        # Reset card status visually
        $script.StatusText.Text = "Pending"
        $script.StatusBadge.Background = [System.Windows.Media.BrushConverter]::ConvertFromString("#1b1b1b")
        $script.StatusText.Foreground = [System.Windows.Media.BrushConverter]::ConvertFromString("#A1A1AA")
        
        if (-not $onlySelected -or $script.Checkbox.IsChecked) {
            $selectedScripts.Add($script) | Out-Null
        }
    }
    
    if ($selectedScripts.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Please select at least one task to run.", "No Tasks Selected", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    # Lock controls during execution
    $BtnRunAll.IsEnabled = $false
    $BtnRunSelected.IsEnabled = $false
    $BtnRefresh.IsEnabled = $false
    
    $TxtProgress.Text = "Executing cleanup..."
    $ProgressIndicator.Value = 0
    
    # Share state with background thread
    $syncHash.Scripts = $selectedScripts
    $syncHash.TotalCount = $selectedScripts.Count
    $syncHash.CompletedCount = 0
    
    # Reset console log
    $TxtConsole.Clear()
    $TxtConsole.AppendText("[*] Starting PC cleanup utility process...`r`n`r`n")
    
    # Setup PowerShell script execution engine
    $ps = [PowerShell]::Create()
    $ps.Runspace = $runspace
    
    $runscript = {
        $scripts = $syncHash.Scripts
        
        for ($i = 0; $i -lt $scripts.Count; $i++) {
            $script = $scripts[$i]
            
            # Transition task card UI status to 'Running'
            $syncHash.Window.Dispatcher.Invoke([Action]{
                $script.StatusText.Text = "Running"
                $script.StatusBadge.Background = [System.Windows.Media.BrushConverter]::ConvertFromString("#393939") # Surface Bright
                $script.StatusText.Foreground = [System.Windows.Media.BrushConverter]::ConvertFromString("#ff6900") # Primary Orange Accent
                
                $syncHash.TxtConsole.AppendText("======================================================================`r`n")
                $syncHash.TxtConsole.AppendText("TASK: $($script.Title)`r`n")
                $syncHash.TxtConsole.AppendText("======================================================================`r`n")
                $syncHash.TxtConsole.ScrollToEnd()
            })
            
            # Prepare process start properties (run sub-script inside hidden process)
            $pinfo = New-Object System.Diagnostics.ProcessStartInfo
            $pinfo.FileName = "powershell.exe"
            $pinfo.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$($script.Path)`" -Force"
            $pinfo.RedirectStandardOutput = $true
            $pinfo.RedirectStandardError = $true
            $pinfo.UseShellExecute = $false
            $pinfo.CreateNoWindow = $true
            
            $proc = New-Object System.Diagnostics.Process
            $proc.StartInfo = $pinfo
            
            $proc.Start() | Out-Null
            
            # Read stdout stream asynchronously line-by-line
            while (-not $proc.HasExited) {
                $line = $proc.StandardOutput.ReadLine()
                if ($line -ne $null) {
                    $syncHash.Window.Dispatcher.Invoke([Action]{
                        $syncHash.TxtConsole.AppendText("$line`r`n")
                        $syncHash.TxtConsole.ScrollToEnd()
                    })
                }
                Start-Sleep -Milliseconds 10
            }
            
            # Capture lingering standard outputs
            $remaining = $proc.StandardOutput.ReadToEnd()
            if (-not [string]::IsNullOrEmpty($remaining)) {
                $syncHash.Window.Dispatcher.Invoke([Action]{
                    $syncHash.TxtConsole.AppendText("$remaining")
                    $syncHash.TxtConsole.ScrollToEnd()
                })
            }
            
            # Capture errors
            $errors = $proc.StandardError.ReadToEnd()
            $exitCode = $proc.ExitCode
            
            # Update Task card UI state based on exit code
            $syncHash.Window.Dispatcher.Invoke([Action]{
                $syncHash.CompletedCount++
                $percent = [math]::Round(($syncHash.CompletedCount / $syncHash.TotalCount) * 100)
                $syncHash.ProgressIndicator.Value = $percent
                $syncHash.TxtStatusCount.Text = "$($syncHash.CompletedCount) / $($syncHash.TotalCount) Completed"
                
                if ($exitCode -eq 0) {
                    $script.StatusText.Text = "Success"
                    $script.StatusBadge.Background = [System.Windows.Media.BrushConverter]::ConvertFromString("#064E3B") # Dark Green
                    $script.StatusText.Foreground = [System.Windows.Media.BrushConverter]::ConvertFromString("#34D399") # Bright Green
                    $syncHash.TxtConsole.AppendText("`r`n[+] Task completed successfully.`r`n`r`n")
                } else {
                    $script.StatusText.Text = "Failed"
                    $script.StatusBadge.Background = [System.Windows.Media.BrushConverter]::ConvertFromString("#7F1D1D") # Dark Red
                    $script.StatusText.Foreground = [System.Windows.Media.BrushConverter]::ConvertFromString("#F87171") # Bright Red
                    if (-not [string]::IsNullOrEmpty($errors)) {
                        $syncHash.TxtConsole.AppendText("`r`n[!] Errors reported:`r`n$errors`r`n")
                     }
                    $syncHash.TxtConsole.AppendText("`r`n[-] Task failed with exit code $exitCode.`r`n`r`n")
                }
                $syncHash.TxtConsole.ScrollToEnd()
            })
        }
        
        # Enable controls after tasks finish
        $syncHash.Window.Dispatcher.Invoke([Action]{
            $syncHash.BtnRunAll.IsEnabled = $true
            $syncHash.BtnRunSelected.IsEnabled = $true
            $syncHash.BtnRefresh.IsEnabled = $true
            $syncHash.TxtProgress.Text = "All tasks processed."
            $syncHash.TxtConsole.AppendText("======================================================================`r`n")
            $syncHash.TxtConsole.AppendText("CLEANUP PROCESS FINISHED`r`n")
            $syncHash.TxtConsole.AppendText("======================================================================`r`n")
            $syncHash.TxtConsole.ScrollToEnd()
        })
    }
    
    $ps.AddScript($runscript).BeginInvoke() | Out-Null
}

# --- Event Handlers ---

# Refresh/Scan Folder Button
$BtnRefresh.Add_Click({
    Load-Scripts
})

# Run Selected Tasks Button
$BtnRunSelected.Add_Click({
    Start-Execution -onlySelected $true
})

# Run All Tasks Button
$BtnRunAll.Add_Click({
    Start-Execution -onlySelected $false
})

# Clear Log Button
$BtnClearLog.Add_Click({
    $TxtConsole.Clear()
})

# Select All Tasks
$BtnSelectAll.Add_Click({
    foreach ($script in $scriptList) {
        $script.Checkbox.IsChecked = $true
    }
})

# Clear/Deselect All Tasks
$BtnDeselectAll.Add_Click({
    foreach ($script in $scriptList) {
        $script.Checkbox.IsChecked = $false
    }
})

# Window closing event - clean up runspaces
$window.Add_Closing({
    try {
        if ($runspace) {
            $runspace.Close()
            $runspace.Dispose()
        }
    } catch {}
})

# Initial script scan and load
Load-Scripts

# Display the main window
$window.ShowDialog() | Out-Null
