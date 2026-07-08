# Requires -Version 5.1
<#
.SYNOPSIS
    Exchange Online Mobile Device Quarantine Release GUI Tool
.DESCRIPTION
    A desktop GUI tool to connect to Exchange Online, search and select licensed mailbox users,
    check their mobile device quarantine status, and release devices to the Allowed list.
#>

# Load WPF assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# Define XAML layout
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Exchange Online - Quarantine Release Manager" Height="720" Width="1000"
        Background="#0F172A" WindowStartupLocation="CenterScreen" FontSize="13" Foreground="#F8FAFC">
    <Window.Resources>
        <!-- Header Typography -->
        <Style x:Key="HeaderStyle" TargetType="TextBlock">
            <Setter Property="FontSize" Value="20"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Foreground" Value="#F8FAFC"/>
        </Style>
        <Style x:Key="SubHeaderStyle" TargetType="TextBlock">
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="Normal"/>
            <Setter Property="Foreground" Value="#94A3B8"/>
        </Style>
        
        <!-- Primary Button Style (Indigo) -->
        <Style TargetType="Button">
            <Setter Property="Background" Value="#3B82F6"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Border" CornerRadius="6" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#2563EB"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#1D4ED8"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Border" Property="Background" Value="#1E293B"/>
                                <Setter Property="Foreground" Value="#475569"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- Secondary Button Style -->
        <Style x:Key="SecondaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#1E293B"/>
            <Setter Property="BorderBrush" Value="#475569"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Border" CornerRadius="6" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#334155"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Border" Property="Background" Value="#0F172A"/>
                                <Setter TargetName="Border" Property="BorderBrush" Value="#1E293B"/>
                                <Setter Property="Foreground" Value="#475569"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- Action Button Style (Emerald Green) -->
        <Style x:Key="ActionButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#10B981"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Border" CornerRadius="6" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#059669"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#047857"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Border" Property="Background" Value="#1E293B"/>
                                <Setter Property="Foreground" Value="#475569"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- Input Textbox Style -->
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#1E293B"/>
            <Setter Property="Foreground" Value="#F8FAFC"/>
            <Setter Property="BorderBrush" Value="#334155"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="CaretBrush" Value="#F8FAFC"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="Border" CornerRadius="6" BorderThickness="{TemplateBinding BorderThickness}" BorderBrush="{TemplateBinding BorderBrush}" Background="{TemplateBinding Background}">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="0"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter TargetName="Border" Property="BorderBrush" Value="#3B82F6"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- Directory ListBox Style -->
        <Style TargetType="ListBox">
            <Setter Property="Background" Value="#0F172A"/>
            <Setter Property="Foreground" Value="#F8FAFC"/>
            <Setter Property="BorderBrush" Value="#334155"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="4"/>
            <Setter Property="ItemContainerStyle">
                <Setter.Value>
                    <Style TargetType="ListBoxItem">
                        <Setter Property="Padding" Value="12,10"/>
                        <Setter Property="Margin" Value="0,2"/>
                        <Setter Property="Background" Value="Transparent"/>
                        <Setter Property="BorderThickness" Value="0"/>
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="ListBoxItem">
                                    <Border x:Name="Border" CornerRadius="4" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                                        <ContentPresenter />
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True">
                                            <Setter TargetName="Border" Property="Background" Value="#1E293B"/>
                                        </Trigger>
                                        <Trigger Property="IsSelected" Value="True">
                                            <Setter TargetName="Border" Property="Background" Value="#3B82F6"/>
                                            <Setter Property="Foreground" Value="White"/>
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
            <RowDefinition Height="Auto"/> <!-- Header and Connection Panel -->
            <RowDefinition Height="*"/>    <!-- Main Body Columns -->
            <RowDefinition Height="160"/>  <!-- Live Console Panel -->
        </Grid.RowDefinitions>

        <!-- TOP BAR: HEADER & CONNECTION CONTROL -->
        <Border Grid.Row="0" Background="#1E293B" BorderBrush="#334155" BorderThickness="0,0,0,1" Padding="20,15">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel VerticalAlignment="Center">
                    <TextBlock Text="M365 Quarantine Release Assistant" Style="{StaticResource HeaderStyle}"/>
                    <TextBlock Text="Search Exchange Online mailboxes and instantly release quarantined ActiveSync devices." Style="{StaticResource SubHeaderStyle}" Margin="0,3,0,0"/>
                </StackPanel>

                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <!-- Status Indicator Dot -->
                    <Border x:Name="StatusDot" Width="10" Height="10" CornerRadius="5" Background="#EF4444" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <TextBlock x:Name="StatusText" Text="Disconnected" Foreground="#94A3B8" VerticalAlignment="Center" Margin="0,0,15,0" FontWeight="SemiBold"/>
                    <Button x:Name="BtnConnect" Content="Connect to Exchange Online" Width="220" Height="38"/>
                    <Button x:Name="BtnDisconnect" Content="Disconnect" Width="100" Height="38" Margin="10,0,0,0" Style="{StaticResource SecondaryButton}" Visibility="Collapsed"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- MAIN LAYOUT -->
        <Grid Grid.Row="1" Margin="20">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="340"/> <!-- Directory Pane -->
                <ColumnDefinition Width="20"/>  <!-- Spacer -->
                <ColumnDefinition Width="*"/>   <!-- Device Details & Execution Pane -->
            </Grid.ColumnDefinitions>

            <!-- LEFT COLUMN: USER DIRECTORY -->
            <Grid Grid.Column="0">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/> <!-- Search bar -->
                    <RowDefinition Height="*"/>    <!-- Users ListBox -->
                    <RowDefinition Height="Auto"/> <!-- Directory load action -->
                </Grid.RowDefinitions>

                <!-- Local Filtering Search Box -->
                <Grid Grid.Row="0" Margin="0,0,0,12">
                    <TextBox x:Name="TxtSearch" Height="38" VerticalContentAlignment="Center" IsEnabled="False"/>
                    <TextBlock IsHitTestVisible="False" Text="Filter users..." VerticalAlignment="Center" HorizontalAlignment="Left" Margin="12,0,0,0" Foreground="#64748B">
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
                <ListBox Grid.Row="1" x:Name="LstUsers" Margin="0,0,0,12" IsEnabled="False">
                    <ListBox.ItemTemplate>
                        <DataTemplate>
                            <StackPanel>
                                <TextBlock Text="{Binding DisplayName}" FontWeight="Bold" Foreground="#F8FAFC" FontSize="13"/>
                                <TextBlock Text="{Binding PrimarySmtpAddress}" Foreground="#64748B" FontSize="11" Margin="0,2,0,0"/>
                            </StackPanel>
                        </DataTemplate>
                    </ListBox.ItemTemplate>
                </ListBox>

                <!-- Fetch Directory Action -->
                <Button Grid.Row="2" x:Name="BtnLoadUsers" Content="Load User Directory" Height="38" IsEnabled="False"/>
            </Grid>

            <!-- RIGHT COLUMN: SELECTION & ACTION PANEL -->
            <Border Grid.Column="2" Background="#1E293B" BorderBrush="#334155" BorderThickness="1" CornerRadius="8" Padding="20">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/> <!-- Target Mailbox Info -->
                        <RowDefinition Height="*"/>    <!-- Quarantined Devices Grid -->
                        <RowDefinition Height="Auto"/> <!-- Actions Panel -->
                    </Grid.RowDefinitions>

                    <!-- Selected Target Header -->
                    <StackPanel Grid.Row="0" Margin="0,0,0,20">
                        <TextBlock x:Name="TxtSelectedUserName" Text="No Mailbox Selected" FontSize="18" FontWeight="Bold" Foreground="#F8FAFC"/>
                        <TextBlock x:Name="TxtSelectedUserEmail" Text="Select a mailbox from the left panel to display quarantine status." FontSize="13" Foreground="#94A3B8" Margin="0,3,0,0"/>
                    </StackPanel>

                    <!-- Quarantined Devices Container -->
                    <Grid Grid.Row="1">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <TextBlock Text="Detected Quarantined Devices" FontWeight="SemiBold" Foreground="#94A3B8" Margin="0,0,0,8"/>
                        
                        <ListBox Grid.Row="1" x:Name="LstDevices" Background="#0F172A" BorderBrush="#334155" BorderThickness="1" Margin="0,0,0,5">
                            <ListBox.ItemTemplate>
                                <DataTemplate>
                                    <Border BorderBrush="#1E293B" BorderThickness="0,0,0,1" Padding="8,10" HorizontalAlignment="Stretch">
                                        <Grid Width="500">
                                            <Grid.ColumnDefinitions>
                                                <ColumnDefinition Width="*"/>
                                                <ColumnDefinition Width="Auto"/>
                                            </Grid.ColumnDefinitions>
                                            
                                            <StackPanel>
                                                <TextBlock Text="{Binding DeviceModel}" FontWeight="Bold" Foreground="#F8FAFC" FontSize="14"/>
                                                <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
                                                    <TextBlock Text="Type: " Foreground="#64748B" FontSize="11"/>
                                                    <TextBlock Text="{Binding DeviceType}" Foreground="#94A3B8" FontSize="11" Margin="0,0,15,0"/>
                                                    <TextBlock Text="ID: " Foreground="#64748B" FontSize="11"/>
                                                    <TextBlock Text="{Binding DeviceId}" Foreground="#38BDF8" FontSize="11"/>
                                                </StackPanel>
                                            </StackPanel>
                                            
                                            <Border Grid.Column="1" Background="#78350F" BorderBrush="#D97706" BorderThickness="1" CornerRadius="4" Padding="8,4" VerticalAlignment="Center" HorizontalAlignment="Right">
                                                <TextBlock Text="{Binding DeviceAccessState}" Foreground="#FBBF24" FontSize="11" FontWeight="Bold"/>
                                            </Border>
                                        </Grid>
                                    </Border>
                                </DataTemplate>
                            </ListBox.ItemTemplate>
                        </ListBox>

                        <!-- Empty/Feedback States overlaying the ListBox -->
                        <TextBlock Grid.Row="1" x:Name="TxtNoDevices" Text="No quarantined devices found for this user." HorizontalAlignment="Center" VerticalAlignment="Center" Foreground="#94A3B8" Visibility="Collapsed"/>
                        <TextBlock Grid.Row="1" x:Name="TxtLoadingDevices" Text="Querying Mobile Devices..." HorizontalAlignment="Center" VerticalAlignment="Center" Foreground="#3B82F6" FontWeight="Bold" Visibility="Collapsed"/>
                    </Grid>

                    <!-- Actions Button Panel -->
                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
                        <Button x:Name="BtnCheckDevices" Content="Re-Scan Status" Style="{StaticResource SecondaryButton}" Width="160" Height="40" Margin="0,0,12,0" IsEnabled="False"/>
                        <Button x:Name="BtnRelease" Content="Release Quarantined Devices" Style="{StaticResource ActionButton}" Width="240" Height="40" IsEnabled="False"/>
                    </StackPanel>
                </Grid>
            </Border>
        </Grid>

        <!-- LIVE LOG OUTPUT -->
        <Border Grid.Row="2" Background="#020617" BorderBrush="#1E293B" BorderThickness="0,1,0,0" Padding="20,10">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <TextBlock Text="Application Activity Logs" FontSize="11" FontWeight="Bold" Foreground="#475569" Margin="0,0,0,6"/>
                <!-- Console textbox -->
                <TextBox Grid.Row="1" x:Name="TxtConsole" Background="Transparent" BorderThickness="0" FontFamily="Consolas" FontSize="12" Foreground="#38BDF8" IsReadOnly="True" VerticalScrollBarVisibility="Auto" AcceptsReturn="True" TextWrapping="Wrap"/>
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

# Console/Log Writer helper
function Write-Log {
    param(
        [string]$Message,
        [string]$Type = "Info"
    )
    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $prefix = "[$timestamp] "
    switch ($Type) {
        "Error"   { $colorPrefix = "[ERROR] "; $msg = "$prefix$colorPrefix$Message" }
        "Success" { $colorPrefix = "[SUCCESS] "; $msg = "$prefix$colorPrefix$Message" }
        "Warning" { $colorPrefix = "[WARNING] "; $msg = "$prefix$colorPrefix$Message" }
        Default   { $colorPrefix = "[INFO] "; $msg = "$prefix$colorPrefix$Message" }
    }
    
    # Run on UI thread to prevent dispatcher exceptions
    $TxtConsole.Dispatcher.Invoke([Action]{
        $TxtConsole.AppendText("$msg`r`n")
        $TxtConsole.ScrollToEnd()
    })
}

# Reset GUI UI elements to initial state
function Reset-UIState {
    $StatusDot.Background = [System.Windows.Media.Brushes]::Red
    $StatusText.Text = "Disconnected"
    $StatusText.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(148, 163, 184)) # Slate 400
    
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
    Write-Log "Connecting to Exchange Online. Please authenticating via modern login dialog..."
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
        $StatusDot.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(16, 185, 129)) # Emerald 500
        $StatusText.Text = "Connected ($adminUser)"
        $StatusText.Foreground = [System.Windows.Media.Brushes]::White
        
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
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
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
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    } catch {}
})

# Splash Message on Startup
Write-Log "Application initialized. Click 'Connect to Exchange Online' to get started."

# Launch the WPF application
[void]$Form.ShowDialog()
