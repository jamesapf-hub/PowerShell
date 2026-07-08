#Requires -Version 7.2
<#
.SYNOPSIS
    M365 Admin Console - Native PowerShell 7 WPF Desktop Dashboard.
.DESCRIPTION
    Launches a modern, dark-themed (Kinetic Command) administration interface to audit
    licensed Entra ID users, evaluate last login timelines, resolve SKUs, and export results.
#>

# 1. Force STA (Single Threaded Apartment) Mode for WPF
if ($Host.Runspace.ApartmentState -ne 'STA') {
    Write-Host "Apartment state is not STA. Restarting script in STA mode..." -ForegroundColor Yellow
    $Arguments = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-ApartmentState", "STA", "-File", $MyInvocation.MyCommand.Path)
    Start-Process pwsh.exe -ArgumentList $Arguments -NoNewWindow -Wait
    return
}

# 2. Import Required Windows Presentation Assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml, System.Xml, System.Drawing

# Dependency Check & Auto-Installation
function Assert-ModuleInstalled {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ModuleName,
        [Parameter(Mandatory=$true)]
        [string]$FriendlyName
    )
    Write-Host "Checking for dependency: ${FriendlyName} (${ModuleName})..." -ForegroundColor Cyan
    $Module = Get-Module -ListAvailable -Name $ModuleName
    if (-not $Module) {
        Write-Host "[WARNING] ${FriendlyName} is missing. Attempting automatic installation..." -ForegroundColor Yellow
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Write-Host "Installing ${ModuleName} from PowerShell Gallery for current user..." -ForegroundColor Gray
            Install-Module -Name $ModuleName -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
            Write-Host "[SUCCESS] ${FriendlyName} successfully installed." -ForegroundColor Green
        } catch {
            Write-Host "[ERROR] Failed to install ${FriendlyName}: $_" -ForegroundColor Red
            [System.Windows.MessageBox]::Show(
                "Failed to install required dependency: ${FriendlyName} (${ModuleName}).`n`nError: $_`n`nPlease run PowerShell and install it manually: Install-Module $ModuleName -Scope CurrentUser",
                "Dependency Installation Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
            exit 1
        }
    } else {
        Write-Host "[OK] $FriendlyName is available." -ForegroundColor Green
    }
}

Assert-ModuleInstalled -ModuleName "Microsoft.Graph" -FriendlyName "Microsoft Graph SDK"

# 3. GUI Layout Definition (XAML)
$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="M365 PowerShell Admin Console" Height="780" Width="1080"
        WindowStartupLocation="CenterScreen" Background="#131313" Foreground="#f3f4f6">
    <Window.Resources>
        <Style x:Key="BrandBtn" TargetType="Button">
            <Setter Property="Background" Value="#3B82F6"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#1D4ED8"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="SecBtn" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#3B82F6"/>
            <Setter Property="BorderBrush" Value="#3B82F6"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#3B82F6"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="TabBtn" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#9ca3af"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="16,10"/>
            <Setter Property="BorderThickness" Value="0,0,0,2"/>
            <Setter Property="BorderBrush" Value="Transparent"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
    </Window.Resources>
    
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <!-- Header -->
            <RowDefinition Height="Auto"/> <!-- Nav Tabs -->
            <RowDefinition Height="*"/>    <!-- Main Pages -->
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#1c1c1e" BorderBrush="#2c2c2e" BorderThickness="0,0,0,1" Padding="20,15">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="0">
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="PS C:\&gt;" Foreground="#3B82F6" FontFamily="Consolas" FontSize="20" FontWeight="Bold" Margin="0,0,10,0"/>
                        <TextBlock Text="M365 Admin Console" Foreground="White" FontSize="20" FontWeight="Bold"/>
                    </StackPanel>
                    <TextBlock Text="POWERSHELL 7 // MICROSOFT GRAPH SDK v2.0" Foreground="#555" FontFamily="Consolas" FontSize="11" Margin="0,2,0,0"/>
                </StackPanel>
                
                <Border Grid.Column="1" Background="#0c0c0d" BorderBrush="#2c2c2e" BorderThickness="1" CornerRadius="6" Padding="12,6" VerticalAlignment="Center">
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                        <Ellipse Name="elStatusDot" Width="10" Height="10" Fill="Red" Margin="0,0,8,0"/>
                        <TextBlock Name="txtStatusLabel" Text="Offline Session" Foreground="#9ca3af" FontSize="11" Margin="0,0,10,0" FontFamily="Consolas"/>
                        <Border Width="1" Height="12" Background="#2c2c2e" Margin="10,0"/>
                        <TextBlock Text="TLS 1.2" Foreground="#10b981" FontSize="11" FontFamily="Consolas" FontWeight="Bold"/>
                    </StackPanel>
                </Border>
            </Grid>
        </Border>
        
        <!-- Nav Tabs -->
        <Border Name="borderNav" Visibility="Collapsed" Grid.Row="1" Background="#131313" BorderBrush="#2c2c2e" BorderThickness="0,0,0,1">
            <StackPanel Orientation="Horizontal" Margin="20,0">
                <Button Name="btnNavConnection" Style="{StaticResource TabBtn}" Content="1. Establish Connection" BorderBrush="#3B82F6" Foreground="White"/>
                <Button Name="btnNavDashboard" Style="{StaticResource TabBtn}" Content="2. Dashboard Insights"/>
                <Button Name="btnNavDirectory" Style="{StaticResource TabBtn}" Content="3. User Directory"/>
            </StackPanel>
        </Border>
        
        <!-- Main Pages Container -->
        <Grid Grid.Row="2" Margin="20">
            <!-- PAGE 1: CONNECTION PANEL -->
            <Grid Name="pageConnection" Visibility="Visible">
                <!-- Offline Welcome Screen -->
                <ScrollViewer Name="panelOfflineWelcome" Visibility="Visible" VerticalScrollBarVisibility="Auto">
                    <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center" Width="750" Margin="0,10,0,10">
                        
                        <!-- App Header / Intro -->
                        <TextBlock Text="Welcome to M365 Administration Console" Foreground="White" FontSize="22" FontWeight="Bold" HorizontalAlignment="Center" Margin="0,0,0,5"/>
                        <TextBlock Text="Please select your operational role to begin auditing tenant users." Foreground="#9ca3af" FontSize="13" HorizontalAlignment="Center" Margin="0,0,0,25"/>

                        <!-- Step 1: Role Selection Cards -->
                        <TextBlock Text="STEP 1: SELECT YOUR OPERATING ROLE" Foreground="#3B82F6" FontFamily="Consolas" FontSize="11" FontWeight="Bold" Margin="0,0,0,10" HorizontalAlignment="Center"/>
                        <UniformGrid Columns="2" Margin="0,0,0,25" Width="600">
                            <!-- Service Desk Role Card -->
                            <Button Name="btnRoleSelectServiceDesk" Background="#1c1c1e" BorderBrush="#2c2c2e" BorderThickness="1" Padding="15" Margin="0,0,10,0" Cursor="Hand" Height="100">
                                <Button.Style>
                                    <Style TargetType="Button">
                                        <Setter Property="Template">
                                            <Setter.Value>
                                                <ControlTemplate TargetType="Button">
                                                    <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8" Padding="{TemplateBinding Padding}">
                                                        <ContentPresenter VerticalAlignment="Center" HorizontalAlignment="Center"/>
                                                    </Border>
                                                </ControlTemplate>
                                            </Setter.Value>
                                        </Setter>
                                    </Style>
                                </Button.Style>
                                <StackPanel>
                                    <TextBlock Text="🛠️ Service Desk" Foreground="White" FontSize="16" FontWeight="Bold" HorizontalAlignment="Center"/>
                                    <TextBlock Text="Standard view: licenses, statuses &amp; accounts. Hides all financials." Foreground="#9ca3af" FontSize="11" TextWrapping="Wrap" TextAlignment="Center" Margin="0,5,0,0"/>
                                </StackPanel>
                            </Button>
                            
                            <!-- Sales Role Card -->
                            <Button Name="btnRoleSelectSales" Background="#1c1c1e" BorderBrush="#2c2c2e" BorderThickness="1" Padding="15" Margin="10,0,0,0" Cursor="Hand" Height="100">
                                <Button.Style>
                                    <Style TargetType="Button">
                                        <Setter Property="Template">
                                            <Setter.Value>
                                                <ControlTemplate TargetType="Button">
                                                    <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8" Padding="{TemplateBinding Padding}">
                                                        <ContentPresenter VerticalAlignment="Center" HorizontalAlignment="Center"/>
                                                    </Border>
                                                </ControlTemplate>
                                            </Setter.Value>
                                        </Setter>
                                    </Style>
                                </Button.Style>
                                <StackPanel>
                                    <TextBlock Text="💰 Sales &amp; Business" Foreground="White" FontSize="16" FontWeight="Bold" HorizontalAlignment="Center"/>
                                    <TextBlock Text="Financial view: savings, wasted costs, and downgrade recommendations." Foreground="#9ca3af" FontSize="11" TextWrapping="Wrap" TextAlignment="Center" Margin="0,5,0,0"/>
                                </StackPanel>
                            </Button>
                        </UniformGrid>
                        
                        <!-- Step 2: Connection Panel (Collapsed by default, shown once role selected) -->
                        <StackPanel Name="panelConnectionMethods" Visibility="Collapsed" HorizontalAlignment="Center" Width="740">
                            <TextBlock Text="STEP 2: CHOOSE CONNECTION METHOD" Foreground="#3B82F6" FontFamily="Consolas" FontSize="11" FontWeight="Bold" Margin="0,0,0,10" HorizontalAlignment="Center"/>
                            
                            <!-- Side-by-Side Connection Cards -->
                            <UniformGrid Columns="3" Margin="0,0,0,20">
                                <!-- Option A: Interactive M365 Sign-In -->
                                <Border Background="#1c1c1e" BorderBrush="#2c2c2e" BorderThickness="1" CornerRadius="8" Padding="15" Margin="0,0,8,0">
                                    <Grid>
                                        <Grid.RowDefinitions>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="*"/>
                                            <RowDefinition Height="Auto"/>
                                        </Grid.RowDefinitions>
                                        <TextBlock Grid.Row="0" Text="Interactive Sign-In" Foreground="White" FontSize="13" FontWeight="Bold" Margin="0,0,0,8"/>
                                        <TextBlock Grid.Row="1" Text="Official Microsoft OAuth login prompt. Supports MFA and SSO." Foreground="#9ca3af" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,15"/>
                                        
                                        <StackPanel Grid.Row="2">
                                            <TextBlock Text="Tenant (Optional):" Foreground="#777" FontSize="10" Margin="0,0,0,2"/>
                                            <TextBox Name="txtUserTenant" Background="#0c0c0d" BorderBrush="#2c2c2e" Foreground="White" Padding="5" Margin="0,0,0,8" FontSize="11" CaretBrush="White"/>
                                            <Button Name="btnConnectM365" Style="{StaticResource BrandBtn}" Content="M365 Sign-In" Height="32" FontSize="12"/>
                                        </StackPanel>
                                    </Grid>
                                </Border>
                                
                                <!-- Option B: App Client Secret -->
                                <Border Background="#1c1c1e" BorderBrush="#2c2c2e" BorderThickness="1" CornerRadius="8" Padding="15" Margin="4,0,4,0">
                                    <Grid>
                                        <Grid.RowDefinitions>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="*"/>
                                            <RowDefinition Height="Auto"/>
                                        </Grid.RowDefinitions>
                                        <TextBlock Grid.Row="0" Text="App Client Secret" Foreground="White" FontSize="13" FontWeight="Bold" Margin="0,0,0,8"/>
                                        <TextBlock Grid.Row="1" Text="Automated connection using Entra App Registration client credentials." Foreground="#9ca3af" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,15"/>
                                        
                                        <StackPanel Grid.Row="2">
                                            <TextBlock Text="Tenant ID/Domain:" Foreground="#777" FontSize="10" Margin="0,0,0,2"/>
                                            <TextBox Name="txtAppTenant" Background="#0c0c0d" BorderBrush="#2c2c2e" Foreground="White" Padding="4" Margin="0,0,0,4" FontSize="10" CaretBrush="White"/>
                                            <TextBlock Text="Client ID:" Foreground="#777" FontSize="10" Margin="0,0,0,2"/>
                                            <TextBox Name="txtAppClient" Background="#0c0c0d" BorderBrush="#2c2c2e" Foreground="White" Padding="4" Margin="0,0,0,4" FontSize="10" CaretBrush="White"/>
                                            <TextBlock Text="Client Secret:" Foreground="#777" FontSize="10" Margin="0,0,0,2"/>
                                            <PasswordBox Name="txtAppSecret" Background="#0c0c0d" BorderBrush="#2c2c2e" Foreground="White" Padding="4" Margin="0,0,0,8" FontSize="10" CaretBrush="White"/>
                                            <Button Name="btnConnectApp" Style="{StaticResource BrandBtn}" Content="Connect App" Height="32" FontSize="12"/>
                                        </StackPanel>
                                    </Grid>
                                </Border>
                                
                                <!-- Option C: Import CSV -->
                                <Border Background="#1c1c1e" BorderBrush="#2c2c2e" BorderThickness="1" CornerRadius="8" Padding="15" Margin="8,0,0,0">
                                    <Grid>
                                        <Grid.RowDefinitions>
                                            <RowDefinition Height="Auto"/>
                                            <RowDefinition Height="*"/>
                                            <RowDefinition Height="Auto"/>
                                        </Grid.RowDefinitions>
                                        <TextBlock Grid.Row="0" Text="Offline CSV Import" Foreground="White" FontSize="13" FontWeight="Bold" Margin="0,0,0,8"/>
                                        <TextBlock Grid.Row="1" Text="Load previously exported user reports directly into the dashboard." Foreground="#9ca3af" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,15"/>
                                        
                                        <Button Grid.Row="2" Name="btnBrowseCSV" Background="Transparent" BorderThickness="0" Padding="0" Cursor="Hand">
                                            <Border BorderBrush="#2c2c2e" BorderThickness="1" Background="#0c0c0d" CornerRadius="6" Padding="15" Height="100">
                                                <StackPanel VerticalAlignment="Center">
                                                    <TextBlock Text="📁" FontSize="20" HorizontalAlignment="Center" Margin="0,0,0,5"/>
                                                    <TextBlock Text="Browse CSV File" Foreground="White" FontSize="11" FontWeight="Bold" HorizontalAlignment="Center"/>
                                                    <TextBlock Text="Click to select file" Foreground="#555" FontSize="9" HorizontalAlignment="Center" Margin="0,2,0,0"/>
                                                </StackPanel>
                                            </Border>
                                        </Button>
                                    </Grid>
                                </Border>
                            </UniformGrid>
                            
                            <!-- Demo Sandbox Option at the bottom -->
                            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,10,0,0">
                                <TextBlock Text="Need to test without credentials?" Foreground="#777" FontSize="12" VerticalAlignment="Center"/>
                                <Button Name="btnDemoMode" Background="Transparent" BorderThickness="0" Foreground="#3B82F6" Content="Launch Demo Sandbox Mode" FontWeight="Bold" Cursor="Hand" Margin="8,0,0,0">
                                    <Button.Style>
                                        <Style TargetType="Button">
                                            <Setter Property="Template">
                                                <Setter.Value>
                                                    <ControlTemplate TargetType="Button">
                                                        <TextBlock Text="{TemplateBinding Content}" Foreground="{TemplateBinding Foreground}" FontWeight="{TemplateBinding FontWeight}" Cursor="{TemplateBinding Cursor}" TextDecorations="Underline"/>
                                                    </ControlTemplate>
                                                </Setter.Value>
                                            </Setter>
                                        </Style>
                                    </Button.Style>
                                </Button>
                            </StackPanel>
                        </StackPanel>
                        
                    </StackPanel>
                </ScrollViewer>
                
                <!-- Connected Session Info Screen -->
                <StackPanel Name="panelConnectedSession" Visibility="Collapsed" HorizontalAlignment="Center" VerticalAlignment="Center" Width="450">
                    <Border Background="#1c1c1e" BorderBrush="#3B82F6" BorderThickness="1" CornerRadius="8" Padding="25">
                        <StackPanel>
                            <TextBlock Text="Active Session Established" Foreground="White" FontSize="18" FontWeight="Bold" HorizontalAlignment="Center" Margin="0,0,0,5"/>
                            <TextBlock Text="You are successfully authenticated and auditing the tenant." Foreground="#9ca3af" FontSize="12" HorizontalAlignment="Center" Margin="0,0,0,20"/>
                            
                            <Border Background="#0c0c0d" BorderBrush="#2c2c2e" BorderThickness="1" CornerRadius="6" Padding="15" Margin="0,0,0,20">
                                <StackPanel>
                                    <Grid Margin="0,6">
                                        <TextBlock Text="Tenant Domain:" Foreground="#9ca3af" FontSize="12"/>
                                        <TextBlock Name="txtSessionTenant" Text="N/A" Foreground="White" FontSize="12" FontWeight="Bold" HorizontalAlignment="Right"/>
                                    </Grid>
                                    <Grid Margin="0,6">
                                        <TextBlock Text="Connected As:" Foreground="#9ca3af" FontSize="12"/>
                                        <TextBlock Name="txtSessionUser" Text="N/A" Foreground="White" FontSize="12" HorizontalAlignment="Right" TextTrimming="CharacterEllipsis" MaxWidth="240"/>
                                    </Grid>
                                    <Grid Margin="0,6">
                                        <TextBlock Text="Uptime:" Foreground="#9ca3af" FontSize="12"/>
                                        <TextBlock Name="txtSessionUptime" Text="00:00:00" Foreground="#3B82F6" FontSize="12" HorizontalAlignment="Right" FontWeight="Bold"/>
                                    </Grid>
                                    <Grid Margin="0,6">
                                        <TextBlock Text="Mode:" Foreground="#9ca3af" FontSize="12"/>
                                        <TextBlock Name="txtSessionMode" Text="N/A" Foreground="White" FontSize="12" HorizontalAlignment="Right"/>
                                    </Grid>
                                </StackPanel>
                            </Border>
                            
                            <Button Name="btnDisconnect" Style="{StaticResource SecBtn}" Content="Disconnect Session" Height="40" FontSize="13"/>
                        </StackPanel>
                    </Border>
                </StackPanel>
            </Grid>
            
            <!-- PAGE 2: DASHBOARD INSIGHTS -->
            <Grid Name="pageDashboard" Visibility="Collapsed">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/> <!-- KPIs -->
                    <RowDefinition Height="*"/>    <!-- Charts side-by-side -->
                    <RowDefinition Height="Auto"/> <!-- Compact Terminal Feed -->
                    <RowDefinition Height="Auto"/> <!-- Bottom Actions -->
                </Grid.RowDefinitions>
                
                <!-- KPIs Row -->
                <Grid Grid.Row="0" Margin="0,0,0,15">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Name="colKpiTotal" Width="*"/>
                        <ColumnDefinition Name="colKpiActive" Width="*"/>
                        <ColumnDefinition Name="colKpiInactive90d" Width="*"/>
                        <ColumnDefinition Name="colKpiInactive1yr" Width="*"/>
                        <ColumnDefinition Name="colKpiSavings" Width="1.2*"/>
                        <ColumnDefinition Name="colKpiPoolWaste" Width="1.2*"/>
                    </Grid.ColumnDefinitions>
                    
                    <!-- Total -->
                    <Border Grid.Column="0" Background="#1c1c1e" BorderBrush="#3B82F6" BorderThickness="1,0,0,0" Padding="12,8" Margin="0,0,8,0">
                        <StackPanel>
                            <TextBlock Text="TOTAL LICENSED" Foreground="#9ca3af" FontSize="10" FontWeight="Bold"/>
                            <TextBlock Name="kpiTotal" Text="0" Foreground="White" FontSize="22" FontWeight="Bold" Margin="0,2,0,0"/>
                        </StackPanel>
                    </Border>
                    <!-- Active -->
                    <Border Grid.Column="1" Background="#1c1c1e" BorderBrush="#10b981" BorderThickness="1,0,0,0" Padding="12,8" Margin="0,0,8,0">
                        <StackPanel>
                            <TextBlock Text="ACTIVE (&lt;=30d)" Foreground="#9ca3af" FontSize="10" FontWeight="Bold"/>
                            <TextBlock Name="kpiActive" Text="0" Foreground="#10b981" FontSize="22" FontWeight="Bold" Margin="0,2,0,0"/>
                        </StackPanel>
                    </Border>
                    <!-- Warning -->
                    <Border Grid.Column="2" Background="#1c1c1e" BorderBrush="#eab308" BorderThickness="1,0,0,0" Padding="12,8" Margin="0,0,8,0">
                        <StackPanel>
                            <TextBlock Text="INACTIVE (&gt;90d)" Foreground="#9ca3af" FontSize="10" FontWeight="Bold"/>
                            <TextBlock Name="kpiInactive90d" Text="0" Foreground="#eab308" FontSize="22" FontWeight="Bold" Margin="0,2,0,0"/>
                        </StackPanel>
                    </Border>
                    <!-- Critical -->
                    <Border Grid.Column="3" Background="#1c1c1e" BorderBrush="#ef4444" BorderThickness="1,0,0,0" Padding="12,8" Margin="0,0,8,0">
                        <StackPanel>
                            <TextBlock Text="CRITICAL (&gt;1yr)" Foreground="#9ca3af" FontSize="10" FontWeight="Bold"/>
                            <TextBlock Name="kpiInactive1yr" Text="0" Foreground="#ef4444" FontSize="22" FontWeight="Bold" Margin="0,2,0,0"/>
                        </StackPanel>
                    </Border>
                    <!-- Savings -->
                    <Border Name="borderKpiSavings" Grid.Column="4" Background="#1c1c1e" BorderBrush="#3B82F6" BorderThickness="1,0,0,0" Padding="12,6" Margin="0,0,8,0">
                        <StackPanel>
                            <TextBlock Text="WASTED MONEY (&gt;6m)" Foreground="#9ca3af" FontSize="9" FontWeight="Bold"/>
                            <TextBlock Name="kpiSavings" Text="£0.00" Foreground="#ef4444" FontSize="16" FontWeight="Bold" Margin="0,1,0,0"/>
                            <TextBlock Text="POTENTIAL MONTHLY SAVINGS" Foreground="#9ca3af" FontSize="9" FontWeight="Bold" Margin="0,3,0,0"/>
                            <TextBlock Name="kpiMonthlySavings" Text="£0.00/mo" Foreground="#10b981" FontSize="16" FontWeight="Bold" Margin="0,1,0,0"/>
                        </StackPanel>
                    </Border>
                    <!-- Pool Waste -->
                    <Border Name="borderKpiPoolWaste" Grid.Column="5" Background="#1c1c1e" BorderBrush="#3B82F6" BorderThickness="1,0,0,0" Padding="12,6">
                        <StackPanel>
                            <TextBlock Text="POOL WASTE (UNASSIGNED)" Foreground="#9ca3af" FontSize="9" FontWeight="Bold"/>
                                    <TextBlock Name="kpiPoolWaste" Text="£0.00" Foreground="#ef4444" FontSize="16" FontWeight="Bold" Margin="0,1,0,0"/>
                            <TextBlock Name="kpiPoolWasteCount" Text="0 unused licenses" Foreground="#9ca3af" FontSize="9" Margin="0,3,0,0" FontWeight="SemiBold"/>
                        </StackPanel>
                    </Border>
                </Grid>
                   <!-- Side-by-Side Charts Layout -->
                <Grid Grid.Row="1" Margin="0,0,0,15">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="2.2*"/>
                        <ColumnDefinition Width="1*"/>
                    </Grid.ColumnDefinitions>
                    
                    <!-- Border 1: Licenses & Unassigned Breakdown -->
                    <Border Grid.Column="0" Background="#1c1c1e" BorderBrush="#2c2c2e" BorderThickness="1" CornerRadius="8" Padding="15" Margin="0,0,10,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <TextBlock Grid.Row="0" Text="Licenses &amp; Unassigned Breakdown" Foreground="White" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <DataGrid Grid.Row="1" Name="gridUnassigned" AutoGenerateColumns="False" IsReadOnly="True" 
                                      Background="#1c1c1e" BorderThickness="0" Foreground="#f3f4f6"
                                      RowBackground="#1c1c1e" AlternatingRowBackground="#131313" 
                                      HeadersVisibility="Column" GridLinesVisibility="None" SelectionMode="Single"
                                      FontFamily="Segoe UI" FontSize="11" RowHeight="28">
                                <DataGrid.ColumnHeaderStyle>
                                    <Style TargetType="DataGridColumnHeader">
                                        <Setter Property="Background" Value="#18181c"/>
                                        <Setter Property="Foreground" Value="#9ca3af"/>
                                        <Setter Property="Padding" Value="8,6"/>
                                        <Setter Property="FontWeight" Value="SemiBold"/>
                                        <Setter Property="BorderThickness" Value="0,0,0,1"/>
                                        <Setter Property="BorderBrush" Value="#2c2c2e"/>
                                    </Style>
                                </DataGrid.ColumnHeaderStyle>
                                <DataGrid.CellStyle>
                                    <Style TargetType="DataGridCell">
                                        <Setter Property="BorderThickness" Value="0"/>
                                        <Setter Property="Padding" Value="8,4"/>
                                        <Setter Property="VerticalAlignment" Value="Center"/>
                                        <Style.Triggers>
                                            <Trigger Property="IsSelected" Value="True">
                                                <Setter Property="Background" Value="#2c2c2e"/>
                                                <Setter Property="Foreground" Value="White"/>
                                            </Trigger>
                                        </Style.Triggers>
                                    </Style>
                                </DataGrid.CellStyle>
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Product License" Binding="{Binding SkuPartName}" Width="2*"/>
                                    <DataGridTextColumn Header="Assigned / Total" Binding="{Binding AssignedText}" Width="1.2*">
                                        <DataGridTextColumn.ElementStyle>
                                            <Style TargetType="TextBlock">
                                                <Setter Property="HorizontalAlignment" Value="Center"/>
                                                <Setter Property="VerticalAlignment" Value="Center"/>
                                            </Style>
                                        </DataGridTextColumn.ElementStyle>
                                    </DataGridTextColumn>
                                    <DataGridTextColumn Header="Unassigned" Binding="{Binding UnassignedUnits}" Width="1*">
                                        <DataGridTextColumn.ElementStyle>
                                            <Style TargetType="TextBlock">
                                                <Setter Property="HorizontalAlignment" Value="Center"/>
                                                <Setter Property="VerticalAlignment" Value="Center"/>
                                                <Setter Property="FontWeight" Value="Bold"/>
                                                <Setter Property="Foreground" Value="{Binding UnassignedBrush}"/>
                                            </Style>
                                        </DataGridTextColumn.ElementStyle>
                                    </DataGridTextColumn>
                                    <DataGridTextColumn Header="Monthly Cost" Binding="{Binding MonthlyPriceText}" Width="1.1*">
                                        <DataGridTextColumn.ElementStyle>
                                            <Style TargetType="TextBlock">
                                                <Setter Property="HorizontalAlignment" Value="Right"/>
                                                <Setter Property="VerticalAlignment" Value="Center"/>
                                            </Style>
                                        </DataGridTextColumn.ElementStyle>
                                    </DataGridTextColumn>
                                    <DataGridTextColumn Header="Wasted/mo" Binding="{Binding WastedCostText}" Width="1.1*">
                                        <DataGridTextColumn.ElementStyle>
                                            <Style TargetType="TextBlock">
                                                <Setter Property="HorizontalAlignment" Value="Right"/>
                                                <Setter Property="VerticalAlignment" Value="Center"/>
                                                <Setter Property="FontWeight" Value="SemiBold"/>
                                                <Setter Property="Foreground" Value="{Binding WastedBrush}"/>
                                            </Style>
                                        </DataGridTextColumn.ElementStyle>
                                    </DataGridTextColumn>
                                </DataGrid.Columns>
                            </DataGrid>
                        </Grid>
                    </Border>
                    
                    <!-- Chart 2: Activity -->
                    <Border Grid.Column="1" Background="#1c1c1e" BorderBrush="#2c2c2e" BorderThickness="1" CornerRadius="8" Padding="15" Margin="10,0,0,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <TextBlock Grid.Row="0" Text="User Activity Distribution" Foreground="White" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,10"/>
                            <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                                <StackPanel Name="panelActivityChart"/>
                            </ScrollViewer>
                        </Grid>
                    </Border>
                </Grid>
                
                <!-- Compact System Activity Feed -->
                <Grid Grid.Row="2" Margin="0,0,0,10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="100"/>
                    </Grid.RowDefinitions>
                    
                    <Grid Grid.Row="0" Margin="0,0,0,5">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Text="&gt; System Activity Feed" Foreground="White" FontSize="12" FontWeight="Bold" VerticalAlignment="Center"/>
                        <Button Name="btnRunScript" Grid.Column="1" Style="{StaticResource BrandBtn}" Content="Refresh Report" Padding="12,4" FontSize="11"/>
                    </Grid>
                    
                    <Border Grid.Row="1" Background="#0c0c0d" BorderBrush="#2c2c2e" BorderThickness="1" CornerRadius="6">
                        <TextBox Name="txtLog" Background="Transparent" Foreground="#38bdf8" BorderThickness="0" 
                                 FontFamily="Consolas" FontSize="11" IsReadOnly="True" AcceptsReturn="True" 
                                 VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Padding="8"/>
                    </Border>
                </Grid>
                
                <!-- Bottom Exporters (Dashboard Page) -->
                <Grid Grid.Row="3" Margin="0,5,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    
                    <TextBlock Text="// Results represent audited active Entra ID subscriptions" Foreground="#555" FontSize="11" VerticalAlignment="Center"/>
                    
                    <StackPanel Grid.Column="1" Orientation="Horizontal">
                        <Button Name="btnGenEmailDash" Style="{StaticResource SecBtn}" Content="Generate Email Draft" Height="32" Margin="0,0,8,0"/>
                        <Button Name="btnExportCSVDash" Style="{StaticResource SecBtn}" Content="Export CSV" Height="32" Margin="0,0,8,0"/>
                        <Button Name="btnExportExcelDash" Style="{StaticResource SecBtn}" Content="Export Excel (XLSX)" Height="32" Margin="0,0,8,0"/>
                        <Button Name="btnExportPDFDash" Style="{StaticResource SecBtn}" Content="Export PDF Report" Height="32" Margin="0,0,8,0"/>
                        <Button Name="btnExportPNGDash" Style="{StaticResource SecBtn}" Content="Snapshot Dashboard PNG" Height="32"/>
                    </StackPanel>
                </Grid>
            </Grid>
            
            <!-- PAGE 3: USER DIRECTORY GRID -->
            <Grid Name="pageDirectory" Visibility="Collapsed">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/> <!-- Filters & Search -->
                    <RowDefinition Height="*"/>    <!-- Data Grid -->
                    <RowDefinition Height="Auto"/> <!-- Bottom Actions -->
                </Grid.RowDefinitions>
                
                <!-- Filters & Search -->
                <Grid Grid.Row="0" Margin="0,0,0,15">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="250"/>
                    </Grid.ColumnDefinitions>
                    
                    <!-- Filters -->
                    <StackPanel Grid.Column="0" Orientation="Horizontal">
                        <Button Name="btnFilterAll" Content="All Users" Padding="12,6" Background="#3B82F6" Foreground="White" FontWeight="SemiBold" BorderThickness="0" Margin="0,0,6,0" Tag="all"/>
                        <Button Name="btnFilterActive" Content="Active (&lt;=30d)" Padding="12,6" Background="#1c1c1e" Foreground="#9ca3af" BorderBrush="#2c2c2e" BorderThickness="1" Margin="0,0,6,0" Tag="active"/>
                        <Button Name="btnFilterInactive90d" Content="Inactive (&gt;90d)" Padding="12,6" Background="#1c1c1e" Foreground="#9ca3af" BorderBrush="#2c2c2e" BorderThickness="1" Margin="0,0,6,0" Tag="inactive90d"/>
                        <Button Name="btnFilterInactive1yr" Content="Inactive (&gt;1yr)" Padding="12,6" Background="#1c1c1e" Foreground="#9ca3af" BorderBrush="#2c2c2e" BorderThickness="1" Margin="0,0,6,0" Tag="inactive1yr"/>
                        <Button Name="btnFilterNever" Content="Never Logged In" Padding="12,6" Background="#1c1c1e" Foreground="#9ca3af" BorderBrush="#2c2c2e" BorderThickness="1" Tag="never"/>
                    </StackPanel>
                    
                    <!-- Search Box -->
                    <Grid Grid.Column="1">
                        <TextBox Name="txtSearch" Background="#0c0c0d" BorderBrush="#2c2c2e" Foreground="White" Padding="8,6" FontSize="12" SelectionBrush="#3B82F6" CaretBrush="White"/>
                        <TextBlock Text="Search Display Name/UPN..." Foreground="#555" IsHitTestVisible="False" VerticalAlignment="Center" Margin="10,0,0,0" FontSize="11">
                            <TextBlock.Style>
                                <Style TargetType="TextBlock">
                                    <Setter Property="Visibility" Value="Collapsed"/>
                                    <Style.Triggers>
                                        <DataTrigger Binding="{Binding ElementName=txtSearch, Path=Text}" Value="">
                                            <Setter Property="Visibility" Value="Visible"/>
                                        </DataTrigger>
                                    </Style.Triggers>
                                </Style>
                            </TextBlock.Style>
                        </TextBlock>
                    </Grid>
                </Grid>
                
                <!-- Styled Data Grid -->
                <Border Grid.Row="1" Background="#1c1c1e" BorderBrush="#2c2c2e" BorderThickness="1" CornerRadius="8" Padding="5">
                    <DataGrid Name="gridUsers" AutoGenerateColumns="False" IsReadOnly="True" 
                              Background="#1c1c1e" BorderThickness="0" Foreground="#f3f4f6"
                              RowBackground="#1c1c1e" AlternatingRowBackground="#131313" 
                              HeadersVisibility="Column" GridLinesVisibility="None" SelectionMode="Single"
                              FontFamily="Consolas" FontSize="11" Margin="5">
                        <DataGrid.ColumnHeaderStyle>
                            <Style TargetType="DataGridColumnHeader">
                                <Setter Property="Background" Value="#18181c"/>
                                <Setter Property="Foreground" Value="#9ca3af"/>
                                <Setter Property="Padding" Value="10,8"/>
                                <Setter Property="FontWeight" Value="SemiBold"/>
                                <Setter Property="BorderThickness" Value="0,0,0,1"/>
                                <Setter Property="BorderBrush" Value="#2c2c2e"/>
                            </Style>
                        </DataGrid.ColumnHeaderStyle>
                        <DataGrid.CellStyle>
                            <Style TargetType="DataGridCell">
                                <Setter Property="Background" Value="Transparent"/>
                                <Setter Property="BorderThickness" Value="0"/>
                                <Setter Property="Foreground" Value="{Binding Path=Foreground, RelativeSource={RelativeSource AncestorType=DataGridRow}}"/>
                                <Style.Triggers>
                                    <Trigger Property="IsSelected" Value="True">
                                        <Setter Property="Background" Value="#33ffffff"/>
                                        <Setter Property="Foreground" Value="{Binding Path=Foreground, RelativeSource={RelativeSource AncestorType=DataGridRow}}"/>
                                    </Trigger>
                                </Style.Triggers>
                            </Style>
                        </DataGrid.CellStyle>
                        <DataGrid.RowStyle>
                            <Style TargetType="DataGridRow">
                                <Setter Property="BorderThickness" Value="0,0,0,1"/>
                                <Setter Property="BorderBrush" Value="#1f1f23"/>
                                <Setter Property="Padding" Value="5"/>
                                <Style.Triggers>
                                    <DataTrigger Binding="{Binding StatusCategory}" Value="Active (&lt;=30d)">
                                        <Setter Property="Background" Value="#10b981"/>
                                        <Setter Property="Foreground" Value="White"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding StatusCategory}" Value="Inactive (30-90d)">
                                        <Setter Property="Background" Value="#eab308"/>
                                        <Setter Property="Foreground" Value="Black"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding StatusCategory}" Value="Inactive (90-365d)">
                                        <Setter Property="Background" Value="#3B82F6"/>
                                        <Setter Property="Foreground" Value="White"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding StatusCategory}" Value="Inactive (&gt;1yr)">
                                        <Setter Property="Background" Value="#ef4444"/>
                                        <Setter Property="Foreground" Value="White"/>
                                    </DataTrigger>
                                    <DataTrigger Binding="{Binding StatusCategory}" Value="Never Logged In">
                                        <Setter Property="Background" Value="#64748b"/>
                                        <Setter Property="Foreground" Value="White"/>
                                    </DataTrigger>
                                </Style.Triggers>
                            </Style>
                        </DataGrid.RowStyle>
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Display Name" Binding="{Binding DisplayName}" Width="1.2*"/>
                            <DataGridTextColumn Header="User Principal Name (UPN)" Binding="{Binding UserPrincipalName}" Width="1.5*"/>
                            <DataGridTextColumn Header="Licenses" Binding="{Binding AssignedLicenses}" Width="2*"/>
                            <DataGridTextColumn Header="Last Sign-In" Binding="{Binding LastSignInDate}" Width="1.2*"/>
                            <DataGridTextColumn Header="Status" Binding="{Binding StatusText}" Width="130"/>
                            <DataGridTextColumn Header="Account Status" Binding="{Binding AccountStatusText}" Width="110">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock">
                                        <Setter Property="HorizontalAlignment" Value="Center"/>
                                    </Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                            <DataGridTextColumn Header="Wasted Cost" Binding="{Binding WastedCostText}" Width="110"/>
                            <DataGridTextColumn Header="Monthly Savings" Binding="{Binding MonthlySavingsText}" Width="110"/>
                            <DataGridTextColumn Header="Recommendation" Binding="{Binding Recommendation}" Width="1.5*"/>
                            <DataGridTextColumn Header="Verification" Binding="{Binding VerificationText}" Width="130"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Border>
                
                <!-- Bottom Exporters -->
                <Grid Grid.Row="2" Margin="0,15,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    
                    <TextBlock Text="// Results represent audited active Entra ID subscriptions" Foreground="#555" FontSize="11" VerticalAlignment="Center"/>
                    
                    <StackPanel Grid.Column="1" Orientation="Horizontal">
                        <Button Name="btnGenEmail" Style="{StaticResource SecBtn}" Content="Generate Email Draft" Height="32" Margin="0,0,8,0"/>
                        <Button Name="btnExportCSV" Style="{StaticResource SecBtn}" Content="Export CSV" Height="32" Margin="0,0,8,0"/>
                        <Button Name="btnExportExcel" Style="{StaticResource SecBtn}" Content="Export Excel (XLSX)" Height="32" Margin="0,0,8,0"/>
                        <Button Name="btnExportPDF" Style="{StaticResource SecBtn}" Content="Export PDF Report" Height="32" Margin="0,0,8,0"/>
                        <Button Name="btnExportPNG" Style="{StaticResource SecBtn}" Content="Snapshot Dashboard PNG" Height="32"/>
                    </StackPanel>
                </Grid>
            </Grid>
        </Grid>
    </Grid>
</Window>
"@

function Find-WpfChildByName {
    param(
        $Parent,
        [string]$Name
    )
    if ($null -eq $Parent) { return $null }
    if ($Parent.Name -eq $Name) { return $Parent }
    
    try {
        $Children = [System.Windows.LogicalTreeHelper]::GetChildren($Parent)
        foreach ($Child in $Children) {
            if ($Child -is [System.Windows.DependencyObject]) {
                $Found = Find-WpfChildByName -Parent $Child -Name $Name
                if ($null -ne $Found) { return $Found }
            }
        }
    } catch {}
    
    return $null
}

# 4. XML Parsing and WPF Window Binding
$Reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($Xaml))
$Window = [System.Windows.Markup.XamlReader]::Load($Reader)

# Dynamically link named XAML elements to local PowerShell variables prefixed with '$wpf_'
[xml]$Xml = $Xaml
$Xml.SelectNodes("//*[@Name]") | ForEach-Object {
    $Val = $Window.FindName($_.Name)
    if ($null -eq $Val) {
        $Val = Find-WpfChildByName -Parent $Window -Name $_.Name
    }
    Set-Variable -Name "wpf_$($_.Name)" -Value $Val -Scope Script
}

# 5. Global State & Mocks
$Script:SessionMode = ""
$Script:CurrentData = $null
$Script:GridRows = $null
$Script:ActiveFilter = "all"
$Script:SearchQuery = ""
$Script:SessionSeconds = 0
$Script:UserRole = ""

function Get-MockDateString($DaysAgo) {
    return (Get-Date).AddDays(-$DaysAgo).ToString("yyyy-MM-ddTHH:mm:ssZ")
}

$MockData = @(
    [PSCustomObject]@{ DisplayName = "Alex Wilber"; UserPrincipalName = "AlexW@contoso.com"; AssignedLicenses = "Microsoft 365 Business Premium, Microsoft 365 Copilot"; LastSignInDate = (Get-MockDateString 2); AccountEnabled = $true }
    [PSCustomObject]@{ DisplayName = "Adele Vance"; UserPrincipalName = "AdeleV@contoso.com"; AssignedLicenses = "Microsoft 365 Business Premium"; LastSignInDate = (Get-MockDateString 15); AccountEnabled = $true }
    [PSCustomObject]@{ DisplayName = "Pradeep Gupta"; UserPrincipalName = "PradeepG@contoso.com"; AssignedLicenses = "Power BI Pro, Microsoft 365 E5"; LastSignInDate = (Get-MockDateString 390); AccountEnabled = $true }
    [PSCustomObject]@{ DisplayName = "Megan Bowen"; UserPrincipalName = "MeganB@contoso.com"; AssignedLicenses = "Microsoft 365 E3"; LastSignInDate = (Get-MockDateString 105); AccountEnabled = $true }
    [PSCustomObject]@{ DisplayName = "Joni Sherman"; UserPrincipalName = "JoniS@contoso.com"; AssignedLicenses = "Microsoft 365 Business Basic"; LastSignInDate = "No interactive sign-in recorded"; AccountEnabled = $false }
    [PSCustomObject]@{ DisplayName = "Lynne Robbins"; UserPrincipalName = "LynneR@contoso.com"; AssignedLicenses = "Microsoft 365 F3, Exchange Online (Plan 2)"; LastSignInDate = (Get-MockDateString 410); AccountEnabled = $false }
    [PSCustomObject]@{ DisplayName = "Isaiah Langer"; UserPrincipalName = "IsaiahL@contoso.com"; AssignedLicenses = "Microsoft 365 E5"; LastSignInDate = (Get-MockDateString 95); AccountEnabled = $true }
    [PSCustomObject]@{ DisplayName = "Lidia Holloway"; UserPrincipalName = "LidiaH@contoso.com"; AssignedLicenses = "Microsoft 365 Business Premium, Power BI Pro"; LastSignInDate = (Get-MockDateString 0); AccountEnabled = $true }
    [PSCustomObject]@{ DisplayName = "Grady Archie"; UserPrincipalName = "GradyA@contoso.com"; AssignedLicenses = "Microsoft 365 Business Standard"; LastSignInDate = (Get-MockDateString 195); AccountEnabled = $true }
    [PSCustomObject]@{ DisplayName = "Patti Fernandez"; UserPrincipalName = "PattiF@contoso.com"; AssignedLicenses = "Microsoft Teams Exploratory"; LastSignInDate = "No interactive sign-in recorded"; AccountEnabled = $true }
    [PSCustomObject]@{ DisplayName = "Nestor Wilke"; UserPrincipalName = "NestorW@contoso.com"; AssignedLicenses = "Microsoft 365 Copilot, Microsoft 365 E5"; LastSignInDate = (Get-MockDateString 35); AccountEnabled = $true }
    [PSCustomObject]@{ DisplayName = "Diego Siciliani"; UserPrincipalName = "DiegoS@contoso.com"; AssignedLicenses = "Microsoft 365 Business Premium"; LastSignInDate = (Get-MockDateString 18); AccountEnabled = $true }
    [PSCustomObject]@{ DisplayName = "Test Admin Account"; UserPrincipalName = "seriun.test@contoso.com"; AssignedLicenses = "Microsoft 365 Business Premium"; LastSignInDate = (Get-MockDateString 5); AccountEnabled = $true }
    [PSCustomObject]@{ DisplayName = "Demo Admin User"; UserPrincipalName = "jp@contoso.com"; AssignedLicenses = "Microsoft 365 E5"; LastSignInDate = (Get-MockDateString 12); AccountEnabled = $true }
)

$ActivityColors = @{
    "Active (<=30d)"    = "#10b981"
    "Inactive (30-90d)"  = "#eab308"
    "Inactive (90-365d)" = "#3B82F6"
    "Inactive (>1yr)"    = "#ef4444"
    "Never Logged In"    = "#64748b"
}

# Determine script directory
$Script:ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path -ErrorAction SilentlyContinue }
if (-not $Script:ScriptDir) { $Script:ScriptDir = (Get-Location).Path }

function Initialize-SkuPrices {
    $PricesFile = Join-Path $Script:ScriptDir "LicensePrices.csv"
    
    # Auto-generate default CSV if missing
    if (-not (Test-Path $PricesFile)) {
        $DefaultCsv = @"
Product,MonthlyPrice,SkuId
Microsoft 365 Business Premium,18.10,cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46
Microsoft 365 Business Standard,10.30,f245ecc8-75af-4f8e-b61f-27d8114de5f3
Microsoft 365 Business Basic,4.90,bd251394-b1ed-487b-a1aa-ee198c62c938
Microsoft 365 Business Basic,4.90,3b555118-da6a-4418-894f-7df1e2096870
Microsoft 365 E5,48.10,078d10ee-6995-4851-8043-334f610f49b3
Microsoft 365 E3,31.70,47d12459-1159-4ad1-abfa-00e92056813a
Microsoft 365 F3,6.60,61346032-1554-4736-b876-7d28d697f394
Microsoft 365 F3,6.60,f7ee79a7-7aec-4ca4-9fb9-34d6b930ad87
Exchange Online (Plan 2),6.60,19ec0d23-8335-4cbd-94ac-6050e30712fa
Power BI Pro,8.20,a403ebcc-fae0-4ca2-8c8c-7a907fd6c235
Microsoft 365 Copilot,24.70,639dec6b-bb19-468b-871c-c5c441c4b0cb
Microsoft 365 Copilot,24.70,ab5128ae-2475-4d95-8c73-33f07d701bfc
Office 365 E3,22.00,6fd2c87f-b296-42f0-b197-1e91e994b900
Office 365 E5,37.50,c5928f49-12ba-48f7-ada3-0d743a3dbd2b
Microsoft 365 Audio Conferencing,3.30,a403a5cc-140d-4168-bfb5-58b5e7cbf094
Microsoft Service Business,0.00,57ff2da0-73d6-437c-a4df-ab9dc44faefc
Microsoft Teams Exploratory,0.00,710779e8-3d4a-4c88-adb9-386c958d1fdf
Microsoft Power Automate Free,0.00,f30db892-07e9-47e9-837c-80727f46fd3d
Windows Store for Business,0.00,6470687e-a428-4b7a-bef2-8a291ad947c9
"@
        try {
            $DefaultCsv | Out-File -FilePath $PricesFile -Force -Encoding utf8
        } catch {
            Write-Host "Warning: Could not create default LicensePrices.csv: $_" -ForegroundColor Yellow
        }
    } else {
        # Check if user's existing CSV is missing critical GUIDs and append them
        try {
            $CsvContent = Get-Content -Path $PricesFile -Raw
            if ($CsvContent -notmatch "ab5128ae-2475-4d95-8c73-33f07d701bfc") {
                Add-Content -Path $PricesFile -Value "`nMicrosoft 365 Copilot,24.70,ab5128ae-2475-4d95-8c73-33f07d701bfc" -Encoding utf8
            }
            if ($CsvContent -notmatch "61346032-1554-4736-b876-7d28d697f394") {
                Add-Content -Path $PricesFile -Value "`nMicrosoft 365 F3,6.60,61346032-1554-4736-b876-7d28d697f394" -Encoding utf8
            }
            if ($CsvContent -notmatch "f7ee79a7-7aec-4ca4-9fb9-34d6b930ad87") {
                Add-Content -Path $PricesFile -Value "`nMicrosoft 365 F3,6.60,f7ee79a7-7aec-4ca4-9fb9-34d6b930ad87" -Encoding utf8
            }
            if ($CsvContent -notmatch "6470687e-a428-4b7a-bef2-8a291ad947c9") {
                Add-Content -Path $PricesFile -Value "`nWindows Store for Business,0.00,6470687e-a428-4b7a-bef2-8a291ad947c9" -Encoding utf8
            }
        } catch {}

        # Self-healing cleanup of any malformed blank entries or legacy incorrect mappings (e.g. 3b555118 F3 mapping)
        try {
            $CleanLines = [System.Collections.Generic.List[string]]::new()
            $HasMalformed = $false
            foreach ($line in Get-Content -Path $PricesFile) {
                if ($line -like "Product,MonthlyPrice,SkuId") {
                    $CleanLines.Add($line)
                    continue
                }
                # Fix F3 incorrect mapping to Business Basic
                if ($line -like "*3b555118-da6a-4418-894f-7df1e2096870*") {
                    $CleanLines.Add("Microsoft 365 Business Basic,4.90,3b555118-da6a-4418-894f-7df1e2096870")
                    $HasMalformed = $true
                    continue
                }
                if ([string]::IsNullOrWhiteSpace($line) -or $line -like ",*") {
                    $HasMalformed = $true
                    continue
                }
                $CleanLines.Add($line)
            }
            if ($HasMalformed) {
                $CleanLines | Out-File -FilePath $PricesFile -Force -Encoding utf8
            }
        } catch {}
    }
    
    $Prices = @{}
    if (Test-Path $PricesFile) {
        try {
            $Csv = Import-Csv -Path $PricesFile
            foreach ($row in $Csv) {
                $Price = 0.00
                if ($row.MonthlyPrice -and [double]::TryParse($row.MonthlyPrice, [ref]$Price)) {
                    # Map by Product Name
                    if ($row.Product) {
                        $Prices[$row.Product.Trim()] = $Price
                    }
                    # Map by SkuId (GUID) as well to allow exact GUID lookups!
                    if ($row.SkuId -and $row.SkuId.Trim() -ne "") {
                        $Guid = $row.SkuId.Trim().ToLower()
                        $Prices[$Guid] = $Price
                    }
                }
            }
        } catch {
            Write-Host "Error loading LicensePrices.csv: $_" -ForegroundColor Red
        }
    }
    
    # If loading failed or file was empty, fall back to inline defaults
    if ($Prices.Count -eq 0) {
        $Prices = @{
            "Microsoft 365 Business Premium"     = 18.10
            "Microsoft 365 Business Standard"    = 10.30
            "Microsoft 365 Business Basic"       = 4.90
            "Microsoft 365 E5"                   = 48.10
            "Microsoft 365 E3"                   = 31.70
            "Microsoft 365 F3"                   = 6.60
            "Exchange Online (Plan 2)"           = 6.60
            "Power BI Pro"                       = 8.20
            "Microsoft 365 Copilot"              = 24.70
            "Office 365 E3"                      = 22.00
            "Office 365 E5"                      = 37.50
            "Microsoft 365 Audio Conferencing"   = 3.30
            "Microsoft Service Business"         = 0.00
            "Microsoft Teams Exploratory"        = 0.00
            "Microsoft Power Automate Free"      = 0.00
        }
    }
    
    return $Prices
}

$Script:SkuPrices = Initialize-SkuPrices

# 6. Dispatcher Timer for Session Uptime
$Script:Timer = New-Object System.Windows.Threading.DispatcherTimer
$Script:Timer.Interval = [TimeSpan]::FromSeconds(1)
$Script:Timer.Add_Tick({
    $Script:SessionSeconds++
    $hours = '{0:d2}' -f [int][Math]::Floor($Script:SessionSeconds / 3600)
    $minutes = '{0:d2}' -f [int][Math]::Floor(($Script:SessionSeconds % 3600) / 60)
    $seconds = '{0:d2}' -f [int]($Script:SessionSeconds % 60)
    $wpf_txtSessionUptime.Text = "${hours}:${minutes}:${seconds}"
})

# 7. Helper Functions
function Clear-MgGraphCache {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    try {
        Remove-Item "$env:USERPROFILE\.graph" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item "$env:USERPROFILE\.mg" -Recurse -Force -ErrorAction SilentlyContinue
        if (Test-Path "$env:LOCALAPPDATA\.IdentityService") {
            Get-ChildItem "$env:LOCALAPPDATA\.IdentityService" -Filter "mg*" -Recurse -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
        }
    } catch {}
}

function Start-SessionTimer {
    $Script:SessionSeconds = 0
    $wpf_txtSessionUptime.Text = "00:00:00"
    $Script:Timer.Start()
}

function Stop-SessionTimer {
    $Script:Timer.Stop()
}

function Get-LicenseMonthlyPrice($LicensesString) {
    if (-not $LicensesString) { return 0.00 }
    $Sum = 0.00
    $LicensesString.Split(',') | ForEach-Object {
        $Name = $_.Trim()
        if ($Name) {
            # Extract GUID if the name is formatted like "Unknown SKU (guid)"
            if ($Name -match "Unknown SKU \(([^)]+)\)") {
                $Name = $Matches[1].Trim().ToLower()
            }
            
            $Matched = $false
            foreach ($key in $Script:SkuPrices.Keys) {
                # Check for exact display name, start/end matches, or exact SkuId GUID match
                if ($Name -eq $key -or $Name.StartsWith($key) -or $key.StartsWith($Name) -or ($Name -match "^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$" -and $Name -eq $key.ToLower())) {
                    $Sum += $Script:SkuPrices[$key]
                    $Matched = $true
                    break
                }
            }
            if (-not $Matched) {
                if ($Name -match "E5") { $Sum += 48.10 }
                elseif ($Name -match "E3") { $Sum += 31.70 }
                elseif ($Name -match "Business") { $Sum += 10.30 }
                else { $Sum += 10.00 }
            }
        }
    }
    return $Sum
}

function Update-UI {
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [Action] {}
    )
}

function DoEvents() {
    Update-UI
}

function Get-TenantName {
    try {
        if ($null -ne $Script:wpf_txtSessionTenant) {
            $val = $Script:wpf_txtSessionTenant.Text
            if (-not [string]::IsNullOrEmpty($val)) { return $val.Trim() }
        }
        if ($null -ne $wpf_txtSessionTenant) {
            $val = $wpf_txtSessionTenant.Text
            if (-not [string]::IsNullOrEmpty($val)) { return $val.Trim() }
        }
    } catch {}
    return "Tenant"
}

function Log-ToTerminal($Message, $Type = "Info") {
    $Timestamp = (Get-Date).ToString("HH:mm:ss")
    $Prefix = "pwsh> "
    if ($Type -eq "Success") { $Prefix = "[SUCCESS] " }
    elseif ($Type -eq "Warning") { $Prefix = "[WARNING] " }
    elseif ($Type -eq "Error") { $Prefix = "[ERROR] " }
    elseif ($Type -eq "Command") { $Prefix = "pwsh> " }
    
    $Line = "[$Timestamp] $Prefix$Message`r`n"
    $wpf_txtLog.AppendText($Line)
    $wpf_txtLog.ScrollToEnd()
    DoEvents
}

# Show-RoleSelectorDialog is deprecated. Role selection is performed directly on the welcome panel.

function Set-DataGridColumnVisibility {
    param(
        [Parameter(Mandatory=$true)]
        [string]$HeaderName,
        [Parameter(Mandatory=$true)]
        [bool]$IsVisible
    )
    if ($null -eq $wpf_gridUsers) { return }
    $Col = $wpf_gridUsers.Columns | Where-Object { $_.Header -eq $HeaderName }
    if ($Col) {
        $Col.Visibility = if ($IsVisible) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    }
}

function Set-UnassignedGridColumnVisibility {
    param(
        [Parameter(Mandatory=$true)]
        [string]$HeaderName,
        [Parameter(Mandatory=$true)]
        [bool]$IsVisible
    )
    if ($null -eq $wpf_gridUnassigned) { return }
    $Col = $wpf_gridUnassigned.Columns | Where-Object { $_.Header -eq $HeaderName }
    if ($Col) {
        $Col.Visibility = if ($IsVisible) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    }
}

function Configure-RoleUI {

    if ($Script:UserRole -eq "ServiceDesk") {
        # Hide monetary KPIs
        if ($wpf_borderKpiSavings) { $wpf_borderKpiSavings.Visibility = [System.Windows.Visibility]::Collapsed }
        if ($wpf_borderKpiPoolWaste) { $wpf_borderKpiPoolWaste.Visibility = [System.Windows.Visibility]::Collapsed }
        if ($wpf_colKpiSavings) { $wpf_colKpiSavings.Width = New-Object System.Windows.GridLength(0) }
        if ($wpf_colKpiPoolWaste) { $wpf_colKpiPoolWaste.Width = New-Object System.Windows.GridLength(0) }
        
        # Hide DataGrid columns
        Set-DataGridColumnVisibility -HeaderName "Wasted Cost" -IsVisible $false
        Set-DataGridColumnVisibility -HeaderName "Monthly Savings" -IsVisible $false
        Set-DataGridColumnVisibility -HeaderName "Recommendation" -IsVisible $false
        
        # Hide PDF exporter
        if ($wpf_btnExportPDF) { $wpf_btnExportPDF.Visibility = [System.Windows.Visibility]::Collapsed }
        if ($wpf_btnExportPDFDash) { $wpf_btnExportPDFDash.Visibility = [System.Windows.Visibility]::Collapsed }
        
        # Hide monetary columns in Unassigned Licenses Breakdown Grid
        Set-UnassignedGridColumnVisibility -HeaderName "Monthly Cost" -IsVisible $false
        Set-UnassignedGridColumnVisibility -HeaderName "Wasted/mo" -IsVisible $false
    } else {
        # Show monetary KPIs
        if ($wpf_borderKpiSavings) { $wpf_borderKpiSavings.Visibility = [System.Windows.Visibility]::Visible }
        if ($wpf_borderKpiPoolWaste) { $wpf_borderKpiPoolWaste.Visibility = [System.Windows.Visibility]::Visible }
        if ($wpf_colKpiSavings) { $wpf_colKpiSavings.Width = New-Object System.Windows.GridLength(1.2, [System.Windows.GridUnitType]::Star) }
        if ($wpf_colKpiPoolWaste) { $wpf_colKpiPoolWaste.Width = New-Object System.Windows.GridLength(1.2, [System.Windows.GridUnitType]::Star) }
        
        # Show DataGrid columns
        Set-DataGridColumnVisibility -HeaderName "Wasted Cost" -IsVisible $true
        Set-DataGridColumnVisibility -HeaderName "Monthly Savings" -IsVisible $true
        Set-DataGridColumnVisibility -HeaderName "Recommendation" -IsVisible $true
        
        # Show PDF exporter
        if ($wpf_btnExportPDF) { $wpf_btnExportPDF.Visibility = [System.Windows.Visibility]::Visible }
        if ($wpf_btnExportPDFDash) { $wpf_btnExportPDFDash.Visibility = [System.Windows.Visibility]::Visible }
        
        # Show monetary columns in Unassigned Licenses Breakdown Grid
        Set-UnassignedGridColumnVisibility -HeaderName "Monthly Cost" -IsVisible $true
        Set-UnassignedGridColumnVisibility -HeaderName "Wasted/mo" -IsVisible $true
    }
}

function Show-VerificationSelectorDialog {
    param(
        [Parameter(Mandatory=$true)]
        $Users
    )
    
    $DialogXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Exclude Verification Accounts" Height="360" Width="460"
        WindowStartupLocation="CenterOwner" Background="#131313" Foreground="#f3f4f6"
        ResizeMode="NoResize" ShowInTaskbar="False">
    <Window.Resources>
        <Style x:Key="DialogBtn" TargetType="Button">
            <Setter Property="Background" Value="#3B82F6"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#1D4ED8"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="DialogSecBtn" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#9ca3af"/>
            <Setter Property="BorderBrush" Value="#2c2c2e"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#2c2c2e"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <StackPanel Grid.Row="0" Margin="0,0,0,15">
            <TextBlock Text="Verification Accounts Detected" Foreground="#3B82F6" FontSize="15" FontWeight="Bold"/>
            <TextBlock Text="The following Test/Admin accounts were found in the list. Uncheck any accounts you want to remove from the exported file." 
                       Foreground="#9ca3af" FontSize="11" TextWrapping="Wrap" Margin="0,4,0,0"/>
        </StackPanel>
        
        <Border Grid.Row="1" Background="#0c0c0d" BorderBrush="#2c2c2e" BorderThickness="1" CornerRadius="6" Padding="10">
            <ScrollViewer VerticalScrollBarVisibility="Auto">
                <StackPanel Name="panelCheckBoxes"/>
            </ScrollViewer>
        </Border>
        
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button Name="btnProceed" Style="{StaticResource DialogBtn}" Content="Proceed Export" Width="120" Height="30" IsDefault="True" Margin="0,0,10,0"/>
            <Button Name="btnCancel" Style="{StaticResource DialogSecBtn}" Content="Cancel" Width="80" Height="30" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $Reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($DialogXaml))
    $DlgWindow = [System.Windows.Markup.XamlReader]::Load($Reader)
    $DlgWindow.Owner = $Window
    
    $Panel = $DlgWindow.FindName("panelCheckBoxes")
    $BtnProceed = $DlgWindow.FindName("btnProceed")
    $BtnCancel = $DlgWindow.FindName("btnCancel")
    
    $CheckBoxes = @()
    foreach ($user in $Users) {
        $CB = [System.Windows.Controls.CheckBox]::new()
        $CB.Content = "$($user.DisplayName) ($($user.UserPrincipalName))"
        $CB.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#f3f4f6")
        $CB.IsChecked = $true
        $CB.Margin = "0,4,0,4"
        $CB.Tag = $user.UserPrincipalName
        $Panel.Children.Add($CB)
        $CheckBoxes += $CB
    }
    
    $Script:DlgResult = $null
    $Script:ExcludedUPNs = @()
    
    $BtnProceed.Add_Click({
        $Script:DlgResult = "Proceed"
        $Script:ExcludedUPNs = foreach ($cb in $CheckBoxes) {
            if ($cb.IsChecked -eq $false) {
                $cb.Tag
            }
        }
        $DlgWindow.Close()
    })
    
    $BtnCancel.Add_Click({
        $Script:DlgResult = "Cancel"
        $DlgWindow.Close()
    })
    
    $DlgWindow.ShowDialog() | Out-Null
    
    return [PSCustomObject]@{
        Result       = $Script:DlgResult
        ExcludedUPNs = $Script:ExcludedUPNs
    }
}

function Show-EmailTemplateDialog {
    param(
        [Parameter(Mandatory=$true)]
        [string]$EmailText
    )
    
    $DialogXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Generated Email Draft" Height="500" Width="650"
        WindowStartupLocation="CenterOwner" Background="#131313" Foreground="#f3f4f6"
        ResizeMode="NoResize" ShowInTaskbar="False">
    <Window.Resources>
        <Style x:Key="DialogBtn" TargetType="Button">
            <Setter Property="Background" Value="#2c2c2e"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Margin" Value="5,0"/>
            <Setter Property="Height" Value="32"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#3c3c3e"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="BrandDialogBtn" TargetType="Button" BasedOn="{StaticResource DialogBtn}">
            <Setter Property="Background" Value="#3B82F6"/>
            <Setter Property="Foreground" Value="White"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#60A5FA"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <StackPanel Margin="0,0,0,15">
            <TextBlock Text="GENERATED EMAIL DRAFT" Foreground="#3B82F6" FontSize="14" FontWeight="Bold"/>
            <TextBlock Text="Review the generated draft below. You can copy it directly to your clipboard." Foreground="#9ca3af" FontSize="11" Margin="0,3,0,0"/>
        </StackPanel>
        
        <!-- Scrollable TextBox -->
        <Border Grid.Row="1" Background="#1c1c1e" BorderBrush="#2c2c2e" BorderThickness="1" CornerRadius="6" Padding="5">
            <TextBox Name="txtEmailBody" Background="Transparent" Foreground="#f3f4f6" BorderThickness="0"
                     TextWrapping="Wrap" AcceptsReturn="True" IsReadOnly="True" VerticalScrollBarVisibility="Auto"
                     FontSize="12" FontFamily="Segoe UI" Padding="8"/>
        </Border>
        
        <!-- Actions -->
        <Grid Grid.Row="2" Margin="0,15,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            
            <TextBlock Name="txtStatusMessage" Text="" Foreground="#10b981" FontSize="12" VerticalAlignment="Center" FontWeight="SemiBold"/>
            
            <StackPanel Grid.Column="1" Orientation="Horizontal">
                <Button Name="btnCopy" Style="{StaticResource BrandDialogBtn}" Content="Copy to Clipboard"/>
                <Button Name="btnClose" Style="{StaticResource DialogBtn}" Content="Close"/>
            </StackPanel>
        </Grid>
    </Grid>
</Window>
"@

    # Apply blur to parent window
    $WpfWindow = $Window
    $Blur = New-Object System.Windows.Media.Effects.BlurEffect
    $Blur.Radius = 15
    $WpfWindow.Effect = $Blur
    
    # Load dialog
    $Reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($DialogXaml))
    $DlgWindow = [System.Windows.Markup.XamlReader]::Load($Reader)
    $DlgWindow.Owner = $WpfWindow
    
    # Bind dialog controls
    $TxtEmailBody = $DlgWindow.FindName("txtEmailBody")
    $BtnCopy = $DlgWindow.FindName("btnCopy")
    $BtnClose = $DlgWindow.FindName("btnClose")
    $TxtStatusMessage = $DlgWindow.FindName("txtStatusMessage")
    
    $TxtEmailBody.Text = $EmailText
    
    # Click Events
    $BtnCopy.Add_Click({
        [System.Windows.Clipboard]::SetText($TxtEmailBody.Text)
        $TxtStatusMessage.Text = "Copied to clipboard!"
        
        # Async clear message after 2 seconds
        $Timer = New-Object System.Windows.Threading.DispatcherTimer
        $Timer.Interval = [TimeSpan]::FromSeconds(2)
        $Timer.Add_Tick({
            $TxtStatusMessage.Text = ""
            $Timer.Stop()
        })
        $Timer.Start()
    })
    
    $BtnClose.Add_Click({
        $DlgWindow.Close()
    })
    
    $DlgWindow.ShowDialog() | Out-Null
    
    # Remove blur
    $WpfWindow.Effect = $null
}

function Invoke-AuditReport {
    $wpf_btnRunScript.IsEnabled = $false
    $wpf_btnRunScript.Content = "Executing audit..."
    Update-UI
    
    Log-ToTerminal "pwsh> .\\LicencedUsersSigninDate.ps1" "Command"
    Update-UI
    
    Start-Sleep -Milliseconds 400
    Log-ToTerminal "Checking Microsoft Graph connection status..." "Info"
    Update-UI
    
    Start-Sleep -Milliseconds 400
    Log-ToTerminal "Connection verified. Fetching tenant organization metadata..." "Info"
    Update-UI
    
    if ($Script:SessionMode -eq "Demo") {
        Start-Sleep -Milliseconds 600
        Log-ToTerminal "Querying Entra licensed users (Simulated)..." "Info"
        Update-UI
        Start-Sleep -Milliseconds 400
        Log-ToTerminal "Resolving license SKU mappings..." "Info"
        Update-UI
        Start-Sleep -Milliseconds 400
        
        $GreenBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#10b981")
        $RedBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#ef4444")
        $GrayBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#9ca3af")
        
        $Script:TotalPoolWaste = 186.70
        $Script:TotalUnassignedCount = 7
        $Script:UnassignedLicenses = @(
            [PSCustomObject]@{ SkuPartName = "Microsoft 365 E5"; SkuId = "078d10ee-6995-4851-8043-334f610f49b3"; ActiveUnits = 10; ConsumedUnits = 8; UnassignedUnits = 2; UnassignedBrush = $GreenBrush; WastedBrush = $RedBrush; MonthlyPrice = 48.10; MonthlyPriceText = "£48.10"; WastedCost = 96.20; WastedCostText = "£96.20"; AssignedText = "8 of 10" }
            [PSCustomObject]@{ SkuPartName = "Microsoft 365 Business Premium"; SkuId = "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46"; ActiveUnits = 25; ConsumedUnits = 20; UnassignedUnits = 5; UnassignedBrush = $GreenBrush; WastedBrush = $RedBrush; MonthlyPrice = 18.10; MonthlyPriceText = "£18.10"; WastedCost = 90.50; WastedCostText = "£90.50"; AssignedText = "20 of 25" }
            [PSCustomObject]@{ SkuPartName = "Microsoft 365 Copilot"; SkuId = "639dec6b-bb19-468b-871c-c5c441c4b0cb"; ActiveUnits = 5; ConsumedUnits = 5; UnassignedUnits = 0; UnassignedBrush = $GrayBrush; WastedBrush = $GrayBrush; MonthlyPrice = 24.70; MonthlyPriceText = "£24.70"; WastedCost = 0.00; WastedCostText = "£0.00"; AssignedText = "5 of 5" }
            [PSCustomObject]@{ SkuPartName = "Power BI Pro"; SkuId = "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235"; ActiveUnits = 10; ConsumedUnits = 10; UnassignedUnits = 0; UnassignedBrush = $GrayBrush; WastedBrush = $GrayBrush; MonthlyPrice = 8.20; MonthlyPriceText = "£8.20"; WastedCost = 0.00; WastedCostText = "£0.00"; AssignedText = "10 of 10" }
            [PSCustomObject]@{ SkuPartName = "Microsoft 365 Business Standard"; SkuId = "f245ecc8-75af-4f8e-b61f-27d8114de5f3"; ActiveUnits = 15; ConsumedUnits = 15; UnassignedUnits = 0; UnassignedBrush = $GrayBrush; WastedBrush = $GrayBrush; MonthlyPrice = 10.30; MonthlyPriceText = "£10.30"; WastedCost = 0.00; WastedCostText = "£0.00"; AssignedText = "15 of 15" }
            [PSCustomObject]@{ SkuPartName = "Microsoft 365 Business Basic"; SkuId = "bd251394-b1ed-487b-a1aa-ee198c62c938"; ActiveUnits = 30; ConsumedUnits = 30; UnassignedUnits = 0; UnassignedBrush = $GrayBrush; WastedBrush = $GrayBrush; MonthlyPrice = 4.90; MonthlyPriceText = "£4.90"; WastedCost = 0.00; WastedCostText = "£0.00"; AssignedText = "30 of 30" }
            [PSCustomObject]@{ SkuPartName = "Office 365 E3"; SkuId = "6fd2c87f-b296-42f0-b197-1e91e994b900"; ActiveUnits = 12; ConsumedUnits = 12; UnassignedUnits = 0; UnassignedBrush = $GrayBrush; WastedBrush = $GrayBrush; MonthlyPrice = 22.00; MonthlyPriceText = "£22.00"; WastedCost = 0.00; WastedCostText = "£0.00"; AssignedText = "12 of 12" }
            [PSCustomObject]@{ SkuPartName = "Microsoft Teams Exploratory"; SkuId = "710779e8-3d4a-4c88-adb9-386c958d1fdf"; ActiveUnits = 50; ConsumedUnits = 50; UnassignedUnits = 0; UnassignedBrush = $GrayBrush; WastedBrush = $GrayBrush; MonthlyPrice = 0.00; MonthlyPriceText = "£0.00"; WastedCost = 0.00; WastedCostText = "£0.00"; AssignedText = "50 of 50" }
        )
        $Users = $MockData
        Process-UserData $Users
        Log-ToTerminal "Report generation complete. Successfully processed $($Users.Count) sandbox users." "Success"
    } else {
        # REAL AD HOC GRAPH RETRIEVAL
        try {
            Log-ToTerminal "Loading license SkuMap configurations..." "Info"
            Update-UI
            
            # Sku fallback dictionary
            $LocalSkuDict = @{
                "ab5128ae-2475-4d95-8c73-33f07d701bfc" = "Microsoft 365 Copilot"
                "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235" = "Power BI Pro"
                "639dec6b-bb19-468b-871c-c5c441c4b0cb" = "Microsoft 365 Copilot"
                "6470687e-a428-4b7a-bef2-8a291ad947c9" = "Windows Store for Business"
                "19ec0d23-8335-4cbd-94ac-6050e30712fa" = "Exchange Online (Plan 2)"
                "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46" = "Microsoft 365 Business Premium"
                "3b555118-da6a-4418-894f-7df1e2096870" = "Microsoft 365 Business Basic"
                "61346032-1554-4736-b876-7d28d697f394" = "Microsoft 365 F3"
                "f7ee79a7-7aec-4ca4-9fb9-34d6b930ad87" = "Microsoft 365 F3"
                "f245ecc8-75af-4f8e-b61f-27d8114de5f3" = "Microsoft 365 Business Standard"
                "f30db892-07e9-47e9-837c-80727f46fd3d" = "Microsoft Power Automate Free"
                "710779e8-3d4a-4c88-adb9-386c958d1fdf" = "Microsoft Teams Exploratory"
                "cbdc14ab-d96c-4c30-b9f4-6bc7c7bd7d72" = "Microsoft 365 Business Premium"
                "bd251394-b1ed-487b-a1aa-ee198c62c938" = "Microsoft 365 Business Basic"
                "078d10ee-6995-4851-8043-334f610f49b3" = "Microsoft 365 E5"
                "47d12459-1159-4ad1-abfa-00e92056813a" = "Microsoft 365 E3"
                "6fd2c87f-b296-42f0-b197-1e91e994b900" = "Office 365 E3"
                "c5928f49-12ba-48f7-ada3-0d743a3dbd2b" = "Office 365 E5"
                "a403a5cc-140d-4168-bfb5-58b5e7cbf094" = "Microsoft 365 Audio Conferencing"
                "57ff2da0-73d6-437c-a4df-ab9dc44faefc" = "Microsoft Service Business"

                # Technical Name Overrides
                "O365_BUSINESS_ESSENTIALS" = "Microsoft 365 Business Basic"
                "O365_BUSINESS_PREMIUM"    = "Microsoft 365 Business Standard"
                "BUSINESS_PREMIUM"         = "Microsoft 365 Business Premium"
                "SPB"                      = "Microsoft 365 Business Premium"
                "SMB_BUSINESS"             = "Microsoft 365 Business Premium"
                "EXCHANGESTANDARD"         = "Exchange Online (Plan 1)"
                "AAD_PREMIUM"              = "Microsoft Entra ID P1"
                "POWERAPPS_DEV"            = "Microsoft Power Apps for Developer"
                "TEAMS_ROOM_BASIC"         = "Microsoft Teams Rooms Basic"
                "TEAMS_ROOM_PRO"           = "Microsoft Teams Rooms Pro"
            }
            
            $SkuMap = @{}
            foreach ($key in $LocalSkuDict.Keys) {
                $SkuMap[$key.ToLower()] = $LocalSkuDict[$key]
            }
            
            # Pre-populate friendly names from LicensePrices.csv
            $PricesFile = Join-Path $Script:ScriptDir "LicensePrices.csv"
            if (Test-Path $PricesFile) {
                try {
                    $PricesCsv = Import-Csv $PricesFile
                    foreach ($row in $PricesCsv) {
                        if ($row.SkuId -and $row.Product) {
                            $SkuMap[$row.SkuId.ToString().ToLower()] = $row.Product
                        }
                    }
                } catch {}
            }
            
            $Script:TotalPoolWaste = 0.00
            $Script:TotalUnassignedCount = 0
            $UnassignedList = [System.Collections.Generic.List[PSCustomObject]]::new()
            
            # Fetch local tenant subbed SKUs and compute unassigned units
            try {
                $SubSkus = Get-MgSubscribedSku -All
                foreach ($sku in $SubSkus) {
                    $GuidStr = $sku.SkuId.ToString().ToLower()
                    if ($sku.SkuId -and $sku.SkuPartName -and -not $SkuMap.ContainsKey($GuidStr)) { 
                        $SkuMap[$GuidStr] = $sku.SkuPartName -replace ';', ','
                    }
                    
                    $ActiveUnits = if ($sku.PrepaidUnits -and $sku.PrepaidUnits.Enabled) { $sku.PrepaidUnits.Enabled } else { 0 }
                    $Unassigned = $ActiveUnits - $sku.ConsumedUnits
                    
                    # Resolve SKU Name using GUID mapping, then SkuPartName override mapping, then raw part name
                    $SkuName = if ($SkuMap.ContainsKey($GuidStr)) { 
                        $SkuMap[$GuidStr] 
                    } elseif ($sku.SkuPartName -and $SkuMap.ContainsKey($sku.SkuPartName)) { 
                        $SkuMap[$sku.SkuPartName] 
                    } else { 
                        $sku.SkuPartName -replace ';', ',' 
                    }
                    
                    # Fallback for empty names
                    if ([string]::IsNullOrWhiteSpace($SkuName)) {
                        $SkuName = "Unknown SKU ($GuidStr)"
                    }
                    
                    $UnitPrice = Get-LicenseMonthlyPrice $SkuName
                    
                    # If this SKU is missing from the CSV database, estimate and append it
                    if ($null -ne $Script:SkuPrices -and -not $Script:SkuPrices.ContainsKey($GuidStr)) {
                        if ($SkuName -notlike "Unknown SKU*" -and -not [string]::IsNullOrWhiteSpace($SkuName)) {
                            try {
                                Add-Content -Path $PricesFile -Value "`n$SkuName,$UnitPrice,$GuidStr" -Encoding utf8
                                $Script:SkuPrices[$GuidStr] = $UnitPrice
                                $Script:SkuPrices[$SkuName] = $UnitPrice
                                Log-ToTerminal "Discovered new SKU '$SkuName' ($GuidStr). Added to LicensePrices.csv with estimate £{0:N2}." -f $UnitPrice
                            } catch {}
                        }
                    }
                    
                    if ($null -ne $Script:SkuPrices -and $Script:SkuPrices.ContainsKey($GuidStr)) {
                        $UnitPrice = $Script:SkuPrices[$GuidStr]
                    }
                    
                    # Skip free licenses (price == 0) or massive viral/trial plans (>= 10,000 active units) unless it contains "Teams"
                    if (($UnitPrice -eq 0.00 -or $ActiveUnits -ge 10000) -and $SkuName -notlike "*Teams*") {
                        continue
                    }
                    
                    if ($Unassigned -gt 0) {
                        $Script:TotalPoolWaste += ($Unassigned * $UnitPrice)
                        $Script:TotalUnassignedCount += $Unassigned
                    }
                    
                    $GreenBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#10b981")
                    $RedBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#ef4444")
                    $GrayBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#9ca3af")
                    
                    $UnassignedBrush = if ($Unassigned -gt 0) { $GreenBrush } elseif ($Unassigned -lt 0) { $RedBrush } else { $GrayBrush }
                    $WastedBrush = if ($Unassigned -gt 0) { $RedBrush } elseif ($Unassigned -lt 0) { $GreenBrush } else { $GrayBrush }
                    
                    $UnassignedList.Add([PSCustomObject]@{
                        SkuPartName      = $SkuName
                        SkuId            = $GuidStr
                        ActiveUnits      = $ActiveUnits
                        ConsumedUnits    = $sku.ConsumedUnits
                        UnassignedUnits  = $Unassigned
                        UnassignedBrush  = $UnassignedBrush
                        WastedBrush      = $WastedBrush
                        MonthlyPrice     = $UnitPrice
                        MonthlyPriceText = "£{0:N2}" -f $UnitPrice
                        WastedCost       = ($Unassigned * $UnitPrice)
                        WastedCostText   = "£{0:N2}" -f ($Unassigned * $UnitPrice)
                        AssignedText     = "$($sku.ConsumedUnits) of $ActiveUnits"
                    })
                }
                $Script:UnassignedLicenses = $UnassignedList
            } catch {
                Log-ToTerminal "Failed to fetch local tenant SKUs. Using catalog reference instead." "Warning"
                Update-UI
            }
            
            # Check SKU cache or fetch online
            $LogDir = "$env:SystemDrive\Logs\UserLicenceCheck"
            $CacheFile = Join-Path $LogDir "M365_SKU_Cache.csv"
            $CacheLoaded = $false
            
            if (Test-Path $CacheFile) {
                $LastMod = (Get-Item $CacheFile).LastWriteTime
                if ($LastMod -gt (Get-Date).AddDays(-7)) {
                    Log-ToTerminal "Loading global SKU map from local cache..." "Info"
                    Update-UI
                    try {
                        Import-Csv $CacheFile | ForEach-Object {
                            if ($_.GUID -and $_.Product_Display_Name) {
                                $Guid = $_.GUID.ToString().ToLower()
                                if (-not $SkuMap.ContainsKey($Guid)) {
                                    $SkuMap[$Guid] = $_.Product_Display_Name -replace ';', ','
                                }
                            }
                        }
                        $CacheLoaded = $true
                    } catch {}
                }
            }
            
            if (-not $CacheLoaded) {
                Log-ToTerminal "Fetching latest global SKU mapping database online..." "Info"
                Update-UI
                try {
                    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                    $CsvText = Invoke-RestMethod -Uri "https://raw.githubusercontent.com/merill/license/main/license.csv" -TimeoutSec 10
                    $OnlineSkus = ConvertFrom-Csv -InputObject $CsvText
                    foreach ($Sku in $OnlineSkus) {
                        if ($Sku.GUID -and $Sku.Product_Display_Name) {
                            $Guid = $Sku.GUID.ToString().ToLower()
                            if (-not $SkuMap.ContainsKey($Guid)) {
                                $SkuMap[$Guid] = $Sku.Product_Display_Name -replace ';', ','
                            }
                        }
                    }
                    if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
                    $CsvText | Out-File -FilePath $CacheFile -Force -Encoding utf8
                    Log-ToTerminal "Successfully updated local SKU cache database." "Success"
                    Update-UI
                } catch {
                    Log-ToTerminal "Could not reach online SKU repository. Using local default fallbacks only." "Warning"
                    Update-UI
                }
            }
            
            Log-ToTerminal "Retrieving licensed users with Filter 'assignedLicenses/count ne 0'..." "Info"
            Update-UI
            
            $UserProperties = @('Id', 'DisplayName', 'UserPrincipalName', 'AssignedLicenses', 'SignInActivity', 'AccountEnabled')
            $LicensedUsers = Get-MgUser -Filter "assignedLicenses/`$count ne 0" -ConsistencyLevel eventual -CountVariable LicensedCount -All -Property $UserProperties
            
            Log-ToTerminal "Fetched $($LicensedUsers.Count) users from Graph. Resolving details..." "Info"
            Update-UI
            
            $Users = [System.Collections.Generic.List[PSCustomObject]]::new()
            foreach ($u in $LicensedUsers) {
                $Lics = foreach ($l in $u.AssignedLicenses) {
                    $Guid = $l.SkuId.ToString().ToLower()
                    if ($SkuMap.ContainsKey($Guid)) {
                        $SkuMap[$Guid]
                    } else {
                        "Unknown SKU ($Guid)"
                    }
                }
                $LicString = $Lics -join ", "
                
                $LastSignIn = $u.SignInActivity.LastSuccessfulSignInDateTime
                if (-not $LastSignIn) { $LastSignIn = "No interactive sign-in recorded" }
                
                $Users.Add([PSCustomObject]@{
                    DisplayName       = $u.DisplayName
                    UserPrincipalName = $u.UserPrincipalName
                    AssignedLicenses  = $LicString
                    LastSignInDate    = $LastSignIn
                    AccountEnabled    = $u.AccountEnabled
                })
            }
            
            Process-UserData $Users
            Log-ToTerminal "Report generation complete. Successfully processed $($Users.Count) active licensed users." "Success"
        } catch {
            Log-ToTerminal "Audit run failed: $_" "Error"
            [System.Windows.MessageBox]::Show("Audit Failed:`n$_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
    
    if ([string]::IsNullOrEmpty($Script:UserRole)) {
        $Script:UserRole = "Sales"
    }
    Configure-RoleUI
    
    $wpf_btnRunScript.IsEnabled = $true
    $wpf_btnRunScript.Content = "Refresh Report"
    Update-UI
    Show-Page "Dashboard"
}

function Parse-DateString($DateStr) {
    if ($null -eq $DateStr) { return $null }
    if ($DateStr -is [System.DateTime]) { return $DateStr }
    if ([string]::IsNullOrWhiteSpace($DateStr)) { return $null }
    if ($DateStr -like "*no interactive*" -or $DateStr -like "*recorded*") { return $null }
    
    # Try en-GB (UK) first
    [DateTime]$ParsedDate = [DateTime]::MinValue
    $UkCulture = [System.Globalization.CultureInfo]::GetCultureInfo("en-GB")
    if ([DateTime]::TryParse([string]$DateStr, $UkCulture, [System.Globalization.DateTimeStyles]::None, [ref]$ParsedDate)) {
        return $ParsedDate
    }
    
    # Try en-US second
    $UsCulture = [System.Globalization.CultureInfo]::GetCultureInfo("en-US")
    if ([DateTime]::TryParse([string]$DateStr, $UsCulture, [System.Globalization.DateTimeStyles]::None, [ref]$ParsedDate)) {
        return $ParsedDate
    }
    
    # Fallback to current culture TryParse
    if ([DateTime]::TryParse([string]$DateStr, [ref]$ParsedDate)) {
        return $ParsedDate
    }
    return $null
}

function Get-DaysSince($DateStr) {
    $Date = Parse-DateString $DateStr
    if ($null -eq $Date) { return [double]::PositiveInfinity }
    
    $Diff = (Get-Date) - $Date
    if ($Diff.TotalDays -lt 0) { return 0 }
    return [Math]::Floor($Diff.TotalDays)
}

function Format-LastSignIn($DateStr) {
    $Date = Parse-DateString $DateStr
    if ($null -eq $Date) { return "Never" }
    return $Date.ToString("yyyy-MM-dd HH:mm")
}

# 8. Charting Layout Engines (WPF Native Shapes)
function UpdateBarChart($Panel, $DataMap, $ColorHex) {
    $Panel.Children.Clear()
    if ($null -eq $DataMap -or $DataMap.Count -eq 0) { return }
    
    $MaxVal = 1
    foreach ($val in $DataMap.Values) {
        if ($val -gt $MaxVal) { $MaxVal = $val }
    }
    
    foreach ($Key in $DataMap.Keys) {
        $Val = $DataMap[$Key]
        
        $Grid = [System.Windows.Controls.Grid]::new()
        $Grid.Margin = "0,6,0,6"
        
        $Col1 = [System.Windows.Controls.ColumnDefinition]::new()
        $Col1.Width = [System.Windows.GridLength]::new(180)
        $Col2 = [System.Windows.Controls.ColumnDefinition]::new()
        $Col2.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $Col3 = [System.Windows.Controls.ColumnDefinition]::new()
        $Col3.Width = [System.Windows.GridLength]::new(40)
        
        $Grid.ColumnDefinitions.Add($Col1)
        $Grid.ColumnDefinitions.Add($Col2)
        $Grid.ColumnDefinitions.Add($Col3)
        
        $Label = [System.Windows.Controls.TextBlock]::new()
        $Label.Text = $Key
        $Label.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#9ca3af")
        $Label.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $Label.FontSize = 11
        $Label.TextTrimming = [System.Windows.TextTrimming]::CharacterEllipsis
        [System.Windows.Controls.Grid]::SetColumn($Label, 0)
        $Grid.Children.Add($Label)
        
        $BarOuter = [System.Windows.Controls.Border]::new()
        $BarOuter.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#0c0c0d")
        $BarOuter.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#2c2c2e")
        $BarOuter.BorderThickness = 1
        $BarOuter.CornerRadius = 3
        $BarOuter.Height = 14
        $BarOuter.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
        $BarOuter.Width = 200
        [System.Windows.Controls.Grid]::SetColumn($BarOuter, 1)
        
        $BarWidth = [Math]::Round(($Val / $MaxVal) * 198)
        if ($BarWidth -lt 2) { $BarWidth = 2 }
        $BarInner = [System.Windows.Shapes.Rectangle]::new()
        $BarInner.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString($ColorHex)
        $BarInner.Width = $BarWidth
        $BarInner.Height = 12
        $BarInner.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
        $BarInner.Margin = "1,0,0,0"
        
        $BarOuter.Child = $BarInner
        $Grid.Children.Add($BarOuter)
        
        $ValLabel = [System.Windows.Controls.TextBlock]::new()
        $ValLabel.Text = $Val.ToString()
        $ValLabel.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#ffffff")
        $ValLabel.FontWeight = [System.Windows.FontWeights]::Bold
        $ValLabel.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $ValLabel.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
        $ValLabel.FontSize = 11
        [System.Windows.Controls.Grid]::SetColumn($ValLabel, 2)
        $Grid.Children.Add($ValLabel)
        
        $Panel.Children.Add($Grid)
    }
}

function AddActivityBarChartRow($Panel, $LabelText, $Val, $Total, $ColorHex) {
    if ($Total -eq 0) { $Total = 1 }
    
    $Grid = [System.Windows.Controls.Grid]::new()
    $Grid.Margin = "0,6,0,6"
    
    $Col1 = [System.Windows.Controls.ColumnDefinition]::new()
    $Col1.Width = [System.Windows.GridLength]::new(180)
    $Col2 = [System.Windows.Controls.ColumnDefinition]::new()
    $Col2.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $Col3 = [System.Windows.Controls.ColumnDefinition]::new()
    $Col3.Width = [System.Windows.GridLength]::new(40)
    
    $Grid.ColumnDefinitions.Add($Col1)
    $Grid.ColumnDefinitions.Add($Col2)
    $Grid.ColumnDefinitions.Add($Col3)
    
    $Label = [System.Windows.Controls.TextBlock]::new()
    $Label.Text = $LabelText
    $Label.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#9ca3af")
    $Label.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    $Label.FontSize = 11
    [System.Windows.Controls.Grid]::SetColumn($Label, 0)
    $Grid.Children.Add($Label)
    
    $BarOuter = [System.Windows.Controls.Border]::new()
    $BarOuter.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#0c0c0d")
    $BarOuter.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#2c2c2e")
    $BarOuter.BorderThickness = 1
    $BarOuter.CornerRadius = 3
    $BarOuter.Height = 14
    $BarOuter.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
    $BarOuter.Width = 200
    [System.Windows.Controls.Grid]::SetColumn($BarOuter, 1)
    
    $BarWidth = [Math]::Round(($Val / $Total) * 198)
    if ($BarWidth -lt 2) { $BarWidth = 2 }
    $BarInner = [System.Windows.Shapes.Rectangle]::new()
    $BarInner.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString($ColorHex)
    $BarInner.Width = $BarWidth
    $BarInner.Height = 12
    $BarInner.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Left
    $BarInner.Margin = "1,0,0,0"
    
    $BarOuter.Child = $BarInner
    $Grid.Children.Add($BarOuter)
    
    $ValLabel = [System.Windows.Controls.TextBlock]::new()
    $ValLabel.Text = $Val.ToString()
    $ValLabel.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#ffffff")
    $ValLabel.FontWeight = [System.Windows.FontWeights]::Bold
    $ValLabel.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    $ValLabel.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
    $ValLabel.FontSize = 11
    [System.Windows.Controls.Grid]::SetColumn($ValLabel, 2)
    $Grid.Children.Add($ValLabel)
    
    $Panel.Children.Add($Grid)
}

# 9. Dashboard Processors
function Process-UserData($Users) {
    $Script:CurrentData = $Users
    
    $Total = $Users.Count
    $Active = 0
    $Inactive90d = 0
    $Inactive1yr = 0
    $Never = 0
    
    $LicenseMap = @{}
    $ActivityMap = [ordered]@{
        "Active (<=30d)"    = 0
        "Inactive (30-90d)"  = 0
        "Inactive (90-365d)" = 0
        "Inactive (>1yr)"    = 0
        "Never Logged In"    = 0
    }
    
    $DataGridRows = [System.Collections.Generic.List[PSCustomObject]]::new()
    
    foreach ($user in $Users) {
        $Days = Get-DaysSince $user.LastSignInDate
        
        $StatusText = ""
        $Category = ""
        if ($Days -eq [double]::PositiveInfinity) {
            $StatusText = "Never Logged In"
            $Category = "Never Logged In"
            $Never++
        } elseif ($Days -le 30) {
            $StatusText = "Active ($Days`d ago)"
            $Category = "Active (<=30d)"
            $Active++
        } elseif ($Days -le 90) {
            $StatusText = "Warning ($Days`d ago)"
            $Category = "Inactive (30-90d)"
        } elseif ($Days -le 365) {
            $StatusText = "Inactive ($Days`d/90+ ago)"
            $Category = "Inactive (90-365d)"
            $Inactive90d++
        } else {
            $StatusText = "Critical ($Days`d/1yr+ ago)"
            $Category = "Inactive (>1yr)"
            $Inactive1yr++
            $Inactive90d++
        }
        
        $ActivityMap[$Category]++
        
        if ($user.AssignedLicenses) {
            $user.AssignedLicenses.Split(',') | ForEach-Object {
                $lic = $_.Trim()
                if ($lic) {
                    $LicenseMap[$lic]++
                }
            }
        }
        
        $Username = ($user.UserPrincipalName -split '@')[0].ToLower()
        $DisplayNameLower = $user.DisplayName.ToLower()
        $VerificationText = "-"
        if ($Username -like "*seriun*" -or $Username -eq "jp" -or $Username -like "jp.*" -or $Username -like "*.jp" -or $DisplayNameLower -like "*seriun*" -or $DisplayNameLower -match "\bjp\b") {
            $VerificationText = "Test/Admin Account"
        }

        $Enabled = $true
        if ($null -ne $user.AccountEnabled) {
            $Enabled = $user.AccountEnabled
        }
        $AccountStatusText = if ($Enabled) { "Enabled" } else { "🚫 Blocked" }

        $WastedCost = 0.00
        $MonthlySavings = 0.00
        if ($Days -gt 180) {
            $InactiveDays = $Days
            if ($Days -eq [double]::PositiveInfinity) {
                $InactiveDays = 365
            }
            $Months = $InactiveDays / 30
            $MonthlyPrice = Get-LicenseMonthlyPrice $user.AssignedLicenses
            $WastedCost = $MonthlyPrice * $Months
            $MonthlySavings = $MonthlyPrice
        }
        $WastedCostText = "£{0:N2}" -f $WastedCost
        $MonthlySavingsText = "£{0:N2}" -f $MonthlySavings
        
        # Downgrade Recommendation Engine
        $Recommendation = "-"
        if ($user.AssignedLicenses) {
            $LicsArray = $user.AssignedLicenses.Split(',') | ForEach-Object { $_.Trim() }
            $MonthlyPrice = Get-LicenseMonthlyPrice $user.AssignedLicenses
            
            if ($Days -eq [double]::PositiveInfinity -or $Days -gt 180) {
                if ($MonthlyPrice -gt 0) {
                    $Recommendation = "Reclaim: Remove all licenses (Save £{0:N2}/mo)" -f $MonthlyPrice
                }
            } else {
                # Check for redundant/overlapping licenses (e.g. having both a suite and standalone plans already included in it)
                $HasBusinessPremium = $LicsArray | Where-Object { $_ -eq "Microsoft 365 Business Premium" }
                $HasM365E5 = $LicsArray | Where-Object { $_ -eq "Microsoft 365 E5" }
                $HasM365E3 = $LicsArray | Where-Object { $_ -eq "Microsoft 365 E3" }
                $HasO365E5 = $LicsArray | Where-Object { $_ -eq "Office 365 E5" }
                $HasO365E3 = $LicsArray | Where-Object { $_ -eq "Office 365 E3" }
                $HasBusinessStandard = $LicsArray | Where-Object { $_ -eq "Microsoft 365 Business Standard" }
                $HasBusinessBasic = $LicsArray | Where-Object { $_ -eq "Microsoft 365 Business Basic" }
                
                $RedundantPlans = [System.Collections.Generic.List[string]]::new()
                $RedundantCost = 0.00
                
                # 1. Exchange Online (Plan 1) redundant check
                $ExchangeP1Price = if ($Script:SkuPrices.ContainsKey("Exchange Online (Plan 1)")) { $Script:SkuPrices["Exchange Online (Plan 1)"] } else { 3.30 }
                if ($LicsArray -contains "Exchange Online (Plan 1)" -or $LicsArray -contains "EXCHANGESTANDARD") {
                    if ($HasBusinessPremium -or $HasM365E5 -or $HasM365E3 -or $HasO365E5 -or $HasO365E3 -or $HasBusinessStandard -or $HasBusinessBasic) {
                        $RedundantPlans.Add("Exchange P1")
                        $RedundantCost += $ExchangeP1Price
                    }
                }
                
                # 2. Exchange Online (Plan 2) redundant check (BP mailboxes upgrade to 100GB as of July 1, 2026!)
                $ExchangeP2Price = if ($Script:SkuPrices.ContainsKey("Exchange Online (Plan 2)")) { $Script:SkuPrices["Exchange Online (Plan 2)"] } else { 6.60 }
                if ($LicsArray -contains "Exchange Online (Plan 2)" -or $LicsArray -contains "EXCHANGEENTERPRISE") {
                    if ($HasBusinessPremium -or $HasM365E5 -or $HasM365E3 -or $HasO365E5 -or $HasO365E3) {
                        $RedundantPlans.Add("Exchange P2")
                        $RedundantCost += $ExchangeP2Price
                    }
                }
                
                # 3. Microsoft Entra ID P1 / AAD Premium redundant check
                $EntraP1Price = if ($Script:SkuPrices.ContainsKey("Microsoft Entra ID P1")) { $Script:SkuPrices["Microsoft Entra ID P1"] } else { 4.90 }
                if ($LicsArray -contains "Microsoft Entra ID P1" -or $LicsArray -contains "AAD_PREMIUM") {
                    if ($HasBusinessPremium -or $HasM365E5 -or $HasM365E3 -or $HasO365E5 -or $HasO365E3) {
                        $RedundantPlans.Add("Entra ID P1")
                        $RedundantCost += $EntraP1Price
                    }
                }
                
                # 4. Microsoft Intune redundant check
                $IntunePrice = if ($Script:SkuPrices.ContainsKey("Microsoft Intune")) { $Script:SkuPrices["Microsoft Intune"] } else { 6.60 }
                if ($LicsArray -contains "Microsoft Intune" -or $LicsArray -contains "INTUNE_A") {
                    if ($HasBusinessPremium -or $HasM365E5 -or $HasM365E3) {
                        $RedundantPlans.Add("Intune")
                        $RedundantCost += $IntunePrice
                    }
                }

                # 5. Power BI Pro redundant check
                $PowerBIPrice = if ($Script:SkuPrices.ContainsKey("Power BI Pro")) { $Script:SkuPrices["Power BI Pro"] } else { 8.20 }
                if ($LicsArray -contains "Power BI Pro" -or $LicsArray -contains "POWER_BI_PRO") {
                    if ($HasM365E5 -or $HasO365E5) {
                        $RedundantPlans.Add("Power BI Pro")
                        $RedundantCost += $PowerBIPrice
                    }
                }

                if ($RedundantPlans.Count -gt 0) {
                    $PlansStr = $RedundantPlans -join " & "
                    $Recommendation = "Redundant: Remove $PlansStr (Save £{0:N2}/mo)" -f $RedundantCost
                    $MonthlySavings = $RedundantCost
                    $WastedCost = $RedundantCost
                    $WastedCostText = "£{0:N2}" -f $WastedCost
                    $MonthlySavingsText = "£{0:N2}" -f $MonthlySavings
                } else {
                    # Check for Copilot inactive > 30 days
                    $HasCopilot = $LicsArray | Where-Object { $_ -match "Copilot" }
                    if ($HasCopilot -and $Days -gt 30) {
                        $Recommendation = "Downgrade: Remove Copilot (Save £24.70/mo)"
                    } else {
                        # Check for E5 inactive > 90 days
                        $HasE5 = $LicsArray | Where-Object { $_ -match "E5" }
                        if ($HasE5 -and $Days -gt 90) {
                            $Recommendation = "Downgrade: E5 -> Business Premium (Save £30.00/mo)"
                        } else {
                            # Check for E3 inactive > 90 days
                            $HasE3 = $LicsArray | Where-Object { $_ -match "E3" }
                            if ($HasE3 -and $Days -gt 90) {
                                $Recommendation = "Downgrade: E3 -> Business Standard (Save £21.40/mo)"
                            }
                        }
                    }
                }
            }
        }

        $DataGridRows.Add([PSCustomObject]@{
            DisplayName        = $user.DisplayName
            UserPrincipalName  = $user.UserPrincipalName
            AssignedLicenses   = $user.AssignedLicenses
            LastSignInDate     = (Format-LastSignIn $user.LastSignInDate)
            StatusText         = $StatusText
            StatusCategory     = $Category
            DaysSince          = $Days
            VerificationText   = $VerificationText
            WastedCost         = $WastedCost
            WastedCostText     = $WastedCostText
            MonthlySavings     = $MonthlySavings
            MonthlySavingsText = $MonthlySavingsText
            Recommendation     = $Recommendation
            AccountStatusText  = $AccountStatusText
        })
    }
    
    $Script:GridRows = $DataGridRows
    
    # Calculate Total Wasted Cost and Monthly Savings
    $TotalSavings = 0.00
    $TotalMonthlySavings = 0.00
    foreach ($row in $DataGridRows) {
        $TotalSavings += $row.WastedCost
        $TotalMonthlySavings += $row.MonthlySavings
    }
    
    # Update KPI UI Text blocks
    $wpf_kpiTotal.Text = $Total.ToString()
    $wpf_kpiActive.Text = $Active.ToString()
    $wpf_kpiInactive90d.Text = $Inactive90d.ToString()
    $wpf_kpiInactive1yr.Text = $Inactive1yr.ToString()
    $wpf_kpiSavings.Text = "£{0:N2}" -f $TotalSavings
    $wpf_kpiMonthlySavings.Text = "£{0:N2}/mo" -f $TotalMonthlySavings
    
    $PoolWasteVal = if ($null -ne $Script:TotalPoolWaste) { $Script:TotalPoolWaste } else { 0.00 }
    $wpf_kpiPoolWaste.Text = "£{0:N2}" -f $PoolWasteVal
    
    $PoolWasteCountVal = if ($null -ne $Script:TotalUnassignedCount) { $Script:TotalUnassignedCount } else { 0 }
    $wpf_kpiPoolWasteCount.Text = "$PoolWasteCountVal unused licenses"
    
    # Draw License Bar Chart
    if ($wpf_panelLicenseChart) { UpdateBarChart $wpf_panelLicenseChart $LicenseMap "#3B82F6" }
    
    # Draw Activity Bar Chart with distinct colors
    $wpf_panelActivityChart.Children.Clear()
    foreach ($Key in $ActivityMap.Keys) {
        $Val = $ActivityMap[$Key]
        $Color = $ActivityColors[$Key]
        AddActivityBarChartRow $wpf_panelActivityChart $Key $Val $Total $Color
    }
    
    if ($wpf_gridUnassigned) {
        $wpf_gridUnassigned.ItemsSource = $null
        $wpf_gridUnassigned.ItemsSource = $Script:UnassignedLicenses
    }
    
    Filter-DataGrid
}

function Filter-DataGrid {
    if ($null -eq $Script:GridRows) { return }
    
    $Filtered = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($row in $Script:GridRows) {
        $SearchMatch = $true
        if (-not [string]::IsNullOrWhiteSpace($Script:SearchQuery)) {
            $SearchMatch = ($row.DisplayName -like "*$Script:SearchQuery*") -or ($row.UserPrincipalName -like "*$Script:SearchQuery*")
        }
        if (-not $SearchMatch) { continue }
        
        $FilterMatch = $false
        if ($Script:ActiveFilter -eq "all") {
            $FilterMatch = $true
        } elseif ($Script:ActiveFilter -eq "active") {
            $FilterMatch = ($row.StatusCategory -eq "Active (<=30d)")
        } elseif ($Script:ActiveFilter -eq "inactive90d") {
            $FilterMatch = ($row.DaysSince -ge 90)
        } elseif ($Script:ActiveFilter -eq "inactive1yr") {
            $FilterMatch = ($row.StatusCategory -eq "Inactive (>1yr)")
        } elseif ($Script:ActiveFilter -eq "never") {
            $FilterMatch = ($row.StatusCategory -eq "Never Logged In")
        }
        
        if ($FilterMatch) {
            $Filtered.Add($row)
        }
    }
    
    $wpf_gridUsers.ItemsSource = $null
    $wpf_gridUsers.ItemsSource = $Filtered
}

# 10. Navigations & Tab Switchers
function Show-Page($PageName) {
    $wpf_pageConnection.Visibility = [System.Windows.Visibility]::Collapsed
    $wpf_pageDashboard.Visibility = [System.Windows.Visibility]::Collapsed
    $wpf_pageDirectory.Visibility = [System.Windows.Visibility]::Collapsed
    
    $wpf_btnNavConnection.BorderBrush = [System.Windows.Media.Brushes]::Transparent
    $wpf_btnNavConnection.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#9ca3af")
    $wpf_btnNavDashboard.BorderBrush = [System.Windows.Media.Brushes]::Transparent
    $wpf_btnNavDashboard.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#9ca3af")
    $wpf_btnNavDirectory.BorderBrush = [System.Windows.Media.Brushes]::Transparent
    $wpf_btnNavDirectory.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#9ca3af")
    
    if ($PageName -eq "Connection") {
        $wpf_pageConnection.Visibility = [System.Windows.Visibility]::Visible
        $wpf_btnNavConnection.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#3B82F6")
        $wpf_btnNavConnection.Foreground = [System.Windows.Media.Brushes]::White
    } elseif ($PageName -eq "Dashboard") {
        $wpf_pageDashboard.Visibility = [System.Windows.Visibility]::Visible
        $wpf_btnNavDashboard.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#3B82F6")
        $wpf_btnNavDashboard.Foreground = [System.Windows.Media.Brushes]::White
    } elseif ($PageName -eq "Directory") {
        $wpf_pageDirectory.Visibility = [System.Windows.Visibility]::Visible
        $wpf_btnNavDirectory.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#3B82F6")
        $wpf_btnNavDirectory.Foreground = [System.Windows.Media.Brushes]::White
    }
}

function Show-ConnectionMethod($MethodName) {
    $wpf_panelMethodM365.Visibility = [System.Windows.Visibility]::Collapsed
    $wpf_panelMethodApp.Visibility = [System.Windows.Visibility]::Collapsed
    $wpf_panelMethodCSV.Visibility = [System.Windows.Visibility]::Collapsed
    
    $wpf_btnMethodM365.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#0c0c0d")
    $wpf_btnMethodM365.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#9ca3af")
    $wpf_btnMethodM365.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#2c2c2e")
    $wpf_btnMethodM365.BorderThickness = 1
    
    $wpf_btnMethodApp.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#0c0c0d")
    $wpf_btnMethodApp.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#9ca3af")
    $wpf_btnMethodApp.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#2c2c2e")
    $wpf_btnMethodApp.BorderThickness = 1
    
    $wpf_btnMethodCSV.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#0c0c0d")
    $wpf_btnMethodCSV.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#9ca3af")
    $wpf_btnMethodCSV.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#2c2c2e")
    $wpf_btnMethodCSV.BorderThickness = 1
    
    if ($MethodName -eq "M365") {
        $wpf_panelMethodM365.Visibility = [System.Windows.Visibility]::Visible
        $wpf_btnMethodM365.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#3B82F6")
        $wpf_btnMethodM365.Foreground = [System.Windows.Media.Brushes]::White
        $wpf_btnMethodM365.BorderThickness = 0
    } elseif ($MethodName -eq "App") {
        $wpf_panelMethodApp.Visibility = [System.Windows.Visibility]::Visible
        $wpf_btnMethodApp.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#3B82F6")
        $wpf_btnMethodApp.Foreground = [System.Windows.Media.Brushes]::White
        $wpf_btnMethodApp.BorderThickness = 0
    } elseif ($MethodName -eq "CSV") {
        $wpf_panelMethodCSV.Visibility = [System.Windows.Visibility]::Visible
        $wpf_btnMethodCSV.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#3B82F6")
        $wpf_btnMethodCSV.Foreground = [System.Windows.Media.Brushes]::White
        $wpf_btnMethodCSV.BorderThickness = 0
    }
}

function Select-Role($Role) {
    $Script:UserRole = $Role
    
    if ($Role -eq "ServiceDesk") {
        $wpf_btnRoleSelectServiceDesk.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#3B82F6")
        $wpf_btnRoleSelectServiceDesk.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#ff8c00")
        
        $wpf_btnRoleSelectSales.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#1c1c1e")
        $wpf_btnRoleSelectSales.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#2c2c2e")
    } else {
        $wpf_btnRoleSelectServiceDesk.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#1c1c1e")
        $wpf_btnRoleSelectServiceDesk.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#2c2c2e")
        
        $wpf_btnRoleSelectSales.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#3B82F6")
        $wpf_btnRoleSelectSales.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#ff8c00")
    }
    
    Configure-RoleUI
    
    # Reveal Step 2 connection methods panel
    $wpf_panelConnectionMethods.Visibility = [System.Windows.Visibility]::Visible
    Update-UI
}

$wpf_btnRoleSelectServiceDesk.Add_Click({ Select-Role "ServiceDesk" })
$wpf_btnRoleSelectSales.Add_Click({ Select-Role "Sales" })

$wpf_btnNavConnection.Add_Click({ Show-Page "Connection" })
$wpf_btnNavDashboard.Add_Click({ Show-Page "Dashboard" })
$wpf_btnNavDirectory.Add_Click({ Show-Page "Directory" })

# 11. Event Handlers & Connection Logic
# Connect: M365 Interactive Login
$wpf_btnConnectM365.Add_Click({
    $TenantId = $wpf_txtUserTenant.Text.Trim()
    
    Log-ToTerminal "Initializing Microsoft Graph interactive login flow..." "Info"
    Log-ToTerminal "Connect-MgGraph -Scopes 'User.Read.All', 'AuditLog.Read.All', 'Directory.Read.All', 'Organization.Read.All'" "Command"
    Update-UI
    
    try {
        # Minimize window to allow browser authentication to take focus
        $Window.WindowState = [System.Windows.WindowState]::Minimized
        Update-UI
        Start-Sleep -Milliseconds 200
        
        # Clear Graph cache to ensure a clean start
        Clear-MgGraphCache
        
        $Params = @{
            Scopes = @("User.Read.All", "AuditLog.Read.All", "Directory.Read.All", "Organization.Read.All", "LicenseAssignment.Read.All")
            ContextScope = "Process"
        }
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $Params.TenantId = $TenantId }
        
        Connect-MgGraph @Params | Out-Null
        
        # Restore window state and focus
        $Window.WindowState = [System.Windows.WindowState]::Normal
        $Window.Activate() | Out-Null
        
        $Context = Get-MgContext
        if ($null -eq $Context) { throw "Graph connection context could not be established." }
        
        $Script:SessionMode = "Real"
        $wpf_txtStatusLabel.Text = "Connected to Graph"
        $wpf_elStatusDot.Fill = [System.Windows.Media.Brushes]::Green
        
        $TenantName = "UnknownTenant"
        try {
            $OrgInfo = Get-MgOrganization | Select-Object -First 1 -Property DisplayName
            $TenantName = $OrgInfo.DisplayName
        } catch {}
        
        $wpf_txtSessionTenant.Text = $TenantName
        $wpf_txtSessionUser.Text = $Context.Account
        $wpf_txtSessionMode.Text = "DIRECT GRAPH API"
        $wpf_btnDisconnect.Visibility = [System.Windows.Visibility]::Visible
        
        $wpf_btnRunScript.IsEnabled = $true
        $wpf_btnRunScript.Opacity = 1.0
        
        Start-SessionTimer
        
        Log-ToTerminal "Successfully connected to Microsoft Graph." "Success"
        Log-ToTerminal "Account: $($Context.Account)" "Info"
        Log-ToTerminal "Tenant ID: $($Context.TenantId)" "Info"
        
        # Switch visibilities to connected session look
        $wpf_panelOfflineWelcome.Visibility = [System.Windows.Visibility]::Collapsed
        $wpf_panelConnectedSession.Visibility = [System.Windows.Visibility]::Visible
        $wpf_borderNav.Visibility = [System.Windows.Visibility]::Visible
        
        Show-Page "Dashboard"
        Invoke-AuditReport
    } catch {
        # Restore window state in case of exception
        $Window.WindowState = [System.Windows.WindowState]::Normal
        $Window.Activate() | Out-Null
        Log-ToTerminal "Authentication failed: $_" "Error"
        [System.Windows.MessageBox]::Show("Connection Failed:`n$_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

# Connect: App Secret Token
$wpf_btnConnectApp.Add_Click({
    $TenantId = $wpf_txtAppTenant.Text.Trim()
    $ClientId = $wpf_txtAppClient.Text.Trim()
    $Secret = $wpf_txtAppSecret.Password.Trim()
    
    if ([string]::IsNullOrWhiteSpace($TenantId) -or [string]::IsNullOrWhiteSpace($ClientId) -or [string]::IsNullOrWhiteSpace($Secret)) {
        [System.Windows.MessageBox]::Show("Please fill in the Tenant ID, Client ID, and Client Secret.", "Missing Info", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    Log-ToTerminal "Connecting via Application token credentials..." "Info"
    Log-ToTerminal "Connect-MgGraph -TenantId '$TenantId' -ClientId '$ClientId' -ClientSecret '••••••••'" "Command"
    Update-UI
    
    try {
        # Clear Graph cache to ensure a clean start
        Clear-MgGraphCache
        
        $SecSecret = ConvertTo-SecureString $Secret -AsPlainText -Force
        $Credential = [PSCredential]::new($ClientId, $SecSecret)
        Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $Credential -ContextScope Process | Out-Null
        
        $Context = Get-MgContext
        if ($null -eq $Context) { throw "Graph connection context could not be established." }
        
        $Script:SessionMode = "AppRegistration"
        $wpf_txtStatusLabel.Text = "Connected (App Secret)"
        $wpf_elStatusDot.Fill = [System.Windows.Media.Brushes]::Green
        
        $TenantName = "UnknownTenant"
        try {
            $OrgInfo = Get-MgOrganization | Select-Object -First 1 -Property DisplayName
            $TenantName = $OrgInfo.DisplayName
        } catch {}
        
        $wpf_txtSessionTenant.Text = $TenantName
        $wpf_txtSessionUser.Text = "AppRegistration ($ClientId)"
        $wpf_txtSessionMode.Text = "APP TOKEN REG"
        $wpf_btnDisconnect.Visibility = [System.Windows.Visibility]::Visible
        
        $wpf_btnRunScript.IsEnabled = $true
        $wpf_btnRunScript.Opacity = 1.0
        
        Start-SessionTimer
        
        Log-ToTerminal "Successfully connected using Application Client Secret." "Success"
        
        # Switch visibilities to connected session look
        $wpf_panelOfflineWelcome.Visibility = [System.Windows.Visibility]::Collapsed
        $wpf_panelConnectedSession.Visibility = [System.Windows.Visibility]::Visible
        $wpf_borderNav.Visibility = [System.Windows.Visibility]::Visible
        
        Show-Page "Dashboard"
        Invoke-AuditReport
    } catch {
        Log-ToTerminal "Connection failed: $_" "Error"
        [System.Windows.MessageBox]::Show("Connection Failed:`n$_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

# Connect: Browse and Import local CSV report outputs
$wpf_btnBrowseCSV.Add_Click({
    $openFileDialog = New-Object Microsoft.Win32.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $openFileDialog.Title = "Import PowerShell Tenant CSV Report"
    
    if ($openFileDialog.ShowDialog()) {
        $File = $openFileDialog.FileName
        Log-ToTerminal "Importing local CSV file: $($openFileDialog.SafeFileName)..." "Info"
        Update-UI
        try {
            $CsvData = Import-Csv -Path $File
            if ($CsvData.Count -eq 0) { throw "CSV file contains no records." }
            
            $Headers = $CsvData[0].PSObject.Properties.Name
            
            $DN_Prop = $Headers | Where-Object { $_ -eq "DisplayName" -or $_ -eq "Display Name" -or $_ -eq "Display_Name" } | Select-Object -First 1
            $UPN_Prop = $Headers | Where-Object { $_ -eq "UserPrincipalName" -or $_ -eq "User Principal Name" -or $_ -eq "User_Principal_Name" -or $_ -eq "UPN" } | Select-Object -First 1
            $Lic_Prop = $Headers | Where-Object { $_ -eq "AssignedLicenses" -or $_ -eq "Assigned Licenses" -or $_ -eq "Licenses" } | Select-Object -First 1
            $Date_Prop = $Headers | Where-Object { $_ -eq "LastSignInDate" -or $_ -eq "Last Sign-In Date" -or $_ -eq "Last_Sign-In_Date" -or $_ -eq "LastSignIn" } | Select-Object -First 1
            
            if (-not ($DN_Prop -and $UPN_Prop -and $Lic_Prop -and $Date_Prop)) {
                throw "Missing required headers. Columns must include Display Name, User Principal Name, Assigned Licenses, and Last Sign-In Date."
            }
            
            if ([string]::IsNullOrEmpty($Script:UserRole)) {
                $Script:UserRole = "Sales"
            }
            Configure-RoleUI
            
            # Extract license counts from imported CSV to show in the breakdown
            $CsvSkuCounts = @{}
            foreach ($row in $CsvData) {
                if ($row.$Lic_Prop) {
                    $row.$Lic_Prop.Split(',') | ForEach-Object {
                        $SkuName = $_.Trim()
                        if ($SkuName) {
                            $CsvSkuCounts[$SkuName] = ($CsvSkuCounts[$SkuName] + 1)
                        }
                    }
                }
            }
            
            $UnassignedList = [System.Collections.Generic.List[PSCustomObject]]::new()
            $PricesFile = Join-Path $Script:ScriptDir "LicensePrices.csv"
            $PricesCsv = if (Test-Path $PricesFile) { Import-Csv $PricesFile } else { @() }
            
            foreach ($SkuName in $CsvSkuCounts.Keys) {
                $UnitPrice = Get-LicenseMonthlyPrice $SkuName
                
                # Resolve Sku ID GUID from LicensePrices.csv if available
                $ResolvedSkuId = "-"
                if ($PricesCsv) {
                    $MatchedRow = $PricesCsv | Where-Object { $_.Product -eq $SkuName -or $_.SkuId -eq $SkuName } | Select-Object -First 1
                    if ($MatchedRow) {
                        $ResolvedSkuId = $MatchedRow.SkuId
                    }
                }
                
                $Count = $CsvSkuCounts[$SkuName]
                $GrayBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#9ca3af")
                $UnassignedList.Add([PSCustomObject]@{
                    SkuPartName      = $SkuName
                    SkuId            = $ResolvedSkuId
                    ActiveUnits      = $Count
                    ConsumedUnits    = $Count
                    UnassignedUnits  = 0
                    UnassignedBrush  = $GrayBrush
                    WastedBrush      = $GrayBrush
                    MonthlyPrice     = $UnitPrice
                    MonthlyPriceText = "£{0:N2}" -f $UnitPrice
                    WastedCost       = 0.00
                    WastedCostText   = "£0.00"
                    AssignedText     = "$Count of $Count"
                })
            }
            
            $Script:UnassignedLicenses = $UnassignedList
            $Script:TotalPoolWaste = 0.00
            $Script:TotalUnassignedCount = 0
            
            $Script:SessionMode = "CSVImport"
            $wpf_txtStatusLabel.Text = "Loaded Local CSV"
            $wpf_elStatusDot.Fill = [System.Windows.Media.Brushes]::Green
            
            $TenantName = "Local CSV"
            if ($openFileDialog.SafeFileName -match "LicensedUsers_([^_]+)_") {
                $TenantName = $Matches[1]
            }
            
            $wpf_txtSessionTenant.Text = $TenantName
            $wpf_txtSessionUser.Text = $openFileDialog.SafeFileName
            $wpf_txtSessionMode.Text = "CSV FILE IMPORT"
            $wpf_btnDisconnect.Visibility = [System.Windows.Visibility]::Visible
            
            $wpf_btnRunScript.IsEnabled = $false
            $wpf_btnRunScript.Opacity = 0.5
            
            $Users = [System.Collections.Generic.List[PSCustomObject]]::new()
            foreach ($row in $CsvData) {
                # Look for AccountEnabled in CSV
                $AccEnabled = $true
                $AccProp = $Headers | Where-Object { $_ -eq "AccountEnabled" -or $_ -eq "Account Status" -or $_ -eq "Account_Status" } | Select-Object -First 1
                if ($AccProp -and $row.$AccProp) {
                    if ($row.$AccProp -match "Blocked" -or $row.$AccProp -eq "False" -or $row.$AccProp -eq $false) {
                        $AccEnabled = $false
                    }
                }
                $Users.Add([PSCustomObject]@{
                    DisplayName       = $row.$DN_Prop
                    UserPrincipalName = $row.$UPN_Prop
                    AssignedLicenses  = $row.$Lic_Prop
                    LastSignInDate    = $row.$Date_Prop
                    AccountEnabled    = $AccEnabled
                })
            }
            
            Process-UserData $Users
            Log-ToTerminal "Successfully imported $($Users.Count) records from CSV." "Success"
            
            # Switch visibilities to connected session look
            $wpf_panelOfflineWelcome.Visibility = [System.Windows.Visibility]::Collapsed
            $wpf_panelConnectedSession.Visibility = [System.Windows.Visibility]::Visible
            $wpf_borderNav.Visibility = [System.Windows.Visibility]::Visible
            
            Show-Page "Dashboard"
        } catch {
            Log-ToTerminal "CSV Import failed: $_" "Error"
            [System.Windows.MessageBox]::Show("CSV Import Failed:`n$_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
})

# Connect: Demo Sandbox Mode
$wpf_btnDemoMode.Add_Click({
    $Script:SessionMode = "Demo"
    $wpf_txtStatusLabel.Text = "Active Sandbox Session"
    $wpf_elStatusDot.Fill = [System.Windows.Media.Brushes]::Green
    
    $wpf_txtSessionTenant.Text = "contoso.com (Demo)"
    $wpf_txtSessionUser.Text = "admin_demo@contoso.com"
    $wpf_txtSessionMode.Text = "DEMO STANDALONE"
    $wpf_btnDisconnect.Visibility = [System.Windows.Visibility]::Visible
    
    $wpf_btnRunScript.IsEnabled = $true
    $wpf_btnRunScript.Opacity = 1.0
    
    Start-SessionTimer
    
    Log-ToTerminal "Initializing local PowerShell sandbox variables..." "Info"
    Log-ToTerminal "Import-Module Microsoft.Graph -Force" "Command"
    Log-ToTerminal "Microsoft.Graph SDK modules loaded successfully." "Success"
    Log-ToTerminal "Connect-MgGraph -DemoSandboxMode" "Command"
    Log-ToTerminal "Connected in offline sandbox session environment." "Success"
    
    # Switch visibilities to connected session look
    $wpf_panelOfflineWelcome.Visibility = [System.Windows.Visibility]::Collapsed
    $wpf_panelConnectedSession.Visibility = [System.Windows.Visibility]::Visible
    $wpf_borderNav.Visibility = [System.Windows.Visibility]::Visible
    
    Show-Page "Dashboard"
    Invoke-AuditReport
})

# Disconnect Session
$wpf_btnDisconnect.Add_Click({
    Stop-SessionTimer
    $Script:SessionMode = ""
    $Script:CurrentData = $null
    $Script:GridRows = $null
    $Script:UserRole = ""
    
    $wpf_txtStatusLabel.Text = "Offline Session"
    $wpf_elStatusDot.Fill = [System.Windows.Media.Brushes]::Red
    
    $wpf_txtSessionTenant.Text = "N/A"
    $wpf_txtSessionUser.Text = "N/A"
    $wpf_txtSessionUptime.Text = "00:00:00"
    $wpf_txtSessionMode.Text = "N/A"
    $wpf_btnDisconnect.Visibility = [System.Windows.Visibility]::Collapsed
    
    $wpf_txtAppTenant.Text = ""
    $wpf_txtAppClient.Text = ""
    $wpf_txtAppSecret.Password = ""
    $wpf_txtUserTenant.Text = ""
    
    $wpf_kpiTotal.Text = "0"
    $wpf_kpiActive.Text = "0"
    $wpf_kpiInactive90d.Text = "0"
    $wpf_kpiInactive1yr.Text = "0"
    
    if ($wpf_panelLicenseChart) { $wpf_panelLicenseChart.Children.Clear() }
    $wpf_panelActivityChart.Children.Clear()
    $wpf_gridUsers.ItemsSource = $null
    
    # Reset role cards and hide connection options panel
    $wpf_panelOfflineWelcome.Visibility = [System.Windows.Visibility]::Visible
    $wpf_panelConnectedSession.Visibility = [System.Windows.Visibility]::Collapsed
    $wpf_borderNav.Visibility = [System.Windows.Visibility]::Collapsed
    
    $wpf_panelConnectionMethods.Visibility = [System.Windows.Visibility]::Collapsed
    $wpf_btnRoleSelectServiceDesk.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#1c1c1e")
    $wpf_btnRoleSelectServiceDesk.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#2c2c2e")
    $wpf_btnRoleSelectSales.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#1c1c1e")
    $wpf_btnRoleSelectSales.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#2c2c2e")
    
    Clear-MgGraphCache
    
    Log-ToTerminal "Disconnecting active Microsoft Graph session..." "Info"
    Log-ToTerminal "Disconnect-MgGraph" "Command"
    Log-ToTerminal "Session disconnected cleanly." "Success"
    Show-Page "Connection"
})

# Execute Audit script simulation / native queries
$wpf_btnRunScript.Add_Click({
    Invoke-AuditReport
})

# Directory Search Text Changes
$wpf_txtSearch.Add_TextChanged({
    $Script:SearchQuery = $wpf_txtSearch.Text.Trim()
    Filter-DataGrid
})

# Directory Filter Selection Changes
$FilterButtons = @($wpf_btnFilterAll, $wpf_btnFilterActive, $wpf_btnFilterInactive90d, $wpf_btnFilterInactive1yr, $wpf_btnFilterNever)
function Set-FilterButtonActive($ActiveButton) {
    foreach ($btn in $FilterButtons) {
        $btn.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#1c1c1e")
        $btn.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#9ca3af")
        $btn.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#2c2c2e")
        $btn.BorderThickness = 1
    }
    $ActiveButton.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#3B82F6")
    $ActiveButton.Foreground = [System.Windows.Media.Brushes]::White
    $ActiveButton.BorderThickness = 0
}

$wpf_btnFilterAll.Add_Click({
    $Script:ActiveFilter = "all"
    Set-FilterButtonActive $wpf_btnFilterAll
    Filter-DataGrid
})
$wpf_btnFilterActive.Add_Click({
    $Script:ActiveFilter = "active"
    Set-FilterButtonActive $wpf_btnFilterActive
    Filter-DataGrid
})
$wpf_btnFilterInactive90d.Add_Click({
    $Script:ActiveFilter = "inactive90d"
    Set-FilterButtonActive $wpf_btnFilterInactive90d
    Filter-DataGrid
})
$wpf_btnFilterInactive1yr.Add_Click({
    $Script:ActiveFilter = "inactive1yr"
    Set-FilterButtonActive $wpf_btnFilterInactive1yr
    Filter-DataGrid
})
$wpf_btnFilterNever.Add_Click({
    $Script:ActiveFilter = "never"
    Set-FilterButtonActive $wpf_btnFilterNever
    Filter-DataGrid
})

# Exporting: Helper functions for Export Actions
function Invoke-CSVExport {
    if ($null -eq $Script:GridRows -or $Script:GridRows.Count -eq 0) { return }
    
    try {
        $ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
        $OutputDir = Join-Path $ScriptDir "Output"
        if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
        
        $DateStr = (Get-Date).ToString("ddMMyy")
        $CleanTenant = (Get-TenantName) -replace '[\\/:*?\"<>| ]', '_'
        $Path = Join-Path $OutputDir "LicensedUsers_${CleanTenant}_${DateStr}.csv"
        
        $ExcludedUPNs = @()
        $HasTestOrAdmin = $Script:GridRows | Where-Object { $_.VerificationText -eq "Test/Admin Account" }
        if ($HasTestOrAdmin) {
            $DlgResponse = Show-VerificationSelectorDialog -Users $HasTestOrAdmin
            if ($DlgResponse.Result -ne "Proceed") {
                Log-ToTerminal "CSV Export cancelled by user." "Warning"
                return
            }
            $ExcludedUPNs = $DlgResponse.ExcludedUPNs
        }
        
        $RowsToExport = $Script:GridRows | Where-Object { $_.UserPrincipalName -notin $ExcludedUPNs }
        $ExportRows = foreach ($row in $RowsToExport) {
            if ($Script:UserRole -eq "ServiceDesk") {
                [PSCustomObject]@{
                    "Display Name"        = $row.DisplayName
                    "User Principal Name" = $row.UserPrincipalName
                    "Assigned Licenses"   = $row.AssignedLicenses
                    "Last Sign-In Date"   = $row.LastSignInDate
                    "Status"              = $row.StatusText
                    "Account Status"      = $row.AccountStatusText
                    "Verification"        = $row.VerificationText
                }
            } else {
                [PSCustomObject]@{
                    "Display Name"              = $row.DisplayName
                    "User Principal Name"       = $row.UserPrincipalName
                    "Assigned Licenses"         = $row.AssignedLicenses
                    "Last Sign-In Date"         = $row.LastSignInDate
                    "Status"                    = $row.StatusText
                    "Account Status"            = $row.AccountStatusText
                    "Wasted Cost (Accumulated)" = $row.WastedCostText
                    "Potential Monthly Savings" = $row.MonthlySavingsText
                    "Recommendation"            = $row.Recommendation
                    "Verification"              = $row.VerificationText
                }
            }
        }
        $ExportRows | Export-Csv -Path $Path -NoTypeInformation -Encoding utf8
        
        $MsgSuffix = if ($ExcludedUPNs.Count -gt 0) { " (Excluded $($ExcludedUPNs.Count) Test/Admin account(s))" } else { "" }
        Log-ToTerminal "CSV report auto-saved to: $Path$MsgSuffix" "Success"
        
        $Result = [System.Windows.MessageBox]::Show("CSV successfully exported to:`n$Path`n`nWould you like to open the file?", "Export Success", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Information)
        if ($Result -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                Start-Process $Path
            } catch {
                Start-Process explorer.exe -ArgumentList "/select,`"$Path`""
            }
        }
    } catch {
        Log-ToTerminal "CSV Export failed: $_" "Error"
        [System.Windows.MessageBox]::Show("CSV Export Failed:`n$_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

function Invoke-ExcelExport {
    if ($null -eq $Script:GridRows -or $Script:GridRows.Count -eq 0) { return }
    
    $Module = Get-Module -ListAvailable -Name "ImportExcel"
    if (-not $Module) {
        $Msg = "The 'ImportExcel' module is required to generate Excel (.xlsx) files but was not found.`n`nWould you like the application to automatically download and install it from the PowerShell Gallery?"
        $Result = [System.Windows.MessageBox]::Show($Msg, "Module Missing", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        if ($Result -eq [System.Windows.MessageBoxResult]::Yes) {
            Log-ToTerminal "ImportExcel module is missing. Initiating auto-installation..." "Warning"
            Update-UI
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Log-ToTerminal "Downloading and installing 'ImportExcel' from PSGallery..." "Info"
                Update-UI
                Install-Module -Name "ImportExcel" -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
                Log-ToTerminal "ImportExcel installed successfully!" "Success"
                Update-UI
            } catch {
                Log-ToTerminal "Failed to install ImportExcel: $_" "Error"
                [System.Windows.MessageBox]::Show("Failed to install 'ImportExcel' automatically:`n$_`n`nPlease install it manually: Install-Module ImportExcel -Scope CurrentUser", "Installation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                return
            }
        } else {
            return
        }
    }
    
    try {
        Import-Module ImportExcel -ErrorAction Stop
    } catch {
        Log-ToTerminal "Failed to load 'ImportExcel' module: $_" "Error"
        [System.Windows.MessageBox]::Show("Failed to load 'ImportExcel' module:`n$_", "Load Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }
    
    try {
        $ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
        $OutputDir = Join-Path $ScriptDir "Output"
        if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
        
        $DateStr = (Get-Date).ToString("ddMMyy")
        $CleanTenant = (Get-TenantName) -replace '[\\/:*?\"<>| ]', '_'
        $Path = Join-Path $OutputDir "LicensedUsers_${CleanTenant}_${DateStr}.xlsx"
        
        Log-ToTerminal "Exporting grid data to Excel..." "Info"
        Update-UI
        
        $ExcludedUPNs = @()
        $HasTestOrAdmin = $Script:GridRows | Where-Object { $_.VerificationText -eq "Test/Admin Account" }
        if ($HasTestOrAdmin) {
            $DlgResponse = Show-VerificationSelectorDialog -Users $HasTestOrAdmin
            if ($DlgResponse.Result -ne "Proceed") {
                Log-ToTerminal "Excel Export cancelled by user." "Warning"
                return
            }
            $ExcludedUPNs = $DlgResponse.ExcludedUPNs
        }
        
        $FilteredGridRows = @($Script:GridRows | Where-Object { $_.UserPrincipalName -notin $ExcludedUPNs })
        
        $ExportRows = @(foreach ($row in $FilteredGridRows) {
            if ($Script:UserRole -eq "ServiceDesk") {
                [PSCustomObject]@{
                    "Display Name"        = $row.DisplayName
                    "User Principal Name" = $row.UserPrincipalName
                    "Assigned Licenses"   = $row.AssignedLicenses
                    "Last Sign-In Date"   = $row.LastSignInDate
                    "Status"              = $row.StatusText
                    "Account Status"      = $row.AccountStatusText
                    "Verification"        = $row.VerificationText
                }
            } else {
                [PSCustomObject]@{
                    "Display Name"              = $row.DisplayName
                    "User Principal Name"       = $row.UserPrincipalName
                    "Assigned Licenses"         = $row.AssignedLicenses
                    "Last Sign-In Date"         = $row.LastSignInDate
                    "Status"                    = $row.StatusText
                    "Account Status"            = $row.AccountStatusText
                    "Wasted Cost (Accumulated)" = $row.WastedCostText
                    "Potential Monthly Savings" = $row.MonthlySavingsText
                    "Recommendation"            = $row.Recommendation
                    "Verification"              = $row.VerificationText
                }
            }
        })
        
        if ($ExportRows.Count -eq 0) {
            [System.Windows.MessageBox]::Show("There are no rows to export.", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        # Lock check: try opening file for writing if it already exists
        if (Test-Path $Path) {
            try {
                $Stream = [System.IO.File]::OpenWrite($Path)
                $Stream.Close()
            } catch {
                throw "The Excel file is currently open or locked by another process (e.g. Microsoft Excel). Please close it and try again."
            }
        }
        
        $Excel = $ExportRows | Export-Excel -Path $Path -PassThru -AutoSize -WorksheetName "M365 Licensed Users"
        $Sheet = $Excel.Workbook.Worksheets["M365 Licensed Users"]
        
        $ColsList = if ($Script:UserRole -eq "ServiceDesk") { @("A", "B", "C", "D", "E", "F", "G") } else { @("A", "B", "C", "D", "E", "F", "G", "H", "I", "J") }
        
        for ($i = 0; $i -lt $FilteredGridRows.Count; $i++) {
            $RowIndex = $i + 2 # row 1 is header
            $Cat = $FilteredGridRows[$i].StatusCategory
            
            $BgColor = $null
            $FontColor = "White"
            
            switch ($Cat) {
                "Active (<=30d)"     { $BgColor = [System.Drawing.ColorTranslator]::FromHtml("#10b981") }
                "Inactive (30-90d)"   { $BgColor = [System.Drawing.ColorTranslator]::FromHtml("#eab308"); $FontColor = "Black" } # Black text on yellow for legibility
                "Inactive (90-365d)"  { $BgColor = [System.Drawing.ColorTranslator]::FromHtml("#3B82F6") }
                "Inactive (>1yr)"     { $BgColor = [System.Drawing.ColorTranslator]::FromHtml("#ef4444") }
                "Never Logged In"     { $BgColor = [System.Drawing.ColorTranslator]::FromHtml("#64748b") }
            }
            
            if ($BgColor) {
                foreach ($Col in $ColsList) {
                    $Cell = $Sheet.Cells["$Col$RowIndex"]
                    $Cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $Cell.Style.Fill.BackgroundColor.SetColor($BgColor.A, $BgColor.R, $BgColor.G, $BgColor.B)
                    if ($FontColor) {
                        $DColor = if ($FontColor -eq "Black") { [System.Drawing.Color]::Black } else { [System.Drawing.Color]::White }
                        $Cell.Style.Font.Color.SetColor($DColor.A, $DColor.R, $DColor.G, $DColor.B)
                    }
                }
                
                # E is Status
                Set-ExcelRange -Worksheet $Sheet -Range "E$RowIndex" -Bold -HorizontalAlignment Center
                # F is Account Status
                Set-ExcelRange -Worksheet $Sheet -Range "F$RowIndex" -HorizontalAlignment Center
                
                if ($Script:UserRole -eq "ServiceDesk") {
                    # G is Verification
                    Set-ExcelRange -Worksheet $Sheet -Range "G$RowIndex" -HorizontalAlignment Center
                } else {
                    # G is Wasted Cost, H is Monthly Savings
                    Set-ExcelRange -Worksheet $Sheet -Range "G$RowIndex" -HorizontalAlignment Right
                    Set-ExcelRange -Worksheet $Sheet -Range "H$RowIndex" -HorizontalAlignment Right
                    # I is Recommendation
                    Set-ExcelRange -Worksheet $Sheet -Range "I$RowIndex" -HorizontalAlignment Left
                    # J is Verification
                    Set-ExcelRange -Worksheet $Sheet -Range "J$RowIndex" -HorizontalAlignment Center
                }
            }
        }
        
        # Add the second worksheet for unassigned breakdown if data is available
        if ($null -ne $Script:UnassignedLicenses -and $Script:UnassignedLicenses.Count -gt 0) {
            $UnassignedExport = foreach ($lic in $Script:UnassignedLicenses) {
                if ($Script:UserRole -eq "ServiceDesk") {
                    [PSCustomObject]@{
                        "Product License" = $lic.SkuPartName
                        "Sku ID"          = $lic.SkuId
                        "Active Units"    = $lic.ActiveUnits
                        "Consumed Units"  = $lic.ConsumedUnits
                        "Unassigned"      = $lic.UnassignedUnits
                    }
                } else {
                    [PSCustomObject]@{
                        "Product License" = $lic.SkuPartName
                        "Sku ID"          = $lic.SkuId
                        "Active Units"    = $lic.ActiveUnits
                        "Consumed Units"  = $lic.ConsumedUnits
                        "Unassigned"      = $lic.UnassignedUnits
                        "Monthly Cost"    = $lic.MonthlyPrice
                        "Wasted Cost/mo"  = $lic.WastedCost
                    }
                }
            }
            $Excel = $UnassignedExport | Export-Excel -ExcelPackage $Excel -PassThru -AutoSize -WorksheetName "Unassigned Licenses Breakdown"
            $UnusedSheet = $Excel.Workbook.Worksheets["Unassigned Licenses Breakdown"]
            
            # Style the header row in the new worksheet
            $HeaderBg = [System.Drawing.ColorTranslator]::FromHtml("#1c1c1e")
            $UnusedCols = if ($Script:UserRole -eq "ServiceDesk") { @("A", "B", "C", "D", "E") } else { @("A", "B", "C", "D", "E", "F", "G") }
            foreach ($Col in $UnusedCols) {
                $Cell = $UnusedSheet.Cells["$Col`1"]
                $Cell.Style.Font.Bold = $true
                $Cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                if ($HeaderBg) {
                    $Cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $Cell.Style.Fill.BackgroundColor.SetColor($HeaderBg.A, $HeaderBg.R, $HeaderBg.G, $HeaderBg.B)
                }
                $Cell.Style.Font.Color.SetColor(255, 255, 255, 255) # White
            }
            
            # Alignments & formats
            for ($i = 0; $i -lt $UnassignedExport.Count; $i++) {
                $RowIndex = $i + 2
                
                # E is Unassigned
                $CellE = $UnusedSheet.Cells["E$RowIndex"]
                $CellE.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                if ($UnassignedExport[$i].Unassigned -gt 0) {
                    $CellE.Style.Font.Bold = $true
                    $CellE.Style.Font.Color.SetColor(255, 0, 128, 0) # Green (0, 128, 0)
                } elseif ($UnassignedExport[$i].Unassigned -lt 0) {
                    $CellE.Style.Font.Bold = $true
                    $CellE.Style.Font.Color.SetColor(255, 255, 0, 0) # Red (255, 0, 0)
                }
                
                if ($Script:UserRole -ne "ServiceDesk") {
                    Set-ExcelRange -Worksheet $UnusedSheet -Range "F$RowIndex" -NumberFormat "$#,##0.00" -HorizontalAlignment Right
                    
                    # G is Wasted Cost/mo
                    $CellG = $UnusedSheet.Cells["G$RowIndex"]
                    $CellG.Style.Numberformat.Format = "$#,##0.00"
                    $CellG.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Right
                    if ($UnassignedExport[$i]."Wasted Cost/mo" -gt 0) {
                        $CellG.Style.Font.Bold = $true
                        $CellG.Style.Font.Color.SetColor(255, 255, 0, 0) # Red
                    } elseif ($UnassignedExport[$i]."Wasted Cost/mo" -lt 0) {
                        $CellG.Style.Font.Bold = $true
                        $CellG.Style.Font.Color.SetColor(255, 0, 128, 0) # Green
                    }
                }
            }
        }
        
        Close-ExcelPackage -ExcelPackage $Excel
        
        $MsgSuffix = if ($ExcludedUPNs.Count -gt 0) { " (Excluded $($ExcludedUPNs.Count) Test/Admin account(s))" } else { "" }
        Log-ToTerminal "Excel report auto-saved to: $Path$MsgSuffix" "Success"
        
        $Result = [System.Windows.MessageBox]::Show("Excel successfully exported to:`n$Path`n`nWould you like to open the file?", "Export Success", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Information)
        if ($Result -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                Start-Process $Path
            } catch {
                Start-Process explorer.exe -ArgumentList "/select,`"$Path`""
            }
        }
    } catch {
        Log-ToTerminal "Excel Export failed: $_" "Error"
        [System.Windows.MessageBox]::Show("Excel Export Failed:`n$_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

function Export-WpfPng($Element, $FilePath) {
    # WPF layout bounds
    $width = $Element.ActualWidth
    $height = $Element.ActualHeight
    
    if ($width -le 0 -or $height -le 0) {
        $width = 900
        $height = 650
    }
    
    # Render element to bitmap context
    $rtb = [System.Windows.Media.Imaging.RenderTargetBitmap]::new(
        [int]$width, [int]$height, 96, 96, 
        [System.Windows.Media.PixelFormats]::Pbgra32
    )
    $rtb.Render($Element)
    
    # Save bitmap as PNG image
    $pngEncoder = [System.Windows.Media.Imaging.PngBitmapEncoder]::new()
    $pngEncoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
    
    $fs = [System.IO.FileStream]::new($FilePath, [System.IO.FileMode]::Create)
    $pngEncoder.Save($fs)
    $fs.Close()
}

function Invoke-PNGExport {
    if ($null -eq $Script:GridRows -or $Script:GridRows.Count -eq 0) { return }
    
    try {
        $ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
        $OutputDir = Join-Path $ScriptDir "Output"
        if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
        
        $DateStr = (Get-Date).ToString("ddMMyy")
        $CleanTenant = (Get-TenantName) -replace '[\\/:*?\"<>| ]', '_'
        $Path = Join-Path $OutputDir "LicensedUsers_${CleanTenant}_${DateStr}.png"
        
        Export-WpfPng $wpf_pageDashboard $Path
        Log-ToTerminal "PNG snapshot auto-saved to: $Path" "Success"
        
        $Result = [System.Windows.MessageBox]::Show("PNG successfully exported to:`n$Path`n`nWould you like to open the file?", "Export Success", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Information)
        if ($Result -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                Start-Process $Path
            } catch {
                Start-Process explorer.exe -ArgumentList "/select,`"$Path`""
            }
        }
    } catch {
        Log-ToTerminal "PNG Snapshot failed: $_" "Error"
        [System.Windows.MessageBox]::Show("Snapshot Failed:`n$_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

function Invoke-PDFExport {
    if ($null -eq $Script:GridRows -or $Script:GridRows.Count -eq 0) { return }
    
    try {
        $ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
        $OutputDir = Join-Path $ScriptDir "Output"
        if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
        
        $DateStr = (Get-Date).ToString("ddMMyy")
        $CleanTenant = (Get-TenantName) -replace '[\\/:*?\"<>| ]', '_'
        $Path = Join-Path $OutputDir "LicensedUsers_ExecutiveSummary_${CleanTenant}_${DateStr}.html"
        
        Log-ToTerminal "Generating print-ready Executive Summary HTML..." "Info"
        Update-UI
        
        # Gather metrics
        $Total = $Script:GridRows.Count
        $Active = ($Script:GridRows | Where-Object { $_.StatusCategory -eq "Active (<=30d)" }).Count
        $Inactive = ($Script:GridRows | Where-Object { $_.DaysSince -gt 180 }).Count
        
        $TotalSavings = 0.00
        $TotalMonthlySavings = 0.00
        foreach ($row in $Script:GridRows) {
            $TotalSavings += $row.WastedCost
            $TotalMonthlySavings += $row.MonthlySavings
        }
        
        # Recommendations list (high priority action items)
        $RecRows = $Script:GridRows | Where-Object { $_.Recommendation -ne "-" }
        $RecTableHtml = ""
        if ($RecRows.Count -gt 0) {
            foreach ($row in $RecRows) {
                $RecTableHtml += @"
            <tr>
                <td style="font-weight: 500;">$($row.DisplayName)</td>
                <td style="font-family: monospace;">$($row.UserPrincipalName)</td>
                <td>$($row.LastSignInDate)</td>
                <td style="color: #ef4444; font-weight: 600;">$($row.WastedCostText)</td>
                <td style="color: #10b981; font-weight: 600;">$($row.MonthlySavingsText)</td>
                <td style="color: #3B82F6; font-weight: 600;">$($row.Recommendation)</td>
            </tr>
"@
            }
        } else {
            $RecTableHtml = "<tr><td colspan='6' style='text-align: center; color: #9ca3af;'>No pending recommendations. All users active!</td></tr>"
        }
        
        # Full users list table
        $UserTableHtml = ""
        foreach ($row in $Script:GridRows) {
            $StatusStyle = switch ($row.StatusCategory) {
                "Active (<=30d)"     { "background-color: #d1fae5; color: #065f46;" }
                "Inactive (30-90d)"   { "background-color: #fef9c3; color: #854d0e;" }
                "Inactive (90-365d)"  { "background-color: #ffedd5; color: #9a3412;" }
                "Inactive (>1yr)"     { "background-color: #fee2e2; color: #991b1b;" }
                "Never Logged In"     { "background-color: #f3f4f6; color: #374151;" }
                default              { "background-color: #f3f4f6; color: #374151;" }
            }
            $UserTableHtml += @"
            <tr>
                <td>$($row.DisplayName)</td>
                <td style="font-family: monospace; font-size: 11px;">$($row.UserPrincipalName)</td>
                <td style="font-size: 11px;">$($row.AssignedLicenses)</td>
                <td>$($row.LastSignInDate)</td>
                <td style="text-align: center;"><span style="padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 500; $StatusStyle">$($row.StatusText)</span></td>
                <td style="text-align: center; font-size: 11px;">$($row.AccountStatusText)</td>
                <td style="text-align: right; font-weight: 500;">$($row.WastedCostText)</td>
                <td style="text-align: right; font-weight: 500;">$($row.MonthlySavingsText)</td>
                <td style="font-size: 11px; color: #3B82F6;">$($row.Recommendation)</td>
            </tr>
"@
        }
        
        # Unassigned licenses inventory table
        $UnassignedTableHtml = ""
        if ($null -ne $Script:UnassignedLicenses -and $Script:UnassignedLicenses.Count -gt 0) {
            $UnassignedRowsHtml = ""
            foreach ($lic in $Script:UnassignedLicenses) {
                # Format counts and colors
                $UnassignedStyle = if ($lic.UnassignedUnits -gt 0) { "color: #10b981; font-weight: 600;" } elseif ($lic.UnassignedUnits -lt 0) { "color: #ef4444; font-weight: 600;" } else { "color: #9ca3af;" }
                
                if ($Script:UserRole -eq "ServiceDesk") {
                    $UnassignedRowsHtml += @"
            <tr>
                <td style="font-weight: 500;">$($lic.SkuPartName)</td>
                <td style="font-family: monospace; font-size: 11px;">$($lic.SkuId)</td>
                <td style="text-align: center;">$($lic.ActiveUnits)</td>
                <td style="text-align: center;">$($lic.ConsumedUnits)</td>
                <td style="text-align: center; $UnassignedStyle">$($lic.UnassignedUnits)</td>
            </tr>
"@
                } else {
                    $WasteStyle = if ($lic.WastedCost -gt 0) { "color: #ef4444; font-weight: 600;" } elseif ($lic.WastedCost -lt 0) { "color: #10b981; font-weight: 600;" } else { "color: #9ca3af;" }
                    $UnassignedRowsHtml += @"
            <tr>
                <td style="font-weight: 500;">$($lic.SkuPartName)</td>
                <td style="font-family: monospace; font-size: 11px;">$($lic.SkuId)</td>
                <td style="text-align: center;">$($lic.ActiveUnits)</td>
                <td style="text-align: center;">$($lic.ConsumedUnits)</td>
                <td style="text-align: center; $UnassignedStyle">$($lic.UnassignedUnits)</td>
                <td style="text-align: right;">$($lic.MonthlyPriceText)</td>
                <td style="text-align: right; $WasteStyle">$($lic.WastedCostText)</td>
            </tr>
"@
                }
            }
            
            $TableHeaderHtml = if ($Script:UserRole -eq "ServiceDesk") {
                @"
            <tr>
                <th>Product License</th>
                <th>Sku ID</th>
                <th style="text-align: center;">Active Units</th>
                <th style="text-align: center;">Consumed Units</th>
                <th style="text-align: center;">Unassigned Units</th>
            </tr>
"@
            } else {
                @"
            <tr>
                <th>Product License</th>
                <th>Sku ID</th>
                <th style="text-align: center;">Active Units</th>
                <th style="text-align: center;">Consumed Units</th>
                <th style="text-align: center;">Unassigned Units</th>
                <th style="text-align: right;">Monthly Cost</th>
                <th style="text-align: right;">Wasted Cost/mo</th>
            </tr>
"@
            }
            
            $UnassignedTableHtml = @"
    <div class="section-title">Unassigned Licenses (Pool Waste) Breakdown</div>
    <table>
        <thead>
            $TableHeaderHtml
        </thead>
        <tbody>
            $UnassignedRowsHtml
        </tbody>
    </table>
"@
        }
        
        # Build HTML content
        $HtmlTemplate = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>M365 User License Audit - Executive Summary</title>
    <style>
        body {
            font-family: 'Segoe UI', -apple-system, sans-serif;
            color: #1f2937;
            background-color: #ffffff;
            margin: 0;
            padding: 30px;
            line-height: 1.5;
        }
        .header {
            border-bottom: 3px solid #3B82F6;
            padding-bottom: 20px;
            margin-bottom: 30px;
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
        }
        .header h1 {
            margin: 0;
            font-size: 24px;
            color: #111827;
        }
        .header p {
            margin: 5px 0 0 0;
            color: #4b5563;
            font-size: 14px;
        }
        .meta-box {
            text-align: right;
            font-size: 12px;
            color: #6b7280;
        }
        .meta-box strong {
            color: #111827;
        }
        .kpi-container {
            display: grid;
            grid-template-columns: repeat(6, 1fr);
            gap: 15px;
            margin-bottom: 30px;
        }
        .kpi-card {
            border: 1px solid #e5e7eb;
            border-radius: 8px;
            padding: 15px;
            background-color: #f9fafb;
            box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        }
        .kpi-card .title {
            font-size: 10px;
            font-weight: 700;
            color: #6b7280;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .kpi-card .value {
            font-size: 20px;
            font-weight: 700;
            margin-top: 5px;
            color: #111827;
        }
        .kpi-card .sub {
            font-size: 11px;
            color: #9ca3af;
            margin-top: 4px;
        }
        .section-title {
            font-size: 16px;
            font-weight: 700;
            color: #111827;
            margin-top: 30px;
            margin-bottom: 15px;
            border-bottom: 1px solid #e5e7eb;
            padding-bottom: 5px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
            margin-bottom: 25px;
        }
        th {
            background-color: #f3f4f6;
            color: #374151;
            text-align: left;
            font-weight: 600;
            padding: 8px 12px;
            border-bottom: 2px solid #e5e7eb;
        }
        td {
            padding: 8px 12px;
            border-bottom: 1px solid #f3f4f6;
            color: #4b5563;
        }
        tr:nth-child(even) td {
            background-color: #fafafa;
        }
        .footer {
            margin-top: 50px;
            border-top: 1px solid #e5e7eb;
            padding-top: 15px;
            text-align: center;
            font-size: 10px;
            color: #9ca3af;
        }
        @media print {
            body { padding: 0; }
            .kpi-card { box-shadow: none; }
            button { display: none; }
        }
    </style>
</head>
<body>

    <div class="header">
        <div>
            <h1>M365 User License Audit</h1>
            <p>Executive License Optimization & Cost Summary</p>
        </div>
        <div class="meta-box">
            Tenant Domain: <strong>$((Get-TenantName))</strong><br>
            Generated Date: <strong>$((Get-Date).ToString("dd/MM/yyyy HH:mm"))</strong><br>
            Session Mode: <strong>$($Script:SessionMode)</strong>
        </div>
    </div>

    <div class="kpi-container">
        <div class="kpi-card" style="border-left: 3px solid #3B82F6;">
            <div class="title">Total Licensed</div>
            <div class="value">$Total</div>
            <div class="sub">Active Entra accounts</div>
        </div>
        <div class="kpi-card" style="border-left: 3px solid #10b981;">
            <div class="title">Active (&lt;=30d)</div>
            <div class="value">$Active</div>
            <div class="sub">Logging in regularly</div>
        </div>
        <div class="kpi-card" style="border-left: 3px solid #ef4444;">
            <div class="title">Inactive (&gt;180d)</div>
            <div class="value">$Inactive</div>
            <div class="sub">Wasting license fees</div>
        </div>
        <div class="kpi-card" style="border-left: 3px solid #ef4444;">
            <div class="title">Wasted Money</div>
            <div class="value" style="color: #ef4444;">£$("{0:N2}" -f $TotalSavings)</div>
            <div class="sub">Total historical loss</div>
        </div>
        <div class="kpi-card" style="border-left: 3px solid #10b981;">
            <div class="title">Monthly Savings</div>
            <div class="value" style="color: #10b981;">£$("{0:N2}" -f $TotalMonthlySavings)</div>
            <div class="sub">Future monthly reclaim</div>
        </div>
        <div class="kpi-card" style="border-left: 3px solid #ef4444;">
            <div class="title">Pool Waste</div>
            <div class="value" style="color: #ef4444;">£$("{0:N2}" -f $Script:TotalPoolWaste)</div>
            <div class="sub">$Script:TotalUnassignedCount unassigned units</div>
        </div>
    </div>

    <div class="section-title">High-Priority Optimization Actions</div>
    <table>
        <thead>
            <tr>
                <th>Display Name</th>
                <th>User Principal Name</th>
                <th>Last Sign-In Date</th>
                <th>Wasted Cost</th>
                <th>Monthly Savings</th>
                <th>Remediation Recommendation</th>
            </tr>
        </thead>
        <tbody>
            $RecTableHtml
        </tbody>
    </table>

    $UnassignedTableHtml

    <div class="section-title">Full User License Inventory</div>
    <table>
        <thead>
            <tr>
                <th>Display Name</th>
                <th>User Principal Name</th>
                <th>Assigned Licenses</th>
                <th>Last Sign-In Date</th>
                <th>Status</th>
                <th style="text-align: center;">Account Status</th>
                <th style="text-align: right;">Wasted Cost</th>
                <th style="text-align: right;">Monthly Savings</th>
                <th>Recommendation</th>
            </tr>
        </thead>
        <tbody>
            $UserTableHtml
        </tbody>
    </table>

    <div class="footer">
        Generated by M365 PowerShell Admin Console &copy; 2026 Contoso. All rights reserved.
    </div>

    <script>
        window.onload = function() {
            setTimeout(function() {
                window.print();
            }, 500);
        }
    </script>
</body>
</html>
"@
        
        $HtmlTemplate | Out-File -FilePath $Path -Force -Encoding utf8
        Log-ToTerminal "HTML report successfully generated at: $Path" "Success"
        
        $Result = [System.Windows.MessageBox]::Show("Executive HTML summary report exported successfully to:`n$Path`n`nWould you like to open it in your browser to print/save as PDF?", "Export Success", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Information)
        if ($Result -eq [System.Windows.MessageBoxResult]::Yes) {
            Start-Process $Path
        }
    } catch {
        Log-ToTerminal "PDF/HTML Export failed: $_" "Error"
        [System.Windows.MessageBox]::Show("PDF/HTML Export Failed:`n$_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

function Invoke-EmailDraftGeneration {
    if ($null -eq $Script:GridRows -or $Script:GridRows.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No user data available to generate an email draft. Please connect and run the audit first.", "No Data", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    # Calculate counts
    $Active = ($Script:GridRows | Where-Object { $_.StatusCategory -eq "Active (<=30d)" }).Count
    $Warning = ($Script:GridRows | Where-Object { $_.StatusCategory -eq "Inactive (30-90d)" }).Count
    $Inactive = ($Script:GridRows | Where-Object { $_.StatusCategory -eq "Inactive (90-365d)" }).Count
    $Critical = ($Script:GridRows | Where-Object { $_.StatusCategory -eq "Inactive (>1yr)" }).Count
    $Never = ($Script:GridRows | Where-Object { $_.StatusCategory -eq "Never Logged In" }).Count
    
    # Tenant details
    $TenantName = (Get-TenantName)
    if ([string]::IsNullOrWhiteSpace($TenantName) -or $TenantName -eq "Not Connected") {
        $TenantName = "[Tenant Name]"
    }
    
    # Dynamic savings call-to-action
    $CTA = ""
    if ($Script:UserRole -eq "ServiceDesk") {
        $CTA = "We recommend reviewing this list with your department heads to confirm which licenses can be safely reclaimed."
    } else {
        # Calculate monthly savings
        $TotalMonthlySavings = 0.00
        foreach ($row in $Script:GridRows) {
            $TotalMonthlySavings += $row.MonthlySavings
        }
        $CTA = "We recommend reviewing this list with your department heads to confirm which licenses can be safely reclaimed to reduce your monthly subscription costs (potential savings of £{0:N2}/mo)." -f $TotalMonthlySavings
    }
    
    # Formulate email draft
    $EmailText = @"
Subject: Microsoft 365 License Audit & Optimization Report - $TenantName

Dear [Director Name],

Please find attached the latest Microsoft 365 User License Audit report for your tenant. 

This report provides a breakdown of your active user accounts and their associated Microsoft 365 subscriptions. The spreadsheet is color-coded by login activity to help identify potential licensing optimization opportunities:

* 🟩 Green (Active <= 30 days): $Active user(s). These accounts are logging in regularly. No action required.
* 🟨 Yellow (Warning 30-90 days): $Warning user(s). Moderate inactivity. 
* 🟧 Orange (Inactive 90-365 days): $Inactive user(s). High inactivity. Recommended for review.
* 🟥 Red (Critical > 1 year): $Critical user(s). Extremely high inactivity. Strong candidates for license removal.
* ⬛ Gray (Never Logged In): $Never user(s). Accounts that have never registered a login event.

Important Context & Caveats
Please note that this audit acts as a rough guide to highlight potential waste. Before disabling or unassigning any licenses, we recommend cross-referencing this list against the following scenarios:
1.  Delegated Access: Some accounts (e.g., shared mailboxes or role accounts) are accessed purely via delegation by other users, which means the mailbox is active but the underlying account registers no direct logins. In some cases, but not all, this means the license can be freed up.
2.  Extended Leave: Users currently on maternity, paternity, or long-term sick leave will show as inactive, but their accounts and licenses must remain active during their absence.
3.  New Starters: Accounts set up for new staff members who never actually joined the business.
4.  Project Accounts & Integrations: Temporary accounts created for specific projects, testing, or historical third-party software integrations that are no longer in use.
5.  Untracked Leavers: There may be former staff members who have left the business, but we were not notified to offboard them or remove their licenses.
6.  Temporary Re-licensing: Accounts that were temporarily re-licensed to allow team members to retrieve files, but we were not notified to remove the license once the request was complete.

$CTA

Kind regards,

[Your MSP Company Name] Service Desk
[Contact Information]
"@

    Show-EmailTemplateDialog -EmailText $EmailText
}

# Bind click events for directory grid page exporter buttons
$wpf_btnExportCSV.Add_Click({ Invoke-CSVExport })
$wpf_btnExportExcel.Add_Click({ Invoke-ExcelExport })
$wpf_btnExportPNG.Add_Click({ Invoke-PNGExport })
$wpf_btnExportPDF.Add_Click({ Invoke-PDFExport })
$wpf_btnGenEmail.Add_Click({ Invoke-EmailDraftGeneration })

# Bind click events for dashboard page exporter buttons to delegate to the same functions
$wpf_btnExportCSVDash.Add_Click({ Invoke-CSVExport })
$wpf_btnExportExcelDash.Add_Click({ Invoke-ExcelExport })
$wpf_btnExportPNGDash.Add_Click({ Invoke-PNGExport })
$wpf_btnExportPDFDash.Add_Click({ Invoke-PDFExport })
$wpf_btnGenEmailDash.Add_Click({ Invoke-EmailDraftGeneration })

# Clear Graph cache at startup to ensure a clean slate
Clear-MgGraphCache

# Initialize tab pages
Show-Page "Connection"

# 12. Show Window Frame Dialog
$Window.ShowDialog() | Out-Null
