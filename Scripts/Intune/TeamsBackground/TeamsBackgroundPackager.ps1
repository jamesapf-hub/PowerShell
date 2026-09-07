<#
.SYNOPSIS
    Microsoft Teams Custom Background Win32 App Packager for Intune
.DESCRIPTION
    A modern WPF GUI tool that packages custom Microsoft Teams background images
    into an Intune Win32 application (.intunewin). Generates automated deployment,
    uninstallation, and detection scripts, downloads IntuneWinAppUtil if needed,
    and outputs all exact Intune configuration settings.
.EXAMPLE
    .\TeamsBackgroundPackager.ps1
#>

# Ensure Single-Threaded Apartment (STA) mode required for WPF
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $scriptPath = $MyInvocation.MyCommand.Definition
    if (-not $scriptPath) { $scriptPath = $PSCommandPath }
    if ($scriptPath) {
        Write-Host "Re-launching in STA mode..." -ForegroundColor Cyan
        Start-Process powershell.exe -ArgumentList "-STA -ExecutionPolicy Bypass -NoProfile -File `"$scriptPath`""
        exit
    }
}

# Add required assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Drawing, System.Windows.Forms, System.Xaml

# Setup working directories
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$ToolsDir = Join-Path $ScriptDir "tools"
$DefaultOutputDir = Join-Path $ScriptDir "Output"

if (-not (Test-Path $ToolsDir)) {
    New-Item -Path $ToolsDir -ItemType Directory -Force | Out-Null
}
if (-not (Test-Path $DefaultOutputDir)) {
    New-Item -Path $DefaultOutputDir -ItemType Directory -Force | Out-Null
}

# Image Items Collection (ObservableCollection for UI data binding)
$Global:ImageItems = New-Object System.Collections.ObjectModel.ObservableCollection[Object]

# Class to store image item details (compiled only once per session)
if (-not ([System.Management.Automation.PSTypeName]'TeamsImageItem').Type) {
    Add-Type -TypeDefinition @"
using System;
using System.Windows.Media.Imaging;

public class TeamsImageItem
{
    public string FileName { get; set; }
    public string FullPath { get; set; }
    public string Resolution { get; set; }
    public string FileSize { get; set; }
    public string AspectRatio { get; set; }
    public bool IsRecommendedSize { get; set; }
    public BitmapImage Thumbnail { get; set; }
}
"@ -ReferencedAssemblies "PresentationCore", "WindowsBase", "PresentationFramework", "System.Xaml"
}

# XAML Definition
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Microsoft Teams Background Win32 App Packager for Intune"
        Height="820" Width="1080" MinHeight="700" MinWidth="950"
        WindowStartupLocation="CenterScreen"
        Background="#0F172A" Foreground="#F8FAFC"
        FontFamily="Segoe UI, Segoe UI Variable, Arial">

    <Window.Resources>
        <!-- Modern ScrollViewer Style -->
        <Style TargetType="ScrollBar">
            <Setter Property="Background" Value="#1E293B"/>
            <Setter Property="Foreground" Value="#475569"/>
        </Style>

        <!-- Base Button Style -->
        <Style TargetType="Button">
            <Setter Property="Background" Value="#334155"/>
            <Setter Property="Foreground" Value="#F8FAFC"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" 
                                CornerRadius="6" Padding="{TemplateBinding Padding}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                BorderBrush="{TemplateBinding BorderBrush}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#475569"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#1E293B"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Primary Accent Button -->
        <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#6264A7"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" 
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#7B7DB7"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#4F5294"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Emerald Action Button -->
        <Style x:Key="SuccessButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#10B981"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" 
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
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
                                <Setter TargetName="border" Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Danger Button -->
        <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#EF4444"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" 
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#DC2626"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#B91C1C"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Secondary Button -->
        <Style x:Key="SecondaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#2563EB"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" 
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#1D4ED8"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#1E40AF"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Styled TextBox -->
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#1E293B"/>
            <Setter Property="Foreground" Value="#F8FAFC"/>
            <Setter Property="BorderBrush" Value="#334155"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="CaretBrush" Value="#6264A7"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border Background="{TemplateBinding Background}" 
                                BorderBrush="{TemplateBinding BorderBrush}" 
                                BorderThickness="{TemplateBinding BorderThickness}" 
                                CornerRadius="6">
                            <ScrollViewer x:Name="PART_ContentHost"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- CheckBox Style -->
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#F8FAFC"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Margin" Value="0,3"/>
        </Style>

        <!-- RadioButton Style -->
        <Style TargetType="RadioButton">
            <Setter Property="Foreground" Value="#F8FAFC"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Margin" Value="0,3"/>
        </Style>

        <!-- Card Style -->
        <Style x:Key="CardBorder" TargetType="Border">
            <Setter Property="Background" Value="#1E293B"/>
            <Setter Property="CornerRadius" Value="10"/>
            <Setter Property="BorderBrush" Value="#334155"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="18"/>
            <Setter Property="Margin" Value="6"/>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header Banner -->
        <Border Grid.Row="0" Background="#1E293B" BorderBrush="#334155" BorderThickness="0,0,0,1" Padding="24,16">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Teams Packager Icon -->
                <Border Grid.Column="0" Background="#6264A7" Width="48" Height="48" CornerRadius="10" Margin="0,0,16,0">
                    <Viewbox Width="26" Height="26" HorizontalAlignment="Center" VerticalAlignment="Center">
                        <Canvas Width="24" Height="24">
                            <Path Fill="#FFFFFF" Data="M16.5 12c1.38 0 2.49-1.12 2.49-2.5S17.88 7 16.5 7 14 8.12 14 9.5s1.11 2.5 2.5 2.5zm-9 0c1.38 0 2.49-1.12 2.49-2.5S8.88 7 7.5 7 5 8.12 5 9.5 6.11 12 7.5 12zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-1.5c0-2.33-4.67-3.5-7-3.5zm9 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-1.5c0-2.33-4.67-3.5-7-3.5z"/>
                        </Canvas>
                    </Viewbox>
                </Border>

                <!-- Titles -->
                <StackPanel Grid.Column="1" VerticalAlignment="Center">
                    <TextBlock Text="Microsoft Teams Background Packager for Intune" FontSize="20" FontWeight="Bold" Foreground="#F8FAFC"/>
                    <TextBlock Text="Package custom video meeting backgrounds into a production-ready Intune Win32 (.intunewin) application" FontSize="13" Foreground="#94A3B8" Margin="0,3,0,0"/>
                </StackPanel>

                <!-- Version Badge -->
                <Border Grid.Column="2" Background="#334155" CornerRadius="12" Padding="12,4" VerticalAlignment="Center">
                    <TextBlock Text="v1.2.0 | Win32 App" FontSize="12" FontWeight="SemiBold" Foreground="#CBD5E1"/>
                </Border>
            </Grid>
        </Border>

        <!-- Main Tab Control -->
        <TabControl Grid.Row="1" x:Name="MainTabs" Background="Transparent" BorderThickness="0" Margin="12,10,12,0">
            <TabControl.Resources>
                <Style TargetType="TabItem">
                    <Setter Property="Template">
                        <Setter.Value>
                            <ControlTemplate TargetType="TabItem">
                                <Border x:Name="TabBorder" Background="Transparent" Padding="20,10" Margin="0,0,6,0" CornerRadius="6,6,0,0">
                                    <ContentPresenter ContentSource="Header" TextBlock.FontSize="14" TextBlock.FontWeight="SemiBold"/>
                                </Border>
                                <ControlTemplate.Triggers>
                                    <Trigger Property="IsSelected" Value="True">
                                        <Setter TargetName="TabBorder" Property="Background" Value="#1E293B"/>
                                        <Setter Property="Foreground" Value="#6264A7"/>
                                    </Trigger>
                                    <Trigger Property="IsSelected" Value="False">
                                        <Setter Property="Foreground" Value="#94A3B8"/>
                                    </Trigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                </Style>
            </TabControl.Resources>

            <!-- TAB 1: BUILD PACKAGE -->
            <TabItem Header="Package Builder">
                <Grid Margin="0,10,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="380"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <!-- Left Column: Settings -->
                    <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto" Margin="0,0,6,0">
                        <StackPanel>
                            <!-- Package Details Card -->
                            <Border Style="{StaticResource CardBorder}">
                                <StackPanel>
                                    <TextBlock Text="PACKAGE CONFIGURATION" FontSize="11" FontWeight="Bold" Foreground="#6264A7" Margin="0,0,0,12"/>

                                    <TextBlock Text="Application Name" FontSize="12" Foreground="#94A3B8" Margin="0,0,0,4"/>
                                    <TextBox x:Name="txtAppName" Text="Teams Backgrounds - Corporate Branding" Margin="0,0,0,12"/>

                                    <TextBlock Text="Publisher" FontSize="12" Foreground="#94A3B8" Margin="0,0,0,4"/>
                                    <TextBox x:Name="txtPublisher" Text="Corporate IT" Margin="0,0,0,12"/>

                                    <TextBlock Text="Version" FontSize="12" Foreground="#94A3B8" Margin="0,0,0,4"/>
                                    <TextBox x:Name="txtVersion" Text="1.0.0" Margin="0,0,0,12"/>

                                    <TextBlock Text="Description" FontSize="12" Foreground="#94A3B8" Margin="0,0,0,4"/>
                                    <TextBox x:Name="txtDescription" Text="Standard corporate branded background images for Microsoft Teams video meetings." TextWrapping="Wrap" AcceptsReturn="True" Height="50" Margin="0,0,0,12"/>

                                    <TextBlock Text="TARGET TEAMS CLIENTS" FontSize="11" FontWeight="Bold" Foreground="#6264A7" Margin="0,6,0,8"/>
                                    <CheckBox x:Name="chkNewTeams" Content="New Microsoft Teams (MSTeams / v2)" IsChecked="True"/>
                                    <CheckBox x:Name="chkClassicTeams" Content="Classic Microsoft Teams (v1)" IsChecked="True" Margin="0,2,0,12"/>

                                    <TextBlock Text="DEPLOYMENT CONTEXT" FontSize="11" FontWeight="Bold" Foreground="#6264A7" Margin="0,6,0,8"/>
                                    <RadioButton x:Name="rbAutoContext" GroupName="Context" Content="Smart Auto-Detect (Recommended)" IsChecked="True"/>
                                    <TextBlock Text="Deploys to all user profiles in C:\Users (including Sandbox &amp; Default profile) when running as SYSTEM. Deploys to current user if run as User." FontSize="11" Foreground="#64748B" Margin="20,0,0,8" TextWrapping="Wrap"/>

                                    <RadioButton x:Name="rbUserContext" GroupName="Context" Content="User Context"/>
                                    <TextBlock Text="Configured for Intune 'User' context (with fail-safe fallback if run as SYSTEM)." FontSize="11" Foreground="#64748B" Margin="20,0,0,8" TextWrapping="Wrap"/>

                                    <RadioButton x:Name="rbSystemContext" GroupName="Context" Content="System Context (All Users)"/>
                                    <TextBlock Text="Always installs to all user profiles under C:\Users (Intune: 'System')." FontSize="11" Foreground="#64748B" Margin="20,0,0,12" TextWrapping="Wrap"/>

                                    <TextBlock Text="LOG DIRECTORY (ENDPOINTS)" FontSize="11" FontWeight="Bold" Foreground="#6264A7" Margin="0,6,0,4"/>
                                    <TextBox x:Name="txtLogDir" Text="C:\Log\TeamsBackgrounds" Margin="0,0,0,12"/>

                                    <TextBlock Text="OUTPUT DIRECTORY" FontSize="11" FontWeight="Bold" Foreground="#6264A7" Margin="0,6,0,4"/>
                                    <Grid Margin="0,0,0,6">
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBox x:Name="txtOutputDir" Grid.Column="0" IsReadOnly="True" Margin="0,0,6,0"/>
                                        <Button x:Name="btnBrowseOutput" Grid.Column="1" Content="Browse..." Padding="10,6"/>
                                    </Grid>
                                </StackPanel>
                            </Border>
                        </StackPanel>
                    </ScrollViewer>

                    <!-- Right Column: Images List -->
                    <Grid Grid.Column="1" Margin="6,0,0,0">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <!-- Action Toolbar -->
                        <Border Grid.Row="0" Style="{StaticResource CardBorder}" Padding="12,10" Margin="0,0,0,8">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>

                                <StackPanel Grid.Column="0" Orientation="Horizontal">
                                    <Button x:Name="btnAddImages" Content="+ Add Images..." Style="{StaticResource PrimaryButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="btnRemoveSelected" Content="Remove Selected" Style="{StaticResource DangerButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="btnClearAll" Content="Clear All" Margin="0,0,8,0"/>
                                </StackPanel>

                                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                                    <TextBlock Text="Backgrounds: " FontSize="13" Foreground="#94A3B8"/>
                                    <TextBlock x:Name="lblImageCount" Text="0 images" FontSize="13" FontWeight="Bold" Foreground="#38BDF8"/>
                                </StackPanel>
                            </Grid>
                        </Border>

                        <!-- Images ListView -->
                        <Border Grid.Row="1" Style="{StaticResource CardBorder}" Padding="4" Margin="0,0,0,0">
                            <Grid>
                                <ListView x:Name="lvImages" Background="Transparent" BorderThickness="0" Foreground="#F8FAFC"
                                          SelectionMode="Extended" ScrollViewer.HorizontalScrollBarVisibility="Disabled">
                                    <ListView.View>
                                        <GridView>
                                            <GridViewColumn Header="Preview" Width="110">
                                                <GridViewColumn.CellTemplate>
                                                    <DataTemplate>
                                                        <Border Background="#0F172A" CornerRadius="4" Padding="2" Margin="2" BorderBrush="#334155" BorderThickness="1">
                                                            <Image Source="{Binding Thumbnail}" Width="96" Height="54" Stretch="UniformToFill"/>
                                                        </Border>
                                                    </DataTemplate>
                                                </GridViewColumn.CellTemplate>
                                            </GridViewColumn>
                                            <GridViewColumn Header="File Name" Width="190" DisplayMemberBinding="{Binding FileName}"/>
                                            <GridViewColumn Header="Resolution" Width="110" DisplayMemberBinding="{Binding Resolution}"/>
                                            <GridViewColumn Header="Size" Width="80" DisplayMemberBinding="{Binding FileSize}"/>
                                            <GridViewColumn Header="Status" Width="110" DisplayMemberBinding="{Binding AspectRatio}"/>
                                            <GridViewColumn Header="Source Path" Width="200" DisplayMemberBinding="{Binding FullPath}"/>
                                        </GridView>
                                    </ListView.View>
                                </ListView>

                                <!-- Empty State Placeholder -->
                                <StackPanel x:Name="pnlEmptyState" VerticalAlignment="Center" HorizontalAlignment="Center">
                                    <Viewbox Width="54" Height="54" HorizontalAlignment="Center" Margin="0,0,0,12">
                                        <Canvas Width="24" Height="24">
                                            <Path Fill="#475569" Data="M21 19V5c0-1.1-.9-2-2-2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2zM8.5 13.5l2.5 3.01L14.5 12l4.5 6H5l3.5-4.5z"/>
                                        </Canvas>
                                    </Viewbox>
                                    <TextBlock Text="No Teams background images added yet" FontSize="16" FontWeight="SemiBold" HorizontalAlignment="Center" Foreground="#94A3B8"/>
                                    <TextBlock Text="Click 'Add Images' to select .PNG or .JPG corporate backgrounds (1920x1080 recommended)" FontSize="12" Foreground="#64748B" HorizontalAlignment="Center" Margin="0,6,0,16"/>
                                    <Button x:Name="btnAddImagesEmpty" Content="+ Select Images" Style="{StaticResource PrimaryButton}" HorizontalAlignment="Center" Padding="20,10"/>
                                </StackPanel>
                            </Grid>
                        </Border>
                    </Grid>
                </Grid>
            </TabItem>

            <!-- TAB 2: INTUNE GUIDANCE -->
            <TabItem x:Name="tabIntuneGuidance" Header="Intune Deployment Details">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="0,10,0,0">
                    <StackPanel>
                        <!-- Package Status Banner -->
                        <Border Style="{StaticResource CardBorder}" Background="#1E3A8A" BorderBrush="#3B82F6">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <Viewbox Grid.Column="0" Width="36" Height="36" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="0,0,16,0">
                                    <Canvas Width="24" Height="24">
                                        <Path Fill="#38BDF8" Data="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 15l-5-5 1.41-1.41L10 14.17l7.59-7.59L19 8l-9 9z"/>
                                    </Canvas>
                                </Viewbox>
                                <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                    <TextBlock Text="Win32 App Package Built Successfully!" FontSize="16" FontWeight="Bold" Foreground="#FFFFFF"/>
                                    <TextBlock x:Name="lblPackagePath" Text="Package path will appear here once built." FontSize="12" Foreground="#BFDBFE" Margin="0,3,0,0" TextWrapping="Wrap"/>
                                </StackPanel>
                                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                                    <Button x:Name="btnOpenOutputDir" Content="Open Folder" Style="{StaticResource SecondaryButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="btnCopyPackagePath" Content="Copy Path"/>
                                </StackPanel>
                            </Grid>
                        </Border>

                        <!-- Step 1: Program Settings -->
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <TextBlock Text="STEP 1: INTUNE PROGRAM SETTINGS" FontSize="12" FontWeight="Bold" Foreground="#6264A7" Margin="0,0,0,12"/>

                                <Grid Margin="0,0,0,12">
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Install Command" FontSize="12" Foreground="#94A3B8" Margin="0,0,0,4"/>
                                    <TextBox x:Name="txtInstallCommand" Grid.Row="1" Grid.Column="0" IsReadOnly="True" Margin="0,0,8,0"/>
                                    <Button x:Name="btnCopyInstallCmd" Grid.Row="1" Grid.Column="1" Content="Copy" Padding="14,6"/>
                                </Grid>

                                <Grid Margin="0,0,0,12">
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Uninstall Command" FontSize="12" Foreground="#94A3B8" Margin="0,0,0,4"/>
                                    <TextBox x:Name="txtUninstallCommand" Grid.Row="1" Grid.Column="0" IsReadOnly="True" Margin="0,0,8,0"/>
                                    <Button x:Name="btnCopyUninstallCmd" Grid.Row="1" Grid.Column="1" Content="Copy" Padding="14,6"/>
                                </Grid>

                                <Grid Margin="0,0,0,6">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>

                                    <StackPanel Grid.Column="0" Margin="0,0,8,0">
                                        <TextBlock Text="Install Behavior" FontSize="12" Foreground="#94A3B8" Margin="0,0,0,4"/>
                                        <TextBox x:Name="txtInstallBehavior" Text="User" IsReadOnly="True"/>
                                    </StackPanel>

                                    <StackPanel Grid.Column="1" Margin="0,0,8,0">
                                        <TextBlock Text="Device Restart Behavior" FontSize="12" Foreground="#94A3B8" Margin="0,0,0,4"/>
                                        <TextBox Text="Determine behavior based on return codes" IsReadOnly="True"/>
                                    </StackPanel>

                                    <StackPanel Grid.Column="2">
                                        <TextBlock Text="Return Code for Success" FontSize="12" Foreground="#94A3B8" Margin="0,0,0,4"/>
                                        <TextBox Text="0 (Success)" IsReadOnly="True"/>
                                    </StackPanel>
                                </Grid>
                            </StackPanel>
                        </Border>

                        <!-- Step 2: Detection Rules -->
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <Grid Margin="0,0,0,10">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <StackPanel Grid.Column="0">
                                        <TextBlock Text="STEP 2: DETECTION RULES CONFIGURATION" FontSize="12" FontWeight="Bold" Foreground="#6264A7"/>
                                        <TextBlock Text="Rules format: 'Use a custom detection script'. Select the generated Detect-TeamsBackgrounds.ps1 file." FontSize="12" Foreground="#94A3B8" Margin="0,2,0,0"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="1" Orientation="Horizontal">
                                        <Button x:Name="btnCopyDetectionScript" Content="Copy Script" Margin="0,0,8,0"/>
                                        <Button x:Name="btnOpenDetectionFile" Content="Open Script File"/>
                                    </StackPanel>
                                </Grid>

                                <TextBox x:Name="txtDetectionScriptPreview" IsReadOnly="True" AcceptsReturn="True" 
                                         FontFamily="Consolas, Courier New" FontSize="12" Background="#0F172A"
                                         Height="180" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                                         Padding="10"/>
                            </StackPanel>
                        </Border>

                        <!-- Step 3: Requirements Guide -->
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <TextBlock Text="STEP 3: REQUIREMENTS &amp; ASSIGNMENT" FontSize="12" FontWeight="Bold" Foreground="#6264A7" Margin="0,0,0,8"/>
                                <TextBlock Text="- Operating System Architecture: 64-bit and 32-bit" FontSize="12" Foreground="#CBD5E1" Margin="0,2"/>
                                <TextBlock Text="- Minimum Operating System: Windows 10 1909 or higher" FontSize="12" Foreground="#CBD5E1" Margin="0,2"/>
                                <TextBlock Text="- Assignment: Assign as 'Required' to targeted Entra ID User / Device groups." FontSize="12" Foreground="#CBD5E1" Margin="0,2"/>
                            </StackPanel>
                        </Border>

                        <!-- Step 4: Enterprise Logging Guide -->
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <Grid Margin="0,0,0,10">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <StackPanel Grid.Column="0">
                                        <TextBlock Text="STEP 4: LOGGING &amp; DIAGNOSTICS" FontSize="12" FontWeight="Bold" Foreground="#6264A7"/>
                                        <TextBlock Text="Full runtime diagnostic logs are automatically written to this folder on target machines:" FontSize="12" Foreground="#94A3B8" Margin="0,2,0,0"/>
                                    </StackPanel>
                                    <Button x:Name="btnOpenLogDir" Grid.Column="1" Content="Open Log Folder" Padding="12,6"/>
                                </Grid>

                                <Grid Margin="0,4,0,3">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="130"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Grid.Column="0" Text="Log Directory:" FontSize="12" Foreground="#94A3B8"/>
                                    <TextBlock x:Name="lblLogDirPath" Grid.Column="1" Text="C:\Log\TeamsBackgrounds" FontSize="12" FontWeight="SemiBold" Foreground="#38BDF8"/>
                                </Grid>
                                <Grid Margin="0,2,0,2">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="130"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Grid.Column="0" Text="Install Log:" FontSize="12" Foreground="#94A3B8"/>
                                    <TextBlock x:Name="lblInstallLogPath" Grid.Column="1" Text="C:\Log\TeamsBackgrounds\Install.log" FontSize="12" Foreground="#CBD5E1"/>
                                </Grid>
                                <Grid Margin="0,2,0,2">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="130"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Grid.Column="0" Text="Detection Log:" FontSize="12" Foreground="#94A3B8"/>
                                    <TextBlock x:Name="lblDetectLogPath" Grid.Column="1" Text="C:\Log\TeamsBackgrounds\Detection.log" FontSize="12" Foreground="#CBD5E1"/>
                                </Grid>
                                <Grid Margin="0,2,0,2">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="130"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Grid.Column="0" Text="Uninstall Log:" FontSize="12" Foreground="#94A3B8"/>
                                    <TextBlock x:Name="lblUninstallLogPath" Grid.Column="1" Text="C:\Log\TeamsBackgrounds\Uninstall.log" FontSize="12" Foreground="#CBD5E1"/>
                                </Grid>
                            </StackPanel>
                        </Border>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>

        <!-- Bottom Action Bar -->
        <Border Grid.Row="2" Background="#1E293B" BorderBrush="#334155" BorderThickness="0,1,0,0" Padding="24,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Status & Progress -->
                <StackPanel Grid.Column="0" VerticalAlignment="Center" Margin="0,0,20,0">
                    <TextBlock x:Name="lblStatus" Text="Ready. Add background images and click 'Build .intunewin Package'." FontSize="13" Foreground="#94A3B8"/>
                    <ProgressBar x:Name="progressBar" Height="6" Margin="0,6,0,0" Background="#334155" Foreground="#10B981" Value="0" Visibility="Collapsed"/>
                </StackPanel>

                <!-- Build Button -->
                <Button x:Name="btnBuildPackage" Grid.Column="1" Content="Build .intunewin Package" Style="{StaticResource SuccessButton}" FontSize="14" Padding="24,10"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Helper to load XAML cleanly
[xml]$xml = $xaml
$reader = New-Object System.Xml.XmlNodeReader $xml
$Window = [System.Windows.Markup.XamlReader]::Load($reader)

# Map UI Controls
$txtAppName = $Window.FindName("txtAppName")
$txtPublisher = $Window.FindName("txtPublisher")
$txtVersion = $Window.FindName("txtVersion")
$txtDescription = $Window.FindName("txtDescription")
$chkNewTeams = $Window.FindName("chkNewTeams")
$chkClassicTeams = $Window.FindName("chkClassicTeams")
$rbAutoContext = $Window.FindName("rbAutoContext")
$rbUserContext = $Window.FindName("rbUserContext")
$rbSystemContext = $Window.FindName("rbSystemContext")
$txtOutputDir = $Window.FindName("txtOutputDir")
$btnBrowseOutput = $Window.FindName("btnBrowseOutput")

$btnAddImages = $Window.FindName("btnAddImages")
$btnAddImagesEmpty = $Window.FindName("btnAddImagesEmpty")
$btnRemoveSelected = $Window.FindName("btnRemoveSelected")
$btnClearAll = $Window.FindName("btnClearAll")
$lblImageCount = $Window.FindName("lblImageCount")
$lvImages = $Window.FindName("lvImages")
$pnlEmptyState = $Window.FindName("pnlEmptyState")

$MainTabs = $Window.FindName("MainTabs")
$tabIntuneGuidance = $Window.FindName("tabIntuneGuidance")
$lblPackagePath = $Window.FindName("lblPackagePath")
$btnOpenOutputDir = $Window.FindName("btnOpenOutputDir")
$btnCopyPackagePath = $Window.FindName("btnCopyPackagePath")

$txtInstallCommand = $Window.FindName("txtInstallCommand")
$txtUninstallCommand = $Window.FindName("txtUninstallCommand")
$btnCopyInstallCmd = $Window.FindName("btnCopyInstallCmd")
$btnCopyUninstallCmd = $Window.FindName("btnCopyUninstallCmd")
$txtInstallBehavior = $Window.FindName("txtInstallBehavior")

$txtDetectionScriptPreview = $Window.FindName("txtDetectionScriptPreview")
$btnCopyDetectionScript = $Window.FindName("btnCopyDetectionScript")
$btnOpenDetectionFile = $Window.FindName("btnOpenDetectionFile")

$lblStatus = $Window.FindName("lblStatus")
$progressBar = $Window.FindName("progressBar")
$btnBuildPackage = $Window.FindName("btnBuildPackage")

$txtLogDir = $Window.FindName("txtLogDir")
$btnOpenLogDir = $Window.FindName("btnOpenLogDir")
$lblLogDirPath = $Window.FindName("lblLogDirPath")
$lblInstallLogPath = $Window.FindName("lblInstallLogPath")
$lblDetectLogPath = $Window.FindName("lblDetectLogPath")
$lblUninstallLogPath = $Window.FindName("lblUninstallLogPath")

# Initialize default values
$txtOutputDir.Text = $DefaultOutputDir
$lvImages.ItemsSource = $Global:ImageItems
$Global:LastBuiltDetectionScriptPath = ""
$Global:LastBuiltPackagePath = ""

# Function to update UI image list counter & empty state
function Update-ImageCounter {
    $count = $Global:ImageItems.Count
    if ($count -eq 1) {
        $lblImageCount.Text = "1 image"
    } else {
        $lblImageCount.Text = "$count images"
    }

    if ($count -eq 0) {
        $pnlEmptyState.Visibility = [System.Windows.Visibility]::Visible
    } else {
        $pnlEmptyState.Visibility = [System.Windows.Visibility]::Collapsed
    }
}

# Function to add image files to collection
function Add-ImageFiles([string[]]$filePaths) {
    foreach ($path in $filePaths) {
        if (-not (Test-Path $path)) { continue }
        
        $fileName = [System.IO.Path]::GetFileName($path)
        
        # Check duplicate
        $existing = $Global:ImageItems | Where-Object { $_.FileName -eq $fileName }
        if ($existing) {
            continue
        }

        try {
            $fileInfo = New-Object System.IO.FileInfo($path)
            $sizeStr = ""
            if ($fileInfo.Length -gt 1MB) {
                $sizeStr = "{0:N2} MB" -f ($fileInfo.Length / 1MB)
            } else {
                $sizeStr = "{0:N0} KB" -f ($fileInfo.Length / 1KB)
            }

            # Extract dimensions using System.Drawing
            $img = [System.Drawing.Image]::FromFile($path)
            $w = $img.Width
            $h = $img.Height
            $img.Dispose()

            $resStr = "${w}x${h}"
            $isOptimal = ($w -eq 1920 -and $h -eq 1080)
            $aspectStr = if ($isOptimal) { "Optimal (16:9)" } else { "Custom (${w}x${h})" }

            # Create BitmapImage with CacheOption OnLoad so file isn't locked
            $bi = New-Object System.Windows.Media.Imaging.BitmapImage
            $bi.BeginInit()
            $bi.UriSource = New-Object System.Uri($path, [System.UriKind]::Absolute)
            $bi.DecodePixelWidth = 120
            $bi.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $bi.EndInit()
            $bi.Freeze()

            $item = New-Object TeamsImageItem
            $item.FileName = $fileName
            $item.FullPath = $path
            $item.Resolution = $resStr
            $item.FileSize = $sizeStr
            $item.AspectRatio = $aspectStr
            $item.IsRecommendedSize = $isOptimal
            $item.Thumbnail = $bi

            $Global:ImageItems.Add($item)
        } catch {
            [System.Windows.MessageBox]::Show("Failed to load image '$fileName': $_", "Image Load Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        }
    }
    Update-ImageCounter
}

# Image selection dialog
$SelectImagesAction = {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Select Teams Background Images"
    $dlg.Filter = "Image Files (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg|PNG Images (*.png)|*.png|JPEG Images (*.jpg;*.jpeg)|*.jpg;*.jpeg|All Files (*.*)|*.*"
    $dlg.Multiselect = $true
    
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Add-ImageFiles $dlg.FileNames
    }
}

$btnAddImages.Add_Click($SelectImagesAction)
$btnAddImagesEmpty.Add_Click($SelectImagesAction)

# Remove selected images
$btnRemoveSelected.Add_Click({
    $selected = @($lvImages.SelectedItems)
    if ($selected.Count -eq 0) { return }
    foreach ($item in $selected) {
        $Global:ImageItems.Remove($item)
    }
    Update-ImageCounter
})

# Clear all images
$btnClearAll.Add_Click({
    if ($Global:ImageItems.Count -eq 0) { return }
    $result = [System.Windows.MessageBox]::Show("Are you sure you want to remove all images from the list?", "Confirm Clear", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        $Global:ImageItems.Clear()
        Update-ImageCounter
    }
})

# Browse output directory
$btnBrowseOutput.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select Output Directory for .intunewin Package"
    $dlg.SelectedPath = $txtOutputDir.Text
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOutputDir.Text = $dlg.SelectedPath
    }
})

# Context radio button changes update guidance display
$rbAutoContext.Add_Checked({
    $txtInstallBehavior.Text = "System (or User)"
})
$rbUserContext.Add_Checked({
    $txtInstallBehavior.Text = "User"
})
$rbSystemContext.Add_Checked({
    $txtInstallBehavior.Text = "System"
})

# Function to ensure IntuneWinAppUtil.exe is available
function Get-IntuneWinAppUtilPath {
    $expectedPath = Join-Path $ToolsDir "IntuneWinAppUtil.exe"
    if (Test-Path $expectedPath) {
        return $expectedPath
    }

    # Check PATH
    $cmd = Get-Command "IntuneWinAppUtil.exe" -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }

    # Attempt download from official Microsoft GitHub repository
    $lblStatus.Text = "Downloading official IntuneWinAppUtil.exe from Microsoft..."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $url = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $url -OutFile $expectedPath -UseBasicParsing -TimeoutSec 30
        if (Test-Path $expectedPath) {
            return $expectedPath
        }
    } catch {
        Write-Warning "Auto-download failed: $_"
    }

    # If download fails, prompt user to locate it
    $result = [System.Windows.MessageBox]::Show("IntuneWinAppUtil.exe was not found. Would you like to browse and locate it manually?", "IntuneWinAppUtil Required", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Information)
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title = "Locate IntuneWinAppUtil.exe"
        $dlg.Filter = "Executable Files (*.exe)|*.exe"
        $dlg.FileName = "IntuneWinAppUtil.exe"
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Copy-Item -Path $dlg.FileName -Destination $expectedPath -Force
            return $expectedPath
        }
    }

    return $null
}

# Function to generate the Install script
function New-InstallScriptContent([string[]]$ImageNames, [bool]$TargetNew, [bool]$TargetClassic, [string]$DeployMode, [string]$Version, [string]$AppName, [string]$LogDir) {
    $imageArrayStr = ($ImageNames | ForEach-Object { "'$_'" }) -join ", "

    $code = @'
<#
.SYNOPSIS
    Installs corporate Microsoft Teams background images.
    Auto-generated by Teams Background Packager for Intune.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$AppName = "__APPNAME__"
$AppVersion = "__APPVERSION__"
$TargetLogDir = "__LOGDIR__"
$ConfiguredMode = "__DEPLOYMODE__" # "Auto", "User", "System"
$TargetNewTeams = __TARGETNEW__
$TargetClassicTeams = __TARGETCLASSIC__
$ImageFiles = @(__IMAGEFILES__)

# Initialize Enterprise Log Directory (Primary: C:\Log\<AppName>, Fallback: %LOCALAPPDATA%\Log\<AppName>)
$LogDir = $TargetLogDir
try {
    if (-not (Test-Path $LogDir)) {
        New-Item -Path $LogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
    # Test write access
    $testFile = Join-Path $LogDir ".write_test"
    [System.IO.File]::WriteAllText($testFile, "test")
    Remove-Item -Path $testFile -Force -ErrorAction SilentlyContinue

    # If running elevated or as SYSTEM, ensure Users have modify access to C:\Log
    $isElevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($isElevated -and (Test-Path "C:\Log")) {
        try {
            $acl = Get-Acl "C:\Log"
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Users", "Modify,Synchronize", "ContainerInherit,ObjectInherit", "None", "Allow")
            $acl.AddAccessRule($rule)
            Set-Acl -Path "C:\Log" -AclObject $acl -ErrorAction SilentlyContinue
        } catch {}
    }
} catch {
    $fallbackDir = "$env:ProgramData\Log\$AppName"
    $LogDir = $fallbackDir
    if (-not (Test-Path $LogDir)) {
        New-Item -Path $LogDir -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
    }
}

$LogFile = Join-Path $LogDir "Install.log"

function Write-Log([string]$message) {
    $stamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $entry = "[$stamp] $message"
    Write-Output $entry
    Add-Content -Path $LogFile -Value $entry -ErrorAction SilentlyContinue
}

# Determine runtime identity
$CurrentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$IsSystemIdentity = $CurrentIdentity.IsSystem -or ($env:USERNAME -match '\$$') -or ($CurrentIdentity.Name -match '(?i)SYSTEM')

# Auto-detect or fail-safe:
# If running as SYSTEM (Sandbox runner, Intune System Context, SCCM), NEVER deploy to systemprofile!
# Always deploy across all real user profiles in C:\Users (including WDAGUtilityAccount in Sandbox & Default)
if ($IsSystemIdentity) {
    $EffectiveMode = "AllProfiles"
    $ContextLabel = "SYSTEM Context (Multi-Profile Auto-Detect)"
} elseif ($ConfiguredMode -eq "System") {
    $EffectiveMode = "AllProfiles"
    $ContextLabel = "System Context (All Profiles Configured)"
} else {
    $EffectiveMode = "CurrentUser"
    $ContextLabel = "User Context ($env:USERNAME)"
}

Write-Log "========================================================"
Write-Log "Starting installation of $AppName (v$AppVersion)..."
Write-Log "Host Computer:      $env:COMPUTERNAME"
Write-Log "Executing Identity: $($CurrentIdentity.Name)"
Write-Log "Execution Context:  $ContextLabel"
if ($IsSystemIdentity -and ($ConfiguredMode -eq "User")) {
    Write-Log "NOTICE: Running as SYSTEM/Sandbox. Auto-routing deployment to all user profiles in C:\Users (including WDAGUtilityAccount & Default) so real users receive backgrounds."
}
Write-Log "Active Log File:    $LogFile"
Write-Log "Background Count:   $($ImageFiles.Count) images"
Write-Log "========================================================"

$SourceImagesFolder = Join-Path $PSScriptRoot "Images"
if (-not (Test-Path $SourceImagesFolder)) {
    Write-Log "ERROR: Source Images folder not found at $SourceImagesFolder"
    exit 1
}

function Deploy-ToFolder([string]$targetFolder, [string]$folderLabel) {
    if (-not (Test-Path $targetFolder)) {
        New-Item -Path $targetFolder -ItemType Directory -Force | Out-Null
        Write-Log "Created directory ($folderLabel): $targetFolder"
    }

    foreach ($img in $ImageFiles) {
        $sourcePath = Join-Path $SourceImagesFolder $img
        $destPath = Join-Path $targetFolder $img
        if (Test-Path $sourcePath) {
            $fInfo = Get-Item $sourcePath
            Copy-Item -Path $sourcePath -Destination $destPath -Force
            $sizeKb = [Math]::Round($fInfo.Length / 1KB, 1)
            Write-Log "Copied ($folderLabel): '$img' ($sizeKb KB) -> '$targetFolder'"
        } else {
            Write-Log "WARNING: Source image '$sourcePath' does not exist."
        }
    }
}

try {
    if ($EffectiveMode -eq "CurrentUser") {
        # Current User Context
        Write-Log "Deploying in User Context for: $env:USERNAME"

        if ($TargetNewTeams) {
            $newTeamsUploads = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Roaming\Microsoft\Teams\Backgrounds\Uploads"
            Deploy-ToFolder -targetFolder $newTeamsUploads -folderLabel "New Teams"
        }

        if ($TargetClassicTeams) {
            $classicTeamsUploads = "$env:APPDATA\Microsoft\Teams\Backgrounds\Uploads"
            Deploy-ToFolder -targetFolder $classicTeamsUploads -folderLabel "Classic Teams"
        }

        # Write User marker file
        $markerDir = "$env:LOCALAPPDATA\TeamsCustomBackgrounds"
        if (-not (Test-Path $markerDir)) { New-Item -Path $markerDir -ItemType Directory -Force | Out-Null }
        $manifest = @{
            AppName       = $AppName
            Version       = $AppVersion
            InstalledDate = (Get-Date).ToString("o")
            Context       = "User"
            Images        = $ImageFiles
            LogDir        = $LogDir
        }
        $manifest | ConvertTo-Json | Set-Content -Path (Join-Path $markerDir "installed.json") -Force
        Write-Log "Wrote user marker file to $markerDir\installed.json"

    } else {
        # Deploy across all user profiles in C:\Users (including Sandbox accounts like WDAGUtilityAccount and Default)
        Write-Log "Deploying across user profiles in C:\Users..."

        $userProfiles = Get-ChildItem -Path "C:\Users" -Directory -Force | Where-Object {
            $_.Name -notmatch '^(Public|All Users|Default User)$' -and
            -not ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -and
            ($_.Name -notmatch '^[\!\(\}]')
        }

        foreach ($profile in $userProfiles) {
            Write-Log "Processing profile: $($profile.FullName)"

            if ($TargetNewTeams) {
                $userNewTeams = Join-Path $profile.FullName "AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Roaming\Microsoft\Teams\Backgrounds\Uploads"
                Deploy-ToFolder -targetFolder $userNewTeams -folderLabel "Profile: $($profile.Name) - New Teams"
            }

            if ($TargetClassicTeams) {
                $userClassicTeams = Join-Path $profile.FullName "AppData\Roaming\Microsoft\Teams\Backgrounds\Uploads"
                Deploy-ToFolder -targetFolder $userClassicTeams -folderLabel "Profile: $($profile.Name) - Classic Teams"
            }

            # Drop user marker inside profile
            $userMarkerDir = Join-Path $profile.FullName "AppData\Local\TeamsCustomBackgrounds"
            if (-not (Test-Path $userMarkerDir)) { New-Item -Path $userMarkerDir -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null }
            $userManifest = @{
                AppName       = $AppName
                Version       = $AppVersion
                InstalledDate = (Get-Date).ToString("o")
                Context       = "System-Deployed"
                Images        = $ImageFiles
            }
            $userManifest | ConvertTo-Json | Set-Content -Path (Join-Path $userMarkerDir "installed.json") -Force -ErrorAction SilentlyContinue
        }

        # Write System marker file
        $markerDir = "$env:ProgramData\TeamsCustomBackgrounds"
        if (-not (Test-Path $markerDir)) { New-Item -Path $markerDir -ItemType Directory -Force | Out-Null }
        $manifest = @{
            AppName       = $AppName
            Version       = $AppVersion
            InstalledDate = (Get-Date).ToString("o")
            Context       = "System"
            Images        = $ImageFiles
            LogDir        = $LogDir
        }
        $manifest | ConvertTo-Json | Set-Content -Path (Join-Path $markerDir "installed.json") -Force
        Write-Log "Wrote system marker file to $markerDir\installed.json"
    }

    Write-Log "Installation completed successfully."
    exit 0
} catch {
    Write-Log "ERROR during installation: $_"
    exit 1
}
'@

    $targetNewStr = if ($TargetNew) { '$true' } else { '$false' }
    $targetClassicStr = if ($TargetClassic) { '$true' } else { '$false' }

    $code = $code.Replace("__APPNAME__", $AppName)
    $code = $code.Replace("__APPVERSION__", $Version)
    $code = $code.Replace("__LOGDIR__", $LogDir)
    $code = $code.Replace("__DEPLOYMODE__", $DeployMode)
    $code = $code.Replace("__TARGETNEW__", $targetNewStr)
    $code = $code.Replace("__TARGETCLASSIC__", $targetClassicStr)
    $code = $code.Replace("__IMAGEFILES__", $imageArrayStr)
    return $code
}

# Function to generate the Detection script
function New-DetectScriptContent([string[]]$ImageNames, [bool]$TargetNew, [bool]$TargetClassic, [string]$DeployMode, [string]$Version, [string]$AppName, [string]$LogDir) {
    $imageArrayStr = ($ImageNames | ForEach-Object { "'$_'" }) -join ", "

    $code = @'
<#
.SYNOPSIS
    Detection script for corporate Microsoft Teams background images.
    Returns 0 and outputs success text if detected; otherwise exits with 1.
    Auto-generated by Teams Background Packager for Intune.
#>
$ErrorActionPreference = 'SilentlyContinue'

$AppName = "__APPNAME__"
$AppVersion = "__APPVERSION__"
$TargetLogDir = "__LOGDIR__"
$ConfiguredMode = "__DEPLOYMODE__"
$TargetNewTeams = __TARGETNEW__
$TargetClassicTeams = __TARGETCLASSIC__
$ExpectedImages = @(__IMAGEFILES__)

# Initialize Log Directory
$LogDir = $TargetLogDir
try {
    if (-not (Test-Path $LogDir)) {
        New-Item -Path $LogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
} catch {
    $fallbackDir = "$env:ProgramData\Log\$AppName"
    $LogDir = $fallbackDir
    if (-not (Test-Path $LogDir)) {
        New-Item -Path $LogDir -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
    }
}

$LogFile = Join-Path $LogDir "Detection.log"

function Write-DetectLog([string]$message) {
    $stamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $entry = "[$stamp] [DETECTION] $message"
    Add-Content -Path $LogFile -Value $entry -ErrorAction SilentlyContinue
}

$CurrentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$IsSystemIdentity = $CurrentIdentity.IsSystem -or ($env:USERNAME -match '\$$') -or ($CurrentIdentity.Name -match '(?i)SYSTEM')

Write-DetectLog "Starting detection check for $AppName (v$AppVersion) under $($CurrentIdentity.Name)..."

function Test-AllImagesExist([string]$folder, [string]$folderLabel) {
    if (-not (Test-Path $folder)) {
        Write-DetectLog "Target folder not found ($folderLabel): $folder"
        return $false
    }
    $allPresent = $true
    foreach ($img in $ExpectedImages) {
        $imgPath = Join-Path $folder $img
        if (-not (Test-Path $imgPath)) {
            Write-DetectLog "Missing image ($folderLabel): $imgPath"
            $allPresent = $false
        } else {
            Write-DetectLog "Found image ($folderLabel): $img"
        }
    }
    return $allPresent
}

if ($IsSystemIdentity -or ($ConfiguredMode -eq "System")) {
    # Check if images exist in any human user profile under C:\Users (e.g. WDAGUtilityAccount in Sandbox) OR system marker
    Write-DetectLog "Running as SYSTEM / Sandbox identity. Checking user profiles under C:\Users..."
    
    $detectedInProfile = $false
    $userProfiles = Get-ChildItem -Path "C:\Users" -Directory -Force | Where-Object {
        $_.Name -notmatch '^(Public|All Users|Default User)$' -and
        -not ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -and
        ($_.Name -notmatch '^[\!\(\}]')
    }

    foreach ($profile in $userProfiles) {
        if ($TargetNewTeams) {
            $userNewTeams = Join-Path $profile.FullName "AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Roaming\Microsoft\Teams\Backgrounds\Uploads"
            if (Test-AllImagesExist -folder $userNewTeams -folderLabel "Profile: $($profile.Name) - New Teams") {
                $detectedInProfile = $true
                break
            }
        }
        if ($TargetClassicTeams) {
            $userClassicTeams = Join-Path $profile.FullName "AppData\Roaming\Microsoft\Teams\Backgrounds\Uploads"
            if (Test-AllImagesExist -folder $userClassicTeams -folderLabel "Profile: $($profile.Name) - Classic Teams") {
                $detectedInProfile = $true
                break
            }
        }
    }

    $systemMarker = "$env:ProgramData\TeamsCustomBackgrounds\installed.json"
    $hasSystemMarker = (Test-Path $systemMarker)

    if ($detectedInProfile -or $hasSystemMarker) {
        Write-DetectLog "STATUS: INSTALLED (Verified via user profile / system marker under SYSTEM context)."
        Write-Output "Teams Backgrounds detected successfully ($($ExpectedImages.Count) images present)."
        exit 0
    } else {
        Write-DetectLog "STATUS: NOT INSTALLED (Images missing from user profiles and system marker missing)."
        Write-Warning "Teams Backgrounds not detected in user profiles."
        exit 1
    }
} else {
    # Current User Context
    $detected = $false

    if ($TargetNewTeams) {
        $newTeamsPath = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Roaming\Microsoft\Teams\Backgrounds\Uploads"
        if (Test-AllImagesExist -folder $newTeamsPath -folderLabel "New Teams") {
            $detected = $true
        }
    }

    if (-not $detected -and $TargetClassicTeams) {
        $classicTeamsPath = "$env:APPDATA\Microsoft\Teams\Backgrounds\Uploads"
        if (Test-AllImagesExist -folder $classicTeamsPath -folderLabel "Classic Teams") {
            $detected = $true
        }
    }

    # Also check user manifest marker
    $markerFile = "$env:LOCALAPPDATA\TeamsCustomBackgrounds\installed.json"
    if (Test-Path $markerFile) {
        $marker = Get-Content $markerFile -Raw | ConvertFrom-Json
        if ($marker.Version -eq "__APPVERSION__") {
            Write-DetectLog "Manifest marker verified: version $($marker.Version)"
            $detected = $true
        }
    }

    if ($detected) {
        Write-DetectLog "STATUS: INSTALLED (All $($ExpectedImages.Count) images verified)."
        Write-Output "Teams Backgrounds detected successfully ($($ExpectedImages.Count) images present)."
        exit 0
    } else {
        Write-DetectLog "STATUS: NOT INSTALLED (Required background images missing)."
        Write-Warning "Teams Backgrounds not detected in user profile."
        exit 1
    }
}
'@

    $targetNewStr = if ($TargetNew) { '$true' } else { '$false' }
    $targetClassicStr = if ($TargetClassic) { '$true' } else { '$false' }

    $code = $code.Replace("__APPNAME__", $AppName)
    $code = $code.Replace("__APPVERSION__", $Version)
    $code = $code.Replace("__LOGDIR__", $LogDir)
    $code = $code.Replace("__DEPLOYMODE__", $DeployMode)
    $code = $code.Replace("__TARGETNEW__", $targetNewStr)
    $code = $code.Replace("__TARGETCLASSIC__", $targetClassicStr)
    $code = $code.Replace("__IMAGEFILES__", $imageArrayStr)
    return $code
}

# Function to generate the Uninstall script
function New-UninstallScriptContent([string[]]$ImageNames, [bool]$TargetNew, [bool]$TargetClassic, [string]$DeployMode, [string]$AppName, [string]$LogDir) {
    $imageArrayStr = ($ImageNames | ForEach-Object { "'$_'" }) -join ", "

    $code = @'
<#
.SYNOPSIS
    Uninstalls corporate Microsoft Teams background images.
    Auto-generated by Teams Background Packager for Intune.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'SilentlyContinue'
$AppName = "__APPNAME__"
$TargetLogDir = "__LOGDIR__"
$ConfiguredMode = "__DEPLOYMODE__"
$TargetNewTeams = __TARGETNEW__
$TargetClassicTeams = __TARGETCLASSIC__
$ImageFiles = @(__IMAGEFILES__)

# Initialize Log Directory
$LogDir = $TargetLogDir
try {
    if (-not (Test-Path $LogDir)) {
        New-Item -Path $LogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
} catch {
    $fallbackDir = "$env:ProgramData\Log\$AppName"
    $LogDir = $fallbackDir
    if (-not (Test-Path $LogDir)) {
        New-Item -Path $LogDir -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
    }
}

$LogFile = Join-Path $LogDir "Uninstall.log"

function Write-UninstallLog([string]$message) {
    $stamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $entry = "[$stamp] [UNINSTALL] $message"
    Write-Output $entry
    Add-Content -Path $LogFile -Value $entry -ErrorAction SilentlyContinue
}

$CurrentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
$IsSystemIdentity = $CurrentIdentity.IsSystem -or ($env:USERNAME -match '\$$') -or ($CurrentIdentity.Name -match '(?i)SYSTEM')

Write-UninstallLog "========================================================"
Write-UninstallLog "Starting uninstallation of $AppName..."
Write-UninstallLog "Executing Identity: $($CurrentIdentity.Name)"
Write-UninstallLog "Host Computer:      $env:COMPUTERNAME"
Write-UninstallLog "Active Log File:    $LogFile"
Write-UninstallLog "========================================================"

function Remove-ImagesFromFolder([string]$folder, [string]$folderLabel) {
    if (-not (Test-Path $folder)) { return }
    foreach ($img in $ImageFiles) {
        $path = Join-Path $folder $img
        if (Test-Path $path) {
            Remove-Item -Path $path -Force -ErrorAction SilentlyContinue
            Write-UninstallLog "Removed ($folderLabel): $path"
        }
    }
}

if ($IsSystemIdentity -or ($ConfiguredMode -eq "System")) {
    Write-UninstallLog "Removing from all user profiles in C:\Users..."
    $userProfiles = Get-ChildItem -Path "C:\Users" -Directory -Force | Where-Object {
        $_.Name -notmatch '^(Public|All Users|Default User)$' -and
        -not ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -and
        ($_.Name -notmatch '^[\!\(\}]')
    }

    foreach ($profile in $userProfiles) {
        if ($TargetNewTeams) {
            $p = Join-Path $profile.FullName "AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Roaming\Microsoft\Teams\Backgrounds\Uploads"
            Remove-ImagesFromFolder -folder $p -folderLabel "Profile: $($profile.Name) - New Teams"
        }
        if ($TargetClassicTeams) {
            $p = Join-Path $profile.FullName "AppData\Roaming\Microsoft\Teams\Backgrounds\Uploads"
            Remove-ImagesFromFolder -folder $p -folderLabel "Profile: $($profile.Name) - Classic Teams"
        }
        $userMarkerDir = Join-Path $profile.FullName "AppData\Local\TeamsCustomBackgrounds"
        if (Test-Path $userMarkerDir) {
            Remove-Item -Path $userMarkerDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    $markerDir = "$env:ProgramData\TeamsCustomBackgrounds"
    if (Test-Path $markerDir) {
        Remove-Item -Path $markerDir -Recurse -Force -ErrorAction SilentlyContinue
        Write-UninstallLog "Removed system deployment marker: $markerDir"
    }
} else {
    if ($TargetNewTeams) {
        $newTeamsUploads = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Roaming\Microsoft\Teams\Backgrounds\Uploads"
        Remove-ImagesFromFolder -folder $newTeamsUploads -folderLabel "New Teams"
    }

    if ($TargetClassicTeams) {
        $classicTeamsUploads = "$env:APPDATA\Microsoft\Teams\Backgrounds\Uploads"
        Remove-ImagesFromFolder -folder $classicTeamsUploads -folderLabel "Classic Teams"
    }

    $markerDir = "$env:LOCALAPPDATA\TeamsCustomBackgrounds"
    if (Test-Path $markerDir) {
        Remove-Item -Path $markerDir -Recurse -Force -ErrorAction SilentlyContinue
        Write-UninstallLog "Removed user deployment marker: $markerDir"
    }
}

Write-UninstallLog "Uninstallation completed successfully."
Write-Output "Teams Backgrounds removed successfully."
exit 0
'@

    $targetNewStr = if ($TargetNew) { '$true' } else { '$false' }
    $targetClassicStr = if ($TargetClassic) { '$true' } else { '$false' }

    $code = $code.Replace("__APPNAME__", $AppName)
    $code = $code.Replace("__LOGDIR__", $LogDir)
    $code = $code.Replace("__DEPLOYMODE__", $DeployMode)
    $code = $code.Replace("__TARGETNEW__", $targetNewStr)
    $code = $code.Replace("__TARGETCLASSIC__", $targetClassicStr)
    $code = $code.Replace("__IMAGEFILES__", $imageArrayStr)
    return $code
}

# Main Packaging Action
$btnBuildPackage.Add_Click({
    # Validate inputs
    if ($Global:ImageItems.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Please add at least one Teams background image before building the package.", "No Images Added", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }

    $appName = $txtAppName.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($appName)) {
        [System.Windows.MessageBox]::Show("Please enter an Application Name.", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }

    $targetNew = $chkNewTeams.IsChecked -eq $true
    $targetClassic = $chkClassicTeams.IsChecked -eq $true
    if (-not $targetNew -and -not $targetClassic) {
        [System.Windows.MessageBox]::Show("Please select at least one target Teams version (New Teams or Classic Teams).", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }

    $outputDir = $txtOutputDir.Text.Trim()
    if (-not (Test-Path $outputDir)) {
        try {
            New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
        } catch {
            [System.Windows.MessageBox]::Show("Failed to create output directory: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }
    }

    $deployMode = if ($rbSystemContext.IsChecked -eq $true) { "System" } elseif ($rbUserContext.IsChecked -eq $true) { "User" } else { "Auto" }
    $isUserContext = ($deployMode -eq "User")
    $version = $txtVersion.Text.Trim()
    if (-not $version) { $version = "1.0.0" }
    $publisher = $txtPublisher.Text.Trim()

    # Disable buttons during build
    $btnBuildPackage.IsEnabled = $false
    $btnAddImages.IsEnabled = $false
    $progressBar.Visibility = [System.Windows.Visibility]::Visible
    $progressBar.IsIndeterminate = $true
    $lblStatus.Text = "Checking Win32 Content Prep Tool (IntuneWinAppUtil.exe)..."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        # Check IntuneWinAppUtil
        $utilPath = Get-IntuneWinAppUtilPath
        if (-not $utilPath) {
            [System.Windows.MessageBox]::Show("IntuneWinAppUtil.exe is required to compile the .intunewin package. Build canceled.", "Missing Tool", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }

        # Create Staging Directory
        $safeName = ($appName -replace '[^\w\-]', '_')
        $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $stagingRoot = Join-Path $outputDir "staging_${safeName}_$timestamp"
        $stagingSource = Join-Path $stagingRoot "Source"
        $stagingImages = Join-Path $stagingSource "Images"

        New-Item -Path $stagingImages -ItemType Directory -Force | Out-Null

        $lblStatus.Text = "Copying background images to package source..."
        [System.Windows.Forms.Application]::DoEvents()

        $imageNames = @()
        foreach ($item in $Global:ImageItems) {
            $destFile = Join-Path $stagingImages $item.FileName
            Copy-Item -Path $item.FullPath -Destination $destFile -Force
            $imageNames += $item.FileName
        }

        $lblStatus.Text = "Generating installation, uninstallation, and detection scripts..."
        [System.Windows.Forms.Application]::DoEvents()

        # Generate scripts
        $logDir = $txtLogDir.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($logDir)) {
            $logDir = "C:\Log\$safeName"
            $txtLogDir.Text = $logDir
        }

        # Generate scripts with full enterprise logging
        $installScript   = New-InstallScriptContent   -ImageNames $imageNames -TargetNew $targetNew -TargetClassic $targetClassic -DeployMode $deployMode -Version $version -AppName $appName -LogDir $logDir
        $detectScript    = New-DetectScriptContent    -ImageNames $imageNames -TargetNew $targetNew -TargetClassic $targetClassic -DeployMode $deployMode -Version $version -AppName $appName -LogDir $logDir
        $uninstallScript = New-UninstallScriptContent -ImageNames $imageNames -TargetNew $targetNew -TargetClassic $targetClassic -DeployMode $deployMode -AppName $appName -LogDir $logDir

        # Save into Staging Source
        $installPath = Join-Path $stagingSource "Install-TeamsBackgrounds.ps1"
        $detectPath = Join-Path $stagingSource "Detect-TeamsBackgrounds.ps1"
        $uninstallPath = Join-Path $stagingSource "Uninstall-TeamsBackgrounds.ps1"

        Set-Content -Path $installPath -Value $installScript -Encoding UTF8 -Force
        Set-Content -Path $detectPath -Value $detectScript -Encoding UTF8 -Force
        Set-Content -Path $uninstallPath -Value $uninstallScript -Encoding UTF8 -Force

        # Also save detection script and install instructions directly to Output directory for easy access
        $outputDetectScript = Join-Path $outputDir "Detect-TeamsBackgrounds.ps1"
        Set-Content -Path $outputDetectScript -Value $detectScript -Encoding UTF8 -Force
        $Global:LastBuiltDetectionScriptPath = $outputDetectScript

        $lblStatus.Text = "Compiling Win32 application package (.intunewin)..."
        [System.Windows.Forms.Application]::DoEvents()

        # Run IntuneWinAppUtil.exe
        $processArgs = @(
            "-c", "`"$stagingSource`"",
            "-s", "`"Install-TeamsBackgrounds.ps1`"",
            "-o", "`"$outputDir`"",
            "-q"
        ) -join " "

        $pinfo = New-Object System.Diagnostics.ProcessStartInfo
        $pinfo.FileName = $utilPath
        $pinfo.Arguments = $processArgs
        $pinfo.RedirectStandardOutput = $true
        $pinfo.RedirectStandardError = $true
        $pinfo.UseShellExecute = $false
        $pinfo.CreateNoWindow = $true

        $process = [System.Diagnostics.Process]::Start($pinfo)
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()

        # Expected output package name is Install-TeamsBackgrounds.intunewin
        $defaultIntuneWin = Join-Path $outputDir "Install-TeamsBackgrounds.intunewin"
        $customIntuneWin = Join-Path $outputDir "${safeName}.intunewin"

        if (Test-Path $defaultIntuneWin) {
            # Rename to application-specific name
            if (Test-Path $customIntuneWin) { Remove-Item -Path $customIntuneWin -Force }
            Move-Item -Path $defaultIntuneWin -Destination $customIntuneWin -Force
            $Global:LastBuiltPackagePath = $customIntuneWin
        } elseif (Test-Path $customIntuneWin) {
            $Global:LastBuiltPackagePath = $customIntuneWin
        } else {
            # Search output directory for any newly created intunewin
            $recent = Get-ChildItem -Path $outputDir -Filter "*.intunewin" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($recent) {
                $Global:LastBuiltPackagePath = $recent.FullName
            } else {
                throw "Packaging failed. IntuneWinAppUtil output: $stdout $stderr"
            }
        }

        # Clean up staging directory
        try {
            Remove-Item -Path $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
        } catch {}

        # Save Intune Setup Instructions in output folder
        $behaviorStr = if ($isUserContext) { 'User' } else { 'System' }
        $instructionsPath = Join-Path $outputDir "Intune_Configuration_Guide.txt"
        $instructions = @"
========================================================================
MICROSOFT INTUNE WIN32 APP CONFIGURATION GUIDE
Teams Backgrounds Package: $appName
Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
========================================================================

1. APP INFORMATION
------------------------------------------------------------------------
Name:                   $appName
Description:            $($txtDescription.Text)
Publisher:              $publisher
App Version:            $version
Package File:           $($Global:LastBuiltPackagePath)

2. PROGRAM SETTINGS
------------------------------------------------------------------------
Install Command:
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "Install-TeamsBackgrounds.ps1"

Uninstall Command:
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "Uninstall-TeamsBackgrounds.ps1"

Install Behavior:       $behaviorStr
Device Restart Behavior: Determine behavior based on return codes
Return Code for Success: 0

3. DETECTION RULES
------------------------------------------------------------------------
Rules Format:           Use a custom detection script
Script File:            $outputDetectScript
Run script as 32-bit:   No
Enforce signature:      No

4. REQUIREMENTS
------------------------------------------------------------------------
OS Architecture:        32-bit, 64-bit
Minimum OS:             Windows 10 1909 or higher

5. LOGGING & DIAGNOSTICS (ENDPOINTS)
------------------------------------------------------------------------
Log Directory:          $logDir
Install Log:            $logDir\Install.log
Detection Log:          $logDir\Detection.log
Uninstall Log:          $logDir\Uninstall.log
Packager Log:           C:\Log\TeamsBackgroundPackager\Packager.log

========================================================================
"@
        Set-Content -Path $instructionsPath -Value $instructions -Encoding UTF8 -Force

        # Update Intune Guidance tab UI
        $lblPackagePath.Text = $Global:LastBuiltPackagePath
        $txtInstallCommand.Text = 'powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "Install-TeamsBackgrounds.ps1"'
        $txtUninstallCommand.Text = 'powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "Uninstall-TeamsBackgrounds.ps1"'
        $txtInstallBehavior.Text = if ($isUserContext) { "User" } else { "System" }
        $txtDetectionScriptPreview.Text = $detectScript
        $lblLogDirPath.Text = $logDir
        $lblInstallLogPath.Text = Join-Path $logDir "Install.log"
        $lblDetectLogPath.Text = Join-Path $logDir "Detection.log"
        $lblUninstallLogPath.Text = Join-Path $logDir "Uninstall.log"

        # Write Packager Tool Log
        try {
            $packagerLogDir = "C:\Log\TeamsBackgroundPackager"
            if (-not (Test-Path $packagerLogDir)) { New-Item -Path $packagerLogDir -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null }
            $packagerLogFile = Join-Path $packagerLogDir "Packager.log"
            $pStamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            Add-Content -Path $packagerLogFile -Value "[$pStamp] [BUILD_SUCCESS] Package: '$appName' (v$version) -> '$Global:LastBuiltPackagePath' | Images: $($imageNames.Count) | Endpoint LogDir: '$logDir'" -ErrorAction SilentlyContinue
        } catch {}

        $lblStatus.Text = "Success! Package created: $([System.IO.Path]::GetFileName($Global:LastBuiltPackagePath))"
        
        # Switch to Intune Guidance Tab
        $MainTabs.SelectedItem = $tabIntuneGuidance

        [System.Windows.MessageBox]::Show("Win32 package successfully created!`n`nFile: $([System.IO.Path]::GetFileName($Global:LastBuiltPackagePath))`n`nReview the 'Intune Deployment Details' tab for the exact command lines and detection rules.", "Package Built Successfully", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)

    } catch {
        $lblStatus.Text = "Error building package."
        [System.Windows.MessageBox]::Show("An error occurred during build:`n$_", "Build Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    } finally {
        $btnBuildPackage.IsEnabled = $true
        $btnAddImages.IsEnabled = $true
        $progressBar.Visibility = [System.Windows.Visibility]::Collapsed
        $progressBar.IsIndeterminate = $false
    }
})

# Copy buttons actions
$btnCopyInstallCmd.Add_Click({
    if ($txtInstallCommand.Text) {
        [System.Windows.Clipboard]::SetText($txtInstallCommand.Text)
        $lblStatus.Text = "Install command copied to clipboard!"
    }
})

$btnCopyUninstallCmd.Add_Click({
    if ($txtUninstallCommand.Text) {
        [System.Windows.Clipboard]::SetText($txtUninstallCommand.Text)
        $lblStatus.Text = "Uninstall command copied to clipboard!"
    }
})

$btnCopyDetectionScript.Add_Click({
    if ($txtDetectionScriptPreview.Text) {
        [System.Windows.Clipboard]::SetText($txtDetectionScriptPreview.Text)
        $lblStatus.Text = "Detection script copied to clipboard!"
    }
})

$btnCopyPackagePath.Add_Click({
    if ($Global:LastBuiltPackagePath) {
        [System.Windows.Clipboard]::SetText($Global:LastBuiltPackagePath)
        $lblStatus.Text = "Package path copied to clipboard!"
    }
})

$btnOpenOutputDir.Add_Click({
    $dir = $txtOutputDir.Text
    if (Test-Path $dir) {
        Start-Process explorer.exe -ArgumentList "`"$dir`""
    }
})

$btnOpenDetectionFile.Add_Click({
    if ($Global:LastBuiltDetectionScriptPath -and (Test-Path $Global:LastBuiltDetectionScriptPath)) {
        Start-Process notepad.exe -ArgumentList "`"$Global:LastBuiltDetectionScriptPath`""
    }
})

$btnOpenLogDir.Add_Click({
    $dir = $txtLogDir.Text.Trim()
    if (-not (Test-Path $dir)) {
        New-Item -Path $dir -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
    }
    if (Test-Path $dir) {
        Start-Process explorer.exe -ArgumentList "`"$dir`""
    }
})

# Auto-update Log Directory when Application Name changes
$txtAppName.Add_TextChanged({
    $safe = ($txtAppName.Text.Trim() -replace '[^\w\-]', '_')
    if (-not $safe) { $safe = "TeamsBackgrounds" }
    $txtLogDir.Text = "C:\Log\$safe"
})

# Launch GUI
$Window.ShowDialog() | Out-Null
