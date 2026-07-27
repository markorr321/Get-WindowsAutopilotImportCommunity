<#PSScriptInfo
.VERSION 1.3.0
.GUID 6f2b9c14-8d3e-4a71-9c5f-1b0e7a4d2c98
.AUTHOR Mark Orr
.COMPANYNAME orr365.tools
.COPYRIGHT (c) 2026 Mark Orr. MIT License.
.TAGS Windows Autopilot Intune EntraID DevicePreparation GUI WPF PowerShell OOBE
.LICENSEURI https://github.com/markorr321/Get-WindowsAutopilotImportCommunity/blob/main/LICENSE
.PROJECTURI https://github.com/markorr321/Get-WindowsAutopilotImportCommunity
.RELEASENOTES
1.3.0 Single-file build. Autopilot v1 and v2 (Device Preparation) support.
#>

<#
.SYNOPSIS
A graphical front end for Windows Autopilot device registration, supporting both Autopilot v1
(hardware hash) and Autopilot v2 (Device Preparation identifiers).

.DESCRIPTION
Registers a Windows device for Autopilot from a resizable dark-themed window, driving the
Windows Autopilot Community script by Andrew S Taylor.

Supports both registration modes: Autopilot v1 uploads the 4K hardware hash, and Autopilot v2
imports the Manufacturer,Model,Serial device identifier used by Device Preparation policies,
which needs no hardware hash and therefore works on virtual machines.

Runs from a single self-contained file with nothing to download first. The window, the theme
and the Autopilot engine are all embedded, so it works during OOBE on a restricted network.

Includes live staged progress with a working cancel button, an offline CSV export of both the
hardware hash and the device identifier, a concurrent network
prerequisite check across the documented Autopilot and Intune endpoints, Autopilot
diagnostics, and a full session log.

Requires Windows PowerShell 5.1 and administrator rights.

.PARAMETER GroupTag
Pre-fills the group tag field.

.PARAMETER AssignedUser
Pre-fills the assigned user UPN field.

.PARAMETER Mode
Pre-selects the registration mode: v1 (hardware hash) or v2 (device preparation).

.PARAMETER NoElevate
Skip the automatic elevation prompt. The hardware hash cannot be read without
administrator rights, so Autopilot v1 will not work in this state.

.EXAMPLE
.\Get-WindowsAutopilotImportGUICommunity.ps1

.EXAMPLE
.\Get-WindowsAutopilotImportGUICommunity.ps1 -GroupTag FINANCE -Mode v2

.NOTES
Author  : Mark Orr (@markorr321)
Website : https://orr365.tools
License : MIT

GENERATED FILE. Do not edit by hand: change the sources under src\ and re-run build.ps1.
Built 2026-07-27 06:14:33 with engine v5.0.16.

Autopilot engine: get-windowsautopilotinfocommunity.ps1 (c) Andrew S Taylor, MIT, embedded
unmodified with its Authenticode signature intact.
Inspired by AutoPilot_Import_GUI (c) 2023 Ugur Koc, MIT.

.LINK
https://orr365.tools

.LINK
https://github.com/andrew-s-taylor/WindowsAutopilotInfo
#>
[CmdletBinding()]
param(
    [string]$GroupTag = '',
    [string]$AssignedUser = '',
    [ValidateSet('v1', 'v2')][string]$Mode = '',
    [switch]$NoElevate
)

$ErrorActionPreference = 'Stop'

# In the single-file build there is no src\ tree; the app root is only used for the
# optional config.json sitting next to this script.
$script:ApAppRoot = $PSScriptRoot
if (-not $script:ApAppRoot) {
    $script:ApAppRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml

#region embedded resources
$script:ApEmbeddedXaml = @{}
$script:ApEmbeddedXaml['MainWindow'] = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Autopilot Import GUI (Community)"
        Width="1140" Height="800" MinWidth="960" MinHeight="620"
        WindowStartupLocation="CenterScreen"
        Background="#161616" Foreground="#FFFFFF"
        FontFamily="Segoe UI Variable, Segoe UI"
        UseLayoutRounding="True" SnapsToDevicePixels="True">

  <!--
    Autopilot Import GUI (Community) main window.

    Replaces the original's fixed 399x636 canvas of absolutely positioned controls with a
    resizable Grid/DockPanel shell: sidebar navigation on the left, one visible page on the
    right, a persistent status and progress strip along the bottom.

    Loaded by XamlReader at runtime, so there is deliberately no x:Class and no event
    handler attributes. Every interactive element carries an x:Name and is wired up in
    Show-AutopilotImportGui.ps1. The THEME token below is replaced with the contents of
    Themes/Dark.xaml before parsing, because StaticResource only resolves against
    resources already present when the tree is built.
  -->

  <Window.Resources>
    <!-- @THEME@ -->
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- ==================== header ==================== -->
    <Border Grid.Row="0" Background="#1F1F1F" BorderBrush="#2A2A2A" BorderThickness="0,0,0,1" Padding="20,12">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
          <Ellipse Width="10" Height="10" Fill="{StaticResource AccentBrush}" VerticalAlignment="Center"/>
          <TextBlock Text="Autopilot Import GUI" FontSize="17" FontWeight="SemiBold" Margin="10,0,0,0" VerticalAlignment="Center"/>
          <TextBlock Text="Community" FontSize="17" Foreground="{StaticResource TextMuted}" Margin="6,0,0,0" VerticalAlignment="Center"/>
        </StackPanel>

        <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
          <Border x:Name="PillAdmin" Style="{StaticResource Pill}" Margin="0,0,8,0">
            <TextBlock x:Name="PillAdminText" Style="{StaticResource PillText}" Text="Checking rights"/>
          </Border>
          <Border x:Name="PillNetwork" Style="{StaticResource Pill}" Margin="0,0,14,0">
            <TextBlock x:Name="PillNetworkText" Style="{StaticResource PillText}" Text="Checking network"/>
          </Border>
          <TextBlock x:Name="HeaderClock" Foreground="{StaticResource TextMuted}" FontSize="12" VerticalAlignment="Center" Text="--:--:--"/>
        </StackPanel>
      </Grid>
    </Border>

    <!-- ==================== body ==================== -->
    <Grid Grid.Row="1">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="196"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <!-- sidebar -->
      <Border Grid.Column="0" Background="#161616" BorderBrush="#2A2A2A" BorderThickness="0,0,1,0">
        <DockPanel LastChildFill="True">
          <!-- Persistent credit, in the same spot the original GUI put its author line. -->
          <StackPanel DockPanel.Dock="Bottom" Margin="16,0,16,14">
            <Border Style="{StaticResource Divider}" Margin="0,0,0,10"/>
            <TextBlock Text="Mark Orr" FontSize="12" FontWeight="SemiBold" Foreground="{StaticResource TextSecondary}"/>
            <TextBlock Text="@markorr321" Style="{StaticResource HintText}" Margin="0,2,0,0"/>
            <Button x:Name="SidebarSiteLink" Style="{StaticResource LinkButton}" Content="orr365.tools" Margin="0,5,0,0"/>
            <Border Style="{StaticResource Divider}" Margin="0,10,0,10"/>
            <TextBlock x:Name="SidebarEngineVersion" Style="{StaticResource HintText}" Margin="0" Text="Engine: resolving"/>
            <TextBlock x:Name="SidebarAppVersion" Style="{StaticResource HintText}" Margin="0,4,0,0" Text=""/>
          </StackPanel>

          <StackPanel DockPanel.Dock="Top" Margin="0,12,0,0">
            <RadioButton x:Name="NavRegister" Style="{StaticResource NavItem}" GroupName="Nav" Content="Register" IsChecked="True"/>
            <RadioButton x:Name="NavDevice"   Style="{StaticResource NavItem}" GroupName="Nav" Content="Device"/>
            <RadioButton x:Name="NavNetwork"  Style="{StaticResource NavItem}" GroupName="Nav" Content="Network check"/>
            <RadioButton x:Name="NavAdvanced" Style="{StaticResource NavItem}" GroupName="Nav" Content="Advanced"/>
            <RadioButton x:Name="NavLogs"     Style="{StaticResource NavItem}" GroupName="Nav" Content="Logs"/>
          </StackPanel>
        </DockPanel>
      </Border>

      <!-- pages -->
      <Grid Grid.Column="1">

        <!-- ================ Register ================ -->
        <Grid x:Name="PageRegister" Margin="24,20,24,16">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <!-- The form scrolls; the action buttons and the live output stay pinned so a
                 tech never has to scroll to find the Register button or the progress. -->
            <RowDefinition Height="*" MinHeight="120"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <StackPanel Grid.Row="0" Margin="0,0,0,16">
            <TextBlock Style="{StaticResource PageTitle}" Text="Register this device"/>
            <TextBlock Style="{StaticResource PageSubtitle}" Text="Signs in to your tenant and registers this machine for Windows Autopilot."/>
          </StackPanel>

          <!-- form + device summary -->
          <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Padding="0,0,6,0">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="336"/>
            </Grid.ColumnDefinitions>

            <StackPanel Grid.Column="0" Margin="0,0,16,0">

              <!-- mode -->
              <Border Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource CardTitle}" Text="Registration mode"/>
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="10"/>
                      <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <RadioButton x:Name="ModeV1" Grid.Column="0" Style="{StaticResource Segment}" GroupName="Mode" IsChecked="True">
                      <StackPanel>
                        <TextBlock Text="Autopilot v1" FontWeight="SemiBold" FontSize="13" Foreground="{Binding Foreground, RelativeSource={RelativeSource AncestorType=RadioButton}}"/>
                        <TextBlock Text="Hardware hash" FontSize="11" Opacity="0.75" Foreground="{Binding Foreground, RelativeSource={RelativeSource AncestorType=RadioButton}}"/>
                      </StackPanel>
                    </RadioButton>
                    <RadioButton x:Name="ModeV2" Grid.Column="2" Style="{StaticResource Segment}" GroupName="Mode">
                      <StackPanel>
                        <TextBlock Text="Device Preparation" FontWeight="SemiBold" FontSize="13" Foreground="{Binding Foreground, RelativeSource={RelativeSource AncestorType=RadioButton}}"/>
                        <TextBlock Text="v2 identifier" FontSize="11" Opacity="0.75" Foreground="{Binding Foreground, RelativeSource={RelativeSource AncestorType=RadioButton}}"/>
                      </StackPanel>
                    </RadioButton>
                  </Grid>
                  <TextBlock x:Name="ModeDescription" Style="{StaticResource HintText}" Margin="2,10,0,0"
                             Text="Uploads the 4K hardware hash and assigns a deployment profile."/>
                </StackPanel>
              </Border>

              <!-- details -->
              <Border x:Name="CardDetails" Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource CardTitle}" Text="Registration details"/>

                  <!-- Wrapped so the whole field can be hidden in Device Preparation mode,
                       where the engine ignores the group tag entirely. -->
                  <StackPanel x:Name="GroupTagSection">
                    <TextBlock Style="{StaticResource FieldLabel}" Text="GROUP TAG (OPTIONAL)"/>
                    <ComboBox x:Name="GroupTagCombo" Style="{StaticResource DarkComboBox}" IsEditable="True" MaxDropDownHeight="220"/>
                    <TextBlock x:Name="GroupTagHint" Style="{StaticResource HintText}" Text="Previously used tags are remembered and offered in the list."/>
                  </StackPanel>

                  <TextBlock Style="{StaticResource FieldLabel}" Text="ASSIGNED USER UPN (OPTIONAL)" Margin="2,14,0,6"/>
                  <TextBox x:Name="AssignedUserBox" Height="38"/>

                  <TextBlock Style="{StaticResource FieldLabel}" Text="COMPUTER NAME (OPTIONAL)" Margin="2,14,0,6"/>
                  <TextBox x:Name="ComputerNameBox" Height="38"/>

                  <TextBlock Style="{StaticResource FieldLabel}" Text="ADD TO ENTRA GROUP (OPTIONAL)" Margin="2,14,0,6"/>
                  <TextBox x:Name="AddToGroupBox" Height="38"/>
                  <TextBlock Style="{StaticResource HintText}" Text="Separate multiple group names with commas."/>
                </StackPanel>
              </Border>

              <!-- options -->
              <Border x:Name="CardOptions" Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource CardTitle}" Text="Options"/>

                  <!-- Autopilot v1 only: all of these map to engine switches that the
                       -identifier code path never reads. -->
                  <StackPanel x:Name="OptionsV1Section">
                    <CheckBox x:Name="WaitAssignCheck" Content="Wait for the deployment profile to be assigned"/>
                    <CheckBox x:Name="RebootCheck" Content="Restart this device once the profile is assigned"/>

                    <TextBlock Style="{StaticResource FieldLabel}" Text="IF THIS SERIAL IS ALREADY REGISTERED" Margin="2,14,0,6"/>
                    <RadioButton x:Name="PolicyUpdate" Style="{StaticResource ChoiceRadio}" GroupName="Existing" IsChecked="True" Content="Update its group tag"/>
                    <RadioButton x:Name="PolicyDelete" Style="{StaticResource ChoiceRadio}" GroupName="Existing" Content="Delete from Autopilot, Intune and Entra ID, then re-add"/>
                    <RadioButton x:Name="PolicySkip" Style="{StaticResource ChoiceRadio}" GroupName="Existing" Content="Assume it is new (skips the lookup, much faster on large tenants)"/>
                  </StackPanel>

                  <!-- Device Preparation only. This restart is performed by this tool, not by
                       the engine: the engine's -Reboot sits inside its assignment wait, which
                       the identifier path never reaches. -->
                  <StackPanel x:Name="OptionsV2Section" Visibility="Collapsed">
                    <CheckBox x:Name="RebootV2Check" Content="Restart this device after the identifier is imported"/>
                    <TextBlock Style="{StaticResource HintText}"
                               Text="Device Preparation targets devices through the Entra group on the policy. Restart only if this device is already a member of that group, otherwise it will return to OOBE before the policy can apply."/>

                    <TextBlock Style="{StaticResource HintText}" Margin="2,10,0,0"
                               Text="To restart later instead, once the device is in that group, use Restart now beside the register button."/>
                  </StackPanel>
                </StackPanel>
              </Border>
            </StackPanel>

            <!-- device summary -->
            <StackPanel Grid.Column="1">
              <Border Style="{StaticResource Card}">
                <StackPanel>
                  <TextBlock Style="{StaticResource CardTitle}" Text="This device"/>
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="Auto"/>
                      <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                      <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <TextBlock Grid.Row="0" Grid.Column="0" Style="{StaticResource InfoLabel}" Text="Serial" Margin="0,0,12,7"/>
                    <TextBlock x:Name="SummarySerial" Grid.Row="0" Grid.Column="1" Style="{StaticResource InfoValue}" Margin="0,0,0,7" Text="..."/>

                    <TextBlock Grid.Row="1" Grid.Column="0" Style="{StaticResource InfoLabel}" Text="Model" Margin="0,0,12,7"/>
                    <TextBlock x:Name="SummaryModel" Grid.Row="1" Grid.Column="1" Style="{StaticResource InfoValue}" Margin="0,0,0,7" Text="..."/>

                    <TextBlock Grid.Row="2" Grid.Column="0" Style="{StaticResource InfoLabel}" Text="Maker" Margin="0,0,12,7"/>
                    <TextBlock x:Name="SummaryManufacturer" Grid.Row="2" Grid.Column="1" Style="{StaticResource InfoValue}" Margin="0,0,0,7" Text="..."/>

                    <TextBlock Grid.Row="3" Grid.Column="0" Style="{StaticResource InfoLabel}" Text="Free" Margin="0,0,12,0"/>
                    <TextBlock x:Name="SummaryFreeSpace" Grid.Row="3" Grid.Column="1" Style="{StaticResource InfoValue}" Text="..."/>
                  </Grid>

                  <Border Style="{StaticResource Divider}"/>

                  <TextBlock Style="{StaticResource FieldLabel}" Text="READINESS"/>
                  <WrapPanel>
                    <Border x:Name="PillHash" Style="{StaticResource Pill}" Margin="0,0,6,6">
                      <TextBlock x:Name="PillHashText" Style="{StaticResource PillText}" Text="Hash"/>
                    </Border>
                    <Border x:Name="PillTpm" Style="{StaticResource Pill}" Margin="0,0,6,6">
                      <TextBlock x:Name="PillTpmText" Style="{StaticResource PillText}" Text="TPM"/>
                    </Border>
                    <Border x:Name="PillSecureBoot" Style="{StaticResource Pill}" Margin="0,0,6,6">
                      <TextBlock x:Name="PillSecureBootText" Style="{StaticResource PillText}" Text="Secure Boot"/>
                    </Border>
                  </WrapPanel>

                  <TextBlock x:Name="IdentifierPreviewLabel" Style="{StaticResource FieldLabel}" Text="DEVICE IDENTIFIER" Margin="2,10,0,6" Visibility="Collapsed"/>
                  <TextBox x:Name="IdentifierPreviewBox" Style="{StaticResource OutputBox}" Height="46" Visibility="Collapsed" TextWrapping="Wrap"/>
                </StackPanel>
              </Border>

              <Border x:Name="NoticeCard" Style="{StaticResource Card}" Visibility="Collapsed"
                      Background="#1E1A12" BorderBrush="#4A3A16">
                <StackPanel>
                  <TextBlock x:Name="NoticeTitle" Style="{StaticResource CardTitle}" Foreground="{StaticResource WarningBrushLight}" Text="Before you continue" Margin="0,0,0,8"/>
                  <TextBlock x:Name="NoticeText" Foreground="{StaticResource TextSecondary}" FontSize="12" TextWrapping="Wrap" Text=""/>
                </StackPanel>
              </Border>
            </StackPanel>
          </Grid>
          </ScrollViewer>

          <!-- Actions. This row is pinned outside the scrolling form, so anything here is
               always reachable. The on-demand restart lives here rather than in the Options
               card for exactly that reason: in the card it sat below the fold. -->
          <Grid Grid.Row="2" Margin="0,12,0,12">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="RegisterButton" Grid.Column="0" Style="{StaticResource PrimaryButton}" Content="REGISTER THIS DEVICE" Margin="0,0,10,0"/>
            <Button x:Name="PreviewButton" Grid.Column="1" Style="{StaticResource SecondaryButton}" Content="Preview command" Height="48" Margin="0,0,10,0"/>
            <Button x:Name="CancelButton" Grid.Column="2" Style="{StaticResource SecondaryButton}" Content="Cancel" Height="48" IsEnabled="False"/>
            <!-- Device Preparation only; shown by Sync-ApModeUi. -->
            <Button x:Name="RestartNowButton" Grid.Column="3" Style="{StaticResource DangerButton}" Content="Restart now" Height="48" Margin="10,0,0,0" Visibility="Collapsed"/>
          </Grid>

          <!-- live output -->
          <Grid Grid.Row="3">
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock Grid.Row="0" Style="{StaticResource FieldLabel}" Text="LIVE OUTPUT"/>
            <TextBox x:Name="RegisterOutput" Grid.Row="1" Style="{StaticResource OutputBox}" Height="132"/>
          </Grid>
        </Grid>

        <!-- ================ Device ================ -->
        <Grid x:Name="PageDevice" Margin="24,20,24,16" Visibility="Collapsed">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <StackPanel Grid.Row="0" Margin="0,0,0,16">
            <TextBlock Style="{StaticResource PageTitle}" Text="Device"/>
            <TextBlock Style="{StaticResource PageSubtitle}" Text="Hardware and OS details, plus offline export of the Autopilot identity. Nothing here contacts your tenant."/>
          </StackPanel>

          <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="336"/>
              </Grid.ColumnDefinitions>

              <Border Grid.Column="0" Style="{StaticResource Card}" Margin="0,0,16,14">
                <StackPanel>
                  <TextBlock Style="{StaticResource CardTitle}" Text="Inventory"/>
                  <DataGrid x:Name="DeviceGrid" MaxHeight="420" BorderThickness="0" HeadersVisibility="None">
                    <DataGrid.Columns>
                      <DataGridTextColumn Header="Property" Binding="{Binding Name}" Width="190"/>
                      <DataGridTextColumn Header="Value" Binding="{Binding Value}" Width="*"/>
                    </DataGrid.Columns>
                  </DataGrid>
                </StackPanel>
              </Border>

              <StackPanel Grid.Column="1">
                <Border Style="{StaticResource Card}">
                  <StackPanel>
                    <TextBlock Style="{StaticResource CardTitle}" Text="Export"/>
                    <Button x:Name="ExportHashButton" Style="{StaticResource SecondaryButton}" Content="Export hardware hash CSV (v1)" HorizontalAlignment="Stretch" Margin="0,0,0,8"/>
                    <Button x:Name="ExportIdentifierButton" Style="{StaticResource SecondaryButton}" Content="Export device identifier CSV (v2)" HorizontalAlignment="Stretch" Margin="0,0,0,8"/>
                    <Button x:Name="CopyIdentifierButton" Style="{StaticResource SecondaryButton}" Content="Copy identifier to clipboard" HorizontalAlignment="Stretch" Margin="0,0,0,8"/>
                    <Button x:Name="CopyHashButton" Style="{StaticResource SecondaryButton}" Content="Copy hardware hash to clipboard" HorizontalAlignment="Stretch"/>
                    <CheckBox x:Name="ExportAppendCheck" Content="Append to the file if it already exists" Margin="0,12,0,0"/>
                    <CheckBox x:Name="ExportPartnerCheck" Content="Use the partner CSV format"/>
                    <TextBlock Style="{StaticResource HintText}" Text="The partner format adds manufacturer and model columns but has no group tag or assigned user column."/>
                  </StackPanel>
                </Border>

                <Border Style="{StaticResource Card}">
                  <StackPanel>
                    <TextBlock Style="{StaticResource CardTitle}" Text="Refresh"/>
                    <Button x:Name="RefreshDeviceButton" Style="{StaticResource SecondaryButton}" Content="Re-read device information" HorizontalAlignment="Stretch"/>
                  </StackPanel>
                </Border>
              </StackPanel>
            </Grid>
          </ScrollViewer>
        </Grid>

        <!-- ================ Network ================ -->
        <Grid x:Name="PageNetwork" Margin="24,20,24,16" Visibility="Collapsed">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <StackPanel Grid.Row="0" Margin="0,0,0,16">
            <TextBlock Style="{StaticResource PageTitle}" Text="Network check"/>
            <TextBlock Style="{StaticResource PageSubtitle}" Text="Tests the endpoints Windows Autopilot and Intune need on TCP 443."/>
          </StackPanel>

          <Grid Grid.Row="1" Margin="0,0,0,14">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="NetworkRunButton" Grid.Column="0" Style="{StaticResource PrimaryButton}" Content="RUN CHECK" Height="40" Margin="0,0,10,0"/>
            <Button x:Name="NetworkCopyButton" Grid.Column="1" Style="{StaticResource SecondaryButton}" Content="Copy results" Height="40" Margin="0,0,16,0"/>
            <TextBlock x:Name="NetworkSummary" Grid.Column="2" VerticalAlignment="Center" Foreground="{StaticResource TextMuted}" FontSize="12" Text="Not run yet."/>
          </Grid>

          <DataGrid x:Name="NetworkGrid" Grid.Row="2">
            <DataGrid.Columns>
              <DataGridTextColumn Header="STATUS" Binding="{Binding Status}" Width="110">
                <DataGridTextColumn.ElementStyle>
                  <Style TargetType="TextBlock">
                    <Setter Property="FontWeight" Value="SemiBold"/>
                    <Setter Property="VerticalAlignment" Value="Center"/>
                    <Setter Property="Foreground" Value="#C0C0C0"/>
                    <Style.Triggers>
                      <!-- Reachable is green; a failing *required* endpoint is red, while a
                           failing optional one is amber so it reads as informational. -->
                      <DataTrigger Binding="{Binding Status}" Value="OK">
                        <Setter Property="Foreground" Value="#3BC77A"/>
                      </DataTrigger>
                      <DataTrigger Binding="{Binding Status}" Value="Failed">
                        <Setter Property="Foreground" Value="#F03A47"/>
                      </DataTrigger>
                      <DataTrigger Binding="{Binding Status}" Value="Unreachable">
                        <Setter Property="Foreground" Value="#E9B44C"/>
                      </DataTrigger>
                    </Style.Triggers>
                  </Style>
                </DataGridTextColumn.ElementStyle>
              </DataGridTextColumn>
              <DataGridTextColumn Header="ENDPOINT" Binding="{Binding Name}" Width="190"/>
              <DataGridTextColumn Header="HOST" Binding="{Binding Host}" Width="*"/>
              <DataGridTextColumn Header="CATEGORY" Binding="{Binding Category}" Width="180"/>
              <DataGridTextColumn Header="REQUIRED" Binding="{Binding Required}" Width="80"/>
              <DataGridTextColumn Header="MS" Binding="{Binding LatencyMs}" Width="60"/>
            </DataGrid.Columns>
          </DataGrid>
        </Grid>

        <!-- ================ Advanced ================ -->
        <Grid x:Name="PageAdvanced" Margin="24,20,24,16" Visibility="Collapsed">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <StackPanel Grid.Row="0" Margin="0,0,0,16">
            <TextBlock Style="{StaticResource PageTitle}" Text="Advanced"/>
            <TextBlock Style="{StaticResource PageSubtitle}" Text="Post-registration actions and maintenance tools. These apply to Autopilot v1 registrations only."/>
          </StackPanel>

          <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="336"/>
              </Grid.ColumnDefinitions>

              <StackPanel Grid.Column="0" Margin="0,0,16,0">
                <Border Style="{StaticResource Card}">
                  <StackPanel>
                    <TextBlock Style="{StaticResource CardTitle}" Text="After the profile is assigned"/>
                    <TextBlock Style="{StaticResource HintText}" Margin="2,0,0,10"
                               Text="The community script performs all of these only after it has waited for profile assignment, so selecting any of them turns that wait on automatically."/>
                    <CheckBox x:Name="AdvPreProvisionCheck" Content="Start pre-provisioning (presses the Windows key five times)"/>
                    <CheckBox x:Name="AdvSysprepCheck" Content="Run sysprep /oobe /reboot"/>
                    <CheckBox x:Name="AdvWipeCheck" Content="Send an Intune wipe to this device"/>
                    <TextBlock x:Name="AdvWipeWarning" Foreground="{StaticResource ErrorBrush}" FontSize="11" TextWrapping="Wrap" Margin="28,2,0,0"
                               Text="A wipe erases this device. It cannot be undone."/>

                    <TextBlock Style="{StaticResource FieldLabel}" Text="CHANGE PRODUCT KEY (OPTIONAL)" Margin="2,16,0,6"/>
                    <TextBox x:Name="AdvChangePkBox" Height="38"/>
                    <TextBlock Style="{StaticResource HintText}" Text="Runs changepk.exe to upgrade the Windows edition, then restarts."/>
                  </StackPanel>
                </Border>

                <Border Style="{StaticResource Card}">
                  <StackPanel>
                    <TextBlock Style="{StaticResource CardTitle}" Text="Tools"/>
                    <Button x:Name="AdvDiagnosticsButton" Style="{StaticResource SecondaryButton}" Content="Run Autopilot diagnostics" HorizontalAlignment="Left" Margin="0,0,0,8"/>
                    <TextBlock Style="{StaticResource HintText}" Margin="2,0,0,8" Text="Reads this machine's Autopilot and enrolment state from the local event logs and registry."/>
                    <CheckBox x:Name="AdvDiagnosticsOnlineCheck" Content="Resolve app and policy names from Intune (-Online)"/>
                    <TextBlock Style="{StaticResource HintText}" Margin="28,2,0,12" Text="Without this, apps and policies appear as GUIDs. With it the run signs in to Graph read-only, so it needs the sign-in module, network access and a browser prompt."/>
                    <Button x:Name="AdvWindowsUpdateButton" Style="{StaticResource SecondaryButton}" Content="Install Windows updates" HorizontalAlignment="Left" Margin="0,0,0,8"/>
                    <TextBlock Style="{StaticResource HintText}" Margin="2,0,0,0" Text="Uses the in-box Windows Update agent, so no module download is needed. Runs in its own window because the device may restart."/>
                  </StackPanel>
                </Border>
              </StackPanel>

              <StackPanel Grid.Column="1">
                <Border Style="{StaticResource Card}">
                  <StackPanel>
                    <TextBlock Style="{StaticResource CardTitle}" Text="Engine"/>
                    <TextBlock x:Name="AdvEnginePath" Style="{StaticResource HintText}" Margin="0,0,0,6" Text=""/>
                    <TextBlock x:Name="AdvEngineIntegrity" Style="{StaticResource HintText}" Margin="0,0,0,10" Text=""/>
                    <Button x:Name="AdvVerifyEngineButton" Style="{StaticResource SecondaryButton}" Content="Verify engine integrity" HorizontalAlignment="Stretch"/>

                    <Border Style="{StaticResource Divider}"/>

                    <TextBlock Style="{StaticResource FieldLabel}" Text="SIGN-IN PREREQUISITE"/>
                    <TextBlock x:Name="AdvGraphStatus" Style="{StaticResource HintText}" Margin="0,0,0,10" Text="Checking..."/>
                    <Button x:Name="AdvRepairGraphButton" Style="{StaticResource SecondaryButton}" Content="Repair sign-in module" HorizontalAlignment="Stretch"/>
                    <TextBlock Style="{StaticResource HintText}" Text="Reinstalls Microsoft.Graph.Authentication for all users at the version the engine expects."/>
                  </StackPanel>
                </Border>

                <Border Style="{StaticResource Card}">
                  <StackPanel>
                    <TextBlock Style="{StaticResource CardTitle}" Text="Diagnostics"/>
                    <CheckBox x:Name="AdvShowConsoleCheck" Content="Show the engine console window"/>
                    <TextBlock Style="{StaticResource HintText}" Text="Useful if a run appears to stall. Output still appears in this window either way."/>
                    <Button x:Name="AdvOpenWorkDirButton" Style="{StaticResource SecondaryButton}" Content="Open working folder" HorizontalAlignment="Stretch" Margin="0,12,0,0"/>
                  </StackPanel>
                </Border>

                <Border Style="{StaticResource Card}">
                  <StackPanel>
                    <TextBlock Style="{StaticResource CardTitle}" Text="About"/>

                    <TextBlock Text="Autopilot Import GUI (Community)" FontSize="13" FontWeight="SemiBold"/>
                    <TextBlock x:Name="AboutVersion" Style="{StaticResource HintText}" Margin="0,3,0,0" Text=""/>

                    <TextBlock Text="Built by Mark Orr" Foreground="{StaticResource TextSecondary}" FontSize="12" Margin="0,12,0,0"/>
                    <Button x:Name="AboutGitHubLink" Style="{StaticResource LinkButton}" Content="github.com/markorr321" Margin="0,5,0,0"/>
                    <Button x:Name="AboutSiteLink" Style="{StaticResource LinkButton}" Content="orr365.tools" Margin="0,4,0,0"/>

                    <Border Style="{StaticResource Divider}"/>

                    <TextBlock Style="{StaticResource FieldLabel}" Text="BUILT ON"/>
                    <TextBlock Text="Windows Autopilot Community script by Andrew S Taylor" Style="{StaticResource HintText}" Margin="0,0,0,0"/>
                    <Button x:Name="AboutEngineLink" Style="{StaticResource LinkButton}" Content="github.com/andrew-s-taylor/WindowsAutopilotInfo" Margin="0,5,0,0"/>
                    <TextBlock Text="Original concept: AutoPilot_Import_GUI by Ugur Koc" Style="{StaticResource HintText}" Margin="0,10,0,0"/>
                    <Button x:Name="AboutOriginalLink" Style="{StaticResource LinkButton}" Content="github.com/ugurkocde/AutoPilot_Import_GUI" Margin="0,5,0,0"/>

                    <TextBlock Text="MIT licensed." Style="{StaticResource HintText}" Margin="0,12,0,0"/>
                  </StackPanel>
                </Border>
              </StackPanel>
            </Grid>
          </ScrollViewer>
        </Grid>

        <!-- ================ Logs ================ -->
        <Grid x:Name="PageLogs" Margin="24,20,24,16" Visibility="Collapsed">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <StackPanel Grid.Row="0" Margin="0,0,0,16">
            <TextBlock Style="{StaticResource PageTitle}" Text="Logs"/>
            <TextBlock Style="{StaticResource PageSubtitle}" Text="Everything this session has done, including the full output of each engine run."/>
          </StackPanel>

          <Grid Grid.Row="1" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="LogsRefreshButton" Grid.Column="0" Style="{StaticResource SecondaryButton}" Content="Refresh" Margin="0,0,8,0"/>
            <Button x:Name="LogsCopyButton" Grid.Column="1" Style="{StaticResource SecondaryButton}" Content="Copy all" Margin="0,0,8,0"/>
            <Button x:Name="LogsOpenFolderButton" Grid.Column="2" Style="{StaticResource SecondaryButton}" Content="Open log folder" Margin="0,0,16,0"/>
            <TextBlock x:Name="LogsPathText" Grid.Column="3" VerticalAlignment="Center" Foreground="{StaticResource TextDisabled}" FontSize="11" TextTrimming="CharacterEllipsis" Text=""/>
          </Grid>

          <TextBox x:Name="LogsOutput" Grid.Row="2" Style="{StaticResource OutputBox}"/>
        </Grid>

      </Grid>
    </Grid>

    <!-- ==================== status strip ==================== -->
    <Border Grid.Row="2" Background="#1F1F1F" BorderBrush="#2A2A2A" BorderThickness="0,1,0,0" Padding="20,10">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid Grid.Row="0">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <TextBlock x:Name="StatusText" Grid.Column="0" Foreground="{StaticResource TextSecondary}" FontSize="12"
                     TextTrimming="CharacterEllipsis" VerticalAlignment="Center" Text="Ready."/>
          <TextBlock x:Name="StatusStage" Grid.Column="1" Foreground="{StaticResource TextDisabled}" FontSize="11"
                     VerticalAlignment="Center" Margin="16,0,0,0" Text=""/>
        </Grid>
        <ProgressBar x:Name="StatusProgress" Grid.Row="1" Margin="0,8,0,0" Minimum="0" Maximum="100" Value="0"/>
      </Grid>
    </Border>
  </Grid>
</Window>

'@
$script:ApEmbeddedXaml['Dark'] = @'
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">

  <!--
    Dark theme for Autopilot Import GUI (Community).

    The palette, typography and the PrimaryButton / FieldLabel / Segment styles are carried
    over from VM-Pilot (AutopilotEnroll.GUI.ps1 and VMPilot.GUI.ps1) so the two tools read
    as one family on a technician's bench.

    Everything here is hand-templated because WPF's stock templates are light-themed: a
    plain Background setter leaves ComboBox popups, DataGrid headers and ScrollBars
    stubbornly grey-on-white. No icon fonts are used; glyphs are plain text so the UI
    renders identically in a stripped-down OOBE image.

    Load order matters: StaticResource only resolves backwards, so brushes come first.
  -->

  <!-- ============================ palette ============================ -->
  <Color x:Key="AccentColor">#FF0078D4</Color>

  <SolidColorBrush x:Key="WindowBackground"   Color="#161616"/>
  <SolidColorBrush x:Key="SurfaceBackground"  Color="#1F1F1F"/>
  <SolidColorBrush x:Key="SurfaceRaised"      Color="#252525"/>
  <SolidColorBrush x:Key="SurfaceHover"       Color="#2A2A2A"/>
  <SolidColorBrush x:Key="BorderBrushSubtle"  Color="#2A2A2A"/>
  <SolidColorBrush x:Key="BorderBrushNormal"  Color="#3A3A3A"/>
  <SolidColorBrush x:Key="BorderBrushStrong"  Color="#4A4A4A"/>

  <SolidColorBrush x:Key="AccentBrush"        Color="#0078D4"/>
  <SolidColorBrush x:Key="AccentBrushHover"   Color="#1F8AE0"/>
  <SolidColorBrush x:Key="AccentBrushPressed" Color="#0061B0"/>

  <SolidColorBrush x:Key="TextPrimary"        Color="#FFFFFF"/>
  <SolidColorBrush x:Key="TextSecondary"      Color="#C0C0C0"/>
  <SolidColorBrush x:Key="TextMuted"          Color="#909090"/>
  <SolidColorBrush x:Key="TextDisabled"       Color="#707070"/>

  <SolidColorBrush x:Key="SuccessBrush"       Color="#107C41"/>
  <SolidColorBrush x:Key="SuccessBrushLight"  Color="#3BC77A"/>
  <SolidColorBrush x:Key="WarningBrush"       Color="#C08A20"/>
  <SolidColorBrush x:Key="WarningBrushLight"  Color="#E9B44C"/>
  <SolidColorBrush x:Key="ErrorBrush"         Color="#F03A47"/>
  <SolidColorBrush x:Key="ErrorBrushDim"      Color="#C92B37"/>

  <!-- ScrollViewer is not hand-templated, and its stock template fills the square where a
       horizontal and a vertical bar meet with SystemColors.ControlBrush: #F0F0F0, a white
       block in the corner of every log pane. The fill is a DynamicResource lookup, so
       overriding the key here is enough; matching the output pane background is what makes
       the corner disappear rather than merely darken. -->
  <SolidColorBrush x:Key="{x:Static SystemColors.ControlBrushKey}" Color="#121212"/>

  <FontFamily x:Key="UiFont">Segoe UI Variable, Segoe UI</FontFamily>
  <FontFamily x:Key="MonoFont">Cascadia Mono, Consolas, Courier New</FontFamily>

  <!-- ============================ text ============================ -->
  <Style TargetType="TextBlock">
    <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
    <Setter Property="FontFamily" Value="{StaticResource UiFont}"/>
    <Setter Property="FontSize" Value="13"/>
    <Setter Property="TextOptions.TextFormattingMode" Value="Ideal"/>
  </Style>

  <!-- Uppercase micro-label above an input. Carried over from VM-Pilot. -->
  <Style x:Key="FieldLabel" TargetType="TextBlock">
    <Setter Property="Foreground" Value="{StaticResource TextMuted}"/>
    <Setter Property="FontSize" Value="11"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="Margin" Value="2,0,0,6"/>
  </Style>

  <Style x:Key="PageTitle" TargetType="TextBlock">
    <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
    <Setter Property="FontSize" Value="21"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
  </Style>

  <Style x:Key="PageSubtitle" TargetType="TextBlock">
    <Setter Property="Foreground" Value="{StaticResource TextMuted}"/>
    <Setter Property="FontSize" Value="13"/>
    <Setter Property="TextWrapping" Value="Wrap"/>
    <Setter Property="Margin" Value="0,3,0,0"/>
  </Style>

  <Style x:Key="CardTitle" TargetType="TextBlock">
    <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
    <Setter Property="FontSize" Value="14"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="Margin" Value="0,0,0,12"/>
  </Style>

  <Style x:Key="InfoLabel" TargetType="TextBlock">
    <Setter Property="Foreground" Value="{StaticResource TextMuted}"/>
    <Setter Property="FontSize" Value="12"/>
  </Style>

  <Style x:Key="InfoValue" TargetType="TextBlock">
    <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
    <Setter Property="FontSize" Value="13"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
  </Style>

  <Style x:Key="HintText" TargetType="TextBlock">
    <Setter Property="Foreground" Value="{StaticResource TextDisabled}"/>
    <Setter Property="FontSize" Value="11"/>
    <Setter Property="TextWrapping" Value="Wrap"/>
    <Setter Property="Margin" Value="2,5,0,0"/>
  </Style>

  <!-- ============================ card ============================ -->
  <Style x:Key="Card" TargetType="Border">
    <Setter Property="Background" Value="{StaticResource SurfaceBackground}"/>
    <Setter Property="BorderBrush" Value="{StaticResource BorderBrushSubtle}"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="CornerRadius" Value="8"/>
    <Setter Property="Padding" Value="18"/>
    <Setter Property="Margin" Value="0,0,0,14"/>
  </Style>

  <Style x:Key="Divider" TargetType="Border">
    <Setter Property="Height" Value="1"/>
    <Setter Property="Background" Value="{StaticResource BorderBrushSubtle}"/>
    <Setter Property="Margin" Value="0,14,0,14"/>
  </Style>

  <!-- ============================ buttons ============================ -->
  <Style x:Key="PrimaryButton" TargetType="Button">
    <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
    <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
    <Setter Property="BorderThickness" Value="0"/>
    <Setter Property="FontFamily" Value="{StaticResource UiFont}"/>
    <Setter Property="FontSize" Value="15"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Height" Value="48"/>
    <Setter Property="Padding" Value="20,0"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="Button">
          <Border x:Name="Bd" Background="{TemplateBinding Background}" CornerRadius="6" SnapsToDevicePixels="True">
            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True">
              <Setter TargetName="Bd" Property="Background" Value="{StaticResource AccentBrushHover}"/>
            </Trigger>
            <Trigger Property="IsPressed" Value="True">
              <Setter TargetName="Bd" Property="Background" Value="{StaticResource AccentBrushPressed}"/>
            </Trigger>
            <Trigger Property="IsEnabled" Value="False">
              <Setter TargetName="Bd" Property="Background" Value="{StaticResource SurfaceHover}"/>
              <Setter Property="Foreground" Value="{StaticResource TextDisabled}"/>
              <Setter Property="Cursor" Value="Arrow"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style x:Key="SecondaryButton" TargetType="Button">
    <Setter Property="Background" Value="{StaticResource SurfaceBackground}"/>
    <Setter Property="Foreground" Value="{StaticResource TextSecondary}"/>
    <Setter Property="BorderBrush" Value="{StaticResource BorderBrushNormal}"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="FontFamily" Value="{StaticResource UiFont}"/>
    <Setter Property="FontSize" Value="13"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Height" Value="36"/>
    <Setter Property="Padding" Value="16,0"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="Button">
          <Border x:Name="Bd" Background="{TemplateBinding Background}"
                  BorderBrush="{TemplateBinding BorderBrush}"
                  BorderThickness="{TemplateBinding BorderThickness}"
                  CornerRadius="6" SnapsToDevicePixels="True">
            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True">
              <Setter TargetName="Bd" Property="Background" Value="{StaticResource SurfaceHover}"/>
              <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource BorderBrushStrong}"/>
              <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            </Trigger>
            <Trigger Property="IsEnabled" Value="False">
              <Setter Property="Foreground" Value="{StaticResource TextDisabled}"/>
              <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource BorderBrushSubtle}"/>
              <Setter Property="Cursor" Value="Arrow"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <!-- Destructive actions: wipe, delete-and-re-add. -->
  <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource SecondaryButton}">
    <Setter Property="Foreground" Value="{StaticResource ErrorBrush}"/>
    <Setter Property="BorderBrush" Value="{StaticResource ErrorBrushDim}"/>
  </Style>

  <!-- ============================ nav ============================ -->
  <Style x:Key="NavItem" TargetType="RadioButton">
    <Setter Property="Foreground" Value="{StaticResource TextSecondary}"/>
    <Setter Property="FontFamily" Value="{StaticResource UiFont}"/>
    <Setter Property="FontSize" Value="13"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Focusable" Value="False"/>
    <Setter Property="Height" Value="40"/>
    <Setter Property="Margin" Value="8,2,8,2"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="RadioButton">
          <Grid>
            <Border x:Name="Bd" Background="Transparent" CornerRadius="6"/>
            <!-- Accent rail on the selected item; clearer than a fill alone at a glance. -->
            <Border x:Name="Rail" Width="3" HorizontalAlignment="Left" Margin="0,8,0,8"
                    CornerRadius="2" Background="Transparent"/>
            <ContentPresenter VerticalAlignment="Center" Margin="16,0,10,0"/>
          </Grid>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True">
              <Setter TargetName="Bd" Property="Background" Value="{StaticResource SurfaceBackground}"/>
              <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            </Trigger>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="Bd" Property="Background" Value="{StaticResource SurfaceRaised}"/>
              <Setter TargetName="Rail" Property="Background" Value="{StaticResource AccentBrush}"/>
              <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <!-- Segmented control (v1 / v2 mode switch). Carried over from VM-Pilot. -->
  <Style x:Key="Segment" TargetType="RadioButton">
    <Setter Property="Background" Value="{StaticResource SurfaceBackground}"/>
    <Setter Property="Foreground" Value="#A8A8A8"/>
    <Setter Property="BorderBrush" Value="{StaticResource BorderBrushNormal}"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="FontFamily" Value="{StaticResource UiFont}"/>
    <Setter Property="FontSize" Value="13"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Focusable" Value="False"/>
    <Setter Property="Height" Value="52"/>
    <Setter Property="Padding" Value="14,0"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="RadioButton">
          <Border x:Name="Bd" Background="{TemplateBinding Background}"
                  BorderBrush="{TemplateBinding BorderBrush}"
                  BorderThickness="{TemplateBinding BorderThickness}"
                  CornerRadius="6" SnapsToDevicePixels="True">
            <ContentPresenter HorizontalAlignment="Left" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True">
              <Setter TargetName="Bd" Property="Background" Value="{StaticResource SurfaceHover}"/>
              <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource BorderBrushStrong}"/>
            </Trigger>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="Bd" Property="Background" Value="{StaticResource AccentBrush}"/>
              <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
              <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            </Trigger>
            <Trigger Property="IsEnabled" Value="False">
              <Setter Property="Foreground" Value="{StaticResource TextDisabled}"/>
              <Setter TargetName="Bd" Property="Background" Value="{StaticResource WindowBackground}"/>
              <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource BorderBrushSubtle}"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <!-- ============================ inputs ============================ -->
  <Style TargetType="TextBox">
    <Setter Property="Background" Value="{StaticResource SurfaceBackground}"/>
    <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
    <Setter Property="BorderBrush" Value="{StaticResource BorderBrushNormal}"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="FontFamily" Value="{StaticResource UiFont}"/>
    <Setter Property="FontSize" Value="14"/>
    <Setter Property="Padding" Value="10,8"/>
    <Setter Property="CaretBrush" Value="{StaticResource TextPrimary}"/>
    <Setter Property="SelectionBrush" Value="{StaticResource AccentBrush}"/>
    <Setter Property="VerticalContentAlignment" Value="Center"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="TextBox">
          <Border x:Name="Bd" Background="{TemplateBinding Background}"
                  BorderBrush="{TemplateBinding BorderBrush}"
                  BorderThickness="{TemplateBinding BorderThickness}"
                  CornerRadius="6" SnapsToDevicePixels="True">
            <!-- Padding is deliberately NOT bound to this ScrollViewer's Margin. The text
                 host applies TextBox.Padding itself, exactly as the stock WPF template
                 relies on, so binding it here applied the inset twice: a 38-high box lost
                 16px to the margin and another 16px internally, leaving a 4px line box that
                 an 18.6px glyph run could not render into. The field kept its value and the
                 caret blinked, but nothing was ever painted, and only for the plain
                 TextBoxes: DarkComboBox uses Padding="10,0" and so lost nothing vertically. -->
            <ScrollViewer x:Name="PART_ContentHost"
                          VerticalAlignment="{TemplateBinding VerticalContentAlignment}"
                          Focusable="False"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True">
              <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource BorderBrushStrong}"/>
            </Trigger>
            <Trigger Property="IsKeyboardFocusWithin" Value="True">
              <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
            </Trigger>
            <Trigger Property="IsEnabled" Value="False">
              <Setter Property="Foreground" Value="{StaticResource TextDisabled}"/>
              <Setter TargetName="Bd" Property="Background" Value="{StaticResource WindowBackground}"/>
              <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource BorderBrushSubtle}"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <!-- Read-only monospaced console pane for live engine output and logs. -->
  <Style x:Key="OutputBox" TargetType="TextBox">
    <Setter Property="Background" Value="#121212"/>
    <Setter Property="Foreground" Value="{StaticResource TextSecondary}"/>
    <Setter Property="BorderBrush" Value="{StaticResource BorderBrushSubtle}"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="FontFamily" Value="{StaticResource MonoFont}"/>
    <Setter Property="FontSize" Value="12"/>
    <Setter Property="Padding" Value="10,8"/>
    <Setter Property="IsReadOnly" Value="True"/>
    <Setter Property="IsReadOnlyCaretVisible" Value="False"/>
    <Setter Property="TextWrapping" Value="NoWrap"/>
    <!-- Always show the vertical bar in a log pane. With Auto it is absent until the content
         happens to overflow, which reads as the pane having no scrollbar at all; a console
         that gains and loses its scrollbar as output arrives is also a moving target to grab.
         Horizontal stays Auto: it is only meaningful when a line actually runs off the edge. -->
    <Setter Property="VerticalScrollBarVisibility" Value="Visible"/>
    <Setter Property="HorizontalScrollBarVisibility" Value="Auto"/>
    <Setter Property="SelectionBrush" Value="{StaticResource AccentBrush}"/>
    <Setter Property="VerticalContentAlignment" Value="Top"/>
  </Style>

  <Style x:Key="ToggleArrow" TargetType="ToggleButton">
    <Setter Property="Focusable" Value="False"/>
    <Setter Property="ClickMode" Value="Press"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="ToggleButton">
          <Border Background="Transparent">
            <Path x:Name="Arrow" Data="M 0 0 L 4 4 L 8 0 Z"
                  Fill="{StaticResource TextMuted}"
                  HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True">
              <Setter TargetName="Arrow" Property="Fill" Value="{StaticResource TextPrimary}"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style TargetType="ComboBoxItem">
    <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
    <Setter Property="FontFamily" Value="{StaticResource UiFont}"/>
    <Setter Property="FontSize" Value="13"/>
    <Setter Property="Padding" Value="10,7"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="ComboBoxItem">
          <Border x:Name="Bd" Background="Transparent" Padding="{TemplateBinding Padding}">
            <ContentPresenter/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True">
              <Setter TargetName="Bd" Property="Background" Value="{StaticResource AccentBrush}"/>
            </Trigger>
            <Trigger Property="IsSelected" Value="True">
              <Setter TargetName="Bd" Property="Background" Value="{StaticResource SurfaceHover}"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <!--
    Editable ComboBox (Group Tag). Fully templated: the stock template renders a white
    popup with a Windows-classic button that ignores Background/Foreground setters.
    PART_EditableTextBox and PART_Popup are the names ComboBox looks up by contract.
  -->
  <Style x:Key="DarkComboBox" TargetType="ComboBox">
    <Setter Property="Background" Value="{StaticResource SurfaceBackground}"/>
    <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
    <Setter Property="BorderBrush" Value="{StaticResource BorderBrushNormal}"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="FontFamily" Value="{StaticResource UiFont}"/>
    <Setter Property="FontSize" Value="14"/>
    <Setter Property="Height" Value="38"/>
    <Setter Property="Padding" Value="10,0"/>
    <Setter Property="SnapsToDevicePixels" Value="True"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="ComboBox">
          <Grid>
            <Border x:Name="Bd" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="6" SnapsToDevicePixels="True">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="28"/>
                </Grid.ColumnDefinitions>

                <TextBox x:Name="PART_EditableTextBox" Grid.Column="0"
                         Background="Transparent" BorderThickness="0"
                         Foreground="{TemplateBinding Foreground}"
                         CaretBrush="{StaticResource TextPrimary}"
                         SelectionBrush="{StaticResource AccentBrush}"
                         FontFamily="{TemplateBinding FontFamily}"
                         FontSize="{TemplateBinding FontSize}"
                         Padding="{TemplateBinding Padding}"
                         VerticalContentAlignment="Center"
                         Visibility="Collapsed"/>

                <ContentPresenter x:Name="ContentSite" Grid.Column="0"
                                  Content="{TemplateBinding SelectionBoxItem}"
                                  ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                  Margin="{TemplateBinding Padding}"
                                  VerticalAlignment="Center" HorizontalAlignment="Left"
                                  IsHitTestVisible="False"/>

                <ToggleButton Grid.Column="1" Style="{StaticResource ToggleArrow}"
                              IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"/>
              </Grid>
            </Border>

            <Popup x:Name="PART_Popup" AllowsTransparency="True" Placement="Bottom"
                   IsOpen="{TemplateBinding IsDropDownOpen}" Focusable="False" PopupAnimation="Fade">
              <Border Background="{StaticResource SurfaceRaised}"
                      BorderBrush="{StaticResource BorderBrushNormal}"
                      BorderThickness="1" CornerRadius="6" Margin="0,4,0,0"
                      MinWidth="{TemplateBinding ActualWidth}"
                      MaxHeight="{TemplateBinding MaxDropDownHeight}">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                  <ItemsPresenter Margin="3"/>
                </ScrollViewer>
              </Border>
            </Popup>
          </Grid>

          <ControlTemplate.Triggers>
            <Trigger Property="IsEditable" Value="True">
              <Setter TargetName="PART_EditableTextBox" Property="Visibility" Value="Visible"/>
              <Setter TargetName="ContentSite" Property="Visibility" Value="Collapsed"/>
            </Trigger>
            <Trigger Property="IsMouseOver" Value="True">
              <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource BorderBrushStrong}"/>
            </Trigger>
            <Trigger Property="IsKeyboardFocusWithin" Value="True">
              <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
            </Trigger>
            <Trigger Property="IsEnabled" Value="False">
              <Setter Property="Foreground" Value="{StaticResource TextDisabled}"/>
              <Setter TargetName="Bd" Property="Background" Value="{StaticResource WindowBackground}"/>
              <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource BorderBrushSubtle}"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style TargetType="CheckBox">
    <Setter Property="Foreground" Value="{StaticResource TextSecondary}"/>
    <Setter Property="FontFamily" Value="{StaticResource UiFont}"/>
    <Setter Property="FontSize" Value="13"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Margin" Value="0,5,0,5"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="CheckBox">
          <Grid Background="Transparent">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Border x:Name="Box" Grid.Column="0" Width="18" Height="18" CornerRadius="4"
                    Background="{StaticResource SurfaceBackground}"
                    BorderBrush="{StaticResource BorderBrushNormal}" BorderThickness="1"
                    VerticalAlignment="Center" SnapsToDevicePixels="True">
              <Path x:Name="Tick" Data="M 2,6 L 6,10 L 12,2" Stroke="{StaticResource TextPrimary}"
                    StrokeThickness="2" Visibility="Collapsed"
                    HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ContentPresenter Grid.Column="1" Margin="10,0,0,0" VerticalAlignment="Center"
                              RecognizesAccessKey="True"/>
          </Grid>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True">
              <Setter TargetName="Box" Property="BorderBrush" Value="{StaticResource BorderBrushStrong}"/>
              <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            </Trigger>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="Box" Property="Background" Value="{StaticResource AccentBrush}"/>
              <Setter TargetName="Box" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
              <Setter TargetName="Tick" Property="Visibility" Value="Visible"/>
            </Trigger>
            <Trigger Property="IsEnabled" Value="False">
              <Setter Property="Foreground" Value="{StaticResource TextDisabled}"/>
              <Setter TargetName="Box" Property="Background" Value="{StaticResource WindowBackground}"/>
              <Setter TargetName="Box" Property="BorderBrush" Value="{StaticResource BorderBrushSubtle}"/>
              <Setter Property="Cursor" Value="Arrow"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style x:Key="ChoiceRadio" TargetType="RadioButton">
    <Setter Property="Foreground" Value="{StaticResource TextSecondary}"/>
    <Setter Property="FontFamily" Value="{StaticResource UiFont}"/>
    <Setter Property="FontSize" Value="13"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Margin" Value="0,5,0,5"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="RadioButton">
          <Grid Background="Transparent">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid Grid.Column="0" Width="18" Height="18" VerticalAlignment="Center">
              <Ellipse x:Name="Ring" Stroke="{StaticResource BorderBrushNormal}" StrokeThickness="1"
                       Fill="{StaticResource SurfaceBackground}"/>
              <Ellipse x:Name="Dot" Width="8" Height="8" Fill="{StaticResource TextPrimary}" Visibility="Collapsed"/>
            </Grid>
            <ContentPresenter Grid.Column="1" Margin="10,0,0,0" VerticalAlignment="Center"/>
          </Grid>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True">
              <Setter TargetName="Ring" Property="Stroke" Value="{StaticResource BorderBrushStrong}"/>
              <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            </Trigger>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="Ring" Property="Fill" Value="{StaticResource AccentBrush}"/>
              <Setter TargetName="Ring" Property="Stroke" Value="{StaticResource AccentBrush}"/>
              <Setter TargetName="Dot" Property="Visibility" Value="Visible"/>
              <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            </Trigger>
            <Trigger Property="IsEnabled" Value="False">
              <Setter Property="Foreground" Value="{StaticResource TextDisabled}"/>
              <Setter TargetName="Ring" Property="Stroke" Value="{StaticResource BorderBrushSubtle}"/>
              <Setter Property="Cursor" Value="Arrow"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <!-- ============================ progress ============================ -->
  <Style TargetType="ProgressBar">
    <Setter Property="Height" Value="6"/>
    <Setter Property="Background" Value="{StaticResource SurfaceBackground}"/>
    <Setter Property="Foreground" Value="{StaticResource AccentBrush}"/>
    <Setter Property="BorderThickness" Value="0"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="ProgressBar">
          <Grid>
            <Border Background="{TemplateBinding Background}" CornerRadius="3"/>
            <!-- PART_Track / PART_Indicator are required by ProgressBar's contract;
                 the indeterminate animation is driven off PART_Indicator. -->
            <Border x:Name="PART_Track" Background="Transparent"/>
            <Border x:Name="PART_Indicator" Background="{TemplateBinding Foreground}"
                    CornerRadius="3" HorizontalAlignment="Left"/>
          </Grid>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <!-- ============================ scrollbar ============================ -->
  <Style x:Key="ScrollThumb" TargetType="Thumb">
    <Setter Property="IsTabStop" Value="False"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="Thumb">
          <!-- Even inset on all four sides so the same thumb reads correctly whether the
               bar is vertical or horizontal. #3A3A3A on the #121212 output pane was too
               close to the background to find: a 6px sliver of near-black on black read as
               "there is no scrollbar at all". Light enough to see at a glance, and it
               brightens further on hover so the grab target is obvious. -->
          <Border x:Name="Bd" Background="#6A6A6A" CornerRadius="4" Margin="3"/>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True">
              <Setter TargetName="Bd" Property="Background" Value="#8C8C8C"/>
            </Trigger>
            <Trigger Property="IsDragging" Value="True">
              <Setter TargetName="Bd" Property="Background" Value="{StaticResource AccentBrush}"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <!-- Width/MinWidth here size the *vertical* bar. They must be released for a horizontal
       one, or it is pinned to a 12x12 stub in the corner instead of spanning the viewport:
       the Logs pane could scroll down but never sideways, so long engine lines were
       unreachable. See the Style.Triggers block below. -->
  <Style TargetType="ScrollBar">
    <!-- A visible trough, not Transparent: it tells you the pane scrolls before you find the
         thumb, and gives the thumb something to read against. -->
    <Setter Property="Background" Value="#242424"/>
    <Setter Property="Width" Value="14"/>
    <Setter Property="MinWidth" Value="14"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="ScrollBar">
          <Grid Background="Transparent">
            <Track x:Name="PART_Track" IsDirectionReversed="True">
              <Track.Thumb>
                <Thumb Style="{StaticResource ScrollThumb}"/>
              </Track.Thumb>
              <!-- Empty repeat buttons: no classic arrow boxes, but the page-scroll
                   click targets above and below the thumb still work. -->
              <Track.DecreaseRepeatButton>
                <RepeatButton Command="ScrollBar.PageUpCommand" Opacity="0" Focusable="False"/>
              </Track.DecreaseRepeatButton>
              <Track.IncreaseRepeatButton>
                <RepeatButton Command="ScrollBar.PageDownCommand" Opacity="0" Focusable="False"/>
              </Track.IncreaseRepeatButton>
            </Track>
          </Grid>
          <ControlTemplate.Triggers>
            <Trigger Property="Orientation" Value="Horizontal">
              <Setter TargetName="PART_Track" Property="IsDirectionReversed" Value="False"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
    <Style.Triggers>
      <!-- A horizontal bar must stretch along its axis and be thin across it: exactly the
           opposite of the setters above. Width="Auto" clears the fixed width rather than
           fighting it. -->
      <Trigger Property="Orientation" Value="Horizontal">
        <Setter Property="Width" Value="Auto"/>
        <Setter Property="MinWidth" Value="0"/>
        <Setter Property="Height" Value="14"/>
        <Setter Property="MinHeight" Value="14"/>
      </Trigger>
    </Style.Triggers>
  </Style>

  <!-- ============================ datagrid ============================ -->
  <Style TargetType="DataGridColumnHeader">
    <Setter Property="Background" Value="{StaticResource SurfaceRaised}"/>
    <Setter Property="Foreground" Value="{StaticResource TextMuted}"/>
    <Setter Property="FontFamily" Value="{StaticResource UiFont}"/>
    <Setter Property="FontSize" Value="11"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="Padding" Value="10,8"/>
    <Setter Property="BorderThickness" Value="0,0,0,1"/>
    <Setter Property="BorderBrush" Value="{StaticResource BorderBrushNormal}"/>
    <Setter Property="HorizontalContentAlignment" Value="Left"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="DataGridColumnHeader">
          <Border Background="{TemplateBinding Background}"
                  BorderBrush="{TemplateBinding BorderBrush}"
                  BorderThickness="{TemplateBinding BorderThickness}"
                  Padding="{TemplateBinding Padding}">
            <ContentPresenter VerticalAlignment="Center"/>
          </Border>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style TargetType="DataGridCell">
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="Foreground" Value="{StaticResource TextSecondary}"/>
    <Setter Property="BorderThickness" Value="0"/>
    <Setter Property="Padding" Value="10,7"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="DataGridCell">
          <Border Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
            <ContentPresenter VerticalAlignment="Center"/>
          </Border>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
    <Style.Triggers>
      <Trigger Property="IsSelected" Value="True">
        <Setter Property="Background" Value="{StaticResource SurfaceHover}"/>
        <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
      </Trigger>
    </Style.Triggers>
  </Style>

  <Style TargetType="DataGridRow">
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="BorderBrush" Value="{StaticResource BorderBrushSubtle}"/>
    <Setter Property="BorderThickness" Value="0,0,0,1"/>
    <Style.Triggers>
      <Trigger Property="IsMouseOver" Value="True">
        <Setter Property="Background" Value="{StaticResource SurfaceBackground}"/>
      </Trigger>
    </Style.Triggers>
  </Style>

  <Style TargetType="DataGrid">
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="Foreground" Value="{StaticResource TextSecondary}"/>
    <Setter Property="BorderBrush" Value="{StaticResource BorderBrushSubtle}"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="RowBackground" Value="Transparent"/>
    <Setter Property="AlternatingRowBackground" Value="Transparent"/>
    <Setter Property="GridLinesVisibility" Value="None"/>
    <Setter Property="HeadersVisibility" Value="Column"/>
    <Setter Property="AutoGenerateColumns" Value="False"/>
    <Setter Property="CanUserAddRows" Value="False"/>
    <Setter Property="CanUserDeleteRows" Value="False"/>
    <Setter Property="CanUserResizeRows" Value="False"/>
    <Setter Property="IsReadOnly" Value="True"/>
    <Setter Property="SelectionMode" Value="Single"/>
    <Setter Property="RowHeaderWidth" Value="0"/>
    <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
    <Setter Property="HorizontalScrollBarVisibility" Value="Auto"/>
    <Setter Property="FontFamily" Value="{StaticResource UiFont}"/>
    <Setter Property="FontSize" Value="12"/>
  </Style>

  <!-- ============================ misc ============================ -->
  <!-- Small state chip: connection status, TPM, hash readiness. -->
  <Style x:Key="Pill" TargetType="Border">
    <Setter Property="CornerRadius" Value="10"/>
    <Setter Property="Padding" Value="9,3"/>
    <Setter Property="Background" Value="{StaticResource SurfaceRaised}"/>
    <Setter Property="BorderBrush" Value="{StaticResource BorderBrushNormal}"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="VerticalAlignment" Value="Center"/>
  </Style>

  <Style x:Key="PillText" TargetType="TextBlock">
    <Setter Property="FontSize" Value="11"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="Foreground" Value="{StaticResource TextSecondary}"/>
  </Style>

  <Style x:Key="LinkButton" TargetType="Button">
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="Foreground" Value="{StaticResource AccentBrushHover}"/>
    <Setter Property="BorderThickness" Value="0"/>
    <Setter Property="FontFamily" Value="{StaticResource UiFont}"/>
    <Setter Property="FontSize" Value="12"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Padding" Value="0"/>
    <Setter Property="HorizontalAlignment" Value="Left"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="Button">
          <Border Background="Transparent">
            <TextBlock x:Name="Txt" Text="{TemplateBinding Content}"
                       Foreground="{TemplateBinding Foreground}"
                       FontSize="{TemplateBinding FontSize}"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True">
              <Setter TargetName="Txt" Property="TextDecorations" Value="Underline"/>
            </Trigger>
            <Trigger Property="IsEnabled" Value="False">
              <Setter TargetName="Txt" Property="Foreground" Value="{StaticResource TextDisabled}"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style TargetType="ScrollViewer">
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="BorderThickness" Value="0"/>
  </Style>

</ResourceDictionary>

'@

$script:ApEmbeddedScripts = @{}
$script:ApEmbeddedScripts['get-windowsautopilotinfocommunity.ps1'] = @'
PCNQU1NjcmlwdEluZm8NCi5WRVJTSU9OIDUuMC4xNg0KLkdVSUQgMzllZmM5YzUtN2I1MS00ZDFmLWI2NTAtMGYzODE4ZTUzMjdhDQouQVVUSE9SIEFuZHJl
d1RheWxvciBmb3JrZWQgZnJvbSB0aGUgb3JpZ2luYWwgYnkgdGhlIGxlZ2VuZCB3aG8gaXMgTWljaGFlbCBOaWVoYXVzDQouQ09NUEFOWU5BTUUgDQouQ09Q
WVJJR0hUIEdQTA0KLlRBR1MgaW50dW5lIGVuZHBvaW50IE1FTSBhdXRvcGlsb3QNCi5MSUNFTlNFVVJJIGh0dHBzOi8vZ2l0aHViLmNvbS9hbmRyZXctcy10
YXlsb3IvV2luZG93c0F1dG9waWxvdEluZm8vYmxvYi9tYWluL0xJQ0VOU0UNCi5QUk9KRUNUVVJJIGh0dHBzOi8vZ2l0aHViLmNvbS9hbmRyZXctcy10YXls
b3IvV2luZG93c0F1dG9waWxvdEluZm8vY29tbXVuaXR5DQouSUNPTlVSSSANCi5FWFRFUk5BTE1PRFVMRURFUEVOREVOQ0lFUw0KLlJFUVVJUkVEU0NSSVBU
UyANCi5FWFRFUk5BTFNDUklQVERFUEVOREVOQ0lFUyANCi5SRUxFQVNFTk9URVMNCnYxLjAuMSAtIEFkZGVkIHN1cHBvcnQgdG8gdXBkYXRlIGdyb3VwIHRh
ZyBvbiBleGlzdGluZyBkZXZpY2VzDQp2MS4wLjIgLSBVcGRhdGVkIGxvZ2ljIHVzZWQgdG8gdXBkYXRlIGdyb3VwIHRhZyBvbiBleGlzdGluZyBkZXZpY2Vz
IFtsaW5lcyAxOTgyLTE5OTAsIDIwNTgtMjA2MF0NCnYxLjAuMyAtIEJ1ZyBGaXgNCnYxLjAuNCAtIFN1cHByZXNzZWQgZXJyb3Igd2hlbiBpbXBvcnRpbmcg
bW9kdWxlcyBpZiBpbiB1c2UNCnYyLjAuMCAtIEFkZGVkIEludHVuZSBXaXBlIGFuZCBTeXNwcmVwIFBhcmFtZXRlcnMNCnYzLjAuMCAtIFN1cHBvcnQgYWRk
ZWQgZm9yIHYyIEdyYXBoIFNESw0KdjMuMC4zIC0gQXV0aGVudGljYXRpb24gZml4ZXMNCnYzLjAuNCAtIFdpcGUgZml4DQp2My4wLjUgLSBBZGRlZCBzdXBw
b3J0IGZvciBwcmUtcHJvdmlzaW9uaW5nDQp2NC4wLjAgLSBBZGRlZCBzdXBwb3J0IHRvIGRlbGV0ZSBleGlzdGluZyBkZXZpY2VzDQp2NC4wLjEgLSBJbXBv
cnQgTW9kdWxlIGZpeA0KdjQuMC4yIC0gQ29kZSBTaWduZWQhIQ0KdjQuMC4zIC0gVGltZXN0YW1wIGZpeA0KdjQuMC40IC0gVXBkYXRlZCBkZXZpY2VzIGdy
YWINCnY0LjAuNSAtIEFkZGVkIG5ld2RldmljZSBwYXJhbWV0ZXIgZm9yIHF1aWNrZXIgaW1wb3J0cw0KdjQuMC42IC0gUmVnaW9uIGZpeA0KdjQuMC43IC0g
QWRkZWQgQ2hhbmdlUEsgc3dpdGNoDQp2NC4wLjggLSBBZGRlZCBsb2dpYyBhcm91bmQgdGhlIHN5bmMgY29tbWFuZCAmIEFkZGVkIEF1dG9JdCBzY3JpcHQg
Zm9yIHByZS1wcm92DQp2NC4wLjkgLSBFeHRlbmRlZCBzeW5jIHRpbWVvdXQNCnY0LjAuMTAgLSBBZGRlZCBhcnJheSBmb3IgZ3JvdXANCnY0LjAuMTEgLSBB
ZGRlZCBjZXJ0IGJhc2VkIE1zR3JhcGggY29ubmVjdGlvbg0KdjUuMC4wIC0gQWRkZWQgc3VwcG9ydCBmb3IgZGV2aWNlIGlkZW50aWZpZXJzDQp2NS4wLjIg
LSBSZW1vdmVkIGRvdHMgYW5kIGNvbW1hcyBmcm9tIG1ha2UgYW5kIG1vZGVsDQp2NS4wLjMgLSBSZW1vdmVkIEdyb3VwIGFuZCBHcm91cE1lbWJlciBzY29w
ZXMgaWYgYWRkIHRvIGdyb3VwIG5vdCBzZWxlY3RlZA0KdjUuMC40IC0gU3dpdGNoZWQgZnJvbSBHcmFwaCBHcm91cHMgbW9kdWxlIHRvIHJhdyByZXF1ZXN0
cw0KdjUuMC41IC0gR3JvdXBzIGZpeA0KdjUuMC42IC0gRG9uJ3QgZGlzcGxheSBFbnRyYSBJRCB0b2tlbjsgZG9uJ3QgcHJvbXB0IGlmIC1Gb3JjZSBpcyBz
cGVjaWZpZWQNCnY1LjAuNyAtIFJlLXdyaXR0ZW4gdXBkYXRlL2RlbGV0ZSB0byBzdG9wIGdyYWJiaW5nIGFsbCBkZXZpY2VzDQp2NS4wLjggLSBBZGRlZCBj
ZXJ0aWZpY2F0ZSBhdXRoZW50aWNhdGlvbg0KdjUuMC45IC0gQW5vdGhlciBncm91cHMgZml4LCB1c2UgbmFtZSBpbnN0ZWFkIG9mIElEIGZvciBsb29rdXAN
CnY1LjAuMTAgLSBSZW1vdmVkIGFuIHVudXNlZCB2YXJpYWJsZSB0byB0cmlnZ2VyIGVycm9yIGZvciBkZXZpY2VzIGluIGFub3RoZXIgdGVuYW50DQp2NS4w
LjExIC0gQWRkZWQgbmV3IHBlcm1pc3Npb25zDQp2NS4wLjEyIC0gV0FNIGZpeA0KdjUuMC4xMyAtIFdBTSBVcGRhdGUNCnY1LjAuMTQgLSBVcGRhdGUgZGV2
aWNlIGZpeCBmcm9tIE1hbmVsIFJvZGVybyBvbiBic2t5DQp2NS4wLjE1IC0gWWV0IGFub3RoZXIgV0FNIGZpeA0KdjUuMC4xNiAtIEZpeGVkIHN5bmMgYnVn
IGFuZCBhZGRlZCBHcmFwaCBkaXNjb25uZWN0DQojPg0KDQo8Iw0KLlNZTk9QU0lTDQpSZXRyaWV2ZXMgdGhlIFdpbmRvd3MgQXV0b1BpbG90IGRlcGxveW1l
bnQgZGV0YWlscyBmcm9tIG9uZSBvciBtb3JlIGNvbXB1dGVycyAtIENvbW11bml0eSBWZXJzaW9uDQpHUEwgTElDRU5TRQ0KUGVybWlzc2lvbiBpcyBoZXJl
YnkgZ3JhbnRlZCwgZnJlZSBvZiBjaGFyZ2UsIHRvIGFueSBwZXJzb24gb2J0YWluaW5nIGEgY29weSBvZiB0aGlzIHNvZnR3YXJlIGFuZCBhc3NvY2lhdGVk
IGRvY3VtZW50YXRpb24gZmlsZXMgKHRoZSAiU29mdHdhcmUiKSwgdG8gZGVhbCBpbiB0aGUgU29mdHdhcmUgd2l0aG91dCByZXN0cmljdGlvbiwgaW5jbHVk
aW5nIHdpdGhvdXQgbGltaXRhdGlvbiB0aGUgcmlnaHRzIHRvIHVzZSwgY29weSwgbW9kaWZ5LCBtZXJnZSwgcHVibGlzaCwgZGlzdHJpYnV0ZSwgc3VibGlj
ZW5zZSwgYW5kL29yIHNlbGwgY29waWVzIG9mIHRoZSBTb2Z0d2FyZSwgYW5kIHRvIHBlcm1pdCBwZXJzb25zIHRvIHdob20gdGhlIFNvZnR3YXJlIGlzIGZ1
cm5pc2hlZCB0byBkbyBzbywgc3ViamVjdCB0byB0aGUgZm9sbG93aW5nIGNvbmRpdGlvbnM6DQpUaGUgYWJvdmUgY29weXJpZ2h0IG5vdGljZSBhbmQgdGhp
cyBwZXJtaXNzaW9uIG5vdGljZSBzaGFsbCBiZSBpbmNsdWRlZCBpbiBhbGwgY29waWVzIG9yIHN1YnN0YW50aWFsIHBvcnRpb25zIG9mIHRoZSBTb2Z0d2Fy
ZS4NClRIRSBTT0ZUV0FSRSBJUyBQUk9WSURFRCAiQVMgSVMiLCBXSVRIT1VUIFdBUlJBTlRZIE9GIEFOWSBLSU5ELCBFWFBSRVNTIE9SIElNUExJRUQsIElO
Q0xVRElORyBCVVQgTk9UIExJTUlURUQgVE8gVEhFIFdBUlJBTlRJRVMgT0YgTUVSQ0hBTlRBQklMSVRZLCBGSVRORVNTIEZPUiBBIFBBUlRJQ1VMQVIgUFVS
UE9TRSBBTkQgTk9OSU5GUklOR0VNRU5ULiBJTiBOTyBFVkVOVCBTSEFMTCBUSEUgQVVUSE9SUyBPUiBDT1BZUklHSFQgSE9MREVSUyBCRSBMSUFCTEUgRk9S
IEFOWSBDTEFJTSwgREFNQUdFUyBPUiBPVEhFUiBMSUFCSUxJVFksIFdIRVRIRVIgSU4gQU4gQUNUSU9OIE9GIENPTlRSQUNULCBUT1JUIE9SIE9USEVSV0lT
RSwgQVJJU0lORyBGUk9NLCBPVVQgT0YgT1IgSU4gQ09OTkVDVElPTiBXSVRIIFRIRSBTT0ZUV0FSRSBPUiBUSEUgVVNFIE9SIE9USEVSIERFQUxJTkdTIElO
IFRIRSBTT0ZUV0FSRS4NCi5ERVNDUklQVElPTg0KVGhpcyBzY3JpcHQgdXNlcyBXTUkgdG8gcmV0cmlldmUgcHJvcGVydGllcyBuZWVkZWQgZm9yIGEgY3Vz
dG9tZXIgdG8gcmVnaXN0ZXIgYSBkZXZpY2Ugd2l0aCBXaW5kb3dzIEF1dG9waWxvdC4gTm90ZSB0aGF0IGl0IGlzIG5vcm1hbCBmb3IgdGhlIHJlc3VsdGlu
ZyBDU1YgZmlsZSB0byBub3QgY29sbGVjdCBhIFdpbmRvd3MgUHJvZHVjdCBJRCAoUEtJRCkgdmFsdWUgc2luY2UgdGhpcyBpcyBub3QgcmVxdWlyZWQgdG8g
cmVnaXN0ZXIgYSBkZXZpY2UuIE9ubHkgdGhlIHNlcmlhbCBudW1iZXIgYW5kIGhhcmR3YXJlIGhhc2ggd2lsbCBiZSBwb3B1bGF0ZWQuDQouUEFSQU1FVEVS
IE5hbWUNClRoZSBuYW1lcyBvZiB0aGUgY29tcHV0ZXJzLiBUaGVzZSBjYW4gYmUgcHJvdmlkZWQgdmlhIHRoZSBwaXBlbGluZSAocHJvcGVydHkgbmFtZSBO
YW1lIG9yIG9uZSBvZiB0aGUgYXZhaWxhYmxlIGFsaWFzZXMsIEROU0hvc3ROYW1lLCBDb21wdXRlck5hbWUsIGFuZCBDb21wdXRlcikuDQouUEFSQU1FVEVS
IE91dHB1dEZpbGUNClRoZSBuYW1lIG9mIHRoZSBDU1YgZmlsZSB0byBiZSBjcmVhdGVkIHdpdGggdGhlIGRldGFpbHMgZm9yIHRoZSBjb21wdXRlcnMuIElm
IG5vdCBzcGVjaWZpZWQsIHRoZSBkZXRhaWxzIHdpbGwgYmUgcmV0dXJuZWQgdG8gdGhlIFBvd2VyU2hlbGwNCnBpcGVsaW5lLg0KLlBBUkFNRVRFUiBBcHBl
bmQNClN3aXRjaCB0byBzcGVjaWZ5IHRoYXQgbmV3IGNvbXB1dGVyIGRldGFpbHMgc2hvdWxkIGJlIGFwcGVuZGVkIHRvIHRoZSBzcGVjaWZpZWQgb3V0cHV0
IGZpbGUsIGluc3RlYWQgb2Ygb3ZlcndyaXRpbmcgdGhlIGV4aXN0aW5nIGZpbGUuDQouUEFSQU1FVEVSIENyZWRlbnRpYWwNCkNyZWRlbnRpYWxzIHRoYXQg
c2hvdWxkIGJlIHVzZWQgd2hlbiBjb25uZWN0aW5nIHRvIGEgcmVtb3RlIGNvbXB1dGVyIChub3Qgc3VwcG9ydGVkIHdoZW4gZ2F0aGVyaW5nIGRldGFpbHMg
ZnJvbSB0aGUgbG9jYWwgY29tcHV0ZXIpLg0KLlBBUkFNRVRFUiBQYXJ0bmVyDQpTd2l0Y2ggdG8gc3BlY2lmeSB0aGF0IHRoZSBjcmVhdGVkIENTViBmaWxl
IHNob3VsZCB1c2UgdGhlIHNjaGVtYSBmb3IgUGFydG5lciBDZW50ZXIgKHVzaW5nIHNlcmlhbCBudW1iZXIsIG1ha2UsIGFuZCBtb2RlbCkuDQouUEFSQU1F
VEVSIEdyb3VwVGFnDQpBbiBvcHRpb25hbCB0YWcgdmFsdWUgdGhhdCBzaG91bGQgYmUgaW5jbHVkZWQgaW4gYSBDU1YgZmlsZSB0aGF0IGlzIGludGVuZGVk
IHRvIGJlIHVwbG9hZGVkIHZpYSBJbnR1bmUgKG5vdCBzdXBwb3J0ZWQgYnkgUGFydG5lciBDZW50ZXIgb3IgTWljcm9zb2Z0IFN0b3JlIGZvciBCdXNpbmVz
cykuDQouUEFSQU1FVEVSIEFzc2lnbmVkVXNlcg0KQW4gb3B0aW9uYWwgdmFsdWUgc3BlY2lmeWluZyB0aGUgVVBOIG9mIHRoZSB1c2VyIHRvIGJlIGFzc2ln
bmVkIHRvIHRoZSBkZXZpY2UuIFRoaXMgY2FuIG9ubHkgYmUgc3BlY2lmaWVkIGZvciBJbnR1bmUgKG5vdCBzdXBwb3J0ZWQgYnkgUGFydG5lciBDZW50ZXIg
b3IgTWljcm9zb2Z0IFN0b3JlIGZvciBCdXNpbmVzcykuDQouUEFSQU1FVEVSIE9ubGluZQ0KQWRkIGNvbXB1dGVycyB0byBXaW5kb3dzIEF1dG9waWxvdCB2
aWEgdGhlIEludHVuZSBHcmFwaCBBUEkNCi5QQVJBTUVURVIgQXNzaWduZWRDb21wdXRlck5hbWUNCkFuIG9wdGlvbmFsIHZhbHVlIHNwZWNpZnlpbmcgdGhl
IGNvbXB1dGVyIG5hbWUgdG8gYmUgYXNzaWduZWQgdG8gdGhlIGRldmljZS4gVGhpcyBjYW4gb25seSBiZSBzcGVjaWZpZWQgd2l0aCB0aGUgLU9ubGluZSBz
d2l0Y2ggYW5kIG9ubHkgd29ya3Mgd2l0aCBBQUQgam9pbiBzY2VuYXJpb3MuDQouUEFSQU1FVEVSIEFkZFRvR3JvdXANClNwZWNpZmllcyB0aGUgbmFtZSBv
ZiB0aGUgRW50cmEgZ3JvdXAgdGhhdCB0aGUgbmV3IGRldmljZSBzaG91bGQgYmUgYWRkZWQgdG8uDQouUEFSQU1FVEVSIEFzc2lnbg0KV2FpdCBmb3IgdGhl
IEF1dG9waWxvdCBwcm9maWxlIGFzc2lnbm1lbnQuIChUaGlzIGNhbiB0YWtlIGEgd2hpbGUgZm9yIGR5bmFtaWMgZ3JvdXBzLikNCi5QQVJBTUVURVIgUmVi
b290DQpSZWJvb3QgdGhlIGRldmljZSBhZnRlciB0aGUgQXV0b3BpbG90IHByb2ZpbGUgaGFzIGJlZW4gYXNzaWduZWQgKG5lY2Vzc2FyeSB0byBkb3dubG9h
ZCB0aGUgcHJvZmlsZSBhbmQgYXBwbHkgdGhlIGNvbXB1dGVyIG5hbWUsIGlmIHNwZWNpZmllZCkuDQouUEFSQU1FVEVSIFdpcGUNCldpcGUgdGhlIGRldmlj
ZSBhZnRlciB0aGUgQXV0b3BpbG90IHByb2ZpbGUgaGFzIGJlZW4gYXNzaWduZWQgKHNlbmRzIGFuIEludHVuZSB3aXBlIGZvciBJbnR1bmUgbWFuYWdlZCBk
ZXZpY2VzIG9ubHkpLg0KLlBBUkFNRVRFUiBTeXNwcmVwDQpLaWNrcyBvZmYgU3lzcHJlcCBhZnRlciB0aGUgQXV0cGlsb3QgcHJvZmlsZSBoYXMgYmVlbiBh
c3NpZ25lZA0KLlBBUkFNRVRFUiBEZWxldGUNClJlbW92ZXMgdGhlIGRldmljZSBpZiBpdCBhbHJlYWR5IGV4aXN0cw0KLlBBUkFNRVRFUiBVcGRhdGV0YWcN
ClVwZGF0ZXMgZ3JvdXAgdGFnIG9uIGV4aXN0aW5nIGRldmljZXMNCi5QQVJBTUVURVIgcHJlcHJvdg0KUHJlc3NlcyBXaW5kb3dzIGtleSA1IHRpbWVzIGZv
ciB3aGl0ZWdsb3ZlIHByZS1wcm92aXNpb25pbmcNCi5QQVJBTUVURVIgaWRlbnRpZmllcg0KQ3JlYXRlcyBkZXZpY2UgaWRlbnRpZmllci4gIENhbiBiZSB1
c2VkIHdpdGggT3V0cHV0RmlsZSBvciBPbmxpbmUuICBDYW4gYWxzbyByZWNlaXZlIGlucHV0cyBmcm9tIElucHV0RmlsZQ0KLlBBUkFNRVRFUiBJbnB1dEZp
bGUNCkNTViBmaWxlIGNvbnRhaW5pbmcgbXVsdGlwbGUgZGV2aWNlIGlkZW50aWZpZXJzDQouUEFSQU1FVEVSIENoYW5nZVBLDQpTcGVjaWZpZXMgYSBwcm9k
dWN0IGtleSB0byBpbmplY3QgaW50byB0aGUgT1MuICBUaGlzIHdpbGwgY2F1c2UgdGhlIGNvbXB1dGVyIHRvIHJlYm9vdC4gIFRoaXMgc2hvdWxkIGJlIGNv
bWJpbmVkIHdpdGgNCnRoZSAtT25saW5lIGFuZCAtQXNzaWduIHN3aXRjaGVzLg0KLkVYQU1QTEUNCi5cR2V0LVdpbmRvd3NBdXRvUGlsb3RJbmZvLnBzMSAt
Q29tcHV0ZXJOYW1lIE1ZQ09NUFVURVIgLU91dHB1dEZpbGUgLlxNeUNvbXB1dGVyLmNzdg0KLkVYQU1QTEUNCi5cR2V0LVdpbmRvd3NBdXRvUGlsb3RJbmZv
LnBzMSAtQ29tcHV0ZXJOYW1lIE1ZQ09NUFVURVIgLU91dHB1dEZpbGUgLlxNeUNvbXB1dGVyLmNzdiAtR3JvdXBUYWcgS2lvc2sNCi5FWEFNUExFDQouXEdl
dC1XaW5kb3dzQXV0b1BpbG90SW5mby5wczEgLUNvbXB1dGVyTmFtZSBNWUNPTVBVVEVSIC1PdXRwdXRGaWxlIC5cTXlDb21wdXRlci5jc3YgLUdyb3VwVGFn
IEtpb3NrIC1Bc3NpZ25lZFVzZXIgSm9obkRvZUBjb250b3NvLmNvbQ0KLkVYQU1QTEUNCi5cR2V0LVdpbmRvd3NBdXRvUGlsb3RJbmZvLnBzMSAtQ29tcHV0
ZXJOYW1lIE1ZQ09NUFVURVIgLU91dHB1dEZpbGUgLlxNeUNvbXB1dGVyLmNzdiAtQXBwZW5kDQouRVhBTVBMRQ0KLlxHZXQtV2luZG93c0F1dG9QaWxvdElu
Zm8ucHMxIC1Db21wdXRlck5hbWUgTVlDT01QVVRFUjEsTVlDT01QVVRFUjIgLU91dHB1dEZpbGUgLlxNeUNvbXB1dGVycy5jc3YNCi5FWEFNUExFDQpHZXQt
QURDb21wdXRlciAtRmlsdGVyICogfCAuXEdldFdpbmRvd3NBdXRvUGlsb3RJbmZvLnBzMSAtT3V0cHV0RmlsZSAuXE15Q29tcHV0ZXJzLmNzdg0KLkVYQU1Q
TEUNCkdldC1DTUNvbGxlY3Rpb25NZW1iZXIgLUNvbGxlY3Rpb25OYW1lICJBbGwgU3lzdGVtcyIgfCAuXEdldFdpbmRvd3NBdXRvUGlsb3RJbmZvLnBzMSAt
T3V0cHV0RmlsZSAuXE15Q29tcHV0ZXJzLmNzdg0KLkVYQU1QTEUNCi5cR2V0LVdpbmRvd3NBdXRvUGlsb3RJbmZvLnBzMSAtQ29tcHV0ZXJOYW1lIE1ZQ09N
UFVURVIxLE1ZQ09NUFVURVIyIC1PdXRwdXRGaWxlIC5cTXlDb21wdXRlcnMuY3N2IC1QYXJ0bmVyDQouRVhBTVBMRQ0KLlxHZXRXaW5kb3dzQXV0b1BpbG90
SW5mby5wczEgLU9ubGluZQ0KLk5PVEVTDQpWZXJzaW9uOiAgICAgICAgNS4wLjE2DQpBdXRob3I6ICAgICAgICAgQW5kcmV3IFRheWxvcg0KV1dXOiAgICAg
ICAgICAgIGFuZHJld3N0YXlsb3IuY29tDQpDcmVhdGlvbiBEYXRlOiAgMTQvMDYvMjAyMw0KIz4NCg0KW0NtZGxldEJpbmRpbmcoRGVmYXVsdFBhcmFtZXRl
clNldE5hbWUgPSAnRGVmYXVsdCcpXQ0KcGFyYW0oDQogICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkRmFsc2UsIFZhbHVlRnJvbVBpcGVsaW5lID0gJFRy
dWUsIFZhbHVlRnJvbVBpcGVsaW5lQnlQcm9wZXJ0eU5hbWUgPSAkVHJ1ZSwgUG9zaXRpb24gPSAwKV1bYWxpYXMoIkROU0hvc3ROYW1lIiwgIkNvbXB1dGVy
TmFtZSIsICJDb21wdXRlciIpXSBbU3RyaW5nW11dICROYW1lID0gQCgibG9jYWxob3N0IiksDQogICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkRmFsc2Up
XSBbU3RyaW5nXSAkT3V0cHV0RmlsZSA9ICIiLCANCiAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICRGYWxzZSldIFtTdHJpbmddICRJbnB1dEZpbGUgPSAi
IiwgDQogICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkRmFsc2UpXSBbU3dpdGNoXSAkaWRlbnRpZmllciA9ICRmYWxzZSwgDQogICAgW1BhcmFtZXRlcihN
YW5kYXRvcnkgPSAkRmFsc2UpXSBbU3RyaW5nXSAkR3JvdXBUYWcgPSAiIiwNCiAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICRGYWxzZSldIFtTdHJpbmdd
ICRBc3NpZ25lZFVzZXIgPSAiIiwNCiAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICRGYWxzZSldIFtTd2l0Y2hdICRBcHBlbmQgPSAkZmFsc2UsDQogICAg
W1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkRmFsc2UpXSBbU3lzdGVtLk1hbmFnZW1lbnQuQXV0b21hdGlvbi5QU0NyZWRlbnRpYWxdICRDcmVkZW50aWFsID0g
JG51bGwsDQogICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkRmFsc2UpXSBbU3dpdGNoXSAkUGFydG5lciA9ICRmYWxzZSwNCiAgICBbUGFyYW1ldGVyKE1h
bmRhdG9yeSA9ICRGYWxzZSldIFtTd2l0Y2hdICRGb3JjZSA9ICRmYWxzZSwNCiAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICRUcnVlLCBQYXJhbWV0ZXJT
ZXROYW1lID0gJ09ubGluZScpXSBbU3dpdGNoXSAkT25saW5lID0gJGZhbHNlLA0KICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJEZhbHNlLCBQYXJhbWV0
ZXJTZXROYW1lID0gJ09ubGluZScpXSBbU3RyaW5nXSAkVGVuYW50SWQgPSAiIiwNCiAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICRGYWxzZSwgUGFyYW1l
dGVyU2V0TmFtZSA9ICdPbmxpbmUnKV0gW1N0cmluZ10gJEFwcElkID0gIiIsDQogICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkRmFsc2UsIFBhcmFtZXRl
clNldE5hbWUgPSAnT25saW5lJyldIFtTdHJpbmddICRBcHBTZWNyZXQgPSAiIiwNCiAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICRGYWxzZSwgUGFyYW1l
dGVyU2V0TmFtZSA9ICdPbmxpbmUnKV0gW1N0cmluZ10gJENlcnRpZmljYXRlU3ViamVjdE5hbWUgPSAiIiwNCiAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9
ICRGYWxzZSwgUGFyYW1ldGVyU2V0TmFtZSA9ICdPbmxpbmUnKV0gW1N0cmluZ10gJENlcnRpZmljYXRlVGh1bWJwcmludCA9ICIiLA0KICAgIFtQYXJhbWV0
ZXIoTWFuZGF0b3J5ID0gJEZhbHNlLCBQYXJhbWV0ZXJTZXROYW1lID0gJ09ubGluZScpXSBbU3RyaW5nW11dICRBZGRUb0dyb3VwID0gIiIsDQogICAgW1Bh
cmFtZXRlcihNYW5kYXRvcnkgPSAkRmFsc2UsIFBhcmFtZXRlclNldE5hbWUgPSAnT25saW5lJyldIFtTdHJpbmddICRBc3NpZ25lZENvbXB1dGVyTmFtZSA9
ICIiLA0KICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJEZhbHNlLCBQYXJhbWV0ZXJTZXROYW1lID0gJ09ubGluZScpXSBbU3dpdGNoXSAkQXNzaWduID0g
JGZhbHNlLCANCiAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICRGYWxzZSwgUGFyYW1ldGVyU2V0TmFtZSA9ICdPbmxpbmUnKV0gW1N3aXRjaF0gJFJlYm9v
dCA9ICRmYWxzZSwNCiAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICRGYWxzZSwgUGFyYW1ldGVyU2V0TmFtZSA9ICdPbmxpbmUnKV0gW1N3aXRjaF0gJFdp
cGUgPSAkZmFsc2UsDQogICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkRmFsc2UsIFBhcmFtZXRlclNldE5hbWUgPSAnT25saW5lJyldIFtTd2l0Y2hdICRT
eXNwcmVwID0gJGZhbHNlLA0KICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJEZhbHNlLCBQYXJhbWV0ZXJTZXROYW1lID0gJ09ubGluZScpXSBbU3dpdGNo
XSAkcHJlcHJvdiA9ICRmYWxzZSwNCiAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICRGYWxzZSwgUGFyYW1ldGVyU2V0TmFtZSA9ICdPbmxpbmUnKV0gW1N3
aXRjaF0gJGRlbGV0ZSA9ICRmYWxzZSwNCiAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICRGYWxzZSwgUGFyYW1ldGVyU2V0TmFtZSA9ICdPbmxpbmUnKV0g
W1N3aXRjaF0gJHVwZGF0ZXRhZyA9ICRmYWxzZSwNCiAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICRGYWxzZSwgUGFyYW1ldGVyU2V0TmFtZSA9ICdPbmxp
bmUnKV0gW1N3aXRjaF0gJG5ld2RldmljZSA9ICRmYWxzZSwNCiAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICRGYWxzZSwgUGFyYW1ldGVyU2V0TmFtZSA9
ICdPbmxpbmUnKV0gW1N0cmluZ10gJENoYW5nZVBLID0gIiINCikNCg0KQmVnaW4gew0KICAgICMgSW5pdGlhbGl6ZSBlbXB0eSBsaXN0DQogICAgJGNvbXB1
dGVycyA9IEAoKQ0KDQogICAgIyBJZiBvbmxpbmUsIG1ha2Ugc3VyZSB3ZSBhcmUgYWJsZSB0byBhdXRoZW50aWNhdGUNCiAgICBpZiAoJE9ubGluZSkgew0K
DQogICAgICAgICMgR2V0IE51R2V0DQogICAgICAgICRwcm92aWRlciA9IEdldC1QYWNrYWdlUHJvdmlkZXIgTnVHZXQgLUVycm9yQWN0aW9uIElnbm9yZQ0K
ICAgICAgICBpZiAoLW5vdCAkcHJvdmlkZXIpIHsNCiAgICAgICAgICAgIFdyaXRlLUhvc3QgIkluc3RhbGxpbmcgcHJvdmlkZXIgTnVHZXQiDQogICAgICAg
ICAgICBGaW5kLVBhY2thZ2VQcm92aWRlciAtTmFtZSBOdUdldCAtRm9yY2VCb290c3RyYXAgLUluY2x1ZGVEZXBlbmRlbmNpZXMNCiAgICAgICAgfQ0KICAg
ICAgICANCiAgICAgICAgIyBHZXQgR3JhcGggQXV0aGVudGljYXRpb24gbW9kdWxlIChhbmQgZGVwZW5kZW5jaWVzKQ0KICAgICAgICAkbW9kdWxlID0gSW1w
b3J0LU1vZHVsZSBtaWNyb3NvZnQuZ3JhcGguYXV0aGVudGljYXRpb24gLVBhc3NUaHJ1IC1FcnJvckFjdGlvbiBJZ25vcmUNCiAgICAgICAgaWYgKC1ub3Qg
JG1vZHVsZSkgew0KICAgICAgICAgICAgV3JpdGUtSG9zdCAiSW5zdGFsbGluZyBtb2R1bGUgbWljcm9zb2Z0LmdyYXBoLmF1dGhlbnRpY2F0aW9uIg0KICAg
ICAgICAgICAgSW5zdGFsbC1Nb2R1bGUgbWljcm9zb2Z0LmdyYXBoLmF1dGhlbnRpY2F0aW9uIC1Gb3JjZSAtRXJyb3JBY3Rpb24gSWdub3JlIC1NYXhpbXVt
VmVyc2lvbiAyLjkuMQ0KICAgICAgICB9DQogICAgICAgICNJbXBvcnQtTW9kdWxlIG1pY3Jvc29mdC5ncmFwaC5hdXRoZW50aWNhdGlvbiAtU2NvcGUgR2xv
YmFsDQoNCiAgICAgICAgIyNBZGQgZnVuY3Rpb25zIGZyb20gbW9kdWxlDQogICAgICAgIEZ1bmN0aW9uIENvbm5lY3QtVG9HcmFwaCB7DQogICAgICAgICAg
ICA8Iw0KLlNZTk9QU0lTDQpBdXRoZW50aWNhdGVzIHRvIHRoZSBHcmFwaCBBUEkgdmlhIHRoZSBNaWNyb3NvZnQuR3JhcGguQXV0aGVudGljYXRpb24gbW9k
dWxlLg0KIA0KLkRFU0NSSVBUSU9ODQpUaGUgQ29ubmVjdC1Ub0dyYXBoIGNtZGxldCBpcyBhIHdyYXBwZXIgY21kbGV0IHRoYXQgaGVscHMgYXV0aGVudGlj
YXRlIHRvIHRoZSBJbnR1bmUgR3JhcGggQVBJIHVzaW5nIHRoZSBNaWNyb3NvZnQuR3JhcGguQXV0aGVudGljYXRpb24gbW9kdWxlLiBJdCBsZXZlcmFnZXMg
YW4gRW50cmEgYXBwIElEIGFuZCBhcHAgc2VjcmV0IGZvciBhdXRoZW50aWNhdGlvbiBvciB1c2VyLWJhc2VkIGF1dGguDQogDQouUEFSQU1FVEVSIFRlbmFu
dA0KU3BlY2lmaWVzIHRoZSB0ZW5hbnQgKGUuZy4gY29udG9zby5vbm1pY3Jvc29mdC5jb20pIHRvIHdoaWNoIHRvIGF1dGhlbnRpY2F0ZS4NCiANCi5QQVJB
TUVURVIgQXBwSWQNClNwZWNpZmllcyB0aGUgRW50cmEgYXBwIElEIChHVUlEKSBmb3IgdGhlIGFwcGxpY2F0aW9uIHRoYXQgd2lsbCBiZSB1c2VkIHRvIGF1
dGhlbnRpY2F0ZS4NCiANCi5QQVJBTUVURVIgQXBwU2VjcmV0DQpTcGVjaWZpZXMgdGhlIEVudHJhIGFwcCBzZWNyZXQgY29ycmVzcG9uZGluZyB0byB0aGUg
YXBwIElEIHRoYXQgd2lsbCBiZSB1c2VkIHRvIGF1dGhlbnRpY2F0ZS4NCg0KLlBBUkFNRVRFUiBTY29wZXMNClNwZWNpZmllcyB0aGUgdXNlciBzY29wZXMg
Zm9yIGludGVyYWN0aXZlIGF1dGhlbnRpY2F0aW9uLg0KIA0KLkVYQU1QTEUNCkNvbm5lY3QtVG9HcmFwaCAtVGVuYW50ICR0ZW5hbnRJRCAtQXBwSWQgJGFw
cCAtQXBwU2VjcmV0ICRzZWNyZXQNCiANCi0jPg0KICAgICAgICAgICAgW2NtZGxldGJpbmRpbmcoKV0NCiAgICAgICAgICAgIHBhcmFtDQogICAgICAgICAg
ICAoDQogICAgICAgICAgICAgICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkZmFsc2UpXSBbc3RyaW5nXSRUZW5hbnQsDQogICAgICAgICAgICAgICAgW1Bh
cmFtZXRlcihNYW5kYXRvcnkgPSAkZmFsc2UpXSBbc3RyaW5nXSRBcHBJZCwNCiAgICAgICAgICAgICAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICRmYWxz
ZSldIFtzdHJpbmddJEFwcFNlY3JldCwNCiAgICAgICAgICAgICAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICRmYWxzZSldIFtzdHJpbmddJENlcnRpZmlj
YXRlU3ViamVjdE5hbWUsDQogICAgICAgICAgICAgICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkZmFsc2UpXSBbc3RyaW5nXSRDZXJ0aWZpY2F0ZVRodW1i
cHJpbnQsDQogICAgICAgICAgICAgICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkZmFsc2UpXSBbc3RyaW5nXSRzY29wZXMNCiAgICAgICAgICAgICkNCg0K
ICAgICAgICAgICAgUHJvY2VzcyB7DQogICAgICAgICAgICAgICAgSW1wb3J0LU1vZHVsZSBNaWNyb3NvZnQuR3JhcGguQXV0aGVudGljYXRpb24NCiAgICAg
ICAgICAgICAgICAkdmVyc2lvbiA9IChHZXQtTW9kdWxlIG1pY3Jvc29mdC5ncmFwaC5hdXRoZW50aWNhdGlvbiB8IFNlbGVjdC1PYmplY3QgLUV4cGFuZFBy
b3BlcnR5IFZlcnNpb24pLm1ham9yDQoNCiAgICAgICAgICAgICAgICBpZiAoJEFwcElkIC1uZSAiIikgew0KICAgICAgICAgICAgICAgICAgICBpZiAoJENl
cnRpZmljYXRlVGh1bWJwcmludCkgew0KICAgICAgICAgICAgICAgICAgICAgICAgJGdyYXBoID0gQ29ubmVjdC1NZ0dyYXBoIC1DZXJ0aWZpY2F0ZVRodW1i
cHJpbnQgJENlcnRpZmljYXRlVGh1bWJwcmludCAtVGVuYW50SWQgJFRlbmFudCAtQXBwSWQgJEFwcElkIA0KICAgICAgICAgICAgICAgICAgICAgICAgV3Jp
dGUtSG9zdCAiQ29ubmVjdGVkIHRvIEludHVuZSB0ZW5hbnQgJFRlbmFudElkIHVzaW5nIGNlcnRpZmljYXRlIHRodW1icHJpbnQgYXV0aGVudGljYXRpb24i
DQogICAgICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICAgICAgZWxzZWlmICgkQ2VydGlmaWNhdGVTdWJqZWN0TmFtZSkgew0KICAgICAgICAg
ICAgICAgICAgICAgICAgJGdyYXBoID0gQ29ubmVjdC1NZ0dyYXBoIC1DZXJ0aWZpY2F0ZU5hbWUgJENlcnRpZmljYXRlU3ViamVjdE5hbWUgLVRlbmFudElk
ICRUZW5hbnQgLUFwcElkICRBcHBJZA0KICAgICAgICAgICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiQ29ubmVjdGVkIHRvIEludHVuZSB0ZW5hbnQgJFRl
bmFudElkIHVzaW5nIGNlcnRpZmljYXRlIHN1YmplY3QgbmFtZSBhdXRoZW50aWNhdGlvbiINCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAg
ICAgICAgICBlbHNlIHsNCg0KICAgICAgICAgICAgICAgICAgICAgICAgJGJvZHkgPSBAew0KICAgICAgICAgICAgICAgICAgICAgICAgICAgIGdyYW50X3R5
cGUgICAgPSAiY2xpZW50X2NyZWRlbnRpYWxzIjsNCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBjbGllbnRfaWQgICAgID0gJEFwcElkOw0KICAgICAg
ICAgICAgICAgICAgICAgICAgICAgIGNsaWVudF9zZWNyZXQgPSAkQXBwU2VjcmV0Ow0KICAgICAgICAgICAgICAgICAgICAgICAgICAgIHNjb3BlICAgICAg
ICAgPSAiaHR0cHM6Ly9ncmFwaC5taWNyb3NvZnQuY29tLy5kZWZhdWx0IjsNCiAgICAgICAgICAgICAgICAgICAgICAgIH0NCiAgICAgDQogICAgICAgICAg
ICAgICAgICAgICAgICAkcmVzcG9uc2UgPSBJbnZva2UtUmVzdE1ldGhvZCAtTWV0aG9kIFBvc3QgLVVyaSBodHRwczovL2xvZ2luLm1pY3Jvc29mdG9ubGlu
ZS5jb20vJFRlbmFudC9vYXV0aDIvdjIuMC90b2tlbiAtQm9keSAkYm9keQ0KICAgICAgICAgICAgICAgICAgICAgICAgJGFjY2Vzc1Rva2VuID0gJHJlc3Bv
bnNlLmFjY2Vzc190b2tlbg0KICAgICANCiAgICAgICAgICAgICAgICAgICAgICAgICRhY2Nlc3NUb2tlbg0KICAgICAgICAgICAgICAgICAgICAgICAgaWYg
KCR2ZXJzaW9uIC1lcSAyKSB7DQogICAgICAgICAgICAgICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiVmVyc2lvbiAyIG1vZHVsZSBkZXRlY3RlZCINCiAg
ICAgICAgICAgICAgICAgICAgICAgICAgICAkYWNjZXNzdG9rZW5maW5hbCA9IENvbnZlcnRUby1TZWN1cmVTdHJpbmcgLVN0cmluZyAkYWNjZXNzVG9rZW4g
LUFzUGxhaW5UZXh0IC1Gb3JjZQ0KICAgICAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICAgICAgZWxzZSB7DQogICAgICAgICAg
ICAgICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiVmVyc2lvbiAxIE1vZHVsZSBEZXRlY3RlZCINCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBTZWxl
Y3QtTWdQcm9maWxlIC1OYW1lIEJldGENCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAkYWNjZXNzdG9rZW5maW5hbCA9ICRhY2Nlc3NUb2tlbg0KICAg
ICAgICAgICAgICAgICAgICAgICAgfQ0KDQogICAgICAgICAgICAgICAgICAgICAgICAkZ3JhcGggPSBDb25uZWN0LU1nR3JhcGggLUFjY2Vzc1Rva2VuICRh
Y2Nlc3N0b2tlbmZpbmFsIA0KICAgICAgICAgICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiQ29ubmVjdGVkIHRvIEludHVuZSB0ZW5hbnQgJFRlbmFudElk
IHVzaW5nIGFwcC1iYXNlZCBhdXRoZW50aWNhdGlvbiAoRW50cmEgYXV0aGVudGljYXRpb24gbm90IHN1cHBvcnRlZCkiDQogICAgICAgICAgICAgICAgICAg
IH0NCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgZWxzZSB7DQogICAgICAgICAgICAgICAgICAgIGlmICgkdmVyc2lvbiAtZXEgMikgew0K
ICAgICAgICAgICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiVmVyc2lvbiAyIG1vZHVsZSBkZXRlY3RlZCINCiAgICAgICAgICAgICAgICAgICAgfQ0KICAg
ICAgICAgICAgICAgICAgICBlbHNlIHsNCiAgICAgICAgICAgICAgICAgICAgICAgIFdyaXRlLUhvc3QgIlZlcnNpb24gMSBNb2R1bGUgRGV0ZWN0ZWQiDQog
ICAgICAgICAgICAgICAgICAgICAgICBTZWxlY3QtTWdQcm9maWxlIC1OYW1lIEJldGENCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAg
ICAgICAkZ3JhcGggPSBDb25uZWN0LU1nR3JhcGggLVNjb3BlcyAkc2NvcGVzDQogICAgICAgICAgICAgICAgICAgIFdyaXRlLUhvc3QgIkNvbm5lY3RlZCB0
byBJbnR1bmUgdGVuYW50ICQoJGdyYXBoLlRlbmFudElkKSINCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICB9DQogICAgICAgIH0gICAgDQogICAg
ICAgICNyZWdpb24gSGVscGVyIG1ldGhvZHMNCg0KICAgICAgICBGdW5jdGlvbiBCb29sVG9TdHJpbmcoKSB7DQogICAgICAgICAgICBwYXJhbQ0KICAgICAg
ICAgICAgKA0KICAgICAgICAgICAgICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJGZhbHNlLCBWYWx1ZUZyb21QaXBlbGluZUJ5UHJvcGVydHlOYW1lID0g
JFRydWUpXSBbYm9vbF0gJHZhbHVlDQogICAgICAgICAgICApDQoNCiAgICAgICAgICAgIFByb2Nlc3Mgew0KICAgICAgICAgICAgICAgIHJldHVybiAkdmFs
dWUuVG9TdHJpbmcoKS5Ub0xvd2VyKCkNCiAgICAgICAgICAgIH0NCiAgICAgICAgfQ0KDQogICAgICAgICNlbmRyZWdpb24NCg0KDQoNCiAgICAgICAgI3Jl
Z2lvbiBDb3JlIG1ldGhvZHMNCg0KICAgICAgICBGdW5jdGlvbiBHZXQtQXV0b3BpbG90RGV2aWNlKCkgew0KICAgICAgICAgICAgPCMNCi5TWU5PUFNJUw0K
R2V0cyBkZXZpY2VzIGN1cnJlbnRseSByZWdpc3RlcmVkIHdpdGggV2luZG93cyBBdXRvcGlsb3QuDQogDQouREVTQ1JJUFRJT04NClRoZSBHZXQtQXV0b3Bp
bG90RGV2aWNlIGNtZGxldCByZXRyaWV2ZXMgZWl0aGVyIHRoZSBmdWxsIGxpc3Qgb2YgZGV2aWNlcyByZWdpc3RlcmVkIHdpdGggV2luZG93cyBBdXRvcGls
b3QgZm9yIHRoZSBjdXJyZW50IEVudHJhIHRlbmFudCwgb3IgYSBzcGVjaWZpYyBkZXZpY2UgaWYgdGhlIElEIG9mIHRoZSBkZXZpY2UgaXMgc3BlY2lmaWVk
Lg0KIA0KLlBBUkFNRVRFUiBpZA0KT3B0aW9uYWxseSBzcGVjaWZpZXMgdGhlIElEIChHVUlEKSBmb3IgYSBzcGVjaWZpYyBXaW5kb3dzIEF1dG9waWxvdCBk
ZXZpY2UgKHdoaWNoIGlzIHR5cGljYWxseSByZXR1cm5lZCBhZnRlciBpbXBvcnRpbmcgYSBuZXcgZGV2aWNlKQ0KIA0KLlBBUkFNRVRFUiBzZXJpYWwNCk9w
dGlvbmFsbHkgc3BlY2lmaWVzIHRoZSBzZXJpYWwgbnVtYmVyIG9mIHRoZSBzcGVjaWZpYyBXaW5kb3dzIEF1dG9waWxvdCBkZXZpY2UgdG8gcmV0cmlldmUN
CiANCi5QQVJBTUVURVIgZXhwYW5kDQpFeHBhbmQgdGhlIHByb3BlcnRpZXMgb2YgdGhlIGRldmljZSB0byBpbmNsdWRlIHRoZSBBdXRvcGlsb3QgcHJvZmls
ZSBpbmZvcm1hdGlvbg0KIA0KLkVYQU1QTEUNCkdldCBhIGxpc3Qgb2YgYWxsIGRldmljZXMgcmVnaXN0ZXJlZCB3aXRoIFdpbmRvd3MgQXV0b3BpbG90DQog
DQpHZXQtQXV0b3BpbG90RGV2aWNlDQojPg0KICAgICAgICAgICAgW2NtZGxldGJpbmRpbmcoKV0NCiAgICAgICAgICAgIHBhcmFtDQogICAgICAgICAgICAo
DQogICAgICAgICAgICAgICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkZmFsc2UsIFZhbHVlRnJvbVBpcGVsaW5lQnlQcm9wZXJ0eU5hbWUgPSAkVHJ1ZSld
ICRpZCwNCiAgICAgICAgICAgICAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICRmYWxzZSldICRzZXJpYWwsDQogICAgICAgICAgICAgICAgW1BhcmFtZXRl
cihNYW5kYXRvcnkgPSAkZmFsc2UpXSBbU3dpdGNoXSRleHBhbmQgPSAkZmFsc2UNCiAgICAgICAgICAgICkNCg0KICAgICAgICAgICAgUHJvY2VzcyB7DQoN
CiAgICAgICAgICAgICAgICAjIERlZmluaW5nIFZhcmlhYmxlcw0KICAgICAgICAgICAgICAgICRncmFwaEFwaVZlcnNpb24gPSAiYmV0YSINCiAgICAgICAg
ICAgICAgICAkUmVzb3VyY2UgPSAiZGV2aWNlTWFuYWdlbWVudC93aW5kb3dzQXV0b3BpbG90RGV2aWNlSWRlbnRpdGllcyINCiAgICANCiAgICAgICAgICAg
ICAgICBpZiAoJGlkIC1hbmQgJGV4cGFuZCkgew0KICAgICAgICAgICAgICAgICAgICAkdXJpID0gImh0dHBzOi8vZ3JhcGgubWljcm9zb2Z0LmNvbS8kZ3Jh
cGhBcGlWZXJzaW9uLyQoJFJlc291cmNlKS8kKCRpZCk/YCRleHBhbmQ9ZGVwbG95bWVudFByb2ZpbGUsaW50ZW5kZWREZXBsb3ltZW50UHJvZmlsZSINCiAg
ICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgZWxzZWlmICgkaWQpIHsNCiAgICAgICAgICAgICAgICAgICAgJHVyaSA9ICJodHRwczovL2dyYXBo
Lm1pY3Jvc29mdC5jb20vJGdyYXBoQXBpVmVyc2lvbi8kKCRSZXNvdXJjZSkvJGlkIg0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBlbHNl
aWYgKCRzZXJpYWwpIHsNCiAgICAgICAgICAgICAgICAgICAgJGVuY29kZWQgPSBbdXJpXTo6RXNjYXBlRGF0YVN0cmluZygkc2VyaWFsKQ0KICAgICAgICAg
ICAgICAgICAgICAjI0NoZWNrIGlmIHNlcmlhbCBjb250YWlucyBhIHNwYWNlDQogICAgICAgICAgICAgICAgICAgICRzZXJpYWxlbGVtZW50cyA9ICRzZXJp
YWwuU3BsaXQoIiAiKQ0KICAgICAgICAgICAgICAgICAgICBpZiAoJHNlcmlhbGVsZW1lbnRzLkNvdW50IC1ndCAxKSB7DQogICAgICAgICAgICAgICAgICAg
ICAgICAkdXJpID0gImh0dHBzOi8vZ3JhcGgubWljcm9zb2Z0LmNvbS8kZ3JhcGhBcGlWZXJzaW9uLyQoJFJlc291cmNlKT9gJGZpbHRlcj1jb250YWlucyhz
ZXJpYWxOdW1iZXIsJyQoJHNlcmlhbGVsZW1lbnRzWzBdKScpIg0KICAgICAgICAgICAgICAgICAgICAgICAgJHNlcmlhbGhhc3NwYWNlcyA9IDENCiAgICAg
ICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICBlbHNlIHsNCiAgICAgICAgICAgICAgICAgICAgICAgICR1cmkgPSAiaHR0cHM6Ly9ncmFw
aC5taWNyb3NvZnQuY29tLyRncmFwaEFwaVZlcnNpb24vJCgkUmVzb3VyY2UpP2AkZmlsdGVyPWNvbnRhaW5zKHNlcmlhbE51bWJlciwnJGVuY29kZWQnKSIN
CiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBlbHNlIHsNCiAgICAgICAgICAgICAgICAgICAgJHVy
aSA9ICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vJGdyYXBoQXBpVmVyc2lvbi8kKCRSZXNvdXJjZSkiDQogICAgICAgICAgICAgICAgfQ0KDQogICAg
ICAgICAgICAgICAgV3JpdGUtVmVyYm9zZSAiR0VUICR1cmkiDQoNCiAgICAgICAgICAgICAgICB0cnkgew0KICAgICAgICAgICAgICAgICAgICAkcmVzcG9u
c2UgPSBJbnZva2UtTWdHcmFwaFJlcXVlc3QgLVVyaSAkdXJpIC1NZXRob2QgR2V0IC1PdXRwdXRUeXBlIFBTT2JqZWN0DQogICAgICAgICAgICAgICAgICAg
IGlmICgkaWQpIHsNCiAgICAgICAgICAgICAgICAgICAgICAgICRyZXNwb25zZQ0KICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgICAg
IGVsc2Ugew0KICAgICAgICAgICAgICAgICAgICAgICAgaWYgKCRzZXJpYWxoYXNzcGFjZXMgLWVxIDEpIHsgIA0KICAgICAgICAgICAgICAgICAgICAgICAg
ICAgICRkZXZpY2VzID0gJHJlc3BvbnNlLnZhbHVlIHwgV2hlcmUtT2JqZWN0IHsgJF8uc2VyaWFsTnVtYmVyIC1lcSAiJCgkc2VyaWFsKSIgfQ0KICAgICAg
ICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICAgICAgZWxzZSB7DQogICAgICAgICAgICAgICAgICAgICAgICAgICAgJGRldmljZXMg
PSAkcmVzcG9uc2UudmFsdWUgDQogICAgICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgICAgICAgICAkZGV2aWNlc05leHRMaW5rID0g
JHJlc3BvbnNlLiJAb2RhdGEubmV4dExpbmsiDQogICAgDQogICAgICAgICAgICAgICAgICAgICAgICB3aGlsZSAoJG51bGwgLW5lICRkZXZpY2VzTmV4dExp
bmspIHsNCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAkZGV2aWNlc1Jlc3BvbnNlID0gKEludm9rZS1NZ0dyYXBoUmVxdWVzdCAtVXJpICRkZXZpY2Vz
TmV4dExpbmsgLU1ldGhvZCBHZXQgLU91dHB1dFR5cGUgUFNPYmplY3QpDQogICAgICAgICAgICAgICAgICAgICAgICAgICAgJGRldmljZXNOZXh0TGluayA9
ICRkZXZpY2VzUmVzcG9uc2UuIkBvZGF0YS5uZXh0TGluayINCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBpZiAoJHNlcmlhbGhhc3NwYWNlcyAtZXEg
MSkgew0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAkZGV2aWNlcyArPSAkZGV2aWNlc1Jlc3BvbnNlLnZhbHVlIHwgV2hlcmUtT2JqZWN0IHsg
JF8uc2VyaWFsTnVtYmVyIC1lcSAiJCgkc2VyaWFsKSIgfQ0KICAgICAgICAgICAgICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICAgICAgICAg
ICAgICBlbHNlIHsNCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgJGRldmljZXMgKz0gJGRldmljZXNSZXNwb25zZS52YWx1ZQ0KICAgICAgICAg
ICAgICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICAgICAgICAgIH0NCiAgICANCiAgICAgICAgICAgICAgICAgICAgICAgIGlmICgkZXhwYW5k
KSB7DQogICAgICAgICAgICAgICAgICAgICAgICAgICAgJGRldmljZXMgfCBHZXQtQXV0b3BpbG90RGV2aWNlIC1leHBhbmQNCiAgICAgICAgICAgICAgICAg
ICAgICAgIH0NCiAgICAgICAgICAgICAgICAgICAgICAgIGVsc2Ugew0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICRkZXZpY2VzDQogICAgICAgICAg
ICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgY2F0Y2ggew0KICAgICAg
ICAgICAgICAgICAgICBXcml0ZS1FcnJvciAkXy5FeGNlcHRpb24gDQogICAgICAgICAgICAgICAgICAgIGJyZWFrDQogICAgICAgICAgICAgICAgfQ0KICAg
ICAgICAgICAgfQ0KICAgICAgICB9DQoNCg0KICAgICAgICBGdW5jdGlvbiBTZXQtQXV0b3BpbG90RGV2aWNlKCkgew0KICAgICAgICAgICAgPCMNCi5TWU5P
UFNJUw0KVXBkYXRlcyBzZXR0aW5ncyBvbiBhbiBBdXRvcGlsb3QgZGV2aWNlLg0KIA0KLkRFU0NSSVBUSU9ODQpUaGUgU2V0LUF1dG9waWxvdERldmljZSBj
bWRsZXQgY2FuIGJlIHVzZWQgdG8gY2hhbmdlIHRoZSB1cGRhdGFibGUgcHJvcGVydGllcyBvbiBhIFdpbmRvd3MgQXV0b3BpbG90IGRldmljZSBvYmplY3Qu
DQogDQouUEFSQU1FVEVSIGlkDQpUaGUgV2luZG93cyBBdXRvcGlsb3QgZGV2aWNlIGlkIChtYW5kYXRvcnkpLg0KIA0KLlBBUkFNRVRFUiB1c2VyUHJpbmNp
cGFsTmFtZQ0KVGhlIHVzZXIgcHJpbmNpcGFsIG5hbWUuDQogDQouUEFSQU1FVEVSIGFkZHJlc3NpYmxlVXNlck5hbWUNClRoZSBuYW1lIHRvIGRpc3BsYXkg
ZHVyaW5nIFdpbmRvd3MgQXV0b3BpbG90IGVucm9sbG1lbnQuIElmIHNwZWNpZmllZCwgdGhlIHVzZXJQcmluY2lwYWxOYW1lIG11c3QgYWxzbyBiZSBzcGVj
aWZpZWQuDQogDQouUEFSQU1FVEVSIGRpc3BsYXlOYW1lDQpUaGUgbmFtZSAoY29tcHV0ZXIgbmFtZSkgdG8gYmUgYXNzaWduZWQgdG8gdGhlIGRldmljZSB3
aGVuIGl0IGlzIGRlcGxveWVkIHZpYSBXaW5kb3dzIEF1dG9waWxvdC4gVGhpcyBpcyBwcmVzZW50bHkgb25seSBzdXBwb3J0ZWQgd2l0aCBFbnRyYSBKb2lu
IHNjZW5hcmlvcy4gTm90ZSB0aGF0IG5hbWVzIHNob3VsZCBub3QgZXhjZWVkIDE1IGNoYXJhY3RlcnMuIEFmdGVyIHNldHRpbmcgdGhlIG5hbWUsIHlvdSBu
ZWVkIHRvIGluaXRpYXRlIGEgc3luYyAoSW52b2tlLUF1dG9waWxvdFN5bmMpIGluIG9yZGVyIHRvIHNlZSB0aGUgbmFtZSBpbiB0aGUgSW50dW5lIG9iamVj
dC4NCiANCi5QQVJBTUVURVIgZ3JvdXBUYWcNClRoZSBncm91cCB0YWcgdmFsdWUgdG8gc2V0IGZvciB0aGUgZGV2aWNlLg0KIA0KLkVYQU1QTEUNCkFzc2ln
biBhIHVzZXIgYW5kIGEgbmFtZSB0byBkaXNwbGF5IGR1cmluZyBlbnJvbGxtZW50IHRvIGEgV2luZG93cyBBdXRvcGlsb3QgZGV2aWNlLg0KIA0KU2V0LUF1
dG9waWxvdERldmljZSAtaWQgJGlkIC11c2VyUHJpbmNpcGFsTmFtZSAkdXNlclByaW5jaXBhbE5hbWUgLWFkZHJlc3NhYmxlVXNlck5hbWUgIkpvaG4gRG9l
IiAtZGlzcGxheU5hbWUgIkNPTlRPU08tMDAwMSIgLWdyb3VwVGFnICJUZXN0aW5nIg0KIz4NCiAgICAgICAgICAgIFtjbWRsZXRiaW5kaW5nKCldDQogICAg
ICAgICAgICBwYXJhbQ0KICAgICAgICAgICAgKA0KICAgICAgICAgICAgICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJHRydWUsIFZhbHVlRnJvbVBpcGVs
aW5lQnlQcm9wZXJ0eU5hbWUgPSAkVHJ1ZSldICRpZCwNCiAgICAgICAgICAgICAgICBbUGFyYW1ldGVyKFBhcmFtZXRlclNldE5hbWUgPSAiUHJvcCIpXSAk
dXNlclByaW5jaXBhbE5hbWUgPSAkbnVsbCwNCiAgICAgICAgICAgICAgICBbUGFyYW1ldGVyKFBhcmFtZXRlclNldE5hbWUgPSAiUHJvcCIpXSAkYWRkcmVz
c2FibGVVc2VyTmFtZSA9ICRudWxsLA0KICAgICAgICAgICAgICAgIFtQYXJhbWV0ZXIoUGFyYW1ldGVyU2V0TmFtZSA9ICJQcm9wIildW0FsaWFzKCJDb21w
dXRlck5hbWUiLCAiQ04iLCAiTWFjaGluZU5hbWUiKV0gJGRpc3BsYXlOYW1lID0gJG51bGwsDQogICAgICAgICAgICAgICAgW1BhcmFtZXRlcihQYXJhbWV0
ZXJTZXROYW1lID0gIlByb3AiKV0gJGdyb3VwVGFnID0gJG51bGwNCiAgICAgICAgICAgICkNCg0KICAgICAgICAgICAgUHJvY2VzcyB7DQogICAgDQogICAg
ICAgICAgICAgICAgIyBEZWZpbmluZyBWYXJpYWJsZXMNCiAgICAgICAgICAgICAgICAkZ3JhcGhBcGlWZXJzaW9uID0gImJldGEiDQogICAgICAgICAgICAg
ICAgJFJlc291cmNlID0gImRldmljZU1hbmFnZW1lbnQvd2luZG93c0F1dG9waWxvdERldmljZUlkZW50aXRpZXMiDQogICAgDQogICAgICAgICAgICAgICAg
JHVyaSA9ICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vJGdyYXBoQXBpVmVyc2lvbi8kUmVzb3VyY2UvJGlkL1VwZGF0ZURldmljZVByb3BlcnRpZXMi
DQoNCiAgICAgICAgICAgICAgICAkanNvbiA9ICJ7Ig0KICAgICAgICAgICAgICAgIGlmICgkUFNCb3VuZFBhcmFtZXRlcnMuQ29udGFpbnNLZXkoJ3VzZXJQ
cmluY2lwYWxOYW1lJykpIHsNCiAgICAgICAgICAgICAgICAgICAgJGpzb24gPSAkanNvbiArICIgdXNlclByaW5jaXBhbE5hbWU6IGAiJHVzZXJQcmluY2lw
YWxOYW1lYCIsIg0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBpZiAoJFBTQm91bmRQYXJhbWV0ZXJzLkNvbnRhaW5zS2V5KCdhZGRyZXNz
YWJsZVVzZXJOYW1lJykpIHsNCiAgICAgICAgICAgICAgICAgICAgJGpzb24gPSAkanNvbiArICIgYWRkcmVzc2FibGVVc2VyTmFtZTogYCIkYWRkcmVzc2Fi
bGVVc2VyTmFtZWAiLCINCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgaWYgKCRQU0JvdW5kUGFyYW1ldGVycy5Db250YWluc0tleSgnZGlz
cGxheU5hbWUnKSkgew0KICAgICAgICAgICAgICAgICAgICAkanNvbiA9ICRqc29uICsgIiBkaXNwbGF5TmFtZTogYCIkZGlzcGxheU5hbWVgIiwiDQogICAg
ICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgIGlmICgkUFNCb3VuZFBhcmFtZXRlcnMuQ29udGFpbnNLZXkoJ2dyb3VwVGFnJykpIHsNCiAgICAgICAg
ICAgICAgICAgICAgJGpzb24gPSAkanNvbiArICIgZ3JvdXBUYWc6IGAiJGdyb3VwVGFnYCIiDQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAg
IGVsc2Ugew0KICAgICAgICAgICAgICAgICAgICAkanNvbiA9ICRqc29uLlRyaW0oIiwiKQ0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICAk
anNvbiA9ICRqc29uICsgIiB9Ig0KDQogICAgICAgICAgICAgICAgV3JpdGUtVmVyYm9zZSAiUE9TVCAkdXJpYG4kanNvbiINCg0KICAgICAgICAgICAgICAg
IHRyeSB7DQogICAgICAgICAgICAgICAgICAgIEludm9rZS1NZ0dyYXBoUmVxdWVzdCAtVXJpICR1cmkgLU1ldGhvZCBQT1NUIC1Cb2R5ICRqc29uIC1Db250
ZW50VHlwZSAiYXBwbGljYXRpb24vanNvbiIgLU91dHB1dFR5cGUgUFNPYmplY3QNCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgY2F0Y2gg
ew0KICAgICAgICAgICAgICAgICAgICBXcml0ZS1FcnJvciAkXy5FeGNlcHRpb24gDQogICAgICAgICAgICAgICAgICAgIGJyZWFrDQogICAgICAgICAgICAg
ICAgfQ0KICAgICAgICAgICAgfQ0KICAgICAgICB9DQoNCiAgICANCiAgICAgICAgRnVuY3Rpb24gUmVtb3ZlLUF1dG9waWxvdERldmljZSgpIHsNCiAgICAg
ICAgICAgIDwjDQouU1lOT1BTSVMNClJlbW92ZXMgYSBzcGVjaWZpYyBkZXZpY2UgY3VycmVudGx5IHJlZ2lzdGVyZWQgd2l0aCBXaW5kb3dzIEF1dG9waWxv
dC4NCiANCi5ERVNDUklQVElPTg0KVGhlIFJlbW92ZS1BdXRvcGlsb3REZXZpY2UgY21kbGV0IHJlbW92ZXMgdGhlIHNwZWNpZmllZCBkZXZpY2UsIGlkZW50
aWZpZWQgYnkgaXRzIElELCBmcm9tIHRoZSBsaXN0IG9mIGRldmljZXMgcmVnaXN0ZXJlZCB3aXRoIFdpbmRvd3MgQXV0b3BpbG90IGZvciB0aGUgY3VycmVu
dCBFbnRyYSB0ZW5hbnQuDQogDQouUEFSQU1FVEVSIGlkDQpTcGVjaWZpZXMgdGhlIElEIChHVUlEKSBmb3IgYSBzcGVjaWZpYyBXaW5kb3dzIEF1dG9waWxv
dCBkZXZpY2UNCiANCi5FWEFNUExFDQpSZW1vdmUgYWxsIFdpbmRvd3MgQXV0b3BpbG90IGRldmljZXMgZnJvbSB0aGUgY3VycmVudCBFbnRyYSB0ZW5hbnQN
CiANCkdldC1BdXRvcGlsb3REZXZpY2UgfCBSZW1vdmUtQXV0b3BpbG90RGV2aWNlDQojPg0KICAgICAgICAgICAgW2NtZGxldGJpbmRpbmcoKV0NCiAgICAg
ICAgICAgIHBhcmFtDQogICAgICAgICAgICAoDQogICAgICAgICAgICAgICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkdHJ1ZSwgVmFsdWVGcm9tUGlwZWxp
bmVCeVByb3BlcnR5TmFtZSA9ICRUcnVlKV0gJGlkLA0KICAgICAgICAgICAgICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJGZhbHNlLCBWYWx1ZUZyb21Q
aXBlbGluZUJ5UHJvcGVydHlOYW1lID0gJFRydWUpXSAkc2VyaWFsTnVtYmVyDQogICAgICAgICAgICApDQoNCiAgICAgICAgICAgIEJlZ2luIHsNCiAgICAg
ICAgICAgICAgICAkYnVsa0xpc3QgPSBAKCkNCiAgICAgICAgICAgIH0NCg0KICAgICAgICAgICAgUHJvY2VzcyB7DQoNCiAgICAgICAgICAgICAgICAjIERl
ZmluaW5nIFZhcmlhYmxlcw0KICAgICAgICAgICAgICAgICRncmFwaEFwaVZlcnNpb24gPSAiYmV0YSINCiAgICAgICAgICAgICAgICAkUmVzb3VyY2UgPSAi
ZGV2aWNlTWFuYWdlbWVudC93aW5kb3dzQXV0b3BpbG90RGV2aWNlSWRlbnRpdGllcyIgICAgDQogICAgICAgICAgICAgICAgJHVyaSA9ICJodHRwczovL2dy
YXBoLm1pY3Jvc29mdC5jb20vJGdyYXBoQXBpVmVyc2lvbi8kUmVzb3VyY2UvJGlkIg0KDQogICAgICAgICAgICAgICAgdHJ5IHsNCiAgICAgICAgICAgICAg
ICAgICAgV3JpdGUtVmVyYm9zZSAiREVMRVRFICR1cmkiDQogICAgICAgICAgICAgICAgICAgIEludm9rZS1NZ0dyYXBoUmVxdWVzdCAtVXJpICR1cmkgLU1l
dGhvZCBERUxFVEUNCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgY2F0Y2ggew0KICAgICAgICAgICAgICAgICAgICBXcml0ZS1FcnJvciAk
Xy5FeGNlcHRpb24gDQogICAgICAgICAgICAgICAgICAgIGJyZWFrDQogICAgICAgICAgICAgICAgfQ0KICAgICAgICANCiAgICAgICAgICAgIH0NCiAgICAg
ICAgfQ0KDQoNCiAgICAgICAgRnVuY3Rpb24gR2V0LUF1dG9waWxvdEltcG9ydGVkRGV2aWNlKCkgew0KICAgICAgICAgICAgPCMNCi5TWU5PUFNJUw0KR2V0
cyBpbmZvcm1hdGlvbiBhYm91dCBkZXZpY2VzIGJlaW5nIGltcG9ydGVkIGludG8gV2luZG93cyBBdXRvcGlsb3QuDQogDQouREVTQ1JJUFRJT04NClRoZSBH
ZXQtQXV0b3BpbG90SW1wb3J0ZWREZXZpY2UgY21kbGV0IHJldHJpZXZlcyBlaXRoZXIgdGhlIGZ1bGwgbGlzdCBvZiBkZXZpY2VzIGJlaW5nIGltcG9ydGVk
IGludG8gV2luZG93cyBBdXRvcGlsb3QgZm9yIHRoZSBjdXJyZW50IEVudHJhIHRlbmFudCwgb3IgaW5mb3JtYXRpb24gZm9yIGEgc3BlY2lmaWMgZGV2aWNl
IGlmIHRoZSBJRCBvZiB0aGUgZGV2aWNlIGlzIHNwZWNpZmllZC4gT25jZSB0aGUgaW1wb3J0IGlzIGNvbXBsZXRlLCB0aGUgaW5mb3JtYXRpb24gaW5zdGFu
Y2UgaXMgZXhwZWN0ZWQgdG8gYmUgZGVsZXRlZC4NCiANCi5QQVJBTUVURVIgaWQNCk9wdGlvbmFsbHkgc3BlY2lmaWVzIHRoZSBJRCAoR1VJRCkgZm9yIGEg
c3BlY2lmaWMgV2luZG93cyBBdXRvcGlsb3QgZGV2aWNlIGJlaW5nIGltcG9ydGVkLg0KIA0KLkVYQU1QTEUNCkdldCBhIGxpc3Qgb2YgYWxsIGRldmljZXMg
YmVpbmcgaW1wb3J0ZWQgaW50byBXaW5kb3dzIEF1dG9waWxvdCBmb3IgdGhlIGN1cnJlbnQgRW50cmEgdGVuYW50Lg0KIA0KR2V0LUF1dG9waWxvdEltcG9y
dGVkRGV2aWNlDQojPg0KICAgICAgICAgICAgW2NtZGxldGJpbmRpbmcoKV0NCiAgICAgICAgICAgIHBhcmFtDQogICAgICAgICAgICAoDQogICAgICAgICAg
ICAgICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkZmFsc2UpXSAkaWQgPSAkbnVsbCwNCiAgICAgICAgICAgICAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9
ICRmYWxzZSldICRzZXJpYWwNCiAgICAgICAgICAgICkNCg0KICAgICAgICAgICAgIyBEZWZpbmluZyBWYXJpYWJsZXMNCiAgICAgICAgICAgICRncmFwaEFw
aVZlcnNpb24gPSAiYmV0YSINCiAgICAgICAgICAgIGlmICgkaWQpIHsNCiAgICAgICAgICAgICAgICAkdXJpID0gImh0dHBzOi8vZ3JhcGgubWljcm9zb2Z0
LmNvbS8kZ3JhcGhBcGlWZXJzaW9uL2RldmljZU1hbmFnZW1lbnQvaW1wb3J0ZWRXaW5kb3dzQXV0b3BpbG90RGV2aWNlSWRlbnRpdGllcy8kaWQiDQogICAg
ICAgICAgICB9IA0KICAgICAgICAgICAgZWxzZWlmICgkc2VyaWFsKSB7DQogICAgICAgICAgICAgICAgIyBoYW5kbGVzIGFsc28gc2VyaWFsIG51bWJlcnMg
d2l0aCBzcGFjZXMgICAgDQogICAgICAgICAgICAgICAgJHVyaSA9ICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vJGdyYXBoQXBpVmVyc2lvbi9kZXZp
Y2VNYW5hZ2VtZW50L2ltcG9ydGVkV2luZG93c0F1dG9waWxvdERldmljZUlkZW50aXRpZXMvP2AkZmlsdGVyPWNvbnRhaW5zKHNlcmlhbE51bWJlciwnJHNl
cmlhbCcpIg0KICAgICAgICAgICAgfQ0KICAgICAgICAgICAgZWxzZSB7DQogICAgICAgICAgICAgICAgJHVyaSA9ICJodHRwczovL2dyYXBoLm1pY3Jvc29m
dC5jb20vJGdyYXBoQXBpVmVyc2lvbi9kZXZpY2VNYW5hZ2VtZW50L2ltcG9ydGVkV2luZG93c0F1dG9waWxvdERldmljZUlkZW50aXRpZXMiDQogICAgICAg
ICAgICB9DQoNCiAgICAgICAgICAgIFdyaXRlLVZlcmJvc2UgIkdFVCAkdXJpIg0KDQogICAgICAgICAgICB0cnkgew0KICAgICAgICAgICAgICAgICRyZXNw
b25zZSA9IEludm9rZS1NZ0dyYXBoUmVxdWVzdCAtVXJpICR1cmkgLU1ldGhvZCBHZXQgLU91dHB1dFR5cGUgUFNPYmplY3QNCiAgICAgICAgICAgICAgICBp
ZiAoJGlkKSB7DQogICAgICAgICAgICAgICAgICAgICRyZXNwb25zZQ0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBlbHNlIHsNCiAgICAg
ICAgICAgICAgICAgICAgJGRldmljZXMgPSAkcmVzcG9uc2UudmFsdWUNCiAgICANCiAgICAgICAgICAgICAgICAgICAgJGRldmljZXNOZXh0TGluayA9ICRy
ZXNwb25zZS4iQG9kYXRhLm5leHRMaW5rIg0KICAgIA0KICAgICAgICAgICAgICAgICAgICB3aGlsZSAoJG51bGwgLW5lICRkZXZpY2VzTmV4dExpbmspIHsN
CiAgICAgICAgICAgICAgICAgICAgICAgICRkZXZpY2VzUmVzcG9uc2UgPSAoSW52b2tlLU1nR3JhcGhSZXF1ZXN0IC1VcmkgJGRldmljZXNOZXh0TGluayAt
TWV0aG9kIEdldCAtT3V0cHV0VHlwZSBQU09iamVjdCkNCiAgICAgICAgICAgICAgICAgICAgICAgICRkZXZpY2VzTmV4dExpbmsgPSAkZGV2aWNlc1Jlc3Bv
bnNlLiJAb2RhdGEubmV4dExpbmsiDQogICAgICAgICAgICAgICAgICAgICAgICAkZGV2aWNlcyArPSAkZGV2aWNlc1Jlc3BvbnNlLnZhbHVlDQogICAgICAg
ICAgICAgICAgICAgIH0NCiAgICANCiAgICAgICAgICAgICAgICAgICAgJGRldmljZXMNCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICB9DQogICAg
ICAgICAgICBjYXRjaCB7DQogICAgICAgICAgICAgICAgV3JpdGUtRXJyb3IgJF8uRXhjZXB0aW9uIA0KICAgICAgICAgICAgICAgIGJyZWFrDQogICAgICAg
ICAgICB9DQoNCiAgICAgICAgfQ0KDQoNCiAgICAgICAgPCMNCi5TWU5PUFNJUw0KQWRkcyBhIG5ldyBkZXZpY2UgdG8gV2luZG93cyBBdXRvcGlsb3QuDQog
DQouREVTQ1JJUFRJT04NClRoZSBBZGQtQXV0b3BpbG90SW1wb3J0ZWREZXZpY2UgY21kbGV0IGFkZHMgdGhlIHNwZWNpZmllZCBkZXZpY2UgdG8gV2luZG93
cyBBdXRvcGlsb3QgZm9yIHRoZSBjdXJyZW50IEVudHJhIHRlbmFudC4gTm90ZSB0aGF0IGEgc3RhdHVzIG9iamVjdCBpcyByZXR1cm5lZCB3aGVuIHRoaXMg
Y21kbGV0IGNvbXBsZXRlczsgdGhlIGFjdHVhbCBpbXBvcnQgcHJvY2VzcyBpcyBwZXJmb3JtZWQgYXMgYSBiYWNrZ3JvdW5kIGJhdGNoIHByb2Nlc3MgYnkg
dGhlIE1pY3Jvc29mdCBJbnR1bmUgc2VydmljZS4NCiANCi5QQVJBTUVURVIgc2VyaWFsTnVtYmVyDQpUaGUgaGFyZHdhcmUgc2VyaWFsIG51bWJlciBvZiB0
aGUgZGV2aWNlIGJlaW5nIGFkZGVkIChtYW5kYXRvcnkpLg0KIA0KLlBBUkFNRVRFUiBoYXJkd2FyZUlkZW50aWZpZXINClRoZSBoYXJkd2FyZSBoYXNoICg0
SyBzdHJpbmcpIHRoYXQgdW5pcXVlbHkgaWRlbnRpZmllcyB0aGUgZGV2aWNlLg0KIA0KLlBBUkFNRVRFUiBncm91cFRhZw0KQW4gb3B0aW9uYWwgaWRlbnRp
ZmllciBvciB0YWcgdGhhdCBjYW4gYmUgYXNzb2NpYXRlZCB3aXRoIHRoaXMgZGV2aWNlLCB1c2VmdWwgZm9yIGdyb3VwaW5nIGRldmljZXMgdXNpbmcgRW50
cmEgZHluYW1pYyBncm91cHMuDQogDQouUEFSQU1FVEVSIGRpc3BsYXlOYW1lDQpUaGUgb3B0aW9uYWwgbmFtZSAoY29tcHV0ZXIgbmFtZSkgdG8gYmUgYXNz
aWduZWQgdG8gdGhlIGRldmljZSB3aGVuIGl0IGlzIGRlcGxveWVkIHZpYSBXaW5kb3dzIEF1dG9waWxvdC4gVGhpcyBpcyBwcmVzZW50bHkgb25seSBzdXBw
b3J0ZWQgd2l0aCBFbnRyYSBKb2luIHNjZW5hcmlvcy4gTm90ZSB0aGF0IG5hbWVzIHNob3VsZCBub3QgZXhjZWVkIDE1IGNoYXJhY3RlcnMuIEFmdGVyIHNl
dHRpbmcgdGhlIG5hbWUsIHlvdSBuZWVkIHRvIGluaXRpYXRlIGEgc3luYyAoSW52b2tlLUF1dG9waWxvdFN5bmMpIGluIG9yZGVyIHRvIHNlZSB0aGUgbmFt
ZSBpbiB0aGUgSW50dW5lIG9iamVjdC4NCiANCi5QQVJBTUVURVIgYXNzaWduZWRVc2VyDQpUaGUgb3B0aW9uYWwgdXNlciBVUE4gdG8gYmUgYXNzaWduZWQg
dG8gdGhlIGRldmljZS4gTm90ZSB0aGF0IG5vIHZhbGlkYXRpb24gaXMgZG9uZSBvbiB0aGUgVVBOIHNwZWNpZmllZC4NCiANCi5FWEFNUExFDQpBZGQgYSBu
ZXcgZGV2aWNlIHRvIFdpbmRvd3MgQXV0b3BpbG90IGZvciB0aGUgY3VycmVudCBFbnRyYSB0ZW5hbnQuDQogDQpBZGQtQXV0b3BpbG90SW1wb3J0ZWREZXZp
Y2UgLXNlcmlhbE51bWJlciAkc2VyaWFsIC1oYXJkd2FyZUlkZW50aWZpZXIgJGhhc2ggLWdyb3VwVGFnICJLaW9zayIgLWFzc2lnbmVkVXNlciAiYW5uYUBj
b250b3NvLmNvbSINCiM+DQogICAgICAgIEZ1bmN0aW9uIEFkZC1BdXRvcGlsb3RJbXBvcnRlZERldmljZSgpIHsNCiAgICAgICAgICAgIFtjbWRsZXRiaW5k
aW5nKCldDQogICAgICAgICAgICBwYXJhbQ0KICAgICAgICAgICAgKA0KICAgICAgICAgICAgICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJHRydWUpXSAk
c2VyaWFsTnVtYmVyLA0KICAgICAgICAgICAgICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJHRydWUpXSAkaGFyZHdhcmVJZGVudGlmaWVyLA0KICAgICAg
ICAgICAgICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJGZhbHNlKV0gW0FsaWFzKCJvcmRlcklkZW50aWZpZXIiKV0gJGdyb3VwVGFnID0gIiIsDQogICAg
ICAgICAgICAgICAgW1BhcmFtZXRlcihQYXJhbWV0ZXJTZXROYW1lID0gIlByb3AyIildW0FsaWFzKCJVUE4iKV0gJGFzc2lnbmVkVXNlciA9ICIiDQogICAg
ICAgICAgICApDQoNCiAgICAgICAgICAgICMgRGVmaW5pbmcgVmFyaWFibGVzDQogICAgICAgICAgICAkZ3JhcGhBcGlWZXJzaW9uID0gImJldGEiDQogICAg
ICAgICAgICAkUmVzb3VyY2UgPSAiZGV2aWNlTWFuYWdlbWVudC9pbXBvcnRlZFdpbmRvd3NBdXRvcGlsb3REZXZpY2VJZGVudGl0aWVzIg0KICAgICAgICAg
ICAgJHVyaSA9ICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vJGdyYXBoQXBpVmVyc2lvbi8kUmVzb3VyY2UiDQogICAgICAgICAgICAkanNvbiA9IEAi
DQp7DQogICAgIkBvZGF0YS50eXBlIjogIiNtaWNyb3NvZnQuZ3JhcGguaW1wb3J0ZWRXaW5kb3dzQXV0b3BpbG90RGV2aWNlSWRlbnRpdHkiLA0KICAgICJn
cm91cFRhZyI6ICIkZ3JvdXBUYWciLA0KICAgICJzZXJpYWxOdW1iZXIiOiAiJHNlcmlhbE51bWJlciIsDQogICAgInByb2R1Y3RLZXkiOiAiIiwNCiAgICAi
aGFyZHdhcmVJZGVudGlmaWVyIjogIiRoYXJkd2FyZUlkZW50aWZpZXIiLA0KICAgICJhc3NpZ25lZFVzZXJQcmluY2lwYWxOYW1lIjogIiRhc3NpZ25lZFVz
ZXIiLA0KICAgICJzdGF0ZSI6IHsNCiAgICAgICAgIkBvZGF0YS50eXBlIjogIm1pY3Jvc29mdC5ncmFwaC5pbXBvcnRlZFdpbmRvd3NBdXRvcGlsb3REZXZp
Y2VJZGVudGl0eVN0YXRlIiwNCiAgICAgICAgImRldmljZUltcG9ydFN0YXR1cyI6ICJwZW5kaW5nIiwNCiAgICAgICAgImRldmljZVJlZ2lzdHJhdGlvbklk
IjogIiIsDQogICAgICAgICJkZXZpY2VFcnJvckNvZGUiOiAwLA0KICAgICAgICAiZGV2aWNlRXJyb3JOYW1lIjogIiINCiAgICB9DQp9DQoiQA0KDQogICAg
ICAgICAgICBXcml0ZS1WZXJib3NlICJQT1NUICR1cmlgbiRqc29uIg0KDQogICAgICAgICAgICB0cnkgew0KICAgICAgICAgICAgICAgIEludm9rZS1NZ0dy
YXBoUmVxdWVzdCAtVXJpICR1cmkgLU1ldGhvZCBQb3N0IC1Cb2R5ICRqc29uIC1Db250ZW50VHlwZSAiYXBwbGljYXRpb24vanNvbiINCiAgICAgICAgICAg
IH0NCiAgICAgICAgICAgIGNhdGNoIHsNCiAgICAgICAgICAgICAgICBXcml0ZS1FcnJvciAkXy5FeGNlcHRpb24gDQogICAgICAgICAgICAgICAgYnJlYWsN
CiAgICAgICAgICAgIH0NCiAgICANCiAgICAgICAgfQ0KDQogICAgDQogICAgICAgIEZ1bmN0aW9uIFJlbW92ZS1BdXRvcGlsb3RJbXBvcnRlZERldmljZSgp
IHsNCiAgICAgICAgICAgIDwjDQouU1lOT1BTSVMNClJlbW92ZXMgdGhlIHN0YXR1cyBpbmZvcm1hdGlvbiBmb3IgYSBkZXZpY2UgYmVpbmcgaW1wb3J0ZWQg
aW50byBXaW5kb3dzIEF1dG9waWxvdC4NCiANCi5ERVNDUklQVElPTg0KVGhlIFJlbW92ZS1BdXRvcGlsb3RJbXBvcnRlZERldmljZSBjbWRsZXQgY2xlYW5z
IHVwIHRoZSBzdGF0dXMgaW5mb3JtYXRpb24gYWJvdXQgYSBuZXcgZGV2aWNlIGJlaW5nIGltcG9ydGVkIGludG8gV2luZG93cyBBdXRvcGlsb3QuIFRoaXMg
c2hvdWxkIGJlIGRvbmUgcmVnYXJkbGVzcyBvZiB3aGV0aGVyIHRoZSBpbXBvcnQgd2FzIHN1Y2Nlc3NmdWwgb3Igbm90Lg0KIA0KLlBBUkFNRVRFUiBpZA0K
VGhlIElEIChHVUlEKSBvZiB0aGUgaW1wb3J0ZWQgZGV2aWNlIHN0YXR1cyBpbmZvcm1hdGlvbiB0byBiZSByZW1vdmVkIChtYW5kYXRvcnkpLg0KIA0KLkVY
QU1QTEUNClJlbW92ZSB0aGUgc3RhdHVzIGluZm9ybWF0aW9uIGZvciBhIHNwZWNpZmllZCBkZXZpY2UuDQogDQpSZW1vdmUtQXV0b3BpbG90SW1wb3J0ZWRE
ZXZpY2UgLWlkICRpZA0KIz4NCiAgICAgICAgICAgIFtjbWRsZXRiaW5kaW5nKCldDQogICAgICAgICAgICBwYXJhbQ0KICAgICAgICAgICAgKA0KICAgICAg
ICAgICAgICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJHRydWUsIFZhbHVlRnJvbVBpcGVsaW5lQnlQcm9wZXJ0eU5hbWUgPSAkVHJ1ZSldICRpZA0KICAg
ICAgICAgICAgKQ0KDQogICAgICAgICAgICBQcm9jZXNzIHsNCg0KICAgICAgICAgICAgICAgICMgRGVmaW5pbmcgVmFyaWFibGVzDQogICAgICAgICAgICAg
ICAgJGdyYXBoQXBpVmVyc2lvbiA9ICJiZXRhIg0KICAgICAgICAgICAgICAgICRSZXNvdXJjZSA9ICJkZXZpY2VNYW5hZ2VtZW50L2ltcG9ydGVkV2luZG93
c0F1dG9waWxvdERldmljZUlkZW50aXRpZXMiICAgIA0KICAgICAgICAgICAgICAgICR1cmkgPSAiaHR0cHM6Ly9ncmFwaC5taWNyb3NvZnQuY29tLyRncmFw
aEFwaVZlcnNpb24vJFJlc291cmNlLyRpZCINCg0KICAgICAgICAgICAgICAgIHRyeSB7DQogICAgICAgICAgICAgICAgICAgIFdyaXRlLVZlcmJvc2UgIkRF
TEVURSAkdXJpIg0KICAgICAgICAgICAgICAgICAgICBJbnZva2UtTWdHcmFwaFJlcXVlc3QgLVVyaSAkdXJpIC1NZXRob2QgREVMRVRFDQogICAgICAgICAg
ICAgICAgfQ0KICAgICAgICAgICAgICAgIGNhdGNoIHsNCiAgICAgICAgICAgICAgICAgICAgV3JpdGUtRXJyb3IgJF8uRXhjZXB0aW9uIA0KICAgICAgICAg
ICAgICAgICAgICBicmVhaw0KICAgICAgICAgICAgICAgIH0NCg0KICAgICAgICAgICAgfQ0KICAgICAgICANCiAgICAgICAgfQ0KDQoNCiAgICAgICAgRnVu
Y3Rpb24gR2V0LUF1dG9waWxvdFByb2ZpbGUoKSB7DQogICAgICAgICAgICA8Iw0KLlNZTk9QU0lTDQpHZXRzIFdpbmRvd3MgQXV0b3BpbG90IHByb2ZpbGUg
ZGV0YWlscy4NCiANCi5ERVNDUklQVElPTg0KVGhlIEdldC1BdXRvcGlsb3RQcm9maWxlIGNtZGxldCByZXR1cm5zIGVpdGhlciBhIGxpc3Qgb2YgYWxsIFdp
bmRvd3MgQXV0b3BpbG90IHByb2ZpbGVzIGZvciB0aGUgY3VycmVudCBFbnRyYSB0ZW5hbnQsIG9yIGluZm9ybWF0aW9uIGZvciB0aGUgc3BlY2lmaWMgcHJv
ZmlsZSBzcGVjaWZpZWQgYnkgaXRzIElELg0KIA0KLlBBUkFNRVRFUiBpZA0KT3B0aW9uYWxseSwgdGhlIElEIChHVUlEKSBvZiB0aGUgcHJvZmlsZSB0byBi
ZSByZXRyaWV2ZWQuDQogDQouRVhBTVBMRQ0KR2V0IGEgbGlzdCBvZiBhbGwgV2luZG93cyBBdXRvcGlsb3QgcHJvZmlsZXMuDQogDQpHZXQtQXV0b3BpbG90
UHJvZmlsZQ0KIz4NCiAgICAgICAgICAgIFtjbWRsZXRiaW5kaW5nKCldDQogICAgICAgICAgICBwYXJhbQ0KICAgICAgICAgICAgKA0KICAgICAgICAgICAg
ICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJGZhbHNlKV0gJGlkDQogICAgICAgICAgICApDQoNCiAgICAgICAgICAgICMgRGVmaW5pbmcgVmFyaWFibGVz
DQogICAgICAgICAgICAkZ3JhcGhBcGlWZXJzaW9uID0gImJldGEiDQogICAgICAgICAgICAkUmVzb3VyY2UgPSAiZGV2aWNlTWFuYWdlbWVudC93aW5kb3dz
QXV0b3BpbG90RGVwbG95bWVudFByb2ZpbGVzIg0KDQogICAgICAgICAgICBpZiAoJGlkKSB7DQogICAgICAgICAgICAgICAgJHVyaSA9ICJodHRwczovL2dy
YXBoLm1pY3Jvc29mdC5jb20vJGdyYXBoQXBpVmVyc2lvbi8kUmVzb3VyY2UvJGlkIg0KICAgICAgICAgICAgfQ0KICAgICAgICAgICAgZWxzZSB7DQogICAg
ICAgICAgICAgICAgJHVyaSA9ICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vJGdyYXBoQXBpVmVyc2lvbi8kUmVzb3VyY2UiDQogICAgICAgICAgICB9
DQoNCiAgICAgICAgICAgIFdyaXRlLVZlcmJvc2UgIkdFVCAkdXJpIg0KDQogICAgICAgICAgICB0cnkgew0KICAgICAgICAgICAgICAgICRyZXNwb25zZSA9
IEludm9rZS1NZ0dyYXBoUmVxdWVzdCAtVXJpICR1cmkgLU1ldGhvZCBHZXQgLU91dHB1dFR5cGUgUFNPYmplY3QNCiAgICAgICAgICAgICAgICBpZiAoJGlk
KSB7DQogICAgICAgICAgICAgICAgICAgICRyZXNwb25zZQ0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBlbHNlIHsNCiAgICAgICAgICAg
ICAgICAgICAgJGRldmljZXMgPSAkcmVzcG9uc2UudmFsdWUNCiAgICANCiAgICAgICAgICAgICAgICAgICAgJGRldmljZXNOZXh0TGluayA9ICRyZXNwb25z
ZS4iQG9kYXRhLm5leHRMaW5rIg0KICAgIA0KICAgICAgICAgICAgICAgICAgICB3aGlsZSAoJG51bGwgLW5lICRkZXZpY2VzTmV4dExpbmspIHsNCiAgICAg
ICAgICAgICAgICAgICAgICAgICRkZXZpY2VzUmVzcG9uc2UgPSAoSW52b2tlLU1nR3JhcGhSZXF1ZXN0IC1VcmkgJGRldmljZXNOZXh0TGluayAtTWV0aG9k
IEdldCAtT3V0cHV0VHlwZSBQU09iamVjdCkNCiAgICAgICAgICAgICAgICAgICAgICAgICRkZXZpY2VzTmV4dExpbmsgPSAkZGV2aWNlc1Jlc3BvbnNlLiJA
b2RhdGEubmV4dExpbmsiDQogICAgICAgICAgICAgICAgICAgICAgICAkZGV2aWNlcyArPSAkZGV2aWNlc1Jlc3BvbnNlLnZhbHVlDQogICAgICAgICAgICAg
ICAgICAgIH0NCiAgICANCiAgICAgICAgICAgICAgICAgICAgJGRldmljZXMNCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICB9DQogICAgICAgICAg
ICBjYXRjaCB7DQogICAgICAgICAgICAgICAgV3JpdGUtRXJyb3IgJF8uRXhjZXB0aW9uIA0KICAgICAgICAgICAgICAgIGJyZWFrDQogICAgICAgICAgICB9
DQoNCiAgICAgICAgfQ0KDQoNCiAgICAgICAgRnVuY3Rpb24gR2V0LUF1dG9waWxvdFByb2ZpbGVBc3NpZ25lZERldmljZSgpIHsNCiAgICAgICAgICAgIDwj
DQouU1lOT1BTSVMNCkdldHMgdGhlIGxpc3Qgb2YgZGV2aWNlcyB0aGF0IGFyZSBhc3NpZ25lZCB0byB0aGUgc3BlY2lmaWVkIFdpbmRvd3MgQXV0b3BpbG90
IHByb2ZpbGUuDQogDQouREVTQ1JJUFRJT04NClRoZSBHZXQtQXV0b3BpbG90UHJvZmlsZUFzc2lnbmVkRGV2aWNlIGNtZGxldCByZXR1cm5zIHRoZSBsaXN0
IG9mIEF1dG9waWxvdCBkZXZpY2VzIHRoYXQgaGF2ZSBiZWVuIGFzc2lnbmVkIHRoZSBzcGVjaWZpZWQgV2luZG93cyBBdXRvcGlsb3QgcHJvZmlsZS4NCiAN
Ci5QQVJBTUVURVIgaWQNClRoZSBJRCAoR1VJRCkgb2YgdGhlIHByb2ZpbGUgdG8gYmUgcmV0cmlldmVkLg0KIA0KLkVYQU1QTEUNCkdldCBhIGxpc3Qgb2Yg
YWxsIFdpbmRvd3MgQXV0b3BpbG90IHByb2ZpbGVzLg0KIA0KR2V0LUF1dG9waWxvdFByb2ZpbGVBc3NpZ25lZERldmljZXMgLWlkICRpZA0KIz4NCiAgICAg
ICAgICAgIFtjbWRsZXRiaW5kaW5nKCldDQogICAgICAgICAgICBwYXJhbQ0KICAgICAgICAgICAgKA0KICAgICAgICAgICAgICAgIFtQYXJhbWV0ZXIoTWFu
ZGF0b3J5ID0gJGZhbHNlLCBWYWx1ZUZyb21QaXBlbGluZUJ5UHJvcGVydHlOYW1lID0gJFRydWUpXSAkaWQNCiAgICAgICAgICAgICkNCg0KICAgICAgICAg
ICAgUHJvY2VzcyB7DQoNCiAgICAgICAgICAgICAgICAjIERlZmluaW5nIFZhcmlhYmxlcw0KICAgICAgICAgICAgICAgICRncmFwaEFwaVZlcnNpb24gPSAi
YmV0YSINCiAgICAgICAgICAgICAgICAkUmVzb3VyY2UgPSAiZGV2aWNlTWFuYWdlbWVudC93aW5kb3dzQXV0b3BpbG90RGVwbG95bWVudFByb2ZpbGVzIg0K
ICAgICAgICAgICAgICAgICR1cmkgPSAiaHR0cHM6Ly9ncmFwaC5taWNyb3NvZnQuY29tLyRncmFwaEFwaVZlcnNpb24vJFJlc291cmNlLyRpZC9hc3NpZ25l
ZERldmljZXMiDQoNCiAgICAgICAgICAgICAgICBXcml0ZS1WZXJib3NlICJHRVQgJHVyaSINCg0KICAgICAgICAgICAgICAgIHRyeSB7DQogICAgICAgICAg
ICAgICAgICAgICRyZXNwb25zZSA9IEludm9rZS1NZ0dyYXBoUmVxdWVzdCAtVXJpICR1cmkgLU1ldGhvZCBHZXQNCiAgICAgICAgICAgICAgICAgICAgJHJl
c3BvbnNlLlZhbHVlDQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgIGNhdGNoIHsNCiAgICAgICAgICAgICAgICAgICAgV3JpdGUtRXJyb3Ig
JF8uRXhjZXB0aW9uIA0KICAgICAgICAgICAgICAgICAgICBicmVhaw0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgIH0NCiAgICAgICAgfQ0KDQoN
Cg0KICAgICAgICBGdW5jdGlvbiBDb252ZXJ0VG8tQXV0b3BpbG90Q29uZmlndXJhdGlvbkpTT04oKSB7DQogICAgICAgICAgICA8Iw0KLlNZTk9QU0lTDQpD
b252ZXJ0cyB0aGUgc3BlY2lmaWVkIFdpbmRvd3MgQXV0b3BpbG90IHByb2ZpbGUgaW50byBhIEpTT04gZm9ybWF0Lg0KIA0KLkRFU0NSSVBUSU9ODQpUaGUg
Q29udmVydFRvLUF1dG9waWxvdENvbmZpZ3VyYXRpb25KU09OIGNtZGxldCBjb252ZXJ0cyB0aGUgc3BlY2lmaWVkIFdpbmRvd3MgQXV0b3BpbG90IHByb2Zp
bGUsIGFzIHJlcHJlc2VudGVkIGJ5IGEgTWljcm9zb2Z0IEdyYXBoIEFQSSBvYmplY3QsIGludG8gYSBKU09OIGZvcm1hdC4NCiANCi5QQVJBTUVURVIgcHJv
ZmlsZQ0KQSBXaW5kb3dzIEF1dG9waWxvdCBwcm9maWxlIG9iamVjdCwgdHlwaWNhbGx5IHJldHVybmVkIGJ5IEdldC1BdXRvcGlsb3RQcm9maWxlDQogDQou
RVhBTVBMRQ0KR2V0IHRoZSBKU09OIHJlcHJlc2VudGF0aW9uIG9mIGVhY2ggV2luZG93cyBBdXRvcGlsb3QgcHJvZmlsZSBpbiB0aGUgY3VycmVudCBFbnRy
YSB0ZW5hbnQuDQogDQpHZXQtQXV0b3BpbG90UHJvZmlsZSB8IENvbnZlcnRUby1BdXRvcGlsb3RDb25maWd1cmF0aW9uSlNPTg0KIz4NCiAgICAgICAgICAg
IFtjbWRsZXRiaW5kaW5nKCldDQogICAgICAgICAgICBwYXJhbQ0KICAgICAgICAgICAgKA0KICAgICAgICAgICAgICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5
ID0gJHRydWUsIFZhbHVlRnJvbVBpcGVsaW5lID0gJFRydWUpXQ0KICAgICAgICAgICAgICAgIFtPYmplY3RdICRwcm9maWxlDQogICAgICAgICAgICApDQoN
CiAgICAgICAgICAgIEJlZ2luIHsNCg0KICAgICAgICAgICAgICAgICMgU2V0IHRoZSBvcmctcmVsYXRlZCBpbmZvDQogICAgICAgICAgICAgICAgJHNjcmlw
dDpUZW5hbnRPcmcgPSBHZXQtT3JnYW5pemF0aW9uDQogICAgICAgICAgICAgICAgZm9yZWFjaCAoJGRvbWFpbiBpbiAkc2NyaXB0OlRlbmFudE9yZy5WZXJp
ZmllZERvbWFpbnMpIHsNCiAgICAgICAgICAgICAgICAgICAgaWYgKCRkb21haW4uaXNEZWZhdWx0KSB7DQogICAgICAgICAgICAgICAgICAgICAgICAkc2Ny
aXB0OlRlbmFudERvbWFpbiA9ICRkb21haW4ubmFtZQ0KICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgfQ0K
DQogICAgICAgICAgICBQcm9jZXNzIHsNCg0KICAgICAgICAgICAgICAgICRvb2JlU2V0dGluZ3MgPSAkcHJvZmlsZS5vdXRPZkJveEV4cGVyaWVuY2VTZXR0
aW5ncw0KDQogICAgICAgICAgICAgICAgIyBCdWlsZCB1cCBwcm9wZXJ0aWVzDQogICAgICAgICAgICAgICAgJGpzb24gPSBAe30NCiAgICAgICAgICAgICAg
ICAkanNvbi5BZGQoIkNvbW1lbnRfRmlsZSIsICJQcm9maWxlICQoJF8uZGlzcGxheU5hbWUpIikNCiAgICAgICAgICAgICAgICAkanNvbi5BZGQoIlZlcnNp
b24iLCAyMDQ5KQ0KICAgICAgICAgICAgICAgICRqc29uLkFkZCgiWnRkQ29ycmVsYXRpb25JZCIsICRfLmlkKQ0KICAgICAgICAgICAgICAgIGlmICgkcHJv
ZmlsZS4iQG9kYXRhLnR5cGUiIC1lcSAiI21pY3Jvc29mdC5ncmFwaC5hY3RpdmVEaXJlY3RvcnlXaW5kb3dzQXV0b3BpbG90RGVwbG95bWVudFByb2ZpbGUi
KSB7DQogICAgICAgICAgICAgICAgICAgICRqc29uLkFkZCgiQ2xvdWRBc3NpZ25lZERvbWFpbkpvaW5NZXRob2QiLCAxKQ0KICAgICAgICAgICAgICAgIH0N
CiAgICAgICAgICAgICAgICBlbHNlIHsNCiAgICAgICAgICAgICAgICAgICAgJGpzb24uQWRkKCJDbG91ZEFzc2lnbmVkRG9tYWluSm9pbk1ldGhvZCIsIDAp
DQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgIGlmICgkcHJvZmlsZS5kZXZpY2VOYW1lVGVtcGxhdGUpIHsNCiAgICAgICAgICAgICAgICAg
ICAgJGpzb24uQWRkKCJDbG91ZEFzc2lnbmVkRGV2aWNlTmFtZSIsICRfLmRldmljZU5hbWVUZW1wbGF0ZSkNCiAgICAgICAgICAgICAgICB9DQoNCiAgICAg
ICAgICAgICAgICAjIEZpZ3VyZSBvdXQgY29uZmlnIHZhbHVlDQogICAgICAgICAgICAgICAgJG9vYmVDb25maWcgPSA4ICsgMjU2DQogICAgICAgICAgICAg
ICAgaWYgKCRvb2JlU2V0dGluZ3MudXNlclR5cGUgLWVxICdzdGFuZGFyZCcpIHsNCiAgICAgICAgICAgICAgICAgICAgJG9vYmVDb25maWcgKz0gMg0KICAg
ICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBpZiAoJG9vYmVTZXR0aW5ncy5oaWRlUHJpdmFjeVNldHRpbmdzIC1lcSAkdHJ1ZSkgew0KICAgICAg
ICAgICAgICAgICAgICAkb29iZUNvbmZpZyArPSA0DQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgIGlmICgkb29iZVNldHRpbmdzLmhpZGVF
VUxBIC1lcSAkdHJ1ZSkgew0KICAgICAgICAgICAgICAgICAgICAkb29iZUNvbmZpZyArPSAxNg0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAg
ICBpZiAoJG9vYmVTZXR0aW5ncy5za2lwS2V5Ym9hcmRTZWxlY3Rpb25QYWdlIC1lcSAkdHJ1ZSkgew0KICAgICAgICAgICAgICAgICAgICAkb29iZUNvbmZp
ZyArPSAxMDI0DQogICAgICAgICAgICAgICAgICAgIGlmICgkXy5sYW5ndWFnZSkgew0KICAgICAgICAgICAgICAgICAgICAgICAgJGpzb24uQWRkKCJDbG91
ZEFzc2lnbmVkTGFuZ3VhZ2UiLCAkXy5sYW5ndWFnZSkNCiAgICAgICAgICAgICAgICAgICAgICAgICMgVXNlIHRoZSBzYW1lIHZhbHVlIGZvciByZWdpb24g
c28gdGhhdCBzY3JlZW4gaXMgc2tpcHBlZCB0b28NCiAgICAgICAgICAgICAgICAgICAgICAgICRqc29uLkFkZCgiQ2xvdWRBc3NpZ25lZFJlZ2lvbiIsICRf
Lmxhbmd1YWdlKQ0KICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgIGlmICgkb29iZVNldHRpbmdzLmRl
dmljZVVzYWdlVHlwZSAtZXEgJ3NoYXJlZCcpIHsNCiAgICAgICAgICAgICAgICAgICAgJG9vYmVDb25maWcgKz0gMzIgKyA2NA0KICAgICAgICAgICAgICAg
IH0NCiAgICAgICAgICAgICAgICAkanNvbi5BZGQoIkNsb3VkQXNzaWduZWRPb2JlQ29uZmlnIiwgJG9vYmVDb25maWcpDQoNCiAgICAgICAgICAgICAgICAj
IFNldCB0aGUgZm9yY2VkIGVucm9sbG1lbnQgc2V0dGluZw0KICAgICAgICAgICAgICAgIGlmICgkb29iZVNldHRpbmdzLmhpZGVFc2NhcGVMaW5rIC1lcSAk
dHJ1ZSkgew0KICAgICAgICAgICAgICAgICAgICAkanNvbi5BZGQoIkNsb3VkQXNzaWduZWRGb3JjZWRFbnJvbGxtZW50IiwgMSkNCiAgICAgICAgICAgICAg
ICB9DQogICAgICAgICAgICAgICAgZWxzZSB7DQogICAgICAgICAgICAgICAgICAgICRqc29uLkFkZCgiQ2xvdWRBc3NpZ25lZEZvcmNlZEVucm9sbG1lbnQi
LCAwKQ0KICAgICAgICAgICAgICAgIH0NCg0KICAgICAgICAgICAgICAgICRqc29uLkFkZCgiQ2xvdWRBc3NpZ25lZFRlbmFudElkIiwgJHNjcmlwdDpUZW5h
bnRPcmcuaWQpDQogICAgICAgICAgICAgICAgJGpzb24uQWRkKCJDbG91ZEFzc2lnbmVkVGVuYW50RG9tYWluIiwgJHNjcmlwdDpUZW5hbnREb21haW4pDQog
ICAgICAgICAgICAgICAgJGVtYmVkZGVkID0gQHt9DQogICAgICAgICAgICAgICAgJGVtYmVkZGVkLkFkZCgiQ2xvdWRBc3NpZ25lZFRlbmFudERvbWFpbiIs
ICRzY3JpcHQ6VGVuYW50RG9tYWluKQ0KICAgICAgICAgICAgICAgICRlbWJlZGRlZC5BZGQoIkNsb3VkQXNzaWduZWRUZW5hbnRVcG4iLCAiIikNCiAgICAg
ICAgICAgICAgICBpZiAoJG9vYmVTZXR0aW5ncy5oaWRlRXNjYXBlTGluayAtZXEgJHRydWUpIHsNCiAgICAgICAgICAgICAgICAgICAgJGVtYmVkZGVkLkFk
ZCgiRm9yY2VkRW5yb2xsbWVudCIsIDEpDQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgIGVsc2Ugew0KICAgICAgICAgICAgICAgICAgICAk
ZW1iZWRkZWQuQWRkKCJGb3JjZWRFbnJvbGxtZW50IiwgMCkNCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgJHp0YyA9IEB7fQ0KICAgICAg
ICAgICAgICAgICR6dGMuQWRkKCJaZXJvVG91Y2hDb25maWciLCAkZW1iZWRkZWQpDQogICAgICAgICAgICAgICAgJGpzb24uQWRkKCJDbG91ZEFzc2lnbmVk
QWFkU2VydmVyRGF0YSIsIChDb252ZXJ0VG8tSnNvbiAkenRjIC1Db21wcmVzcykpDQoNCiAgICAgICAgICAgICAgICAjIFNraXAgY29ubmVjdGl2aXR5IGNo
ZWNrDQogICAgICAgICAgICAgICAgaWYgKCRwcm9maWxlLmh5YnJpZEF6dXJlQURKb2luU2tpcENvbm5lY3Rpdml0eUNoZWNrIC1lcSAkdHJ1ZSkgew0KICAg
ICAgICAgICAgICAgICAgICAkanNvbi5BZGQoIkh5YnJpZEpvaW5Ta2lwRENDb25uZWN0aXZpdHlDaGVjayIsIDEpDQogICAgICAgICAgICAgICAgfQ0KDQog
ICAgICAgICAgICAgICAgIyBIYXJkLWNvZGUgcHJvcGVydGllcyBub3QgcmVwcmVzZW50ZWQgaW4gSW50dW5lDQogICAgICAgICAgICAgICAgJGpzb24uQWRk
KCJDbG91ZEFzc2lnbmVkQXV0b3BpbG90VXBkYXRlRGlzYWJsZWQiLCAxKQ0KICAgICAgICAgICAgICAgICRqc29uLkFkZCgiQ2xvdWRBc3NpZ25lZEF1dG9w
aWxvdFVwZGF0ZVRpbWVvdXQiLCAxODAwMDAwKQ0KDQogICAgICAgICAgICAgICAgIyBSZXR1cm4gdGhlIEpTT04NCiAgICAgICAgICAgICAgICBDb252ZXJ0
VG8tSnNvbiAkanNvbg0KICAgICAgICAgICAgfQ0KDQogICAgICAgIH0NCg0KDQogICAgICAgIEZ1bmN0aW9uIFNldC1BdXRvcGlsb3RQcm9maWxlKCkgew0K
ICAgICAgICAgICAgPCMNCi5TWU5PUFNJUw0KU2V0cyBXaW5kb3dzIEF1dG9waWxvdCBwcm9maWxlIHByb3BlcnRpZXMgb24gYW4gZXhpc3RpbmcgQXV0b3Bp
bG90IHByb2ZpbGUuDQogDQouREVTQ1JJUFRJT04NClRoZSBTZXQtQXV0b3BpbG90UHJvZmlsZSBjbWRsZXQgc2V0cyBwcm9wZXJ0aWVzIG9uIGFuIGV4aXN0
aW5nIEF1dG9waWxvdCBwcm9maWxlLg0KIA0KLlBBUkFNRVRFUiBpZA0KVGhlIEdVSUQgb2YgdGhlIHByb2ZpbGUgdG8gYmUgdXBkYXRlZC4NCiANCi5QQVJB
TUVURVIgZGlzcGxheU5hbWUNClRoZSBuYW1lIG9mIHRoZSBXaW5kb3dzIEF1dG9waWxvdCBwcm9maWxlIHRvIGNyZWF0ZS4gKFRoaXMgdmFsdWUgY2Fubm90
IGNvbnRhaW4gc3BhY2VzLikNCiANCi5QQVJBTUVURVIgZGVzY3JpcHRpb24NClRoZSBkZXNjcmlwdGlvbiB0byBiZSBjb25maWd1cmVkIGluIHRoZSBwcm9m
aWxlLiAoVGhpcyB2YWx1ZSBjYW5ub3QgY29udGFpbiBkYXNoZXMuKQ0KIA0KLlBBUkFNRVRFUiBDb252ZXJ0RGV2aWNlVG9BdXRvcGlsb3QNCkNvbmZpZ3Vy
ZSB0aGUgdmFsdWUgIkNvbnZlcnQgYWxsIHRhcmdldGVkIGRldmljZXMgdG8gQXV0b3BpbG90Ig0KIA0KLlBBUkFNRVRFUiBBbGxFbmFibGVkDQpFbmFibGUg
ZXZlcnl0aGluZyB0aGF0IGNhbiBiZSBlbmFibGVkDQogDQouUEFSQU1FVEVSIEFsbERpc2FibGVkDQpEaXNhYmxlIGV2ZXJ5dGhpbmcgdGhhdCBjYW4gYmUg
ZGlzYWJsZWQNCiANCi5QQVJBTUVURVIgT09CRV9IaWRlRVVMQQ0KQ29uZmlndXJlIHRoZSBPT0JFIG9wdGlvbiB0byBoaWRlIG9yIG5vdCB0aGUgRVVMQQ0K
IA0KLlBBUkFNRVRFUiBPT0JFX0VuYWJsZVdoaXRlR2xvdmUNCkNvbmZpZ3VyZSB0aGUgT09CRSBvcHRpb24gdG8gYWxsb3cgb3Igbm90IFdoaXRlIEdsb3Zl
IE9PQkUNCiANCi5QQVJBTUVURVIgT09CRV9IaWRlUHJpdmFjeVNldHRpbmdzDQpDb25maWd1cmUgdGhlIE9PQkUgb3B0aW9uIHRvIGhpZGUgb3Igbm90IHRo
ZSBwcml2YWN5IHNldHRpbmdzDQogDQouUEFSQU1FVEVSIE9PQkVfSGlkZUNoYW5nZUFjY291bnRPcHRzDQpDb25maWd1cmUgdGhlIE9PQkUgb3B0aW9uIHRv
IGhpZGUgb3Igbm90IHRoZSBjaGFuZ2UgYWNjb3VudCBvcHRpb25zDQogDQouUEFSQU1FVEVSIE9PQkVfVXNlclR5cGVBZG1pbg0KQ29uZmlndXJlIHRoZSB1
c2VyIGFjY291bnQgdHlwZSBhcyBhZG1pbmlzdHJhdG9yLg0KIA0KLlBBUkFNRVRFUiBPT0JFX05hbWVUZW1wbGF0ZQ0KQ29uZmlndXJlIHRoZSBPT0JFIG9w
dGlvbiB0byBhcHBseSBhIGRldmljZSBuYW1lIHRlbXBsYXRlDQogDQouUEFSQU1FVEVSIE9PQkVfbGFuZ3VhZ2UNClRoZSBsYW5ndWFnZSBpZGVudGlmaWVy
IChlLmcuICJlbi11cyIpIHRvIGJlIGNvbmZpZ3VyZWQgaW4gdGhlIHByb2ZpbGUNCiANCi5QQVJBTUVURVIgT09CRV9Ta2lwS2V5Ym9hcmQNCkNvbmZpZ3Vy
ZSB0aGUgT09CRSBvcHRpb24gdG8gc2tpcCBvciBub3QgdGhlIGtleWJvYXJkIHNlbGVjdGlvbiBwYWdlDQogDQouUEFSQU1FVEVSIE9PQkVfSGlkZUNoYW5n
ZUFjY291bnRPcHRzDQpDb25maWd1cmUgdGhlIE9PQkUgb3B0aW9uIHRvIGhpZGUgb3Igbm90IHRoZSBjaGFuZ2UgYWNjb3VudCBvcHRpb25zDQogDQouUEFS
QU1FVEVSIE9PQkVfU2tpcENvbm5lY3Rpdml0eUNoZWNrDQpTcGVjaWZ5IHdoZXRoZXIgdG8gc2tpcCBBY3RpdmUgRGlyZWN0b3J5IGNvbm5lY3Rpdml0eSBj
aGVjayAoVXNlckRyaXZlbkFBRCBvbmx5KQ0KIA0KLkVYQU1QTEUNClVwZGF0ZSBhbiBleGlzdGluZyBBdXRvcGlsb3QgcHJvZmlsZSB0byBzcGVjaWZ5IGEg
bGFuZ3VhZ2U6DQogDQpTZXQtQXV0b3BpbG90UHJvZmlsZSAtSUQgPGd1aWQ+IC1MYW5ndWFnZSAiZW4tdXMiDQogDQouRVhBTVBMRQ0KVXBkYXRlIGFuIGV4
aXN0aW5nIEF1dG9waWxvdCBwcm9maWxlIHRvIHNldCBtdWx0aXBsZSBwcm9wZXJ0aWVzOg0KIA0KU2V0LUF1dG9waWxvdFByb2ZpbGUgLUlEIDxndWlkPiAt
TGFuZ3VhZ2UgImVuLXVzIiAtZGlzcGxheW5hbWUgIk15IHRlc3RpbmcgcHJvZmlsZSIgLURlc2NyaXB0aW9uICJEZXNjcmlwdGlvbiBvZiBteSBwcm9maWxl
IiAtT09CRV9IaWRlRVVMQSAkVHJ1ZSAtT09CRV9oaWRlUHJpdmFjeVNldHRpbmdzICRUcnVlDQojPg0KICAgICAgICAgICAgW2NtZGxldGJpbmRpbmcoKV0N
CiAgICAgICAgICAgIHBhcmFtDQogICAgICAgICAgICAoDQogICAgICAgICAgICAgICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkdHJ1ZSwgVmFsdWVGcm9t
UGlwZWxpbmVCeVByb3BlcnR5TmFtZSA9ICRUcnVlKV0gJGlkLA0KICAgICAgICAgICAgICAgIFtQYXJhbWV0ZXIoUGFyYW1ldGVyU2V0TmFtZSA9ICdub3RB
bGwnKV1bc3RyaW5nXSAkZGlzcGxheU5hbWUsDQogICAgICAgICAgICAgICAgW1BhcmFtZXRlcihQYXJhbWV0ZXJTZXROYW1lID0gJ25vdEFsbCcpXVtzdHJp
bmddICRkZXNjcmlwdGlvbiwNCiAgICAgICAgICAgICAgICBbUGFyYW1ldGVyKFBhcmFtZXRlclNldE5hbWUgPSAnbm90QWxsJyldW1N3aXRjaF0gJENvbnZl
cnREZXZpY2VUb0F1dG9waWxvdCwNCiAgICAgICAgICAgICAgICBbUGFyYW1ldGVyKFBhcmFtZXRlclNldE5hbWUgPSAnbm90QWxsJyldW3N0cmluZ10gJE9P
QkVfbGFuZ3VhZ2UsDQogICAgICAgICAgICAgICAgW1BhcmFtZXRlcihQYXJhbWV0ZXJTZXROYW1lID0gJ25vdEFsbCcpXVtTd2l0Y2hdICRPT0JFX3NraXBL
ZXlib2FyZCwNCiAgICAgICAgICAgICAgICBbUGFyYW1ldGVyKFBhcmFtZXRlclNldE5hbWUgPSAnbm90QWxsJyldW3N0cmluZ10gJE9PQkVfTmFtZVRlbXBs
YXRlLA0KICAgICAgICAgICAgICAgIFtQYXJhbWV0ZXIoUGFyYW1ldGVyU2V0TmFtZSA9ICdub3RBbGwnKV1bU3dpdGNoXSAkT09CRV9FbmFibGVXaGl0ZUds
b3ZlLA0KICAgICAgICAgICAgICAgIFtQYXJhbWV0ZXIoUGFyYW1ldGVyU2V0TmFtZSA9ICdub3RBbGwnKV1bU3dpdGNoXSAkT09CRV9Vc2VyVHlwZUFkbWlu
LA0KICAgICAgICAgICAgICAgIFtQYXJhbWV0ZXIoUGFyYW1ldGVyU2V0TmFtZSA9ICdBbGxFbmFibGVkJywgTWFuZGF0b3J5ID0gJHRydWUpXVtTd2l0Y2hd
ICRBbGxFbmFibGVkLCANCiAgICAgICAgICAgICAgICBbUGFyYW1ldGVyKFBhcmFtZXRlclNldE5hbWUgPSAnQWxsRGlzYWJsZWQnLCBNYW5kYXRvcnkgPSAk
dHJ1ZSldW1N3aXRjaF0gJEFsbERpc2FibGVkLCANCiAgICAgICAgICAgICAgICBbUGFyYW1ldGVyKFBhcmFtZXRlclNldE5hbWUgPSAnbm90QWxsJyldW1N3
aXRjaF0gJE9PQkVfSGlkZUVVTEEsDQogICAgICAgICAgICAgICAgW1BhcmFtZXRlcihQYXJhbWV0ZXJTZXROYW1lID0gJ25vdEFsbCcpXVtTd2l0Y2hdICRP
T0JFX2hpZGVQcml2YWN5U2V0dGluZ3MsDQogICAgICAgICAgICAgICAgW1BhcmFtZXRlcihQYXJhbWV0ZXJTZXROYW1lID0gJ25vdEFsbCcpXVtTd2l0Y2hd
ICRPT0JFX0hpZGVDaGFuZ2VBY2NvdW50T3B0cywNCiAgICAgICAgICAgICAgICBbUGFyYW1ldGVyKFBhcmFtZXRlclNldE5hbWUgPSAnbm90QWxsJyldW1N3
aXRjaF0gJE9PQkVfU2tpcENvbm5lY3Rpdml0eUNoZWNrDQogICAgICAgICAgICApDQoNCiAgICAgICAgICAgICMgR2V0IHRoZSBjdXJyZW50IHZhbHVlcw0K
ICAgICAgICAgICAgJGN1cnJlbnQgPSBHZXQtQXV0b3BpbG90UHJvZmlsZSAtaWQgJGlkDQoNCiAgICAgICAgICAgICMgSWYgdGhpcyBpcyBhIEh5YnJpZCBB
QURKIHByb2ZpbGUsIG1ha2Ugc3VyZSBpdCBoYXMgdGhlIG5lZWRlZCBwcm9wZXJ0eQ0KICAgICAgICAgICAgaWYgKCRjdXJyZW50LidAb2RhdGEudHlwZScg
LWVxICIjbWljcm9zb2Z0LmdyYXBoLmF6dXJlQURXaW5kb3dzQXV0b3BpbG90RGVwbG95bWVudFByb2ZpbGUiKSB7DQogICAgICAgICAgICAgICAgaWYgKC1u
b3QgKCRjdXJyZW50LlBTT2JqZWN0LlByb3BlcnRpZXMgfCBXaGVyZS1PYmplY3QgeyAkXy5OYW1lIC1lcSAiaHlicmlkQXp1cmVBREpvaW5Ta2lwQ29ubmVj
dGl2aXR5Q2hlY2siIH0pKSB7DQogICAgICAgICAgICAgICAgICAgICRjdXJyZW50IHwgQWRkLU1lbWJlciAtTm90ZVByb3BlcnR5TmFtZSBoeWJyaWRBenVy
ZUFESm9pblNraXBDb25uZWN0aXZpdHlDaGVjayAtTm90ZVByb3BlcnR5VmFsdWUgJGZhbHNlDQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgfQ0K
DQogICAgICAgICAgICAjIEZvciBwYXJhbWV0ZXJzIHRoYXQgd2VyZSBzcGVjaWZpZWQsIHVwZGF0ZSB0aGF0IG9iamVjdCBpbiBwbGFjZQ0KICAgICAgICAg
ICAgaWYgKCRQU0JvdW5kUGFyYW1ldGVycy5Db250YWluc0tleSgnZGlzcGxheU5hbWUnKSkgeyAkY3VycmVudC5kaXNwbGF5TmFtZSA9ICRkaXNwbGF5TmFt
ZSB9DQogICAgICAgICAgICBpZiAoJFBTQm91bmRQYXJhbWV0ZXJzLkNvbnRhaW5zS2V5KCdkZXNjcmlwdGlvbicpKSB7ICRjdXJyZW50LmRlc2NyaXB0aW9u
ID0gJGRlc2NyaXB0aW9uIH0NCiAgICAgICAgICAgIGlmICgkUFNCb3VuZFBhcmFtZXRlcnMuQ29udGFpbnNLZXkoJ0NvbnZlcnREZXZpY2VUb0F1dG9waWxv
dCcpKSB7ICRjdXJyZW50LmV4dHJhY3RIYXJkd2FyZUhhc2ggPSBbYm9vbF0kQ29udmVydERldmljZVRvQXV0b3BpbG90IH0NCiAgICAgICAgICAgIGlmICgk
UFNCb3VuZFBhcmFtZXRlcnMuQ29udGFpbnNLZXkoJ09PQkVfbGFuZ3VhZ2UnKSkgeyAkY3VycmVudC5sYW5ndWFnZSA9ICRPT0JFX2xhbmd1YWdlIH0NCiAg
ICAgICAgICAgIGlmICgkUFNCb3VuZFBhcmFtZXRlcnMuQ29udGFpbnNLZXkoJ09PQkVfc2tpcEtleWJvYXJkJykpIHsgJGN1cnJlbnQub3V0T2ZCb3hFeHBl
cmllbmNlU2V0dGluZ3Muc2tpcEtleWJvYXJkU2VsZWN0aW9uUGFnZSA9IFtib29sXSRPT0JFX3NraXBLZXlib2FyZCB9DQogICAgICAgICAgICBpZiAoJFBT
Qm91bmRQYXJhbWV0ZXJzLkNvbnRhaW5zS2V5KCdPT0JFX05hbWVUZW1wbGF0ZScpKSB7ICRjdXJyZW50LmRldmljZU5hbWVUZW1wbGF0ZSA9ICRPT0JFX05h
bWVUZW1wbGF0ZSB9DQogICAgICAgICAgICBpZiAoJFBTQm91bmRQYXJhbWV0ZXJzLkNvbnRhaW5zS2V5KCdPT0JFX0VuYWJsZVdoaXRlR2xvdmUnKSkgeyAk
Y3VycmVudC5lbmFibGVXaGl0ZUdsb3ZlID0gW2Jvb2xdJE9PQkVfRW5hYmxlV2hpdGVHbG92ZSB9DQogICAgICAgICAgICBpZiAoJFBTQm91bmRQYXJhbWV0
ZXJzLkNvbnRhaW5zS2V5KCdPT0JFX1VzZXJUeXBlQWRtaW4nKSkgew0KICAgICAgICAgICAgICAgIGlmICgkT09CRV9Vc2VyVHlwZUFkbWluKSB7DQogICAg
ICAgICAgICAgICAgICAgICRjdXJyZW50Lm91dE9mQm94RXhwZXJpZW5jZVNldHRpbmdzLnVzZXJUeXBlID0gImFkbWluaXN0cmF0b3IiDQogICAgICAgICAg
ICAgICAgfQ0KICAgICAgICAgICAgICAgIGVsc2Ugew0KICAgICAgICAgICAgICAgICAgICAkY3VycmVudC5vdXRPZkJveEV4cGVyaWVuY2VTZXR0aW5ncy51
c2VyVHlwZSA9ICJzdGFuZGFyZCINCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICB9DQogICAgICAgICAgICBpZiAoJFBTQm91bmRQYXJhbWV0ZXJz
LkNvbnRhaW5zS2V5KCdPT0JFX0hpZGVFVUxBJykpIHsgJGN1cnJlbnQub3V0T2ZCb3hFeHBlcmllbmNlU2V0dGluZ3MuaGlkZUVVTEEgPSBbYm9vbF0kT09C
RV9IaWRlRVVMQSB9DQogICAgICAgICAgICBpZiAoJFBTQm91bmRQYXJhbWV0ZXJzLkNvbnRhaW5zS2V5KCdPT0JFX0hpZGVQcml2YWN5U2V0dGluZ3MnKSkg
eyAkY3VycmVudC5vdXRPZkJveEV4cGVyaWVuY2VTZXR0aW5ncy5oaWRlUHJpdmFjeVNldHRpbmdzID0gW2Jvb2xdJE9PQkVfSGlkZVByaXZhY3lTZXR0aW5n
cyB9DQogICAgICAgICAgICBpZiAoJFBTQm91bmRQYXJhbWV0ZXJzLkNvbnRhaW5zS2V5KCdPT0JFX0hpZGVDaGFuZ2VBY2NvdW50T3B0cycpKSB7ICRjdXJy
ZW50Lm91dE9mQm94RXhwZXJpZW5jZVNldHRpbmdzLmhpZGVFc2NhcGVMaW5rID0gW2Jvb2xdJE9PQkVfSGlkZUNoYW5nZUFjY291bnRPcHRzIH0NCiAgICAg
ICAgICAgIGlmICgkUFNCb3VuZFBhcmFtZXRlcnMuQ29udGFpbnNLZXkoJ09PQkVfU2tpcENvbm5lY3Rpdml0eUNoZWNrJykpIHsgJGN1cnJlbnQuaHlicmlk
QXp1cmVBREpvaW5Ta2lwQ29ubmVjdGl2aXR5Q2hlY2sgPSBbYm9vbF0kT09CRV9Ta2lwQ29ubmVjdGl2aXR5Q2hlY2sgfQ0KDQogICAgICAgICAgICBpZiAo
JEFsbEVuYWJsZWQpIHsNCiAgICAgICAgICAgICAgICAkY3VycmVudC5leHRyYWN0SGFyZHdhcmVIYXNoID0gJHRydWUNCiAgICAgICAgICAgICAgICAkY3Vy
cmVudC5vdXRPZkJveEV4cGVyaWVuY2VTZXR0aW5ncy5oaWRlUHJpdmFjeVNldHRpbmdzID0gJHRydWUNCiAgICAgICAgICAgICAgICAkY3VycmVudC5vdXRP
ZkJveEV4cGVyaWVuY2VTZXR0aW5ncy5oaWRlRXNjYXBlTGluayA9ICR0cnVlDQogICAgICAgICAgICAgICAgJGN1cnJlbnQuaHlicmlkQXp1cmVBREpvaW5T
a2lwQ29ubmVjdGl2aXR5Q2hlY2sgPSAkdHJ1ZQ0KICAgICAgICAgICAgICAgICRjdXJyZW50LkVuYWJsZVdoaXRlR2xvdmUgPSAkdHJ1ZQ0KICAgICAgICAg
ICAgICAgICRjdXJyZW50Lm91dE9mQm94RXhwZXJpZW5jZVNldHRpbmdzLmhpZGVFVUxBID0gJHRydWUgDQogICAgICAgICAgICAgICAgJGN1cnJlbnQub3V0
T2ZCb3hFeHBlcmllbmNlU2V0dGluZ3MuaGlkZVByaXZhY3lTZXR0aW5ncyA9ICR0cnVlDQogICAgICAgICAgICAgICAgJGN1cnJlbnQub3V0T2ZCb3hFeHBl
cmllbmNlU2V0dGluZ3MuaGlkZUVzY2FwZUxpbmsgPSAkdHJ1ZQ0KICAgICAgICAgICAgICAgICRjdXJyZW50Lm91dE9mQm94RXhwZXJpZW5jZVNldHRpbmdz
LnNraXBLZXlib2FyZFNlbGVjdGlvblBhZ2UgPSAkdHJ1ZQ0KICAgICAgICAgICAgICAgICRjdXJyZW50Lm91dE9mQm94RXhwZXJpZW5jZVNldHRpbmdzLnVz
ZXJUeXBlID0gImFkbWluaXN0cmF0b3IiDQogICAgICAgICAgICB9DQogICAgICAgICAgICBlbHNlaWYgKCRBbGxEaXNhYmxlZCkgew0KICAgICAgICAgICAg
ICAgICRjdXJyZW50LmV4dHJhY3RIYXJkd2FyZUhhc2ggPSAkZmFsc2UNCiAgICAgICAgICAgICAgICAkY3VycmVudC5vdXRPZkJveEV4cGVyaWVuY2VTZXR0
aW5ncy5oaWRlUHJpdmFjeVNldHRpbmdzID0gJGZhbHNlDQogICAgICAgICAgICAgICAgJGN1cnJlbnQub3V0T2ZCb3hFeHBlcmllbmNlU2V0dGluZ3MuaGlk
ZUVzY2FwZUxpbmsgPSAkZmFsc2UNCiAgICAgICAgICAgICAgICAkY3VycmVudC5oeWJyaWRBenVyZUFESm9pblNraXBDb25uZWN0aXZpdHlDaGVjayA9ICRm
YWxzZQ0KICAgICAgICAgICAgICAgICRjdXJyZW50LkVuYWJsZVdoaXRlR2xvdmUgPSAkZmFsc2UNCiAgICAgICAgICAgICAgICAkY3VycmVudC5vdXRPZkJv
eEV4cGVyaWVuY2VTZXR0aW5ncy5oaWRlRVVMQSA9ICRmYWxzZQ0KICAgICAgICAgICAgICAgICRjdXJyZW50Lm91dE9mQm94RXhwZXJpZW5jZVNldHRpbmdz
LmhpZGVQcml2YWN5U2V0dGluZ3MgPSAkZmFsc2UNCiAgICAgICAgICAgICAgICAkY3VycmVudC5vdXRPZkJveEV4cGVyaWVuY2VTZXR0aW5ncy5oaWRlRXNj
YXBlTGluayA9ICRmYWxzZQ0KICAgICAgICAgICAgICAgICRjdXJyZW50Lm91dE9mQm94RXhwZXJpZW5jZVNldHRpbmdzLnNraXBLZXlib2FyZFNlbGVjdGlv
blBhZ2UgPSAkZmFsc2UNCiAgICAgICAgICAgICAgICAkY3VycmVudC5vdXRPZkJveEV4cGVyaWVuY2VTZXR0aW5ncy51c2VyVHlwZSA9ICJzdGFuZGFyZCIN
CiAgICAgICAgICAgIH0NCg0KICAgICAgICAgICAgIyBDbGVhbiB1cCB1bm5lZWRlZCBwcm9wZXJ0aWVzDQogICAgICAgICAgICAkY3VycmVudC5QU09iamVj
dC5Qcm9wZXJ0aWVzLlJlbW92ZSgibGFzdE1vZGlmaWVkRGF0ZVRpbWUiKQ0KICAgICAgICAgICAgJGN1cnJlbnQuUFNPYmplY3QuUHJvcGVydGllcy5SZW1v
dmUoImNyZWF0ZWREYXRlVGltZSIpIA0KICAgICAgICAgICAgJGN1cnJlbnQuUFNPYmplY3QuUHJvcGVydGllcy5SZW1vdmUoIkBvZGF0YS5jb250ZXh0IikN
CiAgICAgICAgICAgICRjdXJyZW50LlBTT2JqZWN0LlByb3BlcnRpZXMuUmVtb3ZlKCJpZCIpDQogICAgICAgICAgICAkY3VycmVudC5QU09iamVjdC5Qcm9w
ZXJ0aWVzLlJlbW92ZSgicm9sZVNjb3BlVGFnSWRzIikNCg0KICAgICAgICAgICAgIyBEZWZpbmluZyBWYXJpYWJsZXMNCiAgICAgICAgICAgICRncmFwaEFw
aVZlcnNpb24gPSAiYmV0YSINCiAgICAgICAgICAgICRSZXNvdXJjZSA9ICJkZXZpY2VNYW5hZ2VtZW50L3dpbmRvd3NBdXRvcGlsb3REZXBsb3ltZW50UHJv
ZmlsZXMiDQogICAgICAgICAgICAkdXJpID0gImh0dHBzOi8vZ3JhcGgubWljcm9zb2Z0LmNvbS8kZ3JhcGhBcGlWZXJzaW9uLyRSZXNvdXJjZS8kaWQiDQog
ICAgICAgICAgICAkanNvbiA9ICgkY3VycmVudCB8IENvbnZlcnRUby1Kc29uKS5Ub1N0cmluZygpDQogICAgDQogICAgICAgICAgICBXcml0ZS1WZXJib3Nl
ICJQQVRDSCAkdXJpYG4kanNvbiINCg0KICAgICAgICAgICAgdHJ5IHsNCiAgICAgICAgICAgICAgICBJbnZva2UtTWdHcmFwaFJlcXVlc3QgLVVyaSAkdXJp
IC1NZXRob2QgUEFUQ0ggLUJvZHkgJGpzb24gLUNvbnRlbnRUeXBlICJhcHBsaWNhdGlvbi9qc29uIiAtT3V0cHV0VHlwZSBQU09iamVjdA0KICAgICAgICAg
ICAgfQ0KICAgICAgICAgICAgY2F0Y2ggew0KICAgICAgICAgICAgICAgIFdyaXRlLUVycm9yICRfLkV4Y2VwdGlvbiANCiAgICAgICAgICAgICAgICBicmVh
aw0KICAgICAgICAgICAgfQ0KDQogICAgICAgIH0NCg0KDQogICAgICAgIEZ1bmN0aW9uIE5ldy1BdXRvcGlsb3RQcm9maWxlKCkgew0KICAgICAgICAgICAg
PCMNCi5TWU5PUFNJUw0KQ3JlYXRlcyBhIG5ldyBBdXRvcGlsb3QgcHJvZmlsZS4NCiANCi5ERVNDUklQVElPTg0KVGhlIE5ldy1BdXRvcGlsb3RQcm9maWxl
IGNyZWF0ZXMgYSBuZXcgQXV0b3BpbG90IHByb2ZpbGUuDQogDQouUEFSQU1FVEVSIGRpc3BsYXlOYW1lDQpUaGUgbmFtZSBvZiB0aGUgV2luZG93cyBBdXRv
cGlsb3QgcHJvZmlsZSB0byBjcmVhdGUuIChUaGlzIHZhbHVlIGNhbm5vdCBjb250YWluIHNwYWNlcy4pDQogDQouUEFSQU1FVEVSIG1vZGUNClRoZSB0eXBl
IG9mIEF1dG9waWxvdCBwcm9maWxlIHRvIGNyZWF0ZS4gQ2hvaWNlcyBhcmUgIlVzZXJEcml2ZW5BQUQiLCAiVXNlckRyaXZlbkFEIiwgYW5kICJTZWxmRGVw
bG95aW5nQUFEIi4NCiANCi5QQVJBTUVURVIgZGVzY3JpcHRpb24NClRoZSBkZXNjcmlwdGlvbiB0byBiZSBjb25maWd1cmVkIGluIHRoZSBwcm9maWxlLiAo
VGhpcyB2YWx1ZSBjYW5ub3QgY29udGFpbiBkYXNoZXMuKQ0KICAgICANCi5QQVJBTUVURVIgQ29udmVydERldmljZVRvQXV0b3BpbG90DQpDb25maWd1cmUg
dGhlIHZhbHVlICJDb252ZXJ0IGFsbCB0YXJnZXRlZCBkZXZpY2VzIHRvIEF1dG9waWxvdCINCiANCi5QQVJBTUVURVIgT09CRV9IaWRlRVVMQQ0KQ29uZmln
dXJlIHRoZSBPT0JFIG9wdGlvbiB0byBoaWRlIG9yIG5vdCB0aGUgRVVMQQ0KIA0KLlBBUkFNRVRFUiBPT0JFX0VuYWJsZVdoaXRlR2xvdmUNCkNvbmZpZ3Vy
ZSB0aGUgT09CRSBvcHRpb24gdG8gYWxsb3cgb3Igbm90IFdoaXRlIEdsb3ZlIE9PQkUNCiANCi5QQVJBTUVURVIgT09CRV9IaWRlUHJpdmFjeVNldHRpbmdz
DQpDb25maWd1cmUgdGhlIE9PQkUgb3B0aW9uIHRvIGhpZGUgb3Igbm90IHRoZSBwcml2YWN5IHNldHRpbmdzDQogDQouUEFSQU1FVEVSIE9PQkVfSGlkZUNo
YW5nZUFjY291bnRPcHRzDQpDb25maWd1cmUgdGhlIE9PQkUgb3B0aW9uIHRvIGhpZGUgb3Igbm90IHRoZSBjaGFuZ2UgYWNjb3VudCBvcHRpb25zDQogDQou
UEFSQU1FVEVSIE9PQkVfVXNlclR5cGVBZG1pbg0KQ29uZmlndXJlIHRoZSB1c2VyIGFjY291bnQgdHlwZSBhcyBhZG1pbmlzdHJhdG9yLg0KIA0KLlBBUkFN
RVRFUiBPT0JFX05hbWVUZW1wbGF0ZQ0KQ29uZmlndXJlIHRoZSBPT0JFIG9wdGlvbiB0byBhcHBseSBhIGRldmljZSBuYW1lIHRlbXBsYXRlDQogDQouUEFS
QU1FVEVSIE9PQkVfbGFuZ3VhZ2UNClRoZSBsYW5ndWFnZSBpZGVudGlmaWVyIChlLmcuICJlbi11cyIpIHRvIGJlIGNvbmZpZ3VyZWQgaW4gdGhlIHByb2Zp
bGUNCiANCi5QQVJBTUVURVIgT09CRV9Ta2lwS2V5Ym9hcmQNCkNvbmZpZ3VyZSB0aGUgT09CRSBvcHRpb24gdG8gc2tpcCBvciBub3QgdGhlIGtleWJvYXJk
IHNlbGVjdGlvbiBwYWdlDQogDQouUEFSQU1FVEVSIE9PQkVfSGlkZUNoYW5nZUFjY291bnRPcHRzDQpDb25maWd1cmUgdGhlIE9PQkUgb3B0aW9uIHRvIGhp
ZGUgb3Igbm90IHRoZSBjaGFuZ2UgYWNjb3VudCBvcHRpb25zDQogDQouUEFSQU1FVEVSIE9PQkVfU2tpcENvbm5lY3Rpdml0eUNoZWNrDQpTcGVjaWZ5IHdo
ZXRoZXIgdG8gc2tpcCBBY3RpdmUgRGlyZWN0b3J5IGNvbm5lY3Rpdml0eSBjaGVja3MgKFVzZXJEcml2ZW5BQUQgb25seSkNCiANCi5FWEFNUExFDQpDcmVh
dGUgcHJvZmlsZXMgb2YgZGlmZmVyZW50IHR5cGVzOg0KIA0KTmV3LUF1dG9waWxvdFByb2ZpbGUgLW1vZGUgVXNlckRyaXZlbkFBRCAtZGlzcGxheU5hbWUg
Ik15IEFBRCBwcm9maWxlIiAtZGVzY3JpcHRpb24gIk15IHVzZXItZHJpdmVuIEFBRCBwcm9maWxlIiAtT09CRV9RdWlldA0KTmV3LUF1dG9waWxvdFByb2Zp
bGUgLW1vZGUgVXNlckRyaXZlbkFEIC1kaXNwbGF5TmFtZSAiTXkgQUQgcHJvZmlsZSIgLWRlc2NyaXB0aW9uICJNeSB1c2VyLWRyaXZlbiBBRCBwcm9maWxl
IiAtT09CRV9RdWlldA0KTmV3LUF1dG9waWxvdFByb2ZpbGUgLW1vZGUgU2VsZkRlcGxveWluZ0FBRCAtZGlzcGxheU5hbWUgIk15IFNlbGYgRGVwbG95aW5n
IHByb2ZpbGUiIC1kZXNjcmlwdGlvbiAiTXkgc2VsZi1kZXBsb3lpbmcgcHJvZmlsZSIgLU9PQkVfUXVpZXQNCiANCi5FWEFNUExFDQpDcmVhdGUgYSB1c2Vy
LWRyaXZlbiBBQUQgcHJvZmlsZToNCiANCk5ldy1BdXRvcGlsb3RQcm9maWxlIC1tb2RlIFVzZXJEcml2ZW5BQUQgLWRpc3BsYXlOYW1lICJNeSB0ZXN0aW5n
IHByb2ZpbGUiIC1EZXNjcmlwdGlvbiAiRGVzY3JpcHRpb24gb2YgbXkgcHJvZmlsZSIgLU9PQkVfTGFuZ3VhZ2UgImVuLXVzIiAtT09CRV9IaWRlRVVMQSAt
T09CRV9IaWRlUHJpdmFjeVNldHRpbmdzDQogDQojPg0KICAgICAgICAgICAgW2NtZGxldGJpbmRpbmcoKV0NCiAgICAgICAgICAgIHBhcmFtDQogICAgICAg
ICAgICAoDQogICAgICAgICAgICAgICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkdHJ1ZSldW3N0cmluZ10gJGRpc3BsYXlOYW1lLA0KICAgICAgICAgICAg
ICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJHRydWUpXVtWYWxpZGF0ZVNldCgnVXNlckRyaXZlbkFBRCcsICdVc2VyRHJpdmVuQUQnLCAnU2VsZkRlcGxv
eWluZ0FBRCcpXVtzdHJpbmddICRtb2RlLCANCiAgICAgICAgICAgICAgICBbc3RyaW5nXSAkZGVzY3JpcHRpb24sDQogICAgICAgICAgICAgICAgW1N3aXRj
aF0gJENvbnZlcnREZXZpY2VUb0F1dG9waWxvdCwNCiAgICAgICAgICAgICAgICBbc3RyaW5nXSAkT09CRV9sYW5ndWFnZSwNCiAgICAgICAgICAgICAgICBb
U3dpdGNoXSAkT09CRV9za2lwS2V5Ym9hcmQsDQogICAgICAgICAgICAgICAgW3N0cmluZ10gJE9PQkVfTmFtZVRlbXBsYXRlLA0KICAgICAgICAgICAgICAg
IFtTd2l0Y2hdICRPT0JFX0VuYWJsZVdoaXRlR2xvdmUsDQogICAgICAgICAgICAgICAgW1N3aXRjaF0gJE9PQkVfVXNlclR5cGVBZG1pbiwNCiAgICAgICAg
ICAgICAgICBbU3dpdGNoXSAkT09CRV9IaWRlRVVMQSwNCiAgICAgICAgICAgICAgICBbU3dpdGNoXSAkT09CRV9oaWRlUHJpdmFjeVNldHRpbmdzLA0KICAg
ICAgICAgICAgICAgIFtTd2l0Y2hdICRPT0JFX0hpZGVDaGFuZ2VBY2NvdW50T3B0cywNCiAgICAgICAgICAgICAgICBbU3dpdGNoXSAkT09CRV9Ta2lwQ29u
bmVjdGl2aXR5Q2hlY2sNCiAgICAgICAgICAgICkNCg0KICAgICAgICAgICAgIyBBZGp1c3QgdmFsdWVzIGFzIG5lZWRlZA0KICAgICAgICAgICAgc3dpdGNo
ICgkbW9kZSkgew0KICAgICAgICAgICAgICAgICJVc2VyRHJpdmVuQUFEIiB7ICRvZGF0YVR5cGUgPSAiI21pY3Jvc29mdC5ncmFwaC5henVyZUFEV2luZG93
c0F1dG9waWxvdERlcGxveW1lbnRQcm9maWxlIjsgJHVzYWdlID0gInNpbmdsZVVzZXIiIH0NCiAgICAgICAgICAgICAgICAiU2VsZkRlcGxveWluZ0FBRCIg
eyAkb2RhdGFUeXBlID0gIiNtaWNyb3NvZnQuZ3JhcGguYXp1cmVBRFdpbmRvd3NBdXRvcGlsb3REZXBsb3ltZW50UHJvZmlsZSI7ICR1c2FnZSA9ICJzaGFy
ZWQiIH0NCiAgICAgICAgICAgICAgICAiVXNlckRyaXZlbkFEIiB7ICRvZGF0YVR5cGUgPSAiI21pY3Jvc29mdC5ncmFwaC5hY3RpdmVEaXJlY3RvcnlXaW5k
b3dzQXV0b3BpbG90RGVwbG95bWVudFByb2ZpbGUiOyAkdXNhZ2UgPSAic2luZ2xlVXNlciIgfQ0KICAgICAgICAgICAgfQ0KDQogICAgICAgICAgICBpZiAo
JE9PQkVfVXNlclR5cGVBZG1pbikgeyAgICAgICAgDQogICAgICAgICAgICAgICAgJE9PQkVfdXNlclR5cGUgPSAiYWRtaW5pc3RyYXRvciINCiAgICAgICAg
ICAgIH0NCiAgICAgICAgICAgIGVsc2UgeyAgICAgICAgDQogICAgICAgICAgICAgICAgJE9PQkVfdXNlclR5cGUgPSAic3RhbmRhcmQiDQogICAgICAgICAg
ICB9ICAgICAgICANCg0KICAgICAgICAgICAgaWYgKCRPT0JFX0VuYWJsZVdoaXRlR2xvdmUpIHsgICAgICAgIA0KICAgICAgICAgICAgICAgICRPT0JFX0hp
ZGVDaGFuZ2VBY2NvdW50T3B0cyA9ICRUcnVlDQogICAgICAgICAgICB9ICAgICAgICANCiAgICAgICAgDQogICAgICAgICAgICAjIERlZmluaW5nIFZhcmlh
Ymxlcw0KICAgICAgICAgICAgJGdyYXBoQXBpVmVyc2lvbiA9ICJiZXRhIg0KICAgICAgICAgICAgJFJlc291cmNlID0gImRldmljZU1hbmFnZW1lbnQvd2lu
ZG93c0F1dG9waWxvdERlcGxveW1lbnRQcm9maWxlcyINCiAgICAgICAgICAgICR1cmkgPSAiaHR0cHM6Ly9ncmFwaC5taWNyb3NvZnQuY29tLyRncmFwaEFw
aVZlcnNpb24vJFJlc291cmNlIg0KICAgICAgICAgICAgaWYgKCRtb2RlIC1lcSAiVXNlckRyaXZlbkFEIikgew0KICAgICAgICAgICAgICAgICRqc29uID0g
QCINCnsNCiAgICAiQG9kYXRhLnR5cGUiOiAiJG9kYXRhVHlwZSIsDQogICAgImRpc3BsYXlOYW1lIjogIiRkaXNwbGF5bmFtZSIsDQogICAgImRlc2NyaXB0
aW9uIjogIiRkZXNjcmlwdGlvbiIsDQogICAgImxhbmd1YWdlIjogIiRPT0JFX2xhbmd1YWdlIiwNCiAgICAiZXh0cmFjdEhhcmR3YXJlSGFzaCI6ICQoQm9v
bFRvU3RyaW5nKCRDb252ZXJ0RGV2aWNlVG9BdXRvcGlsb3QpKSwNCiAgICAiZGV2aWNlTmFtZVRlbXBsYXRlIjogIiRPT0JFX05hbWVUZW1wbGF0ZSIsDQog
ICAgImRldmljZVR5cGUiOiAid2luZG93c1BjIiwNCiAgICAiZW5hYmxlV2hpdGVHbG92ZSI6ICQoQm9vbFRvU3RyaW5nKCRPT0JFX0VuYWJsZVdoaXRlR2xv
dmUpKSwNCiAgICAiaHlicmlkQXp1cmVBREpvaW5Ta2lwQ29ubmVjdGl2aXR5Q2hlY2siOiAkKEJvb2xUb1N0cmluZygkT09CRV9Ta2lwQ29ubmVjdGl2aXR5
Q2hlY2tzKSksDQogICAgIm91dE9mQm94RXhwZXJpZW5jZVNldHRpbmdzIjogew0KICAgICAgICAiaGlkZVByaXZhY3lTZXR0aW5ncyI6ICQoQm9vbFRvU3Ry
aW5nKCRPT0JFX2hpZGVQcml2YWN5U2V0dGluZ3MpKSwNCiAgICAgICAgImhpZGVFVUxBIjogJChCb29sVG9TdHJpbmcoJE9PQkVfSGlkZUVVTEEpKSwNCiAg
ICAgICAgInVzZXJUeXBlIjogIiRPT0JFX3VzZXJUeXBlIiwNCiAgICAgICAgImRldmljZVVzYWdlVHlwZSI6ICIkdXNhZ2UiLA0KICAgICAgICAic2tpcEtl
eWJvYXJkU2VsZWN0aW9uUGFnZSI6ICQoQm9vbFRvU3RyaW5nKCRPT0JFX3NraXBLZXlib2FyZCkpLA0KICAgICAgICAiaGlkZUVzY2FwZUxpbmsiOiAkKEJv
b2xUb1N0cmluZygkT09CRV9IaWRlQ2hhbmdlQWNjb3VudE9wdHMpKQ0KICAgIH0NCn0NCiJADQogICAgICAgICAgICB9DQogICAgICAgICAgICBlbHNlIHsN
CiAgICAgICAgICAgICAgICAkanNvbiA9IEAiDQp7DQogICAgIkBvZGF0YS50eXBlIjogIiRvZGF0YVR5cGUiLA0KICAgICJkaXNwbGF5TmFtZSI6ICIkZGlz
cGxheW5hbWUiLA0KICAgICJkZXNjcmlwdGlvbiI6ICIkZGVzY3JpcHRpb24iLA0KICAgICJsYW5ndWFnZSI6ICIkT09CRV9sYW5ndWFnZSIsDQogICAgImV4
dHJhY3RIYXJkd2FyZUhhc2giOiAkKEJvb2xUb1N0cmluZygkQ29udmVydERldmljZVRvQXV0b3BpbG90KSksDQogICAgImRldmljZU5hbWVUZW1wbGF0ZSI6
ICIkT09CRV9OYW1lVGVtcGxhdGUiLA0KICAgICJkZXZpY2VUeXBlIjogIndpbmRvd3NQYyIsDQogICAgImVuYWJsZVdoaXRlR2xvdmUiOiAkKEJvb2xUb1N0
cmluZygkT09CRV9FbmFibGVXaGl0ZUdsb3ZlKSksDQogICAgIm91dE9mQm94RXhwZXJpZW5jZVNldHRpbmdzIjogew0KICAgICAgICAiaGlkZVByaXZhY3lT
ZXR0aW5ncyI6ICQoQm9vbFRvU3RyaW5nKCRPT0JFX2hpZGVQcml2YWN5U2V0dGluZ3MpKSwNCiAgICAgICAgImhpZGVFVUxBIjogJChCb29sVG9TdHJpbmco
JE9PQkVfSGlkZUVVTEEpKSwNCiAgICAgICAgInVzZXJUeXBlIjogIiRPT0JFX3VzZXJUeXBlIiwNCiAgICAgICAgImRldmljZVVzYWdlVHlwZSI6ICIkdXNh
Z2UiLA0KICAgICAgICAic2tpcEtleWJvYXJkU2VsZWN0aW9uUGFnZSI6ICQoQm9vbFRvU3RyaW5nKCRPT0JFX3NraXBLZXlib2FyZCkpLA0KICAgICAgICAi
aGlkZUVzY2FwZUxpbmsiOiAkKEJvb2xUb1N0cmluZygkT09CRV9IaWRlQ2hhbmdlQWNjb3VudE9wdHMpKQ0KICAgIH0NCn0NCiJADQogICAgICAgICAgICB9
DQoNCiAgICAgICAgICAgIFdyaXRlLVZlcmJvc2UgIlBPU1QgJHVyaWBuJGpzb24iDQoNCiAgICAgICAgICAgIHRyeSB7DQogICAgICAgICAgICAgICAgSW52
b2tlLU1nR3JhcGhSZXF1ZXN0IC1VcmkgJHVyaSAtTWV0aG9kIFBPU1QgLUJvZHkgJGpzb24gLUNvbnRlbnRUeXBlICJhcHBsaWNhdGlvbi9qc29uIiAtT3V0
cHV0VHlwZSBQU09iamVjdA0KICAgICAgICAgICAgfQ0KICAgICAgICAgICAgY2F0Y2ggew0KICAgICAgICAgICAgICAgIFdyaXRlLUVycm9yICRfLkV4Y2Vw
dGlvbiANCiAgICAgICAgICAgICAgICBicmVhaw0KICAgICAgICAgICAgfQ0KDQogICAgICAgIH0NCg0KDQogICAgICAgIEZ1bmN0aW9uIFJlbW92ZS1BdXRv
cGlsb3RQcm9maWxlKCkgew0KICAgICAgICAgICAgPCMNCi5TWU5PUFNJUw0KUmVtb3ZlIGEgRGVwbG95bWVudCBQcm9maWxlDQouREVTQ1JJUFRJT04NClRo
ZSBSZW1vdmUtQXV0b3BpbG90UHJvZmlsZSBhbGxvd3MgeW91IHRvIHJlbW92ZSBhIHNwZWNpZmljIGRlcGxveW1lbnQgcHJvZmlsZQ0KLlBBUkFNRVRFUiBp
ZA0KTWFuZGF0b3J5LCB0aGUgSUQgKEdVSUQpIG9mIHRoZSBwcm9maWxlIHRvIGJlIHJlbW92ZWQuDQouRVhBTVBMRQ0KUmVtb3ZlLUF1dG9waWxvdFByb2Zp
bGUgLWlkICRpZA0KIz4NCiAgICAgICAgICAgIFtjbWRsZXRiaW5kaW5nKCldDQogICAgICAgICAgICBwYXJhbQ0KICAgICAgICAgICAgKA0KICAgICAgICAg
ICAgICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJFRydWUsIFZhbHVlRnJvbVBpcGVsaW5lQnlQcm9wZXJ0eU5hbWUgPSAkVHJ1ZSldICRpZA0KICAgICAg
ICAgICAgKQ0KDQogICAgICAgICAgICBQcm9jZXNzIHsNCiAgICAgICAgICAgICAgICAjIERlZmluaW5nIFZhcmlhYmxlcw0KICAgICAgICAgICAgICAgICRn
cmFwaEFwaVZlcnNpb24gPSAiYmV0YSINCiAgICAgICAgICAgICAgICAkUmVzb3VyY2UgPSAiZGV2aWNlTWFuYWdlbWVudC93aW5kb3dzQXV0b3BpbG90RGVw
bG95bWVudFByb2ZpbGVzIg0KICAgICAgICAgICAgICAgICR1cmkgPSAiaHR0cHM6Ly9ncmFwaC5taWNyb3NvZnQuY29tLyRncmFwaEFwaVZlcnNpb24vJFJl
c291cmNlLyRpZCINCg0KICAgICAgICAgICAgICAgIFdyaXRlLVZlcmJvc2UgIkRFTEVURSAkdXJpIg0KDQogICAgICAgICAgICAgICAgVHJ5IHsNCiAgICAg
ICAgICAgICAgICAgICAgSW52b2tlLU1nR3JhcGhSZXF1ZXN0IC1VcmkgJHVyaSAtTWV0aG9kIERFTEVURQ0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAg
ICAgICAgICBjYXRjaCB7DQogICAgICAgICAgICAgICAgICAgIFdyaXRlLUVycm9yICRfLkV4Y2VwdGlvbiANCiAgICAgICAgICAgICAgICAgICAgYnJlYWsN
CiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICB9DQogICAgICAgIH0NCg0KDQogICAgICAgIEZ1bmN0aW9uIEdldC1BdXRvcGlsb3RQcm9maWxlQXNz
aWdubWVudHMoKSB7DQogICAgICAgICAgICA8Iw0KLlNZTk9QU0lTDQpMaXN0IGFsbCBhc3NpZ25lZCBkZXZpY2VzIGZvciBhIHNwZWNpZmljIHByb2ZpbGUg
SUQNCi5ERVNDUklQVElPTg0KVGhlIEdldC1BdXRvcGlsb3RQcm9maWxlQXNzaWdubWVudHMgY21kbGV0IHJldHVybnMgdGhlIGxpc3Qgb2YgZ3JvdXBzIHRo
YXQgYWUgYXNzaWduZWQgdG8gYSBzcGNpZmljIGRlcGxveW1lbnQgcHJvZmlsZQ0KLlBBUkFNRVRFUiBpZA0KVHlwZTogSW50ZWdlciAtIE1hbmRhdG9yeSwg
dGhlIElEIChHVUlEKSBvZiB0aGUgcHJvZmlsZSB0byBiZSByZXRyaWV2ZWQuDQouRVhBTVBMRQ0KR2V0LUF1dG9waWxvdFByb2ZpbGVBc3NpZ25tZW50cyAt
aWQgJGlkDQojPg0KICAgICAgICAgICAgW2NtZGxldGJpbmRpbmcoKV0NCiAgICAgICAgICAgIHBhcmFtDQogICAgICAgICAgICAoDQogICAgICAgICAgICAg
ICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkdHJ1ZSwgVmFsdWVGcm9tUGlwZWxpbmVCeVByb3BlcnR5TmFtZSA9ICRUcnVlKV0gJGlkDQogICAgICAgICAg
ICApDQoNCiAgICAgICAgICAgIFByb2Nlc3Mgew0KDQogICAgICAgICAgICAgICAgIyBEZWZpbmluZyBWYXJpYWJsZXMNCiAgICAgICAgICAgICAgICAkZ3Jh
cGhBcGlWZXJzaW9uID0gImJldGEiDQogICAgICAgICAgICAgICAgJFJlc291cmNlID0gImRldmljZU1hbmFnZW1lbnQvd2luZG93c0F1dG9waWxvdERlcGxv
eW1lbnRQcm9maWxlcyINCiAgICAgICAgICAgICAgICAkdXJpID0gImh0dHBzOi8vZ3JhcGgubWljcm9zb2Z0LmNvbS8kZ3JhcGhBcGlWZXJzaW9uLyRSZXNv
dXJjZS8kaWQvYXNzaWdubWVudHMiDQoNCiAgICAgICAgICAgICAgICBXcml0ZS1WZXJib3NlICJHRVQgJHVyaSINCg0KICAgICAgICAgICAgICAgIHRyeSB7
DQogICAgICAgICAgICAgICAgICAgICRyZXNwb25zZSA9IEludm9rZS1NZ0dyYXBoUmVxdWVzdCAtVXJpICR1cmkgLU1ldGhvZCBHZXQNCiAgICAgICAgICAg
ICAgICAgICAgJEdyb3VwX0lEID0gJHJlc3BvbnNlLlZhbHVlLnRhcmdldC5ncm91cElkDQogICAgICAgICAgICAgICAgICAgIEZvckVhY2ggKCRHcm91cCBp
biAkR3JvdXBfSUQpIHsNCiAgICAgICAgICAgICAgICAgICAgICAgIFRyeSB7DQogICAgICAgICAgICAgICAgICAgICAgICAgICAgI0dldC1NZ0dyb3VwIHwg
V2hlcmUtT2JqZWN0IHsgJF8uT2JqZWN0SWQgLWxpa2UgJEdyb3VwIH0NCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAjI1VzZSBHcmFwaCBSZXF1ZXN0
DQogICAgICAgICAgICAgICAgICAgICAgICAgICAgJGd1cmkgPSAiaHR0cHM6Ly9ncmFwaC5taWNyb3NvZnQuY29tL2JldGEvZ3JvdXBzP2AkZmlsdGVyPWlk
IGVxICckR3JvdXAnIg0KICAgICAgICAgICAgICAgICAgICAgICAgICAgIChJbnZva2UtTWdHcmFwaFJlcXVlc3QgLVVyaSAkZ3VyaSAtTWV0aG9kIEdFVCAt
T3V0cHV0VHlwZSBQU09iamVjdCkudmFsdWUNCiAgICAgICAgICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICAgICAgICAgIENhdGNoIHsNCiAg
ICAgICAgICAgICAgICAgICAgICAgICAgICAkR3JvdXANCiAgICAgICAgICAgICAgICAgICAgICAgIH0gICAgICAgICAgICANCiAgICAgICAgICAgICAgICAg
ICAgfQ0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBjYXRjaCB7DQogICAgICAgICAgICAgICAgICAgIFdyaXRlLUVycm9yICRfLkV4Y2Vw
dGlvbiANCiAgICAgICAgICAgICAgICAgICAgYnJlYWsNCiAgICAgICAgICAgICAgICB9DQoNCiAgICAgICAgICAgIH0NCg0KICAgICAgICB9DQoNCg0KICAg
ICAgICBGdW5jdGlvbiBSZW1vdmUtQXV0b3BpbG90UHJvZmlsZUFzc2lnbm1lbnRzKCkgew0KICAgICAgICAgICAgPCMNCi5TWU5PUFNJUw0KUmVtb3ZlcyBh
IHNwZWNpZmljIGdyb3VwIGFzc2lnbnRpb24gZm9yIGEgc3BlY2lmYyBkZXBsb3ltZW50IHByb2ZpbGUNCi5ERVNDUklQVElPTg0KVGhlIFJlbW92ZS1BdXRv
cGlsb3RQcm9maWxlQXNzaWdubWVudHMgY21kbGV0IGFsbG93cyB5b3UgdG8gcmVtb3ZlIGEgZ3JvdXAgYXNzaWduYXRpb24gZm9yIGEgZGVwbG95bWVudCBw
cm9maWxlDQouUEFSQU1FVEVSIGlkDQpUeXBlOiBJbnRlZ2VyIC0gTWFuZGF0b3J5LCB0aGUgSUQgKEdVSUQpIG9mIHRoZSBwcm9maWxlDQouUEFSQU1FVEVS
IGdyb3VwaWQNClR5cGU6IEludGVnZXIgLSBNYW5kYXRvcnksIHRoZSBJRCBvZiB0aGUgZ3JvdXANCi5FWEFNUExFDQpSZW1vdmUtQXV0b3BpbG90UHJvZmls
ZUFzc2lnbm1lbnRzIC1pZCAkaWQNCiM+DQogICAgICAgICAgICBbY21kbGV0YmluZGluZygpXQ0KICAgICAgICAgICAgcGFyYW0NCiAgICAgICAgICAgICgN
CiAgICAgICAgICAgICAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICR0cnVlKV0kaWQsDQogICAgICAgICAgICAgICAgW1BhcmFtZXRlcihNYW5kYXRvcnkg
PSAkdHJ1ZSldJGdyb3VwaWQNCiAgICAgICAgICAgICkNCiAgICAgICAgICAgICMgRGVmaW5pbmcgVmFyaWFibGVzDQogICAgICAgICAgICAkZ3JhcGhBcGlW
ZXJzaW9uID0gImJldGEiDQogICAgICAgICAgICAkUmVzb3VyY2UgPSAiZGV2aWNlTWFuYWdlbWVudC93aW5kb3dzQXV0b3BpbG90RGVwbG95bWVudFByb2Zp
bGVzIg0KICAgIA0KICAgICAgICAgICAgJGZ1bGxfYXNzaWdubWVudF9pZCA9ICRpZCArICJfIiArICRncm91cGlkICsgIl8wIg0KDQogICAgICAgICAgICAk
dXJpID0gImh0dHBzOi8vZ3JhcGgubWljcm9zb2Z0LmNvbS8kZ3JhcGhBcGlWZXJzaW9uLyRSZXNvdXJjZS8kaWQvYXNzaWdubWVudHMvJGZ1bGxfYXNzaWdu
bWVudF9pZCINCg0KICAgICAgICAgICAgV3JpdGUtVmVyYm9zZSAiREVMRVRFICR1cmkiDQoNCiAgICAgICAgICAgIHRyeSB7DQogICAgICAgICAgICAgICAg
SW52b2tlLU1nR3JhcGhSZXF1ZXN0IC1VcmkgJHVyaSAtTWV0aG9kIERFTEVURQ0KICAgICAgICAgICAgfQ0KICAgICAgICAgICAgY2F0Y2ggew0KICAgICAg
ICAgICAgICAgIFdyaXRlLUVycm9yICRfLkV4Y2VwdGlvbiANCiAgICAgICAgICAgICAgICBicmVhaw0KICAgICAgICAgICAgfQ0KDQogICAgICAgIH0NCg0K
DQogICAgICAgIEZ1bmN0aW9uIFNldC1BdXRvcGlsb3RQcm9maWxlQXNzaWduZWRHcm91cCgpIHsNCiAgICAgICAgICAgIDwjDQouU1lOT1BTSVMNCkFzc2ln
bnMgYSBncm91cCB0byBhIFdpbmRvd3MgQXV0b3BpbG90IHByb2ZpbGUuDQouREVTQ1JJUFRJT04NClRoZSBTZXQtQXV0b3BpbG90UHJvZmlsZUFzc2lnbmVk
R3JvdXAgY21kbGV0IGFsbG93cyB5b3UgdG8gYXNzaWduIGEgc3BlY2lmaWMgZ3JvdXAgdG8gYSBzcGVjaWZpYyBkZXBsb3ltZW50IHByb2ZpbGUNCi5QQVJB
TUVURVIgaWQNClR5cGU6IEludGVnZXIgLSBNYW5kYXRvcnksIHRoZSBJRCAoR1VJRCkgb2YgdGhlIHByb2ZpbGUNCi5QQVJBTUVURVIgZ3JvdXBpZA0KVHlw
ZTogSW50ZWdlciAtIE1hbmRhdG9yeSwgdGhlIElEIG9mIHRoZSBncm91cA0KLkVYQU1QTEUNClNldC1BdXRvcGlsb3RQcm9maWxlQXNzaWduZWRHcm91cCAt
aWQgJGlkIC1ncm91cGlkICRncm91cGlkDQojPg0KICAgICAgICAgICAgW2NtZGxldGJpbmRpbmcoKV0NCiAgICAgICAgICAgIHBhcmFtDQogICAgICAgICAg
ICAoDQogICAgICAgICAgICAgICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkdHJ1ZSldJGlkLA0KICAgICAgICAgICAgICAgIFtQYXJhbWV0ZXIoTWFuZGF0
b3J5ID0gJHRydWUpXSRncm91cGlkDQogICAgICAgICAgICApDQogICAgICAgICAgICAkZnVsbF9hc3NpZ25tZW50X2lkID0gJGlkICsgIl8iICsgJGdyb3Vw
aWQgKyAiXzAiICANCiAgDQogICAgICAgICAgICAjIERlZmluaW5nIFZhcmlhYmxlcw0KICAgICAgICAgICAgJGdyYXBoQXBpVmVyc2lvbiA9ICJiZXRhIg0K
ICAgICAgICAgICAgJFJlc291cmNlID0gImRldmljZU1hbmFnZW1lbnQvd2luZG93c0F1dG9waWxvdERlcGxveW1lbnRQcm9maWxlcyIgICAgICAgIA0KICAg
ICAgICAgICAgJHVyaSA9ICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vJGdyYXBoQXBpVmVyc2lvbi8kUmVzb3VyY2UvJGlkL2Fzc2lnbm1lbnRzIiAg
ICAgICAgDQoNCiAgICAgICAgICAgICRqc29uID0gQCINCnsNCiAgICAiaWQiOiAiJGZ1bGxfYXNzaWdubWVudF9pZCIsDQogICAgInRhcmdldCI6IHsNCiAg
ICAgICAgIkBvZGF0YS50eXBlIjogIiNtaWNyb3NvZnQuZ3JhcGguZ3JvdXBBc3NpZ25tZW50VGFyZ2V0IiwNCiAgICAgICAgImdyb3VwSWQiOiAiJGdyb3Vw
aWQiDQogICAgfQ0KfQ0KIkANCg0KICAgICAgICAgICAgV3JpdGUtVmVyYm9zZSAiUE9TVCAkdXJpYG4kanNvbiINCg0KICAgICAgICAgICAgdHJ5IHsNCiAg
ICAgICAgICAgICAgICBJbnZva2UtTWdHcmFwaFJlcXVlc3QgLVVyaSAkdXJpIC1NZXRob2QgUG9zdCAtQm9keSAkanNvbiAtQ29udGVudFR5cGUgImFwcGxp
Y2F0aW9uL2pzb24iIC1PdXRwdXRUeXBlIFBTT2JqZWN0DQogICAgICAgICAgICB9DQogICAgICAgICAgICBjYXRjaCB7DQogICAgICAgICAgICAgICAgV3Jp
dGUtRXJyb3IgJF8uRXhjZXB0aW9uIA0KICAgICAgICAgICAgICAgIGJyZWFrDQogICAgICAgICAgICB9DQogICAgICAgIH0NCg0KDQogICAgICAgIEZ1bmN0
aW9uIEdldC1FbnJvbGxtZW50U3RhdHVzUGFnZSgpIHsNCiAgICAgICAgICAgIDwjDQouU1lOT1BTSVMNCkxpc3QgZW5yb2xsbWVudCBzdGF0dXMgcGFnZQ0K
LkRFU0NSSVBUSU9ODQpUaGUgR2V0LUVucm9sbG1lbnRTdGF0dXNQYWdlIGNtZGxldCByZXR1cm5zIGF2YWlsYWJsZSBlbnJvbGxtZW50IHN0YXR1cyBwYWdl
IHdpdGggdGhlaXIgb3B0aW9ucw0KLlBBUkFNRVRFUiBpZA0KVGhlIElEIChHVUlEKSBvZiB0aGUgc3RhdHVzIHBhZ2UgKG9wdGlvbmFsKQ0KLkVYQU1QTEUN
CkdldC1FbnJvbGxtZW50U3RhdHVzUGFnZQ0KIz4NCg0KICAgICAgICAgICAgW2NtZGxldGJpbmRpbmcoKV0NCiAgICAgICAgICAgIHBhcmFtDQogICAgICAg
ICAgICAoDQogICAgICAgICAgICAgICAgW1BhcmFtZXRlcigpXSAkaWQNCiAgICAgICAgICAgICkNCg0KICAgICAgICAgICAgIyBEZWZpbmluZyBWYXJpYWJs
ZXMNCiAgICAgICAgICAgICRncmFwaEFwaVZlcnNpb24gPSAiYmV0YSINCiAgICAgICAgICAgICRSZXNvdXJjZSA9ICJkZXZpY2VNYW5hZ2VtZW50L2Rldmlj
ZUVucm9sbG1lbnRDb25maWd1cmF0aW9ucyINCg0KICAgICAgICAgICAgaWYgKCRpZCkgew0KICAgICAgICAgICAgICAgICR1cmkgPSAiaHR0cHM6Ly9ncmFw
aC5taWNyb3NvZnQuY29tLyRncmFwaEFwaVZlcnNpb24vJFJlc291cmNlLyRpZCINCiAgICAgICAgICAgIH0NCiAgICAgICAgICAgIGVsc2Ugew0KICAgICAg
ICAgICAgICAgICR1cmkgPSAiaHR0cHM6Ly9ncmFwaC5taWNyb3NvZnQuY29tLyRncmFwaEFwaVZlcnNpb24vJFJlc291cmNlIg0KICAgICAgICAgICAgfQ0K
DQogICAgICAgICAgICBXcml0ZS1WZXJib3NlICJHRVQgJHVyaSINCg0KICAgICAgICAgICAgdHJ5IHsNCiAgICAgICAgICAgICAgICAkcmVzcG9uc2UgPSBJ
bnZva2UtTWdHcmFwaFJlcXVlc3QgLVVyaSAkdXJpIC1NZXRob2QgR2V0IC1PdXRwdXRUeXBlIFBTT2JqZWN0DQogICAgICAgICAgICAgICAgaWYgKCRpZCkg
ew0KICAgICAgICAgICAgICAgICAgICAkcmVzcG9uc2UNCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgZWxzZSB7DQogICAgICAgICAgICAg
ICAgICAgICRyZXNwb25zZS5WYWx1ZSB8IFdoZXJlLU9iamVjdCB7ICRfLidAb2RhdGEudHlwZScgLWVxICIjbWljcm9zb2Z0LmdyYXBoLndpbmRvd3MxMEVu
cm9sbG1lbnRDb21wbGV0aW9uUGFnZUNvbmZpZ3VyYXRpb24iIH0NCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICB9DQogICAgICAgICAgICBjYXRj
aCB7DQogICAgICAgICAgICAgICAgV3JpdGUtRXJyb3IgJF8uRXhjZXB0aW9uIA0KICAgICAgICAgICAgICAgIGJyZWFrDQogICAgICAgICAgICB9DQoNCiAg
ICAgICAgfQ0KDQoNCiAgICAgICAgRnVuY3Rpb24gQWRkLUVucm9sbG1lbnRTdGF0dXNQYWdlKCkgew0KICAgICAgICAgICAgPCMNCi5TWU5PUFNJUw0KQWRk
cyBhIG5ldyBXaW5kb3dzIEF1dG9waWxvdCBFbnJvbGxtZW50IFN0YXR1cyBQYWdlLg0KLkRFU0NSSVBUSU9ODQpUaGUgQWRkLUVucm9sbG1lbnRTdGF0dXNQ
YWdlIGNtZGxldCBzZXRzIHByb3BlcnRpZXMgb24gYW4gZXhpc3RpbmcgQXV0b3BpbG90IHByb2ZpbGUuDQouUEFSQU1FVEVSIERpc3BsYXlOYW1lDQpUeXBl
OiBTdHJpbmcgLSBDb25maWd1cmUgdGhlIGRpc3BsYXkgbmFtZSBvZiB0aGUgZW5yb2xsbWVudCBzdGF0dXMgcGFnZQ0KLlBBUkFNRVRFUiBkZXNjcmlwdGlv
bg0KVHlwZTogU3RyaW5nIC0gQ29uZmlndXJlIHRoZSBkZXNjcmlwdGlvbiBvZiB0aGUgZW5yb2xsbWVudCBzdGF0dXMgcGFnZQ0KLlBBUkFNRVRFUiBIaWRl
UHJvZ3Jlc3MNClR5cGU6IEJvb2xlYW4gLSBDb25maWd1cmUgdGhlIG9wdGlvbjogU2hvdyBhcHAgYW5kIHByb2ZpbGUgaW5zdGFsbGF0aW9uIHByb2dyZXNz
DQouUEFSQU1FVEVSIEFsbG93Q29sbGVjdExvZ3MNClR5cGU6IEJvb2xlYW4gLSBDb25maWd1cmUgdGhlIG9wdGlvbjogQWxsb3cgdXNlcnMgdG8gY29sbGVj
dCBsb2dzIGFib3V0IGluc3RhbGxhdGlvbiBlcnJvcnMNCi5QQVJBTUVURVIgTWVzc2FnZQ0KVHlwZTogU3RyaW5nIC0gQ29uZmlndXJlIHRoZSBvcHRpb246
IFNob3cgY3VzdG9tIG1lc3NhZ2Ugd2hlbiBhbiBlcnJvciBvY2N1cnMNCi5QQVJBTUVURVIgQWxsb3dVc2VPbkZhaWx1cmUNClR5cGU6IEJvb2xlYW4gLSBD
b25maWd1cmUgdGhlIG9wdGlvbjogQWxsb3cgdXNlcnMgdG8gdXNlIGRldmljZSBpZiBpbnN0YWxsYXRpb24gZXJyb3Igb2NjdXJzDQouUEFSQU1FVEVSIEFs
bG93UmVzZXRPbkVycm9yDQpUeXBlOiBCb29sZWFuIC0gQ29uZmlndXJlIHRoZSBvcHRpb246IEFsbG93IHVzZXJzIHRvIHJlc2V0IGRldmljZSBpZiBpbnN0
YWxsYXRpb24gZXJyb3Igb2NjdXJzDQouUEFSQU1FVEVSIEJsb2NrRGV2aWNlVW50aWxDb21wbGV0ZQ0KVHlwZTogQm9vbGVhbiAtIENvbmZpZ3VyZSB0aGUg
b3B0aW9uOiBCbG9jayBkZXZpY2UgdXNlIHVudGlsIGFsbCBhcHBzIGFuZCBwcm9maWxlcyBhcmUgaW5zdGFsbGVkDQouUEFSQU1FVEVSIFRpbWVvdXRJbk1p
bnV0ZXMNClR5cGU6IEludGVnZXIgLSBDb25maWd1cmUgdGhlIG9wdGlvbjogU2hvdyBlcnJvciB3aGVuIGluc3RhbGxhdGlvbiB0YWtlcyBsb25nZXIgdGhh
biBzcGVjaWZpZWQgbnVtYmVyIG9mIG1pbnV0ZXMNCi5FWEFNUExFDQpBZGQtRW5yb2xsbWVudFN0YXR1c1BhZ2UgLU1lc3NhZ2UgIk9vcHMgYW4gZXJyb3Ig
b2NjdXJlZCwgcGxlYXNlIGNvbnRhY3QgeW91ciBzdXBwb3J0IiAtSGlkZVByb2dyZXNzICRUcnVlIC1BbGxvd1Jlc2V0T25FcnJvciAkVHJ1ZQ0KIz4NCiAg
ICAgICAgICAgIFtjbWRsZXRiaW5kaW5nKCldDQogICAgICAgICAgICBwYXJhbQ0KICAgICAgICAgICAgKA0KICAgICAgICAgICAgICAgIFtQYXJhbWV0ZXIo
TWFuZGF0b3J5ID0gJFRydWUpXVtzdHJpbmddJERpc3BsYXlOYW1lLA0KICAgICAgICAgICAgICAgIFtzdHJpbmddJERlc2NyaXB0aW9uLCAgICAgICAgDQog
ICAgICAgICAgICAgICAgW2Jvb2xdJEhpZGVQcm9ncmVzcywgICAgDQogICAgICAgICAgICAgICAgW2Jvb2xdJEFsbG93Q29sbGVjdExvZ3MsDQogICAgICAg
ICAgICAgICAgW2Jvb2xdJGJsb2NrRGV2aWNlU2V0dXBSZXRyeUJ5VXNlciwgICAgDQogICAgICAgICAgICAgICAgW3N0cmluZ10kTWVzc2FnZSwgICAgDQog
ICAgICAgICAgICAgICAgW2Jvb2xdJEFsbG93VXNlT25GYWlsdXJlLA0KICAgICAgICAgICAgICAgIFtib29sXSRBbGxvd1Jlc2V0T25FcnJvciwgICAgDQog
ICAgICAgICAgICAgICAgW2Jvb2xdJEJsb2NrRGV2aWNlVW50aWxDb21wbGV0ZSwgICAgICAgICAgICAgICAgDQogICAgICAgICAgICAgICAgW0ludF0kVGlt
ZW91dEluTWludXRlcyAgICAgICAgDQogICAgICAgICAgICApDQoNCiAgICAgICAgICAgIElmICgkSGlkZVByb2dyZXNzIC1lcSAkRmFsc2UpIHsNCiAgICAg
ICAgICAgICAgICAkYmxvY2tEZXZpY2VTZXR1cFJldHJ5QnlVc2VyID0gJHRydWUNCiAgICAgICAgICAgIH0NCg0KICAgICAgICAgICAgSWYgKCgkRGVzY3Jp
cHRpb24gLWVxICRudWxsKSkgew0KICAgICAgICAgICAgICAgICREZXNjcmlwdGlvbiA9ICRFbnJvbGxtZW50UGFnZV9EZXNjcmlwdGlvbg0KICAgICAgICAg
ICAgfSAgICAgICAgDQoNCiAgICAgICAgICAgIElmICgoJERpc3BsYXlOYW1lIC1lcSAkbnVsbCkpIHsNCiAgICAgICAgICAgICAgICAkRGlzcGxheU5hbWUg
PSAiIg0KICAgICAgICAgICAgfSAgICANCg0KICAgICAgICAgICAgSWYgKCgkVGltZW91dEluTWludXRlcyAtZXEgIiIpKSB7DQogICAgICAgICAgICAgICAg
JFRpbWVvdXRJbk1pbnV0ZXMgPSAiNjAiDQogICAgICAgICAgICB9ICAgICAgICAgICAgICAgIA0KDQogICAgICAgICAgICAjIERlZmluaW5nIFZhcmlhYmxl
cw0KICAgICAgICAgICAgJGdyYXBoQXBpVmVyc2lvbiA9ICJiZXRhIg0KICAgICAgICAgICAgJFJlc291cmNlID0gImRldmljZU1hbmFnZW1lbnQvZGV2aWNl
RW5yb2xsbWVudENvbmZpZ3VyYXRpb25zIg0KICAgICAgICAgICAgJHVyaSA9ICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vJGdyYXBoQXBpVmVyc2lv
bi8kUmVzb3VyY2UiDQogICAgICAgICAgICAkanNvbiA9IEAiDQp7DQogICAgIkBvZGF0YS50eXBlIjogIiNtaWNyb3NvZnQuZ3JhcGgud2luZG93czEwRW5y
b2xsbWVudENvbXBsZXRpb25QYWdlQ29uZmlndXJhdGlvbiIsDQogICAgImRpc3BsYXlOYW1lIjogIiREaXNwbGF5TmFtZSIsDQogICAgImRlc2NyaXB0aW9u
IjogIiRkZXNjcmlwdGlvbiIsDQogICAgInNob3dJbnN0YWxsYXRpb25Qcm9ncmVzcyI6ICIkaGlkZXByb2dyZXNzIiwNCiAgICAiYmxvY2tEZXZpY2VTZXR1
cFJldHJ5QnlVc2VyIjogIiRibG9ja0RldmljZVNldHVwUmV0cnlCeVVzZXIiLA0KICAgICJhbGxvd0RldmljZVJlc2V0T25JbnN0YWxsRmFpbHVyZSI6ICIk
QWxsb3dSZXNldE9uRXJyb3IiLA0KICAgICJhbGxvd0xvZ0NvbGxlY3Rpb25Pbkluc3RhbGxGYWlsdXJlIjogIiRBbGxvd0NvbGxlY3RMb2dzIiwNCiAgICAi
Y3VzdG9tRXJyb3JNZXNzYWdlIjogIiRNZXNzYWdlIiwNCiAgICAiaW5zdGFsbFByb2dyZXNzVGltZW91dEluTWludXRlcyI6ICIkVGltZW91dEluTWludXRl
cyIsDQogICAgImFsbG93RGV2aWNlVXNlT25JbnN0YWxsRmFpbHVyZSI6ICIkQWxsb3dVc2VPbkZhaWx1cmUiLA0KfQ0KIkANCg0KICAgICAgICAgICAgV3Jp
dGUtVmVyYm9zZSAiUE9TVCAkdXJpYG4kanNvbiINCg0KICAgICAgICAgICAgdHJ5IHsNCiAgICAgICAgICAgICAgICBJbnZva2UtTWdHcmFwaFJlcXVlc3Qg
LVVyaSAkdXJpIC1NZXRob2QgUG9zdCAtQm9keSAkanNvbiAtQ29udGVudFR5cGUgImFwcGxpY2F0aW9uL2pzb24iIC1PdXRwdXRUeXBlIFBTT2JqZWN0DQog
ICAgICAgICAgICB9DQogICAgICAgICAgICBjYXRjaCB7DQogICAgICAgICAgICAgICAgV3JpdGUtRXJyb3IgJF8uRXhjZXB0aW9uIA0KICAgICAgICAgICAg
ICAgIGJyZWFrDQogICAgICAgICAgICB9DQoNCiAgICAgICAgfQ0KDQoNCiAgICAgICAgRnVuY3Rpb24gU2V0LUVucm9sbG1lbnRTdGF0dXNQYWdlKCkgew0K
ICAgICAgICAgICAgPCMNCi5TWU5PUFNJUw0KU2V0cyBXaW5kb3dzIEF1dG9waWxvdCBFbnJvbGxtZW50IFN0YXR1cyBQYWdlIHByb3BlcnRpZXMuDQouREVT
Q1JJUFRJT04NClRoZSBTZXQtRW5yb2xsbWVudFN0YXR1c1BhZ2UgY21kbGV0IHNldHMgcHJvcGVydGllcyBvbiBhbiBleGlzdGluZyBBdXRvcGlsb3QgcHJv
ZmlsZS4NCi5QQVJBTUVURVIgaWQNClRoZSBJRCAoR1VJRCkgb2YgdGhlIHByb2ZpbGUgdG8gYmUgdXBkYXRlZC4NCi5QQVJBTUVURVIgRGlzcGxheU5hbWUN
ClR5cGU6IFN0cmluZyAtIENvbmZpZ3VyZSB0aGUgZGlzcGxheSBuYW1lIG9mIHRoZSBlbnJvbGxtZW50IHN0YXR1cyBwYWdlDQouUEFSQU1FVEVSIGRlc2Ny
aXB0aW9uDQpUeXBlOiBTdHJpbmcgLSBDb25maWd1cmUgdGhlIGRlc2NyaXB0aW9uIG9mIHRoZSBlbnJvbGxtZW50IHN0YXR1cyBwYWdlDQouUEFSQU1FVEVS
IEhpZGVQcm9ncmVzcw0KVHlwZTogQm9vbGVhbiAtIENvbmZpZ3VyZSB0aGUgb3B0aW9uOiBTaG93IGFwcCBhbmQgcHJvZmlsZSBpbnN0YWxsYXRpb24gcHJv
Z3Jlc3MNCi5QQVJBTUVURVIgQWxsb3dDb2xsZWN0TG9ncw0KVHlwZTogQm9vbGVhbiAtIENvbmZpZ3VyZSB0aGUgb3B0aW9uOiBBbGxvdyB1c2VycyB0byBj
b2xsZWN0IGxvZ3MgYWJvdXQgaW5zdGFsbGF0aW9uIGVycm9ycw0KLlBBUkFNRVRFUiBNZXNzYWdlDQpUeXBlOiBTdHJpbmcgLSBDb25maWd1cmUgdGhlIG9w
dGlvbjogU2hvdyBjdXN0b20gbWVzc2FnZSB3aGVuIGFuIGVycm9yIG9jY3Vycw0KLlBBUkFNRVRFUiBBbGxvd1VzZU9uRmFpbHVyZQ0KVHlwZTogQm9vbGVh
biAtIENvbmZpZ3VyZSB0aGUgb3B0aW9uOiBBbGxvdyB1c2VycyB0byB1c2UgZGV2aWNlIGlmIGluc3RhbGxhdGlvbiBlcnJvciBvY2N1cnMNCi5QQVJBTUVU
RVIgQWxsb3dSZXNldE9uRXJyb3INClR5cGU6IEJvb2xlYW4gLSBDb25maWd1cmUgdGhlIG9wdGlvbjogQWxsb3cgdXNlcnMgdG8gcmVzZXQgZGV2aWNlIGlm
IGluc3RhbGxhdGlvbiBlcnJvciBvY2N1cnMNCi5QQVJBTUVURVIgQmxvY2tEZXZpY2VVbnRpbENvbXBsZXRlDQpUeXBlOiBCb29sZWFuIC0gQ29uZmlndXJl
IHRoZSBvcHRpb246IEJsb2NrIGRldmljZSB1c2UgdW50aWwgYWxsIGFwcHMgYW5kIHByb2ZpbGVzIGFyZSBpbnN0YWxsZWQNCi5QQVJBTUVURVIgVGltZW91
dEluTWludXRlcw0KVHlwZTogSW50ZWdlciAtIENvbmZpZ3VyZSB0aGUgb3B0aW9uOiBTaG93IGVycm9yIHdoZW4gaW5zdGFsbGF0aW9uIHRha2VzIGxvbmdl
ciB0aGFuIHNwZWNpZmllZCBudW1iZXIgb2YgbWludXRlcw0KLkVYQU1QTEUNClNldC1FbnJvbGxtZW50U3RhdHVzUGFnZSAtaWQgJGlkIC1NZXNzYWdlICJP
b3BzIGFuIGVycm9yIG9jY3VyZWQsIHBsZWFzZSBjb250YWN0IHlvdXIgc3VwcG9ydCIgLUhpZGVQcm9ncmVzcyAkVHJ1ZSAtQWxsb3dSZXNldE9uRXJyb3Ig
JFRydWUNCiM+DQogICAgICAgICAgICBbY21kbGV0YmluZGluZygpXQ0KICAgICAgICAgICAgcGFyYW0NCiAgICAgICAgICAgICgNCiAgICAgICAgICAgICAg
ICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICR0cnVlLCBWYWx1ZUZyb21QaXBlbGluZUJ5UHJvcGVydHlOYW1lID0gJFRydWUpXSAkaWQsDQogICAgICAgICAg
ICAgICAgW3N0cmluZ10kRGlzcGxheU5hbWUsICAgIA0KICAgICAgICAgICAgICAgIFtzdHJpbmddJERlc2NyaXB0aW9uLCAgICAgICAgDQogICAgICAgICAg
ICAgICAgW2Jvb2xdJEhpZGVQcm9ncmVzcywNCiAgICAgICAgICAgICAgICBbYm9vbF0kQWxsb3dDb2xsZWN0TG9ncywNCiAgICAgICAgICAgICAgICBbc3Ry
aW5nXSRNZXNzYWdlLCAgICANCiAgICAgICAgICAgICAgICBbYm9vbF0kQWxsb3dVc2VPbkZhaWx1cmUsDQogICAgICAgICAgICAgICAgW2Jvb2xdJEFsbG93
UmVzZXRPbkVycm9yLCAgICANCiAgICAgICAgICAgICAgICBbYm9vbF0kQWxsb3dVc2VPbkVycm9yLCAgICANCiAgICAgICAgICAgICAgICBbYm9vbF0kQmxv
Y2tEZXZpY2VVbnRpbENvbXBsZXRlLCAgICAgICAgICAgICAgICANCiAgICAgICAgICAgICAgICBbSW50XSRUaW1lb3V0SW5NaW51dGVzICAgICAgICANCiAg
ICAgICAgICAgICkNCg0KICAgICAgICAgICAgUHJvY2VzcyB7DQoNCiAgICAgICAgICAgICAgICAjIExJU1QgRVhJU1RJTkcgVkFMVUVTIEZPUiBUSEUgU0VM
RUNUSU5HIFNUQVVTIFBBR0UNCiAgICAgICAgICAgICAgICAjIERlZmF1bHQgcHJvZmlsZSB2YWx1ZXMNCiAgICAgICAgICAgICAgICAkRW5yb2xsbWVudFBh
Z2VfVmFsdWVzID0gR2V0LUVucm9sbG1lbnRTdGF0dXNQYWdlIC1pZCAkaWQNCiAgICAgICAgICAgICAgICAkRW5yb2xsbWVudFBhZ2VfRGlzcGxheU5hbWUg
PSAkRW5yb2xsbWVudFBhZ2VfVmFsdWVzLmRpc3BsYXlOYW1lDQogICAgICAgICAgICAgICAgJEVucm9sbG1lbnRQYWdlX0Rlc2NyaXB0aW9uID0gJEVucm9s
bG1lbnRQYWdlX1ZhbHVlcy5kZXNjcmlwdGlvbg0KICAgICAgICAgICAgICAgICRFbnJvbGxtZW50UGFnZV9zaG93SW5zdGFsbGF0aW9uUHJvZ3Jlc3MgPSAk
RW5yb2xsbWVudFBhZ2VfVmFsdWVzLnNob3dJbnN0YWxsYXRpb25Qcm9ncmVzcw0KICAgICAgICAgICAgICAgICRFbnJvbGxtZW50UGFnZV9ibG9ja0Rldmlj
ZVNldHVwUmV0cnlCeVVzZXIgPSAkRW5yb2xsbWVudFBhZ2VfVmFsdWVzLmJsb2NrRGV2aWNlU2V0dXBSZXRyeUJ5VXNlcg0KICAgICAgICAgICAgICAgICRF
bnJvbGxtZW50UGFnZV9hbGxvd0RldmljZVJlc2V0T25JbnN0YWxsRmFpbHVyZSA9ICRFbnJvbGxtZW50UGFnZV9WYWx1ZXMuYWxsb3dEZXZpY2VSZXNldE9u
SW5zdGFsbEZhaWx1cmUNCiAgICAgICAgICAgICAgICAkRW5yb2xsbWVudFBhZ2VfYWxsb3dMb2dDb2xsZWN0aW9uT25JbnN0YWxsRmFpbHVyZSA9ICRFbnJv
bGxtZW50UGFnZV9WYWx1ZXMuYWxsb3dMb2dDb2xsZWN0aW9uT25JbnN0YWxsRmFpbHVyZQ0KICAgICAgICAgICAgICAgICRFbnJvbGxtZW50UGFnZV9jdXN0
b21FcnJvck1lc3NhZ2UgPSAkRW5yb2xsbWVudFBhZ2VfVmFsdWVzLmN1c3RvbUVycm9yTWVzc2FnZQ0KICAgICAgICAgICAgICAgICRFbnJvbGxtZW50UGFn
ZV9pbnN0YWxsUHJvZ3Jlc3NUaW1lb3V0SW5NaW51dGVzID0gJEVucm9sbG1lbnRQYWdlX1ZhbHVlcy5pbnN0YWxsUHJvZ3Jlc3NUaW1lb3V0SW5NaW51dGVz
DQogICAgICAgICAgICAgICAgJEVucm9sbG1lbnRQYWdlX2FsbG93RGV2aWNlVXNlT25JbnN0YWxsRmFpbHVyZSA9ICRFbnJvbGxtZW50UGFnZV9WYWx1ZXMu
YWxsb3dEZXZpY2VVc2VPbkluc3RhbGxGYWlsdXJlDQoNCiAgICAgICAgICAgICAgICBJZiAoISgkSGlkZVByb2dyZXNzKSkgew0KICAgICAgICAgICAgICAg
ICAgICAkSGlkZVByb2dyZXNzID0gJEVucm9sbG1lbnRQYWdlX3Nob3dJbnN0YWxsYXRpb25Qcm9ncmVzcw0KICAgICAgICAgICAgICAgIH0gICAgDQogICAg
DQogICAgICAgICAgICAgICAgSWYgKCEoJEJsb2NrRGV2aWNlVW50aWxDb21wbGV0ZSkpIHsNCiAgICAgICAgICAgICAgICAgICAgJEJsb2NrRGV2aWNlVW50
aWxDb21wbGV0ZSA9ICRFbnJvbGxtZW50UGFnZV9ibG9ja0RldmljZVNldHVwUmV0cnlCeVVzZXINCiAgICAgICAgICAgICAgICB9ICAgICAgICANCiAgICAg
ICAgDQogICAgICAgICAgICAgICAgSWYgKCEoJEFsbG93Q29sbGVjdExvZ3MpKSB7DQogICAgICAgICAgICAgICAgICAgICRBbGxvd0NvbGxlY3RMb2dzID0g
JEVucm9sbG1lbnRQYWdlX2FsbG93TG9nQ29sbGVjdGlvbk9uSW5zdGFsbEZhaWx1cmUNCiAgICAgICAgICAgICAgICB9ICAgICAgICAgICAgDQogICAgDQog
ICAgICAgICAgICAgICAgSWYgKCEoJEFsbG93VXNlT25GYWlsdXJlKSkgew0KICAgICAgICAgICAgICAgICAgICAkQWxsb3dVc2VPbkZhaWx1cmUgPSAkRW5y
b2xsbWVudFBhZ2VfYWxsb3dEZXZpY2VVc2VPbkluc3RhbGxGYWlsdXJlDQogICAgICAgICAgICAgICAgfSAgICANCg0KICAgICAgICAgICAgICAgIElmICgo
JE1lc3NhZ2UgLWVxICIiKSkgew0KICAgICAgICAgICAgICAgICAgICAkTWVzc2FnZSA9ICRFbnJvbGxtZW50UGFnZV9jdXN0b21FcnJvck1lc3NhZ2UNCiAg
ICAgICAgICAgICAgICB9ICAgICAgICANCiAgICAgICAgDQogICAgICAgICAgICAgICAgSWYgKCgkRGVzY3JpcHRpb24gLWVxICRudWxsKSkgew0KICAgICAg
ICAgICAgICAgICAgICAkRGVzY3JpcHRpb24gPSAkRW5yb2xsbWVudFBhZ2VfRGVzY3JpcHRpb24NCiAgICAgICAgICAgICAgICB9ICAgICAgICANCg0KICAg
ICAgICAgICAgICAgIElmICgoJERpc3BsYXlOYW1lIC1lcSAkbnVsbCkpIHsNCiAgICAgICAgICAgICAgICAgICAgJERpc3BsYXlOYW1lID0gJEVucm9sbG1l
bnRQYWdlX0Rpc3BsYXlOYW1lDQogICAgICAgICAgICAgICAgfSAgICANCg0KICAgICAgICAgICAgICAgIElmICghKCRBbGxvd1Jlc2V0T25FcnJvcikpIHsN
CiAgICAgICAgICAgICAgICAgICAgJEFsbG93UmVzZXRPbkVycm9yID0gJEVucm9sbG1lbnRQYWdlX2FsbG93RGV2aWNlUmVzZXRPbkluc3RhbGxGYWlsdXJl
DQogICAgICAgICAgICAgICAgfSAgICANCg0KICAgICAgICAgICAgICAgIElmICgoJFRpbWVvdXRJbk1pbnV0ZXMgLWVxICIiKSkgew0KICAgICAgICAgICAg
ICAgICAgICAkVGltZW91dEluTWludXRlcyA9ICRFbnJvbGxtZW50UGFnZV9pbnN0YWxsUHJvZ3Jlc3NUaW1lb3V0SW5NaW51dGVzDQogICAgICAgICAgICAg
ICAgfSAgICAgICAgICAgICAgICANCg0KICAgICAgICAgICAgICAgICMgRGVmaW5pbmcgVmFyaWFibGVzDQogICAgICAgICAgICAgICAgJGdyYXBoQXBpVmVy
c2lvbiA9ICJiZXRhIg0KICAgICAgICAgICAgICAgICRSZXNvdXJjZSA9ICJkZXZpY2VNYW5hZ2VtZW50L2RldmljZUVucm9sbG1lbnRDb25maWd1cmF0aW9u
cyINCiAgICAgICAgICAgICAgICAkdXJpID0gImh0dHBzOi8vZ3JhcGgubWljcm9zb2Z0LmNvbS8kZ3JhcGhBcGlWZXJzaW9uLyRSZXNvdXJjZS8kaWQiDQog
ICAgICAgICAgICAgICAgJGpzb24gPSBAIg0Kew0KICAgICJAb2RhdGEudHlwZSI6ICIjbWljcm9zb2Z0LmdyYXBoLndpbmRvd3MxMEVucm9sbG1lbnRDb21w
bGV0aW9uUGFnZUNvbmZpZ3VyYXRpb24iLA0KICAgICJkaXNwbGF5TmFtZSI6ICIkRGlzcGxheU5hbWUiLA0KICAgICJkZXNjcmlwdGlvbiI6ICIkZGVzY3Jp
cHRpb24iLA0KICAgICJzaG93SW5zdGFsbGF0aW9uUHJvZ3Jlc3MiOiAiJEhpZGVQcm9ncmVzcyIsDQogICAgImJsb2NrRGV2aWNlU2V0dXBSZXRyeUJ5VXNl
ciI6ICIkQmxvY2tEZXZpY2VVbnRpbENvbXBsZXRlIiwNCiAgICAiYWxsb3dEZXZpY2VSZXNldE9uSW5zdGFsbEZhaWx1cmUiOiAiJEFsbG93UmVzZXRPbkVy
cm9yIiwNCiAgICAiYWxsb3dMb2dDb2xsZWN0aW9uT25JbnN0YWxsRmFpbHVyZSI6ICIkQWxsb3dDb2xsZWN0TG9ncyIsDQogICAgImN1c3RvbUVycm9yTWVz
c2FnZSI6ICIkTWVzc2FnZSIsDQogICAgImluc3RhbGxQcm9ncmVzc1RpbWVvdXRJbk1pbnV0ZXMiOiAiJFRpbWVvdXRJbk1pbnV0ZXMiLA0KICAgICJhbGxv
d0RldmljZVVzZU9uSW5zdGFsbEZhaWx1cmUiOiAiJEFsbG93VXNlT25GYWlsdXJlIg0KfQ0KIkANCg0KICAgICAgICAgICAgICAgIFdyaXRlLVZlcmJvc2Ug
IlBBVENIICR1cmlgbiRqc29uIg0KDQogICAgICAgICAgICAgICAgdHJ5IHsNCiAgICAgICAgICAgICAgICAgICAgSW52b2tlLU1nR3JhcGhSZXF1ZXN0IC1V
cmkgJHVyaSAtTWV0aG9kIFBBVENIIC1Cb2R5ICRqc29uIC1Db250ZW50VHlwZSAiYXBwbGljYXRpb24vanNvbiIgLU91dHB1dFR5cGUgUFNPYmplY3QNCiAg
ICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgY2F0Y2ggew0KICAgICAgICAgICAgICAgICAgICBXcml0ZS1FcnJvciAkXy5FeGNlcHRpb24gDQog
ICAgICAgICAgICAgICAgICAgIGJyZWFrDQogICAgICAgICAgICAgICAgfQ0KDQogICAgICAgICAgICB9DQoNCiAgICAgICAgfQ0KDQoNCiAgICAgICAgRnVu
Y3Rpb24gUmVtb3ZlLUVucm9sbG1lbnRTdGF0dXNQYWdlKCkgew0KICAgICAgICAgICAgPCMNCi5TWU5PUFNJUw0KUmVtb3ZlIGEgc3BlY2lmaWMgZW5yb2xs
bWVudCBzdGF0dXMgcGFnZQ0KLkRFU0NSSVBUSU9ODQpUaGUgUmVtb3ZlLUVucm9sbG1lbnRTdGF0dXNQYWdlIGFsbG93cyB5b3UgdG8gcmVtb3ZlIGEgc3Bl
Y2lmaWMgZW5yb2xsbWVudCBzdGF0dXMgcGFnZQ0KLlBBUkFNRVRFUiBpZA0KTWFuZGF0b3J5LCB0aGUgSUQgKEdVSUQpIG9mIHRoZSBwcm9maWxlIHRvIGJl
IHJldHJpZXZlZC4NCi5FWEFNUExFDQpSZW1vdmUtRW5yb2xsbWVudFN0YXR1c1BhZ2UgLWlkICRpZA0KIz4NCiAgICAgICAgICAgIFtjbWRsZXRiaW5kaW5n
KCldDQogICAgICAgICAgICBwYXJhbQ0KICAgICAgICAgICAgKA0KICAgICAgICAgICAgICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJFRydWUsIFZhbHVl
RnJvbVBpcGVsaW5lQnlQcm9wZXJ0eU5hbWUgPSAkVHJ1ZSldICRpZA0KICAgICAgICAgICAgKQ0KDQogICAgICAgICAgICBQcm9jZXNzIHsNCg0KICAgICAg
ICAgICAgICAgICMgRGVmaW5pbmcgVmFyaWFibGVzDQogICAgICAgICAgICAgICAgJGdyYXBoQXBpVmVyc2lvbiA9ICJiZXRhIg0KICAgICAgICAgICAgICAg
ICRSZXNvdXJjZSA9ICJkZXZpY2VNYW5hZ2VtZW50L2RldmljZUVucm9sbG1lbnRDb25maWd1cmF0aW9ucyINCiAgICAgICAgICAgICAgICAkdXJpID0gImh0
dHBzOi8vZ3JhcGgubWljcm9zb2Z0LmNvbS8kZ3JhcGhBcGlWZXJzaW9uLyRSZXNvdXJjZS8kaWQiDQoNCiAgICAgICAgICAgICAgICBXcml0ZS1WZXJib3Nl
ICJERUxFVEUgJHVyaSINCg0KICAgICAgICAgICAgICAgIHRyeSB7DQogICAgICAgICAgICAgICAgICAgIEludm9rZS1NZ0dyYXBoUmVxdWVzdCAtVXJpICR1
cmkgLU1ldGhvZCBERUxFVEUNCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgY2F0Y2ggew0KICAgICAgICAgICAgICAgICAgICBXcml0ZS1F
cnJvciAkXy5FeGNlcHRpb24gDQogICAgICAgICAgICAgICAgICAgIGJyZWFrDQogICAgICAgICAgICAgICAgfQ0KDQogICAgICAgICAgICB9DQoNCiAgICAg
ICAgfQ0KDQoNCiAgICAgICAgRnVuY3Rpb24gSW52b2tlLUF1dG9waWxvdFN5bmMoKSB7DQogICAgICAgICAgICA8Iw0KLlNZTk9QU0lTDQpJbml0aWF0ZXMg
YSBzeW5jaHJvbml6YXRpb24gb2YgV2luZG93cyBBdXRvcGlsb3QgZGV2aWNlcyBiZXR3ZWVuIHRoZSBBdXRvcGlsb3QgZGVwbG95bWVudCBzZXJ2aWNlIGFu
ZCBJbnR1bmUuDQogDQouREVTQ1JJUFRJT04NClRoZSBJbnZva2UtQXV0b3BpbG90U3luYyBjbWRsZXQgaW5pdGlhdGVzIGEgc3luY2hyb25pemF0aW9uIGJl
dHdlZW4gdGhlIEF1dG9waWxvdCBkZXBsb3ltZW50IHNlcnZpY2UgYW5kIEludHVuZS4NClRoaXMgY2FuIGJlIGRvbmUgYWZ0ZXIgaW1wb3J0aW5nIG5ldyBk
ZXZpY2VzLCB0byBlbnN1cmUgdGhhdCB0aGV5IGFwcGVhciBpbiBJbnR1bmUgaW4gdGhlIGxpc3Qgb2YgcmVnaXN0ZXJlZA0KQXV0b3BpbG90IGRldmljZXMu
IFNlZSBodHRwczovL2RldmVsb3Blci5taWNyb3NvZnQuY29tL2VuLXVzL2dyYXBoL2RvY3MvYXBpLXJlZmVyZW5jZS9iZXRhL2FwaS9pbnR1bmVfZW5yb2xs
bWVudF93aW5kb3dzYXV0b3BpbG90c2V0dGluZ3Nfc3luYw0KZm9yIG1vcmUgaW5mb3JtYXRpb24uDQogDQouRVhBTVBMRQ0KSW5pdGlhdGUgYSBzeW5jaHJv
bml6YXRpb24uDQogDQpJbnZva2UtQXV0b3BpbG90U3luYw0KIz4NCiAgICAgICAgICAgIFtjbWRsZXRiaW5kaW5nKCldDQogICAgICAgICAgICBwYXJhbQ0K
ICAgICAgICAgICAgKA0KICAgICAgICAgICAgKQ0KICAgICAgICAgICAgIyBEZWZpbmluZyBWYXJpYWJsZXMNCiAgICAgICAgICAgICRncmFwaEFwaVZlcnNp
b24gPSAiYmV0YSINCiAgICAgICAgICAgICRSZXNvdXJjZSA9ICJkZXZpY2VNYW5hZ2VtZW50L3dpbmRvd3NBdXRvcGlsb3RTZXR0aW5ncy9zeW5jIg0KICAg
ICAgICAgICAgJHVyaSA9ICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vJGdyYXBoQXBpVmVyc2lvbi8kUmVzb3VyY2UiDQoNCiAgICAgICAgICAgIFdy
aXRlLVZlcmJvc2UgIlBPU1QgJHVyaSINCg0KICAgICAgICAgICAgdHJ5IHsNCiAgICAgICAgICAgICAgICBJbnZva2UtTWdHcmFwaFJlcXVlc3QgLVVyaSAk
dXJpIC1NZXRob2QgUG9zdA0KICAgICAgICAgICAgfQ0KICAgICAgICAgICAgY2F0Y2ggew0KICAgICAgICAgICAgICAgIFdyaXRlLUVycm9yICRfLkV4Y2Vw
dGlvbiANCiAgICAgICAgICAgICAgICBicmVhaw0KICAgICAgICAgICAgfQ0KDQogICAgICAgIH0NCg0KICAgICAgICBGdW5jdGlvbiBHZXQtQXV0b3BpbG90
U3luY0luZm8oKSB7DQogICAgICAgICAgICA8Iw0KICAgIC5TWU5PUFNJUw0KICAgIFJldHVybnMgZGV0YWlscyBhYm91dCB0aGUgbGFzdCBBdXRvcGlsb3Qg
c3luYy4NCiAgICAgDQogICAgLkRFU0NSSVBUSU9ODQogICAgVGhlIEdldC1BdXRvcGlsb3RTeW5jSW5mbyBjbWRsZXQgcmV0cmlldmVzIGRldGFpbHMgYWJv
dXQgdGhlIHN5bmMgc3RhdHVzIGJldHdlZW4gSW50dW5lIGFuZCB0aGUgQXV0b3BpbG90IHNlcnZpY2UuDQogICAgU2VlIGh0dHBzOi8vZG9jcy5taWNyb3Nv
ZnQuY29tL2VuLXVzL2dyYXBoL2FwaS9yZXNvdXJjZXMvaW50dW5lLWVucm9sbG1lbnQtd2luZG93c2F1dG9waWxvdHNldHRpbmdzP3ZpZXc9Z3JhcGgtcmVz
dC1iZXRhDQogICAgZm9yIG1vcmUgaW5mb3JtYXRpb24uDQogICAgIA0KICAgIC5FWEFNUExFDQogICAgR2V0LUF1dG9waWxvdFN5bmNJbmZvDQogICAgIz4N
CiAgICAgICAgICAgIFtjbWRsZXRiaW5kaW5nKCldDQogICAgICAgICAgICBwYXJhbQ0KICAgICAgICAgICAgKA0KICAgICAgICAgICAgKQ0KICAgICAgICAg
ICAgIyBEZWZpbmluZyBWYXJpYWJsZXMNCiAgICAgICAgICAgICRncmFwaEFwaVZlcnNpb24gPSAiYmV0YSINCiAgICAgICAgICAgICRSZXNvdXJjZSA9ICJk
ZXZpY2VNYW5hZ2VtZW50L3dpbmRvd3NBdXRvcGlsb3RTZXR0aW5ncyINCiAgICAgICAgICAgICR1cmkgPSAiaHR0cHM6Ly9ncmFwaC5taWNyb3NvZnQuY29t
LyRncmFwaEFwaVZlcnNpb24vJFJlc291cmNlIg0KICAgIA0KICAgICAgICAgICAgV3JpdGUtVmVyYm9zZSAiR0VUICR1cmkiDQogICAgDQogICAgICAgICAg
ICB0cnkgew0KICAgICAgICAgICAgICAgIEludm9rZS1NZ0dyYXBoUmVxdWVzdCAtVXJpICR1cmkgLU1ldGhvZCBHZXQgLU91dHB1dFR5cGUgUFNPYmplY3QN
CiAgICAgICAgICAgIH0NCiAgICAgICAgICAgIGNhdGNoIHsNCiAgICAgICAgICAgICAgICBXcml0ZS1FcnJvciAkXy5FeGNlcHRpb24gDQogICAgICAgICAg
ICAgICAgYnJlYWsNCiAgICAgICAgICAgIH0NCiAgICANCiAgICAgICAgfQ0KICAgIA0KICAgICAgICAjZW5kcmVnaW9uDQoNCg0KICAgICAgICBGdW5jdGlv
biBJbXBvcnQtQXV0b3BpbG90Q1NWKCkgew0KICAgICAgICAgICAgPCMNCi5TWU5PUFNJUw0KQWRkcyBhIGJhdGNoIG9mIG5ldyBkZXZpY2VzIGludG8gV2lu
ZG93cyBBdXRvcGlsb3QuDQogDQouREVTQ1JJUFRJT04NClRoZSBJbXBvcnQtQXV0b3BpbG90Q1NWIGNtZGxldCBwcm9jZXNzZXMgYSBsaXN0IG9mIG5ldyBk
ZXZpY2VzIChjb250YWluZWQgaW4gYSBDU1YgZmlsZSkgdXNpbmcgYSBzZXZlcmFsIG9mIHRoZSBvdGhlciBjbWRsZXRzIGluY2x1ZGVkIGluIHRoaXMgbW9k
dWxlLiBJdCBpcyBhIGNvbnZlbmllbnQgd3JhcHBlciB0byBoYW5kbGUgdGhlIGRldGFpbHMuIEFmdGVyIHRoZSBkZXZpY2VzIGhhdmUgYmVlbiBhZGRlZCwg
dGhlIGNtZGxldCB3aWxsIGNvbnRpbnVlIHRvIGNoZWNrIHRoZSBzdGF0dXMgb2YgdGhlIGltcG9ydCBwcm9jZXNzLiBPbmNlIGFsbCBkZXZpY2VzIGhhdmUg
YmVlbiBwcm9jZXNzZWQgKHN1Y2Nlc3NmdWxseSBvciBub3QpIHRoZSBjbWRsZXQgd2lsbCBjb21wbGV0ZS4gVGhpcyBjYW4gdGFrZSBzZXZlcmFsIG1pbnV0
ZXMsIGFzIHRoZSBkZXZpY2VzIGFyZSBwcm9jZXNzZWQgYnkgSW50dW5lIGFzIGEgYmFja2dyb3VuZCBiYXRjaCBwcm9jZXNzLg0KIA0KLlBBUkFNRVRFUiBj
c3ZGaWxlDQpUaGUgZmlsZSBjb250YWluaW5nIHRoZSBsaXN0IG9mIGRldmljZXMgdG8gYmUgYWRkZWQuDQogDQouUEFSQU1FVEVSIGdyb3VwVGFnDQpBbiBv
cHRpb25hbCBpZGVudGlmaWVyIG9yIHRhZyB0aGF0IGNhbiBiZSBhc3NvY2lhdGVkIHdpdGggdGhpcyBkZXZpY2UsIHVzZWZ1bCBmb3IgZ3JvdXBpbmcgZGV2
aWNlcyB1c2luZyBFbnRyYSBkeW5hbWljIGdyb3Vwcy4gVGhpcyB2YWx1ZSBvdmVycmlkZXMgYW4gR3JvdXAgVGFnIHZhbHVlIHNwZWNpZmllZCBpbiB0aGUg
Q1NWIGZpbGUuDQogDQouRVhBTVBMRQ0KQWRkIGEgYmF0Y2ggb2YgZGV2aWNlcyB0byBXaW5kb3dzIEF1dG9waWxvdCBmb3IgdGhlIGN1cnJlbnQgRW50cmEg
dGVuYW50Lg0KIA0KSW1wb3J0LUF1dG9waWxvdENTViAtY3N2RmlsZSBDOlxEZXZpY2VzLmNzdg0KIz4NCiAgICAgICAgICAgIFtjbWRsZXRiaW5kaW5nKCld
DQogICAgICAgICAgICBwYXJhbQ0KICAgICAgICAgICAgKA0KICAgICAgICAgICAgICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJHRydWUpXSAkY3N2Rmls
ZSwNCiAgICAgICAgICAgICAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICRmYWxzZSldIFtBbGlhcygib3JkZXJJZGVudGlmaWVyIildICRncm91cFRhZyA9
ICIiDQogICAgICAgICAgICApDQogICAgDQogICAgICAgICAgICAjIFJlYWQgQ1NWIGFuZCBwcm9jZXNzIGVhY2ggZGV2aWNlDQogICAgICAgICAgICAkZGV2
aWNlcyA9IEltcG9ydC1Dc3YgJGNzdkZpbGUNCiAgICAgICAgICAgICRpbXBvcnRlZERldmljZXMgPSBAKCkNCiAgICAgICAgICAgIGZvcmVhY2ggKCRkZXZp
Y2UgaW4gJGRldmljZXMpIHsNCiAgICAgICAgICAgICAgICBpZiAoJGdyb3VwVGFnIC1uZSAiIikgew0KICAgICAgICAgICAgICAgICAgICAkbyA9ICRncm91
cFRhZw0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBlbHNlaWYgKCRkZXZpY2UuJ0dyb3VwIFRhZycgLW5lICIiKSB7DQogICAgICAgICAg
ICAgICAgICAgICRvID0gJGRldmljZS4nR3JvdXAgVGFnJw0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBlbHNlIHsNCiAgICAgICAgICAg
ICAgICAgICAgJG8gPSAkZGV2aWNlLidPcmRlcklEJw0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBBZGQtQXV0b3BpbG90SW1wb3J0ZWRE
ZXZpY2UgLXNlcmlhbE51bWJlciAkZGV2aWNlLidEZXZpY2UgU2VyaWFsIE51bWJlcicgLWhhcmR3YXJlSWRlbnRpZmllciAkZGV2aWNlLidIYXJkd2FyZSBI
YXNoJyAtZ3JvdXBUYWcgJG8gLWFzc2lnbmVkVXNlciAkZGV2aWNlLidBc3NpZ25lZCBVc2VyJw0KICAgICAgICAgICAgfQ0KDQogICAgICAgICAgICAjIFdo
aWxlIHdlIGNvdWxkIGtlZXAgYSBsaXN0IG9mIGFsbCB0aGUgSURzIHRoYXQgd2UgYWRkZWQgYW5kIHRoZW4gY2hlY2sgZWFjaCBvbmUsIGl0IGlzDQogICAg
ICAgICAgICAjIGVhc2llciB0byBqdXN0IGxvb3AgdGhyb3VnaCBhbGwgb2YgdGhlbQ0KICAgICAgICAgICAgJHByb2Nlc3NpbmdDb3VudCA9IDENCiAgICAg
ICAgICAgIHdoaWxlICgkcHJvY2Vzc2luZ0NvdW50IC1ndCAwKSB7DQogICAgICAgICAgICAgICAgJGRldmljZVN0YXR1c2VzID0gQChHZXQtQXV0b3BpbG90
SW1wb3J0ZWREZXZpY2UpDQogICAgICAgICAgICAgICAgJGRldmljZUNvdW50ID0gJGRldmljZVN0YXR1c2VzLkxlbmd0aA0KDQogICAgICAgICAgICAgICAg
IyBDaGVjayB0byBzZWUgaWYgYW55IGRldmljZXMgYXJlIHN0aWxsIHByb2Nlc3NpbmcNCiAgICAgICAgICAgICAgICAkcHJvY2Vzc2luZ0NvdW50ID0gMA0K
ICAgICAgICAgICAgICAgIGZvcmVhY2ggKCRkZXZpY2UgaW4gJGRldmljZVN0YXR1c2VzKSB7DQogICAgICAgICAgICAgICAgICAgIGlmICgkZGV2aWNlLnN0
YXRlLmRldmljZUltcG9ydFN0YXR1cyAtZXEgInVua25vd24iKSB7DQogICAgICAgICAgICAgICAgICAgICAgICAkcHJvY2Vzc2luZ0NvdW50ID0gJHByb2Nl
c3NpbmdDb3VudCArIDENCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICJXYWl0
aW5nIGZvciAkcHJvY2Vzc2luZ0NvdW50IG9mICRkZXZpY2VDb3VudCINCg0KICAgICAgICAgICAgICAgICMgU3RpbGwgcHJvY2Vzc2luZz8gU2xlZXAgYmVm
b3JlIHRyeWluZyBhZ2Fpbi4NCiAgICAgICAgICAgICAgICBpZiAoJHByb2Nlc3NpbmdDb3VudCAtZ3QgMCkgew0KICAgICAgICAgICAgICAgICAgICBTdGFy
dC1TbGVlcCAxNQ0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgIH0NCg0KICAgICAgICAgICAgIyBEaXNwbGF5IHRoZSBzdGF0dXNlcw0KICAgICAg
ICAgICAgJGRldmljZVN0YXR1c2VzIHwgRm9yRWFjaC1PYmplY3Qgew0KICAgICAgICAgICAgICAgIFdyaXRlLUhvc3QgIlNlcmlhbCBudW1iZXIgJCgkXy5z
ZXJpYWxOdW1iZXIpOiAkKCRfLnN0YXRlLmRldmljZUltcG9ydFN0YXR1cykgJCgkXy5zdGF0ZS5kZXZpY2VFcnJvckNvZGUpICQoJF8uc3RhdGUuZGV2aWNl
RXJyb3JOYW1lKSINCiAgICAgICAgICAgIH0NCg0KICAgICAgICAgICAgIyBDbGVhbnVwIHRoZSBpbXBvcnRlZCBkZXZpY2UgcmVjb3Jkcw0KICAgICAgICAg
ICAgJGRldmljZVN0YXR1c2VzIHwgRm9yRWFjaC1PYmplY3Qgew0KICAgICAgICAgICAgICAgIFJlbW92ZS1BdXRvcGlsb3RJbXBvcnRlZERldmljZSAtaWQg
JF8uaWQNCiAgICAgICAgICAgIH0NCiAgICAgICAgfQ0KDQoNCiAgICAgICAgRnVuY3Rpb24gR2V0LUF1dG9waWxvdEV2ZW50KCkgew0KICAgICAgICAgICAg
PCMNCi5TWU5PUFNJUw0KR2V0cyBXaW5kb3dzIEF1dG9waWxvdCBkZXBsb3ltZW50IGV2ZW50cy4NCiANCi5ERVNDUklQVElPTg0KVGhlIEdldC1BdXRvcGls
b3RFdmVudCBjbWRsZXQgcmV0cmlldmVzIHRoZSBsaXN0IG9mIGRlcGxveW1lbnQgZXZlbnRzICh0aGUgZGF0YSB0aGF0IHlvdSB3b3VsZCBzZWUgaW4gdGhl
ICJBdXRvcGlsb3QgZGVwbG95bWVudHMiIHJlcG9ydCBpbiB0aGUgSW50dW5lIHBvcnRhbCkuDQogDQouRVhBTVBMRQ0KR2V0IGEgbGlzdCBvZiBhbGwgV2lu
ZG93cyBBdXRvcGlsb3QgZXZlbnRzDQogDQpHZXQtQXV0b3BpbG90RXZlbnQNCiM+DQogICAgICAgICAgICBbY21kbGV0YmluZGluZygpXQ0KICAgICAgICAg
ICAgcGFyYW0NCiAgICAgICAgICAgICgNCiAgICAgICAgICAgICkNCg0KICAgICAgICAgICAgUHJvY2VzcyB7DQoNCiAgICAgICAgICAgICAgICAjIERlZmlu
aW5nIFZhcmlhYmxlcw0KICAgICAgICAgICAgICAgICRncmFwaEFwaVZlcnNpb24gPSAiYmV0YSINCiAgICAgICAgICAgICAgICAkUmVzb3VyY2UgPSAiZGV2
aWNlTWFuYWdlbWVudC9hdXRvcGlsb3RFdmVudHMiDQogICAgICAgICAgICAgICAgJHVyaSA9ICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vJGdyYXBo
QXBpVmVyc2lvbi8kKCRSZXNvdXJjZSkiDQoNCiAgICAgICAgICAgICAgICB0cnkgew0KICAgICAgICAgICAgICAgICAgICAkcmVzcG9uc2UgPSBJbnZva2Ut
TWdHcmFwaFJlcXVlc3QgLVVyaSAkdXJpIC1NZXRob2QgR2V0IC1PdXRwdXRUeXBlIFBTT2JqZWN0DQogICAgICAgICAgICAgICAgICAgICRkZXZpY2VzID0g
JHJlc3BvbnNlLnZhbHVlDQogICAgICAgICAgICAgICAgICAgICRkZXZpY2VzTmV4dExpbmsgPSAkcmVzcG9uc2UuIkBvZGF0YS5uZXh0TGluayINCiAgICAN
CiAgICAgICAgICAgICAgICAgICAgd2hpbGUgKCRudWxsIC1uZSAkZGV2aWNlc05leHRMaW5rKSB7DQogICAgICAgICAgICAgICAgICAgICAgICAkZGV2aWNl
c1Jlc3BvbnNlID0gKEludm9rZS1NZ0dyYXBoUmVxdWVzdCAtVXJpICRkZXZpY2VzTmV4dExpbmsgLU1ldGhvZCBHZXQgLU91dHB1dFR5cGUgUFNPYmplY3Qp
DQogICAgICAgICAgICAgICAgICAgICAgICAkZGV2aWNlc05leHRMaW5rID0gJGRldmljZXNSZXNwb25zZS4iQG9kYXRhLm5leHRMaW5rIg0KICAgICAgICAg
ICAgICAgICAgICAgICAgJGRldmljZXMgKz0gJGRldmljZXNSZXNwb25zZS52YWx1ZQ0KICAgICAgICAgICAgICAgICAgICB9DQogICAgDQogICAgICAgICAg
ICAgICAgICAgICRkZXZpY2VzDQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgIGNhdGNoIHsNCiAgICAgICAgICAgICAgICAgICAgV3JpdGUt
RXJyb3IgJF8uRXhjZXB0aW9uIA0KICAgICAgICAgICAgICAgICAgICBicmVhaw0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgIH0NCiAgICAgICAg
fQ0KDQogICAgICAgIGZ1bmN0aW9uIGdldGRldmljZXNhbmR1c2VycygpIHsNCiAgICAgICAgICAgICRhbGxkZXZpY2VzID0gZ2V0YWxscGFnaW5hdGlvbiAt
dXJsICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vYmV0YS9kZXZpY2VtYW5hZ2VtZW50L21hbmFnZWRkZXZpY2VzIg0KICAgICAgICAgICAgJG91dHB1
dGFycmF5ID0gQCgpDQogICAgICAgICAgICBmb3JlYWNoICgkdmFsdWUgaW4gJGFsbGRldmljZXMpIHsNCiAgICAgICAgICAgICAgICAkb2JqZWN0ZGV0YWls
cyA9IFtwc2N1c3RvbW9iamVjdF1Aew0KICAgICAgICAgICAgICAgICAgICBEZXZpY2VJRCAgICAgICAgPSAkdmFsdWUuaWQNCiAgICAgICAgICAgICAgICAg
ICAgRGV2aWNlTmFtZSAgICAgID0gJHZhbHVlLmRldmljZU5hbWUNCiAgICAgICAgICAgICAgICAgICAgT1NWZXJzaW9uICAgICAgID0gJHZhbHVlLm9wZXJh
dGluZ1N5c3RlbQ0KICAgICAgICAgICAgICAgICAgICBQcmltYXJ5VXNlciAgICAgPSAkdmFsdWUudXNlclByaW5jaXBhbE5hbWUNCiAgICAgICAgICAgICAg
ICAgICAgb3BlcmF0aW5nU3lzdGVtID0gJHZhbHVlLm9wZXJhdGluZ1N5c3RlbQ0KICAgICAgICAgICAgICAgICAgICBBQURJRCAgICAgICAgICAgPSAkdmFs
dWUuYXp1cmVBY3RpdmVEaXJlY3RvcnlEZXZpY2VJZA0KICAgICAgICAgICAgICAgICAgICBTZXJpYWxOdW1iZXIgICAgPSAkdmFsdWUuc2VyaWFsbnVtYmVy
DQoNCiAgICAgICAgICAgICAgICB9DQogICAgDQogICAgDQogICAgICAgICAgICAgICAgJG91dHB1dGFycmF5ICs9ICRvYmplY3RkZXRhaWxzDQogICAgDQog
ICAgICAgICAgICB9DQogICAgDQogICAgICAgICAgICByZXR1cm4gJG91dHB1dGFycmF5DQogICAgICAgIH0NCg0KICAgICAgICBmdW5jdGlvbiBnZXRhbGxw
YWdpbmF0aW9uICgpIHsNCiAgICAgICAgICAgIDwjDQogICAgLlNZTk9QU0lTDQogICAgVGhpcyBmdW5jdGlvbiBpcyB1c2VkIHRvIGdyYWIgYWxsIGl0ZW1z
IGZyb20gR3JhcGggQVBJIHRoYXQgYXJlIHBhZ2luYXRlZA0KICAgIC5ERVNDUklQVElPTg0KICAgIFRoZSBmdW5jdGlvbiBjb25uZWN0cyB0byB0aGUgR3Jh
cGggQVBJIEludGVyZmFjZSBhbmQgZ2V0cyBhbGwgaXRlbXMgZnJvbSB0aGUgQVBJIHRoYXQgYXJlIHBhZ2luYXRlZA0KICAgIC5FWEFNUExFDQogICAgZ2V0
YWxscGFnaW5hdGlvbiAtdXJsICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vdjEuMC9ncm91cHMiDQogICAgIFJldHVybnMgYWxsIGl0ZW1zDQogICAg
Lk5PVEVTDQogICAgIE5BTUU6IGdldGFsbHBhZ2luYXRpb24NCiAgICAjPg0KICAgICAgICAgICAgW2NtZGxldGJpbmRpbmcoKV0NCiAgICAgICAgDQogICAg
ICAgICAgICBwYXJhbQ0KICAgICAgICAgICAgKA0KICAgICAgICAgICAgICAgICR1cmwNCiAgICAgICAgICAgICkNCiAgICAgICAgICAgICRyZXNwb25zZSA9
IChJbnZva2UtTWdHcmFwaFJlcXVlc3QgLVVyaSAkdXJsIC1NZXRob2QgR2V0IC1PdXRwdXRUeXBlIFBTT2JqZWN0KQ0KICAgICAgICAgICAgJGFsbG91dHB1
dCA9ICRyZXNwb25zZS52YWx1ZQ0KICAgICAgICANCiAgICAgICAgICAgICRhbGxvdXRwdXROZXh0TGluayA9ICRyZXNwb25zZS4iQG9kYXRhLm5leHRMaW5r
Ig0KICAgICAgICANCiAgICAgICAgICAgIHdoaWxlICgkbnVsbCAtbmUgJGFsbG91dHB1dE5leHRMaW5rKSB7DQogICAgICAgICAgICAgICAgJGFsbG91dHB1
dFJlc3BvbnNlID0gKEludm9rZS1NZ0dyYXBoUmVxdWVzdCAtVXJpICRhbGxvdXRwdXROZXh0TGluayAtTWV0aG9kIEdldCAtT3V0cHV0VHlwZSBQU09iamVj
dCkNCiAgICAgICAgICAgICAgICAkYWxsb3V0cHV0TmV4dExpbmsgPSAkYWxsb3V0cHV0UmVzcG9uc2UuIkBvZGF0YS5uZXh0TGluayINCiAgICAgICAgICAg
ICAgICAkYWxsb3V0cHV0ICs9ICRhbGxvdXRwdXRSZXNwb25zZS52YWx1ZQ0KICAgICAgICAgICAgfQ0KICAgICAgICANCiAgICAgICAgICAgIHJldHVybiAk
YWxsb3V0cHV0DQogICAgICAgIH0NCg0KICAgICAgICANCiAgICAgICAgZnVuY3Rpb24gY2hlY2staW1wb3J0ZWRkZXZpY2Ugew0KICAgICAgICAgICAgPCMN
CiAgICAuU1lOT1BTSVMNCiAgICBUaGlzIGZ1bmN0aW9uIGlzIHVzZWQgdG8gY2hlY2sgaWYgYSBkZXZpY2UgaWRlbnRpZmllciAoV2luZG93cykgYWxyZWFk
eSBleGlzdHMgaW4gdGhlIEludHVuZSBlbnZpcm9ubWVudA0KICAgIC5ERVNDUklQVElPTg0KICAgIFRoaXMgZnVuY3Rpb24gaXMgdXNlZCB0byBjaGVjayBp
ZiBhIGRldmljZSBpZGVudGlmaWVyIChXaW5kb3dzKSBhbHJlYWR5IGV4aXN0cyBpbiB0aGUgSW50dW5lIGVudmlyb25tZW50DQogICAgLkVYQU1QTEUNCiAg
ICBjaGVjay1pbXBvcnRlZGRldmljZSAtbWFudWZhY3R1cmVyICJNaWNyb3NvZnQgQ29ycG9yYXRpb24iIC1tb2RlbCAiVmlydHVhbCBNYWNoaW5lIiAtc2Vy
aWFsICJ4eHh4eCINCiAgICBSZXR1cm5zIHRydWUgb3IgZmFsc2UNCiAgICAuTk9URVMNCiAgICBOQU1FOiBjaGVjay1pbXBvcnRlZGRldmljZQ0KICAgICM+
DQogICAgICAgICAgICBbY21kbGV0YmluZGluZygpXQ0KICAgIA0KICAgICAgICAgICAgcGFyYW0NCiAgICAgICAgICAgICgNCiAgICAgICAgICAgICAgICAk
bWFudWZhY3R1cmVyLA0KICAgICAgICAgICAgICAgICRtb2RlbCwNCiAgICAgICAgICAgICAgICAkc2VyaWFsDQogICAgICAgICAgICApDQogICAgICAgICAg
ICAjI0NoZWNrIGl0IGV4aXN0cw0KICAgICAgICAgICAgJHVyaSA9ICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vYmV0YS9kZXZpY2VNYW5hZ2VtZW50
L2ltcG9ydGVkRGV2aWNlSWRlbnRpdGllcy9zZWFyY2hFeGlzdGluZ0lkZW50aXRpZXMiDQogICAgICAgICAgICAkanNvbiA9IEAiDQogICAgew0KICAgICAg
ICAiaW1wb3J0ZWREZXZpY2VJZGVudGl0aWVzIjogWw0KICAgICAgICAgICAgew0KICAgICAgICAgICAgICAgICJpbXBvcnRlZERldmljZUlkZW50aWZpZXIi
OiAiJG1hbnVmYWN0dXJlciwkbW9kZWwsJHNlcmlhbCIsDQogICAgICAgICAgICAgICAgImltcG9ydGVkRGV2aWNlSWRlbnRpdHlUeXBlIjogIm1hbnVmYWN0
dXJlck1vZGVsU2VyaWFsIg0KICAgICAgICAgICAgfQ0KICAgICAgICBdDQogICAgfQ0KIkANCiAgICAgICAgICAgICRyZXNwb25zZSA9IChJbnZva2UtTWdH
cmFwaFJlcXVlc3QgLVVyaSAkdXJpIC1NZXRob2QgUG9zdCAtQm9keSAkanNvbiAtT3V0cHV0VHlwZSBQU09iamVjdCkudmFsdWUNCiAgICANCiAgICANCiAg
ICAgICAgICAgIGlmICghJHJlc3BvbnNlKSB7DQogICAgICAgICAgICAgICAgcmV0dXJuICRmYWxzZQ0KICAgICAgICAgICAgfQ0KICAgICAgICAgICAgZWxz
ZSB7DQogICAgICAgICAgICAgICAgcmV0dXJuICR0cnVlDQogICAgICAgICAgICB9DQogICAgDQogICAgICAgIH0NCiAgICANCg0KICAgICAgICBmdW5jdGlv
biBpbXBvcnQtZGV2aWNlaWRlbnRpZmllciB7DQogICAgICAgICAgICA8Iw0KLlNZTk9QU0lTDQpUaGlzIGZ1bmN0aW9uIGlzIHVzZWQgdG8gaW1wb3J0IGEg
ZGV2aWNlIGlkZW50aWZpZXIgKFdpbmRvd3MpIGFscmVhZHkgZXhpc3RzIGluIHRoZSBJbnR1bmUgZW52aXJvbm1lbnQNCi5ERVNDUklQVElPTg0KVGhpcyBm
dW5jdGlvbiBpcyB1c2VkIHRvIGltcG9ydCBhIGRldmljZSBpZGVudGlmaWVyIChXaW5kb3dzKSBhbHJlYWR5IGV4aXN0cyBpbiB0aGUgSW50dW5lIGVudmly
b25tZW50DQouRVhBTVBMRQ0KaW1wb3J0LWRldmljZWlkZW50aWZpZXIgLW1hbnVmYWN0dXJlciAiTWljcm9zb2Z0IENvcnBvcmF0aW9uIiAtbW9kZWwgIlZp
cnR1YWwgTWFjaGluZSIgLXNlcmlhbCAieHh4eHgiDQpSZXR1cm5zIHRydWUgb3IgZmFsc2UNCi5OT1RFUw0KTkFNRTogaW1wb3J0LWRldmljZWlkZW50aWZp
ZXINCiM+DQogICAgICAgICAgICBbY21kbGV0YmluZGluZygpXQ0KDQogICAgICAgICAgICBwYXJhbQ0KICAgICAgICAgICAgKA0KICAgICAgICAgICAgICAg
ICRtYW51ZmFjdHVyZXIsDQogICAgICAgICAgICAgICAgJG1vZGVsLA0KICAgICAgICAgICAgICAgICRzZXJpYWwNCiAgICAgICAgICAgICkNCiAgICAgICAg
ICAgICMjU2VuZCBpdA0KICAgICAgICAgICAgJHVyaSA9ICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vYmV0YS9kZXZpY2VNYW5hZ2VtZW50L2ltcG9y
dGVkRGV2aWNlSWRlbnRpdGllcy9pbXBvcnREZXZpY2VJZGVudGl0eUxpc3QiDQoNCiAgICAgICAgICAgICRqc29uID0gQCINCnsNCiJpbXBvcnRlZERldmlj
ZUlkZW50aXRpZXMiOiBbDQogICAgew0KICAgICAgICAiaW1wb3J0ZWREZXZpY2VJZGVudGlmaWVyIjogIiRtYW51ZmFjdHVyZXIsJG1vZGVsLCRzZXJpYWwi
LA0KICAgICAgICAiaW1wb3J0ZWREZXZpY2VJZGVudGl0eVR5cGUiOiAibWFudWZhY3R1cmVyTW9kZWxTZXJpYWwiDQogICAgfQ0KXSwNCiJvdmVyd3JpdGVJ
bXBvcnRlZERldmljZUlkZW50aXRpZXMiOiBmYWxzZQ0KfQ0KIkANCiAgICAgICAgICAgIEludm9rZS1NZ0dyYXBoUmVxdWVzdCAtVXJpICR1cmkgLU1ldGhv
ZCBQb3N0IC1Cb2R5ICRqc29uIC1PdXRwdXRUeXBlIFBTT2JqZWN0DQogICAgICAgIH0NCg0KDQoNCiAgICAgICAgIyBDb25uZWN0DQogICAgICAgIGlmICgo
JEFwcElkIC1uZSAiIikgLWFuZCAoJEFwcFNlY3JldCAtbmUgIiIpKSB7DQogICAgICAgICAgICBTZXQtTWdHcmFwaE9wdGlvbiAtRGlzYWJsZUxvZ2luQnlX
QU0gJHRydWUNCiAgICAgICAgICAgIENvbm5lY3QtVG9HcmFwaCAtQXBwSWQgJEFwcElkIC1BcHBTZWNyZXQgJEFwcFNlY3JldCAtVGVuYW50ICRUZW5hbnRJ
ZCB8IE91dC1OdWxsDQogICAgICAgIH0NCiAgICAgICAgZWxzZWlmICgkQ2VydGlmaWNhdGVUaHVtYnByaW50IC1uZSAiIikgew0KICAgICAgICAgICAgQ29u
bmVjdC1Ub0dyYXBoIC1BcHBJZCAkQXBwSWQgLUNlcnRpZmljYXRlVGh1bWJwcmludCAkQ2VydGlmaWNhdGVUaHVtYnByaW50IC1UZW5hbnQgJFRlbmFudElk
IHwgT3V0LU51bGwNCiAgICAgICAgfQ0KICAgICAgICBlbHNlIHsNCiAgICAgICAgICAgICMjRGlzYWJsZSBXQU0NCiAgICAgICAgICAgIHNldHggTVNBTF9G
T1JDRV9XQU0gMCANCiAgICAgICAgICAgICRncmFwaCA9IENvbm5lY3QtVG9HcmFwaCAtc2NvcGVzICJEZXZpY2UuUmVhZFdyaXRlLkFsbCwgRGV2aWNlTWFu
YWdlbWVudE1hbmFnZWREZXZpY2VzLlJlYWRXcml0ZS5BbGwsIERldmljZU1hbmFnZW1lbnRTZXJ2aWNlQ29uZmlnLlJlYWRXcml0ZS5BbGwsIERldmljZU1h
bmFnZW1lbnRTY3JpcHRzLlJlYWRXcml0ZS5BbGwiDQogICAgICAgICAgICBXcml0ZS1Ib3N0ICJDb25uZWN0ZWQgdG8gSW50dW5lIHRlbmFudCAkKCRncmFw
aC5UZW5hbnRJZCkiDQogICAgICAgICAgICBpZiAoJEFkZFRvR3JvdXApIHsNCiAgICAgICAgICAgICAgICAkYWFkSWQgPSBDb25uZWN0LVRvR3JhcGggLXNj
b3BlcyAiR3JvdXAuUmVhZFdyaXRlLkFsbCwgRGV2aWNlLlJlYWRXcml0ZS5BbGwsIERldmljZU1hbmFnZW1lbnRNYW5hZ2VkRGV2aWNlcy5SZWFkV3JpdGUu
QWxsLCBEZXZpY2VNYW5hZ2VtZW50U2VydmljZUNvbmZpZy5SZWFkV3JpdGUuQWxsLCBHcm91cE1lbWJlci5SZWFkV3JpdGUuQWxsLCBEZXZpY2VNYW5hZ2Vt
ZW50U2NyaXB0cy5SZWFkV3JpdGUuQWxsIg0KICAgICAgICAgICAgICAgIFdyaXRlLUhvc3QgIkNvbm5lY3RlZCB0byBFbnRyYSB0ZW5hbnQgJCgkYWFkSWQu
VGVuYW50SWQpIg0KICAgICAgICAgICAgfQ0KICAgICAgICB9DQoNCiAgICAgICAgIyBGb3JjZSB0aGUgb3V0cHV0IHRvIGEgZmlsZQ0KICAgICAgICBpZiAo
JE91dHB1dEZpbGUgLWVxICIiKSB7DQogICAgICAgICAgICAkT3V0cHV0RmlsZSA9ICIkKCRlbnY6VEVNUClcYXV0b3BpbG90LmNzdiINCiAgICAgICAgfSAN
CiAgICB9DQp9DQoNCg0KDQpQcm9jZXNzIHsNCg0KICAgICMjQ2hlY2sgSW1wb3J0Q1NWIGlzIGVtcHR5DQogICAgaWYgKCRJbnB1dEZpbGUgLWVxICIiKSB7
DQoNCiAgICAgICAgZm9yZWFjaCAoJGNvbXAgaW4gJE5hbWUpIHsNCiAgICAgICAgICAgICRiYWQgPSAkZmFsc2UNCg0KICAgICAgICAgICAgIyBHZXQgYSBD
SU0gc2Vzc2lvbg0KICAgICAgICAgICAgaWYgKCRjb21wIC1lcSAibG9jYWxob3N0Iikgew0KICAgICAgICAgICAgICAgICRzZXNzaW9uID0gTmV3LUNpbVNl
c3Npb24NCiAgICAgICAgICAgIH0NCiAgICAgICAgICAgIGVsc2Ugew0KICAgICAgICAgICAgICAgICRzZXNzaW9uID0gTmV3LUNpbVNlc3Npb24gLUNvbXB1
dGVyTmFtZSAkY29tcCAtQ3JlZGVudGlhbCAkQ3JlZGVudGlhbA0KICAgICAgICAgICAgfQ0KDQogICAgICAgICAgICAjIEdldCB0aGUgY29tbW9uIHByb3Bl
cnRpZXMuDQogICAgICAgICAgICBXcml0ZS1WZXJib3NlICJDaGVja2luZyAkY29tcCINCiAgICAgICAgICAgICRzZXJpYWwgPSAoR2V0LUNpbUluc3RhbmNl
IC1DaW1TZXNzaW9uICRzZXNzaW9uIC1DbGFzcyBXaW4zMl9CSU9TKS5TZXJpYWxOdW1iZXINCg0KICAgICAgICAgICAgaWYgKCRpZGVudGlmaWVyKSB7DQog
ICAgICAgICAgICAgICAgJGNzID0gR2V0LUNpbUluc3RhbmNlIC1DaW1TZXNzaW9uICRzZXNzaW9uIC1DbGFzcyBXaW4zMl9Db21wdXRlclN5c3RlbQ0KICAg
ICAgICAgICAgICAgICRtYWtlID0gJGNzLk1hbnVmYWN0dXJlci5UcmltKCkuUmVwbGFjZSgiLiIsICIiKS5SZXBsYWNlKCIsIiwgIiIpDQogICAgICAgICAg
ICAgICAgJG1vZGVsID0gJGNzLk1vZGVsLlRyaW0oKS5SZXBsYWNlKCIuIiwgIiIpLlJlcGxhY2UoIiwiLCAiIikNCiAgICAgICAgICAgIA0KICAgICAgICAg
ICAgfQ0KICAgICAgICAgICAgZWxzZSB7DQogICAgICAgICAgICAgICAgIyBHZXQgdGhlIGhhc2ggKGlmIGF2YWlsYWJsZSkNCiAgICAgICAgICAgICAgICAk
ZGV2RGV0YWlsID0gKEdldC1DaW1JbnN0YW5jZSAtQ2ltU2Vzc2lvbiAkc2Vzc2lvbiAtTmFtZXNwYWNlIHJvb3QvY2ltdjIvbWRtL2RtbWFwIC1DbGFzcyBN
RE1fRGV2RGV0YWlsX0V4dDAxIC1GaWx0ZXIgIkluc3RhbmNlSUQ9J0V4dCcgQU5EIFBhcmVudElEPScuL0RldkRldGFpbCciKQ0KICAgICAgICAgICAgICAg
IGlmICgkZGV2RGV0YWlsKSB7DQogICAgICAgICAgICAgICAgICAgICRoYXNoID0gJGRldkRldGFpbC5EZXZpY2VIYXJkd2FyZURhdGENCiAgICAgICAgICAg
ICAgICB9DQogICAgICAgICAgICAgICAgZWxzZSB7DQogICAgICAgICAgICAgICAgICAgICRiYWQgPSAkdHJ1ZQ0KICAgICAgICAgICAgICAgICAgICAkaGFz
aCA9ICIiDQogICAgICAgICAgICAgICAgfQ0KICAgIA0KICAgICAgICAgICAgICAgICMgSWYgdGhlIGhhc2ggaXNuJ3QgYXZhaWxhYmxlLCBnZXQgdGhlIG1h
a2UgYW5kIG1vZGVsDQogICAgICAgICAgICAgICAgaWYgKCRiYWQpIHsNCiAgICAgICAgICAgICAgICAgICAgJGNzID0gR2V0LUNpbUluc3RhbmNlIC1DaW1T
ZXNzaW9uICRzZXNzaW9uIC1DbGFzcyBXaW4zMl9Db21wdXRlclN5c3RlbQ0KICAgICAgICAgICAgICAgICAgICAkbWFrZSA9ICRjcy5NYW51ZmFjdHVyZXIu
VHJpbSgpDQogICAgICAgICAgICAgICAgICAgICRtb2RlbCA9ICRjcy5Nb2RlbC5UcmltKCkNCiAgICAgICAgICAgICAgICAgICAgaWYgKCRQYXJ0bmVyKSB7
DQogICAgICAgICAgICAgICAgICAgICAgICAkYmFkID0gJGZhbHNlDQogICAgICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICB9DQogICAgICAg
ICAgICAgICAgZWxzZSB7DQogICAgICAgICAgICAgICAgICAgICRtYWtlID0gIiINCiAgICAgICAgICAgICAgICAgICAgJG1vZGVsID0gIiINCiAgICAgICAg
ICAgICAgICB9DQogICAgICAgICAgICB9DQogICAgICAgICAgICAjIEdldHRpbmcgdGhlIFBLSUQgaXMgZ2VuZXJhbGx5IHByb2JsZW1hdGljIGZvciBhbnlv
bmUgb3RoZXIgdGhhbiBPRU1zLCBzbyBsZXQncyBza2lwIGl0IGhlcmUNCiAgICAgICAgICAgICRwcm9kdWN0ID0gIiINCg0KICAgICAgICAgICAgIyBEZXBl
bmRpbmcgb24gdGhlIGZvcm1hdCByZXF1ZXN0ZWQsIGNyZWF0ZSB0aGUgbmVjZXNzYXJ5IG9iamVjdA0KICAgICAgICAgICAgaWYgKCRQYXJ0bmVyKSB7DQog
ICAgICAgICAgICAgICAgIyBDcmVhdGUgYSBwaXBlbGluZSBvYmplY3QNCiAgICAgICAgICAgICAgICAkYyA9IE5ldy1PYmplY3QgcHNvYmplY3QgLVByb3Bl
cnR5IEB7DQogICAgICAgICAgICAgICAgICAgICJEZXZpY2UgU2VyaWFsIE51bWJlciIgPSAkc2VyaWFsDQogICAgICAgICAgICAgICAgICAgICJXaW5kb3dz
IFByb2R1Y3QgSUQiICAgPSAkcHJvZHVjdA0KICAgICAgICAgICAgICAgICAgICAiSGFyZHdhcmUgSGFzaCIgICAgICAgID0gJGhhc2gNCiAgICAgICAgICAg
ICAgICAgICAgIk1hbnVmYWN0dXJlciBuYW1lIiAgICA9ICRtYWtlDQogICAgICAgICAgICAgICAgICAgICJEZXZpY2UgbW9kZWwiICAgICAgICAgPSAkbW9k
ZWwNCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgIyBGcm9tIHNwZWM6DQogICAgICAgICAgICAgICAgIyAiTWFudWZhY3R1cmVyIE5hbWUi
ID0gJG1ha2UNCiAgICAgICAgICAgICAgICAjICJEZXZpY2UgTmFtZSIgPSAkbW9kZWwNCg0KICAgICAgICAgICAgfQ0KICAgICAgICAgICAgZWxzZWlmICgk
aWRlbnRpZmllcikgew0KICAgICAgICAgICAgICAgICMgQ3JlYXRlIGEgcGlwZWxpbmUgb2JqZWN0DQogICAgICAgICAgICAgICAgJGMgPSBOZXctT2JqZWN0
IHBzb2JqZWN0IC1Qcm9wZXJ0eSBAew0KICAgICAgICAgICAgICAgICAgICAiU2VyaWFsIiAgICAgICA9ICRzZXJpYWwNCiAgICAgICAgICAgICAgICAgICAg
Ik1hbnVmYWN0dXJlciIgPSAkbWFrZQ0KICAgICAgICAgICAgICAgICAgICAiTW9kZWwiICAgICAgICA9ICRtb2RlbA0KICAgICAgICAgICAgICAgIH0NCiAg
ICAgICAgICAgIH0NCiAgICAgICAgICAgIGVsc2Ugew0KICAgICAgICAgICAgICAgICMgQ3JlYXRlIGEgcGlwZWxpbmUgb2JqZWN0DQogICAgICAgICAgICAg
ICAgJGMgPSBOZXctT2JqZWN0IHBzb2JqZWN0IC1Qcm9wZXJ0eSBAew0KICAgICAgICAgICAgICAgICAgICAiRGV2aWNlIFNlcmlhbCBOdW1iZXIiID0gJHNl
cmlhbA0KICAgICAgICAgICAgICAgICAgICAiV2luZG93cyBQcm9kdWN0IElEIiAgID0gJHByb2R1Y3QNCiAgICAgICAgICAgICAgICAgICAgIkhhcmR3YXJl
IEhhc2giICAgICAgICA9ICRoYXNoDQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgDQogICAgICAgICAgICAgICAgaWYgKCRHcm91cFRhZyAtbmUg
IiIpIHsNCiAgICAgICAgICAgICAgICAgICAgQWRkLU1lbWJlciAtSW5wdXRPYmplY3QgJGMgLU5vdGVQcm9wZXJ0eU5hbWUgIkdyb3VwIFRhZyIgLU5vdGVQ
cm9wZXJ0eVZhbHVlICRHcm91cFRhZw0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBpZiAoJEFzc2lnbmVkVXNlciAtbmUgIiIpIHsNCiAg
ICAgICAgICAgICAgICAgICAgQWRkLU1lbWJlciAtSW5wdXRPYmplY3QgJGMgLU5vdGVQcm9wZXJ0eU5hbWUgIkFzc2lnbmVkIFVzZXIiIC1Ob3RlUHJvcGVy
dHlWYWx1ZSAkQXNzaWduZWRVc2VyDQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgfQ0KDQogICAgICAgICAgICAjIFdyaXRlIHRoZSBvYmplY3Qg
dG8gdGhlIHBpcGVsaW5lIG9yIGFycmF5DQogICAgICAgICAgICBpZiAoJGJhZCkgew0KICAgICAgICAgICAgICAgICMgUmVwb3J0IGFuIGVycm9yIHdoZW4g
dGhlIGhhc2ggaXNuJ3QgYXZhaWxhYmxlDQogICAgICAgICAgICAgICAgV3JpdGUtRXJyb3IgLU1lc3NhZ2UgIlVuYWJsZSB0byByZXRyaWV2ZSBkZXZpY2Ug
aGFyZHdhcmUgZGF0YSAoaGFzaCkgZnJvbSBjb21wdXRlciAkY29tcCIgLUNhdGVnb3J5IERldmljZUVycm9yDQogICAgICAgICAgICB9DQogICAgICAgICAg
ICBlbHNlaWYgKCRPdXRwdXRGaWxlIC1lcSAiIikgew0KICAgICAgICAgICAgICAgICRjDQogICAgICAgICAgICB9DQogICAgICAgICAgICBlbHNlIHsNCiAg
ICAgICAgICAgICAgICAkY29tcHV0ZXJzICs9ICRjDQogICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiR2F0aGVyZWQgZGV0YWlscyBmb3IgZGV2aWNlIHdp
dGggc2VyaWFsIG51bWJlcjogJHNlcmlhbCINCiAgICAgICAgICAgIH0NCg0KICAgICAgICAgICAgUmVtb3ZlLUNpbVNlc3Npb24gJHNlc3Npb24NCiAgICAg
ICAgfQ0KICAgIH0NCiAgICBlbHNlIHsNCiAgICAgICAgd3JpdGUtaG9zdCAiQ1NWIEltcG9ydGVkLCBza2lwcGluZyBkZXZpY2UgY2hlY2siDQogICAgfQ0K
DQp9DQoNCkVuZCB7DQogICAgaWYgKCRPdXRwdXRGaWxlIC1uZSAiIikgew0KICAgICAgICBpZiAoJEFwcGVuZCkgew0KICAgICAgICAgICAgaWYgKFRlc3Qt
UGF0aCAkT3V0cHV0RmlsZSkgew0KICAgICAgICAgICAgICAgICRjb21wdXRlcnMgKz0gSW1wb3J0LUNzdiAtUGF0aCAkT3V0cHV0RmlsZQ0KICAgICAgICAg
ICAgfQ0KICAgICAgICB9DQogICAgICAgIGlmICgkUGFydG5lcikgew0KICAgICAgICAgICAgJGNvbXB1dGVycyB8IFNlbGVjdC1PYmplY3QgIkRldmljZSBT
ZXJpYWwgTnVtYmVyIiwgIldpbmRvd3MgUHJvZHVjdCBJRCIsICJIYXJkd2FyZSBIYXNoIiwgIk1hbnVmYWN0dXJlciBuYW1lIiwgIkRldmljZSBtb2RlbCIg
fCBDb252ZXJ0VG8tQ3N2IC1Ob1R5cGVJbmZvcm1hdGlvbiB8IEZvckVhY2gtT2JqZWN0IHsgJF8gLXJlcGxhY2UgJyInLCAnJyB9IHwgT3V0LUZpbGUgJE91
dHB1dEZpbGUNCiAgICAgICAgfQ0KICAgICAgICBlbHNlaWYgKCRBc3NpZ25lZFVzZXIgLW5lICIiKSB7DQogICAgICAgICAgICAkY29tcHV0ZXJzIHwgU2Vs
ZWN0LU9iamVjdCAiRGV2aWNlIFNlcmlhbCBOdW1iZXIiLCAiV2luZG93cyBQcm9kdWN0IElEIiwgIkhhcmR3YXJlIEhhc2giLCAiR3JvdXAgVGFnIiwgIkFz
c2lnbmVkIFVzZXIiIHwgQ29udmVydFRvLUNzdiAtTm9UeXBlSW5mb3JtYXRpb24gfCBGb3JFYWNoLU9iamVjdCB7ICRfIC1yZXBsYWNlICciJywgJycgfSB8
IE91dC1GaWxlICRPdXRwdXRGaWxlDQogICAgICAgIH0NCiAgICAgICAgZWxzZWlmICgkR3JvdXBUYWcgLW5lICIiKSB7DQogICAgICAgICAgICAkY29tcHV0
ZXJzIHwgU2VsZWN0LU9iamVjdCAiRGV2aWNlIFNlcmlhbCBOdW1iZXIiLCAiV2luZG93cyBQcm9kdWN0IElEIiwgIkhhcmR3YXJlIEhhc2giLCAiR3JvdXAg
VGFnIiB8IENvbnZlcnRUby1Dc3YgLU5vVHlwZUluZm9ybWF0aW9uIHwgRm9yRWFjaC1PYmplY3QgeyAkXyAtcmVwbGFjZSAnIicsICcnIH0gfCBPdXQtRmls
ZSAkT3V0cHV0RmlsZQ0KICAgICAgICB9DQogICAgICAgIGVsc2VpZiAoJGlkZW50aWZpZXIpIHsNCiAgICAgICAgICAgICRjb21wdXRlcnMgfCBTZWxlY3Qt
T2JqZWN0ICJNYW51ZmFjdHVyZXIiLCAiTW9kZWwiLCAiU2VyaWFsIiB8IENvbnZlcnRUby1Dc3YgLU5vVHlwZUluZm9ybWF0aW9uIHwgRm9yRWFjaC1PYmpl
Y3QgeyAkXyAtcmVwbGFjZSAnIicsICcnIH0gfCBTZWxlY3QtT2JqZWN0IC1Ta2lwIDEgfCBPdXQtRmlsZSAkT3V0cHV0RmlsZQ0KICAgICAgICB9DQogICAg
ICAgIGVsc2Ugew0KICAgICAgICAgICAgJGNvbXB1dGVycyB8IFNlbGVjdC1PYmplY3QgIkRldmljZSBTZXJpYWwgTnVtYmVyIiwgIldpbmRvd3MgUHJvZHVj
dCBJRCIsICJIYXJkd2FyZSBIYXNoIiB8IENvbnZlcnRUby1Dc3YgLU5vVHlwZUluZm9ybWF0aW9uIHwgRm9yRWFjaC1PYmplY3QgeyAkXyAtcmVwbGFjZSAn
IicsICcnIH0gfCBPdXQtRmlsZSAkT3V0cHV0RmlsZQ0KICAgICAgICB9DQogICAgfQ0KICAgIGlmICgkT25saW5lKSB7DQoNCiAgICAgICAgaWYgKCRpZGVu
dGlmaWVyKSB7DQogICAgICAgICAgICBpZiAoJElucHV0RmlsZSkgew0KICAgICAgICAgICAgICAgICMjSW1wb3J0IHRoZSBDU1YNCiAgICAgICAgICAgICAg
ICBJbXBvcnQtQ3N2ICRJbnB1dEZpbGUgLWhlYWRlciAiTWFudWZhY3R1cmVyIiwgIk1vZGVsIiwgIlNlcmlhbCIgfCBGb3JFYWNoLU9iamVjdCB7DQogICAg
ICAgICAgICAgICAgICAgICRzZXJpYWwgPSAkXy5TZXJpYWwNCiAgICAgICAgICAgICAgICAgICAgJG1hbnVmYWN0dXJlciA9ICRfLk1hbnVmYWN0dXJlcg0K
ICAgICAgICAgICAgICAgICAgICAkbW9kZWwgPSAkXy5Nb2RlbA0KICAgICAgICAgICAgICAgICAgICB3cml0ZS1ob3N0ICJDaGVja2luZyBpZiBkZXZpY2Ug
JHNlcmlhbCBleGlzdHMgaW4gQXV0b1BpbG90Ig0KICAgICAgICAgICAgICAgICAgICAkZXhpc3RzID0gY2hlY2staW1wb3J0ZWRkZXZpY2UgLW1hbnVmYWN0
dXJlciAkbWFudWZhY3R1cmVyIC1tb2RlbCAkbW9kZWwgLXNlcmlhbCAkc2VyaWFsDQogICAgICAgICAgICAgICAgICAgIGlmICgkZXhpc3RzIC1lcSAkZmFs
c2UpIHsNCiAgICAgICAgICAgICAgICAgICAgICAgIHdyaXRlLWhvc3QgIkRldmljZSAkc2VyaWFsIGRvZXMgbm90IGV4aXN0IGluIEF1dG9QaWxvdCwgYWRk
aW5nIGl0Ig0KICAgICAgICAgICAgICAgICAgICAgICAgJGltcG9ydCA9IGltcG9ydC1kZXZpY2VpZGVudGlmaWVyIC1tYW51ZmFjdHVyZXIgJG1hbnVmYWN0
dXJlciAtbW9kZWwgJG1vZGVsIC1zZXJpYWwgJHNlcmlhbA0KICAgICAgICAgICAgICAgICAgICAgICAgd3JpdGUtaG9zdCAiRGV2aWNlICRzZXJpYWwgYWRk
ZWQgdG8gQXV0b1BpbG90Ig0KICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgICAgIGVsc2Ugew0KICAgICAgICAgICAgICAgICAgICAg
ICAgd3JpdGUtaG9zdCAiRGV2aWNlICRzZXJpYWwgYWxyZWFkeSBleGlzdHMgaW4gQXV0b1BpbG90Ig0KICAgICAgICAgICAgICAgICAgICB9DQogICAgICAg
ICAgICAgICAgfQ0KICAgICAgICAgICAgfQ0KICAgICAgICAgICAgZWxzZSB7DQogICAgICAgICAgICAgICAgJGNvbXB1dGVycyB8IEZvckVhY2gtT2JqZWN0
IHsNCiAgICAgICAgICAgICAgICAgICAgJHNlcmlhbCA9ICRfLlNlcmlhbA0KICAgICAgICAgICAgICAgICAgICAkbWFudWZhY3R1cmVyID0gJF8uTWFudWZh
Y3R1cmVyDQogICAgICAgICAgICAgICAgICAgICRtb2RlbCA9ICRfLk1vZGVsDQogICAgICAgICAgICAgICAgICAgIHdyaXRlLWhvc3QgIkNoZWNraW5nIGlm
IGRldmljZSAkc2VyaWFsIGV4aXN0cyBpbiBBdXRvUGlsb3QiDQogICAgICAgICAgICAgICAgICAgICRleGlzdHMgPSBjaGVjay1pbXBvcnRlZGRldmljZSAt
bWFudWZhY3R1cmVyICRtYW51ZmFjdHVyZXIgLW1vZGVsICRtb2RlbCAtc2VyaWFsICRzZXJpYWwNCiAgICAgICAgICAgICAgICAgICAgaWYgKCRleGlzdHMg
LWVxICRmYWxzZSkgew0KICAgICAgICAgICAgICAgICAgICAgICAgd3JpdGUtaG9zdCAiRGV2aWNlICRzZXJpYWwgZG9lcyBub3QgZXhpc3QgaW4gQXV0b1Bp
bG90LCBhZGRpbmcgaXQiDQogICAgICAgICAgICAgICAgICAgICAgICBpbXBvcnQtZGV2aWNlaWRlbnRpZmllciAtbWFudWZhY3R1cmVyICRtYW51ZmFjdHVy
ZXIgLW1vZGVsICRtb2RlbCAtc2VyaWFsICRzZXJpYWwNCiAgICAgICAgICAgICAgICAgICAgICAgIHdyaXRlLWhvc3QgIkRldmljZSAkc2VyaWFsIGFkZGVk
IHRvIEF1dG9QaWxvdCINCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICBlbHNlIHsNCiAgICAgICAgICAgICAgICAgICAgICAg
IHdyaXRlLWhvc3QgIkRldmljZSAkc2VyaWFsIGFscmVhZHkgZXhpc3RzIGluIEF1dG9QaWxvdCINCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAg
ICAgICAgIH0NCiAgICAgICAgICAgIH0NCg0KICAgICAgICB9DQoNCiAgICAgICAgZWxzZSB7DQogICAgICAgICAgICAjI0NoZWNrIGlmICRuZXdkZXZpY2Ug
aXMgZmFsc2UNCg0KICAgICAgICAgICAgaWYgKCRuZXdkZXZpY2UpIHsNCiAgICAgICAgICAgICAgICAkaW1wb3J0U3RhcnQgPSBHZXQtRGF0ZQ0KICAgICAg
ICAgICAgICAgICRpbXBvcnRlZCA9IEAoKQ0KICAgICAgICAgICAgICAgICRjb21wdXRlcnMgfCBGb3JFYWNoLU9iamVjdCB7DQogICAgICAgICAgICAgICAg
ICAgICMgQWRkIHRoZSBkZXZpY2VzDQogICAgICAgICAgICAgICAgICAgICJBZGRpbmcgTmV3IERldmljZSBzZXJpYWwgJCgkc2VyaWFsKSINCiAgICAgICAg
ICAgICAgICAgICAgJGltcG9ydFN0YXJ0ID0gR2V0LURhdGUNCiAgICAgICAgICAgICAgICAgICAgJGltcG9ydGVkID0gQCgpDQogICAgICAgICAgICAgICAg
ICAgICRjb21wdXRlcnMgfCBGb3JFYWNoLU9iamVjdCB7DQogICAgICAgICAgICAgICAgICAgICAgICAkaW1wb3J0ZWQgKz0gQWRkLUF1dG9waWxvdEltcG9y
dGVkRGV2aWNlIC1zZXJpYWxOdW1iZXIgJF8uJ0RldmljZSBTZXJpYWwgTnVtYmVyJyAtaGFyZHdhcmVJZGVudGlmaWVyICRfLidIYXJkd2FyZSBIYXNoJyAt
Z3JvdXBUYWcgJF8uJ0dyb3VwIFRhZycgLWFzc2lnbmVkVXNlciAkXy4nQXNzaWduZWQgVXNlcicNCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAg
ICAgICAgIH0NCiAgICAgICAgICAgIH0NCiAgICAgICAgICAgIGVsc2Ugew0KICAgICAgICANCiAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICJMb2FkaW5n
IGFsbCBvYmplY3RzLiBUaGlzIGNhbiB0YWtlIGEgd2hpbGUgb24gbGFyZ2UgdGVuYW50cyINCiAgICAgICAgICAgICAgICAjICRhYWREZXZpY2VzID0gZ2V0
YWxscGFnaW5hdGlvbiAtdXJsICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vYmV0YS9kZXZpY2VzIg0KDQogICAgICAgICAgICAgICAgIyRkZXZpY2Vz
ID0gZ2V0ZGV2aWNlc2FuZHVzZXJzDQoNCiAgICAgICAgICAgICAgICAjJGludHVuZWRldmljZXMgPSAkZGV2aWNlcyB8IFdoZXJlLU9iamVjdCB7ICRfLm9w
ZXJhdGluZ1N5c3RlbSAtZXEgIldpbmRvd3MiIH0NCg0KICAgICAgICAgICAgICAgICMgVXBkYXRlIGV4aXN0aW5nIGRldmljZXMgYnkgVGhpYWdvIEJlaWVy
IGh0dHBzOi8vdHdpdHRlci5jb20vdGhpYWdvYmVpZXIgaHR0cHM6Ly93d3cubGlua2VkaW4uY29tL2luL3RiZWllci8NCiAgICAgICAgDQogICAgICAgICAg
ICAgICAgJGltcG9ydFN0YXJ0ID0gR2V0LURhdGUNCiAgICAgICAgICAgICAgICAkaW1wb3J0ZWQgPSBAKCkNCiAgICAgICAgICAgICAgICAkY29tcHV0ZXJz
IHwgRm9yRWFjaC1PYmplY3Qgew0KICAgICAgICAgICAgICAgICAgICAkY3VycmVudFNlcmlhbCA9ICRfLidEZXZpY2UgU2VyaWFsIE51bWJlcicNCiAgICAg
ICAgICAgICAgICAgICAgJGRldmljZSA9IEdldC1BdXRvcGlsb3REZXZpY2UgfCBXaGVyZS1PYmplY3QgeyAkXy5zZXJpYWxOdW1iZXIgLWVxICIkKCRjdXJy
ZW50U2VyaWFsKSIgfQ0KICAgICAgICAgICAgICAgICAgICBpZiAoJGRldmljZSkgew0KICAgICAgICAgICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiRGV2
aWNlIGFscmVhZHkgZXhpc3RzIGluIEF1dG9waWxvdCINCiAgICAgICAgICAgICAgICAgICAgICAgICRzYW5pdHlDaGVja01vZGVsID0gJGRldmljZS5tb2Rl
bA0KICAgICAgICAgICAgICAgICAgICAgICAgJHNhbml0eUNoZWNrTGFzdFNlZW4gPSAkZGV2aWNlLmxhc3RDb250YWN0ZWREYXRlVGltZS5Ub1N0cmluZygi
ZGRkZCBkZC9NTS95eXl5IGhoOm1tIHR0IikNCiAgICAgICAgICAgICAgICAgICAgICAgIFdyaXRlLUhvc3QgIkF1dG9QaWxvdCBpbmRpY2F0ZXMgbW9kZWwg
aXMgYSAkc2FuaXR5Q2hlY2tNb2RlbCwgbGFzdCBjaGVja2VkLWluICRzYW5pdHlDaGVja0xhc3RTZWVuLiINCiAgICAgICAgICAgICAgICAgICAgICAgICMj
Q2hlY2sgaWYgJGRlbGV0ZSBoYXMgYmVlbiBzZXQNCiAgICAgICAgICAgICAgICAgICAgICAgIGlmICgkZGVsZXRlKSB7DQogICAgICAgICAgICAgICAgICAg
ICAgICAgICAgV3JpdGUtSG9zdCAiRGVsZXRpbmcgZGV2aWNlIGZyb20gQXV0b1BpbG90Ig0KICAgICAgICAgICAgICAgICAgICAgICAgICAgIFJlbW92ZS1B
dXRvcGlsb3REZXZpY2UgLWlkICRkZXZpY2UuaWQNCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICJEZXZpY2UgZGVsZXRlZCBmcm9t
IEF1dG9QaWxvdCINCg0KICAgICAgICAgICAgICAgICAgICANCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAjJGludHVuZWRldmljZXRvcmVtb3ZlID0g
JGludHVuZWRldmljZXMgfCBXaGVyZS1PYmplY3QgeyAkXy5TZXJpYWxOdW1iZXIgLWVxICIkKCRzZXJpYWwpIiB9ICAgDQogICAgICAgICAgICAgICAgICAg
ICAgICAgICAgJGludHVuZWRldmljZXRvcmVtb3ZlID0gKEludm9rZS1NZ0dyYXBoUmVxdWVzdCAtVXJpICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20v
YmV0YS9kZXZpY2VNYW5hZ2VtZW50L21hbmFnZWREZXZpY2VzP2AkZmlsdGVyPXN0YXJ0c1dpdGgoc2VyaWFsTnVtYmVyLCAnJHNlcmlhbCcpIiAtTWV0aG9k
IEdFVCAtT3V0cHV0VHlwZSBQU09iamVjdCkudmFsdWUNCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAkaW50dW5lZGV2aWNlaWQgPSAkaW50dW5lZGV2
aWNldG9yZW1vdmUuSUQNCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAkYWFkZGV2aWNlaWQgPSAkaW50dW5lZGV2aWNldG9yZW1vdmUuQXp1cmVBRERl
dmljZUlEICAgIA0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICRhYWR1cmkgPSAiaHR0cHM6Ly9ncmFwaC5taWNyb3NvZnQuY29tL2JldGEvZGV2aWNl
cz9gJGZpbHRlcj1kZXZpY2VJRCBlcSAnJGFhZGRldmljZWlkJyINCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAkYWFkb2JqZWN0aWQgPSAoKEludm9r
ZS1NZ0dyYXBoUmVxdWVzdCAtVXJpICRhYWR1cmkgLU1ldGhvZCBHRVQgLU91dHB1dFR5cGUgUFNPYmplY3QpLnZhbHVlKS5pZA0KICAgICAgICAgICAgICAg
ICAgICAgICAgICAgIFdyaXRlLUhvc3QgIkRlbGV0aW5nIGRldmljZSBmcm9tIEludHVuZSINCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBJbnZva2Ut
TWdHcmFwaFJlcXVlc3QgLVVyaSAiaHR0cHM6Ly9ncmFwaC5taWNyb3NvZnQuY29tL2JldGEvZGV2aWNlTWFuYWdlbWVudC9tYW5hZ2VkRGV2aWNlcy8kaW50
dW5lZGV2aWNlaWQiIC1NZXRob2QgREVMRVRFDQogICAgICAgICAgICAgICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiRGVsZXRlZCBkZXZpY2UgJHNlcmlh
bCBmcm9tIEludHVuZSINCg0KICAgICAgICAgICAgICAgICAgICAgICAgICAgIFdyaXRlLUhvc3QgIkRlbGV0aW5nIERldmljZSBmcm9tIEVudHJhIElEIg0K
ICAgICAgICAgICAgICAgICAgICAgICAgICAgIEludm9rZS1NZ0dyYXBoUmVxdWVzdCAtVXJpICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vYmV0YS9k
ZXZpY2VzLyRhYWRvYmplY3RpZCIgLU1ldGhvZCBERUxFVEUNCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICJEZWxldGVkIGRldmlj
ZSBmcm9tIEVudHJhIg0KDQogICAgICAgICAgICAgICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiQWRkaW5nIGJhY2sgdG8gQXV0b3BpbG90Ig0KICAgICAg
ICAgICAgICAgICAgICAgICAgICAgICRpbXBvcnRlZCArPSBBZGQtQXV0b3BpbG90SW1wb3J0ZWREZXZpY2UgLXNlcmlhbE51bWJlciAkXy4nRGV2aWNlIFNl
cmlhbCBOdW1iZXInIC1oYXJkd2FyZUlkZW50aWZpZXIgJF8uJ0hhcmR3YXJlIEhhc2gnIC1ncm91cFRhZyAkXy4nR3JvdXAgVGFnJyAtYXNzaWduZWRVc2Vy
ICRfLidBc3NpZ25lZCBVc2VyJw0KDQogICAgICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgICAgICAgICAjI0Vsc2VpZiAkZ3JvdXB0
YWcgaXMgc2V0DQogICAgICAgICAgICAgICAgICAgICAgICBlbHNlaWYgKCR1cGRhdGV0YWcpIHsNCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAiVXBk
YXRpbmcgRXhpc3RpbmcgRGV2aWNlIC0gV29ya2luZyBvbiBkZXZpY2Ugc2VyaWFsICQoJHNlcmlhbCkiDQogICAgICAgICAgICAgICAgICAgICAgICAgICAg
JGltcG9ydGVkICs9IFNldC1BdXRvcGlsb3REZXZpY2UgLWlkICRkZXZpY2UuSWQgLWdyb3VwVGFnICRHcm91cFRhZw0KICAgIA0KICAgICAgICAgICAgICAg
ICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICAgICAgZWxzZSB7DQogICAgICAgICAgICAgICAgICAgICAgICAgICAgIyNQcm9tcHQgdG8gZGVsZXRl
IG9yIHVwZGF0ZQ0KICAgICAgICAgICAgICAgICAgICAgICAgICAgIGlmICgkRm9yY2UpIHsNCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgJGNo
b2ljZSA9ICJ1cGRhdGUiDQogICAgICAgICAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICAgICAgICAgIGVsc2Ugew0KICAgICAg
ICAgICAgICAgICAgICAgICAgICAgICAgICAkY2hvaWNlID0gUmVhZC1Ib3N0ICJEbyB5b3Ugd2FudCB0byBkZWxldGUgb3IgdXBkYXRlPyAoZGVsZXRlL3Vw
ZGF0ZSkiDQogICAgICAgICAgICAgICAgICAgICAgICAgICAgfQ0KDQogICAgICAgICAgICAgICAgICAgICAgICAgICAgaWYgKCRjaG9pY2UgLWVxICJkZWxl
dGUiKSB7DQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICMgUGVyZm9ybSBkZWxldGUgYWN0aW9uDQogICAgICAgICAgICAgICAgICAgICAgICAg
ICAgICAgIFdyaXRlLU91dHB1dCAiWW91IGNob3NlIHRvIGRlbGV0ZS4iDQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIFdyaXRlLUhvc3QgIkRl
bGV0aW5nIGRldmljZSBmcm9tIEF1dG9QaWxvdCINCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgUmVtb3ZlLUF1dG9waWxvdERldmljZSAtaWQg
JGRldmljZS5pZA0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICJEZXZpY2UgZGVsZXRlZCBmcm9tIEF1dG9QaWxvdCINCg0K
ICAgIA0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAjJGludHVuZWRldmljZXRvcmVtb3ZlID0gJGludHVuZWRldmljZXMgfCBXaGVyZS1PYmpl
Y3QgeyAkXy5TZXJpYWxOdW1iZXIgLWVxICIkKCRzZXJpYWwpIiB9IA0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAkaW50dW5lZGV2aWNldG9y
ZW1vdmUgPSAoSW52b2tlLU1nR3JhcGhSZXF1ZXN0IC1VcmkgImh0dHBzOi8vZ3JhcGgubWljcm9zb2Z0LmNvbS9iZXRhL2RldmljZU1hbmFnZW1lbnQvbWFu
YWdlZERldmljZXM/YCRmaWx0ZXI9c3RhcnRzV2l0aChzZXJpYWxOdW1iZXIsICckc2VyaWFsJykiIC1NZXRob2QgR0VUIC1PdXRwdXRUeXBlIFBTT2JqZWN0
KS52YWx1ZSAgICAgIA0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAkaW50dW5lZGV2aWNlaWQgPSAkaW50dW5lZGV2aWNldG9yZW1vdmUuSUQN
CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgJGFhZGRldmljZWlkID0gJGludHVuZWRldmljZXRvcmVtb3ZlLkF6dXJlQUREZXZpY2VJRCAgICAN
CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgJGFhZHVyaSA9ICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vYmV0YS9kZXZpY2VzP2AkZmls
dGVyPWRldmljZUlEIGVxICckYWFkZGV2aWNlaWQnIg0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAkYWFkb2JqZWN0aWQgPSAoKEludm9rZS1N
Z0dyYXBoUmVxdWVzdCAtVXJpICRhYWR1cmkgLU1ldGhvZCBHRVQgLU91dHB1dFR5cGUgUFNPYmplY3QpLnZhbHVlKS5pZA0KICAgICAgICAgICAgICAgICAg
ICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICJEZWxldGluZyBkZXZpY2UgZnJvbSBJbnR1bmUiDQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIElu
dm9rZS1NZ0dyYXBoUmVxdWVzdCAtVXJpICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vYmV0YS9kZXZpY2VNYW5hZ2VtZW50L21hbmFnZWREZXZpY2Vz
LyRpbnR1bmVkZXZpY2VpZCIgLU1ldGhvZCBERUxFVEUNCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiRGVsZXRlZCBkZXZp
Y2UgJHNlcmlhbCBmcm9tIEludHVuZSINCg0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICJEZWxldGluZyBEZXZpY2UgZnJv
bSBFbnRyYSBJRCINCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgSW52b2tlLU1nR3JhcGhSZXF1ZXN0IC1VcmkgImh0dHBzOi8vZ3JhcGgubWlj
cm9zb2Z0LmNvbS9iZXRhL2RldmljZXMvJGFhZG9iamVjdGlkIiAtTWV0aG9kIERFTEVURQ0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBXcml0
ZS1Ib3N0ICJEZWxldGVkIGRldmljZSBmcm9tIEVudHJhIg0KDQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIFdyaXRlLUhvc3QgIkFkZGluZyBi
YWNrIHRvIEF1dG9waWxvdCINCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgJGltcG9ydGVkICs9IEFkZC1BdXRvcGlsb3RJbXBvcnRlZERldmlj
ZSAtc2VyaWFsTnVtYmVyICRfLidEZXZpY2UgU2VyaWFsIE51bWJlcicgLWhhcmR3YXJlSWRlbnRpZmllciAkXy4nSGFyZHdhcmUgSGFzaCcgLWdyb3VwVGFn
ICRfLidHcm91cCBUYWcnIC1hc3NpZ25lZFVzZXIgJF8uJ0Fzc2lnbmVkIFVzZXInDQoNCiAgICAgICAgICAgICAgICAgICAgICAgICAgICB9DQogICAgICAg
ICAgICAgICAgICAgICAgICAgICAgZWxzZWlmICgkY2hvaWNlIC1lcSAidXBkYXRlIikgew0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAjIFBl
cmZvcm0gdXBkYXRlIGFjdGlvbg0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBXcml0ZS1PdXRwdXQgIllvdSBjaG9zZSB0byB1cGRhdGUuIg0K
ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAiVXBkYXRpbmcgRXhpc3RpbmcgRGV2aWNlIC0gV29ya2luZyBvbiBkZXZpY2Ugc2VyaWFsICQoJHNl
cmlhbCkiDQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICRpbXBvcnRlZCArPSBTZXQtQXV0b3BpbG90RGV2aWNlIC1pZCAkZGV2aWNlLklkIC1n
cm91cFRhZyAkR3JvdXBUYWcNCg0KICAgICAgICAgICAgICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBlbHNlIHsNCiAg
ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgV3JpdGUtT3V0cHV0ICJJbnZhbGlkIGNob2ljZS4gUGxlYXNlIGVudGVyICdkZWxldGUnIG9yICd1cGRh
dGUnLiINCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgZXhpdA0KICAgICAgICAgICAgICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAg
ICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICBlbHNlIHsNCiAgICAgICAgICAgICAgICAgICAgICAgICMg
QWRkIHRoZSBkZXZpY2VzDQogICAgICAgICAgICAgICAgICAgICAgICAiQWRkaW5nIE5ldyBEZXZpY2Ugc2VyaWFsICQoJHNlcmlhbCkiDQogICAgICAgICAg
ICAgICAgICAgICAgICAkaW1wb3J0U3RhcnQgPSBHZXQtRGF0ZQ0KICAgICAgICAgICAgICAgICAgICAgICAgJGltcG9ydGVkID0gQCgpDQogICAgICAgICAg
ICAgICAgICAgICAgICAkY29tcHV0ZXJzIHwgRm9yRWFjaC1PYmplY3Qgew0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICRpbXBvcnRlZCArPSBBZGQt
QXV0b3BpbG90SW1wb3J0ZWREZXZpY2UgLXNlcmlhbE51bWJlciAkXy4nRGV2aWNlIFNlcmlhbCBOdW1iZXInIC1oYXJkd2FyZUlkZW50aWZpZXIgJF8uJ0hh
cmR3YXJlIEhhc2gnIC1ncm91cFRhZyAkXy4nR3JvdXAgVGFnJyAtYXNzaWduZWRVc2VyICRfLidBc3NpZ25lZCBVc2VyJw0KICAgICAgICAgICAgICAgICAg
ICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgfQ0KICAgICAgICANCiAgICAgICAgICAgIH0NCg0KICAgICAgICAgICAg
IyBXYWl0IHVudGlsIHRoZSBkZXZpY2VzIGhhdmUgYmVlbiBpbXBvcnRlZA0KICAgICAgICAgICAgJHByb2Nlc3NpbmdDb3VudCA9IDENCiAgICAgICAgICAg
IHdoaWxlICgkcHJvY2Vzc2luZ0NvdW50IC1ndCAwKSB7DQogICAgICAgICAgICAgICAgJGN1cnJlbnQgPSBAKCkNCiAgICAgICAgICAgICAgICAkcHJvY2Vz
c2luZ0NvdW50ID0gMA0KICAgICAgICAgICAgICAgICRpbXBvcnRlZCB8IEZvckVhY2gtT2JqZWN0IHsNCiAgICAgICAgICAgICAgICAgICAgIyRkZXZpY2Ug
PSBHZXQtQXV0b3BpbG90SW1wb3J0ZWREZXZpY2UgLWlkICRfLmlkDQogICAgICAgICAgICAgICAgICAgICRkZXZpY2UgPSBHZXQtQXV0b3BpbG90SW1wb3J0
ZWREZXZpY2UgfCBXaGVyZS1PYmplY3QgeyAkXy5zZXJpYWxOdW1iZXIgLWVxICIkKCRzZXJpYWwpIiB9DQogICAgICAgICAgICAgICAgICAgIGlmICgkZGV2
aWNlLnN0YXRlLmRldmljZUltcG9ydFN0YXR1cyAtZXEgInVua25vd24iKSB7DQogICAgICAgICAgICAgICAgICAgICAgICAkcHJvY2Vzc2luZ0NvdW50ID0g
JHByb2Nlc3NpbmdDb3VudCArIDENCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICAkY3VycmVudCArPSAkZGV2aWNlDQogICAg
ICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICRkZXZpY2VDb3VudCA9ICRpbXBvcnRlZC5MZW5ndGgNCiAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0
ICJXYWl0aW5nIGZvciAkcHJvY2Vzc2luZ0NvdW50IG9mICRkZXZpY2VDb3VudCB0byBiZSBpbXBvcnRlZCINCiAgICAgICAgICAgICAgICBpZiAoJHByb2Nl
c3NpbmdDb3VudCAtZ3QgMCkgew0KICAgICAgICAgICAgICAgICAgICBTdGFydC1TbGVlcCAzMA0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgIH0N
CiAgICAgICAgICAgICRpbXBvcnREdXJhdGlvbiA9IChHZXQtRGF0ZSkgLSAkaW1wb3J0U3RhcnQNCiAgICAgICAgICAgICRpbXBvcnRTZWNvbmRzID0gW01h
dGhdOjpDZWlsaW5nKCRpbXBvcnREdXJhdGlvbi5Ub3RhbFNlY29uZHMpDQogICAgICAgICAgICAkc3VjY2Vzc0NvdW50ID0gMA0KICAgICAgICAgICAgJGN1
cnJlbnQgfCBGb3JFYWNoLU9iamVjdCB7DQogICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiJCgkZGV2aWNlLnNlcmlhbE51bWJlcik6ICQoJGRldmljZS5z
dGF0ZS5kZXZpY2VJbXBvcnRTdGF0dXMpICQoJGRldmljZS5zdGF0ZS5kZXZpY2VFcnJvckNvZGUpICQoJGRldmljZS5zdGF0ZS5kZXZpY2VFcnJvck5hbWUp
Ig0KICAgICAgICAgICAgICAgIGlmICgkZGV2aWNlLnN0YXRlLmRldmljZUltcG9ydFN0YXR1cyAtZXEgImNvbXBsZXRlIikgew0KICAgICAgICAgICAgICAg
ICAgICAkc3VjY2Vzc0NvdW50ID0gJHN1Y2Nlc3NDb3VudCArIDENCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICB9DQogICAgICAgICAgICBXcml0
ZS1Ib3N0ICIkc3VjY2Vzc0NvdW50IGRldmljZXMgaW1wb3J0ZWQgc3VjY2Vzc2Z1bGx5LiBFbGFwc2VkIHRpbWUgdG8gY29tcGxldGUgaW1wb3J0OiAkaW1w
b3J0U2Vjb25kcyBzZWNvbmRzIg0KICAgICAgICANCiAgICAgICAgICAgICMgV2FpdCB1bnRpbCB0aGUgZGV2aWNlcyBjYW4gYmUgZm91bmQgaW4gSW50dW5l
IChzaG91bGQgc3luYyBhdXRvbWF0aWNhbGx5KQ0KICAgICAgICAgICAgJHN5bmNTdGFydCA9IEdldC1EYXRlDQogICAgICAgICAgICAkcHJvY2Vzc2luZ0Nv
dW50ID0gMQ0KICAgICAgICAgICAgd2hpbGUgKCRwcm9jZXNzaW5nQ291bnQgLWd0IDApIHsNCiAgICAgICAgICAgICAgICAkYXV0b3BpbG90RGV2aWNlcyA9
IEAoKQ0KICAgICAgICAgICAgICAgICRwcm9jZXNzaW5nQ291bnQgPSAwDQogICAgICAgICAgICAgICAgJGN1cnJlbnQgfCBGb3JFYWNoLU9iamVjdCB7DQog
ICAgICAgICAgICAgICAgICAgIGlmICgkXy5zdGF0ZS5kZXZpY2VJbXBvcnRTdGF0dXMgLWVxICJjb21wbGV0ZSIpIHsNCiAgICAgICAgICAgICAgICAgICAg
ICAgICRyZWdpc3RyYXRpb25JZCA9ICRfLnN0YXRlLmRldmljZVJlZ2lzdHJhdGlvbklkDQogICAgICAgICAgICAgICAgICAgICAgICBpZiAoLW5vdCAkcmVn
aXN0cmF0aW9uSWQpIHsNCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAkcHJvY2Vzc2luZ0NvdW50ID0gJHByb2Nlc3NpbmdDb3VudCArIDENCiAgICAg
ICAgICAgICAgICAgICAgICAgICAgICByZXR1cm4NCiAgICAgICAgICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICAgICAgICAgICRkZXZpY2Ug
PSBHZXQtQXV0b3BpbG90RGV2aWNlIC1pZCAkcmVnaXN0cmF0aW9uSWQNCiAgICAgICAgICAgICAgICAgICAgICAgIGlmICgtbm90ICRkZXZpY2UpIHsNCiAg
ICAgICAgICAgICAgICAgICAgICAgICAgICAkcHJvY2Vzc2luZ0NvdW50ID0gJHByb2Nlc3NpbmdDb3VudCArIDENCiAgICAgICAgICAgICAgICAgICAgICAg
IH0NCiAgICAgICAgICAgICAgICAgICAgICAgICRhdXRvcGlsb3REZXZpY2VzICs9ICRkZXZpY2UNCiAgICAgICAgICAgICAgICAgICAgfSAgICANCiAgICAg
ICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgJGRldmljZUNvdW50ID0gJGF1dG9waWxvdERldmljZXMuTGVuZ3RoDQogICAgICAgICAgICAgICAgV3Jp
dGUtSG9zdCAiV2FpdGluZyBmb3IgJHByb2Nlc3NpbmdDb3VudCBvZiAkZGV2aWNlQ291bnQgdG8gYmUgc3luY2VkIg0KICAgICAgICAgICAgICAgIGlmICgk
cHJvY2Vzc2luZ0NvdW50IC1ndCAwKSB7DQogICAgICAgICAgICAgICAgICAgIFN0YXJ0LVNsZWVwIDMwDQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAg
ICAgfQ0KICAgICAgICAgICAgJHN5bmNEdXJhdGlvbiA9IChHZXQtRGF0ZSkgLSAkc3luY1N0YXJ0DQogICAgICAgICAgICAkc3luY1NlY29uZHMgPSBbTWF0
aF06OkNlaWxpbmcoJHN5bmNEdXJhdGlvbi5Ub3RhbFNlY29uZHMpDQogICAgICAgICAgICBXcml0ZS1Ib3N0ICJBbGwgZGV2aWNlcyBzeW5jZWQuIEVsYXBz
ZWQgdGltZSB0byBjb21wbGV0ZSBzeW5jOiAkc3luY1NlY29uZHMgc2Vjb25kcyINCg0KICAgICAgICAgICAgIyNPbmx5IGRvIHRoaXMgd2hlbiB1cGRhdGV0
YWcgaXMgc2V0DQogICAgICAgICAgICBpZiAoJFVwZGF0ZVRhZykgew0KICAgICAgICAgICAgR2V0LUF1dG9waWxvdEltcG9ydGVkRGV2aWNlIHwgV2hlcmUt
T2JqZWN0IHsgJF8uc2VyaWFsbnVtYmVyIC1lcSAiJHNlcmlhbCIgfSB8IEZvckVhY2gtT2JqZWN0IHsgUmVtb3ZlLUF1dG9waWxvdEltcG9ydGVkRGV2aWNl
IC1pZCAkXy5pZCB9DQogICAgICAgICAgICB0cnkgew0KICAgICAgICAgICAgICAgIEludm9rZS1BdXRvcGlsb3RTeW5jIC1FcnJvckFjdGlvbiBTdG9wDQog
ICAgICAgICAgICB9DQogICAgICAgICAgICBjYXRjaCB7DQogICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiJCgkXy5leGNlcHRpb24ubWVzc2FnZSkiDQog
ICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiQW4gZXJyb3Igb2NjdXJyZWQuIFdhaXRpbmcgZm9yIDEyLDUgbWludXRlcyBiZWZvcmUgcmV0cnlpbmcuLi4i
DQogICAgICAgICAgICAgICAgU3RhcnQtU2xlZXAgLVNlY29uZHMgNzUwDQogICAgICAgICAgICAgICAgSW52b2tlLUF1dG9waWxvdFN5bmMNCiAgICAgICAg
ICAgIH0NCiAgICAgICAgfQ0KICAgICAgICAgICAgIyBBZGQgdGhlIGRldmljZSB0byB0aGUgc3BlY2lmaWVkIEFBRCBncm91cA0KICAgICAgICAgICAgIyBB
ZGQgdGhlIGRldmljZSB0byB0aGUgc3BlY2lmaWVkIEFBRCBncm91cA0KICAgICAgICAgICAgaWYgKCRBZGRUb0dyb3VwKSB7DQogICAgICAgICAgICAgICAg
Zm9yZWFjaCAoJEFER3JvdXAgaW4gJEFkZFRvR3JvdXApIHsNCiAgICAgICAgICAgICAgICAgICAgIyRhYWRHcm91cCA9IEdldC1NZ0dyb3VwIC1GaWx0ZXIg
IkRpc3BsYXlOYW1lIGVxICckQURHcm91cCciDQogICAgICAgICAgICAgICAgICAgICRndXJpID0gImh0dHBzOi8vZ3JhcGgubWljcm9zb2Z0LmNvbS9iZXRh
L2dyb3Vwcz9gJGZpbHRlcj1kaXNwbGF5TmFtZSBlcSAnJEFER3JvdXAnIg0KICAgICAgICAgICAgICAgICAgICAkYWFkR3JvdXAgPSAoSW52b2tlLU1nR3Jh
cGhSZXF1ZXN0IC1VcmkgJGd1cmkgLU1ldGhvZCBHRVQgLU91dHB1dFR5cGUgUFNPYmplY3QpLnZhbHVlDQogICAgICAgICAgICAgICAgICAgIGlmICgkYWFk
R3JvdXApIHsNCiAgICAgICAgICAgICAgICAgICAgICAgICRhdXRvcGlsb3REZXZpY2VzIHwgRm9yRWFjaC1PYmplY3Qgew0KICAgICAgICAgICAgICAgICAg
ICAgICAgICAgICR1cmkgPSAiaHR0cHM6Ly9ncmFwaC5taWNyb3NvZnQuY29tL2JldGEvZGV2aWNlcz9gJGZpbHRlcj1kZXZpY2VJZCBlcSAnIiArICRfLmF6
dXJlQWN0aXZlRGlyZWN0b3J5RGV2aWNlSWQgKyAiJyINCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAkYWFkRGV2aWNlID0gKEludm9rZS1NZ0dyYXBo
UmVxdWVzdCAtVXJpICR1cmkgLU1ldGhvZCBHRVQgLU91dHB1dFR5cGUgUFNPYmplY3QgLVNraXBIdHRwRXJyb3JDaGVjaykudmFsdWUNCiAgICAgICAgICAg
ICAgICAgICAgICAgICAgICBpZiAoJGFhZERldmljZSkgew0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICJBZGRpbmcgZGV2
aWNlICQoJGFhZERldmljZS5kaXNwbGF5TmFtZSkgdG8gZ3JvdXAgJEFER3JvdXAiDQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICRlZ3JwaWQg
PSAkYWFkR3JvdXAuSWQNCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgJGVkdmNpZCA9ICRhYWREZXZpY2UuaWQNCg0KICAgICAgICAgICAgICAg
ICAgICAgICAgICAgICAgICAjTmV3LU1nR3JvdXBNZW1iZXIgLUdyb3VwSWQgJGFhZEdyb3VwLklkIC1EaXJlY3RvcnlPYmplY3RJZCAkYWFkRGV2aWNlLmlk
DQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICMjVXNlIEdyYXBoDQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICRlZ3VyaSA9ICJo
dHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vYmV0YS9ncm91cHMvJGVncnBpZC9tZW1iZXJzL2AkcmVmIg0KICAgICAgICAgICAgICAgICAgICAgICAgICAg
ICAgICAkanNvbiA9IEAiDQoNCiAgICAgICAgICAgICAgICAgICAgICAgICAgICB7DQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICJAb2RhdGEu
aWQiOiAiaHR0cHM6Ly9ncmFwaC5taWNyb3NvZnQuY29tL2JldGEvZGlyZWN0b3J5T2JqZWN0cy8kZWR2Y2lkIg0KICAgICAgICAgICAgICAgICAgICAgICAg
ICAgICAgfQ0KIkANCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgSW52b2tlLU1nR3JhcGhSZXF1ZXN0IC1NZXRob2QgUE9TVCAtVXJpICRlZ3Vy
aSAtQm9keSAkanNvbiAtQ29udGVudFR5cGUgImFwcGxpY2F0aW9uL2pzb24iIC1PdXRwdXRUeXBlIFBTT2JqZWN0DQogICAgICAgICAgICAgICAgICAgICAg
ICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICAgICAgICAgIGVsc2Ugew0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBXcml0ZS1FcnJvciAi
VW5hYmxlIHRvIGZpbmQgRW50cmEgZGV2aWNlIHdpdGggSUQgJCgkYWFkRGV2aWNlLmRldmljZUlkKSINCiAgICAgICAgICAgICAgICAgICAgICAgICAgICB9
DQogICAgICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICJBZGRlZCBkZXZpY2VzIHRvIGdyb3VwICck
QURHcm91cCcgKCQoJGFhZEdyb3VwLklkKSkiDQogICAgICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICAgICAgZWxzZSB7DQogICAgICAgICAg
ICAgICAgICAgICAgICBXcml0ZS1FcnJvciAiVW5hYmxlIHRvIGZpbmQgZ3JvdXAgJEFER3JvdXAiDQogICAgICAgICAgICAgICAgICAgIH0NCiAgICAgICAg
ICAgICAgICB9DQogICAgICAgICAgICB9ICN0byBkZWFsIHdpdGggdGhlIGFycmF5DQoNCiAgICAgICAgICAgICMgQXNzaWduIHRoZSBjb21wdXRlciBuYW1l
DQogICAgICAgICAgICBpZiAoJEFzc2lnbmVkQ29tcHV0ZXJOYW1lIC1uZSAiIikgew0KICAgICAgICAgICAgICAgICRhdXRvcGlsb3REZXZpY2VzIHwgRm9y
RWFjaC1PYmplY3Qgew0KICAgICAgICAgICAgICAgICAgICBTZXQtQXV0b3BpbG90RGV2aWNlIC1pZCAkXy5JZCAtZGlzcGxheU5hbWUgJEFzc2lnbmVkQ29t
cHV0ZXJOYW1lDQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgfQ0KDQogICAgICAgICAgICAjIFdhaXQgZm9yIGFzc2lnbm1lbnQgKGlmIHNwZWNp
ZmllZCkNCiAgICAgICAgICAgIGlmICgkQXNzaWduKSB7DQogICAgICAgICAgICAgICAgJGFzc2lnblN0YXJ0ID0gR2V0LURhdGUNCiAgICAgICAgICAgICAg
ICAkcHJvY2Vzc2luZ0NvdW50ID0gMQ0KICAgICAgICAgICAgICAgIHdoaWxlICgkcHJvY2Vzc2luZ0NvdW50IC1ndCAwKSB7DQogICAgICAgICAgICAgICAg
ICAgICRwcm9jZXNzaW5nQ291bnQgPSAwDQogICAgICAgICAgICAgICAgICAgICRhdXRvcGlsb3REZXZpY2VzIHwgRm9yRWFjaC1PYmplY3Qgew0KICAgICAg
ICAgICAgICAgICAgICAgICAgJGRldmljZSA9IEdldC1BdXRvcGlsb3REZXZpY2UgLWlkICRfLmlkIC1leHBhbmQNCiAgICAgICAgICAgICAgICAgICAgICAg
IGlmICgtbm90ICgkZGV2aWNlLmRlcGxveW1lbnRQcm9maWxlQXNzaWdubWVudFN0YXR1cy5TdGFydHNXaXRoKCJhc3NpZ25lZCIpKSkgew0KICAgICAgICAg
ICAgICAgICAgICAgICAgICAgICRwcm9jZXNzaW5nQ291bnQgPSAkcHJvY2Vzc2luZ0NvdW50ICsgMQ0KICAgICAgICAgICAgICAgICAgICAgICAgfQ0KICAg
ICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgICAgICRkZXZpY2VDb3VudCA9ICRhdXRvcGlsb3REZXZpY2VzLkxlbmd0aA0KICAgICAgICAg
ICAgICAgICAgICBXcml0ZS1Ib3N0ICJXYWl0aW5nIGZvciAkcHJvY2Vzc2luZ0NvdW50IG9mICRkZXZpY2VDb3VudCB0byBiZSBhc3NpZ25lZCINCiAgICAg
ICAgICAgICAgICAgICAgaWYgKCRwcm9jZXNzaW5nQ291bnQgLWd0IDApIHsNCiAgICAgICAgICAgICAgICAgICAgICAgIFN0YXJ0LVNsZWVwIDMwDQogICAg
ICAgICAgICAgICAgICAgIH0gICAgDQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICRhc3NpZ25EdXJhdGlvbiA9IChHZXQtRGF0ZSkgLSAk
YXNzaWduU3RhcnQNCiAgICAgICAgICAgICAgICAkYXNzaWduU2Vjb25kcyA9IFtNYXRoXTo6Q2VpbGluZygkYXNzaWduRHVyYXRpb24uVG90YWxTZWNvbmRz
KQ0KICAgICAgICAgICAgICAgIFdyaXRlLUhvc3QgIlByb2ZpbGVzIGFzc2lnbmVkIHRvIGFsbCBkZXZpY2VzLiBFbGFwc2VkIHRpbWUgdG8gY29tcGxldGUg
YXNzaWdubWVudDogJGFzc2lnblNlY29uZHMgc2Vjb25kcyIgICAgDQogICAgICAgICAgICAgICAgaWYgKCRSZWJvb3QpIHsNCiAgICAgICAgICAgICAgICAg
ICAgUmVzdGFydC1Db21wdXRlciAtRm9yY2UNCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgaWYgKCRXaXBlKSB7DQogICAgICAgICAgICAg
ICAgICAgICRkZXZpY2VzZXJpYWwgPSAkc2VyaWFsDQogICAgICAgICAgICAgICAgICAgICMjRmluZCBkZXZpY2UgSUQNCiAgICAgICAgICAgICAgICAgICAg
JGRldmljZXVyaSA9ICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vdjEuMC9kZXZpY2VNYW5hZ2VtZW50L21hbmFnZWREZXZpY2VzP2AkZmlsdGVyPXNl
cmlhbE51bWJlciBlcSAnJHNlcmlhbCciDQogICAgICAgICAgICAgICAgICAgICRkZXZpY2VpZCA9IChJbnZva2UtTWdHcmFwaFJlcXVlc3QgLVVyaSAkZGV2
aWNldXJpIC1NZXRob2QgR0VUIC1PdXRwdXRUeXBlIFBTT2JqZWN0IC1Ta2lwSHR0cEVycm9yQ2hlY2spLnZhbHVlLmlkDQogICAgICAgICAgICAgICAgICAg
IFdyaXRlLUhvc3QgIlNlbmRpbmcgYSB3aXBlIHRvICRkZXZpY2VpZCINCiAgICAgICAgICAgICAgICAgICAgIyNTZW5kIGEgd2lwZQ0KICAgICAgICAgICAg
ICAgICAgICAkd2lwZXVyaSA9ICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vdjEuMC9kZXZpY2VNYW5hZ2VtZW50L21hbmFnZWREZXZpY2VzLyRkZXZp
Y2VpZC93aXBlIg0KICAgICAgICAgICAgICAgICAgICAkd2lwZWJvZHkgPSBAew0KICAgICAgICAgICAgICAgICAgICAgICAga2VlcEVucm9sbG1lbnREYXRh
ID0gJGZhbHNlDQogICAgICAgICAgICAgICAgICAgICAgICBrZWVwVXNlckRhdGEgICAgICAgPSAkZmFsc2UNCiAgICAgICAgICAgICAgICAgICAgfQ0KICAg
ICAgICAgICAgICAgICAgICBJbnZva2UtTWdHcmFwaFJlcXVlc3QgLVVyaSAkd2lwZXVyaSAtTWV0aG9kIFBPU1QgLUJvZHkgJHdpcGVib2R5IC1Db250ZW50
VHlwZSAiYXBwbGljYXRpb24vanNvbiINCiAgICAgICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiV2lwZSBzZW50IHRvICRkZXZpY2VpZCINCiAgICAgICAg
ICAgICAgICB9DQogICAgICAgICAgICAgICAgaWYgKCRTeXNwcmVwKSB7DQogICAgICAgICAgICAgICAgICAgICMjU2VuZCBhIHN5c3ByZXANCiAgICAgICAg
ICAgICAgICAgICAgU3RhcnQtUHJvY2VzcyAtTm9OZXdXaW5kb3cgLUZpbGVQYXRoICJDOlx3aW5kb3dzXHN5c3RlbTMyXHN5c3ByZXBcc3lzcHJlcC5leGUi
IC1Bcmd1bWVudExpc3QgIi9vb2JlIC9yZWJvb3QgL3F1aWV0Ig0KICAgICAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICJTeXNwcmVwIGV4ZWN1dGVkIg0K
ICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBpZiAoJHByZXByb3YpIHsNCiAgICAgICAgICAgICAgICAgICAgIyNDcmVhdGUgZGlyZWN0b3J5
IGluICV0ZW1wJQ0KICAgICAgICAgICAgICAgICAgICAkcGF0aCA9ICRlbnY6VEVNUCArICJccHJlcHJvdiINCiAgICAgICAgICAgICAgICAgICAgTmV3LUl0
ZW0gLVBhdGggJHBhdGggLUl0ZW1UeXBlIERpcmVjdG9yeQ0KICAgICAgICAgICAgICAgICAgICAkdXJpID0gImh0dHBzOi8vZ2l0aHViLmNvbS9hbmRyZXct
cy10YXlsb3IvV2luZG93c0F1dG9waWxvdEluZm8vcmF3L21haW4vd2luZG93c2tleS1hdXRvaXQuZXhlIg0KICAgICAgICAgICAgICAgICAgICAjI0Rvd25s
b2FkIGl0DQogICAgICAgICAgICAgICAgICAgICRvdXRwdXQgPSAiJHBhdGhcd2luZG93c2tleS1hdXRvaXQuZXhlIg0KICAgICAgICAgICAgICAgICAgICBJ
bnZva2UtV2ViUmVxdWVzdCAtVXJpICR1cmkgLU91dEZpbGUgJG91dHB1dCAtVXNlQmFzaWNQYXJzaW5nDQogICAgICAgICAgICAgICAgICAgIFdyaXRlLUhv
c3QgIkZpbGUgZG93bmxvYWRlZCB0byAkb3V0cHV0Ig0KICAgICAgICAgICAgICAgICAgICAjI1J1biBpdA0KICAgICAgICAgICAgICAgICAgICAmJG91dHB1
dA0KICAgICAgICAgICAgICAgIA0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBpZiAoJENoYW5nZVBLIC1uZSAiIikgew0KICAgICAgICAg
ICAgICAgICAgICAjIFJ1biBDaGFuZ2VQSy5leGUNCiAgICAgICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiU3RhcnRpbmcgQ2hhbmdlUEsiDQogICAgICAg
ICAgICAgICAgICAgIFN0YXJ0LVByb2Nlc3MgLU5vTmV3V2luZG93IC1XYWl0IC1GaWxlUGF0aCAiYzpcd2luZG93c1xzeXN0ZW0zMlxjaGFuZ2Vway5leGUi
IC1Bcmd1bWVudExpc3QgIi9Qcm9kdWN0S2V5ICRDaGFuZ2VQSyAvTm9VSSAvTm9SZWJvb3QiDQogICAgICAgICAgICAgICAgICAgIFJlc3RhcnQtQ29tcHV0
ZXIgLUZvcmNlDQogICAgICAgICAgICAgICAgfQ0KDQogICAgICAgICAgICB9DQogICAgICAgIH0NCiAgICB9DQoNCiAgICAjI1JlLWVuYWJsZSBXQU0NCiAg
ICBzZXR4IE1TQUxfRk9SQ0VfV0FNIDENCg0KICAgICMjRGlzY29ubmVjdCBmcm9tIEdyYXBoIChzaWxlbnRseSBlcnJvciBpZiBub3QgY29ubmVjdGVkKQ0K
ICAgIHRyeSB7DQogICAgICAgIERpc2Nvbm5lY3QtTWdHcmFwaCAtRXJyb3JBY3Rpb24gU2lsZW50bHlDb250aW51ZQ0KICAgIH0gY2F0Y2ggew0KICAgICAg
ICBXcml0ZS1Ib3N0ICJOb3QgY29ubmVjdGVkIHRvIEdyYXBoIg0KICAgIH0NCg0KfQ0KDQoNCg0KDQoNCiMjU2NyaXB0IGVuZHMNCg0KIyBTSUcgIyBCZWdp
biBzaWduYXR1cmUgYmxvY2sNCiMgTUlJb1VBWUpLb1pJaHZjTkFRY0NvSUlvUVRDQ0tEMENBUUV4RHpBTkJnbGdoa2dCWlFNRUFnRUZBREI1Qmdvcg0KIyBC
Z0VFQVlJM0FnRUVvR3N3YVRBMEJnb3JCZ0VFQVlJM0FnRWVNQ1lDQXdFQUFBUVFIOHc3WUZsTENFNjNKTkxHDQojIEtYN3pVUUlCQUFJQkFBSUJBQUlCQUFJ
QkFEQXhNQTBHQ1dDR1NBRmxBd1FDQVFVQUJDRHl0YTNCT3hZMU1VLzcNCiMgU2Rra2U1MjhFbVE2Slo3c3JoU0hoV0Fvc3VsRGpxQ0NJVTB3Z2dXTk1JSUVk
YUFEQWdFQ0FoQU9teGlPK2RBdA0KIyA1Ky9iVU9JSVFCaGFNQTBHQ1NxR1NJYjNEUUVCREFVQU1HVXhDekFKQmdOVkJBWVRBbFZUTVJVd0V3WURWUVFLDQoj
IEV3eEVhV2RwUTJWeWRDQkpibU14R1RBWEJnTlZCQXNURUhkM2R5NWthV2RwWTJWeWRDNWpiMjB4SkRBaUJnTlYNCiMgQkFNVEcwUnBaMmxEWlhKMElFRnpj
M1Z5WldRZ1NVUWdVbTl2ZENCRFFUQWVGdzB5TWpBNE1ERXdNREF3TURCYQ0KIyBGdzB6TVRFeE1Ea3lNelU1TlRsYU1HSXhDekFKQmdOVkJBWVRBbFZUTVJV
d0V3WURWUVFLRXd4RWFXZHBRMlZ5DQojIGRDQkpibU14R1RBWEJnTlZCQXNURUhkM2R5NWthV2RwWTJWeWRDNWpiMjB4SVRBZkJnTlZCQU1UR0VScFoybEQN
CiMgWlhKMElGUnlkWE4wWldRZ1VtOXZkQ0JITkRDQ0FpSXdEUVlKS29aSWh2Y05BUUVCQlFBRGdnSVBBRENDQWdvQw0KIyBnZ0lCQUwvbWtITm8zcnZrWFVv
OE1DSXdhVFBzd3FjbExza2hQZktLMkZuQzRTbW5QVmlyZHByTnJuc2JoQTNFDQojIE1CL3pHNlE0RnV0V3hwZHRIYXV5ZWZMS0VkTGtYOVlGUEZJUFVoL0du
aFdsZnI2ZnFWY1dXVlZ5cjJpVGNNS3kNCiMgdW5XWmFuTXlsTkVRUkJBdTM0THpCNFRtZER0dGNlSXREQnZ1SU5YSklCMWpLUzNPN0Y1T3lKUDRJV0diTk9z
Rg0KIyB4bDdzV3hxODY4blB6YXcwUUYreGVtYnVkOGhJcUdaWFY1OVVXSTRNSzdkUHB6RFpWdTdLZTEzanJjbFBYdVUxDQojIDV6SEwycE5lM0k2UGdOcTJr
WmhBa0huRGVNZTJzY1MxYWhnNEF4Q04yTlEzcEM0RmZZajFnajRRa1hDclZZSkINCiMgTXRmYkJITXFicEVCZkNGTTFMeXVHd04xWFhobTJUb3hSSm96UUw4
STExcEpwTUxtcWFCbjNhUW52S0ZQT2JVUg0KIyBXQmYzSkZ4R2oyVDN3V21JZHBoMlBWbGRRbmFIaVpkcGVranc0S0lTRzJhYWRNcmVTeDduRG1PdTV0VHZr
cEk2DQojIG5qM2NBT1JGSlltMm1rUVpLMzdBbExUU1lXM3JNOW5GMzBzRUFNeDlISlhEai9jaHNySVJ0N3QvOHRXTWNDeEINCiMgWUtxeFl4aEVsUnAyWW43
MmdMRDc2R1NtTTlHSkIrRzl0K1pEcEJpNHBuY0I0UStVRENFZHNsUXBKWWxzNVE1Uw0KIyBVVWQwdmlhc3RrRjEzbnFzWDQwL3lielRRUkVTVytVUVVPc3h4
Y3B5RmlJSjMzeE1kVDlqN0NGZnhDQlJhMit4DQojIHE0YUxUOExXUlYrZElQeWhIc1hBajZLeGZnb21tZlhrYVMrWUhTMzEyYW15SGVVYkFnTUJBQUdqZ2dF
Nk1JSUINCiMgTmpBUEJnTlZIUk1CQWY4RUJUQURBUUgvTUIwR0ExVWREZ1FXQkJUczErT0MwbkZkWkV6ZkxtYy81N3FZcmh3UA0KIyBUekFmQmdOVkhTTUVH
REFXZ0JSRjY2S3Y5SkxMZ2pFdFVZdW5weUdkODIzSUR6QU9CZ05WSFE4QkFmOEVCQU1DDQojIEFZWXdlUVlJS3dZQkJRVUhBUUVFYlRCck1DUUdDQ3NHQVFV
RkJ6QUJoaGhvZEhSd09pOHZiMk56Y0M1a2FXZHANCiMgWTJWeWRDNWpiMjB3UXdZSUt3WUJCUVVITUFLR04yaDBkSEE2THk5allXTmxjblJ6TG1ScFoybGpa
WEowTG1Odg0KIyBiUzlFYVdkcFEyVnlkRUZ6YzNWeVpXUkpSRkp2YjNSRFFTNWpjblF3UlFZRFZSMGZCRDR3UERBNm9EaWdOb1kwDQojIGFIUjBjRG92TDJO
eWJETXVaR2xuYVdObGNuUXVZMjl0TDBScFoybERaWEowUVhOemRYSmxaRWxFVW05dmRFTkINCiMgTG1OeWJEQVJCZ05WSFNBRUNqQUlNQVlHQkZVZElBQXdE
UVlKS29aSWh2Y05BUUVNQlFBRGdnRUJBSENndjBOYw0KIyBWZWM0WDZDamRCczl0aGJYOTc5WEI3MmFyS0dITE95Rlhxa2F1eUw0aHhwcFZDTHRwSWgzYmIw
YUZQUVRTbm92DQojIExiYzQ3L1QvZ0xuNG9mZnljdDRrdkZJRHlFN1FLdDc2TFZiUCtmVDNyREI2bW91eVh0VFAwVU5FbTBNaDY1WnkNCiMgb1VpMG1jdWRU
NmNHQXhOM0owVFU1My9vV2Fqd3Z5OExwdW55TkR6czl3UEhoNmpTVEVBWk5VWnFhVlN3dUtGVw0KIyBqdXlrMVQzb3NkejlITmowZDFwY1ZJeHY3NkZRUGZ4
MkNXaUVuMi9LMnlDTk5XQWNBZ1BMSUxDc1dLQU9RR1BGDQojIG1DTEJzbG4xVld2UEo2dHNkczV2SXkzMGZuRnFJMnNpL3hLNFZDMG5mdGc2MmZDMmg1YjlX
OUZjckJqRFRaOXoNCiMgdHdHcG4xZXFYaWppdVpRd2dnYXdNSUlFbUtBREFnRUNBaEFJclVDeVlOS2NUSjllemFtOWs2N1pNQTBHQ1NxRw0KIyBTSWIzRFFF
QkRBVUFNR0l4Q3pBSkJnTlZCQVlUQWxWVE1SVXdFd1lEVlFRS0V3eEVhV2RwUTJWeWRDQkpibU14DQojIEdUQVhCZ05WQkFzVEVIZDNkeTVrYVdkcFkyVnlk
QzVqYjIweElUQWZCZ05WQkFNVEdFUnBaMmxEWlhKMElGUnkNCiMgZFhOMFpXUWdVbTl2ZENCSE5EQWVGdzB5TVRBME1qa3dNREF3TURCYUZ3MHpOakEwTWpn
eU16VTVOVGxhTUdreA0KIyBDekFKQmdOVkJBWVRBbFZUTVJjd0ZRWURWUVFLRXc1RWFXZHBRMlZ5ZEN3Z1NXNWpMakZCTUQ4R0ExVUVBeE00DQojIFJHbG5h
VU5sY25RZ1ZISjFjM1JsWkNCSE5DQkRiMlJsSUZOcFoyNXBibWNnVWxOQk5EQTVOaUJUU0VFek9EUWcNCiMgTWpBeU1TQkRRVEV3Z2dJaU1BMEdDU3FHU0li
M0RRRUJBUVVBQTRJQ0R3QXdnZ0lLQW9JQ0FRRFZ0QzlDMENpdA0KIyBlTGRkMVRsWkc3R0lRdlV6ak9zOWdaZHd4YnZFaFNZd242U09hTmhjOWVzMEpBZmhT
MC9UZUVQMEY5Y2Uydm5TDQojIDFXY2FVazhPb1ZmOGlKbkJrY3lCQXo1TmNDUmtzNDNpQ0gwMGZVeUFWeEpyUTVxWjhzVTdIL0x2eTBkYUU2Wk0NCiMgc3dF
Z0pmTVEwNHV5K3dqd2l1Q2RDY0JscC9xWWdFazFoejFSR2VpUUlYaEZMcUdmTE9FWXdock14ZTZUU1hCQw0KIyBNby83eHVvYzgyVm9rYUpOVElJUlNGSm8z
aEM5RkZkZDZCZ1RaY1Yvc2srRkxFaWtWb1ExMXZrdW5Lb0FGZEUzDQojIC9ob0dsTUo4eU9vYk11Ykt3dlNub3dNT2RLV3ZPYmFyWUJMajZOYTU5ekhoM0sz
a0dLRFl3U05IUjdPaEQyNmoNCiMgcTIyWUJvTWJ0MnBuTGRLOVJCcVNFSUdQc0RzSjE4ZWJNbHJDLzJwZ1ZJdEp3WlB0NGJSYzRHL3JKdm1NMWJMNQ0KIyBP
QkRtNnM2UjliN1QrMitUWVRSY3ZKTkZLSU0yS21Zb1g3Qnp6b3NtSlFheWc5UmM5aFVaVE8xaTRGNHo4dWpvDQojIDdBcW5zQU1ya2JJMmViNzNyUWdlZGFa
bHpMdmpTRkR6ZDVFYS90dFFva2JJWVZpWTlYd0NGanlES0swNWh1elUNCiMgdHcxVDBQaEg1blV3amV3d2szWVVwbHRMWFhSaFRUOFNrWGJldjFqTGNoQXBR
ZkRWeFcwbWRtZ1JRUk5ZbXR3bQ0KIyBLd0gwaVUxWjIzalBnVW8rUUVkZnlZRlFjNFVRSXlGWllJcGtWTUhNSVJyb09CbDhaaHpOZURoRk1KbFAvMk5QDQoj
IFRMdXFEUWhUUVh4WVBVZXorcmJzakRJSkFzeHNQQXhXRVFJREFRQUJvNElCV1RDQ0FWVXdFZ1lEVlIwVEFRSC8NCiMgQkFnd0JnRUIvd0lCQURBZEJnTlZI
UTRFRmdRVWFEZmc2N1k3K0Y4Umh2ditZWHNJaUdYMFRrSXdId1lEVlIwag0KIyBCQmd3Rm9BVTdOZmpndEp4WFdSTTN5NW5QK2U2bUs0Y0QwOHdEZ1lEVlIw
UEFRSC9CQVFEQWdHR01CTUdBMVVkDQojIEpRUU1NQW9HQ0NzR0FRVUZCd01ETUhjR0NDc0dBUVVGQndFQkJHc3dhVEFrQmdnckJnRUZCUWN3QVlZWWFIUjAN
CiMgY0RvdkwyOWpjM0F1WkdsbmFXTmxjblF1WTI5dE1FRUdDQ3NHQVFVRkJ6QUNoalZvZEhSd09pOHZZMkZqWlhKMA0KIyBjeTVrYVdkcFkyVnlkQzVqYjIw
dlJHbG5hVU5sY25SVWNuVnpkR1ZrVW05dmRFYzBMbU55ZERCREJnTlZIUjhFDQojIFBEQTZNRGlnTnFBMGhqSm9kSFJ3T2k4dlkzSnNNeTVrYVdkcFkyVnlk
QzVqYjIwdlJHbG5hVU5sY25SVWNuVnoNCiMgZEdWa1VtOXZkRWMwTG1OeWJEQWNCZ05WSFNBRUZUQVRNQWNHQldlQkRBRURNQWdHQm1lQkRBRUVBVEFOQmdr
cQ0KIyBoa2lHOXcwQkFRd0ZBQU9DQWdFQU9pTkVQWTBJZHU2UHZEcVowMWJnQWhxbCtFZzA4eXkyNW5SbTk1UnlzUURLDQojIHIyd3dKeE1TbnBCRW4wdjlu
cU44SnRVM3ZEcGRTRzJWMVQ5SjlDZTdGb0ZGVVAyY3ZiYUY0SForTjNITEl2ZGENCiMgcXBEUDlaTnE0K3NnMGRWUWVZaWFpb3JCdHIyaFNCaCszTmlBR2hF
WkdNMWhtWUZXOXNuamR1ZkU1QnRmUS9nKw0KIyBsUDkyT1QyZTFKblBTdDBvNjE4bW9aVllTTlVhL3RjblAvMlEwWGFHM1J5d1lGenpEYWp1NEltaHZUbmhP
RTdhDQojIGJyczJuZnZsSVZOYXc4cnBhdkdpUHR0RHVEUElUemdVa3BuMTNjNVViZGxkQWhRZlFETjhBK0tWc3NJaGRYTlMNCiMgeTBiWXhEUWNvcVZMamMx
dmRqY3NoVDhhemlicEdMNlFCN0JEZjVXSUlJSnc4TXpLNy8wcE5Wd2ZpVGhWOXplSw0KIyBpd21oeXd2cE1Sci9MaGxjT1hIaHZweW5DZ2JXSm1lM2t1Wk9Y
OTU2ckVuUExxUjBrcTNiUEtTY2hoL2p3VlliDQojIEt5UC9qN1hxaUh0d2ErYWd1djA2UDBXbXhPZ1drVktMUWNCSWhFdVdUYXRFUU9PTjhCVW96dTN4R0ZZ
SEtpOFENCiMgeEF3SVpEd3pqNjRvakR6TGo0Z0xEYjg3OU00ZWU0N3Z0ZXZMdC9CM0UrYm5LRCtzRXE2bEx5SnNRZm1DWEJWbQ0KIyB6R3dPeXNXR3cvWW1N
d3dIUzZEVEJ3SnFha0F3U0VzMHFGRWd1NjBiaFFqaVdRMXR5Z1ZRSytwS0hKNmwvYUNuDQojIEh3WjA1L0xXVXBEOXI0VklJZmxYTzdTY0ErMkdSZlMwWVc2
L2FPSW1ZSWJxeUsrcC9wUWQ1Mk1iT29aV2VFNHcNCiMgZ2dhME1JSUVuS0FEQWdFQ0FoQU54NnhYQmY4aG1TNUFReUlNT2ttR01BMEdDU3FHU0liM0RRRUJD
d1VBTUdJeA0KIyBDekFKQmdOVkJBWVRBbFZUTVJVd0V3WURWUVFLRXd4RWFXZHBRMlZ5ZENCSmJtTXhHVEFYQmdOVkJBc1RFSGQzDQojIGR5NWthV2RwWTJW
eWRDNWpiMjB4SVRBZkJnTlZCQU1UR0VScFoybERaWEowSUZSeWRYTjBaV1FnVW05dmRDQkgNCiMgTkRBZUZ3MHlOVEExTURjd01EQXdNREJhRncwek9EQXhN
VFF5TXpVNU5UbGFNR2t4Q3pBSkJnTlZCQVlUQWxWVA0KIyBNUmN3RlFZRFZRUUtFdzVFYVdkcFEyVnlkQ3dnU1c1akxqRkJNRDhHQTFVRUF4TTRSR2xuYVVO
bGNuUWdWSEoxDQojIGMzUmxaQ0JITkNCVWFXMWxVM1JoYlhCcGJtY2dVbE5CTkRBNU5pQlRTRUV5TlRZZ01qQXlOU0JEUVRFd2dnSWkNCiMgTUEwR0NTcUdT
SWIzRFFFQkFRVUFBNElDRHdBd2dnSUtBb0lDQVFDMGVESFRDcGhCY3I0OFJzQWNySFhibzBabw0KIyBkTFJSRjUxTnJZME5sTFdabG9Nc1ZPMURhaEdQTlJj
eWJFS3ErUnV3T25QaG9mNnB2RjR1R2p3anFOamZFdlVpDQojIDZ3dWltNWJhcCswbGdsb00yelg0a2Z0bjVCMUlwWXpUcXB5RlEvNEJ0MG1BeEFIZUhZTm5R
eHFYbVJpbnZ1TmcNCiMgeFZCZEprZjc3UzJ1UG9DajdHSDhCTHV4Qkc1QXZmdEJkc09FQ1MxVWt4QnZNZ0VkZ2tGaUROWWlPVHg0T3RpRg0KIyBjTVNrcVR0
RjJoZlF6M3pRU2t1MldzM0lmRFJlYjZlM21tZGdsVGNhYXJwczB3alVqc1p2a2dGa3JpSzl0VUtKDQojIG0vczgwRmlvY1NrMVZZTFpsRHdGdCtjVkZCVVJK
ZzZ6TVVqWmEvemJDY2xGODNiUlZGTGVHa3VBaEhpR1BNdlMNCiMgR21oZ2FUelZ5aFluNHAwKzh5OW9IUmFRVC9hb2ZFblM1eExyZnhuR3BUWGlVT2VTTHNK
eWdvTFBwNjZia0RYMQ0KIyBabEFlU3BRbDkyUU9NZVJ4eWt2cTZnYnlsc1hRc2tCQkJuR3kzdFcvQU1PTUNaSVZOU2F6N0JYOFZ0WUdxTHQ5DQojIE1tZU9y
ZUdQUmR0QngzeUdPUCtyeDNyS1dERUpsSXFMWHZKV25ZMHY1eWRQcE9qTDZzMzZjend6c3VjdW9LczcNCiMgWWsvZWhiLy9XeCs1a01xSU1SdlVCRHg2ejFl
dis3cHNOT2RnSk1vaXdPclVHMlpkU29RYlUyck1rcExpUTZiRw0KIyBSaW5aYkk0T0x1OUJNSUZtMVVVbDlWbmVQczZCYWFlRVd2akpTak5tMnFBK3NkRlVl
RVkwcVZqUEtPV3VnL0c2DQojIFg1dUFpeW5NN0J1MmF5QmpVd0lEQVFBQm80SUJYVENDQVZrd0VnWURWUjBUQVFIL0JBZ3dCZ0VCL3dJQkFEQWQNCiMgQmdO
VkhRNEVGZ1FVNzI5VFN1bmtCbng2eXVLUVZ2WXYxRW5zeTA0d0h3WURWUjBqQkJnd0ZvQVU3TmZqZ3RKeA0KIyBYV1JNM3k1blArZTZtSzRjRDA4d0RnWURW
UjBQQVFIL0JBUURBZ0dHTUJNR0ExVWRKUVFNTUFvR0NDc0dBUVVGDQojIEJ3TUlNSGNHQ0NzR0FRVUZCd0VCQkdzd2FUQWtCZ2dyQmdFRkJRY3dBWVlZYUhS
MGNEb3ZMMjlqYzNBdVpHbG4NCiMgYVdObGNuUXVZMjl0TUVFR0NDc0dBUVVGQnpBQ2hqVm9kSFJ3T2k4dlkyRmpaWEowY3k1a2FXZHBZMlZ5ZEM1ag0KIyBi
MjB2UkdsbmFVTmxjblJVY25WemRHVmtVbTl2ZEVjMExtTnlkREJEQmdOVkhSOEVQREE2TURpZ05xQTBoakpvDQojIGRIUndPaTh2WTNKc015NWthV2RwWTJW
eWRDNWpiMjB2UkdsbmFVTmxjblJVY25WemRHVmtVbTl2ZEVjMExtTnkNCiMgYkRBZ0JnTlZIU0FFR1RBWE1BZ0dCbWVCREFFRUFqQUxCZ2xnaGtnQmh2MXNC
d0V3RFFZSktvWklodmNOQVFFTA0KIyBCUUFEZ2dJQkFCZk8reGFBSFA0SFBSRjJjVEM5dmd2SXRUU21mODNRaDhXSUdqQi9UOE9iWEFaejhPanVoVXhqDQoj
IGFhRmRsZU1NMGxCcnlQVFFNMnFFSlBlMzZ6d2JTSS9tUzgzYWZzbDNZVGorSVFoUUU3alUva1hqanl0SmdubjANCiMgaHZyVjZocVdHZDNyTEFVdDZ2Snk5
bE1EUGpUTHhMZ1hmOXI1bldNUXdyOE15YjlyRVZLQ2hIeWZwemVlNWtIMA0KIyBGOEhBQkJncjBVZHFpclo3Ym93ZTlWajJBSU1EOGxpeXJ1a1oyaUEvd2RH
MnRoOXkxSXNBMFFGOGRUWHF2Y25UDQojIG1wZmVRaDM1azV6T0NQbVNOcTFVSDQxMEFOVmtvNDMrQ2RtdTR5ODFoamFqVi9neGRFa014MU5LVTR1SFFjS2YN
CiMgWnhBdkJBS3FNVnVxdGU2OU05SjZBNDdPdmdSYVBzKzJ5a2djR1YwMFRZcjJMcjN0eTlxSWlqYW5yVVIzYW56RQ0KIyB3bHZ6WmlpeWZUUGpMYm5GUnNq
c1lnMzlPbFY4Y2lwRG9xNytxTk5qcUZ6ZUd4Y3l0TDVUVExMNFphb0JkcWJoDQojIE9oWjNaUkRVcGhQdlNSbU1UaGkwdnc5dk9EUnpXNkF4bkpsbDM4RjBj
dUpHN3VFQllUcHRNU2JoZGhHUURwT1gNCiMgZ3BJVXNXVGpkNnhwUjZvYVFmL0RKYmczczZLQ0xQQWxaNjZSeklnOXNDK05KcHVkL3Y0KzdSV3NXQ2lLaTlF
Tw0KIyBMTEhmTVIyWnlKLyt4aEN4OXlIYnh0bDVUUGF1MWovMU1JRHBNUHgwTGNrVGV0aVN1RXRRdkxzTnozUWJwN3dHDQojIFdxYklpT1dDbmI1V3F4TDMv
QkFQdklYS1VqUFN4eVpzcThXaGJhTTJ0c3pXa1BaUHViZGNNSUlHN1RDQ0JOV2cNCiMgQXdJQkFnSVFDb0R2R0V1TjhRV0MwY1IycDVWMGFEQU5CZ2txaGtp
Rzl3MEJBUXNGQURCcE1Rc3dDUVlEVlFRRw0KIyBFd0pWVXpFWE1CVUdBMVVFQ2hNT1JHbG5hVU5sY25Rc0lFbHVZeTR4UVRBL0JnTlZCQU1UT0VScFoybERa
WEowDQojIElGUnlkWE4wWldRZ1J6UWdWR2x0WlZOMFlXMXdhVzVuSUZKVFFUUXdPVFlnVTBoQk1qVTJJREl3TWpVZ1EwRXgNCiMgTUI0WERUSTFNRFl3TkRB
d01EQXdNRm9YRFRNMk1Ea3dNekl6TlRrMU9Wb3dZekVMTUFrR0ExVUVCaE1DVlZNeA0KIyBGekFWQmdOVkJBb1REa1JwWjJsRFpYSjBMQ0JKYm1NdU1Uc3dP
UVlEVlFRREV6SkVhV2RwUTJWeWRDQlRTRUV5DQojIE5UWWdVbE5CTkRBNU5pQlVhVzFsYzNSaGJYQWdVbVZ6Y0c5dVpHVnlJREl3TWpVZ01UQ0NBaUl3RFFZ
SktvWkkNCiMgaHZjTkFRRUJCUUFEZ2dJUEFEQ0NBZ29DZ2dJQkFOQkdyQzBTeHA3UTZxNWdWck1yVjdwdlVmK0djQW9CMzhvMw0KIyB6QmxDTUdNeXFKbmZG
Tlp4K3d2QTY5SEZUQmR3Ykh3QlNPZUxwdlBuWjhaTit2bzhkRTIvcFB2T3gvVmo4VGNoDQojIFR5U0EyUjRRS3BWRDdkdk5aaDZ3VzJSNmtTdTlSSnQvNFFo
Z3VTc3NwM3FvbWU3TXJ4VnlmUU85c014NlpBV2oNCiMgRkRZT3pEaThTT2hQVVdsTG5oMDBDbGw4cGpyVWNDVjNLM0UwenowOWxkUS8vbkJaWlJFcjRoL0dJ
NkR4YjJVbw0KIyB5ck4waWp0VURWSFJYZG1uY09PTUEzQ29CL2lVU1JPVUlORFQ5OG9rc291VE1ZRk9uSG9SaDYrODZMdGM1empQDQojIEtIVzVLcUN2cFNk
dVN3aHdVbW90dVFoY2c5dHcyWUQzdzZ5U1NTdSszcVU4REQrbmlnTkpGbXQ2TEFIdkgzS1MNCiMgdU5Mb1pMYzFIZjJKTk1WTDRRMU9wYnlicE1lNDZZY2VO
QTBMZk5zbnFjbnBKZUl0Sy9EaEtiUHhUVHVHb1g3dw0KIyBKTmRvUk9SVmJQUjFWVm5EdVNlSFZabGM0c2VBTys2ZDJzQzI2L1BRUGRQNTFobzF6QnAreFVJ
WmtwU0ZBOHZXDQojIGRvVW9ITFducVdVM2RDQ3lGRzFyb1NyZ0hqU0hscTh4eW1MbmpDYlNMWjQ5a1BtazhpeXlpek5ESVhqLy9jT2cNCiMgclk3cmxSeVRs
YUNDZnc3YVNVUk93bnU3ekVSNkVhSitBbGlMN29qVGRTNVBXUHNXZXVwV3M3TnBDaFVrNTU1Sw0KIyAwOTZWMWhFMHlaSVhlK2dpQXdXMDBhSHpyRGNoSWMy
YlFocHAwSW9LUlI3WXVmQWtwcnhNaVhBSlExWENtbkNmDQojIGdQZjgrM21uQWdNQkFBR2pnZ0dWTUlJQmtUQU1CZ05WSFJNQkFmOEVBakFBTUIwR0ExVWRE
Z1FXQkJUa08venkNCiMgTWUzOS9kZnprWEZqR1ZCRHoyR002REFmQmdOVkhTTUVHREFXZ0JUdmIxTks2ZVFHZkhySzRwQlc5aS9VU2V6TA0KIyBUakFPQmdO
VkhROEJBZjhFQkFNQ0I0QXdGZ1lEVlIwbEFRSC9CQXd3Q2dZSUt3WUJCUVVIQXdnd2daVUdDQ3NHDQojIEFRVUZCd0VCQklHSU1JR0ZNQ1FHQ0NzR0FRVUZC
ekFCaGhob2RIUndPaTh2YjJOemNDNWthV2RwWTJWeWRDNWoNCiMgYjIwd1hRWUlLd1lCQlFVSE1BS0dVV2gwZEhBNkx5OWpZV05sY25SekxtUnBaMmxqWlhK
MExtTnZiUzlFYVdkcA0KIyBRMlZ5ZEZSeWRYTjBaV1JITkZScGJXVlRkR0Z0Y0dsdVoxSlRRVFF3T1RaVFNFRXlOVFl5TURJMVEwRXhMbU55DQojIGREQmZC
Z05WSFI4RVdEQldNRlNnVXFCUWhrNW9kSFJ3T2k4dlkzSnNNeTVrYVdkcFkyVnlkQzVqYjIwdlJHbG4NCiMgYVVObGNuUlVjblZ6ZEdWa1J6UlVhVzFsVTNS
aGJYQnBibWRTVTBFME1EazJVMGhCTWpVMk1qQXlOVU5CTVM1ag0KIyBjbXd3SUFZRFZSMGdCQmt3RnpBSUJnWm5nUXdCQkFJd0N3WUpZSVpJQVliOWJBY0JN
QTBHQ1NxR1NJYjNEUUVCDQojIEN3VUFBNElDQVFCbEtxM3hIQ2NFdWE1Z1FlelJDRVNlWTBCeUlmams5aUpQMnpXTHBRcTFiNFVSR253V0JkRVoNCiMgRDln
QnE5Zk5hTm1GajZFaDgvWW1SRGZ4VDdDMGs4RlVGcU5oK3RzaGdiNE82TGdqZzhLOGVsQzQrb1dDcW5VLw0KIyBNTDlsRmZpbTgvOXlKbVpTZTJGOEFRL1Vk
S0ZPdGo3WU1UbXFQTzltenNrZ2lDM1FZSVVQMlMzSFF2SEcxRkR1DQojICtXVXFXNGRhSXFUb1hGRS9KUS9FQUJnZlpYTFdVMHppVE42UjN5Z1FCSE1VQmFC
NWJkclBiRjZNUllzMDNoNG8NCiMgYkVNbnhZT1g4VkJSS2UxdU5uelFWVGVMbmkybkhrWC9RcXZYbk5iK1lrREZreFVHdE1UYWlMUjl3anhVeHUyaA0KIyBF
Q1pwcXlVMWQwSWJYNldxOC9nVnV0RG9qQklGZVJscUFjdUVWVDBjS3NiK3pKTkVzdUVCN083L2N1dlRRYXNuDQojIE05QVdjSVFmVmpuenJ2d2lDWjg1RUU4
TFVrcVJob1MzWTUwT0hnYVk3VC9sd2Q2VUFyYitCT1ZBa2cyb092b2wNCiMgL0RKZ2RkSjM1WFR4ZlVsUSs4SGdndDhsMll2N3JvYW5jSklGY2JvakJjeGxS
Y0dHMExJaHA2R3ZSZVFHZ01nWQ0KIyB4UWJWMVMzQ3JXcVp6QnQxUjl4SmdLZjQ3Q2R4VlJkL25kVWxRMDVveFl5MnpSV1ZGakY3bWNyNEMzNE1qM29jDQoj
IENWY2NBdmxLVjlqRW5zdHJuaUx2VXh4VlpFL3JwdGI3SVJFMmxza0tQSUpnYmFQNXQybkdqL1VMTGk0OXhUY0INCiMgWlU4YXR1ZmsrRU1GL2NXdWlDN1BP
R1Q3NXFhTDZ2ZEN2SGxzaHRqZE5YT0NJVWpzYXJmTlp6Q0NCMXN3Z2dWRA0KIyBvQU1DQVFJQ0VBaXhuODJ6MnZPd01WVllDQUV2QU9rd0RRWUpLb1pJaHZj
TkFRRUxCUUF3YVRFTE1Ba0dBMVVFDQojIEJoTUNWVk14RnpBVkJnTlZCQW9URGtScFoybERaWEowTENCSmJtTXVNVUV3UHdZRFZRUURFemhFYVdkcFEyVnkN
CiMgZENCVWNuVnpkR1ZrSUVjMElFTnZaR1VnVTJsbmJtbHVaeUJTVTBFME1EazJJRk5JUVRNNE5DQXlNREl4SUVOQg0KIyBNVEFlRncweU16RXhNVFV3TURB
d01EQmFGdzB5TmpFeE1UY3lNelU1TlRsYU1HTXhDekFKQmdOVkJBWVRBa2RDDQojIE1SUXdFZ1lEVlFRSEV3dFhhR2wwYkdWNUlFSmhlVEVlTUJ3R0ExVUVD
aE1WUVU1RVVrVlhVMVJCV1V4UFVpNUQNCiMgVDAwZ1RGUkVNUjR3SEFZRFZRUURFeFZCVGtSU1JWZFRWRUZaVEU5U0xrTlBUU0JNVkVRd2dnSWlNQTBHQ1Nx
Rw0KIyBTSWIzRFFFQkFRVUFBNElDRHdBd2dnSUtBb0lDQVFERHBHSkM2Y3pSK0dOWUZGeGUvZmJoZEFxOEZ2eTVudXUrDQojIHZndldtVGlXTTZhdC93eUd2
TmFGci9XK0c5RnNDNlNWYnRIenBBSnNTT3FITHVMbit0ZDR3ekZ0Qm4xZUhVYUgNCiMgYnJhOG43ZzdvcmVLNTNieVFPZ3lMTkdCdWNUWlNrNUdQQUNMd1Q5
eU1CTTFBOVgrZXllUm9nS2F4eG5xSE9GTA0KIyBiY3hMaGdOOGtxcGJCaElOSUFub1ZpYzUxSklkOGpQRjI1TEF0QzdnWnAyUDNXU2YvSkxRc0FuZC9JSDJS
RHZWDQojIEF3MHBJbnVGVTJOMCsxUlcwNG1oOUc4UGdMMzNFRmN0Z2tzSk1INTVIMkdvRWhaQ21xL2pHTUx1NEtsVjhhNGQNCiMgMWZ4bzcycGVqM1ROQU94
SEU2cHM2d2tiYjVGaUVlbTZjL3R3Q0IraGErc2s3aHQxNGl5QytyQ3Y0aGYvWGVGTg0KIyBqNGg5YnlmOFgzWVJvOUswTi96UVViRkFRdDVkY09OUythdlZG
OVRvZFpVOVRyaWVvVmY3bXA1T2lXTjQ2WnZqDQojIG4yZTJBa3hkaDVNK2N1b2ZVKzdHTkMwNHV2RlpyY1d2eElCTFJ1aVZUVmJLaisxc0JKRUVjYnY5OUty
WThxRi8NCiMgSjgwcmhlMDVyRVlKbWRVZ2ZpRW5KWG83cWt6WFlYTW5BNFl0L3lDRW9GU1RVVXZ4ZW1mbFVCbjkzNGVqbTNVQw0KIyAxY0tFOUNaeVkydy9E
OHlkZGpxQ29GWWswSVozV21XNUg2WWxZbnlkYnMxaWEwMXVjQkt4L3FyN3JSMWJlUDdCDQojIEdRRlV1RHpYem5DVjBkTE5QeStTUjVJN1RPR2hwZXRud0I0
eDQvU2JBcGJySStPM0UrbzBUaUtrQ09Ibzg5YkoNCiMgVHFISlRuclFqUUlEQVFBQm80SUNBekNDQWY4d0h3WURWUjBqQkJnd0ZvQVVhRGZnNjdZNytGOFJo
dnYrWVhzSQ0KIyBpR1gwVGtJd0hRWURWUjBPQkJZRUZOQjNUaFh6OFd2V0htK1R1U2ZWQklpWkt5dmZNRDRHQTFVZElBUTNNRFV3DQojIE13WUdaNEVNQVFR
Qk1Da3dKd1lJS3dZQkJRVUhBZ0VXRzJoMGRIQTZMeTkzZDNjdVpHbG5hV05sY25RdVkyOXQNCiMgTDBOUVV6QU9CZ05WSFE4QkFmOEVCQU1DQjRBd0V3WURW
UjBsQkF3d0NnWUlLd1lCQlFVSEF3TXdnYlVHQTFVZA0KIyBId1NCclRDQnFqQlRvRkdnVDRaTmFIUjBjRG92TDJOeWJETXVaR2xuYVdObGNuUXVZMjl0TDBS
cFoybERaWEowDQojIFZISjFjM1JsWkVjMFEyOWtaVk5wWjI1cGJtZFNVMEUwTURrMlUwaEJNemcwTWpBeU1VTkJNUzVqY213d1U2QlINCiMgb0UrR1RXaDBk
SEE2THk5amNtdzBMbVJwWjJsalpYSjBMbU52YlM5RWFXZHBRMlZ5ZEZSeWRYTjBaV1JITkVOdg0KIyBaR1ZUYVdkdWFXNW5VbE5CTkRBNU5sTklRVE00TkRJ
d01qRkRRVEV1WTNKc01JR1VCZ2dyQmdFRkJRY0JBUVNCDQojIGh6Q0JoREFrQmdnckJnRUZCUWN3QVlZWWFIUjBjRG92TDI5amMzQXVaR2xuYVdObGNuUXVZ
Mjl0TUZ3R0NDc0cNCiMgQVFVRkJ6QUNobEJvZEhSd09pOHZZMkZqWlhKMGN5NWthV2RwWTJWeWRDNWpiMjB2UkdsbmFVTmxjblJVY25Weg0KIyBkR1ZrUnpS
RGIyUmxVMmxuYm1sdVoxSlRRVFF3T1RaVFNFRXpPRFF5TURJeFEwRXhMbU55ZERBSkJnTlZIUk1FDQojIEFqQUFNQTBHQ1NxR1NJYjNEUUVCQ3dVQUE0SUNB
UUJKRVlkajhESXNxMnI2K3VtY01PajQ1ZHVDc3czQm1ERXMNCiMgVkNqaEZPRzhwSERpb0wwdWxkczNtU0xOQS82S25MZDRRU2t4aEVra0ZrZ1B5Wks5UkRB
dE9LaVJ6djA4Sy83QQ0KIyBRd2d3VExTZlZ3TFR1K1NmcUtnM0hEUG9QRDZQbzQ0YW1DY3lyMjRyRlZMM2JENmhaWmVlYjBzMGJ4eGdBT29ZDQojIDhnOG1w
QmkwVG5EV2toV1JuWURpdGdESVVQQkZDNnhQRVlxNHR3OVVUcnFocGxGaXFuRHZXcWp4d1grY0ZNSW0NCiMgdklQZk5MRTE2cmp1WU9FMXBHYWtyOExkUVJ6
SnJ1dlRFZXBhRVFDdGV4N3hYRW9uQ3VqNXRNNG5kU2t0cytKNw0KIyBSZ2VBbjNJUFBHaFM2SU4zSWo5clh6SXRzVTU2amRTdm1KbFBQZUQ1ZFpRcWRjUmIw
K3FhNXRkQWJPZ1ZZaWwzDQojIHczMVJjVjB3ZEpoNkF6YWhvdndmUHE4WDlzKzd1WDZGendzd1dFL2tOM01iYUtiMmJaTnZUS1UyUFBCaHcxU2INCiMgQlRB
WTErOHpPbWdqU3JFeWd0bzIvZGhmbzVaSFBISlFrRUo1ZDFPSTZzRzQrTHE5U1FUK1VNVUthMW9jSFh0UA0KIyBqSW53MWFkTHgyYkZXZEtaMkRBaWtodVcx
Szk4RGxPeEQ5OXRROEw3eExFNHdzdjF3NEw4dklxVk9zTnhvM0VFDQojIEJab0tmd1YvOUNRaDEvdE9FQ2RDWkVmVTB4V1VkUWZzNGR3bkN0V1JMVTdiN21Y
NDBONzRsOEpjTFRTOHVONzINCiMgM3RTZVZqVTdOQjR5VjdIM1c0V09aQk9VWWNuMnNSUWE5dFNtTWZtL25yMk1xSlNFdE1FZWpYc09oYWk0SEhPQg0KIyBu
dmVNbW1GVTZqR0NCbGt3Z2daVkFnRUJNSDB3YVRFTE1Ba0dBMVVFQmhNQ1ZWTXhGekFWQmdOVkJBb1REa1JwDQojIFoybERaWEowTENCSmJtTXVNVUV3UHdZ
RFZRUURFemhFYVdkcFEyVnlkQ0JVY25WemRHVmtJRWMwSUVOdlpHVWcNCiMgVTJsbmJtbHVaeUJTVTBFME1EazJJRk5JUVRNNE5DQXlNREl4SUVOQk1RSVFD
TEdmemJQYTg3QXhWVmdJQVM4QQ0KIyA2VEFOQmdsZ2hrZ0JaUU1FQWdFRkFLQ0JoREFZQmdvckJnRUVBWUkzQWdFTU1Rb3dDS0FDZ0FDaEFvQUFNQmtHDQoj
IENTcUdTSWIzRFFFSkF6RU1CZ29yQmdFRUFZSTNBZ0VFTUJ3R0Npc0dBUVFCZ2pjQ0FRc3hEakFNQmdvckJnRUUNCiMgQVlJM0FnRVZNQzhHQ1NxR1NJYjNE
UUVKQkRFaUJDQU5oQjdIY2pPN29RdXhiZG1KTlgyWUt2SjFWWHl0U004cA0KIyBLZ3c5Q01ucHZ6QU5CZ2txaGtpRzl3MEJBUUVGQUFTQ0FnQzQ4WEJiWk1M
VHR6UDRoZUhZU0ZoUWZiSnVISnR4DQojIGpNdldDM2JLd2hEU0ZwZDkvUVpZQjRHWnRRWGV2U292a2c4QURVT3FMLzQ2ZG12T0VWSGkwRXBheENnUitkQ0wN
CiMgVXJDQWtLRHVwcnp6cFd0TDdJNjlPNjBhOGRWN09EMjVYbjM5cTUydllLaXB6OWlNcnRaZHd1UUZqejBBd0dYMQ0KIyBSaVEyRWZzMmxhSnZ5MFgzTkg3
bHZmQWltTStvWm1Ed1ZSZHRTVmp3OGVJWWZJdFNTMVhXUnRLUDRVNHh6cXU0DQojIExRam1zeXNMUVFzbXZINVdaR09OTEx1U3AzdGhmZC8xY2NKMWpjOUdM
SVVVMUhIbThIVzZCS1diOFlsMHhYSUkNCiMgSUR3VGwvLy8zTXdQKzhsTHRzbnFtQVJxYnpUaVBId2RjTTVYeXA1VHllZkZkVVA1eURlZlZVWFIvRFl0ZnF1
Vg0KIyBCRkgwOEI0NzFNMzg4Mno1SUVMMXh4bUZ2akc4N0dza2NrQkMvMGIzVTlaUk5Bd014MEtGK2ZrQXlhZkxKZzR0DQojIFdoUVQ2OFVmS2t2ZFZ6WSs1
ODlVSzY4c3ZBMElwT0dBMHl2WURPelp6dlVMYVR5REhPSVVxb0grcFFyT1JYa3gNCiMgcXNzMWNMeTYyL24vNFpRbldwb1FCSUUxejdoZ3JmYmxYT2RXYTI3
N0FNL2RUMlBiYXJLOWhRTS8rZ1hNK24rUg0KIyB5NnJ4dk5UL3FmcWdrZkhIMk5zeG56N0NDN2FWbUVWRHl5VzR0NWdUQmJkRjc0eUJmZmdaQ2JiMVdkcnRL
VmkvDQojIDRuQ2VURVIvRU9BZms1bW16cW5pbmVWTW93WTRuTkNkbStGSytEbzlOdldOVytDQXBSMXpvdnBBamZTZFVFUkINCiMgNmV5YVB3cnFHcWh0N3FH
Q0F5WXdnZ01pQmdrcWhraUc5dzBCQ1FZeGdnTVRNSUlERHdJQkFUQjlNR2t4Q3pBSg0KIyBCZ05WQkFZVEFsVlRNUmN3RlFZRFZRUUtFdzVFYVdkcFEyVnlk
Q3dnU1c1akxqRkJNRDhHQTFVRUF4TTRSR2xuDQojIGFVTmxjblFnVkhKMWMzUmxaQ0JITkNCVWFXMWxVM1JoYlhCcGJtY2dVbE5CTkRBNU5pQlRTRUV5TlRZ
Z01qQXkNCiMgTlNCRFFURUNFQXFBN3hoTGpmRUZndEhFZHFlVmRHZ3dEUVlKWUlaSUFXVURCQUlCQlFDZ2FUQVlCZ2txaGtpRw0KIyA5dzBCQ1FNeEN3WUpL
b1pJaHZjTkFRY0JNQndHQ1NxR1NJYjNEUUVKQlRFUEZ3MHlOakExTWpFeE1qTXpNVEZhDQojIE1DOEdDU3FHU0liM0RRRUpCREVpQkNCN3V1MytKekZqL1A3
Y0dtOE94ekRPNUpMZVg3WUxOVXNYNFI0WWJQNU8NCiMgUlRBTkJna3Foa2lHOXcwQkFRRUZBQVNDQWdCQnBpTE5WOGNxL2tnN0VvTkt4a0I0MEZPeHhqbW03
akZiK1ljWA0KIyB2RExRaWFQNU1JNTFXMXVNNWgvQ0lUajg0NVc5Rlp6MTBEa0lJZW40eko1bStHNzFqaEswSm9wZHNEczNEYnRNDQojIGl0NnVJWjdvY09C
dnk0SEI2RG9nWmF3RERGUUdPYjNNOTIwQWQyVGUya1VmYkgrTnkyblV1Rmh0bDV0cDdCMGsNCiMgRCtJUHFYS1BONGxTM1dXUHBwSjdWSG4xODVLek5kSldN
TDFwMXAyMThUbGdPVmlRNG1VQ0RFcm5pejBhOHZkWg0KIyBVWEtFNmZ3WWRmSlpWKzRMR0JZN05JNjJuSWpyOWJYT2QwY1o4R3RYTS9TZXp3OU1BUmlYRjFW
TUh4RmVabnNWDQojIG1QQVNFbkpWbElrazd4bEwwUXBaZWE4ZXJ0cGpPSWs1YVMxWkpJMGlOMXljeTB3UUF0NEwraEhxSitEL0tQOWENCiMgYThyUHZzbURs
WmdpMmlZSjd5Y1Y0ajFFTlJzSmwxck1NZ1YyMmFlYWdjL2NSV1kyT3JPQjFJSjhpVjNtT3RzUg0KIyBMcmtsbjJmemUzb1R6cFZRSU5tVUZ5TmY1T0pqQXds
UFJtVUJZbGdSK2dremRyZUtCY3g1UFVXZWUzb1BhNnBqDQojIEZCWkY0QlJYSXp3ZjNZM2JFTTJiME4xUHRlSUNqK0FVT0xmOG11R2N3SEIydldjMFlQcTdM
Q3JvVnJmR1pRWmINCiMgWEhDZVdXbEtKUTZyb1JxVThtUFlRMXBPZk94OGppZTJnZkZod3IxdjZQaEE2NTZ0d2srNURsRnZqY2MzVHNBZg0KIyArRC83WEVr
aVFNcnVid01Rei93bFQ0dlV3Sk1SMEpuZmdSWE5FSkdnS3RieG81TEZvUkVOdkVwK0dJcVRXWHhzDQojIG93YVhiZz09DQojIFNJRyAjIEVuZCBzaWduYXR1
cmUgYmxvY2sNCg==
'@ -replace '\s',''
$script:ApEmbeddedScripts['Get-AutopilotDiagnosticsCommunity.ps1'] = @'
DQo8I1BTU2NyaXB0SW5mbw0KDQouVkVSU0lPTiA2LjMNCi5HVUlEIGI0NTYwNWI2LTY1YWEtNDVlYy1hMjNjLWY1MjkxZjlmYjUxOQ0KLkFVVEhPUiBBbmRy
ZXdUYXlsb3IsIE1pY2hhZWwgTmllaGF1cyAmIFN0ZXZlbiB2YW4gQmVlaw0KLkNPTVBBTllOQU1FDQouQ09QWVJJR0hUIEdQTA0KLlRBR1MNCi5MSUNFTlNF
VVJJIGh0dHBzOi8vZ2l0aHViLmNvbS9hbmRyZXctcy10YXlsb3IvcHVibGljL2Jsb2IvbWFpbi9MSUNFTlNFDQouUFJPSkVDVFVSSSBodHRwczovL2dpdGh1
Yi5jb20vYW5kcmV3LXMtdGF5bG9yL3B1YmxpYw0KLklDT05VUkkNCi5FWFRFUk5BTE1PRFVMRURFUEVOREVOQ0lFUyANCi5SRVFVSVJFRFNDUklQVFMNCi5F
WFRFUk5BTFNDUklQVERFUEVOREVOQ0lFUw0KLlJFTEVBU0VOT1RFUw0KVmVyc2lvbiA2LjM6IEJ1ZyBmaXhlcy4NClZlcnNpb24gNi4yOiBCdWcgZml4ZXMu
DQpWZXJzaW9uIDYuMTogQnVnIGZpeGVzLg0KVmVyaXNvbiA2LjA6IEFkZGVkIEFQdjIgc3VwcG9ydCBhbmQgdmFyaW91cyBvdGhlciBlbmhhbmNlbWVudHMN
ClZlcnNpb24gNS4xNDogQWRkZWQgYXV0by1pbnN0YWxsIG9mIEdyYXBoIG1vZHVsZXMNClZlcnNpb24gNS4xMzogRml4ZWQgaXNzdWUgd2l0aCBiZWFyZXIN
ClZlcnNpb24gNS4xMjogUmVtb3ZlZCBhbGwgY29tbWFuZGxldHMgYW5kIGFkZGVkIGJlYXJlciBwYXJhbQ0KVmVyc2lvbiA1LjExOiBBZGRlZCBsb2dpYyBh
cm91bmQgRGV2aWNlIFJlZ2lzdHJhdGlvbiBldmVudCBsb2cNClZlcnNpb24gNS4xMDogQWRkaXRpb25hbCBsb2dpYyBmb3IgRE8gZG93bmxvYWRzLCBNU0kg
cHJvZHVjdCBuYW1lcw0KVmVyc2lvbiA1Ljk6IENvZGUgU2lnbmVkDQpWZXJzaW9uIDUuNzogRml4ZWQgTGFzdExvZ2dlZFN0YXRlIGZvciBXaW4zMkFwcHMg
YW5kIEFkZGVkIHN1cHBvcnQgZm9yIG5ldyBHcmFwaCBNb2R1bGUNClZlcnNpb24gNS42OiBGaXhlZCBwYXJhbWV0ZXIgaGFuZGxpbmcNClZlcnNpb24gNS41
OiBBZGRlZCBzdXBwb3J0IGZvciBhIHppcCBmaWxlDQpWZXJzaW9uIDUuNDogQWRkZWQgYWRkaXRpb25hbCBFU1AgZGV0YWlscw0KVmVyc2lvbiA1LjM6IEFk
ZGVkIGhhcmR3YXJlIGFuZCBPUyB2ZXJzaW9uIGRldGFpbHMNClZlcnNpb24gNS4yOiBBZGRlZCBkZXZpY2UgcmVnaXN0cmF0aW9uIGV2ZW50cw0KVmVyc2lv
biA1LjE6IEJ1ZyBmaXhlcw0KVmVyc2lvbiA1LjA6IEJ1ZyBmaXhlcw0KVmVyc2lvbiA0Ljk6IEJ1ZyBmaXhlcw0KVmVyc2lvbiA0Ljg6IEFkZGVkIERlbGl2
ZXJ5IE9wdGltaXphdGlvbiByZXN1bHRzIChidXQgbm90IHdoZW4gdXNpbmcgYSBDQUIgZmlsZSksIGVuc3VyZWQgZXZlbnRzIGFyZSBkaXNwbGF5ZWQgZXZl
biB3aGVuIG5vIEVTUA0KVmVyc2lvbiA0Ljc6IEFkZGVkIEVTUCBzZXR0aW5ncywgZml4ZWQgYnVncw0KVmVyc2lvbiA0LjY6IEZpeGVkIHR5cG8NClZlcnNp
b24gNC41OiBGaXhlZCBidXQgdG8gcHJvcGVybHkgcmVwb3J0ZWQgV2luMzIgYXBwIHN0YXR1cyB3aGVuIGEgV2luMzIgYXBwIGlzIGluc3RhbGxlZCBkdXJp
bmcgdXNlciBFU1ANClZlcnNpb24gNC40OiBBZGRlZCBtb3JlIE9ESiBpbmZvDQpWZXJzaW9uIDQuMzogQWRkZWQgcG9saWN5IHRyYWNraW5nDQpWZXJzaW9u
IDQuMjogQnVnIGZpeGVzIGZvciBXaW5kb3dzIDEwIDIwMDQgKGV2ZW50IElEIGNoYW5nZXMpDQpWZXJzaW9uIDQuMTogUmVuYW1lZCB0byBHZXQtQXV0b3Bp
bG90RGlhZ25vc3RpY3MNClZlcnNpb24gNC4wOiBBZGRlZCBzaWRlY2FyIGluc3RhbGxhdGlvbiBpbmZvDQpWZXJzaW9uIDMuOTogQnVnIGZpeGVzDQpWZXJz
aW9uIDMuODogQnVnIGZpeGVzDQpWZXJzaW9uIDMuNzogTW9kaWZpZWQgT2ZmaWNlIGxvZ2ljIHRvIGVuc3VyZSBpdCBhY2N1cmF0ZWx5IHJlZmxlY3RlZCB3
aGF0IEVTUCB0aGlua3MgdGhlIHN0YXR1cyBpcy4gQWRkZWQgU2hvd1BvbGljaWVzIG9wdGlvbi4NClZlcnNpb24gMy4yOiBGaXhlZCBzaWRlY2FyIGRldGVj
dGlvbiBsb2dpYw0KVmVyc2lvbiAzLjE6IEZpeGVkIE9ESiBhcHBsaWVkIG91dHB1dA0KVmVyc2lvbiAzLjA6IEFkZGVkIHRoZSBhYmlsaXR5IHRvIHByb2Nl
c3MgbG9ncyBhcyB3ZWxsDQpWZXJzaW9uIDIuMjogQWRkZWQgbmV3IElNRSBNU0kgZ3VpZCwgbmV3IC1BbGxTZXNzaW9ucyBzd2l0Y2gNClZlcnNpb24gMi4w
OiBBZGRlZCAtb25saW5lIHBhcmFtZXRlciB0byBsb29rIHVwIGFwcCBhbmQgcG9saWN5IGRldGFpbHMNClZlcnNpb24gMS4wOiBPcmlnaW5hbCBwdWJsaXNo
ZWQgdmVyc2lvbg0KLlBSSVZBVEVEQVRBDQojPg0KDQo8IyANCg0KLkRFU0NSSVBUSU9ODQpUaGlzIHNjcmlwdCBkaXNwbGF5cyBkaWFnbm9zdGljcyBpbmZv
cm1hdGlvbiBmcm9tIHRoZSBjdXJyZW50IFBDIG9yIGEgY2FwdHVyZWQgc2V0IG9mIGxvZ3MuIFRoaXMgaW5jbHVkZXMgZGV0YWlscyBhYm91dCB0aGUgQXV0
b3BpbG90IHByb2ZpbGUgc2V0dGluZ3M7IHBvbGljaWVzLCBhcHBzLCBjZXJ0aWZpY2F0ZSBwcm9maWxlcywgZXRjLiBiZWluZyB0cmFja2VkIHZpYSB0aGUg
RW5yb2xsbWVudCBTdGF0dXMgUGFnZTsgYW5kIGFkZGl0aW9uYWwgaW5mb3JtYXRpb24uDQogDQojPiANCjwjDQouU1lOT1BTSVMNCkRpc3BsYXlzIFdpbmRv
d3MgQXV0b3BpbG90IGRpYWdub3N0aWNzIGluZm9ybWF0aW9uIGZyb20gdGhlIGN1cnJlbnQgUEMgb3IgYSBjYXB0dXJlZCBzZXQgb2YgbG9ncy4NCiANCi5Q
QVJBTUVURVIgT25saW5lDQpMb29rIHVwIHRoZSBhY3R1YWwgcG9saWN5IGFuZCBhcHAgbmFtZXMgdmlhIHRoZSBNaWNyb3NvZnQgR3JhcGggQVBJDQogDQou
UEFSQU1FVEVSIEFsbFNlc3Npb25zDQpTaG93IGFsbCBFU1AgcHJvZ3Jlc3MgaW5zdGVhZCBvZiBqdXN0IHRoZSBmaW5hbCBkZXRhaWxzLg0KIA0KLlBBUkFN
RVRFUiBGaWxlDQpQcm9jZXNzZXMgdGhlIGluZm9ybWF0aW9uIGluIHRoZSBzcGVjaWZpZWQgZmlsZSAoY2FwdHVyZWQgZWl0aGVyIGJ5IE1ETURpYWdub3N0
aWNzVG9vbC5leGUgLWFyZWEgQXV0b3BpbG90IC1jYWIgZmlsZW5hbWUuY2FiIG9yIE1ETURpYWdub3N0aWNzVG9vbC5leGUgLWFyZWEgQXV0b3BpbG90IC16
aXAgZmlsZW5hbWUuemlwKSBpbnN0ZWFkIG9mIGZyb20gdGhlIHJlZ2lzdHJ5Lg0KIA0KLlBBUkFNRVRFUiBTaG93UG9saWNpZXMNClNob3dzIHRoZSBwb2xp
Y3kgZGV0YWlscyBhcyByZWNvcmRlZCBpbiB0aGUgTm9kZUNhY2hlIHJlZ2lzdHJ5IGtleXMsIGluIHRoZSBvcmRlciB0aGF0IHRoZSBwb2xpY2llcyB3ZXJl
IHJlY2VpdmVkIGJ5IHRoZSBjbGllbnQuDQoNCi5QQVJBTUVURVIgVGVuYW50DQpUaGUgR1VJRCAodGV4dCBzdHJpbmcpLCBuZWVkZWQgd2hlbiBzcGVjaWZ5
aW5nIGFuIGFwcCBJRC9zZWNyZXQuDQogDQouUEFSQU1FVEVSIEFwcElkDQpUaGUgYXBwIElEIChHVUlEKSBmb3IgdGhlIEVudHJhIElEIGFwcCBiZWluZyB1
c2VkIHRvIGF1dGhlbnRpY2F0ZSB3aXRoIEludHVuZQ0KDQouUEFSQU1FVEVSIEFwcFNlY3JldA0KVGhlIGFwcCBzZWNyZXQgKGVmZmVjdGl2ZWx5IGEgcGFz
c3dvcmQpIGZvciB0aGUgc3BlY2lmaWVkIGFwcCBJRC4NCiANCi5QQVJBTUVURVIgQmVhcmVyDQpBbiBleGlzdGluZyBiZWFyZXIgdG9rZW4gdGhhdCB3aWxs
IGJlIHVzZWQgdG8gYXV0aGVudGNhdGUgdG8gSW50dW5lLg0KDQouRVhBTVBMRQ0KLlxHZXQtQXV0b3BpbG90RGlhZ25vc3RpY3MucHMxDQogDQouRVhBTVBM
RQ0KLlxHZXQtQXV0b3BpbG90RGlhZ25vc3RpY3MucHMxIC1PbmxpbmUNCiANCi5FWEFNUExFDQouXEdldC1BdXRvcGlsb3REaWFnbm9zdGljcy5wczEgLUFs
bFNlc3Npb25zDQogDQouRVhBTVBMRQ0KLlxHZXQtQXV0b3BpbG90RGlhZ25vc3RpY3MucHMxIC1GaWxlIEM6XEF1dG9waWxvdC5jYWIgLU9ubGluZSAtQWxs
U2Vzc2lvbnMNCiANCi5FWEFNUExFDQouXEdldC1BdXRvcGlsb3REaWFnbm9zdGljcy5wczEgLUZpbGUgQzpcQXV0b3BpbG90LnppcA0KIA0KLkVYQU1QTEUN
Ci5cR2V0LUF1dG9waWxvdERpYWdub3N0aWNzLnBzMSAtU2hvd1BvbGljaWVzDQogDQojPg0KDQpbQ21kbGV0QmluZGluZygpXQ0KcGFyYW0oDQogICAgW0Fs
aWFzKCJDQUJGaWxlIiwiWklQRmlsZSIsIkZ1bGxOYW1lIildW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkRmFsc2UsIFZhbHVlRnJvbVBpcGVsaW5lQnlQcm9w
ZXJ0eU5hbWUgPSAkdHJ1ZSldIFtTdHJpbmddICRGaWxlID0gJG51bGwsDQogICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkRmFsc2UpXSBbU3dpdGNoXSAk
T25saW5lID0gJGZhbHNlLA0KICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJEZhbHNlKV0gW1N3aXRjaF0gJEFsbFNlc3Npb25zID0gJGZhbHNlLA0KICAg
IFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJEZhbHNlKV0gW1N3aXRjaF0gJFNob3dQb2xpY2llcyA9ICRmYWxzZSwNCiAgICBbUGFyYW1ldGVyKE1hbmRhdG9y
eSA9ICRmYWxzZSldIFtzdHJpbmddICRUZW5hbnQsDQogICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkZmFsc2UpXSBbc3RyaW5nXSAkQXBwSWQsDQogICAg
W1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkZmFsc2UpXSBbc3RyaW5nXSAkQXBwU2VjcmV0LA0KICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJGZhbHNlKV0g
W3N0cmluZ10gJEJlYXJlcg0KKQ0KDQpCZWdpbiB7DQoNCiAgICAjIENvbmZpZ3VyZSBjb25zdGFudHMgYW5kIGdsb2JhbCB2YXJpYWJsZXMNCiAgICAkc2Ny
aXB0Om9mZmljZVN0YXR1cyA9IEB7IjAiID0gIk5vbmUiOyAiMTAiID0gIkluaXRpYWxpemVkIjsgIjIwIiA9ICJEb3dubG9hZCBJbiBQcm9ncmVzcyI7ICIy
NSIgPSAiUGVuZGluZyBEb3dubG9hZCBSZXRyeSI7DQogICAgICAgICIzMCIgPSAiRG93bmxvYWQgRmFpbGVkIjsgIjQwIiA9ICJEb3dubG9hZCBDb21wbGV0
ZWQiOyAiNDgiID0gIlBlbmRpbmcgVXNlciBTZXNzaW9uIjsgIjUwIiA9ICJFbmZvcmNlbWVudCBJbiBQcm9ncmVzcyI7IA0KICAgICAgICAiNTUiID0gIlBl
bmRpbmcgRW5mb3JjZW1lbnQgUmV0cnkiOyAiNjAiID0gIkVuZm9yY2VtZW50IEZhaWxlZCI7ICI3MCIgPSAiU3VjY2VzcyAvIEVuZm9yY2VtZW50IENvbXBs
ZXRlZCINCiAgICB9DQogICAgJHNjcmlwdDplc3BTdGF0dXMgPSBAeyIxIiA9ICJOb3QgSW5zdGFsbGVkIjsgIjIiID0gIkRvd25sb2FkaW5nIC8gSW5zdGFs
bGluZyI7ICIzIiA9ICJTdWNjZXNzIC8gSW5zdGFsbGVkIjsgIjQiID0gIkVycm9yIC8gRmFpbGVkIiB9DQogICAgJHNjcmlwdDpwb2xpY3lTdGF0dXMgPSBA
eyIwIiA9ICJOb3QgUHJvY2Vzc2VkIjsgIjEiID0gIlByb2Nlc3NlZCIgfQ0KDQogICAgZW51bSBBdXRvcGlsb3RTY2VuYXJpb0VudW0gew0KICAgICAgICBV
bmtub3duDQogICAgICAgIEF1dG9waWxvdFYxDQogICAgICAgIEF1dG9waWxvdEpzb24NCiAgICAgICAgRXNwT25seQ0KICAgICAgICBBdXRvcGlsb3RWMg0K
ICAgIH0NCg0KICAgIGVudW0gV29ya2xvYWRTdGF0ZQ0KICAgIHsNCiAgICAgICAgTm90U3RhcnRlZA0KICAgICAgICBDb21wbGV0ZWQNCiAgICAgICAgU2tp
cHBlZA0KICAgICAgICBVbmluc3RhbGxlZA0KICAgICAgICBGYWlsZWQNCiAgICAgICAgSW5Qcm9ncmVzcw0KICAgICAgICBSZWJvb3RSZXF1aXJlZA0KICAg
ICAgICBDYW5jZWxsZWQNCiAgICB9DQp9DQoNClByb2Nlc3Mgew0KICAgICMtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0NCiAgICAjIEZ1bmN0aW9ucw0KICAg
ICMtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0NCg0KICAgIGZ1bmN0aW9uIGdldGFsbHBhZ2luYXRpb24gKCkgew0KICAgICAgICA8Iw0KLlNZTk9QU0lTDQpU
aGlzIGZ1bmN0aW9uIGlzIHVzZWQgdG8gZ3JhYiBhbGwgaXRlbXMgZnJvbSBHcmFwaCBBUEkgdGhhdCBhcmUgcGFnaW5hdGVkDQouREVTQ1JJUFRJT04NClRo
ZSBmdW5jdGlvbiBjb25uZWN0cyB0byB0aGUgR3JhcGggQVBJIEludGVyZmFjZSBhbmQgZ2V0cyBhbGwgaXRlbXMgZnJvbSB0aGUgQVBJIHRoYXQgYXJlIHBh
Z2luYXRlZA0KLkVYQU1QTEUNCmdldGFsbHBhZ2luYXRpb24gLXVybCAiaHR0cHM6Ly9ncmFwaC5taWNyb3NvZnQuY29tL3YxLjAvZ3JvdXBzIg0KIFJldHVy
bnMgYWxsIGl0ZW1zDQouTk9URVMNCiBOQU1FOiBnZXRhbGxwYWdpbmF0aW9uDQojPg0KICAgICAgICBbY21kbGV0YmluZGluZygpXQ0KICAgIA0KICAgICAg
ICBwYXJhbQ0KICAgICAgICAoDQogICAgICAgICAgICAkdXJsDQogICAgICAgICkNCiAgICAgICAgJHJlc3BvbnNlID0gKEludm9rZS1NZ0dyYXBoUmVxdWVz
dCAtVXJpICR1cmwgLU1ldGhvZCBHZXQgLU91dHB1dFR5cGUgUFNPYmplY3QpDQogICAgICAgICRhbGxvdXRwdXQgPSAkcmVzcG9uc2UudmFsdWUNCiAgICAN
CiAgICAgICAgJGFsbG91dHB1dE5leHRMaW5rID0gJHJlc3BvbnNlLiJAb2RhdGEubmV4dExpbmsiDQogICAgDQogICAgICAgIHdoaWxlICgkbnVsbCAtbmUg
JGFsbG91dHB1dE5leHRMaW5rKSB7DQogICAgICAgICAgICAkYWxsb3V0cHV0UmVzcG9uc2UgPSAoSW52b2tlLU1nR3JhcGhSZXF1ZXN0IC1VcmkgJGFsbG91
dHB1dE5leHRMaW5rIC1NZXRob2QgR2V0IC1PdXRwdXRUeXBlIFBTT2JqZWN0KQ0KICAgICAgICAgICAgJGFsbG91dHB1dE5leHRMaW5rID0gJGFsbG91dHB1
dFJlc3BvbnNlLiJAb2RhdGEubmV4dExpbmsiDQogICAgICAgICAgICAkYWxsb3V0cHV0ICs9ICRhbGxvdXRwdXRSZXNwb25zZS52YWx1ZQ0KICAgICAgICB9
DQogICAgDQogICAgICAgIHJldHVybiAkYWxsb3V0cHV0DQogICAgfQ0KICAgIA0KICAgIEZ1bmN0aW9uIENvbm5lY3QtVG9HcmFwaCB7DQogICAgICAgIDwj
DQouU1lOT1BTSVMNCkF1dGhlbnRpY2F0ZXMgdG8gdGhlIEdyYXBoIEFQSSB2aWEgdGhlIE1pY3Jvc29mdC5HcmFwaC5BdXRoZW50aWNhdGlvbiBtb2R1bGUu
DQogDQouREVTQ1JJUFRJT04NClRoZSBDb25uZWN0LVRvR3JhcGggY21kbGV0IGlzIGEgd3JhcHBlciBjbWRsZXQgdGhhdCBoZWxwcyBhdXRoZW50aWNhdGUg
dG8gdGhlIEludHVuZSBHcmFwaCBBUEkgdXNpbmcgdGhlIE1pY3Jvc29mdC5HcmFwaC5BdXRoZW50aWNhdGlvbiBtb2R1bGUuIEl0IGxldmVyYWdlcyBhbiBB
enVyZSBBRCBhcHAgSUQgYW5kIGFwcCBzZWNyZXQgZm9yIGF1dGhlbnRpY2F0aW9uIG9yIHVzZXItYmFzZWQgYXV0aC4NCiANCi5QQVJBTUVURVIgVGVuYW50
DQpTcGVjaWZpZXMgdGhlIHRlbmFudCAoZS5nLiBjb250b3NvLm9ubWljcm9zb2Z0LmNvbSkgdG8gd2hpY2ggdG8gYXV0aGVudGljYXRlLg0KIA0KLlBBUkFN
RVRFUiBBcHBJZA0KU3BlY2lmaWVzIHRoZSBBenVyZSBBRCBhcHAgSUQgKEdVSUQpIGZvciB0aGUgYXBwbGljYXRpb24gdGhhdCB3aWxsIGJlIHVzZWQgdG8g
YXV0aGVudGljYXRlLg0KIA0KLlBBUkFNRVRFUiBBcHBTZWNyZXQNClNwZWNpZmllcyB0aGUgQXp1cmUgQUQgYXBwIHNlY3JldCBjb3JyZXNwb25kaW5nIHRv
IHRoZSBhcHAgSUQgdGhhdCB3aWxsIGJlIHVzZWQgdG8gYXV0aGVudGljYXRlLg0KDQouUEFSQU1FVEVSIFNjb3Blcw0KU3BlY2lmaWVzIHRoZSB1c2VyIHNj
b3BlcyBmb3IgaW50ZXJhY3RpdmUgYXV0aGVudGljYXRpb24uDQogDQouRVhBTVBMRQ0KQ29ubmVjdC1Ub0dyYXBoIC1UZW5hbnRJZCAkdGVuYW50SUQgLUFw
cElkICRhcHAgLUFwcFNlY3JldCAkc2VjcmV0DQogDQotIz4NCiAgICAgICAgW2NtZGxldGJpbmRpbmcoKV0NCiAgICAgICAgcGFyYW0NCiAgICAgICAgKA0K
ICAgICAgICAgICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkZmFsc2UpXSBbc3RyaW5nXSAkVGVuYW50LA0KICAgICAgICAgICAgW1BhcmFtZXRlcihNYW5k
YXRvcnkgPSAkZmFsc2UpXSBbc3RyaW5nXSAkQXBwSWQsDQogICAgICAgICAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICRmYWxzZSldIFtzdHJpbmddICRB
cHBTZWNyZXQsDQogICAgICAgICAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICRmYWxzZSldIFtzdHJpbmddICRzY29wZXMsDQogICAgICAgICAgICBbUGFy
YW1ldGVyKE1hbmRhdG9yeSA9ICRmYWxzZSldIFtzdHJpbmddICRCZWFyZXINCiAgICAgICAgKQ0KDQogICAgICAgIFByb2Nlc3Mgew0KICAgICAgICAgICAg
SW1wb3J0LU1vZHVsZSBNaWNyb3NvZnQuR3JhcGguQXV0aGVudGljYXRpb24NCiAgICAgICAgICAgICR2ZXJzaW9uID0gKGdldC1tb2R1bGUgbWljcm9zb2Z0
LmdyYXBoLmF1dGhlbnRpY2F0aW9uIHwgU2VsZWN0LU9iamVjdCAtZXhwYW5kcHJvcGVydHkgVmVyc2lvbikubWFqb3INCg0KICAgICAgICAgICAgaWYgKCRB
cHBJZCAtbmUgIiIpIHsNCiAgICAgICAgICAgICAgICAkYm9keSA9IEB7DQogICAgICAgICAgICAgICAgICAgIGdyYW50X3R5cGUgICAgPSAiY2xpZW50X2Ny
ZWRlbnRpYWxzIjsNCiAgICAgICAgICAgICAgICAgICAgY2xpZW50X2lkICAgICA9ICRBcHBJZDsNCiAgICAgICAgICAgICAgICAgICAgY2xpZW50X3NlY3Jl
dCA9ICRBcHBTZWNyZXQ7DQogICAgICAgICAgICAgICAgICAgIHNjb3BlICAgICAgICAgPSAiaHR0cHM6Ly9ncmFwaC5taWNyb3NvZnQuY29tLy5kZWZhdWx0
IjsNCiAgICAgICAgICAgICAgICB9DQogICAgIA0KICAgICAgICAgICAgICAgICRyZXNwb25zZSA9IEludm9rZS1SZXN0TWV0aG9kIC1NZXRob2QgUG9zdCAt
VXJpIGh0dHBzOi8vbG9naW4ubWljcm9zb2Z0b25saW5lLmNvbS8kVGVuYW50L29hdXRoMi92Mi4wL3Rva2VuIC1Cb2R5ICRib2R5DQogICAgICAgICAgICAg
ICAgJGFjY2Vzc1Rva2VuID0gJHJlc3BvbnNlLmFjY2Vzc190b2tlbg0KICAgICANCiAgICAgICAgICAgICAgICAkYWNjZXNzVG9rZW4NCiAgICAgICAgICAg
ICAgICBpZiAoJHZlcnNpb24gLWVxIDIpIHsNCiAgICAgICAgICAgICAgICAgICAgV3JpdGUtVmVyYm9zZSAiVmVyc2lvbiAyIG1vZHVsZSBkZXRlY3RlZCIN
CiAgICAgICAgICAgICAgICAgICAgJGFjY2Vzc3Rva2VuZmluYWwgPSBDb252ZXJ0VG8tU2VjdXJlU3RyaW5nIC1TdHJpbmcgJGFjY2Vzc1Rva2VuIC1Bc1Bs
YWluVGV4dCAtRm9yY2UNCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgZWxzZSB7DQogICAgICAgICAgICAgICAgICAgIFdyaXRlLVZlcmJv
c2UgIlZlcnNpb24gMSBNb2R1bGUgRGV0ZWN0ZWQiDQogICAgICAgICAgICAgICAgICAgIFNlbGVjdC1NZ1Byb2ZpbGUgLU5hbWUgQmV0YQ0KICAgICAgICAg
ICAgICAgICAgICAkYWNjZXNzdG9rZW5maW5hbCA9ICRhY2Nlc3NUb2tlbg0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBDb25uZWN0LU1n
R3JhcGggLUFjY2Vzc1Rva2VuICRhY2Nlc3N0b2tlbmZpbmFsIC1Ob1dlbGNvbWUNCiAgICAgICAgICAgICAgICBXcml0ZS1WZXJib3NlICJDb25uZWN0ZWQg
dG8gSW50dW5lIHRlbmFudCAkVGVuYW50IHVzaW5nIGFwcC1iYXNlZCBhdXRoZW50aWNhdGlvbiAoQXp1cmUgQUQgYXV0aGVudGljYXRpb24gbm90IHN1cHBv
cnRlZCkiDQogICAgICAgICAgICB9DQogICAgICAgICAgICBlbHNlaWYgKCRCZWFyZXIgLW5lICIiKSB7DQogICAgICAgICAgICAgICAgaWYgKCR2ZXJzaW9u
IC1lcSAyKSB7DQogICAgICAgICAgICAgICAgICAgIFdyaXRlLVZlcmJvc2UgIlZlcnNpb24gMiBtb2R1bGUgZGV0ZWN0ZWQiDQogICAgICAgICAgICAgICAg
ICAgICRhY2Nlc3N0b2tlbmZpbmFsID0gQ29udmVydFRvLVNlY3VyZVN0cmluZyAtU3RyaW5nICRCZWFyZXIgLUFzUGxhaW5UZXh0IC1Gb3JjZQ0KICAgICAg
ICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBlbHNlIHsNCiAgICAgICAgICAgICAgICAgICAgV3JpdGUtVmVyYm9zZSAiVmVyc2lvbiAxIE1vZHVsZSBE
ZXRlY3RlZCINCiAgICAgICAgICAgICAgICAgICAgU2VsZWN0LU1nUHJvZmlsZSAtTmFtZSBCZXRhDQogICAgICAgICAgICAgICAgICAgICRhY2Nlc3N0b2tl
bmZpbmFsID0gJEJlYXJlcg0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBDb25uZWN0LU1nR3JhcGggLUFjY2Vzc1Rva2VuICRhY2Nlc3N0
b2tlbmZpbmFsIC1Ob1dlbGNvbWUNCiAgICAgICAgICAgICAgICBXcml0ZS1WZXJib3NlICJDb25uZWN0ZWQgdG8gSW50dW5lIHRlbmFudCAkVGVuYW50IHVz
aW5nIGFwcC1iYXNlZCBhdXRoZW50aWNhdGlvbiAoQXp1cmUgQUQgYXV0aGVudGljYXRpb24gbm90IHN1cHBvcnRlZCkiDQogICAgICAgICAgICB9DQogICAg
ICAgICAgICBlbHNlIHsNCiAgICAgICAgICAgICAgICBpZiAoJHZlcnNpb24gLWVxIDIpIHsNCiAgICAgICAgICAgICAgICAgICAgV3JpdGUtVmVyYm9zZSAi
VmVyc2lvbiAyIG1vZHVsZSBkZXRlY3RlZCINCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgZWxzZSB7DQogICAgICAgICAgICAgICAgICAg
IFdyaXRlLVZlcmJvc2UgIlZlcnNpb24gMSBNb2R1bGUgRGV0ZWN0ZWQiDQogICAgICAgICAgICAgICAgICAgIFNlbGVjdC1NZ1Byb2ZpbGUgLU5hbWUgQmV0
YQ0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBDb25uZWN0LU1nR3JhcGggLXNjb3BlcyAkc2NvcGVzIC1Ob1dlbGNvbWUNCiAgICAgICAg
ICAgIH0NCiAgICAgICAgICAgICMgUmV0dXJuIHRoZSBjb250ZXh0DQogICAgICAgICAgICAkZ3JhcGggPSBHZXQtTWdDb250ZXh0DQogICAgICAgICAgICBX
cml0ZS1Ib3N0ICJDb25uZWN0ZWQgdG8gSW50dW5lIHRlbmFudCAkKCRncmFwaC5UZW5hbnRJZCkiDQogICAgICAgICAgICAkZ3JhcGgNCiAgICAgICAgfQ0K
ICAgIH0gICAgDQoNCiAgICBGdW5jdGlvbiBSZWNvcmRTdGF0dXMoKSB7DQogICAgICAgIHBhcmFtDQogICAgICAgICgNCiAgICAgICAgICAgIFtQYXJhbWV0
ZXIoTWFuZGF0b3J5ID0gJHRydWUpXSBbU3RyaW5nXSAkZGV0YWlsLA0KICAgICAgICAgICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkdHJ1ZSldIFtTdHJp
bmddICRzdGF0dXMsDQogICAgICAgICAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICR0cnVlKV0gW1N0cmluZ10gJGNvbG9yLA0KICAgICAgICAgICAgW1Bh
cmFtZXRlcihNYW5kYXRvcnkgPSAkdHJ1ZSldIFtkYXRldGltZV0gJGRhdGUNCiAgICAgICAgKQ0KDQogICAgICAgICMgU2VlIGlmIHRoZXJlIGlzIGFscmVh
ZHkgYW4gZW50cnkgZm9yIHRoaXMgcG9saWN5IGFuZCBzdGF0dXMNCiAgICAgICAgJGZvdW5kID0gJHNjcmlwdDpvYnNlcnZlZFRpbWVsaW5lIHwgPyB7ICRf
LkRldGFpbCAtZXEgJGRldGFpbCAtYW5kICRfLlN0YXR1cyAtZXEgJHN0YXR1cyB9DQogICAgICAgIGlmICgtbm90ICRmb3VuZCkgew0KICAgICAgICAgICAg
IyBBcHBseSBhIGZ1ZGdlIHNvIHRoYXQgdGhlIGRvd25sb2FkaW5nIG9mIHRoZSBuZXh0IGFwcCBhcHBlYXJzIG9uZSBzZWNvbmQgYWZ0ZXIgdGhlIHByZXZp
b3VzIGNvbXBsZXRpb24NCiAgICAgICAgICAgIGlmICgkc3RhdHVzIC1saWtlICJEb3dubG9hZGluZyoiKSB7DQogICAgICAgICAgICAgICAgJGFkanVzdGVk
RGF0ZSA9ICRkYXRlLkFkZFNlY29uZHMoMSkNCiAgICAgICAgICAgIH0NCiAgICAgICAgICAgIGVsc2Ugew0KICAgICAgICAgICAgICAgICRhZGp1c3RlZERh
dGUgPSAkZGF0ZQ0KICAgICAgICAgICAgfQ0KICAgICAgICAgICAgJHNjcmlwdDpvYnNlcnZlZFRpbWVsaW5lICs9IE5ldy1PYmplY3QgUFNPYmplY3QgLVBy
b3BlcnR5IEB7DQogICAgICAgICAgICAgICAgIkRhdGUiICAgPSAkYWRqdXN0ZWREYXRlDQogICAgICAgICAgICAgICAgIkRldGFpbCIgPSAkZGV0YWlsDQog
ICAgICAgICAgICAgICAgIlN0YXR1cyIgPSAkc3RhdHVzDQogICAgICAgICAgICAgICAgIkNvbG9yIiAgPSAkY29sb3INCiAgICAgICAgICAgIH0NCiAgICAg
ICAgfQ0KICAgIH0NCg0KICAgIEZ1bmN0aW9uIEFkZERpc3BsYXkoKSB7DQogICAgICAgIHBhcmFtDQogICAgICAgICgNCiAgICAgICAgICAgIFtQYXJhbWV0
ZXIoTWFuZGF0b3J5ID0gJHRydWUpXSBbcmVmXSRpdGVtcw0KICAgICAgICApDQogICAgICAgICRpdGVtcy5WYWx1ZSB8ICUgew0KICAgICAgICAgICAgQWRk
LU1lbWJlciAtSW5wdXRPYmplY3QgJF8gLU5vdGVQcm9wZXJ0eU5hbWUgZGlzcGxheSAtTm90ZVByb3BlcnR5VmFsdWUgJEFsbFNlc3Npb25zDQogICAgICAg
IH0NCiAgICAgICAgJGl0ZW1zLlZhbHVlWyRpdGVtcy5WYWx1ZS5Db3VudCAtIDFdLmRpc3BsYXkgPSAkdHJ1ZQ0KICAgIH0NCiAgICANCiAgICBGdW5jdGlv
biBQcm9jZXNzQXBwcygpIHsNCiAgICAgICAgcGFyYW0NCiAgICAgICAgKA0KICAgICAgICAgICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkdHJ1ZSwgVmFs
dWVGcm9tUGlwZWxpbmUgPSAkVHJ1ZSldIFtNaWNyb3NvZnQuV2luMzIuUmVnaXN0cnlLZXldICRjdXJyZW50S2V5LA0KICAgICAgICAgICAgW1BhcmFtZXRl
cihNYW5kYXRvcnkgPSAkdHJ1ZSldICRjdXJyZW50VXNlciwNCiAgICAgICAgICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJHRydWUsIFZhbHVlRnJvbVBp
cGVsaW5lQnlQcm9wZXJ0eU5hbWUgPSAkVHJ1ZSldIFtib29sXSAkZGlzcGxheQ0KICAgICAgICApDQoNCiAgICAgICAgQmVnaW4gew0KICAgICAgICAgICAg
aWYgKCRkaXNwbGF5KSB7IFdyaXRlLUhvc3QgIkFwcHM6IiB9DQogICAgICAgIH0NCg0KICAgICAgICBQcm9jZXNzIHsNCiAgICAgICAgICAgIGlmICgkZGlz
cGxheSkgeyBXcml0ZS1Ib3N0ICIgJCgoW2RhdGV0aW1lXSRjdXJyZW50S2V5LlBTQ2hpbGROYW1lKS5Ub1N0cmluZygndScpKSIgfQ0KICAgICAgICAgICAg
JGN1cnJlbnRLZXkuUHJvcGVydHkgfCAlIHsNCiAgICAgICAgICAgICAgICBpZiAoJF8uU3RhcnRzV2l0aCgiLi9EZXZpY2UvVmVuZG9yL01TRlQvRW50ZXJw
cmlzZURlc2t0b3BBcHBNYW5hZ2VtZW50L01TSS8iKSkgew0KICAgICAgICAgICAgICAgICAgICAkbXNpS2V5ID0gW1VSSV06OlVuZXNjYXBlRGF0YVN0cmlu
ZygoJF8uU3BsaXQoIi8iKSlbNl0pDQogICAgICAgICAgICAgICAgICAgICRmdWxsUGF0aCA9ICIkbXNpUGF0aFwkY3VycmVudFVzZXJcTVNJXCRtc2lLZXki
DQogICAgICAgICAgICAgICAgICAgIGlmIChUZXN0LVBhdGggJGZ1bGxQYXRoKSB7DQogICAgICAgICAgICAgICAgICAgICAgICAkc3RhdHVzID0gKEdldC1J
dGVtUHJvcGVydHkgLVBhdGggJGZ1bGxQYXRoKS5TdGF0dXMNCiAgICAgICAgICAgICAgICAgICAgICAgICRtc2lGaWxlID0gKEdldC1JdGVtUHJvcGVydHkg
LVBhdGggJGZ1bGxQYXRoKS5DdXJyZW50RG93bmxvYWRVcmwNCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICBpZiAoJHN0YXR1
cyAtZXEgIiIgLW9yICRzdGF0dXMgLWVxICRudWxsKSB7DQogICAgICAgICAgICAgICAgICAgICAgICAkc3RhdHVzID0gMA0KICAgICAgICAgICAgICAgICAg
ICB9IA0KICAgICAgICAgICAgICAgICAgICBpZiAoJG1zaUZpbGUgLW1hdGNoICJJbnR1bmVXaW5kb3dzQWdlbnQubXNpIikgew0KICAgICAgICAgICAgICAg
ICAgICAgICAgJG1zaUtleSA9ICJJbnR1bmUgTWFuYWdlbWVudCBFeHRlbnNpb25zICgkKCRtc2lLZXkpKSINCiAgICAgICAgICAgICAgICAgICAgfQ0KICAg
ICAgICAgICAgICAgICAgICBlbHNlaWYgKCRPbmxpbmUpIHsNCiAgICAgICAgICAgICAgICAgICAgICAgICRmb3VuZCA9ICRhcHBzIHwgPyB7ICRfLlByb2R1
Y3RDb2RlIC1jb250YWlucyAkbXNpS2V5IH0NCiAgICAgICAgICAgICAgICAgICAgICAgICRtc2lLZXkgPSAiJCgkZm91bmQuRGlzcGxheU5hbWUpICgkKCRt
c2lLZXkpKSINCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICBlbHNlaWYgKCRjdXJyZW50VXNlciAtZXEgIlMtMC0wLTAwLTAw
MDAwMDAwMDAtMDAwMDAwMDAwMC0wMDAwMDAwMDAtMDAwIikgew0KICAgICAgICAgICAgICAgICAgICAgICAgIyBUcnkgdG8gcmVhZCB0aGUgbmFtZSBmcm9t
IHRoZSB1bmluc3RhbGwgcmVnaXN0cnkga2V5DQogICAgICAgICAgICAgICAgICAgICAgICBpZiAoVGVzdC1QYXRoICJIS0xNOlxTb2Z0d2FyZVxNaWNyb3Nv
ZnRcV2luZG93c1xDdXJyZW50VmVyc2lvblxVbmluc3RhbGxcJG1zaUtleSIpIHsNCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAkZGlzcGxheU5hbWUg
PSAoR2V0LUl0ZW1Qcm9wZXJ0eSAtUGF0aCAiSEtMTTpcU29mdHdhcmVcTWljcm9zb2Z0XFdpbmRvd3NcQ3VycmVudFZlcnNpb25cVW5pbnN0YWxsXCRtc2lL
ZXkiKS5EaXNwbGF5TmFtZQ0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICRtc2lLZXkgPSAiJGRpc3BsYXlOYW1lICgkKCRtc2lLZXkpKSINCiAgICAg
ICAgICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICBpZiAoJHN0YXR1cyAtZXEgNzApIHsNCiAg
ICAgICAgICAgICAgICAgICAgICAgIGlmICgkZGlzcGxheSkgeyBXcml0ZS1Ib3N0ICIgTVNJICRtc2lLZXkgOiAkc3RhdHVzICgkKCRvZmZpY2VTdGF0dXNb
JHN0YXR1cy5Ub1N0cmluZygpXSkpIiAtRm9yZWdyb3VuZENvbG9yIEdyZWVuIH0NCiAgICAgICAgICAgICAgICAgICAgICAgIFJlY29yZFN0YXR1cyAtZGV0
YWlsICJNU0kgJG1zaUtleSIgLXN0YXR1cyAkb2ZmaWNlU3RhdHVzWyRzdGF0dXMuVG9TdHJpbmcoKV0gLWNvbG9yICJHcmVlbiIgLWRhdGUgJGN1cnJlbnRL
ZXkuUFNDaGlsZE5hbWUNCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICBlbHNlaWYgKCRzdGF0dXMgLWVxIDYwKSB7DQogICAg
ICAgICAgICAgICAgICAgICAgICBpZiAoJGRpc3BsYXkpIHsgV3JpdGUtSG9zdCAiIE1TSSAkbXNpS2V5IDogJHN0YXR1cyAoJCgkb2ZmaWNlU3RhdHVzWyRz
dGF0dXMuVG9TdHJpbmcoKV0pKSIgLUZvcmVncm91bmRDb2xvciBSZWQgfQ0KICAgICAgICAgICAgICAgICAgICAgICAgUmVjb3JkU3RhdHVzIC1kZXRhaWwg
Ik1TSSAkbXNpS2V5IiAtc3RhdHVzICRvZmZpY2VTdGF0dXNbJHN0YXR1cy5Ub1N0cmluZygpXSAtY29sb3IgIlJlZCIgLWRhdGUgJGN1cnJlbnRLZXkuUFND
aGlsZE5hbWUNCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICBlbHNlIHsNCiAgICAgICAgICAgICAgICAgICAgICAgIGlmICgk
ZGlzcGxheSkgeyBXcml0ZS1Ib3N0ICIgTVNJICRtc2lLZXkgOiAkc3RhdHVzICgkKCRvZmZpY2VTdGF0dXNbJHN0YXR1cy5Ub1N0cmluZygpXSkpIiAtRm9y
ZWdyb3VuZENvbG9yIFllbGxvdyB9DQogICAgICAgICAgICAgICAgICAgICAgICBSZWNvcmRTdGF0dXMgLWRldGFpbCAiTVNJICRtc2lLZXkiIC1zdGF0dXMg
JG9mZmljZVN0YXR1c1skc3RhdHVzLlRvU3RyaW5nKCldIC1jb2xvciAiWWVsbG93IiAtZGF0ZSAkY3VycmVudEtleS5QU0NoaWxkTmFtZQ0KICAgICAgICAg
ICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgIGVsc2VpZiAoJF8uU3RhcnRzV2l0aCgiLi9WZW5kb3IvTVNGVC9PZmZp
Y2UvSW5zdGFsbGF0aW9uLyIpKSB7DQogICAgICAgICAgICAgICAgICAgICMgUmVwb3J0IHRoZSBtYWluIHN0YXR1cyBiYXNlZCBvbiB3aGF0IEVTUCBpcyB0
cmFja2luZw0KICAgICAgICAgICAgICAgICAgICAkc3RhdHVzID0gR2V0LUl0ZW1Qcm9wZXJ0eVZhbHVlIC1QYXRoICRjdXJyZW50S2V5LlBTUGF0aCAtTmFt
ZSAkXw0KDQogICAgICAgICAgICAgICAgICAgICMgVGhlbiB0cnkgdG8gZ2V0IHRoZSBkZXRhaWxlZCBPZmZpY2Ugc3RhdHVzDQogICAgICAgICAgICAgICAg
ICAgICRvZmZpY2VLZXkgPSBbVVJJXTo6VW5lc2NhcGVEYXRhU3RyaW5nKCgkXy5TcGxpdCgiLyIpKVs1XSkNCiAgICAgICAgICAgICAgICAgICAgJGZ1bGxQ
YXRoID0gIiRvZmZpY2VwYXRoXCRvZmZpY2VLZXkiDQogICAgICAgICAgICAgICAgICAgIGlmIChUZXN0LVBhdGggJGZ1bGxQYXRoKSB7DQogICAgICAgICAg
ICAgICAgICAgICAgICAkb1N0YXR1cyA9IChHZXQtSXRlbVByb3BlcnR5IC1QYXRoICRmdWxsUGF0aCkuRmluYWxTdGF0dXMNCg0KICAgICAgICAgICAgICAg
ICAgICAgICAgaWYgKCRvU3RhdHVzIC1lcSAkbnVsbCkgew0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICRvU3RhdHVzID0gKEdldC1JdGVtUHJvcGVy
dHkgLVBhdGggJGZ1bGxQYXRoKS5TdGF0dXMNCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBpZiAoJG9TdGF0dXMgLWVxICRudWxsKSB7DQogICAgICAg
ICAgICAgICAgICAgICAgICAgICAgICAgICRvU3RhdHVzID0gIk5vbmUiDQogICAgICAgICAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAg
ICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgICAgIGVsc2Ugew0KICAgICAgICAgICAgICAgICAgICAgICAgJG9T
dGF0dXMgPSAiTm9uZSINCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICBpZiAoJG9mZmljZVN0YXR1cy5LZXlzIC1jb250YWlu
cyAkb1N0YXR1cy5Ub1N0cmluZygpKSB7DQogICAgICAgICAgICAgICAgICAgICAgICAkb2ZmaWNlU3RhdHVzVGV4dCA9ICRvZmZpY2VTdGF0dXNbJG9TdGF0
dXMuVG9TdHJpbmcoKV0NCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICBlbHNlIHsNCiAgICAgICAgICAgICAgICAgICAgICAg
ICRvZmZpY2VTdGF0dXNUZXh0ID0gJG9TdGF0dXMNCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICBpZiAoJHN0YXR1cyAtZXEg
MSkgew0KICAgICAgICAgICAgICAgICAgICAgICAgaWYgKCRkaXNwbGF5KSB7IFdyaXRlLUhvc3QgIiBPZmZpY2UgJG9mZmljZUtleSA6ICRzdGF0dXMgKCQo
JHBvbGljeVN0YXR1c1skc3RhdHVzLlRvU3RyaW5nKCldKSAvICRvZmZpY2VTdGF0dXNUZXh0KSIgLUZvcmVncm91bmRDb2xvciBHcmVlbiB9DQogICAgICAg
ICAgICAgICAgICAgICAgICBSZWNvcmRTdGF0dXMgLWRldGFpbCAiT2ZmaWNlICRvZmZpY2VLZXkiIC1zdGF0dXMgIiQoJHBvbGljeVN0YXR1c1skc3RhdHVz
LlRvU3RyaW5nKCldKSAvICRvZmZpY2VTdGF0dXNUZXh0IiAtY29sb3IgIkdyZWVuIiAtZGF0ZSAkY3VycmVudEtleS5QU0NoaWxkTmFtZQ0KICAgICAgICAg
ICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgICAgIGVsc2Ugew0KICAgICAgICAgICAgICAgICAgICAgICAgaWYgKCRkaXNwbGF5KSB7IFdyaXRlLUhv
c3QgIiBPZmZpY2UgJG9mZmljZUtleSA6ICRzdGF0dXMgKCQoJHBvbGljeVN0YXR1c1skc3RhdHVzLlRvU3RyaW5nKCldKSAvICRvZmZpY2VTdGF0dXNUZXh0
KSIgLUZvcmVncm91bmRDb2xvciBZZWxsb3cgfQ0KICAgICAgICAgICAgICAgICAgICAgICAgUmVjb3JkU3RhdHVzIC1kZXRhaWwgIk9mZmljZSAkb2ZmaWNl
S2V5IiAtc3RhdHVzICIkKCRwb2xpY3lTdGF0dXNbJHN0YXR1cy5Ub1N0cmluZygpXSkgLyAkb2ZmaWNlU3RhdHVzVGV4dCIgLWNvbG9yICJZZWxsb3ciIC1k
YXRlICRjdXJyZW50S2V5LlBTQ2hpbGROYW1lDQogICAgICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgZWxz
ZSB7DQogICAgICAgICAgICAgICAgICAgIGlmICgkZGlzcGxheSkgeyBXcml0ZS1Ib3N0ICIgJF8gOiBVbmtub3duIGFwcCIgfQ0KICAgICAgICAgICAgICAg
IH0NCiAgICAgICAgICAgIH0NCiAgICAgICAgfQ0KDQogICAgfQ0KDQogICAgRnVuY3Rpb24gUHJvY2Vzc01vZGVybkFwcHMoKSB7DQogICAgICAgIHBhcmFt
DQogICAgICAgICgNCiAgICAgICAgICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJHRydWUsIFZhbHVlRnJvbVBpcGVsaW5lID0gJFRydWUpXSBbTWljcm9z
b2Z0LldpbjMyLlJlZ2lzdHJ5S2V5XSAkY3VycmVudEtleSwNCiAgICAgICAgICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJHRydWUpXSAkY3VycmVudFVz
ZXIsDQogICAgICAgICAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICR0cnVlLCBWYWx1ZUZyb21QaXBlbGluZUJ5UHJvcGVydHlOYW1lID0gJFRydWUpXSBb
Ym9vbF0gJGRpc3BsYXkNCiAgICAgICAgKQ0KDQogICAgICAgIEJlZ2luIHsNCiAgICAgICAgICAgIGlmICgkZGlzcGxheSkgeyBXcml0ZS1Ib3N0ICJNb2Rl
cm4gQXBwczoiIH0NCiAgICAgICAgfQ0KDQogICAgICAgIFByb2Nlc3Mgew0KICAgICAgICAgICAgaWYgKCRkaXNwbGF5KSB7IFdyaXRlLUhvc3QgIiAkKChb
ZGF0ZXRpbWVdJGN1cnJlbnRLZXkuUFNDaGlsZE5hbWUpLlRvU3RyaW5nKCd1JykpIiB9DQogICAgICAgICAgICAkY3VycmVudEtleS5Qcm9wZXJ0eSB8ICUg
ew0KICAgICAgICAgICAgICAgICRzdGF0dXMgPSAoR2V0LUl0ZW1Qcm9wZXJ0eVZhbHVlIC1wYXRoICRjdXJyZW50S2V5LlBTUGF0aCAtTmFtZSAkXykuVG9T
dHJpbmcoKQ0KICAgICAgICAgICAgICAgIGlmICgkXy5TdGFydHNXaXRoKCIuL1VzZXIvVmVuZG9yL01TRlQvRW50ZXJwcmlzZU1vZGVybkFwcE1hbmFnZW1l
bnQvQXBwTWFuYWdlbWVudC8iKSkgew0KICAgICAgICAgICAgICAgICAgICAkYXBwSUQgPSBbVVJJXTo6VW5lc2NhcGVEYXRhU3RyaW5nKCgkXy5TcGxpdCgi
LyIpKVs3XSkNCiAgICAgICAgICAgICAgICAgICAgJHR5cGUgPSAiVXNlciBVV1AiDQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgIGVsc2Vp
ZiAoJF8uU3RhcnRzV2l0aCgiLi9EZXZpY2UvVmVuZG9yL01TRlQvRW50ZXJwcmlzZU1vZGVybkFwcE1hbmFnZW1lbnQvQXBwTWFuYWdlbWVudC8iKSkgew0K
ICAgICAgICAgICAgICAgICAgICAkYXBwSUQgPSBbVVJJXTo6VW5lc2NhcGVEYXRhU3RyaW5nKCgkXy5TcGxpdCgiLyIpKVs3XSkNCiAgICAgICAgICAgICAg
ICAgICAgJHR5cGUgPSAiRGV2aWNlIFVXUCINCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgZWxzZSB7DQogICAgICAgICAgICAgICAgICAg
ICRhcHBJRCA9ICRfDQogICAgICAgICAgICAgICAgICAgICR0eXBlID0gIlVua25vd24gVVdQIg0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAg
ICBpZiAoJHN0YXR1cyAtZXEgIjEiKSB7DQogICAgICAgICAgICAgICAgICAgIGlmICgkZGlzcGxheSkgeyBXcml0ZS1Ib3N0ICIgJHR5cGUgJGFwcElEIDog
JHN0YXR1cyAoJCgkcG9saWN5U3RhdHVzWyRzdGF0dXNdKSkiIC1Gb3JlZ3JvdW5kQ29sb3IgR3JlZW4gfQ0KICAgICAgICAgICAgICAgICAgICBSZWNvcmRT
dGF0dXMgLWRldGFpbCAiVVdQICRhcHBJRCIgLXN0YXR1cyAkcG9saWN5U3RhdHVzWyRzdGF0dXNdIC1jb2xvciAiR3JlZW4iIC1kYXRlICRjdXJyZW50S2V5
LlBTQ2hpbGROYW1lDQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgIGVsc2Ugew0KICAgICAgICAgICAgICAgICAgICBpZiAoJGRpc3BsYXkp
IHsgV3JpdGUtSG9zdCAiICR0eXBlICRhcHBJRCA6ICRzdGF0dXMgKCQoJHBvbGljeVN0YXR1c1skc3RhdHVzXSkpIiAtRm9yZWdyb3VuZENvbG9yIFllbGxv
dyB9DQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgfQ0KICAgICAgICB9DQoNCiAgICB9DQoNCiAgICBGdW5jdGlvbiBQcm9jZXNzU2lkZWNhclYy
KCkgew0KICAgICAgICBwYXJhbQ0KICAgICAgICAoDQogICAgICAgICAgICBbUGFyYW1ldGVyKFZhbHVlRnJvbVBpcGVsaW5lQnlQcm9wZXJ0eU5hbWUgPSAk
VHJ1ZSldIFtib29sXSAkZGlzcGxheSA9ICR0cnVlDQogICAgICAgICkNCg0KICAgICAgICBCZWdpbiB7DQogICAgICAgICAgICBpZiAoJGRpc3BsYXkpIHsg
V3JpdGUtSG9zdCAiU2lkZWNhciBhcHBzOiIgfQ0KICAgICAgICAgICAgaWYgKCRudWxsIC1lcSAkc2NyaXB0OkRPRXZlbnRzIC1hbmQgKC1ub3QgJHNjcmlw
dDp1c2VGaWxlKSkgew0KICAgICAgICAgICAgICAgICRzY3JpcHQ6RE9FdmVudHMgPSBHZXQtRGVsaXZlcnlPcHRpbWl6YXRpb25Mb2cgfCBXaGVyZS1PYmpl
Y3QgeyAkXy5GdW5jdGlvbiAtbWF0Y2ggIihEb3dubG9hZFN0YXJ0KXwoRG93bmxvYWRDb21wbGV0ZWQpIiAtYW5kICRfLk1lc3NhZ2UgLWxpa2UgIiouaW50
dW5ld2luLmJpbiwqIiB9DQogICAgICAgICAgICB9DQogICAgICAgIH0NCg0KICAgICAgICBQcm9jZXNzIHsNCiAgICAgICAgICAgIGlmIChUZXN0LVBhdGgg
IiRzaWRlY2FyV2luMzJBcHBzXFByb3Zpc2lvbmluZ1Byb2dyZXNzIikgew0KICAgICAgICAgICAgICAgICRkZXRhaWxzID0gR2V0LUl0ZW1Qcm9wZXJ0eVZh
bHVlIC1QYXRoICIkc2lkZWNhcldpbjMyQXBwc1xQcm92aXNpb25pbmdQcm9ncmVzcyIgLU5hbWUgIlByb3Zpc2lvbmluZ1Byb2dyZXNzIg0KICAgICAgICAg
ICAgICAgIGlmICgkZGV0YWlscykgew0KICAgICAgICAgICAgICAgICAgICAkcHJvdmlzaW9uaW5nUHJvZ3Jlc3MgPSAkZGV0YWlscyB8IENvbnZlcnRGcm9t
LUpzb24NCiAgICAgICAgICAgICAgICAgICAgJHByb3Zpc2lvbmluZ1Byb2dyZXNzLldvcmtsb2FkcyB8IEZvckVhY2gtT2JqZWN0IHsNCiAgICAgICAgICAg
ICAgICAgICAgICAgICMgIldvcmtsb2FkSWQiOiI0MWU5MzFlZi05OTUxLTQ2NDYtYWEwMC02ZGY0NzRhNWQ2NmQiLCJGcmllbmRseU5hbWUiOiJQb3dlclRv
eXMgMC45MC4xIiwiV29ya2xvYWRTdGF0ZSI6MSwiU3RhcnRUaW1lIjoiXC9EYXRlKDE3NDU4NzExNzY2MzQpXC8iLCJFbmRUaW1lIjoiXC9EYXRlKDE3NDU4
NzEyNTY2NzIpXC8iLCJFcnJvckNvZGUiOm51bGwgfQ0KICAgICAgICAgICAgICAgICAgICAgICAgUmVjb3JkU3RhdHVzIC1kZXRhaWwgJF8uRnJpZW5kbHlO
YW1lIC1zdGF0dXMgIkluc3RhbGxhdGlvbiBzdGFydGVkIiAtY29sb3IgIlllbGxvdyIgLWRhdGUgJF8uU3RhcnRUaW1lDQogICAgICAgICAgICAgICAgICAg
ICAgICAkc3RhdHVzID0gW1dvcmtsb2FkU3RhdGVdJF8uV29ya2xvYWRTdGF0ZQ0KICAgICAgICAgICAgICAgICAgICAgICAgaWYgKCRzdGF0dXMgLWVxIFtX
b3JrbG9hZFN0YXRlXTo6Q29tcGxldGVkKSB7DQogICAgICAgICAgICAgICAgICAgICAgICAgICAgaWYgKCRkaXNwbGF5KSB7IFdyaXRlLUhvc3QgIiAkKCRf
LkZyaWVuZGx5TmFtZSkgOiAkc3RhdHVzIiAtRm9yZWdyb3VuZENvbG9yIEdyZWVuIH0NCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBSZWNvcmRTdGF0
dXMgLWRldGFpbCAkXy5GcmllbmRseU5hbWUgLXN0YXR1cyAkc3RhdHVzIC1jb2xvciAiR3JlZW4iIC1kYXRlICRfLkVuZFRpbWUNCiAgICAgICAgICAgICAg
ICAgICAgICAgIH0gZWxzZWlmICgkc3RhdHVzIC1lcSBbV29ya2xvYWRTdGF0ZV06OkZhaWxlZCkgew0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICRl
bmZvcmNlbWVudFN0YXR1cyA9IEdldC1JdGVtUHJvcGVydHlWYWx1ZSAtUGF0aCAiJHNpZGVjYXJXaW4zMkFwcHNcMDAwMDAwMDAtMDAwMC0wMDAwLTAwMDAt
MDAwMDAwMDAwMDAwXCQoJF8uV29ya2xvYWRJZCkqXEVuZm9yY2VtZW50U3RhdGVNZXNzYWdlIiAtTmFtZSBFbmZvcmNlbWVudFN0YXRlTWVzc2FnZSB8IENv
bnZlcnRGcm9tLUpzb24NCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBpZiAoJGRpc3BsYXkpIHsgV3JpdGUtSG9zdCAiICQoJF8uRnJpZW5kbHlOYW1l
KSA6ICRzdGF0dXMsIHJjID0gJCgkZW5mb3JjZW1lbnRTdGF0dXMuRXJyb3JDb2RlKSIgLUZvcmVncm91bmRDb2xvciBSZWQgfQ0KICAgICAgICAgICAgICAg
ICAgICAgICAgICAgIFJlY29yZFN0YXR1cyAtZGV0YWlsICRfLkZyaWVuZGx5TmFtZSAtc3RhdHVzICRzdGF0dXMgLWNvbG9yICJSZWQiIC1kYXRlICRfLkVu
ZFRpbWUNCiAgICAgICAgICAgICAgICAgICAgICAgIH0gZWxzZSB7DQogICAgICAgICAgICAgICAgICAgICAgICAgICAgaWYgKCRkaXNwbGF5KSB7IFdyaXRl
LUhvc3QgIiAkKCRfLkZyaWVuZGx5TmFtZSkgOiAkc3RhdHVzIiAtRm9yZWdyb3VuZENvbG9yIFllbGxvdyB9DQogICAgICAgICAgICAgICAgICAgICAgICAg
ICAgUmVjb3JkU3RhdHVzIC1kZXRhaWwgJF8uRnJpZW5kbHlOYW1lIC1zdGF0dXMgJHN0YXR1cyAtY29sb3IgIlllbGxvdyIgLWRhdGUgJF8uU3RhcnRUaW1l
DQogICAgICAgICAgICAgICAgICAgICAgICB9DQoNCiAgICAgICAgICAgICAgICAgICAgICAgICMgVHJ5IHRvIGZpbmQgdGhlIERPIGV2ZW50cy4NCiAgICAg
ICAgICAgICAgICAgICAgICAgIGlmICgkc2NyaXB0OkRPRXZlbnRzKSB7DQogICAgICAgICAgICAgICAgICAgICAgICAgICAgJGFwcE5hbWUgPSAkXy5Gcmll
bmRseU5hbWUNCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAkYXBwSWQgPSAkXy5Xb3JrbG9hZElkDQogICAgICAgICAgICAgICAgICAgICAgICAgICAg
JHNjcmlwdDpET0V2ZW50cyB8IFdoZXJlLU9iamVjdCB7ICRfLk1lc3NhZ2UgLWlsaWtlICIqJGFwcElkKiIgfSB8IEZvckVhY2gtT2JqZWN0IHsNCiAgICAg
ICAgICAgICAgICAgICAgICAgICAgICAgICAgaWYgKCRfLkZ1bmN0aW9uLkNvbnRhaW5zKCJEb3dubG9hZFN0YXJ0IikpIA0KICAgICAgICAgICAgICAgICAg
ICAgICAgICAgICAgICB7DQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAkb3AgPSAiRG93bmxvYWRTdGFydCINCiAgICAgICAgICAgICAg
ICAgICAgICAgICAgICAgICAgfSBlbHNlIHsNCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICRvcCA9ICJEb3dubG9hZENvbXBsZXRlZCIN
CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBSZWNvcmRTdGF0dXMgLWRldGFpbCAk
YXBwTmFtZSAtc3RhdHVzICJETyAkb3AiIC1jb2xvciAiWWVsbG93IiAtZGF0ZSAkXy5UaW1lQ3JlYXRlZA0KICAgICAgICAgICAgICAgICAgICAgICAgICAg
IH0gICAgDQogICAgICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICB9IGVsc2Ugew0KICAgICAg
ICAgICAgICAgICAgICBpZiAoJGRpc3BsYXkpIHsgV3JpdGUtSG9zdCAiIFByb3Zpc2lvbmluZyBwcm9ncmVzcyBkZXRhaWxzIG5vdCBmb3VuZC4iIH0NCiAg
ICAgICAgICAgICAgICB9ICAgICAgICAgICAgDQogICAgICAgICAgICB9IGVsc2Ugew0KICAgICAgICAgICAgICAgIGlmICgkZGlzcGxheSkgeyBXcml0ZS1I
b3N0ICIgUHJvdmlzaW9uaW5nIHByb2dyZXNzIG5vdCBmb3VuZC4iIH0NCiAgICAgICAgICAgIH0NCiAgICAgICAgfQ0KICAgIH0NCg0KICAgIEZ1bmN0aW9u
IFByb2Nlc3NTaWRlY2FyVjJTY3JpcHRzKCkgew0KICAgICAgICBwYXJhbQ0KICAgICAgICAoDQogICAgICAgICAgICBbUGFyYW1ldGVyKFZhbHVlRnJvbVBp
cGVsaW5lQnlQcm9wZXJ0eU5hbWUgPSAkVHJ1ZSldIFtib29sXSAkZGlzcGxheSA9ICR0cnVlDQogICAgICAgICkNCg0KICAgICAgICBCZWdpbiB7DQogICAg
ICAgICAgICBpZiAoJGRpc3BsYXkpIHsgV3JpdGUtSG9zdCAiU2lkZWNhciBzY3JpcHRzOiIgfQ0KICAgICAgICB9DQoNCiAgICAgICAgUHJvY2VzcyB7DQog
ICAgICAgICAgICBpZiAoVGVzdC1QYXRoICIkc2lkZWNhclBhdGhcUG9saWNpZXNcMDAwMDAwMDAtMDAwMC0wMDAwLTAwMDAtMDAwMDAwMDAwMDAwIikgew0K
ICAgICAgICAgICAgICAgIEdldC1DaGlsZEl0ZW0gLVBhdGggIiRzaWRlY2FyUGF0aFxQb2xpY2llc1wwMDAwMDAwMC0wMDAwLTAwMDAtMDAwMC0wMDAwMDAw
MDAwMDAiIHwgRm9yRWFjaC1PYmplY3Qgew0KICAgICAgICAgICAgICAgICAgICAkc2NyaXB0SWQgPSAkXy5QU0NoaWxkTmFtZQ0KICAgICAgICAgICAgICAg
ICAgICAkc2NyaXB0TmFtZSA9ICRzY3JpcHRJZA0KICAgICAgICAgICAgICAgICAgICAkcHJvcGVydGllcyA9IEdldC1JdGVtUHJvcGVydHkgLVBhdGggJF8u
UFNQYXRoDQogICAgICAgICAgICAgICAgICAgICRyZXN1bHQgPSAkcHJvcGVydGllcy5SZXN1bHQNCiAgICAgICAgICAgICAgICAgICAgJHdoZW4gPSBbRGF0
ZVRpbWVdJHByb3BlcnRpZXMuTGFzdFVwZGF0ZWRUaW1lVXRjDQogICAgICAgICAgICAgICAgICAgIGlmICgkT25saW5lKSB7DQogICAgICAgICAgICAgICAg
ICAgICAgICAkc2NyaXB0cyB8IFdoZXJlLU9iamVjdCB7ICRzY3JpcHRJZCAtZXEgJF8uSWQgfSB8IEZvckVhY2gtT2JqZWN0IHsNCiAgICAgICAgICAgICAg
ICAgICAgICAgICAgICAkc2NyaXB0TmFtZSA9ICIkKCRfLkRpc3BsYXlOYW1lKSAoc2NyaXB0KSIgICAgICAgICAgICAgICAgICAgICAgIA0KICAgICAgICAg
ICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgICAgIGlmICgkcmVzdWx0IC1lcSAiU3VjY2VzcyIpIHsN
CiAgICAgICAgICAgICAgICAgICAgICAgIGlmICgkZGlzcGxheSkgeyBXcml0ZS1Ib3N0ICIgJHNjcmlwdE5hbWUgOiAkcmVzdWx0IiAtRm9yZWdyb3VuZENv
bG9yIEdyZWVuIH0NCiAgICAgICAgICAgICAgICAgICAgICAgIFJlY29yZFN0YXR1cyAtZGV0YWlsICRzY3JpcHROYW1lIC1zdGF0dXMgJHJlc3VsdCAtY29s
b3IgIkdyZWVuIiAtZGF0ZSAkd2hlbg0KICAgICAgICAgICAgICAgICAgICB9IGVsc2VpZiAoJHJlc3VsdCAtZXEgIkZhaWxlZCIpIHsNCiAgICAgICAgICAg
ICAgICAgICAgICAgIGlmICgkZGlzcGxheSkgeyBXcml0ZS1Ib3N0ICIgJHNjcmlwdE5hbWUgOiAkcmVzdWx0IiAtRm9yZWdyb3VuZENvbG9yIFJlZCB9DQog
ICAgICAgICAgICAgICAgICAgICAgICBSZWNvcmRTdGF0dXMgLWRldGFpbCAkc2NyaXB0TmFtZSAtc3RhdHVzICRyZXN1bHQgLWNvbG9yICJSZWQiIC1kYXRl
ICR3aGVuDQogICAgICAgICAgICAgICAgICAgIH0gZWxzZSB7DQogICAgICAgICAgICAgICAgICAgICAgICBpZiAoJGRpc3BsYXkpIHsgV3JpdGUtSG9zdCAi
ICRzY3JpcHROYW1lIDogJHJlc3VsdCIgLUZvcmVncm91bmRDb2xvciBZZWxsb3cgfQ0KICAgICAgICAgICAgICAgICAgICAgICAgUmVjb3JkU3RhdHVzIC1k
ZXRhaWwgJHNjcmlwdE5hbWUgLXN0YXR1cyAkcmVzdWx0IC1jb2xvciAiUmVkIiAtZGF0ZSAkd2hlbg0KICAgICAgICAgICAgICAgICAgICB9DQogICAgICAg
ICAgICAgICAgfSAgICAgICAgICAgIA0KICAgICAgICAgICAgfSBlbHNlIHsNCiAgICAgICAgICAgICAgICBpZiAoJGRpc3BsYXkpIHsgV3JpdGUtSG9zdCAi
IFByb3Zpc2lvbmluZyBzY3JpcHQgaW5mbyBub3QgZm91bmQuIiB9DQogICAgICAgICAgICB9DQogICAgICAgIH0NCiAgICB9DQoNCiAgICBGdW5jdGlvbiBQ
cm9jZXNzU2lkZWNhcigpIHsNCiAgICAgICAgcGFyYW0NCiAgICAgICAgKA0KICAgICAgICAgICAgW1BhcmFtZXRlcihNYW5kYXRvcnkgPSAkdHJ1ZSwgVmFs
dWVGcm9tUGlwZWxpbmUgPSAkVHJ1ZSldIFtNaWNyb3NvZnQuV2luMzIuUmVnaXN0cnlLZXldICRjdXJyZW50S2V5LA0KICAgICAgICAgICAgW1BhcmFtZXRl
cihNYW5kYXRvcnkgPSAkdHJ1ZSldICRjdXJyZW50VXNlciwNCiAgICAgICAgICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJHRydWUsIFZhbHVlRnJvbVBp
cGVsaW5lQnlQcm9wZXJ0eU5hbWUgPSAkVHJ1ZSldIFtib29sXSAkZGlzcGxheQ0KICAgICAgICApDQoNCiAgICAgICAgQmVnaW4gew0KICAgICAgICAgICAg
aWYgKCRkaXNwbGF5KSB7IFdyaXRlLUhvc3QgIlNpZGVjYXIgYXBwczoiIH0NCiAgICAgICAgICAgIGlmICgkbnVsbCAtZXEgJHNjcmlwdDpET0V2ZW50cyAt
YW5kICgtbm90ICRzY3JpcHQ6dXNlRmlsZSkpIHsNCiAgICAgICAgICAgICAgICAkc2NyaXB0OkRPRXZlbnRzID0gR2V0LURlbGl2ZXJ5T3B0aW1pemF0aW9u
TG9nIHwgV2hlcmUtT2JqZWN0IHsgJF8uRnVuY3Rpb24gLW1hdGNoICIoRG93bmxvYWRTdGFydCl8KERvd25sb2FkQ29tcGxldGVkKSIgLWFuZCAkXy5NZXNz
YWdlIC1saWtlICIqLmludHVuZXdpbi5iaW4sKiIgfQ0KICAgICAgICAgICAgfQ0KICAgICAgICB9DQoNCiAgICAgICAgUHJvY2VzcyB7DQogICAgICAgICAg
ICBpZiAoJGRpc3BsYXkpIHsgV3JpdGUtSG9zdCAiICQoKFtkYXRldGltZV0kY3VycmVudEtleS5QU0NoaWxkTmFtZSkuVG9TdHJpbmcoJ3UnKSkiIH0NCiAg
ICAgICAgICAgICRjdXJyZW50S2V5LlByb3BlcnR5IHwgJSB7DQogICAgICAgICAgICAgICAgJHdpbjMyS2V5ID0gW1VSSV06OlVuZXNjYXBlRGF0YVN0cmlu
ZygoJF8uU3BsaXQoIi8iKSlbOV0pDQogICAgICAgICAgICAgICAgJHN0YXR1cyA9IEdldC1JdGVtUHJvcGVydHlWYWx1ZSAtcGF0aCAkY3VycmVudEtleS5Q
U1BhdGggLU5hbWUgJF8NCiAgICAgICAgICAgICAgICBpZiAoJE9ubGluZSkgew0KICAgICAgICAgICAgICAgICAgICAkZm91bmQgPSAkYXBwcyB8ID8geyAk
d2luMzJLZXkgLW1hdGNoICRfLklkIH0NCiAgICAgICAgICAgICAgICAgICAgJHdpbjMyS2V5ID0gIiQoJGZvdW5kLkRpc3BsYXlOYW1lKSAoJCgkd2luMzJL
ZXkpKSINCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgJGFwcEd1aWQgPSAkd2luMzJLZXkuU3Vic3RyaW5nKDkpDQogICAgICAgICAgICAg
ICAgJHNpZGVjYXJBcHAgPSAiJHNpZGVjYXJXaW4zMkFwcHNcJGN1cnJlbnRVc2VyXCRhcHBHdWlkIg0KICAgICAgICAgICAgICAgICRleGl0Q29kZSA9ICRu
dWxsDQogICAgICAgICAgICAgICAgaWYgKFRlc3QtUGF0aCAkc2lkZWNhckFwcCkgew0KICAgICAgICAgICAgICAgICAgICAkZXhpdENvZGUgPSAoR2V0LUl0
ZW1Qcm9wZXJ0eSAtUGF0aCAkc2lkZWNhckFwcCkuRXhpdENvZGUNCiAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgaWYgKCRzdGF0dXMgLWVx
ICIzIikgew0KICAgICAgICAgICAgICAgICAgICBpZiAoJGV4aXRDb2RlIC1uZSAkbnVsbCkgew0KICAgICAgICAgICAgICAgICAgICAgICAgaWYgKCRkaXNw
bGF5KSB7IFdyaXRlLUhvc3QgIiBXaW4zMiAkd2luMzJLZXkgOiAkc3RhdHVzICgkKCRlc3BTdGF0dXNbJHN0YXR1cy5Ub1N0cmluZygpXSksIHJjID0gJGV4
aXRDb2RlKSIgLUZvcmVncm91bmRDb2xvciBHcmVlbiB9DQogICAgICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICAgICAgZWxzZSB7DQogICAg
ICAgICAgICAgICAgICAgICAgICBpZiAoJGRpc3BsYXkpIHsgV3JpdGUtSG9zdCAiIFdpbjMyICR3aW4zMktleSA6ICRzdGF0dXMgKCQoJGVzcFN0YXR1c1sk
c3RhdHVzLlRvU3RyaW5nKCldKSkiIC1Gb3JlZ3JvdW5kQ29sb3IgR3JlZW4gfQ0KICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgICAg
IFJlY29yZFN0YXR1cyAtZGV0YWlsICJXaW4zMiAkd2luMzJLZXkiIC1zdGF0dXMgJGVzcFN0YXR1c1skc3RhdHVzLlRvU3RyaW5nKCldIC1jb2xvciAiR3Jl
ZW4iIC1kYXRlICRjdXJyZW50S2V5LlBTQ2hpbGROYW1lDQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgIGVsc2VpZiAoJHN0YXR1cyAtZXEg
IjQiKSB7DQogICAgICAgICAgICAgICAgICAgIGlmICgkZXhpdENvZGUgLW5lICRudWxsKSB7DQogICAgICAgICAgICAgICAgICAgICAgICBpZiAoJGRpc3Bs
YXkpIHsgV3JpdGUtSG9zdCAiIFdpbjMyICR3aW4zMktleSA6ICRzdGF0dXMgKCQoJGVzcFN0YXR1c1skc3RhdHVzLlRvU3RyaW5nKCldKSwgcmMgPSAkZXhp
dENvZGUiIC1Gb3JlZ3JvdW5kQ29sb3IgUmVkIH0NCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICBlbHNlIHsNCiAgICAgICAg
ICAgICAgICAgICAgICAgIGlmICgkZGlzcGxheSkgeyBXcml0ZS1Ib3N0ICIgV2luMzIgJHdpbjMyS2V5IDogJHN0YXR1cyAoJCgkZXNwU3RhdHVzWyRzdGF0
dXMuVG9TdHJpbmcoKV0pKSIgLUZvcmVncm91bmRDb2xvciBSZWQgfQ0KICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgICAgIFJlY29y
ZFN0YXR1cyAtZGV0YWlsICJXaW4zMiAkd2luMzJLZXkiIC1zdGF0dXMgJGVzcFN0YXR1c1skc3RhdHVzLlRvU3RyaW5nKCldIC1jb2xvciAiUmVkIiAtZGF0
ZSAkY3VycmVudEtleS5QU0NoaWxkTmFtZQ0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBlbHNlIHsNCiAgICAgICAgICAgICAgICAgICAg
aWYgKCRleGl0Q29kZSAtbmUgJG51bGwpIHsNCiAgICAgICAgICAgICAgICAgICAgICAgIGlmICgkZGlzcGxheSkgeyBXcml0ZS1Ib3N0ICIgV2luMzIgJHdp
bjMyS2V5IDogJHN0YXR1cyAoJCgkZXNwU3RhdHVzWyRzdGF0dXMuVG9TdHJpbmcoKV0pLCByYyA9ICRleGl0Q29kZSkiIC1Gb3JlZ3JvdW5kQ29sb3IgWWVs
bG93IH0NCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICBlbHNlIHsNCiAgICAgICAgICAgICAgICAgICAgICAgIGlmICgkZGlz
cGxheSkgeyBXcml0ZS1Ib3N0ICIgV2luMzIgJHdpbjMyS2V5IDogJHN0YXR1cyAoJCgkZXNwU3RhdHVzWyRzdGF0dXMuVG9TdHJpbmcoKV0pKSIgLUZvcmVn
cm91bmRDb2xvciBZZWxsb3cgfQ0KICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgICAgIGlmICgkc3RhdHVzIC1uZSAiMSIpIHsNCiAg
ICAgICAgICAgICAgICAgICAgICAgIFJlY29yZFN0YXR1cyAtZGV0YWlsICJXaW4zMiAkd2luMzJLZXkiIC1zdGF0dXMgJGVzcFN0YXR1c1skc3RhdHVzLlRv
U3RyaW5nKCldIC1jb2xvciAiWWVsbG93IiAtZGF0ZSAkY3VycmVudEtleS5QU0NoaWxkTmFtZQ0KICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAg
ICAgICAgICAgIGlmICgkc3RhdHVzIC1lcSAiMiIpIHsNCiAgICAgICAgICAgICAgICAgICAgICAgICMgVHJ5IHRvIGZpbmQgdGhlIERPIGV2ZW50cy4NCiAg
ICAgICAgICAgICAgICAgICAgICAgICRzY3JpcHQ6RE9FdmVudHMgfCBXaGVyZS1PYmplY3QgeyAkXy5NZXNzYWdlIC1pbGlrZSAiKiRhcHBHdWlkKiIgfSB8
IEZvckVhY2gtT2JqZWN0IHsNCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBpZiAoJF8uRnVuY3Rpb24uQ29udGFpbnMoIkRvd25sb2FkU3RhcnQiKSkg
DQogICAgICAgICAgICAgICAgICAgICAgICAgICAgew0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAkb3AgPSAiRG93bmxvYWRTdGFydCINCiAg
ICAgICAgICAgICAgICAgICAgICAgICAgICB9IGVsc2Ugew0KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAkb3AgPSAiRG93bmxvYWRDb21wbGV0
ZWQiDQogICAgICAgICAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICAgICAgICAgIFJlY29yZFN0YXR1cyAtZGV0YWlsICJXaW4z
MiAkd2luMzJLZXkiIC1zdGF0dXMgIkRPICRvcCIgLWNvbG9yICJZZWxsb3ciIC1kYXRlICRfLlRpbWVDcmVhdGVkLlRvTG9jYWxUaW1lKCkNCiAgICAgICAg
ICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgIH0NCiAgICAgICAgfQ0KDQog
ICAgfQ0KDQogICAgRnVuY3Rpb24gUHJvY2Vzc1BvbGljaWVzKCkgew0KICAgICAgICBwYXJhbQ0KICAgICAgICAoDQogICAgICAgICAgICBbUGFyYW1ldGVy
KE1hbmRhdG9yeSA9ICR0cnVlLCBWYWx1ZUZyb21QaXBlbGluZSA9ICRUcnVlKV0gW01pY3Jvc29mdC5XaW4zMi5SZWdpc3RyeUtleV0gJGN1cnJlbnRLZXks
DQogICAgICAgICAgICBbUGFyYW1ldGVyKE1hbmRhdG9yeSA9ICR0cnVlLCBWYWx1ZUZyb21QaXBlbGluZUJ5UHJvcGVydHlOYW1lID0gJFRydWUpXSBbYm9v
bF0gJGRpc3BsYXkNCiAgICAgICAgKQ0KDQogICAgICAgIEJlZ2luIHsNCiAgICAgICAgICAgIGlmICgkZGlzcGxheSkgeyBXcml0ZS1Ib3N0ICJQb2xpY2ll
czoiIH0NCiAgICAgICAgfQ0KDQogICAgICAgIFByb2Nlc3Mgew0KICAgICAgICAgICAgaWYgKCRkaXNwbGF5KSB7IFdyaXRlLUhvc3QgIiAkKChbZGF0ZXRp
bWVdJGN1cnJlbnRLZXkuUFNDaGlsZE5hbWUpLlRvU3RyaW5nKCd1JykpIiB9DQogICAgICAgICAgICAkY3VycmVudEtleS5Qcm9wZXJ0eSB8ICUgew0KICAg
ICAgICAgICAgICAgICRzdGF0dXMgPSBHZXQtSXRlbVByb3BlcnR5VmFsdWUgLXBhdGggJGN1cnJlbnRLZXkuUFNQYXRoIC1OYW1lICRfDQogICAgICAgICAg
ICAgICAgaWYgKCRzdGF0dXMgLWVxICIxIikgew0KICAgICAgICAgICAgICAgICAgICBpZiAoJGRpc3BsYXkpIHsgV3JpdGUtSG9zdCAiIFBvbGljeSAkXyA6
ICRzdGF0dXMgKCQoJHBvbGljeVN0YXR1c1skc3RhdHVzLlRvU3RyaW5nKCldKSkiIC1Gb3JlZ3JvdW5kQ29sb3IgR3JlZW4gfQ0KICAgICAgICAgICAgICAg
ICAgICBSZWNvcmRTdGF0dXMgLWRldGFpbCAiUG9saWN5ICRfIiAtc3RhdHVzICRwb2xpY3lTdGF0dXNbJHN0YXR1cy5Ub1N0cmluZygpXSAtY29sb3IgIkdy
ZWVuIiAtZGF0ZSAkY3VycmVudEtleS5QU0NoaWxkTmFtZQ0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBlbHNlIHsNCiAgICAgICAgICAg
ICAgICAgICAgaWYgKCRkaXNwbGF5KSB7IFdyaXRlLUhvc3QgIiBQb2xpY3kgJF8gOiAkc3RhdHVzICgkKCRwb2xpY3lTdGF0dXNbJHN0YXR1cy5Ub1N0cmlu
ZygpXSkpIiAtRm9yZWdyb3VuZENvbG9yIFllbGxvdyB9DQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgfQ0KICAgICAgICB9DQoNCiAgICB9DQoN
CiAgICBGdW5jdGlvbiBQcm9jZXNzQ2VydHMoKSB7DQogICAgICAgIHBhcmFtDQogICAgICAgICgNCiAgICAgICAgICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5
ID0gJHRydWUsIFZhbHVlRnJvbVBpcGVsaW5lID0gJFRydWUpXSBbTWljcm9zb2Z0LldpbjMyLlJlZ2lzdHJ5S2V5XSAkY3VycmVudEtleSwNCiAgICAgICAg
ICAgIFtQYXJhbWV0ZXIoTWFuZGF0b3J5ID0gJHRydWUsIFZhbHVlRnJvbVBpcGVsaW5lQnlQcm9wZXJ0eU5hbWUgPSAkVHJ1ZSldIFtib29sXSAkZGlzcGxh
eQ0KICAgICAgICApDQoNCiAgICAgICAgQmVnaW4gew0KICAgICAgICAgICAgaWYgKCRkaXNwbGF5KSB7IFdyaXRlLUhvc3QgIkNlcnRpZmljYXRlczoiIH0N
CiAgICAgICAgfQ0KDQogICAgICAgIFByb2Nlc3Mgew0KICAgICAgICAgICAgaWYgKCRkaXNwbGF5KSB7IFdyaXRlLUhvc3QgIiAkKChbZGF0ZXRpbWVdJGN1
cnJlbnRLZXkuUFNDaGlsZE5hbWUpLlRvU3RyaW5nKCd1JykpIiB9DQogICAgICAgICAgICAkY3VycmVudEtleS5Qcm9wZXJ0eSB8ICUgew0KICAgICAgICAg
ICAgICAgICRjZXJ0S2V5ID0gW1VSSV06OlVuZXNjYXBlRGF0YVN0cmluZygoJF8uU3BsaXQoIi8iKSlbNl0pDQogICAgICAgICAgICAgICAgJHN0YXR1cyA9
IEdldC1JdGVtUHJvcGVydHlWYWx1ZSAtcGF0aCAkY3VycmVudEtleS5QU1BhdGggLU5hbWUgJF8NCiAgICAgICAgICAgICAgICBpZiAoJE9ubGluZSkgew0K
ICAgICAgICAgICAgICAgICAgICAkZm91bmQgPSAkcG9saWNpZXMgfCA/IHsgJGNlcnRLZXkuUmVwbGFjZSgiXyIsICItIikgLW1hdGNoICRfLklkIH0NCiAg
ICAgICAgICAgICAgICAgICAgJGNlcnRLZXkgPSAiJCgkZm91bmQuRGlzcGxheU5hbWUpICgkKCRjZXJ0S2V5KSkiDQogICAgICAgICAgICAgICAgfQ0KICAg
ICAgICAgICAgICAgIGlmICgkc3RhdHVzIC1lcSAiMSIpIHsNCiAgICAgICAgICAgICAgICAgICAgaWYgKCRkaXNwbGF5KSB7IFdyaXRlLUhvc3QgIiBDZXJ0
ICRjZXJ0S2V5IDogJHN0YXR1cyAoJCgkcG9saWN5U3RhdHVzWyRzdGF0dXMuVG9TdHJpbmcoKV0pKSIgLUZvcmVncm91bmRDb2xvciBHcmVlbiB9DQogICAg
ICAgICAgICAgICAgICAgIFJlY29yZFN0YXR1cyAtZGV0YWlsICJDZXJ0ICRjZXJ0S2V5IiAtc3RhdHVzICRwb2xpY3lTdGF0dXNbJHN0YXR1cy5Ub1N0cmlu
ZygpXSAtY29sb3IgIkdyZWVuIiAtZGF0ZSAkY3VycmVudEtleS5QU0NoaWxkTmFtZQ0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAgICBlbHNl
IHsNCiAgICAgICAgICAgICAgICAgICAgaWYgKCRkaXNwbGF5KSB7IFdyaXRlLUhvc3QgIiBDZXJ0ICRjZXJ0S2V5IDogJHN0YXR1cyAoJCgkcG9saWN5U3Rh
dHVzWyRzdGF0dXMuVG9TdHJpbmcoKV0pKSIgLUZvcmVncm91bmRDb2xvciBZZWxsb3cgfQ0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgIH0NCiAg
ICAgICAgfQ0KDQogICAgfQ0KDQogICAgRnVuY3Rpb24gUHJvY2Vzc05vZGVDYWNoZSgpIHsNCg0KICAgICAgICBQcm9jZXNzIHsNCiAgICAgICAgICAgICRu
b2RlQ291bnQgPSAwDQogICAgICAgICAgICB3aGlsZSAoJHRydWUpIHsNCiAgICAgICAgICAgICAgICAjIEdldCB0aGUgbm9kZXMgaW4gb3JkZXIuIFRoaXMg
d29uJ3Qgd29yayBhZnRlciBhIHdoaWxlIGJlY2F1c2UgdGhlIG9sZGVyIG51bWJlcnMgYXJlIGRlbGV0ZWQgYXMgbmV3IG9uZXMgYXJlIGFkZGVkDQogICAg
ICAgICAgICAgICAgIyBidXQgaXQgd2lsbCB3b3JrIG91dCBPSyBzaG9ydGx5IGFmdGVyIHByb3Zpc2lvbmluZy4gVGhlIGFsdGVybmF0aXZlIHdvdWxkIGJl
IHRvIGdldCBhbGwgdGhlIHN1YmtleXMgYW5kIHRoZW4gc29ydA0KICAgICAgICAgICAgICAgICMgdGhlbSBudW1lcmljYWxseSBpbnN0ZWFkIG9mIGFscGhh
YmV0aWNhbGx5LCBidXQgdGhhdCBjYW4gYmUgc2F2ZWQgZm9yIGxhdGVyLi4uDQogICAgICAgICAgICAgICAgJG5vZGUgPSBHZXQtSXRlbVByb3BlcnR5ICIk
cHJvdmlzaW9uaW5nUGF0aFxOb2RlQ2FjaGVcQ1NQXERldmljZVxNUyBETSBTZXJ2ZXJcTm9kZXNcJG5vZGVDb3VudCIgLUVycm9yQWN0aW9uIFNpbGVudGx5
Q29udGludWUNCiAgICAgICAgICAgICAgICBpZiAoJG5vZGUgLWVxICRudWxsKSB7DQogICAgICAgICAgICAgICAgICAgIGJyZWFrDQogICAgICAgICAgICAg
ICAgfQ0KICAgICAgICAgICAgICAgICRub2RlQ291bnQgKz0gMQ0KICAgICAgICAgICAgICAgICRub2RlIHwgU2VsZWN0IE5vZGVVcmksIEV4cGVjdGVkVmFs
dWUNCiAgICAgICAgICAgIH0NCiAgICAgICAgfQ0KDQogICAgfQ0KDQogICAgRnVuY3Rpb24gVHJpbU1TSSgpIHsNCiAgICAgICAgcGFyYW0gKA0KICAgICAg
ICAgICAgW29iamVjdF0gJGUsDQogICAgICAgICAgICBbc3RyaW5nXSAkc2lkZWNhclByb2R1Y3RDb2RlDQogICAgICAgICkNCg0KICAgICAgICAjIEZpeCB1
cCB0aGUgbmFtZQ0KICAgICAgICBpZiAoJGV2ZW50LklkIC1lcSAxOTI0KSB7DQogICAgICAgICAgICAkciA9ICRldmVudC5Qcm9wZXJ0aWVzWzJdLlZhbHVl
DQogICAgICAgIH0gZWxzZSB7DQogICAgICAgICAgICAkciA9ICRldmVudC5Qcm9wZXJ0aWVzWzBdLlZhbHVlDQogICAgICAgIH0NCiAgICAgICAgJHByb2R1
Y3RDb2RlID0gJHIuUmVwbGFjZSgieyIsIiIpLlJlcGxhY2UoIn0iLCIiKQ0KICAgICAgICBpZiAoJHByb2R1Y3RDb2RlIC1lcSAkc2lkZWNhclByb2R1Y3RD
b2RlKSB7DQogICAgICAgICAgICByZXR1cm4gIkludHVuZSBNYW5hZ2VtZW50IEV4dGVuc2lvbiAoJHIpIg0KICAgICAgICB9DQoNCiAgICAgICAgIyBTZWUg
aWYgd2UgY2FuIGZpbmQgdGhlIHJlYWwgbmFtZQ0KICAgICAgICBpZiAoVGVzdC1QYXRoICJIS0xNOlxTb2Z0d2FyZVxNaWNyb3NvZnRcV2luZG93c1xDdXJy
ZW50VmVyc2lvblxVbmluc3RhbGxce3skcHJvZHVjdENvZGV9fSIpIHsNCiAgICAgICAgICAgICRkaXNwbGF5TmFtZSA9IChHZXQtSXRlbVByb3BlcnR5IC1Q
YXRoICJIS0xNOlxTb2Z0d2FyZVxNaWNyb3NvZnRcV2luZG93c1xDdXJyZW50VmVyc2lvblxVbmluc3RhbGxcJHByb2R1Y3RDb2RlIikuRGlzcGxheU5hbWUN
CiAgICAgICAgICAgIHJldHVybiAiJGRpc3BsYXlOYW1lICgkcikiDQogICAgICAgIH0NCiAgICAgICAgZWxzZSB7DQogICAgICAgICAgICByZXR1cm4gJHIN
CiAgICAgICAgfQ0KDQogICAgfQ0KDQogICAgRnVuY3Rpb24gUHJvY2Vzc0V2ZW50cygpIHsNCg0KICAgICAgICBQcm9jZXNzIHsNCg0KICAgICAgICAgICAg
JHByb2R1Y3RDb2RlID0gJ0lNRS1Ob3QtWWV0LUluc3RhbGxlZCcNCiAgICAgICAgICAgIGlmIChUZXN0LVBhdGggIiRtc2lQYXRoXFMtMC0wLTAwLTAwMDAw
MDAwMDAtMDAwMDAwMDAwMC0wMDAwMDAwMDAtMDAwXE1TSSIpIHsNCiAgICAgICAgICAgICAgICBHZXQtQ2hpbGRJdGVtIC1wYXRoICIkbXNpUGF0aFxTLTAt
MC0wMC0wMDAwMDAwMDAwLTAwMDAwMDAwMDAtMDAwMDAwMDAwLTAwMFxNU0kiIHwgJSB7DQogICAgICAgICAgICAgICAgICAgICRmaWxlID0gKEdldC1JdGVt
UHJvcGVydHkgLVBhdGggJF8uUFNQYXRoKS5DdXJyZW50RG93bmxvYWRVcmwNCiAgICAgICAgICAgICAgICAgICAgaWYgKCRmaWxlIC1tYXRjaCAiSW50dW5l
V2luZG93c0FnZW50Lm1zaSIpIHsNCiAgICAgICAgICAgICAgICAgICAgICAgICRwcm9kdWN0Q29kZSA9IChHZXQtSXRlbVByb3BlcnR5VmFsdWUgLVBhdGgg
JF8uUFNQYXRoIC1OYW1lIFByb2R1Y3RDb2RlKS5SZXBsYWNlKCJ7IiwiIikuUmVwbGFjZSgifSIsIiIpDQogICAgICAgICAgICAgICAgICAgIH0NCiAgICAg
ICAgICAgICAgICB9DQogICAgICAgICAgICB9DQoNCiAgICAgICAgICAgICMgUHJvY2VzcyBkZXZpY2UgbWFuYWdlbWVudCBldmVudHMNCiAgICAgICAgICAg
IGlmICgkc2NyaXB0OnVzZUZpbGUpIHsNCiAgICAgICAgICAgICAgICAkZXZlbnRzID0gR2V0LVdpbkV2ZW50IC1QYXRoICIkKCRlbnY6VEVNUClcRVNQU3Rh
dHVzLnRtcFxtaWNyb3NvZnQtd2luZG93cy1kZXZpY2VtYW5hZ2VtZW50LWVudGVycHJpc2UtZGlhZ25vc3RpY3MtcHJvdmlkZXItYWRtaW4uZXZ0eCIgLU9s
ZGVzdCB8ID8geyAoJF8uSWQgLWluIDE5MDUsIDE5MDYsIDE5MjAsIDE5MjIsIDE5MjQpIC1vciAkXy5JZCAtaW4gKDcyLCAxMDAsIDEwNywgMTA5LCAxMTAs
IDExMSkgfQ0KICAgICAgICAgICAgfQ0KICAgICAgICAgICAgZWxzZSB7DQogICAgICAgICAgICAgICAgJGV2ZW50cyA9IEdldC1XaW5FdmVudCAtTG9nTmFt
ZSBNaWNyb3NvZnQtV2luZG93cy1EZXZpY2VNYW5hZ2VtZW50LUVudGVycHJpc2UtRGlhZ25vc3RpY3MtUHJvdmlkZXIvQWRtaW4gLU9sZGVzdCB8ID8geyAo
JF8uSWQgLWluIDE5MDUsIDE5MDYsIDE5MjAsIDE5MjIpIC1vciAkXy5JZCAtaW4gKDcyLCAxMDAsIDEwNywgMTA5LCAxMTAsIDExMSkgfQ0KICAgICAgICAg
ICAgfQ0KICAgICAgICAgICAgJGV2ZW50cyB8ICUgew0KICAgICAgICAgICAgICAgICRtZXNzYWdlID0gJF8uTWVzc2FnZQ0KICAgICAgICAgICAgICAgICRk
ZXRhaWwgPSAiU2lkZWNhciINCiAgICAgICAgICAgICAgICAkY29sb3IgPSAiWWVsbG93Ig0KICAgICAgICAgICAgICAgICRldmVudCA9ICRfDQogICAgICAg
ICAgICAgICAgc3dpdGNoICgkXy5pZCkgew0KICAgICAgICAgICAgICAgICAgICB7ICRfIC1pbiAoMTEwLCAxMDkpIH0geyANCiAgICAgICAgICAgICAgICAg
ICAgICAgICRkZXRhaWwgPSAiT2ZmbGluZSBEb21haW4gSm9pbiINCiAgICAgICAgICAgICAgICAgICAgICAgIHN3aXRjaCAoJGV2ZW50LlByb3BlcnRpZXNb
MF0uVmFsdWUpIHsNCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAwIHsgJG1lc3NhZ2UgPSAiT2ZmbGluZSBkb21haW4gam9pbiBub3QgY29uZmlndXJl
ZCIgfQ0KICAgICAgICAgICAgICAgICAgICAgICAgICAgIDEgeyAkbWVzc2FnZSA9ICJXYWl0aW5nIGZvciBPREogYmxvYiIgfQ0KICAgICAgICAgICAgICAg
ICAgICAgICAgICAgIDIgeyAkbWVzc2FnZSA9ICJQcm9jZXNzZWQgT0RKIGJsb2IiIH0NCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAzIHsgJG1lc3Nh
Z2UgPSAiVGltZWQgb3V0IHdhaXRpbmcgZm9yIE9ESiBibG9iIG9yIGNvbm5lY3Rpdml0eSIgfQ0KICAgICAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAg
ICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgICAgIDExMSB7ICRkZXRhaWwgPSAiT2ZmbGluZSBEb21haW4gSm9pbiI7ICRtZXNzYWdlID0gIlN0
YXJ0aW5nIHdhaXQgZm9yIE9ESiBibG9iIiB9DQogICAgICAgICAgICAgICAgICAgIDEwNyB7ICRkZXRhaWwgPSAiT2ZmbGluZSBEb21haW4gSm9pbiI7ICRt
ZXNzYWdlID0gIlN1Y2Nlc3NmdWxseSBhcHBsaWVkIE9ESiBibG9iIiB9DQogICAgICAgICAgICAgICAgICAgIDEwMCB7ICRkZXRhaWwgPSAiT2ZmbGluZSBE
b21haW4gSm9pbiI7ICRtZXNzYWdlID0gIkNvdWxkIG5vdCBlc3RhYmxpc2ggY29ubmVjdGl2aXR5IjsgJGNvbG9yID0gIlJlZCIgfQ0KICAgICAgICAgICAg
ICAgICAgICA3MiB7ICRkZXRhaWwgPSAiTURNIEVucm9sbG1lbnQiIH0NCiAgICAgICAgICAgICAgICAgICAgMTkwNSB7ICRkZXRhaWwgPSAoVHJpbU1TSSAk
ZXZlbnQgJHByb2R1Y3RDb2RlKTsgJG1lc3NhZ2UgPSAiRG93bmxvYWQgc3RhcnRlZCIgfQ0KICAgICAgICAgICAgICAgICAgICAxOTA2IHsgJGRldGFpbCA9
IChUcmltTVNJICRldmVudCAkcHJvZHVjdENvZGUpOyAkbWVzc2FnZSA9ICJEb3dubG9hZCBmaW5pc2hlZCIgfQ0KICAgICAgICAgICAgICAgICAgICAxOTIw
IHsgJGRldGFpbCA9IChUcmltTVNJICRldmVudCAkcHJvZHVjdENvZGUpOyAkbWVzc2FnZSA9ICJJbnN0YWxsYXRpb24gc3RhcnRlZCIgfQ0KICAgICAgICAg
ICAgICAgICAgICAxOTIyIHsgJGRldGFpbCA9IChUcmltTVNJICRldmVudCAkcHJvZHVjdENvZGUpOyAkbWVzc2FnZSA9ICJJbnN0YWxsYXRpb24gZmluaXNo
ZWQiIH0NCiAgICAgICAgICAgICAgICAgICAgMTkyNCB7ICRkZXRhaWwgPSAoVHJpbU1TSSAkZXZlbnQgJHByb2R1Y3RDb2RlKTsgJG1lc3NhZ2UgPSAiSW5z
dGFsbGF0aW9uIGZhaWxlZCI7ICRjb2xvciA9ICJSZWQiIH0NCiAgICAgICAgICAgICAgICAgICAgeyAkXyAtaW4gKDE5MjIsIDcyKSB9IHsgJGNvbG9yID0g
IkdyZWVuIiB9DQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgIFJlY29yZFN0YXR1cyAtZGV0YWlsICRkZXRhaWwgLWRhdGUgJF8uVGltZUNy
ZWF0ZWQuVG9Vbml2ZXJzYWxUaW1lKCkgLXN0YXR1cyAkbWVzc2FnZSAtY29sb3IgJGNvbG9yDQogICAgICAgICAgICB9DQoNCiAgICAgICAgICAgICMgUHJv
Y2VzcyBkZXZpY2UgcmVnaXN0cmF0aW9uIGV2ZW50cw0KICAgICAgICAgICAgaWYgKCRzY3JpcHQ6dXNlRmlsZSkgew0KICAgICAgICAgICAgICAgICRldmVu
dHMgPSBHZXQtV2luRXZlbnQgLVBhdGggIiQoJGVudjpURU1QKVxFU1BTdGF0dXMudG1wXG1pY3Jvc29mdC13aW5kb3dzLXVzZXIgZGV2aWNlIHJlZ2lzdHJh
dGlvbi1hZG1pbi5ldnR4IiAtT2xkZXN0IHwgPyB7ICRfLklkIC1pbiAoMzA2LCAxMDEpIH0NCiAgICAgICAgICAgIH0NCiAgICAgICAgICAgIGVsc2Ugew0K
ICAgICAgICAgICAgICAgIHRyeSB7DQogICAgICAgICAgICAgICAgICAgICRldmVudHMgPSBHZXQtV2luRXZlbnQgLUxvZ05hbWUgJ01pY3Jvc29mdC1XaW5k
b3dzLVVzZXIgRGV2aWNlIFJlZ2lzdHJhdGlvbi9BZG1pbicgLU9sZGVzdCAtRXJyb3JBY3Rpb24gU3RvcCB8ID8geyAkXy5JZCAtaW4gKDMwNiwgMTAxKSB9
DQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgIGNhdGNoIFtFeGNlcHRpb25dIHsNCiAgICAgICAgICAgICAgICAgICAgaWYgKCRfLkZ1bGx5
UXVhbGlmaWVkRXJyb3JJZCAtbWF0Y2ggIk5vTWF0Y2hpbmdFdmVudHNGb3VuZCIpIHsNCiAgICAgICAgICAgICAgICAgICAgICAgICRldmVudHMgPSBAKCkN
CiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgIH0NCiAgICAgICAgICAgICRldmVudHMgfCAlIHsNCiAgICAg
ICAgICAgICAgICAkbWVzc2FnZSA9ICRfLk1lc3NhZ2UNCiAgICAgICAgICAgICAgICAkZGV0YWlsID0gIkRldmljZSBSZWdpc3RyYXRpb24iDQogICAgICAg
ICAgICAgICAgJGNvbG9yID0gIlllbGxvdyINCiAgICAgICAgICAgICAgICAkZXZlbnQgPSAkXw0KICAgICAgICAgICAgICAgIHN3aXRjaCAoJF8uaWQpIHsN
CiAgICAgICAgICAgICAgICAgICAgMTAxIHsgJGRldGFpbCA9ICJEZXZpY2UgUmVnaXN0cmF0aW9uIjsgJG1lc3NhZ2UgPSAiU0NQIGRpc2NvdmVyeSBzdWNj
ZXNzZnVsIiB9DQogICAgICAgICAgICAgICAgICAgIDMwNCB7ICRkZXRhaWwgPSAiRGV2aWNlIFJlZ2lzdHJhdGlvbiI7ICRtZXNzYWdlID0gIkh5YnJpZCBB
QURKIGRldmljZSByZWdpc3RyYXRpb24gZmFpbGVkIiB9DQogICAgICAgICAgICAgICAgICAgIDMwNiB7ICRkZXRhaWwgPSAiRGV2aWNlIFJlZ2lzdHJhdGlv
biI7ICRtZXNzYWdlID0gIkh5YnJpZCBBQURKIGRldmljZSByZWdpc3RyYXRpb24gc3VjY2VlZGVkIjsgJGNvbG9yID0gJ0dyZWVuJyB9DQogICAgICAgICAg
ICAgICAgfQ0KICAgICAgICAgICAgICAgIFJlY29yZFN0YXR1cyAtZGV0YWlsICRkZXRhaWwgLWRhdGUgJF8uVGltZUNyZWF0ZWQuVG9Vbml2ZXJzYWxUaW1l
KCkgLXN0YXR1cyAkbWVzc2FnZSAtY29sb3IgJGNvbG9yDQogICAgICAgICAgICB9DQoNCiAgICAgICAgICAgICMgQWRkIERPIGV2ZW50cyBmb3IgT2ZmaWNl
IGNsaWNrLXRvLXJ1biBkb3dubG9hZHMNCiAgICAgICAgICAgIGlmICgtbm90ICRzY3JpcHQ6dXNlRmlsZSkgew0KICAgICAgICAgICAgICAgIEdldC1EZWxp
dmVyeU9wdGltaXphdGlvbkxvZyB8IFdoZXJlLU9iamVjdCB7ICRfLkZ1bmN0aW9uIC1tYXRjaCAiKERvd25sb2FkU3RhcnQpfChEb3dubG9hZENvbXBsZXRl
ZCkiIC1hbmQgJF8uTWVzc2FnZSAtbGlrZSAiKk1pY3Jvc29mdCBPZmZpY2UgQ2xpY2stdG8tUnVuKiIgfSB8IEZvckVhY2gtT2JqZWN0IHsNCiAgICAgICAg
ICAgICAgICAgICAgIyBFeHRyYWN0IHRoZSBmaWxlIElEIGJlY2F1c2Ugd2Ugd2FudCB0byBsaXN0IGVhY2ggZmlsZSBkb3dubG9hZGVkDQogICAgICAgICAg
ICAgICAgICAgICRmaWxlSWQgPSAiIg0KICAgICAgICAgICAgICAgICAgICAkZmlsZUlkU3RhcnQgPSAkXy5NZXNzYWdlLkluZGV4T2YoImZpbGVJZDogIikN
CiAgICAgICAgICAgICAgICAgICAgaWYgKCRmaWxlSWRTdGFydCAtZXEgLTEpIHsNCiAgICAgICAgICAgICAgICAgICAgICAgICMgTWlnaHQgYmUgdXNpbmcg
ImZpbGVJZCA9ICIsIGJlY2F1c2UgdGhpcyBETyBldmVudCBpbmZvcm1hdGlvbiBzdWNrcw0KICAgICAgICAgICAgICAgICAgICAgICAgJGZpbGVJZFN0YXJ0
ID0gJF8uTWVzc2FnZS5JbmRleE9mKCJmaWxlSWQgPSAiKQ0KICAgICAgICAgICAgICAgICAgICAgICAgJHNraXAgPSA5DQogICAgICAgICAgICAgICAgICAg
IH0gZWxzZSB7DQogICAgICAgICAgICAgICAgICAgICAgICAkc2tpcCA9IDgNCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICBp
ZiAoJGZpbGVJZFN0YXJ0IC1ndCAwKSB7DQogICAgICAgICAgICAgICAgICAgICAgICAjIEdldCBmcm9tIHRoZSBzdGFydCBvZiB0aGUgYWN0dWFsIElEDQog
ICAgICAgICAgICAgICAgICAgICAgICAkZmlsZUlkID0gJF8uTWVzc2FnZS5TdWJzdHJpbmcoJGZpbGVJZFN0YXJ0ICsgJHNraXApDQogICAgICAgICAgICAg
ICAgICAgICAgICAjIEZpbmQgdGhlIGVuZCBhbmQgY2hvcCBpdCBvZmYNCiAgICAgICAgICAgICAgICAgICAgICAgICRmaWxlSWRFbmQgPSAkZmlsZUlkLklu
ZGV4T2YoIiwiKQ0KICAgICAgICAgICAgICAgICAgICAgICAgJGZpbGVJZCA9ICRmaWxlSWQuU3Vic3RyaW5nKDAsICRmaWxlSWRFbmQpDQogICAgICAgICAg
ICAgICAgICAgICAgICAjIFJlbW92ZSB0aGUgZXh0cmEgR1VJRCBmcm9tIHRoZSBiZWdpbm5pbmcNCiAgICAgICAgICAgICAgICAgICAgICAgICRmaWxlSWQg
PSAkZmlsZUlkLlN1YnN0cmluZygzNykNCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICBpZiAoJF8uRnVuY3Rpb24uQ29udGFp
bnMoIkRvd25sb2FkU3RhcnQiKSkgDQogICAgICAgICAgICAgICAgICAgIHsNCiAgICAgICAgICAgICAgICAgICAgICAgICRvcCA9ICJEb3dubG9hZFN0YXJ0
Ig0KICAgICAgICAgICAgICAgICAgICB9IGVsc2Ugew0KICAgICAgICAgICAgICAgICAgICAgICAgJG9wID0gIkRvd25sb2FkQ29tcGxldGVkIg0KICAgICAg
ICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgICAgIFJlY29yZFN0YXR1cyAtZGV0YWlsICJNaWNyb3NvZnQgT2ZmaWNlIEMyUiAoJGZpbGVJZCki
IC1zdGF0dXMgJG9wIC1jb2xvciAiWWVsbG93IiAtZGF0ZSAkXy5UaW1lQ3JlYXRlZA0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgIH0NCg0KICAg
ICAgICB9DQogICAgDQogICAgfQ0KICAgIA0KICAgICMtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0NCiAgICAjIE1haW4gY29kZQ0KICAgICMtLS0tLS0tLS0t
LS0tLS0tLS0tLS0tLS0NCg0KICAgICRzY3JpcHQ6b2JzZXJ2ZWRUaW1lbGluZSA9IEAoKQ0KDQogICAgIyBJZiBvbmxpbmUsIG1ha2Ugc3VyZSB3ZSBhcmUg
YWJsZSB0byBhdXRoZW50aWNhdGUNCiAgICBpZiAoJE9ubGluZSkgew0KDQogICAgICAgICMjQ2hlY2sgaWYgd2UgbmVlZCB0byBpbnN0YWxsIHRoZSBtb2R1
bGUNCiAgICAgICAgI0luc3RhbGwgTVMgR3JhcGggaWYgbm90IGF2YWlsYWJsZQ0KICAgICAgICBpZiAoR2V0LU1vZHVsZSAtTGlzdEF2YWlsYWJsZSAtTmFt
ZSBNaWNyb3NvZnQuR3JhcGguQXV0aGVudGljYXRpb24pIHsNCiAgICAgICAgICAgIFdyaXRlLVZlcmJvc2UgIk1pY3Jvc29mdCBHcmFwaCBhbHJlYWR5IGlu
c3RhbGxlZCINCiAgICAgICAgfSANCiAgICAgICAgZWxzZSB7DQogICAgICAgICAgICB0cnkgew0KICAgICAgICAgICAgICAgIEluc3RhbGwtTW9kdWxlIC1O
YW1lIE1pY3Jvc29mdC5HcmFwaC5BdXRoZW50aWNhdGlvbiAtUmVwb3NpdG9yeSBQU0dhbGxlcnkgLUZvcmNlIA0KICAgICAgICAgICAgfQ0KICAgICAgICAg
ICAgY2F0Y2ggW0V4Y2VwdGlvbl0gew0KICAgICAgICAgICAgICAgICRfLm1lc3NhZ2UgDQogICAgICAgICAgICB9DQogICAgICAgIH0NCg0KICAgICAgICAj
Q29ubmVjdCB0byBHcmFwaA0KICAgICAgICBpZiAoJEFwcElkIC1hbmQgJEFwcFNlY3JldCAtYW5kICR0ZW5hbnQpIHsNCiAgICAgICAgICAgICRncmFwaCA9
IENvbm5lY3QtVG9HcmFwaCAtVGVuYW50ICR0ZW5hbnQgLUFwcElkICRjbGllbnRpZCAtQXBwU2VjcmV0ICRjbGllbnRzZWNyZXQNCiAgICAgICAgfQ0KICAg
ICAgICBlbHNlaWYgKCRCZWFyZXIpIHsNCiAgICAgICAgICAgICRncmFwaCA9IENvbm5lY3QtVG9HcmFwaCAtYmVhcmVyICRCZWFyZXINCiAgICAgICAgfQ0K
ICAgICAgICBlbHNlIHsNCiAgICAgICAgICAgICRncmFwaCA9IENvbm5lY3QtVG9HcmFwaCAtU2NvcGVzICJEZXZpY2VNYW5hZ2VtZW50QXBwcy5SZWFkLkFs
bCwgRGV2aWNlTWFuYWdlbWVudENvbmZpZ3VyYXRpb24uUmVhZC5BbGwiDQogICAgICAgIH0NCg0KICAgICAgICAjIEdldCBhIGxpc3Qgb2YgYXBwcw0KICAg
ICAgICBXcml0ZS1Ib3N0ICJHZXR0aW5nIGxpc3Qgb2YgYXBwcyINCiAgICAgICAgIyRzY3JpcHQ6YXBwcyA9IEdldC1NZ0RldmljZUFwcE1hbmFnZW1lbnRN
b2JpbGVBcHAgLUFsbA0KICAgICAgICAkYXBwc3VyaSA9ICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vYmV0YS9kZXZpY2VBcHBNYW5hZ2VtZW50L21v
YmlsZUFwcHMiDQogICAgICAgICRzY3JpcHQ6YXBwcyA9IGdldGFsbHBhZ2luYXRpb24gLXVybCAkYXBwc3VyaQ0KICAgICAgICANCiAgICAgICAgIyBHZXQg
YSBsaXN0IG9mIHBvbGljaWVzIChmb3IgY2VydHMpDQogICAgICAgIFdyaXRlLUhvc3QgIkdldHRpbmcgbGlzdCBvZiBwb2xpY2llcyINCiAgICAgICAgJGNv
bmZpZ3VyaSA9ICJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20vYmV0YS9kZXZpY2VNYW5hZ2VtZW50L2NvbmZpZ3VyYXRpb25Qb2xpY2llcyINCiAgICAg
ICAgIyRzY3JpcHQ6cG9saWNpZXMgPSBHZXQtTWdCZXRhRGV2aWNlTWFuYWdlbWVudENvbmZpZ3VyYXRpb25Qb2xpY3kgLUFsbA0KICAgICAgICAkc2NyaXB0
OnBvbGljaWVzID0gZ2V0YWxscGFnaW5hdGlvbiAtdXJsICRjb25maWd1cmkNCg0KICAgICAgICAjIEdldCBhIGxpc3Qgb2YgcGxhdGZvcm0gc2NyaXB0cw0K
ICAgICAgICBXcml0ZS1Ib3N0ICJHZXR0aW5nIGxpc3Qgb2Ygc2NyaXB0cyINCiAgICAgICAgJHNjcmlwdHN1cmkgPSAiaHR0cHM6Ly9ncmFwaC5taWNyb3Nv
ZnQuY29tL2JldGEvZGV2aWNlTWFuYWdlbWVudC9kZXZpY2VNYW5hZ2VtZW50U2NyaXB0cyINCiAgICAgICAgIyRzY3JpcHQ6cG9saWNpZXMgPSBHZXQtTWdC
ZXRhRGV2aWNlTWFuYWdlbWVudENvbmZpZ3VyYXRpb25Qb2xpY3kgLUFsbA0KICAgICAgICAkc2NyaXB0OnNjcmlwdHMgPSBnZXRhbGxwYWdpbmF0aW9uIC11
cmwgJHNjcmlwdHN1cmkNCiAgICB9DQoNCiAgICAjIFByb2Nlc3MgbG9nIGZpbGVzIGlmIG5lZWRlZA0KICAgICRzY3JpcHQ6dXNlRmlsZSA9ICRmYWxzZQ0K
ICAgIGlmICgkRmlsZSkgew0KDQogICAgICAgIFdyaXRlLUhvc3QgIlVzaW5nIGNvbnRlbnRzIG9mIGZpbGU6ICRGaWxlIg0KICAgICAgICBpZiAoVGVzdC1Q
YXRoICIkKCRlbnY6VEVNUClcRVNQU3RhdHVzLnRtcCIpIHsNCiAgICAgICAgICAgIFJlbW92ZS1JdGVtICIkKCRlbnY6VEVNUClcRVNQU3RhdHVzLnRtcCIg
LVJlY3Vyc2UgLUZvcmNlDQogICAgICAgIH0NCiAgICAgICAgTmV3LUl0ZW0gLVBhdGggIiQoJGVudjpURU1QKVxFU1BTdGF0dXMudG1wIiAtSXRlbVR5cGUg
ImRpcmVjdG9yeSIgfCBPdXQtTnVsbA0KICAgICAgICAkc2NyaXB0OnVzZUZpbGUgPSAkdHJ1ZQ0KDQogICAgICAgICMgSWYgdXNpbmcgYSBDQUIgZmlsZSwg
ZXh0cmFjdCB0aGUgbmVlZGVkIGZpbGVzIGZyb20gaXQNCiAgICAgICAgaWYgKCRGaWxlLlRvTG93ZXIoKS5FbmRzV2l0aCgiLmNhYiIpKSB7DQogICAgICAg
ICAgICAkbnVsbCA9ICYgZXhwYW5kLmV4ZSAiJEZpbGUiIC1GOiogIiQoJGVudjpURU1QKVxFU1BTdGF0dXMudG1wXCIgDQogICAgICAgIH0NCiAgICAgICAg
ZWxzZSB7DQogICAgICAgICAgICAjIElmIHVzaW5nIGEgWklQIGZpbGUsIGp1c3QgZXh0cmFjdCB0aGUgZW50aXJlIGNvbnRlbnRzIChub3QgYXMgZWFzeSB0
byBkbyBzZWxlY3RlZCBmaWxlcykNCiAgICAgICAgICAgIEV4cGFuZC1BcmNoaXZlIC1QYXRoICRGaWxlIC1EZXN0aW5hdGlvblBhdGggIiQoJGVudjpURU1Q
KVxFU1BTdGF0dXMudG1wXCINCiAgICAgICAgICAgICMgSWYgdGhpcyBpcyBhbiBJbnR1bmUgZGlhZ25vc3RpY3MgemlwLCB0aGUgInJlYWwiIGxvZ3MgYXJl
IGJ1cmllZCBkZWVwZXIuICBJZiB3ZSBjYW4gZmluZCB0aGVtLCBleHRyYWN0IHRoZW0NCiAgICAgICAgICAgICRyZWFsRmlsZSA9IEdldC1DaGlsZEl0ZW0g
IiQoJGVudjpURU1QKVxFU1BTdGF0dXMudG1wIiAtRmlsdGVyICIqTURNRGlhZ25vc3RpY3MqX2NhYiIgfCBHZXQtQ2hpbGRJdGVtDQogICAgICAgICAgICBp
ZiAoJHJlYWxGaWxlKSB7DQogICAgICAgICAgICAgICAgIyBFeHBhbmQgdGhlbSBpbnRvIHRoZSB0ZW1wIGZvbGRlciAtLSBjcmVhdGVzIGEgYml0IG9mIGEg
bWVzcywgYnV0IGl0J3MgYSB0ZW1wb3JhcnkgbWVzcy4uLg0KICAgICAgICAgICAgICAgICRudWxsID0gJiBleHBhbmQuZXhlICIkKCRyZWFsRmlsZS5GdWxs
TmFtZSkiIC1GOiogIiQoJGVudjpURU1QKVxFU1BTdGF0dXMudG1wXCINCiAgICAgICAgICAgIH0NCiAgICAgICAgfQ0KDQogICAgICAgICMgR2V0IHRoZSBo
YXJkd2FyZSBoYXNoIGluZm9ybWF0aW9uDQogICAgICAgIEdldC1DaGlsZEl0ZW0gIiQoJGVudjpURU1QKVxFU1BTdGF0dXMudG1wIiAtRmlsdGVyICIqLmNz
diIgfCBGb3JFYWNoLU9iamVjdCB7DQogICAgICAgICAgICAkY3N2ID0gR2V0LUNvbnRlbnQgJF8uRnVsbE5hbWUgfCBDb252ZXJ0RnJvbS1Dc3YNCiAgICAg
ICAgICAgICRoYXNoID0gJGNzdi4nSGFyZHdhcmUgSGFzaCcNCiAgICAgICAgfQ0KDQogICAgICAgICMgRWRpdCB0aGUgcGF0aCBpbiB0aGUgLnJlZyBmaWxl
DQogICAgICAgICRjb250ZW50ID0gR2V0LUNvbnRlbnQgLVBhdGggIiQoJGVudjpURU1QKVxFU1BTdGF0dXMudG1wXE1kbURpYWdSZXBvcnRfUmVnaXN0cnlE
dW1wLnJlZyINCiAgICAgICAgJGNvbnRlbnQgPSAkY29udGVudCAtcmVwbGFjZSAiXFtIS0VZX0NVUlJFTlRfVVNFUlxcIiwgIltIS0VZX0NVUlJFTlRfVVNF
UlxFU1BTdGF0dXMudG1wXFVTRVJcIg0KICAgICAgICAkY29udGVudCA9ICRjb250ZW50IC1yZXBsYWNlICJcW0hLRVlfTE9DQUxfTUFDSElORVxcIiwgIltI
S0VZX0NVUlJFTlRfVVNFUlxFU1BTdGF0dXMudG1wXE1BQ0hJTkVcIg0KICAgICAgICAkY29udGVudCA9ICRjb250ZW50IC1yZXBsYWNlICdeICInLCAnIicN
CiAgICAgICAgJGNvbnRlbnQgPSAkY29udGVudCAtcmVwbGFjZSAnXiBAJywgJ0AnDQogICAgICAgICRjb250ZW50ID0gJGNvbnRlbnQgLXJlcGxhY2UgJ0RX
T1JEOicsICdkd29yZDonDQoNCiAgICAgICAgJHN0cmVhbSA9IFtTeXN0ZW0uSU8uU3RyZWFtV3JpdGVyXSAiJCgkZW52OlRFTVApXEVTUFN0YXR1cy50bXBc
TWRtRGlhZ1JlcG9ydF9FZGl0ZWQucmVnIg0KICAgICAgICAkc3RyZWFtLldyaXRlTGluZSgiV2luZG93cyBSZWdpc3RyeSBFZGl0b3IgVmVyc2lvbiA1LjAw
YG4iKQ0KICAgICAgICAkY29udGVudCB8IEZvckVhY2gtT2JqZWN0IHsNCiAgICAgICAgICAgICMgRXNjYXBlIGJhY2tzbGFzaGVzIGFuZCBxdW90ZXMNCiAg
ICAgICAgICAgICRsaW5lID0gJF8uVHJpbSgpDQogICAgICAgICAgICAkdGV4dFN0YXJ0ID0gJGxpbmUuSW5kZXhPZigiIiI9IiIiKSArIDMNCiAgICAgICAg
ICAgIGlmICgkdGV4dFN0YXJ0IC1ndCAzKSB7DQogICAgICAgICAgICAgICAgJHRleHRMZW4gPSAkbGluZS5MZW5ndGggLSAkdGV4dFN0YXJ0DQogICAgICAg
ICAgICAgICAgJHRvRWRpdCA9ICRsaW5lLlN1YnN0cmluZygkdGV4dFN0YXJ0LCAkdGV4dExlbiAtIDEpDQogICAgICAgICAgICAgICAgJHRvRWRpdCA9ICR0
b0VkaXQuUmVwbGFjZSgnXCcsICdcXCcpDQogICAgICAgICAgICAgICAgJHRvRWRpdCA9ICR0b0VkaXQuUmVwbGFjZSgnIicsICdcIicpDQogICAgICAgICAg
ICAgICAgJGxpbmUgPSAiJCgkbGluZS5TdWJzdHJpbmcoMCwgJHRleHRTdGFydCkpJHRvRWRpdCIiIg0KICAgICAgICAgICAgfQ0KICAgICAgICAgICAgIyBB
cHBlbmQgaXQgdG8gdGhlIGZpbGUNCiAgICAgICAgICAgICMgV3JpdGUtSG9zdCAkbGluZQ0KICAgICAgICAgICAgJHN0cmVhbS5Xcml0ZUxpbmUoJGxpbmUp
DQogICAgICAgIH0NCiAgICAgICAgJHN0cmVhbS5DbG9zZSgpDQoNCiAgICAgICAgIyBSZW1vdmUgdGhlIHJlZ2lzdHJ5IGluZm8gaWYgaXQgZXhpc3RzDQog
ICAgICAgIGlmIChUZXN0LVBhdGggIkhLQ1U6XEVTUFN0YXR1cy50bXAiKSB7DQogICAgICAgICAgICBSZW1vdmUtSXRlbSAtUGF0aCAiSEtDVTpcRVNQU3Rh
dHVzLnRtcCIgLVJlY3Vyc2UgLUZvcmNlDQogICAgICAgIH0NCg0KICAgICAgICAjIEltcG9ydCB0aGUgLnJlZyBmaWxlDQogICAgICAgICRudWxsID0gJiBy
ZWcuZXhlIElNUE9SVCAiJCgkZW52OlRFTVApXEVTUFN0YXR1cy50bXBcTWRtRGlhZ1JlcG9ydF9FZGl0ZWQucmVnIiAyPiYxDQoNCiAgICAgICAgIyBDb25m
aWd1cmUgdGhlIChub3QgbGl2ZSkgY29uc3RhbnRzDQogICAgICAgICRzY3JpcHQ6cHJvdmlzaW9uaW5nUGF0aCA9ICJIS0NVOlxFU1BTdGF0dXMudG1wXE1B
Q0hJTkVcc29mdHdhcmVcbWljcm9zb2Z0XHByb3Zpc2lvbmluZyINCiAgICAgICAgJHNjcmlwdDphdXRvcGlsb3REaWFnUGF0aCA9ICJIS0NVOlxFU1BTdGF0
dXMudG1wXE1BQ0hJTkVcc29mdHdhcmVcbWljcm9zb2Z0XHByb3Zpc2lvbmluZ1xEaWFnbm9zdGljc1xBdXRvcGlsb3QiDQogICAgICAgICRzY3JpcHQ6b21h
ZG1QYXRoID0gIkhLQ1U6XEVTUFN0YXR1cy50bXBcTUFDSElORVxzb2Z0d2FyZVxtaWNyb3NvZnRccHJvdmlzaW9uaW5nXE9NQURNIg0KICAgICAgICAkc2Ny
aXB0OnBhdGggPSAiSEtDVTpcRVNQU3RhdHVzLnRtcFxNQUNISU5FXFNvZnR3YXJlXE1pY3Jvc29mdFxXaW5kb3dzXEF1dG9waWxvdFxFbnJvbGxtZW50U3Rh
dHVzVHJhY2tpbmdcRVNQVHJhY2tpbmdJbmZvXERpYWdub3N0aWNzIg0KICAgICAgICAkc2NyaXB0Om1zaVBhdGggPSAiSEtDVTpcRVNQU3RhdHVzLnRtcFxN
QUNISU5FXFNvZnR3YXJlXE1pY3Jvc29mdFxFbnRlcnByaXNlRGVza3RvcEFwcE1hbmFnZW1lbnQiDQogICAgICAgICRzY3JpcHQ6b2ZmaWNlUGF0aCA9ICJI
S0NVOlxFU1BTdGF0dXMudG1wXE1BQ0hJTkVcU29mdHdhcmVcTWljcm9zb2Z0XE9mZmljZUNTUCINCiAgICAgICAgJHNjcmlwdDpzaWRlY2FyUGF0aCA9ICJI
S0NVOlxFU1BTdGF0dXMudG1wXE1BQ0hJTkVcU29mdHdhcmVcTWljcm9zb2Z0XEludHVuZU1hbmFnZW1lbnRFeHRlbnNpb24iDQogICAgICAgICRzY3JpcHQ6
c2lkZWNhcldpbjMyQXBwcyA9ICJIS0NVOlxFU1BTdGF0dXMudG1wXE1BQ0hJTkVcU29mdHdhcmVcTWljcm9zb2Z0XEludHVuZU1hbmFnZW1lbnRFeHRlbnNp
b25cV2luMzJBcHBzIg0KICAgICAgICAkc2NyaXB0OmVucm9sbG1lbnRzUGF0aCA9ICJIS0NVOlxFU1BTdGF0dXMudG1wXE1BQ0hJTkVcc29mdHdhcmVcbWlj
cm9zb2Z0XGVucm9sbG1lbnRzIg0KICAgIH0NCiAgICBlbHNlIHsNCiAgICAgICAgIyBDb25maWd1cmUgbGl2ZSBjb25zdGFudHMNCiAgICAgICAgJHNjcmlw
dDpwcm92aXNpb25pbmdQYXRoID0gIkhLTE06XHNvZnR3YXJlXG1pY3Jvc29mdFxwcm92aXNpb25pbmciDQogICAgICAgICRzY3JpcHQ6YXV0b3BpbG90RGlh
Z1BhdGggPSAiSEtMTTpcc29mdHdhcmVcbWljcm9zb2Z0XHByb3Zpc2lvbmluZ1xEaWFnbm9zdGljc1xBdXRvcGlsb3QiDQogICAgICAgICRzY3JpcHQ6b21h
ZG1QYXRoID0gIkhLTE06XHNvZnR3YXJlXG1pY3Jvc29mdFxwcm92aXNpb25pbmdcT01BRE0iDQogICAgICAgICRzY3JpcHQ6cGF0aCA9ICJIS0xNOlxTb2Z0
d2FyZVxNaWNyb3NvZnRcV2luZG93c1xBdXRvcGlsb3RcRW5yb2xsbWVudFN0YXR1c1RyYWNraW5nXEVTUFRyYWNraW5nSW5mb1xEaWFnbm9zdGljcyINCiAg
ICAgICAgJHNjcmlwdDptc2lQYXRoID0gIkhLTE06XFNvZnR3YXJlXE1pY3Jvc29mdFxFbnRlcnByaXNlRGVza3RvcEFwcE1hbmFnZW1lbnQiDQogICAgICAg
ICRzY3JpcHQ6b2ZmaWNlUGF0aCA9ICJIS0xNOlxTb2Z0d2FyZVxNaWNyb3NvZnRcT2ZmaWNlQ1NQIg0KICAgICAgICAkc2NyaXB0OnNpZGVjYXJQYXRoID0g
IkhLTE06XFNvZnR3YXJlXE1pY3Jvc29mdFxJbnR1bmVNYW5hZ2VtZW50RXh0ZW5zaW9uIg0KICAgICAgICAkc2NyaXB0OnNpZGVjYXJXaW4zMkFwcHMgPSAi
SEtMTTpcU29mdHdhcmVcTWljcm9zb2Z0XEludHVuZU1hbmFnZW1lbnRFeHRlbnNpb25cV2luMzJBcHBzIg0KICAgICAgICAkc2NyaXB0OmVucm9sbG1lbnRz
UGF0aCA9ICJIS0xNOlxTb2Z0d2FyZVxNaWNyb3NvZnRcZW5yb2xsbWVudHMiDQoNCiAgICAgICAgJGhhc2ggPSAoR2V0LVdtaU9iamVjdCAtTmFtZXNwYWNl
IHJvb3QvY2ltdjIvbWRtL2RtbWFwIC1DbGFzcyBNRE1fRGV2RGV0YWlsX0V4dDAxIC1GaWx0ZXIgIkluc3RhbmNlSUQ9J0V4dCcgQU5EIFBhcmVudElEPScu
L0RldkRldGFpbCciKS5EZXZpY2VIYXJkd2FyZURhdGENCiAgICB9DQoNCiAgICAjIERpc3BsYXkgQXV0b3BpbG90IGRpYWcgZGV0YWlscw0KICAgIFdyaXRl
LUhvc3QgIiINCiAgICBXcml0ZS1Ib3N0ICJBVVRPUElMT1QgRElBR05PU1RJQ1MiIC1Gb3JlZ3JvdW5kQ29sb3IgTWFnZW50YQ0KICAgIFdyaXRlLUhvc3Qg
IiINCg0KICAgICMgRGV0ZXJtaW5lIHNjZW5hcmlvDQogICAgJHNjcmlwdDpBdXRvcGlsb3RTY2VuYXJpbyA9IFtBdXRvcGlsb3RTY2VuYXJpb0VudW1dOjpV
bmtub3duDQogICAgaWYgKFRlc3QtUGF0aCAiJGF1dG9waWxvdERpYWdQYXRoXEVzdGFibGlzaGVkQ29ycmVsYXRpb25zIiApDQogICAgew0KICAgICAgICAk
Y29ycmVsYXRpb25zID0gR2V0LUl0ZW1Qcm9wZXJ0eSAiJGF1dG9waWxvdERpYWdQYXRoXEVzdGFibGlzaGVkQ29ycmVsYXRpb25zIg0KICAgIH0NCiAgICBp
ZiAoVGVzdC1QYXRoICIkcHJvdmlzaW9uaW5nUGF0aFxBdXRvcGlsb3RTZXR0aW5ncyIgKQ0KICAgIHsNCiAgICAgICAgJHZhbHVlcyA9IEdldC1JdGVtUHJv
cGVydHkgIiRwcm92aXNpb25pbmdQYXRoXEF1dG9waWxvdFNldHRpbmdzIg0KICAgIH0NCiAgICBpZiAoJG51bGwgLW5lICR2YWx1ZXMuQXV0b3BpbG90RGV2
aWNlUHJlcEhpbnQpIHsNCiAgICAgICAgJHNjcmlwdDpBdXRvcGlsb3RTY2VuYXJpbyA9IFtBdXRvcGlsb3RTY2VuYXJpb0VudW1dOjpBdXRvcGlsb3RWMg0K
ICAgICAgICBXcml0ZS1Ib3N0ICJTY2VuYXJpbzogQXV0b3BpbG90IGRldmljZSBwcmVwYXJhdGlvbiAodjIpIg0KICAgICAgICAkc2V0dGluZ3MgPSBHZXQt
SXRlbVByb3BlcnR5IC1QYXRoICIkcHJvdmlzaW9uaW5nUGF0aFxBdXRvcGlsb3RTZXR0aW5nc1xEZXZpY2VQcmVwYXJhdGlvbiINCiAgICAgICAgJHBhZ2VT
ZXR0aW5ncyA9ICRzZXR0aW5ncy5QYWdlU2V0dGluZ3MgfCBDb252ZXJ0RnJvbS1Kc29uDQogICAgICAgICMgeyJBZ2VudERvd25sb2FkVGltZW91dFNlY29u
ZHMiOjE4MDAsIlBhZ2VUaW1lb3V0U2Vjb25kcyI6MzYwMCwiRXJyb3JNZXNzYWdlIjoiQ29udGFjdCB5b3VyIG9nYW5pemF0aW9uJ3Mgc3VwcG9ydCBwZXJz
b24gZm9yIGhlbHAuIiwiQWxsb3dTa2lwT25GYWlsdXJlIjp0cnVlLCJBbGxvd0RpYWdub3N0aWNzIjp0cnVlfQ0KICAgICAgICBXcml0ZS1Ib3N0ICJBZ2Vu
dERvd25sb2FkVGltZW91dFNlY29uZHM6ICQoJHBhZ2VTZXR0aW5ncy5BZ2VudERvd25sb2FkVGltZW91dFNlY29uZHMpIg0KICAgICAgICBXcml0ZS1Ib3N0
ICJQYWdlVGltZW91dFNlY29uZHM6ICQoJHBhZ2VTZXR0aW5ncy5QYWdlVGltZW91dFNlY29uZHMpIg0KICAgICAgICBXcml0ZS1Ib3N0ICJBbGxvd1NraXBP
bkZhaWx1cmU6ICQoJHBhZ2VTZXR0aW5ncy5BbGxvd1NraXBPbkZhaWx1cmUpIg0KICAgICAgICBXcml0ZS1Ib3N0ICJBbGxvd0RpYWdub3N0aWNzOiAkKCRw
YWdlU2V0dGluZ3MuQWxsb3dEaWFnbm9zdGljcykiDQoNCiAgICAgICAgR2V0LUNoaWxkSXRlbSAkZW5yb2xsbWVudHNQYXRoIHwgRm9yRWFjaC1PYmplY3Qg
ew0KICAgICAgICAgICAgJHByb3BlcnRpZXMgPSBHZXQtSXRlbVByb3BlcnR5IC1QYXRoICRfLlBTUGF0aA0KICAgICAgICAgICAgaWYgKCRwcm9wZXJ0aWVz
LlByb3ZpZGVySWQgLWVxICJNUyBETSBTZXJ2ZXIiKSB7DQogICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiVGVuYW50SUQ6ICQoJHByb3BlcnRpZXMuQUFE
VGVuYW50SUQpIg0KICAgICAgICAgICAgICAgIFdyaXRlLUhvc3QgIlVQTjogJCgkcHJvcGVydGllcy5VUE4pIg0KICAgICAgICAgICAgfQ0KICAgICAgICB9
DQoNCiAgICB9IGVsc2Ugew0KICAgICAgICAkdmFsdWVzID0gR2V0LUl0ZW1Qcm9wZXJ0eSAiJGF1dG9waWxvdERpYWdQYXRoIg0KICAgICAgICBpZiAoJHZh
bHVlcy5DbG91ZEFzc2lnbmVkVGVuYW50SWQpIHsNCiAgICAgICAgICAgIGlmICgkdmFsdWVzLkRlcGxveW1lbnRQcm9maWxlTmFtZSAtYW5kICR2YWx1ZXMu
RGVwbG95bWVudFByb2ZpbGVOYW1lIC1uZSAiIikgew0KICAgICAgICAgICAgICAgICRzY3JpcHQ6QXV0b3BpbG90U2NlbmFyaW8gPSBbQXV0b3BpbG90U2Nl
bmFyaW9FbnVtXTo6QXV0b3BpbG90VjENCiAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICJTY2VuYXJpbzogQXV0b3BpbG90ICh2MSkiDQogICAgICAgICAg
ICAgICAgV3JpdGUtSG9zdCAiUHJvZmlsZTogJCgkdmFsdWVzLkRlcGxveW1lbnRQcm9maWxlTmFtZSkiDQogICAgICAgICAgICB9IGVsc2Ugew0KICAgICAg
ICAgICAgICAgICRzY3JpcHQ6QXV0b3BpbG90U2NlbmFyaW8gPSBbQXV0b3BpbG90U2NlbmFyaW9FbnVtXTo6QXV0b3BpbG90SnNvbg0KICAgICAgICAgICAg
ICAgIFdyaXRlLUhvc3QgIlNjZW5hcmlvOiBBdXRvcGlsb3QgZm9yIGV4aXN0aW5nIGRldmljZXMgKHYxKSINCiAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0
ICJDb3JyZWxhdGlvbiBJRDogJCgkdmFsdWVzLlp0ZENvcnJlbGF0aW9uSWQpIiAgICAgICAgICAgICAgICANCiAgICAgICAgICAgIH0NCiAgICAgICAgICAg
IFdyaXRlLUhvc3QgIlRlbmFudERvbWFpbjogJCgkdmFsdWVzLkNsb3VkQXNzaWduZWRUZW5hbnREb21haW4pIg0KICAgICAgICAgICAgV3JpdGUtSG9zdCAi
VGVuYW50SUQ6ICQoJHZhbHVlcy5DbG91ZEFzc2lnbmVkVGVuYW50SWQpIg0KICAgICAgICAgICAgV3JpdGUtSG9zdCAiT29iZUNvbmZpZzogJCgkdmFsdWVz
LkNsb3VkQXNzaWduZWRPb2JlQ29uZmlnKSINCg0KICAgICAgICAgICAgaWYgKCgkdmFsdWVzLkNsb3VkQXNzaWduZWRPb2JlQ29uZmlnIC1iYW5kIDEwMjQp
IC1ndCAwKSB7DQogICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiIFNraXAga2V5Ym9hcmQ6IFllcyAgICAgICAgICAgICAxIC0gLSAtIC0gLSAtIC0gLSAt
IC0iDQogICAgICAgICAgICB9DQogICAgICAgICAgICBlbHNlIHsNCiAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICIgU2tpcCBrZXlib2FyZDogTm8gICAg
ICAgICAgICAgIDAgLSAtIC0gLSAtIC0gLSAtIC0gLSINCiAgICAgICAgICAgIH0NCiAgICAgICAgICAgIGlmICgoJHZhbHVlcy5DbG91ZEFzc2lnbmVkT29i
ZUNvbmZpZyAtYmFuZCA1MTIpIC1ndCAwKSB7DQogICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiIEVuYWJsZSBwYXRjaCBkb3dubG9hZDogWWVzICAgICAt
IDEgLSAtIC0gLSAtIC0gLSAtIC0iDQogICAgICAgICAgICB9DQogICAgICAgICAgICBlbHNlIHsNCiAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICIgRW5h
YmxlIHBhdGNoIGRvd25sb2FkOiBObyAgICAgIC0gMCAtIC0gLSAtIC0gLSAtIC0gLSINCiAgICAgICAgICAgIH0NCiAgICAgICAgICAgIGlmICgoJHZhbHVl
cy5DbG91ZEFzc2lnbmVkT29iZUNvbmZpZyAtYmFuZCAyNTYpIC1ndCAwKSB7DQogICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiIFNraXAgV2luZG93cyB1
cGdyYWRlIFVYOiBZZXMgICAtIC0gMSAtIC0gLSAtIC0gLSAtIC0iDQogICAgICAgICAgICB9DQogICAgICAgICAgICBlbHNlIHsNCiAgICAgICAgICAgICAg
ICBXcml0ZS1Ib3N0ICIgU2tpcCBXaW5kb3dzIHVwZ3JhZGUgVVg6IE5vICAgIC0gLSAwIC0gLSAtIC0gLSAtIC0gLSINCiAgICAgICAgICAgIH0NCiAgICAg
ICAgICAgIGlmICgoJHZhbHVlcy5DbG91ZEFzc2lnbmVkT29iZUNvbmZpZyAtYmFuZCAxMjgpIC1ndCAwKSB7DQogICAgICAgICAgICAgICAgV3JpdGUtSG9z
dCAiIEFBRCBUUE0gUmVxdWlyZWQ6IFllcyAgICAgICAgICAtIC0gLSAxIC0gLSAtIC0gLSAtIC0iDQogICAgICAgICAgICB9DQogICAgICAgICAgICBlbHNl
IHsNCiAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICIgQUFEIFRQTSBSZXF1aXJlZDogTm8gICAgICAgICAgIC0gLSAtIDAgLSAtIC0gLSAtIC0gLSINCiAg
ICAgICAgICAgIH0NCiAgICAgICAgICAgIGlmICgoJHZhbHVlcy5DbG91ZEFzc2lnbmVkT29iZUNvbmZpZyAtYmFuZCA2NCkgLWd0IDApIHsNCiAgICAgICAg
ICAgICAgICBXcml0ZS1Ib3N0ICIgQUFEIGRldmljZSBhdXRoOiBZZXMgICAgICAgICAgIC0gLSAtIC0gMSAtIC0gLSAtIC0gLSINCiAgICAgICAgICAgIH0N
CiAgICAgICAgICAgIGVsc2Ugew0KICAgICAgICAgICAgICAgIFdyaXRlLUhvc3QgIiBBQUQgZGV2aWNlIGF1dGg6IE5vICAgICAgICAgICAgLSAtIC0gLSAw
IC0gLSAtIC0gLSAtIg0KICAgICAgICAgICAgfQ0KICAgICAgICAgICAgaWYgKCgkdmFsdWVzLkNsb3VkQXNzaWduZWRPb2JlQ29uZmlnIC1iYW5kIDMyKSAt
Z3QgMCkgew0KICAgICAgICAgICAgICAgIFdyaXRlLUhvc3QgIiBUUE0gYXR0ZXN0YXRpb246IFllcyAgICAgICAgICAgLSAtIC0gLSAtIDEgLSAtIC0gLSAt
Ig0KICAgICAgICAgICAgfQ0KICAgICAgICAgICAgZWxzZSB7DQogICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiIFRQTSBhdHRlc3RhdGlvbjogTm8gICAg
ICAgICAgICAtIC0gLSAtIC0gMCAtIC0gLSAtIC0iDQogICAgICAgICAgICB9DQogICAgICAgICAgICBpZiAoKCR2YWx1ZXMuQ2xvdWRBc3NpZ25lZE9vYmVD
b25maWcgLWJhbmQgMTYpIC1ndCAwKSB7DQogICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiIFNraXAgRVVMQTogWWVzICAgICAgICAgICAgICAgICAtIC0g
LSAtIC0gLSAxIC0gLSAtIC0iDQogICAgICAgICAgICB9DQogICAgICAgICAgICBlbHNlIHsNCiAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICIgU2tpcCBF
VUxBOiBObyAgICAgICAgICAgICAgICAgIC0gLSAtIC0gLSAtIDAgLSAtIC0gLSINCiAgICAgICAgICAgIH0NCiAgICAgICAgICAgIGlmICgoJHZhbHVlcy5D
bG91ZEFzc2lnbmVkT29iZUNvbmZpZyAtYmFuZCA4KSAtZ3QgMCkgew0KICAgICAgICAgICAgICAgIFdyaXRlLUhvc3QgIiBTa2lwIE9FTSByZWdpc3RyYXRp
b246IFllcyAgICAgLSAtIC0gLSAtIC0gLSAxIC0gLSAtIg0KICAgICAgICAgICAgfQ0KICAgICAgICAgICAgZWxzZSB7DQogICAgICAgICAgICAgICAgV3Jp
dGUtSG9zdCAiIFNraXAgT0VNIHJlZ2lzdHJhdGlvbjogTm8gICAgICAtIC0gLSAtIC0gLSAtIDAgLSAtIC0iDQogICAgICAgICAgICB9DQogICAgICAgICAg
ICBpZiAoKCR2YWx1ZXMuQ2xvdWRBc3NpZ25lZE9vYmVDb25maWcgLWJhbmQgNCkgLWd0IDApIHsNCiAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICIgU2tp
cCBleHByZXNzIHNldHRpbmdzOiBZZXMgICAgIC0gLSAtIC0gLSAtIC0gLSAxIC0gLSINCiAgICAgICAgICAgIH0NCiAgICAgICAgICAgIGVsc2Ugew0KICAg
ICAgICAgICAgICAgIFdyaXRlLUhvc3QgIiBTa2lwIGV4cHJlc3Mgc2V0dGluZ3M6IE5vICAgICAgLSAtIC0gLSAtIC0gLSAtIDAgLSAtIg0KICAgICAgICAg
ICAgfQ0KICAgICAgICAgICAgaWYgKCgkdmFsdWVzLkNsb3VkQXNzaWduZWRPb2JlQ29uZmlnIC1iYW5kIDIpIC1ndCAwKSB7DQogICAgICAgICAgICAgICAg
V3JpdGUtSG9zdCAiIERpc2FsbG93IGFkbWluOiBZZXMgICAgICAgICAgICAtIC0gLSAtIC0gLSAtIC0gLSAxIC0iDQogICAgICAgICAgICB9DQogICAgICAg
ICAgICBlbHNlIHsNCiAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICIgRGlzYWxsb3cgYWRtaW46IE5vICAgICAgICAgICAgIC0gLSAtIC0gLSAtIC0gLSAt
IDAgLSINCiAgICAgICAgICAgIH0NCg0KICAgICAgICAgICAgIyBJbiB0aGVvcnkgd2UgY291bGQgcmVhZCB0aGVzZSB2YWx1ZXMgZnJvbSB0aGUgcHJvZmls
ZSBjYWNoZSByZWdpc3RyeSBrZXksIGJ1dCBpdCdzIHNvIGJ1bmdsZWQNCiAgICAgICAgICAgICMgdXAgaW4gdGhlIHJlZ2lzdHJ5IGV4cG9ydCB0aGF0IGl0
IGRvZXNuJ3QgaW1wb3J0IHdpdGhvdXQgc29tZSBzZXJpb3VzIG1hc3NhZ2luZyBmb3IgZW1iZWRkZWQNCiAgICAgICAgICAgICMgcXVvdGVzLiBTbyB0aGlz
IGlzIGVhc2llci4NCiAgICAgICAgICAgIGlmICgkc2NyaXB0OnVzZUZpbGUpIHsNCiAgICAgICAgICAgICAgICAkanNvbkZpbGUgPSAiJCgkZW52OlRFTVAp
XEVTUFN0YXR1cy50bXBcQXV0b3BpbG90RERTWlRERmlsZS5qc29uIg0KICAgICAgICAgICAgfQ0KICAgICAgICAgICAgZWxzZSB7DQogICAgICAgICAgICAg
ICAgJGpzb25GaWxlID0gIiQoJGVudjpXSU5ESVIpXFNlcnZpY2VTdGF0ZVx3bWFuc3ZjXEF1dG9waWxvdEREU1pUREZpbGUuanNvbiIgDQogICAgICAgICAg
ICB9DQogICAgICAgICAgICBpZiAoVGVzdC1QYXRoICRqc29uRmlsZSkgew0KICAgICAgICAgICAgICAgICRqc29uID0gR2V0LUNvbnRlbnQgJGpzb25GaWxl
IHwgQ29udmVydEZyb20tSnNvbg0KICAgICAgICAgICAgICAgICRkYXRlID0gW2RhdGV0aW1lXSRqc29uLlBvbGljeURvd25sb2FkRGF0ZQ0KICAgICAgICAg
ICAgICAgIFJlY29yZFN0YXR1cyAtZGF0ZSAkZGF0ZSAtZGV0YWlsICJBdXRvcGlsb3QgcHJvZmlsZSIgLXN0YXR1cyAiUHJvZmlsZSBkb3dubG9hZGVkIiAt
Y29sb3IgIlllbGxvdyIgDQogICAgICAgICAgICAgICAgaWYgKCRqc29uLkNsb3VkQXNzaWduZWREb21haW5Kb2luTWV0aG9kIC1lcSAxKSB7DQogICAgICAg
ICAgICAgICAgICAgIFdyaXRlLUhvc3QgIlN1YnNjZW5hcmlvczogSHlicmlkIEF6dXJlIEFEIEpvaW4iDQogICAgICAgICAgICAgICAgICAgIGlmIChUZXN0
LVBhdGggIiRvbWFkbVBhdGhcU3luY01MXE9ESkFwcGxpZWQiKSB7DQogICAgICAgICAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICJPREogYXBwbGllZDog
WWVzIg0KICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgICAgIGVsc2Ugew0KICAgICAgICAgICAgICAgICAgICAgICAgV3JpdGUtSG9z
dCAiT0RKIGFwcGxpZWQ6IE5vIiAgICAgICAgICAgICAgICANCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgICAgICBpZiAoJGpzb24u
SHlicmlkSm9pblNraXBEQ0Nvbm5lY3Rpdml0eUNoZWNrIC1lcSAxKSB7DQogICAgICAgICAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICJTa2lwIGNvbm5l
Y3Rpdml0eSBjaGVjazogWWVzIg0KICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgICAgIGVsc2Ugew0KICAgICAgICAgICAgICAgICAg
ICAgICAgV3JpdGUtSG9zdCAiU2tpcCBjb25uZWN0aXZpdHkgY2hlY2s6IE5vIg0KICAgICAgICAgICAgICAgICAgICB9DQoNCiAgICAgICAgICAgICAgICB9
DQogICAgICAgICAgICAgICAgZWxzZSB7DQogICAgICAgICAgICAgICAgICAgIFdyaXRlLUhvc3QgIlN1YnNjZW5hcmlvOiBBenVyZSBBRCBKb2luIg0KICAg
ICAgICAgICAgICAgIH0NCiAgICAgICAgICAgIH0NCiAgICAgICAgICAgIGVsc2Ugew0KICAgICAgICAgICAgICAgIFdyaXRlLUhvc3QgIlN1YnNjZW5hcmlv
OiBOb3QgYXZhaWxhYmxlIChKU09OIG5vdCBmb3VuZCkiDQogICAgICAgICAgICB9DQoNCiAgICAgICAgfQ0KICAgIH0NCg0KICAgIGlmICgtbm90ICRzY3Jp
cHQ6dXNlRmlsZSkgew0KICAgICAgICAkb3NWZXJzaW9uID0gKEdldC1XbWlPYmplY3Qgd2luMzJfb3BlcmF0aW5nc3lzdGVtKS5WZXJzaW9uDQogICAgICAg
IFdyaXRlLUhvc3QgIk9TIHZlcnNpb246ICRvc1ZlcnNpb24iDQogICAgfQ0KICAgIFdyaXRlLUhvc3QgIkVudERNSUQ6ICQoJGNvcnJlbGF0aW9ucy5FbnRE
TUlEKSINCg0KICAgICMgR2V0IEVTUCBwcm9wZXJ0aWVzDQogICAgR2V0LUNoaWxkSXRlbSAkZW5yb2xsbWVudHNQYXRoIHwgV2hlcmUtT2JqZWN0IHsgVGVz
dC1QYXRoICIkKCRfLlBTUGF0aClcRmlyc3RTeW5jIiB9IHwgJSB7DQogICAgICAgIGlmICgkc2NyaXB0OkF1dG9waWxvdFNjZW5hcmlvIC1lcSBbQXV0b3Bp
bG90U2NlbmFyaW9FbnVtXTo6VW5rbm93bikgew0KICAgICAgICAgICAgJHNjcmlwdDpBdXRvcGlsb3RTY2VuYXJpbyA9IFtBdXRvcGlsb3RTY2VuYXJpb0Vu
dW1dOjpFc3BPbmx5DQogICAgICAgIH0NCiAgICAgICAgJHByb3BlcnRpZXMgPSBHZXQtSXRlbVByb3BlcnR5ICIkKCRfLlBTUGF0aClcRmlyc3RTeW5jIg0K
ICAgICAgICBXcml0ZS1Ib3N0ICJFbnJvbGxtZW50IHN0YXR1cyBwYWdlOiINCiAgICAgICAgV3JpdGUtSG9zdCAiIERldmljZSBFU1AgZW5hYmxlZDogJCgk
cHJvcGVydGllcy5Ta2lwRGV2aWNlU3RhdHVzUGFnZSAtZXEgMCkiDQogICAgICAgIFdyaXRlLUhvc3QgIiBVc2VyIEVTUCBlbmFibGVkOiAkKCRwcm9wZXJ0
aWVzLlNraXBVc2VyU3RhdHVzUGFnZSAtZXEgMCkiDQogICAgICAgIFdyaXRlLUhvc3QgIiBFU1AgdGltZW91dDogJCgkcHJvcGVydGllcy5TeW5jRmFpbHVy
ZVRpbWVvdXQpIg0KICAgICAgICBpZiAoJHByb3BlcnRpZXMuQmxvY2tJblN0YXR1c1BhZ2UgLWVxIDApIHsNCiAgICAgICAgICAgIFdyaXRlLUhvc3QgIiBF
U1AgYmxvY2tpbmc6IE5vIg0KICAgICAgICB9DQogICAgICAgIGVsc2Ugew0KICAgICAgICAgICAgV3JpdGUtSG9zdCAiIEVTUCBibG9ja2luZzogWWVzIg0K
ICAgICAgICAgICAgaWYgKCRwcm9wZXJ0aWVzLkJsb2NrSW5TdGF0dXNQYWdlIC1iYW5kIDEpIHsNCiAgICAgICAgICAgICAgICBXcml0ZS1Ib3N0ICIgRVNQ
IGFsbG93IHJlc2V0OiBZZXMiDQogICAgICAgICAgICB9DQogICAgICAgICAgICBpZiAoJHByb3BlcnRpZXMuQmxvY2tJblN0YXR1c1BhZ2UgLWJhbmQgMikg
ew0KICAgICAgICAgICAgICAgIFdyaXRlLUhvc3QgIiBFU1AgYWxsb3cgdHJ5IGFnYWluOiBZZXMiDQogICAgICAgICAgICB9DQogICAgICAgICAgICBpZiAo
JHByb3BlcnRpZXMuQmxvY2tJblN0YXR1c1BhZ2UgLWJhbmQgNCkgew0KICAgICAgICAgICAgICAgIFdyaXRlLUhvc3QgIiBFU1AgY29udGludWUgYW55d2F5
OiBZZXMiDQogICAgICAgICAgICB9DQogICAgICAgIH0NCiAgICB9DQoNCiAgICAjIEdldCBEZWxpdmVyeSBPcHRpbWl6YXRpb24gc3RhdGlzdGljcyAod2hl
biBhdmFpbGFibGUpDQogICAgaWYgKC1ub3QgJHNjcmlwdDp1c2VGaWxlKSB7DQogICAgICAgICRzdGF0cyA9IEdldC1EZWxpdmVyeU9wdGltaXphdGlvblBl
cmZTbmFwVGhpc01vbnRoDQogICAgICAgIGlmICgkc3RhdHMuRG93bmxvYWRIdHRwQnl0ZXMgLW5lIDApIHsNCiAgICAgICAgICAgICRwZWVyUGN0ID0gW21h
dGhdOjpSb3VuZCggKCRzdGF0cy5Eb3dubG9hZExhbkJ5dGVzIC8gJHN0YXRzLkRvd25sb2FkSHR0cEJ5dGVzKSAqIDEwMCApDQogICAgICAgICAgICAkY2NQ
Y3QgPSBbbWF0aF06OlJvdW5kKCAoJHN0YXRzLkRvd25sb2FkQ2FjaGVIb3N0Qnl0ZXMgLyAkc3RhdHMuRG93bmxvYWRIdHRwQnl0ZXMpICogMTAwICkNCiAg
ICAgICAgfQ0KICAgICAgICBlbHNlIHsNCiAgICAgICAgICAgICRwZWVyUGN0ID0gMA0KICAgICAgICAgICAgJGNjUGN0ID0gMA0KICAgICAgICB9DQogICAg
ICAgIFdyaXRlLUhvc3QgIkRlbGl2ZXJ5IE9wdGltaXphdGlvbiBzdGF0aXN0aWNzOiINCiAgICAgICAgV3JpdGUtSG9zdCAiIFRvdGFsIGJ5dGVzIGRvd25s
b2FkZWQ6ICQoJHN0YXRzLkRvd25sb2FkSHR0cEJ5dGVzKSINCiAgICAgICAgV3JpdGUtSG9zdCAiIEZyb20gcGVlcnM6ICQoJHBlZXJQY3QpJSAoJCgkc3Rh
dHMuRG93bmxvYWRMYW5CeXRlcykpIg0KICAgICAgICBXcml0ZS1Ib3N0ICIgRnJvbSBDb25uZWN0ZWQgQ2FjaGU6ICQoJGNjUGN0KSUgKCQoJHN0YXRzLkRv
d25sb2FkQ2FjaGVIb3N0Qnl0ZXMpKSINCiAgICB9DQoNCiAgICAjIElmIHRoZSBBREsgaXMgaW5zdGFsbGVkLCBnZXQgc29tZSBrZXkgaGFyZHdhcmUgaGFz
aCBpbmZvDQogICAgJGFka1BhdGggPSBHZXQtSXRlbVByb3BlcnR5VmFsdWUgIkhLTE06XFNvZnR3YXJlXE1pY3Jvc29mdFxXaW5kb3dzIEtpdHNcSW5zdGFs
bGVkIFJvb3RzIiAtTmFtZSBLaXRzUm9vdDEwIC1FcnJvckFjdGlvbiBTaWxlbnRseUNvbnRpbnVlDQogICAgJG9hM1Rvb2wgPSAiJGFka1BhdGhcQXNzZXNz
bWVudCBhbmQgRGVwbG95bWVudCBLaXRcRGVwbG95bWVudCBUb29sc1wkKCRlbnY6UFJPQ0VTU09SX0FSQ0hJVEVDVFVSRSlcTGljZW5zaW5nXE9BMzBcb2Ez
dG9vbC5leGUiDQogICAgaWYgKCRoYXNoIC1hbmQgKFRlc3QtUGF0aCAkb2EzVG9vbCkpIHsNCiAgICAgICAgJGNvbW1hbmRMaW5lQXJncyA9ICIvZGVjb2Rl
aHdoYXNoOiRoYXNoIg0KICAgICAgICAkb3V0cHV0ID0gJiAiJG9hM1Rvb2wiICRjb21tYW5kTGluZUFyZ3MNCiAgICAgICAgW3htbF0gJGhhc2hYTUwgPSAk
b3V0cHV0IHwgU2VsZWN0IC1za2lwIDggLUZpcnN0ICgkb3V0cHV0LkNvdW50IC0gMTIpDQogICAgICAgIFdyaXRlLUhvc3QgIkhhcmR3YXJlIGluZm9ybWF0
aW9uOiINCiAgICAgICAgV3JpdGUtSG9zdCAiIE9wZXJhdGluZyBzeXN0ZW0gYnVpbGQ6ICIgJGhhc2hYTUwuU2VsZWN0U2luZ2xlTm9kZSgiLy9wW0BuPSdP
c0J1aWxkJ10iKS52DQogICAgICAgIFdyaXRlLUhvc3QgIiBNYW51ZmFjdHVyZXI6ICIgJGhhc2hYTUwuU2VsZWN0U2luZ2xlTm9kZSgiLy9wW0BuPSdTbWJp
b3NTeXN0ZW1NYW51ZmFjdHVyZXInXSIpLnYNCiAgICAgICAgV3JpdGUtSG9zdCAiIE1vZGVsOiAiICRoYXNoWE1MLlNlbGVjdFNpbmdsZU5vZGUoIi8vcFtA
bj0nU21iaW9zU3lzdGVtUHJvZHVjdE5hbWUnXSIpLnYNCiAgICAgICAgV3JpdGUtSG9zdCAiIFNlcmlhbCBudW1iZXI6ICIgJGhhc2hYTUwuU2VsZWN0U2lu
Z2xlTm9kZSgiLy9wW0BuPSdTbWJpb3NTeXN0ZW1TZXJpYWxOdW1iZXInXSIpLnYNCiAgICAgICAgV3JpdGUtSG9zdCAiIFRQTSB2ZXJzaW9uOiAiICRoYXNo
WE1MLlNlbGVjdFNpbmdsZU5vZGUoIi8vcFtAbj0nVFBNVmVyc2lvbiddIikudg0KICAgIH0NCiAgICANCiAgICAjIFByb2Nlc3MgZXZlbnQgbG9nIGluZm8N
CiAgICBQcm9jZXNzRXZlbnRzDQoNCiAgICAjIERpc3BsYXkgdGhlIGxpc3Qgb2YgcG9saWNpZXMNCiAgICBpZiAoJFNob3dQb2xpY2llcykgew0KICAgICAg
ICBXcml0ZS1Ib3N0ICIgIg0KICAgICAgICBXcml0ZS1Ib3N0ICJQT0xJQ0lFUyBQUk9DRVNTRUQiIC1Gb3JlZ3JvdW5kQ29sb3IgTWFnZW50YSAgIA0KICAg
ICAgICBQcm9jZXNzTm9kZUNhY2hlIHwgRm9ybWF0LVRhYmxlIC1XcmFwDQogICAgfQ0KICAgIA0KICAgIGlmICgkc2NyaXB0OkF1dG9waWxvdFNjZW5hcmlv
IC1lcSBbQXV0b3BpbG90U2NlbmFyaW9FbnVtXTo6QXV0b3BpbG90VjIpIHsNCg0KICAgICAgICAjIFByb2Nlc3Mgc2NyaXB0cw0KICAgICAgICBXcml0ZS1I
b3N0ICIgIg0KICAgICAgICBXcml0ZS1Ib3N0ICJTQ1JJUFRTOiIgLUZvcmVncm91bmRDb2xvciBNYWdlbnRhDQogICAgICAgIFdyaXRlLUhvc3QgIiAiDQog
ICAgICAgIFByb2Nlc3NTaWRlY2FyVjJTY3JpcHRzDQoNCiAgICAgICAgIyBQcm9jZXNzIFdpbjMyIGFwcHMNCiAgICAgICAgV3JpdGUtSG9zdCAiICINCiAg
ICAgICAgV3JpdGUtSG9zdCAiQVBQUzoiIC1Gb3JlZ3JvdW5kQ29sb3IgTWFnZW50YQ0KICAgICAgICBXcml0ZS1Ib3N0ICIgIg0KICAgICAgICBQcm9jZXNz
U2lkZWNhclYyDQoNCiAgICB9IGVsc2Ugew0KICAgICAgICAjIE1ha2Ugc3VyZSB0aGUgdHJhY2tpbmcgcGF0aCBleGlzdHMNCiAgICAgICAgaWYgKFRlc3Qt
UGF0aCAkcGF0aCkgew0KDQogICAgICAgICAgICAjIFByb2Nlc3MgZGV2aWNlIEVTUCBzZXNzaW9ucw0KICAgICAgICAgICAgV3JpdGUtSG9zdCAiICINCiAg
ICAgICAgICAgIFdyaXRlLUhvc3QgIkRFVklDRSBFU1A6IiAtRm9yZWdyb3VuZENvbG9yIE1hZ2VudGENCiAgICAgICAgICAgIFdyaXRlLUhvc3QgIiAiDQoN
CiAgICAgICAgICAgIGlmIChUZXN0LVBhdGggIiRwYXRoXEV4cGVjdGVkUG9saWNpZXMiKSB7DQogICAgICAgICAgICAgICAgW2FycmF5XSRpdGVtcyA9IEdl
dC1DaGlsZEl0ZW0gIiRwYXRoXEV4cGVjdGVkUG9saWNpZXMiDQogICAgICAgICAgICAgICAgaWYgKCRpdGVtcykgew0KICAgICAgICAgICAgICAgICAgICBB
ZGREaXNwbGF5IChbcmVmXSRpdGVtcykNCiAgICAgICAgICAgICAgICAgICAgJGl0ZW1zIHwgUHJvY2Vzc1BvbGljaWVzDQogICAgICAgICAgICAgICAgfQ0K
ICAgICAgICAgICAgfQ0KICAgICAgICAgICAgaWYgKFRlc3QtUGF0aCAiJHBhdGhcRXhwZWN0ZWRNU0lBcHBQYWNrYWdlcyIpIHsNCiAgICAgICAgICAgICAg
ICBbYXJyYXldJGl0ZW1zID0gR2V0LUNoaWxkSXRlbSAiJHBhdGhcRXhwZWN0ZWRNU0lBcHBQYWNrYWdlcyINCiAgICAgICAgICAgICAgICBpZiAoJGl0ZW1z
KSB7DQogICAgICAgICAgICAgICAgICAgIEFkZERpc3BsYXkgKFtyZWZdJGl0ZW1zKQ0KICAgICAgICAgICAgICAgICAgICAkaXRlbXMgfCBQcm9jZXNzQXBw
cyAtY3VycmVudFVzZXIgIlMtMC0wLTAwLTAwMDAwMDAwMDAtMDAwMDAwMDAwMC0wMDAwMDAwMDAtMDAwIiANCiAgICAgICAgICAgICAgICB9DQogICAgICAg
ICAgICB9DQogICAgICAgICAgICBpZiAoVGVzdC1QYXRoICIkcGF0aFxFeHBlY3RlZE1vZGVybkFwcFBhY2thZ2VzIikgew0KICAgICAgICAgICAgICAgIFth
cnJheV0kaXRlbXMgPSBHZXQtQ2hpbGRJdGVtICIkcGF0aFxFeHBlY3RlZE1vZGVybkFwcFBhY2thZ2VzIg0KICAgICAgICAgICAgICAgIGlmICgkaXRlbXMp
IHsNCiAgICAgICAgICAgICAgICAgICAgQWRkRGlzcGxheSAoW3JlZl0kaXRlbXMpDQogICAgICAgICAgICAgICAgICAgICRpdGVtcyB8IFByb2Nlc3NNb2Rl
cm5BcHBzIC1jdXJyZW50VXNlciAiUy0wLTAtMDAtMDAwMDAwMDAwMC0wMDAwMDAwMDAwLTAwMDAwMDAwMC0wMDAiDQogICAgICAgICAgICAgICAgfQ0KICAg
ICAgICAgICAgfQ0KICAgICAgICAgICAgaWYgKFRlc3QtUGF0aCAiJHBhdGhcU2lkZWNhciIpIHsNCiAgICAgICAgICAgICAgICBbYXJyYXldJGl0ZW1zID0g
R2V0LUNoaWxkSXRlbSAiJHBhdGhcU2lkZWNhciIgfCA/IHsgJF8uUHJvcGVydHkgLW1hdGNoICIuL0RldmljZSIgLWFuZCAkXy5OYW1lIC1ub3RtYXRjaCAi
TGFzdExvZ2dlZFN0YXRlIiB9DQogICAgICAgICAgICAgICAgaWYgKCRpdGVtcykgew0KICAgICAgICAgICAgICAgICAgICBBZGREaXNwbGF5IChbcmVmXSRp
dGVtcykNCiAgICAgICAgICAgICAgICAgICAgJGl0ZW1zIHwgUHJvY2Vzc1NpZGVjYXIgLWN1cnJlbnRVc2VyICIwMDAwMDAwMC0wMDAwLTAwMDAtMDAwMC0w
MDAwMDAwMDAwMDAiDQogICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgfQ0KICAgICAgICAgICAgaWYgKFRlc3QtUGF0aCAiJHBhdGhcRXhwZWN0ZWRT
Q0VQQ2VydHMiKSB7DQogICAgICAgICAgICAgICAgW2FycmF5XSRpdGVtcyA9IEdldC1DaGlsZEl0ZW0gIiRwYXRoXEV4cGVjdGVkU0NFUENlcnRzIg0KICAg
ICAgICAgICAgICAgIGlmICgkaXRlbXMpIHsNCiAgICAgICAgICAgICAgICAgICAgQWRkRGlzcGxheSAoW3JlZl0kaXRlbXMpDQogICAgICAgICAgICAgICAg
ICAgICRpdGVtcyB8IFByb2Nlc3NDZXJ0cw0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgIH0NCg0KICAgICAgICAgICAgIyBQcm9jZXNzIHVzZXIg
RVNQIHNlc3Npb25zDQogICAgICAgICAgICBHZXQtQ2hpbGRJdGVtICIkcGF0aCIgfCA/IHsgJF8uUFNDaGlsZE5hbWUuU3RhcnRzV2l0aCgiUy0iKSB9IHwg
JSB7DQogICAgICAgICAgICAgICAgJHVzZXJQYXRoID0gJF8uUFNQYXRoDQogICAgICAgICAgICAgICAgJHVzZXJTaWQgPSAkXy5QU0NoaWxkTmFtZQ0KICAg
ICAgICAgICAgICAgIFdyaXRlLUhvc3QgIiAiDQogICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiVVNFUiBFU1AgZm9yICQoJHVzZXJTaWQpOiIgLUZvcmVn
cm91bmRDb2xvciBNYWdlbnRhDQogICAgICAgICAgICAgICAgV3JpdGUtSG9zdCAiICINCiAgICAgICAgICAgICAgICBpZiAoVGVzdC1QYXRoICIkdXNlclBh
dGhcRXhwZWN0ZWRQb2xpY2llcyIpIHsNCiAgICAgICAgICAgICAgICAgICAgW2FycmF5XSRpdGVtcyA9IEdldC1DaGlsZEl0ZW0gIiR1c2VyUGF0aFxFeHBl
Y3RlZFBvbGljaWVzIg0KICAgICAgICAgICAgICAgICAgICBpZiAoJGl0ZW1zKSB7DQogICAgICAgICAgICAgICAgICAgICAgICBBZGREaXNwbGF5IChbcmVm
XSRpdGVtcykNCiAgICAgICAgICAgICAgICAgICAgICAgICRpdGVtcyB8IFByb2Nlc3NQb2xpY2llcw0KICAgICAgICAgICAgICAgICAgICB9DQogICAgICAg
ICAgICAgICAgfQ0KICAgICAgICAgICAgICAgIGlmIChUZXN0LVBhdGggIiR1c2VyUGF0aFxFeHBlY3RlZE1TSUFwcFBhY2thZ2VzIikgew0KICAgICAgICAg
ICAgICAgICAgICBbYXJyYXldJGl0ZW1zID0gR2V0LUNoaWxkSXRlbSAiJHVzZXJQYXRoXEV4cGVjdGVkTVNJQXBwUGFja2FnZXMiIA0KICAgICAgICAgICAg
ICAgICAgICBpZiAoJGl0ZW1zKSB7DQogICAgICAgICAgICAgICAgICAgICAgICBBZGREaXNwbGF5IChbcmVmXSRpdGVtcykNCiAgICAgICAgICAgICAgICAg
ICAgICAgICRpdGVtcyB8IFByb2Nlc3NBcHBzIC1jdXJyZW50VXNlciAkdXNlclNpZA0KICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAg
fQ0KICAgICAgICAgICAgICAgIGlmIChUZXN0LVBhdGggIiR1c2VyUGF0aFxFeHBlY3RlZE1vZGVybkFwcFBhY2thZ2VzIikgew0KICAgICAgICAgICAgICAg
ICAgICBbYXJyYXldJGl0ZW1zID0gR2V0LUNoaWxkSXRlbSAiJHVzZXJQYXRoXEV4cGVjdGVkTW9kZXJuQXBwUGFja2FnZXMiDQogICAgICAgICAgICAgICAg
ICAgIGlmICgkaXRlbXMpIHsNCiAgICAgICAgICAgICAgICAgICAgICAgIEFkZERpc3BsYXkgKFtyZWZdJGl0ZW1zKQ0KICAgICAgICAgICAgICAgICAgICAg
ICAgJGl0ZW1zIHwgUHJvY2Vzc01vZGVybkFwcHMgLWN1cnJlbnRVc2VyICR1c2VyU2lkDQogICAgICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgICAg
ICB9DQogICAgICAgICAgICAgICAgaWYgKFRlc3QtUGF0aCAiJHVzZXJQYXRoXFNpZGVjYXIiKSB7DQogICAgICAgICAgICAgICAgICAgIFthcnJheV0kaXRl
bXMgPSBHZXQtQ2hpbGRJdGVtICIkcGF0aFxTaWRlY2FyIiB8ID8geyAkXy5Qcm9wZXJ0eSAtbWF0Y2ggIi4vVXNlciIgfQ0KICAgICAgICAgICAgICAgICAg
ICBpZiAoJGl0ZW1zKSB7DQogICAgICAgICAgICAgICAgICAgICAgICBBZGREaXNwbGF5IChbcmVmXSRpdGVtcykNCiAgICAgICAgICAgICAgICAgICAgICAg
ICRpdGVtcyB8IFByb2Nlc3NTaWRlY2FyIC1jdXJyZW50VXNlciAkdXNlclNpZA0KICAgICAgICAgICAgICAgICAgICB9DQogICAgICAgICAgICAgICAgfQ0K
ICAgICAgICAgICAgICAgIGlmIChUZXN0LVBhdGggIiR1c2VyUGF0aFxFeHBlY3RlZFNDRVBDZXJ0cyIpIHsNCiAgICAgICAgICAgICAgICAgICAgW2FycmF5
XSRpdGVtcyA9IEdldC1DaGlsZEl0ZW0gIiR1c2VyUGF0aFxFeHBlY3RlZFNDRVBDZXJ0cyINCiAgICAgICAgICAgICAgICAgICAgaWYgKCRpdGVtcykgew0K
ICAgICAgICAgICAgICAgICAgICAgICAgQWRkRGlzcGxheSAoW3JlZl0kaXRlbXMpDQogICAgICAgICAgICAgICAgICAgICAgICAkaXRlbXMgfCBQcm9jZXNz
Q2VydHMNCiAgICAgICAgICAgICAgICAgICAgfQ0KICAgICAgICAgICAgICAgIH0NCiAgICAgICAgICAgIH0NCiAgICAgICAgfQ0KICAgIH0NCg0KICAgICMg
RGlzcGxheSB0aW1lbGluZQ0KICAgIFdyaXRlLUhvc3QgIiINCiAgICBXcml0ZS1Ib3N0ICJPQlNFUlZFRCBUSU1FTElORToiIC1Gb3JlZ3JvdW5kQ29sb3Ig
TWFnZW50YQ0KICAgIFdyaXRlLUhvc3QgIiINCiAgICAkb2JzZXJ2ZWRUaW1lbGluZSB8IFNvcnQtT2JqZWN0IC1Qcm9wZXJ0eSBEYXRlIHwNCiAgICBGb3Jt
YXQtVGFibGUgQHsNCiAgICAgICAgTGFiZWwgICAgICA9ICJEYXRlIg0KICAgICAgICBFeHByZXNzaW9uID0geyAkXy5EYXRlLlRvU3RyaW5nKCJ1IikgfSAN
CiAgICB9LCANCiAgICBAew0KICAgICAgICBMYWJlbCAgICAgID0gIlN0YXR1cyINCiAgICAgICAgRXhwcmVzc2lvbiA9DQogICAgICAgIHsNCiAgICAgICAg
ICAgIHN3aXRjaCAoJF8uQ29sb3IpIHsNCiAgICAgICAgICAgICAgICAnUmVkJyB7ICRjb2xvciA9ICI5MSI7IGJyZWFrIH0NCiAgICAgICAgICAgICAgICAn
WWVsbG93JyB7ICRjb2xvciA9ICc5Myc7IGJyZWFrIH0NCiAgICAgICAgICAgICAgICAnR3JlZW4nIHsgJGNvbG9yID0gIjkyIjsgYnJlYWsgfQ0KICAgICAg
ICAgICAgICAgIGRlZmF1bHQgeyAkY29sb3IgPSAiMCIgfQ0KICAgICAgICAgICAgfQ0KICAgICAgICAgICAgJGUgPSBbY2hhcl0yNw0KICAgICAgICAgICAg
IiRlWyR7Y29sb3J9bSQoJF8uU3RhdHVzKSRlWzBtIg0KICAgICAgICB9DQogICAgfSwNCiAgICBEZXRhaWwNCg0KICAgIFdyaXRlLUhvc3QgIiINCn0NCg0K
RW5kIHsNCiAgICAjIFJlbW92ZSB0aGUgcmVnaXN0cnkgaW5mbyBpZiBpdCBleGlzdHMNCiAgICBpZiAoVGVzdC1QYXRoICJIS0NVOlxFU1BTdGF0dXMudG1w
Iikgew0KICAgICAgICBSZW1vdmUtSXRlbSAtUGF0aCAiSEtDVTpcRVNQU3RhdHVzLnRtcCIgLVJlY3Vyc2UgLUZvcmNlDQogICAgfQ0KDQogICAgIyBSZW1v
dmUgdGhlIHRlbXAgZm9sZGVyIGluZm8gaWYgaXQgZXhpc3RzDQogICAgaWYgKFRlc3QtUGF0aCAiJCgkZW52OlRFTVApXEVTUFN0YXR1cy50bXAiKSB7DQog
ICAgICAgIFJlbW92ZS1JdGVtICIkKCRlbnY6VEVNUClcRVNQU3RhdHVzLnRtcCIgLVJlY3Vyc2UgLUZvcmNlDQogICAgfQ0KfQ0KDQoNCg0KIyBTSUcgIyBC
ZWdpbiBzaWduYXR1cmUgYmxvY2sNCiMgTUlJb0V3WUpLb1pJaHZjTkFRY0NvSUlvQkRDQ0tBQUNBUUV4RHpBTkJnbGdoa2dCWlFNRUFnRUZBREI1Qmdvcg0K
IyBCZ0VFQVlJM0FnRUVvR3N3YVRBMEJnb3JCZ0VFQVlJM0FnRWVNQ1lDQXdFQUFBUVFIOHc3WUZsTENFNjNKTkxHDQojIEtYN3pVUUlCQUFJQkFBSUJBQUlC
QUFJQkFEQXhNQTBHQ1dDR1NBRmxBd1FDQVFVQUJDRE9ZU1pPTk1oK0FaMmQNCiMgWTZQOENvVDFpcU1kQjdtWUVhWTJKUHJpR08rTk5xQ0NJUll3Z2dXTk1J
SUVkYUFEQWdFQ0FoQU9teGlPK2RBdA0KIyA1Ky9iVU9JSVFCaGFNQTBHQ1NxR1NJYjNEUUVCREFVQU1HVXhDekFKQmdOVkJBWVRBbFZUTVJVd0V3WURWUVFL
DQojIEV3eEVhV2RwUTJWeWRDQkpibU14R1RBWEJnTlZCQXNURUhkM2R5NWthV2RwWTJWeWRDNWpiMjB4SkRBaUJnTlYNCiMgQkFNVEcwUnBaMmxEWlhKMElF
RnpjM1Z5WldRZ1NVUWdVbTl2ZENCRFFUQWVGdzB5TWpBNE1ERXdNREF3TURCYQ0KIyBGdzB6TVRFeE1Ea3lNelU1TlRsYU1HSXhDekFKQmdOVkJBWVRBbFZU
TVJVd0V3WURWUVFLRXd4RWFXZHBRMlZ5DQojIGRDQkpibU14R1RBWEJnTlZCQXNURUhkM2R5NWthV2RwWTJWeWRDNWpiMjB4SVRBZkJnTlZCQU1UR0VScFoy
bEQNCiMgWlhKMElGUnlkWE4wWldRZ1VtOXZkQ0JITkRDQ0FpSXdEUVlKS29aSWh2Y05BUUVCQlFBRGdnSVBBRENDQWdvQw0KIyBnZ0lCQUwvbWtITm8zcnZr
WFVvOE1DSXdhVFBzd3FjbExza2hQZktLMkZuQzRTbW5QVmlyZHByTnJuc2JoQTNFDQojIE1CL3pHNlE0RnV0V3hwZHRIYXV5ZWZMS0VkTGtYOVlGUEZJUFVo
L0duaFdsZnI2ZnFWY1dXVlZ5cjJpVGNNS3kNCiMgdW5XWmFuTXlsTkVRUkJBdTM0THpCNFRtZER0dGNlSXREQnZ1SU5YSklCMWpLUzNPN0Y1T3lKUDRJV0di
Tk9zRg0KIyB4bDdzV3hxODY4blB6YXcwUUYreGVtYnVkOGhJcUdaWFY1OVVXSTRNSzdkUHB6RFpWdTdLZTEzanJjbFBYdVUxDQojIDV6SEwycE5lM0k2UGdO
cTJrWmhBa0huRGVNZTJzY1MxYWhnNEF4Q04yTlEzcEM0RmZZajFnajRRa1hDclZZSkINCiMgTXRmYkJITXFicEVCZkNGTTFMeXVHd04xWFhobTJUb3hSSm96
UUw4STExcEpwTUxtcWFCbjNhUW52S0ZQT2JVUg0KIyBXQmYzSkZ4R2oyVDN3V21JZHBoMlBWbGRRbmFIaVpkcGVranc0S0lTRzJhYWRNcmVTeDduRG1PdTV0
VHZrcEk2DQojIG5qM2NBT1JGSlltMm1rUVpLMzdBbExUU1lXM3JNOW5GMzBzRUFNeDlISlhEai9jaHNySVJ0N3QvOHRXTWNDeEINCiMgWUtxeFl4aEVsUnAy
WW43MmdMRDc2R1NtTTlHSkIrRzl0K1pEcEJpNHBuY0I0UStVRENFZHNsUXBKWWxzNVE1Uw0KIyBVVWQwdmlhc3RrRjEzbnFzWDQwL3lielRRUkVTVytVUVVP
c3h4Y3B5RmlJSjMzeE1kVDlqN0NGZnhDQlJhMit4DQojIHE0YUxUOExXUlYrZElQeWhIc1hBajZLeGZnb21tZlhrYVMrWUhTMzEyYW15SGVVYkFnTUJBQUdq
Z2dFNk1JSUINCiMgTmpBUEJnTlZIUk1CQWY4RUJUQURBUUgvTUIwR0ExVWREZ1FXQkJUczErT0MwbkZkWkV6ZkxtYy81N3FZcmh3UA0KIyBUekFmQmdOVkhT
TUVHREFXZ0JSRjY2S3Y5SkxMZ2pFdFVZdW5weUdkODIzSUR6QU9CZ05WSFE4QkFmOEVCQU1DDQojIEFZWXdlUVlJS3dZQkJRVUhBUUVFYlRCck1DUUdDQ3NH
QVFVRkJ6QUJoaGhvZEhSd09pOHZiMk56Y0M1a2FXZHANCiMgWTJWeWRDNWpiMjB3UXdZSUt3WUJCUVVITUFLR04yaDBkSEE2THk5allXTmxjblJ6TG1ScFoy
bGpaWEowTG1Odg0KIyBiUzlFYVdkcFEyVnlkRUZ6YzNWeVpXUkpSRkp2YjNSRFFTNWpjblF3UlFZRFZSMGZCRDR3UERBNm9EaWdOb1kwDQojIGFIUjBjRG92
TDJOeWJETXVaR2xuYVdObGNuUXVZMjl0TDBScFoybERaWEowUVhOemRYSmxaRWxFVW05dmRFTkINCiMgTG1OeWJEQVJCZ05WSFNBRUNqQUlNQVlHQkZVZElB
QXdEUVlKS29aSWh2Y05BUUVNQlFBRGdnRUJBSENndjBOYw0KIyBWZWM0WDZDamRCczl0aGJYOTc5WEI3MmFyS0dITE95Rlhxa2F1eUw0aHhwcFZDTHRwSWgz
YmIwYUZQUVRTbm92DQojIExiYzQ3L1QvZ0xuNG9mZnljdDRrdkZJRHlFN1FLdDc2TFZiUCtmVDNyREI2bW91eVh0VFAwVU5FbTBNaDY1WnkNCiMgb1VpMG1j
dWRUNmNHQXhOM0owVFU1My9vV2Fqd3Z5OExwdW55TkR6czl3UEhoNmpTVEVBWk5VWnFhVlN3dUtGVw0KIyBqdXlrMVQzb3NkejlITmowZDFwY1ZJeHY3NkZR
UGZ4MkNXaUVuMi9LMnlDTk5XQWNBZ1BMSUxDc1dLQU9RR1BGDQojIG1DTEJzbG4xVld2UEo2dHNkczV2SXkzMGZuRnFJMnNpL3hLNFZDMG5mdGc2MmZDMmg1
YjlXOUZjckJqRFRaOXoNCiMgdHdHcG4xZXFYaWppdVpRd2dnYXVNSUlFbHFBREFnRUNBaEFITmplM0pGUjgyRWVzL1NobUtsNWJNQTBHQ1NxRw0KIyBTSWIz
RFFFQkN3VUFNR0l4Q3pBSkJnTlZCQVlUQWxWVE1SVXdFd1lEVlFRS0V3eEVhV2RwUTJWeWRDQkpibU14DQojIEdUQVhCZ05WQkFzVEVIZDNkeTVrYVdkcFky
VnlkQzVqYjIweElUQWZCZ05WQkFNVEdFUnBaMmxEWlhKMElGUnkNCiMgZFhOMFpXUWdVbTl2ZENCSE5EQWVGdzB5TWpBek1qTXdNREF3TURCYUZ3MHpOekF6
TWpJeU16VTVOVGxhTUdNeA0KIyBDekFKQmdOVkJBWVRBbFZUTVJjd0ZRWURWUVFLRXc1RWFXZHBRMlZ5ZEN3Z1NXNWpMakU3TURrR0ExVUVBeE15DQojIFJH
bG5hVU5sY25RZ1ZISjFjM1JsWkNCSE5DQlNVMEUwTURrMklGTklRVEkxTmlCVWFXMWxVM1JoYlhCcGJtY2cNCiMgUTBFd2dnSWlNQTBHQ1NxR1NJYjNEUUVC
QVFVQUE0SUNEd0F3Z2dJS0FvSUNBUURHaGpVR1NiUEJQWEpKVVZYSA0KIyBKUVBFOHBFM3FaZFJvZGJTZzlHZVRLSnRvTERNZy9sYTloR2hSQlZDWDZTSTgy
ajZmZk9jaVF0L25SK2VEek1mDQojIFVCTUxKbk9XYmZoWHFBSjkvVU8waE5vUjhYT3hzKzRyZ0lTS0loamY2OW85eEJkL3F4a3JQa0xjWjQ3cVVUM3cNCiMg
MWxiVTV5Z3Q2OU94dFhYbkh3WmxqWlFwMDluc2FkL1prSWRHQUh2YlJFR0ozSHhxVjNyd04zbWZYYXpMNklSaw0KIyB0Rkx5ZGtmM1lZTVozViswVkFzaGFH
NDNJYnRBckYreTNrcDl6dlU1RW1mdkRxVmpiT1NteFIzTk5nMWMxZVliDQojIHFNRmtkRUNud0hMRnVrNGZzYlZZVFhuKzE0OXprNndzT2VLbFNOYndzREVU
cVZjcGxpY3U5WWVtajA1MkZWVW0NCiMgY0pnbWY2QWFSeUJENDBOamdIdDFiaWNsa0pnNk9CR3o5dmFlNWp0YjdJSGVJaFRaZ2lySGtyK2czdU0rb25QNg0K
IyA1eDlhYkpUeVVwVVJLMWgwUUNpcmMwUE8zMHFoSEdzNHhTbnp5cXFXYzBKb243WkdzNTA2bzlVRDRML3dvanpLDQojIFF0d1lTSDhVTk0vU1RLdnZtejMr
RHJoa0t2cDFLQ1JCN1VLL0JaeG1TVkpROUZIek5rbE5peURTTEZjMWVTdW8NCiMgODBWZ3ZDT05XUGZjWWQ2VC9qbkErYkl3cFV6WDZaaEtXRDdUQTRqK3M0
L1RYa3QyRWxHVHlZd01PMXVLSXFqQg0KIyBKZ2o1RkJBU0EzMWZJN3RrNDJQZ3B1RSs5c0owc2o4ZUNYYnNxMTFHZGVKZ28xZ0pBU2dBRG9SVTdzN3BYY2hl
DQojIE1CSzlScDYxMDNhNTBnNXJtUXpTTTdUTnNRSURBUUFCbzRJQlhUQ0NBVmt3RWdZRFZSMFRBUUgvQkFnd0JnRUINCiMgL3dJQkFEQWRCZ05WSFE0RUZn
UVV1aGJaYlUyRkwzTXBkcG92ZFl4cUlJK2V5Rzh3SHdZRFZSMGpCQmd3Rm9BVQ0KIyA3TmZqZ3RKeFhXUk0zeTVuUCtlNm1LNGNEMDh3RGdZRFZSMFBBUUgv
QkFRREFnR0dNQk1HQTFVZEpRUU1NQW9HDQojIENDc0dBUVVGQndNSU1IY0dDQ3NHQVFVRkJ3RUJCR3N3YVRBa0JnZ3JCZ0VGQlFjd0FZWVlhSFIwY0Rvdkwy
OWoNCiMgYzNBdVpHbG5hV05sY25RdVkyOXRNRUVHQ0NzR0FRVUZCekFDaGpWb2RIUndPaTh2WTJGalpYSjBjeTVrYVdkcA0KIyBZMlZ5ZEM1amIyMHZSR2xu
YVVObGNuUlVjblZ6ZEdWa1VtOXZkRWMwTG1OeWREQkRCZ05WSFI4RVBEQTZNRGlnDQojIE5xQTBoakpvZEhSd09pOHZZM0pzTXk1a2FXZHBZMlZ5ZEM1amIy
MHZSR2xuYVVObGNuUlVjblZ6ZEdWa1VtOXYNCiMgZEVjMExtTnliREFnQmdOVkhTQUVHVEFYTUFnR0JtZUJEQUVFQWpBTEJnbGdoa2dCaHYxc0J3RXdEUVlK
S29aSQ0KIyBodmNOQVFFTEJRQURnZ0lCQUgxWmpzQ1R0bStZcVVRaUFYNW0xdGdoUXVHd0dDNFFUUlBQTUZQT3Z4ajd4MUJkDQojIDRrc3ArM0NLRGFvcGFm
eHB3YzhkQitrK1lNallDK1ZjVzlkdGgvcUVJQ1UwTVdmTnRoS1diOFJRVEdJZERBaUMNCiMgcUJhOXFWYlBGWE9OQVNJbHpwVnBQMGQzKzNKMEZOZi9xMCtL
TEhxcmhjMURYKzFndHFwUGtXYWVMSjdnaXF6bA0KIyAvWXk4WkNhSGJKSzluWHpRY0FwODc2aThkVSs2V3ZlcEVMSmQ2ZjhvVkludzFZcHhkbVhhelBCeW95
UDZ3Q2VDDQojIFJLNlpKeHVySkI0bXdiZmVLdXYybnJGNW1ZR2pWb2FyQ2tYSjM4U05vT2VZKy91bW5YS3Z4TWZCd1dweDJjWVQNCiMgZ0FuRXRwL05oNGNr
dTAralNibDNacEh4Y3B6cFN3SlNwemQrazFPc094MElTUStVelRsNjNmOGxZNWtuTEQwLw0KIyBhNmZ4WnNOQnpVKzJRSnNoSVVEUXR4TWt6ZHdkZURya25x
M2xOSEdTMXlacjVEaHpxNllCVDcwL08zaXRUSzM3DQojIHhKVjc3UXBmTXptSFFYaDZPT21jNGQwai9SMG8wOGY1NlBHWVgvc3IySDd5UnAxMUxCNG5MQ2Ji
YnhWN0hobUwNCiMgTnJpVDFPYnlGNWxaeW5Ed043K1lBTjhnRms4bisyQm5GcUZtdXQxVndEb3BockNZb0N2dGxVRzNPdFVWbURHMA0KIyBZZ2tQQ3IyQjJS
UCt2NlRSODFmWnZBVDZndDR5M3dTSjhBRE5YY0w1MENOL0FBdmtkZ0ltMmZCbGRrS21LWWNKDQojIFJ5dm1meHFraFEvOG1KYjJWVlFySDRENndQSU9LK1hX
KzZrdlJCVks1eE1PSGRzM09CcWhLL2J0MW56OE1JSUcNCiMgc0RDQ0JKaWdBd0lCQWdJUUNLMUFzbURTbkV5ZlhzMnB2Wk91MlRBTkJna3Foa2lHOXcwQkFR
d0ZBREJpTVFzdw0KIyBDUVlEVlFRR0V3SlZVekVWTUJNR0ExVUVDaE1NUkdsbmFVTmxjblFnU1c1ak1Sa3dGd1lEVlFRTEV4QjNkM2N1DQojIFpHbG5hV05s
Y25RdVkyOXRNU0V3SHdZRFZRUURFeGhFYVdkcFEyVnlkQ0JVY25WemRHVmtJRkp2YjNRZ1J6UXcNCiMgSGhjTk1qRXdOREk1TURBd01EQXdXaGNOTXpZd05E
STRNak0xT1RVNVdqQnBNUXN3Q1FZRFZRUUdFd0pWVXpFWA0KIyBNQlVHQTFVRUNoTU9SR2xuYVVObGNuUXNJRWx1WXk0eFFUQS9CZ05WQkFNVE9FUnBaMmxE
WlhKMElGUnlkWE4wDQojIFpXUWdSelFnUTI5a1pTQlRhV2R1YVc1bklGSlRRVFF3T1RZZ1UwaEJNemcwSURJd01qRWdRMEV4TUlJQ0lqQU4NCiMgQmdrcWhr
aUc5dzBCQVFFRkFBT0NBZzhBTUlJQ0NnS0NBZ0VBMWJRdlF0QW9yWGkzWGRVNVdSdXhpRUwxTTR6cg0KIyBQWUdYY01XN3hJVW1NSitram1qWVhQWHJOQ1FI
NFV0UDAzaEQ5QmZYSHRyNTB0Vm5HbEpQRHFGWC9JaVp3WkhNDQojIGdRTStUWEFrWkxPTjRnaDlOSDFNZ0ZjU2EwT2FtZkxGT3gveTc4dEhXaE9tVExNQklD
WHpFTk9Mc3ZzSThJcmcNCiMgblFuQVphZjZtSUJKTlljOVVSbm9rQ0Y0UlM2aG55emhHTUlhek1YdWswbHdRaktQKzhicUhQTmxhSkdpVFV5Qw0KIyBFVWhT
YU40UXZSUlhYZWdZRTJYRmY3SlBoU3hJcEZhRU5kYjVMcHlxQUJYUk4vNGFCcFRDZk1qcUd6TG15c0wwDQojIHA2TUREblNscnptMnEyQVM0K2pXdWZjeDRk
eXQ1QmlnMk1FalIwZXpvUTl1bzZ0dG1BYURHN2RxWnkzU3ZVUWENCiMga2hDQmo3QTdDZGZIbXpKYXd2OXFZRlNMU2NHVDdlRzBYT0J2NnliNWpOV3krVGdR
NXVyT2tmVyswL3R2azJFMA0KIyBYTHlUUlNpRE5pcG1LRit3Yzg2TEppVUdzb1BVWFBZVkdVenRZdUJlTS9MbzZPd0twN0FESzVHeU5ubSs5NjBJDQojIEhu
V21aY3k3NDBoUTgzZVJHdjdiVUtKR3lHRlltUFY4QWhZOGd5aXRPWWJzMUxjTlU5RDRSK1oxTUkzc01KTjINCiMgRktaYlMxMTBZVTAvRXBGMjNyOVl5M0lR
S1VIdzFjVnRKblpvRVVFVFdKcmNKaXNCOUlsTldkdDR6NEZLUGtCSA0KIyBYOG1CVUhPRkVDTWhXV0NLWkZUQnpDRWE2RGdaZkdZY3pYZzRSVENaVC85alQw
eTdxZzBJVTBGOFdEMUhzL3EyDQojIDdJd3lDUUxNYkR3TVZoRUNBd0VBQWFPQ0FWa3dnZ0ZWTUJJR0ExVWRFd0VCL3dRSU1BWUJBZjhDQVFBd0hRWUQNCiMg
VlIwT0JCWUVGR2czNE91Mk8vaGZFWWI3L21GN0NJaGw5RTVDTUI4R0ExVWRJd1FZTUJhQUZPelg0NExTY1Yxaw0KIyBUTjh1WnovbnVwaXVIQTlQTUE0R0Ex
VWREd0VCL3dRRUF3SUJoakFUQmdOVkhTVUVEREFLQmdnckJnRUZCUWNEDQojIEF6QjNCZ2dyQmdFRkJRY0JBUVJyTUdrd0pBWUlLd1lCQlFVSE1BR0dHR2gw
ZEhBNkx5OXZZM053TG1ScFoybGoNCiMgWlhKMExtTnZiVEJCQmdnckJnRUZCUWN3QW9ZMWFIUjBjRG92TDJOaFkyVnlkSE11WkdsbmFXTmxjblF1WTI5dA0K
IyBMMFJwWjJsRFpYSjBWSEoxYzNSbFpGSnZiM1JITkM1amNuUXdRd1lEVlIwZkJEd3dPakE0b0RhZ05JWXlhSFIwDQojIGNEb3ZMMk55YkRNdVpHbG5hV05s
Y25RdVkyOXRMMFJwWjJsRFpYSjBWSEoxYzNSbFpGSnZiM1JITkM1amNtd3cNCiMgSEFZRFZSMGdCQlV3RXpBSEJnVm5nUXdCQXpBSUJnWm5nUXdCQkFFd0RR
WUpLb1pJaHZjTkFRRU1CUUFEZ2dJQg0KIyBBRG9qUkQyTkNIYnVqN3c2bWROVzRBSWFwZmhJTlBNc3R1WjBadmVVY3JFQXlxOXNNQ2NURXA2UVJKOUwvWjZq
DQojIGZDYlZON3c2WFVodGxkVS9TZlFudXhhQlJWRDluTDIyaGVCMmZqZHh5eUwzV3FxUXovV1RhdVBySU5IVlVIbUkNCiMgbW9xS3diYTlvVWdZZnR6WWdC
b1JHUmpOWVptQlZ2Yko0M2JueE9RYlgwUDRQcFQvZGprOW50U1p6MHJkS090Zg0KIyBKcUdWV0VqVkd2N1hKei85a05GMmh0MGNzR0JjOHcybzd1Q0pvYjA1
NFRoTzJtNjdOcDM3NVNGVFdzUEs2V3J4DQojIG9qN2JRN2d6eUU4NEZKS1o5ZDNPVkczWlhRSVVIMEF6ZkFQaWxiTENJWFZ6VXN0RzJNUTBIS0tsUzQzTmIz
WTMNCiMgTElVL0dzNG02Umkra0Fld1EzK1ZpQ0NDY1BETXl1LzlLVFZjSDRrNFZmYzNpb3NKb2NzTDZURWEveTRaWERseA0KIyA0YjZjcHdvRzFpWm50NUxt
VGwvZWVxeEp6eTZrZEpLdDJ6eWtuSVlmNDhGV0d5c2ovNCsxNm9oN2NHdm1vTHI5DQojIE9qOUZwc1RvRnBGU2kwSEFTSVJMbGsyclJFRGpqZkFWS003dDhS
aFdCeW92RU1RTUNHUThNNCt1S0l3OHk0K0kNCiMgQ3cyL08vVE9IbnVPNzdYcnk3ZndkeFBtNXlnL3JCS3VwUzhpYkVINWdsd1Zac3hzRHNyRmhzUDJKak1N
QjB1Zw0KIyAwd2NDYW1wQU1FaExOS2hSSUx1dEc0VUk0bGtOYmNvRlVDdnFTaHllcGYyZ3B4OEdkT2Z5MWxLUS9hK0ZTQ0g1DQojIFZ6dTBuQVB0aGtYMHRH
RnV2MmppSm1DRzZzaXZxZjZVSGVkakd6cUdWbmhPTUlJR3ZEQ0NCS1NnQXdJQkFnSVENCiMgQzY1bXZGcTZmNVdIeHZucEJPTXpCREFOQmdrcWhraUc5dzBC
QVFzRkFEQmpNUXN3Q1FZRFZRUUdFd0pWVXpFWA0KIyBNQlVHQTFVRUNoTU9SR2xuYVVObGNuUXNJRWx1WXk0eE96QTVCZ05WQkFNVE1rUnBaMmxEWlhKMElG
UnlkWE4wDQojIFpXUWdSelFnVWxOQk5EQTVOaUJUU0VFeU5UWWdWR2x0WlZOMFlXMXdhVzVuSUVOQk1CNFhEVEkwTURreU5qQXcNCiMgTURBd01Gb1hEVE0x
TVRFeU5USXpOVGsxT1Zvd1FqRUxNQWtHQTFVRUJoTUNWVk14RVRBUEJnTlZCQW9UQ0VScA0KIyBaMmxEWlhKME1TQXdIZ1lEVlFRREV4ZEVhV2RwUTJWeWRD
QlVhVzFsYzNSaGJYQWdNakF5TkRDQ0FpSXdEUVlKDQojIEtvWklodmNOQVFFQkJRQURnZ0lQQURDQ0Fnb0NnZ0lCQUw1cWM1LzJsU0dybGpDNlcyM21XYU8x
NlAyUkh4akUNCiMgaUR0cW1lT2x3ZjBLTUNCREVyNEl4SFJHZDcrTDY2MHg1WGx0U1ZoaEs2NHppOUNlQzlCNmxVZFhNMHM3MUVPYw0KIyBSZTgrQ0VKcCsz
UjJPOG9vNzZFTzdvNXRMdXNseGRyOVFxODJhS2NwQTlPLy9YNlFFK0FjYVUvYnlhQ2FnTEQvDQojIEdMb1ViMzVTZldIaDQzck9IM2JwTEV4N3BaN2F2Vm5w
VVZtUHZreFQ4YzJhMnlDMFdNcDhoTXU2MHRaUjBDaGENCiMgVjc2TmhuajM3REVZVFg5UmVOWjhoSU9ZZTRqbDcvcjQxOUN2RVlWSXJINnNOMDB5eDQ5Ym9V
dXVtRjlpMlQ4VQ0KIyB1S0duOTk2NmZSNVg2a2dYajNvNVdIaEhWTytOQmlrRE8wbWxVaDkwMndTL0VlaDhGL1VGYVJwMXo1U25ST0h3DQojIFNKK1FRUlox
ZmlzRDhVVFZEU3VwV0pOc3RWa2lxTHErSVNUZEVqSktHalZmSWNzZ0E0bDljYms4U21semRkaDQNCiMgRWZ2RnJwVk5uZXM0YzE2SmlkajVYaVBWZHNuNW4x
MGp4bUdweG9NYzZpUGtvYURoaTZKakhkNWliZmRwNXV6SQ0KIyBYcDRQMHdYa2dOcytDTy9DYWNCcVUwUjRrKzhoNmdZbGRwNEZDTWdyWGRLV2ZNNE4wdTI1
T0VBdUVhM0p5aWR4DQojIFc0OGp3QnFJSnFJbWQ5M05SeHZkMWFlcFNlTmVSRVhBdTJ4VURFVzhhcXpGUURZbXI5Wk9OdWMyTWhUTWl6Y2gNCiMgTlVMcFVF
b0E2VnZhN2IxWENCKzFyeHZiS21McWZZL00vU2RWNm13V1R5ZVZ5NVovSmt2TUZwblF5NXdSMTRHSg0KIyBjdjZkUTRhRUtPWDVBZ01CQUFHamdnR0xNSUlC
aHpBT0JnTlZIUThCQWY4RUJBTUNCNEF3REFZRFZSMFRBUUgvDQojIEJBSXdBREFXQmdOVkhTVUJBZjhFRERBS0JnZ3JCZ0VGQlFjRENEQWdCZ05WSFNBRUdU
QVhNQWdHQm1lQkRBRUUNCiMgQWpBTEJnbGdoa2dCaHYxc0J3RXdId1lEVlIwakJCZ3dGb0FVdWhiWmJVMkZMM01wZHBvdmRZeHFJSStleUc4dw0KIyBIUVlE
VlIwT0JCWUVGSjlYTEFOM0RpZ1ZrR2FsWTE3dVQ1SWZkcUJiTUZvR0ExVWRId1JUTUZFd1Q2Qk5vRXVHDQojIFNXaDBkSEE2THk5amNtd3pMbVJwWjJsalpY
SjBMbU52YlM5RWFXZHBRMlZ5ZEZSeWRYTjBaV1JITkZKVFFUUXcNCiMgT1RaVFNFRXlOVFpVYVcxbFUzUmhiWEJwYm1kRFFTNWpjbXd3Z1pBR0NDc0dBUVVG
QndFQkJJR0RNSUdBTUNRRw0KIyBDQ3NHQVFVRkJ6QUJoaGhvZEhSd09pOHZiMk56Y0M1a2FXZHBZMlZ5ZEM1amIyMHdXQVlJS3dZQkJRVUhNQUtHDQojIFRH
aDBkSEE2THk5allXTmxjblJ6TG1ScFoybGpaWEowTG1OdmJTOUVhV2RwUTJWeWRGUnlkWE4wWldSSE5GSlQNCiMgUVRRd09UWlRTRUV5TlRaVWFXMWxVM1Jo
YlhCcGJtZERRUzVqY25Rd0RRWUpLb1pJaHZjTkFRRUxCUUFEZ2dJQg0KIyBBRDJ0SGg5Mm1WdmpPSVFTUjlsRGtmWVIyNXRPQ0IzUktFL1AwOXg3Z1VzbVhx
dDQwb3VSbDNsais4UWlvVllxDQojIDNpZ3B3clB2Qm1aZHJsV0JiMEh2cVQwMG5GU1hnbVVyREtOU1FxR1RkcGpIc1B5K0xhYWxUVzBxVmp2VUJoY0gNCiMg
ekJNdXRCNkh6ZWxlZGJEQ3pGelV5MzRWYXJQbnZJV3JxVm9nSzBxTThnSmhoLytxREVBSWRPL0trWWVzTHlUVg0KIyBPb0o0ZVRxN2dqOVVGQUwxVXJ1Skts
VG5DVmFNMlVlVVVXLzh6M2Z2anhoTjZoZFQ5OFZyMkZZbENTN01iYjRIDQojIHY1c3dPK2FBWHhXVW0zV3BCeVh0Z1ZReGlCbFRWWXpxZkxEYmU5UHBCS0RC
ZmsrcmFiVEZEWlhvVWtlN3pQZ3QNCiMgZDcvZnZXVGxDczMwVkFHRXNzaEptTGJKNlpiUS94bGwvSGpPOUpiTlZla0J2MlRnZW0rbUxwdFI3eUlycGFpZA0K
IyBSSlhySStVekI2dkFsay84YTF1N2NJcVYweWVmNHVhWkZPUk5la1VnUUhUcWRkbXNQQ0VJWVFQN3hHeFpCSWhkDQojIG1tNGJoWXNWQTZHMldnTkZZYWdM
REJ6cG1rOTEwNFdRell1Vk5zeHlvVkxPYmh4M1J1Z2FFR3J1K1Nvalc0ZEgNCiMgUG9XclVoZnROcEZDNUg3UUVZN01oS1J5ckJlN3VjeWtXN2VhQ3VXQnNC
YjRIT0tSRlZEY3JaZ2R3YVNJcU1EaQ0KIyBDTGc0RCtUUFZnS3gyRWdFZGVvSE5IVDlsM1pEQkQrWGdiRisyMy96QmplQ3R4eitkTC85TldSNlAyZVpSaTd6
DQojIGNFTzF4d2NkY3FKc3l6L0pjZUVOYzJTZzhoM0tlRlVDUzd0cEZrN0NyRHFrTUlJSFd6Q0NCVU9nQXdJQkFnSVENCiMgQ0xHZnpiUGE4N0F4VlZnSUFT
OEE2VEFOQmdrcWhraUc5dzBCQVFzRkFEQnBNUXN3Q1FZRFZRUUdFd0pWVXpFWA0KIyBNQlVHQTFVRUNoTU9SR2xuYVVObGNuUXNJRWx1WXk0eFFUQS9CZ05W
QkFNVE9FUnBaMmxEWlhKMElGUnlkWE4wDQojIFpXUWdSelFnUTI5a1pTQlRhV2R1YVc1bklGSlRRVFF3T1RZZ1UwaEJNemcwSURJd01qRWdRMEV4TUI0WERU
SXoNCiMgTVRFeE5UQXdNREF3TUZvWERUSTJNVEV4TnpJek5UazFPVm93WXpFTE1Ba0dBMVVFQmhNQ1IwSXhGREFTQmdOVg0KIyBCQWNUQzFkb2FYUnNaWGtn
UW1GNU1SNHdIQVlEVlFRS0V4VkJUa1JTUlZkVFZFRlpURTlTTGtOUFRTQk1WRVF4DQojIEhqQWNCZ05WQkFNVEZVRk9SRkpGVjFOVVFWbE1UMUl1UTA5TklF
eFVSRENDQWlJd0RRWUpLb1pJaHZjTkFRRUINCiMgQlFBRGdnSVBBRENDQWdvQ2dnSUJBTU9rWWtMcHpOSDRZMWdVWEY3OTl1RjBDcndXL0xtZTY3NitDOWFa
T0pZeg0KIyBwcTMvRElhODFvV3Y5YjRiMFd3THBKVnUwZk9rQW14STZvY3U0dWY2MTNqRE1XMEdmVjRkUm9kdXRyeWZ1RHVpDQojIHQ0cm5kdkpBNkRJczBZ
RzV4TmxLVGtZOEFJdkJQM0l3RXpVRDFmNTdKNUdpQXBySEdlb2M0VXR0ekV1R0EzeVMNCiMgcWxzR0VnMGdDZWhXSnpuVWtoM3lNOFhia3NDMEx1Qm1uWS9k
WkovOGt0Q3dDZDM4Z2ZaRU85VUREU2tpZTRWVA0KIyBZM1Q3VkZiVGlhSDBidytBdmZjUVZ5MkNTd2t3Zm5rZllhZ1NGa0thcitNWXd1N2dxVlh4cmgzVi9H
anZhbDZQDQojIGRNMEE3RWNUcW16ckNSdHZrV0lSNmJweiszQUlINkZyNnlUdUczWGlMSUw2c0svaUYvOWQ0VTJQaUgxdkoveGYNCiMgZGhHajByUTMvTkJS
c1VCQzNsMXc0MUw1cTlVWDFPaDFsVDFPdUo2aFYvdWFuazZKWTNqcG0rT2ZaN1lDVEYySA0KIyBrejV5Nmg5VDdzWTBMVGk2OFZtdHhhL0VnRXRHNkpWTlZz
cVA3V3dFa1FSeHUvMzBxdGp5b1g4bnpTdUY3VG1zDQojIFJnbVoxU0IrSVNjbGVqdXFUTmRoY3ljRGhpMy9JSVNnVkpOUlMvRjZaK1ZRR2YzZmg2T2JkUUxW
d29UMEpuSmoNCiMgYkQ4UHpKMTJPb0tnVmlUUWhuZGFaYmtmcGlWaWZKMXV6V0pyVFc1d0VySCtxdnV0SFZ0NC9zRVpBVlM0UE5mTw0KIyBjSlhSMHMwL0w1
SkhranRNNGFHbDYyZkFIakhqOUpzQ2x1c2o0N2NUNmpST0lxUUk0ZWp6MXNsT29jbE9ldENODQojIEFnTUJBQUdqZ2dJRE1JSUIvekFmQmdOVkhTTUVHREFX
Z0JSb04rRHJ0anY0WHhHRysvNWhld2lJWmZST1FqQWQNCiMgQmdOVkhRNEVGZ1FVMEhkT0ZmUHhhOVllYjVPNUo5VUVpSmtySzk4d1BnWURWUjBnQkRjd05U
QXpCZ1puZ1F3Qg0KIyBCQUV3S1RBbkJnZ3JCZ0VGQlFjQ0FSWWJhSFIwY0RvdkwzZDNkeTVrYVdkcFkyVnlkQzVqYjIwdlExQlRNQTRHDQojIEExVWREd0VC
L3dRRUF3SUhnREFUQmdOVkhTVUVEREFLQmdnckJnRUZCUWNEQXpDQnRRWURWUjBmQklHdE1JR3ENCiMgTUZPZ1VhQlBoazFvZEhSd09pOHZZM0pzTXk1a2FX
ZHBZMlZ5ZEM1amIyMHZSR2xuYVVObGNuUlVjblZ6ZEdWaw0KIyBSelJEYjJSbFUybG5ibWx1WjFKVFFUUXdPVFpUU0VFek9EUXlNREl4UTBFeExtTnliREJU
b0ZHZ1Q0Wk5hSFIwDQojIGNEb3ZMMk55YkRRdVpHbG5hV05sY25RdVkyOXRMMFJwWjJsRFpYSjBWSEoxYzNSbFpFYzBRMjlrWlZOcFoyNXANCiMgYm1kU1Uw
RTBNRGsyVTBoQk16ZzBNakF5TVVOQk1TNWpjbXd3Z1pRR0NDc0dBUVVGQndFQkJJR0hNSUdFTUNRRw0KIyBDQ3NHQVFVRkJ6QUJoaGhvZEhSd09pOHZiMk56
Y0M1a2FXZHBZMlZ5ZEM1amIyMHdYQVlJS3dZQkJRVUhNQUtHDQojIFVHaDBkSEE2THk5allXTmxjblJ6TG1ScFoybGpaWEowTG1OdmJTOUVhV2RwUTJWeWRG
UnlkWE4wWldSSE5FTnYNCiMgWkdWVGFXZHVhVzVuVWxOQk5EQTVObE5JUVRNNE5ESXdNakZEUVRFdVkzSjBNQWtHQTFVZEV3UUNNQUF3RFFZSg0KIyBLb1pJ
aHZjTkFRRUxCUUFEZ2dJQkFFa1JoMlB3TWl5cmF2cjY2Wnd3NlBqbDI0S3pEY0dZTVN4VUtPRVU0YnlrDQojIGNPS2d2UzZWMnplWklzMEQvb3FjdDNoQktU
R0VTU1FXU0EvSmtyMUVNQzA0cUpITy9Ud3Ivc0JEQ0RCTXRKOVgNCiMgQXRPNzVKK29xRGNjTStnOFBvK2pqaHFZSnpLdmJpc1ZVdmRzUHFGbGw1NXZTelJ2
SEdBQTZoanlEeWFrR0xSTw0KIyBjTmFTRlpHZGdPSzJBTWhROEVVTHJFOFJpcmkzRDFST3VxR21VV0txY085YXFQSEJmNXdVd2lhOGc5ODBzVFhxDQojIHVP
NWc0VFdrWnFTdnd0MUJITW11NjlNUjZsb1JBSzE3SHZGY1NpY0s2UG0wemlkMUtTMno0bnRHQjRDZmNnODgNCiMgYUZMb2czY2lQMnRmTWkyeFRucU4xSytZ
bVU4OTRQbDFsQ3AxeEZ2VDZwcm0xMEJzNkJWaUtYZkRmVkZ4WFRCMA0KIyBtSG9ETnFHaS9COCtyeGYyejd1NWZvWFBDekJZVCtRM2N4dG9wdlp0azI5TXBU
WTg4R0hEVkpzRk1Calg3ek02DQojIGFDTktzVEtDMmpiOTJGK2psa2M4Y2xDUVFubDNVNGpxd2JqNHVyMUpCUDVReFFwcldod2RlMCtNaWZEVnAwdkgNCiMg
WnNWWjBwbllNQ0tTRzViVXIzd09VN0VQMzIxRHd2dkVzVGpDeS9YRGd2eThpcFU2dzNHamNRUUZtZ3AvQlgvMA0KIyBKQ0hYKzA0UUowSmtSOVRURlpSMUIr
emgzQ2NLMVpFdFR0dnVaZmpRM3ZpWHdsd3ROTHk0M3ZiZTFKNVdOVHMwDQojIEhqSlhzZmRiaFk1a0U1Umh5ZmF4RkJyMjFLWXgrYitldll5b2xJUzB3UjZO
ZXc2RnFMZ2NjNEdlOTR5YVlWVHENCiMgTVlJR1V6Q0NCazhDQVFFd2ZUQnBNUXN3Q1FZRFZRUUdFd0pWVXpFWE1CVUdBMVVFQ2hNT1JHbG5hVU5sY25Rcw0K
IyBJRWx1WXk0eFFUQS9CZ05WQkFNVE9FUnBaMmxEWlhKMElGUnlkWE4wWldRZ1J6UWdRMjlrWlNCVGFXZHVhVzVuDQojIElGSlRRVFF3T1RZZ1UwaEJNemcw
SURJd01qRWdRMEV4QWhBSXNaL05zOXJ6c0RGVldBZ0JMd0RwTUEwR0NXQ0cNCiMgU0FGbEF3UUNBUVVBb0lHRU1CZ0dDaXNHQVFRQmdqY0NBUXd4Q2pBSW9B
S0FBS0VDZ0FBd0dRWUpLb1pJaHZjTg0KIyBBUWtETVF3R0Npc0dBUVFCZ2pjQ0FRUXdIQVlLS3dZQkJBR0NOd0lCQ3pFT01Bd0dDaXNHQVFRQmdqY0NBUlV3
DQojIEx3WUpLb1pJaHZjTkFRa0VNU0lFSVA2U0pCZ1F6VFdTaFpqNnlwL3VuS0cwc0dud2VyRXFwWWZMc2xuQy83MncNCiMgTUEwR0NTcUdTSWIzRFFFQkFR
VUFCSUlDQUFsSDByM3dJRVBKTXE0bHQxOEQwOEk4T25OZ2VJNlVTNWsxMndVVQ0KIyBEUGl2blU4MlAzelVIeGFQQk0yYXYzWHhza1FZRUdpNVltQ2lNeW0z
b2ZvUU8wTlhSRjRScVZxRjQyTUx4Y1ZIDQojIHh1M0RwcTE1R09KTTZMaGJwVXFydTUxbU9sKzNYdmdhR1FHUENjZUlvN1F6MUNyUHQ0bjRYT01xTzlQQk9h
bkYNCiMgRytYeGVVc1FEclp0djFCU0VFUkpmM0ZudUUxZ3dyL012RTROZjkvU2YvcTZzaHlPS3cwSXFEaXlYRGxYRGxqQQ0KIyBnQTI2R1pxcDZEd1VSS1hr
RGRIM0hxd1hNajYxWnRuRnlvUWx1K012N0poQkRoalc0dFJoRjZCY2ptM3h5UklTDQojIEYwVy9mTUpZbGxmcnRvclhncWhXVGF6V3BxNGU0ckR1cWJ3MkJK
S1BvZlZZYUtlNFl5UTBwOGRST0FGMEdGc1ENCiMgbGRSZ3FEenpUd1B4ZzZJeDJmemNxRFVYZUpIRnQ0OWR3VFh5WjBLMlB3NzYzTktWN2cvT3JRa2J6dXBi
cTlYVA0KIyAwdnZsbjRyNzlmS3JPejhveFVKN1IyTVVLTzRHL2s4OVd3UXZRWUJjbjBWY1E5WFhmY2F1UnBGS2RBeTRMZmtHDQojIHZiS09VaXhCaVdyeEIx
VWorZVZRRFhiOGpaTWdFN2VoNEM4Z3o0WlpzVEZ6c0tEcS9neGdvWFNvSWduL3p0RFENCiMgWkRLSjErQVQzdENKd0lwQ1JaTERYLzBrOTE0V012NmcrWUUw
dzFlTnYvMGo1RzlWY3lpdm1TaE51WndJNVBNNQ0KIyBlZDFCZlJBcWh1Zlo4azZ2Qjh0SVRqZmhIVklVQTU4c1lIZHFmb2pxTWZSYkp6ZVYrcjVMYktjUnRm
akxtZjdnDQojIE9hdTJvWUlESURDQ0F4d0dDU3FHU0liM0RRRUpCakdDQXcwd2dnTUpBZ0VCTUhjd1l6RUxNQWtHQTFVRUJoTUMNCiMgVlZNeEZ6QVZCZ05W
QkFvVERrUnBaMmxEWlhKMExDQkpibU11TVRzd09RWURWUVFERXpKRWFXZHBRMlZ5ZENCVQ0KIyBjblZ6ZEdWa0lFYzBJRkpUUVRRd09UWWdVMGhCTWpVMklG
UnBiV1ZUZEdGdGNHbHVaeUJEUVFJUUM2NW12RnE2DQojIGY1V0h4dm5wQk9NekJEQU5CZ2xnaGtnQlpRTUVBZ0VGQUtCcE1CZ0dDU3FHU0liM0RRRUpBekVM
QmdrcWhraUcNCiMgOXcwQkJ3RXdIQVlKS29aSWh2Y05BUWtGTVE4WERUSTFNRFV5TnpBM05ERTBPVm93THdZSktvWklodmNOQVFrRQ0KIyBNU0lFSUpHdm4v
TUdkaVJoTTlaSkRuRndPQXNid1hBMnA2R1VOamVIK1dYUHJiVUtNQTBHQ1NxR1NJYjNEUUVCDQojIEFRVUFCSUlDQUhSb0gvMElKcFNGYUF4RkhkN1AvTElp
SUwxWGNVZG5rM1dMMVFuazhGaGtuVlU0NW5iZDhLS1oNCiMgL3RqNkQvaGZjbkdvaEg2R2NPRm44Q2VWOW9JZWNIT3NENE9HWGcrQ1RWOTNIZ211SFdaNGlm
SXc1SG56QTlZUw0KIyBaWEJhWnQwS2dMV1ZybXVlQnBtakk5QmRDVG50OVdWNFBNLzRNWmtHK1VVQXphZG0ybkVXOFo3UXRUTUVPbEQyDQojIENlc3c0MzdS
Ty8vMW1lUFdpVUNLcno2eHAzeFRzR2dQcG5PTjBsVnpwdHp5VWdWYkkyc2pVR1R1N1lrTlA5UFANCiMgNktLMEhlWk9seDhOWXllK2NvYmdXUDdzc2VJNTE4
a0k3a3FKU3dpZDhZZVdGdnBnL0Q1ZG9NRFlEb0NVcFhvcQ0KIyBYZG5nOGdXTi9aNzJ2NGYwN09qajVDZU4vUGtmT2dYdktNN3lvYkV3RzVyS0wzME1laHFv
VmFnZXQyWlZxWVdaDQojIHk2OEpDVnBYaCs2VTZqWW5rSzRsdExTbXBoQm1pNGZ2R001ZzcwMElvem5aOU8wQmJkcXhYYTB1RnlydWpmUnANCiMgSXp5UVdm
bnJVNytuWDRCYjRHNFVPNW5nWEdVNVdranR4ZTFXNXZOeTIzaWpoS3VzOGdJRjMrSXdLQ1RCVlVnZg0KIyAvV2NOaXRwUzAvRXdxSFFrU0t3VUtBZDN1VmJN
N3FFbmtpbGJrTXFsdnkrK21KTGI4dHo0L2dNUi9Ea1RIUzR0DQojIE8wUWtkRUpNSERTeGphbkhoQnBUUkkzTnMremdPaVd0MWN0THdPZUtyQ0hLY2p6WjhN
VWtpSUsxVjRybTVFRisNCiMgbGhMcDlzTXhIVEQ2ZXlVN1hBR0FQSDJpV3F3cXhCc0JEbnNyWkpqTHRUTDA3VW55bE1SZA0KIyBTSUcgIyBFbmQgc2lnbmF0
dXJlIGJsb2NrDQo=
'@ -replace '\s',''
$script:ApEmbeddedVersionJson = @'
{
  "_comment": "Third-party scripts vendored byte-exact so their Authenticode signatures remain valid. Do not reformat or re-save these files. Regenerate this manifest with build.ps1 -UpdateVendorManifest.",
  "source": {
    "repository": "https://github.com/andrew-s-taylor/WindowsAutopilotInfo",
    "commit": "a0d2a1c7c1882c38576391b19157ac64b6dc7576",
    "path": "Community Version",
    "retrieved": "2026-07-26",
    "license": "MIT",
    "author": "Andrew S Taylor"
  },
  "scripts": [
    {
      "file": "get-windowsautopilotinfocommunity.ps1",
      "version": "5.0.16",
      "sha256": "EED8FA85F0DC06DFC0476F701E7FE9EAF1744FCE9B71F50E8D5419CAA6B20227",
      "bytes": 125470,
      "signer": "CN=ANDREWSTAYLOR.COM LTD, O=ANDREWSTAYLOR.COM LTD, L=Whitley Bay, C=GB",
      "role": "engine",
      "psGalleryName": "Get-WindowsAutopilotInfoCommunity"
    },
    {
      "file": "Get-AutopilotDiagnosticsCommunity.ps1",
      "version": "6.3",
      "sha256": "35E166838057A3E83B5E79D4A78AE30E382AB5601D69EDA8688F4D11E8058DB4",
      "bytes": 78041,
      "signer": "CN=ANDREWSTAYLOR.COM LTD, O=ANDREWSTAYLOR.COM LTD, L=Whitley Bay, C=GB",
      "role": "diagnostics",
      "psGalleryName": "Get-AutopilotDiagnosticsCommunity"
    }
  ]
}

'@

#endregion embedded resources

#region src\Private\Logging.ps1
# Logging.ps1 -- session log file + in-memory ring buffer for the Logs page.
#
# The original AutoPilot_Import_GUI wrote to the root of C:\ (Get-WindowsAutopilotImportGUI.ps1:520),
# which fails on locked-down machines and litters the system drive. We use
# %ProgramData%\AutopilotImportGUI\Logs and fall back to %TEMP% if that is not writable
# (which happens in some OOBE contexts before ProgramData ACLs settle).

$script:ApLogPath = $null
$script:ApLogBuffer = New-Object System.Collections.Generic.List[string]

function Get-ApLogDirectory {
    <#
    .SYNOPSIS
    Returns a writable directory for logs, creating it if needed.
    #>
    [CmdletBinding()]
    param()

    $candidates = @(
        (Join-Path $env:ProgramData 'AutopilotImportGUI\Logs'),
        (Join-Path $env:TEMP 'AutopilotImportGUI\Logs')
    )

    foreach ($dir in $candidates) {
        if ([string]::IsNullOrWhiteSpace($dir)) { continue }
        try {
            if (-not (Test-Path -LiteralPath $dir)) {
                New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null
            }
            # Prove it is actually writable rather than trusting the ACL.
            $probe = Join-Path $dir ('.write-test-{0}' -f ([guid]::NewGuid().ToString('N')))
            [System.IO.File]::WriteAllText($probe, 'x')
            Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
            return $dir
        }
        catch {
            continue
        }
    }

    throw 'No writable location found for the log directory.'
}

function Initialize-ApLog {
    <#
    .SYNOPSIS
    Starts a new session log file. Safe to call more than once.
    #>
    [CmdletBinding()]
    param(
        [string]$Directory
    )

    if (-not $Directory) { $Directory = Get-ApLogDirectory }

    $name = 'AutopilotImportGUI-{0}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss')
    $script:ApLogPath = Join-Path $Directory $name

    $header = @(
        '=' * 78
        ' Autopilot Import GUI (Community)'
        ' Mark Orr (@markorr321) - https://orr365.tools'
        '-' * 78
        (' Session started : {0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz'))
        (' Computer        : {0}' -f $env:COMPUTERNAME)
        (' User            : {0}\{1}' -f $env:USERDOMAIN, $env:USERNAME)
        (' PowerShell      : {0}' -f $PSVersionTable.PSVersion)
        (' OS              : {0}' -f (Get-ApOsCaption))
        '=' * 78
    ) -join [Environment]::NewLine

    try {
        Set-Content -LiteralPath $script:ApLogPath -Value $header -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        # Logging must never take the GUI down.
        $script:ApLogPath = $null
    }

    return $script:ApLogPath
}

function Get-ApOsCaption {
    try {
        (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).Caption
    }
    catch {
        [Environment]::OSVersion.VersionString
    }
}

function Get-ApLogPath {
    if (-not $script:ApLogPath) { Initialize-ApLog | Out-Null }
    return $script:ApLogPath
}

function Write-ApLog {
    <#
    .SYNOPSIS
    Appends a timestamped line to the session log and the in-memory buffer.

    .PARAMETER Level
    INFO, WARN, ERROR or DEBUG. Rendered into the file; the GUI colours from it.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [AllowEmptyString()]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO'
    )

    process {
        $line = '{0} [{1,-5}] {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message

        $script:ApLogBuffer.Add($line)
        # Keep the buffer bounded; the file on disk remains complete.
        if ($script:ApLogBuffer.Count -gt 5000) {
            $script:ApLogBuffer.RemoveRange(0, 1000)
        }

        $path = $script:ApLogPath
        if ($path) {
            try {
                Add-Content -LiteralPath $path -Value $line -Encoding UTF8 -ErrorAction Stop
            }
            catch {
                # A transient sharing violation must not break the app.
            }
        }
    }
}

function Get-ApLogBuffer {
    <#
    .SYNOPSIS
    Returns the in-memory log lines, newest last.
    #>
    return $script:ApLogBuffer.ToArray()
}

function Clear-ApLogBuffer {
    $script:ApLogBuffer.Clear()
}
#endregion src\Private\Logging.ps1

#region src\Private\Config.ps1
# Config.ps1 -- persisted user preferences (config.json).
#
# Resolution order for the config file:
#   1. next to the script  -- so a USB stick can carry a preconfigured config.json
#      with the site's group tags already populated
#   2. %ProgramData%\AutopilotImportGUI\config.json
# Reads prefer (1); writes go to whichever is writable, preferring (1).

$script:ApConfig = $null
$script:ApConfigPath = $null

function Get-ApDefaultConfig {
    <#
    .SYNOPSIS
    The shape and defaults of config.json. Anything absent from the file on disk
    falls back to these values.
    #>
    return [ordered]@{
        schemaVersion         = 1

        # Editable ComboBox history for the Group Tag field. The original GUI had a
        # bare TextBox, so techs retyped (and mistyped) the same tags all day.
        groupTagHistory       = @()

        # Last-used registration options, restored on next launch.
        lastMode              = 'v1'          # v1 | v2
        lastGroupTag          = ''
        lastAssignedUser      = ''
        lastComputerName      = ''
        lastAddToGroup        = ''
        waitForAssignment     = $true
        rebootWhenAssigned    = $true

        # Device Preparation (v2) restart-after-import. Off by default: a restart straight out
        # of OOBE is premature unless the device is already in the policy's Entra group.
        rebootAfterV2Import   = $false
        existingDevicePolicy  = 'update'      # update | delete
        confirmBeforeRegister = $true

        # Show the child PowerShell console instead of hiding it. Off by default --
        # the whole point of this rewrite is that output lands in the GUI.
        showConsoleWindow     = $false

        # Run the diagnostics script with -Online, which resolves app and policy GUIDs to
        # display names via Graph. Off by default: the local read is the common case and
        # needs neither the sign-in module nor a browser, which matters in OOBE.
        diagnosticsOnline     = $false

        # $null = use the built-in list from Get-ApDefaultEndpoints.
        connectivityEndpoints = $null

        # Prefer the vendored script over any PSGallery-installed copy.
        preferVendoredScript  = $true
    }
}

function Get-ApConfigCandidatePaths {
    $paths = New-Object System.Collections.Generic.List[string]

    if ($script:ApAppRoot) {
        $paths.Add((Join-Path $script:ApAppRoot 'config.json'))
    }
    if ($env:ProgramData) {
        $paths.Add((Join-Path $env:ProgramData 'AutopilotImportGUI\config.json'))
    }

    return $paths.ToArray()
}

function ConvertTo-ApHashtable {
    <#
    .SYNOPSIS
    PSCustomObject (from ConvertFrom-Json) -> ordered hashtable, recursively.
    Windows PowerShell 5.1 has no ConvertFrom-Json -AsHashtable.
    #>
    param($InputObject)

    if ($null -eq $InputObject) { return $null }

    if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
        $out = [ordered]@{}
        foreach ($p in $InputObject.PSObject.Properties) {
            $out[$p.Name] = ConvertTo-ApHashtable $p.Value
        }
        return $out
    }

    if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
        return @($InputObject | ForEach-Object { ConvertTo-ApHashtable $_ })
    }

    return $InputObject
}

function Import-ApConfig {
    <#
    .SYNOPSIS
    Loads config.json, merged over the defaults. Never throws.
    #>
    [CmdletBinding()]
    param()

    $config = Get-ApDefaultConfig

    foreach ($path in (Get-ApConfigCandidatePaths)) {
        if (-not (Test-Path -LiteralPath $path)) { continue }
        try {
            $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
            if ([string]::IsNullOrWhiteSpace($raw)) { continue }

            $loaded = ConvertTo-ApHashtable (ConvertFrom-Json $raw -ErrorAction Stop)
            foreach ($key in $loaded.Keys) {
                # Only accept keys we know about, so a stale file cannot inject junk.
                if ($config.Contains($key)) { $config[$key] = $loaded[$key] }
            }

            $script:ApConfigPath = $path
            Write-ApLog "Loaded configuration from $path"
            break
        }
        catch {
            Write-ApLog "Ignoring unreadable config at ${path}: $($_.Exception.Message)" -Level WARN
        }
    }

    if (-not $script:ApConfigPath) {
        Write-ApLog 'No config.json found; using defaults.'
    }

    # An empty PowerShell array serialises to "{}" rather than "[]" inside an ordered
    # dictionary, and ConvertFrom-Json turns that back into a dictionary. Left alone it
    # became a single bogus "System.Collections.Specialized.OrderedDictionary" entry in the
    # group tag dropdown, so normalise to a clean string array.
    $config.groupTagHistory = @(
        @($config.groupTagHistory) |
            Where-Object { $_ -is [string] -and -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { $_.Trim() } |
            Select-Object -Unique
    )

    $script:ApConfig = $config
    return $config
}

function Get-ApConfig {
    if (-not $script:ApConfig) { Import-ApConfig | Out-Null }
    return $script:ApConfig
}

function Set-ApConfigValue {
    <#
    .SYNOPSIS
    Sets one key in the in-memory config. Call Save-ApConfig to persist.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][AllowNull()][AllowEmptyString()][AllowEmptyCollection()]$Value
    )

    $config = Get-ApConfig
    if (-not $config.Contains($Name)) {
        throw "Unknown configuration key '$Name'."
    }
    $config[$Name] = $Value
}

function Add-ApGroupTagToHistory {
    <#
    .SYNOPSIS
    Records a group tag as most-recently-used, de-duplicated case-insensitively.
    #>
    param([string]$GroupTag)

    if ([string]::IsNullOrWhiteSpace($GroupTag)) { return }

    $tag = $GroupTag.Trim()
    $config = Get-ApConfig
    $existing = @($config.groupTagHistory | Where-Object { $_ -and $_ -ne '' })

    $kept = @($existing | Where-Object { $_ -ne $tag })
    $config.groupTagHistory = @(@($tag) + $kept | Select-Object -First 25)
}

function Save-ApConfig {
    <#
    .SYNOPSIS
    Writes the in-memory config to the first writable candidate path.
    Returns the path written, or $null on failure.
    #>
    [CmdletBinding()]
    param()

    $config = Get-ApConfig
    $json = $config | ConvertTo-Json -Depth 6

    $targets = @()
    if ($script:ApConfigPath) { $targets += $script:ApConfigPath }
    $targets += (Get-ApConfigCandidatePaths)

    foreach ($path in ($targets | Select-Object -Unique)) {
        try {
            $dir = Split-Path -Parent $path
            if ($dir -and -not (Test-Path -LiteralPath $dir)) {
                New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null
            }
            Set-Content -LiteralPath $path -Value $json -Encoding UTF8 -ErrorAction Stop
            $script:ApConfigPath = $path
            return $path
        }
        catch {
            continue
        }
    }

    Write-ApLog 'Could not save configuration to any candidate location.' -Level WARN
    return $null
}
#endregion src\Private\Config.ps1

#region src\Private\Elevation.ps1
# Elevation.ps1 -- administrator detection and one-shot self-relaunch.
#
# Reading the Autopilot hardware hash requires administrator rights: the
# MDM_DevDetail_Ext01 class in root/cimv2/mdm/dmmap is not readable as a standard
# user (see the community script, get-windowsautopilotinfocommunity.ps1:2121).
#
# Elevating the GUI once means every child process it launches inherits the token,
# so the tech sees a single UAC prompt for the whole session instead of one per
# action -- the original GUI used -Verb RunAs on each button.
#
# In OOBE (Shift+F10) the shell already runs as SYSTEM, so this is a no-op there.

function Test-ApElevated {
    <#
    .SYNOPSIS
    True when the current process holds the local Administrators role.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    try {
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Invoke-ApSelfElevate {
    <#
    .SYNOPSIS
    Relaunches the entry script elevated and reports whether the caller should exit.

    .DESCRIPTION
    Returns $true when a new elevated process was started (the caller must exit).
    Returns $false when already elevated, or when the user declined the UAC prompt
    and chose to continue with reduced functionality.

    .PARAMETER ScriptPath
    Path of the entry script to relaunch.

    .PARAMETER BoundParameters
    The entry script's $PSBoundParameters, forwarded to the elevated instance.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [hashtable]$BoundParameters = @{}
    )

    if (Test-ApElevated) { return $false }

    $argList = New-Object System.Collections.Generic.List[string]
    $argList.Add('-NoProfile')
    $argList.Add('-ExecutionPolicy'); $argList.Add('Bypass')
    $argList.Add('-STA')
    $argList.Add('-File'); $argList.Add('"{0}"' -f $ScriptPath)

    foreach ($key in $BoundParameters.Keys) {
        $value = $BoundParameters[$key]
        if ($value -is [switch]) {
            if ($value.IsPresent) { $argList.Add("-$key") }
        }
        elseif ($null -ne $value -and "$value" -ne '') {
            $argList.Add("-$key")
            $argList.Add('"{0}"' -f ($value -replace '"', '""'))
        }
    }

    try {
        Start-Process -FilePath (Get-ApPowerShellPath) `
                      -ArgumentList $argList.ToArray() `
                      -Verb RunAs -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        # 1223 == ERROR_CANCELLED, i.e. the user dismissed the UAC prompt.
        Write-ApLog "Elevation declined or failed: $($_.Exception.Message)" -Level WARN
        return $false
    }
}

function Get-ApPowerShellPath {
    <#
    .SYNOPSIS
    Full path to the Windows PowerShell 5.1 host.

    .DESCRIPTION
    Always Windows PowerShell, never pwsh: the community script pins
    microsoft.graph.authentication to <= 2.9.1 and relies on Windows PowerShell
    behaviour, and WPF hosting is most reliable there. Uses the native-architecture
    path so a 32-bit host does not get redirected into SysWOW64.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()

    $sysNative = Join-Path $env:WINDIR 'sysnative\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path -LiteralPath $sysNative) { return $sysNative }

    $system32 = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path -LiteralPath $system32) { return $system32 }

    return 'powershell.exe'
}
#endregion src\Private\Elevation.ps1

#region src\Private\DeviceInfo.ps1
# DeviceInfo.ps1 -- local hardware/OS facts for the Device page and readiness checks.
#
# Uses CIM throughout. The original GUI mixed Get-WmiObject (deprecated, removed in
# PowerShell 7) with Get-CimInstance -- see Get-WindowsAutopilotImportGUI.ps1:564.
#
# Every getter is individually guarded: on a machine where one WMI provider is
# broken, the Device page should still render everything else rather than blank out.
#
# Nothing in this file may call Write-ApLog or read configuration. The whole set is copied
# into a bare background runspace by Start-ApDeviceLoad, where those helpers do not exist.
# It has to run off the UI thread because two of these queries are slow: on a machine where
# the caller lacks rights, root/cimv2/security/microsofttpm and root/cimv2/mdm/dmmap are
# inaccessible and each call sits on a ~5 second DCOM timeout before failing. Doing that
# synchronously delayed the window by over ten seconds.

function Get-ApCimSafe {
    param(
        [Parameter(Mandatory)][string]$ClassName,
        [string]$Namespace,
        [string]$Filter
    )

    try {
        $params = @{ ClassName = $ClassName; ErrorAction = 'Stop' }
        if ($Namespace) { $params.Namespace = $Namespace }
        if ($Filter) { $params.Filter = $Filter }
        return Get-CimInstance @params
    }
    catch {
        return $null
    }
}

function Get-ApDeviceIdentifierPart {
    <#
    .SYNOPSIS
    Normalises a manufacturer/model string for the v2 device identifier.

    .DESCRIPTION
    Mirrors the community script exactly (get-windowsautopilotinfocommunity.ps1:2115-2116):
    trim, then strip '.' and ',' -- the comma because it is the field separator in the
    "Manufacturer,Model,Serial" triple, the period because Intune's matching rejects it.
    Keeping this identical matters: a mismatch here means the identifier the GUI previews
    is not the one that gets imported.
    #>
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) { return '' }
    return $Value.Trim().Replace('.', '').Replace(',', '')
}

function Test-ApHardwareHashAvailable {
    <#
    .SYNOPSIS
    True when the Autopilot 4K hardware hash can be read (Autopilot v1 prerequisite).
    #>
    [OutputType([bool])]
    param()

    $devDetail = Get-ApCimSafe -ClassName 'MDM_DevDetail_Ext01' `
                               -Namespace 'root/cimv2/mdm/dmmap' `
                               -Filter "InstanceID='Ext' AND ParentID='./DevDetail'"

    return [bool]($devDetail -and $devDetail.DeviceHardwareData)
}

function Get-ApHardwareHash {
    <#
    .SYNOPSIS
    Returns the base64 4K hardware hash, or $null when unavailable.
    #>
    [OutputType([string])]
    param()

    $devDetail = Get-ApCimSafe -ClassName 'MDM_DevDetail_Ext01' `
                               -Namespace 'root/cimv2/mdm/dmmap' `
                               -Filter "InstanceID='Ext' AND ParentID='./DevDetail'"

    if ($devDetail -and $devDetail.DeviceHardwareData) { return $devDetail.DeviceHardwareData }
    return $null
}

function Get-ApTpmInfo {
    <#
    .SYNOPSIS
    TPM presence, enablement and spec version.
    #>
    $tpm = Get-ApCimSafe -ClassName 'Win32_Tpm' -Namespace 'root/cimv2/security/microsofttpm'

    if (-not $tpm) {
        return [ordered]@{ Present = $false; Enabled = $false; Ready = $false; SpecVersion = 'Not detected' }
    }

    # SpecVersion looks like "2.0, 0, 1.38"; the major version is all we display.
    $spec = "$($tpm.SpecVersion)"
    $major = if ($spec -match '^\s*(\d+\.\d+)') { $Matches[1] } else { $spec }

    return [ordered]@{
        # A Win32_Tpm instance only exists when a TPM is physically present.
        Present     = $true
        Enabled     = [bool]$tpm.IsEnabled_InitialValue
        Ready       = ([bool]$tpm.IsEnabled_InitialValue -and [bool]$tpm.IsActivated_InitialValue)
        SpecVersion = $major
    }
}

function Test-ApSecureBoot {
    <#
    .SYNOPSIS
    $true/$false when determinable, $null on legacy BIOS or when the cmdlet is absent.
    #>
    try {
        if (-not (Get-Command Confirm-SecureBootUEFI -ErrorAction SilentlyContinue)) { return $null }
        return [bool](Confirm-SecureBootUEFI -ErrorAction Stop)
    }
    catch {
        # Throws "Cmdlet not supported on this platform" on legacy/BIOS machines.
        return $null
    }
}

function Test-ApInternetConnection {
    <#
    .SYNOPSIS
    Fast reachability probe for the header indicator.

    .DESCRIPTION
    A TCP connect to login.microsoftonline.com:443 -- what actually matters for
    Autopilot -- rather than the original's ICMP ping to 8.8.8.8
    (Get-WindowsAutopilotImportGUI.ps1:630), which fails on any network that blocks
    ICMP or DNS to public resolvers while Microsoft endpoints work fine.
    #>
    [OutputType([bool])]
    param(
        [string]$ComputerName = 'login.microsoftonline.com',
        [int]$Port = 443,
        [int]$TimeoutMs = 3000
    )

    $client = $null
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $async = $client.BeginConnect($ComputerName, $Port, $null, $null)
        if (-not $async.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) { return $false }
        $client.EndConnect($async)
        return $true
    }
    catch {
        return $false
    }
    finally {
        if ($client) { $client.Close() }
    }
}

function Get-ApDeviceInfo {
    <#
    .SYNOPSIS
    Snapshot of everything the GUI shows about the local machine.
    #>
    [CmdletBinding()]
    param()

    $cs   = Get-ApCimSafe -ClassName 'Win32_ComputerSystem'
    $bios = Get-ApCimSafe -ClassName 'Win32_BIOS'
    $os   = Get-ApCimSafe -ClassName 'Win32_OperatingSystem'
    $tpm  = Get-ApTpmInfo

    $serial       = if ($bios) { "$($bios.SerialNumber)".Trim() } else { '' }
    $manufacturer = if ($cs) { "$($cs.Manufacturer)".Trim() } else { '' }
    $model        = if ($cs) { "$($cs.Model)".Trim() } else { '' }

    # System drive free space is what matters for an Autopilot deployment; the
    # original summed every logical disk, which overstated it on multi-disk machines.
    $sysDrive = $env:SystemDrive
    if (-not $sysDrive) { $sysDrive = 'C:' }
    $disk = Get-ApCimSafe -ClassName 'Win32_LogicalDisk' -Filter "DeviceID='$sysDrive'"

    $freeGb  = if ($disk -and $disk.FreeSpace) { [math]::Round($disk.FreeSpace / 1GB, 1) } else { $null }
    $totalGb = if ($disk -and $disk.Size) { [math]::Round($disk.Size / 1GB, 1) } else { $null }
    $ramGb   = if ($cs -and $cs.TotalPhysicalMemory) { [math]::Round($cs.TotalPhysicalMemory / 1GB, 1) } else { $null }

    $hashAvailable = Test-ApHardwareHashAvailable

    $idManufacturer = Get-ApDeviceIdentifierPart $manufacturer
    $idModel        = Get-ApDeviceIdentifierPart $model

    return [ordered]@{
        SerialNumber      = $serial
        Manufacturer      = $manufacturer
        Model             = $model
        ComputerName      = if ($cs) { "$($cs.Name)".Trim() } else { $env:COMPUTERNAME }
        OsCaption         = if ($os) { "$($os.Caption)".Trim() } else { '' }
        OsBuild           = if ($os) { "$($os.Version)" } else { '' }
        OsDisplayVersion  = Get-ApWindowsDisplayVersion
        SystemDrive       = $sysDrive
        FreeSpaceGb       = $freeGb
        TotalSpaceGb      = $totalGb
        MemoryGb          = $ramGb
        TpmPresent        = $tpm.Present
        TpmEnabled        = $tpm.Enabled
        TpmSpecVersion    = $tpm.SpecVersion
        SecureBoot        = Test-ApSecureBoot
        HardwareHashReady = $hashAvailable
        IsElevated        = Test-ApElevated
        IsVirtualMachine  = Test-ApVirtualMachine $manufacturer $model

        # "Manufacturer,Model,Serial" -- exactly what -identifier imports for Autopilot v2.
        DeviceIdentifier  = '{0},{1},{2}' -f $idManufacturer, $idModel, $serial
    }
}

function Get-ApWindowsDisplayVersion {
    try {
        $key = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
        $props = Get-ItemProperty -Path $key -ErrorAction Stop
        if ($props.DisplayVersion) { return $props.DisplayVersion }
        if ($props.ReleaseId) { return $props.ReleaseId }
        return ''
    }
    catch {
        return ''
    }
}

function Test-ApVirtualMachine {
    param([string]$Manufacturer, [string]$Model)

    $needles = @('Virtual', 'VMware', 'Hyper-V', 'KVM', 'QEMU', 'Xen', 'VirtualBox', 'Parallels')
    $haystack = "$Manufacturer $Model"

    foreach ($n in $needles) {
        if ($haystack -like "*$n*") { return $true }
    }
    return $false
}
#endregion src\Private\DeviceInfo.ps1

#region src\Private\VendorScript.ps1
# VendorScript.ps1 -- locate the community engine script that the GUI drives.
#
# Resolution order (mirrors VM-Pilot's AutopilotV2Import.ps1:44-57, which pre-injects
# the script so imports still work when PSGallery is unreachable -- the normal case
# on a locked-down OOBE network):
#
#   1. embedded base64 payload  (single-file dist build)
#   2. vendor\ next to the source tree  (dev / folder deployment)
#   3. an existing PSGallery install    (Get-InstalledScript)
#   4. Install-Script from PSGallery    (last resort, needs internet)
#
# The vendored copy is byte-exact so Andrew Taylor's Authenticode signature survives;
# Test-ApVendorScript verifies both the SHA256 from VERSION.json and the signature.

$script:ApVendorCache = @{}

# Populated by build.ps1 in the single-file dist. Empty in the source tree.
if (-not (Test-Path Variable:script:ApEmbeddedScripts)) {
    $script:ApEmbeddedScripts = @{}
}

function Get-ApVendorManifest {
    <#
    .SYNOPSIS
    Reads vendor\VERSION.json, or the embedded copy in the dist build.
    #>
    if ($script:ApVendorManifestCache) { return $script:ApVendorManifestCache }

    $manifest = $null

    if ($script:ApEmbeddedVersionJson) {
        try { $manifest = ConvertFrom-Json $script:ApEmbeddedVersionJson -ErrorAction Stop } catch { }
    }

    if (-not $manifest -and $script:ApAppRoot) {
        $path = Join-Path $script:ApAppRoot 'vendor\VERSION.json'
        if (Test-Path -LiteralPath $path) {
            try { $manifest = ConvertFrom-Json (Get-Content -LiteralPath $path -Raw) -ErrorAction Stop } catch { }
        }
    }

    $script:ApVendorManifestCache = $manifest
    return $manifest
}

function Get-ApVendorExpectedHash {
    param([Parameter(Mandatory)][string]$FileName)

    $manifest = Get-ApVendorManifest
    if (-not $manifest) { return $null }

    $entry = $manifest.scripts | Where-Object { $_.file -eq $FileName } | Select-Object -First 1
    if ($entry) { return $entry.sha256 }
    return $null
}

function Test-ApVendorScript {
    <#
    .SYNOPSIS
    Verifies a resolved engine script against VERSION.json and its Authenticode signature.

    .DESCRIPTION
    Returns a result object rather than throwing. A checksum mismatch is fatal (the file
    is not what we shipped); an invalid or absent signature is only a warning, because a
    PSGallery-installed copy may legitimately be a newer, differently-signed version.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$SkipHashCheck
    )

    $result = [ordered]@{
        Path            = $Path
        Exists          = (Test-Path -LiteralPath $Path -PathType Leaf)
        HashMatches     = $null
        ActualHash      = $null
        ExpectedHash    = $null
        SignatureStatus = 'NotChecked'
        Signer          = $null
        IsTrusted       = $false
        Messages        = @()
    }

    if (-not $result.Exists) {
        $result.Messages += "Engine script not found at $Path"
        return [pscustomobject]$result
    }

    $fileName = Split-Path -Leaf $Path
    $expected = Get-ApVendorExpectedHash -FileName $fileName
    $result.ExpectedHash = $expected

    if (-not $SkipHashCheck -and $expected) {
        try {
            $actual = (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash
            $result.ActualHash = $actual
            $result.HashMatches = ($actual -eq $expected)
            if (-not $result.HashMatches) {
                $result.Messages += "Checksum mismatch for $fileName (expected $expected, got $actual)"
            }
        }
        catch {
            $result.Messages += "Could not hash ${fileName}: $($_.Exception.Message)"
        }
    }

    try {
        $sig = Get-AuthenticodeSignature -LiteralPath $Path -ErrorAction Stop
        $result.SignatureStatus = "$($sig.Status)"
        if ($sig.SignerCertificate) { $result.Signer = $sig.SignerCertificate.Subject }
        if ($sig.Status -ne 'Valid') {
            $result.Messages += "Authenticode signature is '$($sig.Status)' for $fileName"
        }
    }
    catch {
        $result.Messages += "Could not read signature for ${fileName}: $($_.Exception.Message)"
    }

    # Trusted enough to run: correct bytes, or no manifest entry to compare against.
    $result.IsTrusted = ($result.HashMatches -ne $false)

    return [pscustomobject]$result
}

function Expand-ApEmbeddedScript {
    <#
    .SYNOPSIS
    Writes an embedded base64 payload to disk byte-for-byte.

    .DESCRIPTION
    Uses WriteAllBytes, never Set-Content, so no encoding or line-ending translation
    happens -- that would break the Authenticode signature.
    #>
    param(
        [Parameter(Mandatory)][string]$FileName,
        [Parameter(Mandatory)][string]$Destination
    )

    if (-not $script:ApEmbeddedScripts.ContainsKey($FileName)) { return $null }

    $dir = Split-Path -Parent $Destination
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $bytes = [Convert]::FromBase64String($script:ApEmbeddedScripts[$FileName])
    [System.IO.File]::WriteAllBytes($Destination, $bytes)
    return $Destination
}

function Get-ApWorkingDirectory {
    <#
    .SYNOPSIS
    Per-session scratch directory for extracted scripts, launchers and run logs.
    #>
    if ($script:ApWorkingDirectory -and (Test-Path -LiteralPath $script:ApWorkingDirectory)) {
        return $script:ApWorkingDirectory
    }

    $base = Join-Path $env:TEMP ('AutopilotImportGUI\{0}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    New-Item -ItemType Directory -Path $base -Force | Out-Null
    $script:ApWorkingDirectory = $base
    return $base
}

function Resolve-ApEngineScript {
    <#
    .SYNOPSIS
    Returns the full path to a community script, acquiring it if necessary.

    .PARAMETER Name
    'Engine' for get-windowsautopilotinfocommunity.ps1, 'Diagnostics' for
    Get-AutopilotDiagnosticsCommunity.ps1.

    .PARAMETER AllowInstall
    Permit the PSGallery fallback (needs internet). On by default.
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('Engine', 'Diagnostics')]
        [string]$Name = 'Engine',
        [switch]$AllowInstall = $true
    )

    if ($script:ApVendorCache.ContainsKey($Name)) { return $script:ApVendorCache[$Name] }

    $fileName = if ($Name -eq 'Engine') {
        'get-windowsautopilotinfocommunity.ps1'
    } else {
        'Get-AutopilotDiagnosticsCommunity.ps1'
    }
    $galleryName = if ($Name -eq 'Engine') {
        'Get-WindowsAutopilotInfoCommunity'
    } else {
        'Get-AutopilotDiagnosticsCommunity'
    }

    # 1. embedded payload (single-file build)
    if ($script:ApEmbeddedScripts.ContainsKey($fileName)) {
        $dest = Join-Path (Get-ApWorkingDirectory) $fileName
        if (-not (Test-Path -LiteralPath $dest)) {
            Expand-ApEmbeddedScript -FileName $fileName -Destination $dest | Out-Null
            Write-ApLog "Extracted embedded $fileName to $dest"
        }
        $script:ApVendorCache[$Name] = $dest
        return $dest
    }

    # 2. vendor folder in the source tree
    if ($script:ApAppRoot) {
        $vendored = Join-Path $script:ApAppRoot ('vendor\{0}' -f $fileName)
        if (Test-Path -LiteralPath $vendored -PathType Leaf) {
            Write-ApLog "Using vendored $fileName"
            $script:ApVendorCache[$Name] = $vendored
            return $vendored
        }
    }

    # 3. already installed from PSGallery
    try {
        $installed = Get-InstalledScript -Name $galleryName -ErrorAction Stop
        if ($installed) {
            $candidate = Join-Path $installed.InstalledLocation "$galleryName.ps1"
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                Write-ApLog "Using PSGallery-installed $galleryName $($installed.Version)"
                $script:ApVendorCache[$Name] = $candidate
                return $candidate
            }
        }
    }
    catch {
        # Not installed; fall through.
    }

    # 4. install from PSGallery
    if ($AllowInstall) {
        try {
            Write-ApLog "Installing $galleryName from the PowerShell Gallery..."
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
            Install-Script -Name $galleryName -Force -Scope AllUsers -Confirm:$false -ErrorAction Stop

            $installed = Get-InstalledScript -Name $galleryName -ErrorAction Stop
            $candidate = Join-Path $installed.InstalledLocation "$galleryName.ps1"
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                $script:ApVendorCache[$Name] = $candidate
                return $candidate
            }
        }
        catch {
            Write-ApLog "Could not install ${galleryName}: $($_.Exception.Message)" -Level ERROR
        }
    }

    throw "Unable to locate $fileName. Place it in the vendor folder next to this script, or connect to the internet so it can be installed from the PowerShell Gallery."
}
#endregion src\Private\VendorScript.ps1

#region src\Private\ArgumentBuilder.ps1
# ArgumentBuilder.ps1 -- turns UI state into parameters for the community engine script.
#
# This is deliberately a pure function with no UI or Graph dependencies: it is the piece
# most likely to be wrong in a way that silently does nothing, so it is unit-tested
# (tests\ArgumentBuilder.Tests.ps1) and it also drives the "Preview command" dry-run.
#
# Three non-obvious constraints from get-windowsautopilotinfocommunity.ps1 v5.0.16 are
# encoded here. Getting any of them wrong produces a run that appears to succeed but
# doesn't do what the tech asked:
#
#   1. -Reboot / -Wipe / -Sysprep / -preprov / -ChangePK are all nested inside
#      `if ($Assign)` (line 2541). Without -Assign they are silently ignored.
#   2. With neither -delete nor -updatetag, an already-registered serial hits
#      `Read-Host "Do you want to delete or update?"` (line 2346) and blocks forever
#      behind a hidden console. -Force resolves it to "update" (line 2342).
#   3. -identifier takes a completely separate code path (lines 2233-2270) that ignores
#      GroupTag, AssignedUser, AssignedComputerName, AddToGroup, Assign and Reboot.

function New-ApRegistrationRequest {
    <#
    .SYNOPSIS
    Default UI state for a registration request.

    .DESCRIPTION
    Returned as an ordered hashtable so callers can tweak fields before handing it to
    Build-ApEngineArguments. Keeping the default shape in one place stops the UI and the
    tests from drifting apart.
    #>
    [CmdletBinding()]
    param()

    return [ordered]@{
        # Register = push to tenant; Export = offline CSV only.
        Operation            = 'Register'

        # v1 = hardware hash (windowsAutopilotDeviceIdentities)
        # v2 = device identifier for Device Preparation (importedDeviceIdentities)
        Mode                 = 'v1'

        GroupTag             = ''
        AssignedUser         = ''
        AssignedComputerName = ''
        AddToGroup           = ''

        WaitForAssignment    = $false
        Reboot               = $false

        # Device Preparation (v2) restart. Deliberately NOT an engine switch: the engine's
        # -Reboot is nested inside its assignment wait, which the -identifier path never
        # reaches, so the GUI performs this restart itself once the import succeeds.
        RebootAfterImport    = $false

        # update | delete | skipcheck
        ExistingDevicePolicy = 'update'

        Wipe                 = $false
        Sysprep              = $false
        PreProvision         = $false
        ChangePK             = ''

        OutputFile           = ''
        Append               = $false
        Partner              = $false
    }
}

function Build-ApEngineArguments {
    <#
    .SYNOPSIS
    Maps a registration request to community-script parameters.

    .OUTPUTS
    PSCustomObject with:
      Parameters  ordered hashtable, ready to splat
      Notes       human-readable explanations of any coercion applied
      Warnings    settings that were dropped because the chosen mode ignores them
      Mode/Operation  echoed back for the caller

    .EXAMPLE
    $req = New-ApRegistrationRequest
    $req.GroupTag = 'FINANCE'
    $req.Reboot = $true
    (Build-ApEngineArguments $req).Parameters
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        $Request
    )

    # Accept either a hashtable or a PSCustomObject from the caller.
    $r = if ($Request -is [System.Collections.IDictionary]) { $Request } else { ConvertTo-ApHashtable $Request }

    $p = [ordered]@{}
    $notes = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]

    $operation = Get-ApRequestValue $r 'Operation' 'Register'
    $mode      = Get-ApRequestValue $r 'Mode' 'v1'
    $isV2      = ($mode -eq 'v2')

    # 'Batch' was removed in 1.3.0. The engine only reads -InputFile on its -identifier path
    # (get-windowsautopilotinfocommunity.ps1:2234), and its Process block skips device
    # collection entirely when -InputFile is set, so a v1 batch import collected nothing and
    # reported success having imported zero devices. Fail loudly rather than build that again.
    if ($operation -notin @('Register', 'Export')) {
        throw "'$operation' is not a supported operation. Use Register or Export."
    }

    $groupTag     = (Get-ApRequestValue $r 'GroupTag' '').Trim()
    $assignedUser = (Get-ApRequestValue $r 'AssignedUser' '').Trim()
    $computerName = (Get-ApRequestValue $r 'AssignedComputerName' '').Trim()
    $addToGroup   = (Get-ApRequestValue $r 'AddToGroup' '').Trim()
    $changePk     = (Get-ApRequestValue $r 'ChangePK' '').Trim()
    $outputFile   = (Get-ApRequestValue $r 'OutputFile' '').Trim()

    $wait   = [bool](Get-ApRequestValue $r 'WaitForAssignment' $false)
    $reboot = [bool](Get-ApRequestValue $r 'Reboot' $false)
    $wipe   = [bool](Get-ApRequestValue $r 'Wipe' $false)
    $syspr  = [bool](Get-ApRequestValue $r 'Sysprep' $false)
    $prepro = [bool](Get-ApRequestValue $r 'PreProvision' $false)
    $append = [bool](Get-ApRequestValue $r 'Append' $false)
    $partner = [bool](Get-ApRequestValue $r 'Partner' $false)
    $policy = Get-ApRequestValue $r 'ExistingDevicePolicy' 'update'

    if ($isV2) { $p['identifier'] = $true }

    # ---- offline export -------------------------------------------------------
    if ($operation -eq 'Export') {
        if (-not $outputFile) { throw 'An output file is required for an offline export.' }

        $p['OutputFile'] = $outputFile
        if ($append) { $p['Append'] = $true }

        if ($isV2) {
            # The -identifier branch writes Manufacturer,Model,Serial and ignores the rest.
            if ($groupTag) { $warnings.Add('Group tag is not used by a Device Preparation (v2) identifier export.') }
            if ($assignedUser) { $warnings.Add('Assigned user is not used by a Device Preparation (v2) identifier export.') }
        }
        else {
            if ($partner) {
                $p['Partner'] = $true
                # Line 2215 selects the Partner column set before the GroupTag one,
                # so a group tag supplied alongside -Partner never reaches the CSV.
                if ($groupTag) { $warnings.Add('Partner CSV format does not include a Group Tag column; the tag will be omitted.') }
                if ($assignedUser) { $warnings.Add('Partner CSV format does not include an Assigned User column; the user will be omitted.') }
            }
            else {
                if ($groupTag) { $p['GroupTag'] = $groupTag }
                if ($assignedUser) { $p['AssignedUser'] = $assignedUser }
            }
        }

        return [pscustomobject]@{
            Parameters = $p
            Notes      = $notes.ToArray()
            Warnings   = $warnings.ToArray()
            Mode       = $mode
            Operation  = $operation
        }
    }

    # ---- online: Register ----------------------------------------------------
    $p['Online'] = $true

    if ($isV2) {
        # Constraint 3: the -identifier path ignores all of these. Warn rather than
        # emit them, so the tech is told instead of quietly getting a different result.
        if ($groupTag)     { $warnings.Add('Group tag is ignored in Device Preparation (v2) mode. v2 assigns devices via the Entra group on the Device Preparation policy.') }
        if ($assignedUser) { $warnings.Add('Assigned user is ignored in Device Preparation (v2) mode.') }
        if ($computerName) { $warnings.Add('Computer name is ignored in Device Preparation (v2) mode.') }
        if ($addToGroup)   { $warnings.Add('Add-to-group is ignored in Device Preparation (v2) mode; add the device to the policy''s Entra group afterwards.') }
        if ($wait)         { $warnings.Add('Profile-assignment wait does not apply in Device Preparation (v2) mode.') }
        if ($reboot)       { $warnings.Add('Reboot-when-assigned does not apply in Device Preparation (v2) mode.') }
        if ($wipe -or $syspr -or $prepro -or $changePk) {
            $warnings.Add('Post-registration actions (wipe / sysprep / pre-provision / product key) are only available in Autopilot v1 mode.')
        }

        return [pscustomobject]@{
            Parameters = $p
            Notes      = $notes.ToArray()
            Warnings   = $warnings.ToArray()
            Mode       = $mode
            Operation  = $operation
        }
    }

    # ---- v1 online -----------------------------------------------------------
    if ($groupTag)     { $p['GroupTag'] = $groupTag }
    if ($assignedUser) { $p['AssignedUser'] = $assignedUser }
    if ($computerName) { $p['AssignedComputerName'] = $computerName }

    if ($addToGroup) {
        # -AddToGroup is [String[]]; accept a comma- or semicolon-separated UI field.
        $groups = @($addToGroup -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        if ($groups.Count -gt 0) { $p['AddToGroup'] = $groups }
    }

    # Constraint 2: never leave the existing-device decision to a Read-Host prompt.
    switch ($policy) {
        'delete' {
            $p['delete'] = $true
            $notes.Add('An existing registration for this serial will be deleted from Autopilot, Intune and Entra ID, then re-added.')
        }
        'skipcheck' {
            $p['newdevice'] = $true
            $notes.Add('Skipping the existing-device lookup. Much faster on large tenants, but will fail if the serial is already registered.')
        }
        default {
            $p['updatetag'] = $true
            $notes.Add('An existing registration for this serial will have its group tag updated.')
        }
    }
    # Belt and braces: -Force is only read in the prompt branch (line 2342), so it is
    # harmless alongside -updatetag/-delete and guarantees the run never blocks.
    $p['Force'] = $true

    # Constraint 1: these only execute inside `if ($Assign)`.
    $needsAssign = $reboot -or $wipe -or $syspr -or $prepro -or [bool]$changePk

    if ($wait -or $needsAssign) {
        $p['Assign'] = $true
        if ($needsAssign -and -not $wait) {
            $notes.Add('Waiting for profile assignment was enabled automatically: the community script only performs post-registration actions after assignment completes.')
        }
    }

    if ($reboot) { $p['Reboot'] = $true }
    if ($wipe)   { $p['Wipe'] = $true }
    if ($syspr)  { $p['Sysprep'] = $true }
    if ($prepro) { $p['preprov'] = $true }
    if ($changePk) { $p['ChangePK'] = $changePk }

    if ($reboot -and ($wipe -or $syspr)) {
        $warnings.Add('Reboot runs before wipe and sysprep in the community script, so those actions will not be reached. Choose one.')
    }

    return [pscustomobject]@{
        Parameters = $p
        Notes      = $notes.ToArray()
        Warnings   = $warnings.ToArray()
        Mode       = $mode
        Operation  = $operation
    }
}

function Get-ApRequestValue {
    <#
    .SYNOPSIS
    Reads a key from the request with a default, tolerating hashtable or object input.
    #>
    param(
        [Parameter(Mandatory)]$Request,
        [Parameter(Mandatory)][string]$Name,
        $Default
    )

    if ($Request -is [System.Collections.IDictionary]) {
        if ($Request.Contains($Name) -and $null -ne $Request[$Name]) { return $Request[$Name] }
        return $Default
    }

    $prop = $Request.PSObject.Properties[$Name]
    if ($prop -and $null -ne $prop.Value) { return $prop.Value }
    return $Default
}

function ConvertTo-ApArgumentArray {
    <#
    .SYNOPSIS
    Parameter hashtable -> flat argument array for powershell.exe -File.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Parameters
    )

    $argList = New-Object System.Collections.Generic.List[string]

    foreach ($key in $Parameters.Keys) {
        $value = $Parameters[$key]

        if ($value -is [bool] -or $value -is [switch]) {
            if ($value) { $argList.Add("-$key") }
            continue
        }

        $argList.Add("-$key")
        if ($value -is [Array]) {
            # A [String[]] parameter takes comma-separated values on a command line.
            $argList.Add(($value -join ','))
        }
        else {
            $argList.Add("$value")
        }
    }

    return $argList.ToArray()
}

function Get-ApPreviewCommand {
    <#
    .SYNOPSIS
    Renders the exact invocation for the "Preview command" dry-run.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Parameters,
        [string]$ScriptPath = 'get-windowsautopilotinfocommunity.ps1'
    )

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append(('& "{0}"' -f $ScriptPath))

    foreach ($key in $Parameters.Keys) {
        $value = $Parameters[$key]

        if ($value -is [bool] -or $value -is [switch]) {
            if ($value) { [void]$sb.Append(" -$key") }
            continue
        }

        if ($value -is [Array]) {
            $quoted = @($value | ForEach-Object { "'{0}'" -f ($_ -replace "'", "''") })
            [void]$sb.Append((' -{0} {1}' -f $key, ($quoted -join ',')))
        }
        else {
            [void]$sb.Append((" -{0} '{1}'" -f $key, ("$value" -replace "'", "''")))
        }
    }

    return $sb.ToString()
}
#endregion src\Private\ArgumentBuilder.ps1

#region src\Private\ProgressParser.ps1
# ProgressParser.ps1 -- turns community-script console output into UI progress state.
#
# Wrapping a console script means the only progress signal available is its stdout.
# Fortunately get-windowsautopilotinfocommunity.ps1 emits stable, greppable lines, so we
# can drive a real staged progress bar instead of the indeterminate spinner the original
# GUI never even had. Line references are to v5.0.16.
#
# Stages, in the order the script performs them:
#   Connect  -> Collect -> Import -> Sync -> Assign -> Complete
#
# Pure function, unit-tested in tests\ProgressParser.Tests.ps1. If upstream changes its
# wording the parser degrades to "no progress update" rather than breaking the run --
# Update-ApProgressState returns $null for unrecognised lines and the caller keeps the
# previous state.

$script:ApStageOrder = @('Connect', 'Collect', 'Import', 'Sync', 'Assign', 'Complete')

function New-ApProgressState {
    <#
    .SYNOPSIS
    Initial progress state for a run.
    #>
    [CmdletBinding()]
    param()

    return [ordered]@{
        Stage        = 'Connect'
        StageLabel   = 'Signing in...'
        Current      = 0
        Total        = 0
        Percent      = 0
        IsComplete   = $false
        IsError      = $false
        ErrorMessage = ''
        LastMessage  = ''
    }
}

function Get-ApStageWeight {
    <#
    .SYNOPSIS
    Percentage floor for each stage, so the bar advances monotonically across stages.
    #>
    param([string]$Stage)

    switch ($Stage) {
        'Connect'  { return @{ Floor = 0;   Span = 10 } }
        'Collect'  { return @{ Floor = 10;  Span = 10 } }
        'Import'   { return @{ Floor = 20;  Span = 35 } }
        'Sync'     { return @{ Floor = 55;  Span = 20 } }
        'Assign'   { return @{ Floor = 75;  Span = 24 } }
        'Complete' { return @{ Floor = 100; Span = 0 } }
        default    { return @{ Floor = 0;   Span = 0 } }
    }
}

function Update-ApProgressState {
    <#
    .SYNOPSIS
    Folds one output line into the progress state.

    .DESCRIPTION
    Returns the updated state when the line was meaningful, otherwise $null so the caller
    knows nothing changed. Mutates and returns the same object for cheapness.

    .PARAMETER State
    State from New-ApProgressState.

    .PARAMETER Line
    A single line of script output.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$State,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Line
    )

    if ([string]::IsNullOrWhiteSpace($Line)) { return $null }

    $text = $Line.Trim()
    $changed = $false

    # ---- "Waiting for N of M to be <imported|synced|assigned>" (2415 / 2452 / 2533) ----
    if ($text -match 'Waiting for\s+(\d+)\s+of\s+(\d+)\s+to be\s+(imported|synced|assigned)') {
        $remaining = [int]$Matches[1]
        $total     = [int]$Matches[2]
        $verb      = $Matches[3]

        $stage = switch ($verb) {
            'imported' { 'Import' }
            'synced'   { 'Sync' }
            'assigned' { 'Assign' }
        }

        # The script reports how many are still outstanding, not how many are done.
        $done = [Math]::Max(0, $total - $remaining)

        $State.Stage      = $stage
        $State.Current    = $done
        $State.Total      = $total
        $State.StageLabel = switch ($stage) {
            'Import' { if ($total -gt 1) { "Importing $done of $total devices" } else { 'Importing device into Autopilot' } }
            'Sync'   { if ($total -gt 1) { "Syncing $done of $total devices" }   else { 'Waiting for Intune sync' } }
            'Assign' { if ($total -gt 1) { "Assigning profile to $done of $total" } else { 'Waiting for deployment profile assignment' } }
        }
        $changed = $true
    }
    # ---- stage completion markers ----
    elseif ($text -match 'devices imported successfully') {
        $State.Stage = 'Sync'
        $State.StageLabel = 'Import complete, waiting for Intune sync'
        $changed = $true
    }
    elseif ($text -match 'All devices synced') {
        $State.Stage = 'Assign'
        $State.StageLabel = 'Sync complete, waiting for profile assignment'
        $changed = $true
    }
    elseif ($text -match 'Profiles assigned to all devices') {
        $State.Stage = 'Complete'
        $State.StageLabel = 'Deployment profile assigned'
        $State.IsComplete = $true
        $changed = $true
    }
    # ---- connect / collect ----
    elseif ($text -match 'Connected to Intune tenant\s+(.+)$') {
        $State.Stage = 'Collect'
        $State.StageLabel = 'Signed in, collecting device details'
        $changed = $true
    }
    elseif ($text -match 'Gathered details for device with serial number:\s*(\S+)') {
        $State.Stage = 'Import'
        $State.StageLabel = "Collected hardware details for $($Matches[1])"
        $changed = $true
    }
    elseif ($text -match 'Loading all objects') {
        $State.Stage = 'Import'
        $State.StageLabel = 'Checking whether this device is already registered'
        $changed = $true
    }
    elseif ($text -match 'Adding New Device serial') {
        $State.Stage = 'Import'
        $State.StageLabel = 'Adding device to Autopilot'
        $changed = $true
    }
    elseif ($text -match 'Updating Existing Device') {
        $State.Stage = 'Import'
        $State.StageLabel = 'Updating the existing Autopilot registration'
        $changed = $true
    }
    elseif ($text -match 'Device already exists in Autopilot') {
        $State.Stage = 'Import'
        $State.StageLabel = 'Device is already registered in Autopilot'
        $changed = $true
    }
    # ---- Autopilot v2 / Device Preparation identifier path (2240-2266) ----
    elseif ($text -match 'Checking if device (\S+) exists in AutoPilot') {
        $State.Stage = 'Import'
        $State.StageLabel = "Checking for an existing identifier for $($Matches[1])"
        $changed = $true
    }
    elseif ($text -match 'Device (\S+) added to AutoPilot') {
        $State.Stage = 'Complete'
        $State.StageLabel = "Device identifier imported for $($Matches[1])"
        $State.IsComplete = $true
        $changed = $true
    }
    elseif ($text -match 'Device (\S+) already exists in AutoPilot') {
        $State.Stage = 'Complete'
        $State.StageLabel = "Device identifier already present for $($Matches[1])"
        $State.IsComplete = $true
        $changed = $true
    }
    # ---- failures ----
    elseif ($text -match 'Unable to retrieve device hardware data') {
        $State.IsError = $true
        $State.ErrorMessage = 'Could not read the hardware hash. Run as administrator, and note that virtual machines often cannot provide one.'
        $changed = $true
    }
    elseif ($text -match 'Unable to find group\s+(.+)$') {
        $State.IsError = $true
        $State.ErrorMessage = "Entra group '$($Matches[1])' was not found."
        $changed = $true
    }
    elseif ($text -match '^\s*(Connect-MgGraph|Invoke-MgGraphRequest)\s*:' -or $text -match 'Authentication needed' -or $text -match 'InteractiveBrowserCredential authentication failed') {
        $State.IsError = $true
        $State.ErrorMessage = 'Sign-in to Microsoft Graph failed. Check network connectivity and try again.'
        $changed = $true
    }
    elseif ($text -match 'command was found in the module\s+''([^'']+)''.+could not be loaded') {
        # PowerShell raises this at command resolution, so it is not suppressible with
        # -ErrorAction and it aborts whatever was running.
        $State.IsError = $true
        $State.ErrorMessage = "PowerShell could not load the '$($Matches[1])' module in the engine process. This usually means PSModulePath is broken for that session."
        $changed = $true
    }
    elseif ($text -match '^ERROR:\s*(.+)$') {
        # Emitted by the launcher itself. Surfacing it means the status bar names the real
        # cause instead of only reporting a non-zero exit code.
        $State.IsError = $true
        $State.ErrorMessage = $Matches[1].Trim()
        $changed = $true
    }

    if (-not $changed) { return $null }

    $State.LastMessage = $text
    $State.Percent = Get-ApProgressPercent -State $State
    return $State
}

function Get-ApProgressPercent {
    <#
    .SYNOPSIS
    Percentage for the progress bar, monotonic across stages.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$State
    )

    if ($State.IsComplete) { return 100 }

    $weight = Get-ApStageWeight $State.Stage
    $percent = $weight.Floor

    if ($State.Total -gt 0 -and $weight.Span -gt 0) {
        $fraction = [Math]::Min(1.0, [double]$State.Current / [double]$State.Total)
        $percent = $weight.Floor + [int]([Math]::Floor($fraction * $weight.Span))
    }

    return [Math]::Max(0, [Math]::Min(100, $percent))
}

function Test-ApStageReached {
    <#
    .SYNOPSIS
    True when $Stage is at or beyond $Reference in the pipeline order.
    #>
    param(
        [Parameter(Mandatory)][string]$Stage,
        [Parameter(Mandatory)][string]$Reference
    )

    $a = [Array]::IndexOf($script:ApStageOrder, $Stage)
    $b = [Array]::IndexOf($script:ApStageOrder, $Reference)
    if ($a -lt 0 -or $b -lt 0) { return $false }
    return ($a -ge $b)
}
#endregion src\Private\ProgressParser.ps1

#region src\Private\ScriptRunner.ps1
# ScriptRunner.ps1 -- runs the community engine script and streams its output into the GUI.
#
# Why a child process and a log file rather than a runspace:
#   * the engine is a *script* with a Begin/Process/End pipeline and its own $ErrorActionPreference,
#     setx calls and Disconnect-MgGraph teardown -- running it in-process would pollute the GUI's
#     session state and leave a Graph connection behind
#   * interactive Connect-MgGraph spawns a browser and can call exit; a crash in the engine must not
#     take the window down with it
#   * a log file on disk is also the artefact the tech needs afterwards for a support case
#
# Why the parameters go into a generated launcher rather than onto the command line:
#   quoting a group tag containing a space or an apostrophe through
#   powershell.exe -File is a known source of silent corruption. The launcher embeds a
#   literal hashtable and splats it, so values reach the engine exactly as typed.
#
# Output capture relies on `*>&1` -- in Windows PowerShell 5.1 Write-Host writes to the
# information stream, so redirecting all streams captures the engine's entire narration
# (it uses Write-Host almost exclusively). Without the `*` the log would be empty.

$script:ApCompleteSentinel = '##AP_COMPLETE:'

function Format-ApPowerShellLiteral {
    <#
    .SYNOPSIS
    Renders a value as PowerShell source. Supports the types the argument builder emits.
    #>
    param($Value)

    if ($null -eq $Value) { return '$null' }

    if ($Value -is [bool] -or $Value -is [switch]) {
        if ($Value) { return '$true' } else { return '$false' }
    }

    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double]) { return "$Value" }

    if ($Value -is [Array]) {
        $items = @($Value | ForEach-Object { Format-ApPowerShellLiteral $_ })
        return '@(' + ($items -join ', ') + ')'
    }

    # Single-quoted string with doubled quotes: no expansion, no escape sequences.
    return "'" + ("$Value" -replace "'", "''") + "'"
}

function Format-ApHashtableLiteral {
    <#
    .SYNOPSIS
    Renders a parameter hashtable as a PowerShell hashtable literal for splatting.
    #>
    param([Parameter(Mandatory)][System.Collections.IDictionary]$Parameters)

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('@{')
    foreach ($key in $Parameters.Keys) {
        $lines.Add(('    {0} = {1}' -f $key, (Format-ApPowerShellLiteral $Parameters[$key])))
    }
    $lines.Add('}')

    return ($lines -join [Environment]::NewLine)
}

function Get-ApDependencyPrepBlock {
    <#
    .SYNOPSIS
    Launcher source that installs the engine's online prerequisites without prompting.

    .DESCRIPTION
    The engine bootstraps its own dependencies, but it does so interactively:

        $provider = Get-PackageProvider NuGet -ErrorAction Ignore
        if (-not $provider) { Find-PackageProvider -Name NuGet -ForceBootstrap -IncludeDependencies }
        $module = Import-Module microsoft.graph.authentication -PassThru -ErrorAction Ignore
        if (-not $module) { Install-Module microsoft.graph.authentication -MaximumVersion 2.9.1 }

    On a machine with no NuGet provider and an Untrusted PSGallery, -ForceBootstrap raises a
    Yes/No console prompt. The engine console is hidden, so nobody can answer it and the run
    blocks forever: the browser never appears and the log stops after the banner. That is
    exactly the reported "the sign-in prompt never opens".

    Installing the same prerequisites here, with -Force -Confirm:$false, means the engine
    finds them already present and skips both prompting branches, going straight to
    Connect-MgGraph. Every step is guarded: if the gallery is unreachable the engine still
    runs and fails with a message, rather than hanging.

    Version 2.9.1 matches the engine's own -MaximumVersion pin.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()

    return @'
# ---- prepare the engine's online prerequisites, non-interactively ----
$ConfirmPreference = 'None'

Write-RunLine 'Checking sign-in prerequisites...'

try {
    $nuget = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue
    if (-not $nuget) {
        Write-RunLine '  Installing the NuGet package provider...'
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope AllUsers `
                                -Force -Confirm:$false -ErrorAction Stop | Out-Null
        Write-RunLine '  NuGet package provider installed.'
    }
    else {
        Write-RunLine '  NuGet package provider present.'
    }
}
catch {
    Write-RunLine ('  WARNING: could not install the NuGet provider: ' + $_.Exception.Message)
}

try {
    # Untrusted is the default and makes Install-Module prompt. -Force suppresses that, but
    # trusting it explicitly keeps the engine's own Install-Module call quiet too.
    $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($repo -and $repo.InstallationPolicy -ne 'Trusted') {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
        Write-RunLine '  PSGallery marked trusted for this machine.'
    }
}
catch {
    Write-RunLine ('  WARNING: could not set the PSGallery policy: ' + $_.Exception.Message)
}

# Getting the sign-in module *loaded* is the goal, not merely installed. A present but
# incomplete install (files left as OneDrive cloud-only placeholders, say) is discoverable by
# Get-Module -ListAvailable yet throws on import, and because PowerShell resolves an
# unversioned Import-Module to the HIGHEST version, a broken newer copy defeats both the
# engine's version pin and any attempt to install a good one alongside it.
#
# So: load a known-good version explicitly by manifest path. Once it is loaded, the engine's
# own `Import-Module Microsoft.Graph.Authentication` is a no-op and keeps this one. Nothing is
# uninstalled here; use "Repair sign-in module" on the Advanced page to clean up properly.
function Import-GraphAuthModule {
    $installed = @(Get-Module -ListAvailable -Name Microsoft.Graph.Authentication -ErrorAction SilentlyContinue |
                    Sort-Object Version -Descending)

    if ($installed.Count -gt 0) {
        Write-RunLine ('  Installed versions: ' + (($installed | ForEach-Object { $_.Version.ToString() }) -join ', '))
    }

    # Highest first: if the newest works, that is what the engine would have used anyway.
    foreach ($candidate in $installed) {
        try {
            Import-Module -Name $candidate.Path -ErrorAction Stop
            Write-RunLine ('  Loaded ' + $candidate.Version + ' from ' + $candidate.ModuleBase)
            return $true
        }
        catch {
            Write-RunLine ('  Version ' + $candidate.Version + ' will not load: ' + $_.Exception.Message)
        }
    }

    # Nothing usable on disk; fetch the version the engine pins.
    try {
        Write-RunLine '  Installing Microsoft.Graph.Authentication 2.9.1. This can take a minute...'
        Install-Module -Name Microsoft.Graph.Authentication -RequiredVersion 2.9.1 -Scope AllUsers `
                       -Force -AllowClobber -Confirm:$false -ErrorAction Stop

        $fresh = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication -ErrorAction SilentlyContinue |
                    Where-Object { $_.Version.ToString() -eq '2.9.1' } | Select-Object -First 1
        if ($fresh) {
            Import-Module -Name $fresh.Path -ErrorAction Stop
            Write-RunLine ('  Loaded ' + $fresh.Version + ' from ' + $fresh.ModuleBase)
            return $true
        }
    }
    catch {
        Write-RunLine ('  Install failed: ' + $_.Exception.Message)
    }

    return $false
}

if (Import-GraphAuthModule) {
    Write-RunLine '  Sign-in module ready.'
}
else {
    Write-RunLine '  ERROR: no usable Microsoft.Graph.Authentication could be loaded, so sign-in cannot start. Use "Repair sign-in module" on the Advanced page.'
}

Write-RunLine ('-' * 70)
'@
}

function New-ApLauncherScript {
    <#
    .SYNOPSIS
    Writes the launcher that invokes the engine and tees everything to the run log.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$EnginePath,
        [Parameter(Mandatory)][System.Collections.IDictionary]$Parameters,
        [Parameter(Mandatory)][string]$LogPath,
        [Parameter(Mandatory)][string]$LauncherPath
    )

    $paramLiteral = Format-ApHashtableLiteral $Parameters

    # Rendered here rather than in the launcher: building it there would need an if-expression,
    # which Windows PowerShell 5.1 does not support inside a string subexpression.
    $previewLine = Get-ApPreviewCommand -Parameters $Parameters -ScriptPath (Split-Path -Leaf $EnginePath)

    # Only an online run needs Graph; an offline CSV export must not touch the gallery.
    $prepareBlock = ''
    if ($Parameters.Contains('Online')) { $prepareBlock = Get-ApDependencyPrepBlock }

    # $log/$engine are injected as literals; everything else is escaped for the here-string.
    $body = @"
# Generated by Autopilot Import GUI (Community). Safe to delete.
`$ErrorActionPreference = 'Continue'
`$ProgressPreference    = 'SilentlyContinue'
`$WarningPreference     = 'Continue'

`$logPath = $(Format-ApPowerShellLiteral $LogPath)
`$engine  = $(Format-ApPowerShellLiteral $EnginePath)

`$params = $paramLiteral

# Force PSModulePath to Windows PowerShell locations only, before anything tries to
# auto-load a module. Appending the correct paths is not enough: PowerShell 7 entries sort
# first and shadow in-box modules with Core-only copies, which is how a run died at
# Get-PackageProvider with "not recognized" even though PackageManagement was installed.
# The GUI normalises this too, but the launcher repeats it so the generated script also works
# when run by hand from any shell. Pure string work, so it cannot fail for want of a module.
`$documentsPath = [Environment]::GetFolderPath('MyDocuments')
if (-not `$documentsPath) { `$documentsPath = Join-Path `$env:USERPROFILE 'Documents' }

`$modulePaths = New-Object System.Collections.Generic.List[string]
# Case-insensitive, because System32 and system32 arrive from different sources.
`$seenPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
foreach (`$candidate in @(
    (Join-Path `$documentsPath 'WindowsPowerShell\Modules')
    (Join-Path `$env:ProgramFiles 'WindowsPowerShell\Modules')
    (Join-Path `$env:WINDIR 'System32\WindowsPowerShell\v1.0\Modules')
)) {
    `$trimmedCandidate = "`$candidate".TrimEnd('\')
    if (`$seenPaths.Add(`$trimmedCandidate)) { `$modulePaths.Add(`$trimmedCandidate) }
}
if (`$env:PSModulePath) {
    foreach (`$inherited in (`$env:PSModulePath -split ';')) {
        `$trimmed = "`$inherited".TrimEnd('\')
        if (-not `$trimmed) { continue }
        # A real Windows PowerShell path always contains "WindowsPowerShell"; anything else
        # mentioning PowerShell belongs to PS7. Unrelated paths are preserved.
        if (`$trimmed -like '*powershell*' -and `$trimmed -notlike '*windowspowershell*') { continue }
        if (`$seenPaths.Add(`$trimmed)) { `$modulePaths.Add(`$trimmed) }
    }
}
`$env:PSModulePath = (`$modulePaths -join ';')

# One writer held open for the whole run, with AutoFlush so the GUI sees each line
# immediately. Add-Content was used here originally and intermittently dropped lines:
# it opens and closes the file per call, so it can collide with the GUI's tail reader,
# and its failure is non-terminating -- the line vanishes and the run carries on.
# FileShare.Read lets the reader in while keeping other writers out.
`$logStream = New-Object System.IO.FileStream(`$logPath, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
`$logWriter = New-Object System.IO.StreamWriter(`$logStream, (New-Object System.Text.UTF8Encoding(`$false)))
`$logWriter.AutoFlush = `$true

function Write-RunLine {
    param([string]`$Text)
    # The diagnostics script colours its status column with raw ANSI escapes (Write-Host
    # -ForegroundColor is fine, but it also emits VT sequences of its own). Neither a WPF
    # TextBox nor a log file interprets those, so they showed up verbatim as
    # "[93mSCP discovery successful[0m" once the formatter started rendering the table.
    `$logWriter.WriteLine((`$Text -replace ([char]27 + '\[[0-9;]*[A-Za-z]'), ''))
}

# Environment banner. When a run fails on someone else's bench this log is the only
# diagnostic available, so record what could not be inferred from the failure itself.
Write-RunLine ("Engine   : " + `$engine)
Write-RunLine ("Command  : " + $(Format-ApPowerShellLiteral $previewLine))
Write-RunLine ("Host     : PowerShell " + `$PSVersionTable.PSVersion + " (" + ([IntPtr]::Size * 8) + "-bit)")
Write-RunLine ("Modules  : " + `$env:PSModulePath)
Write-RunLine ("-" * 70)

# Setup is deliberately separate from the engine call and individually guarded. Previously a
# single try wrapped both, so one failing setup line aborted the whole run before sign-in.
#
# Set-ExecutionPolicy is NOT called here on purpose. powershell.exe is already launched with
# -ExecutionPolicy Bypass, which is authoritative for this process, so the call bought
# nothing - and the cmdlet lives in Microsoft.PowerShell.Security, a module that failed to
# load on a real machine ("the module could not be loaded"). That error is raised at command
# resolution, so -ErrorAction SilentlyContinue does not suppress it: it killed the run.
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 }
catch { Write-RunLine ("WARNING: could not enable TLS 1.2: " + `$_.Exception.Message) }

$prepareBlock

`$exitCode = 0

if (-not (Test-Path -LiteralPath `$engine)) {
    Write-RunLine ("ERROR: the engine script is missing: " + `$engine)
    `$exitCode = 1
}
else {
    try {
        # *>&1 folds every stream (including Write-Host's information stream) into the
        # success stream so the GUI sees the engine's full narration.
        #
        # Out-String -Stream rather than "`$_": the diagnostics script renders its observed
        # timeline with Format-Table, and casting those format records to string printed the
        # type names -- 35 consecutive lines of
        # "Microsoft.PowerShell.Commands.Internal.Format.FormatEntryData" -- in place of the
        # ESP phase table, which is the most useful part of the report. Out-String runs the
        # formatter properly; -Stream keeps it line-by-line so the tail reader still sees
        # output as it arrives, and -Width 200 stops the hidden console's narrow default
        # from truncating the table's columns.
        & `$engine @params *>&1 | Out-String -Stream -Width 200 | ForEach-Object { Write-RunLine `$_ }
    }
    catch {
        Write-RunLine ("ERROR: " + `$_.Exception.Message)
        if (`$_.ScriptStackTrace) { Write-RunLine (`$_.ScriptStackTrace) }
        `$exitCode = 1
    }
}

Write-RunLine ('$script:ApCompleteSentinel' + `$exitCode)

`$logWriter.Flush()
`$logWriter.Dispose()
exit `$exitCode
"@

    # The sentinel is a literal in the generated file, not expanded from our session.
    $body = $body.Replace("'`$script:ApCompleteSentinel'", (Format-ApPowerShellLiteral $script:ApCompleteSentinel))

    Set-Content -LiteralPath $LauncherPath -Value $body -Encoding UTF8
    return $LauncherPath
}

function Start-ApEngineRun {
    <#
    .SYNOPSIS
    Starts an engine run and returns a run context for polling.

    .PARAMETER Parameters
    Parameter hashtable from Build-ApEngineArguments.

    .PARAMETER ShowConsole
    Show the child PowerShell window. Off by default -- output belongs in the GUI -- but
    invaluable when diagnosing a run that appears to stall.

    .OUTPUTS
    A run context: Process, LogPath, ProgressState, StartTime, IsFinished, ExitCode.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Parameters,
        [string]$EnginePath,
        [switch]$ShowConsole,
        [string]$Label = 'run'
    )

    if (-not $EnginePath) { $EnginePath = Resolve-ApEngineScript -Name Engine }

    $workDir = Get-ApWorkingDirectory
    $stamp = Get-Date -Format 'HHmmss'
    $logPath = Join-Path $workDir ('{0}-{1}.log' -f $Label, $stamp)
    $launcherPath = Join-Path $workDir ('{0}-{1}-launcher.ps1' -f $Label, $stamp)

    # Create the log up front so the tail reader always has a file to open.
    Set-Content -LiteralPath $logPath -Value '' -Encoding UTF8

    New-ApLauncherScript -EnginePath $EnginePath -Parameters $Parameters `
                         -LogPath $logPath -LauncherPath $launcherPath | Out-Null

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = Get-ApPowerShellPath
    $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -STA -File "{0}"' -f $launcherPath
    $psi.WorkingDirectory = $workDir
    $psi.UseShellExecute = $true   # required for WindowStyle to apply
    $psi.WindowStyle = if ($ShowConsole) { 'Normal' } else { 'Hidden' }

    Write-ApLog "Starting engine run: $(Get-ApPreviewCommand -Parameters $Parameters -ScriptPath (Split-Path -Leaf $EnginePath))"
    Write-ApLog "Run log: $logPath"

    $process = [System.Diagnostics.Process]::Start($psi)

    return [pscustomobject]@{
        Process       = $process
        EnginePath    = $EnginePath
        LogPath       = $logPath
        LauncherPath  = $launcherPath
        Parameters    = $Parameters
        ProgressState = New-ApProgressState
        StartTime     = Get-Date
        IsFinished    = $false
        ExitCode      = $null
        Cancelled     = $false
        BytesRead     = 0L
        LineCount     = 0
        # Incomplete trailing line carried between polls (see Update-ApEngineRun).
        PendingLine   = ''
        # Stall detection: a hidden console cannot show an interactive prompt, so a blocked
        # engine looks identical to a slow one. Tracking silence lets the UI say so.
        LastOutputTime = Get-Date
        StallReported  = $false
    }
}

function Update-ApEngineRun {
    <#
    .SYNOPSIS
    Reads whatever the engine has written since the last call and folds it into progress.

    .DESCRIPTION
    Designed to be called from a DispatcherTimer tick. Reads from the recorded byte offset
    with FileShare.ReadWrite so it never contends with the writing process.

    .OUTPUTS
    PSCustomObject with NewLines, ProgressChanged, IsFinished, ExitCode.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Run
    )

    $newLines = New-Object System.Collections.Generic.List[string]
    $progressChanged = $false

    if (Test-Path -LiteralPath $Run.LogPath) {
        $stream = $null
        $reader = $null
        try {
            $stream = New-Object System.IO.FileStream(
                $Run.LogPath,
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::Read,
                [System.IO.FileShare]::ReadWrite)

            if ($Run.BytesRead -gt 0 -and $Run.BytesRead -le $stream.Length) {
                [void]$stream.Seek($Run.BytesRead, [System.IO.SeekOrigin]::Begin)
            }

            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
            $chunk = $reader.ReadToEnd()
            $Run.BytesRead = $stream.Position

            if ($chunk) {
                # A read can land mid-line. Carry the incomplete tail over to the next tick,
                # otherwise a single log line surfaces as two truncated fragments and the
                # progress regexes stop matching it.
                $chunk = $Run.PendingLine + $chunk
                $Run.PendingLine = ''

                $parts = $chunk -split "`r?`n"
                if (-not ($chunk.EndsWith("`n"))) {
                    $Run.PendingLine = $parts[-1]
                    if ($parts.Count -gt 1) { $parts = $parts[0..($parts.Count - 2)] }
                    else { $parts = @() }
                }

                foreach ($line in $parts) {
                    if ($null -eq $line) { continue }

                    if ($line.StartsWith($script:ApCompleteSentinel)) {
                        $codeText = $line.Substring($script:ApCompleteSentinel.Length).Trim()
                        $parsed = 0
                        if ([int]::TryParse($codeText, [ref]$parsed)) { $Run.ExitCode = $parsed }
                        continue
                    }

                    if ([string]::IsNullOrWhiteSpace($line)) { continue }

                    $newLines.Add($line)
                    $Run.LineCount++
                    $Run.LastOutputTime = Get-Date

                    if (Update-ApProgressState -State $Run.ProgressState -Line $line) {
                        $progressChanged = $true
                    }
                }
            }
        }
        catch {
            # A transient read error just means we retry on the next tick.
        }
        finally {
            if ($reader) { $reader.Dispose() }
            elseif ($stream) { $stream.Dispose() }
        }
    }

    # Only declare the run finished once the process is gone AND the tail is drained,
    # otherwise the last few lines (including the outcome) can be lost.
    if (-not $Run.IsFinished -and $Run.Process -and $Run.Process.HasExited -and $newLines.Count -eq 0) {

        # The process cannot write again, so an unterminated tail is now a whole line.
        if ($Run.PendingLine -and -not [string]::IsNullOrWhiteSpace($Run.PendingLine)) {
            $tail = $Run.PendingLine
            $Run.PendingLine = ''
            $newLines.Add($tail)
            $Run.LineCount++
            if (Update-ApProgressState -State $Run.ProgressState -Line $tail) { $progressChanged = $true }
        }

        $Run.IsFinished = $true
        if ($null -eq $Run.ExitCode) {
            try { $Run.ExitCode = $Run.Process.ExitCode } catch { $Run.ExitCode = -1 }
        }

        $elapsed = (Get-Date) - $Run.StartTime
        Write-ApLog ("Engine run finished with exit code {0} after {1:mm\:ss}" -f $Run.ExitCode, $elapsed)
    }

    # Report a stall once. A run that has gone quiet for this long is usually blocked on an
    # interactive prompt in the hidden console, which the operator can neither see nor answer.
    $stalled = $false
    if (-not $Run.IsFinished -and -not $Run.StallReported) {
        if (((Get-Date) - $Run.LastOutputTime).TotalSeconds -ge 90) {
            $Run.StallReported = $true
            $stalled = $true
            Write-ApLog "Engine run has produced no output for 90 seconds; it may be waiting on a hidden prompt." -Level WARN
        }
    }

    return [pscustomobject]@{
        NewLines        = $newLines.ToArray()
        ProgressChanged = $progressChanged
        IsFinished      = $Run.IsFinished
        ExitCode        = $Run.ExitCode
        Stalled         = $stalled
    }
}

function Stop-ApEngineRun {
    <#
    .SYNOPSIS
    Terminates a run and its child processes.

    .DESCRIPTION
    Uses taskkill /T because the engine may itself have started children (the
    pre-provisioning helper, sysprep, changepk) that a plain Process.Kill would orphan.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Run
    )

    if (-not $Run.Process) { return }

    try {
        if (-not $Run.Process.HasExited) {
            Write-ApLog "Cancelling engine run (PID $($Run.Process.Id))" -Level WARN
            & taskkill.exe /PID $Run.Process.Id /T /F 2>&1 | Out-Null

            if (-not $Run.Process.WaitForExit(5000)) {
                try { $Run.Process.Kill() } catch { }
            }
        }
    }
    catch {
        Write-ApLog "Could not cancel the run cleanly: $($_.Exception.Message)" -Level WARN
    }
    finally {
        $Run.Cancelled = $true
        $Run.IsFinished = $true
        if ($null -eq $Run.ExitCode) { $Run.ExitCode = -1 }
    }
}

function Get-ApRunSummary {
    <#
    .SYNOPSIS
    One-line outcome for the status bar.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Run)

    $elapsed = (Get-Date) - $Run.StartTime
    $duration = '{0:mm\:ss}' -f $elapsed

    if ($Run.Cancelled) { return "Cancelled after $duration" }

    $state = $Run.ProgressState
    if ($state.IsError) { return "Failed after ${duration}: $($state.ErrorMessage)" }
    if ($Run.ExitCode -ne 0) { return "Finished with errors after $duration (exit code $($Run.ExitCode)). See the log for details." }
    if ($state.IsComplete) { return "Completed in $duration" }

    return "Finished in $duration"
}
#endregion src\Private\ScriptRunner.ps1

#region src\Private\Connectivity.ps1
# Connectivity.ps1 -- Autopilot network prerequisite checks.
#
# A rewrite of connectivity_check() from the original GUI
# (Get-WindowsAutopilotImportGUI.ps1:245-508), which never actually worked: every result
# line called `Write-Output -NoNewline -ForegroundColor ...`, and Write-Output has no
# -NoNewline or -ForegroundColor parameter, so all ~29 checks threw a
# ParameterBindingException that was then swallowed by $ErrorActionPreference =
# 'SilentlyContinue'. The tech saw an empty window and concluded the network was fine.
#
# Differences here:
#   * raw TcpClient with an explicit timeout instead of Test-NetConnection -- roughly an
#     order of magnitude faster, reports latency, and needs no extra module
#   * runspace pool, so ~30 endpoints complete in about a second rather than serially
#   * results are returned as objects for a DataGrid rather than printed with colour codes
#   * endpoints are data, overridable via connectivityEndpoints in config.json
#   * dropped the original's bare "azure.net" probe, which is not a resolvable host, and
#     replaced it with the documented TPM attestation endpoint

function Get-ApDefaultEndpoints {
    <#
    .SYNOPSIS
    Documented Windows Autopilot / Intune network requirements.

    .DESCRIPTION
    Required = $true marks endpoints where a failure will actually break enrolment;
    the rest are deployment-time niceties reported for completeness.
    #>
    [CmdletBinding()]
    param()

    return @(
        # --- Entra ID / MDM enrolment -----------------------------------------
        [pscustomobject]@{ Category = 'Entra ID and enrolment'; Name = 'Entra device registration'; Host = 'enterpriseregistration.windows.net'; Port = 443; Required = $true }
        [pscustomobject]@{ Category = 'Entra ID and enrolment'; Name = 'Intune enrolment';          Host = 'enterpriseenrollment-s.manage.microsoft.com'; Port = 443; Required = $true }
        [pscustomobject]@{ Category = 'Entra ID and enrolment'; Name = 'Entra sign-in';             Host = 'login.microsoftonline.com'; Port = 443; Required = $true }
        [pscustomobject]@{ Category = 'Entra ID and enrolment'; Name = 'Microsoft Graph';           Host = 'graph.microsoft.com'; Port = 443; Required = $true }

        # --- Autopilot deployment service -------------------------------------
        [pscustomobject]@{ Category = 'Autopilot service'; Name = 'Autopilot ZTD';        Host = 'ztd.dds.microsoft.com'; Port = 443; Required = $true }
        [pscustomobject]@{ Category = 'Autopilot service'; Name = 'Autopilot CS';         Host = 'cs.dds.microsoft.com'; Port = 443; Required = $true }
        [pscustomobject]@{ Category = 'Autopilot service'; Name = 'Microsoft account';    Host = 'login.live.com'; Port = 443; Required = $true }
        [pscustomobject]@{ Category = 'Autopilot service'; Name = 'Intune service';       Host = 'manage.microsoft.com'; Port = 443; Required = $true }

        # --- TPM attestation (required for pre-provisioning / self-deploying) --
        [pscustomobject]@{ Category = 'TPM attestation'; Name = 'Intel TPM';             Host = 'ekop.intel.com'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'TPM attestation'; Name = 'Qualcomm TPM';          Host = 'ekcert.spserv.microsoft.com'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'TPM attestation'; Name = 'AMD TPM';               Host = 'ftpm.amd.com'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'TPM attestation'; Name = 'Host attestation';      Host = 'has.spserv.microsoft.com'; Port = 443; Required = $false }

        # --- Licensing / activation -------------------------------------------
        [pscustomobject]@{ Category = 'Activation'; Name = 'Activation service';         Host = 'activation.sls.microsoft.com'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'Activation'; Name = 'Validation service';         Host = 'validation.sls.microsoft.com'; Port = 443; Required = $false }

        # --- Windows Update ----------------------------------------------------
        [pscustomobject]@{ Category = 'Windows Update'; Name = 'Windows Update';         Host = 'update.microsoft.com'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'Windows Update'; Name = 'Update download';        Host = 'download.windowsupdate.com'; Port = 443; Required = $false }
        # Microsoft documents these as *.delivery.mp.microsoft.com and
        # *.dsp.mp.microsoft.com. The bare wildcard bases do not resolve, so probe a real
        # host from each rather than reporting a permanent false failure.
        [pscustomobject]@{ Category = 'Windows Update'; Name = 'Content delivery';       Host = 'dl.delivery.mp.microsoft.com'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'Windows Update'; Name = 'Delivery optimisation';  Host = 'tsfe.trafficshaping.dsp.mp.microsoft.com'; Port = 443; Required = $false }

        # --- Single sign-on ----------------------------------------------------
        [pscustomobject]@{ Category = 'Single sign-on'; Name = 'Seamless SSO';           Host = 'autologon.microsoftazuread-sso.com'; Port = 443; Required = $false }

        # --- Configuration and app delivery -----------------------------------
        [pscustomobject]@{ Category = 'Configuration and apps'; Name = 'Office config';  Host = 'config.office.com'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'Configuration and apps'; Name = 'Entra Graph (legacy)'; Host = 'graph.windows.net'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'Configuration and apps'; Name = 'Content CDN (primary)';  Host = 'euprodimedatapri.azureedge.net'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'Configuration and apps'; Name = 'Content CDN (secondary)'; Host = 'euprodimedatasec.azureedge.net'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'Configuration and apps'; Name = 'Content CDN (hotfix)';    Host = 'euprodimedatahotfix.azureedge.net'; Port = 443; Required = $false }

        # --- Diagnostics -------------------------------------------------------
        [pscustomobject]@{ Category = 'Diagnostics'; Name = 'Connectivity test';         Host = 'www.msftconnecttest.com'; Port = 443; Required = $false }
        [pscustomobject]@{ Category = 'Diagnostics'; Name = 'PowerShell Gallery';        Host = 'www.powershellgallery.com'; Port = 443; Required = $false }
    )
}

function Get-ApConfiguredEndpoints {
    <#
    .SYNOPSIS
    The endpoint list to test: the config.json override if present, else the built-in list.
    #>
    [CmdletBinding()]
    param()

    $config = Get-ApConfig
    $override = $config.connectivityEndpoints

    if (-not $override -or @($override).Count -eq 0) { return Get-ApDefaultEndpoints }

    $list = New-Object System.Collections.Generic.List[object]
    foreach ($e in @($override)) {
        $endpointHost = Get-ApRequestValue $e 'Host' ''
        if (-not $endpointHost) { continue }

        $list.Add([pscustomobject]@{
            Category = Get-ApRequestValue $e 'Category' 'Custom'
            Name     = Get-ApRequestValue $e 'Name' $endpointHost
            Host     = $endpointHost
            Port     = [int](Get-ApRequestValue $e 'Port' 443)
            Required = [bool](Get-ApRequestValue $e 'Required' $false)
        })
    }

    if ($list.Count -eq 0) {
        Write-ApLog 'connectivityEndpoints in config.json contained no usable entries; using the built-in list.' -Level WARN
        return Get-ApDefaultEndpoints
    }

    Write-ApLog "Using $($list.Count) custom connectivity endpoints from config.json"
    return $list.ToArray()
}

function Test-ApEndpoint {
    <#
    .SYNOPSIS
    Single TCP reachability probe with latency.

    .DESCRIPTION
    Deliberately standalone with no dependencies on the rest of the app, because this
    function's source is injected into a runspace pool by Invoke-ApConnectivityCheck.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [int]$Port = 443,
        [int]$TimeoutMs = 4000
    )

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $client = $null
    $succeeded = $false
    $errorText = ''

    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $async = $client.BeginConnect($ComputerName, $Port, $null, $null)

        if ($async.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {
            # EndConnect surfaces the real failure (refused, reset, DNS) rather than
            # letting a half-open socket look like a success.
            $client.EndConnect($async)
            $succeeded = $client.Connected
            if (-not $succeeded) { $errorText = 'Connection not established' }
        }
        else {
            $errorText = "Timed out after ${TimeoutMs}ms"
        }
    }
    catch {
        $errorText = $_.Exception.GetBaseException().Message
    }
    finally {
        if ($client) { try { $client.Close() } catch { } }
        $sw.Stop()
    }

    return [pscustomobject]@{
        Host      = $ComputerName
        Port      = $Port
        Succeeded = $succeeded
        LatencyMs = [int]$sw.ElapsedMilliseconds
        Error     = $errorText
    }
}

function Invoke-ApConnectivityCheck {
    <#
    .SYNOPSIS
    Tests every configured endpoint concurrently.

    .PARAMETER OnResult
    Optional scriptblock invoked with each result object as it completes, so the UI can
    populate the grid progressively instead of waiting for the whole sweep.

    .PARAMETER ThrottleLimit
    Concurrent probes. 12 keeps a slow OOBE network from being swamped while still
    finishing the full list in about a second.
    #>
    [CmdletBinding()]
    param(
        [object[]]$Endpoints,
        [scriptblock]$OnResult,
        [int]$ThrottleLimit = 12,
        [int]$TimeoutMs = 4000
    )

    # Deliberately no logging or config access in here: this function is copied into a bare
    # background runspace by Start-ApBackgroundWork, where Write-ApLog and Get-ApConfig do
    # not exist. Callers resolve the endpoint list and do the logging.
    if (-not $Endpoints) { throw 'Invoke-ApConnectivityCheck requires -Endpoints.' }

    $pool = [runspacefactory]::CreateRunspacePool(1, $ThrottleLimit)
    $pool.ApartmentState = 'MTA'
    $pool.Open()

    # Ship the probe function into each runspace as source text.
    $probeSource = "function Test-ApEndpoint {`n$((Get-Command Test-ApEndpoint).Definition)`n}"

    $jobs = New-Object System.Collections.Generic.List[object]

    try {
        foreach ($endpoint in $Endpoints) {
            $ps = [powershell]::Create()
            $ps.RunspacePool = $pool

            [void]$ps.AddScript($probeSource)
            [void]$ps.AddScript('param($h,$p,$t) Test-ApEndpoint -ComputerName $h -Port $p -TimeoutMs $t')
            [void]$ps.AddArgument($endpoint.Host)
            [void]$ps.AddArgument($endpoint.Port)
            [void]$ps.AddArgument($TimeoutMs)

            $jobs.Add([pscustomobject]@{
                Shell    = $ps
                Handle   = $ps.BeginInvoke()
                Endpoint = $endpoint
            })
        }

        $results = New-Object System.Collections.Generic.List[object]

        foreach ($job in $jobs) {
            $probe = $null
            try {
                $probe = @($job.Shell.EndInvoke($job.Handle))[-1]
            }
            catch {
                $probe = $null
            }
            finally {
                $job.Shell.Dispose()
            }

            $succeeded = [bool]($probe -and $probe.Succeeded)

            $result = [pscustomobject]@{
                Category  = $job.Endpoint.Category
                Name      = $job.Endpoint.Name
                Host      = $job.Endpoint.Host
                Port      = $job.Endpoint.Port
                Required  = $job.Endpoint.Required
                Succeeded = $succeeded
                LatencyMs = if ($probe) { $probe.LatencyMs } else { 0 }
                Error     = if ($probe) { $probe.Error } else { 'Probe did not run' }
                Status    = if ($succeeded) { 'OK' } elseif ($job.Endpoint.Required) { 'Failed' } else { 'Unreachable' }
            }

            $results.Add($result)
            if ($OnResult) { & $OnResult $result }
        }

        return $results.ToArray()
    }
    finally {
        $pool.Close()
        $pool.Dispose()
    }
}
#endregion src\Private\Connectivity.ps1

#region src\Private\Preflight.ps1
# Preflight.ps1 -- checks that must pass before an online run can possibly succeed.
#
# Interactive sign-in happens inside the engine, which does:
#
#     $module = Import-Module microsoft.graph.authentication -PassThru -ErrorAction Ignore
#     if (-not $module) { Install-Module microsoft.graph.authentication -MaximumVersion 2.9.1 }
#     ...
#     Import-Module Microsoft.Graph.Authentication     # unversioned: highest wins
#
# Two consequences worth pre-empting:
#
#   * -ErrorAction Ignore means a *broken* install looks the same as a missing one, so the
#     engine tries to install 2.9.1 and then imports unversioned anyway, which picks the
#     highest installed version. A broken newer version therefore defeats both the engine's
#     version pin and its own repair attempt.
#   * A module under a OneDrive-redirected Documents folder is frequently incomplete: the
#     manifest is present but files it references (.format.ps1xml, assemblies) are still
#     cloud-only placeholders. The import fails with "the member 'FormatsToProcess' in the
#     module manifest is not valid".
#
# Without this check the tech sees only a non-zero exit code and no browser prompt.

$script:ApGraphModuleName = 'Microsoft.Graph.Authentication'

# The version the engine pins with -MaximumVersion; what a repair should install.
$script:ApGraphPinnedVersion = '2.9.1'

function Get-ApWindowsPowerShellModulePath {
    <#
    .SYNOPSIS
    A PSModulePath containing only module locations valid for Windows PowerShell 5.1.

    .DESCRIPTION
    Launching this tool from a PowerShell 7 terminal leaks PS7's PSModulePath into the 5.1
    process, and children inherit it. Observed on a real machine:

        C:\Users\<u>\OneDrive\Documents\PowerShell\Modules
        C:\Program Files\PowerShell\Modules
        C:\Program Files\WindowsApps\Microsoft.PowerShell_7.6.4.0_x64__...\Modules
        C:\Program Files\WindowsPowerShell\Modules
        C:\WINDOWS\system32\WindowsPowerShell\v1.0\Modules

    The PS7 entries sort first and shadow in-box modules with Core-only copies, so 5.1 either
    refuses to load them or cannot see them at all. That produced two different symptoms from
    one cause: "the 'Set-ExecutionPolicy' command was found in the module
    'Microsoft.PowerShell.Security', but the module could not be loaded", and
    "the term 'Get-PackageProvider' is not recognized". Both aborted the engine before any
    sign-in prompt could appear.

    Filtering rule: a genuine Windows PowerShell module path always contains
    "WindowsPowerShell". Anything else mentioning "PowerShell" belongs to PS7 and is dropped.
    Unrelated custom paths (a management agent's module directory, say) mention neither and
    are preserved.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()

    # Documents may be redirected (OneDrive), so ask Windows rather than assuming.
    $documents = [Environment]::GetFolderPath('MyDocuments')
    if (-not $documents) { $documents = Join-Path $env:USERPROFILE 'Documents' }

    $canonical = @(
        (Join-Path $documents 'WindowsPowerShell\Modules')
        (Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules')
        (Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\Modules')
    )

    $preserved = @()
    if ($env:PSModulePath) {
        $preserved = @(
            $env:PSModulePath -split ';' |
                Where-Object { $_ -and $_.Trim() } |
                ForEach-Object { $_.TrimEnd('\') } |
                Where-Object {
                    # Drop PowerShell 7 locations; keep everything unrelated.
                    -not ($_ -like '*powershell*' -and $_ -notlike '*windowspowershell*')
                }
        )
    }

    # Case-insensitive dedupe: Windows paths differ only in casing between sources
    # (System32 vs system32), and a List[string].Contains would keep both.
    $ordered = New-Object System.Collections.Generic.List[string]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($path in @($canonical + $preserved)) {
        $trimmed = "$path".TrimEnd('\')
        if (-not $trimmed) { continue }
        if ($seen.Add($trimmed)) { $ordered.Add($trimmed) }
    }

    return ($ordered -join ';')
}

function Repair-ApModulePath {
    <#
    .SYNOPSIS
    Rewrites $env:PSModulePath for this process so Windows PowerShell modules resolve.

    .DESCRIPTION
    Called once at startup. Child processes inherit the corrected value, so the engine
    launcher, the Windows Update console and the Graph repair console all benefit.
    #>
    [CmdletBinding()]
    param()

    $before = "$env:PSModulePath"
    $after = Get-ApWindowsPowerShellModulePath
    $env:PSModulePath = $after

    if ($before -ne $after) {
        $removed = @(
            $before -split ';' | Where-Object {
                $_ -and ($_ -like '*powershell*' -and $_ -notlike '*windowspowershell*')
            }
        )
        if ($removed.Count -gt 0) {
            Write-ApLog "Removed $($removed.Count) PowerShell 7 module path(s) that would shadow Windows PowerShell modules: $($removed -join '; ')" -Level WARN
        }
        Write-ApLog "PSModulePath set to: $after"
    }

    return $after
}

function Test-ApGraphModule {
    <#
    .SYNOPSIS
    Reports whether the Graph authentication module can actually be imported.

    .DESCRIPTION
    Self-contained so it can be copied into a bare background runspace: no logging, no
    config. Imports for real rather than merely checking that a folder exists, because the
    failure mode being detected is precisely a present-but-incomplete install.

    Import it the same way the engine does -- unversioned, so the highest installed version
    wins -- and only then look for a lower version that does work, which is what makes the
    remedy actionable.

    .OUTPUTS
    PSCustomObject: Available, Version, Path, Error, InstalledVersions, WorkingVersion.
    #>
    [CmdletBinding()]
    param()

    $result = [ordered]@{
        Available         = $false
        Version           = ''
        Path              = ''
        Error             = ''
        InstalledVersions = @()
        WorkingVersion    = ''
    }

    $name = 'Microsoft.Graph.Authentication'

    $available = @(Get-Module -ListAvailable -Name $name -ErrorAction SilentlyContinue |
                    Sort-Object Version -Descending)
    $result.InstalledVersions = @($available | ForEach-Object { "$($_.Version)" })

    if ($available.Count -eq 0) {
        $result.Error = 'The module is not installed.'
        return [pscustomobject]$result
    }

    # Exactly what the engine does.
    try {
        $imported = Import-Module -Name $name -PassThru -ErrorAction Stop
        $result.Available = $true
        $result.Version = "$($imported.Version)"
        $result.Path = "$($imported.ModuleBase)"
        return [pscustomobject]$result
    }
    catch {
        $result.Error = $_.Exception.Message
        $result.Version = "$($available[0].Version)"
        $result.Path = "$($available[0].ModuleBase)"
    }

    # Is there an older version that does load? If so, the fix is to remove the broken one.
    foreach ($candidate in ($available | Select-Object -Skip 1)) {
        try {
            Import-Module -Name $candidate.Path -PassThru -ErrorAction Stop | Out-Null
            $result.WorkingVersion = "$($candidate.Version)"
            break
        }
        catch {
            continue
        }
    }

    return [pscustomobject]$result
}

function Get-ApGraphModuleAdvice {
    <#
    .SYNOPSIS
    Turns a Test-ApGraphModule result into text a technician can act on.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Check)

    if ($Check.Available) { return '' }

    $lines = New-Object System.Collections.Generic.List[string]

    if ($Check.InstalledVersions.Count -eq 0) {
        $lines.Add("The $script:ApGraphModuleName module is not installed, so signing in is not possible.")
        $lines.Add('The engine will try to install it automatically, which needs access to the PowerShell Gallery.')
        return ($lines -join [Environment]::NewLine)
    }

    $lines.Add("$script:ApGraphModuleName $($Check.Version) is installed but cannot be loaded, so sign-in will fail before a browser prompt appears.")
    $lines.Add('')
    $lines.Add("Location: $($Check.Path)")
    $lines.Add("Reported: $($Check.Error)")

    if ($Check.Path -like '*OneDrive*') {
        $lines.Add('')
        $lines.Add('That path is inside OneDrive. Redirected Documents folders often leave module files as cloud-only placeholders, so the manifest is present but the files it references are not.')
    }

    if ($Check.WorkingVersion) {
        $lines.Add('')
        $lines.Add("Version $($Check.WorkingVersion) is also installed and does load, but PowerShell imports the highest version, so the broken one wins.")
    }

    return ($lines -join [Environment]::NewLine)
}

function Get-ApGraphRepairScript {
    <#
    .SYNOPSIS
    A script that removes every installed copy of the module and reinstalls the pinned version.

    .DESCRIPTION
    Installs to AllUsers (under Program Files) deliberately: that is outside any redirected
    Documents folder, which is what caused the broken install in the first place.
    #>
    [CmdletBinding()]
    param()

    return @"
`$Host.UI.RawUI.WindowTitle = 'Autopilot Import GUI - Repair Graph module'
`$ErrorActionPreference = 'Continue'

function Write-Step { param([string]`$Text) Write-Host "==> `$Text" -ForegroundColor Cyan }

Write-Host ''
Write-Host ' Repairing $script:ApGraphModuleName' -ForegroundColor White
Write-Host ''

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    Write-Step 'Currently installed versions:'
    Get-Module -ListAvailable -Name $script:ApGraphModuleName |
        Select-Object Version, ModuleBase | Format-Table -AutoSize | Out-String | Write-Host

    Write-Step 'Removing all installed versions...'
    try {
        Uninstall-Module -Name $script:ApGraphModuleName -AllVersions -Force -ErrorAction Stop
        Write-Host '    Removed via Uninstall-Module.' -ForegroundColor Gray
    }
    catch {
        Write-Host "    Uninstall-Module could not remove it: `$(`$_.Exception.Message)" -ForegroundColor Yellow
        # A hand-copied or partially synced module is not registered with PowerShellGet, so
        # Uninstall-Module cannot see it. Delete the folders directly.
        Write-Step 'Deleting module folders directly...'
        foreach (`$found in (Get-Module -ListAvailable -Name $script:ApGraphModuleName)) {
            try {
                Remove-Item -LiteralPath `$found.ModuleBase -Recurse -Force -ErrorAction Stop
                Write-Host "    Deleted `$(`$found.ModuleBase)" -ForegroundColor Gray
            }
            catch {
                Write-Host "    Could not delete `$(`$found.ModuleBase): `$(`$_.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    Write-Step 'Installing version $script:ApGraphPinnedVersion for all users...'
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:`$false -ErrorAction SilentlyContinue | Out-Null
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
    Install-Module -Name $script:ApGraphModuleName -RequiredVersion $script:ApGraphPinnedVersion ``
                   -Scope AllUsers -Force -AllowClobber -Confirm:`$false -ErrorAction Stop

    Write-Step 'Verifying the module imports...'
    Import-Module -Name $script:ApGraphModuleName -Force -ErrorAction Stop
    `$loaded = Get-Module -Name $script:ApGraphModuleName

    Write-Host ''
    Write-Host " Repaired. Loaded version `$(`$loaded.Version) from `$(`$loaded.ModuleBase)" -ForegroundColor Green
    Write-Host ' Close and reopen Autopilot Import GUI, then try signing in again.' -ForegroundColor Gray
}
catch {
    Write-Host ''
    Write-Host " Repair failed: `$(`$_.Exception.Message)" -ForegroundColor Red
    Write-Host ' You can repair it manually with:' -ForegroundColor Gray
    Write-Host '   Uninstall-Module $script:ApGraphModuleName -AllVersions -Force' -ForegroundColor Gray
    Write-Host '   Install-Module $script:ApGraphModuleName -RequiredVersion $script:ApGraphPinnedVersion -Scope AllUsers -Force' -ForegroundColor Gray
}

Write-Host ''
Read-Host 'Press Enter to close'
"@
}
#endregion src\Private\Preflight.ps1

#region src\Private\Background.ps1
# Background.ps1 -- run work off the UI thread and marshal the result back.
#
# WPF has a single UI thread: anything slow on it freezes the window, including the
# "is the network broken" check, which is precisely the moment a frozen window is least
# forgivable (a fully unreachable endpoint list costs one timeout per parallel batch).
#
# Pattern follows VM-Pilot (VMPilot.GUI.ps1:1167-1232): a dedicated runspace for the work,
# a DispatcherTimer to notice completion, and the continuation invoked back on the UI
# thread so it can touch controls safely.

$script:ApBackgroundJobs = New-Object System.Collections.Generic.List[object]

function Start-ApBackgroundWork {
    <#
    .SYNOPSIS
    Runs a scriptblock in a background runspace and calls OnComplete on the UI thread.

    .PARAMETER Work
    Scriptblock to execute. It sees the variables supplied in -Variables and any functions
    named in -FunctionNames.

    .PARAMETER Variables
    Name/value pairs injected into the runspace's session state.

    .PARAMETER FunctionNames
    Functions from the current session to copy into the runspace by source text. Only pass
    self-contained workers; a function that logs or reads config will not find those helpers.

    .PARAMETER OnComplete
    Invoked on the UI thread with two arguments: the work's output, and an error message
    ($null on success).

    .PARAMETER Window
    Window whose Dispatcher is used to marshal the continuation.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$Work,
        [hashtable]$Variables = @{},
        [string[]]$FunctionNames = @(),
        [Parameter(Mandatory)][scriptblock]$OnComplete,
        [Parameter(Mandatory)]$Window,
        [int]$PollIntervalMs = 150
    )

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = 'MTA'    # worker threads do no UI work
    $runspace.ThreadOptions = 'ReuseThread'
    $runspace.Open()

    foreach ($key in $Variables.Keys) {
        $runspace.SessionStateProxy.SetVariable($key, $Variables[$key])
    }

    $shell = [powershell]::Create()
    $shell.Runspace = $runspace

    # Copy requested functions in as source. A runspace starts empty, so anything the
    # work calls must either be a built-in cmdlet or arrive this way.
    foreach ($name in $FunctionNames) {
        $command = Get-Command $name -CommandType Function -ErrorAction SilentlyContinue
        if (-not $command) {
            Write-ApLog "Cannot copy function '$name' into the background runspace: not found." -Level WARN
            continue
        }
        [void]$shell.AddScript("function $name {`n$($command.Definition)`n}")
    }

    [void]$shell.AddScript($Work.ToString())

    $job = [pscustomobject]@{
        Shell    = $shell
        Runspace = $runspace
        Handle   = $null
        Timer    = $null
    }

    $job.Handle = $shell.BeginInvoke()
    $script:ApBackgroundJobs.Add($job)

    # Capture the list as a local so the tick closure holds a direct reference. Inside a
    # GetNewClosure() scriptblock, $script: binds to the closure's own captured session
    # state, where this variable is not visible; calling .Remove() on the resulting $null
    # threw out of the finally block and the continuation never ran, so the UI sat on
    # "reading..." forever with no error shown.
    $jobList = $script:ApBackgroundJobs

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds($PollIntervalMs)
    $job.Timer = $timer

    # Everything inside a DispatcherTimer tick must be caught. An exception that escapes a
    # tick is not merely logged by WPF: it propagates out of Window.ShowDialog and tears the
    # whole application down, which looks to the user like the window vanishing on startup.
    $timer.Add_Tick({
        try {
            if (-not $job.Handle.IsCompleted) { return }

            $timer.Stop()

            $output = $null
            $errorMessage = $null
            try {
                $output = $job.Shell.EndInvoke($job.Handle)

                # A non-terminating error inside the runspace surfaces here, not as an exception.
                if ($job.Shell.HadErrors -and $job.Shell.Streams.Error.Count -gt 0) {
                    $errorMessage = ($job.Shell.Streams.Error | ForEach-Object { "$_" }) -join '; '
                }
            }
            catch {
                $errorMessage = $_.Exception.Message
            }
            finally {
                try { $job.Shell.Dispose() } catch { }
                try { $job.Runspace.Close(); $job.Runspace.Dispose() } catch { }
                try { [void]$jobList.Remove($job) } catch { }
            }

            & $OnComplete $output $errorMessage
        }
        catch {
            try { $timer.Stop() } catch { }
            Write-ApLog "Background work continuation failed: $($_.Exception.Message)" -Level ERROR
            if ($_.ScriptStackTrace) { Write-ApLog $_.ScriptStackTrace -Level DEBUG }
        }
    }.GetNewClosure())

    $timer.Start()
    return $job
}

function Stop-ApBackgroundWork {
    <#
    .SYNOPSIS
    Tears down any still-running background jobs. Called on window close.

    .DESCRIPTION
    Best-effort by design: this runs while the window is closing, and a job may already be
    half-disposed by its own completion handler. Every step is guarded individually so one
    failure cannot abort the rest of the shutdown sequence.
    #>
    [CmdletBinding()]
    param()

    $jobs = @()
    try { $jobs = @($script:ApBackgroundJobs) } catch { $jobs = @() }

    foreach ($job in $jobs) {
        if (-not $job) { continue }
        try { if ($job.Timer) { $job.Timer.Stop() } } catch { }
        try { if ($job.Shell) { $job.Shell.Stop() } } catch { }
        try { if ($job.Shell) { $job.Shell.Dispose() } } catch { }
        try { if ($job.Runspace) { $job.Runspace.Close(); $job.Runspace.Dispose() } } catch { }
    }

    try { $script:ApBackgroundJobs.Clear() }
    catch {
        # Replace rather than clear if the list itself is unusable.
        $script:ApBackgroundJobs = New-Object System.Collections.Generic.List[object]
    }
}
#endregion src\Private\Background.ps1

#region src\Private\XamlLoader.ps1
# XamlLoader.ps1 -- builds the WPF window from the XAML sources.
#
# Two things this handles that a naive loader does not:
#
# 1. Theme injection. Themes/Dark.xaml is authored as a standalone ResourceDictionary so it
#    can be edited and validated on its own, but StaticResource only resolves against
#    resources that already exist when the tree is built. Merging a dictionary into
#    Window.Resources after XamlReader.Load is too late, and referencing a file needs a
#    pack:// URI that does not exist for a loose script. So the dictionary's inner XML is
#    spliced into the window's <Window.Resources> at the @THEME@ token before parsing.
#
# 2. Named-element lookup via the real XAML namespace. The original GUI did
#    `$inputXML -replace "x:N", 'N'` and then selected on //*[@Name]
#    (Get-WindowsAutopilotImportGUI.ps1:113,128) -- a string hack that also rewrites any
#    attribute value containing "x:N". Here an XmlNamespaceManager is used instead.

$script:ApXamlNamespace = 'http://schemas.microsoft.com/winfx/2006/xaml'
$script:ApThemeToken = '<!-- @THEME@ -->'

function Get-ApXamlSource {
    <#
    .SYNOPSIS
    Returns the text of a bundled XAML file.

    .DESCRIPTION
    Prefers the payload embedded by build.ps1 (single-file dist), falling back to the
    file on disk (source tree).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('MainWindow', 'Dark')]
        [string]$Name
    )

    if ($script:ApEmbeddedXaml -and $script:ApEmbeddedXaml.ContainsKey($Name)) {
        return $script:ApEmbeddedXaml[$Name]
    }

    $relative = if ($Name -eq 'MainWindow') { 'src\Views\MainWindow.xaml' } else { 'src\Themes\Dark.xaml' }
    $path = Join-Path $script:ApAppRoot $relative

    if (-not (Test-Path -LiteralPath $path)) {
        throw "Could not find the XAML resource '$Name'. Expected it at $path or embedded in the built script."
    }

    return (Get-Content -LiteralPath $path -Raw)
}

function Get-ApMergedXaml {
    <#
    .SYNOPSIS
    The window XAML with the theme dictionary spliced in.
    #>
    [CmdletBinding()]
    param(
        [string]$WindowXaml,
        [string]$ThemeXaml
    )

    if (-not $WindowXaml) { $WindowXaml = Get-ApXamlSource -Name MainWindow }
    if (-not $ThemeXaml) { $ThemeXaml = Get-ApXamlSource -Name Dark }

    # Take the dictionary's children, discarding its own root element and xmlns declarations:
    # the window already declares the same default and x namespaces.
    $themeDoc = New-Object System.Xml.XmlDocument
    $themeDoc.PreserveWhitespace = $true
    $themeDoc.LoadXml($ThemeXaml)
    $themeInner = $themeDoc.DocumentElement.InnerXml

    if ($WindowXaml -notmatch [regex]::Escape($script:ApThemeToken)) {
        throw "MainWindow.xaml is missing the $script:ApThemeToken token, so the theme cannot be applied."
    }

    return $WindowXaml.Replace($script:ApThemeToken, $themeInner)
}

function New-ApMainWindow {
    <#
    .SYNOPSIS
    Loads the main window and returns it together with a lookup of its named elements.

    .OUTPUTS
    PSCustomObject with:
      Window    the Window instance
      Elements  hashtable of x:Name -> control
    #>
    [CmdletBinding()]
    param(
        [string]$Xaml
    )

    if (-not $Xaml) { $Xaml = Get-ApMergedXaml }

    $doc = New-Object System.Xml.XmlDocument
    try {
        $doc.LoadXml($Xaml)
    }
    catch {
        throw "The window definition is not valid XML: $($_.Exception.Message)"
    }

    $window = $null
    try {
        $reader = New-Object System.Xml.XmlNodeReader $doc
        $window = [Windows.Markup.XamlReader]::Load($reader)
    }
    catch {
        # XamlReader wraps the useful detail (element, line, position) in InnerException.
        $detail = $_.Exception.Message
        $inner = $_.Exception.InnerException
        while ($inner) {
            $detail += " -> $($inner.Message)"
            $inner = $inner.InnerException
        }
        throw "Failed to build the window from XAML: $detail"
    }

    # Resolve every x:Name into a hashtable so the wiring code can index by name
    # instead of repeating $window.FindName(...) for a hundred controls.
    $nsMgr = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $nsMgr.AddNamespace('x', $script:ApXamlNamespace)

    $elements = @{}
    foreach ($node in $doc.SelectNodes('//*[@x:Name]', $nsMgr)) {
        $name = $node.GetAttribute('Name', $script:ApXamlNamespace)
        if (-not $name) { continue }

        $control = $window.FindName($name)
        if ($control) { $elements[$name] = $control }
        else { Write-ApLog "XAML declares '$name' but FindName could not resolve it." -Level DEBUG }
    }

    Write-ApLog "Window loaded with $($elements.Count) named elements."

    return [pscustomobject]@{
        Window   = $window
        Elements = $elements
    }
}
#endregion src\Private\XamlLoader.ps1

#region src\Private\Dialogs.ps1
# Dialogs.ps1 -- themed modal dialogs and file pickers.
#
# The original used [System.Windows.MessageBox] for help and Windows Forms
# FolderBrowserDialog for exports (Get-WindowsAutopilotImportGUI.ps1:198-211,663). A stock
# MessageBox is a light-grey box in the middle of a dark app and cannot show selectable
# text, which matters for the "Preview command" dry-run: the whole point is to copy it.

function Show-ApDialog {
    <#
    .SYNOPSIS
    Themed modal dialog. Returns $true when confirmed.

    .PARAMETER Detail
    Optional monospaced, selectable, read-only block below the message. Used for the
    command preview and for error detail.

    .PARAMETER ConfirmText
    Label of the accept button. Omit -ShowCancel for a plain acknowledgement dialog.

    .PARAMETER Danger
    Style the accept button as destructive.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)][string]$Title,
        [string]$Message = '',
        [string]$Detail = '',
        [string]$ConfirmText = 'OK',
        [string]$CancelText = 'Cancel',
        [switch]$ShowCancel,
        [switch]$ShowCopy,
        [switch]$Danger,
        [int]$Width = 620,
        $Owner
    )

    $xamlText = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        SizeToContent="Height" ResizeMode="NoResize" ShowInTaskbar="False"
        WindowStartupLocation="CenterOwner"
        Background="#161616" Foreground="#FFFFFF"
        FontFamily="Segoe UI Variable, Segoe UI">
  <Window.Resources>
    <!-- @THEME@ -->
  </Window.Resources>
  <Grid Margin="24">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock x:Name="DlgTitle" Grid.Row="0" FontSize="18" FontWeight="SemiBold" TextWrapping="Wrap"/>
    <TextBlock x:Name="DlgMessage" Grid.Row="1" Foreground="#C0C0C0" FontSize="13" TextWrapping="Wrap" Margin="0,10,0,0"/>
    <TextBox   x:Name="DlgDetail" Grid.Row="2" Style="{StaticResource OutputBox}" Margin="0,14,0,0"
               MaxHeight="260" TextWrapping="Wrap" Visibility="Collapsed"/>
    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,20,0,0">
      <Button x:Name="DlgCopy"    Style="{StaticResource SecondaryButton}" Content="Copy" Margin="0,0,8,0" Visibility="Collapsed"/>
      <Button x:Name="DlgCancel"  Style="{StaticResource SecondaryButton}" Content="Cancel" Margin="0,0,8,0" Visibility="Collapsed"/>
      <Button x:Name="DlgConfirm" Style="{StaticResource PrimaryButton}" Content="OK" Height="36" MinWidth="110"/>
    </StackPanel>
  </Grid>
</Window>
'@

    $merged = Get-ApMergedXaml -WindowXaml $xamlText
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($merged)
    $dlg = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $doc))

    $dlg.Title = $Title
    $dlg.Width = $Width
    if ($Owner) {
        $dlg.Owner = $Owner
    }
    else {
        $dlg.WindowStartupLocation = 'CenterScreen'
    }

    $dlg.FindName('DlgTitle').Text = $Title

    $msgBlock = $dlg.FindName('DlgMessage')
    if ($Message) { $msgBlock.Text = $Message } else { $msgBlock.Visibility = 'Collapsed' }

    $detailBox = $dlg.FindName('DlgDetail')
    if ($Detail) {
        $detailBox.Text = $Detail
        $detailBox.Visibility = 'Visible'
    }

    $confirm = $dlg.FindName('DlgConfirm')
    $confirm.Content = $ConfirmText
    if ($Danger) {
        $confirm.Style = $dlg.FindResource('DangerButton')
        $confirm.Height = 36
    }

    $cancel = $dlg.FindName('DlgCancel')
    $cancel.Content = $CancelText
    if ($ShowCancel) { $cancel.Visibility = 'Visible' }

    $copy = $dlg.FindName('DlgCopy')
    if ($ShowCopy -and $Detail) {
        $copy.Visibility = 'Visible'
        $copy.Add_Click({
            try { Set-Clipboard -Value $detailBox.Text } catch { }
            $copy.Content = 'Copied'
        }.GetNewClosure())
    }

    # DialogResult is what ShowDialog returns; setting it closes the window.
    $confirm.Add_Click({ $dlg.DialogResult = $true }.GetNewClosure())
    $cancel.Add_Click({ $dlg.DialogResult = $false }.GetNewClosure())

    $confirm.IsDefault = $true
    $cancel.IsCancel = $true

    return [bool]($dlg.ShowDialog())
}

function Show-ApSaveFileDialog {
    <#
    .SYNOPSIS
    Save-file picker. Returns the chosen path, or $null if cancelled.

    .DESCRIPTION
    Uses Microsoft.Win32.SaveFileDialog (WPF's own, already loaded) rather than the Windows
    Forms FolderBrowserDialog the original used, so the tech can name the file and pick a
    location in one step instead of being handed a fixed AutopilotHWID.csv.
    #>
    [CmdletBinding()]
    param(
        [string]$Title = 'Save file',
        [string]$FileName = 'AutopilotHWID.csv',
        [string]$Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*',
        [string]$InitialDirectory
    )

    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Title = $Title
    $dialog.FileName = $FileName
    $dialog.Filter = $Filter
    $dialog.AddExtension = $true
    $dialog.OverwritePrompt = $false   # -Append is a supported workflow; do not fight it

    if ($InitialDirectory -and (Test-Path -LiteralPath $InitialDirectory)) {
        $dialog.InitialDirectory = $InitialDirectory
    }

    if ($dialog.ShowDialog()) { return $dialog.FileName }
    return $null
}
#endregion src\Private\Dialogs.ps1

#region src\Public\Show-AutopilotImportGui.ps1
# Show-AutopilotImportGui.ps1 -- builds the window and wires every control to behaviour.
#
# Shared state lives in $script: variables rather than function locals because WPF event
# handlers are scriptblocks that execute later, outside the defining function's scope.

$script:ApWin = $null
$script:ApEl = $null
$script:ApDevice = $null
$script:ApRun = $null
$script:ApRunTimer = $null
$script:ApClockTimer = $null
$script:ApActiveOutput = $null
$script:ApNetworkResults = @()
$script:ApGraphCheck = $null
# Set when a v2 register run is launched with the restart option ticked; consumed once the
# run completes successfully.
$script:ApPendingV2Reboot = $false
# The Logs pane has two possible owners: the GUI's own session log, and the streamed output of
# a run launched from the Advanced page (diagnostics). Tracking which one is on screen stops
# the session log from overwriting a report the operator is still reading.
$script:ApLogsShowingRun = $false
$script:ApEnginePath = $null
$script:ApAppVersion = '1.3.0'
$script:ApAuthor = 'Mark Orr'
$script:ApAuthorHandle = '@markorr321'
$script:ApAuthorSite = 'https://orr365.tools'
$script:ApAuthorGitHub = 'https://github.com/markorr321'

# ============================ small UI helpers ============================

function New-ApBrush {
    param([Parameter(Mandatory)][string]$Hex)
    return New-Object System.Windows.Media.SolidColorBrush(
        [System.Windows.Media.ColorConverter]::ConvertFromString($Hex))
}

function Set-ApPill {
    <#
    .SYNOPSIS
    Sets a status chip's text and colour scheme.

    .PARAMETER State
    ok | warn | error | neutral
    #>
    param(
        [Parameter(Mandatory)]$Border,
        [Parameter(Mandatory)]$TextBlock,
        [Parameter(Mandatory)][string]$Text,
        [ValidateSet('ok', 'warn', 'error', 'neutral')][string]$State = 'neutral'
    )

    # Tinted backgrounds rather than saturated fills: a row of bright chips in a dark UI
    # reads as alarming even when everything is fine.
    #
    # The locals are named *Colour, not $border: PowerShell variable names are
    # case-insensitive, so a local $border would silently overwrite the $Border parameter
    # and the assignments below would fail against a string.
    switch ($State) {
        'ok'    { $bgColour = '#12291B'; $borderColour = '#1E5B33'; $fgColour = '#3BC77A' }
        'warn'  { $bgColour = '#2A2110'; $borderColour = '#5B4718'; $fgColour = '#E9B44C' }
        'error' { $bgColour = '#2A1417'; $borderColour = '#6B2027'; $fgColour = '#F03A47' }
        default { $bgColour = '#252525'; $borderColour = '#3A3A3A'; $fgColour = '#C0C0C0' }
    }

    $Border.Background = New-ApBrush $bgColour
    $Border.BorderBrush = New-ApBrush $borderColour
    $TextBlock.Foreground = New-ApBrush $fgColour
    $TextBlock.Text = $Text
}

function Start-ApUrl {
    <#
    .SYNOPSIS
    Opens a URL in the default browser.

    .DESCRIPTION
    Only http and https are accepted. Handing an arbitrary string to Start-Process would
    happily launch a local executable or a file association, which is not what a link in a
    UI should ever do.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Url)

    $parsed = $null
    if (-not [System.Uri]::TryCreate($Url, [System.UriKind]::Absolute, [ref]$parsed) -or
        $parsed.Scheme -notin @('http', 'https')) {
        Write-ApLog "Refusing to open a non-web link: $Url" -Level WARN
        return
    }

    try {
        Start-Process $parsed.AbsoluteUri | Out-Null
        Write-ApLog "Opened $($parsed.AbsoluteUri)"
    }
    catch {
        Set-ApStatus -Text "Could not open the link: $($_.Exception.Message)" -IsError
    }
}

function Set-ApStatus {
    <#
    .SYNOPSIS
    Updates the bottom status strip.
    #>
    param(
        [string]$Text,
        [string]$Stage,
        [Nullable[int]]$Percent,
        [switch]$IsError
    )

    if ($PSBoundParameters.ContainsKey('Text')) {
        $script:ApEl.StatusText.Text = $Text
        $script:ApEl.StatusText.Foreground = if ($IsError) { New-ApBrush '#F03A47' } else { New-ApBrush '#C0C0C0' }
    }
    if ($PSBoundParameters.ContainsKey('Stage')) {
        $script:ApEl.StatusStage.Text = $Stage
    }
    if ($null -ne $Percent) {
        $script:ApEl.StatusProgress.IsIndeterminate = $false
        $script:ApEl.StatusProgress.Value = $Percent
        $script:ApEl.StatusProgress.Foreground = if ($IsError) { New-ApBrush '#F03A47' } else { New-ApBrush '#0078D4' }
    }
}

function Add-ApOutput {
    <#
    .SYNOPSIS
    Appends timestamped lines to an output box and keeps it scrolled to the end.
    #>
    param(
        $Box,
        [string[]]$Lines
    )

    if (-not $Box) { return }
    if (-not $Lines -or $Lines.Count -eq 0) { return }

    # Follow the tail only while the view is already at the bottom. This used to call
    # ScrollToEnd() unconditionally, and because a run appends every 250ms, scrolling up to
    # read an earlier line was undone before it could be read. Measured before appending:
    # afterwards ExtentHeight has already grown and nothing looks like the bottom.
    # A 2px tolerance covers the fractional offsets a partially scrolled line leaves behind.
    $followTail = $true
    try {
        if ($Box.ExtentHeight -gt $Box.ViewportHeight) {
            $followTail = ($Box.VerticalOffset + $Box.ViewportHeight) -ge ($Box.ExtentHeight - 2)
        }
    }
    catch {
        # No layout yet (the box has never been shown), so the tail is trivially in view.
    }

    $stamp = Get-Date -Format 'HH:mm:ss'
    $text = ($Lines | ForEach-Object { "$stamp  $_" }) -join [Environment]::NewLine

    if ($Box.Text.Length -gt 0) { $Box.AppendText([Environment]::NewLine) }
    $Box.AppendText($text)
    if ($followTail) { $Box.ScrollToEnd() }
}

function Show-ApPage {
    <#
    .SYNOPSIS
    Shows one page and hides the rest.
    #>
    param([Parameter(Mandatory)][string]$Name)

    foreach ($page in @('PageRegister', 'PageDevice', 'PageNetwork', 'PageAdvanced', 'PageLogs')) {
        $script:ApEl[$page].Visibility = if ($page -eq $Name) { 'Visible' } else { 'Collapsed' }
    }
}

function Show-ApNotice {
    <#
    .SYNOPSIS
    Shows the warning card on the Register page.
    #>
    param(
        [string]$Title = 'Before you continue',
        [Parameter(Mandatory)][string[]]$Messages
    )

    $messages = @($Messages | Where-Object { $_ })
    if ($messages.Count -eq 0) { Hide-ApNotice; return }

    $script:ApEl.NoticeTitle.Text = $Title
    $script:ApEl.NoticeText.Text = ($messages | ForEach-Object { "- $_" }) -join [Environment]::NewLine
    $script:ApEl.NoticeCard.Visibility = 'Visible'
}

function Hide-ApNotice {
    $script:ApEl.NoticeCard.Visibility = 'Collapsed'
}

# ============================ device page ============================

# Functions Get-ApDeviceInfo needs once it is copied into a bare runspace. Get-ApCimSafe is
# deliberately log-free so this list stays closed.
$script:ApDeviceInfoFunctions = @(
    'Get-ApCimSafe'
    'Get-ApDeviceIdentifierPart'
    'Test-ApHardwareHashAvailable'
    'Get-ApTpmInfo'
    'Test-ApSecureBoot'
    'Get-ApWindowsDisplayVersion'
    'Test-ApVirtualMachine'
    'Test-ApElevated'
    'Get-ApDeviceInfo'
)

function Start-ApDeviceLoad {
    <#
    .SYNOPSIS
    Reads local hardware on a background runspace, then applies it to the UI.

    .DESCRIPTION
    Two of the queries behind Get-ApDeviceInfo block for about five seconds each when the
    process is not elevated, so this must not run on the UI thread: doing so kept the window
    from appearing for over ten seconds. The panes show "reading..." until it lands.
    #>
    [CmdletBinding()]
    param()

    $el = $script:ApEl
    $el.RefreshDeviceButton.IsEnabled = $false
    Set-ApStatus -Text 'Reading device information...'

    Start-ApBackgroundWork -Window $script:ApWin `
        -FunctionNames $script:ApDeviceInfoFunctions `
        -Work { Get-ApDeviceInfo } `
        -OnComplete {
            param($result, $errorMessage)

            $script:ApEl.RefreshDeviceButton.IsEnabled = $true

            if ($errorMessage -or -not $result) {
                Set-ApStatus -Text "Could not read device information: $errorMessage" -IsError
                Write-ApLog "Device information failed: $errorMessage" -Level ERROR
                return
            }

            Set-ApDeviceUi (@($result)[-1])
            Sync-ApModeUi

            # A machine with no hash cannot do v1, so switch unless the caller pinned a mode.
            if (-not $script:ApModePinned -and -not $script:ApDevice.HardwareHashReady -and $script:ApDevice.IsElevated) {
                $script:ApEl.ModeV2.IsChecked = $true
                Write-ApLog 'No hardware hash available; defaulting to Device Preparation (v2).'
            }

            Set-ApStatus -Text 'Ready.' -Percent 0
        } | Out-Null
}

function Set-ApDeviceUi {
    <#
    .SYNOPSIS
    Applies a device-info snapshot to the summary, readiness chips and inventory grid.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]$Info
    )

    $info = $Info
    $script:ApDevice = $info

    $el = $script:ApEl
    $el.SummarySerial.Text = if ($info.SerialNumber) { $info.SerialNumber } else { 'unavailable' }
    $el.SummaryModel.Text = if ($info.Model) { $info.Model } else { 'unavailable' }
    $el.SummaryManufacturer.Text = if ($info.Manufacturer) { $info.Manufacturer } else { 'unavailable' }
    $el.SummaryFreeSpace.Text = if ($null -ne $info.FreeSpaceGb) { "$($info.FreeSpaceGb) GB" } else { 'unknown' }

    # Hardware hash: the hard prerequisite for v1. Virtual machines usually cannot supply
    # one, which is exactly why Device Preparation (v2) exists.
    if ($info.HardwareHashReady) {
        Set-ApPill $el.PillHash $el.PillHashText 'Hash ready' 'ok'
    }
    elseif (-not $info.IsElevated) {
        Set-ApPill $el.PillHash $el.PillHashText 'Hash needs admin' 'error'
    }
    else {
        Set-ApPill $el.PillHash $el.PillHashText 'No hash' 'error'
    }

    # Both the TPM class and Confirm-SecureBootUEFI need administrator rights. Without them
    # the query fails the same way a genuinely absent TPM does, so report "unknown" rather
    # than asserting hardware facts we could not actually read.
    if ($info.TpmPresent -and $info.TpmEnabled) {
        Set-ApPill $el.PillTpm $el.PillTpmText "TPM $($info.TpmSpecVersion)" 'ok'
    }
    elseif ($info.TpmPresent) {
        Set-ApPill $el.PillTpm $el.PillTpmText 'TPM disabled' 'warn'
    }
    elseif (-not $info.IsElevated) {
        Set-ApPill $el.PillTpm $el.PillTpmText 'TPM unknown' 'warn'
    }
    else {
        Set-ApPill $el.PillTpm $el.PillTpmText 'No TPM' 'warn'
    }

    if ($info.SecureBoot -eq $true) {
        Set-ApPill $el.PillSecureBoot $el.PillSecureBootText 'Secure Boot on' 'ok'
    }
    elseif ($info.SecureBoot -eq $false) {
        Set-ApPill $el.PillSecureBoot $el.PillSecureBootText 'Secure Boot off' 'warn'
    }
    elseif (-not $info.IsElevated) {
        Set-ApPill $el.PillSecureBoot $el.PillSecureBootText 'Secure Boot unknown' 'warn'
    }
    else {
        Set-ApPill $el.PillSecureBoot $el.PillSecureBootText 'Legacy BIOS' 'warn'
    }

    $el.IdentifierPreviewBox.Text = $info.DeviceIdentifier

    $rows = @(
        [pscustomobject]@{ Name = 'Serial number';   Value = $info.SerialNumber }
        [pscustomobject]@{ Name = 'Manufacturer';    Value = $info.Manufacturer }
        [pscustomobject]@{ Name = 'Model';           Value = $info.Model }
        [pscustomobject]@{ Name = 'Computer name';   Value = $info.ComputerName }
        [pscustomobject]@{ Name = 'Operating system'; Value = $info.OsCaption }
        [pscustomobject]@{ Name = 'OS build';        Value = $info.OsBuild }
        [pscustomobject]@{ Name = 'OS version';      Value = $info.OsDisplayVersion }
        [pscustomobject]@{ Name = 'Memory';          Value = if ($null -ne $info.MemoryGb) { "$($info.MemoryGb) GB" } else { 'unknown' } }
        [pscustomobject]@{ Name = "Disk $($info.SystemDrive)"; Value = if ($null -ne $info.FreeSpaceGb) { "$($info.FreeSpaceGb) GB free of $($info.TotalSpaceGb) GB" } else { 'unknown' } }
        [pscustomobject]@{ Name = 'TPM';             Value = if ($info.TpmPresent) { "version $($info.TpmSpecVersion), enabled: $($info.TpmEnabled)" } elseif (-not $info.IsElevated) { 'unknown (requires administrator)' } else { 'not detected' } }
        [pscustomobject]@{ Name = 'Secure Boot';     Value = if ($null -ne $info.SecureBoot) { "$($info.SecureBoot)" } elseif (-not $info.IsElevated) { 'unknown (requires administrator)' } else { 'not applicable (legacy BIOS)' } }
        [pscustomobject]@{ Name = 'Hardware hash';   Value = if ($info.HardwareHashReady) { 'available' } elseif (-not $info.IsElevated) { 'unknown (requires administrator)' } else { 'not available' } }
        [pscustomobject]@{ Name = 'Virtual machine'; Value = "$($info.IsVirtualMachine)" }
        [pscustomobject]@{ Name = 'Running elevated'; Value = "$($info.IsElevated)" }
        [pscustomobject]@{ Name = 'Device identifier (v2)'; Value = $info.DeviceIdentifier }
    )
    $el.DeviceGrid.ItemsSource = $rows

    Write-ApLog "Device: $($info.Manufacturer) $($info.Model) serial $($info.SerialNumber); hash ready: $($info.HardwareHashReady)"
    return $info
}

function Set-ApDevicePlaceholders {
    <#
    .SYNOPSIS
    Neutral "reading" state shown until the background device load completes.
    #>
    [CmdletBinding()]
    param()

    $el = $script:ApEl
    foreach ($name in @('SummarySerial', 'SummaryModel', 'SummaryManufacturer', 'SummaryFreeSpace')) {
        $el[$name].Text = 'reading...'
    }
    Set-ApPill $el.PillHash $el.PillHashText 'Hash ...' 'neutral'
    Set-ApPill $el.PillTpm $el.PillTpmText 'TPM ...' 'neutral'
    Set-ApPill $el.PillSecureBoot $el.PillSecureBootText 'Secure Boot ...' 'neutral'
}

# ============================ mode handling ============================

function Sync-ApModeUi {
    <#
    .SYNOPSIS
    Enables or disables controls according to the selected registration mode.

    .DESCRIPTION
    Autopilot v2 (-identifier) runs a separate branch in the community script that ignores
    group tag, assigned user, computer name, Entra group, assignment wait and reboot
    (get-windowsautopilotinfocommunity.ps1:2233-2270). Greying those controls out is
    honest; leaving them enabled would imply they do something.
    #>
    [CmdletBinding()]
    param()

    $el = $script:ApEl
    $isV2 = [bool]$el.ModeV2.IsChecked

    foreach ($name in @('GroupTagCombo', 'AssignedUserBox', 'ComputerNameBox', 'AddToGroupBox')) {
        $el[$name].IsEnabled = -not $isV2
    }

    # The group tag is hidden rather than merely dimmed in v2: it is the one field
    # technicians reach for by habit, and Device Preparation targets devices through the
    # Entra group on the policy instead, so leaving it on screen invites a wasted entry.
    $el.GroupTagSection.Visibility = if ($isV2) { 'Collapsed' } else { 'Visible' }

    # Every field in this card (group tag, assigned user, computer name, Entra group) is
    # ignored by the identifier path, so in v2 the whole card goes. Dimming it was worse than
    # useless: three dead fields pushed the one live v2 option, the restart, below the fold.
    $el.CardDetails.Visibility = if ($isV2) { 'Collapsed' } else { 'Visible' }

    # Swap the whole options block rather than dimming it. The v1 controls map to engine
    # switches the identifier path ignores; the v2 restart is performed by this tool.
    $el.OptionsV1Section.Visibility = if ($isV2) { 'Collapsed' } else { 'Visible' }
    $el.OptionsV2Section.Visibility = if ($isV2) { 'Visible' } else { 'Collapsed' }

    # v1 delegates its restart to the engine's -Reboot, so the on-demand button is a v2 affair.
    $el.RestartNowButton.Visibility = if ($isV2) { 'Visible' } else { 'Collapsed' }

    $el.CardDetails.Opacity = if ($isV2) { 0.55 } else { 1.0 }

    $el.IdentifierPreviewLabel.Visibility = if ($isV2) { 'Visible' } else { 'Collapsed' }
    $el.IdentifierPreviewBox.Visibility = if ($isV2) { 'Visible' } else { 'Collapsed' }

    if ($isV2) {
        $el.ModeDescription.Text = 'Imports the Manufacturer,Model,Serial identifier for Windows Autopilot Device Preparation. No hardware hash is needed, so this works on virtual machines. Devices are targeted by the Entra group on your Device Preparation policy, not by a group tag.'
        Show-ApNotice -Title 'Device Preparation mode' -Messages @(
            'Group tag, assigned user, computer name and profile assignment do not apply to v2.',
            'After importing, add this device to the Entra security group used by your Device Preparation policy.'
        )
    }
    else {
        $el.ModeDescription.Text = 'Uploads the 4K hardware hash to Windows Autopilot, then optionally waits for a deployment profile to be assigned.'
        Hide-ApNotice

        if ($script:ApDevice -and -not $script:ApDevice.HardwareHashReady) {
            $reason = if (-not $script:ApDevice.IsElevated) {
                'This window is not running as administrator, so the hardware hash cannot be read.'
            }
            elseif ($script:ApDevice.IsVirtualMachine) {
                'This looks like a virtual machine, and VMs frequently cannot provide a hardware hash. Device Preparation (v2) is usually the right choice here.'
            }
            else {
                'This device did not return a hardware hash.'
            }
            Show-ApNotice -Title 'Hardware hash unavailable' -Messages @($reason)
        }
    }
}

# ============================ request assembly ============================

function Get-ApUiRequest {
    <#
    .SYNOPSIS
    Reads the current UI state into a registration request.
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('Register', 'Export')][string]$Operation = 'Register'
    )

    $el = $script:ApEl
    $request = New-ApRegistrationRequest
    $request.Operation = $Operation

    $request.Mode = if ($el.ModeV2.IsChecked) { 'v2' } else { 'v1' }

    if ($Operation -eq 'Export') {
        $request.Append = [bool]$el.ExportAppendCheck.IsChecked
        $request.Partner = [bool]$el.ExportPartnerCheck.IsChecked
    }

    # The editable ComboBox's Text is the authoritative value: the tech may have typed a new
    # tag rather than picked a remembered one.
    $request.GroupTag = "$($el.GroupTagCombo.Text)"
    $request.AssignedUser = $el.AssignedUserBox.Text
    $request.AssignedComputerName = $el.ComputerNameBox.Text
    $request.AddToGroup = $el.AddToGroupBox.Text

    if ($Operation -eq 'Register') {
        $request.WaitForAssignment = [bool]$el.WaitAssignCheck.IsChecked
        $request.Reboot = [bool]$el.RebootCheck.IsChecked
        $request.Wipe = [bool]$el.AdvWipeCheck.IsChecked
        $request.Sysprep = [bool]$el.AdvSysprepCheck.IsChecked
        $request.PreProvision = [bool]$el.AdvPreProvisionCheck.IsChecked
        $request.ChangePK = $el.AdvChangePkBox.Text

        $request.ExistingDevicePolicy =
            if ($el.PolicyDelete.IsChecked) { 'delete' }
            elseif ($el.PolicySkip.IsChecked) { 'skipcheck' }
            else { 'update' }
    }

    return $request
}

function Test-ApDestructiveRequest {
    <#
    .SYNOPSIS
    True when the request does something the tech should confirm first.
    #>
    param([Parameter(Mandatory)]$Request)

    return [bool](
        $Request.Wipe -or
        $Request.Sysprep -or
        $Request.ChangePK -or
        ($Request.ExistingDevicePolicy -eq 'delete')
    )
}

# ============================ run lifecycle ============================

function Set-ApRunningState {
    <#
    .SYNOPSIS
    Enables or disables the controls that must not change mid-run.
    #>
    param([bool]$IsRunning)

    $el = $script:ApEl
    $el.RegisterButton.IsEnabled = -not $IsRunning
    $el.PreviewButton.IsEnabled = -not $IsRunning
    $el.AdvDiagnosticsButton.IsEnabled = -not $IsRunning
    $el.AdvDiagnosticsOnlineCheck.IsEnabled = -not $IsRunning
    $el.AdvWindowsUpdateButton.IsEnabled = -not $IsRunning
    $el.ExportHashButton.IsEnabled = -not $IsRunning
    $el.ExportIdentifierButton.IsEnabled = -not $IsRunning
    $el.CancelButton.IsEnabled = $IsRunning

    $el.RegisterButton.Content = if ($IsRunning) { 'REGISTERING...' } else { 'REGISTER THIS DEVICE' }
}

function Start-ApGuiRun {
    <#
    .SYNOPSIS
    Starts an engine run and begins streaming its output into the given box.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Parameters,
        [Parameter(Mandatory)]$OutputBox,
        [string]$Label = 'run',
        [string]$StartMessage = 'Starting...',
        # Defaults to the resolved community engine; the diagnostics script passes its own.
        [string]$EnginePath
    )

    if (-not $EnginePath) { $EnginePath = $script:ApEnginePath }

    if ($script:ApRun -and -not $script:ApRun.IsFinished) {
        Show-ApDialog -Title 'A run is already in progress' -Owner $script:ApWin `
                      -Message 'Wait for the current operation to finish, or cancel it first.' | Out-Null
        return
    }

    $config = Get-ApConfig

    try {
        $script:ApRun = Start-ApEngineRun -Parameters $Parameters `
                                          -EnginePath $EnginePath `
                                          -Label $Label `
                                          -ShowConsole:([bool]$config.showConsoleWindow)
    }
    catch {
        Set-ApStatus -Text "Could not start: $($_.Exception.Message)" -Percent 0 -IsError
        Show-ApDialog -Title 'Could not start the engine' -Owner $script:ApWin `
                      -Message $_.Exception.Message | Out-Null
        return
    }

    $script:ApActiveOutput = $OutputBox
    # A run that streams into the Logs pane owns it until the operator asks for the session
    # log back, otherwise finishing the run would wipe its own output.
    $script:ApLogsShowingRun = [bool]($OutputBox -eq $script:ApEl.LogsOutput)
    $OutputBox.Clear()
    Add-ApOutput -Box $OutputBox -Lines @($StartMessage)

    Set-ApRunningState $true
    Set-ApStatus -Text $StartMessage -Stage 'Connect' -Percent 0

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(250)
    $timer.Add_Tick({
        try {
            Invoke-ApRunTick
        }
        catch {
            # Never let a tick exception escape: it would propagate out of ShowDialog.
            if ($script:ApRunTimer) { $script:ApRunTimer.Stop(); $script:ApRunTimer = $null }
            Set-ApRunningState $false
            Set-ApStatus -Text "Lost track of the run: $($_.Exception.Message)" -IsError
            Write-ApLog "Run tick failed: $($_.Exception.Message)" -Level ERROR
            if ($_.ScriptStackTrace) { Write-ApLog $_.ScriptStackTrace -Level DEBUG }
        }
    })
    $script:ApRunTimer = $timer
    $timer.Start()
}

function Invoke-ApRunTick {
    <#
    .SYNOPSIS
    DispatcherTimer tick: drain new output, update progress, finalise when done.
    #>
    [CmdletBinding()]
    param()

    $run = $script:ApRun
    if (-not $run) { return }

    $update = Update-ApEngineRun -Run $run

    if ($update.NewLines.Count -gt 0 -and $script:ApActiveOutput) {
        Add-ApOutput -Box $script:ApActiveOutput -Lines $update.NewLines
    }

    $state = $run.ProgressState
    if ($update.ProgressChanged) {
        Set-ApStatus -Text $state.StageLabel -Stage $state.Stage -Percent $state.Percent -IsError:([bool]$state.IsError)
    }

    if ($update.Stalled) {
        $hint = 'No output for 90 seconds. The engine may be waiting on a prompt in its hidden console. Turn on "Show the engine console window" on the Advanced page, then cancel and retry.'
        Set-ApStatus -Text $hint -IsError
        Add-ApOutput -Box $script:ApActiveOutput -Lines @('', "[GUI] $hint")
    }

    if (-not $update.IsFinished) { return }

    if ($script:ApRunTimer) { $script:ApRunTimer.Stop(); $script:ApRunTimer = $null }
    Set-ApRunningState $false

    $summary = Get-ApRunSummary -Run $run
    $failed = $run.Cancelled -or $state.IsError -or ($run.ExitCode -ne 0)

    Set-ApStatus -Text $summary -Stage $state.Stage -Percent ($(if ($failed) { $state.Percent } else { 100 })) -IsError:$failed
    Add-ApOutput -Box $script:ApActiveOutput -Lines @('', $summary, "[GUI] Full output saved to $($run.LogPath)")
    Write-ApLog "Run summary: $summary" -Level $(if ($failed) { 'WARN' } else { 'INFO' })

    # Refreshes the log path caption. It deliberately leaves a run report on screen; see
    # Update-ApLogsPage for why that mattered.
    Update-ApLogsPage

    # A finished report is read from the top, but streaming left the view pinned to the tail.
    # Only for runs that own the Logs pane: a registration ends with its outcome, so there the
    # last line is the one worth looking at.
    if ($script:ApLogsShowingRun -and $script:ApActiveOutput) {
        $script:ApActiveOutput.CaretIndex = 0
        $script:ApActiveOutput.ScrollToHome()
    }

    # Device Preparation restart. Only on a clean run: never reboot after a failure or a
    # cancel, or the operator loses the log and the device leaves OOBE unregistered.
    if ($script:ApPendingV2Reboot -and -not $failed -and $state.IsComplete) {
        $script:ApPendingV2Reboot = $false
        Add-ApOutput -Box $script:ApActiveOutput -Lines @('', '[GUI] Identifier imported. Restarting this device now.')
        Set-ApStatus -Text 'Identifier imported. Restarting...' -Percent 100
        Invoke-ApRestartComputer | Out-Null
        return
    }
    $script:ApPendingV2Reboot = $false

    if ($state.IsError) {
        Show-ApDialog -Title 'The run reported a problem' -Owner $script:ApWin `
                      -Message $state.ErrorMessage `
                      -Detail "Full log:`r`n$($run.LogPath)" | Out-Null
    }
}

function Invoke-ApRestartComputer {
    <#
    .SYNOPSIS
    Restarts this machine.

    .DESCRIPTION
    Used for the Device Preparation (v2) restart option. Autopilot v1 delegates its restart to
    the engine's -Reboot switch, but that switch is nested inside the engine's
    assignment-wait block, which the -identifier path never reaches. So for v2 the restart has
    to happen here.

    Restart-Computer comes from Microsoft.PowerShell.Management, which Windows PowerShell loads
    at startup rather than auto-loading, so it is unaffected by the module-shadowing problems
    that affect the engine process. shutdown.exe is kept as a fallback anyway: it is a native
    binary and needs no module resolution at all.
    #>
    [CmdletBinding()]
    param()

    Write-ApLog 'Restarting the computer at the operator''s request.' -Level WARN

    try {
        Restart-Computer -Force -ErrorAction Stop
        return $true
    }
    catch {
        Write-ApLog "Restart-Computer failed: $($_.Exception.Message). Falling back to shutdown.exe." -Level WARN
        try {
            Start-Process -FilePath (Join-Path $env:WINDIR 'System32\shutdown.exe') `
                          -ArgumentList '/r', '/t', '0' -WindowStyle Hidden -ErrorAction Stop | Out-Null
            return $true
        }
        catch {
            Write-ApLog "shutdown.exe also failed: $($_.Exception.Message)" -Level ERROR
            return $false
        }
    }
}

function Stop-ApGuiRun {
    [CmdletBinding()]
    param()

    if (-not $script:ApRun -or $script:ApRun.IsFinished) { return }

    Stop-ApEngineRun -Run $script:ApRun

    if ($script:ApRunTimer) { $script:ApRunTimer.Stop(); $script:ApRunTimer = $null }
    Set-ApRunningState $false
    Set-ApStatus -Text 'Cancelled.' -Stage '' -Percent 0 -IsError
    Add-ApOutput -Box $script:ApActiveOutput -Lines @('', 'Cancelled by the operator.')
}

# ============================ logs page ============================

function Update-ApLogsPage {
    <#
    .SYNOPSIS
    Shows the GUI session log in the Logs pane.

    .DESCRIPTION
    Refuses to overwrite streamed run output unless asked to. A diagnostics run streams into
    this same pane, and this function used to be called unconditionally when any run finished:
    the report was replaced by a handful of session-log lines the instant it completed, which
    looked like the results -- and the scrollbar with them -- simply vanishing.

    .PARAMETER Force
    Replace the pane even when it is showing run output. This is what the Refresh button does,
    because asking for the session log explicitly is unambiguous.
    #>
    [CmdletBinding()]
    param([switch]$Force)

    $script:ApEl.LogsPathText.Text = Get-ApLogPath

    if ($script:ApLogsShowingRun -and -not $Force) { return }
    $script:ApLogsShowingRun = $false

    $lines = Get-ApLogBuffer
    $script:ApEl.LogsOutput.Text = ($lines -join [Environment]::NewLine)
    $script:ApEl.LogsOutput.ScrollToEnd()
}

# ============================ network page ============================

function Start-ApNetworkCheck {
    [CmdletBinding()]
    param()

    $el = $script:ApEl
    $el.NetworkRunButton.IsEnabled = $false
    $el.NetworkSummary.Text = 'Testing...'
    $el.NetworkGrid.ItemsSource = $null

    $endpoints = Get-ApConfiguredEndpoints
    Write-ApLog "Starting connectivity check across $($endpoints.Count) endpoints"
    Set-ApStatus -Text "Testing $($endpoints.Count) endpoints..." -Stage 'Network'
    $el.StatusProgress.IsIndeterminate = $true

    Start-ApBackgroundWork -Window $script:ApWin `
        -Variables @{ ApEndpoints = $endpoints } `
        -FunctionNames @('Test-ApEndpoint', 'Invoke-ApConnectivityCheck') `
        -Work {
            Invoke-ApConnectivityCheck -Endpoints $ApEndpoints
        } `
        -OnComplete {
            param($result, $errorMessage)

            $el = $script:ApEl
            $el.NetworkRunButton.IsEnabled = $true
            $el.StatusProgress.IsIndeterminate = $false

            if ($errorMessage) {
                $el.NetworkSummary.Text = "Check failed: $errorMessage"
                Set-ApStatus -Text "Connectivity check failed: $errorMessage" -Percent 0 -IsError
                Write-ApLog "Connectivity check failed: $errorMessage" -Level ERROR
                return
            }

            $results = @($result)
            $script:ApNetworkResults = $results

            # Sort failures to the top: that is what the tech opened this page to see.
            $el.NetworkGrid.ItemsSource = @(
                $results | Sort-Object @{ Expression = { $_.Succeeded } }, @{ Expression = { -not $_.Required } }, Category, Name
            )

            $ok = @($results | Where-Object { $_.Succeeded }).Count
            $requiredFailures = @($results | Where-Object { $_.Required -and -not $_.Succeeded })

            if ($requiredFailures.Count -gt 0) {
                $names = ($requiredFailures | ForEach-Object { $_.Host }) -join ', '
                $el.NetworkSummary.Text = "$ok of $($results.Count) reachable. Required endpoints failing: $names"
                Set-ApStatus -Text "$($requiredFailures.Count) required endpoint(s) unreachable." -Percent 100 -IsError
            }
            else {
                $optionalFailures = @($results | Where-Object { -not $_.Succeeded }).Count
                $suffix = if ($optionalFailures -gt 0) { " $optionalFailures optional endpoint(s) unreachable." } else { '' }
                $el.NetworkSummary.Text = "$ok of $($results.Count) reachable. All required endpoints are available.$suffix"
                Set-ApStatus -Text 'Connectivity check complete.' -Percent 100
            }

            Write-ApLog "Connectivity check finished: $ok of $($results.Count) reachable, $($requiredFailures.Count) required failing"
        } | Out-Null
}

# ============================ export helpers ============================

function Start-ApCsvExport {
    <#
    .SYNOPSIS
    Runs an offline CSV export (no tenant connection).

    .NOTES
    Named this way deliberately. The obvious "Invoke-Ap" + "Export" spelling produces an
    identifier whose lowercased form starts with the same four letters as a Microsoft
    Defender HackTool signature family, and Defender's script scanner then refuses to load
    the entire file with ScriptContainedMaliciousContent. Verified by elimination: a file
    containing only that one function was blocked while all 22 other functions in this file
    loaded cleanly. Keep the current name.
    #>
    [CmdletBinding()]
    param([ValidateSet('v1', 'v2')][string]$Mode)

    $defaultName = if ($Mode -eq 'v2') { 'AutopilotDeviceIdentifier.csv' } else { 'AutopilotHWID.csv' }
    $path = Show-ApSaveFileDialog -Title 'Save the Autopilot CSV' -FileName $defaultName
    if (-not $path) { return }

    $request = Get-ApUiRequest -Operation Export
    $request.Mode = $Mode
    $request.OutputFile = $path

    try {
        $built = Build-ApEngineArguments $request
    }
    catch {
        Show-ApDialog -Title 'Cannot export' -Owner $script:ApWin -Message $_.Exception.Message | Out-Null
        return
    }

    if ($built.Warnings.Count -gt 0) {
        Show-ApDialog -Title 'Export note' -Owner $script:ApWin `
                      -Message ($built.Warnings -join [Environment]::NewLine) | Out-Null
    }

    Show-ApPage 'PageRegister'
    $script:ApEl.NavRegister.IsChecked = $true

    Start-ApGuiRun -Parameters $built.Parameters -OutputBox $script:ApEl.RegisterOutput `
                   -Label 'export' -StartMessage "Exporting to $path..."
}

# ============================ entry point ============================

function Show-AutopilotImportGui {
    <#
    .SYNOPSIS
    Builds and shows the Autopilot Import GUI.

    .PARAMETER GroupTag
    Pre-fills the group tag field.

    .PARAMETER AssignedUser
    Pre-fills the assigned user field.

    .PARAMETER Mode
    Pre-selects the registration mode.
    #>
    [CmdletBinding()]
    param(
        [string]$GroupTag,
        [string]$AssignedUser,
        [ValidateSet('v1', 'v2')][string]$Mode
    )

    Initialize-ApGui @PSBoundParameters | Out-Null
    [void]$script:ApWin.ShowDialog()
}

function Initialize-ApGui {
    <#
    .SYNOPSIS
    Builds the window, populates it and wires every handler. Does not show it.

    .DESCRIPTION
    Separated from Show-AutopilotImportGui so the whole UI can be constructed and driven in
    a test harness without entering the blocking ShowDialog message loop.

    .OUTPUTS
    The constructed Window.
    #>
    [CmdletBinding()]
    param(
        [string]$GroupTag,
        [string]$AssignedUser,
        [ValidateSet('v1', 'v2')][string]$Mode
    )

    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml

    $ui = New-ApMainWindow
    $script:ApWin = $ui.Window
    $script:ApEl = $ui.Elements
    $el = $script:ApEl
    $config = Get-ApConfig

    # ---------- header state ----------
    if (Test-ApElevated) {
        Set-ApPill $el.PillAdmin $el.PillAdminText 'Administrator' 'ok'
    }
    else {
        Set-ApPill $el.PillAdmin $el.PillAdminText 'Not elevated' 'error'
    }

    $el.SidebarAppVersion.Text = "GUI $script:ApAppVersion"
    $el.AboutVersion.Text = "Version $script:ApAppVersion"

    # ---------- branding links ----------
    $el.SidebarSiteLink.Add_Click({ Start-ApUrl $script:ApAuthorSite })
    $el.AboutSiteLink.Add_Click({ Start-ApUrl $script:ApAuthorSite })
    $el.AboutGitHubLink.Add_Click({ Start-ApUrl $script:ApAuthorGitHub })
    $el.AboutEngineLink.Add_Click({ Start-ApUrl 'https://github.com/andrew-s-taylor/WindowsAutopilotInfo' })
    $el.AboutOriginalLink.Add_Click({ Start-ApUrl 'https://github.com/ugurkocde/AutoPilot_Import_GUI' })

    # ---------- device ----------
    # Only placeholders here. The real read happens on the window's Loaded event so the
    # window is on screen first; see Start-ApDeviceLoad for why it cannot be synchronous.
    Set-ApDevicePlaceholders

    # ---------- engine ----------
    try {
        $script:ApEnginePath = Resolve-ApEngineScript -Name Engine
        $manifest = Get-ApVendorManifest
        $engineEntry = $null
        if ($manifest) {
            $engineEntry = $manifest.scripts | Where-Object { $_.role -eq 'engine' } | Select-Object -First 1
        }
        $versionText = if ($engineEntry) { "Engine v$($engineEntry.version)" } else { 'Engine: resolved' }
        $el.SidebarEngineVersion.Text = $versionText
        $el.AdvEnginePath.Text = $script:ApEnginePath
        Write-ApLog "Engine resolved to $script:ApEnginePath"
    }
    catch {
        $el.SidebarEngineVersion.Text = 'Engine: not found'
        $el.AdvEnginePath.Text = $_.Exception.Message
        Set-ApStatus -Text 'The community engine script could not be found. See the Advanced page.' -IsError
        Write-ApLog "Engine not resolved: $($_.Exception.Message)" -Level ERROR
    }

    # ---------- restore saved preferences ----------
    $history = @($config.groupTagHistory | Where-Object { $_ })
    $el.GroupTagCombo.ItemsSource = $history
    $el.GroupTagCombo.Text = if ($GroupTag) { $GroupTag } else { "$($config.lastGroupTag)" }
    $el.AssignedUserBox.Text = if ($AssignedUser) { $AssignedUser } else { "$($config.lastAssignedUser)" }
    $el.ComputerNameBox.Text = "$($config.lastComputerName)"
    $el.AddToGroupBox.Text = "$($config.lastAddToGroup)"
    $el.WaitAssignCheck.IsChecked = [bool]$config.waitForAssignment
    $el.RebootCheck.IsChecked = [bool]$config.rebootWhenAssigned
    $el.RebootV2Check.IsChecked = [bool]$config.rebootAfterV2Import
    $el.AdvShowConsoleCheck.IsChecked = [bool]$config.showConsoleWindow
    $el.AdvDiagnosticsOnlineCheck.IsChecked = [bool]$config.diagnosticsOnline

    switch ("$($config.existingDevicePolicy)") {
        'delete'    { $el.PolicyDelete.IsChecked = $true }
        'skipcheck' { $el.PolicySkip.IsChecked = $true }
        default     { $el.PolicyUpdate.IsChecked = $true }
    }

    $effectiveMode = if ($Mode) { $Mode } else { "$($config.lastMode)" }
    if ($effectiveMode -eq 'v2') { $el.ModeV2.IsChecked = $true } else { $el.ModeV1.IsChecked = $true }

    # Remembered for the device load, which may switch to v2 when there is no hash: an
    # explicit -Mode from the caller must not be overridden.
    $script:ApModePinned = [bool]$Mode

    Sync-ApModeUi
    Update-ApLogsPage

    # ---------- navigation ----------
    $el.NavRegister.Add_Checked({ Show-ApPage 'PageRegister' })
    $el.NavDevice.Add_Checked({ Show-ApPage 'PageDevice' })
    $el.NavNetwork.Add_Checked({ Show-ApPage 'PageNetwork' })
    $el.NavAdvanced.Add_Checked({ Show-ApPage 'PageAdvanced' })
    $el.NavLogs.Add_Checked({ Show-ApPage 'PageLogs'; Update-ApLogsPage })

    # ---------- mode ----------
    $el.ModeV1.Add_Checked({ Sync-ApModeUi })
    $el.ModeV2.Add_Checked({ Sync-ApModeUi })

    # Reboot is only honoured after assignment, so reflect that in the UI immediately
    # rather than silently coercing it at launch time.
    $el.RebootCheck.Add_Checked({
        if (-not $script:ApEl.WaitAssignCheck.IsChecked) {
            $script:ApEl.WaitAssignCheck.IsChecked = $true
            Set-ApStatus -Text 'Waiting for profile assignment was enabled: a restart only happens after assignment.'
        }
    })

    # ---------- register ----------
    $el.RegisterButton.Add_Click({
        $el = $script:ApEl
        $request = Get-ApUiRequest -Operation Register

        try {
            $built = Build-ApEngineArguments $request
        }
        catch {
            Show-ApDialog -Title 'Cannot register' -Owner $script:ApWin -Message $_.Exception.Message | Out-Null
            return
        }

        if ($built.Warnings.Count -gt 0) { Show-ApNotice -Title 'Note' -Messages $built.Warnings }

        # v1 with no hash will fail in the engine; say so here instead.
        if ($request.Mode -eq 'v1' -and $script:ApDevice -and -not $script:ApDevice.HardwareHashReady) {
            $proceed = Show-ApDialog -Title 'No hardware hash' -Owner $script:ApWin -ShowCancel `
                -ConfirmText 'Try anyway' `
                -Message 'This device did not return an Autopilot hardware hash, so a v1 registration is very likely to fail. Device Preparation (v2) does not need one.'
            if (-not $proceed) { return }
        }

        # A broken sign-in module aborts the engine before any browser prompt appears, which
        # reads as "nothing happened". Say so up front instead.
        if ($script:ApGraphCheck -and -not $script:ApGraphCheck.Available -and $script:ApGraphCheck.InstalledVersions.Count -gt 0) {
            $proceed = Show-ApDialog -Title 'Sign-in will probably fail' -Owner $script:ApWin -ShowCancel `
                -ConfirmText 'Try anyway' `
                -Message 'The Microsoft Graph sign-in module is installed but cannot be loaded, so the engine will stop before a sign-in prompt appears. Use "Repair sign-in module" on the Advanced page.' `
                -Detail (Get-ApGraphModuleAdvice -Check $script:ApGraphCheck)
            if (-not $proceed) { return }
        }

        # Device Preparation restart: confirm, because a restart from OOBE is disruptive and
        # is premature unless the device is already in the policy's Entra group.
        $wantsV2Reboot = ($request.Mode -eq 'v2') -and [bool]$el.RebootV2Check.IsChecked
        if ($wantsV2Reboot) {
            $proceed = Show-ApDialog -Title 'Restart after import' -Owner $script:ApWin -ShowCancel `
                -ConfirmText 'Import and restart' `
                -Message ('This device will restart as soon as the identifier is imported. ' +
                          'Device Preparation targets devices through the Entra group on the policy, so ' +
                          'restart only if this device is already a member of that group. It will not ' +
                          'restart if the import fails or you cancel.')
            if (-not $proceed) { return }
        }
        $script:ApPendingV2Reboot = $wantsV2Reboot

        if (Test-ApDestructiveRequest $request) {
            $detail = Get-ApPreviewCommand -Parameters $built.Parameters -ScriptPath (Split-Path -Leaf $script:ApEnginePath)
            $lines = New-Object System.Collections.Generic.List[string]
            if ($request.ExistingDevicePolicy -eq 'delete') {
                $lines.Add('If this serial is already registered it will be deleted from Autopilot, Intune and Entra ID before being re-added.')
            }
            if ($request.Wipe) { $lines.Add('An Intune wipe will be sent to this device. All data on it will be erased.') }
            if ($request.Sysprep) { $lines.Add('Sysprep will run and the device will restart into OOBE.') }
            if ($request.ChangePK) { $lines.Add('The Windows product key will be changed and the device will restart.') }

            $proceed = Show-ApDialog -Title 'Confirm a destructive action' -Owner $script:ApWin `
                -ShowCancel -Danger -ConfirmText 'Yes, continue' -ShowCopy `
                -Message ($lines -join [Environment]::NewLine) -Detail $detail
            if (-not $proceed) { return }
        }

        # Remember what was used, so the next device is one click.
        Add-ApGroupTagToHistory $request.GroupTag
        Set-ApConfigValue 'lastMode' $request.Mode
        Set-ApConfigValue 'lastGroupTag' $request.GroupTag
        Set-ApConfigValue 'lastAssignedUser' $request.AssignedUser
        Set-ApConfigValue 'lastComputerName' $request.AssignedComputerName
        Set-ApConfigValue 'lastAddToGroup' $request.AddToGroup
        Set-ApConfigValue 'waitForAssignment' ([bool]$request.WaitForAssignment)
        Set-ApConfigValue 'rebootWhenAssigned' ([bool]$request.Reboot)
        Set-ApConfigValue 'rebootAfterV2Import' ([bool]$el.RebootV2Check.IsChecked)
        Set-ApConfigValue 'existingDevicePolicy' $request.ExistingDevicePolicy
        Save-ApConfig | Out-Null
        $el.GroupTagCombo.ItemsSource = @((Get-ApConfig).groupTagHistory)

        foreach ($note in $built.Notes) { Write-ApLog "Note: $note" }

        Start-ApGuiRun -Parameters $built.Parameters -OutputBox $el.RegisterOutput `
                       -Label 'register' -StartMessage 'Signing in to Microsoft Graph. A browser window will open.'
    })

    $el.PreviewButton.Add_Click({
        try {
            $built = Build-ApEngineArguments (Get-ApUiRequest -Operation Register)
        }
        catch {
            Show-ApDialog -Title 'Cannot build the command' -Owner $script:ApWin -Message $_.Exception.Message | Out-Null
            return
        }

        $scriptName = if ($script:ApEnginePath) { $script:ApEnginePath } else { 'get-windowsautopilotinfocommunity.ps1' }
        $detail = Get-ApPreviewCommand -Parameters $built.Parameters -ScriptPath $scriptName

        $message = 'This is exactly what will run. Nothing has been executed.'
        if ($built.Notes.Count -gt 0) {
            $message += [Environment]::NewLine + [Environment]::NewLine + (($built.Notes | ForEach-Object { "- $_" }) -join [Environment]::NewLine)
        }
        if ($built.Warnings.Count -gt 0) {
            $message += [Environment]::NewLine + [Environment]::NewLine + (($built.Warnings | ForEach-Object { "! $_" }) -join [Environment]::NewLine)
        }

        Show-ApDialog -Title 'Preview command' -Owner $script:ApWin -Message $message -Detail $detail -ShowCopy | Out-Null
    })

    $el.CancelButton.Add_Click({ Stop-ApGuiRun })

    $el.RestartNowButton.Add_Click({
        # Refuse mid-run: restarting during an import abandons it half-finished and the log
        # is lost with the session.
        if ($script:ApRun -and -not $script:ApRun.IsFinished) {
            Show-ApDialog -Title 'A run is in progress' -Owner $script:ApWin `
                -Message 'Wait for the current operation to finish, or cancel it, before restarting.' | Out-Null
            return
        }

        $proceed = Show-ApDialog -Title 'Restart this device now' -Owner $script:ApWin -ShowCancel `
            -ConfirmText 'Restart now' -CancelText 'Not yet' -Danger `
            -Message ('This device will restart immediately. Anything unsaved will be lost.' + [Environment]::NewLine + [Environment]::NewLine +
                      'Before restarting, make sure this device is a member of the Entra group targeted by your ' +
                      'Device Preparation policy, or it will return to OOBE before the policy can apply.')
        if (-not $proceed) { return }

        Set-ApStatus -Text 'Restarting...'
        Invoke-ApRestartComputer | Out-Null
    })

    # ---------- device page ----------
    $el.RefreshDeviceButton.Add_Click({
        Set-ApDevicePlaceholders
        Start-ApDeviceLoad
    })

    $el.ExportHashButton.Add_Click({ Start-ApCsvExport -Mode 'v1' })
    $el.ExportIdentifierButton.Add_Click({ Start-ApCsvExport -Mode 'v2' })

    $el.CopyIdentifierButton.Add_Click({
        if (-not $script:ApDevice) { return }
        try {
            Set-Clipboard -Value $script:ApDevice.DeviceIdentifier
            Set-ApStatus -Text 'Device identifier copied to the clipboard.'
        }
        catch {
            Set-ApStatus -Text "Could not copy: $($_.Exception.Message)" -IsError
        }
    })

    $el.CopyHashButton.Add_Click({
        $hash = Get-ApHardwareHash
        if (-not $hash) {
            Show-ApDialog -Title 'No hardware hash' -Owner $script:ApWin `
                -Message 'This device did not return a hardware hash. Administrator rights are required, and virtual machines often cannot provide one.' | Out-Null
            return
        }
        try {
            Set-Clipboard -Value $hash
            Set-ApStatus -Text "Hardware hash copied to the clipboard ($($hash.Length) characters)."
        }
        catch {
            Set-ApStatus -Text "Could not copy: $($_.Exception.Message)" -IsError
        }
    })

    # ---------- network page ----------
    $el.NetworkRunButton.Add_Click({ Start-ApNetworkCheck })

    $el.NetworkCopyButton.Add_Click({
        if (-not $script:ApNetworkResults -or $script:ApNetworkResults.Count -eq 0) {
            Set-ApStatus -Text 'Run the check first.'
            return
        }
        $text = $script:ApNetworkResults |
            Sort-Object Category, Name |
            ForEach-Object { '{0,-12} {1,-32} {2,-45} {3} ms' -f $_.Status, $_.Name, $_.Host, $_.LatencyMs }
        try {
            Set-Clipboard -Value ($text -join [Environment]::NewLine)
            Set-ApStatus -Text 'Connectivity results copied to the clipboard.'
        }
        catch {
            Set-ApStatus -Text "Could not copy: $($_.Exception.Message)" -IsError
        }
    })

    # ---------- advanced page ----------
    $el.AdvShowConsoleCheck.Add_Click({
        Set-ApConfigValue 'showConsoleWindow' ([bool]$script:ApEl.AdvShowConsoleCheck.IsChecked)
        Save-ApConfig | Out-Null
    })

    $el.AdvDiagnosticsOnlineCheck.Add_Click({
        Set-ApConfigValue 'diagnosticsOnline' ([bool]$script:ApEl.AdvDiagnosticsOnlineCheck.IsChecked)
        Save-ApConfig | Out-Null
    })

    $el.AdvVerifyEngineButton.Add_Click({
        $el = $script:ApEl
        if (-not $script:ApEnginePath) {
            $el.AdvEngineIntegrity.Text = 'No engine script resolved.'
            return
        }

        $check = Test-ApVendorScript -Path $script:ApEnginePath
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add("Signature: $($check.SignatureStatus)")
        if ($check.Signer) { $lines.Add("Signer: $($check.Signer)") }
        if ($null -ne $check.HashMatches) {
            $lines.Add("Checksum: $(if ($check.HashMatches) { 'matches vendor manifest' } else { 'DOES NOT MATCH' })")
        }
        else {
            $lines.Add('Checksum: no manifest entry to compare against')
        }
        foreach ($m in $check.Messages) { $lines.Add($m) }

        $el.AdvEngineIntegrity.Text = ($lines -join [Environment]::NewLine)
        Write-ApLog ("Engine integrity: " + ($lines -join ' | '))

        Show-ApDialog -Title 'Engine integrity' -Owner $script:ApWin `
                      -Message $(if ($check.IsTrusted) { 'The engine script matches what was shipped.' } else { 'The engine script does not match the vendor manifest. Do not use it until you know why.' }) `
                      -Detail (($lines -join [Environment]::NewLine) + "`r`n`r`n$($check.Path)") | Out-Null
    })

    $el.AdvOpenWorkDirButton.Add_Click({
        try { Start-Process explorer.exe (Get-ApWorkingDirectory) } catch { }
    })

    $el.AdvRepairGraphButton.Add_Click({
        $detail = ''
        if ($script:ApGraphCheck) { $detail = Get-ApGraphModuleAdvice -Check $script:ApGraphCheck }

        $proceed = Show-ApDialog -Title 'Repair the sign-in module' -Owner $script:ApWin -ShowCancel `
            -ConfirmText 'Repair now' `
            -Message ('This removes every installed copy of Microsoft.Graph.Authentication and reinstalls version ' +
                      "$script:ApGraphPinnedVersion for all users, which is the version the community engine expects. " +
                      'It needs access to the PowerShell Gallery. Continue?') `
            -Detail $detail
        if (-not $proceed) { return }

        Start-ApGraphRepair
    })

    $el.AdvDiagnosticsButton.Add_Click({
        try { $diag = Resolve-ApEngineScript -Name Diagnostics }
        catch {
            Show-ApDialog -Title 'Diagnostics unavailable' -Owner $script:ApWin -Message $_.Exception.Message | Out-Null
            return
        }

        # -Online only adds lookups: app, policy and script GUIDs from the local logs are
        # resolved to display names through Graph. Everything the script reports is still
        # read from this machine, and nothing is written either way.
        $online = [bool]$script:ApEl.AdvDiagnosticsOnlineCheck.IsChecked

        # Same trap as a registration run: a broken sign-in module aborts the engine before
        # any browser prompt appears, which reads as "nothing happened".
        if ($online -and $script:ApGraphCheck -and -not $script:ApGraphCheck.Available -and
            $script:ApGraphCheck.InstalledVersions.Count -gt 0) {
            $proceed = Show-ApDialog -Title 'Sign-in will probably fail' -Owner $script:ApWin -ShowCancel `
                -ConfirmText 'Try anyway' `
                -Message ('The Microsoft Graph sign-in module is installed but cannot be loaded, so the online lookup will stop before a sign-in prompt appears. ' +
                          'Clear the -Online option to run the local diagnostics anyway, or use "Repair sign-in module".') `
                -Detail (Get-ApGraphModuleAdvice -Check $script:ApGraphCheck)
            if (-not $proceed) { return }
        }

        # An 'Online' key is also what makes the launcher install the sign-in prerequisites
        # non-interactively (see Get-ApDependencyPrepBlock); a local run must not touch the gallery.
        $params = [ordered]@{}
        $message = 'Collecting Autopilot diagnostics...'
        if ($online) {
            $params['Online'] = $true
            $message = 'Collecting Autopilot diagnostics and signing in to resolve app and policy names...'
        }

        $script:ApEl.NavLogs.IsChecked = $true
        Start-ApGuiRun -Parameters $params -OutputBox $script:ApEl.LogsOutput `
                       -EnginePath $diag -Label 'diagnostics' `
                       -StartMessage $message
    })

    $el.AdvWindowsUpdateButton.Add_Click({
        $proceed = Show-ApDialog -Title 'Install Windows updates' -Owner $script:ApWin -ShowCancel `
            -ConfirmText 'Install updates' `
            -Message 'This searches for and installs all applicable Windows updates using the in-box Windows Update agent. The device may restart without further warning. Continue?'
        if (-not $proceed) { return }
        Start-ApWindowsUpdate
    })

    $el.AdvWipeCheck.Add_Click({ Sync-ApAdvancedWarnings })
    $el.AdvSysprepCheck.Add_Click({ Sync-ApAdvancedWarnings })
    Sync-ApAdvancedWarnings

    # ---------- logs page ----------
    # -Force: clicking Refresh is an explicit request for the session log, even if a run
    # report is currently on screen. The run's own log file on disk is unaffected.
    $el.LogsRefreshButton.Add_Click({ Update-ApLogsPage -Force })

    $el.LogsCopyButton.Add_Click({
        try {
            Set-Clipboard -Value $script:ApEl.LogsOutput.Text
            Set-ApStatus -Text 'Log copied to the clipboard.'
        }
        catch {
            Set-ApStatus -Text "Could not copy: $($_.Exception.Message)" -IsError
        }
    })

    $el.LogsOpenFolderButton.Add_Click({
        try { Start-Process explorer.exe (Split-Path -Parent (Get-ApLogPath)) } catch { }
    })

    # ---------- timers ----------
    # Seed the clock now; otherwise it reads "--:--:--" for the first second.
    $el.HeaderClock.Text = (Get-Date).ToString('HH:mm:ss')

    $clock = New-Object System.Windows.Threading.DispatcherTimer
    $clock.Interval = [TimeSpan]::FromSeconds(1)
    # Guarded for the same reason as the background tick: an escaping exception would take
    # ShowDialog down with it.
    $clock.Add_Tick({
        try {
            $script:ApEl.HeaderClock.Text = (Get-Date).ToString('HH:mm:ss')

            # Re-probe the network every 15 seconds. In OOBE, Wi-Fi often comes up after the
            # tool is already open, so a one-shot check at startup goes stale immediately.
            $script:ApClockTicks = ($script:ApClockTicks + 1)
            if (($script:ApClockTicks % 15) -eq 1) { Update-ApNetworkPill }
        }
        catch {
            Write-ApLog "Clock tick failed: $($_.Exception.Message)" -Level ERROR
        }
    })
    $script:ApClockTicks = 0
    $script:ApClockTimer = $clock
    $clock.Start()

    Update-ApNetworkPill

    # ---------- shutdown ----------
    # A WPF event handler receives (sender, eventArgs) as positional arguments. $_ is NOT the
    # event args here: using it threw "Argument types do not match" out of Window.OnClosing,
    # so the cancel-the-close path never worked and the exception escaped ShowDialog.
    $script:ApWin.Add_Closing({
        param($closeSender, $closeArgs)

        try {
            if ($script:ApRun -and -not $script:ApRun.IsFinished) {
                $cancelRun = Show-ApDialog -Title 'A run is still in progress' -Owner $script:ApWin -ShowCancel `
                    -ConfirmText 'Cancel the run and close' -CancelText 'Keep it running' -Danger `
                    -Message 'Closing this window will terminate the operation in progress. Registration may be left half-finished.'

                if (-not $cancelRun) {
                    $closeArgs.Cancel = $true
                    return
                }
                Stop-ApEngineRun -Run $script:ApRun
            }

            # Each step guarded separately: a failure in one must not skip the others,
            # especially not the config save.
            try { if ($script:ApRunTimer) { $script:ApRunTimer.Stop() } }
            catch { Write-ApLog "Stopping the run timer failed: $($_.Exception.Message)" -Level WARN }

            try { if ($script:ApClockTimer) { $script:ApClockTimer.Stop() } }
            catch { Write-ApLog "Stopping the clock failed: $($_.Exception.Message)" -Level WARN }

            try { Stop-ApBackgroundWork }
            catch { Write-ApLog "Stopping background work failed: $($_.Exception.Message)" -Level WARN }

            try { Save-ApConfig | Out-Null }
            catch { Write-ApLog "Saving configuration failed: $($_.Exception.Message)" -Level WARN }

            Write-ApLog 'Session closed.'
        }
        catch {
            Write-ApLog "Shutdown failed: $($_.Exception.Message)" -Level ERROR
        }
    })

    # Kick the slow work off only once the window is actually on screen.
    $script:ApWin.Add_Loaded({
        Start-ApDeviceLoad
        Start-ApGraphPreflight
    })

    Set-ApStatus -Text 'Ready.' -Stage '' -Percent 0
    Write-ApLog 'GUI ready.'

    return $script:ApWin
}

function Sync-ApAdvancedWarnings {
    <#
    .SYNOPSIS
    Shows the wipe warning only when a destructive option is actually selected.
    #>
    [CmdletBinding()]
    param()

    $el = $script:ApEl
    $el.AdvWipeWarning.Visibility = if ($el.AdvWipeCheck.IsChecked) { 'Visible' } else { 'Collapsed' }
}

function Update-ApNetworkPill {
    <#
    .SYNOPSIS
    Refreshes the header network chip without blocking the UI thread.
    #>
    [CmdletBinding()]
    param()

    Start-ApBackgroundWork -Window $script:ApWin `
        -FunctionNames @('Test-ApInternetConnection') `
        -Work { Test-ApInternetConnection } `
        -OnComplete {
            param($result, $errorMessage)
            $online = [bool](@($result)[-1])
            if ($online) {
                Set-ApPill $script:ApEl.PillNetwork $script:ApEl.PillNetworkText 'Online' 'ok'
            }
            else {
                Set-ApPill $script:ApEl.PillNetwork $script:ApEl.PillNetworkText 'Offline' 'error'
            }
        } | Out-Null
}

function Start-ApGraphPreflight {
    <#
    .SYNOPSIS
    Checks in the background whether the sign-in module can be imported, and reports it.
    #>
    [CmdletBinding()]
    param()

    $script:ApEl.AdvGraphStatus.Text = 'Checking...'

    Start-ApBackgroundWork -Window $script:ApWin `
        -FunctionNames @('Test-ApGraphModule') `
        -Work { Test-ApGraphModule } `
        -OnComplete {
            param($result, $errorMessage)

            $el = $script:ApEl

            if ($errorMessage -or -not $result) {
                $el.AdvGraphStatus.Text = "Could not check: $errorMessage"
                return
            }

            $check = @($result)[-1]
            $script:ApGraphCheck = $check

            if ($check.Available) {
                $el.AdvGraphStatus.Text = "Microsoft.Graph.Authentication $($check.Version) loads correctly."
                $el.AdvGraphStatus.Foreground = New-ApBrush '#3BC77A'
                return
            }

            $el.AdvGraphStatus.Foreground = New-ApBrush '#E9B44C'
            if ($check.InstalledVersions.Count -eq 0) {
                $el.AdvGraphStatus.Text = 'Not installed for Windows PowerShell. The engine will install it on first sign-in, which needs the PowerShell Gallery.'
            }
            else {
                $el.AdvGraphStatus.Text = "Version $($check.Version) is installed but will not load. Sign-in will fail until this is repaired."
            }

            Write-ApLog "Graph module preflight: available=$($check.Available) versions=$($check.InstalledVersions -join ',') error=$($check.Error)" -Level WARN
        } | Out-Null
}

function Start-ApGraphRepair {
    <#
    .SYNOPSIS
    Runs the Graph module repair in a visible console.

    .DESCRIPTION
    Separate console rather than a streamed run: it uninstalls and reinstalls a module, which
    can prompt, and the tech should see it happen.
    #>
    [CmdletBinding()]
    param()

    $path = Join-Path (Get-ApWorkingDirectory) 'repair-graph-module.ps1'
    Set-Content -LiteralPath $path -Value (Get-ApGraphRepairScript) -Encoding UTF8

    Write-ApLog 'Launching Graph module repair in a separate console.'
    Start-Process -FilePath (Get-ApPowerShellPath) `
                  -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $path) | Out-Null

    Set-ApStatus -Text 'Repairing the sign-in module in a separate window. Restart this tool when it finishes.'
}

function Start-ApWindowsUpdate {
    <#
    .SYNOPSIS
    Searches for and installs available Windows updates in a visible console.

    .DESCRIPTION
    Uses the in-box Windows Update Agent COM API (Microsoft.Update.Session) rather than the
    PSWindowsUpdate module the original GUI installed
    (Get-WindowsAutopilotImportGUI.ps1:171). Three reasons:

      * no dependency to fetch, so it works on an OOBE network that cannot reach the
        PowerShell Gallery, which is exactly where this button gets used
      * no NuGet provider bootstrap, no repository trust change, no module install
      * generating a script that chains Install-PackageProvider, Set-PSRepository -Trusted
        and Install-Module -Force is a well-known malware-dropper shape, and Defender's
        AMSI scanner blocks it on sight

    Runs in its own visible console because a Windows Update install can restart the
    machine; streaming that into the GUI would look like a hang.
    #>
    [CmdletBinding()]
    param()

    $updateScript = @'
$Host.UI.RawUI.WindowTitle = 'Autopilot Import GUI - Windows Update'
$ErrorActionPreference = 'Stop'

function Write-Step { param([string]$Text) Write-Host "==> $Text" -ForegroundColor Cyan }

try {
    Write-Step 'Searching for applicable updates...'
    $session  = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()
    $result   = $searcher.Search("IsInstalled=0 AND Type='Software' AND IsHidden=0")

    if ($result.Updates.Count -eq 0) {
        Write-Host 'No applicable updates were found.' -ForegroundColor Green
    }
    else {
        Write-Host ''
        Write-Host "Found $($result.Updates.Count) update(s):" -ForegroundColor Yellow
        $toInstall = New-Object -ComObject Microsoft.Update.UpdateColl
        foreach ($update in $result.Updates) {
            Write-Host "  - $($update.Title)"
            # Updates with unaccepted licence terms cannot be installed unattended.
            if ($update.EulaAccepted -eq $false) { $update.AcceptEula() }
            [void]$toInstall.Add($update)
        }

        Write-Host ''
        Write-Step 'Downloading...'
        $downloader = $session.CreateUpdateDownloader()
        $downloader.Updates = $toInstall
        $downloadResult = $downloader.Download()
        Write-Host "Download result code: $($downloadResult.ResultCode)"

        Write-Step 'Installing. This can take a while.'
        $installer = $session.CreateUpdateInstaller()
        $installer.Updates = $toInstall
        $installResult = $installer.Install()

        Write-Host ''
        Write-Host "Install result code: $($installResult.ResultCode)" -ForegroundColor Green
        if ($installResult.RebootRequired) {
            Write-Host 'A restart is required to finish installing these updates.' -ForegroundColor Yellow
        }
    }
}
catch {
    Write-Host ''
    Write-Host "Windows Update failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ''
Read-Host 'Press Enter to close'
'@

    $path = Join-Path (Get-ApWorkingDirectory) 'windowsupdate.ps1'
    Set-Content -LiteralPath $path -Value $updateScript -Encoding UTF8

    Write-ApLog 'Launching Windows Update in a separate console.'
    Start-Process -FilePath (Get-ApPowerShellPath) `
                  -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $path) | Out-Null

    Set-ApStatus -Text 'Windows Update is running in a separate window.'
}
#endregion src\Public\Show-AutopilotImportGui.ps1


Initialize-ApLog | Out-Null
Import-ApConfig | Out-Null
Write-ApLog "Starting single-file build from $script:ApAppRoot"

# Must happen before anything auto-loads a module, and before any child process is started:
# a PowerShell 7 parent leaks its module path into this Windows PowerShell process, which
# shadows in-box modules with Core-only copies and breaks the engine before sign-in.
Repair-ApModulePath | Out-Null

# Elevate once for the whole session so child engine runs inherit the token.
if (-not $NoElevate -and -not (Test-ApElevated)) {
    Write-ApLog 'Not elevated; requesting elevation.'
    if (Invoke-ApSelfElevate -ScriptPath $MyInvocation.MyCommand.Path -BoundParameters $PSBoundParameters) {
        exit 0
    }
    Write-ApLog 'Continuing without elevation. The hardware hash will not be readable.' -Level WARN
}

try {
    # -Mode has a ValidateSet that rejects the empty-string default, so only forward
    # parameters that were actually supplied.
    $guiArgs = @{}
    if ($GroupTag) { $guiArgs.GroupTag = $GroupTag }
    if ($AssignedUser) { $guiArgs.AssignedUser = $AssignedUser }
    if ($Mode) { $guiArgs.Mode = $Mode }

    Show-AutopilotImportGui @guiArgs
}
catch {
    $message = $_.Exception.Message
    Write-ApLog "Fatal error: $message" -Level ERROR
    if ($_.ScriptStackTrace) { Write-ApLog $_.ScriptStackTrace -Level DEBUG }

    try {
        [System.Windows.MessageBox]::Show(
            "$message`r`n`r`nLog: $(Get-ApLogPath)",
            'Autopilot Import GUI', 'OK', 'Error') | Out-Null
    }
    catch {
        Write-Host "Autopilot Import GUI failed: $message" -ForegroundColor Red
    }
    exit 1
}