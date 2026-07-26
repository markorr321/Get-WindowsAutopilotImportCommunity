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

function Show-ApOpenFileDialog {
    <#
    .SYNOPSIS
    Open-file picker. Returns the chosen path, or $null if cancelled.
    #>
    [CmdletBinding()]
    param(
        [string]$Title = 'Select a file',
        [string]$Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    )

    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.Title = $Title
    $dialog.Filter = $Filter
    $dialog.Multiselect = $false

    if ($dialog.ShowDialog()) { return $dialog.FileName }
    return $null
}
