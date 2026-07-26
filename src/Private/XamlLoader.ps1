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
