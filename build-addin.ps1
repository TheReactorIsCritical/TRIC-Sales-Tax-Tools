param(
    [string]$ProjectRoot = (Split-Path -Parent $MyInvocation.MyCommand.Path),
    [string]$AddInName = "TRIC Sales Tax Tools",
    [switch]$Force
)

$srcDir = Join-Path $ProjectRoot "src"
if (-not (Test-Path $srcDir)) {
    throw "src folder not found: $srcDir"
}

$basFiles = Get-ChildItem -Path $srcDir -Filter *.bas -File
if ($basFiles.Count -eq 0) {
    throw "No .bas files found in $srcDir"
}

$addinsDir = Join-Path $env:APPDATA "Microsoft\AddIns"
if (-not (Test-Path $addinsDir)) {
    New-Item -ItemType Directory -Path $addinsDir | Out-Null
}

$outXlam = Join-Path $addinsDir ($AddInName + ".xlam")

# Build to a temp location first, then copy into AddIns folder.
# This avoids partially overwriting the live add-in if Excel has it loaded.
$outXlamTemp = Join-Path $env:TEMP ($AddInName + "_build.xlam")
if ((Test-Path $outXlam) -and (-not $Force)) {
    throw "Add-in already exists: $outXlam (re-run with -Force to overwrite)"
}

$tempXlsm = Join-Path $env:TEMP ($AddInName + "_build.xlsm")
if (Test-Path $tempXlsm) { Remove-Item $tempXlsm -Force }
if (Test-Path $outXlamTemp) { Remove-Item $outXlamTemp -Force }

# Excel constants
$xlOpenXMLWorkbookMacroEnabled = 52  # .xlsm
$xlOpenXMLAddIn  = 55                # .xlam

# Ribbon XML
# Keep the inner <ribbon> markup identical for both schemas.
$ribbonBody = @'
  <ribbon>
    <tabs>
      <tab id="tabTRICSalesTaxes" label="TRIC Sales Taxes">
        <group id="grp0" label="Sales Taxes">
          <button
            id="btnPrepareWorkbook"
            size="large"
            label="Prepare Workbook"
            imageMso="TableInsert"
            supertip="Creates or refreshes the Tax Summary worksheet and pulls in the data needed for reporting.&#10;&#10;Does not modify your source data. It only generates the summary output."
            onAction="Button_PrepareWorkbook"/>
          <button 
            id="btnOpenSquarespaceAccounting"
            size="large"
            label="Squarespace Accounting"
            imageMso="WebPagePreview"
            supertip="Opens the Squarespace Accounting page in your default web browser to get the required source data."
            onAction="Button_OpenSquarespaceAccounting"/>
          <button 
            id="btnOpenGithubRepository"
            size="large"
            label="Documentation"
            imageMso="Help"
            supertip="Opens the GitHub repository page in your default web browser."
            onAction="Button_OpenGithubRepository"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
'@

function New-CustomUiXml {
    param(
        [Parameter(Mandatory=$true)][string]$NamespaceUri,
        [Parameter(Mandatory=$true)][string]$OnLoad,
        [Parameter(Mandatory=$true)][string]$BodyXml
    )

    # Use CRLF to keep output stable across Windows tooling.
    $nl = "`r`n"
    return ('<customUI xmlns="{0}" onLoad="{1}">' -f $NamespaceUri, $OnLoad) +
        $nl + $BodyXml + $nl +
        '</customUI>'
}

$ribbonOnLoad = "OnRibbonLoad"

$ribbonXml2006 = New-CustomUiXml `
    -NamespaceUri "http://schemas.microsoft.com/office/2006/01/customui" `
    -OnLoad $ribbonOnLoad `
    -BodyXml $ribbonBody

$ribbonXml14 = New-CustomUiXml `
    -NamespaceUri "http://schemas.microsoft.com/office/2009/07/customui" `
    -OnLoad $ribbonOnLoad `
    -BodyXml $ribbonBody


$excel = $null
$wb = $null

function Test-FileLocked {
    param([Parameter(Mandatory=$true)][string]$Path)

    if (-not (Test-Path $Path)) { return $false }

    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $fs.Dispose()
        return $false
    } catch {
        return $true
    }
}

function Add-RibbonXToOpenXmlFile {
    param(
        [Parameter(Mandatory=$true)][string]$OpenXmlPath,
        [Parameter(Mandatory=$true)][string]$RibbonXml2006,
        [Parameter(Mandatory=$true)][string]$RibbonXml14
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $uiPath2006    = "customUI/customUI.xml"
    $uiPath14      = "customUI/customUi14.xml"
    $relsPath      = "_rels/.rels"
    $ctPath        = "[Content_Types].xml"

    # Match known-good package structure
    $personPath    = "xl/persons/person.xml"
    $wbRelsPath    = "xl/_rels/workbook.xml.rels"

    $uiRelType2006 = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
    $uiRelType2007 = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
    $pkgRelNs      = "http://schemas.openxmlformats.org/package/2006/relationships"

    $ctNs          = "http://schemas.openxmlformats.org/package/2006/content-types"
    $customUiCt    = "application/vnd.ms-office.customui+xml"
    $personCt      = "application/vnd.ms-excel.person+xml"

    $personRelType = "http://schemas.microsoft.com/office/2017/10/relationships/person"

    $personXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
'@

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)

    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::Open($OpenXmlPath, 'Update')

        function Read-ZipText([string]$path) {
            $e = $zip.GetEntry($path)
            if (-not $e) { return $null }
            $s = $e.Open()
            try {
                $r = New-Object System.IO.StreamReader($s, $utf8NoBom, $true)
                try { return $r.ReadToEnd() }
                finally { $r.Dispose() }
            } finally { $s.Dispose() }
        }

        function Write-ZipText([string]$path, [string]$text) {
            $e = $zip.GetEntry($path)
            if ($e) { $e.Delete() }

            $entry  = $zip.CreateEntry($path)
            $stream = $entry.Open()
            try {
                $writer = New-Object System.IO.StreamWriter($stream, $utf8NoBom)
                try {
                    $writer.Write($text)
                    $writer.Flush()
                } finally { $writer.Dispose() }
            } finally { $stream.Dispose() }
        }

        function Write-XmlNoDecl([xml]$xmlDoc) {
            $sb = New-Object System.Text.StringBuilder
            $sw = New-Object System.IO.StringWriter($sb)

            $settings = New-Object System.Xml.XmlWriterSettings
            $settings.OmitXmlDeclaration = $true
            $settings.Indent = $false
            $settings.NewLineHandling = "None"

            $xw = [System.Xml.XmlWriter]::Create($sw, $settings)
            try {
                $xmlDoc.WriteTo($xw)
                $xw.Flush()
            } finally {
                $xw.Dispose()
                $sw.Dispose()
            }

            return $sb.ToString()
        }

        # --- Write both ribbon files ---
        Write-ZipText $uiPath2006 $RibbonXml2006
        Write-ZipText $uiPath14   $RibbonXml14
        Write-ZipText $personPath $personXml

        # --- _rels/.rels : load or create, then ensure BOTH relationships ---
        $relsXml = Read-ZipText $relsPath
        if ([string]::IsNullOrWhiteSpace($relsXml)) {
            $relsXml = "<Relationships xmlns=`"$pkgRelNs`"></Relationships>"
        }

        [xml]$relsDoc = $relsXml

        $relsMgr = New-Object System.Xml.XmlNamespaceManager($relsDoc.NameTable)
        $relsMgr.AddNamespace("r", $pkgRelNs)

        $relsRoot = $relsDoc.SelectSingleNode("/r:Relationships", $relsMgr)
        if (-not $relsRoot) { throw "Invalid _rels/.rels format inside package." }

        function Ensure-Relationship([string]$relType, [string]$target) {
            $existing = $relsDoc.SelectSingleNode("/r:Relationships/r:Relationship[@Type='$relType' and @Target='$target']", $relsMgr)
            if ($existing) { return }

            $ids = @()
            foreach ($rel in $relsDoc.SelectNodes("/r:Relationships/r:Relationship", $relsMgr)) { $ids += $rel.Id }

            $n = 1
            while ($ids -contains ("rId" + $n)) { $n++ }
            $newId = "rId$n"

            $relElem = $relsDoc.CreateElement("Relationship", $pkgRelNs)
            $relElem.SetAttribute("Id", $newId)
            $relElem.SetAttribute("Type", $relType)
            $relElem.SetAttribute("Target", $target)
            $relsRoot.AppendChild($relElem) | Out-Null
        }

        Ensure-Relationship $uiRelType2006 $uiPath2006
        Ensure-Relationship $uiRelType2007 $uiPath14

        Write-ZipText $relsPath (Write-XmlNoDecl $relsDoc)

        # --- xl/_rels/workbook.xml.rels : ensure person relationship exists ---
        $wbRelsXml = Read-ZipText $wbRelsPath
        if ([string]::IsNullOrWhiteSpace($wbRelsXml)) {
            $wbRelsXml = "<Relationships xmlns=`"$pkgRelNs`"></Relationships>"
        }

        [xml]$wbRelsDoc = $wbRelsXml
        $wbRelsMgr = New-Object System.Xml.XmlNamespaceManager($wbRelsDoc.NameTable)
        $wbRelsMgr.AddNamespace("r", $pkgRelNs)

        $wbRelsRoot = $wbRelsDoc.SelectSingleNode("/r:Relationships", $wbRelsMgr)
        if (-not $wbRelsRoot) { throw "Invalid xl/_rels/workbook.xml.rels format inside package." }

        $existingPersonRel = $wbRelsDoc.SelectSingleNode("/r:Relationships/r:Relationship[@Type='$personRelType' and @Target='persons/person.xml']", $wbRelsMgr)
        if (-not $existingPersonRel) {
            $ids = @()
            foreach ($rel in $wbRelsDoc.SelectNodes("/r:Relationships/r:Relationship", $wbRelsMgr)) { $ids += $rel.Id }

            $n = 1
            while ($ids -contains ("rId" + $n)) { $n++ }
            $newId = "rId$n"

            $relElem = $wbRelsDoc.CreateElement("Relationship", $pkgRelNs)
            $relElem.SetAttribute("Id", $newId)
            $relElem.SetAttribute("Type", $personRelType)
            $relElem.SetAttribute("Target", "persons/person.xml")
            $wbRelsRoot.AppendChild($relElem) | Out-Null
        }

        Write-ZipText $wbRelsPath (Write-XmlNoDecl $wbRelsDoc)

        # --- [Content_Types].xml : ensure required overrides ---
        $ctXml = Read-ZipText $ctPath
        if ([string]::IsNullOrWhiteSpace($ctXml)) { throw "Missing [Content_Types].xml in package." }

        [xml]$ctDoc = $ctXml
        $ctMgr = New-Object System.Xml.XmlNamespaceManager($ctDoc.NameTable)
        $ctMgr.AddNamespace("t", $ctNs)

        $typesNode = $ctDoc.SelectSingleNode("/t:Types", $ctMgr)
        if (-not $typesNode) { throw "Invalid [Content_Types].xml format." }

        function Ensure-Override([string]$partName, [string]$contentType) {
            $existing = $ctDoc.SelectSingleNode("/t:Types/t:Override[@PartName='$partName']", $ctMgr)
            if ($existing) { return }

            $overrideElem = $ctDoc.CreateElement("Override", $ctNs)
            $overrideElem.SetAttribute("PartName", $partName)
            $overrideElem.SetAttribute("ContentType", $contentType)
            $typesNode.AppendChild($overrideElem) | Out-Null
        }

        # Do NOT declare customUI parts in [Content_Types].xml
        # (Known-good examples rely on the root-level ui/extensibility relationships only.)
        foreach ($pn in @('/customUI/customUI.xml','/customUI/customUi14.xml')) {
            $node = $ctDoc.SelectSingleNode("/t:Types/t:Override[@PartName='$pn']", $ctMgr)
            if ($node) { $node.ParentNode.RemoveChild($node) | Out-Null }
        }

        # Ensure the person part is declared
        Ensure-Override "/xl/persons/person.xml" $personCt

        Write-ZipText $ctPath (Write-XmlNoDecl $ctDoc)
    }
    finally {
        if ($zip) { $zip.Dispose() }
    }
}



$ErrorActionPreference = 'Stop'

$success = $false
$friendlyError = $null
$installedPath = $outXlam
$wasUpdateInstall = $false

try {
    $wasUpdateInstall = Test-Path $outXlam

    # Preflight: fail fast if Excel has the add-in file locked (avoids slow COM teardown).
    if (Test-FileLocked -Path $outXlam) {
        throw "ADDIN_IN_USE"
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1 # msoAutomationSecurityLow (enable macros during automation)

    # Create a new workbook
    $wb = $excel.Workbooks.Add()

    # Save as .xlsm first (required for VBProject access to persist cleanly)
    $wb.SaveAs($tempXlsm, $xlOpenXMLWorkbookMacroEnabled)

    # Import all .bas modules
    foreach ($f in $basFiles) {
        Write-Host ("- Importing VBA: " + $f.Name)
        $null = $wb.VBProject.VBComponents.Import($f.FullName)
    }
    
    Write-Host ""
    Write-Host "Finished importing VBA modules. Building..."

    # Save as .xlam to TEMP first (so we never corrupt/partially overwrite the live add-in)
    if (Test-Path $outXlamTemp) { Remove-Item $outXlamTemp -Force }
    $wb.SaveAs($outXlamTemp, $xlOpenXMLAddIn)

    # Close workbook before mutating the OpenXML package
    $wb.Close($false)
    $wb = $null

    Add-RibbonXToOpenXmlFile -OpenXmlPath $outXlamTemp -RibbonXml2006 $ribbonXml2006 -RibbonXml14 $ribbonXml14

    # Try to install into the Excel AddIns folder.
    if (Test-FileLocked -Path $outXlam) {
        throw "ADDIN_IN_USE"
    }

    Copy-Item -Path $outXlamTemp -Destination $outXlam -Force

    # Unblock output (helps if user downloaded repo zip)
    try {
        Unblock-File -Path $outXlam -ErrorAction SilentlyContinue
    } catch {}

    $success = $true
}
catch {
    if ($_.Exception.Message -eq 'ADDIN_IN_USE') {
        $friendlyError = "Cannot overwrite '$outXlam' because Excel is currently using it.

Close Excel (or disable/unload the add-in), then run INSTALL.cmd again."
    } else {
        # Prefer a concise message for non-technical users, but keep it truthful.
        $friendlyError = $_.Exception.Message
    }

    $success = $false
}
finally {
    Write-Host ""
    Write-Host "Cleaning up temporary files..."
    if ($wb -ne $null) { try { $wb.Close($false) } catch {} }
    if ($excel -ne $null) {
        try { $excel.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }

    # Cleanup temp artifacts
    if (Test-Path $tempXlsm) { try { Remove-Item $tempXlsm -Force } catch {} }
    if (Test-Path $outXlamTemp) { try { Remove-Item $outXlamTemp -Force } catch {} }
}

# ---- User-facing output AFTER cleanup (prevents "dead air" after messages) ----
Write-Host ""

if (-not $success) {
    Write-Host "--------------------------------------------"
    Write-Host "BUILD FAILED"
    Write-Host "--------------------------------------------"
    Write-Host ""
    Write-Host $friendlyError
    Write-Host ""
    exit 1
}

Write-Host "--------------------------------------------"
Write-Host "BUILD SUCCESS"
Write-Host "--------------------------------------------"
Write-Host ""
Write-Host "Add-in installed at:"
Write-Host "  $installedPath"
Write-Host ""

if ($wasUpdateInstall) {
    Write-Host "Update note:"
    Write-Host "  If the add-in was already enabled, you do NOT need to re-enable it."
    Write-Host "  Just restart Excel to load the updated version."
    Write-Host ""
    Write-Host "If you have never enabled it before, enable it once using:"
} else {
    Write-Host "To enable the add-in:" 
}

Write-Host "  1) Open Excel"
Write-Host "  2) File -> Options -> Add-ins"
Write-Host "  3) Manage: Excel Add-ins -> Go..."
Write-Host "  4) Browse... -> select the add-in file above"
Write-Host "  5) Click OK"


Write-Host ""
Write-Host "Tip: If Excel was open during install, restart Excel to be safe."
exit 0
