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
if ((Test-Path $outXlam) -and (-not $Force)) {
    throw "Add-in already exists: $outXlam (re-run with -Force to overwrite)"
}

$tempXlsm = Join-Path $env:TEMP ($AddInName + "_build.xlsm")
if (Test-Path $tempXlsm) { Remove-Item $tempXlsm -Force }

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
            id="btn0"
            size="large"
            label="Prepare Workbook"
            imageMso="TableInsert"
            supertip="Creates or refreshes the Tax Summary worksheet and pulls in the data needed for reporting.&#10;&#10;Does not modify your source data. It only generates the summary output."
            onAction="Button_PrepareWorkbook"/>
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



try {
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
        Write-Host ("Importing " + $f.Name)
        $null = $wb.VBProject.VBComponents.Import($f.FullName)
    }

    # Save as .xlam into AppData AddIns
    if (Test-Path $outXlam) { Remove-Item $outXlam -Force }
    $wb.SaveAs($outXlam, $xlOpenXMLAddIn)

    $wb.Close($false)

    Add-RibbonXToOpenXmlFile -OpenXmlPath $outXlam -RibbonXml2006 $ribbonXml2006 -RibbonXml14 $ribbonXml14


    # Unblock output (helps if user downloaded repo zip)
    try {
        Unblock-File -Path $outXlam -ErrorAction SilentlyContinue
    } catch {}

    Write-Host ""
    Write-Host "Built add-in:"
    Write-Host $outXlam
    Write-Host ""
    Write-Host "Next: In Excel -> File -> Options -> Add-ins -> Excel Add-ins -> Go... -> Browse -> select it"
}
finally {
    if ($wb -ne $null) { try { $wb.Close($false) } catch {} }
    if ($excel -ne $null) {
        try { $excel.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    if (Test-Path $tempXlsm) { Remove-Item $tempXlsm -Force }
}
