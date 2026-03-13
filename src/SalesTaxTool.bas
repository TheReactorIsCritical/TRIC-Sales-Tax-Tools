Option Explicit

Public MyRibbon As IRibbonUI

Sub OnRibbonLoad(Ribbon As IRibbonUI)
  'customUI Callbackname in XML File "onLoad"
  Set MyRibbon = Ribbon
  
  'We can get the pointer to the ribbon only when the file is loaded.
  MyRibbon.ActivateTab "tabTRICSalesTaxes"
End Sub

Sub Button_PrepareWorkbook(control As IRibbonControl)
    format_tric_tax_workbook
End Sub

Sub Button_OpenSquarespaceAccounting(Optional control As IRibbonControl)
    Const url As String = "https://tric.squarespace.com/config/commerce/selling-tools/accounting"

    On Error GoTo CleanFail

    WB.FollowHyperlink Address:=url, NewWindow:=True
    Exit Sub

CleanFail:
    MsgBox "Couldn't open the Squarespace accounting page in your browser." & vbCrLf & vbCrLf & _
           "URL: " & url, vbExclamation, "TRIC Sales Tax Tools"
End Sub

sub Button_OpenGithubRepository(Optional control As IRibbonControl)
    Const url As String = "https://github.com/TheReactorIsCritical/TRIC-Sales-Tax-Tools"

    On Error GoTo CleanFail

    WB.FollowHyperlink Address:=url, NewWindow:=True
    Exit Sub

CleanFail:
    MsgBox "Couldn't open the tool's GitHub repository page in your browser." & vbCrLf & vbCrLf & _
           "URL: " & url, vbExclamation, "TRIC Sales Tax Tools"
End Sub

Public Sub format_tric_tax_workbook()
    On Error GoTo CleanFail

    ' IMPORTANT: This is an add-in macro. Use the user's active workbook,
    ' not ThisWorkbook (which points at the .xlam).
    Dim WB As Workbook
    Set WB = ActiveWorkbook

    If WB Is Nothing Then
        MsgBox "No active workbook found. Please open the exported tax workbook and try again.", vbExclamation, "TRIC Sales Tax Tools"
        Exit Sub
    End If

    If WB.Name = ThisWorkbook.Name Then
        MsgBox "Please click into the workbook you want to process (not the add-in) and try again.", vbExclamation, "TRIC Sales Tax Tools"
        Exit Sub
    End If

    If Not validate_workbook_data(WB) Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    create_basic_tables
    create_tax_jurisdiction_table
    create_detailed_taxes_table
    create_tax_summary_sheet

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    On Error GoTo 0
    
    With WB.Worksheets("Tax Summary")
        .Activate
        Application.GoTo .Range("A1"), Scroll:=True
    End With
    
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Something went wrong while formatting the tax workbook." & vbCrLf & vbCrLf & _
           "The workbook may be missing required data or be in an unexpected format." & vbCrLf & _
           "No changes were made.", vbCritical, "TRIC Sales Tax Tools"
    
    On Error GoTo 0
End Sub

Private Function WB() As Workbook
    'I'm hardcoding the workbook to be the active workbook so it will run on whatever
    'I have open. I might change that behavior in the future, so that's why I abstracted it here.
    Set WB = ActiveWorkbook
End Function

' ------------------------------
' Validation
' ------------------------------

Private Function validate_workbook_data(WB As Workbook) As Boolean
    Dim problems As Collection
    Set problems = New Collection

    RequireSheet WB, "Orders", problems
    RequireSheet WB, "Taxes", problems
    RequireSheet WB, "Sale Line Items", problems
    RequireSheet WB, "Shipping Line Items", problems


    If WorksheetExists(WB, "Orders") Then
        ' Used by sales and tax aggregation functions (Order ID joins) and date/state logic in formatting.
        RequireHeaders WB, "Orders", Array("Gross Sales", "Net Sales", "Shipping", "Taxes"), problems
    End If
    
    If WorksheetExists(WB, "Taxes") Then
        RequireHeaders WB, "Taxes", Array("Order ID", "Jurisdiction Description", "Amount", "Sale Line Item ID", "Shipping Line Item ID"), problems
    End If

    If WorksheetExists(WB, "Sale Line Items") Then
        RequireHeaders WB, "Sale Line Items", Array("Sale Line Item ID", "Net Sales"), problems
    End If

    If WorksheetExists(WB, "Shipping Line Items") Then
        RequireHeaders WB, "Shipping Line Items", Array("Shipping Line Item ID", "Shipping Amount"), problems
    End If

    If problems.Count > 0 Then
        ShowValidationProblems problems
        validate_workbook_data = False
    Else
        validate_workbook_data = True
    End If
End Function

Private Sub RequireSheet(WB As Workbook, sheetName As String, problems As Collection)
    If Not WorksheetExists(WB, sheetName) Then
        problems.Add "Missing required worksheet: '" & sheetName & "'"
    End If
End Sub

Private Function WorksheetExists(WB As Workbook, sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = WB.Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Sub RequireHeaders(WB As Workbook, sheetName As String, headers As Variant, problems As Collection)
    Dim ws As Worksheet
    Set ws = WB.Worksheets(sheetName)

    Dim header As Variant
    For Each header In headers
        If Not HeaderExistsInRow(ws, CStr(header), 1) Then
            problems.Add "Worksheet '" & sheetName & "' is missing column: '" & header & "'"
        End If
    Next header
End Sub

Private Function HeaderExistsInRow(ws As Worksheet, headerText As String, headerRow As Long) As Boolean
    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Dim col As Long
    For col = 1 To lastCol
        If Trim(ws.Cells(headerRow, col).value) = headerText Then
            HeaderExistsInRow = True
            Exit Function
        End If
    Next col

    HeaderExistsInRow = False
End Function

Private Sub ShowValidationProblems(problems As Collection)
    Dim msg As String
    msg = "This workbook can't be processed yet:" & vbCrLf & vbCrLf

    Dim p As Variant
    For Each p In problems
        msg = msg & "â€¢ " & p & vbCrLf
    Next p

    msg = msg & vbCrLf & "Please fix the issue(s) above and try again."

    MsgBox msg, vbExclamation, "TRIC Sales Tax Tools"
End Sub

Public Sub create_basic_tables()

    Dim Orders As Worksheet
    Dim Taxes As Worksheet
    Dim Shipping As Worksheet
    Dim Sales As Worksheet
    
    Set Orders = Sheets("Orders")
    Set Taxes = Sheets("Taxes")
    Set Shipping = Sheets("Shipping Line Items")
    Set Sales = Sheets("Sale Line Items")
    
    'Make sure that if this script runs twice it will not cause errors
    'because you can only create the tables once.
    On Error Resume Next
    
    Orders.ListObjects.Add(xlSrcRange, Orders.UsedRange, , xlYes).Name = "Orders"
    Taxes.ListObjects.Add(xlSrcRange, Taxes.UsedRange, , xlYes).Name = "Taxes"
    Shipping.ListObjects.Add(xlSrcRange, Shipping.UsedRange, , xlYes).Name = "Shipping"
    Sales.ListObjects.Add(xlSrcRange, Sales.UsedRange, , xlYes).Name = "Sales"
    
    On Error GoTo 0
    
End Sub

Public Sub create_tax_jurisdiction_table()
' This is the new way of getting the tax buckets. WA requires the sum of gross product revenue
' plus the money you collected for shipping expenses. This table gets that, the jurisdiction
' that each order belongs to, and the taxes as well for later sanity checking on taxes collected
' vs. what the state says you owe.

    Dim src As ListObject
    Dim ws As Worksheet
    Dim dict As Object
    Dim orderRange As Range
    Dim jurisdictionRange As Range
    Dim i As Long
    Dim orderId As String
    Dim jurisdictionDescription As String
    Dim rowCount As Long
    Dim outRng As Range
    Dim lo As ListObject
    Dim key As Variant
    Dim writeRow As Long
    
    ' Source table
    Set src = GetTable("Taxes")
    If src Is Nothing Then
        Err.Raise vbObjectError + 3000, , "Taxes table not found."
    End If
    
    If src.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 3001, , "Taxes table is empty."
    End If
    
    Set orderRange = src.ListColumns("Order ID").DataBodyRange
    Set jurisdictionRange = src.ListColumns("Jurisdiction Description").DataBodyRange
    
    ' Build dictionary: Order ID -> longest Jurisdiction Description
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To orderRange.Rows.Count
        orderId = CStr(orderRange.Cells(i, 1).value)
        jurisdictionDescription = CStr(jurisdictionRange.Cells(i, 1).value)
        
        If Not dict.Exists(orderId) Then
            dict.Add orderId, jurisdictionDescription
        ElseIf Len(jurisdictionDescription) > Len(dict(orderId)) Then
            dict(orderId) = jurisdictionDescription
        End If
    Next i
    
    ' Create or clear destination sheet
    Set ws = GetOrCreateWorksheet("Tax Jurisdiction Lookup")
    ws.Cells.Clear
    
    ' Write headers
    ws.Range("A1").value = "Order ID"
    ws.Range("B1").value = "Most Detailed Jurisdiction"
    ws.Range("C1").value = "State"
    ws.Range("D1").value = "County"
    ws.Range("E1").value = "Jurisdiction"
    ws.Range("F1").value = "Gross Plus Shipping"
    ws.Range("G1").value = "Taxes"
    
    ' Write dictionary contents
    writeRow = 2
    For Each key In dict.Keys
        ws.Cells(writeRow, 1).value = key
        ws.Cells(writeRow, 2).value = dict(key)
        writeRow = writeRow + 1
    Next key
    
    rowCount = dict.Count
    
    ' Remove existing table if present
    On Error Resume Next
    ws.ListObjects("TaxJurisdictionLookup").Unlist
    On Error GoTo 0
    
    ' Create table
    Set outRng = ws.Range("A1").Resize(rowCount + 1, 7)
    Set lo = ws.ListObjects.Add(xlSrcRange, outRng, , xlYes)
    lo.Name = "TaxJurisdictionLookup"
    lo.TableStyle = "TableStyleMedium2"
    
    ' Fill helper formulas
    If rowCount > 0 Then
        Dim sep As String
        sep = Application.International(xlListSeparator)
        
        ' State
        lo.ListColumns("State").DataBodyRange.Formula = _
            "=IFERROR(TEXTBEFORE(TEXTAFTER([@[Most Detailed Jurisdiction]],""STATE:""),"",""),"""")"
        
        ' County
        lo.ListColumns("County").DataBodyRange.Formula = _
            "=IFERROR(TEXTBEFORE(TEXTAFTER([@[Most Detailed Jurisdiction]],""COUNTY:""),"",""),"""")"
        
        ' Jurisdiction
        lo.ListColumns("Jurisdiction").DataBodyRange.Formula = _
            "=IF(ISNUMBER(SEARCH(""CITY:"",[@[Most Detailed Jurisdiction]])),IFERROR(TEXTBEFORE(TEXTAFTER([@[Most Detailed Jurisdiction]],""CITY:""),"",""),TEXTAFTER([@[Most Detailed Jurisdiction]],""CITY:"")),IF([@[County]]<>"""",[@County]&"" COUNTY UNINCORPORATED"",""""))"
        
        ' Gross Plus Shipping
        lo.ListColumns("Gross Plus Shipping").DataBodyRange.Formula = _
            "=XLOOKUP([@[Order ID]]&""" & """" & sep & _
            "Orders[Order ID]&""" & """" & sep & _
            "Orders[Gross Sales]" & sep & _
            "0)+XLOOKUP([@[Order ID]]&""" & """" & sep & _
            "Orders[Order ID]&""" & """" & sep & _
            "Orders[Shipping]" & sep & _
            "0)"
        
        ' Taxes
        lo.ListColumns("Taxes").DataBodyRange.Formula = _
            "=XLOOKUP([@[Order ID]]&""" & """" & sep & _
            "Orders[Order ID]&""" & """" & sep & _
            "Orders[Taxes]" & sep & _
            "0)"
        
        ' Put them in currency format
        lo.ListColumns("Gross Plus Shipping").DataBodyRange.NumberFormat = Application.International(xlCurrencyCode) & "#,##0.00"
        lo.ListColumns("Taxes").DataBodyRange.NumberFormat = Application.International(xlCurrencyCode) & "#,##0.00"
    End If
    
    ws.Columns.AutoFit
End Sub

Public Sub create_detailed_taxes_table()
' NOTE: This table is retired in place. It makes a detailed breakdown of where each tax
' collection comes from for every order. This is useful to see where the different taxes
' are going to, but it isn't useful for remitting sales tax to the state. The state just
' buckets these categories, you file your revenue in that bucket, and then they distribute
' it accordingly on their end. This table is left in place in case this method changes
' in the future.

    Dim src As ListObject
    Dim ws As Worksheet
    Dim nRows As Long, nCols As Long
    Dim outRng As Range
    Dim lo As ListObject
    
    ' Source table
    Set src = GetTable("Taxes")
    If src Is Nothing Then
        Err.Raise vbObjectError + 2000, , "Taxes table not found."
    End If
    
    ' Row count (handle empty table)
    If src.DataBodyRange Is Nothing Then
        nRows = 0
    Else
        nRows = src.DataBodyRange.Rows.Count
    End If
    
    ' Destination sheet (create if missing, otherwise clear)
    Set ws = GetOrCreateWorksheet("DetailedTaxes")
    ws.Cells.Clear

    ' Output columns (5 copied + 4 calculated)
    nCols = 9

    ' Create the output range including header row (+ data rows)
    Set outRng = ws.Range("A1").Resize(nRows + 1, nCols)
    
    ' Headers
    outRng.Cells(1, 1).value = "Order ID"
    outRng.Cells(1, 2).value = "Jurisdiction Description"
    outRng.Cells(1, 3).value = "Amount"
    outRng.Cells(1, 4).value = "Sale Line Item ID"
    outRng.Cells(1, 5).value = "Shipping Line Item ID"
    outRng.Cells(1, 6).value = "Sale Revenue"
    outRng.Cells(1, 7).value = "Shipping Revenue"
    outRng.Cells(1, 8).value = "Is WA"
    outRng.Cells(1, 9).value = "Reporting Jurisdiction"
    
    ' Copy data columns (if there are rows)
    If nRows > 0 Then
        outRng.Cells(2, 1).Resize(nRows, 1).value = GetTableColumnValues(src, "Order ID")
        outRng.Cells(2, 2).Resize(nRows, 1).value = GetTableColumnValues(src, "Jurisdiction Description")
        outRng.Cells(2, 3).Resize(nRows, 1).value = GetTableColumnValues(src, "Amount")
        outRng.Cells(2, 4).Resize(nRows, 1).value = GetTableColumnValues(src, "Sale Line Item ID")
        outRng.Cells(2, 5).Resize(nRows, 1).value = GetTableColumnValues(src, "Shipping Line Item ID")
    End If
    
    ' Remove existing DetailedTaxes table if it exists (so reruns are safe)
    On Error Resume Next
    ws.ListObjects("DetailedTaxes").Unlist
    On Error GoTo 0
    
    ' Create the table
    Set lo = ws.ListObjects.Add(xlSrcRange, outRng, , xlYes)
    lo.Name = "DetailedTaxes"
    
    ' Add formulas to the two calculated columns (table will fill down automatically)
    If nRows > 0 Then
        ' Build the formulas using Excel's list separator instead of hardcoding commas.
        ' Without this, assigning the formula via VBA was throwing a runtime 1004 error.
        Dim sep As String
        sep = Application.International(xlListSeparator)
        
        lo.ListColumns("Sale Revenue").DataBodyRange.Formula = _
            "=XLOOKUP([@[Sale Line Item ID]]" & sep & _
            "Sales[Sale Line Item ID]" & sep & _
            "Sales[Net Sales]" & sep & _
            "0)"
            
        lo.ListColumns("Shipping Revenue").DataBodyRange.Formula = _
            "=XLOOKUP([@[Shipping Line Item ID]]" & sep & _
            "Shipping[Shipping Line Item ID]" & sep & _
            "Shipping[Shipping Amount]" & sep & _
            "0)"
            
        lo.ListColumns("Is WA").DataBodyRange.Formula = _
            "=ISNUMBER(SEARCH(""STATE:WA"",[@[Jurisdiction Description]]))"
            
        lo.ListColumns("Reporting Jurisdiction").DataBodyRange.Formula = _
            "=XLOOKUP([@[Order ID]],TaxJurisdictionLookup[Order ID],TaxJurisdictionLookup[Most Detailed Jurisdiction],"""")"
    
    End If
    
    
    ' Pretty formatting
    lo.TableStyle = "TableStyleMedium2"
    ws.Columns.AutoFit

End Sub

Public Sub create_tax_summary_sheet()
    Dim ws As Worksheet
    Dim r As Long

    Set ws = GetOrCreateWorksheetAtEnd("Tax Summary")

    ' Hard reset: wipe everything and rebuild from scratch
    ws.Cells.Clear
    ws.Cells.ClearFormats

    ' Basic layout
    ws.Range("A1").value = "Tax Summary"
    ws.Range("A2").value = "Generated:"
    ws.Range("B2").value = Now

    With ws.Range("A1:B1")
        .Merge
        .Font.Size = 18
        .Font.Bold = True
    End With

    ws.Range("A2").Font.Bold = True
    ws.Range("B2").NumberFormat = "m/d/yyyy h:mm AM/PM"

    r = 4

    ' ---- Totals section ----
    WriteSectionHeader ws, r, "Total"
    r = r + 1

    WriteLineItem ws, r, "Gross Sales", GrossSales(): r = r + 1
    WriteLineItem ws, r, "Net Sales", NetSales(): r = r + 1
    WriteLineItem ws, r, "Shipping Sales", ShippingSales(): r = r + 1
    WriteLineItem ws, r, "Retailing Gross Amount", RetailingGrossAmount(): r = r + 1

    ' Emphasize final line
    With ws.Range("A" & (r - 1) & ":B" & (r - 1))
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204) ' light highlight
    End With
    
    r = r + 1

    ' ---- Washington section ----
    WriteSectionHeader ws, r, "Washington"
    r = r + 1

    WriteLineItem ws, r, "Retailing Gross Amount", RetailingGrossAmount(): r = r + 1
    WriteLineItem ws, r, "Interstate / Foreign Apportionment", InterstateForeignApportionment("WA"): r = r + 1
    WriteLineItem ws, r, "Taxable Income (WA)", TaxableIncome("WA"): r = r + 1

    ' Emphasize final line
    With ws.Range("A" & (r - 1) & ":B" & (r - 1))
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204) ' light highlight
    End With
    
    r = r + 1

    ' ---- Tennessee section ----
    WriteSectionHeader ws, r, "Tennessee"
    r = r + 1

    WriteLineItem ws, r, "Retailing Gross Amount", RetailingGrossAmount(): r = r + 1
    WriteLineItem ws, r, "Interstate / Foreign Apportionment", InterstateForeignApportionment("TN"): r = r + 1
    WriteLineItem ws, r, "Taxable Income (TN)", TaxableIncome("TN"): r = r + 1
    
    ' Emphasize final line
    With ws.Range("A" & (r - 1) & ":B" & (r - 1))
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204) ' light highlight
    End With
    
    r = r + 1

    ' Styling for the whole summary block
    Dim lastRow As Long
    lastRow = r - 1

    With ws.Range("A4:B" & lastRow - 1)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    ws.Columns("A").ColumnWidth = 36
    ws.Columns("B").ColumnWidth = 18

    ' Currency formatting for value column (except header rows)
    ws.Range("B5:B" & lastRow).NumberFormat = "$#,##0.00"

    ' Spacer before pivot
    r = lastRow + 4
    ws.Range("A" & r).value = "Washington Tax Jurisdiction Pivot"
    ws.Range("A" & r).Font.Bold = True
    ws.Range("A" & r).Font.Size = 14

    r = r + 3

    ' Build the pivot table
    Dim lo As ListObject
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim dest As Range
    
    Set dest = ws.Range("A" & r)
    
    ' Clear any old pivot data if there is any; "z" is just a far column to catch everything
    ws.Range(dest, ws.Cells(ws.Rows.Count, "Z")).Clear
    
    Set lo = WB.Worksheets("Tax Jurisdiction Lookup").ListObjects("TaxJurisdictionLookup")

    ' Check for data. If there isn't any, the pivot table creation will fail because it needs at least 1 row
    If lo.DataBodyRange Is Nothing Then
        ws.Range("A" & r - 2).value = "No tax jurisdiction lookup entries were found."
        GoTo Cleanup
    End If
    
    ' Build pivot cache from the TaxJurisdictionLookup table range
    Set pc = WB.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=lo.Range)
    
    ' Create pivot table
    Set pt = pc.CreatePivotTable(TableDestination:=dest, tableName:="TaxSummaryPivot")
    
    ' --- Configure fields ---
    With pt
        .ManualUpdate = True
    
        ' Row fields (in order)
        .PivotFields("Jurisdiction").Orientation = xlRowField
        .PivotFields("Jurisdiction").Position = 1
    
        ' Values
        .AddDataField .PivotFields("Gross Plus Shipping"), "Sum of Gross Plus Shipping", xlSum
        .AddDataField .PivotFields("Taxes"), "Sum of Taxes", xlSum
        
        ' Filter to WA only
        With .PivotFields("State")
            .Orientation = xlPageField
            .Position = 1
            .CurrentPage = "WA"
        End With
    
        ' Make it readable
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
        .NullString = ""
        .DisplayErrorString = True
        .ErrorString = ""
    
        ' Number formatting for value fields
        .DataFields("Sum of Gross Plus Shipping").NumberFormat = "$#,##0.00"
        .DataFields("Sum of Taxes").NumberFormat = "$#,##0.00"
    
        .ManualUpdate = False
    End With
    
Cleanup:
    ' Autofit a reasonable width around the pivot area
    ws.Columns("A:H").AutoFit
    
End Sub

Private Sub WriteSectionHeader(ws As Worksheet, rowNum As Long, title As String)
    With ws.Range("A" & rowNum & ":B" & rowNum)
        .Merge
        .value = title
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242) ' soft section header
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
End Sub

Private Sub WriteLineItem(ws As Worksheet, rowNum As Long, label As String, value As Double)
    ws.Range("A" & rowNum).value = label
    ws.Range("B" & rowNum).value = value
End Sub

Private Function GetOrCreateWorksheetAtEnd(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWorksheetAtEnd = WB.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateWorksheetAtEnd Is Nothing Then
        Set GetOrCreateWorksheetAtEnd = WB.Worksheets.Add(After:=WB.Worksheets(WB.Worksheets.Count))
        GetOrCreateWorksheetAtEnd.Name = sheetName
    Else
        ' Ensure it's last tab (optional)
        GetOrCreateWorksheetAtEnd.Move After:=WB.Worksheets(WB.Worksheets.Count)
    End If
End Function

Private Function GetOrCreateWorksheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWorksheet = WB.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateWorksheet Is Nothing Then
        Set GetOrCreateWorksheet = WB.Worksheets.Add(After:=WB.Worksheets(WB.Worksheets.Count))
        GetOrCreateWorksheet.Name = sheetName
    End If
End Function

Private Function GetTableColumnValues(lo As ListObject, colName As String) As Variant
    Dim rng As Range
    Set rng = lo.ListColumns(colName).DataBodyRange

    If rng Is Nothing Then
        ' Return a 0-row empty variant; caller guards with nRows anyway.
        GetTableColumnValues = Empty
    Else
        GetTableColumnValues = rng.value
    End If
End Function

Public Function GetTable(tableName) As ListObject
    'Assumes worksheet name == table name
    
    Dim ws As Worksheet
    
    On Error GoTo Fail
    Set ws = WB.Worksheets(tableName)
    Set GetTable = ws.ListObjects(tableName)
    Exit Function

Fail:
    Set GetTable = Nothing
    
End Function


Public Function GrossSales() As Double

    Dim rng As Range
    
    Set rng = GetTable("Orders").ListColumns("Gross Sales").DataBodyRange
    
    If rng Is Nothing Then
        GrossSales = 0
    Else
        GrossSales = Application.WorksheetFunction.Sum(rng)
    End If
    
End Function

Public Function NetSales() As Double

    Dim rng As Range
    
    Set rng = GetTable("Orders").ListColumns("Net Sales").DataBodyRange
    
    If rng Is Nothing Then
        NetSales = 0
    Else
        NetSales = Application.WorksheetFunction.Sum(rng)
    End If
    
End Function


Public Function ShippingSales() As Double

    Dim rng As Range
    
    Set rng = GetTable("Orders").ListColumns("Shipping").DataBodyRange
    
    If rng Is Nothing Then
        ShippingSales = 0
    Else
        ShippingSales = Application.WorksheetFunction.Sum(rng)
    End If
    
End Function



Public Function TotalTaxesCollected() As Double

    Dim rng As Range
    
    Set rng = GetTable("Orders").ListColumns("Taxes").DataBodyRange
    
    If rng Is Nothing Then
        TotalTaxesCollected = 0
    Else
        TotalTaxesCollected = Application.WorksheetFunction.Sum(rng)
    End If
    
End Function

Public Function StateGrossSales(stateCode As String) As Double
    Dim lo As ListObject
    Dim criteriaRange As Range
    Dim sumRange As Range

    stateCode = UCase$(Trim$(stateCode))  ' normalize inputs like " wa "

    Set lo = GetTable("Orders")
    If lo Is Nothing Then
        StateGrossSales = 0
        Exit Function
    End If

    Set criteriaRange = lo.ListColumns("Shipping Address Subdivision").DataBodyRange
    Set sumRange = lo.ListColumns("Gross Sales").DataBodyRange

    If criteriaRange Is Nothing Or sumRange Is Nothing Then
        StateGrossSales = 0
    Else
        StateGrossSales = Application.WorksheetFunction.SumIf(criteriaRange, stateCode, sumRange)
    End If
End Function


Public Function StateNetSales(stateCode As String) As Double
    Dim lo As ListObject
    Dim criteriaRange As Range
    Dim sumRange As Range

    stateCode = UCase$(Trim$(stateCode))  ' normalize inputs like " wa "

    Set lo = GetTable("Orders")
    If lo Is Nothing Then
        StateNetSales = 0
        Exit Function
    End If

    Set criteriaRange = lo.ListColumns("Shipping Address Subdivision").DataBodyRange
    Set sumRange = lo.ListColumns("Net Sales").DataBodyRange

    If criteriaRange Is Nothing Or sumRange Is Nothing Then
        StateNetSales = 0
    Else
        StateNetSales = Application.WorksheetFunction.SumIf(criteriaRange, stateCode, sumRange)
    End If
End Function

Public Function StateShippingSales(stateCode As String) As Double
    Dim lo As ListObject
    Dim criteriaRange As Range
    Dim sumRange As Range

    stateCode = UCase$(Trim$(stateCode))  ' normalize inputs like " wa "

    Set lo = GetTable("Orders")
    If lo Is Nothing Then
        StateShippingSales = 0
        Exit Function
    End If

    Set criteriaRange = lo.ListColumns("Shipping Address Subdivision").DataBodyRange
    Set sumRange = lo.ListColumns("Shipping").DataBodyRange

    If criteriaRange Is Nothing Or sumRange Is Nothing Then
        StateShippingSales = 0
    Else
        StateShippingSales = Application.WorksheetFunction.SumIf(criteriaRange, stateCode, sumRange)
    End If
End Function

Public Function StateTaxes(stateCode As String) As Double
    Dim lo As ListObject
    Dim criteriaRange As Range
    Dim sumRange As Range

    stateCode = UCase$(Trim$(stateCode))  ' normalize inputs like " wa "

    Set lo = GetTable("Orders")
    If lo Is Nothing Then
        StateTaxes = 0
        Exit Function
    End If

    Set criteriaRange = lo.ListColumns("Shipping Address Subdivision").DataBodyRange
    Set sumRange = lo.ListColumns("Taxes").DataBodyRange

    If criteriaRange Is Nothing Or sumRange Is Nothing Then
        StateTaxes = 0
    Else
        StateTaxes = Application.WorksheetFunction.SumIf(criteriaRange, stateCode, sumRange)
    End If
End Function

Public Function RetailingGrossAmount() As Double
    RetailingGrossAmount = GrossSales() + ShippingSales()
End Function

Public Function InterstateForeignApportionment(stateCode As String) As Double
    '(all non-WA gross sales) + (all non-WA shipping)
    InterstateForeignApportionment = GrossSales() - StateGrossSales(stateCode) + ShippingSales() - StateShippingSales(stateCode)
End Function

Public Function TaxableIncome(stateCode As String) As Double
    TaxableIncome = RetailingGrossAmount - InterstateForeignApportionment(stateCode)
End Function
