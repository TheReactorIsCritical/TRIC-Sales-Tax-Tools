Attribute VB_Name = "SalesTaxTool"
Option Explicit

Public MyRibbon As IRibbonUI

Sub OnRibbonLoad(Ribbon As IRibbonUI)
  'customUI Callbackname in XML File "onLoad"
  Set MyRibbon = Ribbon
  
  'We can get the pointer to the ribbon only when the file is loaded.
  MyRibbon.ActivateTab "tabTRICSalesTaxes"
End Sub

Sub Button_PrepareWorkbook(Control As IRibbonControl)
    format_tric_tax_workbook
End Sub

Public Sub format_tric_tax_workbook()
    'This is the main script that calls everything needed to prepare for quarterly
    'state sales tax submission.
    create_basic_tables
    create_detailed_taxes_table
    create_tax_summary_sheet

    With WB.Worksheets("Tax Summary")
        .Activate
        Application.GoTo .Range("A1"), Scroll:=True
    End With
    
End Sub

Private Function WB() As Workbook
    'I'm hardcoding the workbook to be the active workbook so it will run on whatever
    'I have open. I might change that behavior in the future, so that's why I abstracted it here.
    Set WB = ActiveWorkbook
End Function

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

Public Sub create_detailed_taxes_table()
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

    ' Output columns (5 copied + 2 calculated)
    nCols = 8

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

    r = r + 1

    ' ---- Washington section ----
    WriteSectionHeader ws, r, "Washington"
    r = r + 1

    WriteLineItem ws, r, "Gross Sales (WA)", StateGrossSales("WA"): r = r + 1
    WriteLineItem ws, r, "Net Sales (WA)", StateNetSales("WA"): r = r + 1
    WriteLineItem ws, r, "Shipping Sales (WA)", StateShippingSales("WA"): r = r + 1

    r = r + 1

    ' ---- Derived amounts ----
    WriteSectionHeader ws, r, "Derived"
    r = r + 1

    WriteLineItem ws, r, "Interstate Discount", InterstateDiscount(): r = r + 1
    WriteLineItem ws, r, "Retailing Gross Amount", RetailingGrossAmount(): r = r + 1
    WriteLineItem ws, r, "Interstate / Foreign Apportionment", InterstateForeignApportionment(): r = r + 1
    WriteLineItem ws, r, "Washington Taxable Income", WashingtonTaxableIncome(): r = r + 1

    ' Styling for the whole summary block
    Dim lastRow As Long
    lastRow = r - 1

    With ws.Range("A4:B" & lastRow)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    ws.Columns("A").ColumnWidth = 36
    ws.Columns("B").ColumnWidth = 18

    ' Currency formatting for value column (except header rows)
    ws.Range("B5:B" & lastRow).NumberFormat = "$#,##0.00"

    ' Emphasize final line (Washington Taxable Income)
    With ws.Range("A" & (lastRow) & ":B" & (lastRow))
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204) ' light highlight
    End With

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
    
    Set lo = WB.Worksheets("DetailedTaxes").ListObjects("DetailedTaxes")
    
    ' Build pivot cache from the DetailedTaxes table range
    Set pc = WB.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=lo.Range)
    
    ' Create pivot table
    Set pt = pc.CreatePivotTable(TableDestination:=dest, tableName:="TaxSummaryPivot")
    
    ' --- Configure fields ---
    With pt
        .ManualUpdate = True

        ' Row fields (in order)
        .PivotFields("Jurisdiction Description").Orientation = xlRowField
        .PivotFields("Jurisdiction Description").Position = 1

        .PivotFields("Order ID").Orientation = xlRowField
        .PivotFields("Order ID").Position = 2

        ' Values (in order)
        .AddDataField .PivotFields("Amount"), "Sum of Amount", xlSum
        .AddDataField .PivotFields("Shipping Revenue"), "Sum of Shipping Revenue", xlSum
        .AddDataField .PivotFields("Sale Revenue"), "Sum of Sale Revenue", xlSum

        ' Make it readable
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
        .NullString = ""
        .DisplayErrorString = True
        .ErrorString = ""

        ' Number formatting for value fields
        .DataFields("Sum of Amount").NumberFormat = "$#,##0.00"
        .DataFields("Sum of Shipping Revenue").NumberFormat = "$#,##0.00"
        .DataFields("Sum of Sale Revenue").NumberFormat = "$#,##0.00"
        
        With .PivotFields("Is WA")
            .Orientation = xlPageField
            .Position = 1
            .CurrentPage = "TRUE"
        End With

        .ManualUpdate = False
    End With
    
    
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

Public Function InterstateDiscount() As Double
    InterstateDiscount = NetSales() - StateGrossSales("WA")
End Function

Public Function RetailingGrossAmount() As Double
    RetailingGrossAmount = GrossSales() + ShippingSales()
End Function

Public Function InterstateForeignApportionment() As Double
    '(all non-WA gross sales) + (all non-WA shipping)
    InterstateForeignApportionment = GrossSales() - StateGrossSales("WA") + ShippingSales() - StateShippingSales("WA")
End Function

Public Function WashingtonTaxableIncome() As Double
    WashingtonTaxableIncome = RetailingGrossAmount - InterstateForeignApportionment()
End Function
