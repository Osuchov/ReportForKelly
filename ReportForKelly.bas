Attribute VB_Name = "ReportForKelly"
Option Explicit

Sub KellysDisputeReport()

Dim DisStart As Date, DisEnd As Date
Dim fd As Office.FileDialog
Dim disputeFile As String
Dim dis As Workbook, rep As Workbook
Dim ctrlSht As Worksheet
Dim disputes As Worksheet, disPivots As Worksheet
Dim pvtCache As PivotCache
Dim pivotDataSource As String

Set ctrlSht = ThisWorkbook.Sheets("Control")
Set fd = Application.FileDialog(msoFileDialogFilePicker)

'check date ranges in "Control" sheet
If ctrlSht.Range("B2") = "" Or ctrlSht.Range("C2") = "" Then
    MsgBox "Invalid date range."
    GoTo CleaningUp
End If

With fd
    .AllowMultiSelect = False
    .title = "Please select the dispute file."  'Set the title of the dialog box.
    .Filters.Clear                            'Clear out the current filters, and add our own.
    .Filters.Add "All Files", "*.*"
    If .Show = True Then                'show the dialog box. Returns true if >=1 file picked
        disputeFile = .SelectedItems(1) 'replace txtFileName with your textbox
    End If
End With

On Error GoTo ErrHandling       'turning off warnings
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False
Application.ScreenUpdating = False

DisStart = ctrlSht.Range("B2")
DisEnd = ctrlSht.Range("C2")

Set rep = ThisWorkbook
Set dis = Workbooks.Open(disputeFile)

'making sure that the file is opened correctly
Do Until dis.ReadOnly = False
    dis.Close
    Application.Wait Now + TimeValue("00:00:01")
    Set dis = Workbooks.Open(disputeFile)
Loop

If dis.Sheets(1).Name <> "Disputes" Then     'checks if 1st sheet is called "Disputes"
    MsgBox "This is not a valid dispute file. Check it and run macro again."
    GoTo CleaningUp
End If

Set disputes = dis.Sheets("Disputes")
Set disPivots = rep.Sheets("Disputes")
pivotDataSource = disputes.Range(Cells(1, 1), Cells(disputes.UsedRange.Rows.Count, 34)).Address(ReferenceStyle:=xlR1C1)

If disputes.FilterMode Then disputes.ShowAllData  'if filter in dispute file applied, turn it off

'############## Pivots #################
Call CreatePivotTable(dis, disputes, pivotDataSource, ctrlSht, "Model Pivot")
Call FilterPivotFieldByDateRange(ctrlSht.PivotTables("Model Pivot"), ctrlSht.PivotTables("Model Pivot").PivotFields("Dispute date"), DisStart, DisEnd)


'######     Disputes Per Week     ######
Call copyModelPivot(disPivots, "Disputes Per Week")
With disPivots.PivotTables("Disputes Per Week")
    .PivotFields("WeekMonthNo").Orientation = xlRowField
    .PivotFields("WeekMonthNo").Position = 1
    .AddDataField disPivots.PivotTables("Disputes Per Week").PivotFields("ShipmentNumber"), "Number of Disputes", xlCount
    .PivotFields("Dispute date").Orientation = xlPageField
    .PivotFields("Dispute date").Position = 1
    .CompactLayoutRowHeader = "Weeks"
End With
Call CreateChartForPivot(disPivots, disPivots.PivotTables("Disputes Per Week"))

'######     Disputes Per Carrier     ######
Call copyModelPivot(disPivots, "Disputes Per Carrier")
With disPivots.PivotTables("Disputes Per Carrier")
    .PivotFields("Carrier").Orientation = xlRowField
    .PivotFields("Carrier").Position = 1
    .AddDataField disPivots.PivotTables("Disputes Per Carrier").PivotFields("ShipmentNumber"), "Number of Disputes", xlCount
    .AddDataField disPivots.PivotTables("Disputes Per Carrier").PivotFields("ShipmentNumber"), "%", xlCount
    .PivotFields("%").Calculation = xlPercentOfTotal
    .CompactLayoutRowHeader = "Carriers"
End With
Call CreateChartForPivot(disPivots, disPivots.PivotTables("Disputes Per Carrier"))

'######     Disputes Per Freight Payer     ######
Call copyModelPivot(disPivots, "Disputes Per Freight Payer")
With disPivots.PivotTables("Disputes Per Freight Payer")
    .PivotFields("CC").Orientation = xlRowField
    .PivotFields("CC").Position = 1
    .AddDataField disPivots.PivotTables("Disputes Per Freight Payer").PivotFields("ShipmentNumber"), "Number of Disputes", xlCount
    .AddDataField disPivots.PivotTables("Disputes Per Freight Payer").PivotFields("ShipmentNumber"), "%", xlCount
    .PivotFields("%").Calculation = xlPercentOfTotal
    .CompactLayoutRowHeader = "Company Codes"
End With
Call CreateChartForPivot(disPivots, disPivots.PivotTables("Disputes Per Freight Payer"))

'######     Disputes Per Reason     ######
Call copyModelPivot(disPivots, "Disputes Per Reason")
With disPivots.PivotTables("Disputes Per Reason")
    .PivotFields("Dispute reason (short)").Orientation = xlRowField
    .PivotFields("Dispute reason (short)").Position = 1
    .AddDataField disPivots.PivotTables("Disputes Per Reason").PivotFields("ShipmentNumber"), "Number of Disputes", xlCount
    .AddDataField disPivots.PivotTables("Disputes Per Reason").PivotFields("ShipmentNumber"), "%", xlCount
    .PivotFields("%").Calculation = xlPercentOfTotal
    .CompactLayoutRowHeader = "Reasons"
End With
Call CreateChartForPivot(disPivots, disPivots.PivotTables("Disputes Per Reason"))

'######     Average dispute resolution time per month     ######
Call CreatePivotTable(dis, disputes, pivotDataSource, disPivots, "Avg dispute resolution time")
With ActiveSheet.PivotTables("Avg dispute resolution time")
    .PivotFields("Dispute closure date").Orientation = xlRowField
    .PivotFields("Dispute closure date").Position = 1
    .PivotFields("Dispute closure date").PivotFilters.Add2 Type:=xlAfter, Value1:="1984-11-25"
    .PivotFields("Dispute closure date").dataRange.Cells(1).Group Periods:=Array(False, False, False, False, True, False, True)
    .CompactLayoutRowHeader = "Months"
    .AddDataField ActiveSheet.PivotTables("Avg dispute resolution time").PivotFields("Days disputes open"), "Avg time [days]", xlCount
    .PivotFields("Avg time [days]").Function = xlAverage
    .PivotFields("Average of Days disputes open").Caption = "Avg time [days]"
    .PivotFields("Avg time [days]").NumberFormat = "0.00"
End With
Call CreateChartForPivot(disPivots, disPivots.PivotTables("Avg dispute resolution time"))

On Error GoTo ErrHandling

CleaningUp:
    On Error Resume Next
    Application.DisplayAlerts = True 'turning on warnings
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True
    'delete model pivot
    ctrlSht.Range("A7").Clear
    ctrlSht.PivotTables("Model Pivot").TableRange2.Clear
    Exit Sub
 
ErrHandling:
    MsgBox Err & ". " & Err.Description
Resume CleaningUp

End Sub

Sub KellysAdditionalCostReport()

Dim acStart As Date, acEnd As Date
Dim fd As Office.FileDialog
Dim acFile As String
Dim ac As Workbook, rep As Workbook
Dim ctrlSht As Worksheet
Dim acs As Worksheet, acPivots As Worksheet
Dim pvtCache As PivotCache
Dim pivotDataSource As String

Set ctrlSht = ThisWorkbook.Sheets("Control")
Set fd = Application.FileDialog(msoFileDialogFilePicker)

'check date ranges in "Control" sheet
If ctrlSht.Range("B3") = "" Or ctrlSht.Range("C3") = "" Then
    MsgBox "Invalid date range."
    GoTo CleaningUp
End If

With fd
    .AllowMultiSelect = False
    .title = "Please select the additional costs file."  'Set the title of the dialog box.
    .Filters.Clear                            'Clear out the current filters, and add our own.
    .Filters.Add "All Files", "*.*"
    If .Show = True Then                'show the dialog box. Returns true if >=1 file picked
        acFile = .SelectedItems(1) 'replace txtFileName with your textbox
    End If
End With

On Error GoTo ErrHandling       'turning off warnings
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False
Application.ScreenUpdating = False

acStart = ctrlSht.Range("B3")
acEnd = ctrlSht.Range("C3")

Set rep = ThisWorkbook
Set ac = Workbooks.Open(acFile)

'making sure that the file is opened correctly
Do Until ac.ReadOnly = False
    ac.Close
    Application.Wait Now + TimeValue("00:00:01")
    Set ac = Workbooks.Open(acFile)
Loop

If ac.Sheets(1).Name <> "Additional Costs" Then     'checks if 1st sheet is called "Additional Costs"
    MsgBox "This is not a valid additional costs file. Check it and run macro again."
    GoTo CleaningUp
End If

Set acs = ac.Sheets("Additional Costs")
Set acPivots = rep.Sheets("AdditionalCosts")
pivotDataSource = acs.Range(Cells(1, 1), Cells(acs.UsedRange.Rows.Count, 34)).Address(ReferenceStyle:=xlR1C1)

If acs.FilterMode Then acs.ShowAllData  'if filter in dispute file applied, turn it off

'############## Pivots #################
Call CreatePivotTable(ac, acs, pivotDataSource, ctrlSht, "Model Pivot")
Call FilterPivotFieldByDateRange(ctrlSht.PivotTables("Model Pivot"), ctrlSht.PivotTables("Model Pivot").PivotFields("Receiving date"), acStart, acEnd)

'######     Additional Costs Per Week     ######
Call copyModelPivot(acPivots, "Additional Costs Per Week")
With acPivots.PivotTables("Additional Costs Per Week")
    .PivotFields("WeekMonthNo").Orientation = xlRowField
    .PivotFields("WeekMonthNo").Position = 1
    .AddDataField acPivots.PivotTables("Additional Costs Per Week").PivotFields("(INFODIS) Shipment number"), "Number of Additionals", xlCount
    .PivotFields("Receiving date").Orientation = xlPageField
    .PivotFields("Receiving date").Position = 1
    .CompactLayoutRowHeader = "Weeks"
End With
Call CreateChartForPivot(acPivots, acPivots.PivotTables("Additional Costs Per Week"))

'######     Additional Costs Per Carrier     ######
Call copyModelPivot(acPivots, "Additional Costs Per Carrier")
With acPivots.PivotTables("Additional Costs Per Carrier")
    .PivotFields("Carrier").Orientation = xlRowField
    .PivotFields("Carrier").Position = 1
    .AddDataField acPivots.PivotTables("Additional Costs Per Carrier").PivotFields("(INFODIS) Shipment number"), "Number of Additionals", xlCount
    .AddDataField acPivots.PivotTables("Additional Costs Per Carrier").PivotFields("(INFODIS) Shipment number"), "%", xlCount
    .PivotFields("%").Calculation = xlPercentOfTotal
    .CompactLayoutRowHeader = "Carriers"
End With
Call CreateChartForPivot(acPivots, acPivots.PivotTables("Additional Costs Per Carrier"))

'######     Additional Costs Per Freight Payer     ######
Call copyModelPivot(acPivots, "Additional Costs Per Freight Payer")
With acPivots.PivotTables("Additional Costs Per Freight Payer")
    .PivotFields("CC").Orientation = xlRowField
    .PivotFields("CC").Position = 1
    .AddDataField acPivots.PivotTables("Additional Costs Per Freight Payer").PivotFields("(INFODIS) Shipment number"), "Number of Additionals", xlCount
    .AddDataField acPivots.PivotTables("Additional Costs Per Freight Payer").PivotFields("(INFODIS) Shipment number"), "%", xlCount
    .PivotFields("%").Calculation = xlPercentOfTotal
    .CompactLayoutRowHeader = "Company Codes"
End With
Call CreateChartForPivot(acPivots, acPivots.PivotTables("Additional Costs Per Freight Payer"))

'######     Additional Costs Per Cost Type     ######
Call copyModelPivot(acPivots, "Additional Costs Per Cost Type")
With acPivots.PivotTables("Additional Costs Per Cost Type")
    .PivotFields("Type of additional costs (see SOP) (red = not in matrix)").Orientation = xlRowField
    .PivotFields("Type of additional costs (see SOP) (red = not in matrix)").Position = 1
    .AddDataField acPivots.PivotTables("Additional Costs Per Cost Type").PivotFields("(INFODIS) Shipment number"), "Number of Additionals", xlCount
    .AddDataField acPivots.PivotTables("Additional Costs Per Cost Type").PivotFields("(INFODIS) Shipment number"), "%", xlCount
    .PivotFields("%").Calculation = xlPercentOfTotal
    .CompactLayoutRowHeader = "Cost Types"
End With
Call CreateChartForPivot(acPivots, acPivots.PivotTables("Additional Costs Per Cost Type"))

'######     Average Additional Costs resolution time per month     ######
'Call CreatePivotTable(ac, acs, pivotDataSource, ctrlSht, "Model Pivot")
Call CreatePivotTable(ac, acs, pivotDataSource, acPivots, "Avg Additional Costs resolution time")
With ActiveSheet.PivotTables("Avg Additional Costs resolution time")
    .PivotFields("Closure date (pre bill / rejected date)").Orientation = xlRowField
    .PivotFields("Closure date (pre bill / rejected date)").Position = 1
    .PivotFields("Closure date (pre bill / rejected date)").PivotFilters.Add2 Type:=xlAfter, Value1:="1984-11-25"
    .PivotFields("Closure date (pre bill / rejected date)").dataRange.Cells(1).Group Periods:=Array(False, False, False, False, True, False, True)
    .CompactLayoutRowHeader = "Months"
    .AddDataField ActiveSheet.PivotTables("Avg Additional Costs resolution time").PivotFields("Days open"), "Avg time [days]", xlCount
    .PivotFields("Avg time [days]").Function = xlAverage
    .PivotFields("Average of Days open").Caption = "Avg time [days]"
    .PivotFields("Avg time [days]").NumberFormat = "0.00"
End With
Call CreateChartForPivot(acPivots, acPivots.PivotTables("Avg Additional Costs resolution time"))


On Error GoTo ErrHandling

CleaningUp:
    On Error Resume Next
    Application.DisplayAlerts = True 'turning on warnings
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True
    
    'delete model pivot
    ctrlSht.Range("A7").Clear
    ctrlSht.PivotTables("Model Pivot").TableRange2.Clear
    Exit Sub
 
ErrHandling:
    MsgBox Err & ". " & Err.Description
Resume CleaningUp

End Sub


Sub CreatePivotTable(disputeFile As Workbook, dataRangeSheet As Worksheet, dataRange As String, targetSheet As Worksheet, PivotName As String)

Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String
Dim target As Long
Dim temprange As Long


'Determine the data range you want to create pivot on
SrcData = "[" & disputeFile.Name & "]" & dataRangeSheet.Name & "!" & dataRange

targetSheet.Activate
target = targetSheet.UsedRange.Rows(targetSheet.UsedRange.Rows.Count).row + 5   '5 rows space in between pivots
With targetSheet.Range("A" & target - 1)
    .Font.Bold = True
    .Value = PivotName
End With

'Where do you want Pivot Table to start?
StartPvt = targetSheet.Name & "!" & targetSheet.Range("A" & target).Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
Set pvtCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

'Create Pivot table from Pivot Cache
Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:=PivotName)

End Sub

Sub clearContents()

Dim rep As Workbook
Dim disPivots As Worksheet
Dim acPivots As Worksheet

Set rep = ThisWorkbook
Set disPivots = rep.Sheets("Disputes")
Set acPivots = rep.Sheets("AdditionalCosts")

disPivots.Columns("A:Z").Delete Shift:=xlToLeft
acPivots.Columns("A:Z").Delete Shift:=xlToLeft

End Sub

Public Function FilterPivotFieldByDateRange(pvt As PivotTable, pvtField As PivotField, dtFrom As Date, dtTo As Date)
'Filters pivot table by given date range
'If there is no date that fulfills the criteria, removes all row, column and value fields

Dim bTemp As Boolean, i As Long
Dim dtTemp As Date, sItem1 As String
Dim rowField As PivotField, columnField As PivotField, valueField As PivotField

With pvtField
    .Orientation = xlPageField
    .Position = 1
End With
    
On Error Resume Next

With pvtField
    'check if there's any date in the date range
    For i = 1 To .PivotItems.Count
        dtTemp = .PivotItems(i)
        bTemp = (dtTemp >= dtFrom) And (dtTemp <= dtTo)
        If bTemp Then
            sItem1 = .PivotItems(i)
            Exit For
        End If
    Next i

    If sItem1 = "" Then 'if there is no such date
        MsgBox "No items are within the specified dates: " & .Parent.Name
        
        'clear row fields, column fields and vaulue fields
        For Each rowField In .Parent.RowFields
            rowField.Orientation = xlHidden
        Next rowField
        
        For Each columnField In .Parent.ColumnFields
            columnField.Orientation = xlHidden
        Next columnField
        
        For Each valueField In .Parent.DataFields
            valueField.Orientation = xlHidden
        Next valueField
        
        Exit Function
    End If

    'Application.Calculation = xlCalculationManual
   '.Parent.ManualUpdate = True

    If .Orientation = xlPageField Then .EnableMultiplePageItems = True
    .PivotItems(sItem1).Visible = True
    
    'filter out dates from outside of the date range
    For i = 1 To .PivotItems.Count
        dtTemp = .PivotItems(i)
        If dtTemp < dtFrom Or dtTemp > dtTo Then
            .PivotItems(i).Visible = False
        End If
    Next i
End With
    
    'pvtField.Parent.ManualUpdate = False
    'Application.Calculation = xlCalculationAutomatic
End Function

Public Function CreateChartForPivot(sht As Worksheet, pivot As PivotTable)

Dim pvtRange As Range
Dim chartStartRow As Long

Set pvtRange = pivot.DataBodyRange
chartStartRow = pvtRange.Cells(1, 1).row - 3

Set pvtRange = pvtRange.Offset(0, -1).Resize(pvtRange.Rows.Count - 1, pvtRange.Columns.Count + 1)

With sht.ChartObjects.Add(Left:=Range("F" & chartStartRow).Left, Top:=Range("F" & chartStartRow).Top, Width:=400, Height:=250).Chart
    .ChartType = xlColumnClustered
    .SetSourceData Source:=pvtRange, PlotBy:=xlColumns
    .HasLegend = False
End With

End Function

Public Function copyModelPivot(targetSheet As Worksheet, PivotName As String)

Dim target As Long
Dim newPivot As PivotTable

targetSheet.Activate
target = targetSheet.UsedRange.Rows(targetSheet.UsedRange.Rows.Count).row + 5   '5 rows space in between pivots

Worksheets("Control").PivotTables("Model Pivot").TableRange2.Copy Destination:=targetSheet.Range("A" & target)

For Each newPivot In targetSheet.PivotTables
    If newPivot.Name Like "Pivot*" Then
        newPivot.Name = PivotName
        Exit For
    End If
Next newPivot

With targetSheet.Range("A" & target - 1)
    .Font.Bold = True
    .Value = PivotName
End With

End Function
