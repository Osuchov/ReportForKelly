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
'Application.ScreenUpdating = False

DisStart = ctrlSht.Range("B2")
DisEnd = ctrlSht.Range("C2")

Set rep = ThisWorkbook
Set dis = Workbooks.Open(disputeFile)

If dis.Sheets(1).Name <> "Disputes" Then     'checks if 1st sheet is called "Disputes"
    MsgBox "This is not a valid dispute file. Check it and run macro again."
    GoTo CleaningUp
End If

Set disputes = dis.Sheets("Disputes")
Set disPivots = rep.Sheets("Disputes")
pivotDataSource = disputes.Range(Cells(1, 1), Cells(disputes.UsedRange.Rows.Count, 34)).Address(ReferenceStyle:=xlR1C1)

If disputes.FilterMode Then disputes.ShowAllData  'if filter in dispute file applied, turn it off

'############## Pivots #################

'######     Disputes Per Week     ######
Call CreatePivotTable(dis, disputes, pivotDataSource, disPivots, "Disputes Per Week")
With ActiveSheet.PivotTables("Disputes Per Week")
    .PivotFields("WeekMonthNo").Orientation = xlRowField
    .PivotFields("WeekMonthNo").Position = 1
    .AddDataField ActiveSheet.PivotTables("Disputes Per Week").PivotFields("ShipmentNumber"), "Number of Disputes", xlCount
    .PivotFields("Dispute date").Orientation = xlPageField
    .PivotFields("Dispute date").Position = 1
    .CompactLayoutRowHeader = "Weeks"
End With
Call FilterPivotFieldByDateRange(ActiveSheet.PivotTables("Disputes Per Week").PivotFields("Dispute date"), DisStart, DisEnd)

'######     Disputes Per Carrier     ######
Call CreatePivotTable(dis, disputes, pivotDataSource, disPivots, "Disputes Per Carrier")
With ActiveSheet.PivotTables("Disputes Per Carrier")
    .PivotFields("Carrier").Orientation = xlRowField
    .PivotFields("Carrier").Position = 1
    .AddDataField ActiveSheet.PivotTables("Disputes Per Carrier").PivotFields("ShipmentNumber"), "Number of Disputes", xlCount
    .AddDataField ActiveSheet.PivotTables("Disputes Per Carrier").PivotFields("ShipmentNumber"), "%", xlCount
    .PivotFields("%").Calculation = xlPercentOfTotal
    .PivotFields("Dispute date").Orientation = xlPageField
    .PivotFields("Dispute date").Position = 1
    .CompactLayoutRowHeader = "Carriers"
End With
Call FilterPivotFieldByDateRange(ActiveSheet.PivotTables("Disputes Per Carrier").PivotFields("Dispute date"), DisStart, DisEnd)

'######     Disputes Per Freight Payer     ######
Call CreatePivotTable(dis, disputes, pivotDataSource, disPivots, "Disputes Per Freight Payer")
With ActiveSheet.PivotTables("Disputes Per Freight Payer")
    .PivotFields("CC").Orientation = xlRowField
    .PivotFields("CC").Position = 1
    .AddDataField ActiveSheet.PivotTables("Disputes Per Freight Payer").PivotFields("ShipmentNumber"), "Number of Disputes", xlCount
    .AddDataField ActiveSheet.PivotTables("Disputes Per Freight Payer").PivotFields("ShipmentNumber"), "%", xlCount
    .PivotFields("%").Calculation = xlPercentOfTotal
    .PivotFields("Dispute date").Orientation = xlPageField
    .PivotFields("Dispute date").Position = 1
    .CompactLayoutRowHeader = "Company Codes"
End With
Call FilterPivotFieldByDateRange(ActiveSheet.PivotTables("Disputes Per Freight Payer").PivotFields("Dispute date"), DisStart, DisEnd)


On Error GoTo ErrHandling

CleaningUp:
    On Error Resume Next
    Application.DisplayAlerts = True 'turning on warnings
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True
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

'temprange = dataRangeSheet.UsedRange.Rows.Count

'Determine the data range you want to create pivot on
'SrcData = "[" & disputeFile.Name & "]" & dataRangeSheet.Name & "!" & dataRangeSheet.UsedRange.Address(ReferenceStyle:=xlR1C1)
SrcData = "[" & disputeFile.Name & "]" & dataRangeSheet.Name & "!" & dataRange

targetSheet.Activate
target = targetSheet.UsedRange.Rows(targetSheet.UsedRange.Rows.Count).Row + 5   '5 rows space in between pivots
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

Public Function FilterPivotFieldByDateRange(pvtField As PivotField, dtFrom As Date, dtTo As Date)
'Filters pivot table by given date range
'If there is no date that fulfills the criteria, removes all row, column and value fields

    Dim bTemp As Boolean, i As Long
    Dim dtTemp As Date, sItem1 As String
    Dim rowField As PivotField, columnField As PivotField, valueField As PivotField
    
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
