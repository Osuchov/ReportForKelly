Attribute VB_Name = "MergePreBills"
Option Explicit

Sub Merge()

Dim ws As Worksheet
Dim wsRoad As Worksheet, wsFCL As Worksheet, wsLCL As Worksheet, wsAir As Worksheet, wsRoadUS As Worksheet, wsALL As Worksheet
Dim wsCons As Worksheet 'consolidated worksheet
Dim file As String
Dim wb As Workbook
Dim wbMerge As Workbook
Dim directory As String
Dim counter As Long, allFiles As Long
Dim i As Integer
Dim target As Range     'where to paste
Dim pb As PreBill
Dim pbNum As Double
Dim fFree As Long, fFreeAll As Long
Dim sht As Variant
Dim arrSheets As Variant
Dim completed As Single
Dim preBills()

directory = pickDir("Pick the directory with Excel files to merge", "Merge")

If Len(directory) = 0 Then
    MsgBox "Directory not picked. Exiting...", vbExclamation
    Unload UserForm1
    Application.ScreenUpdating = True
    Exit Sub
End If

Application.ScreenUpdating = False

Set wbMerge = Application.ThisWorkbook  'naming sheets
Set wsRoad = wbMerge.Sheets("Road")
Set wsRoadUS = wbMerge.Sheets("Road US")
Set wsFCL = wbMerge.Sheets("FCL")
Set wsLCL = wbMerge.Sheets("LCL")
Set wsAir = wbMerge.Sheets("Air")
Set wsALL = wbMerge.Sheets("ALL")

file = Dir(directory & "*.xls")
allFiles = 0

Do While file <> ""
    allFiles = allFiles + 1
    file = Dir()
Loop

file = Dir(directory & "*.xls")

ReDim preBills(0 To 0)      'resetting preBills array

Do Until Len(file) = 0                  'loop on files to be merged
    counter = counter + 1
    completed = Round((counter * 100) / allFiles, 0)
    progress completed
        
    Set wb = Workbooks.Open(directory & file)
    Set ws = wb.Sheets(1)
    
    'Pre bills with status other than "Approved" should be ignored
    For i = 10 To i = 1 Step -1
        If Range("A" & i) = "Invoice status" And Range("B" & i) <> "Approved" Then
            GoTo Exception                              'move on to the next pre bill
        End If
        
        If Range("A" & i) = "Pre-bill Nr" Then          'dynamically check PB number (if Invoice status is "Approved")
            pbNum = CDbl(Range("B" & i))
            
            If IsInArray(pbNum, preBills) = False Then  'add pre bill number to a dynamic array
                preBills(UBound(preBills)) = pbNum      'check for doubles and ommit them if found
                ReDim Preserve preBills(0 To UBound(preBills) + 1)
            Else
                GoTo Exception
            End If
            
            Exit For
        End If
    Next i

    Set pb = New PreBill                'setting new pre bill object with attributes from file
    
    pb.CC = Range("C1")                 'setting company code
    pb.Number = pbNum
    pb.CarrierCode = Range("C2")    'setting the rest of pre bill attributes
    pb.CarrierName = Range("B2")
    pb.Status = Range("B9")
    pb.Vendor = Range("B5")
    pb.Period = Range("B3")
    pb.CreationDate = Range("B7")
    pb.NumberOfColumns = countColumns()
    pb.NumberOfRows = countRows()
    pb.StartRow = findStartRow()
    pb.Mode = ws.Name
    
    pb.Copy                         'copying of pre bill atributes
    
    If pb.Mode = "Road" Or pb.Mode = "Road Azkar" Then  'determining transport mode (pre bill template)
        fFree = firstFree(wsRoad)                       'checking first free cell in the correct sheet
        Set target = wsRoad.Cells(fFree, 9)             'setting pasting target
    ElseIf pb.Mode = "Road US" Then                     'teamplate for US Road pre bills (e.g. CA11)
        fFree = firstFree(wsRoadUS)
        Set target = wsRoadUS.Cells(fFree, 9)
    ElseIf pb.Mode = "FCL" Or pb.Mode = "Sea" Then
        fFree = firstFree(wsFCL)
        Set target = wsFCL.Cells(fFree, 9)
    ElseIf pb.Mode = "Air" Or pb.Mode = "Air 2" Then
        fFree = firstFree(wsAir)
        Set target = wsAir.Cells(fFree, 9)
    ElseIf pb.Mode = "Sea LCL" Then
        fFree = firstFree(wsLCL)
        Set target = wsLCL.Cells(firstFree(wsLCL), 9)
    Else
        MsgBox "WARNING! This workbook contains an unknown transport mode name: " & pb.Mode _
        & " Pre bill " & pb.Number & " (" & pb.CarrierCode & "/" & pb.CC & ") will not be merged!"
        counter = counter - 1
        GoTo Exception                                  'unknown transport mode exception
    End If
    
    target.PasteSpecial Paste:=xlPasteValuesAndNumberFormats    'paste the rest of pre bill data
    pb.PastePBData (fFree)                                      'to a first free cell
    
    fFreeAll = firstFree(wsALL)                         'pasting a whole pre bill file to "ALL" Sheet
    If fFreeAll = 2 Then
        Set target = wsALL.Cells(1, 1)
    Else
        Set target = wsALL.Cells(fFreeAll, 1)
    End If
    
    ws.UsedRange.Copy
    target.PasteSpecial Paste:=xlPasteAllExceptBorders
    
Exception:
    Application.CutCopyMode = False
    wb.Close False
    file = Dir
Loop

arrSheets = Array(Road, RoadUS, FCL, LCL, Air, ALL, check)

For Each sht In arrSheets
    sht.UsedRange.WrapText = False
Next sht

Unload UserForm1
Application.ScreenUpdating = True
MsgBox "Merging of " & counter & " Excel files completed." & vbNewLine & _
UBound(preBills) & " unique pre bill(s) was(were) found and processed.", vbInformation

End Sub

Sub progress(completed As Single)
UserForm1.Text.Caption = completed & "% Completed"
UserForm1.Bar.Width = completed * 2

DoEvents

End Sub


Sub clear_all()
Dim mb As Integer
Dim arrSheets As Variant
Dim sht As Variant

mb = MsgBox("You are about to clear all data from pre bill sheets." & Chr(13) & "Are you sure?", vbOKCancel + vbQuestion)

If mb = 1 Then
    arrSheets = Array(Road, RoadUS, FCL, LCL, Air, ALL, check)
    
    For Each sht In arrSheets
        If sht.Name = "ALL" Then
            sht.Activate
            sht.UsedRange.Select
        Else
            On Error Resume Next   'turn off error reporting
            ActiveSheet.ShowAllData
            sht.Activate
            sht.UsedRange.Offset(1).Select
            On Error GoTo 0
        End If
        Selection.EntireRow.Delete
        Cells(2, 1).Select
    Next sht
    
    MsgBox "Pre bill sheets are now empty."
Else
    MsgBox "Macro cancelled."
End If

End Sub

Sub CountGeneratedPreBills()
Dim arrSheets As Variant, sht As Variant
Dim check As Worksheet
Dim target As Range

Application.ScreenUpdating = False

Set check = Sheets("Check")
arrSheets = Array(Road, RoadUS, FCL, LCL, Air)

For Each sht In arrSheets
    Set target = check.Cells(countRowz(check, 1) + 1, 1)
    sht.Activate
    sht.Range(Cells(2, 1), Cells(countRowz(Sheets(sht.Name), 1) + 1, 1)).Select
    Selection.Copy
    target.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
Next sht

check.UsedRange.RemoveDuplicates columns:=(1), Header:=xlYes
check.Activate
check.Range("A1").Select

Application.ScreenUpdating = True

MsgBox ("There are " & countRowz(check, 1) - 1 & " pre bills.")

End Sub

Function pickDir(winTitle As String, buttonTitle As String) As String

Dim window As FileDialog
Dim picked As String

Set window = Application.FileDialog(msoFileDialogFolderPicker)
window.Title = winTitle
window.ButtonName = buttonTitle

If window.Show = -1 Then
    picked = window.SelectedItems(1)
    If Right(picked, 1) <> "\" Then
        pickDir = picked & "\"
    Else
        pickDir = picked
    End If
    
End If

End Function

Function timestamp() As String

timestamp = Format(Now(), "_yyyymmdd_hhmmss")

End Function

Function countColumns() As Long
    countColumns = ActiveSheet.UsedRange.columns.Count
End Function

Function countRows() As Long

Dim startPBBodyRow As Long

startPBBodyRow = findStartRow()
countRows = ActiveSheet.UsedRange.rows.Count - startPBBodyRow
    
End Function

Function findStartRow() As Long
Dim cell As Range

findStartRow = 12   'by default starting row should be 12. If it is not, the loop will find it

For Each cell In ActiveSheet.Range("A:A").Cells
    If cell.Value = "Referencenr" Then
        findStartRow = cell.row + 1
        Exit For
    End If
Next cell

End Function

Function firstFree(works As Worksheet) As Long
    works.Activate
    firstFree = ActiveSheet.UsedRange.rows.Count + 1
End Function

Function countRowz(ws As Worksheet, column As Long) As Long
'finds last used row (with header)

If ws.Cells(2, column) = "" Then
    countRowz = ws.Cells(1, column).row
Else
    countRowz = ws.Cells(1, column).End(xlDown).row
End If

End Function

Function IsInArray(stringToBeFound As Double, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function
