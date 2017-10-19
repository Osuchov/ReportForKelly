Attribute VB_Name = "ReportForKelly"
Option Explicit

Sub KellysDisputeReport()

Dim DisStart As Date, DisEnd As Date
Dim fd As Office.FileDialog
Dim disputeFile As String
Dim dis As Workbook, rep As Workbook
Dim disputes As Worksheet, disPivots As Worksheet
Dim pvtCache As PivotCache

Set fd = Application.FileDialog(msoFileDialogFilePicker)

With fd
    .AllowMultiSelect = False
    .Title = "Please select the dispute file."  'Set the title of the dialog box.
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

DisStart = Range("B2")
DisEnd = Range("C2")

Set rep = ThisWorkbook
Set dis = Workbooks.Open(disputeFile)

If dis.Sheets(1).Name <> "Disputes" Then     'checks if 1st sheet is called "Disputes"
    MsgBox "This is not a valid dispute file. Check it and run macro again."
    GoTo CleaningUp
End If

Set disputes = dis.Sheets("Disputes")
Set disPivots = rep.Sheets("Disputes")

If disputes.FilterMode Then disputes.ShowAllData  'if filter in dispute file applied, turn it off

disputes.UsedRange.AutoFilter Field:=2, Criteria1:=">=" & DisStart, Operator:=xlAnd, Criteria2:="<=" & DisEnd



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