Public Function CloseAllExcel(Optional SaveChanges as Boolean = True, Optional CloseExcel as Boolean = True)
'Will close all open Excel workbooks.
Dim xl As Excel.Application
Dim wb As Workbook
If isExcelOpen Then
    Set xl = GetObject(, "Excel.Application")
    xl.EnableEvents = False
    xl.AskToUpdateLinks = False
    xl.DisplayAlerts = False
    For Each wb In xl.Workbooks
        wb.Close (SaveChanges)
    Next wb
    xl.DisplayAlerts = True
    xl.EnableEvents = True
    xl.AskToUpdateLinks = True
	If CloseExcel Then
		xl.Quit
	End If
End If
CleanExitTask:
Set xl = Nothing
Set wb = Nothing
End Function


Public Function CloseExcelFile(ByVal FileNamePath As String, Optional SaveChanges as Boolean = True, Optional CloseExcel as Boolean = True)
'Will close specific Excel workbook.
Dim xl As Excel.Application
Dim wb As Workbook
If IsWorkbookOpen(FileNamePath) Then
    Set xl = GetObject(, "Excel.Application")
    xl.EnableEvents = False
    xl.AskToUpdateLinks = False
    xl.DisplayAlerts = False
    For Each wb In xl.Workbooks
        If InStr(FileNamePath, wb.Name) > 0 Then
            wb.Close (SaveChanges)
            Exit For
        End If
    Next wb
    xl.DisplayAlerts = True
    xl.EnableEvents = True
    xl.AskToUpdateLinks = True
	If CloseExcel Then
		xl.Quit
	End If
End If

CleanExitTask:
Set xl = Nothing
Set wb = Nothing
End Function

        
        
Public Function isExcelOpen() As Boolean
'Returns true if an instance of Excel is Open.
On Error GoTo ErrorHandler
Dim xl As Excel.Application
isExcelOpen = False
Do
    Set xl = GetObject(, "Excel.Application")
    Set xl = Nothing
    isExcelOpen = True
Loop Until xl Is Nothing
Exit Function

ErrorHandler:
If Err <> 429 Then
    MsgBox Err.Description
End If

CleanExitTask:
Set xl = Nothing
Set wb = Nothing
End Function




Public Function IsWorkbookOpen(ByVal FileNamePath As String) As Boolean
'Returns true if specified workbook is open.
Dim xl As Excel.Application
Dim wb As Workbook
If isExcelOpen Then
    Set xl = GetObject(, "Excel.Application")
    For Each wb In xl.Workbooks
        If InStr(FileNamePath, wb.Name) > 0 Then
            IsWorkbookOpen = True
            GoTo CleanExitTask
        Else
            IsWorkbookOpen = False
        End If
    Next wb
Else
    IsWorkbookOpen = False
End If

CleanExitTask:
Set xl = Nothing
Set wb = Nothing
End Function




Public Function IsUserCellChangesUncommited() As Boolean
'Determines if user has entered info into a cell, but has not yet commited the data. Returns True if it is uncommited. 
'A commit in Excel is determined by hitting enter or selecting another cell.
'if you are trying to pull info from an already open excel sheet you can use this first make sure there is no uncommited cells in the sheet.
Dim xl As Excel.Application
Dim wb As Workbook
If isExcelOpen Then
    Set xl = GetObject(, "Excel.Application")
    If xl.Interactive = False Then
        IsUserCellChangesUncommited = False
        GoTo CleanExitTask
    End If
    On Error GoTo ErrorHandler
    xl.Interactive = False
    xl.Interactive = True
    IsUserCellChangesUncommited = False
End If
ErrorHandler:
If Err <> 0 Then
    IsUserCellChangesUncommited = True
    GoTo CleanExitTask
End If
CleanExitTask:
Set xl = Nothing
Set wb = Nothing
End Function




Public Function OpenWorkbook(FileNamePath As String)
'Open an Excel Workbook
    Dim xl As Excel.Application
    Dim wb As Excel.Workbook

    If isExcelOpen Then
        Set xl = GetObject(, "Excel.Application")
    Else
        Set xl = New Excel.Application
    End If
    xl.Visible = False  'Not visible by default
    xl.Visible = True  'Not visible by default
    Set wb = xl.Workbooks.Open(FileName:=FileNamePath, ReadOnly:=False)
    AppActivate (wb.Application.Caption)

CleanExitTask:
Set xl = Nothing
Set wb = Nothing
End Function
