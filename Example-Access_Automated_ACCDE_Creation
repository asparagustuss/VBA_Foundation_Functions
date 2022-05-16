'Example - Access ACCDE creation automation using VBA Foundation Functions.
'The below is a demonstration using the VBA Foundation Functions to automate the creation of a completely 
'locked out ACCDE file that would be ready for end user front end deployment. This makes deploying end user front ends very fast.

'---Sub linked to button click even on Access form---
Private Sub Create_ACCDE_Click()
DoCmd.SetWarnings True
Dim db As DAO.Database
Set db = CurrentDb
Dim ErrorCheck As Boolean

If Force_CompileProject = True Then
    Call DeleteAllTempTables(False)
    Application.SetOption "Themed Form Controls", True  'always true
    Application.SetOption "Show Status Bar", True    'always true
    Application.SetOption "DesignWithData", False
    Call ChangeProperty ("AllowFullMenus", vbBoolean, False)
    Call ChangeProperty ("AllowShortcutMenus", vbBoolean, False)
    Call ChangeProperty ("ShowDocumentTabs", vbBoolean, False)
    Call ChangeProperty ("AllowDatasheetSchema", vbBoolean, False)
    Call ChangeProperty ("AllowBypassKey", DB_BOOLEAN, False)
    'Hide navigation Pane
    Call DoCmd.NavigateTo("acNavigationCategoryObjectType")
    Call DoCmd.RunCommand(acCmdWindowHide)

    
    
    'save accde code here
    ErrorCheck = Create_ACCDE_File(True)
    
    Application.SetOption "DesignWithData", True
    Call ChangeProperty ("AllowFullMenus", vbBoolean, True)
    Call ChangeProperty ("AllowShortcutMenus", vbBoolean, True)
    Call ChangeProperty ("ShowDocumentTabs", vbBoolean, True)
    Call ChangeProperty ("AllowDatasheetSchema", vbBoolean, True)
    Call ChangeProperty ("AllowBypassKey", DB_BOOLEAN, True)
    'Show navigation Pane
    Call DoCmd.SelectObject(acTable, "Settings", True)
    
    If ErrorCheck = False Then
        MsgBox "ACCDE Creation Failed"
    Else
        MsgBox "ACCDE/ACCDR File Saved"
    End If
Else
    MsgBox ("Code complie error. Debug code before making ACCDE")
End If

End Sub


'---------Dependent VBA Foundation Functions----------

Function ChangeProperty(strPropName As String, varPropType As Variant, varPropValue As Variant) As Integer
'Changes access option properties. Usefull when automating the creation of an ACCDE file
    Dim dbs As Object, prp As Variant
    Const conPropNotFoundError = 3270
    Set dbs = CurrentDb
    On Error GoTo Change_Err
    dbs.Properties(strPropName) = varPropValue
    ChangeProperty = True
Change_Bye:
	Set dbs = Nothing
    Exit Function
Change_Err:
    If Err = conPropNotFoundError Then ' Property not found.
        Set prp = dbs.CreateProperty(strPropName, _
            varPropType, varPropValue)
        dbs.Properties.Append prp
        Resume Next
    Else
        ' Unknown error.
        ChangeProperty = False
        Resume Change_Bye
    End If
End Function


Public Function Create_ACCDE_File(Optional ByVal ConvertToACCDR As Boolean = False, Optional RemoveDebug as boolean = False) As Boolean
'automatically creates an ACCDE File with appropriate settings and such for a clean accde view
    Dim app As New Access.Application
    Dim ACCDE_FilePath As String
    Dim ACCDE_FileName As String

    If RemoveDebug then		'VBA IDE Foundation. Attempts to remove debug rows from code in IDE. Get most if not all with the below.
		Call VBA_Modules_ReplaceAll(" debug.print", " 'debug.print")
		Call VBA_Modules_ReplaceAll("'On Error GoTo ErrorHandler", "On Error GoTo ErrorHandler")
	End If
    
    ACCDE_FileName = Left(CurrentProject.Name, Len(CurrentProject.Name) - 6)
    ACCDE_FilePath = CurrentProject.Path & "\For Release\"
    
    If FileExists(ACCDE_FilePath & ACCDE_FileName & ".accdb") Then Kill ACCDE_FilePath & ACCDE_FileName & ".accdb"
    If FileExists(ACCDE_FilePath & ACCDE_FileName & "_Archive.DONOTDISTRIBUTE") Then Kill ACCDE_FilePath & ACCDE_FileName & "_Archive.DONOTDISTRIBUTE"
    If FileExists(ACCDE_FilePath & ACCDE_FileName & ".accde") Then Kill ACCDE_FilePath & ACCDE_FileName & ".accde"
    If FileExists(ACCDE_FilePath & ACCDE_FileName & ".accdr") Then Kill ACCDE_FilePath & ACCDE_FileName & ".accdr"
    
    Call CopyFile(CurrentProject.FullName, ACCDE_FilePath, ACCDE_FileName & ".accdb")

    app.AutomationSecurity = 1 ' this is msoAutomationSecurityLow -- can't figure out what the library constant is, so I just used 1.
    app.SysCmd 603, ACCDE_FilePath & ACCDE_FileName & ".accdb", ACCDE_FilePath & ACCDE_FileName & ".accde"
    
    If Not FileExists(ACCDE_FilePath & ACCDE_FileName & ".accde") Then
        Create_ACCDE_File = False
    Else
        Name ACCDE_FilePath & ACCDE_FileName & ".accdb" As ACCDE_FilePath & ACCDE_FileName & "_Archive.DONOTDISTRIBUTE"
        If ConvertToACCDR Then Name ACCDE_FilePath & ACCDE_FileName & ".accde" As ACCDE_FilePath & ACCDE_FileName & ".accdr"
        Create_ACCDE_File = True
    End If
End Function


Public Function FileExists(ByVal FileToTest As String) As Boolean
'Check if file exist
On Error GoTo ErrorHandler
   FileExists = (Dir(FileToTest, vbDirectory) <> "")
ErrorHandler:
If (Err.Number > 0) Then
    Select Case Err.Number
        Case 52
            FileExists = False
    End Select
    Resume Next
End If
End Function


Public Function Remove_Debugs()
Call VBA_Modules_ReplaceAll(" debug.print", " 'debug.print")
Call VBA_Modules_ReplaceAll("'On Error GoTo ErrorHandler", "On Error GoTo ErrorHandler")
End Function



Public Function VBA_Modules_ReplaceAll(ByVal StringToFind As String, _
    ByVal NewString, Optional ByVal FindWholeWord = False, _
    Optional ByVal MatchCase = False, Optional ByVal PatternSearch = False) as long
'Replace matched text in a VBA IDE for all modules
'This can be usefull to comment out debug lines and other task before programmatically compiling your code to something like an Access ACCDE file
'NOTE--this function should be placed in its own module. You should also change the exclusion module name seen in the code below to prevent this function from rewriting itself!
    Dim mdl As Module
    Dim lSLine As Long
    Dim lELine As Long
    Dim lSCol As Long
    Dim lECol As Long
    Dim sLine As String
    Dim lLineLen As Long
    Dim lBefore As Long
    Dim lAfter As Long
    Dim sLeft As String
    Dim sRight As String
    Dim sNewLine As String
    Dim TotalReplaced As Long
    Dim found As Boolean


TotalReplaced = 0
If StringToFind <> NewString Then       'prevents forever loop
    Dim intIndex As Integer
    Dim mods As Modules
    
    Set mods = Application.Modules
    
    For intIndex = 0 To mods.Count - 1

        If mods(intIndex).Name <> "VBA_IDE" Then		'you should change this to whatever module you places this function into so the funciton doesn't rewrite itself!!
            Set mdl = Modules(mods(intIndex).Name)
            Do
                lSCol = 0
                lELine = 0
                lECol = 0
                If mdl.Find(StringToFind, lSLine, lSCol, lELine, lECol, FindWholeWord, _
                        MatchCase, PatternSearch) = True Then
                    If IsMissing(NewString) = False Then
                        sLine = mdl.lines(lSLine, Abs(lELine - lSLine) + 1)
                        lLineLen = Len(sLine)
                        lBefore = lSCol - 1
                        lAfter = lLineLen - CInt(lECol - 1)
                        sLeft = Left$(sLine, lBefore)
                        sRight = Right$(sLine, lAfter)
                        sNewLine = sLeft & NewString & sRight
                        mdl.ReplaceLine lSLine, sNewLine
                    End If
                    TotalReplaced = TotalReplaced + 1
                End If
                lSLine = lELine
            Loop While lELine > 0
        End If
    Next
End If
VBA_Modules_ReplaceAll = TotalReplaced
Set mdl = Nothing
Set mods = Nothing
End Function
