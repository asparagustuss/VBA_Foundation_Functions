Public Function VBA_Module_ReplaceAll(ByVal ModuleName As String, ByVal StringToFind As String, _
        ByVal NewString, Optional ByVal FindWholeWord = False, _
        Optional ByVal MatchCase = False, Optional ByVal PatternSearch = False) As Long
'Replace matched text in a VBA IDE by Module Name
'This can be usefull to comment out debug lines and other task before programmatically compiling your code to something like an Access ACCDE file
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
If StringToFind <> NewString Then       'prevents forever lop
    Set mdl = Modules(ModuleName)
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
VBA_Modules_ReplaceAll = TotalReplaced
Set mdl = Nothing
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
