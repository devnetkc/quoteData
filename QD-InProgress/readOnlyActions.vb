Public Sub askToSaveAsTemp(FName As String)
    Dim msg As String

    msg = "This workbook was opened as Read Only." & vbNewLine &
        "Because your changes will not be saved to QuoteData - In Progress " &
        "would you like to close the workbook now?" & vbNewLine & vbNewLine &
        "To close QuoteData - In Progress select: ""Yes""" & vbNewLine &
        "To work from a temporary backup save select: ""No"""
    iReply = MsgBox(prompt:=msg,
            Buttons:=vbYesNo Or vbExclamation, Title:="This file is Read Only!")

    If iReply = vbYes Then

        Application.ScreenUpdating = False
        Application.DisplayAlerts = False

        ThisWorkbook.Close(False)

        Application.ScreenUpdating = True
        Application.DisplayAlerts = True

        Exit Sub

    ElseIf iReply = vbNo Then
        Dim xlExt As String
        xlExt = ".xlsm"
        If Not Len(Dir(FName & xlExt)) = 0 Then
            iReply2 = MsgBox(prompt:="Do you want to replace the current temporary save file that already exists?",
            Buttons:=vbYesNo Or vbExclamation, Title:="Temporary Save Detected")
            If iReply2 = vbNo Then FName = getLastTempSave(FName)
        End If

        wbSaveExecute FName, ThisWorkbook, False, False, "Workbook Open detected file as Read Only", 52

        Exit Sub

    Else 'They cancelled (VbCancel)

        Exit Sub
    End If
End Sub
Public Sub askToDelteTempSaves(detTmpSvFld As String)
    Dim msg As String

    msg = "Temporary saves have been detected from a Read Only session." & vbNewLine & vbNewLine &
        "Would you like to delete the temporary save files now?" & vbNewLine & vbNewLine &
        "Yes I want to delete those files now, SELECT: ""Yes""" & vbNewLine &
        "No, please show me those files in Windows Explorer, SELECT: ""No"""
    iReply = MsgBox(prompt:=msg,
            Buttons:=vbYesNo Or vbExclamation, Title:="Temporary Saves Detected!")

    If iReply = vbYes Then
        On Error GoTo 0
        Kill detTmpSvFld & "\*.*"
    Else
        Shell "explorer.exe " & detTmpSvFld, vbNormalFocus
    End If

End Sub
Private Function getLastTempSave(FName As String, Optional xlExt As String = ".xlsm") As String
    For i = 2 To 102
        If Len(Dir(FName & "_" & i & xlExt)) = 0 Then
            getLastTempSave = FName & "_" & i
            Exit Function
        End If
    Next i
    errorNotice errNum:=55, funName:="Attempt to get last temporary save file name"
    getLastTempSave = FName & "_probs-Ryan has been emailed" & xlExt
End Function
Private Sub delTempSaveFiles()
    Kill detTmpSvFld & "\*.*"
End Sub


