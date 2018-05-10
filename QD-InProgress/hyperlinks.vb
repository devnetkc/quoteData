Option Explicit
Private pCWB As Workbook, pSessionSaves As Byte, pLastSaveName As String
Public LastSaveName As String
Public pubWBjustOpened As Boolean, forceCloseItAll As Boolean
Public curWB As Workbook
Public Sub ckFormAndSave(nSave As Boolean, Optional cWB As Workbook, Optional wbJustOpened As Boolean = False)
    
    If cWB Is Nothing Then Set cWB = ThisWorkbook
    
    If IsEmpty(pSessionSaves) Then pSessionSaves = 0
    
    Set pCWB = cWB: pubWBjustOpened = wbJustOpened
    
    Call ckformating: If nSave Then execWBbackupSave

    If pubWBjustOpened Then pubWBjustOpened = False
    
    Set pCWB = Nothing

End Sub
Private Sub ckformating()
Dim msg As String, i As Integer, oldColRow As Range, stOldColRow As Long, endOldColRow As Long
Dim ckThis As Range, curWS As Worksheet

Set curWS = pCWB.Sheets("UPDATES")
With curWS

    Set oldColRow = .Range("A1:A" & .UsedRange.Rows.Count).Find("Old Column Widths", lookat:=xlPart)
    
    
    
    
    stOldColRow = oldColRow.Row
    endOldColRow = oldColRow.End(xlDown).Row - stOldColRow
    
    On Error GoTo 0
    
    For i = 1 To endOldColRow
        
        Set ckThis = oldColRow.Offset(i, 1)
        
        If Not ckThis.NumberFormat = "General" Or _
            Not ckThis.NumberFormatLocal = "General" Then
                
                .Activate
                MsgBox prompt:="This workbook is showing an issue with formating." & _
                            vbNewLine & "Please check for formatting errors within the workbook." & _
                            vbNewLine & vbNewLine & _
                            "If this problem persists, please contact your system administrator for help.", _
                        Buttons:=vbOKOnly Or vbCritical, Title:="Formatting Issues Detected"

                        ckThis.Select: End
        End If
        
        Set ckThis = Nothing
        
    Next i

End With

Set curWS = Nothing

End Sub
Private Sub execWBbackupSave()
    
    Call saveQuoteData("\\SVFS02\cfmInfo\cfm Documents\Amos Quote Program\In " & _
        "Progress and Old Quote Data", 666, "SYSTEM", Format(curDateTime, "MMM-d-yyyy"), Format(curDateTime, " h N S AMPM"), _
        "No Formatting errors found")
    
    If delOrNew Then
        If Not Len(Dir(pLastSaveName, vbDirectory)) = 0 Then _
                Kill pLastSaveName
    Else
            pSessionSaves = 1
    End If
        
    pLastSaveName = LastSaveName

End Sub
Private Function delOrNew() As Boolean
    
    delOrNew = False
    
    If pSessionSaves = 1 And Not pubWBjustOpened Then
        Dim res As Integer
        
        res = MsgBox(prompt:="Reminder: Sessions reset each time the workbook opens:" & _
                            vbNewLine & Space(5) & _
                            "You have already saved once this session, " & _
                            "would you like to delete your last saved backup copy for this session?" & _
                            vbNewLine & vbNewLine & "Note: Your new save has already completed successfully", _
                            Buttons:=vbYesNo Or vbQuestion, Title:="Delete last session backup save?")

        If res = vbYes Then delOrNew = True
    End If
End Function


Sub testers()
'MsgBox ThisWorkbook.Sheets("UPDATES").Range("A:A").Find("Old Column Widths", lookat:=xlPart).End(xlDown).Row
   ' ckformating ThisWorkbook
End Sub
