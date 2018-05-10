Public Sub See_Style_Count()



    Dim msg As String, numOfIssues As Long, iReply As Variant



    If ThisWorkbook.Styles.Count > 48 Then
        numOfIssues = ActiveWorkbook.Styles.Count - 47
        msg = "You have " & numOfIssues & " extra styles that needed cleaned." & vbNewLine &
        vbNewLine & "Would you like to clean them up now?"


        iReply = MsgBox(prompt:=msg,
            Buttons:=vbYesNo Or vbExclamation, Title:="Excess Styles Detected!")

        'Set cC = New clsMsgbox

        'With cC

        ' .Title = "Excess Styles Detected!"
        ' .Prompt = Msg
        ' .Icon = Exclamation + DefaultButton1
        ' .ButtonText1 = "Clean Workbook"
        ' .ButtonText2 = "No"
        ' iR = .MessageBox()

        'End With


    Else
        msg = "There are no extra styles to worry about at this time."

        iReply = MsgBox(prompt:=msg,
            Buttons:=vbInformation Or vbOKOnly, Title:="No Issues Detected")

    End If




    If iReply = vbYes Then

        Application.ScreenUpdating = False
        Application.DisplayAlerts = False

        Call repairWorkbookStyles(ThisWorkbook, 0, "See_Style_Count -> iReply Yes")

        Application.ScreenUpdating = True
        Application.DisplayAlerts = True

        Exit Sub

    ElseIf iReply = vbNo Then

        Exit Sub

    Else 'They cancelled (VbCancel)

        Exit Sub

    End If

    Exit Sub



End Sub

Public Sub Un_Protect_All_Sheets(wb As Workbook, Optional funInName As String = "")



    Dim Sh As Worksheet

    '*Dim myPassword As String
    '****myPassword = "password"
    '****

    For Each Sh In wb.Worksheets
        Sh.Unprotect 'Password:=myPassword
    Next Sh

    Exit Sub

End Sub
Public Sub Protect_All_Sheets(wb As Workbook, Optional funInName As String = "")



    Dim Sh As Worksheet

    '*Dim myPassword As String
    '****myPassword = "password"
    '****

    For Each Sh In wb.Worksheets
        Sh.Protect 'Password:=myPassword
    Next Sh


    Exit Sub


End Sub
Private Sub repairWorkbookStyles(wb As Workbook, Optional leaveUnlocked As Integer = 1, Optional funInName As String = "")



    ' Unlock all sheets to prevent errors
    Call Un_Protect_All_Sheets(wb, funInName & " -> repairWorkbookStyles")


    ' Now let's remove all those bad styles
    Dim sty As Style

    ' Frist way to reset styles to prevent major issue
    ' First, remove all styles other than Excel's own.
    ' they may have arrived from pasting from other workbooks
    For Each sty In wb.Styles



        If Not sty.BuiltIn Then sty.Delete
    Next

    ' Lock all sheets again for security

    If leaveUnlocked = 1 Then Call Protect_All_Sheets(wb, funInName & " -> repairWorkbookStyles")

    ' Save the changes made so the style fix caries over to future updates
    Application.EnableEvents = False
    wb.Save
    Application.EnableEvents = True


    Exit Sub

End Sub
