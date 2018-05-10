Public Sub Mail_Selection_Range_Outlook_Body(sheetDate As String, sheetTimeEmail As String,
    wb As Workbook, Optional funInName As String)
    'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
    'Don't forget to copy the function RangetoHTML in the module.
    'Working in Excel 2000-2013
    Dim rng As Range, OutApp As Object, OutMail As Object,
        updateWS As Worksheet, lRow As Integer, rngMsg As String,
        sendTo As String, rngX0 As Range, rngX1 As Range,
        rngX2 As Range, rngX3 As Range

    Set rng = Nothing
    Set updateWS = wb.Sheets("UPDATES")
    
On Error GoTo ErrorOccured1

    With updateWS

        On Error GoTo ErrorOccured2 'resumeNext

        .Unprotect
        
        Set rngX0 = .Range("A1:A10000")
        Set rngX1 = rngX0.Find("SendTo", lookat:=xlPart)
        Set rngX2 = rngX0.Find("Enable Debug Mode", lookat:=xlPart)
        Set rngX3 = rngX0.Find("Debug SendTo", lookat:=xlPart)
        Set rngX4 = rngX0.Find("Auto Send Release E-Mail", lookat:=xlPart)
        
        If Not rngX4 Is Nothing Then sendTo = rngX4.Offset(0, 2).Text
        If Not rngX2 Is Nothing Then _
            If StrComp(rngX2.Offset(0, 1).Text, "True", vbTextCompare) = 0 Then _
                If Not rngX3 Is Nothing Then sendTo = rngX3.Offset(0, 2).Text

        lRow = .Range("A" & 5).End(xlDown).Row
        .Columns(3).ColumnWidth = 97
        .Columns(5).ColumnWidth = 10
    End With

    On Error GoTo ErrorOccured3

    Set rng = updateWS.Range("C6:C" & lRow).SpecialCells(xlCellTypeVisible)
    'ActiveSheet.Range("C5:E150").Copy
    'Set rng = wb.Sheets("UPDATES").Range("A" & 5).End(xlDown).Row
    'Set rng = Selection.SpecialCells(xlCellTypeVisible)
    'You can also use a fixed range if you want
    'Set rng = Sheets("YourSheet").Range("D4:D12").SpecialCells(xlCellTypeVisible)


    If rng Is Nothing Then
        MsgBox "The selection is not a range or the sheet is protected" &
               vbNewLine & "please correct and try again.", vbOKOnly
        Exit Sub
    End If

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    commonReleaseBody = "A new release of QuoteData is now available to you.  Please be sure you are using the latest release " &
        "by opening your QuoteData to the " & """" & "UPDATES" & """" & " tab and verifying the release date matches todays date, " & sheetTimeEmail &
        ".<br><br>" &
        "If you have any questions or see any issues with this release please let me know so I can help get them resolved.  " &
        "Today's release includes these major changes listed below.  " &
        "<br><br><b>Release Notes:</b><br>" & RangetoHTML(rng) & "<br />" & "<br />"

    On Error GoTo ErrorOccured4 ' resumeNext

    With OutMail
        .To = sendTo
        .cC = ""
        .BCC = ""
        .Subject = "New QuoteData release available on " & sheetDate & " at " & sheetTimeEmail
        .Display
        .HTMLBody = commonReleaseBody & .HTMLBody
        If StrComp(rngX4.Offset(0, 1).Text, "True", vbTextCompare) = 0 Then .Send
    End With


    On Error GoTo ErrorOccured5

    'With Application
    '    .EnableEvents = True
    '    .ScreenUpdating = True
    'End With

    On Error GoTo ErrorOccured6

    With updateWS

        .Columns(3).ColumnWidth = 8.57
        .Columns(5).ColumnWidth = 10
        .Protect
    End With


    GoTo ExitSub

ErrorOccured1:
    Call errorNotice(5, funName:=funInName &
        " -> Mail_Selection_Range_Outlook_Body -> location 1")

    GoTo ExitSub

ErrorOccured2:
    Call errorNotice(5, funName:=funInName &
        " -> Mail_Selection_Range_Outlook_Body -> location 2")

    Resume Next

ErrorOccured3:
    Call errorNotice(5, funName:=funInName &
        " -> Mail_Selection_Range_Outlook_Body -> location 3")

    GoTo ExitSub

ErrorOccured4:
    Call errorNotice(5, funName:=funInName &
        " -> Mail_Selection_Range_Outlook_Body -> location 4")

    Resume Next

ErrorOccured5:
    Call errorNotice(5, funName:=funInName &
        " -> Mail_Selection_Range_Outlook_Body -> location 5")

    GoTo ExitSub

ErrorOccured6:
    Call errorNotice(5, funName:=funInName &
        " -> Mail_Selection_Range_Outlook_Body -> location 6")

    GoTo ExitSub

ExitSub:
    Set updateWS = Nothing
    Set rng = Nothing
    Set rngX0 = Nothing
    Set rngX1 = Nothing
    Set rngX2 = Nothing
    Set rngX3 = Nothing
    Set rngX4 = Nothing
    
    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub


Private Function RangetoHTML(rng As Range, Optional funInName As String)
    ' Changed by Ron de Bruin 28-Oct-2006
    ' Working in Office 2000-2013
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    On Error GoTo ErrorOccured1

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)

        On Error GoTo ErrorOccured2

        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    On Error GoTo ErrorOccured3

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add(
         SourceType:=xlSourceRange,
         Filename:=TempFile,
         Sheet:=TempWB.Sheets(1).Name,
         Source:=TempWB.Sheets(1).UsedRange.Address,
         HtmlType:=xlHtmlStatic)
        .Publish(True)
    End With

    On Error GoTo ErrorOccured4

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=",
                          "align=left x:publishsource=")

    On Error GoTo ErrorOccured5

    'Close TempWB
    TempWB.Close SaveChanges:=False

    'Delete the htm file we used in this function
    Kill TempFile


    GoTo ExitFunction

ErrorOccured1:
    Call errorNotice(5, funName:=funInName &
        " -> RangetoHTML -> location 1")

    GoTo ExitFunction

ErrorOccured2:
    Call errorNotice(5, funName:=funInName &
        " -> RangetoHTML -> location 2")

    GoTo ExitFunction

ErrorOccured3:
    Call errorNotice(5, funName:=funInName &
        " -> RangetoHTML -> location 3")

    GoTo ExitFunction

ErrorOccured4:
    Call errorNotice(5, funName:=funInName &
        " -> RangetoHTML -> location 4")

    GoTo ExitFunction

ErrorOccured5:
    Call errorNotice(5, funName:=funInName &
        " -> RangetoHTML -> location 5")

    GoTo ExitFunction

ExitFunction:
    Set TempWB = Nothing
    Set ts = Nothing
    Set fso = Nothing


End Function


