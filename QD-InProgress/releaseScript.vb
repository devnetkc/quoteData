Option Explicit On
Private use_altQD As Boolean
Private name_altQD As String
Public inRelease As Boolean
Public curDateTime As String

Public Sub releaseQD()

    inRelease = True
    curDateTime = DateTime.Now

    On Error GoTo ErrorOccured1

    Dim folderPath As String, sheetDate As String, workbookDir As String,
        inProgDir As String, wbPathCheck As Integer, q1 As Integer, q2 As Integer,
        q3 As String, qdRelName As String, origQD_updater As String,
        defNameCell As String, sheetTime As String, sheetTimeEmail As String,
        inProgWB As Workbook, stylesStart As Integer, stylesFinish As Integer,
        stylesCleaned As Integer, addToPath As Integer, debugCheck As Variant,
        upDateSheet As Worksheet, rngX As Range, debugTestValue As Integer,
        rngX2 As Range, rngX3 As Range, rngX4 As Range, autoRelease As Boolean,
        rngX5 As Range, autoSaveMode As Boolean, rngX6 As Range
        
        



'Range ("A1")

    Set inProgWB = ThisWorkbook
    Set upDateSheet = inProgWB.Sheets("UPDATES")
    
' Get the date and time that is likely an appliplical release time in 15 minute intervals
sheetDate = Format(curDateTime, "MMM-d-yyyy")
    sheetTimeEmail = Format((Application.WorksheetFunction.RoundUp(Now() * 96, 0) / 96), "h:mm AM/PM")
    sheetTime = Replace(sheetTimeEmail, ":", " ")
    'sheetTime = Format((Application.WorksheetFunction.RoundUp(Now() * 96, 0) / 96), "h mm AM/PM")

    ' Find out who is releasing quote data


    ' Default release name is stored in cell A3 on the Updates tab
    'On Error GoTo ErrorOccuredDefaultUpdater

    If Not upDateSheet Is Nothing Then

        On Error GoTo checkDefaultCells
            
            ' Set the Updates admin features range to check
                Set rngX = upDateSheet.Range("A1:A10000")
            
            ' Set the options to check on UPDATES sheet
                Set rngX2 = rngX.Find("Enable Debug Mode", lookat:=xlPart)
                Set rngX3 = rngX.Find("Managed By:", lookat:=xlPart)
                Set rngX4 = rngX.Find("Auto Replace Final Release", lookat:=xlPart)
                Set rngX5 = rngX.Find("End of Release Auto Save", lookat:=xlPart)
                Set rngX6 = rngX.Find("Use Alternate Release Name", lookat:=xlPart)
            
            ' Check if debugging is enabled
                If Not rngX2 Is Nothing Then
            debugCheck = rngX2.Offset(0, 1).Value
        End If

        ' Get the current primary handler of QD releases
        If Not rngX3 Is Nothing Then
            defNameCell = rngX3.Offset(1, 0).Text
        End If

        ' Check if auto releasing mode is enabled
        ' Set our default behavior to false until told otherwise
        autoRelease = False
        ' Check that the option even exists
        If Not rngX4 Is Nothing Then
            ' Now, if it exists and the value is True we will set autoRelease to a boolean True value
            If StrComp(rngX4.Offset(0, 1).Value, "TRUE", vbTextCompare) = 0 Then autoRelease = True
        End If
        ' Check if auto save mode is enabled
        ' Set our default behavior to false until told otherwise
        autoSaveMode = False
        ' Check that the option even exists
        If Not rngX5 Is Nothing Then
            ' Now, if it exists and the value is True we will set autoRelease to a boolean True value
            If StrComp(rngX5.Offset(0, 1).Value, "TRUE", vbTextCompare) = 0 Then autoSaveMode = True
        End If
        ' Check if alternate filename save is enabled
        ' Set our default behavior to false until told otherwise
        use_altQD = False
        ' Check that the option even exists
        If Not rngX6 Is Nothing Then
            ' Now, if it exists and the value is True we will set the alternate filename to use
            If StrComp(rngX6.Offset(0, 1).Value, "TRUE", vbTextCompare) = 0 Then
                name_altQD = rngX6.Offset(0, 2).Value
                use_altQD = True
            End If
        End If

    End If

    On Error GoTo ErrorOccured2

    ' If the cell is Blank then we have a problem and i need to know about it.
    If StrComp(defNameCell, "", vbTextCompare) = 0 Then Call Mail_Ryan_On_Errors(1, "releaseQD", Err.Description)
    ' Break down default name to only the first word
    If InStr(defNameCell, " ") > 0 Then defNameCell = firstWordOnly(defNameCell, "releaseQD -> defNameCell =")
    ' In case they mess this up, i'm manually adding Julie
    'origQD_updater = "Julie"
    qdRelName = defNameCell
    'If Not StrComp(defNameCell, origQD_updater, vbTextCompare) = 0 Then
    'qdRelName = origQD_updater
    ' End If


    ' Ask the release questions
    ' does the user want to release QD

    On Error GoTo ErrorOccured3

    q1 = UserPrompt(q_Number_In:=1, funInName:="releaseQD -> q1")

    If q1 = 0 Or q1 = 2 Then Exit Sub
    Application.EnableEvents = False
    inProgWB.Save
    Application.EnableEvents = True
    On Error GoTo ErrorOccured4

    ' Is the user releasing the default name set above?
    q2 = UserPrompt(2, qdRelName, sheetDate, "releaseQD -> q2")
    If q2 = 0 Then Exit Sub

    If Not q2 = 1 Then
        qdRelName = getUsersName(1, "releaseQD - > q2 was answered as false")

        On Error GoTo ErrorOccured5

        If StrComp(qdRelName, "", vbTextCompare) = 0 Then

            MsgBox "Your name is required to release QuoteData." & vbNewLine &
                    "The QuoteData release process has been canceled."
            Exit Sub

        End If
    End If

    ' We now have our releasor name and can continue on with the nitty gritty stuff

    ' Disable the annoying window prompts and Screen updates
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Before we really let the macro loose to start making big changes or do anything else drastic, we will save a backup in the temp Dir _
    ' \\SVFS02\cfmInfo\cfm Documents\Amos Quote Program\In Progress and Old Quote Data\Temp\Pre-release_auto-backups

    On Error GoTo ErrorOccured6

    ' Specifying the save location & name for the current In_Progress copy
    workbookDir = Application.ThisWorkbook.Path
    inProgDir = "\\SVFS02\cfmInfo\cfm Documents\Amos Quote Program\In Progress and Old Quote Data"
    wbPathCheck = StrComp(workbookDir, inProgDir, vbTextCompare)

    If Not wbPathCheck = 0 Then
        workbookDir = inProgDir
    End If


    ' Execute the save subs ---------



    On Error GoTo ErrorOccured7

    ' Execute the backup save pre-style fixing macros
    ' Makes an overwriting backup copy to fall back on in case the styles macro messes something up

    debugTestValue = StrComp(debugCheck, "True", vbTextCompare)

    If debugTestValue = 0 Then addToPath = 6

    Call saveQuoteData(workbookDir, 0 + addToPath, qdRelName, sheetDate, sheetTime, "releaseQD -> execute save subs -> pre-style backup", debugTestValue)

    ' Now lets clean the styles up
    ' Count the number of styles when we started
    stylesStart = inProgWB.Styles.Count
    ' Call the cleaning crew to tidy things up
    Call repairWorkbookStyles(inProgWB, 0, "releaseQD")
    ' Get the remaing number of styles
    stylesFinish = inProgWB.Styles.Count
    ' Store the number of styles cleaned
    stylesCleaned = stylesStart - stylesFinish

    ' Set default path additon to for functions and then adjust if debug mode is enabled
    addToPath = 0

    'debugCheck = inProgWB.Sheets("UPDATES").Range("B169").Value

    On Error GoTo ErrorOccured8

    If StrComp(debugCheck, "", vbTextCompare) = 0 Then Call Mail_Ryan_On_Errors(2, "releaseQD -> debugcheck is blank", Err.Description)

    If debugTestValue = 0 Then addToPath = 2

    On Error GoTo ErrorOccured9

    ' Save the xlsm to do all the macro work in and saveas from
    Call saveQuoteData(workbookDir, 2 + addToPath, qdRelName, sheetDate, sheetTime, "releaseQD -> execute save subs -> backup xlsm", debugTestValue)

    On Error GoTo ErrorOccured10

    ' Save the backup copy xlsm as xlsx now that it has finalized all workbook macros to prepair release
    Call saveQuoteData(workbookDir, 3 + addToPath, qdRelName, sheetDate, sheetTime, "releaseQD -> execute save subs -> save xlsx", debugTestValue, autoRelease)

    On Error GoTo ErrorOccured11

    ' Prepair the release email
    Call Mail_Selection_Range_Outlook_Body(sheetDate, sheetTimeEmail, inProgWB)

    On Error GoTo ErrorOccured12


    ' Clean up the In Progress workbook a little, and, if enabled, clear out the old update notes
    Call removeUpdateNotes(inProgWB, "debugCheck was false -> releaseQD", rngX)

    ' Do a final save and show completed msgbox
    Application.EnableEvents = False
    If autoSaveMode Then inProgWB.Save
    Application.EnableEvents = True

    MsgBox "Your release file is complete, and can be found here:" & vbNewLine &
        """" & workbookDir & "\Quote Data Releases\QuoteData - " & qdRelName & " released on " & sheetDate & " at " & sheetTime &
        """" & vbNewLine & vbNewLine & "File Health Update: " & stylesCleaned & " styles were removed."

' End of saving subs -------------



    ' Re-Enable the annoying window prompts and ScreenUpdates - ! important !
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ' Congratulations, you have release QD in one click!


    GoTo EndSub

ErrorOccured1:
    Call errorNotice(funName:="releaseQD -> location 1", errDesc:=Err.Description)

    GoTo EndSub

ErrorOccured2:
    Call errorNotice(funName:="releaseQD -> location 2", errDesc:=Err.Description)

    GoTo EndSub

ErrorOccured3:
    Call errorNotice(funName:="releaseQD -> location 3", errDesc:=Err.Description)

    GoTo EndSub

ErrorOccured4:
    Call errorNotice(funName:="releaseQD -> location 4", errDesc:=Err.Description)

    GoTo EndSub

ErrorOccured5:
    Call errorNotice(funName:="releaseQD -> location 5", errDesc:=Err.Description)

    GoTo EndSub

ErrorOccured6:
    Call errorNotice(funName:="releaseQD -> location 6", errDesc:=Err.Description)

    GoTo EndSub

ErrorOccured7:
    Call errorNotice(funName:="releaseQD -> location 7", errDesc:=Err.Description)

    GoTo EndSub

ErrorOccured8:
    Call errorNotice(6, funName:="releaseQD -> location 8", errDesc:=Err.Description)

    Call saveOpenCloseDelete()

    GoTo EndSub

ErrorOccured9:
    Call errorNotice(6, "releaseQD -> location 9", errDesc:=Err.Description)

    Call saveOpenCloseDelete()

    GoTo EndSub

ErrorOccured10:
    Call errorNotice(6, "releaseQD -> location 10", errDesc:=Err.Description)

    Call saveOpenCloseDelete()

    GoTo EndSub

ErrorOccured11:
    Call errorNotice(6, "releaseQD -> location 11", errDesc:=Err.Description)

    Call saveOpenCloseDelete()

    GoTo EndSub

ErrorOccured12:
    Call errorNotice(6, "releaseQD -> location 12", errDesc:=Err.Description)

    Call saveOpenCloseDelete()

    GoTo EndSub

checkDefaultCells:
    Call errorNotice(4, "releaseQD -> checkDefaultCells error", errDesc:=Err.Description)

    MsgBox "There was an error locating important default data.  " &
        "This error has been reported to your developing team " &
        "and the release process will now terminate"

    GoTo EndSub

EndSub:
    Set inProgWB = Nothing
    Set upDateSheet = Nothing
    Set rngX = Nothing
    Set rngX2 = Nothing
    Set rngX3 = Nothing
    Set rngX4 = Nothing
    Set rngX5 = Nothing
    Set rngX6 = Nothing
    inRelease = False


End Sub
' Get a file name based on the type of save execution we are performing
Public Sub saveQuoteData(FPath As String, nameFileNumber As Integer, qdRelName As String,
    Optional dateIn As String, Optional timeIn As String, Optional funInName As String = "",
    Optional debugTestValue As Integer, Optional autoRelease As Boolean = False)

    On Error GoTo ErrorOccured1

    Dim fileType As Integer, QD As String, FName As String, wbk As Workbook,
        xlsmQD_WorkBook As String, workFromQDwb As Workbook, debugPath As String,
        prePath As String, postPath As String, QDR As String, relPath As String
    
    ' Set the workbook high so we can change it in the cases later if needed
    Set wbk = ThisWorkbook
    ' Set default filetype save as 51, .xlsx - if file name is normal inprogress then we will save as .xlsm, aka 52
    fileType = 51
    If nameFileNumber = 0 Or nameFileNumber = 1 Or nameFileNumber = 2 Or nameFileNumber = 4 Then
        fileType = 52
    End If

    ' Set default resonse

    ' Set debugging root directory
    debugPath = "\RECOVERED\Ryan's test folder"
    ' Set release directory
    QDR = "\Quote Data Releases"
    ' Test if debugmode is enabled, and, if needed, add debug root to release path
    If debugTestValue = 0 Then QDR = debugPath & QDR
    ' Append incoming filepath argument to the front of the now established release dir
    QDR = FPath & QDR
    ' Start all QD release saves with this prefix
    QD = "\QuoteData - "
    ' Set the release path to use when saving the .xlsx release files
    relPath = QDR & QD
    ' Set the pre path to use when saving before doing any changes to origional file
    prePath = QDR & "\Pre-Clean Styles Backup" & QD
    ' Set the post path to use when performing actual release process and use as a backup
    postPath = QDR & "\Post-Clean Styles Backup" & QD

    'Debug.Print debugPath & vbNewLine & QDR & vbNewLine & debugTestValue _
    '    & vbNewLine & relPath & vbNewLine & prePath & vbNewLine & postPath

    'Application.EnableEvents = False
    ' Get the file name for backup and release copies

    'On Error GoTo ErrorOccured2
    On Error GoTo 0
    Select Case nameFileNumber

        Case 0, 6
            ' This is the savecopyas for the active QD In Progress.  We always want to have a backup
            ' to rely on before commiting any changes in case an error occurs and we need to reload from it.
            ' This backup will always be overwritten each time the release QD script starts
            FName = prePath & "single backup before style cleanup.xlsm"

            Call wbSaveExecute(FName, wbk, True, False, funInName & " -> saveQuoteData -> Case " & nameFileNumber)

        Case Is = 1
            ' Get the name for In Progress QuoteData saves
            FName = FPath & QD & "In Progress"

            Call wbSaveExecute(FName, wbk, True, False, funInName & " -> saveQuoteData -> Case " & nameFileNumber)

        Case 2, 4

            ' Get the name for backing up QuoteData;
            ' This file will also be used when performing the release prep and where the actual release file will be
            ' saved from.
            FName = postPath & "BackUp Release " & dateIn & " by " & qdRelName & ".xlsm"

            ' Call the save action to create the backup .xlsm file
            Call wbSaveExecute(FName, wbk, True, False, funInName & " -> saveQuoteData -> Case " & nameFileNumber)

        Case 3, 5
            ' This file that is being saved now will ultimatly be the file used when releasing QD

            ' Get the name for the release copy of QuoteData;
            FName = relPath & qdRelName & " released on " & dateIn & " at " & timeIn

            ' Set the path for the workbook to use when performing the actual release process work
            ' This will be the backup .xlsm file created in the select case above, Cases' 2 & 4
            xlsmQD_WorkBook = postPath & "BackUp Release " &
                dateIn & " by " & qdRelName & ".xlsm"
                    
                    ' Open the backup .xlsm file to use when performing the release prep work and saving the
                    ' release .xlsx file when finished.
                    Set workFromQDwb = Workbooks.Open(xlsmQD_WorkBook)
                        
                    ' This will get the backed up .xlsm file ready to release and save as a .xlsx
                    ' If this completes successfully then it will directly call the save action for the .xlsx
                    Call releaseQuoteDataPrep(workFromQDwb, "saveQuoteData -> Case " & nameFileNumber, FName, autoRelease, dateIn, QDR, qdRelName)

        Case 666
            ' Format error detected
            Dim opnd As String
            If pubWBjustOpened Then opnd = "--SESSION START--"

            FName = FPath & "\Pre-Release Backup Saves" &
                "\QuoteData - In Progress Pre-Save Backup_" & opnd &
                dateIn & " at " & timeIn & ".xlsm"



            ' Call the save action to create the backup .xlsm file
            Call wbSaveExecute(FName, wbk, True, False, funInName & " -> Open QD Save")

            LastSaveName = FName

    End Select


    'Application.EnableEvents = True


    GoTo ExitSub

ErrorOccured1:
    Call errorNotice(funName:=funInName & " -> saveQuoteData -> location 1", errDesc:=Err.Description)

    GoTo ExitSub

ErrorOccured2:
    Call errorNotice(funName:=funInName & " -> saveQuoteData -> location 2 -> case " & nameFileNumber, errDesc:=Err.Description)

    GoTo ExitSub

ExitSub:
    Set wbk = Nothing
    Set workFromQDwb = Nothing

End Sub
Public Sub wbSaveExecute(FName As String, wb As Workbook, saveAsCopy As Boolean,
    Optional closeAfter As Boolean = False, Optional funInName As String = "", Optional fileType As Integer = 51)

    On Error GoTo ErrorOccured1

    'Make sure the file doesn't exists or it won't save
    Dim tempFName As String
    tempFName = FName
    If saveAsCopy = True And fileType = 816 Then
        tempFName = FName
        If InStr(FName, ".xl") = 0 Then tempFName = FName & ".xlsx"
    End If
    If Not Len(Dir(tempFName)) = 0 Then Kill FName

    ' What kind of save are we performing
    On Error GoTo ErrorOccured2
    Application.EnableEvents = False
    If saveAsCopy = True Then
        wb.SaveCopyAs Filename:=FName
    Else
        wb.SaveAs Filename:=FName, FileFormat:=fileType, ReadOnlyRecommended:=False, CreateBackup:=False, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    End If

    If closeAfter = True Then wb.Close(True)
    Application.EnableEvents = True

    Exit Sub

ErrorOccured1:
    Call errorNotice(6, funName:=funInName & " -> wbSaveExecute -> location 1 -> " & FName, errDesc:=Err.Description)

    Exit Sub

ErrorOccured2:
    Call errorNotice(6, funName:=funInName & " -> wbSaveExecute -> location 2 -> " & FName, errDesc:=Err.Description)

    Exit Sub

End Sub

Private Function UserPrompt(q_Number_In As Integer, Optional defUserName As String,
    Optional sheetDate As String, Optional funInName As String = "") As Integer

    On Error GoTo ErrorOccured1

    Dim iReply As Integer, msg As String, boxHeader As String, buttonType As Integer

    ' No reason to change the header with the first 3 messages, but the option is available in the future if needed
    boxHeader = "Release QuoteData"
    buttonType = vbYesNoCancel
    On Error GoTo ErrorOccured2

    Select Case q_Number_In

        Case Is = 0
            msg = "Would you like to save QuoteData In Progress before starting and new changes?"
            boxHeader = "Save QuoteData?"
            buttonType = vbYesNo
        Case Is = 1
            msg = "Are you sure you want to prepare QuoteData for release?"
            boxHeader = "Start New QuoteData Release?"
            buttonType = vbYesNo + vbQuestion + vbDefaultButton2
        Case Is = 2
            msg = defUserName & " is the primary person responisible for releasing QuoteData at cfm." & vbNewLine & vbNewLine &
                "Is this " & defUserName & " doing a release of QuoteData on " & sheetDate & "?"
            boxHeader = "Is this " & defUserName & "?"
            buttonType = vbYesNo + vbQuestion + vbDefaultButton2
    End Select

    On Error GoTo ErrorOccured3

    UserPrompt = 0

    iReply = MsgBox(prompt:=msg,
            Buttons:=buttonType, Title:=boxHeader)



    If iReply = vbYes Then

        UserPrompt = 1
        Exit Function

        'Run "UpdateMacro"
        'UserPrompt = 1

        'End Function

    ElseIf iReply = vbNo Then

        UserPrompt = 2
        Exit Function

        'UserPrompt(3, sheetDate)

        'Do Other Stuff

    Else 'They cancelled (VbCancel)

        Exit Function

    End If


    Exit Function

ErrorOccured1:
    Call errorNotice(funName:=funInName & " -> UserPrompt -> location 1", errDesc:=Err.Description)

    Exit Function

ErrorOccured2:
    Call errorNotice(funName:=funInName & " -> UserPrompt -> location 2", errDesc:=Err.Description)

    Exit Function

ErrorOccured3:
    Call errorNotice(funName:=funInName & " -> UserPrompt -> location 3", errDesc:=Err.Description)

    Exit Function

End Function

Private Function getUsersName(MsgNum As Integer, Optional funInName As String = "") As String

    On Error GoTo ErrorOccured1

    Dim Title, defaultValue As String, msg As String
    'Dim myValue As Object

    ' Titles can be set later with more cases as needed, making default title for now
    Title = "Who is releasing QuoteData?"

    On Error GoTo ErrorOccured2

    Select Case MsgNum


        Case Is = 1
            msg = "To continue, please type your name in the space provided below." & vbNewLine &
                    vbNewLine & "Thank you"
    End Select

    On Error GoTo ErrorOccured3

    defaultValue = ""   ' Set default value.

    ' Display message, title, and default value.
    getUsersName = InputBox(prompt:=msg, Title:=Title, Default:=defaultValue)
    ' If user has clicked Cancel, set myValue to defaultValue
    'If getUsersName Is "" Then getUsersName = defaultValue
    If StrComp(getUsersName, "", vbTextCompare) = 0 Then
        getUsersName = defaultValue
    End If
    ' Display dialog box at position 100, 100.
    'myValue = InputBox(message, title, defaultValue, 100, 100)
    ' If user has clicked Cancel, set myValue to defaultValue
    'If myValue Is "" Then myValue = defaultValue


    Exit Function

ErrorOccured1:
    Call errorNotice(funName:=funInName & " -> getUsersName -> location 1", errDesc:=Err.Description)

    Exit Function

ErrorOccured2:
    Call errorNotice(funName:=funInName & " -> getUsersName -> location 2", errDesc:=Err.Description)

    Exit Function

ErrorOccured3:
    Call errorNotice(funName:=funInName & " -> getUsersName -> location 3", errDesc:=Err.Description)

    Exit Function

End Function
Private Sub repairWorkbookStyles(wb As Workbook, Optional leaveUnlocked As Integer = 1, Optional funInName As String = "")

    On Error GoTo ErrorOccured1

    ' Unlock all sheets to prevent errors
    Call Un_Protect_All_Sheets(wb, funInName & " -> repairWorkbookStyles")

    ' Make sure we have the correct number formatting for the "Normal" Style
    Call setDefaultNumberFormat(wb, funInName & " -> repairWorkbookStyles")

    ' Now let's remove all those bad styles
    Dim sty As Style

    ' Frist way to reset styles to prevent major issue
    ' First, remove all styles other than Excel's own.
    ' they may have arrived from pasting from other workbooks
    For Each sty In wb.Styles

        On Error GoTo ErrorOccured2

        If Not sty.BuiltIn Then sty.Delete
    Next

    ' Lock all sheets again for security
    On Error GoTo ErrorOccured3
    If leaveUnlocked = 1 Then Call Protect_All_Sheets(wb, funInName & " -> repairWorkbookStyles")

    ' Save the changes made so the style fix caries over to future updates
    Application.EnableEvents = False
    wb.Save
    Application.EnableEvents = True

    Exit Sub

ErrorOccured1:
    Call errorNotice(funName:=funInName & " -> repairWorkbookStyles -> location 1", errDesc:=Err.Description)

    Exit Sub
ErrorOccured2:
    Call errorNotice(funName:=funInName & " -> repairWorkbookStyles -> location 2", errDesc:=Err.Description)

    Exit Sub
ErrorOccured3:
    Call errorNotice(funName:=funInName & " -> repairWorkbookStyles -> location 3", errDesc:=Err.Description)

    Exit Sub
End Sub
Private Sub releaseQuoteDataPrep(wb As Workbook, Optional funInName As String = "", Optional FName As String,
    Optional autoRelease As Boolean = False, Optional dateIn As String, Optional FPath As String = "", Optional releaserName As String = "")
    Dim releaseDate As String

    On Error GoTo ErrorOccured1


    Dim Sh As Worksheet, upDateSheet As Worksheet

    If Application.ScreenUpdating = True Then
        Application.ScreenUpdating = False
    End If

    ' Unlock the sheets so we don't cause errors
    Call Un_Protect_All_Sheets(wb, funInName & " -> releaseQuoteDataPrep")

    ' Make sure all sheets have the topleft cell selected so users don't have to scroll
    ' Then place the sheet back in protected mode
    For Each Sh In wb.Worksheets

        On Error GoTo ErrorOccured2

        If Sh.Visible Then

            Sh.Activate

            If Not Sh.Tab.ColorIndex = xlColorIndexNone Then
                Sh.Delete

            Else

                If Not Sh.Range("A1").EntireColumn.Hidden Then
                    Sh.Range("A1").Select
                ElseIf Not Sh.Range("B1").EntireColumn.Hidden Then
                    Sh.Range("B1").Select
                Else
                    Sh.Range("C1").Select
                End If
                If Sh.Name = "UPDATES" Then
                    Sh.Range("A100:N600").Delete xlShiftUp
                    Sh.Shapes("exeQDrel").Delete
                    Sh.Shapes("cleanQD").Delete
                    Sh.Range("G2:G3").Clear
                    releaseDate = Format(curDateTime, "m/dd/yyyy")
                    Sh.Range("E2").Value = releaseDate
                    Sh.Columns(5).ColumnWidth = 10
                    ' Select the default cell users have selected when they open QD
                    ' This is usually the release date to try and imporove the chance users will notice
                    ' If they've missed an update or not
                    Sh.Range("E2").Select
                End If
                ActiveWindow.ScrollRow = 1
                ActiveWindow.ScrollColumn = 1
                Sh.Protect
            End If
        Else
            Sh.Delete
        End If

    Next Sh

    For Each Sh In wb.Worksheets

        On Error GoTo ErrorOccured3

        If Sh.Visible Then
            If Sh.Name = "UPDATES" Then
                Sh.Activate
                Exit For
            End If
        End If
    Next Sh

    On Error GoTo ErrorOccured4
    'Protect the workbook structure
    wb.Protect Structure:=True

    ' Save the changes you made to the .xlsm backup workbook before performing the .xlsx save action
    Application.EnableEvents = False
    wb.Save
    Application.EnableEvents = True

    On Error GoTo ErrorOccured5

    ' Call the save action to create the backup .xlsx file.
    ' All the prep work to this point should have been completed successfully
    ' If release prep, or this save process, fails, then the workbook will return
    ' to where the user started origionally-all previous saves will be undone/deleted
    Call wbSaveExecute(FName, wb, False, True, funInName & " -> releaseQuoteDataPrep")


    If autoRelease = True Then

        On Error GoTo ErrorOccured6

        ' Establish our new variables we'll need
        Dim dateInYr As String, dateInMon As String, dateReplaced As String,
            rootQDPath As String, oldQDroot As String,
            yrPath As String, monPath As String, xlsxExt As String,
            archiveCurrentQDFileWithExt As String,
            archiveCurrentQDFile As String, defDisplayAlert As Boolean, name_QD As String, pName_QD As String


        dateInYr = Format(dateIn, "yyyy")
        dateInMon = Format(dateIn, "mmm")
        dateReplaced = Format(dateIn, "mm-d-yyyy")
        rootQDPath = "\\SVFS02\cfmInfo\cfm Documents\Amos Quote Program"
        oldQDroot = FPath & "\Old Quote Versions"
        yrPath = oldQDroot & "\" & Format(dateIn, "yyyy")
        monPath = yrPath & "\" & Format(dateIn, "mmm")

        xlsxExt = ".xlsx"
        archiveCurrentQDFile = monPath & "\QuoteData - Replaced by " & releaserName & " on " & dateReplaced
        archiveCurrentQDFileWithExt = archiveCurrentQDFile & xlsxExt

        '' RETURN HERE

        ' Make sure the appropriate acrchive path exists via year then month
        On Error GoTo ErrorOccured6
        If Len(Dir(oldQDroot, vbDirectory)) = 0 Then MkDir oldQDroot
        oldQDroot = oldQDroot & "\" & Format(dateIn, "yyyy")
        If Len(Dir(oldQDroot, vbDirectory)) = 0 Then MkDir oldQDroot
        oldQDroot = oldQDroot & "\" & Format(dateIn, "mmm")
        If Len(Dir(oldQDroot, vbDirectory)) = 0 Then MkDir oldQDroot
    If Not Len(Dir(archiveCurrentQDFileWithExt)) = 0 Then _
            archiveCurrentQDFileWithExt = archiveCurrentQDFile &
            Format(curDateTime, " h N S AMPM") & xlsxExt

        ' Give the root QD release file a variable name
        pName_QD = rootQDPath & "\QuoteData"
        name_QD = pName_QD & xlsxExt
        ' Check if use alternate relase name is enabled -- if so, set it to that requested filename
        Dim cScreen As Boolean, cAlerts As Boolean, cEvents As Boolean
        cScreen = Application.ScreenUpdating
        cAlerts = Application.DisplayAlerts

        Application.ScreenUpdating = False
        Application.DisplayAlerts = False


        On Error GoTo ErrorOccured7

        Dim QD_WB As Workbook

        If use_altQD Then '& Format(curDateTime, "yy-hhmmss")
            name_QD = pName_QD & name_altQD & xlsxExt
            If Len(Dir(name_QD)) = 0 Then


                On Error Resume Next
                'Set QD_WB = Workbooks(pName_QD & xlsxExt)
                If QD_WB Is Nothing Then Set QD_WB = Workbooks.Open(pName_QD & xlsxExt)
                
            On Error GoTo ErrorOccured7

                Call wbSaveExecute(name_QD, QD_WB, True, False,
                    funInName & "releaseQuoteDataPrep -> fullAutoRelease:duplicateCurrentQD-ReleasedXLSX SaveCopyAs", 816)
                QD_WB.Close(False)
                Set QD_WB = Nothing
                
                'On Error GoTo AltSaveAs
                '    Workbooks(pName_QD).SaveCopyAs name_QD
                    'FileCopy pName_QD & xlsxExt, name_QD
                'On Error GoTo ErrorOccured6
                'Dim fso As Object
                'Set fso = VBA.CreateObject("Scripting.FileSystemObject")
                'Call fso.copyfile(pName_QD, name_QD)
            End If
        End If


        If Not Len(Dir(name_QD)) = 0 Then

            On Error Resume Next
            'Set QD_WB = Workbooks(name_QD)
            If QD_WB Is Nothing Then Set QD_WB = Workbooks.Open(name_QD)
            Call wbSaveExecute(archiveCurrentQDFileWithExt, QD_WB, True, False,
                    funInName & "releaseQuoteDataPrep -> fullAutoRelease:Archive SaveCopyAs")
            QD_WB.Close(False)
        Set QD_WB = Nothing

    On Error Resume Next
            'Set QD_WB = Workbooks(FName & xlsxExt)
            'Application.DisplayAlerts = True

            On Error GoTo ErrorOccured8

            If QD_WB Is Nothing Then Set QD_WB = Workbooks.Open(FName & xlsxExt)
            ' rootQDPath = rootQDPath & "\QuoteData.xlsx"

            Call wbSaveExecute(name_QD, QD_WB, True, False, funInName & "releaseQuoteDataPrep -> fullAutoRelease:FinalRelease SaveCopyAs", 816)
            QD_WB.Close(False)
            Application.DisplayAlerts = True
        Set QD_WB = Nothing
        


    End If

        Application.ScreenUpdating = cScreen
        Application.DisplayAlerts = cAlerts


    End If


    GoTo ExitSub

    'AltSaveAs:

    '    On Error GoTo ErrorOccured6

    '        Workbooks(pName_QD).SaveCopyAs name_QD

    '    Resume Next

ErrorOccured1:
    Call errorNotice(6, funName:=funInName & " -> releaseQuoteDataPrep -> location 1", errDesc:=Err.Description)

    GoTo ExitSub

ErrorOccured2:
    Call errorNotice(6, funName:=funInName & " -> releaseQuoteDataPrep -> location 2", errDesc:=Err.Description)

    GoTo ExitSub

ErrorOccured3:
    Call errorNotice(6, funName:=funInName & " -> releaseQuoteDataPrep -> location 3", errDesc:=Err.Description)

    GoTo ExitSub

ErrorOccured4:
    Call errorNotice(6, funName:=funInName & " -> releaseQuoteDataPrep -> location 4", errDesc:=Err.Description)

    GoTo ExitSub

ErrorOccured5:
    Call errorNotice(6, funName:=funInName & " -> releaseQuoteDataPrep -> location 5", errDesc:=Err.Description)

    GoTo ExitSub

ErrorOccured6:
    Call errorNotice(6, funName:=funInName & " -> releaseQuoteDataPrep -> location 6", errDesc:=Err.Description)

    GoTo ExitSub

ErrorOccured7:
    Call errorNotice(6, funName:=funInName & " -> releaseQuoteDataPrep -> location 7", errDesc:=Err.Description)

    GoTo ExitSub

ErrorOccured8:
    Call errorNotice(6, funName:=funInName & " -> releaseQuoteDataPrep -> location 8", errDesc:=Err.Description)
    
    Set QD_WB = Nothing
    'Set NEW_QD_WB = Nothing
    
    GoTo ExitSub

ExitSub:

End Sub
Private Sub copyfile(SourceFile As String, DestinationFile As String)

    'Application.ScreenUpdating = False
    'Define source file name, not including path.
    'SourceFile = "source.xls"
    'Define target file name. Include path if needed.
    'DestinationFile = "c:\destin.xls"
    'Copy source to target.


End Sub
Private Sub removeUpdateNotes(wb As Workbook, Optional funInName As String = "", Optional rngX As Range)

    Dim Sh As Worksheet, lRow As Integer

    On Error GoTo ErrorOccured1

    wb.Unprotect

    On Error GoTo ErrorOccured2

    For Each Sh In wb.Worksheets

        If Sh.Visible Then
            Sh.Activate
            Sh.Unprotect
            If Not Sh.Range("A1").EntireColumn.Hidden Then
                Sh.Range("A1").Select
            ElseIf Not Sh.Range("B1").EntireColumn.Hidden Then
                Sh.Range("B1").Select
            Else
                Sh.Range("C1").Select
            End If
            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = 1
        End If
    Next Sh

    On Error GoTo ErrorOccured3

    ' Now that you've completed the release file and email template, when enabled, this will clear the update notes
    ' Next it will select the first spot to start adding new update notes on
    ' the UPDATES worksheet.

    With wb.Sheets("UPDATES")

        .Activate
        .Unprotect
        .Columns(5).ColumnWidth = 10

        ' If Auto clear Update Notes is not enabled then it will not delete the release notes!
        If StrComp(
                rngX _
                    .Find _
                        ("Auto Clear Notes",
                        lookat:=xlPart) _
                    .Offset(0, 1) _
                    .Text _
                , "TRUE" _
                , vbTextCompare) _
            = 0 Then

            lRow = .Range("A" & 5).End(xlDown).Row
            .Range("A6:C" & lRow).ClearContents
            .Range("E2").Value = ""

        End If

        .Range("A6").Select

    End With


    Exit Sub

ErrorOccured1:
    Call errorNotice(funName:=funInName & " -> removeUpdateNotes -> location 1", errDesc:=Err.Description)

    Exit Sub

ErrorOccured2:
    Call errorNotice(funName:=funInName & " -> removeUpdateNotes -> location 2", errDesc:=Err.Description)

    Exit Sub

ErrorOccured3:
    Call errorNotice(funName:=funInName & " -> removeUpdateNotes -> location 3", errDesc:=Err.Description)

    Exit Sub

End Sub
Private Function firstWordOnly(str As String, Optional inFunName As String = "") As String

    On Error GoTo ErrorOccured

    'firstWordOnly = Left(str, InStr(str, " ") - 1)
    firstWordOnly = Split(str, " ")(0)

    Exit Function

ErrorOccured:

    Call errorNotice(funName:=inFunName & "firstWordOnly", errDesc:=Err.Description)

    Exit Function

End Function
Public Sub errorNotice(Optional errNum As Integer = 0, Optional funName As String = "", Optional errDesc As String = "")

    Dim bkDir As String

    bkDir = "\\SVFS02\cfmInfo\cfm Documents\Amos Quote Program\In Progress and Old Quote Data\Backup Copies"

    ' Email myself the error occurance

    Call Mail_Ryan_On_Errors(errNum, funName, errDesc)

    If Not errNum = 5 Then MsgBox "An error has occured in your release process.  Please restore your QuoteData " &
        "from the latest saved backup copy found here:" & vbNewLine & vbNewLine &
        """" & bkDir & """" & vbNewLine & vbNewLine & "If you contintue to experience " &
        "difficulties releasing, please contact your system administrator for assistance." &
        vbNewLine & vbNewLine & "We appologize for the inconvenience, please press OK to " &
        "return to where you were before attempting to release QuoteData."
    Application.EnableEvents = True
    If errNum = 6 Then End

End Sub
Private Sub Mail_Ryan_On_Errors(Optional errNum As Integer = 0, Optional funName As String = "", Optional errDesc As String = "")

    On Error GoTo ErrorOccured

    'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
    'Working in Office 2000-2013
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String, myMsg As String, br As String

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    br = " <br> "

    If Not StrComp(errDesc, "", vbTextCompare) = 0 Then _
        errDesc = br & br & "Error Description: " & errDesc & br & br

    If errNum = 0 Or errNum = 6 Then
        myMsg = "I experienced an error trying to release QD.  We may need assistance with this issue."
    Else
        Select Case errNum
            Case Is = 1
                myMsg = "The release script was unable to find a valid user to use as a default user.  " &
                "Is this still working?"
            Case Is = 2
                myMsg = "There seems to be an issue with the debug True/False cell."
            Case Is = 3
                myMsg = "There seems to be an issue with saving the default number format style."
            Case Is = 4
                myMsg = "There seems to be an issue with saving the default values on the updates tab."
            Case Is = 5
                myMsg = "There seems to be an issue with the end of releaseQD email process."
        End Select
    End If

    myMsg = myMsg & br & br & "Your help in investigating this matter is greatly appreciated." & errDesc

    strbody = "Ryan," & br & br & myMsg & br

    ' add the calling function to the end of the email if it is available
    If Not StrComp(funName, "", vbTextCompare) = 0 Then strbody = strbody & "<br><br> " &
        "This error was sent from the " & """" & funName & """" &
        " function in the workbook " & """" & ThisWorkbook.FullName & """"

    On Error Resume Next

    With OutMail
        .To = "r.valizan@cfmkc.com"
        .cC = ""
        .BCC = ""
        .Subject = "This just in from - QuoteData Error Reporting - "
        .HTMLBody = strbody
        .Send
    End With

    GoTo ExitSub

ErrorOccured:

    MsgBox "An error has occured when attempting to notify your QuoteData administrator of potential errors." &
        vbNewLine & vbNewLine & "Please inform Ryan Valizan of this via email at r.valizan@cfmkc.com, Thank You!"

GoTo ExitSub

ExitSub:
    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub
Public Sub setDefaultNumberFormat(Optional wb As Workbook, Optional fnCalledMe As String)

    Dim numFormat As String, rngX As Range, upDateSheet As Worksheet
    ' This is the default number format for the "normal" style
    ' This has proven it can change without anyone knowing.
    ' We will make sure this is always set correctly with each release AND whenever
    ' The workbook is saved or closed in the future!
    ' Also, we will allow this to be modifiably, but only by changing the text we choose
    ' as the answer will be pulled from a cell down by where we check for debug mode!
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Set upDateSheet = wb.Sheets("UPDATES")
    
    On Error GoTo ErrorOccured

    If Not upDateSheet Is Nothing Then
            Set rngX = upDateSheet.Range("A1:A10000").Find("Default Number Format", lookat:=xlPart)
            If Not rngX Is Nothing Then
            numFormat = rngX.Offset(0, 1).Text
            If Not wb.Styles("Normal").NumberFormat = numFormat Then
                wb.Styles("Normal").NumberFormat = numFormat
                Application.EnableEvents = False
                wb.Save
                Application.EnableEvents = True
            End If
            Exit Sub
        End If
    End If
    Set upDateSheet = Nothing
    
Exit Sub

ErrorOccured:
    Set upDateSheet = Nothing
    
    Call errorNotice(3, fnCalledMe & " -> setDefaultNumberFormat", errDesc:=Err.Description)

End Sub
Private Sub saveOpenCloseDelete()

    On Error GoTo killSubError

    ' Private   Public
    Dim tEnv As String, FName As String, brknWBPath As String,
        thisWB As Workbook, pre_Path As String, brknWB_PreBackup As String,
        tempSaveWB As Workbook, backUpWB As Workbook, recoverPath As String,
        rngX2 As Range, saveAsName As String, fullFName As String
        
    
    Set thisWB = ThisWorkbook
    
'    tEnv = Environ("Temp")
'    FName = env & "\temp_" & Format(DateTime.Now, "m-dd-yyyy")
    
    brknWBPath = thisWB.Path
    pre_Path = "\Quote Data Releases\Pre-Clean Styles Backup"
    tEnv = brknWBPath & pre_Path

    FName = tEnv & "\temp_" & Format(curDateTime, "m-dd-yyyy")
    brknWB_PreBackup = tEnv & "\QuoteData - single backup before style cleanup.xlsm"
    
    
    
    
     Set rngX2 = thisWB.Sheets("UPDATES").Range("A1:A10000").Find("Enable Debug Mode", lookat:=xlPart)
     
     saveAsName = brknWBPath & "\" & Left(thisWB.Name, InStrRev(thisWB.Name, ".") - 1)

    'If StrComp(rngX2.Offset(0, 1).Text, "True", vbTextCompare) = 0 Then _
    savePath = brknWBPath"\\SVFS02\cfmInfo\cfm Documents\Amos Quote Program\In Progress and Old Quote Data" & "\RECOVERED\Ryan's test folder"
    'If f Then
    
    Application.EnableEvents = False
    thisWB.SaveAs _
        Filename:=FName,
        FileFormat:=52,
        ReadOnlyRecommended:=False,
        CreateBackup:=False,
        ConflictResolution:=xlLocalSessionChanges
    
    Set backUpWB = Workbooks.Open(brknWB_PreBackup)
    
    backUpWB.SaveAs _
        Filename:=saveAsName,
        FileFormat:=52,
        ReadOnlyRecommended:=False,
        CreateBackup:=False,
        ConflictResolution:=xlLocalSessionChanges
Application.EnableEvents = True

    On Error GoTo killSubError2

    Application.EnableEvents = False
    With ThisWorkbook

        fullFName = .FullName
        .Save
        .ChangeFileAccess Mode:=xlReadOnly
        If fileExists(fullFName, "saveOpenCloseDelete") Then Kill fullFName
        .Close SaveChanges:=False
    End With
    Application.EnableEvents = True

    Exit Sub
killSubError:

    Set thisWB = Nothing
    Set rngX2 = Nothing
    Set backUpWB = Nothing
    
    Call errorNotice(6, "Error deleting current copy and reverting back to origional format at location 1", errDesc:=Err.Description)

    Exit Sub

killSubError2:

    Call errorNotice(6, "Error deleting current copy and reverting back to origional format at location 2", errDesc:=Err.Description)

    Exit Sub

End Sub
Private Function fileExists(FName As String, Optional fnCalledMe As String) As Boolean

    On Error GoTo ErrorOccured

    'Dim obj_fso As Object

    'Set obj_fso = CreateObject("Scripting.FileSystemObject")
    'fileExists = obj_fso.fileExists(FName)

    fileExists = (Dir(FName) > "")

    Exit Function

ErrorOccured:

    Call errorNotice(6, fnCalledMe & " -> fileExists Check", errDesc:=Err.Description)

    Exit Function

End Function
Private Function GetWorkbook(ByVal sFullName As String, Optional rdOnly As Boolean = False) As Workbook

    Dim sFile As String
    Dim wbReturn As Workbook
    Dim defDispAlert As Boolean

    'sFile = Dir(sFullName, vbNormal)
    sFile = sFullName
    defDispAlert = True

    If Application.DisplayAlerts = False Then defDispAlert = False

    On Error Resume Next
    'Set wbReturn = Workbooks(sFile)

    If wbReturn Is Nothing Then _
            Set wbReturn = Workbooks.Open(sFullName, ReadOnly:=rdOnly) 'Or (rdOnly = True And wbReturn.ReadOnly = False) Then
            'If Not wbReturn Is Nothing And rdOnly = True And wbReturn.ReadOnly = False Then'''
'
'                    wbReturn.ChangeFileAccess xlReadOnly
'                    Application.DisplayAlerts = False
'                    wbReturn.Save
'                    Application.DisplayAlerts = defDispAlert
'
'            Else
'
'                Set wbReturn = Workbooks.Open(sFullName, ReadOnly:=rdOnly)
'
'            End If
'        End If
        
    On Error GoTo 0

    Set GetWorkbook = wbReturn
    
    Set wbReturn = Nothing
    
End Function


