Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If Not ThisWorkbook.ReadOnly Then setDefaultNumberFormat ActiveWorkbook, "Workbook_BeforeClose"
    
    Set curWB = Nothing

End Sub
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If Not ThisWorkbook.ReadOnly Then
        If Not inRelease Then
            ckFormAndSave True, curWB
            setDefaultNumberFormat ActiveWorkbook, "Workbook_BeforeClose"
        End If
    End If
End Sub
Private Sub Workbook_Open()
    Dim wbOpenDir As String, wbOpenName As String
    Set curWB = ThisWorkbook: pubWBjustOpened = True
    curDateTime = DateTime.Now

    ckFormAndSave True, curWB, True
    wbOpenDir = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "") & "Quote Data Releases" &
            "\Temporary Save Directory"
    wbOpenName = wbOpenDir & "\QuoteData - In Progress Read Only Temporary Save"
    If Not ThisWorkbook.ReadOnly Then
        If Not Len(Dir(wbOpenName & ".xlsm")) = 0 Then askToDelteTempSaves wbOpenDir
        setDefaultNumberFormat ActiveWorkbook, "Workbook_Open"
    Else
        askToSaveAsTemp wbOpenName
    End If
End Sub
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    If Success Then ckFormAndSave False, curWB
End Sub
