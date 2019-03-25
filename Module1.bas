Option Explicit
Option Base 1
Option Private Module

Dim workPhases As String
Dim x As Long
Dim cellText As String
Dim ws as Worksheet

Sub RunReport()

    Dim startTime As Double
    Dim elapsedTime As Double

    Dim dataArray As Variant
    
    startTime = Timer
    
    Call createLastColumns()
    dataArray = ThisWorkbook.Sheets("Harvest").Range("A1").CurrentRegion.Value
    
    'Call renameColumns(dataArray)
    Call fillLastColumns(dataArray)
    Call createSheets()
    Call customFilter("Harvest", "Ylläpito", Sheets("Makrot").Range("A1:A2"))
    Call customFilter("Harvest", "Tiketit", Sheets("Makrot").Range("B1:B2"))
    Call createPivot("Ylläpito")
    Call createPivot("Tiketit")

    elapsedTime = Round(Timer - startTime, 2)
    workPhases = workPhases & "Suodatus tehty " & elapsedTime & " sekunnissa"
    MsgBox workPhases, vbOKOnly + vbInformation, Title:="Suorituksen tiedot"

End Sub

Sub RestoreAll(restoreDemoData As Boolean)
    
    Dim ws As Worksheet
    Dim wsName As String
    
    Application.DisplayAlerts = False
    For Each ws In Worksheets
        wsName = ws.Name
        Select Case wsName
            Case "Harvest", "Ylläpito", "Tiketit", "YlläpitoPivot", "TiketitPivot"
                Sheets(wsName).Delete
        End Select
    Next ws
    Application.DisplayAlerts = True
    
    If restoreDemoData Then 
        Sheets("HarvestRestore").Visible = True
        Sheets("HarvestRestore").Copy Before:=Sheets("HarvestRestore")
        Sheets(Sheets("HarvestRestore").Index - 1).Name = "Harvest"
        Sheets("HarvestRestore").Visible = False
    Else
        ThisWorkbook.Sheets.Add.Name = "Harvest"
    End If
    
    ThisWorkbook.Sheets("Makrot").Activate
    DoEvents

End Sub

Private Sub renameColumns(ByRef dataArray As Variant)

    For x = 1 to Ubound(dataArray, 2)
        cellText = dataArray(1, x)
        Select Case cellText
            Case ""
        End Select

    Next x

End Sub

Private Sub createLastColumns()
    Dim lastCol As Integer
    Set ws = ThisWorkbook.Sheets("Harvest")
    ws.Activate
    DoEvents
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    'Purposefully overwrite last column 'URL', it is not used
    ws.Cells(1, lastCol).Value = "Ticket?"
    ws.Cells(1, lastCol + 1).Value = "TicketText"
    ws.Cells(1, lastCol + 2).Value = "TicketID"
End Sub

Private Sub fillLastColumns(ByRef dataArray As Variant)

    Dim notesCol As Integer
    Dim leftText As String
    Dim lastRow As Long
    Dim lastCol As Integer

    Dim splitText() As String
    Dim isTicket As Boolean

    notesCol = 6
    lastRow = Ubound(dataArray, 1)
    lastCol = Ubound(dataArray, 2)

    For x = 2 to lastRow
        cellText = dataArray(x, notesCol)
        isTicket = regexTest(cellText, "^(INC|EXT)[0-9]+:")
        'TODO midprio add lower function
        Select Case isTicket
            Case True
                splitText = Split(cellText, ":", 2)
                dataArray(x, lastCol - 2) = "Yes"
                dataArray(x, lastCol - 1) = cellText
                dataArray(x, lastCol) = splitText(0)
            Case Else
                dataArray(x, lastCol - 2) = "No"
        End Select
    Next x

    ThisWorkbook.Sheets("Harvest").Activate
    DoEvents
    ThisWorkbook.Sheets("Harvest").Range(Cells(1, 1), Cells(lastRow, lastCol)) = dataArray

End Sub

Private Sub createSheets()
    ThisWorkbook.Sheets.Add.Name = "Ylläpito"
    ThisWorkbook.Sheets.Add.Name = "Tiketit"
    ThisWorkbook.Sheets.Add.Name = "YlläpitoPivot"
    ThisWorkbook.Sheets.Add.Name = "TiketitPivot"
End Sub

Private Sub customFilter(sourceSheet As String, destinationSheet As String, filterRange As Range)

    Dim dataRange As Range

    Set dataRange = ThisWorkbook.Sheets(sourceSheet).Range("A1").CurrentRegion
    dataRange.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:= _
        filterRange, Unique:=False, CopyToRange:=Sheets(destinationSheet).Range("A1")

End Sub

Private Sub createPivot(sheetName As String)
    Dim pRange As Range
    Dim pCache As PivotCache
    Dim pTable As PivotTable
    Dim pName As String
    Dim destinationSheet As String
    Dim pos As Integer

    Set pRange = ThisWorkbook.Sheets(sheetName).Range("A1").CurrentRegion
    pName = sheetName & "PivotTable"
    destinationSheet = sheetName & "Pivot"
    Set PCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=PRange)
    Set PTable = PCache.CreatePivotTable(TableDestination:=Sheets(destinationSheet).Cells(2, 2), TableName:=pName)

    pos = 1
    With PTable.PivotFields("Task")
        .Orientation = xlRowField
        .Position = pos
    End With

    pos = pos + 1
    If sheetName = "Tiketit" Then
        With PTable.PivotFields("TicketID")
            .Orientation = xlRowField
            .Position = pos
        End With
    pos = pos + 1
    End If

    With PTable.PivotFields("Last Name")
        .Orientation = xlRowField
        .Position = pos
    End With

    With PTable.PivotFields("Hours")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlSum
        .NumberFormat = "# ##0.00"
        .Name = "Tunnit "
    End With
End Sub

Private Function regexTest(text As String, regexPattern) As Boolean

    Dim regExpr As New RegExp

    With regExpr
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = regexPattern
    End With

    regexTest = regExpr.Test(text)

End Function

Sub test()
    

    Dim regEx As New RegExp
    Dim regexResult As String
    Dim testString As String
    Dim strPattern As String

    strPattern = "(INC|EXT)[0-9]+:"

    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = strPattern
    End With

    testString = "INC123:eieui"
    Debug.Print regEx.test(testString)
    testString = "INC123:"
    Debug.Print regEx.test(testString)
    testString = "EXT123:"
    Debug.Print regEx.test(testString)
    testString = "INC123"
    Debug.Print regEx.test(testString)
    testString = "INC:"
    Debug.Print regEx.test(testString)
End Sub



