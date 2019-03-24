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
    
    elapsedTime = Round(Timer - startTime, 2)
    workPhases = workPhases & "Suodatus tehty " & elapsedTime & " sekunnissa"
    MsgBox workPhases, vbOKOnly + vbInformation, Title:="Suorituksen tiedot"

End Sub

Sub RestoreAll()
    
    Dim ws As Worksheet
    Dim wsName As String
    
    Sheets("HarvestRestore").Visible = True

    Application.DisplayAlerts = False
    For Each ws In Worksheets
        wsName = ws.Name
        Select Case wsName
            Case "Harvest", "Ylläpito", "Tiketit"
                Sheets(wsName).Delete
        End Select
    Next ws
    Application.DisplayAlerts = True
    
    Sheets("HarvestRestore").Copy Before:=Sheets("HarvestRestore")
    Sheets(Sheets("HarvestRestore").Index - 1).Name = "Harvest"
    DoEvents
    
    Sheets("HarvestRestore").Visible = False
    
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
    ws.Cells(1, lastCol + 1).Value = "TicketID"
End Sub

Private Sub fillLastColumns(ByRef dataArray As Variant)

    Dim notesCol As Integer
    Dim leftText As String
    Dim lastRow As Long
    Dim lastCol As Integer

    notesCol = 6

    lastRow = Ubound(dataArray, 1)
    lastCol = Ubound(dataArray, 2)

    For x = 2 to lastRow
        cellText = dataArray(x, notesCol)
        leftText = Left(cellText, 3)
        'TODO midprio add lower function
        Select Case leftText
            Case "INC", "EXT"
                dataArray(x, lastCol - 1) = "Yes"
                dataArray(x, lastCol) = cellText
            Case Else
                dataArray(x, lastCol - 1) = "No"
        End Select
    Next x

    ThisWorkbook.Sheets("Harvest").Activate
    DoEvents
    ThisWorkbook.Sheets("Harvest").Range(Cells(1, 1), Cells(lastRow, lastCol)) = dataArray

End Sub

Private Sub createSheets()
    ThisWorkbook.Sheets.Add.Name = "Ylläpito"
    ThisWorkbook.Sheets.Add.Name = "Tiketit"
End Sub

Private Sub customFilter(sourceSheet As String, destinationSheet As String, filterRange As Range)

    Dim dataRange As Range

    Set dataRange = ThisWorkbook.Sheets(sourceSheet).Range("A1").CurrentRegion
    dataRange.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:= _
        filterRange, Unique:=False, CopyToRange:=Sheets(destinationSheet).Range("A1")

End Sub

Sub test()
    'Debug.Print Range("A1:B2").Address
    Debug.Print Sheets("Filters").includeEU.Value
End Sub



