Option Explicit
Option Base 1
Option Private Module

Dim workPhases As String
Dim x As Long
Dim cellText As String

Sub RunReport()

    Dim startTime As Double
    Dim elapsedTime As Double

    Dim ws as Worksheet
    Dim dataArray As Variant
    Dim lastCol As Integer
    
    startTime = Timer
    
    Set ws = ThisWorkbook.Sheets("Harvest")
    ws.Activate
    DoEvents
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    'Purposefully overwrite last column
    ws.Cells(1, lastCol).Value = "Ticket?"
    ws.Cells(1, lastCol + 1).Value = "TicketID"

    dataArray = ThisWorkbook.Sheets("Harvest").Range("A1").CurrentRegion.Value
    
    'Call renameColumns(dataArray)
    Call fillLastColumns(dataArray, lastCol)

    
    elapsedTime = Round(Timer - startTime, 2)
    workPhases = workPhases & "Suodatus tehty " & elapsedTime & " sekunnissa"
    MsgBox workPhases, vbOKOnly + vbInformation, Title:="Suorituksen tiedot"

End Sub

Sub renameColumns(ByRef dataArray As Variant)

    For x = 1 to Ubound(dataArray, 2)
        cellText = dataArray(1, x)
        Select Case cellText
            Case
        End Select

    Next x

End Sub

Sub fillLastColumns(ByRef dataArray As Variant, lastCol As Integer)

    Dim notesCol As Integer
    Dim leftText As String

    notesCol = 6

    For x = 2 to Ubound(dataArray, 1)
        cellText = dataArray(x, notesCol)
        leftText = Left(cellText, 3)
        'TODO midprio add lower function
        Select Case leftText
            Case "INC", "EXT"
                dataArray(x, lastCol) = "Yes"
                dataArray(x, lastCol + 1) = cellText
            Case Else
                dataArray(x, lastCol) = "No"
        End Select
        If x = 5 Then Exit For

    Next x

End Sub

Sub test()

'Debug.Print Range("A1:B2").Address
Debug.Print Sheets("Filters").includeEU.Value

End Sub



