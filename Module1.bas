Attribute VB_Name = "Module1"

Sub AddColumnHeaders()
    Dim ws As Worksheet
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        
       
        
        ' Add column headers
        With ws
            .Cells(1, 9).Value = "Ticker"
            .Cells(1, 10).Value = "Quarterly Change"
            .Cells(1, 11).Value = "Percent Change"
            .Cells(1, 12).Value = "Total Stock Volume"
            
            .Cells(1, 16).Value = "Ticker"
            .Cells(1, 17).Value = "Value"
            .Cells(2, 15).Value = "Greatest % Increase"
            .Cells(3, 15).Value = "Greatest % Decrease"
            .Cells(4, 15).Value = "Greatest Total Volume"


        End With
    Next ws
    
End Sub

Sub FormatDateAllWorksheets()
    Dim ws As Worksheet
    Dim d As Long
    Dim lastRow As Long
        
' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Find last row in column B for the current worksheet
        lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        
' Loop through each cell in column B
        For d = 2 To lastRow
            ' Check if the cell contains a value
            If Not IsEmpty(ws.Cells(d, 2)) Then
                ' Check if the value is a string (assuming original format is YYYYMMDD as string)
                If IsNumeric(ws.Cells(d, 2).Value) And Len(ws.Cells(d, 2).Value) = 8 Then
                    ' Convert YYYYMMDD string to date
                    ws.Cells(d, 2).Value = DateSerial(Left(ws.Cells(d, 2).Value, 4), Mid(ws.Cells(d, 2).Value, 5, 2), Right(ws.Cells(d, 2).Value, 2))
                End If
                
                ' Format as date
                ws.Cells(d, 2).NumberFormat = "mm/dd/yyyy"
            End If
        Next d
    Next ws
    
    
    MsgBox "Date formatting complete for all worksheets!", vbInformation
End Sub

Sub tickerAndVolume()
'Start worksheet loop
Dim ws As Worksheet
Dim lastRow As Long

For Each ws In ThisWorkbook.Worksheets
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Set an initial variable for holding the Ticker Name
    Dim Ticker As String

    ' Set an initial variable for holding the total volume per Ticker Name
    Dim Ticker_Total As Double
    Ticker_Total = 0

    ' Keep track of the location for each Ticker Name in the Ticker Column
    Dim Ticker_Name_Row As Integer
    Ticker_Name_Row = 2

    ' Loop through all Ticker Names
    For i = 2 To lastRow

        ' Check if we are still within the same Ticker Name, if not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Set the Ticker name
            Ticker = ws.Cells(i, 1).Value

            ' Add to the Ticker Total
            Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

            ' Print the Ticker Symbol in the Ticker Column
            ws.Range("I" & Ticker_Name_Row).Value = Ticker

            ' Print the Ticker total to the Ticker Column
            ws.Range("L" & Ticker_Name_Row).Value = Ticker_Total

            ' Add one to the Ticker name row
            Ticker_Name_Row = Ticker_Name_Row + 1

            ' Reset the Ticker Total
            Ticker_Total = 0

        ' If the next cell immediately following a row is the same ticker...
        Else

            ' Add to the Ticker Total
            Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

        End If

    Next i
Next ws
MsgBox "Ticker/volume added to all worksheets!", vbInformation
End Sub

Sub format_percent()
    ' Start worksheet loop
    Dim ws As Worksheet
    Dim lastRow As Long
    
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Format columns for percentages
        For i = 2 To lastRow
           
            
                ' Set number format to display percentages with 2 decimal places
                ws.Cells(i, 11).NumberFormat = "0.00%"
            
        Next i
    Next ws
End Sub

Sub FindMaxVolumeInColumn()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxVol As Double
    Dim resultMaxVol As Range
    Dim resultCorrespondingA As Range

    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row with data in the column
        lastRow = ws.Cells(ws.Rows.Count, 12).End(xlUp).Row

        ' Find the maximum value in column L (12)
        maxVol = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, 12), ws.Cells(lastRow, 12)))

        ' Print maxVol to cell Q4 in the current worksheet
        Set resultMaxVol = ws.Range("Q4")
        resultMaxVol.Value = maxVol

        ' Find the corresponding value from column A (1)
        Dim correspondingA As Variant
        correspondingA = ws.Cells(Application.Match(maxVol, ws.Columns(12), 0), 1)

        ' Print the corresponding value from column A to cell P4 in the current worksheet
        Set resultCorrespondingA = ws.Range("P4")
        resultCorrespondingA.Value = correspondingA

        ' Display the maximum value and the corresponding value
        MsgBox "The maximum value in Column " & Split(ws.Cells(2, 12).Address, "$")(1) & " is: " & maxVol & vbCrLf & _
               "The corresponding value in Column A is: " & correspondingA
    Next ws
End Sub
Sub colorFormatting()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, 12).End(xlUp).Row

        For i = 2 To lastRow
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 10
            ElseIf ws.Cells(i, 10).Value = 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 0
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
    Next ws

End Sub
Sub quarterlyChange()


'quarterlyChange = openingPrice - closingPrice

End Sub
Sub percentChange()

'percentChange = ((closingPrice - openingPrice) / openingPrice) * 100

End Sub

