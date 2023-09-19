Attribute VB_Name = "Module1"

Sub StockAnalysis()

'Set Variables
Dim i As Long
Dim ticker As String
Dim open_value As Double
Dim high_value As Double
Dim low_value As Double
Dim close_value As Double
Dim volume As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_value As Double
Dim output_row As Integer
Dim ws As Worksheet


For Each ws In Worksheets

''Create Column Headings and Format
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Total Stock Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Value"
ws.Columns.AutoFit
ws.Range("I1:L1").Interior.ColorIndex = 6
ws.Range("P1:Q1").Interior.ColorIndex = 6


'Set EndPoint for Loop
Dim lastrow As Long
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
output_row = 2
open_value = ws.Cells(2, 3).Value

'Create Loop to cycle through values
For i = 2 To lastrow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    volume = volume + ws.Cells(i, 7).Value
    close_value = ws.Cells(i, 6).Value
    yearly_change = close_value - open_value
    ticker = ws.Cells(i, 1).Value
    
    If open_value <> 0 Then
        percent_change = yearly_change / open_value
    
    Else
        percent_change = 0
     
     End If
     
'     Output Results to corresponding cells
    ws.Cells(output_row, 9).Value = ticker
    ws.Cells(output_row, 10).Value = yearly_change
    ws.Cells(output_row, 11).Value = percent_change
    ws.Cells(output_row, 12).Value = volume


''Color Fill Yearly Change for Positive=Green, Negative = Red

Select Case yearly_change
    Case Is > 0
            ws.Range("J" & output_row).Interior.ColorIndex = 4
    Case Is < 0
            ws.Range("J" & output_row).Interior.ColorIndex = 3
    Case Else
            ws.Range("J" & output_row).Interior.ColorIndex = 0

    End Select

                    
                    
'           Clear Variables for the next ticker
                    
                    volume = 0
                    open_value = ws.Cells(i, 3).Value
                    output_row = output_row + 1
                    volume = volume + ws.Cells(i, 7).Value

                End If
        Next i
    




' Find Maximum And Minimum Percent Change
'
        
       ws.Range("Q2") = "%" & WorksheetFunction.max(ws.Range("K2:K" & lastrow)) * 100
       ws.Range("Q3") = "%" & WorksheetFunction.min(ws.Range("K2:K" & lastrow)) * 100
       ws.Range("Q4") = WorksheetFunction.max(ws.Range("L2:L" & lastrow))
        
       
        
        
        decrease_number = WorksheetFunction.Match(WorksheetFunction.min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
        increase_number = WorksheetFunction.Match(WorksheetFunction.max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)

         ' final ticker symbol for  total, greatest % of increase and decrease, and average
            ws.Range("P2") = ws.Cells(increase_number + 1, 9)
            ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
            ws.Range("P4") = ws.Cells(volume_number + 1, 9)
        


           
       
    Next ws
End Sub
