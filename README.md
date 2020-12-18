# VBA-challenge

Sub Stocks()

'Loop from Worksheet A TO P
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
    
 'Determine the Last Row
 LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

'Set values to the cells
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly Change"
    Cells(1, "K").Value = " Percent Change"
    Cells(1, "L").Value = "Total Stock Volume"
    
'Declare Variables
    Dim Open_Price As Double
    Dim Closing_Price As Double
    Dim Yearly_Change As Double
    Dim Ticker_Name As String
    Dim Percent_Change As Double
    Dim Volume As Double
'Declare volume as a number
    Volume = 0

'Declare the row and colums
    Dim Row As Double
    Row = 2
    Dim Column As Integer
    Column = 1
    Dim i As Long
    
'Set Open Price
    Open_Price = Cells(2, Column + 2).Value
        
'Loop through all stocks

    For i = 2 To LastRow

    'Output ticker symbol with "Conditional" if
    If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
    
        'Output ticker symbol
        Ticker_Name = Cells(i, Column).Value
        
        'Print Values in cells
        Cells(Row, Column + 8).Value = Ticker_Name
        
        'Set Close Price
    Close_Price = Cells(i, Column + 5).Value
    
        'Subtract Open_Price from Close_Price to get Yearly Change
        Yearly_Change = Close_Price - Open_Price
        'Print Values in cells
        Cells(Row, Column + 9).Value = Yearly_Change
        
        'Calculate Percent Change
        If (Open_Price = 0 And Close_Price = 0) Then
            Percent_Change = 0
        ElseIf (Open_Price = 0 And Close_Price <> 0) Then
            Percent_Change = 1
        Else
            Percent_Change = Yearly_Change / Open_Price
            Cells(Row, Column + 10).Value = Percent_Change
            Cells(Row, Column + 10).NumberFormat = "0.00%"
        End If
        
     'Set Total stock volume
     Volume = Volume + Cells(i, Column + 6).Value
     
     'Print Value In cells
     Cells(Row, Column + 11).Value = Volume
     
     'Add one to the summary table row
     Row = Row + 1
     
     'Reset Open_Price
    OpenPrice = Cells(i + 1, Column + 2)
    
     'Reset Volume Total
     Volume = 0
    
    'If cells are the same
    Else
        Volume = Volume + Cells(i, Column + 6).Value
    End If
 Next i
 
 '------------------------------------------------------------------------------------
    
    ' Last Row of Yearly_Change per WS
    YCLastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Highlight Postive change in green and negative change in red
    For j = 2 To YCLastRow
        If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
        'Set Color Green
            Cells(j, Column + 9).Interior.ColorIndex = 10
        ElseIf Cells(j, Column + 9).Value < 0 Then
        'Set Color Red
            Cells(j, Column + 9).Interior.ColorIndex = 3
        End If
    Next j
    
   Next WS
    
End Sub


