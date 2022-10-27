Attribute VB_Name = "Module1"
Sub Yearly_Stock()


'Define Variables
Dim ws As Worksheet

Dim Ticker As String

Dim Ticker_Total As LongLong

Dim Ticker_Row As LongLong

Dim Lastrow As LongLong

Dim i As LongLong

Dim Open_Price As Double

Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double

Dim cell As Range

'Loop Through All Stocks for the Year
For Each ws In Worksheets


'Column Headings
Ticker_Total = 0
Ticker_Row = 2
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
Open_Price = ws.Cells(2, 3).Value



Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop to Lastrow
For i = 2 To Lastrow

 
'Ticker Symbol Output
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    ws.Activate
    
        Ticker = ws.Cells(i, 1).Value
        
        Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
        
        Close_Price = ws.Cells(i, 6).Value
        
        Yearly_Change = Close_Price - Open_Price
        
        Range("I" & Ticker_Row).Value = Ticker
        
        Range("L" & Ticker_Row).Value = Ticker_Total
        
        Range("J" & Ticker_Row) = Yearly_Change
        
        
          If (Open_Price = 0 And Close_Price = 0) Then

                    Percent_Change = 0
                    
            ElseIf (Open_Price = 0 And Close_Price <> 0) Then
            
                    Percent_Change = 1
                    
            
          Else
                    
            Percent_Change = Yearly_Change / Open_Price
            
            Range("K" & Ticker_Row).Value = Percent_Change
              
            Range("K" & Ticker_Row).NumberFormat = "0.00%"
            
                
          End If
          

        Ticker_Row = Ticker_Row + 1
        
        Open_Price = ws.Cells(i + 1, 3)
        
        Ticker_Total = 0
        
        
     Else
    
        Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
        
        
    End If
    
    
    If Cells(i, 10).Value > 0 Then
    
    
                Cells(i, 10).Interior.ColorIndex = 10
        Else
        
                Cells(i, 10).Interior.ColorIndex = 3
    
    End If

    

Next i


Next ws


End Sub

