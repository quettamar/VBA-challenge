Attribute VB_Name = "Module1"
Sub stocks()

'declare and set worksheet
Dim ws As Worksheet
Set ws = ActiveSheet
'Set ws = Worksheets("Sheet1")
'Set ws = Worksheets(2)


'declare variables this seems to be right
Dim ticker As String
Dim year_open As Double
Dim year_close As Double
Dim high As Double
Dim low As Double
Dim percent_change As Double
Dim yearly_change As Double
Dim volume As Double
Dim total_stock_volume As Double

'create column headers. this seems to be the same across all answers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
        
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'using this this to set the open volume
year_open = Cells(2, 3).Value

'we can use these to move down rows
Summary_Table_Row = 2
Brand_Total = 0

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'doublecheck to make sure lastrow or rowcount is correct
For i = 2 To LastRow
'For j = 1 To 7 ask gretel if I need to worry about this


        'this will check to see if the value below it doesn't
        'match the cells above
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
            'This is adding up the totals
            total_stock_volume = total_stock_volume + Cells(i, 7).Value
            'This is setting the credit card name for the summary table
            ticker = Cells(i, 1).Value
        
            'this is where I will calculate the yearly change in open/close value
            
            year_close = Cells(i, 6).Value
            yearly_change = year_close - year_open
            
            'this is where I will calulate the percent change
             percent_change = (yearly_change / year_open) * 100
        
            
            'this is where I will nest another loop to change to green or red
                If (yearly_change <= 0) Then
            
                    Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                
                Else
                
                    Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                
                End If
        
        'This is putting the totals and brand names in the summary table
        Cells(Summary_Table_Row, 9).Value = ticker
        Cells(Summary_Table_Row, 10).Value = yearly_change
        Cells(Summary_Table_Row, 11).Value = percent_change
        Cells(Summary_Table_Row, 12).Value = total_stock_volume
        
               
        'this will move down to the next row of the summary table
        Summary_Table_Row = Summary_Table_Row + 1
        
        'this will reset the totals and move onto the next group
        total_stock_volume = 0
        percent_change = 0
        yearly_change = 0
               
    Else
    
        'this is moving on down the line
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
        End If

    'Next j

Next i


End Sub
