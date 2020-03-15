Sub stocks()

Dim TickerSymbol As String
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim DateOpen As Integer
Dim DateClose As Integer
Dim LastRow As Long
LastRow = Range("A" & Rows.Count).End(xlUp).Row
Dim TotalVolume As Double
Dim Diff As Double
Dim ConditionValue As Double
Dim ws As Worksheet

'loops through all worksheets in the workbook and applies the ws prefix to all cell locations
For Each ws In ActiveWorkbook.Worksheets
ws.Activate

  

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  TotalVolume = 0
  bBeginingTicker = True
  
'Add Additional headers for calculations
Cells(1, 9).Value = ("Ticker")
Cells(1, 10).Value = ("Yearly Change")
Cells(1, 11).Value = ("Percent Change")
Cells(1, 12).Value = ("Total Stock Volume")
Cells(1, 16).Value = ("Ticker")
Cells(1, 17).Value = ("Value")
  
  ' Loop through all tickers
  For i = 2 To LastRow
  
    'Get the stock open price for first counter
        If bBeginingTicker = True Then
            OpenPrice = Cells(i, 3).Value
        End If
    
    'Print Percentage Change into Column K
    If OpenPrice And ClosePrice > 0 Then
        Range("K" & Summary_Table_Row).Value = (ClosePrice - OpenPrice) / (OpenPrice)
    End If
        
                
    ' Check if we are still within the same Stock ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       
         ' Set the Stock name
         TickerSymbol = Cells(i, 1).Value
        
        ' Print the TickerSymbol in the Summary Table
          Range("I" & Summary_Table_Row).Value = TickerSymbol
       
       ' Print the OpenPrice to the Summary Table Checking OpenPrice
         'Range("M" & Summary_Table_Row).Value = OpenPrice
         
        ' Print Difference price into Column L
         Range("J" & Summary_Table_Row).Value = ClosePrice - OpenPrice
       
       ' Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
          
          'Reset Volume to 0
           TotalVolume = 0
           
        'Reset biginning stock flag
            bBeginingTicker = True
    
    ' If the cell immediately following a row is the same Symbol...
    Else
        
            ClosePrice = Cells(i + 1, 6).Value
                 
           'Add to the ClosePrice Column K checking results
           'Range("N" & Summary_Table_Row).Value = ClosePrice
            
            TotalVolume = TotalVolume + Cells(i, 7).Value
                Range("L" & Summary_Table_Row).Value = TotalVolume
              
        ' Coler the cells with Green for Positive and Red For Negative
        If Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        ElseIf Cells(i, 10).Value < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        End If
             
            'Reset biginning stock flag
            bBeginingTicker = False
      
        End If
   
   
   
  Next i
 
 
Next ws


End Sub










