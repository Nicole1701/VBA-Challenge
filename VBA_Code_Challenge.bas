Attribute VB_Name = "Module1"
Sub WallStreet()

'Loop through all sheets
For Each ws In Worksheets


'Set Variables for Summary Table
    Dim SummaryTableRow As Integer
    SummaryTableRow = 2

'Summary Table Formating
    'Add Header to Table
    ws.Range("I1").Value = "Ticker"
    ws.Range("I1").Font.Bold = True
    ws.Range("I1").HorizontalAlignment = xlCenter
        
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("J1").Font.Bold = True
    ws.Range("J1").HorizontalAlignment = xlCenter
        
    ws.Range("K1").Value = "Percent Change"
    ws.Range("K1").Font.Bold = True
    ws.Range("K1").HorizontalAlignment = xlCenter
        
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("L1").Font.Bold = True
    ws.Range("L1").HorizontalAlignment = xlCenter


'Identify first and last prices for each ticker symbol

'Set ticker range and Calculate

    'Set Initial Variables
    Dim LastRow As Long
    Dim Ticker As String
    Dim TickerOpen As Long
    Dim TickerClose As Long
    Dim LastRowTicker As Long
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearChange As Double
    Dim PercentChange As Double
    Dim StockVolume As Double
        
   LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   LastRowTicker = ws.Cells(Rows.Count, 9).End(xlUp).Row

    'Set first Open Price and Volummn
    OpenPrice = ws.Cells(2, 3).Value
    StockVolume = 0


For i = 2 To LastRow

        'Check to see if ticker symbols are the same and add Volume
            If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
                StockVolume = ws.Cells(i, 7) + StockVolume
    
        'Check if same ticker symbol
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                              
               'Set Ticker Value and add to table
                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & SummaryTableRow).Value = Ticker
                
                'Set Close Price
                ClosePrice = ws.Cells(i, 6).Value
                
                'Find Change in Price and add to table
                YearChange = ClosePrice - OpenPrice
                'MsgBox (YearChange)
                ws.Range("J" & SummaryTableRow).Value = YearChange
                
                'Find Percent of Change and add to Table
                'stackoverflow.com/questions/18598872/vba-script-when-dividing-by-zero
                    If OpenPrice <> 0 Then
                    PercentChange = (YearChange / OpenPrice)
                    'MsgBox (PercentChange)
                    ws.Range("K" & SummaryTableRow).Value = PercentChange
                    
                    Else
                    Formula = 0
                    
                    End If
                
                'Find Total Stock Volume and put value in Table
                StockVolume = ws.Cells(i, 7) + StockVolume
                ws.Range("L" & SummaryTableRow).Value = StockVolume
               
                'Set new open price and Stock Volume
                OpenPrice = ws.Cells(i + 1, 3).Value
                StockVolume = 0
                
                'Add one to summary Table Row
                 SummaryTableRow = SummaryTableRow + 1
         
    
        End If
    
    Next i

'Format Table
    'Number Formatting
    ws.Range("J2:J" & LastRow).NumberFormat = "0.00"
    ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"
    
    For i = 2 To LastRowTicker
    
        If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.Color = 5287936
        ws.Cells(i, 10).Font.ColorIndex = 1
        
        ElseIf ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
        ws.Cells(i, 10).Font.ColorIndex = 2
        
    End If

    Next i
        
 
       
   'Challenge

    'Make Challenge Table
    'Add Header to Table
    ws.Range("O1").Value = "Ticker"
    ws.Range("O1").Font.Bold = True
    ws.Range("O1").HorizontalAlignment = xlCenter
    
    ws.Range("P1").Value = "Value"
    ws.Range("P1").Font.Bold = True
    ws.Range("P1").HorizontalAlignment = xlCenter
    
    'Add Rows to the Table
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N2").Font.Bold = True
    ws.Range("N2").HorizontalAlignment = xlLeft
        
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N3").Font.Bold = True
    ws.Range("N3").HorizontalAlignment = xlLeft
        
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("N4").Font.Bold = True
    ws.Range("N4").HorizontalAlignment = xlLeft
    
    'Number Formatting
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"
    ws.Range("P4").NumberFormat = "0"
    
   'Define Variables
    Dim GreatestIncrease As Double
    Dim SmallestIncrease As Double
    Dim LargestVolume As Double
    Dim Max As Double
    Dim Min As Double
    Dim Largest As Double
    Dim GreatestTicker As String
    Dim SmallestTicker As String
    Dim LargestTicker As String
  
    
    'Find Greatest % Increase and add to Challege Table
     
    Max = Application.WorksheetFunction.Max(ws.Range("K:K"))
     
        For i = 2 To LastRowTicker
            If ws.Cells(i, 11) = Max Then
               GreatestIncrease = ws.Cells(i, 11)
               GreatestTicker = ws.Cells(i, 9)
        End If
    
    Next i
    
        ws.Range("P2") = GreatestIncrease
        ws.Range("O2") = GreatestTicker
        'MsgBox (GreatestIncrease)
        'MsgBox (GreatestTicker)
       

    'Find Smallest % Increase and add to Challege Table
    
    Min = Application.WorksheetFunction.Min(ws.Range("K:K"))
    
        For i = 2 To LastRowTicker
            If ws.Cells(i, 11) = Min Then
               SmallestIncrease = ws.Cells(i, 11)
               SmallestTicker = ws.Cells(i, 9)
        End If
    
    Next i
    
        ws.Range("P3") = SmallestIncrease
        ws.Range("O3") = SmallestTicker
        'MsgBox (SmallestIncrease)
        'MsgBox (SmallestTicker)
        
    'Find Largest Volumne and add to Challege Table
        
        Largest = Application.Max(ws.Range("L:L"))
        
        For i = 2 To LastRowTicker
            If ws.Cells(i, 12) = Largest Then
               LargestVolume = ws.Cells(i, 12)
               LargestTicker = ws.Cells(i, 9)
        End If
    
    Next i
    
        ws.Range("P4") = LargestVolume
        ws.Range("O4") = LargestTicker
        'MsgBox (LargestVolume)
        'MsgBox (LargestTicker)
        
    ' Autofit to display data
    ws.Columns("A:P").AutoFit
    
Next ws

End Sub
