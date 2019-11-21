Attribute VB_Name = "Module2"
Sub DSV_Loop()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call TotalStockVolume
    Next
    Application.ScreenUpdating = True
End Sub


Sub TotalStockVolume()


    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    'Declare and Initialize Current Stock
    Dim CurrentStock As String
    CurrentStock = Cells(2, 1)
    
    'Declare Previous Stock
    Dim PrevStock As String
    PrevStock = Cells(2, 1)
    
    'Declare Sum of Stock Entries, initialize as first stock value
    Dim StockTotal As Double
    StockTotal = 0
    
    'Find length of Excel data
    Dim lastrow As Long
    Set sht = ActiveSheet
    lastrow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row
    
    Dim StockCount As Integer
    StockCount = 1
    
    'Declare initial stock price
    Dim InitialStockPrice As Double
    InitialStockPrice = Cells(2, 3).Value
    
    'Declare Year End Stock Price
    Dim YearEndStockPrice As Double
    YearEndStockPrice = 0
    
    'Declare Percent Change Stock Price
    Dim PercentChange As Double
    PercentChange = 0
    
    'Declare Numerical Change Stock Price
    Dim NumChange As Double
    NumChange = 0
    
    'Declare summary values
    Dim GreatestIncreaseTicker As String
    Dim GreatestIncreaseValue As Double
    Dim GreatestDecreaseTicker As String
    Dim GreatestDecreaseValue As Double
    Dim GreatestTotalTicker As String
    Dim GreatestTotalValue As Double
    GreatestIncreaseValue = 0
    GreatestDecreaseValue = 0
    GreatestTotalValue = 0
    
    'Loop through each stock entry
    For I = 2 To lastrow
    
    'Change current stock to the next stock value
    CurrentStock = Cells(I, 1)
        
        'If current stock is same as previous stock, then add to stock volume
        If CurrentStock = PrevStock Then
            StockTotal = StockTotal + Cells(I, 7)
        End If
        
        'If current stock is different from previous stock, store previous stock volume, reset stock value, set previous stock = current stock
        If CurrentStock <> PrevStock Then
            
            'Store stock Ticker
            Cells(StockCount + 1, 9) = PrevStock
            
            'Store Stock total volume
            Cells(StockCount + 1, 12) = StockTotal
            
            'Assign end of year stock price
            YearEndStockPrice = Cells(I - 1, 6)
            
            'Calculate Stock Price Change, must round to two digits, store these values
            If InitialStockPrice > 0 Then
                PercentChange = Round((YearEndStockPrice / InitialStockPrice - 1) * 100, 2)
            Else
                PercentChange = 0
            End If
            
            NumChange = YearEndStockPrice - InitialStockPrice
            Cells(StockCount + 1, 10) = NumChange
            Cells(StockCount + 1, 11) = PercentChange & "%"
            
            'If NumChange is negative, color red, else color green
            If NumChange < 0 Then
                Cells(StockCount + 1, 10).Interior.Color = RGB(255, 0, 0)
                
                'Check to see if the PercentChange is smallest value, assign it to the smallest value if so
                If PercentChange < GreatestDecreaseValue Then
                GreatestDecreaseValue = PercentChange
                GreatestDecreaseTicker = PrevStock
                End If
                
            Else
                Cells(StockCount + 1, 10).Interior.Color = RGB(0, 255, 0)
                
                'Check to see if the PercentChange is biggest value, assign GreatestIncreaseValue to PercentChange if so
                If PercentChange > GreatestIncreaseValue Then
                GreatestIncreaseValue = PercentChange
                GreatestIncreaseTicker = PrevStock
                End If
            End If
            
            If StockTotal > GreatestTotalValue Then
                GreatestTotalValue = StockTotal
                GreatestTotalTicker = PrevStock
            End If
            
            'Reset Stock Values, increment stock count
            StockTotal = 0
            PrevStock = CurrentStock
            InitialStockPrice = Cells(I, 3)
            StockCount = StockCount + 1
        End If
        
    Next I
    
    'Set summary values
    Cells(2, 16) = GreatestIncreaseTicker
    Cells(2, 17) = GreatestIncreaseValue & "%"
    Cells(3, 16) = GreatestDecreaseTicker
    Cells(3, 17) = GreatestDecreaseValue & "%"
    Cells(4, 16) = GreatestTotalTicker
    Cells(4, 17) = GreatestTotalValue

End Sub

