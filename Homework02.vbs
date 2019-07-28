VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockMarket()

    ' Set an initial variable for holding the stock name
    Dim Ticker As String

    ' Set a variable for holding the stock price at the beginning of the year
    Dim Ticker_Close_Beginning As Double
    Ticker_Close_Beginning = Cells(2, 6).Value
    
    ' Set a variable for holding the stock price at the end of the year
    Dim Ticker_Close_End As Double
  
    ' Set an initial variable for holding the total stock volume
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0

    ' Keep track of the location for each stock name in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Set an initial variable for largest percentage
    Dim Largest_Percentage As Double
    Largest_Percentage = 0
    
    ' Set an initial variable for smallest percentage
    Dim Smallest_Percentage As Double
    Smallest_Percentage = 0
    
    ' Set an initial variable for holding the largest total stock volume
    Dim Largest_Total_Stock_Volume As Double
    Largest_Total_Stock_Volume = 0
    
    ' Set an initial variable for holding the largest percentage stock name
    Dim Largest_Percentage_Ticker As String

    ' Set an initial variable for holding the smallest percentage stock name
    Dim Smallest_Percentage_Ticker As String

    ' Set an initial variable for holding the largest total stock volume stock name
    Dim Largest_Volume_Ticker As String


    ' Loop through all stocks
    For i = 2 To 999999

            ' Check if we are still within the same ticker, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
                ' Set the stock name
                Ticker = Cells(i, 1).Value
              
                ' Set the ticket end value
                Ticker_Close_End = Cells(i, 6).Value
        
                ' Add to the total stock volume
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        
                ' Print the stock name in the Summary Table
                Range("I" & Summary_Table_Row).Value = Ticker
                
                ' Print the yearly change in the Summary Table
                Range("J" & Summary_Table_Row).Value = Ticker_Close_End - Ticker_Close_Beginning
                Range("J" & Summary_Table_Row).Select
                Selection.NumberFormat = "0.000000000"
                
                ' Print the percent change in the Summary Table
                Range("K" & Summary_Table_Row).Value = (Ticker_Close_End - Ticker_Close_Beginning) / Ticker_Close_Beginning
                Range("K" & Summary_Table_Row).Select
                Selection.NumberFormat = "0.00%"
            
                ' Find the largest percentage
                If Range("K" & Summary_Table_Row).Value > Largest_Percentage Then
                
                    Largest_Percentage = Range("K" & Summary_Table_Row).Value
                    Largest_Percentage_Ticker = Cells(i, 1).Value
                    
                End If
                
                ' Find the smallest percentage
                If Range("K" & Summary_Table_Row).Value < Smallest_Percentage Then
                
                    Smallest_Percentage = Range("K" & Summary_Table_Row).Value
                    Smallest_Percentage_Ticker = Cells(i, 1).Value
                    
                End If
    
                ' Print the total stock volume to the Summary Table
                Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                
                ' Find the largest total stock volume
                If Range("L" & Summary_Table_Row).Value > Largest_Total_Stock_Volume Then
                
                    Largest_Total_Stock_Volume = Range("L" & Summary_Table_Row).Value
                    Largest_Volume_Ticker = Cells(i, 1).Value
                    
                End If
        
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset the variable for holding the stock price at the beginning of the year
                Ticker_Close_Beginning = Cells(i + 1, 6).Value
              
                ' Reset the total stock volume
                Total_Stock_Volume = 0
    
            ' If the cell immediately following a row is the same stock...
            Else
            
                ' Add to the total stock volume
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
    
            End If
    
      Next i
      
Cells(2, 15) = Largest_Percentage_Ticker
Cells(2, 16).Value = Largest_Percentage
Cells(2, 16).Select
Selection.NumberFormat = "0.00%"

Cells(3, 15) = Smallest_Percentage_Ticker
Cells(3, 16).Value = Smallest_Percentage
Cells(3, 16).Select
Selection.NumberFormat = "0.00%"

Cells(4, 15) = Largest_Volume_Ticker
Cells(4, 16).Value = Largest_Total_Stock_Volume

End Sub
