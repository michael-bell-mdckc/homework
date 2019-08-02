Attribute VB_Name = "Module1"
Sub WorksheetLoop()

Dim ws As Worksheet

For Each ws In Worksheets

    ' Set a variable for holding the stock name
    Dim Ticker As String
    
    ' Determine the Last Row
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Set a variable for holding the stock price at the beginning of the year
    Dim Ticker_Close_Beginning As Double
    Ticker_Close_Beginning = ws.Cells(2, 6).Value
    
    ' Set a variable for holding the stock price at the end of the year
    Dim Ticker_Close_End As Double
  
    ' Set an initial variable for holding the total stock volume
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0

    ' Set a variable to keep track of the location for each stock name in the summary table
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
    
    ' Set a variable for holding the largest percentage stock name
    Dim Largest_Percentage_Ticker As String

    ' Set a variable for holding the smallest percentage stock name
    Dim Smallest_Percentage_Ticker As String

    ' Set a variable for holding the largest total stock volume stock name
    Dim Largest_Volume_Ticker As String

    ' Loop through all stocks
    For i = 2 To LastRow

            ' Check if we are still within the same ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                ' Set the stock name
                Ticker = ws.Cells(i, 1).Value
              
                ' Set the ticket end value
                Ticker_Close_End = ws.Cells(i, 6).Value
        
                ' Add to the total stock volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
                ' Print the stock name in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                
                ' Print the yearly change in the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = Ticker_Close_End - Ticker_Close_Beginning
                
                    If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                    
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                        
                    Else
                    
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                        
                    End If
                    
                ' Print the percent change in the Summary Table
                If Ticker_Close_Beginning = 0 Then
                
                    ws.Range("K" & Summary_Table_Row).Value = 0
                    
                Else
                
                    ws.Range("K" & Summary_Table_Row).Value = (Ticker_Close_End - Ticker_Close_Beginning) / Ticker_Close_Beginning
                    
                End If
                
                ' Find the largest percentage
                If ws.Range("K" & Summary_Table_Row).Value > Largest_Percentage Then
                
                    Largest_Percentage = ws.Range("K" & Summary_Table_Row).Value
                    Largest_Percentage_Ticker = ws.Cells(i, 1).Value
                    
                End If
                
                ' Find the smallest percentage
                If ws.Range("K" & Summary_Table_Row).Value < Smallest_Percentage Then
                
                    Smallest_Percentage = ws.Range("K" & Summary_Table_Row).Value
                    Smallest_Percentage_Ticker = ws.Cells(i, 1).Value
                    
                End If
    
                ' Print the total stock volume to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                
                ' Find the largest total stock volume
                If ws.Range("L" & Summary_Table_Row).Value > Largest_Total_Stock_Volume Then
                
                    Largest_Total_Stock_Volume = ws.Range("L" & Summary_Table_Row).Value
                    Largest_Volume_Ticker = ws.Cells(i, 1).Value
                    
                End If
        
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset the variable for holding the stock price at the beginning of the year
                Ticker_Close_Beginning = ws.Cells(i + 1, 6).Value
              
                ' Reset the total stock volume
                Total_Stock_Volume = 0
    
            ' If the cell immediately following a row is the same stock...
            Else
            
                ' Add to the total stock volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    
            End If
    
      Next i
      
ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Stock Volume"
ws.Cells(1, 15) = "Ticker"
ws.Cells(1, 16) = "Value"
ws.Cells(2, 14) = "Greatest % Increase"
ws.Cells(3, 14) = "Greatest % Decrease"
ws.Cells(4, 14) = "Greatest Total Volume"
      
ws.Range("J:J").NumberFormat = "#.##0000000"
ws.Range("K:K").NumberFormat = "0.00%"
                
ws.Cells(2, 15) = Largest_Percentage_Ticker
ws.Cells(2, 16).Value = Largest_Percentage
ws.Cells(2, 16).NumberFormat = "0.00%"

ws.Cells(3, 15) = Smallest_Percentage_Ticker
ws.Cells(3, 16).Value = Smallest_Percentage
ws.Cells(3, 16).NumberFormat = "0.00%"

ws.Cells(4, 15) = Largest_Volume_Ticker
ws.Cells(4, 16).Value = Largest_Total_Stock_Volume
ws.Cells(4, 16).NumberFormat = "#"

ws.Range("I:P").EntireColumn.AutoFit

Next ws

End Sub
