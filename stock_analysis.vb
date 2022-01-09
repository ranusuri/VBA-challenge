

Sub stock_analysis()

    'Declare a variable to hold current row in excel.
    Dim row As Long
    
    ' Set an initial variable for holding the ticker name
    Dim ticker_Name As String
    
    ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Add the word Ticker to the First Column Header for Summary table
    Range("I" & 1).Value = "Ticker"
    
    ' Add the word Yearly change to the second Column Header for Summary table
    Range("J" & 1).Value = "Yearly Change"
    
    ' Add the word Percent change to the second Column Header for Summary table
    Range("K" & 1).Value = "Percent Change"
    
    ' Add the word Total Stock Volume to the second Column Header for Summary table
    Range("L" & 1).Value = "Total Stock Volume Change"

    'Variable to hold open value of ticker
    Dim open_value As Double
    
    'Variable to hold close value of ticker
    Dim close_value As Double
    
    'Variable to hold yearly change for given ticker
    Dim yearly_change As Double
    
    'Varilable to hold ticker start status
    Dim ticker_start As Boolean
    
    ticker_start = True
    
    Dim stock_increase As Double
    
    'variable to hold percent change for given ticker
    Dim percent_change As Double
    
    'variable to hold total stock volume
    Dim total_stock_volume As Double
    
    'Loop through each row in excel.
    For row = 2 To Rows.Count - 1
        
        ' Check if we are still within the same ticker, if it is not...
        If Cells(row + 1, 1).Value <> Cells(row, 1).Value Then
    
            ' Set the ticker name
            ticker_Name = Cells(row, 1).Value
            
            ' Print the Ticker in the Summary Table
            Range("I" & Summary_Table_Row).Value = ticker_Name
            
            'read close value at end of ticker row
            close_value = Cells(row, 6).Value
            
            yearly_change = close_value - open_value
            
            ' Print the Yearly change in the Summary Table
            Range("J" & Summary_Table_Row).Font.ColorIndex = 1
            Range("J" & Summary_Table_Row).Value = yearly_change
            
            If yearly_change < 0 Then
                ' Set the Font color to Red
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            Else
                ' Set the Font color to Green
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            End If
            
            stock_increase = close_value - open_value
            
            If open_value <> 0 Then
                percent_change = stock_increase / open_value
            Else
                percent_change = 0
            End If
              
            'Range("K" & Summary_Table_Row).NumberFormat = "Percent"
            
            ' Print the Percentage change in the Summary Table
            'Range("K" & Summary_Table_Row).Value = percent_change
            Range("K" & Summary_Table_Row).Value = Format(percent_change, "Percent")
            
            'add closed colume to total_stock_volumne
            total_stock_volume = total_stock_volume + Cells(row, 7).Value
            
            'Print the total stock volumne in the summary table
            Range("L" & Summary_Table_Row).Value = total_stock_volume

            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            ticker_start = True
            
        Else
            'if loop is at start of the ticker,  store open value and set ticker_start as false
            If ticker_start = True Then
               open_value = Cells(row, 3).Value
               
               ticker_start = False
            End If
            
            'read volume for each row and add it to total stock volume
            total_stock_volume = total_stock_volume + Cells(row, 7).Value
                
            'add totals here
           ' open_value = Cells(row, 3).Value
           ' MsgBox (Cells(row, 1).Value)
           ' MsgBox (open_value)
                
        End If
            
    Next row

    

End Sub




