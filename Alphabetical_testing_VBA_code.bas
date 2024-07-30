Attribute VB_Name = "Module1"
Sub ticker_creator()
    'Set variables for tickers, lastrow of data, i in the for loop, and a refrence to condense the data
    Dim ticker As String
    Dim lastrow As Long
    Dim i As Long
    Dim summary_row_table As Integer
    
    'Set header
    Range("H1").Value = "Ticker"
    
    'Set initial refrence row to 2 so it doesnt overlap into the header
    summary_row_table = 2
    
    'Determine lastrow using count function
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
        'Set a for loop to test if the above cell is different from the one below
        'If true pull that value and display it in column H
        For i = 2 To lastrow
        
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ticker = Cells(i, 1).Value
                Range("H" & summary_row_table).Value = ticker
                summary_row_table = summary_row_table + 1
                
            End If
            
        Next i
        
End Sub
Sub Quarterly_Percent_Change()
    'Set neccessary variables
    Dim lastrow As Long
    Dim i As Long
    Dim summary_row_table As Integer
    Dim openvalue As Double
    Dim closevalue As Double
    Dim quarterlychange As Double
    Dim percentchange As Double
    
    'Set Header
    Range("I1").Value = "Quarterly Change"
    Range("J1").Value = "Percent Change"
    
    'Set initial values, determine lastrow
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    summary_row_table = 2
    openvalue = 0
    closevalue = 0
    
    'Create a for loop to loop through all rows of data
    For i = 2 To lastrow
    
        'Create an if statement to find openvalue based on a change in ticker
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            openvalue = Cells(i, 3).Value
        End If
        
        'Create an if statement to find closevalue based on change in ticker
        'Calculate difference of open and closing values
        'Calculate percent change using percent change formula and open/closing values
        'Move quarterly and percent change to appropriate columns
        'Reset the refrence row, and opening/closing values
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            closevalue = Cells(i, 6).Value
            quarterlychange = closevalue - openvalue
            percentchange = ((closevalue - openvalue) / openvalue)
            Range("I" & summary_row_table).Value = quarterlychange
            Range("J" & summary_row_table).Value = percentchange
            summary_row_table = summary_row_table + 1
            openvalue = 0
            closevalue = 0
        End If
    Next i
End Sub
Sub Total_Volume()
    'Set variables
    Dim lastrow As Long
    Dim i As Long
    Dim summary_row_table As Integer
    Dim totalvolume As LongLong

    'Set header
    Range("K1").Value = "Total Volume"

    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    summary_row_table = 2
    totalvolume = 0

    For i = 2 To lastrow
        totalvolume = totalvolume + Cells(i, 7).Value
    
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Or i = lastrow Then
            Range("K" & summary_row_table).Value = totalvolume
            summary_row_table = summary_row_table + 1
            totalvolume = 0
        End If
    Next i
End Sub
