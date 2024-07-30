Attribute VB_Name = "Module1"
Sub Multiple_Quarter_Master_Script()
    'Set variables for tickers, lastrow of data, i in the for loop, and references to condense the data
    Dim ticker As String
    Dim lastrow As Long
    Dim i As Long
    Dim summary_row_table As Integer
    Dim openvalue As Double
    Dim closevalue As Double
    Dim quarterlychange As Double
    Dim percentchange As Double
    Dim totalvolume As LongLong
    Dim maxpercentchange As Double
    Dim minpercentchange As Double
    Dim maxvolume As LongLong
    
    'Loop through all Worksheets
    For Each ws In Worksheets
    
        'Set headers for the columns
        ws.Range("H1").Value = "Ticker"
        ws.Range("I1").Value = "Quarterly Change"
        ws.Range("J1").Value = "Percent Change"
        ws.Range("K1").Value = "Total Volume"
    
        'Set initial reference row to 2 so it doesn't overlap into the header
        summary_row_table = 2
    
        'Determine last row using count function
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Initialize variables
        openvalue = 0
        closevalue = 0
        totalvolume = 0
        maxpercentchange = -9999
        minpercentchange = 9999
        maxvolume = 0
    
        'Loop through all rows of data
        For i = 2 To lastrow
    
            'Ticker/Open Value(Used to calculate difference and percent change) Creator
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                openvalue = ws.Cells(i, 3).Value
            End If
        
            'Quarterly/Percent Change/Total Volume Creator
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                closevalue = ws.Cells(i, 6).Value
            Else
                closevalue = ws.Cells(i, 6).Value
                quarterlychange = closevalue - openvalue
                percentchange = ((closevalue - openvalue) / openvalue)
                
                'Check for greatest percentage changes
                If percentchange > maxpercentchange Then
                    ws.Cells(2, 16).Value = percentchange
                    maxpercentchange = percentchange
                    ws.Cells(2, 15).Value = ticker
                ElseIf percentchange < minpercentchange Then
                    ws.Cells(3, 16).Value = percentchange
                    minpercentchange = percentchange
                    ws.Cells(3, 15).Value = ticker
                End If
                
                'Color Indexing for percentage column
                If percentchange * 100 > 0 Then
                    ws.Range("J" & summary_row_table).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & summary_row_table).Interior.ColorIndex = 3
                End If
                
                'Color Indexing for Quarterly Change column
                If quarterlychange > 0 Then
                    ws.Range("I" & summary_row_table).Interior.ColorIndex = 4
                Else
                    ws.Range("I" & summary_row_table).Interior.ColorIndex = 3
                End If
                
                ws.Range("H" & summary_row_table).Value = ticker
                ws.Range("I" & summary_row_table).Value = quarterlychange
                ws.Range("J" & summary_row_table).Value = percentchange
                ws.Range("K" & summary_row_table).Value = totalvolume
            
                summary_row_table = summary_row_table + 1
                totalvolume = 0
                
            End If
        
            totalvolume = totalvolume + Cells(i, 7).Value
            
            'Check for greatest total volume
            If totalvolume > maxvolume Then
                ws.Cells(4, 16).Value = totalvolume
                maxvolume = totalvolume
                ws.Cells(4, 15).Value = ticker
            End If
                
        Next i
        
    Next ws
    
End Sub

