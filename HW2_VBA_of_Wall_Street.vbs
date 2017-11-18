Sub total_volume()

    'Loop through each worksheet
    For Each ws In Worksheets
    
        'Declare Variables
        Dim curr_vol As LongLong
        curr_vol = 0
        Dim vol_row As Long
        vol_row = 2
    
        'Set the headers of the rows that we are creating
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Stock Volume"
        
        'Loop through each row until the last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To LastRow
        
            'Once we hit the end of a certain symbol, do...
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                'Add the volume for this last last row to the total volume
                curr_vol = curr_vol + CLng(ws.Cells(i, 7).Value)
                
                'Write ticker symbol to cell
                ws.Cells(vol_row, 9).Value = ws.Cells(i, 1).Value
                
                'Write total volume to cell
                ws.Cells(vol_row, 10).Value = curr_vol
                
                'Increment row
                vol_row = vol_row + 1
                
                'Set current volume back to zero
                curr_vol = 0
                
            'Otherwise, do...
            Else
            
                'Sum the total volume for this symbol
                curr_vol = curr_vol + CLng(ws.Cells(i, 7).Value)
            End If
        Next i
    Next ws
End Sub