Attribute VB_Name = "Module1"
Sub Stock_price()

    'Define worksheet
    Dim ws As Worksheet
    
    'Loop through worksheets
    For Each ws In Worksheets
        
        'Define variables
        Dim i As Long
        Dim Quart As Double
        Dim vol As Double
        Dim lastRow As Long
        Dim Table_Row As Long
        Dim maxValue As Double
        Dim associatedValueMax As Variant
        Dim maxRow As Long
        Dim minValue As Double
        Dim associatedValueMin As Variant
        Dim minRow As Long
        Dim maxVol As Double
        Dim associatedVolMax As Variant
        Dim maxVolRow As Long
        
        'Set counters
        Table_Row = 2
        Quart = 0
        vol = 0
        
        'Define last row
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
                
        'Name columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
                
        ' Apply conditional formatting to column J
        With ws.Range("J2:J" & lastRow).FormatConditions
            ' Clear any existing conditions
            .Delete
            
            ' Add condition for positive values - Green fill
            .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .Item(1).Interior.ColorIndex = 4
            
            ' Add condition for negative values - Red fill
            .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .Item(2).Interior.ColorIndex = 3
        
        End With
        
        ' Format Percent Change and value as percentage
        ws.Range("K2:K" & lastRow).NumberFormat = "0.00%"
        ws.Range("Q2:Q3" & lastRow).NumberFormat = "0.00%"
        ws.Range("Q4:Q4" & lastRow).NumberFormat = "0"
        
        'Loop through rows
        For i = 2 To lastRow
            
            'Check if the ticker changes
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                'Insert ticker
                ws.Cells(Table_Row, 9).Value = ws.Cells(i, 1).Value
                
                'Quarterly change
                ws.Cells(Table_Row, 10).Value = ws.Cells(i, 6).Value - Quart
                
                'Calculate percent change if Quart is not zero
                If Quart <> 0 Then
                    ws.Cells(Table_Row, 11).Value = (ws.Cells(i, 6).Value - Quart) / Quart
                Else
                    ws.Cells(Table_Row, 11).Value = 0
                End If
                
                'Add and insert volume to table
                vol = vol + ws.Cells(i, 7).Value
                ws.Cells(Table_Row, 12).Value = vol
                
                'Move to the next table row
                Table_Row = Table_Row + 1
                
                'Reset counters
                Quart = 0
                vol = 0
            
            ElseIf Quart = 0 Then
                'Set Quart to the opening price for the first occurrence of the ticker
                Quart = ws.Cells(i, 3).Value
            
            End If
            
            'Add volume
            vol = vol + ws.Cells(i, 7).Value
            
        Next i
        
    'Greatest increase
            
            'find max value
        maxValue = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastRow))
            
            'enter max value in cell
        ws.Cells(2, 17).Value = maxValue
            
            'Find max row
        maxRow = Application.WorksheetFunction.Match(maxValue, ws.Range("K2:K" & lastRow), 0) + 1
        associatedValueMax = ws.Cells(maxRow, "I").Value
            
            'enter max row
        ws.Cells(2, 16).Value = associatedValueMax
        
    'Greatest Decrease
            
            'find min value
        minValue = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastRow))
            
            'enter min value in cell
        ws.Cells(3, 17).Value = minValue
            
            'find min row
        minRow = Application.WorksheetFunction.Match(minValue, ws.Range("K2:K" & lastRow), 0) + 1
        associatedValueMin = ws.Cells(minRow, "I").Value
            
            'enter min row
        ws.Cells(3, 16).Value = associatedValueMin
        
    'Greatest total Volume
            
            'find max volume
        maxVol = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
            
            'enter max volume in cell
        ws.Cells(4, 17).Value = maxVol
            
            'find max volume cell
        maxVolRow = Application.WorksheetFunction.Match(maxVol, ws.Range("L2:L" & lastRow), 0) + 1
        associatedVolMax = ws.Cells(maxVolRow, "I").Value
            
            'enter max volume into cell
        ws.Cells(4, 16).Value = associatedVolMax
        
        'fit cells
        ws.Columns.AutoFit
    
    Next ws

End Sub


