Attribute VB_Name = "Module2"
Sub Button1_Click()

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets ' Loop through all Worksheet, loop begins here

    ws.Activate ' Re activate the original woorksheet after loop  is done through all woorksheet

    'variable declearation

    Dim lastrowno As Double
    Dim openprice As Double
    Dim closeprice As Double
    Dim ticker As String
    Dim percentchange As Double
    Dim volume As Double
    Dim i As Double
    Dim j As Double
    Dim counter As Double
    

    'set value of some variable
    Row = 2
    Column = 1
    volume = 0



    ' Adding Header For result
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "yearly Change"
    ws.Cells(1, 11).Value = "Percent change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    

    ' set initial openprice
    openprice = Cells(2, Column + 2).Value

    ' finding Last Row No
    lastrowno = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' loop through all Ticker Symbol
        
        For i = 2 To lastrowno


            If ws.Cells(i + 1, Column).Value <> ws.Cells(i, Column).Value Then ' checks whether ticker symbole match or not
                ' if condition true then
                ticker = ws.Cells(i, Column).Value ' paste cell value to ticker variable
                ws.Cells(Row, Column + 8).Value = ticker ' paste ticker value to cell
        
                closeprice = ws.Cells(i, Column + 5).Value ' set cell value to closing price variable
        
                ' yearly change calculation
                yearlychange = closeprice - openprice
                ws.Cells(Row, Column + 9).Value = yearlychange
        
                'percentchange calculation
                If openprice <> 0 Then
                    percentchange = yearlychange / openprice
                    ws.Cells(Row, Column + 10).Value = percentchange
                    ws.Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                'Total Volume calculation
                volume = volume + ws.Cells(i, Column + 6).Value
                ws.Cells(Row, Column + 11).Value = volume
        
                Row = Row + 1 ' Increase value of row by 1 for next loop
                openprice = ws.Cells(i + 1, Column + 2) 'set open price to next cell open price value
                volume = 0 ' set volume value to 0
                counter = 0
            End If
            
            If counter = 0 Then
            counter = counter + 1
            Else
            volume = volume + ws.Cells(i, Column + 6).Value  'Every time i Increase, it adds coresponding row volume
            End If
        Next i
             
             
             

        ' set color based on positive and negative yearly chnage
        For j = 2 To lastrowno
        If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
        Cells(j, Column + 9).Interior.ColorIndex = 4
        ElseIf Cells(j, Column + 9).Value < 0 Then
        Cells(j, Column + 9).Interior.ColorIndex = 3
        End If
        
        Next j
 
 
        'Hard solution
         
            
        Dim greatestincrease As Double
        Dim greatestdecrese As Double
        Dim greatestvolume As Double
        Dim gi_ticker As String
        Dim gd_ticker As String
        Dim gv_ticker As String
        Dim k As Double
        
        greatestdecrese = Cells(2, 11).Value
        greatestincrease = Cells(2, 11).Value
        greatestvolume = Cells(2, 12).Value
        For k = 2 To lastrowno
        
            If greatestincrease < Cells(k + 1, 11).Value Then
                greatestincrease = Cells(k + 1, 11).Value
                gi_ticker = Cells(k + 1, 9).Value
            ElseIf greatestdecrese > Cells(k + 1, 11).Value Then
                greatestdecrese = Cells(k + 1, 11).Value
                gd_ticker = Cells(k + 1, 9).Value
            ElseIf greatestvolume < Cells(k + 1, 12).Value Then
                greatestvolume = Cells(k + 1, 12).Value
                gv_ticker = Cells(k + 1, 9).Value
            End If
                
         Next k
         
        
        Cells(2, 16).Value = gi_ticker
        Cells(2, 17).Value = greatestincrease
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(3, 16).Value = gd_ticker
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(3, 17).Value = greatestdecrese
        Cells(4, 16).Value = gv_ticker
        Cells(4, 17).Value = greatestvolume

   
Next ws


End Sub
