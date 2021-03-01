Public Sub Stockmarket()

'LOOP THROUGH ALL THE SHEETS
        For Each ws In Worksheets
    
'DETERMINE THE LAST ROW
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        worksheetname = ws.Name

'INSERT TABLE 1 ON EVERY SHEET
        ws.Range("K1").Value = " Ticker Symbol"
        ws.Range("l1").Value = " Yearly Change"
        ws.Range("m1").Value = "Percent Change"
        ws.Range("n1").Value = "Total Stock Volume"
   
'INSERT TITLES FOR TABLE 2 ON EVERY SHEETS
        ws.Range("q1").Value = " Ticker"
        ws.Range("r1").Value = "Value"
        ws.Range("p3").Value = " Greatest % Increase"
        ws.Range("p4").Value = " Greatest % decrease"
        ws.Range("P5").Value = "Greatest total Volume"

'DECLARE VARIABLES
        Dim tickersymbol As String
        Dim totalvolume As Double
            totalvolume = 0
        Dim openprice  As Double
        Dim closeprice As Double
        Dim table As Integer
            table = 3
        Dim yearly_change As Double
            yearly_change = 0
        Dim percent_change As Double
            percent_change = 0
            
            
        openprice = ws.Cells(2, 3).Value
        
'LOOP TROUGH ALL TICKERS
        For i = 2 To lastrow

                    
           
         'last ticker row so i is the last row
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                closeprice = ws.Cells(i, 6).Value
                tickersymbol = ws.Cells(i, 1).Value
            
                'YEARLY CHANGE
                yearly_change = closeprice - openprice
                
                If yearly_change > 0 Then
                     'highlight green
                     ws.Range("l" & table).Interior.ColorIndex = 4
                Else
                    'highlight red
                    ws.Range("l" & table).Interior.ColorIndex = 3
                End If
                
                'PERCENTAGE CHANGE
                If openprice And yearly_change <> 0 Then
                    percent_change = yearly_change / openprice
                ElseIf openprice Or yearly_change = 0 Then
                    percent_change = 0
                End If
                
                ws.Range("m" & table).Style = "Percent"
            
                'grab new openprice; we are on the last row and its on the next row
                openprice = ws.Cells(i + 1, 3).Value
            
                'ADD VOLUMES
                totalvolume = totalvolume + ws.Cells(i, 7).Value
        
                'Print ticker symbol, yearly change, percentage change and volume in the table
                ws.Range("k" & table).Value = tickersymbol
                ws.Range("n" & table).Value = totalvolume
                ws.Range("l" & table).Value = yearly_change
                ws.Range("m" & table).Value = percent_change
         
                'add other rows
                table = table + 1
                
                totalvolume = 0
                
               Else
               totalvolume = totalvolume + ws.Cells(i, 7).Value
               
         
            End If
        Next i
    
    
'###CHALLENGE: PULL MAX/MIN PERCENT CHANGE AND MAX VOLUME

'First Declare variables
        Dim maxincrease As Double
        maxincrease = 0
        
        Dim maxdecrease As Double
        maxdecrease = 0
        
        Dim maxvolume As Double
        maxvolume = 0
        
        Dim ticker As String
        ticker = ws.Cells(i, 11).Value
              
    
'Pull Max/min and max volume
For i = 3 To 291


        If maxincrease < ws.Cells(i, 13).Value Then
             maxincrease = ws.Cells(i, 13).Value
             ws.Range("r3").Value = maxincrease
             ws.Range("r3").Style = "percent"
             ws.Range("q3").Value = ws.Cells(i, 11).Value
             
             
        ElseIf maxdecrease > ws.Cells(i, 13).Value Then
             maxdecrease = ws.Cells(i, 13).Value
             ws.Range("r4").Value = maxdecrease
             ws.Range("r4").Style = "Percent"
             ws.Range("q4").Value = ws.Cells(i, 11).Value
    
             
        ElseIf maxvolume < ws.Cells(i, 14).Value Then
             maxvolume = ws.Cells(i, 14).Value
             ws.Range("r5").Value = maxvolume
             ws.Range("q5").Value = ws.Cells(i, 11).Value
            
        End If
        

        
Next i

Next ws

End Sub


