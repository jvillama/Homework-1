Sub Stocks()

    For Each ws In Worksheets

        Dim WorksheetName As String
        
        WorksheetName = ws.Name
        
        MsgBox ("Sheet " & WorksheetName)
        
            Dim Stock_Name As String
            
            Dim Stock_Total_Vol As Double
            
            Dim Sub_Total_Vol As Integer
            Sub_Total_Vol = 2
            
            'Dim First_Value As Double
            
            'Dim Last_Value As Double
            
            last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
                        
            For I = 2 To last_row
                
                
                If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                    
                    
                    ws.Range("K1").Value = "Ticketer Name"
                    
                    ws.Range("L1").Value = "Yearly Change"
                    
                    ws.Range("M1").Value = "Percent Change"
                    
                    ws.Range("N1").Value = "Total Stock Volume"
                    
                    Stock_Name = ws.Cells(I, 1).Value
                    
                    'First_Value = ws.Cells(I, 3).Value
                    
                    'Last_Value = ws.Cells(I, 6).Value
                    
                    Stock_Total_Vol = Stock_Total_Vol + ws.Cells(I, 7).Value
            
                    ws.Range("K" & Sub_Total_Vol).Value = Stock_Name
                    
                    ws.Range("L" & Sub_Total_Vol).Value = First_Value
                    
                    ws.Range("M" & Sub_Total_Vol).Value = Last_Value
            
                    ws.Range("N" & Sub_Total_Vol).Value = Stock_Total_Vol

                    Sub_Total_Vol = Sub_Total_Vol + 1
            
                    Stock_Total_Vol = 0
                    
                    ws.Range("N2:N" & Sub_Total_Vol).Style = "Currency"

                                     
                Else
            
                    Stock_Total_Vol = Stock_Total_Vol + ws.Cells(I, 7).Value
            
                    'ws.Columns("A:N").AutoFit
                    
                    End If
            
            Next I

    Next ws
    
    MsgBox ("Done")
    
End Sub
