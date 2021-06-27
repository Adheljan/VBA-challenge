Attribute VB_Name = "Module1"
Sub Stocks_Analysis()


        Dim Ws As Worksheet
       
        For Each Ws In Worksheets
            Dim Ticker_Name As String
            Dim Open_Year As Double
            Open_Price = 0
            Dim Close_Year As Double
            Close_Price = 0
            Dim Yearly_Change As Double
            Delta_Price = 0
            Dim Percent_Change As Double
            Delta_Percent = 0
            Dim Volume As Double
            Volume = 0
    
            Dim Summary_Table_Row As Long
            Summary_Table_Row = 2
            
            Dim Lastrow As Long
            Dim i As Long
            
            Lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
            
                Ws.Range("I1").Value = "Ticker"
                Ws.Range("J1").Value = "Yearly Change"
                Ws.Range("K1").Value = "Percent Change"
                Ws.Range("L1").Value = "Total Stock Volume"

            Open_Year = Ws.Cells(2, 3).Value
            
            For i = 2 To Lastrow
            
                If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
                
                    Ticker_Name = Ws.Cells(i, 1).Value
                    Close_Year = Ws.Cells(i, 6).Value
                    Yearly_Change = Close_Year - Open_Year

                    If Open_Year <> 0 Then
                        Percent_Change = (Yearly_Change / Open_Year) * 100
                    Else

                    End If

                    Volume = Volume + Ws.Cells(i, 7).Value
                  
                    
                    Ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

                    Ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

                    If (Yearly_Change > 0) Then

                        Ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    ElseIf (Yearly_Change <= 0) Then

                        Ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If
                    

                    Ws.Range("K" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")

                    Ws.Range("L" & Summary_Table_Row).Value = Volume
                    
                    Summary_Table_Row = Summary_Table_Row + 1

                    Yearly_Change = 0
                    Percent_Change = 0
                    Close_Year = 0

                    Open_Year = Ws.Cells(i + 1, 3).Value
                    Volume = 0

                Else

                    Volume = Volume + Ws.Cells(i, 7).Value
                    
                End If

          
            Next i
            
         Next Ws
End Sub
