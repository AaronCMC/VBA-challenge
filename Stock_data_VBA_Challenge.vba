Attribute VB_Name = "Module1"
Sub Quarterly_Summary()

    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        ' Insert Header
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        ' Determine days within quarter
        Dim k As Integer
        Dim Days As Integer
        
        For k = 2 To 100
            If ws.Cells(k + 1, 1).Value <> ws.Cells(k, 1).Value Then
                Days = k - 1
                Exit For
            End If
        Next k
        
        ' Extraction
        Dim i As Long
        Dim Stock_Symbol As String
        Dim Total_Stock_Vol As Double
        Dim Daily_Open() As Double
        Dim Daily_Close() As Double
        Dim Summary_Table_Row As Integer
        Dim lastIndex As Integer
        Dim Quarterly_Change As Double
        Dim Percent_Change As Double
        
        Summary_Table_Row = 2
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Stock_Symbol = ws.Cells(i, 1).Value
                ws.Range("I" & Summary_Table_Row).Value = Stock_Symbol
                Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7).Value
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Vol
                Daily_Open(i - 1) = ws.Cells(i, 3).Value
                Daily_Close(i - 1) = ws.Cells(i, 6).Value
                Quarterly_Change = Daily_Close(i - 1) - Daily_Open(i - Days)
                ws.Range("J" & Summary_Table_Row).Value = Quarterly_Change
                Percent_Change = Quarterly_Change / Daily_Open(i - Days)
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                Summary_Table_Row = Summary_Table_Row + 1
                Total_Stock_Vol = 0
                Quarterly_Change = 0
                Erase Daily_Open
                Erase Daily_Close
            Else
                Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7).Value
                ReDim Preserve Daily_Open(1 To lastRow)
                ReDim Preserve Daily_Close(1 To lastRow)
                Daily_Open(i - 1) = ws.Cells(i, 3).Value
                Daily_Close(i - 1) = ws.Cells(i, 6).Value
            End If
        Next i
        
        ' Conditional Formatting
        Dim n As Long
        
        For n = 2 To lastRow
            If ws.Cells(n, 10).Value > 0 Then
                ws.Cells(n, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(n, 10).Value < 0 Then
                ws.Cells(n, 10).Interior.ColorIndex = 3
            End If
        Next n
        
        ' Percentage Formatting
        ws.Range("K2:K" & lastRow).NumberFormat = "0.00%"
        
        ' Max % Increase and Decrease
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest & Decrease"
        
        Dim Max_Percent_Increase As Double
        Dim Max_Percent_Decrease As Double
        Dim x As Integer
        
        Max_Percent_Increase = 0
        Max_Percent_Decrease = 0
        lastRow_Extremes = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For x = 2 To lastRow_Extremes
            If ws.Cells(x, 11).Value > Max_Percent_Increase Then
                Max_Percent_Increase = ws.Cells(x, 11).Value
                ws.Range("Q2").Value = Max_Percent_Increase
                ws.Range("P2").Value = ws.Cells(x, 9).Value
            ElseIf ws.Cells(x, 11).Value < Max_Percent_Decrease Then
                Max_Percent_Decrease = ws.Cells(x, 11).Value
                ws.Range("Q3").Value = Max_Percent_Decrease
                ws.Range("P3").Value = ws.Cells(x, 9).Value
            End If
        Next x
        
        ' Percentage Formatting in Extremes Table
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        ' Max Total Volume
        ws.Range("O4").Value = "Greatest Total Volume"
        
        Dim Max_Total_Volume As Double
        Dim y As Integer
        
        Max_Total_Volume = 0
        
        For y = 2 To lastRow_Extremes
            If ws.Cells(y, 12).Value > Max_Total_Volume Then
                Max_Total_Volume = ws.Cells(y, 12).Value
                ws.Range("Q4") = Max_Total_Volume
                ws.Range("P4") = ws.Cells(y, 9).Value
            End If
        Next y
    
    Next ws
    
End Sub



