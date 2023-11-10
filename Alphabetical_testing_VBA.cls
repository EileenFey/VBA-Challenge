VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub alphabet()
Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Volume As LongLong
Dim LastRow As Long
Dim Ticker_Cell As Integer
Dim Ticker_Opening As Integer
Dim i As Long
Dim currMaxIncrease As Double
Dim currMaxIncreaseTicker As String
Dim currMaxDecrease As Double
Dim currMaxDecreaseTicker As String
Dim currMaxVolume As LongLong
Dim currMaxVolumeTicker As String
Dim ws As Worksheet



For Each ws In Worksheets

    Ticker_Cell = 2
    
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    Ticker_Opening = ws.Cells(2, 3).Value
    currMaxIncrease = 0#
    currMaxDecrease = 0#
    currMaxVolume = 0
    
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            ws.Range("O2").Value = "Greatest %  Increase"
            ws.Range("O3").Value = "Greatest %  Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
    
    For i = 2 To LastRow
        
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
        If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
        
            Yearly_Change = ws.Cells(i, 6).Value - Ticker_Opening
            Percent_Change = ((ws.Cells(i, 6).Value - Ticker_Opening) / Ticker_Opening)
            Ticker = ws.Cells(i, 1).Value
            
            ws.Range("I" & Ticker_Cell).Value = Ticker
            ws.Range("J" & Ticker_Cell).Value = Yearly_Change
            ws.Range("J" & Ticker_Cell).NumberFormat = "0.##"
            ws.Range("K" & Ticker_Cell).Value = Percent_Change
            ws.Range("K" & Ticker_Cell).NumberFormat = "##%"
            ws.Cells(Ticker_Cell, 12).Value = Total_Stock_Volume
            
            
            If Total_Stock_Volume > currMaxVolume Then
                currMaxVolume = Total_Stock_Volume
                currMaxVolumeTicker = Ticker
            End If
            
             Select Case Percent_Change
                Case Is > currMaxIncrease
                    currMaxIncrease = Percent_Change
                    currMaxIncreaseTicker = Ticker
                Case Is < currMaxDecrease
                    currMaxDecrease = Percent_Change
                    currMaxDecreaseTicker = Ticker
            End Select
            
            Select Case Yearly_Change
                Case Is < 0
                ws.Range("J" & Ticker_Cell).Interior.Color = RGB(255, 0, 0)
                Case Is > 0
                ws.Range("J" & Ticker_Cell).Interior.Color = RGB(0, 255, 0)
            End Select
            
            Select Case Percent_Change
                Case Is < 0
                ws.Range("K" & Ticker_Cell).Interior.Color = RGB(255, 0, 0)
                Case Is > 0
                ws.Range("K" & Ticker_Cell).Interior.Color = RGB(0, 255, 0)
            End Select
            
            Ticker_Cell = Ticker_Cell + 1
            Total_Stock_Volume = 0
            Ticker_Opening = ws.Cells(i + 1, 6).Value
            
        End If
      
    Next i
    
        ws.Range("Q2") = currMaxIncrease
        ws.Range("Q3") = currMaxDecrease
        ws.Range("Q4") = currMaxVolume
        ws.Range("P2") = currMaxIncreaseTicker
        ws.Range("P3") = currMaxDecreaseTicker
        ws.Range("P4") = currMaxVolumeTicker

Next ws

End Sub
