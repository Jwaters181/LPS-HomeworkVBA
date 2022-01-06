Attribute VB_Name = "Module1"

Sub StockTicker()
Dim CurrentWs As Worksheet
Dim Req_Summary_Table As Boolean
Req_Summary_Table = False
For Each CurrentWs In Worksheets
Dim Ticker_Name As String
Ticker_Name = " "
Dim Total_Ticker_Volume As Double
Total_Ticker_Volume = 0
Dim Start_Price As Double
Start_Price = 0
Dim Close_Price As Double
Close_Price = 0
Dim Yearly_Change As Double
Yearly_Change = 0
Dim Percent_Change As Double
Percent_Change = 0
Dim Ticker_Table_Row As Long
Ticker_Table_Row = 2
Dim Lastrow As Long
Dim i As Long
Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row
If Req_Summary_Table Then
CurrentWs.Range("I1").Value = "Ticker"
CurrentWs.Range("J1").Value = "Yearly Change"
CurrentWs.Range("K1").Value = "Percent Change"
CurrentWs.Range("L1").Value = "Total Stock Volume"
Else
Req_Summary_Table = True
End If
Start_Price = CurrentWs.Cells(2, 3).Value
For i = 2 To Lastrow
If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
Ticker_Name = CurrentWs.Cells(i, 1).Value
Close_Price = CurrentWs.Cells(i, 6).Value
Yearly_Change = Close_Price - Start_Price
If Start_Price <> 0 Then
Percent_Change = (Yearly_Change / Start_Price) * 100
Else: MsgBox ("For" & Ticker_Name & ", Row" & CStr(i) & ": Start Price = " & Start_Price & ".Fix <open> field manually and save the spreadsheet.")
End If
Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
CurrentWs.Range("I" & Ticker_Table_Row).Value = Ticker_Name
CurrentWs.Range("J" & Ticker_Table_Row).Value = Yearly_Change
If (Yearly_Change > 0) Then
CurrentWs.Range("J" & Ticker_Table_Row).Interior.ColorIndex = 4
ElseIf (Yearly_Change <= 0) Then
CurrentWs.Range("J" & Ticker_Table_Row).Interior.ColorIndex = 3
End If
CurrentWs.Range("K" & Ticker_Table_Row).Value = (CStr(Percent_Change) & "%")
CurrentWs.Range("L" & Ticker_Table_Row).Value = Total_Ticker_Volume
Ticker_Table_Row = Ticker_Table_Row + 1
Yearly_Change = 0
Percent_Change = 0
Start_Price = CurrentWs.Cells(i + 1, 3).Value
Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
End If
Next i
Next CurrentWs










End Sub
