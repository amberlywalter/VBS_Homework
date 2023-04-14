Attribute VB_Name = "Module1"
Sub stock_data()

Dim a As Integer

a = Application.Worksheets.Count

For x = 1 To a

Worksheets(x).Activate



    Dim ticker As String
    Dim Yearly_change As Double
    Dim Percentage_change As Double
    Dim Open_price As Double
    Dim Total_Stock_Volume As Double
    
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
Open_price = Cells(2, 3).Value
    
    For i = 2 To 759001
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker = Cells(i, 1).Value
        
        Close_price = Cells(i, 6).Value
        Yearly_change = Close_price - Open_price
        Percentage_change = Yearly_change / Open_price * 100
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        Range("J1").Value = "Ticker"
        Range("K1").Value = "Yearly Change"
        Range("L1").Value = "Percentage Change"
        Range("M1").Value = "Total Stock Volume"
        Range("J" & Summary_Table_Row).Value = ticker
        Range("K" & Summary_Table_Row).Value = Yearly_change
        Range("L" & Summary_Table_Row).Value = (CStr(Percentage_change) & "%")
        Range("M" & Summary_Table_Row).Value = Total_Stock_Volume
    Summary_Table_Row = Summary_Table_Row + 1
    Yearly_change = 0
    Percentage_change = 0
    Total_Stock_Volume = 0
 
Open_price = Cells(i + 1, 3).Value

Else
    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
       
End If

Next i

  
Dim maxvalue As Double
Dim maxstock As Variant
Dim currentvalue As Double
Dim currentstock As Variant
Dim smallestvalue As Double
Dim smalleststock As Variant

maxvalue = 0

 For j = 2 To 3001
 currentvalue = Range("L" & j).Value
 currentstock = Range("J" & j).Value
 If currentvalue > maxvalue Then
    maxvalue = currentvalue
    Range("R2").Value = maxvalue
    Range("R2").NumberFormat = "0.00%"
    Range("Q2").Value = currentstock
    
End If

        Range("P2").Value = "Greatest % Increase"
        Range("Q1").Value = "Ticker"

Next j

smallestvalue = 0

For K = 2 To 3001
 currentvalue = Range("L" & K).Value
 currentstock = Range("J" & K).Value
 If currentvalue < smallestvalue Then
    smallestvalue = currentvalue
    Range("R3").Value = smallestvalue
    Range("R3").NumberFormat = "0.00%"
    Range("Q3").Value = currentstock
End If

 Range("P3").Value = "Greatest % Decrease"
 Range("R1").Value = "Value"
Next K

maxvalue = 0

 For L = 2 To 3001
 currentvalue = Range("M" & L).Value
 currentstock = Range("J" & L).Value
 If currentvalue > maxvalue Then
    maxvalue = currentvalue
    Range("R4").Value = maxvalue
    Range("Q4").Value = currentstock
    
End If
    Range("P4").Value = "Greatest Total Volume"
Next L

Dim YearlyChange As Double

YearlyChange = 0

 For N = 2 To 3001
 YearlyChange = Range("K" & N).Value
 If (YearlyChange > 0) Then
   Range("K" & N).Interior.ColorIndex = 4
    
 ElseIf (YearlyChange < 0) Then
    Range("K" & N).Interior.ColorIndex = 3
    
End If
    
Next N

Next x


End Sub


