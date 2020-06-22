VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True



Sub multi():

Dim ws As Worksheet
For Each ws In Worksheets

'variables'
Dim openingprice As Double
Dim closingprice As Double
Dim tickerincrease As String
Dim tickerdecrease As String
Dim yearlychange As Double
Dim percentage As Double
Dim vol As Double
Dim maxvolume As Double
Dim maxpercent As Double
Dim minpercent As Double
Dim min As Double
vol = 0
summary2 = 2
Dim lastrow As Long
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'headers'

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "YearlyChange"
ws.Cells(1, 11).Value = "PercentageChange"
ws.Cells(1, 12).Value = "TotalVolume"
ws.Cells(1, 15).Value = "SignificantChanges"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"


For i = 2 To lastrow

tickerincrease = ws.Cells(i + 1, 1).Value
tickerdecrease = ws.Cells(i - 1, 1).Value
ticker = ws.Cells(i, 1).Value
yearlychange = openingprice - closingprice


'to calulate opening price'
If tickerdecrease <> ticker Then
openingprice = ws.Cells(i, 3).Value
End If

'to isolate values in column A'
If tickerincrease <> ticker Then
ticker = ws.Cells(i, 1).Value

'to calculate total vol and print to column L'
vol = vol + ws.Cells(i, 7).Value
ws.Range("L" & summary2).Value = vol


'isolate tickers from column A and print to column I'
ws.Range("I" & summary2).Value = ticker
closingprice = ws.Cells(i, 6).Value

'calculate yearly change and print to column J'
yearlychange = openingprice - closingprice

ws.Range("J" & summary2).Value = Format(yearlychange, "Currency")
If yearlychange >= 0 Then
ws.Range("J" & summary2).Interior.ColorIndex = 4
Else
ws.Range("J" & summary2).Interior.ColorIndex = 3
End If

'calculate percentage of yearly change and print to column K'
If openingprice = 0 Then
percentage = 0
Else
percentage = yearlychange / openingprice
ws.Range("k" & summary2).Value = Format(percentage, "percent")
End If


'to find min and max ranges'
If ws.Range("K" & summary2).Value > maxpercent Then
maxpercent = ws.Range("K" & summary2).Value
ws.Cells(2, 17).Value = Format(maxpercent, "percent")
ws.Cells(2, 16).Value = ws.Range("I" & summary2).Value
ElseIf ws.Range("K" & summary2).Value < minpercent Then
minpercent = ws.Range("K" & summary2).Value
ws.Cells(3, 17).Value = Format(minpercent, "percent")
ws.Cells(3, 16).Value = ws.Range("I" & summary2).Value
End If

'to find max volume'
If ws.Range("L" & summary2).Value > Max Then
Max = ws.Range("L" & summary2).Value
ws.Cells(4, 17).Value = Max
ws.Cells(4, 16).Value = ws.Range("I" & summary2).Value
End If


vol = 0
summary2 = summary2 + 1
Else
vol = vol + ws.Cells(i, 7).Value
End If

Next i
Next ws
End Sub


