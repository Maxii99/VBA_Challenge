Attribute VB_Name = "Module1"
Sub stock_data()

Dim ws As Worksheet
For Each ws In Worksheets

'Define variables

Dim Ticker As String
Dim year_open As Double
Dim yer_close As Double
Dim percent_change As Double
Dim Total_volume As Double
Dim LastRow As Long
Dim start As Long
Dim vol As Double
Dim Summary_Table_Row As Long
start = 2
vol = 0

'To loop through worksheets, define LastRow and other variables
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "yearly_change"
ws.Cells(1, 11).Value = "percent_change"
ws.Cells(1, 12).Value = "Total_volume"
Summary_Table_Row = 2

For i = 2 To LastRow
    If year_open = 0 Then
    year_open = ws.Cells(start, 3).Value
    End If
 
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Ticker = ws.Cells(i, 1).Value
    year_close = ws.Cells(i, 6).Value
    vol = vol + ws.Cells(i, 7).Value
    yearly_change = year_close - year_open

'Find summary values
    ws.Cells(Summary_Table_Row, 9).Value = Ticker
    ws.Cells(Summary_Table_Row, 10).Value = yearly_change
    ws.Cells(Summary_Table_Row, 11).Value = percent_change
    ws.Cells(Summary_Table_Row, 12).Value = vol
    Summary_Table_Row = Summary_Table_Row + 1

    Else
    vol = vol + ws.Cells(i, 7).Value

    End If
    
'Round up percentage change to 2 decimal places (with suggestion from Zach, TA)

    If year_open = 0 Then
    percentage_change = 0
    Else
    percent_change = Round((yearly_change / year_open) * 100, 2)
 
    End If
 

'Color formatting using range
'------------------------------------------------------

'Fill yearly change with green and red for good and bad respectively

    If ws.Range("J" & Summary_Table_Row).Value > 0 Then
    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    
    ElseIf ws.Range("J" & Summary_Table_Row).Value <= 0 Then
    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3

    End If
  
'Trying the challenge bonus
'Define new variables for max and min values for challenge bonus
Dim Max_ticker As Double
Dim Min_ticker As Double
Dim Max_tickername As String
Dim Min_tickername As String
Max_tickername = " "
Min_tickername = " "
Dim Max_percent As Double
Dim Min_percent As Double
Dim Max_tickervol As String
Dim Min_tickervol As String
Dim Max_volume As Double

'Fill in new Summary Table with additional headers and values as follows:

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest_Percent_Increase"
ws.Range("O3").Value = "Greatest_Percent_Increase"
ws.Range("O4").Value = "Greatest_Total_Volume"

'Determine values for for max, min percent changes and greatest total volume

    If percent_change > Max_percent Then
    Max_percent = percent_change
    Max_tickername = Ticker
    
    ElseIf percent_change < Min_percent Then
    Min_percent = percent_change
    Min_tickername = Ticker
    
    End If

    If vol > Max_volume Then
    Max_volume = vol
    Max_tickervol = Ticker

    End If
 
Next i

Next ws


End Sub

