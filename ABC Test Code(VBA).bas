Attribute VB_Name = "Module1"

Sub stockmarket()

Dim ws As Worksheet

'Loop through worksheets

For Each ws In Worksheets
 ws.Activate
 
'Dim each variable

  Dim Ticker As String
Dim Yrc As Double
Dim opn As Double
Dim cls As Double
Dim Yrc2 As Double
Dim Pc As Double
  Dim SV As Double
  Dim max As Double
  Dim min As Double
  Dim maxv As Double
  

  Pc = 0
  Yrc = 0
  opn = 0
  Yrc2 = 2
  SV = 0

 
  Dim Summary_Table_Row As Long
  Summary_Table_Row = 2

  
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  

  For i = 2 To LastRow

    ' Check for different ticker
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker name
      Ticker = Cells(i, 1).Value

      ' Add to the Volume Total
      SV = SV + Cells(i, 7).Value

      'Input Ticker
      Range("I" & Summary_Table_Row).Value = Ticker

      ' Input Total Stock Volume
      Range("L" & Summary_Table_Row).Value = SV

' Find open and close value
        opn = Cells(i - i + Yrc2, 3).Value
        cls = Cells(i, 6).Value
        
 'Find and Input Year Change
  
       Yrc = cls - opn
        Range("J" & Summary_Table_Row).Value = Yrc
        
 'Adjust for open row
        Yrc2 = 0
        Yrc2 = i + 1
        
'Find and Input Percent Change
If opn = 0 Then
    Pc = 0

Else

    Pc = Yrc / opn
    Range("k" & Summary_Table_Row).Value = Pc
    Range("k" & Summary_Table_Row).NumberFormat = "0.00%"
    
End If
    
'Change the color
If Range("k" & Summary_Table_Row).Value > 0 Then
    Range("k" & Summary_Table_Row).Interior.ColorIndex = 4
Else
    Range("k" & Summary_Table_Row).Interior.ColorIndex = 3

End If

' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
' Reset the Totals
      SV = 0
      Yrc = 0
      Pc = 0
      opn = 0
      cls = 0
      
'Add all Identical Volumes

    Else

      ' Add to the Volume Total
      SV = SV + Cells(i, 7).Value

    End If

  Next i
    
'Find and Input Max/Min
Dim Ticker2 As String
Dim Ticker3 As String
Dim Ticker4 As String
max = 0
min = 0
maxv = 0

'If for max value

  For j = 2 To LastRow
    If Cells(j, 11).Value > max Then
        max = Cells(j, 11).Value
        Ticker2 = Cells(j, 9)
    End If
    
'If for min value

    If Cells(j, 11).Value < min Then
        min = Cells(j, 11).Value
        Ticker3 = Cells(j, 9)
    End If
    
'If for Max Volume

    If Cells(j, 12).Value > maxv Then
        maxv = Cells(j, 12).Value
        Ticker4 = Cells(j, 9)
    End If
  Next j
  
  'Input Values
  
    Range("P2").Value = Ticker2
  Range("Q2").Value = max
    Range("P3").Value = Ticker3
  Range("Q3").Value = min
    Range("P4").Value = Ticker4
  Range("Q4").Value = maxv
 
'Format to percentages

  Range("Q2").NumberFormat = "0.00%"
  Range("Q3").NumberFormat = "0.00%"
  
'Labels
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'Fit the data
Columns("A:Q").AutoFit

Next ws


End Sub


