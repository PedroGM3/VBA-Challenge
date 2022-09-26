Attribute VB_Name = "Module1"
Sub Stocks()

'Define variable for the ticker

Dim ticker As String

'Variable for yearly change

Dim yearlyChange As Double

Dim begyr As Double

Dim endyr As Double

'variable for percentage change

Dim percentChange As Double

'varible to add the total volume

Dim volume As Double

volume = 0

'Define counter for rows in summary table
'starts in first row

Dim counter As Integer
counter = 1

'print headers

        Cells(counter, 10).Value = "Ticker"
        Cells(counter, 11).Value = "Yearly Change"
        Cells(counter, 12).Value = "Percent Change"
        Cells(counter, 13).Value = "Total Stock Volume"


'Determine the Last Row

Dim LastRow As Double

LastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row

'MsgBox (LastRow)

' Loop the tickers adding volume to the total counter

For i = 2 To LastRow


    volume = volume + Cells(i, 7).Value
    
    'if previous cell is different, collect open value
    
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
    
        begyr = Cells(i, 3).Value
        
    End If
    
     'If next cell is different, collect closing value
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        endyr = Cells(i, 6).Value
        
        'collect ticker into variable

        ticker = Cells(i, 1).Value
        
        'calculate change and % change
        
        yearlyChange = endyr - begyr
        
        percentChange = endyr / begyr - 1
    
        'print ticker and data in summary table
        
        counter = counter + 1
        
        Cells(counter, 10).Value = ticker
        Cells(counter, 11).Value = yearlyChange
        If yearlyChange > 0 Then
            Cells(counter, 11).Interior.Color = 5287936
            Else
            Cells(counter, 11).Interior.Color = 255
        End If
        
        Cells(counter, 12).Value = percentChange
        Cells(counter, 13).Value = volume
        
        'set variables to 0 again
        
        volume = 0
        
        
    End If
        

Next i

'Format summary table
    Columns("L:L").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Columns("K:K").Select
    Selection.Style = "Comma"
    Columns("M:M").Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
   
   
' "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
   
   'define variables
   
   Dim GreatInc As Double
   Dim GreatDec As Double
   Dim GreatVol As Double
   Dim ticker2 As String
   Dim ticker3 As String
   Dim ticker4 As String
    GreatInc = 0
    GreatDec = 0
    GreatVol = 0
   
   
  'select new last row for summary table
  
   LastRow2 = ActiveSheet.Cells(Rows.Count, 10).End(xlUp).Row
   
   'run loop to select high and low values
   
   For j = 2 To LastRow2

    If Cells(j, 12).Value > GreatInc Then
        GreatInc = Cells(j, 12).Value
        ticker2 = Cells(j, 10).Value
    End If
    
    If Cells(j, 12).Value < GreatDec Then
        GreatDec = Cells(j, 12).Value
        ticker3 = Cells(j, 10).Value
    End If
    
    If Cells(j, 13).Value > GreatVol Then
        GreatVol = Cells(j, 13).Value
        ticker4 = Cells(j, 10).Value
    End If
    
    Next j
    
    'print results
    
    Cells(2, 15).Value = "Greatest % increase"
    Cells(3, 15).Value = "Greatest % decrease"
    Cells(4, 15).Value = "Greatest total volume"
    Cells(2, 16).Value = ticker2
    Cells(3, 16).Value = ticker3
    Cells(4, 16).Value = ticker4
    Cells(2, 17).Value = GreatInc
    Cells(3, 17).Value = GreatDec
    Cells(4, 17).Value = GreatVol
    
    'format cells
    
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(4, 17).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
End Sub
