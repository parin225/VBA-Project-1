Sub Multipleyearstockdata()

'Declare the variables
Dim ticker As String
Dim vol As Double
Dim i As Double
Dim LastRow As Double
Dim xsheet As Worksheet
Dim summarytable As Double
Dim totalstockvolume As Double
Dim yearlychange As Double
Dim percentchange As Double
Dim yearopen As Double
Dim yearclose As Double
Dim row As Double

'Assign the variables
summarytable = 2
vol = 0
totalstockvolume = 0
Application.ScreenUpdating = False

'Loop through all worksheets
For Each xsheet In ThisWorkbook.Worksheets
      xsheet.Activate

       'Insert headlines on excel on each sheet
       xsheet.cells(1, 9).Value = "ticker"
       xsheet.cells(1, 10).Value = "yearly change"
       xsheet.cells(1, 11).Value = "percent change"
       xsheet.cells(1, 12).Value = "total stock volume"

       'Loop through tickers
       LastRow = cells(rows.Count, 1).End(xlUp).row
       ticker = xsheet.cells(2, 1).Value
       yearopen = xsheet.cells(2, 3).Value
       For i = 2 To LastRow

           'Check if we are still within the same stock ticker, do an if then conditional statement
           If ticker <> xsheet.cells(i, 1).Value Then

               'Print the ticker name into the summary table
               xsheet.Range("I" & summarytable).Value = ticker
               ticker = xsheet.cells(i, 1).Value

              'Print the yearly change into the summary table
               yearclose = xsheet.cells(i - 1, 6).Value
               yearlychange = yearclose - yearopen
               xsheet.Range("J" & summarytable).Value = yearlychange

              'Print the percent change into the summary table
              If yearopen = 0 Then
                   For row = i To LastRow
                       If xsheet.cells(i, 3).Value > 0 Or xsheet.cells(i, 3).Value < 0 Then
                           yearopen = xsheet.cells(i, 3).Value
                           Exit For
                       End If
                   Next row
              End If
               percentchange = (yearclose - yearopen) / yearopen
               xsheet.Range("K" & summarytable).Value = percentchange

              'Set year open
              yearopen = xsheet.cells(i, 3).Value

              'Change format to percent
               xsheet.Columns("K").NumberFormat = "0.00%"

              'Print the total stock volume into the summary table
               xsheet.Range("L" & summarytable).Value = vol
               'totalstockvolume = totalstockvolume + vol

               'Add one to the summary table
               summarytable = summarytable + 1

               'Reset the total stock volume amount
               vol = 0


           End If

           vol = vol + xsheet.cells(i, 7).Value

           'Formatting color

           If xsheet.cells(i, 10).Value > 0 Then
                xsheet.cells(i, 10).Interior.ColorIndex = 4
           Else
               xsheet.cells(i, 10).Interior.ColorIndex = 3
           End If

       Next i
       
'reset summarytable to 2 before moving onto next sheet
 summarytable = 2
 
'move to next worksheet
Next xsheet



Application.ScreenUpdating = True


End Sub


