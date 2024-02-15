Attribute VB_Name = "Module1"
Sub Stock_Data_Analysis():
    
    
     'Loop for each worksheet
     For Each ws In Worksheets
     'Enter ticker name
     Dim Ticker_Name As String
     'This variable keeps track of the volume for the current ticker
     'set this ticker to 0 before processing any new ticker
      Dim Ticker_Volume As Double
      Ticker_Volume = 0
        
     ' These variables keep track of a ticker's opening and closing price
     ' We will start by reading/storing the first ticker's opening price
      Dim Opening_Price As Single
      Dim Closing_Price As Single
      Opening_Price = ws.Cells(2, 3).Value
        
        ' These variables will be used at the end to find the MIN and the MAX and the greatest stock volume
        Dim Greatest_Increase As Double
        Dim Greatest_Increase_Ticker As String
        
        Dim Greatest_Decrease As Double
        Dim Greatest_Decrease_Ticker As String
        Dim Greatest_Volume As Double
        Dim Greatest_Volume_Ticker As String
    
        ' This variable keeps track of which row we want to print to
        Dim Output_Row As Integer
        Output_Row = 2
        
       ' Label headers for table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        
        ' Loop through all ticker rows
        ' got the lastrow code from the CensusData Inclass Activity from Week2 Day3 files
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastRow
        
          ' Compare the current ticker name (found in Cells(i, 1)) to the next row's ticker name (Cells(i+1, 1))
          '  If they are different, that means you've arrived at the last row for your current ticker, and
           ' we can calculate open/close and print results
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ' Set the ticker name
                Ticker_Name = ws.Cells(i, 1).Value
                
                ' Add the volume for this row (column 7) to the total we've been keeping in Ticker_Volume
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
                
                ' Find the yearly closing price
                Closing_Price = ws.Cells(i, 6).Value
                
                ' Print the ticker name
                ws.Cells(Output_Row, 9).Value = Ticker_Name
                
                ' Print the Yearly Change (difference between opening and closing price)
                ' Formatting option found here: https://stackoverflow.com/questions/53099470/how-to-round-to-2-decimal-places-in-vba
                ws.Cells(Output_Row, 10).Value = Format((Closing_Price - Opening_Price), "0.00")
                
                ' Print the Percent Change
                ' Format option found here: https://excelvbatutor.com/vba_lesson9.htm
                ws.Cells(Output_Row, 11).Value = Format((Closing_Price - Opening_Price) / Opening_Price, "0.00%")
                
                ' Print the volume total
                ws.Cells(Output_Row, 12).Value = Ticker_Volume
                
                ' Add one to the Output_Row so we don't keep printing in the same row
                'Output row started at 2
                Output_Row = Output_Row + 1
                
                ' Reset the volume total since we're moving onto a new ticker
                Ticker_Volume = 0
                
                ' Update the Opening_Price to the next value
                Opening_Price = ws.Cells(i + 1, 3)
            
            ' If Cells(i, 1) and Cells (i+1, 1) have the same ticker, that means we're still working with
            ' the same ticker. In this case, all we need to do is add the volume to our total and move on
            Else
        
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
        
            End If
        
        Next i
        
        ' Apply conditional formatting to % difference row
        ' This logic was provided by Xpert Learning Assistant
        Dim rng As Range
        Dim cond1 As FormatCondition
        Dim cond2 As FormatCondition
        
        ' Set the range to which you want to apply the conditional formatting
        Set rng = ws.Range("J2:J" & Output_Row - 1)
        
        ' Clear any existing conditional formatting rules
        rng.FormatConditions.Delete
        
        ' Add the first conditional format for negative numbers
        'got the formula from Xpert Learning Assistant
        Set cond1 = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        With cond1
            .Interior.Color = RGB(255, 0, 0) ' Red color
        End With
        
        ' Add the second conditional format for positive numbers
        Set cond2 = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        With cond2
            .Interior.Color = RGB(0, 255, 0) ' Green color
        End With
        
        ' Find Greatest % Increase+ Decrease
        ' Use the first row to start our variables + loop
        Greatest_Increase = ws.Range("K2").Value
        Greatest_Increase_Ticker = ws.Range("I2").Value
        Greatest_Decrease = ws.Range("K2").Value
        Greatest_Decrease_Ticker = ws.Range("I2").Value
        Greatest_Total_Volume = ws.Range("L2").Value
        Greatest_Total_Volume_Ticker = ws.Range("I2").Value
        
        
        ' Now it's time to loop through rest of rows
        For x = 3 To Output_Row - 1
            If ws.Cells(x, 11).Value > Greatest_Increase Then
                Greatest_Increase = ws.Cells(x, 11).Value
                Greatest_Increase_Ticker = ws.Cells(x, 9).Value
            End If
            
            If ws.Cells(x, 11).Value < Greatest_Decrease Then
                Greatest_Decrease = ws.Cells(x, 11).Value
                Greatest_Decrease_Ticker = ws.Cells(x, 9).Value
            End If
             
            If ws.Cells(x, 12).Value > Greatest_Total_Volume Then
                Greatest_Total_Volume = ws.Cells(x, 12).Value
                Greatest_Total_Volume_Ticker = ws.Cells(x, 9).Value
            End If
            
        Next x
            
        ws.Range("Q2").Value = Format(Greatest_Increase, "0.00%")
        ws.Range("P2").Value = Greatest_Increase_Ticker
        ws.Range("Q3").Value = Format(Greatest_Decrease, "0.00%")
        ws.Range("P3").Value = Greatest_Decrease_Ticker
        ws.Range("P4").Value = Greatest_Total_Volume_Ticker
        ws.Range("q4").Value = Greatest_Total_Volume
        
        
    
    Next ws

End Sub


