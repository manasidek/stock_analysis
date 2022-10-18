Sub AllStocksAnalysisRefactored()


    Dim startTime As Single
    Dim endTime  As Single

Start:
    yearValue = inputbox("What year would you like to run the analysis on?")
    
    'Checking if the year request is valid.
    On Error Resume Next
        If yearValue = "" Then
    
                Exit Sub
        
        ElseIf yearValue = (Worksheets(yearValue) Is Nothing) Then
    
                MsgBox "Sheet Not Found.", vbCritical, "Invalid Input"
                GoTo Start
        
        End If
    On Error GoTo 0
    
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
      Worksheets("All Stocks Analysis").Activate
    
      Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
      Cells(3, 1).Value = "Ticker"
    
      Cells(3, 2).Value = "Total Daily Volume"
    
      Cells(3, 3).Value = "Return"
    

    'Initialize array of all tickers
      Dim tickers(12) As String
    
        tickers(0) = "AY"
        tickers(1) = "CSIQ"
        tickers(2) = "DQ"
        tickers(3) = "ENPH"
        tickers(4) = "FSLR"
        tickers(5) = "HASI"
        tickers(6) = "JKS"
        tickers(7) = "RUN"
        tickers(8) = "SEDG"
        tickers(9) = "SPWR"
        tickers(10) = "TERP"
        tickers(11) = "VSLR"
    
        'Activate data worksheet
         Worksheets(yearValue).Activate
    
        'Get the number of rows to loop over
         RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
        '1a) Create a ticker Index
         Dim tickerIndex As String
        
        
        '1b) Create three output arrays
            Dim tickerStartingPrices(12) As Single
            Dim tickerEndingPrices(12) As Single
            Dim tickerVolumes(12) As Long
    
                '2a) Create a for loop to initialize the tickerVolumes to zero.
                        For i = 0 To 11

                                tickerVolumes(i) = 0
                        
                        Next i
                        '2b) Loop over all the rows in the spreadsheet.
                        
                            i = 0
                            
                            Worksheets(yearValue).Activate
                        
                                For j = 2 To RowCount
            
                        
                                    tickerIndex = tickers(i)
                                    '3a) Increase volume for current ticker
                                        If Cells(j, 1).Value = tickerIndex Then
                                    
                                                tickerVolumes(i) = tickerVolumes(i) + Cells(j, 8).Value
                                        
                                        End If
                                     
                                        '3b) Check if the current row is the first row with the selected tickerIndex.
                                            If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
                                     
                                                        tickerStartingPrices(i) = Cells(j, 6).Value
                                      
                                        '3c) check if the current row is the last row with the selected ticker
                                         ElseIf Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
                                     
                                                tickerEndingPrices(i) = Cells(j, 6).Value
                                                 
                                                 '3d) Increase the tickerIndex.
                                                 i = i + 1
                                                 
                                            End If
                                            
                                     Next j
    
                    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
                       Worksheets("All Stocks Analysis").Activate
                         
                        For i = 0 To 11
                            
                            Cells(4 + i, 1).Value = tickers(i)
            
                            Cells(4 + i, 2).Value = tickerVolumes(i)
            
                            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
                
                        Next i
    
            'Formatting
            Worksheets("All Stocks Analysis").Activate
            Range("A3:C3").Font.FontStyle = "Bold"
            Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
            Range("B4:B15").NumberFormat = "#,##0"
            Range("C4:C15").NumberFormat = "0.0%"
            Columns("B").AutoFit
        
            dataRowStart = 4
            dataRowEnd = 15

                For i = dataRowStart To dataRowEnd
                    
                    If Cells(i, 3) > 0 Then
                        
                        Cells(i, 3).Interior.Color = vbGreen
                        
                    Else
                    
                        Cells(i, 3).Interior.Color = vbRed
                        
                    End If
                    
                Next i
             
                endTime = Timer
                
                MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
            
   End Sub

Sub ClearWork()

    Worksheets("All Stocks Analysis").Activate
    
    Cells.Clear

End Sub


