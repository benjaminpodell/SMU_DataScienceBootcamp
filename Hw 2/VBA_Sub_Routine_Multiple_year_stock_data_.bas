Attribute VB_Name = "Module1"
Sub WallStreet_Stocks()
        
        'Runs routine through each worksheet
        For Each WS In Worksheets
        
        'Creates variables that are declared in memory
        Dim stock_name As String
        Dim stock_open As Double
        Dim stock_close As Double
        Dim year_change As Double
        Dim percent_change As Double
        Dim stock_volume As Double
        Dim table_row As Double
        Dim i As Long
        
        'Creates a counter for summing the Total Stock Volume
        stock_volume = 0
        'Creates a starting point for looping through
        table_row = 2
        'Defining Stock Open Price for use in later for loop
        stock_open = WS.Cells(2, 3).Value
        'Defining the last row for Initial Stock Ticker
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        'Creates Titles for respective column cells
        WS.Cells(1, "I").Value = "Ticker"
        WS.Cells(1, "J").Value = "Yearly Change"
        WS.Cells(1, "K").Value = "Percent Change"
        WS.Cells(1, "L").Value = "Total Stock Volume"
        WS.Cells(2, "O").Value = "Greatest % Increase"
        WS.Cells(3, "O").Value = "Greatest % Decrease"
        WS.Cells(4, "O").Value = "Greatest Total Volume"
        WS.Cells(1, "P").Value = "Ticker"
        WS.Cells(1, "Q").Value = "Value"
        
        'For loop creating summation of Total Stock Volume and Yearly Change across all stocks
        For i = 2 To LastRow
            'Conditional Loop that moves down rows unless each ticker is the same per row
            If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
                'Adds Total Stock Volume using counter by starting at 0 and adding respective cells until next ticker is descovered and defines location of Ticker
                stock_volume = stock_volume + WS.Cells(i, 7).Value
                WS.Cells(table_row, 12).Value = stock_volume
                'Loops throgh each ticker and places ticker in designated location
                stock_name = WS.Cells(i, 1).Value
                WS.Cells(table_row, 9).Value = stock_name
                'Defining Stock Closing Price that loops through rows and helps remove fullstack error and divide by 0 error for percentage
                stock_close = WS.Cells(i, 6).Value
                'Formula that calculates stocks yearly change and places it in designated location
                year_change = stock_close - stock_open
                WS.Cells(table_row, 10).Value = year_change
               
                    'Conditional to prevent percentage formula from dividing by 0 or getting fullstack error
                    If (stock_close = 0 And stock_open = 0) Then
                        percent_change = 0
                    ElseIf (stock_close <> 0 And stock_open = 0) Then
                        percent_change = 1
                    Else
                        'Formula to calculate percent
                        percent_change = year_change / stock_open
                        'Excel formula to convert percentage value into an actual percentage along with location for Percent
                        WS.Cells(table_row, 11).NumberFormat = "0.00%"
                        WS.Cells(table_row, 11).Value = percent_change
                    End If
   
                'Pushes loop to move to next row for formulas using table increments
                table_row = table_row + 1
                'Creates next row for stock open to calculate for
                stock_open = WS.Cells(i + 1, 3)
                'Resets counter for Total Stock Volume for next row
                stock_volume = 0
            'If stock ticker is the same for next row
            Else
                'This adds to the current ticker for it's stock volume
                stock_volume = stock_volume + WS.Cells(i, 7).Value
            End If
            
        Next i
    
        'Defining the last row for Percent Change Column
        lastrow2 = WS.Cells(Rows.Count, 10).End(xlUp).Row
        'Another for loop that will color cells in Percent Change either green for positive values or red for negative values
        For j = 2 To lastrow2
            'Conditional to check if The Yearly Change is a Positive Value or Zero and if so colors interior of cell Light Green
            If (WS.Cells(j, 10).Value > 0 Or WS.Cells(j, 10).Value = 0) Then
                WS.Cells(j, 10).Interior.ColorIndex = 4
            'Checks to see if the value of the cell is less than 0 and if so colors the interior of the cell red
            ElseIf WS.Cells(j, 10).Value < 0 Then
                WS.Cells(j, 10).Interior.ColorIndex = 3
            
            End If
        Next j
        
        '[Hard]
        'Does not work for Ticker , Percent Decrease, and Greatest Total Volume...I am not sure why?
        'Defining the last row for Stock Ticker
        lastrow3 = WS.Cells(Rows.Count, 9).End(xlUp).Row
        'New for loop to last row for /stock Ticker
        For k = 2 To lastrow3
        
        'Conditional that uses a function to calculate the Maximum Percentage Value and places it in designated cell
        'https://docs.microsoft.com/en-us/office/vba/api/excel.worksheetfunction

            If WS.Cells(k, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & lastrow3)) Then
            WS.Cells(2, 16).Value = WS.Cells(k, 9).Value
            WS.Cells(2, 17).Value = WS.Cells(k, 11).Value
            'Formats cell value as a percentage with two decimals
            WS.Cells(2, 17).NumberFormat = "0.00%"

         'Conditional that uses an excel function object to calculate the Minimum Percentage Value and places it in designated cell
          
            
            ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & lastrow3)) Then
            WS.Cells(3, 16).Value = WS.Cells(k, 9).Value
            WS.Cells(3, 17).Value = WS.Cells(k, 11).Value
            WS.Cells(3, 17).NumberFormat = "0.00%"
           
        'Conditional that uses an excel function object to calculate the Maximum Total Stock Volume and places it in designated cell
           
            ElseIf Cells(k, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & lastrow3)) Then
            WS.Cells(4, 16).Value = WS.Cells(k, 9).Value
            WS.Cells(4, 17).Value = WS.Cells(k, 12).Value
            End If

        Next k

    'Moves sub routine to next worksheet
    Next WS
        
End Sub


