VB Script for multi year

Sub Calc_Tot_Volume_Year()

Dim row As Double
Dim col As Integer
Dim x As Integer
Dim y As Integer
Dim u As Integer
Dim v As Integer
Dim LastRow As Double
Dim LastColumn As Integer
Dim ActiveSheetNm() As String
Dim Tot_Volume As Double
Dim Ticker_Symbol As String
'Dim ws_name As Worksheet
Dim ws_count As Integer
Dim stk_yr_Open As Double
Dim stk_yr_Close As Double
Dim cnt As Integer
Dim cnt_2 As Integer

Dim grtst_pct_incr As Double
Dim grtst_pct_incr_tkr As String

Dim grtst_pct_dcrs As Double
Dim grts_pct_dcrs_tkr As String

Dim grtst_vol As Double
Dim grts_pct_vol_tkr As String



ws_count = ActiveWorkbook.Worksheets.Count

MsgBox ("Total Worksheets" & Str(ws_count))



         ' Declare Current as a worksheet object variable.
'         Dim Current As Worksheet

         ' Loop through all of the worksheets in the active workbook.
'         For Each Current In Worksheets
'
'            MsgBox Current.Name
'         Next
' --------------------------------------------
' LOOP THROUGH ALL SHEETS
' --------------------------------------------

For Each ws In Worksheets

    ws.Activate
    
    ' Identify last row and last column in the active worksheet
    LastRow = Cells(Rows.Count, 1).End(xlUp).row
    LastColumn = Cells(2, Columns.Count).End(xlToLeft).Column
    
    ' Initialize variables
    ' These two variables are used to read the input table cols A through G
    row = 2
    col = 1
    ' These two variables are used to write the Moderate challange table with four values ticker, price change, %change and volume
    x = 1
    y = LastColumn + 2
    'These two variables are used to write the Hard table with the higest and lowest percent change and the largest volume
    u = 1
    v = y + 5
    cnt = 1
    cnt_2 = 1
    ' These variables are used to store the greatest pct incr/dec and greatest vol as well as the ticker
    grtst_pct_incr = 0
    grtst_pct_dcrs = 0
    grtst_vol = 0
    grtst_pct_incr_tkr = ""
    grtst_pct_dcrs_tkr = ""
    grtst_vol_tkr = ""
    'These variable is used to store the total volume by each ticker symbol
    Tot_Volume = 0

    
 ' Create Header Record for result set  - Moderate
    Cells(x, y).Value = "Ticker Symbol"
    Cells(x, y + 1).Value = "Yearly Change"
    Cells(x, y + 2).Value = "Percentage Change"
    Cells(x, y + 3).Value = "Total Volume"
 ' Create Header Record for result set  - Hard
    Cells(u, v + 1).Value = "Ticker Symbol"
    Cells(u, v + 2).Value = "Value"
    
    'Increment the value of x used for the moderate table so that we can write the values and since in the prior section we create the header(label) record
    x = x + 1


        MsgBox (" Last Row " & Str(LastRow) & "  Last Column " & Str(LastColumn))
                
            ' Loop through the input table
            For row = 2 To LastRow
                    
                    ' This piece of code stores the first open price for a ticker and if the price is 0 it sets it to 1 because of a division by zero calculation further down
                      If cnt = 1 Then
                          stk_yr_Open = Cells(row, 3).Value
                         
                          cnt = cnt + 1
                     End If
        
                    ' Condition to check if the next ticker symbol is the same or different than the current ticker symbol
                    If Cells(row, col).Value <> Cells(row + 1, col).Value Then
                    
                        'MsgBox ("Current Ticker" & Cells(row, col).Value & " Next Ticker " & Cells(row + 1, col).Value)
                        Tot_Volume = Tot_Volume + Cells(row, col + 6)
                        Ticker_Symbol = Cells(row, col).Value
                        
                        stk_yr_Close = Cells(row, 6).Value
                        
                        'Publish values
                        Cells(x, y).Value = Ticker_Symbol
                        Cells(x, y + 1).Value = stk_yr_Close - stk_yr_Open
                        If stk_yr_Open = 0 Then
                                Cells(x, y + 2).Value = 0
                        Else
                                Cells(x, y + 2).Value = ((stk_yr_Close - stk_yr_Open) / stk_yr_Open)
                        End If
                        Cells(x, y + 3).Value = Tot_Volume
                        
                        Cells(x, y + 2).NumberFormat = "0.00%"
                        
                                ' Condition to check if the percentage is greater than color green else color red
                                If (Cells(x, y + 2).Value > 0) Then
                                    Cells(x, y + 2).Interior.ColorIndex = 4
                                ElseIf (Cells(x, y + 2).Value < 0) Then
                                      Cells(x, y + 2).Interior.ColorIndex = 3
                                End If
                        
                        ' Calculating greatest and lowest percent changes and greatest volume
                                If grtst_pct_incr < Cells(x, y + 2).Value Then
                                  grtst_pct_incr = Cells(x, y + 2).Value
                                  grtst_pct_incr_tkr = Ticker_Symbol
                                End If
                                
                                If grtst_pct_dcrs > Cells(x, y + 2).Value Then
                                  grtst_pct_dcrs = Cells(x, y + 2).Value
                                  grtst_pct_dcrs_tkr = Ticker_Symbol
                                End If
                                
                                If grtst_vol < Tot_Volume Then
                                  grtst_vol = Tot_Volume
                                  grtst_vol_tkr = Ticker_Symbol
                                End If

                        
                        'Increment publish row by 1
                        x = x + 1
                        
                        'Re initialize key variables
                        Tot_Volume = 0
                        stk_yr_Open = 0
                        stk_yr_Close = 0
                        Ticker_Symbol = ""
                        cnt = 1
                        
                    Else
                    
                        Tot_Volume = Tot_Volume + Cells(row, col + 6)
                        
                    End If
                      
            Next row ' End of reading a row from the input table
            
            ' Publishing Hard values
                
                Cells(u + 1, v).Value = "Greatest % Increase"
                Cells(u + 1, v + 1).Value = grtst_pct_incr_tkr
                Cells(u + 1, v + 2).Value = grtst_pct_incr
                Cells(u + 1, v + 2).NumberFormat = "0.00%"
                  
                Cells(u + 2, v).Value = "Greatest % Decrease"
                Cells(u + 2, v + 1).Value = grtst_pct_dcrs_tkr
                Cells(u + 2, v + 2).Value = grtst_pct_dcrs
                Cells(u + 2, v + 2).NumberFormat = "0.00%"
                                                
                Cells(u + 3, v).Value = "Greatest Total Volume"
                Cells(u + 3, v + 1).Value = grtst_vol_tkr
                Cells(u + 3, v + 2).Value = grtst_vol

Next ' This will cycle to the next worksheet
End Sub





