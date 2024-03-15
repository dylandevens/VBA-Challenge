# VBA-Challenge
VBA-Challenge for Module 2 UCB Data Analytics and Visualization

Included is all the files used for the VBA Challenge. 
I have included 3 screenshots to show the final results of all three worksheets in excel with manipulated data.

The script used from VBA is copy and pasted below this. 

I have useda couple sources for troubleshooting some of the (many) problems that came up, including but not limited to course materials for this module, stack exchange, the articles on wallstreetmojo.com, and LLMs to help explain principles in common english and help me understand where to explicitly state a variable's value.







Sub fill_columns()

'All Worksheets to analyze:
  Dim sheets_used As Variant
  sheets_used = Array("2018", "2019", "2020")
  
  Dim worksheets As Variant
    For Each worksheets In sheets_used
        Sheets(worksheets).Activate
      
      
      ' A variable to hold the Ticker symbol (i)
      Dim ticker As String
      Dim ticker_unique As Long
      ticker_unique = 2
      ' Variable for the Yearly Change Column (j)
      Dim yearly_change As Double
      ' Variable to hold the initial stock value at beginning of year to calcuate percent change
      Dim initial_value As Double
      ' Variable for the Percent Change Column (k)
      Dim percent_change As Double
      ' Variable for the Total Stock Volume Column (l)
      Dim total_stock_volume As Double
        total_stock_volume = 0
      Dim last_row As Double
        last_row = Cells(Rows.Count, 1).End(xlUp).Row
        
        
        
        'Variables for Calculating the Greatests(columns Q & R)
        Dim max_increase_ticker As String
        Dim max_decrease_ticker As String
        Dim max_volume_ticker As String
        Dim max_increase As Double
        Dim max_decrease As Double
        Dim max_volume As Double
        Dim last_row_summary As Double
    
    
        
      
      ' Loop through all data
      For i = 2 To last_row
        ticker = (Cells(i, 1).Value)
       ' yearly_change = Cells(i,6
         ' Set the ticker string in column i
        If Cells(ticker_unique, 9).Value = "" Then  ' ***** need to update ticker_unique right before moving to next one *****
            total_stock_volume = 0 'reset t_s_v
            initial_value = 0 'rest and fill i_v
            initial_value = Cells(i, 3).Value
            Cells(ticker_unique, 9).Value = ticker
            Cells(ticker_unique, 10).Value = Cells(i, 3).Value 'fill out initial opening value to calculate yearly change at end
            total_stock_volume = total_stock_volume + Cells(i, 7).Value ' Begin adding Volume summation to column L
        Else
            total_stock_volume = total_stock_volume + Cells(i, 7).Value 'Increment the total_stock_value variable as i increases
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                'Calculate yearly_change
                Cells(ticker_unique, 10).Value = (Cells(i, 6).Value - Cells(ticker_unique, 10).Value)
                'Calculate percent_change
                Cells(ticker_unique, 11).Value = (Cells(ticker_unique, 10).Value / initial_value)
                Cells(ticker_unique, 12).Value = total_stock_volume
                ticker_unique = (ticker_unique + 1) 'increment the next row for ticker_unique
                total_stock_volume = 0
                
                End If
            
            End If
        Next i
        
        'Set the Greatest Variables = 0
        max_increase = 0
        max_decrease = 999
        max_volume = 0
        
        last_row_summary = Cells(Rows.Count, 11).End(xlUp).Row
        
        max_increase = Application.WorksheetFunction.Max(Range("K2:K" & last_row_summary))
        max_decrease = Application.WorksheetFunction.Min(Range("K2:K" & last_row_summary))
        max_volume = Application.WorksheetFunction.Max(Range("L2:L" & last_row_summary))
        
        For i = 2 To last_row_summary
            If Cells(i, 11).Value = max_increase Then
                max_increase_ticker = Cells(i, 9).Value
                Exit For
            End If
        Next i
            
        For i = 2 To last_row_summary
            If Cells(i, 11).Value = max_decrease Then
                max_decrease_ticker = Cells(i, 9).Value
                Exit For
            End If
        Next i
            
        For i = 2 To last_row_summary
            If Cells(i, 12).Value = max_volume Then
                max_volume_ticker = Cells(i, 9).Value
                Exit For
            End If
        Next i
        
        
        Cells(2, 17).Value = max_increase_ticker
        Cells(2, 18).Value = max_increase
        Cells(3, 17).Value = max_decrease_ticker
        Cells(3, 18).Value = max_decrease
        Cells(4, 17).Value = max_volume_ticker
        Cells(4, 18).Value = max_volume
     
        Cells(2, 18).NumberFormat = "0.00%"
        Cells(3, 18).NumberFormat = "0.00%"
        
    Next worksheets
        
    
End Sub
Sub Format()
'
' Format Macro
'

'
    Columns("J:K").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=-1000", Formula2:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=1000", Formula2:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub

