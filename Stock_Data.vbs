Attribute VB_Name = "Stock-Data"
Sub Stock_Data()

For Each ws In Worksheets

'Define variables

Dim i As Long
Dim j As Long
Dim Summary_Table_Row As Long
Dim LastRow As Long
Dim ticker As String
Dim total_stock As Double
Dim Opening_Price As Variant
Dim Closing_Price As Variant
Dim yearly_change As Variant
Dim percent_change As Variant


'Determine Last Row of the table
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Setting titles for Summary table
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 17).Value = "Greatest Increase %"
ws.Cells(3, 17).Value = "Greatest Decrease %"
ws.Cells(4, 17).Value = "Greatest Total Volume"
ws.Cells(1, 18).Value = "Ticker"
ws.Cells(1, 19).Value = "Value"

'Setting initial values for some variables
total_stock = 0
yearly_change = 0
percent_change = 0
Summary_Table_Row = 2

'Loop
For i = 2 To LastRow
    ' Determine Opening Price
   If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then

       Opening_Price = ws.Cells(i, 3).Value
   End If
    
    'Determine Closing Price, Yearly Change , Percentage Change and Total Stock volume
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        ticker = ws.Cells(i, 1).Value
        Closing_Price = ws.Cells(i, 6).Value
        yearly_change = Closing_Price - Opening_Price
        percent_change = yearly_change / Opening_Price
        total_stock = total_stock + ws.Cells(i, 7).Value
        
       'Putting values in Summary table
       ws.Range("I" & Summary_Table_Row).Value = ticker
       ws.Range("J" & Summary_Table_Row).Value = yearly_change
       ws.Range("K" & Summary_Table_Row).Value = percent_change
       ws.Range("L" & Summary_Table_Row).Value = total_stock
                
                
       Summary_Table_Row = Summary_Table_Row + 1
                
      'Resetting values to start loop again
       yearly_change = 0
       percent_change = 0
       total_stock = 0
                
        Else
               
        total_stock = total_stock + ws.Cells(i, 7).Value
                
        End If
        
        Next i
'Conditional Formatting for Yearly change values also Percent Change

For j = 2 To Summary_Table_Row - 1

 'Format Percent Change
 
  ws.Cells(j, 11).Value = FormatPercent(ws.Cells(j, 11).Value, 2)
            
    If (ws.Cells(j, 10).Value) >= 0 Then
                    
        ws.Cells(j, 10).Interior.ColorIndex = 4
                            
    Else
                            
        ws.Cells(j, 10).Interior.ColorIndex = 3
                        
    End If

'Determine values in the bonus question


 ws.Cells(2, 19).Value = FormatPercent(Application.WorksheetFunction.Max(ws.Range("K2" & ":" & "K" & Summary_Table_Row)), 2)
                    
            If ws.Cells(j, 11).Value = ws.Cells(2, 19).Value Then
                    
            ws.Cells(2, 18).Value = ws.Cells(j, 9).Value
                    
            End If
            
            
            
            If ws.Cells(j, 11).Value < 0 Then
                            
            ws.Cells(3, 19).Value = FormatPercent(Application.WorksheetFunction.Min(ws.Range("K2" & ":" & "K" & Summary_Table_Row)), 2)
                           
            End If
                        
            If ws.Cells(j, 11).Value = ws.Cells(3, 19).Value Then
                        
            ws.Cells(3, 18).Value = ws.Cells(j, 9).Value
                        
            End If
            
            
            
            ws.Cells(4, 19).Value = Application.WorksheetFunction.Max(ws.Range("L2" & ":" & "L" & Summary_Table_Row))
                            
            If ws.Cells(j, 12).Value = ws.Cells(4, 19).Value Then
                            
            ws.Cells(4, 18).Value = ws.Cells(j, 9).Value
                            
            End If
            
                            
            ws.Range("S4").EntireRow.AutoFit
            
        Next j
        
        ws.Range("J2" & ":" & "J" & Summary_Table_Row).NumberFormat = "[$$-en-US]#,##0.00"
        ws.Range("I1:S1").EntireColumn.AutoFit

Next ws


End Sub


