Attribute VB_Name = "Module1"
Sub stock_values()

'define variables
Dim worksheet_count As Integer
Dim i As Double  'record counter
Dim j As Integer ' worksheet counter
Dim list_row As Integer 'display line counter

Dim open_date As Double
Dim open_price As Double
Dim close_date As Double
Dim close_price As Double
Dim stock_volume As Double
Dim headers(8) As String


'assign initial values to worksheet variables
headers(0) = "Ticker"
headers(1) = "Annual Mvmt"
headers(2) = "Percent change"
headers(3) = "Stock Volume"
headers(6) = "ticker"
headers(7) = "value"


'count number of sheets in workbook
worksheet_count = ActiveWorkbook.Worksheets.Count



'navigate through the sheets
For j = 1 To worksheet_count

  'activate current sheet
  Worksheets(j).Activate
   
  'find last record row in sheet
  end_row = Cells(Rows.Count, "a").End(xlUp).Row
 
   
  'assign initial values to sheet variables
  start_row = 2
  list_row = start_row
  open_date = 20200530
  close_date = 0
  stock_volume = 0


  'display titles for ticker output
  Range("j1:q1").Value = headers()
  
  
  'count through all records
   For i = start_row To end_row
    
    'sum stock volume
     stock_volume = Cells(i, 7).Value + stock_volume
           
    'update current open and close date values if less than stored date values
    If Cells(i, 2).Value < open_date Then
         
       open_date = Cells(i, 2).Value
       open_price = Cells(i, 3).Value
         
    End If
        
    If Cells(i, 2).Value > close_date Then
       close_date = Cells(i, 2).Value
       close_price = Cells(i, 6).Value
    End If
   
    'check if last record for current ticker value
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
     
       'if last record is true value then print and format output
       Cells(list_row, 10) = Cells(i, 1)
    
       price_var = close_price - open_price
       Cells(list_row, 11) = price_var
     
       If price_var < 0 Then
          Cells(list_row, 11).Interior.ColorIndex = 3
       Else
          Cells(list_row, 11).Interior.ColorIndex = 4
       End If
        
       'test for zeros in percentage formula
       If price_var = 0 Or open_price = 0 Then
          Cells(list_row, 12) = 0
       Else
          Cells(list_row, 12) = price_var / open_price
       End If
                
       Cells(list_row, 12).NumberFormat = "#0.00%"
       Cells(list_row, 13) = stock_volume
     
       'increment counter
       list_row = list_row + 1
     
       'reset print variables
       open_date = 20200530
       close_date = 0
       stock_volume = 0
                  
    End If
 
   Next i
   
   'reset record and print counters at end of sheet
   list_row = start_row
   i = start_row
 
 
   'create greatest values table
   Range("j:m").Sort Key1:=Range("l1"), Order1:=xlDescending, Header:=xlYes

   Range("o2").Value = "Greatest% increase"
   Range("p2").Value = Range("j2").Value
   Range("q2").Value = Range("l2").Value
   Range("q2").NumberFormat = "#0.00%"
 
   Range("j:m").Sort Key1:=Range("l1"), Order1:=xlAscending, Header:=xlYes
   Range("o3").Value = "Greatest% decrease"
   Range("p3").Value = Range("j2").Value
   Range("q3").Value = Range("l2").Value
   Range("q3").NumberFormat = "#0.00%"
 
   Range("j:m").Sort Key1:=Range("m1"), Order1:=xlDescending, Header:=xlYes
   Range("o4").Value = "Greatest total volume"
   Range("p4").Value = Range("j2").Value
   Range("q4").Value = Range("m2").Value
   
   Range("j:m").Sort Key1:=Range("j1"), Order1:=xlAscending, Header:=xlYes
   Range("j:m", "p:q").Select
   Selection.ColumnWidth = 12.25
   Columns("o").Select
   Selection.ColumnWidth = 19
   Cells(1, 1).Select
   
 
Next j

End Sub
