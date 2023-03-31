Attribute VB_Name = "Module1"
Public Sub MacroStockMkt() 'START PROGRAM

'SET MAIN WORKSHEET AS WORKSHEET OBJECT VARIABLE
Dim ws As Worksheet
Dim wb As Workbook
Dim headers() As Variant
Dim i As Integer

'SET THE CURRENT WORKBOOK TO ALIAS FOR EASE OF COMMANDS
Set wb = ActiveWorkbook

headers() = Array("Ticker", "Date", "Open", "High", "Low", "Close", "Volume", " ", "Ticker", "Yearly_Change", "Percent_Change", "Stock_Volum", " ", " ", " ", "Ticker", "Value")

'For loop through worksheets in workbook for header setup
For Each ws In wb.Sheets
    With ws 'with loop allows disregarding alia for worksheet shared between commands
         .Rows(1).Value = " " ' clears current cell
    
    For i = LBound(headers()) To UBound(headers()) ' for loop through headers array
        .Cells(1, 1 + i).Value = headers(i) 'ADD HEADER TITLE FOR EACH COLUMN ON SHEET
    
    Next i 'for loop step into next header index
        .Rows(1).Font.Bold = True 'format header as bold
        .Rows(1).VerticalAlignment = xlCenter 'format header as center justification
    
    End With 'for loop step into next header index

Next ws 'for loop through worksheets in workbook for stock calculations

'*******************************************************************************************************
'for loop through worksheets in workbook for stock calculations
For Each ws In Worksheets

    'set initial variables for calculations
        Dim Ticker_Name As String
        Ticker_Name = " "
        Dim Total_Ticker_Volume As Double
        Total_Ticker_Volume = 0
        Dim Beg_Price As Double
        Beg_Price = 0
        Dim End_Price As Double
        End_Price = 0
        Dim Yearly_Price_Change As Double
        Yearly_Price_Change = 0
        Dim Max_Ticker_Name As String
        Max_Ticker_Name = " "
        Dim Min_Ticker_Name As String
        Min_Ticker_Name = " "
        Dim Max_Percent As Double
        Max_Percent = 0
        Dim Min_Percent As Double
        Min_Percent = 0
        Dim Max_Volume_Ticker_Name As String
        Max_Volume_Ticker_Name = " "
        Dim Max_Volume As Double
        Max_Volume = 0
    '*************************************************************************
    'Set location for Variable
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    
    'Set Row count for all sheets in the workbook
    Dim LastRow As Long
    
    'loop through all sheets to find last cell that isnt empty
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'set initial value of beginning stock value for the first ticker of ws
    Beg_Price = ws.Cells(2, 3).Value
    
    'loop from the beginning of the main worksheet row 2 until last row of the last worksheet o is just a place holder for the value of the cell row number and counts all the way to the end of the last row
    For o = 2 To LastRow

        'check confirmation on same ticker name
        If ws.Cells(o + 1, 1) <> ws.Cells(o, 1).Value Then
                      
            'set ticker name starting point
            Ticker_Name = ws.Cells(o, 1).Value
             
            'calculate yearly price change
            End_Price = ws.Cells(o, 6).Value
            Yearly_Price_Change = End_Price - Beg_Price
             
            'set condition for a zero value
            If Beg_Price <> 0 Then
               Yearly_Price_change_Percent = (Yearly_Price_Change / Beg_Price) * 100
                 
            End If 'if statement termination
             
            'calcuate the ticker name total volume
            Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(o, 7).Value
             
            'print ticker name in summary table, column I
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
             
            'print the yearly price change in the summary table in colume J
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Price_Change
             
            'Change color format if negative number make red if positive number make green
            If (Yearly_Price_Change > 0) Then
                 ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
             
            ElseIf (Yearly_Price_Change <= 0) Then
                 ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
             
            End If 'if statement termination
             
            'calculate the yearly price change as percent in the summary table in column K
            ws.Range("K" & Summary_Table_Row).Value = (CStr(Yearly_Price_change_Percent) & "%")
             
             'calculate the total volume in the summary table put in column L
             ws.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
             
             'add 1 to the summary table row count
             Summary_Table_Row = Summary_Table_Row + 1
             
             'get the beginning price
             Beg_Price = ws.Cells(o + 1, 3).Value
             
             'calculate worksheet percentage changes
              If (Yearly_Price_change_Percent > Max_Percent) Then
                     Max_Percent = Yearly_Price_change_Percent
                     Max_Ticker_Name = Ticker_Name
                     
              ElseIf (Yearly_Price_change_Percent < Min_Percent) Then
                     Min_Percent = Yearly_Price_change_Percent
                     Min_Ticker_Name = Ticker_Name
              End If 'if statement termination
                          
              If (Total_Ticker_Volume > Max_Volume) Then
                     Max_Volume = Total_Ticker_Volume
                     Max_Volume_Ticker_Name = Ticker_Name
              End If ' If Statement Termination
             
             'Reset Values
              Yearly_Price_change_Percent = 0
              Total_Ticker_Volume = 0
             
        'else if in next ticker name then enter new ticker stock volume
         Else
              Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(o, 7).Value
                 
         End If ' If Statement Termination
                
                
    Next o ' for loop step into next row index
                
            'Input values in assigned cells
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P2").Value = Max_Ticker_Name
            ws.Range("P3").Value = Min_Ticker_Name
            ws.Range("P4").Value = Max_Volume_Ticker_Name
            ws.Range("Q2").Value = (CStr(Max_Percent) & "%")
            ws.Range("Q3").Value = (CStr(Min_Percent) & "%")
            ws.Range("Q4").Value = Max_Volume
                
            ' autofit all column within worksheet for presentation
            ws.Cells.Columns.AutoFit

Next ws ' for loop step into next worksheet index
    
    
End Sub ' end program
