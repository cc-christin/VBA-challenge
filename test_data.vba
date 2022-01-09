Attribute VB_Name = "Module1"
Sub Stock_Market_Analysis()

' Test Data

' Defining All Variables (Ticker, Yearly Change, Percent Change, Total Stock Volume)

' Defining A Variable For Ticker
Dim Ticker As String

' Defining A Variable For year open
Dim year_open As Double

' Defining A Variable For year close
Dim year_close As Double

' Defining A Variable For Yearly Change
Dim Yearly_Change As Double

' Defining A Variable For Percent Change
Dim Percent_Change As Double

' Defining A Variable For Total Stock Volume
Dim Total_Stock_Volume As Double

' Defining A Variable Set Up The Starting Row
Dim starting_row As Integer

' Defining Variable To Excute Code In All Worksheets At Once For The Workbook
Dim ws As Worksheet

' Initiate In All Worksheet Excuting Code At Once

For Each ws In Worksheets

    ' Assigning Column Headers (Ticker, Yearly Change, Percent Change, Total Stock Volume)
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Assinging Starting Intiger For Loop
    starting_row = 2
    
    previous_i = 1
    
    ' Assinging Inital Variable For Total Stock Volume Per Ticker
    
    Total_Stock_Volume = 0
    
    ' Setting Last Row For Column A, 1
    EndRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
        ' Looping For Data On (Yearly Change, Percent Change, And Total Stock Volume) per Ticker
        
        For i = 2 To EndRow
        
            ' If Ticker Not Equal To Previous, Excute Code
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ' Retrive Ticker
            
            Ticker = ws.Cells(i, 1).Value
        
            ' Moving On To The Next Ticker
            
            previous_i = previous_i + 1
        
            ' Calculating Yearly Change from year open (column C, 3) and year close (column F, 6)
        
            year_open = ws.Cells(previous_i, 3).Value
            year_close = ws.Cells(i, 6).Value
        
            ' Loop For Sum Of Total Stock Volume
        
            For j = previous_i To i
        
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value
            
            Next j
        
            ' Loop Value Goes To Zero
        
            If year_open = 0 Then
        
                Percent_Change = year_close
            
            Else
        
                ' Yearly Change Condition
                Yearly_Change = year_close - year_open
            
                ' Percentage Change Condition
                Percent_Change = Yearly_Change / year_open
            
            End If
        
            ' Values For Summary Table For Starting Row
        
            ws.Cells(starting_row, 9).Value = Ticker
            ws.Cells(starting_row, 10).Value = Yearly_Change
            ws.Cells(starting_row, 11).Value = Percent_Change
        
            ' Percent Formating
        
            ws.Cells(starting_row, 11).NumberFormat = "0.00%"
        
            ' Total Stock Volume
        
            ws.Cells(starting_row, 12).Value = Total_Stock_Volume
        
            ' Next Row
        
            starting_row = starting_row + 1
        
            ' Reset Variables To Zero
        
            Total_Stock_Volume = 0
            Yearly_Change = 0
            Percent_Change = 0
        
            ' Setting i to a variable, previous_i
        
            previous_i = i
        
        
        End If
        
    ' Loop Completed
    
    Next i
    
' Bonus Summary Table (Hard Solution)
    
    ' Last Row of Column K
        
    kEndRow = ws.Cells(Rows.Count, "K").End(xlUp).Row
        
    ' Intial variables for bonus summary table (compare to hard_solution.png)
    Increase = 0
    Decrease = 0
    Greatest = 0
        
        ' calculate percent change
        For k = 3 To kEndRow
            
            ' last k
            last_k = k - 1
                
            ' current k
            current_k = ws.Cells(last_k, 11).Value
                
            ' previous k
            previous_k = ws.Cells(last_k, 12).Value
                
            ' previous greatest volume
            prevous_vol = ws.Cells(last_k, 12).Value
                
    
            ' Greatest Increase
            If Increase > current_k And Increase > prevous_k Then
                
                Increase = Increase
                    
            ElseIf current_k > Increase And current_k > prevous_k Then
                
                Increase = current_k
                    
                increase_name = ws.Cells(k, 9).Value
                    
            ElseIf prevous_k > Increase And prevous_k > current_k Then
                
                Increase = prevous_k
                    
                increase_name = ws.Cells(last_k, 9).Value
                    
            End If
                
            
            ' Greatest Decrease
                
            If Decrease < current_k And Decrease < prevous_k Then
            
                Decrease = Decrease
                
            ElseIf current_k < Increase And current_k < prevous_k Then
                
                Decrease = current_k
                    
                decrease_name = ws.Cells(k, 9).Value
                    
            ElseIf prevous_k < Increase And prevous_k < current_k Then
                
                Decrease = previous_k
                    
                decrease_name = ws.Cells(last_k, 9).Value
                    
            End If
                
            
            ' Find Greatest Total Volume
                
            If Greatest > volume And Greatest > prevous_vol Then
                
                Greatest = Greatest
                    
            ElseIf volume > Greatest And volume > prevous_vol Then
                
                Greatest = volume
                    
                greatest_name = ws.Cells(last_k, 9).Value
                    
            End If
                
        Next k
            
        
  ' (Bonus) Ranges for greatest % increase, greatest % decrease, greatest total volume
            
    ws.Range("N1").Value = "Column Name"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("O2").Value = increase_name
    ws.Range("O3").Value = decrease_name
    ws.Range("O4").Value = greatest_name
    ws.Range("P2").Value = Increase
    ws.Range("P3").Value = Decrease
    ws.Range("P4").Value = Greatest
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"
    
' conditional formating

    jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
    
        For j = 2 To jEndRow
        
        ' conditional for green
            If ws.Cells(j, 10) > 0 Then
            
                ws.Cells(j, 10).Interior.ColorIndex = 4
                
            Else
            
                ' red
                ws.Cells(j, 10).Interior.ColorIndex = 3
            
            End If
            
        Next j
        
        ' autofit columns
        
            Columns("I:Q").AutoFit
        
' Go To Next Worksheet
Next ws
 
End Sub







