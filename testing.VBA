Attribute VB_Name = "Module1"
Sub Stock_Market_Analysis()

' Defining All Variables
' ----------------------------

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
'---------------------------------------------------

For Each ws In Worksheets

    ' Assigning Column Headers
    
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
    last_row = Cells(Rows.Count, 1).End(x1Up).Row
    
        ' Looping For Data On (Yearly Change, Percent Change, And Total Stock Volume) per Ticker
        
        For i = 2 To last_row
        
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
    ' --------------------------------------------------
        
        ' Values For Summary Table
        
        
        






End Sub
