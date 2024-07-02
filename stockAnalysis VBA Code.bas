Attribute VB_Name = "Module1"
Option Explicit
Sub stockAnalysis():

   'Create and set variables
    Dim ws As Worksheet
    Dim i As Long
    Dim lastrowA As Long
    Dim lastrowK As Long
    Dim lastrowL As Long
    Dim columnI As Long
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalStock As Double

    
    'Loop through each worksheet in this workbook
    For Each ws In ThisWorkbook.Worksheets
    ws.Activate

    
    
    'Set Headers of Summary Tables
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Quarterly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    
    
    'Set Column I Initial Value
    columnI = 2
    
  'First Summary Table
   'Loop through rows in Column A
    lastrowA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastrowA

        'Search for different value in next cell
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then

            'Set opening price
            OpenPrice = ws.Cells(i, 3).Value
            
            'Set Total Stock Volume Inial Amount
             totalStock = Cells(i, 7).Value
            
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Set the Ticker Name
            Ticker = ws.Cells(i, 1).Value
        
            'Print Ticker Name in Summary Table
            ws.Range("I" & columnI).Value = Ticker
            
            'Set Closing Price
            ClosePrice = ws.Cells(i, 6).Value
            
               
            'Calculate Quarterly Change
            quarterlyChange = ClosePrice - OpenPrice
               
            'Print Quarterly Change
            ws.Range("J" & columnI).Value = quarterlyChange
                     
            'Conditional Format Percent Change
            If ws.Range("J" & columnI).Value < 0 Then
                ws.Range("J" & columnI).Interior.ColorIndex = 3
                
            ElseIf ws.Range("J" & columnI).Value > 0 Then
                ws.Range("J" & columnI).Interior.ColorIndex = 4
                
            ElseIf ws.Range("J" & columnI).Value = 0 Then
                ws.Range("J" & columnI).Interior.ColorIndex = 0
            End If
               
            'Calculate Percent Change
            percentChange = ws.Range("J" & columnI).Value / OpenPrice
                
            'Print Percent Change
            ws.Range("K" & columnI).Value = percentChange
            
            'Format Percent Change as Percentage
            ws.Range("K:K").NumberFormat = "0.00%"
              
            'Add to the Total Stock Volume
            totalStock = totalStock + ws.Cells(i, 7).Value
               
            'Print Total Stock Volume to Summary Table
            ws.Range("L" & columnI).Value = totalStock
               
            'Add one to Summary Table Row
            columnI = columnI + 1
               
            'Reset Total Stock Volume
            totalStock = Cells(i, 7).Value
            
        Else
            'Add to the Total Stock Volume
            totalStock = totalStock + ws.Cells(i, 7).Value
 
        End If
        
    Next i

   
   
'Second Summary Table

    'Loop through rows in Column K
    lastrowK = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
    For i = 2 To lastrowK

        'Find Greatest% Increase and Print Ticker & Values
        If ws.Cells(i, 11).Value > ws.Range("Q2").Value Then
            ws.Range("Q2") = ws.Cells(i, 11).Value
            ws.Range("P2") = ws.Cells(i, 9).Value
       
        End If
    Next i
      
    'Loop through rows in Column K
    lastrowK = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
    For i = 2 To lastrowK

        'Find Greatest% Decrease and Print Ticker & Values
        If ws.Cells(i, 11).Value < ws.Range("Q3").Value Then
            ws.Range("Q3") = ws.Cells(i, 11).Value
            ws.Range("P3") = ws.Cells(i, 9).Value

        End If
     Next i

     'Loop through rows in Column L
     lastrowL = ws.Cells(ws.Rows.Count, 12).End(xlUp).Row
     For i = 2 To lastrowL
    
        'Find Greatest Total Volume and Print Ticker & Value
        If ws.Cells(i, 12).Value > ws.Range("Q4").Value Then
            ws.Range("Q4") = ws.Cells(i, 12).Value
            ws.Range("P4") = ws.Cells(i, 9).Value

        End If
     Next i
    

        ws.Columns("I:Q").AutoFit

    Next ws
     
End Sub
