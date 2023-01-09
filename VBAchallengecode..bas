Attribute VB_Name = "Module1"
Sub marketstats()

'declare variables

Dim lra As Long
Dim lrk As Long
Dim lrj As Long
Dim ws As Integer
Dim ticker As String
Dim tvolume As Double
Dim oprice As Double
Dim cprice As Double
Dim pchange As Double
Dim pchange2 As Double
Dim j As Integer
Dim n As Integer
Dim rng As Range
Dim cell As Range

ws = ActiveWorkbook.Worksheets.Count

'loop through all worksheets

For n = 1 To ws

    Sheets(n).Activate
    
    'find last rows

    lra = Cells(Rows.Count, 1).End(xlUp).Row
    lrj = Cells(Rows.Count, 10).End(xlUp).Row
    lrk = Cells(Rows.Count, 11).End(xlUp).Row
    
    'fill in static cell values

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Largest Total Volume"
    
    'counter variable value

    j = 2
    
    'loop through all data, storing necessary variables

    For i = 2 To (lra + 1) Step 1
    
        'if ticker is not equal to checked cell, it is either the first loop or the ticker has changed. Save and print variables to respective cells as necessary

        If ticker <> Cells(i, 1) Then
        
            If ticker = "" Then
        
                ticker = Cells(2, 1).Value
                oprice = Cells(2, 3).Value
            
            Else
        
                cprice = Cells((i - 1), 6)
                
                'yearly change and % change
        
                Cells(j, 10).Value = (cprice - oprice)
                Cells(j, 11).Value = ((cprice - oprice) / oprice)
                
                'ticker and tvolume of previous ticker
                
                Cells(j, 9).Value = ticker
                Cells(j, 12).Value = tvolume
                
                'set to new ticker and restart tvolume count
        
                ticker = Cells(i, 1)
                tvolume = Cells(i, 7)
                
                'save oprice on first day of year for new ticker
                
                oprice = Cells(i, 3).Value
        
        
                j = (j + 1)
        
            End If
          
        'ticker hasn't change, add up total volume
          
        Else
       
            tvolume = (tvolume + Cells(i, 7))
        
        End If
        
    Next i
    
    'reset variables, another loop going through newly created column to find largest % increase, % decrease, total volume
    
    pchange = 0
    pchange2 = 0
    tvolume = 0
    
    For i = 2 To (lrk + 1) Step 1
    
        If Cells(i, 11).Value > pchange Then
    
            pchange = Cells(i, 11).Value
            Cells(2, 16).Value = Cells(i, 9).Value
            Cells(2, 17).Value = pchange
        
        ElseIf Cells(i, 11).Value < pchange2 Then
    
            pchange2 = Cells(i, 11).Value
            Cells(3, 16).Value = Cells(i, 9).Value
            Cells(3, 17).Value = pchange2
        
        ElseIf Cells(i, 12).Value > tvolume Then
    
            tvolume = Cells(i, 12).Value
            Cells(4, 16).Value = Cells(i, 9).Value
            Cells(4, 17).Value = tvolume
    
        End If
    
    Next i

    'set range to Yearly Change column, change interior color based on value

    Set rng = Range("J:J")

    For Each cell In rng
    
        If cell.Value = "Yearly Change" Then
            cell.Interior.Color = vbWhite
        
        ElseIf cell.Value > 0 Then
            cell.Interior.Color = vbGreen
        
        ElseIf cell.Value < 0 Then
            cell.Interior.Color = vbRed
    
        End If

    Next cell
        
    'set column K & Q cells to percentage format
        
    Range("K:K,Q2:Q3").NumberFormat = "0.00%"

    'autofit data for whole sheet
    
    ActiveSheet.UsedRange.EntireColumn.AutoFit
    ActiveSheet.UsedRange.EntireRow.AutoFit
    
Next n
    
End Sub
