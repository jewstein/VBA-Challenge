Attribute VB_Name = "Module1"
Sub ticker()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        ws.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
End Sub
Sub RunCode()
    Cells(1, 9).Value = "Ticker"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Value"

Dim column  As Integer
column = 1
Dim tickername As String
Dim tickerrow As Integer
tickerrow = 2
Dim sumcount As Integer
Dim voltotal As Variant
Dim curcol As String
Dim nextcol As String
sumcount = 2
voltotal = 0
Dim openingprice As Double
Dim closingprice As Double
Dim yearlychange As Double
Dim percentchange As Double
tablerow = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).Row


    For a = 2 To lastrow
    
        curcol = Cells(a, column).Value
        nextcol = Cells(a + 1, column).Value
        
        If curcol <> nextcol Then
        
        ' find and assign ticker
        
            tickername = Cells(a, column).Value
            
            Range("I" & tickerrow).Value = tickername
            
            tickerrow = tickerrow + 1
            
        ' find and assign stock volume
            
            Cells(sumcount, 9).Value = curcol
            
            voltotal = voltotal + Cells(a, 7).Value
            
            Cells(sumcount, 12).Value = voltotal
            
            sumcount = sumcount + 1
            
            voltotal = 0
            
            Else: voltotal = voltotal + Cells(a, 7).Value
            
        End If
        
    Next a
    
    For b = 2 To lastrow
    
        curcol = Cells(b, 1).Value
        nextcol = Cells(b + 1, 1).Value
        
            If curcol <> nextcol Then
            openingprice = Cells(b, 3)
            closingprice = Cells(b, 6)
            yearlychange = closingprice - openingprice
            Range("J" & tablerow).Value = yearlychange
            
                If openingprice = 0 Then
                    percentchange = 0
                    
                    Else: openingprice = Range("C" & b)
                            percentchange = yearlychange / openingprice
                        
                End If
                
                Range("K" & tablerow).Value = percentchange
                
                    If Range("J" & tablerow).Value >= 0 Then
                    Range("J" & tablerow).Interior.ColorIndex = 4
                    
                        Else: Range("J" & tablerow).Interior.ColorIndex = 3
                        
                    End If
                    tablerow = tablerow + 1
            End If
            yearlychange = 0
            percentchange = 0
    Next b
    
    
End Sub
