Attribute VB_Name = "Module3"
Sub yearchange()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        ws.Select
        Call RunCode3
    Next
    Application.ScreenUpdating = True
End Sub
Sub RunCode3()

Dim openingprice As Double
Dim closingprice As Double
Dim yearlychange As Double
Dim curcol As String
Dim nextcol As String
Dim percentchange As Double
tablerow = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    For c = 2 To lastrow
    
        curcol = Cells(c, 1).Value
        nextcol = Cells(c + 1, 1).Value
        
            If curcol <> nextcol Then
            openingprice = Cells(c, 3)
            closingprice = Cells(c, 6)
            yearlychange = closingprice - openingprice
            Range("J" & tablerow).Value = yearlychange
            
                If openingprice = 0 Then
                    percentchange = 0
                    
                    Else: openingprice = Range("C" & c)
                            percentchange = yearlychange / openingprice
                        
                End If
                
                Range("K" & tablerow).Value = percentchange
                
                    If Range("J" & tablerow).Value >= 0 Then
                    Range("J" & tablerow).Interior.ColorIndex = 4
                    
                        Else: Range("J" & tablerow).Interior.ColorIndex = 3
                        
                    End If
                    
            End If
            
    Next c
            
                    
                    
End Sub
