Attribute VB_Name = "Module2"
Sub total()

Dim sumcount As Integer
Dim total As Variant
Dim curcol As String
Dim nextcol As String
sumcount = 2
voltotal = 0
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    For a = 2 To lastrow
    
        curcolvalue = Cells(a, 1).Value
        nextcolvalue = Cells(a + 1, 1).Value
        
            If curcolvalue <> nextcolvalue Then
            Cells(sumcount, 9).Value = curcolvalue
            voltotal = voltotal + Cells(a, 7).Value
            Cells(sumcount, 12).Value = voltotal
            sumcount = sumcount + 1
            voltotal = 0
            
            Else: voltotal = voltotal + Cells(a, 7).Value
        End If
        
    Next a
    

End Sub
