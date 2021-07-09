Attribute VB_Name = "Module1"
Sub ticker()

Dim column  As Integer
column = 1
Dim tickername As String
Dim tickerrow As Integer
tickerrow = 2

lastrow = Cells(Rows.Count, 1).End(xlUp).Row


    For a = 2 To lastrow
        
        If Cells(a + 1, column).Value <> Cells(a, column).Value Then
        
            tickername = Cells(a, column).Value
            
            Range("I" & tickerrow).Value = tickername
            
            tickerrow = tickerrow + 1
            
        End If
        
    Next a


End Sub
