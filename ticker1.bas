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
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

Dim column  As Integer
column = 1
Dim tickername As String
Dim tickerrow As Integer
tickerrow = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).Row


    For a = 2 To lastrow
    
        curtick = Cells(a, column).Value
        nexttick = Cells(a + 1, column).Value
        
        If curtick <> nexttick Then
        
            tickername = Cells(a, column).Value
            
            Range("I" & tickerrow).Value = tickername
            
            tickerrow = tickerrow + 1
            
        End If
        
    Next a
End Sub
