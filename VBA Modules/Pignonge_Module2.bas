Attribute VB_Name = "Module2"
Function ColumnNumberToLetter(ByVal ColumnNumber As Integer)
    ColumnNumberToLetter = Replace(Replace(Cells(1, ColumnNumber).Address, "1", ""), "$", "")
End Function


Sub Timeout(seconds As Double)
Starting_Time = Timer

Do

DoEvents
Loop Until (Timer - Starting_Time) >= seconds

End Sub
