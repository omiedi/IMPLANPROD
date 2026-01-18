Attribute VB_Name = "SCREEN_FUNCTIONS"
Sub LOCATE(A%, B%)
    Form1.CurrentX = B% * Form1.ScaleWidth / 80
    Form1.CurrentY = A% * Form1.ScaleHeight / 25
End Sub
