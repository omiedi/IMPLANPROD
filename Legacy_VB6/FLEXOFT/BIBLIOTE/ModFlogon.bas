Attribute VB_Name = "ModFlogon"
Sub APLICACIONSKINS(frm As Form, Optional PictureVerticalBotones As PictureBox)
Exit Sub
End Sub
Function VALOBADA(NOTAB$, NOCAM$, CONDI$, Optional MODOMESS$)
VALOBADA = ""
End Function
Function XVALO(valor) As Double
 XVALO = Val(valor)

End Function
Sub ACTUREGISTRO(TABLA$, CAMPO$, CONDI$, NUEVALOR, REGISTROSAFECTADOS&, Optional MODO$, Optional TIMUT)
   Exit Sub
End Sub

Function SAFETEXT$(TXTXT)
    SAFETEXT$ = ""
    On Error Resume Next
    SAFETEXT$ = TXTXT
    On Error GoTo 0
End Function

