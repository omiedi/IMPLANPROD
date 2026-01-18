Attribute VB_Name = "OBJPLA1"


Public LABE$(64), TIVAA%(64), RENG%(64), COLU%(64), Form$(64), VDS$(64)
Public AVAL$(64), POED%(64), RGEC%(64), CLEC%(64), LGEC%(64), PXGR%(64), LGGR%(64), MALEN%(64)
Public OKEY%
Public KPLA%, LOREC%
Public IANT%
'

Sub PREPLA(OBY$, VDEF$, TTY$)
       '
       Dim AECO$(32)
       LBX$ = LUBAS$
       '
       OBJNAME$ = AJUSTI$(UCase$(OBY$), 16)
       FORMUPLA.OBNAME = OBJNAME$
       '
       Dim X1 As String * 256
       n1% = FreeFile
       Open LBX$ + "OBJNAME.DRV" For Random Access Read Shared As #n1% Len = 256
       '
       For i& = 1 To LOF(n1%) / 256
           Get #n1%, i&, X1$
           If Left$(X1$, 16) = OBJNAME$ Then GoTo 10100
       Next i&
       '
       LBX1$ = REPLACAR$(LBX$, "\", "/")
       Call COMUNI("Error de Invocación de Rutina 'OBJPLA':\  \No Existe un Objeto denominado '" + TRIM$(OBJNAME$) + "'\en Carpeta '" + LBX1$ + ".")
       Close #n1%: n1% = 0
       Call FINAL("")
       '
10100  X$ = X1$
       Close #n1%: n1% = 0
       OBJDESC$ = Mid$(X$, 17, 48)
       OBJTYPE% = CVI(Mid$(X$, 65, 2))
       If OBJTYPE% <> 2 Then
            Call COMUNI("ERROR FATAL: El Objeto\" + TRIM$(OBJNAME$) + " no es\Caja Entrada de Datos.")
            Call FINAL("")
       End If
       OBJPUNT& = CVS(Mid$(X$, 67, 4))
       '
       BSUP% = Asc(Mid$(X$, 71, 1))
       BIZQ% = Asc(Mid$(X$, 72, 1))
       ALTU% = Asc(Mid$(X$, 73, 1))
       ALIB% = Asc(Mid$(X$, 74, 1))
       '
       For UU% = 1 To 32
         AECO$(UU%) = ""
         ARGATRI$ = "ARCHECO(" + TRIM$(Str$(UU%)) + ")="
         PX% = InStr(UCase$(TTY$), ARGATRI$)
         If PX% > 0 Then
            XX$ = Mid$(TTY$, PX% + Len(ARGATRI$))
            PY% = InStr(XX$, "/"): If PY% < 1 Then PY% = 1 + Len(XX$)
            XX$ = Left$(XX$, PY% - 1)
            If TRIM$(XX$) <> "" Then AECO$(UU%) = XX$
            TTY$ = Left$(TTY$, PX% - 1) + Mid$(TTY$, PX% + Len(ARGATRI$) + Len(XX$))
         End If
       Next UU%
       '
       TTZ$ = REPLACAR$(TTY$, "/", "")
       TITU$ = TTZ$
       If TRIM$(TITU$) = "" Then TITU$ = Mid$(X$, 75, 32)
       If TRIM$(TITU$) = "" Then TITU$ = OBJDESC$
       If TRIM$(TITU$) = "" Then TITU$ = OBJNAME$
       COLO$ = Mid$(X$, 107, 3): If MONICOL% <> 1 Then COLO$ = "B-N"
       CIER% = Asc(Mid$(X$, 110, 1)): CIERS$ = "NO"
          If CIER% = 1 Then CIERS$ = "SI"
       CONF% = Asc(Mid$(X$, 111, 1)): CONFS$ = "NO"
          If CONF% = 1 Then CONFS$ = "SI"
          '
10200  PUNTE& = OBJPUNT&: KPLA% = 0: PRICA% = 0
       '
       Dim X2 As String * 128
       n1% = FreeFile
       Open LBX$ + "OBJDETA.DRV" For Random Access Read Shared As #n1% Len = 128
       '
       LOREC% = Len(VDEF$)
       '
       LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
       If LONCODI% < 15 Or LONCODI% > 24 Then LONCODI% = 15
       FORMACO$ = "C" + TRIM$(Str$(LONCODI%))
       '
       For U& = 1 To LOF(n1%) / 128
         Get #n1%, U&, X2$
         If UCase$(Left$(X2$, 16)) = OBJNAME$ Then
           X$ = Mid$(X2$, 17)
           KPLA% = KPLA% + 1: If KPLA% > 64 Then GoTo 10300
           '
           PUNTE& = CVS(Left$(X$, 4))
           '
           LABE$(KPLA%) = REPLACEN$(TRIM$(Mid$(X$, 5, 30)))
           TIVAA%(KPLA%) = Asc(Mid$(X$, 35, 1))
           RENG%(KPLA%) = Asc(Mid$(X$, 36, 1))
           COLU%(KPLA%) = Asc(Mid$(X$, 37, 1))
           Form$(KPLA%) = TRIM$(Mid$(X$, 38, 6))
              If TIVAA%(KPLA%) = 7 Then Form$(KPLA%) = FORMACO$
           MALEN%(KPLA%) = Int(XVALO(Mid$(Form$(KPLA%), 2)))
           PPXXYY% = InStr(Form$(KPLA%), "/")
           If PPXXYY% > 0 Then MALEN%(KPLA%) = Int(XVALO(Mid$(Form$(KPLA%), PPXXYY% + 1)))
           '
           AVAL$(KPLA%) = TRIM$(Mid$(X$, 44, 8))
           If AECO$(KPLA%) <> "" Then AVAL$(KPLA%) = AECO$(KPLA%)
           POED%(KPLA%) = Asc(Mid$(X$, 52, 1))
              If POED%(KPLA%) = 1 Then If PRICA% = 0 Then PRICA% = KPLA%
           RGEC%(KPLA%) = Asc(Mid$(X$, 53, 1))
           CLEC%(KPLA%) = Asc(Mid$(X$, 54, 1))
           LGEC%(KPLA%) = Asc(Mid$(X$, 55, 1))
           PXGR%(KPLA%) = CVI(Mid$(X$, 56, 2))
           LGGR%(KPLA%) = CVI(Mid$(X$, 58, 2))
           '
           If LOREC% < PXGR%(KPLA%) + LGGR%(KPLA%) - 1 Then LOREC% = PXGR%(KPLA%) + LGGR%(KPLA%) - 1
           '
         End If
       Next U&
       Close #n1%: n1% = 0
       '
10300  ' Call SETCOLOR(COLO$)
       ' If Mid$(UCase$(COLO$), 3, 1) = "B" Then COLOR 15
       '
       For II% = 0 To 31
          FORMUPLA.ENTRA(II%).Visible = False
          FORMUPLA.ECO(II%).Visible = False
          FORMUPLA.Label(II%).Visible = False
          FORMUPLA.BOTSEL(II%).Visible = False
       Next II%
       '
       FORMUPLA.REFLABEL.Caption = Space$(ALIB%)
       COEF = FORMUPLA.REFLABEL.Width / ALIB%
       COEFV = Screen.Height / 25
       FORMUPLA.Caption = TITU$
       MATOP = FORMUPLA.Height - FORMUPLA.ScaleHeight
       '
       FORMUPLA.Top = BSUP% * COEFV
       FORMUPLA.Height = MATOP + COEFV * (ALTU% + 0.5) + 1.5 * FORMUPLA.Command1.Height
       '
       FORMUPLA.Left = BIZQ% * Screen.Width / 80
       FORMUPLA.Width = COEF * (ALIB% + 1)
       '
       FORMUPLA.Command1.Left = (FORMUPLA.ScaleWidth - 2 * FORMUPLA.Command1.Width) / 3
       FORMUPLA.Command2.Left = 2 * FORMUPLA.Command1.Left + FORMUPLA.Command1.Width
       FORMUPLA.Command1.Top = FORMUPLA.ScaleHeight - 1.3 * FORMUPLA.Command1.Height
       FORMUPLA.Command2.Top = FORMUPLA.ScaleHeight - 1.3 * FORMUPLA.Command1.Height
       FORMUPLA.Command1.Visible = True
       FORMUPLA.Command2.Visible = True
       '
       For II% = 1 To KPLA%
          If TRIM$(LABE$(II%)) <> "" Then
              CLL% = COLU%(II%) - 2
              FORMUPLA.Label(II% - 1).Font = FORMUPLA.Label(0).Font
              FORMUPLA.Label(II% - 1).FontBold = FORMUPLA.Label(0).FontBold
              FORMUPLA.Label(II% - 1).Caption = LABE$(II%) + " "
              FORMUPLA.Label(II% - 1).Left = COEF * (COLU%(II%) - Len(LABE$(II%)) - 2)
              FORMUPLA.Label(II% - 1).Top = 20 + COEFV * RENG%(II%) - COEFV / 2
              FORMUPLA.Label(II% - 1).Width = COEF * (1 + Len(LABE$(II%)))
              FORMUPLA.Label(II% - 1).Visible = True
              '
          End If
          '
          FORMUPLA.ENTRA(II% - 1).Left = COEF * (COLU%(II%) - 1)
          FORMUPLA.ENTRA(II% - 1).Top = COEFV * RENG(II%) - COEFV / 2
          FORMUPLA.ENTRA(II% - 1).Width = COEF * Int(XVALO(Mid$(Form$(II%), 2)))
          FORMUPLA.ENTRA(II% - 1).MaxLength = MALEN%(II%)
          FORMUPLA.ENTRA(II% - 1).Visible = True
          '
          If TRIM$(AVAL$(II%)) <> "" Then
            If LGEC%(II%) > 0 Then
               FORMUPLA.ECO(II% - 1).Left = COEF * (CLEC%(II%) - 1)
               FORMUPLA.ECO(II% - 1).Top = COEFV * RGEC%(II%) - COEFV / 2
               FORMUPLA.ECO(II% - 1).Width = COEF * (LGEC%(II%) + 1)
               FORMUPLA.ECO(II% - 1).Visible = True
            End If
          End If
          '
          If TRIM$(AVAL$(II%)) <> "" Or TIVAA%(II%) = 5 Then
            If POED%(II%) = 1 Then
              FORMUPLA.BOTSEL(II% - 1).Top = FORMUPLA.ENTRA(II% - 1).Top
              FORMUPLA.BOTSEL(II% - 1).Height = FORMUPLA.ENTRA(II% - 1).Height
              FORMUPLA.BOTSEL(II% - 1).Width = FORMUPLA.BOTSEL(II% - 1).Height * 0.6
              FORMUPLA.ENTRA(II% - 1).Width = FORMUPLA.ENTRA(II% - 1).Width + FORMUPLA.BOTSEL(II% - 1).Height * 0.6
              FORMUPLA.BOTSEL(II% - 1).Left = FORMUPLA.ENTRA(II% - 1).Left + FORMUPLA.ENTRA(II% - 1).Width - FORMUPLA.BOTSEL(II% - 1).Width
              FORMUPLA.BOTSEL(II% - 1).Visible = True
              If FORMUPLA.Width < FORMUPLA.BOTSEL(II% - 1).Left + FORMUPLA.BOTSEL(II% - 1).Width + 350 Then
                 FORMUPLA.Width = FORMUPLA.BOTSEL(II% - 1).Left + FORMUPLA.BOTSEL(II% - 1).Width + 350
              End If
            End If
          End If
          '
       Next II%
       '
End Sub


Function OBJPLA$(OBX$, VDEF$)
       '
       Static CAOPLA%, OBNAME$(16), OBVAL$(16), LONCO%
       '
       If LONCO% < 1 Then
          LONCO% = XVALO(Control$("MASTER", "LOMAXCOD"))
          If LONCO% < 16 Or LONCO% > 24 Then LONCO% = 16
       End If
       '
       PPXX% = InStr(OBX$, "/")
       If PPXX% < 2 Then PPXX% = 1 + Len(OBX$)
       OBY$ = Left$(OBX$, PPXX% - 1)
       TTY$ = Mid$(OBX$, PPXX% + 1)
       '
       If FORMUPLA.Visible = True Then
          Call MENSERR(24, "Imposible Presentar Formulario '" + TRIM$(OBY$) + "'\en el Contexto Actual del Sistema.")
          Exit Function
       End If
       '
       Call PREPLA(OBY$, VDEF$, TTY$)
       OKEY% = 0
       '
10302  XXY$ = Left$(VDEF$ + String$(LOREC%, 0), LOREC%)
       For II% = 1 To KPLA%
          FORMUPLA.ENTRA(II% - 1).TEXT = ""
          XXZ$ = Mid$(XXY$, PXGR%(II%), LGGR%(II%))
          FORMUPLA.ENTRA(II% - 1).TEXT = TRIM$(MSTRING$(XXZ$, Form$(II%) + "<"))
       Next II%
       '
       Call APLICACIONSKINS(FORMUPLA)
       
       SMP = Screen.MousePointer
       Screen.MousePointer = 1
       FORMUPLA.Show 1
       Screen.MousePointer = SMP
       Call FRESHALL
       OBJPLA$ = ""
       '
       If OKEY% = (-1) Then
           GoTo 10302
           '
         ElseIf OKEY% = 1 Then
           '
           For II% = 1 To KPLA%
              '
              VDS$(II%) = TRIM$(FORMUPLA.ENTRA(II% - 1).TEXT)
              If TIVAA%(II%) = 7 Then VDS$(II%) = AJUSTI$(VDS$(II%), LONCO%)
              XXZ$ = CBIN$(VDS$(II%), TIVAA%(II%), LGGR%(II%))
              Mid$(XXY$, PXGR%(II%), LGGR%(II%)) = XXZ$
              '
           Next II%
           OBJPLA$ = XXY$
           '
       End If
       '
End Function

Sub SETNEXTFOCPLA(IIAA%)
    '
    For II% = IIAA% + 1 To KPLA%
       If POED%(II%) = 1 Then
           On Error Resume Next
           FORMUPLA.ENTRA(II% - 1).SetFocus
           On Error GoTo 0
           Exit Sub
       End If
    Next II%
    FORMUPLA.Command1.SetFocus
    '
End Sub
