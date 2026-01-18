Attribute VB_Name = "TEXT_FUNCTIONS"
Public CADENATEX$()
Function CONVERLECBAR$(Texto$)
    'Reemplaza el caracter (') por (-) y todos a Mayuscula
    TENTRA$ = REPLACAR(Texto$, "'", "-")
    CONVERLECBAR = UCase(TENTRA$)
End Function
 Sub Duplicados(COMBO1 As comboBox)
        'UTILIZAR Duplicados1 FUNCIONA MEJOR
        'depura duplicados de un combo box
        Dim Search1 As Long
        Dim Search2 As Long
        Dim KillDupe As Long
        KillDupe = 0
        For Search1& = 0 To COMBO1.ListCount - 1
          For Search2& = Search1& + 1 To COMBO1.ListCount - 1
            KillDupe = KillDupe + 1
            If COMBO1.List(Search1&) = COMBO1.List(Search2&) Then
              If Search2& < COMBO1.ListCount - 1 Then
                COMBO1.RemoveItem Search2&
                Search2& = Search2& - 1
              End If
            End If
          Next Search2&
        Next Search1&
End Sub

 Sub DuplicadosEnList(List As ListBox)
        'UTILIZAR DuplicadosList1 QUE FUNCIONA MEJOR
        'depura duplicados de un combo box
        Dim Search1 As Long
        Dim Search2 As Long
        Dim KillDupe As Long
        KillDupe = 0
        For Search1& = 0 To List.ListCount - 1
          For Search2& = Search1& + 1 To List.ListCount - 1
            KillDupe = KillDupe + 1
            If List.List(Search1&) = List.List(Search2&) Then
              If Search2& < List.ListCount - 1 Then
                List.RemoveItem Search2&
                Search2& = Search2& - 1
              End If
            End If
          Next Search2&
        Next Search1&
End Sub
Function TieneCaracteresAsciiMayora127(textox$)
 
  'devuleve EL ASCII QUE ENCUENTRA > 127
  For i% = 1 To Len(textox$)
     LETRA$ = Mid$(textox$, i%, 1)
     valor = Asc(LETRA)
     If valor > 127 Then
        TieneCaracteresAsciiMayora127 = valor
     End If
  Next i%
  
End Function
Function CBIN$(XXZ$, TIVA%, LOCA%)
    '
    CBIN$ = ""
    
    On TIVA% GoTo 910, 910, 930, 930, 950, 960, 970, 980
    Exit Function
    '
910 If LOCA% = 1 Then
        If XVALO(XXZ$) >= 0 And XVALO(XXZ$) < 256 Then
           CBIN$ = Chr$(XVALO(XXZ$))
        End If
      ElseIf LOCA% = 2 Then
        If Abs(XVALO(XXZ$)) < 32768 Then
           CBIN$ = MKI$(XVALO(XXZ$))
        End If
    End If
    Exit Function
    '
930 If LOCA% = 4 Then
        CBIN$ = MKS$(XVALO(XXZ$))
      ElseIf LOCA% = 8 Then
        CBIN$ = MKD$(XVALO(XXZ$))
    End If
    Exit Function
    '
950 If Len(XXZ$) = 2 Then
        CBIN$ = XXZ$
      ElseIf Len(XXZ$) >= 6 Then
        FFX% = FECHANUM(TRIM$(XXZ$))
        CBIN$ = MKI$(FFX%)
    End If
    Exit Function
    '
960 CBIN$ = Left$(TRIM$(XXZ$) + Space$(LOCA%), LOCA%)
    Exit Function
    '
970 If TRIM$(XXZ$) = "" Then
        FFX% = 0
        CBIN$ = MKI$(0)
      Else
        FFX% = CODINT%(XXZ$)
        CBIN$ = MKI$(FFX%)
    End If
    Exit Function
    '
980 FFX% = Int(HORANUM(XXZ$))
    CBIN$ = MKI$(FFX%)
    '
    Exit Function
    '
End Function
'

Function CSTRING$(XXZ$, TIVA%, LOTOT%, LODEC%, JUSTI%)
   '
    CSTRING$ = Space$(LOTOT%)
    On TIVA% GoTo 910, 910, 930, 930, 950, 960, 970, 980
    Exit Function
    '
910 If Len(XXZ$) < 2 Then
          VDI% = Asc(XXZ$ + Chr$(0))
       Else
          VDI% = CVI(XXZ$)
    End If
    '
    If JUSTI% = 1 Then
          CSTRING$ = AJUSTI$(Str$(VDI%), LOTOT%)
       Else
          CSTRING$ = AJUSTD$(Str$(VDI%), LOTOT%)
    End If
    '
    If XVALO(CSTRING$) = 0 Then CSTRING$ = Space$(LOTOT%)
    Exit Function
    '
930 If Len(XXZ$) < 8 Then
          VDN# = CVS(XXZ$ + String$(4, 0))
       Else
          VDN# = CVD(XXZ$)
    End If
    '
    CS$ = NUMSTRI$(VDN#, LOTOT%, LODEC%, 0)
    '
    If JUSTI% = 1 Then
          CSTRING$ = AJUSTI$(CS$, LOTOT%)
       Else
          CSTRING$ = AJUSTD$(CS$, LOTOT%)
    End If
    '
    Exit Function
    '
950 FEX = CVI(XXZ$ + String$(2, 0))
    CS$ = FECHATEX$(FEX)
    '
    If JUSTI% = 1 Then
          CSTRING$ = Left$(CS$ + Space$(LOTOT%), LOTOT%)
       Else
          CSTRING$ = Right$(Space$(LOTOT%) + CS$, LOTOT%)
    End If
    '
    Exit Function
    '
960 CS$ = XXZ$
    '
    If JUSTI% = 1 Then
          CSTRING$ = AJUSTI$(CS$, LOTOT%)
       Else
          CSTRING$ = AJUSTD$(CS$, LOTOT%)
    End If
    '
    Exit Function
    '
970 CII% = CVI(XXZ$ + String$(2, 0))
    CS$ = CODEXT$(CII%)
    '
    If JUSTI% = 1 Then
          CSTRING$ = AJUSTI$(CS$, LOTOT%)
       Else
          CSTRING$ = AJUSTD$(CS$, LOTOT%)
    End If
    '
    Exit Function
    '
980 If Len(XXZ$) < 4 Then
         CS$ = HORATEX$(CVI(XXZ$ + String$(2, 0)))
       Else
         CS$ = HORMINSE$(CVS(XXZ$))
    End If
    '
    If JUSTI% = 1 Then
          CSTRING$ = TRIM$(CS$)
          If InStr(CSTRING$, ":") = 2 Then
             CSTRING$ = " " + CSTRING$
          End If
          While Len(CSTRING$) < LOTOT%
             CSTRING$ = CSTRING$ + " "
          Wend
          CSTRING$ = Left$(CSTRING$, LOTOT%)
       Else
          CSTRING$ = AJUSTD$(CS$, LOTOT%)
    End If
    '
    Exit Function
    '
End Function
'
Function MSTRING$(XXZ$, Formato$)
    '
    FRX$ = UCase$(TRIM$(Formato$))
    '
    LETRA$ = Left$(FRX$, 1)
    If LETRA$ = "I" Then
         TV% = 2
       ElseIf LETRA$ = "D" Then
         TV% = 5
       ElseIf LETRA$ = "A" Then
         TV% = 6
       ElseIf LETRA$ = "C" Then
         TV% = 7
       ElseIf LETRA$ = "F" Then
         TV% = 4
       ElseIf LETRA$ = "H" Or LETRA$ = "T" Then
         TV% = 8
       Else
         GoTo 99
    End If
    '
    FRX$ = Mid$(FRX$, 2)
    LCC% = XVALO(FRX$)
    If LCC% < 1 Then GoTo 99
    '
    LD% = 0
    If TV% = 3 Or TV% = 4 Then
        LD% = 10 * (XVALO(FRX$) - Int(XVALO(FRX$)))
    End If
    '
    JTT% = 1
    If TV% <= 4 Then JTT% = 2
    If InStr(FRX$, "<") > 0 Then JTT% = 1
    If InStr(FRX$, ">") > 0 Then JTT% = 2
    '
    MSTRING$ = CSTRING$(XXZ$, TV%, LCC%, LD%, JTT%)
    Exit Function
    '
99  Call MENSERR(24, "Error Invocacion Funcion 'MSTRING' - Formato no Procesable")
    Close: Call FINAL("")
    '
End Function
'
Function BINVAL%(A$)   ' determina un integer por suma de los
    '                  ' valores ASCII de un string.
    BV& = 0
    For KI% = 1 To Len(A$)
      BV& = BV& + Asc(Mid$(A$, KI%, 1))
    Next KI%
    While BV& > 32767
      BV& = BV& - 32768
    Wend
    BINVAL% = BV&
    '
End Function
'
Function TEXTFORM$(VDEF$, Formato$, Optional Moneda$, Optional MONEDB$, Optional PRINT_MONEI%)

'SI PRINT_MONEI% = 0 IMPRIME LA MONEDA EN CASO QUE SEA -1 NO IMPRIME LA MONEDA SOLO IMPRIME EL IMPORTE EN LETRAS

If Left$(VDEF$, 3) = "===" Then
AAA = BBB
End If

FRX$ = UCase$(TRIM$(Formato$))
If Left$(FRX$, 1) = "<" Then FRX$ = Mid$(FRX$, 2) + "<"
If Left$(FRX$, 1) = ">" Then FRX$ = Mid$(FRX$, 2) + ">"
'
EXPA% = 0
If Left$(FRX$, 1) = "@" Then
    FRX$ = Mid$(FRX$, 2)
    EXPA% = 1
End If

LETRA$ = Left$(FRX$, 1)
If LETRA$ = "I" Then
     TV% = 2
  ElseIf LETRA$ = "D" Then
     TV% = 5
  ElseIf LETRA$ = "A" Or Left$(VDEF$, 3) = "===" Then
     TV% = 6
  ElseIf LETRA$ = "C" Or LETRA$ = "K" Then
     TV% = 7
  ElseIf LETRA$ = "F" Or LETRA$ = "G" Or LETRA$ = "$" Then
     TV% = 4
  ElseIf LETRA$ = "H" Or LETRA$ = "T" Then
     TV% = 8
  ElseIf LETRA$ = "M" Then
     TV% = 99
  ElseIf LETRA$ = "X" Then
     TV% = 99
  ElseIf LETRA$ = "Z" Then
     TV% = 99
  Else
     Exit Function
End If
'
FRX$ = Mid$(FRX$, 2)
LCII% = XVALO(FRX$)
If LCII% < 1 Then Exit Function
If TV% = 99 Then
   PPXYZ% = InStr(FRX$, "/")
   CAREMAX% = XVALO(Mid$(FRX$, PPXYZ% + 1))
   If CAREMAX% < 1 Then CAREMAX% = 1
   VDEF1$ = VDEF$
   If VDEF$ = "" Then GoTo 5
   '
   ' DETERMINA SI ES UN NUMERO
   If IsNumeric(VDEF$) = False Or LETRA$ = "M" Then GoTo 5
   VDEF2$ = TRIM$(VDEF$)
'   For KKII% = 1 To Len(VDEF2$)
'      AASS% = Asc(Mid$(VDEF2$, KKII%, 1))
'      If AASS% < 45 Or AASS% > 57 Then GoTo 5
'   Next KKII%
   '
   If Abs(XVALO(VDEF$)) > 999999999 Then GoTo 5
   '
   VDEF1$ = NUMTEX$(XVALO(VDEF2$))
   
   If Moneda$ = "" Then Moneda$ = "Pesos"
   If PRINT_MONEI% >= 0 Then VDEF1$ = Moneda$ + " " + VDEF1$
   If MONEDB$ = "" Then MONEDB$ = "Ctvs"
      VDEF1$ = VDEF1$ + " " + MONEDB$
      '
5  Call FRATEXTO(VDEF1$, LCII%)
   If CARETEX% > CAREMAX% Then CARETEX% = CAREMAX%
   VDSS$ = ""
   For KKIIX% = 1 To CARETEX%: VDSS$ = VDSS$ + " " + TRIM$(RETEX$(KKIIX%)) + vbCrLf: Next KKIIX%
   GoTo 999
End If
'
On TV% GoTo 10, 20, 30, 40, 50, 60, 70, 80
'
10 ' continue
20 If LCII% > 4 Then TV% = 4: GoTo 40
   LKII% = LCII%: LD% = 0
   JTII% = 2: If InStr(FRX$, "<") > 0 Then JTII% = 1
   If InStr(FRX$, "+") > 0 Then TV% = 1
   VD# = CDbl(XVALO(VDEF$))
   GoTo 100
30 ' continue
40 LKII% = LCII%: LD% = 10 * (XVALO(FRX$) - Int(XVALO(FRX$)))
   JTII% = 2: If InStr(FRX$, "<") > 0 Then JTII% = 1
   If InStr(FRX$, "+") > 0 Then TV% = 3
   If Left$(VDEF$, 1) = "*" Then
       VDSS$ = VDEF$            ' RELLENA CON ASTERISCOS
       JTII% = 1
       GoTo 110
   End If
   VD# = CDbl(XVALO(VDEF$))
   If LETRA$ = "$" Or LETRA$ = "G" Then
     FORMAX$ = Formato$
     VDSS$ = FORMATNUM$(VD#, FORMAX$)
     GoTo 999
   End If
   GoTo 100
50 If XVALO(FRX$) > 0 Then LKII% = 8: LCII% = XVALO(FRX$)
   LD% = 0: JTII% = 1
   VDSS$ = VDEF$
   GoTo 100
60 LD% = 0: JTII% = 1
   If InStr(FRX$, ">") > 0 Then JTII% = 2
   If InStr(FRX$, Chr$(94)) > 0 Then JTII% = 3
   VDSS$ = VDEF$
   GoTo 100
   '
70 LKII% = 16
   JTII% = 1: LD% = 0
   VDSS$ = VDEF$: TV% = 6      ' :VD%=CODINT%(VDEF$)
   If LETRA$ = "K" Then VDSS$ = "Codigo: " + VDEF$
   GoTo 100
   '
80 If XVALO(FRX$) > 0 Then LKII% = 5: LCII% = XVALO(FRX$)
   LD% = 0: JTII% = 2
   VDSS$ = Left$(VDEF$, LCII%)
   GoTo 100
   '
100 If TV% <= 4 Then
       SEMIL% = 0
       If LETRA$ = "G" Or TV% < 3 Then SEMIL% = 1
       VDSS$ = NUMSTRI$(VD#, LCII%, LD%, SEMIL%)
    End If
    '
110 If JTII% = 1 Then                 ' izquierda
       VDSS$ = AJUSTI$(VDSS$, LCII%)
      ElseIf JTII% = 2 Then           ' derecha
       VDSS$ = AJUSTD$(VDSS$, LCII%)
       While (Len(VDSS$) > 0 And Len(VDSS$) > LCII%)
         VDSS$ = Mid$(VDSS$, 2)
       Wend
      ElseIf JTII% = 3 Then           ' centrado
       VDSS$ = TRIM$(VDSS$)
       While Len(VDSS$) < LCII%: VDSS$ = " " + VDSS$ + " ": Wend
       VDSS$ = Left$(VDSS$, LCII%)
    End If
    '
999 TEXTFORM$ = VDSS$
    '
End Function
'

Sub MENSERR(A%, B$, Optional c$)
    '
    Static PRIMEN%
    'MODIFICACION A PARTIR DE PARAMETRO C
    'PARAMETRO C$ PARA QUE PUEDA TOMAR DE LA BASE DE DATOS UN MENSAJE PRECONFIGURADO
    'VALIDAR SI TIENE NUMERO DE MENSAJE
    'SI TIENE NUMERO RECORRE PARA PASAR ASIGNAR LOS PARMETROS
    'REEMPLAZAR LOS PARAMETROS POR LOS CARACTERES %% DE LA CADENA OBTENIDA DE LA BD
    'SI SE OBTIENE MENSAJE REEMPLAZAR ASIGNARLO A VARIABLE B$
    
    PARAME$ = TRIM$(c$)
    '
    Static MENSANT$, TACCANT
    If Timer < TACCANT + 3 Then
       If B$ = MENSANT$ Then
          Exit Sub
       End If
    End If
    TACCANT = Timer
    '
    MENSANT$ = B$
    
    'DESDE AQUI NUEVO
    MENSAB$ = B$
    '
    'VALIDA SI EXISTE EL ARCHIVO
    If EXISTE%(App.Path + "\MENSAFLEX.MDB") < 1 Then GoTo 90
    IDNUME& = 0
    '
    'OPCIONES DE IDIOMA
    Select Case LANGUA%
      Case 0: CAMPOIDIOMA$ = "MENSAESP"
      Case 1: CAMPOIDIOMA$ = "MENSAING"
      Case 5: CAMPOIDIOMA$ = "MENSAPOR"
      Case Else: CAMPOIDIOMA$ = "MENSAESP"
    End Select

    If Left$(MENSAB$, 1) = "#" Then 'SI TIENE SIMBOLO
       POS1% = InStr(MENSAB$, Chr$(124))
       If POS1% < 1 Then GoTo 90
       IDNUME& = XVALO(Mid$(MENSAB$, 2, POS1% - 1))
    End If
    'SI NO TIENE NUMERO SIGUE COMO ANTES
    If IDNUME& < 1 Then
       POS1% = InStr(MENSAB$, Chr$(124))
       MENSAB$ = TRIM$(Mid$(MENSAB$, POS1% + 1))
       GoTo 85
    End If
    '
    'OBTENGO LA FRASE DEL MENSAJE DE LA BASE DE DATOS Y LA VALIDO
    CONDI$ = "NUMMENSA =" & IDNUME&
    TXMEN$ = VALOMENSA("MENSAJES", CAMPOIDIOMA$, CONDI$, "NOMESS")
    '
    'SI NO EXISTE LE QUITO LO QUE ESTA A LA IZQUIERDA DEL CHR(124)
    If TRIM$(TXMEN$) = "" Then
       POS1% = InStr(MENSAB$, Chr$(124))
       MENSAB$ = TRIM$(Mid$(MENSAB$, POS1% + 1))
       GoTo 85
    End If
    
    'RECORRE LOS PARAMETROS ENVIADOS
    FRASETOTAL$ = ""
    While Len(PARAME$) > 0
       'OBTENGO LOS PARAMETROS EN ORDEN DEL PRIMERO AL ULTIMO
       POS1% = InStr(PARAME$, Chr$(124))
       If POS1% < 1 Then POS1% = 1 + Len(PARAME$)
       Parametro = Left$(PARAME$, POS1% - 1)
       '
       'RECORRE EL TEXTO DE LA BASE DE DATOS.REEMPLAZO LOS SIMBOLOES %%
       'EN ORDEN DEL PRIMERO AL ULTIMO
       POS1TX% = InStr(TXMEN$, "%%")
       If POS1TX% < 1 Then GoTo 80
       FRASE$ = Left$(TXMEN$, POS1TX% - 1)
       FRASETOTAL$ = FRASETOTAL$ + FRASE$ + Parametro
       TXMEN$ = Mid$(TXMEN$, POS1TX% + 2)
       '
       'TOMO EL SALDO DE LA CADENA
       PARAME$ = Mid$(PARAME$, POS1% + 1)
     Wend
     '
80   FRASETOTAL$ = FRASETOTAL$ & TXMEN$ 'ARMO LA FRASE PARA LOS CASOS QUE NO FINALIZA CON PARAMETRO
     '
     'POR SI NO SE PASA ALGUN PARAMETRO QUE NO PRESENTE LOS SIMBOLOS DE PORCENTAJE
     FRASETOTAL$ = REPLACAR(FRASETOTAL$, "%%", "")
     FRASETOTAL$ = REPLACAR(FRASETOTAL$, CRLF, "\")
     'SI OBTIENE EL MENSAJE LE ASIGNO A MENSAB$
     If TRIM$(FRASETOTAL$) <> "" Then MENSAB$ = FRASETOTAL$
     GoTo 90
   '
     'BUSCA MENSAJE VIEJO
85   CONDI$ = "MensaViejo = '" & MENSAB$ & "'"
     TEXTOV$ = VALOMENSA("MENSAJES", "MensaViejo", CONDI$, "NOMESS")
     'SI LO OBTIENE BUSCA SI TIENE TEXTO EN CAMPO SEGUN IDIOMA
     If TRIM$(TEXTOV$) <> "" Then
        TEXTOV$ = VALOMENSA("MENSAJES", CAMPOIDIOMA$, CONDI$, "NOMESS")
     End If
     If TRIM$(TEXTOV$) <> "" Then MENSAB$ = TEXTOV$
     '
90   TIMENSA% = 2
     If PRIMEN% <> 1 Then
        TINI = Timer
        While Timer < TINI + 1: Wend
        ' PRIMEN% = 1
     End If
     MMXX% = COMALTER%(MENSAB$ + "\\Aceptar Mensaje")
    '
End Sub


Function REPLACEN$(AA1$)
    '
    aa2$ = REPLACAR$(AA1$, Chr$(160), "a")
    aa3$ = REPLACAR$(aa2$, Chr$(130), "e")
    AA4$ = REPLACAR$(aa3$, Chr$(161), "i")
    AA5$ = REPLACAR$(AA4$, Chr$(162), "o")
    AA6$ = REPLACAR$(AA5$, Chr$(163), "u")
    AA7$ = REPLACAR$(AA6$, Chr$(164), "n")
    AA8$ = REPLACAR$(AA7$, Chr$(165), "N")
    '
    For KKII% = 1 To Len(AA8$)
      VLETRA% = Asc(Mid$(AA8$, KKII%, 1))
      If VLETRA% < 32 Or VLETRA% > 128 Then
         Mid$(AA8$, KKII%, 1) = "."
      End If
    Next KKII%
    '
    REPLACEN$ = AA8$
    '
End Function
'

Function REPLACE$(Texto$, SUBTEX$, PINI%, LON%)
    '
    If Texto$ = "" Then
       TX$ = ""
       SUBTEX$ = ""
       PINI% = 0
       LON% = 0
       Exit Function
    End If
    '
    TX$ = Texto$
    ST$ = SUBTEX$
    PX% = PINI%: If PX% < 1 Then PX% = 1
    Lx% = LON%: If Lx% = 0 Then Lx% = Len(ST$)
    '
    While Len(TX$) < PX% + Lx% - 1
        TX$ = TX$ + " "
    Wend
    '
    ST$ = Left$(ST$ + Space$(Lx%), Lx%)
    Mid$(TX$, PX%, Lx%) = ST$
    '
    Texto$ = TX$
    REPLACE$ = TX$
    '
End Function
'
Sub SUPRACENTO(TTXX$)
    '
    TTXX1$ = REPLACAR$(TTXX$, Chr$(130), Chr$(101))
    TTXX2$ = REPLACAR$(TTXX1$, Chr$(160), Chr$(97))
    TTXX$ = REPLACAR$(TTXX2$, Chr$(162), Chr$(111))
    '
End Sub
'
Sub OPNOIN(KXX$)
    '
    TTXX$ = "La Opción que ha Elegido"
    If TRIM$(KXX$) <> "" Then
        TTXX$ = TTXX$ + ":\  \'" + TRIM$(KXX$) + "'\   "
    End If
    TTXX$ = TTXX$ + "\no se Encuentra Disponible\en la Versión FLEXOFT Actualmente Instalada.\  \Consulte con su Soporte de Sistemas\o su Proveedor de Software para la\Adquisición e Instalación de la Misma."
    Call COMUNI(TTXX$)
    '
End Sub
'

Sub HELPONLINE(CAPI$)
    '
    APLICA$ = AJUSTI$(CAPI$, 8)
    FORMUHLP.TEPRUEBA.Caption = String$(72, "*")
    ALIX = FORMUHLP.TEPRUEBA.Width
    FORMUHLP.Width = ALIX + 240
    FORMUHLP.Lista.Width = FORMUHLP.Width
    FORMUHLP.Frame1.Left = FORMUHLP.Lista.Width - FORMUHLP.Frame1.Width
    '
    FORMUHLP.TEPRUEBA.Caption = String$(60, "*")
    ALIX = FORMUHLP.TEPRUEBA.Width
    FORMUHLP.LIAPLI.Width = ALIX + 240
    FORMUHLP.LIAPLI.Left = (FORMUHLP.Lista.Width - FORMUHLP.LIAPLI.Width) / 2
    '
    FORMUHLP.Refresh
    SMP = Screen.MousePointer
    Screen.MousePointer = 1
    FORMUHLP.Show 1
    Screen.MousePointer = SMP
    Call FRESHALL
    '
End Sub
'
Sub CARGAHELP()
    '
    Static PRIVEZ%
    Screen.MousePointer = 11
    FORMUHLP.Refresh
    FORMUHLP.Lista.Clear
    NHX% = FreeFile
    Dim RNX As String * 80
    NHX% = FILEOPEN%(LUBAS$ + "FLEXHLP.DAT", 0, 80)
    FIN& = LOF(NHX%) / 80
    If FIN& < 1 Then FIN& = 1
    '
    For i& = 1 To FIN&
       Call TRACE(20, i&, FIN&)
       Get #NHX%, i&, RNX$
       APLIX$ = Left$(RNX$, 8)
       '
       If PRIVEZ% = 0 Then
          If APLIX$ = APLIXA$ Then GoTo 20
          If TRIM$(APLIX$) = "" Then GoTo 20
          If Asc(Left$(APLIX$, 1)) < 35 Then GoTo 20
          If Asc(Left$(APLIX$, 1)) > 122 Then GoTo 20
          If TRIM$(Mid$(RNX$, 9)) = "" Then GoTo 20
          APLIXA$ = APLIX$
          For J& = 0 To FORMUHLP.LIAPLI.ListCount - 1
             FORMUHLP.LIAPLI.ListIndex = J&
             If Left$(FORMUHLP.LIAPLI.TEXT, 8) = APLIX$ Then GoTo 20
          Next J&
          FORMUHLP.LIAPLI.AddItem APLIX$ + "  " + Mid$(RNX$, 9)
       End If
       '
20     If APLIX$ = APLICA$ Then
          FORMUHLP.Lista.AddItem Mid$(RNX$, 9)
       End If
    Next i&
    Screen.MousePointer = 1
    PRIVEZ% = (-1)
    '
End Sub
'
Sub COMUNI(B$)
    '
    TIMENSA% = 1
    MMXX% = COMALTER%(B$ + "\\Aceptar Mensaje")
    '
End Sub
'

Sub FRATEXTO(TX$, LOTEX%, Optional NOTRIM%, Optional AUTANCH%)
   '
   For KK1% = 0 To 256: RETEX$(KK1%) = "": Next KK1%
   If LOTEX% < 2 Then LOTEX% = 16
   '
   DE$ = RTrim(TX$)
   While Len(DE$) > 0
      If Asc(Mid$(DE$, Len(DE$), 1)) > 32 Then GoTo 2
      DE$ = Left$(DE$, Len(DE$) - 1)
   Wend
   '
2  If NOTRIM% < 1 Then     ' RECORTA BLANCOS AL PRINCIPIO
      On Error Resume Next
      While Len(DE$) > 0
        If Asc(Left$(DE$, 1)) > 32 Then GoTo 3
        DE$ = Mid$(DE$, 2)
      Wend
3     While Len(DE$) > 0
        If Asc(Mid$(DE$, Len(DE$), 1)) > 32 Then GoTo 5
        DE$ = Left$(DE$, Len(DE$) - 1)
      Wend
   End If
5  On Error GoTo 0
   kk% = 0
   '
   If AUTANCH% = 1 Then
      FORMUPLA.AutoAnch.Caption = String$(LOTEX%, "*")
      LONTEX& = FORMUPLA.AutoAnch.Width
   End If
   '
10 If Len(DE$) < 1 Then GoTo 99
   '
   PPXX% = InStr(DE$, Chr$(13) + Chr$(10))
   If PPXX% < 1 Then PPXX% = InStr(DE$, CRLF$)
   If PPXX% < 1 Then PPXX% = 1 + Len(DE$)
   If AUTANCH% = 1 Then
      FORMUPLA.AutoAnch.Caption = Left$(DE$, PPXX% - 1)
      If FORMUPLA.AutoAnch.Width <= LONTEX& Then
        If kk% >= 256 Then GoTo 99
        kk% = kk% + 1
        RETEX$(kk%) = Left$(DE$, PPXX% - 1) + "  "
        DE$ = Mid$(DE$, PPXX% + 2)
        GoTo 10
      End If
      LOTEX1% = PPXX% * LONTEX& / FORMUPLA.AutoAnch.Width - 1
      GoTo 20
   End If
   '
   If PPXX% <= LOTEX% + 1 Then
      If kk% >= 256 Then GoTo 99
      kk% = kk% + 1
      RETEX$(kk%) = Left$(DE$, PPXX% - 1) + "  "
      DE$ = Mid$(DE$, PPXX% + 2)
      GoTo 10
   End If
   LOTEX1% = LOTEX%
   '
20 For PPXX% = LOTEX1% + 1 To 1 Step -1
      If Mid$(DE$, PPXX%, 1) = " " Then
         '
         If AUTANCH% = 1 Then
           FORMUPLA.AutoAnch.Caption = Left$(DE$, PPXX% - 1)
           If FORMUPLA.AutoAnch.Width <= LONTEX& Then
             If kk% >= 256 Then GoTo 99
             kk% = kk% + 1
             RETEX$(kk%) = Left$(DE$, PPXX% - 1) + "  "
             DE$ = Mid$(DE$, PPXX% + 1)
             GoTo 10
           End If
           If CInt(PPXX% * LONTEX& / FORMUPLA.AutoAnch.Width - 1) <> LOTEX1% Then
              LOTEX1% = PPXX% * LONTEX& / FORMUPLA.AutoAnch.Width - 1
              GoTo 20
           End If
         End If
         '
         If kk% >= 256 Then GoTo 99
         kk% = kk% + 1
         RETEX$(kk%) = Left$(DE$, PPXX% - 1)
         DE$ = Mid$(DE$, PPXX% + 1)
         GoTo 10
      End If
      '
      If Mid$(DE$, PPXX%, 1) = "," Then
         If kk% >= 256 Then GoTo 99
         kk% = kk% + 1
         RETEX$(kk%) = Left$(DE$, PPXX%)
         DE$ = Mid$(DE$, PPXX% + 1)
         GoTo 10
      End If
   Next PPXX%
   '
   If kk% >= 256 Then GoTo 99
   If LOTEX% > 1 Then
      kk% = kk% + 1
      'RETEX$(KK%) = Left$(DE$, LOTEX% - 1) + "-" ASI ESTABA ANTES SE MODIFICA PARA QUE NO AGREGUE EL GUION
      RETEX$(kk%) = Left$(DE$, LOTEX% - 1)        'QUE UNE PALABRAS CUANDO CORTA ALGUNA POR SOBREPASAR LA LONGITUD PARAMETRIZADA.
   End If
   DE$ = Mid$(DE$, LOTEX%)
   GoTo 10
   '
99 CARETEX% = kk%
   '
End Sub
'
Sub FRATEXFONT(TX$, LOTEMM, FONAM, FOSIZ, FOBOL)
   '
   For KK1% = 0 To 256: RETEX$(KK1%) = "": Next KK1%
   '
   DE$ = RTrim(TX$)
   While Len(DE$) > 0
      If Asc(Mid$(DE$, Len(DE$), 1)) > 32 Then GoTo 2
      DE$ = Left$(DE$, Len(DE$) - 1)
   Wend
   '
2  If NOTRIM% < 1 Then     ' RECORTA BLANCOS AL PRINCIPIO
      On Error Resume Next
      While Len(DE$) > 0
        If Asc(Left$(DE$, 1)) > 32 Then GoTo 3
        DE$ = Mid$(DE$, 2)
      Wend
3     While Len(DE$) > 0
        If Asc(Mid$(DE$, Len(DE$), 1)) > 32 Then GoTo 5
        DE$ = Left$(DE$, Len(DE$) - 1)
      Wend
   End If
5  On Error GoTo 0
   kk% = 1
   '
   FORMUPLA.AutoAnch.FontName = FONAM
   FORMUPLA.AutoAnch.FontSize = FOSIZ
   FORMUPLA.AutoAnch.FontBold = FOBOL
   '
10 If Len(DE$) < 1 Then GoTo 99
   '
   NUREN% = 0
   PPXX% = InStr(DE$, Chr$(13) + Chr$(10)): NUREN% = 1
   If PPXX% < 1 Then PPXX% = InStr(DE$, CRLF$): NUREN% = 1
   If PPXX% < 1 Then PPXX% = 1 + Len(DE$): NUREN% = 0
   '
   PPYY% = InStr(DE$, " "): If PPYY% < 2 Then PPYY% = 1 + Len(DE$)
   If PPYY% < PPXX% Then PPXX% = PPYY%: NUREN% = 0
   '
   FORMUPLA.AutoAnch.Caption = LTrim$(RETEX$(kk%) + " " + Left$(DE$, PPXX% - 1))
   If FORMUPLA.AutoAnch.Width > 56.7 * LOTEMM Then
      If TRIM$(RETEX$(kk%)) <> "" Then kk% = kk% + 1
   End If
   RETEX$(kk%) = LTrim$(RETEX$(kk%) + " " + Left$(DE$, PPXX% - 1))
   If NUREN% = 1 Then kk% = kk% + 1
   DE$ = LTrim$(Mid$(DE$, PPXX% + 1))
15 If Left$(DE$, 2) = Chr$(13) + Chr$(10) Or Left$(DE$, 2) = CRLF$ Then
      kk% = kk% + 1
      DE$ = Mid$(DE$, 3)
   End If
   '
   If DE$ <> "" Then
     If Asc(Left$(DE$, 1)) <= 32 Then
        DE$ = Mid$(DE$, 2)
        GoTo 15
     End If
   End If
   GoTo 10
   '
   'If AUTANCH% = 1 Then
      FORMUPLA.AutoAnch.Caption = Left$(DE$, PPXX% - 1)
      If FORMUPLA.AutoAnch.Width <= 56.7 * LOTEMM Then
        If kk% >= 256 Then GoTo 99
        kk% = kk% + 1
        RETEX$(kk%) = Left$(DE$, PPXX% - 1) + "  "
        DE$ = Mid$(DE$, PPXX% + 2)
        GoTo 10
      End If
      LOTEX1% = PPXX% * LONTEX& / FORMUPLA.AutoAnch.Width - 1
      GoTo 20
   'End If
   '
   If PPXX% <= LOTEX% + 1 Then
      If kk% >= 256 Then GoTo 99
      kk% = kk% + 1
      RETEX$(kk%) = Left$(DE$, PPXX% - 1) + "  "
      DE$ = Mid$(DE$, PPXX% + 2)
      GoTo 10
   End If
   LOTEX1% = LOTEX%
   '
20 For PPXX% = LOTEX1% + 1 To 1 Step -1
      If Mid$(DE$, PPXX%, 1) = " " Then
         '
         If AUTANCH% = 1 Then
           FORMUPLA.AutoAnch.Caption = Left$(DE$, PPXX% - 1)
           If FORMUPLA.AutoAnch.Width <= LONTEX& Then
             If kk% >= 256 Then GoTo 99
             kk% = kk% + 1
             RETEX$(kk%) = Left$(DE$, PPXX% - 1) + "  "
             DE$ = Mid$(DE$, PPXX% + 1)
             GoTo 10
           End If
           LOTEX1% = PPXX% * LONTEX& / FORMUPLA.AutoAnch.Width - 1
           GoTo 20
         End If
         '
         If kk% >= 256 Then GoTo 99
         kk% = kk% + 1
         RETEX$(kk%) = Left$(DE$, PPXX% - 1)
         DE$ = Mid$(DE$, PPXX% + 1)
         GoTo 10
      End If
      '
      If Mid$(DE$, PPXX%, 1) = "," Then
         If kk% >= 256 Then GoTo 99
         kk% = kk% + 1
         RETEX$(kk%) = Left$(DE$, PPXX%)
         DE$ = Mid$(DE$, PPXX% + 1)
         GoTo 10
      End If
   Next PPXX%
   '
   If kk% >= 256 Then GoTo 99
   If LOTEX% > 1 Then
      kk% = kk% + 1
      'RETEX$(KK%) = Left$(DE$, LOTEX% - 1) + "-"
      RETEX$(kk%) = Left$(DE$, LOTEX% - 1)
   End If
   DE$ = Mid$(DE$, LOTEX%)
   GoTo 10
   '
99 CARETEX% = kk%
   '
End Sub
'
Function TEXTOMEDI$(MEDIX$)
   '
   TXMD$ = ""
   If Len(MEDIX$) = 12 Then
      LARGOX = CVS(Mid$(MEDIX$, 1, 4))
      ANCHOX = CVS(Mid$(MEDIX$, 5, 4))
      ALTOX = CVS(Mid$(MEDIX$, 9, 4))
      '
      If LARGOX > 0.5 Then
        If ANCHOX > 0.5 Then
          If ALTOX > 0.5 Then
            TXMD$ = TRIM$(Str$(LARGOX)) + " x " + TRIM$(Str$(ANCHOX)) + " x " + TRIM$(Str$(ALTOX))
          End If
        End If
      End If
      '
   End If
   '
   TEXTOMEDI$ = TXMD$
   '
End Function

Function ULOCASE$(TTXX$)
   '
   TTYY$ = ""
   For KKI% = 1 To Len(TTXX$)
     CARA% = Asc(Mid$(TTXX$, KKI%, 1))
     If CARA% >= 65 Then
       If CARA% <= 90 Then
         CARA% = CARA% + 32
         GoTo 50
       End If
     End If
     '
     If CARA% >= 97 Then
       If CARA% <= 122 Then
         CARA% = CARA% - 32
         GoTo 50
       End If
     End If
     '
50   TTYY$ = TTYY$ + Chr$(CARA%)
   Next KKI%
   '
   ULOCASE$ = TTYY$
End Function

Function SAFETEXT$(TXTXT)
    SAFETEXT$ = ""
    On Error Resume Next
    SAFETEXT$ = TXTXT
    On Error GoTo 0
End Function
'
Sub FRACLECBAR(CajaEntrada, CampoCodigo, CampoLote, CampoCant)
   '
   On Error Resume Next
   textox$ = CONVERLECBAR(CajaEntrada.TEXT)
   If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Imposible Procesar Lectura Codigo de Barras")
      Exit Sub
   End If
   On Error GoTo 0
   '
   TOT% = Len(textox$)
   '
   Pi% = 1
   PF% = InStr(Pi%, textox$, Chr$(124))
   If PF% = 0 Then
     CODEX$ = TRIM$(Mid$(textox$, 1, 16))
     LOTE$ = TRIM$(Mid$(textox$, 17, 12))
     CANTI$ = TRIM$(Mid$(textox$, 29, 8))
     GoTo 99
   End If
   '
   CODEX$ = Mid$(textox$, 1, PF% - 1)
   '
   Pi% = PF% + 1
   PF% = InStr(Pi%, textox$, Chr$(124))
   TAMCAD% = PF% - Pi%
   LOTE$ = Mid$(textox$, Pi%, TAMCAD%)
   '
   Pi% = PF% + 1
   PF% = InStr(Pi%, textox$, Chr$(124))
   If PF% = 0 Then PF% = TOT%
   TAMCAD% = PF% - Pi% + 1
   CANTI$ = Mid$(textox$, Pi%, TAMCAD%)
99 '
   On Error Resume Next
   CampoCodigo.Caption = TRIM$(CODEX$)
   CampoCodigo.TEXT = TRIM$(CODEX$)
   CampoLote.Caption = TRIM$(LOTE$)
   CampoLote.TEXT = TRIM$(LOTE$)
   CIXI% = CODINT%(CODEX$)
   CampoCant.Caption = FORMATNUM(XVALO(CANTI$), "F13." & CNTDC$(CIXI%))
   CampoCant.TEXT = FORMATNUM(XVALO(CANTI$), "F13." & CNTDC$(CIXI%))
   On Error GoTo 0
   '
End Sub

Function CONMUTEX$(TACT$, LREP$)
    Dim OPX$(32)
    CANOP% = 0
    LIOP$ = LREP$
    Do While LIOP$ <> ""
       PPXX% = InStr(LIOP$, "/")
       If PPXX% < 1 Then PPXX% = 1 + Len(LIOP$)
       If TRIM$(Left$(LIOP$, PPXX% - 1)) <> "" Then
         If CANOP% < 32 Then
            CANOP% = CANOP% + 1
            OPX$(CANOP%) = TRIM$(Left$(LIOP$, PPXX% - 1))
         End If
       End If
       LIOP$ = TRIM$(Mid$(LIOP$, PPXX% + 1))
    Loop
    '
    For U% = 1 To CANOP%
       If TRIM$(UCase$(TACT$)) = UCase$(OPX$(U%)) Then
          CONMUTEX = OPX$(U% + 1)
          If CONMUTEX <> "" Then Exit Function
       End If
    Next U%
    '
    CONMUTEX$ = OPX$(1)
    '
End Function
'
Function OBTIENEDIGITO$(XX$)
       
   ETAPA1 = 0
   ETAPA2 = 0
   ETAPA3 = 0
   ETAPA4 = 0
    
   For i = 1 To Len(XX$) Step 2
        DIGITO = XVALO(Mid$(XX$, i, 1))
        ETAPA1 = ETAPA1 + DIGITO
   Next i
    
   ETAPA2 = ETAPA1 * 3
   
   For i = 2 To Len(XX$) Step 2
        DIGITO = XVALO(Mid$(XX$, i, 1))
        ETAPA3 = ETAPA3 + DIGITO
   Next i
   
   ETAPA4 = ETAPA2 + ETAPA3
   
   DIGITO = 10 - (ETAPA4 Mod 10)
   
   OBTIENEDIGITO$ = CStr(DIGITO)
   
End Function

Function CONVIERTOCODIGOBARRAS$(YY$)

    Codigo$ = YY$
    XX$ = "("
    
    For J = 1 To Len(Codigo$) Step 2
        NUMERO = XVALO(Mid$(Codigo$, J, 2))
        i = Int(NUMERO / 10)
        d = NUMERO Mod 10
        If NUMERO < 50 Then
            XX$ = XX$ + Chr$(d + 10 * i + 48)
        Else
            XX$ = XX$ + Chr$(192 + d + 10 * (i - 5))
        End If
    Next J
    
    XX$ = XX$ + ")"
    CONVIERTOCODIGOBARRAS$ = XX$

End Function

Function CODIBARRA$(XX$, Optional MODO%)
    '
    YY$ = ""
    If TRIM$(XX$) = "" Then GoTo 98
       
    If MODO% = 1 Then
        YY$ = CONVIERTOCODIGOBARRAS$(XX$)       ' INTERLEAVED 2 DE 5. SOLO NUMEROS. UTILIZADO PARA AFIP
    ElseIf MODO% = 2 Then
        YY$ = CODIBARRA_BC128$(XX$)             ' BCODE 128. ALFANUMERICO
    ElseIf MODO% = 3 Then
        YY$ = CODIBARRA_EAN13$(XX$)             ' EAN 13 NUMERICO
    ElseIf MODO% = 4 Then
        YY$ = CODIBARRA_CODE128$(XX$)           ' CODE 128. ALFANUMERICO
    End If
    '
98  CODIBARRA$ = YY$

'    If CONTROL$("", "CODIBAR") = "COBAR128" Then
'        YY$ = CODIBARRA_BC128$(XX$)
'        GoTo 98
'    End If
'
'    If TRIM$(XX$) = "" Then Exit Function
'
'    ' Convierte al formato del codigo de barras
'    YY$ = CONVIERTOCODIGOBARRAS$(XX$)
'    '
'98  CODIBARRA$ = YY$
    '
End Function

'
Function CODIBARRA_CODE128$(XX$)
  
 'This function is governed by the GNU Lesser General Public License (GNU LGPL)
  'V 2.0.0
  'Parameters : a string
  'Return : * a string which give the bar code when it is dispayed with CODE128.TTF font
  '         * an empty string if the supplied parameter is no good
  Dim i%, checksum&, MINI%, dummy%, tableB As Boolean
  chaine$ = XX$
  CODIBARRA_CODE128$ = ""
  If Len(chaine$) > 0 Then
  'Check for valid characters
    For i% = 1 To Len(chaine$)
      Select Case Asc(Mid$(chaine$, i%, 1))
      Case 32 To 126, 203
      Case Else
        i% = 0
        Exit For
      End Select
    Next
    'Calculation of the code string with optimized use of tables B and C
    CODIBARRA_CODE128$ = ""
    tableB = True
    If i% > 0 Then
      i% = 1 'i% become the string index
      Do While i% <= Len(chaine$)
        If tableB Then
          'See if interesting to switch to table C
          'yes for 4 digits at start or end, else if 6 digits
          MINI% = IIf(i% = 1 Or i% + 3 = Len(chaine$), 4, 6)
          GoSub testnum
          If MINI% < 0 Then 'Choice of table C
            If i% = 1 Then 'Starting with table C
              CODIBARRA_CODE128$ = Chr$(210)
            Else 'Switch to table C
              CODIBARRA_CODE128$ = CODIBARRA_CODE128$ & Chr$(204)
            End If
            tableB = False
          Else
            If i% = 1 Then CODIBARRA_CODE128$ = Chr$(209) 'Starting with table B
          End If
        End If
        If Not tableB Then
          'We are on table C, try to process 2 digits
          MINI% = 2
          GoSub testnum
          If MINI% < 0 Then 'OK for 2 digits, process it
            dummy% = Val(Mid$(chaine$, i%, 2))
            dummy% = IIf(dummy% < 95, dummy% + 32, dummy% + 105)
            CODIBARRA_CODE128$ = CODIBARRA_CODE128$ & Chr$(dummy%)
            i% = i% + 2
          Else 'We haven't 2 digits, switch to table B
            CODIBARRA_CODE128$ = CODIBARRA_CODE128$ & Chr$(205)
            tableB = True
          End If
        End If
        If tableB Then
          'Process 1 digit with table B
          CODIBARRA_CODE128$ = CODIBARRA_CODE128$ & Mid$(chaine$, i%, 1)
          i% = i% + 1
        End If
      Loop
      'Calculation of the checksum
      For i% = 1 To Len(CODIBARRA_CODE128$)
        dummy% = Asc(Mid$(CODIBARRA_CODE128$, i%, 1))
        dummy% = IIf(dummy% < 127, dummy% - 32, dummy% - 105)
        If i% = 1 Then checksum& = dummy%
        checksum& = (checksum& + (i% - 1) * dummy%) Mod 103
      Next
      'Calculation of the checksum ASCII code
      checksum& = IIf(checksum& < 95, checksum& + 32, checksum& + 105)
      'Add the checksum and the STOP
      CODIBARRA_CODE128$ = CODIBARRA_CODE128$ & Chr$(checksum&) & Chr$(211)
    End If
  End If
  Exit Function
testnum:
  'if the mini% characters from i% are numeric, then mini%=0
  MINI% = MINI% - 1
  If i% + MINI% <= Len(chaine$) Then
    Do While MINI% >= 0
      If Asc(Mid$(chaine$, i% + MINI%, 1)) < 48 Or Asc(Mid$(chaine$, i% + MINI%, 1)) > 57 Then Exit Do
      MINI% = MINI% - 1
    Loop
  End If
Return
'    Dim lcStart, lcStop, lcRet$, lcCheck
'    Dim lnLong, lnI%, lnCheckSum, lnAsc, AA
'    '
'    tcString = XX$
'    lcStart = Chr(104 + 32)
'    lcStop = Chr(106 + 32)
'    lnCheckSum = Asc(lcStart) - 32
'    lcRet = tcString
'    lnLong = Len(lcRet)
'    For lnI = 1 To lnLong
'        lnAsc = Asc(Mid(lcRet, lnI, 1)) - 32
'        If Not lnAsc >= 0 And lnAsc <= 99 Then
'            lcRet = REPLACE$(lcRet, Chr(32), lnI, 1)
'            lnAsc = Asc(Mid(lcRet, lnI, 1)) - 32
'        End If
'        lnCheckSum = lnCheckSum + (lnAsc * lnI)
'    Next
'    AA = lnCheckSum Mod 103
'    AA = AA + 32
'    lcCheck = Chr(AA)
'    lcRet = lcStart + lcRet + lcCheck + lcStop
'    '*--- Esto es para cambiar los espacios y caracteres invalidos
'    lcRet = REPLACAR$(lcRet, Chr(32), Chr(232))
'    lcRet = REPLACAR$(lcRet, Chr(127), Chr(192))
'    lcRet = REPLACAR$(lcRet, Chr(128), Chr(193))
'    '*---
'    CODIBARRA_CODE128$ = lcRet
'    '
End Function

Function CODIBARRA_BC128$(XX$)
  
    Dim lcStart, lcStop, lcRet$, lcCheck
    Dim lnLong, lnI%, lnCheckSum, lnAsc, AA
    '
    tcString = XX$
    lcStart = Chr(104 + 32)
    lcStop = Chr(106 + 32)
    lnCheckSum = Asc(lcStart) - 32
    lcRet = tcString
    lnLong = Len(lcRet)
    For lnI = 1 To lnLong
        lnAsc = Asc(Mid(lcRet, lnI, 1)) - 32
        If Not lnAsc >= 0 And lnAsc <= 99 Then
            lcRet = REPLACE$(lcRet, Chr(32), lnI, 1)
            lnAsc = Asc(Mid(lcRet, lnI, 1)) - 32
        End If
        lnCheckSum = lnCheckSum + (lnAsc * lnI)
    Next
    AA = lnCheckSum Mod 103
    AA = AA + 32
    lcCheck = Chr(AA)
    lcRet = lcStart + lcRet + lcCheck + lcStop
    '*--- Esto es para cambiar los espacios y caracteres invalidos
    lcRet = REPLACAR$(lcRet, Chr(32), Chr(232))
    lcRet = REPLACAR$(lcRet, Chr(127), Chr(192))
    lcRet = REPLACAR$(lcRet, Chr(128), Chr(193))
    '*---
    CODIBARRA_BC128$ = lcRet
    '
End Function

Function CODIBARRA_EAN13$(XX$)

  EANLen% = 12
  CodeClair$ = Mid$(XX$, 1, 12) & String$(EANLen% - Len(TRIM$(Mid$(XX$, 1, 12))), "0")
  chaine$ = CodeClair$
  '
  'Cette fonction est régie par la Licence Générale Publique Amoindrie GNU (GNU LGPL)
  'This function is governed by the GNU Lesser General Public License (GNU LGPL)
  'V 1.1.1
  'Paramètres : une chaine de 12 chiffres
  'Parameters : a 12 digits length string
  'Retour : * une chaine qui, affichée avec la police EAN13.TTF, donne le code barre
  '         * une chaine vide si paramètre fourni incorrect
  'Return : * a string which give the bar code when it is dispayed with EAN13.TTF font
  '         * an empty string if the supplied parameter is no good
  Dim i%, checksum%, first%, CodeBarre$, tableA As Boolean
  CODIBARRA_EAN13$ = ""
  'Vérifier qu'il y a 12 caractères
  'Check for 12 characters
  If Len(chaine$) = 12 Then
    'Et que ce sont bien des chiffres
    'And they are really digits
    For i% = 1 To 12
      If Asc(Mid$(chaine$, i%, 1)) < 48 Or Asc(Mid$(chaine$, i%, 1)) > 57 Then
        i% = 0
        Exit For
      End If
    Next
    If i% = 13 Then
      'Calcul de la clé de contrôle
      'Calculation of the checksum
      For i% = 12 To 1 Step -2
        checksum% = checksum% + Val(Mid$(chaine$, i%, 1))
      Next
      checksum% = checksum% * 3
      For i% = 11 To 1 Step -2
        checksum% = checksum% + Val(Mid$(chaine$, i%, 1))
      Next
      chaine$ = chaine$ & (10 - checksum% Mod 10) Mod 10
      'Le premier chiffre est pris tel quel, le deuxième vient de la table A
      'The first digit is taken just as it is, the second one come from table A
      CodeBarre$ = Left$(chaine$, 1) & Chr$(65 + Val(Mid$(chaine$, 2, 1)))
      first% = Val(Left$(chaine$, 1))
      For i% = 3 To 7
        tableA = False
         Select Case i%
         Case 3
           Select Case first%
           Case 0 To 3
             tableA = True
           End Select
         Case 4
           Select Case first%
           Case 0, 4, 7, 8
             tableA = True
           End Select
         Case 5
           Select Case first%
           Case 0, 1, 4, 5, 9
             tableA = True
           End Select
         Case 6
           Select Case first%
           Case 0, 2, 5, 6, 7
             tableA = True
           End Select
         Case 7
           Select Case first%
           Case 0, 3, 6, 8, 9
             tableA = True
           End Select
         End Select
       If tableA Then
         CodeBarre$ = CodeBarre$ & Chr$(65 + Val(Mid$(chaine$, i%, 1)))
       Else
         CodeBarre$ = CodeBarre$ & Chr$(75 + Val(Mid$(chaine$, i%, 1)))
       End If
     Next
      CodeBarre$ = CodeBarre$ & "*"   'Ajout séparateur central / Add middle separator
      For i% = 8 To 13
        CodeBarre$ = CodeBarre$ & Chr$(97 + Val(Mid$(chaine$, i%, 1)))
      Next
      CodeBarre$ = CodeBarre$ & "+"   'Ajout de la marque de fin / Add end mark
      CODIBARRA_EAN13$ = CodeBarre$
    End If
  End If

End Function
Sub PINTATEXT(TXT As TextBox)
   TXT.SelStart = 0
   'Selecciona Todo el contenido de la caja de texto
   If Len(TXT.TEXT) = 0 Then TXT = " "
   TXT.SelLength = Len(TXT.TEXT)
   On Error Resume Next
   TXT.SetFocus
   On Error GoTo 0
End Sub
Sub PINTATEXTCMB(CMB As comboBox)
   CMB.SelStart = 0
   'Selecciona Todo el contenido de la caja de texto
   If Len(CMB.TEXT) = 0 Then TXT = "     "
   CMB.SelLength = Len(CMB.TEXT)
   On Error Resume Next
   CMB.SetFocus
   On Error GoTo 0
End Sub

Function CEROIGUALVACIO(NUME)
   If XVALO(NUME) < 1 Then
      CEROIGUALVACIO = ""
   Else
      CEROIGUALVACIO = NUME
   End If
End Function

Function TEXTOLOTES$(TEXTO1$, SEPARADOR$, Optional SEPARADOR_CIERRE$)
      ' FUNCIONA  PARA TEXTOS SEPARADOS POR UN CARACTER, DEVOLVIENDO EN UN  VECTOR LOS TROZOS LEIDOS EN UN VECTOR
      'FUNCIONA PARA TEXTOS SEPARADOS CON UN SEPARADOR INCIAL Y UN SEPARADOR DE CIERRE EJEMPLO: [| Y |]
      'VALIDACIONES
      If TRIM$(TEXTO1$) = "" Then
        ReDim Preserve CADENATEX$(0)
        Exit Function
      End If
      If InStr(TEXTO1$, SEPARADOR$) < 1 Then
        TEXTOLOTES$ = TEXTO1$
        ReDim Preserve CADENATEX$(1)
        CADENATEX$(1) = TEXTO1$ 'AGREGADO PARA QUE SI NO TIENE SEPARADOR EN EL VECTOR INDICE 1 DEVUELVA EL TEXTO COMPLETO
        Exit Function
      End If
      '
      'PROCESOS
      i% = 0
      textox$ = TEXTO1$
      'Debug.Print TEXTOX$
20    POS1% = InStr(textox$, SEPARADOR$) 'POSICION INICIAL DE PRIMER SEPARADOR
      CNTCARACT1% = Len(SEPARADOR$) 'POR SI TIENE MAS DE UN CARACTER EL SEPARADOR
      'VALIDA SI TIENE SEPARADOR DE CIERRE
      If SEPARADOR_CIERRE$ <> "" Then
         TEXTOINTERNO$ = Mid$(textox$, POS1% + CNTCARACT1%)
         POS2% = InStr(TEXTOINTERNO$, SEPARADOR_CIERRE$)
         CNTCARACT2% = Len(SEPARADOR_CIERRE$)
      End If
      If POS1% > 0 Then
        i% = i% + 1
        TROZO$ = Left$(textox$, POS1% - 1)
        If POS2% > 0 Then
           TROZO$ = Mid$(textox$, POS1% + CNTCARACT1%, POS2% - 1)
        End If
        ReDim Preserve CADENATEX$(i%)
        CADENATEX$(i%) = TROZO$
        'Debug.Print TEXTOX$
        'Debug.Print TROZO$
        textox$ = Mid$(textox$, POS1% + 1)
        If POS2% > 0 Then
           textox$ = Mid$(TEXTO1$, POS1% + 1 + POS2% + CNTCARACT2%)
        End If
        'SI EL TEXTO ESTA VACIO SE VA
        If TRIM$(textox$) = "" Then Exit Function ' SI ESTA VACIO EL RESTO SE VA
        GoTo 20 ' SI NO VUELVE A ARRIBA
      End If
      i% = i% + 1
      'EN CASO QUE HAYA SEPARADOR DE CIERRE NO DEVUELVE EL TEXTO FINAL Y SE VA
      If SEPARADOR_CIERRE$ <> "" Then Exit Function
      ReDim Preserve CADENATEX$(i%)
      CADENATEX$(i%) = textox$
      
      
End Function

    Sub POSICIONARALFINALTXT(TXTCAJA As TextBox)
       On Error Resume Next
       TXTCAJA.SelStart = Len(TXTCAJA)
       On Error GoTo 0
    End Sub
    
Sub LIMPIARTEXT(FRM As Form, LIMPIOTEXT%, Optional SOLO_LOCKED_TRUE%)
  Dim CONTROLX As Control
  For Each CONTROLX In FRM.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
        If TypeOf CONTROLX Is TextBox And LIMPIOTEXT% > 0 Then
          If SOLO_LOCKED_TRUE% > 0 Then
            If CONTROLX.Locked = True Then
              CONTROLX.TEXT = ""
            End If
          Else
           CONTROLX.TEXT = ""
           
          End If
        End If
  Next
End Sub
Sub LIMPIARTEXTSEGUNCOLOR(FRM As Form, LIMPIOTEXT%, Optional COLOR1$, Optional COLOR2$)
  Dim CONTROLX As Control
  For Each CONTROLX In FRM.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
        If TypeOf CONTROLX Is TextBox And LIMPIOTEXT% > 0 Then
          If CONTROLX.BackColor = COLOR1$ Or CONTROLX.BackColor = COLOR2$ Then
              CONTROLX.TEXT = ""
          End If
        End If
  Next
  
End Sub

Sub LIMPIAROPTION(FRM As Form)
  Dim CONTROLX As Control
  For Each CONTROLX In FRM.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
        If TypeOf CONTROLX Is OptionButton Then
              CONTROLX.Value = False
        End If
  Next
End Sub
Sub LIMPIARCHECK(FRM As Form)
  Dim CONTROLX As Control
  For Each CONTROLX In FRM.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
        If TypeOf CONTROLX Is CheckBox Then
              CONTROLX.Value = 0
        End If
  Next
End Sub
Sub PorcionesDeTexto(TXT$, sep$)
   Dim strTest As String
   Dim strArray() As String
   Dim intCount As Integer
   
   strArray = Split(TXT$, sep$)
   Text2 = ""
   For i& = LBound(strArray) To UBound(strArray)
     If i& = 0 Then
      Text2 = Text2 + "SQLX$ = " & Chr(34) & TRIM(strArray(i&)) & Chr(34) & vbCrLf
     Else
      Text2 = Text2 + "SQLX$ = SQLX$ + " & Chr(34) & TRIM(strArray(i&)) & Chr(34) & vbCrLf
     End If
   Next i&
End Sub

Function TextoEnFormatoUTF8(TextoOriginal$)
    
    For i& = 1 To Len(TextoOriginal$)
        LETRA$ = Mid$(TextoOriginal$, i&, 1)
        If (Asc(LETRA$) >= 65 And Asc(LETRA$) <= 90) Or (Asc(LETRA$) >= 97 And Asc(LETRA$) <= 122) Or LETRA$ = " " Or LETRA$ = "." Or LETRA$ = "-" Or LETRA$ = "," Then
            TextoConvertido$ = TextoConvertido$ & LETRA$
        End If
    Next i&
    
    TextoEnFormatoUTF8 = TextoConvertido$
    
End Function
Function FORMATOCON0$(NUM&, CANTI_0%)
      'SE LE PASA UN NUMERO Y LA CANTIDAD DE 0 A COMPLETAR A IZQUIERDA
      CAN0$ = String$(CANTI_0%, "0")
      NUCL$ = Right(CAN0$ & TRIM$(Str(NUM&)), CANTI_0%)
      FORMATOCON0$ = NUCL$
   
End Function
Public Function FORMATO_RTO_WALDB(REMILAR$)
   'CAMBIA EL FORMATO RT-X-0001-00002286 A 000100002286
   FORMATW$ = Mid$(REMILAR$, 6, 4)
   FORMATW$ = FORMATOCON0$(XVALO(FORMATW$), 5) 'le agrego un 0 mas ahora es con 5 caracteres el prefijo
   FORMATW$ = FORMATW$ + Mid$(REMILAR$, 11, 8)
   FORMATO_RTO_WALDB = FORMATW$
End Function
Function LetraProvArbaSegunLetraFlexoft$(LetraFlex$)
   'EL RESTO COINCIDE CON LA TABLA FLEXOFT.
   Select Case LetraFlex$
     Case "U": LetraFlex$ = "Q"
      Case "S": LetraFlex$ = "A"
       Case "L": LetraFlex$ = "D"
        Case "W": LetraFlex$ = "K"
         Case "F": LetraFlex$ = "V"
          Case "O": LetraFlex$ = "N"
           Case "I": LetraFlex$ = "W"
            Case "A": LetraFlex$ = "H"
             Case "N": LetraFlex$ = "P"
              Case "P": LetraFlex$ = "L"
               Case "K": LetraFlex$ = "F"
                Case "Q": LetraFlex$ = "S"
   End Select
   LetraProvArbaSegunLetraFlexoft$ = LetraFlex$

End Function
Function UnidMedidaArbaSegunUnidFlexoft$(UniFlex$, RetornoNum%)
   UniFlex1$ = UCase$(TRIM$(UniFlex$))
   Select Case UniFlex1$
     Case "U.": UniFlex$ = "UNIDADES": RETORNO% = 3
      Case "KG": UniFlex$ = "Kilogramos": RETORNO% = 1
       Case "M.": UniFlex$ = "Metros": RETORNO% = 5
        Case "M2": UniFlex$ = "Metros Cuadrados": RETORNO% = 4
         Case "M3": UniFlex$ = "Metros Cúbicos": RETORNO% = 6
          Case "LT": UniFlex$ = "Litros": RETORNO% = 2
       Case Else
          UniFlex$ = "UNIDADES": RETORNO% = 3
   End Select
   UnidMedidaArbaSegunUnidFlexoft$ = UniFlex$
   RetornoNum% = RETORNO%
End Function
Function FORMATOSERIEORD$(ORD%)
   
   ORDEX$ = SAFETEXT$(ORD%)
   'LEE LA CONFIGURACION DE LA TABLA FORMANSERIE
   If EXISTE_TABLA%("FORMANSERIE") < 1 Then Exit Function
   
   SQLX$ = "SELECT * FROM FORMANSERIE WITH(NOLOCK) WHERE CORRESPONDE_A= 'NUMEROS DE SERIE' AND REFER_DATO = 'ORDEN'"
   Set rs = READSET(SQLX$)
   With rs
     XANTICARACTERES% = XVALO(!CANTICARACT)
     COMPLETACOX$ = TRIM$(SAFETEXT$(!COMPLETACON))
     EXTREMOX$ = TRIM$(SAFETEXT$(!EXTREMO$))
   End With
   rs.Close
   Set rs = Nothing
   '
   'CON LOS DATOS FORMATEA
   If EXTREMOX$ = "" Then EXTREMOX$ = "I"
   If XANTICARACTERES% > 0 Then
     If COMPLETACOX$ <> "" Then
        If EXTREMOX$ = "I" Then
           While Len(ORDEX$) < XANTICARACTERES%: ORDEX$ = COMPLETACOX$ & ORDEX$: Wend
        Else
           While Len(ORDEX$) < XANTICARACTERES%: ORDEX$ = ORDEX$ & COMPLETACOX$: Wend
        End If
     End If
   End If
   '
   FORMATOSERIEORD$ = ORDEX$
End Function


Function FORMATOCONMASCARA$(NSERIE$)
   
   FORMATOCONMASCARA$ = NSERIE$ 'POR  DEFECTO
   'LEE LA CONFIGURACION DE LA TABLA FORMANSERIE
   If EXISTE_TABLA%("FORMANSERIE") < 1 Then Exit Function
   
   SQLX$ = "SELECT * FROM FORMANSERIE WITH(NOLOCK) WHERE CORRESPONDE_A= 'NUMEROS DE SERIE' AND REFER_DATO = 'MASCARA'"
   Set rs = READSET(SQLX$)
   With rs
     MASX$ = TRIM$(SAFETEXT$(!MASCARA))
   End With
   rs.Close
   Set rs = Nothing
   '
   If MASX$ = "" Then Exit Function
   LONGX% = Len(MASX$)
   For i% = 1 To LONGX%
     LEO$ = Mid$(MASX$, i%, 1)
     'ABCDEFG
     'X-XX-X-XX-X
     If LEO$ <> "X" Then
        PARTELEIDA$ = Left(NSERIE$, i% - 1)
        PARTERESTANTE$ = Mid$(NSERIE$, i%)
        FORMATOCONMASCARA$ = PARTELEIDA$ & LEO$ & PARTERESTANTE$
        NSERIE$ = FORMATOCONMASCARA$
        J% = J% + 1
     End If
   Next
End Function
Function TEXTOMENSACTRLSCARGO$(CTRLSINCARGO%)
    '0= TODO BIEN
    '1= TIENE ITEMS SIN CARGO PERO ESTAN APROBADOS
    '-1 TIENE ITEMS SIN CARGO QUE ESTAN EN LA LISTA Y NECESITAN APROBACION
    '-2 TIENE ITEMS  QUE NO ESTAN EN LA LISTA Y LA DIFERENCIA ES MAYOR A LA TOLERADA CONTRA LA LISTA
    '3 TIENE ITEMS CON VALOR 0 QUE NO ESTAN EN LA LISTA (SE CONSIDERA POR SI NO HAY PORCENTAJE DE TOLERANCIA
    Select Case CTRLSINCARGO%
       Case -1
        TEXTOMENSACTRLSCARGO$ = "Items Configurados Sin Cargo, Con Aprobación Posterior."
       Case -2
        TEXTOMENSACTRLSCARGO$ = "Items con Diferencia de Precio de Lista Que superan el Porcentaje de Tolerancia."
       Case -3
        TEXTOMENSACTRLSCARGO$ = "Items Sin Cargo No Configurados."
    End Select
End Function
Sub APLICAR_REVISION_FORMATEADA(MSF As msFlexGrid, COL%)
   'LEE EL MSF Y SEGUN LA POSICION ACTUALIZA CADA CELDA
   For i% = 1 To MSF.Rows - 1
      LEO$ = MSF.TextMatrix(i%, COL%)
      MSF.TextMatrix(i%, COL%) = FormatoEYMRevision(LEO$)
   Next i%
End Sub


Function FormatoEYMRevision$(REVISIOX$)
   'LE DA FORMATO A LA REVISION TOMA SOLO LOS NUMEROS
   'EJ SI ES AB 12 , LA FUNCION DEVUELVE 12, RECORRE DE DERECHA A IZQUIERDA HASTA QUE ENCUENTRA UNA LETRA
   'SI EL NUMERO ESTA SEPARADO POR ESPACIOS LO CONSIDERA TAMBIEN SOLO CORTA CUANDO ENCUENTRA UNA LETRA
   
   REVISIOX$ = TRIM$(REVISIOX$)
   For i% = 1 To Len(REVISIOX$)
      CARACT_LEIDO$ = Right(REVISIOX$, 1)
      If IsNumeric(CARACT_LEIDO$) Or Asc(CARACT_LEIDO$) = 32 Then
         FormatoEYMRevision$ = TRIM$(CARACT_LEIDO$) & FormatoEYMRevision$
         REVISIOX$ = Mid$(REVISIOX$, 1, Len(REVISIOX$) - 1)
      Else
         Exit For
      End If
   Next i%
   FormatoEYMRevision$ = TRIM$(FormatoEYMRevision$)
End Function
Function FORMATOBR$(TEXBR$, Optional CI1%)
  TEXBR$ = UCase$(TEXBR$)
  FORMATOBR$ = TEXBR$
  'SI A LA PARTE CENTRAL ENTRE LOS GUIONES LE FALTAN CEROS PARA COMPLETAR LOS 6 DIGITOS, LOS AGREGA
  'SE CAMBIO POR QUE EL INGRESO EN RECEPCION UTILIZA EJ.: BR-234-1
  If Left(TEXBR$, 3) = "BR-" Or Left(TEXBR$, 3) = "PR-" Or Left(TEXBR$, 3) = "OF-" Then
     On Error Resume Next
     POSLEFT% = InStr(TEXBR$, "-")
     POSRIGHT% = InStrRev(TEXBR$, "-")
     TXCEN$ = Mid(TEXBR$, POSLEFT% + 1, POSRIGHT% - (POSLEFT% + 1))
     If Left(TEXBR$, 3) = "BR-" Or Left(TEXBR$, 3) = "PR-" Then
        FORMATOBR$ = "BR-" & SAFETEXT$(XVALO(TXCEN$)) & Mid$(TEXBR$, POSRIGHT%)
     Else
        FORMATOBR$ = "OF-" & FORMATOCON0$(XVALO(TXCEN$), 6) & Mid$(TEXBR$, POSRIGHT%) 'LA OF SI SE COMPLETA CON 0
     End If
     On Error GoTo 0
  ElseIf Left(TEXBR$, 2) = "BR" Or Left(TEXBR$, 2) = "PR" Then 'EJ. SI TIENE FORMATO BR125 DEVUELVE BR-125-1
    LEIDORECE$ = Mid$(TEXBR$, 3)
    If IsNumeric(LEIDORECE$) Then
       If CI1% > 0 Then
             NUMERECE = XVALO(LEIDORECE$)
             SQLX$ = "SELECT * FROM bolrecep WITH(NOLOCK) WHERE NUMEBORE = " & NUMERECE
             SQLX$ = SQLX$ + " AND COD_INTE = " & CI1%
             Set RSS = READSET(SQLX$)
             With RSS
               If Not .EOF Then
                  LOTEBOLRECEP$ = SAFETEXT$(!idenlote)
                  FORMATOBR$ = LOTEBOLRECEP$
               End If
             End With
             RSS.Close: Set RSS = Nothing
        End If
   ElseIf Left(TEXBR$, 2) = "OF" Then 'EJ.
        LEIDORECE$ = Mid$(TEXBR$, 3)
        If IsNumeric(LEIDORECE$) Then
             NUMERECE = XVALO(LEIDORECE$)
             FORMATOBR$ = "OF-" & FORMATOCON0$(XVALO(LEIDORECE$), 6) & "-1" 'LA OF SI SE COMPLETA CON 0
        End If
    End If
  End If
  
End Function
Function FORMATO_BR_OF_SIN_0$(LOTE$)
    'LE QUITA LOS CEROS DE LA PARTE CENTRAL
    'EJ: de OF-002220-1 queda OF-220-1 'realizado para eym
    FORMATO_BR_OF_SIN_0$ = LOTE$
    If Left(LOTE$, 3) = "BR-" Or Left(LOTE$, 3) = "OF-" Then
     On Error Resume Next
     POSLEFT% = InStr(LOTE$, "-")
     POSRIGHT% = InStrRev(LOTE$, "-")
     TXCEN$ = Mid(LOTE$, POSLEFT% + 1, POSRIGHT% - (POSLEFT% + 1))
     For i% = 1 To Len(TXCEN$)
        If XVALO(Mid$(TXCEN$, i%, 1)) > 0 Then
           TXCEN$ = Mid$(TXCEN$, i%)
           Exit For
        End If
     Next i%
     FORMATO_BR_OF_SIN_0$ = Mid$(LOTE$, 1, 3) & TXCEN$ & Mid$(LOTE$, POSRIGHT%)
     On Error GoTo 0
    End If
End Function

Sub LIMPIAR_CONTROL_ENCONTENEDOR(FORMULARIO As Form, OBJETOCONTENEDOR As Object, Optional COLOR1$, Optional COLOR2$)
 
    Dim CONTROLX As Control
    For Each CONTROLX In FORMULARIO.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
          CONTENEDORPARAM = OBJETOCONTENEDOR.Name
          On Error Resume Next
          CONTENEDORLEIDO = CONTROLX.Container.Name
          If Err <> 0 Then
             On Error GoTo 0
             GoTo 30
          End If
          If CONTENEDORPARAM = CONTENEDORLEIDO Then 'CONTROLA SI ESTA DENTRO DEL CONTENEDOR PASADO COMO PARAMETRO
            If TypeOf CONTROLX Is TextBox Then
              If TRIM(COLOR1$) <> "" Then
                 If CONTROLX.BackColor = COLOR1$ Or CONTROLX.BackColor = COLOR2$ Then
                   CONTROLX.TEXT = ""
                End If
              End If
              Else
               CONTROLX.TEXT = ""
              End If
         End If
         
30   Next
End Sub

Public Sub Centrar(ByVal poForm As Form)
    With poForm
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
     End With
End Sub

Sub AJUSTARFRM(FRM As Form, ANCHO, ALTO)
    If ANCHO > 0 Then FRM.Width = ANCHO
    If ALTO > 0 Then FRM.Height = ALTO
    Call Centrar(FRM)
End Sub
Sub PRINT_NOMBRE_TEXTBOX_X_COLOR(FRM As Form, Optional COLOR1$, Optional COLOR2$)

  Dim CONTROLX As Control
  For Each CONTROLX In FRM.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
        If TypeOf CONTROLX Is TextBox Then
          If CONTROLX.BackColor = COLOR1$ Or CONTROLX.BackColor = COLOR2$ Then
              CONTROLX.TEXT = CONTROLX.Name
          End If
        End If
  Next
  
End Sub
Function CenterString(ByVal sText As String, ByVal nWidth As Integer, Optional CaracteresPrincipioyFin$)
    Dim nSpaces As Integer
    Dim sCenteredText As String
    'nWidth es la cantidad de caracteres que deberia tener de ancho
'CENTRAR UN TEXTO COMPLETANDO EL RESTO CON ESPACIOS, SI SE PASA CaracteresPrincipioyFin$ LO AGREGA AL PRINCIPIO Y AL FIN
    ' Calcula la cantidad de espacios en blanco necesarios para centrar el texto
    nSpaces = (nWidth - Len(sText)) \ 2

    ' Construye el string centrado
    sCenteredText = String(nSpaces, " ") & sText & String(nSpaces, " ")

    If Len(sText) <= (nWidth - 2) And TRIM$(CaracteresPrincipioyFin$) <> "" Then
       sCenteredText = Right(sCenteredText, Len(sCenteredText) - 1)
       sCenteredText = CaracteresPrincipioyFin$ & sCenteredText
       If (Len(sText) Mod 2) = 0 Then 'SI EL TEXTO ES PAR LE QUITA 1 Y REEMPLAZA POR EL CARACTER
          sCenteredText = Left(sCenteredText, Len(sCenteredText) - 1)
       Else 'CASO CONTRARIO LE AGREGA EL CARACTER SIN QUITAR NINGUN ESPACIO
          sCenteredText = Left(sCenteredText, Len(sCenteredText))
       End If
       sCenteredText = sCenteredText & CaracteresPrincipioyFin$
      
    End If
    ' Imprime el texto centrado
    Debug.Print sCenteredText
    CenterString = sCenteredText
End Function

Function NumeroEstaEnCadena(ByVal cadena As String, ByVal NUMERO As Integer) As Boolean
    Dim numeros() As String
    Dim i As Integer
    Dim numeroConComillas As String
    Dim numeroSinComillas As String
    ' Dividir la cadena en un arreglo utilizando la coma como separador
    numeros = Split(cadena, ",")
    
    ' Recorrer el arreglo para buscar el número
    For i = 0 To UBound(numeros)
        numeroConComillas = numeros(i) ' Número entre comillas simples
        numeroSinComillas = REPLACAR(numeroConComillas, "'", "") ' Elimina las comillas simples

        If XVALO(numeroSinComillas) = NUMERO Then
        
            NumeroEstaEnCadena = True
            Exit Function
        End If
       
    Next i
    
    ' Si llegamos hasta aquí, el número no se encontró en la cadena
    NumeroEstaEnCadena = False
End Function
