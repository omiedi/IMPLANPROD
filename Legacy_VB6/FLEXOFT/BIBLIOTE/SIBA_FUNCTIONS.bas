Attribute VB_Name = "SIBA_FUNCTIONS"
Public CLAVE As String * 12
Public CLAV As String * 12
Public CONTRA As String * 12
Public QWERTY%
Public NUSER%
Dim DENUS$

'

Sub DATSYS(NSX$, EMX$, SCX$, PSX$, FLX$, DSX$, FLICX%)
        '
        Static CTX%, NS$, EM$, SC$, PS$, FL$, DS$, FLIC%
        '
        If CTX% <> 1 Then
            CTX% = 1
            '
            VERCD% = 0
            APRUEBA$ = "NULLPRUE.$$$"
            Dim RPRU As String * 128
            NX% = FreeFile
            On Error Resume Next
            Open APRUEBA$ For Random Access Read Write Shared As #NX% Len = 128
            RPRU$ = String$(128, "*")
            Put #NX%, 1, RPRU$
            Close #NX%: NX% = 0
            If Err Then
               VERCD% = 1
               LUDAT$ = "C:\FLEXOFT\DATOS\"
            End If
            On Error GoTo 0
            '
            If VERCD% = 1 Then GoTo 20
            '
            If EXISTE%("IDENSYS.DRV") > 0 Then
10              Dim IDS As String * 96
                n1% = FILEOPEN%("IDENSYS.DRV", 9, 96)
                If LOF(n1%) > 0 Then
                    Get #n1%, 1, IDS$
                    DEXX$ = DESENCRI$(Left$(IDS$, 94))
                    '
                    Get #n1%, 2, IDS$
                    DF$ = DESENCRI$(Left$(IDS$, 94))
                  Else
                    Close #n1%: n1% = 0
                    Call CARDASYS
                    GoTo 10
                End If
                Close #n1%: n1% = 0
            End If
            '
20          NS$ = "000000-00"
            EM$ = "Su Empresa"
            SC$ = "Sistema de Demostración"
            PS$ = "Usuario Demo"
            FL$ = "S/Limite"
            FLIC% = HOY(HO$)
            '
            If DEXX$ <> "" Then
                NS$ = TRIM$(Left$(DEXX$, 12))
                EM$ = TRIM$(Mid$(DEXX$, 13, 30))
                SC$ = TRIM$(Mid$(DEXX$, 43, 20))
                PS$ = TRIM$(Mid$(DEXX$, 63, 30))
                   FLIM = CVI(Mid$(DEXX$, 93, 2))
                FL$ = FECHATEX$(FLIM)
            End If
            '
            DS$ = "SISTEMA PRECONFIGURADO FLEXOFT (r)"
            If DF$ <> "" Then
                DS$ = TRIM$(Mid$(DF$, 11, 40))
                FLIC% = CVI(Left$(DF$, 2))
            End If
            '
        End If
        '
        Dim N0S As String * 128
        n0% = FILEOPEN%("\IEEE.DRV", 9, 128)
        If LOF(n0%) < 128 Then
          If VERDEMO% < 1 Then
            Close #n0%: n0% = 0
            n0% = FILEOPEN%("\IEEE.DRV", 0, 128)
            FLI1% = HOY(HOS$) + 100
            N0SX$ = ENCRIPTA$(Space$(128))
            Call REPLA(N0SX$, MKI$(FLI1%), 45, 2)
            N0S$ = N0SX$
            Put #n0%, 1, N0S$
            N0SX$ = ENCRIPTA$(String$(128, 45))
            Call REPLA(N0SX$, MKI$(9876 - FLI1%), 73, 2)
            N0S$ = N0SX$
            Put #n0%, 2, N0S$
          End If
        End If
        '
        Get #n0%, 1, N0S$: FLIM = CVI(Mid$(N0S$, 45, 2))
        Get #n0%, 2, N0S$: FLXX = CVI(Mid$(N0S$, 73, 2))
        Close #n0%: n0% = 0
        If (FLIM + FLXX) <> 9876 Then
            Call FUERASYS
        End If
        '
        FX = FLIM: FL$ = FECHATEX$(FX)
        If FL$ = "31/12/99" Then FL$ = "S/Limite"
        '
        If FLIC% < 8000 Or FLIC% > HOY(HO$) Then
            FLIC% = HOY(HO$)
        End If
        '
        NSX$ = NS$
        EMX$ = EM$
        SCX$ = SC$
        PSX$ = PS$
        FLX$ = FL$
        DSX$ = DS$
        FLICX% = FLIC%
        '
End Sub
'

Sub CARDASYS()
    '
    Randomize Timer
    INTEN% = 0
    '
5   CLAX$ = ""
10  PPX% = Int(Rnd * 63)
    If PPX% < 48 Or PPX% > 57 Then GoTo 10
    CLAX$ = CLAX$ + Chr$(PPX%)
    If Len(CLAX$) < 6 Then GoTo 10
    '
    RESPUES$ = ""
    For m% = 1 To 6
        POSI% = (5 * m%) Mod 6
        While POSI% < 1: POSI% = POSI% + 6: Wend
        VORI% = XVALO(Mid$(CLAX$, POSI%, 1))
        VTRA% = (VORI% + 3 * m%) Mod 10
        RESPUES$ = RESPUES$ + Mid$(Str$(VTRA%), 2)
    Next m%
    '
    VDEF$ = CLAX$ + Space$(6)
    CTRS$ = OBJPLA$("CLACONTRA", VDEF$)
    If CTRS$ = "" Then
        Call FUERASYS
    End If
    CONTRASE$ = Mid$(CTRS$, 7, 6)
    If CONTRASE$ <> RESPUES$ Then
        INTEN% = INTEN% + 1
        If INTEN% < 3 Then GoTo 5
        Call FUERASYS
    End If
    '
15  Dim IDS As String * 96
    n1% = FILEOPEN%("IDENSYS.DRV", 0, 96)
    If LOF(n1%) >= 192 Then
        Get #n1%, 1, IDS$
        DEXX$ = DESENCRI$(Left$(IDS$, 94))
        '
        Get #n1%, 2, IDS$
        DF$ = DESENCRI$(Left$(IDS$, 94))
        '
        Get #n1%, 3, IDS$
        PUETRA% = Asc(Mid$(IDS$, 17, 1))
        If PUETRA% < 1 Then PUETRA% = 1
        '
        NS$ = TRIM$(Left$(DEXX$, 12))
        EM$ = TRIM$(Mid$(DEXX$, 13, 30))
        SC$ = TRIM$(Mid$(DEXX$, 43, 20))
        PS$ = TRIM$(Mid$(DEXX$, 63, 30))
            FLIM = CVI(Mid$(DEXX$, 93, 2))
            FLI1% = FLIM
        FL$ = FECHATEX$(FLIM)
        DS$ = TRIM$(Mid$(DF$, 11, 40))
        FELI% = CVI(Left$(DF$, 2))
        If FELI% < 8000 Or FELI% > HOY(HO$) Then
            FELI% = HOY(HO$)
        End If
        '
      Else
        HOII = HOY(HO$)
        FLIM = HOII: While DIENTRE%(HOII, FLIM) < 100
                        FLIM = FLIM + 1
                     Wend
        FLI1% = FLIM
        NS$ = Mid$(HO$, 7, 2) + Mid$(HO$, 4, 2) + Left$(HO$, 2) + "-01"
        If XVALO(Left$(NS$, 1)) < 9 Then NS$ = "2" + NS$ Else NS$ = "1" + NS$
        '
        EM$ = ""
        SC$ = ""
        DS$ = "SISTEMA PRECONFIGURADO FLEXOFT (r)"
        FELI% = HOY(HO$)
        PUETRA% = 1
        '
    End If
    '
    MSX$ = REGLEIDO$("MASTER", 1&)
    FEX = CVI(Mid$(MSX$, 124, 2))
    Call CIERRARCH("MASTER")
    FES$ = FECHATEX$(FEX)
    CLAV9$ = Mid$(FES$, 7, 2) + Mid$(FES$, 4, 2) + Mid$(FES$, 1, 2)
    UI% = 0: For KI% = 1 To Len(CLAV9$)
                UI% = UI% + XVALO(Mid$(CLAV9$, KI%, 1))
             Next KI%
             UI% = UI% Mod 10
    CLAV9$ = CLAV9$ + TRIM$(Str$(UI%))
    CLAV9$ = CLAV9$ + Chr$(65 + UI%)
    '
    VDEF$ = String$(180, 0)
    Call REPLA(VDEF$, "FLEXOFT", 1, 8)
    Call REPLA(VDEF$, DS$, 9, 40)
    Call REPLA(VDEF$, NS$, 49, 12)
    Call REPLA(VDEF$, EM$, 61, 30)
    Call REPLA(VDEF$, SC$, 91, 20)
    Call REPLA(VDEF$, PS$, 111, 30)
    Call REPLA(VDEF$, MKI$(FLI1%), 141, 2)
    Call REPLA(VDEF$, MKI$(FELI%), 159, 2)
    Call REPLA(VDEF$, CLAV9$, 151, 8)
    Call REPLA(VDEF$, Chr$(PUETRA%), 161, 1)
    '
    OBX$ = OBJPLA$("CARIDENSYS", VDEF$)
    If OBX$ = "" Then
        Call FUERASYS
    End If
    '
    PUETRA% = Asc(Mid$(OBX$, 161, 1))
    If PUETRA% < 1 Then PUETRA% = 1
    '
    CLAV9$ = Mid$(OBX$, 151, 8)
    If Len(TRIM$(CLAV9$)) <> 8 Then
        Call FUERASYS
    End If
    UI% = 0: For KI% = 1 To 6
                UI% = UI% + XVALO(Mid$(CLAV9$, KI%, 1))
             Next KI%
             UI% = UI% Mod 10
    If XVALO(Mid$(CLAV9$, 7, 1)) <> UI% Then
        Call FUERASYS
    End If
    If Asc(Mid$(CLAV9$, 8, 1)) <> UI% + 65 Then
        Call FUERASYS
    End If
    FES$ = Mid$(CLAV9$, 5, 2) + "/" + Mid$(CLAV9$, 3, 2) + "/" + Left$(CLAV9$, 2)
    FLI1% = FECHANUM(FES$)
    '
    MSX$ = REGLEIDO$("MASTER", 1&)
    Call REPLA(MSX$, MKI$(FLI1%), 124, 2)
    Call REPLA(MSX$, MKI$(9876 - FLI1%), 126, 2)
    PPX% = 23: For KI% = 124 To 127
                   PPX% = PPX% + (KI% - 123) * Asc(Mid$(MSX$, KI%, 1))
               Next KI%
               PPX% = PPX% Mod 256
               Mid$(MSX$, 128, 1) = Chr$(PPX%)
    Call GRAREG("MASTER", MSX$, 1&)
    Call CIERRARCH("MASTER")
    '
    X$ = ENCRIPTA$(Mid$(OBX$, 49, 94))
    SUCAR% = 0: For i% = 1 To Len(X$): SUCAR% = SUCAR% + Asc(Mid$(X$, i%, 1)) - 71: Next i%
    IDENSYS$ = X$: DIVERI$ = MKI$(SUCAR%)
    '
    XX$ = Space$(94)
    FELI% = CVI(Mid$(OBX$, 159, 2))
    Call REPLA(XX$, MKI$(FELI%), 1, 2)
    Call REPLA(XX$, Left$(OBX$, 48), 3, 48)
    Call REPLA(XX$, DIAPLI$, 51, 34)
    X$ = ENCRIPTA$(XX$)
    SUCAR% = 0: For i% = 1 To Len(X$): SUCAR% = SUCAR% + Asc(Mid$(X$, i%, 1)) - 83: Next i%
    IDENSYS2$ = X$: DIVERI2$ = MKI$(SUCAR%)
    '
    Get #n1%, 3, IDS$
    IDENSYS3$ = Left$(IDS$, 94)
    Call REPLA(IDENSYS3$, Chr$(PUETRA%), 17, 1)
    SUCAR% = 0: For i% = 1 To Len(IDENSYS3$): SUCAR% = SUCAR% + Asc(Mid$(IDENSYS3$, i%, 1)) - 83: Next i%
    DIVERI3$ = MKI$(SUCAR%)
    '
    LSet IDS$ = IDENSYS$ + DIVERI$: Put #n1%, 1, IDS$
    LSet IDS$ = IDENSYS2$ + DIVERI2$: Put #n1%, 2, IDS$
    LSet IDS$ = IDENSYS3$ + DIVERI3$: Put #n1%, 3, IDS$
    Close #n1%: n1% = 0
    '
    If VERDEMO% > 0 Then Exit Sub
    '
    FLX1% = 9876 - FLI1%
    n0% = FILEOPEN%("\IEEE.DRV", 0, 128)
    Dim N0S As String * 128
    Get #n0%, 73, N0S$: Put #n0%, 79, N0S$
    '
    Get #n0%, 1, N0S$
    N1S$ = N0S$
    Call REPLA(N1S$, MKI$(FLI1%), 45, 2)
    LSet N0S$ = N1S$: Put #n0%, 1, N0S$
    Get #n0%, 2, N0S$
    N1S$ = N0S$: Call REPLA(N1S$, MKI$(FLX1%), 73, 2)
    LSet N0S$ = N1S$: Put #n0%, 2, N0S$
    Close #n0%: n0% = 0
    '
End Sub
'

Function DerechoSupervisor%()
    NUMUS% = USNBR% Mod 100
    Static DERESUP%

    If DERESUP% > 0 Then GoTo 99
    If NUMUS% = 1 Then DERESUP% = 1: GoTo 99
    
    '
    'BUSCO SOLO EN SQL POR QUE EL DERECHO SE DEBE DAR DE AQUI, ASI QUE DEBERIA ESTAR GRABADO YA JUNTO CON EL PROCESO INICIAL
    DERECHOX$ = "DERUSER"
    DERESUP% = 0
    
    CONDI$ = " NUM_USER =  " & USNBR%
    CONDI$ = CONDI$ + " AND COD_APLI = '" & DERECHOX$ & "'"
    SQLX$ = "SELECT DER_READ,DER_WRITE FROM DER_ACCE WITH(NOLOCK) WHERE " & CONDI$
    Set rs = READSET(SQLX$)
    If XVALO(rs!DER_READ) > 0 And XVALO(rs!DER_WRITE) > 0 Then
       DERESUP% = 1
       GoTo 99
    End If
    rs.Close: Set rs = Nothing
    Exit Function
    
    
    
'    On Error GoTo ERROR_MDB
'    Call ABRETACON
'    SQLX$ = "SELECT * FROM DER_ACCE "
'    SQLX$ = SQLX$ + " WHERE NUM_USER = " & USNBR%
'    SQLX$ = SQLX$ + " AND COD_APLI = '" & DERECHOX$ & "'"
'    '
'    Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenSnapshot)
'    With PARATEMP
'      On Error Resume Next
'      .MoveFirst
'      If Err Then
'          On Error GoTo 0
'        Else
'          On Error GoTo 0
'          If XVALO(!DER_READ) > 0 And XVALO(!DER_WRITE) > 0 Then
'             DERESUP% = 1
'          End If
'      End If
'    End With
'    PARATEMP.Close
'    BaDaConsul.Close
99  DerechoSupervisor% = DERESUP%
'    Exit Function
'ERROR_MDB:
''    Call COMUNI("No se Ha Podido Guardar el Derecho en Tabla SQLCTRL.MDB\Consulte a su Soporte de Sistema")
'    On Error GoTo 0
End Function

Function SERINUM$()
    '
    Call DATSYS(NS$, EM$, SC$, PS$, FL$, DS$, FLIC%)
    SERINUM$ = NS$
    '
End Function

' ***************************************************************************

        Function FELICEN%()
        '
        Call DATSYS(NS$, EM$, SC$, PS$, FL$, DS$, FLIC%)
        FELICEN% = FLIC%
        End Function

' ***************************************************************************

        Function EMPRELI$()
        '
        Call DATSYS(NS$, EM$, SC$, PS$, FL$, DS$, FLIC%)
        EMPRELI$ = EM$
        '
        If InStr(UCase$(EMPREAC$), "FRIOLATI") > 0 Then
          EMPRELI$ = "FRIOLATINA S.A."
        End If
        '
        If InStr(UCase$(EMPREAC$), "PERGAMINO") > 0 Then
          EMPRELI$ = "BIG BAG PERGAMINO"
        End If
        End Function

' ***************************************************************************

        Function SECTOLI$()
        '
        Call DATSYS(NS$, EM$, SC$, PS$, FL$, DS$, FLIC%)
        SECTOLI$ = SC$
        End Function

' ***************************************************************************

        Function RESPOSI$()
        '
        Call DATSYS(NS$, EM$, SC$, PS$, FL$, DS$, FLIC%)
        RESPOSI$ = PS$
        End Function

' ***************************************************************************

        Function FECHALI$()
        '
        Call DATSYS(NS$, EM$, SC$, PS$, FL$, DS$, FLIC%)
        FECHALI$ = FL$
        End Function

' ***************************************************************************

        Function DENOSYS$()
        '
        Call DATSYS(NS$, EM$, SC$, PS$, FL$, DS$, FLIC%)
        DENOSYS$ = DS$
        End Function
        '

' ***************************************************************************
            
        Function MULTIEM%()
        '
        Static MT%
        If MT% = 0 Then
           On Error GoTo 90
           n1% = FreeFile
           Open App.Path + "\FLEXMUL.DRV" For Input Shared As #n1%
           If LOF(n1%) > 0 Then
                MT% = 1
              Else
                MT% = (-1)
           End If
           Close #n1%: n1% = 0
           On Error GoTo 0
        End If
        '
        If MT% > 0 Then MULTIEM% = 1
        Exit Function
        '
90      Resume 91
91      MT% = (-1)
        MULTIEM% = 0
        '
        End Function

' ***************************************************************************

Sub FINAL(MODU$)
    '
    If TRIM$(MODU$) = "" Then
         Dim NOFORMU$(256)
         CAFORMU% = Forms.Count
         For KKI% = 1 To Forms.Count
            NOFORMU$(KKI%) = Forms(KKI% - 1).Name
         Next KKI%
         Call CIERRARCH("*.*")
         Call BLOQEXE("*")
         Call LIBARCH("*.*")
         Close
         '
         If TRIM$(JOBID$) <> "" Then
           WAITFILE$ = LUCOM$ + JOBID$ + ".DRV"
15         EXIWAIT% = EXISTE%(WAITFILE$)
           If EXIWAIT% > 0 Then
             On Error Resume Next
             Kill WAITFILE$
             On Error GoTo 0
             GoTo 15
           End If
         End If
         '
         ESFINAL% = 1
         On Error GoTo 0
         On Error Resume Next
10       For KKI% = CAFORMU% To 1 Step -1
            For KKJ% = 1 To Forms.Count
              If Forms(KKJ% - 1).Name = NOFORMU$(KKI%) Then
                Forms(KKJ% - 1).Hide
                On Error Resume Next
                Unload Forms(KKJ% - 1)
                DoEvents
              End If
            Next KKJ%
         Next KKI%
         On Error GoTo 0
         Call CANCELUCOM
         Close: End
       ElseIf Left$(MODU$, 1) = "*" Then
         Call CIERRARCH("*.*")
         Call CONECRUN(MODU$, "")
       ElseIf Left$(MODU$, 1) = "?" Then
         Call CIERRARCH("*.*")
         Call CONECRUN("GENREP07/" + TRIM$(Mid$(MODU$, 2)), "")
       Else
         Call MENSERR(24, "Error de Conexion\a Módulo " + MODU$)
         Call CIERRARCH("*.*")
         Close: End
    End If
    '
End Sub

Sub CONECRUN(MODU$, APLI$)
      '
'MsgBox "CONECRUN" + MODU$
      If DERACCE%(MODU$, APLI$) < 1 Then
         Exit Sub
      End If
      '
      If Left$(MODU$, 1) = "*" Or Left$(MODU$, 1) = "?" Then
           CONECTA$ = UCase$(MODU$)
         Else
           CONECTA$ = UCase$(AJUSTI$(MODU$, 8))
      End If
      '
      If Left$(CONECTA$, 1) = "?" Then
         CONECTA$ = "GENREP07/" + TRIM$(Mid$(MODU$, 2))
         MODU$ = CONECTA$
      End If
      '
199   TITU$ = AJUSTI$(APLI$, 32)
      If TITU$ = "" Then TITU$ = "MODULO SIN IDENTIFICACION"
      OPERA$ = TITU$
      '
      If Left$(UCase$(MODU$), 3) = "FAC" Then
        HOII% = HOY(HOS$)
        If EXISTE%(LUDAT$ + "VENCICAI.RFS") = 1 Then
          For UU& = 1 To ULTREG&("VENCICAI")
            XYZ$ = REGLEIDO$("VENCICAI", UU&)
            FEVENCAI% = CVI(Mid$(XYZ$, 63, 2))
            If FEVENCAI% > 0 Then
              If FEVENCAI% < HOII% Then
                  DEFORMU$ = TRIM$(Left$(XYZ$, 32))
                  Call MENSERR(24, "Imposible Acceder a Módulo '" + TRIM$(UCase$(MODU$)) + "'\CAI de " + DEFORMU$ + " está Vencido.")
                  Exit Sub
                ElseIf FEVENCAI% < HOII% + 32 Then
                  DEFORMU$ = TRIM$(Left$(XYZ$, 32))
                  Call MENSERR(24, "ATENCION: \CAI de " + DEFORMU$ + "\Próximo a Vencer !!!")
              End If
            End If
          Next UU&
        End If
      End If
      '
      If Left$(MODU$, 8) = "WEDITABU" Then
          '
          DIRA$ = TRIM$(CurDir(""))
          If InStr(DIRA$, " ") > 0 Then
               MsgBox "Esta Aplicación no Puede Ejecutarse en la Carpeta de Instalación " + DIRA$ + ". reinstale el Sistema en \FLEXOFT\APLICA."
               Exit Sub
          End If
          '
          If EXISTE%(LUBAS$ + MODU$ + ".EXE") < 1 Then
              BRXX$ = REPLACAR$(LUBAS$ + MODU$ + ".EXE", "\", "/")
              Call MENSERR(24, "No Existe " + BRXX$)
              Exit Sub
          End If
          '
          DRIVA$ = Left$(DIRA$, 2)
          DIRA$ = Mid$(DIRA$, 3)
          If Mid$(DRIVA$, 2, 1) <> ":" Then
               Call COMUNI("Error de Instalación y/o de Configuración.\Imposible Identificar  la Unidad de Disco.\ \Consulte a su Soporte de Sistemas FLEXOFT.")
               Exit Sub
          End If
          '
          NX% = FILEOPEN%(LUCOM$ + "FLEXBAT.BAT", 2, 0)
          Print #NX%, "@ECHO OFF"
          Print #NX%, DRIVA$
          Print #NX%, "CD " + DIRA$
          Print #NX%, LUBAS$ + MODU$
          Print #NX%, "EXIT"
          Close #NX%: NX% = 0
          '     '
          MDX$ = AJUSTI$("*" + Mid$(MODU$, 13, 8), 8): Call GRARAM(21, MDX$)
          APX$ = AJUSTI$(APLI$, 32): Call GRARAM(29, APX$)
          Call BAJRAM
          '
          CCI$(0) = "WEDITABU\" + Mid$(MDX$, 2)
          Call GRATRACO(RAZO$, EMPRE$, SE$, LU$, QWERTY%, CCI$(), EJU%, menu%)
          '
          LCB$ = LUBAS$
          If EXISTE%(LUCOM$ + "FLEXPIF.PIF") = 1 Then
              Shell LUCOM$ + "FLEXPIF.PIF", vbNormalFocus
            Else
              TXT$ = REPLACAR$(LUCOM$ + "FLEXPIF.PIF", "\", "/")
              Call MENSERR(24, "No Existe " + TXT$)
          End If
          '
          Exit Sub
          '
      End If
      '
210   If Left$(CONECTA$, 1) = "#" Then
          '
          ATRIBU$ = ""
          PPX% = InStr(MODU$, "/")
          If PPX% > 0 Then
              CONECTA$ = Left$(MODU$, PPX% - 1)
              ATRIBU$ = Mid$(MODU$, PPX% + 1)
          End If
          '
          CONECTA$ = TRIM$(Mid$(CONECTA$, 2))
          AX$ = CONECTA$ + ".EXE"
          If Left$(CONECTA$, 4) = "FLEX" Then AX$ = LUBAS$ + AX$
          If Left$(CONECTA$, 4) = "PARA" Then AX$ = LUBAS$ + AX$
          '
          If EXISTE%(AX$) < 1 Then
             BRXX$ = REPLACAR$(AX$, "\", "/")
             Call OPNOIN(BRXX$)
             Exit Sub
          End If
          '
          If COMALTER%("No se Encontró la Versión para Windows\de la Aplicación Elegida.\  \ Puede ejecutar la Versión D.O.S. o Abandonar\\Ejecutar D.O.S.\Abandonar") <> 1 Then
             Exit Sub
          End If
          '
          DIRA$ = CurDir("")
          If InStr(DIRA$, " ") > 0 Then
             MsgBox "Esta Aplicación no Puede Ejecutarse en la Carpeta de Instalación " + DIRA$ + ". reinstale el Sistema en \FLEXOFT\WINDOWS."
             Exit Sub
          End If
          '
          DRIVA$ = Left$(DIRA$, 2)
          DIRA$ = Mid$(DIRA$, 3)
          If Mid$(DRIVA$, 2, 1) <> ":" Then
             Call COMUNI("Error de Instalación y/o de Configuración.\Imposible Identificar  la Unidad de Disco.\ \Consulte a su Soporte de Sistemas FLEXOFT.")
             Exit Sub
          End If
          '
          INVOCA$ = AX$
          If ATRIBU$ <> "" Then INVOCA$ = INVOCA$ + " " + ATRIBU$
          NX% = FILEOPEN%(LUCOM$ + "FLEXBAT.BAT", 2, 0)
          Print #NX%, "@ECHO OFF"
          Print #NX%, DRIVA$
          Print #NX%, "CD " + DIRA$
          Print #NX%, INVOCA$
          Print #NX%, "EXIT"
          Close #NX%: NX% = 0
          '
          MDX$ = AJUSTI$(MODU$, 8): Call GRARAM(21, MDX$)
          APX$ = AJUSTI$(APLI$, 32): Call GRARAM(29, APX$)

          Call BAJRAM
          '
          If CONECTA$ <> "FILTABU" Then
              CCI$(0) = CONECTA$
          End If
          Call GRATRACO(RAZO$, EMPRE$, SE$, LU$, QWERTY%, CCI$(), EJU%, menu%)
          Call CIERRARCH("*.*")
          Close
          '
          LCB$ = LUBAS$
          If EXISTE%(LUCOM$ + "FLEXPIF.PIF") = 1 Then
                Shell LUCOM$ + "FLEXPIF.PIF", vbNormalFocus
              Else
                TXT$ = REPLACAR$(LUCOM$ + "FLEXPIF.PIF", "\", "/")
                Call MENSERR(24, "No Existe " + TXT$)
          End If
          Exit Sub
          '
      ElseIf Left$(CONECTA$, 1) = "*" Then
           '
          If EXISTE%(LUBAS$ + "WINTABU.EXE") < 1 Then
            If (EXISTE%(LUBAS$ + "WINTAB04.EXE") < 1 And EXISTE%(LUBAS$ + "WINTAB07.EXE") < 1) Then
              BRXX$ = REPLACAR$(LUBAS$ + "WINTABU.EXE", "\", "/")
              Call MENSERR(24, "No Existe " + BRXX$)
              Exit Sub
            End If
          End If
          '
          CONECTB$ = CONECTA$ + "/"
          PPXX% = InStr(CONECTB$, "/")
          If PPXX% > 0 Then
            CONECTB$ = Left$(CONECTB$, PPXX% - 1)
            ARZZ$ = UCase$(Mid$(CONECTB$, 2)) + ".PRF"
            Call GRATABCONTRO(ARZZ$)
          End If
          '
          MDX$ = AJUSTI$(MODU$, 8): Call GRARAM(21, MDX$)
          APX$ = APLI$
          PPX% = InStr(APX$, "/WAIT")
          ESPERA% = 0
          If PPX% > 0 Then
             APX$ = Left$(APX$, PPX% - 1)
             ESPERA% = 1
          End If
          APY$ = AJUSTI$(APX$, 32)
          Call GRARAM(29, APY$)
          Call GRARAM(99, AJUSTI$(LUDAT$, 46))
          Call GRARAM(145, AJUSTI$(LUBAS$, 46))
          Call GRARAM(193, AJUSTI$(EMPREAC$, 30))
          Call GRAJOBID
          Call BAJRAM
          '
          YACARWI% = 0
          TINI = 3600 * XVALO(Left$(Time$, 2)) + 60 * XVALO(Mid$(Time$, 4, 2)) + XVALO(Mid$(Time$, 7, 2))
          HOII% = HOY(HOS$)
          '
          If UCase$(App.EXEName) = "WINTABU" Then
             Call BLOQEXE("*")
          End If
          '
' elije generador residente o normal
          If NUEWINTA% = 1 Then
            If XVALO(Right$(App.EXEName, 2)) < 4 Then
              On Error Resume Next
              CONECWIN.Label1.Caption = CONECTA$
              If Err Then GoTo 230
              On Error GoTo 0
              CONECWIN.Show 1
              DoEvents
              On Error Resume Next
              For KKU% = Forms.Count To 1 Step -1
                If Forms(KKU%).Name <> "CONECWIN" Then
                  Forms(KKU% - 1).Enabled = True
                  If Forms(KKU% - 1).Visible = True Then
                    Forms(KKU%).SetFocus
                    GoTo 19
                  End If
                End If
              Next KKU%
19            On Error GoTo 0
              Exit Sub
            End If
          End If
          '
230       On Error GoTo 0
          ACONE$ = "WINTABU"
          
          If XVALO(Right$(App.EXEName, 2)) >= 4 Then ACONE$ = "WINTAB04"
          If XVALO(Right$(App.EXEName, 2)) >= 7 Then
             If UCase$(Left$(App.EXEName, 8)) <> "LIDIRMAI" Then
                ACONE$ = "WINTAB07"
             End If
          End If
          If App.EXEName = "AMASAHP" Then ACONE$ = "WINTAB07"
          '
          On Error Resume Next
          Shell LUBAS$ + ACONE$ + " " + CONECTA$ + "#" + JOBID$ + "#", 1
          If Err Then
            MsgBox "Error " & Err & " - Imposible Conectar con " & LUBAS$ + ACONE$ + " " + CONECTA$ + "#" + JOBID$ + "#"
            On Error GoTo 0
            Exit Sub
          End If
          On Error GoTo 0
          '
          If UCase$(LUCOM$) = UCase$(CurDir) Or UCase$(LUCOM$) = UCase$(CurDir) + "\" Then Exit Sub
          '
          CONECWIN.Timer2.Enabled = True
          CONECWIN.Top = Screen.Height
          CONECWIN.Label2 = ""
          CONECWIN.Label3 = ""
          CONECWIN.Show 1
          CONECWIN.Timer2.Enabled = False
          Exit Sub
          '
          Screen.MousePointer = 11
          YACARWI% = 0
          Dim WTT As String * 128
240       'TIM1 = Timer
          'While Timer < TIM1 + 0.5: DoEvents: Wend
          On Error GoTo 250
          NK% = FreeFile
          Open LUCOM$ + ACONE$ + ".EXK" For Random Access Read Write As #NK% Len = 128
          Get #NK%, 1, WTT$
          Put #NK%, 1, WTT$
          On Error GoTo 0
          Close #NK%: NK% = 0
          If Mid$(WTT$, 69, 8) <> HOS$ Then GoTo 240
          TIMF$ = Mid$(WTT$, 79, 8)
          tinj = 3600 * XVALO(Left$(TIMF$, 2)) + 60 * XVALO(Mid$(TIMF$, 4, 2)) + XVALO(Mid$(TIMF$, 7, 2))
          If tinj < TINI Then GoTo 240
          '
          Call FRESHALL
          Screen.MousePointer = 1
          Exit Sub
          '
250       Resume 251
251       YACARWI% = 1
          'If NK% > 0 Then
          '  Close #NK%
          '  NK% = 0
          'End If
          GoTo 240
          '
      Else
          MDY$ = MODU$
          ATRIBU$ = ""
          PPX% = InStr(MDY$, "/")
          If PPX% > 0 Then
              MDY$ = Left$(MODU$, PPX% - 1)
              ATRIBU$ = Mid$(MODU$, PPX% + 1)
          End If
          '
          RECARGA% = 0
          MDY$ = TRIM$(MDY$)
          MDX$ = AJUSTI$(MDY$, 8): Call GRARAM(21, MDX$)
          APX$ = APLI$
          PPX% = InStr(APX$, "/WAIT")
          ESPERA% = 0
          If PPX% > 0 Then
             APX$ = Left$(APX$, PPX% - 1)
             ESPERA% = 1
          End If
          APY$ = AJUSTI$(APX$, 32)
          Call GRARAM(29, APY$)
          Call GRAJOBID
          Call BAJRAM
          If UCase$(TRIM$(MDX$)) = "WINTABU" Then
              MDX$ = LUBAS$ + "WINTABU"
          End If
          '
          If ESPERA% = 1 Then
             n1% = FILEOPEN%(LUCOM$ + JOBID$ + ".DRV", 2, 0)
             Print #n1%, "ESPERAR"
             Close #n1%: n1% = 0
          End If
          '
          If MDX$ <> App.EXEName Then ' para posibilitar conectarse consigo mismo
            If REACTIVA%(MDX$) = 1 Then Exit Sub
          End If
          '
          If TRIM$(UCase$(MDX$)) = App.EXEName Then
            Call CIERRARCH("*.*")
            Call BAJALDISCO
            Close
          End If
          '
          On Error GoTo OPNOIN1
          Shell TRIM$(MDX$) + ".EXE " + ATRIBU$ + "#" + JOBID$ + "#", 1
          On Error GoTo 0
          If ESPERA% = 1 Then
             While EXISTE%(LUCOM$ + JOBID$ + ".DRV") > 0
             Wend
          End If
          Exit Sub
          '
      End If
      '
OPNOIN:
      Call COMUNI(MENSA$)
      Exit Sub
      '
OPNOIN1:
      Resume OPNOIN2
OPNOIN2:
      On Error GoTo 0
      TXT$ = "La Opción Elegida  no se Encuentra Disponible\en la Versión FLEXOFT  Actualmente Instalada."
      TXT$ = TXT$ + "\  \Aplicación: " + UCase$(MDX$) + ".EXE" + "\  \"
      TXT$ = TXT$ + "Consulte con su Soporte de Sistemas o su Pro-\veedor de Software para la Adquisición e Insta-\lación de la misma."
      Call COMUNI(TXT$)
      Exit Sub
      '
End Sub
'

Function DERACCE%(MDX$, APL$)
    '
    ' MDX$ = MODULO INVOCADO
    ' APX$ = DESCRIPCION DEL MÓDULO O APLICACION
    '
    Static CTRX%, CAPLI$(), DAPLI$()
    '
    'CONSIDERA EL DERECHO TRADICIONAL POR SI EXISITIA
    MDX1$ = TRIM$(UCase$(MDX$))
    If Left$(MDX1$, 2) = "R-" Then MDX$ = Mid$(MDX1$, 3)
    
    DRXXX% = 0
    MDY$ = REPLACAR$(TRIM$(UCase$(MDX$)), "'", "")
      If MDY$ = "" Then GoSub CARGALISDER: DRXXX% = 2: GoTo 99
    APM$ = REPLACAR$(TRIM$(UCase$(APL$)), "'", "")
    If TRIM$(UCase$(MDX$)) = "FSETUP04/CLAVUSER" Then
      GoSub CARGALISDER: DRXXX% = 2: GoTo 99  ' PARA PODER CAMBIAR CLAVE PERSONAL
    End If
    '
    ATRICON$ = ""
    ' ESTO ES PARA RESOLVER EL ERROR DE ACHERNAR SEGÚN OT-11-0046
    If Left$(UCase$(MDY$), 16) = "TRAZAB07/BORECEP" Then
        MDY$ = "TRAZAB07/BORECEP"
    End If
    '
    If Left$(UCase$(MDY$), 6) = "GENREP" Then
       If Mid$(UCase$(MDY$), 9, 1) = "/" Then
          MDY$ = Mid$(MDY$, 10)
          If TRIM$(APM$) = "" Then
             APM$ = VALOBADA("SQL_QUERIES", "DES_QUERY", "COD_QUERY = '" & MDY$ & "'", "NOMESS")
             If TRIM$(APM$) = "" Then APM$ = VALOTACO("SQL_QUERIES", "DES_QUERY", "COD_QUERY = '" & MDY$ & "'", "NOMESS")
             If APM$ = "" Then APM$ = "Reporte " + MDY$
          End If
       End If
    End If
    '
    PPXX% = InStr(MDY$, "/")
    If PPXX% > 1 Then
       ATRICON$ = Mid$(MDY$, PPXX% + 1)
       If ATRICON$ = "NOMESS" Then MDY$ = Left$(MDY$, PPXX% - 1)
    End If
    '
    PPXX% = InStr(MDY$, "(")
    If PPXX% > 1 Then
       MDY$ = Left$(MDY$, PPXX% - 1)
    End If
    '
    PPY$ = Left$(MDY$, 1)
    If Asc(PPY$) < 48 Or (Asc(PPY$) > 57 And Asc(PPY$) < 65) Or Asc(PPY$) > 90 Then
        MDY$ = Mid$(MDY$, 2)
    End If
    '
    MDY$ = TRIM$(MDY$): GoSub CARGALISDER
    If MDY$ = "" Then DRXXX% = 2: GoTo 99
    '
    QWR% = USNBR% Mod 100
    If QWR% = 1 Then DRXXX% = 2: GoTo 99  ' supervisor entra directo
    If DerechoSupervisor% = 1 Then DRXXX% = 2: GoTo 99
    '
    '
    MDY32$ = Left$(MDY$, 32)
    'ESTO AGREGADO NUEVO PARA QUE PRIMERO MIRE EN TABLA SQL SI YA TIENE EL DERECHO NO SIGUE
    '__________________________________________________  agregado esto entre las lineas 13_06_24
    If CantRegistros("DER_ACCE", "NUM_USER > 0", "NOMESS") > 0 Then
        CONDI$ = " NUM_USER =  " & QWR%
        CONDI$ = CONDI$ + " AND COD_APLI = '" & MDY32$ & "'"
        SQLX$ = "SELECT DER_READ,DER_WRITE FROM DER_ACCE WITH(NOLOCK) WHERE " & CONDI$
        Set RSX = READSET(SQLX$)
        If Not RSX.EOF Then
            If XVALO(RSX!DER_READ) > 0 Then DRXXX% = 1
            If XVALO(RSX!DER_WRITE) > 0 Then DRXXX% = 2
        End If
        RSX.Close: Set RSX = Nothing
    End If
    If DRXXX% > 1 Then
      DERACCE% = DRXXX%
      Exit Function
    End If
    '________________________________________________________
    ' BUSCAR EL DERECHO EN MDB
    
    DRXXX% = 0
    Call ABRETACON
    SQLX$ = "SELECT * FROM DER_ACCE "
    SQLX$ = SQLX$ + " WHERE NUM_USER = " & QWR%
    SQLX$ = SQLX$ + " AND COD_APLI = '" & MDY32$ & "'"
    '
    Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenSnapshot)
    With PARATEMP
      On Error Resume Next
      .MoveFirst
      If Err Then
          On Error GoTo 0
        Else
          On Error GoTo 0
          If XVALO(!DER_READ) > 0 Then DRXXX% = 1
          If XVALO(!DER_WRITE) > 0 Then DRXXX% = 2
      End If
    End With
25  PARATEMP.Close
    BaDaConsul.Close
    '
    If DRXXX% < 2 Then
       ' BUSCAR DERECHO EN FLEXACC
       MDY8$ = Left$(MDY$, 8)
       PPXX% = InStr(MDY8$, "/")
       If PPXX% > 1 Then
          MDY8$ = Left$(MDY8$, PPXX% - 1)
       End If
       RFF$ = RECFILT$("FLEXACC", 2, MDY8$)
       For KU& = 1 To Len(RFF$) Step 4
         TYU& = CVS(Mid$(RFF$, KU&, 4))
         XTY$ = REGLEIDO$("FLEXACC", TYU&)
         If CVI(Left$(XTY$, 2)) = QWR% Then
           If TRIM$(UCase$(Mid$(XTY$, 3, 8))) = MDY8$ Then
             If Mid$(XTY$, 11, 1) = "S" Then DRXXX% = 1
             If Mid$(XTY$, 12, 1) = "S" Then DRXXX% = 2
             Call CIERRARCH("FLEXACC")
             GoTo 99
           End If
         End If
       Next KU&
       Call CIERRARCH("FLEXACC")
       '
       If APL$ = "" Then APL$ = APM$
       APM$ = TRIM$(UCase$(APL$))
       If APM$ = "" Then
          If Left$(MDX$, 1) = "*" Then
             MDZ$ = TRIM$(Mid$(MDX$, 2))
             If EXISTE%(MDZ$ + ".PRF") > 0 Then
                APM$ = TRIM$(Left$(LOADFILE$(MDZ$ + ".PRF"), 48))
             End If
          End If
       End If
       '
20     If Left$(ATRICON$, 6) <> "NOMESS" Then  ' NO PRESENTA MENSAJE
          MENSA$ = "Usuario no Habilitado\"
          MENSA$ = MENSA$ + "para Ejecutar esta Opción.\  \"
          MENSA$ = MENSA$ + "Modulo: " + TRIM$(UCase$(MDY$)) + "\"
          MENSA$ = MENSA$ + "Aplicacion: " + TRIM$(APM$) + "\   \"
          MENSA$ = MENSA$ + "Solicite la Autorizacion Correspondiente\"
          MENSA$ = MENSA$ + "al Supervisor del Sistema FLEXOFT."
          Call COMUNI(MENSA$)
       End If
       '
    End If
    '
99  DERACCE% = DRXXX%
    '
Exit Function
'
CARGALISDER:
    '
    Call ABRECONEXION
    Call CREATABLAS_SQLCTRL
    
    If APM$ <> "" And TRIM$(MDY$) <> "" Then
       ' AGREGA A LA LISTA DE DERECHOS DEFINIDOS EN SQL
       MDY32$ = Left$(MDY$, 32)
       CONDI$ = "COD_APLI = '" & MDY32$ & "' "
       CONDI$ = CONDI$ + " AND DES_APLI = '" & Left$(APM$, 128) & "'"
       'If CantRegistros&("INDI_APLI", CONDI$) > 0 Then GoTo 100
       If CantRegistros&("INDI_APLI", CONDI$) > 0 Then GoTo 120 'linea x linea 13_06_24 (si esta en sql se va no sigue)
       SQLX$ = "SELECT * FROM INDI_APLI WHERE COD_APLI = '" & MDY32$ & "'"
       Dim rs As ADODB.Recordset
       Set rs = New ADODB.Recordset
       rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       With rs
         If .EOF Then
             .AddNew
         End If
         !COD_APLI = MDY32$
         !DES_APLI = Left$(APM$, 128)
         .Update
       End With
       rs.Close
       Set rs = Nothing
       
'       ' AGREGA A LA LISTA DE DERECHOS DEFINIDOS EN MDB
'100    On Error GoTo ERROR_MDB
'       Call ABRETACON
'       SQLX$ = "SELECT * FROM INDI_APLI "
'       SQLX$ = SQLX$ + " WHERE COD_APLI = '" & MDY32$ & "' "
'       SQLX$ = SQLX$ + " AND DES_APLI = '" & Left$(APM$, 128) & "'"
'       Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenSnapshot)
'       With PARATEMP
'          If Not (.BOF And .EOF) Then
'            PARATEMP.Close
'            GoTo 120
'          End If
'       End With
'       PARATEMP.Close
'       BaDaConsul.Close
'       '
'       SQLX$ = "SELECT * FROM INDI_APLI WHERE COD_APLI = '" & MDY32$ & "'"
'       Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenDynaset)
'       With PARATEMP
'         On Error Resume Next
'         .MoveFirst
'         If Err Then
'             On Error GoTo 0
'             .AddNew
'           Else
'             On Error GoTo 0
'             .Edit
'         End If
'         !COD_APLI = MDY32$
'         !DES_APLI = Left$(APM$, 128)
'         .Update
'       End With
'       PARATEMP.Close
'       '
'       ' AGREGA A INDIAPLI.DAT
'       X$ = REGBLAN$("INDIAPLI")
'       Call REPLA(X$, MDY$, 1, 12)
'       Call REPLA(X$, APM$, 13, 36)
'       Call FILTERPUTRECORD("INDIAPLI", 1, MDY$, X$)
'       Call CIERRARCH("INDIAPLI")
    End If
'    '
120 Return
'    '
'ERROR_MDB:
''    Call COMUNI("No se Ha Podido Guardar el Derecho en Tabla SQLCTRL.MDB\Consulte a su Soporte de Sistema")
'    On Error GoTo 0
End Function
Function DERACCE_ORIGINAL%(MDX$, APL$)
    '
    '*****IF SE VUELVE A USAR ESTA FUNCION MAS ALLA DE CAMBIAR EL NOMBRE SE DEBERA CAMBIAR LA FILA *****
    '*****: 99  DERACCE_ORIGINAL% = DRXXX% POR 99  DERACCE% = DRXXX%                               *****
    
    ' MDX$ = MODULO INVOCADO
    ' APX$ = DESCRIPCION DEL MÓDULO O APLICACION
    
    '
    Static CTRX%, CAPLI$(), DAPLI$()
    '
    'CONSIDERA EL DERECHO TRADICIONAL POR SI EXISITIA
    MDX1$ = TRIM$(UCase$(MDX$))
    If Left$(MDX1$, 2) = "R-" Then MDX$ = Mid$(MDX1$, 3)
    
    DRXXX% = 0
    MDY$ = REPLACAR$(TRIM$(UCase$(MDX$)), "'", "")
      If MDY$ = "" Then GoSub CARGALISDER: DRXXX% = 2: GoTo 99
    APM$ = REPLACAR$(TRIM$(UCase$(APL$)), "'", "")
    If TRIM$(UCase$(MDX$)) = "FSETUP04/CLAVUSER" Then
      GoSub CARGALISDER: DRXXX% = 2: GoTo 99  ' PARA PODER CAMBIAR CLAVE PERSONAL
    End If
    '
    ATRICON$ = ""
    ' ESTO ES PARA RESOLVER EL ERROR DE ACHERNAR SEGÚN OT-11-0046
    If Left$(UCase$(MDY$), 16) = "TRAZAB07/BORECEP" Then
        MDY$ = "TRAZAB07/BORECEP"
    End If
    '
    If Left$(UCase$(MDY$), 6) = "GENREP" Then
       If Mid$(UCase$(MDY$), 9, 1) = "/" Then
          MDY$ = Mid$(MDY$, 10)
          If TRIM$(APM$) = "" Then
             APM$ = VALOBADA("SQL_QUERIES", "DES_QUERY", "COD_QUERY = '" & MDY$ & "'", "NOMESS")
             If TRIM$(APM$) = "" Then APM$ = VALOTACO("SQL_QUERIES", "DES_QUERY", "COD_QUERY = '" & MDY$ & "'", "NOMESS")
             If APM$ = "" Then APM$ = "Reporte " + MDY$
          End If
       End If
    End If
    '
    PPXX% = InStr(MDY$, "/")
    If PPXX% > 1 Then
       ATRICON$ = Mid$(MDY$, PPXX% + 1)
       If ATRICON$ = "NOMESS" Then MDY$ = Left$(MDY$, PPXX% - 1)
    End If
    '
    PPXX% = InStr(MDY$, "(")
    If PPXX% > 1 Then
       MDY$ = Left$(MDY$, PPXX% - 1)
    End If
    '
    PPY$ = Left$(MDY$, 1)
    If Asc(PPY$) < 48 Or (Asc(PPY$) > 57 And Asc(PPY$) < 65) Or Asc(PPY$) > 90 Then
        MDY$ = Mid$(MDY$, 2)
    End If
    '
    MDY$ = TRIM$(MDY$): GoSub CARGALISDER
    If MDY$ = "" Then DRXXX% = 2: GoTo 99
    '
    QWR% = USNBR% Mod 100
    If QWR% = 1 Then DRXXX% = 2: GoTo 99  ' supervisor entra directo
    If DerechoSupervisor% = 1 Then DRXXX% = 2: GoTo 99
    '
    ' BUSCAR EL DERECHO EN MDB
    MDY32$ = Left$(MDY$, 32)
    DRXXX% = 0
    Call ABRETACON
    SQLX$ = "SELECT * FROM DER_ACCE "
    SQLX$ = SQLX$ + " WHERE NUM_USER = " & QWR%
    SQLX$ = SQLX$ + " AND COD_APLI = '" & MDY32$ & "'"
    '
    Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenSnapshot)
    With PARATEMP
      On Error Resume Next
      .MoveFirst
      If Err Then
          On Error GoTo 0
        Else
          On Error GoTo 0
          If XVALO(!DER_READ) > 0 Then DRXXX% = 1
          If XVALO(!DER_WRITE) > 0 Then DRXXX% = 2
      End If
    End With
25  PARATEMP.Close
     
    '
    If DRXXX% < 2 Then
       ' BUSCAR DERECHO EN FLEXACC
       MDY8$ = Left$(MDY$, 8)
       PPXX% = InStr(MDY8$, "/")
       If PPXX% > 1 Then
          MDY8$ = Left$(MDY8$, PPXX% - 1)
       End If
       RFF$ = RECFILT$("FLEXACC", 2, MDY8$)
       For KU& = 1 To Len(RFF$) Step 4
         TYU& = CVS(Mid$(RFF$, KU&, 4))
         XTY$ = REGLEIDO$("FLEXACC", TYU&)
         If CVI(Left$(XTY$, 2)) = QWR% Then
           If TRIM$(UCase$(Mid$(XTY$, 3, 8))) = MDY8$ Then
             If Mid$(XTY$, 11, 1) = "S" Then DRXXX% = 1
             If Mid$(XTY$, 12, 1) = "S" Then DRXXX% = 2
             Call CIERRARCH("FLEXACC")
             GoTo 99
           End If
         End If
       Next KU&
       Call CIERRARCH("FLEXACC")
       '
       If APL$ = "" Then APL$ = APM$
       APM$ = TRIM$(UCase$(APL$))
       If APM$ = "" Then
          If Left$(MDX$, 1) = "*" Then
             MDZ$ = TRIM$(Mid$(MDX$, 2))
             If EXISTE%(MDZ$ + ".PRF") > 0 Then
                APM$ = TRIM$(Left$(LOADFILE$(MDZ$ + ".PRF"), 48))
             End If
          End If
       End If
       '
20     If Left$(ATRICON$, 6) <> "NOMESS" Then  ' NO PRESENTA MENSAJE
          MENSA$ = "Usuario no Habilitado\"
          MENSA$ = MENSA$ + "para Ejecutar esta Opción.\  \"
          MENSA$ = MENSA$ + "Modulo: " + TRIM$(UCase$(MDY$)) + "\"
          MENSA$ = MENSA$ + "Aplicacion: " + TRIM$(APM$) + "\   \"
          MENSA$ = MENSA$ + "Solicite la Autorizacion Correspondiente\"
          MENSA$ = MENSA$ + "al Supervisor del Sistema FLEXOFT."
          Call COMUNI(MENSA$)
       End If
       '
    End If
    '
99  DERACCE_ORIGINAL% = DRXXX%
    '
Exit Function
'
CARGALISDER:
    '
    If APM$ <> "" And TRIM$(MDY$) <> "" Then
       ' AGREGA A LA LISTA DE DERECHOS DEFINIDOS
       Call ABRETACON
       MDY32$ = Left$(MDY$, 32)
       SQLX$ = "SELECT * FROM INDI_APLI "
       SQLX$ = SQLX$ + " WHERE COD_APLI = '" & MDY32$ & "' "
       SQLX$ = SQLX$ + " AND DES_APLI = '" & Left$(APM$, 128) & "'"
       Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenSnapshot)
       With PARATEMP
          If Not (.BOF And .EOF) Then
            PARATEMP.Close
            GoTo 120
          End If
       End With
       PARATEMP.Close
       '
       SQLX$ = "SELECT * FROM INDI_APLI WHERE COD_APLI = '" & MDY32$ & "'"
       Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenDynaset)
       With PARATEMP
         On Error Resume Next
         .MoveFirst
         If Err Then
             On Error GoTo 0
             .AddNew
           Else
             On Error GoTo 0
             .Edit
         End If
         !COD_APLI = MDY32$
         !DES_APLI = Left$(APM$, 128)
         .Update
       End With
       PARATEMP.Close
       '
       ' AGREGA A INDIAPLI.DAT
       X$ = REGBLAN$("INDIAPLI")
       Call REPLA(X$, MDY$, 1, 12)
       Call REPLA(X$, APM$, 13, 36)
       Call FILTERPUTRECORD("INDIAPLI", 1, MDY$, X$)
       Call CIERRARCH("INDIAPLI")
    End If
    '
120 Return
    '
    '
End Function
'
Function DERACTU%()
    Static DRXXX%
    If DRXXX% = 0 Then
        '
        QWR% = USNBR% Mod 100
        If QWR% = 1 Then
            DRXXX% = 1
            GoTo 99                 ' supervisor entra directo
        End If
        '
        MDY$ = TRIM$(UCase$(App.EXEName))
        If MDY$ = "" Then
            DRXXX% = 1
            GoTo 99
        End If
        '
        If MDY$ = "WINTABU" Then
            DRXXX% = 1
            GoTo 99
        End If
        '
        LU$ = LUDAT$
        Dim XYZT As String * 48
        NUY% = FILEOPEN%(LU$ + "INDIAPLI.DAT", 0, 48)
        For U& = 1 To LOF(NUY%) / 48
            Get #NUY%, U&, XYZT$
            If TRIM$(Left$(XYZT$, 12)) = MDY$ Then
                Close #NUY%: NUY% = 0
                GoTo 20
            End If
        Next U&
        Close #NUY%: NUY% = 0
        DRXXX% = 1
        GoTo 99
        '
20      NUY% = FILEOPEN%(LU$ + "FLEXACC.DAT", 0, 48)
        For U& = 1 To LOF(NUY%) / 48
            Get #NUY%, U&, XYZT$
            If CVI(Left$(XYZT$, 2)) = QWR% Then
                If TRIM$(UCase$(Mid$(XYZT$, 3, 8))) = TRIM$(UCase$(MDY$)) Then
                    If Mid$(XYZT$, 12, 1) = "S" Then
                         DRXXX% = 1
                       Else
                         DRXXX% = (-1)
                    End If
                    Close #NUY%: NUY% = 0
                    GoTo 99
                End If
            End If
        Next U&
        Close #NUY%: NUY% = 0
        DRXXX% = (-1)
        '
    End If
    '
99  DERACTU% = DRXXX%
    '
End Function
'

Sub ACTUSER()
    '
    Call BLANARCH("!FLEXUSR")
    '
    n1% = FILEOPEN%("USER.DAT", 9, 64)
    Dim F1 As String * 64
    '
    For i% = 1 To LOF(n1%) / 64
       Get #n1%, i%, F1$
       If Asc(Left$(F1$, 1)) = 0 Then GoTo 305
       NOM$ = Left$(F1$, 30)
       If Left$(NOM$, 1) = "@" Then
           NOM$ = DESENCRI$(Mid$(F1$, 2, 29))
       End If
       '
       IDE$ = Mid$(F1$, 31, 12)
       If Left$(IDE$, 1) = "@" Then
           IDE$ = DESENCRI$(Mid$(F1$, 32, 11))
       End If
       CLA$ = Mid$(F1$, 43, 12)
       If Left$(CLA$, 1) = "@" Then
           CLA$ = DESENCRI$(Mid$(F1$, 44, 11))
       End If
       FCO% = CVI(Mid$(F1$, 55, 2))
       HCO% = CVI(Mid$(F1$, 57, 2)): If HCO% < 0 Or HCO% > 2359 Then HCO% = 0
       Lim% = CVI(Mid$(F1$, 59, 2))
       '
       X$ = REGBLAN$("FLEXUSR")
       Call REPLA(X$, MKI$(i%), 1, 2)
       Call REPLA(X$, UCase$(NOM$), 3, 30)
       Call REPLA(X$, UCase$(IDE$), 33, 12)
       Call REPLA(X$, CLA$, 45, 12)
       Call REPLA(X$, MKI$(FCO%), 57, 2)
       Call REPLA(X$, MKI$(HCO%), 59, 2)
       Call REPLA(X$, MKI$(Lim%), 61, 2)
       Call REGAPP("!FLEXUSR", X$)
    Next i%
    '
305 Close #n1%: n1% = 0
    Call CIERRARCH("!FLEXUSR")
    '
320 XUX% = VENTABU%("FLEXUSR")
    If XUX% < 0 Then
        GoTo 340
    End If
    '
330 n1% = FILEOPEN%("USER.DAT", 0, 64)
    For U& = 1 To ULTREG&("!FLEXUSR")
        X$ = REGLEIDO$("!FLEXUSR", U&)
        '
        Get #n1%, U&, F1$
        Y$ = F1$
        NOM$ = "@" + ENCRIPTA$(UCase$(Mid$(X$, 3, 29)))
        Call REPLA(Y$, NOM$, 1, 30)
        IDE$ = "@" + ENCRIPTA$(UCase$(Mid$(X$, 33, 11)))
        Call REPLA(Y$, IDE$, 31, 12)
        Call REPLA(Y$, Mid$(X$, 61, 2), 59, 2)
        LSet F1$ = Y$
        Put #n1%, U&, F1$
        '
    Next U&
    '
    For J& = U& + 1 To LOF(n1%) / 64
       LSet F1$ = String$(64, 0)
       Put #n1%, U&, F1$
    Next J&
    Close #n1%: n1% = 0
    '
340 For U& = 1 To ULTREG&("!FLEXUSR")
        Call GRAREG("!FLEXUSR", String$(64, 255), U&)
    Next U&
    Call CIERRARCH("!FLEXUSR")
    '
End Sub
'

Sub ACDERACC()
    '
    Screen.MousePointer = 11
    Call BLOQARCH("FLEXUSR")
    FUSER$ = LOADFILE$("USER.DAT")
    JJI& = 0: KKI% = 0
    For i% = 1 To Len(FUSER$) Step 64
       F1$ = Mid$(FUSER$, i%, 64)
       If Asc(F1$) = 0 Then GoTo 9
       NOM$ = Left$(F1$, 30)
       If Left$(NOM$, 1) = "@" Then NOM$ = DESENCRI$(Mid$(F1$, 2, 29))
       '
       KKI% = KKI% + 1
       X$ = REGBLAN$("FLEXUSR")
       Call REPLA$(X$, MKI$(KKI%), 1, 2)
       Call REPLA$(X$, NOM$, 3, 30)
       JJI& = JJI& + 1
       Call GRAREG("FLEXUSR", X$, JJI&)
    Next i%
    '
9   Call SETULTREG("FLEXUSR", JJI&)
10  Call CIERRARCH("FLEXUSR")
    '
    Call SELECHO("FLEXUSR")
    Screen.MousePointer = 1
    If VALACT1$("FLEXUSR") = "" Then Exit Sub
    '
    USEL% = XVALO(TRIM$(VALACT1$("FLEXUSR")))
    DEUSER$ = VALACT2$("FLEXUSR")
    Call ACLISDER(USEL%, DEUSER$)
    FORDERAC.Show 1
    GoTo 10
    '
End Sub
'
Sub AGREGALIST2(CAPLI$, DAPLI$)
    '
    
    If Left$(CAPLI$, 1) = "*" Then CAPLI$ = Mid$(CAPLI$, 2)
    If Left$(CAPLI$, 1) = "?" Then CAPLI$ = Mid$(CAPLI$, 2)
    If UCase$(Left$(CAPLI$, 16)) = "TRAZAB07/BORECEP" Then CAPLI$ = "TRAZAB07/BORECEP"
    If UCase$(Left$(CAPLI$, 15)) = "VENGES07/PEDIDO" Then CAPLI$ = "VENGES07/PEDIDO"
    '
    Z$ = Space$(128)
    Call REPLA(Z$, DAPLI$, 1, 36)
    Call REPLA(Z$, "      ", 45, 6)
    Call REPLA(Z$, AJUSTI$(CAPLI$, 32), 53, 32)
    '
    If TRIM$(Z$) <> "" Then
       If Left$(DAPLI$, 6) <> "Modulo" Then
          For KKJ& = 1 To FORDERAC.MSF.Rows - 1
             If TRIM$(FORDERAC.MSF.TextMatrix(KKJ&, 2)) = TRIM$(CAPLI$) Then
                Exit Sub      ' YA ESTA EN LA LISTA
             End If
'             If Len(CAPLI$) = 12 Then
'                If MTRIM$(FORDERAC.MSF.TextMatrix(kkj&, 2)) = CAPLI$ Then
'                   Exit Sub
'                End If
'             End If
             '
             If Len(TRIM$(Mid$(FORDERAC.List2.List(KKJ& - 1), 53, 32))) = 12 Then
                If Left$(CAPLI$, 12) = TRIM$(Mid$(FORDERAC.List2.List(KKJ& - 1), 53, 32)) Then
                   TTXX$ = FORDERAC.List2.List(KKJ& - 1)
                   Mid$(TTXX$, 53, 32) = CAPLI$
                   FORDERAC.List2.List(KKJ& - 1) = TTXX$
                   Exit Sub
                End If
             End If
          Next KKJ&
          'FORDERAC.List2.AddItem Z$
          FORDERAC.MSF.Rows = KKJ& + 1
          FORDERAC.MSF.TextMatrix(KKJ&, 1) = TRIM$(DAPLI$)
          FORDERAC.MSF.TextMatrix(KKJ&, 4) = TRIM$(CAPLI$)
       End If
    End If
End Sub
Sub AGREGALIST(CAPLI$, DAPLI$)
    '

    If Left$(CAPLI$, 1) = "*" Then CAPLI$ = Mid$(CAPLI$, 2)
    If Left$(CAPLI$, 1) = "?" Then CAPLI$ = Mid$(CAPLI$, 2)
    If UCase$(Left$(CAPLI$, 16)) = "TRAZAB07/BORECEP" Then CAPLI$ = "TRAZAB07/BORECEP"
    If UCase$(Left$(CAPLI$, 15)) = "VENGES07/PEDIDO" Then CAPLI$ = "VENGES07/PEDIDO"
    '
    Z$ = Space$(128)
    Call REPLA(Z$, DAPLI$, 1, 36)
    Call REPLA(Z$, "      ", 45, 6)
    Call REPLA(Z$, AJUSTI$(CAPLI$, 32), 53, 32)
    '
    If TRIM$(Z$) <> "" Then
       If Left$(DAPLI$, 6) <> "Modulo" Then
          For KKJ& = 1 To FORDERAC.List2.ListCount

             If TRIM$(Mid$(FORDERAC.List2.List(KKJ& - 1), 53, 32)) = TRIM$(CAPLI$) Then
                Exit Sub      ' YA ESTA EN LA LISTA
             End If
             If Len(CAPLI$) = 12 Then
                If Mid$(FORDERAC.List2.List(KKJ& - 1), 53, 12) = CAPLI$ Then
                   Exit Sub
                End If
             End If
             '
             If Len(TRIM$(Mid$(FORDERAC.List2.List(KKJ& - 1), 53, 32))) = 12 Then
                If Left$(CAPLI$, 12) = TRIM$(Mid$(FORDERAC.List2.List(KKJ& - 1), 53, 32)) Then
                   TTXX$ = FORDERAC.List2.List(KKJ& - 1)
                   Mid$(TTXX$, 53, 32) = CAPLI$
                   FORDERAC.List2.List(KKJ& - 1) = TTXX$
                   Exit Sub
                End If
             End If
          Next KKJ&
          FORDERAC.List2.AddItem Z$
       End If
    End If
End Sub


Sub CARGADER(CAPLI$, DR1$, DR2$)
    '
    If Left$(CAPLI$, 1) = "*" Then CAPLI$ = Mid$(CAPLI$, 2)
    If Left$(CAPLI$, 1) = "?" Then CAPLI$ = Mid$(CAPLI$, 2)
    '
    VUELTA% = 0
10  For JJI& = 1 To FORDERAC.List2.ListCount
      Z$ = FORDERAC.List2.List(JJI& - 1)
      If TRIM$(Mid$(Z$, 53, 32)) = CAPLI$ Then
         Call REPLA(Z$, DR1$, 45, 2)
         Call REPLA(Z$, DR2$, 49, 2)
         FORDERAC.List2.List(JJI& - 1) = Z$
         Exit Sub
      End If
    Next JJI&
    '
    ' NO ENCONTRO EL DERECHO
    If VUELTA% = 0 Then
       VUELTA% = 1
       Call AGREGALIST(CAPLI$, "Aplicacion " + CAPLI$)
       GoTo 10
    End If
    '
End Sub
'
Sub ACLISDER(USEL%, DEUSER$)
    '
    Call ACLISDER2(USEL%, DEUSER$)
    GoTo 90
    Screen.MousePointer = 11
    FORDERAC.List1.Clear
    FORDERAC.List2.Clear
    FORDERAC.Label2.Caption = TRIM$(Str$(USEL%))
    FORDERAC.Label4.Caption = TRIM$(DEUSER$)
    '
    ' CARGAR LISTA DE TODAS LAS APLICACIONES DEL MDB
    Call ABRETACON
    SQLX$ = "SELECT * FROM INDI_APLI "
    Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenSnapshot)
    With PARATEMP
      Do While Not .EOF
         CAPLI$ = AJUSTI$(SAFETEXT$(!COD_APLI), 32)
         DAPLI$ = TRIM$(SAFETEXT$(!DES_APLI))
         Call AGREGALIST(CAPLI$, DAPLI$)
      .MoveNext
      Loop
    End With
    PARATEMP.Close
    BaDaConsul.Close
    '
    UJ& = 0
    For U& = 1 To ULTREG&("INDIAPLI")
       X$ = REGLEIDO$("INDIAPLI", U&)
       CAPLI$ = TRIM$(Left$(X$, 12))
       DAPLI$ = Mid$(X$, 13, 36)
       Call AGREGALIST(CAPLI$, DAPLI$)
    Next U&
    '
    ' CARGAR DERECHOS DE ACCESO DEL MDB
    Screen.MousePointer = 11
'    FORDERAC.Show 1
    Call ABRETACON
    SQLX$ = "SELECT * FROM DER_ACCE "
    SQLX$ = SQLX$ + " WHERE NUM_USER = " & USEL%
    Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenSnapshot)
    With PARATEMP
      Do While Not .EOF
         CAPLI$ = SAFETEXT$(!COD_APLI)
         DR1$ = "  ": If XVALO(!DER_READ) > 0 Then DR1$ = "Si"
         DR2$ = "  ": If XVALO(!DER_WRITE) > 0 Then DR2$ = "Si"
         Call CARGADER(CAPLI$, DR1$, DR2$)
39    .MoveNext
      Loop
    End With
    PARATEMP.Close
    '

'    FORDERAC.Show 1
    ' CARGAR DERECHOS DEL DAT
    Screen.MousePointer = 11
    v& = 0
    For U& = 1 To ULTREG&("FLEXACC")
        X$ = REGLEIDO$("FLEXACC", U&)
        UI% = CVI(Left$(X$, 2))
        If CVI(Left$(X$, 2)) = USEL% Then
           CAPLI$ = TRIM$(UCase$(Mid$(X$, 3, 8)))
           DR1$ = "  ": If Mid$(X$, 11, 1) = "S" Then DR1$ = "Si"
           DR2$ = "  ": If Mid$(X$, 12, 1) = "S" Then DR2$ = "Si"
           Call CARGADER(CAPLI$, DR1$, DR2$)
        End If
49  Next U&
    '
    ' TRASLADA A LIST1 PARA ORDENAR
    FORDERAC.List1.Clear
    For JJI& = 1 To FORDERAC.List2.ListCount
      Z$ = FORDERAC.List2.List(JJI& - 1)
      FORDERAC.List1.AddItem Z$
    Next JJI&
90  Screen.MousePointer = 1
    '
End Sub
'
Sub GRATRACO(RAZO$, EMPRE$, SE$, LU$, QWERTY%, CCI$(), EJU%, menu%)
    '
    EJU% = EJER%
    LU$ = LUDAT$
    Dim CCP$(10)

200 GoSub 32000

    n% = FILEOPEN%("TRACOMO", 2, 0)
            Print #n%, RAZO$
            Print #n%, EMPRE$
            Print #n%, SE$
            Print #n%, LU$
            Print #n%, QWERTY%
                   For i% = 1 To 32
                        Print #n%, CCI$(i%)
                   Next i%
            Print #n%, 0
            Print #n%, EJU%
            Print #n%, menu%
            Print #n%, TC1
            Print #n%, CCI$(0)
            Close #n%: n% = 0
    Exit Sub

32000 n% = FreeFile
      Dim PAR As String * 8
      GoSub 32026
      If CTX% = 1 Then GoTo 32020
32002 n% = FILEOPEN%("SVD0200.EXE", 0, 8)
32006 If LOF(n%) <= 128 Then GoSub 32038: GoTo 32020
32008 For II% = 1 To 32
32010 Get #n%, II%, PAR$
      PAR1$ = PAR$
32012 '
32014 JJ% = 9: If InStr(PAR1$, Chr$(255)) > 0 Then JJ% = InStr(PAR1$, Chr$(255))
32016 CCI$(II%) = Left$(PAR1$, JJ% - 1)
32018 Next II%
32020 Close #n%: n% = 0
32022 Return
32024 '
32026 CTX% = 0
32028 n% = FILEOPEN%("SVD0200.EXE", 0, 8)
      If LOF(n%) < 16 Then GoTo 32036
32030 Close #n%: n% = 0
32032 GoTo 32048
32034 '
32036 Close #n%: n% = 0
32038 'Kill "SVD0200.EXE"
32040 'Data 27 - 64 - 27 - 120 - 0 - 18, 12, 27 - 50, 27 - 48, 27 - 51, 27 - 65, 27 - 67, 27 - 118, 8, 27 - 80, 27 - 77, 27 - 15, 18, 27 - 87 - 1, 27 - 87 - 0, 27 - 71, 27 - 72, 124, 80, 150
32042 'RESTORE 32040
32044 'For JJ% = 1 To 20: Read CCI$(JJ%): GoSub 32052: GoSub 32084: Next JJ%
32046 CTX% = 1
32048 Return
32050 '
32052 For X2% = 1 To 8
32054 CCP$(X2%) = ""
32056 Next X2%
32058 X2% = 1
32060 '
32062 CON11$ = CCI$(JJ%)
32064 For X% = 1 To Len(CON11$)
32066 X1$ = Mid$(CON11$, X%, 1)
32068 If X1$ = " " Then GoTo 32078
32070 If X1$ = "-" Then X2% = X2% + 1: GoTo 32076
32072 If CCP$(X2%) = "" Then CCP$(X2%) = X1$: GoTo 32076
32074 CCP$(X2%) = CCP$(X2%) + X1$
32076 Next X%
32078 '
32080 Return
32082 '
32084 B$ = "": If JJ% = 20 Then B$ = MKI$(XVALO(CCP$(1))): GoTo 32094
32086 For XX% = 1 To 8
32088 BX$ = "": If Len(CCP$(XX%)) > 0 Then BX$ = Chr$(XVALO(CCP$(XX%)))
32090 B$ = B$ + BX$
32092 Next XX%
32094 CCI$(JJ%) = B$
32096 Return
32098 '
'
End Sub
'

Sub ACLISDER2(USEL%, DEUSER$)
    '
    
    Screen.MousePointer = 11
    FORDERAC.List1.Clear
    FORDERAC.List2.Clear
    FORDERAC.Label2.Caption = TRIM$(Str$(USEL%))
    FORDERAC.Label4.Caption = TRIM$(DEUSER$)
    '
    


    SQLX$ = "SELECT  I.COD_APLI AS COD_APLI, D.NUM_USER, i.DES_APLI"
    SQLX$ = SQLX$ + " ,Case D.DER_READ"
    SQLX$ = SQLX$ + " WHEN  1 THEN 'Si'"
    SQLX$ = SQLX$ + " WHEN  0 THEN ''"
    SQLX$ = SQLX$ + " End"
    SQLX$ = SQLX$ + " AS LECTURA"
    SQLX$ = SQLX$ + " ,Case D.DER_WRITE"
    SQLX$ = SQLX$ + "  WHEN  1 THEN 'Si'"
    SQLX$ = SQLX$ + "  WHEN  0 THEN ''"
    SQLX$ = SQLX$ + "  End"
    SQLX$ = SQLX$ + " AS ESCRITURA"
    SQLX$ = SQLX$ + "  FROM  INDI_APLI I LEFT JOIN DER_ACCE D WITH(NOLOCK)"
    SQLX$ = SQLX$ + " ON D.COD_APLI = I.COD_APLI"
    SQLX$ = SQLX$ + " order by i.COD_APLI"
    

    Set rs = READSET(SQLX$)
    With rs
       Do While Not .EOF
         DESAPLI$ = SAFETEXT$(!DES_APLI)
         COAPLI$ = SAFETEXT$(!COD_APLI)
         If COAPLI$ <> COAPLI_ANT$ Then
            J& = J& + 1
            FORDERAC.MSF.Rows = J& + 1
            FORDERAC.MSF.TextMatrix(J&, 1) = DESAPLI$
            FORDERAC.MSF.TextMatrix(J&, 2) = ""
            FORDERAC.MSF.TextMatrix(J&, 3) = ""
            FORDERAC.MSF.TextMatrix(J&, 4) = COAPLI$
          End If
          If XVALO(!NUM_USER) = USEL% Then
              FORDERAC.MSF.TextMatrix(J&, 2) = SAFETEXT$(!LECTURA)
              FORDERAC.MSF.TextMatrix(J&, 3) = SAFETEXT$(!ESCRITURA)
           End If
          COAPLI_ANT$ = COAPLI$
'         CAPLI$ = AJUSTI$(SAFETEXT$(!COD_APLI), 32)
'         DAPLI$ = TRIM$(SAFETEXT$(!DES_APLI))
         .MoveNext
       Loop
    End With
    rs.Close
    Set rs = Nothing
'    For U& = 1 To ULTREG&("INDIAPLI")
'       X$ = REGLEIDO$("INDIAPLI", U&)
'       CAPLI$ = TRIM$(Left$(X$, 12))
'       DAPLI$ = Mid$(X$, 13, 36)
'            J& = J& + 1
'            FORDERAC.MSF.Rows = J& + 1
'            FORDERAC.MSF.TextMatrix(J&, 1) = DAPLI$
'            FORDERAC.MSF.TextMatrix(J&, 2) = ""
'            FORDERAC.MSF.TextMatrix(J&, 3) = ""
'            FORDERAC.MSF.TextMatrix(J&, 4) = CAPLI$
'
''       Call AGREGALIST(CAPLI$, DAPLI$)
'    Next U&
    'ORDENO LA GRILLA POR LA COLUMNA DE LA DESCRIPCION
    FORDERAC.MSF.COL = 1
    FORDERAC.MSF.Sort = 1

    Exit Sub
    
    
    
    
'    ' CARGAR LISTA DE TODAS LAS APLICACIONES DEL MDB
'    Call ABRETACON
'    SQLX$ = "SELECT * FROM INDI_APLI "
'    Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenSnapshot)
'    With PARATEMP
'         CAPLI$ = AJUSTI$(SAFETEXT$(!COD_APLI), 32)
'         DAPLI$ = TRIM$(SAFETEXT$(!DES_APLI))
'         Call AGREGALIST(CAPLI$, DAPLI$)
'      .MoveNext
'      Loop
'    End With
'    PARATEMP.Close
'    '
'    UJ& = 0
''    For U& = 1 To ULTREG&("INDIAPLI")
''       X$ = REGLEIDO$("INDIAPLI", U&)
''       CAPLI$ = TRIM$(Left$(X$, 12))
''       DAPLI$ = Mid$(X$, 13, 36)
''       Call AGREGALIST(CAPLI$, DAPLI$)
''    Next U&
'    '
'    ' CARGAR DERECHOS DE ACCESO DEL MDB
'    Screen.MousePointer = 11
'    FORDERAC.Show 1
'    Call ABRETACON
'    SQLX$ = "SELECT * FROM DER_ACCE "
'    SQLX$ = SQLX$ + " WHERE NUM_USER = " & USEL%
'    Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenSnapshot)
'    With PARATEMP
'      Do While Not .EOF
'         CAPLI$ = SAFETEXT$(!COD_APLI)
'         DR1$ = "  ": If XVALO(!DER_READ) > 0 Then DR1$ = "Si"
'         DR2$ = "  ": If XVALO(!DER_WRITE) > 0 Then DR2$ = "Si"
'         Call CARGADER(CAPLI$, DR1$, DR2$)
'39    .MoveNext
'      Loop
'    End With
'    PARATEMP.Close
'    '
'
'    FORDERAC.Show 1
'    ' CARGAR DERECHOS DEL DAT
'    Screen.MousePointer = 11
'    V& = 0
'    For U& = 1 To ULTREG&("FLEXACC")
'        X$ = REGLEIDO$("FLEXACC", U&)
'        UI% = CVI(Left$(X$, 2))
'        If CVI(Left$(X$, 2)) = USEL% Then
'           CAPLI$ = TRIM$(UCase$(Mid$(X$, 3, 8)))
'           DR1$ = "  ": If Mid$(X$, 11, 1) = "S" Then DR1$ = "Si"
'           DR2$ = "  ": If Mid$(X$, 12, 1) = "S" Then DR2$ = "Si"
'           Call CARGADER(CAPLI$, DR1$, DR2$)
'        End If
'49  Next U&
'    '
'    ' TRASLADA A LIST1 PARA ORDENAR
'    FORDERAC.List1.Clear
'    For JJI& = 1 To FORDERAC.List2.ListCount
'      Z$ = FORDERAC.List2.List(JJI& - 1)
'      FORDERAC.List1.AddItem Z$
'    Next JJI&
    Screen.MousePointer = 1
    '
End Sub

Sub FUERASYS()
    '
    If TRIM(Dir(LUCOM$ + "*.*")) <> "" Then
       Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
    End If
    Call CIERRARCH("*.*")
    Call LIBARCH("*.*")
    Call BLOQEXE("*")
    Close
    End
    '
End Sub
'
Sub BUSER()
    '
    Static CAIN%
    CLA$ = Left$(UCase$(CLAVE$) + Space$(12), 12)
    '
    If TRIM$(CLA$) = "???" Then
      If CLACONTRA% = 1 Then
        HOII% = HOY(HOS$)
        Has = HOII% + 30
        Des = HOII
        While DIENTRE%(Des, Has) < 30: Has = Has + 1: Wend
        FLAB% = Has
        XYZT$ = REGLEIDO$("MASTER", 1)
        Call REPLA(XYZT$, MKI$(FLAB%), 122, 2)
        Call GRAREG("MASTER", XYZT$, 1)
        Call CIERRARCH("MASTER")
        Call GRACONTROL("FLEXOFT", "LMTBNMNT", Space$(128))
        QWERTY% = 101
        USNBR% = 101
        USERNAME$ = "Supervisor General"
        Call CARDATRAM
        Exit Sub
      End If
    End If
    '
    If EXISTE%("USER.DAT") < 1 Then
        QWERTY% = 101
        USNBR% = 101
        USERNAME$ = "Supervisor General"
        Call CARDATRAM
        Exit Sub
    End If
    '
    Dim F0 As String * 64
    NX5% = FILEOPEN%("USER.DAT", 9, 64)
    '
    USNBR% = 0
    USERNAME$ = ""
    CONTRA$ = ""
    '
    CAUS% = LOF(NX5%) / 64
    MNGR% = 0
    For JJ% = 1 To LOF(NX5%) / 64
         Get #NX5%, JJ%, F0$
         Ident$ = Mid$(F0$, 31, 12): If Left$(Ident$, 1) = "@" Then Ident$ = DESENCRI$(Mid$(F0$, 32, 11)) + " "
         DENO$ = Left$(F0$, 30): If Left$(DENO$, 1) = "@" Then DENO$ = DESENCRI$(Mid$(F0$, 2, 29)) + " "
         CLAV$ = Mid$(F0$, 43, 12): If Left$(CLAV$, 1) = "@" Then CLAV$ = DESENCRI$(Mid$(F0$, 44, 11)) + " "
         If CLA$ = Ident$ Then
              NUSER% = JJ%
              FLESCRIT.Text2.TEXT = DENO$
              DENUS$ = DENO$
              FLESCRIT.Text3.Enabled = True
              FLESCRIT.Text3.SetFocus
              Close #NX5%: NX5% = 0
              Exit Sub
         End If
    Next JJ%
    '
    Close #NX5%: NX5% = 0
    MsgBox "USUARIO NO DEFINIDO"
    CAIN% = CAIN% + 1
    If CAIN% > 3 Then
        MsgBox "Superado Limite Permitido Intentos de Acceso !!!"
        Call FUERASYS: Close: End
    End If
    FLESCRIT.Text1.TEXT = ""
    '
End Sub

Sub BUCLA()
    '
    Static CAIN%
    CLA$ = Left$(UCase$(CONTRA$) + Space$(12), 12)
    '
    If TRIM$(CLA$) = TRIM$(CLAV$) Then
           USNBR% = NUSER
           USERNAME$ = DENUS$
           If USNBR% = 1 Then USNBR% = 101
           Call CARDATRAM
        Else
           FLESCRIT.Image5.Visible = True
           CAIN% = CAIN% + 1
           If CAIN% > 3 Then
              MsgBox "Superado Limite Permitido Intentos de Acceso !!!"
              Call FUERASYS
           End If
           '
           Call MENSERR(24, "Contraseña Incorrecta")
           FLESCRIT.Image5.Visible = False
           FLESCRIT.Text3.TEXT = ""
           CONTRA$ = ""
           FLESCRIT.Text3.SetFocus
    End If
    '
End Sub
'

Function CLACONTRA%()
   '
   CLACO% = 0
   CLAX$ = ""
   Randomize Timer
10 PPX% = Int(Rnd * 63)
   If PPX% < 48 Or PPX% > 57 Then GoTo 10
   CLAX$ = CLAX$ + Chr$(PPX%)
   If Len(CLAX$) < 6 Then GoTo 10
   '
   RESPUES$ = CONTRACLA$(CLAX$)
   VDEF$ = CLAX$ + Space$(6)
   CTRS$ = OBJPLA$("CLACONTRA/Santo y Seña", VDEF$)
   If CTRS$ = "" Then GoTo 99
   CONTRASE$ = Mid$(CTRS$, 7, 6)
   If CONTRASE$ = RESPUES$ Then
     CLACO% = 1
   End If
   '
99 CLACONTRA% = CLACO%
   '
End Function

Sub GENBACKUP()
   '
   Static YABAK%
   If YABAK% > 0 Then Exit Sub
   '
   DoEvents
   If Control$(IDENTER$, "SUPRIBAK") = "SI" Then Exit Sub
   '
   Screen.MousePointer = 11
   ULOWP$ = LUCOM$
   If Left$(JOBID$, 2) = "WF" Then
      Exit Sub
      'ULOWP$ = App.Path + "\"
   End If
   '
   VENTBAK$ = TRIM$(UCase$(Control$(IDENTER$, "VENTBAK"))) ' = "MINIMIZ" Then
      If VENTBAK$ <> "MINIMIZ" Then VENTBAK$ = "NORMAL"
   UTIBAK$ = TRIM$(UCase$(Control$(IDENTER$, "UTIBACK")))
      If UTIBAK$ <> "PKZIP" Then UTIBAK$ = "ARJ A -R"    '= "PKZIP"
   EXTEBAK$ = "ARJ": If UTIBAK$ = "PKZIP" Then EXTEBAK$ = "ZIP"
    '
   YABAK% = 1
   CUDI$ = App.Path
   LUBA$ = LUBAS$
   HOII% = HOY(HOS$)
   '
   LUDA$ = LUDAT$
   If Mid$(LUDA$, 2, 1) = ":" Then LUDA$ = Mid$(LUDA$, 3)
   If Left$(LUDA$, 2) = "\\" Then LUDA$ = Mid$(LUDA$, 2)
   '
   On Error Resume Next
   'MkDir "C:\FLEXOFT"
   MkDir ULOWP$ + "BACKUPS"
   LUDACO$ = LUDA$
   PPX1% = 1
   While Len(LUDACO$) > 0
     PPX1% = InStr(1 + PPX1%, LUDACO$, "\")
     If PPX1% < 1 Then GoTo 15
     'MkDir "C:\FLEXOFT\BACKUPS" + Left$(LUDACO$, PPX1% - 1)
     MkDir ULOWP$ + "BACKUPS" + Left$(LUDACO$, PPX1% - 1)
   Wend
15 On Error GoTo 0
   '
   'VERIFICA EXISTENCIA BACKUP DEL MES ANTERIOR
   FEX = DIANTE(COMESACT)
   FMXX$ = FECHATEX$(FEX)
   AMESJ$ = UCase$(ABREMESAN$(12 * XVALO(Mid$(FMXX$, 7, 2)) + XVALO(Mid$(FMXX$, 4, 2)), 0))
   AMESI$ = REPLACAR$(AMESJ$, " ", "")
   ABAK$ = ULOWP$ + "BACKUPS" + LUDA$ + "BKM" + AMESI$
   If EXISTE%(ABAK$ + "." + EXTEBAK$) < 1 Then
     CRSPX$ = "de " + ABREMESAN$(12 * XVALO(Mid$(FMXX$, 7, 2)) + XVALO(Mid$(FMXX$, 4, 2)), 1)
     GoTo 20
   End If
   '
   DSX% = DIASEM%(HOII%)
   ADIAS$ = UCase$(Left$(DSEM$(HOII%), 5))
   If DSX% = 5 Then ' PARA EL VIERNES
     DMX% = 1 + Int((XVALO(Left$(HOS$, 2)) - 1) / 7)
     ASEMA$ = "-0" + TRIM$(Str$(DMX%))
     ABAK$ = ULOWP$ + "BACKUPS" + LUDA$ + "BKSEM" + ASEMA$
     If EXISTE%(ABAK$ + "." + EXTEBAK$) < 1 Then
       CRSPX$ = "de Semana Nro. " + TRIM$(Str$(DMX%))
       GoTo 20
     End If
   End If
   '
   ' borra los ARJTEMP
   On Error Resume Next
   Kill ULOWP$ + "BACKUPS" + LUDA$ + "ARJTEMP.*"
   On Error GoTo 0
   '
   ABAK$ = ULOWP$ + "BACKUPS" + LUDA$ + "BK-" + ADIAS$
   If EXISTE%(ABAK$ + "." + EXTEBAK$) > 0 Then
     TTXX$ = FileDateTime(ABAK$ + "." + EXTEBAK$)
     If Int(CDbl(FileDateTime(ABAK$ + "." + EXTEBAK$))) = Int(CDbl(Now)) Then
       TTXX$ = "Ya Existe Archivo de Respaldo del " + HOS$ + "\Generado el " + TTXX$
       TTXX$ = TTXX$ + "\  \Opciones de Trabajo:\\Cancelar Operación\Sobre-Escribir Respaldo"
       If COMALTER%(TTXX$) <> 2 Then
         GoTo 99
       End If
     End If
   End If
   CRSPX$ = "del " + HOS$
   '
20 cmd$ = ""
   If Right$(LUBA$, 1) = "\" Then LUBA$ = Left$(LUBA$, Len(LUBA$) - 1)
   If Mid$(LUBA$, 2, 1) = ":" Then
     cmd$ = cmd$ + Left$(LUBA$, 2) + CRLF$
     LUBA$ = Mid$(LUBA$, 3)
   End If
   '
   cmd$ = cmd$ + "CHDIR " + LUBA$ + " " + CRLF$
   cmd$ = cmd$ + UTIBAK$ + " " + ABAK$ + " " + LUDAT$ + "*.MDB " + CRLF$
   cmd$ = cmd$ + UTIBAK$ + " " + ABAK$ + " " + LUDAT$ + "*.DAT " + CRLF$
   cmd$ = cmd$ + UTIBAK$ + " " + ABAK$ + " " + LUDAT$ + "*.X* " + CRLF$
   cmd$ = cmd$ + UTIBAK$ + " " + ABAK$ + " " + LUDAT$ + "*.OLD " + CRLF$
   cmd$ = cmd$ + UTIBAK$ + " " + ABAK$ + " " + LUDAT$ + "*.RFS " + CRLF$
   cmd$ = cmd$ + UTIBAK$ + " " + ABAK$ + " " + LUDAT$ + "*.FRM " + CRLF$
   If Mid$(CUDI$, 2, 1) = ":" Then
     cmd$ = cmd$ + Left$(CUDI$, 2) + CRLF$
     CUDI$ = Mid$(CUDI$, 3)
   End If
   '
   cmd$ = cmd$ + "CHDIR " + CUDI$ + CRLF$
   cmd$ = cmd$ + "ECHO OFF" + CRLF$
   cmd$ = cmd$ + "ECHO Fin de la Operacion de Respaldo" ' + CRLF$
   If VENTBAK$ <> "MINIMIZ" Then
     ' CMD$ = CMD$ + "PAUSE " + CRLF$
     cmd$ = cmd$ + CRLF$ + "PAUSE "
   End If
   Call DOSCOMMAND(cmd$)
'   CMD$ = CMD$ + "EXIT " + CRLF$
'   Call SAVEFILE(LUCOM$ + "DOSCOMM.BAT", CMD$)
   '
   'Call COMUNI("Preparando Generación de Respaldo " + CRSPX$)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   MODOBA = vbNormalFocus: If VENTBAK$ = "MINIMIZ" Then MODOBA = vbMinimizedFocus
   On Error Resume Next
   SBX1 = Shell(LUCOM$ + "DOSCOMM.BAT", MODOBA)
   On Error GoTo 0
   '
99 Screen.MousePointer = 1
   '
End Sub

