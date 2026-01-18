Attribute VB_Name = "FLEXOFT1"
Public VERIPAN%
'

Sub Main()
    '
    If VERIPAN% < 1 Then
       If Screen.Width < 12000 Then
          ANCHOPAN.Show 1
       End If
       VERIPAN% = 1
    End If
    '
' numero y nombre de usuario ***************************
    '
    USNBR% = 0
    USERNAME$ = ""
    '
' driver del menu principal activo *********************
    '
    MENUPRIN$ = ""
    CMD$ = TRIM$(Command$)
    If CMD$ <> "" Then
        MENUPRIN$ = CMD$
    End If
    '
' unidad local LUCOM$
    '
    LUCOM$ = "C:\FLEXOFT\"
    '
    If REACTIVA%("FLEXOFT") > 0 Then
      Call FINAL("")
      Close: End
      Exit Sub
    End If
    '
    FLEXOFTA.Show
    '
End Sub
'

Sub VERIDIR()
    '
    'Call SELUDAT
    '
    VERDEMO% = 0
    '
    FLEXOFTA.MousePointer = 11
    FLEXOFTA.List1.Clear
    FLEXOFTA.List1.Visible = True
    '
5   Call ADDLISVERI("Disco Local 'C:'")
    ARXX$ = "C:\"
    BRXX$ = CurDir
    On Error GoTo 6
    ChDir ARXX$
    ChDir BRXX$
    On Error GoTo 0
    GoTo 10
    '
6   Resume 7
7   On Error GoTo 0
    CRXX$ = REPLACAR(ARXX$, "\", "/")
    FLEXOFTA.List1.Visible = False
    TTXX$ = "No se ha Detectado la Presencia\de un Disco Rígido Local (Unidad C:).\ \"
    TTXX$ = TTXX$ + "Para su Funcionamiento, FLEXOFT\"
    TTXX$ = TTXX$ + "Requiere este Elemento en su Terminal.\  \"
    TTXX$ = TTXX$ + "Pulse Tecla <ENTER> o Haga Click\"
    TTXX$ = TTXX$ + "Sobre el Botón 'Aceptar Mensaje'\"
    TTXX$ = TTXX$ + "Para Abandonar la Ejecución FLEXOFT."
    Call COMUNI(TTXX$)
    Close: Call FUERASYS
    '
10  Call ADDLISVERI("Carpeta de Trabajo Actual")
    BRXX$ = TRIM$(CurDir)
    If InStr(BRXX$, " ") > 0 Then
        TTXX$ = "Su Sistema Aplicativo FLEXOFT\"
        TTXX$ = TTXX$ + "está Instalado en la Carpeta de Trabajo\  \"
        TTXX$ = TTXX$ + REPLACAR$(CurDir, "\", "/") + "\  \"
        TTXX$ = TTXX$ + "Es Posible que Algunas Aplicaciones\"
        TTXX$ = TTXX$ + "no Funcionen Correctamente.\  \"
        TTXX$ = TTXX$ + "Re-instale el Sistema FLEXOFT\"
        TTXX$ = TTXX$ + "en /FLEXOFT/APLICA."
        FLEXOFTA.List1.Visible = False
        If COMALTER%(TTXX$ + "\\Continuar\Cancelar") <> 1 Then
            Call FUERASYS
        End If
    End If
    '
20  Call ADDLISVERI("Escritura en Carpeta Actual")
    Dim RPRU As String * 128
    NX% = FreeFile
    On Error GoTo 21
    Open "NULL" For Random Access Read Write Shared As #NX% Len = 128
    RPRU$ = String$(128, "*")
    Put #NX%, 1, RPRU$
    Close #NX%: NX% = 0
    On Error GoTo 0
    GoTo 30
    '
21  Resume 22
22  On Error GoTo 0
    Close #NX%: NX% = 0
    VERDEMO% = 1
    ' Call MENSADEMO
    '
30  Call ADDLISVERI("Borrado de Archivos en Carpeta Actual")
    On Error GoTo 31
    Kill "NULL"
    On Error GoTo 0
    GoTo 40
    '
31  Resume 32
32  On Error GoTo 0
    VERDEMO% = 1
    ' Call MENSADEMO
    '
40  Call ADDLISVERI("Carpeta de Trabajo Local")
41  ARXX$ = "C:\FLEXOFT"
    BRXX$ = CurDir
    On Error GoTo 43
    ChDir ARXX$
    ChDir BRXX$
    On Error GoTo 0
    Call BLOQEXE(App.EXEName)
    GoTo 50
    '
43  Resume 44
44  On Error GoTo 0
    CRXX$ = REPLACAR(ARXX$, "\", "/")
    FLEXOFTA.List1.Visible = False
    TTXX$ = "No Existe Directorio " + CRXX$ + ".\ \"
    TTXX$ = TTXX$ + "Para su Funcionamiento, FLEXOFT\"
    TTXX$ = TTXX$ + "Requiere Instalar esta Carpeta\"
    TTXX$ = TTXX$ + "de Trabajo en su Disco Local.\  \"
    TTXX$ = TTXX$ + "Puede Crearla en este Momento\"
    TTXX$ = TTXX$ + "o Abandonar la Ejecución FLEXOFT.\\"
    TTXX$ = TTXX$ + "Crear Ahora\Abandonar"
    OPP% = COMALTER%(TTXX$)
    If OPP% = 1 Then
       MkDir ARXX$
       GoTo 41
    End If
    Close: Call FUERASYS
    '
50  Call ADDLISVERI("Carpeta Sistema Basico")
    LB$ = ""
    If EXISTE%("FLEXCFG.DAT") > 0 Then
        Dim FF0 As String * 128
        N1% = FILEOPEN%("FLEXCFG.DAT", 9, 128)
        '
        Get #N1%, 3, FF0$
        LB$ = TRIM$(Left$(FF0$, 46))
        Close #N1%: N1% = 0
    End If
    '
    If LB$ = "" Then
        LB$ = TRIM$(Mid$(UCase$(CurDir), 3))
    End If
    '
    If Len(LB$) > 0 Then
        If Right$(LB$, 1) <> "\" Then
            LB$ = LB$ + "\"
        End If
    End If
    LUBAS$ = LB$
    If Mid$(LUBAS$, 2, 1) <> ":" Then
        ULOAC$ = Left$(CurDir, 2)
        LUBAS$ = ULOAC$ + LUBAS$
    End If
    '
    Call CARDATRAM
    '
60  MENBAS$ = "Imposible Continuar Ejecución FLEXOFT.\  \"
    FINME$ = "\  \Consulte a su Soporte de Sistemas FLEXOFT."
    Call ADDLISVERI("Controlador Estructura de Archivos")
    If EXISTE%(LUBAS$ + "FLEXRFS.ZIP") = 1 Then GoTo 70
    FLEXOFTA.List1.Visible = False
    OPP% = COMALTER%(MENBAS$ + "Falta Controlador de Estructura de Archivos\en Carpeta " + REPLACAR$(LUBAS$, "\", "/") + FINME$ + "\\Aceptar Mensaje")
    Call FUERASYS
    '
70  Call ADDLISVERI("Utilitario 'PKUNZIP.EXE'")
    If EXISTE%(LUBAS$ + "PKUNZIP.EXE") = 1 Then GoTo 80
    FLEXOFTA.List1.Visible = False
    OPP% = COMALTER%(MENBAS$ + "Falta Utilitario 'PKUNZIP.EXE'\en Carpeta " + REPLACAR$(LUBAS$, "\", "/") + FINME$ + "\\Aceptar Mensaje")
    Call FUERASYS
    '
80  Call ADDLISVERI("Encabezado de Objetos 'OBJNAME.DRV'")
    If EXISTE%(LUBAS$ + "OBJNAME.DRV") = 1 Then GoTo 90
    FLEXOFTA.List1.Visible = False
    OPP% = COMALTER%(MENBAS$ + "Falta Controlador 'OBJNAME.DRV'\en Carpeta " + REPLACAR$(LUBAS$, "\", "/") + FINME$ + "\\Aceptar Mensaje")
    Call FUERASYS
    '
90  Call ADDLISVERI("Detalle de Objetos 'OBJDETA.DRV'")
    If EXISTE%(LUBAS$ + "OBJDETA.DRV") = 1 Then GoTo 100
    FLEXOFTA.List1.Visible = False
    OPP% = COMALTER%(MENBAS$ + "Falta Controlador 'OBJDETA.DRV'\en Carpeta " + REPLACAR$(LUBAS$, "\", "/") + FINME$ + "\\Aceptar Mensaje")
    Call FUERASYS
    '
100 Call VERIPIF
    '
End Sub
'

Sub VERIPIF()
    '
    NOEXI% = 0
    ARXX$ = "FLEXPIF.PIF": GoSub 10
    ARXX$ = "FLEXQB1.EXE": GoSub 10
    ARXX$ = "FLEXQB2.EXE": GoSub 10
    ARXX$ = "DOSCOMM.PIF": GoSub 10
    ARXX$ = "DOSCOMM.EXE": GoSub 10
    ARXX$ = "BRUN45.EXE": GoSub 10
    If NOEXI% > 0 Then Close: Call FUERASYS
    ARXX$ = "ARJ.EXE": GoSub 10
    Exit Sub
    '
10  Call ADDLISVERI("Instalación Utilitario " + ARXX$)
    If InStr(ARXX$, ".PIF") > 0 Then
        LCB$ = LUBAS$
        If EXISTE%(LCB$ + ARXX$) <> 1 Then
            NX1% = FreeFile
            Open LCB$ + "PIFFILES.DRV" For Random Access Read Shared As #NX1% Len = 8
            Dim PF As String * 8
            NX2% = FreeFile
            Open LCB$ + ARXX$ For Output As #NX2%
            For U& = 1 To LOF(NX1%) / 8
                Get #NX1%, U&, PF$
                If Left$(PF$, 7) = Left$(ARXX$, 7) Then
                    PY$ = Mid$(PF$, 8, 1)
                    Print #NX2%, PY$;
                End If
            Next U&
            Close #NX2%
            Close #NX1%
        End If
    End If
    '
    If EXISTE%("C:\FLEXOFT\" + ARXX$) <> 1 Then
        On Error Resume Next
        LCB$ = LUBAS$
        FileCopy LCB$ + ARXX$, "C:\FLEXOFT\" + ARXX$
        On Error GoTo 0
        If EXISTE%("C:\FLEXOFT\" + ARXX$) < 1 Then
           BRXX$ = REPLACAR$("C:\FLEXOFT\" + ARXX$, "\", "/")
           FLEXOFTA.List1.Visible = False
           MMXX% = COMALTER%("No se Verificó la Existencia del Programa Utilitario\  \'" + BRXX$ + "'.\  \La Aplicación no Podrá Ejecutarse Normalmente.\  \Consulte a su Soporte de Sistemas.\\Aceptar Mensaje")
           NOEXI% = NOEXI% + 1
        End If
    End If
    '
    Return
    '
End Sub
'

Sub ADDLISVERI(TTX$)
    '
    Static TANTU
    INCRETI = 0.3
    '
    FLEXOFTA.List1.AddItem TTX$
    FLEXOFTA.List1.ListIndex = (-1)
    FLEXOFTA.List1.Visible = True
    If FLEXOFTA.List1.ListCount >= 6 Then
        FLEXOFTA.List1.TopIndex = FLEXOFTA.List1.ListCount - 4
    End If
    FLEXOFTA.List1.Refresh
    '
    While Timer < TANTU + INCRETI: Wend
    TANTU = Timer
    '
End Sub
'

Sub VERIFLIM(MDXI%)
    '
    Static CTX%
    EPL$ = UCase$(EMPRELI$)
    'If Left$(EPL$, 6) = "AMBEST" Then Exit Sub
    'If InStr(EPL$, "BORENS") > 0 Then Exit Sub
    '
    If MDXI% = 1 Then CTX% = 0
    If CTX% = 0 Then
        Dim CHKY As String * 127
        Dim MSX As String * 128
10      FECHAC% = HOY(HOS$)
        LU$ = LUDAT$
        NX% = FILEOPEN%(LU$ + "MASTER.DAT", 0, 128)
        If LOF(NX%) < 128 Then
            HAS = FECHAC% + 60
            DES = FECHAC%
            While DIENTRE%(DES, HAS) < 60
                HAS = HAS + 1
            Wend
            FLIM% = HAS
            MSY$ = String$(128, 0)
            Mid$(MSY$, 124, 2) = MKI$(FLIM%)
            Mid$(MSY$, 126, 2) = MKI$(9876 - FLIM%)
            PPX% = 23
            For I% = 124 To 127
                PPX% = PPX% + (I% - 123) * Asc(Mid$(MSY$, I%, 1))
            Next I%
            PPX% = PPX% Mod 256
            Mid$(MSY$, 128, 1) = Chr$(PPX%)
            LSet MSX$ = MSY$
            Put #NX%, 1, MSX$
            '
          Else
            Get #NX%, 1, MSX$
            'If Mid$(MSX$, 124, 5) = String$(5, 0) Then
               'If FECHAC% < 8847 Then
                  HAS = FECHAC% + 60
                  DES = FECHAC%
                  While DIENTRE%(DES, HAS) < 60
                    HAS = HAS + 1
                  Wend
                  'FLIM% = HAS
FLIMI1% = FECHANUM("15/07/01")
FLIMI2% = FECHANUM("21/12/01")
FLIMI3% = FECHANUM("25/03/02")
FLIMI4% = 32767
If InStr(EMPRELI$, "CIME-SOI") > 0 Then
  If HOY(HOS$) > FLIMI2% - 25 Then
    If VERCD% <> 1 Then '
      'MsgBox "Verificar Fechas de Vencimiento"
    End If
  End If
End If
EPL$ = UCase$(EMPREAC$)
EPLI$ = TRIM$(UCase$(EMPRELI$))
If Left$(EPL$, 4) = "SABO" Then FLIM% = FLIMI4%
If Left$(EPL$, 7) = "LODIGIA" Then FLIM% = FLIMI1%
If Left$(EPL$, 4) = "NOVO" Then FLIM% = FLIMI4%
If Left$(EPL$, 7) = "YALTRES" Then FLIM% = FLIMI4%
If Left$(EPL$, 6) = "AMBEST" Then FLIM% = FLIMI4%
If Left$(EPL$, 5) = "FORJA" Then FLIM% = FLIM1%
If Left$(EPL$, 7) = "TALEMEC" Then FLIM% = FLIMI4%
If Left$(EPL$, 5) = "FERVI" Then FLIM% = FLIMI4%
If Left$(EPL$, 5) = "STADT" Then FLIM% = FLIMI1%
If Left$(EPL$, 7) = "NEUMANN" Then FLIM% = FLIMI4%
If InStr(EPLI$, "DOVAL") > 0 Then FLIM% = FLIMI4%
If InStr(EPLI$, "DEVOTO") > 0 Then FLIM% = FLIMI4%
If InStr(EPL$, "ROLLAN") > 0 Then FLIM% = FLIMI4%
If InStr(EPL$, "FAEM") > 0 Then FLIM% = FLIMI4%
If InStr(EPL$, "BORENS") > 0 Then FLIM% = FLIMI3%
If InStr(EPL$, "DOSIVAC") > 0 Then FLIM% = FLIMI4%
If InStr(EPL$, "FUINY") > 0 Then FLIM% = FLIMI4%
If InStr(EPL$, "ARCON") > 0 Then FLIM% = FLIMI4%
If InStr(EPL$, "MAINIERI") > 0 Then FLIM% = FLIMI4%
If InStr(EPL$, "SEALU") > 0 Then FLIM% = FLIMI4%
If InStr(EPL$, "DURALAMP") > 0 Then FLIM% = FLIMI4%
If InStr(EPLI$, "MAINIERI") > 0 Then FLIM% = FLIMI4%
If Left$(EPL$, 6) = "CENTRO" Then FLIM% = FLIMI4%
If InStr(EPL$, "DIPRO") > 0 Then FLIM% = FLIMI4%
If InStr(EPL$, "GRISFER") > 0 Then FLIM% = FLIMI4%
If InStr(EPLI$, "CIME-SOI") > 0 Then FLIM% = FLIMI4%
                  '
                  MSY$ = MSX$
                  Mid$(MSY$, 124, 2) = MKI$(FLIM%)
                  Mid$(MSY$, 126, 2) = MKI$(9876 - FLIM%)
                  PPX% = 23
                  For I% = 124 To 127
                    PPX% = PPX% + (I% - 123) * Asc(Mid$(MSY$, I%, 1))
                  Next I%
                  PPX% = PPX% Mod 256
                  Mid$(MSY$, 128, 1) = Chr$(PPX%)
                  LSet MSX$ = MSY$
                  Put #NX%, 1, MSX$
               'End If
            'End If
            '
        End If
        '
        Get #NX%, 1, MSX$
        FLIM% = CVI(Mid$(MSX$, 124, 2))
        '
        If FECHAC% < 8975 Then ' Corresponde a 15/07/2001
           GoTo 98
        End If
        '
        If FLIM% < FECHAC% Then GoTo 99
        'If CVI(Mid$(MSX$, 126, 2)) <> 9876 - FLIM% Then GoTo 99
        'PPX% = 23
        'For I% = 124 To 127
        '    PPX% = PPX% + (I% - 123) * Asc(Mid$(MSX$, I%, 1))
        'Next I%
        'PPX% = PPX% Mod 256
        'If Asc(Mid$(MSX$, 128, 1)) <> PPX% Then GoTo 99
        '
        
'VERIFICAR SERIVOL$
'
98      Close #NX%
        CTX% = 1
        Exit Sub
        '
99      JK$ = "": JK2$ = "": QQ% = 0
        SVL$ = SERIVOL$("C:")
        Randomize Timer
100     PPX% = Int(Rnd * 63)
        If PPX% < 48 Then GoTo 100
        If PPX% > 57 Then PPX% = PPX% + 7
        QQ% = QQ% + 1
        If QQ% Mod 2 = 1 Then
            JK$ = JK$ + Chr$(PPX%)
          Else
            JK2$ = JK2$ + Chr$(PPX%)
        End If
        If Len(JK$) = 4 Then JK$ = JK$ + ":"
        If Len(JK2$) = 4 Then JK2$ = JK2$ + ":"
        If Len(JK$) < 9 Or Len(JK2$) < 9 Then GoTo 100
        FEX = FLIM: FFF$ = FECHATEX$(FEX)
        FFY$ = Mid$(FFF$, 7, 2) + Mid$(FFF$, 4, 2) + Left$(FFF$, 2)
        FFZ$ = TRIM$(FFY$)
        '
        MENSA$ = "ATENCION: Sistema Inestable. Se ha Producido" + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
        MENSA$ = MENSA$ + "IQL Interrupt(" + FFZ$ + ") at FLIM_ID " + JK$ + Chr$(10) + Chr$(13)
        MENSA$ = MENSA$ + "SVL_ID " + SVL$ + " at Address " + JK2$ + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
        MENSA$ = MENSA$ + "Consulte Urgente a su Soporte de Sistemas."
        MsgBox MENSA$, 4096 + 16
        '
        On Error GoTo 200
        MODOAC% = 1
        NY% = FreeFile
        Open "A:FLEXCHK.DRV" For Random Access Read Write Shared As #NY% Len = 127
150     Get #NY%, 1953, CHKY$
        SVL1$ = DESENCRI$(Mid$(CHKY$, 2, 9))
        SVLA1$ = DESENCRI$(Mid$(CHKY$, 73, 9))
        If MODOAC% <> 2 Then SVLA$ = TRIM$(SERIVOL$("A:"))
        Close #NY%
        '
        If SVL$ <> SVL1$ Then GoTo 300
        If MODOAC% <> 2 Then
           If SVLA$ = "" Or SVLA$ = "0000-0000" Then GoTo 300
           If SVLA$ <> SVLA1$ Then GoTo 300
        End If
        '
160     PPXA% = 0
        For KK% = 1 To 94
            PPXA% = PPXA% + Asc(Mid$(CHKY$, KK%, 1)) - 128
        Next KK%
        If PPXA% <> CVI(Mid$(CHKY$, 95, 2)) Then GoTo 300
        FLIM% = CVI(Mid$(CHKY$, 92, 2))
        If FLIM% < FECHAC% Then GoTo 300
        '
        On Error GoTo 0
        Get #NX%, 1, MSX$
        MSY$ = MSX$
        Mid$(MSY$, 124, 2) = MKI$(FLIM%)
        Mid$(MSY$, 126, 2) = MKI$(9876 - FLIM%)
        PPX% = 23
        For I% = 124 To 127
            PPX% = PPX% + (I% - 123) * Asc(Mid$(MSY$, I%, 1))
        Next I%
        PPX% = PPX% Mod 256
        Mid$(MSY$, 128, 1) = Chr$(PPX%)
        LSet MSX$ = MSY$
        Put #NX%, 1, MSX$
        Close #NX%
        GoTo 10
        '
        '
200     Resume 210
210     On Error GoTo 0
        If EXISTE%("FLEXCHK.DRV") > 0 Then
           NY% = FreeFile
           Open "FLEXCHK.DRV" For Random Access Read Write Shared As #NY% Len = 127
           MODOAC% = 2
           GoTo 150
        End If
        '
300     On Error GoTo 0
        Close #NX%
        HAS = FLIM% + 11
        DES = FLIM%
        While DIENTRE%(DES, HAS) < 11
             HAS = HAS + 1
        Wend
        If FECHAC% > HAS Then
           Call FUERASYS
        End If
' CONTADOR DE ACCESOS
        '
    End If
    '
End Sub
'

Sub VERITER()
    '
    If EXISTE%(LUCOM$ + "IDENTER.DRV") < 1 Then
        If EXISTE%("C:\IDENTER.DRV") > 0 Then
            Call COPYFILE("C:\IDENTER.DRV", LUCOM$ + "IDENTER.DRV")
        End If
    End If
    '
    SVL$ = SERIVOL$("C:")
    SVL1$ = AJUSTI$(SVL$, 12)
    SVLX$ = ENCRIPTA$(SVL1$)
    '
    N1% = FILEOPEN%(LUCOM$ + "IDENTER.DRV", 0, 40)
    Dim IDXX As String * 40
    '
    If LOF(N1%) < 80 Then
       ID = Timer
       While ID > 9999.99
           ID = ID - 9000
       Wend
       ID = Int(100 * ID)
       IDX$ = Mid$(Str$(ID), 2)
       While Len(IDX$) < 6
           IDX$ = "0" + IDX$
       Wend
       '
       VDEF$ = Space$(52)
       Call REPLA(VDEF$, IDX$, 1, 6)
       Call REPLA(VDEF$, SVL$, 41, 12)
10     OBX$ = OBJPLA$("DATERMINAL", VDEF$)
       If OBX$ = "" Then
           Call FUERASYS
       End If
       IDT$ = TRIM$(Mid$(OBX$, 7, 32))
       If IDT$ = "" Then GoTo 10
       '
       NTID$ = Space$(38)
       Call REPLA(NTID$, Mid$(OBX$, 7, 32), 1, 32)
       Call REPLA(NTID$, Left$(OBX$, 6), 33, 6)
       IDENTER$ = TRIM$(Mid$(OBX$, 1, 6))
       NOMTER$ = TRIM$(Mid$(OBX$, 7, 32))
       '
       X$ = ENCRIPTA$(NTID$)
       SUDI% = 0: For I% = 1 To Len(X$)
                     SUDI% = SUDI% + Asc(Mid$(X$, I%, 1))
                  Next I%
       Y$ = String$(40, 0)
       Call REPLA(Y$, X$, 1, 38)
       Call REPLA(Y$, MKI$(SUDI%), 39, 2)
       IDXX$ = Y$
       Put #N1%, 1, IDXX$
       '
       SVL1$ = AJUSTI$(SVL$, 12)
       Y$ = String$(40, 0)
       Call REPLA(Y$, "B800", 1, 4)
       Call REPLA(Y$, MKS$(47104), 5, 4)
       Call REPLA(Y$, "SI", 9, 2)
       X$ = ENCRIPTA$(SVL1$)
       Call REPLA(Y$, X$, 11, 12)
       IDXX$ = Y$
       Put #N1%, 2, IDXX$
       Close #N1%: N1% = 0
       '
       Call COPYFILE(LUCOM$ + "IDENTER.DRV", "C:\IDENTER.DRV")
       '
    End If
    '
    If N1% > 0 Then
        Close #N1%: N1% = 0
    End If
    '
    If EXISTE%("C:\IDENTER.DRV") < 1 Then
       Call COPYFILE(LUCOM$ + "IDENTER.DRV", "C:\IDENTER.DRV")
    End If
    '
    N1% = FILEOPEN%(LUCOM$ + "IDENTER.DRV", 0, 40)
    Dim X1 As String * 40
    If LOF(N1%) < 80 Then
        Call MENSERR(24, "Falta Identificador de Terminal")
        Call FUERASYS
    End If
    '
    MCX% = 0
    Get #N1%, 1, X1$
    SUDI% = 0: For I% = 1 To 38
                        SUDI% = SUDI% + Asc(Mid$(X1$, I%, 1))
               Next I%
               If SUDI% <> CVI(Mid$(X1$, 39, 2)) Then
                   Call MENSERR(24, "Falta Identificador de Terminal")
                   Call FUERASYS
               End If
               '
               J1$ = DESENCRI$(Left$(X1$, 38))
               IDX$ = TRIM$(Mid$(J1$, 33, 6))
               NTX$ = TRIM$(Mid$(J1$, 1, 32))
               '
    If IDX$ = "" Then
        Call MENSERR(24, "Falta Identificador de Terminal")
        Call FUERASYS
    End If
    Close #N1%
    '
    IDENTER$ = IDX$
    NOMTER$ = NTX$
    '
    ' VERIFICAR CANTIDAD DE PUESTOS DE TRABAJO HABILITADOS
    If VERCD% = 1 Then
      PUETRA% = 1
      Exit Sub
    End If
    '
    Dim IDS As String * 96
    N1% = FILEOPEN%("IDENSYS.DRV", 0, 96)
    If LOF(N1%) < 96 * 3 Then
       PUETRA% = 1
       IDENSYS3$ = RANDSTRING$(94)
       Call REPLA(IDENSYS3$, Chr$(PUETRA%), 17, 1)
       SUCAR% = 0: For I% = 1 To Len(IDENSYS3$): SUCAR% = SUCAR% + Asc(Mid$(IDENSYS3$, I%, 1)) - 83: Next I%
       DIVERI3$ = MKI$(SUCAR%)
       '
       LSet IDS$ = IDENSYS3$ + DIVERI3$
       Put #N1%, 3, IDS$
    End If
    '
    Get #N1%, 3, IDS$
    IDENSYS3$ = Left$(IDS$, 94): SUCAR% = 0
    For I% = 1 To Len(IDENSYS3$)
       SUCAR% = SUCAR% + Asc(Mid$(IDENSYS3$, I%, 1)) - 83
    Next I%
    If CVI(Mid$(IDS$, 95, 2)) <> SUCAR% Then
       Call COMUNI("Puesto de Trabajo no Habilitado")
       Call FUERASYS
    End If
    '
    EPL$ = UCase$(EMPRELI$)
    If Left$(EPL$, 4) = "SABO" Then PUETRAX% = 50
    If Left$(EPL$, 7) = "LODIGIA" Then PUETRAX% = 8
    If Left$(EPL$, 4) = "NOVO" Then PUETRAX% = 8
    If Left$(EPL$, 7) = "YALTRES" Then PUETRAX% = 8
    If Left$(EPL$, 6) = "AMBEST" Then PUETRAX% = 4
    If Left$(EPL$, 7) = "TALEMEC" Then PUETRAX% = 2
    If Left$(EPL$, 5) = "FERVI" Then PUETRAX% = 16
    If Left$(EPL$, 5) = "STADT" Then PUETRAX% = 4
    If Left$(EPL$, 7) = "NEUMANN" Then PUETRAX% = 20
    If Left$(EPL$, 7) = "GRISFER" Then PUETRAX% = 3
    If InStr(EPL$, "MAINIERI") > 0 Then PUETRAX% = 99
    If InStr(EPL$, "DOVAL") > 0 Then PUETRAX% = 4
    If InStr(EPL$, "CIME-SOI") > 0 Then PUETRAX% = 99
    '
    PUETRA% = Asc(Mid$(IDENSYS3$, 17, 1))
    If PUETRAX% > PUETRA% Then PUETRA% = PUETRAX%
    '
    I1% = LOF(N1%) / 96
    For I% = I1% + 1 To 10
       Get #N1%, I%, IDS$
       If I% > I1% Then IDS$ = String$(96, 0)
       Put #N1%, I%, IDS$
    Next I%
    I1% = LOF(N1%) / 96
    '
    CAPUE% = 0
    For I% = 11 To I1%
       Get #N1%, I%, IDS$
       SUCAR% = 0
       For KI% = 1 To 94
          SUCAR% = SUCAR% + Asc(Mid$(IDS$, KI%, 1)) - 83
       Next KI%
       If CVI(Mid$(IDS$, 95, 2)) = SUCAR% Then
          SVLL$ = Mid$(IDS$, 8, 12)
          If SVLL$ = SVLX$ Then
             GoTo 99
          End If
          CAPUE% = CAPUE% + 1
          If CAPUE% >= PUETRA% Then GoTo 89
       End If
    Next I%
    '
    If CAPUE% < PUETRA% Then
       IDENSYS4$ = RANDSTRING$(94)
       Call REPLA(IDENSYS4$, SVLX$, 8, 12)
       SUCAR% = 0: For I% = 1 To Len(IDENSYS4$): SUCAR% = SUCAR% + Asc(Mid$(IDENSYS4$, I%, 1)) - 83: Next I%
       DIVERI4$ = MKI$(SUCAR%)
       '
       LSet IDS$ = IDENSYS4$ + DIVERI4$
       Put #N1%, I1% + 1, IDS$
       GoTo 99
    End If
    '
89  Close #N1%
    Call MENSERR(24, "Imposible Utilizar Este Puesto de Trabajo.\Supera Máximo Habilitado de " + TRIM$(Str$(PUETRA%)) + " Puestos.\  \Consulte a FLEXOFT para Ampliar su Licencia.")
    Call FUERASYS
    '
99  Close #N1%
    '
End Sub
'

Sub SELUDAT()
    '
    LCB$ = "C"
10  Call GRARAM(97, LCB$)
    On Error GoTo 800
    N1% = FreeFile
    Open LCB$ + ":\COMMAND.COM" For Input Shared As #N1%
    If N1% > 0 Then Close #N1%: N1% = 0
    On Error GoTo 0
    '
    If LEERAM$(97, 1) = LCB$ Then GoTo 11
    If LCB$ = "A" Then
        LCB$ = "C"
        Call GRARAM(97, LCB$)
        GoTo 11
    End If
    LCB$ = "A": GoTo 10
    '
11  LUBOOT$ = LCB$
    '
    VERCD% = 0
    On Error GoTo 15
    FILI% = FreeFile
    Open "FLEXMUL.DRV" For Random Access Read Write Shared As #FILI% Len = 79
    Close #FILI%
    GoTo 20
    '
15  Resume 16
16  On Error GoTo 0
    VERCD% = 1
    LD$ = "C:\FLEXOFT\DATOS\"
    GoTo 50
    '
20  On Error GoTo 0
    If EXISTE%("FLEXMUL.DRV") = 1 Then
        FORSELUDAT.List2.Clear
        FORSELUDAT.List2.Visible = True
        '
        NX% = FILEOPEN("FLEXMUL.DRV", 0, 79)
        Dim FMX As String * 79
        '
        For I% = 1 To LOF(NX%) / 79
           Get #NX%, I%, FMX$
           SUCAR% = 0
           For J% = 1 To 77
              SUCAR% = SUCAR% + Asc(Mid$(FMX$, J%, 1))
           Next J%
           If CVI(Mid$(FMX$, 78, 2)) = SUCAR% Then
              RASO$ = DESENCRI$(Left$(FMX$, 30))
              PACCE$ = DESENCRI$(Mid$(FMX$, 32, 46))
              Z$ = Space$(96)
              Call REPLA(Z$, RASO$, 1, 30)
              Call REPLA(Z$, PACCE$, 51, 46)
              FORSELUDAT.List2.AddItem Z$
           End If
        Next I%
        '
        Close #NX%
        If FORSELUDAT.List2.ListCount > 0 Then
            FORSELUDAT.List2.ListIndex = 0
            FORSELUDAT.Show 1
            Call FRESHALL
            Exit Sub
        End If
        '
    End If
    '
    LD$ = ""
    If EXISTE%("FLEXCFG.DAT") > 0 Then
        Dim FF0 As String * 128
        NX% = FILEOPEN%("FLEXCFG.DAT", 9, 128)
        Get #NX%, 1, FF0$
        LD$ = TRIM$(Mid$(FF0$, 51, 45))
        Close #NX%: NX% = 0
    End If
    '
    If Len(LD$) > 0 Then
        If Right$(LD$, 1) <> "\" Then
            LD$ = LD$ + "\"
        End If
    End If
    '
50  If LD$ = "" Then
        LD$ = "C:\FLEXOFT\DATOS\"
    End If
    '
    BRXX$ = CurDir
    ARXX$ = TRIM$(LD$)
    If Len(ARXX$) > 0 Then
        If Right$(ARXX$, 1) = "\" Then
            ARXX$ = Left$(ARXX$, Len(ARXX$) - 1)
        End If
    End If
    '
53  On Error GoTo 51
    ChDir ARXX$
    ChDir BRXX$
    On Error GoTo 0
    GoTo 60
    '
51  Resume 52
52  On Error GoTo 0
    TTXX$ = "Imposible Acceder a Carpeta de Datos\   \"
    TTXX$ = TTXX$ + REPLACAR$(ARXX$, "\", "/") + "\   \"
    TTXX$ = TTXX$ + "Puede Crear Ahora la Carpeta de Datos."
    FLEXOFTA.List1.Visible = False
    On COMALTER%(TTXX$ + "\\Crear Carpeta\Abandonar") GoTo 56
    Call FUERASYS
'
56  On Error GoTo 57
    MkDir ARXX$
    If VERCD% = 1 Then
      While EXTRACTFORM.Visible = True: Wend
      EXTRACTFORM.Label18.Caption = "Controladores"
      Screen.MousePointer = 11
      EXTRACTFORM.Show 1
    End If
    On Error GoTo 0
    GoTo 53
    '
57  Resume 58
58  On Error GoTo 0
    TTXX$ = "Imposible Crear la Carpeta\  \"
    TTXX$ = TTXX$ + ARXX$
    FLEXOFTA.List1.Visible = False
    On COMALTER%(TTXX$ + "\\Aceptar Mensaje") GoTo 59
59  GoTo 52
    '
60  If Mid$(LD$, 2, 1) = ":" Then LD$ = Mid$(LD$, 3)
    LUDAT$ = ULODATOS$ + TRIM$(LD$)
    '
    If VERCD% = 1 Then
      While EXTRACTFORM.Visible = True: Wend
      EXTRACTFORM.Label18.Caption = "Datos de Prueba"
      Screen.MousePointer = 11
      EXTRACTFORM.Show 1
    End If
    '
    Call CARDATRAM
    Exit Sub
    '
    '
800 LCB$ = LEERAM$(97, 1)
    If LCB$ = "C" Then
            Call GRARAM(97, "A")
          Else
            Call GRARAM(97, "C")
    End If
    Resume Next
    '
End Sub
'

Sub CARMULTIEM()
    '
    Call BLANARCH("!CARMULTI")
    If EXISTE%("FLEXMUL.DRV") = 1 Then
        '
        NX% = FILEOPEN("FLEXMUL.DRV", 0, 79)
        Dim FMX As String * 79
        '
        For I% = 1 To LOF(NX%) / 79
           Get #NX%, I%, FMX$
           SUCAR% = 0
           For J% = 1 To 77
              SUCAR% = SUCAR% + Asc(Mid$(FMX$, J%, 1))
           Next J%
           If CVI(Mid$(FMX$, 78, 2)) = SUCAR% Then
              RASO$ = DESENCRI$(Left$(FMX$, 30))
              PACCE$ = DESENCRI$(Mid$(FMX$, 32, 46))
              Z$ = Space$(96)
              Call REPLA(Z$, RASO$, 1, 30)
              Call REPLA(Z$, PACCE$, 31, 46)
              Call REGAPP("!CARMULTI", Z$)
           End If
        Next I%
        '
        Close #NX%
        Call CIERRARCH("!CARMULTI")
        '
    End If
    '
    If ULTREG&("!CARMULTI") < 1 Then
        '
        Randomize Timer
        INTEN% = 0
        '
5       CLAX$ = ""
10      PPX% = Int(Rnd * 63)
        If PPX% < 48 Or PPX% > 57 Then GoTo 10
        CLAX$ = CLAX$ + Chr$(PPX%)
        If Len(CLAX$) < 6 Then GoTo 10
        '
        RESPUES$ = ""
        For M% = 1 To 6
            POSI% = (5 * M%) Mod 6
            While POSI% < 1: POSI% = POSI% + 6: Wend
            VORI% = Val(Mid$(CLAX$, POSI%, 1))
            VTRA% = (VORI% + 3 * M%) Mod 10
            RESPUES$ = RESPUES$ + Mid$(Str$(VTRA%), 2)
        Next M%
        '
        VDEF$ = CLAX$ + Space$(6)
        CTRS$ = OBJPLA$("CLACONTRA/Configuración Multiempresa", VDEF$)
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
    End If
    '
    VTB% = VENTABU%("CARMULTI")
    If VTB% >= 0 Then
        '
        NX% = FILEOPEN("FLEXMUL.DRV", 0, 79)
        Dim FMY As String * 79
        '
        For U& = 1 To ULTREG&("!CARMULTI")
            Z$ = REGLEIDO$("!CARMULTI", U&)
            RASO$ = AJUSTI$(Left$(Z$, 30), 30)
            PACCE$ = AJUSTI$(Mid$(Z$, 31, 46), 46)
            If TRIM$(RASO$) <> "" Then
                If TRIM$(PACCE$) <> "" Then
                    Z1$ = String$(79, 0)
                    Call REPLA(Z1$, ENCRIPTA$(RASO$), 1, 30)
                    Call REPLA(Z1$, ENCRIPTA$(PACCE$), 32, 46)
                    SUCAR% = 0
                    For J% = 1 To 77
                       SUCAR% = SUCAR% + Asc(Mid$(Z1$, J%, 1))
                    Next J%
                    Call REPLA(Z1$, MKI$(SUCAR%), 78, 2)
                    V& = V& + 1
                    LSet FMY$ = Z1$
                    Put #NX%, V&, FMY$
                End If
            End If
        Next U&
        '
        For U& = V& + 1 To LOF(NX%) / 79
            LSet FMY$ = String$(79, 0)
            Put #NX%, U&, FMY$
        Next U&
        '
        Close #NX%
        Call CIERRARCH("!CARMULTI")
        '
    End If
    '
    Call COMUNI("Para Efectivizar la Nueva Configuración\Debe Re-Iniciarse el Sistema FLEXOFT")
    Call FUERASYS
    '
End Sub
'

Sub MENSADEMO()
    '
    Static CTX%
    If CTX% = 0 Then
        TXT$ = "La Presente Versión FLEXOFT es un\"
        TXT$ = TXT$ + "Ejemplar Limitado  para el Fin de\"
        TXT$ = TXT$ + "una Demostración.\  \"
        TXT$ = TXT$ + "Tiene Restricciones de Uso ya que\"
        TXT$ = TXT$ + "sólo Puede Utilizarse para Evaluar\"
        TXT$ = TXT$ + "su Licenciamiento y Adquisición Definitivos."
        Call COMUNI(TXT$)
        CTX% = 1
    End If
    '
End Sub
'

Function PROCACTI%()
   '
   TERMINAR.File1.Path = "C:\FLEXOFT"
   TERMINAR.File1.Pattern = "*.EXK"
   TERMINAR.File1.Refresh
   For U& = 1 To TERMINAR.File1.ListCount
      TERMINAR.File1.ListIndex = U& - 1
      X$ = TRIM$(UCase$(TERMINAR.File1.filename))
      If X$ <> "" Then
        If X$ <> "FLEXOFT.EXK" Then
          If FileLen("C:\FLEXOFT\" + X$) > 0 Then
            On Error GoTo 10
            Kill "C:\FLEXOFT\" + X$
            On Error GoTo 0
          End If
        End If
      End If
   Next U&
   PROCACTI% = 0
   Exit Function
   '
10 Resume 11
11 On Error GoTo 0
   PROCACTI% = 1
   X1$ = Left$(X$, Len(X$) - 1) + "E"
   TTXX$ = "Imposible Cerrar Normalmente la Sesión de Trabajo.\   \"
   TTXX$ = TTXX$ + "Se Encuentra Abierto y Ejecutando el Proceso\"
   TTXX$ = TTXX$ + "'" + X1$ + "'\  \"
   TTXX$ = TTXX$ + "Cierre este Proceso Antes de Terminar FLEXOFT."
   Call COMUNI(TTXX$)
   TERMINAR.Option1.Value = True
   '
End Function
