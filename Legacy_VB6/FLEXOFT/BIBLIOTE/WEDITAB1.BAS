Attribute VB_Name = "WEDITAB1"
' Option Explicit
'
Public TIPRO%
Public ARCHISAL$, ASAL%
Dim SUBTI$(16)
Dim XX As String * 128

Public PRINTERSET%, LARGOJA, ANCHOJA, ALTUREN, TALEBASI, TAFUESEL
Public TOTALI%, POSITOT%, RESUB$, RETRA$, ULTIHOJ%
Public HAYTOTA%
Public CXC$
Public NUHOJ%
Public DATLISTOK%, SOBRESCRI%
Public SRY As String * 255
Public GU As String * 255
Public LII$(10)
Public CCC$(10)
'
Dim RGFJ%(64), TFF%(64)
Dim VFMIN$(64), VFMAX$(64)
'
Dim VDS$(64, 16), VDX#(64, 16), VD%(64, 16), VFS$(64), VFSA$(64), VFX#(64), VF%(64), FI$(64)
Dim TVAL%(20), CHE$(20), BAN$(20), IMPCHE#(20), FECHE%(20)
Dim VDP%(64), VDPX#(64), VDPS$(64), VDA$(64), VDAA%(64), VDAB#(64)
Dim SUMA#(64, 2), SUPAR#(64)
Dim LEYE$(16)
'
Dim CALIMI As Integer
Dim CALIMS As Integer
Dim COEFH As Single
'
Dim MIZQ$
'
Public YALIS%(64)
Public YAPRE%
'
Public COLUAC%
Public MODIFI%
'
Dim PPZZ As String * 128
Dim FUFIJA$(128)
'
Public ANCHOPAP, LARGOPAP, MARGENI, MARGENS
Public ORIENTA%, CACOPIAS%
'
Public SALEXCEL%
'

Sub Main()
    '
    Screen.MousePointer = 11
    '
'MsgBox "001"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
'MsgBox "002"
    LU$ = LUDAT$
    LCB$ = LUCOM$
    HOO% = HOY(HOS$)
    '
10  Call SELPROCE                   ' seleccion proceso
'MsgBox "003"
    If TRIM$(Sali$) <> "" Then
        If Left$(Sali$, 1) = "*" Then
            Call FINAL(Sali$)
        End If
    End If
    '
    If COMMA$ = "" Then GoTo 10     ' entro por qweries
    '
    Screen.MousePointer = 11
    LBB$ = LUBAS$
    If EXISTE%(LUBAS$ + "PROTABU.EXE") > 0 Then
           Shell LUBAS$ + "PROTABU " + PROPRIN$, vbNormalFocus
        Else
           'Call COMUNI("Falta Utilitario 'PROTABU.EXE'\en Carpeta '" + REPLACAR$(LBB$, "\", "/") + "'.\   \Consulte a su Soporte de Sistemas.")
    End If
    '
    Screen.MousePointer = 1
    Call FINAL("")
    '
End Sub
'
Sub SELPROCE()
    '
    COMMA$ = TRIM$(Command$)
    PPX% = InStr(COMMA$, "#" + JOBID$ + "#")
    If PPX% > 0 Then
        COMMA$ = Left$(COMMA$, PPX% - 1) + Mid$(COMMA$, PPX% + 10)
    End If
    '
    If Left$(COMMA$, 1) = "*" Then
        COMMA$ = Mid$(COMMA$, 2)
    End If
    '
    
    PROPRIN$ = ""
    ATRIBU$ = ""
    PPX% = InStr(COMMA$, "/")
    If PPX% > 0 Then
          PROPRIN$ = TRIM$(Left$(COMMA$, PPX% - 1))
          ATRIBU$ = TRIM$(Mid$(COMMA$, PPX% + 1))
        Else
          PROPRIN$ = TRIM$(COMMA$)
    End If
    '
'PROPRIN$ = "OPERFAB"
    If PROPRIN$ = "" Then
        Screen.MousePointer = 1
        INITABU.Show 1
        If PROPRIN$ = "" Then
            Call FINAL("")
        End If
    End If
    '
    If ATRA$ = "" Then ATRA$ = CurDir
    FORMA$ = ATRA$ + "\" + PROPRIN$ + ".PRF"
    If EXISTE%(FORMA$) < 1 Then
        Call EXTRACT(FORMA$)
    End If
    '
    NX% = FILEOPEN%(FORMA$, 9, 128)
    If LOF(NX%) < 128 Then
        On Error Resume Next
        XX$ = String$(128, 0)
        Put #NX%, 1, XX$
        Close #NX%: NX% = 0
        Kill FORMA$
        On Error GoTo 0
        FORMB$ = REPLACAR$(FORMA$, "\", "/")
        Call COMUNI("Proceso Interrumpido.\  \No se ha Podido Hallar o Extraer\el Controlador Correspondiente\al Proceso Elegido:\  \" + FORMB$ + "\ \Consulte a su Soporte de Sistemas.")
        Call FINAL("")
    End If
    '
    Get #NX%, 1, XX$
    TIPRO% = CVI(Mid$(XX$, 127, 2))
    Close #NX%: NX% = 0
    '
    Screen.MousePointer = 1
    '
    Select Case TIPRO%
        '
        Case 1              ' SELECCIONA EDICION TABULADA
           Call CARDATFORM
           Call DIBUFORM
           YAFIL% = 0
           WEDITAB.Show 1
           Exit Sub
        Case 2              ' SELECCIONA LISTADO TABULADO
           WLISTAB.Show 1
           WLISTAB.Refresh
           Exit Sub
        Case 3              ' SELECCIONA FILTRO TABULADO
           'CCI$(0) = PROPRIN$
           'Call CONECRUN("#FILTABU\" + PROPRIN$, "")
           'Exit Sub
 '              WFILTAB.Show
 '              WFILTAB.Refresh
               Exit Sub
        Case 4              ' SELECCIONA ORDENAMIENTO
           Call CARDATSOR
           WSORTAB.Show 1
           Exit Sub
        Case 5              ' SELECCIONA PANTALLA PLANA
 '              WPLATAB.Show
 '              WPLATAB.Refresh
               Exit Sub
    End Select
    '
    Call MENSERR(24, "No se Encuentra Controlador\Correspondiente a " + PROPRIN$ + ".PRF")
    Call FINAL("")
    '
End Sub
'
        
Sub CARLIPRO()
    '
    Screen.MousePointer = 11
    If ATRA$ = "" Then ATRA$ = CurDir
    INITABU.LIPROCE.Clear
    INITABU.File1.Path = ATRA$
    '
    Dim XX As String * 128
    For KCL% = 1 To INITABU.File1.ListCount
        INITABU.File1.ListIndex = KCL% - 1
        X$ = UCase$(INITABU.File1.filename)
        '
        PX% = InStr(X$, ".PRF")
        If PX% > 1 Then
            APLX$ = TRIM$(UCase$(Left$(X$, PX% - 1)))
            If APLX$ <> "" Then
                NX% = FILEOPEN%(ATRA$ + "\" + APLX$ + ".PRF", 9, 128)
                Get #NX%, 1, XX$
                If CVI(Mid$(XX$, 127, 2)) = TIPRO% Then ' SELECCIONA TIPO DE CONTROLADOR
                    INITABU.LIPROCE.AddItem AJUSTI$(APLX$, 10) + Left$(XX$, 47)
                End If
                Close #NX%: NX% = 0
            End If
        End If
    Next KCL%
    '
    Screen.MousePointer = 1
    '
    CAPRO% = INITABU.LIPROCE.ListCount
    If CAPRO% < 1 Then
          Call MENSERR(24, "No se Encuentran Controladores\para el Tipo de Proceso Elegido.")
          Exit Sub
       ElseIf CAPRO% = 1 Then
          INITABU.LIPROCE.ListIndex = 0
          PROPRIN$ = TRIM$(Left$(INITABU.LIPROCE.Text, 8))
          ATRIBU$ = ""
          INITABU.Hide
          Exit Sub
       Else
          INITABU.LIPROCE.Top = 1300
          INITABU.LIPROCE.Left = 200
          INITABU.LIPROCE.Width = 6250
          INITABU.LIPROCE.Height = 2800
          INITABU.LIPROCE.Visible = True
          INITABU.LIPROCE.SetFocus
    End If
    '
End Sub
'

Sub CARDATFORM()
     '
     TIFORMED% = 2
     For II% = 0 To 31
       WEDITAB.FILABEL(II%).Visible = False
       WEDITAB.FIENTRA(II%).Visible = False
       WEDITAB.FIECO(II%).Visible = False
       WEDITAB.FIBOTSE(II%).Visible = False
       WEDITAB.LINVER(II%).Visible = False
       If II% > 0 Then
         WEDITAB.FILABEL(II%).AutoSize = WEDITAB.FILABEL(0).AutoSize
         WEDITAB.FILABEL(II%).Alignment = WEDITAB.FILABEL(0).Alignment
         WEDITAB.FILABEL(II%).Font = WEDITAB.FILABEL(0).Font
         WEDITAB.FILABEL(II%).FontName = WEDITAB.FILABEL(0).FontName
         WEDITAB.FILABEL(II%).FontSize = WEDITAB.FILABEL(0).FontSize
         WEDITAB.FILABEL(II%).FontBold = WEDITAB.FILABEL(0).FontBold
       End If
     Next II%
     '
5800 GRABACT% = 0
     Dim YYY As String * 128
     If N15% > 0 Then Close #N15%
     N15% = FILEOPEN%(FORMA$, 9, 128)
     Get #N15%, 1, YYY$
     '
     If LOF(N15%) <= 128 Then
         On Error Resume Next
         YYY$ = String$(128, 0)
         Put #N15%, 1, YYY$
         Close #N15%
         Kill FORMA$
         On Error GoTo 0
         FORMB$ = UCase$(REPLACAR$(FORMA$, "\", "/"))
         Call MENSERR(24, "Proceso Interrumpido.\Falta Controlador de Procesos '" + FORMB$ + "'.")
         Close: End
     End If
     '
     Archi$ = Mid$(YYY$, 57, 8)
     ARCHITRA$ = TRIM$(Archi$)
     ARCHILOC$ = "!" + ARCHITRA$
     PROCE$ = TRIM$(Mid$(YYY$, 65, 8))
     LOR1% = Asc(Mid$(YYY$, 106, 1)): If LOR1% > 16 Then LOR1% = 0
     LOR% = 256 * LOR1% + Asc(Mid$(YYY$, 107))
     CAMAX& = CVI(Mid$(YYY$, 108))
     PEOF% = 1 + Asc(Mid$(YYY$, 110))
     EOFS$ = Mid$(YYY$, 111, 16)
5803 S1$ = Left$(YYY$, 48)
     Sali$ = Mid$(YYY$, 49, 8)
     COD$ = Mid$(YYY$, 81)
     RGI& = Asc(Mid$(COD$, 1))
     RGF& = Asc(Mid$(COD$, 2))
     OFS% = Asc(Mid$(COD$, 3))
5804 BLP% = Asc(Mid$(COD$, 4))
     VIV% = Asc(Mid$(COD$, 5))
     APQ% = Asc(Mid$(COD$, 6))
        WEDITAB.BOTAGRE.Enabled = True
        If APQ% < 1 Then
            WEDITAB.BOTAGRE.Enabled = False
        End If
     IDL% = Asc(Mid$(COD$, 7))
        WEDITAB.BOTINSE.Enabled = True
        WEDITAB.BOTERASE.Enabled = True
        If IDL% < 1 Then
            WEDITAB.BOTINSE.Enabled = False
            WEDITAB.BOTERASE.Enabled = False
        End If
     WEDITAB.BOTEDIT.Enabled = False
     HDR& = Asc(Mid$(COD$, 8))
     RFT% = Asc(Mid$(COD$, 9))
     If IDL% > 0 Or APQ% > 0 Then GRABACT% = 1
     PRX% = Asc(Mid$(COD$, 10))
     NCL% = 0
     NFT% = 0
     NCLF% = 0
     TCL% = 0
     HAYTOTA% = 0
     RG% = 1
     '
5805 RG% = RG% + 1
     If RG% > LOF(N15%) / 128 Or RG% > 33 Then GoTo 5820
5806 Get #N15%, RG%, YYY$
     If YYY$ = String$(128, 0) Then GoTo 5820
     '
5807 TCL% = TCL% + 1
     TVL%(TCL%) = Asc(Mid$(YYY$, 33))
     LDL%(TCL%) = Asc(Mid$(YYY$, 34))
         If RG% = 2 Then TVL%(TCL%) = 3: LDL%(TCL%) = 0
5808 PR%(TCL%) = Asc(Mid$(YYY$, 35))
     LF%(TCL%) = Asc(Mid$(YYY$, 36))
         If RG% = 2 Then
           PR%(TCL%) = 0
           LF%(TCL%) = 2
         End If
5809 TT%(TCL%) = Asc(Mid$(YYY$, 88))
     If TT%(TCL%) > 0 Then
        HAYTOTA% = 1
        If TT%(TCL%) <> TCL% Then
            RTT%(TT%(TCL%) Mod 100) = 1
        End If
     End If
     '
5811 If Asc(Mid$(YYY$, 69)) < 1 Then GoTo 5816 ' SUPRIME COLUMNAS LONGITUD 0
     '
5812 NCL% = NCL% + 1
     CLT%(TCL%) = NCL%
     CFT%(NCL%) = RG% - 1
     CL0%(NCL%) = RG% - 1
     ECL$(NCL%) = TRIM$(Mid$(YYY$, 37, 32))
     If Left$(ECL$(NCL%), 1) = "@" Then
          NCLF% = NCLF% + 1
     End If
5813 ACL%(NCL%) = Asc(Mid$(YYY$, 69))
     LK%(NCL%) = Asc(Mid$(YYY$, 70))
5814 MG%(NCL%) = Asc(Mid$(YYY$, 71))
     PE%(NCL%) = Asc(Mid$(YYY$, 73))
          If RG% = 2 Then PE%(NCL%) = 2
     If PE%(NCL%) = 1 Then
        GRABACT% = 1
        WEDITAB.BOTEDIT.Enabled = True
     End If
5815 JT%(NCL%) = Asc(Mid$(YYY$, 72))
     AE$(NCL%) = Mid$(YYY$, 74, 8)
     CE%(NCL%) = Asc(Mid$(YYY$, 82))
5816 GoTo 5805
5817 '
5820 ATOT% = NCL% - 1 - NCLF%
     CLII% = OFS% + 2
     For KCL% = 1 To NCL%
        LC%(KCL%) = ACL%(KCL%)
        If Left$(ECL$(KCL%), 1) <> "@" Then
            LC%(KCL%) = ACL%(KCL%)
            ATOT% = ATOT% + ACL%(KCL%)
            CL%(KCL%) = CLII%
            CLII% = CLII% + ACL%(KCL%) + 1
        End If
     Next KCL%
     '
     If ATOT% > 85 Then
        FORMAA$ = REPLACAR$(FORMA$, "\", "/")
        MENSA$ = "Error: Imposible Presentar Formulario Tabulado.\Ancho de Pantalla Tabulada " + FORMAA$ + " =" + Str$(ATOT% + 2) + " Pos."
        If USNBR% < 100 Then
          MENSA$ = MENSA$ + "\  \Informe al Supervisor del Sistema."
        End If
        Call MENSERR(24, MENSA$)
        Call CIERRARCH("*.*")
        If USNBR% > 100 Then
          If EXISTE%("PARAMED.EXE") > 0 Then
            Call CONECRUN("#PARAMED/" + PROPRIN$, "Definición Formulario Consulta Tabulada")
          End If
        End If
        Call FINAL("")
     End If
     '
5821 For KCL% = 1 To NCL%
        If CE%(KCL%) <= 0 Then GoTo 5824
5822    For KCL1% = 1 To NCL%
           If CL0%(KCL1%) = CE%(KCL%) Then
               CE%(KCL%) = KCL1%: PE%(KCL1%) = 3
               GoTo 5824
           End If
5823    Next KCL1%
5824 Next KCL%
     '
5825 If BLP% > 0 Then
'         If OFS% = 0 Then
'             OFS% = (78 - ATOT%) / 2
'             For KCL% = 1 To NCL%
'                 CL%(KCL%) = CL%(KCL%) + OFS%
'             Next KCL%
'         End If
     End If
     '
5830 NFT% = 0
     RG% = 1
     RG0% = RGI&
     RGFI% = 25: RGFS% = 1
     '
5835 RG% = RG% + 1
     If RG% > LOF(N15%) / 128 Or RG% > 33 Then GoTo 5850
5840 Get #N15%, RG%, YYY$
     If YYY$ = String$(128, 0) Then GoTo 5850
5841 If Asc(Mid$(YYY$, 83)) < 1 Then GoTo 5847 ' SUPRIME RG-FILTRO=0
5842 NFT% = NFT% + 1
     CFF%(NFT%) = RG% - 1
     LCF%(NFT%) = Asc(Mid$(YYY$, 83))
     RGFL&(NFT%) = Asc(Mid$(YYY$, 84))
     CLF%(NFT%) = Asc(Mid$(YYY$, 85))
5843 If RGFL&(NFT%) - 1 < RG0% Then
         RG0% = RGFL&(NFT%) - 1
     End If
     If RGFL&(NFT%) < RGFI% Then RGFI% = RGFL&(NFT%)
     If RGFL&(NFT%) > RGFS% Then RGFS% = RGFL&(NFT%)
     '
5844 ECLF$(NFT%) = RTrim$(Mid$(YYY$, 37, 32))
     If Len(ECLF$(NFT%)) > 0 Then
         ECLF$(NFT%) = ECLF$(NFT%) + ":"
     End If
5845 LKF%(NFT%) = Asc(Mid$(YYY$, 70))
     LEF%(NFT%) = Asc(Mid$(YYY$, 86))
     CEF%(NFT%) = Asc(Mid$(YYY$, 87))
5846 AEF$(NFT%) = Mid$(YYY$, 74, 8)
5847 GoTo 5835
 '
5850 Close #N15%: N15% = 0
     '
'     Call CLEARFORM
     '
End Sub
'

Sub DIBUFORM()     ' Rutina de dibujo pantalla tabulada
'
HREN% = WEDITAB.ScaleHeight / 25
WEDITAB.Caption = TRIM$(S1$)
'
EPAC$ = TRIM$(EMPREAC$)
If EPAC$ <> "" Then
    WEDITAB.Caption = WEDITAB.Caption + "  -  " + EPAC$
End If
'
MAIZQ = 0
MAINT = 170
MADER = 220
'
ENCA$ = " "
'
CALINVE% = 0
For KCL% = 1 To NCL%
      '
      If Left$(ECL$(KCL%), 1) <> "@" Then
          If Left$(ECL$(KCL%), 1) = "<" Then
              ENCACO$ = AJUSTI$(Mid$(ECL$(KCL%), 2), LC%(KCL%))
              ENCACO$ = Space$(MG%(KCL%)) + ENCACO$
              ENCACO$ = Left$(ENCACO$, LC%(KCL%))
            ElseIf Left$(ECL$(KCL%), 1) = ">" Then
              ENCACO$ = AJUSTD$(Mid$(ECL$(KCL%), 2), LC%(KCL%))
              ENCACO$ = Mid$(ENCACO$, MG%(KCL%) + 1) + Space$(MG%(KCL%))
            Else
              ENCACO$ = TRIM$(ECL$(KCL%))
              While Len(ENCACO$) < LC%(KCL%)
                 ENCACO$ = " " + ENCACO$ + " "
              Wend
              ENCACO$ = Left$(ENCACO$, LC%(KCL%))
          End If
          ENCA$ = ENCA$ + ENCACO$
          ENCA$ = ENCA$ + " "
          WEDITAB.TEPRUEBA.Caption = ENCA$
          CALINVE% = CALINVE% + 1
          WEDITAB.LINVER(CALINVE% - 1).Left = WEDITAB.TEPRUEBA.Width
        Else
          ' CL%(KCL%) = 4 + Len(ECL$(KCL%))
      End If
      '
Next KCL%
'
ENCA$ = ENCA$ + "   "
WEDITAB.TEPRUEBA.Caption = ENCA$
WEDITAB.LITAB.Width = WEDITAB.TEPRUEBA.Width - 50
WEDITAB.FRAMENCA.Width = WEDITAB.TEPRUEBA.Width - 50
WEDITAB.ENCACOL.Caption = ENCA$
WEDITAB.Picture1.Width = WEDITAB.TEPRUEBA.Width - 50
WEDITAB.ENCACOL.Width = WEDITAB.TEPRUEBA.Width - 50
WEDITAB.FRAMETOT.Width = WEDITAB.TEPRUEBA.Width - 50
WEDITAB.TOTALI.Width = WEDITAB.TEPRUEBA.Width - 50
WEDITAB.LITAB.Left = MAIZQ
WEDITAB.FRAMENCA.Left = MAIZQ
WEDITAB.ENCACOL.Left = 50
WEDITAB.FRAMETOT.Left = MAIZQ
WEDITAB.TOTALI.Left = 50
WEDITAB.Width = MAIZQ + WEDITAB.LITAB.Width + MAINT + WEDITAB.FRAMEBOT1.Width + MADER
COEFH = WEDITAB.LITAB.Width / Len(ENCA$)
'
For II% = 1 To NFT%
   '
   If TRIM$(ECLF$(II%)) <> "" Then
      CLL% = CLF%(II%) - 2
      WEDITAB.FILABEL(II% - 1).Caption = ECLF$(II%) + " "
      WEDITAB.FILABEL(II% - 1).Left = MAIZQ + COEFH * (CLF%(II%) - Len(ECLF$(II%)) - 3)
      WEDITAB.FILABEL(II% - 1).Top = 150 + (0.5 + (RGFL&(II%) - RGFI%) * 1.5) * WEDITAB.FIENTRA(II% - 1).Height
      WEDITAB.FILABEL(II% - 1).Width = COEFH * (2 + Len(ECLF$(II%)))
      WEDITAB.FILABEL(II% - 1).Visible = True
   End If
   '
   WEDITAB.FIENTRA(II% - 1).Left = COEFH * (CLF%(II%) - 1)
   WEDITAB.FIENTRA(II% - 1).Top = 70 + (0.5 + (RGFL&(II%) - RGFI%) * 1.5) * WEDITAB.FIENTRA(II% - 1).Height
   '
   LCFF% = LCF%(II%)
   WEDITAB.TEPRU.Font = WEDITAB.FIENTRA(II% - 1).Font
   WEDITAB.TEPRU.FontSize = WEDITAB.FIENTRA(II% - 1).FontSize
   WEDITAB.TEPRU.FontBold = WEDITAB.FIENTRA(II% - 1).FontBold
   WEDITAB.TEPRU.Caption = String$(LCFF%, "M") + "."
   WEDITAB.FIENTRA(II% - 1).Width = WEDITAB.TEPRU.Width
   If LCFF% >= 6 Then
        WEDITAB.FIENTRA(II% - 1).Width = COEFH * (1 + LCFF%)
   End If
   '
   WEDITAB.FIENTRA(II% - 1).Visible = True
   '
   If TRIM$(AEF$(II%)) <> "" Then
      If LEF%(II%) > 0 Then
         WEDITAB.FIECO(II% - 1).Left = COEFH * (CEF%(II%) - 1)
         WEDITAB.FIECO(II% - 1).Top = 70 + (0.5 + (RGFL&(II%) - RGFI%) * 1.5) * WEDITAB.FIENTRA(II% - 1).Height
         WEDITAB.FIECO(II% - 1).Width = COEFH * (LEF%(II%) + 1)
         WEDITAB.FIECO(II% - 1).Visible = True
      End If
   End If
   '
   If TRIM$(AEF$(II%)) <> "" Or TVL%(CFF%(II%)) = 5 Then
      If Left$(UCase$(AEF$(II%)), 4) <> "AUTO" Then
         WEDITAB.FIBOTSE(II% - 1).Top = WEDITAB.FIENTRA(II% - 1).Top
         WEDITAB.FIBOTSE(II% - 1).Height = WEDITAB.FIENTRA(II% - 1).Height
         WEDITAB.FIBOTSE(II% - 1).Width = WEDITAB.FIBOTSE(II% - 1).Height * 0.6
         WEDITAB.FIBOTSE(II% - 1).Left = WEDITAB.FIENTRA(II% - 1).Left + WEDITAB.FIENTRA(II% - 1).Width
         WEDITAB.FIBOTSE(II% - 1).Visible = True
      End If
   End If
   '          '
Next II%
'
If NFT% > 0 Then
     WEDITAB.Picture1.Visible = True
     WEDITAB.FRAMENCA.Top = (2.5 + (RGFS% - RGFI%) * 1.5) * WEDITAB.FIENTRA(0).Height
   Else
     WEDITAB.Picture1.Visible = False
     WEDITAB.FRAMENCA.Top = 0.5 * WEDITAB.FIENTRA(0).Height
End If
WEDITAB.FRAMENCA.Height = 2 * WEDITAB.ENCACOL.Height
WEDITAB.ENCACOL.Top = (WEDITAB.FRAMENCA.Height - WEDITAB.ENCACOL.Height) / 2
WEDITAB.ENCACOL.Caption = ENCA$
WEDITAB.LITAB.Top = WEDITAB.FRAMENCA.Top + WEDITAB.FRAMENCA.Height
'
WEDITAB.FRAMETOT.Height = 2.2 * WEDITAB.TOTALI.Height
WEDITAB.LITAB.Height = WEDITAB.Height - 200 - WEDITAB.LITAB.Top - 0.2 * WEDITAB.FIENTRA(0).Height - WEDITAB.FRAMETOT.Height
WEDITAB.FRAMETOT.Top = WEDITAB.LITAB.Top + WEDITAB.LITAB.Height
WEDITAB.TOTALI.Top = (WEDITAB.FRAMETOT.Height - WEDITAB.TOTALI.Height) / 2
WEDITAB.FRAMETOT.Visible = True
'
WEDITAB.Picture1.Height = WEDITAB.FRAMETOT.Top + WEDITAB.FRAMETOT.Height
'
WEDITAB.Image3.Top = WEDITAB.FRAMETOT.Top '- 60
WEDITAB.Image3.Left = WEDITAB.FRAMETOT.Width - 370
'
For KKI% = 1 To CALINVE%
  WEDITAB.LINVER(KKI% - 1).Top = WEDITAB.FRAMENCA.Top
  WEDITAB.LINVER(KKI% - 1).Height = WEDITAB.FRAMETOT.Top - WEDITAB.FRAMENCA.Top + WEDITAB.FRAMETOT.Height
  WEDITAB.LINVER(KKI% - 1).Visible = True
Next KKI%
'
WEDITAB.Line1.X1 = WEDITAB.FRAMETOT.Left
WEDITAB.Line1.X2 = 12000 ' WEDITAB.FRAMETOT.Left + WEDITAB.FRAMETOT.Width
WEDITAB.Line1.Y1 = WEDITAB.FRAMETOT.Top + WEDITAB.FRAMETOT.Height
WEDITAB.Line1.Y2 = WEDITAB.FRAMETOT.Top + WEDITAB.FRAMETOT.Height
'
WEDITAB.Line2.X1 = WEDITAB.FRAMENCA.Left + WEDITAB.FRAMENCA.Width
WEDITAB.Line2.X2 = WEDITAB.FRAMENCA.Left + WEDITAB.FRAMENCA.Width
WEDITAB.Line2.Y1 = 0  ' WEDITAB.FRAMENCA.Top
WEDITAB.Line2.Y2 = WEDITAB.FRAMETOT.Top + WEDITAB.FRAMETOT.Height
'
WEDITAB.FRAMEBOT1.Top = WEDITAB.FIENTRA(0).Height * 0.5
WEDITAB.FRAMEBOT1.Left = MAIZQ + WEDITAB.LITAB.Width + MAINT
WEDITAB.FRAMEBOT2.Top = WEDITAB.FRAMETOT.Top - 2560
WEDITAB.FRAMEBOT2.Left = MAIZQ + WEDITAB.LITAB.Width + MAINT
'
WEDITAB.MARCOBARRA.Top = WEDITAB.FRAMEBOT2.Top - 300
WEDITAB.MARCOBARRA.Left = WEDITAB.FRAMEBOT2.Left
WEDITAB.MARCOBARRA.Width = WEDITAB.FRAMEBOT2.Width
'
WEDITAB.LITAB.Clear
While WEDITAB.LITAB.TopIndex <= 0
   WEDITAB.LITAB.AddItem "   "
   'Call VERISCROLL
   WEDITAB.LITAB.ListIndex = WEDITAB.LITAB.ListCount - 1
Wend
CARELI% = WEDITAB.LITAB.ListCount - 1
' AQUI ESTA BIEN
WEDITAB.LITAB.Clear
'Call VERISCROLL
'
WEDITAB.VScroll1.Top = WEDITAB.LITAB.Top
WEDITAB.VScroll1.Height = WEDITAB.LITAB.Height
WEDITAB.VScroll1.Left = WEDITAB.LITAB.Left + WEDITAB.LITAB.Width - WEDITAB.VScroll1.Width - 20
WEDITAB.VScroll1.Visible = True
'
WEDITAB.Refresh
'
End Sub

Sub CLEARFORM()
    '
    YAFIL% = 0
    WEDITAB.LITAB.Clear
    Call VERISCROLL
    For FT% = 0 To 31
        WEDITAB.FIENTRA(FT%).Text = ""
        WEDITAB.FIENTRA(FT%).Enabled = True
    Next FT%
    '
    For K2% = 1 To 64
       TTX#(K2%) = 0
       TXX#(K2%) = 0
    Next K2%
    '
End Sub
'
Sub ENTRAFIL()
    '
    On Error Resume Next
    WEDITAB.SetFocus
    On Error GoTo 0
    If YAFIL% < 1 Then
       WEDITAB.LITAB.Clear
       Call VERISCROLL
       For K2% = 1 To 64
          TTX#(K2%) = 0
          TXX#(K2%) = 0
       Next K2%
       '
       For FT% = 1 To NFT%
          WEDITAB.FIENTRA(FT% - 1).Visible = True
          WEDITAB.FIENTRA(FT% - 1).Enabled = True
          ARCHECO$ = UCase$(TRIM$(AEF$(FT%)))
          If ARCHECO$ = "AUTO" Then
             If Val(WEDITAB.FIENTRA(FT% - 1).Text) < 1 Then
                REG& = ULTREG&(ARCHITRA$)
                TF1% = TIPFIELD%(ARCHITRA$, CFF%(FT%) - 1)
                LF1% = LENFIELD%(ARCHITRA$, CFF%(FT%) - 1)
                PF1% = POSFIELD%(ARCHITRA$, CFF%(FT%) - 1)
                XXZ$ = Mid$(REGLEIDO$(ARCHITRA$, REG&), PF1%, LF1%)
                NRO = 1 + Val(CSTRING$(XXZ$, TF1%, 12, 0, 1))
                WEDITAB.FIENTRA(FT% - 1).Text = TRIM$(Str$(NRO))
             End If
          End If
       Next FT%
       '
       For FT% = 1 To NFT%
          If FT% = 1 Then
            If ATRIBU$ <> "" Then
               WEDITAB.FIENTRA(0).Text = ATRIBU$
               ATRIBU$ = ""
            End If
          End If
          '
          If LKF%(FT%) > 0 Then
             If TRIM$(WEDITAB.FIENTRA(FT% - 1).Text) = "" Then
                 WEDITAB.FIENTRA(FT% - 1).Enabled = True
                 WEDITAB.FIENTRA(FT% - 1).SetFocus
                 Exit Sub
             End If
             '
             If TRIM$(WEDITAB.FIENTRA(FT% - 1).Text) = "." Then
                 Exit Sub
             End If
             '
           End If
       Next FT%
       '
       Call CARLITAB(FINOVA%)
       If FINOVA% > 0 Then
          'YAFIL% = 1
          Call CLEARFORM
          Exit Sub
       End If
       '
       If WEDITAB.LITAB.ListCount > 0 Then
           YAFIL% = 2
           WEDITAB.LITAB.ListIndex = 0
           '
           If TRIM$(UCase$(PROPRIN$)) = "KARDEX" Then
              If WEDITAB.LITAB.ListCount > 20 Then
                 WEDITAB.LITAB.TopIndex = WEDITAB.LITAB.ListCount - 20
              End If
              WEDITAB.LITAB.ListIndex = WEDITAB.LITAB.ListCount - 1
           End If
           '
           Exit Sub
         Else
           '
           If APQ% < 1 Then
               If NFT% < 1 Then
                    YAFIL% = (-1)
                    Call MENSERR(24, "Archivo de Trabajo Vacío")
                    'MsgBox "Archivo de Trabajo Vacío"
                    Exit Sub
                  Else
                    Call MENSERR(24, "Ningun Elemento Cumple\Condición de Filtro")
                    'MsgBox "Ningun Elemento Cumple Condición de Filtro"
                    If NFT% = 1 Then
                       If LKF%(1) < 1 Then
                          Unload WEDITAB
                       End If
                    End If
                    
               End If
               Call CLEARFORM
               Exit Sub
           End If
           '
           For FT% = 1 To NFT%
              If TRIM$(WEDITAB.FIENTRA(FT% - 1).Text) = "" Then
                 WEDITAB.FIENTRA(FT% - 1).Enabled = True
                 WEDITAB.FIENTRA(FT% - 1).SetFocus
                 Exit Sub
              End If
           Next FT%
           YAFIL% = 1
           If APQ% > 0 Then
               Call ADDLISTA
               Exit Sub
             Else
               If NFT% < 1 Then
                    YAFIL% = (-1)
                    Call MENSERR(24, "Archivo de Trabajo Vacío")
                    'MsgBox "Archivo de Trabajo Vacío"
                    Exit Sub
                  Else
                    Call MENSERR(24, "Ningun Elemento Cumple\Condición de Filtro")
                    'MsgBox "Ningun Elemento Cumple Condición de Filtro"
               End If
               Call CLEARFORM
               Exit Sub
           End If
       End If
       '
    End If
    '
    If YAFIL% < 2 Then
       For FT% = 1 To NFT%
          If LKF%(FT%) < 1 Then
              If TRIM$(WEDITAB.FIENTRA(FT% - 1).Text) = "" Then
                 WEDITAB.FIENTRA(FT% - 1).SetFocus
                 Exit Sub
              End If
          End If
       Next FT%
       YAFIL% = 2
    End If
    '
End Sub
'

Sub CARLITAB(FINOVA%)
    '
    Static N18%
    Static IXT As String * 1024
    '
    Screen.MousePointer = 11
    FINOVA% = 0
    '
    WEDITAB.SORTLIST.Clear
    WEDITAB.LITAB.Visible = False
    WEDITAB.LITAB.Clear
    Call VERISCROLL
    Call BLANARCH(ARCHILOC$)
    '
    For K2% = 1 To 64
       TTX#(K2%) = 0
       TXX#(K2%) = 0
       WEDITAB.TOTALI.Caption = ""
    Next K2%
    '
    If ULTREG&(ARCHITRA$) > 0 Then
       WEDITAB.MARCOBARRA.Top = WEDITAB.FRAMEBOT2.Top - 300
       WEDITAB.MARCOBARRA.Left = WEDITAB.FRAMEBOT2.Left
       WEDITAB.MARCOBARRA.Width = WEDITAB.FRAMEBOT2.Width
    End If
    '
    ' FINUII% = FILENBR%(ARCHITRA$)
    LOREII% = RECLEN%(ARCHITRA$)
    '
    If NFT% > 0 Then
       '
       CONTA& = 0
       IDX% = 0
       CARESEL& = 0
       ORDE% = 0: If TVL%(CFT%(1)) = 5 Then ORDE% = 1
       '
       FT1% = 0
       For FT% = 1 To NFT%
          FI$(FT%) = ""
          If LKF%(FT%) > 0 Then
             FT1% = FT%
             FI$(FT%) = CBIN$(WEDITAB.FIENTRA(FT% - 1).Text, TVL%(CFF%(FT%)), LF%(CFF%(FT%)))
             If TRIM$(AEF$(FT%)) <> "" Then
                If Left$(AEF$(FT%), 4) <> "AUTO" Then
                   If ECOARCH$(AEF$(FT%), FI$(FT%)) = "" Then
                      On Error GoTo 999
                      WEDITAB.FIENTRA(FT% - 1).SetFocus
                      On Error GoTo 0
                      Call MENSERR(24, "Filtro no Válido")
                      FINOVA% = 1
                      GoTo 999
                   End If
                End If
             End If
          End If
       Next FT%
       '
       If FT1% = 1 Then
          ' intenta seleccion indexada
          'COLUX$ = TRIM$(Str$(CFF%(1) - 1))
          'While Len(COLUX$) < 2: COLUX$ = "0" + COLUX$: Wend
          'ARIX$ = ARCHITRA$ + ".X" + COLUX$
          '
          'If N18% > 0 Then Close #N18%
          'N18% = FILEOPEN%(LU$ + ARIX$, 0, 1024)
          '
          'If LOF(N18%) < 1 Then
          '      IXT$ = String$(1024, 0)
          '      Put #N18%, 1, IXT$
          '      Close #N18%: N18% = 0
          '      Kill LU$ + ARIX$
          '   Else             ' leer en forma secuencial
                IDX% = 1
                ARGU$ = FI$(1)
                LARGU& = Len(ARGU$)
                RFKX$ = RECFILT$(ARCHITRA$, CFF%(1) - 1, ARGU$)
                For IX& = 1 To Len(RFKX$) Step 4
                  u& = CVS(Mid$(RFKX$, IX&, 4))
                  GoSub 701
                Next IX&
                '
                'For IX& = 1 To LOF(N18%) / 1024 + 1
                '   Get #N18%, IX&, IXT$
                '   PINI& = 1
650             '   POSI& = InStr(PINI&, IXT$, ARGU$)
                '   If POSI& >= PINI& Then
                '      If (POSI& Mod LARGU&) = 1 Then
                '            U& = (1024& * (IX& - 1) + (POSI& + LARGU& - 1)) / LARGU&
                '            GoSub 701                ' EXTRACTAR
                '            PINI& = POSI& + LARGU&
                '         Else
                '            PINI& = POSI& + 1
                '      End If
                '      GoTo 650
                '   End If
                'Next IX&
                GoTo 990
          'End If
       End If
    End If
    '
    TINI = Timer
    Screen.MousePointer = 11
    '
    FIN& = ULTREG&(ARCHITRA$)
    For u& = 1 To FIN&
        '
        If (Timer - TINI) > 1 Then
            WEDITAB.MARCOBARRA.Visible = True
        End If
        '
        Call TRACE(20, u&, FIN&)
701     'XXX$ = LEEREG$(FINUII%, LOREII%, U&)
        XXX$ = REGLEIDO$(ARCHITRA$, u&)
        RELIB$ = Space$(ATOT% + 2)
        '
        For FT% = 1 To NFT%
            If LKF%(FT%) > 0 Then
               If Mid$(XXX$, PR%(CFF%(FT%)), LF%(CFF%(FT%))) <> FI$(FT%) Then GoTo 979
            End If
        Next FT%
        '
        CARESEL& = CARESEL& + 1
        If CARESEL& = 1 Then
            For FT% = 1 To NFT%
               If LKF%(FT%) < 1 Then
                  XXZ$ = Mid$(XXX$, PR%(CFF%(FT%)), LF%(CFF%(FT%)))
                  TXT$ = CSTRING$(XXZ$, TVL%(CFF%(FT%)), LCF%(FT%), LDL%(CFF%(FT%)), 1)
                  WEDITAB.FIENTRA(FT% - 1).Text = TXT$
'                  WEDITAB.FIENTRA(FT% - 1).Enabled = False
               End If
            Next FT%
        
' SELECCIONAR Y PRESENTAR FILTROS
        End If
        '
        For K2% = 1 To NCL%
           If CFT%(K2%) <> 1 Then
             If RTT%(CFT%(K2%)) < 1 Then
               If LF%(CFT%(K2%)) > 0 Then
                   XXZ$ = Mid$(XXX$, PR%(CFT%(K2%)), LF%(CFT%(K2%)))
                   ACLX% = ACL%(K2%) - MG%(K2%)
                   TXT$ = CSTRING$(XXZ$, TVL%(CFT%(K2%)), ACLX%, LDL%(CFT%(K2%)), JT%(K2%))
                   '
                   If JT%(K2%) = 2 Then
                       TXT$ = TXT$ + Space$(MG%(K2%))
                     Else
                       TXT$ = Space$(MG%(K2%)) + TXT$
                   End If
                   If CL%(K2%) > 0 Then Mid$(RELIB$, CL%(K2%), LC%(K2%)) = TXT$
                   '
                   If ORDE% = 1 Then
                      If K2% = 1 Then
                         SORTAR$ = CSTRING$(XXZ$, 1, 6, 0, 2) + CSTRING$(MKS$(u&), 3, 9, 0, 2)
                      End If
                   End If
                   '
                   If TRIM$(AE$(K2%)) <> "" Then
                      XXE$ = ECOARCH$(AE$(K2%), XXZ$)
                      COLOG% = K2%
                      CLE% = CL%(CE%(COLOG%))
                      LCE% = LC%(CE%(COLOG%))
                      JTE% = JT%(CE%(COLOG%))
                      TXT$ = CSTRING$(XXE$, 6, LCE%, 0, JTE%)
                      If CLE% > 0 Then Mid$(RELIB$, CLE%, LCE%) = TXT$
                   End If
                   '
                   TTT% = TT%(CFT%(K2%))
                   If TTT% > 0 Then
                      If (TTT% Mod 100) <= TCL% Then
                         TTX#(K2%) = TTX#(K2%) + Val(TXT$)
                         If TTT% < 100 Then TXX#(TTT%) = TXX#(TTT%) + Val(TXT$) Else TXX#(TTT% Mod 100) = TXX#(TTT% Mod 100) - Val(TXT$)
                         If (TTT% Mod 100) <> CFT%(K2%) Then
                             COLOG% = CLT%(TTT% Mod 100)   ' CFT%(TTT% Mod 100)
                             CLE% = CL%(COLOG%)
                             LCE% = LC%(COLOG%)
                             JTE% = JT%(COLOG%)
                             TXT$ = CSTRING$(MKD$(TXX#(TTT% Mod 100)), TVL%(CFT%(COLOG%)), LCE%, LDL%(CFT%(COLOG%)), JTE%)
                             If CLE% > 0 Then Mid$(RELIB$, CLE%, LCE%) = TXT$
                         End If
                      End If
                   End If
                   '
                End If
             End If
           End If
        Next K2%
        '
        On Error GoTo 980
        If ORDE% = 1 Then
              WEDITAB.SORTLIST.AddItem SORTAR$ + RELIB$ + "   " + Str$(u&)
            Else
              WEDITAB.LITAB.AddItem RELIB$ + "   " + Str$(u&)
              'Call VERISCROLL
        End If
        On Error GoTo 0
        '
        If IDX% = 1 Then
            Return
        End If
        '
979 Next u&
    GoTo 990
    '
980 Resume 981
981 On Error GoTo 0
    Call MENSERR(24, "Imposible Presentar Consulta Completa.\La Base de Datos es Demasiado Grande.")
    '
990 WEDITAB.LITAB.Visible = False
    If ORDE% = 1 Then
        For K2% = 1 To 64: TXX#(K2%) = 0: Next K2%
        FIN& = WEDITAB.SORTLIST.ListCount
        For u& = 1 To WEDITAB.SORTLIST.ListCount
            Call TRACE(1, u&, FIN&)
            WEDITAB.SORTLIST.ListIndex = u& - 1
            RELIB$ = Mid$(WEDITAB.SORTLIST.Text, 16)
            '
            For K2% = 1 To NCL%
               TTT% = TT%(CFT%(K2%))
               If TTT% > 0 Then
                   If (TTT% Mod 100) <= TCL% Then
                      VALO# = Val(Mid$(RELIB$, CL%(K2%), LC%(K2%)))
                      If TTT% < 100 Then TXX#(TTT%) = TXX#(TTT%) + VALO# Else TXX#(TTT% Mod 100) = TXX#(TTT% Mod 100) - VALO
                      If (TTT% Mod 100) <> CFT%(K2%) Then
                          COLOG% = CLT%(TTT% Mod 100)   ' CFT%(TTT% Mod 100)
                          CLE% = CL%(COLOG%)
                          LCE% = LC%(COLOG%)
                          JTE% = JT%(COLOG%)
                          TXT$ = CSTRING$(MKD$(TXX#(TTT% Mod 100)), TVL%(CFT%(COLOG%)), LCE%, LDL%(CFT%(COLOG%)), JTE%)
                          If CLE% > 0 Then Mid$(RELIB$, CLE%, LCE%) = TXT$
                       End If
                   End If
               End If                   '
            Next K2%
            '
            WEDITAB.LITAB.AddItem RELIB$
            'Call VERISCROLL
        Next u&
    End If
    WEDITAB.SORTLIST.Clear
    '
    WEDITAB.LITAB.Visible = False
    WEDITAB.LIREGSEL.Clear
    FIN& = WEDITAB.LITAB.ListCount
    For u& = 1 To WEDITAB.LITAB.ListCount
        Call TRACE(1, u&, FIN&)
        WEDITAB.LITAB.ListIndex = u& - 1
        RELIB$ = Mid$(WEDITAB.LITAB.Text, ATOT% + 3)
        WEDITAB.LIREGSEL.AddItem RELIB$
        '
        REGPRI& = Val(RELIB$)
        XXX$ = REGLEIDO$(ARCHITRA$, REGPRI&)
        Call REGAPP(ARCHILOC$, XXX$)
        '
    Next u&
    WEDITAB.LITAB.Visible = True
    '
    RELIB$ = Space$(ATOT% + 2)
    For K2% = 1 To NCL%
       If TT%(CFT%(K2%)) > 0 Then
           TVLL% = 4: XXZ$ = MKD$(TTX#(K2%))
           ACO% = ACL%(K2%)
           If JT%(K2%) = 2 Then ACO% = ACO% - MG%(K2%)
           TXT$ = CSTRING$(XXZ$, TVLL%, ACO%, LDL%(CFT%(K2%)), JT%(K2%))
           If CL%(K2%) > 0 Then Mid$(RELIB$, CL%(K2%), LC%(K2%)) = TXT$
       End If
    Next K2%
    WEDITAB.TOTALI.Caption = RELIB$
    '
999 On Error GoTo 0
    WEDITAB.LITAB.Visible = True
    WEDITAB.MARCOBARRA.Visible = False
    Screen.MousePointer = 1
    WEDITAB.TOTALI.Visible = True
    Call VERISCROLL
    MODIFI% = 0
    '
End Sub
'

Sub GRABALI()
    '
    If WEDITAB.BOTAGRE.Enabled = True Then GoTo 5
    If WEDITAB.BOTINSE.Enabled = True Then GoTo 5
    If WEDITAB.BOTERASE.Enabled = True Then GoTo 5
    If WEDITAB.BOTEDIT.Enabled = True Then GoTo 5
    Exit Sub
    '
5   UU& = 0
    For u& = 1 To WEDITAB.LITAB.ListCount
        '
        XXX$ = REGLEIDO$(ARCHILOC$, u&)
        For FT% = 1 To NFT%
            XXZ$ = WEDITAB.FIENTRA(FT% - 1).Text
            XXY$ = CBIN$(XXZ$, TVL%(CFF%(FT%)), LF%(CFF%(FT%)))
            Call REPLA$(XXX$, XXY$, PR%(CFF%(FT%)), LF%(CFF%(FT%)))
        Next FT%
        '
        REGGRA& = 0
        UU& = UU& + 1
        If UU& <= WEDITAB.LIREGSEL.ListCount Then
            WEDITAB.LIREGSEL.ListIndex = UU& - 1
            REGGRA& = Val(WEDITAB.LIREGSEL.Text)
        End If
        '
        If REGGRA& < 1 Or REGGRA& > ULTREG&(ARCHITRA$) Then
              Call REGAPP(ARCHITRA$, XXX$)
              REGGRA& = ULTREG&(ARCHITRA$)
              WEDITAB.LIREGSEL.AddItem TRIM$(Str$(REGGRA&))
            Else
              Call GRAREG(ARCHITRA$, XXX$, REGGRA&)
        End If
        '
    Next u&
    '
    For u& = UU& + 1 To WEDITAB.LIREGSEL.ListCount
        WEDITAB.LIREGSEL.ListIndex = u& - 1
        REGGRA& = Val(WEDITAB.LIREGSEL.Text)
        If REGGRA& > 0 Then
           If REGGRA& <= ULTREG&(ARCHITRA$) Then
              LORE% = RECLEN%(ARCHITRA$)
              XXX$ = String$(LORE%, 0)
              Call GRAREG(ARCHITRA$, XXX$, REGGRA&)
           End If
        End If
    Next u&
    '
10  Call CIERRARCH(ARCHITRA$)
    '
End Sub
'
Sub SELECFIL(Index)
    '
    WEDITAB.FIENTRA(Index).Text = TRIM$(WEDITAB.FIENTRA(Index).Text)
    '
    K2% = Index + 1
    If TRIM$(AEF$(K2%)) = "" Then
        Call COMUNI("No Existe Tabla de Selección\para el Campo Elegido")
        WEDITAB.FIENTRA(Index).Enabled = True
        WEDITAB.FIENTRA(Index).SetFocus
        Exit Sub
    End If
    '
    ARCHECO$ = UCase$(TRIM$(AEF$(K2%)))
    If ARCHECO$ = "AUTO" Then
        WEDITAB.FIENTRA(Index).Enabled = True
        Exit Sub
    End If
    '
    If ARCHECO$ = "MASTER" Then ARCHECO$ = "ECOMAST"
    '
    If YALIS%(Index) = 0 Then
        Screen.MousePointer = 11
        TF1% = TIPFIELD%(ARCHECO$, 1)
        LF1% = LENFIELD%(ARCHECO$, 1)
        PF1% = POSFIELD%(ARCHECO$, 1)
        LT1% = LF1%: LD1% = 0: JT1% = 1
        If TF1% <= 2 Then
              LT1% = 7
              JT1% = 2
            ElseIf TF1% <= 4 Then
              LT1% = 15
              LD1% = 2
              JT1% = 2
            ElseIf TF1% = 5 Then
              LT1% = 8
            ElseIf TF1% = 7 Then
              LT1% = 16
        End If
        tf2% = TIPFIELD%(ARCHECO$, 2)
        LF2% = LENFIELD%(ARCHECO$, 2)
        pf2% = POSFIELD%(ARCHECO$, 2)

        LT2% = LF2%: LD2% = 0: JT2% = 1
        If tf2% <= 2 Then
              LT2% = 7
              JT2% = 2
            ElseIf tf2% <= 4 Then
              LT2% = 15
              LD2% = 2
              JT2% = 2
            ElseIf tf2% = 5 Then
              LT2% = 8
            ElseIf tf2% = 7 Then
              LT2% = 16
        End If
        '
       FIN& = ULTREG&(ARCHECO$)
        For I& = 1 To FIN&
           X$ = REGLEIDO$(ARCHECO$, I&)
           XXZ$ = Mid$(X$, PF1%, LF1%)
           FI1$ = CSTRING$(XXZ$, TF1%, LT1%, LD1%, JT1%)
           XXZ$ = Mid$(X$, pf2%, LF2%)
           FI1$ = FI1$ + "  " + CSTRING$(XXZ$, tf2%, LT2%, LD2%, JT2%)
        Next I&
        YALIS%(Index) = 1
        Screen.MousePointer = 1
    End If
    '
    '
End Sub
'

Sub LISTABU()
     
     '
     Screen.MousePointer = 11
     '
     largomax = LARGOJA
     If ORIENTA% = 1 Then
         If largomax > LARGOPAP * 56.7 Then
             largomax = LARGOPAP * 56.7
         End If
       ElseIf ORIENTA% = 2 Then
         If largomax > ANCHOPAP * 56.7 Then
             largomax = ANCHOPAP * 56.7
         End If
     End If
     '
     PRIHOJ% = 1
     If ARCHISAL$ = "" Or Left$(ARCHISAL$, 3) = "LPT" Or ATRAN% = 1 Then
         CAHOJ& = ULTREG&(ARCHITRA$) / (IDL% - RGF& - 4)
'         If CAHOJ& >= 3 Then
'             PRIHOJ% = Int(Val(InputBox$("Imprime desde Hoja Numero: 1", 4)))
'         End If
       Else
         IDL% = 32767
     End If
     '
     PRIENCA% = 1
     '
     FINX& = ULTREG&(ARCHITRA$)
     'WLISTAB.ProgressBar1.Min = 0
     'If FINX& > 0 Then
     '   WLISTAB.ProgressBar1.Max = FINX&
     'End If
     '
     For K1& = HDR& + 1 To FINX&
        '
        CORTECON% = 0
        'WLISTAB.ProgressBar1.Value = K1&
        Call TRACE(20, K1&, FINX&)
        '
        'If BOSTOP%("Pulse Boton de Parada\o Tecla de  <ESCAPE>\para Interrumpir Listador") Then
         ' TINTE$ = Time$
'          146011  X% = COMALTER%("Impresion Interrumpida a su Comando\  \Fecha: " + HO$ + "    Hora: " + TINTE$ + "\\Reanudar\Encuadrar\Abortar Listado")
         ' On X% GoTo 14602, 14603, 14604
'          14602  ' Call TRAZA(21, FINX% + 1, FINX% + 1)
          'GoTo 14605
'          14603   'Call MENSERR(24, "Opcion no Disponible por el Momento")
          'GoTo 146011
'          14604   'Printer.Print , Space$(OFS%) + "INTERRUMPIDO" + Mid$(LII$(3), 13 + OFS%)
          'GoSub 25500                                  ' PRINTER.PRINT, CCI$(2);
          'INTE% = 1
          'GoTo 3800
        'End If
        '
        XXX$ = REGLEIDO$(ARCHITRA$, K1&)
        If K1& = HDR& + 1 Then
            If TNC% > 0 Then
                Call ENCAHOJA1(XXX$)
                Call ENCAHOJA
            End If
        End If
        '
        If XXX$ = String$(RECLEN%(ARCHITRA$), "=") Then
            ULTIHOJ% = 1
            If K1& < ULTREG&(ARCHITRA$) Then
              Call LITRANSPO
              XXX2$ = REGLEIDO$(ARCHITRA$, K1& + 1)
              If ARCHISAL$ <> "" Then
                  Print #ASAL%, "   " ' LII$(0)
                  Print #ASAL%, LII$(4)
                  Print #ASAL%, "   " ' LII$(0)
                  Call INITOTA
                  Call ENCAHOJA1(XXX2$)
                  If TNC% > 0 Then
                      Call ENCAHOJA
                  End If
                Else
                  Call EJECT
                  Printer.NewPage
                  NUHOJ% = 0
                  Call INITOTA
                  Call ENCAHOJA1(XXX2$)
                  Call ENCAHOJA
              End If
            End If
            '
            For K2% = 1 To NCL%
                VFSA$(K2%) = ""
            Next K2%
            GoTo 99
        End If
        '
        If ARCHISAL$ = "" Then
           '
           If Printer.CurrentY > largomax - (7 + 6 * TOTALI%) * ALTUREN Then
           'If ARCHISAL$ = "" Then
               ULTIHOJ% = 0
               Call LITRANSPO
               '
               Call EJECT
               For K2% = 1 To NCL%
                  VFSA$(K2%) = ""
               Next K2%
               Printer.NewPage
               Call ENCAHOJA1(XXX$)
               Call ENCAHOJA
               '
               If TOTALI% > 0 Then
                  If ARCHISAL$ = "" Then
                      Printer.Print MIZQ$ + LII$(0)
                      If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                      Printer.Print MIZQ$ + RETRA$
                      If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                      Printer.Print MIZQ$ + LII$(0)
                      If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                    Else
                      Print #ASAL%, LII$(0)
                      Print #ASAL%, RETRA$
                      Print #ASAL%, LII$(0)
                  End If
               End If
           End If
           '
        End If
        '
        If XXX$ = REGBLAN$(ARCHITRA$) Then
            RENGLO$ = LII$(0)
            For K3% = 1 To NCL%
               VFSA$(K3%) = 0
            Next K3%
            GoTo 5
        End If
        '
        If XXX$ = String$(RECLEN%(ARCHITRA$), "-") Then
            RENGLO$ = LII$(4)
            GoTo 5
        End If
        '
        If XXX$ = String$(RECLEN%(ARCHITRA$), "*") Then
            RENGLO$ = LII$(5)
            GoTo 5
        End If
        '
        KI% = 0: KF% = 0
        For K2% = 1 To NCL%
            If PE%(K2%) < 3 Then
               If CFT%(K2%) = 1 Then
                     XXZ$ = MKS$(K1& + HDR&)
                     If NFT% > 0 Then XXZ$ = RECNO$
                   Else
                     XXZ$ = Mid$(XXX$, PR%(CFT%(K2%)), LF%(CFT%(K2%)))
               End If
               '
               TIVA% = TVL%(CFT%(K2%))
               If TIVA% = 6 Then
                 If Left$(XXZ$, 1) = "@" Then
                   For K3% = 1 To NCL%
                     VFS$(K3%) = Space$(ACL%(K3%))
                   Next K3%
                   VFS$(K2%) = Mid$(XXZ$, 2)
                   GoTo 4
                 End If
               End If
               '
               LOTOT% = ACL%(K2%)
               LODEC% = LDL%(CFT%(K2%))
               JUSTI% = JT%(K2%)
               '
               VFS$(K2%) = CSTRING$(XXZ$, TIVA%, LOTOT%, LODEC%, JUSTI%)
               If CE%(K2%) > 0 Then
                  VFS$(CE%(K2%)) = AJUSTI$(ECOARCH$(AE$(K2%), XXZ$), ACL%(CE%(K2%)))
               End If
               '
               VALO# = 0: If TIVA% <= 4 Then VALO# = Val(VFS$(K2%))
               If TIVA% = 8 Then
                  VALO# = Val(CSTRING$(XXZ$, 4, 12, 2, 1))
                  VFX#(K2%) = VALO#
               End If
               If TT%(K2%) > K2% Then
                  TT1% = TT%(K2%) Mod 100: SIG% = 1
                  If TT%(K2%) > 100 Then SIG% = (-1)
                  If TT1% <> 99 Then SUPAR#(TT1%) = SUPAR#(TT1%) + VALO# * SIG%
               End If
               '
            End If
        Next K2%
        '
        For K2% = 1 To NCL%
          On Error GoTo 0
          If TT%(K2%) = 99 Then
            YYY$ = REGLEIDO$(ARCHITRA$, K1& + 1)
            If Mid$(YYY$, PR%(CFT%(K2%)), LF%(CFT%(K2%))) <> Mid$(XXX$, PR%(CFT%(K2%)), LF%(CFT%(K2%))) Then
              CORTECON% = 1
            End If
          End If
          '
          If TVL%(CFT%(K2%)) < 5 Or TVL%(CFT%(K2%)) = 8 Then
            If TVL%(CFT%(K2%)) <> 8 Then VFX#(K2%) = 0
            LOTOT% = ACL%(K2%)
            LODEC% = LDL%(CFT%(K2%))
            TIVA% = TVL%(CFT%(K2%))
            '
            If TT%(K2%) = 255 Then
                 VFX#(K2%) = SUPAR#(K2%) ' : GoTo 14660
                 VFS$(K2%) = NUMSTRI$(VFX#(K2%), LOTOT%, LODEC%, 0)
               Else
                 '
                 If TT%(K2%) > 0 And TT%(K2%) < K2% Then
                    VFX#(K2%) = 0
                    YYY$ = REGLEIDO$(ARCHITRA$, K1& + 1)
                    If Mid$(YYY$, PR%(CFT%(TT%(K2%))), LF%(CFT%(TT%(K2%)))) <> Mid$(XXX$, PR%(CFT%(TT%(K2%))), LF%(CFT%(TT%(K2%)))) Then
                       If KI% < TT%(K2%) Then KI% = TT%(K2%)
                       VFX#(K2%) = SUPAR#(K2%)
                       SUPAR#(K2%) = 0
                       If K2% > KF% Then KF% = K2%
                    End If
                    XXX$ = REGLEIDO$(ARCHITRA$, K1&)
                    VFS$(K2%) = NUMSTRI$(VFX#(K2%), LOTOT%, LODEC%, 0)
                    If TIVA% = 8 Then
                       TIETO = VFX#(K2%)
                       JUSTI% = JT%(K2%)
                       VFS$(K2%) = CSTRING$(MKS$(TIETO), TIVA%, LOTOT%, LODEC%, JUSTI%)
                       If TIETO < 0.01 Then VFS$(K2%) = Space$(LOTOT%)
                    End If
                End If
            End If
            '
            If TT%(K2%) > 0 Then
               For K4% = 0 To 2
                  TIVA% = TVL%(CFT%(K2%))
                  If TIVA% <> 8 Then
                       SUMA#(K2%, K4%) = SUMA#(K2%, K4%) + CDbl(Val(VFS$(K2%)))
                     Else
                       SUMA#(K2%, K4%) = SUMA#(K2%, K4%) + VFX#(K2%)
                  End If
                  If TT%(K2%) = 255 Then
                      SUMA#(K2%, 1) = 0
                      SUMA#(K2%, 2) = CDbl(Val(VFS$(K2%)))
                  End If
               Next K4%
            End If
            '
          End If
        Next K2%
        '
4       RENGLO$ = Space$(OFS%) + CXC$
        For K2% = 1 To NCL%
           '
           TPAR$ = VFS$(K2%)
           If LK%(K2%) = 0 Then
               If VFS$(K2%) = VFSA$(K2%) Then
                   TPAR$ = Space$(Len(VFS$(K2%)))
                 Else
                   For KK22% = K2% + 1 To NCL%
                     VFSA$(KK22%) = ""
                   Next KK22%
               End If
           End If
           VFSA$(K2%) = VFS$(K2%)
           If JT%(K2%) = 1 Then
               TPAR$ = Space$(MG%(K2%)) + TPAR$
             ElseIf JT%(K2%) = 2 Then
               TPAR$ = TPAR$ + Space$(MG%(K2%))
           End If
           '
           RENGLO$ = RENGLO$ + TPAR$ + CXC$
        Next K2%
        '
        '
5       RGXX1$ = RENGLO$
        RENGLO$ = REPLACAR$(RGXX1$, "@", "~")
        If ARCHISAL$ = "" Then
             Printer.Print MIZQ$ + RENGLO$
             If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
           Else
             Print #ASAL%, RENGLO$
        End If
        '
        If KI% > 0 Then
            RESEP$ = CXC$: If KI% = 1 Then RESEP$ = CCC$(5)
            
            For K2% = 1 To NCL%
               If K2% < KI% Or K2% > KF% Then
                    RESEP$ = RESEP$ + Space$(MG%(K2%) + Len(VFS$(K2%)))
                    If (K2% < KI% - 1 Or K2% > KF%) Then RESEP$ = RESEP$ + CXC$ Else RESEP$ = RESEP$ + CCC$(5)
                  Else
                    RESEP$ = RESEP$ + String$(MG%(K2%) + Len(VFS$(K2%)), "-") + CCC$(5)
                End If
                VFSA$(K2%) = ""
            Next K2%
            '
            If K1& >= FINX& Then
                RESEP$ = LII$(3)
            End If
            '
            If ARCHISAL$ = "" Then
                 Printer.Print MIZQ$ + RESEP$
                 If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                 UIU% = 1
               Else
                 Print #ASAL%, RESEP$
                 UIU% = 2
            End If
        End If
        '
        If CORTECON% = 1 Then
          KI% = 0: KF% = 0
          For K2% = 1 To NCL%
            If KI% = 0 Then If TT%(K2%) = 99 Then KI% = K2%
            TIVA% = TVL%(CFT%(K2%))
            If TIVA% > 2 And TIVA% < 5 Then KF% = K2%
          Next K2%
          '
          RESEP$ = CXC$: If KI% = 1 Then RESEP$ = CCC$(5)
          For K2% = 1 To NCL%
            If K2% < KI% Or K2% > KF% Then
                 RESEP$ = RESEP$ + Space$(MG%(K2%) + Len(VFS$(K2%)))
                 If (K2% < KI% - 1 Or K2% > KF%) Then RESEP$ = RESEP$ + CXC$ Else RESEP$ = RESEP$ + CCC$(5)
               Else
                 RESEP$ = RESEP$ + String$(MG%(K2%) + Len(VFS$(K2%)), "-") + CCC$(5)
             End If
          Next K2%
          '
          RENGLO$ = Space$(OFS%) + CXC$
          For K2% = 1 To NCL%
            TIVA% = TVL%(CFT%(K2%))
            LOTOT% = ACL%(K2%)
            TPAR$ = Space$(LOTOT%)
            If TIVA% > 2 And TIVA% < 5 Then
              LODEC% = LDL%(CFT%(K2%))
              JUSTI% = JT%(K2%)
              TPAR$ = CSTRING$(MKD$(SUMA#(K2%, 0)), TIVA%, LOTOT%, LODEC%, JUSTI%)
            End If
            If JT%(K2%) = 1 Then
               TPAR$ = Space$(MG%(K2%)) + TPAR$
              ElseIf JT%(K2%) = 2 Then
               TPAR$ = TPAR$ + Space$(MG%(K2%))
            End If
            RENGLO$ = RENGLO$ + TPAR$ + CXC$
          Next K2%
          '
          RESUB$ = LII$(0)
          For KU% = 1 To Len(RESUB$)
            LTXY% = Asc(Mid$(RENGLO$, KU%, 1) + " ")
            If LTXY% > 32 Then
              If CXC$ = "" Or LTXY% <> Asc(CXC$ + " ") Then
                Mid$(RESUB$, KU%, 1) = "="
              End If
            End If
          Next KU%
          '
          If ARCHISAL$ = "" Then
               Printer.Print MIZQ$ + RESEP$
               Printer.Print MIZQ$ + RENGLO$
               Printer.Print MIZQ$ + RESUB$
               Printer.Print MIZQ$ + LII$(0)
               If INTERLIX > 0 Then CURRY = CURRY + 4 * INTERLIX: Printer.CurrentY = CURRY
             Else
               Print #ASAL%, RESEP$
               Print #ASAL%, RENGLO$
               Print #ASAL%, RESUB$
               Print #ASAL%, LII$(0)
          End If
          '
          For K2% = 1 To NCL%
            SUMA#(K2%, 0) = 0
          Next K2%
          
        End If
        '
99  Next K1&
    '
    If ARCHISAL$ = "" Then
         Printer.Print MIZQ$ + LII$(3)
         If INTERLIX > 0 Then CURRY = CURRY + 4 * INTERLIX: Printer.CurrentY = CURRY
       Else
         Print #ASAL%, LII$(3)
    End If
    '
    ULTIHOJ% = 1
    Call LITRANSPO
    If ARCHISAL$ = "" Then
        Call EJECT
        For K2% = 1 To NCL%
            VFSA$(K2%) = ""
        Next K2%
190     On Error GoTo 199
        Printer.EndDoc
        On Error GoTo 0
    End If
    '
    If ASAL% > 0 Then
        Close #ASAL%
        ASAL% = 0
    End If
    '
    Screen.MousePointer = 1
    Exit Sub
    '
199 ERNU$ = TRIM$(Str$(Err.Number))
    MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") en la Impresión de\"
    MENSA$ = MENSA$ + "este Listado.\  \"
    '
    MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
    MENSA$ = MENSA$ + "   1.- Impresora Apagada o fuera de Línea.\"
    MENSA$ = MENSA$ + "   2.- Falta Papel o Papel Atorado.\"
    MENSA$ = MENSA$ + "   3.- Problemas en Cable de Señal de Impresora.\"
    MENSA$ = MENSA$ + "   4.- Problemas de Red - No hay Acceso a Impre-\"
    MENSA$ = MENSA$ + "        sora Remota.\"
    MENSA$ = MENSA$ + "   5.- Impresora de Matriz de Puntos Configurada\"
    MENSA$ = MENSA$ + "        para 'Calidad de Imprenta'. Corrija la Con-\"
    MENSA$ = MENSA$ + "        figuración para 'Borrador'.\"
    MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
    '
    C$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(C$, 53, "Error de Impresión de Listado")
    If RESPU% = 4 Then
        Resume 190
    End If
    Call CIERRARCH("*.*")
    Call FINAL("")
    Close: End
    '
End Sub
'

Sub LITRANSPO()
    '
    If TOTALI% > 0 Then
        '
        RETRA$ = Space$(OFS%) + CXC$
        For K2% = 1 To NCL%
           LOTOT% = ACL%(K2%)
           LODEC% = LDL%(CFT%(K2%))
           '
           TPAR$ = Space$(LOTOT%)
           If K2% = POSITOT% Then
                TPAR$ = AJUSTI$("Transporte", LOTOT%)
                If ULTIHOJ% = 1 Then
                   TPAR$ = AJUSTI$("T O T A L   G E N E R A L", LOTOT%)
                End If
              ElseIf TT%(K2%) > 0 Then
                TPAR$ = NUMSTRI$(SUMA#(K2%, 2), LOTOT%, LODEC%, 0)
                If TVL%(CFT%(K2%)) = 8 Then
                   TIETO = SUMA#(K2%, 2)
                   TPAR$ = CSTRING$(MKS$(TIETO), 8, LOTOT%, LODEC%, JT%(K2%))
                End If
           End If
           '
           If JT%(K2%) = 1 Then
                TPAR$ = Space$(MG%(K2%)) + TPAR$
              ElseIf JT%(K2%) = 2 Then
                TPAR$ = TPAR$ + Space$(MG%(K2%))
           End If
           '
           RETRA$ = RETRA$ + TPAR$ + CXC$
           '
        Next K2%
        '
        If ULTIHOJ% = 0 Then
            If ARCHISAL$ = "" Then
                Printer.Print MIZQ$ + LII$(0)  ' renglon en blanco
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                Printer.Print MIZQ$ + RETRA$
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                Printer.Print MIZQ$ + LII$(0)
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
              Else
                Print #ASAL%, LII$(0)    ' renglon en blanco
                Print #ASAL%, RETRA$
                Print #ASAL%, LII$(0)
             End If
        End If
        '
        If ARCHISAL$ = "" Then
             Printer.Print MIZQ$ + LII$(2)
             If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
        End If
        '
        RESUB$ = Space$(OFS%) + CXC$
        For K2% = 1 To NCL%
           LOTOT% = ACL%(K2%)
           LODEC% = LDL%(CFT%(K2%))
           '
           TPAR$ = Space$(LOTOT%)
           If K2% = POSITOT% Then
                TPAR$ = AJUSTI$("Parciales", LOTOT%)
              ElseIf TT%(K2%) > 0 Then
                TPAR$ = NUMSTRI$(SUMA#(K2%, 1), LOTOT%, LODEC%, 0)
                If TVL%(CFT%(K2%)) = 8 Then
                   TIETO = SUMA#(K2%, 1)
                   TPAR$ = CSTRING$(MKS$(TIETO), 8, LOTOT%, LODEC%, JT%(K2%))
                End If
           End If
           SUMA#(K2%, 1) = 0
           '
           If JT%(K2%) = 1 Then
                TPAR$ = Space$(MG%(K2%)) + TPAR$
              ElseIf JT%(K2%) = 2 Then
                TPAR$ = TPAR$ + Space$(MG%(K2%))
           End If
           '
           RESUB$ = RESUB$ + TPAR$ + CXC$
           '
        Next K2%
        '
        If ARCHISAL$ = "" Then
             Printer.Print MIZQ$ + RESUB$
             If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
             Printer.Print MIZQ$ + LII$(3)
             If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
        End If
        '
        If ULTIHOJ% = 1 Then
            '
            Dim HDRX As String * 256
            Dim HDRZ$(32)
            CASBT% = 0
            If EXISTE%(LUCOM$ + ARCHITRA$ + ".SBT") > 0 Then
               NXX% = FILEOPEN%(LUCOM$ + ARCHITRA$ + ".SBT", 0, 256)
               For u& = 1 To LOF(NXX%) / 256
                 If u& <= 32 Then
                   Get #NXX%, u&, HDRX$
                   HDRZ$(u&) = Left$(HDRX$, ATOT%)
                   CASBT% = u&
                 End If
               Next u&
               Close #NXX%
            End If
            '
            If ARCHISAL$ = "" Then
                Printer.Print MIZQ$ + " "
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                Printer.Print MIZQ$ + LII$(1)
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                Printer.Print MIZQ$ + LII$(0)  ' renglon en blanco
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                If CASBT% > 0 Then
                  For SBT% = 1 To CASBT%
                    Printer.Print MIZQ$ + HDRZ$(SBT%)
                    If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                  Next SBT%
                  Printer.Print MIZQ$ + LII$(0)
                  If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                End If
                Printer.Print MIZQ$ + RETRA$
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                Printer.Print MIZQ$ + LII$(0)
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                Printer.Print MIZQ$ + LII$(3)
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
              Else
                Print #ASAL%, " "
                Print #ASAL%, LII$(1)
                Print #ASAL%, LII$(0)    ' renglon en blanco
                If CASBT% > 0 Then
                  For SBT% = 1 To CASBT%
                    Print #ASAL%, HDRZ$(SBT%)
                  Next SBT%
                  Print #ASAL%, LII$(0)
                End If
                Print #ASAL%, RETRA$
                Print #ASAL%, LII$(0)
                Print #ASAL%, LII$(3)
            End If
        End If
        '
      Else
        '
        If ARCHISAL$ = "" Then
             Printer.Print MIZQ$ + LII$(3)
             If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
           Else
             Print #ASAL%, LII$(3)
        End If
        '
    End If
    '
End Sub
'

Sub CARDATLIS()
     '
5800 GRABACT% = 0
     TOTALI% = 0
     POSITOT% = 0
     '
     Dim YYY As String * 128
     If N15% > 0 Then Close #N15%
     N15% = FILEOPEN%(FORMA$, 9, 128)
     Get #N15%, 1, YYY$
     '
     If LOF(N15%) <= 128 Then
         On Error Resume Next
         YYY$ = String$(128, 0)
         Put #N15%, 1, YYY$
         Close #N15%
         Kill FORMA$
         On Error GoTo 0
         FORMB$ = UCase$(REPLACAR$(FORMA$, "\", "/"))
         Call MENSERR(24, "Proceso Interrumpido.\Falta Controlador de Procesos '" + FORMB$ + "'.")
         Call FINAL("")
     End If
     '
     Archi$ = Mid$(YYY$, 57, 8)
     ARCHITRA$ = TRIM$(Archi$)
     WLISTAB.Label5 = ARCHITRA$
     WLISTAB.Label6 = TITUARCH$(ARCHITRA$)
5801 Get #N15%, 1, YYY$
     LOR1% = Asc(Mid$(YYY$, 106, 1)): If LOR1% > 16 Then LOR1% = 0
     LOR% = 256 * LOR1% + Asc(Mid$(YYY$, 107))
     CAMAX& = CVI(Mid$(YYY$, 108))
     PEOF% = 1 + Asc(Mid$(YYY$, 110))
     EOFS$ = Mid$(YYY$, 111, 16)
     WLISTAB.Label8 = PROPRIN$
5803 S1$ = TRIM$(Left$(YYY$, 48))
     WLISTAB.Label9 = TRIM$(S1$)
     '
     EPAC$ = TRIM$(EMPREAC$)
     If EPAC$ <> "" Then
         WLISTAB.Caption = WLISTAB.Caption + "  -  " + EPAC$
     End If
    '
     Sali$ = Mid$(YYY$, 49, 8)
     ARCHISAL$ = Mid$(YYY$, 65, 12)
     If Left$(ARCHISAL$, 3) = "LPT" Then ARCHISAL$ = ""
     COD$ = Mid$(YYY$, 81)
     RGI& = Asc(Mid$(COD$, 1))
     RGF& = Asc(Mid$(COD$, 2))
     OFS% = Asc(Mid$(COD$, 3))
5804 BLP% = Asc(Mid$(COD$, 4))
        CXC$ = "": If BLP% = 1 Then CXC$ = "¦"
     VIV% = Asc(Mid$(COD$, 5))
     APQ% = Asc(Mid$(COD$, 6))
     IDL% = REHOJ%: REHOJ1% = REHOJ%
     If Asc(Mid$(COD$, 7)) < REHOJ% Then IDL% = Asc(Mid$(COD$, 7))
     '
     If VIV% = 1 Then IDL% = 8 * IDL% / 6
     REHOJX% = 0
     If Asc(Mid$(COD$, 7)) > 96 Then
       REHOJX% = Asc(Mid$(COD$, 7))
     End If
     SUPRIFE% = 0
     HDR& = Asc(Mid$(COD$, 8))
     If HDR& = 99 Then
        HDR& = 0
        SUPRIFE% = 1
     End If
     RFT% = Asc(Mid$(COD$, 9))
     NCL% = 0: TCL% = 0: TNC% = 0
     RG% = 1: HST% = 0
     HOJINI% = CVI(Mid$(COD$, 10, 2))
5805 RG% = RG% + 1
     If RG% > LOF(N15%) / 128 Or RG% > 64 Then GoTo 5820
5806 Get #N15%, RG%, YYY$
     If YYY$ = String$(128, 0) Then GoTo 5820
     '
5807 TCL% = TCL% + 1
     TVL%(TCL%) = Asc(Mid$(YYY$, 33))
         If RG% = 2 Then TVL%(TCL%) = 1
     LDL%(TCL%) = Asc(Mid$(YYY$, 34))
5808 PR%(TCL%) = Asc(Mid$(YYY$, 35))
     While PR%(TCL%) < PR%(TCL% - 1)
       PR%(TCL%) = PR%(TCL%) + 256
     Wend
     LF%(TCL%) = Asc(Mid$(YYY$, 36))
         If RG% = 2 Then PR%(TCL%) = 0: LF%(TCL%) = 2
         '
     If Asc(Mid$(YYY$, 72)) = 99 Then
        TNC% = TNC% + 1
        CLNC%(TNC%) = TCL%
     End If
     '
5811 If Asc(Mid$(YYY$, 69)) < 1 Then GoTo 5817   ' SUPRIME COLUMNAS LONGITUD 0
5812 NCL% = NCL% + 1
     CFT%(NCL%) = RG% - 1
     CL0%(NCL%) = RG% - 1
     ECL$(NCL%) = TRIM$(Mid$(YYY$, 37, 32))
5813 ACL%(NCL%) = Asc(Mid$(YYY$, 69))
     LK%(NCL%) = Asc(Mid$(YYY$, 70))
5814 MG%(NCL%) = Asc(Mid$(YYY$, 71))
     TT%(NCL%) = Asc(Mid$(YYY$, 73))
        If TT%(NCL%) = 1 Then TT%(NCL%) = TCL%
        If TT%(NCL%) = 255 Then HST% = 1
5815 PE%(NCL%) = 1
        If RG% = 2 Then PE%(NCL%) = 2
5816 JT%(NCL%) = Asc(Mid$(YYY$, 72))
     AE$(NCL%) = Mid$(YYY$, 74, 8)
     CE%(NCL%) = Asc(Mid$(YYY$, 82))
     If TT%(NCL%) > 0 Then
        NTT% = NTT% + 1
        TOTALI% = 1
     End If
     '
     If POSITOT% = 0 Then
        If TT%(NCL%) < 1 Then
           If TVL%(TCL%) = 6 Then
              If ACL%(NCL%) >= 10 Then
                 POSITOT% = NCL%
              End If
           End If
        End If
     End If
     
5817 GoTo 5805
     '
5820 If POSITOT% = 0 Then POSITOT% = 1
     ATOT% = BLP% * (NCL% + 1)
     CLII% = OFS% + 2
     For KCL% = 1 To NCL%
          LC%(KCL%) = ACL%(KCL%)
          ATOT% = ATOT% + ACL%(KCL%) + MG%(KCL%)
          CL%(KCL%) = CLII%
          CLII% = CLII% + ACL%(KCL%) + 1
     Next KCL%
5821 For KCL% = 1 To NCL%
          If CE%(KCL%) <= 0 Then GoTo 5824
5822      For KCL1% = 1 To NCL%
              If CL0%(KCL1%) = CE%(KCL%) Then CE%(KCL%) = KCL1%: PE%(KCL1%) = 3: GoTo 5824
5823      Next KCL1%
5824 Next KCL%
5825 For KCL% = 1 To NCL%
          If TT%(KCL%) <= 0 Then GoTo 5828
5826      For KCL1% = 1 To NCL%
             If CL0%(KCL1%) = TT%(KCL%) Then TT%(KCL%) = KCL1%: GoTo 5828
             If 100 + CL0%(KCL1%) = TT%(KCL%) Then TT%(KCL%) = 100 + KCL1%: GoTo 5828
5827      Next KCL1%
5828 Next KCL%
     '
5830 NFT% = 0    ' NO HAY FILTROS
5850 If N15% > 0 Then
         Close #N15%
         N15% = 0
     End If
     '
End Sub
'
Sub CARDATSOR()
     '
     Dim YYY As String * 128
     If N15% > 0 Then Close #N15%
     N15% = FILEOPEN%(FORMA$, 9, 128)
     Get #N15%, 1, YYY$
     '
     If LOF(N15%) <= 128 Then
         On Error Resume Next
         YYY$ = String$(128, 0)
         Put #N15%, 1, YYY$
         Close #N15%
         Kill FORMA$
         On Error GoTo 0
         FORMB$ = UCase$(REPLACAR$(FORMA$, "\", "/"))
         Call MENSERR(24, "Proceso Interrumpido.\Falta Controlador de Procesos '" + FORMB$ + "'.")
         Call FINAL("")
     End If
     '
     ARCHITRA$ = TRIM$(UCase$(Mid$(YYY$, 57, 8)))
     WSORTAB.Text1.Text = ARCHITRA$
     WSORTAB.Text3.Text = TITUARCH$(ARCHITRA$)
     ARCHISAL$ = TRIM$(UCase$(Mid$(YYY$, 65, 8)))
     If ARCHISAL$ <> ARCHITRA$ Then
         Call COPYSTRU(ARCHITRA$, ARCHISAL$)
     End If
     WSORTAB.Text4.Text = ARCHISAL$
     WSORTAB.Text2.Text = TITUARCH$(ARCHISAL$)
     WSORTAB.Refresh
     '
     CALIMI% = Asc(Mid$(YYY$, 73, 1))
     MORDE% = Asc(Mid$(YYY$, 74, 1))
     CALIMS% = Asc(Mid$(YYY$, 75, 1))
        If CALIMS% < CALIMI% Then CALIMS% = CALIMI%
     LOR1% = Asc(Mid$(YYY$, 106, 1)): If LOR1% > 16 Then LOR1% = 0
     LOR% = 256 * LOR1% + Asc(Mid$(YYY$, 107))
     CAMAX& = CVI(Mid$(YYY$, 108))
     PEOF% = Asc(Mid$(YYY$, 110))
     EOFS$ = Mid$(YYY$, 111, 16)
     '
5803 S1$ = Left$(YYY$, 48)
     Sali$ = Mid$(YYY$, 49, 8)
     COD$ = Mid$(YYY$, 81)
     RGI& = Asc(Mid$(COD$, 1))
     RGF& = Asc(Mid$(COD$, 2))
     OFS% = Asc(Mid$(COD$, 3))
5804 BLP% = Asc(Mid$(COD$, 4))
     VIV% = Asc(Mid$(COD$, 5))
     APQ% = Asc(Mid$(COD$, 6))
     IDL% = Asc(Mid$(COD$, 7))
     HDR& = Asc(Mid$(COD$, 8))
     RFT% = Asc(Mid$(COD$, 9))
     NCL% = 0: TCL% = 0: RG% = 2
     '
5805 RG% = RG% + 1
     If RG% > LOF(N15%) / 128 Or RG% > 33 Then GoTo 5820
5806 Get #N15%, RG%, YYY$
     If YYY$ = String$(128, 0) Then GoTo 5820
     '
5807 TCL% = TCL% + 1
     TVL%(TCL%) = Asc(Mid$(YYY$, 33))
     LDL%(TCL%) = Asc(Mid$(YYY$, 34))
     If RG% = 2 Then
         TVL%(TCL%) = 1
         LDL%(TCL%) = 0
     End If
5808 PR%(TCL%) = Asc(Mid$(YYY$, 35))
     LF%(TCL%) = Asc(Mid$(YYY$, 36))
     If RG% = 2 Then
         PR%(TCL%) = 0
         LF%(TCL%) = 2
     End If
5812 NCL% = NCL% + 1
     CFT%(NCL%) = RG% - 1
     CL0%(NCL%) = RG% - 1
     ECL$(NCL%) = TRIM$(Mid$(YYY$, 37, 32))
5813 ACL%(NCL%) = Asc(Mid$(YYY$, 69))
     LK%(NCL%) = Asc(Mid$(YYY$, 70))
5814 MG%(NCL%) = Asc(Mid$(YYY$, 71))
     PE%(NCL%) = Asc(Mid$(YYY$, 73))
     If RG% = 2 Then PE%(NCL%) = 2
5815 JT%(NCL%) = Asc(Mid$(YYY$, 72))
     AE$(NCL%) = Mid$(YYY$, 74, 8)
     CE%(NCL%) = Asc(Mid$(YYY$, 82))
5816 GoTo 5805
5817 '
5820 ATOT% = NCL% - 1
     CLXX% = OFS% + 2
     For KCL% = 1 To NCL%
        LC%(KCL%) = ACL%(KCL%)
        ATOT% = ATOT% + ACL%(KCL%)
        CL%(KCL%) = CLXX%
        CLXX% = CLXX% + ACL%(KCL%) + 1
     Next KCL%
5821 For KCL% = 1 To NCL%
        If CE%(KCL%) <= 0 Then GoTo 5824
5822    For KCL1% = 1 To NCL%
           If CL0%(KCL1%) = CE%(KCL%) Then
               CE%(KCL%) = KCL1%: PE%(KCL1%) = 3: GoTo 5824
           End If
5823    Next KCL1%
5824 Next KCL%
 '
5830 NFT% = 0
     RG% = 1
     RG0% = RGI&
5835 RG% = RG% + 1
     If RG% > LOF(N15%) / 128 Or RG% > 33 Then GoTo 5850
5840 Get #N15%, RG%, YYY$
     If YYY$ = String$(128, 0) Then GoTo 5850
5841 If Asc(Mid$(YYY$, 88)) < 1 Then GoTo 5847   ' SUPRIME RG-FILTRO=0
5842 NFT% = NFT% + 1
     CFF%(NFT%) = RG% - 1
     LCF%(NFT%) = Asc(Mid$(YYY$, 83))
     CLF%(NFT%) = Asc(Mid$(YYY$, 85))
     RGFJ%(NFT%) = Asc(Mid$(YYY$, 84))
5843 If RGFJ%(NFT%) - 1 < RG0% Then RG0% = RGFJ%(NFT%) - 1
5844 ECL$(NFT%) = TRIM$(Mid$(YYY$, 37, 32))
5845 LKF%(NFT%) = Asc(Mid$(YYY$, 70)): LEF%(NFT%) = Asc(Mid$(YYY$, 86)): CEF%(NFT%) = Asc(Mid$(YYY$, 87)): TFF%(NFT%) = Asc(Mid$(YYY$, 88))
5846 AEF$(NFT%) = Mid$(YYY$, 74, 8): VFMIN$(NFT%) = Mid$(YYY$, 113, 8): VFMAX$(NFT%) = Mid$(YYY$, 121, 8)
5847 GoTo 5835
 '
5850 S1$ = TRIM$(S1$)
     '
5860 Close #N15%: N15% = 0
     '
End Sub
'

Sub SETPRINTER()
    '
    Dim FUENTESEL
    Dim FUENPRUE
    Dim NADAFUEN
    '
    If IDL% = 32767 Then
        For u% = 1 To 17
           CCI$(u%) = ""
        Next u%
        Exit Sub
    End If
    '
    SRY$ = String$(255, "*")
    GU$ = String$(255, "-")
    CXC$ = "": CC$ = ""
       If BLP% = 1 Then CXC$ = "¦": CC$ = "="
    For I% = 1 To 9
         CCC$(I%) = "": If BLP% = 1 Then CCC$(I%) = "+"
    Next I%
    '
    For I% = 1 To 3
       LII$(I%) = Space$(OFS%) + CCC$(I%)
       For J% = 1 To NCL%
          LII$(I%) = LII$(I%) + String$(LC%(J%) + MG%(J%), "-")
          If J% < NCL% Then
              LII$(I%) = LII$(I%) + CCC$(I% + 1)
            Else
              LII$(I%) = LII$(I%) + CCC$(I% + 2)
          End If
       Next J%
    Next I%
    '
    LII$(0) = Space$(OFS%) + CXC$
    LII$(4) = Space$(OFS%) + CC$
    LII$(5) = Space$(OFS%) + CC$
    '
    For J% = 1 To NCL%
       LII$(0) = LII$(0) + Space$(LC%(J%) + MG%(J%)) + CXC$
       LII$(4) = LII$(4) + String$(LC%(J%) + MG%(J%), "=") + CC$
       LII$(5) = LII$(5) + String$(LC%(J%) + MG%(J%), "*") + CC$
    Next J%
    '
    If TRIM$(ARCHISAL$) <> "" Then
        MIZQ$ = ""
        PRINTERSET = 1
    End If
    '
    If PRINTERSET% = 0 Then
        '
        If Printer.CurrentX > 0 Then
            Printer.NewPage
        End If
        '
10      NOMIMPRE$ = TRIM$(Printer.DeviceName)
        If NOMIMPRE$ = "" Then
            Call MENSERR(24, "Falta Seleccionar Impresora\en la Configuración de Windows")
            Call FINAL("")
        End If
        '
        YAESTA% = 0
        CAFUFIJ% = 0
        NX% = FILEOPEN(LUCOM$ + "PRINTERS.DRV", 0, 128)
        FIN& = LOF(NX%) / 128
        For IPRIN& = 1 To FIN&
            Get #NX%, IPRIN&, PPZZ$
            PPXX$ = PPZZ$
            If TRIM$(Left$(PPXX$, 48)) = NOMIMPRE$ Then
                YAESTA% = 1
                If UCase$(Mid$(PPXX$, 125, 2)) = "SI" Then
                    CAFUFIJ% = CAFUFIJ% + 1
                    FUFIJA$(CAFUFIJ%) = PPZZ$
                End If
            End If
        Next IPRIN&
        '
        Close #NX%
        NOMIMP$ = REPLACAR$(NOMIMPRE$, "\", "/")
        If YAESTA% < 1 Then
            OPX% = COMALTER%("Se ha Detectado un Modelo de Impresora\Desconocido para el Sistema Aplicativo\  \'" + NOMIMP$ + "'\  \\Configurar\Cancelar")
            If OPX% <> 1 Then Call FINAL("")
            Call SETUPRINTER
            GoTo 10
        End If
        '
        If CAFUFIJ% < 1 Then
            OPX% = COMALTER%("No hay Fuentes Utilizables para Impresora\  \'" + NOMIMP$ + "'\  \\Configurar\Cancelar")
            If OPX% <> 1 Then Call FINAL("")
            Call SETUPRINTER
            GoTo 10
        End If
        '
        Printer.ScaleMode = 1
        ANCHOJA = 56.7 * (ANCHOPAP - MARGENI)
        If ORIENTA% = 2 Then ANCHOJA = 56.7 * (LARGOPAP - MARGENI)
        'ANCHOMAX = Int(56.7 * CVS(Mid$(FUFIJA$(1), 49, 4)))
        'If ANCHOMAX < ANCHOJA Then ANCHOJA = ANCHOMAX
        '
        ATEXMAX = 0
        TAFUESEL = 0
        '
        TALEBASI = 24
        '
100     TEPRUEBA$ = String$(ATOT%, "*")
        Printer.FontBold = False
        Printer.FontItalic = False
        For I% = 1 To CAFUFIJ%
            '
            Printer.FontName = TRIM$(Mid$(FUFIJA$(I%), 65, 48))
            Printer.FontBold = False
            Printer.FontItalic = False
            '
            FUENPRUE = TRIM$(Mid$(FUFIJA$(I%), 65, 48))
            If Printer.FontName <> FUENPRUE Then
'MsgBox Printer.FontName & TRIM$(Mid$(FUFIJA$(I%), 65, 48))
                GoTo 109   ' modificado 20/11/01 para que busque la fuente
                           ' disponible
                Call COMUNI("Se ha Detectado un Error\en la Configuración Local\de la Impresora Predeterminada.\  \Imposible Emitir el Listado Solicitado.")
                Call CIERRARCH("*.*")
                Close
                Call FILEKILL(LUCOM$ + "PRINTERS.DRV")
                Call FINAL("")
            End If
            '
            FUENTEMIN = CVS(Mid$(FUFIJA$(I%), 113, 4))
            FUENTEMAX = CVS(Mid$(FUFIJA$(I%), 117, 4))
            '
            TALEBA = FUENTEMAX + 0.4
            While TALEBA > FUENTEMIN
               TALEBA = TALEBA - 0.4
               Printer.FontSize = TALEBA
               If Printer.FontSize > TALEBA Then GoTo 105
               ATEXTO = Printer.TextWidth(TEPRUEBA$)
               If ATEXTO <= ANCHOJA Then
                  If ATEXTO > ATEXMAX Then
                    FUENTESEL = FUENPRUE    ' Printer.FontName
                    TAFUESEL = Printer.FontSize ' TALEBASI
                    ATEXMAX = ATEXTO
                  End If
                  GoTo 109
               End If
               TALEBA = Printer.FontSize
105         Wend
        '
109     Next I%
        '
        If FUENTESEL = NADAFUEN Or TAFUESEL < 1 Then
            Call MENSERR(24, "No hay Fuentes Adecuadas\para la Impresion de Este Listado")
            Call FINAL("")
        End If
        '
        Printer.FontName = FUENTESEL
        Printer.FontSize = TAFUESEL
        WLISTAB.Text5.Text = FUENTESEL
        WLISTAB.Text5.Refresh
        WLISTAB.Text7.Text = CSTRING$(MKS$(TAFUESEL), 3, 6, 1, 1)
        WLISTAB.Text7.Refresh
       '
        If FUENTESEL = "" Or TAFUESEL < 0.4 Then
            Call MENSERR(24, "No puede Establecerse\el Formato de Impresion.\Consulte al Soporte de Sistemas.")
            Call FINAL("")
        End If
        '
110     Printer.FontName = FUENTESEL
        Printer.FontSize = TAFUESEL
        LARGOJA = Val(Printer.ScaleHeight)
        'ANCHOJA = Printer.ScaleWidth
        TEPRUEBB$ = TEPRUEBA$ ' : If LX810 = 1 Then TEPRUEBB$ = TEPRUEBA$ + "*****"
        ANCHOJB = Val(Printer.TextWidth(TEPRUEBB$))
        TALEBASI = Printer.FontSize
        ALTUREN = Printer.TextHeight(TEPRUEBA$)
        Printer.FontBold = False
        Printer.FontItalic = False
        MIZMAX = ANCHOJA - ANCHOJB
        If MARGENI > 0 Then
            MIZMAX = 56.7 * MARGENI
        End If
        MIZQ$ = ""
        While Printer.TextWidth(MIZQ$ + "*") <= MIZMAX
           MIZQ$ = MIZQ$ + " "
        Wend
        MIZQ$ = Mid$(MIZQ$, 2)
        PRINTERSET% = 1
        '
    End If
    '
End Sub
'

Sub SETUPRINTER()
    '
    ' WLISTAB.Hide
    '
    NOMIMPRE$ = TRIM$(Printer.DeviceName)
    If NOMIMPRE$ = "" Then
        Call MENSERR(24, "Falta Seleccionar Impresora\en la Configuración de Windows")
        Call FINAL("")
    End If
    '
    YAESTA% = 0
    Call BLANARCH("!PRINTERS")
    NOMI$ = AJUSTI$(NOMIMPRE$, 48)
    NX% = FILEOPEN(LUCOM$ + "PRINTERS.DRV", 0, 128)
    For IPRIN& = 1 To LOF(NX%) / 128
        Get #NX%, IPRIN&, PPZZ$
        PPXX$ = PPZZ$
        If TRIM$(Left$(PPXX$, 48)) = NOMIMPRE$ Then
            Call REGAPP("!PRINTERS", PPXX$)
            YAESTA% = 1
            GoTo 10
        End If
    Next IPRIN&
    '
10  Close #NX%
    If YAESTA% = 1 Then
        PPXX$ = Left$(PPXX$, 52)
        'GoTo 50
    End If
    '
    Printer.ScaleMode = 1
    ANCHOJMM = Int(Printer.ScaleWidth / 56.7)
    Call REPLA(PPXX$, MKS$(ANCHOJMM), 49, 4)
    '
    CAFUFIJ% = 0
    For I% = 1 To Printer.FontCount
        Printer.FontName = Printer.Fonts(I% - 1)
        A1 = 0: A2 = 0
        On Error Resume Next
        A1 = Printer.TextWidth(String$(80, "l"))
        A2 = Printer.TextWidth(String$(80, "M"))
        On Error GoTo 0
        If A1 > 0 And A2 > 0 Then
          If A1 = A2 Then
            If CAFUFIJ% < 128 Then
              CAFUFIJ% = CAFUFIJ% + 1
              FUFIJA$(CAFUFIJ%) = Printer.Fonts(I% - 1)
            End If
          End If
        End If
    Next I%
    '
    If CAFUFIJ% < 1 Then
       Call MENSERR(24, "No se Encontraron Fuentes de Ancho Fijo\para la Impresora Seleccionada.\  \No es Posible Emitir Listado por Impresora.\  \Consulte a su Soporte de Sistemas.")
       Call FINAL("")
    End If
    '
    Call BLANARCH("!PRINTERS")
    TALEBASI = 24
    TEPRUEBA$ = String$(128, "*")
    Printer.FontBold = False
    Printer.FontItalic = False
    For I% = 1 To CAFUFIJ%
        Printer.FontName = FUFIJA$(I%)
        If Printer.FontName <> FUFIJA$(I%) Then GoTo 19
        '
        Printer.FontSize = 24
        FUENTEMAX = Printer.FontSize
        Printer.FontSize = 2
        FUENTEMIN = Printer.FontSize
        '
14      If FUENTEMIN <= FUENTEMAX Then
            XX$ = REGBLAN$("!PRINTERS")
            Call REPLA(XX$, NOMIMPRE$, 1, 48)
            Call REPLA(XX$, MKS$(ANCHOJMM), 49, 4)
            Call REPLA(XX$, FUFIJA$(I%), 65, 48)
            Call REPLA(XX$, MKS$(FUENTEMIN), 113, 4)
            Call REPLA(XX$, MKS$(FUENTEMAX), 117, 4)
            Call REPLA(XX$, "SI", 125, 2)
            Call REGAPP("!PRINTERS", XX$)
        End If
        '
19  Next I%
    Call CIERRARCH("!PRINTERS")
    '
    PPXX$ = AJUSTI$(NOMIMPRE$, 48) + MKS$(ANCHOJMM) + String$(12, 0)
    '
50  PPYY$ = OBJPLA$("DATIMPRE", PPXX$)
    If PPYY$ = "" Then
        Call CIERRARCH("PRINTERS")
        GoTo 99
    End If
    '
    ANCHOJMM = CVS(Mid$(PPYY$, 49, 8))
    If ANCHOJMM > Int(Printer.ScaleWidth / 56.7) Or ANCHOJMM < 48 Then
        PPXX$ = PPYY$
        GoTo 50
    End If
    '
    VTB% = VENTABU%("PRINSET/" + AJUSTI$(NOMIMPRE$, 48))
    If VTB% >= 0 Then
        NX% = FILEOPEN%(LUCOM$ + "PRINTERS.DRV", 0, 128)
        NOMI$ = AJUSTI$(NOMIMPRE$, 48)
        JJ& = 0
        For IPRIN& = 1 To LOF(NX%) / 128
            Get #NX%, IPRIN&, PPZZ$
            If Left$(PPZZ$, 48) <> NOMI$ Then
               JJ& = JJ& + 1
               Put #NX%, JJ&, PPZZ$
            End If
        Next IPRIN&
        For IPRIN& = 1 To ULTREG&("!PRINTERS")
            PPXX$ = REGLEIDO$("!PRINTERS", IPRIN&)
            Call REPLA(PPXX$, NOMI$, 1, 48)
            Call REPLA(PPXX$, MKS$(ANCHOJMM), 49, 4)
            If UCase$(Mid$(PPXX$, 125, 2)) <> "SI" Then
                Call REPLA(PPXX$, "  ", 125, 2)
            End If
            PPZZ$ = PPXX$
            JJ& = JJ& + 1
            Put #NX%, JJ&, PPZZ$
        Next IPRIN&
        For IPRIN& = JJ& + 1 To LOF(NX%) / 128
            PPZZ$ = String$(128, 0)
            Put #NX%, IPRIN&, PPZZ$
        Next IPRIN&
        Close #NX%
    End If
    Call CIERRARCH("!PRINTERS")
    '
99  FORMA$ = FORMALI$
    If TRIM$(FORMA$) <> "" Then
        Call CARDATLIS
        Call CARDATIMPRE
      Else
        Call COMUNI("Para Efectivizar la nueva Configuración\Debe Volver a pedir el Listado.")
        Call FINAL("")
    End If
    '
End Sub
'
Sub BLANHEADERS()
   For u& = 0 To 16
     LEYE$(u&) = ""
   Next u&
End Sub
'
Sub CARHEADERS()
     '
     For u& = 1 To HDR&
         SUBTI$(u&) = REGLEIDO$(ARCHITRA$, u&)
     Next u&
     '
     X1% = 0
     ARCH$ = AJUSTI$(UCase$(ARCHITRA$), 8)
     For X% = 1 To 10
        If LTrim$(RTrim$(LEYE$(X%))) <> "" Then
             X1% = X1% + 1
             LEYE$(X1%) = LEYE$(X%)
        End If
     Next X%
     '
     For u& = 1 To ULTREG&("HEADERS")
        XXX$ = REGLEIDO$("HEADERS", u&)
        If UCase$(Left$(XXX$, 8)) = ARCH$ Then
            For X4% = 1 To X1%
                 If LEYE$(X4%) = Mid$(XXX$, 9) Then GoTo 1024
            Next X4%
            X1% = X1% + 1
            If X1% < 10 Then
                 LEYE$(X1%) = Mid$(XXX$, 9)
            End If
        End If
1024 Next u&
     '
     Call CIERRARCH("HEADERS")
     For X3% = X1% + 1 To 10
            LEYE$(X3%) = ""
     Next X3%
     '
End Sub
'
Sub ARMAENCA()
    '
    If SALEXCEL% = 1 Then
      If EXISTE%(LUCOM$ + ARCHITRA$ + ".HDP") > 0 Then
        HDRX$ = LOADFILE$(LUCOM$ + ARCHITRA$ + ".HDP")
        ENCA$ = AJUSTI(TRIM$(HDRX$), ATOT%)
        '
        For KCL% = 1 To NCL%
          ECLX1$ = TRIM$(Left$(ENCA$, ACL%(KCL%)))
          If Left$(ECL$(KCL%), 1) = "<" Or JT%(KCL%) = 1 Then
              ECL$(KCL%) = "<" + ECLX1$
            ElseIf Left$(ECL$(KCL%), 1) = ">" Or JT%(KCL%) = 2 Then
              ECL$(KCL%) = ">" + ECLX1$
            Else
              ECL$(KCL%) = ECLX1$
          End If
          ENCA$ = Mid$(ENCA$, 1 + ACL%(KCL%) + MG%(KCL%))
        Next KCL%
      End If
    End If
    '
    ENCA$ = Space$(OFS%) + CXC$
    '
    For KCL% = 1 To NCL%
        If Left$(ECL$(KCL%), 1) = "<" Then
              ECL$(KCL%) = Left$(Space$(MG%(KCL%)) + Mid$(ECL$(KCL%), 2) + Space$(ACL%(KCL%)), ACL%(KCL%) + MG%(KCL%))
          ElseIf Left$(ECL$(KCL%), 1) = ">" Then
              ECL$(KCL%) = Right$(Space$(ACL%(KCL%)) + Mid$(ECL$(KCL%), 2) + Space$(MG%(KCL%)), ACL%(KCL%) + MG%(KCL%))
          Else
              While Len(ECL$(KCL%)) < ACL%(KCL%)
                  ECL$(KCL%) = " " + ECL$(KCL%) + " "
              Wend
              ECL$(KCL%) = Space$(MG%(KCL%)) + Left$(ECL$(KCL%), ACL%(KCL%))
        End If
        '
        If Right$(ECL$(KCL%), 2) = ". " Then
            ECL$(KCL%) = " " + Left$(ECL$(KCL%), Len(ECL$(KCL%)) - 1)
        End If
        '
        ENCA$ = ENCA$ + ECL$(KCL%) + CXC$
        '
    Next KCL%
    '
End Sub
'
Sub INITOTA()
    '
    For K3% = 1 To NCL%
        For K4% = 0 To 2
           SUMA#(K3%, K4%) = 0
        Next K4%
    Next K3%
    '
    NUHOJ% = 0
    PPXX% = InStr(ATRIBU$, "PRIHOJA=")
    If PPXX% > 0 Then
       NUHOJ% = Val(Mid$(ATRIBU$, PPXX% + 8)) - 1
    End If
    '
End Sub
'
Sub ENCAHOJA()
    '
    Dim HDRX As String * 256
    Dim HDRZ$(32)
    Dim ENCATAB$(32)
    '
    If InStr(ARCHISAL$, ".XLS") > 0 Then
      Exit Sub
    End If
    '
    NUHOJ% = NUHOJ% + 1
    NH# = CDbl(NUHOJ%)
    '
    For K3% = 1 To NCL%
         VDAA%(K3%) = 0
         VDAB#(K3%) = 0
         VDA$(K3%) = ""
    Next K3%
    '
    If NUHOJ% >= PRIHOJ% Then
           ' Call MEMBRETE
         If ARCHISAL$ = "" Then
           If TNC% < 1 Then
              Printer.CurrentY = MARGENS * 56.7
              CURRY = Printer.CurrentY
              For RI% = 1 To RGI&
                 Printer.Print , " "
                 If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
              Next RI%
           End If
         End If
         '
         If Left$(ARCHISAL$, 3) = "LPT" Then ARCHISAL$ = ""
         '
         ESPAMED% = ATOT% - Len(S1$) - 35
         S2$ = Left$(SRY$, Len(S1$))
         '
         If TNC% > 0 Then GoTo 10
         '
         TEFE$ = "Fecha: " + HOS$: If SUPRIFE% = 1 Then TEFE$ = "               "
         SUFE$ = "******": If SUPRIFE% = 1 Then SUFE$ = "      "
         HDRZ$(1) = S1$ + Space$(ESPAMED%) + "   " + TEFE$ + "    Hoja No:#####"
         HDRZ$(2) = S2$ + Space$(ESPAMED%) + "   " + SUFE$ + "             ********"
         CAENCA% = 2
         '
         If Left$(S1$, 1) = "%" Then
           HDRZ$(1) = Space$(10)
           HDRZ$(2) = Space$(10)
         End If
         '
         CAENCOL% = 0
         If EXISTE%(LUCOM$ + ARCHITRA$ + ".HDR") > 0 Then
            NXX% = FILEOPEN%(LUCOM$ + ARCHITRA$ + ".HDR", 0, 256)
            For u& = 1 To LOF(NXX%) / 256
               If u& <= 32 Then
                  Get #NXX%, u&, HDRX$
                  ENCATAB$(u&) = Left$(HDRX$, ATOT%)
                  CAENCOL% = u&
               End If
            Next u&
            Close #NXX%
         End If
         '
         If NUHOJ% <= PRIHOJ% Or NUHOJ% = 1 Then
            If EXISTE%(LUCOM$ + ARCHITRA$ + ".HDP") > 0 Then
               'CAENCA% = CAENCA% + 1: HDRZ$(CAENCA%) = Space$(ATOT%)
               '
               NXX% = FILEOPEN%(LUCOM$ + ARCHITRA$ + ".HDP", 0, 256)
               'For U& = 1 To 1 + Int((LOF(NXX%) - 1) / 256)
               Get #NXX%, 1, HDRX$
               ENCA$ = AJUSTI(TRIM$(HDRX$), ATOT%)
               '   CAENCA% = CAENCA% + 1
               '   HDRZ$(CAENCA%) = Left$(HDRX$, ATOT%)
               'Next U&
               Close #NXX%
            End If
         End If
         '
         PPXX% = InStr(HDRZ$(1), "#####")
         If PPXX% > 0 Then
            Call REPLA(HDRZ$(1), NUMSTRI$(NH#, 5, 0, 0), PPXX%, 5)
         End If
         '
         For u& = 1 To CAENCA%
            If ARCHISAL$ = "" Then
                 Printer.Print MIZQ$ + Space$(OFS%) + HDRZ$(u&)
                 If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
               Else
                 Print #ASAL%, Space$(OFS%);
                 Print #ASAL%, HDRZ$(u&)
            End If
         Next u&
         '
10       For RI% = 1 To 10
              LEYEN$ = TRIM$(LEYE$(RI%))
              If LEYEN$ <> "" Then
                 If Left$(LEYEN$, 1) <> ":" Then
                    PRI% = InStr(LEYEN$, ":")
                    If PRI% = 0 Then PRI% = Len(LEYEN$) + 1
                    SUBRA$ = String$(PRI%, "*")
                    While Len(SUBRA$) < Len(S1$)
                       LEYEN$ = " " + LEYEN$
                       SUBRA$ = " " + SUBRA$
                    Wend
                    '
                    If ARCHISAL$ = "" Then
                        Printer.Print " "
                        If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                      Else
                        Print #ASAL%, " "
                    End If
                    '
                    LEYF$ = ""
                    If Len(LEYEN$) > ATOT% Then
                       PPXY% = InStr(UCase$(LEYEN$), "HASTA:")
                       If PPXY% > 0 Then
                          LEYF$ = Mid$(LEYEN$, PPXY%)
                          LEYEN$ = Left$(LEYEN$, PPXY% - 1)
                       End If
                    End If
                    '
                    If ARCHISAL$ = "" Then
                         Printer.Print MIZQ$ + Space$(OFS%) + Left$(LEYEN$ + Space$(ATOT%), ATOT%)
                         If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                         Printer.Print MIZQ$ + Space$(OFS%) + Left$(SUBRA$ + Space$(ATOT%), ATOT%)
                         If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                       Else
                         Print #ASAL%, Space$(OFS%) + Left$(LEYEN$ + Space$(ATOT%), ATOT%)
                         Print #ASAL%, Space$(OFS%) + Left$(SUBRA$ + Space$(ATOT%), ATOT%)
                    End If
                    '
                    If Len(LEYF$) > 0 Then
                       If ARCHISAL$ = "" Then
                           Printer.Print MIZQ$ + Space$(OFS%) + Left$(Space$(1 + Len(SUBRA$)) + LEYF$, ATOT%)
                           If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                         Else
                           Print #ASAL%, Space$(OFS%) + Left$(Space$(1 + Len(SUBRA$)) + LEYF$, ATOT%)
                       End If
                    End If
                    '
                 End If
              End If
              '
         Next RI%
         '
         If HDR& > 0 Then
            Printer.Print " "
            If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
            For RI% = 1 To HDR&
               If ARCHISAL$ = "" Then
                    Printer.Print MIZQ$ + Space$(OFS%) + Left$(SUBTI$(RI%), ATOT%)
                    If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                  Else
                    Print #ASAL%, Space$(OFS%) + Left$(SUBTI$(RI%), ATOT%)
               End If
            Next RI%
         End If
         '
30       For RI% = 1 To RFT%
              If ARCHISAL$ = "" Then
                  Printer.Print " "
                  If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                Else
                  Print #ASAL%, " "
              End If
         Next RI%
         '
         If ARCHISAL$ = "" Then
                Printer.Print MIZQ$ + LII$(1)
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                If CAENCOL% < 1 Then
                     Printer.Print MIZQ$ + ENCA$
                     If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                   Else
                     For KKCOL% = 1 To CAENCOL%
                       Printer.Print MIZQ$ + ENCATAB$(KKCOL%)
                       If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                     Next KKCOL%
                End If
                Printer.Print MIZQ$ + LII$(2)
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
              Else
                Print #ASAL%, LII$(1)
                If CAENCOL% < 1 Then
                     Print #ASAL%, ENCA$
                   Else
                     For KKCOL% = 1 To CAENCOL%
                       Print #ASAL%, ENCATAB$(KKCOL%)
                     Next KKCOL%
                End If
                Print #ASAL%, LII$(2)
         End If
         '
    End If
    '
    For K3% = 1 To NCL%
       SUMA#(K3%, 1) = 0
    Next K3%
    '
End Sub
'

Sub ENCAHOJA1(NCXX$)
    '
    NRENCA% = 0
    If ARCHISAL$ = "" Then
        If TNC% > 0 Then
            Printer.CurrentY = 56.7 * MARGENS
            CURRY = Printer.CurrentY
            For RI% = 1 To RGI&
                 Printer.Print , " "
                 If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
            Next RI%
        End If
    End If
    '
    For CLIJ% = 1 To TNC%
        TTXI$ = Mid$(NCXX$, PR%(CLNC%(CLIJ%)), LF%(CLNC%(CLIJ%)))
        If TVL%(CLNC%(CLIJ%)) = 6 Then
            If ARCHISAL$ = "" Then
                Printer.Print MIZQ$ + TTXI$
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
              Else
                Print #ASAL%, TTXI$
            End If
            NRENCA% = NRENCA% + 1
        End If
    Next CLIJ%
    '
    '
End Sub
'

Sub EJECT()
    '
10  On Error GoTo 99
    While Printer.CurrentY < Printer.ScaleHeight - 5 * ALTUREN
        Printer.Print " "
        If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
    Wend
    RNGL$ = MIZQ$ + Space$(OFS%) + String$(ATOT%, "-")
    ANCHOJB = Printer.TextWidth(RNGL$)
    If ULTIHOJ% > 0 Or TOTALI% = 0 Then
        Printer.Print RNGL$
        If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
    End If
    Printer.FontSize = 6
    '
    TXT$ = "Sistema FLEXOFT s/n " + SERINUM$
    TXT$ = TXT$ + " licenciado a " + EMPRELI$
    'If UCase$(FECHALI$) <> "S/LIMITE" Then
    '    TXT$ = TXT$ + " hasta " + FECHALI$
    'End If
    TXT$ = TXT$ + " - Derechos Reservados L.11723"
    TXT$ = MIZQ$ + Space$(OFS%) + TXT$
    '
    While Printer.TextWidth(TXT$) < ANCHOJB
       TXT$ = " " + TXT$ + " "
    Wend
    Printer.Print TXT$
    If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
    Printer.FontSize = TAFUESEL
    On Error GoTo 0
    Exit Sub
    '
99  ERNU$ = TRIM$(Str$(Err.Number))
100 MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") en la Impresión de\"
    MENSA$ = MENSA$ + "este Listado.\  \"
    '
    MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
    MENSA$ = MENSA$ + "   1.- Impresora Apagada o fuera de Línea.\"
    MENSA$ = MENSA$ + "   2.- Falta Papel o Papel Atorado.\"
    MENSA$ = MENSA$ + "   3.- Problemas en Cable de Señal de Impresora.\"
    MENSA$ = MENSA$ + "   4.- Problemas de Red - No hay Acceso a Impre-\"
    MENSA$ = MENSA$ + "        sora Remota.\"
    MENSA$ = MENSA$ + "   5.- Impresora de Matriz de Puntos Configurada\"
    MENSA$ = MENSA$ + "        para 'Calidad de Imprenta'. Corrija la Con-\"
    MENSA$ = MENSA$ + "        figuración para 'Borrador'.\"
    MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
    '
    C$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(C$, 53, "Error de Impresión de Listado")
    If RESPU% = 4 Then
        Resume 10
    End If
    Call CIERRARCH("*.*")
    Call FINAL("")
    Close: End
    '
End Sub
'


Sub VERDATLIS()
    '
    DATLISTOK% = 0
    If ULTREG&(ARCHITRA$) < 1 Then
        Call MENSERR(24, "Archivo de Trabajo Vacío")
        DATLISTOK% = (-1)
        Exit Sub
    End If
    '
    If SALEXCEL% < 1 Then
      ASALI$ = TRIM$(UCase$(WLISTAB.Text2.Text))
      If ASALI$ = "" Then
         DATLISTOK% = 1
         Exit Sub
      End If
    End If
    '
    If SALEXCEL% = 1 Then
      ASAL% = FILEOPEN%(LUCOM$ + "TREXCEL.TRF", 2, 0)
      DATLISTOK% = 1
      Exit Sub
    End If
    '
    'PPUN% = InStr(ASALI$, ".")
    'If PPUN% < 1 Then PPUN% = 1 + Len(ASALI$)
    'ASALI$ = Left$(ASALI$, PPUN% - 1) + ".TXT"
    '
    'ARCHISAL$ = ASALI$
    If Left$(ARCHISAL$, 3) = "LPT" Then ARCHISAL$ = ""
    'WLISTAB.Text2.Text = ARCHISAL$
    '
10  On Error GoTo 100
    ASAL% = FreeFile
    If SOBRESCRI% = 1 Then
        ASAL% = FILEOPEN%(ARCHISAL$, 2, 0)
      Else
        ASAL% = FILEOPEN%(ARCHISAL$, 3, 0)
    End If
    DATLISTOK% = 1
    Exit Sub
    '
100 Resume 101
101 Close #ASAL%
    ASAL% = 0
    MsgBox "Error de Acceso al Archivo" + Chr$(10) + Chr$(13) + ARCHISAL$
    DATLISTOK% = 0
    ARCHISAL$ = LUDAT$ + TRIM$(ARCHITRA$) + ".TXT"
    WLISTAB.Text2.Text = ARCHISAL$

    '
End Sub

Sub SORTABU()
    '
     WSORTAB.Timer1.Enabled = False
     WSORTAB.Refresh
     Screen.MousePointer = 11
     FIN& = ULTREG&(ARCHITRA$)
     If FIN& < 1 Then GoTo 199
     '
     FINDEX% = FILEINDEX%(ARCHITRA$)
     FINUME% = FILENBR%(ARCHITRA$)
     LOREGI% = LOREGIS%(FINDEX%)
     '
     WSORTAB.LISORT.Clear
     For RECNO& = 1 To FIN&
        '
        Call TRACE(20, RECNO&, FIN&)
        'XXX$ = REGLEIDO$(ARCHITRA$, RECNO&)
        XXX$ = LEEREGIS$(FINDEX%, FINUME%, LOREGI%, RECNO&)
        '
        ARGU$ = "": LARGU% = Len(ARGU$)
        For J% = CALIMI% To CALIMS%
           XXZ$ = Mid$(XXX$, PR%(J%), LF%(J%))
           TV1% = TVL%(J%): If TV1% = 5 Or TV1% = 8 Then TV1% = 1
           LF1% = LF%(J%)
           AC1% = LF%(J%)
           JT1% = 1: LD1% = LDL%(J%)
           '
           If TV1% = 1 Or TV1% = 2 Then
                  JT1% = 2: LD1% = 0
                  If LF1% < 2 Then AC1% = 3 Else AC1% = 5
               ElseIf TV1% = 3 Or TV1% = 4 Then
                  JT1% = 2
                  If LF1% < 8 Then AC1% = 7 Else AC1% = 12
                  If LD1% > 0 Then AC1% = AC1% + 2
               ElseIf TV1% = 5 Then
                  AC1% = 8: LD1% = 0
               ElseIf TV1% = 7 Then
                  AC1% = 16: LD1% = 0
           End If
           '
           ARGX$ = CSTRING$(XXZ$, TV1%, AC1%, LD1%, JT1%)
           ARGU$ = ARGU$ + ARGX$
        Next J%
        '
        If LARGU% = 0 Then
            LARGU% = Len(ARGU$)
        End If
        '
        ARGU$ = ARGU$ + CSTRING$(MKS$(RECNO&), 3, 8, 0, 2)
        WSORTAB.LISORT.AddItem ARGU$
        '
    Next RECNO&
    '
    PPX% = LARGU% + 1
    ARCHINTER$ = ARCHISAL$
    If ARCHISAL$ = ARCHITRA$ Then
        ARCHINTER$ = "!TRANSIT"
        Call COPYSTRU(ARCHITRA$, "TRANSIT")
    End If
    '
    JJ& = 0
    FINPRO& = WSORTAB.LISORT.ListCount - 1
    URETRA& = ULTREG&(ARCHITRA$)
    For u& = 0 To WSORTAB.LISORT.ListCount - 1
       Call TRACE(20, u&, FINPRO&)
       RECNO& = Val(Mid$(WSORTAB.LISORT.List(u&), PPX% + 1))
       If RECNO& > 0 Then
          If RECNO& <= URETRA& Then
             'XXX$ = REGLEIDO$(ARCHITRA$, RECNO&)
             XXX$ = LEEREGIS$(FINDEX%, FINUME%, LOREGI%, RECNO&)
             JJ& = JJ& + 1
             Call GRAREG(ARCHINTER$, XXX$, JJ&)
          End If
       End If
    Next u&
    '
    Call SETULTREG(ARCHINTER$, JJ&)
    Call CIERRARCH(ARCHINTER$)
    Call CIERRARCH(ARCHITRA$)
    If ARCHISAL$ = ARCHITRA$ Then
       JJ& = 0
       FINPRO& = ULTREG&(ARCHINTER$)
       For u& = 1 To ULTREG&(ARCHINTER$)
          Call TRACE(20, u&, FINPRO&)
          XXX$ = REGLEIDO$(ARCHINTER$, u&)
          JJ& = JJ& + 1
          Call GRAREG(ARCHITRA$, XXX$, JJ&)
       Next u&
    End If
    Call SETULTREG(ARCHITRA$, JJ&)
    '
199 Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    WSORTAB.Hide
    Call CIERRARCH("*.*")
    Call LIBARCH("*.*")
    Call BLOQEXE("*")
    '
    If Left$(Sali$, 1) = "*" Then
         Call CONECRUN(Sali$, "")
    End If
    Close: End
    '
End Sub
'

Sub PRENTRA()
    '
    'If YAPRE% = 1 Then
    '    Exit Sub
    'End If
    ACAREN = WEDITAB.LITAB.Height / CARELI%
    RENGLOAC = WEDITAB.LITAB.ListIndex - WEDITAB.LITAB.TopIndex + 1
    If RENGLOAC < 1 Or RENGLOAC > CARELI% Then
        Exit Sub
    End If
    YAPRE% = 1
    BORDESUPE = WEDITAB.Top + WEDITAB.LITAB.Top + RENGLOAC * ACAREN
    KPLA% = 0: PRICA% = 0
    '
    BLANCO = &H80000005
    For II% = 0 To 31
        '
        FORMEDIT.ENTRA(II%).FontSize = 7
        FORMEDIT.ENTRA(II%).FontBold = True
        FORMEDIT.ENTRA(II%).Text = ""
        FORMEDIT.ENTRA(II%).Height = 285
        FORMEDIT.ENTRA(II%).Top = 0
        FORMEDIT.ENTRA(II%).BackColor = BLANCO
        FORMEDIT.ENTRA(II%).Visible = False
        FORMEDIT.BOTSEL(II%).Visible = False
        '
    Next II%
    '
    ALTUCON% = 285
    FORMEDIT.Top = BORDESUPE
    FORMEDIT.Left = WEDITAB.Left
    FORMEDIT.Width = WEDITAB.Width
    FORMEDIT.Height = ALTUCON% * 1.2
    '
    If MODEDI% = 5 Then
        ' FORMEDIT.Top = WEDITAB.Top + WEDITAB.FRAMETOT.Top
        FORMEDIT.Top = WEDITAB.Top + WEDITAB.LITAB.Top + (CARELI% + 1) * ACAREN + (WEDITAB.FRAMETOT.Height - FORMEDIT.Height) / 2
    End If
    '
    FORMEDIT.Command1.Left = WEDITAB.FRAMEBOT2.Left
    FORMEDIT.Command1.Width = WEDITAB.BOTAGRE.Width * 0.6
    FORMEDIT.Command2.Left = FORMEDIT.Command1.Left + FORMEDIT.Command1.Width
    FORMEDIT.Command2.Width = WEDITAB.BOTAGRE.Width * 0.4
    FORMEDIT.Command1.Height = ALTUCON%
    FORMEDIT.Command2.Height = ALTUCON%
    '
    POSEDI% = 0
    For II% = 1 To NCL%
       '
       If CL%(II%) > 0 Then
          FORMEDIT.ENTRA(II% - 1).Left = COEFH * (CL%(II%) - 0.9)
          FORMEDIT.ENTRA(II% - 1).Top = 0
          FORMEDIT.ENTRA(II% - 1).Width = COEFH * (ACL%(II%) + 0.9)
          FORMEDIT.ENTRA(II% - 1).Visible = True
          FORMEDIT.ENTRA(II% - 1).FontName = WEDITAB.LITAB.FontName
          FORMEDIT.ENTRA(II% - 1).FontSize = WEDITAB.LITAB.FontSize
          FORMEDIT.ENTRA(II% - 1).FontBold = WEDITAB.LITAB.FontBold
          If PE%(II%) <> 1 Then
            If PE%(II%) = 3 Or MODEDI% = 4 Then
              FORMEDIT.ENTRA(II% - 1).BackColor = &H80000013
            End If
          End If
       End If
       '
       If FORMEDIT.ENTRA(II% - 1).BackColor = BLANCO Then
           POSEDI% = 1
       End If
       '
    Next II%
    '
    If MODEDI% = 4 Then                     ' editar registro
        RELIB$ = WEDITAB.LITAB.Text
        XXX$ = REGLEIDO(ARCHILOC$, WEDITAB.LITAB.ListIndex + 1)
        For II% = 1 To NCL%
           If CL%(II%) > 0 Then
              FORMEDIT.ENTRA(II% - 1).Text = Mid$(RELIB$, CL%(II%), ACL%(II%))
              If TVL%(CFT%(II%)) <= 4 Then
                  If Left$(TRIM$(FORMEDIT.ENTRA(II% - 1)), 1) = "*" Then
                      XXZ$ = Mid$(XXX$, PR%(CFT%(II%)), LF%(CFT%(II%)))
                      FORMEDIT.ENTRA(II% - 1).Text = TRIM$(CSTRING$(XXZ$, TVL%(CFT%(II%)), 15, LDL%(CFT%(II%)), 1))
                  End If
              End If
           End If
        Next II%
    End If
    '
    If POSEDI% = 1 Or MODEDI% < 4 Then
        FORMEDIT.Refresh
        On Error Resume Next
        COLUAC% = 0
        FORMEDIT.Show 1
        On Error GoTo 0
      Else
        WEDITAB.BOTEDIT.Enabled = False
    End If
    '
End Sub
'

Sub ADDLISTA()
    '
    If NFT% > 0 Then
       If YAFIL% < 1 Then
          Call ENTRAFIL
          Exit Sub
       End If
    End If
    '
    WEDITAB.LITAB.AddItem Space$(ATOT%)
    Call VERISCROLL
    WEDITAB.LITAB.ListIndex = WEDITAB.LITAB.ListCount - 1
    If WEDITAB.LITAB.TopIndex < WEDITAB.LITAB.ListIndex - CARELI% Then
        WEDITAB.LITAB.TopIndex = WEDITAB.LITAB.ListCount - CARELI% - 3
    End If
    MODEDI% = 1
    NUREGIS& = WEDITAB.LITAB.ListCount + 1
    REACTIVO$ = REGBLAN$(ARCHILOC$)
    Call PRENTRA
    '
End Sub
'

Sub INSLISTA()
    '
    If NFT% > 0 Then
       If YAFIL% < 1 Then
          Exit Sub
       End If
    End If
    '
    ILISTA& = WEDITAB.LITAB.ListIndex
    If ILISTA& < 0 Or ILISTA& > WEDITAB.LITAB.ListCount - 1 Then
        ILISTA& = 0
    End If
    '
    WEDITAB.LITAB.AddItem Space$(ATOT%), ILISTA&
    Call VERISCROLL
    WEDITAB.LITAB.ListIndex = ILISTA&
    If WEDITAB.LITAB.TopIndex < WEDITAB.LITAB.ListIndex - CARELI% Then
        WEDITAB.LITAB.TopIndex = WEDITAB.LITAB.ListIndex - CARELI% / 2
    End If
    '
    For u& = ULTREG&(ARCHILOC$) To ILISTA& + 1 Step -1
       REACTIVO$ = REGLEIDO$(ARCHILOC$, u&)
       Call GRAREG(ARCHILOC$, REACTIVO$, u& + 1)
    Next u&
    REACTIVO$ = REGBLAN$(ARCHILOC$)
    Call GRAREG(ARCHILOC$, REACTIVO$, ILISTA& + 1)
    '
    MODEDI% = 2
    NUREGIS& = ILISTA& + 1
    REACTIVO$ = REGBLAN$(ARCHILOC$)
    Call PRENTRA
    '
End Sub
'
Sub BORLISTA()
    '
    If NFT% > 0 Then
       If YAFIL% < 1 Then
          Exit Sub
       End If
    End If
    '
    ILISTA& = WEDITAB.LITAB.ListIndex
    If ILISTA& < 0 Or ILISTA& > WEDITAB.LITAB.ListCount - 1 Then
        Exit Sub
    End If
    '
    WEDITAB.LITAB.RemoveItem ILISTA&
    Call VERISCROLL
    For u& = ILISTA& + 2 To ULTREG&(ARCHILOC$)
       REACTIVO$ = REGLEIDO$(ARCHILOC$, u&)
       Call GRAREG(ARCHILOC$, REACTIVO$, u& - 1)
    Next u&
    REACTIVO$ = String$(RECLEN%(ARCHILOC$), 0)
    Call GRAREG(ARCHILOC$, REACTIVO$, ULTREG&(ARCHILOC$))
    Call CIERRARCH(ARCHILOC$)
    '
    If ILISTA& > WEDITAB.LITAB.ListCount - 1 Then
       ILISTA& = WEDITAB.LITAB.ListCount - 1
    End If
    WEDITAB.LITAB.ListIndex = ILISTA&
    '
    NUREGIS& = ILISTA& + 1
    REACTIVO$ = REGLEIDO$(ARCHILOC$, NUREGIS&)
    '
End Sub
'

Sub EDLISTA()
    '
    If WEDITAB.LITAB.ListIndex < 0 Then
        Exit Sub
    End If
    '
    If WEDITAB.LITAB.ListIndex >= WEDITAB.LITAB.ListCount Then
        Exit Sub
    End If
    '
    MODEDI% = 4
    NUREGIS& = WEDITAB.LITAB.ListIndex + 1
    REACTIVO$ = REGLEIDO$(ARCHILOC$, NUREGIS&)
    '
    Call PRENTRA
    '
End Sub
'
Sub BULISTA()
    '
    Call OPNOIN("")
    Exit Sub
    '
    If WEDITAB.LITAB.ListCount < 1 Then
        Exit Sub
    End If
    MODEDI% = 5
    Call PRENTRA
    '
End Sub
'

Sub CARDATIMPRE()
    '
    NOMIMPRE$ = TRIM$(Printer.DeviceName)
    Printer.Orientation = 1
    Printer.ScaleMode = 1
    ANCHOPAP = Int(Printer.ScaleWidth / 56.7)
    LARGOPAP = Int(Printer.ScaleHeight / 56.7)
    MARGENI = 0
    MARGENS = 0
    ORIENTA% = 1
    CACOPIAS% = 1
    '
    Dim PPZZ As String * 128
    NX% = FILEOPEN(LUCOM$ + "LISTSET.DRV", 0, 128)
    FIN& = LOF(NX%) / 128
    For IPRIN& = 1 To FIN&
         Get #NX%, IPRIN&, PPZZ$
         PPXX$ = PPZZ$
         If TRIM$(Left$(PPXX$, 48)) = NOMIMPRE$ Then
           If TRIM$(UCase$(Mid$(PPXX$, 57, 8))) = TRIM$(UCase$(PROPRIN$)) Then
             ANCHOPAP1 = CVS(Mid$(PPXX$, 49, 4))
             LARGOPAP1 = CVS(Mid$(PPXX$, 53, 4))
             '
             If ANCHOPAP1 > 0 Then
                If ANCHOPAP1 < ANCHOPAP Then
                   ANCHOPAP = ANCHOPAP1
                End If
             End If
             '
             If LARGOPAP1 > 0 Then
                If LARGOPAP1 < LARGOPAP Then
                   LARGOPAP = LARGOPAP1
                End If
             End If
             '
             GoTo 10
             '
           End If
         End If
    Next IPRIN&
    '
10  Close #NX%
    NX% = FILEOPEN%(LUCOM$ + "LISTABU.DRV", 0, 128)
    FIN& = LOF(NX%) / 128
    For IPRIN& = 1 To FIN&
         Get #NX%, IPRIN&, PPZZ$
         PPXX$ = PPZZ$
         If TRIM$(Left$(PPXX$, 8)) = TRIM$(PROPRIN$) Then
             MARGENI1 = CVS(Mid$(PPXX$, 9, 4))
             MARGENS1 = CVS(Mid$(PPXX$, 13, 4))
             ORIENTA1% = Asc(Mid$(PPXX$, 17, 1))
             CACOPIAS1% = Asc(Mid$(PPXX$, 18, 1))
             '
             If MARGENI1 > MARGENI Then
                 MARGENI = MARGENI1
             End If
             '
             If MARGENS1 > MARGENS Then
                   MARGENS = MARGENS1
             End If
             '
             If ORIENTA1% > 0 Then
                If ORIENTA1% <= 2 Then
                    ORIENTA% = ORIENTA1%
                End If
             End If
             '
             If CACOPIAS1% > 0 Then
                If CACOPIAS1% <= 32 Then
                    CACOPIAS% = CACOPIAS1%
                End If
             End If
             '
             GoTo 20
             '
         End If
    Next IPRIN&
    '
20  If ORIENTA% = 1 Then
         WLISTAB.Option2(0).Value = True        ' RETRATO
       Else
         WLISTAB.Option2(1).Value = True        ' PAISAJE
    End If
    WLISTAB.Text4(0).Text = CSTRING$(MKS$(ANCHOPAP), 3, 6, 1, 1)
    WLISTAB.Text4(1).Text = CSTRING$(MKS$(LARGOPAP), 3, 6, 1, 1)
    WLISTAB.Text4(2).Text = CSTRING$(MKS$(MARGENS), 3, 6, 1, 1)
    WLISTAB.Text4(3).Text = CSTRING$(MKS$(MARGENI), 3, 6, 1, 1)
    WLISTAB.Text6.Text = CACOPIAS%
    '
End Sub


Sub GRADATIMPRE()
    '
5   On Error GoTo 199
    NOMIMPRE$ = TRIM$(Printer.DeviceName)
    Printer.Orientation = 1
    Printer.ScaleMode = 1
    ANCHOPAP = Int(Printer.ScaleWidth / 56.7)
    LARGOPAP = Int(Printer.ScaleHeight / 56.7)
    On Error GoTo 0
    '
    If Val(WLISTAB.Text4(0).Text) > 0 Then
        If Val(WLISTAB.Text4(0).Text) < ANCHOPAP Then
            ANCHOPAP = Val(WLISTAB.Text4(0).Text)
        End If
    End If
    If Val(WLISTAB.Text4(1).Text) > 0 Then
        If Val(WLISTAB.Text4(1).Text) < LARGOPAP Then
            LARGOPAP = Val(WLISTAB.Text4(1).Text)
        End If
    End If
    MARGENI = Val(WLISTAB.Text4(3).Text)
    MARGENS = Val(WLISTAB.Text4(2).Text)
    ORIENTA% = 1
    If WLISTAB.Option2(1).Value = True Then
        ORIENTA% = 2
    End If
    CACOPIAS% = 1
    If Val(WLISTAB.Text6.Text) > 0 Then
        If Val(WLISTAB.Text6.Text) <= 32 Then
            CACOPIAS% = Val(WLISTAB.Text6.Text)
        End If
    End If
    '
    Dim PPZZ As String * 128
    NX% = FILEOPEN(LUCOM$ + "LISTSET.DRV", 0, 128)
    FIN& = LOF(NX%) / 128
    For IPRIN& = 1 To FIN&
         Get #NX%, IPRIN&, PPZZ$
         PPXX$ = PPZZ$
         If TRIM$(Left$(PPXX$, 48)) = NOMIMPRE$ Then
           If TRIM$(UCase$(Mid$(PPXX$, 57, 8))) = TRIM$(UCase$(PROPRIN$)) Then
             Call REPLA(PPXX$, MKS$(ANCHOPAP), 49, 4)
             Call REPLA(PPXX$, MKS$(LARGOPAP), 53, 4)
             PPZZ$ = PPXX$
             Put #NX%, IPRIN&, PPZZ$
             GoTo 9
           End If
         End If
    Next IPRIN&
    PPXX$ = Space$(48) + String$(80, 0)
    Call REPLA(PPXX$, NOMIMPRE$, 1, 48)
    Call REPLA(PPXX$, MKS$(ANCHOPAP), 49, 4)
    Call REPLA(PPXX$, MKS$(LARGOPAP), 53, 4)
    Call REPLA(PPXX$, PROPRIN$, 57, 8)
    PPZZ$ = PPXX$
    Put #NX%, FIN& + 1, PPZZ$
9   Close #NX%
    '
10  NX% = FILEOPEN%(LUCOM$ + "LISTABU.DRV", 0, 128)
    FIN& = LOF(NX%) / 128
    For IPRIN& = 1 To FIN&
         Get #NX%, IPRIN&, PPZZ$
         PPXX$ = PPZZ$
         If TRIM$(Left$(PPXX$, 8)) = TRIM$(PROPRIN$) Then
             GoTo 20
         End If
    Next IPRIN&
    '
    IPRIN& = FIN& + 1
    PPXX$ = String$(128, 0)
    '
20  Call REPLA(PPXX$, PROPRIN$, 1, 8)
    Call REPLA(PPXX$, MKS$(MARGENI), 9, 4)
    Call REPLA(PPXX$, MKS$(MARGENS), 13, 4)
    Call REPLA(PPXX$, Chr$(ORIENTA%), 17, 1)
    Call REPLA(PPXX$, Chr$(CACOPIAS%), 18, 1)
    PPZZ$ = PPXX$
    Put #NX%, IPRIN&, PPZZ$
    Close #NX%
    Exit Sub
    '
199 ERNU$ = TRIM$(Str$(Err.Number))
    MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") en la Impresión de\"
    MENSA$ = MENSA$ + "este Listado.\  \"
    '
    MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
    MENSA$ = MENSA$ + "   1.- Impresora Apagada o fuera de Línea.\"
    MENSA$ = MENSA$ + "   2.- Falta Papel o Papel Atorado.\"
    MENSA$ = MENSA$ + "   3.- Problemas en Cable de Señal de Impresora.\"
    MENSA$ = MENSA$ + "   4.- Problemas de Red - No hay Acceso a Impre-\"
    MENSA$ = MENSA$ + "        sora Remota.\"
    MENSA$ = MENSA$ + "   5.- Impresora de Matriz de Puntos Configurada\"
    MENSA$ = MENSA$ + "        para 'Calidad de Imprenta'. Corrija la Con-\"
    MENSA$ = MENSA$ + "        figuración para 'Borrador'.\"
    MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
    '
    C$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(C$, 53, "Error de Impresión de Listado")
    If RESPU% = 4 Then
        Resume 5
    End If
    Call CIERRARCH("*.*")
    Call FINAL("")
    Close: End
    '
End Sub
'

Sub PRINTABU()
   '
   NOMIMPRE$ = TRIM$(Printer.DeviceName)
   Printer.Orientation = 1
   Printer.ScaleMode = 1
   Printer.FontBold = False
   Printer.FontItalic = False
   Printer.ColorMode = 1     ' blanco y negro
   On Error Resume Next
   Printer.ForeColor = QBColor(0)
   Printer.DrawMode = 13     ' linea negra
   Printer.FillStyle = 1
   Printer.FillColor = QBColor(15)
   Printer.FontTransparent = True
   On Error GoTo 0
   '
   MARSEGUR = Int((Printer.Width - Printer.ScaleWidth) / 56.7 + 1)
   ANCHOPAP = Int(Printer.ScaleWidth / 56.7) - MARSEGUR
   If ANCHOPAP > 196 Then ANCHOPAP = 196
   LARGOPAP = Int(Printer.ScaleHeight / 56.7)
   MARGENI = 0: MARGENS = 0
   '
   ENCA$ = ""
   For KCL% = 1 To NCL%
      If Left$(ECL$(KCL%), 1) <> "@" Then
         If Left$(ECL$(KCL%), 1) = "<" Then
              ENCACO$ = AJUSTI$(Mid$(ECL$(KCL%), 2), LC%(KCL%))
              ENCACO$ = Space$(MG%(KCL%)) + ENCACO$
              ENCACO$ = Left$(ENCACO$, LC%(KCL%))
           ElseIf Left$(ECL$(KCL%), 1) = ">" Then
              ENCACO$ = AJUSTD$(Mid$(ECL$(KCL%), 2), LC%(KCL%))
              ENCACO$ = Mid$(ENCACO$, MG%(KCL%) + 1) + Space$(MG%(KCL%))
           Else
              ENCACO$ = TRIM$(ECL$(KCL%))
              While Len(ENCACO$) < LC%(KCL%)
                 ENCACO$ = " " + ENCACO$ + " "
              Wend
              ENCACO$ = Left$(ENCACO$, LC%(KCL%))
         End If
         ENCA$ = ENCA$ + ENCACO$
         If KCL% < NCL% Then
            ENCA$ = ENCA$ + " "
         End If
      End If
   Next KCL%
   ATOT% = 2 + Len(ENCA$)
   '
10 YAESTA% = 0: CAFUFIJ% = 0
   NX% = FILEOPEN(LUCOM$ + "PRINTERS.DRV", 0, 128)
   For IPRIN& = 1 To LOF(NX%) / 128
      Get #NX%, IPRIN&, PPZZ$
      If TRIM$(Left$(PPZZ$, 48)) = NOMIMPRE$ Then
         YAESTA% = 1
         If UCase$(Mid$(PPZZ$, 125, 2)) = "SI" Then
            CAFUFIJ% = CAFUFIJ% + 1
            FUFIJA$(CAFUFIJ%) = PPZZ$
         End If
      End If
   Next IPRIN&
   '
   Close #NX%
   NOMIMP$ = REPLACAR$(NOMIMPRE$, "\", "/")
   If YAESTA% < 1 Then
      OPX% = COMALTER%("Se ha Detectado un Modelo de Impresora\Desconocido para el Sistema Aplicativo\  \'" + NOMIMP$ + "'\  \\Configurar\Cancelar")
      If OPX% <> 1 Then Exit Sub
      Call CIERRARCH("*.*")
      WEDITAB.Hide
      Call SETUPRINTER
      GoTo 10
   End If
   '
   If CAFUFIJ% < 1 Then
      OPX% = COMALTER%("No hay Fuentes Utilizables para Impresora\  \'" + NOMIMP$ + "'\  \\Configurar\Cancelar")
      If OPX% <> 1 Then Exit Sub
      Call CIERRARCH("*.*")
      WEDITAB.Hide
      Call SETUPRINTER
      GoTo 10
   End If
   '
   ANCHOJA = 56.7 * (ANCHOPAP - 16)
   ATEXMAX = 0
   FUENTESEL$ = ""
   TAFUESEL = 0
   TALEBASI = 24
   '
20 TEPRUEBA$ = String$(ATOT%, "*")
   For I% = 1 To CAFUFIJ%
      Printer.FontName = TRIM$(Mid$(FUFIJA$(I%), 65, 48))
      If Printer.FontName <> TRIM$(Mid$(FUFIJA$(I%), 65, 48)) Then
         Call COMUNI("Se ha Detectado un Error\en la Configuración Local\de la Impresora Predeterminada.\  \Imposible Emitir el Listado Solicitado.")
         Call CIERRARCH("*.*")
         Close
         Call FILEKILL(LUCOM$ + "PRINTERS.DRV")
         Call FINAL("")
      End If
      '
      FUENTEMIN = CVS(Mid$(FUFIJA$(I%), 113, 4))
      FUENTEMAX = CVS(Mid$(FUFIJA$(I%), 117, 4))
      '
      TALEBA = FUENTEMAX + 0.4
      While TALEBA > FUENTEMIN
         TALEBA = TALEBA - 0.4
         Printer.FontSize = TALEBA
         If Printer.FontSize > TALEBA Then GoTo 25
         ATEXTO = 0
         On Error Resume Next
         ATEXTO = Printer.TextWidth(TEPRUEBA$)
         On Error GoTo 0
         If ATEXTO < 1 Then GoTo 25
         If ATEXTO <= ANCHOJA Then
            If ATEXTO > ATEXMAX Then
               FUENTESEL$ = Printer.FontName
               TAFUESEL = Printer.FontSize ' TALEBASI
               ATEXMAX = ATEXTO
            End If
            GoTo 29
         End If
         TALEBA = Printer.FontSize
25    Wend
      '
29 Next I%
   '
   ATEXTO = ATEXMAX
   If TRIM$(FUENTESEL$) = "" Or TAFUESEL < 1 Then
      Call MENSERR(24, "No hay Fuentes Adecuadas\para la Impresion de Este Listado")
      Exit Sub
   End If
   '
   Printer.FontName = FUENTESEL$
   Printer.FontSize = TAFUESEL
   HREN = Printer.TextHeight(TEPRUEBA$)
   MARGENI = 16 * 56.7 + ANCHOJA - ATEXTO: If MARGENI < 0 Then MARGENI = 0
   '
   MODOSEL% = 0
   CARESEL& = WEDITAB.LITAB.SelCount
   If CARESEL& > 0 Then
      MODOSEL% = COMALTER%("Seleccione Rango de Impresión\\Todo\Texto Seleccionado\Cancelar") - 1
      If MODOSEL% < 0 Or MODOSEL% > 1 Then Exit Sub
   End If
   If MODOSEL% = 0 Then CARESEL& = WEDITAB.LITAB.ListCount
   '
   NUHOJ% = 0: CAHO% = 0
   Printer.Orientation = 1
   GoSub IMPREFORM
   '
   On Error Resume Next
   Printer.ForeColor = QBColor(0)
   Printer.FillStyle = 1
   Printer.FillColor = QBColor(15)
   Printer.FontTransparent = True
   On Error GoTo 0
   '
   For u& = 1 To WEDITAB.LITAB.ListCount
      If MODOSEL% = 0 Or WEDITAB.LITAB.Selected(u& - 1) = True Then
35       POVERT = POVERT + HREN
         Printer.FontName = FUENTESEL$
         Printer.FontSize = TAFUESEL
         If POVERT > (LARGOPAP - 25) * 56.7 - 3.2 * HREN Then
            GoSub LEPIEPA
            Printer.NewPage
            '
            GoSub IMPREFORM
            GoTo 35
         End If
         '
         Printer.CurrentX = MARGENI + 100
         Printer.CurrentY = POVERT
         Printer.Print RTrim$(Mid$(WEDITAB.LITAB.List(u& - 1), 2, ATOT% - 1))
      End If
   Next u&
   '
   If MODOSEL% = 0 Then
      Printer.CurrentX = MARGENI + 100
      Printer.CurrentY = (LARGOPAP - 25) * 56.7 - 1.5 * HREN
      Printer.Print RTrim$(Mid$(WEDITAB.TOTALI.Caption, 2, ATOT% - 1))
   End If
   '
   GoSub LEPIEPA
   Printer.EndDoc
   Exit Sub

IMPREFORM:
   '
   Printer.ScaleMode = 1
   Printer.FontBold = False
   Printer.FontItalic = False
   '
   Printer.FontName = FUENTESEL$
   Printer.FontSize = TAFUESEL
   '
   On Error Resume Next
   Printer.ForeColor = QBColor(0)
   Printer.DrawMode = 13    ' linea negra
   Printer.DrawWidth = 2
   Printer.DrawStyle = 0
   Printer.FillStyle = 1
   Printer.FillColor = QBColor(15)
   Printer.FontTransparent = True
   On Error GoTo 0
   '
   Printer.CurrentX = 0
   Printer.CurrentY = 0
   Printer.Print "    "       ' para inicializar impresora
   '
   H1 = MARGENI: V1 = HREN
   H2 = H1 + ATEXTO: V2 = 3 * HREN
   Printer.Line (H1, V1)-(H2, V2), QBColor(0), B
   '
   Printer.CurrentX = MARGENI + 100
   Printer.CurrentY = 1.7 * HREN
   Printer.Print S1$
   '
   VFILMAX = 0
   COEF = ATEXTO / (WEDITAB.LITAB.Width - 240)
   For II% = 1 To NFT%
      If WEDITAB.FILABEL(II% - 1).Visible = True Then
         V2 = 4 * HREN + WEDITAB.FIENTRA(II% - 1).Top + HREN * 1.3
         If V2 > VFILMAX Then VFILMAX = V2
      End If
   Next II%
   '
   Printer.DrawWidth = 2
   TOPETAB = 4 * HREN
   If NFT% > 0 Then
      H1 = MARGENI: V1 = 4 * HREN
      H2 = H1 + ATEXTO: V2 = VFILMAX + 0.5 * HREN
      TOPETAB = V2 + HREN
      Printer.Line (H1, V1)-(H2, V2), QBColor(0), B
   End If
   '
   Printer.DrawWidth = 1
   For II% = 1 To NFT%
      '
      If WEDITAB.FILABEL(II% - 1).Visible = True Then
         '
         TTXX$ = WEDITAB.FILABEL(II% - 1).Caption
         H1 = COEF * WEDITAB.FILABEL(II% - 1).Left + MARGENI
         V1 = 4 * HREN + WEDITAB.FILABEL(II% - 1).Top
         Printer.CurrentX = H1
         Printer.CurrentY = V1
         Printer.Print TTXX$
         '
         H1 = COEF * WEDITAB.FIENTRA(II% - 1).Left + MARGENI
         H2 = H1 + COEF * WEDITAB.FIENTRA(II% - 1).Width
         If H2 > MARGENI + ATEXTO - 150 Then H2 = MARGENI + ATEXTO - 150
         ANCHOCUA = H2 - H1 - 100
         V1 = 4 * HREN + WEDITAB.FIENTRA(II% - 1).Top
         V2 = V1 + HREN * 1.3
         If V2 > VFILMAX Then VFILMAX = V2
         Printer.Line (H1, V1)-(H2, V2), QBColor(0), B
         '
         TTXX$ = WEDITAB.FIENTRA(II% - 1).Text
34       If Printer.TextWidth(TTXX$) > ANCHOCUA Then
            If Len(TTXX$) > 0 Then
               TTXX$ = Left$(TTXX$, Len(TTXX$) - 1)
               GoTo 34
            End If
         End If
         Printer.CurrentX = H1 + 50
         Printer.CurrentY = V1 + 0.2 * HREN
         Printer.Print TTXX$
         '
         If WEDITAB.FIECO(II% - 1).Visible = True Then
            H1 = COEF * WEDITAB.FIECO(II% - 1).Left + MARGENI
            H2 = H1 + COEF * WEDITAB.FIECO(II% - 1).Width
            If H2 > MARGENI + ATEXTO - 150 Then H2 = MARGENI + ATEXTO - 150
            ANCHOCUA = H2 - H1 - 100
            V1 = 4 * HREN + WEDITAB.FIECO(II% - 1).Top
            V2 = V1 + HREN * 1.3
            If V2 > VFILMAX Then VFILMAX = V2
            Printer.Line (H1, V1)-(H2, V2), QBColor(0), B
            '
            TTXX$ = WEDITAB.FIECO(II% - 1).Text
36          If Printer.TextWidth(TTXX$) > ANCHOCUA Then
               If Len(TTXX$) > 0 Then
                  TTXX$ = Left$(TTXX$, Len(TTXX$) - 1)
                  GoTo 36
               End If
            End If
            Printer.CurrentX = H1 + 50
            Printer.CurrentY = V1 + 0.2 * HREN
            Printer.Print TTXX$
         End If
         '
      End If
   Next II%
   '
   H1 = MARGENI: V1 = TOPETAB
   H2 = H1 + ATEXTO: V2 = (LARGOPAP - 25) * 56.7
   Printer.Line (H1, V1)-(H2, V2), QBColor(0), B
   '
   H1 = MARGENI: V1 = TOPETAB + 2 * HREN
   H2 = H1 + ATEXTO: V2 = V1
   Printer.Line (H1, V1)-(H2, V2), QBColor(0)
   POVERT = V2 - 0.5 * HREN
   '
   Printer.CurrentX = MARGENI + 100
   Printer.CurrentY = TOPETAB + 0.7 * HREN
   TTXX$ = ENCA$
   Printer.Print TTXX$
   '
   H1 = MARGENI: V1 = (LARGOPAP - 25) * 56.7 - 2 * HREN
   H2 = H1 + ATEXTO: V2 = V1
   Printer.Line (H1, V1)-(H2, V2), QBColor(0)
   '
   If CAHO% < 1 Then
      CAREHO% = (V1 - TOPETAB) / HREN
      CAHOJAS = CARESEL& / CAREHO%
      If CAHOJAS > Int(CAHOJAS) Then CAHOJAS = CAHOJAS + 1
      CAHO% = Int(CAHOJAS)
   End If
   '
   NUHOJ% = NUHOJ% + 1
   TTXX$ = "Hoja No. " + TRIM$(Str$(NUHOJ%)) + "/" + TRIM$(Str$(CAHO%))
   Printer.CurrentX = MARGENI + ATEXTO - 150 - Printer.TextWidth(TTXX$)
   Printer.CurrentY = 1.7 * HREN
   Printer.Print TTXX$
   H1 = MARGENI + ATEXTO - 300 - Printer.TextWidth(TTXX$)
   H2 = H1
   V1 = HREN: V2 = 3 * HREN
   Printer.DrawWidth = 2
   Printer.Line (H1, V1)-(H2, V2), QBColor(0)
   '
   H3 = H1
   HOII% = HOY(HOS$)
   TTXX$ = HOS$
   Printer.CurrentX = H3 - 150 - Printer.TextWidth(TTXX$)
   Printer.CurrentY = 1.7 * HREN
   Printer.Print TTXX$
   H1 = H3 - 300 - Printer.TextWidth(TTXX$)
   H2 = H1
   V1 = HREN: V2 = 3 * HREN
   Printer.DrawWidth = 2
   Printer.Line (H1, V1)-(H2, V2), QBColor(0)
   '
Return
'
LEPIEPA:
   '
   Printer.FontSize = 6
   TXT$ = "Sistema FLEXOFT s/n " + SERINUM$
   TXT$ = TXT$ + " licenciado a " + EMPRELI$
   TXT$ = TXT$ + " - Derechos Reservados L.11723"
   TXT$ = MIZQ$ + Space$(OFS%) + TXT$
   '
   While Printer.TextWidth(TXT$) < ATEXTO
      TXT$ = " " + TXT$ + " "
   Wend
   Printer.CurrentX = MARGENI
   Printer.CurrentY = (LARGOPAP - 25) * 56.7 + 100
   Printer.Print TXT$
   Printer.FontSize = TAFUESEL
   '
Return
'
End Sub

Sub LISEXCEL()
     '
     Screen.MousePointer = 11
     '
     Print #ASAL%, ARCHISAL$
     '
     For K2% = 1 To NCL%
       '
       TPAR$ = ECL$(K2%)
       If JT%(K2%) = 1 Then
           TPAR$ = Space$(MG%(K2%)) + TPAR$
         ElseIf JT%(K2%) = 2 Then
           TPAR$ = TPAR$ + Space$(MG%(K2%))
       End If
       '
       Print #ASAL%, "#T#" + TPAR$
     Next K2%
     Print #ASAL%, "@NEWLINE@"
     '
     For K2% = 1 To NCL%
       '
       TPAR$ = String$(Len(ECL$(K2%)), "-")
       If JT%(K2%) = 1 Then
           TPAR$ = Space$(MG%(K2%)) + TPAR$
         ElseIf JT%(K2%) = 2 Then
           TPAR$ = TPAR$ + Space$(MG%(K2%))
       End If
       '
       Print #ASAL%, "#T#" + TPAR$
     Next K2%
     Print #ASAL%, "@NEWLINE@"
     Print #ASAL%, "@NEWLINE@"
     '
     FINX& = ULTREG&(ARCHITRA$)
     '
     For K1& = HDR& + 1 To FINX&
        Call TRACE(20, K1&, FINX&)
        XXX$ = REGLEIDO$(ARCHITRA$, K1&)
        '
        If XXX$ = String$(RECLEN%(ARCHITRA$), "=") Then
            GoTo 19
        End If
        '
        If XXX$ = REGBLAN$(ARCHITRA$) Then
            GoTo 19
        End If
        '
        If XXX$ = String$(RECLEN%(ARCHITRA$), "-") Then
            GoTo 19
        End If
        '
        If XXX$ = String$(RECLEN%(ARCHITRA$), "*") Then
            GoTo 19
        End If
        '
        KI% = 0: KF% = 0
        For K2% = 1 To NCL%
            If PE%(K2%) < 3 Then
               If CFT%(K2%) = 1 Then
                     XXZ$ = MKS$(K1& + HDR&)
                     If NFT% > 0 Then XXZ$ = RECNO$
                   Else
                     XXZ$ = Mid$(XXX$, PR%(CFT%(K2%)), LF%(CFT%(K2%)))
               End If
               '
               TIVA% = TVL%(CFT%(K2%))
               If TIVA% = 6 Then
                 If Left$(XXZ$, 1) = "@" Then
                   For K3% = 1 To NCL%
                     VFS$(K3%) = Space$(ACL%(K3%))
                   Next K3%
                   VFS$(K2%) = "#T#" + Mid$(XXZ$, 2)
                   GoTo 4
                 End If
               End If
               '
               LOTOT% = ACL%(K2%)
               LODEC% = LDL%(CFT%(K2%))
               JUSTI% = JT%(K2%)
               '
               If TIVA% = 5 Then
                 VFS$(K2%) = "#F#" + FETEXCOR$(CVI(XXZ$))
                 GoTo 3
               End If
               '
               VFS$(K2%) = CSTRING$(XXZ$, TIVA%, LOTOT%, LODEC%, JUSTI%)
               If TIVA% < 5 Then
                    VFS$(K2%) = "#N#" + VFS$(K2%)
                  ElseIf TIVA% = 6 Or TIVA% = 7 Or TIVA% = 8 Then
                    VFS$(K2%) = "#T#" + VFS$(K2%)
               End If
               '
               If CE%(K2%) > 0 Then
                  VFS$(CE%(K2%)) = "#T#" + AJUSTI$(ECOARCH$(AE$(K2%), XXZ$), ACL%(CE%(K2%)))
               End If
               '
               VALO# = 0: If TIVA% <= 4 Then VALO# = Val(Mid$(VFS$(K2%), 4))
               If TIVA% = 8 Then
                  VALO# = Val(CSTRING$(XXZ$, 4, 12, 2, 1))
                  VFX#(K2%) = VALO#
               End If
               If TT%(K2%) > K2% Then
                  TT1% = TT%(K2%) Mod 100: SIG% = 1
                  If TT%(K2%) > 100 Then SIG% = (-1)
                  SUPAR#(TT1%) = SUPAR#(TT1%) + VALO# * SIG%
               End If
               '
            End If
3       Next K2%
        '
        For K2% = 1 To NCL%
          If TVL%(CFT%(K2%)) < 5 Or TVL%(CFT%(K2%)) = 8 Then
            If TVL%(CFT%(K2%)) <> 8 Then VFX#(K2%) = 0
            LOTOT% = ACL%(K2%)
            LODEC% = LDL%(CFT%(K2%))
            TIVA% = TVL%(CFT%(K2%))
            '
            If TT%(K2%) = 255 Then
                 VFX#(K2%) = SUPAR#(K2%) ' : GoTo 14660
                 VFS$(K2%) = "#N#" + NUMSTRI$(VFX#(K2%), LOTOT%, LODEC%, 0)
               Else
                 If TT%(K2%) > 0 And TT%(K2%) < K2% Then
                    VFX#(K2%) = 0
                    YYY$ = REGLEIDO$(ARCHITRA$, K1& + 1)
                    If Mid$(YYY$, PR%(CFT%(TT%(K2%))), LF%(CFT%(TT%(K2%)))) <> Mid$(XXX$, PR%(CFT%(TT%(K2%))), LF%(CFT%(TT%(K2%)))) Then
                       If KI% < TT%(K2%) Then KI% = TT%(K2%)
                       VFX#(K2%) = SUPAR#(K2%)
                       SUPAR#(K2%) = 0
                       If K2% > KF% Then KF% = K2%
                    End If
                    XXX$ = REGLEIDO$(ARCHITRA$, K1&)
                    VFS$(K2%) = "#N#" + NUMSTRI$(VFX#(K2%), LOTOT%, LODEC%, 0)
                    If TIVA% = 8 Then
                       TIETO = VFX#(K2%)
                       JUSTI% = JT%(K2%)
                       VFS$(K2%) = "#T#" + CSTRING$(MKS$(TIETO), TIVA%, LOTOT%, LODEC%, JUSTI%)
                       If TIETO < 0.01 Then VFS$(K2%) = Space$(LOTOT%)
                    End If
                End If
            End If
            '
            If TT%(K2%) > 0 Then
               For K4% = 1 To 2
                  TIVA% = TVL%(CFT%(K2%))
                  If TIVA% <> 8 Then
                       SUMA#(K2%, K4%) = SUMA#(K2%, K4%) + CDbl(Val(VFS$(K2%)))
                     Else
                       SUMA#(K2%, K4%) = SUMA#(K2%, K4%) + VFX#(K2%)
                  End If
                  If TT%(K2%) = 255 Then
                      SUMA#(K2%, 1) = 0
                      SUMA#(K2%, 2) = CDbl(Val(VFS$(K2%)))
                  End If
               Next K4%
            End If
            '
          End If
        Next K2%
        '
4       RENGLO$ = Space$(OFS%) + CXC$
        For K2% = 1 To NCL%
           TPAR$ = VFS$(K2%)
           Print #ASAL%, TPAR$
        Next K2%
        '
        '
19      Print #ASAL%, "@NEWLINE@"
        '
99  Next K1&
    '
    Close #ASAL%
    ASAL% = 0
    Shell "FLEXCEL.EXE " + LUCOM$, vbNormalFocus
    '
    Screen.MousePointer = 1
    '
End Sub
