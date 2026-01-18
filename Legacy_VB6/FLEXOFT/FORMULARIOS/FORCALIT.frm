VERSION 5.00
Begin VB.Form FORCALIT 
   BackColor       =   &H00D0F0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Re-Cálculo Consumos I.T."
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Opcion de Cálculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   105
      TabIndex        =   8
      Top             =   1320
      Width           =   2535
      Begin VB.CheckBox Check1 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Volver a Re-Calcular todo, desde el 02/07 de 2003"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Inicio del Régimen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   105
      TabIndex        =   4
      Top             =   210
      Width           =   2535
      Begin VB.CommandButton Command7 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2100
         TabIndex        =   5
         Top             =   435
         Width           =   195
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   330
         Left            =   210
         TabIndex        =   7
         Top             =   480
         Width           =   750
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   6
         Top             =   450
         Width           =   960
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   2835
      ScaleHeight     =   75
      ScaleWidth      =   765
      TabIndex        =   2
      Top             =   1650
      Width           =   825
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   3
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.CommandButton COMMAND1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      Picture         =   "FORCALIT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cierra - Abandona la Aplicación"
      Top             =   300
      Width           =   720
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      Picture         =   "FORCALIT.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Inicia proceso de Cálculo"
      Top             =   960
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   2250
      Picture         =   "FORCALIT.frx":0454
      Top             =   1860
      Width           =   2010
   End
End
Attribute VB_Name = "FORCALIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim INTERRUM%

Private Sub Command1_Click()
   INTERRUM% = 1
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command11_Click()
   If FECHANUM(Label9) < 1 Then
     Call MENSERR(24, "Falta Fecha Inicio Régimen")
     Exit Sub
   End If
   '
   Command11.Enabled = False
   INTERRUM% = 0
   Call GRACONTROL("", "FINITEMP", Label9)
   '
   Screen.MousePointer = 11
   '
'LAS SIGUIENTES INSTRUCCIONES SON PARA REALIZAR EL CALCULO DE CERO
'*****************************************************************
   If Check1.Value = 1 Then
     Call BLANARCH("CONSIMPT")
     Call BLANARCH("EXPOTIP")
     Call BLANARCH("MOVSTEX0")
   End If
   '
   ' 1.- Determinar Clientes del Exterior
   TECLEX$ = "/"
   FIN& = ULTREG&("DEUDOR1")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("DEUDOR1", U&)
     NCLIE% = CVI(Left$(X$, 2))
     If PIVACLI%(NCLIE%) = 5 Then
       TECLEX$ = TECLEX$ + TRIM$(Str$(NCLIE%)) + "/"
     End If
   Next U&
   '
   ' 2.- Movimientos de Exportación
   Call COPYSTRU("MOVISTO", "MOVSTEX0")
   FIN& = ULTREG&("MOVISTO")
   FINITEM% = FECHANUM(Label9)
   PRIMOVI& = 1
   '
   ' LEE MOVISTO Y DETERMINA MOVIMIENTOS QUE SE ANULAN ENTRE SI
   ' Y LOS MARCA EN 112,6
   ' ADEMAS DETERMINA REGISTRO DE PRIMER MOVIMIENTO POSTERIOR A COMIENZO DEL REGIMEN EN PRIMOVI&
   For U& = 1 To FIN&
     If (U& Mod 100) = 0 Then Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("MOVISTO", U&)
     If CVI(Mid$(X$, 1, 2)) >= FINITEM% Then
       If PRIMOVI& <= 1 Then PRIMOVI& = U&
       NCLIE% = CVI(Mid$(X$, 84, 2))
       If NCLIE% > 0 Then
         NCLEX$ = "/" + TRIM$(Str$(NCLIE%)) + "/"
         If InStr(TECLEX$, NCLEX$) > 0 Then
           If TRIM$(Mid$(X$, 112, 6)) = "" Then ' no esta anulado
             CADEVU# = CVD(Mid$(X$, 88, 8)) - CVD(Mid$(X$, 96, 8))
             If CADEVU# >= 0.05 Then
               FEDEVU% = CVI(Left$(X$, 2))
               CIXI% = CVI(Mid$(X$, 3, 2))
               If CIXI% > 0 Then
                 '
                 RGG$ = RECFILT$("MOVISTO", 2, MKI$(CIXI%))
                 VUELTA% = 1
                 '
10               For UK1& = Len(RGG$) - 3 To 1 Step -4
                   If UK1& > 0 Then
                     UK& = CVS(Mid$(RGG$, UK1&, 4))
                     YY$ = REGLEIDO$("MOVISTO", UK&)
                     If (VUELTA% = 1 And CVI(Left$(YY$, 2)) = FEDEVU%) Or (VUELTA% = 2 And UK& < U&) Then
                       If TRIM$(Mid$(YY$, 112, 6)) = "" Then
                         If CVI(Mid$(YY$, 84, 2)) = NCLIE% Then
                           If CVI(Mid$(YY$, 3, 2)) = CIXI% Then
                             CASALI# = CVD(Mid$(YY$, 96, 8))
                             If Abs(CADEVU# - CASALI#) < 0.05 Then
                               MOANU$ = TRIM$(Mid$(YY$, 26, 7)): While Len(MOANU$) > 3: MOANU$ = Mid$(MOANU$, 2): Wend
                               MOANU$ = Mid$(YY$, 23, 3) + MOANU$
                               Call REPLA(X$, MOANU$, 112, 6)
                               Call GRAREG("MOVISTO", X$, U&)
                               MOANU$ = TRIM$(Mid$(X$, 26, 7)): While Len(MOANU$) > 3: MOANU$ = Mid$(MOANU$, 2): Wend
                               MOANU$ = Mid$(X$, 23, 3) + MOANU$
                               Call REPLA(YY$, MOANU$, 112, 6)
                               Call GRAREG("MOVISTO", YY$, UK&)
                               GoTo 12
                             End If
                           End If
                         End If
                       End If
                     End If
                   End If
                 Next UK1&
                 '
                 If VUELTA% = 1 Then
                   VUELTA% = 2
                   GoTo 10
                 End If
                 '
               End If
             End If
           End If
         End If
       End If
     End If
12 Next U&
   '
   ' LEE SECUENCIALMENTE MOVISTO Y CARGA MOVSTEX0
   For U& = PRIMOVI& To FIN&
     Call TRACE(20, 1 + U& - PRIMOVI&, 1 + FIN& - PRIMOVI&)
     X$ = REGLEIDO$("MOVISTO", U&)
     If CVI(Mid$(X$, 1, 2)) >= FINITEM% Then
       NCLIE% = CVI(Mid$(X$, 84, 2))
       If NCLIE% > 0 Then
         NCLEX$ = "/" + TRIM$(Str$(NCLIE%)) + "/"
         If InStr(TECLEX$, NCLEX$) > 0 Then
           CIXI% = CVI(Mid$(X$, 3, 2))
           If CIXI% > 0 Then
             CASALI# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
             If Abs(CASALI#) >= 0.05 Then
               '
               BARGU$ = Left$(X$, 103)
               RFF$ = RECFILT$("MOVSTEX0", 2, MKI$(CIXI%))
               '
               For KU& = 1 To Len(RFF$) Step 4
                 JK& = CVS(Mid$(RFF$, KU&, 4))
                 YY$ = REGLEIDO$("MOVSTEX0", JK&)
                 If Left$(YY$, 103) = BARGU$ Then GoTo 19
               Next KU&
               '
               For JK& = 1 To ULTREG&("MOVSTEX0")
                 YY$ = REGLEIDO$("MOVSTEX0", JK&)
                 If Left$(YY$, 103) = BARGU$ Then GoTo 19
               Next JK&
               '
               Call REGAPP("MOVSTEX0", X$)
               Call CIERRARCH("MOVSTEX0")
               '
             End If
             '
           End If
           '
         End If
       End If
     End If
19   DoEvents
     If INTERRUM% > 0 Then GoTo 96
   Next U&
   '
   ' CONVIERTE LOS MOVIMIENTOS DE EXPORTACION LEIDOS DE MOVSTEX0
   ' EN CONSUMOS DE MATERIA PRIMA Y LOS CARGA / AGREGA A EXPOTIP
   '
20 FIN& = ULTREG&("MOVSTEX0")
   REBLA$ = REGBLAN$("EXPOTIP")
   NMS% = NUMAS%
   ILOTEC& = 0
   UREXPO& = ULTREG&("EXPOTIP")
   If UREXPO& > 0 Then
     YY$ = REGLEIDO$("EXPOTIP", UREXPO&)
     ILOTEC& = CVS(Mid$(YY$, 45, 4))
   End If
   If ILOTEC& < UREXPO& Then ILOTEC& = UREXPO&
   '
   For U& = 1 To FIN&    ' ULTREG&("MOVSTEX0")
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("MOVSTEX0", U&)
     If TRIM$(Mid$(XX$, 112, 6)) = "" Then   ' DESCARTA ANULADOS
       FEXPOR% = CVI(Mid$(XX$, 1, 2))
       NREMI& = Val(Mid$(XX$, 26, 7))
       NCLIE% = CVI(Mid$(XX$, 84, 2))
       CIXI% = CVI(Mid$(XX$, 3, 2))
       CANEX# = CVD(Mid$(XX$, 96, 8))
       '
       If CANEX# >= 0.05 Then
         RFF$ = RECFILT$("EXPOTIP", 4, MKS$(NREMI&))
         For KU& = 1 To Len(RFF$) Step 4
           REGO& = CVS(Mid$(RFF$, KU&, 4))
           YY$ = REGLEIDO$("EXPOTIP", REGO&)
           If CVI(Left$(YY$, 2)) = FEXPOR% Then
             If CVI(Mid$(YY$, 3, 2)) = NCLIE% Then
               If CVI(Mid$(YY$, 55, 2)) = CIXI% Then
                 If Abs(CVD(Mid$(YY$, 57, 8)) - CANEX#) < 0.05 Then
                   GoTo 29   ' PROXIMO REGISTRO DE MOVSTEX0 - ya está en EXPOTIP
                 End If
               End If
             End If
           End If
         Next KU&
         '
         YY0$ = REBLA$
         Call REPLA(YY0$, MKI$(FEXPOR%), 1, 2)
         Call REPLA(YY0$, MKI$(NCLIE%), 3, 2)
         Call REPLA(YY0$, MKS$(NREMI&), 5, 4)
           NRELAR$ = Mid$(XX$, 23, 10)
         Call REPLA(YY0$, NRELAR$, 9, 16)
           NFACTX$ = Mid$(XX$, 33, 10)
         Call REPLA(YY0$, MKS$(Val(Mid$(NFACTX$, 4, 7))), 25, 4)
         Call REPLA(YY0$, NFACTX$, 29, 16)
         Call REPLA(YY0$, MKI$(CIXI%), 55, 2)
         Call REPLA(YY0$, MKD$(CANEX#), 57, 8)
         '
         RFF$ = RECFILT$("TIPIFIC", 1, MKI$(CIXI%))
         If Len(RFF$) > 0 Then
           For UK& = 1 To Len(RFF$) Step 4
             RETI& = CVS(Mid$(RFF$, UK&, 4))
             ZZ$ = REGLEIDO$("TIPIFIC", RETI&)
             CICO% = CVI(Mid$(ZZ$, 49, 2))
             If CICO% > 0 Then
               If CICO% <= NMS% Then
                 UNICO$ = Mid$(ZZ$, 51, 4)
                 If TRIM$(UNICO$) = "" Then
                   UNICO$ = UNIMED$(CICO%)
                 End If
                 USNETO# = CVD(Mid$(ZZ$, 55, 8))
                 PORMER = CVS(Mid$(ZZ$, 71, 4))
                 PORPER = CVS(Mid$(ZZ$, 83, 4))
                 USBRUT# = CVD(Mid$(ZZ$, 87, 8))
' AQUI DEFINIR CON CUANTOS DECIMALES TRABAJAR
                 '
                 YY$ = YY0$
                 ILOTEC& = ILOTEC& + 1
                 Call REPLA(YY$, MKS$(ILOTEC&), 45, 4)
                 Call REPLA(YY$, MKI$(CICO%), 65, 2)
                 Call REPLA(YY$, UNICO$, 67, 4)
                 Call REPLA(YY$, MKD$(USNETO#), 71, 8)
                 Call REPLA(YY$, MKS$(PORMER), 79, 4)
                 Call REPLA(YY$, MKS$(PORPER), 83, 4)
                 Call REPLA(YY$, MKD$(USBRUT#), 87, 8)
                 Call REPLA(YY$, MKD$(USBRUT# * CANEX#), 95, 8)
                 Call REGAPP("EXPOTIP", YY$)
               End If
             End If
           Next UK&
         End If
       End If
     End If
29 Next U&
   Call CIERRARCH("EXPOTIP")
   Call CIERRARCH("*.*")
   '
   ' LEE LAS RECEPCIONES DE MATERIAS PRIMAS DE IMPORTACION TEMPORARIA
   ' Y LAS CARGA EN CONSIMPT Y SALIMPT
   J1& = 0
   RFF$ = RECFILT$("REMAPRI", 23, "XT")  ' es importacion (X) y temporaria (T)
   For UK& = 1 To Len(RFF$) Step 4
     U& = CVS(Mid$(RFF$, UK&, 4))
     XX$ = REGLEIDO$("REMAPRI", U&)
     If Asc(Mid$(XX$, 98, 1)) = 9 Then GoTo 39 ' marca anulado
     '
     ILOTREC& = CVS(Mid$(XX$, 237, 4))
     CIREC% = CVI(Mid$(XX$, 83, 2))
     CANREC# = CVD(Mid$(XX$, 163, 8))
     FEREC% = CVI(Mid$(XX$, 195, 2)) ' FECHA ADUANA
     '
     RGG$ = RECFILT$("CONSIMPT", 1, MKS$(ILOTREC&))
     For JK& = 1 To Len(RGG$) Step 4
       JJ& = CVS(Mid$(RGG$, JK&, 4))
       YY$ = REGLEIDO$("CONSIMPT", JJ&)
       If CVI(Mid$(YY$, 9, 2)) <> CIREC% Then
         Call MENSERR(24, "Inconsistencia de Datos - Consulte.")
         Call CIERRARCH("*.*")
         Call FINAL("")
       End If
       CANREC# = CANREC# - CVD(Mid$(YY$, 11, 8))
     Next JK&
     '
     If CANREC# > 0 Then
       IDENLO$ = Mid$(XX$, 69, 12)
       PERMIT$ = Mid$(XX$, 213, 24)
       FCPROV$ = Mid$(XX$, 135, 16)
       If TRIM$(FCPROV$) = "" Then FCPROV$ = Mid$(XX$, 119, 16) ' reemplaza por remito
       UNIDX$ = Mid$(XX$, 85, 4)
       CATORE# = CVD(Mid$(XX$, 163, 8))
       
       ZZ$ = REGBLAN$("SALIMPT")
       Call REPLA(ZZ$, MKS$(ILOTREC&), 1, 4)
       Call REPLA(ZZ$, MKI$(CIREC%), 9, 2)
       Call REPLA(ZZ$, IDENLO$, 21, 12)
       '
       Call REPLA(ZZ$, PERMIT$, 33, 24)
       Call REPLA(ZZ$, MKI$(FEREC%), 57, 2)
       Call REPLA(ZZ$, FCPROV$, 59, 16)
       Call REPLA(ZZ$, MKI$(CIREC%), 75, 2)
       Call REPLA(ZZ$, UNIDX$, 77, 4)
       Call REPLA(ZZ$, MKD$(CATORE#), 81, 8)
       Call REPLA(ZZ$, MKD$(CATORE# - CANREC#), 89, 8)
       Call REPLA(ZZ$, MKD$(CANREC#), 97, 8)
       Call REPLA(ZZ$, String$(24, 0), 105, 24)
       '
       J1& = J1& + 1
       Call GRAREG("SALIMPT", ZZ$, J1&)
     End If
39 Next UK&
   '
   Call SETULTREG("SALIMPT", J1&)
   Call CIERRARCH("SALIMPT")
   '
   J2& = 0
   FIN& = ULTREG&("EXPOTIP")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("EXPOTIP", U&)
     ILOTEXP& = CVS(Mid$(XX$, 45, 4))
     CIEXP% = CVI(Mid$(XX$, 65, 2))
     CANEXP# = CVD(Mid$(XX$, 95, 8))
     FEXPO% = CVI(Mid$(XX$, 1, 2)) ' FECHA FACTURA
     '
     RGG$ = RECFILT$("CONSIMPT", 2, MKS$(ILOTEXP&))
     For JK& = 1 To Len(RGG$) Step 4
       JJ& = CVS(Mid$(RGG$, JK&, 4))
       YY$ = REGLEIDO$("CONSIMPT", JJ&)
       If CVI(Mid$(YY$, 9, 2)) <> CIEXP% Then
         Call MENSERR(24, "Inconsistencia de Datos - Consulte.")
         Call CIERRARCH("*.*")
         Call FINAL("")
       End If
       CANEXP# = CANEXP# - CVD(Mid$(YY$, 11, 8))
     Next JK&
     '
     If CANEXP# > 0 Then
       ZZ$ = REGBLAN$("MEXPSIMT")
       Call REPLA(ZZ$, MKS$(ILOTEXP&), 1, 4)
       Call REPLA(ZZ$, MKS$(U&), 5, 4)
       Call REPLA(ZZ$, MKI$(CIEXP%), 9, 2)
       Call REPLA(ZZ$, MKI$(FEXPO%), 57, 2)
       Call REPLA(ZZ$, MKD$(CANEXP#), 97, 8)
       Call REPLA(ZZ$, String$(24, 0), 105, 24)
       '
       J2& = J2& + 1
       Call GRAREG("MEXPSIMT", ZZ$, J2&)
     End If
   Next U&
   '
   Call SETULTREG("MEXPSIMT", J2&)
   Call CIERRARCH("MEXPSIMT")
   '
   '
40 FIN& = ULTREG&("SALIMPT")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("SALIMPT", U&)
     IDLOTIN& = CVS(Left$(XX$, 4))
     CIXI% = CVI(Mid$(XX$, 9, 2))
     SALOTE# = CVD(Mid$(XX$, 97, 8))
     FRECEP% = CVI(Mid$(XX$, 57, 2))
     '
     If SALOTE# >= 0.0005 Then
       For KJ& = 1 To ULTREG&("MEXPSIMT")
         YY$ = REGLEIDO$("MEXPSIMT", KJ&)
         If CVI(Mid$(YY$, 9, 2)) = CIXI% Then
           CSALEX# = CVD(Mid$(YY$, 97, 8))
           If CSALEX# >= 0.0005 Then
             If CVI(Mid$(YY$, 57, 2)) >= FRECEP% Then
               IDLOCON& = CVS(Left$(YY$, 4))
               REXPOTIP& = CVS(Mid$(YY$, 5, 4))
               CACONSU# = SALOTE#
               If CSALEX# < CACONSU# Then CACONSU# = CSALEX#
               Call REPLA(YY$, MKD$(CSALEX# - CACONSU#), 97, 8)
               Call GRAREG("MEXPSIMT", YY$, KJ&)
               '
               XX$ = REGLEIDO$("SALIMPT", U&)
               COLOTE# = CVD(Mid$(XX$, 89, 8)) + CACONSU#
               SALOTE# = CVD(Mid$(XX$, 97, 8)) - CACONSU#
               Call REPLA(XX$, MKD$(COLOTE#), 89, 8)
               Call REPLA(XX$, MKD$(SALOTE#), 97, 8)
               Call GRAREG("SALIMPT", XX$, U&)
               '
               ZZ$ = REGBLAN$("CONSIMPT")
               Call REPLA(ZZ$, MKS$(IDLOTIN&), 1, 4)
               Call REPLA(ZZ$, MKS$(IDLOCON&), 5, 4)
               Call REPLA(ZZ$, MKI$(CIXI%), 9, 2)
               Call REPLA(ZZ$, MKD$(CACONSU#), 11, 8)
                 IDLX$ = Mid$(XX$, 19, 70)
               Call REPLA(ZZ$, IDLX$, 19, 70)
               '
               YY$ = REGLEIDO$("EXPOTIP", REXPOTIP&)
               Call REPLA(ZZ$, YY$, 129, 128)
               Call REPLA(ZZ$, MKD$(CACONSU#), 223, 8)
               Call REPLA(ZZ$, MKD$(SALOTE#), 231, 8)
               '
               USOTIPI# = CVD(Mid$(ZZ$, 215, 8))
               If USOTIPI# >= 0.0000005 Then
                 If CVD(Mid$(ZZ$, 185, 8)) > CACONSU# / USOTIPI# Then
                   UNICON# = CACONSU# / USOTIPI#
                   Call REPLA(ZZ$, MKD$(UNICON#), 185, 8)
                 End If
               End If
               '
               Call REGAPP("CONSIMPT", ZZ$)
               '
               Call CIERRARCH("SALIMPT")
               Call CIERRARCH("MEXPSIMT")
               GoTo 40    ' If CACONSU# > 0.00005 Then GoTo 40
             End If
           End If
         End If
       Next KJ&
     End If
   Next U&
   '
   '
95 Call CIERRARCH("*.*")
   HOII% = HOY(HOS$)
   FULTICA$ = HOS$ + " - " + Time$ + " hs"
   Call GRACONTROL("CALCONIT", "FULTICAL", FULTICA$)
   Call COMUNI("Cálculo Completo")
   '
96 Call CIERRARCH("*.*")
   INTERRUM% = 0
   Screen.MousePointer = 1
   '
   Command11.Enabled = True
End Sub

Private Sub Command7_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
     Label9 = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Form_Load()
   Label9 = CONTROL$("", "FINITEMP")
   DoEvents
End Sub

Private Sub Label9_DblClick()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
     Label9 = VALACT1$("SELFECHA")
   End If
End Sub
