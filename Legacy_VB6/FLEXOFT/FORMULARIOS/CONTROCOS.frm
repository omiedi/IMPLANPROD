VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CONTROCOS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procesos de Control de Costos"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8715
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   8715
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   31
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   7665
      Picture         =   "CONTROCOS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Control Costo Reposición vs. Precio Ultima Compra "
      Top             =   1995
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D0E0FF&
      Caption         =   "Antiguedad de Costo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   3885
      TabIndex        =   26
      Top             =   1575
      Width           =   3585
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   27
         Top             =   420
         Width           =   2325
         Begin VB.CommandButton Command6 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "CONTROCOS.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Control Costos Anteriores a Fecha"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Control por Fecha Costo Reposición"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   630
            TabIndex        =   29
            Top             =   105
            Width           =   1620
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0E0FF&
      Caption         =   "Proveedores Habituales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   3885
      TabIndex        =   22
      Top             =   210
      Width           =   3585
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   23
         Top             =   420
         Width           =   2325
         Begin VB.CommandButton Command5 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "CONTROCOS.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Control Proveedor Habitual Vs. Ultima Compra "
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Ultima Compra vs. Proveed.Habitual"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   630
            TabIndex        =   25
            Top             =   105
            Width           =   1620
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0E0FF&
      Caption         =   "Control Operaciones Standard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   210
      TabIndex        =   15
      Top             =   2310
      Width           =   3585
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   19
         Top             =   1050
         Width           =   2430
         Begin VB.CommandButton Command4 
            Caption         =   "MOD"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "CONTROCOS.frx":091E
            TabIndex        =   20
            ToolTipText     =   "Control Mano de Obra Directa Material Fabricado"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Mano de Obra Di- recta Fabricación"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   21
            Top             =   105
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   16
         Top             =   525
         Width           =   2430
         Begin VB.CommandButton Command2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "CONTROCOS.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Operaciones Standard con Datos Faltantes"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Operaciones Stan- dard Incompletas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   630
            TabIndex        =   18
            Top             =   40
            Width           =   1620
         End
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00D0E0FF&
      Caption         =   "Procesos por Unica Vez"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   3885
      TabIndex        =   11
      Top             =   2940
      Width           =   3585
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   12
         Top             =   420
         Width           =   2325
         Begin VB.CommandButton Command3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "CONTROCOS.frx":0D72
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Control y Corrección Moneda de Valuación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Refrescar Moneda de Valuación"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   630
            TabIndex        =   14
            Top             =   105
            Width           =   1620
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00D0E0FF&
      Caption         =   "Costos de Mercadería Vendida"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   210
      TabIndex        =   4
      Top             =   210
      Width           =   3585
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   5
         Top             =   1050
         Width           =   2430
         Begin VB.CommandButton Command14 
            Caption         =   "MOD"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "CONTROCOS.frx":107C
            TabIndex        =   6
            ToolTipText     =   "Control Mano de Obra Directa Material Vendido"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Mano de Obra Di- recta Mat.Vendido"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   7
            Top             =   105
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture14 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   8
         Top             =   525
         Width           =   2430
         Begin VB.CommandButton Command22 
            Caption         =   "ABC"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Análisis ABC Costos Material Vendido"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "ABC Costo Insumos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Index           =   0
            Left            =   630
            TabIndex        =   10
            Top             =   105
            Width           =   1695
         End
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   7665
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3360
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   7665
      Picture         =   "CONTROCOS.frx":11C6
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   7665
      Picture         =   "CONTROCOS.frx":1310
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1155
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "CONTROCOS.frx":161A
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "CONTROCOS.frx":184E
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   7140
      Picture         =   "CONTROCOS.frx":18C8
      Top             =   3675
      Width           =   2010
   End
End
Attribute VB_Name = "CONTROCOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   '
   Dim CANTUX(32767), COSTOT(32767), CACONS(32767)
   '
10 FEX = HOY(HO$)
   HAS% = FEX: HASS$ = FECHATEX$(FEX)
   DES% = FECHANUM("01" + Mid$(HO$, 3))
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
   OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
     GoTo 99
   End If
   '
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If DES% > HAS% Then
      GoTo 10
   End If
   FEX = DES%: PERIO$ = FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   TTXX$ = "Desea Registrar como Standard\los Tiempos Promedio Registrados\para Operaciones sin T.Standard ?\\"
   OPXX% = COMALTER%(TTXX$ + "No, Conservar Actuales\Si, Registrar Promedios")
   '
11 Screen.MousePointer = 11
   COSTOMA = 0
   FIN& = ULTREG&("MOVISTO")
   For i& = 1 To FIN&
     Call TRACE(20, i&, FIN&)
     XS$ = REGLEIDO$("MOVISTO", i&)
     FF% = CVI(Left$(XS$, 2))
     If FF% >= DES% Then
       If FF% <= HAS% Then
         COMO$ = UCase$(Mid$(XS$, 33, 2))
         If COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
           If TRIM$(UCase$(Mid$(XS$, 23, 2))) <> "PL" Then
             CIXI% = CVI(Mid$(XS$, 3, 2))
             If CIXI% > 0 Then
               CANTUX(CIXI%) = CANTUX(CIXI%) + CVD(Mid$(XS$, 96, 8)) - CVD(Mid$(XS$, 88, 8))
             End If
           End If
         End If
       End If
     End If
   Next i&
   '
   SUMACO = 0
   Screen.MousePointer = 11
   FIN& = NUMAS%
   For UI% = 1 To NUMAS%
     U1& = UI%: Call TRACE(20, U1&, FIN&)
     If Abs(CANTUX(UI%)) >= 0.1 Then
       Call EXPLOMUL(UI%)
       If CONTAMUL% > 0 Then
         For KKI% = 1 To CONTAMUL%
           CIXI% = CIK%(KKI%)
           CACOM = CANT(KKI%)
           If CIXI% > 0 Then
             CACONS(CIXI%) = CACONS(CIXI%) + CANTUX(UI%) * CACOM
           End If
         Next KKI%
       End If
     End If
   Next UI%
   '
   URE& = ULTREG&("SECOPER")
   REBLA$ = REGBLAN$("COMODIR")
   JJ& = 0
   Screen.MousePointer = 11
   FIN& = NUMAS%
   For UI% = 1 To NUMAS%
     U& = UI%
     Call TRACE(20, U&, FIN&)
     If Abs(CACONS(UI%)) > 0.01 Then
       RFF$ = RECFILT$("SECOPER", 1, MKI$(UI%))
       RCT$ = RECFILT$("COSOPTER", 1, MKI$(UI%))
       If Len(RFF$) > 0 Then
         For UJ& = 1 To Len(RFF$) Step 4
           i& = CVS(Mid$(RFF$, UJ&, 4))
           If i& > 0 Then
             If i& <= URE& Then
               SCP$ = REGLEIDO$("SECOPER", i&)
               NOPY% = CVI(Mid$(SCP$, 3, 2))
               DOPE$ = Mid$(SCP$, 5, 48)
               LOTESTA = LOTESTAN(CIXI%): If LOTESTA < 1 Then LOTESTA = 1
               CICLOHE = CVS(Mid$(SCP$, 99, 4)): If CICLOHE < 1 Then CICLOHE = 1
               MINPREP = 60 * CVS(Mid$(SCP$, 59, 4))
               MINFIJO = 60 * CVS(Mid$(SCP$, 63, 4))
               MINVARI = LOTESTA * CVS(Mid$(SCP$, 67, 4))
               MINHERR = LOTESTA * 60 * CVS(Mid$(SCP$, 95, 4)) / CICLOHE
               MINDEMO = (MINFIJO + MINVARI + MINHERR) * CVS(Mid$(SCP$, 75, 4)) / 100
               MINLOTE = MINPREP + MINFIJO + MINVARI + MINHERR + MINDEMO
               HORAPROD = CACONS(UI%) * MINLOTE / 60 / LOTESTA
               '
               CAOPS = CVS(Mid$(SCP$, 83, 4))
               CATOP% = Asc(Mid$(SCP$, 87, 1))
               EQUIPO$ = TRIM$(Mid$(SCP$, 53, 6))
               COSMOD = CAOPS * HORAPROD * VALHORA(CATOP%)
               COSMAQ = HORAPROD * CUHORMAQ(EQUIPO$)
               '
               TERCOP% = 0: COSTER = 0
               For UT& = 1 To Len(RCT$) Step 4
                 IT& = CVS(Mid$(RCT$, UT&, 4))
                 If IT& > 0 Then
                   If IT& <= UCT& Then
                     SCT$ = REGLEIDO$("COSOPTER", IT&)
                     If CVI(Mid$(SCT$, 3, 2)) = NOPY% Then
                       MONEI% = CVI(Mid$(SCT$, 15, 2))
                       COSTER = CACOM * CVD(Mid$(SCT$, 7, 8)) * TICAMONE(MONEI%)
                       If MONEI% <= 1 Then
                         FEBAS% = CVI(Mid$(SCT$, 17, 2))
                         FEACT% = HOY(HOS$)
                         COSTER = CACOM * CVD(Mid$(SCT$, 7, 8)) * IAJUINF#(FEBAS%, FEACT%)
                       End If
                       If COSTER >= 0.00005 Then
                         If (COSTER < COSMOD + COSMAQ) Or ((COSMOD + COSMAQ) < 0.00005) Then
                           COSMOD = COSTER
                           COSMAQ = 0
                           HORAPROD = 0
                           TERCOP% = 1
                         End If
                       End If
                     End If
                   End If
                 End If
               Next UT&
               '
               Z$ = REBLA$
               Call REPLA(Z$, MKI$(UI%), 1, 2)
               Call REPLA(Z$, MKI$(NOPY%), 3, 2)
               Call REPLA(Z$, DOPE$, 5, 28)
               Call REPLA(Z$, MKS$(CACONS(UI%)), 33, 4)
               Call REPLA(Z$, MKS$(3600 * HORAPROD / CACONS(UI%)), 37, 4)
               Call REPLA(Z$, MKS$(3600 * HORAPROD), 45, 4)
               Call REPLA(Z$, MKS$(i&), 69, 4)
               '
               JJ& = JJ& + 1
               Call GRAREG("COMODIR", Z$, JJ&)
             End If
           End If
         Next UJ&
       End If
     End If
   Next UI%
   '
   Call SETULTREG("COMODIR", JJ&)
   Call CIERRARCH("COMODIR")
   '
   FIN& = ULTREG&("COMODIR")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("COMODIR", U&)
     CII% = CVI(Left$(XX$, 2))
     NPI% = CVI(Mid$(XX$, 3, 2))
     TISTAU = CVS(Mid$(XX$, 37, 4))
     RFF$ = RECFILT$("OPERFAB", 2, MKI$(CII%))
     SUMATI = 0: SUMACA = 0: TIUNMA = 0: TIUNMI = 0
     For KU& = 1 To Len(RFF$) Step 4
       REGE& = CVS(Mid$(RFF$, KU&, 4))
       YY$ = REGLEIDO$("OPERFAB", REGE&)
       If CVI(Mid$(YY$, 15, 2)) = NPI% Then
         CAREA = CVD(Mid$(YY$, 87, 8))
         If CAREA > 0.05 Then
           TITOT = 60 * CVS(Mid$(YY$, 161, 4))
           If TISTAU > 0.1 Then
             If TITOT / CAREA > 100 * TISTAU Then
               Call REPLA(YY$, MKS$(0), 161, 4)
               Call GRAREG("OPERFAB", YY$, REGE&)
               TITOT = 0
             End If
           End If
           If TITOT > 1 Then
             If TIUNMI = 0 Then TIUNMI = 9999999
             TIREUNI = TITOT / CAREA
             If TIREUNI > TIUNMA Then TIUNMA = TIREUNI
             If TIREUNI < TIUNMI Then TIUNMI = TIREUNI
             SUMATI = SUMATI + TITOT
             SUMACA = SUMACA + CAREA
           End If
         End If
       End If
     Next KU&
     '
     TIUNPR = 0
     If SUMACA > 0.05 Then
       TIUNPR = SUMATI / SUMACA
       Call REPLA(XX$, MKS$(TIUNMI), 49, 4)
       Call REPLA(XX$, MKS$(TIUNMA), 53, 4)
       Call REPLA(XX$, MKS$(TIUNPR), 57, 4)
       '
       TIUNST = CVS(Mid$(XX$, 37, 4))
       RELA = 999.9: If TIUNST > 0.1 Then RELA = 100 * TIUNPR / TIUNST
       If RELA > 999.9 Then RELA = 999
       Call REPLA(XX$, MKS$(RELA), 61, 4)
     End If
     '
     If TIUNPR < 0.1 Then TIUNPR = CVS(Mid$(XX$, 37, 4)) ' TOMA EL STANDARD COMO REAL
     CACOU = CVS(Mid$(XX$, 33, 4))
     TITOU = CACOU * TIUNPR
     Call REPLA(XX$, MKS$(TITOU), 65, 4)
     '
     Call GRAREG("COMODIR", XX$, U&)
     '
     If OPXX% = 2 Then
       If CVS(Mid$(XX$, 37, 4)) < 0.1 Then ' NO HAY T.STANDARD
         RESE& = CVS(Mid$(XX$, 69, 4))
         ZZ$ = REGLEIDO$("SECOPER", RESE&)
         Call REPLA(ZZ$, MKS$(TIUNPR / 60), 67, 4)
         Call GRAREG("SECOPER", ZZ$, RESE&)
       End If
     End If
     '
   Next U&
   '
   Call CIERRARCH("SECOPER")
   Call FILESORT("COMODIR", "", 1, 3, "ASC")
   Call CONECRUN("*COMODIR", "Control Mano de Obra Directa")
   '
99 Command14.Enabled = True
End Sub

Private Sub Command2_Click()
   '
   Command2.Enabled = False
   URE& = ULTREG&("SECOPER")
   'Call COPYSTRU("SECOPER", "SECOPFI")
   JJ& = 0
   Screen.MousePointer = 11
   FIN& = ULTREG&("SECOPER")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     SCP$ = REGLEIDO$("SECOPER", U&)
     CI0% = CVI(Left$(SCP$, 2))
     If CI0% < 1 Then GoTo 19
     '
     NOPY% = CVI(Mid$(SCP$, 3, 2))
     DOPE$ = Mid$(SCP$, 5, 48)
     LOTESTA = LOTESTAN(CIXI%): If LOTESTA < 1 Then LOTESTA = 1
     CICLOHE = CVS(Mid$(SCP$, 99, 4)): If CICLOHE < 1 Then CICLOHE = 1
     MINPREP = 60 * CVS(Mid$(SCP$, 59, 4))
     MINFIJO = 60 * CVS(Mid$(SCP$, 63, 4))
     MINVARI = LOTESTA * CVS(Mid$(SCP$, 67, 4))
     MINHERR = LOTESTA * 60 * CVS(Mid$(SCP$, 95, 4)) / CICLOHE
     MINDEMO = (MINFIJO + MINVARI + MINHERR) * CVS(Mid$(SCP$, 75, 4)) / 100
     MINLOTE = MINPREP + MINFIJO + MINVARI + MINHERR + MINDEMO
     '
     CAOPS = CVS(Mid$(SCP$, 83, 4))
     CATOP% = Asc(Mid$(SCP$, 87, 1))
     EQUIPO$ = TRIM$(Mid$(SCP$, 53, 6))
     '
     If MINLOTE > 0.1 Then
       If CAOPS > 0.1 Then
         If CATOP% > 0 Then
           If TRIM$(EQUIPO$) <> "" Then
             GoTo 19
           End If
         End If
       End If
     End If
     '
     RNC$ = RECFILT$("COSOPTER", 1, MKI$(CI0%))
     For UL& = 1 To Len(RNC$) Step 4
       RENOTA& = CVS(Mid$(RNC$, UL&, 4))
       TTZZ$ = REGLEIDO$("COSOPTER", RENOTA&)
       NOPZ% = CVI(Mid$(TTZZ$, 3, 2))
       If NOPZ% = NOPY% Then GoTo 19
     Next UL&
     '
     JJ& = JJ& + 1
     Y1$ = SCP$
     NOPES$ = TRIM$(CSTRING$(Mid$(Y1$, 3, 2), 1, 6, 0, 2))
     While Len(NOPES$) < 4: NOPES$ = "." + NOPES$: Wend
     DOPES$ = Mid$(Y1$, 5, 62)
     TVAR = CVS(Mid$(Y1$, 67, 4)) * LOTESTA
     y$ = Y1$
     Call REPLA(y$, NOPES$, 3, 6)
     Call REPLA(y$, DOPES$, 9, 62)
     Call REPLA(y$, MKS$(TVAR), 71, 4)
     Call GRAREG("SECOPSEL", y$, JJ&)
     '
19 Next U&
   '
   Call SETULTREG("SECOPSEL", JJ&)
   '
   Call FILESORT("SECOPSEL", "", 1, 3, "ASC")
   Call FINAL("*LISECOP")
   '
   Command2.Enabled = True
   '
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   Dim CANTUX(32767), COSTOT(32767), CACONS(32767)
   '
   On Error Resume Next
   CODIACTI$ = ANACOSTO.Text1
   If Err > 0 Then
     Call MENSERR(24, "Imposible Ejecutar\desde esta Aplicación.\  \Ejecute desde Módulo de Análisis de Costos.")
     On Error GoTo 0
     GoTo 999
   End If
   On Error GoTo 0
   ANACOSTO.SSTab1.Tab = 3
10 FEX = HOY(HO$)
   HAS% = FEX: HASS$ = FECHATEX$(FEX)
   DES% = FECHANUM("01" + Mid$(HO$, 3))
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
   OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
     GoTo 99
   End If
   '
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If DES% > HAS% Then
      GoTo 10
   End If
   FEX = DES%: PERIO$ = FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
11 Screen.MousePointer = 11
   COSTOMA = 0
   FIN& = ULTREG&("MOVISTO")
   For i& = 1 To FIN&
     Call TRACE(20, i&, FIN&)
     XS$ = REGLEIDO$("MOVISTO", i&)
     FF% = CVI(Left$(XS$, 2))
     If FF% >= DES% Then
       If FF% <= HAS% Then
         COMO$ = UCase$(Mid$(XS$, 33, 2))
         If COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
           If TRIM$(UCase$(Mid$(XS$, 23, 2))) <> "PL" Then
             CIXI% = CVI(Mid$(XS$, 3, 2))
             If CIXI% > 0 Then
               CANTUX(CIXI%) = CANTUX(CIXI%) + CVD(Mid$(XS$, 96, 8)) - CVD(Mid$(XS$, 88, 8))
             End If
           End If
         End If
       End If
     End If
   Next i&
   '
   SUMACO = 0
   Screen.MousePointer = 11
   For UI% = 1 To NUMAS%
     If Abs(CANTUX(UI%)) >= 0.1 Then
       ANACOSTO.Text1 = CODEXT$(UI%)
       Screen.MousePointer = 11
       For UV& = 1 To ANACOSMAT.List1.ListCount
         TTXX$ = ANACOSMAT.List1.List(UV& - 1)
         CIXI% = CODINT%(TRIM$(Left$(TTXX$, 15)))
         If CIXI% > 0 Then
           COSTOT(CIXI%) = COSTOT(CIXI%) + CANTUX(UI%) * XVALO(Mid$(TTXX$, 85, 10))
           CACONS(CIXI%) = CACONS(CIXI%) + CANTUX(UI%) * XVALO(Mid$(TTXX$, 40, 10))
           SUMACO = SUMACO + CANTUX(UI%) * XVALO(Mid$(TTXX$, 85, 10))
           If COSTOT(CIXI%) > COSTOMA Then COSTOMA = COSTOT(CIXI%)
         End If
       Next UV&
     End If
   Next UI%
   '
   JJ& = 0
   Screen.MousePointer = 11
   FIN& = NUMAS%
   For UI% = 1 To NUMAS%
     U& = UI%
     Call TRACE(20, U&, FIN&)
     If Abs(COSTOT(UI%)) > 0.01 Then
       Z$ = REGBLAN$("RANKCOS")
       Call REPLA(Z$, MKS$(COSTOMA - COSTOT(UI%)), 1, 4)
       Call REPLA(Z$, MKI$(0), 5, 2)
       Call REPLA(Z$, "  ", 7, 2)
       Call REPLA(Z$, MKI$(UI%), 9, 2)
       Call REPLA(Z$, MKS$(COSUNI(UI%)), 11, 4)
       Call REPLA(Z$, DEMONECO$(MONECOSTO%(UI%)), 15, 2)
       Call REPLA(Z$, MKI$(FECOSTO%(UI%)), 17, 2)
       FEX1 = FECOSTO%(UI%): FEX2 = HOY(HOS$)
         ANTICOS = (DIENTRE%(FEX1, FEX2)) / 30
         If ANTICOS < 1 Then ANTICOS = 0
       Call REPLA(Z$, MKS$(ANTICOS), 19, 4)
         CUNI = COSUNI(UI%) * TICAMONE(MONECOSTO%(UI%)) * (1 + PORNACIO(UI%) / 100)
       Call REPLA(Z$, MKS$(CUNI), 23, 4)
       Call REPLA(Z$, MKS$(CACONS(UI%)), 27, 4)
       Call REPLA(Z$, MKS$(COSTOT(UI%)), 31, 4)
       Call REPLA(Z$, MKS$(100 * COSTOT(UI%) / SUMACO), 35, 4)
       '
       VUCOM = VULTICOM#(UI%, NPRX%)
       PORDIF = 999.9
       If CUNI > 0.005 Then PORDIF = 100 * (VUCOM - CUNI) / CUNI
       If Abs(PORDIF) > 3 Then
          Call REPLA(Z$, MKS$(VUCOM), 43, 4)
          If CUNI > 0.00005 Then
            Call REPLA(Z$, MKS$(PORDIF), 47, 4)
          End If
       End If
       If NPRX% <> PROVHABI(UI%) Then
         If NPRX% > 0 Then
           Call REPLA(Z$, rasocli$((-1) * NPRX%), 51, 30)
         End If
       End If
       '
       JJ& = JJ& + 1
       Call GRAREG("RANKCOS", Z$, JJ&)
     End If
   Next UI%
   Call SETULTREG("RANKCOS", JJ&)
   Call CIERRARCH("RANKCOS")
   '
   Call FILESORT("RANKCOS", "", 1, 1, "ASC")
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("RANKCOS")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     UI% = U&
     Z$ = REGLEIDO$("RANKCOS", U&)
     SUMAPOR = SUMAPOR + CVS(Mid$(Z$, 35, 4))
     CATABC$ = "A": If SUMAPOR > 80 Then CATABC$ = "B"
                    If SUMAPOR > 95 Then CATABC$ = "C"
     Call REPLA(Z$, MKI$(UI%), 5, 2)
     Call REPLA(Z$, CATABC$, 7, 2)
     Call GRAREG("RANKCOS", Z$, U&)
   Next U&
   Call CIERRARCH("RANKCOS")
   Screen.MousePointer = 1
   '
   Call CONECRUN("*RANKCOS", "Ranking de Costos Material Vendido")
   '
99 ANACOSTO.Text1 = CODIACTI$
999 Command22.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   If DERACCE%("FRESHMVA", "Actualizar Moneda de Valuación") < 2 Then GoTo 99
   
   FIN& = ULTREG&("OCOMDETA")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("OCOMDETA", U&)
     '
     STAT% = Asc(Mid$(XX$, 124, 1))
     If STAT% <> 9 Then           ' no esta anulada
       NUORSE& = CVS(Mid$(XX$, 117, 4))
       If NUORSE& < 1 Then ' excluye orden de servicios
         MONEMISION% = Asc(Mid$(XX$, 122, 1))
         If MONEMISION% < 1 Then MONEMISION% = 1
         MONEVALUA% = Asc(Mid$(XX$, 121, 1))
         If MONEVALUA% < 1 Then MONEVALUA% = 1
         VALOUNI$ = Mid$(XX$, 43, 12)
         NPRO% = CVI(Mid$(XX$, 7, 2))
         '
         MONEVAC% = MONEVACO%(NPRO%)
         If MONEVAC% <= 1 Then
           If MONEMISION% <= 1 Then
             COEFUNI = CVS(Mid$(XX$, 63, 4))
             If COEFUNI < 0.0001 Then COEFUNI = 1
             IUNI# = CVD(Mid$(XX$, 79, 8)) / COEFUNI
             Z1$ = Space$(12)
             Call REPLA(Z1$, "$", 1, 1)
             Call REPLA(Z1$, CSTRING$(MKD$(IUNI#), 3, 10, 4, 2), 3, 10)
             '
             Call REPLA(XX$, Z1$, 43, 12)
             Call REPLA(XX$, Chr$(MONEVAC%), 121, 1)
             Call GRAREG("OCOMDETA", XX$, U&)
             '
             CIXI% = CVI(Mid$(XX$, 9, 2))
             If NPRO% = PROVHABI%(CIXI%) Then
               VUNI = IUNI#
               FECOI% = CVI(Mid$(XX$, 5, 2))
               If CIXI% > 0 Then
                 If CIXI% <= NUMAS% Then
                   REMA& = CIXI% + 1
                   ZZ$ = REGLEIDO$("MASTER", REMA&)
                   Call REPLA(ZZ$, MKS$(VUNI), 111, 4)
                   Call REPLA(ZZ$, MKI$(FECOI%), 115, 2)
                   Call REPLA(ZZ$, Chr$(MONEVAC%), 95, 1)
                   Call GRAREG("MASTER", ZZ$, REMA&)
                 End If
               End If
             End If
           End If
         End If
       End If
     End If
   Next U&
   Call CIERRARCH("OCOMDETA")
99 Command3.Enabled = True
   '
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   '
   Dim CACONS(32767)
   '
10 FEX = HOY(HO$)
   HAS% = FEX: HASS$ = FECHATEX$(FEX)
   DES% = FECHANUM("01" + Mid$(HO$, 3))
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
   OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
     GoTo 99
   End If
   '
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If DES% > HAS% Then
      GoTo 10
   End If
   FEX = DES%: PERIO$ = FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   TTXX$ = "Desea Registrar como Standard\los Tiempos Promedio Registrados\para Operaciones sin T.Standard ?\\"
   OPXX% = COMALTER%(TTXX$ + "No, Conservar Actuales\Si, Registrar Promedios")
   '
11 Screen.MousePointer = 11
   COSTOMA = 0
   FIN& = ULTREG&("MOVISTO")
   For i& = 1 To FIN&
     Call TRACE(20, i&, FIN&)
     XS$ = REGLEIDO$("MOVISTO", i&)
     FF% = CVI(Left$(XS$, 2))
     If FF% >= DES% Then
       If FF% <= HAS% Then
         COMO$ = UCase$(Mid$(XS$, 23, 2))
         If COMO$ = "OF" Then
           If CVD(Mid$(XS$, 88, 8)) >= 0.05 Then
             CIXI% = CVI(Mid$(XS$, 3, 2))
             If CIXI% > 0 Then
               CACONS(CIXI%) = CACONS(CIXI%) + CVD(Mid$(XS$, 88, 8)) - CVD(Mid$(XS$, 96, 8))
             End If
           End If
         End If
       End If
     End If
   Next i&
   '
   SUMACO = 0
   Screen.MousePointer = 11
   URE& = ULTREG&("SECOPER")
   REBLA$ = REGBLAN$("COMODIR")
   JJ& = 0
   Screen.MousePointer = 11
   FIN& = NUMAS%
   For UI% = 1 To NUMAS%
     U& = UI%
     Call TRACE(20, U&, FIN&)
     If Abs(CACONS(UI%)) > 0.01 Then
       RFF$ = RECFILT$("SECOPER", 1, MKI$(UI%))
       RCT$ = RECFILT$("COSOPTER", 1, MKI$(UI%))
       If Len(RFF$) > 0 Then
         For UJ& = 1 To Len(RFF$) Step 4
           i& = CVS(Mid$(RFF$, UJ&, 4))
           If i& > 0 Then
             If i& <= URE& Then
               SCP$ = REGLEIDO$("SECOPER", i&)
               NOPY% = CVI(Mid$(SCP$, 3, 2))
               DOPE$ = Mid$(SCP$, 5, 48)
               LOTESTA = LOTESTAN(CIXI%): If LOTESTA < 1 Then LOTESTA = 1
               CICLOHE = CVS(Mid$(SCP$, 99, 4)): If CICLOHE < 1 Then CICLOHE = 1
               MINPREP = 60 * CVS(Mid$(SCP$, 59, 4))
               MINFIJO = 60 * CVS(Mid$(SCP$, 63, 4))
               MINVARI = LOTESTA * CVS(Mid$(SCP$, 67, 4))
               MINHERR = LOTESTA * 60 * CVS(Mid$(SCP$, 95, 4)) / CICLOHE
               MINDEMO = (MINFIJO + MINVARI + MINHERR) * CVS(Mid$(SCP$, 75, 4)) / 100
               MINLOTE = MINPREP + MINFIJO + MINVARI + MINHERR + MINDEMO
               HORAPROD = CACONS(UI%) * MINLOTE / 60 / LOTESTA
               '
               CAOPS = CVS(Mid$(SCP$, 83, 4))
               CATOP% = Asc(Mid$(SCP$, 87, 1))
               EQUIPO$ = TRIM$(Mid$(SCP$, 53, 6))
               COSMOD = CAOPS * HORAPROD * VALHORA(CATOP%)
               COSMAQ = HORAPROD * CUHORMAQ(EQUIPO$)
               '
               TERCOP% = 0: COSTER = 0
               For UT& = 1 To Len(RCT$) Step 4
                 IT& = CVS(Mid$(RCT$, UT&, 4))
                 If IT& > 0 Then
                   If IT& <= UCT& Then
                     SCT$ = REGLEIDO$("COSOPTER", IT&)
                     If CVI(Mid$(SCT$, 3, 2)) = NOPY% Then
                       MONEI% = CVI(Mid$(SCT$, 15, 2))
                       COSTER = CACOM * CVD(Mid$(SCT$, 7, 8)) * TICAMONE(MONEI%)
                       If MONEI% <= 1 Then
                         FEBAS% = CVI(Mid$(SCT$, 17, 2))
                         FEACT% = HOY(HOS$)
                         COSTER = CACOM * CVD(Mid$(SCT$, 7, 8)) * IAJUINF#(FEBAS%, FEACT%)
                       End If
                       If COSTER >= 0.00005 Then
                         If (COSTER < COSMOD + COSMAQ) Or ((COSMOD + COSMAQ) < 0.00005) Then
                           COSMOD = COSTER
                           COSMAQ = 0
                           HORAPROD = 0
                           TERCOP% = 1
                         End If
                       End If
                     End If
                   End If
                 End If
               Next UT&
               '
               Z$ = REBLA$
               Call REPLA(Z$, MKI$(UI%), 1, 2)
               Call REPLA(Z$, MKI$(NOPY%), 3, 2)
               Call REPLA(Z$, DOPE$, 5, 28)
               Call REPLA(Z$, MKS$(CACONS(UI%)), 33, 4)
               Call REPLA(Z$, MKS$(3600 * HORAPROD / CACONS(UI%)), 37, 4)
               Call REPLA(Z$, MKS$(3600 * HORAPROD), 45, 4)
               Call REPLA(Z$, MKS$(i&), 69, 4)
               '
               JJ& = JJ& + 1
               Call GRAREG("COMODIR", Z$, JJ&)
             End If
           End If
         Next UJ&
       End If
     End If
   Next UI%
   '
   Call SETULTREG("COMODIR", JJ&)
   Call CIERRARCH("COMODIR")
   '
   FIN& = ULTREG&("COMODIR")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("COMODIR", U&)
     CII% = CVI(Left$(XX$, 2))
     NPI% = CVI(Mid$(XX$, 3, 2))
     TISTAU = CVS(Mid$(XX$, 37, 4))
     RFF$ = RECFILT$("OPERFAB", 2, MKI$(CII%))
     SUMATI = 0: SUMACA = 0: TIUNMA = 0: TIUNMI = 0
     For KU& = 1 To Len(RFF$) Step 4
       REGE& = CVS(Mid$(RFF$, KU&, 4))
       YY$ = REGLEIDO$("OPERFAB", REGE&)
       If CVI(Mid$(YY$, 15, 2)) = NPI% Then
         CAREA = CVD(Mid$(YY$, 87, 8))
         If CAREA > 0.05 Then
           TITOT = 60 * CVS(Mid$(YY$, 161, 4))
           If TISTAU > 0.1 Then
             If TITOT / CAREA > 100 * TISTAU Then
               Call REPLA(YY$, MKS$(0), 161, 4)
               Call GRAREG("OPERFAB", YY$, REGE&)
               TITOT = 0
             End If
           End If
           If TITOT > 1 Then
             If TIUNMI = 0 Then TIUNMI = 9999999
             TIREUNI = TITOT / CAREA
             If TIREUNI > TIUNMA Then TIUNMA = TIREUNI
             If TIREUNI < TIUNMI Then TIUNMI = TIREUNI
             SUMATI = SUMATI + TITOT
             SUMACA = SUMACA + CAREA
           End If
         End If
       End If
     Next KU&
     '
     TIUNPR = 0
     If SUMACA > 0.05 Then
       TIUNPR = SUMATI / SUMACA
       Call REPLA(XX$, MKS$(TIUNMI), 49, 4)
       Call REPLA(XX$, MKS$(TIUNMA), 53, 4)
       Call REPLA(XX$, MKS$(TIUNPR), 57, 4)
       '
       TIUNST = CVS(Mid$(XX$, 37, 4))
       RELA = 999.9: If TIUNST > 0.1 Then RELA = 100 * TIUNPR / TIUNST
       If RELA > 999.9 Then RELA = 999
       Call REPLA(XX$, MKS$(RELA), 61, 4)
     End If
     '
     If TIUNPR < 0.1 Then TIUNPR = CVS(Mid$(XX$, 37, 4)) ' TOMA EL STANDARD COMO REAL
     CACOU = CVS(Mid$(XX$, 33, 4))
     TITOU = CACOU * TIUNPR
     Call REPLA(XX$, MKS$(TITOU), 65, 4)
     '
     Call GRAREG("COMODIR", XX$, U&)
     '
     If OPXX% = 2 Then
       If CVS(Mid$(XX$, 37, 4)) < 0.1 Then ' NO HAY T.STANDARD
         RESE& = CVS(Mid$(XX$, 69, 4))
         ZZ$ = REGLEIDO$("SECOPER", RESE&)
         Call REPLA(ZZ$, MKS$(TIUNPR / 60), 67, 4)
         Call GRAREG("SECOPER", ZZ$, RESE&)
       End If
     End If
     '
   Next U&
   '
   Call CIERRARCH("SECOPER")
   Call FILESORT("COMODIR", "", 1, 3, "ASC")
   Call CONECRUN("*COMODIR", "Control Mano de Obra Directa")
   '
99 Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Screen.MousePointer = 11
   Dim UCOM%(32767), PCOM%(32767)
   '
   FIN& = ULTREG&("OCOMDETA")
   NMS% = NUMAS%
   For REOCO& = 1 To FIN&
     Call TRACE(20, REOCO&, FIN&)
     XX$ = REGLEIDO$("OCOMDETA", REOCO&)
     STAT% = Asc(Mid$(XX$, 124, 1))
     If STAT% <> 9 Then           ' no esta anulada
       NUORSE& = CVS(Mid$(XX$, 117, 4))
       If NUORSE& < 1 Then ' excluye orden de servicios
         CIIX% = CVI(Mid$(XX$, 9, 2))
         If CIIX% > 0 Then
           If CIIX% <= NMS% Then
             FECO% = CVI(Mid$(XX$, 5, 2))
             If FECO% >= UCOM%(CIIX%) Then
               UCOM%(CIIX%) = FECO%
               PCOM%(CIIX%) = CVI(Mid$(XX$, 7, 2))
             End If
           End If
         End If
       End If
     End If
   Next REOCO&
   '
   Call CIERRARCH("OCOMDETA")
   '
   REBLA$ = REGBLAN$("CPROVHAB")
   FIN& = NUMAS%: JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     CIIX% = U&
     If UCOM%(CIIX%) <> 0 Then
       If STATUIT%(CIIX%) > 0 Then
         PHBL% = PROVHABI%(CIIX%)
         If PCOM%(CIIX%) <> PHBL% Then
           ZZ$ = REBLA$
           Call REPLA(ZZ$, MKI$(CIIX%), 1, 2)
           Call REPLA(ZZ$, MKI$(PHBL%), 3, 2)
           Call REPLA(ZZ$, MKI$(PCOM%(CIIX%)), 5, 2)
           Call REPLA(ZZ$, MKI$(UCOM%(CIIX%)), 7, 2)
           JJ& = JJ& + 1
           Call GRAREG("CPROVHAB", ZZ$, JJ&)
         End If
       End If
     End If
   Next U&
   Call SETULTREG("CPROVHAB", JJ&)
   '
   Call FILESORT("CPROVHAB", "", 1, 1, "ASC")
   Call CONECRUN("*CPROVHAB", "Control Proveedores Habituales")
   Screen.MousePointer = 1
   '
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   HOII% = HOY(HOS$)
   VDEF$ = MKI$(HOII%) + MKS$(1)
   OBX$ = OBJPLA$("CONFECOREP", VDEF$)
   If OBX$ <> "" Then
     Screen.MousePointer = 1
     FECOLI% = CVI(Left$(OBX$, 2))
     VALIM = CVS(Mid$(OBX$, 3, 4))
     '
     FIN& = NUMAS%: JJ& = 0
     REBLA$ = REGBLAN$("COREPVIE")
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       Screen.MousePointer = 11
       'REMAS$ = REGLEIDO$("MASTER", U& + 1)
       'FECOST% = CVI(Mid$(REMAS$, 115, 2))
       CIIX% = U&
       FECOST% = FECOSTO%(CIIX%)
       If STATITEM%(CIIX%) > 0 Then
         If FECOST% < FECOLI% Then
           PRHBL% = PROVHABI%(CIIX%)
           If PRHBL% > 0 Then
             COSUPE = COSUNI(CIIX%) * TICAMONE(MONECOSTO%(CIIX%)) * (1 + PORNACIO(CIIX%))
             If COSUPE > VALIM Then
               ZZ$ = REBLA$
               Call REPLA(ZZ$, MKI$(CIIX%), 1, 2)
               Call REPLA(ZZ$, ECOARCH$("TAMONED", Chr$(MONECOSTO%(CIIX%))), 3, 2)
               Call REPLA(ZZ$, MKS$(COSUNI(CIIX%)), 5, 4)
               Call REPLA(ZZ$, MKI$(FECOST%), 9, 2)
               Call REPLA(ZZ$, MKI$(PRHBL%), 11, 2)
               JJ& = JJ& + 1
               Call GRAREG("COREPVIE", ZZ$, JJ&)
             End If
           End If
         End If
       End If
     Next U&
     Call SETULTREG("COREPVIE", JJ&)
     '
     Call FILESORT("COREPVIE", "", 1, 1, "ASC")
     Call CONECRUN("*COREPVIE", "Control por Antiguedad de Precio")
     Screen.MousePointer = 1
     '
   End If
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   Screen.MousePointer = 11
   Dim PREUCO#(32767), MONEUCO%(32767)
   '
   For KLI% = 1 To 32766
     PREUCO#(KLI%) = 0
     MONEUCO%(KLI%) = 0
   Next KLI%
   '
   FIN& = ULTREG&("OCOMDETA")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("OCOMDETA", U&)
     If Asc(Mid$(XX$, 124, 1)) < 9 Then
       NUORSE& = CVS(Mid$(XX$, 117, 4))
       If NUORSE& < 1 Then ' excluye orden de servicios
         PRUNI# = CVD(Mid$(XX$, 67, 8))
         If PRUNI# >= 0.00005 Then
           CI0% = CVI(Mid$(XX$, 9, 2))
           COEFI = CVS(Mid$(XX$, 63, 4))
           If COEFI < 0.0001 Then COEFI = 1
           PRUNI# = PRUNI# / COEFI
           MONEMIX% = Asc(Mid$(XX$, 122, 2))
           If MONEMIX% < 1 Then MONEMIX% = 1
           PREUCO#(CI0%) = PRUNI#
           MONEUCO%(CI0%) = MONEMIX%
         End If
       End If
     End If
   Next U&
   '
   JJ& = 0
   REBLA$ = REGBLAN$("RANKCOS")
   FIN& = NUINV%
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     IK1$ = REGLEIDO$("INVERT", U& + 1)
     CI0% = CVI(Mid$(IK1$, 17, 2))
     If CI0% > 0 Then
       If CI0% <= NUMAS% Then
         If STATUIT%(CI0%) > 0 Then
           If PREUCO#(CI0%) > 0.00001 Then
             NMC% = MONECOSTO%(CI0%): If MNC% < 1 Then MNC% = 1
             If MONEUCO%(CI0%) <> MNC% Or Abs((PREUCO#(CI0%) - COSUNI(CI0%)) / PREUCO#(CI0%)) > 0.01 Then
               Z$ = REGBLAN$("RANKCOS")
               Call REPLA(Z$, MKI$(CI0%), 9, 2)
               Call REPLA(Z$, MKS$(COSUNI(CI0%)), 11, 4)
               Call REPLA(Z$, DEMONECO$(MONECOSTO%(CI0%)), 15, 2)
               Call REPLA(Z$, MKI$(FECOSTO%(CI0%)), 17, 2)
                 CUNI = COSUNI(CI0%) * TICAMONE(MONECOSTO%(CI0%)) * (1 + PORNACIO(CI0%) / 100)
               Call REPLA(Z$, MKS$(CUNI), 23, 4)
                 VUCOM = PREUCO#(CI0%) * TICAMONE(MONEUCO%(CI0%)) * (1 + PORNACIO(CI0%) / 100)
               PORDIF = 999.9
               If CUNI > 0.005 Then PORDIF = 100 * (VUCOM - CUNI) / CUNI
               If Abs(PORDIF) > 3 Then
                  Call REPLA(Z$, MKS$(VUCOM), 43, 4)
                  If CUNI > 0.00005 Then
                    Call REPLA(Z$, MKS$(PORDIF), 47, 4)
                  End If
               End If
               '
               JJ& = JJ& + 1
               Call GRAREG("RANKCOS", Z$, JJ&)
             End If
           End If
         End If
       End If
     End If
   Next U&
   '
   Call SETULTREG("RANKCOS", JJ&)
   Call CIERRARCH("RANKCOS")
   MsgBox "Revisar " + Str$(JJ&) + " Valores."
   Call FINAL("*LCONCOS")
   Screen.MousePointer = 1
   '
   Command7.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Openforms = DoEvents
   Wend
End Sub

Private Sub Form_Load()
   Call APLICACIONSKINS(Me)
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Function VULTICOM#(CIXI%, NPRX%)
   '
   Static URE&
   If URE& < 1 Then
     URE& = ULTREG&("OCOMDETA")
   End If
   '
   NPRX% = 0
   VULTICOM# = 0
   RFF$ = RECFILT$("OCOMDETA", 5, MKI$(CIXI%))
   For KKI% = 1 To Len(RFF$) Step 4
     REOCO& = CVS(Mid$(RFF$, KKI%, 4))
     If REOCO& > 0 Then
       If REOCO& <= URE& Then
         XX$ = REGLEIDO$("OCOMDETA", REOCO&)
         STAT% = Asc(Mid$(XX$, 124, 1))
         If STAT% <> 9 Then           ' no esta anulada
           NUORSE& = CVS(Mid$(XX$, 117, 4))
           If NUORSE& < 1 Then ' excluye orden de servicios
                'Z$ = Space$(64)
                'FEX = CVI(Mid$(XX$, 5, 2))
                'Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
             NPXX% = CVI(Mid$(XX$, 7, 2))
                'RPP$ = ECOARCH$("ACREDOR", MKI$(NPXX%))
                'If RPP$ <> "" Then
                '  Call REPLA(Z$, RPP$, 11, 12)
                'End If
             COEFUNI = CVS(Mid$(XX$, 63, 4))
             If COEFUNI < 0.0001 Then COEFUNI = 1
             IUNI# = CVD(Mid$(XX$, 79, 8)) / COEFUNI
             MONEX$ = "$"
                'Call REPLA(Z$, "$", 25, 1)
                'Call REPLA(Z$, CSTRING$(MKD$(IUNI#), 3, 10, 4, 2), 27, 10)
             If TRIM$(Mid$(XX$, 43, 12)) <> "" Then
               MONEX$ = TRIM$(Mid$(XX$, 43, 2))
               IUNI# = XVALO(Mid$(XX$, 45, 10))
             End If
             If MONEX$ <> "$" Then
               CIXI% = CVI(Mid$(XX$, 9, 2))
               IUNI# = IUNI# * TICAMONE(2) * (1 + PORNACIO(CIXI%) / 100)
             End If
             If IUNI# > 0.00005 Then
               NPRX% = NPXX%
               VULTICOM# = IUNI#
             End If
           End If
         End If
       End If
     End If
   Next KKI%
   '

End Function

