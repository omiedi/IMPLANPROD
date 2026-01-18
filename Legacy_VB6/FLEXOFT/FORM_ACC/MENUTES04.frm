VERSION 5.00
Begin VB.Form MENUTES04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tesorería - Gestión de Valores"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5565
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4515
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5145
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0E0&
      Caption         =   "Sub-Sistema de Tesorería"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5760
      Left            =   210
      TabIndex        =   4
      Top             =   210
      Width           =   4110
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0E0&
         Caption         =   "Otros Movimientos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   420
         TabIndex        =   16
         Top             =   2415
         Width           =   3270
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   17
            Top             =   945
            Width           =   2325
            Begin VB.CommandButton Command3 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   0
               Picture         =   "MENUTES04.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "Transferencia de Fondos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Transferencia de Fondos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   840
               TabIndex        =   19
               Top             =   105
               Width           =   1380
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   20
            Top             =   420
            Width           =   2325
            Begin VB.CommandButton Command8 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   0
               Picture         =   "MENUTES04.frx":014A
               Style           =   1  'Graphical
               TabIndex        =   21
               ToolTipText     =   "Depósito de Valores"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Depósito Cuen- ta Bancaria"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   840
               TabIndex        =   22
               Top             =   50
               Width           =   1380
            End
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0E0&
         Caption         =   "Gestión de Valores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   420
         TabIndex        =   12
         Top             =   4305
         Width           =   3270
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   13
            Top             =   420
            Width           =   2220
            Begin VB.CommandButton Command14 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   0
               Picture         =   "MENUTES04.frx":0454
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Gestión de Cheques y Valores en Cartera"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Gestión de Che- ques y Valores"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   15
               Top             =   105
               Width           =   1485
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0E0&
         Caption         =   "Cobranzas y Pagos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   420
         TabIndex        =   5
         Top             =   525
         Width           =   3270
         Begin VB.PictureBox Picture14 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   8
            Top             =   945
            Width           =   2325
            Begin VB.CommandButton Command22 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   0
               Picture         =   "MENUTES04.frx":075E
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Emitir Ordenes de Pago a Proveedores"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Orden de Pago Cuenta Corriente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   9
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture16 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   10
            Top             =   420
            Width           =   2325
            Begin VB.CommandButton Command7 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   0
               Picture         =   "MENUTES04.frx":0A68
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Generar Recibos de Cobranza"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "Recibo Cobranza"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   11
               Top             =   105
               Width           =   1590
            End
         End
      End
   End
   Begin VB.CommandButton Command1 
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
      Height          =   800
      Left            =   4515
      Picture         =   "MENUTES04.frx":0BB2
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   4515
      Picture         =   "MENUTES04.frx":0CFC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1155
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3990
      Picture         =   "MENUTES04.frx":1006
      Top             =   5440
      Width           =   2010
   End
End
Attribute VB_Name = "MENUTES04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Sub Command14_Click()
   Command14.Enabled = False
   GESVAL04.Show 1
   Command14.Enabled = True
End Sub


Sub Command22_Click()
    Command22.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("ACRE_GES/ORPAGO", "Gestión de Compras y Acreedores")
    Command22.Enabled = True
End Sub

Sub Command3_Click()
   Command3.Enabled = False
   Call TRANFONDOS
   Command3.Enabled = True
End Sub

Sub Command7_Click()
   Command7.Enabled = False
   Call CIERRARCH("*.*")
   Call CONECRUN("VENT_GES/RECIBO", "Gestión de Ventas y Deudores")
   'RECIBO00.Show 1
   Command7.Enabled = True
End Sub

Sub Command8_Click()
   Command8.Enabled = False
   Call ASIDECHE
   Command8.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub


Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Sub ASIDECHE()
   '
   HOII% = HOY(HOS$)
   CUECHECA$ = CONTROL$("", "CUECHECA")
   CUEICHEQ% = CUEINT%(CUECHECA$)
   If TRIM$(CUECHECA$) = "" Or CUEICHEQ% < 1 Or CUEICHEQ% > NUCUEN% Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Cartera de Valores")
      OPCHEQ04.Show 1
      Exit Sub
   End If
   '
   If ULTREG&("CUEBANC") < 1 Then
      Call MENSERR(24, "Falta Definir Cuentas Bancarias\para Deposito de Cheques.")
      Call CARCUEBANC
      Exit Sub
   End If
   '
5  VDEF$ = Space$(80)
   Call REPLA(VDEF$, MKI$(HOII%), 13, 2)
   Call REPLA(VDEF$, "DEPOSITO CHEQUES ", 15, 40)
   Call REPLA(VDEF$, CUECHECA$, 55, 12)
   Call REPLA(VDEF$, "48 Hs", 77, 10)
10 OBX$ = OBJPLA$("DEPOCHECAR", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   VDEF$ = OBX$
   CUEBA$ = Left$(OBX$, 12)
   If CUEINT%(CUEBA$) < 1 Or CUEINT(CUEBA$) > NUCUEN% Then GoTo 10
   DESTICHE$ = ECOARCH$("CUEBANC", CUEBA$)
   If DESTICHE$ = "" Then
     Call MENSERR(24, "Cuenta " + CUEBA$ + " no Definida\como Cuenta Bancaria")
     GoTo 10
   End If
   '
   CUECHE$ = Mid$(OBX$, 55, 12)
   FEDEPO% = CVI(Mid$(OBX$, 13, 2))
   If FEDEPO% < COMEJ% Or FEDEPO% > FINEJ% Then GoTo 10
   PLADEPO$ = TRIM$(Mid$(OBX$, 77, 10))
   HORACRE% = Val(Left$(PLADEPO$, 3))
   If HORACRE% < 24 Then
     If Left$(PLADEPO$, 2) <> "0 " Then
       Call MENSERR(24, "Plazo de Acreditación no Válido.")
       GoTo 10
     End If
   End If
   If HORACRE% Mod 24 <> 0 Then
     Call MENSERR(24, "Plazo de Acreditación no Válido.")
     GoTo 10
   End If
   '
   Screen.MousePointer = 11
   FEX = FEDEPO%
   REFE$ = TRIM$(Mid$(OBX$, 15, 40)) + " " + FECHATEX$(FEX)
   Call GRACONTROL("", "CUECHE", CUECHE$)
   '
   SELCHE04.List1.Clear
   SELCHE04.LICHESEL.Clear
   RFF$ = RECFILT$("CAJABANC", 16, Chr$(1))
   JJ& = 0
   FIN& = Len(RFF$)
   FORMALTER.LINOSORT.Clear
   '
   For I& = 1 To FIN& Step 4
     Call TRACE(20, I&, FIN&)
     RECHE& = CVS(Mid$(RFF$, I&, 4))
     X$ = REGLEIDO$("CAJABANC", RECHE&)
     '
     If UCase$(Mid$(X$, 69, 2)) <> "CH" Then  ' no es un cheque
       GoTo 19
     End If
     '
     If CVI(Mid$(X$, 269, 2)) > 0 Then ' esta anulado
       Call REPLA(X$, Chr$(9), 121, 1)
       Call GRAREG("CAJABANC", X$, RECHE&)
       GoTo 19
     End If
     '
     If CVI(Mid$(X$, 275, 2)) > 0 Then ' salio de cartera
       Call REPLA(X$, Chr$(2), 121, 1)
       Call GRAREG("CAJABANC", X$, RECHE&)
       GoTo 19
     End If
     '
     FEACRED% = CVI(Mid$(X$, 63, 2))
     BANUFE$ = Mid$(X$, 71, 40)
     ICHEQ# = CVD(Mid$(X$, 153, 8))
     '
     If TRIM$(Left$(BANUFE$, 12)) = "" Then
        Call REPLA(BANUFE$, "Cheque Nro:", 1, 11)
     End If
     Z$ = "A" + Space$(80)
     Call REPLA(Z$, CSTRING$(MKI$(FEACRED%), 1, 5, 0, 2), 2, 5)
     Call REPLA(Z$, TRIM$(Left$(BANUFE$, 40)), 9, 40)
     Call REPLA(Z$, CSTRING$(MKD$(ICHEQ#), 4, 12, 2, 2), 49, 12)
     Call REPLA(Z$, CSTRING$(MKS$(RECHE&), 3, 8, 0, 2), 73, 8)
     SELCHE04.List1.AddItem Z$
19 Next I&
   '
   SUCHETO# = 0
   For I& = 1 To SELCHE04.List1.ListCount
      Z$ = Mid$(SELCHE04.List1.List(I& - 1), 9)
      SELCHE04.LICHESEL.AddItem Z$
      SUCHETO# = SUCHETO# + Val(Mid$(Z$, 41, 12))
   Next I&
   '
   SELCHE04.Label4.Caption = CSTRING$(MKD$(SUCHETO#), 4, 11, 2, 2)
   SELCHE04.Label4.Refresh
   '
   Screen.MousePointer = 1
   SELCHE04.Show 1
   '
   IPASE# = Val(SELCHE04.Label2.Caption)
   FEACR% = FECHACRE%(FEDEPO%, HORACRE%)
   If IPASE# > 0.005 Then
      FEX = FEDEPO%
      '
      Z$ = REGBLAN$("CAJABANC")
      Call REPLA(Z$, CUEBA$, 1, 12)
      Call REPLA(Z$, "DEPOSITO " + PLADEPO$, 14, 27)
      Call REPLA(Z$, CSTRING$(MKD$(IPASE#), 3, 12, 2, 2), 41, 12)
      PASDEBE$ = Z$
      'List1.AddItem Z$
      '
      FEX = FEDEPO%
      FETX$ = FECHATEX$(FEX)
      NRECIB$ = "DB-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
      ORMA% = (-1)
      RFF$ = RECFILT$("CAJABANC", 3, MKI$(FEDEPO%))
      For I& = 1 To Len(RFF$) Step 4
        RECAJ& = CVS(Mid$(RFF$, I&, 4))
        XY$ = REGLEIDO$("CAJABANC", RECAJ&)
        If Left$(XY$, 10) = NRECIB$ Then
          ORMB% = Val(Mid$(XY$, 11, 2))
          If ORMB% > ORMA% Then ORMA% = ORMB%
        End If
      Next I&
      ORMX$ = TRIM$(Str$(ORMA% + 1))
      While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
      NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
      '
      Z$ = REGBLAN$("CAJABANC")
      Call REPLA(Z$, NRECIB$, 1, 12)
      Call REPLA(Z$, REFE$, 13, 48)
      Call REPLA(Z$, MKI$(FEDEPO%), 61, 2)
      Call REPLA(Z$, MKI$(FEACR%), 63, 2)
      Call REPLA(Z$, MKI$(CUEINT%(CUEBA$)), 65, 2)
      Call REPLA(Z$, MKI$(0), 67, 2)
      Call REPLA(Z$, "DP", 69, 2)
      Call REPLA(Z$, "Depósito Cheques " + PLADEPO$ + " hs", 71, 40)
      Call REPLA(Z$, MKI$(FEDEPO%), 111, 2)
      Call REPLA(Z$, MKI$(0), 115, 2)
      Call REPLA(Z$, MKS$(0), 117, 4)
      Call REPLA(Z$, Chr$(0), 121, 1)
      Call REPLA(Z$, MKD$(IPASE#), 145, 8)
      Call REPLA(Z$, MKD$(IPASE#), 153, 8)
      Call REPLA(Z$, "CAJABANC", 245, 8)
      Call REPLA(Z$, MKS$(0), 253, 4)
      Call REPLA(Z$, IDACTU$, 257, 6)  ' FECHA - HORA - USUARIO - TERMINAL
      Call REGAPP("CAJABANC", Z$)
      '
      SUMACHESEL# = 0
      FEX = FEDEPO%: HOS$ = FECHATEX$(FEX)
      For I& = 1 To SELCHE04.LICHESEL.ListCount
        If SELCHE04.LICHESEL.Selected(I& - 1) = True Then
          X$ = SELCHE04.LICHESEL.List(I& - 1)
          BACH$ = TRIM$(Left$(X$, 16))
          NCHE$ = Mid$(X$, 19, 12)
          IPASE# = Val(Mid$(X$, 41, 12))
          SUMACHESEL# = SUMACHESEL# + IPASE#
          REGBA& = Val(Mid$(X$, 65, 8))
          '
          If REGBA& > 0 Then ' es un cheque de cartera
             '
             TTS$ = REGLEIDO$("CAJABANC", REGBA&)
             Call REPLA(TTS$, MKI$(FEDEPO%), 275, 2)
             Call REPLA(TTS$, NRECIB$, 277, 16)
             Call REPLA(TTS$, DESTICHE$, 293, 30)
             Call REPLA(TTS$, REFE$, 323, 40)
             Call REPLA(TTS$, HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 363, 40)
             Call REPLA(TTS$, "Usuario: " + TRIM$(USERNAME$), 403, 40)
             Call REPLA(TTS$, Chr$(2), 121, 1)  ' Marca Salida de Cartera
             Call REPLA(TTS$, IDACTU$, 263, 6)  ' Marca Modificacion de registro - FECHA - HORA - USUARIO - TERMINAL
             Call GRAREG("CAJABANC", TTS$, REGBA&)
             '
             Z$ = TTS$
             Call REPLA(Z$, Mid$(REBLA$, 209), 177, 336)
             IVAX$ = Mid$(Z$, 71, 40)
             IDEVA$ = TRIM$(Left$(IVAX$, 18)) + " " + TRIM$(Mid$(IVAX$, 19, 14)) + " " + TRIM$(Mid$(IVAX$, 33))
             '
             Call REPLA(Z$, NRECIB$, 1, 12)
                 FEX = FEDEPO%
             Call REPLA(Z$, REFE$, 13, 48)
             Call REPLA(Z$, MKI$(FEDEPO%), 61, 2)
             Call REPLA(Z$, "DP", 69, 2)
             Call REPLA(Z$, MKI$(0), 115, 2)
             Call REPLA(Z$, MKS$(NRO&), 117, 4)
                 STATUI% = 0
             Call REPLA(Z$, Chr$(STATUI%), 121, 1) ' STATUS%=1 ES CHEQUE EN CARTERA
             Call REPLA(Z$, MKS$(0), 125, 4)
             Call REPLA(Z$, String$(16, 0), 129, 16)
             Call REPLA(Z$, MKD$(0), 153, 8)
             Call REPLA(Z$, MKD$(IPASE#), 161, 8)
             Call REPLA(Z$, MKD$(0), 169, 8)
             Call REPLA(Z$, IDEVA$, 177, 32)
             Call REPLA(Z$, "CAJABANC", 245, 8)
             Call REPLA(Z$, MKS$(REGBA&), 253, 4)
             Call REPLA(Z$, IDACTU$, 257, 6)  ' FECHA - HORA - USUARIO - TERMINAL
             Call REGAPP("CAJABANC", Z$)
             '
          End If
        End If
      Next I&
   End If
   '
   Call CIERRARCH("CAJABANC")
   Call BAJALDISCO
   Screen.MousePointer = 1
   '
   GoTo 5
   '
End Sub

Sub TRANFONDOS()
   '
   HOII% = HOY(HOS$)
   CUECHECA$ = CONTROL$("", "CUECHECA")
   CUEICHEQ% = CUECOINT%(CUECHECA$)
   If TRIM$(CUECHECA$) = "" Or CUEICHEQ% < 1 Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Cartera de Valores")
      OPCHEQ04.Show 1
      Exit Sub
   End If
   '
   If ULTREG&("CUEBANC") < 1 Then
      Call MENSERR(24, "Falta Definir Cuentas Bancarias\para Deposito de Cheques.")
      Call CARCUEBANC
      Exit Sub
   End If
   '
5  VDEF$ = Space$(96)
   Call REPLA(VDEF$, MKI$(HOII%), 1, 2)
   Call REPLA(VDEF$, "TRANSFERENCIA DE FONDOS ", 35, 40)
   '
10 OBX$ = OBJPLA$("TRANFONDOS", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   VDEF$ = OBX$
   CUEDEBE$ = TRIM$(Mid$(OBX$, 23, 12))
   CUDEB1% = CUECOINT%(CUEDEBE$)
   If CUEDEBE$ = "" Or CUDEB1% < 1 Or CUDEB1% = CUEICHEQ% Then
     Call MENSERR(24, "Cuenta no Válida")
     GoTo 10
   End If
   '
   CUEHABE$ = TRIM$(Mid$(OBX$, 11, 12))
   CUHAB1% = CUECOINT%(CUEHABE$)
   If CUEHABE$ = "" Or CUHAB1% < 1 Or CUHAB1% = CUEICHEQ% Then
     Call MENSERR(24, "Cuenta no Válida")
     GoTo 10
   End If
   '
   FEDEP% = CVI(Mid$(OBX$, 1, 2))
   If FEDEP% < 1 Then
     Call MENSERR(24, "Fecha no Válida")
     GoTo 10
   End If
   '
   ITRAN# = CVD(Mid$(OBX$, 3, 8))
   If ITRAN# < 0.005 Then
     Call MENSERR(24, "Falta Importe")
     GoTo 10
   End If
   '
   If TRIM$(Mid$(OBX$, 35, 40)) = "" Then
     Call MENSERR(24, "Falta Referencia")
     GoTo 10
   End If
   REFE$ = TRIM$(Mid$(OBX$, 35, 40)) + " " + FECHATEX$(FEX)
   '
   Screen.MousePointer = 11
   FEX = FEDEP%
   FETX$ = FECHATEX$(FEX)
   NRECIB$ = "TF-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
   ORMA% = (-1)
   RFF$ = RECFILT$("CAJABANC", 3, MKI$(FEDEP%))
   For I& = 1 To Len(RFF$) Step 4
     RECAJ& = CVS(Mid$(RFF$, I&, 4))
     XY$ = REGLEIDO$("CAJABANC", RECAJ&)
     If Left$(XY$, 10) = NRECIB$ Then
       ORMB% = Val(Mid$(XY$, 11, 2))
       If ORMB% > ORMA% Then ORMA% = ORMB%
     End If
   Next I&
   ORMX$ = TRIM$(Str$(ORMA% + 1))
   While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
   NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
   '
   Z$ = REGBLAN$("CAJABANC")
   Call REPLA(Z$, NRECIB$, 1, 12)
   Call REPLA(Z$, REFE$, 13, 48)
   Call REPLA(Z$, MKI$(FEDEP%), 61, 2)
   Call REPLA(Z$, MKI$(FEDEP%), 63, 2)
   Call REPLA(Z$, MKI$(CUDEB1%), 65, 2)
   Call REPLA(Z$, MKI$(0), 67, 2)
   Call REPLA(Z$, "TF", 69, 2)
   Call REPLA(Z$, "De " + TRIM$(ECOARCH$("ECUECON", CUEHABE$)), 71, 40)
   Call REPLA(Z$, MKI$(FEDEP%), 111, 2)
   Call REPLA(Z$, MKI$(0), 115, 2)
   Call REPLA(Z$, MKS$(0), 117, 4)
   Call REPLA(Z$, Chr$(0), 121, 1)
   Call REPLA(Z$, MKD$(ITRAN#), 145, 8)
   Call REPLA(Z$, MKD$(ITRAN#), 153, 8)
   Call REPLA(Z$, "CAJABANC", 245, 8)
   Call REPLA(Z$, MKS$(0), 253, 4)
   Call REPLA(Z$, IDACTU$, 257, 6)  ' FECHA - HORA - USUARIO - TERMINAL
   Call REGAPP("CAJABANC", Z$)
   '
   Call REPLA(Z$, MKI$(CUHAB1%), 65, 2)
   Call REPLA(Z$, MKI$(0), 67, 2)
   Call REPLA(Z$, "TF", 69, 2)
   Call REPLA(Z$, "A " + TRIM$(ECOARCH$("ECUECON", CUEDEBE$)), 71, 40)
   Call REPLA(Z$, MKD$(0), 153, 8)
   Call REPLA(Z$, MKD$(ITRAN#), 161, 8)
   Call REGAPP("CAJABANC", Z$)
   '
   Call CIERRARCH("CAJABANC")
   Screen.MousePointer = 1
   Call COMUNI("Transferencia Completa y Verificada")
   GoTo 5
   '
End Sub
