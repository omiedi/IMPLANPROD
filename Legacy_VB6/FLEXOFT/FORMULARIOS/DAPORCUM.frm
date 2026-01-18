VERSION 5.00
Begin VB.Form DAPORCUM 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Compra Pendientes - Marcar Items Cumplidos"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11745
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox LICOPEN 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      ToolTipText     =   "Tilde los Items que Dará por Cumplidos"
      Top             =   2040
      Width           =   10545
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   10485
      TabIndex        =   8
      Top             =   1630
      Width           =   10545
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Código           Descripción                    U/S      Pedido     Recep.    Saldo"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   225
         Left            =   220
         TabIndex        =   9
         ToolTipText     =   "Haga Doble-Click para Blanquear Pre-Carga y Habilitar Carga Manual"
         Top             =   105
         Width           =   10515
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   5775
      Left            =   10800
      ScaleHeight     =   5715
      ScaleWidth      =   1035
      TabIndex        =   3
      Top             =   0
      Width           =   1095
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "DAPORCUM.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   4080
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   5
         Top             =   3720
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton BOTREC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "DAPORCUM.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aprueba la Selección de Items Cumplidos"
         Top             =   4755
         Width           =   650
      End
      Begin VB.CommandButton BOTEXIT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "DAPORCUM.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "DAPORCUM.frx":075E
         Top             =   1620
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   10680
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   10680
      Y1              =   5415
      Y2              =   5415
   End
   Begin VB.Line Line2 
      X1              =   10660
      X2              =   10660
      Y1              =   1650
      Y2              =   5400
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   7440
      TabIndex        =   25
      Top             =   1140
      Width           =   1695
   End
   Begin VB.Label Label15 
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
      Left            =   9560
      TabIndex        =   24
      Top             =   450
      Width           =   1095
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
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
      Left            =   9680
      TabIndex        =   23
      Top             =   1110
      Width           =   975
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Número de O/C:"
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
      Left            =   8160
      TabIndex        =   22
      Top             =   860
      Width           =   1695
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Emisión:"
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
      Left            =   8160
      TabIndex        =   21
      Top             =   195
      Width           =   1695
   End
   Begin VB.Label Label11 
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
      Left            =   5760
      TabIndex        =   20
      Top             =   1110
      Width           =   1695
   End
   Begin VB.Label Label10 
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
      Left            =   5760
      TabIndex        =   19
      Top             =   450
      Width           =   1695
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
      Left            =   1440
      TabIndex        =   18
      Top             =   1110
      Width           =   4095
   End
   Begin VB.Label Label8 
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
      Left            =   2280
      TabIndex        =   17
      Top             =   780
      Width           =   3255
   End
   Begin VB.Label Label7 
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
      Left            =   1440
      TabIndex        =   16
      Top             =   780
      Width           =   735
   End
   Begin VB.Label Label6 
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
      Left            =   1440
      TabIndex        =   15
      Top             =   450
      Width           =   4095
   End
   Begin VB.Label Label4 
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
      Left            =   1440
      TabIndex        =   14
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      Left            =   480
      TabIndex        =   13
      Top             =   1110
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor:"
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
      Left            =   120
      TabIndex        =   12
      Top             =   855
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Posición IVA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   5790
      TabIndex        =   11
      Top             =   180
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C.U.I.T.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   5790
      TabIndex        =   10
      Top             =   855
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7050
      Y1              =   4340
      Y2              =   4340
   End
End
Attribute VB_Name = "DAPORCUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTEXIT_Click()
   LICOPEN.Clear
   Unload Me
End Sub
'
Private Sub BOTREC_Click()
   If COMALTER%("Acepta Dar por Cumplidos\los Items Tildados ?\\No, Cancelar\Si, Dar por Cumplidos") = 2 Then
     Hide
   End If
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   TPI = LICOPEN.TopIndex
   For U& = 1 To LICOPEN.ListCount
     If LICOPEN.Selected(U& - 1) = True Then
         LICOPEN.Selected(U& - 1) = False
       Else
         LICOPEN.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LICOPEN.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   Command1.Enabled = True
End Sub

Private Sub Command9_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Label2 = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Label14_Change()
   NUOCO& = Val(Label14)
   If NUOCO& > 0 Then
     ENCAOC$ = FILTERGETRECORD$("OCOMENCA", 1, MKS$(NUOCO&))
     If CVS(Left$(ENCAOC$, 4)) <> NUOCO& Then
       Call MENSERR(24, "Imposible Presentar O/C Número" + Str$(NUOCO&))
       Label14 = ""
       Hide
       Exit Sub
     End If
     '
     Label15 = FECHATEX$(CVI(Mid$(ENCAOC$, 49, 2)))
     NPRO% = CVI(Mid$(ENCAOC$, 51, 2))
     NCLIE% = (-1) * NPRO%
     ATEN$ = TRIM$(Mid$(ENCAOC$, 163, 30))
     '
     Label3 = NPRO%
     Label4 = RASOCLI$(NCLIE%)
     Label6 = DOMICLI$(NCLIE%)
     Label7 = CPOSCLI$(NCLIE%)
     Label8 = LOCACLI$(NCLIE%)
     Label9 = ATEN$
     '
     Label10 = POSIVAC$(NCLIE%)
     Label11 = CUITCLI$(NCLIE%)
     DoEvents
     '
     RFF$ = RECFILT$("OCOMDETA", 1, MKS$(NUOCO&))
     LICOPEN.Clear
     For U& = 1 To Len(RFF$) Step 4
       REOC& = CVS(Mid$(RFF$, U&, 4))
       X$ = REGLEIDO$("OCOMDETA", REOC&)
       TTXX$ = Space$(128)
       If Asc(Mid$(X$, 124, 1)) < 9 Then ' abierta
         CII% = CVI(Mid$(X$, 9, 2))
         COEFCAN = CVS(Mid$(X$, 63, 4))
           If COEFCAN < 0.005 Then COEFCAN = 1
         CAORI = CVS(Mid$(X$, 55, 4)) * COEFCAN
         CAREC = CVS(Mid$(X$, 99, 4)) * COEFCAN
         CAPEN = CAORI - CAREC: If CAPEN < 0 Then CAPEN = 0
         UNICO$ = Mid$(X$, 11, 4)
         COEFUNI = CVS(Mid$(X$, 63, 4))
         STAT% = Asc(Mid$(X$, 124, 1))
         '
         Call REPLA(TTXX$, CODEXT$(CII%), 1, 16)
         Call REPLA(TTXX$, DESCRIT0$(CII%), 18, 30)
             ' esto es para o/c's de texto libre
             RFF1$ = RECFILT$("TXOCOLIB", 1, MKS$(NUOCO&))
             If Len(RFF1$) >= 4 Then
                For KKL& = 1 To Len(RFF1$) Step 4
                   RETX& = CVS(Mid$(RFF1$, KKL&))
                   XYX$ = REGLEIDO$("TXOCOLIB", RETX&)
                   DECOR$ = TRIM$(Mid$(XYX$, 5, 64))
                   Call REPLA(TTXX$, DECOR$, 18, 30)
                   GoTo 15
                Next KKL&
             End If
15       Call REPLA(TTXX$, UNICO$, 49, 4)
         Call REPLA(TTXX$, FORMATNUM$(CAORI, "F10.1"), 54, 10)
         Call REPLA(TTXX$, FORMATNUM$(CAREC, "F10.1"), 64, 10)
         Call REPLA(TTXX$, FORMATNUM$(CAPEN, "F10.1"), 74, 10)
         '
         If STAT% = 8 Then
           Call REPLA(TTXX$, "Cumplido", 76, 10)
         End If
         Call REPLA(TTXX$, TRIM$(Str$(REOC&)), 121, 8)
         LICOPEN.AddItem TTXX$
         '
       End If
     Next U&
     '
   End If
End Sub

