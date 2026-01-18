VERSION 5.00
Begin VB.Form LISOCAB04 
   BackColor       =   &H00E0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Compra Abiertas - Proveedores Externos"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00006060&
      Height          =   4215
      Left            =   8700
      ScaleHeight     =   4155
      ScaleWidth      =   1005
      TabIndex        =   29
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   140
         ScaleHeight     =   75
         ScaleWidth      =   495
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2520
         Width           =   555
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -105
            TabIndex        =   35
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   420
         Picture         =   "LISOCAB04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   1890
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Direct"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "LISOCAB04.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Acceso Directo a Otros Procesos"
         Top             =   1680
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "LISOCAB04.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   1005
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "LISOCAB04.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -255
         Picture         =   "LISOCAB04.frx":0BA0
         Stretch         =   -1  'True
         Top             =   2730
         Width           =   1410
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Consultas Varias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   4410
      TabIndex        =   3
      Top             =   210
      Width           =   4005
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0FFFF&
         Caption         =   "Salida"
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
         Left            =   210
         TabIndex        =   36
         Top             =   2100
         Width           =   3585
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Impresión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1995
            TabIndex        =   38
            Top             =   250
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0FFFF&
            Caption         =   "Pantalla"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   420
            TabIndex        =   37
            Top             =   250
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0FFFF&
         Caption         =   "Seguimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   210
         TabIndex        =   22
         Top             =   240
         Width           =   3585
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   26
            Top             =   840
            Width           =   2220
            Begin VB.CommandButton Command2 
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
               Left            =   -20
               Picture         =   "LISOCAB04.frx":2AC2
               Style           =   1  'Graphical
               TabIndex        =   27
               ToolTipText     =   "Aprobar Ordenes de Compra"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Aprobar Orde- nes de Compra"
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
               TabIndex        =   28
               Top             =   105
               Width           =   1485
            End
         End
         Begin VB.PictureBox Picture14 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   23
            Top             =   405
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
               Picture         =   "LISOCAB04.frx":2DCC
               Style           =   1  'Graphical
               TabIndex        =   24
               ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Ver y Re-Imprimir"
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
               TabIndex        =   25
               Top             =   80
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0FFFF&
         Caption         =   "Control de Ultimas Compras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   210
         TabIndex        =   4
         Top             =   2730
         Visible         =   0   'False
         Width           =   3585
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   5
            Top             =   735
            Width           =   2220
            Begin VB.CommandButton Command12 
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
               Picture         =   "LISOCAB04.frx":30D6
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Ultimas Compras por Proveedor"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Control por Proveedor"
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
               TabIndex        =   6
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture12 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   7
            Top             =   420
            Width           =   2220
            Begin VB.CommandButton Command20 
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
               Left            =   -20
               Picture         =   "LISOCAB04.frx":33E0
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Ultimas Compras por Material"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Por Artículo"
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
               TabIndex        =   8
               Top             =   50
               Width           =   1590
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Control de Compras en Curso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   4005
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0FFFF&
         Caption         =   "Opciones de Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Left            =   210
         TabIndex        =   9
         Top             =   420
         Width           =   3585
         Begin VB.PictureBox Picture11 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   10
            Top             =   2205
            Visible         =   0   'False
            Width           =   2325
            Begin VB.CommandButton Command9 
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
               Picture         =   "LISOCAB04.frx":352A
               Style           =   1  'Graphical
               TabIndex        =   11
               ToolTipText     =   "Material en Consignación"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Material en Consignación"
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
               TabIndex        =   12
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   13
            Top             =   1365
            Width           =   2325
            Begin VB.CommandButton Command5 
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
               Left            =   -20
               Picture         =   "LISOCAB04.frx":3834
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Pendientes por Material"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Pendientes Por Material"
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
               TabIndex        =   15
               Top             =   105
               Width           =   1200
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   16
            Top             =   945
            Width           =   2325
            Begin VB.CommandButton Command4 
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
               Left            =   -20
               Picture         =   "LISOCAB04.frx":3B3E
               Style           =   1  'Graphical
               TabIndex        =   17
               ToolTipText     =   "Pendientes por Proveedor"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Por Proveedor"
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
               TabIndex        =   18
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   19
            Top             =   525
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
               Picture         =   "LISOCAB04.frx":3F80
               Style           =   1  'Graphical
               TabIndex        =   20
               ToolTipText     =   "Ordenes de Compra por Número"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "O/C's por Número"
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
               TabIndex        =   21
               Top             =   105
               Width           =   1590
            End
         End
      End
   End
End
Attribute VB_Name = "LISOCAB04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HAYMATERIALIB%

'*****/////COMANDO OCULTO/////*****
'Se oculto en el frame8 ("Opciones de Consulta")
'el Command9("Listado de Material en Consignación)"
'*****/////**************/////*****

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

 Sub Command12_Click()
    Command12.Enabled = False
    MODO% = 0
    Call SELCOMRAFE(OKKX%, MODO%)
    '
    If HAYMATERIALIB% = 1 Then
        If OKKX% = 1 Then Call CONECRUN("*LICOMPL", "Listado de Compras por Proveedor")
      Else
        If OKKX% = 1 Then Call CONECRUN("*LICOMPR", "Listado de Compras por Proveedor")
    End If
    '
    Screen.MousePointer = 1
99  Command12.Enabled = True
End Sub

Private Sub Command17_Click()
    '
    Command17.Enabled = False
    Screen.MousePointer = 11
    J& = 0: K& = 0: OCOA& = 0
    '
    For I& = 1 To ULTREG&("OCOMPRA")
         XX$ = REGLEIDO$("OCOMPRA", I&)
         '
         If CVS(Left(XX$, 4)) >= 0.5 Then    ' depura O/C's
             K& = K& + 1                     ' sin numero
             Call GRAREG("OCOMPRA", XX$, K&)
         End If
         '
         ST% = Asc(Mid$(XX$, 62, 1))
         '
         If ST% < 1 Then     '    excluye cerradas
            If CVS(Left$(XX$, 4)) <> OCOA& Then
               NP% = CVI(Mid$(XX$, 7, 2))
               If NP% < 1 Or NP% = 9999 Then
                  J& = J& + 1
                  Call GRAREG("!OCOMPRA", XX$, J&)
                  OCOA& = CVS(Left$(XX$, 4))
               End If
            End If
         End If
    Next I&
    '
    Call SETULTREG("OCOMPRA", K&)
    Call SETULTREG("!OCOMPRA", J&)
    Call CIERRARCH("!OCOMPRA")
    '
    If J& < 1 Then
          Call MENSERR(24, "No hay Pedidos de Suministro\Pendientes que Puedan Asignarse.")
          Screen.MousePointer = 1
          GoTo 99
    End If
    '
    INDI% = 1
    Screen.MousePointer = 1
    '
    XX$ = REGSEL$("!OCOMPRA")
    If Len(XX$) <= 4 Then GoTo 99
    '
    NUOCO& = CVS(Left$(XX$, 4))
    '
    YY$ = REGSEL$("PROVED1")
    If Len(YY$) <= 4 Then GoTo 99
    NP1% = CVI(Left$(YY$, 2))
    PROVEX$ = Mid$(YY$, 3, 30)
    '
    Screen.MousePointer = 11
    JJ& = 0
    FIN& = ULTREG&("OCOMPRA")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("OCOMPRA", U&)
       '
       If CVS(Left$(XX$, 4)) = NUOCO& Then
          CI0% = CVI(Mid$(XX$, 11, 2))
          CAOR = CVS(Mid$(XX$, 13, 4))
          FENTRE% = CVI(Mid$(XX$, 9, 2))
          YY$ = REGBLAN$("CARDETOC")
          Call REPLA(YY$, MKI$(CI0%), 1, 2)
          Call REPLA(YY$, UNIMED$(CI0%), 3, 4)
          Call REPLA(YY$, MKS$(CAOR), 47, 4)
          Call REPLA(YY$, MKI$(FENTRE%), 67, 2)
          Call REPLA(YY$, MKS$(U&), 71, 4)
          JJ& = JJ& + 1
          Call GRAREG("!CARDETOC", YY$, JJ&)
       End If
    Next U&
    '
    Call SETULTREG("!CARDETOC", JJ&)
    Call CIERRARCH("!CARDETOC")
    Screen.MousePointer = 1
    '
    VTB% = VENTABU%("CAROCOM2/TITUPAN=P.S.Nro " + TRIM$(Str$(NUOCO&)) + " - " + TRIM$(RASOCLI$((-1) * NP1%)))
    If VTB% < 0 Then GoTo 99
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("OCOMPRA")
    For U& = 1 To ULTREG&("!CARDETOC")
       YY$ = REGLEIDO$("!CARDETOC", U&)
       CAOR = CVS(Mid$(YY$, 47, 4))
       REOCO& = CVS(Mid$(YY$, 71, 4))
       If REOCO& > 0 Then
          If REOCO& <= FIN& Then
             XX$ = REGLEIDO$("OCOMPRA", REOCO&)
             If CVS(Left$(XX$, 4)) = NUOCO& Then
                Call REPLA(XX$, MKI$(NP1%), 7, 2)
                Call REPLA(XX$, MKS$(CAOR), 13, 4)
                Call GRAREG("OCOMPRA", XX$, REOCO&)
             End If
          End If
       End If
    Next U&
    '
    FIN& = ULTREG&("MACONSIG")
    For U& = 1 To FIN&
       XX$ = REGLEIDO$("MACONSIG", U&)
       If Val(Mid$(XX$, 24, 6)) = NUOCO& Then
          Call REPLA(XX$, MKI$(NP1%), 3, 2)
          Call GRAREG("MACONSIG", XX$, U&)
       End If
    Next U&
    '
    Call CIERRARCH("*.*")
    Call FRESHECHO("!OCOMPRA")
    Screen.MousePointer = 1
    '
99  Command17.Enabled = True
    '
End Sub

Private Sub Command18_Click()
    '
    Command18.Enabled = False
    'SETUPCOM.Show 1
    Call CONECRUN("FLSETUP/SETUPCOM", "Configuración de Funcionamiento")
    Command18.Enabled = True
    '
End Sub


Private Sub Command2_Click()
    Command2.Enabled = False
    FAPOCAB04.Show 1
    Command2.Enabled = True
End Sub

 Sub Command20_Click()
    Command20.Enabled = False
    MODO% = 1
    Call SELCOMRAFE(OKKX%, MODO%)
'    If OKKX% = 1 Then Call CONECRUN("*LICOMAR", "Listado de Compras por Articulo")
    'HAY QUE DERIVAR PARA CADDA LISTADO SEGUN SI ES MAT. TEXTO LIBRE
    If HAYMATERIALIB = 1 Then
        If OKKX% = 1 Then Call CONECRUN("*LICOMAL", "Listado de Compras por Articulo")
    Else
        If OKKX% = 1 Then Call CONECRUN("*LICOMAR", "Listado de Compras por Articulo")
    End If
    Screen.MousePointer = 1
    Command20.Enabled = True
End Sub

Private Sub Command21_Click()
   Command21.Enabled = False
   Screen.MousePointer = 11
   JJ& = 0
   FIN& = ULTREG&("OCOMPRA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("OCOMPRA", U&)
      '
      ST% = Asc(Mid$(XX$, 62, 1))
      If ST% < 1 Then              '    excluye cerradas
         NP% = CVI(Mid$(XX$, 7, 2))
         If NP% < 1 Or NP% = 9999 Then
            CAPEN = CVS(Mid$(XX$, 13, 4)) - CVS(Mid$(XX$, 58, 4))
            If CAPEN > 0 Then
               JJ& = JJ& + 1
               Call GRAREG("!OCOMPRA", XX$, JJ&)
            End If
         End If
      End If
   Next U&
   '
   Call SETULTREG("!OCOMPRA", JJ&)
   Call CIERRARCH("!OCOMPRA")
   '
   Screen.MousePointer = 1
   If JJ& < 1 Then
      Call MENSERR(24, "No Hay Pedidos de Suministro Pendientes.")
      GoTo 99
   End If
   '
   VTB% = VENTABU%("PEDSUPEN")
   '
99 Command21.Enabled = True
   '
End Sub

 Sub Command22_Click()
    '
    Command22.Enabled = False
    'Cargo la lista con las O/C
    Call ABRECOMPRAS
    SQLX$ = " SELECT Nume_Ocom, Refe_Ocom FROM ENCOCABI"
    SQLX$ = SQLX$ + " ORDER BY Nume_Ocom "
    Set VRIMPTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    With VRIMPTEMP
        On Error Resume Next ' si hay algun error msgbox err
        .MoveFirst
        If Err Then
          On Error GoTo 0
          Call MENSERR(24, "No Aparecen Ordenes de Compra Emitidas")
          GoTo 99
        End If
        On Error GoTo 0
        JJ& = 0
        Do While (.EOF = False)
            XXY$ = REGBLAN$("INDIOCOM")
            Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4)
            Call REPLA(XXY$, !Refe_OcoM, 5, 44)
            JJ& = JJ& + 1
            Call GRAREG("!INDIOCOM", XXY$, JJ&)
            .MoveNext
        Loop
    End With
    Set VRIMPTEMP = Nothing
    Call SETULTREG("!INDIOCOM", JJ&)
    Call CIERRARCH("!INDIOCOM")
    '
    Call FRESHECHO("!INDIOCOM")
3   Call SELECHO("!INDIOCOM/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
    NPX& = Val(VALACT1$("!INDIOCOM")) ' ACA LEO DEL INDIOCOM
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Nota de Pedido - Abierta"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU1$ = " " + NUDOCU$ + " "
    NUDOCU2$ = " " + NUDOCU$ + "-"
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    GoTo 99
    '
5   For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    GoTo 99
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    FIN& = Len(TPSP$)
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7070
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7070
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    UI& = UI& + 1
    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
    If InStr(XX$, TIDOCUM$) > 4 Then
      ' If InStr(XX$, NUDOCU$) > 4 Then
      If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.HScroll1.Min = 0
    EDITFORM.HScroll1.Value = 0
    EDITFORM.HScroll1.Max = 1
    If ANCHPAGINA > EDITFORM.Frame1.Width Then
      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command1.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 3
    '
99  Command22.Enabled = True
    '
90 End Sub
Sub Command3_Click()
    OPXX% = ALTERNA%("Base de Datos Completa\Ordenes Pendientes")
    If Option1.Value = True Then 'PANTALLA
        Select Case OPXX%
          Case 1
            Call GENINDOC
            Call COCOMP0 ' SUB DE BdDATOS COMPLET.
            '
            If HAYMATERIALIB = 1 Then
                Call COMUNI("La Consulta por Pantalla\Permite Visualizar Solamente\Materiales CODIFICADOS.")
            End If
            '
            Call FINAL("*COCAMP0")
          Case 2
            ORD% = 0
            Call CoCoPen(ORD%)        ' ORDENES PENDIENTES
            '
            If HAYMATERIALIB = 1 Then
                Call COMUNI("La Consulta por Pantalla\Permite Visualizar Solamente\Materiales CODIFICADOS.")
            End If
            '
            Call FINAL("*COCAPEN")
          Case Else
            Exit Sub
        End Select
        '
      ElseIf Option2.Value = True Then 'IMPRESORA
        Select Case OPXX%
          Case 1
            Call COCOMP0
            If HAYMATERIALIB% = 1 Then 'SI HAY MAT. TEXTO LIBRE
                Call FINAL("*LICOMALI")
              Else
                HAYOCX% = 0: If EXISTE%(LUDAT$ + "OCOMDETX.RFS") > 0 Then HAYOCX% = 1
                If HAYOCX% < 1 Then
                    Call FINAL("*LOCABPO")
                  Else
                    Call FINAL("*LICOMPX")
                End If
            End If
          Case 2
            ORD% = 0
            Call CoCoPen(ORD%)        ' ORDENES PENDIENTES
            If HAYMATERIALIB% = 1 Then
                Call FINAL("*LICOPEL")
            Else
                Call FINAL("*LOCAPEN")
            End If
          Case Else
            Exit Sub
        End Select
    End If
End Sub

Sub Command4_Click()
    Command4.Enabled = False
    ORD% = 1
    Call CoCoPen(ORD%)    ' SELECCIONAR SOLO PENDIENTES
    If Option1.Value = True Then
        '
        If HAYMATERIALIB = 1 Then
                Call COMUNI("La Consulta por Pantalla\Permite Visualizar Solamente\Materiales CODIFICADOS.")
        End If
        '
        Call FINAL("*COCAPRO")
      ElseIf Option2.Value = True Then
        If HAYMATERIALIB% = 1 Then
            Call FINAL("*LOCPRLI")
          Else
            Call FINAL("*LOCABPRO")
        End If
    End If
    Command4.Enabled = True
End Sub

Sub Command5_Click()
    Command5.Enabled = False
    ORD% = 2
    Call CoCoPen(ORD%)    ' SELECCIONAR SOLO PENDIENTES
    If Option1.Value = True Then
        '
        If HAYMATERIALIB = 1 Then
                Call COMUNI("La Consulta por Pantalla\Permite Visualizar Solamente\Materiales CODIFICADOS.")
        End If
        '
        Call FINAL("*COCAMAT")
      ElseIf Option2.Value = True Then
        If HAYMATERIALIB = 1 Then
            Call FINAL("*LOCMATL")
          Else
            Call FINAL("*LOCABMAT")
        End If
    End If
    Command5.Enabled = True
End Sub



Private Sub Command50_Click()
    Call CIERRARCH("*.*")
    ACONEC$ = "brecep04/forecep04"
    Call CONECRUN(ACONEC$, "Recpecion de Proveedores")
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   OPX% = ALTERNA%("Control de Recepciones\Otros Accesos Directos")
   If OPX% = 1 Then
       Call CONECRUN("BRECEP04/CONRECEP", "Recepcion de Proveedores")
     ElseIf OPX% = 2 Then
       Call ACCDIR("OCOMPRA")
   End If
   Command8.Enabled = True
End Sub

Sub Command9_Click()
   Command9.Enabled = False
   'ACA LISTADO DE MAT EN CONSIGNACION
       If Option1.Value = True Then
            MODOEMI% = COMALTER%("Opciones de Consulta:\\Por Proveedor\Por Artículo\Cancelar")
            Call GENEMATCON
            If MODOEMI% = 1 Then
                 Call FINAL("*COMATCO")
               ElseIf MODOEMI% = 2 Then
                 Call FINAL("*COMACON")
            End If
            Command9.Enabled = True
            Exit Sub
          ElseIf Option2.Value = True Then
            MODOEMI% = COMALTER%("Opciones de Emisión de Listado:\\Secuencial Completo\Pendientes Por Proveedor\Pendientes Por Artículo\Cancelar")
            Call LITMAOC(MODOEMI%)
            If MODOEMI% = 1 Then
                 Call FINAL("*LIMATCO")
               ElseIf MODOEMI% = 2 Or MODOEMI% = 3 Then
                 'Call GENCONPEN(MODOEMI%)
                 Call GENCONPEN1(MODOEMI%)
            End If
        End If
    Command9.Enabled = True
End Sub

Sub GENCONPEN1(MODOEMI%)
   Screen.MousePointer = 11
   Call SETULTREG("LCONPEN", 0)
   Call ABRESTOCKS
   SQLX$ = "Select * from MACONSIG"
   SQLX$ = SQLX$ + " WHERE ( Cant_Entreg - Cant_Rendid )>0 "
   'SQLX$ = SQLX$ + " WHERE (Cant_Asig - Cant_Entreg )>0 "
   '
   Select Case MODOEMI%
    Case 2
        SQLX$ = SQLX$ + " ORDER BY Npro_Ocom "
    Case 3
        SQLX$ = SQLX$ + " ORDER BY Cod_Inte "
   End Select
   '
   Set GECONPENTEMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
   With GECONPENTEMP
   On Error Resume Next
       .MoveLast
       'Valida que la tabla no este vacia
       If Err Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         On Error GoTo 0
         MODOEMI% = 0
         GoTo 90
       End If
       On Error Resume Next
   FIN& = .RecordCount
   .MoveFirst
   JJ& = 0
   Do While (.EOF = False)
         XXY$ = REGBLAN$("LCONPEN")
         Call REPLA(XXY$, MKS$(CVINT(!Femi_Ocom)), 1, 2)
         Call REPLA(XXY$, MKI$(!Npro_Ocom), 3, 2)
         Call REPLA(XXY$, MKD$(!Secc_Ejec), 5, 4)
         If IsNull(!NRO_ORFA) Then
            Call REPLA(XXY$, "", 9, 12)
         Else
            Call REPLA(XXY$, !NRO_ORFA, 9, 12)
         End If
         Call REPLA(XXY$, !Nro_Ocom, 21, 12)
         Call REPLA(XXY$, MKI$(!Cod_inte), 33, 2)
         If IsNull(!LOTE_OCOM) Then
            Call REPLA(XXY$, "", 35, 16)
         Else
            Call REPLA(XXY$, !LOTE_OCOM, 35, 16)
         End If
         Call REPLA(XXY$, !Unidad, 51, 4)
         Call REPLA(XXY$, MKD$(!Cant_Asig), 55, 8)
         Call REPLA(XXY$, MKD$(!Cant_Entreg), 63, 8)
         Call REPLA(XXY$, MKI$(CVINT(!Fecha_Entreg)), 71, 2)
         If IsNull(!NRemi_Consig) Then
            Call REPLA(XXY$, "", 73, 6)
         Else
            Call REPLA(XXY$, !NRemi_Consig, 73, 6)
         End If
         Call REPLA(XXY$, MKD(!Valo_Unit), 93, 8)
         Call REPLA(XXY$, MKI$(!MONEDA), 101, 2)
         Call REPLA(XXY$, MKD$(!Valo_Tot), 103, 8)
         Call REPLA(XXY$, MKD$(!Cant_Rendid), 111, 8)
           SALPE# = !Cant_Entreg - !Cant_Rendid
           If SALPE# < 0 Then SALPE# = 0
         Call REPLA(XXY$, MKD$(SALPE#), 119, 8)
         If IsNull(!Observa) Then
            Call REPLA(XXY$, "", 129, 16)
         Else
            Call REPLA(XXY$, !Observa, 129, 16)
         End If
         JJ& = JJ& + 1
         Call GRAREG("LCONPEN", XXY$, JJ&)
     .MoveNext
   Loop
  End With
  '
  Call SETULTREG("LCONPEN", JJ&)
  Call CIERRARCH("LCONPEN")
  '
  Select Case MODOEMI%
    Case 2
      Call FINAL("*LCONPE2")
    Case 3
      Call FILESORT("LCONPEN", "LCONPEN", 8, 8, "ASC")
      Call FINAL("*LCONPE3")
  End Select
  '
90 Screen.MousePointer = 1
   End Sub

Sub LITMAOC(MODO%)
Screen.MousePointer = 11
   'NUEVO
   Call SETULTREG("MACONSIG", 0)
   Call SETULTREG("LCONPEN", 0)
   Call ABRESTOCKS
   SQLX$ = "Select * from MACONSIG"
   '
   Select Case MODO%
    Case 1
        SQLX$ = SQLX$
    Case 2
        SQLX$ = SQLX$ + " ORDER BY Npro_Ocom "
    Case 3
        SQLX$ = SQLX$ + " ORDER BY Cod_Inte "
   End Select
   '
   Set LITMAOCTEMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
   With LITMAOCTEMP
     On Error Resume Next
     .MoveLast
     'Valida que la tabla no este vacia
     If Err Then
       Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
       On Error GoTo 0
       MODO% = 0
       GoTo 90
     End If
     On Error GoTo 0
     FIN& = .RecordCount
     .MoveFirst
     JJ& = 0
     '
     Do While (.EOF = False)
        Select Case MODO%
          Case 1
            XXY$ = REGBLAN$("MACONSIG")
          Case Else
            XXY$ = REGBLAN$("LCONPEN")
        End Select
        '
        If !Cod_inte > 0 Then
          Call REPLA(XXY$, MKS$(CVINT(!Femi_Ocom)), 1, 2)
          Call REPLA(XXY$, MKI$(!Npro_Ocom), 3, 2)
              SEJEC# = 0
              On Error Resume Next
              SEJEC# = !Secc_Ejec
              On Error GoTo 0
          Call REPLA(XXY$, MKD$(SEJEC#), 5, 4)
            If IsNull(!NRO_ORFA) = False Then Call REPLA(XXY$, !NRO_ORFA, 9, 12)
          Call REPLA(XXY$, !Nro_Ocom, 21, 12)
          Call REPLA(XXY$, MKI$(!Cod_inte), 33, 2)
            If IsNull(!LOTE_OCOM) = False Then Call REPLA(XXY$, !LOTE_OCOM, 35, 16)
          Call REPLA(XXY$, !Unidad, 51, 4)
          Call REPLA(XXY$, MKD$(!Cant_Asig), 55, 8)
          Call REPLA(XXY$, MKD$(!Cant_Entreg), 63, 8)
          Call REPLA(XXY$, MKI$(CVINT(!Fecha_Entreg)), 71, 2)
            If IsNull(!NRemi_Consig) = False Then Call REPLA(XXY$, !NRemi_Consig, 73, 6)
              VUNIT# = 0
              On Error Resume Next
              VUNIT# = !Valo_Unit
              On Error GoTo 0
          Call REPLA(XXY$, MKD(VUNIT#), 93, 8)
              MONEII% = 0
              On Error Resume Next
              MONEII% = !MONEDA
              On Error GoTo 0
          Call REPLA(XXY$, MKI$(MONEII%), 101, 2)
              VATOTA# = 0
              On Error Resume Next
              VATOTA# = !Valo_Tot
              On Error GoTo 0
          Call REPLA(XXY$, MKD$(VATOTA#), 103, 8)
              CARENDI# = 0
              On Error Resume Next
              CARENDI# = !Cant_Rendid
              On Error GoTo 0
          Call REPLA(XXY$, MKD$(CARENDI#), 111, 8)
              SALDOPE# = 0
              On Error Resume Next
              SALDOPE# = !Saldo_Pend
              On Error GoTo 0
          Call REPLA(XXY$, MKD$(SALDOPE#), 119, 8)
            If IsNull(!Observa) = False Then Call REPLA(XXY$, !Observa, 129, 16)
          JJ& = JJ& + 1
          Select Case MODO%
            Case 1
               Call GRAREG("MACONSIG", XXY$, JJ&)
            Case Else
               Call GRAREG("LCONPEN", XXY$, JJ&)
          End Select
        End If
     .MoveNext
     Loop
   End With
   '
   Select Case MODO%
     Case 1
        Call SETULTREG("MACONSIG", JJ&)
        Call CIERRARCH("MACONSIG")
     Case Else
        Call SETULTREG("LCONPEN", JJ&)
        Call CIERRARCH("LCONPEN")
   End Select
   '
90 Screen.MousePointer = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
End Sub

Sub PEDISUMI()             ' NUEVO PEDIDO DE SUMINISTRO
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
100 Call CIERRARCH("*.*")
    REGXX& = ULTREG&("OCOMPRA")
    NRO& = 1
    If REGXX& > 0 Then
        NRO& = 1 + CVS(Left$(REGLEIDO$("OCOMPRA", REGXX&), 4))
    End If
    If NRO& <= REGXX& Then
        Screen.MousePointer = 11
        For U& = 1 To REGXX&
            X0$ = REGLEIDO$("OCOMPRA", U&)
            If CVS(Left$(X0$, 4)) >= NRO& Then
                NRO& = 1 + CVS(Left$(X0$, 4))
            End If
        Next U&
        Screen.MousePointer = 1
    End If
    '
    X0$ = REGBLAN$("OCOMPRA")
    Call REPLA(X0$, MKS$(NRO&), 1, 4)
    Call REPLA(X0$, MKI$(HO%), 5, 2)
    Call REPLA(X0$, MKI$(HO%), 9, 2)
    Call REPLA(X0$, "S", 57, 1)
    '
101 OBJE$ = "PEDISUMI"
    '
    X1$ = OBJPLA$(OBJE$, X0$)
    If X1$ = "" Then GoTo 299
    '
    X0$ = X1$
    NRO& = CVS(Left$(X1$, 4))
    If NRO& < 1 Then GoTo 101
    '
    FEMI% = CVI(Mid$(X1$, 5, 2))
    If FEMI% < 1 Then
        Call MENSERR(24, "Falta Fecha de Emisión")
        GoTo 101
    End If
    '
    NPX$ = MKI$(0)
    NPRO% = CVI(NPX$)
    '
    Call BLANARCH("!CARDETOC")
    '
102 TITU$ = TRIM$(Str$(NRO&)) + " - Uso Interno"
    VTB% = VENTABU%("CARPESU1/TITUPAN=Pedido de Suministro Nro. " + TITU$)
    If VTB% < 0 Then GoTo 101
    '
    CAICO% = 0
    For J& = 1 To ULTREG&("!CARDETOC")
       X$ = REGLEIDO$("!CARDETOC", J&)
       CI0% = CVI(Left$(X$, 2))
       If CODEXT$(CI0%) = "" Then
          Call MENSERR(24, "Código de Artículo no Válido\en Item Número " + TRIM$(Str$(J&)) + ".")
          GoTo 102
       End If
       CANTU = CVS(Mid$(X$, 47, 4))
       If CANTU < 0.05 Then
          Call MENSERR(24, "Falta Indicar Cantidad en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          GoTo 102
       End If
       CAICO% = J&
       '
       Call REPLA(X$, UNIMED$(CI0%), 3, 4)
       If CVI(Mid$(X$, 67, 2)) < HO% Then
          Call REPLA(X$, MKI$(HO%), 67, 2)
       End If
       Call REPLA(X$, "  ", 69, 2)
       '
       Call LEEEXPLO(CODEXT$(CI0%), 1)
       If CAIT% > 0 Then
          Call REPLA(X$, " *", 69, 2)
       End If
       '
       Call GRAREG("!CARDETOC", X$, J&)
       '
    Next J&
    '
    If CAICO% < 1 Then
       Call MENSERR(24, "Falta Ingresar Detalle Ordenado")
       GoTo 102
    End If
    '
103 VTB% = VENTABU%("CARPESU2/TITUPAN=Pedido de Suministro Nro. " + TITU$)
    If VTB% < 0 Then GoTo 102
    '
    TOTANE# = 0
    For J& = 1 To ULTREG&("!CARDETOC")
       X$ = REGLEIDO$("!CARDETOC", J&)
       CI0% = CVI(Left$(X$, 2))
       CANTU = CVS(Mid$(X$, 47, 4))
       FENTRE% = CVI(Mid$(X$, 67, 2))
       '
       If FENTRE% < FEMI% Then
          Call MENSERR(24, "Fecha de Entrega Incoherente\en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          Call REPLA(X0$, MKI$(FEMI%), 9, 2)
          GoTo 103
       End If
       '
    Next J&
    '
    Call CIERRARCH("!CARDETOC")
    '
    Call BLOQARCH("OCOMPRA")
    Call BLOQARCH("MACONSIG")
    '
    For J& = 1 To ULTREG&("!CARDETOC")
       X$ = REGLEIDO$("!CARDETOC", J&)
       CI0% = CVI(Left$(X$, 2))
       CANTU = CVS(Mid$(X$, 47, 4))
       PREUNID# = CVD(Mid$(X$, 51, 8))
       'PORDE = CVS(Mid$(X$, 55, 4))
       FENTRE% = CVI(Mid$(X$, 67, 2))
       CONSIMAT% = 0: If TRIM$(Mid$(X$, 69, 2)) <> "" Then CONSIMAT% = 1
       '
       X2$ = X1$
       Call REPLA(X2$, MKI$(FENTRE%), 9, 2)
       Call REPLA(X2$, MKI$(CI0%), 11, 2)
       Call REPLA(X2$, MKS$(CANTU), 13, 4)
       'Call REPLA(X2$, MKS$(0), 58, 4)
       'Call REPLA(X2$, Chr$(1 + CONSIMAT%), 57, 1) ' normal
       Call REPLA(X1$, Chr$(0), 62, 1)          ' abierta
       '
       Call REGAPP("OCOMPRA", X2$)
       '
       If CONSIMAT% = 1 Then
          '
          NOCOM$ = TRIM$(Str$(NRO&))
          While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
          ITNU$ = TRIM$(Str$(J&))
          While Len(ITNU$) < 2: ITNU$ = "0" + ITNU$: Wend
          NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-" + ITNU$
          '
          Call LEEEXPLO(CODEXT$(CI0%), CANTU)
          For I% = 1 To CAIT%
             YY$ = REGBLAN$("MACONSIG")
             Call REPLA(YY$, MKI$(FEMI%), 1, 2)
             Call REPLA(YY$, NPX$, 3, 2)
             Call REPLA(YY$, NOCOM$, 21, 12)
             '
             CII% = CIJ%(I%)
             CANTI# = CDbl(CANTU * USOI(I%))
             '
             If TRIM$(CODEXT$(CII%)) <> "" Then
               If CANTI# > 0.05 Then
                 Call REPLA(YY$, MKI$(CII%), 33, 2)
                 Call REPLA(YY$, UNIMED$(CII%), 51, 4)
                 Call REPLA(YY$, MKD$(CANTI#), 55, 8)
                 Call REGAPP("MACONSIG", YY$)
               End If
             End If
          Next I%
       End If
       '
    Next J&
    '
    Call CIERRARCH("OCOMPRA")
    Call CIERRARCH("MACONSIG")
    '
' preguntar si imprime si/no
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      FECHDOC$ = FETEXTO$(FEMI%)
      NUMEDOC$ = TRIM$(Str$(NRO&))
      While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Pedido de Suministro"
      '
      CODPARAM$(1) = "COND-PAG"
      CODPARAM$(2) = "OBSERVA"
      CODPARAM$(3) = "IMP-NETO"
      CODPARAM$(4) = "POR-IVA"
      CODPARAM$(5) = "IMP-IVA"
      CODPARAM$(6) = "IMP-TOTA"
      CAPARDOC% = 6
      '
      DOCPARAM$(1) = FORPAG$
      DOCPARAM$(2) = OBSE$
      DOCPARAM$(3) = Str$(TOTANE#)
      DOCPARAM$(4) = Str$(ALIVA)
      DOCPARAM$(5) = Str$(TOTIVA#)
      DOCPARAM$(6) = Str$(TOTATO#)
      
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "UNIMED"
      CODTABU1$(4) = "CANTIDAD"
      CODTABU1$(5) = "FECH-VEN"
      CODTABU1$(6) = "REF-MAT1"
      CODTABU1$(7) = "PRE-LIST"
      CODTABU1$(8) = "PORDESCU"
      CODTABU1$(9) = "PRE-NETO"
      CODTABU1$(10) = "NET-ITEM"
      CACOLTAB1% = 10
      '
      DESTIDOC% = NC1%
      CAITAB1% = 0
      CAITCRU% = 0
      '
      For J& = 1 To ULTREG&("!CARDETOC")
         X$ = REGLEIDO$("!CARDETOC", J&)
         CI0% = CVI(Left$(X$, 2))
         CANTU = CVS(Mid$(X$, 47, 4))
         PREUNID# = CVD(Mid$(X$, 51, 8))
         'PORDE = CVS(Mid$(X$, 55, 4))
         FEX = CVI(Mid$(X$, 67, 2))
         CONSIMAT% = 0: If TRIM$(Mid$(X$, 69, 2)) <> "" Then CONSIMAT% = 1
         '
         CAITAB1% = CAITAB1% + 1
         TETABU1$(1, CAITAB1%) = CODEXT$(CI0%)
         TETABU1$(2, CAITAB1%) = DESCRIT0$(CI0%)
         TETABU1$(3, CAITAB1%) = Mid$(X$, 3, 4)
         TETABU1$(4, CAITAB1%) = TRIM$(Str$(CANTU))
         TETABU1$(5, CAITAB1%) = FECHATEX$(FEX)
         TETABU1$(6, CAITAB1%) = Mid$(X$, 69, 2)
         TETABU1$(7, CAITAB1%) = Str$(PREUNID#)
         TETABU1$(8, CAITAB1%) = Str$(PORDE)
         TETABU1$(9, CAITAB1%) = Str$(PREUNID# * (1 - PORDE / 100))
         TETABU1$(10, CAITAB1%) = Str$(CANTU * PREUNID# * (1 - PORDE / 100))
        '
109   Next J&
      '
      Call PRINTDOC("FORMOCOM")   ' Orden de Compra
      '
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
    End If
    
'consignacion inmediata si/no

    GoTo 100
    '
299 Call CIERRARCH("*.*")
    '
End Sub
'
Sub CoCoPen(ORDEN%)
   '
   If RECLEN("OPENDETA") < 160 Then
     Call COMUNI("ATENCION: Controlador 'OPENDETA.RFS' Desactualizado.\Hay Campos que no Serán Listados.\  \Solicite Actualización S/Cargo a Soporte de Sistemas.")
   End If
   '
   Screen.MousePointer = 11
   'NUEVO 11/03/2005
   'MARCA PARA SABER SI HAY MATERIALES DE TEXTO LIBRE
   HAYMATERIALIB% = 0
   '
   Call SETULTREG("OPENDETA", 0)
   Call SETULTREG("OPENDETL", 0)
   Call ABRECOMPRAS
   SQLX$ = "Select * from DETOCABI"
   SQLX$ = SQLX$ + " Where Stat_Ocom < 8 "
   SQLX$ = SQLX$ + " AND Stat_Ocom <> -1 "
   '//**** que condicion se toma para que sean pendientes?***//
   'SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
   '\\\OT-5-0397-FG-13/06/05///
   Select Case ORDEN%
    Case 0
        SQLX$ = SQLX$ + " ORDER BY Nume_Ocom, Cod_Exte, FVig_Ocom "
    Case 1
        SQLX$ = SQLX$ + " ORDER BY Npro_Ocom, Nume_Ocom, Cod_Exte, FVig_Ocom "
    Case 2
        SQLX$ = SQLX$ + " ORDER BY Cod_Exte, Nume_Ocom, FVig_Ocom "
    End Select
   '\\\OT-5-0397-FG-FIN///
   Set OCOMTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   With OCOMTEMP
  On Error Resume Next
       .MoveLast
       'Valida que la tabla no este vacia
       If Err Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         On Error GoTo 0
         GoTo 90
       End If
   On Error GoTo 0
   .MoveFirst
  JJ& = 0
  TC& = 0
  Do While (.EOF = False)

     XXX$ = REGBLAN$("OPENDETA")
         XXY$ = XXX$ + String$(32, 0)
         MONEMI% = !Mone_Emis
         If MONEMI% <= 1 Then MONEX$ = "   "
         If MONEMI% = 2 Then MONEX$ = "U$S "
         If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
         TICAMBIO = 1
         If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
         '
         Call REPLA(XXY$, MONEX$, 63, 4)
         CAPEN = !Cant_Ocom - !Cant_Rec
         If CAPEN < 0.05 Then CAPEN = 0
         PREUNID# = !PrUn_Neto * TICAMBIO
         Call REPLA(XXY$, MKS$(!Nume_Ocom), 1, 4) 'AGREGO EL Nº DE ORDEN DE COMPRA
         Call REPLA(XXY$, MKI$(CVINT%(!Femi_Ocom)), 5, 2)
         Call REPLA(XXY$, MKI$(!Npro_Ocom), 7, 2)
         Call REPLA(XXY$, MKI$(!Cod_inte), 9, 2)
         Call REPLA(XXY$, MKS$(!Cant_Ocom), 55, 4)
         Call REPLA(XXY$, !Uni_Com, 59, 4)
         Call REPLA(XXY$, MKD$(!PrUn_Neto), 79, 8)
         Call REPLA(XXY$, MKD$(!Cant_Ocom * !PrUn_Neto), 87, 8)
         Call REPLA(XXY$, MKI$(CVINT%(!FVig_Ocom)), 95, 2)
         Call REPLA(XXY$, !CONSIMAT, 97, 2)
         Call REPLA(XXY$, MKS$(!Cant_Rec), 99, 4)
         Call REPLA(XXY$, MKI$(MONEMI%), 121, 2)
         Call REPLA(XXY$, MKS$(CAPEN), 129, 4)
           PREUNID# = !PrUn_Neto * TICAMONE(MONEMI%)
         Call REPLA(XXY$, MKD$(PREUNID#), 133, 8)
         Call REPLA(XXY$, MKD$(PREUNID# * CAPEN), 141, 8)
         '
         If !Cod_inte <= 0 Then HAYMATERIALIB% = 1 'HAY MATERIAL TEXTO LIBRE
         '
         JJ& = JJ& + 1
         If !Cod_inte > 0 Then
            TC& = TC& + 1
            Call GRAREG("OPENDETA", XXY$, TC&)
         End If
         '
         XXZ$ = REGBLAN("OPENDETL")
         XXZ$ = Left$(XXY$, 10) + Space$(118) + Mid$(XXY$, 11)
         If !Cod_inte <= 0 Then
             Call REPLA(XXZ$, MKI$(!Cod_inte), 9, 2)
             Call REPLA(XXZ$, !Cod_Exte, 11, 16)
             '\\\OT-05-0690-WAR-21/10/05///
             'Call REPLA(XXZ$, !Descritem, 27, 64)
             Call REPLA(XXZ$, !Delibre, 27, 64)
             '\\\OT-05-0690-WAR-FIN///
           Else
             CIXI% = !Cod_inte
             Call REPLA(XXZ$, CODEXT$(CIXI%), 11, 16)
             Call REPLA(XXZ$, DESCRIT0$(CIXI%), 27, 64)
         End If
         '
         If IsNull(!Nro_Otrab) Then
                NOTR$ = ""
           Else
                NOTR$ = !Nro_Otrab
         End If
         If IsNull(!Nro_Grup) Then
                NGR$ = ""
           Else
                NGR$ = !Nro_Grup
         End If
         DEST1$ = NOTR$
         If DEST1$ <> "" Then
           If NGR$ <> "" Then
             DEST1$ = DEST1$ + " / "
           End If
         End If
         DEST1$ = DEST1$ + NGR$
         If TRIM$(DEST1$) = "" Then
           CCO% = !IContab
           DEST1$ = CODCUE$(CCO%)
         End If
         Call REPLA(XXZ$, DEST1$, 91, 24)
         If IsNull(!Sect_Ocom) = False Then Call REPLA(XXZ$, !Sect_Ocom, 115, 16)
         '
         Call GRAREG("OPENDETL", XXZ$, JJ&)
         '
     .MoveNext
    Loop
   End With
   '
   Call SETULTREG("OPENDETA", TC&)
   Call CIERRARCH("OPENDETA")
   '
   Call SETULTREG("OPENDETL", JJ&)
   Call CIERRARCH("OPENDETL")
   '
90 Screen.MousePointer = 1
   End Sub

Sub COCOMP0()
   '
    Screen.MousePointer = 11
    'MARCA PARA SABER SI HAY MATERIALES DE TEXTO LIBRE
    HAYMATERIALIB% = 0
    HAYOCX% = 0: If EXISTE%(LUDAT$ + "OCOMDETX.RFS") > 0 Then HAYOCX% = 1
    'NUEVa subrutina para base de datos completa de OC
    Call SETULTREG("OCOMDETA", 0)
    Call SETULTREG("OCOMDETL", 0)
    Call ABRECOMPRAS
    SQLX$ = "Select * from DETOCABI where Stat_Ocom < 9 "
    Set OCOMP0TEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    With OCOMP0TEMP
      On Error Resume Next
      .MoveLast
      'Valida que la tabla no este vacia
      If Err Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         On Error GoTo 0
         GoTo 90
      End If
      On Error GoTo 0
      FIN& = .RecordCount
      .MoveFirst
      '
      JJ& = 0
      TC& = 0
      Do While (.EOF = False)
        XXX$ = REGBLAN$("OCOMDETA") + Space$(32)
        XXY$ = XXX$
        '
        MONEMI% = !Mone_Emis
        If MONEMI% <= 1 Then MONEX$ = "   "
        If MONEMI% = 2 Then MONEX$ = "U$S "
        If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
        TICAMBIO = 1
        If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
        '
        Call REPLA(XXY$, MONEX$, 121, 2)
        CAPEN = !Cant_Ocom - !Cant_Rec
        If CAPEN < 0.05 Then CAPEN = 0
        PREUNID# = !PrUn_Neto * TICAMBIO
        
        Call REPLA(XXY$, MKS$(!Nume_Ocom), 1, 4) 'AGREGO EL Nº DE ORDEN DE COMPRA
        Call REPLA(XXY$, MKI$(CVINT%(!Femi_Ocom)), 5, 2)
        Call REPLA(XXY$, MKI$(!Npro_Ocom), 7, 2)
        Call REPLA(XXY$, MKI$(!Cod_inte), 9, 2)
        '
        If !Cod_inte <= 0 Then HAYMATERIALIB% = 1 'HAY MATERIAL TEXTO LIBRE
        '
        Call REPLA(XXY$, MKS$(!Cant_Ocom), 55, 4)
        If IsNull(!Uni_Com) Then
          UNICOMOC$ = "  "
        Else
          UNICOMOC$ = !Uni_Com
        End If
        Call REPLA(XXY$, UNICOMOC$, 59, 4)
        Call REPLA(XXY$, MKD$(!PrUn_Neto), 79, 8)
        Call REPLA(XXY$, MKD$(!Cant_Ocom * !PrUn_Neto), 87, 8)
        Call REPLA(XXY$, MKI$(CVINT%(!FVig_Ocom)), 95, 2)
        If IsNull(!CONSIMAT) Then
          CONSIMATOC$ = "  "
        Else
          CONSIMATOC$ = !CONSIMAT
        End If
        Call REPLA(XXY$, CONSIMATOC$, 97, 2)
        Call REPLA(XXY$, MKS$(!Cant_Rec), 99, 4)
        Call REPLA(XXY$, MKI$(MONEMI%), 121, 2)
        Call REPLA(XXY$, MKI$(!Stat_Ocom), 124, 1)
        '
        If !Cod_inte > 0 Then
           TC& = TC& + 1
           Call GRAREG("OCOMDETA", XXY$, TC&)
           If HAYOCX% > 0 Then
               CUEII% = !IContab
             Call REPLA(XXY$, CODCUE$(CUEII%), 129, 16)
               SECCP$ = !Sect_Ocom
             Call REPLA(XXY$, SECCP$, 145, 16)
             Call GRAREG("OCOMDETX", XXY$, TC&)
           End If
        End If
        '
        XXZ$ = Left$(XXY$, 10) + Space$(118) + Mid$(XXY$, 11)
        If !Cod_inte <= 0 Then
            Call REPLA(XXY$, MKI$(!Cod_inte), 9, 2)
            Call REPLA(XXZ$, !Cod_Exte, 11, 16)
            '\\\OT-05-0690-WAR-21/10/05///
            'Call REPLA(XXZ$, !Descritem, 27, 64)
            Call REPLA(XXZ$, !Delibre, 27, 64)
            '\\\OT-05-0690-WAR-FIN///
          Else
            CIXI% = !Cod_inte
            Call REPLA(XXZ$, CODEXT$(CIXI%), 11, 16)
            Call REPLA(XXZ$, DESCRIT0$(CIXI%), 27, 64)
        End If
        '
        If IsNull(!Nro_Otrab) Then
           NOTR$ = ""
        Else
           NOTR$ = !Nro_Otrab
        End If
        If IsNull(!Nro_Grup) Then
           NGR$ = ""
        Else
           NGR$ = !Nro_Grup
        End If
        DEST1$ = NOTR$
        If DEST1$ <> "" Then
          If NGR$ <> "" Then
            DEST1$ = DEST1$ + " / "
          End If
        End If
        DEST1$ = DEST1$ + NGR$
        '
        '\\\OT-05-0690-WAR-21/10/05///
         If TRIM$(DEST1$) = "" Then
           CCO% = !IContab
           DEST1$ = CODCUE$(CCO%)
         End If
         If IsNull(!Sect_Ocom) Then
                SECT1$ = ""
           Else
                SECT1$ = !Sect_Ocom
         End If
         Call REPLA(XXZ$, SEXT1$, 15, 16)
        '\\\OT-05-0690-WAR-FIN///
        '
        Call REPLA(XXZ$, DEST1$, 91, 24)
        JJ& = JJ& + 1
        Call GRAREG("OCOMDETL", XXZ$, JJ&)
        '
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("OCOMDETA", TC&)
    If HAYOCX% > 0 Then Call SETULTREG("OCOMDETX", TC&)
    '\\\OT-05-0693-WAR-26/10/05///
    Call FILESORT("OCOMDETL", "OCOMDETL", 1, 1, "ASC")
    Call FILESORT("OCOMDETA", "OCOMDETA", 1, 1, "ASC")
    '\\\OT-05-0693-WAR-FIN///
    Call SETULTREG("OCOMDETL", JJ&)
    Call CIERRARCH("OCOMDETA")
    Call CIERRARCH("OCOMDETL")
    Screen.MousePointer = 1
    '
90 Screen.MousePointer = 1
   End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Sub SELCOMRAFE(OKKX%, MODO%)
   '
   'CODIGO NUEVO AL 11/3/2005 WAR
   'MARCA PARA SABER SI HAY MAT. TEXTO LIBRE
   HAYMATERIALIB% = 0
   '
   HAS% = HOY(HOS$)
   FEX = HAS%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   DES% = FECHANUM(DESS$)
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
10 OKKX% = 0
   OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
      Exit Sub
   End If
   OKKX% = 1
   DES% = CVI(Left$(OBX$, 2)): DES1 = CDbl(CVDAT(DES%))
   HAS% = CVI(Mid$(OBX$, 3, 2)): HAS1 = CDbl(CVDAT(HAS%))
   If HAS% < DES% Then GoTo 10
   FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   Screen.MousePointer = 11
   'Nuevo
   REBLA$ = REGBLAN$("COPRECOM")
   XXZ$ = REGBLAN("COPRECOL")
   Call HEADERS("COPRECOM", "")
   Call HEADERS("COPRECOM", "Periodo: " + PERIO$)
   '
   Call ABRECOMPRAS
   SQLX$ = "SELECT * FROM OCOMDETA"
   SQLX$ = SQLX$ + " WHERE Stat_Ocom < 9 "
   SQLX$ = SQLX$ + " AND CDBL(Femi_Ocom) >= " & DES1 & " "
   SQLX$ = SQLX$ + " AND CDBL(Femi_Ocom) <=   " & HAS1 & "  "
   If MODO% = 1 Then
        SQLX$ = SQLX$ + " ORDER BY  Cod_Exte "
      ElseIf MODO% = 0 Then
        SQLX$ = SQLX$ + " ORDER BY Npro_Ocom  "
   End If
   
  
    Set OCOMPTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    On Error GoTo 90
    With OCOMPTEMP
        .MoveFirst
        JJ& = 0
        TC& = 0
        Do While (.EOF = False)
               FECHEMI% = CVINT(!Femi_Ocom)
               STAT% = !Stat_Ocom
               NPRO% = !Npro_Ocom
               CIXI% = !Cod_inte
               DOPE$ = ""
               '
               NOPE% = !Cod_Oper
               If NOPE% > 0 Then
                 DOPE$ = DESCRIOP$(CIXI%, NOPE%)
               End If
               NUOC& = !Nume_Ocom
               '
               UMEDI$ = !Uni_Com
               If TRIM$(UMEDI$) = "" Then
                 UMEDI$ = UNIMED$(CIXI%)
               End If
               CACOM = !Cant_Ocom
               '
               If STAT% = 8 Then                 ' SI SE DIO POR CUMPLIDA,
                 CACOM = !Cant_Rec
               End If
               '
               PREUNI# = !Pre_Unit
               '
               MONEMI% = !Mone_Emis
               If MONEMI% <= 1 Then MONEX$ = "    "
               If MONEMI% = 2 Then MONEX$ = "U$S "
               If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
               TICAMBIO = 1
               If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
               '
               YYY$ = REBLA$
               Call REPLA(YYY$, MKI$(CIXI%), 1, 2)
               Call REPLA(YYY$, MKI$(FECHEMI%), 3, 2)
               Call REPLA(YYY$, MKI$(NPRO%), 5, 2)
               Call REPLA(YYY$, MKI$(CIXI%), 7, 2)
               Call REPLA(YYY$, MKI$(FECHEMI%), 9, 2)
               Call REPLA(YYY$, DOPE$, 11, 22)
               Call REPLA(YYY$, MKS$(NUOC&), 33, 4)
               Call REPLA(YYY$, UMEDI$, 37, 4)
               Call REPLA(YYY$, MKS$(CACOM), 41, 4)
               Call REPLA(YYY$, MONEX$, 49, 4)
               Call REPLA(YYY$, MKD$(PREUNI#), 53, 8)
               ITOCOM# = PREUNI# * CACOM
               Call REPLA(YYY$, MKD$(ITOCOM#), 61, 8)
               Call REPLA(YYY$, MKD$(PREUNI# * TICAMBIO), 69, 8)
               Call REPLA(YYY$, MKD$(ITOCOM# * TICAMBIO), 77, 8)
               '
               If !Cod_inte <= 0 Then HAYMATERIALIB% = 1
               '
               JJ& = JJ& + 1
               If !Cod_inte > 0 Then
                    TC& = TC& + 1
                    Call GRAREG("COPRECOM", YYY$, TC&)
               End If
               '
                XXZ$ = Left$(YYY$, 6) + Space$(122) + Mid$(YYY$, 7)
                If !Cod_inte <= 0 Then
                    Call REPLA(XXY$, MKI$(!Cod_inte), 1, 2)
                    Call REPLA(XXZ$, !Cod_Exte, 7, 16)
                    '\\\OT-05-0690-WAR-24/10/05///
                    'Call REPLA(XXZ$, !Descritem, 23, 64)
                    Call REPLA(XXZ$, !Delibre, 23, 64)
                    '\\\OT-05-0690-WAR-FIN///
                  Else
                    CIXI% = !Cod_inte
                    Call REPLA(XXZ$, CODEXT$(CIXI%), 7, 16)
                    Call REPLA(XXZ$, DESCRIT0$(CIXI%), 23, 64)
                End If
                '
                If IsNull(!Nro_Otrab) Then
                       NOTR$ = ""
                  Else
                       NOTR$ = !Nro_Otrab
                End If
                If IsNull(!Nro_Grup) Then
                       NGR$ = ""
                  Else
                       NGR$ = !Nro_Grup
                End If
                DEST1$ = NOTR$
                If DEST1$ <> "" Then
                  If NGR$ <> "" Then
                    DEST1$ = DEST1$ + " / "
                  End If
                End If
                DEST1$ = DEST1$ + NGR$
                '\\\OT-05-0690-WAR-24/10/05///
                If TRIM$(DEST1$) = "" Then
                  CCO% = !IContab
                  DEST1$ = CODCUE$(CCO%)
                End If
                '
                If IsNull(!Sect_Ocom) Then
                       SECT1$ = ""
                  Else
                       SECT1$ = !Sect_Ocom
                End If
                Call REPLA(XXZ$, SECT1$, 111, 18)
                '\\\OT-05-0690-WAR-FIN///
                Call REPLA(XXZ$, DEST1$, 87, 24)
                '
'                Call REPLA(XXZ$, DOPE$, 133, 22)
'                Call REPLA(XXZ$, MKS$(NUOC&), 155, 4)
'                Call REPLA(XXZ$, UMEDI$, 159, 4)
'                Call REPLA(XXZ$, MKS$(CACOM), 163, 4)
'                Call REPLA(XXZ$, MONEX$, 171, 4)
'                Call REPLA(XXZ$, MKD$(PREUNI#), 175, 8)
'                ITOCOM# = PREUNI# * CACOM
'                Call REPLA(XXZ$, MKD$(ITOCOM#), 183, 8)
'                Call REPLA(XXZ$, MKD$(PREUNI# * TICAMBIO), 191, 8)
'                Call REPLA(XXZ$, MKD$(ITOCOM# * TICAMBIO), 199, 8)
                '
                Call GRAREG("COPRECOL", XXZ$, JJ&)
         .MoveNext
      Loop
    End With
   '
   Call SETULTREG("COPRECOM", TC&)
   Call CIERRARCH("COPRECOM")
   'Call FINAL("*COPRECOM")
   '
   Call SETULTREG("COPRECOL", JJ&)
   Call CIERRARCH("COPRECOL")
   Screen.MousePointer = 1
   '
'   'CODIGO ANTIGUO
'   HAS% = HOY(HOS$)
'   FEX = HAS%
'   HASS$ = FECHATEX$(FEX)
'   DESS$ = "01" + Mid$(HASS$, 3)
'   DES% = FECHANUM(DESS$)
'   VDEF$ = MKI$(DES%) + MKI$(HAS%)
'   '
'10 OKKX% = 0
'   OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
'   If OBX$ = "" Then
'      Exit Sub
'   End If
'   OKKX% = 1
'   DES% = CVI(Left$(OBX$, 2)): DES1 = CDbl(CVDAT(DES%))
'   HAS% = CVI(Mid$(OBX$, 3, 2)): HAS1 = CDbl(CVDAT(HAS%))
'   If HAS% < DES% Then GoTo 10
'   FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
'   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
'   '
'   Screen.MousePointer = 11
'   'Nuevo
'   REBLA$ = REGBLAN$("COPRECOM")
'   Call HEADERS("COPRECOM", "")
'   Call HEADERS("COPRECOM", "Periodo: " + PERIO$)
'
'   Call ABRECOMPRAS
'   SQLX$ = "SELECT * FROM OCOMDETA"
'   SQLX$ = SQLX$ + " WHERE Stat_Ocom <> 9 "
'   SQLX$ = SQLX$ + " AND CDBL(Femi_Ocom) >= " & DES1 & " "
'   SQLX$ = SQLX$ + " AND CDBL(Femi_Ocom) <=   " & HAS1 & "  "
'   If MODO% = 1 Then
'        SQLX$ = SQLX$ + " ORDER BY  Cod_Exte "
'      ElseIf MODO% = 0 Then
'        SQLX$ = SQLX$ + " ORDER BY Npro_Ocom  "
'   End If
'
'
'    Set OCOMPTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
'    On Error GoTo 90
'    With OCOMPTEMP
'        .MoveFirst
'        JJ& = 0
'   '
'        Do While (.EOF = False)
'          FECHEMI% = CVINT(!Femi_Ocom)
'            STAT% = !Stat_Ocom
'               NPRO% = !Npro_Ocom
'               CIXI% = !COD_INTE
'               DOPE$ = ""
'               '
'               NOPE% = !Cod_Oper
'               If NOPE% > 0 Then
'                 DOPE$ = DESCRIOP$(CIXI%, NOPE%)
'               End If
'               NUOC& = !Nume_Ocom
'               '
'               UMEDI$ = !Uni_Com
'               If TRIM$(UMEDI$) = "" Then
'                 UMEDI$ = UNIMED$(CIXI%)
'               End If
'               CACOM = !Cant_Ocom
'               '
'               If STAT% = 8 Then                 ' SI SE DIO POR CUMPLIDA,
'                 CACOM = !Cant_Rec
'               End If
'               '
'               PREUNI# = !Pre_Unit
'               '
'               MONEMI% = !Mone_Emis
'               If MONEMI% <= 1 Then MONEX$ = "    "
'               If MONEMI% = 2 Then MONEX$ = "U$S "
'               If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
'               TICAMBIO = 1
'               If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
'               '
'               YYY$ = REBLA$
'               Call REPLA(YYY$, MKI$(CIXI%), 1, 2)
'               Call REPLA(YYY$, MKI$(FECHEMI%), 3, 2)
'               Call REPLA(YYY$, MKI$(NPRO%), 5, 2)
'               Call REPLA(YYY$, MKI$(CIXI%), 7, 2)
'               Call REPLA(YYY$, MKI$(FECHEMI%), 9, 2)
'               Call REPLA(YYY$, DOPE$, 11, 22)
'               Call REPLA(YYY$, MKS$(NUOC&), 33, 4)
'               Call REPLA(YYY$, UMEDI$, 37, 4)
'               Call REPLA(YYY$, MKS$(CACOM), 41, 4)
'               Call REPLA(YYY$, MONEX$, 49, 4)
'               Call REPLA(YYY$, MKD$(PREUNI#), 53, 8)
'               ITOCOM# = PREUNI# * CACOM
'               Call REPLA(YYY$, MKD$(ITOCOM#), 61, 8)
'               Call REPLA(YYY$, MKD$(PREUNI# * TICAMBIO), 69, 8)
'               Call REPLA(YYY$, MKD$(ITOCOM# * TICAMBIO), 77, 8)
'               JJ& = JJ& + 1
'               Call GRAREG("COPRECOM", YYY$, JJ&)
'         .MoveNext
'      Loop
'    End With
'   '
'   Call SETULTREG("COPRECOM", JJ&)
'   Call CIERRARCH("COPRECOM")
'   Screen.MousePointer = 1
'   '
90 End Sub
'
Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
   Call BAJALDISCO
End Sub
Private Sub GENINDOC()
   '
   Call SETULTREG("OCOMENCA", 0)
   Call ABRECOMPRAS
   SQLX$ = "Select Nume_Ocom,Refe_Ocom  from ENCOCABI"
   Set OCOMP0TEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   With OCOMP0TEMP
       On Error Resume Next
       .MoveFirst
       'Valida que la tabla no este vacia
       If Err Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         On Error GoTo 0
         GoTo 90
       End If
       On Error GoTo 0
       JJ& = 0
       Do While (.EOF = False)
         XXX$ = REGBLAN$("OCOMENCA")
         Call REPLA(XXX$, MKS$(!Nume_Ocom), 1, 4)
         Call REPLA(XXX$, !Refe_OcoM, 5, 44)
         JJ& = JJ& + 1
         Call GRAREG("OCOMENCA", XXX$, JJ&)
         .MoveNext
        Loop
        Call SETULTREG("OCOMENCA", JJ&)
        Call CIERRARCH("OCOMENCA")
    End With
90 End Sub
Sub GENEMATCON()
    '
    Call SETULTREG("MACONSIG", 0)
    Call ABRESTOCKS
    SQLX$ = "SELECT * FROM MACONSIG "
    SQLX$ = SQLX$ + " ORDER BY Femi_Ocom "
    Set LIMACONTMP = Stocks.OpenRecordset(SQLX$, 4)
    '
    With LIMACONTMP
     On Error Resume Next
       .MoveFirst
       'Valida que la tabla no este vacia
       If Err Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         On Error GoTo 0
         GoTo 91
       End If
       On Error GoTo 0
       JJ& = 0
       Do While (.EOF = False)
         XXX$ = REGBLAN$("MACONSIG")
         On Error Resume Next
         Call REPLA(XXX$, MKS$(CVINT(!Femi_Ocom)), 1, 2)
         Call REPLA(XXX$, MKI$(!Npro_Ocom), 3, 2)
         Call REPLA(XXX$, MKD$(!Secc_Ejec), 5, 4)
         If IsNull(!NRO_ORFA) Then
            Call REPLA(XXX$, "", 9, 12)
         Else
            Call REPLA(XXX$, !NRO_ORFA, 9, 12)
         End If
         Call REPLA(XXX$, !Nro_Ocom, 21, 12)
         Call REPLA(XXX$, MKI$(!Cod_inte), 33, 2)
         If IsNull(!LOTE_OCOM) Then
            Call REPLA(XXX$, "", 35, 16)
         Else
            Call REPLA(XXX$, !LOTE_OCOM, 35, 16)
         End If
         Call REPLA(XXX$, !Unidad, 51, 4)
         Call REPLA(XXX$, MKD$(!Cant_Asig), 55, 8)
         Call REPLA(XXX$, MKD$(!Cant_Entreg), 63, 8)
         Call REPLA(XXX$, MKI$(CVINT(!Fecha_Entreg)), 71, 2)
         If IsNull(!NRemi_Consig) Then
            Call REPLA(XXX$, "", 73, 6)
         Else
            Call REPLA(XXX$, !NRemi_Consig, 73, 6)
         End If
         Call REPLA(XXX$, MKD(!Valo_Unit), 93, 8)
         Call REPLA(XXX$, MKI$(!MONEDA), 101, 2)
         Call REPLA(XXX$, MKD$(!Valo_Tot), 103, 8)
         Call REPLA(XXX$, MKD$(!Cant_Rendid), 111, 8)
           SALPE# = !Cant_Entreg - !Cant_Rendid
           If SALPE# < 0 Then SALPE# = 0
         Call REPLA(XXX$, MKD$(SALPE#), 119, 8)
         If IsNull(!Observa) Then
            Call REPLA(XXX$, "", 129, 16)
         Else
            Call REPLA(XXX$, !Observa, 129, 16)
         End If
         On Error GoTo 0
         JJ& = JJ& + 1
         Call GRAREG("MACONSIG", XXX$, JJ&)
         .MoveNext
        Loop
        Call SETULTREG("MACONSIG", JJ&)
        Call CIERRARCH("MACONSIG")
    End With
    '
91 End Sub

