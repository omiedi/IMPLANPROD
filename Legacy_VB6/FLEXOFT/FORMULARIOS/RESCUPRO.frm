VERSION 5.00
Begin VB.Form RESCUPRO 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F0F0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmenes de Cuenta Corriente de Acreedores (Gestion de Compras)"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Re-apli"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9975
      Picture         =   "RESCUPRO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Re-Aplicacion de Creditos y Pagos a Cuenta"
      Top             =   3990
      Width           =   855
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   9975
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   36
      Top             =   6930
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   2430
      Left            =   105
      ScaleHeight     =   2370
      ScaleWidth      =   9615
      TabIndex        =   19
      Top             =   2415
      Width           =   9675
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   330
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   10515
         Begin VB.Line Line3 
            X1              =   7665
            X2              =   7665
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "   Fecha       Documento             Número                        Debe               Haber                Saldo              "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   0
            TabIndex        =   21
            Top             =   50
            Width           =   10410
         End
         Begin VB.Line Line1 
            X1              =   1155
            X2              =   1155
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Line Line4 
            X1              =   3100
            X2              =   3100
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Line Line5 
            X1              =   4725
            X2              =   4725
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Line Line6 
            X1              =   6090
            X2              =   6090
            Y1              =   0
            Y2              =   1050
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0F0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   6500
         Left            =   0
         TabIndex        =   22
         Top             =   -1050
         Width           =   10515
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2085
            Left            =   -50
            TabIndex        =   23
            Top             =   1365
            Width           =   9670
         End
         Begin VB.Line Line7 
            X1              =   1155
            X2              =   1155
            Y1              =   1470
            Y2              =   3255
         End
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9975
      Picture         =   "RESCUPRO.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Resúmenes de Cuenta Corriente por Impresora"
      Top             =   2415
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Notas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9975
      Picture         =   "RESCUPRO.frx":07B4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   2800
      Left            =   105
      ScaleHeight     =   2745
      ScaleWidth      =   9615
      TabIndex        =   25
      Top             =   4990
      Width           =   9675
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   5630
         Left            =   0
         TabIndex        =   28
         Top             =   315
         Width           =   10515
         Begin VB.ListBox List2 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2085
            Left            =   -30
            TabIndex        =   29
            Top             =   0
            Width           =   9675
         End
         Begin VB.Line Line19 
            X1              =   7770
            X2              =   7770
            Y1              =   2100
            Y2              =   2415
         End
         Begin VB.Line Line18 
            X1              =   6300
            X2              =   6300
            Y1              =   2100
            Y2              =   2415
         End
         Begin VB.Line Line17 
            X1              =   4830
            X2              =   4830
            Y1              =   2100
            Y2              =   2415
         End
         Begin VB.Line Line16 
            X1              =   3570
            X2              =   3570
            Y1              =   2100
            Y2              =   2415
         End
         Begin VB.Line Line15 
            X1              =   0
            X2              =   9660
            Y1              =   2420
            Y2              =   2420
         End
         Begin VB.Line Line14 
            BorderWidth     =   2
            X1              =   0
            X2              =   9660
            Y1              =   2085
            Y2              =   2085
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Total General"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   25
            TabIndex        =   32
            Top             =   2130
            Width           =   1695
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4330
            TabIndex        =   31
            Top             =   2130
            Width           =   1695
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   7455
            TabIndex        =   30
            Top             =   2130
            Width           =   1695
         End
         Begin VB.Line Line12 
            X1              =   1890
            X2              =   1890
            Y1              =   1575
            Y2              =   3360
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   330
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   10515
         Begin VB.Line Line11 
            X1              =   6300
            X2              =   6300
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Line Line10 
            X1              =   4830
            X2              =   4830
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Line Line9 
            X1              =   3570
            X2              =   3570
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Line Line8 
            X1              =   1890
            X2              =   1890
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   " Documento             Número                 Emision            Importe         Vencim.                 Saldo              "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   0
            TabIndex        =   27
            Top             =   50
            Width           =   10410
         End
         Begin VB.Line Line2 
            X1              =   7770
            X2              =   7770
            Y1              =   0
            Y2              =   1050
         End
      End
   End
   Begin VB.ListBox LINDICE 
      Height          =   1035
      Left            =   9870
      Sorted          =   -1  'True
      TabIndex        =   24
      Top             =   5460
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F0F0C0&
      Caption         =   "Datos de la Cuenta Activa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   105
      TabIndex        =   5
      Top             =   210
      Width           =   9675
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7560
         TabIndex        =   34
         Text            =   " "
         Top             =   840
         Width           =   1905
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7560
         TabIndex        =   17
         Text            =   " "
         Top             =   1560
         Width           =   1905
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1365
         TabIndex        =   14
         Text            =   " "
         Top             =   1560
         Width           =   3690
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   210
         TabIndex        =   0
         Text            =   " "
         Top             =   800
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1050
         TabIndex        =   1
         Top             =   800
         Width           =   180
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1365
         TabIndex        =   11
         Text            =   " "
         Top             =   420
         Width           =   3690
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1365
         TabIndex        =   10
         Text            =   " "
         Top             =   800
         Width           =   3690
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1365
         TabIndex        =   9
         Text            =   " "
         Top             =   1180
         Width           =   750
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2205
         TabIndex        =   8
         Text            =   " "
         Top             =   1180
         Width           =   2850
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5355
         TabIndex        =   7
         Text            =   " "
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5355
         TabIndex        =   6
         Text            =   " "
         Top             =   800
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cuenta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   7560
         TabIndex        =   35
         Top             =   525
         Width           =   1800
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Acreedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   7560
         TabIndex        =   18
         Top             =   1260
         Width           =   1800
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax-Nbr:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   5355
         TabIndex        =   16
         Top             =   1260
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   15
         Top             =   1575
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   13
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   5355
         TabIndex        =   12
         Top             =   480
         Width           =   960
      End
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
      Height          =   855
      Left            =   9975
      Picture         =   "RESCUPRO.frx":0ABE
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1365
      Width           =   855
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
      Height          =   855
      Left            =   9975
      Picture         =   "RESCUPRO.frx":0DC8
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Finaliza la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.Line Line24 
      X1              =   9765
      X2              =   11000
      Y1              =   7250
      Y2              =   7250
   End
   Begin VB.Line Line23 
      X1              =   105
      X2              =   9765
      Y1              =   4845
      Y2              =   4845
   End
   Begin VB.Line Line22 
      X1              =   9775
      X2              =   9775
      Y1              =   2415
      Y2              =   4830
   End
   Begin VB.Line Line21 
      X1              =   9775
      X2              =   9775
      Y1              =   4960
      Y2              =   7770
   End
   Begin VB.Line Line20 
      X1              =   9765
      X2              =   10920
      Y1              =   7770
      Y2              =   7770
   End
   Begin VB.Line Line13 
      X1              =   4830
      X2              =   6090
      Y1              =   3570
      Y2              =   3990
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   9400
      Picture         =   "RESCUPRO.frx":0F12
      Top             =   7260
      Width           =   2010
   End
End
Attribute VB_Name = "RESCUPRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
    For K% = 0 To Forms.Count - 1
        If Forms(K%).Name = PROLLAMA$ Then
            On Error Resume Next
            Forms(K%).Show
            PROLLAMA$ = ""
            On Error GoTo 0
            Exit Sub
          ElseIf Forms(K%).Name = "COGECOMP" Then
            COGECOMP.Show
            Exit Sub
        End If
    Next K%
End Sub

Private Sub Command2_Click()
    '
    Call SELECHO("PROVED1")
    NC% = Val(TRIM$(VALACT1$("PROVED1")))
    Text1.TEXT = TRIM$(Str$(NC%))
    Call CARESCUEPRO
    '
End Sub

Private Sub Command29_Click()
   Call HELPONLINE("COGECOM")
End Sub

Sub Command3_Click()
    '
    Screen.MousePointer = 11
    '
    If ULTREG&("PROVED1") > 0 Then
        X$ = REGLEIDO$("PROVED1", 1)      ' fuerza la
        Call GRAREG("PROVED1", X$, 1)     ' reindexacion
        Call CIERRARCH("PROVED1")         ' del archivo
    End If
    '
    Screen.MousePointer = 11
    '
    If Int(Val(Text1.TEXT)) > 0 Then
        If Int(Val(Text1.TEXT)) <= 32767 Then
            NCMIN% = Int(Val(Text1.TEXT))
            NCMAX% = NCMIN%
            GoTo 10
        End If
    End If
    '
    NCMIN% = 32767: NCMAX% = 0
    FIN& = ULTREG&("PROVED1")
    For U& = 1 To FIN&
        NCX% = CVI(Left$(REGLEIDO$("PROVED1", U&), 2))
        If NCX% > 0 Then
            If NCX% > NCMAX% Then NCMAX% = NCX%
            If NCX% < NCMIN% Then NCMIN% = NCX%
        End If
    Next U&
    '
10  Screen.MousePointer = 1
    RESCUEPRO.Text10.TEXT = TRIM$(Str$(NCMIN%))
    RESCUEPRO.Text9.TEXT = RASOCLI$((-1) * NCMIN%)
    RESCUEPRO.Text8.TEXT = TRIM$(Str$(NCMAX%))
    RESCUEPRO.Text7.TEXT = RASOCLI$((-1) * NCMAX%)
    '
    HOO% = HOY(HO$)
    RESCUEPRO.Text6.TEXT = "01" + Mid$(HO$, 3)
    RESCUEPRO.Text5.TEXT = HO$
    '
    RESCUEPRO.Show 1
    '
End Sub

Private Sub Command4_Click()
    '
    Command4.Enabled = False
    NC% = Val(Text1.TEXT)
    If NC% < 1 Or ECOARCH$("PROVED1", MKI$(NC%)) = "" Then
       Call MENSERR(24, "Número de Cuenta no Válido")
       Text1.SetFocus
       GoTo 99
    End If
    '
    If TICUEPRO%((-1) * NC%) <> 0 Then
       Call MENSERR(24, "Proceso Imposible - Cuenta de Gastos")
       GoTo 99
    End If
    '
    If COMALTER%("El Proceso Invocado Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Pagos y Créditos\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar") <> 2 Then
       GoTo 99
    End If
    '
    Call REAPLICREPRO(NC%)
    '
99  Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    NCLI$ = TRIM$(Text1.TEXT)
    If NCLI$ <> "" Then
        Command5.Enabled = False
        Call CONECRUN("*NOTAPRO/" + NCLI$, "Notas de Proveedores")
        Command5.Enabled = True
      Else
        Call MENSERR(24, "Falta Ingresar Número de Proveedor")
    End If
End Sub

Private Sub Form_Activate()
    If Val(Text1.TEXT) > 0 Then
        NC% = Val(Text1.TEXT)
      ElseIf NC% > 0 Then
        Text1.TEXT = NC%
    End If
    List1.Clear
    List2.Clear
    If NC% > 0 Then
      If ECOARCH$("PROVED1", MKI$(NC%)) <> "" Then
        Call CARESCUEPRO
      End If
    End If
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    For K% = 0 To Forms.Count - 1
        If Forms(K%).Name = PROLLAMA$ Then
            Forms(K%).Show
            PROLLAMA$ = ""
            Exit Sub
          ElseIf Forms(K%).Name = "COGECOMP" Then
            COGECOMP.Show
            Exit Sub
        End If
    Next K%
End Sub

Private Sub List1_Click()
   RELIS$ = List1.TEXT
   REG& = Val(Mid$(RELIS$, 81, 8))
   If Left$(Text2.TEXT, 6) = "VARIOS" Then
     Text9.TEXT = ""
     If REG& > 0 Then
       If REG& <= ULTREG&("CUECORR") Then
         TS$ = REGLEIDO$("CUECORR", REG&)
         If Mid$(TS$, 295, 6) <> String$(6, 0) Then
           Text9.TEXT = Mid$(TS$, 25, 24)
         End If
       End If
     End If
   End If
End Sub

Private Sub List1_DBLClick()
   '
   RELIS$ = List1.TEXT
   TIDOCUM$ = TRIM$(Mid$(RELIS$, 11, 16))
   If UCase$(TIDOCUM$) <> "ORDEN DE PAGO" Then
      Call MUEFACO
      Exit Sub
   End If
   '
   NPX& = Val(Mid$(RELIS$, 32, 8))
   NUDOCU$ = " " + TRIM$(Str$(NPX&)) + " -"
   Screen.MousePointer = 11
   '
   For UI& = ULTREG&("PDOCLST") To 1 Step -1
     XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
     If InStr(XX$, TIDOCUM$) > 4 Then
       PPXX% = InStr(XX$, "Nro.")
       If PPXX% > 0 Then
         XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
       End If
       If InStr(XX$, NUDOCU$) > 4 Then
         DOCUNU& = CVS(Left$(XX$, 4))
         GoTo 5
       End If
     End If
   Next UI&
   Screen.MousePointer = 1
3  Call COMUNI("Imposible Mostrar Documento")
   Exit Sub
   '
5  Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
   '
End Sub

Private Sub Text1_DblClick()
    '
    Call SELECHO("PROVED1")
    NC% = Val(TRIM$(VALACT1$("PROVED1")))
    Text1.TEXT = TRIM$(Str$(NC%))
    Call CARESCUEPRO
    '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Val(Text1.TEXT) > 0 Then
          If Val(Text1.TEXT) <= 32767 Then
             NC% = Val(Text1.TEXT)
             If ECOARCH$("PROVED1", MKI$(NC%)) <> "" Then
                Call CARESCUEPRO
                Exit Sub
             End If
          End If
       End If
       '
       Beep
       Text1.TEXT = ""
       '
    End If
    '
End Sub


Private Sub Text10_Click()
   Text1.SetFocus
End Sub

Private Sub Text2_CLICK()
   Text1.SetFocus
End Sub

Private Sub Text3_Click()
   Text1.SetFocus
End Sub

Private Sub Text4_CLICK()
   Text1.SetFocus
End Sub

Private Sub Text5_Click()
   Text1.SetFocus
End Sub

Private Sub Text6_Click()
   Text1.SetFocus
End Sub

Private Sub Text7_CLICK()
   Text1.SetFocus
End Sub

Private Sub Text8_Click()
   Text1.SetFocus
End Sub

Private Sub Text9_CLICK()
   Text1.SetFocus
End Sub

Sub MUEFACO()
   '
   RELIS$ = List1.TEXT
   REG& = Val(Mid$(RELIS$, 81, 8))
   If REG& > 0 Then
     If REG& <= ULTREG&("CUECORR") Then
       X$ = REGLEIDO$("CUECORR", REG&)
       NF$ = Mid$(X$, 7, 18)
       NC% = CVI(Mid$(X$, 5, 2))
       IMPO# = Abs(CVD(Mid$(X$, 115, 8)) - CVD(Mid$(X$, 123, 8)))
       Y6$ = Mid$(X$, 199, 80)
       GoTo 20
     End If
   End If
   Call MENSERR(24, "Documento no Encontrado - Consulte.")
   GoTo 99
   '
20 ATRI$ = TRIM$(NF$) + " - " + TRIM$(RASOCLI$((-1) * NC%)) + " ($ " + TRIM$(CSTRING$(MKD$(IMPO#), 4, 12, 2, 2)) + ")"
   '
   Call BLANARCH("!CARICOM")
   IMPO# = 0
   For J% = 1 To 8
      IPA# = 0
      CUE% = CVI(Mid$(Y6$, 10 * J% - 9, 2))
      If CUE% > 0 Then
         IPA# = CVD(Mid$(Y6$, 10 * J% - 7, 8))
         Y$ = REGBLAN$("!CARICOM")
         Call REPLA(Y$, MKI$(CUE%), 1, 2)
         Call REPLA(Y$, MKD$(IPA#), 3, 8)
         IMPO# = IMPO# + Abs(IPA#)
         Call REGAPP("!CARICOM", Y$)
      End If
   Next J%
   Call CIERRARCH("!CARICOM")
   '
30 VTB% = VENTABU%("CARIFACO/TITUPAN=" + ATRI$)
   If VTB% < 0 Then GoTo 99
   '
   TOTIMPU# = 0: Y6$ = ""
   For U& = 1 To ULTREG&("!CARICOM")
       Y$ = REGLEIDO$("!CARICOM", U&)
       CUEB$ = Left$(Y$, 2)
       If CVI(CUEB$) > 0 Then
          DECUE$ = ECOARCH$("CUECON", CUEB$)
          If DECUE$ <> "" Then
             IMBB$ = Mid$(Y$, 3, 8)
             If CVD(IMBB$) < 0 Then
                Call MENSERR(24, "No se Admiten Valores Negativos")
                GoTo 30
             End If
             Y6$ = Y6$ + CUEB$ + IMBB$
             TOTIMPU# = TOTIMPU# + CVD(IMBB$)
          End If
       End If
   Next U&
   '
   If Abs(IMPO# - TOTIMPU#) >= 0.005 Then
      Call MENSERR(24, "Suma de Imputaciones Difiere de Total")
      GoTo 30
   End If
   '
   Y6$ = Y6$ + String$(80, 0)
   X$ = REGLEIDO$("CUECORR", REG&)
   Call REPLA(X$, Y6$, 199, 80)
   Call GRAREG("CUECORR", X$, REG&)
   '
99 Call CIERRARCH("CUECORR")
   '
End Sub

