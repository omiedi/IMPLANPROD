VERSION 5.00
Begin VB.Form FORLICOMAI 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liquidación de Comisiones por Ventas"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Vend's"
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
      Left            =   5040
      Picture         =   "FORLICOMAI.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command4 
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
      Left            =   7875
      Picture         =   "FORLICOMAI.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Listado"
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
      Left            =   6930
      Picture         =   "FORLICOMAI.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   210
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Totales por Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2640
      Left            =   5040
      TabIndex        =   13
      Top             =   1260
      Width           =   4635
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
         Left            =   210
         TabIndex        =   14
         Top             =   420
         Width           =   4215
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H80000016&
      Height          =   1275
      Left            =   210
      ScaleHeight     =   1215
      ScaleWidth      =   4575
      TabIndex        =   7
      Top             =   210
      Width           =   4635
      Begin VB.Label Label6 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2940
         TabIndex        =   20
         Top             =   735
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1575
         TabIndex        =   19
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1575
         TabIndex        =   18
         Top             =   210
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Período:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   10
         Left            =   -735
         TabIndex        =   12
         Top             =   315
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Generado el:"
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
         Index           =   13
         Left            =   -735
         TabIndex        =   11
         Top             =   840
         Width           =   2220
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Hs"
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
         Index           =   12
         Left            =   3990
         TabIndex        =   10
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   225
         Left            =   2835
         TabIndex        =   9
         Top             =   840
         Width           =   15
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2730
         TabIndex        =   8
         Top             =   735
         Width           =   225
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Avance de Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   210
      TabIndex        =   2
      Top             =   1680
      Width           =   4635
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   525
         ScaleHeight     =   75
         ScaleWidth      =   3525
         TabIndex        =   16
         Top             =   1785
         Width           =   3585
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cálculo de Comisiones"
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
         Left            =   525
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1365
         Width           =   2745
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Valorización por Producto"
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
         Left            =   525
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1050
         Width           =   2745
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Detalle de Venta por Factura"
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
         Left            =   525
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   735
         Width           =   3690
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Selección de Facturación del Período"
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
         Left            =   525
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   420
         Width           =   3585
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cálculo"
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
      Left            =   5985
      Picture         =   "FORLICOMAI.frx":0C7E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   210
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
      Left            =   8820
      Picture         =   "FORLICOMAI.frx":0F88
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   855
   End
End
Attribute VB_Name = "FORLICOMAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    Check1.Value = 0
End Sub

Private Sub Check2_Click()
    Check2.Value = 0
End Sub

Private Sub Check3_Click()
    Check3.Value = 0
End Sub

Private Sub Check4_Click()
    Check4.Value = 0
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    Call CALCOMCHE
    'Call CALCOMVEN     ESTA ERA LA MODALIDAD SETEADA
    'Call CALCOMAI
99  Command2.Enabled = True
    '
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    '
    If FORLICOMAI.List1.ListIndex >= 0 Then
        TXT$ = FORLICOMAI.List1.Text
        VENDE% = Val(Left$(TXT$, 3))
        NOVEN$ = "Clientes de Direccion"
        If VENDE% > 0 Then
            NOVEN$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        End If
        '
        CORRESP$ = AJUSTI$(FORLICOMAI.Label3.Caption, 12)
        '
        Screen.MousePointer = 11
        Call COPYSTRU("LICOMCHE", "LICOMCHX")
        VENDE% = Val(Left$(List1.Text, 4))
        VENDX$ = Mid$(List1.Text, 6, 17)
        '
        JJ& = 0
        FIN& = ULTREG&("LICOMCHE")
        For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           X$ = REGLEIDO$("LICOMCHE", U&)
           If Asc(Left$(X$, 1)) = VENDE% Then
              JJ& = JJ& + 1
              Call GRAREG("LICOMCHX", X$, JJ&)
           End If
        Next U&
        Call SETULTREG("LICOMCHX", JJ&)
        Call CIERRARCH("LICOMCHX")
        Call CIERRARCH("LICOMCHE")
        '
        Call HEADERS("LICOMCHX", "")
        Call HEADERS("LICOMCHX", "Corresponde a: " + CORRESP$)
        Call HEADERS("LICOMCHX", "Vendedor: " + NOVEN$)
        Call FINAL("*SOCOCHE")
        Screen.MousePointer = 1
    End If
    Command3.Enabled = True
    '
End Sub

Private Sub Command4_Click()
    Call HELPONLINE("LICOMIS")
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Call FINAL("*VENDEDOR")
   Command5.Enabled = True
End Sub

Private Sub Form_Activate()
    If List1.ListCount < 1 Then
        Command3.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Show
    Refresh
    FORSELMEC.Show 1
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub

Private Sub MESAN_GOTFOCUS()
    Command2.SetFocus
End Sub

Private Sub Option1_Click()
    If Option2.Value = False Then
        Option1.Value = False
    End If
End Sub

Private Sub Option2_Click()
    List1.Clear
    Command3.Enabled = False
    Command4.Enabled = False
End Sub

Private Sub Option3_Click()
    List1.Clear
    Command3.Enabled = False
    Command4.Enabled = False
    Option1.Value = False
End Sub

Private Sub Option4_Click()
    List1.Clear
    Command3.Enabled = False
    Command4.Enabled = False
    Option1.Value = False
End Sub

Private Sub Option5_Click()
    List1.Clear
    Command3.Enabled = False
    Command4.Enabled = False
    Option1.Value = False
End Sub

Private Sub Text4_GOTFOCUS(Index As Integer)
    Command2.SetFocus
End Sub

Private Sub List1_DblClick()
   Screen.MousePointer = 11
   CORRESP$ = AJUSTI$(FORLICOMAI.Label3.Caption, 12)
   Call COPYSTRU("LICOMCHE", "LICOMCHX")
   VENDE% = Val(Left$(List1.Text, 4))
   VENDX$ = Mid$(List1.Text, 6, 17)
   '
   JJ& = 0
   FIN& = ULTREG&("LICOMCHE")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("LICOMCHE", U&)
      If Asc(Left$(X$, 1)) = VENDE% Then
         JJ& = JJ& + 1
         Call GRAREG("LICOMCHX", X$, JJ&)
      End If
   Next U&
   Call SETULTREG("LICOMCHX", JJ&)
   Call CIERRARCH("LICOMCHX")
   Call CIERRARCH("LICOMCHE")
   '
   Call HEADERS("LICOMCHX", "")
   Call HEADERS("LICOMCHX", "Corresponde a: " + CORRESP$)
   Call HEADERS("LICOMCHX", "Vendedor: " + VENDX$)
   Call FINAL("*SOCOCHE")
   Screen.MousePointer = 1
End Sub
