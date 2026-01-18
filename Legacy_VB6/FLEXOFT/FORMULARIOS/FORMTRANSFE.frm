VERSION 5.00
Begin VB.Form FORMTRANSFE 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transferencias de Stocks"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5610
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command27 
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
      Height          =   855
      Left            =   2100
      Picture         =   "FORMTRANSFE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command26 
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
      Height          =   855
      Left            =   210
      Picture         =   "FORMTRANSFE.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Kardex"
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
      Left            =   1155
      Picture         =   "FORMTRANSFE.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command6 
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
      Left            =   3570
      Picture         =   "FORMTRANSFE.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   210
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Operaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   17
      Top             =   1365
      Width           =   4110
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   18
         Top             =   420
         Width           =   2535
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
            Left            =   -20
            Picture         =   "FORMTRANSFE.frx":0D60
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Informar Trans- ferencias de Stock"
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
            TabIndex        =   19
            Top             =   105
            Width           =   1800
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Consultas - Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4005
      Left            =   210
      TabIndex        =   15
      Top             =   2835
      Width           =   5160
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   210
         ScaleHeight     =   75
         ScaleWidth      =   4680
         TabIndex        =   26
         Top             =   3675
         Width           =   4740
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Rango de Fechas"
         Height          =   915
         Left            =   210
         TabIndex        =   23
         Top             =   1575
         Width           =   4740
         Begin VB.CommandButton BOTSEL1 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   4305
            Picture         =   "FORMTRANSFE.frx":106A
            TabIndex        =   6
            Top             =   420
            Width           =   175
         End
         Begin VB.CommandButton BOTSEL 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1995
            Picture         =   "FORMTRANSFE.frx":1374
            TabIndex        =   4
            Top             =   420
            Width           =   175
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   945
            TabIndex        =   3
            Text            =   " "
            Top             =   420
            Width           =   1065
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   3255
            TabIndex        =   5
            Text            =   " "
            Top             =   420
            Width           =   1065
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
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
            Left            =   0
            TabIndex        =   25
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
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
            Left            =   2415
            TabIndex        =   24
            Top             =   420
            Width           =   750
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Rango Numérico"
         Height          =   915
         Left            =   210
         TabIndex        =   20
         Top             =   420
         Width           =   4740
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   3255
            TabIndex        =   2
            Top             =   420
            Width           =   1065
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   945
            TabIndex        =   1
            Top             =   420
            Width           =   1065
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
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
            Left            =   2415
            TabIndex        =   22
            Top             =   420
            Width           =   750
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
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
            Left            =   0
            TabIndex        =   21
            Top             =   420
            Width           =   855
         End
      End
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
         Left            =   4095
         Picture         =   "FORMTRANSFE.frx":167E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2830
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Salida"
         Height          =   705
         Left            =   210
         TabIndex        =   16
         Top             =   2730
         Width           =   3690
         Begin VB.OptionButton Option5 
            BackColor       =   &H00C0FFFF&
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
            Left            =   315
            TabIndex        =   7
            Top             =   315
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0FFFF&
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
            Left            =   2100
            TabIndex        =   8
            Top             =   315
            Width           =   1275
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
      Height          =   855
      Left            =   4540
      Picture         =   "FORMTRANSFE.frx":1988
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   210
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   4305
      Picture         =   "FORMTRANSFE.frx":1AD2
      Stretch         =   -1  'True
      Top             =   2070
      Width           =   5250
   End
End
Attribute VB_Name = "FORMTRANSFE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTSEL_Click(Index As Integer)
    Call SELECHO("SELFECHA")
    Text2.TEXT = VALACT1$("SELFECHA")
End Sub

Private Sub BOTSEL1_Click(Index As Integer)
    Call SELECHO("SELFECHA")
    Text1.TEXT = VALACT1$("SELFECHA")
End Sub

Private Sub Command1_Click()
    Hide
    Call FRESHALL
End Sub

Private Sub Command2_Click()
    DESNUM& = 1: If Val(Text7.TEXT) > 0 Then DESNUM& = Val(Text7.TEXT)
    HASNUM& = 999999
    If TRIM$(Text8.TEXT) <> "" Then
        If Val(Text8.TEXT) >= DESNUM& Then HASNUM& = Val(Text8.TEXT)
    End If
    DESFE% = 0: If FECHANUM(TRIM$(Text2.TEXT)) > 0 Then DESFE% = FECHANUM(TRIM$(Text2.TEXT))
    HASFE% = 32767
    If TRIM$(Text1.TEXT) <> "" Then
        If FECHANUM(TRIM$(Text1.TEXT)) >= DESFE% Then HASFE% = FECHANUM(TRIM$(Text1.TEXT))
    End If
    '
    Screen.MousePointer = 11
    Call BLANARCH("LISTRAJ")
    FIN& = ULTREG&("TRANSFE")
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        X$ = REGLEIDO$("TRANSFE", I&)
        FF% = CVI(Left$(X$, 2))
        NRO& = CVS(Mid$(X$, 45, 4))
        If FF% >= DESFE% Then
          If FF% <= HASFE% Then
            NRO& = CVS(Mid$(X$, 45, 4))
            If NRO& >= DESNUM& Then
              If NRO& <= HASNUM& Then
                REFE$ = Mid$(X$, 9, 16)
                CI1% = CVI(Mid$(X$, 5, 2))
                DP1% = Asc(Mid$(X$, 3, 1))
                CI2% = CVI(Mid$(X$, 7, 2))
                DP2% = Asc(Mid$(X$, 4, 1))
                CAN# = CVS(Mid$(X$, 41, 4))
                Y$ = REGBLAN$("LISTRAJ")
                Call REPLA(Y$, MKI$(FF%), 1, 2)
                Call REPLA(Y$, REFE$, 3, 48)
                Call REPLA(Y$, MKS$(NRO&), 51, 4)
                Call REPLA(Y$, MKI$(CI1%), 55, 2)
                Call REPLA(Y$, UNIMED$(CI1%), 57, 4)
                Call REPLA(Y$, MKD$(0), 61, 8)
                Call REPLA(Y$, MKD$(CAN#), 69, 8)
                Call REPLA(Y$, Chr$(DP1%), 80, 1)
                Call REGAPP("LISTRAJ", Y$)
                Call REPLA(Y$, MKI$(CI2%), 55, 2)
                Call REPLA(Y$, UNIMED$(CI2%), 57, 4)
                Call REPLA(Y$, MKD$(CAN#), 61, 8)
                Call REPLA(Y$, MKD$(0), 69, 8)
                Call REPLA(Y$, Chr$(DP2%), 80, 1)
                Call REGAPP("LISTRAJ", Y$)
              End If
            End If
          End If
        End If
    Next I&
    Call CIERRARCH("LISTRAJ")
    Screen.MousePointer = 1
    '
    OPCI% = 1
    If Option4.Value = True Then OPCI% = 2
    If OPCI% = 1 Then
        Call OPNOIN("")
        'Call CONECRUN("*COTRAST", "")
      ElseIf OPCI% = 2 Then
        Call CONECRUN("*LITRAST", "")
    End If
    '
End Sub

Private Sub Command26_Click()
    OPTRAJUST.Show 1
End Sub

Private Sub Command27_Click()
    Call ACCDIR("TRANSFE")
End Sub

Private Sub Command3_Click()
    Call TRASTOCK
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call GENKARDEX("")
    'Call CONECRUN("*KARDEX", "")
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Call OPNOIN("")
End Sub

Private Sub Command6_Click()
    Call HELPONLINE("TRANSFE")
End Sub

Private Sub Form_Load()
   Option5.Value = True
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("SELFECHA")
    Text1.TEXT = VALACT1$("SELFECHA")
End Sub

Private Sub Text2_DblClick()
    Call SELECHO("SELFECHA")
    Text2.TEXT = VALACT1$("SELFECHA")
End Sub
