VERSION 5.00
Begin VB.Form SALOTES04 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Saldos por Lote"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      Height          =   6075
      Left            =   5880
      ScaleHeight     =   6015
      ScaleWidth      =   1035
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   6075
         Left            =   0
         Picture         =   "SALOTES04.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3250
      Index           =   1
      Left            =   2810
      TabIndex        =   15
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3250
      Index           =   3
      Left            =   5460
      TabIndex        =   14
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3250
      Index           =   2
      Left            =   4180
      TabIndex        =   13
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3250
      Index           =   0
      Left            =   1080
      TabIndex        =   12
      Top             =   1890
      Width           =   15
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   105
      TabIndex        =   8
      Top             =   630
      Width           =   5640
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   210
         TabIndex        =   11
         Top             =   360
         Width           =   5265
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   105
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   2205
      Width           =   5660
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00004000&
      FillColor       =   &H00400000&
      ForeColor       =   &H00400000&
      Height          =   6075
      Left            =   5880
      ScaleHeight     =   6015
      ScaleWidth      =   2295
      TabIndex        =   0
      Top             =   0
      Width           =   2355
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   9
         Top             =   4515
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   10
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
         Height          =   620
         Left            =   120
         Picture         =   "SALOTES04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   210
         Width           =   650
      End
      Begin VB.CommandButton Command8 
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
         Height          =   620
         Left            =   120
         Picture         =   "SALOTES04.frx":24E2
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprime Ficha de Movimientos de Material (Ficha Kardex)"
         Top             =   860
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "SALOTES04.frx":2B4C
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4740
         Width           =   1515
      End
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Totales"
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
      Left            =   210
      TabIndex        =   19
      Top             =   5250
      Width           =   1485
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4305
      TabIndex        =   18
      Top             =   5250
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2730
      TabIndex        =   17
      Top             =   5250
      Width           =   1450
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1050
      TabIndex        =   16
      Top             =   210
      Width           =   2010
   End
   Begin VB.Line Line2 
      X1              =   5760
      X2              =   5760
      Y1              =   1935
      Y2              =   5155
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   5760
      Y1              =   5145
      Y2              =   5145
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4830
      TabIndex        =   7
      Top             =   210
      Width           =   750
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U-M:"
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
      Left            =   3885
      TabIndex        =   6
      Top             =   255
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   250
      Width           =   960
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha       Id.Lote                       Cantidad            Saldo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   4
      Top             =   1920
      Width           =   5660
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu imprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "SALOTES04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call MENSERR(24, "No Disponible en la Presente Versión.")
   Command8.Enabled = True
End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture6.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2(1).BackColor = &HFCD7B6
    End If

End Sub

Private Sub imprimir_Click()
   Call Command8_Click
End Sub

Private Sub Label4_Change()
   Label5 = ""
   Label6 = ""
   Label1 = ""
   Label7 = ""
   List1.Clear
   '
   CIXI% = CODINT%(Label4)
   If CIXI% > 0 Then
     Label5 = UNIMED$(CIXI%)
     Label6 = DESCRIT0$(CIXI%)
     '
     Call ABRESTOCKS
     SQLX$ = " SELECT * FROM LOTINGRE "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
     SQLX$ = SQLX$ + "ORDER BY FechAlta desc "
     Set LotIngTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
     '
     SUMALOT# = 0
     SUMASAL = 0
     With LotIngTemp
25     On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           TTXX$ = Space$(128)
             FEX = CVINT(!FechAlta)
           Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
           Call REPLA(TTXX$, TRIM$(!IdenLote), 11, 16)
           Call REPLA(TTXX$, FORMATNUM$(!CanAlta, "F12,1"), 27, 12)
             SUMALOT = SUMALOT# + !CanAlta
           Call REPLA(TTXX$, FORMATNUM$(!CanAlta - !CanConsu, "F12,1"), 39, 12)
             SUMASAL = SUMASAL + !CanAlta - !CanConsu
           List1.AddItem TTXX$
         .MoveNext
         Loop
       End If
     End With
     '
     Label1 = TRIM$(FORMATNUM$(SUMALOT#, "F12.1"))
     Label7 = TRIM$(FORMATNUM$(SUMASAL, "F12.1"))
   End If
   Screen.MousePointer = 1
   '
End Sub

Private Sub List1_DblClick()
   '
   TTXX$ = List1.TEXT
   CANTING = XVALO(Mid$(TTXX$, 27, 12))
   '
   If CANTING > 0 Then
      IDLOT$ = TRIM$(Mid$(TTXX$, 11, 16))
      If IDLOT$ <> "" Then
        CONSUMLOT.Label4 = Label4
        CONSUMLOT.Label1 = IDLOT$
        CONSUMLOT.Show 1
      End If
   End If
   '
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub
