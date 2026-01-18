VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form RIGIDIEL07 
   BackColor       =   &H00E0E0D0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe de Rigidez Dielectrica"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8580
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   650
      Left            =   7770
      Picture         =   "RIGIDIEL07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Aprobar"
      Top             =   1400
      Width           =   650
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   1380
      Left            =   7665
      ScaleHeight     =   1320
      ScaleWidth      =   1110
      TabIndex        =   3
      Top             =   0
      Width           =   1170
      Begin VB.CommandButton Command1 
         Height          =   640
         Left            =   105
         Picture         =   "RIGIDIEL07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Salir"
         Top             =   50
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   7
         Top             =   700
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "RIGIDIEL07.frx":0454
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   880
         Width           =   1515
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0D0&
      Caption         =   "Rigidez Dielectrica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   7470
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6720
         TabIndex        =   15
         Top             =   1155
         Width           =   645
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4830
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "60"
         Top             =   1155
         Width           =   645
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1785
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "1500"
         Top             =   1170
         Width           =   1065
      End
      Begin VB.Label Label98 
         Alignment       =   2  'Center
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
         Height          =   360
         Left            =   1785
         TabIndex        =   18
         ToolTipText     =   "Doble Click Para Conmutar"
         Top             =   525
         Width           =   2430
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Veredicto:"
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
         Left            =   5775
         TabIndex        =   12
         Top             =   1125
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo en Segundos:"
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
         Left            =   2650
         TabIndex        =   11
         Top             =   1125
         Width           =   2115
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tension Utilizada:"
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
         Left            =   180
         TabIndex        =   10
         Top             =   1125
         Width           =   1590
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Analizador de Rigidez Dieléctrica"
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
         Left            =   4305
         TabIndex        =   9
         Top             =   525
         Width           =   3165
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   630
         TabIndex        =   4
         Top             =   735
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Instrumento:"
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
         TabIndex        =   2
         Top             =   525
         Width           =   1170
      End
   End
   Begin VB.CommandButton Command7 
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
      Left            =   5600
      Picture         =   "RIGIDIEL07.frx":2376
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3990
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7035
      OleObjectBlob   =   "RIGIDIEL07.frx":2680
      Top             =   105
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "RIGIDIEL07.frx":28B4
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   6090
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   5040
      Picture         =   "RIGIDIEL07.frx":292E
      Top             =   6950
      Width           =   2010
   End
End
Attribute VB_Name = "RIGIDIEL07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   
   Hide
   
   Command1.Enabled = True
   
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   '
   If TRIM$(Text2) = "" Then ' AHORA ES FIJO SEGUN FABIAN CORNEJ0
      Call COMUNI("Falta Tensión Utilizada")
      GoTo 99
   End If
   '
   If XVALO(Text3) < 1 Then  ' AHORA ES FIJO SEGUN FABIAN CORNEJ0
      Call COMUNI("Tiempo en Segundos No Válido")
      GoTo 99
   End If

   If TRIM$(Label98) = "" Then
      Call COMUNI("Falta Instrumento")
      GoTo 99
   End If

   If TRIM$(Text5) = "" Then
      Call COMUNI("Falta Veredicto")
      GoTo 99
   End If
   
   If TRIM$(UCase$(Text5)) <> "NC" And TRIM$(UCase$(Text5)) <> "A" Then
      Call COMUNI("Veredicto No Válido \Valores Posibles 'NC o A'")
      GoTo 99
   End If
   
    'oculto el frm
   Hide
99 Command14.Enabled = True
  
End Sub

Private Sub Label7_Click()

End Sub

Private Sub Form_Load()
       
'       SQLX$ = "SELECT * FROM ENSAR072"
'       NUSERIE$ = COFUTIE07.Text8
'       ULTREVIS% = XVALO(COFUTIE07.Label13)
'       CONDI$ = CONDI$ + "NumeroSerie ='" & NUSERIE$ & "'"
'       CONDI$ = CONDI$ + " AND REVISION = " & ULTREVIS% & ""
'       If CantRegistros("ENSAR072", CONDI$, "NOMESS") > 0 Then
'            Label98 = VALOBADA("ENSAR072", "RIDIEINSTRU", CONDI$, "NOMESS")
'            Text2 = VALOBADA("ENSAR072", "RiDieTension", CONDI$, "NOMESS")
'            Text3 = VALOBADA("ENSAR072", "RIDIETIEMPO", CONDI$, "NOMESS")
'            Text5 = VALOBADA("ENSAR072", "RIDIEVEREDICTO", CONDI$, "NOMESS")
'       End If
End Sub

Private Sub Label98_DblClick()
    Call SELECHO("INSTRUME")
    CODINSTRU% = XVALO(VALACT1$("INSTRUME"))
    If CODINSTRU% < 1 Then Exit Sub
    Label98 = ECOARCH$("INSTRUME", Chr$(CODINSTRU%))

'        If TRIM$(Label98) = "" Then
'           Label98 = "I133"
'        ElseIf TRIM$(Label98) = "I133" Then
'           Label98 = "I394"
'        ElseIf TRIM$(Label98) = "I394" Then
'           Label98 = "I395"
'        ElseIf TRIM$(Label98) = "I395" Then
'           Label98 = "I133"
'        End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text3.SetFocus
    End If

End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text5.SetFocus
    End If
End Sub




