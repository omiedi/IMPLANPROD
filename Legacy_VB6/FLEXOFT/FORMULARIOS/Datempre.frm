VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form DATEMPRE 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos Generales de la Empresa"
   ClientHeight    =   4230
   ClientLeft      =   2445
   ClientTop       =   2580
   ClientWidth     =   8580
   ControlBox      =   0   'False
   Icon            =   "Datempre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3390
      TabIndex        =   26
      Top             =   2280
      Width           =   440
   End
   Begin VB.CommandButton Command6 
      Caption         =   "."
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
      Left            =   3840
      TabIndex        =   25
      ToolTipText     =   "Selección de Provincia"
      Top             =   2280
      Width           =   175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4140
      TabIndex        =   23
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox TxtCodPostal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3390
      TabIndex        =   21
      Top             =   1920
      Width           =   3500
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2400
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K."
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
      Left            =   7320
      Picture         =   "Datempre.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
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
      Left            =   7350
      Picture         =   "Datempre.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   525
      Width           =   855
   End
   Begin VB.CommandButton Command3 
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
      Left            =   7350
      Picture         =   "Datempre.frx":0896
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   3405
      TabIndex        =   15
      Top             =   3735
      Width           =   3500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   3405
      TabIndex        =   14
      Top             =   3375
      Width           =   3500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   3405
      TabIndex        =   13
      Top             =   3015
      Width           =   3500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   3405
      TabIndex        =   12
      Top             =   2655
      Width           =   3500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   3405
      TabIndex        =   11
      Top             =   1575
      Width           =   3500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   3405
      TabIndex        =   10
      Top             =   1215
      Width           =   3500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   3405
      TabIndex        =   9
      Top             =   855
      Width           =   3500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   3405
      TabIndex        =   8
      Top             =   495
      Width           =   3500
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "Datempre.frx":0BA0
      TabIndex        =   20
      Top             =   1080
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "Datempre.frx":0C1A
      Top             =   1200
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Provincia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2310
      TabIndex        =   24
      Top             =   2295
      Width           =   990
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cód.Postal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2130
      TabIndex        =   22
      Top             =   1935
      Width           =   1140
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caja Previsional"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1605
      TabIndex        =   7
      Top             =   3750
      Width           =   1725
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Ingresos Brutos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   450
      TabIndex        =   6
      Top             =   3390
      Width           =   2835
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de C.U.I.T."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1365
      TabIndex        =   5
      Top             =   3030
      Width           =   1980
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2100
      TabIndex        =   4
      Top             =   2670
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Localidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2175
      TabIndex        =   3
      Top             =   1590
      Width           =   1110
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Domicilio: Calle y Numero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   495
      TabIndex        =   2
      Top             =   1230
      Width           =   2700
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Razon Social de la Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   300
      TabIndex        =   1
      Top             =   870
      Width           =   2970
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de Fantasia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1140
      TabIndex        =   0
      Top             =   510
      Width           =   2130
   End
End
Attribute VB_Name = "DATEMPRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    '
    LU$ = LUDAT$
    Cuit$ = TRIM$(Text1(5).TEXT)
    If Cuit$ <> "" Then
        If VALICUIT%(Cuit$) <> 1 Then
           Call MENSERR(24, "Numero de C.U.I.T.\no Válido.")
           Text1(5).SetFocus
           Exit Sub
        End If
        Text1(5).TEXT = Cuit$
    End If
    '
    N1% = FILEOPEN%(LU$ + "SVD147.EXE", 0, 41)
    Dim X As String * 41
    '
    For KK% = 1 To 8
        X = ENCRIPTA$(AJUSTI$(Text1(KK% - 1).TEXT, 41))
        Put #N1%, KK%, X
    Next KK%
    Close #N1%
    '
    Call GRACONTROL("FSTCOPOS", "CODIPOST", Me.TxtCodPostal)
    Call GRACONTROL("FSTCOPOS", "LETRAPRO", Me.Text8)
    Hide
End Sub

Private Sub Command2_Click()
    Hide
End Sub

Private Sub Command3_Click()
    Call HELPONLINE("DATEMPRE")
End Sub

Private Sub Command6_Click()
   Call SELECHO("PROVINC")
   VDEVU$ = VALACT1$("PROVINC")
   If TRIM$(VDEVU$) <> "" Then
      Text8.TEXT = VDEVU$
   End If

End Sub

Private Sub Form_Load()
    '
    Call APLICACIONSKINS(Me)
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
    End If

    LU$ = LUDAT$
    N1% = FILEOPEN(LU$ + "SVD147.EXE", 0, 41)
    Dim X As String * 41
    '
    For KK% = 1 To 8
        '
        If KK% > LOF(N1%) / 41 Then
              Text1(KK% - 1).TEXT = Space$(41)
           Else
              Get #N1%, KK%, X
              Text1(KK% - 1).TEXT = DESENCRI$(X)
              A = DESENCRI$(X)
        End If
        '
    Next KK%
    Close #N1%
    '
    Me.TxtCodPostal = CONTROL("FSTCOPOS", "CODIPOST")
    Me.Text8 = CONTROL("FSTCOPOS", "LETRAPRO")

End Sub

Private Sub Text8_Change()
    If TRIM$(Text8.TEXT) = "" Then Text2 = ""
    DPX1$ = UCase$(TRIM$(Text8.TEXT))
    If DPX1$ <> "" Then
        Text8.TEXT = DPX1$
        Text2 = ECOARCH$("PROVINC", DPX1$)
    End If

End Sub
