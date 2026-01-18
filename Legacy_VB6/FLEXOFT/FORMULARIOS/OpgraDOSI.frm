VERSION 5.00
Begin VB.Form OPGRADOSI 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bases de Datos Gráficas"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10710
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
   ScaleHeight     =   2370
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Archivo Imagen Isologo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   4410
      TabIndex        =   8
      Top             =   210
      Width           =   5055
      Begin VB.TextBox Text3 
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
         Left            =   2310
         MaxLength       =   20
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   420
         Width           =   2550
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del Archivo:"
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
         Left            =   -1260
         TabIndex        =   10
         Top             =   525
         Width           =   3480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ubicación de Archivos Gráficos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   5
      Top             =   1260
      Width           =   9255
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
         Height          =   360
         Left            =   735
         MaxLength       =   40
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   420
         Width           =   8325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta:"
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
         TabIndex        =   7
         Top             =   525
         Width           =   645
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Tipo de Archivo (JPG - GIF - etc)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   4005
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
         Height          =   360
         Left            =   2835
         MaxLength       =   8
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " *.*"
         Top             =   420
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código Tipo de Archivo:"
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
         Left            =   -735
         TabIndex        =   4
         Top             =   525
         Width           =   3480
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
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
      Left            =   9660
      Picture         =   "OpgraDOSI.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1365
      Width           =   855
   End
   Begin VB.CommandButton command1 
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
      Left            =   9660
      Picture         =   "OpgraDOSI.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   315
      Width           =   855
   End
End
Attribute VB_Name = "OPGRADOSI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FORMUNAME(5), FORMUECO(5)

Private Sub BOTOPENFORM_Click(Index As Integer)
    Call CONECRUN("PRIFORMS/" + TRIM$(FORMUNAME(Index).Text), "Formularios de Impresion")
    If EXISTE%("PRIFORMS.EXE") < 1 Then
        Call FINAL("")
    End If
End Sub

Private Sub BOTSELFORM_Click(Index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        FORMUNAME(Index).Text = VDEVU$
        FORMUECO(Index).Caption = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    TIARGRA$ = TRIM$(Text1.Text)
    Call GRACONTROL("", "TIARGRA", Text1.Text)
    '
    RUTARGRA$ = TRIM$(Text2.Text)
    Call GRACONTROL("", "RUTARGRA", Text2.Text)
    '
    ISOLOGO$ = TRIM$(Text3.Text)
    Call GRACONTROL("", "ISOLOGO", Text3.Text)
    '
    For KK% = 0 To 5
       On Error Resume Next
       Call GRACONTROL("FORGRAFI", "DOCUGRA" + TRIM$(Str$(KK%)), FORMUNAME(KK%).Text)
    Next KK%
    '
    On Error GoTo 0
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Form_Load()
    '
    TIARGRA$ = TRIM$(Control$("", "TIARGRA"))
    If TIARGRA$ <> "" Then
       Text1.Text = TIARGRA$
    End If
    '
    RUTARGRA$ = TRIM$(Control$("", "RUTARGRA"))
    If RUTARGRA$ <> "" Then
       Text2.Text = RUTARGRA$
    End If
    '
    ISOLOGO$ = TRIM$(Control$("", "ISOLOGO"))
    If ISOLOGO$ <> "" Then
       Text3.Text = ISOLOGO$
    End If
    '
    For KK% = 0 To 5
       On Error Resume Next
       FORMUNAME(KK%).Text = Control$("FORGRAFI", "DOCUGRA" + TRIM$(Str$(KK%)))
       FORMUECO(KK%).Caption = ECOARCH$("INDIFRM", FORMUNAME(KK%).Text)
    Next KK%
    On Error GoTo 0
    '
End Sub

Private Sub FORMUNAME_DblClick(Index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        FORMUNAME(Index).Text = VDEVU$
        FORMUECO(Index).Caption = VALACT2$("INDIFRM")
    End If
End Sub
