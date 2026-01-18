VERSION 5.00
Begin VB.Form OPCALCOS 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Criterio de Calculo"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5130
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
   ScaleHeight     =   2250
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cálculo Costos de Reposición"
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
      Left            =   210
      TabIndex        =   2
      Top             =   315
      Width           =   3690
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Pesificado correspondiente al momento del cálculo"
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
         Left            =   525
         TabIndex        =   4
         Top             =   1155
         Width           =   2745
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Dolarizado y luego ajustado al período mensual a calcular"
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
         Left            =   525
         TabIndex        =   3
         Top             =   525
         Value           =   -1  'True
         Width           =   2850
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
      Height          =   800
      Left            =   4095
      Picture         =   "OPCALCOS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1315
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
      Height          =   800
      Left            =   4095
      Picture         =   "OPCALCOS.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   420
      Width           =   855
   End
End
Attribute VB_Name = "OPCALCOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    PORCALCOS$ = "DOLARIZ"
    If Option2.Value = True Then PORCALCOS$ = "PESIFIC"
    Call GRACONTROL("IGESVEN", "MOCALCOS", PORCALCOS$)
    Hide
    '
End Sub

Private Sub Form_Load()
    '
    PORCALCOS$ = TRIM$(CONTROL$("IGESVEN", "MOCALCOS"))
    If PORCALCOS$ = "PESIFIC" Then
       Option2.Value = True
    End If
    '
End Sub

