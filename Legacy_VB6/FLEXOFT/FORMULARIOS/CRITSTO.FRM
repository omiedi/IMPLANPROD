VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form CRITSTO 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Critico - Items Sub-Mínimo"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   210
      TabIndex        =   3
      Top             =   315
      Width           =   3690
      Begin VB.OptionButton Option3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Revisar Stocks Actuales"
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
         Left            =   420
         TabIndex        =   6
         Top             =   1260
         Width           =   2955
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Listado por Impresora"
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
         Left            =   420
         TabIndex        =   5
         Top             =   840
         Width           =   2745
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Consulta por Pantalla"
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
         Left            =   420
         TabIndex        =   4
         Top             =   420
         Width           =   2430
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
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
      Left            =   4620
      Picture         =   "CRITSTO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
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
      Left            =   4620
      Picture         =   "CRITSTO.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   315
      Width           =   855
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   210
      TabIndex        =   0
      Top             =   2625
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   397
      _Version        =   327680
      Appearance      =   1
      MouseIcon       =   "CRITSTO.frx":0454
   End
End
Attribute VB_Name = "CRITSTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    CRITSTO.Hide
End Sub

Private Sub Command2_Click()
    If CRITSTO.Option1.Value = True Then
        Call CONECRUN("*CRITSTS", "")
      ElseIf CRITSTO.Option2.Value = True Then
        Call CONECRUN("*CRITSTL", "")
      Else
        Call CRITSTO2
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Screen.MousePointer = 1
    Option3.Value = True
End Sub

