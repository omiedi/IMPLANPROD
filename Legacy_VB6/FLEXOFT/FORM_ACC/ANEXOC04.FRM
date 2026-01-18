VERSION 5.00
Begin VB.Form ANEXOC04 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Emisión de O/C's - Documentos Anexos "
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5850
   ControlBox      =   0   'False
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
   ScaleHeight     =   2220
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Emitir / enviar por"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   210
      TabIndex        =   8
      Top             =   105
      Width           =   4740
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "E-Mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2625
         TabIndex        =   10
         Top             =   420
         Width           =   1800
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   525
         TabIndex        =   9
         Top             =   420
         Value           =   -1  'True
         Width           =   1800
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Documentos a Anexar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   210
      TabIndex        =   2
      Top             =   1050
      Width           =   4740
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Fotografía - 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   2625
         TabIndex        =   6
         Top             =   630
         Width           =   1905
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Fotografía - 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   2625
         TabIndex        =   5
         Top             =   315
         Width           =   1905
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Croquis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   525
         TabIndex        =   4
         Top             =   630
         Width           =   1905
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Plano de Pieza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   525
         TabIndex        =   3
         Top             =   315
         Width           =   1905
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   5040
      Picture         =   "ANEXOC04.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   210
      Width           =   645
   End
   Begin VB.CommandButton command1 
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
      Height          =   920
      Left            =   5040
      Picture         =   "ANEXOC04.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1155
      Width           =   645
   End
   Begin VB.Label CONTROLABEL 
      Height          =   225
      Left            =   5775
      TabIndex        =   7
      Top             =   105
      Visible         =   0   'False
      Width           =   645
   End
End
Attribute VB_Name = "ANEXOC04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
   CONTROLABEL.Caption = "OK"
   Hide
End Sub

Private Sub Command7_Click()
   CONTROLABEL.Caption = ""
   Hide
End Sub

Private Sub Form_Load()
    '
    For KKI% = 0 To 3
      Check1(KKI%).Value = 0
      If CONTROL$("PRESUPUE", "PRIGRA-" + TRIM$(Str$(KKI%))) = "SI" Then
        Check1(KKI%).Value = 1
      End If
    Next KKI%
    '
End Sub

