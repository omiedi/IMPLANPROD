VERSION 5.00
Begin VB.Form CONPRIRE 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Impresión"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command14 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   875
      Left            =   6160
      Picture         =   "CONPRIRE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Aprueba, Graba e Imprime Presupuesto de Reparación"
      Top             =   300
      Width           =   540
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00CDDADA&
      Caption         =   "IMPRESION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   5955
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Replica por Mail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   4680
         TabIndex        =   6
         Top             =   315
         Width           =   1050
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Impresión Fantasma"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   3360
         TabIndex        =   3
         Top             =   315
         Width           =   1170
      End
      Begin VB.CheckBox Check8 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Presupuesto Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   1800
         TabIndex        =   2
         Top             =   315
         Value           =   1  'Checked
         Width           =   1560
      End
      Begin VB.CheckBox Check7 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Presupuesto Interno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   240
         TabIndex        =   1
         Top             =   315
         Value           =   1  'Checked
         Width           =   1620
      End
   End
   Begin VB.Label Label1 
      Height          =   435
      Left            =   105
      TabIndex        =   4
      Top             =   1890
      Width           =   1380
   End
End
Attribute VB_Name = "CONPRIRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command14_Click()
    Label1.Caption = "OK"
    Hide
End Sub

Private Sub Form_Load()
    '
'    Check7.Value = 1
'    If CONTROL$("PEDIDO", "PRINOVEN") = "NO" Then Check7.Value = 0
'    Check8.Value = 0
'    If CONTROL$("PEDIDO", "PRINODES") = "SI" Then Check8.Value = 1
'    Check2.Value = 0
'    If CONTROL$("PEDCLIEN", "SUPRINT") = "SI" Then Check2.Value = 1
    '
End Sub
