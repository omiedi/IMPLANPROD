VERSION 5.00
Begin VB.Form FORSELCHE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccion de Cheques a Depositar"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8310
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   7245
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   9
      Top             =   2835
      Width           =   855
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000004&
      Height          =   540
      Left            =   105
      ScaleHeight     =   480
      ScaleWidth      =   6885
      TabIndex        =   4
      Top             =   3220
      Width           =   6945
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4790
         TabIndex        =   8
         Top             =   90
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Valores:"
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
         Left            =   3360
         TabIndex        =   7
         Top             =   150
         Width           =   1380
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1470
         TabIndex        =   6
         Top             =   90
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Seleccionados:"
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
         Left            =   105
         TabIndex        =   5
         Top             =   150
         Width           =   1695
      End
   End
   Begin VB.CommandButton BOTREC 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   7245
      Picture         =   "FORSELCHE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Grabar - Guardar Modificaciones"
      Top             =   1050
      Width           =   855
   End
   Begin VB.CommandButton BOTEXIT 
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
      Height          =   825
      Left            =   7245
      Picture         =   "FORSELCHE.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancelar la Aplicación"
      Top             =   105
      Width           =   855
   End
   Begin VB.ListBox LICHESEL 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3120
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   105
      Width           =   6945
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   3045
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   2205
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   6720
      Picture         =   "FORSELCHE.frx":0614
      Top             =   3250
      Width           =   2010
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   8500
      Y1              =   3765
      Y2              =   3765
   End
End
Attribute VB_Name = "FORSELCHE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTEXIT_Click()
    LICHESEL.Clear
    Label2.Caption = ""
    Hide
End Sub
'

Private Sub BOTREC_Click()
    SUANT# = Val(Label2)
    Call SUCHESEL
    If Val(Label2) = SUANT# Then
        Hide
      Else
        Call MENSERR(24, "Revise Valores Seleccionados")
    End If
End Sub

Private Sub Form_Activate()
    Call SUCHESEL
    LICHESEL.SetFocus
End Sub

Sub SUCHESEL()
    SUCHESE# = 0
    For I& = 1 To LICHESEL.ListCount
       If LICHESEL.Selected(I& - 1) = True Then
         SUCHESE# = SUCHESE# + Val(Mid$(LICHESEL.List(I& - 1), 41, 12))
       End If
    Next I&
    Label2.Caption = CSTRING$(MKD$(SUCHESE#), 4, 11, 2, 2)
    Label2.Refresh
End Sub

Private Sub LICHESEL_Click()
   Call SUCHESEL
End Sub




