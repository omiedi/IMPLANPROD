VERSION 5.00
Begin VB.Form FORSELMEC 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccion Periodo Mensual"
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BOTREC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      Picture         =   "FORSELMEC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Genera Reporte Solicitado"
      Top             =   420
      Width           =   750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3780
      TabIndex        =   1
      Top             =   420
      Width           =   175
   End
   Begin VB.TextBox MESAN 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1575
      TabIndex        =   0
      Text            =   " "
      Top             =   420
      Width           =   2220
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mes y año:"
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
      Index           =   10
      Left            =   -735
      TabIndex        =   2
      Top             =   525
      Width           =   2220
   End
End
Attribute VB_Name = "FORSELMEC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTREC_Click()
    '
    MMXX0$ = TRIM$(MESAN.TEXT)
    MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
    '
    On Error Resume Next
    FORLICODOS.Label3.Caption = MMXX1$
    FORLICODO07.Label3.Caption = MMXX1$
    On Error GoTo 0
    Hide
    '
End Sub

Private Sub Command3_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
        Exit Sub
    End If
    MESAN.TEXT = MESASEL$
End Sub

