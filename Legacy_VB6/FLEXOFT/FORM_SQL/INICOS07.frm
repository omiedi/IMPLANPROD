VERSION 5.00
Begin VB.Form INICOS07 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuestación de Productos"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5130
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   2  'Horizontal Line
      Height          =   3105
      Left            =   0
      Picture         =   "INICOS07.frx":0000
      ScaleHeight     =   3105
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   2725.888
      TabIndex        =   0
      Top             =   0
      Width           =   5370
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "    Análisis de Costos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   3465
         TabIndex        =   1
         Top             =   210
         Width           =   1575
      End
   End
End
Attribute VB_Name = "INICOS07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   '
   'Show
   'Refresh
   VERANTER$ = "ANACOS01"
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "GRISFER") > 0 Then
     If HOY(HOS$) > 9647 Then    ' corresponde 15/02/03
       Call GRACONTROL("ANACOSTO", "LMTBNMNT", "SI")
     End If
   End If
   '
   If TRIM$(CONTROL$("ANACOSTO", "LMTBNMNT")) = "SI" Then
     Call CIERRARCH("*.*")
     Call FINAL("")
   End If
   '
   ANACOSTO07.Show
   Hide
   '
End Sub

