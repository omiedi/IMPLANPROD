VERSION 5.00
Begin VB.Form FORSELTAL 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Talles"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4545
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   3805
      ScaleHeight     =   75
      ScaleWidth      =   585
      TabIndex        =   4
      Top             =   2940
      Width           =   650
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   12000
      End
   End
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
      Height          =   650
      Left            =   3800
      Picture         =   "FORSELTAL.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Aprueba la Selección"
      Top             =   840
      Width           =   650
   End
   Begin VB.CommandButton BOTEXIT 
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
      Left            =   3800
      Picture         =   "FORSELTAL.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancela la Selección"
      Top             =   105
      Width           =   650
   End
   Begin VB.ListBox LITALSEL 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   105
      Width           =   3585
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   2415
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2205
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   3370
      Picture         =   "FORSELTAL.frx":0614
      Top             =   3255
      Width           =   1515
   End
End
Attribute VB_Name = "FORSELTAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTEXIT_Click()
    LITALSEL.Clear
    Unload Me
End Sub
'

Private Sub BOTREC_Click()
    '
    For U& = 1 To LITALSEL.ListCount
      If LITALSEL.Selected(U& - 1) = True Then
        COD$ = TRIM$(AMASTOCK.Text1.TEXT) & "/" & TRIM$(Mid$(LITALSEL.List(U& - 1), 24, 2))
        DESC$ = TRIM$(AMASTOCK.Text3) & " " & TRIM$(Mid$(LITALSEL.List(U& - 1), 1, 23))
        TM% = Int(AMASTOCK.Text6)
        UM$ = TRIM$(AMASTOCK.Text4)
        If Len(COD$) <= 15 Then
          If CODINT%(COD$) < 0 Then
            Call AGRECOD(COD$, DESC$, TM%, UM$, OKX%)
            If OKX% < 1 Then
              Call MENSERR(24, "Error al Agregar Codigo " + COD$)
            End If
          Else
              Call MENSERR(24, "Codigo " + COD$ & " " & "Existente o de Baja")
          End If
        End If
      End If
    Next U&
    '
End Sub

Private Sub Form_Load()
    '
    LITALSEL.Clear
    For U& = 1 To ULTREG&("TABLATA")
      X$ = REGLEIDO$("TABLATA", U&)
      XX$ = TRIM$(Mid$(X$, 15, 16)): AA% = Len(XX$)
      LITALSEL.AddItem XX$ & Space(23 - AA%) & Mid$(X$, 13, 2)
    Next U&
    '
    Call CIERRARCH("TABLATA")
    '
End Sub
