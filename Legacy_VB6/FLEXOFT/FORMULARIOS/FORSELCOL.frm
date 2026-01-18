VERSION 5.00
Begin VB.Form FORSELCOL 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4470
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   4575
      Left            =   3570
      ScaleHeight     =   4515
      ScaleWidth      =   1035
      TabIndex        =   2
      Top             =   0
      Width           =   1095
      Begin VB.CommandButton Command1 
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
         Left            =   120
         Picture         =   "FORSELCOL.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   780
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   4
         Top             =   2310
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
         Left            =   120
         Picture         =   "FORSELCOL.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Aprueba la Pre-Selección"
         Top             =   1440
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
         Left            =   120
         Picture         =   "FORSELCOL.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "FORSELCOL.frx":075E
         Top             =   2550
         Width           =   1515
      End
   End
   Begin VB.ListBox LICOLSEL 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2865
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   105
      Width           =   3375
   End
End
Attribute VB_Name = "FORSELCOL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTEXIT_Click()
   Unload Me
End Sub

Private Sub BOTREC_Click()
   '
   BOTREC.Enabled = False
   Screen.MousePointer = 11
   CODIBA$ = AMASTO00.Text1
   CIBA% = CODINT%(CODIBA$)
   DECRIBA$ = TRIM$(AMASTO00.Text3)
   TIMA% = TIMATE%(CIBA%)
   UNID$ = UNIMED$(CIBA%)
   '
   For U& = 1 To LICOLSEL.ListCount
     If LICOLSEL.Selected(U& - 1) = True Then
       TTXX$ = LICOLSEL.List(U& - 1)
       CODICOLO$ = TRIM$(Left$(TTXX$, 2))
       DECRICOLO$ = TRIM$(Mid$(TTXX$, 6))
       '
       CODINUE$ = CODIBA$ + "/" + CODICOLO$
       If CODINT%(CODINUE$) < 1 Then
         DESCRINUE$ = DECRIBA$ + " " + DECRICOLO$
         Call AGRECOD(CODINUE$, DESCRINUE$, TIMA%, UNID$, OKX%)
       End If
     End If
   Next U&
   '
   Screen.MousePointer = 1
   BOTREC.Enabled = True
   '
   Call COMUNI("Se Ha Completado la Operación")
   Unload Me
   '
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   '
   TPI = LICOLSEL.TopIndex
   For U& = 1 To LICOLSEL.ListCount
     If LICOLSEL.Selected(U& - 1) = True Then
         LICOLSEL.Selected(U& - 1) = False
       Else
         LICOLSEL.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LICOLSEL.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   '
   Command1.Enabled = True
End Sub

Private Sub Form_Load()
    '
    FORSELCOL.Caption = TRIM$(AMASTO00.Text1) + " - " + TRIM$(AMASTO00.Text3)
    LICOLSEL.Clear
    '
    For U& = 1 To ULTREG&("TACOLOR")
      X$ = REGLEIDO$("TACOLOR", U&)
      CODCOLOR$ = TRIM$(Mid$(X$, 13, 2))
      If Len(CODCOLOR$) < 2 Then CODCOLOR$ = Space(1) + CODCOLOR$
      LICOLSEL.AddItem CODCOLOR$ + " - " + TRIM$(Mid$(X$, 15, 16))
    Next U&
    '
End Sub
