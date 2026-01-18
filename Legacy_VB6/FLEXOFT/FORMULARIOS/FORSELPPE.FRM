VERSION 5.00
Begin VB.Form FORSELPPE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8280
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   4575
      Left            =   7350
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
         Picture         =   "FORSELPPE.frx":0000
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
         Picture         =   "FORSELPPE.frx":014A
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
         Picture         =   "FORSELPPE.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "FORSELPPE.frx":075E
         Top             =   2550
         Width           =   1515
      End
   End
   Begin VB.ListBox LIPEDSEL 
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
      Width           =   7155
   End
End
Attribute VB_Name = "FORSELPPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTEXIT_Click()
   OFABRIC.Label1 = ""
   Unload Me
End Sub

Private Sub BOTREC_Click()
   '
   OFABRIC.Label1 = ""
   OFABRIC.Label12 = ""
   For U% = 0 To (LIPEDSEL.ListCount - 1)
     If LIPEDSEL.Selected(U%) = True Then
       CODIGO$ = Mid$(LIPEDSEL.List(U%), 1, 16)
       NUMCOD% = CODINT%(CODIGO$)
       SALPEND! = Val(Mid$(LIPEDSEL.List(U%), 46, 9))
       ' Guarda en label1 los 2 bytes del codigo seleccionado y
       ' los 4 de la cantidad pendiente
       OFABRIC.Label1 = OFABRIC.Label1 & MKI$(NUMCOD%) & MKS$(SALPEND!)
     End If
   Next U%
   '
   Unload Me
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   '
   TPI = LIPEDSEL.TopIndex
   For U& = 1 To LIPEDSEL.ListCount
     If LIPEDSEL.Selected(U& - 1) = True Then
         LIPEDSEL.Selected(U& - 1) = False
       Else
         LIPEDSEL.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LIPEDSEL.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   '
   Command1.Enabled = True
End Sub

Private Sub Form_Load()
    '
    TITU$ = "Lista de Articulos del Pedido Nº " & OFABRIC.Label1.Caption
    FORSELPPE.Caption = TITU$
    LIPEDSEL.Clear
    '
    NROPED& = Val(OFABRIC.Label1)
    RFF$ = RECFILT$("PEDDETA", 1, MKS$(NROPED&))
    For U& = 1 To Len(RFF$) Step 4
      UU& = CVS(Mid$(RFF$, U&, 4))
      X$ = REGLEIDO$("PEDDETA", UU&)
      CI% = CVI((Mid$(X$, 33, 2)))
      If CI% > 0 Then
        CODIGO$ = CODEXT$(CI%)
        DESCRIP$ = Mid$(ECOARCH$("MASTER", MKI$(CI%)), 1, 28)
        II% = 28 - Len(TRIM$(DESCRIP$)) ' Ajusta la descripcion a 28 caracteres
        SALPEND! = CVS(Mid$(X$, 99, 4))
        SALPENDP$ = FORMATNUM$(SALPEND!, "F9.1")
        LIPEDSEL.AddItem CODIGO$ + " " + DESCRIP$ + Space(II%) + " " + SALPENDP$
       End If
    Next U&
    '
    OFABRIC.Label1 = ""
End Sub
