VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FRMPEDPE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos Pendientes"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9915
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   7
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   3900
      Left            =   9030
      ScaleHeight     =   3840
      ScaleWidth      =   1845
      TabIndex        =   1
      Top             =   0
      Width           =   1905
      Begin VB.CommandButton Command2 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Left            =   100
         Picture         =   "FRMPEDPE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Búsqueda por  Texto"
         Top             =   840
         Width           =   680
      End
      Begin VB.CommandButton BOTEXIT 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Left            =   100
         Picture         =   "FRMPEDPE.frx":1366
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cancelar la Aplicación"
         Top             =   105
         Width           =   680
      End
      Begin VB.CommandButton BOTREC 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Left            =   100
         Picture         =   "FRMPEDPE.frx":1670
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cargar Mail"
         Top             =   1600
         Width           =   680
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   140
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   2
         Top             =   2310
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "FRMPEDPE.frx":197A
         Top             =   2520
         Width           =   1515
      End
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
      Height          =   2610
      Left            =   90
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   210
      Width           =   8835
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FRMPEDPE.frx":389C
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "FRMPEDPE.frx":3AD0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "FRMPEDPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODORDE%
Dim MUESTRACAR%

Private Sub BOTEXIT_Click()
    Hide
End Sub
'
Private Sub BOTREC_Click()
    'RECORRO LOS MAIL Y CONCATENO
    'Se concatena a manera que quede del modo del ejemplo siguiente
    'omardemaio@hotmail.com>, <fgrasso@zetti.com.ar>, <lmuraca@zetti.com.ar
   'Nota el simbolo '<' de inicio y el de cierre lo pone solo el outlook
    MAIX$ = ""
    VUELTA% = 0
    SIMBO$ = ""
    For I& = 1 To LICHESEL.ListCount
       If LICHESEL.Selected(I& - 1) = True Then
'       If VUELTA% > 0 Then SIMBO$ = "<"
'        MAIX$ = MAIX$ + SIMBO$ + TRIM$(Mid$(LICHESEL.List(I& - 1), 44)) + ">, "
'        VUELTA% = 1
        MAIX$ = MAIX$ + ";" + TRIM$(Mid$(LICHESEL.List(I& - 1), 44))
       End If
    Next I&
    MAIX$ = TRIM$(MAIX$)
    If InStr(MAIX$, ">,") > 0 Then MAIX$ = Mid$(MAIX$, 1, Len(MAIX$) - 2)
    '
    'ASIGNA EL VALOR A LA VARIABLE
    SENDBYMAIL$ = MAIX$
    '
    Call BOTEXIT_Click
End Sub
'
Private Sub Command2_Click()
   Static ABUSQUE$
   ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", "Introduzca Argumento de Búsqueda", ABUSQUE$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     INI& = LICHESEL.ListIndex
10   For U& = INI& + 2 To LICHESEL.ListCount
       If InStr(UCase$(LICHESEL.List(U& - 1)), ABUSQUE$) > 0 Then
         LICHESEL.ListIndex = U& - 1
         Exit Sub
       End If
     Next U&
     If INI& >= 0 Then
       INI& = -1
       GoTo 10
     End If
     '
     Call MENSERR(24, "Argumento no Encontrado.")
   End If
End Sub

Private Sub Form_Load()
   Call APLICACIONSKINS(Me)
End Sub
