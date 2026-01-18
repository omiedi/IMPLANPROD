VERSION 5.00
Begin VB.Form CONSUDOC09 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Documentos"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   3240
      Left            =   6150
      TabIndex        =   7
      Top             =   990
      Width           =   30
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   45
      TabIndex        =   4
      Top             =   915
      Width           =   8415
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   4590
      Left            =   8670
      ScaleHeight     =   4530
      ScaleWidth      =   1005
      TabIndex        =   0
      Top             =   -15
      Width           =   1065
      Begin VB.CommandButton Command3 
         Caption         =   "Help"
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
         Picture         =   "CONSUDOC09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   885
         Width           =   650
      End
      Begin VB.CommandButton AGenRep 
         Height          =   600
         Left            =   120
         Picture         =   "CONSUDOC09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   2880
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Height          =   630
         Left            =   120
         Picture         =   "CONSUDOC09.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Conectar Documentos"
         Top             =   1995
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   2
         Top             =   3750
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   3
            Top             =   -30
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command7 
         Height          =   660
         Left            =   120
         Picture         =   "CONSUDOC09.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   150
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -240
         Picture         =   "CONSUDOC09.frx":0BA0
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   3990
         Width           =   1515
      End
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1260
      TabIndex        =   12
      Top             =   225
      Width           =   750
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conf. Lec."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6150
      TabIndex        =   9
      Top             =   600
      Width           =   2300
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   45
      TabIndex        =   8
      Top             =   600
      Width           =   6285
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2085
      TabIndex        =   6
      Top             =   225
      Width           =   3570
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   840
   End
End
Attribute VB_Name = "CONSUDOC09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SHELLEXECUTE Lib "SHELL32.DLL" Alias "ShellExecuteA" (ByVal HWND As Long, ByVal IPOPERATION As String, ByVal IPFILE As String, ByVal IPPARAMETERS As String, ByVal IPDIRECTORY As String, ByVal NSHOWCMD As Long) As Long

Private Sub Command4_Click()
  Call SELECHO("FLEXUSR")
    If VALACT1$("FLEXUSR") = "" Then
       Exit Sub
    Else
       Text3 = VALACT1$("FLEXUSR")
       Label6 = VALACT2$("FLEXUSR")
    End If
End Sub

Private Sub Command7_Click()
   End
End Sub


Private Sub Command8_Click()
    CARDOC09.Show 1
End Sub

Private Sub Form_Activate()
   Call CargaDocumentos
End Sub

Private Sub Form_Load()
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   Label1 = USNBR% Mod 100
   
End Sub


Private Sub Label1_Change()
   Call CargaDocumentos
End Sub


Private Sub List2_DblClick()
   If List2.List(List2.ListIndex) <> "" Then
      UBICA$ = Mid(List2.List(List2.ListIndex), 211, 128)
      Call SHELLEXECUTE(Me.HWND, "OPEN", TRIM(UBICA$), "", "", 1)
   End If
   
End Sub

Sub CargaDocumentos()
    Label6 = ECOARCH("FLEXUSR", MKI(XVALO((Label1))))
    List2.Clear
    For I& = 1 To ULTREG("VINCUDOC")
     REGVINCU$ = REGLEIDO$("VINCUDOC", I&)
     CODUSU% = Asc(Mid$(REGVINCU$, 1, 1))
     CODDOCU% = CVI(Mid(REGVINCU$, 2, 2))

      If CODUSU% = XVALO(Label1) Then
         REGDOCU$ = FILTERGETRECORD$("DOCUMEN", 1, MKI$(CODDOCU%))
         DENOM$ = Mid(REGDOCU$, 3, 64)
         UBICA$ = Mid(REGDOCU$, 67, 128)
           
         FILA$ = Space(340)
         Call REPLA(FILA$, DENOM$, 1, 64) 'DENOMINACION
         Call REPLA(FILA$, SAFETEXT(CODDOCU%), 200, 10) 'CODDOCU
         Call REPLA(FILA$, UBICA$, 211, 128)
         List2.AddItem FILA$
      End If
   
    Next I&

End Sub
