VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form LITBAJ07 
   BackColor       =   &H00B9DFEB&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Items de Stock Dados de Baja"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2760
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00404080&
      Height          =   7575
      Left            =   7980
      ScaleHeight     =   7515
      ScaleWidth      =   1005
      TabIndex        =   1
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command2 
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
         Left            =   105
         Picture         =   "LITBAJ07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Buscar Item "
         Top             =   1575
         Width           =   640
      End
      Begin VB.CommandButton Command11 
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
         Left            =   105
         Picture         =   "LITBAJ07.frx":1366
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   860
         Width           =   640
      End
      Begin VB.CommandButton command1 
         Caption         =   "Exit"
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
         Left            =   105
         Picture         =   "LITBAJ07.frx":1670
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   210
         Width           =   640
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   2
         Top             =   6090
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   3
            Top             =   105
            Width           =   12000
         End
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Left            =   105
      TabIndex        =   0
      ToolTipText     =   "Haga un Doble Click para Re-Vivir el Item"
      Top             =   210
      Width           =   7785
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "LITBAJ07.frx":17BA
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "LITBAJ07.frx":19EE
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "LITBAJ07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Unload Me
End Sub

Private Sub Command2_Click()
      '\\\OT-06-0523-RG-01/12/06///
      Command2.Enabled = False
    '''  TTXX0$ = REGBLAN("ECOMAST")
    '''  For I& = 0 To List1.ListCount - 1
    '''    TTXX$ = TTXX0$
    '''    Call REPLA(TTXX$, Mid(List1.List(I), 1, 16), 1, 16)
    '''    Call REPLA(TTXX$, Mid(List1.List(I), 19, 46), 17, 46)
    '''    CI% = XVALO(Mid(List1.List(I), 73, 8))
    '''    Call REPLA(TTXX$, MKI$(CI%), 63, 2)
    '''    Call GRAREG("!ECOMAST", TTXX$, I& + 1)
    '''  Next I&
        
    '''  Call SETULTREG("!ECOMAST", List1.ListCount)
    '''  Call FRESHECHO("!ECOMAST") 'AGREGADO POR QUE NO SE REFRESCABA INMEDIATAMNTE
    
      Call SELECHO("ECOMABAJ")
      COD$ = TRIM$(VALACT1$("ECOMABAJ"))
      If COD$ = "" Then GoTo 99
      For i& = 0 To List1.ListCount - 1
        If TRIM(Mid(List1.List(i), 1, 16)) = COD$ Then
            List1.ListIndex = i
        End If
      Next i&
      '
99    Command2.Enabled = True
  '\\\OT-06-0523-RG-FIN///
End Sub
Private Sub Form_Load()
   'Call APLICACIONSKINS(Me, Picture6)
End Sub

Private Sub List1_DblClick()
   '
   If DERACCE%("BAJAITEM", "Baja y Restauración de un Item de Stock") < 2 Then
     Exit Sub
   End If
   '
   If INCLUALIAS% = 0 Then
     INCLUALIAS% = (-1)
     If CONTROL$("", "IDSECUN") = "ALIAS" Then INCLUALIAS% = 1
   End If
   '
   Screen.MousePointer = 11
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Screen.MousePointer = 11
   LONCODI% = XVALO(CONTROL$("MASTER", "LOMAXCOD"))
   If LONCODI% < 15 Or LONCODI% > 24 Then LONCODI% = 15
   '
   CDXX$ = TRIM$(Left$(List1.TEXT, LONCODI%))
   If CDXX$ = "" Then GoTo 99
   '
   If INCLUALIAS% = 1 Then
      Call APERBASDAT("MASTER")
      Dim myrs As ADODB.Recordset
      strsql = ("Select codigo,NPlano from Master where NPlano='" & CDXX$ & "'")
      Set myrs = FLEXCONN.Execute(FILTSQL$(strsql))
      If Not (myrs.EOF And myrs.BOF) Then
         Call MENSERR(24, "Imposible Restaurar Código '" + CDXX$ + "'\por ser Alias de '" + myrs!Codigo + "'.")
         GoTo 99
      End If
   End If
   '
   CII% = Abs(CODINT%(CDXX$))
   ITESEL& = List1.ListIndex
   If Abs(CII%) > 0 Then
      CI1% = Abs(CII%)
      strsql = ("Update Master set status=1 where codint=" & CI1%)
       FLEXCONN.Execute (FILTSQL$(strsql))
      UPDATMASTER% = 1
      On Error Resume Next
      List1.RemoveItem ITESEL&
      List1.Refresh
      Call GENEECOMABAJ 'ACTUALIZA LISTA DE SELECCION
      Call COMUNI("Código '" + CDXX$ + "' Restaurado.")
   End If
   '
99 Screen.MousePointer = 1
   '
   On Error Resume Next
   myrs.Close
   Set myrs = Nothing
   On Error GoTo 0
   '
End Sub
