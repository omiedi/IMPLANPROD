VERSION 5.00
Begin VB.Form LITBAJ04 
   BackColor       =   &H00B0C0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Items de Stock Dados de Baja"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   4005
      Left            =   7980
      ScaleHeight     =   3945
      ScaleWidth      =   1005
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
      Begin VB.Image Image3 
         Height          =   4095
         Left            =   0
         Picture         =   "LITBAJ04.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00000080&
      Height          =   4000
      Left            =   7980
      ScaleHeight     =   3945
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
         Picture         =   "LITBAJ04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "LITBAJ04.frx":36FE
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "LITBAJ04.frx":3A08
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   210
         Width           =   640
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "LITBAJ04.frx":3B52
         Top             =   3650
         Width           =   1515
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
      Height          =   3885
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Haga un Doble Click para Re-Vivir el Item"
      Top             =   120
      Width           =   7785
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnubuscaritem 
         Caption         =   "Buscar Item"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "LITBAJ04"
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
  TTXX0$ = REGBLAN("ECOMAST")
  For I& = 0 To List1.ListCount - 1
    TTXX$ = TTXX0$
    Call REPLA(TTXX$, Mid(List1.List(I), 1, 16), 1, 16)
    Call REPLA(TTXX$, Mid(List1.List(I), 19, 46), 17, 46)
    CI% = XVALO(Mid(List1.List(I), 73, 8))
    Call REPLA(TTXX$, MKI$(CI%), 63, 2)
    Call GRAREG("!ECOMAST", TTXX$, I& + 1)
  Next I&
  
  Call SETULTREG("!ECOMAST", List1.ListCount)
  Call SELECHO("!ECOMAST")
  COD$ = VALACT1$("!ECOMAST")
  For I& = 0 To List1.ListCount - 1
    If TRIM(Mid(List1.List(I), 1, 16)) = COD$ Then
        List1.ListIndex = I
    End If
  Next I&
  Command2.Enabled = True
  '\\\OT-06-0523-RG-FIN///
End Sub

Private Sub Form_Load()

  If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
      Picture2.Visible = True
  End If
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
   CDXX$ = TRIM$(Left$(List1.TEXT, 15))
   If CDXX$ = "" Then GoTo 99
   '
   If INCLUALIAS% = 1 Then
     Call APERBASDAT("MASTER")
     SQLX$ = "NPlano = '" + CDXX$ + "' "
     Master.FindFirst SQLX$
     If Master.NoMatch = False Then
       Call MENSERR(24, "Imposible Restaurar Código '" + CDXX$ + "'\por ser Alias de '" + Master!CODIGO + "'.")
       GoTo 99
     End If
   End If
   '
   CII% = CODINT%(CDXX$)
   ITESEL& = List1.ListIndex
   If Abs(CII%) > 0 Then
     CI1% = Abs(CII%)
     Master.FindFirst ("CODINT = " & CI1%)
     If Master.NoMatch = False Then
       Master.Edit
       Master!Status = 1
       If Mid$(Master!CODIGO, 16, 1) = Chr$(96) Then
         Master!CODIGO = CDXX$
       End If
       Master.Update
       UPDATMASTER% = 1
       On Error Resume Next
       List1.RemoveItem ITESEL&
       List1.Refresh
       Call COMUNI("Código '" + CDXX$ + "' Restaurado.")
     End If
   End If
   '
99 Screen.MousePointer = 1
   '
End Sub

Private Sub mnubuscaritem_Click()
    Call Command2_Click
End Sub

Private Sub mnusalir_Click()
    Call Command1_Click
End Sub
