VERSION 5.00
Begin VB.Form INIGSMA14 
   Caption         =   "Sistema para Estudio Contable"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   0
      Picture         =   "INIGSMA14.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6135
   End
   Begin VB.Menu mnuabm 
      Caption         =   "A.B.M"
      Begin VB.Menu MnuABMClientes 
         Caption         =   "A.B.M. de Clientes"
      End
      Begin VB.Menu mnuimp 
         Caption         =   "A.B.M. de Impuestos"
      End
   End
   Begin VB.Menu mnuHon 
      Caption         =   "Honorarios"
      Begin VB.Menu mnuadminhon 
         Caption         =   "Administrar Honorarios"
      End
   End
   Begin VB.Menu mnui 
      Caption         =   "Impuestos"
      Begin VB.Menu mnuimpuestos 
         Caption         =   "Administrar Impuestos"
      End
   End
   Begin VB.Menu mnusetup 
      Caption         =   "Setup"
      Begin VB.Menu Mnuconfig 
         Caption         =   "Configuración"
      End
   End
   Begin VB.Menu mnucc1 
      Caption         =   "C.Contables"
      Begin VB.Menu mnucc 
         Caption         =   "C.Contables      "
      End
   End
   Begin VB.Menu mnudis 
      Caption         =   "Diseño        "
      Begin VB.Menu mnudiseñofrm 
         Caption         =   "Diseñador de Formularios     "
      End
      Begin VB.Menu mnugrep 
         Caption         =   "Generador de Reportes"
      End
   End
End
Attribute VB_Name = "INIGSMA14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const LocalMachine As String = "HKEY_LOCAL_MACHINE"
Private Const CurrentUser As String = "HKEY_CURRENT_USER"


Private Sub Form_Load()
   '
   If App.PrevInstance = True Then
      Close: End
   End If
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '

15  EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
'    Call APLICAR_SKIN(Me, LUDAT$ & "\SKINS\dogmax.skn")
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    Call ABRECONEXION
    Call ABRECLIEN
    Call ABREMASTER
    Call ABRECABAN
    Call CREACAMPO("", "ITEMS", "CODINT", 3, 0, 1)
    Call CREACAMPO("", "ITEMS", "DESCRIPCION", 10, 128)
    Call CREACAMPO("", "ITEMS", "STATUS", 3, 0, 1)
    Call CREACAMPO("", "ITEMS", "MARBORRA", 3, 0)
    '
    Call ABMIMP14.ACTUITEMDAT
    If ULTREG&("CAPOSIV") < 8 Then
      IQ& = 0
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(0) + AJUSTI$("IVA Consumidor Final            ", 31) + AJUSTI$("Cons.Fin", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(1) + AJUSTI$("IVA Responsable Inscripto   ", 31) + AJUSTI$("Resp.Ins", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(2) + AJUSTI$("IVA Responsable no Inscripto", 31) + AJUSTI$("R.No Ins", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(3) + AJUSTI$("IVA No Responsable          ", 31) + AJUSTI$("No Resp.", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(4) + AJUSTI$("IVA Responsable Exento      ", 31) + AJUSTI$("Exento  ", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(5) + AJUSTI$("IVA Operac.Comercio Exterior", 31) + AJUSTI$("Com.Ext. ", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(6) + AJUSTI$("IVA Contrib.Monotributo     ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(7) + AJUSTI$("IVA Exento Zona Promovida   ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
      Call SETULTREG("CAPOSIV", IQ&)
      Call CIERRARCH("CAPOSIV")
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Close
   End

End Sub



Private Sub MnuABMClientes_Click()
   ABMCLI14.Show 1
End Sub
    
Private Sub mnuadminhon_Click()
   Call CONECRUN("HONORA14", "Sistema de Honoarios")
End Sub

Private Sub mnucc_Click()
   If CLACONTRA% > 0 Then
     CONFICTA14.Show 1
   End If
End Sub

Private Sub Mnuconfig_Click()
   Call CONECRUN("FSETUP04", "Configuración de Funcionamiento")
   Call FINAL("")
End Sub

Private Sub mnudiseñofrm_Click()
   Call CONECRUN("DSGFRM09", "Diseñador de Formularios")
End Sub

Private Sub mnugrep_Click()
   Call CONECRUN("GENREP07", "Generador de Reportes SQL")
   
End Sub

Private Sub mnuimp_Click()
   ABMIMP14.Show 1
End Sub

Private Sub mnuimpuestos_Click()
   Call CONECRUN("IMPUES14", "Administración de Impuestos")

End Sub

