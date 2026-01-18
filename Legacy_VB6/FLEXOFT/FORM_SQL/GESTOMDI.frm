VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.MDIForm GESTOMDI 
   AutoShowChildren=   0   'False
   BackColor       =   &H00404040&
   Caption         =   "Administración de Alquileres"
   ClientHeight    =   7635
   ClientLeft      =   -105
   ClientTop       =   660
   ClientWidth     =   8850
   Icon            =   "GESTOMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture999 
      Align           =   1  'Align Top
      Height          =   375
      Index           =   0
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   8790
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8850
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
         Height          =   225
         Index           =   0
         Left            =   0
         OleObjectBlob   =   "GESTOMDI.frx":058A
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   1485
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "GESTOMDI.frx":05E8
      Top             =   360
   End
   Begin VB.Menu MnuABM 
      Caption         =   "A.B.M."
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
      Begin VB.Menu mnucc 
         Caption         =   "C.Ctables"
      End
   End
   Begin VB.Menu mnudis 
      Caption         =   "Diseño"
      Begin VB.Menu mnudiseñofrm 
         Caption         =   "Diseñador de formularios"
      End
      Begin VB.Menu mnugrep 
         Caption         =   "Generador de Reportes"
      End
   End
End
Attribute VB_Name = "GESTOMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    Call CENTRAFORM(Me)
'    Show
    '
MsgBox "1"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    
 MsgBox "1.4"   '
15  EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
MsgBox "2"
'    Call APLICAR_SKIN(Me, LUDAT$ & "\SKINS\dogmax.skn")
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
MsgBox "3"        '
    Call ABRECONEXION
    Call ABRECLIEN
    Call ABREMASTER
    Call ABRECABAN
MsgBox "4"
    Call CREACAMPO("", "ITEMS", "CODINT", 3, 0, 1)
    Call CREACAMPO("", "ITEMS", "DESCRIPCION", 10, 128)
    Call CREACAMPO("", "ITEMS", "STATUS", 3, 0, 1)
    Call CREACAMPO("", "ITEMS", "MARBORRA", 3, 0)
    '
MsgBox "5"
    Call ABMIMP14.ACTUITEMDAT
MsgBox "6"    '
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
MsgBox "7"
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Call FINAL("")

End Sub

Private Sub MnuABMClientes_Click()
   ABMCLI14.Show 1
End Sub

Private Sub mnuadminhon_Click()
   Call CONECRUN("HONORA14", "Sistema de Honoarios")
End Sub

Private Sub mnucc_Click()
   CONFICTA14.Show 1
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
