VERSION 5.00
Begin VB.Form INIBAP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buenos Aires Players"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   2600
      Left            =   0
      Picture         =   "INIBAP.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4500
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu MnuCerrar 
         Caption         =   "Cerrar"
      End
   End
   Begin VB.Menu mnuReporte 
      Caption         =   "Reportes"
      Begin VB.Menu mnuReportes 
         Caption         =   "Reportes"
      End
   End
End
Attribute VB_Name = "INIBAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   '
'   Call VERJOBID(OKEY%)
'   If OKEY% < 1 Then Call FINAL("")
'   '
'   '
'   EPAC$ = TRIM$(UCase$(EMPREAC$))
   '
'   UTILIZA_SKIN% = 1
'   AMASTOK07.Show
   '
    Call iniciar_variables
    Call ABRECONEXION_PUNTODEVENTA
'    Call APLICAR_SKIN(Me, LUDAT$ & "\SKINS\dogmax.skn")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Cancel = 0 Then Call FINAL("")
End Sub

Private Sub MnuCerrar_Click()
  Call FINAL("")
End Sub

Private Sub mnuReporte1_Click()
End Sub


Private Sub mnuReportes_Click()
    MUESTRABAP.Show 1
End Sub

