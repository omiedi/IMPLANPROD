VERSION 5.00
Begin VB.Form SHOWGRAF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   645
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   900
      TabIndex        =   0
      ToolTipText     =   "Haga Doble-Click para Imprimir Documento"
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "SHOWGRAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub Form_Unload(Cancel As Integer)
   HORIMAGEN = Left
   VERTIMAGEN = Top
End Sub

Private Sub Picture2_DblClick()
   On Error Resume Next
     CIXI% = CODINT%(AMASTOCK.Text1.TEXT)
   If Err > 0 Then
     CIXI% = CODINT%(FORPLAN04.Text1.TEXT)
   End If
   On Error GoTo 0
   If CIXI% > 0 Then
     If CIXI% <= NUMAS% Then
       IGRA% = INDIMAGEN%
       HOII% = HOY(HOS$)
       DOCUORIG$ = "AMASTOCK-" + HOS$ + "-" + Left$(Time$, 5)
       Call PRIDOGRAF(CIXI%, IGRA%)
     End If
   End If
End Sub
