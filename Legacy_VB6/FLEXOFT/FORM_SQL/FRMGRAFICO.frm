VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form FRMGRAFICO 
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15075
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   15075
   StartUpPosition =   3  'Windows Default
   Begin MSChart20Lib.MSChart MC 
      Height          =   7935
      Left            =   0
      OleObjectBlob   =   "FRMGRAFICO.frx":0000
      TabIndex        =   1
      Top             =   0
      Width           =   15015
   End
End
Attribute VB_Name = "FRMGRAFICO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CANTIFILASGRID As Integer
Public CAMPOLABEL As Integer
Public SERIE1 As Integer
Public SERIE2 As Integer
Public GRID As MSFlexGrid


Private Sub Form_Load()
SERIE% = 1
If SERIE2 > 0 Then SERIE% = 2
For i& = 1 To CANTIFILASGRID
    With MC.DataGrid
        .SetSize i&, SERIE%, i&, SERIE% '( 2:CANTIDAD DE SERIES)
        .RowLabel(i&, 1) = Left$(TRIM$(GRID.TextMatrix(i&, CAMPOLABEL)), 24)
        .SetData i&, 1, XVALO(GRID.TextMatrix(i&, SERIE1)), 0 'DATOS DE LA SERIE  (DEBE SER NUMERICO)
        If SERIE2 > 0 Then .SetData i&, 2, XVALO(GRID.TextMatrix(i&, SERIE2)), 0 'DATOS DE LA SERIE  (DEBE SER NUMERICO)
    End With
 Next i&
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   Me.MC.Width = Me.Width - 300
   Me.MC.Height = Me.Height - 300
   On Error GoTo 0

End Sub
