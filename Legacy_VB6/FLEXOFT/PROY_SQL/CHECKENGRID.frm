VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CHECKENGRID 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5106
      _Version        =   393216
   End
End
Attribute VB_Name = "CHECKENGRID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  MSF.Rows = 6
  MSF.Cols = 6
End Sub

Private Sub MSF_Click()
    A = MSF.MouseCol
    B = MSF.MouseRow
    With MSF
        .CellFontName = "Wingdings"
        .CellFontSize = 14
        .CellAlignment = flexAlignCenterCenter
        ' edita la celda
        If .TextMatrix(B, A) = Chr(254) Then
            .TextMatrix(B, A) = Chr(168) ' false
        Else
            .TextMatrix(B, A) = Chr(254) ' true
        End If
    End With
End Sub
