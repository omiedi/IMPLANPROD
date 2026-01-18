VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Begin VB.Form GRAFUNPR 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis Estadistico de Ventas"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   Begin MSChartLib.MSChart MSChart1 
      Height          =   4000
      Left            =   105
      OleObjectBlob   =   "GRAFUNPR.frx":0000
      TabIndex        =   0
      Top             =   105
      Width           =   11700
   End
   Begin MSChartLib.MSChart MSChart2 
      Height          =   2000
      Left            =   105
      OleObjectBlob   =   "GRAFUNPR.frx":1E57
      TabIndex        =   1
      Top             =   4200
      Width           =   11700
   End
   Begin MSChartLib.MSChart MSChart3 
      Height          =   1995
      Left            =   105
      OleObjectBlob   =   "GRAFUNPR.frx":38C0
      TabIndex        =   2
      Top             =   6300
      Width           =   11700
   End
End
Attribute VB_Name = "GRAFUNPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Hide
End Sub


Private Sub Form_Load()
    Call CENTRAFORM(Me)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub

Private Sub MSChart1_DblClick()
   If MSChart1.Height < 7900 Then
      MSChart1.Top = 100
      MSChart1.Height = 4000
      MSChart1.Refresh
      MSChart2.Top = 4200
      MSChart2.Height = 2000
      MSChart2.Refresh
      MSChart3.Top = 6300
      MSChart3.Height = 2000
      MSChart3.Refresh
   End If
End Sub

Private Sub MSChart2_DblClick()
MSChart1.Top = 100
MSChart1.Height = 2000
MSChart1.Refresh
MSChart2.Top = 2200
MSChart2.Height = 4000
MSChart2.Refresh
MSChart3.Top = 6300
MSChart3.Height = 2000
MSChart3.Refresh
End Sub

Private Sub MSChart3_DblClick()
MSChart1.Top = 100
MSChart1.Height = 2000
MSChart1.Refresh
MSChart2.Top = 2200
MSChart2.Height = 2000
MSChart2.Refresh
MSChart3.Top = 4300
MSChart3.Height = 4000
MSChart3.Refresh

End Sub
