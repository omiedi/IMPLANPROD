VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.ocx"
Begin VB.Form GRABAR07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis Gráfico FLEXOFT"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
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
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   135
      Left            =   840
      OleObjectBlob   =   "GRABAR07.frx":0000
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy"
      Height          =   645
      Left            =   9840
      Picture         =   "GRABAR07.frx":1F23
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Exportar - Copia al Portapapeles"
      Top             =   105
      Width           =   645
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      Height          =   645
      Left            =   10500
      Picture         =   "GRABAR07.frx":222D
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprime Ficha"
      Top             =   105
      Width           =   645
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   645
      Left            =   11130
      Picture         =   "GRABAR07.frx":2897
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   105
      Width           =   645
   End
End
Attribute VB_Name = "GRABAR07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DESDEFE%, HASTAFE%
Dim CORRESP$, ABREMES$, CORRESQ$
Dim DESFE$, HASFE$
Dim CODIMES%
Dim CAMESA%

Private Sub Command1_Click()
    MSChart1.EditCopy
    MsgBox "Grafico Copiado al Portapapeles"
End Sub
'
Private Sub Command2_Click()
   Call CIERRARCH("*.*")
   Unload Me
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Printer.TrackDefault = False
   Printer.Orientation = 2
   PRINTFORM
   Printer.Orientation = 1
   Printer.TrackDefault = True
   Command3.Enabled = True
End Sub

Private Sub Form_Load()
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
     Caption = Caption + "  -  " + EPAC$
   End If
   '
   Call CENTRAFORM(Me)
   Refresh
   DoEvents
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub
