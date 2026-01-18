VERSION 5.00
Begin VB.Form LISECUEN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Secuencias de Demostración Disponibles"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      Left            =   735
      TabIndex        =   0
      Top             =   1260
      Width           =   6945
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "haciendo un doble-click:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Left            =   2310
      TabIndex        =   2
      Top             =   735
      Width           =   5370
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Elija la secuencia a ejecutar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Left            =   210
      TabIndex        =   1
      Top             =   210
      Width           =   5370
   End
   Begin VB.Image Image1 
      Height          =   5625
      Left            =   0
      Picture         =   "LISECUEN.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8670
   End
End
Attribute VB_Name = "LISECUEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   N1% = FreeFile
   Dim BUFLE As String * 128
   Open DEMOPATH$ + "\SCRIPTS\SECUDEMO.DAT" For Random Access Read Shared As #N1% Len = 128
   For KKI% = 1 To LOF(N1%) / 128
     Get #N1%, KKI%, BUFLE$
     If Trim$(BUFLE$) <> "" Then
       List1.AddItem Trim$(BUFLE$)
     End If
   Next KKI%
   Close #N1%
End Sub

Private Sub List1_DblClick()
   Hide
End Sub
