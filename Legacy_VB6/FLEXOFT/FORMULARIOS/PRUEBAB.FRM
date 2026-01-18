VERSION 5.00
Begin VB.Form PRUEBAB 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   105
      TabIndex        =   3
      Top             =   2205
      Width           =   6315
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   1155
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1470
      Width           =   5160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   540
      Left            =   735
      TabIndex        =   1
      Top             =   840
      Width           =   2955
   End
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   105
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   210
      Width           =   4425
   End
End
Attribute VB_Name = "PRUEBAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
archi$ = "c:\autoexec.bat"
aa$ = FileDateTime(archi$)
Exit Sub
TINI = Timer
'Reset
lu$ = "\clientes\garin\datos\"
NX% = FreeFile
Dim TTXX As String * 32767
Open lu$ + "ECOMAST.DAT" For Random Access Read Write Shared As #NX% Len = 32767
LONGI& = 1 + LOF(NX%) / 32767
LOFI& = LOF(NX%)
TTYY$ = ""
For KK& = 1 To LONGI&
Get #NX%, KK&, TTXX$
TTYY$ = TTYY$ + TTXX$
Next KK&
AAA = Len(TTYY$)
'BBB = InStr(TTYY$, "LTC018GO")
'List1.Clear
'For KK& = 1 To LOFI& Step 64
'List1.AddItem Mid$(TTYY$, KK&, 62)
'Next KK&
Text2.Text = Timer - TINI & "     " & KK&
End Sub
