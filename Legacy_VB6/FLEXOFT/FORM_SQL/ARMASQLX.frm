VERSION 5.00
Begin VB.Form ARMASQLX 
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   2175
      Left            =   4680
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Arma SQLX$"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "ARMASQLX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Dim strTest As String
   Dim strArray() As String
   Dim intCount As Integer
   
   strArray = Split(Text1, Chr(13))
   Text2 = ""
   For i& = LBound(strArray) To UBound(strArray)
     If i& = 0 Then
      Text2 = Text2 + "SQLX$ = " & Chr(34) & " " & strArray(i&) & " " & Chr(34) & vbCrLf
     Else
      Text2 = Text2 + "SQLX$ = SQLX$ + " & Chr(34) & " " & strArray(i&) & " " & Chr(34) & vbCrLf
     End If
   Next i&
End Sub
            
