VERSION 5.00
Begin VB.Form MSJESREGINFO 
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8565
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3480
      Width           =   3135
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "MSJESREGINFO.frx":0000
      Left            =   120
      List            =   "MSJESREGINFO.frx":0002
      TabIndex        =   1
      Top             =   480
      Width           =   8295
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000003&
      Caption         =   "Lista de errores encontrados durante la generación:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   160
      Width           =   4575
   End
End
Attribute VB_Name = "MSJESREGINFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Command1_Click()

    Dim ContinuaCiclo As Boolean
    Dim PuntoEncuentro As Integer
    
    ListaDeErrores$ = Text1.TEXT
    ContinuaCiclo = True
    Do While ContinuaCiclo
        PuntoEncuentro = InStr(1, ListaDeErrores$, vbCrLf)
        If PuntoEncuentro = 0 And Len(ListaDeErrores$) > 0 Then PuntoEncuentro = Len(ListaDeErrores$)
        If PuntoEncuentro > 0 Then
            List1.AddItem Mid$(ListaDeErrores$, 1, PuntoEncuentro - 1)
            ListaDeErrores$ = Mid$(ListaDeErrores$, PuntoEncuentro + 1)
        Else
            ContinuaCiclo = False
        End If
    Loop

End Sub

Private Sub Form_Load()
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)
End Sub

Private Sub Text1_GotFocus()
    Call Me.SetFocus
End Sub


