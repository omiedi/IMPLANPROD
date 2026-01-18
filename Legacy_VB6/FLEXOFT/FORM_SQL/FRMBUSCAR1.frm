VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FRMBUSCAR1 
   Caption         =   "BUSCAR"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5670
   Begin VB.ListBox List2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   495
      Left            =   120
      OleObjectBlob   =   "FRMBUSCAR1.frx":0000
      TabIndex        =   2
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   2520
      Width           =   4455
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FRMBUSCAR1.frx":006C
      Top             =   2280
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "FRMBUSCAR1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
'
End Sub

Private Sub Form_Load()
   Me.Top = 500
   Me.Left = 500
   Me.Height = 3690
   Me.Width = 5790
   Me.Top = mdiprincipal.Top + mdiprincipal.Height / 2 - Me.Height / 2
   Me.Left = mdiprincipal.Left + mdiprincipal.Width / 2 - Me.Width / 2
   Label1 = ""
   
   Call APLICAR_SKIN(Me, App.Path & "\SKINS\dogmax.skn")
End Sub

Private Sub List1_DblClick()
   If Me.Caption = "USUARIOS" Then
      ABMUSUARIOS.Label9 = XVALO(Mid$(List1.Text, 1, 10))
   ElseIf Me.Caption = "CLIENTES" Then
      FRMCLIENTES.Label9 = XVALO(Mid$(List1.Text, 1, 10))
   ElseIf Me.Caption = "DERECHOS" Then
      FRMDERAC.Label2 = XVALO(Mid$(List1.Text, 1, 10))
'   ElseIf Me.Caption = "HISTORIAL - CLIENTES" Then
'      ANOTACLI.Label7 = XVALO(Mid$(List1.Text, 1, 10))
'   ElseIf Me.Caption = "POSIBLES CLIENTES" Then
'      PROSPECT.Label9 = TRIM(Mid$(List1.Text, 1, 14)) '14 ES CORRECTO
'   ElseIf Me.Caption = "HISTORIAL PROSPECT" Then
'      ANOTAPRO.Label7 = TRIM(Mid$(List1.Text, 1, 14)) '14 ES CORRECTO
'   ElseIf Me.Caption = "AGENDA PROSPECT" Then
'      AGENDA.Text1 = TRIM(Mid$(List1.Text, 1, 14)) '14 ES CORRECTO
   ElseIf Me.Caption = "BUSQUEDA" Then
      RESPUESTA_BUSQUEDA$ = TRIM$(Str$(XVALO(Mid$(List1.Text, 1, 10))))
   End If

   Unload Me
End Sub

Private Sub List2_DblClick()
   If Me.Caption = "USUARIOS" Then
      ABMUSUARIOS.Label9 = XVALO(Mid$(List2.Text, 1, 10))
   ElseIf Me.Caption = "CLIENTES" Then
      FRMCLIENTES.Label9 = XVALO(Mid$(List2.Text, 1, 10))
   ElseIf Me.Caption = "DERECHOS" Then
      FRMDERAC.Label2 = XVALO(Mid$(List2.Text, 1, 10))
'   ElseIf Me.Caption = "HISTORIAL - CLIENTES" Then
'      ANOTACLI.Label7 = XVALO(Mid$(List2.Text, 1, 10))
'   ElseIf Me.Caption = "POSIBLES CLIENTES" Then
'      PROSPECT.Label9 = TRIM(Mid$(List2.Text, 1, 14)) '14 ES CORRECTO
'   ElseIf Me.Caption = "HISTORIAL PROSPECT" Then
'      ANOTAPRO.Label7 = TRIM(Mid$(List2.Text, 1, 14)) '14 ES CORRECTO
'   ElseIf Me.Caption = "AGENDA PROSPECT" Then
'      AGENDA.Text1 = TRIM(Mid$(List2.Text, 1, 14)) '14 ES CORRECTO
   ElseIf Me.Caption = "BUSQUEDA" Then
      RESPUESTA_BUSQUEDA$ = TRIM$(Str$(XVALO(Mid$(List2.Text, 1, 10))))

   End If

   
   Unload Me
End Sub

Private Sub mnuSalir_Click()
   Unload Me
End Sub

Private Sub Text1_Change()
10 ARGU$ = UCase$(TRIM$(Text1))
   'LIMPIO EL LIST GRIS
   List2.Clear
   If ARGU$ <> "" Then
    'SI HAY UN ARGUMENTO DE BUSQUEDA CONMUTA LOS LIST
    List1.Visible = False
    List2.Visible = True
    If List1.ListCount > 0 Then
     For I& = 1 To List1.ListCount
        If InStr(UCase$(List1.List(I& - 1)), ARGU$) > 0 Then  ' SI COINCIDE LO AGREGA AL LIST GRIS
          List2.AddItem List1.List(I& - 1)
        End If
     Next I&
    End If
   Else
    List1.Visible = True
    List2.Visible = False
   End If

End Sub
