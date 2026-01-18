VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FRMINSERTAR 
   Caption         =   "Mover/Inserar/Eliminar Fila "
   ClientHeight    =   1695
   ClientLeft      =   1920
   ClientTop       =   3315
   ClientWidth     =   14550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleMode       =   0  'User
   ScaleWidth      =   13726.42
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame2 
      Height          =   580
      Left            =   4680
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
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
         Height          =   285
         Left            =   840
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "En Fila:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   290
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   580
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   4695
      Begin VB.OptionButton Option2 
         Caption         =   "Insertar y Desplazar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Insertar y Eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   4200
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2520
      OleObjectBlob   =   "FRMINSERTAR.frx":0000
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insertar"
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
      Left            =   10920
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1150
      Width           =   1695
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "Cancelar"
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
      Left            =   12840
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1150
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   1035
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   1826
      _Version        =   393216
      BackColorFixed  =   12632256
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   855
   End
End
Attribute VB_Name = "FRMINSERTAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FORMULARIOX As Form
Public MSF  As MSFlexGrid
Private Sub cmdcancelar_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
    '
    If Option1.Value = False And Option2.Value = False Then
       Call COMUNI("Debe Seleccionar una Opción a Realizar")
       GoTo 99
    End If
    '
    If XVALO(Text1) < 1 Then GoTo 98
    '
    If XVALO(Text1) = XVALO(MSF1.TextMatrix(1, 0)) Then GoTo 98
    '
    If Option1.Value = True Then
       Call INSERTARYELMINAR(XVALO(Text1))
    Else
       Call INSERTYDESPLAZAR(XVALO(Text1))
    End If
    '
98  Unload Me
99  Command1.Enabled = True
End Sub

Private Sub Command1_GotFocus()
       Command1.BackColor = vbYellow

End Sub

Private Sub Command1_LostFocus()
    Command1.BackColor = &H8000000F
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Call cmdcancelar_Click
   End If
End Sub

Private Sub Form_Load()
    '
    '
    Timer1.Enabled = True ' PARA QUE PROVOQUE UNA DEMORA EN PINTAR EL TEXTLISTA
    'POR QUE SE VE QUE POR EL APLICAR SKIN NO FUNCIONA POR QUE PIERDE EL FOCO
 End Sub
   
Private Sub Text1_GotFocus()
    Text1.BackColor = vbYellow
    Call PINTATEXT(Text1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      If Option1.Value = True Or Option2.Value = True Then
         Command1.SetFocus
      End If
    End If
End Sub

Private Sub Text1_LostFocus()
    Text1.BackColor = vbWhite
End Sub

Private Sub Timer1_Timer()
    Call PINTATEXT(Text1)
    Timer1.Enabled = False
End Sub

Private Sub Text1_Change()
    If XVALO(Text1) > MSF.Rows - 1 Then
       MsgBox "Número de Fila Inválido"
    End If
End Sub
Sub INSERTARYELMINAR(FILA&)
   If FILA& < 1 Then Exit Sub
   Call SOBRE_ESCRIBIR_FILA(FILA&)
   Call borrafila(XVALO(MSF1.TextMatrix(1, 0)), FORMULARIOX, MSF)
End Sub
   Sub SOBRE_ESCRIBIR_FILA(FILA&)
    For i& = 1 To MSF1.Cols - 1
       DATO$ = MSF1.TextMatrix(1, i&)
       MSF.TextMatrix(FILA&, i&) = DATO$
    Next i&
   End Sub
Sub INSERTYDESPLAZAR(FILA&)
    Screen.MousePointer = 11
    '
    UR& = MSF.Rows - 1
    MSF.Rows = MSF.Rows + 1
    For U& = UR& To FILA& Step -1
      For i& = 1 To MSF.Cols - 1
         DATO$ = MSF.TextMatrix(U&, i&)
         MSF.TextMatrix(U& + 1, i&) = DATO$
      Next i&
    Next U&
    '
    'SE REALIZA ESTA CONSIDERACION POR QUE CAMBIA CUANDO SE SELECCIONA UNA POSICION MAYOR Y SE INTENTA INCLUIR EN UNA INFERIOR POR EL HECHO
    'QUE ESTO FUNCIONA PRIMERO CREANDO UN  REGISTRO PARA DEJAR LIBRE LA FILA DONDE SE VA A INSERTAR LUEGO SE SOBRE ESCRIBE ESTA
    'Y POR ULTIMO SE BORRA EL REGISTRO DE LA POSICION ORIGINAL
    valor% = 0
    VALOR2% = 1
    If FILA& > XVALO(MSF1.TextMatrix(1, 0)) Then
       valor% = 1
       VALOR2% = 0
    End If
    '
    Call SOBRE_ESCRIBIR_FILA(FILA& + valor%) ' SOBRE ESCRIBE LA FILA
    Call borrafila(XVALO(MSF1.TextMatrix(1, 0)) + VALOR2%, FORMULARIOX, MSF)  ' BORRA LA FILA ORIGINAL
    Screen.MousePointer = 1
    '
End Sub

