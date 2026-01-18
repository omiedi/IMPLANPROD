VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FRMEDITAHP 
   Caption         =   "Edición de Itmes de Lista de Precios Activa"
   ClientHeight    =   1320
   ClientLeft      =   1920
   ClientTop       =   3315
   ClientWidth     =   12525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleMode       =   0  'User
   ScaleWidth      =   11816.04
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   7049
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "   "
      Top             =   390
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7049
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "Cantidad"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox TXTDESCRIPCION 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   390
      Width           =   4905
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10770
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "Importe"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Descripción"
      Top             =   120
      Width           =   4905
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Código"
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtcodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   390
      Width           =   1935
   End
   Begin VB.TextBox txtlista 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   10770
      TabIndex        =   7
      ToolTipText     =   "('+')  PARA AGREGAR ('-') PARA BORRAR"
      Top             =   390
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8115
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Costo"
      Top             =   120
      Width           =   1218
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   8115
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "   "
      Top             =   390
      Width           =   1218
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   4200
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2520
      OleObjectBlob   =   "FRMEDITAHP.frx":0000
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdelinminar 
      Caption         =   "Quitar de la lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox TXTNROORIGINAL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   9330
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "   "
      Top             =   390
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9330
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "P.Venta"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   855
   End
End
Attribute VB_Name = "FRMEDITAHP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ENTER EN CUALQUIER LUGAR DEL FORMULARIO Y REALIZA CLICK EN EL BOTON
'SE DEBE MARCAR EN TRUE LA PROPIEDAD DEFAULT DEL BOTON
Public MODO1$
Private Sub cmdcancelar_Click()
   
   Unload Me

End Sub

Private Sub cmdelinminar_Click()
    'QUITAR DE LA LISTA
    If TRIM$(txtcodigo.TEXT) = "" Then
       MsgBox "CODIGO INEXISTENTE O NO VÁLIDO"
       GoTo 99
    End If

    MODO1$ = "ELIMINAR"
'    Call borrafila(LISTAPRECIOS.MSF1, LISTAPRECIOS.MSF1.Row)
    
99  Hide

End Sub


Private Sub Command1_Click()
   'ACTUALIZAR
    If TRIM$(txtcodigo.TEXT) = "" Then
       MsgBox "CODIGO INEXISTENTE O NO VÁLIDO"
       GoTo 99
    End If
    '
    MODO1$ = "ACTUALIZAR"
99 Hide

End Sub

Private Sub Form_Load()
    '
    '
    Timer1.Enabled = True ' PARA QUE PROVOQUE UNA DEMORA EN PINTAR EL TEXTLISTA
    'POR QUE SE VE QUE POR EL APLICAR SKIN NO FUNCIONA POR QUE PIERDE EL FOCO
 End Sub
   
Private Sub Timer1_Timer()
    Call PINTATEXT(txtlista)
    Timer1.Enabled = False
End Sub

Private Sub TXTLISTA_GotFocus()
   txtlista.SelStart = 0
   txtlista.SelLength = Len(txtlista.TEXT)

End Sub

Private Sub txtlista_LostFocus()

   txtlista.SelStart = 0
   txtlista.SelLength = Len(txtlista.TEXT)
   NN% = XVALO(ACLISPRE07.Text3)
   If NN% = 0 Then NN% = 4
   txtlista = FORMATNUM$(txtlista, "F14." & CStr(NN%))

End Sub
Sub ACTUALIZA_FLEX(MSF As MSFlexGrid, RENGLON%, FILA%, VALOR1$)
    '
    If MSF.Row < 1 Then Exit Sub
    '
    MSF.TextMatrix(RENGLON%, FILA%) = TRIM$(VALOR1$)
    
End Sub

