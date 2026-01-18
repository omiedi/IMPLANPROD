VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form EDITLPAG 
   Caption         =   "Ingreso de Pago Real."
   ClientHeight    =   1965
   ClientLeft      =   1920
   ClientTop       =   3315
   ClientWidth     =   9690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   34.66
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   170.921
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox CUENTA 
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
      Index           =   2
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "Datos de Comprobante de Pago"
      Top             =   822
      Width           =   2895
   End
   Begin VB.TextBox TXTTICKET 
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
      Left            =   240
      TabIndex        =   15
      Top             =   1134
      Width           =   9375
   End
   Begin VB.TextBox TXTID 
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
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1644
      Width           =   1575
   End
   Begin VB.TextBox CUENTA 
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
      Index           =   1
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "Id."
      Top             =   1644
      Width           =   495
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
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   390
      Width           =   4065
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
      Left            =   8010
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "Valor Real"
      Top             =   120
      Width           =   1575
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
      Left            =   7170
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Còd.Imp."
      Top             =   120
      Width           =   855
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
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Descripción"
      Top             =   120
      Width           =   4065
   End
   Begin VB.TextBox CUENTA 
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
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Cta."
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox TXTCUENTA 
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
      TabIndex        =   7
      Top             =   390
      Width           =   1335
   End
   Begin VB.TextBox TXTVALORREAL 
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
      Left            =   8010
      TabIndex        =   6
      Top             =   390
      Width           =   1575
   End
   Begin VB.TextBox TXTCODIGOIMP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   7170
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "   "
      Top             =   390
      Width           =   855
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
      Left            =   5595
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "Importe"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox TXTIMPORTE 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   5595
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "   "
      Top             =   390
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   4200
      Top             =   2160
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2520
      OleObjectBlob   =   "EDITLPAG.frx":0000
      Top             =   2160
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
      Left            =   5520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1560
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
      Left            =   7920
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   855
   End
End
Attribute VB_Name = "EDITLPAG"
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



Private Sub Command1_Click()
   'ACTUALIZAR
    If TRIM$(TXTCUENTA.TEXT) = "" Then
       MsgBox "CUENA INEXISTENTE O NO VÁLIDO"
       GoTo 99
    End If
    '
    MODO1$ = "ACTUALIZAR"
99 Hide

End Sub

Private Sub Form_Load()
    '
    Timer1.Enabled = True ' PARA QUE PROVOQUE UNA DEMORA EN PINTAR EL TEXTLISTA
    'POR QUE SE VE QUE POR EL APLICAR SKIN NO FUNCIONA POR QUE PIERDE EL FOCO
 End Sub
   
Private Sub Timer1_Timer()
    Call PINTATEXT(TXTVALORREAL)
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

Private Sub TXTVALORREAL_LostFocus()
     TXTVALORREAL = FORMATNUM$(XVALO(TXTVALORREAL), "F12.2")
End Sub
