VERSION 5.00
Begin VB.Form RETEDAT 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Datos de  Retención "
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CMBCOMPROBANTE 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1560
      TabIndex        =   6
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox TXTCERTIFICADO 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      Text            =   " "
      Top             =   720
      Width           =   2385
   End
   Begin VB.TextBox TXTLETRA 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Text            =   " "
      Top             =   1200
      Width           =   2385
   End
   Begin VB.CommandButton Command3 
      Height          =   1320
      Left            =   4080
      Picture         =   "RETEDAT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   550
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2D3C0&
      Caption         =   "Certificado/Código de Operación:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   -15
      TabIndex        =   5
      Top             =   645
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2D3C0&
      Caption         =   "Letra:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2D3C0&
      Caption         =   "Tipo de Comprobante:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   -15
      TabIndex        =   3
      Top             =   105
      Width           =   1455
   End
End
Attribute VB_Name = "RETEDAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command3_Click()
   If Me.CMBCOMPROBANTE.Visible = True Then
      LETR$ = Left$(TRIM$(Me.CMBCOMPROBANTE.TEXT), 1)
      If LETR$ <> "F" And LETR$ <> "R" And LETR$ <> "D" And LETR$ <> "C" And LETR$ <> "O" Then
         Call COMUNI("Tipo de Comprobante No Válida")
         Exit Sub
      End If
   End If
   Hide
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
End Sub

Private Sub Form_Load()
   CMBCOMPROBANTE.AddItem ""
   CMBCOMPROBANTE.AddItem "F  Factura"
   CMBCOMPROBANTE.AddItem "R  Recibo"
   CMBCOMPROBANTE.AddItem "D  Débito"
   CMBCOMPROBANTE.AddItem "C  Crédito"
   CMBCOMPROBANTE.AddItem "O  Ord.Pago"
   
   CMBCOMPROBANTE.TEXT = "F  Factura"

   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)
End Sub

Private Sub TXTCERTIFICADO_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TXTIPOCOMPROBANTE_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TXTLETRA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{tab}"
End Sub
