VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form NUEVOCODI 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nuevo Código"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   3525
   BeginProperty Font 
      Name            =   "Arial"
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
   ScaleHeight     =   2340
   ScaleWidth      =   3525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TXTCODIGO 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1920
      Width           =   2160
   End
   Begin VB.TextBox TXTVERSION 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1260
      Width           =   1500
   End
   Begin VB.TextBox TXTORDEN 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   1920
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   960
      Width           =   1500
   End
   Begin VB.TextBox TXTCLIENTE 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   1920
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   660
      Width           =   1500
   End
   Begin VB.TextBox TXTTIPOBOLSA 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   1920
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   1500
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "TIPO BOLSA  "
      Top             =   360
      Width           =   1800
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "ORDEN ITEM "
      Top             =   960
      Width           =   1800
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "CTA. CLIENTE "
      Top             =   660
      Width           =   1800
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "VERSIÓN "
      Top             =   1260
      Width           =   1800
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2160
      OleObjectBlob   =   "NUEVOCODI.frx":0000
      Top             =   -120
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "NUEVOCODI.frx":0234
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Código Resultante"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      TabIndex        =   2
      Top             =   4725
      Width           =   750
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      TabIndex        =   1
      Top             =   5250
      Width           =   795
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuAceptar 
         Caption         =   "Aceptar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "NUEVOCODI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mnuAceptar_Click()
   Hide
End Sub

Private Sub mnuSalir_Click()
  Unload Me
End Sub

Private Sub TXTCLIENTE_Change()
  Call ARMARCODIGO
End Sub

Private Sub TXTCLIENTE_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Call PINTATEXT(Me.TXTORDEN)
  End If

End Sub

Private Sub TXTORDEN_Change()
  Call ARMARCODIGO

End Sub

Private Sub TXTORDEN_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Call PINTATEXT(Me.TXTVERSION)
  End If

End Sub

Private Sub TXTTIPOBOLSA_Change()
   Call ARMARCODIGO
End Sub
Sub ARMARCODIGO()
   Me.TXTCODIGO = ""
   TIPOBOLSA$ = "00" & XVALO(Me.TXTTIPOBOLSA): TIPOBOLSA$ = Right$(TIPOBOLSA$, 2)
   CTACLIENTE$ = "00000" & XVALO(Me.TXTCLIENTE): CTACLIENTE$ = Right$(CTACLIENTE$, 5)
   If XVALO(TXTORDEN) < 100 Then 'PARA MANTENER 2 DIGITOS EN CODIGOS EXISTENTES
      ORDEN1$ = "00" & XVALO(Me.TXTORDEN): ORDEN1$ = Right$(ORDEN1$, 2)
   Else  ' Y SI ES MAYOR 99 LE DA FORMATO DE 3 DIGITOS
      ORDEN1$ = "000" & XVALO(Me.TXTORDEN): ORDEN1$ = Right$(ORDEN1$, 3)
   End If
   LETVERSION$ = TRIM$(Left$(TRIM$(Me.TXTVERSION), 2))
   Me.TXTCODIGO = TIPOBOLSA$ & "-" & CTACLIENTE$ & "-" & ORDEN1$ & "-" & LETVERSION$
   Me.TXTCODIGO = UCase$(Me.TXTCODIGO)
End Sub

Private Sub TXTTIPOBOLSA_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Call PINTATEXT(Me.TXTCLIENTE)
  End If
End Sub

Private Sub TXTVERSION_Change()
  If TRIM$(TXTVERSION) = "" Then Exit Sub
  If Asc(TXTVERSION) < 65 Or Asc(TXTVERSION) > 90 Then
    If Asc(TXTVERSION) < 97 Or Asc(TXTVERSION) > 122 Then
       MsgBox "SOLO LETRAS"
       TXTVERSION = ""
       Call PINTATEXT(TXTVERSION)
   End If
  End If
  Call ARMARCODIGO
End Sub

Private Sub TXTVERSION_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Call PINTATEXT(Me.TXTTIPOBOLSA)
  End If

End Sub
