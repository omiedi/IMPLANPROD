VERSION 5.00
Begin VB.Form CIERFACIL 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre de Orden de Fabricación"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   7185
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00CDDADA&
      Caption         =   "Codigo de barras. Sin separador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5280
      TabIndex        =   4
      Top             =   1200
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   5220
   End
   Begin VB.TextBox TXTCANTI 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox TXTRECHAZO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton CMDOF 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   6650
      TabIndex        =   18
      Top             =   120
      Width           =   300
   End
   Begin VB.CommandButton CMDCM 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4290
      TabIndex        =   17
      Top             =   3000
      Width           =   175
   End
   Begin VB.CommandButton CMDDE 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   690
      TabIndex        =   15
      Top             =   3000
      Width           =   175
   End
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   3120
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   3
      Top             =   -120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   240
      TabIndex        =   21
      Top             =   4680
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4500
      TabIndex        =   20
      Top             =   3000
      Width           =   2505
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   900
      TabIndex        =   19
      Top             =   3000
      Width           =   2500
   End
   Begin VB.Label LBLDEPOCONS 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   3720
      TabIndex        =   16
      Top             =   3000
      Width           =   555
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "DEPÓSITO DE CONSUMO"
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
      Left            =   3720
      TabIndex        =   14
      Top             =   2760
      Width           =   3495
   End
   Begin VB.Label LBLDEPOENT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   550
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "DEPÓSITO DE ENTRADA"
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
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD APROBADA"
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
      Left            =   3720
      TabIndex        =   11
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD RECHAZADA"
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
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   6855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCIÓN"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ORDEN DE FABRICACIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   350
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CÓDIGO"
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
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuGrabar 
         Caption         =   "Grabar Cierre de Orden de Fabricación"
      End
      Begin VB.Menu mnusalir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "CIERFACIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public STRINGORFA$





Private Sub CMDCM_Click()
    CMDCM.Enabled = False
    Call SELECHO("TADEPOSI")
    VDEV$ = VALACT1$("TADEPOSI")
    If VDEV$ <> "" Then
      Me.LBLDEPOCONS = VDEV$
    End If
    CMDCM.Enabled = True
End Sub

Private Sub CMDDE_Click()
    CMDDE.Enabled = False
    Call SELECHO("TADEPOSI")
    VDEV$ = VALACT1$("TADEPOSI")
    If VDEV$ <> "" Then
      LBLDEPOENT = VDEV$
    End If
    CMDDE.Enabled = True

End Sub

Private Sub CMDOF_Click()
     Condi$ = "STATUORF < 3 AND (CANPROY-CANAPRO)> 0.05 ORDER BY IDSUBOR"
     Call CARGALISEL("CIEROFAB", "ORDEFABR", "IDSUBOR/A12;REFEREN/A48;ID_ORFA/A16", Condi$)
     Call SELECHO("CIEROFAB/CIERRE DE ORDEN DE FABRICACION")
     NROOF$ = TRIM$(VALACT1$("CIEROFAB"))
     If TRIM$(NROOF$) = "" Then
         Exit Sub
     End If
     Text1 = NROOF$
     Call Text1_KeyPress(13)
End Sub




Private Sub Form_Unload(Cancel As Integer)
'  If Cancel = 0 Then
'     Cancel = 1
'     Hide
'     CIERFACIL.STRINGORFA = "CERRAR"
'     OFABRI07.Command46_Click
'  End If
End Sub


Private Sub LBLDEPOCONS_Change()
   Label10.Caption = TRIM$(ECOARCH$("TADEPOSI", MKI$(XVALO(Me.LBLDEPOCONS))))

End Sub

Private Sub LBLDEPOENT_Change()
   Label6.Caption = TRIM$(ECOARCH$("TADEPOSI", MKI$(XVALO(Me.LBLDEPOENT))))

End Sub

Private Sub mnuGrabar_Click()
   STRINGORFA$ = ""
   If TRIM$(Text1) = "" Then
      Call COMUNI("Falta Seleccionar Orden de Fabricación")
      GoTo 99
   End If
   '
   If CODINT%(Label1) < 1 Then
      Call COMUNI("Código No Válido o Inexistente ")
      GoTo 99
   End If
   '
   If XVALO(Me.LBLDEPOENT) < 1 Or TRIM$(Label6) = "" Then
      Call COMUNI("Depósito de Ingreso No Válido o Inexistente")
      GoTo 99
   End If
   '
   If XVALO(Me.LBLDEPOCONS) < 1 Or TRIM$(Label10) = "" Then
      Call COMUNI("Depósito de Consumo No Válido o Inexistente")
      GoTo 99
   End If
   '
   If XVALO(Me.TXTRECHAZO) > XVALO(Me.Label5) Then
      Call COMUNI("Rechazo Mayor a Pendiente a Cerrar")
      GoTo 99
   End If
   '
   If (XVALO(Me.TXTRECHAZO) + XVALO(Me.TXTCANTI)) > XVALO(Label5) Then
      Call COMUNI("Rechazo + Cantidad Aprobada Mayor a Pendiente de Cierre")
      GoTo 99
   End If
   '
   If XVALO(Me.TXTCANTI) < 1 Then
      Call COMUNI("Falta Ingresar Cantidad")
      GoTo 99
   End If
   '
   Call CARCOMPO(Label1, Text1, TXTCANTI, VACONTROL%, 2)
   VUELTA% = 0: TXTX$ = ""
   If VACONTROL% >= 0 Then
     FIN& = ULTREG&("!MATEROF")
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       ZZ$ = REGLEIDO$("!MATEROF", U&)
       CI2% = CVI(Left$(ZZ$, 2))
       CAN = CVS(Mid$(ZZ$, 5, 4))
       STKK = STODEPOS(CI2%, XVALO(Me.LBLDEPOCONS))
       If CAN > STKK Then
          TXTX$ = TXTX$ + CODEXT$(CI2%) & " Necesidad: " & CAN & " STOCKS: " & STKK & "\"
       End If
     Next U&
     
     If TXTX$ <> "" Then
       Call COMUNI("Stocks Insuficiente para Código/s:\ " & TXTX$ & "")
       GoTo 99
     End If
   End If
   STRINGORFA$ = Text1 & Chr$(124) & SAFETEXT$(XVALO(Me.TXTCANTI)) & Chr$(124) & Me.LBLDEPOENT & Chr$(124) & (Me.LBLDEPOCONS) & Chr$(124) & TXTRECHAZO
  
99
   Hide
   OFABRI07.Hide
   OFABRI07.Command46_Click

End Sub

Private Sub mnusalir_Click()
  
  Call CIERRARCH("*.*")
  Screen.MousePointer = 1
  Hide
  OFABRI07.Hide
  OFABRI07.Command46_Click

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Check1.Value = 1 And InStr(1, Text1, "-") < 1 Then
           Text1 = Left(Text1, 2) & "-" & Mid$(Text1, 3, 6) & "-" & Mid$(Text1, 9)
        End If
        CI1% = CODIFAB%(Text1)
        Label1 = CODEXT$(CI1%)
        Label8 = DESCRIT$(CI1%)
        TXTCANTI = CANT_OF_PENDIENTEDECIERRE(CI1%, Text1)
        Label5 = TXTCANTI
        LABEL12 = ""
             
        DEPENT1% = DEPENART%(CI1%) 'DE ENTRADA QUE TIENE POR DEFAULT
        DEPOENTRA% = XVALO(Control("", "DEPOCIE"))
        If DEPOENTRA% > 0 Then DEPENT1% = DEPOENTRA%
        '
        'BUSCA PRIMERO EL DEPOSITO ASOCIADO POR EL SECTOR
        CS1 = SECTPRODOF(Text1)
        REGI_SECTOR$ = FILTERGETRECORD$("SECPROD", 1, MKS$(CS1))
        DEPCONS1% = Asc(Mid$(REGI_SECTOR$, 81, 1))
        If DEPCONS1% < 1 Then DEPCONS1% = DEPCONART%(CI1%)   'EL ARTICULO CUANDO HACEMOS EL CIERRE
        '
        Me.LBLDEPOENT = DEPENT1%
        Me.LBLDEPOCONS = DEPCONS1%
        Call PINTATEXT(TXTCANTI)
    End If
End Sub

