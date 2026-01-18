VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form SHOWRETGA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle Retencion de Ganancias"
   ClientHeight    =   5160
   ClientLeft      =   1050
   ClientTop       =   5535
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   28
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Height          =   620
      Left            =   7110
      Picture         =   "SHOWRETGA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4410
      Width           =   440
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5670
      TabIndex        =   19
      Top             =   4725
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0E0FF&
      Height          =   1905
      Left            =   210
      ScaleHeight     =   1845
      ScaleWidth      =   7140
      TabIndex        =   0
      Top             =   1050
      Width           =   7200
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Left            =   -30
         TabIndex        =   2
         Top             =   210
         Width           =   7200
      End
      Begin VB.Label Label1 
         Caption         =   "Comprobante            Importe   B.Imponible    Retenido"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   10095
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "SHOWRETGA.frx":030A
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "SHOWRETGA.frx":0384
      Top             =   1200
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Label24"
      Height          =   330
      Left            =   6405
      TabIndex        =   27
      Top             =   3675
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Parcial O/P Actual:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -840
      TabIndex        =   26
      Top             =   3465
      Width           =   3060
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Importe Neto a Retener:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3360
      TabIndex        =   25
      Top             =   4755
      Width           =   2220
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Base Imponible Neta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      TabIndex        =   24
      Top             =   4410
      Width           =   2220
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Importe no Gravado (M.N.I):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   735
      TabIndex        =   23
      Top             =   4095
      Width           =   3165
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Base ImponibleTotal:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   630
      TabIndex        =   22
      Top             =   3780
      Width           =   3270
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal - Acumulado Mensual:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   840
      TabIndex        =   21
      Top             =   3150
      Width           =   3060
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5670
      TabIndex        =   18
      Top             =   4410
      Width           =   1380
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3990
      TabIndex        =   17
      Top             =   4410
      Width           =   1590
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3990
      TabIndex        =   16
      Top             =   4095
      Width           =   1590
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3990
      TabIndex        =   15
      Top             =   3780
      Width           =   1590
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2310
      TabIndex        =   14
      Top             =   3465
      Width           =   1590
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3990
      TabIndex        =   13
      Top             =   3465
      Width           =   1590
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5670
      TabIndex        =   12
      Top             =   3150
      Width           =   1380
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3990
      TabIndex        =   11
      Top             =   3150
      Width           =   1590
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5185
      TabIndex        =   10
      Top             =   420
      Width           =   2220
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7080
      TabIndex        =   9
      Top             =   105
      Width           =   330
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Categoría:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4725
      TabIndex        =   8
      Top             =   105
      Width           =   2220
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3570
      TabIndex        =   7
      Top             =   105
      Width           =   1485
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "C.U.I.T:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2100
      TabIndex        =   6
      Top             =   105
      Width           =   1380
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1260
      TabIndex        =   5
      Top             =   420
      Width           =   3795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1260
      TabIndex        =   4
      Top             =   105
      Width           =   645
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -1050
      TabIndex        =   3
      Top             =   105
      Width           =   2220
   End
End
Attribute VB_Name = "SHOWRETGA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ALIGAN(8), CTX%

Private Sub Command1_Click()
   If TRIM$(Text1) <> "" Then
     If XVALO(Text1) < 0.005 Then
         Text1 = "0.00"
       Else
         Text1 = FORMATNUM$(XVALO(Text1), "F10.2")
     End If
     Label24 = "OK"
   End If
   Hide
End Sub

Private Sub Form_Load()
   '
   If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
    End If

   If CTX% = 0 Then
     CTX% = 1
     For U& = 1 To ULTREG&("TACRETG")
        TCX$ = REGLEIDO$("TACRETG", U&)
        TCY% = Asc(Mid$(TCX$, 1, 1))
        If TCY% > 0 Then
          If TCY% <= 8 Then
            ALIGAN(TCY%) = CVS(Mid$(TCX$, 41, 4))
          End If
        End If
     Next U&
     Call CIERRARCH("TACRETG")
   End If
   '
   Label1.Caption = "Comprobante            Importe   B.Imponible   Retencion"
    Call APLICACIONSKINS(Me)
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 1
   If XVALO(Text1) < 0.005 Then
       Text1 = "0.00"
     Else
       Text1 = FORMATNUM$(XVALO(Text1), "F10.2")
   End If
   Label24 = ""
   On Error Resume Next
   Hide
   On Error GoTo 0
End Sub

Private Sub Label10_Change()
   NETOARET# = XVALO(Label17) - XVALO(Label10)
   If NETOARET# < 0 Then NETOARET# = 0
   Text1 = TRIM$(FORMATNUM$(NETOARET#, "F11.2"))
End Sub

Private Sub Label11_Change()
   Label14 = FORMATNUM$(XVALO(Label9) + XVALO(Label11), "F12.2")
     BANETA# = XVALO(Label14) - XVALO(Label15)
     If BANETA# < 0 Then BANETA# = 0
   Label16 = FORMATNUM$(BANETA#, "F12.2")
End Sub

Private Sub Label16_Change()
   '
   RETGACAL# = 0
   '
   If XVALO(Label16) > 0 Then
     CATIB1% = XVALO(Label7)
     RETGACAL# = XVALO(Label16) * ALIGAN(CATIB1%) / 100
     '
     BASIMNE# = XVALO(Label16)
     RFF$ = RECFILT$("ESCALGAN", 1, MKI$(CATIB1%))
     If Len(RFF$) > 0 Then
       Call FILTERFILE("ESCALGAN", "!ESCALGAN", 1, MKI$(CATIB1%))
       For KU& = 1 To ULTREG&("!ESCALGAN")
         ESCX$ = REGLEIDO$("!ESCALGAN", KU&)
         If BASIMNE# > CVD(Mid$(ESCX$, 3, 8)) Then
           If BASIMNE# <= CVD(Mid$(ESCX$, 11, 8)) Then
             RETGACAL# = CVD(Mid$(ESCX$, 19, 8)) + (BASIMNE# - CVD(Mid$(ESCX$, 3, 8))) * CVS(Mid$(ESCX$, 27, 4)) / 100
             GoTo 98
           End If
         End If
       Next KU&
     End If
98 End If
   '
   Label17 = FORMATNUM(RETGACAL#, "F11.2")
   '
End Sub

Private Sub Label17_Change()
   NETOARET# = XVALO(Label17) - XVALO(Label10)
   If NETOARET# < 0 Then NETOARET# = 0
   Text1 = TRIM$(FORMATNUM$(NETOARET#, "F11.2"))
End Sub

Private Sub Label2_Change()
   Label5 = CUITCLI$((-1) * XVALO(Label2))
   Label3 = rasocli$((-1) * XVALO(Label2))
End Sub

Private Sub Label7_Change()
   Label8 = ECOARCH$("TACRETG", Chr$(XVALO(Label7)))
End Sub

Private Sub Label9_Change()
   Label14 = FORMATNUM$(XVALO(Label9) + XVALO(Label11), "F12.2")
     BANETA# = XVALO(Label14) - XVALO(Label15)
     If BANETA# < 0 Then BANETA# = 0
   Label16 = FORMATNUM$(BANETA#, "F12.2")
End Sub

