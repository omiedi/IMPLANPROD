VERSION 5.00
Begin VB.Form PARDI015 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ensayos Bombas Dosificadoras"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11910
   Begin VB.Frame Frame12 
      Caption         =   "Frame12"
      Height          =   5895
      Left            =   10300
      TabIndex        =   24
      Top             =   1470
      Width           =   15
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   5880
      Index           =   1
      Left            =   3570
      TabIndex        =   16
      Top             =   1470
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   5880
      Left            =   1365
      TabIndex        =   18
      Top             =   1470
      Width           =   15
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   5880
      Index           =   0
      Left            =   2625
      TabIndex        =   17
      Top             =   1470
      Width           =   15
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      Height          =   435
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   10545
      TabIndex        =   20
      Top             =   1470
      Width           =   10600
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Bomba  Modelo     Caudal Observaciones "
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   45
         TabIndex        =   21
         Top             =   105
         Width           =   10305
      End
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5460
      ItemData        =   "PARDI015.frx":0000
      Left            =   105
      List            =   "PARDI015.frx":0002
      TabIndex        =   19
      Top             =   1890
      Width           =   10605
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ensayo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   6300
      TabIndex        =   8
      Top             =   105
      Width           =   4425
      Begin VB.CommandButton Command9 
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
         Height          =   285
         Left            =   4095
         TabIndex        =   23
         Top             =   315
         Width           =   180
      End
      Begin VB.CommandButton Command8 
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
         Height          =   285
         Left            =   2100
         TabIndex        =   22
         Top             =   315
         Width           =   180
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1155
         TabIndex        =   13
         Top             =   705
         Width           =   3120
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Operario:"
         Height          =   225
         Left            =   2205
         TabIndex        =   12
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label Label71 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1155
         TabIndex        =   11
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label73 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3465
         TabIndex        =   10
         Top             =   315
         Width           =   645
      End
      Begin VB.Label Label72 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   225
         Left            =   -105
         TabIndex        =   9
         Top             =   315
         Width           =   1185
      End
   End
   Begin VB.TextBox Text25 
      Height          =   750
      Left            =   1260
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   7455
      Width           =   9465
   End
   Begin VB.CommandButton Command6 
      Caption         =   "New"
      Height          =   855
      Left            =   10920
      Picture         =   "PARDI015.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Nuevo Modelo de Bomba"
      Top             =   3150
      Width           =   855
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   10920
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   6
      Top             =   7430
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   735
         TabIndex        =   7
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton command1 
      Caption         =   "Exit"
      Height          =   855
      Left            =   10920
      Picture         =   "PARDI015.frx":014E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Help"
      Height          =   855
      Left            =   10920
      Picture         =   "PARDI015.frx":0298
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1260
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      Height          =   855
      Left            =   10920
      Picture         =   "PARDI015.frx":05A2
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guarda e Imprime Informe de Ensayo"
      Top             =   6500
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Setup"
      Height          =   855
      Left            =   10920
      Picture         =   "PARDI015.frx":0C0C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Configuración - Formularios de Impresión Informes"
      Top             =   2205
      Width           =   855
   End
   Begin VB.Line Line7 
      X1              =   10720
      X2              =   10720
      Y1              =   7455
      Y2              =   8190
   End
   Begin VB.Line Line6 
      X1              =   1260
      X2              =   10710
      Y1              =   8200
      Y2              =   8200
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Observ:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   25
      Top             =   7455
      Width           =   1080
   End
   Begin VB.Line Line3 
      X1              =   10710
      X2              =   10710
      Y1              =   1470
      Y2              =   7350
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   105
      X2              =   10710
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   525
      X2              =   6090
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "015"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   4725
      TabIndex        =   15
      Top             =   570
      Width           =   1380
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Parte Diario Prueba"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   0
      TabIndex        =   14
      Top             =   660
      Width           =   4530
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   10710
      Picture         =   "PARDI015.frx":104E
      Stretch         =   -1  'True
      Top             =   7630
      Width           =   4725
   End
End
Attribute VB_Name = "PARDI015"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Private Sub BOTSEL_Click(Index As Integer)

End Sub

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub Command5_Click()
   VTB% = VENTABU%("PARDIDOS")
End Sub

Private Sub Command3_Click()
   '
   URN& = ULTREG&("PARDIDOS")
   RNC$ = VECPUNPAR$
   OBSE$ = Text25
   '
   For KK& = 1 To Len(RNC$) Step 4
     RENOTA& = CVS(Mid$(RNC$, KK&, 4))
     TTYY$ = REGLEIDO$("PARDIDOS", RENOTA&)
     Call REPLA(TTYY$, OBSE$, 129, 384)
     Call GRAREG("PARDIDOS", TTYY$, RENOTA&)
   Next KK&
   Call CIERRARCH("PARDIDOS")
   '
   Call PRIENSADOS
   '
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   CONFENSA.Show 1
   Command4.Enabled = True
End Sub

Private Sub Command6_Click()
   '
   Command6.Enabled = False
   VDEF$ = REGBLAN$("MOBOMDOS")
   '
10 Screen.MousePointer = 1
   OBX$ = OBJPLA$("NUEMOBOMDVR/Nuevo Modelo Bomba Dosificadora", VDEF$)
   If OBX$ = "" Then
     GoTo 99
   End If
   '
   VDEF$ = OBX$
   CODINS$ = UCase$(TRIM$(Left$(OBX$, 4)))
   If CODINS$ = "" Then
      Call MENSERR(24, "Falta Modelo de Bomba")
      GoTo 10
   End If
   Call REPLA(OBX$, CODINS$, 1, 4)
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("MOBOMDOS")
   For UI& = 1 To FIN&
      XX$ = REGLEIDO$("MOBOMDOS", UI&)
      If UCase$(TRIM$(Left$(XX$, 4))) = CODINS$ Then
         Call MENSERR(24, "Modelo Pre-Existente")
         GoTo 10
      End If
   Next UI&
   '
   MARCAMOT$ = TRIM$(Mid$(OBX$, 5, 28))
   If MARCAMOT$ = "" Then
      Call MENSERR(24, "Falta Descripción.")
      GoTo 10
   End If
   '
   Call CIERRARCH("MOBOMDOS")
   Call BLOQARCH("MOBOMDOS")
   Call REGAPP("MOBOMDOS", OBX$)
   Call CIERRARCH("MOBOMDOS")
   Call FRESHECHO("MOBOMDOS")
   Label4 = Left$(OBX$, 4)
   Call BAJALDISCO
   '
   Screen.MousePointer = 1
   '
99 Command6.Enabled = True
   '
End Sub


Private Sub Command8_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Label71 = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Command9_Click()
   Call SELECHO("PADRON")
   VDEVU$ = TRIM$(VALACT1$("PADRON"))
   If VDEVU$ <> "" Then
      Label73 = VALACT1$("PADRON")
   End If
End Sub

Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    HOII% = HOY(HOS$)
    Label71 = HOS$
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub

Private Sub Label4_Change()
   Call MUEPADIAR
End Sub

Private Sub Label4_DblClick()
   XX$ = REGSEL$("MOBOMDOS")
   If Len(XX$) > 4 Then
     Label4 = Left$(XX$, 4)
   End If
End Sub

Private Sub Label71_Change()
   Call MUEPADIAR
End Sub

Private Sub Label71_DblClick()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Label71 = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Label73_Change()
   Label2 = ""
   If Val(Label73) > 0 Then
     Label2 = ECOARCH$("PADRON", MKI$(Val(Label73)))
     Call MUEPADIAR
   End If
End Sub

Private Sub Label73_DblClick()
   Call SELECHO("PADRON")
   VDEVU$ = TRIM$(VALACT1$("PADRON"))
   If VDEVU$ <> "" Then
      Label73 = VALACT1$("PADRON")
   End If
End Sub

Sub MUEPADIAR()
   '
   Dim MODEPAR$(32), TENSA%(32), TAPRO%(32), TRECH%(32), QRECH%(10), MRECH%(32, 10)
   '
   List1.Clear
   Text25 = ""
   '
   FF% = FECHANUM(Label71)
   If FF% > 0 Then
     NOPE% = Val(Label73)
     If NOPE% > 0 Then
       RFF$ = VECPUNPAR$
       JJ& = 0
       List1.Clear
       CAMODE% = 0
       For KKI& = 1 To Len(RFF$) Step 4
         REG& = CVS(Mid$(RFF$, KKI&, 4))
         XX$ = REGLEIDO$("PARDIDOS", REG&)
         '
         If KKI& = 1 Then
           Text25 = Mid$(XX$, 129, 384)
         End If
         '
         YY$ = Space$(128)
         Call REPLA(YY$, Mid$(XX$, 5, 12), 1, 10)
         Call REPLA(YY$, Mid$(XX$, 21, 12), 12, 10)
         Call REPLA(YY$, FORMATNUM$(CVS(Mid$(XX$, 33, 4)), "F7.1"), 21, 7)
         Call REPLA(YY$, Mid$(XX$, 49, 80), 30, 80)
         List1.AddItem YY$
         '
20     Next KKI&
       '
       If List1.ListCount < 1 Then
         Call CARPADIAR
       End If
       '
     End If
   End If
End Sub

Sub CARPADIAR()
   '
   JJ& = 0
   REBLA$ = REGBLAN$("PARDIDOS")
   For KKI& = 1 To List1.ListCount
     YY$ = List1.List(KKI& - 1)
     XX$ = REBLA$
     Call REPLA(XX$, Mid$(YY$, 1, 10), 5, 12)
     Call REPLA(XX$, Mid$(YY$, 12, 9), 21, 12)
     Call REPLA(XX$, MKS$(Val(Mid$(YY$, 21, 7))), 33, 4)
     Call REPLA(XX$, Mid$(YY$, 30, 80), 49, 80)
     JJ& = JJ& + 1
     Call GRAREG("!PARDIDOS", XX$, JJ&)
   Next KKI&
   '
   Call SETULTREG("!PARDIDOS", JJ&)
   Call CIERRARCH("!PARDIDOS")
   '
   VTB% = VENTABU%("PARDIDOS")
   If VTB% >= 0 Then
     '
     Screen.MousePointer = 11
     FF% = FECHANUM(Label71)
     NOPE% = Val(Label73)
     MODE$ = AJUSTI$(Label4, 4)
     OBSE$ = Text25
     '
     URN& = ULTREG&("PARDIDOS")
     RNC$ = VECPUNPAR$
     KK& = (-3)
     '
     For LL& = 1 To ULTREG&("!PARDIDOS")
       TTYY$ = REGLEIDO$("!PARDIDOS", LL&)
       Call REPLA(TTYY$, MKI$(FF%), 1, 2)
       Call REPLA(TTYY$, MKI$(NOPE%), 3, 2)
       Call REPLA(TTYY$, MODE$, 17, 4)
       Call REPLA(TTYY$, OBSE$, 129, 384)
       KK& = KK& + 4
       If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
       End If
       Call GRAREG("PARDIDOS", TTYY$, RENOTA&)
     Next LL&
     '
     While KK& < Len(RNC$) - 3
       KK& = KK& + 4
       RENOTA& = CVS(Mid$(RNC$, KK&, 4))
       Call GRAREG("PARDIDOS", String$(512, 0), RENOTA&)
     Wend
     '
     Call CIERRARCH("PARDIDOS")
     Call MUEPADIAR
     Screen.MousePointer = 1
   End If
   '
   '
End Sub

Function VECPUNPAR$()
   '
   RFF$ = ""
   FF% = FECHANUM(Label71)
   If FF% > 0 Then
     NOPE% = Val(Label73)
     If NOPE% > 0 Then
       MODE$ = AJUSTI$(Label4, 4)
       URE& = ULTREG&("PARDIDOS")
       RFF$ = ""
       RGG$ = RECFILT$("PARDIDOS", 1, MKI$(FF%))
       For KKI& = 1 To Len(RGG$) Step 4
         REG& = CVS(Mid$(RGG$, KKI&, 4))
         If REG& > 0 Then
           If REG& <= URE& Then
             XX$ = REGLEIDO$("PARDIDOS", REG&)
             If CVI(Mid$(XX$, 3, 2)) = NOPE% Then
               If Mid$(XX$, 17, 4) = MODE$ Then
                 RFF$ = RFF$ + MKS$(REG&)
               End If
             End If
           End If
         End If
       Next KKI&
     End If
   End If
   '
   VECPUNPAR$ = RFF$
   '
End Function

Private Sub List1_DblClick()
   Call CARPADIAR
End Sub

Sub PRIENSADOS()
   '
10 FORIPRE$ = TRIM$(CONTROL$("FORENSA", "FORENSA2"))
   If FORIPRE$ = "" Then
     Call MENSERR(24, "Falta Definir Formulario de Impresión")
     CONFENSA.Show 1
     Exit Sub
   End If
   '
   Screen.MousePointer = 11
   Call FRATEXTO(Text25, 64)
   JJ& = 0
   For U& = 1 To CARETEX%
     Z$ = RETEX$(U&)
     JJ& = JJ& + 1
     Call GRAREG("!OBSERVOF", Z$, JJ&)
   Next U&
   CARE& = CARETEX%
   Call SETULTREG("!OBSERVOF", CARE&)
   Call CIERRARCH("!OBSERVOF")
   '
   FECHDOC$ = Label71
   NUMEDOC$ = ""
   TIPODOC$ = "Parte Diario Prueba " + TRIM$(Label4) + " - " + TRIM$(Label71) + " - " + TRIM$(Label2)
   '
   CAPAGIN% = Int((List1.ListCount + 49) / 50)
   If CAPAGIN% < 1 Then CAPAGIN% = 1
   '
   CODPARAM$(1) = "COD-CONJ"
   CODPARAM$(2) = "VENDEDOR"
   CODPARAM$(3) = "FECH-LAN"
   CODPARAM$(4) = "FECH-VEN"
   CAPARDOC% = 4
   '
   DOCPARAM$(1) = TRIM$(Label4)
   DOCPARAM$(2) = TRIM$(Label2)
   DOCPARAM$(4) = TRIM$(Str$(CAPAGIN%))
   '
   CODTABU1$(1) = "MEDIDA-1"
   CODTABU1$(2) = "MEDIDA-2"
   CODTABU1$(3) = "MEDIDA-3"
   CODTABU1$(4) = "N-PLANO"
   CACOLTAB1% = 4
   '
   INIT% = 1
   IPAGI% = 1
   '
20 DOCPARAM$(3) = TRIM$(Str$(IPAGI%))
   
   CAITAB1% = 0
   For KKI% = INIT% To List1.ListCount
     If KKI% > INIT% + 49 Then
       INIT% = KKI%
       NOCONFIRM% = 1
       Call PRINTDOC("@" + FORIPRE$)
       IPAGI% = IPAGI% + 1
       GoTo 20
     End If
     '
     YY$ = List1.List(KKI% - 1)
     CAITAB1% = CAITAB1% + 1
     TETABU1$(1, CAITAB1%) = Left$(YY$, 10)
     TETABU1$(2, CAITAB1%) = Mid$(YY$, 12, 9)
     TETABU1$(3, CAITAB1%) = Mid$(YY$, 21, 7)
     TETABU1$(4, CAITAB1%) = Mid$(YY$, 30, 80)
   Next KKI%
   '
   NOCONFIRM% = 1
   Call PRINTDOC("@" + FORIPRE$)
   Screen.MousePointer = 1
   '
End Sub

