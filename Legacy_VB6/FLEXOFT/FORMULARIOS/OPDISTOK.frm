VERSION 5.00
Begin VB.Form OPDISTOK 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disponibilidad General de Stocks - Configuración"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6660
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
   ScaleHeight     =   7620
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   5600
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   29
      Top             =   6615
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5565
      Picture         =   "OPDISTOK.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3675
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Presentación de Costo y Precios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   24
      Top             =   6510
      Width           =   5160
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Calculado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   27
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Precio Base (C.Rep. + 15 %)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3360
         TabIndex        =   26
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Informado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   25
         Top             =   360
         Value           =   -1  'True
         Width           =   1905
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Impresión de Situación de Stocks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   210
      TabIndex        =   18
      Top             =   5040
      Width           =   5160
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   880
         Left            =   4095
         Picture         =   "OPDISTOK.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   315
         Width           =   855
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   3750
      End
      Begin VB.CommandButton Command3 
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
         Height          =   360
         Left            =   3780
         TabIndex        =   20
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2415
         TabIndex        =   19
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   23
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Origen de Stocks Remotos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   210
      TabIndex        =   9
      Top             =   2415
      Width           =   5160
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   735
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1995
         Width           =   4230
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   735
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1470
         Width           =   4230
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   735
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   945
         Width           =   4230
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   735
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   420
         Width           =   4230
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-4:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   0
         TabIndex        =   17
         Top             =   2100
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-3:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   0
         TabIndex        =   15
         Top             =   1575
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   0
         TabIndex        =   13
         Top             =   1050
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   525
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Destino Archivo de Transferencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   6
      Top             =   1365
      Width           =   5160
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   735
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   420
         Width           =   4230
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   8
         Top             =   525
         Width           =   645
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Identificador de Sede"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   3
      Top             =   315
      Width           =   5160
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
         Height          =   360
         Left            =   3780
         MaxLength       =   8
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   420
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código Identificador de Sede:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   5
         Top             =   525
         Width           =   3480
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5565
      Picture         =   "OPDISTOK.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2310
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5600
      Picture         =   "OPDISTOK.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1365
      Width           =   855
   End
   Begin VB.CommandButton command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5600
      Picture         =   "OPDISTOK.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   420
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   5040
      Picture         =   "OPDISTOK.frx":0D72
      Top             =   6950
      Width           =   2010
   End
End
Attribute VB_Name = "OPDISTOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    IDENSEDE$ = TRIM$(Text1.TEXT)
    Call GRACONTROL("", "IDENSEDE", Text1.TEXT)
    '
    DESTITRA$ = TRIM$(Text2.TEXT)
    Call GRACONTROL("", "DESTITRA", Text2.TEXT)
    '
    For U& = 0 To 3
       ORITRA$(U&) = TRIM$(Text3(U&).TEXT)
    Next U&
    '
    Call GRACONTROL("", "ORITRA-1", Text3(0).TEXT)
    Call GRACONTROL("", "ORITRA-2", Text3(1).TEXT)
    Call GRACONTROL("", "ORITRA-3", Text3(2).TEXT)
    Call GRACONTROL("", "ORITRA-4", Text3(3).TEXT)
    '
    Call GRACONTROL("", "FORDISTO", Text4.TEXT)
    '
    MODOCORE% = 1: MODOCOREP$ = "COREPOS"
    If Option2.Value = True Then
         MODOCORE% = 2: MODOCOREP$ = "PREBASE"
       ElseIf Option3.Value = True Then
         MODOCORE% = 3: MODOCOREP$ = "COSCALCU"
    End If
    Call GRACONTROL("SITUASTO", "MODOCORE", MODOCOREP$)
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text4.TEXT = VDEVU$
        Text5.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    HUBOERR% = 0
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DOSIVA") > 0 Then
      Call DESHASFECHA(DES%, HAS%, PERIO$)
      If PERIO$ <> "" Then
        Screen.MousePointer = 11
        FIN& = ULTREG&("MOVISTO")
        For U& = 1 To FIN&
          Call TRACE(20, U&, FIN&)
          XX$ = REGLEIDO$("MOVISTO", U&)
          FF1% = CVI(Mid$(XX$, 1, 2))
          If FF1% >= DES% Then
            If FF1% <= HAS% Then
              CMOV$ = TRIM$(Mid$(XX$, 21, 2))
              If CMOV$ = "RT" Then
                CI1% = CVI(Mid$(XX$, 3, 2))
                DS1$ = TRIM$(Mid$(XX$, 33, 10))
                RF1$ = TRIM$(Mid$(XX$, 43, 30))
                QT1# = CVD(Mid$(XX$, 96, 8))
                If QT1# < 0.05 Then GoTo 19    ' NO ES MOVIMIENTO DE SALIDA
                '
                For U2& = U& + 1 To FIN&
                  YY$ = REGLEIDO$("MOVISTO", U2&)
                  If Mid$(YY$, 21, 2) = "RP" Then
                    FF2% = CVI(Mid$(YY$, 1, 2))
                    CI2% = CVI(Mid$(YY$, 3, 2))
                    DS2$ = TRIM$(Mid$(YY$, 33, 10))
                    QT2# = CVD(Mid$(YY$, 88, 8))
                    If FF2% = FF1% Then
                     If CI2% = CI1% Then
                      If DS2$ = DS1$ Then
                       If Abs(QT2# - QT1#) < 0.005 Then
                         GoTo 19    ' COINCIDE REGISTRO DE PL
                       End If
                      End If
                     End If
                    End If
                  End If
                Next U2&
                '
                RFF$ = RECFILT$("MOVISTO", 1, MKI$(FF1%))
                For KKX& = 1 To Len(RFF$) Step 4
                  U2& = CVS(Mid$(RFF$, KKX&, 4))
                  If U2& <> U& Then
                    YY$ = REGLEIDO$("MOVISTO", U2&)
                    If Mid$(YY$, 21, 2) = "RT" Then
                      FF2% = CVI(Mid$(YY$, 1, 2))
                      CI2% = CVI(Mid$(YY$, 3, 2))
                      RF2$ = TRIM$(Mid$(YY$, 43, 30))
                      QT2# = CVD(Mid$(YY$, 88, 8))
                      If FF2% = FF1% Then
                       If CI2% = CI1% Then
                        If RF2$ = "Factura " + DS1$ Then
                         If Abs(QT2# - QT1#) < 0.005 Then
                          GoTo 19    ' COINCIDE REGISTRO DE PL
                         End If
                        End If
                       End If
                      End If
                    End If
                  End If
                Next KKX&
                '
                HUBOERR% = 1
                FEX = FF1%
                DOPRI1$ = TRIM$(Mid$(XX$, 23, 10))
                TTXX$ = "Error en Código '" + TRIM$(CODEXT$(CI1%)) + " - Remito " + DOPRI1$ + " del " + FECHATEX$(FEX)
                TTXX$ = TTXX$ + "\¿ Desea Recuperar el Movimiento del 'Packing List' ?\\Si, Recuperar\No, Conservar"
                If COMALTER%(TTXX$) = 1 Then
                  ZZ$ = XX$
                  Call BLOQARCH("MOVISTO")
                  Call REPLA(ZZ$, "RP", 21, 2)
                  Call REPLA(ZZ$, "PL", 23, 2)
                  Call REPLA(ZZ$, "Packing List " + RF1$, 43, 30)
                  Call REPLA(ZZ$, MKI$(0), 84, 2)
                  Call REPLA(ZZ$, MKD$(QT1#), 88, 8)
                  Call REPLA(ZZ$, MKD$(0), 96, 8)
                  Call REGAPP("MOVISTO", ZZ$)
                  Call CIERRARCH("MOVISTO")
                End If
                '
              End If
            End If
          End If
19      Next U&
      End If
    End If
    '
    TTXX$ = "Control Completo"
    If HUBOERR% > 0 Then TTXX$ = TTXX$ + " con Errores"
    Call COMUNI(TTXX$)
    Screen.MousePointer = 1
    '
'
'           If Left$(DOPRI$, 2) = "RT" Then
'             DOPRIQ$ = "PL." + TRIM$(Mid$(XQ$, 371, 16))
'             If TRIM$(DOPRIQ$) = "PL." Then   ' REMITO DIRECTO SIN O/VENTA
'               NRXLX& = 900000 + Val(Text10)
'               DOPRIQ$ = DOPRIQ$ + TRIM$(Str$(NRXLX&))
'             End If
'             REFEQ$ = "Packing-List " + REFE$
'             If CMOV$ = "*R" Then
'                  CMOVQ$ = "*D"
'               Else
'                  CMOVQ$ = "RP"
'             End If
'             Call MOVISTO(FF%, CI%, LOTE$, CMOVQ$, DOPRIQ$, DOSEC$, REFEQ$, 0, "$", 0, DEPOSALI%, (-1) * SIG% * CANTAA)
'           End If
'         End If
'         '
'
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text4.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Dim CMX$(128), QTM(128)
   '
   CATIMO% = 0
   FIN& = ULTREG&("MOVISTO")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("MOVISTO", U&)
     CMOV$ = TRIM$(Mid$(XX$, 21, 2))
     For KKI% = 1 To CATIMO%
       If CMX$(KKI%) = CMOV$ Then
         QTM(KKI%) = QTM(KKI%) + 1
         GoTo 19
       End If
     Next KKI%
     CATIMO% = CATIMO% + 1
     CMX$(CATIMO%) = CMOV$
     QTM(CATIMO%) = 1
19 Next U&
   '
   For KKI% = 1 To CATIMO%
     MsgBox CMX$(KKI%) & "    " & QTM(KKI%)
   Next KKI%
   '
End Sub

Private Sub Form_Load()
    '
    Call TRANSLABELS(Name)
    '
    Call CARUTRA
    '
    Text1.TEXT = IDENSEDE$
    Text2.TEXT = DESTITRA$
    For U& = 0 To 3
       Text3(U&).TEXT = ORITRA$(U&)
    Next U&
    '
    Text4.TEXT = CONTROL$("", "FORDISTO")
    Text5.TEXT = ECOARCH$("INDIFRM", Text4.TEXT)
    '
    Option1.Value = True
    MODOCORE% = 1: MODOCOREP$ = "COREPOS"
    If TRIM$(UCase$(CONTROL$("SITUASTO", "MODOCORE"))) = "PREBASE" Then
       Option2.Value = True
       MODOCORE% = 2: MODOCOREP$ = "PREBASE"
     ElseIf TRIM$(UCase$(CONTROL$("SITUASTO", "MODOCORE"))) = "COSCALCU" Then
       Option3.Value = True
       MODOCORE% = 3: MODOCOREP$ = "COSCALCU"
    End If
    '
End Sub

Private Sub Text4_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text4.TEXT = VDEVU$
        Text5.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text5_Click()
   Text4.SetFocus
End Sub
