VERSION 5.00
Begin VB.Form PRINFORM 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generador de Formularios de Impresion"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "Ver Documentos"
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
      Left            =   4830
      TabIndex        =   26
      Top             =   2205
      Width           =   2010
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Repetir Impresion"
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
      Left            =   2520
      TabIndex        =   25
      Top             =   2205
      Width           =   2115
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Archivos Gráficos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   210
      TabIndex        =   20
      Top             =   2730
      Width           =   8100
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
         Left            =   3885
         MaxLength       =   20
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   420
         Width           =   4020
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
         Left            =   735
         MaxLength       =   40
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   945
         Width           =   7170
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del Archivo de Imagen ISOLOGO:"
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
         Left            =   50
         TabIndex        =   24
         Top             =   525
         Width           =   3690
      End
      Begin VB.Label Label4 
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
         TabIndex        =   22
         Top             =   1050
         Width           =   645
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   4725
      TabIndex        =   18
      Top             =   2940
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tipos de Objetos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   4830
      TabIndex        =   12
      Top             =   210
      Width           =   3480
      Begin VB.CommandButton BOTIOBJ 
         Caption         =   "Imagen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   2310
         TabIndex        =   19
         Top             =   525
         Width           =   950
      End
      Begin VB.CommandButton BOTIOBJ 
         Caption         =   "Param."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   2310
         TabIndex        =   17
         Top             =   1365
         Width           =   950
      End
      Begin VB.CommandButton BOTIOBJ 
         Caption         =   "Caja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   1260
         TabIndex        =   16
         Top             =   1365
         Width           =   950
      End
      Begin VB.CommandButton BOTIOBJ 
         Caption         =   "Línea"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   1260
         TabIndex        =   15
         Top             =   945
         Width           =   950
      End
      Begin VB.CommandButton BOTIOBJ 
         Caption         =   "Leyenda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1260
         TabIndex        =   14
         Top             =   525
         Width           =   950
      End
      Begin VB.CommandButton BOTIOBJ 
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   210
         TabIndex        =   13
         Top             =   525
         Width           =   950
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   8400
      Picture         =   "PRINFORM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Guarda Cambios en Formulario Activo"
      Top             =   1050
      Width           =   765
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   8400
      Picture         =   "PRINFORM.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Imprime Muestra del Formulario Activo"
      Top             =   2730
      Width           =   765
   End
   Begin VB.CommandButton Command3 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   8400
      Picture         =   "PRINFORM.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Genera Nuevo Formulario de Impresión"
      Top             =   1890
      Width           =   765
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Selección de Formulario Activo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   210
      TabIndex        =   4
      Top             =   210
      Width           =   4425
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
         Height          =   570
         Left            =   210
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "PRINFORM.frx":0C7E
         Top             =   1050
         Width           =   3950
      End
      Begin VB.CommandButton Command1 
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
         Height          =   330
         Left            =   3990
         TabIndex        =   7
         Top             =   525
         Width           =   170
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2100
         TabIndex        =   6
         Text            =   " "
         Top             =   525
         Width           =   1905
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Left            =   1365
         TabIndex        =   5
         Top             =   630
         Width           =   855
      End
   End
   Begin VB.CommandButton Command6 
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
      Height          =   750
      Left            =   8400
      Picture         =   "PRINFORM.frx":0C82
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   3570
      Width           =   750
   End
   Begin VB.CommandButton Command2 
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
      Height          =   750
      Left            =   8400
      Picture         =   "PRINFORM.frx":0F8C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Termina - Cierra la Aplicación"
      Top             =   210
      Width           =   750
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   1470
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   2940
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9400
      Y1              =   4500
      Y2              =   4500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"PRINFORM.frx":10D6
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3255
      TabIndex        =   1
      Top             =   4660
      Width           =   7365
   End
   Begin VB.Image Image3 
      Height          =   885
      Left            =   -105
      Picture         =   "PRINFORM.frx":115E
      Stretch         =   -1  'True
      Top             =   4515
      Width           =   9660
   End
End
Attribute VB_Name = "PRINFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTIOBJ_Click(Index As Integer)
    Select Case Index
      Case 0
        Call DEFLAYOUTW
      Case 1
        Call DEFLEYENW
      Case 2
        Call DEFLINEAW
      Case 3
        Call DEFCAJASW
      Case 4
        Call DEFPARAMW
      Case 5
        Call DEFIMAGEW
    End Select
End Sub

Private Sub Command1_Click()
    '
    Call SELECHO("INDIFRM")
    CODFOR$ = VALACT1$("INDIFRM")
    If TRIM$(CODFOR$) = "" Then
        Text1.SetFocus
      Else
        Text1.TEXT = CODFOR$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
    '
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Close
    Call FINAL("")
End Sub

Private Sub Command3_Click()
    Call ALTANUEFORM
End Sub

Private Sub Command4_Click()
    '
    MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
    MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
    '
    PORIAN = Printer.Orientation  ' retiene orientacion
    Printer.Orientation = 1       ' anterior
    PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "ORIENTA")))
    If PORI$ = "PAISAJE" Or PORI$ = "2" Then
        Printer.Orientation = 2
    End If
    '
    Call CARFORWIN
    'Printer.EndDoc
    Call SETSPOOL("ENDDOC", "")
    DESCRIDOC$ = "Formulario " + TRIM$(Text1.TEXT) + " - " + TRIM$(Text2.TEXT)
    HOII% = HOY(HOS$)
    DESCRIDOC$ = DESCRIDOC$ + " - " + HOS$
    Call SAVESPOOL(DESCRIDOC$, OKX%)
    '
    Printer.Orientation = PORIAN
    '
End Sub

Private Sub Command5_Click()
    '
    ISOLOGO$ = TRIM$(Text3.TEXT)
    Call GRACONTROL("", "ISOLOGO", Text3.TEXT)
    '
    RUTARGRA$ = TRIM$(Text4.TEXT)
    Call GRACONTROL("", "RUTARGRA", Text4.TEXT)
    '
    Call GRATITUFORM
    '
End Sub

Private Sub Command6_Click()
    Call HELPONLINE("PRIFORMS")
End Sub

Private Sub Command7_Click()
   '
   Printer.DrawWidth = 3
   Printer.DrawStyle = 0
   Printer.FillStyle = 1
   '
5  Call SELECHO("PDOCLST")
   DOCUNU& = Val(VALACT1$("PDOCLST"))
   If DOCUNU& < 1 Then
      Exit Sub
   End If
   REDOCU$ = REGACT$("PDOCLST")
   '
   TPSP$ = ""
   Screen.MousePointer = 11
   '
   For KKL& = 1 To ULTREG&("PDOCSPL")
     TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
       TPSP$ = TPSP$ + Mid$(TBUF$, 5)
     End If
   Next KKL&
   '
   Screen.MousePointer = 11
   On Error Resume Next
   Printer.PrintQuality = (-1)
   MATRIXPUN% = 0
   If InStr(Printer.DeviceName, "810") > 0 Then MATRIXPUN% = 1
   If InStr(Printer.DeviceName, "286") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "APEX") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "FACTURA") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "REMITO") > 0 Then MATRIXPUN% = 1
   'If InStr(Printer.DriverName, "EPSON9") > 0 Then MATRIXPUN% = 1
   FORMALTER.List1.Clear
   On Error GoTo 0
   '
   Printer.EndDoc        ' RESETEA IMPRESORA
   '
10 PPXX% = InStr(TPSP$, Chr$(255))
   If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call PRINTCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
   End If
   '
   Printer.EndDoc
   Screen.MousePointer = 1
   GoTo 5
   '
End Sub

Private Sub Command8_Click()
   '
   For KKI% = 0 To 5
     EDIFORMU.Picture1(KKI%).Cls
     EDIFORMU.Picture1(KKI%).Top = 0
     EDIFORMU.Picture1(KKI%).Refresh
     EDIFORMU.Picture1(KKI%).Height = 7070
     EDIFORMU.Picture1(KKI%).Width = 11700
     EDIFORMU.Picture1(KKI%).DrawWidth = 3
     EDIFORMU.Picture1(KKI%).DrawStyle = 0
     EDIFORMU.Picture1(KKI%).FillStyle = 1
   Next KKI%
   ALTUPAGINA = 7070
   TOPEPAGINA = 0
   PAGINACTIVA% = 0
   CAPAGINAS% = 0
   '
5  Call SELECHO("PDOCLST")
   DOCUNU& = Val(VALACT1$("PDOCLST"))
   If DOCUNU& < 1 Then
      Exit Sub
   End If
   REDOCU$ = REGACT$("PDOCLST")
   ORIPRINTER$ = TRIM$(Mid$(REDOCU$, 129))
   '
   URE& = ULTREG&("PDOCSPL")
   LI& = 0: LS& = URE&
6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
   DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
   GoTo 8
   '
7  Call MENSERR(24, "Documento no Registrado")
   GoTo 5
   '
8  TPSP$ = ""
   Screen.MousePointer = 11
   While L& > 1
     DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
     If DCC& < DOCUNU& Then GoTo 9
     L& = L& - 1
   Wend
   '
9  CAPAGINAS% = 0
   For KKL& = L& To URE&
     TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
     If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
       TPSP$ = TPSP$ + Mid$(TBUF$, 5)
     End If
   Next KKL&
   '
95 SPOOLSTRING = TPSP$
   '
10 PPXX% = InStr(TPSP$, Chr$(255))
   If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
   End If
   '
   Screen.MousePointer = 1
   PAGINACTIVA% = 0
   ANCHPAGINA = EDIFORMU.Picture1(PAGINACTIVA%).Width
   EDIFORMU.Picture1(PAGINACTIVA%).Refresh
   EDIFORMU.VScroll1.Min = 0
   EDIFORMU.VScroll1.Value = 0
   EDIFORMU.VScroll1.Max = 1
   If ALTUPAGINA > EDIFORMU.Frame1.Height Then
      EDIFORMU.VScroll1.Max = (ALTUPAGINA - EDIFORMU.Frame1.Height) / 100
   End If
   EDIFORMU.HScroll1.Min = 0
   EDIFORMU.HScroll1.Value = 0
   EDIFORMU.HScroll1.Max = 1
   If ANCHPAGINA > EDIFORMU.Frame1.Width Then
      EDIFORMU.HScroll1.Max = (ANCHPAGINA - EDIFORMU.Frame1.Width) / 100
   End If
   '
   EDIFORMU.Show
   '
End Sub
'
Private Sub Form_Load()
    '
    Show
    Refresh
    '   MsgBox LUDAT$
    Call VERJOBID
    COMMB$ = Command$
    PPX% = InStr(COMMB$, "#")
    If PPX% > 0 Then
        COMMB$ = Left$(COMMB$, PPX% - 1)
    End If
    '
    Call CARLISFUEN
    Call CARINDIFORM
    CODFOR$ = COMMB$
    Text1.TEXT = CODFOR$
    Text2.TEXT = ECOARCH$("INDIFRM", CODFOR$)
    '
    ISOLOGO$ = TRIM$(CONTROL$("", "ISOLOGO"))
    If ISOLOGO$ <> "" Then
       Text3.TEXT = ISOLOGO$
    End If
    '
    RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
    If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
    If RUTARGRA$ <> "" Then
       Text4.TEXT = RUTARGRA$
    End If
    '
End Sub

Private Sub Text1_DblClick()
    '
    Call SELECHO("INDIFRM")
    CODFOR$ = VALACT1$("INDIFRM")
    If TRIM$(CODFOR$) = "" Then
        Text1.SetFocus
      Else
        Text1.TEXT = CODFOR$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
    '
End Sub

Private Sub Text2_LOSTFOCUS()
    Call GRATITUFORM
End Sub

Sub CARLISFUEN()
    '
    NOMIMPRE$ = TRIM$(Printer.DeviceName)
    If NOMIMPRE$ = "" Then
         Call MENSERR(24, "Falta Seleccionar Impresora\en la Configuración de Windows")
         Call FINAL("")
    End If
    '
    Call BLANARCH("FUENPRIN")
    For I% = 1 To Printer.FontCount
        Printer.FontName = Printer.Fonts(I% - 1)
        List1.AddItem Printer.FontName
    Next I%
    '
    For I% = 1 To List1.ListCount
        PRINFORM.List1.ListIndex = I% - 1
        YY$ = TRIM$(PRINFORM.List1.TEXT)
        If YY$ <> "" Then
            XX$ = REGBLAN$("FUENPRIN")
            Call REPLA(XX$, YY$, 1, 42)
            Printer.FontName = YY$
            A1 = 0: A2 = 0
            On Error Resume Next
            A1 = Printer.TextWidth(String$(80, "l"))
            A2 = Printer.TextWidth(String$(80, "M"))
            On Error GoTo 0
            If A1 > 0 And A2 > 0 Then
              If A1 = A2 Then
                Call REPLA(XX$, "(fija)", 43, 6)
              End If
              Call REGAPP("FUENPRIN", XX$)
            End If
        End If
    Next I%
    '
    Call CIERRARCH("FUENPRIN")
    '
End Sub

Sub CARINDIFORM()
    '
    Screen.MousePointer = 11
    Dim RF As String * 128
    File1.Path = LUDAT$
    For U& = 0 To File1.ListCount - 1
        File1.ListIndex = U&
        ARXYZ$ = UCase$(File1.filename)
        PPXX% = InStr(ARXYZ$, ".FRM")
        If PPXX% > 1 Then
           ARXYZ$ = TRIM$(Left$(ARXYZ$, PPXX% - 1))
           If ARXYZ$ <> "" Then
               NX% = FILEOPEN%(LU$ + ARXYZ$ + ".FRM", 0, 128)
               TITUFOR$ = ""
               For I% = 1 To LOF(NX%) / 128
                   Get #NX%, I%, RF$
                   If TRIM$(Left$(RF$, 8)) = "TITUFORM" Then
                       TITUFOR$ = Mid$(RF$, 9, 48)
                       GoTo 10
                   End If
               Next I%
10             Close #NX%: NX% = 0
               '
               For V& = 1 To ULTREG&("INDIFRM")
                   INF$ = REGLEIDO$("INDIFRM", V&)
                   If Left$(INF$, 8) = AJUSTI$(ARXYZ$, 8) Then
                       If TITUFOR$ <> "" Then
                           Call REPLA(INF$, TITUFOR$, 9, 48)
                           Call GRAREG("INDIFRM", INF$, V&)
                       End If
                       GoTo 19
                   End If
               Next V&
               '
               INF$ = REGBLAN$("INDIFRM")
               Call REPLA(INF$, ARXYZ$, 1, 8)
               Call REPLA(INF$, TITUFOR$, 9, 48)
               Call REGAPP("INDIFRM", INF$)
               '
           End If
        End If
19  Next U&
    '
    Call CIERRARCH("INDIFRM")
    Call FRESHECHO("INDIFRM")
    Screen.MousePointer = 1
    '
End Sub
'

