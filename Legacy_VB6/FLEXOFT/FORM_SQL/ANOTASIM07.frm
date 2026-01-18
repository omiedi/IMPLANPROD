VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ANOTASIM07 
   BackColor       =   &H00E0FFD0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anotador Solicitud Interna de Materiales"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   34
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command6 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   9870
      Picture         =   "ANOTASIM07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Consultas y Re-Impresión"
      Top             =   1020
      Width           =   750
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Solicitud Interna de Materiales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1550
      Left            =   105
      TabIndex        =   15
      Top             =   210
      Width           =   9570
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   18
         Top             =   315
         Width           =   180
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sector:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3720
         TabIndex        =   33
         Top             =   345
         Width           =   1275
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         TabIndex        =   32
         Top             =   315
         Width           =   945
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   960
         TabIndex        =   29
         ToolTipText     =   "Doble Click para Restaurar Situación de Preparacion y Despacho"
         Top             =   1050
         Width           =   570
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "User:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -480
         TabIndex        =   28
         Top             =   1110
         Width           =   1275
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5040
         TabIndex        =   27
         Top             =   300
         Width           =   1410
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Left            =   1770
         TabIndex        =   23
         Top             =   735
         Width           =   7575
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Destino:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6390
         TabIndex        =   22
         Top             =   345
         Width           =   960
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7440
         TabIndex        =   21
         Top             =   300
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Left            =   1680
         TabIndex        =   20
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3000
         TabIndex        =   19
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   -320
         TabIndex        =   17
         Top             =   720
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   -840
         TabIndex        =   16
         Top             =   380
         Width           =   1905
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Novedades"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   105
      TabIndex        =   4
      Top             =   6300
      Width           =   10620
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1175
         Left            =   1155
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   735
         Width           =   7155
      End
      Begin VB.CommandButton Command3 
         Caption         =   "O.K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   745
         Left            =   8360
         Picture         =   "ANOTASIM07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1140
         Width           =   750
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0C0&
         Height          =   425
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   10515
         TabIndex        =   6
         Top             =   315
         Width           =   10580
         Begin VB.Line Line2 
            X1              =   0
            X2              =   1155
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Label Label5 
            Caption         =   "Label5"
            Height          =   15
            Left            =   0
            TabIndex        =   8
            Top             =   420
            Width           =   1380
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   7
            Top             =   75
            Width           =   10200
         End
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   10665
         Y1              =   1900
         Y2              =   1900
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
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
         Left            =   8400
         TabIndex        =   26
         Top             =   840
         Width           =   4000
      End
      Begin VB.Line Line7 
         X1              =   10580
         X2              =   10580
         Y1              =   315
         Y2              =   1900
      End
      Begin VB.Line Line6 
         X1              =   8295
         X2              =   10605
         Y1              =   1130
         Y2              =   1130
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Registrar Anotación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   9345
         TabIndex        =   25
         Top             =   1310
         Width           =   960
      End
      Begin VB.Line Line5 
         X1              =   8295
         X2              =   8295
         Y1              =   735
         Y2              =   1995
      End
      Begin VB.Label Label9 
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
         Height          =   1165
         Left            =   0
         TabIndex        =   9
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label10 
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
         Height          =   1170
         Left            =   8310
         TabIndex        =   10
         Top             =   735
         Width           =   2260
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Anotaciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4450
      Left            =   105
      TabIndex        =   2
      Top             =   1800
      Width           =   10620
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   3795
         Left            =   10290
         TabIndex        =   31
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   4005
         Left            =   8085
         TabIndex        =   14
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   4110
         Left            =   1155
         TabIndex        =   13
         Top             =   630
         Width           =   15
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
         Height          =   3660
         Left            =   0
         TabIndex        =   12
         Top             =   735
         Width           =   10580
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   10515
         TabIndex        =   3
         Top             =   315
         Width           =   10580
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   24
            Top             =   105
            Width           =   10200
         End
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   10605
         Y1              =   4410
         Y2              =   4410
      End
      Begin VB.Line Line1 
         X1              =   10580
         X2              =   10580
         Y1              =   320
         Y2              =   6225
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   9870
      Picture         =   "ANOTASIM07.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Sale de la Aplicación"
      Top             =   360
      Width           =   750
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "ANOTASIM07.frx":075E
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "ANOTASIM07.frx":0992
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -105
      TabIndex        =   1
      Top             =   2625
      Width           =   1695
   End
End
Attribute VB_Name = "ANOTASIM07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command3_Click()
   '
   Command3.Enabled = False
   Call ANOTASIM(XVALO(Label7), Text1)
   Text1 = ""
   Call Label7_Change
   Command3.Enabled = True
   '
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
10  If ULTREG&("INSOLINT") > 0 Then ' SI TIENE INFORMACION MUESTRA EL INDICE
        Call SELECHO("INSOLINT/Indice General de Solicitud de Materiales")
        Label7.Caption = VALACT1$("INSOLINT")
    Else
        GoTo 99
    End If
99  Command4.Enabled = True
End Sub

Private Sub Command6_Click()
    '
    Command6.Enabled = False
    '
10  If ULTREG&("!INSOLINT") > 0 Then ' SI TIENE INFORMACION MUESTRA EL INDICE
    Call SELECHO("!INSOLINT/Indice General de Solicitud de Materiales")
    Else
     GoTo 99
    End If
    '
    NUDOCU$ = VALACT1$("!INSOLINT") ' ASIGNO EL NUMERO DEL FORMULARIO IMPRESO  SELECCIONAD0
    TIDOCUM$ = "Solicitud de Materiales"
    If NUDOCU$ = "" Then GoTo 99 ' por si no elige ninguno y cierra el indice
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
         XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
         If InStr(XX$, TIDOCUM$) > 4 Then
           PPXX% = InStr(XX$, "Nro.")
           If PPXX% > 0 Then
             XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
           End If
           If InStr(XX$, NUDOCU$) > 4 Then
             DOCUNU& = CVS(Left$(XX$, 4))
             GoTo 5
           End If
         End If
    Next UI&
       '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a Esta Solicitud de Material.\ ")
    GoTo 10
    
5   Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
    '
    GoTo 10
    '
99  Command6.Enabled = True

End Sub

Private Sub Command9_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub Form_Load()
    
    Call CENTRAFORM(Me)
    '
    DoEvents
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Screen.MousePointer = 1
    HOII% = HOY(HOS$)
    Label9 = " " + HOS$
    Label12 = Left$(USERNAME, 17)
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
End Sub


Private Sub Label7_Change()
    '
    If Label7.Caption = "" Then Exit Sub
    '
    CONDI$ = "NUME_SOLINT = " & XVALO(Label7.Caption)
    SECTO = VALOBADA("ENCASOLINT", "SECT_SOLINT", CONDI$)
    Label16.Caption = SECTO
    FEXX = VALOBADA("ENCASOLINT", "FEMI_SOLINT", CONDI$)
    Label8.Caption = (FEXX)
    CUECO$ = VALOBADA("DETASOLINT", "ICONTA_SOLINT", CONDI$)
    Label11.Caption = CUECO$
    USSER = XVALO(VALOBADA("ENCASOLINT", "USUA_GEN", CONDI$))
    Label18 = USSER
    OBSER = VALOBADA("ENCASOLINT", "OBSE_SOLINT", CONDI$)
    Label14.Caption = OBSER
    '
    XX$ = VALOBADA("ENCASOLINT", "AnotaSolint", "NUME_SOLINT = " & Label7)
    LONGI& = Len(XX$)
    List1.Clear
    '
    INICIO& = 1
    Do While INICIO& < LONGI&
       FIN& = InStr(INICIO&, XX$, Chr$(10))
       If FIN& > 0 Then
          TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO& - 1)
          List1.AddItem TTXX$
          INICIO& = FIN& + 1
       Else
          TTXX$ = Mid$(XX$, INICIO&)
          List1.AddItem TTXX$
          Exit Do
       End If
    Loop
    Screen.MousePointer = 1
    '
End Sub

Sub ANOTASIM(NPX&, TANOTA$)
   '
   If NPX& > 0 Then
      If TRIM$(TANOTA$) <> "" Then
         Call FRATEXTO(TANOTA$, 56)
         '
         TTXX$ = ""
         For KKI% = 1 To CARETEX%
            TTX1$ = Space$(76)
        Call REPLA(TTX1$, RETEX$(KKI%), 11, 56)
            If KKI% = 1 Then
              Call REPLA(TTX1$, HOS$, 1, 8)
              Call REPLA(TTX1$, USERTER$, 69, 24)
            End If
            TTXX$ = TTXX$ + TTX1$ + vbCrLf$
         Next KKI%
         TTXX$ = TTXX$ + String$(88, "=") + vbCrLf$
         '
         ANOTANTE$ = VALOBADA("ENCASOLINT", "AnotaSolint", "Nume_Solint = " & NPX&)
         ANOTANUE$ = ANOTANTE$ + TTXX$
         Call ACTUREGISTRO("ENCASOLINT", "AnotaSolint", "Nume_Solint = " & NPX&, ANOTANUE$, REGAF&)
         '
      End If
   End If
   '
End Sub





