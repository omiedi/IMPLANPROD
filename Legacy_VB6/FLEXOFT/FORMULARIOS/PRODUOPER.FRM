VERSION 5.00
Begin VB.Form PRODUOPER 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Productividad de Mano de Obra"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   ClipControls    =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox SORTLIST 
      Height          =   645
      Left            =   105
      Sorted          =   -1  'True
      TabIndex        =   34
      Top             =   7035
      Visible         =   0   'False
      Width           =   11580
   End
   Begin VB.CommandButton Command2 
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
      Left            =   11025
      Picture         =   "PRODUOPER.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Impresión - Vuelco de Pantalla"
      Top             =   1155
      Width           =   750
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5100
      Left            =   105
      TabIndex        =   27
      Top             =   2620
      Width           =   11700
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Resumen de Productividad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   105
      TabIndex        =   10
      Top             =   210
      Width           =   6840
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
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
         Index           =   7
         Left            =   4515
         TabIndex        =   26
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   5565
         TabIndex        =   25
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rendim. Personal: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   6
         Left            =   4830
         TabIndex        =   24
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5775
         TabIndex        =   23
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Index           =   2
         Left            =   3620
         TabIndex        =   22
         Top             =   1260
         Width           =   1250
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Index           =   1
         Left            =   3620
         TabIndex        =   21
         Top             =   840
         Width           =   1250
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Difer.en Tiempo:"
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
         Index           =   5
         Left            =   1350
         TabIndex        =   20
         Top             =   1320
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo Tot.Empl:"
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
         Index           =   4
         Left            =   1350
         TabIndex        =   19
         Top             =   900
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo Tot.Estim:"
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
         Left            =   1770
         TabIndex        =   18
         Top             =   480
         Width           =   1800
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Index           =   0
         Left            =   3620
         TabIndex        =   17
         Top             =   420
         Width           =   1250
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Index           =   2
         Left            =   1260
         TabIndex        =   16
         Top             =   1260
         Width           =   700
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Negativas:"
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
         Left            =   -1050
         TabIndex        =   15
         Top             =   1320
         Width           =   2220
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Index           =   1
         Left            =   1260
         TabIndex        =   14
         Top             =   840
         Width           =   700
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Positivas:"
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
         Left            =   -1050
         TabIndex        =   13
         Top             =   900
         Width           =   2220
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Index           =   0
         Left            =   1260
         TabIndex        =   12
         Top             =   420
         Width           =   700
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tot.Ordenes:"
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
         Left            =   -1050
         TabIndex        =   11
         Top             =   480
         Width           =   2220
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   120
      Left            =   210
      TabIndex        =   9
      Top             =   1575
      Width           =   15
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Operario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7035
      TabIndex        =   1
      Top             =   210
      Width           =   3860
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cat:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   8
         Left            =   2940
         TabIndex        =   33
         Top             =   525
         Width           =   750
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
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
         Left            =   3360
         TabIndex        =   32
         Top             =   420
         Width           =   330
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000016&
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
         Left            =   210
         TabIndex        =   31
         Top             =   420
         Width           =   2535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
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
         Left            =   2625
         TabIndex        =   8
         Top             =   1260
         Width           =   750
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1260
         TabIndex        =   7
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
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
         Left            =   1260
         TabIndex        =   6
         Top             =   840
         Width           =   2115
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   2370
         TabIndex        =   5
         Top             =   1215
         Width           =   225
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Hs"
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
         Index           =   12
         Left            =   3405
         TabIndex        =   4
         Top             =   1365
         Width           =   1590
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Generado el:"
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
         Index           =   13
         Left            =   105
         TabIndex        =   3
         Top             =   1365
         Width           =   2220
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Período:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   10
         Left            =   490
         TabIndex        =   2
         Top             =   945
         Width           =   2220
      End
   End
   Begin VB.CommandButton Command1 
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
      Left            =   11025
      Picture         =   "PRODUOPER.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cerrar - Salir de la Aplicación"
      Top             =   315
      Width           =   750
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   5145
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fecha  Comprob Tarea           Modelo        Oper  Ped  Ent x Uni  x Uni  x Uni    Total  Rend"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   150
      TabIndex        =   29
      Top             =   2355
      Width           =   11565
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                                             Ctro  Can  Can T.Est  T.Emp  Difer    Difer"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   105
      TabIndex        =   28
      Top             =   2100
      Width           =   11700
   End
End
Attribute VB_Name = "PRODUOPER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Unload PRODUOPER
End Sub

Private Sub Command2_Click()
   '
   Command2.Enabled = False
   FORIPRE$ = TRIM$(CONTROL$("", "FORPROIN"))
   If FORIPRE$ = "" Then
        PRINTFORM
      Else
        Screen.MousePointer = 11
        FEX = HOY(HOS$)
        FECHDOC$ = HOS$
        NUMEDOC$ = ""
        TIPODOC$ = "PRODUCTIvidad Individual"
        '
        CODPARAM$(1) = "FECH-LAN"
        CODPARAM$(2) = "FECH-VEN"
        CODPARAM$(3) = "CUEN-CLI"
        CODPARAM$(4) = "RASO-CLI"
        CODPARAM$(5) = "CIVA-CLI"
        CODPARAM$(6) = "PRE-LIST"
        CODPARAM$(7) = "TOTA-CAN"
        CODPARAM$(8) = "CANT-DIS"
        CODPARAM$(9) = "CANT-FAL"
        CODPARAM$(10) = "TIE-PREP"
        CODPARAM$(11) = "TIE-PROD"
        CODPARAM$(12) = "TIE-TOTA"
        CODPARAM$(13) = "CAKILOS"
        CAPARDOC% = 13
        '
        DOCPARAM$(1) = TRIM$(Left$(Label1.Caption, 8))
        DOCPARAM$(2) = TRIM$(Mid$(Label1.Caption, 12, 8))
        TTX0$ = PRODUCTI.List1.TEXT
        DOCPARAM$(3) = TRIM$(Mid$(TTX0$, 123, 6))
        DOCPARAM$(4) = TRIM$(Label12.Caption)
        DOCPARAM$(5) = TRIM$(Label13.Caption)
        DOCPARAM$(6) = ""
        DOCPARAM$(7) = TRIM$(Label6(0).Caption)
        DOCPARAM$(8) = TRIM$(Label6(1).Caption)
        DOCPARAM$(9) = TRIM$(Label6(2).Caption)
        DOCPARAM$(10) = AJUSTD$(Label7(0).Caption, 12)
        DOCPARAM$(11) = AJUSTD$(Label7(1).Caption, 12)
        DOCPARAM$(12) = AJUSTD$(Label7(2).Caption, 12)
        DOCPARAM$(13) = TRIM$(Label8.Caption)
        '
        CODTABU1$(1) = "REF-MAT1"
        CACOLTAB1% = 1
        '
        CAITAB1% = 0
        For U& = 1 To List1.ListCount
           Z$ = List1.List(U& - 1)
           CAITAB1% = CAITAB1% + 1
           TETABU1$(1, CAITAB1%) = Z$
        Next U&
        '
        Call PRINTDOC("FORPROIN")
        Screen.MousePointer = 1
        '
      End If
      '
   Command2.Enabled = True
   '
End Sub

Private Sub Form_Load()
   TTX0$ = PRODUCTI.List1.TEXT
   LEGA% = Val(Mid$(TTX0$, 123, 6))
   '
   Label1.Caption = PRODUCTI.Label1.Caption
   Label3.Caption = PRODUCTI.Label3.Caption
   Label4.Caption = PRODUCTI.Label4.Caption
   '
   Label12.Caption = ECOARCH$("PADRON", MKI$(LEGA%))
   Label13.Caption = TRIM$(Mid$(TTX0$, 23, 1))
   Label6(0) = TRIM$(Mid$(TTX0$, 24, 6))
   Label6(1) = TRIM$(Mid$(TTX0$, 30, 5))
   Label6(2) = TRIM$(Mid$(TTX0$, 35, 5))
   Label7(0) = TRIM$(Mid$(TTX0$, 40, 12))
   Label7(1) = TRIM$(Mid$(TTX0$, 52, 12))
   Label7(2) = TRIM$(Mid$(TTX0$, 67, 10))
   Label7(2).ForeColor = RGB(0, 0, 0)
   If Mid$(TTX0$, 66, 1) = "-" Then Label7(2).ForeColor = RGB(255, 0, 0)
   Label8.Caption = TRIM$(Mid$(TTX0$, 77, 8))
   Label8.ForeColor = RGB(0, 0, 0)
   If Val(Label8) < 100 Then Label8.ForeColor = RGB(255, 0, 0)
   '
   Screen.MousePointer = 11
   SORTLIST.Clear
   List1.Clear
   VALODIF = 0
   RFF$ = RECFILT$("COMOBRA", 1, MKI$(LEGA%))
   For KL& = 1 To Len(RFF$) Step 4
     TTXX$ = Space$(128)
     RECO& = CVS(Mid$(RFF$, KL&, 4))
     XX$ = REGLEIDO$("COMOBRA", RECO&)
     VALODIF = VALODIF + CVS(Mid$(XX$, 115, 4))
     FEX = CVI(Mid$(XX$, 75, 2))
     SORFE$ = CSTRING$(MKS$(FEX), 3, 6, 0, 2)
     Call REPLA(TTXX$, Left$(FECHATEX$(FEX), 5), 1, 5)
     NORFOP$ = Mid$(XX$, 6, 6) + "-" + TRIM$(Str$(CVI(Mid$(XX$, 17, 2))))
     If Left$(NORFOP$, 3) = "000" Then Mid$(NORFOP$, 1, 3) = "   "
     If Left$(NORFOP$, 2) = "00" Then Mid$(NORFOP$, 1, 2) = "  "
     If Left$(NORFOP$, 1) = "0" Then Mid$(NORFOP$, 1, 1) = " "
     Call REPLA(TTXX$, NORFOP$, 7, 10)
     TAREA$ = Mid$(XX$, 19, 48)
     Call REPLA(TTXX$, TAREA$, 18, 17)
     CIXI% = CVI(Mid$(XX$, 15, 2))
     Call REPLA(TTXX$, CODEXT$(CIXI%), 36, 15)
     MAQUI$ = Mid$(XX$, 101, 6)
     Call REPLA(TTXX$, MAQUI$, 52, 6)
     CAPED = CVD(Mid$(XX$, 107, 8))
     Call REPLA(TTXX$, CSTRING$(MKS$(CAPED), 3, 5, 0, 2), 58, 5)
     CAREA = CVD(Mid$(XX$, 67, 8))
     Call REPLA(TTXX$, CSTRING$(MKS$(CAREA), 3, 5, 0, 2), 63, 5)
     TEMPLE = CVS(Mid$(XX$, 81, 4))
     TEMPLEUNI = TEMPLE: If CAREA > 0.5 Then TEMPLEUNI = TEMPLE / CAREA
     TSTAN = CVS(Mid$(XX$, 85, 4))
     TSTANUNI = TSTAN: If CAREA > 0.5 Then TSTANUNI = TSTAN / CAREA
     DIFEUNI = TSTANUNI - TEMPLEUNI
     DIFETOT = TSTAN - TEMPLE
     TESTUNI$ = Mid$(HORMINSE$(TSTANUNI), 3): If Left$(TESTUNI$, 2) = "0:" Then Mid$(TESTUNI$, 1, 2) = "  "
     Call REPLA(TTXX$, TESTUNI$, 68, 7)
     TEMPUNI$ = Mid$(HORMINSE$(TEMPLEUNI), 3): If Left$(TEMPUNI$, 2) = "0:" Then Mid$(TEMPUNI$, 1, 2) = "  "
     Call REPLA(TTXX$, TEMPUNI$, 76, 7)
     DIFUNI$ = Mid$(HORMINSE$(Abs(DIFEUNI)), 3): If Left$(DIFUNI$, 2) = "0:" Then Mid$(DIFUNI$, 1, 2) = "  "
     If DIFEUNI < 0 Then DIFUNI$ = DIFUNI$ + "-" Else DIFUNI$ = DIFUNI$ + " "
     Call REPLA(TTXX$, DIFUNI$, 84, 8)
     DIFTOT$ = HORMINSE$(Abs(DIFETOT))
     If DIFETOT < 0 Then DIFTOT$ = DIFTOT$ + "-" Else DIFTOT$ = DIFTOT$ + " "
     Call REPLA(TTXX$, DIFTOT$, 92, 10)
     RENDIM = 999: If TEMPLE > 0.01 Then RENDIM = 100 * TSTAN / TEMPLE
     If RENDIM > 999 Then RENDIM = 999
     RDXX$ = CSTRING$(MKS$(RENDIM), 3, 6, 1, 2)
     If RENDIM >= 998.9 Then RDXX$ = " *****"
     Call REPLA(TTXX$, RDXX$, 102, 6)
     SORTLIST.AddItem SORFE$ + TTXX$
   Next KL&
   '
   Label9.ForeColor = RGB(0, 0, 0)
   If VALODIF < 0 Then Label9.ForeColor = RGB(255, 0, 0)
   Label9.Caption = TRIM$(CSTRING$(MKS$(Abs(VALODIF)), 3, 10, 2, 2))
   '
   For KL& = 1 To SORTLIST.ListCount
     List1.AddItem Mid$(SORTLIST.List(KL& - 1), 7)
   Next KL&
   '
   Screen.MousePointer = 1
   
End Sub

Private Sub Label12_DblClick()
   Unload PRODUOPER
End Sub

Private Sub List1_DblClick()
   '
10 For KIL& = 1 To List1.ListCount - 1
     AAA$ = Mid$(List1.List(KIL& - 1), 36, 22)
     BBB$ = Mid$(List1.List(KIL&), 36, 22)
     If AAA$ > BBB$ Then
       CCC$ = List1.List(KIL& - 1)
       List1.RemoveItem (KIL& - 1)
       List1.AddItem CCC$
       GoTo 10
     End If
   Next KIL&
   '
End Sub
