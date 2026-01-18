VERSION 5.00
Begin VB.Form ANOTAORF 
   BackColor       =   &H00E0FFD0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anotador Ordenes de Fabricación"
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
   Begin VB.CommandButton Command6 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   9870
      Picture         =   "ANOTAORF.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Consultas y Re-Impresión"
      Top             =   1020
      Width           =   750
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Orden de Fabricación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
         Left            =   2520
         TabIndex        =   18
         Top             =   315
         Width           =   180
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1155
         TabIndex        =   31
         ToolTipText     =   "Doble Click para Restaurar Situación de Preparacion y Despacho"
         Top             =   1050
         Width           =   330
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
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
         Left            =   -210
         TabIndex        =   30
         Top             =   1110
         Width           =   1275
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5460
         TabIndex        =   29
         Top             =   300
         Width           =   1170
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant:"
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
         Left            =   4095
         TabIndex        =   28
         Top             =   345
         Width           =   1275
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Left            =   1890
         TabIndex        =   24
         Top             =   735
         Width           =   7485
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo:"
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
         Left            =   6510
         TabIndex        =   23
         Top             =   345
         Width           =   960
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7560
         TabIndex        =   22
         Top             =   300
         Width           =   1815
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
         Left            =   2205
         TabIndex        =   21
         Top             =   375
         Width           =   1275
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1155
         TabIndex        =   20
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3570
         TabIndex        =   19
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
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
         Index           =   1
         Left            =   -105
         TabIndex        =   17
         Top             =   735
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
         Picture         =   "ANOTAORF.frx":030A
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
         TabIndex        =   27
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   9345
         TabIndex        =   26
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
         TabIndex        =   33
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
            TabIndex        =   25
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   9870
      Picture         =   "ANOTAORF.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Sale de la Aplicación"
      Top             =   300
      Width           =   750
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
Attribute VB_Name = "ANOTAORF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command3_Click()
   Command3.Enabled = False
   NPX& = Val(Mid$(Label7, 4, 6))
   If NPX& > 0 Then
     If TRIM$(Text1) <> "" Then
       Call FRATEXTO(Text1, 56)
       For KKI% = 1 To CARETEX%
         TTXX$ = Space$(76)
         Call REPLA(TTXX$, RETEX$(KKI%), 11, 56)
         If KKI% = 1 Then
           Call REPLA(TTXX$, HOS$, 1, 8)
           Call REPLA(TTXX$, USERTER$, 69, 24)
         End If
         Call REGAPP("ANOTAORF", Left$(MKS$(NPX&) + TTXX$, 96))
       Next KKI%
       Call REGAPP("ANOTAORF", MKS$(NPX&) + String$(88, "="))
       Call CIERRARCH("ANOTAORF")
       Text1 = ""
       '
       List1.Clear
       RFF$ = RECFILT$("ANOTAORF", 1, MKS$(NPX&))
       '
       For U& = 1 To Len(RFF$) Step 4
         RENOTA& = CVS(Mid$(RFF$, U&, 4))
         XX$ = REGLEIDO$("ANOTAORF", RENOTA&)
         List1.AddItem Mid$(XX$, 5, 92)
       Next U&
       '
     End If
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
3   Call SELECHO("ECORDEF/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
    NORFA$ = TRIM$(VALACT1$("ECORDEF"))
    If NORFA$ = "" Then GoTo 99
    '
    Label7 = TRIM$(NORFA$)
    '
99  Command4.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call OFABRIC.VERORFA(Label7)
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
End Sub


Private Sub Label7_Change()
     List1.Clear
     If TRIM$(Label7) = "" Then Exit Sub
     '
     NORFA$ = AJUSTI$(Label7, 12)
     If TRIM$(ECOARCH$("ECORDEF", NORFA$)) = "" Then
        Call MENSERR(24, "Orden no Registrada")
        Exit Sub
     End If
     '
     For I& = ULTREG&("ORDEFABR") To 1 Step -1
       XX$ = REGLEIDO$("ORDEFABR", I&)
       If Mid$(XX$, 69, 12) = NORFA$ Then
         FEGEN = CVI(Mid$(XX$, 81, 2))
         CATOT# = CVD(Mid$(XX$, 89, 8))
         STAT% = Asc(Mid$(XX$, 108, 1))
         ' CAPEN# = CATOT# - CAPRA#
         ' If CAPEN# < 0 Then CAPEN# = 0
         CIXI% = CVI(Mid$(XX$, 83, 2))
         GoTo 1420
       End If
     Next I&
     '
     Call COMUNI("Inconsistencia en Base de Datos\de Ordenes de Fabricación.\  \Consulte al Soporte de Sistemas.")
     INDI% = 0
     Exit Sub
     '
1420 Label8 = FECHATEX$(FEGEN)
     Label16 = FORMATNUM(CATOT#, "F8.1")
     Label11 = CODEXT$(CIXI%)
     Label14 = DESCRIT$(CIXI%)
     Label18 = TRIM$(Str$(STAT%))
     '
20   NPX& = Val(Mid$(NORFA$, 4, 6))
     RFF$ = RECFILT$("ANOTAORF", 1, MKS$(NPX&))
     For U& = 1 To Len(RFF$) Step 4
       RENOTA& = CVS(Mid$(RFF$, U&, 4))
       XX$ = REGLEIDO$("ANOTAORF", RENOTA&)
       List1.AddItem Mid$(XX$, 5, 92)
     Next U&
     '
End Sub
