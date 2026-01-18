VERSION 5.00
Begin VB.Form ANOTABOL 
   BackColor       =   &H00EEDDEE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anotador de Boletas de Recepción"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00EEDDEE&
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   105
      TabIndex        =   20
      Top             =   210
      Width           =   7680
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.B.R.:"
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
         Left            =   210
         TabIndex        =   29
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label16 
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
         Left            =   1260
         TabIndex        =   28
         Top             =   380
         Width           =   855
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.Remito:"
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
         TabIndex        =   27
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label Label13 
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
         Left            =   3570
         TabIndex        =   26
         Top             =   375
         Width           =   1800
      End
      Begin VB.Label Label8 
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
         Left            =   5355
         TabIndex        =   25
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label7 
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
         Left            =   6405
         TabIndex        =   24
         Top             =   375
         Width           =   1065
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
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
         Left            =   210
         TabIndex        =   23
         Top             =   780
         Width           =   960
      End
      Begin VB.Label Label11 
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
         Left            =   1260
         TabIndex        =   22
         Top             =   735
         Width           =   6210
      End
      Begin VB.Label Label14 
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
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEDDEE&
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
      Height          =   2060
      Left            =   105
      TabIndex        =   5
      Top             =   5985
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
         Height          =   1275
         Left            =   1155
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
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
         Height          =   855
         Left            =   8360
         Picture         =   "ANOTABOL.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1140
         Width           =   750
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0C0&
         Height          =   425
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   10515
         TabIndex        =   7
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
            TabIndex        =   9
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
            TabIndex        =   8
            Top             =   75
            Width           =   10200
         End
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
         TabIndex        =   18
         Top             =   840
         Width           =   4000
      End
      Begin VB.Line Line7 
         X1              =   10580
         X2              =   10580
         Y1              =   315
         Y2              =   1995
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
         TabIndex        =   17
         Top             =   1365
         Width           =   960
      End
      Begin VB.Line Line5 
         X1              =   8295
         X2              =   8295
         Y1              =   735
         Y2              =   1995
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   10605
         Y1              =   2010
         Y2              =   2010
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
         Height          =   1275
         Left            =   0
         TabIndex        =   10
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
         Height          =   1280
         Left            =   8310
         TabIndex        =   11
         Top             =   735
         Width           =   2260
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEDDEE&
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
      TabIndex        =   3
      Top             =   1470
      Width           =   10620
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   3795
         Left            =   10290
         TabIndex        =   19
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   4005
         Left            =   8085
         TabIndex        =   15
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   4110
         Left            =   1155
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   735
         Width           =   10580
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   10515
         TabIndex        =   4
         Top             =   315
         Width           =   10580
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   Usuario/Terminal"
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
            Left            =   80
            TabIndex        =   16
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
   Begin VB.CommandButton Command2 
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
      Height          =   1110
      Left            =   9030
      Picture         =   "ANOTABOL.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   315
      Width           =   800
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
      Height          =   1110
      Left            =   9920
      Picture         =   "ANOTABOL.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Sale de la Aplicación"
      Top             =   315
      Width           =   800
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   7420
      Picture         =   "ANOTABOL.frx":075E
      Top             =   900
      Width           =   2010
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
      TabIndex        =   2
      Top             =   2625
      Width           =   1695
   End
End
Attribute VB_Name = "ANOTABOL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command12_Click()

End Sub

Private Sub Command26_Click()

End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   NPX& = Val(Label16)
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
         Call REGAPP("ANOTABOL", Left$(MKS$(NPX&) + TTXX$, 96))
       Next KKI%
       Call REGAPP("ANOTABOL", MKS$(NPX&) + String$(88, "="))
       Call CIERRARCH("ANOTABOL")
       Text1 = ""
       '
       List1.Clear
       RFF$ = RECFILT$("ANOTABOL", 1, MKS$(NPX&))
       '
       For U& = 1 To Len(RFF$) Step 4
         RENOTA& = CVS(Mid$(RFF$, U&, 4))
         XX$ = REGLEIDO$("ANOTABOL", RENOTA&)
         List1.AddItem Mid$(XX$, 5, 92)
       Next U&
       '
     End If
   End If
   Command3.Enabled = True
End Sub

Private Sub Command9_Click()
   Label7 = ""
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub Form_Load()
    '
    If Screen.Height > 9000 Then
       Top = (Screen.Height - Height) / 2
       Left = (Screen.Width - Width) / 2
    End If
    '
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

Private Sub Label16_Change()
   '
   NBR& = Val(Label16)
   List1.Clear
   If NBR& > 0 Then
     RFF$ = RECFILT$("BOLRECEP", 1, MKS$(NBR&))
     For U& = 1 To Len(RFF$) Step 4
       REBO& = CVS(Mid$(RFF$, U&, 4))
       XX$ = REGLEIDO$("BOLRECEP", REBO&)
       '
       If List1.ListCount < 1 Then
         NPRO% = CVI(Mid$(XX$, 105, 2))
         Label14 = TRIM$(Str$(NPRO%))
         Label11 = RASOCLI$((-1) * NPRO%)
         Label13 = TRIM$(Mid$(XX$, 119, 16)) ' NUMERO DE REMITO
         FEX = CVI(Mid$(XX$, 81, 2))
         Label7 = FECHATEX$(FEX)
       End If
     Next U&
   End If
   '
20 RFF$ = RECFILT$("ANOTABOL", 1, MKS$(NBR&))
   For U& = 1 To Len(RFF$) Step 4
     RENOTA& = CVS(Mid$(RFF$, U&, 4))
     XX$ = REGLEIDO$("ANOTABOL", RENOTA&)
     List1.AddItem Mid$(XX$, 5, 92)
   Next U&
   '
End Sub

