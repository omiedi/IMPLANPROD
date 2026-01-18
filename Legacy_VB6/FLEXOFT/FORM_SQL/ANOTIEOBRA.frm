VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ANOTIEOBRA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historial de Tiempos de Mano de Obra"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame7 
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
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   6360
      Width           =   10635
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00C0C0C0&
         Height          =   425
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   10275
         TabIndex        =   23
         Top             =   315
         Width           =   10330
         Begin VB.Label Label44 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   330
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   10320
         End
         Begin VB.Label Label23 
            Caption         =   "Label5"
            Height          =   15
            Left            =   0
            TabIndex        =   24
            Top             =   420
            Width           =   1380
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   1155
            Y1              =   420
            Y2              =   420
         End
      End
      Begin VB.TextBox Text14 
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
         Left            =   1230
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   720
         Width           =   6870
      End
      Begin VB.CommandButton Command30 
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
         Height          =   650
         Left            =   8240
         Picture         =   "ANOTIEOBRA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Agrega el Texo al Anotador"
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
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
         Height          =   225
         Left            =   8240
         TabIndex        =   28
         Top             =   840
         Width           =   2100
      End
      Begin VB.Line Line24 
         Index           =   1
         X1              =   0
         X2              =   1170
         Y1              =   1100
         Y2              =   1100
      End
      Begin VB.Line Line25 
         X1              =   8180
         X2              =   8180
         Y1              =   735
         Y2              =   1900
      End
      Begin VB.Label Label51 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Height          =   650
         Left            =   9120
         TabIndex        =   27
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Line Line24 
         Index           =   0
         X1              =   8240
         X2              =   10310
         Y1              =   1120
         Y2              =   1120
      End
      Begin VB.Line Line22 
         X1              =   0
         X2              =   10665
         Y1              =   1900
         Y2              =   1900
      End
      Begin VB.Label Label57 
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
         Index           =   0
         Left            =   0
         TabIndex        =   26
         Top             =   735
         Width           =   1170
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   315
      Index           =   0
      Left            =   8280
      ScaleHeight     =   255
      ScaleWidth      =   1110
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00004040&
      Height          =   8415
      Left            =   10830
      ScaleHeight     =   8355
      ScaleWidth      =   1140
      TabIndex        =   12
      Top             =   0
      Width           =   1200
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
         Height          =   650
         Left            =   105
         Picture         =   "ANOTIEOBRA.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Sale de la Aplicación"
         Top             =   210
         Width           =   750
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   705
         TabIndex        =   13
         Top             =   7440
         Width           =   770
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image3 
         Height          =   390
         Left            =   -220
         Picture         =   "ANOTIEOBRA.frx":0454
         Top             =   7770
         Width           =   1515
      End
   End
   Begin VB.Frame Frame6 
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
      Height          =   825
      Left            =   105
      TabIndex        =   6
      Top             =   0
      Width           =   10650
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
         Left            =   2100
         TabIndex        =   8
         Top             =   315
         Width           =   180
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000016&
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
         Left            =   1000
         TabIndex        =   29
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label Label81 
         BackColor       =   &H80000016&
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
         Left            =   9240
         TabIndex        =   19
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Alta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7560
         TabIndex        =   18
         Top             =   420
         Width           =   1635
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000016&
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
         Left            =   2280
         TabIndex        =   9
         Top             =   330
         Width           =   5580
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
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
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   10620
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   4935
         Index           =   0
         Left            =   10290
         TabIndex        =   11
         Top             =   330
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   4935
         Left            =   8085
         TabIndex        =   5
         Top             =   330
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   4935
         Left            =   1155
         TabIndex        =   4
         Top             =   330
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
         Height          =   4560
         Left            =   0
         TabIndex        =   3
         Top             =   735
         Width           =   10580
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   10515
         TabIndex        =   2
         Top             =   315
         Width           =   10580
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   10560
         End
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   10605
         Y1              =   4410
         Y2              =   4410
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "ANOTIEOBRA.frx":2376
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "ANOTIEOBRA.frx":25AA
      TabIndex        =   17
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
      TabIndex        =   0
      Top             =   2625
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "ANOTIEOBRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command30_Click()
   Command30.Enabled = False
   CI1% = CODINT%(Label8)
   If CI1% > 0 Then
     Call ANOTADATO_M_OBRA(CI1%, Text14)
     Call CARGAELLIST
     Text14 = ""
   End If
   
   Command30.Enabled = True
   
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     Label8 = COD$
   End If
99 Command4.Enabled = True
End Sub

Private Sub Command9_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub




Private Sub Form_Load()
    Call CENTRAFORM(Me)
    '
    Call APLICACIONSKINS(Me)

    DoEvents
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Screen.MousePointer = 1
    HOII% = HOY(HOS$)
    Label57(0) = " " + HOS$
    Label81 = HOS$
    Label45 = Left$(USERNAME, 17)
    '
End Sub

Sub ANOTADATO_M_OBRA(CI1%, TXTX$, Optional FF%, Optional MODO$)
    '
   If CI1% < 1 Then Exit Sub
   If TRIM$(TXTX$) = "" Then Exit Sub

   CONDI$ = "CODICONJ = " & CI1%
   If CantRegistros&("HISTOMOBRA", CONDI$, "NOMESS") < 1 Then
      Call CreaRegistro("HISTOMOBRA", "CODICONJ", CI1%, "NOMESS")
   End If
   '
   HOII% = HOY(HOS$)
   fech$ = HOS$
   If FF% > 0 Then
      fech$ = FECHATEX(FF%) ' SI TRAE FECHA EL PARAMETRO TOMO ESTA
   End If
   
   Call FRATEXTO(TXTX$, 56) ' FRACCIONO EL TEXTO
   '
   For KKI% = 1 To CARETEX% ' RECORRO DESDE 1 A LA CANTIDAD DE FRACCIONES
    TTYY$ = Space$(90)
    If KKI% = 1 Then
       Call REPLA(TTYY$, fech$, 1, 8) ' SI ES LA PRIMERA VEZ GRABA LA FECHA Y USUARIO
       Call REPLA(TTYY$, USERTER$, 69, 15)
    End If
    Call REPLA(TTYY$, RETEX$(KKI%), 11, 56) ' SIEMPRE QUE PASA GUARDA EL CENTRO DEL TEXTO
    ANOTANTE$ = ANOTANTE$ + TTYY$ + vbCrLf
   Next KKI%
   '
   ANOTANTE$ = ANOTANTE$ & String$(85, "=") & vbCrLf
   '
   '
   ANOTRIOR$ = VALOBADA("HISTOMOBRA", "ANOTMOBRA", CONDI$, "NOMESS")
   '
   ANOTA$ = ANOTRIOR$ & ANOTANTE$   ' CONCATENO LA ANTERIOR CON LA NUEVA
   Call ACTUREGISTRO("HISTOMOBRA", "ANOTMOBRA", CONDI$, ANOTA$, REGAF&, "NOMESS") 'GRABO
   '
End Sub



Private Sub Label8_Change()
  Label11 = DESCRIT$(CODINT%(Label8))
  Call CARGAELLIST
   '

End Sub
Sub CARGAELLIST()

   CI1% = CODINT%(Label8)
   If CI1% < 1 Then Exit Sub
   '
   CONDI$ = "CODICONJ = " & CI1%
   XX$ = VALOBADA("HISTOMOBRA", "ANOTMOBRA", CONDI$, "NOMESS")
   LONGI& = Len(XX$)
   '
   List1.Clear
   INICIO& = 1
   Do While INICIO& < LONGI&
        FIN& = InStr(INICIO&, XX$, Chr$(13))
        If FIN& > 0 Then
           TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO&)
           List1.AddItem TTXX$
           INICIO& = FIN& + 2
        Else
           TTXX$ = Mid$(XX$, INICIO&)
           List1.AddItem TTXX$
           Exit Do
        End If
   Loop
   
   ' ESTO ES PARA POSICIONARSE AL FINAL DE LA LISTA
   TPXI& = List1.ListCount - 15
   If TPXI& < 0 Then TPXI& = 0
   List1.TopIndex = TPXI&
   '
   List1.Refresh
   Openforms = DoEvents
   '
End Sub

