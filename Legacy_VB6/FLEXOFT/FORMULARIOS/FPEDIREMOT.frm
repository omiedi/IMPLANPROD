VERSION 5.00
Begin VB.Form FPEDIREMOT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0FFEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos Remotos "
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9510
   Icon            =   "FPEDIREMOT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0FFEE&
      Caption         =   "Datos de Pedido de Cliente"
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
      Height          =   2175
      Left            =   105
      TabIndex        =   20
      Top             =   105
      Width           =   8100
      Begin VB.Frame Frame2 
         BackColor       =   &H00D0FFEE&
         Caption         =   "Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1330
         Left            =   5775
         TabIndex        =   26
         Top             =   630
         Width           =   2115
         Begin VB.OptionButton Option2 
            BackColor       =   &H00D0FFEE&
            Caption         =   "B"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1260
            TabIndex        =   8
            Top             =   1000
            Width           =   645
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0FFEE&
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   525
            TabIndex        =   7
            Top             =   1000
            Value           =   -1  'True
            Width           =   645
         End
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
            Left            =   1785
            TabIndex        =   5
            Top             =   315
            Width           =   180
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nro:"
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
            TabIndex        =   28
            Top             =   690
            Width           =   960
         End
         Begin VB.Label Label4 
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
            Left            =   840
            TabIndex        =   6
            Top             =   630
            Width           =   1065
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
            Left            =   -210
            TabIndex        =   27
            Top             =   375
            Width           =   960
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
            Left            =   840
            TabIndex        =   4
            Top             =   315
            Width           =   1065
         End
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   1
         Top             =   315
         Width           =   5370
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1575
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1365
         Width           =   4110
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   1575
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   735
         Width           =   4110
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1575
         TabIndex        =   0
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
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
         Left            =   105
         TabIndex        =   24
         Top             =   1365
         Width           =   1380
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
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
         Left            =   525
         TabIndex        =   22
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Entregar en:"
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
         TabIndex        =   21
         Top             =   735
         Width           =   1275
      End
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   105
      TabIndex        =   11
      Top             =   2835
      Width           =   8220
   End
   Begin VB.PictureBox Picture1 
      Height          =   435
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   8160
      TabIndex        =   19
      Top             =   2415
      Width           =   8220
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   $"FPEDIREMOT.frx":030A
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   9
         Top             =   60
         Width           =   10410
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   8445
      ScaleHeight     =   75
      ScaleWidth      =   945
      TabIndex        =   17
      Top             =   5920
      Width           =   1000
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00D0FFEE&
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2800
      Left            =   8445
      TabIndex        =   16
      Top             =   105
      Width           =   1000
      Begin VB.CommandButton COMMAND3 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   740
         Left            =   105
         Picture         =   "FPEDIREMOT.frx":0396
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   1900
         Width           =   750
      End
      Begin VB.CommandButton Command1 
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
         Height          =   740
         Left            =   105
         Picture         =   "FPEDIREMOT.frx":07D8
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   375
         Width           =   750
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   740
         Left            =   105
         Picture         =   "FPEDIREMOT.frx":0922
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1140
         Width           =   750
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00D0FFEE&
      Caption         =   "Accion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2800
      Left            =   8445
      TabIndex        =   10
      ToolTipText     =   "Maestro de Items de Stock"
      Top             =   3000
      Width           =   1000
      Begin VB.CommandButton Command7 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   740
         Left            =   105
         Picture         =   "FPEDIREMOT.frx":0C2C
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Limpriar Formulario"
         Top             =   1890
         Width           =   750
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   740
         Left            =   105
         Picture         =   "FPEDIREMOT.frx":0F36
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Enviar Pedido por Mail"
         Top             =   375
         Width           =   750
      End
      Begin VB.CommandButton Command21 
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
         Height          =   740
         Left            =   105
         Picture         =   "FPEDIREMOT.frx":1240
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Impresion de Pedido"
         Top             =   1140
         Width           =   750
      End
   End
   Begin VB.Line Line3 
      X1              =   105
      X2              =   11865
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line1 
      X1              =   8325
      X2              =   8325
      Y1              =   -105
      Y2              =   6720
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   7950
      Picture         =   "FPEDIREMOT.frx":18AA
      Top             =   6210
      Width           =   2010
   End
End
Attribute VB_Name = "FPEDIREMOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command21_Click()
   Command1.Enabled = False
   NCLIE% = Val(Text1)
   DCLIEN$ = UCase$(TRIM$(Text4))
   LENTRE$ = Text2
   OBSER$ = Text3
   FEPED$ = Label8: FENTBA% = FECHANUM(FEPED$)
   GRUE% = 1: If Option2.Value = True Then GRUE% = 2
   '
   CRLF$ = ""
   Printer.Print "Numero Cliente = " + TRIM$(Str$(NCLIE%))
   Printer.Print "Denominacion Cliente = " + TRIM$(DCLIEN$)
   Printer.Print "Fecha Pedido = " + TRIM$(FEPED$) + CRLF$
   Printer.Print "Sucursal = " + TRIM$(Str$(GRUE%)) + CRLF$
   Printer.Print "Lugar de Entrega = " + TRIM$(LENTRE$) + CRLF$
   Printer.Print "Observaciones = " + TRIM$(OBSER$) + CRLF$
   Printer.Print ""
   Printer.Print ""
   '
   For U& = 1 To ULTREG&("!PEDREMOT")
     TREN$ = ""
     XX$ = REGLEIDO$("!PEDREMOT", U&)
     CODD$ = TRIM$(Left$(XX$, 16))
     CANT = CVS(Mid$(XX$, 17, 4))
     CANTX$ = TRIM$(FORMATNUM$(CANT, "F12.0"))
     FEX = CVI(Mid$(XX$, 21, 2))
     FENTX$ = TRIM$(FECHATEX$(FEX))
     OBSX$ = TRIM$(Mid$(XX$, 23, 42))
     '
     TREN$ = CODD$ + ", " + CANTX$ + ", " + FENTX$ + ", " + OBSX$
     Printer.Print TREN$ + CRLF$
   Next U&
   Printer.Print String$(80, "-")
   Printer.EndDoc
   '
   Command1.Enabled = True
End Sub

Private Sub COMMAND3_Click()
   COMMAND3.Enabled = False
   SETUPVEN.Show 1
   COMMAND3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call SELECHO("SELFECHA")
   If VALACT1$("SELFECHA") <> "" Then
     Label8 = VALACT1$("SELFECHA")
   End If
   Command4.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call SELECHO("PROVED1")
   If VALACT1$("PROVED1") <> "" Then
     Label6 = VALACT1$("PROVED1")
     On Error Resume Next
     Text2.SetFocus
     On Error GoTo 0
   End If
   Command6.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Call COMUNI("Ayuda no Disponible")
   Command5.Enabled = True
End Sub

Private Sub Command7_Click()
   '
   Command7.Enabled = False
   If COMALTER%("Está Seguro de Blanquear\el Pedido en Curso ?\\No, Mantenerlo\Si, Borrarlo") = 2 Then
      Text1 = ""
      Text2 = ""
      Text3 = ""
      Text4 = ""
      List2.Clear
      Call SETULTREG("!PEDREMOT", 0)
   End If
   Command7.Enabled = True
   '
End Sub


Private Sub Command9_Click()
   Command9.Enabled = False
   '
   NCLIE% = Val(Text1)
   If NCLIE% < 1 Or NCLIE% > 9999 Then
     Call MENSERR(24, "Número de Cliente no Válido.")
     GoTo 99
   End If
   '
   DCLIEN$ = UCase$(TRIM$(Text4))
   Text4 = DCLIEN$
   If DCLIEN$ = "" Then
     Call MENSERR(24, "Falta Denominación del Cliente")
     GoTo 99
   End If
   '
   LENTRE$ = Text2
   OBSER$ = Text3
   FEPED$ = Label8: FENTBA% = FECHANUM(FEPED$)
   GRUE% = 1: If Option2.Value = True Then GRUE% = 2
   '
   CAITE% = 0: ATX% = 0
   For U& = 1 To ULTREG&("!PEDREMOT")
     XX$ = REGLEIDO$("!PEDREMOT", U&)
     CODD$ = TRIM$(Left$(XX$, 16))
     If Len(CODD$) < 3 Then ATX% = 1
     CANT = CVS(Mid$(XX$, 17, 4))
     If CANT < 0.05 Then ATX% = 1
     If CVI(Mid$(XX$, 21, 2)) < 1 Then
       Call REPLA(XX$, MKI$(FENTBA%), 21, 2)
       Call GRAREG("!PEDREMOT", XX$, U&)
       ATX% = 1
     End If
     CAITE% = CAITE% + 1
   Next U&
   '
   If CAITE% < 1 Then
     Call MENSERR(24, "Falta Cargar Detalle de Pedido.")
     GoTo 99
   End If
   '
   If ATX% > 0 Then
     Call MENSERR(24, "Datos Incompletos o no Válidos.")
     GoTo 99
   End If
   '
   DVENDE$ = TRIM$(CONTROL$("VENREMOT", "VENDEDOR"))
   DESTIM$ = TRIM$(CONTROL$("VENREMOT", "DESTIMAI"))
   DIRELM$ = TRIM$(CONTROL$("VENREMOT", "DIRELMAI"))
   '
   '
   CRLF$ = Chr$(13) + Chr$(10)
   ASUNTO$ = "P-Elec." + DVENDE$ + " " + DCLIEN$
   TEXTOMAI$ = TEXTOMAI$ + "N-CLIEN=" + TRIM$(Str$(NCLIE%)) + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "D-CLIEN=" + TRIM$(DCLIEN$) + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "F-PEDID=" + TRIM$(FEPED$) + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "N-SUCUR=" + TRIM$(Str$(GRUE%)) + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "L-ENTRE=" + TRIM$(LENTRE$) + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "T-OBSER=" + TRIM$(OBSER$) + CRLF$
   '
   For U& = 1 To ULTREG&("!PEDREMOT")
     TREN$ = ""
     XX$ = REGLEIDO$("!PEDREMOT", U&)
     CODD$ = AJUSTI$(Left$(XX$, 16), 16)
     CANT = CVS(Mid$(XX$, 17, 4))
     CANTX$ = AJUSTD$(FORMATNUM$(CANT, "F12.1"), 12)
     FEX = CVI(Mid$(XX$, 21, 2))
     FENTX$ = FECHATEX$(FEX)
     OBSX$ = TRIM$(Mid$(XX$, 23, 42))
     '
     TREN$ = CODD$ + "  " + CANTX$ + "  " + FENTX$ + "  " + OBSX$
     TEXTOMAI$ = TEXTOMAI$ + TREN$ + CRLF$
   Next U&
   '
   MAI_DESTI$ = DESTIM$
   MAI_DIREL$ = DIRELM$
   MAI_ASUNT$ = ASUNTO$
   MAI_TEXTO$ = TEXTOMAI$
   MAI_ADJUN$(0) = ""
   Call GENERAMAIL
   '
   MsgBox "FLEXOFT ha Despachado su Pedido."
   '
   Text1 = ""
   Text2 = ""
   Text3 = ""
   Text4 = ""
   List2.Clear
   Call SETULTREG("!PEDREMOT", 0)
   '
   FEX = HOY(HOS$)
   Label8 = FECHATEX$(FEX)
   
   
99 Command9.Enabled = True
End Sub

Private Sub Form_Load()
   '
   Call AUTOEXEC
   '
   Show
   Refresh
   '
   FEX = HOY(HOS$)
   Label8 = FECHATEX$(FEX)
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Label15_DblClick()
   Call CARDETPEDI
End Sub

Private Sub Label6_Change()
   Label11 = ""
   If Val(Label6) > 0 Then
     Label11 = RASOCLI$((-1) * Val(Label6))
   End If
End Sub

Private Sub Label7_Change()
   Label9 = ECOARCH$("VALIMONE", Chr$(Val(Label7)))
End Sub

Private Sub Label7_DblClick()
   Call SELECHO("VALIMONE")
   If VALACT1$("VALIMONE") <> "" Then
     Label7 = VALACT1$("VALIMONE")
   End If
End Sub

Private Sub Label8_DblClick()
   Call SELECHO("SELFECHA")
   If VALACT1$("SELFECHA") <> "" Then
     Label8 = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub List2_DblClick()
   '
   Call CARDETPEDI
   '
End Sub

Private Sub List2_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
     Call List2_DblClick
   End If
End Sub

Private Sub Text1_GotFocus()
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.TEXT)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
     On Error Resume Next
     Command4.SetFocus
     On Error GoTo 0
   End If
End Sub

Private Sub Text2_GotFocus()
   Text2.SelStart = 0
   Text2.SelLength = Len(Text2.TEXT)
End Sub

Private Sub Text3_GotFocus()
   Text3.SelStart = 0
   Text3.SelLength = Len(Text3.TEXT)
End Sub

Sub CARDETPEDI()
   '
   HOII% = HOY(HOS$)
   '
   ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=Detalle Pedido de Cliente"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left + 50))
   '
10 XXX% = VENTABU%("PEDREMOT" + ATRI$)
   If XXX% < 0 Then
      Call CIERRARCH("!PEDREMOT")
      Call BLANARCH("!PEDREMOT")
      Exit Sub
   End If
   '
   List2.Clear
   FIN& = ULTREG&("!PEDREMOT")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!PEDREMOT", U&)
     TTXX$ = Space$(96)
     CODD$ = TRIM$(Left$(X$, 16))
     CAN = CVS(Mid$(X$, 17, 4))
     FENTRESOL% = CVI(Mid$(X$, 21, 2))
     OBSER$ = TRIM$(Mid$(X$, 23, 42))
     '
     Call REPLA(TTXX$, CODD$, 1, 16)
     Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 19, 6)
     FEX = FENTRESOL%
     Call REPLA(TTXX$, FECHATEX$(FEX), 27, 8)
     Call REPLA(TTXX$, OBSER$, 37, 42)
     List2.AddItem TTXX$
   Next U&

End Sub
'

