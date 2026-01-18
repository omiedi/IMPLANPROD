VERSION 5.00
Begin VB.Form FOPCMAQ04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedido de Cotización"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0F0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8590
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12000
      Begin VB.Frame Frame9 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Identificación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   1770
         Left            =   6480
         TabIndex        =   26
         Top             =   120
         Width           =   4280
         Begin VB.TextBox Text2 
            BackColor       =   &H00E0E0E0&
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
            Left            =   2085
            TabIndex        =   35
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1320
            Width           =   2010
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
            Left            =   2085
            TabIndex        =   33
            Text            =   " "
            Top             =   960
            Width           =   2010
         End
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
            Left            =   3900
            TabIndex        =   29
            Top             =   240
            Width           =   195
         End
         Begin VB.TextBox Text9 
            BackColor       =   &H00E0E0E0&
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
            Left            =   2955
            TabIndex        =   28
            TabStop         =   0   'False
            Text            =   " "
            Top             =   240
            Width           =   960
         End
         Begin VB.TextBox Text6 
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
            Left            =   2085
            MaxLength       =   16
            TabIndex        =   27
            Text            =   " "
            Top             =   600
            Width           =   2010
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Solicitante:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   330
            Left            =   600
            TabIndex        =   36
            Top             =   1365
            Width           =   1230
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Destino:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   300
            Left            =   360
            TabIndex        =   34
            Top             =   1035
            Width           =   1485
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
            ForeColor       =   &H00404000&
            Height          =   330
            Left            =   2040
            TabIndex        =   31
            Top             =   285
            Width           =   750
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Referencia:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   300
            Left            =   360
            TabIndex        =   30
            Top             =   670
            Width           =   1485
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Invitados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   1770
         Left            =   1920
         TabIndex        =   23
         Top             =   120
         Width           =   4140
         Begin VB.ListBox List4 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   960
            Left            =   120
            TabIndex        =   24
            ToolTipText     =   "Doble-Click para Ingresar Proveedores Invitados"
            Top             =   630
            Width           =   3900
         End
         Begin VB.Label Label92 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Pr.            Razon Social       "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   300
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   3900
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00D0F0F0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   2990
         Left            =   160
         TabIndex        =   12
         Top             =   2640
         Width           =   10580
         Begin VB.Frame Frame6 
            Caption         =   "Frame18"
            Height          =   4530
            Left            =   7680
            TabIndex        =   22
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame3 
            Caption         =   "Frame18"
            Height          =   4530
            Left            =   6580
            TabIndex        =   21
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame7 
            Caption         =   "Frame7"
            Height          =   4530
            Left            =   1810
            TabIndex        =   15
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame18 
            Caption         =   "Frame18"
            Height          =   4530
            Left            =   5520
            TabIndex        =   14
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame19 
            Caption         =   "Frame19"
            Height          =   4530
            Left            =   10290
            TabIndex        =   13
            Top             =   315
            Width           =   15
         End
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
            Height          =   2400
            Left            =   0
            TabIndex        =   18
            ToolTipText     =   "Doble Click para Cargar Detalle"
            Top             =   600
            Width           =   10560
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00C0C0C0&
            Height          =   435
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   10500
            TabIndex        =   16
            Top             =   315
            Width           =   10565
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   $"FOPCMAQ04.frx":0000
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404000&
               Height          =   225
               Left            =   0
               TabIndex        =   17
               ToolTipText     =   "Doble Click para Cargar Detalle"
               Top             =   0
               Width           =   10515
            End
         End
         Begin VB.Line Line4 
            X1              =   10560
            X2              =   10560
            Y1              =   315
            Y2              =   4880
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "!"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1680
         TabIndex        =   19
         ToolTipText     =   "Lista de Proveedores con Entregas Programadas"
         Top             =   9520
         Width           =   195
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Condiciones / Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   2325
         Left            =   160
         TabIndex        =   10
         Top             =   5820
         Width           =   10590
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   320
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            ToolTipText     =   "Click para Ingresar las Condiciones / Observaciones"
            Top             =   330
            Width           =   9960
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00004000&
         Height          =   9000
         Left            =   11025
         ScaleHeight     =   8940
         ScaleWidth      =   1005
         TabIndex        =   1
         Top             =   0
         Width           =   1065
         Begin VB.CommandButton Command3 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   640
            Left            =   105
            Picture         =   "FOPCMAQ04.frx":009D
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   790
            Width           =   650
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Prove."
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FOPCMAQ04.frx":03A7
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Acceso a Base de Datos de Proveedores"
            Top             =   1550
            Width           =   650
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Rep."
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FOPCMAQ04.frx":06B1
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Acceso a Maestro de Repuestos"
            Top             =   2240
            Width           =   650
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Setup"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FOPCMAQ04.frx":09BB
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Configuración de Funcionamiento de Pedido de Cotización"
            Top             =   3045
            Width           =   650
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FOPCMAQ04.frx":0DFD
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   100
            Width           =   650
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H00E0E0E0&
            Caption         =   "O.K."
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1110
            Left            =   120
            Picture         =   "FOPCMAQ04.frx":0F47
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Aprueba e Imprime Pedido de Cotización"
            Top             =   6930
            Width           =   650
         End
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   105
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   2
            Top             =   6045
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   3
               Top             =   105
               Width           =   12000
            End
         End
         Begin VB.Image Image1 
            Height          =   390
            Left            =   -315
            Picture         =   "FOPCMAQ04.frx":1251
            Top             =   6285
            Width           =   1515
         End
      End
      Begin VB.Line Line7 
         X1              =   165
         X2              =   10740
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Por el Presente solicitamos la Cotización de los siguientes Items:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   2280
         Width           =   10455
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   750
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   630
         Shape           =   4  'Rounded Rectangle
         Top             =   375
         Width           =   540
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000C000&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   390
         Left            =   525
         Shape           =   4  'Rounded Rectangle
         Top             =   315
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   420
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   960
      End
   End
End
Attribute VB_Name = "FOPCMAQ04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command11_Click()
   Command11.Enabled = False
'   Call CIERRARCH("*.*")
'   ACONEC$ = "AMAPRO"
'   Call CONECRUN(ACONEC$, "Padron Maestro de Proveedores")
   ACRE_MAQ04.Show 1
   Command11.Enabled = True
End Sub

Private Sub Command14_Click()
  '
  'IMPRIME COTIZACION
  '
  FORIPRE$ = CONTROL$("", "FORMPECO")
  '
  If TRIM$(FORIPRE$) <> "" Then
    '
    For JX% = 0 To List4.ListCount - 1 'UNO POR CADA PROVEEDOR
      L1$ = List4.List(JX%)
      NPRO% = Val(Mid$(L1$, 1, 5))
      Screen.MousePointer = 11
      FECHDOC$ = Text9.TEXT
      NUMEDOC$ = TRIM$(Text6)
      TIPODOC$ = "Pedido de Cotización"
      '
      CODPARAM$(1) = "USUARIO"
      CODPARAM$(2) = "REF-MAT2"
      CAPARDOC% = 2
      '
      DOCPARAM$(1) = USERNAME$
      DOCPARAM$(2) = Text1.TEXT
      '
      CODTABU1$(1) = "COD-MPRI"
      CODTABU1$(2) = "DES-MPRI"
      CODTABU1$(3) = "CAN-MPRI"
      CODTABU1$(4) = "FECH-REC"
      CODTABU1$(5) = "PAR-MPRI"
      CACOLTAB1% = 5
      '
      DESTIDOC% = NPRO% * (-1)
      CAITAB1% = 0
      CAITCRU% = 0
      '
      NC1% = DESTIDOC%
      For J1% = 0 To List1.ListCount - 1
         X$ = List1.List(J1%)
         CI0% = CODINT(Mid$(X$, 1, 16))
         CANTU = Val(Mid$(X$, 54, 10))
         FEX = FECHANUM(TRIM$(Mid$(X$, 65, 10)))
         OBS1$ = Mid$(X$, 76, 25)
         CAITAB1% = CAITAB1% + 1
         CODDX$ = TRIM$(CODEXT$(CI0%))
         DECRIT$ = TRIM$(DESCRIT0$(CI0%))
         '
         TETABU1$(1, CAITAB1%) = CODDX$
         TETABU1$(2, CAITAB1%) = DECRIT$
         TETABU1$(3, CAITAB1%) = TRIM$(Str$(CANTU))
         TETABU1$(4, CAITAB1%) = FECHATEX$(FEX)
         TETABU1$(5, CAITAB1%) = OBS1$
      Next J1%
      '
      CODTABU2$(1) = "MEDIDA-2" 'DESCRIPCION
      CACOLTAB2% = 1
      '
      CAITAB2% = 0
      DESGRAL$ = Text13.TEXT 'OBSERVACIONES
      Call FRATEXTO(DESGRAL$, 92)
      For U& = 1 To CARETEX%
        Z$ = "." + RETEX$(U&)
        CAITAB2% = CAITAB2% + 1
        TETABU2$(1, CAITAB2%) = Z$
      Next U&
      '
      Call PRINTDOC("FORMPECO")   ' Pedido de Cotización
      '
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
      For KK1% = 1 To CACOLTAB2%
        For KK2% = 1 To CAITAB2%
          TETABU2$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
    Next JX%
     '
99  Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
    Call BLANFORMU
    '
   End If
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("PEDICOTI")
    Command3.Enabled = True
End Sub

Private Sub Command5_Click()
  '
  OPPCMAQ04.Show 1
  '
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
'    Call CIERRARCH("*.*")
'    ACONEC$ = "AMASTO01"
'    If EXISTE%("AMASTO02.EXE") > 0 Then ACONEC$ = "AMASTO02"
'    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
    AMAS_MAQ04.Show 1
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text9.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text9.TEXT)
End Sub

Private Sub Form_Load()
  '
  Call CENTRAFORM(Me)
  EPAC$ = TRIM$(EMPREAC$)
  If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
  End If
  Call GRATITFOR(Caption)
  '
  FECHA1$ = Format(Now, "yymmdd")
  ID1$ = CONTROL$("", "IDENPECO")
  If ID1$ = "" Then ID1$ = "PC"
  Text6.TEXT = ID1$ + "-" + FECHA1$
  '
  HOII% = HOY(HOS$)
  Text9.TEXT = HOS$
  Text2.TEXT = Left$(USERNAME, 20)
  '
End Sub
Sub BLANFORMU()
  '
  FECHA1$ = Format(Now, "yymmdd")
  ID1$ = CONTROL$("", "IDENPECO")
  If ID1$ = "" Then ID1$ = "PC"
  Text6.TEXT = ID1$ + "-" + FECHA1$
  '
  HOII% = HOY(HOS$)
  Text9.TEXT = HOS$
  '
  Text13.TEXT = ""
  List1.Clear
  List4.Clear
  Text1.TEXT = ""
  '
End Sub

Private Sub Label22_DBLClick()
    '
    Call BLANARCH("!LIPECOTI")
    JJ& = 0
    For JX% = 0 To List1.ListCount - 1
      L1$ = List1.List(JX%)
      CODI0$ = Mid$(L1$, 1, 16)
        CODI1% = CODINT%(CODI0$)
      DESC1$ = Mid$(L1$, 18, 35)
      CAN1 = Val(Mid$(L1$, 54, 10))
      FES1$ = TRIM$(Mid$(L1$, 65, 10))
      OBS1$ = Mid$(L1$, 76, 25)
      '
      YY$ = REGBLAN$("LIPECOTI")
      Call REPLA(YY$, MKI$(CODI1%), 1, 2)
      Call REPLA(YY$, MKS$(CAN1), 3, 4)
      Call REPLA(YY$, MKI$(FECHANUM(FES1$)), 7, 2)
      Call REPLA(YY$, OBS1$, 9, 25)
      JJ& = JJ& + 1
      Call GRAREG("!LIPECOTI", YY$, JJ&)
      '
    Next JX%
    Call SETULTREG("!LIPECOTI", JJ&)
    Call CIERRARCH("!LIPECOTI")
    '
    PRF$ = "LIPECOTI"
    ATRI$ = ATRI$ + "/TITUPAN=Lista de Repuestos /"
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + FOPCMAQ04.Left + 170))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + FOPCMAQ04.Top + 2250))
    VTB% = VENTABU%(PRF$ + ATRI$)
    If VTB% >= 0 Then
      '
      List1.Clear
      '
      FIN& = ULTREG&("!LIPECOTI")
      For J& = 1 To FIN&
        X$ = REGLEIDO$("!LIPECOTI", J&)
        CODI1% = CVI(Mid$(X$, 1, 2))
        CODI2$ = CODEXT$(CODI1%)
        DESC2$ = DESCRIT0(CODI1%)
        CAN1 = CVS(Mid$(X$, 3, 4))
        CAN2$ = FORMATNUM$(CAN1, "F9.1")
        FES2% = CVI(Mid$(X$, 7, 2))
        OBS2$ = Mid$(X$, 9, 55)
        'AGREGO AL LIST
        L1$ = Space(128)
        Call REPLA(L1$, CODI2$, 1, 16)
        Call REPLA(L1$, DESC2$, 18, 35)
        Call REPLA(L1$, CAN2$, 54, 10)
        Call REPLA(L1$, FECHATEX$(FES2%), 65, 10)
        Call REPLA(L1$, OBS2$, 76, 25)
        List1.AddItem L1$
      Next J&
      '
    End If
End Sub

Private Sub Label92_dblClick()
  '
  Call BLANARCH("!INVITA2")
  JJ& = 0
  For JX% = 0 To List4.ListCount - 1
    L1$ = List4.List(JX%)
    NPRO% = Mid$(L1$, 1, 5)
    '
    YY$ = REGBLAN$("INVITA2")
    Call REPLA(YY$, MKI$(NPRO%), 1, 2)
    JJ& = JJ& + 1
    Call GRAREG("!INVITA2", YY$, JJ&)
    '
  Next JX%
  Call SETULTREG("!INVITA2", JJ&)
  Call CIERRARCH("!INVITA2")
  '
  PRF$ = "INVITA2"
  ATRI$ = ATRI$ + "/TITUPAN=Lista de Proveedores Invitados /"
  ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + FOPCMAQ04.Left + 1000))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + FOPCMAQ04.Top + 5))
  VTB% = VENTABU%(PRF$ + ATRI$)
  If VTB% >= 0 Then
    '
    List4.Clear
    '
    FIN& = ULTREG&("!INVITA2")
    For J& = 1 To FIN&
      X$ = REGLEIDO$("!INVITA2", J&)
      NPRO1% = CVI(Mid$(X$, 1, 2))
        NPRO2$ = FORMATNUM$(NPRO1%, "F5.0")
      PRO1$ = ECOARCH$("PROVED1", MKI$(NPRO1%))
      'AGREGO AL LIST
      L1$ = Space(128)
      Call REPLA(L1$, NPRO2$, 1, 5)
      Call REPLA(L1$, PRO1$, 8, 20)
      List4.AddItem L1$
    Next J&
    '
  End If
  '
End Sub

Private Sub List1_DblClick()
  '
  Call Label22_DBLClick
  '
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  '
  CODI3$ = Left(List1.List(List1.ListIndex), 16)
  CODI2% = CODINT%(CODI3$)
  List1.ToolTipText = DESCRIT0(CODI2%)
  '
End Sub

Private Sub List4_Click()
  '
  Call Label92_dblClick
  '
End Sub
