VERSION 5.00
Begin VB.Form OPPRFLE04 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Presupuestos - Configuración"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6690
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
   ScaleHeight     =   4095
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Formularios de Impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   210
      TabIndex        =   4
      Top             =   2205
      Width           =   5415
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Detalle de Presupuesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1150
         Left            =   210
         TabIndex        =   5
         Top             =   420
         Width           =   4950
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
            Height          =   285
            Left            =   2310
            TabIndex        =   9
            Text            =   " "
            Top             =   315
            Width           =   1380
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
            Height          =   285
            Left            =   3675
            TabIndex        =   8
            Top             =   315
            Width           =   175
         End
         Begin VB.TextBox Text5 
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
            Left            =   210
            TabIndex        =   7
            TabStop         =   0   'False
            Text            =   " "
            Top             =   700
            Width           =   3645
         End
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
            Height          =   680
            Left            =   3990
            Picture         =   "OPPRFLE.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   315
            Width           =   680
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
            Left            =   -105
            TabIndex        =   10
            Top             =   420
            Width           =   2325
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Lista de Precios Por Rubro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   210
      TabIndex        =   3
      Top             =   210
      Width           =   5385
      Begin VB.CommandButton Command11 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   22
         Top             =   1260
         Width           =   180
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
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
         Left            =   1260
         TabIndex        =   21
         Text            =   " "
         Top             =   1260
         Width           =   435
      End
      Begin VB.TextBox Text12 
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
         Left            =   1995
         TabIndex        =   20
         Text            =   " "
         Top             =   1260
         Width           =   3165
      End
      Begin VB.TextBox Text14 
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
         Left            =   1995
         TabIndex        =   19
         Text            =   " "
         Top             =   840
         Width           =   3165
      End
      Begin VB.CommandButton Command12 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   17
         Top             =   840
         Width           =   180
      End
      Begin VB.TextBox Text13 
         Alignment       =   1  'Right Justify
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
         Left            =   1260
         TabIndex        =   16
         Text            =   " "
         Top             =   840
         Width           =   435
      End
      Begin VB.TextBox TxtLiPre 
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
         Left            =   1995
         TabIndex        =   14
         Text            =   " "
         Top             =   420
         Width           =   3165
      End
      Begin VB.TextBox TEXT21 
         Alignment       =   1  'Right Justify
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
         Left            =   1260
         TabIndex        =   13
         Text            =   " "
         Top             =   420
         Width           =   435
      End
      Begin VB.CommandButton CmdLiPre 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   420
         Width           =   180
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Varios:"
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
         Index           =   2
         Left            =   -1470
         TabIndex        =   23
         Top             =   1335
         Width           =   2640
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Migración:"
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
         Left            =   -1470
         TabIndex        =   18
         Top             =   915
         Width           =   2640
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Licencias:"
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
         Left            =   -1470
         TabIndex        =   15
         Top             =   495
         Width           =   2640
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   6705
      Left            =   5775
      ScaleHeight     =   6645
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   -105
      Width           =   1170
      Begin VB.CommandButton Command9 
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
         Left            =   105
         Picture         =   "OPPRFLE.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Tasas de I.V.A."
         Top             =   2800
         Width           =   640
      End
      Begin VB.CommandButton Command8 
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
         Left            =   105
         Picture         =   "OPPRFLE.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Tabla de Tipo de Cambio"
         Top             =   2200
         Width           =   640
      End
      Begin VB.CommandButton Command15 
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
         Left            =   105
         Picture         =   "OPPRFLE.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Acceso a Listas de Precios"
         Top             =   1600
         Width           =   640
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
         Height          =   610
         Left            =   105
         Picture         =   "OPPRFLE.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   300
         Width           =   645
      End
      Begin VB.CommandButton Command2 
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
         Height          =   600
         Left            =   105
         Picture         =   "OPPRFLE.frx":0BB2
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   885
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -300
         Picture         =   "OPPRFLE.frx":0EBC
         Top             =   3620
         Width           =   1515
      End
   End
End
Attribute VB_Name = "OPPRFLE04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdLiPre_Click()
  CmdLiPre.Enabled = False
    '
    ARCHEC$ = "LISTAS"
    Call SELECHO(ARCHEC$)
    VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
    If VDEVU$ <> "" Then
       TEXT21.TEXT = VDEVU$
    End If
    '
  CmdLiPre.Enabled = True
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command11_Click()
  '
  Command11.Enabled = False
  ARCHEC$ = "LISTAS"
  Call SELECHO(ARCHEC$)
  VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
  If VDEVU$ <> "" Then
    Text11.TEXT = VDEVU$
  End If
  Command11.Enabled = True
  '
End Sub

Private Sub Command12_Click()
  '
  Command12.Enabled = False
  ARCHEC$ = "LISTAS"
  Call SELECHO(ARCHEC$)
  VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
  If VDEVU$ <> "" Then
    Text13.TEXT = VDEVU$
  End If
  Command12.Enabled = True
  '
End Sub

Private Sub COMMAND15_CLICK()
   '
   Command15.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "LISPRE04"
   Call CONECRUN(ACONEC$, "Listas de Precios de Venta")
99 Command15.Enabled = True
   '
End Sub

Private Sub Command2_Click()
    '
    'guardo formularios electronicos
    Call GRACONTROL("", "DETAPRE", Text4.TEXT)
    'guardo listas de precios x grupo
    Call GRACONTROL("LIPREVEN", "LICENCIA", TEXT21.TEXT)
    Call GRACONTROL("LIPREVEN", "MIGRADAT", Text13.TEXT)
    Call GRACONTROL("LIPREVEN", "VARIOS", Text11.TEXT)
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


Private Sub Command5_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text4.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call CARTAMONED
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    Call CARDATCENCALC
    Command9.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call TRANSLABELS(Name)
    '
    'LEO LOS FORMULARIOS SETEADOS
    Text4.TEXT = CONTROL$("", "DETAPRE")
    Text5.TEXT = ECOARCH$("INDIFRM", Text4.TEXT)
    'Text7.TEXT = CONTROL$("", "DETAPRE")
    'Text6.TEXT = ECOARCH$("INDIFRM", Text7.TEXT)
    'LEO LAS LISTAS DE PRECIOS SETEADAS
    TEXT21 = CONTROL("LIPREVEN", "LICENCIA")
    Text13 = CONTROL("LIPREVEN", "MIGRADAT")
    Text11 = CONTROL("LIPREVEN", "VARIOS")
    'Text1 = CONTROL("LIPREVEN", "EQUIPO")
    'Text10 = CONTROL("LIPREVEN", "VARIOS")
    '
End Sub

Private Sub Text11_Change()
    If Val(Text11.TEXT) < 0 Or Val(Text11.TEXT) > 255 Then
       Text11.TEXT = ""
    End If
    '
    If Val(Text11) < 1 Then
        Text12.TEXT = "CONCERTADOS FIJOS"
      Else
        Text12.TEXT = ECOARCH$("LISTAS", MKI$(Val(Text11.TEXT)))
    End If
    '
End Sub

Private Sub Text12_GotFocus()
   Text11.SetFocus
End Sub

Private Sub Text13_Change()
    '
    If Val(Text13.TEXT) < 0 Or Val(Text13.TEXT) > 255 Then
       Text13.TEXT = ""
    End If
    '
    If Val(Text13) < 1 Then
        Text14.TEXT = "CONCERTADOS FIJOS"
      Else
        Text14.TEXT = ECOARCH$("LISTAS", MKI$(Val(Text13.TEXT)))
    End If
    '
End Sub

Private Sub Text14_GotFocus()
  On Error Resume Next
   Text13.SetFocus
  On Error GoTo 0
End Sub

Private Sub Text21_Change()
    '
    If Val(TEXT21.TEXT) < 0 Or Val(TEXT21.TEXT) > 255 Then
       TEXT21.TEXT = ""
    End If
    '
    If Val(TEXT21) < 1 Then
        TxtLiPre.TEXT = "CONCERTADOS FIJOS"
      Else
        TxtLiPre.TEXT = ECOARCH$("LISTAS", MKI$(Val(TEXT21.TEXT)))
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

Sub CARTAMONED()
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAMONED/NC/TITUPAN=Monedas y Tipo de Cambio")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAMONED")
         X$ = REGLEIDO$("!TAMONED", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG("TAMONED", X$, U&)
      Next U&
      FIN& = ULTREG&("!TAMONED")
      Call SETULTREG("TAMONED", FIN&)
   End If
   Call CIERRARCH("TAMONED")
End Sub


Private Sub TxtLiPre_GotFocus()
   TEXT21.SetFocus
End Sub
