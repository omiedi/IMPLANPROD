VERSION 5.00
Begin VB.Form OPCFPRE04 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Presupuestos - Configuración"
   ClientHeight    =   6300
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
   ScaleHeight     =   6300
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
      Height          =   3105
      Left            =   210
      TabIndex        =   4
      Top             =   3045
      Width           =   5415
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Presupuestos de Cámara"
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
         TabIndex        =   11
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
            TabIndex        =   15
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
            TabIndex        =   14
            Top             =   315
            Width           =   175
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
            Height          =   285
            Left            =   210
            TabIndex        =   13
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
            Picture         =   "OPCFPRE04.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   12
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
            TabIndex        =   16
            Top             =   420
            Width           =   2325
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Detalle"
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
         Top             =   1785
         Width           =   4950
         Begin VB.CommandButton Command4 
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
            Picture         =   "OPCFPRE04.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   315
            Width           =   680
         End
         Begin VB.TextBox Text6 
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
            TabIndex        =   8
            TabStop         =   0   'False
            Text            =   " "
            Top             =   700
            Width           =   3645
         End
         Begin VB.CommandButton Command6 
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
            TabIndex        =   7
            Top             =   315
            Width           =   175
         End
         Begin VB.TextBox Text7 
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
            TabIndex        =   6
            Text            =   " "
            Top             =   315
            Width           =   1380
         End
         Begin VB.Label Label5 
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
      Height          =   2625
      Left            =   210
      TabIndex        =   3
      Top             =   210
      Width           =   5385
      Begin VB.CommandButton Command7 
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
         TabIndex        =   36
         Top             =   2100
         Width           =   180
      End
      Begin VB.TextBox Text10 
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
         TabIndex        =   35
         Text            =   " "
         Top             =   2100
         Width           =   435
      End
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
         Left            =   1995
         TabIndex        =   34
         Text            =   " "
         Top             =   2100
         Width           =   3165
      End
      Begin VB.CommandButton Command10 
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
         TabIndex        =   32
         Top             =   1680
         Width           =   180
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
         Left            =   1260
         TabIndex        =   31
         Text            =   " "
         Top             =   1680
         Width           =   435
      End
      Begin VB.TextBox Text3 
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
         TabIndex        =   30
         Text            =   " "
         Top             =   1680
         Width           =   3165
      End
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
         TabIndex        =   28
         Top             =   1260
         Width           =   180
      End
      Begin VB.TextBox Text11 
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   23
         Top             =   840
         Width           =   180
      End
      Begin VB.TextBox Text13 
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
         TabIndex        =   22
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
         TabIndex        =   20
         Text            =   " "
         Top             =   420
         Width           =   3165
      End
      Begin VB.TextBox TEXT21 
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         Index           =   4
         Left            =   -1470
         TabIndex        =   37
         Top             =   2175
         Width           =   2640
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Equipos;"
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
         Index           =   3
         Left            =   -1470
         TabIndex        =   33
         Top             =   1755
         Width           =   2640
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Accesorios:"
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
         TabIndex        =   29
         Top             =   1335
         Width           =   2640
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Puertas:"
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
         TabIndex        =   24
         Top             =   915
         Width           =   2640
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cámara:"
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
         TabIndex        =   21
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
         Picture         =   "OPCFPRE04.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Tasas de I.V.A."
         Top             =   3780
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
         Picture         =   "OPCFPRE04.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Tabla de Tipo de Cambio"
         Top             =   3150
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
         Picture         =   "OPCFPRE04.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Acceso a Listas de Precios"
         Top             =   2520
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
         Picture         =   "OPCFPRE04.frx":0D72
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
         Picture         =   "OPCFPRE04.frx":0EBC
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   885
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "OPCFPRE04.frx":11C6
         Top             =   5775
         Width           =   1515
      End
   End
End
Attribute VB_Name = "OPCFPRE04"
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

Private Sub Command10_Click()
  '
  Command10.Enabled = False
  ARCHEC$ = "LISTAS"
  Call SELECHO(ARCHEC$)
  VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
  If VDEVU$ <> "" Then
    Text1.TEXT = VDEVU$
  End If
  Command10.Enabled = True
  '
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
    'guardo formularios electronicos
    IDORTRAB = TRIM$(Text4.TEXT)
    Call GRACONTROL("", "ESPVENTA", Text4.TEXT)
    Call GRACONTROL("", "DETAPRE", Text7.TEXT)
    'guardo listas de precios x grupo
    Call GRACONTROL("LIPREVEN", "CAMARA", TEXT21.TEXT)
    Call GRACONTROL("LIPREVEN", "PUERTA", Text13.TEXT)
    Call GRACONTROL("LIPREVEN", "ACCESO", Text11.TEXT)
    Call GRACONTROL("LIPREVEN", "EQUIPO", Text1.TEXT)
    Call GRACONTROL("LIPREVEN", "VARIOS", Text10.TEXT)
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
    Call CONECRUN("FLFORMS/" + TRIM$(Text7.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command5_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text4.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command6_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Text6.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command7_Click()
  '
  Command7.Enabled = False
  ARCHEC$ = "LISTAS"
  Call SELECHO(ARCHEC$)
  VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
  If VDEVU$ <> "" Then
    Text10.TEXT = VDEVU$
  End If
  Command7.Enabled = True
  '
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
    Call CARUTRA
    '
    Text1.TEXT = IDENSEDE$
    Text2.TEXT = DESTITRA$
    'LEO LOS FORMULARIOS SETEADOS
    Text4.TEXT = CONTROL$("", "ESPVENTA")
    Text5.TEXT = ECOARCH$("INDIFRM", Text4.TEXT)
    Text7.TEXT = CONTROL$("", "DETAPRE")
    Text6.TEXT = ECOARCH$("INDIFRM", Text7.TEXT)
    'LEO LAS LISTAS DE PRECIOS SETEADAS
    TEXT21 = CONTROL("LIPREVEN", "CAMARA")
    Text13 = CONTROL("LIPREVEN", "PUERTA")
    Text11 = CONTROL("LIPREVEN", "ACCESO")
    Text1 = CONTROL("LIPREVEN", "EQUIPO")
    Text10 = CONTROL("LIPREVEN", "VARIOS")
    '
End Sub

Private Sub Text1_Change()
    If Val(Text1.TEXT) < 0 Or Val(Text1.TEXT) > 255 Then
       Text1.TEXT = ""
    End If
    '
    If Val(Text1) < 1 Then
        Text3.TEXT = "CONCERTADOS FIJOS"
      Else
        Text3.TEXT = ECOARCH$("LISTAS", MKI$(Val(Text1.TEXT)))
    End If
    '
End Sub

Private Sub Text10_Change()
    If Val(Text10) < 0 Or Val(Text10) > 255 Then
       Text10 = ""
    End If
    '
    If Val(Text10) < 1 Then
        Text2.TEXT = "CONCERTADOS FIJOS"
      Else
        Text2.TEXT = ECOARCH$("LISTAS", MKI$(Val(Text10)))
    End If
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

Private Sub Text2_GotFocus()
  On Error Resume Next
    Text10.SetFocus
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

Private Sub Text3_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
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

Private Sub Text8_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text8.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("INDIFRM")
    End If
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
