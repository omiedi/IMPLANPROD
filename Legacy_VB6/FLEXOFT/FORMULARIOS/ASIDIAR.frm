VERSION 5.00
Begin VB.Form ASIDIAR 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asientos Contables"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   8610
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   42
      Top             =   4620
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   8610
      Sorted          =   -1  'True
      TabIndex        =   41
      Top             =   4410
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   390
      Left            =   8085
      TabIndex        =   39
      Top             =   735
      Value           =   16384
      Width           =   175
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Referencia"
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
      Left            =   210
      TabIndex        =   38
      Top             =   7140
      Width           =   9255
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
         Left            =   1260
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   250
         Width           =   7785
      End
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6510
      TabIndex        =   1
      Top             =   735
      Width           =   1590
   End
   Begin VB.CommandButton Command21 
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
      Height          =   390
      Left            =   8085
      TabIndex        =   25
      Top             =   210
      Width           =   175
   End
   Begin VB.CommandButton Command5 
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
      Height          =   800
      Left            =   8610
      Picture         =   "ASIDIAR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Imprime Comprobante del Asiento"
      Top             =   3570
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Registrar Pase - Movimiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      Left            =   210
      TabIndex        =   21
      Top             =   4830
      Width           =   9255
      Begin VB.CommandButton Command7 
         Height          =   645
         Left            =   8505
         Picture         =   "ASIDIAR.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Activa Calculadora del Sistema"
         Top             =   1470
         Width           =   480
      End
      Begin VB.CommandButton Command6 
         Height          =   645
         Left            =   7980
         Picture         =   "ASIDIAR.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Agrega Pase al Asiento"
         Top             =   1470
         Width           =   480
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFFF&
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
         Height          =   855
         Left            =   105
         TabIndex        =   37
         Top             =   1260
         Width           =   4215
         Begin VB.TextBox Text12 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   105
            TabIndex        =   5
            TabStop         =   0   'False
            Text            =   " "
            Top             =   315
            Width           =   4005
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Balance / Diferencia"
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
         Left            =   4515
         TabIndex        =   30
         Top             =   1260
         Width           =   3270
         Begin VB.Label Label10 
            BackColor       =   &H80000016&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   330
            Left            =   1785
            TabIndex        =   32
            Top             =   315
            Width           =   1380
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000016&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   330
            Left            =   105
            TabIndex        =   31
            Top             =   315
            Width           =   1380
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Haber"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   6195
         TabIndex        =   29
         Top             =   210
         Width           =   1590
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   105
            MaxLength       =   11
            TabIndex        =   7
            Text            =   " "
            Top             =   420
            Width           =   1380
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Debe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   4515
         TabIndex        =   28
         Top             =   210
         Width           =   1590
         Begin VB.TextBox Text14 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   105
            MaxLength       =   11
            TabIndex        =   6
            Text            =   " "
            Top             =   420
            Width           =   1380
         End
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   7980
         Picture         =   "ASIDIAR.frx":0C7E
         TabIndex        =   12
         ToolTipText     =   "Eldición el Pase Activo para su Modificación"
         Top             =   1115
         Width           =   1000
      End
      Begin VB.CommandButton Command31 
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
         Height          =   360
         Left            =   3990
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   315
         Width           =   175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   7980
         Picture         =   "ASIDIAR.frx":0F88
         TabIndex        =   11
         ToolTipText     =   "Elimina el Pase Activo del Asiento"
         Top             =   840
         Width           =   1000
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Insert"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   7980
         Picture         =   "ASIDIAR.frx":1292
         TabIndex        =   10
         ToolTipText     =   "Intercala Pase Encima del Activo"
         Top             =   575
         Width           =   1000
      End
      Begin VB.TextBox Text11 
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
         Left            =   2205
         TabIndex        =   3
         Top             =   315
         Width           =   1800
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Append"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Left            =   7980
         Picture         =   "ASIDIAR.frx":159C
         TabIndex        =   9
         ToolTipText     =   "Agrega Pase al Final del Asiento"
         Top             =   300
         Width           =   1000
      End
      Begin VB.Label Label4 
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
         Height          =   330
         Left            =   210
         TabIndex        =   27
         Top             =   840
         Width           =   4005
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         Left            =   1155
         TabIndex        =   22
         Top             =   420
         Width           =   960
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   3350
      Left            =   180
      TabIndex        =   19
      Top             =   1365
      Width           =   8100
      Begin VB.PictureBox Picture1 
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   8040
         TabIndex        =   34
         Top             =   2940
         Width           =   8100
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000016&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   6370
            TabIndex        =   36
            Top             =   30
            Width           =   1380
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000016&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   4915
            TabIndex        =   35
            Top             =   30
            Width           =   1380
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2670
         Left            =   0
         TabIndex        =   20
         Top             =   315
         Width           =   8100
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   360
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   9000
         Begin VB.Label Label5 
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Cuenta                    Detalle                                                 Debe             Haber    "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   8625
         End
      End
      Begin VB.Line Line4 
         X1              =   7000
         X2              =   10
         Y1              =   0
         Y2              =   420
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6510
      TabIndex        =   0
      Top             =   210
      Width           =   1590
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
      Height          =   800
      Left            =   8610
      Picture         =   "ASIDIAR.frx":18A6
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   2730
      Width           =   855
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
      Height          =   800
      Left            =   8610
      Picture         =   "ASIDIAR.frx":1BB0
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Abandona la Aplicación"
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Restore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8610
      Picture         =   "ASIDIAR.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Restaura Valores Anteriores del Asiento"
      Top             =   1890
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8610
      Picture         =   "ASIDIAR.frx":2004
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Aprueba y Graba el Asiento"
      Top             =   1050
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
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
      Left            =   210
      TabIndex        =   40
      Top             =   840
      Width           =   4110
   End
   Begin VB.Line Line5 
      X1              =   8275
      X2              =   8280
      Y1              =   4305
      Y2              =   4725
   End
   Begin VB.Line Line3 
      X1              =   8275
      X2              =   8275
      Y1              =   1365
      Y2              =   1680
   End
   Begin VB.Line Line2 
      X1              =   8265
      X2              =   8265
      Y1              =   1365
      Y2              =   1785
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   180
      X2              =   8275
      Y1              =   4725
      Y2              =   4725
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Asiento de Diario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   210
      TabIndex        =   33
      Top             =   315
      Width           =   5475
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Número:"
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
      Left            =   5145
      TabIndex        =   26
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
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
      Left            =   5145
      TabIndex        =   18
      Top             =   315
      Width           =   1275
   End
End
Attribute VB_Name = "ASIDIAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FEAS%
Dim SUDEBE#, SUHABE#
Dim CTA%(1024), IMPA#(1024), DETA$(1024)
Dim ACTEDIT%
Dim PRIMO%, PRIMOL&
'
Private Sub Command1_Click()
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   '
   NROAS = (Int(10 * Val(Text2.TEXT) + 0.5)) / 10
   FEAS% = FECHANUM(Text1.TEXT)
   REFE$ = Text4.TEXT
   CONTA% = List1.ListCount
   SUMATOT# = 0
   '
   If TIASI% = 0 Then
      If POSEDI% < 0 Then
         TIASORI% = Asc(Left$(E8$, 1))
         If TIASORI% = 1 Or TIASORI% = 2 Then
            If List1.ListCount < 1 Then
               GoTo 10
            End If
         End If
         Call MENSERR(24, "Imposible Grabar este Asiento")
         Exit Sub
      End If
   End If
   '
   If CONTROL$("PROPSCG", "CONTAGEN") = "GUEST" Then
      Call MENSERR(24, "Sistema Distribuido - Invitado en Contabilidad General\Imposible Grabar Este Asiento.")
      Exit Sub
   End If
   '
10 For U& = 1 To List1.ListCount
      Z$ = List1.List(U& - 1)
      CTA%(U&) = CUEINT%(Left$(Z$, 12))
      If CTA%(U&) < 1 Or CTA%(U&) > NUCUEN% Then
         Call MENSERR(24, "Imposible Grabar este Asiento.\  \Número de Cuenta Incorrecto (" + TRIM$(Left$(Z$, 12)) + ")\en Renglon " + TRIM$(Str$(U&)) + ".")
         Exit Sub
      End If
      IMPA#(U&) = Val(Mid$(Z$, 41, 12)) - Val(Mid$(Z$, 53, 12))
      SUMATOT# = SUMATOT# + IMPA#(U&)
      DETA$(U&) = Mid$(Z$, 14, 27)
      If TRIM$(DETA$(U&)) = TRIM$(Left$(DENOCUE$(CTA%(U&)), 27)) Then
         DETA$(U&) = ""
      End If
   Next U&
   '
   If Abs(SUMATOT#) >= 0.005 Then
      Call MENSERR(24, "Imposible Grabar este Asiento.\No Balancea.")
      Exit Sub
   End If
   '
   Call BLOQARCH("ACONT" + EXTE$)
   '
   If YAESTABAS% = 0 Then
     LI% = 0: LS% = MAXENCDIARIO% + 1
82   E% = Int((LS% - LI%) / 2)
     If E% < 1 Then GoTo 85
     L% = LI% + E%
     If L% < 1 Or L% > MAXENCDIARIO% Then GoTo 85
     Call LEENCDIARIO(L%)
     If CVS(E1$) < NROAS - 0.01 Then LI% = L%: GoTo 82
     If CVS(E1$) > NROAS + 0.01 Then LS% = L%: GoTo 82
     '
     Call MENSERR(24, "Imposible Grabar este Asiento.\Número Ocupado.")
     Exit Sub
   End If
   '
85 Call GRASIDIAR(NROAS, FEAS%, REFE$, CONTA%, CTA%(), IMPA#(), DETA$())
   If TIASI% = 3 Then
      Call MARCASCAJA
   End If
   '
   If CONTROL$("ASICONTA", "IMPRECOM") <> "NO" Then
      Call PRINASIDIAR
   End If
   '
   Call CLEARASIDIAR
   If TIASI% > 0 Then
      Unload ASIDIAR
   End If
   '
End Sub

Private Sub Command11_Click()
   Call AGREPASE(1)
End Sub

Private Sub Command12_Click()
   Call AGREPASE(2)
End Sub

Private Sub Command3_Click()
   If TIASI% > 0 Then command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then command1.SetFocus: Exit Sub
   Call CLEARASIDIAR
End Sub

Private Sub Command4_Click()
   Call BORRAPASE
End Sub

Private Sub Command5_Click()
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   'PRINTFORM
   Call PRINASIDIAR
End Sub

Private Sub Command6_Click()
   Call AGREPASE(0)
End Sub

Private Sub Command7_Click()
   If EXISTE%("C:\WINDOWS\CALC.EXE") > 0 Then
        Shell "C:\WINDOWS\CALC.EXE", 4
      Else
        Call MENSERR(24, "No se Encuentra Utilitario\'CALC.EXE'\en Carpeta 'WINDOWS' del Disco 'C:'.")
   End If
End Sub

Private Sub Command8_Click()
   Call EDITPASE
End Sub

'
Private Sub Form_Load()
    '
    Caption = Caption + " - " + TRIM$(EJACT$)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    If TIASI% = 0 Then
         VV% = VALIFECHAS%
         Label3.Caption = "(Asiento Manual)"
      ElseIf TIASI% = 1 Then
         Label3.Caption = "(Resumen Mensual de Ventas)"
         Call ASIVENTAS
         If List1.ListCount < 1 Then
            Call MENSERR(24, "No hay Movimientos\en el Período Elegido")
            Text1.TEXT = "********"
         End If
      ElseIf TIASI% = 2 Then
         Label3.Caption = "(Resumen Mensual de Compras)"
         Call ASICOMPRAS
         If List1.ListCount < 1 Then
            Call MENSERR(24, "No hay Movimientos\en el Período Elegido")
            Text1.TEXT = "********"
         End If
      ElseIf TIASI% = 3 Then
         Label3.Caption = "(Resumen Mensual de Caja-Bancos)"
         Call RECAJABAN
         If List1.ListCount < 1 Then
            Call MENSERR(24, "No hay Movimientos\en el Período Elegido")
            Text1.TEXT = "********"
         End If
      ElseIf TIASI% = 4 Then
         Label3.Caption = "(Asiento de Apertura del Ejercicio)"
         If COMEJ% < 1 Then
            Call MENSERR(24, "Falta Abrir Ejercicio Contable")
            Text1.TEXT = "********"
            Exit Sub
         End If
         FEX = COMEJ%: Text1.TEXT = FECHATEX$(FEX)
         Text2.TEXT = "1.0"
         Call CARGASIDIAR
    End If
    '
    ACTEDIT% = 0
    '
End Sub


Private Sub Form_Unload(Cancel As Integer)
    TIASI% = 0
    Call CIERRACONTA
    Call LIBARCH("ASICONT")
End Sub

Private Sub Label11_Click()
   OPASICONT.Show 1
End Sub

Private Sub List1_DblClick()
   Call EDITPASE
End Sub

Private Sub Text1_GotFocus()
   If TIASI% = 4 Then
      Text4.SetFocus
      Exit Sub
   End If
   '
   If TIASI% > 0 Then
      If Text1.TEXT = "********" Then
         Unload ASIDIAR
         Exit Sub
      End If
   End If
   '
   If TIASI% > 0 And TIASI% <> 4 Then command1.SetFocus: Exit Sub
   If ACTEDIT% = 1 Then Text11.SetFocus: Exit Sub
   VV% = VALIFECHAS%
   Text1.SelStart = 0
   Text1.SelLength = 8
End Sub
'
Private Sub Text1_DblClick()
   If TIASI% > 0 And TIASI% <> 4 Then command1.SetFocus: Exit Sub
   If TIASI% = 4 Then Exit Sub
   If POSEDI% < 0 Then command1.SetFocus: Exit Sub
   If ACTEDIT% = 1 Then Text11.SetFocus: Exit Sub
   If FECHANUM(Text1.TEXT) < 1 Then
      HOII% = HOY(HOS$)
      Text1.TEXT = HOS$
   End If
   Call SELECHO("SELFECHA")
   Text1.TEXT = VALACT1$("SELFECHA")
   If VALIFECHAS% > 0 Then
         NPA = NUPROAS
         Text2.SetFocus
      Else
         Call MENSERR(24, "Fecha no Válida")
         Text1.SetFocus
   End If
End Sub
'
Private Sub COMMAND21_CLICK()
   If TIASI% > 0 And TIASI% <> 4 Then command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then command1.SetFocus: Exit Sub
   If FECHANUM(Text1.TEXT) < 1 Then
      HOII% = HOY(HOS$)
      Text1.TEXT = HOS$
   End If
   Call SELECHO("SELFECHA")
   Text1.TEXT = VALACT1$("SELFECHA")
   If VALIFECHAS% > 0 Then
         NPA = NUPROAS
         Text2.SetFocus
      Else
         Call MENSERR(24, "Fecha no Válida")
         Text1.SetFocus
   End If
End Sub
'
Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If VALIFECHAS% > 0 Then
           NPA = NUPROAS
           Text2.SetFocus
         Else
           Call MENSERR(24, "Fecha no Válida")
           Text1.SetFocus
      End If
   End If
End Sub
'

Private Sub Text1_LostFocus()
   If TIASI% > 0 Then Exit Sub
   If VALIFECHAS% > 0 Then
        NPA = NUPROAS
        'Text2.SetFocus
      Else
        Call MENSERR(24, "Fecha no Válida")
        On Error Resume Next
        Text1.SetFocus
        On Error GoTo 0
   End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 27 Then
      If List1.ListCount > 0 Then
         Call SUMASIEN
         If Abs(SUDEBE# - SUHABE#) < 0.005 Then
            If SUDEBE# >= 0.005 Then
               command1.SetFocus
            End If
         End If
      End If
      Command9.SetFocus
   End If
   '
   If KeyAscii% = 13 Then
      Text12.SetFocus
   End If
   '
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 27 Then
      If List1.ListCount > 0 Then
         Call SUMASIEN
         If Abs(SUDEBE# - SUHABE#) < 0.005 Then
            If SUDEBE# >= 0.005 Then
               command1.SetFocus
            End If
         End If
      End If
      Command9.SetFocus
   End If
   '
   If KeyAscii% = 13 Then
      Text14.SetFocus
   End If
   '
End Sub

Private Sub Text14_LostFocus()
   Text14.TEXT = CSTRING$(MKD$(Val(Text14.TEXT)), 3, 11, 2, 2)
End Sub

Private Sub Text15_LostFocus()
   Text15.TEXT = CSTRING$(MKD$(Val(Text15.TEXT)), 3, 11, 2, 2)
End Sub

Private Sub Text2_Change()
   List1.Clear
   YAESTABAS% = 0
   Text11.TEXT = ""
   Label4.Caption = ""
   Text12.TEXT = ""
   Text4.TEXT = Space$(40)
   Text14.TEXT = ""
   Text15.TEXT = ""
   Label21.Caption = ""
   Label22.Caption = ""
   POSEDI% = 0
End Sub

Private Sub Text2_GotFocus()
   If TIASI% > 0 And TIASI% <> 4 Then command1.SetFocus: Exit Sub
   If TIASI% = 4 Then Text4.SetFocus
   If ACTEDIT% = 1 Then Text11.SetFocus: Exit Sub
   VV% = VALIFECHAS%
   Text2.SelStart = 0
   Text2.SelLength = 8
End Sub
'
Private Sub Text14_GotFocus()
   If TIASI% > 0 And TIASI% <> 4 Then command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then command1.SetFocus: Exit Sub
   If VALICUENTA% <> 1 Then Exit Sub
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   If Abs(Val(Text14.TEXT)) < 0.005 Then
      Text14.TEXT = "       0.00"
   End If
   Text14.SelStart = 0
   Text14.SelLength = Len(Text14.TEXT)
End Sub

Private Sub Text15_GotFocus()
   If TIASI% > 0 And TIASI% <> 4 Then command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then command1.SetFocus: Exit Sub
   If VALICUENTA% <> 1 Then Exit Sub
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   If Abs(Val(Text15.TEXT)) < 0.005 Then
      Text15.TEXT = "       0.00"
   End If
   Text15.SelStart = 0
   Text15.SelLength = Len(Text14.TEXT)
End Sub
Private Sub Text14_Change()
   If Abs(Val(Text14.TEXT)) >= 0.005 Then
      Text15.TEXT = "       0.00"
   End If
End Sub

Private Sub Text15_Change()
   If Abs(Val(Text15.TEXT)) >= 0.005 Then
      Text14.TEXT = "       0.00"
   End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 27 Then
      If List1.ListCount > 0 Then
         Call SUMASIEN
         If Abs(SUDEBE# - SUHABE#) < 0.005 Then
            If SUDEBE# >= 0.005 Then
               command1.SetFocus
            End If
         End If
      End If
      Command9.SetFocus
   End If
   '
   If KeyAscii% = 13 Then
      If Abs(Val(Text14.TEXT)) > 0 Then
           Call AGREPASE(0)
         Else
           Text15.SetFocus
      End If
   End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 27 Then
      If List1.ListCount > 0 Then
         Call SUMASIEN
         If Abs(SUDEBE# - SUHABE#) < 0.005 Then
            If SUDEBE# >= 0.005 Then
               command1.SetFocus
            End If
         End If
      End If
      Command9.SetFocus
   End If
   '
   If KeyAscii% = 13 Then
      If Abs(Val(Text15.TEXT)) > 0 Then
           Call AGREPASE(0)
         Else
           Text14.SetFocus
      End If
   End If
End Sub

Private Sub Command2_Click()
    Call HELPONLINE("ASIDIAR")
End Sub
'

Private Sub Command9_Click()
    '
    FEX = FEPROASDIA
    Text1.TEXT = FECHATEX$(FEX)
    Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
    List1.Clear
    YAESTABAS% = 0
    Text11.TEXT = ""
    Label4.Caption = ""
    Text12.TEXT = ""
    Text4.TEXT = Space$(40)
    Text14.TEXT = ""
    Text15.TEXT = ""
    Call SUMASIEN
    Call CIERRARCH("*.*")
    Unload ASIDIAR
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 13 Then
      '
      If VALINUMAS% <> 1 Then
         Text2.SetFocus
         Exit Sub
      End If
      '
      Text4.SetFocus
      '
   End If

End Sub

Private Sub Text2_LostFocus()
   '
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   '
   Call CARGASIDIAR
   On Error Resume Next
   Text4.SetFocus
   On Error GoTo 0
   '
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text11.SetFocus
   End If
End Sub

Private Sub Text4_GOTFOCUS()
   If TIASI% > 0 And TIASI% <> 4 Then command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then command1.SetFocus: Exit Sub
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   Text4.SelStart = 0
   Text4.SelLength = 40
End Sub

Private Sub Text11_DBLCLICK()
   If TIASI% > 0 And TIASI% <> 4 Then command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then command1.SetFocus: Exit Sub
   RG1$ = REGSEL$("ECUECON")
   If Len(RG1$) > 4 Then
      Text11.TEXT = Left$(RG1$, 12)
      Label4.Caption = Mid$(RG1$, 13, 28)
      Text12.SetFocus
   End If
End Sub
'
Private Sub Text11_CHANGE()
   'If VALICUENTA% <> 1 Then Exit Sub
End Sub
Private Sub COMMAND31_CLICK()
   If TIASI% > 0 And TIASI% <> 4 Then command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then command1.SetFocus: Exit Sub
   Text11.SetFocus
   RG1$ = REGSEL$("ECUECON")
   If Len(RG1$) > 4 Then
      Text11.TEXT = Left$(RG1$, 12)
      Label4.Caption = Mid$(RG1$, 13, 28)
      Text12.SetFocus
   End If
End Sub
'
Private Sub Text11_GotFocus()
   If TIASI% > 0 And TIASI% <> 4 Then command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then command1.SetFocus: Exit Sub
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   '
   Text11.SelStart = 0
   Text11.SelLength = 12
End Sub
'
Private Sub Text12_GotFocus()
   If TIASI% > 0 And TIASI% <> 4 Then command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then command1.SetFocus: Exit Sub
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   '
   If VALICUENTA% <> 1 Then Exit Sub
   '
   If Text12.TEXT = "" Then
     Text12.TEXT = Label4.Caption
   End If
   '
   Text12.SelStart = 0
   Text12.SelLength = 40
End Sub
'

Private Sub VScroll1_Change()
   If TIASI% > 0 Then Exit Sub
   Static VALANT%
   If VALANT% = 0 Then VALANT% = 16384
   NROACT = Val(Text2.TEXT)
   '
   If NROACT > ULTASDIA + 0.01 Then
      RGENC% = MAXENCDIARIO% + 1
      GoTo 85
   End If
   '
   LI% = 0: LS% = MAXENCDIARIO% + 1
82 E% = Int((LS% - LI%) / 2)
   If E% < 1 Then GoTo 83
   L% = LI% + E%
   If L% < 1 Then RGENC% = 1: GoTo 85
   If L% > MAXENCDIARIO% Then RGENC% = MAXENCDIARIO% + 1: GoTo 85
   Call LEENCDIARIO(L%)
   If CVS(E1$) < NROACT - 0.01 Then LI% = L%: GoTo 82
   If CVS(E1$) > NROACT + 0.01 Then LS% = L%: GoTo 82
83 RGENC% = L%
   '
85 If VScroll1.Value < VALANT% Then
        RGENC% = RGENC% + 1
      ElseIf VScroll1.Value > VALANT% Then
        RGENC% = RGENC% - 1
   End If
   VALANT% = VScroll1.Value
   If RGENC% > 0 Then
      If RGENC% <= MAXENCDIARIO% Then
           Call LEENCDIARIO(RGENC%)
           NROAS = CVS(E1$)
           Text2.TEXT = CSTRING$(MKS$(NROAS), 3, 8, 1, 2)
           Call CARGASIDIAR
         Else
           Text1.TEXT = ""
           VV% = VALIFECHAS%
           Text4.SetFocus
      End If
   End If
End Sub
'
Function VALIFECHAS%()
   '
   FEAS% = FECHANUM(Text1.TEXT)
   FEASA% = FEAS%
   If FEAS% < COMEJ% Or FEAS% > FINEJ% Then
      FEX = FEPROASDIA%
      FEAS% = FEX
      Text1.TEXT = FECHATEX$(FEX)
      Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
   End If
   '
   If FEAS% = FEASA% Then
       VALIFECHAS% = 1
     Else
       VALIFECHAS% = 0
   End If
   '
End Function

Function VALINUMAS%()
   '
   Static VNA%, NROAA
   
   VALINUMAS% = 0
   NROAS = (Int(10 * Val(Text2.TEXT) + 0.5)) / 10
   If NROAS = NROAA Then
      VALINUMAS% = VNA%
      Text2.TEXT = CSTRING$(MKS$(NROAS), 3, 8, 1, 2)
      Exit Function
   End If
   '
   NROAA = NROAS
   Text2.TEXT = CSTRING$(MKS$(NROAS), 3, 8, 1, 2)
   If NROAS < 1 Then GoTo 59
   If NROAS > 9999.9 Then GoTo 59
   VALINUMAS% = 1: VNA% = 1
   Exit Function
   '
59 FEX = FEPROASDIA%
   FEAS% = FEX
   Text1.TEXT = FECHATEX$(FEX)
   Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
   VNA% = 0
   '
End Function
'
Function VALICUENTA%()
   '
   CUE$ = TRIM$(Text11.TEXT)
   If CUE$ <> "" Then
      If ECOARCH("ECUECON", AJUSTI$(CUE$, 12)) <> "" Then
         Label4.Caption = ECOARCH("ECUECON", AJUSTI$(Text11.TEXT, 12))
         VALICUENTA% = 1
         Exit Function
      End If
   End If
   Label4.Caption = ""
   VALICUENTA% = 0
   Text11.SetFocus
   '
End Function
'
Function REGENCAS%()
   '
   REGENCAS% = 0
   VV% = VALINUMAS%
   NRO = Val(Text2.TEXT)
   LI% = 0: LS% = MAXENCDIARIO% + 1
   L% = MAXENCDIARIO%: ERRO% = 0
   If NRO > ULTASDIA + 0.01 Then
      FEASA = FECHANUM(Text1.TEXT)
      If FEASA < FULTASDIA% Then
         FEASA = FULTASDIA%
      End If
   GoTo 99
   End If
   '
82 E% = Int((LS% - LI%) / 2)
   If E% < 1 Then GoTo 98
   L% = LI% + E%
   If L% < 1 Or L% > MAXENCDIARIO% Then GoTo 98
   Call LEENCDIARIO(L%)
   If CVS(E1$) < NRO - 0.01 Then LI% = L%: GoTo 82
   If CVS(E1$) > NRO + 0.01 Then LS% = L%: GoTo 82
   REGENCAS% = L%
   Exit Function
   '
98 FEASA = FECHANUM(Text1.TEXT)
   If FEASA < COMEJ% Then FEASA = COMEJ%
   If LI% > 0 Then
      Call LEENCDIARIO(LI%)
      If FEASA < CVI(E3$) Then FEASA = CVI(E3$)
   End If
   If LS% < LI% + 1 Then LS% = LI% + 1
   If LS% <= MAXENCDIARIO% Then
      Call LEENCDIARIO(LS%)
      If FEASA > CVI(E3$) Then FEASA = CVI(E3$)
   End If
   '
99 Text1.TEXT = FECHATEX$(FEASA)
   '
End Function

Function NUPROAS()
   '
   NUPROAS = 0
   If VALIFECHAS% <> 1 Then
      Exit Function
   End If
   '
   FEAS% = FECHANUM(Text1.TEXT)
   FEASPE = FEAS%
   LI% = 0: LS% = MAXENCDIARIO% + 1
   L% = MAXENCDIARIO%: ERRO% = 0
   If FEAS% >= FULTASDIA% Then
      NUPROAS = Int(1 + ULTASDIA)
      Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
      Exit Function
   End If
   '
81 E% = Int((LS% - LI%) / 2)
   If E% < 1 Then L% = LS%: GoTo 82
   L% = LI% + E%
   If L% < 1 Or L% > MAXENCDIARIO% Then GoTo 99
   Call LEENCDIARIO(L%)
   If CVI(E3$) < FEAS% Then LI% = L%: GoTo 81
   If CVI(E3$) > FEAS% Then LS% = L%: GoTo 81
   
82 For L1% = L% To 1 Step -1
      Call LEENCDIARIO(L1%)
     If CVI(E3$) < FEAS% Then GoTo 83
   Next L1%
   L1% = 0
   '
83 L2% = L1%
   NASPROX = 0
   For L% = L1% + 1 To MAXENCDIARIO%
      Call LEENCDIARIO(L%)
      If CVI(E3$) > FEAS% Then GoTo 84
      L2% = L%
   Next L%
   '
84 If L2% < 1 Then
      Call LEENCDIARIO(1)
      NUPROAS = CVS(E1$)
      GoTo 85
   End If
   '
   Call LEENCDIARIO(L2%)
   NUPROAS = (1 + Int(10 * CVS(E1$) + 0.5)) / 10
   If L2% < MAXENCDIARIO% Then
      Call LEENCDIARIO(L2% + 1)
      If Abs(CVS(E1$) - NUPROAS) < 0.02 Then
         GoTo 99
      End If
   End If
   '
85 Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
   RGENC% = REGENCAS%
   If RGENC% > 0 Then
      Call LEENCDIARIO(RGENC%)
      If CVI(E2$) > 0 And CVI(E2$) < FEAS% Then GoTo 99
      NUPROAS = (Int(10 * NUPROAS + 0.5) - 1) / 10
      GoTo 85
   End If
   If NUPROAS < 1.05 Then GoTo 99
   FEX = FEAS%: Text1.TEXT = FECHATEX$(FEX)
   Exit Function
   '
99 Text1.TEXT = FECHATEX$(FEASPE)
   If FORMALTER.Visible = True Then Exit Function
   '
   TTXX$ = "Imposible Registrar Asiento\en la Fecha Elegida."
   TTXX$ = TTXX$ + "\No hay Números Libres en esa Fecha."
   TTXX$ = TTXX$ + "\  \Puede Cancelar o Re-Numerar Todos los Asientos\Posteriores al " + FECHATEX$(FEASPE) + "."
   OPXX% = COMALTER%(TTXX$ + "\\Cancelar\Re-Numerar Ahora")
   If OPXX% = 2 Then
      Call RENUMERAS(FEASPE)
   End If
   Call FINAL("")
   '
End Function

Sub RENUMERAS(FEASPE)
   Screen.MousePointer = 11
   Call ABRENCDIARIO
   FINALIX% = MAXENCDIARIO%
   For UKI% = 1 To FINALIX%
     Get #NX9%, UKI%, E0$
     NUASAN = CVS(Mid$(E0$, 1, 4))
     FEANTE% = CVI(Mid$(E0$, 7, 2))
     If FEANTE% > FEASPE + 0.1 Then
       NUASNUE = (10 + Int(10 * NUASAN + 0.5)) / 10
       Mid$(E0$, 1, 4) = MKS$(NUASNUE)
       Put #NX9%, UKI%, E0$
     End If
   Next UKI%
   '
   Call ABREDIARIO
   FINALIX% = MAXDIARIO%
   For UKI% = 1 To FINALIX%
     Get #NX10%, UKI%, E10$
     NUASAN = CVS(Mid$(E10$, 3, 4))
     FEANTE% = CVI(Mid$(E10$, 11, 2))
     If FEANTE% > FEASPE + 0.1 Then
       NUASNUE = (10 + Int(10 * NUASAN + 0.5)) / 10
       Mid$(E10$, 3, 4) = MKS$(NUASNUE)
       Put #NX10%, UKI%, E10$
     End If
   Next UKI%
   '
   Close #NX9%: NX9% = 0
   Close #NX10%: NX10% = 0
   Screen.MousePointer = 1
   Call COMUNI("Re-Numeración Completada.\Re-Inicie Ahora el Sistema de Contabilidad")
   Call FINAL("")
   '
End Sub
'
Sub CARGASIDIAR()
   '
   POSEDI% = 0
   RGE% = REGENCAS%
   If RGE% > 0 Then
      NROAS = Val(Text2.TEXT)
      Call LEENCDIARIO(RGE%)
      FF% = CVI(E3$)
      CONTA% = CVI(E6$)
      FEX = FF%
      Text1.TEXT = FECHATEX$(FEX)
      Text4.TEXT = E7$
      '
81    ERRO% = 0
      LI% = 0: LS% = MAXDIARIO% + 1
82    E% = Int((LS% - LI%) / 2)
      If E% < 1 Then
         NABORX$ = Val(Text2.TEXT)
         MENSA$ = "Se ha Detectado un Error de Consistencia\en el Asiento Numero '" + TRIM$(NABORX$) + "'.\  \"
         MENSA$ = MENSA$ + "FLEXOFT Puede Recuperar la Consistencia\de la Base de Datos, Borrando la Cabecera\de este Asiento.\\"
         If COMALTER%(MENSA$ + "Cancelar\Borrar Cabecera") = 2 Then
            Call BLOQARCH("ACONT" + EXTE$)
            Call BLOQARCH("ASICONT")
            RGX% = 0
            FINX% = MAXENCDIARIO%
            NABOR = Val(Text2.TEXT)
            Call ABRENCDIARIO
            For UI% = 1 To FINX%
               Get #NX9%, UI%, E0$
               NOAS = CVS(Left$(E0$, 4))
'MsgBox NOAS & "   " & NABOR & "   " & (NOAS - NABOR)
               If Abs(NOAS - NABOR) > 0.01 Then
                  RGX% = RGX% + 1
                  Put #NX9%, RGX%, E0$
               End If
            Next UI%
            '
            For UI% = RGX% + 1 To FINX%
               LSet E0$ = String$(64, 0)
               Put #NX9%, UI%, E0$
            Next UI%
            '
            Close #NX9%: NX9% = 0
            Call COMUNI("Recuperación Completa.\Re-Inicie Ahora el Sistema Contable.")
         End If
         Call FINAL("")
      End If
      L% = LI% + E%
      If L% < 1 Or L% > MAXDIARIO% Then
         ERRO% = 1: GoTo 89
      End If
      Call LEEDIARIO(L%)
      If CVS(E12$) < NROAS - 0.01 Then LI% = L%: GoTo 82
      If CVS(E12$) > NROAS + 0.01 Then LS% = L%: GoTo 82
      '
      For K1% = L% To 1 Step -1
         Call LEEDIARIO(K1%)
         If Abs(CVS(E12$) - NROAS) > 0.01 Then GoTo 83
         REGINI% = K1%
      Next K1%
      '
83    List1.Clear
      YAESTABAS% = 0
      For K1% = 1 To CONTA%
         Call LEEDIARIO(REGINI% + K1% - 1)
         If Abs(CVS(E12$) - NROAS) > 0.01 Or CVI(E15$) <> FF% Then
            ERRO% = 1: GoTo 89
         End If
         IPASE# = CVD(E16$)
         If CVS(E19$) > 0.01 Then POSEDI% = (-1)
         ' la condic. siguiente modificada 6/2/04 por asiento apertura VDR
         If Asc(Left$(E8$, 1)) > 0 And NROAS < 1.05 Then POSEDI% = (-1)
         If CVI(E2$) > 0 Then POSEDI% = (-1)
         '
         Z$ = Space$(128)
         Call REPLA(Z$, CODCUE$(CVI(E11$)), 1, 12)
         DEPASE$ = E00$: If TRIM$(DEPASE$) = "" Then DEPASE$ = DENOCUE$(CVI(E11$))
         Call REPLA(Z$, DEPASE$, 14, 27)
         If IPASE# > 0 Then
              Call REPLA(Z$, CSTRING$(MKD$(IPASE#), 3, 12, 2, 2), 41, 12)
            Else
              Call REPLA(Z$, CSTRING$(MKD$(Abs(IPASE#)), 3, 12, 2, 2), 53, 12)
         End If
         List1.AddItem Z$
         YAESTABAS% = 1
      Next K1%
      List1.ListIndex = (-1)
      '
      If ERRO% > 0 Then
         CONTA% = 0
         Text1.TEXT = ""
         VV% = VALIFECHAS%
         Exit Sub
      End If
      '
      If REGINI% > 1 Then
         Call LEEDIARIO(REGINI% - 1)
         If CVS(E12$) >= NROAS - 0.01 Then ERRO% = 1: GoTo 89
      End If
      '
      Call LEEDIARIO(REGINI% + CONTA%)
      If CVI(E11$) <> 0 Then
         If CVS(E12$) <= NROAS + 0.01 Then ERRO% = 1
         If Abs(CVS(E12$) - NROAS) < 0.01 Then ERRO% = 2
      End If
89    If ERRO% > 0 Then
         Call MENSERR(24, "Error de Consistencia de Archivos.\  \Consulte a su Soporte de Sistemas FLEXOFT.")
         If ERRO% = 2 Then
            CONTA% = CONTA% + 1
            GoTo 81
         End If
         Call FINAL("")
      End If
      CONTANT% = CONTA%
     '
   End If
   If TIASI% = 4 Then POSEDI% = 0
   Call SUMASIEN
   '
End Sub
'
Sub CLEARASIDIAR()
   '
   FEX = FEPROASDIA%
   FEAS% = FEX
   Text1.TEXT = FECHATEX$(FEX)
   Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
   List1.Clear
   YAESTABAS% = 0
   Text11.TEXT = ""
   Label4.Caption = ""
   Text12.TEXT = ""
   Text4.TEXT = Space$(40)
   Text14.TEXT = ""
   Text15.TEXT = ""
   Call SUMASIEN
   ACTEDIT% = 0
   '
   Text1.SetFocus
   '
End Sub

Sub AGREPASE(MODO%)
   '
   If TIASI% > 0 And TIASI% <> 4 Then command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then command1.SetFocus: Exit Sub
   If VALICUENTA% <> 1 Then
      Exit Sub
   End If
   '
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   '
   If Val(Text14.TEXT) < 0.005 Then
      If Val(Text15.TEXT) < 0.005 Then
         Text14.SetFocus
         Exit Sub
      End If
   End If
   '
   REAC% = List1.ListIndex
   If REAC% < 0 Or REAC% >= List1.ListCount Then
      MODO% = 1
   End If
   If MODO% = 0 Then
      If REAC% >= 0 Then
         If REAC% < List1.ListCount Then
            MODO% = 2
         End If
      End If
   End If
   '
   Z$ = Space$(128)
   Call REPLA(Z$, Text11.TEXT, 1, 12)
   DEPASE$ = Text12.TEXT: If TRIM$(DEPASE$) = "" Then DEPASE$ = DENOCUE$(CUEINT%(Text11.TEXT))
   Call REPLA(Z$, DEPASE$, 14, 27)
   Call REPLA(Z$, CSTRING$(MKD$(Val(Text14.TEXT)), 3, 12, 2, 2), 41, 12)
   Call REPLA(Z$, CSTRING$(MKD$(Val(Text15.TEXT)), 3, 12, 2, 2), 53, 12)
   '
   If MODO% = 1 Then
        List1.AddItem Z$
        List1.ListIndex = (-1)
        If List1.TopIndex < List1.ListCount - 9 Then
           List1.TopIndex = List1.ListCount - 9
        End If
      ElseIf MODO% = 2 Then
        List1.AddItem Z$, REAC%
   End If
   '
   Call SUMASIEN
   Text14.TEXT = Label8.Caption
   Text15.TEXT = Label10.Caption
   ACTEDIT% = 0
   Command4.Enabled = True
   Command8.Enabled = True
   '
   'If InStr(EMPREAC$, "DOSIVA") > 0 Then
     Text12.TEXT = ""
   'End If
   '
End Sub

Sub EDITPASE()
   If TIASI% > 0 And TIASI% <> 4 Then command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then command1.SetFocus: Exit Sub
   If ACTEDIT% > 0 Then
      Exit Sub
   End If
   '
   U& = List1.ListIndex
   If U& >= 0 Then
      If U& < List1.ListCount Then
         Z$ = List1.List(U&)
         Text11.TEXT = Left$(Z$, 12)
         VV% = VALICUENTA%
         Text12.TEXT = Mid$(Z$, 14, 27)
         IDEBE# = Val(Mid$(Z$, 41, 12))
         IHABE# = Val(Mid$(Z$, 53, 12))
         Text14.TEXT = CSTRING$(MKD$(IDEBE#), 3, 11, 2, 2)
         Text15.TEXT = CSTRING$(MKD$(IHABE#), 3, 11, 2, 2)
         List1.RemoveItem U&
         If U& < List1.ListCount Then
              List1.ListIndex = U&
            Else
              List1.ListIndex = (-1)
         End If
         ACTEDIT% = 1
         Command4.Enabled = False
         Command8.Enabled = False
         Call SUMASIEN
      End If
   End If
End Sub

Sub BORRAPASE()
   If TIASI% > 0 And TIASI% <> 4 Then command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then
      TIASORI% = Asc(Left$(E8$, 1))
      If TIASORI% = 1 Or TIASORI% = 2 Then
         If COMALTER%("Realmente Quiere Borrar Este Asiento ?\\No, Continuar\Si, Borrar") = 2 Then
            List1.Clear
            Label21.Caption = ""
            Label22.Caption = ""
            List1.Refresh
            Label21.Refresh
            Label22.Refresh
            Exit Sub
         End If
      End If
      command1.SetFocus: Exit Sub
   End If
   U& = List1.ListIndex
   If U& >= 0 Then
      If U& < List1.ListCount Then
         List1.RemoveItem U&
         If U& < List1.ListCount Then
              List1.ListIndex = U&
            Else
              List1.ListIndex = (-1)
         End If
         Call SUMASIEN
      End If
   End If
End Sub
'
Sub SUMASIEN()
   '
   SUDEBE# = 0: SUHABE# = 0
   For U& = 1 To List1.ListCount
      Z$ = List1.List(U& - 1)
      SUDEBE# = SUDEBE# + Val(Mid$(Z$, 41, 12))
      SUHABE# = SUHABE# + Val(Mid$(Z$, 53, 12))
   Next U&
   Label21.Caption = CSTRING$(MKD$(SUDEBE#), 3, 11, 2, 2)
   Label22.Caption = CSTRING$(MKD$(SUHABE#), 3, 11, 2, 2)
   '
   Label8.Caption = ""
   Label10.Caption = ""
   If SUDEBE# >= SUHABE# + 0.005 Then
        Label10.Caption = CSTRING$(MKD$(SUDEBE# - SUHABE#), 3, 11, 2, 2)
      ElseIf SUHABE# >= SUDEBE# + 0.005 Then
        Label8.Caption = CSTRING$(MKD$(SUHABE# - SUDEBE#), 3, 11, 2, 2)
   End If
   '
   On Error Resume Next
   Text11.SetFocus
   On Error GoTo 0
   '
End Sub

Sub ASIVENTAS()
   '
   Call FRESHALL
   Call CARSELMESA
10 Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) = "" Then
      Text1.TEXT = "********"
      Exit Sub
   End If
   '
   MMXX0$ = TRIM$(MESASEL$)
   MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
   If TRIM$(MMXX1$) = "" Then
      Text1.TEXT = "********"
      Exit Sub
   End If
   MESASEL$ = MMXX1$
   '
   If DES% < COMEJ% Or HAS% > FINEJ% Then
      Call MENSERR(24, "Período Fuera de Límites\del Ejercicio Activo.")
      GoTo 10
   End If
   '
   FEAS% = HAS%
   FEX = FEAS%: Text1.TEXT = FECHATEX$(FEX)
   FF$ = Text1.TEXT
   If VALIFECHAS% > 0 Then
         NPA = NUPROAS
      Else
         Call MENSERR(24, "Fecha no Válida")
         GoTo 10
   End If
   '
   REFE$ = AJUSTI$("RESUMEN MENSUAL DE VENTAS DE " + MESASEL$, 40)
   LI% = 0: LS% = MAXENCDIARIO% + 1
   L% = MAXENCDIARIO%: ERRO% = 0
   If FEAS% > FULTASDIA% Then
      Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
      GoTo 90
   End If
   '
82 E% = Int((LS% - LI%) / 2)
   If E% < 1 Then GoTo 90
   L% = LI% + E%
   If L% < 1 Or L% > MAXENCDIARIO% Then GoTo 90
   Call LEENCDIARIO(L%)
   If CVI(E3$) < FEAS% Then LI% = L%: GoTo 82
   If CVI(E3$) > FEAS% Then LS% = L%: GoTo 82
   '
   ENCUAS% = 0: NROAS = 0
   For L1% = L% To 1 Step -1
      Call LEENCDIARIO(L1%)
      If CVI(E3$) <> FEAS% Then GoTo 83
   Next L1%
   L1% = 0
   '
83 L1% = L1% + 1
   NASPROX = 0
   For L% = L1% To MAXENCDIARIO%
      Call LEENCDIARIO(L%)
      If CVI(E3$) <> FEAS% Then GoTo 84
      If TRIM$(E7$) = TRIM$(REFE$) Then              ' ENCONTRO EL ASIENTO
         ENCUAS% = 1
         NROAS = CVS(E1$)
         GoTo 84
      End If
      L2% = L%
   Next L%
   '
84 If ENCUAS% = 1 Then
      Text1.TEXT = FF$
      Text2.TEXT = CSTRING$(MKS$(NROAS), 3, 8, 1, 2)
      Text4.TEXT = REFE$
      VV% = VALINUMAS%
      YAESTABAS% = 1
      CMDX% = COMALTER%("Asiento de Ventas de " + MESASEL$ + "\ya Registrado.\\Mantener\Refrescar")
      If CMDX% <> 2 Then
           Call CARGASIDIAR
           Exit Sub
      End If
    Else
      Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
      Text4.TEXT = REFE$
   End If
   '
90 If Val(Text2.TEXT) < 1 Then
      Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
   End If
   Text4.TEXT = "RESUMEN MENSUAL DE VENTAS DE " + MESASEL$
   CONTA% = 0: SUMA# = 0
   List1.Clear: List2.Clear
   FIN& = ULTREG&("ASIDIAR")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      ASX$ = REGLEIDO$("ASIDIAR", U&)
      If Left$(ASX$, 4) = Chr$(1) + Chr$(TIASI%) + MKI$(FEAS%) Then
         If CVI(Mid$(ASX$, 5, 2)) > 0 Then
            CONTA% = CONTA% + 1
            IPASE# = CVD(Mid$(ASX$, 39, 8))
            SUMA# = SUMA# + (Int(100 * IPASE#) + 0.5) / 100
            '
            Z$ = Space$(128)
            Call REPLA(Z$, CODCUE$(CVI(Mid$(ASX$, 5, 2))), 1, 12)
            DEPASE$ = DENOCUE$(CVI(Mid$(ASX$, 5, 2)))
            Call REPLA(Z$, DEPASE$, 14, 27)
            If IPASE# > 0 Then
                 Call REPLA(Z$, CSTRING$(MKD$(IPASE#), 3, 12, 2, 2), 41, 12)
               Else
                 Call REPLA(Z$, CSTRING$(MKD$(Abs(IPASE#)), 3, 12, 2, 2), 53, 12)
            End If
            List2.AddItem "A" + Z$
         End If
      End If
   Next U&
   '
   For U& = 1 To List2.ListCount
      Z$ = Mid$(List2.List(U& - 1), 2)
      List1.AddItem Z$
   Next U&
   '
   Call SUMASIEN
   '
End Sub

Sub ASICOMPRAS()
   '
   Call FRESHALL
   Call CARSELMESA
10 Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) = "" Then
      Text1.TEXT = "********"
      Exit Sub
   End If
   '
   MMXX0$ = TRIM$(MESASEL$)
   MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
   If TRIM$(MMXX1$) = "" Then
      Text1.TEXT = "********"
      Exit Sub
   End If
   MESASEL$ = MMXX1$
   '
   If DES% < COMEJ% Or HAS% > FINEJ% Then
      Call MENSERR(24, "Período Fuera de Límites\del Ejercicio Activo.")
      GoTo 10
   End If
   '
   FEAS% = HAS%
   FEX = FEAS%: Text1.TEXT = FECHATEX$(FEX)
   FF$ = Text1.TEXT
   If VALIFECHAS% > 0 Then
         NPA = NUPROAS
      Else
         Call MENSERR(24, "Fecha no Válida")
         GoTo 10
   End If
   '
   REFE$ = AJUSTI$("RESUMEN MENSUAL DE COMPRAS DE " + MESASEL$, 40)
   LI% = 0: LS% = MAXENCDIARIO% + 1
   L% = MAXENCDIARIO%: ERRO% = 0
   If FEAS% > FULTASDIA% Then
      Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
      GoTo 90
   End If
   '
82 E% = Int((LS% - LI%) / 2)
   If E% < 1 Then GoTo 90
   L% = LI% + E%
   If L% < 1 Or L% > MAXENCDIARIO% Then GoTo 90
   Call LEENCDIARIO(L%)
   If CVI(E3$) < FEAS% Then LI% = L%: GoTo 82
   If CVI(E3$) > FEAS% Then LS% = L%: GoTo 82
   '
   ENCUAS% = 0: NROAS = 0
   For L1% = L% To 1 Step -1
      Call LEENCDIARIO(L1%)
      If CVI(E3$) <> FEAS% Then GoTo 83
   Next L1%
   L1% = 0
   '
83 L1% = L1% + 1
   NASPROX = 0
   For L% = L1% To MAXENCDIARIO%
      Call LEENCDIARIO(L%)
      If CVI(E3$) <> FEAS% Then GoTo 84
      If TRIM$(E7$) = TRIM$(REFE$) Then              ' ENCONTRO EL ASIENTO
         ENCUAS% = 1
         NROAS = CVS(E1$)
         GoTo 84
      End If
      L2% = L%
   Next L%
   '
84 If ENCUAS% = 1 Then
      Text1.TEXT = FF$
      Text2.TEXT = CSTRING$(MKS$(NROAS), 3, 8, 1, 2)
      Text4.TEXT = REFE$
      VV% = VALINUMAS%
      YAESTABAS% = 1
      CMDX% = COMALTER%("Asiento de Compras de " + MESASEL$ + "\ya Registrado.\\Mantener\Refrescar")
      If CMDX% <> 2 Then
           Call CARGASIDIAR
           Exit Sub
      End If
    Else
      Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
      Text4.TEXT = REFE$
   End If
   '
90 If Val(Text2.TEXT) < 1 Then
      Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
   End If
   Text4.TEXT = "RESUMEN MENSUAL DE COMPRAS DE " + MESASEL$
   CONTA% = 0: SUMA# = 0
   List1.Clear: List2.Clear
   FIN& = ULTREG&("ASIDIAR")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      ASX$ = REGLEIDO$("ASIDIAR", U&)
      If Left$(ASX$, 4) = Chr$(1) + Chr$(TIASI%) + MKI$(FEAS%) Then
         If CVI(Mid$(ASX$, 5, 2)) > 0 Then
            CONTA% = CONTA% + 1
            IPASE# = CVD(Mid$(ASX$, 39, 8))
            SUMA# = SUMA# + (Int(100 * IPASE#) + 0.5) / 100
            '
            Z$ = Space$(128)
            Call REPLA(Z$, CODCUE$(CVI(Mid$(ASX$, 5, 2))), 1, 12)
            DEPASE$ = DENOCUE$(CVI(Mid$(ASX$, 5, 2)))
            Call REPLA(Z$, DEPASE$, 14, 27)
            If IPASE# > 0 Then
                 Call REPLA(Z$, CSTRING$(MKD$(IPASE#), 3, 12, 2, 2), 41, 12)
               Else
                 Call REPLA(Z$, CSTRING$(MKD$(Abs(IPASE#)), 3, 12, 2, 2), 53, 12)
            End If
            List2.AddItem "A" + Z$
         End If
      End If
   Next U&
   '
   For U& = 1 To List2.ListCount
      Z$ = Mid$(List2.List(U& - 1), 2)
      List1.AddItem Z$
   Next U&
   '
   Call SUMASIEN
   '
End Sub

Sub RECAJABAN()
   '
   Dim CTA%(2048), IPAS#(2048)
10 Call SELECHO("SELFECHA")
   FF$ = VALACT1$("SELFECHA")
   If FF$ = "" Then
      Text1.TEXT = "********"
      Exit Sub
   End If
   '
   FEAS% = FECHANUM(FF$)
   If FEAS% < COMEJ% Or FEAS% > FINEJ% Or FEAS% > HOY(HOS$) Then
      Call MENSERR(24, "Fecha no Válida")
      Text1.TEXT = "********"
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   List1.Clear: List2.Clear
   Text1.TEXT = FF$
   '
   If VALIFECHAS% > 0 Then
         NPA = NUPROAS
      Else
         Call MENSERR(24, "Fecha no Válida")
         GoTo 10
   End If
   REFE$ = "RESUMEN DE CAJA-BANCOS AL " + FF$ + Space$(40)
   Text4.TEXT = REFE$
   '
   PRIMOL& = 0: HAYMO% = 0: ASHAS = 999999
   FIN& = MAXDIA&
   For IL& = 1 To MAXDIA&
      U& = IL&
      Call TRACE(20, U&, FIN&)
      Call LEESUCAJA(IL&)
      If CVI(D11$) > 0 Then
         If CVI(D15$) <= FEAS% Then
            If CVS(D19$) < 0.05 Then
               If PRIMOL& = 0 Then PRIMOL& = IL&
               CUE% = CVI(D11$)
               IM# = CVD(D16$)
               For J% = 1 To CONTA%
                  If CTA%(J%) = CUE% Then
                     IPAS#(J%) = IPAS#(J%) + (Int(100 * IM# + 0.5)) / 100
                     GoTo 75
                  End If
               Next J%
               CONTA% = CONTA% + 1
               CTA%(CONTA%) = CUE%
               IPAS#(CONTA%) = (Int(100 * IM# + 0.5)) / 100
            End If
         End If
      End If
75 Next IL&
   '
   K% = 0
   For I% = 1 To CONTA%
      If Abs(IPAS#(I%)) >= 0.005 Then
         K% = K% + 1
         CTA%(K%) = CTA%(I%)
         IPAS#(K%) = IPAS#(I%)
      End If
   Next I%
   CONTA% = K%
   '
   For I% = 1 To CONTA%
      Z$ = Space$(128)
      Call REPLA(Z$, CODCUE$(CTA%(I%)), 1, 12)
      DEPASE$ = DENOCUE$(CTA%(I%))
      Call REPLA(Z$, DEPASE$, 14, 27)
      If IPAS#(I%) > 0 Then
           Call REPLA(Z$, CSTRING$(MKD$(IPAS#(I%)), 3, 12, 2, 2), 41, 12)
         Else
           Call REPLA(Z$, CSTRING$(MKD$(Abs(IPAS#(I%))), 3, 12, 2, 2), 53, 12)
      End If
      List2.AddItem "A" + Z$
   Next I%
   '
   For U& = 1 To List2.ListCount
      Z$ = Mid$(List2.List(U& - 1), 2)
      List1.AddItem Z$
   Next U&
   '
   Call SUMASIEN
   Screen.MousePointer = 1
   '
End Sub

Sub MARCASCAJA()
   '
   Screen.MousePointer = 11
   NROAS = Val(Text2.TEXT)
   FEAS% = FECHANUM(Text1.TEXT)
   If PRIMOL& < 1 Then PRIMOL& = 1
   FIN& = MAXDIA&
   For IL& = PRIMOL& To MAXDIA&
      U& = IL&
      Call TRACE(20, U&, FIN&)
      Call LEESUCAJA(IL&)
      If CVI(D11$) > 0 Then
         If CVI(D15$) <= FEAS% Then
            If CVS(D19$) < 0.05 Then
               DXX$ = D10$
               Call REPLA(DXX$, MKS$(NROAS), 55, 4)
               LSet D10$ = DXX$
               Put #NX6%, IL&, D10$
            End If
         End If
      End If
   Next IL&
   '
   Close #NX6%: NX6% = 0
   Screen.MousePointer = 1
   '
End Sub

Sub PRINASIDIAR()
   '
   FORIPRE$ = CONTROL$("", "ASICONTA")
   '
   If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      FECHDOC$ = Text1.TEXT
      NUMEDOC$ = TRIM$(Text2.TEXT)
      TIPODOC$ = "Comprobante de Diario"
      '
      CODPARAM$(1) = "REF-ASIE"
      CODPARAM$(2) = "TOT-DEBE"
      CODPARAM$(3) = "TOT-HABE"
      CAPARDOC% = 3
      '
      DOCPARAM$(1) = TRIM$(Text4.TEXT)
      DOCPARAM$(2) = Label21.Caption
      DOCPARAM$(3) = Label22.Caption
      '
      CODTABU1$(1) = "NRO-PASE"
      CODTABU1$(2) = "COD-CUEN"
      CODTABU1$(3) = "DEN-CUEN"
      CODTABU1$(4) = "DET-PASE"
      CODTABU1$(5) = "IMP-DEBE"
      CODTABU1$(6) = "IMP-HABE"
      CODTABU1$(7) = "NRO-PASE"
      CACOLTAB1% = 7
      '
      CAITAB1% = 0
      For U& = 1 To List1.ListCount
         Z$ = List1.List(U& - 1)
         CAITAB1% = CAITAB1% + 1
         TETABU1$(1, CAITAB1%) = "." ' Str$(CAITAB1%)
         TETABU1$(2, CAITAB1%) = Left$(Z$, 12)
         CUE% = CUEINT%(Left$(Z$, 12))
         TETABU1$(3, CAITAB1%) = DENOCUE$(CUE%)
         TETABU1$(4, CAITAB1%) = Mid$(Z$, 14, 27)
         TETABU1$(5, CAITAB1%) = Mid$(Z$, 41, 12)
         TETABU1$(6, CAITAB1%) = Mid$(Z$, 53, 12)
         TETABU1$(7, CAITAB1%) = Str$(CAITAB1%)
      Next U&
      '
      Call PRINTDOC("ASICONTA")   ' asiento contable
      Screen.MousePointer = 1
      '
   End If
   '
End Sub

