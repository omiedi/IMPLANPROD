VERSION 5.00
Begin VB.Form ASICAJA 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asientos Contables"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Height          =   675
      Left            =   8190
      Picture         =   "ASICAJ00.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Activa Calculadora del Sistema"
      Top             =   7310
      Width           =   1000
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   8610
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   41
      Top             =   4620
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   12000
      End
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
      BackColor       =   &H00EFEFEF&
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
      Top             =   7245
      Width           =   7785
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
         Width           =   6420
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
      TabIndex        =   24
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
      Picture         =   "ASICAJ00.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Imprime Comprobante del Asiento"
      Top             =   3570
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EFEFEF&
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
      TabIndex        =   20
      Top             =   4935
      Width           =   9255
      Begin VB.CommandButton Command6 
         Height          =   645
         Left            =   7980
         Picture         =   "ASICAJ00.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Agrega Pase al Asiento"
         Top             =   1470
         Width           =   1000
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00EFEFEF&
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
            TabIndex        =   4
            TabStop         =   0   'False
            Text            =   " "
            Top             =   315
            Width           =   4005
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00EFEFEF&
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
         BackColor       =   &H00EFEFEF&
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
            MaxLength       =   10
            TabIndex        =   6
            Text            =   " "
            Top             =   420
            Width           =   1380
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EFEFEF&
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
            MaxLength       =   10
            TabIndex        =   5
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
         Picture         =   "ASICAJ00.frx":0C7E
         TabIndex        =   11
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
         TabIndex        =   25
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
         Picture         =   "ASICAJ00.frx":0F88
         TabIndex        =   10
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
         Picture         =   "ASICAJ00.frx":1292
         TabIndex        =   9
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
         Picture         =   "ASICAJ00.frx":159C
         TabIndex        =   8
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
         TabIndex        =   21
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
      TabIndex        =   18
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
         TabIndex        =   19
         Top             =   315
         Width           =   8100
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   360
         Left            =   0
         TabIndex        =   22
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
            TabIndex        =   23
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
      Picture         =   "ASICAJ00.frx":18A6
      Style           =   1  'Graphical
      TabIndex        =   14
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
      Picture         =   "ASICAJ00.frx":1BB0
      Style           =   1  'Graphical
      TabIndex        =   16
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
      Picture         =   "ASICAJ00.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   13
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
      Picture         =   "ASICAJ00.frx":2004
      Style           =   1  'Graphical
      TabIndex        =   12
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
      Left            =   945
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
      Caption         =   "Asiento de Caja - Bancos"
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
      TabIndex        =   17
      Top             =   315
      Width           =   1275
   End
End
Attribute VB_Name = "ASICAJA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FEAS%
Dim SUDEBE#, SUHABE#
Dim CTA%(1024), IMPA#(1024), DETA$(1024)
Dim ACTEDIT%
'
Private Sub Command1_Click()
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   '
   NROAS = (Int(10 * Val(Text2.TEXT) + 0.5)) / 10
   NUASORIGI = NROAS
   FEAS% = FECHANUM(Text1.TEXT)
   FEASORIGI% = FEAS%
   REFE$ = Text4.TEXT
   CONTA% = List1.ListCount
   SUMATOT# = 0
   '
   On Error Resume Next
   TIASORI% = Asc(Left$(D8$, 1))
   On Error GoTo 0
   If TIASI% = 0 Then
      If TIASORI% = 4 And CONTA% < 1 Then GoTo 10
      If POSEDI% < 0 Then
         Call MENSERR(24, "Imposible Grabar este Asiento")
         Exit Sub
      End If
   End If
   '
10 For U& = 1 To List1.ListCount
      Z$ = List1.List(U& - 1)
      CTA%(U&) = CUEINT%(Left$(Z$, 12))
      If CTA%(U&) < 1 Or CTA%(U&) > NUCUEN% Then
         Call MENSERR(24, "Imposible Grabar este Asiento.\  \Número de Cuenta Incorrecto.")
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
   Call GRASICAJA(NROAS, FEAS%, REFE$, CONTA%, CTA%(), IMPA#(), DETA$())
   If TIASI% = 2 Then Call CARBACHEQ ' carga BADACHE
   If TIASI% = 4 Then Call BLABACHEQ ' marca salida en BADACHE
   If TIASI% = 0 And TIASORI% = 4 And CONTA% < 1 Then Call RECARCHE(FEASORIGI%, NUASORIGI)
   '
   If CONTROL$("ASICONTA", "IMPRECOM") <> "NO" Then
      Call PRINASICAJA
   End If
   '
   Call CLEARASICAJA
   If TIASI% > 0 Then
      Unload ASICAJA
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
   If TIASI% > 0 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
   Call CLEARASICAJA
End Sub

Private Sub Command4_Click()
   Call BORRAPASE
End Sub

Private Sub Command5_Click()
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   'PrintForm
   Call PRINASICAJA
   '
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
         Label3.Caption = "(Resumen de Pagos)"
         Call ASIPAGOS
         If List1.ListCount < 1 Then
            Call MENSERR(24, "No hay Movimientos\en la Fecha Elegida")
            Text1.TEXT = "********"
         End If
      ElseIf TIASI% = 2 Then
         Label3.Caption = "(Resumen de Cobranza)"
         Call ASICOBRA
         If List1.ListCount < 1 Then
            Call MENSERR(24, "No hay Movimientos\en la Fecha Elegida")
            Text1.TEXT = "********"
         End If
      ElseIf TIASI% = 3 Then
         Label3.Caption = "(Resumen de Fondo Fijo)"
         Call ASIFOFIJ
         If List1.ListCount < 1 Then
            Call MENSERR(24, "No hay Movimientos\en la Fecha Elegida")
            Text1.TEXT = "********"
         End If
      ElseIf TIASI% = 4 Then
         Label3.Caption = "(Depósito de Cheques)"
         Call ASIDECHE
         If List1.ListCount < 1 Then
            Text1.TEXT = "********"
         End If
    End If
    '
    ACTEDIT% = 0
    '
End Sub


Private Sub Form_Unload(Cancel As Integer)
    TIASI% = 0
    Call CIERRACONTA
    'Call LIBARCH("ASICONT")
End Sub

Private Sub Label1_DblClick()
   Call CORRNUASICAJA
End Sub

Private Sub Label11_DblClick()
   OPASICONT.Show 1
End Sub

Private Sub List1_DblClick()
   Call EDITPASE
End Sub

Private Sub Text1_GotFocus()
   If TIASI% > 0 Then
      If Text1.TEXT = "********" Then
         Unload ASICAJA
         Exit Sub
      End If
   End If
   '
   If TIASI% > 0 Then Command1.SetFocus: Exit Sub
   If ACTEDIT% = 1 Then Text11.SetFocus: Exit Sub
   VV% = VALIFECHAS%
   Text1.SelStart = 0
   Text1.SelLength = 8
End Sub
'
Private Sub Text1_DblClick()
   If TIASI% > 0 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
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
   If TIASI% > 0 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
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
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
   If VALIFECHAS% > 0 Then
        NPA = NUPROAS
        'Text2.SetFocus
      Else
        Call MENSERR(24, "Fecha no Válida")
        Text1.SetFocus
   End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 27 Then
      If List1.ListCount > 0 Then
         Call SUMASIEN
         If Abs(SUDEBE# - SUHABE#) < 0.005 Then
            If SUDEBE# >= 0.005 Then
               Command1.SetFocus
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
               Command1.SetFocus
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
   Text14.TEXT = CSTRING$(MKD$(Val(Text14.TEXT)), 3, 10, 2, 2)
End Sub

Private Sub Text15_LostFocus()
   Text15.TEXT = CSTRING$(MKD$(Val(Text15.TEXT)), 3, 10, 2, 2)
End Sub

Private Sub Text2_Change()
   List1.Clear
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
   If TIASI% > 0 Then Command1.SetFocus: Exit Sub
   If ACTEDIT% = 1 Then Text11.SetFocus: Exit Sub
   VV% = VALIFECHAS%
   Text2.SelStart = 0
   Text2.SelLength = 8
End Sub
'
Private Sub Text14_GotFocus()
   If TIASI% > 0 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
   If VALICUENTA% <> 1 Then Exit Sub
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   If Abs(Val(Text14.TEXT)) < 0.005 Then
      Text14.TEXT = "      0.00"
   End If
   Text14.SelStart = 0
   Text14.SelLength = Len(Text14.TEXT)
End Sub

Private Sub Text15_GotFocus()
   If TIASI% > 0 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
   If VALICUENTA% <> 1 Then Exit Sub
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   If Abs(Val(Text15.TEXT)) < 0.005 Then
      Text15.TEXT = "      0.00"
   End If
   Text15.SelStart = 0
   Text15.SelLength = Len(Text14.TEXT)
End Sub
Private Sub Text14_Change()
   If Abs(Val(Text14.TEXT)) >= 0.005 Then
      Text15.TEXT = "      0.00"
   End If
End Sub

Private Sub Text15_Change()
   If Abs(Val(Text15.TEXT)) >= 0.005 Then
      Text14.TEXT = "      0.00"
   End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 27 Then
      If List1.ListCount > 0 Then
         Call SUMASIEN
         If Abs(SUDEBE# - SUHABE#) < 0.005 Then
            If SUDEBE# >= 0.005 Then
               Command1.SetFocus
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
               Command1.SetFocus
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
    Call HELPONLINE("ASICAJA")
End Sub
'

Private Sub Command9_Click()
    '
    FEX = FEPROAS
    Text1.TEXT = FECHATEX$(FEX)
    Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
    List1.Clear
    Text11.TEXT = ""
    Label4.Caption = ""
    Text12.TEXT = ""
    Text4.TEXT = Space$(40)
    Text14.TEXT = ""
    Text15.TEXT = ""
    Call SUMASIEN
    Call CIERRARCH("*.*")
    Unload ASICAJA
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
      'Call CARGASICAJ
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
   Call CARGASICAJ
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
   If TIASI% > 0 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   Text4.SelStart = 0
   Text4.SelLength = 40
End Sub

Private Sub Text11_DBLCLICK()
   If TIASI% > 0 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
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
   If TIASI% > 0 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
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
   If TIASI% > 0 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
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
   If TIASI% > 0 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   '
   If VALICUENTA% <> 1 Then Exit Sub
   If Text12 = "" Then
     Text12 = Label4
   End If
   '
   Text12.SelStart = 0
   Text12.SelLength = 40
End Sub
'

Private Sub VScroll1_Change()
   '
   If TIASI% > 0 Then Exit Sub
   Static VALANT%
   If VALANT% = 0 Then VALANT% = 16384
   NROACT = Val(Text2.TEXT)
   '
   If NROACT > ULTAS + 0.01 Then
      RGENC% = MAXENC% + 1
      GoTo 85
   End If
   '
   LI% = 0: LS% = MAXENC% + 1
82 E% = Int((LS% - LI%) / 2)
   If E% < 1 Then GoTo 83
   L% = LI% + E%
   If L% < 1 Then RGENC% = 1: GoTo 85
   If L% > MAXENC% Then RGENC% = MAXENC% + 1: GoTo 85
   Call LEENCAJA(L%)
   If CVS(D1$) < NROACT - 0.01 Then LI% = L%: GoTo 82
   If CVS(D1$) > NROACT + 0.01 Then LS% = L%: GoTo 82
83 RGENC% = L%
   '
85 If VScroll1.Value < VALANT% Then
        RGENC% = RGENC% + 1
      ElseIf VScroll1.Value > VALANT% Then
        RGENC% = RGENC% - 1
   End If
   VALANT% = VScroll1.Value
   If RGENC% > 0 Then
      If RGENC% <= MAXENC% Then
           Call LEENCAJA(RGENC%)
           NROAS = CVS(D1$)
           Text2.TEXT = CSTRING$(MKS$(NROAS), 3, 8, 1, 2)
           Call CARGASICAJ
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
      FEX = FEPROAS%
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
   Text2.TEXT = CSTRING$(MKS$(NROAS), 3, 8, 1, 2)
   If NROAS = NROAA Then
      VALINUMAS% = VNA%
      Exit Function
   End If
   '
   NROAA = NROAS
   If NROAS < 1 Then GoTo 59
   If NROAS > 9999.9 Then GoTo 59
   VALINUMAS% = 1: VNA% = 1
   Exit Function
   '
59 FEX = FEPROAS%
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
   LI% = 0: LS% = MAXENC% + 1
   L% = MAXENC%: ERRO% = 0
   If NRO > ULTAS + 0.01 Then
      FEASA = FECHANUM(Text1.TEXT)
      If FEASA < FULTAS% Then
         FEASA = FULTAS%
      End If
      GoTo 99
   End If
   '
82 E% = Int((LS% - LI%) / 2)
   If E% < 1 Then GoTo 98
   L% = LI% + E%
   If L% < 1 Or L% > MAXENC% Then GoTo 98
   Call LEENCAJA(L%)
   If CVS(D1$) < NRO - 0.01 Then LI% = L%: GoTo 82
   If CVS(D1$) > NRO + 0.01 Then LS% = L%: GoTo 82
   REGENCAS% = L%
   Exit Function
   '
98 FEASA = FECHANUM(Text1.TEXT)
   If FEASA < COMEJ% Then FEASA = COMEJ%
   If LI% > 0 Then
      Call LEENCAJA(LI%)
      If FEASA < CVI(D3$) Then FEASA = CVI(D3$)
   End If
   If LS% < LI% + 1 Then LS% = LI% + 1
   If LS% <= MAXENC% Then
      Call LEENCAJA(LS%)
      If FEASA > CVI(D3$) Then FEASA = CVI(D3$)
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
   LI% = 0: LS% = MAXENC% + 1
   L% = MAXENC%: ERRO% = 0
   If FEAS% >= FULTAS% Then
      NUPROAS = Int(1 + ULTAS)
      Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
      Exit Function
   End If
   '
81 E% = Int((LS% - LI%) / 2)
   If E% < 1 Then GoTo 82
   L% = LI% + E%
   If L% < 1 Or L% > MAXENC% Then GoTo 99
   Call LEENCAJA(L%)
   If CVI(D3$) < FEAS% Then LI% = L%: GoTo 81
   If CVI(D3$) > FEAS% Then LS% = L%: GoTo 81
   
82 For L1% = L% To 1 Step -1
      Call LEENCAJA(L1%)
      If CVI(D3$) < FEAS% Then GoTo 83
   Next L1%
   L1% = 0
   '
83 L2% = L1%
   NASPROX = 0
   For L% = L1% + 1 To MAXENC%
      Call LEENCAJA(L%)
      If CVI(D3$) > FEAS% Then GoTo 84
      L2% = L%
   Next L%
   '
84 If L2% < 1 Then
      Call LEENCAJA(1)
      NUPROAS = Int(CVS(D1$)) - 1
      If NUPROAS < 1 Then GoTo 99
      If NUPROAS > 10 Then NUPROAS = 10
      GoTo 85
   End If
   '
   Call LEENCAJA(L2%)
   NUPROAS = (1 + Int(10 * CVS(D1$) + 0.5)) / 10
   If L2% < MAXENC% Then
      Call LEENCAJA(L2% + 1)
      If CVS(D1$) < NUPROAS - 0.01 Then
         GoTo 99
      End If
   End If
   '
85 Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
   RGENC% = REGENCAS%
   If RGENC% > 0 Then
      Call LEENCAJA(RGENC%)
      If CVI(D2$) < FEAS% Then GoTo 99
      NUPROAS = (Int(10 * NUPROAS + 0.5) - 1) / 10
      GoTo 85
   End If
   If NUPROAS < 0.1 Then GoTo 99
   FEX = FEAS%: Text1.TEXT = FECHATEX$(FEX)
   Exit Function
   '
99 On Error Resume Next
   FORMALTER.Hide
   On Error GoTo 0
   Call MENSERR(24, "Imposible Registrar Asiento\en la Fecha Elegida.")
   Call FINAL("")
   '
End Function

Sub CARGASICAJ()
   RGE% = REGENCAS%
   If RGE% > 0 Then
      NROAS = Val(Text2.TEXT)
      Call LEENCAJA(RGE%)
      FF% = CVI(D3$)
      CONTA% = CVI(D6$)
      FEX = FF%
      Text1.TEXT = FECHATEX$(FEX)
      Text4.TEXT = D7$
      '
81    ERRO% = 0
      LI% = 0: LS% = MAXDIA% + 1
82    E% = Int((LS% - LI%) / 2)
      If E% < 1 Then
         NABORX$ = Text2.TEXT
         MENSA$ = "Se ha Detectado un Error de Consistencia\en el Asiento Numero '" + TRIM$(NABORX$) + "'.\  \"
         MENSA$ = MENSA$ + "FLEXOFT Puede Recuperar la Consistencia\de la Base de Datos, Borrando la Cabecera\de este Asiento.\\"
         If COMALTER%(MENSA$ + "Cancelar\Borrar Cabecera") = 2 Then
            Call BLOQARCH("ACONT" + EXTE$)
            'Call BLOQARCH("ASICONT")
            RGX% = RGE% - 1
            FINX% = MAXENC%
            NABOR = Val(Text2.TEXT)
            Call ABRENCAJA
            For UI% = RGE% To FINX% + 1
               Get #NX5%, UI%, D0$
               NOAS = CVS(Left$(D0$, 4))
               If Abs(NOAS - NABOR) > 0.01 Then
                  RGX% = RGX% + 1
                  Put #NX5%, RGX%, D0$
               End If
            Next UI%
            '
            For UI% = RGX% + 1 To FINX%
               LSet D0$ = String$(64, 0)
               Put #NX5%, UI%, D0$
            Next UI%
            '
            Close #NX5%: NX5% = 0
            Call COMUNI("Recuperación Completa.\Re-Inicie Ahora el Sistema Contable.")
         End If
         Call FINAL("")
      End If
      L% = LI% + E%
      If L% < 1 Or L% > MAXDIA% Then
         ERRO% = 1: GoTo 89
      End If
      Call LEESUCAJA(L%)
      If CVS(D12$) < NROAS - 0.01 Then LI% = L%: GoTo 82
      If CVS(D12$) > NROAS + 0.01 Then LS% = L%: GoTo 82
      '
      For K1% = L% To 1 Step -1
         Call LEESUCAJA(K1%)
         If Abs(CVS(D12$) - NROAS) > 0.01 Then GoTo 83
         REGINI% = K1%
      Next K1%
      '
83    List1.Clear
      For K1% = 1 To CONTA%
         Call LEESUCAJA(REGINI% + K1% - 1)
         If Abs(CVS(D12$) - NROAS) > 0.01 Or CVI(D15$) <> FF% Then
            ERRO% = 1: GoTo 89
         End If
         IPASE# = CVD(D16$)
         If CVS(D19$) > 0.01 Then POSEDI% = (-1)
         If Asc(Left$(D8$, 1)) > 0 Then POSEDI% = (-1)
         If CVI(D2$) > 0 Then POSEDI% = (-1)
         '
         Z$ = Space$(128)
         Call REPLA(Z$, CODCUE$(CVI(D11$)), 1, 12)
         Call REPLA(Z$, D18$, 14, 27)
         If IPASE# > 0 Then
              Call REPLA(Z$, CSTRING$(MKD$(IPASE#), 3, 12, 2, 2), 41, 12)
            Else
              Call REPLA(Z$, CSTRING$(MKD$(Abs(IPASE#)), 3, 12, 2, 2), 53, 12)
         End If
         List1.AddItem Z$
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
         Call LEESUCAJA(REGINI% - 1)
         If CVS(D12$) >= NROAS + 0.01 Then ERRO% = 1: GoTo 89
      End If
      '
      Call LEESUCAJA(REGINI% + CONTA%)
      If CVI(D11$) <> 0 Then
         If CVS(D12$) <= NROAS - 0.01 Then ERRO% = 1
         If Abs(CVS(D12$) - NROAS) < 0.01 Then ERRO% = 2
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
   Call SUMASIEN
   '
End Sub
'
Sub CLEARASICAJA()
   '
   FEX = FEPROAS%
   FEAS% = FEX
   Text1.TEXT = FECHATEX$(FEX)
   Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
   List1.Clear
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
   If TIASI% > 0 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
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
   Call REPLA(Z$, Text12.TEXT, 14, 27)
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
   If TIASI% > 0 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
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
         Text14.TEXT = CSTRING$(MKD$(IDEBE#), 3, 10, 2, 2)
         Text15.TEXT = CSTRING$(MKD$(IHABE#), 3, 10, 2, 2)
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
   If TIASI% > 0 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then
      TIASORI% = Asc(Left$(D8$, 1))
      If TIASORI% = 4 Then    ' PERMITE BORRAR DEPOSITO BANCARIO
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
      Command1.SetFocus: Exit Sub
   End If
   '
   If TIASI% > 0 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
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
        Label10.Caption = CSTRING$(MKD$(SUDEBE# - SUHABE#), 3, 10, 2, 2)
      ElseIf SUHABE# >= SUDEBE# + 0.005 Then
        Label8.Caption = CSTRING$(MKD$(SUHABE# - SUDEBE#), 3, 10, 2, 2)
   End If
   '
   On Error Resume Next
   Text11.SetFocus
   On Error GoTo 0
   '
End Sub

Sub ASIPAGOS()
   '
   Call SELECHO("SELFECHA")
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
   Text1.TEXT = FF$
   REFE$ = AJUSTI$("RESUMEN DIARIO CAJA DEL " + FF$, 40)
   Text4.TEXT = REFE$
   '
   ' BUSQUEDA POR FECHA Y REFERENCIA
   LI% = 0: LS% = MAXENC% + 1
   L% = MAXENC%: ERRO% = 0
   If FEAS% > FULTAS% Then
      Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
      GoTo 90
   End If
   '
82 E% = Int((LS% - LI%) / 2)
   If E% < 1 Then GoTo 90
   L% = LI% + E%
   If L% < 1 Or L% > MAXENC% Then GoTo 90
   Call LEENCAJA(L%)
   If CVI(D3$) < FEAS% Then LI% = L%: GoTo 82
   If CVI(D3$) > FEAS% Then LS% = L%: GoTo 82
   '
   ENCUAS% = 0: NROAS = 0
   For L1% = L% To 1 Step -1
      Call LEENCAJA(L1%)
      If CVI(D3$) <> FEAS% Then GoTo 83
   Next L1%
   L1% = 0
   '
83 L1% = L1% + 1
   NASPROX = 0
   For L% = L1% To MAXENC%
      Call LEENCAJA(L%)
      If CVI(D3$) <> FEAS% Then GoTo 84
      If TRIM$(D7$) = TRIM$(REFE$) Then              ' ENCONTRO EL ASIENTO
         ENCUAS% = 1
         NROAS = CVS(D1$)
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
      CMDX% = COMALTER%("Asiento de Pagos del " + FF$ + "\ya Registrado.\\Mantener\Refrescar")
      If CMDX% <> 2 Then
           Call CARGASICAJ
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
   Text4.TEXT = REFE$
   Call GENASIPAGO
   '
End Sub

Sub GENASIPAGO()
   '
   Static CAMP%, CUEMP%(2048), MEPA%(2048)
   '
   If CAMP% <= 0 Then
      For I& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", I&)
         CAMP% = I&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUE$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUE$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUE$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next I&
   End If
   '
   Screen.MousePointer = 11
   '
   FFAA$ = ASICAJA.Text1.TEXT
   FF% = FECHANUM(FFAA$)
   '
   Dim NOPFF(2048)
   CANOPFF% = 0
   FIN& = ULTREG&("CUECORR")
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      CC0$ = REGLEIDO$("CUECORR", I&)
      If Mid$(CC0$, 7, 2) = "FF" Then
         If CVI(Left$(CC0$, 2)) = FF% Then
            NOPA = Val(Mid$(CC0$, 20, 6))
            If NOPA <> NOPFF(CANOPFF%) Then
               CANOPFF% = CANOPFF% + 1
               NOPFF(CANOPFF%) = NOPA
            End If
         End If
      End If
   Next I&
   '
   NOPANT = 0: NPANT% = 0
   ATX% = 0
   FIN& = ULTREG&("CHEQUEMI")
   For IAX& = 1 To FIN&
      Call TRACE(20, IAX&, FIN&)
      CH0$ = REGLEIDO$("CHEQUEMI", IAX&)
      CH1$ = Mid$(CH0$, 1, 2)
      CH5$ = Mid$(CH0$, 33, 8)
      CH6$ = Mid$(CH0$, 41, 4)
      CH7$ = Mid$(CH0$, 45, 2)
      CH8$ = Mid$(CH0$, 47, 2)
      '
      If CVI(CH8$) = FF% Then                ' CORRESPONDE FECHA
         IM# = CVD(CH5$)
         If Abs(IM#) >= 0.005 Then
            NOPA = CVS(CH6$)
            '
            For KKK% = 1 To CANOPFF%
               If NOPFF(KKK%) = NOPA Then GoTo 76 ' DESCARTA FONDO FIJO
            Next KKK%
            '
            NP% = CVI(CH1$)         ' NRO DE PROVEEDOR
            TIPOVA% = CVI(CH7$)
            QUE% = CVI(CH7$)
            '
            If NOPA <> NOPANT Or NP% <> NPANT% Then
               ITOT# = 0
               For J& = IAX& To FIN&
                  CH0$ = REGLEIDO$("CHEQUEMI", J&)
                  CH1$ = Mid$(CH0$, 1, 2)
                  CH5$ = Mid$(CH0$, 33, 8)
                  CH6$ = Mid$(CH0$, 41, 4)
                  If CVI(CH1$) = NP% Then
                     If CVS(CH6$) = NOPA Then
                        ITOT# = ITOT# + CVD(CH5$)
                     End If
                  End If
               Next J&
               '
               If TICUEPRO((-1) * NP%) = 1 Then ' CUENTA DE GASTOS
                  For UDP& = 1 To ULTREG&("DETPAGAS")
                     DTP$ = REGLEIDO$("DETPAGAS", UDP&)
                     If CVS(Mid$(DTP$, 3, 4)) = NOPA Then
                        CUEDP% = CVI(Mid$(DTP$, 17, 2))
                        If TRIM$(CODCUE$(CUEDP%)) <> "" Then
                           IDP# = CVD(Mid$(DTP$, 9, 8))
                           Z$ = Space$(128)
                           Call REPLA(Z$, CODCUE$(CUEDP%), 1, 12)
                           DETAX$ = "OP " + TRIM$(Str$(NOPA)) + " " + RASOCLI$((-1) * NP%)
                           Call REPLA(Z$, DETAX$, 14, 27)
                           Call REPLA(Z$, CSTRING$(MKD$(IDP#), 3, 12, 2, 2), 41, 12)
                           List1.AddItem Z$
                           ITOT# = ITOT# - IDP#
                        End If
                     End If
                  Next UDP&
               End If
               '
               If Abs(ITOT#) >= 0.005 Then
                  Z$ = Space$(128)
                  Call REPLA(Z$, CUEMADRE$((-1) * NP%), 1, 12)
                  DETAX$ = "OP " + TRIM$(Str$(NOPA)) + " " + RASOCLI$((-1) * NP%)
                  Call REPLA(Z$, DETAX$, 14, 27)
                  Call REPLA(Z$, CSTRING$(MKD$(ITOT#), 3, 12, 2, 2), 41, 12)
                  List1.AddItem Z$
               End If
               '
               NOPANT = NOPA
               NPANT% = NP%
            End If
            '
            CH0$ = REGLEIDO$("CHEQUEMI", IAX&)
            CH1$ = Mid$(CH0$, 1, 2)
            CH2$ = Mid$(CH0$, 3, 12)
            CH3$ = Mid$(CH0$, 15, 16)
            CH5$ = Mid$(CH0$, 33, 8)
            CH6$ = Mid$(CH0$, 41, 4)
            CH7$ = Mid$(CH0$, 45, 2)
            CH8$ = Mid$(CH0$, 47, 2)
            '
            CHE$ = ""
            For JI% = 1 To Len(CH2$)
               If Asc(Mid$(CH2$, JI%)) > 32 Then CHE$ = CHE$ + Mid$(CH2$, JI%, 1)
            Next JI%
            '
            NCHE$ = ""
            For JI% = 1 To Len(CH3$)
               If Asc(Mid$(CH3$, JI%)) > 32 Then NCHE$ = NCHE$ + Mid$(CH3$, JI%, 1)
            Next JI%
            '
            For UX% = 1 To CAMP%
               If MEPA%(UX%) = QUE% Then GoTo 18
            Next UX%
            Call MENSERR(24, "Medio de Pago" + Str$(QUE%) + " no Definido")
            Call FINAL("")
            '
18          CUEH% = CUEMP%(UX%)
            IPASE# = CVD(CH5$)
            DTX$ = TRIM$(CHE$)
            If DTX$ = "" Then DTX$ = TRIM$(NCHE$)
            If DTX$ <> "" Then DTX$ = DTX$ + " - "
            DTX$ = DTX$ + "OP" + Str$(NOPA) + " " + RASOCLI$((-1) * NP%)
            Z$ = Space$(128)
            Call REPLA(Z$, CODCUE$(CUEH%), 1, 12)
            Call REPLA(Z$, DTX$, 14, 27)
            Call REPLA(Z$, CSTRING$(MKD$(IPASE#), 3, 12, 2, 2), 53, 12)
            List1.AddItem Z$
            '
         End If
      End If
76 Next IAX&
   '
   Call SUMASIEN
   Screen.MousePointer = 1
   '
End Sub

Sub ASICOBRA()
   '
   Call SELECHO("SELFECHA")
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
   Text1.TEXT = FF$
   REFE$ = AJUSTI$("RESUMEN DIARIO COBRANZA DEL " + FF$, 40)
   Text4.TEXT = REFE$
   '
   ' BUSQUEDA POR FECHA Y REFERENCIA
   LI% = 0: LS% = MAXENC% + 1
   L% = MAXENC%: ERRO% = 0
   If FEAS% > FULTAS% Then
      Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
      GoTo 90
   End If
   '
82 E% = Int((LS% - LI%) / 2)
   If E% < 1 Then GoTo 90
   L% = LI% + E%
   If L% < 1 Or L% > MAXENC% Then GoTo 90
   Call LEENCAJA(L%)
   If CVI(D3$) < FEAS% Then LI% = L%: GoTo 82
   If CVI(D3$) > FEAS% Then LS% = L%: GoTo 82
   '
   ENCUAS% = 0: NROAS = 0
   For L1% = L% To 1 Step -1
      Call LEENCAJA(L1%)
      If CVI(D3$) <> FEAS% Then GoTo 83
   Next L1%
   L1% = 0
   '
83 L1% = L1% + 1
   NASPROX = 0
   For L% = L1% To MAXENC%
      Call LEENCAJA(L%)
      If CVI(D3$) <> FEAS% Then GoTo 84
      If TRIM$(D7$) = TRIM$(REFE$) Then              ' ENCONTRO EL ASIENTO
         ENCUAS% = 1
         NROAS = CVS(D1$)
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
      CMDX% = COMALTER%("Asiento de Cobranza del " + FF$ + "\ya Registrado.\\Mantener\Refrescar")
      If CMDX% <> 2 Then
           Call CARGASICAJ
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
   Text4.TEXT = REFE$
   Call GENASICOBRA
   '
End Sub

Sub GENASICOBRA()
   '
   Static CAMP%, CUEMP%(2048), MEPA%(2048), DEVAL$(2048)
   '
   If CAMP% <= 0 Then
      For I& = 1 To ULTREG&("MECOBRA")
         F0$ = REGLEIDO$("MECOBRA", I&)
         CAMP% = I&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         DEVAL$(CAMP%) = Mid$(F0$, 3, 12)
         CUE$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUE$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUE$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Cobranza '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next I&
   End If
   '
   Screen.MousePointer = 11
   FFAA$ = ASICAJA.Text1.TEXT
   FF% = FECHANUM(FFAA$)
   '
   NOPANT = 0: NPANT% = 0
   ATX% = 0
   FIN& = ULTREG&("CHECAR")
   FBC& = ULTREG&("BADACHE")
   '
   For IAX& = 1 To FIN&
      Call TRACE(20, IAX&, FIN&)
      CH0$ = REGLEIDO$("CHECAR", IAX&)
      CH1$ = Mid$(CH0$, 3, 2)
      CH5$ = Mid$(CH0$, 33, 8)
      CH6$ = Mid$(CH0$, 47, 4)
      CH9$ = Mid$(CH0$, 51, 2)
      CH7$ = Mid$(CH0$, 1, 2)
      CH8$ = Mid$(CH0$, 45, 2)
      '
      If CVI(CH8$) = FF% Then                ' CORRESPONDE FECHA
         IM# = CVD(CH5$)
         If Abs(IM#) >= 0.005 Then
            NOPA = CVS(CH6$)
            '
            NP% = CVI(CH1$)         ' NRO DE CLIENTE
            TIPOVA% = CVI(CH7$)
            QUE% = CVI(CH7$)
            '
            If NOPA <> NOPANT Or NP% <> NPANT% Then
               ITOT# = 0
               For J& = IAX& To FIN&
                  CH0$ = REGLEIDO$("CHECAR", J&)
                  CH1$ = Mid$(CH0$, 3, 2)
                  CH5$ = Mid$(CH0$, 33, 8)
                  CH6$ = Mid$(CH0$, 47, 4)
                  CH8$ = Mid$(CH0$, 45, 2)
                  If CVI(CH1$) = NP% Then
                    If CVS(CH6$) = NOPA Then
                      If CVI(CH8$) = FF% Then
                        ITOT# = ITOT# + CVD(CH5$)
                      End If
                    End If
                  End If
               Next J&
               '
               Z$ = Space$(128)
               Call REPLA(Z$, CUEMADRE$(NP%), 1, 12)
               DETAX$ = "RB " + TRIM$(Str$(NOPA)) + " " + RASOCLI$(NP%)
               Call REPLA(Z$, DETAX$, 14, 27)
               Call REPLA(Z$, CSTRING$(MKD$(ITOT#), 3, 12, 2, 2), 53, 12)
               List1.AddItem Z$
               '
               NOPANT = NOPA
               NPANT% = NP%
            End If
            '
            CH0$ = REGLEIDO$("CHECAR", IAX&)
            CH1$ = Mid$(CH0$, 3, 2)
            CH5$ = Mid$(CH0$, 33, 8)
            CH6$ = Mid$(CH0$, 47, 4)
            CH7$ = Mid$(CH0$, 1, 2)
            CH8$ = Mid$(CH0$, 45, 2)
            CH2$ = Mid$(CH0$, 5, 12)
            CH3$ = Mid$(CH0$, 17, 16)
            RBC& = CVS(Mid$(CH0$, 51, 4))
            '
            CHE$ = ""
            For JI% = 1 To Len(CH2$)
               If Asc(Mid$(CH2$, JI%)) > 32 Then CHE$ = CHE$ + Mid$(CH2$, JI%, 1)
            Next JI%
            '
            NCHE$ = ""
            For JI% = 1 To Len(CH3$)
               If Asc(Mid$(CH3$, JI%)) > 32 Then NCHE$ = NCHE$ + Mid$(CH3$, JI%, 1)
            Next JI%
            '
            MEPAX$ = ""
            For UX% = 1 To CAMP%
               If MEPA%(UX%) = QUE% Then
                  MEPAX$ = TRIM$(DEVAL$(UX%)) + " "
                  GoTo 18
               End If
            Next UX%
            Call MENSERR(24, "Medio de Cobranza" + Str$(QUE%) + " no Definido")
            Call FINAL("")
            '
18          CUEH% = CUEMP%(UX%)
            IPASE# = CVD(CH5$)
            DTX$ = MEPAX$ + TRIM$(CHE$)
            If DTX$ = "" Then DTX$ = TRIM$(NCHE$)
            If DTX$ <> "" Then DTX$ = DTX$ + " - "
            DTX$ = DTX$ + "RB" + Str$(NOPA) + " " + RASOCLI$(NP%)
            Z$ = Space$(128)
            Call REPLA(Z$, CODCUE$(CUEH%), 1, 12)
            Call REPLA(Z$, DTX$, 14, 27)
            Call REPLA(Z$, CSTRING$(MKD$(IPASE#), 3, 12, 2, 2), 41, 12)
EPAC$ = TRIM$(UCase$(EMPREAC$))
If InStr(EPAC$, "FERVI") > 0 Then
   If RBC& < 0 Then
      RBC& = 0
   End If
End If
            If RBC& <> (-1) Then
               If RBC& < 1 Or RBC& > FBC& Then
                  Call REPLA(Z$, CSTRING$(MKS$(IAX&), 3, 8, 0, 2), 121, 8)
               End If
            End If
            List1.AddItem Z$
            '
         End If
      End If
   Next IAX&
   '
   Call SUMASIEN
   Screen.MousePointer = 1
   '
End Sub

Sub ASIFOFIJ()
   '
   Call SELECHO("SELFECHA")
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
   Text1.TEXT = FF$
   REFE$ = AJUSTI$("RESUMEN DIARIO F.FIJO DEL " + FF$, 40)
   Text4.TEXT = REFE$
   '
   ' BUSQUEDA POR FECHA Y REFERENCIA
   LI% = 0: LS% = MAXENC% + 1
   L% = MAXENC%: ERRO% = 0
   If FEAS% > FULTAS% Then
      Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
      GoTo 90
   End If
   '
82 E% = Int((LS% - LI%) / 2)
   If E% < 1 Then GoTo 90
   L% = LI% + E%
   If L% < 1 Or L% > MAXENC% Then GoTo 90
   Call LEENCAJA(L%)
   If CVI(D3$) < FEAS% Then LI% = L%: GoTo 82
   If CVI(D3$) > FEAS% Then LS% = L%: GoTo 82
   '
   ENCUAS% = 0: NROAS = 0
   For L1% = L% To 1 Step -1
      Call LEENCAJA(L1%)
      If CVI(D3$) <> FEAS% Then GoTo 83
   Next L1%
   L1% = 0
   '
83 L1% = L1% + 1
   NASPROX = 0
   For L% = L1% To MAXENC%
      Call LEENCAJA(L%)
      If CVI(D3$) <> FEAS% Then GoTo 84
      If TRIM$(D7$) = TRIM$(REFE$) Then              ' ENCONTRO EL ASIENTO
         ENCUAS% = 1
         NROAS = CVS(D1$)
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
      CMDX% = COMALTER%("Asiento de Fondo Fijo del " + FF$ + "\ya Registrado.\\Mantener\Refrescar")
      If CMDX% <> 2 Then
           Call CARGASICAJ
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
   Text4.TEXT = REFE$
   Call GENASIFOFIJ
   '
End Sub

Sub GENASIFOFIJ()
   '
   Static CAMP%, CUEMP%(2048), MEPA%(2048)
   '
   If CAMP% <= 0 Then
      For I& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", I&)
         CAMP% = I&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUE$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUE$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUE$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next I&
   End If
   '
   Screen.MousePointer = 11
   FFAA$ = ASICAJA.Text1.TEXT
   FF% = FECHANUM(FFAA$)
   '
   Dim NOPFF(2048)
   CANOPFF% = 0
   REFE$ = Left$("RESUMEN DIARIO F.FIJO DEL " + FFAA$ + Space$(40), 40)
   '
   FIN& = ULTREG&("CUECORR")
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      CC0$ = REGLEIDO$("CUECORR", I&)
      If Mid$(CC0$, 7, 2) = "FF" Then
         If CVI(Left$(CC0$, 2)) = FF% Then
            NP% = CVI(Mid$(CC0$, 5, 2))
            NOPA = Val(Mid$(CC0$, 20, 6))
            If NOPA <> NOPFF(CANOPFF%) Then
               CANOPFF% = CANOPFF% + 1
               NOPFF(CANOPFF%) = NOPA
            End If
            ITOT# = CVD(Mid$(CC0$, 115, 8))
            '
            Z$ = Space$(134)
            Call REPLA(Z$, CUEMADRE$((-1) * NP%), 1, 12)
            DETAX$ = "FF " + TRIM$(Str$(NOPA)) + " " + RASOCLI$((-1) * NP%)
            Call REPLA(Z$, DETAX$, 14, 27)
            Call REPLA(Z$, CSTRING$(MKD$(ITOT#), 3, 12, 2, 2), 41, 12)
            Call REPLA(Z$, Str$(NP%), 129, 6)
            List1.AddItem Z$
            '
         End If
      End If
   Next I&
   Call CIERRARCH("CUECORR")
   '
   NOPANT = 0
   FIN& = ULTREG&("CHEQUEMI")
   For IAX& = 1 To FIN&
      Call TRACE(20, IAX&, FIN&)
      CH0$ = REGLEIDO$("CHEQUEMI", IAX&)
      CH1$ = Mid$(CH0$, 1, 2)
      CH5$ = Mid$(CH0$, 33, 8)
      CH6$ = Mid$(CH0$, 41, 4)
      CH7$ = Mid$(CH0$, 45, 2)
      CH8$ = Mid$(CH0$, 47, 2)
      '
      If CVI(CH8$) = FF% Then                ' CORRESPONDE FECHA
         IM# = CVD(CH5$)
         If Abs(IM#) < 0.005 Then
            NPROVX% = CVI(CH1$)
            For KK11% = 1 To List1.ListCount
              Z$ = List1.List(KK11% - 1)
              If Val(Mid$(Z$, 129, 6)) = NPROVX% Then
                IM# = Val(Mid$(Z$, 41, 12))
                Call REPLA(CH0$, MKD$(IM#), 33, 8)
                Call GRAREG("CHEQUEMI", CH0$, IAX&)
                CH5$ = MKD$(IM#)
                GoTo 5
              End If
            Next KK11%
         End If
         '
5        If Abs(IM#) >= 0.005 Then
            NOPA = CVS(CH6$)
            '
            For KKK% = 1 To CANOPFF%
               If NOPFF(KKK%) = NOPA Then GoTo 10
            Next KKK%
            GoTo 76
            '
10          NP% = CVI(CH1$)         ' NRO DE PROVEEDOR
            TIPOVA% = CVI(CH7$)
            QUE% = CVI(CH7$)
      '
            For UX% = 1 To CAMP%
               If MEPA%(UX%) = QUE% Then GoTo 18
            Next UX%
            Call MENSERR(24, "Medio de Pago" + Str$(QUE%) + " no Definido")
            Call FINAL("")
            '
18          CUEH% = CUEMP%(UX%)
            IPASE# = CVD(CH5$)
            DTX$ = "LIQ. FONDO FIJO Nro." + Str$(NOPA)
            '
            Z$ = Space$(128)
            Call REPLA(Z$, CODCUE$(CUEH%), 1, 12)
            Call REPLA(Z$, DTX$, 14, 27)
            If List1.ListCount > 0 Then
               If Left$(Z$, 40) = Left$(List1.List(List1.ListCount - 1), 40) Then
                  IPASE# = IPASE# + Val(Mid$(List1.List(List1.ListCount - 1), 53, 12))
                  List1.RemoveItem (List1.ListCount - 1)
               End If
            End If
            Call REPLA(Z$, CSTRING$(MKD$(IPASE#), 3, 12, 2, 2), 53, 12)
            List1.AddItem Z$
            '
         End If
      End If
      '
76 Next IAX&
   '
   Call SUMASIEN
   Screen.MousePointer = 1
   '
End Sub

Sub ASIDECHE()
   '
   HOII% = HOY(HOS$)
   CUECHECA$ = CONTROL$("", "CUECHECA")
   CUEICHEQ% = CUEINT%(CUECHECA$)
   If TRIM$(CUECHECA$) = "" Or CUEICHEQ% < 1 Or CUEICHEQ% > NUCUEN% Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Cartera de Valores")
      OPASICONT.Show 1
      Exit Sub
   End If
   '
   If ULTREG&("CUEBANC") < 1 Then
      Call MENSERR(24, "Falta Definir Cuentas Bancarias\para Deposito de Cheques.")
      Call CARCUEBANC
      Exit Sub
   End If
   '
   VDEF$ = Space$(80)
   Call REPLA(VDEF$, MKI$(HOII%), 13, 2)
   Call REPLA(VDEF$, "DEPOSITO CHEQUES ", 15, 40)
   Call REPLA(VDEF$, CUECHECA$, 55, 12)
   Call REPLA(VDEF$, "48 Hs", 77, 10)
10 OBX$ = OBJPLA$("DEPOCHECAR", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   VDEF$ = OBX$
   CUEBA$ = Left$(OBX$, 12)
   If CUEINT%(CUEBA$) < 1 Or CUEINT(CUEBA$) > NUCUEN% Then GoTo 10
   CUECHE$ = Mid$(OBX$, 55, 12)
   FEDEPO% = CVI(Mid$(OBX$, 13, 2))
   If FEDEPO% < COMEJ% Or FEDEPO% > FINEJ% Then GoTo 10
   PLADEPO$ = TRIM$(Mid$(OBX$, 77, 10))
   HORACRE% = Val(Left$(PLADEPO$, 3))
   If HORACRE% < 24 Then
     If Left$(PLADEPO$, 2) <> "0 " Then
       Call MENSERR(24, "Plazo de Acreditación no Válido.")
       GoTo 10
     End If
   End If
   If HORACRE% Mod 24 <> 0 Then
     Call MENSERR(24, "Plazo de Acreditación no Válido.")
     GoTo 10
   End If
   '
   Screen.MousePointer = 11
   FEX = FEDEPO%
   REFE$ = TRIM$(Mid$(OBX$, 15, 40)) + " " + FECHATEX$(FEX)
   Call GRACONTROL("", "CUECHE", CUECHE$)
   '
   FORSELCHE.List1.Clear
   FORSELCHE.LICHESEL.Clear
   FIN& = ULTREG&("BADACHE")
   '
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("BADACHE", I&)
      If CVI(Mid$(X$, 133, 2)) < 1 Then
         If CVD(Mid$(X$, 49, 8)) >= 0.005 Then
            If TRIM$(Left$(X$, 12)) = "" Then
               Call REPLA(X$, "Cheque Nro:", 1, 11)
            End If
            FEVE$ = Mid$(X$, 33, 8)
            Z$ = "A" + Space$(80)
            FEXI% = FECHANUM(FEVE$)
            Call REPLA(Z$, CSTRING$(MKI$(FEXI%), 1, 5, 0, 2), 2, 5)
            Call REPLA(Z$, TRIM$(Left$(X$, 40)), 9, 40)
            Call REPLA(Z$, CSTRING$(Mid$(X$, 49, 8), 4, 12, 2, 2), 49, 12)
            Call REPLA(Z$, CSTRING$(MKS$(I&), 3, 8, 0, 2), 73, 8)
            FORSELCHE.List1.AddItem Z$
         End If
      End If
   Next I&
   '
   SUCHETO# = 0
   For I& = 1 To FORSELCHE.List1.ListCount
      Z$ = Mid$(FORSELCHE.List1.List(I& - 1), 9)
      FORSELCHE.LICHESEL.AddItem Z$
      SUCHETO# = SUCHETO# + Val(Mid$(Z$, 41, 12))
   Next I&
   '
   FORSELCHE.Label4.Caption = CSTRING$(MKD$(SUCHETO#), 4, 11, 2, 2)
   FORSELCHE.Label4.Refresh
   '
   Screen.MousePointer = 1
   FORSELCHE.Show 1
   '
   List1.Clear
   IPASE# = Val(FORSELCHE.Label2.Caption)
   If IPASE# > 0.005 Then
      FEX = FEDEPO%
      Text1.TEXT = FECHATEX(FEX)
      Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
      Text4.TEXT = REFE$
      '
      Z$ = Space$(128)
      Call REPLA(Z$, CUEBA$, 1, 12)
      Call REPLA(Z$, "DEPOSITO " + PLADEPO$, 14, 27)
      Call REPLA(Z$, CSTRING$(MKD$(IPASE#), 3, 12, 2, 2), 41, 12)
      PASDEBE$ = Z$
      'List1.AddItem Z$
      '
      SUMACHESEL# = 0
      For I& = 1 To FORSELCHE.LICHESEL.ListCount
         If FORSELCHE.LICHESEL.Selected(I& - 1) = True Then
            X$ = FORSELCHE.LICHESEL.List(I& - 1)
            BACH$ = TRIM$(Left$(X$, 16))
            NCHE$ = Mid$(X$, 19, 12)
            IPASE# = Val(Mid$(X$, 41, 12))
            SUMACHESEL# = SUMACHESEL# + IPASE#
            REGBA$ = Mid$(X$, 65, 8)
            Z$ = Space$(128)
            Call REPLA(Z$, CUECHE$, 1, 12)
            Call REPLA(Z$, BACH$ + " " + NCHE$, 14, 27)
            Call REPLA(Z$, CSTRING$(MKD$(IPASE#), 3, 12, 2, 2), 53, 12)
            Call REPLA(Z$, REGBA$, 121, 8)
            List1.AddItem Z$
         End If
      Next I&
      Call REPLA(PASDEBE$, CSTRING$(MKD$(SUMACHESEL#), 3, 12, 2, 2), 41, 12)
      List1.AddItem PASDEBE$, 0
      Call SUMASIEN
   End If
   
End Sub

Sub PRINASICAJA()
   '
   FORIPRE$ = CONTROL$("", "ASICONTA")
   '
   If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      FECHDOC$ = Text1.TEXT
      NUMEDOC$ = TRIM$(Text2.TEXT)
      TIPODOC$ = "Comprobante de Caja"
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
         TETABU1$(1, CAITAB1%) = "."  ' Str$(CAITAB1%)
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

Sub CORRNUASICAJA()
   '
   If DERACCE%("RENUASI", "Re-Numeración de Asientos de Caja") < 2 Then
      GoTo 99
   End If
   '
   NROASVIE = (Int(10 * Val(Text2.TEXT) + 0.5)) / 10
   If NROASVIE < 1 Then Exit Sub
   NROASNUE = (Int(10 * Val(Text2.TEXT) + 1.5)) / 10
   '
   Screen.MousePointer = 11
   NUVIE% = 0: NUNUE% = 0
   Call ABRENCAJA
   For I% = 1 To LOF(NX5%) / 64
      Call LEENCAJA(I%)
      If Abs(CVS(D1$) - NROASVIE) <= 0.01 Then NUVIE% = 1
      If Abs(CVS(D1$) - NROASNUE) <= 0.01 Then NUNUE% = 1
   Next I%
   '
   Screen.MousePointer = 1
   If NUVIE% <> 1 Then
      Call MENSERR(24, "Asiento No Registrado")
      GoTo 99
   End If
   '
   If NUNUE% = 1 Then
      Call MENSERR(24, "Número Siguiente Ocupado")
      GoTo 99
   End If
   '
   ASVIE$ = TRIM$(CSTRING$(MKS$(NROASVIE), 3, 10, 1, 2))
   ASNUE$ = TRIM$(CSTRING$(MKS$(NROASNUE), 3, 10, 1, 2))
   If COMALTER%("Confirma Re-Numerar Asiento de Caja\de " + ASVIE$ + " a " + ASNUE$ + ".\\Cancelar\Re-Numerar") <> 2 Then
      GoTo 99
   End If
   '
   Screen.MousePointer = 11
   For I% = LOF(NX5%) / 64 To 1 Step -1
      Call LEENCAJA(I%)
      If Abs(CVS(D1$) - NROASVIE) <= 0.01 Then
         DD00$ = D0$
         Call REPLA(DD00$, MKS$(NROASNUE), 1, 4)
         LSet D0$ = DD00$
         Put #NX5%, I%, D0$
      End If
   Next I%
   '
   Call ABRESUCAJA
   For I% = 1 To LOF(NX6%) / 64
      Call LEESUCAJA(I%)
      If Abs(CVS(D12$) - NROASVIE) <= 0.01 Then
         DD00$ = D10$
         Call REPLA(DD00$, MKS$(NROASNUE), 3, 4)
         LSet D10$ = DD00$
         Put #NX6%, I%, D10$
      End If
   Next I%
   '
   Call INDEXCAJA(1)
   '
   Close #NX5%: NX5% = 0
   Close #NX6%: NX6% = 0
   '
   For REGBA& = 1 To ULTREG&("BADACHE")
      TTXX$ = REGLEIDO$("BADACHE", REGBA&)
      TTYY$ = Mid$(TTXX$, 135, 40)
      If Left$(TTYY$, 20) = "Asiento Contable C&B" Then
         If Abs(Val(Mid$(TTYY$, 21, 12)) - NROASVIE) < 0.01 Then
            Call REPLA(TTXX$, "Asiento Contable C&B " + Str$(NROASNUE), 135, 40)
            Call GRAREG("BADACHE", TTXX$, REGBA&)
         End If
      End If
   Next REGBA&
   Call CIERRARCH("BADACHE")
   '
   Text2.TEXT = CSTRING$(MKS$(NROASNUE), 3, 8, 1, 2)
   Call CARGASICAJ
   '
99 Screen.MousePointer = 1
   '
End Sub
