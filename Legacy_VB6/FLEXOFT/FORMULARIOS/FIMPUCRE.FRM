VERSION 5.00
Begin VB.Form FIMPUCRE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operaciones Comerciales - Imputación de Créditos y Anticipos"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11850
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFE0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8200
      Left            =   0
      TabIndex        =   1
      Top             =   -420
      Width           =   12000
      Begin VB.CommandButton Command1 
         Caption         =   "Autom."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   10840
         Picture         =   "FIMPUCRE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Imputación Automática Recibos y Creditos Viejos."
         Top             =   3045
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox TICAMBIS 
         Enabled         =   0   'False
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
         Left            =   10640
         MaxLength       =   14
         TabIndex        =   38
         Top             =   6195
         Width           =   1065
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Documento a Imputar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   5145
         TabIndex        =   29
         Top             =   630
         Width           =   5550
         Begin VB.Label Label15 
            BackColor       =   &H00FFFFFF&
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
            Left            =   315
            TabIndex        =   37
            Top             =   1260
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo no Imputado:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   945
            TabIndex        =   36
            Top             =   1785
            Width           =   2325
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Bruto:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   945
            TabIndex        =   35
            Top             =   1365
            Width           =   2325
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Emisión:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2625
            TabIndex        =   34
            Top             =   735
            Width           =   1275
         End
         Begin VB.Label Label25 
            BackColor       =   &H00FFFFFF&
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
            Height          =   330
            Left            =   315
            TabIndex        =   33
            Top             =   630
            Width           =   2220
         End
         Begin VB.Label Label24 
            BackColor       =   &H00FFFFFF&
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
            Left            =   3990
            TabIndex        =   32
            Top             =   630
            Width           =   1170
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   3360
            TabIndex        =   31
            Top             =   1260
            Width           =   1800
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   3360
            TabIndex        =   30
            Top             =   1680
            Width           =   1800
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Cuenta / Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   210
         TabIndex        =   23
         Top             =   630
         Width           =   4635
         Begin VB.Label Label12 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1155
            TabIndex        =   28
            Top             =   1680
            Width           =   3270
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
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
            Top             =   1680
            Width           =   750
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   26
            Top             =   1260
            Width           =   4215
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   25
            Top             =   840
            Width           =   4215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Height          =   330
            Left            =   3570
            TabIndex        =   24
            Top             =   420
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         Height          =   2800
         Left            =   210
         ScaleHeight     =   2745
         ScaleWidth      =   10395
         TabIndex        =   6
         Top             =   3045
         Width           =   10450
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   330
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   10515
            Begin VB.Line Line1 
               X1              =   8400
               X2              =   8400
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line2 
               X1              =   6720
               X2              =   6720
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   " Documento             Nro      Emision                Importe       Vencim.                  Saldo                Acredit.       "
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
               Left            =   0
               TabIndex        =   13
               Top             =   50
               Width           =   10410
            End
            Begin VB.Line Line8 
               X1              =   1680
               X2              =   1680
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line9 
               X1              =   2520
               X2              =   2520
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line10 
               X1              =   3730
               X2              =   3730
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line11 
               X1              =   5460
               X2              =   5460
               Y1              =   0
               Y2              =   1050
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5630
            Left            =   0
            TabIndex        =   7
            Top             =   315
            Width           =   10515
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
               Height          =   2085
               Left            =   -30
               TabIndex        =   8
               Top             =   0
               Width           =   10450
            End
            Begin VB.ListBox List3 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2085
               Left            =   -30
               TabIndex        =   22
               Top             =   0
               Width           =   10450
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Left            =   8295
               TabIndex        =   15
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Left            =   0
               TabIndex        =   14
               Top             =   0
               Width           =   1695
            End
            Begin VB.Line Line4 
               X1              =   8400
               X2              =   8400
               Y1              =   0
               Y2              =   2415
            End
            Begin VB.Line Line12 
               X1              =   1680
               X2              =   1680
               Y1              =   1575
               Y2              =   3360
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Left            =   6615
               TabIndex        =   11
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Left            =   3570
               TabIndex        =   10
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Total General"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   25
               TabIndex        =   9
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Line Line14 
               BorderWidth     =   2
               X1              =   0
               X2              =   10395
               Y1              =   2100
               Y2              =   2100
            End
            Begin VB.Line Line15 
               X1              =   0
               X2              =   10395
               Y1              =   2415
               Y2              =   2415
            End
            Begin VB.Line Line16 
               X1              =   2520
               X2              =   2520
               Y1              =   2100
               Y2              =   2415
            End
            Begin VB.Line Line17 
               X1              =   3730
               X2              =   3730
               Y1              =   2100
               Y2              =   2415
            End
            Begin VB.Line Line18 
               X1              =   5460
               X2              =   5460
               Y1              =   2100
               Y2              =   2415
            End
            Begin VB.Line Line19 
               X1              =   6720
               X2              =   6720
               Y1              =   2100
               Y2              =   2415
            End
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "O.K."
         Enabled         =   0   'False
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
         Left            =   10840
         Picture         =   "FIMPUCRE.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Aprueba - Graba la Imputación"
         Top             =   5000
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "C.Cte."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   10840
         Picture         =   "FIMPUCRE.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Resumen de Cuenta Corriente"
         Top             =   1260
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   10840
         Picture         =   "FIMPUCRE.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Notas de Clientes"
         Top             =   2100
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10840
         Picture         =   "FIMPUCRE.frx":0A68
         TabIndex        =   3
         ToolTipText     =   "Cancela la Imputaciión"
         Top             =   760
         Width           =   855
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "T-C:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10395
         TabIndex        =   39
         ToolTipText     =   "Doble-Click para Ajustar T-C"
         Top             =   5950
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Line Line5 
         X1              =   10660
         X2              =   10660
         Y1              =   3045
         Y2              =   5800
      End
      Begin VB.Image Image2 
         Height          =   510
         Left            =   10290
         Picture         =   "FIMPUCRE.frx":0D72
         Top             =   4320
         Width           =   2010
      End
      Begin VB.Line Line24 
         X1              =   10620
         X2              =   11870
         Y1              =   4305
         Y2              =   4305
      End
      Begin VB.Line Line25 
         X1              =   10620
         X2              =   11870
         Y1              =   4830
         Y2              =   4830
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8610
         TabIndex        =   21
         Top             =   6195
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pendiente:"
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
         Left            =   6300
         TabIndex        =   20
         Top             =   6300
         Width           =   2220
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5145
         TabIndex        =   19
         Top             =   6195
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Imputado:"
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
         Left            =   2835
         TabIndex        =   18
         Top             =   6300
         Width           =   2220
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1785
         TabIndex        =   17
         Top             =   6195
         Width           =   1695
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Importe Total:"
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
         Left            =   -525
         TabIndex        =   16
         Top             =   6300
         Width           =   2220
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pos.IVA:"
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
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "FIMPUCRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$

Private Sub Command1_Click()
   Command1.Enabled = False
   '
   Screen.MousePointer = 11
   FELIVIE% = FECHANUM("31/08/01")
   For UKU& = 1 To ULTREG&("!SELDOCAP")
     RGSEL$ = REGLEIDO$("!SELDOCAP", UKU&)
     If FECHANUM(Mid$(RGSEL$, 19, 8)) <= FELIVIE% Then
       Label25.Caption = Left$(RGSEL$, 18)
       Label24.Caption = Mid$(RGSEL$, 19, 8)
       Label23.Caption = Mid$(RGSEL$, 27, 14)
       Label22.Caption = Mid$(RGSEL$, 41, 14)
       Label15.Caption = CVS(Mid$(RGSEL$, 61, 4))
       '
       Label14.Caption = Mid$(RGSEL$, 41, 14)
       Label20.Caption = Mid$(RGSEL$, 41, 14)
       Label18 = ""
       '
       For U& = 1 To List2.ListCount
           Z$ = List2.List(U& - 1)
           LIST2ANT$ = Z$
           ILISTA& = U& - 1
           IACRE$ = TRIM$(Mid$(Z$, 70, 14))
           ACREMAX# = CDbl(Val(Label14.Caption)) - CDbl(Val(Label18.Caption))
           If ACREMAX# < 0 Then ACREMAX# = 0
           IPEN$ = Mid$(Z$, 56, 14)
           If Val(IPEN$) > ACREMAX# Then
              IPEN$ = CSTRING$(MKD$(ACREMAX#), 3, 14, 2, 2)
           End If
           SUMACRE# = Val(Label18.Caption)
           If IACRE$ = "" Then
               Call REPLA(Z$, IPEN$, 70, 14)
               SUMACRE# = SUMACRE# + CDbl(Val(IPEN$))
             Else
               Call REPLA(Z$, Space$(14), 70, 14)
               SUMACRE# = SUMACRE# - CDbl(Val(IACRE$))
           End If
           VACON% = 1
           List2.AddItem Z$, ILISTA&
           List2.RemoveItem ILISTA& + 1
           List2.Refresh
           List2.ListIndex = ILISTA&
           VACON% = 0
           '
           Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
           Label10.Refresh
           Label18.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
           Label18.Refresh
           '
           ANTI# = CDbl(Val(Label14.Caption)) - SUMACRE#
           If ANTI# < 0 Then
                 Label19.Caption = "Descuentos:"
               Else
                 Label19.Caption = "Pendiente:"
           End If
           Label20.Caption = TRIM$(CSTRING$(MKD$(ANTI#), 4, 14, 2, 2))
           Label20.Refresh
           '
           If Val(Label20) < 0.005 Then GoTo 10
           '
       Next U&
       '
10     Call IMPUCRE2
     End If
     '
     List2.Clear
     List2.Refresh
     NC% = Val(Label1)
     MAXDEU& = ULTREG&("SVD202")
     RFD$ = RECFILT$("SVD202", 1, MKI$(NC%))
     '
     SUMADEU# = 0: SUMAFAC# = 0: KQ& = 0
     HAYVIEJOS% = 0: FELIVIE% = 0
     If InStr(UCase$(EMPREAC$), "DOSIVA") > 0 Then
       FELIVIE% = FECHANUM("31/08/01")
     End If
     '
     For U& = 1 To Len(RFD$) Step 4
       REG& = CVS(Mid$(RFD$, U&, 4))
       If REG& > 0 Then
         If REG& <= MAXDEU& Then
           X$ = REGLEIDO$("SVD202", REG&)
           '
           TI% = Asc(Mid$(X$, 3, 1))
           VA% = Asc(Mid$(X$, 4, 1))
           TIPOS$ = "Factura": If TI% = 0 Then TIPOS$ = "Ajuste Saldo"
                               If TI% = 1 Then TIPOS$ = "Venta Contado"
                               If TI% = 3 Then TIPOS$ = "Nota Debito"
                               If TI% = 4 Then TIPOS$ = "Anticipo"
                               If TI% = 5 Then TIPOS$ = "Desc. Caja"
                               If TI% = 6 Then TIPOS$ = "Nota Credito"
                               If TI% = 7 Then TIPOS$ = "Ajuste Saldo"
                               '
           SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
           IMPO# = CVD(Mid$(X$, 21, 8))
           If TI% <> 0 Then IMPO# = Abs(IMPO#)
           IMPO# = SG% * IMPO#
           '
           SALDO# = CVD(Mid$(X$, 11, 8))
           If TI% <> 0 Then SALDO# = Abs(SALDO#)
           SALDO# = SG% * SALDO#
           '
           If SALDO# >= 0.005 Then
             SUMAFAC# = SUMAFAC# + IMPO#
             SUMADEU# = SUMADEU# + SALDO#
             NROL& = CVS(Mid$(X$, 5, 4))
             FEMI = CVI(Mid$(X$, 19, 2))
             FEVE = CVI(Mid$(X$, 9, 2))
             NUSUCUVE% = CVI(Mid$(X$, 29, 2))
             '
             IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 14, 2, 2)
             SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
             DOCU$ = FORMDOC$(NROL&, TI%, VA%, NUSUCUVE%)
             '
             Z$ = Space$(128)
             Call REPLA(Z$, TIPOS$, 1, 14)
             Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 13, 8)
             Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
             Call REPLA(Z$, IMPTEX$, 31, 14)
             Call REPLA(Z$, FECHATEX$(FEVE), 48, 8)
             Call REPLA(Z$, SALTEX$, 56, 14)
             Call REPLA(Z$, DOCU$, 97, 18)
             Call REPLA(Z$, TRIM$(Str$(REG&)), 121, 6)
             '
             List2.AddItem Z$
             '
           End If
                '
         End If
       End If
       '
     Next U&
     '
     List2.ListIndex = -1
     List2.Refresh
     Label8.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
     Label10.Caption = ""
     List2.Visible = True
     '
   Next UKU&
   '
20 Screen.MousePointer = 1
   Command1.Enabled = True
   Command1.Visible = False
   '
   Call CIERRARCH("SVD202")
   Call CIERRARCH("APLICOB")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
   '
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    If InStr(UCase$(App.EXEName), "GES") > 0 Then
      Call FINAL("")
      Exit Sub
    End If
    Hide
    Call FRESHALL
    On Error Resume Next
    If FORECISTA.Visible = False Then
      FORECISTA.Show
      Call FRESHALL
      On Error GoTo 0
      Exit Sub
    End If
    On Error GoTo 0
    Call FINAL("")
    'Call IMPUCRE1
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NCLI$ = TRIM$(Label1.Caption)
    'Call CONECRUN("*NOTACLI/" + NCLI$, "Notas de Clientes")
    NC% = Val(Label1.Caption)
    Call NOTACLI(NC%, 2)
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    'If EXISTE%("COGEVEN.EXE") < 1 Then
    '    Call OPNOIN("")
    '    Command4.Enabled = True
    '    Exit Sub
    'End If
    If InStr(UCase$(App.EXEName), "GES") > 0 Then
      RECUEC00.Text1 = Label1
      RECUEC00.Show 1
      Exit Sub
    End If
    '
    RESCUDEU.Text1.TEXT = Label1.Caption
    PROLLAMA$ = Name
    Hide
    RESCUDEU.Show
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    TOTAVAL# = CDbl(Val(Label14.Caption))
    TOTACAN# = CDbl(Val(Label18.Caption))
    If TOTACAN# > TOTAVAL# Then
        Call COMUNI("Operación Incorrecta.\  \Las Cancelaciones no Pueden\Exceder al Documento a Aplicar.")
        Command5.Enabled = True
        Exit Sub
    End If
    '
    Call IMPUCRE2
    '
    Call IMPUCRE1
    '
End Sub


Private Sub Form_Load()
    '
    Show
    Refresh
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
    If Val(TICAMBIS) < 1 Then TICAMBIS = "   1.0000"
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If FORECISTA.Visible = False Then
      FORECISTA.Show
    End If
    On Error GoTo 0
End Sub

Private Sub List2_Click()
    '
    If VACON% <> 0 Then
        Exit Sub
    End If
    '
    Z$ = List2.TEXT
    LIST2ANT$ = Z$
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    ACREMAX# = CDbl(Val(Label14.Caption)) - CDbl(Val(Label18.Caption))
    If ACREMAX# < 0 Then ACREMAX# = 0
    IPEN$ = Mid$(Z$, 56, 14)
    If Val(IPEN$) > ACREMAX# Then
        IPEN$ = CSTRING$(MKD$(ACREMAX#), 3, 14, 2, 2)
    End If
    SUMACRE# = Val(Label18.Caption)
    If IACRE$ = "" Then
         Call REPLA(Z$, IPEN$, 70, 14)
         SUMACRE# = SUMACRE# + CDbl(Val(IPEN$))
       Else
         Call REPLA(Z$, Space$(14), 70, 14)
         SUMACRE# = SUMACRE# - CDbl(Val(IACRE$))
    End If
    VACON% = 1
    List2.AddItem Z$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    List2.Refresh
    List2.ListIndex = ILISTA&
    VACON% = 0
    '
    Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label10.Refresh
    Label18.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label18.Refresh
    '
    ANTI# = CDbl(Val(Label14.Caption)) - SUMACRE#
    If ANTI# < 0 Then
          Label19.Caption = "Descuentos:"
        Else
          Label19.Caption = "Pendiente:"
    End If
    Label20.Caption = TRIM$(CSTRING$(MKD$(ANTI#), 4, 14, 2, 2))
    Label20.Refresh
    '
End Sub
'

Private Sub List2_DblClick()
    '
    VACON% = 1
    ILIST2& = List2.ListIndex
    TOLIST2& = List2.TopIndex
    '
    Z$ = List2.TEXT
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    IPEN$ = Mid$(Z$, 56, 14)
    SUMACRE# = Val(Label18.Caption)
    SUMACRE# = SUMACRE# - CDbl(Val(IACRE$))
    SUMACRE# = SUMACRE# + CDbl(Val(Mid$(LIST2ANT$, 70, 14)))
    List2.Visible = False
    List2.AddItem LIST2ANT$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    '
    Call IMPUCRE3
    '
    VACON% = 1
    List2.TopIndex = TOLIST2&
    List2.Visible = True
    If List2.ListCount > ILIST2& Then
          List2.ListIndex = ILIST2&
        Else
          Call CIERRARCH("!CARMECO")
          Call SETULTREG("!CARMECO", 0)
          Call FRESHGRID
          Call BLANARCH("!CARMECO")
          Call CIERRARCH("!APLIRECI")
          Call BLANARCH("!APLIRECI")
          List1.Clear
          List2.Clear
          Label4.Caption = ""
          Label8.Caption = ""
          Label10.Caption = ""
          Label14.Caption = ""
          Label18.Caption = ""
          Label20.Caption = ""
          BOTNEXT.Enabled = True
          Command5.Enabled = False
          Text1.SetFocus
    End If
    VACON% = 0
    '
End Sub

Private Sub Label29_DblClick()
   TICAMBIS.Enabled = True
   TICAMBIS.SetFocus
End Sub

Private Sub TICAMBIS_Change()
   If TICAMBIS.Enabled = True Then
     VH(MONEMI%) = Val(TICAMBIS)
   End If
End Sub

Private Sub TICAMBIS_LostFocus()
   If Val(TICAMBIS) < 1 Then TICAMBIS = "1"
   TICAMBIS = FORMATNUM(Val(TICAMBIS), "F9.4")
   TICAMBIS.Enabled = False
End Sub


