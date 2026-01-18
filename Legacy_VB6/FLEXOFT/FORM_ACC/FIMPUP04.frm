VERSION 5.00
Begin VB.Form FIMPUP04 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imputación de Créditos y Anticipos a Proveedores"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
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
      Left            =   -105
      TabIndex        =   0
      Top             =   -525
      Width           =   12000
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00800000&
         Height          =   7365
         Left            =   10815
         ScaleHeight     =   7305
         ScaleWidth      =   915
         TabIndex        =   33
         Top             =   525
         Width           =   975
         Begin VB.PictureBox Picture16 
            Height          =   6630
            Left            =   0
            ScaleHeight     =   6570
            ScaleWidth      =   1005
            TabIndex        =   46
            Top             =   0
            Visible         =   0   'False
            Width           =   1065
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
               Height          =   675
               Left            =   130
               Picture         =   "FIMPUP04.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   47
               ToolTipText     =   "Aprueba - Graba la Imputación"
               Top             =   5760
               Width           =   650
            End
            Begin VB.Image Image3 
               Height          =   6700
               Left            =   0
               Picture         =   "FIMPUP04.frx":030A
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1065
            End
         End
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   140
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   5670
            Width           =   640
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   -420
               TabIndex        =   40
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.CommandButton Command2 
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
            Height          =   690
            Left            =   140
            Picture         =   "FIMPUP04.frx":26A2
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Terminar - Salir de la Aplicación"
            Top             =   200
            Width           =   650
         End
         Begin VB.CommandButton Command29 
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
            Height          =   690
            Left            =   140
            Picture         =   "FIMPUP04.frx":27EC
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Ayuda Flexoft en Línea"
            Top             =   890
            Width           =   650
         End
         Begin VB.CommandButton Command5 
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
            Height          =   790
            Left            =   140
            Picture         =   "FIMPUP04.frx":2AF6
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Aprueba - Graba la Imputación"
            Top             =   4650
            Width           =   650
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
            Left            =   140
            Picture         =   "FIMPUP04.frx":2E00
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Notas de Clientes"
            Top             =   3310
            Width           =   650
         End
         Begin VB.CommandButton Command4 
            Caption         =   "R-Cue"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   790
            Left            =   140
            Picture         =   "FIMPUP04.frx":310A
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   2520
            Width           =   650
         End
         Begin VB.Image Image1 
            Height          =   390
            Left            =   -290
            Picture         =   "FIMPUP04.frx":3414
            Top             =   6075
            Width           =   1515
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Documento a Imputar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   5145
         TabIndex        =   25
         Top             =   630
         Width           =   5520
         Begin VB.Label Label32 
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
            Height          =   330
            Left            =   1575
            TabIndex        =   45
            Top             =   1260
            Visible         =   0   'False
            Width           =   330
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
            Height          =   330
            Left            =   840
            TabIndex        =   44
            Top             =   1680
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Label Label28 
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
            Height          =   330
            Left            =   840
            TabIndex        =   43
            Top             =   1260
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "M/E:"
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
            Left            =   0
            TabIndex        =   42
            Top             =   1365
            Width           =   750
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cotiz:"
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
            Left            =   0
            TabIndex        =   41
            Top             =   1785
            Width           =   750
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
            Top             =   735
            Width           =   1275
         End
         Begin VB.Label Label25 
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   29
            Top             =   630
            Width           =   2220
         End
         Begin VB.Label Label24 
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
            Height          =   330
            Left            =   3990
            TabIndex        =   28
            Top             =   630
            Width           =   1170
         End
         Begin VB.Label Label23 
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
            Height          =   330
            Left            =   3360
            TabIndex        =   27
            Top             =   1260
            Width           =   1800
         End
         Begin VB.Label Label22 
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
            Height          =   330
            Left            =   3360
            TabIndex        =   26
            Top             =   1680
            Width           =   1800
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cuenta / Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   210
         TabIndex        =   19
         Top             =   630
         Width           =   4635
         Begin VB.Label Label12 
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
            Height          =   330
            Left            =   1155
            TabIndex        =   24
            Top             =   1680
            Width           =   3270
         End
         Begin VB.Label Label6 
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
            Height          =   330
            Left            =   210
            TabIndex        =   23
            Top             =   1680
            Width           =   750
         End
         Begin VB.Label Label3 
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
            Height          =   330
            Left            =   210
            TabIndex        =   22
            Top             =   1260
            Width           =   4215
         End
         Begin VB.Label Label2 
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
            Height          =   330
            Left            =   210
            TabIndex        =   21
            Top             =   840
            Width           =   4215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   20
            Top             =   420
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         Height          =   3400
         Left            =   210
         ScaleHeight     =   3345
         ScaleWidth      =   10395
         TabIndex        =   2
         Top             =   3045
         Width           =   10450
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C000&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   330
            Left            =   0
            TabIndex        =   8
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
               Caption         =   " Documento Numero          Emision                Importe       Vencim.                  Saldo                Acredit.       "
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
               TabIndex        =   9
               Top             =   50
               Width           =   10410
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
            BackColor       =   &H00C0C000&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5630
            Left            =   0
            TabIndex        =   3
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
               Height          =   2535
               Left            =   -30
               TabIndex        =   4
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
               Height          =   2535
               Left            =   -30
               TabIndex        =   18
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
               TabIndex        =   11
               Top             =   2660
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
               TabIndex        =   10
               Top             =   0
               Width           =   1695
            End
            Begin VB.Line Line4 
               X1              =   8400
               X2              =   8400
               Y1              =   0
               Y2              =   3045
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
               TabIndex        =   7
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
               TabIndex        =   6
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
               Left            =   30
               TabIndex        =   5
               Top             =   2660
               Width           =   1695
            End
            Begin VB.Line Line14 
               X1              =   0
               X2              =   10395
               Y1              =   2540
               Y2              =   2540
            End
            Begin VB.Line Line15 
               X1              =   0
               X2              =   10395
               Y1              =   3010
               Y2              =   3010
            End
            Begin VB.Line Line16 
               X1              =   2520
               X2              =   2520
               Y1              =   2100
               Y2              =   3045
            End
            Begin VB.Line Line17 
               X1              =   3735
               X2              =   3735
               Y1              =   2100
               Y2              =   3045
            End
            Begin VB.Line Line18 
               X1              =   5460
               X2              =   5460
               Y1              =   2100
               Y2              =   3045
            End
            Begin VB.Line Line19 
               X1              =   6720
               X2              =   6720
               Y1              =   2100
               Y2              =   3045
            End
         End
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         X1              =   10665
         X2              =   10665
         Y1              =   3045
         Y2              =   6405
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   17
         Top             =   6700
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
         TabIndex        =   16
         Top             =   6800
         Width           =   2220
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   15
         Top             =   6700
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
         TabIndex        =   14
         Top             =   6800
         Width           =   2220
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   13
         Top             =   6700
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
         TabIndex        =   12
         Top             =   6800
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
      TabIndex        =   1
      Top             =   0
      Width           =   960
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu aprobar 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu ver 
      Caption         =   "Ver"
      Begin VB.Menu resumencuenta 
         Caption         =   "Resumen de Cuenta"
      End
      Begin VB.Menu notasdecliente 
         Caption         =   "Notas de Cliente"
      End
   End
End
Attribute VB_Name = "FIMPUP04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$

Private Sub aprobar_Click()
   Call Command5_Click
End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
    Call Command5_Click
    Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NCLI$ = TRIM$(Label1.Caption)
    Call CONECRUN("*NOTAPRO/" + NCLI$, "Notas de Proveedores")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RECUEP04.Text1.TEXT = Label1.Caption
    On Error Resume Next
    RECUEP04.Show 1
    If Err Then
      On Error GoTo 0
      Hide
    End If
    On Error GoTo 0
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    TOTAVAL# = CDbl(XVALO(Label14.Caption))
    TOTACAN# = CDbl(XVALO(Label18.Caption))
    If TOTACAN# > TOTAVAL# Then
        Call COMUNI("Operación Incorrecta.\  \Las Cancelaciones no Pueden\Exceder al Documento a Aplicar.")
        Command5.Enabled = True
        Exit Sub
    End If
    '
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESACRED", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
    End If
    '
    If MODOMONEI% = 2 Then
      If XVALO(Label32) <= 1 Then
        For U& = 1 To List2.ListCount
          Z$ = List2.List(U& - 1)
          DOCU$ = Left$(Z$, 20)
          If DOCU$ <> "Anticipo no Aplicado" Then
             CANCE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
             MEMI% = XVALO(Mid$(Z$, 147, 2))
             SALDOPESI# = XVALO(Mid$(Z$, 56, 13))
             SALDODOLI# = 0
             CANCEDOL# = 0
             If MEMI% > 1 Then
                On Error Resume Next
                CANCEDOL# = CANCE# / TICAMONE(2)
                If Err Then
                   On Error GoTo 0
                   Call MENSERR(24, "Imposible Completar la Transacción Solicitada.\División por Cero en Cotización Dolar.")
                   Call FINAL("")
                End If
                On Error GoTo 0
                SALDODOLI# = XVALO(Mid$(Z$, 149, 12))
             End If
             'If CANCEDOL# > SALDODOLI# Then
             '   Call MENSERR(24, "Imposible Grabar.\Las Aplicaciones en Moneda Extranjera\Exceden el Saldo Pendiente\del Comprobante '" + TRIM$(DOCU$) + "'.")
             '   Command5.Enabled = True
             '   Exit Sub
             'End If
             If CANCEDOL# > SALDODOLI# Then
                'Call COMUNI("Esta Transacción se Aplicará al Saldo en PESOS\Dejando Cancelado el Saldo en Moneda Extranjera.")
                'CANCEDOL# = SALDODOLI#
                Call MENSERR(24, "Imposible Grabar.\Las Aplicaciones en Moneda Extranjera\Exceden el Saldo Pendiente\del Comprobante '" + TRIM$(DOCU$) + "'.")
                Label21 = ""
                Command5.Enabled = True
                Exit Sub
             End If
             If SALDOPESI# - CANCE# <= 0.005 Then ' ANULARÁ SALDO EN PESOS
               If SALDODOLI# - CANCEDOL# >= 0.005 Then ' MANTENDRÁ SALDO EN DOLARES
                 TTXX$ = "Esta Transacción Anulará el Saldo en Pesos\del Comprobante '" + TRIM$(DOCU$) + "' Manteniendo\un Saldo no Cancelado en Moneda Extranjera.\  \Se Recomienda Previamente Cancelar\el Saldo en Moneda Extranjera."
                 Call MENSERR(24, TTXX$)
                 Label21 = ""
                 Command5.Enabled = True
                 Exit Sub
               End If
             End If
          End If
        Next U&
      End If
      '
      IAJU% = CUEINT%(CONTROL$("", "CUEIVA-7"))
      If IAJU% < 1 Then
         Call MENSERR(24, "Inposible Grabar esta Imputación.\   \Falta Definir Cuenta Contable\Correspondiente a Diferencia de Cambio.")
         If DERACCE%("SETUPACR", "Configuración de Funcionamiento ACREEDORES") = 2 Then
            OPFACOM04.Show 1
         End If
         Call COMUNI("Re-Inicie Ahora la Aplicación")
         Call FINAL("")
      End If
      '
    End If
    '
    Call APERBASDAT("CABAN")
    RECILAR$ = Label25
    CoComLar$ = RECILAR$
    CodiEmp% = 0
    CodiSuc% = SUC%
    NumeClie% = 0
    NumeProv% = XVALO(Label1)
    ATX% = 1
    '
    MAPLI% = XVALO(Label32)
    If MAPLI% <= 1 Or MODOMONEI% < 2 Then
         TOTACANPES# = TOTACAN#
         On Error Resume Next
         TOTACANDOL# = TOTACAN# / TICAMONE(MAPLI%)
         If Err Then
            On Error GoTo 0
            Call MENSERR(24, "Imposible Completar la Transacción Solicitada.\División por Cero en Cotización Moneda " + TRIM$(Str$(MAPLI%)) + ".")
            Call FINAL("")
         End If
         On Error GoTo 0
      Else
         TOTACANDOL# = TOTACAN#
         TOTACANPES# = TOTACAN# * XVALO(Label16)
    End If
    '
    With SacrePen
      .FindFirst "NuProve = " & NumeProv% & " AND CODICOM_LAR = '" & RECILAR$ & "' AND SALDDebe > 0"
      If .NoMatch = False Then
        CoComCor$ = !CodiCom_Cor
        .Edit
        !SaldDebe = !SaldDebe - TOTACANPES#
        If !SaldDebe < 0 Then !SaldDebe = 0
        !SaldDebe_Dola = !SaldDebe_Dola - TOTACANDOL#
        If !SaldDebe_Dola < 0 Then !SaldDebe_Dola = 0
        .Update
        ATX% = 0
      End If
    End With
    If ATX% > 0 Then
      Call MENSERR(24, "Imposible Completar Transacción\de Imputación de Creditos")
      GoTo 99
    End If
    '
    NRECIB$ = CoComCor$
    RefApli$ = "Imputación Manual " + NRECIB$
    '
    XX% = 0: VACON% = 1: SUCANCE# = 0
    For U& = 1 To List2.ListCount
      Z$ = List2.List(U& - 1)
      DOCU$ = Left$(Z$, 20)
      If DOCU$ <> "Anticipo no Aplicado" Then
        CANCE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
        On Error Resume Next
        CANCEDOL# = CANCE# / TICAMONE(2)
        If Err Then
           On Error GoTo 0
           Call MENSERR(24, "Imposible Completar la Transacción Solicitada.\División por Cero en Cotización Dolar.")
           Call FINAL("")
        End If
        On Error GoTo 0
        MMXX% = XVALO(Label32)
        If MMXX% > 1 Then
           CANCEDOL# = CANCE#
           CANCE# = ROUND(CANCEDOL# * XVALO(Label16))
        End If
        If CANCE# > 0 Then
          DOCUCA$ = TRIM$(Left$(Z$, 18))
          ICAN# = CANCE#
          ICANDOL# = CANCEDOL#
          TOTAPLI# = 0
          '
          DocOrCor$ = ""
          TotOrig# = 0
          IvaOrig# = 0
          '
          With SacrePen
            .FindFirst "NuProve = " & NumeProv% & " AND CODICOM_LAR = '" & DOCUCA$ & "' AND SALDACRE > 0"
850           If .NoMatch = False Then
                 PARCAN# = ICAN#
                 PARCANDOL# = ICANDOL#
                 If !SaldAcre < ICAN# Then PARCAN# = !SaldAcre
                 If PARCANDOL# > !SaldAcre_Dola Then PARCANDOL# = !SaldAcre_Dola
                 .Edit
                 !SaldAcre = !SaldAcre - PARCAN#
                 !SaldAcre_Dola = !SaldAcre_Dola - PARCANDOL#
                 .Update
                 '
                 DocOrCor$ = !CodiCom_Cor
                 TotOrig# = !IBruComp
                 IIVAORIG# = 0
                 On Error Resume Next
                 IvaOrig# = !IIvaComp
                 On Error GoTo 0
                 '
                 ICAN# = ICAN# - PARCAN#
                 ICANDOL# = ICANDOL# - PARCANDOL#
                 TOTAPLI# = TOTAPLI# + PARCAN#   ' SUMA APLICACIONES -  agregado por tema bimonetario 22/03
                 If ICAN# >= 0.005 Or ICANDOL# >= 0.005 Then
                   .FindNext "NuProve = " & NumeProv% & " AND CODICOM_LAR = '" & DOCUCA$ & "' AND SALDACRE > 0"
                   GoTo 850
                 End If
            End If
          End With
          '
          SUCANCE# = SUCANCE# + TOTAPLI#
          '
          NPX% = XVALO(Label1)
          DocOrLar$ = TRIM$(Left$(Z$, 18))
          ImpApli# = CANCE#
          MonOri% = MONEMISORP%(TRIM$(Left$(Z$, 18)), NPX%)
          ImpApliOr# = CANCE#
          If MonOri% > 1 Then
             On Error Resume Next
             TPCC = XVALO(Label16)
             If TPCC < 0.005 Then TPCC = 1   ' SI NO HAY TIPO DE CAMBIO ASUME 1 = PESOS
             ImpApliOr# = ROUND#(CANCE# / TPCC)
             If Err Then
                On Error GoTo 0
                Call MENSERR(24, "Imposible Completar la Transacción Solicitada.\División por Cero Valor de Cotización.")
                Call FINAL("")
             End If
             On Error GoTo 0
          End If
          Call GRAPLIPAG
          '
          If TOTAPLI# < CANCE# - 0.005 Then
              ' acumula la diferencia para ajuste de saldo
              Z1$ = Left$(Z$ + Space$(256), 256)
              DIFAPLI$ = FORMATNUM$(CANCE# - TOTAPLI#, "F12.2")
              Call REPLA(Z1$, DIFAPLI$, 241, 12)
              List2.RemoveItem U& - 1
              List2.AddItem Z1$, U& - 1
          End If
          '
        End If
      End If
    Next U&
    '
    MONEMANT% = MONEMI%
    MONEMI% = 1
    ' graba ajuste diferencia de cambio
    CoComCor$ = "AS" + Mid$(NRECIB$, 3)
    CoComLar$ = "AS" + Mid$(RECILAR$, 3)
    FC% = FECHANUM(Label24)
    FeContab = CVDAT(FC%)
    NC1% = (-1) * XVALO(Label1)
    NC% = XVALO(Label1)
    For U5& = 1 To List2.ListCount
       DIFE1# = XVALO(Mid$(List2.List(U5& - 1), 241, 12))
       If DIFE1# >= 0.005 Then
             CUECO% = CUECOCLI%(NC1%)
          CuentaCo% = CUECO%
          NumeClie% = 0
          NumeProv% = NC%
          TiMovi = "AJUP"            ' AJUSTE DE SALDO DEUDORES
          OpMovi = "AS"              ' AJUSTE DE SALDO
          VaMovi = "CCORR"           ' CUENTA CORRIENTE
          Descrip = "Aj.Saldo - " + RASOCLI$(NC1%)
          FechEmi = CVDAT(FC%)
          FeAcredi = CVDAT(FC%)
          NumeTal = COBRACT%
          CodiEmp% = 0
          CodiSuc% = SUCUI%
          NumeCom = NRO&
          StatInit = 0
          ImpoDeb# = (-1) * DIFE1#
          Call GRACABAN
          '
          ' graba contrapartida del ajuste
          CuentaCo% = IAJU%
          TiMovi = "AJUSALP"         ' AJUSTE DE SALDO DEUDORES
          OpMovi = "AS"              ' AJUSTE DE SALDO
          VaMovi = "NO-IMP"          ' NO IMPUTADO
          Descrip = "Aj.Saldo - " + RASOCLI$(NC1%)
          ImpoDeb# = DIFE1#
          Call GRACABAN
       End If
    Next U5&
    '
    MONEMI% = MONEMANT%
    '
    Call COMUNI("Imputación Completada.")
    Hide
    '
99  VACON% = 0
    Command5.Enabled = True
End Sub

Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.Picture16.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      MARCOBARRA.Top = 5500
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If

    On Error Resume Next
    Show
    Refresh
    On Error GoTo 0
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESACRED", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    On Error Resume Next
    Hide
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
    ACREMAX# = CDbl(XVALO(Label14.Caption)) - CDbl(XVALO(Label18.Caption))
    If ACREMAX# < 0 Then ACREMAX# = 0
    IPEN$ = Mid$(Z$, 55, 14)
    If XVALO(IPEN$) > ACREMAX# Then
        IPEN$ = CSTRING$(MKD$(ACREMAX#), 3, 14, 2, 2)
    End If
    SUMACRE# = XVALO(Label18.Caption)
    If IACRE$ = "" Then
         Call REPLA(Z$, IPEN$, 70, 14)
         SUMACRE# = SUMACRE# + CDbl(XVALO(IPEN$))
       Else
         Call REPLA(Z$, Space$(14), 70, 14)
         SUMACRE# = SUMACRE# - CDbl(XVALO(IACRE$))
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
    ANTI# = CDbl(XVALO(Label14.Caption)) - SUMACRE#
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
    If TICUEPRO%((-1) * NC%) = 1 Then   ' cuenta de gastos
       Call MENSERR(24, "Cuenta de Gastos")
       Exit Sub
    End If
    '
    VACON% = 1
    ILIST2& = List2.ListIndex
    TOLIST2& = List2.TopIndex
    '
    Z$ = List2.TEXT
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    IPEN$ = Mid$(Z$, 57, 13)
    SUMACRE# = XVALO(Label18.Caption)
    SUMACRE# = SUMACRE# - CDbl(XVALO(IACRE$))
    SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(LIST2ANT$, 70, 14)))
    List2.Visible = False
    List2.AddItem LIST2ANT$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    '
    Call BLANARCH("!CARGAPLI")
    For U& = 1 To List2.ListCount
        Z$ = List2.List(U& - 1)
        NROX$ = Mid$(Z$, 1, 20)
        FEMII% = FECHANUM(Mid$(Z$, 23, 8))
        IPED# = CDbl(XVALO(Mid$(Z$, 32, 13)))
        FEVEI% = FECHANUM(Mid$(Z$, 47, 8))
        SALD# = CDbl(XVALO(Mid$(Z$, 57, 13)))
        ACRE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
        DOCU$ = Mid$(Z$, 1, 18)
        REG2& = XVALO(Mid$(Z$, 119, 8))
        IRE% = XVALO(Mid$(Z$, 127, 2))
        SALDOL# = XVALO(Mid$(Z$, 149, 12))
        '
        Y$ = REGBLAN$("!CARGAPLI")
        Call REPLA(Y$, NROX$, 1, 20)
        Call REPLA(Y$, MKI$(FEMII%), 21, 2)
        Call REPLA(Y$, MKD$(IPED#), 23, 8)
        Call REPLA(Y$, MKI$(FEVEI%), 31, 2)
        Call REPLA(Y$, MKD$(SALD#), 33, 8)
        Call REPLA(Y$, MKD$(ACRE#), 41, 8)
        Call REPLA(Y$, DOCU$, 49, 18)
        Call REPLA(Y$, MKS$(REG2&), 67, 4)
        Call REPLA(Y$, MKI$(IRE%), 71, 2)
        Call REPLA(Y$, MKD$(SALDOL#), 73, 8)
        '
        Call REGAPP("!CARGAPLI", Y$)
    Next U&
    Call CIERRARCH("!CARGAPLI")
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Imputación de Créditos y Anticipos"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture2.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture2.Left))
    '
    VTB% = VENTABU%("CARGAPLI" + ATRI$)
    '
    VACON% = 1
    If VTB% >= 0 Then
        List2.Clear
        For U& = 1 To ULTREG&("!CARGAPLI")
            Y$ = REGLEIDO$("!CARGAPLI", U&)
            NROX$ = Mid$(Y$, 1, 20)
            If TRIM$(NROX$) = "Anticipo no Aplicado" Then
               HOI% = HOY(HO$)
               ACRE# = CVD(Mid$(Y$, 41, 8))
               Call REPLA(Y$, MKI$(HOI%), 31, 2)
               Call REPLA(Y$, MKD$(ACRE#), 23, 8)
               Call REPLA(Y$, MKD$(ACRE#), 33, 8)
               Call GRAREG("!CARGAPLI", Y$, U&)
            End If
            FEMI = CVI(Mid$(Y$, 21, 2))
            IPED# = CVD(Mid$(Y$, 23, 8))
            IMPTEX$ = CSTRING$(MKD$(IPED#), 4, 13, 2, 2)
            FEVE = CVI(Mid$(Y$, 31, 2))
            SALD# = CVD(Mid$(Y$, 33, 8))
            SALTEX$ = CSTRING$(MKD$(SALD#), 4, 13, 2, 2)
            ACRE# = CVD(Mid$(Y$, 41, 8))
            '
            If ACRE# < 0 Then
                Call MENSERR(24, "No se Permite Aplicación Negativa")
                GoTo 200
            End If
            If ACRE# > SALD# Then
                Call MENSERR(24, "El Importe Acreditado Debe Ser\Inferior al Saldo del Documento")
                GoTo 200
            End If
            ACRTEX$ = CSTRING$(MKD$(ACRE#), 4, 14, 2, 2)
            DOCU$ = Mid$(Y$, 49, 18)
            REG2& = CVS(Mid$(Y$, 67, 4))
            IRE% = CVI(Mid$(Y$, 71, 2))
            REGTEX$ = CSTRING$(MKS$(REG2&), 3, 8, 0, 2)
            IRETEX$ = CSTRING$(MKI$(IRE%), 1, 2, 0, 2)
            SALDOX$ = FORMATNUM$(CVD(Mid$(Y$, 73, 8)), "F12.2")
            '
            Z$ = Space$(128)
            Call REPLA(Z$, NROX$, 1, 20)
            Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
            Call REPLA(Z$, IMPTEX$, 32, 13)
            Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
            Call REPLA(Z$, SALTEX$, 57, 13)
            Call REPLA(Z$, ACRTEX$, 70, 14)
            Call REPLA(Z$, REGTEX$, 119, 8)
            Call REPLA(Z$, IRETEX$, 127, 2)
            Call REPLA(Z$, SALDOX$, 149, 12)
            List2.AddItem Z$
            '
        Next U&
    End If
    '
    SUMACRE# = 0
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        Z$ = List2.TEXT
        SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(Z$, 70, 14)))
    Next U&
    '
    Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label10.Refresh
    Label18.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    Label18.Refresh
    Label20.Caption = CSTRING$(MKD$(XVALO(Label14.Caption) - XVALO(Label18.Caption)), 4, 14, 2, 2)
    Label20.Refresh
    '
    List2.Refresh
    VACON% = 1
    '
    List2.TopIndex = TOLIST2&
    List2.Visible = True
    If List2.ListCount > ILIST2& Then
         List2.ListIndex = ILIST2&
       Else
         Call CIERRARCH("!CARMEPA")
         Call BLANARCH("!CARMEPA")
         Call CIERRARCH("!CARGAPLI")
         Call BLANARCH("!CARGAPLI")
         Call CIERRARCH("!CARGAST")
         Call BLANARCH("!CARGAST")
         List1.Clear
         List2.Clear
         Label4.Caption = ""
         Label8.Caption = ""
         Label10.Caption = ""
         Label14.Caption = ""
         Label18.Caption = ""
         Label20.Caption = ""
         Label21.Caption = ""
         Hide
    End If
    VACON% = 0
    '
End Sub

Private Sub notasdecliente_Click()
    Call Command3_Click
End Sub

Private Sub resumencuenta_Click()
    Call Command4_Click
End Sub

'
Private Sub salir_Click()
   Call Command2_Click
End Sub
