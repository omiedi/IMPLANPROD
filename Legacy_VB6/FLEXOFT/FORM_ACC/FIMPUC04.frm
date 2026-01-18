VERSION 5.00
Begin VB.Form FIMPUC04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operaciones Comerciales - Imputación de Créditos y Anticipos"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   615
   ClientWidth     =   11745
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   7365
      Left            =   10815
      ScaleHeight     =   7305
      ScaleWidth      =   1230
      TabIndex        =   36
      Top             =   0
      Width           =   1290
      Begin VB.PictureBox Picture16 
         Height          =   6050
         Left            =   0
         ScaleHeight     =   5985
         ScaleWidth      =   1020
         TabIndex        =   50
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         Begin VB.Image Image3 
            Height          =   6050
            Left            =   0
            Picture         =   "FIMPUC04.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Auto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   140
         Picture         =   "FIMPUC04.frx":2398
         TabIndex        =   49
         ToolTipText     =   "Aprueba - Graba la Imputación"
         Top             =   4515
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
         Picture         =   "FIMPUC04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Resumen de Cuenta Corriente"
         Top             =   1785
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
         Picture         =   "FIMPUC04.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Notas de Clientes"
         Top             =   2580
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
         Picture         =   "FIMPUC04.frx":2CB6
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Aprueba - Graba la Imputación"
         Top             =   3675
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
         Picture         =   "FIMPUC04.frx":2FC0
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   890
         Width           =   650
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
         Picture         =   "FIMPUC04.frx":32CA
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   200
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   135
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   5040
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -420
            TabIndex        =   38
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -285
         Picture         =   "FIMPUC04.frx":3414
         Top             =   5445
         Width           =   1515
      End
   End
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
      TabIndex        =   0
      Top             =   -420
      Width           =   12000
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
         Left            =   7065
         MaxLength       =   14
         TabIndex        =   34
         Top             =   7455
         Width           =   1065
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFE0&
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
         Height          =   2115
         Left            =   5040
         TabIndex        =   25
         Top             =   630
         Width           =   5655
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
            TabIndex        =   48
            Top             =   1155
            Visible         =   0   'False
            Width           =   330
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
            TabIndex        =   47
            Top             =   1620
            Width           =   750
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
            TabIndex        =   46
            Top             =   1200
            Width           =   750
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
            TabIndex        =   45
            Top             =   1155
            Visible         =   0   'False
            Width           =   645
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
            Top             =   1575
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Label Label15 
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
            Left            =   4410
            TabIndex        =   33
            Top             =   105
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
            Left            =   1155
            TabIndex        =   32
            Top             =   1620
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
            Left            =   1155
            TabIndex        =   31
            Top             =   1200
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
            Left            =   2835
            TabIndex        =   30
            Top             =   570
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
            Left            =   525
            TabIndex        =   29
            Top             =   525
            Width           =   2220
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
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
            Left            =   4200
            TabIndex        =   28
            Top             =   525
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
            Left            =   3570
            TabIndex        =   27
            Top             =   1155
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
            Left            =   3570
            TabIndex        =   26
            Top             =   1575
            Width           =   1800
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Cuenta / Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
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
            Top             =   1575
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
            Top             =   1575
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
            Top             =   1155
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
            Top             =   735
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
            Top             =   315
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         Height          =   2800
         Left            =   210
         ScaleHeight     =   2745
         ScaleWidth      =   10395
         TabIndex        =   2
         Top             =   2950
         Width           =   10450
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFF80&
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
               TabIndex        =   9
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
               Height          =   2085
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
               Height          =   2085
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
               TabIndex        =   10
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
               Left            =   25
               TabIndex        =   5
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
         Left            =   6825
         TabIndex        =   35
         ToolTipText     =   "Doble-Click para Ajustar T-C"
         Top             =   7215
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Line Line5 
         X1              =   10660
         X2              =   10660
         Y1              =   2950
         Y2              =   5700
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
         Top             =   5985
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
         Top             =   6090
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
         Top             =   5985
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
         Top             =   6090
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
         Top             =   5985
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
         Top             =   6090
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
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu aprobarimputacion 
         Caption         =   "Aprobar Imputación"
      End
      Begin VB.Menu imputacionautomatica 
         Caption         =   "Imputación Automática"
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
      Begin VB.Menu notasdeclientes 
         Caption         =   "Notas de Clientes"
      End
   End
   Begin VB.Menu accesodirecto 
      Caption         =   "Accesos Directos"
      Begin VB.Menu resumendecuenta 
         Caption         =   "Resumen de Cuenta "
      End
   End
End
Attribute VB_Name = "FIMPUC04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$
Dim MODOMONEI%

Private Sub aprobarimputacion_Click()
   Call Command5_Click
End Sub

Sub Command1_Click()
    Command1.Enabled = False
    For U& = 1 To List2.ListCount
       If TRIM$(Mid$(List2.List(U& - 1), 70, 14)) <> "" Then
          Call MENSERR(24, "Imposible Aplicación Automática.")
          GoTo 99
       End If
    Next U&
    For U& = 1 To List2.ListCount
       List2.ListIndex = U& - 1
    Next U&
    APLIAUTO% = 1
    Command1.Enabled = True
    Call Command5_Click
    Exit Sub
    '
99  Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NCLI$ = TRIM$(Label1.Caption)
    'Call CONECRUN("*NOTACLI/" + NCLI$, "Notas de Clientes")
    NC% = XVALO(Label1.Caption)
    Call NOTACLI(NC%, 2)
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RECUEC04.Text1 = Label1
    RECUEC04.Show 1
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
    ' VERIFICAR COINCIDENCIA DE MONEDA DE EMISIÓN DE DOCUMENTOS
    MEMI% = XVALO(Label32)
    If MODOMONEI% = 2 Then
       YACAN% = 0
       For U& = 1 To List2.ListCount
         z$ = List2.List(U& - 1)
         DOCU$ = Left$(z$, 20)
         If DOCU$ <> "Anticipo no Aplicado" Then
            CANCE# = CDbl(XVALO(Mid$(z$, 70, 14)))
            If CANCE# > 0 Then
               YACAN% = YACAN% + 1
               DOCUCA$ = TRIM$(Left$(z$, 18))
               FEMIX = FECHANUM(Mid$(z$, 23, 8))
               NCX1% = XVALO(Label1)
               MORI% = MONEMISOR%(DOCUCA$, FEMIX, NCX1%)
               If MEMI% > 1 Then
                  If MORI% <> MEMI% Then
                     Call MENSERR(24, "Imposible Aplicar Sobre Comprobante\Número '" + DOCUCA$ + "'\por Diferir Moneda de Emisión.")
                     Command5.Enabled = True
                     Exit Sub
                  End If
               End If
            End If
         End If
       Next U&
       '
       If YACAN% > 1 Then
          Call MENSERR(24, "Transaccion Cancelada.\Debe Aplicar a Un Solo\Documento por Vez.")
          Command5.Enabled = True
          Exit Sub
       End If
       '
    End If
    '
    Call APERBASDAT("CABAN")
    RECILAR$ = Label25
    CoComLar$ = RECILAR$
    CodiEmp% = 0
    CodiSuc% = SUC%
    NumeClie% = XVALO(Label1)
    NumeProv% = 0
    ATX% = 1
    '
    ' DESCUENTA SALDO DEL RECIBO APLICADO
    SUDEBE# = 0
    TICAPLI# = 1
    If MODOMONEI% = 2 Then
       TICAPLI# = XVALO(Label16)
    End If
    With SaDeuPen
      .FindFirst "NuClien = " & NumeClie% & " AND CODICOM_LAR = '" & RECILAR$ & "' AND SALDAcre > 0"
      If .NoMatch = False Then
        CoComCor$ = !CodiCom_Cor
        .Edit
        !SaldAcre = !SaldAcre - ROUND#(TOTACAN# * TICAPLI#)
        If MODOMONEI% = 2 Then
           If MEMI% > 1 Then
              !SaldAcre_Dola = !SaldAcre_Dola - TOTACAN#
              If !SaldAcre_Dola < 0 Then
                 !SaldAcre_Dola = !SaldAcre / TICAPLI#
              End If
              If Left$(CoComCor$, 2) = "RB" Then
                !SaldDebe_Dola = 0
                !SaldAcre_Dola = 0
              End If
           End If
        End If
        .Update
        SUDEBE# = SUDEBE# + ROUND#(TOTACAN# * TICAPLI#)
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
    XX% = 0: VACON% = 1
    SUCANCE# = 0
    For U& = 1 To List2.ListCount
      z$ = List2.List(U& - 1)
      DOCU$ = Left$(z$, 20)
      If DOCU$ <> "Anticipo no Aplicado" Then
        CANCE# = ROUND#(TICAPLI# * CDbl(XVALO(Mid$(z$, 70, 14))))
        If CANCE# > 0 Then
          DOCUCA$ = TRIM$(Left$(z$, 18))
          ICAN# = CANCE#
          TOTAPLI# = 0
          '
          DocOrCor$ = ""
          TotOrig# = 0
          IvaOrig# = 0
          '
          Call APERBASDAT("CABAN")
          With SaDeuPen
            .FindFirst "NuClien = " & NumeClie% & " AND CODICOM_LAR = '" & DOCUCA$ & "' AND SALDdebe > 0"
850         If .NoMatch = False Then
                 PARCAN# = ICAN#: If !SaldDebe < ICAN# Then PARCAN# = !SaldDebe
                 .Edit
                 !SaldDebe = !SaldDebe - PARCAN#
                 .Update
                 TOTAPLI# = TOTAPLI# + PARCAN#   ' SUMA APLICACIONES -  agregado por tema bimonetario 22/03
                 '
                 DocOrCor$ = !CodiCom_Cor
                 TotOrig# = !IBruComp
                 IvaOrig# = 0
                 On Error Resume Next
                 IvaOrig# = !IIvaComp
                 On Error GoTo 0
                 '
                 ICAN# = ICAN# - PARCAN#
                 If ICAN# >= 0.005 Then
                   .FindNext "NuClien = " & NumeClie% & "AND CODICOM_LAR = '" & DOCUCA$ & "' AND SALDdebe > 0"
                   GoTo 850
                 End If
            End If
          End With
          '
          DocOrLar$ = TRIM$(Left$(z$, 18))
          ImpApli# = TOTAPLI#
          SUCANCE# = SUCANCE# + TOTAPLI#
          Call GRAPLICOB
          '
        End If
      End If
    Next U&
    '
    SUDEBE# = SUCANCE#
    DIFECAN# = SUDEBE# - SUCANCE#
    If DIFECAN# >= 0.005 Then
       If MODOMONEI% = 2 Then
           IAJU% = CUEINT%(CONTROL$("SVD451S", "CUEVE-9"))
          Else
           IAJU% = CUEINT%(CONTROL$("SVD451S", "CUEVE-11"))
       End If
       If IAJU% < 1 Then
          Call MENSERR(24, "Errores al Grabar este Recibo.\Debe Anularlo y Cargarlo de Nuevo.")
          GoTo 99
       End If
       '
       MONEMANT% = MONEMI%
       MONEMI% = 1
       NCX1% = XVALO(Label1)
       ' graba ajuste diferencia de cambio
       RECILA$ = Label25
       CoComCor$ = "AS-" + Right$(RECILA$, 7)
       CoComLar$ = "AS" + Mid$(RECILA$, 3)
       FeContab = CVDAT(Int((FECHANUM(Label24))))
         CUECO% = CUECOCLI%(NCX1%)
       CuentaCo% = CUECO%
       NumeClie% = NCX1%
       NumeProv% = 0
       TiMovi = "AJUD"            ' AJUSTE DE SALDO DEUDORES
       OpMovi = "AS"              ' AJUSTE DE SALDO
       VaMovi = "CCORR"           ' CUENTA CORRIENTE
       Descrip = "Aj.Saldo - " + RASOCLI$(NCX1%)
       FechEmi = FeContab
       FeAcredi = FeContab
       NumeTal = XVALO(Mid$(RECILA$, 6, 4))
       CodiEmp% = 0
       CodiSuc% = SUCUI%
       NumeCom = XVALO(Mid$(RECILA$, 11))
       StatInit = 0
       ImpoDeb# = DIFECAN#
       Call GRACABAN
       '
       ' graba contrapartida del ajuste
       CuentaCo% = IAJU%
       TiMovi = "AJUSALD"         ' AJUSTE DE SALDO DEUDORES
       OpMovi = "AS"              ' AJUSTE DE SALDO
       VaMovi = "IMP-CO"          ' NO IMPUTADO
       Descrip = "Aj.Saldo - " + RASOCLI$(NCX1%)
       ImpoDeb# = (-1) * DIFECAN#
       Call GRACABAN
       '
       MONEMI% = MONEMANT%
       '
    End If
    '
    If MODOMONEI% > 0 Then
      NCI% = NumeClie%
      RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCI%))
      If Len(RFF$) > 0 Then
        For U& = 1 To List2.ListCount
          X$ = List2.List(U& - 1)
          CANCE# = CDbl(XVALO(Mid$(X$, 70, 14)))
          If CANCE# >= 0.005 Then
            ABUSQUE$ = TRIM$(Left$(X$, 18))
            For JKLX& = 1 To Len(RFF$) Step 4
              RECABI& = CVS(Mid$(RFF$, JKLX&, 4))
              YCB$ = REGLEIDO$("CARBIMON", RECABI&)
              If TRIM$(Mid$(YCB$, 13, 18)) = ABUSQUE$ Or TRIM$(Mid$(YCB$, 45, 18)) = ABUSQUE$ Then
                DOCUCABE$ = TRIM$(Mid$(YCB$, 13, 18))
                For JKLY& = 1 To Len(RFF$) Step 4
                  RECABI& = CVS(Mid$(RFF$, JKLY&, 4))
                  YCB$ = REGLEIDO$("CARBIMON", RECABI&)
                  If TRIM$(Mid$(YCB$, 13, 18)) = DOCUCABE$ Then
                      MONEMORI% = CVI(Mid$(YCB$, 31, 2))
                      TICAMBX# = TICAMONE(MONEMORI%)
                      If MODOMONEI% = 2 Then
                         TICAMBX# = XVALO(Label16)
                      End If
                      FEDOAPLI% = FECHANUM(Label24)
                      NROO& = XVALO(Mid$(RECILA$, 11))
                      Call REPLA(YCB$, MKI$(FEDOAPLI%), 33, 2)
                      Call REPLA(YCB$, MKI$(NCI%), 35, 2)
                      Call REPLA(YCB$, Chr$(TIDOAP%), 37, 1)
                      Call REPLA(YCB$, Chr$(VADOAP%), 38, 1)
                      Call REPLA(YCB$, MKS$(NROO&), 39, 4)
                      Call REPLA(YCB$, MKI$(SUC%), 43, 2)
                      Call REPLA(YCB$, RECILAR$, 45, 18)
                      Call REPLA(YCB$, MKI$(1), 63, 2) ' RECIBO SIEMPRE EN PESOS
                        If MODOMONEI% = 2 Then
                          Call REPLA(YCB$, MKI$(XVALO(Label32)), 63, 2) ' RECIBO SIEMPRE EN PESOS
                        End If
                      Call REPLA(YCB$, MKD$(0), 65, 8)
                      Call REPLA(YCB$, MKD$(CANCE# * TICAMBX#), 73, 8)
                      Call REPLA(YCB$, MKD$(0), 81, 8)
                      Call REPLA(YCB$, MKD$(TICAMBX#), 89, 8)
                      Call REPLA(YCB$, MKD$(0), 97, 8)
                      Call REPLA(YCB$, MKD$(ROUND#(CANCE#)), 105, 8)
                      Call REPLA(YCB$, MKD$(0), 113, 8)
                        ' ESTO ES PARA QUE SI SE APLICA DOCUMENTO EN PESOS,
                        ' SOLO CANCELE EL SALDO EN PESOS
                        If MODOMONEI% = 2 And MEMI% <= 1 Then
                           Call REPLA(YCB$, String$(32, 0), 89, 32)
                        End If
                      Call REGAPP("CARBIMON", YCB$)
                      '
                      If DIFECAN# >= 0.005 Then
                        Call REPLA(YCB$, "AS" + Mid$(RECILA$, 3), 45, 18)
                        Call REPLA(YCB$, MKI$(1), 63, 2) ' RECIBO SIEMPRE EN PESOS
                        Call REPLA(YCB$, MKD$(DIFECAN#), 65, 8)
                        Call REPLA(YCB$, MKD$(0), 73, 8)
                        Call REPLA(YCB$, MKD$(0), 81, 8)
                        Call REPLA(YCB$, MKD$(TICAMBX#), 89, 8)
                        Call REPLA(YCB$, MKD$(0), 97, 8)
                        Call REPLA(YCB$, MKD$(0), 105, 8)
                        Call REPLA(YCB$, MKD$(0), 113, 8)
                        Call REGAPP("CARBIMON", YCB$)
                      End If
                      GoTo 835
                  End If
                Next JKLY&
              End If
            Next JKLX&
          End If
835     Next U&
      End If
      Call CIERRARCH("CARBIMON")
      '
      ' RECUPERAR EL SALDO EN DOLARES
      For U& = 1 To List2.ListCount
        X$ = List2.List(U& - 1)
        CANCE# = CDbl(XVALO(Mid$(X$, 70, 14)))
        If CANCE# >= 0.005 Then
          ABUSQUE$ = TRIM$(Left$(X$, 18))
          Call RECUSALDOL(ABUSQUE$, NCI%)
        End If
      Next U&
      '
    End If
    '
    If APLIAUTO% < 1 Then Call COMUNI("Imputación Completada.")
    Hide
    '
99  VACON% = 0
    Command5.Enabled = True
End Sub

Private Sub Form_Activate()
    If APLIAUTO% = 1 Then
       Call Command1_Click
    End If
End Sub

Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      '
      MARCOBARRA.Top = 5650
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If

    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
    If XVALO(TICAMBIS) < 0.1 Then TICAMBIS = "   1.0000"
    '
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
    End If
    '
    APLIAUTO% = 0
    
End Sub

Private Sub imputacionautomatica_Click()
    Call Command1_Click
End Sub

Private Sub List2_Click()
    '
    If VACON% <> 0 Then
        Exit Sub
    End If
    '
    z$ = List2.TEXT
    LIST2ANT$ = z$
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(z$, 70, 14))
    ACREMAX# = CDbl(XVALO(Label14.Caption)) - CDbl(XVALO(Label18.Caption))
    If ACREMAX# < 0 Then ACREMAX# = 0
    IPEN$ = Mid$(z$, 55, 14)
    If XVALO(IPEN$) > ACREMAX# Then
        IPEN$ = CSTRING$(MKD$(ACREMAX#), 3, 14, 2, 2)
    End If
    SUMACRE# = XVALO(Label18.Caption)
    If IACRE$ = "" Then
         Call REPLA(z$, IPEN$, 70, 14)
         SUMACRE# = SUMACRE# + CDbl(XVALO(IPEN$))
       Else
         Call REPLA(z$, Space$(14), 70, 14)
         SUMACRE# = SUMACRE# - CDbl(XVALO(IACRE$))
    End If
    VACON% = 1
    List2.AddItem z$, ILISTA&
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
    VACON% = 1
    ILIST2& = List2.ListIndex
    TOLIST2& = List2.TopIndex
    '
    z$ = List2.TEXT
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(z$, 70, 14))
    IPEN$ = Mid$(z$, 56, 14)
    SUMACRE# = XVALO(Label18.Caption)
    SUMACRE# = SUMACRE# - CDbl(XVALO(IACRE$))
    SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(LIST2ANT$, 70, 14)))
    List2.Visible = False
    List2.AddItem LIST2ANT$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    '
    Call IMPUCREVEN
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

Private Sub notasdeclientes_Click()
   Call Command3_Click
End Sub

Private Sub resumendecuenta_Click()
    Call Command4_Click
End Sub

Private Sub salir_Click()
    Call Command2_Click
End Sub

Private Sub TICAMBIS_Change()
   If TICAMBIS.Enabled = True Then
     TICACH(MONEMI%) = XVALO(TICAMBIS)
   End If
End Sub

Private Sub TICAMBIS_LostFocus()
   If XVALO(TICAMBIS) < 0.1 Then TICAMBIS = "1"
   TICAMBIS = FORMATNUM(XVALO(TICAMBIS), "F9.4")
   TICAMBIS.Enabled = False
End Sub

Sub IMPUCREVEN()
    '
    If List2.ListCount < 1 Then
        Exit Sub
    End If
    '
    Call BLANARCH("!APLIRECI")
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        z$ = List2.TEXT
        TINUMS$ = Left$(z$, 20)
        'NROL& = XVALO(Mid$(Z$, 13, 8))
        FEMII% = FECHANUM(Mid$(z$, 23, 8))
        IPED# = CDbl(XVALO(Mid$(z$, 31, 14)))
        FEVEI% = FECHANUM(Mid$(z$, 48, 8))
        SALD# = CDbl(XVALO(Mid$(z$, 56, 14)))
        ACRE# = CDbl(XVALO(Mid$(z$, 70, 14)))
        DOCU$ = Mid$(z$, 97, 18)
        REG2& = XVALO(Mid$(z$, 121, 6))
        '
        Y$ = REGBLAN$("!APLIRECI")
        Call REPLA(Y$, TINUMS$, 1, 20)
        'Call REPLA(Y$, MKS$(NROL&), 17, 4)
        Call REPLA(Y$, MKI$(FEMII%), 21, 2)
        Call REPLA(Y$, MKD$(IPED#), 23, 8)
        Call REPLA(Y$, MKI$(FEVEI%), 31, 2)
        Call REPLA(Y$, MKD$(SALD#), 33, 8)
        Call REPLA(Y$, MKD$(ACRE#), 41, 8)
        Call REPLA(Y$, DOCU$, 49, 18)
        Call REPLA(Y$, MKS$(REG2&), 67, 4)
        '
        Call REGAPP("!APLIRECI", Y$)
    Next U&
    Call CIERRARCH("!APLIRECI")
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Aplicación de Cobranza"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture2.Top + Frame1.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture2.Left))
    '
    VTB% = VENTABU%("APLIRECJ" + ATRI$)
    '
    VACON% = 1
    If VTB% >= 0 Then
        List2.Clear
        For U& = 1 To ULTREG&("!APLIRECI")
            Y$ = REGLEIDO$("!APLIRECI", U&)
            TINUMS$ = Left$(Y$, 20)
            'NROL& = CVS(Mid$(Y$, 17, 4))
            FEMI = CVI(Mid$(Y$, 21, 2))
            IPED# = CVD(Mid$(Y$, 23, 8))
            IMPTEX$ = CSTRING$(MKD$(IPED#), 4, 14, 2, 2)
            FEVE = CVI(Mid$(Y$, 31, 8))
            SALD# = CVD(Mid$(Y$, 33, 8))
            SALTEX$ = CSTRING$(MKD$(SALD#), 4, 14, 2, 2)
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
            '
            z$ = Space$(128)
            Call REPLA(z$, TINUMS$, 1, 20)
            'Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 13, 8)
            Call REPLA(z$, FECHATEX$(FEMI), 23, 8)
            Call REPLA(z$, IMPTEX$, 31, 14)
            Call REPLA(z$, FECHATEX$(FEVE), 48, 8)
            Call REPLA(z$, SALTEX$, 56, 14)
            Call REPLA(z$, ACRTEX$, 70, 14)
            Call REPLA(z$, DOCU$, 97, 18)
            Call REPLA(z$, TRIM$(Str$(REG2&)), 121, 6)
            '
            List2.AddItem z$
        Next U&
    End If
    '
    SUMACRE# = 0
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        z$ = List2.TEXT
        SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(z$, 70, 14)))
    Next U&
    '
    Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label10.Refresh
    Label18.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    Label18.Refresh
    ANTI# = CDbl(XVALO(Label14.Caption)) - SUMACRE#
    If ANTI# < 0 Then
          Label19.Caption = "Descuentos:"
        Else
          Label19.Caption = "Pendiente:"
    End If
    Label20.Caption = TRIM$(CSTRING$(MKD$(ANTI#), 4, 14, 2, 2))
    Label20.Refresh
    '
    List2.Refresh
    VACON% = 1
    '
End Sub
'
