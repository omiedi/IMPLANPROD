VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FIMPUP07 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imputación de Créditos y Anticipos a Proveedores"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   11580
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   39
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C7D9CC&
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
         Appearance      =   0  'Flat
         BackColor       =   &H0050765A&
         ForeColor       =   &H80000008&
         Height          =   7365
         Left            =   10800
         ScaleHeight     =   7335
         ScaleWidth      =   945
         TabIndex        =   33
         Top             =   525
         Width           =   975
         Begin VB.PictureBox MARCOBARRA 
            Appearance      =   0  'Flat
            BackColor       =   &H000C93DC&
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            ScaleHeight     =   165
            ScaleWidth      =   615
            TabIndex        =   46
            Top             =   6120
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H0076C9F5&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   -120
               TabIndex        =   47
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.CommandButton Command2 
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
            Height          =   690
            Left            =   120
            Picture         =   "FIMPUP07.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Terminar - Salir de la Aplicación"
            Top             =   180
            Width           =   650
         End
         Begin VB.CommandButton Command29 
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
            Height          =   690
            Left            =   120
            Picture         =   "FIMPUP07.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Ayuda Flexoft en Línea"
            Top             =   890
            Width           =   650
         End
         Begin VB.CommandButton Command5 
            Caption         =   "O.K."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   790
            Left            =   120
            Picture         =   "FIMPUP07.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Aprueba - Graba la Imputación"
            Top             =   4650
            Width           =   650
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Notas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   120
            Picture         =   "FIMPUP07.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Notas de Clientes"
            Top             =   3310
            Width           =   650
         End
         Begin VB.CommandButton Command4 
            Caption         =   "R-Cue"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   790
            Left            =   120
            Picture         =   "FIMPUP07.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   2520
            Width           =   650
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C7D9CC&
         Caption         =   "DOCUMENTO A IMPUTAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   5145
         TabIndex        =   23
         Top             =   630
         Width           =   5520
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1560
            TabIndex        =   45
            Top             =   1260
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cotiza"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -120
            TabIndex        =   44
            Top             =   1750
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -120
            TabIndex        =   43
            Top             =   1365
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Label Label29 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   720
            TabIndex        =   42
            Top             =   1260
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   720
            TabIndex        =   41
            Top             =   1680
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Label Label16 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1470
            TabIndex        =   32
            Top             =   960
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Label Label15 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   315
            TabIndex        =   31
            Top             =   960
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo N-Imp."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2145
            TabIndex        =   30
            Top             =   1770
            Width           =   1125
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Bruto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   945
            TabIndex        =   29
            Top             =   1365
            Width           =   2325
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Emisión"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2625
            TabIndex        =   28
            Top             =   735
            Width           =   1275
         End
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   315
            TabIndex        =   27
            Top             =   630
            Width           =   2220
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3990
            TabIndex        =   26
            Top             =   630
            Width           =   1170
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3360
            TabIndex        =   25
            Top             =   1260
            Width           =   1800
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3360
            TabIndex        =   24
            Top             =   1680
            Width           =   1800
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C7D9CC&
         Caption         =   "CUENTA | PROVEEDOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   210
         TabIndex        =   17
         Top             =   630
         Width           =   4635
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1155
            TabIndex        =   22
            Top             =   1680
            Width           =   3270
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   210
            TabIndex        =   21
            Top             =   1680
            Width           =   750
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   210
            TabIndex        =   20
            Top             =   1260
            Width           =   4215
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   210
            TabIndex        =   19
            Top             =   840
            Width           =   4215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   3570
            TabIndex        =   18
            Top             =   420
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3360
         Left            =   210
         ScaleHeight     =   3330
         ScaleWidth      =   10425
         TabIndex        =   2
         Top             =   3000
         Width           =   10450
         Begin VB.Frame Frame2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   0
            TabIndex        =   48
            Top             =   0
            Width           =   10515
            Begin VB.Label Label35 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   " Documento Número      Emisión      Importe  Vencimiento       Saldo   Acreditación       "
               BeginProperty Font 
                  Name            =   "Roboto Mono"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Left            =   0
               TabIndex        =   50
               Top             =   0
               Width           =   10995
            End
            Begin VB.Line Line11 
               X1              =   5460
               X2              =   5460
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line10 
               X1              =   3730
               X2              =   3730
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line9 
               X1              =   2520
               X2              =   2520
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line2 
               X1              =   6720
               X2              =   6720
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line1 
               X1              =   8400
               X2              =   8400
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label9 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   $"FIMPUP07.frx":0D72
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   0
               TabIndex        =   49
               Top             =   240
               Visible         =   0   'False
               Width           =   10995
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5625
            Left            =   0
            TabIndex        =   3
            Top             =   315
            Width           =   10515
            Begin VB.ListBox List2 
               BeginProperty Font 
                  Name            =   "Roboto Mono Medium"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   500
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2490
               ItemData        =   "FIMPUP07.frx":0E19
               Left            =   0
               List            =   "FIMPUP07.frx":0E1B
               TabIndex        =   4
               Top             =   0
               Width           =   10455
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
               Left            =   0
               TabIndex        =   16
               Top             =   0
               Width           =   10450
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Roboto Mono Medium"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   500
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Left            =   8505
               TabIndex        =   9
               Top             =   2655
               Width           =   1560
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
               TabIndex        =   8
               Top             =   0
               Visible         =   0   'False
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
               Visible         =   0   'False
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
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Total General"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
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
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Roboto Mono Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8610
         TabIndex        =   15
         Top             =   6700
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pendiente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6300
         TabIndex        =   14
         Top             =   6800
         Width           =   2220
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Roboto Mono Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5145
         TabIndex        =   13
         Top             =   6700
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Imputado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2835
         TabIndex        =   12
         Top             =   6800
         Width           =   2220
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Roboto Mono Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1785
         TabIndex        =   11
         Top             =   6700
         Width           =   1695
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Importe Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -525
         TabIndex        =   10
         Top             =   6800
         Width           =   2220
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "FIMPUP07.frx":0E1D
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "FIMPUP07.frx":0E97
      Top             =   1200
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
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "FIMPUP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$

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
    RECUEP07.Text1.TEXT = Label1.Caption
    On Error Resume Next
    RECUEP07.Show 1
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
    Call APERBASDAT("CABAN")
    
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(Control$("GESACRED", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
    End If
    
    ' VALIDACIONES
    ' ************
    If MODOMONEI% = 2 Then
      If XVALO(Label32) <= 1 Then
        For U& = 1 To List2.ListCount
          Z$ = List2.LIST(U& - 1)
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
      IAJU% = CUEINT%(Control$("", "CUEIVA-7"))
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
    ' NETEA ORDEN DE PAGO / NOTA DE CREDITO / AJUSTE
    ' **********************************************
    Call APERBASDAT("CABAN")
    RECILAR$ = Label25
    CoComLar$ = RECILAR$
    'CodiEmp% = 0
    'CodiSuc% = SUC%
    NumeClie = 0
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
         TOTACANPES# = TOTACAN# * XVALO(Label28)
    End If
    
    RECILAR$ = Label25
    CoComLar$ = RECILAR$
'    CodiEmp% = 0
'    CodiSuc% = SUC%
    NumeClie = 0
    NumeProv% = XVALO(Label1)
    ATX% = 1
    '
    Dim Sacrepen  As ADODB.Recordset
    Set Sacrepen = New ADODB.Recordset
    strsql = "Select * from SacrePen"
    strsql = strsql & " Where NuProve= " & NumeProv%
    strsql = strsql & " AND CODICOM_LAR = '" & RECILAR$ & "'"
    strsql = strsql & " AND SALDDebe > 0"
26  On Error Resume Next
    Sacrepen.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 26
    End If
    On Error GoTo 0
    With Sacrepen
      If Not (.EOF And .BOF) Then
           .MoveFirst
           CoComCor$ = !CODICOM_COR
           !SaldDebe = !SaldDebe - TOTACANPES#
           
           If !SaldDebe < 0 Then !SaldDebe = 0
           !SaldDebe_DOLA = !SaldDebe_DOLA - TOTACANDOL#
           If !SaldDebe_DOLA < 0 Then !SaldDebe_DOLA = 0
           .Update
          ATX = 0
      End If
    End With
    Sacrepen.Close
    Set Sacrepen = Nothing
    
    If ATX% > 0 Then
      Call MENSERR(24, "Imposible Completar Transacción\de Imputación de Creditos")
      GoTo 99
    End If
    '
    ' NETEA FACTURA / DEBITO
    ' **********************
    NRECIB$ = CoComCor$
    RefApli$ = "Imputación Manual " + NRECIB$
    '
    XX% = 0: VACON% = 1
    For U& = 1 To List2.ListCount
      Z$ = List2.LIST(U& - 1)
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
           CANCE# = CANCEDOL# * XVALO(Label28) ' IMPORTE EN PESOS SEGUN EL TIPO DE CAMBIO DEL CREDITO (O/PAGO, NC, AJ, ETC)
        End If
        
        If CANCE# > 0 Then
          DOCUCA$ = TRIM$(Left$(Z$, 18))
          ICAN# = CANCE#
          '
          ICANDOL# = CANCEDOL#
          TOTAPLI# = 0
          '
          DocOrCor$ = ""
          TotOrig# = 0
          IvaOrig# = 0
          '
          AplicacionEnPesosSegunFactura# = 0
          TipoDeCambioFactura# = 0
          DiferenciaPorTipoCambio# = 0
          '
          Dim SacrePenAcre As ADODB.Recordset
          Set SacrePenAcre = New ADODB.Recordset
          strsql = "Select * from SacrePen"
          strsql = strsql & " Where NuProve= " & NumeProv%
          strsql = strsql & " AND CODICOM_LAR = '" & DOCUCA$ & "'"
          strsql = strsql & " AND SALDACRE > 0"
25        On Error Resume Next
          SacrePenAcre.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 25
          End If
          On Error GoTo 0
          With SacrePenAcre
            If Not (.EOF And .BOF) Then
              .MoveFirst
              Do While Not .EOF()
                 PARCAN# = ICAN#: If !SaldAcre < ICAN# Then PARCAN# = !SaldAcre
                 
                 PARCANDOL# = ICANDOL#: If PARCANDOL# > !SaldAcre_DOLA Then PARCANDOL# = !SaldAcre_DOLA
                 
                 MonedaFactura% = MMXX%
                 If MonedaFactura% > 1 Then
                    ' SI ICAN# < PARCAN#, QUIERE DECIR QUE FALTAN PESOS EN LA FACTURA PARA COMPENZAR LA O-PAGO. ESTO ESTA REFLEJADO
                    ' EN EL SALDO EN PESOS PENDIENTE QUE TIENE LA FACTURA EN EL SACREPEN. DICHA FALTA SE CONSIDERA COMO UNA
                    ' DIFERENCIA DE CAMBIO Y POR ENDE UN AJUSTE DEL LADO DEL DEBE.
                    If ICAN# > PARCAN# Then
                        DiferenciaPorTipoCambio# = DiferenciaPorTipoCambio# + (ICAN# - PARCAN#)
                    Else
                        ' *** CALCULA LA DIFERENCIA EN PESOS CUANDO EL TIPO DE CAMBIO DE LA O/PAGO (PARCAN#) ES MENOR AL DE LA FACTURA (AplicacionEnPesosSegunFactura#).
                        ' *** PARA ESO TOMA LOS DOLARES A CANCELAR (PARCANDOL#), LOS CONVIERTE A PESOS SEGUN EL TIPO DE CAMBIO
                        ' *** DE LA FACTURA Y LO COMPARA CON LOS PESOS DE LA ORDEN DE PAGO (PARCAN#)
                        TipoDeCambioFactura# = XVALO(VALOBADA("CAJABANC", "Tipo_Cam", "CODICOM_LAR='" & DOCUCA$ & "' AND NuProve= " & NumeProv%))
                        If TipoDeCambioFactura# > 0.00005 Then AplicacionEnPesosSegunFactura# = PARCANDOL# * TipoDeCambioFactura#  ' CONVIERTE LOS DOLARES DE LA O-P EN PESOS S/TIPO DE CAMBIO DE LA FACTURA
                        If (AplicacionEnPesosSegunFactura# - PARCAN#) >= 0.01 Then DiferenciaPorTipoCambio# = DiferenciaPorTipoCambio# + (PARCAN# - AplicacionEnPesosSegunFactura#)
                    End If
                 End If
                 
                 !SaldAcre = !SaldAcre - PARCAN# + DiferenciaPorTipoCambio#
                 !SaldAcre_DOLA = !SaldAcre_DOLA - PARCANDOL#
                 .Update
                 
                 DocOrCor$ = !CODICOM_COR
                 TotOrig# = !IBruComp
                 IIVAORIG# = 0
                 On Error Resume Next
                 IvaOrig# = !IIvaComp
                 On Error GoTo 0
                 '
                 ICAN# = ICAN# - PARCAN#
                 ICANDOL# = ICANDOL# - PARCANDOL#
                 TOTAPLI# = TOTAPLI# + PARCAN#
                 If ICAN# >= 0.005 Or ICANDOL# >= 0.005 Then
                   '.FindNext "CODICOM_LAR = '" & DOCUCA$ & "' AND SALDACRE > 0"
                   'GoTo 850
                   .MoveNext
                 Else
                    Exit Do
                 End If
            Loop
            End If
          End With
          SacrePenAcre.Close
          Set SacrePenAcre = Nothing
          
          SUCANCE# = SUCANCE# + TOTAPLI#
          '
          NPX% = XVALO(Label1)
          DocOrLar$ = TRIM$(Left$(Z$, 18))
          'ImpApli# = CANCE#
          ImpApli# = TOTAPLI#
          MonOri% = MONEMISORP%(TRIM$(Left$(Z$, 18)), NPX%)
          ImpApliOr# = CANCE#
          If MonOri% > 1 Then
             On Error Resume Next
             TPCC = XVALO(Label28)
             If TPCC < 0.005 Then TPCC = 1   ' SI NO HAY TIPO DE CAMBIO ASUME 1 = PESOS
             ImpApliOr# = ROUND#(CANCE# / TPCC)
             If Err Then
                On Error GoTo 0
                Call MENSERR(24, "Imposible Completar la Transacción Solicitada.\División por Cero Valor de Cotización.")
                Call FINAL("")
             End If
             On Error GoTo 0
          End If
          '
          Call GRAPLIPAG
          '
          'If TOTAPLI# < CANCE# - 0.005 Then
          If (TOTAPLI# < CANCE# - 0.005 Or Abs(DiferenciaPorTipoCambio#) > 0.005) And MonOri% > 1 Then
              ' acumula la diferencia para ajuste de saldo
              Z1$ = Left$(Z$ + Space$(256), 256)
              DIFAPLI$ = FORMATNUM$(CANCE# - TOTAPLI#, "F12.4")
              If Abs(DiferenciaPorTipoCambio#) >= 0.01 Then DIFAPLI$ = FORMATNUM$(DiferenciaPorTipoCambio#, "F12.4")
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
    NumeroOrdenAjuste% = XVALO(VALOBADA("cajabanc", "cast(substring(CodiCom_Lar,6,4) as int)", "nuprove=" & TRIM$(Label1) & " and tipomovim='AJUSALP' and codicom_lar like '%" & Mid$(RECILAR$, 11) & "' order by valorse desc"))
    FC% = FECHANUM(Label24)
    FeContab = CVDAT(FC%)
    NC1% = (-1) * XVALO(Label1)
    NC% = XVALO(Label1)
    For U5& = 1 To List2.ListCount
       DIFE1# = XVALO(Mid$(List2.LIST(U5& - 1), 241, 12))
       If Abs(DIFE1#) >= 0.005 Then
          
          NumeroOrdenAjuste% = NumeroOrdenAjuste% + 1
          PREFIJO$ = CStr(NumeroOrdenAjuste%)
          Do While Len(PREFIJO$) < 4: PREFIJO$ = "0" & PREFIJO$: Loop
          CoComCor$ = "AS" + Mid$(NRECIB$, 3, 8) + Right(PREFIJO$, 2)
          CoComLar$ = "AS" + Mid$(RECILAR$, 3, 3) & PREFIJO$ & Mid$(RECILAR$, 10)
          
             CUECO% = CUECOCLI%(NC1%)
          CuentaCo% = CUECO%
          NumeClie = 0
          NumeProv% = NC%
          TiMovi = "AJUP"            ' AJUSTE DE SALDO DEUDORES
          OpMovi = "AS"              ' AJUSTE DE SALDO
          VaMovi = "CCORR"           ' CUENTA CORRIENTE
          Descrip = "Aj.Saldo - " + rasocli$(NC1%)
            FechaMovimiento% = FC%
            fechaFactura% = FECHANUM(Mid$(List2.LIST(U5& - 1), 23, 8))
            If fechaFactura% > FC% Then FechaMovimiento% = fechaFactura%
          FechEmi = CVDAT(FechaMovimiento%)
          FeContab = CVDAT(FechaMovimiento%)
          FeAcredi = CVDAT(FechaMovimiento%)
          NumeTal = COBRACT%
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
          Descrip = "Aj.Saldo - " + rasocli$(NC1%)
          ImpoDeb# = DIFE1#
          Call GRACABAN
          
          ' *** GRABA MOVIMIENTO EN APLIPAGO
          If DIFE1# > 0.05 Then
            ImpApli# = DIFE1#
            RefApli$ = "Aj.Saldo - " + rasocli$(NC1%)
            DocOrCor$ = "AS" + Mid$(NRECIB$, 3, 8) + Right(PREFIJO$, 2)
            DocOrLar$ = "AS" + Mid$(RECILAR$, 3, 3) & PREFIJO$ & Mid$(RECILAR$, 10) ' EL QUE APLICA
            CoComCor$ = NRECIB$                                                     ' AL QUE SE APLICA
            CoComLar$ = RECILAR$
            TotOrig# = DIFE1#
            IvaOrig# = 0
            ImpApliOr# = 0
          Else
            ImpApli# = Abs(DIFE1#)
            RefApli$ = "Aj.Saldo - " + rasocli$(NC1%)
            CoComCor$ = "AS" + Mid$(NRECIB$, 3, 8) + Right(PREFIJO$, 2)
            CoComLar$ = "AS" + Mid$(RECILAR$, 3, 3) & PREFIJO$ & Mid$(RECILAR$, 10)
                DOCU$ = TRIM$(Left$(List2.LIST(U5& - 1), 18))
            DocOrLar$ = DOCU$
            DocOrCor$ = Mid$(DOCU$, 1, 5) + Mid$(DOCU$, 9, 1) + CStr(XVALO(Mid$(DOCU$, 13, 6)))
            TotOrig# = Abs(DIFE1#)
            IvaOrig# = 0
            ImpApliOr# = 0
          End If
          
          Call GRAPLIPAG
       End If
    Next U5&
    '
    MONEMI% = MONEMANT%
    
    Call RAPLIPRO(NC%)
    
    Call COMUNI("Imputación Completada.")
    Hide
    '
99  VACON% = 0
    Command5.Enabled = True
End Sub

Private Sub Form_Load()
    '
    On Error Resume Next
'    Show
    Refresh
    On Error GoTo 0
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    'Call AplicaFuenteAnchoFijo(List2, , 10)
    
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture1)
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Hide
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
    IPEN$ = Mid$(Z$, 56, 14)
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
    Label10.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    Label10.Refresh
    Label18.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
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
        Z$ = List2.LIST(U& - 1)
        NROX$ = Mid$(Z$, 1, 20)
        FEMII% = FECHANUM(Mid$(Z$, 23, 8))
        IPED# = CDbl(XVALO(Mid$(Z$, 32, 13)))
        FEVEI% = FECHANUM(Mid$(Z$, 47, 8))
        SALD# = CDbl(XVALO(Mid$(Z$, 57, 13)))
        ACRE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
        DOCU$ = Mid$(Z$, 1, 18)
        REG2& = XVALO(Mid$(Z$, 119, 8))
        IRE% = XVALO(Mid$(Z$, 127, 2))
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
    Label20.Caption = TRIM$(CSTRING$(MKD$(XVALO(Label14.Caption) - XVALO(Label18.Caption)), 4, 14, 2, 2))
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
'
Private Sub Picture1_Click()

End Sub
