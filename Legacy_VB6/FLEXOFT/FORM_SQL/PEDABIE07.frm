VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form PEDABIE07 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos de Clientes - Formato Lineal"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   Icon            =   "PEDABIE07.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2730
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   86
      Top             =   630
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command21 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   5880
      TabIndex        =   69
      Top             =   8715
      Width           =   180
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
      Height          =   300
      Index           =   2
      Left            =   5460
      TabIndex        =   68
      Text            =   " "
      Top             =   8715
      Width           =   435
   End
   Begin VB.TextBox Eco21 
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
      Height          =   300
      Index           =   2
      Left            =   6195
      TabIndex        =   67
      Text            =   " "
      Top             =   8715
      Width           =   1800
   End
   Begin VB.TextBox Text44 
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
      Left            =   8925
      TabIndex        =   66
      Top             =   8715
      Width           =   540
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00B0D0E0&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   0
      TabIndex        =   20
      Top             =   4650
      Width           =   10800
      Begin VB.Frame Frame4 
         BackColor       =   &H00B0D0E0&
         Caption         =   "Condiciones Comerciales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505000&
         Height          =   1440
         Left            =   5670
         TabIndex        =   53
         Top             =   105
         Width           =   5055
         Begin VB.Frame Frame6 
            BackColor       =   &H00B0D0E0&
            BorderStyle     =   0  'None
            Height          =   825
            Left            =   170
            TabIndex        =   57
            Top             =   555
            Width           =   4800
            Begin VB.CommandButton Command21 
               Caption         =   "."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   1470
               TabIndex        =   63
               Top             =   165
               Width           =   180
            End
            Begin VB.TextBox TEXT21 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   1050
               TabIndex        =   62
               Text            =   " "
               Top             =   165
               Width           =   435
            End
            Begin VB.TextBox Eco21 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   1785
               TabIndex        =   61
               Text            =   " "
               Top             =   165
               Width           =   2955
            End
            Begin VB.CommandButton Command21 
               Caption         =   "."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   1470
               TabIndex        =   60
               Top             =   495
               Width           =   180
            End
            Begin VB.TextBox TEXT21 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   1050
               TabIndex        =   59
               Text            =   " "
               Top             =   495
               Width           =   435
            End
            Begin VB.TextBox Eco21 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   1785
               TabIndex        =   58
               Text            =   " "
               Top             =   495
               Width           =   2955
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "L.Precios:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   -630
               TabIndex        =   65
               Top             =   240
               Width           =   1590
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "C.Pago:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   -3150
               TabIndex        =   64
               Top             =   585
               Width           =   4110
            End
         End
         Begin VB.TextBox TICAMBIS 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3840
            MaxLength       =   14
            TabIndex        =   55
            Top             =   330
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label MONEMIS 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pesos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00005500&
            Height          =   330
            Left            =   1680
            TabIndex        =   54
            ToolTipText     =   "Doble-Click para 'Dólares'"
            Top             =   285
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "T-C:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3360
            TabIndex        =   56
            Top             =   390
            Visible         =   0   'False
            Width           =   435
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   1590
         Left            =   3150
         TabIndex        =   31
         Top             =   1680
         Visible         =   0   'False
         Width           =   2220
         Begin VB.TextBox TOTIVA 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   210
            TabIndex        =   34
            Text            =   " "
            Top             =   210
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox TOTIVA 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   210
            TabIndex        =   33
            Text            =   " "
            Top             =   525
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox TOTIVA 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   210
            TabIndex        =   32
            Text            =   " "
            Top             =   840
            Visible         =   0   'False
            Width           =   870
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00B0D0E0&
         Caption         =   "Condiciones Particulares"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505000&
         Height          =   3120
         Left            =   210
         TabIndex        =   26
         Top             =   105
         Width           =   5330
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2640
            Left            =   1155
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   315
            Width           =   4005
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Items:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   210
            TabIndex        =   36
            Top             =   2415
            Width           =   645
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   420
            TabIndex        =   35
            Top             =   2680
            Width           =   540
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Observac:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   -3045
            TabIndex        =   27
            Top             =   315
            Width           =   4110
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00B0D0E0&
         Caption         =   "Entrega"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505000&
         Height          =   1620
         Left            =   5670
         TabIndex        =   25
         Top             =   1605
         Width           =   5055
         Begin VB.TextBox LENTREGA 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   1150
            MaxLength       =   32
            TabIndex        =   83
            Top             =   555
            Width           =   3750
         End
         Begin VB.TextBox LENTREGA 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   1150
            MaxLength       =   32
            TabIndex        =   82
            Top             =   555
            Visible         =   0   'False
            Width           =   3750
         End
         Begin VB.CommandButton Command24 
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
            Left            =   1805
            TabIndex        =   77
            Top             =   1200
            Width           =   195
         End
         Begin VB.TextBox LENTREGA 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   2045
            MaxLength       =   48
            TabIndex        =   76
            Top             =   1200
            Width           =   2865
         End
         Begin VB.TextBox Text19 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1150
            TabIndex        =   75
            Top             =   1200
            Width           =   645
         End
         Begin VB.CommandButton Command25 
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
            Left            =   1805
            TabIndex        =   74
            Top             =   870
            Width           =   195
         End
         Begin VB.TextBox Text20 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1150
            TabIndex        =   73
            Top             =   870
            Width           =   645
         End
         Begin VB.TextBox LENTREGA 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   1150
            MaxLength       =   32
            TabIndex        =   72
            Top             =   240
            Width           =   3750
         End
         Begin VB.Label Label9 
            Height          =   225
            Left            =   210
            TabIndex        =   85
            Top             =   790
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label8 
            Height          =   225
            Left            =   210
            TabIndex        =   84
            Top             =   525
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label27 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2045
            TabIndex        =   81
            Top             =   870
            Width           =   2865
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   13
            Left            =   -840
            TabIndex        =   80
            Top             =   975
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Transporte:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   -840
            TabIndex        =   79
            Top             =   1305
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Domicilio:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   -840
            TabIndex        =   78
            Top             =   360
            Width           =   1905
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00B0D0E0&
      BorderStyle     =   0  'None
      Height          =   8520
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   12000
      Begin VB.CommandButton Command23 
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
         Height          =   300
         Left            =   10560
         TabIndex        =   51
         Top             =   490
         Width           =   180
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9615
         TabIndex        =   50
         TabStop         =   0   'False
         Text            =   " "
         Top             =   490
         Width           =   960
      End
      Begin VB.Timer Timer2 
         Interval        =   5000
         Left            =   735
         Top             =   0
      End
      Begin VB.CommandButton Command20 
         Height          =   720
         Left            =   5040
         Picture         =   "PEDABIE07.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Búsqueda por CUIT, Domicilio, Nombre de Fantasía, etc"
         Top             =   890
         Width           =   540
      End
      Begin VB.ListBox List1 
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
         Left            =   210
         TabIndex        =   28
         Top             =   2100
         Width           =   10515
      End
      Begin VB.CommandButton Command10 
         Height          =   260
         Left            =   420
         TabIndex        =   24
         Top             =   6855
         Width           =   2825
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   210
         ScaleHeight     =   375
         ScaleWidth      =   10455
         TabIndex        =   22
         Top             =   1700
         Width           =   10515
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo           Descripcion                                 Pr.Lista  % D    Pr.Unit"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   225
            Left            =   30
            TabIndex        =   23
            Top             =   105
            Width           =   10515
         End
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1470
         MaxLength       =   40
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1280
         Width           =   3435
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9615
         TabIndex        =   11
         Text            =   " "
         Top             =   890
         Width           =   1120
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8820
         MaxLength       =   16
         TabIndex        =   12
         Top             =   1280
         Width           =   1920
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9615
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   105
         Width           =   960
      End
      Begin VB.CommandButton Command9 
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
         Left            =   10560
         TabIndex        =   10
         Top             =   105
         Width           =   180
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5775
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   490
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5775
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1280
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2310
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   890
         Width           =   2595
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1470
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   890
         Width           =   750
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1470
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   490
         Width           =   4065
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1470
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   105
         Width           =   4065
      End
      Begin VB.CommandButton Command1 
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
         Left            =   1050
         TabIndex        =   1
         Top             =   1280
         Width           =   180
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   210
         MaxLength       =   5
         TabIndex        =   0
         Text            =   " "
         Top             =   1260
         Width           =   855
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00004040&
         Height          =   8100
         Left            =   11025
         ScaleHeight     =   8040
         ScaleWidth      =   945
         TabIndex        =   37
         Top             =   -20
         Width           =   1000
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   105
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   48
            Top             =   6300
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   49
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.CommandButton Command14 
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
            Height          =   640
            Left            =   105
            Picture         =   "PEDABIE07.frx":17A8
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Aprueba, Graba e Imprime Pedido"
            Top             =   7200
            Width           =   650
         End
         Begin VB.CommandButton Command3 
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
            Height          =   590
            Left            =   105
            Picture         =   "PEDABIE07.frx":1AB2
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   685
            Width           =   650
         End
         Begin VB.CommandButton Command13 
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
            Height          =   580
            Left            =   105
            Picture         =   "PEDABIE07.frx":1DBC
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Opciones Varias - Modificaciones y Consultas"
            Top             =   3960
            Width           =   650
         End
         Begin VB.CommandButton Command11 
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
            Height          =   580
            Left            =   105
            Picture         =   "PEDABIE07.frx":1F06
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Acceso a Base de Datos de Clientes"
            Top             =   1420
            Width           =   650
         End
         Begin VB.CommandButton Command8 
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
            Height          =   590
            Left            =   105
            Picture         =   "PEDABIE07.frx":2348
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Acceso a Maestro de Artículos"
            Top             =   2000
            Width           =   650
         End
         Begin VB.CommandButton Command7 
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
            Height          =   570
            Left            =   105
            Picture         =   "PEDABIE07.frx":2652
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Acceso a Condiciones Comerciales"
            Top             =   5380
            Width           =   650
         End
         Begin VB.CommandButton Command5 
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
            Height          =   580
            Left            =   105
            Picture         =   "PEDABIE07.frx":295C
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Configuración de Funcionamiento de Pedidos"
            Top             =   4800
            Width           =   650
         End
         Begin VB.CommandButton Command4 
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
            Height          =   580
            Left            =   105
            Picture         =   "PEDABIE07.frx":2D9E
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   2790
            Width           =   650
         End
         Begin VB.CommandButton Command2 
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
            Height          =   580
            Left            =   105
            Picture         =   "PEDABIE07.frx":30A8
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   105
            Width           =   650
         End
         Begin VB.CommandButton Command12 
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
            Height          =   580
            Left            =   105
            Picture         =   "PEDABIE07.frx":31F2
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Consultas y Re-Impresión"
            Top             =   3380
            Width           =   650
         End
         Begin VB.Image Image3 
            Height          =   390
            Left            =   -310
            Picture         =   "PEDABIE07.frx":34FC
            Top             =   6615
            Width           =   1515
         End
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vigente Hasta:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505000&
         Height          =   240
         Left            =   8010
         TabIndex        =   52
         Top             =   555
         Width           =   1560
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   11865
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "OV"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   855
         Left            =   150
         TabIndex        =   29
         Top             =   180
         Width           =   1140
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   420
         Shape           =   4  'Rounded Rectangle
         Top             =   465
         Width           =   540
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000C000&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   390
         Left            =   315
         Shape           =   4  'Rounded Rectangle
         Top             =   390
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   210
         Shape           =   4  'Rounded Rectangle
         Top             =   315
         Width           =   960
      End
      Begin VB.Line Line7 
         X1              =   210
         X2              =   10710
         Y1              =   4640
         Y2              =   4640
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505000&
         Height          =   240
         Left            =   5775
         TabIndex        =   21
         Top             =   975
         Width           =   1695
      End
      Begin VB.Line Line8 
         X1              =   10725
         X2              =   10725
         Y1              =   1700
         Y2              =   4640
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido Nro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505000&
         Height          =   270
         Left            =   8130
         TabIndex        =   19
         Top             =   945
         Width           =   1440
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "O/C o N/P:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505000&
         Height          =   300
         Left            =   7515
         TabIndex        =   18
         Top             =   1335
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505000&
         Height          =   210
         Left            =   8820
         TabIndex        =   17
         Top             =   150
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Límite Crédito:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505000&
         Height          =   210
         Left            =   5775
         TabIndex        =   16
         Top             =   195
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505000&
         Height          =   210
         Left            =   315
         TabIndex        =   15
         Top             =   945
         Width           =   960
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "PEDABIE07.frx":541E
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "PEDABIE07.frx":5652
      TabIndex        =   87
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor:"
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
      Left            =   1260
      TabIndex        =   71
      Top             =   8790
      Width           =   4110
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "% Comis:"
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
      Left            =   7980
      TabIndex        =   70
      Top             =   8790
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   105
      Picture         =   "PEDABIE07.frx":56CC
      ToolTipText     =   "Recepción Pedidos Electrónicos"
      Top             =   8715
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Menu Archi 
      Caption         =   "Archivos"
      Begin VB.Menu nueped 
         Caption         =   "Nuevo Pedido"
      End
      Begin VB.Menu MoPeCur 
         Caption         =   "Modificar Pedido en Curso"
      End
      Begin VB.Menu php0 
         Caption         =   "-"
      End
      Begin VB.Menu LiPrVe 
         Caption         =   "Listas de Precios de Venta"
      End
      Begin VB.Menu php0a 
         Caption         =   "-"
      End
      Begin VB.Menu maeart 
         Caption         =   "Maestro de Articulos"
      End
      Begin VB.Menu maecli 
         Caption         =   "Maestro de Clientes"
      End
      Begin VB.Menu php1 
         Caption         =   "-"
      End
      Begin VB.Menu sali 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Edic 
      Caption         =   "Edición"
      Begin VB.Menu MoPeCu 
         Caption         =   "Modificación de Pedido en Curso"
      End
      Begin VB.Menu AnuPed 
         Caption         =   "Anulación de Pedidos"
      End
      Begin VB.Menu RevAnu 
         Caption         =   "Revertir Anulación de Pedidos"
      End
      Begin VB.Menu php1a 
         Caption         =   "-"
      End
      Begin VB.Menu AnuItem 
         Caption         =   "Anulación por Items de Pedidos"
      End
   End
   Begin VB.Menu ConLis 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu PePeCli 
         Caption         =   "Pedidos Pendientes por Cliente"
         Begin VB.Menu PeClPan 
            Caption         =   "Por Pantalla"
         End
         Begin VB.Menu PeCliEx 
            Caption         =   "Imprimir/Exportar"
         End
      End
      Begin VB.Menu PePeCo 
         Caption         =   "Pedidos Pendientes por Código de Producto"
         Begin VB.Menu PeCoPa 
            Caption         =   "Por Pantalla"
         End
         Begin VB.Menu PeCoEx 
            Caption         =   "Imprimir/Exportar"
         End
      End
      Begin VB.Menu php1b 
         Caption         =   "-"
      End
      Begin VB.Menu EncaPe 
         Caption         =   "Encabezados de Pedidos"
      End
      Begin VB.Menu DeNumPe 
         Caption         =   "Detalle por Número de Pedido"
      End
      Begin VB.Menu php2 
         Caption         =   "-"
      End
      Begin VB.Menu ConRimp 
         Caption         =   "Consultas y Re-Impresión"
      End
   End
   Begin VB.Menu AccDirec 
      Caption         =   "Accesos Directos"
      Begin VB.Menu LiPreVen 
         Caption         =   "Listas de Precios de Ventas"
      End
      Begin VB.Menu php2a 
         Caption         =   "-"
      End
      Begin VB.Menu MaesArt 
         Caption         =   "Maestro de Articulos"
      End
      Begin VB.Menu MaesCli 
         Caption         =   "Maestro de Clientes"
      End
      Begin VB.Menu php3 
         Caption         =   "-"
      End
      Begin VB.Menu RCtaCte 
         Caption         =   "Resumen de Cuenta Corriente"
      End
      Begin VB.Menu php4 
         Caption         =   "-"
      End
      Begin VB.Menu otr 
         Caption         =   "Otros Accesos Directos"
      End
   End
   Begin VB.Menu Configura 
      Caption         =   "Configuración"
      Begin VB.Menu ConGral 
         Caption         =   "Configuración General"
      End
      Begin VB.Menu php5 
         Caption         =   "-"
      End
      Begin VB.Menu ConCom 
         Caption         =   "Condiciones Comerciales"
         Begin VB.Menu LiPrVta 
            Caption         =   "Listas de Precios "
         End
         Begin VB.Menu CoPag 
            Caption         =   "Condiciones de Pago"
         End
      End
   End
   Begin VB.Menu ayuda 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "PEDABIE07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHIPRE$

Private Sub AnuItem_Click()
Call OPPED07.Command4_Click
End Sub

Private Sub AnuPed_Click()
Call OPPED07.Command45_Click
End Sub

Private Sub ayuda_Click()
Call Command3_Click
End Sub

Private Sub Command1_Click()
   Call SELECHO("DEUDOR1")
   NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   Text1.TEXT = TRIM$(Str$(NCLIE%))
   If NCLIE% > 0 Then
     If REGICLI&(NCLIE%) > 0 Then
       Call CARDACLI
       Call CARDETPEDI
     End If
   End If
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG07/AMACLI"
   Call CONECRUN(ACONEC$, "Padron Maestro de Clientes")
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    '
    Command12.Enabled = False
    Screen.MousePointer = 11
'    Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 0 ORDER BY NROPED DESC")
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "TANIT") < 1 Then
      Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 0 ORDER BY NROPED DESC")
    Else
      Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 199000 ORDER BY NROPED DESC")
    End If

    '
'    SQLX$ = "SELECT * FROM PEDENCA ORDER BY NroPed ASC"
'    'Set PEDENCTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
'    Dim PEDENCTEMP As ADODB.Recordset
'    Set PEDENCTEMP = New ADODB.Recordset
'25  On Error Resume Next
'      PEDENCTEMP.Open SQLX$, FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 25
'   End If
'   On Error GoTo 0
'    '
'    '
'    JJ& = 0
'    With PEDENCTEMP
'      On Error Resume Next
'      .MoveFirst
'      If Err < 1 Then
'        Do While Not .EOF
'          NPX& = !NroPed
'          REFEPED$ = !Referencia
'          NCLIE% = !NroClie
'          FEX = CVINT(!FechEmis)
'          ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
'          YY$ = MKS$(NPX&) + ANULX$ + AJUSTI$(REFEPED$, 60)
'          JJ& = JJ& + 1
'          Call GRAREG("INDIPEDI", YY$, JJ&)
'        .MoveNext
'        Loop
'      End If
'    End With
'    '
'    Call SETULTREG("INDIPEDI", JJ&)
'    Call CIERRARCH("INDIPEDI")
'    Call FRESHECHO("INDIPEDI")
    Screen.MousePointer = 1
    '
3   Call SELECHO("INDIPEDI/Indice General de Pedidos (Pendientes-Cumplidos-Anulados)")
    NPX& = XVALO(VALACT1$("INDIPEDI"))
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Pedido de Cliente"
    TIDOCUX$ = "Pedido ABIERTO"
    TIDOCUY$ = "Pedido Interno"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU$ = NUDOCU$ + " "
    Screen.MousePointer = 11
    '
4   For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Or InStr(XX$, TIDOCUX$) > 4 Or InStr(XX$, TIDOCUY$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
    On Error Resume Next
    OPPED07.Command46.SetFocus
    On Error GoTo 0
    OPPED07.Show 1
    GoTo 99
    '
5   For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    l& = LI& + E&: If l& < 1 Or l& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", l&), 4))
    If DCC& < DOCUNU& Then LI& = l&: GoTo 6
    If DCC& > DOCUNU& Then LS& = l&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    GoTo 99
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While l& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", l&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      l& = l& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = l& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7170
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7170
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    EDITFORM.Command3.Enabled = False
    ' EDITFORM.Command5.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 3
    '
99  Command12.Enabled = True
    '
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   PROLLAMA$ = "ABIE"
   OPPED07.Show 1
   If MONEMI% < 1 Then MONEMI% = 1
   MONEMIS = DEMONECO$(MONEMI%)
   TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   '
   Call GRAPEDIDO(OKX%)
   If OKX% > 0 Then
     Call PRIPEDIDO    ' Imprime Orden de Venta. EN CASO DE PEDIDOS ABIERTOS NO IMPRIME O/DESPACHO
     Call BLANFORMU
     Call SETULTREG("!CARDETPE", 0)
     Call CIERRARCH("*.*")
     Text1.TEXT = ""
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
   End If
99 Command14.Enabled = True
End Sub

Private Sub Command15_Click()
   MARCONOTA.Visible = False
   Text1.TEXT = ""
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Command16_Click()
   Call MENSERR(24, "No Disponible por el Momento")
End Sub

Private Sub Command16_LostFocus()
   On Error Resume Next
   Command6.SetFocus
   On Error GoTo 0
End Sub

Private Sub Command17_Click()
   Call MENSERR(24, "No Disponible por el Momento")
End Sub

Private Sub Command18_Click()
   MARSELPRESU.Visible = False
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_LostFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub


Private Sub Command20_Click()
   Static ARGU$
   Command20.Enabled = False
10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda Cliente por Texto", ARGU$)))
   If ARGU$ <> "" Then
     Screen.MousePointer = 11
     PUNCLI$ = ""
     Call COPYSTRU("DEUDOR1", "SELECLI")
     '
     Call ABRECLIEN
     Dim Deudor12 As ADODB.Recordset
     Set Deudor12 = New ADODB.Recordset
     Deudor12.CursorLocation = adUseClient
     strSQL = "Select * from Deudor12 "
25   On Error Resume Next
     Deudor12.Open FILTSQL$(strSQL), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 25
     End If
     On Error GoTo 0
     
     On Error Resume Next
     With Deudor12
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           For KKI% = 1 To .Fields.Count
             On Error Resume Next
             VATEX$ = ""
             VATEX$ = .Fields(KKI% - 1).Value
             On Error GoTo 0
             If InStr(VATEX$, ARGU$) > 0 Then
               NCLIE% = !nume_cli
               RACLI$ = !raso_cli
               XX$ = MKI$(NCLIE%) + RACLI$
               jj& = jj& + 1
               Call GRAREG("!SELECLI", XX$, jj&)
               Exit For
             End If
           Next KKI%
         .MoveNext
         Loop
       End If
     End With
     Deudor12.Close
     Set Deudor12 = Nothing
     
     Call SETULTREG("!SELECLI", jj&)
     '
     Screen.MousePointer = 1
     If jj& < 1 Then
         Call MENSERR(24, "No se Encontraron Coincidencias")
         GoTo 10
       ElseIf jj& = 1 Then
         Text1 = TRIM$(Str$(CVI(Left$(REGLEIDO$("!SELECLI", 1), 2))))
         Call CARDACLI
         Call CARDETPEDI
       Else
         Call FRESHECHO("!SELECLI")
         Call SELECHO("!SELECLI")
         NCLIX$ = VALACT1$("!SELECLI")
         If NCLIX$ <> "" Then
           Text1 = NCLIX$
           Call CARDACLI
           Call CARDETPEDI
       End If
     End If
   End If
   Command20.Enabled = True
End Sub

Private Sub Command22_Click()
    '
    Command22.Enabled = False
    Screen.MousePointer = 11
    'Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 0 ORDER BY NROPED DESC")
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "TANIT") < 1 Then
      Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 0 ORDER BY NROPED DESC")
    Else
      Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 199000 ORDER BY NROPED DESC")
    End If
    
    '
'    SQLX$ = "SELECT * FROM PEDENCA ORDER BY NroPed DESC"
'    Set PEDENCTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
'    '
'    JJ& = 0
'    With PEDENCTEMP
'      On Error Resume Next
'      .MoveFirst
'      If Err < 1 Then
'        Do While Not .EOF
'          NPX& = !NroPed
'          NCLIE% = !NroClie
'          FEX = CVINT(!FechEmis)
'          ANULX$ = "": If !Status = 9 Then ANULX$ = "(Anul) "
'          YY$ = MKS$(NPX&) + ANULX$ + TRIM$(Mid$(XX$, 5, 48))
'          JJ& = JJ& + 1
'          Call GRAREG("INDIPEDI", YY$, JJ&)
'        .MoveNext
'        Loop
'      End If
'    End With
'    '
'    Call SETULTREG("INDIPEDI", JJ&)
    Screen.MousePointer = 1
    CONPED07.Show 1
    '
    Command22.Enabled = True
    '
End Sub

Private Sub Command23_Click()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text14.TEXT = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Command24_Click()
  '\\\OT-6-422-FG-05/10/06
   Call SELECHO("TRANSPOR")
   NTRAN% = XVALO(VALACT1$("TRANSPOR"))
   Text19 = TRIM$(FORMATNUM(NTRAN%, "I4"))
  '\\\OT-6-422-FG-FIN
End Sub

Private Sub Command25_Click()
    '\\\OT-6-422-FG-05/10/06
    Call SELECHO("!SUCENTRE")
    VDEVU$ = VALACT1$("!SUCENTRE")
    Text20 = VDEVU$
    '\\\OT-6-422-FG-FIN
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("PRESUP01")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RECUEC07.Text1.TEXT = Text1.TEXT
    RECUEC07.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call CIERRARCH("*.*")
    If (USNBR% Mod 100) <> 1 Then
      Call COMUNI("La Opción Elegida Requiere\Privilegios de Supervisor.")
      Exit Sub
    End If
    CFGPED07.Show 1
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   MARCONOTA.Visible = False
   Call SELPRESUPEN
End Sub
'
Sub SELPRESUPEN()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   'UREPRE& = ULTREG&("PRESUPU")
   Screen.MousePointer = 11
   List2.Clear
   NCLIE% = XVALO(Text1.TEXT)
   If UCase$(Left$(Label11, 6)) = "REPARA" Then GoTo 50
   '
   SQLX$ = "SELECT * FROM PRESUVEN "
   SQLX$ = SQLX$ + "WHERE NroClie = " & NCLIE% & " "
   SQLX$ = SQLX$ + "AND NroPedi < 1 "
   SQLX$ = SQLX$ + "ORDER BY NroPres ASC"
   '
   NPX& = 0
   Call ABREPRESVEN
   Set IPRESPEN = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   With IPRESPEN
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       Do While Not .EOF
         NRO& = !NROPRES
           TTXX$ = "Presup.No: " + TRIM$(Str$(NRO&))
         FEX = CVINT(!FECHEMIS)
           TTXX$ = TTXX$ + " - " + FECHATEX$(FEX)
         If TRIM$(!Referen) <> "" Then
           TTXX$ = TTXX$ + " - Ref: " + TRIM$(!Referen)
         End If
         CIXI% = !CODIINT
         '
         If NRO& = NPX& Then
           TTXX$ = TTXX$ + " - VARIOS"
           List2.RemoveItem (List2.ListCount - 1)
           GoTo 15
         End If
         '
         NPX& = NRO&
         CAPRESU& = !CANTIT
         If CAPRESU& < 1 Then CAPRESU& = 1
         TTXX$ = TTXX$ + " - " + TRIM$(CODEXT$(CIXI%)) + " (" + TRIM$(Str$(CAPRESU&)) + ") " + DESCRIT0$(CIXI%)
         '
15       TTXX$ = AJUSTI$(TTXX$, 256)
         Call REPLA(TTXX$, TRIM$(Str$(NRO&)), 251, 6)
         List2.AddItem TTXX$
         '
19     .MoveNext
       Loop
     End If
   End With
   GoTo 90
   '
50 UREPRE& = ULTREG&("PRESREPA")
   Screen.MousePointer = 11
   RFF$ = RECFILT$("PRESREPA", 4, MKI$(NCLIE%))
   For KKI% = 1 To Len(RFF$) Step 4
      REPRESU& = CVS(Mid$(RFF$, KKI%, 4))
      If REPRESU& > 0 Then
        If REPRESU& <= UREPRE& Then
          XX$ = REGLEIDO$("PRESREPA", REPRESU&)
          If CVI(Mid$(XX$, 121, 2)) > 0 Then
            If Abs(CVS(Mid$(XX$, 125, 4))) < 0.5 Then
              TTXX$ = "Reparación No: " + TRIM$(CSTRING$(Left$(XX$, 4), 3, 8, 0, 2))
              TTXX$ = TTXX$ + " - " + Mid$(XX$, 5, 60)
              TTXX$ = AJUSTI$(TTXX$, 128)
              Call REPLA(TTXX$, Str$(REPRESU&), 121, 8)
              List2.AddItem TTXX$
            End If
          End If
        End If
      End If
   Next KKI%
   '
90 Screen.MousePointer = 1
   '
   If List2.ListCount < 1 Then
      Call SETULTREG("!CARDETPE", 0)
      Call CIERRARCH("!CARDETPE")
      Call CARDETPEDI
      Exit Sub
   End If
   '
End Sub
'
 Private Sub Command7_Click()
    Command7.Enabled = False
    Call CIERRARCH("*.*")
    MODX% = ALTERNA%("Listas de Precios\Condiciones de Pago")
    If MODX% = 1 Then
         Call CONECRUN("LPREAC07", "Listas de Precios de Venta")
       ElseIf MODX% = 2 Then
         Call CARCONPAGOS
    End If
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CIERRARCH("*.*")
    ACONEC$ = "ARCHIG07/AMASTO"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text9.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text9.TEXT)
End Sub



Private Sub ConGral_Click()
Call Command5_Click
End Sub

Private Sub ConRimp_Click()
Call Command12_Click
End Sub

Private Sub CoPag_Click()
MODO% = 2
Call MODOCON(MODO%)
End Sub

Private Sub DeNumPe_Click()
Call OPPED07.Command63_Click
End Sub

Private Sub Eco21_GotFocus(Index As Integer)
   On Error Resume Next
   TEXT21(Index).SetFocus
   On Error GoTo 0
End Sub


Private Sub EncaPe_Click()
Call OPPED07.Command61_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  While WindowState <> 0
    WindowState = 0
    Openforms = DoEvents
    AppActivate Caption
    SendKeys "%" & Chr$(32) & Chr$(13), True
  Wend
End Sub

Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    Frame4.Height = 1440
    Openforms = DoEvents
    '
    DoEvents
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    
    MONEMI% = 1  ' PESOS
    MONEMIS = DEMONECO$(MONEMI%)
    MONEMIS.Visible = True
    CONVERPRECIO% = 1
    TICAMBIS.Visible = True
    Label12.Visible = True
    'Image1.Visible = True
    TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
    '
    If ECOPRECIO% < 0 Then
      Label22.Caption = "  Cant Código          Descripción                                           Entrega "
      ' Image1.Visible = False
    End If
    '
    Call APLICACIONSKINS(Me, Picture2)
    
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$
    Text11.TEXT = TRIM$(Str$(NUPEDIDO&))
    '
    CMO$ = TRIM$(Control$("", "CODPECLI"))
    If CMO$ <> "" Then Label10.Caption = CMO$
    DoEvents
    '

    Call SETULTREG("!CARDETPE", 0)
    '
End Sub
'

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Image1_DblClick()
   Static CTXX%
   '
   If ULTREG&("DAPIBCLI") < 1 Then
     Call MENSERR(24, "Imposible desde esta Aplicación")
     Exit Sub
   End If
   '
   NCLIE% = XVALO(TRIM$(Text1.TEXT))
   If NCLIE% > 0 Then
     If rasocli$(NCLIE%) <> "" Then
       '
       If CTXX% < 1 Then
         Call GRAREG("TACATIBC", Chr$(0) + "Exento Percepcion IB", 1)
         Call GRAREG("TACATIBC", Chr$(1) + "Sujeto Percepcion IB", 2)
         Call SETULTREG("TACATIBC", 2)
         CTXX% = 1
       End If
       '
       RECORD$ = FILTERGETRECORD$("DAPIBCLI", 1, MKI$(NCLIE%))
       Call REPLA(RECORD$, MKI$(NCLIE%), 1, 2)
       CATIB1% = Asc(Mid$(RECORD$, 3, 1))
       If CATIB1% < 0 Or CATIB1% > 1 Then
         Call REPLA(RECORD$, Chr$(1), 3, 1)
         Call REPLA(RECORD$, Space$(32), 9, 32)
       End If
       '
10     OBX$ = OBJPLA$("CARDAPERIB", RECORD$)
       If OBX$ <> "" Then
         RECORD$ = OBX$
         CATIB1% = Asc(Mid$(RECORD$, 3, 1))
         If CATIB1% < 0 Or CATIB1% > 1 Then
           Call MENSERR(24, "Mal Categoría I.Brutos")
           GoTo 10
         End If
         If CATIB1% = 0 Then
           If TRIM$(Mid$(RECORD$, 9, 32)) = "" Then
             Call MENSERR(24, "Mandatorio Indicar Motivo de Exención")
             GoTo 10
           End If
         End If
         If CATIB1% = 1 Then
           Call REPLA(RECORD$, Space$(32), 9, 32)
         End If
         PORCONVE = CVS(Mid$(RECORD$, 5, 4))
         If PORCONVE < 0 Or PORCONVE > 100 Then
           Call MENSERR(24, "% Convenio Multilateral\Fuera de Rango")
           GoTo 10
         End If
         '
         Call REPLA(RECORD$, MKI$(NCLIE%), 1, 2)
         Call REPLA(RECORD$, USERNAME$, 41, 20)
         Call REPLA(RECORD$, MKI$(HOY(HOS$)), 61, 2)
         Call REPLA(RECORD$, MKI$(Int(HORANUM(Time$))), 63, 2)
         Call FILTERPUTRECORD("DAPIBCLI", 1, MKI$(NCLIE%), RECORD$)
         Call CIERRARCH("DAPIBCLI")
         '
       End If
     End If
   End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If ECOPRECIO >= 0 Then
     Frame4.Height = 3120
   End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Frame4.Height = 1440
End Sub

Private Sub Image2_DblClick()
   Call CIERRARCH("*.*")
   Call CONECRUN("REPEDREM", "Recepción Pedidos Remotos por Mail")
End Sub

Private Sub Image4_Click()

End Sub

Private Sub Image4_DblClick()
   PEDABIER.Show 1
End Sub

Private Sub Label11_DblClick()
   If TRIM$(UCase$(Label11)) = "VENTA" Then
       Label11.ForeColor = RGB(255, 0, 0)
       Label11 = "Reparación"
       Label11.ToolTipText = "Doble-Click para 'Venta'"
     Else
       Label11.ForeColor = RGB(0, 0, 255)
       Label11 = "Venta"
       Label11.ToolTipText = "Doble-Click para 'Reparación'"
   End If
End Sub


Private Sub Label8_Click()
   On Error Resume Next
   Text13.SetFocus
   On Error GoTo 0
End Sub

Private Sub Label22_DblClick()
   Call CARDETPEDI
End Sub

Private Sub LENTREGA_GotFocus(Index As Integer)
   '\\\OT-6-422-FG-05/10/06
   If Index = 2 Then
     If XVALO(Text19) > 0 Then
       Call MENSERR(24, "Campo no Editable")
       On Error Resume Next
       Text19.SetFocus
       On Error GoTo 0
       Exit Sub ' LE AGREGUE ESTA LINEA PARA QUE NO SE EJECUTE EL RESTO DEL CODIGO
     End If
   End If
   '\\\OT-6-422-FG-FIN
   '
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   LENTREGA(Index).TEXT = TRIM$(LENTREGA(Index).TEXT)
   LENTREGA(Index).SelStart = 0
   LENTREGA(Index).SelLength = Len(LENTREGA(Index).TEXT)
End Sub

Private Sub LiPreVen_Click()
   Call CIERRARCH("*.*")
   Call CONECRUN("LPREAC07", "Listas de Precios de Venta")
End Sub

Private Sub LiPrVe_Click()
   Call CIERRARCH("*.*")
   Call CONECRUN("LPREAC07", "Listas de Precios de Venta")
End Sub

Private Sub LiPrVta_Click()
   MODO% = 1
   Call MODOCON(MODO%)
End Sub

Private Sub List1_DblClick()
   Call CARDETPEDI
End Sub

Private Sub maeart_Click()
   Call Command11_Click
End Sub

Private Sub maecli_Click()
   Call Command8_Click
End Sub

Private Sub MaesArt_Click()
   Call Command11_Click
End Sub

Private Sub MaesCli_Click()
   Call Command8_Click
End Sub

Private Sub MONEMIS_Change()
Exit Sub
   MONEMA% = MONEMI%
   If Left$(MONEMIS, 1) = "D" Then
        MONEMI% = 2
        MONEMIS.ForeColor = RGB(150, 0, 0)
        MONEMIS.ToolTipText = "Doble-Click para 'Pesos'"
        'Label12.Visible = True
        If TICAMBIS.Visible = True Then TICAMBIS.Enabled = True
        TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
      Else
        MONEMI% = 1
        MONEMIS.ForeColor = RGB(0, 100, 0)
        MONEMIS.ToolTipText = "Doble-Click para 'Dólares'"
        'Label12.Visible = False
        If TICAMBIS.Visible = True Then TICAMBIS.Enabled = False
   End If
   '
   COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
   FIN& = ULTREG&("!CARDETPE")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDETPE", U&)
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     PRECARGA = COEFCON * CVS(Mid$(X$, 39, 4))
     PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
     TOT# = CDbl(CAN * PREUNI)
     '
     Call REPLA(X$, MKS$(PRECARGA), 39, 4)
     Call REPLA(X$, MKS$(PRECARGA), 51, 4)
     Call REPLA(X$, MKD$(TOT#), 55, 8)
     Call REPLA(X$, MKS$(PRECARGA), 65, 4)
     Call REPLA(X$, MKS$(PREUNI), 73, 4)
     Call REPLA(X$, MKD$(TOT#), 77, 8)
     Call GRAREG("!CARDETPE", X$, U&)
   Next U&
   '
   Call CARLISPE
   Call CIERRARCH("!CARDETPE")
   '
End Sub

Private Sub MONEMIS_DblClick()
'   If TRIM$(UCase$(MONEMIS)) = "PESOS" Then
'       MONEMIS = "Dólares"
'     Else
'       MONEMIS = "Pesos"
'   End If
   Static CAMON%, DMONE$(), TCMONE()
   If CAMON% < 1 Then
     CAMON% = ULTREG&("TAMONED")
     If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
           VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
           DMONE$(U&) = DMX$
           TCMONE(U) = VHX
         Next U&
     End If
   End If
   '
   MONEMIA% = MONEMI%
15 MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   If DMONE$(MONEMI%) = "" Then GoTo 15
   '
   MONEMIS = DMONE$(MONEMI%)
   MONEMIS.ForeColor = RGB(0, 0, 255)
   If MONEMI% > 1 Then
      MONEMIS.ForeColor = RGB(255, 0, 0)
   End If
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   '
   COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
   TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
   '
   If MONEMI% > 1 Then
     TICAMBIS.Visible = True
     TICAMBIS.Enabled = False
   End If
   '
   FIN& = ULTREG&("!CARDETPE")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDETPE", U&)
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     PRECARGA = COEFCON * CVS(Mid$(X$, 39, 4))
     PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
     TOT# = CDbl(CAN * PREUNI)
     '
     Call REPLA(X$, MKS$(PRECARGA), 39, 4)
     Call REPLA(X$, MKS$(PRECARGA), 51, 4)
     Call REPLA(X$, MKD$(TOT#), 55, 8)
     Call REPLA(X$, MKS$(PRECARGA), 65, 4)
     Call REPLA(X$, MKS$(PREUNI), 73, 4)
     Call REPLA(X$, MKD$(TOT#), 77, 8)
     Call GRAREG("!CARDETPE", X$, U&)
   Next U&
   '
   Call CARLISPE
   Call CIERRARCH("!CARDETPE")
   '
End Sub

Private Sub MoPeCu_Click()
   Call OPPED07.Command46_Click
End Sub

Private Sub nueped_Click()
   Call Command1_Click
End Sub

Private Sub otr_Click()
   Call MENSERR(24, "No Disponible por Ahora")
End Sub

Private Sub PeCliEx_Click()
   OPPED07.Option2 = True
   Call OPPED07.Command62_Click
End Sub

Private Sub PeClPan_Click()
   OPPED07.Option1 = True
   Call OPPED07.Command62_Click
End Sub

Private Sub PeCoEx_Click()
   OPPED07.Option2 = True
   Call OPPED07.Command64_Click
End Sub

Private Sub PeCoPa_Click()
   OPPED07.Option1 = True
   Call OPPED07.Command64_Click
End Sub

Private Sub Picture3_DblClick()
   Call CARDETPEDI
End Sub

Private Sub RCtaCte_Click()
   Call Command4_Click
End Sub

Private Sub RevAnu_Click()
   Call OPPED07.Command43_Click
End Sub

Private Sub Sali_Click()
   Call Command2_Click
End Sub

Private Sub Text1_Change()
   Call BLANFORMU
End Sub

Private Sub Text1_DblClick()
  Call SELECHO("DEUDOR1")
  NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
  Text1.TEXT = TRIM$(Str$(NCLIE%))
  If NCLIE% > 0 Then
    If REGICLI&(NCLIE%) > 0 Then
      Call CARDACLI
      Call CARDETPEDI
    End If
  End If
End Sub

Private Sub Text1_GotFocus()
   Text1.TEXT = TRIM$(Text1.TEXT)
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.TEXT)
   '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If VALICLIE% > 0 Then
      Call BLANFORMU
      If XVALO(Text1.TEXT) > 0 Then
        If XVALO(Text1.TEXT) <= 32767 Then
          NC% = XVALO(Text1)
          If REGICLI&(NC%) > 0 Then
            Call CARDACLI
            Call CARDETPEDI
          End If
        End If
      End If
    End If
  End If
End Sub

Private Sub Text10_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   Text10.TEXT = TRIM$(Text10.TEXT)
   Text10.SelStart = 0
   Text10.SelLength = Len(Text10.TEXT)
End Sub

Private Sub Text11_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   If ULTREG&("PEDENCA") > 2 Then
      On Error Resume Next
      Text10.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text11_LostFocus()
   NUMEPRE$ = ""
   For KK% = 1 To Len(Text11.TEXT)
     CCX% = Asc(Mid$(Text11.TEXT, KK%, 1))
     If CCX% >= 48 Then
       If CCX% <= 57 Then
         NUMEPRE$ = NUMEPRE$ + Chr$(CCX%)
       End If
     End If
   Next KK%
   '
   NPP& = NUPEDIDO&
   If XVALO(NUMEPRE$) < NPP& Then
     Call MENSERR(24, "Número no Válido")
     Text11.TEXT = TRIM$(Str$(NPP&))
     Exit Sub
   End If
   '
   If NUMEPRE$ <> TRIM$(Text11.TEXT) Then
     Text11.TEXT = NUMEPRE$
     Call MENSERR(24, "Número de Presupuesto Ajustado")
   End If
   '
End Sub

Private Sub Text12_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text13_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text14_DblClick()
   Call Command23_Click
End Sub

Private Sub Text19_Change()
  '\\\OT-07-0040-RG-02/02/07///
  Label23.Caption = ""
  Label24.Caption = ""
  LENTREGA(2) = ""
  TRANACT% = 0
  '\\\OT-07-0040-RG-FIN///
  '\\\OT-6-422-FG-05/10/06
  NTRA = XVALO(Text19)
  If NTRA > 0 Then
    If NTRA < 32767 Then
      NTRAI% = NTRA
      LENTREGA(2) = DENOTRANS$(NTRAI%)
      If LENTREGA(2) <> "" Then
         TRANACT% = NTRAI%
      End If
      '\\\OT-07-0040-RG-02/02/07///
      Label23.Caption = DOMITRANS$(NTRAI%)
      Label24.Caption = LOCATRANS$(NTRAI%)
      '\\\OT-07-0040-RG-FIN///
    End If
  End If
  '\\\OT-6-422-FG-FIN
End Sub

Private Sub Text2_gotfocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text20_Change()
    '\\\OT-6-422-FG-05/10/06
    NSUC& = XVALO(Text20)
    Label27 = ""
    If NSUC& > 0 Then
      If NSUC& <= ULTREG&("!SUCENTRE") Then
        XX$ = REGLEIDO$("!SUCENTRE", NSUC&)
        Label27 = TRIM$(Mid$(XX$, 3, 62))
        LENTREGA(0) = TRIM$(Mid$(XX$, 65, 48))
        LENTREGA(1) = TRIM$(Mid$(XX$, 113, 48))
      End If
    End If
    '\\\OT-6-422-FG-FIN
End Sub

Private Sub TEXT21_DblClick(Index As Integer)
    '
    If VALICLIE% < 1 Then
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
       Exit Sub
    End If
    '
    If Index = 0 Then
        ARCHEC$ = "LISTAS": LOCA% = 2
      ElseIf Index = 1 Then
        ARCHEC$ = "CONPAG": LOCA% = 2
      ElseIf Index = 2 Then
        ARCHEC$ = "VENDEDOR": LOCA% = 1
    End If
    ARGU$ = CBIN$(TRIM$(TEXT21(Index).TEXT), 1, LOCA%)
    Call SELECHO(ARCHEC$)
    VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
    If VDEVU$ <> "" Then
       TEXT21(Index).TEXT = VDEVU$
    End If
    '
End Sub

Private Sub Text21_GotFocus(Index As Integer)
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      Exit Sub
   End If
   TEXT21(Index).TEXT = TRIM$(TEXT21(Index).TEXT)
   TEXT21(Index).SelStart = 0
   TEXT21(Index).SelLength = Len(TEXT21(Index).TEXT)
End Sub

Private Sub Text3_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text4_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text5_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text21_Change(Index As Integer)
    '
    If XVALO(TEXT21(Index).TEXT) < 0 Or XVALO(TEXT21(Index).TEXT) > 255 Then
       TEXT21(Index).TEXT = ""
    End If
    '
    If Index = 0 Then
        If XVALO(TEXT21(Index).TEXT) < 1 Then
            Eco21(Index).TEXT = "CONCERTADOS FIJOS"
            ARCHIPRE$ = "MASTER"
            LIPRACTI% = 0
          Else
            Eco21(Index).TEXT = ECOARCH$("LISTAS", MKI$(XVALO(TEXT21(Index).TEXT)))
            ARCHIPRE$ = TRIM$(TEXT21(0).TEXT): LIPRACTI% = XVALO(TRIM$(ARCHIPRE$))
            While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
            ARCHIPRE$ = "LIPRE" + ARCHIPRE$
        End If
        If MONELISTA%(XVALO(TEXT21(Index))) = 2 Then
             MONEMIS = DEMONECO$(2)
             TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
           Else
             MONEMIS = DEMONECO$(1)
             TICAMBIS = FORMATNUM$(TICAMONE(1), "F9.4")
        End If
      ElseIf Index = 1 Then
        Eco21(Index).TEXT = ECOARCH$("CONPAG", MKI$(XVALO(TEXT21(Index).TEXT)))
      ElseIf Index = 2 Then
        Eco21(Index).TEXT = ECOARCH$("VENDEDOR", Chr$(XVALO(TEXT21(Index).TEXT)))
    End If
    '
End Sub

Private Sub Command21_Click(Index As Integer)
    If VALICLIE% < 1 Then
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
       Exit Sub
    End If
    '
    If Index = 0 Then
        ARCHEC$ = "LISTAS": LOCA% = 2
      ElseIf Index = 1 Then
        ARCHEC$ = "CONPAG": LOCA% = 2
      ElseIf Index = 2 Then
        ARCHEC$ = "VENDEDOR": LOCA% = 1
    End If
    ARGU$ = CBIN$(TRIM$(TEXT21(Index).TEXT), 1, LOCA%)
    Call SELECHO(ARCHEC$)
    VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
    If VDEVU$ <> "" Then
       TEXT21(Index).TEXT = VDEVU$
    End If
    '
End Sub
   '
Function VALICLIE%()
   VCC% = 0
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 32767 Then
       NCLIE% = XVALO(Text1.TEXT)
       If rasocli$(NCLIE%) <> "" Then
         VCC% = 1
       End If
     End If
   End If
   '
99 VALICLIE% = VCC%
   '
End Function

Sub BLANFORMU()
   Text2.TEXT = ""
   Text3.TEXT = ""
   Text4.TEXT = ""
   Text5.TEXT = ""
   Text7.TEXT = ""
   Text8.TEXT = ""
   Text10.TEXT = ""
   Text12.TEXT = ""
   Text13.TEXT = ""
   Text15 = ""
   Label18 = ""
   '
   '\\\OT-6-422-FG-05/10/06
   Text19 = ""
   Text20 = ""
   Call SETULTREG("!SUCENTRE", 0)
   Call FRESHECHO("!SUCENTRE")
   '\\\OT-6-422-FG-FIN
   '
   List1.Clear
   Call SETULTREG("!CARDETPE", 0)
   '
   TEXT21(0).TEXT = "": Eco21(0) = ""
   TEXT21(1).TEXT = "": Eco21(1) = ""
   Text44 = ""
   '
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   Text11.TEXT = TRIM$(Str$(NUPEDIDO&))
   '
End Sub
'
Sub CARDACLI()
   '
   Screen.MousePointer = 11
   If Control$("ADMIPED", "VERLICRE") = "SI" Then VERLICRE% = 1
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   NC% = XVALO(Text1.TEXT)
   If NC% > 0 Then
     '
     If InStr(EPAC$, "DOSIVA") > 0 Then
       PVCL% = PIVACLI%(NC%)
       If PVCL% <> 0 And PVCL% <> 5 Then
         PRVC$ = TRIM$(PROVCLI$(NC%))
         If PRVC$ = "" Or DEPROVI$(PRVC$) = "" Then
           Call COMUNI("Imposible Registrar O/Venta.\Falta Código de Provincia del Cliente.")
           Text1 = ""
           GoTo 99
         End If
       End If
     End If
     '
     '\\\OT-07-0010-RG-12/01/07///
     '\\\OT-6-422-FG-04/10/06
     ' CARGA LISTA DE SUCURSALES DEL CLIENTE ACTIVO
     Call FILTERFILE("SUCENTRE", "", 5, MKI$(NC%))
     '\\\OT-6-422-FG-FIN
     '\\\OT-07-0010-RG-FIN///
     '
     If rasocli$(NC%) <> "" Then
       Text2.TEXT = rasocli$(NC%)
       Text3.TEXT = DOMICLI$(NC%)
       Text4.TEXT = CPOSCLI$(NC%)
       Text5.TEXT = LOCACLI$(NC%)
       '
       '27/03/07 (OT-07-0133)
       LMCR# = LIMICRE(NC%)  'PONE EN 'LIMITE DE CREDITO' LA LEYENDA 'SUSPENDIDO'
        If LMCR# < 0 Then
           Text7 = "Suspendido"
        Else
           Text7.TEXT = CSTRING$(MKS$(LIMICRE(NC%)), 3, 13, 2, 2)
        End If
       '27/03/07 (FIN)
       '
       Text12.TEXT = TRIM$(POSIVAC$(NC%)) + " - CUIT " + CuitCli$(NC%)
       CREDISPO# = LIMICRE(NC%) - SALDCLI#(NC%) - CHEPEACRE#(NC%)
       Text8.ForeColor = RGB(0, 0, 0)
       If CREDISPO# < 0 Then Text8.ForeColor = RGB(255, 0, 0)
       Text8.TEXT = CSTRING$(MKD$(Abs(CREDISPO#)), 3, 13, 2, 2)
       '
       TEXT21(0).TEXT = Str$(LIPRCLI%(NC%))
       TEXT21(1).TEXT = Str$(CPAGCLI%(NC%))
       '
       For KKI% = 0 To 2
         LENTREGA(KKI%) = LUENTRECLI$(NC%, KKI%)
       Next KKI%
       '
       '23/03/07 (OT-07-0121)'NO PERMITE GENERAR UN PEDIDO A UN CLIENTE CON LA
       LMCR# = LIMICRE(NC%)  'CUENTA SUSPENDIDA
       If LMCR# < 0 Then
          Call COMUNI("Imposible Generar Pedido.\ \La Cuenta esta Suspendida")
          Text1 = ""
          Text1.SetFocus
          GoTo 99
       End If
       '23/03/07 (FIN)
       '
       If VERLICRE% = 1 Then
          SDEU# = SALDCLI#(NC%)
          CHPA# = CHEPEACRE#(NC%)
          If SDEU# + CHPA# > LMCR# Then
             Call COMUNI("Este Pedido no Podrá Facturarse.\  \Límite de Crédito Asignado ya está Superado.\  \Informe al Supervisor de Ventas.")
          End If
       End If
       '
     End If
   End If
99 Screen.MousePointer = 1
   '
End Sub
   '
Function NUPEDIDO&()
   '
   NPP& = 1
   Call ABREPEDCLI

   NPXX& = 1 + XVALO(VALOBADA("PEDENCA", "MAX(NROPED)", "NROPED > 0", "NOMESS"))
   If NPXX& > NPP& Then NPP& = NPXX&
   NUPEDIDO& = NPP&
   '
End Function
'
Private Sub Text6_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text7_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text8_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text9_GotFocus()
   On Error Resume Next
   Command9.SetFocus
   On Error GoTo 0
End Sub

Private Sub TICAMBIS_Change()
   If TICAMBIS.Enabled = True Then
     VH(MONEMI%) = XVALO(TICAMBIS)
   End If
End Sub

Private Sub TICAMBIS_LostFocus()
   TICAMBIS = FORMATNUM(XVALO(TICAMBIS), "F9.4")
End Sub

Sub PRIPEDIDO()
    '
    FORIPRE$ = TRIM$(Control$("", "FORPECLI"))
    If FORIPRE$ = "" Then Exit Sub
    '
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Pedido de Cliente.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    CANCELPRINT = 0
    If Control$("PEDCLIEN", "SUPRINT") = "SI" Then
       CANCELPRINT% = 1
    End If
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DOSIVAC") > 0 Then
      CANCELPRINT% = 0
      Screen.MousePointer = 1
      OPRIPEDI.Show 1
      If OPRIPEDI.Check1(0).Value <> 1 Then
        CANCELPRINT% = 1
      End If
    End If
    '
    Screen.MousePointer = 11
    NC% = XVALO(Text1)
    FECHDOC$ = Text9.TEXT
    NUMEDOC$ = TRIM$(Text11.TEXT)
    While Len(NUMEDOC$) < 6
      NUMEDOC$ = "0" + NUMEDOC$
    Wend
    '
    TIPODOC$ = "Pedido ABIERTO"
    If NC% < 1 Then
      TIPODOC$ = "Pedido Interno"
    End If
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CODTABU1$(4) = "PRE-LIST"
    CODTABU1$(5) = "PORDESCU"
    CODTABU1$(6) = "PRE-NETO"
    CODTABU1$(7) = "NET-ITEM"
    CODTABU1$(8) = "FECH-REC"
    CODTABU1$(9) = "ORD-ITEM"
    CACOLTAB1% = 9
    '
    DESTIDOC% = NC%
    CAITAB1% = 0
    CAITCRU% = 0
    '
    For U& = 1 To ULTREG&("!CARDETPE")
      XX$ = REGLEIDO$("!CARDETPE", U&)
      CIPRE% = CVI(Left$(XX$, 2))
      CAPEX$ = "Abier"   ' Str$(CVS(Mid$(XX$, 47, 4)))  es la cantidad en cero por pedido abierto
      '
      If CIPRE% > 0 Then
        If CIPRE% <= NUMAS% Then
          '
          PRELI$ = Str$(CVS(Mid$(XX$, 39, 4)))
          PREUX$ = Str$(CVS(Mid$(XX$, 51, 4)))
          DESCX$ = "0": On Error Resume Next
                        DESCX$ = Str$(100 * (XVALO(PRELI$) - XVALO(PREUX$)) / XVALO(PRELI$))
                        On Error GoTo 0
          ITOTX$ = Str$((XVALO(CAPEX$)) * (XVALO(PREUX$)))
          FEX = CVI(Mid$(XX$, 63, 2))
          FENTX$ = FECHATEX$(FEX)
          '
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = CODEXT$(CIPRE%)
          TETABU1$(2, CAITAB1%) = DESCRIT0$(CIPRE%)
          TETABU1$(3, CAITAB1%) = CAPEX$
          TETABU1$(4, CAITAB1%) = PRELI$
          TETABU1$(5, CAITAB1%) = DESCX$
          TETABU1$(6, CAITAB1%) = PREUX$
          TETABU1$(7, CAITAB1%) = ITOTX$
          TETABU1$(8, CAITAB1%) = FENTX$
          TETABU1$(9, CAITAB1%) = Str$(CAITAB1%)
          '
        End If
      End If
109 Next U&
    '
    SUMATOTA# = 0    ' XVALO(Total(0).TEXT)
    IDESCUEN# = 0    ' XVALO(Total(1).TEXT)
    '
    CODPARAM$(1) = "IMP-NETO": DOCPARAM$(1) = Str$(SUMATOTA#)
    CODPARAM$(2) = "VENDEDOR":     DOCPARAM$(2) = ""  'Eco21(2).TEXT
    CODPARAM$(3) = "COND-PAG":     DOCPARAM$(3) = Eco21(1).TEXT
    CODPARAM$(4) = "NRO-OCOM":     DOCPARAM$(4) = Text10.TEXT
    CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = LENTREGA(0).TEXT
    CODPARAM$(6) = "DESTINO": DOCPARAM$(6) = LENTREGA(1).TEXT
    CODPARAM$(7) = "TRANSPOR": DOCPARAM$(7) = LENTREGA(2).TEXT
    CODPARAM$(8) = "IMP-DESC": DOCPARAM$(8) = 0   '  Total(1).TEXT
    CODPARAM$(9) = "IMPNEDES": DOCPARAM$(9) = 0   '  Total(2).TEXT
    CODPARAM$(10) = "IMP-IVA": DOCPARAM$(10) = ""  ' TOTIVA(3).TEXT
    CODPARAM$(11) = "IMP-TOTA": DOCPARAM$(11) = ""   ' Total(3).TEXT
    CODPARAM$(12) = "POR-DESC": DOCPARAM$(12) = ""   ' TRIM$(PORDESCU.TEXT)
    'CODPARAM$(13) = "MONEDEMI": DOCPARAM$(13) = MONEMIS
    CODPARAM$(13) = "PERC-IVA": DOCPARAM$(13) = Text15
    CODPARAM$(14) = "FECH-VEN": DOCPARAM$(14) = Text14
    '
    '\\\OT-6-422-FG-05/10/06
    CODPARAM$(15) = "NRO-PASE": DOCPARAM$(15) = Label27
    '\\\OT-6-422-FG-FIN
    '
    '\\\OT-07-0007-RG-08/01/07///
    HORA$ = HORAENT$(NC%)
    CODPARAM$(16) = "Z-PLANO": DOCPARAM$(16) = TRIM$(HORA$)
    '\\\OT-07-0007-RG-FIN///
    '
    '\\\OT-07-0010-RG-12/01/07///
    AISUC% = XVALO(Text20)
    TELSUC$ = TELESUC$(NC%, AISUC%)
    CODPARAM$(17) = "REF-MAT1": DOCPARAM$(17) = TRIM$(TELSUC$)
    '\\\OT-07-0010-RG-FIN///
    '
    CAPARDOC% = 17
    '
    Call BLANARCH("!OBSERVOF")
    Call FRATEXTO(Text13.TEXT, 32)
    If CARETEX% >= 1 Then Call GRAREG("!OBSERVOF", RETEX$(1), 1)
    If CARETEX% >= 2 Then Call GRAREG("!OBSERVOF", RETEX$(2), 2)
    If CARETEX% >= 3 Then Call GRAREG("!OBSERVOF", RETEX$(3), 3)
    URETE& = CARETEX%
    Call SETULTREG("!OBSERVOF", URETE&)
    Call CIERRARCH("!OBSERVOF")
    '
    FORIPRE$ = TRIM$(Control$("", "FORPECLI"))
    If FORIPRE$ <> "" Then
      Call PRINTDOC("FORPECLI")   ' PEDIDO DE CLIENTE
    End If
    '
End Sub

Private Sub Total_GotFocus(Index As Integer)
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Sub GRAPEDIDO(OKX%)
   '
   OKX% = 0
   If ULTREG&("!CARDETPE") < 1 Then Exit Sub
   '
   NPP& = NUPEDIDO&
   If NPP& > XVALO(Text11.TEXT) Then
      Call COMUNI("FLEXOFT ha Re-Numerado el Pedido\como Pedido Número '" + TRIM$(Str$(NPP&)) + "'.\  \Verifique y Vuelva a Grabar.")
      Text11.TEXT = TRIM$(Str$(NPP&))
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLIE% = XVALO(Text1.TEXT)
   '
   If XVALO(TEXT21(0).TEXT) > 255 Then
      MERRO$ = "Lista de Precios no Válida"
      GoTo 98
   End If
   '
   If XVALO(TEXT21(1).TEXT) > 255 Then
      MERRO$ = "Condición de Pago no Válida"
      GoTo 98
   End If
   '
   HOII% = HOY(HOS$)
   FEMI% = FECHANUM(Text9.TEXT)
   If FEMI% < 1 Or FEMI% > HOII% Then
      MERRO$ = "Fecha de Emisión no Válida"
      GoTo 98
   End If
   '
   FEVENPE% = FECHANUM(Text14)
   If FEVENPE% < FEMI% Then
      MERRO$ = "Fecha de Vigencia no Válida - Restablecida."
      Text14 = Text9
      GoTo 98
   End If
   '
   If TRIM$(Text10) = "" Then
      MERRO$ = "Falta Orden de Compra Cliente"
      GoTo 98
   End If
   '
   NROPE& = XVALO(Text11.TEXT)
   REFE$ = " - " + TRIM$(Text10.TEXT)
   If REFE$ = " - " Then REFE$ = ""
   FEX = FEMI%
   REFE$ = REFE$ + " - " + FECHATEX$(FEX)
   LOMA% = 48 - Len(REFE$)
   REFE$ = Left$(TRIM$(rasocli$(NCLIE%)), LOMA%) + REFE$
   '
   Call ABREPEDCLI
   'jandy sql
   strSQL = "Select * from Pedenca where NroPed=" & NROPE&
   Dim PedEnca As ADODB.Recordset
   Set PedEnca = New ADODB.Recordset
25 On Error Resume Next
   PedEnca.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
 
   With PedEnca
      If Not (.EOF And .BOF) Then 'si lo encuentra
         Call CIERRARCH("PEDENCA")
         Call COMUNI("Número Ocupado")
         Exit Sub
      End If
      ' aca se inserta uno nuevo
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CODIEMPR = CodiEmp%
      On Error GoTo 0
      !NROPED = NROPE&
      !REFERENCIA = REFE$
      !TipoPed = "ABIERTO"
      !FECHEMIS = CVDAT(FEMI%)
      !NROCLIE = NCLIE%
      !RaSoClie = rasocli$(NCLIE%)
      !NROOCOM = TRIM$(Text10.TEXT)
      !DOMIENT = LENTREGA(0)
      !LOCAENT = LENTREGA(1)
      !TranspEnt = LENTREGA(2)
      !NTranspo = XVALO(Text19)
      !NSucent = XVALO(Text20)
      !Observa = Text13
      !FeSolEnt = CVDAT(HOII%)
      !finvigen = CVDAT(FEVENPE%)
         '\\\OT-05-0236-WAR-02/05/05///
         STATUPC% = (-1)
         If Left$(Control$("PEDCLI", "MODOAPRO"), 6) = "DIRECT" Then STATUPC% = 0
       !Status = STATUPC%              ' Status  de Aprobación
       '\\\OT-05-0236-FIN///
      !ListPre = XVALO(TEXT21(0))
      !CondPag = XVALO(TEXT21(1))
      !VENDED = 0
      !TDescApli = " "
     .Update
   End With
   PedEnca.Close
   Set PedEnca = Nothing
   '
   ANOTA$ = ""
   TTXYZ$ = Space$(76)
   HOII% = HOY(HOS$)
   Call REPLA(TTXYZ$, HOS$, 1, 8)
   Call REPLA(TTXYZ$, "Generación de Orden de Venta:", 11, 56)
   Call REPLA(TTXYZ$, USERTER$, 69, 24)
   ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
   '
   strSQL = "Select * from PedDeta where NroPed=" & NROPE&
   Dim PedDeta As ADODB.Recordset
   Set PedDeta = New ADODB.Recordset
26 On Error Resume Next
   PedDeta.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 26
   End If
   On Error GoTo 0
   '
   CAVER$ = TRIM$(Str$(NROPE&))
   While Len(CAVER$) > 6: CAVER$ = Mid$(CAVER$, 2): Wend
   While Len(CAVER$) < 6: CAVER$ = "0" + CAVER$: Wend
   IKITEM% = 0
   For U& = 1 To ULTREG&("!CARDETPE")
     X$ = REGLEIDO$("!CARDETPE", U&)
     IKITEM% = IKITEM% + 1
     IK$ = TRIM$(Str$(IKITEM%))
     CAVERI$ = CAVER$ + "-" + IK$
     '
     With PedDeta
       .AddNew   ' YA CONVERTIDO
       On Error Resume Next
       !CODIEMPR = CodiEmp%
       On Error GoTo 0
       !NROPED = NROPE&
       !TipoPed = "ABIERTO"
       !FECHEMIS = CVDAT(FEMI%)
       !NROCLIE = NCLIE%
       !RaSoClie = rasocli$(NCLIE%)
          NORCOMX$ = TRIM$(Mid$(X$, 93, 16))
          If NORCOMX$ = "" Then NORCOMX$ = TRIM$(Text10.TEXT)
       !NROOCOM = NORCOMX$
       !FeSolEnt = CVDAT(HOII%)
       !finvigen = CVDAT(FEVENPE%)
       !Status = -1 'pendiente de aprobación.
       !ListPre = XVALO(TEXT21(0).TEXT)
       !CondPag = XVALO(TEXT21(1).TEXT)
       !VENDED = 0
       !PorComVend = 0
       !NuOrItem = IKITEM%
         CIXI% = CVI(Left$(X$, 2))
       !CODIINT = CIXI%
       !CODIEXT = CODEXT$(CIXI%)
       !SubCod = Mid$(X$, 109, 8)
       !Descrip = DESCRIT0$(CIXI%)
       !TCOLOR = Mid$(X$, 3, 16)
       !TTALLE = Mid$(X$, 19, 8)
       !TLargo = CVS(Mid$(X$, 27, 4))
       !TAncho = CVS(Mid$(X$, 31, 4))
       !TEspesor = CVS(Mid$(X$, 35, 4))
       !VenRep = "V"     ' Venta o Reparacion
       !MoneEmis = MONEMI%
       !TIPOCAM = XVALO(TICAMBIS)
       !PreList = CVS(Mid$(X$, 39, 4))
       !CanAsig = 0   ' CVS(Mid$(X$, 43, 4))
       !CanPed = 0    ' indica que es pedido abierto
       !PreUnid = CVS(Mid$(X$, 51, 4))
       !ImpoTot = 0
       !CANTENT = 0
       !CanSaldo = 0
       !FeUltEnt = CVDAT(0)
       !NroRemi = ""
       !TPedItem = CAVERI$
       '
       TTXYZ$ = Space$(76)
       Call REPLA(TTXYZ$, "- " + CODEXT$(CIXI%), 11, 18)
       Call REPLA(TTXYZ$, DESCRIT0$(CIXI%), 29, 38)
       Call REPLA(TTXYZ$, CSTRING$(MKS$(CAN), 3, 10, 1, 2), 57, 10)
       ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
       '
       .Update
     End With
     '
   Next U&
   PedDeta.Close
   Set PedDeta = Nothing
   '
   ANOTA$ = ANOTA$ + Space$(10) + String$(56, "-") + Chr$(10) + Chr$(13)
   Call FRATEXTO(Text13.TEXT, 32)
   For KKI% = 1 To CARETEX%
     TTXX$ = Space$(92)
     Call REPLA(TTXX$, RETEX$(KKI%), 16, 32)
     If KKI% = 1 Then Call REPLA(TTXX$, "Observaciones:", 1, 15)
     ANOTA$ = ANOTA$ + TTXX$ + Chr$(10) + Chr$(13)
   Next KKI%
   ANOTA$ = ANOTA$ + String$(88, "=") + Chr$(10) + Chr$(13)
   '
   SQLX$ = " SELECT * FROM PEDENCA "
   SQLX$ = SQLX$ + "WHERE  NroPed = " & NROPE & ""
   Dim PEDDDENCA As ADODB.Recordset
   Set PEDDDENCA = New ADODB.Recordset
27 On Error Resume Next
   PEDDDENCA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 27
   End If
   On Error GoTo 0
   On Error Resume Next
   With PEDDDENCA
      If Err < 1 Then
        On Error GoTo 0
        If IsNull(!AnotaPed) = True Or !AnotaPed = "" Then
            !AnotaPed = ANOTA$
          Else
            VarAnotaped$ = !AnotaPed
            !AnotaPed = VarAnotaped$ + ANOTA$
         End If
          .Update
        End If
   End With
   PEDDDENCA.Close
   Set PEDDDENCA = Nothing
   '
   Call ABRECLIEN
   strSQL = "Update Deudor12 set Lentre0_Cli ='" & TRIM$(LENTREGA(0).TEXT) & "'"
   strSQL = strSQL & ", Lentre1_Cli ='" & TRIM$(LENTREGA(1).TEXT) & "'"
   strSQL = strSQL & ", Lentre2_Cli ='" & TRIM$(LENTREGA(2).TEXT) & "'"
   strSQL = strSQL & " where Nume_Cli = " & NCLIE%
    FLEXCONN.Execute (FILTSQL$(strSQL))
   '
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Call BAJALDISCO
   '
   OKX% = 1
   Screen.MousePointer = 1
   Exit Sub
   '
98 Screen.MousePointer = 1
   Call MENSERR(24, MERRO$)
   '
End Sub

Private Sub Total_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
     If Index = 0 Then
       On Error Resume Next
       PORDESCU.SetFocus
       On Error GoTo 0
     End If
   End If
End Sub

Sub CARDETPEDI()
   '
   HOII% = HOY(HOS$)
   '
   ATRI$ = "/NC/ARCHECO(1)=" + ARCHIPRE$
   ATRI$ = ATRI$ + "/TITUPAN=Detalle Pedido de Cliente"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top + 400))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left))
   '
10 PRF$ = "CARPEDAB"
   XXX% = VENTABU%(PRF$ + ATRI$)
   '
   If XXX% < 0 Then
      Call CIERRARCH("!CARDETPE")
      Call BLANARCH("!CARDETPE")
      Text1.TEXT = ""
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
20 If ULTREG&("!CARDETPE") < 1 Then
      Screen.MousePointer = 1
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   Screen.MousePointer = 11   ' Registrando Pedido ...
   FIN& = ULTREG&("!CARDETPE")
   VALIREPE% = 1 ' BANDERA DE VALIDACION DE PRECIOS
   If Control("ADMIPED", "VALIREPE") = "NO" Then VALIREPE% = 0

   If FIN& > 0 Then
     For U& = 1 To FIN&
       X$ = REGLEIDO$("!CARDETPE", U&)
       Ci% = CVI(Left$(X$, 2))
       If Ci% < 1 Or Ci% > NUMAS% Then
         Call MENSERR(24, "Código no Válido")
         Call REPLA(X$, MKI$(0), 1, 2)
         Call GRAREG("!CARDETPE", X$, U&)
         GoTo 10
       End If
       '
       If VALIREPE% > 0 Then
        For W& = 1 To U& - 1
          Y$ = REGLEIDO$("!CARDETPE", W&)
          If CVI(Left$(Y$, 2)) = Ci% Then
            Screen.MousePointer = 1
            TTXX$ = "Código " + TRIM$(CODEXT$(Ci%)) + "\Duplicado en el Pedido."
            If COMALTER%(TTXX$ + "\\Corregir\Continuar") <> 2 Then
              GoTo 10
            End If
          End If
        Next W&
       End If
       '
       PORDESCUIN = 0
       If ECOPRECIO% < 0 Then
         PRECARGA = PRELISTA#(ARCHIPRE$, Ci%)
         Call REPLA(X$, MKS$(PRECARGA), 39, 4)
         Call GRAREG("!CARDETPE", X$, U&)
       End If
       PRECARGA = CVS(Mid$(X$, 39, 4))
       PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
       '
       CADECI% = DECILISTA%(XVALO(Mid$(ARCHIPRE$, 6, 3)))
       VAMINI1 = 0.5 / (10 ^ CADECI%)
       If PRECARGA < VAMINI1 Or PRECARGA > 99999.99 Then
         If ECOPRECIO% >= 0 Then
           Call MENSERR(24, "Artículo sin Precio")
           Call REPLA(X$, MKS$(0), 39, 4)
           Call GRAREG("!CARDETPE", X$, U&)
         End If
       End If
       '
       Call REPLA(X$, MKS$(PRECARGA), 39, 4)
       Call REPLA(X$, MKS$(PRECARGA), 51, 4)
       Call REPLA(X$, MKS$(PRECARGA), 65, 4)
       Call REPLA(X$, MKS$(PREUNI), 73, 4)
       Call GRAREG("!CARDETPE", X$, U&)
       '
     Next U&
   End If
   '
   Call CARLISPE
   Call CIERRARCH("!CARDETPE")
   '
   List1.Refresh
   On Error Resume Next
   Command14.SetFocus
   On Error GoTo 0
   Screen.MousePointer = 1
   '
End Sub

Sub CARLISPE()
   List1.Clear
   FIN& = ULTREG&("!CARDETPE")
   '
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDETPE", U&)
     TTXX$ = Space$(96)
     Ci% = CVI(Left$(X$, 2))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     PRECARGA = CVS(Mid$(X$, 39, 4))
     PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
     '
     Call REPLA(TTXX$, CODEXT$(Ci%), 1, 15)
     Call REPLA(TTXX$, DESCRIT0$(Ci%), 18, 42)
     '
     PRELIX = PRELISTA#(ARCHIPRE$, Ci%)
     If PRELIX < PRECARGA Then PRELIX = PRECARGA
     If PRELIX < PREUNI Then PRELIX = PREUNI
     PORDESCUIN = 0: If PRELIX > 0.005 Then PORDESCUIN = 100 * (PRELIX - PREUNI) / PRELIX
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PRELIX), 3, 10, 2, 2), 60, 10)
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PORDESCUIN), 3, 5, 1, 2), 70, 5)
     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PREUNI), 3, 10, 4, 2), 76, 10)
     List1.AddItem TTXX$
   Next U&
'TTX = "123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 "
'TTX = "Codigo           Descripcion                                 Pr.Lista  % D    Pr.Unit"
End Sub
'
Sub PONEMARPREP()
   '
   If COMALTER%("FLEXOFT Debe Ahora Reorganizar\la Base de Datos de Pedidos\  \Todos los Usuarios Deben Salir\del Módulo de Pedidos de Clientes\\Cancelar\Re-Organizar Ahora") <> 2 Then
     Call FINAL("")
   End If
   '
   Call INDEXARCH("PEDDETA", 1)
   Call BLOQARCH("PEDDETA")
   Call BLOQARCH("PEDDETP")
   Call BLOQARCH("DESPAPED")
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("PEDDETA")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("PEDDETA", U&)
     NROPE& = CVS(Left$(XX$, 4))
     If NROPE& > 0 Then
       CAVER$ = TRIM$(Str$(NROPE&))
       While Len(CAVER$) > 6: CAVER$ = Mid$(CAVER$, 2): Wend
       While Len(CAVER$) < 6: CAVER$ = "0" + CAVER$: Wend
       '
       If Mid$(XX$, 113, 6) <> CAVER$ Then
         IKITEM% = 0
         RFF$ = RECFILT$("PEDDETA", 1, MKS$(NROPE&))
         For UL& = 1 To Len(RFF$) Step 4
           REPED& = CVS(Mid$(RFF$, UL&, 4))
           YY$ = REGLEIDO$("PEDDETA", REPED&)
           '
           IKITEM% = IKITEM% + 1
           IK$ = TRIM$(Str$(IKITEM%))
           CAVERI$ = CAVER$ + "-" + IK$
           Call REPLA(YY$, CAVERI$, 113, 10)
           Call GRAREG("PEDDETA", YY$, REPED&)
         Next UL&
       End If
     End If
   Next U&
   Call CIERRARCH("PEDDETA")
   '
   FIN& = ULTREG&("PEDDETP")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("PEDDETP", U&)
     NROPE& = CVS(Left$(XX$, 4))
     REPEDA& = CVS(Mid$(XX$, 109, 4))
     '
     If REPEDA& > 0 Then
       YY$ = REGLEIDO$("PEDDETA", REPEDA&)
       If CVS(Left$(YY$, 4)) = NROPE& Then
         CAVERI$ = Mid$(YY$, 113, 10)
         Call REPLA(XX$, CAVERI$, 113, 10)
         Call GRAREG("PEDDETP", XX$, U&)
       End If
     End If
   Next U&
   Call CIERRARCH("PEDDETP")
   '
   FIN& = ULTREG&("DESPAPED")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("DESPAPED", U&)
     NROPE& = CVS(Left$(XX$, 4))
     REPEDA& = CVS(Mid$(XX$, 109, 4))
     '
     If REPEDA& > 0 Then
       YY$ = REGLEIDO$("PEDDETA", REPEDA&)
       If CVS(Left$(YY$, 4)) = NROPE& Then
         CAVERI$ = Mid$(YY$, 113, 10)
         Call REPLA(XX$, CAVERI$, 113, 10)
         Call GRAREG("DESPAPED", XX$, U&)
       End If
     End If
   Next U&
   Call CIERRARCH("DESPAPED")
   '
   Call GRACONTROL("PEDCLIEN", "MARPREPA", "SI")
   Screen.MousePointer = 1
   Call COMUNI("Re-Organización Completada")
   '
End Sub
'NUEVO SUB
Private Sub MODOCON(MODO%)
    Call CIERRARCH("*.*")
    MODX% = MODO%
    If MODX% = 1 Then
        Call CONECRUN("LPREAC07", "Listas de Precios de Venta")
      ElseIf MODX% = 2 Then
        Call CARCONPAGOS
    End If
End Sub

Private Sub Timer2_Timer()
   TOTITEMS% = ULTREG&("!CARDETPE")
   Label18 = TRIM$(Str$(TOTITEMS%))
End Sub
