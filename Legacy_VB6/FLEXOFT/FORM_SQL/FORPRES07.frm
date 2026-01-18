VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FORPRES07 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PRESUPUESTOS A CLIENTES"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   12585
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FORPRES07.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   12585
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8520
      Left            =   0
      TabIndex        =   17
      Top             =   -50
      Width           =   12600
      Begin VB.PictureBox Picture999 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   1470
         ScaleHeight     =   300
         ScaleWidth      =   1140
         TabIndex        =   101
         Top             =   525
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2D3C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2565
         Left            =   0
         TabIndex        =   58
         Top             =   5475
         Width           =   11745
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Index           =   0
            Left            =   10155
            TabIndex        =   83
            Text            =   " "
            Top             =   105
            Width           =   1500
         End
         Begin VB.TextBox Eco21 
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   0
            Left            =   2100
            TabIndex        =   82
            Text            =   " "
            Top             =   100
            Width           =   4290
         End
         Begin VB.TextBox TEXT21 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   0
            Left            =   1365
            TabIndex        =   81
            Text            =   " "
            Top             =   100
            Width           =   435
         End
         Begin VB.CommandButton Command21 
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
            Height          =   280
            Index           =   0
            Left            =   1785
            TabIndex        =   80
            Top             =   100
            Width           =   180
         End
         Begin VB.TextBox TICAMBIS 
            Appearance      =   0  'Flat
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
            Left            =   7230
            MaxLength       =   14
            TabIndex        =   79
            Top             =   760
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E2D3C0&
            Caption         =   "CONDICIONES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1380
            Left            =   210
            TabIndex        =   70
            Top             =   1080
            Width           =   6195
            Begin VB.CommandButton Command30 
               Caption         =   "."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   5850
               TabIndex        =   100
               Top             =   1010
               Width           =   180
            End
            Begin VB.TextBox Total 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   2
               Left            =   315
               TabIndex        =   74
               Text            =   " "
               Top             =   1000
               Width           =   810
            End
            Begin VB.TextBox Total 
               Appearance      =   0  'Flat
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
               Index           =   4
               Left            =   1995
               TabIndex        =   73
               Text            =   " "
               Top             =   1000
               Width           =   3810
            End
            Begin VB.TextBox Total 
               Appearance      =   0  'Flat
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
               Index           =   6
               Left            =   1995
               TabIndex        =   72
               Text            =   " "
               Top             =   250
               Width           =   4035
            End
            Begin VB.TextBox Total 
               Appearance      =   0  'Flat
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
               Index           =   5
               Left            =   1995
               TabIndex        =   71
               Text            =   " "
               Top             =   630
               Width           =   4035
            End
            Begin VB.Label Ltota 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Plazo"
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
               Index           =   2
               Left            =   0
               TabIndex        =   78
               Top             =   760
               Width           =   1050
            End
            Begin VB.Label Ltota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Lugar"
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
               Index           =   4
               Left            =   1260
               TabIndex        =   77
               Top             =   1120
               Width           =   645
            End
            Begin VB.Label Ltota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Validez de Oferta"
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
               Index           =   6
               Left            =   0
               TabIndex        =   76
               Top             =   390
               Width           =   1905
            End
            Begin VB.Label Ltota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Garantia"
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
               Index           =   5
               Left            =   0
               TabIndex        =   75
               Top             =   765
               Width           =   1905
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E2D3C0&
            Caption         =   "OBSERVACIONES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1380
            Left            =   6720
            TabIndex        =   68
            Top             =   1065
            Width           =   4950
            Begin VB.TextBox Text17 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   945
               Left            =   240
               MaxLength       =   96
               MultiLine       =   -1  'True
               TabIndex        =   69
               Top             =   330
               Width           =   4470
            End
         End
         Begin VB.TextBox PORDESCU 
            Appearance      =   0  'Flat
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
            Left            =   8640
            MaxLength       =   14
            TabIndex        =   67
            Top             =   430
            Width           =   1065
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   1
            Left            =   10155
            TabIndex        =   66
            TabStop         =   0   'False
            Text            =   " "
            Top             =   430
            Width           =   1500
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   3
            Left            =   10155
            TabIndex        =   65
            TabStop         =   0   'False
            Text            =   " "
            Top             =   760
            Width           =   1500
         End
         Begin VB.TextBox Eco21 
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   2
            Left            =   2100
            TabIndex        =   64
            Text            =   " "
            Top             =   760
            Width           =   4290
         End
         Begin VB.TextBox TEXT21 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   2
            Left            =   1365
            TabIndex        =   63
            Text            =   " "
            Top             =   760
            Width           =   435
         End
         Begin VB.CommandButton Command21 
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
            Height          =   280
            Index           =   2
            Left            =   1785
            TabIndex        =   62
            Top             =   760
            Width           =   180
         End
         Begin VB.TextBox Eco21 
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   1
            Left            =   2100
            TabIndex        =   61
            Text            =   " "
            Top             =   430
            Width           =   4290
         End
         Begin VB.TextBox TEXT21 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   1
            Left            =   1365
            TabIndex        =   60
            Text            =   " "
            Top             =   420
            Width           =   435
         End
         Begin VB.CommandButton Command21 
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
            Height          =   280
            Index           =   1
            Left            =   1785
            TabIndex        =   59
            Top             =   430
            Width           =   180
         End
         Begin VB.Label TOTIVA 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6705
            TabIndex        =   93
            Top             =   0
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Descuento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   1
            Left            =   7080
            TabIndex        =   88
            Top             =   120
            Width           =   1905
         End
         Begin VB.Label PORIVA1 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6825
            TabIndex        =   99
            Top             =   1680
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label PORIVA0 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6825
            TabIndex        =   98
            Top             =   1365
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label IMPIVA105 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6720
            TabIndex        =   97
            Top             =   1680
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label IMPIVA21 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6720
            TabIndex        =   96
            Top             =   1365
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label14 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   7320
            TabIndex        =   95
            Top             =   750
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label TOTPER 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6705
            TabIndex        =   94
            Top             =   315
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "T-C"
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
            Left            =   6765
            TabIndex        =   92
            Top             =   840
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label MONEMIS 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   6720
            TabIndex        =   91
            ToolTipText     =   "Doble-Click para 'Dólares'"
            Top             =   345
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   9915
            TabIndex        =   90
            Top             =   540
            Width           =   225
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Neto"
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
            Index           =   3
            Left            =   9000
            TabIndex        =   89
            Top             =   870
            Width           =   1065
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Subtotal"
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
            Index           =   0
            Left            =   9120
            TabIndex        =   87
            Top             =   220
            Width           =   960
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vendedor"
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
            Index           =   2
            Left            =   -2835
            TabIndex        =   86
            Top             =   845
            Width           =   4110
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "C.Pago"
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
            Index           =   1
            Left            =   -2835
            TabIndex        =   85
            Top             =   510
            Width           =   4110
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "L.Precios"
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
            Index           =   0
            Left            =   -315
            TabIndex        =   84
            Top             =   185
            Width           =   1590
         End
      End
      Begin VB.CommandButton Command17 
         Caption         =   "!"
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
         Left            =   1245
         TabIndex        =   57
         Top             =   885
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
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
         Left            =   9960
         MaxLength       =   2
         TabIndex        =   55
         Top             =   525
         Width           =   450
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   10545
         TabIndex        =   53
         Top             =   525
         Width           =   660
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   210
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3675
         Left            =   240
         ScaleHeight     =   3645
         ScaleWidth      =   11250
         TabIndex        =   33
         Top             =   1680
         Visible         =   0   'False
         Width           =   11280
         Begin MSFlexGridLib.MSFlexGrid MSF 
            Height          =   4260
            Left            =   0
            TabIndex        =   112
            Top             =   0
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   7514
            _Version        =   393216
            Cols            =   4
            BackColor       =   16777215
            BackColorFixed  =   16501405
            BackColorBkg    =   16777215
            FillStyle       =   1
            SelectionMode   =   1
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Frame Frame21 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3600
            Left            =   10300
            TabIndex        =   110
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame20 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3600
            Left            =   9000
            TabIndex        =   109
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame19 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3600
            Left            =   7900
            TabIndex        =   108
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame18 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3600
            Left            =   7200
            TabIndex        =   107
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame17 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3600
            Left            =   6200
            TabIndex        =   106
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame16 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3600
            Left            =   2700
            TabIndex        =   105
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame15 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3600
            Left            =   840
            TabIndex        =   104
            Top             =   0
            Width           =   15
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
            Height          =   2985
            ItemData        =   "FORPRES07.frx":000C
            Left            =   -10
            List            =   "FORPRES07.frx":000E
            TabIndex        =   36
            Top             =   420
            Width           =   10575
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   490
            Left            =   -30
            ScaleHeight     =   435
            ScaleWidth      =   10515
            TabIndex        =   34
            ToolTipText     =   "Haga Doble-Click para Cargar Detalle."
            Top             =   -30
            Width           =   10570
            Begin VB.Label Label13 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "  Cant  Código          Descripción                   Lista  Desc  P.Unit     Importe"
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
               Height          =   465
               Left            =   0
               TabIndex        =   35
               ToolTipText     =   "Haga Doble-Click para Cargar Detalle."
               Top             =   120
               Width           =   10635
            End
         End
         Begin VB.Line Line6 
            X1              =   0
            X2              =   10600
            Y1              =   3620
            Y2              =   3620
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   10500
            Y1              =   3905
            Y2              =   3905
         End
         Begin VB.Line Line1 
            X1              =   10545
            X2              =   10545
            Y1              =   1440
            Y2              =   5360
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2D3C0&
         Caption         =   "ARTICULO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   700
         Left            =   570
         TabIndex        =   30
         Top             =   4620
         Width           =   8625
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
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
            TabIndex        =   15
            Text            =   " "
            Top             =   280
            Width           =   2325
         End
         Begin VB.CommandButton Command18 
            Caption         =   "."
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
            Left            =   2520
            TabIndex        =   16
            Top             =   280
            Width           =   180
         End
         Begin VB.Label Label8 
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
            Left            =   2940
            TabIndex        =   31
            Top             =   285
            Width           =   5520
         End
      End
      Begin VB.TextBox Text28 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Left            =   1830
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   2215
         Width           =   7890
      End
      Begin VB.CommandButton Command10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   420
         TabIndex        =   26
         Top             =   6855
         Width           =   2825
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   570
         ScaleHeight     =   405
         ScaleWidth      =   10575
         TabIndex        =   24
         Top             =   1785
         Width           =   10605
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción General - Especificaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   225
            Left            =   0
            TabIndex        =   25
            Top             =   105
            Width           =   10515
         End
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
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
         Text            =   " "
         Top             =   1280
         Width           =   3690
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   10815
         TabIndex        =   13
         Text            =   " "
         Top             =   890
         Width           =   870
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
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
         Left            =   10185
         TabIndex        =   14
         Text            =   " "
         Top             =   1280
         Width           =   1500
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
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
         Left            =   10530
         TabIndex        =   11
         Text            =   " "
         Top             =   105
         Width           =   960
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
         Height          =   315
         Left            =   11505
         TabIndex        =   12
         Top             =   105
         Width           =   195
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
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
         Left            =   5715
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   490
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   5715
         TabIndex        =   10
         Text            =   " "
         Top             =   1280
         Width           =   1695
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
         Height          =   300
         Left            =   7215
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   105
         Width           =   180
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
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
         Left            =   6795
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   105
         Width           =   435
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
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
         Width           =   2850
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
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
         Appearance      =   0  'Flat
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
         Width           =   3690
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
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
         Width           =   3690
      End
      Begin VB.CommandButton Command1 
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
         Left            =   1050
         TabIndex        =   1
         Top             =   885
         Width           =   180
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         TabIndex        =   0
         Text            =   " "
         Top             =   890
         Width           =   855
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00624E28&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   8520
         Left            =   11745
         ScaleHeight     =   8490
         ScaleWidth      =   1035
         TabIndex        =   37
         Top             =   0
         Width           =   1065
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
            Height          =   650
            Left            =   120
            Picture         =   "FORPRES07.frx":0010
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Opciones Varias- Consultas y Listados"""
            Top             =   3485
            Visible         =   0   'False
            Width           =   600
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
            Height          =   650
            Left            =   120
            Picture         =   "FORPRES07.frx":015A
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   750
            Width           =   600
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
            Height          =   500
            Left            =   120
            Picture         =   "FORPRES07.frx":0464
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Aprueba, Graba e Imprime Presupuesto"
            Top             =   6900
            Width           =   600
         End
         Begin VB.CommandButton Command15 
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
            Height          =   500
            Left            =   120
            Picture         =   "FORPRES07.frx":076E
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Vista Previa de Impresion"
            Top             =   6350
            Width           =   600
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
            Height          =   650
            Left            =   120
            Picture         =   "FORPRES07.frx":0A78
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   4200
            Visible         =   0   'False
            Width           =   600
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
            Height          =   650
            Left            =   120
            Picture         =   "FORPRES07.frx":0EBA
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Acceso a Base de Datos de Clientes"
            Top             =   2120
            Visible         =   0   'False
            Width           =   600
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
            Height          =   650
            Left            =   120
            Picture         =   "FORPRES07.frx":12FC
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Acceso a Maestro de Artículos"
            Top             =   1470
            Visible         =   0   'False
            Width           =   600
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
            Height          =   650
            Left            =   120
            Picture         =   "FORPRES07.frx":1606
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Generación Presupuesto por Copia"
            Top             =   5500
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.CommandButton Command5 
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
            Height          =   650
            Left            =   120
            Picture         =   "FORPRES07.frx":1910
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Condiciones Comerciales: Listas de Precios / Condiciones de Pago"
            Top             =   4850
            Visible         =   0   'False
            Width           =   600
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
            Height          =   650
            Left            =   120
            Picture         =   "FORPRES07.frx":1A5A
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   2835
            Visible         =   0   'False
            Width           =   600
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
            Height          =   650
            Left            =   120
            Picture         =   "FORPRES07.frx":1D64
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   100
            Width           =   600
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E2D3C0&
         Caption         =   "CANTIDAD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   700
         Index           =   1
         Left            =   9285
         TabIndex        =   50
         Top             =   4620
         Width           =   1905
         Begin VB.CommandButton Command16 
            Caption         =   "."
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
            Left            =   2520
            TabIndex        =   52
            Top             =   280
            Width           =   180
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   240
            TabIndex        =   51
            Text            =   " "
            Top             =   280
            Width           =   1500
         End
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   2625
         OleObjectBlob   =   "FORPRES07.frx":1EAE
         Top             =   0
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
         Height          =   225
         Index           =   0
         Left            =   0
         OleObjectBlob   =   "FORPRES07.frx":20E2
         TabIndex        =   102
         Top             =   0
         Visible         =   0   'False
         Width           =   1485
      End
      Begin MSFlexGridLib.MSFlexGrid GRID1 
         Height          =   765
         Left            =   1080
         TabIndex        =   103
         Top             =   3840
         Visible         =   0   'False
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   1349
         _Version        =   393216
         Rows            =   23
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   8433840
         BackColorSel    =   -2147483643
         ForeColorSel    =   12582912
         BackColorBkg    =   16777215
         Redraw          =   -1  'True
         FocusRect       =   2
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   111
         Top             =   0
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Left            =   11340
         TabIndex        =   56
         Top             =   525
         Width           =   330
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Id          -            -"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   9660
         TabIndex        =   54
         Top             =   630
         Width           =   330
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         TabIndex        =   49
         Top             =   1260
         Width           =   1200
      End
      Begin VB.Line Line5 
         X1              =   570
         X2              =   11175
         Y1              =   4545
         Y2              =   4545
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   12000
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "It. Unico"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   7965
         TabIndex        =   32
         ToolTipText     =   "Doble-Click para 'Lista Repuestos'"
         Top             =   90
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2325
         Left            =   9720
         TabIndex        =   29
         Top             =   2220
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Left            =   570
         TabIndex        =   28
         Top             =   2220
         Width           =   1245
      End
      Begin VB.Line Line7 
         X1              =   570
         X2              =   11170
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5715
         TabIndex        =   23
         Top             =   1000
         Width           =   1695
      End
      Begin VB.Line Line8 
         X1              =   10810
         X2              =   10810
         Y1              =   1785
         Y2              =   4540
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   8925
         TabIndex        =   22
         Top             =   990
         Width           =   1800
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8280
         TabIndex        =   21
         Top             =   1380
         Width           =   1800
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   9720
         TabIndex        =   20
         Top             =   210
         Width           =   750
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pos.IVA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6120
         TabIndex        =   19
         Top             =   200
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   210
         TabIndex        =   18
         Top             =   600
         Width           =   960
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Menu Archi 
      Caption         =   "Transacciones"
      Begin VB.Menu nuepre 
         Caption         =   "Nuevo Presupuesto"
      End
      Begin VB.Menu mnuPresupuestoPorCopia 
         Caption         =   "Nuevo Presupuesto por Copia"
      End
      Begin VB.Menu mnuMarcarNoVigente 
         Caption         =   "Marcar Presupuesto como No Vigente"
      End
      Begin VB.Menu php0 
         Caption         =   "-"
      End
      Begin VB.Menu Sali 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu ConLis 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu ReCuCo 
         Caption         =   "Resumen de Cuenta Corriente"
      End
      Begin VB.Menu php3 
         Caption         =   "-"
      End
      Begin VB.Menu LisRang 
         Caption         =   "Listados por Rango de Fecha"
      End
      Begin VB.Menu separa1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportesVarios 
         Caption         =   "Reportes Varios"
      End
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuración"
      Begin VB.Menu mnuAsociarReporte 
         Caption         =   "Vincular Reporte con Formulario"
      End
      Begin VB.Menu php4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVisPlantform 
         Caption         =   "Visualizar Plantilla de Formula de Items"
      End
      Begin VB.Menu separa2 
         Caption         =   "-"
      End
      Begin VB.Menu Configur 
         Caption         =   "Configuración  General"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
      Begin VB.Menu ConReimp 
         Caption         =   "Consulta y Re-impresión Presupuesto"
      End
      Begin VB.Menu mnuReimpresionProforma 
         Caption         =   "Consulta y Re-impresión Proforma"
      End
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuNuevoItem 
         Caption         =   "Nuevo Item "
      End
      Begin VB.Menu CondPag 
         Caption         =   "Condiciones de Pago"
      End
   End
   Begin VB.Menu mnuAccesosDirectos 
      Caption         =   "Accesos Directos"
      Begin VB.Menu MaCli 
         Caption         =   "Maestro de Clientes"
      End
      Begin VB.Menu MaAr 
         Caption         =   "Maestro de Articulos"
      End
      Begin VB.Menu LisPreVen 
         Caption         =   "Listas de Precios de Venta"
      End
   End
End
Attribute VB_Name = "FORPRES07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHIPRE$
Dim ALPORCOPIA%
Dim APLIVA%


Sub DARDEALTA()
    NUEVOCODI.TXTCLIENTE = Text1
    NUEVOCODI.Show 1
    If NUEVOCODI.TXTCODIGO = "" Then Exit Sub
       VDEF$ = String$(128, 0)
       '
       'PASA LOS DATOS DEL
       LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
       COD$ = TRIM$(NUEVOCODI.TXTCODIGO)
       CI_NEW% = CODINT_NEW
       Call REPLA(VDEF$, MKI$(CI_NEW%), 1, 2)
       Call REPLA(VDEF$, COD$, 77, LONCODI%)
       Call REPLA(VDEF$, "U.", 109, 2)
       Call REPLA(VDEF$, Chr$(1), 101, 1)
       '
       OKX% = 0
       
       Call ALTACODIGO(VDEF$, OKX%)
       
       If LONCODI% < 15 Or LONCODI% > 24 Then LONCODI% = 15
       If OKX% = 1 Then
          CODINUE$ = Mid$(VDEF$, 77, LONCODI%)
       End If
   
       'ACTUALIZO EL MASTER.
       Condi$ = " CODIGO <> '' AND STATUS > 0 ORDER BY CODIGO ASC"
       If Control$("", "IDSECUN") = "ALIAS" Then
           Call CARGALISEL("ECOMAST", "MASTER", "CODIGO/A24;DESCRIPCION/A36;NPLANO/A24", Condi$, "NOMESS")
       Else
           Call CARGALISEL("ECOMAST", "MASTER", "CODIGO/A24;DESCRIPCION/A50", Condi$, "NOMESS")
       End If
        '
        On Error Resume Next
        Unload NUEVOCODI
        On Error GoTo 0
        
        FRMENVAP.TXTOBS(31) = XVALO(Me.Text1)
        FRMENVAP.TXTRAZOCLI = Text2
        FRMENVAP.TXTCODIGO = COD$
        FRMENVAP.TXTPREPARO = USERNAME$
        FRMENVAP.TXTFECHPREPARO = FECHATEX$(HOY(HO$))
        FRMENVAP.Show 1

End Sub

Private Sub Command17_Click()
    '
    Call SELECHO("INDIPROS") ' PERMITE CARGA PRESUPUESTO A PARTIR DE PROSPECTOS
    NCLIE$ = TRIM$(VALACT1$("INDIPROS"))
    If NCLIE$ <> "" Then
      Text1.TEXT = "0"
      Label26.Caption = TRIM$(NCLIE$)
      NCLIE = XVALO(VALOBADA("DEUDOR12", "Nume_Cli", "Codi_Cli = '" & NCLIE$ & "'"))
      If NCLIE > 0 Then Text1 = NCLIE
      Call CARDACLI
      Text12.SetFocus
    End If
End Sub




Private Sub Command30_Click()
    Call SELECHO("PRELUENT")
    Total(4).TEXT = VALACT1$("PRELUENT")
End Sub


Private Sub GRID1_DblClick()
    If GRID1.Row > 0 Then
       On Error Resume Next
       List1.ListIndex = GRID1.Row - 1
       Call List1_DblClick
       On Error GoTo 0
    End If
End Sub



Private Sub mnuAsociarReporte_Click()
    Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
End Sub

Private Sub mnuMarcarNoVigente_Click()
    Call OPLIPRE07.Command3_Click
End Sub

Private Sub mnuNuevoItem_Click()
       If DERACCE%("ALTAITEM", "Alta de un Nuevo Item de Stock") < 2 Then
         Exit Sub
       End If
       Call DARDEALTA
End Sub

Private Sub mnuPresupuestoPorCopia_Click()
    Call Command7_Click
End Sub

Private Sub mnuReimpresionProforma_Click()
    Call OPLIPRE07.Command5_Click
End Sub

Private Sub mnuReportesVarios_Click()
    Call OPLIPRE07.AGenRep_Click
End Sub

Private Sub mnuVisPlantform_Click()
  FRMENVAP.Show 1
End Sub

Private Sub Picture2_Click()
    Call Label13_Click
End Sub

'\\\OT-06-0168-WAR-09/05/06///
Private Sub Text14_Change()
   '
   CANTI = XVALO(Text14.TEXT)
   SUBTOT = XVALO(Label14) * CANTI
   Total(0) = FORMATNUM$(SUBTOT, "F11.2")
   '
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
'
If KeyAscii = 13 Then
  CANTI = XVALO(Text14.TEXT)
  Text14.TEXT = FORMATNUM$(CANTI, "F11.1")
End If
'
End Sub

Private Sub Text14_LostFocus()
   '
   CANTI = XVALO(Text14.TEXT)
   Text14.TEXT = FORMATNUM$(CANTI, "F11.1")
   '
End Sub
'\\\OT-06-0168-WAR-FIN///
Private Sub ayud_Click()
   Call Command3_Click
End Sub

Private Sub CondPag_Click()
   Call CARCONPAGOS
End Sub

Private Sub Configur_Click()
   Call Command13_Click
End Sub

Private Sub ConReimp_Click()
   Call OPLIPRE07.Command1_Click
End Sub

Private Sub Label12_DblClick()
   Dim TRUE_FALSE As Boolean
   
   If TRIM$(UCase$(Label12)) = "IT. UNICO" Then
       Label12.ForeColor = RGB(200, 0, 0)
       Label12 = "Lista Repuestos"
       Label12.ToolTipText = "Doble-Click para 'Item Unico'"
       Picture1.Visible = True
       Picture2.Visible = True
       TRUE_FALSE = False
       MSF.Visible = True
     Else
       Label12.ForeColor = RGB(0, 80, 80)
       Label12 = "It. Unico"
       Label12.ToolTipText = "Doble-Click para 'Lista Repuestos'"
       Picture1.Visible = False
       TRUE_FALSE = True
       GRID1.Visible = False
       MSF.Visible = False
   End If
   
       'SE OBTIENE EL CONTENEDOR PICTURE PARA EL CASO DE SKINS PARA CAMBIAR LA PROPIEDAD VISIBLE
       Set CNT_L10 = Label10.Container
       If InStr(UCase(CNT_L10.Name), "PICTURE") > 0 Then CNT_L10.Visible = TRUE_FALSE
       Set CNT_L9 = Label9.Container
       If InStr(UCase(CNT_L9.Name), "PICTURE") > 0 Then CNT_L9.Visible = TRUE_FALSE

       If TRIM$(UCase$(Label12)) <> "IT. UNICO" Then Call REMPLAGRID(Me, GRID1, Picture1, Label13, List1)
End Sub


Private Sub Command1_Click()
   Call SELECHO("DEUDOR1")
   NCLIE = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE > 0 Then
      Text1.TEXT = TRIM$(Str$(NCLIE))
      Call CARDACLI
      Text12.SetFocus
   End If
   End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
  '\\\OT-05-0966-OD-28/10/08///
   If EXISTE%("SEGCOM07.EXE") > 0 Or EXISTE%("SEGCOM09.EXE") > 0 Then
    RTA% = ALTERNA%("Padron Maestro de Clientes\Seguimiento Comercial")
    If RTA% < 1 Or RTA% > 2 Then Command11.Enabled = True: Exit Sub
    If RTA% > 1 Then
       Call CONECRUN("SEGCOM07", "Seguimiento Comercial")
       Command11.Enabled = True
       Exit Sub
    End If
   End If
  '\\\OT-05-0966-OD-fin///
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG07/AMACLI"
   Call CONECRUN(ACONEC$, "Padron Maestro de Clientes")
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   OPLIPRE07.Show 1
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   OPPRES07.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   '
   
   ANEXPRESU.Show 1
   Refresh
   If ANEXPRESU.CONTROLABEL.Caption = "OK" Then
     Call GRAPRESU(OKX%)
     If OKX% > 0 Then
       CANCELPRINT% = 0
       '
       SENDBYMAIL$ = ""
       If ANEXPRESU.Option2.Value = True Then
          NCX1 = XVALO(Text1)
          'PERMITE ELEGIR MAIL DE ENVIA A PARTIR DE LA LISTA DE MAILS DE CONTACTOS
          Call CONTACTO_MAIL(Label26, NCX1) 'AGREGA A SENDBYMAIL LOS MAIL DE LOS CONTACTOS SELECCIONADOS
          If SENDBYMAIL$ = "" Then
            SENDBYMAIL$ = EMAILCLI$(NCX1)
          End If
          ' VALIDACION DE MAIL
          If SENDBYMAIL$ = "" Or InStr(SENDBYMAIL$, "@") < 1 Then 'TIENE DIRECCION VALIDA
             Call MENSERR(24, "Imposible Enviar por Mail.\  \Falta Direccion de Email o la Existente es Invalida.")
          End If
       End If
       '
       If ANEXPRESU.Option3.Value = True Then CANCELPRINT% = 1  ' IMPRESION FANTASMA
       '
       Call PRIPRESU(0) ' ES IMPRESION DEFINITIVA
       '
       'CARGA LA LINEA DE CORTE Y LAS OBSERVACIONES EN EL ANOTADOR DEL PROSPECT
       ANOTA$ = ANOTA$ + Space$(10) + String$(56, "-") + Chr$(10) + Chr$(13)
       Call FRATEXTO(ENTRAOBS.Text28.TEXT, 32)
       For KKI% = 1 To CARETEX%
          TTXX$ = Space$(92)
          Call REPLA(TTXX$, RETEX$(KKI%), 16, 32)
          If KKI% = 1 Then Call REPLA(TTXX$, "Observaciones:", 1, 15)
          ANOTA$ = ANOTA$ + Space$(10) + TTXX$ + Chr$(10) + Chr$(13)
       Next KKI%
       ANOTA$ = ANOTA$ + String$(88, "=")
       Call AGREGA_ANOTAPROSPECT(Label26.Caption, ANOTA$)  ' AGREGA LAS OBSERVACIONES Y LA LINEA DE FIN EN EL ANOTADOR DEL PROSPECT
       '
       Call BLANFORMU
       Text1.TEXT = ""
       Text1.SetFocus
     End If
   End If
99 Command14.Enabled = True
End Sub
Private Sub Command15_Click()
   If XVALO(Text1.TEXT) < 1 Then Exit Sub
   Command15.Enabled = False
   '\\\OT-08-1058-OD-19/11/08///'""MEDITEA""
   If DERACCE%("VISTAPRE", "Vista Previa de Impresion") < 2 Then ' DERECHO ACCESO A VISTA PREVIA DEL PRESUPUESTO
     GoTo 99
   End If
   '\\\OT-08-1058-OD-FIN///

   CANCELPRINT% = 1
   Screen.MousePointer = 11
   Call PRIPRESU(1)     ' ES VISTA PREVIA
   '
   TIDOCUM$ = "Presup"
   ' NUDOCU$ = " " + TRIM$(Text11.TEXT) + " -"
   NUDOCU$ = "(VP-" + TRIM$(Text11.TEXT) + ")"
   Screen.MousePointer = 11
   '
   'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
   DOCUNU& = DOCID&("PRESUPUESTO", NUDOCU$, "")

   If DOCUNU& > 0 Then
       Call MUESTRADOC(DOCUNU&)
   Else
       Screen.MousePointer = 1
       Call COMUNI("Documento no Encontrado")
   End If
   '
   GoTo 99

'''   For UI& = ULTREG&("PDOCLST") To 1 Step -1
'''      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'''      If InStr(XX$, TIDOCUM$) > 4 Then
'''        PPXX% = InStr(XX$, "Nro.")
'''        If PPXX% > 0 Then
'''          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
'''        End If
'''        If InStr(XX$, NUDOCU$) > 4 Then
'''          DOCUNU& = CVS(Left$(XX$, 4))
'''          GoTo 5
'''        End If
'''      End If
'''   Next UI&
'''   Screen.MousePointer = 1
'''   Call COMUNI("Documento no Encontrado")
'''   GoTo 99
'''   '
'''5 Call MUESTRADOC(DOCUNU&)
'''  GoTo 99
  '
  '***************************************************************************'
  'COMENTEADO LO DE ABAJO POR QUE TODO ESTO LO REALIZA DENTRO DEL MUESTRA DOC
  'APARTE ASI NO FUNCIONABA POR QUE AHORA GUARDA EN LA TABLA SQL PDOC_TBL
  '***************************************************************************'
''''''  For KKI% = 0 To 5
''''''      EDITFORM.Picture1(KKI%).Cls
''''''      EDITFORM.Picture1(KKI%).DrawWidth = 3
''''''      EDITFORM.Picture1(KKI%).DrawStyle = 0
''''''      EDITFORM.Picture1(KKI%).FillStyle = 1
''''''   Next KKI%
''''''   '
''''''   REDOCU$ = REGACT$("PDOCLST")
''''''   URE& = ULTREG&("PDOCSPL")
''''''   LI& = 0: LS& = URE&
''''''6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
''''''   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
''''''   DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
''''''   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
''''''   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
''''''   GoTo 8
''''''   '
''''''7  Screen.MousePointer = 1
''''''   Call MENSERR(24, "Documento no Registrado")
''''''   GoTo 99
''''''   '
''''''8  TPSP$ = ""
''''''   Screen.MousePointer = 11
''''''   While L& > 1
''''''      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
''''''      If DCC& < DOCUNU& Then GoTo 9
''''''      L& = L& - 1
''''''   Wend
''''''   '
''''''9  CAPAGINAS% = 0
''''''   For KKL& = L& To URE&
''''''      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
''''''      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
''''''      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
''''''        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
''''''      End If
''''''   Next KKL&
''''''   '
''''''95 SPOOLSTRING$ = TPSP$
''''''   For KKI% = 0 To 5
''''''      EDITFORM.Picture1(KKI%).Cls
''''''      EDITFORM.Picture1(KKI%).Top = 0
''''''      EDITFORM.Picture1(KKI%).Refresh
''''''      EDITFORM.Picture1(KKI%).Height = 7170
''''''      EDITFORM.Picture1(KKI%).Width = 11700
''''''      EDITFORM.Picture1(KKI%).DrawWidth = 3
''''''      EDITFORM.Picture1(KKI%).DrawStyle = 0
''''''      EDITFORM.Picture1(KKI%).FillStyle = 1
''''''   Next KKI%
''''''   ALTUPAGINA = 7170
''''''   TOPEPAGINA = 0
''''''   PAGINACTIVA% = 0
''''''   CAPAGINAS% = 0
''''''   '
''''''10 PPXX% = InStr(TPSP$, Chr$(255))
''''''   If PPXX% > 0 Then
''''''      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
''''''      TPSP$ = Mid$(TPSP$, PPXX% + 1)
''''''      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
''''''      '
''''''      If Len(COPRIAT$) > 0 Then
''''''        COPRI% = Asc(Left$(COPRIAT$, 1))
''''''        ATRIP$ = Mid$(COPRIAT$, 2)
''''''        Call SHOWCOMMAND(COPRI%, ATRIP$)
''''''      End If
''''''      GoTo 10
''''''      '
''''''   End If
''''''   '
''''''   UI& = UI& + 1
''''''   XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
''''''   If InStr(XX$, TIDOCUM$) > 4 Then
''''''      If InStr(XX$, NUDOCU$) > 4 Then
''''''        DOCUNU& = CVS(Left$(XX$, 4))
''''''        GoTo 7
''''''      End If
''''''   End If
''''''   '
''''''   Screen.MousePointer = 1
''''''   PAGINACTIVA% = 0
''''''   For KKI% = 0 To 5
''''''       EDITFORM.Picture1(KKI%).Visible = False
''''''   Next KKI%
''''''   EDITFORM.Picture1(PAGINACTIVA%).Visible = True
''''''   EDITFORM.Picture1(PAGINACTIVA%).Refresh
''''''   EDITFORM.VScroll1.Min = 0
''''''   EDITFORM.VScroll1.Value = 0
''''''   EDITFORM.VScroll1.Max = 1
''''''   If ALTUPAGINA > EDITFORM.Frame1.Height Then
''''''       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
''''''   End If
''''''   EDITFORM.Command3.Enabled = False
''''''   'EDITFORM.Command5.Enabled = False
''''''   EDITFORM.Show 1
''''''   'EDITFORM.Command5.Enabled = True
   '
99 Command15.Enabled = True
   '
End Sub

Private Sub Command18_Click()
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
   ARCHIPRE$ = "MASTER"
   LPREX% = XVALO(TEXT21(0))
   If LPREX% > 0 Then
     If LPREX% < 256 Then
       If ECOARCH$("LISTAS", Chr$(LPREX%)) <> "" Then
         ARCHIPRE$ = TRIM$(Str$(LPREX%))
         While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
         ARCHIPRE$ = "LIPRE" + ARCHIPRE$
         DELISECO$ = DENOLISTA$(LPREX%)
         ARCHIPRE$ = ARCHIPRE$ + "/" + DELISECO$
       End If
     End If
   End If
   Call SELECHO(ARCHIPRE$)
   Text13 = VALACT1$(ARCHIPRE$)
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("PRESUP04")
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
    RTA% = ALTERNA%("Listas de Precios de Venta\Condiciones de Pago")
    If RTA% = 1 Then
        Call CIERRARCH("*.*")
        '\\\OT-5-869-FG-28/12/05///
        Call CONECRUN("LPREAC07", "Listas de Precios de Venta")
        '\\\OT-5-869-FG-FIN///
      ElseIf RTA% = 2 Then
        Call CARCONPAGOS
    End If
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Text1.SetFocus
End Sub

Sub Command7_Click()
    Command7.Enabled = False
    If ALPORCOPIA% = 1 Then
      Call MENSERR(24, "Imposible Continuar.Generación por Copia en Proceso.")
      GoTo 99
    End If
    '
    Call OPLIPRE07.GENINDIPRES
    Call FRESHECHO("INDIPRES")
    '
3   Screen.MousePointer = 1
    '\\\OT-05-0758-WAR-18/11/05///
4    OBJ1$ = OBJPLA$("SELECLIPRE", VDEF$)
    If TRIM$(OBJ1$) <> "" Then
      OBJANT$ = OBJ1$
      NPX& = CVS(Mid$(OBJ1$, 1, 4))
      NCLIE = CVS(Mid$(OBJ1$, 5, 4))
      'VALIDACIONES
      If NPX& <= 0 Then
        Call COMUNI("Numero de Presupuesto no Válido")
        VDEF$ = OBJANT$
        GoTo 4
      End If
      If NCLIE <= 0 Then
        Call COMUNI("Numero de Cliente no Válido o Inexistente")
        VDEF$ = OBJANT$
        GoTo 4
      End If
      '
      Call PREPORCOP(NPX&, NCLIE)
      '
    End If
    '
99  Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CIERRARCH("*.*")
    ACONEC$ = "ARCHIG07/AMASTO"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
   Call SELECHO("SELFECHA")
   Text9.TEXT = VALACT1$("SELFECHA")
   FF% = FECHANUM(Text9.TEXT)
End Sub

Private Sub Eco21_GotFocus(Index As Integer)
   On Error Resume Next
   TEXT21(Index).SetFocus
   On Error GoTo 0
End Sub

Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
'    Show
    Refresh
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    If Left$(APLINVO$, 8) <> "ANACOSTO" Then
       Call CONAPLI(APLINVO$, TERMINA%) ' Busca la aplicacion a mostrar
       If TERMINA% > 0 Then Call FINAL("")
    End If
    '
    Call CONAPLI(APLINVO$, TERMINA%) ' Busca la aplicacion a mostrar
    If TERMINA% > 0 Then Call FINAL("")
    '
    
'15  EPAC$ = TRIM$(EMPREAC$)
'    If EPAC$ <> "" Then
'        Caption = Caption + "  -  " + EPAC$
'    End If
'    Call GRATITFOR(Caption)
    '
15  EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    '
     If ControlMenu%(Me.Name, Me) < 1 Then Call FINAL("")

    If InStr(UCase$(EMPREAC$), "ENVAPLAST") > 0 Then
        mnuNuevoItem.Visible = True
        mnuVisPlantform.Visible = True
    End If
    '
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$
    Text11.TEXT = TRIM$(Str$(NUPRESUP&))
    Call NUCORUB
    '
    ARCHIPRE$ = "MASTER"
    '
    MONEMI% = 1  ' PESOS
    MONEMIS = DEMONECO$(1)
    MONEMIS.Visible = True
    TICAMBIS = FORMATNUM$(TICAMONE(1), "F9.4")
    MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
    CONVERPRECIO% = 1
    TICAMBIS.Visible = True
    Label11.Visible = True
    '
    Label12.Visible = True
    '
    For U& = 1 To 2
       X$ = REGLEIDO$("GRUPIVA", U&)
       PRIVAX = CVS(Mid$(X$, 13, 4))
       If U& = 1 Then
           PORIVA0 = TRIM$(FORMATNUM$(PRIVAX, "F8.2"))
         Else
           PORIVA1 = TRIM$(FORMATNUM$(PRIVAX, "F8.2"))
       End If
    Next U&
    Call CIERRARCH("GRUPIVA")
    '
    If Left$(APLINVO$, 8) = "ANACOSTO" Then
       Call SELECHO("DEUDOR1")
       NC% = XVALO(VALACT1$("DEUDOR1"))
       If NC% < 1 Then Call FINAL("")
       Text1 = NC%
       Call CARDACLI
       Text14 = "1"
       Text13 = TRIM$(Mid$(APLINVO$, 9, 16))
       CIXI% = CODINT%(Text13)
       Text28 = ESPECIFIT$(CIXI%)
       Total(0) = TRIM$(Mid$(APLINVO$, 25))
       Exit Sub
    End If
    '
    If (EXISTE%("SEGCOM07.EXE") > 0 Or EXISTE%("SEGCOM09.EXE") > 0) And TRIM$(APLINVO$) <> "" Then ' LEE EL NUMERO DE PROSPECTO Y LO CARGA EN PANTALLA
       CODICL$ = TRIM$(APLINVO$)
       Call ABRECLIEN
       NC% = XVALO(VALOBADA("DEUDOR12", "Nume_Cli", "Codi_Cli = '" & CODICL$ & "'"))
       Text1 = NC%: Label26 = CODICL$
       Call CARDACLI
    End If
    '
    Campos$ = "Cant./F8.1;Código/A16;Descripción/A24;Lista/F12.2;Desc/F7.2;P.Unit/F12.2;Importe/F12.2"
    Call CARGA_ENCABEZADO(Me.MSF, Campos$, Me)
    
    If EXISTE%("SEGCOM07.EXE") > 0 Or EXISTE%("SEGCOM09.EXE") > 0 Then Command17.Visible = True
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    
    Call APERBASDAT("PEDCLI")
    Screen.MousePointer = 11
    Condi$ = "abs(preunit * cantit - imtone)>0.05"
    QTR& = CantRegistros("PRESUVEN", Condi$)
    If QTR& > 0 Then
       Dim PRESEL As ADODB.Recordset
       Set PRESEL = New ADODB.Recordset
       Set PRESEL = FLEXCONN.Execute(FILTSQL$("SELECT * FROM PRESUVEN WHERE " + Condi$))
       With PRESEL
          Do While Not .EOF
            NPRES& = XVALO(!NroPres)
            NORIX% = XVALO(!NuOrdIt)
            ITOTA$ = FORMATNUM$(XVALO(!PREUNIT) * XVALO(!Cantit), "F16.6")
            Call ACTUREGISTRO("PRESUVEN", "IMTONE", "NroPres = " & NPRES& & " AND NuOrdIt = " & NORIX%, ITOTA$, RAFE&, "NOMESS")
          .MoveNext
          Loop
       End With
       PRESEL.Close
       Set PRESEL = Nothing
    End If
    Screen.MousePointer = 1
    '
End Sub
'

Private Sub Form_Resize()
   If WindowState = 1 Then
      Call CIERRARCH("*.*")
      Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Openforms = DoEvents
   Wend
End Sub

Private Sub Label13_Click()
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
   Call CARDETPRES
End Sub

Private Sub Label26_DblClick()
   Call SELECHO("INDICLIE")
   CPROX$ = TRIM$(VALACT1$("INDICLIE"))
   If CPROX$ <> "" Then
     NCXI = NUMECLI(CPROX$)
     If NCXI > 0 Then
       If REGICLI&(NCXI) > 0 Then
         Label26 = CPROX$
         Text1.TEXT = TRIM$(Str$(NCXI))
         Call CARDACLI
       End If
     End If
   End If
End Sub

Private Sub Label8_Click()
   Text13.SetFocus
End Sub

Private Sub LisPreVen_Click()
   Call CIERRARCH("*.*")
   Call CONECRUN("LPREAC07", "Listas de Precios de Venta")
End Sub

Private Sub LisRang_Click()
   Call OPLIPRE07.Command6_Click
End Sub

Private Sub List1_DblClick()
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
   Call CARDETPRES
End Sub

Private Sub MSF_DblClick()
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
   Call CARDETPRES
End Sub

Private Sub MaAr_Click()
   Call Command8_Click
End Sub

Private Sub MaCli_Click()
   Call Command11_Click
End Sub

Private Sub MONEMIS_Change()
Exit Sub
   MONEMA% = MONEMI%
   If Left$(MONEMIS, 1) = "D" Then
        MONEMIS.ForeColor = RGB(255, 0, 0)
        MONEMIS.ToolTipText = "Doble-Click para 'Pesos'"
        Label11.Visible = True
        TICAMBIS.Enabled = True
        TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
        MONEMI% = 2
      Else
        MONEMIS.ForeColor = RGB(0, 0, 255)
        MONEMIS.ToolTipText = "Doble-Click para 'Dólares'"
        Label11.Visible = False
        TICAMBIS.Enabled = False
        MONEMI% = 1
   End If '
   Total(0) = FORMATNUM$(XVALO(Total(0)) * TICAMONE(MONEMA%) / TICAMONE(MONEMI%), "F11.2")
   '
   If Picture1.Visible = True Then
      COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
      FIN& = ULTREG&("!CARDETPR")
      For U& = 1 To FIN&
        X$ = REGLEIDO$("!CARDETPR", U&)
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
        Call GRAREG("!CARDETPR", X$, U&)
      Next U&
      '
      Call CARLISPRES
      Call CIERRARCH("!CARDETPR")
      Call CALTOPRES
   End If
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
   '
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
   '
   If MONEMI% > 1 Then
     TICAMBIS.Visible = True
     TICAMBIS.Enabled = False
   End If
   '
   Total(0) = FORMATNUM$(XVALO(Total(0)) * TICAMONE(MONEMIA%) / TICAMONE(MONEMI%), "F11.2")
   '
   If Picture1.Visible = True Then
      COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
      FIN& = ULTREG&("!CARDETPR")
      For U& = 1 To FIN&
        X$ = REGLEIDO$("!CARDETPR", U&)
        CAN = CVS(Mid$(X$, 47, 4))
        PORDESCUIN = CVS(Mid$(X$, 69, 4))
        PRECARGA = COEFCON * CVS(Mid$(X$, 39, 4))
        If Abs(PORDESCUIN) > 0 Then
           PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
        Else
           PREUNI = PRECARGA
        End If
        TOT# = CDbl(CAN * PREUNI)
        '
        Call REPLA(X$, MKS$(PRECARGA), 39, 4)
        Call REPLA(X$, MKS$(PRECARGA), 51, 4)
        Call REPLA(X$, MKD$(TOT#), 55, 8)
        Call REPLA(X$, MKS$(PRECARGA), 65, 4)
        Call REPLA(X$, MKS$(PREUNI), 73, 4)
        Call REPLA(X$, MKD$(TOT#), 77, 8)
        Call GRAREG("!CARDETPR", X$, U&)
      Next U&
      '
      Call CARLISPRES
      Call CIERRARCH("!CARDETPR")
      Call CALTOPRES
   End If
   '
End Sub

Private Sub nuepre_Click()
   Call Command1_Click
End Sub

Private Sub PORDESCU_Change()
   DESC$ = PORDESCU.TEXT
   If TRIM$(DESC$) <> "" Then
     TTT1 = XVALO(Total(0).TEXT) * COEFDESCU(DESC$) / 100
   Else
     TTT1 = 0
   End If
   
   Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   TTT3 = XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT)
   'Total(3).TEXT = CSTRING$(MKS$(TTT3), 3, 11, 2, 2)
   Total(3).TEXT = FORMATNUM$(TTT3, "F14.2")
End Sub

Private Sub PORDESCU_GotFocus()
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
End Sub

Private Sub ReCuCo_Click()
   Call Command4_Click
End Sub

Private Sub Sali_Click()
Call Command2_Click
End Sub

Private Sub TEXT1_Change()
   Call BLANFORMU
End Sub

Private Sub Text1_DblClick()
  Call SELECHO("DEUDOR1")
  NCLIE = XVALO(TRIM$(VALACT1$("DEUDOR1")))
  If NCLIE > 0 Then
    Text1.TEXT = TRIM$(Str$(NCLIE))
    Call CARDACLI
    Text12.SetFocus
  End If
End Sub

Private Sub Text1_GotFocus()
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.TEXT)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If VALICLIE% > 0 Then
      Call CARDACLI
      Text12.SetFocus
    End If
  End If
End Sub

Private Sub Text10_GotFocus()
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
End Sub

Private Sub Text11_GotFocus()
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
   If ULTREG&("PRESUPU") > 2 Then
      Text10.SetFocus
   End If
End Sub

Private Sub Text11_LostFocus()
   NUMEPRE$ = ""
   For kk% = 1 To Len(Text11.TEXT)
     CCX% = Asc(Mid$(Text11.TEXT, kk%, 1))
     If CCX% >= 48 Then
       If CCX% <= 57 Then
         NUMEPRE$ = NUMEPRE$ + Chr$(CCX%)
       End If
     End If
   Next kk%
   '
   NPP& = NUPRESUP&
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
   If VALICLIE% < 1 Then Text1.SetFocus
End Sub

Private Sub Text13_Change()
   If CODINT%(Text13) > 0 Then
      CI% = CODINT%(Text13)
      Label8.Caption = DESCRIT0$(CI%)
      LPREX% = XVALO(TEXT21(0))
      If LPREX% = 0 Then Total(0).TEXT = "" '\\\OT-06-0168-WAR-09/05/06///
      If LPREX% > 0 Then
        If LPREX% < 256 Then
          If ECOARCH$("LISTAS", Chr$(LPREX%)) <> "" Then
            ARCHIPRE$ = TRIM$(Str$(LPREX%))
            While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
            ARCHIPRE$ = "LIPRE" + ARCHIPRE$
            PREUNI = PRELISTA(ARCHIPRE$, CI%)
            Total(0).TEXT = CSTRING$(MKS$(PREUNI), 3, 11, 2, 2)
            '\\\OT-06-0168-WAR-09/05/06///
            Label14.Caption = "" 'LIMPIA PRECIO UNITARIO ANTERIOR
            Label14.Caption = Total(0).TEXT ' GUARDA EL PRECIO UNITARIO
            Text14 = FORMATNUM$("1", "F11.1") ' ESTABLECE CANTIDAD A 1
            '\\\OT-06-0168-WAR-FIN///
          End If
        End If
      End If
      Call CARESPECI
    End If
End Sub

Private Sub Text13_DBLCLICK()
   ARCHIPRE$ = "MASTER"
   LPREX% = XVALO(TEXT21(0))
   If LPREX% > 0 Then
     If LPREX% < 256 Then
       If ECOARCH$("LISTAS", Chr$(LPREX%)) <> "" Then
         ARCHIPRE$ = TRIM$(Str$(LPREX%))
         While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
         ARCHIPRE$ = "LIPRE" + ARCHIPRE$
         DELISECO$ = DENOLISTA$(LPREX%)
         ARCHIPRE$ = ARCHIPRE$ + "/" + DELISECO$
       End If
     End If
   End If
   Call SELECHO(ARCHIPRE$)
   Text13 = VALACT1$(ARCHIPRE$)
End Sub

Private Sub Text13_GotFocus()
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
End Sub

Private Sub Text16_Change()
   Call NUCORUB
End Sub

Private Sub Text16_LostFocus()
  Text16 = UCase(Text16)
End Sub

Private Sub Text2_GotFocus()
   Text1.SetFocus
End Sub

Private Sub TEXT21_GotFocus(Index As Integer)
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
End Sub

Private Sub Text3_GotFocus()
   Text1.SetFocus
End Sub

Private Sub Text4_GotFocus()
   Text1.SetFocus
End Sub

Private Sub Text5_GotFocus()
   Text1.SetFocus
End Sub

Private Sub Text6_Change()
    PIVA% = XVALO(Text6.TEXT)
    If ECOARCH$("CAPOSIV", Chr$(PIVA%)) <> "" Then
        XX$ = REGACT$("CAPOSIV")
        Text7.TEXT = UCase$(Mid$(XX$, 33, 8))
        FACIVA = 1
        CATIVA$ = "A"
        VA% = 1
        If PIVA% = 0 Or PIVA% = 3 Or PIVA% = 4 Or PIVA% = 6 Then
           FACIVA = 1 + ALICU / 100
           CATIVAS$ = "B"
           VA% = 2
           '
          ElseIf PIVA% = 2 Then
           ALICU = ALI1 + ALI2
           RNI% = 1
           '
          ElseIf PIVA% = 5 Then
           ALICU = 0
           CATIVAS$ = " "
           VA% = 0
           If LETREXPO$ = "A" Then
              VA% = 1
              CATIVAS$ = "A"
           End If
        End If
      Else
        Text6.TEXT = "1"
    End If
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
          Else
            Eco21(Index).TEXT = ECOARCH$("LISTAS", MKI$(XVALO(TEXT21(Index).TEXT)))
            ARCHIPRE$ = TRIM$(TEXT21(Index))
            While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
            ARCHIPRE$ = "LIPRE" + ARCHIPRE$
        End If
        If MONELISTA%(XVALO(TEXT21(Index))) = 2 Then
             MONEMIS = DEMONECO$(2)
             MONEMI% = 2
             TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
           Else
             MONEMIS = DEMONECO$(1)
             MONEMI% = 1
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
    If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
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
    TEXT21(Index).TEXT = VALACT1$(ARCHEC$)

    '
End Sub
   '
Function VALICLIE%()
   VCC% = 0
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 999999 Then
       NCLIE = XVALO(Text1.TEXT)
       If REGICLI&(NCLIE) > 0 Then
         VCC% = 1
       End If
     End If
   End If
   If VCC% = 0 Then
      CODIXX$ = TRIM$(Label26)
      CODICL$ = SAFETEXT$(VALOBADA("BSEGCOM", "Codi_Cli", "Codi_Cli ='" & CODIXX$ & "'"))
      If CODICL$ <> "" Then VCC% = 1
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
   Text6.TEXT = ""
   Text7.TEXT = ""
   Text8.TEXT = ""
   Text10.TEXT = ""
   Text12.TEXT = ""
   Text13.TEXT = ""
   Text28.TEXT = ""
   Text17.TEXT = "" 'Blanquea el campo observaciones
   Label26 = ""
   Label8.Caption = ""
   TOTIVA = ""
   TOTPER = ""
   '
   Text14.TEXT = FORMATNUM$(1, "F11.1") '\\\OT-06-0168-WAR-09/05/06///
   Label14.Caption = "" '\\\OT-06-0168-WAR-09/05/06///
   Text15.TEXT = ""
   '
   TEXT21(0).TEXT = "": Eco21(0) = ""
   TEXT21(1).TEXT = "": Eco21(1) = ""
   TEXT21(2).TEXT = "": Eco21(2) = ""
   '
   PORDESCU.TEXT = ""
   For kk% = 0 To 6
      Total(kk%).TEXT = ""
   Next kk%
   '
   '\\\LIMPIA LA OBSERVACION QUE SE CARGO EN EL FRM ENTRAOBS--OT-08-0887-OD-24/09/08///
   ENTRAOBS.Text28 = ""
   '\\\OT-08-0887-OD-24/09/08///
   List1.Clear
   MSF.Rows = 1
   Call SETULTREG("!CARDETPR", 0)
   Label12 = "It. Unico"
   If Picture1.Visible = True Then
      Call REMPLAGRID(Me, GRID1, Picture1, Label13, List1)
   End If
   Picture1.Visible = False
   '
   Set CNT_L10 = Label10.Container
   If InStr(UCase(CNT_L10.Name), "PICTURE") > 0 Then CNT_L10.Visible = True
   Set CNT_L9 = Label9.Container
   If InStr(UCase(CNT_L9.Name), "PICTURE") > 0 Then CNT_L9.Visible = True
   '
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   Text11.TEXT = TRIM$(Str$(NUPRESUP&))
   Call NUCORUB
   ALPORCOPIA% = 0
   '
End Sub
'
Sub CARDACLI()
    '
    Screen.MousePointer = 11
    NC = XVALO(Text1.TEXT)
    If NC > 0 Then
      'If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
      If CantRegistros("deudor12", "nume_cli=" & NC) > 0 Then
        Text2.TEXT = rasocli$(NC)
        Text3.TEXT = DOMICLI$(NC)
        Text4.TEXT = CPOSCLI$(NC)
        Text5.TEXT = LOCACLI$(NC)
        Text6.TEXT = PIVACLI%(NC)
        Text7.TEXT = POSIVAC$(NC)
        Text8.TEXT = TEL1CLI$(NC)
        Label26 = CODICLI$(NC)
        Text12.TEXT = CONTACTCLI$(NC)
        '
        TEXT21(0).TEXT = Str$(LIPRCLI%(NC))
        TEXT21(1).TEXT = Str$(CPAGCLI%(NC))
        TEXT21(2).TEXT = Str$(VENDCLI%(NC))
        '
        If XVALO(TEXT21(2).TEXT) < 1 Then
           Eco21(2).TEXT = "GERENCIA"
        End If
        '
        PORDESCU.TEXT = DESCUCLI$(NC)
        '
        LMCR# = LIMICRE(NC)  'CUENTA SUSPENDIDA
        If LMCR# < 0 Then
          Call COMUNI("Imposible Generar Presupuesto.\ \La Cuenta esta Suspendida")
          Text1 = ""
          On Error Resume Next
          Text1.SetFocus
          On Error GoTo 0
          GoTo 99
        End If
      End If
    ElseIf NC = 0 Then ' CARGA DATOS DEL PROSPECTO
        NCLIE$ = TRIM$(Label26)
        Text2.TEXT = VALOBADA("BSEGCOM", "Raso_Cli", "Codi_Cli = '" & NCLIE$ & "'") 'SELECCIONA LA RAZON SOCIAL DEL PROSPECT
        Text3.TEXT = VALOBADA("BSEGCOM", "Domi_Cli", "Codi_Cli = '" & NCLIE$ & "'")
        Text4.TEXT = VALOBADA("BSEGCOM", "Cpos_Cli", "Codi_Cli = '" & NCLIE$ & "'")
        Text5.TEXT = VALOBADA("BSEGCOM", "Loca_Cli", "Codi_Cli = '" & NCLIE$ & "'")
        Text8.TEXT = VALOBADA("BSEGCOM", "Tele_Cli", "Codi_Cli = '" & NCLIE$ & "'")
        Text12.TEXT = VALOBADA("BSEGCOM", "Ctac_Cli", "Codi_Cli = '" & NCLIE$ & "'")
        TEXT21(2).TEXT = VALOBADA("BSEGCOM", "Vend_Cli", "Codi_Cli = '" & NCLIE$ & "'")
        TEXT21(0).TEXT = TRIM$(Control("AMACLI", "LISTAS"))
        TEXT21(1).TEXT = TRIM$(Control("AMACLI", "CONPAG"))
    End If
    Total(5).TEXT = Control$("", "GARANDEF")
99  Screen.MousePointer = 1
    '
End Sub
   '
Function NUPRESUP&()
   '
   NPP& = 1
   Call ABREPRESVEN
   '
   Dim MIRS As ADODB.Recordset
   strSQL = ("SELECT max(NroPres) as maximo FROM PresuVen")
   Set MIRS = FLEXCONN.Execute(FILTSQL$(strSQL))
   If Not IsNull(MIRS!maximo) Then
      NPP& = 1 + MIRS!maximo
   End If
   NUPRESUP& = NPP&
   MIRS.Close
   Set MIRS = Nothing
   '
End Function
'
Sub NUCORUB()
   '
   CORUB$ = Text16
   CODAN$ = Label16
   If TRIM$(CORUB$) = "" Then Exit Sub
   NPP& = 1
   Call ABREPRESVEN
   '
   SQLX$ = "SELECT NumeCRubro FROM PresuVen "
   SQLX$ = SQLX$ + "WHERE Cod_Rubro = '" & CORUB$ & "' "
   SQLX$ = SQLX$ + "AND Cod_Anio = '" & CODAN$ & "' "
   SQLX$ = SQLX$ + "ORDER BY NumeCRubro DESC;"
   'Set PRESUTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim PRESUTEMP As ADODB.Recordset
   Set PRESUTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

   On Error Resume Next
   With PRESUTEMP
     .MoveFirst
     If Err < 1 Then
       NPP& = 1 + !NumeCRubro
     End If
   End With
   On Error GoTo 0
   '
   Text15 = NPP&
   '
End Sub
'
Sub CARESPECI()
      '
    CODD$ = Text13.TEXT
    CI% = CODINT%(CODD$)
    If CI% < 1 Then Exit Sub
    '
    Text28.TEXT = ESPECIFIT$(CI%)
    '
End Sub
'


Private Sub Text6_GotFocus()
   Text1.SetFocus
End Sub

Private Sub Text7_GotFocus()
   Text1.SetFocus
End Sub

Private Sub Text8_GotFocus()
   'Text1.SetFocus
End Sub

Private Sub Text9_Change()
   Label16 = Mid$(Text9, 7)
End Sub

Private Sub Text9_GotFocus()
   Command9.SetFocus
End Sub

Private Sub TICAMBIS_Change()
   If TICAMBIS.Enabled = True Then
     VH(MONEMI%) = XVALO(TICAMBIS)
   End If
End Sub

Private Sub TICAMBIS_LostFocus()
   TICAMBIS = FORMATNUM(XVALO(TICAMBIS), "F9.4")
End Sub

Private Sub Timer1_Timer()
   '
   Call CALTOPRES
   '
End Sub

Private Sub Total_Change(Index As Integer)
'   If Index = 0 Then
'      DESC$ = PORDESCU.TEXT
'      TTT1 = XVALO(Total(0).TEXT) * COEFDESCU(DESC$) / 100
'      Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
'      TTT3 = XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT)
'      Total(3).TEXT = CSTRING$(MKS$(TTT3), 3, 11, 2, 2)
'    ElseIf Index = 3 Then
'      Call CALTOIVA
'   End If
   '\\\OT-06-0168-WAR-09/05/06///
   If Index = 0 Then
      DESC$ = PORDESCU.TEXT
      TTT1 = Val(Total(0).TEXT) * COEFDESCU(DESC$) / 100
      Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
      TTT3 = Val(Total(0).TEXT) - Val(Total(1).TEXT)
      Total(3).TEXT = CSTRING$(MKS$(TTT3), 3, 11, 2, 2)
      CANTI = XVALO(Text14.TEXT)
      If CANTI < 0.1 Then CANTI = 1
      SUBTOT = Val(Total(0))
        PREUNI = XVALO(Total(0)) / CANTI ' CALCULA EL PRECIO UNITARIO
        If XVALO(Total(0)) >= 0.1 Then
          Label14 = FORMATNUM$(PREUNI, "F12.2") ' GUARDA EL ULTIMO P.UNITARIO
        End If
    ElseIf Index = 3 Then
      Call CALTOIVA
   End If
   '\\\OT-06-0168-WAR-FIN///
End Sub
'

Sub PRIPRESU(VIPREVIA%)
    '
    Call LIMPIATETAB
    
    ITEMUNICO% = 0
    If Picture1.Visible = False Then ITEMUNICO% = 1
    
    FORIPRE$ = TRIM$(Control$("", "FORPRESU"))
    FPXX$ = "FORPRESU"
    If Picture1.Visible = True Then
       FORIPRE$ = TRIM$(Control$("", "FORPRERE"))
       FPXX$ = "FORPRERE"
    End If

    If FORIPRE$ = "" Then
       Call MENSERR(24, "Imposible Imprimir Presupuesto.\Falta Definir Formulario Electrónico.")
       Exit Sub
    End If
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Presupuesto.\No existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    'SE AGREGA ESTE CODIGO PARA QUE TOME LA IMPRESORA CONFIGURADA
    PORTUPRE$ = TRIM$(Control$(IDENTER$, "PORTPRES"))
    If PORTUPRE$ = "" Then PORTUPRE$ = Control$("*", "PORTPRES")
    
    'DESDE AQUI AGREGADO PARA VER SI CORRESPONDE APLICAR IVA O NO
    NCX = XVALO(Text1)
    APLIVA% = 1
    If NCX > 0 Then
     PVCLI% = PIVACLI%(NCX)
     If PVCLI% = 5 Or PVCLI% = 7 Then APLIVA% = 0
    End If
    'FIN AGREGADO PARA APLICAR IVA O NO

10  NCX = XVALO(Text1.TEXT)
    TIPODOC$ = "Presupuesto"
    DESTIDOCNUE = NCX
    NUMEDOC$ = TRIM$(Text11.TEXT)
      If VIPREVIA% = 1 Then NUMEDOC$ = "(VP-" + NUMEDOC$ + ")"
    FECHDOC$ = TRIM$(Text9.TEXT)
    '
    CODPARAM$(1) = "REF-MAT3": DOCPARAM$(1) = TRIM$(Text10.TEXT)
    CODPARAM$(2) = "REF-MAT2": DOCPARAM$(2) = TRIM$(Text12.TEXT)
    CODPARAM$(3) = "PROV-CLI": DOCPARAM$(3) = NPROCLI$(NCX)
    CODPARAM$(4) = "COD-MAT": DOCPARAM$(4) = TRIM$(Text13.TEXT)
    CODPARAM$(5) = "DES-MAT": DOCPARAM$(5) = TRIM$(Label8.Caption)
    CODPARAM$(6) = "IMP-NETO": DOCPARAM$(6) = Total(0).TEXT
    CODPARAM$(7) = "POR-DESC": DOCPARAM$(7) = TRIM$(PORDESCU.TEXT)
    CODPARAM$(8) = "IMP-DESC": DOCPARAM$(8) = Total(1).TEXT
    CODPARAM$(9) = "IMPNEDES": DOCPARAM$(9) = Total(3).TEXT
    CODPARAM$(10) = "FECH-LAN": DOCPARAM$(10) = TRIM$(Total(2).TEXT)
    CODPARAM$(11) = "COND-PAG": DOCPARAM$(11) = TRIM$(Eco21(1).TEXT)
    CODPARAM$(12) = "DESTINO": DOCPARAM$(12) = TRIM$(Total(4).TEXT)
    CODPARAM$(13) = "FECH-VEN": DOCPARAM$(13) = TRIM$(Total(6).TEXT)
    CODPARAM$(14) = "FECH-REC": DOCPARAM$(14) = TRIM$(Total(5).TEXT) '\\\OT-06-0269-WAR-29/06/06///
    CODPARAM$(15) = "VENDEDOR": DOCPARAM$(15) = TRIM$(Eco21(2).TEXT)
    CODPARAM$(16) = "EQUIPO1": DOCPARAM$(16) = TRIM$(Text8.TEXT)
    '\\\OT-5-869-28/12/05///
    'CODPARAM$(17) = "MONEDEMI": DOCPARAM$(17) = TRIM$(MONEMIS)
    '\\\OT-5-869-FIN///
    CODPARAM$(17) = "TICAMBIO": DOCPARAM$(17) = TRIM$(TICAMBIS)
    CODPARAM$(18) = "IMP-IVA": DOCPARAM$(18) = IMPIVA21
    CODPARAM$(19) = "PERC-IVA": DOCPARAM$(19) = TRIM$(TOTPER)
    '\\\OT-5-544-FG-26/07/05///
      TOTPRESU$ = TRIM$(FORMATNUM$(XVALO(Total(3).TEXT) + XVALO(TRIM$(TOTIVA)) + XVALO(TRIM$(TOTPER)), "F12.2"))
    CODPARAM$(20) = "IMP-TOTA": DOCPARAM$(20) = TRIM$(TOTPRESU$) ' TOTAL DEL PRESUPUESTO
    CODPARAM$(21) = "MEDIDA-1": DOCPARAM$(21) = EMAILCLI(NCX) ' E-MAIL
    '\\\OT-5-544-FG-FIN///
    '\\\OT-5-869-FG-28/12/05///
    CODPARAM$(22) = "NRO-LOTE": DOCPARAM$(22) = TRIM$(Total(5).TEXT)
    CODPARAM$(23) = "CAN-CONJ": DOCPARAM$(23) = TRIM$(Text14.TEXT) '\\\OT-06-0168-WAR-09/05/06///
    '
    CR1$ = TRIM$(Text16):   NCR1 = XVALO(Text15):   CA1$ = TRIM$(Label16)
    IDPRESU$ = ""
    If TRIM$(CR1$) <> "" And XVALO(NCR1) > 0 Then
      NCR2$ = TRIM$(Str$(NCR1)):      While Len(NCR2$) < 4: NCR2$ = "0" + NCR2$: Wend
      IDPRESU$ = CR1$ + "-" + NCR2$ + "-" + CA1$ + "   "
    End If
    CODPARAM$(24) = "NRO-PASE": DOCPARAM$(24) = IDPRESU$
    CODPARAM$(25) = "IMP-IVA1": DOCPARAM$(25) = IMPIVA105
    CODPARAM$(26) = "POR-IVA": DOCPARAM$(26) = PORIVA0
    CODPARAM$(27) = "POR-IVA1": DOCPARAM$(27) = PORIVA1
    '
    ' LO QUE SIGUE ES PARA QUE EN CASO DE QUE NO DISCRIMINA IVA'S JUNTE TODO COMO IVA 21 %
    If TRIM$(ATRIFORM$(FORIPRE$, "IMP-IVA1/FORMATO")) = "" Then
      DOCPARAM$(18) = FORMATNUM$(XVALO(IMPIVA21) + XVALO(IMPIVA105), "F12.2")
      DOCPARAM$(25) = ""
      PORCENTAIVA = PORIVA0
      If XVALO(Total(3)) > 0.5 Then
        PORCENTAIVA = 100 * (XVALO(IMPIVA21) + XVALO(IMPIVA105)) / XVALO(Total(3))
      End If
      If PORCENTAIVA >= 0.005 Then
          DOCPARAM$(26) = FORMATNUM$(PORCENTAIVA, "F5.2")
        Else
          DOCPARAM$(26) = " 0.00"
      End If
    End If
    '
    CODPARAM$(28) = "PRE-NETO": DOCPARAM$(28) = ""
    CODPARAM$(29) = "IMPO-CHE": DOCPARAM$(29) = ""
    '
    CANTI# = XVALO(Text14.TEXT)
    If XVALO(Text14) >= 0.5 Then
      DOCPARAM$(28) = FORMATNUM$(XVALO(Total(3)) / XVALO(Text14) * CANTI#, "F12.2")
      DOCPARAM$(29) = FORMATNUM$(XVALO(Total(3)) / XVALO(Text14), "F12.2")
    End If
    
    PRIVA$ = Control$("PEDCLIEN", "PRECOIVA")
    If PRIVA$ = "SI" Then
      CI% = CODINT%(TRIM$(Text13.TEXT))
      
      If APLIVA% <> 0 Then ' VALIDACION POR SI ES CLIENTE DE EXPORTACION
        PIVA# = 1 + PORCEIVA(CI%) / 100
      End If
      '
      IMPIVA$ = FORMATNUM$(XVALO(Total(3)) / XVALO(Text14) * PIVA# * CANTI#, "F12.2")
      '
      DOCPARAM$(28) = TRIM$(IMPIVA$)
      IMPIVAUNI$ = FORMATNUM$(XVALO(Total(3)) / XVALO(Text14) * PIVA#, "F12.2")
      DOCPARAM$(29) = TRIM$(IMPIVAUNI$) '(PRECIO UNITARIO CON IVA)
    End If
    '
    CODPARAM$(30) = "NET-ITEM": DOCPARAM$(30) = TRIM$(Label14) ' PRECIO UNITARIO (LABEL14 OCULTO)
    
    If Picture1.Visible = False Then
       CODPARAM$(31) = "UNIMED": DOCPARAM$(31) = Unimed(CODINT%(TRIM$(Text13.TEXT)))
    End If
    CAPARDOC% = 31
    '
    If NCX < 1 Then ' Se agregan estos parametros para los casos en que esta tomando la informacion del prospect y todavia no son clientes (No figuran en la base de datos de clientes)
      NCLIE$ = TRIM$(Label26)
      '
      CODPARAM$(31) = "RASO-CLI": DOCPARAM$(31) = TRIM$(Text2.TEXT)
      CODPARAM$(32) = "DOMI-CLI": DOCPARAM$(32) = TRIM$(Text3.TEXT)
      CODPARAM$(33) = "LOCA-CLI": DOCPARAM$(33) = TRIM$(Text4.TEXT) + "-" + TRIM$(Text5.TEXT)
      CODPARAM$(34) = "TELE-CLI": DOCPARAM$(34) = TRIM$(Text8.TEXT)
      CODPARAM$(35) = "FAX-CLI": DOCPARAM$(35) = VALOBADA("BSEGCOM", "Fax_Cli", "Codi_Cli = '" & NCLIE$ & "'")
      CODPARAM$(36) = "MAIL-CLI": DOCPARAM$(36) = VALOBADA("BSEGCOM", "Mail_Cli", "Codi_Cli = '" & NCLIE$ & "'")
      '
      CAPARDOC% = 36
      '
    End If
    CAPARDOC% = CAPARDOC% + 1
    CODPARAM$(CAPARDOC%) = "OBSERVA": DOCPARAM$(CAPARDOC%) = Text17
    
    '\\\OT-5-869-FG-FIN///
    '
    CODTABU1$(1) = "REF-MAT1"
    CACOLTAB1% = 1
    '
    Call FRATEXTO(Text28.TEXT, 62)
    CAITAB1% = 0
    '
    For YX& = 1 To CARETEX%
       CAITAB1% = CAITAB1% + 1
       TETABU1$(1, CAITAB1%) = "." + RETEX$(YX&)
    Next YX&
    '
    If Picture1.Visible = True Then
      CODPARAM$(4) = "": DOCPARAM$(4) = ""
      CODPARAM$(5) = "": DOCPARAM$(5) = ""
      CODPARAM$(23) = "" 'AGREGADO POR QUE SI NO QUEDA CON 1 POR DEFECTO COMO CANTIDAD
      DOCPARAM$(30) = "" 'AGREGADO POR QUE SI NO QUEDABA EL IMPORTE NETO EN LA PRIMERA LINEA DEL ENCABEZADO
      '
      CODTABU1$(1) = "CANTIDAD"
      CODTABU1$(2) = "COD-MAT"
      CODTABU1$(3) = "DES-MAT"
      CODTABU1$(4) = "PRE-LIST"
      CODTABU1$(5) = "NET-ITEM"
      CACOLTAB1% = 5
      '
      
      CAITAB1% = 0
      For U& = 1 To ULTREG&("!CARDETPR")
        X$ = REGLEIDO$("!CARDETPR", U&)
        CIIX% = CVI(Mid$(X$, 1, 2))
        CAN = CVS(Mid$(X$, 47, 4))

        INETO = CVS(Mid$(X$, 39, 4)) * (100 - CVS(Mid$(X$, 69, 4))) / 100
        If PRIVA$ = "SI" Then
           PIVA# = 1
           If APLIVA% <> 0 Then ' VALIDACION POR SI ES CLIENTE DE EXPORTACION
             'SE CAMBIO LA VARIABLE CI% POR CIIX% QUE ES LA QUE CORRESPONDE
             PIVA# = 1 + PORCEIVA(CIIX%) / 100
           End If

'          PIVA# = 1 + PORCEIVA(CIIX%) / 100  ' VERIFICA QUE TIENE SETEADO EN EL SETUP PARA VER SI IMPRIME CON IVA INCLUIDO
          INETO = INETO * PIVA#
        End If
        If CAN > 0.05 Then
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = Str$(CAN)
          TETABU1$(2, CAITAB1%) = CODEXT$(CIIX%)
          TETABU1$(3, CAITAB1%) = DESCRIT0$(CIIX%)
          TETABU1$(4, CAITAB1%) = FORMATNUM$(INETO, "F12.4")
          TETABU1$(5, CAITAB1%) = FORMATNUM$(INETO * CAN, "F12.2")
          '
          If Control("PRESUVEN", "INCLUSPE") = "SI" Then
             XX$ = ESPECIFIT$(CIIX%)
             Call PRINTEXT(XX$, 3, "DES-MAT")
          End If
          '
          Call CARLINSEP         'Carga linea de separacion entre codigo y codigo
          '
        End If
      Next U&
    End If
    '
    NOCONFIRM% = 1
    'SE AGREGA ESTE CODIGO PARA QUE TOME LA IMPRESORA CONFIGURADA
    If FPXX$ <> "" Then
       If PORTUPRE$ <> "" Then Call SETPRINPORT(PORTUPRE$)
       Call PRINTDOC(FPXX$)  ' PRESUPUESTO DE CLIENTE
       If FPXX$ <> "" Then Call SETPRINPORT("RESTORE")
    End If

    '
    If Picture1.Visible = False Then
        CIXI% = CODINT%(Text13.TEXT)
        DOCUORIG$ = "Presupuesto " + NUMEDOC$
        TIPODOC$ = "Anexo Presup."
        For KKI% = 0 To 3
          If ANEXPRESU.Check1(KKI%).Value > 0 Then
            Call PRIDOGRAF(CIXI%, KKI%)
          End If
        Next KKI%
        If EXISTE_ESPECIFIC_ENVAPLAST%(CIXI%) > 0 Then
           OPX% = COMALTER%("Impresión de Receta\\No Imprimir\Imprimir")
           If OPX% > 1 Then
              FILTX$ = TRIM$(CIXI%)
              Call STDFORM("RCTA-ENV", FILTX$, "PRINT")
           End If
        End If
      Else
        For U& = 1 To ULTREG&("!CARDETPR")
          X$ = REGLEIDO$("!CARDETPR", U&)
          CIXI% = CVI(Mid$(X$, 1, 2))
          DOCUORIG$ = "Presupuesto " + NUMEDOC$
          TIPODOC$ = "Anexo Presup."
          For KKI% = 0 To 3
            If ANEXPRESU.Check1(KKI%).Value > 0 Then
              Call PRIDOGRAF(CIXI%, KKI%)
            End If
          Next KKI%
        Next U&
    End If
    '
    Call LIMPIATETAB
    'FACTURA PROFORMA
    If ANEXPRESU.Option4.Value = True Then
        '
        FILTX$ = TRIM$(NUMEDOC$) & ";" & TRIM$(Str$(NCX))
        If ITEMUNICO% = 1 Then
          FRMITEMUNICO$ = Control("PRESIUNI", "PROFLREP")
          Call STDFORM(FRMITEMUNICO$, FILTX$)
        Else
           FRMLISTAREP$ = Control("PRESUPUE", "PROFLREP")
           Call STDFORM(FRMLISTAREP$, FILTX$)
        End If
    End If
    
    If CANCELPRINT% < 1 Then
      If COMALTER%(DOCUORIG$ + " Registrado y Emitido por Impresora\\Repetir Impresión\Continuar") = 1 Then
        GoTo 10
      End If
    End If
    '
    
End Sub

Private Sub Total_GotFocus(Index As Integer)
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
   If Index = 1 Or Index = 3 Then
     Total(0).SetFocus
   End If
   If Index = 0 Then
     CII% = CODINT%(Text13.TEXT)
     If CII% < 1 Or CII% > NUMAS% Then
       Text13.SetFocus
     End If
   End If
End Sub

Sub GRAPRESU(OKX%)
   '
   OKX% = 0
   Call BLOQARCH("PRESUPU")
   '
   NPP& = NUPRESUP&
   If NPP& > XVALO(Text11.TEXT) Then
      Call CIERRARCH("PRESUPU")
      Call COMUNI("Número Ocupado")
      Text11.TEXT = TRIM$(Str$(NPP&))
      Exit Sub
   End If
   '
   '\\\OT-06-0168-WAR-09/05/06///
   'validar no repeticion de codigo y numero de rubro
   Call ABREPRESVEN
   NCPR1& = XVALO(Text15.TEXT)
   NCO1$ = Text16.TEXT
   NCO2$ = Label16
   If NCPR1& > 0 And TRIM$(NCO1$) <> "" Then
     SQLX$ = SQLX$ + " NumeCRubro = " & NCPR1& & ""
     SQLX$ = SQLX$ + " AND Cod_Rubro = '" & NCO1$ & "'"
     SQLX$ = SQLX$ + " AND Cod_Anio = '" & NCO2$ & "'"
     'jandy sql
     Dim MIRS As ADODB.Recordset
     Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With MIRS
       '.FindFirst SQLX$
       'If .NoMatch = False Then
       If Not (.EOF And Not .BOF) Then 'si lo encuentra
         Call CIERRARCH("PRESUPU")
         Call COMUNI("Código Identificador Ocupado")
         Call NUCORUB
         Exit Sub
       End If
     End With
     MIRS.Close
     Set MIRS = Nothing
   End If
   '\\\OT-06-0168-WAR-FIN///
   '
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLIE = XVALO(Text1.TEXT)
   '
   If Picture1.Visible = False Then
     CIIX% = CODINT%(Text13.TEXT)
     If CIIX% < 1 Or CIIX% > NUMAS% Then
       MERRO$ = "Código de Artículo no Válido"
       GoTo 98
     End If
   End If
   '
   If XVALO(Total(0).TEXT) < 1 Then
      MERRO$ = "Falta(n) Precio(s) de Lista"
      GoTo 98
   End If
   '
   If XVALO(Total(3).TEXT) < 1 Then
      MERRO$ = "Importe no Válido"
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
   '
   Call REPLA(XX$, Chr$(MONEMI%), 68, 1)
   Call REPLA(XX$, MKS$(XVALO(Total(0).TEXT)), 71, 4)
   Call REPLA(XX$, TRIM$(PORDESCU.TEXT), 75, 14)
      DESC$ = PORDESCU.TEXT
      TTT1 = COEFDESCU(DESC$)
   Call REPLA(XX$, MKS$(TTT1), 89, 4)
   Call REPLA(XX$, MKS$(XVALO(Total(1).TEXT)), 93, 4)
   Call REPLA(XX$, MKS$(XVALO(Total(3).TEXT)), 97, 4)
   Call REPLA(XX$, MKS$(1), 101, 4)
   Call REPLA(XX$, MKD$(CDbl(1 * XVALO(Total(3).TEXT))), 105, 8)
   '
   NUOIT% = 0
   Call ABREPRESVEN
   'jandy sql
   strSQL = "Select * from Presuven"
   strSQL = strSQL + " where NroPres=" & XVALO(Text11.TEXT)
   Dim PresuVen As ADODB.Recordset
   Set PresuVen = New ADODB.Recordset
   PresuVen.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   On Error Resume Next
   With PresuVen
   If Not (.EOF And .BOF) Then
      Call CIERRARCH("PRESUPU")
      Call COMUNI("Número Ocupado")
      Text11.TEXT = TRIM$(Str$(NPP&))
      Exit Sub
   End If
     
     CODICL$ = TRIM$(Label26) ' CARGA EL NUMERO DE PROSPECT
     Call REPLA(TTXYZ$, HOS$, 1, 8)
     Call REPLA(TTXYZ$, "Generación de Presupuesto:", 11, 56)
     Call REPLA(TTXYZ$, USERTER$, 69, 24)
     Call AGREGA_ANOTAPROSPECT(CODICL$, TTXYZ$) ' AGREGA EL ENCABEZADO EN EL ANOTADOR DEL PROSPECT
     'no lo encuentra entonces Inserta
     If Picture1.Visible = False Then
        '\\\OT-06-0168-WAR-09/05/06///
        CANTI0 = XVALO(Text14)
        If CANTI0 < 1 Then CANTI0 = 1
        '\\\OT-06-0168-WAR-FIN///
       .AddNew   ' YA CONVERTIDO
       On Error Resume Next
       !CodiEmpr = CodiEmp%
       On Error GoTo 0
        !NroPres = XVALO(Text11.TEXT)
        '\\\OT-06-0168-WAR-09/05/06///
         !NumeCRubro = XVALO(Text15.TEXT)
         !Cod_Rubro = Text16.TEXT
         !Cod_Anio = Label16
        '\\\OT-06-0168-WAR-FIN///
        !FECHEMIS = CVDAT(FEMI%)
        !NROCLIE = NCLIE
        !RaSoClie = Left$(rasocli$(NCLIE), 64)
        !Referen = Left$(Text10, 16)
        !AteSen = Left$(Text12, 48)
        !LisPre = XVALO(TEXT21(0))
        !ConPag = XVALO(TEXT21(1))
        !VENDED = XVALO(TEXT21(2))
        !MonedEmi = MONEMI%
        !NuOrdIt = NUOIT%
        !CODIINT = CIIX%
        !CODIEXT = Left$(CODEXT$(CIIX%), 16)
        !Descrip = Left$(DESCRIT0$(CIIX%), 64)
        !Especif = Text28
        !Observa = Text17
        !PreList = XVALO(Total(0)) / CANTI0 'XVALO(Total(0))'\\\OT-06-0168-WAR-09/05/06///
        !TDESCUEN = Left$(PORDESCU, 32)
           DESC$ = PORDESCU
        !PorcDesc = COEFDESCU(DESC$)
        '
        If !PorcDesc = 0 Then
            !ImpoDesc = 0
        Else
            !ImpoDesc = XVALO(Total(1)) / CANTI0 'XVALO(Total(1))'\\\OT-06-0168-WAR-09/05/06///
        End If
        '
        !PREUNIT = XVALO(Total(3)) / CANTI0  'XVALO(Total(3))'\\\OT-06-0168-WAR-09/05/06///
        !Cantit = CANTI0 '1'\\\OT-06-0168-WAR-09/05/06///
        !ImToNe = XVALO(Total(3)) '1 * XVALO(Total(3))'\\\OT-06-0168-WAR-09/05/06///
        !FechVent = CVDAT(0)
        !NroPedi = 0
        !NroFac_Cor = ""
        !NroFac_Lar = ""
        !Codi_Cli = CODICL$ ' GRABA CODIGO DE CUENTA DEL CLIENTE / PROSPECT
        !TTIVA21 = XVALO(IMPIVA21)
        !TTIVA105 = XVALO(IMPIVA105)
        !TOTPER = XVALO(TOTPER)

       .Update
        TTXYZ$ = Space$(76)
        Call REPLA(TTXYZ$, "- " + CODEXT$(CIIX%), 11, 18)
        Call REPLA(TTXYZ$, DESCRIT0$(CIIX%), 29, 38)
        Call REPLA(TTXYZ$, CSTRING$(MKS$(CANTI0), 3, 10, 1, 2), 57, 10)
        Call AGREGA_ANOTAPROSPECT(CODICL$, TTXYZ$) ' AGREGA EL DETALLE EN EL ANOTADOR DEL PROSPECT

      Else
       For U& = 1 To ULTREG&("!CARDETPR")
         X$ = REGLEIDO$("!CARDETPR", U&)
         CIIX% = CVI(Mid$(X$, 1, 2))
         CAN = CVS(Mid$(X$, 47, 4))
         INETO = CVS(Mid$(X$, 39, 4)) * (100 - CVS(Mid$(X$, 69, 4))) / 100
         If CAN > 0.05 Then
           .AddNew   ' YA CONVERTIDO
           On Error Resume Next
           !CodiEmpr = CodiEmp%
           On Error GoTo 0
             !NroPres = XVALO(Text11.TEXT)
             '\\\OT-06-0168-WAR-09/05/06///
              !NumeCRubro = XVALO(Text15.TEXT)
              !Cod_Rubro = Text16.TEXT
              !Cod_Anio = Label16
             '\\\OT-06-0168-WAR-FIN///
             !FECHEMIS = CVDAT(FEMI%)
             !NROCLIE = NCLIE
             !RaSoClie = Left$(rasocli$(NCLIE), 64)
             !Referen = Left$(Text10, 16)
             !AteSen = Left$(Text12, 48)
             !LisPre = XVALO(TEXT21(0))
             !ConPag = XVALO(TEXT21(1))
             !VENDED = XVALO(TEXT21(2))
             !MonedEmi = MONEMI%
             !NuOrdIt = NUOIT%
               NUOIT% = NUOIT% + 1
             !CODIINT = CIIX%
             !CODIEXT = Left$(CODEXT$(CIIX%), 16)
             !Descrip = Left$(DESCRIT0$(CIIX%), 64)
             !Especif = DESCRIT0$(CIIX%)
             !PreList = INETO
             !TDESCUEN = Left$(PORDESCU, 32)
                DESC$ = PORDESCU
                TTT1 = COEFDESCU(DESC$)
             !PorcDesc = TTT1
             !ImpoDesc = INETO * TTT1 / 100
             !PREUNIT = INETO * (1 - TTT1 / 100)
             !Cantit = CAN
             !ImToNe = CAN * INETO * (1 - TTT1 / 100)
             !FechVent = CVDAT(0)
             !NroPedi = 0
             !NroFac_Cor = ""
             !NroFac_Lar = ""
             !Codi_Cli = CODICL$
             !TTIVA21 = XVALO(IMPIVA21) ' 0
             !TTIVA105 = XVALO(IMPIVA105) ' 1280.24
             !TOTPER = XVALO(TOTPER)
             !Observa = Text17

            .Update
            TTXYZ$ = Space$(76)
            Call REPLA(TTXYZ$, "- " + CODEXT$(CIIX%), 11, 18)
            Call REPLA(TTXYZ$, DESCRIT0$(CIIX%), 29, 38)
            Call REPLA(TTXYZ$, CSTRING$(MKS$(CAN), 3, 10, 1, 2), 57, 10)
            Call AGREGA_ANOTAPROSPECT(CODICL$, TTXYZ$) ' AGREGA EL DETALLE EN EL ANOTADOR DEL PROSPECT
         End If
       Next U&
     End If
   End With
   PresuVen.Close
   Set PresuVen = Nothing
   '
   'VERIFICO SI ESTA INICIALIZADO EN EL SETUP CUAL ES LA PRIMER PRESUPUESTO CON LOS DATOS COMPLETOS PARA LA PROFORMA
   INICIO_PROFORMA& = XVALO(Control("PRESUPUE", "INIPROFO"))
   If INICIO_PROFORMA& < 1 Then
      Call GRACONTROL("PRESUPUE", "INIPROFO", SAFETEXT$(XVALO(Text11.TEXT)))
   End If
   '
   Call OPLIPRE07.ACTINDIPRES(XVALO(Text11))
   Call CIERRARCH("PRESUPU")
   '
'   CR1$ = TRIM$(Text16)
'   NCR1 = XVALO(Text15)
'   CA1$ = TRIM$(Label16)
'   '
'   If TRIM$(CR1$) <> "" And XVALO(NCR1) > 0 Then
'     NCR2$ = TRIM$(Str$(NCR1))
'     While Len(NCR2$) < 4: NCR2$ = "0" + NCR2$: Wend
'     CN1$ = CR1$ + "-" + NCR2$ + "-" + CA1$ + "   "
'   End If
'   YY2$ = TRIM$(AJUSTI$(CN1$ + TRIM$(RASOCLI$(NCLIE%)), 30)) + " - "
'   If Picture1.Visible = False Then
'       YY2$ = YY2$ + TRIM$(CODEXT$(CIIX%))
'     Else
'       YY2$ = YY2$ + "VARIOS"
'   End If
'   FEX = FEMI%
'   YY2$ = YY2$ + " - " + FECHATEX$(FEX)
'   Call REGAPP("INDIPRES", MKS$(XVALO(Text11)) + YY2$)
'   Call CIERRARCH("INDIPRES")
   '
   OKX% = 1
   Exit Sub
   '
98 Call CIERRARCH("PRESUPU")
   Call CIERRARCH("INDIPRES")
   Call MENSERR(24, MERRO$)
   '
End Sub

Private Sub Total_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
     If Index = 0 Then
       PORDESCU.SetFocus
     End If
   End If
End Sub

Private Sub Total_LostFocus(Index As Integer)
   If Index = 0 Then
     TTT1 = XVALO(Total(0).TEXT)
     Total(0) = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   End If
End Sub

Sub CARDETPRES()
   '
   
   HOII% = HOY(HOS$)
   '
   ATRI$ = "/NC/ARCHECO(1)=" + ARCHIPRE$
   ATRI$ = ATRI$ + "/TITUPAN=Detalle Presupuesto (Repuestos)"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top + 300))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
   '
   Timer1.Enabled = True
   
10 PRF$ = "CARPRESR"
   XXX% = VENTABU%(PRF$ + ATRI$)
   Call CALTOPRES
   Timer1.Enabled = False
   '
   If XXX% < 0 Then
      Call CIERRARCH("!CARDETPR")
      Call BLANARCH("!CARDETPR")
      Total(0).TEXT = ""
      Text1.TEXT = ""
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
20 If ULTREG&("!CARDETPR") < 1 Then
      Screen.MousePointer = 1
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   Screen.MousePointer = 11   ' Registrando Pedido ...
   FIN& = ULTREG&("!CARDETPR")
   If FIN& > 0 Then
     Y$ = REGLEIDO$("!CARDETPR", FIN&)
     CI% = CVI(Left$(Y$, 2))
     If CI% < 1 Or CI% > NUMAS% Then
       Call SETULTREG("!CARDETPR", FIN& - 1)
       Call CIERRARCH("!CARDETPR")
       GoTo 20
     End If
   End If
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   For U& = 1 To FIN&
     Y$ = REGLEIDO$("!CARDETPR", U&)
     CI% = CVI(Left$(Y$, 2))
     If CI% < 1 Or CI% > NUMAS% Then
       Call MENSERR(24, "Código no Válido")
       Call REPLA(Y$, MKI$(0), 1, 2)
       Call GRAREG("!CARDETPR", Y$, U&)
       GoTo 10
     End If
     '
     FENTRESOL% = CVI(Mid$(Y$, 63, 2))
     If FENTRESOL% < HOII% Then
        FENTRESOL% = HOII%
        Call REPLA(Y$, MKI$(FENTRESOL%), 63, 2)
        Call GRAREG("!CARDETPR", Y$, U&)
     End If
     '
     ' para DOSIVAC no verifica codigos repetidos
     If InStr(EPAC$, "DOSIVA") < 1 Then
       For W& = 1 To U& - 1
         Y$ = REGLEIDO$("!CARDETPR", W&)
         If CVI(Left$(Y$, 2)) = CI% Then
           If CVI(Mid$(Y$, 63, 2)) = FENTRESOL% Then
             Screen.MousePointer = 1
             TTXX$ = "Código " + TRIM$(CODEXT$(CI%)) + "\Duplicado en el Pedido."
             If COMALTER%(TTXX$ + "\\Corregir\Continuar") <> 2 Then
               GoTo 10
             End If
           End If
         End If
       Next W&
     End If
     '
     X$ = REGLEIDO$("!CARDETPR", U&)
     CI% = CVI(Left$(X$, 2))
     TTXX$ = Space$(96)
     FENTRESOL% = CVI(Mid$(X$, 63, 2))
     '
     CAN = CVS(Mid$(X$, 47, 4))
     If CAN > 999999 Then
       Call MENSERR(24, "Cantidad Excesiva - Max 999999")
       Call REPLA(X$, MKS$(0), 47, 4)
       Call GRAREG("!CARDETPR", X$, U&)
       GoTo 10
     End If
     If CAN < 0.5 Then
       Call MENSERR(24, "Falta Cantidad")
       Call REPLA(X$, MKS$(0), 47, 4)
       Call GRAREG("!CARDETPR", X$, U&)
       GoTo 10
     End If
     '
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     If PORDESCUIN < 0 Or PORDESCUIN > 99.9 Then
       Call MENSERR(24, "Descuento no Válido")
       Call REPLA(X$, MKS$(0), 69, 4)
       Call GRAREG("!CARDETPR", X$, U&)
       GoTo 10
     End If
     '
     If ECOPRECIO% < 0 Then
       PRECARGA = PRELISTA#(ARCHIPRE$, CI%)
       Call REPLA(X$, MKS$(PRECARGA), 39, 4)
       Call GRAREG("!CARDETPR", X$, U&)
     End If
     PRECARGA = CVS(Mid$(X$, 39, 4))
     PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
     TOT# = CDbl(CAN * PREUNI)
     '
     If PRECARGA < 0.005 Or PRECARGA > 9999999.99 Then
       If MODO% < 2 Then
         If ECOPRECIO% >= 0 Then
           Call MENSERR(24, "Artículo sin Precio")
           Call REPLA(X$, MKS$(0), 39, 4)
           Call GRAREG("!CARDETPR", X$, U&)
           MODO% = 2
           GoTo 10
         End If
       End If
     End If
     '
     Call REPLA(X$, MKS$(PRECARGA), 39, 4)
     Call REPLA(X$, MKS$(PRECARGA), 51, 4)
     Call REPLA(X$, MKD$(TOT#), 55, 8)
     Call REPLA(X$, MKS$(PRECARGA), 65, 4)
     Call REPLA(X$, MKS$(PREUNI), 73, 4)
     Call REPLA(X$, MKD$(TOT#), 77, 8)
     Call GRAREG("!CARDETPR", X$, U&)
     '
   Next U&
   '
   Call CARLISPRES
   Call CIERRARCH("!CARDETPR")
   Call CALTOPRES
   '
   List1.Refresh
   '
   Call REMPLAGRID(Me, GRID1, Picture1, Label13, List1)
   On Error Resume Next
   Command14.SetFocus
   On Error GoTo 0
   Screen.MousePointer = 1
   '
End Sub

Sub CARLISPRES()
   '
   List1.Clear
   MSF.Rows = 1
   FIN& = ULTREG&("!CARDETPR")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDETPR", U&)
     TTXX$ = Space$(96)
     CI% = CVI(Left$(X$, 2))
     FENTRESOL% = CVI(Mid$(X$, 63, 2))
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     PRECARGA = CVS(Mid$(X$, 39, 4))
     If Abs(PORDESCUIN) > 0 Then
       PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
     Else
       PREUNI = PRECARGA
     End If
     TOT# = CDbl(CAN * PREUNI)
     'Call REPLA(TTXX$, DESCRIT0$(CI%), 24, 28)
     PRELIX = PRELISTA#(ARCHIPRE$, CI%)
     If PRELIX < PRECARGA Then PRELIX = PRECARGA
     If PRELIX < PREUNI Then PRELIX = PREUNI
     PORDESCUIN = 0: If PRELIX > 0.005 Then PORDESCUIN = 100 * (PRELIX - PREUNI) / PRELIX
     '
     MSF.Rows = U& + 1
     MSF.TextMatrix(U&, 1) = CAN
     MSF.TextMatrix(U&, 2) = CODEXT$(CI%)
     MSF.TextMatrix(U&, 3) = DESCRIT0$(CI%)
     MSF.TextMatrix(U&, 4) = TRIM$(FORMATNUM$(PRELIX, "F12.2"))
     MSF.TextMatrix(U&, 5) = TRIM$(FORMATNUM$(PORDESCUIN, "F12.2"))
     MSF.TextMatrix(U&, 6) = TRIM$(FORMATNUM$(PREUNI, "F12.2"))
     MSF.TextMatrix(U&, 7) = TRIM$(FORMATNUM$(TOT#, "F14.2"))
     '
'     MSF.TextMatrix(U&, 1) = CAN
'     MSF.TextMatrix(U&, 2) = CODEXT$(CI%)
'     MSF.TextMatrix(U&, 3) = DESCRIT0$(CI%)
'     MSF.TextMatrix(U&, 4) = ROUND(PRELIX, 2)
'     MSF.TextMatrix(U&, 5) = ROUND(PORDESCUIN, 2)
'     MSF.TextMatrix(U&, 6) = ROUND(PREUNI, 2)
'     MSF.TextMatrix(U&, 7) = ROUND(TOT#, 2)

'     Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 1, 6)
'     Call REPLA(TTXX$, CODEXT$(CI%), 8, 15)
'     FEX = FENTRESOL%
'     Call REPLA(TTXX$, DESCRIT0$(CI%), 24, 28)
'     PRELIX = PRELISTA#(ARCHIPRE$, CI%)
'     If PRELIX < PRECARGA Then PRELIX = PRECARGA
'     If PRELIX < PREUNI Then PRELIX = PREUNI
'     PORDESCUIN = 0: If PRELIX > 0.005 Then PORDESCUIN = 100 * (PRELIX - PREUNI) / PRELIX
'     Call REPLA(TTXX$, CSTRING$(MKS$(PRELIX), 3, 8, 2, 2), 52, 8)
'     Call REPLA(TTXX$, CSTRING$(MKS$(PORDESCUIN), 3, 5, 1, 2), 61, 5)
'     Call REPLA(TTXX$, CSTRING$(MKS$(PREUNI), 3, 8, 2, 2), 66, 8)
'     Call REPLA(TTXX$, CSTRING$(MKD$(TOT#), 3, 10, 2, 2), 76, 10)
'     List1.AddItem TTXX$
     '
     
   Next U&
   '
   Call REMPLAGRID(Me, GRID1, Picture1, Label13, List1)
   '
End Sub
'
Sub CALTOPRES()
   '
   Static CLIA$, RNI%, APLIVA%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE = XVALO(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIE > 0 Then
       PVCLI% = PIVACLI%(NCLIE)
       If PVCLI% = 2 Then RNI% = 1
       If PVCLI% = 5 Or PVCLI% = 7 Then APLIVA% = 0
     End If
   End If
   '
   SUTOTPED# = 0: SUIVA# = 0: SUIVARNI# = 0: TOIVA# = 0: TOTOTA# = 0
   For U& = 1 To ULTREG&("!CARDETPR")
     X$ = REGLEIDO$("!CARDETPR", U&)
     CIXXI% = CVI(Left$(X$, 2))
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     If PORDESCUEIN < 0 Or PORDESCUEIN > 99 Then PORDESCUEIN = 0
     If CAN > 0.05 Then
       PRELI = CVS(Mid$(X$, 39, 4))
       'If PRELI < 0.01 Then
       '  PRELI = PRELISTA(ARCHIPRE$, CIXXI%)
       'End If
       If Abs(PORDESCUIN) > 0 Then
         PREUNI = PRELI * (100 - PORDESCUIN) / 100
       Else
         PREUNI = PRELI
       End If
       TOT# = CDbl(CAN * PREUNI)
       SUTOTPED# = SUTOTPED# + TOT#
       SUIVA# = SUIVA# + (Int(TOT# * PORCEIVA(CIXXI%) + 0.5)) / 100
       Call REPLA(X$, MKS$(PREUNI) + MKD$(TOT#), 51, 12)
       Call GRAREG("!CARDETPR", X$, U&)
     End If
   Next U&
   '
   Total(0).TEXT = CSTRING$(MKD$(SUTOTPED#), 3, 11, 2, 2)
   Total(0).Refresh
   '
   DESC$ = TRIM$(PORDESCU.TEXT)
   If TRIM$(DESC$) <> "" Then
      TTT1 = XVALO(Total(0).TEXT) * COEFDESCU(DESC$) / 100
   Else
      TTT1 = 0
   End If
   
   Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   '
   TTT3 = XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT)
End Sub

Sub CALTOIVA()
   '
   Static CLIA$, RNI%, APLIVA%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE = XVALO(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIE > 0 Then
       PVCLI% = PIVACLI%(NCLIE)
       If PVCLI% = 2 Then RNI% = 1
       If PVCLI% = 5 Or PVCLI% = 7 Then APLIVA% = 0
     End If
   End If
   '
   CIXXI% = CODINT%(Text13)
   TOT# = XVALO(Total(3))
   '
   IMPIVA21 = ""
   IMPIVA105 = ""
   If GRUPIVA%(CIXXI%) < 1 Then
       IMPOIVA# = APLIVA% * (XVALO(IMPIVA21) + ROUND#(TOT# * PORCEIVA(CIXXI%) / 100, 2))
       IMPIVA21 = TRIM$(FORMATNUM$(IMPOIVA#, "F12.2"))
     Else
       IMPOIVA# = APLIVA% * (XVALO(IMPIVA105) + ROUND#(TOT# * PORCEIVA(CIXXI%) / 100, 2))
       IMPIVA105 = TRIM$(FORMATNUM$(IMPOIVA#, "F12.2"))
   End If
   '
   If Picture1.Visible = True Then
      IMPIVA21 = ""
      IMPIVA105 = ""
      DESC$ = TRIM$(PORDESCU.TEXT)
      FIN& = ULTREG&("!CARDETPR")
      For U& = 1 To FIN&
        X$ = REGLEIDO$("!CARDETPR", U&)
        CIXXI% = CVI(Left$(X$, 2))
        TOT# = CVD(Mid$(X$, 55, 8))
        SUIVA# = SUIVA# + (Int(TOT# * PORCEIVA(CIXXI%) + 0.5)) / 100
        '
        If GRUPIVA%(CIXXI%) < 1 Then
            IMPOIVA# = APLIVA% * (XVALO(IMPIVA21) + ROUND#(TOT# * (PORCEIVA(CIXXI%) / 100) * (1 - COEFDESCU(DESC$) / 100), 2))
            IMPIVA21 = TRIM$(FORMATNUM$(IMPOIVA#, "F12.2"))
          Else
            IMPOIVA# = APLIVA% * (XVALO(IMPIVA105) + ROUND#(TOT# * (PORCEIVA(CIXXI%) / 100) * (1 - COEFDESCU(DESC$) / 100), 2))
            IMPIVA105 = TRIM$(FORMATNUM$(IMPOIVA#, "F12.2"))
        End If
      Next U&
   End If
   '
   SUIVA# = XVALO(IMPIVA21) + XVALO(IMPIVA105)
   '
   SUIVARNI# = SUIVA# * RNI% / 2
   TOTIVA = TRIM$(CSTRING$(MKD$(SUIVA# + SUIVARNI#), 4, 12, 2, 2))
   '
   Call CALPERIB
   '
End Sub

Sub CALPERIB()

   NCLIE1 = XVALO(Text1)
   
   PIVADOC% = PIVACLI%(NCLIE)
   '
   VARIDOC% = 1
   If PIVADOC% = 0 Or PIVADOC% = 3 Or PIVADOC% = 4 Or PIVADOC% = 6 Then
        VARIDOC% = 2
   ElseIf PIVADOC% = 5 Or PIVADOC% = 7 Then
        VARIDOC% = 3
   End If
   '
   TASAX = ALIPERBRU#(NCLIE1, MINIPER#, VARIDOC%)
   '
   TOTPER = ""
   BIMPON# = XVALO(Total(3))
   If BIMPON# * TICAMONE(MONEMI%) >= MINIPER# Then
      PERSALTA# = BIMPON# * (TASAX / 100)
      TOTPER = FORMATNUM$(PERSALTA#, "F11.2")
   End If
   '
End Sub
'
Sub PREPORCOP(NPRESU&, NCLIE)
   '
   Refrescar% = COMALTER%("Puede Mantener los Precios del Presupuesto de Origen o bien Actualizarlos \Segun la Lista de Precios Vigente\\No, Continuar\Si, Actualizar")
   If Refrescar% = 0 Then Exit Sub
   '
   NUCLI = NCLIE 'CVI(Mid$(XX$, 7, 2))'CLIENTE QUE SE ELIGIO
   Call ABREPRESVEN
   '
   SQLX$ = " SELECT * "
   SQLX$ = SQLX$ + " FROM presuven "
   SQLX$ = SQLX$ + " WHERE NroPres = " & NPRESU& & " "
   SQLX$ = SQLX$ + " ORDER BY NuOrdIt "
   '
   'Set PRESUTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   Dim PRESUTEMP As ADODB.Recordset
   Set PRESUTEMP = New ADODB.Recordset
   PRESUTEMP.CursorLocation = adUseClient
   PRESUTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   With PRESUTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        REFE$ = !Referen
        ATESR$ = !AteSen
        LIPREC% = !LisPre
        COPA% = !ConPag
        VENDE1% = !VENDED
        MONEPRE% = !MonedEmi
        CODI1% = !CODIINT
        PORDESCU1 = !PorcDesc
        PREUNI1 = !PreList
        ESPE$ = !Especif
        CANTI0 = !Cantit '\\\OT-06-0168-WAR-09/05/06///
        '
        
        '
        'CARGA DE ENCABEZADO DE PRESUPUESTO
          Text1 = TRIM$(Str$(NUCLI))
          Call CARDACLI
          TEXT21(0) = TRIM$(Str$(LIPREC%))
          TEXT21(1) = TRIM$(Str$(COPA%))
          TEXT21(2) = TRIM$(Str$(VENDE1%))
          Text10 = REFE$
          Text12 = ATESR$
          '
          'CARGO MONEDA DE ORIGEN
          MONEMI% = MONEPRE%
          MONEMIS = DEMONECO$(MONEMI%)
          MONEMIS.ForeColor = RGB(0, 0, 255)
             If MONEMI% > 1 Then MONEMIS.ForeColor = RGB(255, 0, 0)
          MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
          TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
          '
        'CARGO COMPLEMENTOS DEL PRESUPUESTO SI LOS TIENE
          '  Z$ = RECFILT$("PRESUPU1", 1, MKS$(NPRESU&))
          '  NREG1& = CVS(Z$)
          '  If NREG1& >= 1 Then
          '    ZZ$ = REGLEIDO$("PRESUPU1", NREG1&)
          '    Total(2) = Mid$(ZZ$, 5, 48)
          '    Total(4) = Mid$(ZZ$, 53, 48)
          '    Total(6) = Mid$(ZZ$, 101, 32)
          '    Total(5) = Mid$(ZZ$, 133, 32)
          '  End If
        
        .MoveLast
        NUIT1& = .RecordCount
        If NUIT1& = 1 Then
            'ITEM UNICO
            Label12 = "LISTA REPUESTOS"
            Call Label12_DblClick
            '\\\OT-06-0168-WAR-09/05/06///
            Text13 = CODEXT$(CODI1%)
            Text14 = FORMATNUM$(CANTI0, "F11.1")
            Total(0) = FORMATNUM$(PREUNI1 * CANTI0, "F11.2")
            PORDESCU = FORMATNUM$(PORDESCU1, "F5.2")
            '\\\OT-06-0168-WAR-FIN///
'            Total(0) = FORMATNUM$(PREUNI1, "F11.2")
'            PORDESCU = FORMATNUM$(PORDESCU1, "F5.2")
            Text28 = ESPE$
            '
            'SI SE ELIGIO ACTUALIZAR PRECIO
            If Refrescar% = 2 Then
               Condi$ = "Nume_Cli = " & NCLIE
               LISTA% = XVALO(VALOBADA("DEUDOR12", "Lpre_Cli", Condi$))
               Condi$ = "COD_LISTA = " & LISTA% & " AND COD_PIEZA = " & CODI1%
               '
               PREUNIX = XVALO(VALOBADA("PREVENTA", "PRECIO_UNITARIO", Condi$))
               If PREUNIX < 0.0001 Then
                  Msg$ = "No Se Ha Encontrado el Precio del Articulo " & CODEXT(CODI1%)
                    Msg$ = Msg$ & "\en la Lista de Precios " & LISTA% & ".\"
                    Msg$ = Msg$ & "Se Mantiene el que Proviene del Presupuesto."
                  Call COMUNI(Msg$)
                 Else
                  PREUNI1 = PREUNIX
                End If
            End If
          Else
            'LISTA DE REPUESTOS
            Label12 = "IT. UNICO"
            Call Label12_DblClick
            '
            XX$ = REGBLAN$("CARDETPR")
            JJ& = 0
            Call BLANARCH("!CARDETPR")
            .MoveFirst
            Condi$ = "Nume_Cli = " & NCLIE
            LISTA% = XVALO(VALOBADA("DEUDOR12", "Lpre_Cli", Condi$))
            Do While Not .EOF
              CODI1% = !CODIINT
              PRELI1 = !PreList
              IMPOUN = !PREUNIT
              CAN1 = !Cantit
              '
              If Refrescar% = 2 Then
                 Condi$ = "COD_LISTA = " & LISTA% & " AND COD_PIEZA = " & CODI1%
                 PREUNIX = XVALO(VALOBADA("PREVENTA", "PRECIO_UNITARIO", Condi$))
                 If PREUNIX < 0.0001 Then
                    Msg$ = "No Se Ha Encontrado el Precio del Articulo " & CODEXT(CODI1%)
                    Msg$ = Msg$ & "\en la Lista de Precios " & LISTA% & ".\"
                    Msg$ = Msg$ & "Se Mantiene el que Proviene del Presupuesto."
                    Call COMUNI(Msg$)
                   Else
                    PREUNI1 = PREUNIX
                 End If
              End If
              '
              Call REPLA(XX$, MKI$(CODI1%), 1, 2)
              Call REPLA(XX$, MKS$(PRELI1), 39, 4)
              Call REPLA(XX$, MKS$(IMPOUN), 65, 4)
              Call REPLA(XX$, MKS$(CAN1), 47, 4)
              JJ& = JJ& + 1
              Call GRAREG("!CARDETPR", XX$, JJ&)
            .MoveNext
            Loop
            Call SETULTREG("!CARDETPR", JJ&)
            Call CARDETPRES
        End If
      End If
   End With
   PRESUTEMP.Close
   Set PRESUTEMP = Nothing
   '
   ALPORCOPIA% = 1 'MARCA PARA SABER SI ESTA EN ALTA POR COPIA
   '
End Sub


Sub CONAPLI(APLINVO$, TERMINA%)
   TERMINA% = 1
   Select Case APLINVO$
       '
       'OPCIONES VARIAS DE PRESUPUESTOS
       '
       Case "OPVARPRE"
           Call Command12_Click             ' OPCIONES VARIAS - PRESUPUESTOS
       Case "PRENOVIG"
           Call OPLIPRE07.Command3_Click    ' MARCA PRESUPUESTO COMO NO VIGENTE
       Case "GENPRECOP"
           Call Command7_Click              ' GENERACION DE PRESUPUESTO POR COPIA
           TERMINA% = 0
       '
       'CONSULTAS Y LISTADOS DE PEDIDOS
       '
       Case "CONLISPRE"
           Call Command12_Click             ' CONSULTAS Y LISTADOS
       Case "VERIMPRESU"
           Call OPLIPRE07.Command1_Click    ' ENCABEZA LISTADO DE PEDIDOS
       Case "PRERANFE"
           Call OPLIPRE07.Command6_Click    ' LISTADO POR RANGO DE FECHAS
       '
       'CONFIGURACION DE PRESUPUESTOS
       '
       Case "CONGENPRES"
           Call Command13_Click              ' CONFIGURACION GENERAL DE PRESUPUESTOS
      
       Case Else
           TERMINA% = 0
   End Select
End Sub

Sub PRINTEXT(ESPECIFIX$, NPARA%, PARAMEXT$)

       FORTELIB$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, PARAMEXT$ + "/FORMATO")))
       ATELIB% = XVALO(Mid$(FORTELIB$, 2))
       '
       XX$ = ESPECIFIX$
       Call FRATEXTO(XX$, ATELIB%)
       For JJ& = 1 To CARETEX%
            CAITAB1% = CAITAB1% + 1
            TETABU1$(NPARA%, CAITAB1%) = RETEX$(JJ&)
       Next JJ&

End Sub

