VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FACPROFORMA 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Factura Proforma"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   64
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   7965
      Left            =   7320
      ScaleHeight     =   7935
      ScaleWidth      =   1185
      TabIndex        =   60
      Top             =   0
      Width           =   1215
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   84
         Top             =   6840
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   85
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
         Left            =   120
         Picture         =   "FACPROFORMA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Aprueba y Graba Documento"
         Top             =   5040
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
         Height          =   640
         Left            =   105
         Picture         =   "FACPROFORMA.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Resumen de Cuenta Corriente"
         Top             =   1890
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
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
         Left            =   105
         Picture         =   "FACPROFORMA.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   210
         Width           =   650
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   -50
      TabIndex        =   0
      Top             =   120
      Width           =   12000
      Begin VB.CommandButton Command6 
         Caption         =   "Abrir"
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
         Left            =   5460
         Picture         =   "FACPROFORMA.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   7080
         Width           =   650
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
         Height          =   300
         Left            =   2760
         TabIndex        =   88
         Top             =   7080
         Width           =   175
      End
      Begin VB.TextBox Text7 
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
         Left            =   1680
         TabIndex        =   87
         Top             =   7080
         Width           =   1065
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Guardar Formulario"
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
         Left            =   6160
         Picture         =   "FACPROFORMA.frx":0A68
         TabIndex        =   81
         ToolTipText     =   "Aprueba y Graba Documento"
         Top             =   7080
         Width           =   1005
      End
      Begin VB.TextBox Text111 
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
         Left            =   1695
         TabIndex        =   78
         Top             =   7440
         Width           =   1065
      End
      Begin VB.CommandButton Command44 
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
         Left            =   2760
         TabIndex        =   77
         Top             =   7440
         Width           =   175
      End
      Begin VB.CommandButton Command66 
         Caption         =   "Abrir"
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
         Left            =   5460
         Picture         =   "FACPROFORMA.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   7440
         Width           =   650
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   2
         Top             =   1280
         Width           =   610
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "COMERCIAL"
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
         Height          =   2500
         Index           =   0
         Left            =   210
         TabIndex        =   38
         Top             =   1600
         Width           =   6960
         Begin VB.CommandButton Command210 
            Caption         =   "Proformas Aplicadas al Pedido"
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
            Left            =   3360
            TabIndex        =   93
            Top             =   2160
            Width           =   3420
         End
         Begin VB.TextBox TEXT21 
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
            Index           =   2
            Left            =   1410
            Locked          =   -1  'True
            TabIndex        =   92
            Text            =   " "
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CommandButton Command211 
            Caption         =   "."
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
            Left            =   2520
            TabIndex        =   74
            Top             =   1770
            Width           =   180
         End
         Begin VB.TextBox TEXT21 
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
            Index           =   4
            Left            =   4275
            Locked          =   -1  'True
            TabIndex        =   72
            Text            =   " "
            Top             =   1770
            Width           =   2550
         End
         Begin VB.TextBox TEXT211 
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
            Left            =   1410
            TabIndex        =   70
            Text            =   " "
            Top             =   1770
            Width           =   1100
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
            Height          =   300
            Index           =   1
            Left            =   2130
            Locked          =   -1  'True
            TabIndex        =   68
            Text            =   " "
            Top             =   1340
            Width           =   4695
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
            Height          =   300
            Index           =   1
            Left            =   1410
            TabIndex        =   67
            Text            =   " "
            Top             =   1340
            Width           =   435
         End
         Begin VB.CommandButton Command21 
            Caption         =   "."
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
            Left            =   1860
            TabIndex        =   66
            Top             =   1340
            Width           =   180
         End
         Begin VB.TextBox DescriTxt 
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
            Height          =   660
            Left            =   1410
            MaxLength       =   64
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   240
            Width           =   5400
         End
         Begin VB.CommandButton Command21 
            Caption         =   "."
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
            Left            =   1860
            TabIndex        =   41
            Top             =   955
            Width           =   180
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
            Height          =   300
            Index           =   0
            Left            =   1410
            TabIndex        =   40
            Text            =   " "
            Top             =   955
            Width           =   435
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
            Height          =   300
            Index           =   0
            Left            =   2130
            Locked          =   -1  'True
            TabIndex        =   39
            Text            =   " "
            Top             =   955
            Width           =   4695
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Imp. Total Ped."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Index           =   6
            Left            =   120
            TabIndex        =   91
            Top             =   2250
            Width           =   1185
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "O.C. Cliente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   3000
            TabIndex        =   73
            Top             =   1860
            Width           =   1185
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Pedido"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   110
            TabIndex        =   71
            Top             =   1840
            Width           =   1185
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cond.Pago"
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
            Index           =   1
            Left            =   240
            TabIndex        =   69
            Top             =   1430
            Width           =   1065
         End
         Begin VB.Label Ltota 
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
            Height          =   330
            Index           =   0
            Left            =   -165
            TabIndex        =   43
            Top             =   690
            Width           =   1485
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
            Index           =   0
            Left            =   345
            TabIndex        =   42
            Top             =   1030
            Width           =   945
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Contable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   870
         Left            =   210
         TabIndex        =   30
         Top             =   3120
         Visible         =   0   'False
         Width           =   6960
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
            Index           =   3
            Left            =   1890
            TabIndex        =   34
            Text            =   " "
            Top             =   1130
            Width           =   3270
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
            Index           =   3
            Left            =   1155
            TabIndex        =   33
            Text            =   " "
            Top             =   1130
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
            Height          =   300
            Index           =   3
            Left            =   1575
            TabIndex        =   32
            Top             =   1130
            Width           =   180
         End
         Begin VB.TextBox Text16 
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
            Left            =   1890
            TabIndex        =   5
            Top             =   345
            Width           =   1170
         End
         Begin VB.CommandButton Command12 
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
            Left            =   3045
            TabIndex        =   31
            Top             =   345
            Width           =   180
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
            Index           =   3
            Left            =   -3045
            TabIndex        =   37
            Top             =   1205
            Width           =   4110
         End
         Begin VB.Label Label21 
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
            Height          =   330
            Index           =   5
            Left            =   210
            TabIndex        =   36
            Top             =   420
            Width           =   1590
         End
         Begin VB.Label Label16 
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
            Height          =   300
            Left            =   3330
            TabIndex        =   35
            Top             =   345
            Width           =   3375
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "RESUMEN FINAL"
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
         Height          =   2790
         Left            =   210
         TabIndex        =   21
         Top             =   4150
         Width           =   6960
         Begin VB.Frame Frame9 
            BackColor       =   &H00E2D3C0&
            BorderStyle     =   0  'None
            Height          =   1365
            Left            =   1890
            TabIndex        =   44
            Top             =   840
            Width           =   4900
            Begin VB.TextBox Total 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   5
               Left            =   3150
               TabIndex        =   47
               TabStop         =   0   'False
               Text            =   " "
               Top             =   930
               Width           =   1650
            End
            Begin VB.TextBox Total 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   4
               Left            =   3150
               TabIndex        =   46
               TabStop         =   0   'False
               Text            =   " "
               Top             =   540
               Width           =   1650
            End
            Begin VB.TextBox Total 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   3
               Left            =   3150
               TabIndex        =   45
               TabStop         =   0   'False
               Text            =   " "
               Top             =   145
               Width           =   1650
            End
            Begin VB.Label Ltota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "IVA-NI"
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
               Left            =   -405
               TabIndex        =   59
               Top             =   1010
               Width           =   960
            End
            Begin VB.Label Ltota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "IVA"
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
               Left            =   -525
               TabIndex        =   58
               Top             =   620
               Width           =   1065
            End
            Begin VB.Label Ltota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "IVA"
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
               Left            =   -315
               TabIndex        =   57
               Top             =   220
               Width           =   855
            End
            Begin VB.Label Ltota 
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
               Index           =   7
               Left            =   1365
               TabIndex        =   56
               Top             =   220
               Width           =   225
            End
            Begin VB.Label Ltota 
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
               Index           =   8
               Left            =   1365
               TabIndex        =   55
               Top             =   620
               Width           =   225
            End
            Begin VB.Label Ltota 
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
               Index           =   9
               Left            =   1365
               TabIndex        =   54
               Top             =   1010
               Width           =   225
            End
            Begin VB.Label PORIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Left            =   630
               TabIndex        =   53
               Top             =   220
               Width           =   645
            End
            Begin VB.Label PORIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Index           =   1
               Left            =   630
               TabIndex        =   52
               Top             =   620
               Width           =   645
            End
            Begin VB.Label PORIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Left            =   600
               TabIndex        =   51
               Top             =   1010
               Width           =   645
            End
            Begin VB.Label BASIVA 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   0
               Left            =   1680
               TabIndex        =   50
               Top             =   145
               Width           =   1380
            End
            Begin VB.Label BASIVA 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   1
               Left            =   1680
               TabIndex        =   49
               Top             =   540
               Width           =   1380
            End
            Begin VB.Label BASIVA 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   330
               Index           =   2
               Left            =   1680
               TabIndex        =   48
               Top             =   930
               Width           =   1380
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00E2D3C0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1890
            TabIndex        =   26
            Top             =   2025
            Width           =   4860
            Begin VB.TextBox Text15 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   3360
               TabIndex        =   8
               Top             =   330
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.TextBox Text17 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   1155
               TabIndex        =   7
               Top             =   330
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Perc. I-B: "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   2625
               TabIndex        =   28
               ToolTipText     =   "Doble Click para Configurar Cliente"
               Top             =   180
               Visible         =   0   'False
               Width           =   645
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Perc. I.V.A.:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   315
               TabIndex        =   27
               ToolTipText     =   "Doble Click para Configurar Cliente"
               Top             =   225
               Visible         =   0   'False
               Width           =   750
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00E2D3C0&
            BorderStyle     =   0  'None
            Height          =   675
            Left            =   240
            TabIndex        =   23
            Top             =   230
            Width           =   6525
            Begin VB.TextBox Total 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   2
               Left            =   4800
               TabIndex        =   6
               TabStop         =   0   'False
               Text            =   " "
               Top             =   240
               Width           =   1650
            End
            Begin VB.Label Ltota 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Importe Total Neto"
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
               Index           =   10
               Left            =   2610
               TabIndex        =   25
               Top             =   345
               Width           =   2010
            End
            Begin VB.Label MONEMIS 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Pesos"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   12
                  Charset         =   0
                  Weight          =   900
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00005500&
               Height          =   330
               Left            =   975
               TabIndex        =   24
               ToolTipText     =   "Doble-Click para 'Dólares'"
               Top             =   240
               Width           =   1380
            End
         End
         Begin VB.TextBox Total 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   6
            Left            =   5040
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2400
            Width           =   1650
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL DOCUMENTO"
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
            Index           =   11
            Left            =   2970
            TabIndex        =   29
            Top             =   2520
            Width           =   1905
         End
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
         TabIndex        =   16
         Top             =   890
         Width           =   180
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
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   " "
         Top             =   105
         Width           =   3480
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
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   " "
         Top             =   490
         Width           =   3480
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
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
         Left            =   1470
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   " "
         Top             =   890
         Width           =   750
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
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   " "
         Top             =   890
         Width           =   2640
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
         Left            =   6990
         TabIndex        =   11
         Top             =   105
         Width           =   180
      End
      Begin VB.TextBox Text9 
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
         Height          =   300
         Left            =   6045
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   105
         Width           =   960
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5775
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   3
         Text            =   " "
         Top             =   1280
         Width           =   1440
      End
      Begin VB.TextBox Text12 
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
         MaxLength       =   40
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1280
         Width           =   3480
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   0
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
         MaxLength       =   5
         TabIndex        =   1
         Text            =   " "
         Top             =   890
         Width           =   855
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
         Height          =   285
         Left            =   3000
         TabIndex        =   90
         Top             =   7080
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Prof. Mat.Codific.:"
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
         Left            =   240
         TabIndex        =   86
         Top             =   7080
         Width           =   1365
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
         Height          =   195
         Left            =   210
         TabIndex        =   83
         Top             =   650
         Width           =   960
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   770
         Left            =   6550
         TabIndex        =   82
         Top             =   480
         Width           =   655
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Height          =   765
         Left            =   6550
         Top             =   480
         Width           =   670
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Prof. Texto Libre:"
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
         Index           =   3
         Left            =   240
         TabIndex        =   80
         Top             =   7530
         Width           =   1365
      End
      Begin VB.Label TEXT112 
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
         Height          =   285
         Left            =   3000
         TabIndex        =   79
         Top             =   7440
         Visible         =   0   'False
         Width           =   2385
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
         ForeColor       =   &H00004040&
         Height          =   330
         Left            =   5200
         TabIndex        =   20
         Top             =   200
         Width           =   750
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5633
         TabIndex        =   19
         Top             =   1260
         Width           =   150
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Factura Proforma"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   645
         Left            =   5250
         TabIndex        =   18
         Top             =   600
         Width           =   1080
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
         TabIndex        =   17
         Top             =   1280
         Width           =   1035
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FACPROFORMA.frx":107C
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "FACPROFORMA.frx":12B0
      TabIndex        =   65
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   75
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Menu mnuTransaccion 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuAnular 
         Caption         =   "Anular Proforma (Pendiente de Factura)"
      End
      Begin VB.Menu mnuRevAnul 
         Caption         =   "Revertir Anulación"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
      Begin VB.Menu mnuProforFacturadas 
         Caption         =   "Visualizar Proformas Facturadas"
      End
   End
End
Attribute VB_Name = "FACPROFORMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TIFADOC%    ' 2=FACTURA / 3=N.DEBITO / 6=N.CREDITO
Dim MODODOC%    ' 0=MAT.CODIFICADO / 1=TEXTO LIBRE
Dim VARIDOC%
Dim RNIDOC%
Dim DATEDOC%
Dim PIVADOC%
Dim VERLICRE%
Dim MODESCU%
'
Dim CONSICLI%
'
Dim NUOCOMPRA$
Dim NROPEDIDO$
Dim NROREMICO$
'
Dim ARCHIPRE$
Dim DATCLI$(32)
'
Dim MONECO$, TICAMBX#

Private Sub BASIVA_DBLClick(Index As Integer)
   '
   If Index <= 1 Then
       NCLIE% = XVALO(Text1)
       RNI% = 0: APLIVA% = 1
       If NCLIE% > 0 Then
         PPVI% = PIVACLI%(NCLIE%)
         If PPVI% = 2 Then RNI% = 1
         If PPVI% = 5 Or PPVI% = 7 Then APLIVA% = 0
       End If
   End If
   '
   Select Case Index
       Case 0 '21%
          If XVALO(BASIVA(0)) < 0.005 Then
            BASIVA(0) = BASIVA(1)
            Total(3).TEXT = CSTRING$(MKD$(XVALO(BASIVA(0).Caption) * XVALO(PORIVA(0).Caption) * APLIVA% / 100), 3, 12, 2, 2)
            BASIVA(1) = ""
            Total(4).TEXT = CSTRING$(MKD$(XVALO(BASIVA(1).Caption) * XVALO(PORIVA(1).Caption) * APLIVA% / 100), 3, 12, 2, 2)
          End If
          Call CALTOFA2
       Case 1 '10.5%
          If XVALO(BASIVA(1)) < 0.005 Then
            BASIVA(1) = BASIVA(0)
            AA1 = XVALO(BASIVA(1))
            aa2 = PORIVA(1)
            aa3 = APLIVA%
            Total(4).TEXT = CSTRING$(MKD$(XVALO(BASIVA(1).Caption) * XVALO(PORIVA(1).Caption) * APLIVA% / 100), 3, 12, 2, 2)
            BASIVA(0) = ""
            Total(3).TEXT = CSTRING$(MKD$(XVALO(BASIVA(0).Caption) * XVALO(PORIVA(0).Caption) * APLIVA% / 100), 3, 12, 2, 2)
          End If
          Call CALTOFA2
    End Select
    '
End Sub

Private Sub BASIVA_LostFocus(Index As Integer)
    BASIVA(Index) = FORMATNUM(BASIVA(Index), "F11.2")
End Sub

Private Sub Command1_Click()
   
   NCLIE = ListaSeleccionClientes()
   If NCLIE > 0 Then
     Text1.TEXT = TRIM$(Str$(NCLIE))
     If XVALO(Text1.TEXT) > 0 Then
         NC = XVALO(Text1)
           Call CARDACLI
       End If
     End If
   End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   Call CONECRUN("ARCHIG07", "Archivos Maestros")
   Command11.Enabled = True
End Sub


Private Sub Command12_Click()
   Call SELECHO("ECUECON")
   VDEVU$ = TRIM$(VALACT1$("ECUECON"))
   If VDEVU$ <> "" Then
     Text16 = VDEVU$
   End If
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   COGEVE07.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   '
   Screen.MousePointer = 11
   '
   FECHALIM% = XVALO(CONTROL$("", "CIEMEVEN"))
   If FECHANUM(Text9) <= FECHALIM% Then
     Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
     Command14.Enabled = True
     Exit Sub
   End If
   '
   'Valida que el Nº no sea 0. -21/12/04-
   If XVALO(TRIM$(Text11)) <= 0 Then
        Call MENSERR(24, "Imposible Registrar este Documento.\Numero de Proforma Incorrecto. - Consulte.")
        Screen.MousePointer = 1
        Command14.Enabled = True
        Exit Sub
   End If
   '
   If TRIM$(Text6) = "" Or TRIM$(Text11) = "" Then
        Call MENSERR(24, "Imposible Registrar este Documento.\Numero de Proforma Incompleto. - Consulte.")
        Screen.MousePointer = 1
        Command14.Enabled = True
        Exit Sub
   End If
   '
   For J = 0 To 2
    If XVALO(Total(J + 3)) > 0 Then
      If XVALO(BASIVA(J)) < 0.005 Then
        Call MENSERR(24, "Imposible Registrar este Documento.\Base Imponible de Iva de la Proforma Incompleta. - Consulte.")
        Screen.MousePointer = 1
        Command14.Enabled = True
        Exit Sub
      End If
    End If
   Next J
   '
   For J = 0 To 2
    If XVALO(Total(J + 3)) < 0.005 Then
      If XVALO(BASIVA(J)) > 0.005 Then
        OPX% = COMALTER("Este Documento No Posee Importe de Iva.\¿Desea Continuar o Cancelar?\\Cancelar\Continuar")
        If OPX% <= 1 Or OPX% > 2 Then
            Screen.MousePointer = 1
            Command14.Enabled = True
            Exit Sub
        End If
      End If
    End If
   Next J
   '
   If TRIM$(DescriTxt.TEXT) = "" Then
    Call MENSERR(24, "Imposible Registrar este Documento.\Descripcion de la Proforma Incompleta. - Consulte.")
    Screen.MousePointer = 1
    Command14.Enabled = True
    Exit Sub
   End If
   '
   If XVALO(Total(6)) <= 0.0005 Then
    Call MENSERR(24, "Imposible Registrar este Documento.\Debe Completar el Importe de la Proforma. - Consulte.")
    Screen.MousePointer = 1
    Command14.Enabled = True
    Exit Sub
   End If
   '
'   If TRIM$(Text16.TEXT) = "" Then
'    Call MENSERR(24, "Imposible Registrar este Documento.\Cuenta Contable de la Proforma Incompleta. - Consulte.")
'    Screen.MousePointer = 1
'    Command14.Enabled = True
'    Exit Sub
'   End If
   '
'   If CUECOINT(Text16.TEXT) <= 0 Then
'    Call MENSERR(24, "Imposible Registrar este Documento.\Cuenta Contable de la Nota de Debito Inexistente. - Consulte.")
'    Screen.MousePointer = 1
'    Command14.Enabled = True
'    Exit Sub
'   End If
   '
   
   '
'6 Call REGISNODE
  Call REGISTRAPROFORMA
  FORIMPRE$ = TRIM$(CONTROL("", "FACPRO-A"))
  If FORIMPRE$ = "" Then
      Call COMUNI("Falta Configurar Formulario de Impresión de Factura Proforma")
  Else
    FILTX$ = TRIM$(Str$(XVALO(Text11))) & ";" & TRIM$(Str$(Text1) & ";" & TRIM$(Label13) & ";" & TRIM$(Text11))
    Call STDFORM(FORIMPRE$, FILTX$)
  End If
  
   '
10 Call BLANFORMU
   FACPROFORMA.Text11 = FORMATOCON0$(PROXIMONUMPROFORMA&, 8)
   FACPROFORMA.Text6 = "0000"

   Call CIERRARCH("*.*")
   Text1.TEXT = ""
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
   '
   Screen.MousePointer = 1
   Command14.Enabled = True
End Sub

Sub REGISTRAPROFORMA()
    '
    TIPOCOMP$ = "FP"
    numeroComp& = XVALO(Mid$(Text11, 1))
    NUMEROTALO$ = XVALO(Mid$(Text6, 1))
    'CORTO
    NCCMP$ = TIPOCOMP$ & "-" & FORMATOCON0$(numeroComp&, 4)
    'LARGO
    NLCMP$ = TIPOCOMP$ & "-" & TRIM$(Text6) & "-" & TRIM$(Text11)
    '
    TOTDOCUM# = XVALO(Total(6))
    '
    NUCOM& = PROXIMONUMPROFORMA&
    NC = XVALO(Text1)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    SQLX$ = "SELECT * FROM PROFORMTL WHERE NumeComp < 1 "
    rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

    With rs
        .AddNew
        !NumeComp = NUCOM&
        !NUMETALO = XVALO(Text6)
        !NUCLIEN = XVALO(Text1)
        !BASE1 = XVALO(BASIVA(0))
        !BASE2 = XVALO(BASIVA(1))
        !TASA1 = FORMATNUM$(XVALO(PORIVA(0)), "F6.2")
        !TASA2 = FORMATNUM$(XVALO(PORIVA(1)), "F6.2")
        !CODICONPAG = XVALO(TEXT21(1))
        !VEND_CLI = XVALO(TEXT21(0))
        !ImpoTot = XVALO(Total(6))
        !TOTNETO = XVALO(Total(2))
        AS1 = FORMATNUM$(XVALO(PORIVA(0)), "F6.2")
        !FechFactura = CVDAT(CVINT("01/01/80"))
        !FacturaReal = ""
        
        !FECHEMISI = CVDAT(FECHANUM(Text9))
        !FeHoReal = Now
        
        !TIPO_CAM = TICAMONE(MONEMI%)
        !MONE_EMIS = MONEMI%
        !Status = 0
        
        !COTIZA_DOLAR = TICAMONE(2)
        !REFERENCIA = DescriTxt
        !NROPED = XVALO(TEXT211)
        !IMPOIVA1 = XVALO(Total(3))
        !IMPOIVA2 = XVALO(Total(4))
        !VARIMOVIM = Label13
        !TIPOFACT = "TL" 'TEXTO LIBRE
        .Update
        '
    End With
    rs.Close: Set rs = Nothing
   '
End Sub

Private Sub Command21_Click(Index As Integer)
    If Index = 0 Then
        ARCHEC$ = "VENDEDOR": LOCA% = 1
      ElseIf Index = 1 Then
        ARCHEC$ = "CONPAG": LOCA% = 2
    End If
    ARGU$ = CBIN$(TRIM$(TEXT21(Index).TEXT), 1, LOCA%)
    '\\\OT-06-0570-RG-27/12/06///
    Call CommSel(ARCHEC$, Command21(Index), TEXT21(Index))

End Sub

Private Sub Command210_Click()
    Command210.Enabled = False
    If XVALO(TEXT211) < 1 Then GoTo 99
    NPED& = XVALO(TEXT211)
    CAMPOS$ = "Nro. Cot./F12;M./A3;Imorte Neto/F12.2;Factura/A18;Referencia/A64"
    
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, CAMPOS$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
    FRMZELECHO.Caption = "Proformas Asociadas al Pedido Activo"
    Dim obj As New RECORXET
    Set RS1 = obj.RS_ProformasXPedido(NPED&)
    Call Carga_MSF_Recordset(RS1, CAMPOS$, FRMZELECHO.MSF2, 1)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    FRMZELECHO.Width = 12000
    FRMZELECHO.Show 1
    
    
    
99  Command210.Enabled = True
End Sub

Private Sub Command211_Click()
    Command211.Enabled = False
    If XVALO(Text1) < 1 Then GoTo 99
    NC = XVALO(Text1)
    CAMPOS$ = "Nro. Pedido/F12;Referencia/A64"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, CAMPOS$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
    FRMZELECHO.Caption = "PEDIDOS PENDIENTES DEL CLIENTE ACTIVO"

    Dim obj As New RECORXET
    Set RS1 = obj.RS_PedidoPendxCte2(NC)
    Call Carga_MSF_Recordset(RS1, CAMPOS$, FRMZELECHO.MSF2, 1)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    FRMZELECHO.Width = 10000
    FRMZELECHO.Show 1
    NPF& = 0
    If XVALO(RESP_ZELE_VECT(1)) > 0 Then
      TEXT211 = XVALO(RESP_ZELE_VECT(1))
      NPF& = PedidoTeineProforma&(XVALO(TEXT211))
    Else
      TEXT211 = ""
    End If
    If NPF& > 0 Then
         OPX% = COMALTER%("Atención !!! El Pedido Seleccionado Yá Generó Una Factura Proforma (N° " & NPF& & ")\\Cancelar\Continuar")
         If OPX% <= 1 Then Call BLANFORMU
         Total(2) = ""
    End If

99  Command211.Enabled = True
End Sub


Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Label2 = VALACT2$("INDIFRM")
    End If

End Sub

Private Sub Command4_Click()
    Command4 = False
    RECUEC07.Text1.TEXT = Text1.TEXT
    RECUEC07.Show 1
    Command4.Enabled = True

End Sub

Private Sub Command44_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text111.TEXT = VDEVU$
        TEXT112 = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command66_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text111.TEXT), "Formularios de Impresion")

End Sub

Private Sub mnuAnular_Click()
    NPROF& = SELECCION_FACTURAPROFORMA%(0)
    If NPROF& > 0 Then
       CONDI$ = "NUMECOMP = " & NPROF&
       Call ACTUREGISTRO("PROFORMTL", "STATUS", CONDI$, 9, REGAF&, "NOMESS")
       Call COMUNI("La Proforma N°: " & NPROF& & " Se Ha Anulado")
    End If
End Sub

Private Sub mnuProforFacturadas_Click()
    CAMPOS$ = "Fecha/D8;N°Prof./F8;Referencia/A48;Imp.Neto/F12.2;Imp.Total/F12.2;Factura/A18;F.Fact./D8;NC/A0"
    CAMPOXSEL$ = "Fechemisi,NumeComp,REFERENCIA,TotNeto,ImpoTot,FacturaReal,FechFactura,NUCLIEN"
10  NPROF& = SELECCION_FACTURAPROFORMA%(0, CAMPOS$, CAMPOXSEL$, 3)
    If NPROF& < 1 Then
      Exit Sub
    End If

    If NPROF& > 0 Then
            If EsProformaXDetalle%(NPROF&) > 0 Then
                FORIMPRE$ = TRIM$(CONTROL("", "FAPRMACO"))
                If FORIMPRE$ = "" Then
                    Call COMUNI("Falta Configurar Formulario de Impresión de Factura Proforma por Detalle")
                    Exit Sub
                Else
                    FILTX$ = TRIM$(Str$(NPROF&) & ";" & FORMATOCON0$(NPROF&, 8))
                    Call STDFORM(FORIMPRE$, FILTX$)
                End If
                GoTo 10

            Else
                NC = XVALO(RESP_ZELE_VECT(8))
                If NC > 0 Then
                     FORIMPRE$ = TRIM$(CONTROL("", "FACPRO-A"))
                     If FORIMPRE$ = "" Then
                         Call COMUNI("Falta Configurar Formulario de Impresión de Factura Proforma")
                         Exit Sub
                     Else
                         LETR$ = LETRAPROFORMA$(NPROF&)
                         FILTX$ = TRIM$(Str$(NPROF&)) & ";" & TRIM$(Str$(NC) & ";" & LETR$ & ";" & FORMATOCON0$(NPROF&, 8))
                         Call STDFORM(FORIMPRE$, FILTX$)
                     End If
                End If
                GoTo 10
           End If
        End If
     







   GoTo 10



End Sub


Private Sub mnuRevAnul_Click()
    NPROF& = SELECCION_FACTURAPROFORMA%(0, "", "", 4)
    If NPROF& > 0 Then
       CONDI$ = "NUMECOMP = " & NPROF&
       Call ACTUREGISTRO("PROFORMTL", "STATUS", CONDI$, 0, REGAF&, "NOMESS")
       Call COMUNI("La Proforma N°: " & NPROF& & " Se Ha Revertido la Anulación")
    End If

End Sub


Private Sub Text15_Change()
   Call CALTOFA2
End Sub

Private Sub Text15_GotFocus()
   If VALICLIE% < 1 Then GoTo 99
   If TIFADOC% < 3 Then GoTo 98
   If XVALO(Total(2)) < 0.5 Then GoTo 98
   Exit Sub
   '
98 On Error Resume Next
   PORDESCU.SetFocus
   On Error GoTo 0
   Exit Sub
   '
99 On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Text15 = FORMATNUM$(XVALO(Text15), "F11.2")
  End If
End Sub

Private Sub Text15_LostFocus()
   Text15 = FORMATNUM$(XVALO(Text15), "F11.2")
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
   Command17.Enabled = False
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Then
       CONPEDID.Show 1
     Else
       Call SHOWPEDIPEN
   End If
   Command17.Enabled = True
End Sub

Private Sub Command18_Click()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Command19_Click()
   '
   NUOCOMPRA$ = ""
   NROPEDIDO$ = ""
   '
   
   '
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

Private Sub Command22_Click()
    Command22.Enabled = False
        
    Command22.Enabled = True
End Sub



Private Sub Command5_Click()
    Command5.Enabled = False
    If DERACCE%("STPFACTU", "Configuración de Funcionamiento Facturacion") > 1 Then
       Call GRACONTROL("", "FACPRO-A", Text111.TEXT)
       Call GRACONTROL("", "FAPRMACO", Text7.TEXT)
    End If
    
    Command5.Enabled = True
End Sub

'
Sub CARCONSIPEN()
   '
   Static DES%, HAS%
   '
   If DES% < 1 Then
2    FEX = HOY(HO$)
     HAS% = FEX
     HASS$ = FECHATEX$(FEX)
     DESS$ = "01" + Mid$(HASS$, 3)
     DES% = FECHANUM(DESS$)
   End If
   '
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
     CONSICLI% = 0
     Exit Sub
   End If
   '
   DES% = CVI(Left$(OBX$, 2))
   HAS% = CVI(Mid$(OBX$, 3, 2))
   If DES% > HAS% Then
      GoTo 2
   End If
   FEX = DES%: PERIO$ = FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " al " + FECHATEX$(FEX)
   '
'   DES1# = CDbl(CVDAT(DES%))
'   HAS1# = CDbl(CVDAT(HAS%))
    DES1 = CVDAT(DES%)
    HAS1 = CVDAT(HAS%)

   '
   'LPRE% = XVALO(TEXT21(0).TEXT)
   NCLIE% = XVALO(Text1.TEXT)
   ARCHIPRE$ = TRIM$(Str$(LPRE%)): LIPRACTI% = LPRE%
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   If LPRE% < 1 Then ARCHIPRE$ = "MASTER": LIPRACTI% = 0
   CONSICLI% = 0
   '
   REGODES& = 0
   Call SETULTREG("!CARDEFAC", 0)
   jj& = 0: CONSICLI% = 0
   '
   CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
   If Len(CMO$) <> 2 Then
     CMO$ = "RT"
   End If
   Screen.MousePointer = 11
   Call APERBASDAT("STOCKS")
   '
   SQLX$ = " SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCLIE% & " "
   SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
   SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
   SQLX$ = SQLX$ + " AND DoPriLar <> '' "
   SQLX$ = SQLX$ + " AND ((DoSecLar IS NULL) OR (DoSecLar = '') OR (DoSecLar LIKE 'P/F%')) "
   SQLX$ = SQLX$ + " AND FechMov >= '" & DES1 & "' "
   SQLX$ = SQLX$ + " AND FechMov <= '" & HAS1 & "' "
   '
   Dim CONSITMP As ADODB.Recordset
   Set CONSITMP = New ADODB.Recordset
   '
'25 On Error Resume Next
'   CONSITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 25
'   End If
'   On Error GoTo 0
'   '
   Set CONSITMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   With CONSITMP
     'On Error Resume Next
     '.MoveFirst
     'If Err < 1 Then
       RECONSI$ = REGBLAN("CONSICLI")
       J& = 0
       Do While .EOF = False
         CODIMAT% = !Cod_Inte
         CAN = !Cantsalid - !CantIngre
         NPEDID& = 0: If IsNull(!NUPEDCLI) = False Then NPEDID& = !NUPEDCLI
5        Z$ = REGBLAN$("!CARDEFAC")
         Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CODIMAT%)
         Call REPLA(Z$, MKI$(CODIMAT%), 5, 2)
         Call REPLA(Z$, DESCRIT$(CODIMAT%), 7, 64)
           NREMIT$ = !DoPriCor
         Call REPLA(Z$, NREMIT$, 387, 16) ' numero de remito
           FEREMIT% = CVINT%(!Fechmov)
         Call REPLA(Z$, MKI$(FEREMIT%), 419, 2)
         Call REPLA(Z$, NUPLANO$(CODIMAT%), 315, 24)
           '
           MONEMA% = !MoneVal
           COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
           PRECIO = !ValoUnit * COEFCON
           PRELI = PRELISTA#(ARCHIPRE$, CODIMAT%)
           If PRELI < PRECIO Then PRELI = PRECIO
           PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
           '
         Call REPLA(Z$, MKS$(PRELI), 421, 4)
         Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
         Call REPLA(Z$, MKS$(PRECIO), 429, 4)
           IMPO# = CDbl(PRECIO) * CAN
         Call REPLA(Z$, MKD$(IMPO#), 433, 8)
         Call REPLA(Z$, MKS$(PORCEIVA(CODIMAT%)), 441, 4)
           NORIT% = !NuOrItem
         Call REPLA(Z$, MKI$(NORIT%), 503, 2)
         Call REPLA(Z$, MKS$(0), 505, 4)
         Call REPLA(Z$, MKS$(0), 509, 4)
         Call REPLA(Z$, TRIM$(Str$(NPEDID&)), 371, 16)
           NUOCOX$ = NUOCOCLI$(NPEDID&, CODIMAT%)
10       Call REPLA(Z$, NUOCOX$, 355, 16)
         '
         Call REGAPP("!CARDEFAC", Z$)
         CAITCAR% = CAITCAR% + 1
         '
       .MoveNext
       Loop
12   'End If
   End With
   CONSITMP.Close
   Set CONSITMP = Nothing
   '
   jj& = 0
   For U& = 1 To ULTREG&("!CARDEFAC")
     Z$ = REGLEIDO$("!CARDEFAC", U&)
     If Abs(CVS(Mid$(Z$, 1, 4))) >= 0.01 Then
       jj& = jj& + 1
       Call GRAREG("!CARDEFAC", Z$, jj&)
     End If
   Next U&
   Call SETULTREG("!CARDEFAC", jj&)
   Call CIERRARCH("!CARDEFAC")
   '
   CONSICLI% = 1
   If jj& < 1 Then
     Call MENSERR(24, "No hay Remitos Pendientes de Facturación\del " + PERIO$ + " para la Cuenta Elegida.")
     CONSICLI% = 0
   End If
   '
End Sub
'
Sub CARDICAMBIO()
   '
   Dim DOCABX$(2048), SALPES#(2048), SALDOL#(2048), REDOCA&(2048)
   '
   NC% = XVALO(Text1)
   Call BLANARCH("!CARDEFAC")
   '
   Screen.MousePointer = 11
   CADOCPE% = 0
   RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
   For KUL& = 1 To Len(RFG$) Step 4
     RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
     XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
     DOCABY$ = Mid$(XXC$, 13, 18)
     SALPESY# = CVD(Mid$(XXC$, 65, 8)) - CVD(Mid$(XXC$, 73, 8))
     SALDOLY# = CVD(Mid$(XXC$, 97, 8)) - CVD(Mid$(XXC$, 105, 8))
     '
     For KKI% = 1 To CADOCPE%
       If DOCABX$(KKI%) = DOCABY$ Then GoTo 20
     Next KKI%
     CADOCPE% = CADOCPE% + 1
     DOCABX$(CADOCPE%) = DOCABY$
     REDOCA&(CADOCPE%) = RECABIN&
     KKI% = CADOCPE%
     '
20   SALPES#(KKI%) = SALPES#(KKI%) + SALPESY#
     SALDOL#(KKI%) = SALDOL#(KKI%) + SALDOLY#
     '
   Next KUL&
   '
   Call BLANARCH("SELDOCAP")
   For KKI% = 1 To CADOCPE%
     If TIFADOC% = 6 Then    ' N.CREDITO
       If SALDOL#(KKI%) <= 0.005 Then
         If SALPES#(KKI%) > 0 Then
           XXC$ = REGLEIDO$("CARBIMON", REDOCA&(KKI%))
           FEMI = CVI(Mid$(XXC$, 1, 2))
           Z$ = Space$(60) + MKS$(0)
           Call REPLA(Z$, DOCABX$(KKI%), 1, 18)
           Call REPLA(Z$, FECHATEX$(FEMI), 19, 8)
           Call REPLA(Z$, "$" + FORMATNUM$(SALPES#(KKI%), "F11.2"), 29, 12)
           Call REPLA(Z$, "u$s" + FORMATNUM$(SALDOL#(KKI%), "F11.2"), 44, 14)
           Call REPLA(Z$, MKS$(REDOCA&(KKI%)), 61, 4)
           Call REGAPP("SELDOCAP", Z$)
         End If
       End If
     ElseIf TIFADOC% = 3 Then  ' N.DEBITO
       If SALDOL#(KKI%) > 0 Then
         If SALPES#(KKI%) <= 0.005 Then
           XXC$ = REGLEIDO$("CARBIMON", REDOCA&(KKI%))
           FEMI = CVI(Mid$(XXC$, 1, 2))
           Z$ = Space$(60) + MKS$(0)
           Call REPLA(Z$, DOCABX$(KKI%), 1, 18)
           Call REPLA(Z$, FECHATEX$(FEMI), 19, 8)
           Call REPLA(Z$, "$" + FORMATNUM$(SALPES#(KKI%), "F11.2"), 29, 12)
           Call REPLA(Z$, "u$s" + FORMATNUM$(SALDOL#(KKI%), "F11.2"), 44, 14)
           Call REPLA(Z$, MKS$(REDOCA&(KKI%)), 61, 4)
           Call REGAPP("SELDOCAP", Z$)
         End If
       End If
     End If
   Next KKI%
   Call CIERRARCH("SELDOCAP")
   '
   Screen.MousePointer = 1
   If ULTREG&("SELDOCAP") < 1 Then
     Exit Sub
   End If
   '
   Call SELECHO("SELDOCAP")
   DOCABY$ = Left$(VALACT1$("SELDOCAP"), 18)
   If TRIM$(DOCABY$) = "" Then Exit Sub
   DOCUDIF = DOCABY$
   BBXX$ = VALACT2$("SELDOCAP")
   '
   MONEMIS = "Pesos"
   IDIFE = XVALO(Mid$(BBXX$, 12, 11))
   If TIFADOC% = 3 Then
     IDIFE = XVALO(Mid$(BBXX$, 29, 11)) * XVALO(TICAMBIS)
   End If
    '
   ZZ$ = REGBLAN$("CARDEFAC")
   Call REPLA(ZZ$, MKS$(1), 1, 4)
   Call REPLA(ZZ$, "Por Diferencia de Cambio", 7, 64)
   Call REPLA(ZZ$, MKS$(IDIFE), 429, 4)
   Call REPLA(ZZ$, MKD$(CDbl(IDIFE)), 433, 8)
   Call REGAPP("!CARDEFAC", ZZ$)
   ZZ$ = REGBLAN$("CARDEFAC")
   Call REPLA(ZZ$, "s/Factura " + DOCABY$, 7, 64)
   Call REGAPP("!CARDEFAC", ZZ$)
   Call CIERRARCH("!CARDEFAC")
   '
End Sub
'
Private Sub Command7_Click()
    Command7.Enabled = False
    Call CIERRARCH("*.*")
    MODX% = ALTERNA%("Listas de Precios\Condiciones de Pago\Alícuotas de I.V.A.")
    If MODX% = 1 Then
         Call CONECRUN("LISPRE02", "Listas de Precios de Venta")
       ElseIf MODX% = 2 Then
         Call CARCONPAGOS
       ElseIf MODX% = 3 Then
         Call CARTASIVA
         Call COMUNI("Re-Inicie Ahora la Aplicación.")
         Call CIERRARCH("*.*")
         Call FINAL("")
    End If
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    '
    MODO% = COMALTER%("Opciones de Salida\\Pantalla\Impresora\Cancelar")
    If MODO% < 1 Or MODO% > 2 Then GoTo 99
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
      CMO$ = "RT"
    End If
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    DES1 = CVDAT(DES%)
    HAS1 = CVDAT(HAS%)
    '
    Screen.MousePointer = 11
    Call APERBASDAT("STOCKS")
    '
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND ((DoSecLar IS NULL) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F%')) "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= '" & DES1 & "' "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= '" & HAS1 & "' "
    '
    Dim CONSITMP As ADODB.Recordset
    Set CONSITMP = New ADODB.Recordset
25  On Error Resume Next
    CONSITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0

    'Set CONSITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    With CONSITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Remitos Pendientes de Facturación.")  'Generados para este Cliente")
        On Error GoTo 0
        Screen.MousePointer = 1
        Command8.Enabled = True
        Exit Sub
      End If
      On Error GoTo 0
      '
      RECONSI$ = REGBLAN("CONSICLI")
      RECONSJ$ = REGBLAN("REPENFAC")
      J& = 0
      Do While .EOF = False
         NCLIE% = !Nume_Clie
         Call REPLA(RECONSI$, MKI$(NCLIE%), 1, 2)
         Call REPLA(RECONSI$, MKI$(CVINT(!Fechmov)), 3, 2)
         Call REPLA(RECONSI$, "", 5, 18)
         Call REPLA(RECONSI$, !DoPriCor, 23, 12)
         Call REPLA(RECONSI$, MKI$(!Cod_Inte), 35, 2)
         Call REPLA(RECONSI$, MKD$(!Cantsalid), 37, 8)
         Call REPLA(RECONSI$, MKD$(!CantIngre), 45, 8)
         Call REPLA(RECONSI$, MKD$(!CantSaldo), 53, 8)
         Call REPLA(RECONSI$, !ReferCor, 61, 48)
         Call REPLA(RECONSI$, String$(12, 0), 109, 12)
         Call REPLA(RECONSI$, MKS$(XVALO(Mid(!DoPriCor, 5))), 109, 4)
         Call REPLA(RECONSI$, MKI$(CVINT(!FeReMovi)), 122, 2)
          HOR = Format(!FeReMovi, "HH:MM:SS")
         Call REPLA(RECONSI$, MKI$(CVINT(HOR)), 124, 2)
         Call REPLA(RECONSI$, MKI$(!NumUsuar), 127, 2)
         '
         Call REPLA(RECONSJ$, MKI$(NCLIE%), 1, 2)
         Call REPLA(RECONSJ$, MKI$(!Cod_Inte), 3, 2)
         Call REPLA(RECONSJ$, !DoPriCor, 5, 12)
         Call REPLA(RECONSJ$, !DoPriLar, 17, 18)
         Call REPLA(RECONSJ$, MKI$(CVINT(!FeReMovi)), 35, 2)
         If IsNull(!NUPEDCLI) Then
            Call REPLA(RECONSJ$, MKS$(0), 37, 4)
         Else
            Call REPLA(RECONSJ$, MKS$(!NUPEDCLI), 37, 4)
         End If
         If IsNull(!NUPEDCLI) Then
             NUMEOCXX$ = ""
           Else
             NUMEOCXX$ = NUOCOCLI$(!NUPEDCLI, !Cod_Inte)
         End If
         Call REPLA(RECONSJ$, NUMEOCXX$, 41, 24)
         Call REPLA(RECONSJ$, MKD$(!Cantsalid), 65, 8)
         Call REPLA(RECONSJ$, MKD$(!CantIngre), 73, 8)
         Call REPLA(RECONSJ$, MKD$(!CantSaldo), 81, 8)
         Call REPLA(RECONSJ$, MKD$(!ValoUnit), 89, 8)
           TOTI# = !ValoUnit * (!Cantsalid - !CantIngre)
         Call REPLA(RECONSJ$, MKD$(TOTI#), 97, 8)
         '
         J& = J& + 1
         Call GRAREG("CONSICLI", RECONSI$, J&)
         Call GRAREG("REPENFAC", RECONSJ$, J&)
      .MoveNext
      Loop
    End With
    CONSITMP.Close
    Set CONSITMP = Nothing
    '
    Call SETULTREG("CONSICLI", J&)
    Call FILESORT("CONSICLI", "", 1, 6, "ASC")
    Call CIERRARCH("CONSICLI")
    '
    Call SETULTREG("REPENFAC", J&)
    Call FILESORT("REPENFAC", "", 1, 6, "ASC")
    Call CIERRARCH("REPENFAC")
    '
    If MODO% = 1 Then
        Call CONECRUN("*CREPENFA", "Remitos Sin Facturar (Clientes)")
      Else
        RESP% = COMALTER%("Opciones de Agrupamiento\\Comprobante\Articulo\Cancelar")
        If RESP% = 1 Then
            Call CONECRUN("*LIREPEFA", "Remitos Sin Facturar (Clientes)")
          ElseIf RESP% = 2 Then
            Call CONECRUN("*LREPENFA", "Remitos Sin Facturar (Clientes)")
        End If
    End If
    '
99  Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   Call SELECHO("SELFECHA")
   Text9.TEXT = VALACT1$("SELFECHA")
   FF% = FECHANUM(Text9.TEXT)
End Sub


Private Sub Form_Load()
'    '
    Call CENTRAFORM(Me)
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$: DATEDOC% = HOII%
    '
    'Call FORFASQL.CARGAPUERTOS
    '
    MONEMI% = 1  ' PESOS
    Call LEETASIVA
    '
    Text111 = CONTROL("", "FACPRO-A")
    Text7 = CONTROL("", "FAPRMACO")
    
    Screen.MousePointer = 1
    '
End Sub
'

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call BAJALDISCO
   End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
'   Call CIERRARCH("*.*")
'   Hide
End Sub

Private Sub Label19_DblClick()
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
         
         '
       End If
     End If
   End If

End Sub

Private Sub Label26_DblClick()
   Call SELECHO("INDICLIE")
   CPROX$ = TRIM$(VALACT1$("INDICLIE"))
   If CPROX$ <> "" Then
     NCXI% = NUMECLI(CPROX$)
     If NCXI% > 0 Then
       If REGICLI&(NCXI%) > 0 Then
         Label26 = CPROX$
         NC% = NCXI%
         Text1.TEXT = TRIM$(Str$(NC%))
         Call CARDACLI
       End If
     End If
   End If
End Sub

Private Sub Label6_Click()
   If ULTREG&("!CARDEFAC") > 0 Then Exit Sub
   'If TIFADOC% = 2 Or TIFADOC% = 6 Then
      If MODODOC% = 0 Then
          Label6.Caption = "(T.Libre)": MODODOC% = 1
        ElseIf MODODOC% = 1 Then
          Label6.Caption = "(Mat.Codif)": MODODOC% = 0
      End If
   'End If
   '
   NCLIE% = XVALO(Text1.TEXT)
   If NCLIE% > 0 Then
     If MODODOC% > 0 Then
         PORDESCU.TEXT = ""
       Else
         PORDESCU.TEXT = DESCUCLI$(NCLIE%)
     End If
   End If
End Sub

Private Sub Label7_Click()
'   If ULTREG&("!CARDEFAC") > 0 Then Exit Sub
'   If Label7.Caption = "Factura:" Then
'       Label7.Caption = "N.Débito:": TIFADOC% = 3
'       Label6.Caption = "(T.Libre)": MODODOC% = 1
'     ElseIf Label7.Caption = "N.Débito:" Then
'       Label7.Caption = "N.Crédito:": TIFADOC% = 6
'       Label6.Caption = "(Mat.Codif)": MODODOC% = 0
'     Else
'       Label7.Caption = "Factura:": TIFADOC% = 2
'       Label6.Caption = "(Mat.Codif)": MODODOC% = 0
'   End If
   
   '
   NCLIE% = XVALO(Text1.TEXT)
'   If NCLIE% > 0 Then
'     If MODODOC% > 0 Then
'         PORDESCU.TEXT = ""
'       Else
'         PORDESCU.TEXT = DESCUCLI$(NCLIE%)
'     End If
'   End If
End Sub

Private Sub Label8_Click()
   On Error Resume Next
   Text13.SetFocus
   On Error GoTo 0
End Sub

Private Sub List3_DblClick()
   
End Sub

Private Sub MONEMIS_Change()
   
   MONEMA% = MONEMI%
   If Left$(MONEMIS, 1) = "D" Then
        MONEMI% = 2
        MONEMIS.ForeColor = RGB(150, 0, 0)
'        MONEMIS.ToolTipText = "Doble-Click para 'Pesos'"
        'Label12.Visible = True
'        If TICAMBIS.Visible = True Then TICAMBIS.Enabled = True
        TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
        TICAMBX# = XVALO(TICAMBIS)
      Else
        MONEMI% = 1
        MONEMIS.ForeColor = RGB(0, 100, 0)
'        MONEMIS.ToolTipText = "Doble-Click para 'Dólares'"
        'Label12.Visible = False
'        If TICAMBIS.Visible = True Then TICAMBIS.Enabled = False
        TICAMBX# = 1
   End If
   '
'   COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
'   FIN& = ULTREG&("!CARDEFAC")
'   For U& = 1 To FIN&
'     x$ = REGLEIDO$("!CARDEFAC", U&)
'     CAN = CVS(Mid$(x$, 1, 4))
'     PRELI = COEFCON * CVS(Mid$(x$, 421, 4))
'     PRECIO = COEFCON * CVS(Mid$(x$, 429, 4))
'     IMPO# = CDbl(CAN * PRECIO)
'     '
'     Call REPLA(x$, MKS$(PRELI), 421, 4)
'     Call REPLA(x$, MKS$(PRECIO), 429, 4)
'     Call REPLA(x$, MKD$(IMPO#), 433, 8)
'     Call GRAREG("!CARDEFAC", x$, U&)
'   Next U&
   '
'   Call CARLISFAC
'   Call CIERRARCH("!CARDEFAC")
'   Call CALTOFAC
   '
End Sub

Private Sub MONEMIS_DblClick()
'   If TRIM$(UCase$(MONEMIS)) = "PESOS" Then
'       MONEMIS = "Dólares"
'     Else
'       MONEMIS = "Pesos"
'   End If
End Sub

Private Sub PORDESCU_Change()
   Call CALTOFAC
End Sub

Private Sub PORDESCU_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
End Sub



Private Sub Text1_Change()
   Call BLANFORMU
End Sub

Private Sub Text1_DblClick()
  Call SELECHO("DEUDOR1")
  NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
  If NCLIE% > 0 Then
    Text1.TEXT = TRIM$(Str$(NCLIE%))
    If XVALO(Text1.TEXT) > 0 Then
      If XVALO(Text1.TEXT) <= 32767 Then
        NC% = XVALO(Text1)
        If REGICLI&(NC%) > 0 Then
          Call CARDACLI
        End If
      End If
    End If
  End If
End Sub

Private Sub Text1_GotFocus()
   Text1.TEXT = TRIM$(Text1.TEXT)
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.TEXT)

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
            On Error Resume Next
            Command6.SetFocus
            On Error GoTo 0
          End If
        End If
      End If
    End If
  End If
End Sub

Sub CARLETDOC()
   '
   Label13.Caption = "A"
   VARIDOC% = 1
   RNIDOC% = 0
   '
   If PIVADOC% = 0 Or PIVADOC% = 3 Or PIVADOC% = 4 Or PIVADOC% = 6 Then
        Label13.Caption = "B"
        VARIDOC% = 2
        '
      ElseIf PIVADOC% = 2 Then
        RNIDOC% = 1
        '
      ElseIf PIVADOC% = 5 Or PIVADOC% = 7 Then
        Label13.Caption = "X"
        VARIDOC% = 3
        If TRIM$(LETREXPO$) <> "" Then
          Label13.Caption = LETREXPO$
        End If
        If LETREXPO$ = "A" Then
          VARIDOC% = 1
          Label13.Caption = "A"
        End If
    End If
    '
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
   If XVALO(Text11.TEXT) > 9 Then
      On Error Resume Next
      Text10.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text11_LostFocus()
    '
    On Error Resume Next
    Do While Len(TRIM$(Text11)) < 8
     Text11 = "0" + Text11
    Loop
    On Error GoTo 0
    '
End Sub

Private Sub Text12_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text13_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
   End If
End Sub



Private Sub Text16_Change()
   Label16 = Left$(ECOARCH("ECUECON", Text16), 20)
End Sub

''///****NUEVO***//////////
Private Sub TEXT17_CHANGE()
   Call CALTOFA2
End Sub
'
Private Sub Text17_GotFocus()
   Text17.SelStart = 0
   Text17.SelLength = Len(Text17)
End Sub

Private Sub Text17_LostFocus()
    '
    Text17 = FORMATNUM$(XVALO(Text17), "F11.2")
    '
End Sub

Private Sub Text2_gotfocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub



Private Sub TEXT21_Change(Index As Integer)
    If Index = 0 Then
        Eco21(Index).TEXT = ECOARCH$("VENDEDOR", Chr$(XVALO(TEXT21(Index).TEXT)))
      ElseIf Index = 1 Then
        Eco21(Index).TEXT = ECOARCH$("CONPAG", MKI$(XVALO(TEXT21(Index).TEXT)))
    End If
End Sub

Private Sub TEXT211_Change()
    NP& = XVALO(TEXT211)
    If NP& > 0 Then
      TEXT21(4) = NOCOMCLI$(XVALO(TEXT211))
      Call CARGAR_DATOS_PEDIDO(NP&)
    Else
      TEXT21(4) = ""
    End If
End Sub
Sub CARGAR_DATOS_PEDIDO(NP&)
   
   TEXT21(2) = FORMATNUM$(IMPORTEPEDI(NP&), "F12.2")
   NROCOT& = PED_NROCOTIZACION&(NP&)
   Total(2) = FORMATNUM$(ANTICIPO_COTIZACION(NP&, NROCOT&), "F12.2")
   MONEMI% = MONEPEDI%(NP&)
   MONEMIS = UCase(DEMONECO$(MONEMI%))
End Sub

Private Sub TEXT211_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      NPF& = PedidoTeineProforma&(XVALO(TEXT211))
      If NPF& > 0 Then
      
         OPX% = COMALTER%("Atención !!! El Pedido Seleccionado Yá Generó Una Factura Proforma (N° " & NPF& & ")\\Cancelar\Continuar")
         If OPX% <= 1 Then Call BLANFORMU
         Total(2) = ""

      End If
   End If
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
'   Text6.TEXT = ""
'   Text11.TEXT = ""
   DescriTxt.TEXT = ""
   '
   For J = 2 To 6
    Total(J) = ""
   Next J
   '
   For J = 0 To 2
    BASIVA(J) = ""
   Next J
   '
   Text12.TEXT = ""

   Text15 = ""
   Text16 = ""
   Label16 = ""
   Label26 = ""
   Text17 = "" 'LIMPIA RET-IVA
   '
   DOCUDIF = ""
   '
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   Call SETULTREG("!CARDEFAC", 0)
   If EXISTE%(LUCOM$ + "1CARDEFA.WKF") > 0 Then
     Call SETULTREG("!1CARDEFA", 0)
   End If
   TEXT21(4) = ""
   TEXT211 = ""
   TEXT21(0) = ""
   TEXT21(1) = ""
   TEXT21(2) = ""
   '
   Call CIERRARCH("*.*")
   '
End Sub
'
Sub CARDACLI()
    '
    Screen.MousePointer = 11
    '
    NC% = XVALO(Text1.TEXT): Text1.Refresh
    If NC% > 0 Then
      If rasocli$(NC%) <> "" Then
        Text2.TEXT = rasocli$(NC%): Text2.Refresh
        Text3.TEXT = DOMICLI$(NC%): Text3.Refresh
        Text4.TEXT = CPOSCLI$(NC%): Text4.Refresh
        Text5.TEXT = LOCACLI$(NC%): Text5.Refresh
        Label26 = CODICLI$(NC%)
        Text12.TEXT = TRIM$(POSIVAC$(NC%)) + " - CUIT " + CuitCli$(NC%)
          Text12.Refresh
        CREDISPO# = LIMICRE(NC%) - SALDCLI#(NC%) - CHEPEACRE#(NC%)

        NCLIE% = XVALO(Text1.TEXT)
        If NCLIE% > 0 Then
        End If
        PIVADOC% = PIVACLI%(NCLIE%)
        
        TEXT21(1).TEXT = Str$(CPAGCLI%(NCLIE))
        TEXT21(0).TEXT = Str$(VENDCLI%(NCLIE))

        '
        PPVI% = PIVADOC%
        If PPVI% < 0 Or PPVI% > 7 Then PPVI% = 1
        Call CARLETDOC

        Text16 = CONTROL$("SVD451S", "CUEVE-" + TRIM$(Str$(PPVI%)))
        '
        LENTREG$ = Space$(96)
        RFEX$ = RECFILT$("LENTRECL", 1, MKI$(NC%))
        If Len(RFEX$) >= 4 Then
          RELENT& = CVS(Left$(RFEX$, 4))
          If RELENT& > 0 Then
            If RELENT& <= ULTREG&("LENTRECL") Then
              LENTREG$ = Mid$(REGLEIDO$("LENTRECL", RELENT&), 3)
            End If
          End If
        End If
        '
        RNC$ = RECFILT$("NOTACLI", 1, MKI$(NC%))
        URN& = ULTREG&("NOTACLI")

        '
        For U& = 1 To Len(RNC$) Step 4
          RENOTA& = CVS(Mid$(RNC$, U&, 4))
          If RENOTA& > 0 Then
            If RENOTA& <= URN& Then
              TTYY$ = RTrim$(Mid$(REGLEIDO$("NOTACLI", RENOTA&), 3))
              If Text6.TEXT <> "" Then
                Text6.TEXT = Text6.TEXT + Chr$(13) + Chr$(10)
              End If
              Text6.TEXT = Text6.TEXT + RTrim$(TTYY$)
            End If
          End If
        Next U&


        '
        If VERLICRE% = 1 Then
          LMCR# = LIMICRE(NC%)
          SDEU# = SALDCLI#(NC%)
          CHPA# = CHEPEACRE#(NC%)
          If SDEU# + CHPA# >= LMCR# Then
            Call COMUNI("Imposible Facturar a este Cliente.\  \Límite de Crédito Asignado ya está Superado.\  \Informe al Supervisor de Ventas.")
            Call BLANFORMU
            On Error Resume Next
            Text1.SetFocus
            On Error GoTo 0
          End If
        End If
        '
      End If
    End If
    Screen.MousePointer = 1
    '
End Sub
   '

Private Sub Text6_lostFocus()
   On Error Resume Next
   J = Len(TRIM$(Text6))
   Do While Len(TRIM$(Text6)) < 4
    Text6 = "0" + Text6
   Loop
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
     TICACH(MONEMI%) = XVALO(TICAMBIS)
   End If
   TICAMBX# = 1
   If MONEMI% = 2 Then
     TICAMBX# = XVALO(TICAMBIS)
   End If
End Sub

Private Sub TICAMBIS_LostFocus()
   TICAMBIS = FORMATNUM(XVALO(TICAMBIS), "F9.4")
End Sub

Private Sub Timer1_Timer()
   '
   Call CALTOFAC
   '
End Sub

'Sub PRIFACTURA()
'    '
'    CAHOJAS% = 0
'    FORIPRE$ = TRIM$(CONTROL$("", "FORFAC-" + Label13.Caption))
'    If FORIPRE$ = "" Then
'       Call MENSERR(24, "Falta Definir Controlador de Impresión")
'       Exit Sub
'    End If
'    '
'    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
'       Call MENSERR(24, "Imposible Imprimir Documento.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
'       Exit Sub
'    End If
'    '
'    Screen.MousePointer = 11
'    ITPORHOJA% = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
'    If ITPORHOJA% < 1 Then
'      ITPORHOJA% = 32767
'    End If
'    '
'    CANCELPRINT% = 0
'    FECHDOC$ = Text9.TEXT
'    NUMEDOC$ = TRIM$(Text11.TEXT)
'    While Len(NUMEDOC$) < 8
'      NUMEDOC$ = "0" + NUMEDOC$
'    Wend
'    NUMEDOC$ = Label12.Caption + "-" + NUMEDOC$
'    '
'
'    '
''    If InStr(UCase$(TIPODOC$), "FACTURA") > 0 Then'\\\OT-06-0110-WAR-29/03/06///
''      If InStr(UCase$(TIPODOC$), "CONTADO") < 1 Then
''        If CPAG% < 1 Or InStr(UCase$(DECONPA$(CPAG%)), "CONTADO") > 0 Then
''          TIPODOC$ = TIPODOC$ + " CONTADO"
''        End If
''      End If
''    End If
'    '
'    DESTIDOC% = XVALO(Text1.TEXT)
'    '
'    CODTABU1$(1) = "COD-MAT"
'    CODTABU1$(2) = "DES-MAT"
'    CODTABU1$(3) = "CANTIDAD"
'    'CODTABU1$(4) = "PRE-LIST"
'    'CODTABU1$(5) = "PORDESCU"
'    'CODTABU1$(6) = "PRE-NETO"
'    'CODTABU1$(7) = "NET-ITEM"
'    'CODTABU1$(8) = "FECH-REC"
'    CODTABU1$(9) = "ORD-ITEM"
'    CODTABU1$(10) = "UNIMED"
'    CODTABU1$(11) = "REF-MAT1"
'    CODTABU1$(12) = "REF-MAT2"
'    CODTABU1$(13) = "REF-MAT3"
'    CODTABU1$(14) = "N-PLANO"
'    CODTABU1$(15) = "N-ORIGIN"
'    CODTABU1$(16) = "N-OCOMPR"
'    CODTABU1$(17) = "N-REMITO"
'    CODTABU1$(18) = "N-BOLREC"
'    CODTABU1$(19) = "FECH-REC"
'    CODTABU1$(20) = "PRE-LIST"
'    CODTABU1$(21) = "PORDESCU"
'    CODTABU1$(22) = "PRE-NETO"
'    CODTABU1$(23) = "NET-ITEM"
'    CODTABU1$(24) = "ALICUIVA"
'    CODTABU1$(25) = "IVA-ITEM"
'    CODTABU1$(26) = "BRU-ITEM"
'    '
'    CACOLTAB1% = 26
'    '
'    CAITAB1% = 0
'    CAITCRU% = 0
'    '
'    SUMATOTA# = XVALO(Total(0).TEXT)
'    IDESCUEN# = XVALO(Total(1).TEXT)
'    '
'    NroRemi$ = TRIM$(Text10.TEXT)
'    If NroRemi$ <> "" Then
'      While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
'      PREFIREM$ = CONTROL$("", "PREFIREM")
'      If PREFIREM$ = "" Then PREFIREM$ = "0001"
'      NroRemi$ = PREFIREM$ + "-" + NroRemi$
'    End If
'    '
'    If CONSICLI% = 1 Then
'      NroRemi$ = NROREMICO$
'    End If
'    '
'    COEFIVAINCLU# = 1
'    If XVALO(Total(2).TEXT) > 0.005 Then
'      COEFIVAINCLU# = (XVALO(Total(2).TEXT) + XVALO(Total(3).TEXT) + XVALO(Total(4).TEXT) + XVALO(Total(5).TEXT)) / XVALO(Total(2).TEXT)
'    End If
'    '
'    If TRIM$(Text14.TEXT) <> "" Then
'      NUOCOMPRA$ = TRIM$(Text14.TEXT)
'    End If
'    '
'    CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = Label13.Caption
'    CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = NroRemi$
'    CODPARAM$(3) = "NRO-OCOM":  DOCPARAM$(3) = NUOCOMPRA$
'    CODPARAM$(4) = "NRO-AENT":  DOCPARAM$(4) = NROPEDIDO$
'    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = ""
'    'CODPARAM$(6) = "VENDEDOR": DOCPARAM$(6) = Eco21(2).TEXT
'    'CODPARAM$(7) = "COND-PAG": DOCPARAM$(7) = Eco21(1).TEXT
'    'CODPARAM$(8) = "POR-IVA": DOCPARAM$(8) = PORIVA(0).Caption
'    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = ""
'    'CODPARAM$(10) = "POR-IVA1": DOCPARAM$(10) = PORIVA(1).Caption
'    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = ""
'    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = ""
'    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = ""
'    'CODPARAM$(4) = "NRO-OCOM": DOCPARAM$(4) = Text10.TEXT
'    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = ""
'    CODPARAM$(15) = "POR-DESC": DOCPARAM$(15) = TRIM$(PORDESCU.TEXT)
'    CODPARAM$(16) = "IMP-RNI": DOCPARAM$(16) = ""
'    CODPARAM$(17) = "IMP-LET1": DOCPARAM$(17) = ""
'    CODPARAM$(18) = "IMP-LET2": DOCPARAM$(18) = ""
'    CODPARAM$(19) = "IMP-LET3": DOCPARAM$(19) = ""
'    CODPARAM$(23) = "TOTA-DOL": DOCPARAM$(23) = ""
'    '
'   ' CODPARAM$(20) = "DOMI-TRA": DOCPARAM$(20) = LENTREGA(0).TEXT
'   ' CODPARAM$(21) = "DESTINO": DOCPARAM$(21) = LENTREGA(1).TEXT
'   ' CODPARAM$(22) = "TRANSPOR": DOCPARAM$(22) = LENTREGA(2).TEXT
'    CODPARAM$(24) = "MONEDEMI": DOCPARAM$(24) = MONEMIS
'    CODPARAM$(25) = "PERC-BRU": DOCPARAM$(25) = ""
'    CODPARAM$(26) = "SUB-TOTA": DOCPARAM$(26) = ""
'    CODPARAM$(27) = "PERC-IVA": DOCPARAM$(27) = "" 'PERCIVA
'    CAPARDOC% = 27
'    '
'    For KKIL& = 0 To 5
'       Call GRAREG("!OBSERVOF", FINFACTURA.Text1(KKIL&).TEXT, KKIL& + 1)
'    Next KKIL&
'    Call SETULTREG("!OBSERVOF", 6)
'    Call CIERRARCH("!OBSERVOF")
'    '
'    PRINTERPORT$ = PORTUFAC$
'    CACOPDOC% = XVALO(FINFACTURA.Text2(0).TEXT)
'    '
'    NOCONFIRM% = 0
'    If TRIM$(UCase$(CONTROL$("FACTUREM", "IMPRENUM"))) <> "AUTONUME" Then
'      If CONTROL$("FACTUREM", "FORMULAR") <> "HOJASUEL" Then
'        NOCONFIRM% = 1
'      End If
'    End If
'    '
'    HOJASUEL% = 0
'    If CONTROL$("FACTUREM", "FORMULAR") = "HOJASUEL" Then
'       HOJASUEL% = 1
'    End If
'    '
'    FIN& = ULTREG&("!CARDEFAC")
'    SUBTOTA# = 0
'    For U& = 1 To FIN&
'      If (U& Mod ITPORHOJA%) = 1 Then
'        ' PONE ENCABEZADO
'        CAITAB1% = 0
'        CAHOJAS% = CAHOJAS% + 1
'      End If
'      '
'      XXX$ = REGLEIDO$("!CARDEFAC", U&)
'      CIPRE% = CVI(Mid$(XXX$, 5, 2))
'      CAPEX$ = Str$(CVS(Left$(XXX$, 4)))       ' cantidad
'      '
'      CAITAB1% = CAITAB1% + 1
'      TETABU1$(1, CAITAB1%) = CODEXT$(CIPRE%)
'      TETABU1$(2, CAITAB1%) = Mid$(XXX$, 7, 64)
'      TETABU1$(3, CAITAB1%) = CAPEX$
'      'TETABU1$(4, CAITAB1%) = PRELI$
'      'TETABU1$(5, CAITAB1%) = DESCX$
'      'TETABU1$(6, CAITAB1%) = PREUX$
'      'TETABU1$(7, CAITAB1%) = ITOTX$
'      'TETABU1$(8, CAITAB1%) = FENTX$
'      TETABU1$(9, CAITAB1%) = Str$(U&)
'      TETABU1$(10, CAITAB1%) = Mid$(XXX$, 71, 4)                           ' unidad de medida
'      TETABU1$(11, CAITAB1%) = Mid$(XXX$, 75, 64)                ' referencia 1
'      TETABU1$(12, CAITAB1%) = Mid$(XXX$, 139, 64)                ' referencia 1
'      TETABU1$(13, CAITAB1%) = Mid$(XXX$, 203, 64)                ' referencia 1
'      TETABU1$(14, CAITAB1%) = Mid$(XXX$, 315, 24)  ' numero de plano
'      TETABU1$(15, CAITAB1%) = Mid$(XXX$, 339, 16)  ' numero original
'      TETABU1$(16, CAITAB1%) = Mid$(XXX$, 355, 16)  ' numero de o.compra
'      TETABU1$(17, CAITAB1%) = Mid$(XXX$, 387, 16)  ' numero de remito
'      TETABU1$(18, CAITAB1%) = Mid$(XXX$, 403, 16)  ' numero de doc-recepcion
'      TETABU1$(19, CAITAB1%) = FECHATEX$(CVI(Mid$(XXX$, 419, 2))) ' fecha de recepcion
'      '
'      'FACIVITEM = 1: If Label13.Caption = "B" Then FACIVITEM = 1 + (CVS(Mid$(XXX$, 441, 4))) / 100
'      'LPRE% = XVALO(TEXT21(0).TEXT)
'      'If LPRE% > 0 Then
'      '  TETABU1$(20, CAITAB1%) = Str$(FACIVITEM * CVS(Mid$(XXX$, 421, 4))) ' precio de lista
'      '  TETABU1$(21, CAITAB1%) = Str$(CVS(Mid$(XXX$, 425, 4))) ' PORCENTAJE DE DESCUENTO
'      'End If
'      '
'      TETABU1$(22, CAITAB1%) = Str$(FACIVITEM * CVS(Mid$(XXX$, 429, 4))) ' precio unitario
'      TETABU1$(23, CAITAB1%) = Str$(CVD(Mid$(XXX$, 433, 8))) ' total neto
'      TETABU1$(24, CAITAB1%) = Str$(CVS(Mid$(XXX$, 441, 4))) ' alicuota iva
'      TETABU1$(25, CAITAB1%) = Str$(CVD(Mid$(XXX$, 445, 8))) ' importe iva del item
'      TETABU1$(26, CAITAB1%) = Str$(CVD(Mid$(XXX$, 453, 8))) ' total bruto
'      '
'      SUBTOTA# = SUBTOTA# + XVALO(TETABU1$(23, CAITAB1%))
'      '
'      If (U& Mod ITPORHOJA%) = 0 Then
'        If U& < FIN& Then
'          If CACOPDOC% > 0 Then
'            DOCPARAM$(26) = TRIM$(CSTRING$(MKD$(SUBTOTA#), 4, 12, 2, 2))
'            Call PRINTDOC("@" + FORIPRE$)
'          End If
'        End If
'      End If
'109 Next U&
'    '
'    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = Total(0).TEXT
'    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = Total(3).TEXT
'    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = Total(4).TEXT
'    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = Total(1).TEXT
'    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = Total(2).TEXT
'    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = Total(6).TEXT
'    CODPARAM$(16) = "IMP-RNI": DOCPARAM$(16) = Total(5).TEXT
'    If TICAMBIO > 0 Then
'       VADOLA# = XVALO(Total(6).TEXT) / TICAMBIO
'       DOCPARAM$(23) = TRIM$(CSTRING$(MKD$(VADOLA#), 4, 12, 2, 2))
'    End If
'    CODPARAM$(25) = "PERC-BRU": DOCPARAM$(25) = Text15
'    DOCPARAM$(26) = ""
'    DOCPARAM$(27) = Text17                          'PARA IMPRIMIR PERC-IVA
'    '
'    If Label13.Caption = "B" Then
'      DOCPARAM$(5) = Str$(COEFIVAINCLU# * XVALO(Total(0).TEXT))
'      DOCPARAM$(12) = Str$(COEFIVAINCLU# * XVALO(Total(1).TEXT))
'      DOCPARAM$(13) = Str$(COEFIVAINCLU# * XVALO(Total(2).TEXT))
'    End If
'    '
'    CODFOR$ = FORIPRE$
'    LOIMPLET% = XVALO(ATRIFORM$(CODFOR$, "LOIMPLET"))
'    If LOIMPLET% < 16 Then LOIMPLET% = 32
'    TTTT# = XVALO(Total(6).TEXT)
'    TOTFAC$ = NUMTEX$(TTTT#)
'    If MONECO$ <> "" Then
'      PPXX% = InStr(TOTFAC$, "PESOS")
'      If PPXX% > 0 Then
'        TOTFAC$ = Left$(TOTFAC$, PPXX% - 1) + UCase$(MONECO$) + Mid$(TOTFAC$, PPXX% + 5)
'      End If
'    End If
'    Call FRATEX(TOTFAC$, NRX1$, NRX2$, NRX3$, LOIMPLET%)
'    CODPARAM$(17) = "IMP-LET1": DOCPARAM$(17) = NRX1$
'    CODPARAM$(18) = "IMP-LET2": DOCPARAM$(18) = NRX2$
'    CODPARAM$(19) = "IMP-LET3": DOCPARAM$(19) = NRX3$
'    '
'    If CACOPDOC% > 0 Then
'      Call PRINTDOC("@" + FORIPRE$)
'    End If
'    '
'    CACOPDOC% = XVALO(FINFACTURA.Text2(1).TEXT)
'    If CACOPDOC% > 0 Then
'      If TRIM$(Text10.TEXT) <> "" Then
'         For KKI% = 0 To CACOLTAB1%
'           For KKJ% = 0 To CAITAB1%
'             TETABU1$(KKI%, KKJ%) = ""
'           Next KKJ%
'         Next KKI%
'         '
'        Call PRIREMIFA
'      End If
'    End If
'    '
'    Call CIERRARCH("*.*")
'    Call BAJALDISCO
'    '
'    If NOCONFIRM% > 0 Then
'       Call COMUNI("Se han Emitido los Documentos Solicitados")
'    End If
'    '
'End Sub

'Sub PRIREMIFA()
'    '
'    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
'    If TRIM$(FORIPRE$) <> "" Then
'       '
'       For KKIL& = 0 To 2
'          Call GRAREG("!OBSERVOF", FINFACTURA.Text1(KKIL& + 6).TEXT, KKIL& + 1)
'       Next KKIL&
'       Call SETULTREG("!OBSERVOF", 3)
'       Call CIERRARCH("!OBSERVOF")
'       '
'       PREFIREM$ = CONTROL$("", "PREFIREM")
'       If PREFIREM$ = "" Then PREFIREM$ = "0001"
'       NroRemi$ = TRIM$(Text10.TEXT)
'       While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
'       NroRemi$ = PREFIREM$ + "-" + NroRemi$
'       '
'       NC% = XVALO(Text1.TEXT)
'       FECHDOC$ = Text9.TEXT
'       NUMEDOC$ = NroRemi$
'       TIPODOC$ = "Remito"
'       '
'       If TRIM$(Text14.TEXT) <> "" Then
'         NUOCOMPRA$ = TRIM$(Text14.TEXT)
'       End If
'       '
'       CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(TRXP$, 1, 32))
'       CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = TRIM$(FINFACTURA.Text2(2).TEXT)
'       CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(TRXP$, 65, 16))
'       CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = TRIM$(FINFACTURA.Text2(3).TEXT)
'       CODPARAM$(5) = "COND-TRA": DOCPARAM$(5) = TRIM$(Mid$(TRXP$, 97, 32))
'       CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(TRXP$, 129, 24))
'       CODPARAM$(7) = "DEPOSAL ": 'DOCPARAM$(7) = TRIM$(DEPOSAL$)
'       CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = TRIM$(NroRemi$)
'       CODPARAM$(9) = "NRO-OCOM":  DOCPARAM$(9) = NUOCOMPRA$
'       CODPARAM$(10) = "PATE-TRA":  DOCPARAM$(10) = TRIM$(Mid$(TRXP$, 81, 16))
'       'CODPARAM$(11) = "DOMI-TRA":  DOCPARAM$(11) = LENTREGA(0).TEXT
'       CODPARAM$(12) = "LOCA-TRA":  DOCPARAM$(12) = TRIM$(Mid$(TRXP$, 159, 32))
'       'CODPARAM$(13) = "COND-PAG": DOCPARAM$(13) = Eco21(1).TEXT
'       CODPARAM$(14) = "NRO-AENT":  DOCPARAM$(14) = NROPEDIDO$
'       'CODPARAM$(15) = "DESTINO": DOCPARAM$(15) = LENTREGA(1).TEXT
'       'CODPARAM$(16) = "TRANSPOR": DOCPARAM$(16) = LENTREGA(2).TEXT
'       CAPARDOC% = 16
'       '
'       CODTABU1$(1) = "COD-MAT"
'       CODTABU1$(2) = "DES-MAT"
'       CODTABU1$(3) = "CANTIDAD"
'       CODTABU1$(4) = "ORD-ITEM"
'       CACOLTAB1% = 4
'       '
'       DESTIDOC% = NC%
'       CAITAB1% = 0
'       '
'       If FINFACTURA.Text3 <> "" Then
'         Call FRATEXTO(FINFACTURA.Text3, 40)
'         For KKII% = 1 To CARETEX%
'           CAITAB1% = CAITAB1% + 1
'           TETABU1$(1, CAITAB1%) = ""
'           TETABU1$(2, CAITAB1%) = RETEX$(KKII%)
'           TETABU1$(3, CAITAB1%) = ""
'           TETABU1$(4, CAITAB1%) = Str$(CAITAB1%)
'         Next KKII%
'         FINFACTURA.Text3 = ""
'         FINFACTURA.Frame5.Visible = False
'         GoTo 80
'       End If
'       '
'       For U& = 1 To ULTREG&("!CARDEFAC")
'          XXX$ = REGLEIDO$("!CARDEFAC", U&)
'          CANS$ = Str$(CVS(Left$(XXX$, 4)))       ' cantidad
'          If XVALO(CANS$) >= 0.05 Then
'             If TRIM$(Mid$(XXX$, 387, 16)) = "" Then
'                CAITAB1% = CAITAB1% + 1
'                CDXXS$ = CODEXT$(CVI(Mid$(XXX$, 5, 2)))     ' codigo
'                DDTXT$ = Mid$(XXX$, 7, 64)                  ' descripcion
'                '
'                TETABU1$(1, CAITAB1%) = CDXXS$
'                TETABU1$(2, CAITAB1%) = DDTXT$
'                TETABU1$(3, CAITAB1%) = CANS$
'                TETABU1$(4, CAITAB1%) = Str$(CAITAB1%)
'                '
'             End If
'          End If
'       Next U&
'       '
'80     CACOPDOC% = XVALO(FINFACTURA.Text2(1).TEXT)
'       PRINTERPORT$ = PORTUREM$
'       Call PRINTDOC("FOREMIFA")   ' Remito de Facturacion
'       '
'    End If
'    '
'End Sub

Private Sub Total_Change(Index As Integer)
   '
   If Index <= 2 Then
     Call CALTOFAC
   End If
   Call CALTOFA2
   '
End Sub

'
Private Sub Total_GotFocus(Index As Integer)
   '
   If VALICLIE% < 1 Then GoTo 99
   If Index < 3 Or Index > 4 Then GoTo 98
   If TIFADOC% < 3 Then GoTo 98
   Exit Sub
   '
98 On Error Resume Next
   PORDESCU.SetFocus
   On Error GoTo 0
   Exit Sub
   '
99 On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

'Sub GRAFACTURA(OKX%)
'   '
'   OKX% = 0
'   NCLIE% = XVALO(Text1.TEXT)
'   If MARCONOTA.Visible = True Then Exit Sub
'   If ULTREG&("!CARDEFAC") < 1 Then Exit Sub
'   '
'   Call APROCOPIAS(OKX%)
'   COPIASFAC% = XVALO(FINFACTURA.Text2(0))
'   COPIASREM% = XVALO(FINFACTURA.Text2(1))
'   If OKX% < 0 Then Exit Sub
'   If COPIASFAC% + COPIASREM% < 1 Then Exit Sub
'   '
'   If COPIASFAC% > 0 Then
'     If XVALO(Total(6).TEXT) < 0.01 Then
'       Call MENSERR(24, "Imposible Emitir Factura SIN CARGO !!!")
'       Exit Sub
'     End If
'   End If
'   '
'   If COPIASFAC% < 1 Then
'      Call GRAREMITO(OKX%)
'      Screen.MousePointer = 1
'      Exit Sub
'   End If
'   '
'   PPVI% = PIVACLI%(NCLIE%)
'   If PPVI% < 0 Or PPVI% > 7 Then PPVI% = 1
'   CUEVE$ = CONTROL$("SVD451S", "CUEVE-" + TRIM$(Str$(PPVI%)))
'   QQQ% = CUECOINT%(CUEVE$)
'   '
'   If TIFADOC% = 3 Or TIFADOC% = 6 Then
'     '
'     OPIMPUTA.Show 1
'     VDEF$ = "ND": If TIFADOC% = 6 Then VDEF$ = "NC"
'     VDEF$ = VDEF$ + String$(94, 0)
'     NRO& = XVALO(Text11)
'     Call REPLA(VDEF$, MKS$(NRO&), 3, 4)
'     Call REPLA(VDEF$, MKI$(NCLIE%), 7, 2)
'     FFI% = FECHANUM(Text9)
'     Call REPLA(VDEF$, MKI$(FFI%), 9, 2)
'     Call REPLA(VDEF$, MKD$(XVALO(Total(2)) * TICAMBX#), 19, 8)
'     '
'     CUEVE$ = CONTROL$("SVD451S", "CUEVE-8")
'     If OPIMPUTA.Option2.Value = True Then CUEVE$ = CONTROL$("SVD451S", "CUEVE-9")
'     If OPIMPUTA.Option3.Value = True Then CUEVE$ = CONTROL$("SVD451S", "CUEVE-11")
'     QQQ% = CUECOINT%(CUEVE$)
'     Call REPLA(VDEF$, MKI$(QQQ%), 17, 2)
'     XD1$ = VDEF$
'     '
'   End If
'   '
'   NUPANTALLA& = XVALO(Text11.TEXT)
'
'   If XVALO(Text11.TEXT) > NUPANTALLA& Then
'      Call CIERRARCH("SVD202")
'      Call CIERRARCH("DEUDOR2")
'      Call CIERRARCH("FACTUR")
'      Call CIERRARCH("CAJABANC")
'      Call CIERRARCH("MOVISTO")
'      Call CIERRARCH("SALSTOCK")
'      Call COMUNI("FLEXOFT ha Re-Numerado el Documento\como Número '" + TRIM$(Text11.TEXT) + "'.\  \Verifique y Vuelva a Grabar.")
'      Exit Sub
'   End If
'   Text11.TEXT = TRIM$(Str$(NUPANTALLA&))
'   '
'   COPIASREM% = XVALO(FINFACTURA.Text2(1))
'   If COPIASREM% < 1 Then
'      Text10.TEXT = ""
'   End If
'   '
'   Screen.MousePointer = 11
'   If VALICLIE% < 1 Then
'      MERRO$ = "Número de Cliente no Válido"
'      GoTo 98
'   End If
'   NCLIE% = XVALO(Text1.TEXT)
'   '
'   If XVALO(Total(6).TEXT) * TICAMBX# < 0.01 Then
'      MERRO$ = "Importe no Válido"
'      GoTo 98
'   End If
'   '
'   HOII% = HOY(HOS$): FEACT = HOII%
'   FEMI% = FECHANUM(Text9.TEXT)
'   ' LA SIGUIENTE VALIDACION SE SUPRIME A PEDIDO DE S.IGLESIAS DE I.GUIDI
'   ' If FEMI% < DIANTE(FEACT) Or FEMI% > DIPOST(FEACT) Then
'   '    MERRO$ = "Fecha de Emisión no Válida"
'   '    GoTo 98
'   ' End If
'   '
'   If VERLICRE% = 1 Then
'     If TIFADOC% < 4 Then ' EXCLUYE CREDITOS
'       '
'       For UKK& = 1 To ULTREG&("!CARDEFAC")
'         XKK$ = REGLEIDO$("!CARDEFAC", UKK&)
'         NROPEDXX& = XVALO(TRIM$(Mid$(XKK$, 371, 16)))
'         If NROPEDXX& < 1 Then GoTo 22
'         RFFKK$ = RECFILT$("AUTORENT", 1, MKS$(NROPEDXX&))
'         If Len(RFFKK$) < 4 Then GoTo 22 ' NO AUTORIZADO
'         REGAUTKK& = CVS(RFFKK$)
'         XAUT$ = REGLEIDO$("AUTORENT", REGAUTKK&)
'         If UCase$(Mid$(XAUT$, 31, 2)) <> "SI" Then GoTo 22
'       Next UKK&
'       GoTo 25
'       '
'22     LMCR# = LIMICRE(NCLIE%)
'       SDEU# = SALDCLI#(NCLIE%)
'       CHPA# = CHEPEACRE#(NCLIE%)
'       TPED# = XVALO(Total(6).TEXT) * TICAMBX#
'       If SDEU# + CHPA# + TPED# > LMCR# Then
'         TLIMI$ = TRIM$(CSTRING$(MKD$(LMCR#), 3, 12, 2, 2))
'         If TLIMI$ = "" Then TLIMI$ = "0.00"
'         Call COMUNI("Imposible Emitir Este Documento.\  \Supera el Límite de Crédito Asignado\para esta Cuenta de $ " + TLIMI$ + ".\  \Informe al Supervisor de Ventas.")
'         Exit Sub
'       End If
'     End If
'   End If
'   '
'25 DESC$ = TRIM$(PORDESCU.TEXT)
'   PODES = COEFDESCU(DESC$)
'   '
'   For KK% = 1 To 10: TOTAFAC#(1, KK%) = 0: Next KK%
'   'TOTAFAC#(1, 1) = XVALO(BASIVA(0).Caption) / (1 - PODES / 100)
'   'TOTAFAC#(1, 2) = XVALO(BASIVA(1).Caption) / (1 - PODES / 100)
'   TOTAFAC#(1, 3) = XVALO(Total(0).TEXT)
'   TOTAFAC#(1, 4) = XVALO(Total(1).TEXT)
'   'TOTAFAC#(1, 5) = XVALO(BASIVA(0).Caption)
'   TOTAFAC#(1, 6) = XVALO(Total(3).TEXT)
'   TOTAFAC#(1, 7) = XVALO(Total(5).TEXT)
'   TOTAFAC#(1, 8) = XVALO(Total(6).TEXT)
'   'TOTAFAC#(1, 9) = XVALO(BASIVA(1).Caption)
'   TOTAFAC#(1, 10) = XVALO(Total(4).TEXT)
'   '
'   TICA = 1
'   For KK% = 1 To 10
'     TOTAFAC#(1, KK%) = TOTAFAC#(1, KK%) * TICAMBX#
'     TOTAFAC#(2, KK%) = TOTAFAC#(1, KK%) / TICA
'   Next KK%
'   '
'   LU$ = LUDAT$
'   ARCHIDEU$ = "SVD202"
'   ARCHICLI$ = "DEUDOR"
'   ARCHIFAC$ = "FACTUR"
'   RD% = 1
'   TIPO% = TIFADOC%
'   VA% = VARIDOC%
'   NRO = XVALO(Text11.TEXT)
'   NC% = NCLIE%
'   FFI% = FECHANUM(Text9.TEXT)
'   TOTAGEN# = XVALO(Total(6).TEXT) * TICAMBX#
'   '
'   'Call BLIMESS("Actualizando Saldo de Cuenta ...")
'   SIGX% = 1: If TIPO% = 6 Then SIGX% = (-1)
'   TTF# = TOTAGEN# * SIGX%
'   Call GRACLI(LU$, ARCHICLI$, NC%, TTF#, DATCLI$(), 2)
'   Call CIERRARCH("DEUDOR2")
'   '
'   FLIM% = FECHANUM("01/01/99")
'   If ULTREG&("FORMFAC") > 0 Then
'      FFS$ = REGLEIDO$("FORMFAC", 2&)
'      FLIM% = CVI(Mid$(FFS$, 19, 2))
'   End If
'   '
'   'Call BLIMESS("Grabando Factura ...")
'   ' LO QUE SIGUE ES PARA CONSIDERAR LAS HOJAS QUE SE CONSUMIRAN
'   FORIPRE$ = TRIM$(CONTROL$("", "FORFAC-" + Label13.Caption))
'   CARENGLO& = ULTREG&("!CARDEFAC")
'   ITPORHOJA% = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
'   If ITPORHOJA% < 1 Then
'     ITPORHOJA% = 80
'   End If
'   CAHOJNEC% = Int(0.99 + CARENGLO& / ITPORHOJA%)
'
'   NPP& = XVALO(Text11) + CAHOJNEC% - 1
'   Text11.TEXT = TRIM$(Str$(NPP&))
'   Text11.Refresh
'   '
'   TICOMX$ = "FC": OPMOVIX$ = "FAC"
'   If TIPO% = 3 Then TICOMX$ = "ND": OPMOVIX$ = "DEB"
'   If TIPO% = 6 Then TICOMX$ = "NC": OPMOVIX$ = "CRE"
'   UREFA& = ULTREG&("FACTUR")
'   NFCZ$ = TRIM$(Text11): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
'   NFCX$ = TICOMX$ + "-" + Mid$(NFCZ$, 3) + "-" + Label13
'   NFCY$ = TICOMX$ + "-" + Label12 + "-" + NFCZ$ + "-" + Label13
'   '
'   OPMOVIX$ = "FC"
'   If TIFADOC% = 3 Then OPMOVIX$ = "ND"
'   If TIFADOC% = 6 Then OPMOVIX$ = "NC"
'   VAMOVIX$ = Label13
'   '
'   SIGX% = 1: If TIPO% = 6 Then SIGX% = (-1)
'   '
'' AQUI GRABAR CAJABANC PARA FACTURA
'   CoComCor$ = NFCX$
'   CoComLar$ = NFCY$
'   CodiEmp% = 0
'   CodiSuc% = 0
'   NumeClie = GRUDIVCLI%(NC%)
'   NumeProv% = 0
'   FechEmi = CVDAT(FFI%)
'   Referen$ = RASOCLI$(NC%)
'   FeContab = CVDAT(FFI%)
'   FeAcredi = CVDAT(FFI%)
'   CuentaCo% = CUECOCLI%(NC%)
'   TiMovi$ = "FVEN"
'   OpMovi$ = OPMOVIX$
'   VaMovi$ = Label13
'   Descrip$ = Label7 + " " + Label13.Caption + " " + Label6.Caption + " " + Label12.Caption + "-" + Text11.TEXT
'   NumeTal = XVALO(Label12)
'   NumeCom = CDbl(Text11)
'   StatInit = 0
'   ImpoDeb = TOTAGEN# * SIGX%
'   Observa$ = RASOCLI$(NC%)
'   Call GRACABAN
'   '
'   BasImpo# = 0
'   '
'   If CONSICLI% < 1 Then
'       Call GRAM2000(NC%, NRO1, REMI1, FFI%)
'     Else
'       Call GRACONSI
'   End If
'   '
'80 OKX% = 1
'   Screen.MousePointer = 1
'   Exit Sub
'
'98 Screen.MousePointer = 1
'   Call MENSERR(24, MERRO$)
'   '
'End Sub
'
Sub GRADEUFA(Tipo%, VA%, NRO, NC%, FFAC%, COPAGI%, TOTFAC#)
   '
   Dim FVI%(6), PLAII%(6), IM#(6), FEX(3)
   '
   LU$ = LUDAT$
   IORIG# = TOTFAC#
   '
   If Tipo% = 0 Then
     If IORIG# < 0 Then
       Tipo% = 7
     End If
   End If
   '
   If Tipo% > 3 Then
     NUPAG% = 1
     FVI%(1) = FFAC%
     IM#(1) = TOTFAC#     ' IORIG#
     GoTo 41
   End If
   '
   ' busca condicion de pago segun COPAGI%
20 Call BUCONPA(LU$, COPAGI%, CoPag$, FREF%, PLA$)
   '
   ' determina vencimientos segun total factura y condicion de pago
30 FF = FFAC%: If FF < 1 Then FF = FECHANUM("01/01/80")
   FFS$ = FECHATEX$(FF)
   PLA$ = Left$(PLA$ + String$(6, 0), 6)

   If (FREF% < 1 Or FREF% > 4) Then
        Call MENSERR(24, "Mal Fecha Referencia en Cond.Pago Nro." + Str$(COPAGI%))
           FREF% = 1
   End If

   If FREF% = 2 Then FREF% = 1   ' AUN NO FUNCIONA SEMANA DE FACTURA
   If FREF% <> 3 Then GoTo 35
   If XVALO(Left$(FFS$, 2)) <= 15 Then FEBAS$ = "15" + Mid$(FFS$, 3)
   If XVALO(Left$(FFS$, 2)) > 15 Then FEBAS$ = "31" + Mid$(FFS$, 3)
35 If FREF% = 1 Then FEBAS$ = FFS$
   If FREF% = 4 Then FEBAS$ = "31" + Mid$(FFS$, 3)
   '
   For i% = 1 To 6
     FVI%(i%) = 0
   Next i%
   '
   FEFA = FECHANUM(FEBAS$)
   For i% = 1 To 6
     PLAI% = Asc(Mid$(PLA$, i%, 1))
     PLAII%(i%) = PLAI%
     If i% > 1 Then
       If (PLAII%(i%) < PLAII%(i% - 1) Or PLAII%(i%) <= 0) Then
         GoTo 31
       End If
     End If
     '
     FEXU = FEFA + PLAI%
     While DIENTRE(FEFA, FEXU) < PLAI%
       FEXU = DIPOST(FEXU)
     Wend
     '
     FVI%(i%) = FEXU
     NUPAG% = i%
   Next i%

31 FVIJ% = FVI%(NUPAG%)
   If NUPAG% < 2 Then FVIJ% = FFAC%
   '
   ' determina importes segun numero de pagos
40 If NUPAG% < 1 Then NUPAG% = 1
   IM#(1) = (Int(100 * IORIG# / NUPAG% + 0.5)) / 100

   For H% = 2 To NUPAG% - 1
     IM#(H%) = IM#(1)
   Next H%
   IM#(NUPAG%) = IORIG#

   For H% = 1 To NUPAG% - 1
     IM#(NUPAG%) = IM#(NUPAG%) - IM#(H%)
   Next H%
   '
   ICANCE# = IORIG# - TOTFAC#
   For H% = 1 To NUPAG%
     If ICANCE# > 0 Then
       If ICANCE# >= IM#(H%) Then
           ICANCE# = ICANCE# - IM#(H%)
           IM#(H%) = 0
         Else
           IM#(H%) = IM#(H%) - ICANCE#
           ICANCE# = 0
       End If
     End If
   Next H%
   '
41 NuorVen% = 0
   SIGX% = 1: If Tipo% > 3 Then SIGX% = (-1)
   For i% = 1 To NUPAG%
     If Abs(IM#(i%)) >= 0.005 Then
       FechEmi = CVDAT(FFAC%)
       IBruTot# = IORIG#
       NuorVen% = NuorVen% + 1
       FeVenc = CVDAT(FVI%(i%))
       ImVenc# = (Abs(IM#(i%))) * SIGX%
       Call GRADEUPEN
     End If
   Next i%
   '
End Sub
'
Sub GRAMOCABE()
   '
   FEMI% = FECHANUM(Text9)
   NCLIE% = XVALO(Text1)
   Tipo% = TIFADOC%
   VA% = VARIDOC%
   NRO& = XVALO(Text11.TEXT)
   NUDOCU$ = FORMDOC$(NRO&, Tipo%, VA%, SUCUVEN%)
   '
   SUCUVEN% = 0
   On Error Resume Next
   SUCUVEN% = XVALO(Label12)
   On Error GoTo 0
   NUDOCU$ = FORMDOC$(NRO&, Tipo%, VA%, SUCUVEN%)
   '
   TOTAPESOS# = TOTAFAC#(1, 8)
   TOTADOLAR# = TOTAPESOS# / TICAMBX#
   '
   RAPLI$ = REGBLAN$("CARBIMON")
   ' documento cabeza
   Call REPLA(RAPLI$, MKI$(FEMI%), 1, 2)
   Call REPLA(RAPLI$, MKI$(NCLIE%), 3, 2)
   Call REPLA(RAPLI$, Chr$(Tipo%), 5, 1)
   Call REPLA(RAPLI$, MKI$(VA%), 6, 1)
   Call REPLA(RAPLI$, MKS$(NRO), 7, 4)
   Call REPLA(RAPLI$, MKI$(SUCUVEN%), 11, 2)
   Call REPLA(RAPLI$, NUDOCU$, 13, 18)
   Call REPLA(RAPLI$, MKI$(MONEMI%), 31, 2)
   '
   ' documento aplicado
   Call REPLA(RAPLI$, MKI$(FEMI%), 33, 2)
   Call REPLA(RAPLI$, MKI$(NCLIE%), 35, 2)
   Call REPLA(RAPLI$, Chr$(Tipo%), 37, 1)
   Call REPLA(RAPLI$, MKI$(VA%), 38, 1)
   Call REPLA(RAPLI$, MKS$(NRO), 39, 4)
   Call REPLA(RAPLI$, MKI$(SUCUVEN%), 43, 2)
   Call REPLA(RAPLI$, NUDOCU$, 45, 18)
   Call REPLA(RAPLI$, MKI$(MONEMI%), 63, 2)
   '
   'importe en pesos
   Call REPLA(RAPLI$, MKD$(TOTAPESOS#), 65, 8)
   Call REPLA(RAPLI$, MKD$(TICAMBX#), 89, 8)
   '
   'importe en dolares
   Call REPLA(RAPLI$, MKD$(TOTADOLAR#), 97, 8)
   '
   Call REGAPP("CARBIMON", RAPLI$)
   Call CIERRARCH("CARBIMON")
   '
End Sub
'


Sub GRAREMITO(OKX%)
   '
   Screen.MousePointer = 11
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLIE% = XVALO(Text1.TEXT)
   '
   Call BLOQARCH("MOVISTO")
   Call BLOQARCH("CONSICLI")
   '
   AINDI% = 0
   CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
   If Len(CMO$) <> 2 Then
     CMO$ = "RT"
   End If
   '
   REMI$ = "RT." + TRIM$(Text10.TEXT)
   NROO& = XVALO(Mid$(REMI$, 4))
   FF% = FECHANUM(Text9.TEXT)
   DEPOSALI% = DEPOEXPE%
   '
   PREFIREM$ = CONTROL$("", "PREFIREM")
   If PREFIREM$ = "" Then PREFIREM$ = "0001"
   NroRemi$ = TRIM$(Mid$(REMI$, 4))
   While Len(NroRemi$) < 8: NroRemi$ = "0" + NroRemi$: Wend
   NroRemi$ = CMO$ + "." + PREFIREM$ + "-" + NroRemi$
   '
   HO% = HOY(HOS$)
   HP = 100 * XVALO(Left$(Time$, 2)) + XVALO(Mid$(Time$, 4, 2))
   User% = USNBR% Mod 100
   CTRXX$ = MKI$(HO%) + MKI$(Int(HP)) + Chr$(User%) + String$(2, 0)
   '
   RECONSI$ = REGBLAN$("CONSICLI")
   Call REPLA(RECONSI$, String$(12, 0), 109, 12)
   Call REPLA(RECONSI$, MKI$(NCLIE%), 1, 2)
   Call REPLA(RECONSI$, MKI$(FF%), 3, 2)
   Call REPLA(RECONSI$, NroRemi$, 5, 18)
   Call REPLA(RECONSI$, String$(12, 0), 109, 12)
   Call REPLA(RECONSI$, CTRXX$, 122, 7)
   '
   REFE$ = rasocli$(NCLIE%)
   '
   YXI% = 0
   FIN& = ULTREG&("!CARDEFAC")
   For i& = 1 To FIN&
     Call TRACE(20, i&, FIN&)
     x$ = REGLEIDO$("!CARDEFAC", i&)
     CANTAA = CVS(Left$(x$, 4))
     If CANTAA > 0 Then
       Ci% = CVI(Mid$(x$, 5, 2))
       If Ci% > 0 Then
         NPEDID& = XVALO(Mid$(x$, 371, 8))
         COD$ = CODEXT$(Ci%)
         PRE# = CDbl(CVS(Mid$(x$, 429, 4)))
         CMOV$ = "RT"
         If Mid$(x$, 461, 1) = "R" Then CMOV$ = "*R" ' REPARACION
         DOPRI$ = REMI$
         DOSEC$ = ""
         LOTE$ = ""
         SIG% = (-1)
         '
         'Call MOVISTO(FF%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, PRE#, "$", NCLIE%, DEPOSALI%, SIG% * CANTAA)
         '
         YXI% = YXI% + 1
         ZZ$ = RECONSI$
         Call REPLA(ZZ$, TRIM$(REMI$) + "-" + TRIM$(Str$(YXI%)), 23, 12)
         Call REPLA(ZZ$, MKI$(Ci%), 35, 2)
         Call REPLA(ZZ$, MKD$(CDbl(Abs(CANTAA))), 37, 8)
         Call REPLA(ZZ$, MKS$(NPEDID&), 109, 4)
         Call REPLA(ZZ$, MKD$(PRE#), 113, 8)
         Call REPLA(ZZ$, Chr$(MONEMI%), 121, 1)
         Call REGAPP("CONSICLI", ZZ$)
         '
       End If
     End If
   Next i&
   '
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("CONSICLI")
   Screen.MousePointer = 1
   OKX% = 1
   Exit Sub
   '
   '
98 Call CIERRARCH("MOVISTO")
   Call CIERRARCH("CONSICLI")
   Screen.MousePointer = 1
   Call MENSERR(24, MERRO$)
   '
End Sub
'
Sub APROCOPIAS(OKX%)
   '\\\OT-06-0114-WAR-03/04/06///
'   TIPCOM$ = TRIM$(REPLACAR$(Label7.Caption, ":", ""))
'   CATIVAS$ = Label13.Caption
'   '
'   TCMP$ = TRIM$(TIPCOM$)
''   If InStr(UCase$(TIPCOM$), "FACTURA") > 0 Then  '\\\OT-06-0110-WAR-29/03/06///
''     If InStr(UCase$(TIPCOM$), "CONTADO") < 1 Then
''       If CPAG% < 1 Or InStr(UCase$(DECONPA$(CPAG%)), "CONTADO") > 0 Then
''         TCMP$ = TCMP$ + " CONTADO"
''       End If
''     End If
''   End If
'   '
'   TICOMPROBAN$ = TCMP$
'   FINFACTURA.Label2.Caption = TCMP$
'   FINFACTURA.Label1(0) = TCMP$ + ":"
'   FINFACTURA.Label4.Caption = CATIVAS$
'   '
'   CODFOR$ = TRIM$(CONTROL$("", "FORFAC-" + Label13.Caption))
'   NROCON& = XVALO(Text11.TEXT)
'   NCN$ = TRIM$(Str$(NROCON&))
'   While Len(NCN$) < 8: NCN$ = "0" + NCN$: Wend
'   PREFI$ = Label12.Caption
'   NCN$ = PREFI$ + "-" + NCN$
'   NUCOMPROBAN$ = NCN$
'   FINFACTURA.Label3.Caption = NCN$
'   '
'5  COPIASFAC% = XVALO(CONTROL$("", "COPIAFAC"))
'   COPIASREM% = XVALO(CONTROL$("", "COPIAREM"))
'   If COPIASFAC% < 1 Then COPIASFAC% = 1
'   If COPIASREM% < 1 Then COPIASREM% = 1
'   '
'   If TIFADOC% > 2 Then
'      'COPIASFAC% = 2: COPIASREM% = 0
'      COPIASREM% = 0   ' POR TEMA CANTIDAD DE COPIAS FRIOLATINA
'   End If
'   '
'   If CONSICLI% = 1 Then
'      COPIASREM% = 0
'   End If
'   '
'   FINFACTURA.Text2(0) = TRIM$(Str$(COPIASFAC%))
'   FINFACTURA.Text2(1) = TRIM$(Str$(COPIASREM%))
'   FINFACTURA.Text2(2) = " "
'   FINFACTURA.Text2(3) = TRIM$(Str$(Int(TICAMBX# * XVALO(Total(2).TEXT)))) + ".="
'   FINFACTURA.Label5.Caption = ""
'   For U& = 0 To 8
'      FINFACTURA.Text1(U&).TEXT = ""
'   Next U&
''AQUI LIMPIAR
'   '
'   EPAC$ = TRIM$(UCase$(EMPREAC$))
'   TTXX$ = ""
'   '
'   FINFACTURA.Label6.Enabled = False
'   FINFACTURA.CLAREDOL = "0"
'   If MONEMI% = 1 Then
'     FINFACTURA.Label6.Enabled = True
'   End If
'   '
'   FINFACTURA.Label6.Visible = False
'   FINFACTURA.Text3 = ""
'   FINFACTURA.Frame5.Visible = False
'   '
'10 FINFACTURA.Show 1
'   Refresh
'   '
'   If FINFACTURA.Label5.Caption <> "APRO" Then
'      If COMALTER%("Realmente Quiere CANCELAR este Documento\\No, Continuar\Si, Cancelar") <> 2 Then GoTo 10
'      OKX% = (-1)
'      Exit Sub
'   End If
'   '
'   COPIASFAC% = XVALO(FINFACTURA.Text2(0))
'   If COPIASFAC% < 0 Or COPIASFAC% > 6 Then GoTo 5
'   COPIASREM% = XVALO(FINFACTURA.Text2(1))
'   If COPIASREM% > 6 Then GoTo 5
'   If COPIASREM% < 1 Then
'      Text10.TEXT = ""
'   End If
'   '
End Sub
'
Private Sub Total_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
     If Index = 0 Then
       On Error Resume Next
       PORDESCU.SetFocus
       On Error GoTo 0
      Else
       On Error Resume Next
       Command14.SetFocus
       On Error GoTo 0
     End If
   End If
End Sub

Private Sub Total_LostFocus(Index As Integer)
   If Index >= 2 Then
     TTT1 = XVALO(Total(Index).TEXT)
     Total(Index) = CSTRING$(MKS$(TTT1), 3, 12, 2, 2)
   End If
End Sub

'
Sub CARLISFAC()
    List1.Clear
   FIN& = ULTREG&("!CARDEFAC")
   For U& = 1 To FIN&
      y$ = REGLEIDO$("!CARDEFAC", U&)
      TTXX$ = Space$(128)
      Ci% = CVI(Mid$(y$, 5, 2))
      PRECIO = CVS(Mid$(y$, 429, 4))
      PRELI = PRELISTA#(ARCHIPRE$, Ci%)
      PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
      MAREPA$ = UCase$(Mid$(y$, 461, 1)): If MAREPA$ <> "R" Then MAREPA$ = " "
      NPEDID& = XVALO(Mid$(y$, 371, 8))
      CAN = CVS(Mid$(y$, 1, 4))
      '
      Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 1, 6)
      Call REPLA(TTXX$, CODEXT$(Ci%), 8, 15)
      Call REPLA(TTXX$, Mid$(y$, 7, 64), 24, 26)
      PRELI = CVS(Mid$(y$, 421, 4))
      Call REPLA(TTXX$, CSTRING$(MKS$(PRELI), 3, 8, 2, 2), 50, 8)
      PORDESCUIN = CVS(Mid$(y$, 425, 4))
      Call REPLA(TTXX$, CSTRING$(MKS$(PORDESCUIN), 3, 5, 1, 2), 59, 5)
      PREUNI = CVS(Mid$(y$, 429, 4))
      Call REPLA(TTXX$, CSTRING$(MKS$(PREUNI), 3, 8, 2, 2), 64, 8)
      TOT# = CVD(Mid$(y$, 433, 8))
      Call REPLA(TTXX$, CSTRING$(MKD$(TOT#), 3, 10, 2, 2), 74, 10)
      Call REPLA(TTXX$, Mid$(y$, 461, 2), 86, 2) ' MARCA REPARACION
      List1.AddItem TTXX$
      '
   Next U&
   List1.Refresh
   '
End Sub
'
Sub CALTOFAC()
   '
   Static CLIA$, RNI%, APLIVA%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE% = XVALO(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIE% > 0 Then
       PPVI% = PIVACLI%(NCLIE%)
       If PPVI% = 2 Then RNI% = 1
       If PPVI% = 5 Or PPVI% = 7 Then APLIVA% = 0
     End If
   End If
   '
   TASA1 = XVALO(PORIVA(0).Caption)
   TASA2 = XVALO(PORIVA(1).Caption)
   TASA3 = XVALO(PORIVA(2).Caption)
   '
   TTTX# = XVALO(Total(2))
   BASIVA(0).Caption = TRIM$(CSTRING$(MKD$(TTTX#), 3, 12, 2, 2))
   Total(3).TEXT = CSTRING$(MKD$(XVALO(BASIVA(0).Caption) * XVALO(PORIVA(0).Caption) * APLIVA% / 100), 3, 12, 2, 2)
   '
   TTTX# = (XVALO(Total(2).TEXT) - XVALO(BASIVA(0).Caption))
   BASIVA(1).Caption = TRIM$(CSTRING$(MKD$(TTTX#), 3, 12, 2, 2))
   Total(4).TEXT = CSTRING$(MKD$(XVALO(BASIVA(1).Caption) * XVALO(PORIVA(1).Caption) * APLIVA% / 100), 3, 12, 2, 2)
   '
   Call CALTOFA2
   '
End Sub
'
Sub CALTOFA2()
   '
   Static CLIA$, RNI%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE% = XVALO(CLIA$)
     RNI% = 0
     If NCLIE% > 0 Then
       PPVI% = PIVACLI%(NCLIE%)
       If PPVI% = 2 Then RNI% = 1
     End If
   End If
   '
   TTTX# = (XVALO(Total(3).TEXT) + XVALO(Total(4).TEXT)) * RNI%
   '
   TTTX# = 0: For KKI% = 2 To 5 'ANTES DE 3 A 5
                 TTTX# = TTTX# + XVALO(Total(KKI%).TEXT)
              Next KKI%
              TTTX# = TTTX# + XVALO(Text15) 'PERCEP IVA
              TTTX# = TTTX# + XVALO(Text17) 'PERCEP I.BRUTOS
   Total(6).TEXT = CSTRING$(MKD$(TTTX#), 3, 12, 2, 2)
   '

End Sub
Sub CALTOFA1()
    '
    
    TTTX# = TTTX# + XVALO(Total(7)) 'TOTAL + NETO
    Total(6).TEXT = CSTRING$(MKD$(TTTX#), 3, 12, 2, 2)
    '
End Sub
Sub LEETASIVA()
   '
   URETA& = ULTREG&("GRUPIVA")
   If URETA& < 2 Then
     Call MENSERR(24, "Falta Definir Grupos de I.V.A.")
     Call CIERRARCH("*.*")
     Call FINAL("")
   End If
   '
   For U& = 1 To 2
      x$ = REGLEIDO$("GRUPIVA", U&)
      TAIVA(U& - 1) = CVS(Mid$(x$, 13, 4))
      PORIVA(U& - 1) = TRIM$(CSTRING$(MKS$(TAIVA(U& - 1)), 3, 6, 2, 2))
   Next U&
   PORIVA(2) = "50.00" ' responsable no inscripto
   Call CIERRARCH("GRUPIVA")
   ALIIVA = ALIVA
   '
End Sub

Sub GRACONSI()
   '
   NUFACLAR$ = CoComLar$
   NUFACORT$ = CoComCor$
   '
   NC1% = XVALO(Text1.TEXT)
   Call APERBASDAT("STOCKS")
   
   FIN& = ULTREG&("!CARDEFAC")
   '
   For i& = 1 To FIN&
     Call TRACE(20, i&, FIN&)
     x$ = REGLEIDO$("!CARDEFAC", i&)
     CANTAA = CVS(Left$(x$, 4))
     If Abs(CANTAA) > 0 Then
       Ci% = CVI(Mid$(x$, 5, 2))
       If Ci% > 0 Then
         DOPRIC$ = TRIM$(Mid$(x$, 387, 16))
         NORIT% = CVI(Mid$(x$, 503, 2))
         PREUNI = CVS(Mid$(x$, 429, 4))
         NPEDCAR& = XVALO(Mid$(x$, 371, 16))
         '
         SQLX$ = "DoPriCor = '" & DOPRIC$ & "' "
         SQLX$ = SQLX$ + "AND Cod_Inte = " & Ci% & " "
         SQLX$ = SQLX$ + "AND Nume_Clie = " & NC1% & " "
         SQLX$ = SQLX$ + "AND NuorItem = " & NORIT% & " "
         Dim MOVVVISTO As ADODB.Recordset
         Set MOVVVISTO = New ADODB.Recordset
20       On Error Resume Next
         MOVVVISTO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
         If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 20
         End If
         On Error GoTo 0
         '
         With MOVVVISTO
           On Error Resume Next
           '.FindFirst SQLX$
25         'If .NoMatch = False Then
           If Not (.EOF And .BOF) Then
             On Error GoTo 0
             Do While (Not .EOF)
               NPEDID& = 0
               If IsNull(!NUPEDCLI) = False Then NPEDID& = !NUPEDCLI
               If NPEDID& = NPEDCAR& Then
    '             .FindNext SQLX$
    '             GoTo 25
                  !DoSecCor = NUFACORT$
                  !DoSecLar = NUFACLAR$
                  !ValoUnit = PREUNI
                  If CANTAA > 0 Then
                      !Cantsalid = CANTAA
                      !CantIngre = 0
                    Else
                      !Cantsalid = 0
                      !CantIngre = Abs(CANTAA)
                  End If
                  .Update
                  Exit Do
               End If
               .MoveNext
             Loop
           End If
           On Error GoTo 0
         End With
         MOVVVISTO.Close
         Set MOVVVISTO = Nothing
       End If
     End If
   Next i&
   '
End Sub

Sub AJUPREVE()
   '
   NC1% = XVALO(Text1.TEXT)
   Call APERBASDAT("STOCKS")
   
   FIN& = ULTREG&("!DIFPRECO")
   '
   For i& = 1 To FIN&
     Call TRACE(20, i&, FIN&)
     x$ = REGLEIDO$("!DIFPRECO", i&)
     PRENUEVO# = CVD(Mid$(x$, 45, 8))
     If PRENUEVO# > 0.00005 Then
       Ci% = CVI(Mid$(x$, 1, 2))
       If Ci% > 0 Then
         DOPRIL$ = TRIM$(Mid$(x$, 81, 18))
         NORIT% = CVI(Mid$(x$, 117, 2))
         PREUNI = CVD(Mid$(x$, 45, 8))
         '
         SQLX$ = "DoPriLar = '" & DOPRIL$ & "' "
         SQLX$ = SQLX$ + "AND Cod_Inte = " & Ci% & " "
         SQLX$ = SQLX$ + "AND Nume_Clie = " & NC1% & " "
         SQLX$ = SQLX$ + "AND NuorItem = " & NORIT% & " "
         '
         Dim MOVVVISTO As ADODB.Recordset
         Set MOVVVISTO = New ADODB.Recordset
20       On Error Resume Next
         MOVVVISTO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
         If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 20
         End If
         On Error GoTo 0
         '
         With MOVVVISTO
           On Error Resume Next
           '.FindFirst SQLX$
25         'If .NoMatch = False Then
             '.Edit
           If Not (.EOF And .BOF) Then
             On Error GoTo 0
             !ValoUnit = PREUNI
             .Update
           End If
         End With
         MOVVVISTO.Close
         Set MOVVVISTO = Nothing
       End If
     End If
   Next i&
   '
End Sub

Sub CARTAMONED()
   If DERACCE%("TICAMBIO", "Ajustar Tipo de Cambio Central") < 2 Then
     Exit Sub
   End If
   '
   If ULTREG&("TAMONED") > 6 Then
      Call SETULTREG("TAMONED", 6)
      Call COMUNI("Tabla de Monedas Truncada\al Máximo de 6 Registros")
   End If
   '
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAMONED/NC/TITUPAN=Monedas y Tipo de Cambio")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAMONED")
         x$ = REGLEIDO$("!TAMONED", U&)
         Call REPLA(x$, Chr$(U&), 1, 1)
         Call GRAREG("TAMONED", x$, U&)
      Next U&
      FIN& = ULTREG&("!TAMONED")
      Call SETULTREG("TAMONED", FIN&)
   End If
   Call CIERRARCH("TAMONED")
   TICAMBIS = FORMATNUM(TICAMONE(2), "F9.4")
End Sub

Sub SHOWPEDIPEN()
     '
     Static UFEPEDET#
     If CDbl(FileDateTime(LUDAT$ + "PEDDETA.DAT")) <> UFEPEDET# Then
        UFEPEDET# = CDbl(FileDateTime(LUDAT$ + "PEDDETA.DAT"))
        '
        jj& = 0
        UREPRE& = ULTREG&("PEDDETA")
        Screen.MousePointer = 11
        jj& = 0
        NROPEDA& = 0: CAPED% = 0
        FINF = 999999: FSUP = 0
        For REPRESU& = 1 To UREPRE&
          Call TRACE(20, REPRESU&, UREPRE&)
          x$ = REGLEIDO$("PEDDETA", REPRESU&)
          If Asc(Mid$(x$, 27, 1)) <> 9 Then ' no esta anulado
            CAN = CVS(Mid$(x$, 79, 4))
            CADESPACH = CVS(Mid$(x$, 95, 4))
            CAPENDENT = CAN - CADESPACH
            If CAPENDENT > 0.5 Then
              NRODES& = CVS(Left$(x$, 4))
              If NRODES& <> NRODESA& Then
                PRIREDETA& = REPRESU&
                NRODESA& = NRODES&
                TTXX$ = MKS$(NRODES&)
                NOCO$ = TRIM$(Mid$(x$, 9, 16))
                If NOCO$ <> "" Then
                  TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
                End If
                FEX = CVI(Mid$(x$, 5, 2))
                TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
                NCLIE% = CVI(Mid$(x$, 7, 2))
                TTXX$ = TTXX$ + TRIM$(rasocli$(NCLIE%)) + Space$(52)
                REGENCA& = CVS(Mid$(x$, 109, 4))
                jj& = jj& + 1
                Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), jj&)
              End If
            End If
          End If
        Next REPRESU&
        Call SETULTREG("INDIPEP", jj&)
        Call CIERRARCH("INDIPEP")
        '
        Screen.MousePointer = 1
        YAINDI% = 1
     End If
     '
2112 If ULTREG&("INDIPEP") < 1 Then
        Call COMUNI("No hay O/V's Pendientes\que Puedan Consultarse")
        GoTo 9999
     End If
     '
2110 x$ = REGSEL$("INDIPEP/Ordenes de Venta Pendientes de Entrega")
     If Len(x$) <= 4 Then
        GoTo 9999
     End If
     '
     NPX& = CVS(Left$(x$, 4))
     If NPX& < 1 Then GoTo 9999
     '
     TIDOCUM$ = "Orden de Despacho"
     NUDOCU$ = TRIM$(Str$(NPX&))
     While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
     NUDOCU$ = NUDOCU$ + " "
     Screen.MousePointer = 11
     '
     For UI& = ULTREG&("PDOCLST") To 1 Step -1
       XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
       If InStr(XX$, TIDOCUM$) > 4 Then
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
     Screen.MousePointer = 1
     Call COMUNI("Documento no Encontrado")
     Exit Sub
     '
5    Call MUESTRADOC(DOCUNU&)
     '
     GoTo 2110
     '
9999 End Sub
'
Sub REGISNODE()
    'REGISTRAR NOTA DE DEBITO DEL CLIENTE
    '
    TIPOCOMP$ = "ND" 'CI(CREDITO INTERNO) ES NOTA DE DEBITO DEL CLIENTE
    OPCICOMP$ = "A" 'O "B" 'DEPENDIENDO DE LA POSICION DEL IVA DEL CLIENTE?
    '
    numeroComp$ = XVALO(Mid$(Text11, 1))
    NUMEROTALO$ = XVALO(Mid$(Text6, 1))
    'CORTO
    NCCMP$ = TIPOCOMP$ & "-" & numeroComp & "-" & OPCICOMP$
    'LARGO
    NLCMP$ = TIPOCOMP$ & "-" & TRIM$(Text6) & "-" & TRIM$(Text11) & "-" & OPCICOMP$
    '
    TOTDOCUM# = XVALO(Total(6))
    '
    NC% = XVALO(Text1)
    CoComCor = NCCMP$
    CoComLar = NLCMP$
    NumeTal = NUMEROTALO$
    NumeCom = numeroComp$
    Referen = Mid$(TRIM$(DescriTxt), 1, 48)
    FechEmi = Text9
    FeContab = Text9
    FeAcredi = Text9
    TiMovi$ = "FVEN"
    OpMovi = "CI"   ' CREDITO INTERNO
    VaMovi = OPCICOMP$
    NumeProv = 0
    NumeClie = Int(GRUDIVCLI(NC))
    CuentaCo% = CUECOCLI%(NC%)
'    CodiEmp% = 0
'    CodiSuc% = 0
    Referen$ = rasocli$(NC%)
    BasImpo# = 0
    ImpoDeb# = TOTDOCUM# * (-1)
    Descrip$ = "N/Debito Cliente " + NLCMP$
    Observa$ = rasocli$(NC%)
    StatInit = 0
    Call GRACABAN
    '
    'graba venta neta - Verificar que QQQ sea la IC de combinacion Cliente-Tipo de Comprobante
   TOTANETX# = XVALO(Total(2))
   If TOTANETX# >= 0.005 Then
     QQ1% = CUECOINT%(Text16)
     CuentaCo% = QQ1%   ' TOMA LA QUE APARECE EN PANTALLA
     TiMovi$ = "NDCLI"
     OpMovi$ = "CI"
     VaMovi$ = OPCICOMP$
     StatInit = 0
     ImpoDeb = TOTANETX#
     Call GRACABAN
   End If
   '
   'graba iva inscripto
   For KKI% = 3 To 4
     TOTANETX# = XVALO(Total(KKI%))
     If TOTANETX# >= 0.005 Then
       CUEIVA$ = CONTROL$("SVD451S", "CUEIVA")
       CuentaCo% = CUECOINT%(CUEIVA$)
       TiMovi$ = "IVA"
       OpMovi$ = "DeFi"
       VaMovi$ = PORIVA(KKI% - 3) + " %"
       Descrip$ = "Iva Debito Fiscal " + NCCMP$
       StatInit = 0
       ImpoDeb = TOTANETX#
       BasImpo# = XVALO(BASIVA(KKI% - 3))
       Call GRACABAN
     End If
   Next KKI%
   BasImpo# = 0
   '
   'graba iva NO inscripto
   TOTANETX# = XVALO(Total(5))
   If TOTANETX# >= 0.005 Then
     CUEIVA$ = CONTROL$("SVD451S", "CUEADU")
     CuentaCo% = CUECOINT%(CUEIVA$)
     TiMovi$ = "IVA"
     OpMovi$ = "DeFi"
     VaMovi$ = "RNI"
     Descrip$ = "IVA Resp.No Inscripto " + NCCMP$
     StatInit = 0
     ImpoDeb = TOTANETX#
     BasImpo# = XVALO(BASIVA(2))
     Call GRACABAN
   End If
   BasImpo# = 0
   '
   'graba percepcion I.Brutos
   TOTANETX# = XVALO(Text15)
   If TOTANETX# >= 0.005 Then
     CUEIVA$ = CONTROL$("PERIBRU", "CUECONTA")
     CuentaCo% = CUECOINT%(CUEIVA$)
     TiMovi$ = "IBRU"
     OpMovi$ = "IB-B"
     VaMovi$ = "PERC"
     Descrip$ = "Perc.IB.Bs.As. " + NCCMP$
     StatInit = 0
     ImpoDeb = TOTANETX#
     BasImpo# = XVALO(Total(2))
     Call GRACABAN
   End If
   BasImpo# = 0
   '
   'graba percepcion IVA
   TOTANETX# = XVALO(Text17)
   If TOTANETX# >= 0.005 Then
     CUEPER$ = CONTROL$("SVD451S", "CUERET")
     CuentaCo% = CUECOINT%(CUEPER$)
     TiMovi$ = "IVA"
     OpMovi$ = "DeFi"
     VaMovi$ = "PER-IVA"
     Descrip$ = "Percep.IVA " + NCCMP$
     StatInit = 0
     ImpoDeb = TOTANETX#
     BasImpo# = XVALO(Total(2))
     Call GRACABAN
   End If
   '
   'GRABA DEUDA PENDIENTE
   IBruTot# = TOTDOCUM#
   IIvaTot# = XVALO(Total(3)) + XVALO(Total(4))
   NuorVen% = 0
   FeVenc = FechEmi
   ImVenc# = (-1) * TOTDOCUM
   Call GRADEUPEN
   '
End Sub
