VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form RECALCOS07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Re-Cálculo Integral de Costos"
   ClientHeight    =   6150
   ClientLeft      =   165
   ClientTop       =   2205
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2520
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   90
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00CDDADA&
      BorderStyle     =   0  'None
      Caption         =   "Frame16"
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10515
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   6960
         ScaleHeight     =   165
         ScaleWidth      =   1695
         TabIndex        =   98
         Top             =   120
         Width           =   1725
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   99
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Pausa"
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
         Left            =   8925
         TabIndex        =   89
         ToolTipText     =   "Click para Interrumpir"
         Top             =   105
         Width           =   750
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00CDDADA&
         Caption         =   "COSTO TOTAL CALCULADO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2745
         Left            =   2520
         TabIndex        =   56
         Top             =   3150
         Width           =   7890
         Begin VB.Label Label99 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Index           =   5
            Left            =   6120
            TabIndex        =   96
            Top             =   1920
            Width           =   210
         End
         Begin VB.Label Label99 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Index           =   4
            Left            =   6120
            TabIndex        =   95
            Top             =   1200
            Width           =   210
         End
         Begin VB.Label Label99 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Index           =   3
            Left            =   6120
            TabIndex        =   94
            Top             =   840
            Width           =   210
         End
         Begin VB.Label Label99 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Index           =   2
            Left            =   2320
            TabIndex        =   93
            Top             =   1200
            Width           =   330
         End
         Begin VB.Label Label99 
            Alignment       =   1  'Right Justify
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
            Height          =   225
            Index           =   1
            Left            =   2320
            TabIndex        =   92
            Top             =   840
            Width           =   330
         End
         Begin VB.Label Label239 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   5460
            TabIndex        =   87
            Top             =   1785
            Width           =   645
         End
         Begin VB.Label Label243 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   5460
            TabIndex        =   86
            Top             =   1065
            Width           =   645
         End
         Begin VB.Label Label242 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   5460
            TabIndex        =   85
            Top             =   690
            Width           =   645
         End
         Begin VB.Label Label236 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   1680
            TabIndex        =   84
            Top             =   1065
            Width           =   750
         End
         Begin VB.Label Label238 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   1680
            TabIndex        =   83
            Top             =   1815
            Width           =   750
         End
         Begin VB.Label Label241 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   1680
            TabIndex        =   82
            Top             =   690
            Width           =   750
         End
         Begin VB.Label Label65 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   2730
            TabIndex        =   81
            Top             =   1440
            Width           =   1275
         End
         Begin VB.Label Label66 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "COSTO STD DE MANUFACTURA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -420
            TabIndex        =   80
            Top             =   1515
            Width           =   3060
         End
         Begin VB.Label Label67 
            BackStyle       =   0  'Transparent
            Caption         =   "GG.Materiales "
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
            Left            =   600
            TabIndex        =   79
            Top             =   820
            Width           =   1350
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2730
            TabIndex        =   78
            Top             =   1065
            Width           =   1275
         End
         Begin VB.Label Label69 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Costos Directos"
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
            Left            =   525
            TabIndex        =   77
            Top             =   440
            Width           =   2115
         End
         Begin VB.Label Label70 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   2730
            TabIndex        =   76
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label77 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2730
            TabIndex        =   75
            Top             =   1815
            Width           =   1275
         End
         Begin VB.Label Label96 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   2730
            TabIndex        =   74
            Top             =   2250
            Width           =   1275
         End
         Begin VB.Label Label97 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Gastos Grales Admin. y Ventas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   73
            Top             =   1740
            Width           =   1485
         End
         Begin VB.Label Label98 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "COSTO STANDARD TOTAL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -420
            TabIndex        =   72
            Top             =   2390
            Width           =   3060
         End
         Begin VB.Label Label99 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "%:"
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
            Left            =   2350
            TabIndex        =   71
            Top             =   1935
            Width           =   330
         End
         Begin VB.Label Label101 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "PRECIO BÁSICO DE VENTA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4200
            TabIndex        =   70
            Top             =   2400
            Width           =   2115
         End
         Begin VB.Label Label102 
            BackStyle       =   0  'Transparent
            Caption         =   "Impuestos"
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
            Left            =   4650
            TabIndex        =   69
            Top             =   1200
            Width           =   1035
         End
         Begin VB.Label Label103 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   6405
            TabIndex        =   68
            Top             =   2250
            Width           =   1275
         End
         Begin VB.Label Label104 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   6405
            TabIndex        =   67
            Top             =   1065
            Width           =   1275
         End
         Begin VB.Label Label105 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   6405
            TabIndex        =   66
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label106 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Costo Standard Total"
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
            Left            =   4200
            TabIndex        =   65
            Top             =   440
            Width           =   2115
         End
         Begin VB.Label Label107 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   6405
            TabIndex        =   64
            Top             =   1815
            Width           =   1275
         End
         Begin VB.Label Label108 
            BackStyle       =   0  'Transparent
            Caption         =   "Utilidad"
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
            Left            =   4880
            TabIndex        =   63
            Top             =   1920
            Width           =   825
         End
         Begin VB.Label Label109 
            BackStyle       =   0  'Transparent
            Caption         =   " Costo s/Venta"
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
            Left            =   4310
            TabIndex        =   62
            Top             =   840
            Width           =   1365
         End
         Begin VB.Label Label110 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   6405
            TabIndex        =   61
            Top             =   690
            Width           =   1275
         End
         Begin VB.Label Label93 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2730
            TabIndex        =   60
            Top             =   690
            Width           =   1275
         End
         Begin VB.Label Label124 
            BackStyle       =   0  'Transparent
            Caption         =   "GG.Elaboración"
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
            Left            =   480
            TabIndex        =   59
            Top             =   1190
            Width           =   1455
         End
         Begin VB.Label Label100 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "COSTO TOTAL BRUTO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4215
            TabIndex        =   58
            Top             =   1515
            Width           =   2100
         End
         Begin VB.Label Label111 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   6405
            TabIndex        =   57
            Top             =   1440
            Width           =   1275
         End
      End
      Begin VB.Frame Frame26 
         BackColor       =   &H00CDDADA&
         Caption         =   "COBERTURA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2745
         Left            =   0
         TabIndex        =   42
         Top             =   3150
         Width           =   2430
         Begin VB.Label Label244 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Height          =   280
            Left            =   330
            TabIndex        =   88
            Top             =   640
            Width           =   540
         End
         Begin VB.Label Label113 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Precio Lista"
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
            Left            =   -1230
            TabIndex        =   55
            Top             =   420
            Width           =   2115
         End
         Begin VB.Label Label114 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Height          =   280
            Left            =   945
            TabIndex        =   54
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label115 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Height          =   280
            Left            =   945
            TabIndex        =   53
            Top             =   640
            Width           =   1275
         End
         Begin VB.Label Label113 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Precio Neto"
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
            Index           =   1
            Left            =   -1230
            TabIndex        =   52
            Top             =   1070
            Width           =   2115
         End
         Begin VB.Label Label116 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Height          =   280
            Left            =   945
            TabIndex        =   51
            Top             =   1290
            Width           =   1275
         End
         Begin VB.Label Label113 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Margen"
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
            Left            =   -1230
            TabIndex        =   50
            Top             =   1365
            Width           =   2115
         End
         Begin VB.Label Label119 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   285
            Left            =   1420
            TabIndex        =   49
            Top             =   2325
            Width           =   795
         End
         Begin VB.Label Label113 
            BackStyle       =   0  'Transparent
            Caption         =   "Indice Calidad de Precio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   4
            Left            =   120
            TabIndex        =   48
            Top             =   2250
            Width           =   1395
         End
         Begin VB.Line Line5 
            BorderColor     =   &H80000003&
            X1              =   10
            X2              =   2400
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Line Line6 
            BorderColor     =   &H80000005&
            X1              =   15
            X2              =   2400
            Y1              =   1695
            Y2              =   1695
         End
         Begin VB.Label Label64 
            BackStyle       =   0  'Transparent
            Caption         =   "Dto"
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
            Left            =   45
            TabIndex        =   47
            Top             =   720
            Width           =   330
         End
         Begin VB.Label Label112 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Height          =   280
            Left            =   945
            TabIndex        =   46
            Top             =   965
            Width           =   1275
         End
         Begin VB.Label Label113 
            BackStyle       =   0  'Transparent
            Caption         =   "Mg.Bto(%):      Mg.Neto(%):   "
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
            Index           =   5
            Left            =   350
            TabIndex        =   45
            Top             =   1725
            Width           =   2100
         End
         Begin VB.Label Label117 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   285
            Left            =   1420
            TabIndex        =   44
            Top             =   1980
            Width           =   795
         End
         Begin VB.Label Label118 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   285
            Left            =   315
            TabIndex        =   43
            Top             =   1980
            Width           =   795
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Stop"
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
         Left            =   9630
         TabIndex        =   41
         ToolTipText     =   "Cancela Proceso de Cálculo"
         Top             =   105
         Width           =   765
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00CDDADA&
         Caption         =   "AVANCE DE PROCESO"
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
         Left            =   6720
         TabIndex        =   23
         Top             =   420
         Width           =   3690
         Begin VB.Frame Frame4 
            BackColor       =   &H00CDDADA&
            Caption         =   "REGISTROS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   160
            TabIndex        =   34
            Top             =   1520
            Width           =   1560
            Begin VB.Label Label16 
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
               Height          =   255
               Index           =   0
               Left            =   720
               TabIndex        =   40
               Top             =   210
               Width           =   795
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Falta"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   -1500
               TabIndex        =   39
               Top             =   720
               Width           =   2175
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -1500
               TabIndex        =   38
               Top             =   255
               Width           =   2175
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Proceso"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -1500
               TabIndex        =   37
               Top             =   480
               Width           =   2175
            End
            Begin VB.Label Label16 
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
               Height          =   255
               Index           =   1
               Left            =   720
               TabIndex        =   36
               Top             =   450
               Width           =   795
            End
            Begin VB.Label Label16 
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
               Height          =   255
               Index           =   2
               Left            =   720
               TabIndex        =   35
               Top             =   685
               Width           =   795
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00CDDADA&
            Caption         =   "TIEMPOS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Index           =   0
            Left            =   1850
            TabIndex        =   27
            Top             =   1520
            Width           =   1660
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Proceso"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -1490
               TabIndex        =   33
               Top             =   480
               Width           =   2175
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Estimado"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -1490
               TabIndex        =   32
               Top             =   255
               Width           =   2175
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Falta"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -1490
               TabIndex        =   31
               Top             =   720
               Width           =   2175
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "00.00.00"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   580
               TabIndex        =   30
               Top             =   225
               Width           =   975
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "00.00.00"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   580
               TabIndex        =   29
               Top             =   455
               Width           =   975
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "00.00.00"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   580
               TabIndex        =   28
               Top             =   685
               Width           =   975
            End
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00CDDADA&
            Caption         =   "Artículo"
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
            Left            =   210
            TabIndex        =   26
            Top             =   490
            Width           =   1170
         End
         Begin VB.Label Label2 
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
            Height          =   540
            Left            =   160
            TabIndex        =   25
            Top             =   795
            Width           =   3320
         End
         Begin VB.Label Label1 
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
            Left            =   1470
            TabIndex        =   24
            Top             =   360
            Width           =   2010
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00CDDADA&
         Caption         =   "COSTO DE ELABORACIÓN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Index           =   4
         Left            =   0
         TabIndex        =   15
         Top             =   1890
         Width           =   6630
         Begin VB.Label Label72 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   3780
            TabIndex        =   22
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label74 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   3780
            TabIndex        =   21
            Top             =   690
            Width           =   1275
         End
         Begin VB.Label Label75 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   2415
            TabIndex        =   20
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label76 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mano de Obra Directa"
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
            Left            =   210
            TabIndex        =   19
            Top             =   450
            Width           =   2115
         End
         Begin VB.Label Label78 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Equipos de Producción"
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
            Left            =   0
            TabIndex        =   18
            Top             =   810
            Width           =   2325
         End
         Begin VB.Label Label79 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   2415
            TabIndex        =   17
            Top             =   690
            Width           =   1275
         End
         Begin VB.Label Label80 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   5145
            TabIndex        =   16
            ToolTipText     =   "Doble-click Para Ver Análisis de Costos de Mano de Obra "
            Top             =   690
            Width           =   1275
         End
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00CDDADA&
         Caption         =   "COSTO DE SUMINISTROS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1850
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   6630
         Begin VB.Frame Frame3 
            Height          =   40
            Index           =   1
            Left            =   10
            TabIndex        =   97
            Top             =   480
            Width           =   6615
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Totales"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   5230
            TabIndex        =   14
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Recibidos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   3890
            TabIndex        =   13
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "C.Propios"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   2550
            TabIndex        =   12
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label Label81 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Trabajos de Terceros"
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
            Left            =   210
            TabIndex        =   11
            Top             =   1490
            Width           =   2115
         End
         Begin VB.Label Label82 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   3780
            TabIndex        =   10
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label Label83 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   3780
            TabIndex        =   9
            Top             =   1005
            Width           =   1275
         End
         Begin VB.Label Label84 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   3780
            TabIndex        =   8
            Top             =   1365
            Width           =   1275
         End
         Begin VB.Label Label85 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   5145
            TabIndex        =   7
            ToolTipText     =   "Doble-click para Ver Analisis de Costos de Materiales y Servicios"
            Top             =   1365
            Width           =   1275
         End
         Begin VB.Label Label86 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   2415
            TabIndex        =   6
            Top             =   1365
            Width           =   1275
         End
         Begin VB.Label Label87 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Piezas Compradas"
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
            Left            =   210
            TabIndex        =   5
            Top             =   1130
            Width           =   2115
         End
         Begin VB.Label Label88 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   2415
            TabIndex        =   4
            Top             =   1005
            Width           =   1275
         End
         Begin VB.Label Label89 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Materias Primas"
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
            Left            =   210
            TabIndex        =   3
            Top             =   770
            Width           =   2115
         End
         Begin VB.Label Label90 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   2415
            TabIndex        =   2
            Top             =   630
            Width           =   1275
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "RECALCOS07.frx":0000
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "RECALCOS07.frx":0234
      TabIndex        =   91
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "RECALCOS07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PROVINT%(32767)
Dim PAUSA%
'
Dim COCAMOD(32767), COCAMAQ(32767), COSUMAT(32767), UMEDIX$(32767), PRHABI%(32767)
Dim PORGAGE(7, 255), MODOCA%(7, 255)
Dim PREVENETO#, MODOCALCU%(7)
Dim HAYRECON%
'
Sub RECACOS()
   '
   Dim DP%(16384)
   '
   Screen.MousePointer = 11
   '
   TINI1 = Timer
   FECO% = HOY(HOS$)
   NVI% = NUINV%
   NMM% = NUMAS%
   URE& = ULTREG&("SECOPER")
   UCT& = ULTREG&("COSOPTER")
   '
   ' INICIALIZANDO ARCHIVO
   ' IDENTIFICANDO LOS ITEMS SIN EXPLOSION
   '
   HAYRECON% = 0
   If EXISTE%(LUDAT$ + "RECOSCON.RFS") > 0 Then HAYRECON% = 1
   '
   Call COPYSTRU("RECOSCAL", "TRACALCO")
   '\\\OT-06-0229-WAR-17/07/06///
   If HAYRECON% > 0 Then Call COPYSTRU("RECOSCON", "TRCOSCON")
   '\\\OT-06-0229-WAR-FIN///
   FIN& = NUMAS%
   For i& = 1 To FIN&
     Call TRACE(20, i&, FIN&)
     CII% = i&
     TIMA% = TIMATE%(CII%) ' Asc(Mid$(XX$, 93, 1))
     UMED$ = Unimed$(CII%) ' AJUSTI$(Mid$(XX$, 109, 2), 2)
     YY$ = String$(128, 0)
     Call REPLA(YY$, MKI$(CII%), 1, 2)
     Call REPLA(YY$, MKI$(TIMA%), 3, 2)
     Call REPLA(YY$, UMED$, 5, 2)
     Call REPLA(YY$, MKI$(FECO%), 7, 2)
     Call GRAREG("TRACALCO", YY$, i&)
     '\\\OT-06-0229-WAR-17/07/06///
     If HAYRECON% > 0 Then
        ZZ$ = String$(128, 0)
        Call REPLA(ZZ$, MKI$(CII%), 1, 2)
        Call REPLA(ZZ$, MKI$(TIMA%), 3, 2)
        Call REPLA(ZZ$, UMED$, 5, 2)
        Call REPLA(ZZ$, MKI$(FECO%), 7, 2)
        Call GRAREG("TRCOSCON", ZZ$, i&)
     End If
     '\\\OT-06-0229-WAR-FIN///
     '
     'CODYY$ = CODEXT$(CII%)
     'Call LEEEXPLO(CODYY$, 1)
     'Call AREXPLO(CII%, 1, 3)
     CAIT1% = CANCOMPO%(CII%)
     If CAIT1% > 0 Then
       DP%(CII%) = 1
     End If
   Next i&
   '
   Call SETULTREG("TRACALCO", FIN&)
   Call CIERRARCH("TRACALCO")
   '\\\OT-06-0229-WAR-17/07/06///
   If HAYRECON% > 0 Then
     Call SETULTREG("TRCOSCON", FIN&)
     Call CIERRARCH("TRCOSCON")
   End If
   '\\\OT-06-0229-WAR-FIN///
   '
   FINCALCOS% = 0
   '
   Screen.MousePointer = 1
   'Set MASTERECO = Articulos.OpenRecordset("SELECT CODIGO,DESCRIPCION,CODINT FROM Master ORDER BY CODIGO ASC;", dbOpenSnapshot, dbReadOnly)
   SQLX$ = "SELECT CODIGO,DESCRIPCION,CODINT "
   SQLX$ = SQLX$ + " FROM Master "
   SQLX$ = SQLX$ + " ORDER BY CODIGO ASC "
   Dim MASTERECO As ADODB.Recordset
   Set MASTERECO = New ADODB.Recordset
   MASTERECO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   
   On Error Resume Next
   MASTERECO.MoveLast
   If Err Then On Error GoTo 0: Exit Sub
   '
   CAREC% = MASTERECO.RecordCount
   If CAREC% < 1 Then Exit Sub
    '
   FIN& = CAREC%
   For IPIE& = 1 To FIN&
     CI% = IPIE&
     CODD$ = CODEXT$(CI%)
     If STATUIT%(CI%) > 0 Then
       If CI% > 0 Then
         If CI% <= NMM% Then
            CICON% = CI%
            '
            Label1 = CODD$
            Label2 = DESCRIT$(CI%)
            '
            GRCV% = GRUCOVEI%(CICON%)
            If GRCV% < 1 Or GRCV% > 255 Then GRCV% = 0
            Label241 = FORMATNUM$(PORGAGE(1, GRCV%), "F5.1")
            Label236 = FORMATNUM$(PORGAGE(2, GRCV%), "F5.1")
            Label238 = FORMATNUM$(PORGAGE(3, GRCV%), "F5.1")
            Label242 = FORMATNUM$(PORGAGE(4, GRCV%), "F4.1")
            Label243 = FORMATNUM$(PORGAGE(5, GRCV%), "F4.1")
            Label239 = FORMATNUM$(PORGAGE(6, GRCV%), "F4.1")
            Label244 = FORMATNUM$(PORGAGE(7, GRCV%), "F4.1")
            '
            For KKII% = 0 To 7
              MODOCALCU%(KKII%) = MODOCA%(KKII%, GRCV%)
            Next KKII%
            '
            Call EXPLOMUL(CI%)
            SUMPPROP = 0: SUMPREC = 0
            SUPIPROP = 0: SUPIREC = 0
            SUSEPROP = 0: SUSEREC = 0
            SUMODPRO = 0: SUMODREC = 0
            SUMAQPRO = 0: SUMAQREC = 0
            '
            SUPRI = 0: SUPIE = 0: SUSER = 0
            SUTIE = 0: SUMOD = 0: SUMAQ = 0
            '
            ' costo de materias primas y piezas de terceros
            '
            For KKI% = 1 To CONTAMUL
              If DP%(CIK%(KKI%)) < 1 Then   ' NO TIENE ESTRUCTURA
                '
                MONEI% = MONECOSTO%(CIK%(KKI%))
                COSUN = COSUNI(CIK%(KKI%)) * TICAMONE(MONEI%) * (1 + PORNACIO(CIK%(KKI%)) / 100)
                UNIME$ = UCase$(Unimed$(CIK%(KKI%)))
                '
16              If Left$(UNIME$, 1) = "U" Or Left$(UNIME$, 2) = "PZ" Then
                     If KKI% = 1 Then
                          SUPIPROP = SUPIPROP + CANT(KKI%) * COSUN
                        Else
                          SUPIREC = SUPIREC + CANT(KKI%) * COSUN
                     End If
                     SUPIE = SUPIE + CANT(KKI%) * COSUN
                   Else
                     If KKI% = 1 Then
                          SUMPPROP = SUMPPROP + COSUN * CANT(KKI%)
                        Else
                          SUMPREC = SUMPREC + COSUN * CANT(KKI%)
                     End If
                     SUPRI = SUPRI + CANT(KKI%) * COSUN
                End If
                '
              Else
                '
                If COSUNI(CIK%(KKI%)) > 0 Then
                  PHB% = PROVHABI%(CIK%(KKI%))
                  If PHB% > 0 Then
                    If PROVINT%(PHB%) = 0 Then
                      MONEI% = MONECOSTO%(CIK%(KKI%))
                      COSUN = COSUNI(CIK%(KKI%)) * TICAMONE(MONEI%) * (1 + PORNACIO(CIK%(KKI%)) / 100)
                      '
                      If KKI% = 1 Then
                           SUSEPROP = SUSEPROP + COSUN * CANT(KKI%)
                         Else
                           SUSEREC = SUSEREC + CANT(KKI%) * COSUN
                      End If
                      SUSER = SUSER + CANT(KKI%) * COSUN
                    End If
                  End If
                End If
                '
              End If
            Next KKI%
            '
            '
            ' costo de mano de obra directa
            '
            FIXX& = CONTAMUL%
            For UI& = 1 To FIXX&
              '
              CIXI% = CIK%(UI&)
              CACOM = CANT(UI&)
              '
              SQLX$ = "SELECT * FROM Secoper"
              SQLX$ = SQLX$ & " WHERE CodIConj = " & CIXI% & " "
              SQLX$ = SQLX$ & "AND StatOper > 0 "
              'Set SECOPTEMP = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot)
              Dim SECOPTEMP As ADODB.Recordset
              Set SECOPTEMP = New ADODB.Recordset
              SECOPTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
              '
              On Error Resume Next
              With SECOPTEMP
                .MoveFirst
                If Err Then GoTo 25
                '
                Do Until .EOF = True
                  If IsNull(!AltOper) = True Or !AltOper < 1 Then       '
                     NOPY% = !NUMEOPER
                     DOPE$ = !descoper
                     LOTESTA = LOTESTAN(CIXI%)
                        If LOTESTA < 1 Then LOTESTA = 1
                     CICLOHE = !PieCiclo
                        If CICLOHE < 1 Then CICLOHE = 1
                     MINPREP = 60 * !HorsPrep
                     MINFIJO = 60 * !HorsFijo
                     MINVARI = LOTESTA * !MinuStan
                     MINHERR = LOTESTA * 60 * !HorsMHer / CICLOHE
                     PIECICL = !PieCiclo
                     HORLIMP = !HorsLimp
                     DEMODIS = !PorDemor
                     MINDEMO = (MINFIJO + MINVARI + MINHERR) * DEMODIS / 100
                     MINLOTE = MINFIJO + MINVARI + MINHERR + MINDEMO
                     HORAPROD = CACOM * MINLOTE / 60 / LOTESTA
                     '
                     SOLAPAI% = !CoefSolp
                     '
                     CAOPS = !CantiOps
                     CATOP% = !CATEGOPS
                     EQUIPO$ = !CodMaq0
                     EQUIP1$ = !CodMaq1
                     EQUIP2$ = !CodMaq2
                     EQUIP3$ = !CodMaq3
                     CUOTOPE = VALHORA(CATOP%)
                     CUOTMAQ = CUHORMAQ(EQUIPO$)
                     '
                     PROTERI% = !ProTerOp  ' PROVEEDOR OPERACION TERCERIZADA
                     PRECTER# = !PreUnid
                     MONETER% = !MonePrec
                     FEPRETER% = FECHANUM(Format$(!FechPrec, "dd/mm/yy"))
                     '
                     EQUIPO$ = TRIM$(Mid$(SCP$, 53, 6))
                     CAOPREP = CANOPREP(CIXI%, NOPY%)
                     If CAOPREP < 1 Then CAOPREP = CAOPS
                     CTOPREP% = CATOPREP%(CIXI%, NOPY%)
                     If CTOPREP% < 1 Then CTOPREP% = CATOP%
                     CUOTOPREP = VALHORA(CTOPREP%)
                     '
                     COSMOD = CUOTOPE * CAOPS * HORAPROD + CUOTOPREP * CAOPREP * MINPREP / 60 / LOTESTA
                     COSMAQ = CUOTMAQ * (HORAPROD + MINPREP / 60 / LOTESTA)
                     '
                     TERCOP% = 0
                     'If PROTERI% > 0 Then
                       MONEI% = !MonePrec  'MONEI% = MONETER%
                       COSTER = CACOM * !PreUnid * TICAMONE(MONEI%)  'COSTER = CACOM * PRECTER# * TICAMONE(MONEI%)
                       If MONEI% <= 1 Then
                         COSTER = CACOM * !PreUnid * IAJUINF#(FEBAS%, FEACT%)
                       End If
                       If COSTER >= 0.00005 Then
                         If (COSTER < COSMOD + COSMAQ) Or ((COSMOD + COSMAQ) < 0.00005) Then
                           COSMOD = COSTER
                           COSMAQ = 0
                           HORAPROD = 0
                           TERCOP% = 1
                         End If
                       End If
                     'End If
                     '
                     SUTIE = SUTIE + HORAPROD
                     SUMOD = SUMOD + COSMOD
                     SUMAQ = SUMAQ + COSMAQ
                     '
                     If UI& = 1 Then
                          SUMODPRO = SUMODPRO + COSMOD
                          SUMAQPRO = SUMAQPRO + COSMAQ
                        Else
                          SUMODREC = SUMODREC + COSMOD
                          SUMAQREC = SUMAQREC + COSMAQ
                     End If
                  End If
                .MoveNext
                Loop
                '
25            End With
              Set SECOPTEMP = Nothing

            Next UI&
            '
            GoSub 90
            '
            PRELISX# = PRELISTA#("LIPRE001", CICON%) * TICAMONE(MONELISTA%(1))
            
            Label114.Caption = TRIM$(CSTRING$(MKD$(PRELISX#), 3, 12, DECICALCOS%, 2))
            DoEvents
            PREVENETO# = XVALO(Label112)
            '
            Call LABEL241_Change
            Call LABEL236_Change
            Call LABEL238_Change
            Call LABEL242_Change
            Call LABEL243_Change
            Call LABEL239_Change
            DoEvents
            '
            RETRA& = CICON%
            YY$ = REGLEIDO$("TRACALCO", RETRA&)
            Call REPLA(YY$, MKD$(Val(Label90) + Val(Label82)), 9, 8)
            Call REPLA(YY$, MKD$(Val(Label88) + Val(Label83)), 17, 8)
            Call REPLA(YY$, MKD$(Val(Label86) + Val(Label84)), 25, 8)
            Call REPLA(YY$, MKD$(Val(Label85)), 33, 8)
            Call REPLA(YY$, MKD$(Val(Label75) + Val(Label72)), 41, 8)
            Call REPLA(YY$, MKD$(Val(Label79) + Val(Label74)), 49, 8)
            Call REPLA(YY$, MKD$(0), 57, 8)
            Call REPLA(YY$, MKD$(Val(Label80)), 65, 8)
            Call REPLA(YY$, MKD$(Val(Label70)), 73, 8)
            Call REPLA(YY$, MKD$(Val(Label65)), 81, 8)
            Call REPLA(YY$, MKD$(Val(Label96)), 89, 8)
            Call REPLA(YY$, MKD$(Val(Label103)), 97, 8)
            Call REPLA(YY$, MKD$(Val(Label112)), 105, 8)
            Call REPLA(YY$, MKD$(Val(Label116)), 113, 8)
            Call REPLA(YY$, MKS$(Val(Label117)), 121, 4)
            Call REPLA(YY$, MKS$(Val(Label119)), 125, 4)
            Call GRAREG("TRACALCO", YY$, RETRA&)
            '
            '\\\OT-06-0299-WAR-17/07/06///
            If HAYRECON% > 0 Then
               ZZ$ = REGLEIDO$("TRCOSCON", RETRA&)
               Call REPLA(ZZ$, MKD$(Val(Label70)), 9, 8)
               Call REPLA(ZZ$, MKD$(Val(Label104)), 17, 8)
                 GSINDI1# = XVALO(Label93) + XVALO(Label68) + XVALO(Label77) + XVALO(Label110)
               Call REPLA(ZZ$, MKD$(GSINDI1#), 25, 8)
                 GSNAZ1# = 0 'GS.NACIONALIZACION
                 MONECO1% = MONECOSTO(CICON%)
                 GSNAZ1# = (COSUNI(CICON%) * TICAMONE(MONECO1%) * PORNACIO(CICON%)) / 100
               Call REPLA(ZZ$, MKD$(GSNAZ1#), 33, 8)
                 TOTA1# = XVALO(Label70) + XVALO(Label104) + GSINDI1# + GSNAZ1#
               Call REPLA(ZZ$, MKD$(TOTA1#), 41, 8)
                 PRVTA1# = XVALO(Label112)
               Call REPLA(ZZ$, MKD$(PRVTA1#), 49, 8)
                 MARG1# = 0
'                 If PRVTA1# > 0 Then 'Evita División por Cero.
'                   marg1# = TOTA1# / PRVTA1#
'                 End If
                 ' modificado segun solic.Pablo Pierri 25/07/06
                 If TOTA1# > 0 Then 'Evita División por Cero.
                   MARG1# = 100 * (PRVTA1# / TOTA1# - 1)
                 End If
               Call REPLA(ZZ$, MKD$(MARG1#), 57, 8)
               Call GRAREG("TRCOSCON", ZZ$, RETRA&)
            End If
            '\\\OT-06-0299-WAR-FIN///
            '
         End If
       End If
     End If
     '
     Label16(0) = TRIM$(Str$(FIN&))
     Label16(1) = TRIM$(Str$(IPIE&))
     Label16(2) = TRIM$(Str$(FIN& - IPIE&))
     Label10(0) = TRIM$(HORMINSE$((Timer - TINI1) * FIN& / IPIE&))
     Label10(1) = TRIM$(HORMINSE$(Timer - TINI1))
     Label10(2) = TRIM$(HORMINSE$((Timer - TINI1) * (FIN& / IPIE& - 1)))
     '
     DoEvents
     '
     While PAUSA% = 1
     DoEvents
     Wend
     '
   Next IPIE&
   '
   Screen.MousePointer = 11
   FEX = HOY(HOS$)
   Call GRACONTROL("ANACOSTO", "FECALCOS", HOS$)
   Call GRACONTROL("ANACOSTO", "HOCALCOS", Left$(Time$, 5))
   '
   ANACOSTO07.Label34(4).Caption = HOS$
   ANACOSTO07.Label34(4).Refresh
   ANACOSTO07.Label34(5).Caption = Left$(Time$, 5)
   ANACOSTO07.Label34(5).Refresh
   '
   FIN& = ULTREG&("TRACALCO")
   For i& = 1 To FIN&
     Call TRACE(20, i&, FIN&)
     XX$ = REGLEIDO$("TRACALCO", i&)
     Call GRAREG("RECOSCAL", XX$, i&)
   Next i&
   Call SETULTREG("RECOSCAL", FIN&)
   Call CIERRARCH("RECOSCAL")
   '\\\OT-06-0299-WAR-17/07/06///
   If HAYRECON% > 0 Then
      FIN& = ULTREG&("TRCOSCON")
      For i& = 1 To FIN&
        Call TRACE(20, i&, FIN&)
        XX$ = REGLEIDO$("TRCOSCON", i&)
        Call GRAREG("RECOSCON", XX$, i&)
      Next i&
      Call SETULTREG("RECOSCON", FIN&)
      Call CIERRARCH("RECOSCON")
   End If
   '\\\OT-06-0299-WAR-FIN///
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   FINCALCOS% = 1
   Call COMUNI("Re-Cálculo Completo")
   Unload RECALCOS07
   Screen.MousePointer = 1
   '
Exit Sub

90 Label90.Caption = TRIM$(CSTRING$(MKS$(SUMPPROP), 3, 12, DECICALCOS%, 2))
   Label90.Refresh
   Label82.Caption = TRIM$(CSTRING$(MKS$(SUMPREC), 3, 12, DECICALCOS%, 2))
   Label82.Refresh
   Label88.Caption = TRIM$(CSTRING$(MKS$(SUPIPROP), 3, 12, DECICALCOS%, 2))
   Label88.Refresh
   Label83.Caption = TRIM$(CSTRING$(MKS$(SUPIREC), 3, 12, DECICALCOS%, 2))
   Label83.Refresh
   Label86.Caption = TRIM$(CSTRING$(MKS$(SUSEPROP), 3, 12, DECICALCOS%, 2))
   Label86.Refresh
   Label84.Caption = TRIM$(CSTRING$(MKS$(SUSEREC), 3, 12, DECICALCOS%, 2))
   Label84.Refresh
   '
   Label75.Caption = TRIM$(CSTRING$(MKS$(SUMODPRO), 3, 12, DECICALCOS%, 2))
   Label75.Refresh
   Label72.Caption = TRIM$(CSTRING$(MKS$(SUMODREC), 3, 12, DECICALCOS%, 2))
   Label72.Refresh
   Label79.Caption = TRIM$(CSTRING$(MKS$(SUMAQPRO), 3, 12, DECICALCOS%, 2))
   Label79.Refresh
   Label74.Caption = TRIM$(CSTRING$(MKS$(SUMAQREC), 3, 12, DECICALCOS%, 2))
   Label74.Refresh
   '
   SUMA1 = 0
   SUMA1 = SUMA1 + Val(Label90.Caption)
   SUMA1 = SUMA1 + Val(Label88.Caption)
   SUMA1 = SUMA1 + Val(Label86.Caption)
   SUMA1 = SUMA1 + Val(Label82.Caption)
   SUMA1 = SUMA1 + Val(Label83.Caption)
   SUMA1 = SUMA1 + Val(Label84.Caption)
   Label85.Caption = TRIM$(CSTRING$(MKS$(SUMA1), 3, 12, DECICALCOS%, 2))
   Label85.Refresh
   '
   SUMA2 = 0
   SUMA2 = SUMA2 + Val(Label75.Caption)
   SUMA2 = SUMA2 + Val(Label72.Caption)
   SUMA2 = SUMA2 + Val(Label79.Caption)
   SUMA2 = SUMA2 + Val(Label74.Caption)
   Label80.Caption = TRIM$(CSTRING$(MKS$(SUMA2), 3, 12, DECICALCOS%, 2))
   Label80.Refresh
   '
   SUMA3 = SUMA1 + SUMA2
   Label70.Caption = TRIM$(CSTRING$(MKS$(SUMA3), 3, 12, DECICALCOS%, 2))
   Label70.Refresh
   '
Return
   
End Sub
'
Sub RECACOS_NUE()
   '
   Dim DP%(32767)
   '
   Screen.MousePointer = 11
   '
   TINI1 = Timer
   FECO% = HOY(HOS$)
   NVI% = NUINV%
   NMM% = NUMAS%
   URE& = ULTREG&("SECOPER")
   UCT& = ULTREG&("COSOPTER")
   '
   ' INICIALIZANDO ARCHIVO
   ' IDENTIFICANDO LOS ITEMS SIN EXPLOSION
   '
   HAYRECON% = 0
   If EXISTE%(LUDAT$ + "RECOSCON.RFS") > 0 Then HAYRECON% = 1
   '
   Call COPYSTRU("RECOSCAL", "TRACALCO")
   '\\\OT-06-0229-WAR-17/07/06///
   If HAYRECON% > 0 Then Call COPYSTRU("RECOSCON", "TRCOSCON")
   '\\\OT-06-0229-WAR-FIN///
   FIN& = NMM%
   For i& = 1 To FIN&
     Call TRACE(24, i&, FIN&)
     CII% = i&
     'Label1 = CODEXT$(CII%)
     'Label2 = DESCRIT$(CII%)
     TIMA% = TIMATE%(CII%) ' Asc(Mid$(XX$, 93, 1))
     UMED$ = Unimed$(CII%) ' AJUSTI$(Mid$(XX$, 109, 2), 2)
     YY$ = String$(128, 0)
     Call REPLA(YY$, MKI$(CII%), 1, 2)
     Call REPLA(YY$, MKI$(TIMA%), 3, 2)
     Call REPLA(YY$, UMED$, 5, 2)
     Call REPLA(YY$, MKI$(FECO%), 7, 2)
     Call GRAREG("TRACALCO", YY$, i&)
     '\\\OT-06-0229-WAR-17/07/06///
     If HAYRECON% > 0 Then
        ZZ$ = String$(128, 0)
        Call REPLA(ZZ$, MKI$(CII%), 1, 2)
        Call REPLA(ZZ$, MKI$(TIMA%), 3, 2)
        Call REPLA(ZZ$, UMED$, 5, 2)
        Call REPLA(ZZ$, MKI$(FECO%), 7, 2)
        Call GRAREG("TRCOSCON", ZZ$, i&)
     End If
     '
     COCAMOD(CII%) = 0: COCAMAQ(CII%) = 0: COSUMAT(CII%) = 0: PRHABI%(CII%) = 0
     Call CALCOSMOD(CII%)
     Call CALCOSMAT(CII%)
     '
     CAIT1% = CANCOMPO%(CII%)
     If CAIT1% > 0 Then
       DP%(CII%) = 1
     End If
     If IPIE& > 0 Then Label10(0) = TRIM$(HORMINSE$((Timer - TINI1) * FIN& / IPIE&))
     Label10(1) = TRIM$(HORMINSE$(Timer - TINI1))
   Next i&
   '
   Call SETULTREG("TRACALCO", FIN&)
   Call CIERRARCH("TRACALCO")
   '\\\OT-06-0229-WAR-17/07/06///
   If HAYRECON% > 0 Then
     Call SETULTREG("TRCOSCON", FIN&)
     Call CIERRARCH("TRCOSCON")
   End If
   '\\\OT-06-0229-WAR-FIN///
   '
   FINCALCOS% = 0
   '
   CAREC& = CantRegistros&("master", "status>0")
   
   Screen.MousePointer = 1
   'Set MASTERECO = Articulos.OpenRecordset("SELECT CODIGO,DESCRIPCION,CODINT FROM Master ORDER BY CODIGO ASC;", dbOpenSnapshot, dbReadOnly)
   SQLX$ = "SELECT CODIGO, DESCRIPCION, CODINT, GRUCOVEN "
   SQLX$ = SQLX$ + " FROM Master WITH(NOLOCK) where status>0"
   SQLX$ = SQLX$ + " ORDER BY CODIGO ASC "
   Dim MASTERECO As ADODB.Recordset
   Set MASTERECO = New ADODB.Recordset
   MASTERECO.CursorLocation = adUseClient
   MASTERECO.CursorType = adOpenForwardOnly
   MASTERECO.LockType = adLockReadOnly
   MASTERECO.Open SQLX$, FLEXCONN, , , adCmdText
'   Set MASTERECO = New ADODB.Recordset
'   MASTERECO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'   Set MASTERECO = READSET(SQLX$)
   
'   On Error Resume Next
'   MASTERECO.MoveLast
'   If Err Then On Error GoTo 0: Exit Sub
'   '
'   CAREC% = MASTERECO.RecordCount

   If CAREC& < 1 Then Exit Sub
    '
   'FIN& = CAREC%
   'For IPIE& = 1 To FIN&
   IPIE& = 0
   Do While Not MASTERECO.EOF
     IPIE& = IPIE& + 1
     Call TRACE(24, IPIE&, CAREC&)
     'CI% = IPIE&
     CI% = XVALO(MASTERECO!CODINT)
     'CODD$ = CODEXT$(CI%)
     CODD$ = SAFETEXT$(MASTERECO!Codigo)
     'If STATUIT%(CI%) > 0 Then
       If CI% > 0 Then
         If CI% <= NMM% Then
            CICON% = CI%
            '
            Label1 = CODD$
            'Label2 = DESCRIT$(CI%)
            Label2 = SAFETEXT$(MASTERECO!Descripcion)
            '
            'grcv% = GRUCOVEI%(CICON%)
            GRCV% = XVALO(MASTERECO!Grucoven)
            If GRCV% < 1 Or GRCV% > 255 Then GRCV% = 0
            Label241 = FORMATNUM$(PORGAGE(1, GRCV%), "F5.1")
            Label236 = FORMATNUM$(PORGAGE(2, GRCV%), "F5.1")
            Label238 = FORMATNUM$(PORGAGE(3, GRCV%), "F5.1")
            Label242 = FORMATNUM$(PORGAGE(4, GRCV%), "F4.1")
            Label243 = FORMATNUM$(PORGAGE(5, GRCV%), "F4.1")
            Label239 = FORMATNUM$(PORGAGE(6, GRCV%), "F4.1")
            Label244 = FORMATNUM$(PORGAGE(7, GRCV%), "F4.1")
            '
            For KKII% = 0 To 7
              MODOCALCU%(KKII%) = MODOCA%(KKII%, GRCV%)
            Next KKII%
            '
            Call EXPLOMUL(CI%)
            SUMPPROP = 0: SUMPREC = 0
            SUPIPROP = 0: SUPIREC = 0
            SUSEPROP = 0: SUSEREC = 0
            SUMODPRO = 0: SUMODREC = 0
            SUMAQPRO = 0: SUMAQREC = 0
            '
            SUPRI = 0: SUPIE = 0: SUSER = 0
            SUTIE = 0: SUMOD = 0: SUMAQ = 0
            '
            ' costo de materias primas y piezas de terceros
            '
            For KKI% = 1 To CONTAMUL
              If DP%(CIK%(KKI%)) < 1 Then   ' NO TIENE ESTRUCTURA
                '
                COSUN = COSUMAT(CIK%(KKI%))
                UNIME$ = UMEDIX$(CIK%(KKI%))
                '
16              If Left$(UNIME$, 1) = "U" Or Left$(UNIME$, 2) = "PZ" Then
                     If KKI% = 1 Then
                          SUPIPROP = SUPIPROP + CANT(KKI%) * COSUN
                        Else
                          SUPIREC = SUPIREC + CANT(KKI%) * COSUN
                     End If
                     SUPIE = SUPIE + CANT(KKI%) * COSUN
                   Else
                     If KKI% = 1 Then
                          SUMPPROP = SUMPPROP + COSUN * CANT(KKI%)
                        Else
                          SUMPREC = SUMPREC + COSUN * CANT(KKI%)
                     End If
                     SUPRI = SUPRI + CANT(KKI%) * COSUN
                End If
                '
              Else
                '
                If COSUMAT(CIK%(KKI%)) > 0 Then
                  PHB% = PRHABI%(CIK%(KKI%))
                  If PHB% > 0 Then
                    If PROVINT%(PHB%) = 0 Then
                      COSUN = COSUMAT(CIK%(KKI%))
                      '
                      If KKI% = 1 Then
                           SUSEPROP = SUSEPROP + COSUN * CANT(KKI%)
                         Else
                           SUSEREC = SUSEREC + CANT(KKI%) * COSUN
                      End If
                      SUSER = SUSER + CANT(KKI%) * COSUN
                    End If
                  End If
                End If
                '
              End If
            Next KKI%
            '
            '
            ' costo de mano de obra directa
            '
            FIXX& = CONTAMUL%
            For UI& = 1 To FIXX&
               '
               CIXI% = CIK%(UI&)
               CACOM = CANT(UI&)
               '
               COSMOD = CACOM * COCAMOD(CIXI%)
               COSMAQ = CACOM * COCAMAQ(CIXI%)
               '
               If UI& = 1 Then
                    SUMODPRO = SUMODPRO + COSMOD
                    SUMAQPRO = SUMAQPRO + COSMAQ
                  Else
                    SUMODREC = SUMODREC + COSMOD
                    SUMAQREC = SUMAQREC + COSMAQ
               End If
               '
            Next UI&
            '
40          GoSub 90
            '
            'PRELISX# = PRELISTA#("LIPRE001", CICON%) * TICAMONE(MONELISTA%(1))
            PRELISX# = XVALO(VALOBADA("preventa", "Precio_Unitario", "cod_lista=1 and cod_pieza=" & CICON%)) * TICAMONE(MONELISTA%(1))
            If PRELISX# <= 0 And CONTROL$("ANACOSTO", "PRELISTA") = "SI" Then
               FIN& = CantRegistros&("LISTAPRE", "Cod_Lista > 0 ", "NOMESS")
               For H% = 2 To FIN&
                 NMBR$ = TRIM$(Str$(H%))
                 PRELISX# = PRELISTA#(NMBR$, CI%)
                 PRELISX# = PRELISX# * TICAMONE(MONELISTA%(H%))
                 If PRELISX# > 0.00005 Then Exit For
               Next H%
            End If

            Label114.Caption = TRIM$(CSTRING$(MKD$(PRELISX#), 3, 12, DECICALCOS%, 2))
            DoEvents
            PREVENETO# = XVALO(Label112)
            '
            Call LABEL241_Change
            Call LABEL236_Change
            Call LABEL238_Change
            Call LABEL242_Change
            Call LABEL243_Change
            Call LABEL239_Change
            DoEvents
            '
            RETRA& = CICON%
            YY$ = REGLEIDO$("TRACALCO", RETRA&)
            Call REPLA(YY$, MKD$(Val(Label90) + Val(Label82)), 9, 8)
            Call REPLA(YY$, MKD$(Val(Label88) + Val(Label83)), 17, 8)
            Call REPLA(YY$, MKD$(Val(Label86) + Val(Label84)), 25, 8)
            Call REPLA(YY$, MKD$(Val(Label85)), 33, 8)
            Call REPLA(YY$, MKD$(Val(Label75) + Val(Label72)), 41, 8)
            Call REPLA(YY$, MKD$(Val(Label79) + Val(Label74)), 49, 8)
            Call REPLA(YY$, MKD$(0), 57, 8)
            Call REPLA(YY$, MKD$(Val(Label80)), 65, 8)
            Call REPLA(YY$, MKD$(Val(Label70)), 73, 8)
            Call REPLA(YY$, MKD$(Val(Label65)), 81, 8)
            Call REPLA(YY$, MKD$(Val(Label96)), 89, 8)
            Call REPLA(YY$, MKD$(Val(Label103)), 97, 8)
            Call REPLA(YY$, MKD$(Val(Label112)), 105, 8)
            Call REPLA(YY$, MKD$(Val(Label116)), 113, 8)
            Call REPLA(YY$, MKS$(Val(Label117)), 121, 4)
            Call REPLA(YY$, MKS$(Val(Label119)), 125, 4)
            Call GRAREG("TRACALCO", YY$, RETRA&)
            '\\\OT-06-0299-WAR-17/07/06///
            If HAYRECON% > 0 Then
               ZZ$ = REGLEIDO$("TRCOSCON", RETRA&)
               Call REPLA(ZZ$, MKD$(Val(Label70)), 9, 8)
               Call REPLA(ZZ$, MKD$(Val(Label104)), 17, 8)
                 GSINDI1# = XVALO(Label93) + XVALO(Label68) + XVALO(Label77) + XVALO(Label110)
               Call REPLA(ZZ$, MKD$(GSINDI1#), 25, 8)
                 GSNAZ1# = 0 'GS.NACIONALIZACION
                 MONECO1% = MONECOSTO(CICON%)
                 GSNAZ1# = (COSUNI(CICON%) * TICAMONE(MONECO1%) * PORNACIO(CICON%)) / 100
               Call REPLA(ZZ$, MKD$(GSNAZ1#), 33, 8)
                 TOTA1# = XVALO(Label70) + XVALO(Label104) + GSINDI1# + GSNAZ1#
               Call REPLA(ZZ$, MKD$(TOTA1#), 41, 8)
                 PRVTA1# = XVALO(Label112)
               Call REPLA(ZZ$, MKD$(PRVTA1#), 49, 8)
                 MARG1# = 0
'                 If PRVTA1# > 0 Then 'Evita División por Cero.
'                   marg1# = TOTA1# / PRVTA1#
'                 End If
                 ' modificado segun solic.Pablo Pierri 25/07/06
                 If TOTA1# > 0 Then 'Evita División por Cero.
                   MARG1# = 100 * (PRVTA1# / TOTA1# - 1)
                 End If
               Call REPLA(ZZ$, MKD$(MARG1#), 57, 8)
               Call GRAREG("TRCOSCON", ZZ$, RETRA&)
            End If
            '\\\OT-06-0299-WAR-FIN///
            '
         End If
       End If
     'End If
     '
     Label16(0) = TRIM$(Str$(FIN&))
     Label16(1) = TRIM$(Str$(IPIE&))
     Label16(2) = TRIM$(Str$(FIN& - IPIE&))
     Label10(0) = TRIM$(HORMINSE$((Timer - TINI1) * FIN& / IPIE&))
     Label10(1) = TRIM$(HORMINSE$(Timer - TINI1))
     Label10(2) = TRIM$(HORMINSE$((Timer - TINI1) * (FIN& / IPIE& - 1)))
     '
     'DoEvents
     '
     While PAUSA% = 1
     DoEvents
     Wend
     '
     'GRABA COSTO CALCULADO
     SQLX$ = "select * from master where codint = " & CI%
     Dim Rstactucos As ADODB.Recordset
     Set Rstactucos = New ADODB.Recordset
     Rstactucos.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     If Not (Rstactucos.BOF And Rstactucos.EOF) Then
        Rstactucos!COSCAL = XVALO(Label103)
        Rstactucos!FECOSCAL = CVDAT(HOY(HO$))
        Rstactucos.Update
     End If
     Rstactucos.Close
     Set Rstactucos = Nothing
     'Call ACTUREGISTRO("Master", "CosCal", CONDI$, XVALO(Label103), REG&)
     'Call ACTUREGISTRO("Master", "FeCosCal", CONDI$, FETEXCOR1$(HOY(HOS$)), REG&)
'   Next IPIE&
     MASTERECO.MoveNext
   Loop
   '
   Screen.MousePointer = 11
   FEX = HOY(HOS$)
   Call GRACONTROL("ANACOSTO", "FECALCOS", HOS$)
   Call GRACONTROL("ANACOSTO", "HOCALCOS", Left$(Time$, 5))
   '
   ANACOSTO07.Label34(4).Caption = HOS$
   ANACOSTO07.Label34(4).Refresh
   ANACOSTO07.Label34(5).Caption = Left$(Time$, 5)
   ANACOSTO07.Label34(5).Refresh
   '
   FIN& = ULTREG&("TRACALCO")
   For i& = 1 To FIN&
     Call TRACE(20, i&, FIN&)
     XX$ = REGLEIDO$("TRACALCO", i&)
     Call GRAREG("RECOSCAL", XX$, i&)
   Next i&
   Call SETULTREG("RECOSCAL", FIN&)
   Call CIERRARCH("RECOSCAL")
   '\\\OT-06-0299-WAR-17/07/06///
   If HAYRECON% > 0 Then
      FIN& = ULTREG&("TRCOSCON")
      For i& = 1 To FIN&
        Call TRACE(20, i&, FIN&)
        XX$ = REGLEIDO$("TRCOSCON", i&)
        Call GRAREG("RECOSCON", XX$, i&)
      Next i&
      Call SETULTREG("RECOSCON", FIN&)
      Call CIERRARCH("RECOSCON")
   End If
   '\\\OT-06-0299-WAR-FIN///
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   
   ' GUARDA EL RESULTADO DEL CALCULO ALMACENADO EN RECOSCAL EN TABLA HISTOCOSTOS
   Call FRESHECHO("SELMESA")
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
     MESAN$ = MESASEL$
   End If
   Call MESANO$(MESAN$, Des%, Has%)
   '
   FF1$ = FECHATEXLAR$(Des%)
   MESHOY& = Val(Mid$(FF1$, 4, 2))
   ANIOHOY& = Val(Mid$(FF1$, 7, 4))
   NN% = InStr(1, MESAN, " ")
   DENMES$ = TRIM$(Mid$(MESAN, 1, NN%))
   'DENMES$ = UCase(TRIM$(Mid$(Format(CVDAT(FF%), "DD-MMMM"), 4)))
   IDMESANIO& = (MESHOY& * 10000) + ANIOHOY&
   '
   FechaHoraDeHoy = Now
   '
   Dim RST_HISTO As ADODB.Recordset
   Set RST_HISTO = New ADODB.Recordset
   SQLX$ = "SELECT * FROM HISTOCOSTOS WHERE IDPERIODO=" & Str$(IDMESANIO&)
   RST_HISTO.Open SQLX$, FLEXCONN, adOpenKeyset, adLockOptimistic, adCmdText
   '
   Screen.MousePointer = 11
   MESCAL$ = DENMES$ & " " & Mid$(FF1$, 9, 2)
   With RST_HISTO
        FIN& = ULTREG&("RECOSCAL")
        For II& = 1 To FIN&
            Call TRACE(24, II&, FIN&)
            '
            XX$ = REGLEIDO$("RECOSCAL", II&)
            .AddNew
            !PERIODO = MESCAL$
            !IdPeriodo = IDMESANIO&
            !FECOSCAL = FechaHoraDeHoy
            !COD_INTE = CVI(Mid$(XX$, 1, 2))
            !TIMATE = CVI(Mid$(XX$, 3, 2))
            !UMEDIDA = Mid$(XX$, 5, 2)
            !MoneCalCo = CVI(Mid$(XX$, 7, 2))
            !costomp = CVD(Mid$(XX$, 9, 8))
            !Piezas = CVD(Mid$(XX$, 17, 8))
            !TTerceros = CVD(Mid$(XX$, 25, 8))
            !TotaSuminis = CVD(Mid$(XX$, 33, 8))
            !ManoObra = CVD(Mid$(XX$, 41, 8))
            !Equipos = CVD(Mid$(XX$, 49, 8))
            !Otros = CVD(Mid$(XX$, 57, 8))
            !TotElaboracion = CVD(Mid$(XX$, 65, 8))
            !TotCostoDirecto = CVD(Mid$(XX$, 73, 8))
            !CostoStandard = CVD(Mid$(XX$, 81, 8))
            !CostoStdTotal = CVD(Mid$(XX$, 89, 8))
            !PrecioVenta = CVD(Mid$(XX$, 97, 8))
            !PRECIOLISTA = CVD(Mid$(XX$, 105, 8))
            !Margen = CVD(Mid$(XX$, 113, 8))
            !PorcentajeMBruto = CVS(Mid$(XX$, 121, 4))
            !IndiceCalidad = CVS(Mid$(XX$, 125, 4))
            .Update
       Next II&
   End With
   On Error Resume Next
   RST_HISTO.Close
   Set RST_HISTO = Nothing
   On Error GoTo 0
   '
   Screen.MousePointer = 1
   '
   FINCALCOS% = 1
   Call COMUNI("Re-Cálculo Completo")
   '
   Call FINAL("")
   Screen.MousePointer = 1
   '
Exit Sub

90 Label90.Caption = TRIM$(CSTRING$(MKS$(SUMPPROP), 3, 12, DECICALCOS%, 2))
   Label90.Refresh
   Label82.Caption = TRIM$(CSTRING$(MKS$(SUMPREC), 3, 12, DECICALCOS%, 2))
   Label82.Refresh
   Label88.Caption = TRIM$(CSTRING$(MKS$(SUPIPROP), 3, 12, DECICALCOS%, 2))
   Label88.Refresh
   Label83.Caption = TRIM$(CSTRING$(MKS$(SUPIREC), 3, 12, DECICALCOS%, 2))
   Label83.Refresh
   Label86.Caption = TRIM$(CSTRING$(MKS$(SUSEPROP), 3, 12, DECICALCOS%, 2))
   Label86.Refresh
   Label84.Caption = TRIM$(CSTRING$(MKS$(SUSEREC), 3, 12, DECICALCOS%, 2))
   Label84.Refresh
   '
   Label75.Caption = TRIM$(CSTRING$(MKS$(SUMODPRO), 3, 12, DECICALCOS%, 2))
   Label75.Refresh
   Label72.Caption = TRIM$(CSTRING$(MKS$(SUMODREC), 3, 12, DECICALCOS%, 2))
   Label72.Refresh
   Label79.Caption = TRIM$(CSTRING$(MKS$(SUMAQPRO), 3, 12, DECICALCOS%, 2))
   Label79.Refresh
   Label74.Caption = TRIM$(CSTRING$(MKS$(SUMAQREC), 3, 12, DECICALCOS%, 2))
   Label74.Refresh
   '
   SUMA1 = 0
   SUMA1 = SUMA1 + Val(Label90.Caption)
   SUMA1 = SUMA1 + Val(Label88.Caption)
   SUMA1 = SUMA1 + Val(Label86.Caption)
   SUMA1 = SUMA1 + Val(Label82.Caption)
   SUMA1 = SUMA1 + Val(Label83.Caption)
   SUMA1 = SUMA1 + Val(Label84.Caption)
   Label85.Caption = TRIM$(CSTRING$(MKS$(SUMA1), 3, 12, DECICALCOS%, 2))
   Label85.Refresh
   '
   SUMA2 = 0
   SUMA2 = SUMA2 + Val(Label75.Caption)
   SUMA2 = SUMA2 + Val(Label72.Caption)
   SUMA2 = SUMA2 + Val(Label79.Caption)
   SUMA2 = SUMA2 + Val(Label74.Caption)
   Label80.Caption = TRIM$(CSTRING$(MKS$(SUMA2), 3, 12, DECICALCOS%, 2))
   Label80.Refresh
   '
   SUMA3 = SUMA1 + SUMA2
   Label70.Caption = TRIM$(CSTRING$(MKS$(SUMA3), 3, 12, DECICALCOS%, 2))
   Label70.Refresh
   '
Return
   
End Sub
'
Sub CALCOSMOD(CIXI%)
    '
    CACOM = 1
    '
    SQLX$ = "SELECT * FROM Secoper WITH(NOLOCK) "
    SQLX$ = SQLX$ & " WHERE CodIConj = " & CIXI% & " "
    SQLX$ = SQLX$ & "AND StatOper > 0 "
    'Set SECOPTEMP = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim SECOPTEMP As ADODB.Recordset
    Set SECOPTEMP = New ADODB.Recordset
    SECOPTEMP.CursorLocation = adUseClient
    SECOPTEMP.CursorType = adOpenForwardOnly
    SECOPTEMP.LockType = adLockReadOnly
    SECOPTEMP.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
    'Set SECOPTEMP = READSET(SQLX$)
    '
    SUCMOD = 0: SUCMAQ = 0
    On Error Resume Next
    With SECOPTEMP
      .MoveFirst
      If Err Then GoTo 25
      '
      Do Until .EOF = True
        If IsNull(!AltOper) = True Or !AltOper < 1 Then       '
           NOPY% = XVALO(!NUMEOPER)
           DOPE$ = SAFETEXT$(!descoper)
           LOTESTA = LOTESTAN(CIXI%)
              If LOTESTA < 1 Then LOTESTA = 1
           CICLOHE = XVALO(!PieCiclo)
              If CICLOHE < 1 Then CICLOHE = 1
           MINPREP = 60 * XVALO(!HorsPrep)
           MINFIJO = 60 * XVALO(!HorsFijo)
           MINVARI = LOTESTA * XVALO(!MinuStan)
           MINHERR = LOTESTA * 60 * XVALO(!HorsMHer) / CICLOHE
           PIECICL = XVALO(!PieCiclo)
           HORLIMP = XVALO(!HorsLimp)
           DEMODIS = XVALO(!PorDemor)
           MINDEMO = (MINFIJO + MINVARI + MINHERR) * DEMODIS / 100
           MINLOTE = MINFIJO + MINVARI + MINHERR + MINDEMO
           HORAPROD = CACOM * MINLOTE / 60 / LOTESTA
           '
           SOLAPAI% = XVALO(!CoefSolp)
           '
           CAOPS = XVALO(!CantiOps)
           CATOP% = XVALO(!CATEGOPS)
           EQUIPO$ = SAFETEXT$(!CodMaq0)
           EQUIP1$ = SAFETEXT$(!CodMaq1)
           EQUIP2$ = SAFETEXT$(!CodMaq2)
           EQUIP3$ = SAFETEXT$(!CodMaq3)
           CUOTOPE = VALHORA(CATOP%)
           CUOTMAQ = CUHORMAQ(EQUIPO$)
           '
           PROTERI% = XVALO(!ProTerOp)  ' PROVEEDOR OPERACION TERCERIZADA
           PRECTER# = XVALO(!PreUnid)
           MONETER% = XVALO(!MonePrec)
           FEPRETER% = FECHANUM(CVINT%(!FechPrec))
           '
           EQUIPO$ = TRIM$(Mid$(SCP$, 53, 6))
           CAOPREP = CANOPREP(CIXI%, NOPY%)
           If CAOPREP < 1 Then CAOPREP = CAOPS
           CTOPREP% = CATOPREP%(CIXI%, NOPY%)
           If CTOPREP% < 1 Then CTOPREP% = CATOP%
           CUOTOPREP = VALHORA(CTOPREP%)
           '
           COSMOD = CUOTOPE * CAOPS * HORAPROD + CUOTOPREP * CAOPREP * MINPREP / 60 / LOTESTA
           COSMAQ = CUOTMAQ * (HORAPROD + MINPREP / 60 / LOTESTA)
           '
           TERCOP% = 0
           MONEI% = XVALO(!MonePrec)  'MONEI% = MONETER%
           COSTER = CACOM * XVALO(!PreUnid) * TICAMONE(MONEI%)  'COSTER = CACOM * PRECTER# * TICAMONE(MONEI%)
           If MONEI% <= 1 Then
             COSTER = CACOM * XVALO(!PreUnid) * IAJUINF#(FEBAS%, FEACT%)
           End If
           If COSTER >= 0.00005 Then
             If (COSTER < COSMOD + COSMAQ) Or ((COSMOD + COSMAQ) < 0.00005) Then
               COSMOD = COSTER
               COSMAQ = 0
               HORAPROD = 0
               TERCOP% = 1
             End If
           End If
           '
           SUCMOD = SUCMOD + COSMOD
           SUCMAQ = SUCMAQ + COSMAQ
        End If
      .MoveNext
      Loop
      '
25  End With
    '
    SECOPTEMP.Close
    Set SECOPTEMP = Nothing
    '
    COCAMOD(CIXI%) = SUCMOD
    COCAMAQ(CIXI%) = SUCMAQ
    '
End Sub
'
Sub CALCOSMAT(CIXI%)
    '
    SQLX$ = "select codint, monecos, cosuni, pornacio, fecosto, umedida, provhabi from master with(nolock) where codint=" & CIXI%
    Set Rst = READSET(SQLX$)
    With Rst
        'MONEI% = MONECOSTO%(CIXI%)
        'COSUDOL = COSUNI(CIXI%)
        MONEI% = XVALO(!Monecos)
        COSUDOL = XVALO(!COSUNI)
        If MONEI% > 1 Then
             COSUMAT(CIXI%) = COSUDOL * TICAMONE(MONEI%) * (1 + XVALO(!PORNACIO) / 100)
        Else
             'FEBAS% = FECOSTO%(CIXI%)
             FEBAS% = CVINT(!FECOSTO)
             FEACT% = HOY(HOS$)
             AA = IAJUINF#(FEBAS%, FEACT%)
             COSUN = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
             COSUMAT(CIXI%) = COSUN
        End If
        'COSUMAT(CIXI%) = COSUNI(CIXI%) * TICAMONE(MONEI%) * (1 + (PORNACIO(CIXI%)) / 100)
        'UMEDIX$(CIXI%) = UCase$(UNIMED$(CIXI%))
        UMEDIX$(CIXI%) = UCase$(SAFETEXT$(!UMEDIDA))
        '
        'PRHABI%(CIXI%) = PROVHABI%(CIXI%)
        PRHABI%(CIXI%) = XVALO(!PROVHABI)
    End With
End Sub
'
Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Command2_Click()
   If PAUSA% = 0 Then
       PAUSA% = 1
       Command2.ToolTipText = "Click para Continuar"
     Else
       PAUSA% = 0
       Command2.ToolTipText = "Click para Interrumpir"
   End If
End Sub

Private Sub Form_Load()
    '
    ENPRORECA% = 1    ' en proceso de recalculo
    '
    If RECALCOSTO% <> 1 Then Exit Sub
    '
    RECALCOSTO% = 0
    Top = ANACOSTO07.Top + ANACOSTO07.SSTab1.Top + 520
    Left = ANACOSTO07.Left + ANACOSTO07.SSTab1.Left
    Call APLICACIONSKINS(Me)

    Show
'    EPAC$ = TRIM$(EMPREAC$)
'    If EPAC$ <> "" Then
'        Caption = Caption + "  -  " + EPAC$
'    End If
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    DoEvents
    '
    '
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "GAGENMAT"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(1, KU&) = GGG: MODOCA%(1, KU&) = 1: Next
      Label241 = CSTRING$(MKS$(Val(DCC$)), 3, 5, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "GAGENFAB"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(2, KU&) = GGG: MODOCA%(2, KU&) = 1: Next
      Label236 = CSTRING$(MKS$(Val(DCC$)), 3, 5, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "GAGENADM"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(3, KU&) = GGG: MODOCA%(3, KU&) = 1: Next
      Label238 = CSTRING$(MKS$(Val(DCC$)), 3, 5, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "PORCOMIS"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(4, KU&) = GGG: MODOCA%(4, KU&) = 1: Next
      Label242 = CSTRING$(MKS$(Val(DCC$)), 3, 4, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "PORIMPUE"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(5, KU&) = GGG: MODOCA%(5, KU&) = 1: Next
      Label243 = CSTRING$(MKS$(Val(DCC$)), 3, 4, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "PORUTILI"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(6, KU&) = GGG: MODOCA%(6, KU&) = 1: Next
      Label239 = CSTRING$(MKS$(Val(DCC$)), 3, 4, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "PORDESCU"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(7, KU&) = GGG: MODOCA%(7, KU&) = 1: Next
      Label244 = CSTRING$(MKS$(Val(DCC$)), 3, 4, 1, 2)
      '
      If EXISTE%(LUDAT$ + "GAGENCOS.DAT") > 0 Then
         For KU& = 1 To ULTREG&("GAGENCOS")
           XXX$ = REGLEIDO$("GAGENCOS", KU&)
           GRGR% = CVI(Left$(XXX$, 2))
           If GRGR% >= 0 Then
             If GRGR% <= 255 Then
               NORI% = CVI(Mid$(XXX$, 3, 2))
               If NORI% >= 0 Then
                 If NORI% <= 7 Then
                   PORGAGE(NORI%, GRGR%) = CVS(Mid$(XXX$, 37, 4))
                   MODOCA%(NORI%, GRGR%) = CVI(Mid$(XXX$, 41, 2))
                 End If
               End If
             End If
           End If
         Next KU&
      End If
      '
   For UI& = 0 To 32767
     PROVINT%(UI&) = 0
   Next UI&
   '
   For U& = 1 To ULTREG&("TAPROINT")
      x$ = REGLEIDO$("TAPROINT", U&)
      NP% = CVI(Left$(x$, 2))
      If NP% > 0 Then
        PROVINT%(NP%) = 1
      End If
   Next U&
   Call CIERRARCH("TAPROINT")
   '
   DoEvents
   'OPX% = ALTERNA%("Proceso VIEJO\Proceso NUEVO")
   'If OPX% = 1 Then
   '      Call RECACOS
   '   ElseIf OPX% = 2 Then
         Call RECACOS_NUE
   'End If
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Screen.MousePointer = 1
   ENPRORECA% = 0
   Cancel = 1
   If FINCALCOS% <> 1 Then
     If COMALTER%("Realmente Desea Cancelar\el Proceso de Cálculo de Costos ?\\No, Continuar\Si, Cancelar") <> 2 Then
        Exit Sub
     End If
   End If
   Cancel = 0
   On Error Resume Next
   ANACOSTO07.Command55.Enabled = True
   ANACOSTO07.Text1.SetFocus
   On Error GoTo 0
   DoEvents
   Call FINAL("")
End Sub




Private Sub Label103_Change()
   If MODOCALCU%(4) = 1 Then Label110 = TRIM$(CSTRING$(MKD$(Val(Label103) * Val(Label242) / 100), 3, 12, DECICALCOS%, 2))
   If MODOCALCU%(5) = 1 Then Label104 = TRIM$(CSTRING$(MKD$(Val(Label103) * Val(Label243) / 100), 3, 12, DECICALCOS%, 2))
   If MODOCALCU%(6) = 1 Then Label107 = TRIM$(CSTRING$(MKD$(Val(Label103) * Val(Label239) / 100), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label104_Change()
   Label111 = TRIM$(CSTRING$(MKD$(Val(Label105) + Val(Label110) + Val(Label104)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label105_Change()
   PREVEN# = 0
   If (1 - (Val(Label239) + Val(Label242) + Val(Label243)) / 100) > 0.0001 Then
     PREVEN# = Val(Label105) / (1 - (Val(Label239) + Val(Label242) + Val(Label243)) / 100)
   End If
   Label103 = TRIM$(CSTRING$(MKD$(PREVEN#), 3, 12, DECICALCOS%, 2))
   Label111 = TRIM$(CSTRING$(MKD$(Val(Label105) + Val(Label110) + Val(Label104)), 3, 12, DECICALCOS%, 2))
   Label116 = ""
   If Val(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(Val(Label112) * (1 - Val(Label242) / 100 - Val(Label243) / 100) - Val(Label105)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label110_Change()
   Label111 = TRIM$(CSTRING$(MKD$(Val(Label105) + Val(Label110) + Val(Label104)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label111_Change()
   Label116 = ""
   If Val(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(Val(Label112) - Val(Label111)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label112_Change()
   Label116 = ""
   If Val(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(Val(Label112) * (1 - Val(Label242) / 100 - Val(Label243) / 100) - Val(Label105)), 3, 12, DECICALCOS%, 2))
   Label117 = ""
   If Abs(Val(Label112)) > 0.005 Then
     PORMA# = 100 * Val(Label116) / Val(Label112)
     If Abs(PORMA#) <= 99.9 Then
       Label117 = TRIM$(CSTRING$(MKD$(PORMA#), 3, 5, 1, 2))
     End If
   End If
   Label117.ForeColor = RGB(0, 0, 0)
   If Val(Label117) < 0 Then Label117.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Label114_Change()
   Label115 = TRIM$(CSTRING$(MKD$(Val(Label114) * Val(Label244) / 100), 3, 12, DECICALCOS%, 2))
   Label112 = TRIM$(CSTRING$(MKD$(Val(Label114) - Val(Label115)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label115_Change()
   Label116 = ""
   If Val(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(Val(Label114) - Val(Label115)), 3, 12, DECICALCOS%, 2))
   Label112 = TRIM$(CSTRING$(MKD$(Val(Label114) - Val(Label115)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label116_Change()
   Label117 = ""
   If Abs(Val(Label112)) > 0.005 Then
     PORMA# = 100 * Val(Label116) / Val(Label112)
     If Abs(PORMA#) <= 99.9 Then
       Label117 = TRIM$(CSTRING$(MKD$(PORMA#), 3, 5, 1, 2))
     End If
   End If
   Label117.ForeColor = RGB(0, 0, 0)
   If Val(Label117) < 0 Then Label117.ForeColor = RGB(255, 0, 0)
   '
End Sub

Private Sub Label117_Change()
   Label119 = ""
   If Val(Label239) > 0.005 Then
     PORMA# = 100 * Val(Label117) / Val(Label239)
     If Abs(PORMA#) <= 999.9 Then
       Label119 = TRIM$(CSTRING$(MKD$(PORMA#), 3, 5, 1, 2))
     End If
   End If
   Label119.ForeColor = RGB(0, 0, 0)
   If Val(Label119) < 100 Then Label119.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Label65_Change()
   Label77 = TRIM$(CSTRING$(MKD$(Val(Label65) * Val(Label238) / 100), 3, 12, DECICALCOS%, 2))
   Label96 = TRIM$(CSTRING$(MKD$(Val(Label65) + Val(Label77)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label68_Change()
   Label65 = TRIM$(CSTRING$(MKD$(Val(Label70) + Val(Label68) + Val(Label93)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label70_Change()
   Label65 = TRIM$(CSTRING$(MKD$(Val(Label70) + Val(Label68) + Val(Label93)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label77_Change()
   Label96 = TRIM$(CSTRING$(MKD$(Val(Label65) + Val(Label77)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label80_Change()
   Label68 = TRIM$(CSTRING$(MKD$(Val(Label80) * Val(Label236) / 100), 3, 12, DECICALCOS%, 2))
   Label68.Refresh
End Sub

Private Sub Label80_DBLClick()
   Label80.Enabled = False
   If ANACOSMOD07.List1.ListCount > 0 Then
     ANACOSMOD07.Show 1
   End If
   Label80.Enabled = True
End Sub

Private Sub Label85_Change()
   Label93 = TRIM$(CSTRING$(MKD$(Val(Label85) * Val(Label241) / 100), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label85_dblClick()
   Label85.Enabled = False
   If ANACOSMAT07.List1.ListCount > 0 Then
     ANACOSMAT07.Show 1
   End If
   Label85.Enabled = True
End Sub

Private Sub Label93_Change()
   Label65 = TRIM$(CSTRING$(MKD$(Val(Label70) + Val(Label68) + Val(Label93)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label96_Change()
   Label105 = Label96
End Sub


Private Sub LABEL236_Change()
   'Label68 = TRIM$(CSTRING$(MKD$(Val(Label80) * Val(Label236) / 100), 3, 12, DECICALCOS%, 2))
   If MODOCALCU%(2) = 1 Then
        Label68 = TRIM$(CSTRING$(MKD$(Val(Label80) * Val(Label236) / 100), 3, 12, DECICALCOS%, 2))
      Else
        Label68 = TRIM$(CSTRING$(MKD$(PREVENETO# * XVALO(Label236) / 100), 3, 12, DECICALCOS%, 2))
   End If
End Sub

Private Sub LABEL238_Change()
   'Label77 = TRIM$(CSTRING$(MKD$(Val(Label65) * Val(Label238) / 100), 3, 12, DECICALCOS%, 2))
   If MODOCALCU%(3) = 1 Then
        Label77 = TRIM$(CSTRING$(MKD$(Val(Label65) * Val(Label238) / 100), 3, 12, DECICALCOS%, 2))
      Else
        Label77 = TRIM$(CSTRING$(MKD$(PREVENETO# * XVALO(Label238) / 100), 3, 12, DECICALCOS%, 2))
   End If
End Sub

Private Sub LABEL239_Change()
   PREVEN# = 0
   If (1 - (Val(Label239) + Val(Label242) + Val(Label243)) / 100) > 0.0001 Then
     PREVEN# = Val(Label105) / (1 - (Val(Label239) + Val(Label242) + Val(Label243)) / 100)
   End If
   Label103 = TRIM$(CSTRING$(MKD$(PREVEN#), 3, 12, DECICALCOS%, 2))
   'Label107 = TRIM$(CSTRING$(MKD$(PREVEN# * Val(Label239) / 100), 3, 12, DECICALCOS%, 2))
   '
   If MODOCALCU%(6) = 1 Then
        Label107 = TRIM$(CSTRING$(MKD$(PREVEN# * Val(Label239) / 100), 3, 12, DECICALCOS%, 2))
      Else
        Label107 = TRIM$(CSTRING$(MKD$(PREVENETO# * XVALO(Label239) / 100), 3, 12, DECICALCOS%, 2))
   End If
   '
   Label119 = ""
   If Val(Label239) > 0.005 Then
     PORMA# = 100 * Val(Label117) / Val(Label239)
     If Abs(PORMA#) <= 999.9 Then
       Label119 = TRIM$(CSTRING$(MKD$(PORMA#), 3, 5, 1, 2))
     End If
   End If
   Label119.ForeColor = RGB(0, 0, 0)
   If Val(Label119) < 100 Then Label119.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub LABEL241_Change()
   'Label93 = TRIM$(CSTRING$(MKD$(Val(Label85) * Val(Label241) / 100), 3, 12, DECICALCOS%, 2))
   If MODOCALCU%(1) = 1 Then
        Label93 = TRIM$(CSTRING$(MKD$(Val(Label85) * Val(Label241) / 100), 3, 12, DECICALCOS%, 2))
     Else
        Label93 = TRIM$(CSTRING$(MKD$(PREVENETO# * XVALO(Label241) / 100), 3, 12, DECICALCOS%, 2))
   End If
End Sub

Private Sub LABEL242_Change()
   PREVEN# = 0
   If (1 - (Val(Label239) + Val(Label242) + Val(Label243)) / 100) > 0.0001 Then
     PREVEN# = Val(Label105) / (1 - (Val(Label239) + Val(Label242) + Val(Label243)) / 100)
   End If
   Label103 = TRIM$(CSTRING$(MKD$(PREVEN#), 3, 12, DECICALCOS%, 2))
   'Label110 = TRIM$(CSTRING$(MKD$(PREVEN# * Val(Label242) / 100), 3, 12, DECICALCOS%, 2))
   '
   If MODOCALCU%(4) = 1 Then
        Label110 = TRIM$(CSTRING$(MKD$(PREVEN# * Val(Label242) / 100), 3, 12, DECICALCOS%, 2))
      Else
        Label110 = TRIM$(CSTRING$(MKD$(PREVENETO# * XVALO(Label242) / 100), 3, 12, DECICALCOS%, 2))
   End If
   '
   Label116 = ""
   If Val(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(Val(Label112) * (1 - Val(Label242) / 100 - Val(Label243) / 100) - Val(Label105)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub LABEL243_Change()
   PREVEN# = 0
   If (1 - (Val(Label239) + Val(Label242) + Val(Label243)) / 100) > 0.0001 Then
     PREVEN# = Val(Label105) / (1 - (Val(Label239) + Val(Label242) + Val(Label243)) / 100)
   End If
   Label103 = TRIM$(CSTRING$(MKD$(PREVEN#), 3, 12, DECICALCOS%, 2))
   'Label104 = TRIM$(CSTRING$(MKD$(PREVEN# * Val(Label243) / 100), 3, 12, DECICALCOS%, 2))
   '
   If MODOCALCU%(5) = 1 Then
        Label104 = TRIM$(CSTRING$(MKD$(PREVEN# * Val(Label243) / 100), 3, 12, DECICALCOS%, 2))
      Else
        Label104 = TRIM$(CSTRING$(MKD$(PREVENETO# * XVALO(Label243) / 100), 3, 12, DECICALCOS%, 2))
   End If
   '
   If Val(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(Val(Label112) * (1 - Val(Label242) / 100 - Val(Label243) / 100) - Val(Label105)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub LABEL244_Change()
   Label115 = TRIM$(CSTRING$(MKD$(Val(Label114) * Val(Label244) / 100), 3, 12, DECICALCOS%, 2))
End Sub

Sub CALCOSMOD_NUEVO()
    '
    CACOM = 1
    '
    SQLX$ = "SELECT * FROM Secoper WITH(NOLOCK) "
    'SQLX$ = SQLX$ & " WHERE CodIConj = " & CIXI% & " "
    SQLX$ = SQLX$ & "WHERE StatOper > 0 "
    'Set SECOPTEMP = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim SECOPTEMP As ADODB.Recordset
    Set SECOPTEMP = New ADODB.Recordset
    'SECOPTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    Set SECOPTEMP = READSET(SQLX$)
    '
    SUCMOD = 0: SUCMAQ = 0
    On Error Resume Next
    With SECOPTEMP
      .MoveFirst
      If Err Then GoTo 25
      '
      Do Until .EOF = True
        If IsNull(!AltOper) = True Or !AltOper < 1 Then       '
           NOPY% = !NUMEOPER
           DOPE$ = !descoper
           LOTESTA = LOTESTAN(CIXI%)
              If LOTESTA < 1 Then LOTESTA = 1
           CICLOHE = !PieCiclo
              If CICLOHE < 1 Then CICLOHE = 1
           MINPREP = 60 * !HorsPrep
           MINFIJO = 60 * !HorsFijo
           MINVARI = LOTESTA * !MinuStan
           MINHERR = LOTESTA * 60 * !HorsMHer / CICLOHE
           PIECICL = !PieCiclo
           HORLIMP = !HorsLimp
           DEMODIS = !PorDemor
           MINDEMO = (MINFIJO + MINVARI + MINHERR) * DEMODIS / 100
           MINLOTE = MINFIJO + MINVARI + MINHERR + MINDEMO
           HORAPROD = CACOM * MINLOTE / 60 / LOTESTA
           '
           SOLAPAI% = !CoefSolp
           '
           CAOPS = !CantiOps
           CATOP% = !CATEGOPS
           EQUIPO$ = !CodMaq0
           EQUIP1$ = !CodMaq1
           EQUIP2$ = !CodMaq2
           EQUIP3$ = !CodMaq3
           CUOTOPE = VALHORA(CATOP%)
           CUOTMAQ = CUHORMAQ(EQUIPO$)
           '
           PROTERI% = !ProTerOp  ' PROVEEDOR OPERACION TERCERIZADA
           PRECTER# = !PreUnid
           MONETER% = !MonePrec
           FEPRETER% = FECHANUM(Format$(!FechPrec, "dd/mm/yy"))
           '
           EQUIPO$ = TRIM$(Mid$(SCP$, 53, 6))
           CAOPREP = CANOPREP(CIXI%, NOPY%)
           If CAOPREP < 1 Then CAOPREP = CAOPS
           CTOPREP% = CATOPREP%(CIXI%, NOPY%)
           If CTOPREP% < 1 Then CTOPREP% = CATOP%
           CUOTOPREP = VALHORA(CTOPREP%)
           '
           COSMOD = CUOTOPE * CAOPS * HORAPROD + CUOTOPREP * CAOPREP * MINPREP / 60 / LOTESTA
           COSMAQ = CUOTMAQ * (HORAPROD + MINPREP / 60 / LOTESTA)
           '
           TERCOP% = 0
           MONEI% = !MonePrec  'MONEI% = MONETER%
           COSTER = CACOM * !PreUnid * TICAMONE(MONEI%)  'COSTER = CACOM * PRECTER# * TICAMONE(MONEI%)
           If MONEI% <= 1 Then
             COSTER = CACOM * !PreUnid * IAJUINF#(FEBAS%, FEACT%)
           End If
           If COSTER >= 0.00005 Then
             If (COSTER < COSMOD + COSMAQ) Or ((COSMOD + COSMAQ) < 0.00005) Then
               COSMOD = COSTER
               COSMAQ = 0
               HORAPROD = 0
               TERCOP% = 1
             End If
           End If
           '
           SUCMOD = SUCMOD + COSMOD
           SUCMAQ = SUCMAQ + COSMAQ
        End If
      .MoveNext
      Loop
      '
25  End With
    '
    SECOPTEMP.Close
    Set SECOPTEMP = Nothing
    '
    COCAMOD(CIXI%) = SUCMOD
    COCAMAQ(CIXI%) = SUCMAQ
    '
End Sub


