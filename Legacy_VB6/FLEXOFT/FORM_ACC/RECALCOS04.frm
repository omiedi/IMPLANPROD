VERSION 5.00
Begin VB.Form RECALCOS04 
   BackColor       =   &H00C0E0EE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Re-Cálculo Integral de Costos"
   ClientHeight    =   6150
   ClientLeft      =   165
   ClientTop       =   2205
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   10695
   Begin VB.Frame Frame16 
      BackColor       =   &H00C0E0EE&
      BorderStyle     =   0  'None
      Caption         =   "Frame16"
      Height          =   5895
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   10515
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
         BackColor       =   &H00C0E0EE&
         Caption         =   "Costo Total Calculado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         Begin VB.Label Label239 
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
            Left            =   5460
            TabIndex        =   87
            Top             =   1785
            Width           =   645
         End
         Begin VB.Label Label243 
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
            Left            =   5460
            TabIndex        =   86
            Top             =   1065
            Width           =   645
         End
         Begin VB.Label Label242 
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
            Left            =   5460
            TabIndex        =   85
            Top             =   690
            Width           =   645
         End
         Begin VB.Label Label236 
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
            Left            =   1680
            TabIndex        =   84
            Top             =   1065
            Width           =   750
         End
         Begin VB.Label Label238 
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
            Left            =   1680
            TabIndex        =   83
            Top             =   1815
            Width           =   750
         End
         Begin VB.Label Label241 
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
            Left            =   1680
            TabIndex        =   82
            Top             =   690
            Width           =   750
         End
         Begin VB.Label Label65 
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
            Left            =   2730
            TabIndex        =   81
            Top             =   1440
            Width           =   1275
         End
         Begin VB.Label Label66 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Costo Std de Manufactura:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
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
            Caption         =   "GG.Materiales:               %:"
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
            Left            =   315
            TabIndex        =   79
            Top             =   810
            Width           =   2430
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2730
            TabIndex        =   78
            Top             =   1065
            Width           =   1275
         End
         Begin VB.Label Label69 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Costos Directos:"
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
            Left            =   525
            TabIndex        =   77
            Top             =   420
            Width           =   2115
         End
         Begin VB.Label Label70 
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
            Left            =   2730
            TabIndex        =   76
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label77 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2730
            TabIndex        =   75
            Top             =   1815
            Width           =   1275
         End
         Begin VB.Label Label96 
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
            Left            =   2730
            TabIndex        =   74
            Top             =   2250
            Width           =   1275
         End
         Begin VB.Label Label97 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Gastos Grales Admin. y Ventas:"
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
            Left            =   105
            TabIndex        =   73
            Top             =   1785
            Width           =   1485
         End
         Begin VB.Label Label98 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Costo Standard Total:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -420
            TabIndex        =   72
            Top             =   2370
            Width           =   3060
         End
         Begin VB.Label Label99 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "%:"
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
            Left            =   2350
            TabIndex        =   71
            Top             =   1935
            Width           =   330
         End
         Begin VB.Label Label101 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Precio Básico de Venta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3360
            TabIndex        =   70
            Top             =   2370
            Width           =   2955
         End
         Begin VB.Label Label102 
            BackStyle       =   0  'Transparent
            Caption         =   " Impuestos:             %:"
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
            Left            =   4410
            TabIndex        =   69
            Top             =   1170
            Width           =   2115
         End
         Begin VB.Label Label103 
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
            Left            =   6405
            TabIndex        =   68
            Top             =   2250
            Width           =   1275
         End
         Begin VB.Label Label104 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6405
            TabIndex        =   67
            Top             =   1065
            Width           =   1275
         End
         Begin VB.Label Label105 
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
            Left            =   6405
            TabIndex        =   66
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label106 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Costo Standard Total:"
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
            Left            =   4200
            TabIndex        =   65
            Top             =   420
            Width           =   2115
         End
         Begin VB.Label Label107 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6405
            TabIndex        =   64
            Top             =   1815
            Width           =   1275
         End
         Begin VB.Label Label108 
            BackStyle       =   0  'Transparent
            Caption         =   " Utilidad:             %:"
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
            Left            =   4620
            TabIndex        =   63
            Top             =   1920
            Width           =   1905
         End
         Begin VB.Label Label109 
            BackStyle       =   0  'Transparent
            Caption         =   " Costo s/Venta:             %:"
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
            Left            =   4050
            TabIndex        =   62
            Top             =   810
            Width           =   2325
         End
         Begin VB.Label Label110 
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
            Left            =   6405
            TabIndex        =   61
            Top             =   690
            Width           =   1275
         End
         Begin VB.Label Label93 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2730
            TabIndex        =   60
            Top             =   690
            Width           =   1275
         End
         Begin VB.Label Label124 
            BackStyle       =   0  'Transparent
            Caption         =   "GG.Elaboración:               %:"
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
            Left            =   180
            TabIndex        =   59
            Top             =   1170
            Width           =   2535
         End
         Begin VB.Label Label100 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Costo Total Bruto:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3255
            TabIndex        =   58
            Top             =   1515
            Width           =   3060
         End
         Begin VB.Label Label111 
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
            Left            =   6405
            TabIndex        =   57
            Top             =   1440
            Width           =   1275
         End
      End
      Begin VB.Frame Frame26 
         BackColor       =   &H00C0E0EE&
         Caption         =   "Cobertura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
            Height          =   280
            Left            =   330
            TabIndex        =   88
            Top             =   640
            Width           =   540
         End
         Begin VB.Label Label113 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Pr.Lista:"
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
            Index           =   0
            Left            =   -1260
            TabIndex        =   55
            Top             =   420
            Width           =   2115
         End
         Begin VB.Label Label114 
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
            Height          =   280
            Left            =   945
            TabIndex        =   54
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label115 
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
            Height          =   280
            Left            =   945
            TabIndex        =   53
            Top             =   640
            Width           =   1275
         End
         Begin VB.Label Label113 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Pr.Neto:"
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
            Index           =   1
            Left            =   -1260
            TabIndex        =   52
            Top             =   1070
            Width           =   2115
         End
         Begin VB.Label Label116 
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
            Height          =   280
            Left            =   945
            TabIndex        =   51
            Top             =   1290
            Width           =   1275
         End
         Begin VB.Label Label113 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Margen:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   -1260
            TabIndex        =   50
            Top             =   1365
            Width           =   2115
         End
         Begin VB.Label Label119 
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
            Height          =   285
            Left            =   840
            TabIndex        =   49
            Top             =   2325
            Width           =   795
         End
         Begin VB.Label Label113 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "I.C.Pr:"
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
            Index           =   4
            Left            =   -420
            TabIndex        =   48
            Top             =   2400
            Width           =   1170
         End
         Begin VB.Line Line5 
            BorderColor     =   &H80000003&
            X1              =   0
            X2              =   2310
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Line Line6 
            BorderColor     =   &H80000005&
            X1              =   0
            X2              =   2310
            Y1              =   1695
            Y2              =   1695
         End
         Begin VB.Label Label64 
            BackStyle       =   0  'Transparent
            Caption         =   "D:"
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
            Left            =   105
            TabIndex        =   47
            Top             =   730
            Width           =   2010
         End
         Begin VB.Label Label112 
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
            Height          =   280
            Left            =   945
            TabIndex        =   46
            Top             =   965
            Width           =   1275
         End
         Begin VB.Label Label113 
            BackStyle       =   0  'Transparent
            Caption         =   "Mg.Bto(%):    Mg.Neto(%):   "
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   105
            TabIndex        =   45
            Top             =   1780
            Width           =   2100
         End
         Begin VB.Label Label117 
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
            Height          =   285
            Left            =   1420
            TabIndex        =   44
            Top             =   1980
            Width           =   795
         End
         Begin VB.Label Label118 
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
         BackColor       =   &H00C0E0EE&
         Caption         =   "Avance de Proceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
            BackColor       =   &H00C0E0EE&
            Caption         =   "Registros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
               BackStyle       =   0  'Transparent
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
               Height          =   255
               Index           =   0
               Left            =   630
               TabIndex        =   40
               Top             =   225
               Width           =   795
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Falta:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   -1620
               TabIndex        =   39
               Top             =   720
               Width           =   2175
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tot:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -1620
               TabIndex        =   38
               Top             =   255
               Width           =   2175
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Pro:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -1620
               TabIndex        =   37
               Top             =   480
               Width           =   2175
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Height          =   255
               Index           =   1
               Left            =   630
               TabIndex        =   36
               Top             =   455
               Width           =   795
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Height          =   255
               Index           =   2
               Left            =   630
               TabIndex        =   35
               Top             =   685
               Width           =   795
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0E0EE&
            Caption         =   "Tiempos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   1850
            TabIndex        =   27
            Top             =   1520
            Width           =   1660
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Trans:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -1515
               TabIndex        =   33
               Top             =   480
               Width           =   2175
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Estim:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -1515
               TabIndex        =   32
               Top             =   255
               Width           =   2175
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Falta:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -1515
               TabIndex        =   31
               Top             =   720
               Width           =   2175
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "00.00.00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   630
               TabIndex        =   30
               Top             =   225
               Width           =   975
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "00.00.00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   630
               TabIndex        =   29
               Top             =   455
               Width           =   975
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "00.00.00"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   630
               TabIndex        =   28
               Top             =   685
               Width           =   975
            End
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0EE&
            Caption         =   "Artículo:"
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
            Left            =   210
            TabIndex        =   26
            Top             =   465
            Width           =   1170
         End
         Begin VB.Label Label2 
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
            Height          =   540
            Left            =   160
            TabIndex        =   25
            Top             =   795
            Width           =   3320
         End
         Begin VB.Label Label1 
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
            Left            =   1470
            TabIndex        =   24
            Top             =   360
            Width           =   2010
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0EE&
         Caption         =   "Costo de Elaboración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
            Left            =   3780
            TabIndex        =   22
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label74 
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
            Left            =   3780
            TabIndex        =   21
            Top             =   690
            Width           =   1275
         End
         Begin VB.Label Label75 
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
            Left            =   2415
            TabIndex        =   20
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label76 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mano de Obra Directa:"
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
            Left            =   210
            TabIndex        =   19
            Top             =   420
            Width           =   2115
         End
         Begin VB.Label Label78 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Equipos de Producción:"
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
            Left            =   0
            TabIndex        =   18
            Top             =   795
            Width           =   2325
         End
         Begin VB.Label Label79 
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
            Left            =   2415
            TabIndex        =   17
            Top             =   690
            Width           =   1275
         End
         Begin VB.Label Label80 
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
            Left            =   5145
            TabIndex        =   16
            ToolTipText     =   "Doble-click Para Ver Análisis de Costos de Mano de Obra "
            Top             =   690
            Width           =   1275
         End
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00C0E0EE&
         Caption         =   "Costo de Suministros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Totales"
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
            Index           =   2
            Left            =   5145
            TabIndex        =   14
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Recibidos"
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
            Index           =   1
            Left            =   3780
            TabIndex        =   13
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "C.Propios"
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
            Index           =   3
            Left            =   2415
            TabIndex        =   12
            Top             =   210
            Width           =   1170
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000003&
            X1              =   0
            X2              =   6600
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Label Label81 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Trabajos de Terceros:"
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
            Left            =   210
            TabIndex        =   11
            Top             =   1470
            Width           =   2115
         End
         Begin VB.Label Label82 
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
            Left            =   3780
            TabIndex        =   10
            Top             =   630
            Width           =   1275
         End
         Begin VB.Label Label83 
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
            Left            =   3780
            TabIndex        =   9
            Top             =   1005
            Width           =   1275
         End
         Begin VB.Label Label84 
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
            Left            =   3780
            TabIndex        =   8
            Top             =   1365
            Width           =   1275
         End
         Begin VB.Label Label85 
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
            Left            =   5145
            TabIndex        =   7
            ToolTipText     =   "Doble-click para Ver Analisis de Costos de Materiales y Servicios"
            Top             =   1365
            Width           =   1275
         End
         Begin VB.Label Label86 
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
            Left            =   2415
            TabIndex        =   6
            Top             =   1365
            Width           =   1275
         End
         Begin VB.Label Label87 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Piezas Compradas:"
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
            Left            =   210
            TabIndex        =   5
            Top             =   1110
            Width           =   2115
         End
         Begin VB.Label Label88 
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
            Left            =   2415
            TabIndex        =   4
            Top             =   1005
            Width           =   1275
         End
         Begin VB.Label Label89 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Materias Primas:"
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
            Left            =   210
            TabIndex        =   3
            Top             =   735
            Width           =   2115
         End
         Begin VB.Label Label90 
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
            Left            =   2415
            TabIndex        =   2
            Top             =   630
            Width           =   1275
         End
         Begin VB.Line Line7 
            BorderColor     =   &H80000005&
            X1              =   0
            X2              =   6600
            Y1              =   435
            Y2              =   435
         End
      End
   End
End
Attribute VB_Name = "RECALCOS04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PROVINT%(32767)
Dim PAUSA%

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
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     CII% = I&
     TIMA% = TIMATE%(CII%) ' Asc(Mid$(XX$, 93, 1))
     UMED$ = UNIMED$(CII%) ' AJUSTI$(Mid$(XX$, 109, 2), 2)
     YY$ = String$(128, 0)
     Call REPLA(YY$, MKI$(CII%), 1, 2)
     Call REPLA(YY$, MKI$(TIMA%), 3, 2)
     Call REPLA(YY$, UMED$, 5, 2)
     Call REPLA(YY$, MKI$(FECO%), 7, 2)
     Call GRAREG("TRACALCO", YY$, I&)
     '\\\OT-06-0229-WAR-17/07/06///
     If HAYRECON% > 0 Then
        ZZ$ = String$(128, 0)
        Call REPLA(ZZ$, MKI$(CII%), 1, 2)
        Call REPLA(ZZ$, MKI$(TIMA%), 3, 2)
        Call REPLA(ZZ$, UMED$, 5, 2)
        Call REPLA(ZZ$, MKI$(FECO%), 7, 2)
        Call GRAREG("TRCOSCON", ZZ$, I&)
     End If
     '\\\OT-06-0229-WAR-FIN///
     '
     CODYY$ = CODEXT$(CII%)
     Call LEEEXPLO(CODYY$, 1)
     If CAIT% > 0 Then
       DP%(CII%) = 1
     End If
   Next I&
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
   TINI = Timer
   FINCALCOS% = 0
   '
   Screen.MousePointer = 1
   Set MASTERECO = Articulos.OpenRecordset("SELECT CODIGO,DESCRIPCION,CODINT FROM Master ORDER BY CODIGO ASC;", dbOpenSnapshot, dbReadOnly)
   On Error Resume Next
   MASTERECO.MoveLast
   If Err Then Exit Sub
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
                MONEI% = MONECOSTO%(CIK%(KKI%))
                COSUN = COSUNI(CIK%(KKI%)) * TICAMONE(MONEI%) * (1 + PORNACIO(CIK%(KKI%)) / 100)
                '
                UNIME$ = UCase$(UNIMED$(CIK%(KKI%)))
                If Left$(UNIME$, 1) = "U" Or Left$(UNIME$, 2) = "PZ" Then
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
              Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot)
              '
              On Error Resume Next
              With SecopTemp
                .MoveFirst
                If Err Then GoTo 25
                '
                Do Until .EOF = True
                  If IsNull(!AltOper) = True Or !AltOper < 1 Then       '
                     NOPY% = !NumeOper
                     DOPE$ = !DescOper
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
                     CATOP% = !CategOps
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
            Call REPLA(YY$, MKD$(XVALO(Label90) + XVALO(Label82)), 9, 8)
            Call REPLA(YY$, MKD$(XVALO(Label88) + XVALO(Label83)), 17, 8)
            Call REPLA(YY$, MKD$(XVALO(Label86) + XVALO(Label84)), 25, 8)
            Call REPLA(YY$, MKD$(XVALO(Label85)), 33, 8)
            Call REPLA(YY$, MKD$(XVALO(Label75) + XVALO(Label72)), 41, 8)
            Call REPLA(YY$, MKD$(XVALO(Label79) + XVALO(Label74)), 49, 8)
            Call REPLA(YY$, MKD$(0), 57, 8)
            Call REPLA(YY$, MKD$(XVALO(Label80)), 65, 8)
            Call REPLA(YY$, MKD$(XVALO(Label70)), 73, 8)
            Call REPLA(YY$, MKD$(XVALO(Label65)), 81, 8)
            Call REPLA(YY$, MKD$(XVALO(Label96)), 89, 8)
            Call REPLA(YY$, MKD$(XVALO(Label103)), 97, 8)
            Call REPLA(YY$, MKD$(XVALO(Label112)), 105, 8)
            Call REPLA(YY$, MKD$(XVALO(Label116)), 113, 8)
            Call REPLA(YY$, MKS$(XVALO(Label117)), 121, 4)
            Call REPLA(YY$, MKS$(XVALO(Label119)), 125, 4)
            Call GRAREG("TRACALCO", YY$, RETRA&)
            '\\\OT-06-0299-WAR-17/07/06///
            If HAYRECON% > 0 Then
               ZZ$ = REGLEIDO$("TRCOSCON", RETRA&)
               Call REPLA(ZZ$, MKD$(XVALO(Label70)), 9, 8)
               Call REPLA(ZZ$, MKD$(XVALO(Label104)), 17, 8)
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
     Label10(0) = TRIM$(HORMINSE$((Timer - TINI) * FIN& / IPIE&))
     Label10(1) = TRIM$(HORMINSE$(Timer - TINI))
     Label10(2) = TRIM$(HORMINSE$((Timer - TINI) * (FIN& / IPIE& - 1)))
     '
     DoEvents
     '
     While PAUSA% = 1
     DoEvents
     Wend
   Next IPIE&
   '
   Screen.MousePointer = 11
   FEX = HOY(HOS$)
   Call GRACONTROL("ANACOSTO", "FECALCOS", HOS$)
   Call GRACONTROL("ANACOSTO", "HOCALCOS", Left$(Time$, 5))
   '
   ANACOSTO04.Label34(4).Caption = HOS$
   ANACOSTO04.Label34(4).Refresh
   ANACOSTO04.Label95.Caption = Left$(Time$, 5)
   ANACOSTO04.Label95.Refresh
   '
   FIN& = ULTREG&("TRACALCO")
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     XX$ = REGLEIDO$("TRACALCO", I&)
     Call GRAREG("RECOSCAL", XX$, I&)
   Next I&
   Call SETULTREG("RECOSCAL", FIN&)
   Call CIERRARCH("RECOSCAL")
   '\\\OT-06-0299-WAR-17/07/06///
   If HAYRECON% > 0 Then
      FIN& = ULTREG&("TRCOSCON")
      For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        XX$ = REGLEIDO$("TRCOSCON", I&)
        Call GRAREG("RECOSCON", XX$, I&)
      Next I&
      Call SETULTREG("RECOSCON", FIN&)
      Call CIERRARCH("RECOSCON")
   End If
   '\\\OT-06-0299-WAR-FIN///
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   FINCALCOS% = 1
   Unload RECALCOS04
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
   SUMA1 = SUMA1 + XVALO(Label90.Caption)
   SUMA1 = SUMA1 + XVALO(Label88.Caption)
   SUMA1 = SUMA1 + XVALO(Label86.Caption)
   SUMA1 = SUMA1 + XVALO(Label82.Caption)
   SUMA1 = SUMA1 + XVALO(Label83.Caption)
   SUMA1 = SUMA1 + XVALO(Label84.Caption)
   Label85.Caption = TRIM$(CSTRING$(MKS$(SUMA1), 3, 12, DECICALCOS%, 2))
   Label85.Refresh
   '
   SUMA2 = 0
   SUMA2 = SUMA2 + XVALO(Label75.Caption)
   SUMA2 = SUMA2 + XVALO(Label72.Caption)
   SUMA2 = SUMA2 + XVALO(Label79.Caption)
   SUMA2 = SUMA2 + XVALO(Label74.Caption)
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
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
      Me.Frame1(4).BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame16.BackColor = &HFCD7B6

      Me.Frame17.BackColor = &HFCD7B6
      Me.Frame19.BackColor = &HFCD7B6
      Me.Frame26.BackColor = &HFCD7B6
      Me.Label3.BackColor = &HFCD7B6
  End If

    If RECALCOSTO% <> 1 Then Exit Sub
    '
    RECALCOSTO% = 0
    Top = ANACOSTO04.Top + ANACOSTO04.SSTab1.Top + 520
    Left = ANACOSTO04.Left + ANACOSTO04.SSTab1.Left
    Show
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    DoEvents
    '
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "GAGENMAT"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(1, KU&) = GGG: MODOCA%(1, KU&) = 1: Next
      Label241 = CSTRING$(MKS$(XVALO(DCC$)), 3, 5, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "GAGENFAB"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(2, KU&) = GGG: MODOCA%(2, KU&) = 1: Next
      Label236 = CSTRING$(MKS$(XVALO(DCC$)), 3, 5, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "GAGENADM"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(3, KU&) = GGG: MODOCA%(3, KU&) = 1: Next
      Label238 = CSTRING$(MKS$(XVALO(DCC$)), 3, 5, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "PORCOMIS"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(4, KU&) = GGG: MODOCA%(4, KU&) = 1: Next
      Label242 = CSTRING$(MKS$(XVALO(DCC$)), 3, 4, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "PORIMPUE"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(5, KU&) = GGG: MODOCA%(5, KU&) = 1: Next
      Label243 = CSTRING$(MKS$(XVALO(DCC$)), 3, 4, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "PORUTILI"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(6, KU&) = GGG: MODOCA%(6, KU&) = 1: Next
      Label239 = CSTRING$(MKS$(XVALO(DCC$)), 3, 4, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "PORDESCU"))
      GGG = XVALO(DCC$)
      For KU& = 0 To 255: PORGAGE(7, KU&) = GGG: MODOCA%(7, KU&) = 1: Next
      Label244 = CSTRING$(MKS$(XVALO(DCC$)), 3, 4, 1, 2)
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
   Call RECACOS
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Screen.MousePointer = 1
   Cancel = 1
   If FINCALCOS% <> 1 Then
     If COMALTER%("Realmente Desea Cancelar\el Proceso de Cálculo de Costos ?\\No, Continuar\Si, Cancelar") <> 2 Then
        Exit Sub
     End If
   End If
   Cancel = 0
   On Error Resume Next
   ANACOSTO04.Command55.Enabled = True
   ANACOSTO04.Text1.SetFocus
   On Error GoTo 0
   DoEvents
   Call FINAL("")
End Sub


Private Sub Label103_Change()
   If MODOCALCU%(4) = 1 Then Label110 = TRIM$(CSTRING$(MKD$(XVALO(Label103) * XVALO(Label242) / 100), 3, 12, DECICALCOS%, 2))
   If MODOCALCU%(5) = 1 Then Label104 = TRIM$(CSTRING$(MKD$(XVALO(Label103) * XVALO(Label243) / 100), 3, 12, DECICALCOS%, 2))
   If MODOCALCU%(6) = 1 Then Label107 = TRIM$(CSTRING$(MKD$(XVALO(Label103) * XVALO(Label239) / 100), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label104_Change()
   Label111 = TRIM$(CSTRING$(MKD$(XVALO(Label105) + XVALO(Label110) + XVALO(Label104)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label105_Change()
   PREVEN# = 0
   If (1 - (XVALO(Label239) + XVALO(Label242) + XVALO(Label243)) / 100) > 0.0001 Then
     PREVEN# = XVALO(Label105) / (1 - (XVALO(Label239) + XVALO(Label242) + XVALO(Label243)) / 100)
   End If
   Label103 = TRIM$(CSTRING$(MKD$(PREVEN#), 3, 12, DECICALCOS%, 2))
   Label111 = TRIM$(CSTRING$(MKD$(XVALO(Label105) + XVALO(Label110) + XVALO(Label104)), 3, 12, DECICALCOS%, 2))
   Label116 = ""
   If XVALO(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(XVALO(Label112) * (1 - XVALO(Label242) / 100 - XVALO(Label243) / 100) - XVALO(Label105)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label110_Change()
   Label111 = TRIM$(CSTRING$(MKD$(XVALO(Label105) + XVALO(Label110) + XVALO(Label104)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label111_Change()
   Label116 = ""
   If XVALO(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(XVALO(Label112) - XVALO(Label111)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label112_Change()
   Label116 = ""
   If XVALO(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(XVALO(Label112) * (1 - XVALO(Label242) / 100 - XVALO(Label243) / 100) - XVALO(Label105)), 3, 12, DECICALCOS%, 2))
   Label117 = ""
   If Abs(XVALO(Label112)) > 0.005 Then
     PORMA# = 100 * XVALO(Label116) / XVALO(Label112)
     If Abs(PORMA#) <= 99.9 Then
       Label117 = TRIM$(CSTRING$(MKD$(PORMA#), 3, 5, 1, 2))
     End If
   End If
   Label117.ForeColor = RGB(0, 0, 0)
   If XVALO(Label117) < 0 Then Label117.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Label114_Change()
   Label115 = TRIM$(CSTRING$(MKD$(XVALO(Label114) * XVALO(Label244) / 100), 3, 12, DECICALCOS%, 2))
   Label112 = TRIM$(CSTRING$(MKD$(XVALO(Label114) - XVALO(Label115)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label115_Change()
   Label116 = ""
   If XVALO(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(XVALO(Label114) - XVALO(Label115)), 3, 12, DECICALCOS%, 2))
   Label112 = TRIM$(CSTRING$(MKD$(XVALO(Label114) - XVALO(Label115)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label116_Change()
   Label117 = ""
   If Abs(XVALO(Label112)) > 0.005 Then
     PORMA# = 100 * XVALO(Label116) / XVALO(Label112)
     If Abs(PORMA#) <= 99.9 Then
       Label117 = TRIM$(CSTRING$(MKD$(PORMA#), 3, 5, 1, 2))
     End If
   End If
   Label117.ForeColor = RGB(0, 0, 0)
   If XVALO(Label117) < 0 Then Label117.ForeColor = RGB(255, 0, 0)
   '
End Sub

Private Sub Label117_Change()
   Label119 = ""
   If XVALO(Label239) > 0.005 Then
     PORMA# = 100 * XVALO(Label117) / XVALO(Label239)
     If Abs(PORMA#) <= 999.9 Then
       Label119 = TRIM$(CSTRING$(MKD$(PORMA#), 3, 5, 1, 2))
     End If
   End If
   Label119.ForeColor = RGB(0, 0, 0)
   If XVALO(Label119) < 100 Then Label119.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Label65_Change()
   Label77 = TRIM$(CSTRING$(MKD$(XVALO(Label65) * XVALO(Label238) / 100), 3, 12, DECICALCOS%, 2))
   Label96 = TRIM$(CSTRING$(MKD$(XVALO(Label65) + XVALO(Label77)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label68_Change()
   Label65 = TRIM$(CSTRING$(MKD$(XVALO(Label70) + XVALO(Label68) + XVALO(Label93)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label70_Change()
   Label65 = TRIM$(CSTRING$(MKD$(XVALO(Label70) + XVALO(Label68) + XVALO(Label93)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label77_Change()
   Label96 = TRIM$(CSTRING$(MKD$(XVALO(Label65) + XVALO(Label77)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label80_Change()
   Label68 = TRIM$(CSTRING$(MKD$(XVALO(Label80) * XVALO(Label236) / 100), 3, 12, DECICALCOS%, 2))
   Label68.Refresh
End Sub

Private Sub Label80_DBLClick()
   Label80.Enabled = False
   If ANACOSMOD.List1.ListCount > 0 Then
     ANACOSMOD.Show 1
   End If
   Label80.Enabled = True
End Sub

Private Sub Label85_Change()
   Label93 = TRIM$(CSTRING$(MKD$(XVALO(Label85) * XVALO(Label241) / 100), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label85_dblClick()
   Label85.Enabled = False
   If ANACOSMAT.List1.ListCount > 0 Then
     ANACOSMAT.Show 1
   End If
   Label85.Enabled = True
End Sub

Private Sub Label93_Change()
   Label65 = TRIM$(CSTRING$(MKD$(XVALO(Label70) + XVALO(Label68) + XVALO(Label93)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label96_Change()
   Label105 = Label96
End Sub


Private Sub LABEL236_Change()
   'Label68 = TRIM$(CSTRING$(MKD$(xvalo(Label80) * xvalo(Label236) / 100), 3, 12, DECICALCOS%, 2))
   If MODOCALCU%(2) = 1 Then
        Label68 = TRIM$(CSTRING$(MKD$(XVALO(Label80) * XVALO(Label236) / 100), 3, 12, DECICALCOS%, 2))
      Else
        Label68 = TRIM$(CSTRING$(MKD$(PREVENETO# * XVALO(Label236) / 100), 3, 12, DECICALCOS%, 2))
   End If
End Sub

Private Sub LABEL238_Change()
   'Label77 = TRIM$(CSTRING$(MKD$(xvalo(Label65) * xvalo(Label238) / 100), 3, 12, DECICALCOS%, 2))
   If MODOCALCU%(3) = 1 Then
        Label77 = TRIM$(CSTRING$(MKD$(XVALO(Label65) * XVALO(Label238) / 100), 3, 12, DECICALCOS%, 2))
      Else
        Label77 = TRIM$(CSTRING$(MKD$(PREVENETO# * XVALO(Label238) / 100), 3, 12, DECICALCOS%, 2))
   End If
End Sub

Private Sub LABEL239_Change()
   PREVEN# = 0
   If (1 - (XVALO(Label239) + XVALO(Label242) + XVALO(Label243)) / 100) > 0.0001 Then
     PREVEN# = XVALO(Label105) / (1 - (XVALO(Label239) + XVALO(Label242) + XVALO(Label243)) / 100)
   End If
   Label103 = TRIM$(CSTRING$(MKD$(PREVEN#), 3, 12, DECICALCOS%, 2))
   'Label107 = TRIM$(CSTRING$(MKD$(PREVEN# * xvalo(Label239) / 100), 3, 12, DECICALCOS%, 2))
   '
   If MODOCALCU%(6) = 1 Then
        Label107 = TRIM$(CSTRING$(MKD$(PREVEN# * XVALO(Label239) / 100), 3, 12, DECICALCOS%, 2))
      Else
        Label107 = TRIM$(CSTRING$(MKD$(PREVENETO# * XVALO(Label239) / 100), 3, 12, DECICALCOS%, 2))
   End If
   '
   Label119 = ""
   If XVALO(Label239) > 0.005 Then
     PORMA# = 100 * XVALO(Label117) / XVALO(Label239)
     If Abs(PORMA#) <= 999.9 Then
       Label119 = TRIM$(CSTRING$(MKD$(PORMA#), 3, 5, 1, 2))
     End If
   End If
   Label119.ForeColor = RGB(0, 0, 0)
   If XVALO(Label119) < 100 Then Label119.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub LABEL241_Change()
   'Label93 = TRIM$(CSTRING$(MKD$(xvalo(Label85) * xvalo(Label241) / 100), 3, 12, DECICALCOS%, 2))
   If MODOCALCU%(1) = 1 Then
        Label93 = TRIM$(CSTRING$(MKD$(XVALO(Label85) * XVALO(Label241) / 100), 3, 12, DECICALCOS%, 2))
     Else
        Label93 = TRIM$(CSTRING$(MKD$(PREVENETO# * XVALO(Label241) / 100), 3, 12, DECICALCOS%, 2))
   End If
End Sub

Private Sub LABEL242_Change()
   PREVEN# = 0
   If (1 - (XVALO(Label239) + XVALO(Label242) + XVALO(Label243)) / 100) > 0.0001 Then
     PREVEN# = XVALO(Label105) / (1 - (XVALO(Label239) + XVALO(Label242) + XVALO(Label243)) / 100)
   End If
   Label103 = TRIM$(CSTRING$(MKD$(PREVEN#), 3, 12, DECICALCOS%, 2))
   'Label110 = TRIM$(CSTRING$(MKD$(PREVEN# * xvalo(Label242) / 100), 3, 12, DECICALCOS%, 2))
   '
   If MODOCALCU%(4) = 1 Then
        Label110 = TRIM$(CSTRING$(MKD$(PREVEN# * XVALO(Label242) / 100), 3, 12, DECICALCOS%, 2))
      Else
        Label110 = TRIM$(CSTRING$(MKD$(PREVENETO# * XVALO(Label242) / 100), 3, 12, DECICALCOS%, 2))
   End If
   '
   Label116 = ""
   If XVALO(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(XVALO(Label112) * (1 - XVALO(Label242) / 100 - XVALO(Label243) / 100) - XVALO(Label105)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub LABEL243_Change()
   PREVEN# = 0
   If (1 - (XVALO(Label239) + XVALO(Label242) + XVALO(Label243)) / 100) > 0.0001 Then
     PREVEN# = XVALO(Label105) / (1 - (XVALO(Label239) + XVALO(Label242) + XVALO(Label243)) / 100)
   End If
   Label103 = TRIM$(CSTRING$(MKD$(PREVEN#), 3, 12, DECICALCOS%, 2))
   'Label104 = TRIM$(CSTRING$(MKD$(PREVEN# * xvalo(Label243) / 100), 3, 12, DECICALCOS%, 2))
   '
   If MODOCALCU%(5) = 1 Then
        Label104 = TRIM$(CSTRING$(MKD$(PREVEN# * XVALO(Label243) / 100), 3, 12, DECICALCOS%, 2))
      Else
        Label104 = TRIM$(CSTRING$(MKD$(PREVENETO# * XVALO(Label243) / 100), 3, 12, DECICALCOS%, 2))
   End If
   '
   If XVALO(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(XVALO(Label112) * (1 - XVALO(Label242) / 100 - XVALO(Label243) / 100) - XVALO(Label105)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub LABEL244_Change()
   Label115 = TRIM$(CSTRING$(MKD$(XVALO(Label114) * XVALO(Label244) / 100), 3, 12, DECICALCOS%, 2))
End Sub

