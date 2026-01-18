VERSION 5.00
Begin VB.Form RECALCOS 
   BackColor       =   &H00C0E0EE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Re-Cálculo Integral de Costos"
   ClientHeight    =   6150
   ClientLeft      =   165
   ClientTop       =   2205
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   10410
   Begin VB.Frame Frame16 
      BackColor       =   &H00C0E0EE&
      BorderStyle     =   0  'None
      Caption         =   "Frame16"
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10200
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
         Left            =   2415
         TabIndex        =   56
         Top             =   3150
         Width           =   7785
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
         Width           =   2325
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
         Caption         =   "Detener"
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
         Left            =   9030
         TabIndex        =   41
         Top             =   105
         Width           =   1170
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
         Width           =   3480
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
            Left            =   105
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
            Left            =   1725
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
            Left            =   105
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
            Left            =   105
            TabIndex        =   25
            Top             =   795
            Width           =   3270
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
            Width           =   1905
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
Attribute VB_Name = "RECALCOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PROVINT%(32767)

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
   '
   Call COPYSTRU("RECOSCAL", "TRACALCO")
   FIN& = NUMAS%
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     XX$ = REGLEIDO$("MASTER", I& + 1)
     CII% = I&
     TIMA% = Asc(Mid$(XX$, 93, 1))
     UMED$ = AJUSTI$(Mid$(XX$, 109, 2), 2)
     YY$ = String$(128, 0)
     Call REPLA(YY$, MKI$(CII%), 1, 2)
     Call REPLA(YY$, MKI$(TIMA%), 3, 2)
     Call REPLA(YY$, UMED$, 5, 2)
     Call REPLA(YY$, MKI$(FECO%), 7, 2)
     Call GRAREG("TRACALCO", YY$, I&)
   Next I&
   '
   Call SETULTREG("TRACALCO", FIN&)
   Call CIERRARCH("TRACALCO")
   '
   ' IDENTIFICANDO LOS ITEMS SIN EXPLOSION
   '
   FIN& = ULTREG&("ESTRUNUE")
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     R0$ = REGLEIDO$("ESTRUNUE", I&)
     CI1% = CVI(Left$(R0$, 2))
     If CI1% > 0 Then
       If CI1% <= NMM% Then
         If CVS(Mid$(R0$, 5, 4)) > 0.00005 Then
           DP%(CI1%) = 1
         End If
       End If
     End If
   Next I&
   '
   '
   TINI = Timer
   FINCALCOS% = 0
   '
   Screen.MousePointer = 1
   FIN& = NUINV%
   
   For IPIE& = 1 To FIN&
     Call TRACE(20, IPIE&, FIN&)
     I0$ = REGLEIDO$("INVERT", IPIE& + 1)
     CODD$ = TRIM$(Left$(I0$, 16))
     If Mid$(CODD$, 16, 1) <> Chr$(96) Then
       CI% = CVI(Mid$(I0$, 17, 2))
       If CI% > 0 Then
         If CI% <= NMM% Then
            CICON% = CI%
            '
            Label1 = CODD$
            Label2 = DESCRIT$(CI%)
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
              RFF$ = RECFILT$("SECOPER", 1, MKI$(CIXI%))
              RCT$ = RECFILT$("COSOPTER", 1, MKI$(CIXI%))
              '
              For UJ& = 1 To Len(RFF$) Step 4
                I& = CVS(Mid$(RFF$, UJ&, 4))
                If I& > 0 Then
                  If I& <= URE& Then
                    SCP$ = REGLEIDO$("SECOPER", I&)
                    NOPY% = CVI(Mid$(SCP$, 3, 2))
                    LOTESTA = LOTESTAN(CIXI%): If LOTESTA < 1 Then LOTESTA = 1
                    CICLOHE = CVS(Mid$(SCP$, 99, 4)): If CICLOHE < 1 Then CICLOHE = 1
                    MINPREP = 60 * CVS(Mid$(SCP$, 59, 4))
                    MINFIJO = 60 * CVS(Mid$(SCP$, 63, 4))
                    MINVARI = LOTESTA * CVS(Mid$(SCP$, 67, 4))
                    MINHERR = LOTESTA * 60 * CVS(Mid$(SCP$, 95, 4)) / CICLOHE
                    MINDEMO = (MINFIJO + MINVARI + MINHERR) * CVS(Mid$(SCP$, 75, 4)) / 100
                    MINLOTE = MINPREP + MINFIJO + MINVARI + MINHERR + MINDEMO
                    HORAPROD = CACOM * MINLOTE / 60 / LOTESTA
                    '
                    CAOPS = CVS(Mid$(SCP$, 83, 4))
                    CATOP% = Asc(Mid$(SCP$, 87, 1))
                    EQUIPO$ = TRIM$(Mid$(SCP$, 53, 6))
                    COSMOD = CAOPS * HORAPROD * VALHORA(CATOP%)
                    COSMAQ = HORAPROD * CUHORMAQ(EQUIPO$)
                    '
                    TERCOP% = 0
                    For UT& = 1 To Len(RCT$) Step 4
                      IT& = CVS(Mid$(RCT$, UT&, 4))
                      If IT& > 0 Then
                        If IT& <= UCT& Then
                          SCT$ = REGLEIDO$("COSOPTER", IT&)
                          If CVI(Mid$(SCT$, 3, 2)) = NOPY% Then
                            MONEI% = CVI(Mid$(SCT$, 15, 2))
                            COSTER = CACOM * CVD(Mid$(SCT$, 7, 8)) * TICAMONE(MONEI%)
                            If COSTER >= 0.00005 Then
                              If (COSTER < COSMOD + COSMAQ) Or ((COSMOD + COSMAQ) < 0.00005) Then
                                COSMOD = COSTER
                                COSMAQ = 0
                                HORAPROD = 0
                                TERCOP% = 1
                              End If
                            End If
                          End If
                        End If
                      End If
                    Next UT&
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
                    '
                  End If
                End If
              Next UJ&
              '
            Next UI&
            '
            '
            GoSub 90
            '
            PRELISX# = PRELISTA#("LIPRE001", CICON%) * TICAMONE(MONELISTA%(1))
            Label114.Caption = TRIM$(CSTRING$(MKD$(PRELISX#), 3, 12, DECICALCOS%, 2))
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
69 Next IPIE&
   '
   Screen.MousePointer = 11
   FEX = HOY(HOS$)
   Call GRACONTROL("ANACOSTO", "FECALCOS", HOS$)
   Call GRACONTROL("ANACOSTO", "HOCALCOS", Left$(Time$, 5))
   '
   ANACOSTO.Label34(4).Caption = HOS$
   ANACOSTO.Label34(4).Refresh
   ANACOSTO.Label95.Caption = Left$(Time$, 5)
   ANACOSTO.Label95.Refresh
   '
   FIN& = ULTREG&("TRACALCO")
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     XX$ = REGLEIDO$("TRACALCO", I&)
     Call GRAREG("RECOSCAL", XX$, I&)
   Next I&
   Call SETULTREG("RECOSCAL", FIN&)
   Call CIERRARCH("RECOSCAL")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   FINCALCOS% = 1
   Unload RECALCOS
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

Private Sub Command1_Click()
   Unload RECALCOS
End Sub

Private Sub Form_Load()
    '
    If RECALCOSTO% <> 1 Then Exit Sub
    '
    RECALCOSTO% = 0
    Show
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    DoEvents
    '
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "GAGENMAT"))
      Label241 = CSTRING$(MKS$(Val(DCC$)), 3, 5, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "GAGENFAB"))
      Label236 = CSTRING$(MKS$(Val(DCC$)), 3, 5, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "GAGENADM"))
      Label238 = CSTRING$(MKS$(Val(DCC$)), 3, 5, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "PORUTILI"))
      Label239 = CSTRING$(MKS$(Val(DCC$)), 3, 4, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "PORCOMIS"))
      Label242 = CSTRING$(MKS$(Val(DCC$)), 3, 4, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "PORIMPUE"))
      Label243 = CSTRING$(MKS$(Val(DCC$)), 3, 4, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "PORDESCU"))
      Label244 = CSTRING$(MKS$(Val(DCC$)), 3, 4, 1, 2)
      '
   For UI& = 0 To 32767
     PROVINT%(UI&) = 0
   Next UI&
   '
   For U& = 1 To ULTREG&("TAPROINT")
      X$ = REGLEIDO$("TAPROINT", U&)
      NP% = CVI(Left$(X$, 2))
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
   ANACOSTO.Command55.Enabled = True
   ANACOSTO.Text1.SetFocus
   On Error GoTo 0
   DoEvents
End Sub


Private Sub Label103_Change()
   Label107 = TRIM$(CSTRING$(MKD$(Val(Label103) * Val(Label239) / 100), 3, 12, DECICALCOS%, 2))
   Label110 = TRIM$(CSTRING$(MKD$(Val(Label103) * Val(Label242) / 100), 3, 12, DECICALCOS%, 2))
   Label104 = TRIM$(CSTRING$(MKD$(Val(Label103) * Val(Label243) / 100), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label104_Change()
   Label111 = TRIM$(CSTRING$(MKD$(Val(Label105) + Val(Label110) + Val(Label104)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label105_Change()
   PREVEN# = Val(Label105) / (1 - (Val(Label239) + Val(Label242) + Val(Label243)) / 100)
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
   If ANACOSMOD.List1.ListCount > 0 Then
     ANACOSMOD.Show 1
   End If
   Label80.Enabled = True
End Sub

Private Sub Label85_Change()
   Label93 = TRIM$(CSTRING$(MKD$(Val(Label85) * Val(Label241) / 100), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label85_dblClick()
   Label85.Enabled = False
   If ANACOSMAT.List1.ListCount > 0 Then
     ANACOSMAT.Show 1
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
   Label68 = TRIM$(CSTRING$(MKD$(Val(Label80) * Val(Label236) / 100), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub LABEL238_Change()
   Label77 = TRIM$(CSTRING$(MKD$(Val(Label65) * Val(Label238) / 100), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub LABEL239_Change()
   PREVEN# = Val(Label105) / (1 - (Val(Label239) + Val(Label242) + Val(Label243)) / 100)
   Label103 = TRIM$(CSTRING$(MKD$(PREVEN#), 3, 12, DECICALCOS%, 2))
   Label107 = TRIM$(CSTRING$(MKD$(PREVEN# * Val(Label239) / 100), 3, 12, DECICALCOS%, 2))
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
   Label93 = TRIM$(CSTRING$(MKD$(Val(Label85) * Val(Label241) / 100), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub LABEL242_Change()
   PREVEN# = Val(Label105) / (1 - (Val(Label239) + Val(Label242) + Val(Label243)) / 100)
   Label103 = TRIM$(CSTRING$(MKD$(PREVEN#), 3, 12, DECICALCOS%, 2))
   Label110 = TRIM$(CSTRING$(MKD$(PREVEN# * Val(Label242) / 100), 3, 12, DECICALCOS%, 2))
   Label116 = ""
   If Val(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(Val(Label112) * (1 - Val(Label242) / 100 - Val(Label243) / 100) - Val(Label105)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub LABEL243_Change()
   PREVEN# = Val(Label105) / (1 - (Val(Label239) + Val(Label242) + Val(Label243)) / 100)
   Label103 = TRIM$(CSTRING$(MKD$(PREVEN#), 3, 12, DECICALCOS%, 2))
   Label104 = TRIM$(CSTRING$(MKD$(PREVEN# * Val(Label243) / 100), 3, 12, DECICALCOS%, 2))
   If Val(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(Val(Label112) * (1 - Val(Label242) / 100 - Val(Label243) / 100) - Val(Label105)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub LABEL244_Change()
   Label115 = TRIM$(CSTRING$(MKD$(Val(Label114) * Val(Label244) / 100), 3, 12, DECICALCOS%, 2))
End Sub

