VERSION 5.00
Begin VB.Form REMECOMP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FDD0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Gestión de Compras - Resúmenes Mensuales"
   ClientHeight    =   8355
   ClientLeft      =   435
   ClientTop       =   300
   ClientWidth     =   11115
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   10080
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   107
      Top             =   7320
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   108
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FDD0&
      Caption         =   "Período Mensual Activo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8100
      Left            =   210
      TabIndex        =   11
      Top             =   105
      Width           =   9675
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FDD0&
         Caption         =   "Percepciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3510
         Left            =   5460
         TabIndex        =   87
         Top             =   4410
         Width           =   3900
         Begin VB.Label TOTAGEN 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   2415
            TabIndex        =   100
            Top             =   3045
            Width           =   1275
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total General:"
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
            Index           =   14
            Left            =   -315
            TabIndex        =   99
            Top             =   3150
            Width           =   2640
         End
         Begin VB.Label SUTOPER 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   2415
            TabIndex        =   98
            Top             =   2415
            Width           =   1275
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Subtotal Percepciones:"
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
            Index           =   11
            Left            =   -315
            TabIndex        =   97
            Top             =   2520
            Width           =   2640
         End
         Begin VB.Label PERIBRU 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   2415
            TabIndex        =   96
            Top             =   1305
            Width           =   1275
         End
         Begin VB.Label PERGAN 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   2415
            TabIndex        =   95
            Top             =   630
            Width           =   1275
         End
         Begin VB.Label PERIVA 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   2415
            TabIndex        =   94
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Percepciones Ingr.Brutos:"
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
            Index           =   9
            Left            =   -315
            TabIndex        =   93
            Top             =   1410
            Width           =   2640
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Percepción Ganancias:"
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
            Index           =   8
            Left            =   315
            TabIndex        =   89
            Top             =   735
            Width           =   2010
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Percepcion I.V.A:"
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
            Index           =   7
            Left            =   105
            TabIndex        =   88
            Top             =   420
            Width           =   2220
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FDD0&
         Caption         =   "Posicion Mensual I.V.A."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3510
         Left            =   315
         TabIndex        =   55
         Top             =   4410
         Width           =   4950
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0FDD0&
            BorderStyle     =   0  'None
            Caption         =   "Frame6"
            Height          =   2745
            Left            =   105
            TabIndex        =   57
            Top             =   630
            Width           =   4740
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Index           =   8
               Left            =   0
               TabIndex        =   106
               Top             =   2055
               Width           =   855
            End
            Begin VB.Label IMPOIVA 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   8
               Left            =   2310
               TabIndex        =   104
               Top             =   1995
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.Label IMPOIVA 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   7
               Left            =   2310
               TabIndex        =   84
               Top             =   1710
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.Label IMPOIVA 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   6
               Left            =   2310
               TabIndex        =   78
               Top             =   1425
               Width           =   1275
            End
            Begin VB.Label IMPOGRA 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   8
               Left            =   945
               TabIndex        =   105
               Top             =   1995
               Width           =   1275
            End
            Begin VB.Label TAREAL 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   8
               Left            =   3675
               TabIndex        =   103
               Top             =   1995
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Index           =   7
               Left            =   5
               TabIndex        =   92
               Top             =   1770
               Width           =   855
            End
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Index           =   6
               Left            =   5
               TabIndex        =   91
               Top             =   1485
               Width           =   855
            End
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Index           =   0
               Left            =   5
               TabIndex        =   90
               Top             =   2475
               Width           =   855
            End
            Begin VB.Label IMPOIVA 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   5
               Left            =   2310
               TabIndex        =   86
               Top             =   1140
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.Label TAREAL 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   7
               Left            =   3675
               TabIndex        =   85
               Top             =   1710
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.Label IMPOGRA 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   7
               Left            =   945
               TabIndex        =   83
               Top             =   1710
               Width           =   1275
            End
            Begin VB.Label TAREAL 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   0
               Left            =   3675
               TabIndex        =   82
               Top             =   2415
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.Label IMPOIVA 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   0
               Left            =   2310
               TabIndex        =   81
               Top             =   2415
               Width           =   1275
            End
            Begin VB.Label IMPOGRA 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   0
               Left            =   945
               TabIndex        =   80
               Top             =   2415
               Width           =   1275
            End
            Begin VB.Label TAREAL 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   6
               Left            =   3675
               TabIndex        =   79
               Top             =   1425
               Width           =   960
            End
            Begin VB.Label IMPOGRA 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   6
               Left            =   945
               TabIndex        =   77
               Top             =   1425
               Width           =   1275
            End
            Begin VB.Label TAREAL 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   5
               Left            =   3675
               TabIndex        =   76
               Top             =   1140
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Index           =   5
               Left            =   5
               TabIndex        =   75
               Top             =   1200
               Width           =   850
            End
            Begin VB.Label IMPOGRA 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   5
               Left            =   945
               TabIndex        =   74
               Top             =   1140
               Width           =   1275
            End
            Begin VB.Label TAREAL 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   4
               Left            =   3675
               TabIndex        =   73
               Top             =   855
               Width           =   960
            End
            Begin VB.Label IMPOIVA 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   4
               Left            =   2310
               TabIndex        =   72
               Top             =   855
               Width           =   1275
            End
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Index           =   4
               Left            =   5
               TabIndex        =   71
               Top             =   915
               Width           =   850
            End
            Begin VB.Label IMPOGRA 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   4
               Left            =   945
               TabIndex        =   70
               Top             =   855
               Width           =   1275
            End
            Begin VB.Label TAREAL 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   3
               Left            =   3675
               TabIndex        =   69
               Top             =   570
               Width           =   960
            End
            Begin VB.Label IMPOIVA 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   3
               Left            =   2310
               TabIndex        =   68
               Top             =   570
               Width           =   1275
            End
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Index           =   3
               Left            =   5
               TabIndex        =   67
               Top             =   630
               Width           =   850
            End
            Begin VB.Label IMPOGRA 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   3
               Left            =   945
               TabIndex        =   66
               Top             =   570
               Width           =   1275
            End
            Begin VB.Label TAREAL 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   2
               Left            =   3675
               TabIndex        =   65
               Top             =   285
               Width           =   960
            End
            Begin VB.Label IMPOIVA 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   2
               Left            =   2310
               TabIndex        =   64
               Top             =   285
               Width           =   1275
            End
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Index           =   2
               Left            =   5
               TabIndex        =   63
               Top             =   345
               Width           =   850
            End
            Begin VB.Label IMPOGRA 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   2
               Left            =   945
               TabIndex        =   62
               Top             =   285
               Width           =   1275
            End
            Begin VB.Label TAREAL 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   1
               Left            =   3675
               TabIndex        =   61
               Top             =   0
               Width           =   960
            End
            Begin VB.Label IMPOIVA 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   1
               Left            =   2310
               TabIndex        =   60
               Top             =   0
               Width           =   1275
            End
            Begin VB.Label TASIVA 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Index           =   1
               Left            =   5
               TabIndex        =   59
               Top             =   60
               Width           =   850
            End
            Begin VB.Label IMPOGRA 
               Alignment       =   1  'Right Justify
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
               Height          =   300
               Index           =   1
               Left            =   945
               TabIndex        =   58
               Top             =   0
               Width           =   1275
            End
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "T.Nomin.  Neto Gravado    Importe I.V.A.      T.Real"
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
            TabIndex        =   56
            Top             =   380
            Width           =   6315
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0FDD0&
         Caption         =   "Avance de Proceso de Cálculo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2985
         Left            =   5145
         TabIndex        =   39
         Top             =   1365
         Visible         =   0   'False
         Width           =   4215
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FDD0&
            Caption         =   "Resumenes Mensuales Compras y Pagos"
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
            Index           =   6
            Left            =   210
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   2130
            Width           =   3795
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FDD0&
            Caption         =   "Saldos Mensuales por Cuenta"
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
            Index           =   7
            Left            =   210
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   2435
            Width           =   3795
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FDD0&
            Caption         =   "Asiento Resumen Mensual de Compras"
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
            Index           =   5
            Left            =   210
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   1845
            Width           =   3795
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FDD0&
            Caption         =   "Generación del Subdiario de  Compras"
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
            Left            =   210
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1560
            Width           =   3795
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FDD0&
            Caption         =   "Cálculo Posición Fiscal I.V.A.- Compras"
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
            Left            =   210
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   1275
            Width           =   3795
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FDD0&
            Caption         =   "Generación del Resumen Mensual "
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
            Left            =   210
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   990
            Width           =   3795
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FDD0&
            Caption         =   "Clasificación de Cuentas por Tipo"
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
            Left            =   210
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   705
            Width           =   3795
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FDD0&
            Caption         =   "Validación de Cuentas Contables"
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
            Left            =   210
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   420
            Width           =   3795
         End
      End
      Begin VB.TextBox MESAN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1575
         TabIndex        =   0
         Text            =   " "
         Top             =   735
         Width           =   2220
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FDD0&
         Caption         =   "Resumen Mensual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2985
         Left            =   315
         TabIndex        =   23
         Top             =   1365
         Width           =   4530
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   2835
            TabIndex        =   37
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2100
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   2835
            TabIndex        =   35
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1470
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   2835
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   420
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1155
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   840
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   1155
            TabIndex        =   27
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1155
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   1155
            TabIndex        =   26
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1470
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   2835
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   " "
            Top             =   1785
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   2835
            TabIndex        =   24
            TabStop         =   0   'False
            Text            =   " "
            Top             =   2520
            Width           =   1485
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Final del Período"
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
            Index           =   6
            Left            =   -105
            TabIndex        =   38
            Top             =   2625
            Width           =   2745
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ajustes ........................."
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
            Left            =   -525
            TabIndex        =   36
            Top             =   2205
            Width           =   3165
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo Inicial del Período"
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
            Left            =   -105
            TabIndex        =   34
            Top             =   525
            Width           =   2745
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Facturas"
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
            Left            =   -525
            TabIndex        =   33
            Top             =   840
            Width           =   1590
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Débitos"
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
            Left            =   -525
            TabIndex        =   32
            Top             =   1155
            Width           =   1590
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Créditos"
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
            Left            =   -525
            TabIndex        =   31
            Top             =   1470
            Width           =   1590
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Pagos ........................."
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
            Index           =   5
            Left            =   -525
            TabIndex        =   30
            Top             =   1890
            Width           =   3165
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FDD0&
         Caption         =   "Cuentas Contables"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2985
         Left            =   5145
         TabIndex        =   16
         Top             =   1365
         Width           =   4215
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   315
            TabIndex        =   21
            Text            =   " "
            Top             =   2520
            Width           =   3615
         End
         Begin VB.CommandButton Command4 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   3780
            TabIndex        =   7
            Top             =   2100
            Width           =   150
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   2625
            TabIndex        =   6
            Text            =   " "
            Top             =   2100
            Width           =   1170
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   315
            TabIndex        =   19
            Text            =   " "
            Top             =   1680
            Width           =   3615
         End
         Begin VB.CommandButton Command4 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   3780
            TabIndex        =   5
            Top             =   1260
            Width           =   150
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   2625
            TabIndex        =   4
            Text            =   " "
            Top             =   1260
            Width           =   1170
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   315
            TabIndex        =   17
            Text            =   " "
            Top             =   840
            Width           =   3615
         End
         Begin VB.CommandButton Command4 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   3780
            TabIndex        =   3
            Top             =   420
            Width           =   150
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   2625
            TabIndex        =   2
            Text            =   " "
            Top             =   420
            Width           =   1170
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Acreed.Comerc (Omis):"
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
            Left            =   210
            TabIndex        =   22
            Top             =   2130
            Width           =   2325
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Aduana:"
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
            Left            =   630
            TabIndex        =   20
            Top             =   1290
            Width           =   1905
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Compras:"
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
            Left            =   105
            TabIndex        =   18
            Top             =   450
            Width           =   2430
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FDD0&
         Caption         =   "Generación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   5145
         TabIndex        =   13
         Top             =   315
         Width           =   4215
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   3150
            TabIndex        =   49
            TabStop         =   0   'False
            Text            =   " "
            Top             =   420
            Width           =   750
         End
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   1050
            TabIndex        =   48
            TabStop         =   0   'False
            Text            =   " "
            Top             =   420
            Width           =   1065
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha:"
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
            Index           =   13
            Left            =   -1260
            TabIndex        =   15
            Top             =   525
            Width           =   2220
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hora:"
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
            Index           =   12
            Left            =   1470
            TabIndex        =   14
            Top             =   525
            Width           =   1590
         End
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
         Height          =   360
         Left            =   3780
         TabIndex        =   1
         Top             =   735
         Width           =   175
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   4305
         Picture         =   "REMECOMP.frx":0000
         Top             =   700
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4305
         Picture         =   "REMECOMP.frx":030A
         Top             =   630
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mes y año:"
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
         Index           =   10
         Left            =   -735
         TabIndex        =   12
         Top             =   840
         Width           =   2220
      End
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   102
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "I.V.A"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   10080
      Picture         =   "REMECOMP.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Emite Subdiario de I.V.A.- Compras por Impresora"
      Top             =   2950
      Width           =   855
   End
   Begin VB.ListBox LISORT 
      Height          =   840
      Left            =   9975
      Sorted          =   -1  'True
      TabIndex        =   101
      Top             =   2940
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10080
      Picture         =   "REMECOMP.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Emite Reporte Resumen por Impresora"
      Top             =   6615
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Informes"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   10080
      Picture         =   "REMECOMP.frx":0F88
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Resúmenes Estadísticos de Compras"
      Top             =   5650
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Asiento"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   10080
      Picture         =   "REMECOMP.frx":1292
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Presenta Asiento Contable de Compras"
      Top             =   4750
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Subdiario"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   10080
      Picture         =   "REMECOMP.frx":159C
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Emite Subdiario de Compras por Impresora"
      Top             =   3850
      Width           =   855
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
      Height          =   825
      Left            =   10080
      Picture         =   "REMECOMP.frx":18A6
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1150
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cálculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   10080
      Picture         =   "REMECOMP.frx":1BB0
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Re-Calcula Resúmenes Mensuales"
      Top             =   2050
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Height          =   825
      Left            =   10080
      Picture         =   "REMECOMP.frx":1EBA
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Finaliza la Aplicación"
      Top             =   250
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   510
      Left            =   9520
      Picture         =   "REMECOMP.frx":2004
      Top             =   7680
      Width           =   2010
   End
End
Attribute VB_Name = "REMECOMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub
'

Private Sub Command2_Click()
   Command2.Enabled = False
   On Error Resume Next
   Command1.SetFocus
   On Error GoTo 0
   Call CALRECOM
   Command2.Enabled = True
End Sub
'

Private Sub Command29_Click()
   Call HELPONLINE("COGECOM")
End Sub

Private Sub Command3_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
        MESAN.TEXT = MESASEL$
        Call CARGARESCOM
        Call FRESHALL
        MESAN.SetFocus
    End If
    Call FRESHALL
End Sub
'

Private Sub Command4_Click(Index As Integer)
    Call SELECHO("ECUECON")
    If VALACT1$("ECUECON") <> "" Then
        Text2(Index).TEXT = VALACT1$("ECUECON")
        Text3(Index).TEXT = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    PRINTFORM
    Command5.Enabled = True
End Sub

Sub Command6_Click()
    If FECHANUM(Text4(0).TEXT) < 1 Then
        OPX% = COMALTER%("La Generación de Resumen Mensual\no se ha Completado.\  \Puede Ahora Re-Calcular el Resumen\o Abandonar el Proceso.\\Re-Calcular\Abandonar")
        If OPX% = 1 Then
            Call CALRECOM
        End If
        Exit Sub
    End If
    '
    Command6.Enabled = False
    Call BACKUPCOM(3, HAYERRO%)
    If HAYERRO% > 0 Then
        Command6.Enabled = True
        Exit Sub
    End If
    '
    Screen.MousePointer = 11
    X$ = REGLEIDO$("PROVED1", 1)
    Call GRAREG("PROVED1", X$, 1)
    Call CIERRARCH("PROVED1")
    Screen.MousePointer = 1
    '
    Call FINAL("*SUCOSOR")
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    '
    If FECHANUM(Text4(0).TEXT) < 1 Then
        OPX% = COMALTER%("La Generación de Resumen Mensual\no se ha Completado.\  \Puede Ahora Re-Calcular el Resumen\o Consultar en Forma Parcial.\\Re-Calcular\Consultar")
        If OPX% = 1 Then
            Call CALRECOM
            Exit Sub
        End If
    End If
    Call CONASIDIACOM
    '
End Sub
'

Private Sub Command8_Click()
    If FECHANUM(Text4(0).TEXT) < 1 Then
        OPX% = COMALTER%("La Generación de Resumen Mensual\no se ha Completado.\  \Puede Ahora Re-Calcular el Resumen\o Abandonar el Proceso.\\Re-Calcular\Abandonar")
        If OPX% = 1 Then
            Call CALRECOM
        End If
        Exit Sub
    End If
    '
    Command8.Enabled = False
    Call BACKUPCOM(3, HAYERRO%)
    If HAYERRO% > 0 Then
        Command8.Enabled = True
        Exit Sub
    End If
    '
    RESUMENCO.MESAN.TEXT = MESAN.TEXT
    RESUMENCO.Text4(0).TEXT = Text4(0).TEXT
    RESUMENCO.Text4(1).TEXT = Text4(1).TEXT
    Command8.Enabled = True
    RESUMENCO.Show 1
End Sub

Sub Command9_Click()
    If FECHANUM(Text4(0).TEXT) < 1 Then
        OPX% = COMALTER%("La Generación de Resumen Mensual\no se ha Completado.\  \Puede Ahora Re-Calcular el Resumen\o Abandonar el Proceso.\\Re-Calcular\Abandonar")
        If OPX% = 1 Then
            Call CALRECOM
        End If
        Exit Sub
    End If
    '
    Command9.Enabled = False
    Call BACKUPCOM(3, HAYERRO%)
    If HAYERRO% > 0 Then
        Command9.Enabled = True
        Exit Sub
    End If
    '
    Screen.MousePointer = 11
    X$ = REGLEIDO$("PROVED1", 1)
    Call GRAREG("PROVED1", X$, 1)
    Call CIERRARCH("PROVED1")
    Screen.MousePointer = 1
    '
    EPAC$ = EMPREAC$
    'If InStr(EPAC$, "SABO") > 0 Then
    '     Call FINAL("*SUIVASO")
    '   Else
         CMX% = COMALTER%("Emite el Subdiario IVA-Compras\\Completo\Solo Resumen\Cancelar")
         If CMX% = 1 Then
             Call CONECRUN("*LIVCOMEN", "Subdiario de IVA - Compras")
           ElseIf CMX% = 2 Then
             Call CONECRUN("*SOREIVA", "Subdiario de IVA - Compras")
         End If
    'End If
    '
    Command9.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Dim TETASA$(10)
    Call CENTRAFORM(Me)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    For IXX% = 1 To 4
      TXXX$ = TRIM$(Str$(IXX% - 1))
      TASA = Val(CONTROL$("", "TASIVA-" + TXXX$))
      If TASA > 0.005 Then
        TETASA$(IXX%) = TRIM$(CSTRING$(MKS$(TASA), 3, 6, 2, 2) + "%")
      End If
    Next IXX%
    TETASA$(5) = "Exento"
    TETASA$(6) = "Aduana"
    TETASA$(7) = "Monotr"
    TETASA$(8) = "I.Int."
    '
    TETASA$(1) = "19/21%"
    TETASA$(2) = "9/10,5%"
    '
    For IXX% = 1 To 8
      REMECOMP.TASIVA(IXX%).Caption = TETASA(IXX%)
    Next IXX%
    REMECOMP.TASIVA(0).Caption = "Total"
    '
    If InStr(EPAC$, "SABO") > 0 Then
      PERIBRU.ToolTipText = "Doble Click para Subdiarios (por Provincia)"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub

Private Sub Image2_Click()
   If USNBR% Mod 100 = 1 Then
      N5% = FILEOPEN%("USER.DAT", 9, 64)
      CAUS% = LOF(N5%) / 64
      Close #N5%
      If CAUS% > 1 Then
         Call MENSERR(24, "Usuario Supervisor no Autorizado\Para la Operación Solicitada.\   \Delegue el Derecho de Acceso\en Otro Usuario del Sistema.")
         Exit Sub
      End If
   End If
   '
   TXT$ = MESAN.TEXT
   If MESANO$(TXT$, DES%, HAS%) = TXT$ Then
      If DERACCE%("CIEMECOM", "Cierre y Re-Apertura Mensual Compras") >= 2 Then
         Call GRACONTROL("", "CIEMECOM", TRIM$(Str$(HAS%)))
         Image2.Visible = False
         Image1.Visible = True
      End If
   End If
End Sub

Private Sub Image1_Click()
   If USNBR% Mod 100 = 1 Then
      N5% = FILEOPEN%("USER.DAT", 9, 64)
      CAUS% = LOF(N5%) / 64
      Close #N5%
      If CAUS% > 1 Then
         Call MENSERR(24, "Usuario Supervisor no Autorizado\Para la Operación Solicitada.\   \Delegue el Derecho de Acceso\en Otro Usuario del Sistema.")
         Exit Sub
      End If
   End If
   '
   TXT$ = MESAN.TEXT
   If MESANO$(TXT$, DES%, HAS%) = TXT$ Then
      If DERACCE%("CIEMECOM", "Cierre y Re-Apertura Mensual Compras") >= 2 Then
         If COMALTER%("Confirme Re-Apertura del Período Mensual\\Cancelar\Re-Abrir") = 2 Then
            Call GRACONTROL("", "CIEMECOM", TRIM$(Str$(DES% - 1)))
            Image1.Visible = False
            Image2.Visible = True
         End If
      End If
   End If
End Sub

Private Sub IMPOGRA_Change(Index As Integer)
   If Index > 0 Then
      VALOR# = 0
      For KKI% = 1 To 8
         VALOR# = VALOR# + Val(REMECOMP.IMPOGRA(KKI%).Caption)
      Next KKI%
      REMECOMP.IMPOGRA(0).Caption = TRIM$(CSTRING$(MKD$(VALOR#), 4, 12, 2, 2))
      If Abs(Val(REMECOMP.IMPOGRA(Index).Caption)) > 0.005 Then
         TASARE = 100 * Val(REMECOMP.IMPOIVA(Index).Caption) / Val(REMECOMP.IMPOGRA(Index).Caption)
         REMECOMP.TAREAL(Index).Caption = TRIM$(CSTRING$(MKS$(TASARE), 4, 12, 2, 2))
         REMECOMP.TAREAL(Index).Refresh
      End If
     Else
      Call CALTOTA
   End If
   REMECOMP.IMPOGRA(Index).Refresh
End Sub

Private Sub IMPOIVA_Change(Index As Integer)
   If Index > 0 Then
      VALOR# = 0
      For KKI% = 1 To 8
         VALOR# = VALOR# + Val(REMECOMP.IMPOIVA(KKI%).Caption)
      Next KKI%
      REMECOMP.IMPOIVA(0).Caption = TRIM$(CSTRING$(MKD$(VALOR#), 4, 12, 2, 2))
      If Abs(Val(REMECOMP.IMPOGRA(Index).Caption)) > 0.005 Then
         TASARE = 100 * Val(REMECOMP.IMPOIVA(Index).Caption) / Val(REMECOMP.IMPOGRA(Index).Caption)
         REMECOMP.TAREAL(Index).Caption = TRIM$(CSTRING$(MKS$(TASARE), 4, 12, 2, 2))
         REMECOMP.TAREAL(Index).Refresh
      End If
     Else
      Call CALTOTA
   REMECOMP.TOTAGEN.Refresh
   End If
   REMECOMP.IMPOIVA(Index).Refresh
End Sub

Private Sub MESAN_Change()
   TXT$ = MESAN.TEXT
   If MESANO$(TXT$, DES%, HAS%) = TXT$ Then
      If Val(CONTROL$("", "CIEMECOM")) < HAS% Then
          Image1.Visible = False
          Image2.Visible = True
         Else
          Image2.Visible = False
          Image1.Visible = True
      End If
   End If
End Sub

Private Sub MESAN_DblClick()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) <> "" Then
        MESAN.TEXT = MESASEL$
        Call CARGARESCOM
        Call FRESHALL
    End If
    Call FRESHALL
End Sub

Private Sub MESAN_GOTFOCUS()
    '
    If CARGAMEN% = 0 Then
        '
        Refresh
        Screen.MousePointer = 11
        '
        CORRESP$ = TRIM$(CORRAR$("SUIVACO"))
        MESAN.TEXT = CORRESP$
        If TRIM$(MESAN.TEXT) = "" Then
            Call CARSELMESA
            MESAN.TEXT = Left$(REGLEIDO$("SELMESA", 1), 12)
            If TRIM$(MESAN.TEXT) = "" Then
                Call COMUNI("Imposible Inicializar Período Mensual.\Consulte a su Soporte de Sistemas.")
                Call FINAL("")
            End If
        End If
        '
        CUEIVA$ = CONTROL$("", "CUEIVA-0")
        Text2(0) = CUEIVA$
        Text3(0) = ECOARCH$("ECUECON", CUEIVA$)
        '
        CUEADU$ = CONTROL$("", "CUEIVA-4")
        Text2(1) = CUEADU$
        Text3(1) = ECOARCH$("ECUECON", CUEADU$)
        '
        CUEACRE$ = CONTROL$("SUIVACO", "CUEACRE")
        Text2(2) = CUEACRE$
        Text3(2) = ECOARCH$("ECUECON", CUEACRE$)
        '
        Call CARGARESCOM
        Call FRESHALL
        CARGAMEN% = 1
        '
    End If
    Screen.MousePointer = 1
    '
End Sub
'

Private Sub MESAN_KeyPress(KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
        Command2.SetFocus
    End If
    '
End Sub

Private Sub MESAN_LostFocus()
    '
    'MESA$ = MESAN.TEXT
    'If TRIM$(MESA$) = "" Then
    '    MESAN.SetFocus
    '    Exit Sub
    'End If
    '
    'MESB$ = MESANO$(MESA$, DES%, HAS%)
    'If MESB$ <> MESA$ Then
    '    If MESB$ <> "" Then
    '        MESAN.TEXT = MESB$
    '    End If
    '    MESAN.SetFocus
    '    Call CARGARESCOM
    '    If FECHANUM(Text4(0).TEXT) < 1 Then
    '        Command2.SetFocus
    '    End If
        '
    'End If
    '
End Sub

Private Sub PERGAN_Change()
   Call CALTOPER
End Sub

Private Sub PERIBRU_Change()
   Call CALTOPER
End Sub

Private Sub PERIBRU_DblClick()
   '
   Screen.MousePointer = 11
   Dim COPROVIN$(64), DEPROVIN$(64)
   '
   For U& = 1 To ULTREG&("PROVINB")
     If U& < 65 Then
       X$ = REGLEIDO$("PROVINB", U&)
       CAPROVI% = U&
       COPROVIN$(U&) = Left$(X$, 1)
       DEPROVIN$(U&) = Mid$(X$, 2, 25)
     End If
   Next U&
   '
   Call MESANO(MESAN, DES%, HAS%)
10 Call SELECHO("PROVINC")
   PSELE$ = TRIM$(VALACT1$("PROVINC"))
   DPRX$ = VALACT2$("PROVINC")
   If PSELE$ <> "" Then
     Screen.MousePointer = 11
     FIN& = ULTREG&("CUECORR")
     JJ& = 0
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       TS$ = REGLEIDO$("CUECOSS", U&)
       FF% = CVI(Left$(TS$, 2))
       If FF% >= DES% Then
         If FF% <= HAS% Then
           COPROV$ = Mid$(TS$, 493, 1)
           SIG% = Sgn(CVD(Mid$(TS$, 123, 8)) - CVD(Mid$(TS$, 115, 8)))
           PERBRU# = SIG% * Abs(CVD(Mid$(TS$, 107, 8)))
           If Abs(PERBRU#) > 0 Then
             NPRX% = CVI(Mid$(TS$, 5, 2))
             NFXX$ = TRIM$(Mid$(TS$, 7, 18))
             '
             If TRIM$(COPROV$) = "" Then
               DEPROVYI$ = TRIM$(Mid$(TS$, 494, 19))
               For KKI% = 1 To CAPROVI%
                 If TRIM$(UCase$(DEPROVYI$)) = TRIM$(UCase$(Left$(DEPROVIN$(KKI%), 19))) Then
                   COPROV$ = COPROVIN$(KKI%)
                   GoTo 52
                 End If
               Next KKI%
             End If
             '
52           If COPROV$ = PSELE$ Then
                  DEPROV$ = Mid$(TS$, 494, 19)
                  REBLA$ = REGBLAN$("SUPEIBFA")
                  Call REPLA(REBLA$, MKI$(FF%), 1, 2)
                  Call REPLA(REBLA$, MKI$(NPRX%), 3, 2)
                  Call REPLA(REBLA$, NFXX$, 5, 18)
                  Call REPLA(REBLA$, RASOCLI$((-1) * NPRX%), 23, 30)
                  Call REPLA(REBLA$, CUITCLI$((-1) * NPRX%), 53, 14)
                  Call REPLA(REBLA$, NIBRCLI$((-1) * NPRX%), 67, 14)
                  Call REPLA(REBLA$, MKD$(PERBRU#), 89, 8)
                  Call REPLA(REBLA$, COPROV$ + DEPROV$, 97, 20)
                    CUCO% = CUEIPROVIN%(COPROV$)
                  Call REPLA(REBLA$, MKI$(CUCO%), 117, 2)
                  JJ& = JJ& + 1
                  Call GRAREG("SUPEIBFA", REBLA$, JJ&)
               ElseIf COPROV$ = "*" Then
                  RFG$ = RECFILT$("PERCIBFA", 2, MKI$(NPRX%))
                  For UKY& = 1 To Len(RFG$) Step 4
                    RGY& = CVS(Mid$(RFG$, UKY&, 4))
                    RGZ$ = REGLEIDO$("PERCIBFA", RGY&)
                    If TRIM$(Mid$(RGZ$, 5, 18)) = TRIM$(NFXX$) Then
                      If CVI(Left$(RGZ$, 2)) = FF% Then
                        COJU$ = TRIM$(Mid$(RGZ$, 97, 1))
                        If COJU$ = PSELE$ Then
                          IMPE# = Abs(CVD(Mid$(RGZ$, 89, 8)))
                          If IMPE# >= 0.005 Then
                            REBLA$ = REGBLAN$("SUPEIBFA")
                            Call REPLA(REBLA$, MKI$(FF%), 1, 2)
                            Call REPLA(REBLA$, MKI$(NPRX%), 3, 2)
                            Call REPLA(REBLA$, NFXX$, 5, 18)
                            Call REPLA(REBLA$, RASOCLI$((-1) * NPRX%), 23, 30)
                            Call REPLA(REBLA$, CUITCLI$((-1) * NPRX%), 53, 14)
                            Call REPLA(REBLA$, NIBRCLI$((-1) * NPRX%), 67, 14)
                            Call REPLA(REBLA$, MKD$(SIG% * IMPE#), 89, 8)
                            Call REPLA(REBLA$, Mid$(RGZ$, 97, 20), 97, 20)
                              CUCO% = CUEIPROVIN%(COJU$)
                            Call REPLA(REBLA$, MKI$(CUCO%), 117, 2)
                            JJ& = JJ& + 1
                            Call GRAREG("SUPEIBFA", REBLA$, JJ&)
                          End If
                        End If
                      End If
                    End If
                  Next UKY&
             End If
           End If
         End If
       End If
     Next U&
     '
     If JJ& > 0 Then
         Call SETULTREG("SUPEIBFA", JJ&)
         Call HEADERS("SUPEIBFA", "")
         Call HEADERS("SUPEIBFA", "Corresponde a: " + MESAN + " - " + DPRX$)
         Call FINAL("*SUPEIBFA")
       Else
         Call MENSERR(24, "Sin Movimientos")
     End If
     Screen.MousePointer = 1
     GoTo 10
     '
   End If
   '
End Sub

Private Sub PERIVA_Change()
   Call CALTOPER
End Sub

Private Sub SUTOPER_Change()
   Call CALTOTA
End Sub

Private Sub Text1_Click(Index As Integer)
    MESAN.SetFocus
End Sub

Private Sub Text2_Change(Index As Integer)
    DECUE$ = ECOARCH("ECUECON", TRIM$(Text2(Index).TEXT))
    If DECUE$ <> "" Then
        Text3(Index).TEXT = DECUE$
      Else
        Text3(Index).TEXT = ""
    End If
End Sub

Private Sub Text2_DblClick(Index As Integer)
    Call SELECHO("ECUECON")
    If VALACT1$("ECUECON") <> "" Then
        Text2(Index).TEXT = VALACT1$("ECUECON")
        Text3(Index).TEXT = VALACT2$("ECUECON")
    End If
End Sub

Sub CALTOPER()
   VALOR# = 0
   VALOR# = VALOR# + Val(PERIVA.Caption)
   VALOR# = VALOR# + Val(PERGAN.Caption)
   VALOR# = VALOR# + Val(PERIBRU.Caption)
   REMECOMP.SUTOPER.Caption = TRIM$(CSTRING$(MKD$(VALOR#), 4, 12, 2, 2))
   REMECOMP.SUTOPER.Refresh
End Sub
'
Sub CALTOTA()
   VALOR# = 0
   VALOR# = VALOR# + Val(IMPOGRA(0).Caption)
   VALOR# = VALOR# + Val(IMPOIVA(0).Caption)
   VALOR# = VALOR# + Val(SUTOPER.Caption)
   REMECOMP.TOTAGEN.Caption = TRIM$(CSTRING$(MKD$(VALOR#), 4, 12, 2, 2))
   REMECOMP.TOTAGEN.Refresh
End Sub
