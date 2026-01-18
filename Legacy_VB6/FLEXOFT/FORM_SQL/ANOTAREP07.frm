VERSION 5.00
Begin VB.Form ANOTAREP07 
   BackColor       =   &H00EEDDEE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anotador de Ordenes de Reparación "
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   10940
      ScaleHeight     =   75
      ScaleWidth      =   810
      TabIndex        =   53
      Top             =   6615
      Width           =   870
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   54
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00EEDDEE&
      Caption         =   "Ver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   10900
      TabIndex        =   47
      Top             =   4020
      Width           =   885
      Begin VB.OptionButton Option4 
         BackColor       =   &H00EEDDEE&
         Caption         =   "O/R"
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
         Left            =   80
         TabIndex        =   52
         Top             =   1170
         Width           =   645
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00EEDDEE&
         Caption         =   "P-Ext"
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
         Left            =   80
         TabIndex        =   51
         Top             =   885
         Width           =   750
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00EEDDEE&
         Caption         =   "P-Int"
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
         Left            =   80
         TabIndex        =   50
         Top             =   600
         Width           =   750
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EEDDEE&
         Caption         =   "Plan"
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
         Left            =   80
         TabIndex        =   49
         Top             =   315
         Value           =   -1  'True
         Width           =   645
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   105
         Picture         =   "ANOTAREP07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Consultas y Re-Impresión"
         Top             =   1520
         Width           =   645
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00EEDDEE&
      Caption         =   "Planilla / Orden de Reparación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1550
      Left            =   105
      TabIndex        =   20
      Top             =   210
      Width           =   5370
      Begin VB.CommandButton Command4 
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
         Left            =   1995
         TabIndex        =   23
         Top             =   315
         Width           =   180
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cierr:"
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
         Left            =   3360
         TabIndex        =   42
         Top             =   1170
         Width           =   750
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4200
         TabIndex        =   41
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4830
         TabIndex        =   36
         Top             =   315
         Width           =   330
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "St:"
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
         Left            =   3465
         TabIndex        =   35
         Top             =   380
         Width           =   1275
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2625
         TabIndex        =   34
         Top             =   1125
         Width           =   855
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "O/V:"
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
         Left            =   1260
         TabIndex        =   33
         Top             =   1170
         Width           =   1275
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   945
         TabIndex        =   29
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente::"
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
         Left            =   105
         TabIndex        =   28
         Top             =   765
         Width           =   750
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   945
         TabIndex        =   27
         Top             =   720
         Width           =   4230
      End
      Begin VB.Label Label3 
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
         Height          =   330
         Left            =   1890
         TabIndex        =   26
         Top             =   380
         Width           =   1275
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   945
         TabIndex        =   25
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3255
         TabIndex        =   24
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Presup:"
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
         Index           =   1
         Left            =   -1050
         TabIndex        =   22
         Top             =   1170
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
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
         Index           =   0
         Left            =   -1050
         TabIndex        =   21
         Top             =   375
         Width           =   1905
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00EEDDEE&
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1550
      Left            =   5670
      TabIndex        =   17
      Top             =   210
      Width           =   5055
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ref:"
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
         Index           =   6
         Left            =   1680
         TabIndex        =   46
         Top             =   1170
         Width           =   1905
      End
      Begin VB.Label LENTREGA 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   3675
         TabIndex        =   45
         Top             =   1125
         Width           =   1185
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Talón:"
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
         Left            =   1890
         TabIndex        =   44
         Top             =   375
         Width           =   1905
      End
      Begin VB.Label LENTREGA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   3885
         TabIndex        =   43
         Top             =   315
         Width           =   975
      End
      Begin VB.Label LENTREGA 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   1155
         TabIndex        =   39
         Top             =   1125
         Width           =   1815
      End
      Begin VB.Label LENTREGA 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   1155
         TabIndex        =   38
         Top             =   720
         Width           =   3705
      End
      Begin VB.Label LENTREGA 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   1155
         TabIndex        =   37
         Top             =   315
         Width           =   1815
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Left            =   -840
         TabIndex        =   19
         Top             =   380
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Serie Nro:"
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
         Left            =   -840
         TabIndex        =   18
         Top             =   1170
         Width           =   1905
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEDDEE&
      Caption         =   "Novedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2060
      Left            =   105
      TabIndex        =   6
      Top             =   6510
      Width           =   10620
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
         Height          =   1275
         Left            =   1155
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   735
         Width           =   7155
      End
      Begin VB.CommandButton Command3 
         Caption         =   "O.K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8360
         Picture         =   "ANOTAREP07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1140
         Width           =   750
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0C0&
         Height          =   425
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   10515
         TabIndex        =   8
         Top             =   315
         Width           =   10580
         Begin VB.Line Line2 
            X1              =   0
            X2              =   1155
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Label Label5 
            Caption         =   "Label5"
            Height          =   15
            Left            =   0
            TabIndex        =   10
            Top             =   420
            Width           =   1380
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   9
            Top             =   75
            Width           =   10200
         End
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
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
         Left            =   8400
         TabIndex        =   32
         Top             =   840
         Width           =   4000
      End
      Begin VB.Line Line7 
         X1              =   10580
         X2              =   10580
         Y1              =   315
         Y2              =   1995
      End
      Begin VB.Line Line6 
         X1              =   8295
         X2              =   10605
         Y1              =   1130
         Y2              =   1130
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Registrar Anotación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   9345
         TabIndex        =   31
         Top             =   1365
         Width           =   960
      End
      Begin VB.Line Line5 
         X1              =   8295
         X2              =   8295
         Y1              =   735
         Y2              =   1995
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   10605
         Y1              =   2010
         Y2              =   2010
      End
      Begin VB.Label Label9 
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
         Height          =   1275
         Left            =   0
         TabIndex        =   11
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label10 
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
         Height          =   1280
         Left            =   8310
         TabIndex        =   12
         Top             =   735
         Width           =   2260
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEDDEE&
      Caption         =   "Anotaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4450
      Left            =   105
      TabIndex        =   4
      Top             =   1890
      Width           =   10620
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   3795
         Left            =   10290
         TabIndex        =   40
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   4005
         Left            =   8085
         TabIndex        =   16
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   4110
         Left            =   1155
         TabIndex        =   15
         Top             =   630
         Width           =   15
      End
      Begin VB.ListBox List1 
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
         Height          =   3660
         Left            =   0
         TabIndex        =   14
         Top             =   735
         Width           =   10580
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   10515
         TabIndex        =   5
         Top             =   315
         Width           =   10580
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   Usuario/Terminal"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   80
            TabIndex        =   30
            Top             =   105
            Width           =   10200
         End
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   10605
         Y1              =   4410
         Y2              =   4410
      End
      Begin VB.Line Line1 
         X1              =   10580
         X2              =   10580
         Y1              =   320
         Y2              =   6225
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   10920
      Picture         =   "ANOTAREP07.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprime Ficha de Ingenieria"
      Top             =   6930
      Width           =   855
   End
   Begin VB.CommandButton Command2 
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
      Height          =   800
      Left            =   10920
      Picture         =   "ANOTAREP07.frx":0C7E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1155
      Width           =   855
   End
   Begin VB.CommandButton Command9 
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
      Height          =   800
      Left            =   10920
      Picture         =   "ANOTAREP07.frx":0F88
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Sale de la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10340
      Picture         =   "ANOTAREP07.frx":10D2
      Top             =   8020
      Width           =   2010
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
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
      Left            =   -105
      TabIndex        =   3
      Top             =   2625
      Width           =   1695
   End
End
Attribute VB_Name = "ANOTAREP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command12_Click()

End Sub

Private Sub Command26_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   NPX& = Val(Label7)
   norventa& = Val(Label16)
   If NPX& > 0 Then
     If TRIM$(Text1) <> "" Then
       Call FRATEXTO(Text1, 56)
       For KKI% = 1 To CARETEX%
         TTXX$ = Space$(76)
         Call REPLA(TTXX$, RETEX$(KKI%), 11, 56)
         If KKI% = 1 Then
           Call REPLA(TTXX$, HOS$, 1, 8)
           Call REPLA(TTXX$, USERTER$, 69, 24)
         End If
'
         Call REGAPP("ANOTAREP", Left$(MKS$(NPX&) + TTXX$, 96))

         If norventa& > 0 Then
           '\\\OT-08-0817-OD-08/09/08///
           'Call REGAPP("ANOTAPED", Left$(MKS$(norventa&) + TTXX$, 96))
           Call AGREGA_ANOTAPED(norventa&, TTXX$)
           '\\\OT-08-0817-OD-FIN///
         End If
       Next KKI%
       Call REGAPP("ANOTAREP", MKS$(NPX&) + String$(88, "="))
       If norventa& > 0 Then
         'Call REGAPP("ANOTAPED", MKS$(norventa&) + String$(88, "="))
         Call AGREGA_ANOTAPED(norventa&, String$(88, "="))
         '\\\OT-08-0817-OD-FIN///
       End If
       Call CIERRARCH("ANOTAREP")
       Call CIERRARCH("ANOTAPED")
       Text1 = ""
       '
       List1.Clear
       RFF$ = RECFILT$("ANOTAREP", 1, MKS$(NPX&))
       '
       For U& = 1 To Len(RFF$) Step 4
         RENOTA& = CVS(Mid$(RFF$, U&, 4))
         XX$ = REGLEIDO$("ANOTAREP", RENOTA&)
         List1.AddItem Mid$(XX$, 5, 92)
       Next U&
       '
     End If
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call SELECHO("PRESREPA/Indice de Planillas de Reparación")
   NPX& = Val(VALACT1$("PRESREPA"))
   If NPX& < 1 Then GoTo 99
   '
   Label7 = TRIM$(Str$(NPX&))
99 Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   PRINTFORM
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   NPX& = Val(Label7)
   '
   If NPX& > 0 Then
       If Option1.Value = True Then
            TIDOCX$ = TRIM$("Planilla de Reparación")
          ElseIf Option2.Value = True Then
            '
            If DERACCE%("VEPREPAR", "Visualizar Presupuestos de Reparación") < 2 Then
              GoTo 99
            End If
            '
            TIDOCX$ = TRIM$("Presupuesto de Reparación")
          ElseIf Option3.Value = True Then
            '
            If DERACCE%("VEPREPAR", "Visualizar Presupuestos de Reparación") < 2 Then
              GoTo 99
            End If
            '
            TIDOCX$ = TRIM$("Presup.Reparación (Externo)")
          ElseIf Option4.Value = True Then
            TIDOCX$ = TRIM$("Orden de Reparación")
      End If
      '
      Screen.MousePointer = 11
      NUDOCU$ = TRIM$(Str$(NPX&))
      TIDOCU$ = TIDOCX$ + " " + NUDOCU$ + " "
      '
      For UI& = ULTREG&("PDOCLST") To 1 Step -1
        XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
        If InStr(XX$, TIDOCU$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      Next UI&
      '
      Screen.MousePointer = 1
      Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a esta Reparación.")
      On Error Resume Next
      GoTo 99
      '
      '28/03/07 (OT-07-0113)
5     Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
      '28/03/07 (FIN)
'5     For KKI% = 0 To 5
'        EDITFORM.Picture1(KKI%).Cls
'        EDITFORM.Picture1(KKI%).DrawWidth = 3
'        EDITFORM.Picture1(KKI%).DrawStyle = 0
'        EDITFORM.Picture1(KKI%).FillStyle = 1
'      Next KKI%
'      '
'      REDOCU$ = REGACT$("PDOCLST")
'      URE& = ULTREG&("PDOCSPL")
'      LI& = 0: LS& = URE&
'6     E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'      L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'      If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'      GoTo 8
'      '
'7     Screen.MousePointer = 1
'      Call MENSERR(24, "Documento no Registrado")
'      GoTo 99
'      '
'8     TPSP$ = String$(2048, 0): PINSE& = 1
'      Screen.MousePointer = 11
'      While L& > 1
'        DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'        If DCC& < DOCUNU& Then GoTo 9
'        L& = L& - 1
'      Wend
'      '
'9     CAPAGINAS% = 0
'      For KKL& = L& To URE&
'        TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'        If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'        If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TBUF$ = Mid$(TBUF$, 5)
'          If PINSE& + Len(TBUF$) > Len(TPSP$) Then TPSP$ = TPSP$ + String$(2048, 0)
'          Mid$(TPSP$, PINSE&, Len(TBUF$)) = TBUF$
'          PINSE& = PINSE& + Len(TBUF$)
'        End If
'      Next KKL&
'      '
'95    SPOOLSTRING$ = Left$(TPSP$, PINSE& - 1)
'      FIN& = Len(SPOOLSTRING$)
'      For KKI% = 0 To 5
'        EDITFORM.Picture1(KKI%).Cls
'        EDITFORM.Picture1(KKI%).Top = 0
'        EDITFORM.Picture1(KKI%).Refresh
'        EDITFORM.Picture1(KKI%).Height = 7170
'        EDITFORM.Picture1(KKI%).Width = 11700
'        EDITFORM.Picture1(KKI%).DrawWidth = 3
'        EDITFORM.Picture1(KKI%).DrawStyle = 0
'        EDITFORM.Picture1(KKI%).FillStyle = 1
'      Next KKI%
'      ALTUPAGINA = 7170
'      TOPEPAGINA = 0
'      PAGINACTIVA% = 0
'      CAPAGINAS% = 0
'      '
'10    PPXX% = InStr(TPSP$, Chr$(255))
'      If PPXX% > 0 Then
'        COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'        TPSP$ = Mid$(TPSP$, PPXX% + 1)
'        Call TRACE(20, FIN& - Len(TPSP$), FIN&)
'        '
'        If Len(COPRIAT$) > 0 Then
'          COPRI% = Asc(Left$(COPRIAT$, 1))
'          ATRIP$ = Mid$(COPRIAT$, 2)
'          Call SHOWCOMMAND(COPRI%, ATRIP$)
'        End If
'        GoTo 10
'        '
'      End If
'      '
'      Screen.MousePointer = 1
'      PAGINACTIVA% = 0
'      For KKI% = 0 To 5
'         EDITFORM.Picture1(KKI%).Visible = False
'      Next KKI%
'      EDITFORM.Picture1(PAGINACTIVA%).Visible = True
'      EDITFORM.Picture1(PAGINACTIVA%).Refresh
'      EDITFORM.VScroll1.Min = 0
'      EDITFORM.VScroll1.Value = 0
'      EDITFORM.VScroll1.Max = 1
'      If ALTUPAGINA > EDITFORM.Frame1.Height Then
'         EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'      End If
'      EDITFORM.Command3.Enabled = False
'      EDITFORM.Command5.Enabled = False
'      EDITFORM.Show 1
'      '
    End If
'    '
99  Command6.Enabled = True
    '
   End Sub

Private Sub Command9_Click()
   Label7 = ""
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub Form_Load()
    '
    If Screen.Height > 9000 Then
       Top = (Screen.Height - Height) / 2
       Left = (Screen.Width - Width) / 2
    End If
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    If UCase$(Left$(App.EXEName, 6)) = "REPABO" Then
      Label15 = "O/R:"
    End If
    '
    Screen.MousePointer = 1
    HOII% = HOY(HOS$)
    Label9 = " " + HOS$
    Label12 = Left$(USERNAME, 17)
    '
End Sub

Private Sub Label7_Change()
   '
   NPX& = Val(Label7)
   Label8 = ""
   Label11 = ""
   Label14 = ""
   Label16 = ""
   Label18 = ""
   Label19 = ""
   For KKI% = 0 To 4
     LENTREGA(KKI%) = ""
   Next KKI%
   Text1 = ""
   List1.Clear
   Option1.Value = True
   Option1.Enabled = False
   Option2.Enabled = False
   Option3.Enabled = False
   Option4.Enabled = False
   '
   RFF$ = RECFILT$("PRESREPA", 1, MKS$(NPX&))
   If Len(RFF$) >= 4 Then
     REPED& = CVS(Left$(RFF$, 4))
     XX$ = REGLEIDO$("PRESREPA", REPED&)
     FEX = CVI(Mid$(XX$, 65, 2))
     Label8 = FECHATEX$(FEX)
     Label18 = "1"
     NCLIE% = CVI(Mid$(XX$, 67, 2))
     Label11 = RASOCLI$(NCLIE%)
     FEX = CVI(Mid$(XX$, 121, 2))
     If FEX > 0 Then Label18 = "2"
     Label14 = FECHATEX$(FEX)
     Label16 = CSTRING$(Mid$(XX$, 125, 4), 3, 6, 0, 2)
     If Val(Label16) > 0 Then Label18 = "3"
     FEX = CVI(Mid$(XX$, 123, 2))
     If FEX > 0 Then Label18 = "4"
     Label19 = FECHATEX$(FEX)
     '
     CIXI% = CVI(Mid$(XX$, 69, 2))
     LENTREGA(0) = CODEXT$(CIXI%)
     LENTREGA(1) = DESCRIT0$(CIXI%)
     LENTREGA(2) = Mid$(XX$, 71, 15)
     LENTREGA(3) = TRIM$(Mid$(XX$, 95, 6))
     LENTREGA(4) = TRIM$(Mid$(XX$, 86, 9))
     '
     If Val(Label18) > 0 Then
       Option1.Enabled = True
       Option1.Value = True
     End If
     If Val(Label18) > 1 Then
       Option2.Enabled = True
       Option3.Enabled = True
       Option3.Value = True
     End If
     If Val(Label18) > 2 Then
       Option4.Enabled = True
       Option4.Value = True
     End If
   End If
   '
20 RFF$ = RECFILT$("ANOTAREP", 1, MKS$(NPX&))
   For U& = 1 To Len(RFF$) Step 4
     RENOTA& = CVS(Mid$(RFF$, U&, 4))
     XX$ = REGLEIDO$("ANOTAREP", RENOTA&)
     List1.AddItem Mid$(XX$, 5, 92)
   Next U&
   '
End Sub

