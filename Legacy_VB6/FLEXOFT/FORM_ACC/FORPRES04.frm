VERSION 5.00
Begin VB.Form FORPRES04 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuestos - Generación y Emisión"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   11910
   Icon            =   "FORPRES04.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   11910
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2325
      Left            =   0
      TabIndex        =   37
      Top             =   5565
      Width           =   10905
      Begin VB.TextBox TICAMBIS 
         Enabled         =   0   'False
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
         Left            =   6510
         MaxLength       =   14
         TabIndex        =   65
         Top             =   840
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Otros"
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
         Height          =   1065
         Left            =   5985
         TabIndex        =   58
         Top             =   1260
         Width           =   4830
         Begin VB.TextBox Total 
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
            Index           =   6
            Left            =   1995
            TabIndex        =   29
            Text            =   " "
            Top             =   250
            Width           =   2595
         End
         Begin VB.TextBox Total 
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
            Index           =   5
            Left            =   1995
            TabIndex        =   30
            Text            =   " "
            Top             =   630
            Width           =   2595
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Validez de Oferta:"
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
            Left            =   0
            TabIndex        =   60
            Top             =   315
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Garantia:"
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
            Left            =   0
            TabIndex        =   59
            Top             =   705
            Width           =   1905
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Entrega"
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
         Height          =   1065
         Left            =   210
         TabIndex        =   55
         Top             =   1260
         Width           =   5580
         Begin VB.TextBox Total 
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
            Index           =   4
            Left            =   1155
            TabIndex        =   28
            Text            =   " "
            Top             =   630
            Width           =   4275
         End
         Begin VB.TextBox Total 
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
            Left            =   1155
            TabIndex        =   27
            Text            =   " "
            Top             =   250
            Width           =   4275
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Lugar:"
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
            Left            =   -840
            TabIndex        =   57
            Top             =   680
            Width           =   1905
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Plazo:"
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
            TabIndex        =   56
            Top             =   315
            Width           =   1905
         End
      End
      Begin VB.TextBox PORDESCU 
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
         Left            =   7560
         MaxLength       =   14
         TabIndex        =   24
         Top             =   420
         Width           =   1065
      End
      Begin VB.TextBox Total 
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
         Left            =   9080
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   " "
         Top             =   405
         Width           =   1500
      End
      Begin VB.TextBox Total 
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
         Left            =   9080
         TabIndex        =   23
         Text            =   " "
         Top             =   0
         Width           =   1500
      End
      Begin VB.TextBox Total 
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
         Index           =   3
         Left            =   9080
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   1500
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
         Left            =   2100
         TabIndex        =   43
         Text            =   " "
         Top             =   945
         Width           =   3690
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
         Left            =   1365
         TabIndex        =   19
         Text            =   " "
         Top             =   945
         Width           =   435
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
         Height          =   330
         Index           =   2
         Left            =   1785
         TabIndex        =   20
         Top             =   945
         Width           =   180
      End
      Begin VB.TextBox Eco21 
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
         Index           =   1
         Left            =   2100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Text            =   "FORPRES04.frx":030A
         Top             =   380
         Width           =   3690
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
         Index           =   1
         Left            =   1365
         TabIndex        =   17
         Text            =   " "
         Top             =   380
         Width           =   435
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
         Left            =   1785
         TabIndex        =   18
         Top             =   380
         Width           =   180
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
         Index           =   0
         Left            =   2100
         TabIndex        =   39
         Text            =   " "
         Top             =   0
         Width           =   3690
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
         Index           =   0
         Left            =   1365
         TabIndex        =   15
         Text            =   " "
         Top             =   0
         Width           =   435
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
         Index           =   0
         Left            =   1785
         TabIndex        =   16
         Top             =   0
         Width           =   180
      End
      Begin VB.Label PORIVA1 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   6825
         TabIndex        =   96
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label PORIVA0 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   6825
         TabIndex        =   95
         Top             =   1365
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label IMPIVA105 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   5880
         TabIndex        =   94
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label IMPIVA21 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   5880
         TabIndex        =   93
         Top             =   1365
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   5880
         TabIndex        =   88
         Top             =   630
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label TOTPER 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   5985
         TabIndex        =   71
         Top             =   315
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label TOTIVA 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   5985
         TabIndex        =   70
         Top             =   0
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "T-C:"
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
         Left            =   6050
         TabIndex        =   64
         Top             =   915
         Visible         =   0   'False
         Width           =   435
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
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   5985
         TabIndex        =   63
         ToolTipText     =   "Doble-Click para 'Dólares'"
         Top             =   420
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label7 
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
         Height          =   330
         Left            =   8715
         TabIndex        =   47
         Top             =   510
         Width           =   225
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Neto:"
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
         Left            =   7560
         TabIndex        =   46
         Top             =   915
         Width           =   1380
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento:"
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
         Index           =   1
         Left            =   5985
         TabIndex        =   45
         Top             =   175
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal:"
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
         Left            =   7035
         TabIndex        =   44
         Top             =   75
         Width           =   1905
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
         Left            =   -2835
         TabIndex        =   42
         Top             =   1000
         Width           =   4110
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C.Pago:"
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
         Left            =   -2835
         TabIndex        =   40
         Top             =   420
         Width           =   4110
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L.Precios:"
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
         Left            =   -315
         TabIndex        =   38
         Top             =   75
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   12000
      Begin VB.PictureBox Picture6 
         Height          =   8300
         Left            =   11025
         ScaleHeight     =   8235
         ScaleWidth      =   1215
         TabIndex        =   102
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
         Begin VB.CommandButton Command17 
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
            Height          =   600
            Left            =   150
            Picture         =   "FORPRES04.frx":030C
            Style           =   1  'Graphical
            TabIndex        =   103
            ToolTipText     =   "Aprueba, Graba e Imprime Presupuesto"
            Top             =   7250
            Width           =   600
         End
         Begin VB.Image Image3 
            Height          =   8520
            Left            =   0
            Picture         =   "FORPRES04.frx":0616
            Stretch         =   -1  'True
            Top             =   -120
            Width           =   1065
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Formato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   1485
         Left            =   7140
         TabIndex        =   97
         Top             =   105
         Width           =   1485
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Por Medida"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   330
            Left            =   105
            MaskColor       =   &H80000012&
            TabIndex        =   99
            Top             =   1050
            Visible         =   0   'False
            Width           =   1300
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Item Unico"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   11.25
               Charset         =   0
               Weight          =   900
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   750
            Left            =   45
            TabIndex        =   98
            ToolTipText     =   "Doble-Click para 'Lista Repuestos'"
            Top             =   315
            Width           =   1380
         End
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
         Left            =   9120
         MaxLength       =   2
         TabIndex        =   91
         Top             =   525
         Width           =   450
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
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
         Left            =   9700
         TabIndex        =   89
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
         BackColor       =   &H80000009&
         Height          =   3680
         Left            =   210
         ScaleHeight     =   3615
         ScaleWidth      =   10560
         TabIndex        =   66
         Top             =   1800
         Visible         =   0   'False
         Width           =   10620
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
            ForeColor       =   &H00000000&
            Height          =   3210
            Left            =   -30
            TabIndex        =   69
            Top             =   450
            Width           =   10575
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00E0E0E0&
            Height          =   490
            Left            =   -30
            ScaleHeight     =   435
            ScaleWidth      =   10515
            TabIndex        =   67
            Top             =   -30
            Width           =   10570
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "  Cant Código          Descripción                    Lista  Desc  P.Unit     Importe"
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
               Left            =   0
               TabIndex        =   68
               ToolTipText     =   "Haga Doble-Click para Cargar Detalle."
               Top             =   120
               Width           =   10515
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
            X1              =   10550
            X2              =   10550
            Y1              =   0
            Y2              =   3920
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Artículo"
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
         Height          =   700
         Left            =   210
         TabIndex        =   61
         Top             =   4620
         Width           =   8625
         Begin VB.TextBox Text13 
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
            Left            =   210
            TabIndex        =   21
            Text            =   " "
            Top             =   280
            Width           =   2325
         End
         Begin VB.CommandButton Command18 
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
            Height          =   300
            Left            =   2520
            TabIndex        =   22
            Top             =   280
            Width           =   180
         End
         Begin VB.Label Label8 
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
            Height          =   300
            Left            =   2940
            TabIndex        =   62
            Top             =   285
            Width           =   5520
         End
      End
      Begin VB.TextBox Text28 
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
         Left            =   1470
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   52
         Top             =   2215
         Width           =   7890
      End
      Begin VB.CommandButton Command10 
         Height          =   260
         Left            =   420
         TabIndex        =   51
         Top             =   6855
         Width           =   2825
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   210
         ScaleHeight     =   375
         ScaleWidth      =   10545
         TabIndex        =   49
         Top             =   1785
         Width           =   10600
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción General - Especificaciones"
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
            Height          =   225
            Left            =   0
            TabIndex        =   50
            Top             =   105
            Width           =   10515
         End
      End
      Begin VB.TextBox Text12 
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
         MaxLength       =   40
         TabIndex        =   6
         Text            =   " "
         Top             =   1280
         Width           =   3690
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
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
         Left            =   9975
         TabIndex        =   13
         Text            =   " "
         Top             =   890
         Width           =   870
      End
      Begin VB.TextBox Text10 
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
         Left            =   9345
         TabIndex        =   14
         Text            =   " "
         Top             =   1280
         Width           =   1500
      End
      Begin VB.TextBox Text9 
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
         Left            =   9720
         TabIndex        =   11
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
         Left            =   10665
         TabIndex        =   12
         Top             =   105
         Width           =   180
      End
      Begin VB.TextBox Text7 
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
         Left            =   5355
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   490
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FFFFFF&
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
         Left            =   5355
         TabIndex        =   10
         Text            =   " "
         Top             =   1280
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
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
         Left            =   6855
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   105
         Width           =   180
      End
      Begin VB.TextBox Text6 
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
         Left            =   6435
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   105
         Width           =   435
      End
      Begin VB.TextBox Text5 
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
         Left            =   2310
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   890
         Width           =   2850
      End
      Begin VB.TextBox Text4 
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
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   890
         Width           =   750
      End
      Begin VB.TextBox Text3 
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
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   490
         Width           =   3690
      End
      Begin VB.TextBox Text2 
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
         Top             =   890
         Width           =   180
      End
      Begin VB.TextBox Text1 
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
         Left            =   210
         TabIndex        =   0
         Text            =   " "
         Top             =   890
         Width           =   855
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00404040&
         Height          =   8520
         Left            =   11025
         ScaleHeight     =   8460
         ScaleWidth      =   1005
         TabIndex        =   72
         Top             =   0
         Width           =   1065
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   130
            ScaleHeight     =   75
            ScaleWidth      =   540
            TabIndex        =   100
            Top             =   7200
            Width           =   600
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   -315
               TabIndex        =   101
               Top             =   105
               Width           =   12000
            End
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
            Height          =   600
            Left            =   130
            Picture         =   "FORPRES04.frx":29AE
            Style           =   1  'Graphical
            TabIndex        =   83
            ToolTipText     =   "Opciones Varias- Consultas y Listados"""
            Top             =   3330
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
            Left            =   130
            Picture         =   "FORPRES04.frx":2AF8
            Style           =   1  'Graphical
            TabIndex        =   82
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
            Height          =   600
            Left            =   130
            Picture         =   "FORPRES04.frx":2E02
            Style           =   1  'Graphical
            TabIndex        =   81
            ToolTipText     =   "Aprueba, Graba e Imprime Presupuesto"
            Top             =   6450
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
            Height          =   600
            Left            =   130
            Picture         =   "FORPRES04.frx":310C
            Style           =   1  'Graphical
            TabIndex        =   80
            ToolTipText     =   "Vista Previa de Impresion"
            Top             =   5850
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
            Height          =   600
            Left            =   130
            Picture         =   "FORPRES04.frx":3416
            Style           =   1  'Graphical
            TabIndex        =   79
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   3990
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
            Height          =   600
            Left            =   130
            Picture         =   "FORPRES04.frx":3858
            Style           =   1  'Graphical
            TabIndex        =   78
            ToolTipText     =   "Acceso a Base de Datos de Clientes"
            Top             =   2070
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
            Height          =   600
            Left            =   130
            Picture         =   "FORPRES04.frx":3C9A
            Style           =   1  'Graphical
            TabIndex        =   77
            ToolTipText     =   "Acceso a Maestro de Artículos"
            Top             =   1470
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
            Height          =   600
            Left            =   130
            Picture         =   "FORPRES04.frx":3FA4
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Generación Presupuesto por Copia"
            Top             =   5190
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
            Height          =   600
            Left            =   130
            Picture         =   "FORPRES04.frx":42AE
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Condiciones Comerciales: Listas de Precios / Condiciones de Pago"
            Top             =   4590
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
            Height          =   600
            Left            =   130
            Picture         =   "FORPRES04.frx":43F8
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   2730
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
            Left            =   130
            Picture         =   "FORPRES04.frx":4702
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   100
            Width           =   600
         End
         Begin VB.Image Image2 
            Height          =   390
            Left            =   -300
            Picture         =   "FORPRES04.frx":484C
            Top             =   7455
            Width           =   1515
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad"
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
         Height          =   700
         Left            =   8925
         TabIndex        =   85
         Top             =   4620
         Width           =   1905
         Begin VB.CommandButton Command16 
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
            Height          =   300
            Left            =   2520
            TabIndex        =   87
            Top             =   280
            Width           =   180
         End
         Begin VB.TextBox Text14 
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
            Left            =   240
            TabIndex        =   86
            Text            =   " "
            Top             =   280
            Width           =   1500
         End
      End
      Begin VB.Label Label16 
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
         Height          =   300
         Left            =   10500
         TabIndex        =   92
         Top             =   525
         Width           =   330
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Id:         -            -"
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
         Height          =   330
         Left            =   8820
         TabIndex        =   90
         Top             =   570
         Width           =   2010
      End
      Begin VB.Label Label26 
         BackColor       =   &H00EEEEEE&
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
         Left            =   210
         TabIndex        =   84
         Top             =   1280
         Width           =   1035
      End
      Begin VB.Line Line5 
         X1              =   210
         X2              =   10815
         Y1              =   4545
         Y2              =   4545
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   12000
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   2325
         Left            =   9360
         TabIndex        =   54
         Top             =   2220
         Width           =   1450
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   2325
         Left            =   210
         TabIndex        =   53
         Top             =   2220
         Width           =   1275
      End
      Begin VB.Line Line7 
         X1              =   210
         X2              =   10810
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono:"
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
         Height          =   330
         Left            =   5355
         TabIndex        =   48
         Top             =   975
         Width           =   1695
      End
      Begin VB.Line Line8 
         X1              =   10810
         X2              =   10810
         Y1              =   1785
         Y2              =   4540
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   420
         Picture         =   "FORPRES04.frx":676E
         ToolTipText     =   "Click Aquí para cambiar Modalidad de Facturación"
         Top             =   105
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Presup.Nro:"
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
         Height          =   330
         Left            =   8120
         TabIndex        =   36
         Top             =   945
         Width           =   1800
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ref:"
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
         Height          =   330
         Left            =   7455
         TabIndex        =   35
         Top             =   1365
         Width           =   1800
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Height          =   330
         Left            =   8925
         TabIndex        =   34
         Top             =   150
         Width           =   750
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   5460
         TabIndex        =   33
         Top             =   150
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   0
         TabIndex        =   32
         Top             =   630
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
      Caption         =   "Archivo"
      Begin VB.Menu nuepre 
         Caption         =   "Nuevo Presupuesto"
      End
      Begin VB.Menu php0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuvp 
         Caption         =   "Vista Previa de Impresion"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMnv 
         Caption         =   "Marcar Presupuesto Como No Vigente"
      End
      Begin VB.Menu m 
         Caption         =   "-"
      End
      Begin VB.Menu MaAr 
         Caption         =   "Maestro de Articulos"
      End
      Begin VB.Menu MaCli 
         Caption         =   "Maestro de Clientes"
      End
      Begin VB.Menu php1 
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
      Begin VB.Menu ConReimp 
         Caption         =   "Consultas y Re - impresión"
      End
      Begin VB.Menu LisRang 
         Caption         =   "Listados por Rango de Fecha"
      End
   End
   Begin VB.Menu ConfiGral 
      Caption         =   "Configuración"
      Begin VB.Menu Configur 
         Caption         =   "Configuración  General"
      End
      Begin VB.Menu php4 
         Caption         =   "-"
      End
      Begin VB.Menu LisPreVen 
         Caption         =   "Listas de Precios de Venta"
      End
      Begin VB.Menu CondPag 
         Caption         =   "Tablas de Condiciones de Pago"
      End
   End
   Begin VB.Menu fp 
      Caption         =   "Formato de Presupuesto"
      Begin VB.Menu itemunico 
         Caption         =   "Item  Único"
         Checked         =   -1  'True
      End
      Begin VB.Menu listaderepuestos 
         Caption         =   "Lista de Repuestos"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu ayud 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "FORPRES04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHIPRE$
Dim ALPORCOPIA%

Sub PRIFACPRO()
    FORIPRE$ = TRIM$(CONTROL$("", "FORFAPRO"))
    FPXX$ = "FORFAPRO"
'    If Picture1.Visible = True Then
'       FORIPRE$ = TRIM$(CONTROL$("", "FORPRERE"))
'       FPXX$ = "FORPRERE"
'    End If

    If FORIPRE$ = "" Then
       Call MENSERR(24, "Imposible Imprimir Presupuesto.\Falta Definir Formulario Electrónico.")
       Exit Sub
    End If
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Presupuesto.\No existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
10  NC% = XVALO(Text1.TEXT)
    TIPODOC$ = "Factura Proforma"
    DESTIDOC% = NC%
    NUMEDOC$ = TRIM$(Text11.TEXT)
      If VIPREVIA% = 1 Then NUMEDOC$ = "(VP-" + NUMEDOC$ + ")"
    FECHDOC$ = TRIM$(Text9.TEXT)
    '
    'SE LEE LA INFORMACION DE PROFORMA
    Call ABREPRESVEN
    SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM PRESUVEN "
    SQLX$ = SQLX$ + "WHERE NroPres = " & XVALO(Text11)
    '
    Set LISPRETEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    With LISPRETEMP
      .MoveFirst
      BULTOS$ = SAFETEXT(!CantBultos)
      PNETO$ = SAFETEXT(!PesoNeto)
      PTOTAL$ = SAFETEXT(!PesoTotal)
      NCMX$ = SAFETEXT(!Ncm)
      CONDVENTA$ = SAFETEXT(!CondVta)
      .Close
    End With
    
    '
    CODPARAM$(1) = "REF-MAT3": DOCPARAM$(1) = TRIM$(Text10.TEXT)
    CODPARAM$(2) = "REF-MAT2": DOCPARAM$(2) = TRIM$(Text12.TEXT)
    CODPARAM$(3) = "PROV-CLI": DOCPARAM$(3) = NPROCLI$(NC%)
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
    CODPARAM$(21) = "MEDIDA-1": DOCPARAM$(21) = EMAILCLI(NC%) ' E-MAIL
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
    CODPARAM$(28) = "REF-MAT4": DOCPARAM$(28) = BULTOS$
    CODPARAM$(29) = "REF-MAT5": DOCPARAM$(29) = PNETO$
    CODPARAM$(30) = "REF-MAT6": DOCPARAM$(30) = PTOTAL$
    CODPARAM$(31) = "MEDIDA-3": DOCPARAM$(31) = NCMX$
    CODPARAM$(32) = "MEDIDA-2": DOCPARAM$(32) = CONDVENTA$
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
    CAPARDOC% = 32
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
      '
      CODTABU1$(1) = "CANTIDAD"
      CODTABU1$(2) = "COD-MAT"
      CODTABU1$(3) = "DES-MAT"
      CODTABU1$(4) = "PRE-NETO"
      CODTABU1$(5) = "NET-ITEM"
      CODTABU1$(6) = "N-PLANO"
      CODTABU1$(7) = "R-PLANO"
      CODTABU1$(8) = "F-PLANO"
      CACOLTAB1% = 8
      '
      CAITAB1% = 0
      For U& = 1 To ULTREG&("!CARDETPR")
        X$ = REGLEIDO$("!CARDETPR", U&)
        CIIX% = CVI(Mid$(X$, 1, 2))
        CAN = CVS(Mid$(X$, 47, 4))
        INETO = CVS(Mid$(X$, 39, 4)) * (100 - CVS(Mid$(X$, 69, 4))) / 100
        PRELIX = PRELISTA#(ARCHIPRE$, CIIX%)
        DESCULISTA# = CVS(Mid$(X$, 85, 4)) 'Descuento de lista
        ' Suma de los descuentos de grupo cliente/linea + Articulo/Cliente
        DescFamGruArtClien# = CVS(Mid$(X$, 89, 4))
        If CAN > 0.05 Then
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = Str$(CAN)
          TETABU1$(2, CAITAB1%) = CODEXT$(CIIX%)
          TETABU1$(3, CAITAB1%) = DESCRIT0$(CIIX%)
          TETABU1$(4, CAITAB1%) = FORMATNUM$(INETO, "F12.4")
          TETABU1$(5, CAITAB1%) = FORMATNUM$(INETO * CAN, "F12.2")
          TETABU1$(6, CAITAB1%) = FORMATNUM$(PRELIX, "F12.2")
          TETABU1$(7, CAITAB1%) = FORMATNUM$(DESCULISTA#, "F4.1")
          TETABU1$(8, CAITAB1%) = FORMATNUM$(DescFamGruArtClien#, "F4.1")
          COMPDESCRI$ = ""
          If CONTROL$("PEDIDO", "VARCOLOR") = "SI" Then
            COLORBA$ = TRIM$(Mid$(X$, 129, 12))
            TELA$ = TRIM$(Mid$(X$, 141, 12))
            COLTELA$ = TRIM$(Mid$(X$, 153, 12))
            If COLORBA$ <> "" Or TELA$ <> "" Or COLTELA$ <> "" Then
              COMPDESCRI$ = "CB: " + TRIM$(COLORBA$) + " - T: " + TRIM$(TELA$) + " - CT: " + TRIM$(COLTELA$)
              CAITAB1% = CAITAB1% + 1
              TETABU1$(3, CAITAB1%) = COMPDESCRI$
            End If
          End If
        End If
      Next U&
    End If
    '
    NOCONFIRM% = 1
    Call PRINTDOC("@FAPR-JYM")
    '
'    If Picture1.Visible = False Then
'        CIXI% = CODINT%(Text13.TEXT)
'        DOCUORIG$ = "Presupuesto " + NUMEDOC$
'        TIPODOC$ = "Anexo Presup."
'        For KKI% = 0 To 3
'          If ANEXPRESU.Check1(KKI%).Value > 0 Then
'            Call PRIDOGRAF(CIXI%, KKI%)
'          End If
'        Next KKI%
'      Else
'        For U& = 1 To ULTREG&("!CARDETPR")
'          X$ = REGLEIDO$("!CARDETPR", U&)
'          CIXI% = CVI(Mid$(X$, 1, 2))
'          DOCUORIG$ = "Presupuesto " + NUMEDOC$
'          TIPODOC$ = "Anexo Presup."
'          For KKI% = 0 To 3
'            If ANEXPRESU.Check1(KKI%).Value > 0 Then
'              Call PRIDOGRAF(CIXI%, KKI%)
'            End If
'          Next KKI%
'        Next U&
'    End If
'    '
'    If CANCELPRINT% < 1 Then
'      If COMALTER%(DOCUORIG$ + " Registrado y Emitido por Impresora\\Repetir Impresión\Continuar") = 1 Then
'        GoTo 10
'      End If
'    End If
End Sub

Private Sub Command17_Click()
   Call Command14_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Cancel = 0 Then
       Call Command2_Click
    End If
    
End Sub

Private Sub itemunico_Click()
    If Label12.Caption <> "Item Unico" Then
        Call Label12_DblClick
        listaderepuestos.Checked = False
        itemunico.Checked = True
    End If
End Sub

Private Sub Label12_Change()
    If Label12 = "Lista Productos" Then
         Check1.Visible = True
       Else
         Check1.Visible = False
    End If
End Sub

Private Sub listaderepuestos_Click()
    If Label12.Caption <> "Lista Productos" Then
        Call Label12_DblClick
        itemunico.Checked = False
        listaderepuestos.Checked = True
    End If

End Sub

Private Sub mnuMnv_Click()
    Call OPLIPRE04.Command3_Click
End Sub

Private Sub mnuvp_Click()
   Call Command15_Click
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
   Call OPLIPRE04.Command1_Click
End Sub

Private Sub Label12_DblClick()
   If TRIM$(UCase$(Label12)) = "ITEM UNICO" Then
       Label12.ForeColor = RGB(200, 0, 0)
       Label12 = "Lista Productos"
       Label12.ToolTipText = "Doble-Click para 'Item Unico'"
       Picture1.Visible = True
     Else
       Label12.ForeColor = RGB(0, 80, 80)
       Label12 = "Item Unico"
       Label12.ToolTipText = "Doble-Click para 'Lista Productos'"
       Picture1.Visible = False
       Check1.Value = 0
   End If
End Sub


Private Sub Command1_Click()
   Call SELECHO("DEUDOR1")
   NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE% > 0 Then
      Text1.TEXT = TRIM$(Str$(NCLIE%))
      Call CARDACLI
      Text12.SetFocus
   End If
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG04/AMACLI"
   Call CONECRUN(ACONEC$, "Padron Maestro de Clientes")
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   OPLIPRE04.Show 1
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   OPPRES04.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   If UCase(Label12) = "LISTA PRODUCTOS" Then
    OPX% = COMALTER%("Formato de Salida:\\Presupuesto\Factura Proforma")
    If OPX% = 0 Then GoTo 99
    If OPX% = 2 Then
       Call FACPROFORMA
       Call BLANFORMU
       Text1.TEXT = ""
       Text1.SetFocus
       GoTo 99
    End If
   End If
   ANEXPRESU.Show 1
   Refresh
   If ANEXPRESU.CONTROLABEL.Caption = "OK" Then
     Call GRAPRESU(OKX%)
     If OKX% > 0 Then
       '
       CANCELPRINT% = 0
       SENDBYMAIL$ = ""
       If ANEXPRESU.Option2.Value = True Then
            NCX1% = XVALO(Text1)
            SENDBYMAIL$ = EMAILCLI$(NCX1%)
         ElseIf ANEXPRESU.Option3.Value = True Then
            CANCELPRINT% = 1
       End If
       '
       Call PRIPRESU(0) ' ES IMPRESION DEFINITIVA
       Call BLANFORMU
       Text1.TEXT = ""
       Text1.SetFocus
     End If
   End If
99 Command14.Enabled = True
End Sub

Sub FACPROFORMA()
   Call GRAPRESU(OKX%)
   If XVALO(Text6) <> 5 Then GoTo 5
  'LLAMAR A FORMULARIO EN CASO DE SER CLIENTE DEL EXTERIOR
   CARFAPRO.Show 1
   If CARFAPRO.OK = False Then GoTo 5
   '
   'Obtener datod del formulario de carga datos proforma
   BULTOS$ = CARFAPRO.Text1
   PNETO$ = Left(CARFAPRO.Text2, 12)
   PTOTAL$ = Left(CARFAPRO.Text3, 12)
   NCMX$ = CARFAPRO.Text4
   CONDVENTA$ = Left(CARFAPRO.Text5, 128)
   '
   CARFAPRO.Text1 = ""
   CARFAPRO.Text2 = ""
   CARFAPRO.Text3 = ""
   CARFAPRO.Text4 = ""
   CARFAPRO.Text5 = ""
   'Actualiza la informacion para el registro
   Call ABREPRESVEN
   '
   SQLX$ = " SELECT * "
   SQLX$ = SQLX$ + " FROM PRESUVEN "
   SQLX$ = SQLX$ + "WHERE NroPres = " & XVALO(Text11)
   '
   Set LISPRETEMP = Ventas.OpenRecordset(SQLX$, dbOpenDynaset)
   With LISPRETEMP
     .Edit
      !CantBultos = BULTOS$
      !PesoNeto = PNETO$
      !PesoTotal = PTOTAL$
      !Ncm = NCMX$
      !CondVta = CONDVENTA$
      .Update
      .Close
   End With
   '
5  'Imprimir Factura Proforma
   Call PRIFACPRO
   '
99
End Sub

Private Sub Command15_Click()
   If XVALO(Text11.TEXT) < 1 Then Exit Sub
   Command15.Enabled = False
   CANCELPRINT% = 1
   Screen.MousePointer = 11
   Call PRIPRESU(1)     ' ES VISTA PREVIA
   '
   TIDOCUM$ = "Presup"
   ' NUDOCU$ = " " + TRIM$(Text11.TEXT) + " -"
   NUDOCU$ = "(VP-" + TRIM$(Text11.TEXT) + ")"
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
   GoTo 99
   '
5  For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
   Next KKI%
   '
   REDOCU$ = REGACT$("PDOCLST")
   URE& = ULTREG&("PDOCSPL")
   LI& = 0: LS& = URE&
6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
   DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
   GoTo 8
   '
7  Screen.MousePointer = 1
   Call MENSERR(24, "Documento no Registrado")
   GoTo 99
   '
8  TPSP$ = ""
   Screen.MousePointer = 11
   While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
   Wend
   '
9  CAPAGINAS% = 0
   For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
   Next KKL&
   '
95 SPOOLSTRING$ = TPSP$
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
10 PPXX% = InStr(TPSP$, Chr$(255))
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
   UI& = UI& + 1
   XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
   If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
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
   EDITFORM.Command1.Enabled = False
   EDITFORM.Show 1
   EDITFORM.Command1.Enabled = True
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
    RECUEC04.Text1.TEXT = Text1.TEXT
    RECUEC04.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    RTA% = ALTERNA%("Listas de Precios de Venta\Condiciones de Pago")
    If RTA% = 1 Then
        Call CIERRARCH("*.*")
        '\\\OT-5-869-FG-28/12/05///
        Call CONECRUN("LISPRE04", "Listas de Precios de Venta")
        '\\\OT-5-869-FG-FIN///
      ElseIf RTA% = 2 Then
        Call CARCONPAGOS
    End If
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Text1.SetFocus
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    If ALPORCOPIA% = 1 Then
      Call MENSERR(24, "Imposible Continuar.Generación por Copia en Proceso.")
      GoTo 99
    End If
    Screen.MousePointer = 11
    Call ABREPRESVEN
    '
    SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM PRESUVEN "
    SQLX$ = SQLX$ + " ORDER BY NroPres DESC "
    '
    Set LISPRETEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    With LISPRETEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        XXX$ = REGBLAN("INDIPRES")
        JJ& = 0: NPXA& = 0
        Do While (.EOF = False)
          NPX& = !NroPres
          NCLIE% = !NroClie
          CIXI% = !CodiInt
          FEX = CVINT(!FECHEMIS)
          '
          CN1$ = ""
          On Error Resume Next
          CR1$ = !Cod_Rubro
          NCR1 = !NumeCRubro
          CA1$ = !Cod_Anio
          On Error GoTo 0
          '
          If TRIM$(CR1$) <> "" And XVALO(NCR1) > 0 Then
            NCR2$ = TRIM$(Str$(NCR1))
            While Len(NCR2$) < 4: NCR2$ = "0" + NCR2$: Wend
            CN1$ = CR1$ + "-" + NCR2$ + "-" + CA1$ + "   "
          End If
          '\\\OT-06-0168-WAR-FIN///
          YY2$ = TRIM$(AJUSTI$(CN1$ + TRIM$(RASOCLI$(NCLIE%)), 30)) + " - "
          If NPX& <> NPXA& Then
               YY2$ = YY2$ + TRIM$(CODEXT$(CIXI%))
             Else
               YY2$ = YY2$ + "VARIOS"
          End If
          YY2$ = YY2$ + " - " + FECHATEX$(FEX)
          '
          Call REPLA(XXX$, MKS$(NPX&), 1, 4)
          Call REPLA(XXX$, YY2$, 5, 60)
          If NPX& <> NPXA& Then
            JJ& = JJ& + 1
          End If
          Call GRAREG("INDIPRES", XXX$, JJ&)
          '
          NPXA& = NPX&
        .MoveNext
        Loop
      End If
    End With
    '
    Call SETULTREG("INDIPRES", JJ&)
    Call CIERRARCH("INDIPRES")
    Call FRESHECHO("INDIPRES")
    '
3   Screen.MousePointer = 1
    '\\\OT-05-0758-WAR-18/11/05///
4    OBJ1$ = OBJPLA$("SELECLIPRE", VDEF$)
    If TRIM$(OBJ1$) <> "" Then
      OBJANT$ = OBJ1$
      NPX& = CVS(Mid$(OBJ1$, 1, 4))
      NCLIE% = CVI(Mid$(OBJ1$, 5, 2))
      'VALIDACIONES
      If NPX& <= 0 Then
        Call COMUNI("Numero de Presupuesto no Válido")
        VDEF$ = OBJANT$
        GoTo 4
      End If
      If NCLIE% <= 0 Then
        Call COMUNI("Numero de Cliente no Válido o Inexistente")
        VDEF$ = OBJANT$
        GoTo 4
      End If
      '
      Call PREPORCOP(NPX&, NCLIE%)
    End If
    '
99  Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CIERRARCH("*.*")
    ACONEC$ = "ARCHIG04/AMASTO"
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
'   On Error Resume Next
'   TEXT21(Index).SetFocus
'   On Error GoTo 0
End Sub

Private Sub Form_Load()
    '
    '
    Me.Visible = False
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Frame7.BackColor = &HFCD7B6
      Picture6.Visible = True
    End If
    
    Call CENTRAFORM(Me)
    listaderepuestos.Checked = False

    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If EXISTE%(LUDAT$ + "PEDMEDID.RFS") < 1 Then   ' PARA OCULTARLO
       Check1.Value = Unchecked
       Check1.Top = 2000
       Label12.Top = 500
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
     Me.Visible = True
     Refresh
     Show
     
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
     OPENFORMS = DoEvents
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
     NCXI% = NUMECLI%(CPROX$)
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

Private Sub Label8_Click()
   Text13.SetFocus
End Sub

Private Sub LisPreVen_Click()
   Call CIERRARCH("*.*")
   Call CONECRUN("LISPRE04", "Listas de Precios de Venta")
End Sub

Private Sub LisRang_Click()
   Call OPLIPRE04.Command6_Click
End Sub

Private Sub List1_DBLClick()
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

Private Sub nuepre_Click()
   Call Command1_Click
End Sub

Private Sub PORDESCU_Change()
   DESC$ = PORDESCU.TEXT
   TTT1 = XVALO(Total(0).TEXT) * COEFDESCU(DESC$) / 100
   Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   TTT3 = XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT)
   Total(3).TEXT = CSTRING$(MKS$(TTT3), 3, 11, 2, 2)
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

Private Sub Text1_Change()
   Call BLANFORMU
End Sub

Private Sub Text1_DblClick()
  Call SELECHO("DEUDOR1")
  NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
  If NCLIE% > 0 Then
    Text1.TEXT = TRIM$(Str$(NCLIE%))
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
   For KK% = 1 To Len(Text11.TEXT)
     CCX% = Asc(Mid$(Text11.TEXT, KK%, 1))
     If CCX% >= 48 Then
       If CCX% <= 57 Then
         NUMEPRE$ = NUMEPRE$ + Chr$(CCX%)
       End If
     End If
   Next KK%
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

Private Sub Text13_DblClick()
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

Private Sub Text13_GOTFOCUS()
   If VALICLIE% < 1 Then Text1.SetFocus: Exit Sub
End Sub

Private Sub Text16_Change()
   Call NUCORUB
End Sub

Private Sub Text16_LostFocus()
  Text16 = UCase(Text16)
End Sub

Private Sub Text2_gotfocus()
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
     If XVALO(Text1.TEXT) <= 32767 Then
       NCLIE% = XVALO(Text1.TEXT)
       If REGICLI&(NCLIE%) > 0 Then
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
   Text6.TEXT = ""
   Text7.TEXT = ""
   Text8.TEXT = ""
   Text10.TEXT = ""
   Text12.TEXT = ""
   Text13.TEXT = ""
   Text28.TEXT = ""
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
   For KK% = 0 To 6
      Total(KK%).TEXT = ""
   Next KK%
   '
   List1.Clear
   Call SETULTREG("!CARDETPR", 0)
   If EXISTE%(LUDAT$ + "PEDMEDID.RFS") > 0 Then
      Call SETULTREG("!PEDMEDID", 0)
   End If
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
    NC% = XVALO(Text1.TEXT)
    If NC% > 0 Then
      '
      '\\\OT-06-0607-OD-28/07/08///
      'EVALUA SI LA CUENTA NO ESTA SUSPENDIDA
      '
      If LIMICRE(NC%) < (-0.5) Then
          Call MENSERR(24, "Cuenta Suspendida.")
          Text1 = ""
          Call BLANFORMU
          GoTo 99
      End If
      '\\\OT-06-0607-OD-FIN///
      '
      If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
        Text2.TEXT = RASOCLI$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        Text6.TEXT = PIVACLI%(NC%)
        Text7.TEXT = POSIVAC$(NC%)
        Text8.TEXT = TEL1CLI$(NC%)
        Label26 = CODICLI$(NC%)
        '
        TEXT21(0).TEXT = Str$(LIPRCLI%(NC%))
        TEXT21(1).TEXT = Str$(CPAGCLI%(NC%))
        TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
        '
        If XVALO(TEXT21(2).TEXT) < 1 Then
           Eco21(2).TEXT = "GERENCIA"
        End If
        '
        PORDESCU.TEXT = DESCUCLI$(NC%)
        '
      End If
    End If
99  Screen.MousePointer = 1
    '
End Sub
   '
Function NUPRESUP&()
   '
   NPP& = 1
   Call ABREPRESVEN
   '
   SQLX$ = "SELECT NroPres FROM PresuVen "
   SQLX$ = SQLX$ + "ORDER BY NroPres DESC;"
   Set PRESUTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   On Error Resume Next
   With PRESUTEMP
     .MoveFirst
     If Err < 1 Then
       NPP& = 1 + !NroPres
     End If
   End With
   On Error GoTo 0
   '
   NUPRESUP& = NPP&
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
   Set PRESUTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
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
    FORIPRE$ = TRIM$(CONTROL$("", "FORPRESU"))
    FPXX$ = "FORPRESU"
    If Picture1.Visible = True Then
       FORIPRE$ = TRIM$(CONTROL$("", "FORPRERE"))
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
10  NC% = XVALO(Text1.TEXT)
    TIPODOC$ = "Presupuesto"
    DESTIDOC% = NC%
    NUMEDOC$ = TRIM$(Text11.TEXT)
      If VIPREVIA% = 1 Then NUMEDOC$ = "(VP-" + NUMEDOC$ + ")"
    FECHDOC$ = TRIM$(Text9.TEXT)
    '
    CODPARAM$(1) = "REF-MAT3": DOCPARAM$(1) = TRIM$(Text10.TEXT)
    CODPARAM$(2) = "REF-MAT2": DOCPARAM$(2) = TRIM$(Text12.TEXT)
    CODPARAM$(3) = "PROV-CLI": DOCPARAM$(3) = NPROCLI$(NC%)
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
    CODPARAM$(21) = "MEDIDA-1": DOCPARAM$(21) = EMAILCLI(NC%) ' E-MAIL
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
    If XVALO(Text14) >= 0.5 Then
      DOCPARAM$(28) = FORMATNUM$(XVALO(Total(3)) / XVALO(Text14), "F12.2")
    End If
    '
    CAPARDOC% = 28
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
      '
      CODTABU1$(1) = "CANTIDAD"
      CODTABU1$(2) = "COD-MAT"
      CODTABU1$(3) = "DES-MAT"
      CODTABU1$(4) = "PRE-LIST"
      CODTABU1$(5) = "NET-ITEM"
      CODTABU1$(6) = "N-PLANO"
      CODTABU1$(7) = "R-PLANO"
      CODTABU1$(8) = "F-PLANO"
      CACOLTAB1% = 8
      '
      CAITAB1% = 0
      For U& = 1 To ULTREG&("!CARDETPR")
        X$ = REGLEIDO$("!CARDETPR", U&)
        CIIX% = CVI(Mid$(X$, 1, 2))
        CAN = CVS(Mid$(X$, 47, 4))
        INETO = CVS(Mid$(X$, 39, 4)) * (100 - CVS(Mid$(X$, 69, 4))) / 100
        PRELIX = PRELISTA#(ARCHIPRE$, CIIX%)
        DESCULISTA# = CVS(Mid$(X$, 85, 4)) 'Descuento de lista
        ' Suma de los descuentos de grupo cliente/linea + Articulo/Cliente
        DescFamGruArtClien# = CVS(Mid$(X$, 89, 4))
        If CAN > 0.05 Then
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = Str$(CAN)
          TETABU1$(2, CAITAB1%) = CODEXT$(CIIX%)
          TETABU1$(3, CAITAB1%) = DESCRIT0$(CIIX%)
          TETABU1$(4, CAITAB1%) = FORMATNUM$(INETO, "F12.4")
          TETABU1$(5, CAITAB1%) = FORMATNUM$(INETO * CAN, "F12.2")
          TETABU1$(6, CAITAB1%) = FORMATNUM$(PRELIX, "F12.2")
          TETABU1$(7, CAITAB1%) = FORMATNUM$(DESCULISTA#, "F4.1")
          TETABU1$(8, CAITAB1%) = FORMATNUM$(DescFamGruArtClien#, "F4.1")
          COMPDESCRI$ = ""
          If CONTROL$("PEDIDO", "VARCOLOR") = "SI" Then
            COLORBA$ = TRIM$(Mid$(X$, 129, 12))
            TELA$ = TRIM$(Mid$(X$, 141, 12))
            COLTELA$ = TRIM$(Mid$(X$, 153, 12))
            If COLORBA$ <> "" Or TELA$ <> "" Or COLTELA$ <> "" Then
              COMPDESCRI$ = "CB: " + TRIM$(COLORBA$) + " - T: " + TRIM$(TELA$) + " - CT: " + TRIM$(COLTELA$)
              CAITAB1% = CAITAB1% + 1
              TETABU1$(3, CAITAB1%) = COMPDESCRI$
            End If
          End If
        End If
      Next U&
    End If
    '
    NOCONFIRM% = 1
    Call PRINTDOC(FPXX$)
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
     With PresuVen
       .FindFirst SQLX$
       If .NoMatch = False Then
          Call CIERRARCH("PRESUPU")
         Call COMUNI("Código Identificador Ocupado")
         Call NUCORUB
         Exit Sub
       End If
     End With
   End If
   '\\\OT-06-0168-WAR-FIN///
   '
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLIE% = XVALO(Text1.TEXT)
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
   With PresuVen
     If Picture1.Visible = False Then
        '\\\OT-06-0168-WAR-09/05/06///
        CANTI0 = XVALO(Text14)
        If CANTI0 < 1 Then CANTI0 = 1
        '\\\OT-06-0168-WAR-FIN///
       .AddNew
        !NroPres = XVALO(Text11.TEXT)
        '\\\OT-06-0168-WAR-09/05/06///
         !NumeCRubro = XVALO(Text15.TEXT)
         !Cod_Rubro = Text16.TEXT
         !Cod_Anio = Label16
        '\\\OT-06-0168-WAR-FIN///
        !FECHEMIS = CVDAT(FEMI%)
        !NroClie = NCLIE%
        !RasoClie = Left$(RASOCLI$(NCLIE%), 64)
        !Referen = Left$(Text10, 16)
        !AteSen = Left$(Text12, 48)
        !LisPre = XVALO(TEXT21(0))
        !ConPag = XVALO(TEXT21(1))
        !VENDED = XVALO(TEXT21(2))
        !MonedEmi = MONEMI%
        !NuOrdIt = NUOIT%
        !CodiInt = CIIX%
        !CodiExt = Left$(CODEXT$(CIIX%), 16)
        !Descrip = Left$(DESCRIT0$(CIIX%), 64)
        !Especif = Text28
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
        !PreUnit = XVALO(Total(3)) / CANTI0  'XVALO(Total(3))'\\\OT-06-0168-WAR-09/05/06///
        !Cantit = CANTI0 '1'\\\OT-06-0168-WAR-09/05/06///
        !ImToNe = CANTI0 * XVALO(Total(3)) '1 * XVALO(Total(3))'\\\OT-06-0168-WAR-09/05/06///
        !FechVent = CVDAT(0)
        !NroPedi = 0
        !NroFac_Cor = ""
        !NroFac_Lar = ""
       .Update
      Else
       For U& = 1 To ULTREG&("!CARDETPR")
         X$ = REGLEIDO$("!CARDETPR", U&)
         CIIX% = CVI(Mid$(X$, 1, 2))
         CAN = CVS(Mid$(X$, 47, 4))
         INETO = CVS(Mid$(X$, 39, 4)) * (100 - CVS(Mid$(X$, 69, 4))) / 100
         CLRBASE$ = Mid$(X$, 129, 12)
         TTTLA$ = Mid$(X$, 141, 12)
         CLRTELA$ = Mid$(X$, 153, 12)
         
         If CAN > 0.05 Then
           .AddNew
             !NroPres = XVALO(Text11.TEXT)
             '\\\OT-06-0168-WAR-09/05/06///
              !NumeCRubro = XVALO(Text15.TEXT)
              !Cod_Rubro = Text16.TEXT
              !Cod_Anio = Label16
             '\\\OT-06-0168-WAR-FIN///
             !FECHEMIS = CVDAT(FEMI%)
             !NroClie = NCLIE%
             !RasoClie = Left$(RASOCLI$(NCLIE%), 64)
             !Referen = Left$(Text10, 16)
             !AteSen = Left$(Text12, 48)
             !LisPre = XVALO(TEXT21(0))
             !ConPag = XVALO(TEXT21(1))
             !VENDED = XVALO(TEXT21(2))
             !MonedEmi = MONEMI%
             !NuOrdIt = NUOIT%
               NUOIT% = NUOIT% + 1
             !CodiInt = CIIX%
             !CodiExt = Left$(CODEXT$(CIIX%), 16)
             !Descrip = Left$(DESCRIT0$(CIIX%), 64)
             !Especif = DESCRIT0$(CIIX%)
             !PreList = INETO
             !TDESCUEN = Left$(PORDESCU, 32)
                DESC$ = PORDESCU
                TTT1 = COEFDESCU(DESC$)
             !PorcDesc = TTT1
             !ImpoDesc = INETO * TTT1 / 100
             !PreUnit = INETO * (1 - TTT1 / 100)
             !Cantit = CAN
             !ImToNe = CAN * INETO * (1 - TTT1 / 100)
             !FechVent = CVDAT(0)
             !NroPedi = 0
             !NroFac_Cor = ""
             !NroFac_Lar = ""
             !CBASE = TRIM$(Mid$(CLRBASE$, 1, 12))
             !TTALLE = TRIM$(Mid$(TTTLA$, 1, 8))
             !TCOLOR = TRIM$(Mid$(CLRTELA$, 1, 16))
            .Update
         End If
       Next U&
     End If
   End With
   '
   Call CIERRARCH("PRESUPU")
   '
   CR1$ = TRIM$(Text16)
   NCR1 = XVALO(Text15)
   CA1$ = TRIM$(Label16)
   '
   If TRIM$(CR1$) <> "" And XVALO(NCR1) > 0 Then
     NCR2$ = TRIM$(Str$(NCR1))
     While Len(NCR2$) < 4: NCR2$ = "0" + NCR2$: Wend
     CN1$ = CR1$ + "-" + NCR2$ + "-" + CA1$ + "   "
   End If
   YY2$ = TRIM$(AJUSTI$(CN1$ + TRIM$(RASOCLI$(NCLIE%)), 30)) + " - "
   If Picture1.Visible = False Then
       YY2$ = YY2$ + TRIM$(CODEXT$(CIIX%))
     Else
       YY2$ = YY2$ + "VARIOS"
   End If
   FEX = FEMI%
   YY2$ = YY2$ + " - " + FECHATEX$(FEX)
   Call REGAPP("INDIPRES", MKS$(XVALO(Text11)) + YY2$)
   Call CIERRARCH("INDIPRES")
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
   If Check1.Value = 1 Then ' si la opcion de pedidos x medida esta en true
      Call CargaPorMedida(OKEY2%)
      If OKEY2% < 1 Then Exit Sub
      'Call ActualizaTotales
      'Exit Sub
   End If

   HOII% = HOY(HOS$)
   CLIEN% = XVALO(Text1.TEXT)
   NULISTA% = XVALO(TEXT21(0))

   '
   ATRI$ = "/NC/ARCHECO(1)=" + ARCHIPRE$
   ATRI$ = ATRI$ + "/TITUPAN=Detalle Presupuesto (Productos)"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top + 300))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
   '
   Timer1.Enabled = True
10 PRF$ = "CARPRESR"
   If CONTROL("PEDIDO", "VARCOLOR") = "SI" Then
     PRF$ = "CARPRESD"
   End If

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
     If PRECARGA < 0.005 Or PRECARGA > 99999.99 Then
       If modo% < 2 Then
         If ECOPRECIO% >= 0 Then
           Call MENSERR(24, "Artículo sin Precio")
           Call REPLA(X$, MKS$(0), 39, 4)
           Call GRAREG("!CARDETPR", X$, U&)
           modo% = 2
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
     '
     ' Trae la combinacion del descuento que tiene el articulo/Cliente
     DESCARCLI# = ROUND(DESCUCLIART(CLIEN%, CI%), 1)
       KOEFIARCLI# = 1 - DESCARCLI# / 100 'Calcula el coeficiente
     'Trae el descuento de lista
     DESCULISTA# = ROUND(DESCULISPRE(NULISTA%), 1)
       KOEFILISTA# = 1 - DESCULISTA# / 100 ' Calcula el coeficiente de Dto. de lista
     'trae el descuento combinado de la linea y grupo de clientes
     DESGRUCLIENTE# = ROUND(DESCUGRULIN(CLIEN%, CI%), 1)
       KOEFIGRUCLI# = 1 - DESGRUCLIENTE# / 100
     'Agrupa en un solo valor el descuento de la Linea-Grupo de Cliente y Articulo/Cliente
     DescFamGruArtClien# = ROUND#(100 * (1 - KOEFIGRUCLI# * KOEFIARCLI#), 1)
     'Agrupa todos los descuentos en un solo valor
     DESCGEN# = ROUND#(100 * (1 - KOEFILISTA# * KOEFIGRUCLI# * KOEFIARCLI#), 1)
     Call REPLA(X$, MKS$(DESCGEN#), 69, 4)
     Call REPLA(X$, MKS$(DESCULISTA#), 85, 4)
     Call REPLA(X$, MKS$(DescFamGruArtClien#), 89, 4)
     '
     Call GRAREG("!CARDETPR", X$, U&)
     '
   Next U&
   '
   Call CARLISPRES
   Call CIERRARCH("!CARDETPR")
   Call CALTOPRES
   '
   List1.Refresh
   On Error Resume Next
   Command14.SetFocus
   On Error GoTo 0
   Screen.MousePointer = 1
   '
End Sub

Sub CARLISPRES()
   '
   List1.Clear
   FIN& = ULTREG&("!CARDETPR")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDETPR", U&)
     TTXX$ = Space$(96)
     CI% = CVI(Left$(X$, 2))
     FENTRESOL% = CVI(Mid$(X$, 63, 2))
     CAN = CVS(Mid$(X$, 47, 4))
     PORDESCUIN = CVS(Mid$(X$, 69, 4))
     PRECARGA = CVS(Mid$(X$, 39, 4))
     PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
     TOT# = CDbl(CAN * PREUNI)
     '
     Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 1, 6)
     Call REPLA(TTXX$, CODEXT$(CI%), 8, 15)
     FEX = FENTRESOL%
     Call REPLA(TTXX$, DESCRIT0$(CI%), 24, 28)
     PRELIX = PRELISTA#(ARCHIPRE$, CI%)
     If PRELIX < PRECARGA Then PRELIX = PRECARGA
     If PRELIX < PREUNI Then PRELIX = PREUNI
     PORDESCUIN = 0: If PRELIX > 0.005 Then PORDESCUIN = 100 * (PRELIX - PREUNI) / PRELIX
     Call REPLA(TTXX$, CSTRING$(MKS$(PRELIX), 3, 8, 2, 2), 52, 8)
     Call REPLA(TTXX$, CSTRING$(MKS$(PORDESCUIN), 3, 5, 1, 2), 61, 5)
     Call REPLA(TTXX$, CSTRING$(MKS$(PREUNI), 3, 8, 2, 2), 66, 8)
     Call REPLA(TTXX$, CSTRING$(MKD$(TOT#), 3, 10, 2, 2), 76, 10)
     List1.AddItem TTXX$
     '
   Next U&
   '
End Sub
'
Sub CALTOPRES()
   '
   Static CLIA$, RNI%, APLIVA%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE% = XVALO(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIE% > 0 Then
       PVCLI% = PIVACLI%(NCLIE%)
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
       PREUNI = PRELI * (100 - PORDESCUIN) / 100
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
   TTT1 = XVALO(Total(0).TEXT) * COEFDESCU(DESC$) / 100
   Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   '
   TTT3 = XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT)
End Sub

Sub CALTOIVA()
   '
   Static CLIA$, RNI%, APLIVA%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE% = XVALO(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIE% > 0 Then
       PVCLI% = PIVACLI%(NCLIE%)
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
       IMPOIVA# = APLIVA% * XVALO(IMPIVA21) + ROUND#(TOT# * PORCEIVA(CIXXI%) / 100, 2)
       IMPIVA21 = TRIM$(FORMATNUM$(IMPOIVA#, "F12.2"))
     Else
       IMPOIVA# = APLIVA% * XVALO(IMPIVA105) + ROUND#(TOT# * PORCEIVA(CIXXI%) / 100, 2)
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
            IMPOIVA# = APLIVA% * XVALO(IMPIVA21) + ROUND#(TOT# * (PORCEIVA(CIXXI%) / 100) * (1 - COEFDESCU(DESC$) / 100), 2)
            IMPIVA21 = TRIM$(FORMATNUM$(IMPOIVA#, "F12.2"))
          Else
            IMPOIVA# = APLIVA% * XVALO(IMPIVA105) + ROUND#(TOT# * (PORCEIVA(CIXXI%) / 100) * (1 - COEFDESCU(DESC$) / 100), 2)
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
'   '
'   Static CTTXX%, TASAA, TASAB, TASAEXP, MINIPER
'   '
'   If CTTXX% < 1 Then
'     TASAA = XVALO(CONTROL$("PERIBRU", "TASA-A"))
'     TASAB = XVALO(CONTROL$("PERIBRU", "TASA-B"))
'     TASAEXP = XVALO(CONTROL$("PERIBRU", "TASA-EXP"))
'     MINIPER = XVALO(CONTROL$("PERIBRU", "MINIPER"))
'     CTTXX% = 1
'   End If
'   '
'   TASAX = TASAA
'   NCLIE% = XVALO(Text1)
'   If NCLIE% < 1 Then Exit Sub
'   '
'   VARIDC% = 1: PVCLX% = PIVACLI%(NCLIE%)
'   If PVCLX% = 0 Or PVCLX% = 4 Then VARIDC% = 2
'   If PVCLX% = 5 Or PVCLX% = 7 Then VARIDC% = 3
'   '
'   If VARIDC% = 2 Then TASAX = TASAB
'   If VARIDC% = 3 Then TASAX = TASAEXP
'   '
'   'Text15 = ""'\\\OT-06-0168-WAR-09/05/06///
'   REPERCI$ = FILTERGETRECORD$("DAPIBCLI", 1, MKI$(NCLIE%))
'   CATIB1% = Asc(Mid$(REPERCI$, 3, 1))
'   PORCONVE = CVS(Mid$(REPERCI$, 5, 4))
'   If PORCONVE > 100 Then PORCONVE = 100
'   '
'   If CATIB1% > 0 Then         ' NO ESTA EXENTO
'     If PORCONVE >= 10.1 Then  ' PARTICIPACION CONVENIO MULTILATERAL
'       BIMPON# = XVALO(Total(3))
'       If VARIDC% = 2 Then BIMPON# = BIMPON# + XVALO(TOTIVA)
'       If BIMPON# * TICAMONE(MONEMI%) >= MINIPER Then
'         PERSALTA# = (BIMPON# * PORCONVE / 100) * (TASAX / 100)
'         TOTPER = FORMATNUM$(PERSALTA#, "F11.2")
'       End If
'     End If
'   End If
   '
         '
   NCLIE% = XVALO(Text1)
   TASAX = ALIPERBRU#(NCLIE%, MINIPER#, VARIDOC%)
   '
   Text15 = ""
   BIMPON# = XVALO(Total(3))
   If BIMPON# * TICAMONE(MONEMI%) >= MINIPER# Then
      PERSALTA# = BIMPON# * (TASAX / 100)
      TOTPER = FORMATNUM$(PERSALTA#, "F11.2")
      'Text15 = FORMATNUM$(PERSALTA#, "F11.2")
   End If
   '
End Sub
'
Sub PREPORCOP(NPRESU&, NCLIE%)
   '
   NUCLI% = NCLIE% 'CVI(Mid$(XX$, 7, 2))'CLIENTE QUE SE ELIGIO
   Call ABREPRESVEN
   '
   SQLX$ = " SELECT * "
   SQLX$ = SQLX$ + " FROM presuven "
   SQLX$ = SQLX$ + " WHERE NroPres = " & NPRESU& & " "
   SQLX$ = SQLX$ + " ORDER BY NuOrdIt "
   '
   Set PRESUTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
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
        CODI1% = !CodiInt
        PORDESCU1 = !PorcDesc
        PREUNI1 = !PreList
        ESPE$ = !Especif
        CANTI0 = !Cantit '\\\OT-06-0168-WAR-09/05/06///
        'CARGA DE ENCABEZADO DE PRESUPUESTO
          Text1 = TRIM$(Str$(NUCLI%))
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
            Label12 = "LISTA PRODUCTOS"
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
          Else
            'LISTA DE REPUESTOS
            Label12 = "ITEM UNICO"
            Call Label12_DblClick
            '
            XX$ = REGBLAN$("CARDETPR")
            JJ& = 0
            Call BLANARCH("!CARDETPR")
            .MoveFirst
            Do While Not .EOF
              CODI1% = !CodiInt
              PRELI1 = !PreList
              IMPOUN = !PreUnit
              CAN1 = !Cantit
              COLORBA$ = SAFETEXT$(TRIM$(!CBASE))
              TELA$ = SAFETEXT$(TRIM$(!TTALLE))
              COLTELA$ = SAFETEXT$(TRIM$(!TCOLOR))
              '
              Call REPLA(XX$, MKI$(CODI1%), 1, 2)
              Call REPLA(XX$, MKS$(PRELI1), 39, 4)
              Call REPLA(XX$, MKS$(IMPOUN), 65, 4)
              Call REPLA(XX$, MKS$(CAN1), 47, 4)
              Call REPLA(XX$, COLORBA$, 129, 12)
              Call REPLA(XX$, TELA$, 141, 12)
              Call REPLA(XX$, COLTELA$, 153, 12)

              JJ& = JJ& + 1
              Call GRAREG("!CARDETPR", XX$, JJ&)
            .MoveNext
            Loop
            Call SETULTREG("!CARDETPR", JJ&)
            Call CARDETPRES
        End If
      End If
   End With
   '
   ALPORCOPIA% = 1 'MARCA PARA SABER SI ESTA EN ALTA POR COPIA
   '
End Sub

Sub CargaPorMedida(OKEY2%) '\\OT-07-0420
    '
    OKEY2% = 0
    HOII% = HOY(HOS$)
    '******CONCATENO EL NOMBRE DE LA LISTA CON EL NUMERO DE LISTA QUE ES DE TRES POSICIONES********
     'EL while si el text21 es de
    '***********************************************************************************************
    'VTB% = VENTABU%("VENCICAI/NC/TITUPAN=Tabla de Vencimientos de CAI")
    ATRI$ = "PEDMEDID/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Carga de Presupuesto por Medida"""
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top + 600))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
    
    'VTB% = VENTABU("PEDMEDID/NC/TITUPAN=Carga de Presupuesto por Medida")  'LA VENTANA QUE SE ABRE con los datos y al confirmala seguimos aca abajo
10  VTB% = VENTABU(ATRI$)  'LA VENTANA QUE SE ABRE con los datos y al confirmala seguimos aca abajo
    If VTB% >= 0 Then
        FIN& = ULTREG&("!PEDMEDID") 'NUMERO DE ULTIMO REGISTRO
        List1.Clear
         
        ' VALIDACIONES
        For U& = 1 To FIN& ' dentro de este form se van leyendo los campos que se ingresaron en la ventana
           Y$ = REGLEIDO$("!PEDMEDID", U&) 'NOS DEVUELVE EL REGISTRO LEIDO EN Y$
           CODFAM$ = TRIM$(Mid$(Y$, 1, 2)) 'MID$ le pasamos como arg.en y$ el registro en 1 donde comienza y en 2 los bytes de acuerdo a la estructura
           M1% = CVI(Mid$(Y$, 3, 2)) ' aqui con cvi paara que devuel. un numero decimal
           M2% = CVI(Mid$(Y$, 5, 2))
           M3% = CVI(Mid$(Y$, 7, 2))
           VARIANTE$ = TRIM$(Mid$(Y$, 9, 4))
           CAN = CVS(Mid$(Y$, 13, 4))
           FECHA% = CVI(Mid$(Y$, 17, 2)) 'CVI pasaba la fecha a un numero
           '
           If CODFAM$ = "" Then
                 Call MENSERR(24, "Falta Código de Familia")
                 GoTo 10
              ElseIf ECOARCH$("TAFAMFIL", CODFAM$) = "" Then
                 Call MENSERR(24, "Código de Familia no Válido Según Tabla")
                 GoTo 10
           End If
           '
           If CDbl(M1) * CDbl(M2) < 1 Then
               Call MENSERR(24, "Faltan Medidas Básicas")
               GoTo 10
           End If
           '
           If CAN < 0.5 Then
               Call MENSERR(24, "Falta Cantidad")
               GoTo 10
           End If
           '
           CODIGO$ = CODFAM$ & M1 & M2   ' concateno el nuevo codigo
           If M3 > 0 Then CODIGO$ = CODIGO & M3   ' concateno el nuevo codigo
           If VARIANTE$ <> "" Then CODIGO$ = CODIGO$ + "-" + VARIANTE$
           If Len(CODIGO$) > 15 Then
              Call MENSERR(24, "Imposible Registrar Código '" + CODIGO$ + "'")
              GoTo 10
           End If
        Next U&
        '
        '
        For U& = 1 To FIN&
           Y$ = REGLEIDO$("!PEDMEDID", U&) 'NOS DEVUELVE EL REGISTRO LEIDO EN Y$
           CODFAM$ = TRIM$(Mid$(Y$, 1, 2)) 'MID$ le pasamos como arg.en y$ el registro en 1 donde comienza y en 2 los bytes de acuerdo a la estructura
           M1% = CVI(Mid$(Y$, 3, 2)) ' aqui con cvi paara que devuel. un numero decimal
           M2% = CVI(Mid$(Y$, 5, 2))
           M3% = CVI(Mid$(Y$, 7, 2))
           VARIANTE$ = TRIM$(Mid$(Y$, 9, 4))
           CAN = CVS(Mid$(Y$, 13, 4))
           FECHA% = CVI(Mid$(Y$, 17, 2)) 'CVI pasaba la fecha a un numero
           '
           CODIGO$ = CODFAM$ & M1 & M2   ' concateno el nuevo codigo
           If M3 > 0 Then CODIGO$ = CODIGO & M3   ' concateno el nuevo codigo
           If VARIANTE$ <> "" Then CODIGO$ = CODIGO$ + "-" + VARIANTE$
           '
30         If CODINT%(CODIGO$) < 1 Then
              OKEY% = 0
              Call CREACODIGO(CODFAM$, M1%, M2%, M3%, VARIANTE$, OKEY%)
              If OKEY% > 0 Then GoTo 30
              Call FINAL("")
           End If
           '
           CII% = CODINT%(CODIGO$)
           Call LEEEXPLO(CODIGO$, 1)
           If CAIT% < 1 Then
              Call CREAFORMULA(CII%, CODFAM$, M1%, M2%, M3%, VARIANTE$)
           End If
           '
           NLISTA% = XVALO(TEXT21(0).TEXT)
           If NLISTA% > 0 Then
              If PRELISTA#(ARCHIPRE$, CII%) < 0.005 Then
                 ' CALCULA EL PRECIO DE VENTA
                 PREUNI# = PRECALCULADO(CODFAM$, M1, M2)
                 TTXX$ = REGBLAN$("LIPREVEN") 'EL ESPACIO DESTINADO A LA VARIABLE
                 'ARMAMOS EL REGISTRO
                 Call REPLA(TTXX$, MKI$(CII%), 1, 2)
                 Call REPLA(TTXX$, DENOMI$, 3, 42)
                 Call REPLA(TTXX$, MKD$(PREUNI#), 62, 8)
                 Call FILTERPUTRECORD(ARCHIPRE$, 1, MKI$(CII%), TTXX$)
                 Call CIERRARCH(ARCHIPRE$)
                 PREUNI2 = PRELISTA(ARCHIPRE$ + "/NEW", CII%)
              End If
           End If
           '
           YY$ = REGBLAN$("CARDETPR")
           Call REPLA(YY$, MKI$(CII%), 1, 2)
           PREUNI = PRELISTA(ARCHIPRE$ + "/NEW", CII%)
'           If PREUNI < 0.005 Then
'              PREUNI = PRECALCULADO(CODFAM$, M1, M2)
'              TTXX$ = REGBLAN$("LIPREVEN") 'EL ESPACIO DESTINADO A LA VARIABLE
'              Call REPLA(TTXX$, MKI$(CII%), 1, 2)
'              Call REPLA(TTXX$, DESCRIT0$(CII%), 3, 42)
'              Call REPLA(TTXX$, MKD$(CDbl(PREUNI)), 62, 8)
'              Call REGAPP(ARCHIPRE$, TTXX$) 'aqui regapp lo agrega en la ultima posicion por que ya sabemos que no existe
'              Call CIERRARCH(ARCHIPRE$)
'              PREUNI2 = PRELISTA(ARCHIPRE$ + "/NEW", CII%)
'           End If
           Call REPLA(YY$, MKS$(PREUNI), 39, 4)
           Call REPLA(YY$, MKS$(CAN), 47, 4)
           Call GRAREG("!CARDETPR", YY$, U&)
        Next U&
        Call SETULTREG("!CARDETPR", FIN&)
        OKEY2% = 1
        '
    End If
    Call CIERRARCH("PEDMEDID")
    '\\OT-07-0420
End Sub

Sub CREACODIGO(CODFAM$, M1%, M2%, M3%, VARIANTE$, OKEY%)
    '
    OKEY% = 0
    '
    ' ARMA EL NUEVO CODIGO
    CODIGO$ = CODFAM$ & M1 & M2   ' concateno el nuevo codigo
    If M3 > 0 Then CODIGO$ = CODIGO & M3   ' concateno el nuevo codigo
    If VARIANTE$ <> "" Then CODIGO$ = CODIGO$ + "-" + VARIANTE$
    '
    ' ARMA LA NUEVA DESCRIPCION
    DENOMI$ = TRIM(ECOARCH$("TAFAMFIL", CODFAM$)) 'dev.el segundo campo pasandole el primero como argumento
    DENOMI$ = DENOMI$ & " " & M1 & " x " & M2
    If M3 > 0 Then DENOMI$ = DENOMI$ & " x " & M3
    DENOMI$ = DENOMI$ & " mm." 'concateno la descripcion
    If VARIANTE$ <> "" Then DENOMI$ = DENOMI$ + " " & VARIANTE$
    '
    ' AGREGA EL CODIGO AL MAESTRO DE ARTICULOS
    Call AGRECOD(CODIGO$, DENOMI$, 5, "U.", OKX%) 'se graban los datos
    If OKX% < 1 Then     ' FALLO LA CREACIÓN
       Exit Sub
    End If
    CII% = CODINT%(CODIGO$)
    '
    OKEY% = 1
    '
End Sub

Function PRECALCULADO(CODFAM$, M1, M2)
    RECORD$ = FILTERGETRECORD$("TAFAMFIL", 1, CODFAM$) 'lee en la tabla tafamfil en la posicion 1 con el argumento codigo de familia
    PRECITAFAMFIL = CVD(Mid$(RECORD$, 33, 8)) 'devuelve el precio que viene de la tabla tafamfil
    PRECIFIJO = CVD(Mid$(RECORD$, 47, 8))
    MEDIDA = (M1 / 1000) * (M2 / 1000) 'calculamos la medida en M2 (dividimos * 1000 por que son mm.
    PRECALCULADO = PRECIFIJO + PRECITAFAMFIL * MEDIDA 'precio unitario surge del precio que trajimos de la tabla tafamfil por los M2
' *** CODIGO ANTERIOR
'    RECORD$ = FILTERGETRECORD$("TAFAMFIL", 1, CODFAM$) 'lee en la tabla tafamfil en la posicion 1 con el argumento codigo de familia
'    PRECITAFAMFIL = CVD(Mid$(RECORD$, 33, 8)) 'devuelve el precio que viene de la tabla tafamfil
'    MEDIDA = (M1 / 1000) * (M2 / 1000) 'calculamos la medida en M2 (dividimos * 1000 por que son mm.
'    PRECALCULADO = PRECITAFAMFIL * MEDIDA 'precio unitario surge del precio que trajimos de la tabla tafamfil por los M2
End Function

Sub CREAFORMULA(CII%, CODFAM$, M1%, M2%, M3%, VARIANTE$)
    '
    'ARMAMOS LA ESTRUCTURA DE PRODUCTO
20  Call FILTERFILE("STRUCFAM", "!STRUCFAM", 1, CODFAM$)
    CAIT% = ULTREG&("!STRUCFAM")
    For U& = 1 To CAIT%
       XX$ = REGLEIDO$("!STRUCFAM", U&)
       CODCOMPO% = CVI(Mid$(XX$, 3, 2))
       CANCOMPO = CVD(Mid$(XX$, 5, 8))
       MODOCALC$ = TRIM$(UCase$(Mid$(XX$, 13, 32)))
       MODOCALC$ = REPLACAR$(MODOCALC$, " ", "")
       If MODOCALC$ = "M1XM2" Or MODOCALC$ = "M1*M2" Then
            CANTU = CANCOMPO * (M1% / 1000) * (M2% / 1000)
          ElseIf MODOCALC$ = "M1+M2" Then
            CANTU = CANCOMPO * ((M1% / 1000) + (M2% / 1000))
          Else
            Call MENSERR(24, "Modo de Cálculo '" + MODOCALC$ + "' no Procesable.")
            Call FINAL("")
       End If
       '
       EXPLOLIN$(U&) = Space$(128)
       Mid$(EXPLOLIN$(U&), 1, 16) = CODEXT$(CODCOMPO%)
       Mid$(EXPLOLIN$(U&), 49, 16) = FORMATNUM$(CANTU, "F12.4")
    Next U&
    Call AREXPLO(CII%, 1, 2)
    '
End Sub




