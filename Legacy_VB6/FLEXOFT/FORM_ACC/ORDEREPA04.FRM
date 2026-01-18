VERSION 5.00
Begin VB.Form ORDEREPA04 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orden de Reparación"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11805
   Icon            =   "ORDEREPA04.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11805
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2325
      Left            =   0
      TabIndex        =   6
      Top             =   7035
      Width           =   10800
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Generó"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   760
         Left            =   210
         TabIndex        =   31
         Top             =   420
         Width           =   3270
         Begin VB.Label Label19 
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
            Left            =   210
            TabIndex        =   34
            Top             =   315
            Width           =   2850
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
         Height          =   1060
         Left            =   3570
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   105
         Width           =   7155
      End
      Begin VB.Label Label21 
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
         Height          =   300
         Left            =   1680
         TabIndex        =   37
         Top             =   105
         Width           =   960
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Estim.Hs:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   -1680
         TabIndex        =   35
         Top             =   105
         Width           =   3270
      End
      Begin VB.Line Line2 
         X1              =   3570
         X2              =   10725
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Line1 
         X1              =   10725
         X2              =   10725
         Y1              =   105
         Y2              =   1150
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Obs:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   210
         TabIndex        =   30
         Top             =   105
         Width           =   3270
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12000
      Begin VB.CommandButton Command14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   10920
         Picture         =   "ORDEREPA04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Aprueba, Graba e Imprime Presupuesto"
         Top             =   7150
         Width           =   800
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
         Left            =   6585
         TabIndex        =   39
         Text            =   " "
         Top             =   885
         Width           =   1485
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Código - Modelo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   915
         Left            =   210
         TabIndex        =   7
         Top             =   1365
         Width           =   6630
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
            Height          =   300
            Left            =   210
            TabIndex        =   33
            Top             =   495
            Width           =   2325
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
            Height          =   510
            Left            =   2625
            TabIndex        =   8
            Top             =   285
            Width           =   3840
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "O/V Nro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   915
         Left            =   9555
         TabIndex        =   32
         Top             =   1365
         Width           =   1170
         Begin VB.TextBox Text2 
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
            Left            =   105
            TabIndex        =   36
            Text            =   " "
            Top             =   495
            Width           =   960
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Repuestos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   4700
         Left            =   210
         TabIndex        =   24
         Top             =   2310
         Width           =   10515
         Begin VB.Frame Frame9 
            Caption         =   "Controles no Visibles"
            Height          =   2220
            Left            =   2940
            TabIndex        =   48
            Top             =   1365
            Visible         =   0   'False
            Width           =   3375
            Begin VB.Label TOTAL 
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
               Height          =   255
               Index           =   0
               Left            =   1785
               TabIndex        =   54
               Top             =   735
               Width           =   1275
            End
            Begin VB.Label TOTAL 
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
               Height          =   255
               Index           =   2
               Left            =   1785
               TabIndex        =   53
               Top             =   1515
               Width           =   1275
            End
            Begin VB.Label TOTPER 
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
               Height          =   255
               Left            =   1785
               TabIndex        =   52
               Top             =   1260
               Width           =   1275
            End
            Begin VB.Label TOTAL 
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
               Height          =   255
               Index           =   1
               Left            =   1785
               TabIndex        =   51
               Top             =   990
               Width           =   1275
            End
            Begin VB.Label SUMAREP 
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
               Height          =   300
               Left            =   1785
               TabIndex        =   50
               Top             =   315
               Width           =   1275
            End
            Begin VB.Label TOTAMOBRA 
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
               Height          =   300
               Left            =   315
               TabIndex        =   49
               Top             =   315
               Width           =   1275
            End
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
            Height          =   4110
            Left            =   0
            TabIndex        =   26
            Top             =   555
            Width           =   10515
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo          Descripción                                   Cant.   Dispon.  Falta"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   225
            Left            =   70
            TabIndex        =   25
            Top             =   315
            Width           =   10100
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Serie Número"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   915
         Left            =   6930
         TabIndex        =   22
         Top             =   1365
         Width           =   2535
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
            Left            =   105
            TabIndex        =   23
            Text            =   " "
            Top             =   495
            Width           =   2325
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00404080&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   6990
         Left            =   10920
         TabIndex        =   9
         Top             =   0
         Width           =   1200
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   120
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   56
            Top             =   6300
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   57
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Notas"
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
            Picture         =   "ORDEREPA04.frx":0974
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Anotaciones"
            Top             =   4450
            Width           =   650
         End
         Begin VB.CommandButton Command15 
            BackColor       =   &H00E0E0E0&
            Caption         =   "View"
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
            Picture         =   "ORDEREPA04.frx":0C7E
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Ver y Reimprimir Planillas y Presupuestos"
            Top             =   3800
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
            Height          =   650
            Left            =   120
            Picture         =   "ORDEREPA04.frx":0F88
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   100
            Width           =   650
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "C.Cte."
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
            Picture         =   "ORDEREPA04.frx":10D2
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Resumen de Cuenta Corriente de Clientes"
            Top             =   2900
            Width           =   650
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Master"
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
            Picture         =   "ORDEREPA04.frx":13DC
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Acceso a Maestro de Artículos"
            Top             =   1600
            Width           =   650
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cust."
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
            Picture         =   "ORDEREPA04.frx":16E6
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Acceso a Base de Datos de Clientes"
            Top             =   2250
            Width           =   650
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Setup"
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
            Picture         =   "ORDEREPA04.frx":1B28
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   5350
            Width           =   650
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00E0E0E0&
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
            Height          =   650
            Left            =   120
            Picture         =   "ORDEREPA04.frx":1F6A
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   750
            Width           =   650
         End
         Begin VB.Image Image2 
            Height          =   390
            Left            =   -315
            Picture         =   "ORDEREPA04.frx":2274
            Top             =   6615
            Width           =   1515
         End
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
         Left            =   9615
         TabIndex        =   0
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
         TabIndex        =   1
         Top             =   105
         Width           =   180
      End
      Begin VB.Line Line8 
         X1              =   10725
         X2              =   10725
         Y1              =   2445
         Y2              =   7020
      End
      Begin VB.Label TALON 
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
         Left            =   9615
         TabIndex        =   46
         Top             =   890
         Width           =   1110
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   5325
         TabIndex        =   45
         Top             =   150
         Width           =   960
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   4725
         TabIndex        =   44
         Top             =   945
         Width           =   1800
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   5325
         TabIndex        =   43
         Top             =   525
         Width           =   1695
      End
      Begin VB.Label Label16 
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
         Left            =   5850
         TabIndex        =   42
         Top             =   105
         Width           =   2220
      End
      Begin VB.Label Label17 
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
         Left            =   5850
         TabIndex        =   41
         Top             =   495
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Talón Nro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   7770
         TabIndex        =   40
         Top             =   945
         Width           =   1800
      End
      Begin VB.Label REPRESUP 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   225
         Left            =   6510
         TabIndex        =   38
         Top             =   0
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
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
         Left            =   4935
         TabIndex        =   27
         Top             =   1785
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label18 
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
         Height          =   300
         Left            =   9615
         TabIndex        =   21
         Top             =   495
         Width           =   1110
      End
      Begin VB.Label Label15 
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
         Left            =   2310
         TabIndex        =   20
         Top             =   890
         Width           =   2850
      End
      Begin VB.Label Label14 
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
         Height          =   300
         Left            =   1470
         TabIndex        =   19
         Top             =   890
         Width           =   750
      End
      Begin VB.Label Label13 
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
         Left            =   1470
         TabIndex        =   18
         Top             =   490
         Width           =   3690
      End
      Begin VB.Label Label12 
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
         Left            =   1470
         TabIndex        =   17
         Top             =   105
         Width           =   3690
      End
      Begin VB.Label Label11 
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
         Height          =   300
         Left            =   420
         TabIndex        =   16
         Top             =   885
         Width           =   750
      End
      Begin VB.Line Line7 
         X1              =   210
         X2              =   10710
         Y1              =   7000
         Y2              =   7000
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   420
         Picture         =   "ORDEREPA04.frx":4196
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
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   7770
         TabIndex        =   5
         Top             =   525
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
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   8820
         TabIndex        =   4
         Top             =   150
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   210
         TabIndex        =   3
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
End
Attribute VB_Name = "ORDEREPA04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   ANOTAREP04.Label7 = Label18
   ANOTAREP04.Show 1
   Command1.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG04/AMACLI"
   Call CONECRUN(ACONEC$, "Padron Maestro de Clientes")
   Command11.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   CONFREPA04.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call GRAOREPA(OKX%)
   If OKX% > 0 Then
     CANCELPRINT% = 0
     Call PRIOREPA
     Call BLANFORMU
     Command14.Enabled = True
     Label11 = ""
     Hide
   End If
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
    '
    Command15.Enabled = False
    '
3   Call SELECHO("PRESREPA/Indice General Planillas de Reparación")
    NPX& = Val(VALACT1$("PRESREPA"))
    If NPX& < 1 Then GoTo 99
    '
    For UI& = ULTREG&("PRESREPA") To 1 Step -1
      XX$ = REGLEIDO$("PRESREPA", UI&)
      If CVS(Left$(XX$, 4)) = NPX& Then GoTo 4
    Next UI&
    Call COMUNI("Error de Indices - Consulte.")
    Exit Sub
    '
4   If CVS(Mid$(XX$, 125, 4)) < 1 Then
      Call COMUNI("No se ha Emitido Orden de Reparación\para este Presupuesto.")
      GoTo 3
    End If
    '
    NUDOCU$ = TRIM$(Str$(NPX&))
    TIDOCUM$ = "Orden de Reparación " + NUDOCU$ + " "
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
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
5   Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
    '28/03/07 (FIN)
'5   For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    '
'    REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'    GoTo 8
'    '
'7   Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    GoTo 99
'    '
'8   TPSP$ = ""
'    Screen.MousePointer = 11
'    While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 9
'      L& = L& - 1
'    Wend
'    '
'9   CAPAGINAS% = 0
'    For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'95  SPOOLSTRING$ = TPSP$
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).Top = 0
'      EDITFORM.Picture1(KKI%).Refresh
'      EDITFORM.Picture1(KKI%).Height = 7170
'      EDITFORM.Picture1(KKI%).Width = 11700
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    ALTUPAGINA = 7170
'    TOPEPAGINA = 0
'    PAGINACTIVA% = 0
'    CAPAGINAS% = 0
'    '
'10  PPXX% = InStr(TPSP$, Chr$(255))
'    If PPXX% > 0 Then
'      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'      TPSP$ = Mid$(TPSP$, PPXX% + 1)
'      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
'      '
'      If Len(COPRIAT$) > 0 Then
'        COPRI% = Asc(Left$(COPRIAT$, 1))
'        ATRIP$ = Mid$(COPRIAT$, 2)
'        Call SHOWCOMMAND(COPRI%, ATRIP$)
'      End If
'      GoTo 10
'      '
'    End If
'    '
'    Screen.MousePointer = 1
'    PAGINACTIVA% = 0
'    For KKI% = 0 To 5
'       EDITFORM.Picture1(KKI%).Visible = False
'    Next KKI%
'    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
'    EDITFORM.Picture1(PAGINACTIVA%).Refresh
'    EDITFORM.VScroll1.Min = 0
'    EDITFORM.VScroll1.Value = 0
'    EDITFORM.VScroll1.Max = 1
'    If ALTUPAGINA > EDITFORM.Frame1.Height Then
'       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'    End If
'    EDITFORM.Command3.Enabled = False
'    EDITFORM.Command5.Enabled = False
'    EDITFORM.Show 1
'    '
    GoTo 3
    '
99  Command15.Enabled = True
    '
End Sub


Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Label11 = ""
    Hide
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("PRESUP01")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    If DERACCE("RESCUDEU", "Resumen de Cuenta de Deudores") >= 2 Then
      RECUEC04.Text1.TEXT = Label11
      PROLLAMA$ = "PLAPRERE"
      RECUEC04.Show 1
    End If
    Command4.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("AMASTO04", "Maestro de Artículos de Stock")
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Call SELECHO("SELFECHA")
   Text9.TEXT = VALACT1$("SELFECHA")
   FF% = FECHANUM(Text9.TEXT)
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
    Call GRATITFOR(Caption)
    '
    Label19 = USERNAME$
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$
    Label18 = TRIM$(Str$(NUPRESUP&))
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
   Label11 = ""
   Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Label11_Change()
   '
   NC% = Val(Label11)
   Label12 = ""
   Label13 = ""
   Label14 = ""
   Label15 = ""
   Label16 = ""
   Label17 = ""
   '
   If NC% > 0 Then
     If RASOCLI$(NC%) <> "" Then
       Label12 = RASOCLI$(NC%)
       Label13 = DOMICLI$(NC%)
       Label14 = CPOSCLI$(NC%)
       Label15 = LOCACLI$(NC%)
       Label16 = TEL1CLI$(NC%)
       Label17 = FAXCLI$(NC%)
     End If
   End If
   '
End Sub

Private Sub Label18_Change()
   '
   NPX& = Val(Label18)
   Screen.MousePointer = 11
   Call BLANFORMU
   For UI& = ULTREG&("PRESREPA") To 1 Step -1
     XX$ = REGLEIDO$("PRESREPA", UI&)
     If CVS(Left$(XX$, 4)) = NPX& Then
       '
       If CVS(Mid$(XX$, 125, 4)) > 0.5 Then
         Call COMUNI("Imposible Modificar esta Orden de Reparación.\Corresponde O.Venta " + TRIM$(Str$(CVS(Mid$(XX$, 125, 4)))))
         Hide
         Exit Sub
       End If
       '
       REPRESUP = TRIM$(Str$(UI&))
       NCLIE% = CVI(Mid$(XX$, 67, 2))
       Label11 = TRIM$(Str$(NCLIE%))
       CII% = CVI(Mid$(XX$, 69, 2))
       Label10 = CODEXT$(CII%)
       TEPRUEBA = DESCRIT$(CII%)
       Text1 = Mid$(XX$, 71, 15)
       Text10 = Mid$(XX$, 86, 9)
       TALON = Mid$(XX$, 95, 6)
       Text28 = Mid$(XX$, 129, 384)
       Label21 = CSTRING$(Mid$(XX$, 101, 4), 4, 8, 2, 2)
       '
       TOTAMOBRA = TRIM$(CSTRING$(Mid$(XX$, 109, 4), 4, 12, 2, 2))
       SUMAREP = TRIM$(CSTRING$(Mid$(XX$, 113, 4), 4, 12, 2, 2))
       '
       Call CAREPUES
       Call PREREPUES
       '
       On Error Resume Next
       Text10.SetFocus
       On Error GoTo 0
       OPENFORMS = DoEvents
       Screen.MousePointer = 1
       Exit Sub
     End If
   Next UI&
   Screen.MousePointer = 1
   '
End Sub

Private Sub Label8_Click()
   Text13.SetFocus
End Sub

Private Sub SUMAREP_Change()
   TOTAMO = Val(SUMAREP) + Val(TOTAMOBRA)
   TOTAL(0) = CSTRING$(MKS$(TOTAMO), 3, 11, 2, 2)
End Sub

Private Sub TEPRUEBA_Change()
   DESCRI$ = TEPRUEBA.Caption
   If TEPRUEBA.Width < Label8.Width Then
     DESCRI$ = Chr$(10) + Chr$(13) + DESCRI$
   End If
   Label8.Caption = DESCRI$
End Sub

Function VALICLIE%()
   VCC% = 0
   If Val(Label11) > 0 Then
     If Val(Label11) <= 32767 Then
       NCLIE% = Val(Label11)
       If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
         VCC% = 1
       End If
     End If
   End If
   '
99 VALICLIE% = VCC%
   '
End Function
'
Sub CAREPUES()
   '
   Screen.MousePointer = 11
   Call SETULTREG("!MATEROF", 0)
   NPX& = Val(Label18)
   NPI% = NPX& Mod 10000
   RFF$ = RECFILT$("REPUREPA", 2, MKI$(NPI%))
   '
   URE& = ULTREG&("REPUREPA")
30 For KK& = 1 To Len(RFF$) Step 4
     REG& = CVS(Mid$(RFF$, KK&, 4))
     If REG& > 0 Then
       If REG& <= URE& Then
         XX$ = REGLEIDO$("REPUREPA", REG&)
         If CVS(Left$(XX$, 4)) = NPX& Then
           CII% = CVI(Mid$(XX$, 7, 2))
           CANTI = CVS(Mid$(XX$, 9, 4))
           PRUNI = CVS(Mid$(XX$, 13, 4))
           '
           ZZ$ = REGBLAN$("MATEROF")
           Call REPLA(ZZ$, MKI$(CII%), 1, 2)
           Call REPLA(ZZ$, UNIMED$(CII%), 3, 2)
           Call REPLA(ZZ$, MKS$(CANTI), 5, 4)
           Call REPLA(ZZ$, MKD$(CDbl(CANTI)), 9, 8)
           Call REPLA(ZZ$, MKS$(0), 17, 4)
           STDIS = STOCKDIS(CII%)
           Call REPLA(ZZ$, MKS$(STDIS), 21, 4)
           FALTA = CANTI - STDIS: If FALTA < 0 Then FALTA = 0
           Call REPLA(ZZ$, MKS$(FALTA), 25, 4)
           Call REPLA(ZZ$, MKS$(PRUNI), 29, 4)
           Call REGAPP("!MATEROF", ZZ$)
           '
         End If
       End If
     End If
   Next KK&
   Call CIERRARCH("!MATEROF")
   Screen.MousePointer = 1
   '
End Sub
'
Sub PREREPUES()
   '
   Screen.MousePointer = 11
   List1.Clear
   SUMA# = 0
   For U& = 1 To ULTREG&("!MATEROF")
     ZZ$ = REGLEIDO$("!MATEROF", U&)
     TTXX$ = Space$(128)
     CII% = CVI(Left$(ZZ$, 2))
     Call REPLA(TTXX$, CODEXT$(CII%), 1, 16)
     Call REPLA(TTXX$, DESCRIT$(CII%), 17, 42)
     NEBRUX$ = Mid$(ZZ$, 5, 4)
     Call REPLA(TTXX$, CSTRING$(NEBRUX$, 4, 8, 1, 2), 59, 8)
     COSUX$ = Mid$(ZZ$, 21, 4)
     Call REPLA(TTXX$, CSTRING$(COSUX$, 4, 10, 1, 2), 67, 10)
     ITOX$ = Mid$(ZZ$, 25, 4)
     Call REPLA(TTXX$, CSTRING$(ITOX$, 4, 8, 1, 2), 77, 8)
     SUMA# = SUMA# + Val(Mid$(TTXX$, 74, 11))
     List1.AddItem TTXX$
   Next U&
   Screen.MousePointer = 1
   '
End Sub
'
Function NUPRESUP&()
   '
   NPP& = 1
   '
   FIN& = ULTREG&("PRESREPA")
   If FIN& > 0 Then
      For U& = FIN& To FIN& - 128 Step -1
         If U& < 1 Then GoTo 99
         XX$ = REGLEIDO$("PRESREPA", U&)
         NPPX& = CVS(Mid$(XX$, 1, 4))
         If NPPX& >= NPP& Then
            NPP& = 1 + NPPX&
         End If
      Next U&
   End If
   '
99 NUPRESUP& = NPP&
   '
End Function
'
Private Sub Text9_GotFocus()
   On Error Resume Next
   Command9.SetFocus
   On Error GoTo 0
End Sub

Sub PRIOREPA()
    '
    FORIPRE$ = TRIM$(CONTROL$("FOREPAR", "DOCUREP2"))
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Imposible Imprimir Planilla de Reparación.\Falta Definir Formulario Electrónico.")
       Exit Sub
    End If
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Planilla de Reparación.\No existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
10  NC% = Val(Label11)
    TIPODOC$ = "Orden de Reparación"
    DESTIDOC% = NC%
    NUMEDOC$ = TRIM$(Label18)
    FECHDOC$ = TRIM$(Text9.TEXT)
    '
    CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = TRIM$(Label10)
    CODPARAM$(2) = "DES-CONJ": DOCPARAM$(2) = TRIM$(TEPRUEBA)
    CODPARAM$(3) = "REF-MAT1": DOCPARAM$(3) = TRIM$(Text1)
    CODPARAM$(4) = "MEDIDA-1": DOCPARAM$(4) = TRIM$(Text10)
    CODPARAM$(5) = "MEDIDA-2": DOCPARAM$(5) = TRIM$(Text2)
    CODPARAM$(6) = "REF-MAT2": DOCPARAM$(6) = TRIM$(Label21)
    CODPARAM$(7) = "MEDIDA-3": DOCPARAM$(7) = TRIM$(TALON)
    '
    CODPARAM$(8) = "PRE-NETO": DOCPARAM$(8) = TRIM$(SUMAREP)
    CODPARAM$(9) = "BRU-ITEM": DOCPARAM$(9) = TRIM$(TOTAMOBRA)
    CODPARAM$(10) = "IMP-NETO": DOCPARAM$(10) = TRIM$(TOTAL(0))
    CODPARAM$(11) = "IMP-IVA": DOCPARAM$(11) = TRIM$(TOTAL(1))
    CODPARAM$(12) = "IMP-TOTA": DOCPARAM$(12) = TRIM$(TOTAL(2))
    CODPARAM$(13) = "PERC-IVA": DOCPARAM$(13) = TRIM$(TOTPER)
    CAPARDOC% = 13
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CODTABU1$(4) = "PRE-LIST"
    CACOLTAB1% = 4
    '
    CAITAB1% = 0
    For U& = 1 To ULTREG&("!MATEROF")
      ZZ$ = REGLEIDO$("!MATEROF", U&)
      CII% = CVI(Left$(ZZ$, 2))
      CAITAB1% = CAITAB1% + 1
      TETABU1$(1, CAITAB1%) = CODEXT$(CII%)
      TETABU1$(2, CAITAB1%) = DESCRIT$(CII%)
      NEBRUX$ = Mid$(ZZ$, 5, 4)
      TETABU1$(3, CAITAB1%) = CSTRING$(NEBRUX$, 4, 8, 1, 2)
      TETABU1$(4, CAITAB1%) = ".............................."
    Next U&
    '
    Call SETULTREG("!OBSERVOF", 0)
    Call FRATEXTO(Text28.TEXT, 56)
    For YX& = 1 To CARETEX%
       Call REGAPP("!OBSERVOF", RETEX$(YX&))
    Next YX&
    '
    Call PRINTDOC("@" + FORIPRE$)
    '
End Sub

Sub GRAOREPA(OKX%)
   '
   OKX% = 0
   Call BLOQARCH("PRESREPA")
   Call BLOQARCH("REPUREPA")
   '
   REG& = Val(REPRESUP)
   If REG& < 1 Or REG& > ULTREG&("PRESREPA") Then
      Call COMUNI("Imposible Registrar - Error de Indices")
      GoTo 98
   End If
   '
   XX$ = REGLEIDO$("PRESREPA", REG&)
   If CVS(Left$(XX$, 4)) <> Val(Label18) Then
      Call COMUNI("Error de Indices - Imposible Registrar")
      GoTo 98
   End If
   '
   Screen.MousePointer = 11
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLIE% = Val(Label11)
   '
   CIIX% = CODINT%(Label10)
   If CIIX% < 1 Or CIIX% > NUMAS% Then
      MERRO$ = "Código de Artículo no Válido"
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
   If CVS(Mid$(XX$, 125, 4)) > 0.5 Then
      MERRO$ = "Imposible Registrar - O.Venta " + TRIM$(Str$(CVS(Mid$(XX$, 125, 4))))
      GoTo 98
   End If
   '
   If Val(Text2) < 1 Then
      MERRO$ = "Imposible Generar O/Rep. - Falta Número O/Venta"
      OKX% = 0
      GoTo 98
   End If
   NORVENTA& = Val(Text2)
   '
   Call ABRESERVA
   '
   Call REPLA(XX$, TRIM$(Text1), 71, 15)
   Call REPLA(XX$, TRIM$(Text10), 86, 9)
   Call REPLA(XX$, TRIM$(TALON), 95, 6)
   Call REPLA(XX$, MKI$(0), 123, 2)
   Call REPLA(XX$, MKS$(Val(Text2)), 125, 4)
   MODIOBSE% = 0: If TRIM$(Text28) <> TRIM$(Mid$(XX$, 129, 384)) Then MODIOBSE% = 1
   Call REPLA(XX$, Text28 + Space$(384), 129, 384)
   Call GRAREG("PRESREPA", XX$, REG&)
   Call CIERRARCH("PRESREPA")
   '
   NPLAL& = Val(Label18)
   NPLAI% = NPLAL& Mod 10000
   NOREP$ = "OR-" + TRIM$(Str$(NPLAL&))
   REBLA$ = REGBLAN$("REPUREPA")
   Call REPLA(REBLA$, MKS$(NPLAL&), 1, 4)
   Call REPLA(REBLA$, MKI$(NPLAI%), 5, 2)
   Call REPLA(REBLA$, NOREP$, 33, 12)
   '
   RFF$ = RECFILT$("REPUREPA", 2, MKI$(NPLAI%))
   RCC$ = ""
   URE& = ULTREG&("REPUREPA")
   For KK& = 1 To Len(RFF$) Step 4
     RGG& = CVS(Mid$(RFF$, KK&, 4))
     If RGG& > 0 Then
       If RGG& <= URE& Then
         XX$ = REGLEIDO$("REPUREPA", RGG&)
         If CVS(Left$(XX$, 4)) = NPLAL& Then
           RCC$ = RCC$ + MKS$(RGG&)
         End If
       End If
     End If
   Next KK&
   '
   KK& = -3
   For U& = 1 To ULTREG&("!MATEROF")
     ZZ$ = REGLEIDO$("!MATEROF", U&)
     XX$ = REBLA$
     CII% = CVI(Left$(ZZ$, 2))
     CANTI = CVS(Mid$(ZZ$, 5, 4))
     PRUNI = CVS(Mid$(ZZ$, 29, 4))
     '
     Call REPLA(XX$, MKI$(CII%), 7, 2)
     Call REPLA(XX$, MKS$(CANTI), 9, 4)
     Call REPLA(XX$, MKS$(PRUNI), 13, 4)
     Call REPLA(XX$, MKS$(PRUNI * CANTI), 17, 4)
     Call REPLA(XX$, MKS$(CANTI), 21, 4)
     Call REPLA(XX$, MKS$(0), 25, 4)
     Call REPLA(XX$, MKS$(CANTI), 29, 4)
     KK& = KK& + 4
     If KK& <= Len(RCC$) Then
          RGG& = CVS(Mid$(RCC$, KK&, 4))
          Call GRAREG("REPUREPA", XX$, RGG&)
        Else
          Call REGAPP("REPUREPA", XX$)
     End If
     '
     With Reserva
        .AddNew
        !Cod_Inte = CII%
        !COD_EXTE = CODEXT$(CII%)
        !Descrip = DESCRIT0$(CII%)
        !IdSubOr = NOREP$
        !FechRes = Now
        !CantRes = CANTI
        !CantCons = 0
        !Saldo_Entre = CANTI
        !IdenLote = ""
        !Deposito = DEPOPRE%(CII%)
        !UsoUnit = CANTI
        .Update
      End With
     
   Next U&
   Reserva.Close
   '
   While KK& < Len(RCC$) - 3
     KK& = KK& + 4
     RGG& = CVS(Mid$(RCC$, KK&, 4))
     Call GRAREG("REPUREPA", String$(128, 0), RGG&)
   Wend
   '
   Call CIERRARCH("REPUREPA")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   TTXYZ$ = Space$(92)
   NROREP& = Val(Label18)
   HOII% = HOY(HOS$)
   Call REPLA(TTXYZ$, HOS$, 1, 8)
   Call REPLA(TTXYZ$, USERTER$, 69, 24)
   Call REPLA(TTXYZ$, "Orden de Reparación " + TRIM$(Str$(NROREP&)) + " - O/Venta " + TRIM$(Text2) + " - Talón " + TRIM$(TALON), 11, 56)
   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   '
   If MODIOBSE% > 0 Then
      TTXYZ$ = Space$(92)
      Call REPLA(TTXYZ$, Space$(56), 11, 56)
      Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
      Call FRATEXTO(Text28, 46)
      For KKI% = 1 To CARETEX%
        TTXX$ = Space$(92)
        Call REPLA(TTXX$, RETEX$(KKI%), 10, 46)
        If KKI% = 1 Then Call REPLA(TTXX$, " Observ: ", 1, 9)
        Call REGAPP("ANOTAREP", Left$(MKS$(NROREP&) + Space$(10) + TTXX$, 96))
      Next KKI%
   End If
   '
   Call REGAPP("ANOTAREP", MKS$(NROREP&) + String$(88, "="))
   Call CIERRARCH("ANOTAREP")
   '
   RFF$ = RECFILT$("ANOTAREP", 1, MKS$(NROREP&))
   For U& = 1 To Len(RFF$) Step 4
     REAREP& = CVS(Mid$(RFF$, U&, 4))
     TTXX$ = REGLEIDO$("ANOTAREP", REAREP&)
     Call REPLA(TTXX$, MKS$(NORVENTA&), 1, 4)
     Call REGAPP("ANOTAPED", TTXX$)
   Next U&
   Call CIERRARCH("ANOTAPED")
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   OKX% = 1
   Screen.MousePointer = 1
   Exit Sub
   '
98 Call CIERRARCH("PRESREPA")
   Call CIERRARCH("REPUREPA")
   Screen.MousePointer = 1
   Call MENSERR(24, MERRO$)
   '
End Sub

Sub BLANFORMU()
   '
   Label8 = ""
   Label12 = ""
   Label13 = ""
   Label14 = ""
   Label15 = ""
   Label16 = ""
   Label17 = ""
   Label21 = ""
   Text1 = ""
   Text2 = ""
   Text10 = ""
   Text13 = ""
   Text28 = ""
   TEPRUEBA = ""
   TALON = ""
   List1.Clear
   '
End Sub

Private Sub Total_Change(Index As Integer)
   If Index = 0 Then
     NCLIE% = Val(Label11)
     If NCLIE% > 0 Then
       PIVA% = PIVACLI%(NCLIE%)
       FACIVA = ALIVA / 100
       If PIVA% = 2 Then
            FACIVA = 1.5 * ALICU / 100
          ElseIf PIVA% = 5 Then
            FACIVA = 0
            ALICU = 0
       End If
       TOTAL(1) = CSTRING$(MKS$(FACIVA * Val(TOTAL(0))), 4, 11, 2, 2)
       Call CALPERIB
       TOTAL(2) = CSTRING$(MKS$(Val(TOTAL(0)) + Val(TOTAL(1)) + Val(TOTPER)), 4, 11, 2, 2)
     End If
   End If
End Sub

Private Sub TOTAMOBRA_Change()
   TOTAMO = Val(SUMAREP) + Val(TOTAMOBRA)
   TOTAL(0) = CSTRING$(MKS$(TOTAMO), 3, 11, 2, 2)
End Sub

Sub CALPERIB()
   '
   Static CTTXX%, TASAA, TASAB, TASAEXP, MINIPER
   '
   If CTTXX% < 1 Then
     TASAA = Val(CONTROL$("PERIBRU", "TASA-A"))
     TASAB = Val(CONTROL$("PERIBRU", "TASA-B"))
     TASAEXP = Val(CONTROL$("PERIBRU", "TASA-EXP"))
     MINIPER = Val(CONTROL$("PERIBRU", "MINIPER"))
     CTTXX% = 1
   End If
   '
   TASAX = TASAA
   NCLIE% = Val(Label11)
   If NCLIE% < 1 Then Exit Sub
   '
   VARIDC% = 1: PVCLX% = PIVACLI%(NCLIE%)
   If PVCLX% = 0 Or PVCLX% = 4 Then VARIDC% = 2
   If PVCLX% = 5 Or PVCLX% = 7 Then VARIDC% = 3
   '
   If VARIDC% = 2 Then TASAX = TASAB
   If VARIDC% = 3 Then TASAX = TASAEXP
   '
   Text15 = ""
   REPERCI$ = FILTERGETRECORD$("DAPIBCLI", 1, MKI$(NCLIE%))
   CATIB1% = Asc(Mid$(REPERCI$, 3, 1))
   PORCONVE = CVS(Mid$(REPERCI$, 5, 4))
   If PORCONVE > 100 Then PORCONVE = 100
   '
   If CATIB1% > 0 Then         ' NO ESTA EXENTO
     If PORCONVE >= 10.1 Then  ' PARTICIPACION CONVENIO MULTILATERAL
       BIMPON# = Val(TOTAL(0))
       If VARIDC% = 2 Then BIMPON# = BIMPON# + Val(TOTAL(1))
       If BIMPON# * TICAMONE(MONEMI%) >= MINIPER Then
         PERSALTA# = (BIMPON# * PORCONVE / 100) * (TASAX / 100)
         TOTPER = FORMATNUM$(PERSALTA#, "F11.2")
       End If
     End If
   End If
   '
End Sub
'
