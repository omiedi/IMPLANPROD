VERSION 5.00
Begin VB.Form PRESURE07 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuestos por Reparaciones - Generación y Emisión"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11790
   Icon            =   "PRESURE07.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11790
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2325
      Left            =   0
      TabIndex        =   13
      Top             =   7170
      Width           =   10800
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
         Height          =   1020
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   0
         Width           =   7155
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Percepción IB:"
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
         Left            =   7140
         TabIndex        =   57
         Top             =   555
         Width           =   1905
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
         Height          =   260
         Index           =   1
         Left            =   9135
         TabIndex        =   38
         Top             =   260
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
         Left            =   9135
         TabIndex        =   56
         Top             =   520
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
         Height          =   260
         Index           =   2
         Left            =   9135
         TabIndex        =   39
         Top             =   780
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
         Height          =   260
         Index           =   0
         Left            =   9135
         TabIndex        =   37
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Presupuesto:"
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
         Left            =   7140
         TabIndex        =   16
         Top             =   820
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A:"
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
         Left            =   7140
         TabIndex        =   15
         Top             =   300
         Width           =   1905
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
         Index           =   0
         Left            =   7140
         TabIndex        =   14
         Top             =   40
         Width           =   1905
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12000
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "O-Rep."
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
         Left            =   10880
         TabIndex        =   58
         ToolTipText     =   "Marque Casilla para Orden de Reparación Directa"
         Top             =   7200
         Width           =   1200
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
         TabIndex        =   49
         Text            =   " "
         Top             =   885
         Width           =   1485
      End
      Begin VB.TextBox TALON 
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
         TabIndex        =   48
         Text            =   " "
         Top             =   885
         Width           =   1110
      End
      Begin VB.TextBox Text3 
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
         Left            =   3360
         TabIndex        =   1
         Top             =   6650
         Width           =   960
      End
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
         Left            =   5565
         TabIndex        =   2
         Top             =   6650
         Width           =   960
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
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
         Height          =   4230
         Left            =   210
         TabIndex        =   32
         ToolTipText     =   "Doble-Click Aquí para Editar Repuestos"
         Top             =   2310
         Width           =   10515
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
            Height          =   3660
            Left            =   0
            TabIndex        =   34
            Top             =   555
            Width           =   10515
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo          Descripción                     St.Disp    Cant.  Pr.Unit    Importe"
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
            TabIndex        =   33
            ToolTipText     =   "Doble-Click Aquí para Editar Repuestos"
            Top             =   315
            Width           =   10100
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
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
         Left            =   7980
         TabIndex        =   31
         Top             =   1365
         Width           =   2745
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
            Top             =   495
            Width           =   2325
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   7095
         Left            =   10880
         TabIndex        =   19
         Top             =   0
         Width           =   1200
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   150
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   60
            Top             =   6405
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   61
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
            Left            =   150
            Picture         =   "PRESURE07.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Anotaciones"
            Top             =   4550
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
            Left            =   150
            Picture         =   "PRESURE07.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Ver y Re-Imprimir"
            Top             =   3900
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
            Left            =   150
            Picture         =   "PRESURE07.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   5
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
            Left            =   150
            Picture         =   "PRESURE07.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   2990
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
            Left            =   150
            Picture         =   "PRESURE07.frx":0D72
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Acceso a Maestro de Artículos"
            Top             =   1700
            Width           =   650
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cust"
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
            Left            =   150
            Picture         =   "PRESURE07.frx":107C
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Acceso a Base de Datos de Clientes"
            Top             =   2340
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
            Left            =   150
            Picture         =   "PRESURE07.frx":14BE
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   5460
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
            Left            =   150
            Picture         =   "PRESURE07.frx":1900
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   750
            Width           =   650
         End
         Begin VB.Image Image2 
            Height          =   390
            Left            =   -300
            Picture         =   "PRESURE07.frx":1C0A
            Top             =   6720
            Width           =   1515
         End
      End
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
         Height          =   690
         Left            =   10880
         Picture         =   "PRESURE07.frx":3B2C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Aprueba, Graba e Imprime Presupuesto"
         Top             =   7530
         Width           =   855
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   17
         Top             =   1365
         Width           =   7680
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
            TabIndex        =   8
            Top             =   495
            Width           =   180
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
            Height          =   300
            Left            =   210
            TabIndex        =   46
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
            Left            =   2835
            TabIndex        =   18
            Top             =   285
            Width           =   4680
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
         TabIndex        =   6
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
         TabIndex        =   7
         Top             =   105
         Width           =   180
      End
      Begin VB.Label Label2 
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
         Top             =   495
         Width           =   2220
      End
      Begin VB.Label Label19 
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
         TabIndex        =   50
         Top             =   945
         Width           =   1800
      End
      Begin VB.Line Line1 
         X1              =   9030
         X2              =   9030
         Y1              =   6510
         Y2              =   7060
      End
      Begin VB.Label REPRESUP 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   225
         Left            =   6510
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "M.de Obra:"
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
         Left            =   -735
         TabIndex        =   45
         Top             =   6690
         Width           =   2115
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Estimado de Horas:"
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
         Left            =   1365
         TabIndex        =   43
         Top             =   6720
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Hora:"
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
         Left            =   3570
         TabIndex        =   42
         Top             =   6720
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Importe:"
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
         Left            =   5670
         TabIndex        =   41
         Top             =   6720
         Width           =   1905
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
         Left            =   7665
         TabIndex        =   40
         Top             =   6645
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
         Left            =   9135
         TabIndex        =   36
         Top             =   6645
         Width           =   1275
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
         TabIndex        =   35
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
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9615
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   885
         Width           =   750
      End
      Begin VB.Line Line7 
         X1              =   210
         X2              =   10710
         Y1              =   7065
         Y2              =   7065
      End
      Begin VB.Line Line8 
         X1              =   10725
         X2              =   10725
         Y1              =   2445
         Y2              =   7060
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   420
         Picture         =   "PRESURE07.frx":4196
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
Attribute VB_Name = "PRESURE07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   If Val(Label18) > 0 Then
     ANOTAREP07.Label7 = Label18
     ANOTAREP07.Show 1
   End If
   Command1.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG07/AMACLI"
   Call CONECRUN(ACONEC$, "Padron Maestro de Clientes")
   Command11.Enabled = True
End Sub


Private Sub Command13_Click()
   Command13.Enabled = False
   CONFREPA07.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call GRAPRESU(OKX%)
   If OKX% > 0 Then
     CANCELPRINT% = 0
     If Check1.Value = 1 Then CANCELPRINT = 1
     Call PRIPRESU
     '
     If Check1.Value = 1 Then
       ORDEREPA07.Label18 = TRIM$(Label18)
       ORDEREPA07.Show 1
     End If
     
     Call BLANFORMU
     Label18 = ""
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
    NUDOCU$ = TRIM$(Str$(NPX&))
    TIDOCU1$ = "Planilla de Reparación " + NUDOCU$ + " "
    TIDOCU2$ = "Presupuesto de Reparación " + NUDOCU$ + " "
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCU1$) > 4 Or InStr(XX$, TIDOCU2$) > 4 Then
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
    '
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
      RECUEC07.Text1.TEXT = Label11
      PROLLAMA$ = "PRESURE"
      RECUEC07.Show 1
    End If
    Command4.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("AMASTOO4", "Maestro de Artículos de Stock")
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
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$
    '
    CUHOR = Val(CONTROL$("REPARA", "VALHORA"))
    COEFI = Val(CONTROL$("REPARA", "COEFPREA"))
    Text2 = CSTRING$(MKS$(CUHOR * COEFI), 4, 8, 2, 2)
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

Private Sub Frame4_DblClick()
   Call EDITREPU
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
   Call BLANFORMU
   For UI& = ULTREG&("PRESREPA") To 1 Step -1
     XX$ = REGLEIDO$("PRESREPA", UI&)
     If CVS(Left$(XX$, 4)) = NPX& Then
       '
       If CVS(Mid$(XX$, 125, 4)) > 0.5 Then
         Call COMUNI("Imposible Modificar este Presupuesto.\Corresponde O.Venta " + TRIM$(Str$(CVS(Mid$(XX$, 125, 4)))))
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
       Text3 = CSTRING$(Mid$(XX$, 101, 4), 4, 8, 2, 2)
       If CVS(Mid$(XX$, 105, 4)) > 0.5 Then
         Text2 = CSTRING$(Mid$(XX$, 105, 4), 4, 8, 2, 2)
       End If
       '
       Call CAREPUES
       Call PREREPUES
       '
       On Error Resume Next
       Text10.SetFocus
       On Error GoTo 0
       OPENFORMS = DoEvents
       Exit Sub
     End If
   Next UI&
   '
End Sub

Private Sub Label9_DblClick()
   Call EDITREPU
End Sub

Private Sub List1_DblClick()
   '
   REACTI& = List1.ListIndex + 1
   CODACTI$ = TRIM$(Left$(List1.TEXT, 16))
   '
   Call LEEEXPLOMUL(CODACTI$, 1, 1)
   If CAIT% < 1 Then             ' producto sin estructura
     Call COMUNI("Imposible Abrir\este Repuesto (" + TRIM$(CODACTI$) + ").")
     Exit Sub
   End If
   '
   Screen.MousePointer = 11
   Call COPYSTRU("MATEROF", "MATEROG")
   Call SETULTREG("!MATEROG", 0)
   '
   
   For UI& = 1 To REACTI& - 1
     ZZ$ = REGLEIDO$("!MATEROF", UI&)
     Call REGAPP("!MATEROG", ZZ$)
   Next UI&
   ZZ$ = REGLEIDO$("!MATEROF", REACTI&)
   CABACO = CVS(Mid$(ZZ$, 5, 4))
   '
   For I1% = 1 To CAIT%
     If CIJ%(I1%) > 0 Then
       If CIJ%(I1%) <= NUMAS% Then
         If Abs(USOI(I1%)) > 0 Then
           '
           ZZ$ = REGBLAN$("MATEROF")
           Call REPLA(ZZ$, MKI$(CIJ%(I1%)), 1, 2)
           Call REPLA(ZZ$, UNIMED$(CIJ%(I1%)), 3, 2)
           NEBRU = USOI(I1%)
           Call REPLA(ZZ$, MKS$(CABACO * NEBRU), 5, 4)
           Call REPLA(ZZ$, MKD$(CDbl(CABACO * USOI(I1%))), 9, 8)
           Call REPLA(ZZ$, MKS$(0), 17, 4)
           'STDIS = STOCKDIS(CIJ%(I1%))
           'Call REPLA(ZZ$, MKS$(STDIS), 21, 4)
           'FALTA = NEBRU - STDIS: If FALTA < 0 Then FALTA = 0
           'Call REPLA(ZZ$, MKS$(FALTA), 25, 4)
           Call REGAPP("!MATEROG", ZZ$)
           '
         End If
       End If
     End If
   Next I1% '
   
   For UI& = REACTI& + 1 To ULTREG&("!MATEROF")
     ZZ$ = REGLEIDO$("!MATEROF", UI&)
     Call REGAPP("!MATEROG", ZZ$)
   Next UI&
   '
   Call CIERRARCH("!MATEROG")
   Call SETULTREG("!MATEROF", 0)
   For UI& = 1 To ULTREG&("!MATEROG")
     ZZ$ = REGLEIDO$("!MATEROG", UI&)
     Call REGAPP("!MATEROF", ZZ$)
   Next UI&
   Call CIERRARCH("!MATEROF")
   Call CIERRARCH("!MATEROF")
   '
   Call PREREPUES
   Screen.MousePointer = 1
   '
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


Private Sub Text2_Change()
   TOTAMO = Val(Text3) * Val(Text2)
   TOTAMOBRA = CSTRING$(MKS$(TOTAMO), 3, 11, 2, 2)
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
           '
           ZZ$ = REGBLAN$("MATEROF")
           Call REPLA(ZZ$, MKI$(CII%), 1, 2)
           Call REPLA(ZZ$, UNIMED$(CII%), 3, 2)
           Call REPLA(ZZ$, MKS$(CANTI), 5, 4)
           Call REPLA(ZZ$, MKD$(CDbl(CANTI)), 9, 8)
           Call REPLA(ZZ$, MKS$(0), 17, 4)
           Call REGAPP("!MATEROF", ZZ$)
           '
         End If
       End If
     End If
   Next KK&
   Call CIERRARCH("!MATEROF")
   '
End Sub
'
Sub PREREPUES()
   '
   Screen.MousePointer = 11
   ORIPRE$ = TRIM$(CONTROL$("REPARA", "LIPREREP"))
   ARCHIPRE$ = ""
   If Val(ORIPRE$) > 0 Then
     While Len(ORIPRE$) < 3: ORIPRE$ = "0" + ORIPRE$: Wend
     ARCHIPRE$ = "LIPRE" + ORIPRE$
   End If
   '
   List1.Clear
   SUMA# = 0
   SUMAREP = ""
   COEFI = Val(CONTROL$("REPARA", "COEFPREA"))
   If COEFI < 0.05 Then COEFI = 1
   For U& = 1 To ULTREG&("!MATEROF")
     ZZ$ = REGLEIDO$("!MATEROF", U&)
     TTXX$ = Space$(128)
     CII% = CVI(Left$(ZZ$, 2))
     Call REPLA(TTXX$, CODEXT$(CII%), 1, 16)
     Call REPLA(TTXX$, DESCRIT$(CII%), 17, 31)
     STDIX$ = MKS$(STOCKDIS(CII%))
     Call REPLA(TTXX$, CSTRING$(STDIX$, 4, 8, 1, 2), 48, 8)
     NEBRUX$ = Mid$(ZZ$, 5, 4)
     Call REPLA(TTXX$, CSTRING$(NEBRUX$, 4, 8, 1, 2), 56, 8)
     '
     PREBASI = CVS(Mid$(ZZ$, 29, 4)) / COEFI
     If PREBASI < 0.005 Then
       If ARCHIPRE$ = "" Then
            PREBASI = COSCALCU(CII%, MODOCO%)
          Else
            PREBASI = PRELISTA#(ARCHIPRE$, CII%)
       End If
       Call REPLA(ZZ$, MKS$(PREBASI * COEFI), 29, 4)
       Call GRAREG("!MATEROF", ZZ$, U&)
     End If
     '
     COSUX$ = MKS$(COEFI * PREBASI) ' modificado a partir de 24 mar 2004 para tomar costo calculado de costos
     Call REPLA(TTXX$, CSTRING$(COSUX$, 4, 10, 2, 2), 64, 10)
     ITOX$ = MKS$(CVS(NEBRUX$) * CVS(COSUX$))
     Call REPLA(TTXX$, CSTRING$(ITOX$, 4, 11, 2, 2), 74, 11)
     SUMA# = SUMA# + Val(Mid$(TTXX$, 74, 11))
     List1.AddItem TTXX$
   Next U&
   Call CIERRARCH("!MATEROF")
   SUMAREP = CSTRING$(MKD$(SUMA#), 4, 11, 2, 2)
   Screen.MousePointer = 1
   '
End Sub
'
Sub BLANFORMU()
   '
   Label8 = ""
   Label10 = ""
   Label12 = ""
   Label13 = ""
   Label14 = ""
   Label15 = ""
   Label16 = ""
   Label17 = ""
   Text1 = ""
   Text10 = ""
   TALON = ""
   TEPRUEBA = ""
   Text28 = ""
   Text3 = ""
   Text2 = ""
   List1.Clear
   '
   Check1.Value = 0
   '
   CUHOR = Val(CONTROL$("REPARA", "VALHORA"))
   COEFI = Val(CONTROL$("REPARA", "COEFPREA"))
   Text2 = CSTRING$(MKS$(CUHOR * COEFI), 4, 8, 2, 2)
   '
   For KK% = 0 To 2
      TOTAL(KK%) = ""
   Next KK%
   TOTPER = ""
   '
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   '
End Sub
'
Function NUPRESUP&()
   '
   NPP& = 1
   '
   FIN& = ULTREG&("PRESUPU")
   If FIN& > 0 Then
      For U& = FIN& To FIN& - 128 Step -1
         If U& < 1 Then GoTo 99
         XX$ = REGLEIDO$("PRESUPU", U&)
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

Private Sub Text2_LostFocus()
   Text2 = CSTRING$(MKS$(Val(Text2)), 4, 8, 2, 2)
End Sub

Private Sub Text3_Change()
   TOTAMO = Val(Text3) * Val(Text2)
   TOTAMOBRA = CSTRING$(MKS$(TOTAMO), 3, 11, 2, 2)
End Sub

Private Sub Text3_LostFocus()
   Text3 = CSTRING$(MKS$(Val(Text3)), 4, 8, 2, 2)
End Sub

Private Sub Text9_GotFocus()
   Command9.SetFocus
End Sub

Private Sub Total_Change(Index As Integer)
   If Index = 0 Then
     NCLIE% = Val(Label11)
     If NCLIE% > 0 Then
'       PIVA% = PIVACLI%(NCLIE%)
'       FACIVA = ALIVA / 100
'       If PIVA% = 2 Then
'            FACIVA = 1.5 * ALICU / 100
'          ElseIf PIVA% = 5 Then
'            FACIVA = 0
'            ALICU = 0
'       End If
'       TOTAL(1) = CSTRING$(MKS$(FACIVA * Val(TOTAL(0))), 4, 11, 2, 2)
       Call CALTOIVA
       Call CALPERIB
       TOTAL(2) = CSTRING$(MKS$(Val(TOTAL(0)) + Val(TOTAL(1)) + Val(TOTPER)), 4, 11, 2, 2)
     End If
   End If
End Sub
'
Sub PRIPRESU()
    '
    FORIPRE$ = TRIM$(CONTROL$("FOREPAR", "DOCUREP1"))
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Imposible Imprimir Presupuesto de Reparación.\Falta Definir Formulario Electrónico.")
       Exit Sub
    End If
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Presupuesto de Reparación.\No existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
10  NC% = Val(Label11)
    TIPODOC$ = "Presupuesto de Reparación"
    DESTIDOC% = NC%
    NUMEDOC$ = TRIM$(Label18)
    FECHDOC$ = TRIM$(Text9.TEXT)
    '
    CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = TRIM$(Label10)
    CODPARAM$(2) = "DES-CONJ": DOCPARAM$(2) = TRIM$(TEPRUEBA)
    CODPARAM$(3) = "REF-MAT1": DOCPARAM$(3) = TRIM$(Text1)
    CODPARAM$(4) = "MEDIDA-1": DOCPARAM$(4) = TRIM$(Text10)
    CODPARAM$(5) = "REF-MAT2": DOCPARAM$(5) = TRIM$(Text3)
    CODPARAM$(6) = "REF-MAT3": DOCPARAM$(6) = TRIM$(Text2)
    CODPARAM$(7) = "PRE-NETO": DOCPARAM$(7) = TRIM$(SUMAREP)
    CODPARAM$(8) = "BRU-ITEM": DOCPARAM$(8) = TRIM$(TOTAMOBRA)
    CODPARAM$(9) = "IMP-NETO": DOCPARAM$(9) = TRIM$(TOTAL(0))
    CODPARAM$(10) = "IMP-IVA": DOCPARAM$(10) = TRIM$(TOTAL(1))
    CODPARAM$(11) = "IMP-TOTA": DOCPARAM$(11) = TRIM$(TOTAL(2))
    PORI$ = ""
    On Error Resume Next
    PORI$ = TRIM$(Str$(100 * TOTAL(1) / TOTAL(0)))
    On Error GoTo 0
    CODPARAM$(12) = "POR-IVA": DOCPARAM$(12) = TRIM$(PORI$)
    CODPARAM$(13) = "MEDIDA-3": DOCPARAM$(13) = TRIM$(TALON)
    CODPARAM$(14) = "PERC-IVA": DOCPARAM$(14) = TRIM$(TOTPER)
    CAPARDOC% = 14
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CODTABU1$(4) = "PRE-LIST"
    CODTABU1$(5) = "NET-ITEM"
    CACOLTAB1% = 5
    '
    CAITAB1% = 0
    For U& = 1 To List1.ListCount
      ZZ$ = List1.List(U& - 1)
      CAITAB1% = CAITAB1% + 1
      TETABU1$(1, CAITAB1%) = Left$(ZZ$, 16)
      CII% = CODINT%(Left$(ZZ$, 16))
      TETABU1$(2, CAITAB1%) = DESCRIT$(CII%)
      TETABU1$(3, CAITAB1%) = TRIM$(Mid$(ZZ$, 56, 8))
      TETABU1$(4, CAITAB1%) = TRIM$(Mid$(ZZ$, 64, 10))
      TETABU1$(5, CAITAB1%) = TRIM$(Mid$(ZZ$, 74, 11))
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
    '
    FORIPRE$ = TRIM$(CONTROL$("FOREPAR", "DOCUREP3"))
    If FORIPRE$ = "" Then
       Exit Sub
    End If
    '
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Presupuesto de Reparación.\No existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    VDEF$ = Space$(96) + Space$(48)
    Call REPLA(VDEF$, "5 Días hábiles de recibida su órden de compra.", 1, 48)
    Call REPLA(VDEF$, "Contado contra entrega", 49, 48)
15  OBX$ = OBJPLA$("PRESUREPA", VDEF$)
    If OBX$ = "" Then GoTo 15
    PLAENTRE$ = Left$(OBX$, 48)
    CONDIPAG$ = Mid$(OBX$, 49, 48)
    ATENTX$ = Mid$(OBX$, 97, 48)
    If TRIM$(PLAENTRE$) = "" Or TRIM$(CONDIPAG$) = "" Then
      Call MENSERR(24, "Datos Incompletos")
      VDEF$ = OBX$
      GoTo 15
    End If
    '
20  NC% = Val(Label11)
    TIPODOC$ = "Presup.Reparación (Externo)"
    DESTIDOC% = NC%
    NUMEDOC$ = TRIM$(Label18)
    FECHDOC$ = TRIM$(Text9.TEXT)
    '
    CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = TRIM$(Label10)
    CODPARAM$(2) = "DES-CONJ": DOCPARAM$(2) = TRIM$(TEPRUEBA)
    CODPARAM$(3) = "REF-MAT1": DOCPARAM$(3) = TRIM$(Text1)
    CODPARAM$(4) = "MEDIDA-1": DOCPARAM$(4) = TRIM$(Text10)
    CODPARAM$(5) = "REF-MAT2": DOCPARAM$(5) = ""
    CODPARAM$(6) = "REF-MAT3": DOCPARAM$(6) = ""
    CODPARAM$(7) = "PRE-NETO": DOCPARAM$(7) = ""
    CODPARAM$(8) = "BRU-ITEM": DOCPARAM$(8) = ""
    CODPARAM$(9) = "IMP-NETO": DOCPARAM$(9) = TRIM$(TOTAL(0))
    CODPARAM$(10) = "IMP-IVA": DOCPARAM$(10) = TRIM$(TOTAL(1))
    CODPARAM$(11) = "IMP-TOTA": DOCPARAM$(11) = TRIM$(TOTAL(2))
    PORI$ = ""
    On Error Resume Next
    PORI$ = TRIM$(Str$(100 * TOTAL(1) / TOTAL(0)))
    On Error GoTo 0
    CODPARAM$(12) = "POR-IVA": DOCPARAM$(12) = TRIM$(PORI$)
    CODPARAM$(13) = "MEDIDA-3": DOCPARAM$(13) = TRIM$(TALON)
    CODPARAM$(14) = "REF-MAT2": DOCPARAM$(14) = TRIM$(PLAENTRE$)
    CODPARAM$(15) = "COND-PAG": DOCPARAM$(15) = TRIM$(CONDIPAG$)
    CODPARAM$(16) = "REF-MAT3": DOCPARAM$(16) = TRIM$(ATENTX$)
    CODPARAM$(17) = "PERC-IVA": DOCPARAM$(17) = TRIM$(TOTPER)
    CAPARDOC% = 17
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CACOLTAB1% = 3
    '
    CAITAB1% = 0
    For U& = 1 To List1.ListCount
      ZZ$ = List1.List(U& - 1)
      CAITAB1% = CAITAB1% + 1
      TETABU1$(1, CAITAB1%) = Left$(ZZ$, 16)
      CII% = CODINT%(Left$(ZZ$, 16))
      TETABU1$(2, CAITAB1%) = DESCRIT$(CII%)
      TETABU1$(3, CAITAB1%) = TRIM$(Mid$(ZZ$, 56, 8))
    Next U&
    '
    Call PRINTDOC("@" + FORIPRE$)
    '
End Sub

Sub GRAPRESU(OKX%)
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
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Then
     If TRIM$(TALON) = "" Then
          MERRO$ = "Falta Número de Talón"
          GoTo 98
        ElseIf Len(TRIM$(TALON)) > 6 Then
          MERRO$ = "Exceso Longitud Número de Talón"
          TALON = Left$(TRIM$(TALON), 6)
          GoTo 98
     End If
   End If
   '
   UREPRE& = ULTREG&("PRESREPA")
   TALOT$ = TRIM$(TALON)
   If TALOT$ <> "" Then
     For U& = UREPRE& To UREPRE& - 512 Step -1
       If U& > 0 Then
         If U& <> REG& Then
           YY$ = REGLEIDO$("PRESREPA", U&)
           If TRIM$(Mid$(YY$, 95, 6)) = TALOT$ Then
             MERRO$ = "Talón Repetido"
             GoTo 98
           End If
         End If
       End If
     Next U&
   End If
   '
   If Len(TRIM$(Text1)) > 15 Then
        MERRO$ = "Exceso Longitud Número de Serie"
        Text1 = Left$(TRIM$(Text1), 15)
        GoTo 98
   End If
   '
   If Len(TRIM$(Text10)) > 9 Then
        MERRO$ = "Exceso Longitud Referencia"
        Text10 = Left$(TRIM$(Text10), 9)
        GoTo 98
   End If
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
      MERRO$ = "Imposible Modificar - O.Venta " + TRIM$(Str$(CVS(Mid$(XX$, 125, 4))))
      GoTo 98
   End If
   '
   REFEPRE$ = "T." + TRIM$(TALON) + " - " + TRIM$(Label10) + " - " + Left$(TRIM$(Label12), 20) + " - R: " + TRIM$(Text10) + " - " + TRIM$(Text9)
   Call REPLA(XX$, REFEPRE$, 5, 60)
   Call REPLA(XX$, TRIM$(Text1), 71, 15)
   Call REPLA(XX$, TRIM$(Text10), 86, 9)
   Call REPLA(XX$, TRIM$(TALON), 95, 6)
   Call REPLA(XX$, MKS$(Val(Text3)), 101, 4)
   Call REPLA(XX$, MKS$(Val(Text2)), 105, 4)
   Call REPLA(XX$, MKS$(Val(TOTAMOBRA)), 109, 4)
   Call REPLA(XX$, MKS$(Val(SUMAREP)), 113, 4)
   Call REPLA(XX$, MKS$(Val(TOTAL(0))), 117, 4)
   DOPERA$ = "Emisión": If CVI(Mid$(XX$, 121, 2)) > 0 Then DOPERA$ = "Actualización"
   Call REPLA(XX$, MKI$(FEMI%), 121, 2)
   MODIOBSE% = 0: If TRIM$(Text28) <> TRIM$(Mid$(XX$, 129, 384)) Then MODIOBSE% = 1
   Call REPLA(XX$, Text28 + Space$(384), 129, 384)
   Call GRAREG("PRESREPA", XX$, REG&)
   Call CIERRARCH("PRESREPA")
   '
   NPLAL& = Val(Label18)
   NPLAI% = NPLAL& Mod 10000
   REBLA$ = REGBLAN$("REPUREPA")
   Call REPLA(REBLA$, MKS$(NPLAL&), 1, 4)
   Call REPLA(REBLA$, MKI$(NPLAI%), 5, 2)
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
     PRUNI = CVS(Mid$(ZZ$, 29, 4))
     XX$ = REBLA$
     CII% = CVI(Left$(ZZ$, 2))
     CANTI = CVS(Mid$(ZZ$, 5, 4))
     PRUNI = CVS(Mid$(ZZ$, 29, 4))
     '
     Call REPLA(XX$, MKI$(CII%), 7, 2)
     Call REPLA(XX$, MKS$(CANTI), 9, 4)
     Call REPLA(XX$, MKS$(PRUNI), 13, 4)
     Call REPLA(XX$, MKS$(CANTI * PRUNI), 17, 4)
     '
     KK& = KK& + 4
     If KK& <= Len(RCC$) Then
          RGG& = CVS(Mid$(RCC$, KK&, 4))
          Call GRAREG("REPUREPA", XX$, RGG&)
        Else
          Call REGAPP("REPUREPA", XX$)
     End If
   Next U&
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
   OKX% = 1
   '
   TTXYZ$ = Space$(92)
   NROREP& = Val(Label18)
   HOII% = HOY(HOS$)
   Call REPLA(TTXYZ$, HOS$, 1, 8)
   Call REPLA(TTXYZ$, USERTER$, 69, 24)
   Call REPLA(TTXYZ$, DOPERA$ + " Presupuesto de Reparación:", 11, 56)
   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   '
   TTXYZ$ = Space$(92)
   Call REPLA(TTXYZ$, Space$(56), 11, 56)
   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   Call REPLA(TTXYZ$, "Estimado de Horas: " + TRIM$(Text3), 11, 56)
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
   '
   'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
   '
'   NSERIE$ = TRIM$(Text1)
'   If NSERIE$ <> "" Then
'    CUERPO$ = "Reparación - Presupuesto de Reparación"
'    Call ANOTANUSE(NSERIE$, CUERPO$)
'   End If
   '
   Call CIERRARCH("ANOTAREP")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   Screen.MousePointer = 1
   Exit Sub
   '
98 Call CIERRARCH("PRESREPA")
   Call CIERRARCH("REPUREPA")
   Screen.MousePointer = 1
   Call MENSERR(24, MERRO$)
   '
End Sub

Private Sub TOTAMOBRA_Change()
   TOTAMO = Val(SUMAREP) + Val(TOTAMOBRA)
   TOTAL(0) = CSTRING$(MKS$(TOTAMO), 3, 11, 2, 2)
End Sub

Sub EDITREPU()
   '
20 ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=Repuestos para Reparación"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame4.Top + Frame1.Top + 110)) 'Height - ScaleHeight))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(60 + Left + Frame4.Left + Frame1.Left))
   '
   VTB% = VENTABU%("EDITREPE" + ATRI$)
   If VTB% < 0 Then
     Exit Sub
   End If
   Call PREREPUES
   '
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
Sub CALTOIVA()
   '
   Static CLIA$, RNI%, APLIVA%
   If Label11.Caption <> CLIA$ Then
     CLIA$ = Label11.Caption: NCLIE% = XVALO(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIE% > 0 Then
       PVCLI% = PIVACLI%(NCLIE%)
       If PVCLI% = 2 Then RNI% = 1
       If PVCLI% = 5 Or PVCLI% = 7 Then APLIVA% = 0
     End If
   End If
   '
   IMPIVA21 = TRIM$(FORMATNUM$(APLIVA% * XVALO(TOTAMOBRA) * (PORCEIVA(0) / 100), "F12.2"))
   IMPIVA105 = ""
   For U& = 1 To ULTREG&("!MATEROF")
     ZZ$ = REGLEIDO$("!MATEROF", U&)
     CIXXI% = CVI(Left$(ZZ$, 2))
     CANTU = CVS(Mid$(ZZ$, 5, 4))
     PRUNI = CVS(Mid$(ZZ$, 29, 4))
     '
     TOT# = ROUND(CANTU * PRUNI)
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
   '
   SUIVA# = XVALO(IMPIVA21) + XVALO(IMPIVA105)
   SUIVARNI# = SUIVA# * RNI% / 2
   TOTIVA = TRIM$(CSTRING$(MKD$(SUIVA# + SUIVARNI#), 4, 12, 2, 2))
   '
   TOTAL(1) = TOTIVA
   '
End Sub




