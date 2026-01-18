VERSION 5.00
Begin VB.Form PLAPRERE04 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilla de Presupuestación por Reparaciones"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11625
   Icon            =   "PLAPRERE04.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11625
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2325
      Left            =   -100
      TabIndex        =   14
      Top             =   7170
      Width           =   10800
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   32
         Top             =   280
         Width           =   3270
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
            Left            =   420
            TabIndex        =   33
            Top             =   315
            Width           =   2535
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
         Height          =   1020
         Left            =   3570
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   0
         Width           =   7155
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
         Y1              =   0
         Y2              =   1020
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
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
         TabIndex        =   31
         Top             =   0
         Width           =   3270
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   -100
      TabIndex        =   8
      Top             =   0
      Width           =   12000
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00004000&
         Height          =   7155
         Left            =   10815
         ScaleHeight     =   7095
         ScaleWidth      =   1845
         TabIndex        =   37
         Top             =   -105
         Width           =   1905
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   105
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   45
            Top             =   6300
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   46
               Top             =   0
               Width           =   12000
            End
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
            Left            =   105
            Picture         =   "PLAPRERE04.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   800
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
            Left            =   105
            Picture         =   "PLAPRERE04.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   5205
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
            Left            =   105
            Picture         =   "PLAPRERE04.frx":0A56
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Acceso a Base de Datos de Clientes"
            Top             =   2625
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
            Left            =   105
            Picture         =   "PLAPRERE04.frx":0E98
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Acceso a Maestro de Artículos"
            Top             =   1980
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
            Left            =   105
            Picture         =   "PLAPRERE04.frx":11A2
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Resumen de Cuenta Corriente de Clientes"
            Top             =   3270
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
            Left            =   105
            Picture         =   "PLAPRERE04.frx":14AC
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   150
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
            Left            =   105
            Picture         =   "PLAPRERE04.frx":15F6
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Ver y Reimprimir Planillas y Presupuestos"
            Top             =   4560
            Width           =   650
         End
         Begin VB.Image Image2 
            Height          =   390
            Left            =   -315
            Picture         =   "PLAPRERE04.frx":1900
            Top             =   6720
            Width           =   1515
         End
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Presup."
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
         Left            =   10830
         TabIndex        =   36
         ToolTipText     =   "Marque Casilla para Presupuesto Directo"
         Top             =   7200
         Width           =   1200
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
         Height          =   705
         Left            =   10830
         Picture         =   "PLAPRERE04.frx":3822
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Aprueba, Graba e Imprime Presupuesto"
         Top             =   7515
         Width           =   800
      End
      Begin VB.TextBox TALON 
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
         TabIndex        =   3
         Text            =   " "
         Top             =   890
         Width           =   1110
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   27
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
            Height          =   4110
            Left            =   0
            TabIndex        =   29
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
            TabIndex        =   28
            Top             =   315
            Width           =   10100
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   26
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
            TabIndex        =   6
            Text            =   " "
            Top             =   495
            Width           =   2325
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   16
         Top             =   1365
         Width           =   7680
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
            TabIndex        =   4
            Text            =   " "
            Top             =   495
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
            TabIndex        =   5
            Top             =   495
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
            Height          =   510
            Left            =   2835
            TabIndex        =   17
            Top             =   285
            Width           =   4680
         End
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
         Left            =   6615
         TabIndex        =   2
         Text            =   " "
         Top             =   890
         Width           =   1485
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
         TabIndex        =   34
         Top             =   945
         Width           =   1800
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
         TabIndex        =   30
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
         TabIndex        =   25
         Top             =   495
         Width           =   1110
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
         Left            =   5880
         TabIndex        =   24
         Top             =   490
         Width           =   2220
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
         Left            =   5880
         TabIndex        =   23
         Top             =   105
         Width           =   2220
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   885
         Width           =   750
      End
      Begin VB.Line Line7 
         X1              =   210
         X2              =   12000
         Y1              =   7000
         Y2              =   7000
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
         Left            =   5355
         TabIndex        =   15
         Top             =   525
         Width           =   1695
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   420
         Picture         =   "PLAPRERE04.frx":3E8C
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
         TabIndex        =   13
         Top             =   525
         Width           =   1800
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
         Left            =   4760
         TabIndex        =   12
         Top             =   945
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
         Left            =   5355
         TabIndex        =   10
         Top             =   150
         Width           =   960
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
         TabIndex        =   9
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
Attribute VB_Name = "PLAPRERE04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Command11_Click()
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
   Call GRAPLARE(OKX%)
   If OKX% > 0 Then
     CANCELPRINT% = 0
     If Check1.Value = 1 Then CANCELPRINT% = 1
     Call PRIPLAREPA
     '
     If Check1.Value < 1 Then
       ANOTAREP04.Label7 = Label18
       ANOTAREP04.Show 1
     End If
     '
     If Check1.Value = 1 Then
      '\\\OT-08-0868-OD-22/09/08///
       PRESURE04.Label18 = Label18
       PRESURE04.Show 1
      '\\\OT-08-0868-OD-FIN///
     End If
     '
     Call BLANFORMU
     Command14.Enabled = True
     Label11 = ""
     '
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
    TIDOCUM$ = "Planilla de Reparación " + NUDOCU$ + " "
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
    '
    GoTo 3
    '
99  Command15.Enabled = True
    '
End Sub


Private Sub Command18_Click()
   ARCHIPRE$ = "MASTER"
   Call SELECHO(ARCHIPRE$)
   Text13 = VALACT1$(ARCHIPRE$)
   Text13.Refresh
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

Sub Command8_Click()
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
    Label10 = USERNAME$
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$
    Label18 = TRIM$(Str$(NUPRESUP&))
    '
End Sub
'

Private Sub Form_Resize()
   If WindowState = 1 Then
      Call CIERRARCH("*.*")
      'Call BAJALDISCO
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
   Call BLANFORMU
   '
   NC% = Val(Label11)
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

Private Sub Label8_Click()
   Text13.SetFocus
End Sub

Private Sub List1_DBLClick()
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
           STDIS = STOCKDIS(CIJ%(I1%))
           Call REPLA(ZZ$, MKS$(STDIS), 21, 4)
           FALTA = NEBRU - STDIS: If FALTA < 0 Then FALTA = 0
           Call REPLA(ZZ$, MKS$(FALTA), 25, 4)
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
   On Error Resume Next
   List1.ListIndex = REACTI& - 1
   On Error GoTo 0
   Screen.MousePointer = 1
   '
End Sub

Private Sub TEPRUEBA_Change()
   DESCRI$ = TEPRUEBA.Caption
   If TEPRUEBA.Width < Label8.Width Then
     DESCRI$ = Chr$(10) + Chr$(13) + DESCRI$
   End If
   Label8.Caption = DESCRI$
End Sub

Private Sub Text13_Change()
   If CODINT%(Text13) > 0 Then
      CI% = CODINT%(Text13)
      TEPRUEBA.Caption = DESCRIT0$(CI%)
      Call CAREPUES
      Call PREREPUES
    End If
End Sub

Private Sub Text13_DblClick()
   ARCHIPRE$ = "MASTER"
   Call SELECHO(ARCHIPRE$)
   Text13 = VALACT1$(ARCHIPRE$)
   Text13.Refresh
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
   CODPRO$ = Text13.TEXT
   Call LEEEXPLOMUL(CODPRO$, 1, 1)
   If CAIT% < 1 Then             ' producto sin estructura
     Call COMUNI("No hay Fórmula de Consumo Standard\para este Producto (" + TRIM$(CODPRO$) + ").")
   End If

30 Screen.MousePointer = 11
   For I1% = 1 To CAIT%
     If CIJ%(I1%) > 0 Then
       If CIJ%(I1%) <= NUMAS% Then
         If Abs(USOI(I1%)) > 0 Then
           '
           ZZ$ = REGBLAN$("MATEROF")
           Call REPLA(ZZ$, MKI$(CIJ%(I1%)), 1, 2)
           Call REPLA(ZZ$, UNIMED$(CIJ%(I1%)), 3, 2)
           NEBRU = USOI(I1%)
           Call REPLA(ZZ$, MKS$(NEBRU), 5, 4)
           Call REPLA(ZZ$, MKD$(CDbl(USOI(I1%))), 9, 8)
           Call REPLA(ZZ$, MKS$(0), 17, 4)
           STDIS = STOCKDIS(CIJ%(I1%))
           Call REPLA(ZZ$, MKS$(STDIS), 21, 4)
           FALTA = NEBRU - STDIS: If FALTA < 0 Then FALTA = 0
           Call REPLA(ZZ$, MKS$(FALTA), 25, 4)
           Call REGAPP("!MATEROF", ZZ$)
           '
         End If
       End If
     End If
   Next I1%
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
   SUMAREP = ""
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

Sub PRIPLAREPA()
    '
    FORIPRE$ = TRIM$(CONTROL$("FOREPAR", "DOCUREP0"))
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
    TIPODOC$ = "Planilla de Reparación"
    DESTIDOC% = NC%
    NUMEDOC$ = TRIM$(Label18)
    FECHDOC$ = TRIM$(Text9.TEXT)
    '
    CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = TRIM$(Text13.TEXT)
    CODPARAM$(2) = "DES-CONJ": DOCPARAM$(2) = TRIM$(TEPRUEBA)
    CODPARAM$(3) = "REF-MAT1": DOCPARAM$(3) = TRIM$(Text1)
    CODPARAM$(4) = "MEDIDA-1": DOCPARAM$(4) = TRIM$(Text10)
    CODPARAM$(5) = "MEDIDA-3": DOCPARAM$(5) = TRIM$(TALON)
    CAPARDOC% = 5
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CODTABU1$(4) = "PRE-LIST"
    CODTABU1$(5) = "NET-ITEM"
    CACOLTAB1% = 5
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
      ITOX$ = Mid$(ZZ$, 25, 4)
      TETABU1$(4, CAITAB1%) = CSTRING$(ITOX$, 4, 8, 1, 2)
      TETABU1$(5, CAITAB1%) = ".............................."
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

Sub GRAPLARE(OKX%)
   '
   OKX% = 0
   Call BLOQARCH("PRESREPA")
   Call BLOQARCH("REPUREPA")
   '
   NPP& = NUPRESUP&
   If NPP& > Val(Label18) Then
      Call CIERRARCH("PRESREPA")
      Call COMUNI("Número Ocupado")
      Label18 = TRIM$(Str$(NPP&))
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
   If TRIM$(TALON) = "" Then
        Call COMUNI("ATENCION: Falta Número de Talón")
        GoTo 10
      ElseIf Len(TRIM$(TALON)) > 6 Then
        MERRO$ = "Exceso Longitud Número de Talón"
        TALON = Left$(TRIM$(TALON), 6)
        GoTo 98
   End If
   '
   UREPRE& = ULTREG&("PRESREPA")
   TALOT$ = TRIM$(TALON)
   For U& = UREPRE& To UREPRE& - 512 Step -1
     If U& > 0 Then
       YY$ = REGLEIDO$("PRESREPA", U&)
       If TRIM$(Mid$(YY$, 95, 6)) = TALOT$ Then
         MERRO$ = "Talón Repetido"
         GoTo 98
       End If
     End If
   Next U&
   '
10 If Len(TRIM$(Text1)) > 15 Then
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
   CIIX% = CODINT%(Text13.TEXT)
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
   TTXYZ$ = Space$(92)
   NROREP& = Val(Label18)
   HOII% = HOY(HOS$)
   Call REPLA(TTXYZ$, HOS$, 1, 8)
   Call REPLA(TTXYZ$, USERTER$, 69, 24)
   Call REPLA(TTXYZ$, "Planilla de Presupuestación de Reparación:", 11, 56)
   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   '
   TTXYZ$ = Space$(92)
   Call REPLA(TTXYZ$, Space$(56), 11, 56)
   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   Call REPLA(TTXYZ$, "Cliente: " + Label12, 11, 56)
   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   Call REPLA(TTXYZ$, "    Tel: " + TRIM$(Label16), 11, 56)
   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   Call REPLA(TTXYZ$, "    Fax: " + TRIM$(Label17), 11, 56)
   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   '
   Call FRATEXTO(Text28, 46)
   For KKI% = 1 To CARETEX%
     TTXX$ = Space$(92)
     Call REPLA(TTXX$, RETEX$(KKI%), 10, 46)
     If KKI% = 1 Then Call REPLA(TTXX$, " Observ: ", 1, 9)
     Call REGAPP("ANOTAREP", Left$(MKS$(NROREP&) + Space$(10) + TTXX$, 96))
   Next KKI%
   '
   Call REGAPP("ANOTAREP", MKS$(NROREP&) + String$(88, "="))
   Call CIERRARCH("ANOTAREP")
   '
   REFEPRE$ = "T." + TRIM$(TALON) + " - " + TRIM$(Text13.TEXT) + " - " + Left$(TRIM$(Label12), 20) + " - R: " + TRIM$(Text10) + " - " + TRIM$(Text9)
   XX$ = REGBLAN$("PRESREPA")
   Call REPLA(XX$, MKS$(Val(Label18)), 1, 4)
   Call REPLA(XX$, REFEPRE$, 5, 60)
   Call REPLA(XX$, MKI$(FEMI%), 65, 2)
   Call REPLA(XX$, MKI$(NCLIE%), 67, 2)
   Call REPLA(XX$, MKI$(CIIX%), 69, 2)
   Call REPLA(XX$, TRIM$(Text1), 71, 15)
   Call REPLA(XX$, TRIM$(Text10), 86, 9)
   Call REPLA(XX$, TRIM$(TALON), 95, 6)
   Call REPLA(XX$, String$(28, 0), 101, 28)
   Call REPLA(XX$, Text28 + Space$(384), 129, 384)
   Call REGAPP("PRESREPA", XX$)
   Call CIERRARCH("PRESREPA")
   '
   NPLAL& = Val(Label18)
   NPLAI% = NPLAL& Mod 10000
   REBLA$ = REGBLAN$("REPUREPA")
   Call REPLA(REBLA$, MKS$(NPLAL&), 1, 4)
   Call REPLA(REBLA$, MKI$(NPLAI%), 5, 2)
   '
   For U& = 1 To ULTREG&("!MATEROF")
     ZZ$ = REGLEIDO$("!MATEROF", U&)
     XX$ = REBLA$
     CII% = CVI(Left$(ZZ$, 2))
     CANTI = CVS(Mid$(ZZ$, 5, 4))
     Call REPLA(XX$, MKI$(CII%), 7, 2)
     Call REPLA(XX$, MKS$(CANTI), 9, 4)
     Call REGAPP("REPUREPA", XX$)
   Next U&
   Call CIERRARCH("REPUREPA")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
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
   TALON = ""
   Text1 = ""
   Text10 = ""
   Text13 = ""
   Text28 = ""
   List1.Clear
   '
   Check1.Value = 0
   '
End Sub
