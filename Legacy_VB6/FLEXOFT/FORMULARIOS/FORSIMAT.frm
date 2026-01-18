VERSION 5.00
Begin VB.Form FORSIMAT 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud Interna de Materiales"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   Icon            =   "FORSIMAT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1800
      Left            =   0
      TabIndex        =   19
      Top             =   6510
      Width           =   10800
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   1590
         Left            =   3045
         TabIndex        =   35
         Top             =   105
         Visible         =   0   'False
         Width           =   2220
         Begin VB.TextBox TOTIVA 
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
            Index           =   0
            Left            =   210
            TabIndex        =   38
            Text            =   " "
            Top             =   210
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox TOTIVA 
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
            Left            =   210
            TabIndex        =   37
            Text            =   " "
            Top             =   525
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox TOTIVA 
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
            Index           =   2
            Left            =   210
            TabIndex        =   36
            Text            =   " "
            Top             =   840
            Visible         =   0   'False
            Width           =   870
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
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
         Height          =   1650
         Left            =   210
         TabIndex        =   31
         Top             =   105
         Width           =   10530
         Begin VB.TextBox Text44 
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
            Left            =   4620
            TabIndex        =   43
            Top             =   1230
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1170
            Left            =   1155
            MaxLength       =   96
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   315
            Width           =   9220
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Observac:"
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
            TabIndex        =   44
            Top             =   315
            Width           =   4110
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Items:"
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
            Left            =   210
            TabIndex        =   40
            Top             =   945
            Width           =   645
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
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
            Height          =   270
            Left            =   420
            TabIndex        =   39
            Top             =   1200
            Width           =   540
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   12000
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
         Left            =   8715
         MaxLength       =   24
         TabIndex        =   46
         Top             =   1400
         Width           =   2025
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   10920
         ScaleHeight     =   75
         ScaleWidth      =   795
         TabIndex        =   41
         Top             =   6405
         Width           =   855
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   735
         Top             =   0
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   6180
         Left            =   10920
         TabIndex        =   23
         Top             =   105
         Width           =   850
         Begin VB.CommandButton Command22 
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
            Height          =   770
            Left            =   0
            Picture         =   "FORSIMAT.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Anotador de SIM"
            Top             =   3860
            Width           =   855
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
            Height          =   770
            Left            =   0
            Picture         =   "FORSIMAT.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Obras"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORSIMAT.frx":0896
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Base de Datos de Obras"
            Top             =   3090
            Width           =   855
         End
         Begin VB.CommandButton Command5 
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
            Height          =   770
            Left            =   0
            Picture         =   "FORSIMAT.frx":0BA0
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Configuración de Funcionamiento de SIM"
            Top             =   5410
            Width           =   855
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Articulos"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORSIMAT.frx":0FE2
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Acceso a Maestro de Artículos"
            Top             =   1540
            Width           =   855
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Clientes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORSIMAT.frx":12EC
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Acceso a Base de Datos de Clientes"
            Top             =   2315
            Width           =   855
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Opción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORSIMAT.frx":172E
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Opciones Varias - Consultas"
            Top             =   4642
            Width           =   855
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
            Height          =   770
            Left            =   0
            Picture         =   "FORSIMAT.frx":1878
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   770
            Width           =   855
         End
         Begin VB.CommandButton Command12 
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
            Height          =   770
            Left            =   0
            Picture         =   "FORSIMAT.frx":1B82
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Consultas y Re-Impresión"
            Top             =   3860
            Width           =   855
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
         Height          =   1520
         Left            =   10920
         Picture         =   "FORSIMAT.frx":1E8C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Aprueba, Graba e Imprime SIM"
         Top             =   6750
         Width           =   855
      End
      Begin VB.CommandButton Command10 
         Height          =   260
         Left            =   420
         TabIndex        =   22
         Top             =   6855
         Width           =   2825
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   210
         ScaleHeight     =   375
         ScaleWidth      =   10455
         TabIndex        =   20
         Top             =   1785
         Width           =   10515
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Código          Descripción                       UM   Cantidad   St.Disp.  F.Solic."
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
            Left            =   30
            TabIndex        =   21
            Top             =   105
            Width           =   10515
         End
      End
      Begin VB.TextBox Text12 
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
         Left            =   1575
         MaxLength       =   40
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1280
         Width           =   4065
      End
      Begin VB.TextBox Text11 
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
         TabIndex        =   9
         Text            =   " "
         Top             =   455
         Width           =   1120
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
         Left            =   8400
         MaxLength       =   24
         TabIndex        =   10
         Top             =   945
         Width           =   2340
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
         TabIndex        =   7
         TabStop         =   0   'False
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
         TabIndex        =   8
         Top             =   105
         Width           =   180
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
         Left            =   2625
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   890
         Width           =   3015
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
         Left            =   1575
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   890
         Width           =   960
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
         Left            =   1575
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   490
         Width           =   4065
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
         Left            =   1575
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   105
         Width           =   4065
      End
      Begin VB.CommandButton Command1 
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
         Left            =   1260
         TabIndex        =   1
         Top             =   1280
         Width           =   180
      End
      Begin VB.TextBox Text1 
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
         MaxLength       =   8
         TabIndex        =   0
         Top             =   1260
         Width           =   1065
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
         Height          =   4335
         Left            =   210
         TabIndex        =   32
         Top             =   2205
         Width           =   10515
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
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
         Left            =   6800
         TabIndex        =   47
         Top             =   1450
         Width           =   1800
      End
      Begin VB.Label NCLIEN 
         Caption         =   " "
         Height          =   330
         Left            =   840
         TabIndex        =   45
         Top             =   210
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   6195
         Picture         =   "FORSIMAT.frx":2196
         ToolTipText     =   "Gestión de Pedidos y Ordenes de Compra a Terceros."
         Top             =   525
         Width           =   480
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "SIM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   855
         Left            =   150
         TabIndex        =   33
         Top             =   180
         Width           =   1380
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   420
         Shape           =   4  'Rounded Rectangle
         Top             =   465
         Width           =   540
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000C000&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   390
         Left            =   315
         Shape           =   4  'Rounded Rectangle
         Top             =   390
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   210
         Shape           =   4  'Rounded Rectangle
         Top             =   315
         Width           =   960
      End
      Begin VB.Line Line7 
         X1              =   210
         X2              =   10710
         Y1              =   6550
         Y2              =   6550
      End
      Begin VB.Line Line8 
         X1              =   10730
         X2              =   10730
         Y1              =   1785
         Y2              =   6550
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SIM Nro:"
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
         Left            =   7755
         TabIndex        =   18
         Top             =   485
         Width           =   1800
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. / Parte:"
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
         Left            =   6345
         TabIndex        =   17
         Top             =   990
         Width           =   2010
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
         Left            =   8820
         TabIndex        =   16
         Top             =   150
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Obra:"
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
         Left            =   250
         TabIndex        =   15
         Top             =   945
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
Attribute VB_Name = "FORSIMAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHIPRE$

Private Sub command1_Click()
   OBRA$ = REGSEL$("INDIOBR")
   Text1.TEXT = TRIM$(Mid$(OBRA$, 1, 8))
   '
   NCLIE% = CVI(Mid$(OBRA$, 49, 2))
   If NCLIE% > 0 Then
     Text2.TEXT = RASOCLI$(NCLIE%)
     Text3.TEXT = DOMICLI$(NCLIE%)
     Text4.TEXT = CPOSCLI$(NCLIE%)
     Text5.TEXT = LOCACLI$(NCLIE%)
     Text12.TEXT = TRIM$(Mid$(OBRA$, 9, 40))
     NCLIEN.Caption = CStr(NCLIE%)
     On Error Resume Next
     Command6.SetFocus
     On Error GoTo 0
   End If
   List1.Clear
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   AMACLI00.Show 1
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    '
    Command12.Enabled = False
    'Screen.MousePointer = 11
    '
    'JJ& = 0
    'For U& = 1 To ULTREG&("PEDENCA")
2   '  XX$ = REGLEIDO$("PEDENCA", U&)
    '  If CVS(Left$(XX$, 4)) > 0 Then
    '    NPX& = CVS(Left$(XX$, 4))
    '    NCLIE% = CVI(Mid$(XX$, 55, 2))
    '    FEX = CVI(Mid$(XX$, 53, 2))
    '    ANULX$ = "": If Mid$(XX$, 267, 1) = Chr$(9) Then ANULX$ = "(Anul) "
    '    YY$ = MKS$(NPX&) + ANULX$ + TRIM$(Mid$(XX$, 5, 48))
    '    JJ& = JJ& + 1
    '    Call GRAREG("INDIPEDI", YY$, JJ&)
    '  End If
    'Next U&
    'Call SETULTREG("INDIPEDI", JJ&)
    'Screen.MousePointer = 1
    '
'3   Call SELECHO("INDIPEDI/Indice General de Pedidos (Pendientes-Cumplidos-Anulados)")
    'NPX& = Val(VALACT1$("INDIPEDI"))
3   Call SELECHO("PEDENCA/Indice General de Pedidos (Pendientes-Cumplidos-Anulados)")
    NPX& = Val(VALACT1$("PEDENCA"))
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Pedido de Cliente"
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
    '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
    On Error Resume Next
    OPPEDIDO.Command46.SetFocus
    On Error GoTo 0
    OPPEDIDO.Show 1
    GoTo 99
    '
5   For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    GoTo 99
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
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
10  PPXX% = InStr(TPSP$, Chr$(255))
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
    EDITFORM.Command5.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 3
    '
99  Command12.Enabled = True
    '
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   OPSIMAT.Show 1
   If MONEMI% < 1 Or MONEMI% > 2 Then MONEMI% = 1
   If MONEMI% = 2 Then
       MONEMIS = "Dólares"
     Else
       MONEMIS = "Pesos"
   End If
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   '
   OKX% = 0
   'If ULTREG&("!MATEROF") > 0 Then OKX% = 1
   '
   Call GRAOBRA(OKX%)
   If OKX% > 0 Then
     'Call GRAOBRA(OKX%)
     Call PRIPEDIDO
     Call BLANFORMU
     Text1.TEXT = ""
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
   End If
   'Call BLANARCH("!MATEROF")
99 Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
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
   Call MENSERR(24, "No Disponible por el Momento")
End Sub

Private Sub Command18_Click()
   MARSELPRESU.Visible = False
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Command19_Click()
   '
   If Left$(UCase$(Label11), 6) = "REPARA" Then GoTo 50
   '
   UREPRE& = ULTREG&("PRESUPU")
   HOII% = HOY(HOS$)
   JJ& = 0: ATX% = 0: COEFCON = 1
   For KKI% = 1 To List2.ListCount
     If List2.Selected(KKI% - 1) = True Then
       REPRESU& = Val(Mid$(List2.List(KKI% - 1), 121, 8))
       If REPRESU& > 0 Then
         If REPRESU& <= UREPRE& Then
           NPRESX& = CVS(Left$(REGLEIDO$("PRESUPU", REPRESU&), 4))
           For REPRES1& = REPRESU& To UREPRE&
             XX$ = REGLEIDO$("PRESUPU", REPRES1&)
             If CVS(Left$(XX$, 4)) <> NPRESX& Then GoTo 19
             If CVI(Mid$(XX$, 113, 2)) > 0 Then GoTo 19 ' FECHA O.VENTA
             '
             ZZ$ = REGBLAN$("!CARDETPE")
             CIXI% = CVI(Mid$(XX$, 69, 2))
             Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
             PRELI = CVS(Mid(XX$, 71, 4))
             DESCUTEX$ = TRIM$(Mid$(XX$, 75, 14))
             CPPP = (100 - COEFDESCU(DESCUTEX$)) / 100
             If CPPP < 1 Then ATX% = 1
             '
             MONEMA% = Asc(Mid$(XX$, 68, 1))
             If JJ& < 1 Then
               MONEMI% = Asc(Mid$(XX$, 68, 1))
               If MONEMI% < 1 Then MONEMI% = 1
               If MONEMI% = 2 Then MONEMIS = "Dólares" Else MONEMIS = "Pesos"
             End If
             COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
             CAPRESU& = CVS(Mid$(XX$, 101, 4))
             If CAPRESU& < 1 Then CAPRESU& = 1
             '
             Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 39, 4)
             Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 51, 4)
             Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 65, 4)
             Call REPLA(ZZ$, MKS$(CAPRESU&), 47, 4)   ' CANTIDAD
             DESCUTEX$ = TRIM$(Mid$(XX$, 75, 14))
             'Call REPLA(ZZ$, MKS$(COEFDESCU(DESCUTEX$)), 69, 4)
             Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 73, 4)
             Call REPLA(ZZ$, MKI$(HOII%), 63, 2)
             Call REPLA(ZZ$, MKS$(REPRES1&), 125, 4)
             JJ& = JJ& + 1
             Call GRAREG("!CARDETPE", ZZ$, JJ&)
           Next REPRES1&
         End If
       End If
     End If
19 Next KKI%
   Call SETULTREG("!CARDETPE", JJ&)
   Call CIERRARCH("!CARDETPE")
   GoTo 90
   '
50 UREPRE& = ULTREG&("PRESREPA")
   HOII% = HOY(HOS$)
   'TEXT21(0) = "0"
   JJ& = 0: ATX% = 0
   For KKI% = 1 To List2.ListCount
     If List2.Selected(KKI% - 1) = True Then
       REPRESU& = Val(Mid$(List2.List(KKI% - 1), 121, 8))
       If REPRESU& > 0 Then
         If REPRESU& <= UREPRE& Then
           XX$ = REGLEIDO$("PRESREPA", REPRESU&)
           ZZ$ = REGBLAN$("!CARDETPE")
           CIXI% = CVI(Mid$(XX$, 69, 2))
           Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
           PRELI = CVS(Mid(XX$, 117, 4))
           'DESCUTEX$ = TRIM$(Mid$(XX$, 75, 14))
           CPPP = 1  ' (100 - COEFDESCU(DESCUTEX$)) / 100
           'If CPPP < 1 Then ATX% = 1
           '
           MONEMA% = 1   ' REPARACIONES SIEMPRE EN PESOS
           MONEMI% = 1
           MONEMIS = "Pesos"
           COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
           '
           Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 39, 4)
           Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 51, 4)
           Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 65, 4)
           Call REPLA(ZZ$, MKS$(1), 47, 4)   ' CANTIDAD
           DESCUTEX$ = TRIM$(Mid$(XX$, 75, 14))
           'Call REPLA(ZZ$, MKS$(COEFDESCU(DESCUTEX$)), 69, 4)
           Call REPLA(ZZ$, MKS$(PRELI * CPPP * COEFCON), 73, 4)
           Call REPLA(ZZ$, MKI$(HOII%), 63, 2)
           Call REPLA(ZZ$, MKS$((-1) * REPRESU&), 125, 4)
           JJ& = JJ& + 1
           Call GRAREG("!CARDETPE", ZZ$, JJ&)
         End If
       End If
     End If
   Next KKI%
   Call SETULTREG("!CARDETPE", JJ&)
   Call CIERRARCH("!CARDETPE")
   '
90 MARSELPRESU.Visible = False
   '
   If ATX% > 0 Then
      PORDESCU.TEXT = ""
      Call COMUNI("ATENCION: Descuento Aplicado según Presupuesto Previo !!!")
   End If
   ' Call CARDETPEDI(1)
   '
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command2_LostFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub


Private Sub Command20_Click()
   Static ARGU$
   Command20.Enabled = False
10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda Cliente por Texto", ARGU$)))
   If ARGU$ <> "" Then
     Screen.MousePointer = 11
     PUNCLI$ = ""
     Call COPYSTRU("DEUDOR1", "SELECLI")
     FIN& = ULTREG&("DEUDOR2")
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = UCase$(REGLEIDO$("DEUDOR2", U&))
       If InStr(X$, ARGU$) > 0 Then
         X$ = REGLEIDO$("DEUDOR1", U&)
         PUNCLI$ = PUNCLI$ + Left$(X$, 2)
         JJ& = JJ& + 1
         Call GRAREG("!SELECLI", X$, JJ&)
       End If
     Next U&
     '
     FIN& = ULTREG&("BADACLIE")
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = UCase$(REGLEIDO$("BADACLIE", U&))
       If InStr(X$, ARGU$) > 0 Then
         NCX$ = Left$(REGLEIDO$("BADACLIE", U&), 2)
         For KKI% = 1 To Len(PUNCLI$) Step 2
           If Mid$(PUNCLI$, KKI%, 2) = NCX$ Then GoTo 19 ' ya esta
         Next KKI%
         UK& = REGBUS&("DEUDOR1", 1, NCX$)
         If UK& > 0 Then
           X$ = REGLEIDO$("DEUDOR1", UK&)
           JJ& = JJ& + 1
           Call GRAREG("!SELECLI", X$, JJ&)
         End If
       End If
19   Next U&
     '
     Call SETULTREG("!SELECLI", JJ&)
     '
     Screen.MousePointer = 1
     If JJ& < 1 Then
         Call MENSERR(24, "No se Encontraron Coincidencias")
         GoTo 10
       ElseIf JJ& = 1 Then
         Text1 = TRIM$(Str$(CVI(Left$(REGLEIDO$("!SELECLI", 1), 2))))
         Call CARDACLI
       Else
         Call FRESHECHO("!SELECLI")
         Call SELECHO("!SELECLI")
         NCLIX$ = VALACT1$("!SELECLI")
         If NCLIX$ <> "" Then
           Text1 = NCLIX$
           Call CARDACLI
         End If
     End If
   End If
   Command20.Enabled = True
End Sub

Private Sub Command22_Click()
    Command22.Enabled = False
    Screen.MousePointer = 11
    '
    'JJ& = 0
    'For U& = 1 To ULTREG&("PEDENCA")
2   '  XX$ = REGLEIDO$("PEDENCA", U&)
    '  If CVS(Left$(XX$, 4)) > 0 Then
    '    NPX& = CVS(Left$(XX$, 4))
    '    NCLIE% = CVI(Mid$(XX$, 55, 2))
    '    FEX = CVI(Mid$(XX$, 53, 2))
    '    ANULX$ = "": If Mid$(XX$, 267, 1) = Chr$(9) Then ANULX$ = "(Anul) "
    '    YY$ = MKS$(NPX&) + ANULX$ + TRIM$(Mid$(XX$, 5, 48))
    '    JJ& = JJ& + 1
    '    Call GRAREG("INDIPEDI", YY$, JJ&)
    '  End If
    'Next U&
    'Call SETULTREG("INDIPEDI", JJ&)
    Screen.MousePointer = 1
    '
    CONSIMAT.Show 1
    '
    Command22.Enabled = True
End Sub

Private Sub Command23_Click()
For U& = 1 To ULTREG&("PEDDETA")
XX$ = REGLEIDO$("PEDDETA", U&)
YY$ = Mid$(XX$, 113, 16)
If YY$ <> String$(16, 32) Then
TTXX$ = "": For KI% = 1 To 16: TTXX$ = TTXX$ + Str$(Asc(Mid$(YY$, KI%, 1))): Next KI%
MsgBox TTXX$
End If
Next U&
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("PRESUP01")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call FINAL("*INDIOBR")
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call CIERRARCH("*.*")
    If (USNBR% Mod 100) <> 1 Then
      Call COMUNI("La Opción Elegida Requiere\Privilegios de Supervisor.")
      Exit Sub
    End If
    CFGSIMAT.Show 1
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   MARCONOTA.Visible = False
   Call SELPRESUPEN
End Sub
'
Sub SELPRESUPEN()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   '
   UREPRE& = ULTREG&("PRESUPU")
   Screen.MousePointer = 11
   List2.Clear
   NCLIE% = Val(Text1.TEXT)
   If UCase$(Left$(Label11, 6)) = "REPARA" Then GoTo 50
   '
10 RFF$ = RECFILT$("PRESUPU", 3, MKI$(NCLIE%))
   For KKI% = 1 To Len(RFF$) Step 4
      REPRESU& = CVS(Mid$(RFF$, KKI%, 4))
      If REPRESU& > 0 Then
        If REPRESU& <= UREPRE& Then
12        XX$ = REGLEIDO$("PRESUPU", REPRESU&)
          If CVI(Mid$(XX$, 7, 2)) <> NCLIE% Then
            XX$ = REGLEIDO$("PRESUPU", 1)
            Call GRAREG("PRESUPU", XX$, 1)
            Call CIERRARCH("PRESUPU")
            GoTo 10
          End If
          If CVI(Mid$(XX$, 113, 2)) < 1 Then
            TTXX$ = "Presup.No: " + TRIM$(CSTRING$(Left$(XX$, 4), 3, 8, 0, 2))
            FEX = CVI(Mid$(XX$, 5, 2))
            TTXX$ = TTXX$ + " - " + FECHATEX$(FEX)
            If TRIM$(Mid$(XX$, 9, 16)) <> "" Then
              TTXX$ = TTXX$ + " - Ref: " + TRIM$(Mid$(XX$, 9, 16))
            End If
            CIXI% = CVI(Mid$(XX$, 69, 2))
            '
            If CVS(Left$(XX$, 4)) = NPX& Then GoTo 19
            '
            NPX& = CVS(Left$(XX$, 4))
            CAPRESU& = CVS(Mid$(XX$, 101, 4))
            If CAPRESU& < 1 Then CAPRESU& = 1
            Call RECODOSI(CIXI%, CIXJ%, NPX&)
            If CIXJ% <> CIXI% Then
              Call REPLA(XX$, MKI$(CIXJ%), 69, 2)
              Call GRAREG("PRESUPU", XX$, REPRESU&)
              GoTo 12
            End If
            '
            If CVS(Left$(REGLEIDO$("PRESUPU", REPRESU& + 1), 4)) = NPX& Then
                TTXX$ = TTXX$ + " - " + "VARIOS"
               Else
                TTXX$ = TTXX$ + " - " + TRIM$(CODEXT$(CIXI%)) + " (" + TRIM$(Str$(CAPRESU&)) + ") " + DESCRIT0$(CIXI%)
            End If
            TTXX$ = AJUSTI$(TTXX$, 128)
            Call REPLA(TTXX$, Str$(REPRESU&), 121, 8)
            List2.AddItem TTXX$
          End If
        End If
      End If
19 Next KKI%
   GoTo 90
   '
50 UREPRE& = ULTREG&("PRESREPA")
   Screen.MousePointer = 11
   RFF$ = RECFILT$("PRESREPA", 4, MKI$(NCLIE%))
   For KKI% = 1 To Len(RFF$) Step 4
      REPRESU& = CVS(Mid$(RFF$, KKI%, 4))
      If REPRESU& > 0 Then
        If REPRESU& <= UREPRE& Then
          XX$ = REGLEIDO$("PRESREPA", REPRESU&)
          If CVI(Mid$(XX$, 121, 2)) > 0 Then
            If Abs(CVS(Mid$(XX$, 125, 4))) < 0.5 Then
              TTXX$ = "Reparación No: " + TRIM$(CSTRING$(Left$(XX$, 4), 3, 8, 0, 2))
              TTXX$ = TTXX$ + " - " + Mid$(XX$, 5, 60)
              TTXX$ = AJUSTI$(TTXX$, 128)
              Call REPLA(TTXX$, Str$(REPRESU&), 121, 8)
              List2.AddItem TTXX$
            End If
          End If
        End If
      End If
   Next KKI%
   '
90 Screen.MousePointer = 1
   '
   If List2.ListCount < 1 Then
      Call SETULTREG("!CARDETPE", 0)
      Call CIERRARCH("!CARDETPE")
      'Call CARDETPEDI(1)
      Exit Sub
   End If
   '
   MARSELPRESU.Visible = True
   Command19.SetFocus
   '
End Sub
'
Private Sub Command7_Click()
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    'AMASTO00.Show 1
    ALTAPORFAM.Show 1
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text9.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text9.TEXT)
End Sub

Private Sub Eco21_GotFocus(Index As Integer)
   On Error Resume Next
   'TEXT21(Index).SetFocus
   On Error GoTo 0
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  While WindowState <> 0
    WindowState = 0
    OPENFORMS = DoEvents
    AppActivate Caption
    SendKeys "%" & Chr$(32) & Chr$(13), True
  Wend
End Sub


Private Sub Form_Load()
    '
    OPENFORMS = DoEvents
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    
    Call GRATITFOR(Caption)
    '
    If IDENTER$ <> "" Then
      If CONTROL$("CAPRINT", IDENTER$) = "SI" Then
        Timer2.Enabled = True
      End If
    End If
    '
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$
    Text11.TEXT = TRIM$(Str$(NUOBRA&))
    Text6.TEXT = USERNAME$
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
   Call LIBARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Image1_DblClick()
   Static CTXX%
   '
   If ULTREG&("DAPIBCLI") < 1 Then
     Call MENSERR(24, "Imposible desde esta Aplicación")
     Exit Sub
   End If
   '
   NCLIE% = Val(TRIM$(Text1.TEXT))
   If NCLIE% > 0 Then
     If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
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
         Call REPLA(RECORD$, MKI$(HORANUM%(Time$)), 63, 2)
         Call FILTERPUTRECORD("DAPIBCLI", 1, MKI$(NCLIE%), RECORD$)
         Call CIERRARCH("DAPIBCLI")
         '
         Call CALPERIB
         '
       End If
     End If
   End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If ECOPRECIO >= 0 Then
     Frame4.Height = 3120
   End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Frame4.Height = 1545
End Sub



Private Sub Image2_DblClick()
   Call CIERRARCH("*.*")
   If EXISTE%("ORCOMP03.EXE") > 0 Then
     Call CONECRUN("ORCOMP03", "Gestión de Pedidos y Ordenes de Compra a Terceros")
   End If
End Sub

Private Sub Label11_DblClick()
   If TRIM$(UCase$(Label11)) = "VENTA" Then
       Label11.ForeColor = RGB(255, 0, 0)
       Label11 = "Reparación"
       Label11.ToolTipText = "Doble-Click para 'Venta'"
     Else
       Label11.ForeColor = RGB(0, 0, 255)
       Label11 = "Venta"
       Label11.ToolTipText = "Doble-Click para 'Reparación'"
   End If
End Sub


Private Sub Label8_Click()
   On Error Resume Next
   Text13.SetFocus
   On Error GoTo 0
End Sub

Private Sub Label22_Click()
   If VALICLIE% = 1 Then
     Call SETULTREG("!CARDETPE", 0)
     Call CARDETOBRA(2)
   End If
End Sub

Private Sub MONEMIS_Change()
   MONEMA% = MONEMI%
   If Left$(MONEMIS, 1) = "D" Then
        MONEMI% = 2
        MONEMIS.ForeColor = RGB(150, 0, 0)
        MONEMIS.ToolTipText = "Doble-Click para 'Pesos'"
        'Label12.Visible = True
        If TICAMBIS.Visible = True Then TICAMBIS.Enabled = True
        TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
      Else
        MONEMI% = 1
        MONEMIS.ForeColor = RGB(0, 100, 0)
        MONEMIS.ToolTipText = "Doble-Click para 'Dólares'"
        'Label12.Visible = False
        If TICAMBIS.Visible = True Then TICAMBIS.Enabled = False
   End If
   '
   COEFCON = TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
   FIN& = ULTREG&("!CARDETPE")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!CARDETPE", U&)
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
     Call GRAREG("!CARDETPE", X$, U&)
   Next U&
   '
   Call CARLISPE
   Call CIERRARCH("!CARDETPE")
   Call CALTOPED
   '
End Sub

Private Sub MONEMIS_DblClick()
   If TRIM$(UCase$(MONEMIS)) = "PESOS" Then
       MONEMIS = "Dólares"
     Else
       MONEMIS = "Pesos"
   End If
End Sub

Private Sub PORDESCU_Change()
   Call CALTOPED
End Sub

Private Sub PORDESCU_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
End Sub

Private Sub List1_DblClick()
  If VALICLIE% = 1 Then
    Call CARDETOBRA(2)
  End If
End Sub

Private Sub Text1_Change()
   If VALICLIE% < 1 Then
      Call BLANFORMU
   End If
End Sub

Private Sub Text1_DblClick()
  Call command1_Click
End Sub

Private Sub Text1_GotFocus()
'   If MARCONOTA.Visible = True Then
'      Command6.SetFocus
'      Exit Sub
'    ElseIf MARSELPRESU.Visible = True Then
'      Command19.SetFocus
'      Exit Sub
'   End If
   Text1.TEXT = TRIM$(Text1.TEXT)
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.TEXT)
'   MARCONOTA.Visible = False
'   MARSELPRESU.Visible = False
   '
   If TRIM$(Text1) = "" Then
'     Label11.ForeColor = RGB(0, 0, 255)
     Label11 = "Venta"
'     Label11.ToolTipText = "Doble-Click para 'Reparación'"
   End If
   '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call TRALOCAL("INDIOBR", "")
    For U& = 1 To ULTREG&("INDIOBR")
      X$ = REGLEIDO$("INDIOBR", U&)
      If TRIM$(Text1.TEXT) = TRIM$(Mid$(X$, 1, 8)) Then
        NCLIE% = CVI(Mid$(X$, 49, 2))
        If NCLIE% > 0 Then
          Text2.TEXT = RASOCLI$(NCLIE%)
          Text3.TEXT = DOMICLI$(NCLIE%)
          Text4.TEXT = CPOSCLI$(NCLIE%)
          Text5.TEXT = LOCACLI$(NCLIE%)
          Text12.TEXT = TRIM$(Mid$(X$, 51, 20))
          NCLIEN.Caption = CStr(NCLIE%)
          On Error Resume Next
          Command6.SetFocus
          On Error GoTo 0
          List1.Clear
          Exit For
        End If
      End If
    Next U&
  End If
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
'   If ULTREG&("PEDENCA") > 2 Then
'      On Error Resume Next
'      Text10.SetFocus
'      On Error GoTo 0
'   End If
End Sub

Private Sub Text11_LostFocus()
'   NUMEPRE$ = ""
'   For KK% = 1 To Len(Text11.TEXT)
'     CCX% = Asc(Mid$(Text11.TEXT, KK%, 1))
'     If CCX% >= 48 Then
'       If CCX% <= 57 Then
'         NUMEPRE$ = NUMEPRE$ + Chr$(CCX%)
'       End If
'     End If
'   Next KK%
'   '
'   NPP& = NUOBRA&
'   If Val(NUMEPRE$) < NPP& Then
'     Call MENSERR(24, "Número no Válido")
'     Text11.TEXT = TRIM$(Str$(NPP&))
'     Exit Sub
'   End If
'   '
'   If NUMEPRE$ <> TRIM$(Text11.TEXT) Then
'     Text11.TEXT = NUMEPRE$
'     Call MENSERR(24, "Número de Presupuesto Ajustado")
'   End If
'   '
End Sub

Private Sub Text12_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text13_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text2_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
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

Private Sub Text44_LostFocus()
   If Val(Text44) > 99.9 Then Text42 = "99.9"
   If TRIM$(Text44) <> "" Then Text44 = Format$(Val(Text44), "#0.0")
End Sub

Private Sub Text5_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub COMMAND21_CLICK(Index As Integer)
    If VALICLIE% < 1 Then
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
       Exit Sub
    End If
    '
    If Index = 0 Then
        ARCHEC$ = "LISTAS": LOCA% = 2
      ElseIf Index = 1 Then
        ARCHEC$ = "CONPAG": LOCA% = 2
      ElseIf Index = 2 Then
        ARCHEC$ = "VENDEDOR": LOCA% = 1
    End If
    'ARGU$ = CBIN$(TRIM$(TEXT21(Index).TEXT), 1, LOCA%)
    Call SELECHO(ARCHEC$)
    VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
    If VDEVU$ <> "" Then
       'TEXT21(Index).TEXT = VDEVU$
    End If
    '
End Sub
   '
Function VALICLIE%()
   VCC% = 0
   If TRIM$(Text1.TEXT) <> "" Then
     If ECOARCH$("INDIOBR", Text1) <> "" Then
       VCC% = 1
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
'   Text7.TEXT = ""
'   Text8.TEXT = ""
   Text10.TEXT = ""
   Text12.TEXT = ""
   Text13.TEXT = ""
   Text15 = ""
   'Text6.TEXT = ""
   Label18 = ""
   List1.Clear
   Call SETULTREG("!CARDETPE", 0)
   '
'   TEXT21(0).TEXT = "": Eco21(0) = ""
'   TEXT21(1).TEXT = "": Eco21(1) = ""
'   TEXT21(2).TEXT = "": Eco21(2) = ""
   Text44 = ""
   '
'   PORDESCU.TEXT = ""
'   For KK% = 0 To 3
'      Total(KK%).TEXT = ""
'      TOTIVA(KK%).TEXT = ""
'      LENTREGA(KK%).TEXT = ""
'   Next KK%
'   '
  ' MARCONOTA.Visible = False
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   Text11.TEXT = TRIM$(Str$(NUOBRA&))
   '
End Sub
'
Sub CARDACLI()
    '
    Screen.MousePointer = 11
    MARCONOTA.Visible = False
    MARSELPRESU.Visible = False
    If CONTROL$("ADMIPED", "VERLICRE") = "SI" Then VERLICRE% = 1
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    NC% = Val(Text1.TEXT)
    If NC% > 0 Then
      '
      If InStr(EPAC$, "DOSIVA") > 0 Then
        PVCL% = PIVACLI%(NC%)
        If PVCL% <> 0 And PVCL% <> 5 Then
          PRVC$ = TRIM$(PROVCLI$(NC%))
          If PRVC$ = "" Or DEPROVI$(PRVC$) = "" Then
            Call COMUNI("Imposible Registrar O/Venta.\Falta Código de Provincia del Cliente.")
            Text1 = ""
            GoTo 99
          End If
        End If
      End If
      '
      If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
        Text2.TEXT = RASOCLI$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        Text7.TEXT = CSTRING$(MKS$(LIMICRE(NC%)), 3, 13, 2, 2)
        Text12.TEXT = TRIM$(POSIVAC$(NC%)) + " - CUIT " + CUITCLI$(NC%)
        CREDISPO# = LIMICRE(NC%) - SALDCLI#(NC%) - CHEPEACRE#(NC%)
        Text8.ForeColor = RGB(0, 0, 0)
        If CREDISPO# < 0 Then Text8.ForeColor = RGB(255, 0, 0)
        Text8.TEXT = CSTRING$(MKD$(Abs(CREDISPO#)), 3, 13, 2, 2)
        '
'        TEXT21(0).TEXT = Str$(LIPRCLI%(NC%))
'        TEXT21(1).TEXT = Str$(CPAGCLI%(NC%))
'        TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
'        '
'        If Val(TEXT21(2).TEXT) < 1 Then
'           Eco21(2).TEXT = "GERENCIA"
'        End If
'        '
        PORDESCU.TEXT = DESCUCLI$(NC%)
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
'        LENTREGA(0) = Left$(LENTREG$, 32)
'        LENTREGA(1) = Mid$(LENTREG$, 33, 32)
'        LENTREGA(2) = Mid$(LENTREG$, 65, 32)
        '
        RNC$ = RECFILT$("NOTACLI", 1, MKI$(NC%))
        URN& = ULTREG&("NOTACLI")
        Text6.TEXT = ""
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
        MARCONOTA.Visible = True
        '
        If VERLICRE% = 1 Then
          LMCR# = LIMICRE(NC%)
          SDEU# = SALDCLI#(NC%)
          CHPA# = CHEPEACRE#(NC%)
          If SDEU# + CHPA# > LMCR# Then
            Call COMUNI("Este Pedido no Podrá Facturarse.\  \Límite de Crédito Asignado ya está Superado.\  \Informe al Supervisor de Ventas.")
          End If
        End If
        '
      End If
    End If
99  Screen.MousePointer = 1
    '
End Sub
   '
Function NUOBRA&()
   '
   NOBRA& = 1
   '
   FIN& = ULTREG&("BDSIMAT")
   If FIN& > 0 Then
      For U& = FIN& To FIN& - 16 Step -1
         If U& < 1 Then GoTo 99
         XX$ = REGLEIDO$("BDSIMAT", U&)
         NOBRAX& = CVS(Mid$(XX$, 1, 4))
         If NOBRAX& >= NOBRA& Then
            NOBRA& = 1 + NOBRAX&
         End If
      Next U&
   End If
   '
99 NUOBRA& = NOBRA&
   '
End Function
'
Private Sub Text6_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text7_GotFocus()
   On Error Resume Next
   Text1.SetFocus
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
     VH(MONEMI%) = Val(TICAMBIS)
   End If
End Sub

Private Sub TICAMBIS_LostFocus()
   TICAMBIS = FORMATNUM(Val(TICAMBIS), "F9.4")
End Sub

Private Sub Timer1_Timer()
   '
   Call CALTOPED
   '
End Sub

Private Sub Timer2_Timer()
  '
  Static FIDAPRIN#
  '
  FDP# = FIDATIM#("IPEDPRIN")
  If FDP# <> FIDAPRIN# Then
    FIDAPRIN# = FDP#
    RFF$ = RECFILT$("IPEDPRIN", 3, MKI$(0))
    '
    Screen.MousePointer = 11
    For UL& = 1 To Len(RFF$) Step 4
      U& = CVS(Mid$(RFF$, UL&, 4))
      YY$ = REGLEIDO$("IPEDPRIN", U&)
      TDOC$ = UCase$(Left$(YY$, 2))
      NPX& = Val(Mid$(YY$, 3, 6))
      '
      If NPX& > 0 Then
        TIDOCUM$ = "Pedido de Cliente"
        If TDOC$ = "OD" Then TIDOCUM$ = "Orden de Despacho"
        '
        NUDOCU$ = TRIM$(Str$(NPX&))
        '
        ' esto para forzar que haya bajado al disco
        Call BLOQARCH("PDOCSPL")
        Call LIBARCH("PDOCSPL")
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
        '
        GoTo 49
        '
5       REDOCU$ = REGACT$("PDOCLST")
        URE& = ULTREG&("PDOCSPL")
        LI& = 0: LS& = URE&
6       E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 49
        L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 49
        DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
        If DCC& < DOCUNU& Then LI& = L&: GoTo 6
        If DCC& > DOCUNU& Then LS& = L&: GoTo 6
        '
8       TPSP$ = ""
        Screen.MousePointer = 11
        While L& > 1
          DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
          If DCC& < DOCUNU& Then GoTo 9
          L& = L& - 1
        Wend
        '
9       For KKL& = L& To URE&
          TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
          If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 10
          If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
            TPSP$ = TPSP$ + Mid$(TBUF$, 5)
          End If
        Next KKL&
        '
10      PPXX% = InStr(TPSP$, Chr$(255))
        If PPXX% > 0 Then
          COPRIAT$ = Left$(TPSP$, PPXX% - 1)
          TPSP$ = Mid$(TPSP$, PPXX% + 1)
          '
          If Len(COPRIAT$) > 0 Then
            COPRI% = Asc(Left$(COPRIAT$, 1))
            ATRIP$ = Mid$(COPRIAT$, 2)
            Call PRINTCOMMAND(COPRI%, ATRIP$)
          End If
          GoTo 10
        End If
        '
      End If
      '
      YY$ = REGLEIDO$("IPEDPRIN", U&)
      HOII% = HOY(HOS$)
      Call REPLA(YY$, MKI$(HOII%), 57, 2)
      Call REPLA(YY$, Time$, 59, 6)
      Call GRAREG("IPEDPRIN", YY$, U&)
      Call CIERRARCH("IPEDPRIN")
      '
49  Next UL&
    Screen.MousePointer = 1
    '
  End If
  '
End Sub

Private Sub Total_Change(Index As Integer)
   If Index = 0 Then
     Call CALTOPED
'      DESC$ = PORDESCU.TEXT
'      TTT1 = Val(Total(0).TEXT) * COEFDESCU(DESC$) / 100
'      Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
'      TTT3 = Val(Total(0).TEXT) - Val(Total(1).TEXT)
'      Total(2).TEXT = CSTRING$(MKS$(TTT3), 3, 11, 2, 2)
   End If
'   If Index < 3 Then
'      TTT1 = Val(Total(2).TEXT)
'      TTT2 = Val(TOTIVA(3).TEXT)
'      Total(3).TEXT = Str$(TTT1 + TTT2)
'   End If
End Sub
'
Sub PRIPEDIDO()
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORSIMT"))
    If FORIPRE$ = "" Then Exit Sub
    '
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Solicitud Interna de Materiales.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    CANCELPRINT = 0
    If CONTROL$("FORSIMT", "SUPRINT") = "SI" Then
       CANCELPRINT% = 1
    End If
    '
'    EPAC$ = TRIM$(UCase$(EMPREAC$))
'    If InStr(EPAC$, "DOSIVAC") > 0 Then
'      CANCELPRINT% = 0
'      Screen.MousePointer = 1
'      OPRIPEDI.Show 1
'      If OPRIPEDI.Check1(0).Value <> 1 Then
'        CANCELPRINT% = 1
'      End If
'    End If
    '
    Screen.MousePointer = 11
    FECHDOC$ = Text9.TEXT
    NUMEDOC$ = TRIM$(Text11.TEXT)
    While Len(NUMEDOC$) < 6
      NUMEDOC$ = "0" + NUMEDOC$
    Wend
    '
    TIPODOC$ = "Solicitud Interna de Materiales"
    If NC% < 1 Then
      TIPODOC$ = "Solicitud Interna de Materiales"
    End If
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "UNIMED"
    CODTABU1$(4) = "CANTIDAD"
    CODTABU1$(5) = "CANT-DIS"
    CODTABU1$(6) = "FECH-VEN"
'    CODTABU1$(7) = "NET-ITEM"
'    CODTABU1$(8) = "FECH-REC"
'    CODTABU1$(9) = "ORD-ITEM"
    CACOLTAB1% = 6
    '
    NC% = CInt(NCLIEN.Caption)
    DESTIDOC% = NC%
    CAITAB1% = 0
    CAITCRU% = 0
    '
    For U& = 1 To ULTREG&("!MATEROF")
      XX$ = REGLEIDO$("!MATEROF", U&)
      CIPRE% = CVI(Left$(XX$, 2))
      CAPEX$ = Str$(CVS(Mid$(XX$, 5, 4)))
      '
      If CIPRE% > 0 Then
        If CIPRE% <= NUMAS% Then
          If Val(CAPEX$) > 0 Then
            '
      '      PRELI$ = Str$(CVS(Mid$(XX$, 39, 4)))
      '      PREUX$ = Str$(CVS(Mid$(XX$, 51, 4)))
      '      DESCX$ = "0": On Error Resume Next
      '                    DESCX$ = Str$(100 * (Val(PRELI$) - Val(PREUX$)) / Val(PRELI$))
      '                    On Error GoTo 0
      '      ITOTX$ = Str$((Val(CAPEX$)) * (Val(PREUX$)))
      '      FEX = CVI(Mid$(XX$, 63, 2))
      '      FENTX$ = FECHATEX$(FEX)
            '
            CAITAB1% = CAITAB1% + 1
            TETABU1$(1, CAITAB1%) = CODEXT$(CIPRE%)
            TETABU1$(2, CAITAB1%) = DESCRIT0$(CIPRE%)
            TETABU1$(3, CAITAB1%) = TRIM$(Mid$(XX$, 3, 2))
            TETABU1$(4, CAITAB1%) = CAPEX$
            TETABU1$(5, CAITAB1%) = Str$(CVS(Mid$(XX$, 21, 4)))
              FEX = CVI(Mid$(XX$, 29, 2))
            TETABU1$(6, CAITAB1%) = FECHATEX$(FEX)
            'TETABU1$(7, CAITAB1%) = ITOTX$
            'TETABU1$(8, CAITAB1%) = FENTX$
            'TETABU1$(9, CAITAB1%) = Str$(CAITAB1%)
            '
          End If
        End If
      End If
109 Next U&
    '
'    SUMATOTA# = Val(Total(0).TEXT)
'    IDESCUEN# = Val(Total(1).TEXT)
'    '
    CODPARAM$(1) = "DESTINO": DOCPARAM$(1) = TRIM$(Text1) & " - " & TRIM$(Text12)
'    CODPARAM$(2) = "VENDEDOR":     DOCPARAM$(2) = Eco21(2).TEXT
'    CODPARAM$(3) = "COND-PAG":     DOCPARAM$(3) = Eco21(1).TEXT
'    CODPARAM$(4) = "NRO-OCOM":     DOCPARAM$(4) = Text10.TEXT
    'CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = LENTREGA(0).TEXT
    'CODPARAM$(6) = "DESTINO": DOCPARAM$(6) = LENTREGA(1).TEXT
    'CODPARAM$(7) = "TRANSPOR": DOCPARAM$(7) = LENTREGA(2).TEXT
'    CODPARAM$(8) = "IMP-DESC": DOCPARAM$(8) = Total(1).TEXT
'    CODPARAM$(9) = "IMPNEDES": DOCPARAM$(9) = Total(2).TEXT
'    CODPARAM$(10) = "IMP-IVA": DOCPARAM$(10) = TOTIVA(3).TEXT
'    CODPARAM$(11) = "IMP-TOTA": DOCPARAM$(11) = Total(3).TEXT
'    CODPARAM$(12) = "POR-DESC": DOCPARAM$(12) = TRIM$(PORDESCU.TEXT)
'    CODPARAM$(13) = "MONEDEMI": DOCPARAM$(13) = MONEMIS
'    CODPARAM$(14) = "PERC-IVA": DOCPARAM$(14) = Text15
    CAPARDOC% = 1
    '
    Call BLANARCH("!OBSERVOF")
    Call FRATEXTO(Text13.TEXT, 70)
    If CARETEX% >= 1 Then Call GRAREG("!OBSERVOF", RETEX$(1), 1)
    If CARETEX% >= 2 Then Call GRAREG("!OBSERVOF", RETEX$(2), 2)
    If CARETEX% >= 3 Then Call GRAREG("!OBSERVOF", RETEX$(3), 3)
    URETE& = CARETEX%
    Call SETULTREG("!OBSERVOF", URETE&)
    Call CIERRARCH("!OBSERVOF")
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORSIMT"))
    If FORIPRE$ <> "" Then
      Call PRINTDOC("FORSIMT")   ' SOLICITUD INTERNA DE MATERIALES
    End If
    '
'    FORIPRE$ = TRIM$(CONTROL$("", "FORDESPA"))
'    If FORIPRE$ <> "" Then
'      If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") > 0 Then
'        TIPODOC$ = "Orden de Despacho"
'        CANCELPRINT = 0
'        If CONTROL$("ORDESPA", "SUPRINT") = "SI" Then
'           CANCELPRINT% = 1
'        End If
'        '
'        If InStr(EPAC$, "DOSIVAC") > 0 Then
'          CANCELPRINT% = 0
'          If OPRIPEDI.Check1(1).Value <> 1 Then
'            CANCELPRINT% = 1
'          End If
'        End If
'        '
'        If CARETEX% >= 1 Then Call GRAREG("!OBSERVOF", RETEX$(1), 1)
'        If CARETEX% >= 2 Then Call GRAREG("!OBSERVOF", RETEX$(2), 2)
'        If CARETEX% >= 3 Then Call GRAREG("!OBSERVOF", RETEX$(3), 3)
'        URETE& = CARETEX%
'        Call SETULTREG("!OBSERVOF", URETE&)
'        Call CIERRARCH("!OBSERVOF")
'        '
'        Call PRINTDOC("FORDESPA")   ' ORDEN DE DESPACHO
'      End If
'    End If
    '
End Sub

Private Sub Total_GotFocus(Index As Integer)
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Sub GRAOBRA(OKX%)
   '
   KKK% = 0
   OKX% = 0
   '
   If ULTREG&("!MATEROF") < 1 Then
     Call MENSERR(24, "Falta Detalle de los Materiales")
     Exit Sub
   End If
   '
   Call CIERRARCH("BDSIMAT")
   Call CIERRARCH("INDISIM")
   Call CIERRARCH("ANOTASIM")
   '
   Call BLOQARCH("BDSIMAT")
   Call BLOQARCH("INDISIM")
   Call BLOQARCH("ANOTASIM")
   '
   NOBRA& = NUOBRA&
   If NOBRA& > Val(Text11.TEXT) Then
      Call CIERRARCH("BDSIMAT")
      Call CIERRARCH("INDISIM")
      Call CIERRARCH("ANOTASIM")
      Call COMUNI("FLEXOFT ha Re-Numerado la S.I.M.\como Solicitud Nro '" + TRIM$(Str$(NOBRA&)) + "'.\  \Verifique y Vuelva a Grabar.")
      Text11.TEXT = TRIM$(Str$(NOBRA&))
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   If VALICLIE% < 1 Then
      MERRO$ = "Código de Obra no Válido"
      GoTo 98
   End If
   NCLIE% = Val(NCLIEN.Caption)
   '
   HOII% = HOY(HOS$)
   FEMI% = FECHANUM(Text9.TEXT)
   If FEMI% < 1 Or FEMI% > HOII% Then
      MERRO$ = "Fecha de Emisión no Válida"
      GoTo 98
   End If
   '
   REFE$ = " - " + TRIM$(Text10.TEXT)
   If REFE$ = " - " Then REFE$ = ""
   FEX = FEMI%
   REFE$ = REFE$ + " - " + FECHATEX$(FEX)
   LOMA% = 48 - Len(REFE$)
   REFE$ = Left$(TRIM$(ECOARCH$("DEUDOR1", MKI$(NCLIE%))), LOMA%) + REFE$
   '
   TTXX$ = Space$(76)
   HOII% = HOY(HOS$)
   Call REPLA(TTXX$, HOS$, 1, 8)
   Call REPLA(TTXX$, "Generación de Solicitud de Materiales:", 11, 56)
   Call REPLA(TTXX$, USERNAME$, 69, 24)
   Call REGAPP("ANOTASIM", MKS$(NOBRA&) + TTXX$)
   Call REGAPP("ANOTASIM", MKS$(NOBRA&) + "       ")
   '
   SIMANT& = 0
   For U& = 1 To ULTREG&("!MATEROF")
     X$ = REGLEIDO$("!MATEROF", U&)
     CAN = CVS(Mid$(X$, 5, 4))
     If CAN > 0.05 Then
       SIM$ = REGBLAN$("BDSIMAT")
       Call REPLA(SIM$, MKS$(NOBRA&), 1, 4)
       Call REPLA(SIM$, MKI$(FEMI%), 5, 2)
       Call REPLA(SIM$, TRIM$(Text1), 7, 8)
       Call REPLA(SIM$, TRIM$(Text10), 15, 24)
       Call REPLA(SIM$, MKI$(NCLIE%), 39, 2)
       Call REPLA(SIM$, MKI$((USNBR% Mod 100)), 41, 2)
       Call REPLA(SIM$, Mid$(USERNAME$, 1, 22), 43, 22)
         CI% = CVI(Mid$(X$, 1, 2))
       Call REPLA(SIM$, Mid$(X$, 1, 2), 65, 2)
       Call REPLA(SIM$, UNIMED$(CI%), 67, 4)
         CAN = CVS(Mid$(X$, 5, 4))
       Call REPLA(SIM$, Mid$(X$, 5, 4), 71, 4)
         FENTRESOL% = CVI(Mid$(X$, 29, 2))
       Call REPLA(SIM$, MKI$(FENTRESOL%), 75, 2)
       Call REGAPP("BDSIMAT", SIM$)
       '
       ' GRABA EL ARCHIVO DE INDICES DE SIM
       If SIMANT& <> NOBRA& Then
         SIM$ = REGBLAN$("INDISIM")
         Call REPLA(SIM$, MKS$(NOBRA&), 1, 4)
         Call REPLA(SIM$, TRIM$(Text1) & " - " & TRIM$(Text12), 5, 60)
         Call REGAPP("INDISIM", SIM$)
         SIMANT& = NOBRA&
       End If
       '
       TTXX$ = Space$(76)
       CIXI% = CI%
       Call REPLA(TTXX$, "- " + CODEXT$(CI%), 11, 18)
       Call REPLA(TTXX$, DESCRIT0$(CI%), 29, 38)
       Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 10, 1, 2), 57, 10)
       Call REGAPP("ANOTASIM", Left$(MKS$(NOBRA&) + TTXX$, 96))
       '
     End If
   Next U&
   '
   Call FRATEXTO(Text13.TEXT, 32)
   For KKI% = 1 To CARETEX%
     TTXX$ = Space$(92)
     Call REPLA(TTXX$, RETEX$(KKI%), 16, 32)
     If KKI% = 1 Then
       Call REGAPP("ANOTASIM", MKS$(NOBRA&) + Space$(96))
       Call REPLA(TTXX$, "Observaciones:", 1, 15)
     End If
     Call REGAPP("ANOTASIM", Left$(MKS$(NOBRA&) + Space$(10) + TTXX$, 96))
   Next KKI%
   '
   Call REGAPP("ANOTASIM", MKS$(NOBRA&) + String$(88, "="))
   Call CIERRARCH("ANOTASIM")
   '
   Call CIERRARCH("BDSIMAT")
   Call CIERRARCH("INDISIM")
   Call CIERRARCH("ANOTASIM")
   '
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Call BAJALDISCO
   '
   OKX% = 1
   '
98 Call CIERRARCH("BDSIMAT")
   Call CIERRARCH("INDISIM")
   Call CIERRARCH("ANOTASIM")
   Screen.MousePointer = 1
   '
End Sub

Private Sub Total_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
     If Index = 0 Then
       On Error Resume Next
       PORDESCU.SetFocus
       On Error GoTo 0
     End If
   End If
End Sub

Sub CARDETOBRA(MODO%)
   '
   HOII% = HOY(HOS$)
   '
   ATRI$ = "/NC/ARCHECO(1)=" + ARCHIPRE$
   ATRI$ = ATRI$ + "/TITUPAN=Detalle de Materiales"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture3.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture3.Left + 50))
   '
   Timer1.Enabled = True
10
   PRF$ = "EDITOBRA"
   XXX% = VENTABU%(PRF$ + ATRI$)
   If XXX% < 0 Then
     Call CIERRARCH("!MATEROF")
     Call BLANARCH("!MATEROF")
     Text1 = ""
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
   Timer1.Enabled = False
   '
20 If ULTREG&("!MATEROF") < 1 Then
     Screen.MousePointer = 1
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
   '
   Screen.MousePointer = 11   ' Registrando Obra ...
   FIN& = ULTREG&("!MATEROF")
   If FIN& > 0 Then
     Y$ = REGLEIDO$("!MATEROF", FIN&)
     CI% = CVI(Left$(Y$, 2))
     If CI% < 1 Or CI% > NUMAS% Then
       Call SETULTREG("!MATEROF", FIN& - 1)
       Call CIERRARCH("!MATEROF")
       Call MENSERR(24, "Código No Reconocido o No Válido\Debe Correr Control de Consistencia de Maestro.")
       GoTo 20
     End If
   End If
   '
   For U& = 1 To FIN&
     Y$ = REGLEIDO$("!MATEROF", U&)
     CI% = CVI(Left$(Y$, 2))
     If CI% < 1 Or CI% > NUMAS% Then
       Call MENSERR(24, "Código no Válido en Registro " + TRIM$(Str$(U&)))
       Call REPLA(Y$, MKI$(0), 1, 2)
       Call GRAREG("!MATEROF", Y$, U&)
       GoTo 10
     End If
     '
     FENTRESOL% = CVI(Mid$(Y$, 29, 2))
     If FENTRESOL% < HOII% Then
        FENTRESOL% = HOII%
        Call REPLA(Y$, MKI$(FENTRESOL%), 29, 2)
        Call GRAREG("!MATEROF", Y$, U&)
     End If
     '
     X$ = REGLEIDO$("!MATEROF", U&)
     CI% = CVI(Left$(X$, 2))
     TTXX$ = Space$(96)
     FENTRESOL% = CVI(Mid$(X$, 29, 2))
     '
     CAN = CVS(Mid$(X$, 5, 4))
     If CAN > 999999 Then
       Call MENSERR(24, "Cantidad Excesiva - Max 999999")
       Call REPLA(X$, MKS$(0), 5, 4)
       Call GRAREG("!MATEROF", X$, U&)
       GoTo 10
     End If
     '
     If CAN < 0.5 Then
       Call MENSERR(24, "Falta Cantidad")
       Call REPLA(X$, MKS$(0), 5, 4)
       Call GRAREG("!MATEROF", X$, U&)
       GoTo 10
     End If
     '
     X$ = REGLEIDO$("!MATEROF", U&)
     Call REPLA(X$, UNIMED(CI%), 3, 2)
     Call GRAREG("!MATEROF", X$, U&)
     '
     X$ = REGLEIDO$("!MATEROF", U&)
     Call REPLA(X$, MKS$(STOCKDIS(CI%)), 21, 4)
     Call GRAREG("!MATEROF", X$, U&)
     '
   Next U&
   '
   Call CARLISPE
   Call CIERRARCH("!MATEROF")
   GoTo 99
   Call CALTOPED
   '
99 List1.Refresh
   On Error Resume Next
   Command14.SetFocus
   On Error GoTo 0
   Screen.MousePointer = 1
   '
End Sub

Sub CARLISPE()
   List1.Clear
   FIN& = ULTREG&("!MATEROF")
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!MATEROF", U&)
     TTXX$ = Space$(96)
     CI% = CVI(Left$(X$, 2))
     FENTRESOL% = CVI(Mid$(X$, 29, 2))
     CAN = CVS(Mid$(X$, 5, 4))
     UM$ = Mid$(X$, 3, 2)
     STDIS = CVS(Mid$(X$, 21, 4))
'     PORDESCUIN = CVS(Mid$(X$, 69, 4))
'     PRECARGA = CVS(Mid$(X$, 39, 4))
'     PREUNI = PRECARGA * (100 - PORDESCUIN) / 100
'     TOT# = CDbl(CAN * PREUNI)
     '
'     Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 1, 6)
'     Call REPLA(TTXX$, CODEXT$(CI%), 8, 15)
'     FEX = FENTRESOL%
'     If ECOPRECIO% < 0 Then
'          Call REPLA(TTXX$, DESCRIT0$(CI%), 24, 52)
'          Call REPLA(TTXX$, FECHATEX$(FEX), 78, 8)
'        Else
'          Call REPLA(TTXX$, DESCRIT0$(CI%), 24, 22)
'          Call REPLA(TTXX$, Left$(FECHATEX$(FEX), 5), 47, 5)
'     End If
'     PRELIX = PRELISTA#(ARCHIPRE$, CI%)
'     If PRELIX < PRECARGA Then PRELIX = PRECARGA
'     If PRELIX < PREUNI Then PRELIX = PREUNI
'     PORDESCUIN = 0: If PRELIX > 0.005 Then PORDESCUIN = 100 * (PRELIX - PREUNI) / PRELIX
'     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PRELIX), 3, 8, 2, 2), 52, 8)
'     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PORDESCUIN), 3, 5, 1, 2), 61, 5)
'     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKS$(PREUNI), 3, 8, 2, 2), 66, 8)
'     If ECOPRECIO% >= 0 Then Call REPLA(TTXX$, CSTRING$(MKD$(TOT#), 3, 10, 2, 2), 76, 10)

     Call REPLA(TTXX$, CODEXT$(CI%), 1, 16)
     Call REPLA(TTXX$, DESCRIT0$(CI%), 17, 32)
     Call REPLA(TTXX$, UM$, 51, 2)
     Call REPLA(TTXX$, FORMATNUM$(CAN, "F8.1"), 56, 8)
     Call REPLA(TTXX$, FORMATNUM$(STDIS, "F8.1"), 67, 8)
     FEX = FENTRESOL%
     Call REPLA(TTXX$, FECHATEX$(FEX), 77, 8)
     List1.AddItem TTXX$
     '
   Next U&

End Sub
'
Sub CALTOPED()
   '
   Static CLIA$, RNI%, APLIVA%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE% = Val(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIE% > 0 Then
       PVCLI% = PIVACLI%(NCLIE%)
       If PVCLI% = 2 Then RNI% = 1
       If PVCLI% = 5 Or PVCLI% = 7 Then APLIVA% = 0
     End If
   End If
   '
   SUTOTPED# = 0: SUIVA# = 0: SUIVARNI# = 0: TOIVA# = 0: TOTOTA# = 0
   TOTITEMS% = 0
   For U& = 1 To ULTREG&("!CARDETPE")
     X$ = REGLEIDO$("!CARDETPE", U&)
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
       Call GRAREG("!CARDETPE", X$, U&)
       TOTITEMS% = TOTITEMS% + 1
     End If
   Next U&
   '
   Label18 = TRIM$(Str$(TOTITEMS%))
   '
'   Total(0).TEXT = CSTRING$(MKD$(SUTOTPED#), 3, 11, 2, 2)
'   Total(0).Refresh
'   '
'   DESC$ = TRIM$(PORDESCU.TEXT)
'   TTT1 = Val(Total(0).TEXT) * COEFDESCU(DESC$) / 100
'   Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
'   '
'   TTT3 = Val(Total(0).TEXT) - Val(Total(1).TEXT)
'   Total(2).TEXT = CSTRING$(MKS$(TTT3), 3, 11, 2, 2)
'   '
'   On Error Resume Next
'   TOTIVA(0).TEXT = CSTRING$(MKD$(APLIVA% * SUIVA# * Val(Total(2).TEXT) / Val(Total(0).TEXT)), 4, 11, 2, 2)
'   On Error GoTo 0
'   TOTIVA(2).TEXT = CSTRING$(MKD$(CDbl(Val(TOTIVA(0).TEXT) * 0.5 * RNI%)), 4, 11, 2, 2)
'   TOTIVA(3).TEXT = CSTRING$(MKD$(CDbl(Val(TOTIVA(0).TEXT) + Val(TOTIVA(2).TEXT))), 4, 11, 2, 2)
'   '
'   Call CALPERIB
'   '
'   Total(3).TEXT = CSTRING$(MKD$(CDbl(Val(Total(2).TEXT) + Val(TOTIVA(3).TEXT)) + Val(Text15)), 4, 11, 2, 2)
'   '
End Sub

Private Sub TOTIVA_Change(Index As Integer)
   'If Index = 3 Then
   '  Total(3).TEXT = Str$(Val(Total(2).TEXT) + Val(TOTIVA(3).TEXT))
   'End If
End Sub

Sub GENERODES()
   '
   If CONTROL$("ADMIPED", "GENDESPA") = "REVISTOCK" Then
     Call GENODESPA(Val(Text11), NODES&, 1, OKKY%)
     If OKKY% = 1 Then
        Call PRIODESPA(NODES&)
     End If
   End If
   '
End Sub
'
Sub RECODOSI(CIXI%, CIXJ%, NPX&)
   '
   CIBIN$ = MKI$(CIXI%)
   CIXJ% = CIXI%
   If Asc(Mid$(CIBIN$, 2, 1)) = 0 Then
     If NPX& > 1600 Then
       If NPX& <= 1614 Then
         If NPX& = 1601 Then CIXJ% = CODINT%("041130M-PVT0")
         If NPX& = 1602 Then CIXJ% = CODINT%("042121T-PHH0")
         If NPX& = 1603 Then CIXJ% = CODINT%("041015M-PVT0")
         If NPX& = 1604 Then CIXJ% = CODINT%("041015M-PVT0")
         If NPX& = 1605 Then CIXJ% = CODINT%("041130M-PVT0")
         If NPX& = 1606 Then CIXJ% = CODINT%("041130M-PVT0")
         If NPX& = 1607 Then CIXJ% = CODINT%("042501T-PHH0")
         If NPX& = 1608 Then CIXJ% = CODINT%("44234")
         If NPX& = 1609 Then CIXJ% = CODINT%("44234")
         If NPX& = 1610 Then CIXJ% = CODINT%("LD54S2-PHCA-D13")
         If NPX& = 1611 Then CIXJ% = CODINT%("044501T-PTT0")
         If NPX& = 1612 Then CIXJ% = CODINT%("044251T-PTT0")
         If NPX& = 1613 Then CIXJ% = CODINT%("041130M-PVT0")
         If NPX& = 1614 Then CIXJ% = CODINT%("A4190")
       End If
     End If
   End If
   If CIXJ% < 1 Then CIXJ% = CIXI%
   '
End Sub

Sub CALPERIB()
   '
   Static CTTXX%, TASAA, TASAB, TASAEXP, MINIPER
   '
   'If TIFADOC% > 2 Then
   '   Exit Sub
   'End If
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
   NCLIE% = Val(Text1)
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
'       BIMPON# = Val(Total(2))
'       If VARIDC% = 2 Then BIMPON# = BIMPON# + Val(TOTIVA(3))
'       If BIMPON# * TICAMONE(MONEMI%) >= MINIPER Then
'         PERSALTA# = (BIMPON# * PORCONVE / 100) * (TASAX / 100)
'         Text15 = FORMATNUM$(PERSALTA#, "F11.2")
'       End If
     End If
   End If
   '
End Sub
'
Sub PONEMARPREP()
   '
   If COMALTER%("FLEXOFT Debe Ahora Reorganizar\la Base de Datos de Pedidos\  \Todos los Usuarios Deben Salir\del Módulo de Pedidos de Clientes\\Cancelar\Re-Organizar Ahora") <> 2 Then
     Call FINAL("")
   End If
   '
   Call INDEXARCH("PEDDETA", 1)
   Call BLOQARCH("PEDDETA")
   Call BLOQARCH("PEDDETP")
   Call BLOQARCH("DESPAPED")
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("PEDDETA")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("PEDDETA", U&)
     NROPE& = CVS(Left$(XX$, 4))
     If NROPE& > 0 Then
       CAVER$ = TRIM$(Str$(NROPE&))
       While Len(CAVER$) > 6: CAVER$ = Mid$(CAVER$, 2): Wend
       While Len(CAVER$) < 6: CAVER$ = "0" + CAVER$: Wend
       '
       If Mid$(XX$, 113, 6) <> CAVER$ Then
         IKITEM% = 0
         RFF$ = RECFILT$("PEDDETA", 1, MKS$(NROPE&))
         For UL& = 1 To Len(RFF$) Step 4
           REPED& = CVS(Mid$(RFF$, UL&, 4))
           YY$ = REGLEIDO$("PEDDETA", REPED&)
           '
           IKITEM% = IKITEM% + 1
           IK$ = TRIM$(Str$(IKITEM%))
           CAVERI$ = CAVER$ + "-" + IK$
           Call REPLA(YY$, CAVERI$, 113, 10)
           Call GRAREG("PEDDETA", YY$, REPED&)
         Next UL&
       End If
     End If
   Next U&
   Call CIERRARCH("PEDDETA")
   '
   FIN& = ULTREG&("PEDDETP")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("PEDDETP", U&)
     NROPE& = CVS(Left$(XX$, 4))
     REPEDA& = CVS(Mid$(XX$, 109, 4))
     '
     If REPEDA& > 0 Then
       YY$ = REGLEIDO$("PEDDETA", REPEDA&)
       If CVS(Left$(YY$, 4)) = NROPE& Then
         CAVERI$ = Mid$(YY$, 113, 10)
         Call REPLA(XX$, CAVERI$, 113, 10)
         Call GRAREG("PEDDETP", XX$, U&)
       End If
     End If
   Next U&
   Call CIERRARCH("PEDDETP")
   '
   FIN& = ULTREG&("DESPAPED")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("DESPAPED", U&)
     NROPE& = CVS(Left$(XX$, 4))
     REPEDA& = CVS(Mid$(XX$, 109, 4))
     '
     If REPEDA& > 0 Then
       YY$ = REGLEIDO$("PEDDETA", REPEDA&)
       If CVS(Left$(YY$, 4)) = NROPE& Then
         CAVERI$ = Mid$(YY$, 113, 10)
         Call REPLA(XX$, CAVERI$, 113, 10)
         Call GRAREG("DESPAPED", XX$, U&)
       End If
     End If
   Next U&
   Call CIERRARCH("DESPAPED")
   '
   Call GRACONTROL("PEDCLIEN", "MARPREPA", "SI")
   Screen.MousePointer = 1
   Call COMUNI("Re-Organización Completada")
   '
End Sub
