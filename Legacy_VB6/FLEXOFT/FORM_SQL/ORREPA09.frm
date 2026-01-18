VERSION 5.00
Begin VB.Form ORREPA09 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orden de Reparación"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   12000
   Icon            =   "ORREPA09.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   12000
   StartUpPosition =   1  'CenterOwner
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
         TabIndex        =   23
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
            TabIndex        =   25
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
         TabIndex        =   21
         Top             =   105
         Width           =   7155
      End
      Begin VB.Line Line4 
         X1              =   3570
         X2              =   10740
         Y1              =   1170
         Y2              =   1170
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
         Left            =   1695
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   22
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
      Begin VB.PictureBox Picture14 
         BackColor       =   &H00400000&
         Height          =   8325
         Left            =   11025
         ScaleHeight     =   8265
         ScaleWidth      =   1155
         TabIndex        =   40
         Top             =   15
         Width           =   1215
         Begin VB.CommandButton Command4 
            Caption         =   "Print"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   5640
            Width           =   720
         End
         Begin VB.CommandButton Command1 
            Caption         =   "save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   5160
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton Command14 
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
            Height          =   1030
            Left            =   60
            Picture         =   "ORREPA09.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Aprueba, Graba e Imprime Planilla"
            Top             =   7155
            Width           =   855
         End
         Begin VB.CommandButton Command3 
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
            Height          =   660
            Left            =   105
            Picture         =   "ORREPA09.frx":0974
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   1095
            Width           =   750
         End
         Begin VB.CommandButton Command13 
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
            Height          =   660
            Left            =   105
            Picture         =   "ORREPA09.frx":0C7E
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   3915
            Width           =   750
         End
         Begin VB.CommandButton Command11 
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
            Height          =   660
            Left            =   105
            Picture         =   "ORREPA09.frx":10C0
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Acceso a Base de Datos de Clientes"
            Top             =   2505
            Width           =   750
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Articulos"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   105
            Picture         =   "ORREPA09.frx":1502
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Acceso a Maestro de Artículos"
            Top             =   1800
            Width           =   750
         End
         Begin VB.CommandButton Command2 
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
            Height          =   660
            Left            =   105
            Picture         =   "ORREPA09.frx":180C
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   390
            Width           =   750
         End
         Begin VB.CommandButton Command15 
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
            Height          =   660
            Left            =   105
            Picture         =   "ORREPA09.frx":1956
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Ver y Reimprimir Ordenes de Reparación"
            Top             =   3210
            Width           =   750
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   150
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   41
            Top             =   6240
            Width           =   645
            Begin VB.Frame Frame8 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   42
               Top             =   105
               Width           =   12000
            End
         End
         Begin VB.Image Image3 
            Height          =   390
            Left            =   -285
            Picture         =   "ORREPA09.frx":1C60
            Top             =   6555
            Width           =   1515
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
         Left            =   6590
         TabIndex        =   29
         Text            =   " "
         Top             =   890
         Width           =   1485
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Descripción"
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
         Width           =   7890
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
            TabIndex        =   24
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
            Width           =   5100
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
         TabIndex        =   17
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
            TabIndex        =   19
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
            TabIndex        =   18
            Top             =   315
            Width           =   10100
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "S/No."
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
         Height          =   780
         Left            =   8190
         TabIndex        =   15
         Top             =   1500
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
            TabIndex        =   16
            Text            =   " "
            Top             =   360
            Width           =   2325
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
      Begin VB.Label Label26 
         Caption         =   " "
         Height          =   225
         Left            =   5145
         TabIndex        =   39
         Top             =   1155
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label25 
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
         Left            =   9600
         TabIndex        =   38
         Top             =   1140
         Width           =   1110
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "O/Vta:"
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
         TabIndex        =   37
         Top             =   1170
         Width           =   1800
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "O/R Nro:"
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
         TabIndex        =   36
         Top             =   475
         Width           =   1800
      End
      Begin VB.Label Label2 
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
         TabIndex        =   35
         Top             =   445
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
         Top             =   495
         Width           =   2220
      End
      Begin VB.Label REPRESUP 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   225
         Left            =   6510
         TabIndex        =   28
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
         TabIndex        =   20
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
         Left            =   9600
         TabIndex        =   14
         Top             =   795
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   885
         Width           =   750
      End
      Begin VB.Line Line7 
         X1              =   210
         X2              =   10710
         Y1              =   7000
         Y2              =   7000
      End
      Begin VB.Line Line8 
         X1              =   10725
         X2              =   10725
         Y1              =   2445
         Y2              =   7020
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   420
         Picture         =   "ORREPA09.frx":3B82
         ToolTipText     =   "Click Aquí para cambiar Modalidad de Facturación"
         Top             =   105
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Presupuesto:"
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
         Top             =   845
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
Attribute VB_Name = "ORREPA09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub GENERAPEDIDO(NROREP&)

   ' GRABA TABLA PEDENCA
   '
10 CONDI$ = "NroPed > 0"
   NPP& = XVALO(VALOBADA("PEDENCA", "Max(NroPed)", CONDI$)) + 1
   Label25 = TRIM$(Str$(NPP&))
  
   strSQL = "Select * from Pedenca where NroPed = " & NPP&
   Dim PedEnca As ADODB.Recordset
   Set PedEnca = New ADODB.Recordset
25 On Error Resume Next
   PedEnca.Open FILTSQL$(strSQL), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   
   With PedEnca
     If Not (.EOF And .BOF) Then 'si lo encuentra
       PedEnca.Close
       Set PedEnca = Nothing
       GoTo 10
     End If
     .AddNew
     On Error Resume Next
     !CodiEmpr = CodiEmp%
     On Error GoTo 0
     !NROPED = NPP&
     .Update
   End With
   PedEnca.Close
   Set PedEnca = Nothing
   '
   Set PedEnca = New ADODB.Recordset
35 On Error Resume Next
   PedEnca.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 35
   End If
   On Error GoTo 0
   
   With PedEnca
     If (.EOF And .BOF) Then 'si NO lo encuentra
        PedEnca.Close
        Set PedEnca = Nothing
        GoTo 10
     End If
     CONDI$ = "NumRepa = " & NROREP&
     CODIGENERI% = CODINT%(Control$("REPARA", "CODGENER"))
     NCLIE% = Val(Label11)
     REFE$ = "Reparación - " + TRIM$(Label10)
     If REFE$ = " - " Then REFE$ = ""
     FEMI% = FECHANUM(Text9.TEXT)
     FEX = FEMI%
     REFE$ = REFE$ + " - " + FECHATEX$(FEX)
     LOMA% = 48 - Len(REFE$)
     REFE$ = Left$(TRIM$(ECOARCH$("DEUDOR1", MKI$(NCLIE%))), LOMA%) + REFE$
     
     !REFERENCIA = REFE$
     !FECHEMIS = CVDAT(FEMI%)
     !NROCLIE = NCLIE%
     !RaSoClie = rasocli$(NCLIE%)
        NROOCOM$ = SAFETEXT$(VALOBADA("REPARA", "Or_Compra", CONDI$, "NOMESS"))
     !NROOCOM = NROOCOM$
     !DOMIENT = ""
     !LOCAENT = ""
     !TranspEnt = ""
     !NTranspo = 0
     !NSucent = 0
     !Observa = Text28
     !FeSolEnt = Now
     !Status = STATUPC%              ' Status  de Aprobación
     !ListPre = 0
     !CondPag = 1
        VEND% = XVALO(VALOBADA("DEUDOR12", "Vend_Cli", "Nume_Cli = " & NCLIE%))
     !VENDED = VEND%
     !OV_Motivo = 0
     !OV_AUTORIZA2 = ""
     !OV_AUTORIZA1 = ""
     !PorComVend = 0
     !Rec_ConPa = ""
     !TDescApli = 0
     !VenRep = "R"  ' Venta o Reparacion
     !AnotaPed = ""
     .Update
   End With
   PedEnca.Close
   Set PedEnca = Nothing
   
   ' GRABA TABLA PEDDETA
   
    strSQL = "Select * from PedDeta where NroPed=" & NPP&
    Dim PedDeta As ADODB.Recordset
    Set PedDeta = New ADODB.Recordset
26  On Error Resume Next
    PedDeta.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 26
    End If
    On Error GoTo 0
    With PedDeta
      .AddNew
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
        !NROPED = NPP&
        !FECHEMIS = CVDAT(FEMI%)
        !NROCLIE = NCLIE%
        !RaSoClie = rasocli$(NCLIE%)
        !NROOCOM = NROOCOM$
          FENTRESOL% = FEMI%
        !FeSolEnt = CVDAT(FENTRESOL%)
        !Status = STATUPC%              ' Status  de Aprobación
        !ListPre = 0
        !CondPag = 1
        !VENDED = VEND%
        !PorComVend = 0
        !NuOrItem = 0
        !CODIINT = CODIGENERI%
        !CODIEXT = CODEXT$(CODIGENERI%)
        !SubCod = 0
        !Descrip = DESCRIT0$(CODIGENERI%)
        !TCOLOR = ""
        !TTALLE = ""
        !TLargo = 0
        !TAncho = 0
        !TEspesor = 0
        !VenRep = "R"  ' Venta o Reparacion
        !MoneEmis = 1
        !TIPOCAM = 1
        !PreList = PREUNINE
          CAN = 1
        !CanAsig = CAN
        !CANPED = CAN
          PREUNINE = XVALO(VALOBADA("REPARA", "PrUn_Neto", CONDI$, "NOMESS"))
        !PreUnid = PREUNINE
        !ImpoTot = CAN * PREUNINE
        !CANTENT = 0
        !CanSaldo = CAN
        !FeUltEnt = CVDAT(0)
        !NROREMI = ""
        !TPedItem = ""
        '
        '\\\OT-8-766-19/09/08
        !ImpoAnti = 0
        !FApro_Nope = CVDAT(FECHANUM("01/01/80")) 'PARA EVITAR PROBLEMAS CON FORMATO REGIONAL
        !AcuPrepa = 0
        !AcuDespa = 0
        !TipoPed = "" ' se corroboro que este CAMPO es de  TEXTO
        !FinVigen = CVDAT(FECHANUM("01/01/80"))
        !StaDespa = 0
        !CBase = ""
        !DescLista = 0
        !DescCombi = ""
        !DescRubro = 0
        !DescTotal = 0
        !NUMEROSERIE = "" ' SE ESTABLECE VALOR POR DEFAULT PARA EL NUMERO DE SERIE
        '\\\OT-8-766-FIN
        '
        TTXYZ$ = Space$(76)
        Call REPLA(TTXYZ$, "- " + CODEXT$(CIXI%), 11, 18)
        Call REPLA(TTXYZ$, DESCRIT0$(CIXI%), 29, 38)
        Call REPLA(TTXYZ$, CSTRING$(MKS$(CAN), 3, 10, 1, 2), 57, 10)
        ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
        '
      .Update
    End With
    PedDeta.Close
    Set PedDeta = Nothing
    
    Call AGREGA_ANOTAPED(NPP&, ANOTA$)
   
'   CODIGENERI% = CODINT%(CONTROL$("REPARA", "CODGENER"))
'   NCLIE% = Val(Label11)
'   REFE$ = "Reparación - " + TRIM$(Label10)
'   If REFE$ = " - " Then REFE$ = ""
'   FEX = FEMI%
'   REFE$ = REFE$ + " - " + FECHATEX$(FEX)
'   LOMA% = 48 - Len(REFE$)
'   REFE$ = Left$(TRIM$(ECOARCH$("DEUDOR1", MKI$(NCLIE%))), LOMA%) + REFE$
'   '
'   PDC$ = REGBLAN$("PEDENCA")
'   NROPE& = Val(Label25)
'   Call REPLA(PDC$, MKS$(NROPE&), 1, 4)
'   Call REPLA(PDC$, REFE$, 5, 48)
'   Call REPLA(PDC$, MKI$(FEMI%), 53, 2)
'   Call REPLA(PDC$, MKI$(NCLIE%), 55, 2)
'   Call REPLA(PDC$, TRIM$(Text10.TEXT), 57, 16)
'   ' Call REPLA(PDC$, TRIM$(LENTREGA(0).TEXT), 73, 32)
'   ' Call REPLA(PDC$, TRIM$(LENTREGA(1).TEXT), 105, 32)
'   ' Call REPLA(PDC$, TRIM$(LENTREGA(2).TEXT), 137, 32)
'   Call FRATEXTO(Text28.TEXT, 32)
'   If CARETEX% >= 1 Then Call REPLA(PDC$, RETEX$(1), 169, 32)
'   If CARETEX% >= 2 Then Call REPLA(PDC$, RETEX$(2), 201, 32)
'   If CARETEX% >= 3 Then Call REPLA(PDC$, RETEX$(3), 233, 32)
'   Call REPLA(PDC$, MKI$(0), 265, 2)
'   Call REPLA(PDC$, Chr$(1), 267, 1)
'   Call REPLA(PDC$, Chr$(0), 268, 1)
'   Call REPLA(PDC$, Chr$(1), 269, 1)   ' CONDICION DE PAGO CONTADO
'   Call REPLA(PDC$, Chr$(0), 270, 1)
'   Call REPLA(PDC$, MKS$(1 + ULTREG&("PEDDETA")), 271, 4)
'   Call REPLA(PDC$, "", 275, 14)
'   Call REPLA(PDC$, CAVERI$, 289, 16)
'   Call REGAPP("PEDENCA", PDC$)
'   REGENCA& = ULTREG&("PEDENCA")
'   Call CIERRARCH("PEDENCA")
'   '
'   CAN = 1
'   PDC$ = REGBLAN$("PEDDETA")
'   Call REPLA(PDC$, MKS$(NROPE&), 1, 4)
'   Call REPLA(PDC$, MKI$(FEMI%), 5, 2)
'   Call REPLA(PDC$, MKI$(NCLIE%), 7, 2)
'   Call REPLA(PDC$, TRIM$(Text10.TEXT), 9, 16)
'      FENTRESOL% = FEMI%
'   Call REPLA(PDC$, MKI$(FENTRESOL%), 25, 2)
'   Call REPLA(PDC$, Chr$(1), 27, 1)
'   Call REPLA(PDC$, Chr$(0), 28, 1)
'   Call REPLA(PDC$, Chr$(1), 29, 1)  ' CONDICION DE PAGO CONTADO
'   Call REPLA(PDC$, Chr$(0), 30, 1)
'   Call REPLA(PDC$, MKI$(0), 31, 2)
'   Call REPLA(PDC$, MKI$(CODIGENERI%), 33, 2)
'   Call REPLA(PDC$, MKS$(PREUNINE), 71, 4)
'   Call REPLA(PDC$, MKS$(CAN), 79, 4)
'   Call REPLA(PDC$, MKS$(PREUNINE), 83, 4)
'   Call REPLA(PDC$, MKD$(CDbl(PREUNINE)), 87, 8)
'       '
'   Call REPLA(PDC$, String$(8, 0), 51, 8)
'   Call REPLA(PDC$, "R", 51, 1) ' venta o reparacion
'   '
'   Call REPLA(PDC$, String$(18, 0), 95, 18)
'   Call REPLA(PDC$, MKS$(CAN), 99, 4)
'   Call REPLA(PDC$, MKS$(REGENCA&), 109, 4)  ' puntero a ENCABEZADO
'   Call REPLA(PDC$, CAVERI$, 113, 16)
'   '
'   Call REGAPP("PEDDETA", PDC$)
'   RGX& = ULTREG&("PEDDETA")
'   Call REPLA(PDC$, MKS$(RGX&), 109, 4)  ' puntero a PEDDETA
'   Call REGAPP("PEDDETP", PDC$)
'   '
'   Call CIERRARCH("PEDDETA")
'   Call CIERRARCH("PEDDETP")

End Sub

Sub GENERARESERVA(NROREP$)

    Call BORRAREGISTROS("RESERVA", "IDSUBOR = '" & NROREP$ & "'", REGAF&, "NOMESS")

    SQLX1$ = "SELECT * FROM RESERVA"
    SQLX1$ = SQLX1$ + " WHERE COD_INTE < 1 "
    Dim RESETMP As ADODB.Recordset
    Set RESETMP = New ADODB.Recordset
    RESETMP.Open FILTSQL$(SQLX1$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
    With RESETMP
      KK& = 0
      For UI% = 1 To List1.ListCount
         zz$ = List1.List(UI% - 1)
         CODD$ = TRIM$(Left$(zz$, 16))
         CIIX% = CODINT%(CODD$)
         If CIIX% > 0 Then
           If CIIX% <= NUMAS% Then
             CANTX# = Val(Mid$(zz$, 59, 10))
             .AddNew
             On Error Resume Next
             !CodiEmpr = CodiEmp%
             On Error GoTo 0
             !cod_inte = CIIX%
             !Cod_Exte = CODEXT$(CIIX%)
             !Descrip = DESCRIT$(CIIX%)
             !IdSubOr = NROREP$
             !FechRes = Now
             !CANTRES = CANTX#
             !CantCons = 0
             !Saldo_Entre = CANTX#
             !IDENLOTE = ""
             !DEPOSITO = 0
             !UsoUnit = CANTX#
                KK& = KK& + 1
             !NuOrItem = KK&
             .Update
           End If
         End If
      Next UI%
   End With
   '
   On Error Resume Next
   RESETMP.Close
   Set RESETMP = Nothing
   On Error GoTo 0
   '
End Sub

Private Sub Command1_Click()
  If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLIE% = Val(Label11)
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
   HOII% = HOY(HOS$)
   FEMI% = FECHANUM(Text9.TEXT)
   If FEMI% < 1 Or FEMI% > HOII% Then
      MERRO$ = "Fecha de Emisión no Válida"
      GoTo 98
   End If
   '
   If XVALO(Label2) < 1 Then
        MERRO$ = "Falta N° de Reparación"
        GoTo 98
   End If
   
   If XVALO(Label18) < 1 Then
        MERRO$ = "Falta N° de Presupuesto"
        GoTo 98
   End If
   
   
   CONDI$ = "NumRepa = " & XVALO(Label18)
   
   PREUNINE = VALOBADA("REPARA", "Imp_Tot", CONDI$, "NOMESS")
   Label26 = FORMATNUM$(PREUNINE, "F12.2")
   Call ACTUREGISTRO("REPARA", "NuOrdRep", CONDI$, TRIM$(Label2), REGS&, "NOMESS")
   'texto existente
   textoexistente$ = VALOBADA("REPARA", "RefOrdRep", CONDI$, "NOMESS")
   REFEPRE$ = textoexistente$ & " " & TRIM$(Label10) + " - " + Left$(TRIM$(Label12), 30) + " - R: " + TRIM$(Text10) + " - " + TRIM$(Text9)
   '
   Call ACTUREGISTRO("REPARA", "RefOrdRep", CONDI$, Left$(REFEPRE$, 128), REGS&)
   'Call ACTUREGISTRO("REPARA", "NroPed", CONDI$, Label25, REGS&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "Referen2", CONDI$, Text10, REGS&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "NumeroSerie", CONDI$, Text1, REGS&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "FechOrdRep", CONDI$, FETEXCOR1$(FEMI%), REGS&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "Observa", CONDI$, Text28, REGS&, "NOMESS")
   
   Call COMUNI("Grabación Realizada")
   '
98
99 Command1.Visible = False
   Command14.Visible = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   Call CONECRUN("ARCHIG07", "Padron Maestro de Clientes")
   Command11.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   CONFRE09.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   
   Call GRAOREPA(OKX%)
   If OKX% > 0 Then
     CANCELPRINT% = 0
     Call PRIOREPA
     Call BLANFORMU
     Label18 = ""
     Label2 = ""
     Command14.Enabled = True
     Label11 = ""
     Label10 = ""
     Hide
   End If
   Command14.Enabled = True
End Sub

Sub Command15_Click()
    '
    Command15.Enabled = False
    '
3   CONDI$ = "NuOrdRep > 0 ORDER BY NuOrdRep"
    Call CARGALISEL("PREPAEMI", "REPARA", "NuOrdRep/I10;RefOrdRep/A128", CONDI$)
    Call SELECHO("PREPAEMI/Indice Presupuestos de Reparación Pendientes")
    NPX& = Val(VALACT1$("PREPAEMI"))
    If NPX& < 1 Then GoTo 99
    '
'3   Call SELECHO("ORDEREPA/Indice General Ordenes de Reparación")
'    NPX& = Val(VALACT1$("ORDEREPA"))
'    If NPX& < 1 Then GoTo 99
'    '
'    For UI& = ULTREG&("ORDEREPA") To 1 Step -1
'      XX$ = REGLEIDO$("ORDEREPA", UI&)
'      If CVS(Left$(XX$, 4)) = NPX& Then GoTo 4
'    Next UI&
'    Call COMUNI("Error de Indices - Consulte.")
'    Exit Sub
    '
4   NUDOCU$ = TRIM$(Str$(NPX&))
    TIDOCUM$ = "Orden de Reparación " + NUDOCU$
    TIDOCU2$ = "Procedimiento de Reparación OR-" + NUDOCU$
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
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
    On Error Resume Next
    GoTo 99
    '
5   Call MUESTRADOC(DOCUNU&)
    '
    GoTo 3
    '
99  Command15.Enabled = True
    '
End Sub


Private Sub Command2_Click()
    Label18 = "": Label10 = ""
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
   Command14.Enabled = False
'   Call COPYSTRU("PRESREPA", "PREPAEMI")
'   JJ& = 0
'
'   '
''   CONDI$ = "NROPED > 0 AND STATUS < 9 AND NuOrdRep < 0.5"
'   CONDI$ = "NuOrdRep > 0"
'
''   If InStr(1, EMPREAC$, "NEUQUEN") > 1 Then CONDI$ = "NuOrdRep < 0.5 AND STATUS < 9"
'   JJ& = CantRegistros("REPARA", CONDI$)
'   If JJ& < 1 Then
'     Call COMUNI("No Hay Ordenes de Venta Pendientes\para Emitir Orden de Reparación.")
'     GoTo 99
'   End If
'
'   CONDI$ = CONDI$ + " ORDER BY NUMREPA"
'   Call CARGALISEL("PREPAEMI", "REPARA", "NumRepa/I10;Referen/A128", CONDI$)
'
'   Call SELECHO("PREPAEMI/Indice Ordenes de Venta de Reparación Pendientes")
'   NPX& = Val(VALACT1$("PREPAEMI"))
'   If NPX& < 1 Then GoTo 99
'   '
'   CONDI$ = "NumRepa = " & NPX&
'   NPP& = XVALO(VALOBADA("REPARA", "NroPed", CONDI$))
'   Label18 = TRIM$(Str$(NPX&))
'   Label2 = TRIM$(Str$(NUORDREP&))
'   Label25 = TRIM$(Str$(NPP&))
   If DERACCE%("ACTUREPA", "Actualizar Orden de Reparación") < 2 Then
      Exit Sub
   End If
   
3   CONDI$ = "NuOrdRep > 0 ORDER BY NuOrdRep"
    Call CARGALISEL("PREPAEMI", "REPARA", "NuOrdRep/I10;RefOrdRep/A128", CONDI$)
    Call SELECHO("PREPAEMI/Indice Presupuestos de Reparación Pendientes")
    NPX& = Val(VALACT1$("PREPAEMI"))
    If NPX& < 1 Then GoTo 99
    '

    CONDI$ = " STATUS < 10" 'menor a 10 asi trae todas
    JJ& = CantRegistros("REPARA", CONDI$)
    If JJ& < 1 Then
     Call COMUNI("No Hay Ordenes de Venta Pendientes\para Emitir Orden de Reparación.")
     GoTo 99
    End If
   
'   CONDI$ = CONDI$ + " ORDER BY NUMREPA"
'   Call CARGALISEL("PREPAEMI", "REPARA", "NumRepa/I10;Referen/A128", CONDI$)
   
'   Call SELECHO("PREPAEMI/Indice Ordenes de Venta de Reparación ")
'   NPX& = Val(VALACT1$("PREPAEMI"))
'   If NPX& < 1 Then GoTo 99
   
   CONDI$ = "NuOrdRep = " & NPX& 'NUMERO DE ORDEN DE REPARACION
   NUMREPA& = XVALO(VALOBADA("REPARA", "NumRepa", CONDI$, "NOMESS"))
   
   CONDI$ = "NumRepa = " & NUMREPA& ' NUMERO DE REPARACION
   NPP& = XVALO(VALOBADA("REPARA", "NroPed", CONDI$))


   ORREPA09.Label18 = TRIM$(Str$(NUMREPA&))
   ORREPA09.Label2 = TRIM$(Str$(NPX&))
   ORREPA09.Label25 = TRIM$(Str$(NPP&))
   
   CANCELPRINT% = 0
   Call PRIOREPA
   Call BLANFORMU
   Label18 = ""
   Label2 = ""
   Label11 = ""
   Label10 = ""
   Hide
   Command14.Enabled = True
99 Command4.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("AMASTO07", "Maestro de Artículos de Stock")
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Call SELECHO("SELFECHA")
   Text9.TEXT = VALACT1$("SELFECHA")
   FF% = FECHANUM(Text9.TEXT)
End Sub

Private Sub Form_Load()
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
   Label18 = "": Label10 = ""
   Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Openforms = DoEvents
   Wend
End Sub

Private Sub Label10_Change()
   CODD$ = TRIM$(Label10)
   Label8 = "": TTYZ$ = ""
   Screen.MousePointer = 11
   RNC$ = RECFILT$("DESCREP", 1, CODD$)
   URN& = ULTREG&("DESCREP")
   '
   For U& = 1 To Len(RNC$) Step 4
     RENOTA& = CVS(Mid$(RNC$, U&, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTYY$ = RTrim$(Mid$(REGLEIDO$("DESCREP", RENOTA&), 17))
         If TTYZ$ <> "" Then
           TTYZ$ = TTYZ$ + Chr$(13) + Chr$(10)
         End If
         TTYZ$ = TTYZ$ + RTrim$(TTYY$)
       End If
     End If
   Next U&
   '
   Screen.MousePointer = 1
   LOCADE& = Len(TTYZ$)
   For U& = LOCADE& To 1 Step -1
     If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 30
   Next U&
   U& = 0
30 Label8 = Left$(TTYZ$, U&)
   Label8.Refresh
   '
End Sub

Private Sub Label11_Change()
   '
   Call BLANFORMU
   '
   NC% = Val(Label11)
   If NC% > 0 Then
     If rasocli$(NC%) <> "" Then
       Label12 = rasocli$(NC%)
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
   If NPX& < 1 Then Exit Sub
   '
'   For UI& = ULTREG&("PRESREPA") To 1 Step -1
'     XX$ = REGLEIDO$("PRESREPA", UI&)
'     If CVS(Left$(XX$, 4)) = NPX& Then
'       REPRESUP = TRIM$(Str$(UI&))
'       NCLIE% = CVI(Mid$(XX$, 67, 2))
'       Label11 = TRIM$(Str$(NCLIE%))
'       CODREBO$ = Mid$(XX$, 497, 16)
'       Label10 = CODREBO$
'       Text1 = Mid$(XX$, 71, 15)
'       Text10 = Mid$(XX$, 86, 9)
'       Text28 = TRIM$(Mid$(XX$, 129, 368))
'       Label21 = CSTRING$(Mid$(XX$, 101, 4), 4, 8, 2, 2)
       '
       CONDI$ = "NumRepa = " & NPX& & ""
       Label11 = VALOBADA("REPARA", "Nume_Cli", CONDI$, "NOMESS")
       Label10 = VALOBADA("REPARA", "CodRepa", CONDI$, "NOMESS")
       Text1 = VALOBADA("REPARA", "NumeroSerie", CONDI$, "NOMESS")
       Text10 = VALOBADA("REPARA", "Referen2", CONDI$, "NOMESS")
       Text28 = VALOBADA("REPARA", "Observa", CONDI$, "NOMESS")
       REPRESUP = VALOBADA("REPARA", "NumRepa", CONDI$, "NOMESS")
       Label21 = VALOBADA("REPARA", "HoraMObra", CONDI$, "NOMESS")
       '
       Call CAREPUES
       Call PREREPUES
       '
       On Error Resume Next
       Text10.SetFocus
       On Error GoTo 0
       Openforms = DoEvents
       Screen.MousePointer = 1
'       Exit Sub
'     End If
'   Next UI&
   Screen.MousePointer = 1
   '
End Sub

Private Sub Label6_DblClick()
   If DERACCE%("ACTUREPA", "Actualizar Orden de Reparación") < 2 Then
      Exit Sub
   End If
   
3   CONDI$ = "NuOrdRep > 0 ORDER BY NuOrdRep"
    Call CARGALISEL("PREPAEMI", "REPARA", "NuOrdRep/I10;RefOrdRep/A128", CONDI$)
    Call SELECHO("PREPAEMI/Indice Presupuestos de Reparación Pendientes")
    NPX& = Val(VALACT1$("PREPAEMI"))
    If NPX& < 1 Then GoTo 99
    '

    CONDI$ = " STATUS < 10" 'menor a 10 asi trae todas
    JJ& = CantRegistros("REPARA", CONDI$)
    If JJ& < 1 Then
     Call COMUNI("No Hay Ordenes de Venta Pendientes\para Emitir Orden de Reparación.")
     GoTo 99
    End If
   
'   CONDI$ = CONDI$ + " ORDER BY NUMREPA"
'   Call CARGALISEL("PREPAEMI", "REPARA", "NumRepa/I10;Referen/A128", CONDI$)
   
'   Call SELECHO("PREPAEMI/Indice Ordenes de Venta de Reparación ")
'   NPX& = Val(VALACT1$("PREPAEMI"))
'   If NPX& < 1 Then GoTo 99
   
   CONDI$ = "NuOrdRep = " & NPX& 'NUMERO DE ORDEN DE REPARACION
   NUMREPA& = XVALO(VALOBADA("REPARA", "NumRepa", CONDI$, "NOMESS"))
   
   CONDI$ = "NumRepa = " & NUMREPA& ' NUMERO DE REPARACION
   NPP& = XVALO(VALOBADA("REPARA", "NroPed", CONDI$))


   ORREPA09.Label18 = TRIM$(Str$(NUMREPA&))
   ORREPA09.Label2 = TRIM$(Str$(NPX&))
   ORREPA09.Label25 = TRIM$(Str$(NPP&))
   
   Command1.Visible = True
   Command14.Visible = False
99
   
End Sub

Private Sub List1_Click()
   '
   Call EDITREPU
   '
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
'   NPI% = NPX& Mod 10000
'   RFF$ = RECFILT$("REPUREPA", 2, MKI$(NPI%))
'   '
'   URE& = ULTREG&("REPUREPA")
'30 For KK& = 1 To Len(RFF$) Step 4
'     REG& = CVS(Mid$(RFF$, KK&, 4))
'     If REG& > 0 Then
'       If REG& <= URE& Then
'         XX$ = REGLEIDO$("REPUREPA", REG&)
'         If CVS(Left$(XX$, 4)) = NPX& Then
'           CII% = CVI(Mid$(XX$, 7, 2))
'           CANTI = CVS(Mid$(XX$, 9, 4))
'           '
'           ZZ$ = REGBLAN$("MATEROF")
'           Call REPLA(ZZ$, MKI$(CII%), 1, 2)
'           Call REPLA(ZZ$, UNIMED$(CII%), 3, 2)
'           Call REPLA(ZZ$, MKS$(CANTI), 5, 4)
'           Call REPLA(ZZ$, MKD$(CDbl(CANTI)), 9, 8)
'           Call REPLA(ZZ$, MKS$(0), 17, 4)
'           STDIS = STOCKDIS(CII%)
'           Call REPLA(ZZ$, MKS$(STDIS), 21, 4)
'           FALTA = CANTI - STDIS: If FALTA < 0 Then FALTA = 0
'           Call REPLA(ZZ$, MKS$(FALTA), 25, 4)
'           Call REGAPP("!MATEROF", ZZ$)
'           '
'         End If
'       End If
'     End If
'   Next KK&
   '
30
   Dim OREPAR As ADODB.Recordset
   SQLX$ = "SELECT * FROM REPUREPA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NumRepa = " & NPX& & " AND ORIGEN=1 ORDER BY NuOrd_Item ASC"
   Set OREPAR = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With OREPAR
           Do While Not .EOF
                CII% = !cod_inte
                CANTI = !CANT_SOL
                '
                zz$ = REGBLAN$("MATEROF")
                Call REPLA(zz$, MKI$(CII%), 1, 2)
                Call REPLA(zz$, Unimed$(CII%), 3, 2)
                Call REPLA(zz$, MKS$(CANTI), 5, 4)
                Call REPLA(zz$, MKD$(CDbl(CANTI)), 9, 8)
                Call REPLA(zz$, MKS$(0), 17, 4)
                'STDIS = STOCKDIS(CII%)
                Call REPLA(zz$, MKS$(STDIS), 21, 4)
                FALTA = CANTI - STDIS: If FALTA < 0 Then FALTA = 0
                Call REPLA(zz$, MKS$(FALTA), 25, 4)
                Call REGAPP("!MATEROF", zz$)
                '
                .MoveNext
            Loop
   End With
   OREPAR.Close
   Set OREPAR = Nothing
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
     zz$ = REGLEIDO$("!MATEROF", U&)
     TTXX$ = Space$(128)
     CII% = CVI(Left$(zz$, 2))
     Call REPLA(TTXX$, CODEXT$(CII%), 1, 16)
     Call REPLA(TTXX$, DESCRIT$(CII%), 17, 42)
     NEBRUX$ = Mid$(zz$, 5, 4)
     Call REPLA(TTXX$, CSTRING$(NEBRUX$, 4, 8, 1, 2), 59, 8)
     CANTI = CVS(NEBRUX$)
     COSUX$ = MKS$(STOCKDIS(CII%))
     Call REPLA(zz$, COSUX$, 21, 4)
     Call REPLA(TTXX$, CSTRING$(COSUX$, 4, 10, 1, 2), 67, 10)
     STDIS = CVS(COSUX$)
     FALTA = CANTI - STDIS: If FALTA < 0 Then FALTA = 0
     ITOX$ = MKS$(FALTA)
     Call REPLA(zz$, MKS$(FALTA), 25, 4)
     Call GRAREG("!MATEROF", zz$, U&)
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
   'Command9.SetFocus
End Sub

Sub PRIOREPA()
    '
    FORIPRE$ = TRIM$(Control$("FOREPAR", "DOCUREP2"))
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Imposible Imprimir Orden de Reparación.\Falta Definir Formulario Electrónico.")
       Exit Sub
    End If
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Orden de Reparación.\No existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    CODD$ = Label10: TIETOT = 0
    RFFX$ = RECFILT$("SECOPREP", 1, CODD$)
    For I1& = 1 To Len(RFFX$) Step 4
      i& = CVS(Mid$(RFFX$, I1&, 4))
      X$ = REGLEIDO$("SECOPREP", i&)
      TIEX = CVS(Mid$(X$, 73, 4))
      CAOPS = CVS(Mid$(X$, 97, 4)): If CAOPS < 1 Then CAOPS = 1
      TIETOT = TIETOT + TIEX * CAOPS
    Next I1&
    '
10  NC% = Val(Label11)
    TIPODOC$ = "Orden de Reparación"
    DESTIDOC% = NC%
    NUMEDOC$ = TRIM$(Label2)
    FECHDOC$ = TRIM$(Text9.TEXT)
    '
    CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = TRIM$(Label10)
    CODPARAM$(2) = "DES-CONJ": DOCPARAM$(2) = TRIM$(ECOARCH$("INDIREP", Label10))
    CODPARAM$(3) = "REF-MAT1": DOCPARAM$(3) = TRIM$(Text1)
    CODPARAM$(4) = "MEDIDA-1": DOCPARAM$(4) = TRIM$(Text10)
    CODPARAM$(5) = "REF-MAT2": DOCPARAM$(5) = ""
    CODPARAM$(6) = "REF-MAT3": DOCPARAM$(6) = ""
    CODPARAM$(7) = "PRE-NETO": DOCPARAM$(7) = ""
    CODPARAM$(8) = "BRU-ITEM": DOCPARAM$(8) = ""
    CODPARAM$(9) = "IMP-NETO": DOCPARAM$(9) = ""
    CODPARAM$(10) = "IMP-IVA": DOCPARAM$(10) = ""
    CODPARAM$(11) = "IMP-TOTA": DOCPARAM$(11) = ""
    CODPARAM$(12) = "POR-IVA": DOCPARAM$(12) = ""
    CODPARAM$(13) = "MEDIDA-3": DOCPARAM$(13) = TRIM$(Text10)
    CODPARAM$(14) = "REF-MAT2": DOCPARAM$(14) = TRIM$(Str$(TIETOT))
    CODPARAM$(15) = "COND-PAG": DOCPARAM$(15) = ""
    CODPARAM$(16) = "REF-MAT3": DOCPARAM$(16) = ""
    CODPARAM$(17) = "DES-MPRI": DOCPARAM$(17) = Label8
    CODPARAM$(18) = "MEDIDA-2": DOCPARAM$(18) = TRIM$(Label25)
    CAPARDOC% = 18
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CODTABU1$(4) = "PRE-LIST"
    CACOLTAB1% = 4
    '
    CAITAB1% = 0
    For U& = 1 To List1.ListCount
      zz$ = List1.List(U& - 1)
      CAITAB1% = CAITAB1% + 1
      TETABU1$(1, CAITAB1%) = Left$(zz$, 16)
      CII% = CODINT%(Left$(zz$, 16))
      TETABU1$(2, CAITAB1%) = DESCRIT$(CII%)
      TETABU1$(3, CAITAB1%) = TRIM$(Mid$(zz$, 59, 10))
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
    '
'    FORIPRE$ = TRIM$(CONTROL$("FOREPAR", "DOCUREP0"))
'    NC% = Val(Label11)
'    TIPODOC$ = "Procedimiento de Reparación"
'    DESTIDOC% = NC%
'    NUMEDOC$ = "OR-" + TRIM$(Label2)
'    '
'    CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = TRIM$(Label10)
'    CODPARAM$(2) = "REF-MAT1": DOCPARAM$(2) = TRIM$(Text1)
'    CODPARAM$(3) = "MEDIDA-3": DOCPARAM$(3) = TRIM$(Text10)
'    CODPARAM$(4) = "DES-MPRI": DOCPARAM$(4) = Label8
'    CODPARAM$(5) = "TOT-TIEM": DOCPARAM$(5) = Left$(HORMINSE$(3600 * TIETOT), 6)
'    CAPARDOC% = 5
'    '
'    CODTABU1$(1) = "NOP-OPER"
'    CODTABU1$(2) = "OBS-OPER"
'    CODTABU1$(3) = "TIE-PROD"
'    CACOLTAB1% = 3
'    '
'    CAITAB1% = 0
'    URN& = ULTREG&("OBSOPREP")
'    CODD$ = Label10
'    RFFX$ = RECFILT$("SECOPREP", 1, CODD$)
'    For I1& = 1 To Len(RFFX$) Step 4
'      i& = CVS(Mid$(RFFX$, I1&, 4))
'      X$ = REGLEIDO$("SECOPREP", i&)
'      NOPE% = CVI(Mid$(X$, 17, 2))
'      DOPE$ = Mid$(X$, 19, 48)
'      TIEX = CVS(Mid$(X$, 73, 4))
'      '
'      RNC$ = RECFILT$("OBSOPREP", 1, CODD$)
'      '
'      For U& = 1 To Len(RNC$) Step 4
'        RENOTA& = CVS(Mid$(RNC$, U&, 4))
'        If RENOTA& > 0 Then
'          If RENOTA& <= URN& Then
'            TTZZ$ = REGLEIDO$("OBSOPREP", RENOTA&)
'            NOPY% = CVI(Mid$(TTZZ$, 17, 2))
'            If NOPY% = NOPE% Then
'              DOPE$ = DOPE$ + Mid$(TTZZ$, 19)
'            End If
'          End If
'        End If
'      Next U&
'      '
'      CAITAB1% = CAITAB1% + 1
'      TETABU1$(1, CAITAB1%) = Str$(NOPE%)
'      TETABU1$(2, CAITAB1%) = DOPE$
'      TETABU1$(3, CAITAB1%) = Left$(HORMINSE$(3600 * TIEX), 6)
'      '
'    Next I1&
'    '
'    Call PRINTDOC("@" + FORIPRE$)
    '
    ' AQUI IMPRIMIR ORDEN DE VENTA
    'Call PRIPEDIDO
    '
End Sub

Sub PRIPEDIDO()
    '
    FORIPRE$ = TRIM$(Control$("", "FORPECLI"))
    If FORIPRE$ = "" Then Exit Sub
    '
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Pedido de Cliente.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    CANCELPRINT = 0
    If Control$("PEDCLIEN", "SUPRINT") = "SI" Then
       CANCELPRINT% = 1
    End If
    '
    Screen.MousePointer = 11
    FECHDOC$ = Text9.TEXT
    NUMEDOC$ = TRIM$(Label25)
    While Len(NUMEDOC$) < 6
      NUMEDOC$ = "0" + NUMEDOC$
    Wend
    '
    TIPODOC$ = "Pedido de Cliente"
    If NC% < 1 Then
      TIPODOC$ = "Pedido Interno"
    End If
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CODTABU1$(4) = "PRE-LIST"
    CODTABU1$(5) = "PORDESCU"
    CODTABU1$(6) = "PRE-NETO"
    CODTABU1$(7) = "NET-ITEM"
    CODTABU1$(8) = "FECH-REC"
    CODTABU1$(9) = "ORD-ITEM"
    CACOLTAB1% = 9
    '
    NCLIE% = Val(Label11)
    DESTIDOC% = NCLIE%
    CAITAB1% = 0
    CAITCRU% = 0
    '
    CAITAB1% = 1
    TETABU1$(1, CAITAB1%) = Control$("REPARA", "CODGENER")
    CODIGENERI% = CODINT%(TETABU1$(1, CAITAB1%))
    TETABU1$(2, CAITAB1%) = DESCRIT0$(CODIGENERI%)
    '
    CAPEX$ = "1"
    PRELI$ = Label26
    PREUX$ = Label26
    DESCX$ = "0"
    ITOTX$ = Label26
    FEX = HOY(HOS$)
    FENTX$ = HOS$
    '
    TETABU1$(3, CAITAB1%) = CAPEX$
    TETABU1$(4, CAITAB1%) = PRELI$
    TETABU1$(5, CAITAB1%) = DESCX$
    TETABU1$(6, CAITAB1%) = PREUX$
    TETABU1$(7, CAITAB1%) = ITOTX$
    TETABU1$(8, CAITAB1%) = FENTX$
    TETABU1$(9, CAITAB1%) = Str$(CAITAB1%)
    '
    CAITAB1% = CAITAB1% + 1
    For KKI% = 1 To 9: TETABU1$(KKI%, CAITAB1%) = "": Next KKI%
    TETABU1$(2, CAITAB1%) = "Serie Nro " + Text1
    '
    SUMATOTA# = Val(Label26)
    IDESCUEN# = 0
    '
    CODPARAM$(1) = "IMP-NETO": DOCPARAM$(1) = Str$(SUMATOTA#)
    CODPARAM$(2) = "VENDEDOR":     DOCPARAM$(2) = ""
    CODPARAM$(3) = "COND-PAG":     DOCPARAM$(3) = DECONPA$(1)
    CODPARAM$(4) = "NRO-OCOM":     DOCPARAM$(4) = Text10.TEXT
    CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = ""
    CODPARAM$(6) = "DESTINO": DOCPARAM$(6) = ""
    CODPARAM$(7) = "TRANSPOR": DOCPARAM$(7) = ""
    CODPARAM$(8) = "IMP-DESC": DOCPARAM$(8) = 0
    CODPARAM$(9) = "IMPNEDES": DOCPARAM$(9) = Str$(SUMATOTA#)
    CODPARAM$(10) = "IMP-IVA": DOCPARAM$(10) = Str$(SUMATOTA# * (ALIVA / 100))
    CODPARAM$(11) = "IMP-TOTA": DOCPARAM$(11) = Str$(SUMATOTA# * (1 + ALIVA / 100))
    CODPARAM$(12) = "POR-DESC": DOCPARAM$(12) = ""
    CAPARDOC% = 12
    '
    Call BLANARCH("!OBSERVOF")
    Call FRATEXTO(Text28.TEXT, 32)
    If CARETEX% >= 1 Then Call GRAREG("!OBSERVOF", RETEX$(1), 1)
    If CARETEX% >= 2 Then Call GRAREG("!OBSERVOF", RETEX$(2), 2)
    If CARETEX% >= 3 Then Call GRAREG("!OBSERVOF", RETEX$(3), 3)
    URETE& = CARETEX%
    Call SETULTREG("!OBSERVOF", URETE&)
    Call CIERRARCH("!OBSERVOF")
    '
    FORIPRE$ = TRIM$(Control$("", "FORPECLI"))
    If FORIPRE$ <> "" Then
      Call PRINTDOC("FORPECLI")   ' PEDIDO DE CLIENTE
    End If
    '
    FORIPRE$ = TRIM$(Control$("", "FORDESPA"))
    If FORIPRE$ <> "" Then
      If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") > 0 Then
        TIPODOC$ = "Orden de Despacho"
        CANCELPRINT = 0
        If Control$("ORDESPA", "SUPRINT") = "SI" Then
           CANCELPRINT% = 1
        End If
        '
        If CARETEX% >= 1 Then Call GRAREG("!OBSERVOF", RETEX$(1), 1)
        If CARETEX% >= 2 Then Call GRAREG("!OBSERVOF", RETEX$(2), 2)
        If CARETEX% >= 3 Then Call GRAREG("!OBSERVOF", RETEX$(3), 3)
        URETE& = CARETEX%
        Call SETULTREG("!OBSERVOF", URETE&)
        Call CIERRARCH("!OBSERVOF")
        '
        Call PRINTDOC("FORDESPA")   ' ORDEN DE DESPACHO
      End If
    End If
    '
End Sub

Sub GRAOREPA(OKX%)
   '
   OKX% = 0
'   Call BLOQARCH("ORDEREPA")
'   Call BLOQARCH("RESERVA")
'   Call BLOQARCH("PEDENCA")
'   Call BLOQARCH("PEDDETA")
'   '
'   NPP& = 1
'   FIN& = ULTREG&("PEDENCA")
'   If FIN& > 0 Then
'      For U& = FIN& To FIN& - 16 Step -1
'         If U& < 1 Then GoTo 10
'         XX$ = REGLEIDO$("PEDENCA", U&)
'         NPPX& = CVS(Mid$(XX$, 1, 4))
'         If NPPX& >= NPP& Then
'            NPP& = 1 + NPPX&
'         End If
'      Next U&
'   End If
   '
   'CONDI$ = "NroPed > 0"
   'NPP& = XVALO(VALOBADA("PEDENCA", "Max(NroPed)", CONDI$))
   
'10 ORREPA09.Label25 = TRIM$(Str$(NPP&))
   '
'   REPRESU& = Val(REPRESUP)
'   REG& = 1 + ULTREG&("ORDEREPA")
'   If REPRESU& < 1 Or REPRESU& > ULTREG&("PRESREPA") Then
'      Call COMUNI("Imposible Registrar - Error de Indices")
'      GoTo 98
'   End If
   '
'   XX$ = REGLEIDO$("PRESREPA", REPRESU&)
   '
   Screen.MousePointer = 11
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLIE% = Val(Label11)
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
   HOII% = HOY(HOS$)
   FEMI% = FECHANUM(Text9.TEXT)
   If FEMI% < 1 Or FEMI% > HOII% Then
      MERRO$ = "Fecha de Emisión no Válida"
      GoTo 98
   End If
   '
   CONDI$ = "NumRepa = " & Label18 & ""
   PREUNINE = VALOBADA("REPARA", "Imp_Tot", CONDI$, "NOMESS")
   Label26 = FORMATNUM$(PREUNINE, "F12.2")
   Call ACTUREGISTRO("REPARA", "NuOrdRep", CONDI$, TRIM$(Label2), REGS&, "NOMESS")
   REFEPRE$ = TRIM$(Label10) + " - " + Left$(TRIM$(Label12), 30) + " - R: " + TRIM$(Text10) + " - " + TRIM$(Text9)
   '
   Call ACTUREGISTRO("REPARA", "RefOrdRep", CONDI$, REFEPRE$, REGS&, "NOMESS")
   'Call ACTUREGISTRO("REPARA", "NroPed", CONDI$, Label25, REGS&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "Referen2", CONDI$, Text10, REGS&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "NumeroSerie", CONDI$, Text1, REGS&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "FechOrdRep", CONDI$, FETEXCOR1$(FEMI%), REGS&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "Observa", CONDI$, Text28, REGS&, "NOMESS")
   '
   NOREP$ = TRIM$(Str$(Label2))
   While Len(NOREP$) < 6: NOREP$ = "0" + NOREP$: Wend
   NOREP$ = "OR-" + NOREP$
   '
   Call GENERARESERVA(NOREP$)
   NROREP& = Val(Label18)
'   Call GENERAPEDIDO(NROREP&)
   '
'   RFF$ = RECFILT$("RESERVA", 3, NOREP$)
'   RCC$ = ""
'   URE& = ULTREG&("RESERVA")
'   For KK& = 1 To Len(RFF$) Step 4
'     RGG& = CVS(Mid$(RFF$, KK&, 4))
'     If RGG& > 0 Then
'       If RGG& <= URE& Then
'         XX$ = REGLEIDO$("RESERVA", RGG&)
'         If TRIM$(Mid$(XX$, 3, 12)) = NOREP$ Then
'           RCC$ = RCC$ + MKS$(RGG&)
'         End If
'       End If
'     End If
'   Next KK&
'   '
'   REBLA$ = REGBLAN$("RESERVA")
'   KK& = -3
'   For UI% = 1 To List1.ListCount
'     ZZ$ = List1.List(UI% - 1)
'     CODD$ = TRIM$(Left$(ZZ$, 16))
'     CIIX% = CODINT%(CODD$)
'     If CIIX% > 0 Then
'       If CIIX% <= NUMAS% Then
'         CANTX# = Val(Mid$(ZZ$, 59, 10))
'         USOII = CANTX#
'         YY$ = REBLA$
'         Call REPLA(YY$, MKI$(CIIX%), 1, 2)
'         Call REPLA(YY$, NOREP$, 3, 12)
'         Call REPLA(YY$, MKI$(CIIX%), 15, 2)
'         Call REPLA(YY$, MKI$(FEMI%), 17, 2)
'         Call REPLA(YY$, MKD$(CANTX#), 19, 8)
'         Call REPLA(YY$, MKD$(CANTX#), 35, 8)
'         Call REPLA(YY$, MKS$(USOII), 75, 4)
'         '
'         KK& = KK& + 4
'         If KK& <= Len(RCC$) Then
'             RGG& = CVS(Mid$(RCC$, KK&, 4))
'             Call GRAREG("RESERVA", YY$, RGG&)
'           Else
'             Call REGAPP("RESERVA", YY$)
'         End If
'       End If
'     End If
'   Next UI%
'   '
'   While KK& < Len(RCC$) - 3
'     KK& = KK& + 4
'     RGG& = CVS(Mid$(RCC$, KK&, 4))
'     Call GRAREG("RESERVA", String$(128, 0), RGG&)
'   Wend
'   '
'   Call CIERRARCH("RESERVA")
'   '
'   CODIGENERI% = CODINT%(CONTROL$("REPARA", "CODGENER"))
'   NCLIE% = Val(Label11)
'   REFE$ = "Reparación - " + TRIM$(Label10)
'   If REFE$ = " - " Then REFE$ = ""
'   FEX = FEMI%
'   REFE$ = REFE$ + " - " + FECHATEX$(FEX)
'   LOMA% = 48 - Len(REFE$)
'   REFE$ = Left$(TRIM$(ECOARCH$("DEUDOR1", MKI$(NCLIE%))), LOMA%) + REFE$
'   '
'   PDC$ = REGBLAN$("PEDENCA")
'   NROPE& = Val(Label25)
'   Call REPLA(PDC$, MKS$(NROPE&), 1, 4)
'   Call REPLA(PDC$, REFE$, 5, 48)
'   Call REPLA(PDC$, MKI$(FEMI%), 53, 2)
'   Call REPLA(PDC$, MKI$(NCLIE%), 55, 2)
'   Call REPLA(PDC$, TRIM$(Text10.TEXT), 57, 16)
'   ' Call REPLA(PDC$, TRIM$(LENTREGA(0).TEXT), 73, 32)
'   ' Call REPLA(PDC$, TRIM$(LENTREGA(1).TEXT), 105, 32)
'   ' Call REPLA(PDC$, TRIM$(LENTREGA(2).TEXT), 137, 32)
'   Call FRATEXTO(Text28.TEXT, 32)
'   If CARETEX% >= 1 Then Call REPLA(PDC$, RETEX$(1), 169, 32)
'   If CARETEX% >= 2 Then Call REPLA(PDC$, RETEX$(2), 201, 32)
'   If CARETEX% >= 3 Then Call REPLA(PDC$, RETEX$(3), 233, 32)
'   Call REPLA(PDC$, MKI$(0), 265, 2)
'   Call REPLA(PDC$, Chr$(1), 267, 1)
'   Call REPLA(PDC$, Chr$(0), 268, 1)
'   Call REPLA(PDC$, Chr$(1), 269, 1)   ' CONDICION DE PAGO CONTADO
'   Call REPLA(PDC$, Chr$(0), 270, 1)
'   Call REPLA(PDC$, MKS$(1 + ULTREG&("PEDDETA")), 271, 4)
'   Call REPLA(PDC$, "", 275, 14)
'   Call REPLA(PDC$, CAVERI$, 289, 16)
'   Call REGAPP("PEDENCA", PDC$)
'   REGENCA& = ULTREG&("PEDENCA")
'   Call CIERRARCH("PEDENCA")
'   '
'   CAN = 1
'   PDC$ = REGBLAN$("PEDDETA")
'   Call REPLA(PDC$, MKS$(NROPE&), 1, 4)
'   Call REPLA(PDC$, MKI$(FEMI%), 5, 2)
'   Call REPLA(PDC$, MKI$(NCLIE%), 7, 2)
'   Call REPLA(PDC$, TRIM$(Text10.TEXT), 9, 16)
'      FENTRESOL% = FEMI%
'   Call REPLA(PDC$, MKI$(FENTRESOL%), 25, 2)
'   Call REPLA(PDC$, Chr$(1), 27, 1)
'   Call REPLA(PDC$, Chr$(0), 28, 1)
'   Call REPLA(PDC$, Chr$(1), 29, 1)  ' CONDICION DE PAGO CONTADO
'   Call REPLA(PDC$, Chr$(0), 30, 1)
'   Call REPLA(PDC$, MKI$(0), 31, 2)
'   Call REPLA(PDC$, MKI$(CODIGENERI%), 33, 2)
'   Call REPLA(PDC$, MKS$(PREUNINE), 71, 4)
'   Call REPLA(PDC$, MKS$(CAN), 79, 4)
'   Call REPLA(PDC$, MKS$(PREUNINE), 83, 4)
'   Call REPLA(PDC$, MKD$(CDbl(PREUNINE)), 87, 8)
'       '
'   Call REPLA(PDC$, String$(8, 0), 51, 8)
'   Call REPLA(PDC$, "R", 51, 1) ' venta o reparacion
'   '
'   Call REPLA(PDC$, String$(18, 0), 95, 18)
'   Call REPLA(PDC$, MKS$(CAN), 99, 4)
'   Call REPLA(PDC$, MKS$(REGENCA&), 109, 4)  ' puntero a ENCABEZADO
'   Call REPLA(PDC$, CAVERI$, 113, 16)
'   '
'   Call REGAPP("PEDDETA", PDC$)
'   RGX& = ULTREG&("PEDDETA")
'   Call REPLA(PDC$, MKS$(RGX&), 109, 4)  ' puntero a PEDDETA
'   Call REGAPP("PEDDETP", PDC$)
'   '
'   Call CIERRARCH("PEDDETA")
'   Call CIERRARCH("PEDDETP")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   OKX% = 1
   '
   TTXYZ$ = Space$(92)
   NROREP& = Val(Label18)
   HOII% = HOY(HOS$)
   Call REPLA(TTXYZ$, HOS$, 1, 8)
   Call REPLA(TTXYZ$, USERTER$, 69, 24)
   Call REPLA(TTXYZ$, "Generación Orden de Reparación " + TRIM$(Label2) + ":", 11, 56)
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   '
   TTXYZ$ = Space$(92)
   Call REPLA(TTXYZ$, Space$(56), 11, 56)
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   Call REPLA(TTXYZ$, "Cliente: " + Label12, 11, 56)
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   Call REPLA(TTXYZ$, "    Tel: " + TRIM$(Label16), 11, 56)
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   Call REPLA(TTXYZ$, "    Fax: " + TRIM$(Label17), 11, 56)
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   '
   Call REPLA(TTXYZ$, Space$(56), 11, 56)
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   Call REPLA(TTXYZ$, "  S/Nro: " + Text1, 11, 56)
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   Call REPLA(TTXYZ$, "Referen: " + Text10, 11, 56)
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   Call REPLA(TTXYZ$, "O.Venta: " + TRIM$(Label25), 11, 56)
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   Call REPLA(TTXYZ$, Space$(56), 11, 56)
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   '
   Call FRATEXTO(Text28, 46)
   For KKI% = 1 To CARETEX%
     TTXX$ = Space$(92)
     Call REPLA(TTXX$, RETEX$(KKI%), 20, 46)
     If KKI% = 1 Then Call REPLA(TTXX$, " Observ: ", 11, 9)
'     Call REGAPP("ANOTAREP", Left$(MKS$(NROREP&) + Space$(10) + TTXX$, 96))
     Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, Space$(10) + TTXX$)
   Next KKI%
   '
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, String$(88, "="))
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + String$(88, "="))
   Call CIERRARCH("ANOTAREP")
   '
   Screen.MousePointer = 1
   Exit Sub
   '
98 Call CIERRARCH("ORDEREPA")
   Call CIERRARCH("RESERVA")
   Call CIERRARCH("PEDDETA")
   Call CIERRARCH("PEDDETP")
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

Sub EDITREPU()
   '
   Screen.MousePointer = 11
   JJ& = 0
   For UI% = 1 To List1.ListCount
     zz$ = List1.List(UI% - 1)
     CODD$ = TRIM$(Left$(zz$, 16))
     CIXI% = CODINT%(CODD$)
     If CIXI% > 0 Then
       If CIXI% <= NUMAS% Then
         CANTU = Val(Mid$(zz$, 59, 10))
         XX$ = REGBLAN$("MATEROF")
         Call REPLA(XX$, MKI$(CIXI%), 1, 2)
         Call REPLA(XX$, Unimed$(CIXI%), 3, 2)
         Call REPLA(XX$, MKS$(CANTU), 5, 4)
         Call REPLA(XX$, MKD$(CDbl(CANTU)), 9, 8)
         Call REPLA(XX$, MKS$(0), 17, 4)
         STDIS = STOCKDIS(CIXI%)
         Call REPLA(XX$, MKS$(STDIS), 21, 4)
         FALTA = CANTU - STDIS: If FALTA < 0 Then FALTA = 0
         Call REPLA(XX$, MKS$(FALTA), 25, 4)
         JJ& = JJ& + 1
         Call GRAREG("!MATEROF", XX$, JJ&)
       End If
     End If
   Next UI%
   '
   Call SETULTREG("!MATEROF", JJ&)
   '
20 ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=Repuestos para Reparación"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Frame4.Top + Frame1.Top + 110)) 'Height - ScaleHeight))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(60 + Frame4.Left + Frame1.Left))
   '
   VTB% = VENTABU%("EDITREPU" + ATRI$)
   If VTB% < 0 Then
     Exit Sub
   End If
   Call PREREPUES
   '
End Sub


