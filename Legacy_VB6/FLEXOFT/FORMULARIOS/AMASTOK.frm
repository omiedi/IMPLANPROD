VERSION 5.00
Begin VB.Form AMASTOK 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   Caption         =   "Archivo Maestro de Items de Stock"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   10875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   5145
      TabIndex        =   10
      Top             =   420
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.ListBox SORTLIST 
      Height          =   1425
      Left            =   5145
      Sorted          =   -1  'True
      TabIndex        =   61
      Top             =   1785
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   3000
      Left            =   5370
      ScaleHeight     =   2940
      ScaleWidth      =   60
      TabIndex        =   59
      Top             =   3530
      Width           =   120
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   5000
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Control"
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
      Left            =   5775
      Picture         =   "AMASTOK.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command20 
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
      Height          =   855
      Left            =   6720
      Picture         =   "AMASTOK.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Direct"
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
      Left            =   7665
      Picture         =   "AMASTOK.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command17 
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
      Height          =   855
      Left            =   8610
      Picture         =   "AMASTOK.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   315
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
      Height          =   855
      Left            =   9555
      Picture         =   "AMASTOK.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Clasificación del Maestro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3585
      Left            =   5775
      TabIndex        =   5
      Top             =   2940
      Width           =   4635
      Begin VB.PictureBox Picture10 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2415
         ScaleHeight     =   585
         ScaleWidth      =   1845
         TabIndex        =   49
         Top             =   1995
         Width           =   1905
         Begin VB.CommandButton Command14 
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
            Left            =   0
            Picture         =   "AMASTOK.frx":0EAA
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Clasificación por Grupos Contables"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Clasificación Contable"
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
            Left            =   630
            TabIndex        =   51
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   1845
         TabIndex        =   46
         Top             =   1995
         Width           =   1905
         Begin VB.CommandButton Command7 
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
            Left            =   0
            Picture         =   "AMASTOK.frx":11B4
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Clasificación del Maestro por Criterios Técnicos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Clasificación Técnica"
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
            Left            =   630
            TabIndex        =   48
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Operación"
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
         Left            =   315
         TabIndex        =   39
         Top             =   2730
         Width           =   4005
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Interactiva"
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
            Left            =   525
            TabIndex        =   41
            Top             =   250
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Por Rango"
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
            Left            =   2415
            TabIndex        =   40
            Top             =   250
            Width           =   1275
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Criterios y Tablas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   315
         TabIndex        =   38
         Top             =   360
         Width           =   4005
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Tabla de Tipos de Material"
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
            TabIndex        =   58
            Top             =   315
            Width           =   2640
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Abrir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   700
            Left            =   3045
            Picture         =   "AMASTOK.frx":14BE
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   525
            Width           =   750
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Criterios Técnicos"
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
            Index           =   7
            Left            =   315
            TabIndex        =   44
            Top             =   565
            Width           =   2640
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Tablas de Clasificación"
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
            Index           =   6
            Left            =   315
            TabIndex        =   43
            Top             =   815
            Width           =   2745
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Grupos Contables"
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
            Index           =   5
            Left            =   315
            TabIndex        =   42
            Top             =   1065
            Width           =   2640
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Actualizacion Base de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6315
      Left            =   420
      TabIndex        =   1
      Top             =   210
      Width           =   4635
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Complementos"
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
         Left            =   315
         TabIndex        =   4
         Top             =   4830
         Width           =   4005
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   1635
            TabIndex        =   30
            Top             =   420
            Width           =   1695
            Begin VB.CommandButton Command10 
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
               Left            =   0
               Picture         =   "AMASTOK.frx":17C8
               Style           =   1  'Graphical
               TabIndex        =   31
               ToolTipText     =   "Acceso a Fórmulas y Estructuras de Producto"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Fórmulas y Estructuras"
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
               Left            =   630
               TabIndex        =   32
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1995
            ScaleHeight     =   585
            ScaleWidth      =   1740
            TabIndex        =   27
            Top             =   420
            Width           =   1800
            Begin VB.CommandButton Command11 
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
               Left            =   0
               Picture         =   "AMASTOK.frx":1AD2
               Style           =   1  'Graphical
               TabIndex        =   28
               ToolTipText     =   "Acceso a Secuencia de Operaciones"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Rutas y Operaciones"
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
               Left            =   630
               TabIndex        =   29
               Top             =   105
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Alta - Modificaciones - Bajas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   315
         TabIndex        =   2
         Top             =   420
         Width           =   4005
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2895
            TabIndex        =   33
            Top             =   3360
            Width           =   2955
            Begin VB.CommandButton Command19 
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
               Left            =   0
               Picture         =   "AMASTOK.frx":1DDC
               Style           =   1  'Graphical
               TabIndex        =   35
               ToolTipText     =   "Baja Lógica Individual"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command18 
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
               Left            =   2340
               Picture         =   "AMASTOK.frx":20E6
               Style           =   1  'Graphical
               TabIndex        =   34
               ToolTipText     =   "Baja Automática por Falta de Movimientos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Baja Lógica de Artículo de Stock"
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
               Left            =   735
               TabIndex        =   36
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2895
            TabIndex        =   23
            Top             =   2625
            Width           =   2955
            Begin VB.CommandButton Command12 
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
               Left            =   2340
               Picture         =   "AMASTOK.frx":23F0
               Style           =   1  'Graphical
               TabIndex        =   26
               ToolTipText     =   "Sustitución de Código de Artículo"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command13 
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
               Left            =   0
               Picture         =   "AMASTOK.frx":253A
               Style           =   1  'Graphical
               TabIndex        =   24
               ToolTipText     =   "Modificación de Atributos Generales"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Modificación de Datos y Atributos"
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
               Left            =   735
               TabIndex        =   25
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1470
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   11
            Top             =   1785
            Width           =   2325
            Begin VB.CommandButton Command2 
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
               Left            =   0
               Picture         =   "AMASTOK.frx":2844
               Style           =   1  'Graphical
               TabIndex        =   12
               ToolTipText     =   "Alta de Nuevo Item por Teclado"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Alta Directa de Nuevo Artículo"
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
               Left            =   735
               TabIndex        =   13
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture9 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   14
            Top             =   1365
            Width           =   2325
            Begin VB.CommandButton Command3 
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
               Left            =   0
               Picture         =   "AMASTOK.frx":298E
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "Copia Completa de Datos de Artículo Existente"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Alta por Copia"
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
               Left            =   735
               TabIndex        =   16
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   17
            Top             =   945
            Width           =   2325
            Begin VB.CommandButton Command5 
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
               Left            =   0
               Picture         =   "AMASTOK.frx":2AD8
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "Importación de Bases de Datos - No Disponible"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Alta por Importación"
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
               Left            =   840
               TabIndex        =   19
               Top             =   0
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   210
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   20
            Top             =   525
            Width           =   2325
            Begin VB.CommandButton Command4 
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
               Left            =   -20
               Picture         =   "AMASTOK.frx":2C22
               Style           =   1  'Graphical
               TabIndex        =   21
               ToolTipText     =   "Alta por Generación - No Disponible"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Generacion de Serie"
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
               Left            =   840
               TabIndex        =   22
               Top             =   0
               Width           =   1170
            End
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Emisión de Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   5775
      TabIndex        =   0
      Top             =   1365
      Width           =   4635
      Begin VB.PictureBox Picture12 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   1845
         TabIndex        =   55
         Top             =   525
         Width           =   1905
         Begin VB.CommandButton Command9 
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
            Left            =   0
            Picture         =   "AMASTOK.frx":2F2C
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Listado de Atributos Generales de Articulos de Stock"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributos Generales"
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
            Left            =   735
            TabIndex        =   57
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2415
         ScaleHeight     =   585
         ScaleWidth      =   1845
         TabIndex        =   52
         Top             =   525
         Width           =   1905
         Begin VB.CommandButton Command15 
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
            Left            =   0
            Picture         =   "AMASTOK.frx":3076
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Listado de captura de Datos para Inventarios"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Captura para Inventarios"
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
            Left            =   630
            TabIndex        =   54
            Top             =   105
            Width           =   1590
         End
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "    Sistema Modular Integrado de Procesamiento de Datos                              Linea FLEXOFT (r) - Version 9.03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3465
      TabIndex        =   3
      Top             =   6825
      Width           =   7365
   End
   Begin VB.Image Image3 
      Height          =   885
      Left            =   0
      Picture         =   "AMASTOK.frx":3380
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   10920
   End
End
Attribute VB_Name = "AMASTOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command10_Click()
    Call CONECRUN("FORMULAS", "")
End Sub

Private Sub COMMAND11_CLICK()
    Command11.Enabled = False
    Call CONECRUN("HORUTAS", "Secuencia de Operaciones Standard por Producto")
    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    ' Call CONECRUN("SUSTICOD", "")
    '
    If Control$("", "AMASTO") <> "NORMAL" Then
       FEMAL$ = TRIM$(Mid$(Control$("", "AMASTO"), 7))
       Call MENSERR(24, "La Base de Datos del Maestro de Items de Stock\no se ha Cerrado Correctamente el " + FEMAL$ + ".\  \Ejecute un Control General del Maestro\para Recuperar los Indices.")
       Exit Sub
    End If
    '
    Call BLOQARCH("AMASTO")
    '
    SUSTICO.Show 1
    Call FRESHECHO("MASTER")
End Sub

Private Sub Command13_Click()
    '
    If Control$("", "AMASTO") <> "NORMAL" Then
       FEMAL$ = TRIM$(Mid$(Control$("", "AMASTO"), 7))
       Call MENSERR(24, "La Base de Datos del Maestro de Items de Stock\no se ha Cerrado Correctamente el " + FEMAL$ + ".\  \Ejecute un Control General del Maestro\para Recuperar los Indices.")
       Exit Sub
    End If
    '
    Call BLOQARCH("AMASTO")
    MODIITEM.Show
End Sub

Private Sub Command14_Click()
    Call OPNOIN("")
End Sub

Private Sub COMMAND15_CLICK()
    Command9.Enabled = False
    Call LIMASTER(2)
    Command9.Enabled = True
End Sub

Private Sub Command16_Click()
    Call ACCDIR("AMASTO")
End Sub

Private Sub Command17_Click()
    Call HELPONLINE("AMASTO")
End Sub

Private Sub Command18_Click()
    Call OPNOIN("")
End Sub

Private Sub Command19_Click()
    '
    If Control$("", "AMASTO") <> "NORMAL" Then
       FEMAL$ = TRIM$(Mid$(Control$("", "AMASTO"), 7))
       Call MENSERR(24, "La Base de Datos del Maestro de Items de Stock\no se ha Cerrado Correctamente el " + FEMAL$ + ".\  \Ejecute un Control General del Maestro\para Recuperar los Indices.")
       Exit Sub
    End If
    '
    Call BLOQARCH("AMASTO")
    '
    Call BAJAMAS
End Sub

Private Sub Command2_Click()
    If Control$("", "AMASTO") <> "NORMAL" Then
       FEMAL$ = TRIM$(Mid$(Control$("", "AMASTO"), 7))
       Call MENSERR(24, "La Base de Datos del Maestro de Items de Stock\no se ha Cerrado Correctamente el " + FEMAL$ + ".\  \Ejecute un Control General del Maestro\para Recuperar los Indices.")
       Exit Sub
    End If
    '
    Call BLOQARCH("AMASTO")
    ALTAITEM.Show
End Sub

Private Sub Command20_Click()
    OPGENPRE.Show 1
End Sub


Private Sub Command22_Click()
    Call BLOQARCH("AMASTO")
    Call REVIMAS
End Sub

Private Sub Command26_Click()
    Call LIESTRU
End Sub

Private Sub Command27_Click()
    Command27.Enabled = False
    Call CONECRUN("HORUTAS", "Secuencia de Operaciones Standard por Producto")
    Command27.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    '
    If Control$("", "AMASTO") <> "NORMAL" Then
       FEMAL$ = TRIM$(Mid$(Control$("", "AMASTO"), 7))
       Call MENSERR(24, "La Base de Datos del Maestro de Items de Stock\no se ha Cerrado Correctamente el " + FEMAL$ + ".\  \Ejecute un Control General del Maestro\para Recuperar los Indices.")
       GoTo 99
    End If
    '
    Call BLOQARCH("AMASTO")
    '
    Call ALTAPORCOPIA
    If SUCCESS% = 1 Then
      Call BLOQARCH("MASTER")
      Call BLOQARCH("INVERT")
      Screen.MousePointer = 11
      '
      NM% = ULTREG&("MASTER") - 1
      If CVI(Left$(REGLEIDO$("MASTER", 1&), 2)) <> NM% Then
          X$ = REGLEIDO$("MASTER", 1&)
          Call REPLA(X$, MKI$(NM%), 1, 2)
          Call GRAREG("MASTER", X$, 1&)
      End If
      '
      Screen.MousePointer = 11
      '
      FINVE% = FILENBR%("INVERT")
      Dim VL As String * 18
      FIECO% = FILENBR%("ECOMAST")
      Dim EL As String * 64
      FICIA% = FILENBR%("CIACTI")
      Dim CL As String * 2
      FIMAS% = FILENBR%("MASTER")
      Dim ML As String * 128
      '       '
10      J1& = 1: J2& = 0: J3& = 1: ATX% = 0
      FIN& = AMASTOK.List1.ListCount
      For II& = 1 To FIN&
          Call TRACE(1, II&, FIN&)
          CODD$ = Mid$(AMASTOK.List1.List(II& - 1), 2, 16)
          CIIX% = Val(Mid$(AMASTOK.List1.List(II& - 1), 18))
          J1& = J1& + 1
          Get #FINVE%, J1&, VL$
          If CODD$ <> Left$(VL$, 16) Then ATX% = 1
          If CIIX% < 0 Then ATX% = 1
          If ATX% = 1 Then
             ZZ1$ = Space$(18)
             Call REPLA(ZZ1$, CODD$, 1, 16)
             Call REPLA(ZZ1$, MKI$(Abs(CIIX%)), 17, 2)
             LSet VL$ = ZZ1$: Put #FINVE%, J1&, VL$
          End If
          '
          If Mid$(CODD$, 16, 1) <> Chr$(96) Then
             If TRIM$(CODD$) <> "" Then
                J2& = J2& + 1
                J3& = J3& + 1
                If J2& > LOF(FIECO%) / 64 Then ATX% = 1
                '
                If ATX% = 1 Then
                   CI% = Abs(CIIX%)
                   REGMA& = CI% + 1
                   Get #FIMAS%, REGMA&, ML$
                   DENO$ = Mid$(ML$, 5, 64)
                   '
                   EC$ = Space$(62) + MKI$(CI%)
                   Call REPLA(EC$, MKI$(CI%), 63, 2)
                   Call REPLA(EC$, CODD$, 1, 16)
                   Call REPLA(EC$, DENO$, 17, 46)
                   '
                   LSet EL$ = EC$: Put #FIECO%, J2&, EL$
                   LSet CL$ = MKI$(CI%): Put #FICIA%, J3&, CL$
                   '
                End If
             End If
          End If
      Next II&
      '
      ZZ$ = String$(18, 0)
      NVI% = AMASTOK.List1.ListCount
      Call REPLA(ZZ$, MKI$(NVI%), 17, 2)
      LSet VL$ = ZZ$: Put #FINVE%, 1, VL$
      Call SETULTREG("INVERT", J1&)
      Call CIERRARCH("INVERT")
      '
      Call SETULTREG("ECOMAST", J2&)
      Call CIERRARCH("ECOMAST")
      '
      NVJ% = J2&
      LSet CL$ = MKI$(NVJ%): Put #FICIA%, 1, CL$
      Call SETULTREG("CIACTI", J3&)
      Call CIERRARCH("CIACTI")
      '
      Call GRACONTROL("", "AMASTO", "NORMAL")
      '
      Call FRESHECHO("MASTER")
      Call FRESHECHO("ECOMAST")
      '
      Call LIBARCH("AMASTO")
      Screen.MousePointer = 1
    End If
    '
99  Command3.Enabled = True
End Sub

Private Sub COMMAND4_CLICK()
    Call OPNOIN("")
End Sub

Private Sub Command5_Click()
   If InStr(UCase$(EMPREAC$), "GRISFER") > 0 Then
        Call IMPORGRI
      Else
        Call OPNOIN("")
   End If
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    If Option1(1) = True Then
        Call CARTATIMA
      ElseIf Option1(5).Value = True Then
        Call CARGRUCOVE
      ElseIf Option1(6).Value = True Then
        Call CARTACLAS
      ElseIf Option1(7).Value = True Then
        Call CARCRITEC
      Else
        Call OPNOIN("")
    End If
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   '
   If Control$("", "AMASTO") <> "NORMAL" Then
       FEMAL$ = TRIM$(Mid$(Control$("", "AMASTO"), 7))
       Call MENSERR(24, "La Base de Datos del Maestro de Items de Stock\no se ha Cerrado Correctamente el " + FEMAL$ + ".\  \Ejecute un Control General del Maestro\para Recuperar los Indices.")
       Exit Sub
   End If
   '
   Call BLOQARCH("AMASTO")
   '
   If Option1(0).Value = True Then
        Call CLASINTER
      ElseIf Option2.Value = True Then
        Call CLASIRANGO
   End If
End Sub

Private Sub Command8_Click()
'Call IMPORMAK
'List2.Clear
'For U& = 30 To 150
'A$ = "A  " + Chr$(U&) + "     " + Str$(U&)
'List2.AddItem A$
'Next U&
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    Call LIMASTER(1)
    Command9.Enabled = True
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    On Error Resume Next
    Show
    On Error GoTo 0
    Screen.MousePointer = 11
    Call CARTAVALISTO
    '
End Sub

Private Sub Option1_DblClick(Index As Integer)
    '
    If Index = 1 Then
        Call CARTATIMA
     ElseIf Index = 5 Then
        Call CARGRUCOVE
     ElseIf Index = 6 Then
        Call CARTACLAS
     ElseIf Index = 7 Then
        Call CARCRITEC
    End If
    '
End Sub
