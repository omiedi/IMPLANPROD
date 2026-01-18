VERSION 5.00
Begin VB.Form FORMINVEN 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventarios de Materiales - Generación y Consulta"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   9540
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   8505
      ScaleHeight     =   75
      ScaleWidth      =   810
      TabIndex        =   58
      Top             =   5880
      Width           =   865
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   59
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Opciones de Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   210
      TabIndex        =   36
      Top             =   5775
      Width           =   8100
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Dispositivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   210
         TabIndex        =   40
         Top             =   315
         Width           =   3585
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pantalla"
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
            Left            =   840
            TabIndex        =   42
            Top             =   230
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Impresión"
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
            Left            =   2310
            TabIndex        =   41
            Top             =   230
            Width           =   1170
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Valorización"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Left            =   4305
         TabIndex        =   37
         Top             =   315
         Width           =   3585
         Begin VB.OptionButton Option5 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Importes"
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
            Left            =   2205
            TabIndex        =   39
            Top             =   230
            Width           =   1065
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Cantidades"
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
            Left            =   735
            TabIndex        =   38
            Top             =   230
            Value           =   -1  'True
            Width           =   1275
         End
      End
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Proveed"
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
      Left            =   8505
      Picture         =   "FORMINVEN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Alta de Nuevo Item de Stock"
      Top             =   4750
      Width           =   855
   End
   Begin VB.CommandButton Command14 
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
      Left            =   8505
      Picture         =   "FORMINVEN.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   2960
      Width           =   855
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Articulos"
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
      Left            =   8505
      Picture         =   "FORMINVEN.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Alta de Nuevo Item de Stock"
      Top             =   3860
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Productos en Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5475
      Left            =   4305
      TabIndex        =   6
      Top             =   105
      Width           =   4005
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2100
         ScaleHeight     =   585
         ScaleWidth      =   1635
         TabIndex        =   55
         Top             =   4620
         Width           =   1695
         Begin VB.CommandButton Command16 
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
            Picture         =   "FORMINVEN.frx":0896
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Costos Informados"
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
            TabIndex        =   57
            Top             =   105
            Width           =   960
         End
      End
      Begin VB.PictureBox Picture12 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   210
         ScaleHeight     =   585
         ScaleWidth      =   1635
         TabIndex        =   52
         Top             =   4620
         Width           =   1695
         Begin VB.CommandButton Command17 
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
            Picture         =   "FORMINVEN.frx":09E0
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Ajustes   de Stocks"
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
            Width           =   960
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Fábrica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   210
         TabIndex        =   8
         Top             =   420
         Width           =   3585
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2475
            TabIndex        =   30
            Top             =   840
            Width           =   2535
            Begin VB.CommandButton Command6 
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
               Picture         =   "FORMINVEN.frx":0B2A
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Por Número de Orden de Fabricación"
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
               Left            =   840
               TabIndex        =   32
               Top             =   0
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture7 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2475
            TabIndex        =   33
            Top             =   420
            Width           =   2535
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
               Left            =   -20
               Picture         =   "FORMINVEN.frx":0E34
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Totales por Material"
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
               TabIndex        =   35
               Top             =   0
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Proveedores (Consignación)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   210
         TabIndex        =   7
         Top             =   2205
         Width           =   3585
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   43
            Top             =   1260
            Width           =   2325
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
               Left            =   -20
               Picture         =   "FORMINVEN.frx":0F7E
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Materiales Reservados     a Consignar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   750
               Left            =   840
               TabIndex        =   45
               Top             =   0
               Width           =   1380
            End
         End
         Begin VB.PictureBox Picture9 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   46
            Top             =   840
            Width           =   2325
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
               Left            =   -20
               Picture         =   "FORMINVEN.frx":10C8
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Por Material"
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
               TabIndex        =   48
               Top             =   0
               Width           =   855
            End
         End
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   49
            Top             =   420
            Width           =   2325
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
               Left            =   0
               Picture         =   "FORMINVEN.frx":13D2
               Style           =   1  'Graphical
               TabIndex        =   50
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Clasificado     por Proveedor"
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
               TabIndex        =   51
               Top             =   0
               Width           =   1485
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Depósitos Propios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5475
      Left            =   210
      TabIndex        =   3
      Top             =   105
      Width           =   4005
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Consultas y Listados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   210
         TabIndex        =   5
         Top             =   3150
         Width           =   3585
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   17
            Top             =   1260
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
               Picture         =   "FORMINVEN.frx":16DC
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "General Analítico      por Depósito"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   750
               Left            =   840
               TabIndex        =   19
               Top             =   0
               Width           =   1380
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   15
            Top             =   840
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
               Picture         =   "FORMINVEN.frx":19E6
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Parcial por Depósito"
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
               TabIndex        =   20
               Top             =   0
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   12
            Top             =   420
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
               Picture         =   "FORMINVEN.frx":1CF0
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Inventario General"
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
               TabIndex        =   14
               Top             =   0
               Width           =   1485
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Inventario Continuo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   210
         TabIndex        =   4
         Top             =   420
         Width           =   3585
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2475
            TabIndex        =   27
            Top             =   1680
            Width           =   2535
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
               Left            =   -20
               Picture         =   "FORMINVEN.frx":1FFA
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Listado de Cap- tura de Datos"
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
               TabIndex        =   29
               Top             =   105
               Width           =   1590
            End
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
            Height          =   330
            Left            =   315
            TabIndex        =   26
            Text            =   " "
            Top             =   735
            Width           =   2955
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
            Height          =   330
            Left            =   2205
            TabIndex        =   24
            Text            =   " "
            Top             =   315
            Width           =   1065
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2475
            TabIndex        =   9
            Top             =   1260
            Width           =   2535
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
               Picture         =   "FORMINVEN.frx":2304
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Re-Calcular Exis- tencias a Fecha"
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
               TabIndex        =   11
               Top             =   0
               Width           =   1590
            End
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Inventario:"
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
            TabIndex        =   25
            Top             =   420
            Width           =   1905
         End
      End
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
      Left            =   8505
      Picture         =   "FORMINVEN.frx":260E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   250
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
      Height          =   855
      Left            =   8505
      Picture         =   "FORMINVEN.frx":2758
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1155
      Width           =   855
   End
   Begin VB.CommandButton Command8 
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
      Left            =   8505
      Picture         =   "FORMINVEN.frx":2A62
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2050
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   7980
      Picture         =   "FORMINVEN.frx":2D6C
      Top             =   6260
      Width           =   2010
   End
End
Attribute VB_Name = "FORMINVEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command10_Click()
    '
    If Option1.Value = True Then
        Call FINAL("*COMATCO")
      ElseIf Option2.Value = True Then
        Call FINAL("*LIMATCO")
    End If
    '
End Sub

Private Sub Command13_Click()
    Call NEWITEM
    Call FRESHECHO("MASTER")
End Sub

Private Sub Command14_Click()
    Call OPNOIN("")
End Sub

Private Sub Command15_Click()
    Call OPNOIN("")
End Sub

Private Sub Command16_Click()
    Call CONECRUN("AMASTO02", "Maestro de Items de Stock")
End Sub

Private Sub Command17_Click()
    PREPA% = 0
    FORMAJU% = 0
    AJUSTOCK.Show
End Sub

Private Sub Command2_Click()
    Call LIMASTER(2)
End Sub

Private Sub Command3_Click()
    If Option1.Value = True Then
        If Option6.Value = True Then
            Call FINAL("*COINGEN1")
          ElseIf Option5.Value = True Then
            Call FINAL("*COINGEN2")
        End If
      ElseIf Option2.Value = True Then
        Call GENINV3            ' filtrar y listar
    End If
End Sub

Private Sub Command4_Click()
    '
    If Option1.Value = True Then
        If Option6.Value = True Then
            Call FINAL("*COINDEP1")
          ElseIf Option5.Value = True Then
            Call FINAL("*COINDEP2")
        End If
      ElseIf Option2.Value = True Then
        Call GENINV4            ' filtrar y listar
    End If
    '
End Sub

Private Sub Command5_Click()
    '
    If Option1.Value = True Then
        If Option6.Value = True Then
            Call FINAL("*COINANA1")
          ElseIf Option5.Value = True Then
            Call FINAL("*COINANA2")
        End If
      ElseIf Option2.Value = True Then
        Call GENINV5            ' filtrar y listar
    End If
    '
End Sub

Private Sub Command7_Click()
    GENINVE.Show 1
End Sub

Private Sub Command8_Click()
    Call ACCDIR("RECEPMA")
End Sub

Private Sub Form_Load()
    '
    VERANTER$ = "INVENTAR"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = EMPREAC$
    If EPAC$ <> "" Then
      Caption = Caption + " - " + EMPREAC$
    End If
    '
    FECHAINV$ = TRIM$(CONTROL$("INVENTAR", "FECHAINV"))
    If FECHANUM(FECHAINV$) < 1 Then
        FECHAINV$ = ""
    End If
    Text1.TEXT = FECHAINV$
    Text2.TEXT = TRIM$(CONTROL$("INVENTAR", "FEGENINV"))
    Show
    Refresh
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Text1_GotFocus()
    Command7.SetFocus
End Sub

Private Sub Text2_gotfocus()
    Command7.SetFocus
End Sub
