VERSION 5.00
Begin VB.Form FORECOMO 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Despacho de Materiales "
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFD0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   12000
      Begin VB.CommandButton Command10 
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
         Left            =   10500
         TabIndex        =   5
         Top             =   1155
         Width           =   180
      End
      Begin VB.CommandButton BOTNEXT 
         Caption         =   "Next"
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
         Left            =   10920
         Picture         =   "FORECOMO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Continúa carga de Remito."
         Top             =   735
         Width           =   855
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00E0FFD0&
         BorderStyle     =   0  'None
         Caption         =   "Frame15"
         Height          =   540
         Left            =   10650
         TabIndex        =   67
         Top             =   1050
         Width           =   1170
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00E0FFD0&
         Caption         =   "Variante"
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
         Left            =   6195
         TabIndex        =   54
         Top             =   210
         Width           =   1800
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Consignación"
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
            Top             =   420
            Width           =   1485
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Pedido"
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
            TabIndex        =   55
            Top             =   210
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   210
         ScaleHeight     =   75
         ScaleWidth      =   960
         TabIndex        =   44
         Top             =   1300
         Width           =   1020
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   45
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command12 
         Caption         =   "List"
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
         Left            =   10920
         Picture         =   "FORECOMO.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Listado Resumen de Entregas"
         Top             =   4480
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Kardex"
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
         Left            =   10920
         Picture         =   "FORECOMO.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Ficha de Movimientos de Material"
         Top             =   3780
         Width           =   855
      End
      Begin VB.CommandButton Command3 
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
         Height          =   700
         Left            =   10920
         Picture         =   "FORECOMO.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Notas de Clientes"
         Top             =   2390
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Maestro"
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
         Left            =   10920
         Picture         =   "FORECOMO.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Visualizar Datos de Item de Stock"
         Top             =   3080
         Width           =   855
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
         Height          =   700
         Left            =   10920
         Picture         =   "FORECOMO.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "A-B-M Archivo Maestro de Clientes"
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Trasla."
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
         Left            =   10920
         Picture         =   "FORECOMO.frx":1374
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Opciones Complementarias - Listados y Reportes"
         Top             =   6300
         Width           =   855
      End
      Begin VB.CommandButton Command7 
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
         Left            =   10920
         Picture         =   "FORECOMO.frx":167E
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   5355
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Imprimir"
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
         Height          =   855
         Left            =   10920
         Picture         =   "FORECOMO.frx":1AC0
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Aprueba - Graba e Imprime Remito"
         Top             =   7370
         Width           =   855
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0FFD0&
         Height          =   700
         Left            =   6930
         TabIndex        =   38
         Top             =   5250
         Width           =   3795
         Begin VB.TextBox Text17 
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
            Left            =   2415
            TabIndex        =   8
            Top             =   250
            Width           =   1170
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Declarado:"
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
            Left            =   210
            TabIndex        =   39
            Top             =   285
            Width           =   2115
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0FFD0&
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
         Height          =   2270
         Left            =   6930
         TabIndex        =   37
         Top             =   5985
         Width           =   3795
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
            Height          =   285
            Index           =   4
            Left            =   210
            TabIndex        =   65
            Top             =   1785
            Width           =   3375
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
            Height          =   285
            Index           =   3
            Left            =   210
            TabIndex        =   12
            Top             =   1470
            Width           =   3375
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
            Height          =   285
            Index           =   2
            Left            =   210
            TabIndex        =   11
            Top             =   1155
            Width           =   3375
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
            Height          =   285
            Index           =   1
            Left            =   210
            TabIndex        =   10
            Top             =   840
            Width           =   3375
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
            Height          =   285
            Index           =   0
            Left            =   210
            TabIndex        =   9
            Top             =   525
            Width           =   3375
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0FFD0&
         Height          =   3020
         Left            =   210
         TabIndex        =   35
         Top             =   5250
         Width           =   6525
         Begin VB.Frame Frame14 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Conductor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1170
            Left            =   210
            TabIndex        =   57
            Top             =   1680
            Width           =   2325
            Begin VB.TextBox Text19 
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
               Left            =   105
               TabIndex        =   62
               Top             =   340
               Width           =   2115
            End
            Begin VB.TextBox Text13 
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
               Left            =   630
               TabIndex        =   58
               Top             =   735
               Width           =   1590
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "DNI:"
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
               TabIndex        =   64
               Top             =   800
               Width           =   435
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Orden de Compra"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   210
            TabIndex        =   51
            Top             =   210
            Width           =   2325
            Begin VB.TextBox Text20 
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
               Left            =   315
               MaxLength       =   16
               TabIndex        =   52
               Top             =   315
               Width           =   1695
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Numero de Bultos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   210
            TabIndex        =   50
            Top             =   945
            Width           =   2325
            Begin VB.TextBox Text15 
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
               Left            =   525
               TabIndex        =   53
               Top             =   315
               Width           =   1275
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Transportista"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2640
            Left            =   2625
            TabIndex        =   36
            Top             =   210
            Width           =   3690
            Begin VB.Frame Frame5 
               BackColor       =   &H00E0FFD0&
               Caption         =   "Domicilio"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1130
               Left            =   105
               TabIndex        =   59
               Top             =   1365
               Width           =   3480
               Begin VB.TextBox Text21 
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
                  Left            =   105
                  TabIndex        =   63
                  Top             =   700
                  Width           =   3270
               End
               Begin VB.TextBox Text14 
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
                  Left            =   105
                  TabIndex        =   60
                  Top             =   315
                  Width           =   3270
               End
            End
            Begin VB.Frame Frame10 
               BackColor       =   &H00E0FFD0&
               Caption         =   "Patente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   620
               Left            =   1995
               TabIndex        =   48
               Top             =   735
               Width           =   1590
               Begin VB.TextBox Text18 
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
                  Left            =   105
                  TabIndex        =   49
                  Top             =   210
                  Width           =   1380
               End
            End
            Begin VB.Frame Frame9 
               BackColor       =   &H00E0FFD0&
               Caption         =   "C.U.I.T."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   620
               Left            =   105
               TabIndex        =   46
               Top             =   735
               Width           =   1800
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
                  Height          =   285
                  Left            =   105
                  TabIndex        =   47
                  Top             =   210
                  Width           =   1590
               End
            End
            Begin VB.TextBox Text8 
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
               Left            =   210
               TabIndex        =   7
               Top             =   315
               Width           =   3270
            End
         End
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   5565
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "X"
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   315
         Width           =   435
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
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
         Left            =   10920
         Picture         =   "FORECOMO.frx":212A
         TabIndex        =   24
         ToolTipText     =   "Ababdona la Aplicación"
         Top             =   315
         Width           =   855
      End
      Begin VB.TextBox Text11 
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
         Left            =   9660
         TabIndex        =   4
         Text            =   " "
         Top             =   735
         Width           =   1020
      End
      Begin VB.TextBox Text9 
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
         TabIndex        =   2
         Text            =   " "
         Top             =   315
         Width           =   1170
      End
      Begin VB.CommandButton Command9 
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
         Left            =   10500
         TabIndex        =   3
         Top             =   315
         Width           =   180
      End
      Begin VB.TextBox Text5 
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
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   2850
      End
      Begin VB.TextBox Text4 
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
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   750
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
         Height          =   300
         Left            =   1470
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   3690
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
         Height          =   300
         Left            =   1470
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   " "
         Top             =   315
         Width           =   3690
      End
      Begin VB.CommandButton Command1 
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
         Left            =   1050
         TabIndex        =   1
         Top             =   735
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
         TabIndex        =   0
         Text            =   " "
         Top             =   735
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   3480
         Left            =   210
         ScaleHeight     =   3420
         ScaleWidth      =   10455
         TabIndex        =   25
         Top             =   1680
         Width           =   10515
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C000&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   480
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   10515
            Begin VB.Line Line7 
               X1              =   6300
               X2              =   6300
               Y1              =   0
               Y2              =   1260
            End
            Begin VB.Line Line6 
               X1              =   9345
               X2              =   9345
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line5 
               X1              =   7665
               X2              =   7665
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line4 
               X1              =   1785
               X2              =   1785
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line1 
               X1              =   5670
               X2              =   5670
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "  Codigo                  Descripcion                                           U-M     Cantidad      Lote / Color        N/Ped."
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
               Left            =   0
               TabIndex        =   28
               Top             =   140
               Width           =   10410
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5580
            Left            =   0
            TabIndex        =   26
            Top             =   -1050
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
               Height          =   2985
               Left            =   -30
               TabIndex        =   41
               Top             =   1530
               Visible         =   0   'False
               Width           =   10515
            End
            Begin VB.Line Line10 
               X1              =   0
               X2              =   10500
               Y1              =   1530
               Y2              =   1530
            End
            Begin VB.Line Line2 
               Index           =   1
               X1              =   5670
               X2              =   5670
               Y1              =   525
               Y2              =   5565
            End
            Begin VB.Line Line2 
               Index           =   4
               X1              =   1785
               X2              =   1785
               Y1              =   105
               Y2              =   5145
            End
            Begin VB.Line Line2 
               Index           =   3
               X1              =   9345
               X2              =   9345
               Y1              =   630
               Y2              =   5670
            End
            Begin VB.Line Line2 
               Index           =   2
               X1              =   7665
               X2              =   7665
               Y1              =   525
               Y2              =   5565
            End
            Begin VB.Line Line2 
               Index           =   0
               X1              =   6300
               X2              =   6300
               Y1              =   1260
               Y2              =   6300
            End
         End
      End
      Begin VB.Label Label8 
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
         Height          =   300
         Left            =   8505
         TabIndex        =   66
         Top             =   1155
         Width           =   6000
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         X1              =   210
         X2              =   10710
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         X1              =   10720
         X2              =   10720
         Y1              =   1680
         Y2              =   5160
      End
      Begin VB.Label NORVENTA 
         Caption         =   "Label8"
         Height          =   225
         Left            =   0
         TabIndex        =   61
         Top             =   0
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Documento no Válido como Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5355
         TabIndex        =   34
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   1050
         Width           =   2115
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   7980
         TabIndex        =   23
         Top             =   735
         Width           =   1590
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Almacén:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7140
         TabIndex        =   22
         Top             =   1155
         Width           =   1275
      End
      Begin VB.Label Label3 
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
         Height          =   330
         Left            =   8505
         TabIndex        =   21
         Top             =   315
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         Left            =   315
         TabIndex        =   15
         Top             =   315
         Width           =   960
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Label Label5 
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
      Height          =   330
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "FORECOMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CDESPA$
'
Private Sub BOTNEXT_Click()
    '
    If TRIM$(Label8) = "" Then
      Call MENSERR(24, "Falta Almacén de Entrega")
      Exit Sub
    End If
    '
    Call SIGUEREMI1
    If ULTREG&("!BOLREDET") < 1 Then
        List1.Clear
        List1.Visible = False
        Command5.Enabled = False
        BOTNEXT.Enabled = True
        Text1.SetFocus
        Exit Sub
    End If
    List1.Clear
    List1.Visible = True
    V& = 0
    For U& = 1 To ULTREG&("!BOLREDET")
        X$ = REGLEIDO$("!BOLREDET", U&)
        MEDIDAS$ = Mid$(X$, 193, 12)
        RELI$ = Space$(80)
        CI% = CVI(Mid$(X$, 83, 2))
        If TRIM$(CODEXT$(CI%)) <> "" Then
            V& = V& + 1
            Call GRAREG("!BOLREDET", X$, V&)
            Call REPLA(RELI$, CODEXT$(CI%), 1, 15)
            Call REPLA(RELI$, DESCRIT0$(CI%), 16, 30)
            TXMD$ = TEXTOMEDI$(MEDIDAS$)
            If TXMD$ <> "" Then
               Call REPLA(RELI$, "P.C. " + TXMD$, 16, 30)
            End If
            Call REPLA(RELI$, UNIMED$(CI%), 49, 4)
            CANS$ = CSTRING$(Mid$(X$, 89, 8), 3, 10, 1, 2)
            Call REPLA(RELI$, CANS$, 53, 10)
            Call REPLA(RELI$, Mid$(X$, 107, 12), 66, 12)
            NOPE$ = CSTRING$(Mid$(X$, 151, 4), 3, 7, 0, 2)
            Call REPLA(RELI$, NOPE$, 78, 7)
            List1.AddItem RELI$
        End If
    Next U&
    Call SETULTREG("!BOLREDET", V&)
    '
    Command5.Enabled = True
    Text15.SetFocus
    '
End Sub

Private Sub Command1_Click()
    '
    Call SELECHO("DEUDOR1")
    NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
    If NCXI% > 0 Then
        NC% = NCXI%
        Text1.TEXT = TRIM$(Str$(NC%))
        Text2.TEXT = RASOCLI$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        Call CARTRANSPOR
    End If
    '
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   NCXI% = Val(Text1)
   If ECOARCH("DEUDOR1", MKI$(NCXI%)) = "" Then
      Call MENSERR(24, "Ingrese Número de Cliente")
      GoTo 99
   End If
   '
   RFF$ = RECFILT$("ALMAENT", 3, MKI$(NCXI%))
   JJ& = 0
   For U& = 1 To Len(RFF$) Step 4
     REGE& = CVS(Mid$(RFF$, U&, 4))
     X$ = REGLEIDO$("ALMAENT", REGE&)
     JJ& = JJ& + 1
     Call GRAREG("!ALMAENT", X$, JJ&)
   Next U&
   Call SETULTREG("!ALMAENT", JJ&)
   Call FRESHECHO("!ALMAENT")
   '
   Call SELECHO("!ALMAENT")
   Label8 = VALACT1$("!ALMAENT")
   '
99 Command10.Enabled = True
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    AMACLIBAS.Show 1
    Call FRESHECHO("DEUDOR1")
    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
10 DES% = COMESACT%
   HAS% = HOY(HOS$)
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ <> "" Then
     DES% = CVI(Left$(OBX$, 2))
     HAS% = CVI(Mid$(OBX$, 3, 2))
     If DES% > HAS% Then GoTo 10
     '
     JJ& = 0
     Screen.MousePointer = 11
     Call COPYSTRU("ENTRECOM", "ENTRECOX")
     For U& = 1 To ULTREG&("ENTRECOM")
       X$ = REGLEIDO$("ENTRECOM", U&)
       FF% = CVI(Left$(X$, 2))
       If FF% >= DES% Then
         If FF% <= HAS% Then
           JJ& = JJ& + 1
           Call GRAREG("ENTRECOX", X$, JJ&)
         End If
       End If
     Next U&
     Call SETULTREG("ENTRECOX", JJ&)
     '
     If JJ& < 1 Then
        Call COMUNI("Sin Movimientos")
      Else
        Call FINAL("*ENTRECOM")
     End If
     Screen.MousePointer = 1
   End If
   '
   Command12.Enabled = True
End Sub

Private Sub Command2_Click()
    '
    If ULTREG&("!BOLREDET") > 0 Then
        Call CIERRARCH("!BOLREDET")
        Call BLANARCH("!BOLREDET")
        List1.Clear
        List1.Visible = False
        Command5.Enabled = False
        BOTNEXT.Enabled = True
        Text1.SetFocus
       Else
        Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        Hide
    End If
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NC% = Val(TRIM$(Text1.TEXT))
    If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
       Call NOTACLI(NC%, 2)
    End If
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call SELECHO("MASTER")
    ARGUBUS$ = VALACT1$("MASTER")
    If CODINT%(ARGUBUS$) > 0 Then
      Call MUESTRADATCOD(ARGUBUS$)
    End If
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    For U& = 1 To 4
       TTXXX$ = TRIM$(Text16(U& - 1).TEXT)
       Call GRAREG("!OBSERVOF", TTXXX$, U&)
    Next U&
    Call SETULTREG("!OBSERVOF", 4)
    Call CIERRARCH("!OBSERVOF")
    '
    Call PRIREMIFAC
    'DEDOC$ = TIPODOC$ + " " + NUMEDOC$ + " " + TRIM$(Text1.TEXT) + "-" + TRIM$(Text2.TEXT)
    'Call SAVESPOOL(DEDOC$, OKX%)      ' guarda el spool de impresion
    '
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    DP1% = Val(CONTROL$("", "DEPOEXPE"))
    '
    NROPRO& = 0
    Call NUREMITO(NROPRO&)
    Text11.TEXT = CDESPA$ + "." + TRIM$(Str$(NROPRO&))
    '
    Call CIERRARCH("!BOLREDET")
    Call BLANARCH("!BOLREDET")
    List1.Clear
    List1.Visible = False
    Command5.Enabled = False
    BOTNEXT.Enabled = True
    '
End Sub

Private Sub Command6_Click()
    '
    Command6.Enabled = False
    If COMALTER%("Genera Archivo de Transferencia\a la Central Avellaneda ?\\Cancelar\Generar Transferencia") = 2 Then
      Call GENARTRA
    End If
    Command6.Enabled = True
    '
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    OPREMICOMO.Show 1
    '
    NROPRO& = 0
    Call NUREMITO(NROPRO&)
    Text11.TEXT = CDESPA$ + "." + TRIM$(Str$(NROPRO&))
    '
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    'Call CONECRUN("*KARDEX", "Ficha de Movimientos por Pieza")
    Call GENKARDEX("")
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
    End If
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
10  If ULTREG("ECODEP") < 1 Then
        Call MENSERR(24, "Falta Definir Tabla de Depósitos")
        Call CARTADEPOSI
        GoTo 10
    End If
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIT0"))
    If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Remito.")
        OPREMICOMO.Show 1
        GoTo 10
    End If
    '
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito de Facturación:\  \'" + RMCC$ + "'.")
            OPREMICOMO.Show 1
            GoTo 10
        End If
    End If
    '
     FORIPRE$ = TRIM$(CONTROL$("", "FOREMIT1"))
    If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Nota de Suministro.")
        OPREMICOMO.Show 1
        GoTo 10
    End If
    '
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Nota de Suministro:\  \'" + RMCC$ + "'.")
            OPREMICOMO.Show 1
            GoTo 10
        End If
    End If
    '
   FF% = HOY(HO$)
    Text9.TEXT = HO$
    '

    CDESPA$ = "RT"
    NROPRO& = 0
    Call NUREMITO(NROPRO&)
    Text11.TEXT = CDESPA$ + "." + TRIM$(Str$(NROPRO&))
    '
    Refresh
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   On Error Resume Next
   Hide
   On Error GoTo 0
End Sub

Private Sub Label2_DblClick()
    '
    OPREMICOMO.Show 1
    '
    NROPRO& = 0
    Call NUREMITO(NROPRO&)
    Text11.TEXT = CDESPA$ + "." + TRIM$(Str$(NROPRO&))
    '
End Sub

Private Sub Label6_DblClick()
   If TRIM$(UCase$(Label6)) = "REMITO:" Then
       Label6.ForeColor = RGB(255, 0, 0)
       Label6 = "N.Recepción:"
       CDESPA$ = "NR"
     Else
       Label6.ForeColor = RGB(0, 0, 255)
       Label6 = "Remito:"
       CDESPA$ = "RT"
   End If
   NROPRO& = 0
   Call NUREMITO(NROPRO&)
   Text11.TEXT = CDESPA$ + "." + TRIM$(Str$(NROPRO&))
End Sub

Private Sub List1_DblClick()
    '

    Call SIGUEREMI2
    If ULTREG&("!BOLREDET") < 1 Then
        List1.Clear
        List1.Visible = False
        Command5.Enabled = False
        BOTNEXT.Enabled = True
        Text1.SetFocus
        Exit Sub
    End If
    List1.Clear
    List1.Visible = True
    V& = 0
    For U& = 1 To ULTREG&("!BOLREDET")
        X$ = REGLEIDO$("!BOLREDET", U&)
        RELI$ = Space$(80)
        CI% = CVI(Mid$(X$, 83, 2))
        If TRIM$(CODEXT$(CI%)) <> "" Then
            V& = V& + 1
            Call GRAREG("!BOLREDET", X$, V&)
            Call REPLA(RELI$, CODEXT$(CI%), 1, 15)
            Call REPLA(RELI$, DESCRIT0$(CI%), 16, 30)
            Call REPLA(RELI$, UNIMED$(CI%), 49, 4)
            CANS$ = CSTRING$(Mid$(X$, 89, 8), 3, 10, 1, 2)
            Call REPLA(RELI$, CANS$, 53, 10)
            Call REPLA(RELI$, Mid$(X$, 107, 12), 66, 12)
            NOPE$ = CSTRING$(Mid$(X$, 151, 4), 3, 7, 0, 2)
            Call REPLA(RELI$, NOPE$, 78, 7)
            List1.AddItem RELI$
        End If
    Next U&
    Call SETULTREG("!BOLREDET", V&)
    '
    Command5.Enabled = True
    Text15.SetFocus
    '
End Sub

Private Sub Text1_Change()
   Label8 = ""
End Sub

Private Sub Text1_DblClick()
    '
    Call SELECHO("DEUDOR1")
    NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
    If NCXI% > 0 Then
        NC% = NCXI%
        Text1.TEXT = TRIM$(Str$(NC%))
        Text2.TEXT = RASOCLI$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        Call CARTRANSPOR
    End If
    '
End Sub

Private Sub Text1_GotFocus()
    Static CTXX%
    If CTXX% = 0 Then
        Refresh
        ' MsgBox "AQUI"
        CTXX% = 1
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Val(Text1.TEXT) > 0 Then
          If Val(Text1.TEXT) <= 32767 Then
             NC% = Val(Text1.TEXT)
             If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
                Text2.TEXT = RASOCLI$(NC%)
                Text3.TEXT = DOMICLI$(NC%)
                Text4.TEXT = CPOSCLI$(NC%)
                Text5.TEXT = LOCACLI$(NC%)
                Call CARTRANSPOR
                BOTNEXT.SetFocus
                Exit Sub
             End If
          End If
       End If
       '
       Beep
       Text1.TEXT = ""
       '
    End If
    '
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("ECODEP")
    DPX1% = Val(VALACT1$("ECODEP"))
    If DPX1% > 0 Then
        DP1% = DPX1%
        Text10.TEXT = TRIM$(Str$(DP1%))
        Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    End If
End Sub

Private Sub Text11_LostFocus()
    '
    VDEFF$ = TRIM$(Text11.TEXT)
    Text11.TEXT = VDEFF$
    If Left$(VDEFF$, 3) <> CDESPA$ + "." Then GoTo 90
    NRONUE& = Val(Mid$(VDEFF$, 4))
    NROPRO& = NRONUE&
    Call NUREMITO(NROPRO&)
    If NROPRO& <> NRONUE& Then
        Beep
        GoTo 95
    End If
    Exit Sub
    '
90  Beep
    NROPRO& = 0
    Call NUREMITO(NROPRO&)
95  Text11.TEXT = CDESPA$ + "." + TRIM$(Str$(NROPRO&))
    Text11.SetFocus
    '
End Sub

Private Sub Text2_GotFocus()
    Text1.SetFocus
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

Private Sub Text6_Click()
    Text1.SetFocus
End Sub

Private Sub Text6_DblClick()
    OPREMICOMO.Show 1
    '
    NROPRO& = 0
    Call NUREMITO(NROPRO&)
    Text11.TEXT = CDESPA$ + "." + TRIM$(Str$(NROPRO&))
    '
End Sub

Private Sub Text7_GotFocus()
    Text10.SetFocus
End Sub

Private Sub Text9_DblClick()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
    End If
End Sub

Private Sub Text9_LostFocus()
    VDEVV$ = TRIM$(Text9.TEXT)
    If FECHANUM(VDEVV$) > 0 Then
          FF% = FECHANUM(VDEVV$)
        Else
          Call MENSERR(24, "Fecha no Válida")
          FEX = HOY(HO$)
          Text9.TEXT = HO$
          Text9.SetFocus
    End If
End Sub

Sub CARTRANSPOR()
    '
    NCLIE% = Val(Text1.TEXT)
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
    Text8.TEXT = TRIM$(Mid$(LENTREG$, 65, 32))
    '
End Sub

Sub SIGUEREMI1()
    '
    NCXI = Val(Text1.TEXT)
    If NCXI < 1 Or NCXI > 32767 Then
        Call MENSERR(24, "Numero de Cuenta no Válido")
        Text1.SetFocus
        Exit Sub
    End If
    '
    NC% = NCXI
    If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
        Call MENSERR(24, "Cuenta no Figura\en Base de Datos de Clientes")
        Text1.SetFocus
        Exit Sub
    End If
    '
    FE$ = TRIM$(Text9.TEXT)
    FF% = FECHANUM(FE$)
    If FF% < 1 Or FF% > HOY(HO$) + 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text9.SetFocus
        Exit Sub
    End If
    '
    Call NOTACLI(NC%, 1)
    '
    Call CIERRARCH("*.*")
    Call BLANARCH("!BOLREDET")
    '
    If InStr(UCase$(EMPREAC$), "SABO") > 0 Then GoTo 10
    '
    FIN& = ULTREG&("PEDDETA")
    RPEDX$ = RECFILT$("PEDDETA", 3, MKI$(NC%))
    NROPEDA& = 0
    FIN& = Len(RPEDX$)
    For UU& = 1 To Len(RPEDX$) Step 4
        U& = CVS(Mid$(RPEDX$, UU&, 4))
        If FIN& > 0 Then
            Call TRACE(20, UU&, FIN&)
        End If
        X$ = REGLEIDO$("PEDDETA", U&)
        CIPD% = CVI(Mid$(X$, 33, 2))
        If CVI(Mid$(X$, 7, 2)) = NC% Then
          Status% = Asc(Mid$(X$, 27, 1))
          If Status% <= 1 Then
            CAN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
            If CAN > 0.05 Then
              NROPED& = CVS(Left$(X$, 4))
              MEDIDAS$ = Mid$(X$, 59, 12)
              Y$ = REGBLAN$("BOLREDET")
              Call REPLA(Y$, MKI$(CIPD%), 83, 2)
              Call REPLA(Y$, MKD$(0), 89, 8)
              Call REPLA(Y$, MKS$(NROPED&), 151, 4)
              Call REPLA(Y$, MKD$(CDbl(CAN)), 163, 8)
              Call REPLA(Y$, MKS$(U&), 189, 4)
              Call REPLA(Y$, MEDIDAS$, 193, 12)
              Call REGAPP("!BOLREDET", Y$)
            End If
          End If
        End If
    Next UU&
    '
10  BOTNEXT.Enabled = False
    Call SIGUEREMI2
    '
End Sub
'

Sub SIGUEREMI2()
    '
    ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Remito de Facturación - Detalle de Material a Despachar"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
    '
    ARCHIPRE$ = "MASTER"
    LPREX% = Val(CONTROL$("", "LIPREXPE"))
    If LPREX% > 0 Then
      If LPREX% < 256 Then
        If ECOARCH$("LISTAS", Chr$(LPREX%)) <> "" Then
          ARCHIPRE$ = TRIM$(Str$(LPREX%))
          While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
          ARCHIPRE$ = "LIPRE" + ARCHIPRE$
        End If
      End If
    End If
    '
    VERIST% = (-1)
    If CONTROL$("REMSTAN", "VERISTOK") = "VERISTOK" Then
       VERIST% = 1
    End If
    '
    AVENTA$ = "DETAREMI"
    If EXISTE%("DETRENQN.PRF") > 0 Then AVENTA$ = "DETRENQN"
    '
    ATRI$ = ATRI$ + "/ARCHECO(1)=" + ARCHIPRE$
10  VTB% = VENTABU%(AVENTA$ + ATRI$)
    If VTB% < 0 Then
       If COMALTER%("Realmente Desea Cancelar Este Remito ?\\No, Continuar\Si, Cancelar") <> 2 Then GoTo 10
       Call BLANARCH("!BOLREDET")
       Exit Sub
    End If
    '
    ATXX% = 0
    FIN& = ULTREG&("!BOLREDET")
    For U& = 1 To FIN&
       YY$ = REGLEIDO$("!BOLREDET", U&)
       CI% = CVI(Mid$(YY$, 83, 2))
       CAN# = CVD(Mid$(YY$, 89, 8))
       CAPEPE# = CVD(Mid$(YY$, 163, 8))
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           If CAN# > 0.05 Then
             If CAN# > CAPEPE# + 0.05 Then
               ' redistribuir
               For JJ& = U& + 1 To FIN&
                 ZZ$ = REGLEIDO$("!BOLREDET", JJ&)
                 If CVI(Mid$(ZZ$, 83, 2)) = CI% Then
                   CANUEN# = CAN# - CAPEPE# + CVD(Mid$(ZZ$, 89, 8))
                   Call REPLA(ZZ$, MKD$(CANUEN#), 89, 8)
                   Call GRAREG("!BOLREDET", ZZ$, JJ&)
                   YY$ = REGLEIDO$("!BOLREDET", U&)
                   Call REPLA(YY$, MKD$(CAPEPE#), 89, 8)
                   Call GRAREG("!BOLREDET", YY$, U&)
                   ATXX% = 1
                   GoTo 19
                 End If
               Next JJ&
             End If
           End If
         End If
       End If
19  Next U&
    '
    If ATXX% > 0 Then
      Call COMUNI("FLEXOFT ha Re-Distribuido la Entrega\Según los Pedidos Pendientes.\  \Pulse 'Aceptar' para ver la Distribución.")
      GoTo 10
    End If
    '
    UJ& = 0: AJUCAN% = 0
    NORVENTA = ""
    For U& = 1 To ULTREG&("!BOLREDET")
       YY$ = REGLEIDO$("!BOLREDET", U&)
       REPED& = CVS(Mid$(YY$, 189, 4))
       CI% = CVI(Mid$(YY$, 83, 2))
       CAN# = CVD(Mid$(YY$, 89, 8))
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           If CAN# > 0.05 Then
             Call VERISAL(CI%)
             Call VESARRAS(CI%)
             CADISPONI# = STODEPOS(CI%, DEPOEXPE%) - STOCKASI(CI%)
             If CADISPONI# < 0 Then CADISPONI# = 0
             If CADISPONI# - CAN# < (-0.005) Then
               CONTI% = COMALTER%("ATENCION !!!\Stock Insuficiente en Código '" + TRIM$(CODEXT$(CI%)) + "'\" + DESCRIT0$(CI%) + "\\Ajustar\Continuar\Abortar")
               If CONTI% <= 1 Then ' ajusta cantidad
                    AJUCAN% = 1
                    Call REPLA(YY$, MKD$(CADISPONI#), 89, 8)
                   ElseIf CONTI% = 3 Then
                    Call SETULTREG("!BOLREDET", 0)
                    Call CIERRARCH("!BOLREDET")
                    Exit Sub
               End If
             End If
             UJ& = UJ& + 1
             Call GRAREG("!BOLREDET", YY$, UJ&)
             '
             NORVE& = CVS(Mid$(YY$, 151, 4))
             If NORVE& > 0 Then
               If NORVE& <> NORVEA& Then
                 NORVEA& = NORVE&
                 If NORVENTA <> "" Then NORVENTA = NORVENTA + " - "
                 NORVENTA = NORVENTA + TRIM$(Str$(NORVE&))
               End If
             End If
             '
           End If
         End If
       End If
    Next U&
    '
    If UJ& < ULTREG&("!BOLREDET") Or AJUCAN% > 0 Then
       Call SETULTREG("!BOLREDET", UJ&)
       Call CIERRARCH("!BOLREDET")
       GoTo 10
    End If
    '
End Sub
'
Sub PRIREMIFAC()
    '
    DEPOEX% = 1
    CMXX$ = "0": If CDESPA$ = "NR" Then CMXX$ = "1"
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIT" + CMXX$))
    If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de " + TRIM$(Label6))
        Exit Sub
    End If
    '
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de " + TRIM$(Label6) + ":\  \'" + RMCC$ + "'.")
            Exit Sub
        End If
    End If
    '
    CONTA% = ULTREG&("!BOLREDET")
    If CONTA% < 1 Then
        Call COMUNI("No se ha Ingresado Detalle")
        Exit Sub
    End If
    '
    VERIST% = (-1)
    If CONTROL$("REMSTAN", "VERISTOK") = "VERISTOK" Then
       VERIST% = 1
    End If
    '
    REPLAORI% = 0
    If CONTROL$("OCOMPRA", "REPLAORI") = "SI" Then
       REPLAORI% = 1
    End If
    '
    If VERIST% = 1 Then
       For YX& = 1 To ULTREG&("!BOLREDET")
          YY$ = REGLEIDO$("!BOLREDET", YX&)
          CI% = CVI(Mid$(YY$, 83, 2))
          CAN# = CVD(Mid$(YY$, 89, 8))
          LOTE$ = Mid$(YY$, 107, 12)
          If CI% > 0 Then
             If CI% <= NUMAS% Then
                If CAN# > 0.005 Then
                   If STODEPOS(CI%, DEPOEX%) - CAN# < (-0.005) Then
                      CONTI% = COMALTER%("ATENCION !!!\Stock Insuficiente en Código '" + TRIM$(CODEXT$(CI%)) + "'\" + DESCRIT0$(CI%) + "\\Cancelar\Continuar")
                      If CONTI% <> 2 Then
                         Exit Sub
                      End If
                   End If
                End If
             End If
          End If
       Next YX&
    End If
    '
    Screen.MousePointer = 11      ' Registrando Remito
    '
    AINDI% = 0
    '
    NC% = Val(Text1.TEXT)
    REMI$ = TRIM$(Text11.TEXT)
    NROO& = Val(Mid$(REMI$, 4))
    FF% = FECHANUM(Text9.TEXT)
    DEPO% = DEPOEX%
    '
    PREFIREM$ = CONTROL$("", "PREFIRE" + CMXX$)
    If PREFIREM$ = "" Then PREFIREM$ = "0001"
    NROREMI$ = TRIM$(Mid$(REMI$, 4))
    While Len(NROREMI$) < 8: NROREMI$ = "0" + NROREMI$: Wend
    NROREMI$ = CDESPA$ + "." + PREFIREM$ + "-" + NROREMI$
    '
    HO% = HOY(HOS$)
    HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
    User% = USNBR% Mod 100
    CTRXX$ = MKI$(HO%) + MKI$(Int(HP)) + Chr$(User%) + String$(2, 0)
    '
    RECONSI$ = REGBLAN$("ENTRECOM")
    Call REPLA(RECONSI$, MKI$(FF%), 1, 2)
    Call REPLA(RECONSI$, REMI$, 3, 14)
      TIPODO$ = TRIM$(Option1.Caption)
      If Option2.Value = True Then TIPODO$ = TRIM$(Option2.Caption)
    Call REPLA(RECONSI$, TIPODO$, 17, 16)
    Call REPLA(RECONSI$, MKI$(NC%), 33, 2)
    Call REPLA(RECONSI$, RASOCLI$(NC%), 35, 30)
    Call REPLA(RECONSI$, Label8, 65, 32)
    Call REPLA(RECONSI$, MKI$(1), 191, 2)
    Call REPLA(RECONSI$, MKI$(0), 193, 2)
    IDOCUM& = 10000& * (FF% Mod 1000) + 100& * (NROO& Mod 100)
    '
    YXI% = 0
    For YX& = 1 To ULTREG&("!BOLREDET")
      YY$ = REGLEIDO$("!BOLREDET", YX&)
      CI% = CVI(Mid$(YY$, 83, 2))
      CAN# = CVD(Mid$(YY$, 89, 8))
      LOTE$ = Mid$(YY$, 107, 12)
      MEDIDAS$ = Mid$(YY$, 193, 12)
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          If CAN# > 0.005 Then
            NUOC& = CVS(Mid$(YY$, 151, 4))
            DOSEC$ = ""
            CANTI = (-1) * CAN#
            '
            DEPX% = DEPO%: If DEPO% = 99 Then DEPX% = DEPOPRE%(CI%)
            Call MOVISTO(FF%, CI%, LOTE$, CDESPA$, REMI$, DOSEC$, RASOCLI$(NC%), 0, "$", NC%, DEPX%, CANTI)
            '
            RTRASY$ = RECONSI$
            IDOCUM& = IDOCUM& + 1
            Call REPLA(RTRASY$, MKS$(IDOCUM&), 97, 4)
            Call REPLA(RTRASY$, TRIM$(Text20) + "/" + TRIM$(LOTE$), 101, 16) ' ORDEN DE COMPRA
            Call REPLA(RTRASY$, MKI$(CI%), 117, 2)
            Call REPLA(RTRASY$, CODEXT$(CI%), 119, 16)
            Call REPLA(RTRASY$, DESCRIT0$(CI%), 135, 48)
            Call REPLA(RTRASY$, MKD$(CAN#), 183, 8)
            Call REGAPP("ENTRECOM", RTRASY$)
            '
          End If
        End If
      End If
    Next YX&
    '
    ' blanquear pedidos pendientes
    For YX& = 1 To ULTREG&("!BOLREDET")
        YY$ = REGLEIDO$("!BOLREDET", YX&)
        CI% = CVI(Mid$(YY$, 83, 2))
        CAN# = CVD(Mid$(YY$, 89, 8))
        LOTE$ = Mid$(YY$, 107, 12)
        MEDIDAS$ = Mid$(YY$, 193, 12)
        REGPEDPE& = CVS(Mid$(YY$, 189, 4))
        '
        If REGPEDPE& > 0 Then
          If REGPEDPE& <= ULTREG&("PEDDETA") Then
            CIENT% = CI%
            CAENTRE = CAN#
            Y$ = REGLEIDO$("PEDDETA", REGPEDPE&)
            If CVI(Mid$(Y$, 33, 2)) = CIENT% Then ' debe coincidir codigo
              TOENTRE = CAENTRE + CVS(Mid$(Y$, 95, 4))
              SALPEN = CVS(Mid$(Y$, 79, 4)) - TOENTRE: If SALPEN < 0 Then SALPEN = 0
              '
              CASIG = 0: If Mid$(Y$, 75, 4) <> "    " Then CASIG = CVS(Mid$(Y$, 75, 4))
              CASIG = CASIG - CAENTRE: If CASIG < 0 Then CASIG = 0
              Call REPLA(Y$, MKS$(CASIG), 75, 4)
              '
              Call REPLA(Y$, MKS$(TOENTRE), 95, 4)
              Call REPLA(Y$, MKS$(SALPEN), 99, 4)
              Call REPLA(Y$, MKI$(FF%), 103, 2)
              Call REPLA(Y$, MKS$(NROO&), 105, 4)
              Call GRAREG("PEDDETA", Y$, REGPEDPE&)
              '
              For UK& = 1 To ULTREG&("PEDDETP")
                 Z$ = REGLEIDO$("PEDDETP", UK&)
                 If Left$(Z$, 30) = Left$(Y$, 30) Then
                   If Mid$(Z$, 33, 42) = Mid$(Y$, 33, 42) Then
                     Call REPLA(Z$, MKS$(CASIG), 75, 4)
                     Call REPLA(Z$, MKS$(TOENTRE), 95, 4)
                     Call REPLA(Z$, MKS$(SALPEN), 99, 4)
                     Call REPLA(Z$, MKI$(FF%), 103, 2)
                     Call REPLA(Z$, MKS$(NROO&), 105, 4)
                     Call GRAREG("PEDDETP", Z$, UK&)
                   End If
                 End If
              Next UK&
              '
            End If
          End If
        End If
    Next YX&
    '
5   Call CIERRARCH("MOVISTO")
    Call CIERRARCH("ENTRECOM")
    Call CIERRARCH("PEDDETP")
    Call CIERRARCH("PEDDETA")
    '
    GoSub PRIDOCREMI
    '
GoTo 999
    '
    '
    '
PRIDOCREMI:
    '
    PRINTERPORT$ = CONTROL$(IDENTER$, "PORTREM")
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIT" + CMXX$))
    If TRIM$(FORIPRE$) <> "" Then
       '
       PREFIREM$ = CONTROL$("", "PREFIRE" + CMXX$)
       If PREFIREM$ = "" Then PREFIREM$ = "0001"
       NROREMI$ = TRIM$(Mid$(REMI$, 4))
       While Len(NROREMI$) < 8: NROREMI$ = "0" + NROREMI$: Wend
       NROREMI$ = PREFIREM$ + "-" + NROREMI$
       '
       FEX = FF%
       FECHDOC$ = FECHATEX$(FEX)
       NUMEDOC$ = NROREMI$
       TIPODOC$ = TRIM$(Label6) + " ("
       If Option1.Value = True Then TIPODOC$ = TIPODOC$ + TRIM$(Option1.Caption) + ")"
       If Option2.Value = True Then TIPODOC$ = TIPODOC$ + TRIM$(Option2.Caption) + ")"
       '
       CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Text8.TEXT)
       CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = TRIM$(Text15.TEXT)
       CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Text12.TEXT)
       CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = TRIM$(Text17.TEXT)
       CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = TRIM$(Text14.TEXT)
       CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Text13.TEXT)
       CODPARAM$(7) = "DEPOSAL ": DOCPARAM$(7) = TRIM$(DEPOSAL$)
       CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = TRIM$(NROREMI$)
       CODPARAM$(9) = "NRO-OCOM": DOCPARAM$(9) = TRIM$(Text20.TEXT)
       CODPARAM$(10) = "PATE-TRA": DOCPARAM$(10) = TRIM$(Text18.TEXT)
       CODPARAM$(11) = "N-AUTENT": DOCPARAM$(11) = TRIM$(NORVENTA)
       CODPARAM$(12) = "COND-TRA": DOCPARAM$(12) = TRIM$(Text19)
       CODPARAM$(13) = "LOCA-TRA": DOCPARAM$(13) = TRIM$(Text21)
       CODPARAM$(14) = "RASO-CLI": DOCPARAM$(14) = TRIM$(RASOCLI$(NC%)) + " - Almacén " + Label8
       CAPARDOC% = 14
       '
       CODTABU1$(1) = "COD-MAT"
       CODTABU1$(2) = "DES-MAT"
       CODTABU1$(3) = "CANTIDAD"
       CODTABU1$(4) = "NRO-LOTE"
       CODTABU1$(5) = "UNIMED"
       CODTABU1$(6) = "N-OCOMPR"
       CODTABU1$(7) = "ORD-ITEM"
       CODTABU1$(8) = "REF-MAT1"
       CACOLTAB1% = 8
       '
       DESTIDOC% = NC%
       CAITAB1% = 0
       '
       For YX& = 1 To ULTREG&("!BOLREDET")
          YY$ = REGLEIDO$("!BOLREDET", YX&)
          CI% = CVI(Mid$(YY$, 83, 2))
          CAN# = CVD(Mid$(YY$, 89, 8))
          CAC# = 0
          If CI% > 0 Then
             If CI% <= NUMAS% Then
                If CAN# + CAC# > 0.005 Then
                   '
                   UNID$ = UNIMED$(CI%)
                   If UNID$ = "" Then UNID$ = UNIMED$(CI%)
                   CANS$ = CSTRING$(MKD$(CAN#), 3, 12, 1, 2)
                   If CAC# >= 0.05 Then
                      CANS$ = CSTRING$(MKD$(CAC#), 3, 12, 1, 2)
                      Call REPLA(CANS$, "(*)", 1, 3)
                   End If
                   NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
                   NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
                   '
                   CAITAB1% = CAITAB1% + 1
                   CODDX$ = CODEXT$(CI%)
                   TETABU1$(1, CAITAB1%) = CODDX$
                   TETABU1$(2, CAITAB1%) = TRIM$(DESCRIT0$(CI%))
                     If REPLAORI% = 1 Then TETABU1$(2, CAITAB1%) = TETABU1$(2, CAITAB1%) + " (" + TRIM$(CODEXT$(CI%)) + ")"
                   TETABU1$(3, CAITAB1%) = CANS$
                   TETABU1$(4, CAITAB1%) = NULOTE$
                   TETABU1$(5, CAITAB1%) = UNID$
                   TETABU1$(6, CAITAB1%) = NUOCO$
                   TETABU1$(7, CAITAB1%) = Str$(CAITAB1%)
                   CODDY$ = TRIM$(CODORIG$(CI%, NC%))
                   If CODDY$ <> TRIM$(CODDX$) Then
                     TETABU1$(8, CAITAB1%) = CODDY$
                   End If
                   '
                End If
             End If
          End If
       Next YX&
       '
       Call PRINTDOC("FOREMIT" + CMXX$) ' Remito de Facturacion
       '
    End If
    '
Return
    '
    '
999 End Sub
'
Sub NUREMITO(NROPRO&)
    '
    Screen.MousePointer = 11
    '
    FENUE$ = CONTROL$("", "FENUREFA")
    FLIM% = FECHANUM(FENUE$)
    If FLIM% > HOY(HO$) Then FLIM% = HOY(HO$)
    '
    CAMOLE% = 0
    NROMA& = 0
    RETOLE& = 0
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    
    For U& = ULTREG&("MOVISTO") To 1 Step -1
       XX$ = REGLEIDO$("MOVISTO", U&)
       RETOLE& = RETOLE& + 1
       If RETOLE& > 2048 Then
         If CAMOLE% >= 1 Then
           GoTo 85
         End If
       End If
       '
       If UCase$(Mid$(XX$, 23, 2)) = CDESPA$ Then
         CAMOLE% = CAMOLE% + 1
         If CAMOLE% > 256 Then GoTo 85
         '
         FEMO% = CVI(Left$(XX$, 2))
         If FEMO% >= FLIM% Then
           NROLE& = Val(Mid$(XX$, 26, 7))
           If NROLE& = NROPRO& Then
               NROPRO& = 0
           End If
           If NROLE& > NROMA& Then NROMA& = NROLE&
         End If
       End If
80  Next U&
    '
85  If NROPRO& < 1 Then NROPRO& = 1 + NROMA&
    Screen.MousePointer = 1
    '
End Sub
'
Sub GENARTRA()
    Screen.MousePointer = 11
    Call COPYSTRU("ENTRECOM", "TRAENTRE")
    Call BLOQARCH("ATRACEN")
    Call BLANARCH("ATRACEN")
    '
    JJ& = 0
    For U& = 1 To ULTREG("ENTRECOM")
      XX$ = REGLEIDO$("ENTRECOM", U&)
      STATU% = CVI(Mid$(X$, 191, 2))
      If STATU% <> 9 Then        ' NO ESTA ANULADO
        VENVI% = CVI(Mid$(XX$, 193, 2))     ' veces enviado
        If VENVI% < 5 Then
          Call REPLA(XX$, MKI$(VENVI% + 1), 193, 2)
          Call GRAREG("ENTRECOM", XX$, U&)
          JJ& = JJ& + 1
          Call GRAREG("TRAENTRE", XX$, JJ&)
        End If
      End If
    Next U&
    Call SETULTREG("TRAENTRE", JJ&)
    Call CIERRARCH("ENTRECOM")
    Call CIERRARCH("TRAENTRE")
    '
    Dim BUFLE As String * 128
    LU$ = LUDAT$
    N1% = FILEOPEN(LU$ + "TRAENTRE.DAT", 0, 128)
    For KU& = 1 To 1 + LOF(N1%) / 128
      Get #N1%, KU&, BUFLE$
      TTXX$ = Space$(8) + String$(128, 0)
      Call REPLA(TTXX$, "TRAENTRE", 1, 8)
      Call REPLA(TTXX$, BUFLE$, 9, 128)
      Call REGAPP("ATRACEN", TTXX$)
    Next KU&
    Close #N1%
    '
    Call CIERRARCH("ATRACEN")
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 1
    '
    FILEORIG.Caption = "Destino de Archivo"
    FILEORIG.Show 1
    If FILEORIG.Label1.Caption = "" Then GoTo 99
    '
40  LU$ = LUDAT$
    ORIGARCH$ = Left$(FILEORIG.Label1.Caption, 2)
    If Mid$(ORIGARCH$, 2, 1) <> ":" Then
       Call MENSERR(24, "Unidad '" + ORIGARCH$ + "' No Disponible.")
       GoTo 99
    End If
    On Error GoTo 50
    FileCopy LU$ + "ATRACEN.DAT", ORIGARCH$ + "\ATRACEN.DAT"
    On Error GoTo 0
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call COMUNI("Archivo de Transferencia Electrónica\para Central Avellaneda\Generado Exitosamente EN '" + ORIGARCH$ + "\.")
    GoTo 99
    '
50  Resume 51
51  On Error GoTo 0
    Call MENSERR(24, "Error de Copiado a Unidad.\   \ Intente de Nuevo.")
    GoTo 99
    '
99  Screen.MousePointer = 1
End Sub

