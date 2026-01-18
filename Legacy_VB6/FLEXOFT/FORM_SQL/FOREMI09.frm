VERSION 5.00
Begin VB.Form FOREMI09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Despacho de Materiales "
   ClientHeight    =   8355
   ClientLeft      =   1995
   ClientTop       =   1485
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
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
      Begin VB.PictureBox Picture14 
         BackColor       =   &H00400000&
         Height          =   8535
         Left            =   10905
         ScaleHeight     =   8475
         ScaleWidth      =   1155
         TabIndex        =   56
         Top             =   -45
         Width           =   1215
         Begin VB.CommandButton Command5 
            Caption         =   "Imprimir"
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
            Left            =   90
            Picture         =   "FOREMI09.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Aprueba - Graba e Imprime Remito"
            Top             =   6270
            Width           =   750
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Factura"
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
            Left            =   90
            Picture         =   "FOREMI09.frx":066A
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Acceso Directo a Facturación"
            Top             =   4965
            Width           =   750
         End
         Begin VB.CommandButton Command6 
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
            Height          =   660
            Left            =   90
            Picture         =   "FOREMI09.frx":0AAC
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Opciones Complementarias - Listados y Reportes"
            Top             =   5625
            Width           =   750
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
            Height          =   660
            Left            =   105
            Picture         =   "FOREMI09.frx":0BF6
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Ababdona la Aplicación"
            Top             =   165
            Width           =   750
         End
         Begin VB.CommandButton Command11 
            Caption         =   "New"
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
            Picture         =   "FOREMI09.frx":0D40
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Apertura de Nueva Cuenta de Cliente"
            Top             =   1230
            Width           =   750
         End
         Begin VB.CommandButton Command4 
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
            Height          =   660
            Left            =   105
            Picture         =   "FOREMI09.frx":1182
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   1890
            Width           =   750
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
            Height          =   660
            Left            =   105
            Picture         =   "FOREMI09.frx":148C
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Notas de Clientes"
            Top             =   2550
            Width           =   750
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
            Height          =   660
            Left            =   105
            Picture         =   "FOREMI09.frx":1796
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Ficha de Movimientos de Material"
            Top             =   3210
            Width           =   750
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Abrir"
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
            Height          =   660
            Left            =   105
            Picture         =   "FOREMI09.frx":1AA0
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Descompone Conjuntos en Partes Componentes"
            Top             =   3870
            Width           =   750
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   90
            ScaleHeight     =   75
            ScaleWidth      =   690
            TabIndex        =   57
            Top             =   7530
            Width           =   750
            Begin VB.Frame Frame13 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   58
               Top             =   105
               Width           =   12000
            End
         End
         Begin VB.Image Image3 
            Height          =   390
            Left            =   -270
            Picture         =   "FOREMI09.frx":1DAA
            Top             =   7800
            Width           =   1515
         End
      End
      Begin VB.TextBox Text22 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2970
         Left            =   210
         MultiLine       =   -1  'True
         TabIndex        =   54
         Text            =   "FOREMI09.frx":3CCC
         Top             =   2170
         Width           =   6330
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   210
         ScaleHeight     =   75
         ScaleWidth      =   960
         TabIndex        =   35
         Top             =   1300
         Width           =   1020
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFC0&
         Height          =   700
         Left            =   6930
         TabIndex        =   33
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
            TabIndex        =   9
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
            TabIndex        =   34
            Top             =   285
            Width           =   2115
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   32
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
            TabIndex        =   53
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
            TabIndex        =   13
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
            TabIndex        =   12
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
            TabIndex        =   11
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
            TabIndex        =   10
            Top             =   525
            Width           =   3375
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Height          =   3020
         Left            =   210
         TabIndex        =   30
         Top             =   5250
         Width           =   6525
         Begin VB.Frame Frame14 
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   45
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
               TabIndex        =   50
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
               TabIndex        =   46
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
               TabIndex        =   52
               Top             =   800
               Width           =   435
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   42
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
               TabIndex        =   43
               Top             =   315
               Width           =   1695
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   41
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
               TabIndex        =   44
               Top             =   315
               Width           =   1275
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   31
            Top             =   210
            Width           =   3690
            Begin VB.Frame Frame5 
               BackColor       =   &H00FFFFC0&
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
               TabIndex        =   47
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
                  TabIndex        =   51
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
                  TabIndex        =   48
                  Top             =   315
                  Width           =   3270
               End
            End
            Begin VB.Frame Frame10 
               BackColor       =   &H00FFFFC0&
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
               TabIndex        =   39
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
                  TabIndex        =   40
                  Top             =   210
                  Width           =   1380
               End
            End
            Begin VB.Frame Frame9 
               BackColor       =   &H00FFFFC0&
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
               TabIndex        =   37
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
                  TabIndex        =   38
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
               TabIndex        =   8
               Top             =   315
               Width           =   3270
            End
         End
      End
      Begin VB.TextBox Text7 
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
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   1335
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
         Left            =   6195
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "X"
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   315
         Width           =   435
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
         Left            =   9030
         TabIndex        =   4
         Text            =   " "
         Top             =   735
         Width           =   1650
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
         Left            =   8715
         TabIndex        =   6
         Text            =   " "
         Top             =   1155
         Width           =   330
      End
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
         Left            =   9030
         TabIndex        =   5
         Top             =   1155
         Width           =   180
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
         Left            =   240
         TabIndex        =   0
         Text            =   " "
         Top             =   720
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   3480
         Left            =   210
         ScaleHeight     =   3420
         ScaleWidth      =   10455
         TabIndex        =   24
         Top             =   1680
         Width           =   10515
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C000&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   480
            Left            =   0
            TabIndex        =   26
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
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   $"FOREMI09.frx":3CD3
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
               TabIndex        =   27
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
            TabIndex        =   25
            Top             =   -1050
            Width           =   10515
            Begin VB.Line Line10 
               X1              =   0
               X2              =   10500
               Y1              =   1530
               Y2              =   1530
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
      Begin VB.Label REPRESUP 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   225
         Left            =   2415
         TabIndex        =   55
         Top             =   0
         Visible         =   0   'False
         Width           =   1065
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
         Height          =   225
         Left            =   0
         TabIndex        =   49
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
         TabIndex        =   29
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   1050
         Width           =   2115
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro:"
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
         Left            =   7560
         TabIndex        =   23
         Top             =   735
         Width           =   1380
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito:"
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
         Left            =   7350
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
Attribute VB_Name = "FOREMI09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub PRIDEVOL_MSC()
   '
    NUOREP& = Me.REPRESUP
    TALONX& = XVALO(VALOBADA("REPARA", "NumRepa", "NuOrdRep = " & NUOREP&))
    Call MENURE09.CONSUMO_MAT_SUM_CLIENTE(TALONX&, XVALO(Mid$(Text11, 4)))

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
    Call SELECHO("ECODEP")
    DPX1% = Val(VALACT1$("ECODEP"))
    If DPX1% > 0 Then
        DP1% = DPX1%
        Text10.TEXT = TRIM$(Str$(DP1%))
        Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    End If
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    If EXISTE%("ARCHIG07.EXE") > 0 Then
         Call CONECRUN("ARCHIG07/AMACLI", "Alta de Nueva Cuenta de Cliente")
       Else
         Call CONECRUN("AMACLI/NEW", "Alta de Nueva Cuenta de Cliente")
    End If
    Call FRESHECHO("DEUDOR1")
    Command11.Enabled = True
End Sub

Private Sub Command2_Click()
    '
    If ULTREG&("!BOLREDET") > 0 Then
        Call CIERRARCH("!BOLREDET")
        Call BLANARCH("!BOLREDET")
        'List1.Clear
        'List1.Visible = False
        Command5.Enabled = False
        'BOTNEXT.Enabled = True
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


Private Sub Command5_Click()
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
    Command5.Enabled = False
    For U& = 1 To 4
       TTXXX$ = TRIM$(Text16(U& - 1).TEXT)
       Call GRAREG("!OBSERVOF", TTXXX$, U&)
    Next U&
    Call SETULTREG("!OBSERVOF", 4)
    Call CIERRARCH("!OBSERVOF")
    '
    Call PRIDEVOL_MSC
    Call PRIREMIFAC
    'DEDOC$ = TIPODOC$ + " " + NUMEDOC$ + " " + TRIM$(FOREMIFAC.Text1.TEXT) + "-" + TRIM$(FOREMIFAC.Text2.TEXT)
    'Call SAVESPOOL(DEDOC$, OKX%)      ' guarda el spool de impresion
    '
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    DP1% = Val(CONTROL$("", "DEPOEXPE"))
    Text10.TEXT = TRIM$(Str$(DP1%))
    Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If CMO$ = "" Then CMO$ = "RT"
    NROPRO& = 0
    PREFIREM$ = CONTROL$("", "PREFIREM")
    If PREFIREM$ = "" Then PREFIREM$ = "0001"
    Call NUREMIFA(NROPRO&, PREFIREM$)
    Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
    '
    Call CIERRARCH("!BOLREDET")
    Call BLANARCH("!BOLREDET")
    Text22 = ""
    Command5.Enabled = True
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Unload Me
    '
End Sub
Private Sub Command6_Click()
    '
    Command6.Enabled = False
    'OPREMSTAN.Show 1
    Command6.Enabled = True
    '
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    'Call CONECRUN("*KARDEX", "Ficha de Movimientos por Pieza")
    'Call GENKARDEX("")
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
    TEXPE$ = "Remito de Ventas"
    If CONTROL$("", "MODOEXPE") = "ODESPACH" Then
        TEXPE$ = "Orden de Despacho"
    End If
    Caption = Caption + " - " + TEXPE$
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
10  'If ULTREG("ECODEP") < 1 Then
    '    Call MENSERR(24, "Falta Definir Tabla de Depósitos")
    '    Call CARTADEPOSI
    '    GoTo 10
    'End If
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
    If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Remito.")
        OPREMIFAC.Show 1
        GoTo 10
    End If
    '
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito de Facturación:\  \'" + RMCC$ + "'.")
            OPREMIFAC.Show 1
            GoTo 10
        End If
    End If
    '
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    DP1% = Val(CONTROL$("", "DEPOEXPE"))
    If DP1% < 1 Or ECOARCH$("ECODEP", Chr$(DP1%)) = "" Then
       Call COMUNI("Error de Configuración:\  \Falta Definir Depósito de Expedición.")
       OPREMIFAC.Show 1
       GoTo 10
    End If
    '
    Text10.TEXT = TRIM$(Str$(DP1%))
    Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If CMO$ = "" Then CMO$ = "RT"
    NROPRO& = 0
    PREFIREM$ = CONTROL$("", "PREFIREM")
    If PREFIREM$ = "" Then PREFIREM$ = "0001"
    Call NUREMIFA(NROPRO&, PREFIREM$)
    Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
    '
    Refresh
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Cancel = 0
   Screen.MousePointer = 1
   Hide
End Sub

Private Sub Label2_DblClick()
    OPREMIFAC.Show 1
End Sub

Private Sub Label9_DblClick()
   Frame15.Visible = True
   List1.Visible = True
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
                List1.Clear
                Call CARTRANSPOR
                BOTNEXT.Enabled = True
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
    Static CMO$
    If CMO$ = "" Then
       CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
       If CMO$ = "" Then CMO$ = "RT"
    End If
    '
    VDEFF$ = TRIM$(Text11.TEXT)
    Text11.TEXT = VDEFF$
    If Left$(VDEFF$, 3) <> CMO$ + "." Then GoTo 90
    NRONUE& = Val(Mid$(VDEFF$, 4))
    NROPRO& = NRONUE&
    PREFIREM$ = CONTROL$("", "PREFIREM")
    If PREFIREM$ = "" Then PREFIREM$ = "0001"
    Call NUREMIFA(NROPRO&, PREFIREM$)
    If NROPRO& <> NRONUE& Then
        Beep
        GoTo 95
    End If
    Exit Sub
    '
90  Beep
    NROPRO& = 0
    PREFIREM$ = CONTROL$("", "PREFIREM")
    If PREFIREM$ = "" Then PREFIREM$ = "0001"
    Call NUREMIFA(NROPRO&, PREFIREM$)
95  Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
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
    OPREMIFAC.Show 1
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

Sub PRIREMIFAC()
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
    If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Remito.")
        Exit Sub
    End If
    '
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito de Facturación:\  \'" + RMCC$ + "'.")
            Exit Sub
        End If
    End If
    '
    VERIST% = (-1)
    REPLAORI% = 0
    '
20  Screen.MousePointer = 11      ' Registrando Remito
    '
    AINDI% = 0
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    NC = Val(Text1.TEXT)
    REMI$ = TRIM$(Text11.TEXT)
    NROO& = Val(Mid$(REMI$, 4))
    FF% = FECHANUM(Text9.TEXT)
    DEPO% = Val(Text10.TEXT)
    '
    PREFIREM$ = CONTROL$("", "PREFIREM")
    If PREFIREM$ = "" Then PREFIREM$ = "0001"
    NROREMI$ = TRIM$(Mid$(REMI$, 4))
    While Len(NROREMI$) < 8: NROREMI$ = "0" + NROREMI$: Wend
    NROREMI$ = CMO$ + "-X-" + PREFIREM$ + "-" + NROREMI$
    '
    HO% = HOY(HOS$)
    HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
    User% = USNBR% Mod 100
    CTRXX$ = MKI$(HO%) + MKI$(Int(HP)) + Chr$(User%) + String$(2, 0)
    '
    YXI% = 0
    Ci% = CODINT%(CONTROL$("REPARA", "CODGENER"))
    CAN# = 1
'   XX$ = REGLEIDO$("PRESREPA", Val(REPRESUP))
'   LOTE$ = TRIM$(Mid$(XX$, 71, 15))
    CONDI$ = "NuOrdRep = " & REPRESUP
    LOTE$ = SAFETEXT$(VALOBADA("REPARA", "NumeroSerie", CONDI$))
    MEDIDAS$ = ""
    NUOC& = Val(Text20)
    DOSEC$ = NROREMI$
    CANTI = (-1) * CAN#
    '
    DEPX% = DEPO%: If DEPO% = 99 Then DEPX% = DEPOPRE%(Ci%)
'   Call Movisto(FF%, CI%, LOTE$, CMO$, REMI$, DOSEC$, RASOCLI$(NC), 0, "$", NC, DEPX%, CANTI)
    Call MOVISTOK(FF%, Ci%, LOTE$, CMO$, DOSEC$, REMI$, DOSEC$, REMI$, REFE$, NUORIT%, 0, 0, NC, DepOrig%, CANTI)
    '
    ' blanquear pedidos pendientes
    'XX$ = REGLEIDO$("PRESREPA", Val(REPRESUP))
    'Call REPLA(XX$, MKI$(FF%), 123, 2)
    'Call GRAREG("PRESREPA", XX$, Val(REPRESUP))
    Call ACTUREGISTRO("REPARA", "FechApro", CONDI$, FETEXCOR1$(FF%), REG&)
    '
5   Call CIERRARCH("MOVISTO")
    Call CIERRARCH("PRESREPA")
    '
    GoSub PRIDOCREMI
    '
    XX$ = REGLEIDO("PRESREPA", Val(REPRESUP))
    TTXYZ$ = Space$(92)
    NROREP& = CVS(Left$(XX$, 4))
    HOII% = HOY(HOS$)
    Call REPLA(TTXYZ$, HOS$, 1, 8)
    Call REPLA(TTXYZ$, USERTER$, 69, 24)
    Call REPLA(TTXYZ$, "Despacho de Equipo Reparado", 11, 56)
    ' Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
    Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
    '
    TTXYZ$ = Space$(92)
    Call REPLA(TTXYZ$, Space$(56), 11, 56)
    ' Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
    Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
    Call REPLA(TTXYZ$, "Cliente: " + Text2, 11, 56)
    ' Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
    Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
    Call REPLA(TTXYZ$, " Remito: " + Text11, 11, 56)
    ' Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
    Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
    '
    Call REPLA(TTXYZ$, Space$(56), 11, 56)
    ' Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
    Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
    '
    Call FRATEXTO(Text22, 46)
    For KKI% = 1 To CARETEX%
      TTXX$ = Space$(92)
      Call REPLA(TTXX$, RETEX$(KKI%), 10, 46)
      If KKI% = 1 Then Call REPLA(TTXX$, "Descrip: ", 1, 9)
      ' Call REGAPP("ANOTAREP", Left$(MKS$(NROREP&) + Space$(10) + TTXX$, 96))
      Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, Space$(10) + TTXX$)
    Next KKI%
    '
    ' Call REGAPP("ANOTAREP", MKS$(NROREP&) + String$(88, "="))
    Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, String$(88, "="))
    Call CIERRARCH("ANOTAREP")
    '
GoTo 999
    '
    '
PRIDOCREMI:
    '
    PRINTERPORT$ = CONTROL$(IDENTER$, "PORTREM")
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
    If TRIM$(FORIPRE$) <> "" Then
       '
       PREFIREM$ = CONTROL$("", "PREFIREM")
       If PREFIREM$ = "" Then PREFIREM$ = "0001"
       NROREMI$ = TRIM$(Mid$(REMI$, 4))
       While Len(NROREMI$) < 8: NROREMI$ = "0" + NROREMI$: Wend
       NROREMI$ = PREFIREM$ + "-" + NROREMI$
       '
       FEX = FF%
       FECHDOC$ = FECHATEX$(FEX)
       NUMEDOC$ = NROREMI$
       TIPODOC$ = "Remito"
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
       CAPARDOC% = 13
       '
       CODTABU1$(1) = "COD-MAT"
       CODTABU1$(2) = "DES-MAT"
       CODTABU1$(3) = "CANTIDAD"
       CODTABU1$(4) = "NRO-LOTE"
       CODTABU1$(5) = "UNIMED"
       CODTABU1$(6) = "N-OCOMPR"
       CODTABU1$(7) = "ORD-ITEM"
       CACOLTAB1% = 7
       '
       DESTIDOC% = NC
       CAITAB1% = 0
       '
       Call FRATEXTO(Text22, 30)
       
       For YX& = 1 To CARETEX%
          CANS$ = "": If YX& = 1 Then CANS$ = "1.00"
          CAITAB1% = CAITAB1% + 1
          '
          TETABU1$(1, CAITAB1%) = ""
          TETABU1$(2, CAITAB1%) = RETEX$(YX&)
          TETABU1$(3, CAITAB1%) = CANS$
          TETABU1$(4, CAITAB1%) = ""
          TETABU1$(5, CAITAB1%) = ""
          TETABU1$(6, CAITAB1%) = ""
          TETABU1$(7, CAITAB1%) = ""
          '
       Next YX&
       '
       Call PRINTDOC("FOREMIFA")   ' Remito de Facturacion
       '
    End If
    '
Return
    '
999 End Sub
'
