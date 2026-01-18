VERSION 5.00
Begin VB.Form LIBREMIT04 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remito de Texto Libre"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11940
   Begin VB.CommandButton Command12 
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
      Height          =   630
      Left            =   12915
      Picture         =   "LIBREMIT04.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Descompone Conjuntos en Partes Componentes"
      Top             =   420
      Width           =   650
   End
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
      TabIndex        =   11
      Top             =   0
      Width           =   12000
      Begin VB.CommandButton Command7 
         Height          =   260
         Left            =   210
         TabIndex        =   59
         Top             =   4850
         Width           =   5565
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Selecc Todo"
         Height          =   260
         Left            =   9345
         TabIndex        =   58
         Top             =   4850
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Pegar"
         Height          =   260
         Left            =   8085
         TabIndex        =   57
         Top             =   4850
         Width           =   1275
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Copiar"
         Height          =   260
         Left            =   6930
         TabIndex        =   56
         Top             =   4850
         Width           =   1170
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Cortar"
         Height          =   260
         Left            =   5775
         TabIndex        =   55
         Top             =   4850
         Width           =   1170
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame16"
         Height          =   3300
         Left            =   105
         TabIndex        =   60
         Top             =   1600
         Visible         =   0   'False
         Width           =   10830
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
            Height          =   2850
            Left            =   1470
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   63
            Top             =   430
            Width           =   7890
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00E0E0E0&
            Height          =   435
            Left            =   105
            ScaleHeight     =   375
            ScaleWidth      =   10515
            TabIndex        =   61
            Top             =   0
            Width           =   10575
            Begin VB.Label Label13 
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
               ForeColor       =   &H00404000&
               Height          =   225
               Left            =   0
               TabIndex        =   62
               Top             =   105
               Width           =   10515
            End
         End
         Begin VB.Line Line6 
            X1              =   10680
            X2              =   10680
            Y1              =   0
            Y2              =   3465
         End
         Begin VB.Label Label14 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            Height          =   2850
            Left            =   9360
            TabIndex        =   65
            Top             =   435
            Width           =   1320
         End
         Begin VB.Label Label15 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            Height          =   2850
            Left            =   105
            TabIndex        =   64
            Top             =   435
            Width           =   1380
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9360
         TabIndex        =   52
         Top             =   6000
         Width           =   1335
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
            Left            =   450
            TabIndex        =   53
            Top             =   270
            Width           =   540
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Orden  de Suministro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   6930
         TabIndex        =   49
         Top             =   5250
         Width           =   3795
         Begin VB.TextBox Text22 
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
            MaxLength       =   24
            TabIndex        =   50
            Top             =   300
            Width           =   3375
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00800000&
         Height          =   9045
         Left            =   11025
         ScaleHeight     =   8985
         ScaleWidth      =   3210
         TabIndex        =   43
         Top             =   -105
         Width           =   3270
         Begin VB.PictureBox Picture16 
            Height          =   8450
            Left            =   0
            ScaleHeight     =   8385
            ScaleWidth      =   1020
            TabIndex        =   67
            Top             =   0
            Visible         =   0   'False
            Width           =   1080
            Begin VB.PictureBox MARCOBARRA 
               BackColor       =   &H000000FF&
               Height          =   135
               Left            =   120
               ScaleHeight     =   75
               ScaleWidth      =   540
               TabIndex        =   68
               Top             =   7680
               Width           =   600
               Begin VB.Frame BARRATRAZA 
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   0  'None
                  Height          =   500
                  Left            =   0
                  TabIndex        =   69
                  Top             =   105
                  Width           =   12000
               End
            End
            Begin VB.Image Image3 
               Height          =   8450
               Left            =   0
               Picture         =   "LIBREMIT04.frx":030A
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1095
            End
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
            Height          =   690
            Left            =   105
            Picture         =   "LIBREMIT04.frx":26A2
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Configuracion de Funcionamiento"
            Top             =   5085
            Width           =   650
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
            Height          =   690
            Left            =   105
            Picture         =   "LIBREMIT04.frx":2AE4
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Terminar - Salir de la Aplicación"
            Top             =   200
            Width           =   650
         End
         Begin VB.CommandButton Command5 
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
            Left            =   105
            Picture         =   "LIBREMIT04.frx":2C2E
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Aprueba - Graba e Imprime Remito"
            Top             =   7080
            Width           =   650
         End
         Begin VB.CommandButton Command6 
            Caption         =   "More"
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
            Left            =   105
            Picture         =   "LIBREMIT04.frx":3298
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Opciones Complementarias - Listados y Reportes"
            Top             =   4395
            Width           =   650
         End
         Begin VB.Image Image2 
            Height          =   390
            Left            =   -300
            Picture         =   "LIBREMIT04.frx":33E2
            Top             =   7980
            Width           =   1515
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Valor Declarado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   6930
         TabIndex        =   25
         Top             =   5985
         Width           =   2235
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
            Left            =   480
            TabIndex        =   6
            Top             =   250
            Width           =   1170
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
         Height          =   1635
         Left            =   6930
         TabIndex        =   24
         Top             =   6615
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
            TabIndex        =   42
            Top             =   2500
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
            TabIndex        =   10
            Top             =   1260
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
            TabIndex        =   9
            Top             =   945
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
            TabIndex        =   8
            Top             =   630
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
            TabIndex        =   7
            Top             =   315
            Width           =   3375
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Height          =   3020
         Left            =   210
         TabIndex        =   22
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
            TabIndex        =   34
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
               TabIndex        =   39
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
               TabIndex        =   35
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
               TabIndex        =   41
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
            TabIndex        =   31
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
               MaxLength       =   8
               TabIndex        =   32
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
            TabIndex        =   30
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
               TabIndex        =   33
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
            TabIndex        =   23
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
               TabIndex        =   36
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
                  TabIndex        =   40
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
                  TabIndex        =   37
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
               TabIndex        =   28
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
                  TabIndex        =   29
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
               TabIndex        =   26
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
                  TabIndex        =   27
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
               TabIndex        =   5
               Top             =   315
               Width           =   3270
            End
         End
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
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
         Left            =   6090
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "X"
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   315
         Width           =   540
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
         Top             =   1155
         Width           =   1650
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
            Size            =   12
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
      Begin VB.Label TEPRUE 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   66
         Top             =   1470
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000040&
         Height          =   300
         Left            =   8340
         TabIndex        =   54
         ToolTipText     =   "Doble Click  para Cambiar el Punto de Venta"
         Top             =   1155
         Width           =   660
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
         TabIndex        =   51
         Top             =   1155
         Width           =   1035
      End
      Begin VB.Label NORVENTA 
         Height          =   225
         Left            =   0
         TabIndex        =   38
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
         TabIndex        =   21
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
         Left            =   6825
         TabIndex        =   19
         Top             =   1190
         Width           =   1380
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
         TabIndex        =   18
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
         TabIndex        =   12
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
      TabIndex        =   17
      Top             =   0
      Width           =   960
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "LIBREMIT04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
        Label26 = CODICLI$(NC%)
        Call CARTRANSPOR
        Text28.SetFocus
    End If
    '
End Sub


Private Sub Command13_Click()
    Command13.Enabled = False
    '\\\OT-05-0299-WAR-18/05/05///
    If DERACCE%("SETREMIT", "Configuración - Setup de Remitos.") >= 2 Then
      OPREMIT04.Show 1
    End If
    '\\\OT-05-0299-WAR-FIN///
    Command13.Enabled = True
End Sub


Private Sub Command17_Click()
    '
    COMIE = Text28.SelStart
    LOSEL = Text28.SelLength
    Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
    Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
    '
End Sub

Private Sub Command2_Click()
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command3_Click()
   '
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Clipboard.GetText + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
   '
End Sub

Private Sub Command4_Click()
   '
   Text28.SelStart = 0
   Text28.SelLength = Len(Text28.TEXT)
   Text28.SetFocus
   '
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    '
    ' TRANSFIERE LAS OBSERVACIONES
    For U& = 1 To 4
       TTXXX$ = TRIM$(Text16(U& - 1).TEXT)
       Call GRAREG("!OBSERVOF", TTXXX$, U&)
    Next U&
    Call SETULTREG("!OBSERVOF", 4)
    Call CIERRARCH("!OBSERVOF")
    '
    '\\\OT-05-0323-EPB-24/05/05///
    Call BLOQARCH("MOVISTO")
    NPRESENTADO& = Val(Mid(Text11, 4))
    'OT-05-0419-WAR-26/06/05
    NUPUVE$ = TRIM(Label10)
    NROPRO& = NPRESENTADO&
    '
    Call NUREMIFA(NROPRO&, NUPUVE$)
    If NPRESENTADO& <> NROPRO& Then
        Call MENSERR(24, "Número de Remito ya Utilizado.\FLEXOFT Asignará el Próximo Número Libre.\Vuelva a Pedir la Impresión.")
        CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
        If CMO$ = "" Then CMO$ = "RT"
        Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
        Command5.Enabled = True
        Exit Sub
    End If
    '
10  Call PRIREMIFAC
    Call LIBARCH("MOVISTO")
    '\\\OT-05-0323-EPB-FIN///
    '
    If COMALTER%("Desea Guardar el Texto\Para el Próximo Remito Libre ?\\Si, Guardarlo\No, Descartarlo") = 1 Then
        Call SAVEFILE(LUDAT$ + "TEREMILI.TXT", Text28)
      Else
        Call SAVEFILE(LUDAT$ + "TEREMILI.TXT", "")
    End If
    '
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If CMO$ = "" Then CMO$ = "RT"
    NROPRO& = 0
    'OT-05-0419-WAR-26/06/05///
    NUPUVE$ = TRIM(Label10)
    Call NUREMIFA(NROPRO&, NUPUVE$)
    Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
    '
    Call CIERRARCH("!BOLREDET")
    Call BLANARCH("!BOLREDET")
    
    Command5.Enabled = False
    Command12.Enabled = False
    
    Text15 = ""
    Text17 = ""
    Unload Me
    '
End Sub

Private Sub Command6_Click()
    '
    Command6.Enabled = False
    OPREMSTA04.Show 1
    Command6.Enabled = True
    '
End Sub

Private Sub Command8_Click()
    '
    Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
    '
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
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame10.BackColor = &HFCD7B6
      Me.Frame11.BackColor = &HFCD7B6
      Me.Frame12.BackColor = &HFCD7B6
      Me.Frame14.BackColor = &HFCD7B6
      Me.Frame16.BackColor = &HFCD7B6
      Me.Frame17.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      
      Me.Frame6.BackColor = &HFCD7B6
      Me.Frame7.BackColor = &HFCD7B6
      Me.Frame8.BackColor = &HFCD7B6
      Me.Frame12.BackColor = &HFCD7B6
      Me.Frame9.BackColor = &HFCD7B6
'      MARCOBARRA.Top = 3600
'      Set MARCOBARRA.Container = Me.Picture16
'      MARCOBARRA.ZOrder 0
      ' &H8000000F
    End If

    FORIPRE$ = CONTROL$("", "FOREMIFA")
        If FORIPRE$ <> "" Then
          If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") > 0 Then
            AA1$ = ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")
            ACOLDE = Val(Mid$(AA1$, 2))
            If ACOLDE > 0 Then
              If ACOLDE <= 128 Then
                TEPRUE.Font = Text28.Font
                TEPRUE.FontSize = Text28.FontSize
                TEPRUE.FontBold = Text28.FontBold
                TEPRUE = String$(ACOLDE + 3, "*")
                ATEX = TEPRUE.Width + 50
                '
                ANCHOTO = Label14.Width + Text28.Width + Label15.Width
                Label14.Width = (ANCHOTO - ATEX) / 2
                Label15.Width = (ANCHOTO - ATEX) / 2
                Text28.Left = Label15.Left + Label15.Width
                Text28.Width = ATEX
                Label14.Left = Text28.Left + Text28.Width
                '
              End If
            End If
          End If
        End If
        '
        Frame3.Top = 1570
        Frame3.Visible = True
        'Frame2.Top = 5985
        DoEvents
      'Else
        'Frame16.Visible = False
        'Frame2.Top = 3960
    '
    Call CENTRAFORM(Me)
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
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    DP1% = Val(CONTROL$("", "DEPOEXPE"))
    If DP1% < 1 Or ECOARCH$("ECODEP", Chr$(DP1%)) = "" Then
       Call COMUNI("Error de Configuración:\  \Falta Definir Depósito de Expedición.")
       OPREMIT04.Show 1
       GoTo 10
    End If
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If CMO$ = "" Then CMO$ = "RT"
    '
    Label10 = ""
    PREFIREM$ = TRIM$(CONTROL$("", "PREFIREM"))
    If PREFIREM$ <> "*" Then
      Label10 = PREFIREM$
    End If
    '
    NROPRO& = 0
    Call NUREMIFA(NROPRO&)
    Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
    '
    Text28 = LOADFILE$(LUDAT$ + "TEREMILI.TXT")
    '\\\OT-05-420-WAR-26/06/05///
    '\\\OT-05-420-WAR-FIN///
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Hide
   Cancel = 0
End Sub
'\\\OT-05-0419-WAR-26/06/05///
Private Sub Label10_dblClick()
  TALO$ = TRIM$(Label10)
'  Select Case TALO$
'    Case ""
'        Label10 = "0001"
'    Case "0001"
'        Label10 = "0002"
'    Case "0002"
'        Label10 = "0001"
'  End Select
  Select Case TALO$
    Case ""
        Label10 = "0001"
    Case "0001"
        Label10 = "0002"
    Case "0002"
        Label10 = "0003"
    Case "0003"
        Label10 = "0004"
    Case "0004"
        Label10 = "0001"
  End Select
  NUPUVE$ = TRIM(Label10)
  CMO$ = "RT"
  NROPRO& = 0
  Call NUREMIFA(NROPRO&, NUPUVE$)
  Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
End Sub

Private Sub Label18_Change()
   '\\\OT-06-0001-WAR-02/01/06///
   TOTITEMS% = Val(Label18)
   If TOTITEMS% > TOTITFO& Then
     Call MENSERR(24, "Renglones de Este Remito (" & TRIM$(Str$(TOTITEMS%)) & ") Superan\Máximo Permitido (" & TRIM$(Str$(TOTITFO&)) & ") del Formulario\de Impresión de Remito")
   End If
   '
End Sub

'\\\OT-05-0419-WAR-FIN///
Private Sub Label2_DblClick()
'    OPREMIFAC.Show 1
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
         Text2.TEXT = RASOCLI$(NC%)
         Text3.TEXT = DOMICLI$(NC%)
         Text4.TEXT = CPOSCLI$(NC%)
         Text5.TEXT = LOCACLI$(NC%)
         Call CARTRANSPOR
       End If
     End If
   End If
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
        Label26 = CODICLI$(NC%)
        Call CARTRANSPOR
    End If
    '
End Sub

Private Sub Text1_GotFocus()
    Static CTXX%
    If CTXX% = 0 Then
      Refresh
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
                Label26 = CODICLI$(NC%)
                Unload BCODE04 '\\\OT-06-0182-WAR-29/05/06///
                Call CARTRANSPOR
                Text28.SetFocus
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
    '\\\OT-05-0419-WAR-26/06/05///
    'Call NUREMIFA(NROPRO&)
    NUPUVE$ = TRIM(Label10)
    Call NUREMIFA(NROPRO&, NUPUVE$)
    'FIN
    If NROPRO& <> NRONUE& Then
        Call MENSERR(24, "Número de Remito ya Utilizado.\FLEXOFT Asignara el Próximo Número libre.")
        Beep
        GoTo 95
    End If
    Exit Sub
    '
90  Beep
    NROPRO& = 0
    'OT-05-0419-WAR-26/06/05///
    NUPUVE$ = TRIM(Label10)
    Call NUREMIFA(NROPRO&, NUPUVE$)
    'Call NUREMIFA(NROPRO&)
    'FIN OT
95  Text11.TEXT = CMO$ + "." + TRIM$(Str$(NROPRO&))
    Text11.SetFocus
    '
End Sub

Private Sub Text2_gotfocus()
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
    Call Command13_Click
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
    If TRIM$(Text28) = "" Then
      Call COMUNI("No se ha Ingresado Detalle")
      Exit Sub
    End If
    '
    CAITMO% = 0
    '
20  HOJASUEL% = 0
    If CONTROL$("", "TIPAREMI") = "HOJASUEL" Then HOJASUEL% = 1
    '
    Screen.MousePointer = 11      ' Registrando Remito
    '
    AINDI% = 0
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
      CMO$ = "RT"
    End If
    '
    NC% = Val(Text1.TEXT)
    REMI$ = TRIM$(Text11.TEXT)
    NROO& = Val(Mid$(REMI$, 4))
    FF% = FECHANUM(Text9.TEXT)
    DEPO% = 0
    PREFIREM$ = TRIM(Label10)
    NROREMIT$ = TRIM$(Mid$(REMI$, 4))
    While Len(NROREMIT$) < 8: NROREMIT$ = "0" + NROREMIT$: Wend
    NROREMIT$ = CMO$ + "-X-" + PREFIREM$ + "-" + NROREMIT$
    '
    HO% = HOY(HOS$)
    HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
    User% = USNBR% Mod 100
    CTRXX$ = MKI$(HO%) + MKI$(Int(HP)) + Chr$(User%) + String$(2, 0)
    '
    YXI% = 0
    NUORIT% = 0
    Call APERBASDAT("STOCKS")
    NOCLOSE = 1: NOARRAS% = 1
    CI% = 0    ' texto libre
    CAN# = 1
    LOTE$ = ""
    MEDIDAS$ = ""
    NUOC& = XVALO(Text20)
    DOSEC$ = ""
    CANTI = (-1) * CAN#
    VUNIST# = XVALO(Text17)
    MONEVA% = 1
    '
    DOSELA$ = "P/F - PC." + TRIM$(Str$(NUOC&))
    DOSEC$ = "PC." + TRIM$(Str$(NUOC&))
    '
    REFEMO$ = RASOCLI$(NC%)
    '
    NUORIT% = 1
    DEPX% = DEPOEXPE%
    REMICOR$ = Mid$(Text11.TEXT, 1, 3)
    NURE1& = Val(Mid$(Text11.TEXT, 4))
    NURE2$ = TRIM$(Str$(NURE1&))
    SUC1% = Val(Label10)
    SUC2$ = TRIM$(Str$(SUC1%))
    REMICOR1$ = REMICOR$ + SUC2$ + "-" + NURE2$
    '
    GoSub GRAREMILI
    '
5   Call BAJALDISCO
    '
    GoSub PRIDOCREMI
    '
GoTo 999
    '
    '
GRAREMILI:
   Call ABRESTOCKS
   With Movisto
     .AddNew
        FECHA% = FECHANUM(Text9)
     !FECHMOV = CVDAT(FECHA%)
     !COD_INTE = 0
     !COD_EXTE = ""
     !Descrip = TRIM$(AJUSTI$(Text28, 64))
     !IdenLote = TRIM$(LOTE$)
        CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
        If CMO$ = "" Then CMO$ = "RT"
     !CodiMovi = CMO$
     !DoPriCor = REMICOR1$
     !DoPriLar = NROREMIT$
     !DoSecCor = DOSEC$
     !DoSecLar = DOSELA$
     !NuOrItem = NUORIT%
     !NUPEDCLI = XVALO(Text20)
     !ReferCor = REFEMO$
     !ValoUnit = VUNIST#
     !MoneVal = 1
       On Error Resume Next      ' POR SI NO ESTA DEFINIDO EL CAMPO
     !TipoCam = 1
       On Error GoTo 0
     !DEPOMOVI = DEPO%
     If NC% >= 0 Then
         !Nume_Clie = NC%
         !Nume_Prov = 0
       Else
         !Nume_Clie = 0
         !Nume_Prov = Abs(NC%)
     End If
     !CANTINGRE = 0
     !CANTSALID = 1
     !CantSaldo = 0
     !DeLaRemi = Text28
     !FeReMovi = Now
     !NumUsuar = USNBR% Mod 100
     .Update
   End With
   Movisto.Close

Return
    '
PRIDOCREMI:
    '
    PRINTERPORT$ = CONTROL$(IDENTER$, "PORTREM")
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
    If TRIM$(FORIPRE$) <> "" Then
      '
      'PREFIREM$ = CONTROL$("", "PREFIREM")
      PREFIREM$ = TRIM(Label10)
      If PREFIREM$ = "" Then PREFIREM$ = "0001"
      NROREMIT$ = TRIM$(Mid$(REMI$, 4))
      While Len(NROREMIT$) < 8: NROREMIT$ = "0" + NROREMIT$: Wend
      NROREMIT$ = PREFIREM$ + "-" + NROREMIT$
      '
      FEX = FF%
      FECHDOC$ = FECHATEX$(FEX)
      NUMEDOC$ = NROREMIT$
      TIPODOC$ = "Remito"
      '
      CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Text8.TEXT)
      CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = TRIM$(Text15.TEXT)
      CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Text12.TEXT)
      CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = TRIM$(Text17.TEXT)
      CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = TRIM$(Text14.TEXT)
      CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Text13.TEXT)
      CODPARAM$(7) = "DEPOSAL ": DOCPARAM$(7) = TRIM$(DEPOSAL$)
      CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = TRIM$(NROREMIT$)
      CODPARAM$(9) = "NRO-OCOM": DOCPARAM$(9) = TRIM$(Text20.TEXT)
      CODPARAM$(10) = "PATE-TRA": DOCPARAM$(10) = TRIM$(Text18.TEXT)
      CODPARAM$(11) = "N-AUTENT": DOCPARAM$(11) = TRIM$(NORVENTA)
      CODPARAM$(12) = "COND-TRA": DOCPARAM$(12) = TRIM$(Text19)
      CODPARAM$(13) = "LOCA-TRA": DOCPARAM$(13) = TRIM$(Text21)
      CODPARAM$(14) = "NRO-AENT": DOCPARAM$(14) = TRIM$(Text22)
      '\\\OT-07-0007-RG-08/01/07///
      HORA$ = HORAENT$(NC%)
      CODPARAM$(15) = "Z-PLANO": DOCPARAM$(15) = TRIM$(HORA$)
      CAPARDOC% = 15
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "CANTIDAD"
      CODTABU1$(4) = "NRO-LOTE"
      CODTABU1$(5) = "UNIMED"
      CODTABU1$(6) = "N-OCOMPR"
      CODTABU1$(7) = "ORD-ITEM"
      CODTABU1$(8) = "REF-MAT1"
      CODTABU1$(9) = "REF-MAT2"
      CACOLTAB1% = 9
      '
      DESTIDOC% = NC%
      CAITAB1% = 0
      '
      CI% = 0
      CAN# = 1
      CAC# = 0
      UNID$ = ""
      CANS$ = CSTRING$(MKD$(CAN#), 3, 12, 1, 2)
      NULOTE$ = ""
      NUOCO$ = Text20
      decrit$ = ""
      '
      AA1$ = ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")
      ACOLDEI% = Val(Mid$(AA1$, 2))
      Call FRATEXTO(Text28, ACOLDEI%)
      '
      CAITAB1% = CAITAB1% + 1
      CODDX$ = ""
      TETABU1$(1, CAITAB1%) = CODDX$
      TETABU1$(2, CAITAB1%) = RETEX$(1)
      TETABU1$(3, CAITAB1%) = CANS$
      TETABU1$(4, CAITAB1%) = NULOTE$
      TETABU1$(5, CAITAB1%) = UNID$
      TETABU1$(6, CAITAB1%) = NUOCO$
      TETABU1$(7, CAITAB1%) = Str$(CAITAB1%)
      TETABU1$(8, CAITAB1%) = NUOCOCLI$(Val(NUOCO$), CI%)
      TETABU1$(9, CAITAB1%) = NUPLANO$(CI%)
      '
      For KKJ% = 2 To CARETEX%
         CAITAB1% = CAITAB1% + 1
         For KKI% = 0 To CACOLTAB1%
             TETABU1$(KKI%, CAITAB1%) = ""
         Next KKI%
         '
         TETABU1$(2, CAITAB1%) = RETEX$(KKJ%)
      Next KKJ%
      '
      Call PRINTDOC("FOREMIFA")   ' Remito de Facturacion
      '
    End If
    '
Return
    '
    '
999 End Sub
'
