VERSION 5.00
Begin VB.Form RECIBO04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recibo de Cobranza en Cuenta Corriente"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11805
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox LINDICE 
      Height          =   1035
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   38
      Top             =   5460
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0F0&
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
      TabIndex        =   5
      Top             =   0
      Width           =   12000
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5880
         TabIndex        =   73
         Top             =   3920
         Width           =   1275
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Saldo Cancela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   2100
         TabIndex        =   71
         Top             =   7200
         Width           =   1830
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   210
            TabIndex        =   72
            Top             =   315
            Width           =   1485
         End
      End
      Begin VB.CommandButton Command23 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   10860
         Picture         =   "RECIBO04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Ver e Imprimir Copia de Recibo ya Emitido"
         Top             =   6420
         Width           =   840
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Descuento (N/Cred)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   4040
         TabIndex        =   52
         Top             =   7200
         Visible         =   0   'False
         Width           =   2820
         Begin VB.TextBox DESCUEN 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   1155
            MaxLength       =   14
            TabIndex        =   59
            Top             =   315
            Width           =   1485
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   105
            TabIndex        =   58
            Top             =   315
            Width           =   645
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   840
            TabIndex        =   53
            Top             =   300
            Width           =   330
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Anticipos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   8880
         TabIndex        =   56
         Top             =   7200
         Width           =   1815
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   210
            TabIndex        =   57
            Top             =   315
            Width           =   1485
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Aplicados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   6960
         TabIndex        =   54
         Top             =   7200
         Width           =   1815
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   210
            TabIndex        =   55
            Top             =   315
            Width           =   1485
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Neto Cobranza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   210
         TabIndex        =   50
         Top             =   7200
         Width           =   1815
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   210
            TabIndex        =   51
            Top             =   315
            Width           =   1485
         End
      End
      Begin VB.CommandButton Command5 
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
         Height          =   675
         Left            =   10860
         Picture         =   "RECIBO04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Aprueba, Graba e Imprime Documento"
         Top             =   7270
         Width           =   840
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
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   265
         Width           =   3165
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
         Height          =   675
         Left            =   5250
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "X"
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   315
         Width           =   435
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00004040&
         Height          =   6225
         Left            =   10860
         ScaleHeight     =   6165
         ScaleWidth      =   1005
         TabIndex        =   39
         Top             =   0
         Width           =   1065
         Begin VB.PictureBox Picture16 
            Height          =   6200
            Left            =   0
            ScaleHeight     =   6135
            ScaleWidth      =   1020
            TabIndex        =   77
            Top             =   0
            Visible         =   0   'False
            Width           =   1080
            Begin VB.PictureBox MARCOBARRA 
               BackColor       =   &H000000FF&
               Height          =   135
               Left            =   120
               ScaleHeight     =   75
               ScaleWidth      =   585
               TabIndex        =   78
               TabStop         =   0   'False
               Top             =   5880
               Width           =   650
               Begin VB.Frame BARRATRAZA 
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   0  'None
                  Height          =   500
                  Left            =   0
                  TabIndex        =   79
                  Top             =   0
                  Width           =   12000
               End
            End
            Begin VB.Image Image3 
               Height          =   6200
               Left            =   0
               Picture         =   "RECIBO04.frx":0614
               Stretch         =   -1  'True
               Top             =   0
               Width           =   1095
            End
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
            Left            =   125
            Picture         =   "RECIBO04.frx":29AC
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Abandona la Aplicación"
            Top             =   210
            Width           =   650
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
            Height          =   690
            Left            =   125
            Picture         =   "RECIBO04.frx":2AF6
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Notas de Clientes"
            Top             =   3885
            Width           =   650
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
            Height          =   690
            Left            =   125
            Picture         =   "RECIBO04.frx":2E00
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   1700
            Width           =   650
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
            Height          =   690
            Left            =   125
            Picture         =   "RECIBO04.frx":310A
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Apertura de Nueva Cuenta de Cliente"
            Top             =   3200
            Width           =   650
         End
         Begin VB.CommandButton Command6 
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
            Left            =   125
            Picture         =   "RECIBO04.frx":354C
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   4725
            Width           =   650
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
            Height          =   690
            Left            =   125
            Picture         =   "RECIBO04.frx":398E
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Continúa Carga de Recibo."
            Top             =   900
            Width           =   650
         End
         Begin VB.CommandButton Command8 
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
            Left            =   125
            Picture         =   "RECIBO04.frx":3C98
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Anular, Re-Imprimir, Imputación de Créditos, Tablas"
            Top             =   2380
            Width           =   650
         End
         Begin VB.Image Image2 
            Height          =   390
            Left            =   -300
            Picture         =   "RECIBO04.frx":3DE2
            Top             =   5780
            Width           =   1515
         End
      End
      Begin VB.CommandButton Command7 
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
         TabIndex        =   36
         Top             =   630
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00D0F0F0&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   435
         Left            =   10670
         TabIndex        =   37
         Top             =   630
         Width           =   1170
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   2355
         Left            =   210
         ScaleHeight     =   2295
         ScaleWidth      =   10395
         TabIndex        =   26
         Top             =   1480
         Width           =   10450
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   330
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   10515
            Begin VB.Line Line23 
               X1              =   8505
               X2              =   8505
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line22 
               X1              =   4935
               X2              =   4935
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "  M-P    Medio de Pago      Numero              P/T       Cl.         Emision           Vencim.               Importe"
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
               Top             =   50
               Width           =   10410
            End
            Begin VB.Line Line20 
               X1              =   5670
               X2              =   5670
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line13 
               X1              =   4200
               X2              =   4200
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line7 
               X1              =   2625
               X2              =   2625
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line6 
               X1              =   630
               X2              =   630
               Y1              =   0
               Y2              =   1155
            End
            Begin VB.Line Line5 
               X1              =   7140
               X2              =   7140
               Y1              =   0
               Y2              =   1050
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0F0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   2600
            Left            =   -315
            TabIndex        =   29
            Top             =   0
            Width           =   10800
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
               Height          =   2310
               Left            =   285
               TabIndex        =   30
               Top             =   315
               Width           =   10500
            End
            Begin VB.Line Line21 
               X1              =   1155
               X2              =   1155
               Y1              =   1470
               Y2              =   3255
            End
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         Height          =   2800
         Left            =   210
         ScaleHeight     =   2745
         ScaleWidth      =   10395
         TabIndex        =   16
         Top             =   4300
         Width           =   10450
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5630
            Left            =   0
            TabIndex        =   17
            Top             =   315
            Width           =   10515
            Begin VB.ListBox List2 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2085
               Left            =   -30
               TabIndex        =   18
               Top             =   0
               Width           =   10450
            End
            Begin VB.ListBox List3 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2085
               Left            =   -30
               TabIndex        =   33
               Top             =   0
               Width           =   10450
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Left            =   3570
               TabIndex        =   20
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Left            =   8295
               TabIndex        =   25
               Top             =   2100
               Width           =   1695
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Left            =   0
               TabIndex        =   24
               Top             =   0
               Width           =   1695
            End
            Begin VB.Line Line4 
               X1              =   8400
               X2              =   8400
               Y1              =   0
               Y2              =   2415
            End
            Begin VB.Line Line12 
               X1              =   1680
               X2              =   1680
               Y1              =   1575
               Y2              =   3360
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Left            =   6600
               TabIndex        =   21
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Total General"
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
               Left            =   25
               TabIndex        =   19
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Line Line14 
               BorderWidth     =   2
               X1              =   0
               X2              =   10395
               Y1              =   2085
               Y2              =   2085
            End
            Begin VB.Line Line15 
               X1              =   0
               X2              =   10450
               Y1              =   2415
               Y2              =   2415
            End
            Begin VB.Line Line16 
               X1              =   2520
               X2              =   2520
               Y1              =   2100
               Y2              =   2415
            End
            Begin VB.Line Line17 
               X1              =   3730
               X2              =   3730
               Y1              =   2100
               Y2              =   2415
            End
            Begin VB.Line Line18 
               X1              =   5460
               X2              =   5460
               Y1              =   2100
               Y2              =   2415
            End
            Begin VB.Line Line19 
               X1              =   6720
               X2              =   6720
               Y1              =   2100
               Y2              =   2415
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   330
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   10515
            Begin VB.Line Line1 
               X1              =   8400
               X2              =   8400
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line2 
               X1              =   6720
               X2              =   6720
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   " Documento             Nro      Emision                Importe       Vencim.                  Saldo                Acredit.       "
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
               TabIndex        =   23
               Top             =   50
               Width           =   10410
            End
            Begin VB.Line Line8 
               X1              =   1680
               X2              =   1680
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line9 
               X1              =   2520
               X2              =   2520
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line10 
               X1              =   3730
               X2              =   3730
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line11 
               X1              =   5460
               X2              =   5460
               Y1              =   0
               Y2              =   1050
            End
         End
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
         TabIndex        =   1
         Text            =   " "
         Top             =   1000
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
         TabIndex        =   3
         Text            =   " "
         Top             =   265
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
         TabIndex        =   4
         Top             =   265
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
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1000
         Width           =   2325
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
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1000
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
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   630
         Width           =   3165
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
         TabIndex        =   2
         Top             =   630
         Width           =   180
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
         Height          =   300
         Left            =   210
         TabIndex        =   0
         Text            =   " "
         Top             =   630
         Width           =   855
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Prom. de Pagos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1250
         Left            =   6300
         TabIndex        =   64
         Top             =   210
         Width           =   1665
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00D0F0F0&
            Caption         =   "Dias Atr:"
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
            Left            =   60
            TabIndex        =   70
            Top             =   945
            Width           =   855
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00D0F0F0&
            Caption         =   "Pr.Val:"
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
            Left            =   60
            TabIndex        =   69
            Top             =   630
            Width           =   855
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00D0F0F0&
            Caption         =   "Pr.venc:"
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
            Left            =   60
            TabIndex        =   68
            Top             =   315
            Width           =   855
         End
         Begin VB.Label Label23 
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
            Height          =   225
            Left            =   975
            TabIndex        =   67
            Top             =   945
            Width           =   540
         End
         Begin VB.Label Label19 
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
            Height          =   225
            Left            =   975
            TabIndex        =   66
            Top             =   630
            Width           =   540
         End
         Begin VB.Label Label17 
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
            Height          =   225
            Left            =   975
            TabIndex        =   65
            Top             =   315
            Width           =   540
         End
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Valores:"
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
         Left            =   6615
         TabIndex        =   76
         Top             =   3980
         Width           =   2115
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ajuste de Centavos:"
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
         Left            =   2520
         TabIndex        =   75
         ToolTipText     =   "Doble Click Para Ajuste Automatico"
         Top             =   3975
         Width           =   3270
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8820
         TabIndex        =   74
         Top             =   3920
         Width           =   1485
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "T-C:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10750
         TabIndex        =   62
         ToolTipText     =   "Doble Click para Configurar Cliente"
         Top             =   6300
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label MONEMIS 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pesos"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   11.25
            Charset         =   0
            Weight          =   900
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00005500&
         Height          =   330
         Left            =   10815
         TabIndex        =   61
         ToolTipText     =   "T.C.: 1.0000 - Doble-Click para Cambiar Moneda"
         Top             =   6810
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Height          =   120
         Left            =   3360
         TabIndex        =   60
         Top             =   7980
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label TICAMBIS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   10860
         TabIndex        =   49
         ToolTipText     =   "Doble Click para Ajustar Tipo de Cambio"
         Top             =   6500
         Visible         =   0   'False
         Width           =   840
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
         TabIndex        =   48
         Top             =   1000
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Docum. no Válido  como Factura"
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
         Left            =   4620
         TabIndex        =   15
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   1050
         Width           =   1800
      End
      Begin VB.Line Line30 
         X1              =   0
         X2              =   12000
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line28 
         X1              =   210
         X2              =   10660
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line27 
         X1              =   10665
         X2              =   10665
         Y1              =   1480
         Y2              =   3840
      End
      Begin VB.Line Line26 
         X1              =   10665
         X2              =   10665
         Y1              =   4305
         Y2              =   7090
      End
      Begin VB.Label Label22 
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
         TabIndex        =   35
         Top             =   630
         Visible         =   0   'False
         Width           =   3330
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tal:"
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
         Left            =   7840
         TabIndex        =   34
         Top             =   705
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   3780
         TabIndex        =   32
         Top             =   5775
         Width           =   1695
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Valores:"
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
         Left            =   1470
         TabIndex        =   31
         Top             =   5880
         Width           =   2220
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Numero:"
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
         TabIndex        =   13
         Top             =   1060
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
         TabIndex        =   12
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
         TabIndex        =   6
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
      TabIndex        =   11
      Top             =   0
      Width           =   960
   End
   Begin VB.Menu Archiv 
      Caption         =   "Archivos"
      Begin VB.Menu ReciCobran 
         Caption         =   "Nuevo Recibo de Cobranza"
      End
      Begin VB.Menu php0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuEdi 
      Caption         =   "Edición"
      Begin VB.Menu mnuAnuRec 
         Caption         =   "Anulación Recibo de Cobranza"
      End
      Begin VB.Menu php2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImpCre 
         Caption         =   "Imputación  de Créditos"
      End
      Begin VB.Menu mnuAjuSalDeu 
         Caption         =   "Ajuste de Saldos Deudores"
      End
      Begin VB.Menu compensacionclienteproveedor 
         Caption         =   "Compensación Cliente-Proveedor"
      End
      Begin VB.Menu REIMPUTACIONDECOMPROBANTESDEVENTAS 
         Caption         =   "Re-Imputación de Comprobantes de Ventas"
      End
   End
   Begin VB.Menu ver 
      Caption         =   "Ver"
      Begin VB.Menu mnuNoCli 
         Caption         =   "Notas de Clientes"
      End
      Begin VB.Menu mnuVerReimp 
         Caption         =   "Ver Recibos y Re-Imprimir"
      End
      Begin VB.Menu php1 
         Caption         =   "-"
      End
      Begin VB.Menu visualizarimputaciondecreditosyanticipos 
         Caption         =   "Visualizar Imputación de Créditos y Anticipos"
      End
   End
   Begin VB.Menu mnuConsuList 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuResCuCo 
         Caption         =   "Resumen de Cuenta Corriente"
      End
      Begin VB.Menu mnuLiCuCoes 
         Caption         =   "Listado de Cuentas Corrientes"
      End
      Begin VB.Menu php4a 
         Caption         =   "-"
      End
      Begin VB.Menu LiReReg 
         Caption         =   "Listado de Recibos Registrados"
      End
      Begin VB.Menu retencionesdeingresosbrutos 
         Caption         =   "Subdiario de Retenciones de Ingresos Brutos"
      End
      Begin VB.Menu atrasodecobranza 
         Caption         =   "Atraso de Cobranza"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu listadodeaplicaciondiferenciasdecambio 
         Caption         =   "Listado Aplicación Diferencias de Cambio"
      End
   End
   Begin VB.Menu mnuAccDir 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuMaeClie 
         Caption         =   "Maestro de Clientes"
      End
      Begin VB.Menu php4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOtAccDir 
         Caption         =   "Otros Accesos Directos"
      End
   End
   Begin VB.Menu mnuConfi 
      Caption         =   "Configuración"
      Begin VB.Menu mnuConfiGral 
         Caption         =   "Configuración General de Recibo"
      End
   End
End
Attribute VB_Name = "RECIBO04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$, AUTONUME%
Dim MODORECI%
Dim MODOMONEI%


Private Sub atrasodecobranza_Click()
    RECICO04.LISTADORECIBOS_SUBDIARIORETENCIONES_ATRASOCOBRANZA (3)
End Sub

Private Sub BOTNEXT_Click()
    '
    Call SETULTREG("!CARMECO", 0)
    NRONUE& = XVALO(Mid$(Text11, 4))
    If NRONUE& < 1 Or NRONUE& > 999999 Then
      Call MENSERR(24, "Número de Recibo Fuera de Rango")
      Call NURECIBO(NROPRO&)
      Call PRENURE(NROPRO&)
      Exit Sub
    End If
    '
    FECHALIM% = XVALO(CONTROL$("", "CIEMEVEN"))
    If FECHANUM(Text9) <= FECHALIM% Then
      Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
      Exit Sub
    End If
    '
    Call GRACONTROL("RECISTA", "COBRACTI", TRIM$(Label22.Caption))
    Call SIGUERECI1
    Command5.Enabled = True
    '
End Sub

Private Sub Command1_Click()
    '
    Call SELECHO("DEUDOR1")
    NCXI% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
    If NCXI% > 0 Then
      NC% = NCXI%
      '\\\OT-06-0607-OD-28/07/08///
      'EVALUA SI LA CUENTA NO ESTA SUSPENDIDA
      '
      If LIMICRE(NCXI%) < (-0.5) Then
          Call BorraEncabezado
          Call MENSERR(24, "Cuenta Suspendida.")
          Exit Sub
      End If
      '\\\OT-06-0607-OD-FIN///
        Text1.TEXT = TRIM$(Str$(NC%))
        Text2.TEXT = RASOCLI$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        Label26 = CODICLI$(NC%)
        Call NOTACLI(NC%, 1)
    End If
    '
End Sub

Sub BorraEncabezado()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Label26 = ""
End Sub

Private Sub Command10_Click()
    Call SELECHO("ECODEP")
    DPX1% = XVALO(VALACT1$("ECODEP"))
    If DPX1% > 0 Then
        DP1% = DPX1%
        Text10.TEXT = TRIM$(Str$(DP1%))
        Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    End If
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("ARCHIG04/AMACLI", "Archivos Maestros")
99  Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   For U& = 1 To ULTREG&("mecobra")
     XX$ = REGLEIDO$("MECOBRA", U&)
     TTXX$ = TTXX$ + Mid$(XX$, 45, 12)
   Next U&
   MsgBox TTXX$
End Sub

Private Sub Command2_Click()
    If List1.ListCount > 0 Or List2.ListCount > 0 Then
          Call CIERRARCH("!CARMECO")
          Call BLANARCH("!CARMECO")
          Call CIERRARCH("!APLIRECI")
          Call BLANARCH("!APLIRECI")
          List1.Clear
          List2.Clear
          Label4.Caption = ""
          Label8.Caption = ""
          Label10.Caption = ""
          Label14.Caption = "": Label30 = "": Text8 = ""
          Label18.Caption = ""
          Label20.Caption = ""
          '\\\OT-07-0029-RG-26/01/07///
          DESCUEN(0) = ""
          Text7 = ""
          Frame8.Caption = "Descuento (N./Cred)"
          '\\\OT-07-0029-RG-FIN///
          BOTNEXT.Enabled = True
          Command5.Enabled = False
          Text1.SetFocus
          Call LIMPPROM
        Else
          If TRIM$(CONTROL$("RECISTA", "AVAHOJA")) = "NEWPAGE" Then
              'Printer.EndDoc
              Call SETSPOOL("ENDDOC", "")
          End If
          Call CIERRARCH("*.*")
          Call BAJALDISCO
          Hide
    End If
End Sub

Private Sub Command23_Click()
   Command23.Enabled = False
   Call RECICO04.Command23_Click
   Command23.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NCLI$ = TRIM$(Text1.TEXT)
    'Call CONECRUN("*NOTACLI/" + NCLI$, "Notas de Clientes")
    NC% = XVALO(Text1.TEXT)
    If NC% > 0 Then Call NOTACLI(NC%, 2)
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RECUEC04.Text1.TEXT = Text1.TEXT
    RECUEC04.Show 1
    Command4.Enabled = True
End Sub

Function VALIDARECIBO%()
    '
    Screen.MousePointer = 11   ' Grabando Recibo - Espere un Momento, por Favor"
    VDRX% = 0
    '
    ' NUMERO DE CLIENTE EN TEXT1
    NC% = XVALO(Text1)
    If NC% <> XVALO(Text1.TEXT) Then
       Call MENSERR(24, "Imposible Grabar - Intente de Nuevo")
       Call CIERRARCH("*.*")
       Call FINAL("")
    End If
    APRORECIB.Label1 = TRIM$(Str$(NC%))
    APRORECIB.Label2 = RASOCLI$(NC%)
    '
    ' BALANCE ENTRE VALORES Y CANCELACIONES
    TOTAVAL# = CDbl(XVALO(Label14.Caption)) + CDbl(XVALO(DESCUEN(0)))
    TOTACAN# = CDbl(XVALO(Label18.Caption))
    If ROUND#(TOTACAN#, 2) > ROUND#(TOTAVAL#, 2) Then
        Call COMUNI("Operación Incorrecta.\  \Las Cancelaciones no Pueden\Exceder la Suma de Valores Recibidos.")
        VDRX% = (-1)
        GoTo 99
    End If
    '
    If TOTACAN# < (-0.005) Then
        Call MENSERR(24, "Valores no Procesables")
        VDRX% = (-1)
        GoTo 99
    End If
    '
    If XVALO(Label20) >= 0.005 Then
      If XVALO(DESCUEN(0)) >= 0.005 Then
        Call MENSERR(24, "Para Generar Descuento\Debe Aplicarse la Totalidad del Pago.")
        VDRX% = (-1)
        GoTo 99
      End If
    End If
    '
    APRORECIB.Label12 = Label14
    '
    ANTI# = CDbl(XVALO(Label20.Caption))
    APRORECIB.Label14 = FORMATNUM$(ANTI#, "F12.2")
    '
    ' FECHA DE EMISIÓN DEL RECIBO
    FECHALIM% = XVALO(CONTROL$("", "CIEMEVEN"))
    FLIMI = CVDAT(FECHALIM%)
    FF% = FECHANUM(Text9.TEXT)
    If FF% < 1 Then
        Call MENSERR(24, "Fecha de Emisión no Válida")
        VDRX% = (-1)
        GoTo 99
    End If
    '
    If FF% <= FECHALIM% Then
        Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
        VDRX% = (-1)
        GoTo 99
    End If
    FEX = FF%
    APRORECIB.Label6 = FECHATEX$(FEX)
    '
    Call SETULTREG("!PRINRECI", 0)
    RENDISPO& = XVALO(CONTROL$("RECISTA", "CARENDET"))
    If RENDISPO& < 1 Then RENDISPO& = 9999
    If RENGLONEC& > RENDISPO& Then
      Call MENSERR(24, "Imposible Registrar este Recibo.\   \Excede Renglones Disponibles\en Formulario de Impresión.")
      VDRX% = (-1)
      GoTo 99
    End If
    '
    NROO& = XVALO(Mid$(TRIM$(Text11.TEXT), 4))
    If Val(NROO&) < 1 Then
      Call MENSERR(24, "Número de Recibo no Válido")
      VDRX% = (-1)
      GoTo 99
    End If
    APRORECIB.Label4 = TRIM$(Str$(NROO&))
    '
    COBRACT% = 0
    On Error Resume Next
    COBRACT% = XVALO(Label22.Caption)
    On Error GoTo 0
    APRORECIB.Label3 = TRIM$(Str$(COBRACT%))
    '
    CUECARTE% = CUEMECOB%(-1)
    APRORECIB.Label16 = TRIM$(Str$(CUECARTE%))
    APRORECIB.Label17 = ECOARCH$("CUECOT", MKI$(CUECARTE%))
    '
    APRORECIB.Label29 = TRIM$(Mid$(Frame8.Caption, 9))
      NFCY$ = APRORECIB.Label29
    APRORECIB.Label31 = Left$(NFCY$, 3) + Mid$(NFCY$, 11)
      NEFA7# = ROUND#(XVALO(DESCUEN(0)) / (1 + ALIVA / 100), 2)
      NIVA8# = XVALO(DESCUEN(0)) - NEFA7#
    APRORECIB.Label27 = TRIM$(FORMATNUM$(NEFA7#, "F12.2"))
    APRORECIB.Label25 = TRIM$(FORMATNUM$(NIVA8#, "F12.2"))
      CUEDESCU$ = TRIM$(CONTROL$("", "CUEDESCU"))
      CUECONDE% = CUECOINT%(CUEDESCU$)
    APRORECIB.Label22 = TRIM$(Str$(CUECONDE%))
    APRORECIB.Label21 = ECOARCH$("CUECOT", MKI$(CUECONDE%))
    '
    If NEFA7# >= 0.005 Then
      If CUECONDE% < 1 Or TRIM$(APRORECIB.Label21) = "" Then
        Call MENSERR(24, "Imposible Grabar N/Cred. Descuento.\   \Falta Definir Cuenta Contable.")
        VDRX% = (-1)
        GoTo 99
      End If
    End If
    '
    'BLOQUEA ARCHIVO DE COBRANZA
    Call BLOQARCH("CUECORR")
    '
    RECICOR$ = TRIM$(Str$(COBRACT%))
    While Len(RECICOR$) < 2: RECICOR$ = "0" + RECICOR$: Wend
    RECICOR$ = TRIM$(Str$(NROO&)) + "-" + RECICOR$
    While Len(RECICOR$) < 9: RECICOR$ = "0" + RECICOR$: Wend
    RECICOR$ = "RB-" + RECICOR$
    APRORECIB.Label8 = RECICOR$
    '
    RECILAR$ = TRIM$(Str$(NROO&))
    While Len(RECILAR$) < 8: RECILAR$ = "0" + RECILAR$: Wend
    RECILAR$ = TRIM$(Str$(COBRACT%)) + "-" + RECILAR$
    While Len(RECILAR$) < 13: RECILAR = "0" + RECILAR$: Wend
    RECILAR$ = "RB-X-" + RECILAR$
    APRORECIB.Label10 = RECILAR$
    '
    FENUE$ = CONTROL$("", "FENURECI")
    If FENUE$ = "" Then FENUE$ = "01/01/80"
    FLIMI = CVDAT(FECHANUM(FENUE$))
    If FLIMI > Now Then FLIMI = CVDAT(HOY(HO$))
    FLIMI1 = CDbl(FLIMI)
    '
    SQLX$ = "SELECT NumeComp FROM CajaBanc "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + "AND (CodiCom_Cor = '" & RECICOR$ & "' "
    SQLX$ = SQLX$ + "OR CodiCom_Lar = '" & RECILAR$ & "') "
    SQLX$ = SQLX$ + "AND Status < 9 "
    SQLX$ = SQLX$ + "AND CDBL(FechEmisi) >= " & FLIMI1 & " "
    Set OrpaTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    With OrpaTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Call LIBARCH("CUECORR")
        Call MENSERR(24, "Al Grabar este Recibo\el Número Propuesto Estaba Ocupado.\  \Verifique el Nuevo Número Propuesto\y Vuelva a Grabar e Imprimir.")
        Call NURECIBO(NROPRO&)
        Call PRENURE(NROPRO&)
        VDRX% = (-1)
        GoTo 99
      End If
      On Error GoTo 0
    End With
    '
    If NEFA7# >= 0.005 Then
      If FORFACCS.VERINUFA%(NFCY$) < 1 Then
        Call LIBARCH("CUECORR")
        Call MENSERR(24, "Al Grabar este Recibo\el Número de N/Cred. Estaba Ocupado.\  \Verifique el Nuevo Número Propuesto\y Vuelva a Grabar e Imprimir.")
        Frame8.Caption = "N./Cred. " + PRONOCRE$
        VDRX% = (-1)
        GoTo 99
      End If
    End If
    '
    VDRX% = 1
    '
99  VALIDARECIBO% = VDRX%
    Screen.MousePointer = 1
    '
End Function
'
Private Sub Command5_Click()
    '
    Command5.Enabled = False
    If VALIDARECIBO% <> 1 Then
       Call LIBARCH("CUECORR")
       Command5.Enabled = True
       Exit Sub
    End If
    '
    Call GRARECISTA(OKEY%)
    Call LIBARCH("CUECORR")
    '
    If OKEY% = 0 Then
        Command5.Enabled = True
        Exit Sub
      ElseIf OKEY% <> 6 Then
        Call COMUNI("La Registración de este Recibo\no Pudo Completarse. (" + TRIM$(Str$(OKEY%)) + ")\  \Consulte a su Soporte de Sistemas FLEXOFT.")
        Call FINAL("")
    End If
    
    On Error GoTo 0
    HOJASUEL% = 1
    If AUTONUME% = 1 Then HOJASUEL% = 0
    Call PRIRECISTA
    '
    If Val(DESCUEN(0)) >= 0.005 Then
       Call PRICREDESC
    End If
    '
    Call CIERRARCH("!CARMECO")
    Call BLANARCH("!CARMECO")
    Call CIERRARCH("!APLIRECI")
    Call BLANARCH("!APLIRECI")
    List1.Clear
    List2.Clear
    Label4.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    Label14.Caption = "": Label30 = "": Text8 = ""
    Label18.Caption = ""
    Label20.Caption = ""
    '\\\OT-07-0029-RG-26/01/07///
    DESCUEN(0) = ""
    Text7 = ""
    Frame8.Caption = "Descuento (N./Cred)"
    '\\\OT-07-0029-RG-FIN///
    BOTNEXT.Enabled = True
    Command5.Enabled = False
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    
    NROPRO& = 0
    Call NURECIBO(NROPRO&)
    Call PRENURE(NROPRO&)
    '
    'Text11 = "RB.      "
    Text11.SelStart = 3
    Text11.SelLength = Len(Text11.TEXT) - 3
    Text11.SetFocus
    '
End Sub

Sub PRENURE(NROPRO&)
    '
    Static MODONUM$
    If MODONUM$ = "" Then
      MODONUM$ = CONTROL$("", "NURECIST")
    End If
    '
    TNUME$ = "RB.      "
    'If MODONUM$ = "AUTONUME" Then TNUME$ = TRIM$(TNUME$) + TRIM$(Str$(NROPRO&))
    TNUME$ = TRIM$(TNUME$) + TRIM$(Str$(NROPRO&))
    Text11.TEXT = TNUME$
    '
End Sub
'
Sub Command6_Click()
    Command6.Enabled = False
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    OPRECIB00.Show 1
'    Call CONECRUN("FSETUP04/OPRECIBO", "Configuración de Funcionamiento")
    Call FINAL("")
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    '
    If CTX% = 0 Then
       CTX% = 1
       If ULTREG&("COBRADOR") > 1 Then
          XXX$ = REGSEL$("COBRADOR")
          If Len(XXX$) < 64 Then
             Call FINAL("")
          End If
          Label22.Caption = Left$(TRIM$(Asc(Left$(XXX$, 1))) + " - " + Mid$(XXX$, 2, 30), 19)
          NROPRO& = 0
          Call NURECIBO(NROPRO&)
          Call PRENURE(NROPRO&)
          'Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
          Text11.SelStart = 3
          Text11.SelLength = Len(Text11.TEXT) - 3
       End If
    End If
    '
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   RECICO04.Show 1
   '
   NROPRO& = 0
   Call NURECIBO(NROPRO&)
   Call PRENURE(NROPRO&)
   'Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
   Text11.SelStart = 3
   Text11.SelLength = Len(Text11.TEXT) - 3
   '
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

Private Sub compensacionclienteproveedor_Click()
    Call RECICO04.Command10_Click
    
End Sub

Private Sub DESCUEN_Change(index As Integer)
    '
    Label20.Caption = CSTRING$(MKD$(XVALO(Label14) + XVALO(DESCUEN(0)) - XVALO(Label18)), 4, 14, 2, 2)
    If XVALO(Label20) < 0 Then
          Frame10.Caption = "Descuentos:"
        Else
          Frame10.Caption = "Anticipos:"
    End If
    Call Timer1_Timer
    '
End Sub

Private Sub DESCUEN_GotFocus(index As Integer)
    If TRIM$(Text1) = "" Then Text1.SetFocus
    NCX1 = XVALO(Text1)
    If NCX1 < 1 Or NCX1 > 32767 Then Text1.SetFocus
    NCX2% = NCX1
    If ECOARCH$("DEUDOR1", MKI$(NCX2%)) = "" Then Text1.SetFocus
    DESCUEN(index) = TRIM$(DESCUEN(index))
End Sub

Private Sub DESCUEN_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
      Call DESCUEN_LostFocus(index)
    End If
End Sub

Private Sub DESCUEN_LostFocus(index As Integer)
    FDESCU$ = TRIM$(FORMATNUM(XVALO(DESCUEN(index)), "F12.2"))
    PDX = 0
    IDES# = XVALO(DESCUEN(0))
    L14# = XVALO(Label14)
    If L14# >= 1 Then
      PDX = 100 - 100 / (IDES# / L14# + 1)
    End If
    Text7 = TRIM$(FORMATNUM$(PDX, "F6.1"))
    DESCUEN(index) = FDESCU$
End Sub

Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Frame7.BackColor = &HFCD7B6
      Me.Frame8.BackColor = &HFCD7B6
      Me.Frame9.BackColor = &HFCD7B6
      Me.Frame10.BackColor = &HFCD7B6
      Me.Frame11.BackColor = &HFCD7B6
      Me.Frame12.BackColor = &HFCD7B6
      
      Me.Label25.BackColor = &HFCD7B6
      Me.Label27.BackColor = &HFCD7B6
      Me.Label28.BackColor = &HFCD7B6
      
      MARCOBARRA.Top = 5900
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
      
      BOTNEXT.Top = 5100
      Set BOTNEXT.Container = Me.Picture16
      BOTNEXT.ZOrder 0
      
      
    End If

    Dim BUFLE As String * 64
    '
    Call CENTRAFORM(Me)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Refresh
    Frame1.Refresh
    DoEvents
    '\\\OT-07-0029-RG-26/01/07///
    If CONTROL$("RECISTA", "DESCAJA") = "CREDESCU" Then
       Frame9.Left = 6960
       Frame8.Visible = True
    End If
    '\\\OT-07-0029-RG-FIN///
    '
    MONEMI% = 1
    MONEMIS = DEMONECO$(1)
    TICAMBIS = "1.0000"
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
       Command23.Visible = False
       MONEMIS.Visible = True
       Label24.Visible = True
       TICAMBIS.Visible = True
    End If
    '
10  FORIPRE$ = TRIM$(CONTROL$("", "FORECIST"))
    If FORIPRE$ <> "" Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
        RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
        Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Recibo Standard de Cobranza:\  \'" + RMCC$ + "'.")
        OPRECIB00.Show 1
        'Call CONECRUN("FSETUP04/OPRECIBO", "Configuración de Funcionamiento")
        Call CIERRARCH("*.*")
        Call FINAL("")
      End If
    End If
    '
    CUEDEUVE$ = TRIM$(CONTROL$("AMACLI", "ECUECON"))
    If CUEDEUVE$ = "" Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Correspondiente a 'Deudores por Ventas'.")
      OPRECIB00.Show 1
'      Call CONECRUN("FSETUP04/OPRECIBO", "Configuración de Funcionamiento")
      Call CIERRARCH("*.*")
      Call FINAL("")
    End If
    If ECOARCH$("ECUECON", CUEDEUVE$) = "" Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Correspondiente a 'Deudores por Ventas'.")
      OPRECIB00.Show 1
      'Call CONECRUN("FSETUP04/OPRECIBO", "Configuración de Funcionamiento")
      Call CIERRARCH("*.*")
      Call FINAL("")
    End If
    '
    CUECHECA$ = TRIM$(CONTROL$("", "CUECHECA"))
    If CUECHECA$ = "" Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a la Cartera de Valores.")
      OPRECIB00.Show 1
'      Call CONECRUN("FSETUP04/OPRECIBO", "Configuración de Funcionamiento")
      Call CIERRARCH("*.*")
      Call FINAL("")
    End If
    If ECOARCH$("ECUECON", CUECHECA$) = "" Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a la Cartera de Valores.")
      OPRECIB00.Show 1
      'Call CONECRUN("FSETUP04/OPRECIBO", "Configuración de Funcionamiento")
      Call CIERRARCH("*.*")
      Call FINAL("")
    End If
    '
    If Frame8.Visible = True Then
       CUEDESCU$ = TRIM$(CONTROL$("", "CUEDESCU"))
       If CUEDESCU$ = "" Then
         Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Descuentos de Caja.")
         OPRECIB00.Show 1
         'Call CONECRUN("FSETUP04/OPRECIBO", "Configuración de Funcionamiento")
         Call CIERRARCH("*.*")
         Call FINAL("")
       End If
       If ECOARCH$("ECUECON", CUEDESCU$) = "" Then
         Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Descuentos de Caja.")
         OPRECIB00.Show 1
         'Call CONECRUN("FSETUP04/OPRECIBO", "Configuración de Funcionamiento")
         Call CIERRARCH("*.*")
         Call FINAL("")
       End If
     End If
     '
20  If ULTREG&("MECOBRA") < 1 Then
      TTXX$ = "Falta Tabla de Medios de Cobranza\Puede Instalar una Tabla Preconfigurada"
      OPX% = COMALTER%(TTXX$ + "\\Instalar Ahora\Carga Manual")
      If OPX% = 1 Then
         NX1% = FILEOPEN%(CurDir$ + "\MECOBRA.DRV", 0, 64)
         For U& = 1 To LOF(NX1%) / 64
           Get #NX1%, U&, BUFLE$
           NMP% = CVI(Left$(BUFLE$, 2))
           If NMP% > 0 Then
             If NMP% < 900 Then
               XXX$ = BUFLE$
               Call REPLA(XXX$, CUECHECA$, 45, 12)
               Call REGAPP("MECOBRA", XXX$)
             End If
           End If
         Next U&
         Call CIERRARCH("MECOBRA")
         Close #NX1%
         OPRECIB00.Show 1
         'Call CONECRUN("FSETUP04/OPRECIBO", "Configuración de Funcionamiento")
         Call CIERRARCH("*.*")
         Call FINAL("")
       ElseIf OPX% = 2 Then
         OPRECIB00.Show 1
         'Call CONECRUN("FSETUP04/OPRECIBO", "Configuración de Funcionamiento")
         Call CIERRARCH("*.*")
         Call FINAL("")
      End If
      GoTo 20
    End If
    '
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    If ULTREG&("COBRADOR") > 0 Then
       Label21.Visible = True
       Label22.Visible = True
       Command7.Visible = True
       XXX$ = REGLEIDO$("COBRADOR", 1)
       RECIBO04.Label22.Caption = Left$(TRIM$(Asc(Left$(XXX$, 1))) + " - " + Mid$(XXX$, 2, 30), 19)
    End If
    '
    AUTONUME% = 0
    If CONTROL$("", "NURECIST") = "AUTONUME" Then
      AUTONUME% = 1
    End If
    '
    NROPRO& = 0
    Call NURECIBO(NROPRO&)
    Call PRENURE(NROPRO&)
    'Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
    Text11.SelStart = 3
    Text11.SelLength = Len(Text11.TEXT) - 3
    '
    If AUTONUME% <> 0 Then
      Text1.TEXT = "     "
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
    End If
    '
    Refresh
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BLOQEXE("*")
     Call LIBARCH("*.*")
     Reset
     Call BLOQEXE(App.EXEName)
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub Label10_Change()
   '\\\OT-08-0203-OD-29/04/08///
   PROVENCI# = 0
   FIN = List2.ListCount - 1
   For U& = 0 To FIN
     Z$ = List2.LIST(U&)
     FEVEN$ = Mid$(Z$, 47, 8)
     IMPRT# = XVALO(Mid$(Z$, 70, 14))
     ResultDientre = DIENTRE(FECHANUM(FEVEN$), HOY(HO$))
     If ResultDientre > 0 Then
       PROVENCI# = PROVENCI + (DIENTRE(FECHANUM(FEVEN$), HOY(HO$)) * IMPRT#)
      Else
       PROVENCI# = PROVENCI + (DIENTRE(HOY(HO$), FECHANUM(FEVEN$)) * IMPRT#) * (-1)
     
     End If
   Next U&
   TOTAPLI# = XVALO(TRIM$(Label10))
   PROMVEN# = 0
   If Abs(TOTAPLI#) > 0.1 Then
     PROMVEN# = (PROVENCI# / TOTAPLI#)
   End If
   Label17 = FORMATNUM(PROMVEN#, "F4.0")
   If CONTROL("", "OPCVALOR") = "NO" Then
      PORDEX = XVALO(Text7): If PORDEX > 99 Then PORDEX = 99
      If PORDEX < 0 Then PORDEX = 0
      IDES# = XVALO(Label10) * PORDEX / 100
      DESCUEN(0) = CSTRING$(MKD$(IDES#), 4, 14, 2, 2)
   End If
   
   '\\\OT-08-0203-OD-FIN///
End Sub

Private Sub Label14_Change()
   '\\\OT-07-0029-RG-26/01/07///
   TT77$ = REPLACAR$(TRIM$(Text7), ",", ".")
   PORDEX = Val(Text7): If PORDEX > 99 Then PORDEX = 99
   If PORDEX < 0 Then PORDEX = 0
   IDES# = Val(Label14) * (100 / (100 - PORDEX) - 1)
   DESCUEN(0) = CSTRING$(MKD$(IDES#), 4, 14, 2, 2)
   '\\\OT-07-0029-RG-FIN///
   Label20 = TRIM$(FORMATNUM$(XVALO(Label14) + XVALO(DESCUEN(0)) - XVALO(Label18), "F12.2"))
End Sub

Private Sub Label17_Change()
Label23 = FORMATNUM$(XVALO(Label17) + XVALO(Label19), "F4.0")
If XVALO(Label23) < 0 Then Label23 = 0
End Sub

Private Sub Label19_Change()
Label23 = FORMATNUM$(XVALO(Label17) + XVALO(Label19), "F4.0")
End Sub

Private Sub Label2_DblClick()
   OPRECIB00.Show 1
   'Call CONECRUN("FSETUP04/OPRECIBO", "Configuración de Funcionamiento")
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Label22_DblClick()
    '
    If CTX% = 0 Then
       CTX% = 1
       If ULTREG&("COBRADOR") > 1 Then
          XXX$ = REGSEL$("COBRADOR")
          If Len(XXX$) < 64 Then
             Call FINAL("")
          End If
          Label22.Caption = Left$(TRIM$(Asc(Left$(XXX$, 1))) + " - " + Mid$(XXX$, 2, 30), 19)
          NROPRO& = 0
          Call NURECIBO(NROPRO&)
          Call PRENURE(NROPRO&)
          'Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
          Text11.SelStart = 3
          Text11.SelLength = Len(Text11.TEXT) - 3
       End If
    End If
    '
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
         Call NOTACLI(NC%, 1)
       End If
     End If
   End If
End Sub

Private Sub Label32_DblClick()
   If Abs(XVALO(Label29)) < 1 Then
      Text8 = TRIM$(Label29)
   End If
End Sub

Private Sub Label8_Change()
    Call Timer1_Timer
End Sub

Private Sub LiReReg_Click()
    RECICO04.LISTADORECIBOS_SUBDIARIORETENCIONES_ATRASOCOBRANZA (1)
End Sub

Private Sub List1_DBLClick()
    '
    Call SIGUERECI2
    '
End Sub

Private Sub List2_Click()
    '
    If VACON% <> 0 Then
        Exit Sub
    End If
    '
    Z$ = List2.TEXT
    LIST2ANT$ = Z$
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    ACREMAX# = CDbl(XVALO(Label14.Caption)) + CDbl(XVALO(DESCUEN(0))) - CDbl(XVALO(Label18.Caption))
    If ACREMAX# < 0 Then ACREMAX# = 0
    IPEN$ = Mid$(Z$, 56, 14)
    If XVALO(IPEN$) > ACREMAX# Then
        IPEN$ = CSTRING$(MKD$(ACREMAX#), 3, 14, 2, 2)
    End If
    SUMACRE# = XVALO(Label18.Caption)
    If IACRE$ = "" Then
         Call REPLA(Z$, IPEN$, 70, 14)
         SUMACRE# = SUMACRE# + CDbl(XVALO(IPEN$))
       Else
         Call REPLA(Z$, Space$(14), 70, 14)
         SUMACRE# = SUMACRE# - CDbl(XVALO(IACRE$))
    End If
    VACON% = 1
    List2.AddItem Z$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    List2.Refresh
    List2.ListIndex = ILISTA&
    VACON% = 0
    '
    Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label10.Refresh
    Label18.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label18.Refresh
    '
    '\\\OT-07-0029-RG-26/01/07///
    ANTI# = CDbl(XVALO(Label14.Caption)) + XVALO(DESCUEN(0)) - SUMACRE#
    If ANTI# < 0 Then
          Frame10.Caption = "Descuentos:"
        Else
          Frame10.Caption = "Anticipos:"
    End If
    '\\\OT-07-0029-RG-FIN///
    Label20.Caption = TRIM$(CSTRING$(MKD$(ANTI#), 4, 14, 2, 2))
    Label20.Refresh
    '
End Sub
'

Private Sub List2_DblClick()
    '
    VACON% = 1
    ILIST2& = List2.ListIndex
    TOLIST2& = List2.TopIndex
    '
    Z$ = List2.TEXT
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    IPEN$ = Mid$(Z$, 56, 14)
    SUMACRE# = XVALO(Label18.Caption)
    SUMACRE# = SUMACRE# - CDbl(XVALO(IACRE$))
    SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(LIST2ANT$, 70, 14)))
    List2.Visible = False
    List2.AddItem LIST2ANT$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    '
    Call SIGUERECI3
    '
    VACON% = 1
    List2.TopIndex = TOLIST2&
    List2.Visible = True
    If List2.ListCount > ILIST2& Then
          List2.ListIndex = ILIST2&
        Else
          Call CIERRARCH("!CARMECO")
          Call BLANARCH("!CARMECO")
          Call CIERRARCH("!APLIRECI")
          Call BLANARCH("!APLIRECI")
          List1.Clear
          List2.Clear
          Label4.Caption = ""
          Label8.Caption = ""
          Label10.Caption = ""
          Label14.Caption = "": Label30 = "": Text8 = ""
          Label18.Caption = ""
          Label20.Caption = ""
          '\\\OT-07-0029-RG-26/01/07///
          DESCUEN(0) = ""
          Text7 = ""
          Frame8.Caption = "Descuento (N./Cred)"
          '\\\OT-07-0029-RG-FIN///
          BOTNEXT.Enabled = True
          Command5.Enabled = False
          Text1.SetFocus
    End If
    VACON% = 0
    '
End Sub

Private Sub listadodeaplicaciondiferenciasdecambio_Click()
    Call RECICO04.Command6_Click
End Sub

Private Sub mnuAjuSalDeu_Click()
    Call RECICO04.Command4_Click
End Sub

Private Sub mnuAnuRec_Click()
    Call RECICO04.Command25_Click
End Sub

Private Sub mnuConfiGral_Click()
    Call Command6_Click
End Sub

Private Sub mnuImpCre_Click()
     Call RECICO04.Command3_Click
End Sub

Private Sub mnuLiCuCoes_Click()
    Call RECUEC04.Command3_Click
End Sub

Private Sub mnuMaeCli_Click()
    Call Command11_Click
End Sub

Private Sub mnuMaeClie_Click()
    Call Command11_Click
End Sub

Private Sub mnuNoCli_Click()
    Call Command3_Click
End Sub

Private Sub mnuOtAccDir_Click()
    Call COMUNI("No Existen Accesos Directos \ para esta Aplicación")
End Sub

Private Sub mnuResCuCo_Click()
    Call Command4_Click
End Sub

Private Sub mnuSaInCuCor_Click()
    Call RECICO04.Command5_Click
End Sub

Private Sub mnuSalir_Click()
    Call Command2_Click
End Sub

Private Sub mnuVerReimp_Click()
     Call RECICO04.Command23_Click
End Sub

Private Sub MONEMIS_Change()
    '
    '\\\OT-07-0086-RG-06/03/07///
    ENTRATICA.Label6.Caption = MONEMIS.Caption
    ENTRATICA.Text4 = TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4"))
    '\\\OT-07-0086-RG-FIN///
    '
End Sub

Private Sub MONEMIS_DblClick()
  Static CAMON%, DMONE$(), TCMONE()
   ' Exit Sub    ' PORQUE TODAVÍA NO ESTA TERMINADO BIMONETARIO
   If CAMON% < 1 Then
     CAMON% = ULTREG&("TAMONED")
10   If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         U2& = 0
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
           If TRIM$(DMX$) <> "" Then
             VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
             U2& = U2& + 1
             DMONE$(U2&) = DMX$
             TCMONE(U2&) = VHX
           End If
         Next U&
         CAMON% = U2&
         If CAMON% < 1 Then GoTo 10
     End If
   End If
   '
   MONEMIA% = MONEMI%
   MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   MONEMIS = DMONE$(MONEMI%)
   MONEMIS.ForeColor = RGB(0, 0, 255)
   MONEMIS.FontSize = 11
   If MONEMI% > 1 Then
      MONEMIS.ForeColor = RGB(255, 0, 0)
      MONEMIS.FontSize = 10
   End If
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   '
   COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
   TICAMBX# = TICAMONE(MONEMI%)
   TICAMBIS = FORMATNUM(TICAMBX#, "F9.4")
   '
   List1.Clear
   List2.Clear
   Label4.Caption = ""
   Label8.Caption = ""
   Label10.Caption = ""
   Label14.Caption = "": Label30 = "": Text8 = ""
   Label18.Caption = ""
   Label20.Caption = ""
   DESCUEN(0) = ""
   Text7 = ""
   Frame8.Caption = "Descuento (N./Cred)"
   BOTNEXT.Enabled = True
   '
End Sub


Private Sub ReciCobran_Click()
    Call Command1_Click
End Sub

Private Sub reimputaciondecomprobantesdeventas_Click()
   Call RECICO04.Command1_Click
End Sub

Private Sub retencionesdeingresosbrutos_Click()
    RECICO04.LISTADORECIBOS_SUBDIARIORETENCIONES_ATRASOCOBRANZA (2)
End Sub

Private Sub Text1_Change()
    List1.Clear
    List2.Clear
    Label4.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    Label14.Caption = "": Label14 = "": Text8 = ""
    Label18.Caption = ""
    Label20.Caption = ""
    Label17.Caption = ""
    Label19.Caption = ""
    Label23.Caption = ""
    '\\\OT-07-0029-RG-26/01/07///
    DESCUEN(0) = ""
    Text7 = ""
    Frame8.Caption = "Descuento (N./Cred)"
    '\\\OT-07-0029-RG-FIN///
    BOTNEXT.Enabled = True
    Command5.Enabled = False
    Call SETULTREG("!CARMECO", 0)
End Sub
Private Sub Text1_DblClick()
    '
    Call SELECHO("DEUDOR1")
    NCXI% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
    If NCXI% > 0 Then
        NC% = NCXI%
        '\\\OT-06-0607-OD-28/07/08///
        'EVALUA SI LA CUENTA NO ESTA SUSPENDIDA
        '
        If LIMICRE(NC%) < (-0.5) Then
          Call BorraEncabezado
          Call MENSERR(24, "Cuenta Suspendida.")
          Exit Sub
        End If
        '\\\OT-06-0607-OD-FIN///

        Text1.TEXT = TRIM$(Str$(NC%))
        Text2.TEXT = RASOCLI$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        Label26 = CODICLI$(NC%)
        Call NOTACLI(NC%, 1)
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
       If XVALO(Text1.TEXT) > 0 Then
          If XVALO(Text1.TEXT) <= 32767 Then
             NC% = XVALO(Text1.TEXT)
             '\\\OT-06-0607-OD-28/07/08///
             'EVALUA SI LA CUENTA NO ESTA SUSPENDIDA
             '
              If LIMICRE(NC%) < (-0.5) Then
                  Call BorraEncabezado
                  Call MENSERR(24, "Cuenta Suspendida.")
                  Exit Sub
              End If
              '\\\OT-06-0607-OD-FIN///

             If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
                Text2.TEXT = RASOCLI$(NC%)
                Text3.TEXT = DOMICLI$(NC%)
                Text4.TEXT = CPOSCLI$(NC%)
                Text5.TEXT = LOCACLI$(NC%)
                Label26 = CODICLI$(NC%)
                Call NOTACLI(NC%, 1)
                On Error Resume Next
                BOTNEXT.SetFocus
                On Error GoTo 0
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
    DPX1% = XVALO(VALACT1$("ECODEP"))
    If DPX1% > 0 Then
        DP1% = DPX1%
        Text10.TEXT = TRIM$(Str$(DP1%))
        Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    End If
End Sub

Private Sub Text11_GotFocus()
    If AUTONUME% <> 0 Then
      NROPRO& = 0
      Call NURECIBO(NROPRO&)
      Call PRENURE(NROPRO&)
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
    End If
    '
    Text11.SelStart = 3
    On Error Resume Next
    Text11.SelLength = Len(Text11.TEXT) - 3
    On Error GoTo 0
End Sub

Private Sub Text11_LostFocus()
    '
    VDEFF$ = TRIM$(Text11.TEXT)
    Text11.TEXT = VDEFF$
    If Left$(VDEFF$, 3) <> "RB." Then GoTo 90
    NRONUE& = XVALO(Mid$(VDEFF$, 4))
    If NRONUE& > 999999 Then
        Call MENSERR(24, "Número de Recibo Fuera de Rango")
        Call NURECIBO(NROPRO&)
        GoTo 95
    End If
    '
    NROPRO& = NRONUE&
    If NROPRO& > 0 Then
      Call NURECIBO(NROPRO&)
      If NROPRO& <> NRONUE& Then
        Beep
        Call MENSERR(24, "Número no Válido o Pre-existente")
        GoTo 95
      End If
    End If
    Exit Sub
    '
90  Beep
    If FORMALTER.Visible = True Then Exit Sub
    'COMENTE LA SIG. LINEA XQ EL OPRECIB00 NO EXISTE
    'If OPRECIB00.Visible = True Then Exit Sub
    Call MENSERR(24, "Número no Válido")
    NROPRO& = 0
    Call NURECIBO(NROPRO&)
95  Call PRENURE(NROPRO&)
    'Text11.TEXT = "RB." + TRIM$(Str$(NROPRO&))
    On Error Resume Next
    Text11.SelStart = 3
    Text11.SelLength = Len(Text11.TEXT) - 3
    Text11.SetFocus
    On Error GoTo 0
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

Private Sub Text6_DBLCLICK()
   OPRECIB00.Show 1
   'Call CONECRUN("FSETUP04/OPRECIBO", "Configuración de Funcionamiento")
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Text7_Change()
   '
   '\\\OT-07-0029-RG-26/01/07///
   Static OPCIVALO$
   If OPCIVALO$ = "" Then
      OPCIVALO$ = "SI"
      If CONTROL("", "OPCVALOR") = "NO" Then OPCIVALO$ = "NO"
   End If
   '
   If OPCIVALO$ = "NO" Then
      PORDEX = XVALO(Text7): If PORDEX > 99 Then PORDEX = 99
      If PORDEX < 0 Then PORDEX = 0
      IDES# = XVALO(Label10) * PORDEX / 100
      DESCUEN(0) = CSTRING$(MKD$(IDES#), 4, 14, 2, 2)
    Else
      TT77$ = REPLACAR$(TRIM$(Text7), ",", ".")
      PORDEX = XVALO(Text7): If PORDEX > 99 Then PORDEX = 99
      If PORDEX < 0 Then PORDEX = 0
      IDES# = Val(Label14) * (100 / (100 - PORDEX) - 1)
      DESCUEN(0) = CSTRING$(MKD$(IDES#), 4, 14, 2, 2)
      '\\\OT-07-0029-RG-FIN///
      'Call CONDOLA
   End If
   '
   If Text7 <> "" Then
     If Left$(Frame8.Caption, 7) = "Descuen" Then
       Frame8.Caption = "N./Cred. " + PRONOCRE$
     End If
   End If
   
   '
End Sub

Private Sub Text7_GotFocus()
    If TRIM$(Text1) = "" Then Text1.SetFocus: Exit Sub
    NCX1 = XVALO(Text1)
    If NCX1 < 1 Or NCX1 > 32767 Then Text1.SetFocus: Exit Sub
    NCX2% = NCX1
    If ECOARCH$("DEUDOR1", MKI$(NCX2%)) = "" Then Text1.SetFocus
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call Text7_LostFocus
    End If
End Sub

Private Sub Text7_LostFocus()
    Text7 = FORMATNUM(Text7, "F4.1")
End Sub

Private Sub Text8_Change()
   Call Timer1_Timer
End Sub

Private Sub Text8_LostFocus()
    Text8 = FORMATNUM$(XVALO(Text8), "F10.2")
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

Sub NURECIBO(NROPRO&)
    '
    Static FLIMI, FENUE$, AUTONURE%
    '
    Screen.MousePointer = 11
    '
    If FENUE$ = "" Then
      FENUE$ = CONTROL$("", "FENURECI")
      If FENUE$ = "" Then FENUE$ = "01/01/80"
      FLIMI = CVDAT(FECHANUM(FENUE$))
      If FLIMI > Now Then FLIMI = CVDAT(HOY(HO$))
      '
      AUTONURE% = 0
      If CONTROL$("", "NURECIST") = "AUTONUME" Then
        AUTONURE% = 1
      End If
    End If
    '
    COBRACT% = 0
    On Error Resume Next
    COBRACT% = XVALO(Label22.Caption)
    On Error GoTo 0
    '
    NROMA& = 0
    FLIMI1 = CDbl(FLIMI)
    COBRACT% = XVALO(Label22)
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT NumeComp, CodiCom_Lar FROM CajaBanc "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + "AND CodiCom_Cor LIKE 'RB*' "
    SQLX$ = SQLX$ + "AND VAL(MID$(CodiCom_Lar,6,4)) = " & COBRACT% & " "
    SQLX$ = SQLX$ + "AND Status < 9 "
    SQLX$ = SQLX$ + "AND CDBL(FechEmisi) >= " & FLIMI1 & " "
    SQLX$ = SQLX$ + "ORDER BY NumeComp DESC;"
    Set OrpaTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    On Error Resume Next
    With OrpaTemp
      .MoveFirst
      If Err < 1 Then
        NROMA& = !NumeComp
      End If
      '
      .FindFirst ("NumeComp = " & NROPRO& & " AND VAL(MID$(CodiCom_Lar,6,4)) = " & COBRACT%)
      If .NoMatch = False Then
        NROPRO& = 0
      End If
    End With
    On Error GoTo 0
    '
90  If NROPRO& < 1 Then NROPRO& = 1 + NROMA&
    Screen.MousePointer = 1
    '
End Sub
'
Sub SIGUERECI1()
    '
    If MODORECI% = 0 Then
        MODORECI% = 1       ' por emision
        If CONTROL$("RECISTA", "MODORDE") = "VENCIM" Then
           MODORECI% = 2
        End If
    End If
    '
    NCXI = XVALO(Text1.TEXT)
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
    NCX1X% = GRUDIVCLI%(NC%)
    If NCX1X% <> NC% Then
      Call MENSERR(24, "Sub-Cuenta Divisional - Utilice " + TRIM$(Str$(NCX1X%)))
      Text1 = TRIM$(Str$(NCX1X%))
      Text1.SetFocus
      Exit Sub
    End If
    '
    FE$ = TRIM$(Text9.TEXT)
    FF% = FECHANUM(FE$)
    HOI% = HOY(HO$)
    If FF% < 1 Or FF% > HOI% + 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text9.SetFocus
        Exit Sub
    End If
    '
' VENTANA PARA GRUPO BRASIL DEL 8/03/2008 AL 18/03/2008
If HOI% >= 11752 Then
  If HOI% <= 11762 Then
    If FF% > 10000 Then
      If FF% <= HOI% Then
        GoTo 29
      End If
    End If
  End If
End If
    '
    If DIENTRE%(FF%, HOI%) > 30 Then
      A = DIENTRE%(FF%, HOI%)
      EPAC$ = TRIM$(UCase$(EMPREAC$))
      If InStr(EPAC$, "NEUMAN") < 1 Or XVALO(Left$(TRIM$(Label22.Caption), 2)) <> 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text9.SetFocus
        Exit Sub
      End If
    End If
    '
29  Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 11
    NC% = XVALO(Text1.TEXT)
    If NC% > 0 Then
        Text2.TEXT = RASOCLI$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        'Text7.TEXT = TEL1CLI$(NC%)
        'Text10.TEXT = CSTRING$(MKD$(SALDCLI#(NC%)), 4, 13, 2, 2)
        Refresh
    End If
    '
100 List1.Clear
    List1.Refresh
    List2.Clear
    List2.Refresh
    SUMADEU# = 0: SUIMPO# = 0: SUSALDO# = 0
    '
    Z$ = Space$(128)
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM SADEUPEN "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NC% & " "
    SQLX$ = SQLX$ + "ORDER BY FeVencim ASC;"
    Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With FacPenTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          TIX$ = UCase$(Left$(!CODICOM_LAR, 2))
          x$ = !CODICOM_LAR
          SG% = (-1)
          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                              If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                              If TIX$ = "NC" Then TIPOS$ = "Nota de Credito": SG% = 1
                              If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1 '\\\OT-06-0112-WAR-30/03/06///
                              If TIX$ = "RB" Then TIPOS$ = "Anticipo": SG% = 1 '\\\OT-06-0112-WAR-30/03/06///
                              '
          If TIX$ = "FC" Or TIX$ = "ND" Or TIX$ = "AS" Or TIX$ = "RB" Or TIX$ = "NC" Then
            SALDO# = !SaldDebe - !SaldAcre
            If MODOMONEI% = 2 Then
               SALDOLA# = !SaldDebe_Dola - !SaldAcre_Dola
               MONEDOCU% = !Mone_Emis
               If MONEMI% > 1 Then
                 If MONEMI% <> MONEDOCU% Then GoTo 95
                 SALDO# = SALDOLA#
               End If
            End If
            '
'            If SALDO# >= 0.005 Then
            If Abs(SALDO#) >= 0.005 Then
              IMPO# = (-1) * !IBruComp * SG%
              SUMAFAC# = SUMAFAC# + IMPO#
              IMPO# = Abs(IMPO#)
              '
              NROX$ = !CODICOM_LAR
              FEMI = CVINT(!FechEmisi)
              FEVE = CVINT(!FEVENCIM)
              '
              IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 13, 2, 2)
              IVATEX$ = CSTRING$(MKD$(IVAO#), 3, 13, 2, 2)
              SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
              '
                    REGTEX$ = CSTRING$(MKS$(REG&), 3, 8, 0, 2)
                    IRETEX$ = CSTRING$(MKI$(KKI%), 1, 2, 0, 2)
              '
              Z$ = Space$(128)
              Call REPLA(Z$, NROX$, 1, 18)
              Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
              Call REPLA(Z$, IMPTEX$, 32, 13)
              Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
              Call REPLA(Z$, SALTEX$, 56, 14)
              Call REPLA(Z$, REGTEX$, 119, 8)
              Call REPLA(Z$, IRETEX$, 127, 2)
              '
              If MODOMONEI% = 2 Then GoTo 80
              ' ESTO ES PARA PROCESO BIMONETARIO
              MONEMORI% = 0: SALDOLAX# = 0
              RFF$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
              DOCUCABE$ = TRIM$(NROX$)
              For JKLY& = 1 To Len(RFF$) Step 4
                RECABI& = CVS(Mid$(RFF$, JKLY&, 4))
                YCB$ = REGLEIDO$("CARBIMON", RECABI&)
                If TRIM$(Mid$(YCB$, 13, 18)) = DOCUCABE$ Then
                  If TRIM$(Mid$(YCB$, 45, 18)) = DOCUCABE$ Then
                    MONEMORI% = CVI(Mid$(YCB$, 31, 2))
                    If MONEMORI% > ULTREG&("TAMONED") Then
                       MONEMORI% = MONEMORI% Mod 6         ' ENJUAGUE ASEMBLI - GRAMY
                    End If
                  End If
                  SALDOLAX# = SALDOLAX# + CVD(Mid$(YCB$, 97, 8)) - CVD(Mid$(YCB$, 105, 8))
                  Call REPLA(YCB$, MKD$(0), 113, 8)
                End If
              Next JKLY&
              If MONEMORI% > 1 Then
                SALDO1# = SALDOLAX# * TICAMONE(MONEMORI%)
                If SALDO1# < 0 Then SALDO1# = 0
                If SALDO1# > SALDO# Then SALDO1 = SALDO#
                SALTEX$ = CSTRING$(MKD$(SALDO1#), 3, 14, 2, 2)
                Call REPLA(Z$, SALTEX$, 56, 14)
              End If
              GoTo 90
              ' FIN ESTO ES PARA PROCESO BIMONETARIO
              '
80            ' esto es para MULTIMONEDA
'              MONEMORI% = 0: SALDOLAX# = 0: IMPODOL# = 0
'              RFF$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
'              DOCUCABE$ = TRIM$(NROX$)
'              For JKLY& = 1 To Len(RFF$) Step 4
'                RECABI& = CVS(Mid$(RFF$, JKLY&, 4))
'                YCB$ = REGLEIDO$("CARBIMON", RECABI&)
'                If TRIM$(Mid$(YCB$, 13, 18)) = DOCUCABE$ Then
'                  If TRIM$(Mid$(YCB$, 45, 18)) = DOCUCABE$ Then
'                    MONEMORI% = CVI(Mid$(YCB$, 31, 2))
'                    If MONEMORI% > ULTREG&("TAMONED") Then
'                       MONEMORI% = MONEMORI% Mod 6         ' ENJUAGUE ASEMBLI - GRAMY
'                    End If
'                    IMPODOL# = Abs(CVD(Mid$(YCB$, 97, 8)) - CVD(Mid$(YCB$, 105, 8)))
'                  End If
'                  SALDOLAX# = SALDOLAX# + CVD(Mid$(YCB$, 97, 8)) - CVD(Mid$(YCB$, 105, 8))
'                  Call REPLA(YCB$, MKD$(0), 113, 8)
'                End If
'              Next JKLY&
'              If MONEMORI% <> MONEMI% Then GoTo 95
'              SALDO1# = SALDOLAX#
'              If SALDO1# < 0 Then SALDO1# = 0
'              SALTEX$ = CSTRING$(MKD$(SALDO1#), 3, 14, 2, 2)
'              Call REPLA(Z$, SALTEX$, 56, 14)
'              IMPTEX$ = CSTRING$(MKD$(IMPODOL#), 3, 13, 2, 2)
'              Call REPLA(Z$, IMPTEX$, 32, 13)
'              ' FIN PROCESO MULTIMONEDA
              
90            SUMADEU# = SUMADEU# + XVALO(SALTEX$)
              SUIMPO# = SUIMPO# + XVALO(IMPTEX$)
              '
              List2.AddItem Z$
              '
            End If
          End If
95      .MoveNext
        Loop
      End If
    End With
    '
200 List2.ListIndex = -1
    List2.Refresh
    Label4.Caption = CSTRING$(MKD$(SUIMPO#), 4, 14, 2, 2)
    Label4.Refresh
    Label8.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    Label8.Refresh
    Label10.Caption = ""
    Label10.Refresh
    List2.Visible = True
    List2.SetFocus
    If List2.ListCount > 0 Then
        VACON% = 1
        List2.ListIndex = 0
        VACON% = 0
    End If
    '
    BOTNEXT.Enabled = False
    Call BLANARCH("!CARMECO")
    Screen.MousePointer = 1
    '
    Call SIGUERECI2
    '
End Sub
'

Sub SIGUERECI2()
    '
    HOI% = HOY(HO$)
    '\\\OT-08-0203-OD-25/07/08///
    OPX% = COMALTER%("Opciones de Carga de Cheques:\\Lectora de Cheques\Carga Manual")

    If OPX% < 1 Or OPX% > 2 Then Exit Sub
    '
    If OPX% = 2 Then
       GoTo 200
      Else
       GoTo 100
    End If
100 OBX$ = OBJPLA$("LECTORCHEQUE", VDEF$)
    If OBX$ = "" Then GoTo 200
    '
    'Valida  que el String pasado por la lectora se haya cargado
    DatoCheque$ = TRIM$(Mid$(OBX$, 2, 29))
    If DatoCheque$ = "" Then
        Call MENSERR(24, "Faltan Datos del Cheque")
        VDEF$ = OBX$
        GoTo 100
    End If
    '
    'Valida que el String pasado por la lectora tenga  29 Caracteres
    If Len(DatoCheque$) <> 29 Then
      Call MENSERR(24, "Lectura Mal Realizada")
      VDEF$ = OBX$
      GoTo 100
    End If
    '
    'Valida que tenga el importe
    ImporCheque# = CVS(Mid$(OBX$, 41, 4))
    If ImporCheque# < 0.0005 Then
        Call MENSERR(24, "Falta el Importe del Cheque")
        VDEF$ = OBX$
        GoTo 100
    End If
    '
    'Valida que el medio de pago sea valido
    '
     NumeroBanco% = XVALO(Mid$(DatoCheque$, 1, 3))
    '
    EcoMedioPago$ = TRIM$(ECOARCH$("MECOBRA", MKI$(NumeroBanco%)))
    If EcoMedioPago$ <> "" Then
      Else
       Call MENSERR(24, "Medio de Pago Inexistente")
       VDEF$ = OBX$
       GoTo 100
    End If
    '
    'Valida que tenga fechas
    '
    FechaEmision% = CVI(Mid$(OBX$, 45, 2))
    FechaVencim% = CVI(Mid$(OBX$, 47, 2))
    If FechaEmision% < 1 Or FechaVencim% < 1 Then
       Call MENSERR(24, "Falta completar Fechas")
       VDEF$ = OBX$
       GoTo 100
    End If
    '
    'Validacion que fecha de vencimiento no sea menor que fecha de emision
    '
    If FechaEmision% > FechaVencim% Then
       Call MENSERR(24, "Fecha de Emisión mayor a Fecha de Pago")
       VDEF$ = OBX$
       GoTo 100
    End If
    '
    'Verifica si el cheque ya fue ingresado anteriormente
    '
    NumeroCheque$ = TRIM$(Mid$(DatoCheque$, 11, 8))
    BANCO1$ = Mid$(EcoMedioPago$, 4, 18)
    VeriChek% = YAESTACHEQUE%(BANCO1$, NumeroCheque$)
    If VeriChek% > 0 Then
      Call MENSERR(24, "Cheque N.: " + NumeroCheque$ + "  Banco : " + EcoMedioPago$ + " Ya Ingresado")
      VDEF$ = OBX$
      GoTo 100
    End If
    '
    'Verifica si el cheque esta repetido
    '
    For i& = 1 To ULTREG("!CARMECO")
      x$ = REGLEIDO$("!CARMECO", i&)
      NuBanco% = TRIM$(CVI(Mid$(x$, 1, 2)))
      NuCheque$ = TRIM$(Mid$(x$, 19, 12))
      If NumeroBanco% = NuBanco% And NumeroCheque$ = NuCheque$ Then
        Call MENSERR(24, "Cheque N.: " + NumeroCheque$ + "  Banco : " + EcoMedioPago$ + " Repetido")
        VDEF$ = OBX$
        GoTo 100
      End If
    Next i&
    '
    'VALIDACION DE EXISTENCIA DE CUIT
    '
    TX$ = ""
    CuitLibrador$ = TRIM$(Mid$(OBX$, 49, 14))
    If CuitLibrador$ = "" Then
      OPXX% = COMALTER%("Falta Ingresar C.U.I.T:\\Agregar \Continuar")
      If OPXX% = 1 Then
        VDEF$ = OBX$
        GoTo 100
      End If
       Else
        If VALICUIT%(CuitLibrador$) <= 0 Then
           Call MENSERR(24, "C.U.I.T. no Válido")
           VDEF$ = OBX$
           GoTo 100
          Else
           'EN CASO DE INGRESAR CUIT SE CONSIDERA DE TERCEROS
           'Y SE PONE LA LETRA T EN EL CAMPO DE PROPIOS O TERCEROS
           Call REPLA(TX$, "T", 43, 2)
        End If
    End If
    '
    'Graba la informacion ingresada  por Lectora de Cheques
    '
    Call REPLA(TX$, MKI(NumeroBanco%), 1, 2)
    Call REPLA(TX$, NumeroCheque$, 19, 12)
    Call REPLA(TX$, MKD(ImporCheque#), 31, 8)
    Call REPLA(TX$, MKI$(FechaEmision%), 39, 2)
    Call REPLA(TX$, MKI$(FechaVencim%), 41, 2)
    Call REPLA(TX$, MKI$(48), 45, 2)
    Call REPLA(TX$, CuitLibrador$, 47, 16)
    Call REGAPP("!CARMECO", TX$)
    TX$ = ""
    '
    VDEF$ = ""
    GoTo 100
    '\\\OT-08-0203-OD-Fin///
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Ingreso de Medios de Pago Codificados"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top + 100))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
    '
    Timer1.Enabled = True
    VTB% = VENTABU%("CARMECOW" + ATRI$)
    Timer1.Enabled = False
    Call CIERRARCH("!CARMECO")
    AAA = ULTREG&("!CARMECO")
    Call Timer1_Timer
    If VTB% < 0 Then
        BOTNEXT.Enabled = True
        Exit Sub
    End If
    '
    ATXY% = 0
    For U& = 1 To ULTREG&("!CARMECO")
      x$ = REGLEIDO$("!CARMECO", U&)
      PROTER$ = UCase$(TRIM$(Mid$(x$, 43, 2)))
      If PROTER$ = "" Or Left$(PROTER$, 1) = "P" Then
           PROTER$ = "  "
         Else
           PROTER$ = "3 "
      End If
      If TRIM$(PROTER$) <> "" Then ATXY% = 1
    Next U&
    '
    If ATXY% > 0 Then
      ATRI$ = "/NC"
      ATRI$ = ATRI$ + "/TITUPAN=Ingreso CUIT Cheques de Terceros"
      ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture1.Top + 100))
      ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture1.Left))
      '
      VTB% = VENTABU%("CARMECOX" + ATRI$)
      Call CIERRARCH("!CARMECO")
      AAA = ULTREG&("!CARMECO")
      If VTB% < 0 Then
          BOTNEXT.Enabled = True
          Exit Sub
      End If
    End If
    '
    List1.Clear
    SUMAVA# = 0
    PROVALO# = 0
    TOTVALO# = 0
    Label14 = "": Label30 = "": Text8 = ""
    For U& = 1 To ULTREG&("!CARMECO")
        x$ = REGLEIDO$("!CARMECO", U&)
        If x$ = REBLA$ Then GoTo 290
        MECO% = CVI(Left$(x$, 2))
        BANCO$ = TRIM$(ECOARCH$("MECOBRA", MKI$(MECO%)))
        '
        If BANCO$ = "" Then
            Call MENSERR(24, "Medio de Pago no Válido")
            GoTo 200
        End If
        '
        NUVA$ = Mid$(x$, 19, 12)
        '
        'VERIFICA QUE SEA CHEQUE
        CUECO% = CUEMECOB%(MECO%)
        If CUECO% <> CUEICARCH% Then GoTo 19
        
        'VERIFICA SI EL CHEQUE  FUE INGRESADO
        BANCO1$ = Mid$(BANCO$, 4, 18)
        VeriChek% = YAESTACHEQUE%(BANCO1$, NUVA$)
        If VeriChek% > 0 Then
          Call MENSERR(24, "Cheque N.: " + NUVA$ + "  Banco : " + BANCO1$ + " Ya Ingresado")
          GoTo 200
        End If
        '
        'Verifica si el cheque esta repetido
        For i& = 1 To (U& - 1)
          X1$ = REGLEIDO$("!CARMECO", i&)
          If Mid$(X1$, 1, 32) = Mid$(x$, 1, 32) Then
            Call MENSERR(24, "Cheque N.: " + NUVA$ + "  Banco : " + BANCO1$ + " Repetido")
            GoTo 200
          End If
        Next i&

19      IMVA# = CVD(Mid$(x$, 31, 8))
        IMVABS# = IMVA#: If UCase$(BANCO$) = "EFECTIVO" Then IMVABS# = Abs(IMVA#)
        ' LO ANTERIOR ES PARA ACEPTAR EFECTIVO EN NEGATIVO
        If IMVABS# < 0.005 Then
            Call MENSERR(24, "Falta Importe")
            GoTo 200
        End If
        '
        If UCase$(Left$(BANCO$, 2)) = "CH" Then
          FEMI = CVI(Mid$(x$, 39, 2))
          FEVE = CVI(Mid$(x$, 41, 2))
          '
          ' ESTO ES PARA EL CASO DE QUE NO SE INGRESA FECHA DE EMISIÓN Y
          ' SI DE VENCIMIENTO - PEDIDO POR ANTONIO DE TANIT EL 16/05/2007
          If Abs(FEMI) < 0.5 And FEVE > 0.5 Then
            FEMII% = FEVE
            If FEMII% > HOI% Then FEMII% = HOI%
            Call REPLA(x$, MKI$(FEMII%), 39, 2)
            Call GRAREG("!CARMECO", x$, U&)
            FEMI = FEMII%
          End If
          '
          If Abs(FEMI) < 0.5 Then
            Call REPLA(x$, MKI$(HOI%), 39, 2)
            Call GRAREG("!CARMECO", x$, U&)
            FEMI = HOI%
          End If
          F1 = FEMI: F2 = HOI%
' VENTANA GRUPO BRASIL
If HOI% >= 11752 Then
  If HOI% <= 11762 Then
    If F1 > 10000 Then
      If F1 <= HOI% Then
        GoTo 29
      End If
    End If
  End If
End If
          If FEMI < HOI% - 32 * 13 Or FEMI > HOI% + 30 Or DIENTRE%(F1, F2) > 365 Then
            Call MENSERR(24, "Fecha de Emisión no Válida")
            GoTo 200
          End If
          '
29        FEVE = CVI(Mid$(x$, 41, 2))
          If Abs(FEVE) < 0.5 Then
            Call REPLA(x$, MKI$(HOI%), 41, 2)
            Call GRAREG("!CARMECO", x$, U&)
            FEVE = HOI%
          End If
          F1 = FEVE: F2 = HOI%
' VENTANA GRUPO BRASIL
If HOI% >= 11752 Then
  If HOI% <= 11762 Then
    If F1 > 10000 Then
      If F1 <= HOI% Then
        GoTo 39
      End If
    End If
  End If
End If
          If FEVE < HOI% - 64 Or DIENTRE%(F1, F2) > 30 Or FEVE < FEMI - 0.001 Then
            Call MENSERR(24, "Fecha de Vencimiento no Válida")
            GoTo 200
          End If
        End If
        '
39      PROTER$ = UCase$(TRIM$(Mid$(x$, 43, 2)))
        If PROTER$ = "" Or Left$(PROTER$, 1) = "P" Then
             PROTER$ = "  "
           Else
             PROTER$ = "3 "
        End If
        Call REPLA(x$, PROTER$, 43, 2)
        Call GRAREG("!CARMECO", x$, U&)
        '
        CLEA% = CVI(Mid$(x$, 45, 2))
        If Mid$(x$, 45, 2) = "  " Then
          Call REPLA(x$, MKI$(24), 45, 2)
          Call GRAREG("!CARMECO", x$, U&)
          CLEA% = 24
        End If
        If CLEA% < 24 Then CLEA% = 24
        If (CLEA% Mod 24) <> 0 Then
            Call MENSERR(24, "Clearing no Válido")
            GoTo 200
        End If
        Call REPLA(x$, MKI$(CLEA%), 45, 2)
        Call GRAREG("!CARMECO", x$, U&)
        '
        Z$ = Space$(96)
        SUMAVA# = SUMAVA# + IMVA#
        IMPTEX$ = CSTRING$(MKD$(IMVA#), 3, 14, 2, 2)
        Call REPLA(Z$, CSTRING$(MKI$(MECO%), 1, 4, 0, 2), 1, 4)
        Call REPLA(Z$, BANCO$, 7, 14)
        Call REPLA(Z$, NUVA$, 23, 12)
        Call REPLA(Z$, PROTER$, 38, 2)
        Call REPLA(Z$, CSTRING$(MKI$(CLEA%), 1, 4, 0, 2), 42, 4)
        Call REPLA(Z$, FECHATEX$(FEMI), 50, 8)
        Call REPLA(Z$, FECHATEX$(FEVE), 62, 8)
        Call REPLA(Z$, IMPTEX$, 70, 14)
        '\\\OT-08-0203-OD-29/04/08///
        'CALCULO PARA PROMEDIO DE VALORES
        PROVALO# = PROVALO# + (DIENTRE(HOY(HO$), FEVE) * XVALO(IMPTEX$))
        TOTVALO# = TOTVALO# + XVALO(IMPTEX$)
        '
        List1.AddItem Z$
        '
290 Next U&
    
    If Abs(TOTVALO#) > 0.001 Then
      Label19 = FORMATNUM(PROVALO# / TOTVALO#, "F4.0")
     Else
      Label19 = 0
    End If
    '
      Label23 = FORMATNUM$(XVALO(Label17) + XVALO(Label19), "F4.0")
    '\\\OT-08-0203-OD-FIN///
    '
    If List1.ListCount > 0 Then
        List1.ListIndex = 0
    End If
    '
    Label30.Caption = TRIM$(CSTRING$(MKD$(SUMAVA#), 4, 14, 2, 2))
    Label30.Refresh
    Label14.Caption = TRIM$(CSTRING$(MKD$(SUMAVA# - XVALO(Text8)), 4, 14, 2, 2))
    Label14.Refresh
    '\\\OT-07-0029-RG-26/01/07///
    ANTI# = XVALO(Label14) + XVALO(DESCUEN(0)) - CDbl(XVALO(Label18.Caption))
    If ANTI# < 0 Then
          Frame10.Caption = "Descuentos:"
        Else
          Frame10.Caption = "Anticipos:"
    End If
    '\\\OT-07-0029-RG-FIN///
    Label20.Caption = TRIM$(CSTRING$(MKD$(ANTI#), 4, 14, 2, 2))
    Label20.Refresh
    '
    If Abs(XVALO(Label29)) < 1 Then
       On Error Resume Next
       Text8.SetFocus
       On Error GoTo 0
    End If
End Sub
'

Sub SIGUERECI3()
    '
    '
    If List2.ListCount < 1 Then
        Exit Sub
    End If
    '
    Call BLANARCH("!CARGAPLI")
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        Z$ = List2.TEXT
        NROX$ = Mid$(Z$, 1, 20)
        FEMII% = FECHANUM(Mid$(Z$, 23, 8))
        IPED# = CDbl(XVALO(Mid$(Z$, 32, 13)))
        FEVEI% = FECHANUM(Mid$(Z$, 47, 8))
        SALD# = CDbl(XVALO(Mid$(Z$, 56, 14)))
        ACRE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
        DOCU$ = Mid$(Z$, 1, 18)
        REG2& = XVALO(Mid$(Z$, 119, 8))
        IRE% = XVALO(Mid$(Z$, 127, 2))
        '
        Y$ = REGBLAN$("!CARGAPLI")
        Call REPLA(Y$, NROX$, 1, 20)
        Call REPLA(Y$, MKI$(FEMII%), 21, 2)
        Call REPLA(Y$, MKD$(IPED#), 23, 8)
        Call REPLA(Y$, MKI$(FEVEI%), 31, 2)
        Call REPLA(Y$, MKD$(SALD#), 33, 8)
        Call REPLA(Y$, MKD$(ACRE#), 41, 8)
        Call REPLA(Y$, DOCU$, 49, 18)
        Call REPLA(Y$, MKS$(REG2&), 67, 4)
        Call REPLA(Y$, MKI$(IRE%), 71, 2)
        '
        Call REGAPP("!CARGAPLI", Y$)
    Next U&
    Call CIERRARCH("!CARGAPLI")
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Aplicación de Cobranza"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture2.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture2.Left))
    '
    VTB% = VENTABU%("CARGAPLI" + ATRI$)
    '
    VACON% = 1
    If VTB% >= 0 Then
        List2.Clear
        For U& = 1 To ULTREG&("!CARGAPLI")
            Y$ = REGLEIDO$("!CARGAPLI", U&)
            NROX$ = Mid$(Y$, 1, 20)
            If TRIM$(NROX$) = "Anticipo no Aplicado" Then
               HOI% = HOY(HO$)
               ACRE# = CVD(Mid$(Y$, 41, 8))
               Call REPLA(Y$, MKI$(HOI%), 31, 2)
               Call REPLA(Y$, MKD$(ACRE#), 23, 8)
               Call REPLA(Y$, MKD$(ACRE#), 33, 8)
               Call GRAREG("!CARGAPLI", Y$, U&)
            End If
            FEMI = CVI(Mid$(Y$, 21, 2))
            IPED# = CVD(Mid$(Y$, 23, 8))
            IMPTEX$ = CSTRING$(MKD$(IPED#), 4, 13, 2, 2)
            FEVE = CVI(Mid$(Y$, 31, 2))
            SALD# = CVD(Mid$(Y$, 33, 8))
            SALTEX$ = CSTRING$(MKD$(SALD#), 4, 14, 2, 2)
            ACRE# = CVD(Mid$(Y$, 41, 8))
            '
            'If ACRE# < 0 Then
            If Abs(ACRE#) >= 0.005 And Sgn(ACRE#) <> Sgn(SALD#) Then
                Call MENSERR(24, "No se Permite Aplicación Negativa")
                GoTo 200
            End If
            If U& > 1 Then
              If Abs(ACRE#) > Abs(SALD#) Then
                Call MENSERR(24, "El Importe Acreditado Debe Ser\Inferior al Saldo del Documento")
                 GoTo 200
              End If
            End If
            ACRTEX$ = CSTRING$(MKD$(ACRE#), 4, 14, 2, 2)
            DOCU$ = Mid$(Y$, 49, 18)
            REG2& = CVS(Mid$(Y$, 67, 4))
            IRE% = CVI(Mid$(Y$, 71, 2))
            REGTEX$ = CSTRING$(MKS$(REG2&), 3, 8, 0, 2)
            IRETEX$ = CSTRING$(MKI$(IRE%), 1, 2, 0, 2)
            '
            Z$ = Space$(128)
            Call REPLA(Z$, NROX$, 1, 20)
            Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
            Call REPLA(Z$, IMPTEX$, 32, 13)
            Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
            Call REPLA(Z$, SALTEX$, 56, 14)
            Call REPLA(Z$, ACRTEX$, 70, 14)
            Call REPLA(Z$, REGTEX$, 119, 8)
            Call REPLA(Z$, IRETEX$, 127, 2)
            List2.AddItem Z$
            '
        Next U&
    End If
    '
    SUMACRE# = 0
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        Z$ = List2.TEXT
        SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(Z$, 70, 14)))
    Next U&
    '
    Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label10.Refresh
    Label18.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    Label18.Refresh
    '\\\OT-07-0029-RG-26/01/07///
    Label20.Caption = CSTRING$(MKD$(XVALO(Label14.Caption) + XVALO(DESCUEN(0)) - XVALO(Label18.Caption)), 4, 14, 2, 2)
    '\\\OT-07-0029-RG-FIN///
    Label20.Refresh
    '
    List2.Refresh
    VACON% = 1
    '
End Sub
'

Sub GRARECISTA(OKEY%)
    '
700 OKEY% = 0
    'APRORECIB.Show 1
    'If APRORECIB.Label19 <> "OK" Then Exit Sub
    Screen.MousePointer = 11   ' Grabando Recibo - Espere un Momento, por Favor"
    '
    COTIMONEM# = XVALO(TICAMBIS)
    NCLIEN% = XVALO(APRORECIB.Label1)
       NROO& = XVALO(Mid$(TRIM$(Text11.TEXT), 4))
    FC% = FECHANUM(TRIM$(APRORECIB.Label6))
    TOTAVAL# = ROUND#(CDbl(XVALO(APRORECIB.Label12)) * COTIMONEM#)
    ANTI# = ROUND#(CDbl(XVALO(APRORECIB.Label14)) * COTIMONEM#)
    COBRACT% = XVALO(APRORECIB.Label3)
    CUECARTE% = XVALO(APRORECIB.Label16)
    '
    RECICORTO$ = TRIM$(APRORECIB.Label8)
    RECILARGO$ = TRIM$(APRORECIB.Label10)
    '
    CoComCor$ = RECICORTO$
    CoComLar$ = RECILARGO$
    Referen$ = RASOCLI$(NCLIEN%)
    FeContab = CVDAT(FC%)
    FeAcredi = CVDAT(FC%)
    '
840 ' ***** aqui grabar cheques en CAJABANC.DAT
    NOCLOSE% = 1
    SUDEBE# = 0
    For U& = 1 To ULTREG&("!CARMECO")
       x$ = REGLEIDO$("!CARMECO", U&)
       TIPO% = CVI(Left$(x$, 2))
       BANCO$ = TRIM$(Mid$(x$, 3, 16))
       If BANCO$ = "" Then
         BANCO$ = TRIM$(ECOARCH$("MECOBRA", MKI$(TIPO%)))
       End If
       If Left$(UCase$(BANCO$), 3) = "CH." Then
         If Len(BANCO$) > 3 Then
           BANCO$ = Mid$(BANCO$, 4)
         End If
       End If
       NUCHE$ = TRIM$(Mid$(x$, 19, 12))
       IMCHE# = ROUND(COTIMONEM# * CVD(Mid$(x$, 31, 8)))
       FECHE% = CVI(Mid$(x$, 39, 2))
       FEVEN% = CVI(Mid$(x$, 41, 2))
       CUITLIB$ = TRIM$(Mid$(x$, 47, 16))
       '
       If Abs(IMCHE#) > 0 Then
          SUDEBE# = SUDEBE# + IMCHE#
             FEX = FEVEN%: FEVECHE$ = FECHATEX$(FEX)
          FeContab = CVDAT(FC%)
             CUECO% = CUEMECOB%(TIPO%)
          CuentaCo% = CUECO%
          NumeClie% = NCLIEN%
          NumeProv% = 0
          TiMovi$ = "CHRECI"
          OpMovi$ = "CH"
          VaMovi$ = ""
             IDEVA$ = AJUSTI$(BANCO$, 18)
             IDEVA$ = IDEVA$ + AJUSTI$(NUCHE$, 14)
             IDEVA$ = IDEVA$ + FEVECHE$
          Descrip$ = IDEVA$
             If FECHE% < 1 Then FECHE% = FC%
          FechEmi = CVDAT(FECHE%)
          FeAcredi = CVDAT(FEVEN%)
          NumeTal = COBRACT%
          CodiEmp% = 0
          CodiSuc% = SUCUI%
          NumeCom = NROO&
             STATUI% = 0
             If CUECO% = CUECARTE% Then STATUI% = 1
          StatInit = STATUI%
          ImpoDeb# = IMCHE#
          Observa$ = ""
            If CUITLIB$ <> "" Then
              Observa$ = "CUIT Libr: " + CUITLIB$
            End If
          Call GRACABAN
       End If
    Next U&
    OKEY% = 6
    ' fin graba de cheques en CAJABANC **********************************
    '
    ' AQUI GRABA AJUSTE DE REDONDEO
    IREDON# = XVALO(Text8)
    If Abs(IREDON#) >= 0.005 Then
       CUEREDO$ = TRIM$(CONTROL$("RECISTA", "CUEREDON"))
       If CUEREDO$ <> "" Then CUERED% = CUEINT%(CUEREDO$)
       '
       SUDEBE# = SUDEBE# + IREDON#
          FEX = FC%: FEVECHE$ = FECHATEX$(FEX)
       FeContab = CVDAT(FC%)
       CuentaCo% = CUERED%
       NumeClie% = NCLIEN%
       NumeProv% = 0
       TiMovi$ = "CHRECI"
       OpMovi$ = "REDON"
       VaMovi$ = ""
       Descrip$ = "Ajuste Redondeo"
       FechEmi = CVDAT(FC%)
       FeAcredi = CVDAT(FC%)
       NumeTal = COBRACT%
       CodiEmp% = 0
       CodiSuc% = SUCUI%
       NumeCom = NROO&
       StatInit = 0
       ImpoDeb# = IREDON#
       Observa$ = ""
       Call GRACABAN
    End If
    '
843 ' acumula anticipo entre detalle y casillero
    For U& = 1 To List2.ListCount
      Z$ = List2.LIST(U& - 1)
      DOCU$ = Left$(Z$, 20)
      If DOCU$ = "Anticipo no Aplicado" Then
        ANTI# = ANTI# + ROUND#(CDbl(XVALO(Mid$(Z$, 70, 14))) * COTIMONEM#)
        List2.RemoveItem (U& - 1)
        GoTo 843
      End If
    Next U&
    '
    ' graba el anticipo como deuda pendiente
    If ANTI# >= 0.005 Then
      IBruTot# = (-1) * ANTI#
      IIvaTot# = 0
      NuorVen% = 0
      FeVenc = CVDAT(FC%)
      ImVenc# = (-1) * ANTI#
      Call GRADEUPEN
    End If
    OKEY% = 1
    '
800 ' *********************************** aqui blanquear facturas canceladas
    CoComCor$ = RECICORTO$
    CoComLar$ = RECILARGO$
    CodiEmp% = 0
    CodiSuc% = SUC%
    NumeClie% = NCLIEN%
    NumeProv% = 0
    FechEmi = CVDAT(FC%)
    RefApli$ = "Cobranza " + RECICORTO$
    '
    XX% = 0
    VACON% = 1
    SUCANCE# = 0: ENCUENDE% = 0
    For U& = 1 To List2.ListCount
       Z$ = List2.LIST(U& - 1)
       DOCU$ = Left$(Z$, 20)
       If DOCU$ <> "Anticipo no Aplicado" Then
          CANCE# = ROUND#(COTIMONEM# * CDbl(XVALO(Mid$(Z$, 70, 14))))
          If Abs(CANCE#) > 0 Then
             DOCUCA$ = TRIM$(Left$(Z$, 18))
             ICAN# = CANCE#
             SICAN% = Sgn(CANCE#)
             TOTAPLI# = 0
             '
             DocOrCor$ = ""
             TotOrig# = 0
             IvaOrig# = 0
             '
             Call APERBASDAT("CABAN") 'WAR MODIFICADO X CT-101463-FG
             '
             With SaDeuPen
                .FindFirst "NuClien = " & NCLIEN% & " AND CODICOM_LAR = '" & DOCUCA$ & "'" ' AND SALDDEBE > 0"
850             If .NoMatch = False Then
                   ENCUENDE% = 1
                   PARCAN# = ICAN#
                   If SICAN% * (!SaldDebe - !SaldAcre) < SICAN% * ICAN# Then PARCAN# = !SaldDebe - !SaldAcre
                   .Edit
                   If PARCAN# > 0 Then
                        !SaldDebe = ROUND#(!SaldDebe) - PARCAN#
                      Else
                        !SaldAcre = ROUND#(!SaldAcre) + PARCAN#
                        If !Mone_Emis = 2 Then ' ESTO SE AGREGA PARA ACTUALIZAR EL SALDO EN DOLAR DE LOS CREDITOS
                            IMPDOLA# = ROUND#(CDbl(XVALO(Mid$(Z$, 70, 14))))
                            !SaldAcre_Dola = !SaldAcre_Dola + IMPDOLA#
                        End If
                   End If
                   .Update
                   TOTAPLI# = TOTAPLI# + PARCAN#   ' SUMA APLICACIONES -  agregado por tema bimonetario 22/03
                   '
                   DocOrCor$ = !CodiCom_Cor
                   TotOrig# = SICAN% * XVALO(!IBruComp)
                   IIVAORIG# = 0
                   On Error Resume Next
                   IvaOrig# = !IIvaComp
                   On Error GoTo 0
                   '
                   ICAN# = ICAN# - PARCAN#
                   If Abs(ICAN#) >= 0.005 Then
                     .FindNext "Nuclien = " & NumeClie% & " AND CODICOM_LAR = '" & DOCUCA$ & "'" ' AND SALDDEBE > 0"
                     GoTo 850
                   End If
                End If
             End With
             '
             DocOrLar$ = TRIM$(Left$(Z$, 18))
             ImpApli# = TOTAPLI#
             SUCANCE# = SUCANCE# + TOTAPLI#
             If XVALO(DESCUEN(0)) < 0.005 Then
                  Call GRAPLICOB
               Else
                  ImpApli# = ROUND#(TOTAPLI# * TOTAVAL# / (TOTAVAL# + XVALO(DESCUEN(0))), 2)
                  Call GRAPLICOB
                  ImpApli# = TOTAPLI# - ImpApli#
                  CoComCor$ = APRORECIB.Label31
                  CoComLar$ = APRORECIB.Label29
                  Call GRAPLICOB
                  ' LO QUE SIQUE ES PARA RESTAURAR LOS VALORES ORIGINALES
                  CoComCor$ = RECICORTO$
                  CoComLar$ = RECILARGO$
             End If
             '
             If Abs(TOTAPLI#) < Abs(CANCE#) - 0.005 Then
                 ' acumula la diferencia para ajuste de saldo
                 Z1$ = Left$(Z$ + Space$(256), 256)
                 DIFAPLI$ = FORMATNUM$(CANCE# - TOTAPLI#, "F12.2")
                 Call REPLA(Z1$, DIFAPLI$, 241, 12)
                 List2.RemoveItem U& - 1
                 List2.AddItem Z1$, U& - 1
             Else
             
             End If
          End If
       End If
    Next U&
    VACON% = 0
    '
    ' aqui graba el importe bruto del recibo
    FeContab = CVDAT(FC%)
      CUECO% = CUECOCLI%(NCLIEN%)
    CuentaCo% = CUECO%
    NumeClie% = NCLIEN%
    NumeProv% = 0
    TiMovi$ = "RCOB"
    OpMovi$ = "CCORR"
    VaMovi$ = ""
    Descrip$ = "Cobranza Cta.Corriente " + RECICORTO$
    FechEmi = CVDAT(FC%)
    FeAcredi = CVDAT(FC%)
    NumeTal = COBRACT%
    CodiEmp% = 0
    CodiSuc% = SUCUI%
    NumeCom = NROO&
    StatInit = 0
    ImpoDeb# = (-1) * SUDEBE#
    AtraCheque& = XVALO(Label17)
    AtraCance& = XVALO(Label19)
    Call GRACABAN
    OKEY% = 4
    '
    ' aqui graba la nota de credito por descuento
    DESCAJ# = XVALO(DESCUEN(0))
    If DESCAJ# >= 0.005 Then
       GoSub GRANOCRED
    End If
    '
    DIFECAN# = SUDEBE# - SUCANCE# - ANTI#
    If DIFECAN# >= 0.005 Then
       If MODOMONEI% = 2 Then
           IAJU% = CUEINT%(CONTROL$("SVD451S", "CUEVE-9"))
          Else
           IAJU% = CUEINT%(CONTROL$("SVD451S", "CUEVE-11"))
       End If
       If IAJU% < 1 Then
          Call MENSERR(24, "Errores al Grabar este Recibo.\Debe Anularlo y Cargarlo de Nuevo.")
          GoTo 890
       End If
       '
       If MODOMONEI% <> 2 Then
          TTXX$ = "Ocurrio Un Error al Grabar este Recibo.\  \"
          TTXX$ = TTXX$ + "Comunique los Siguientes Valores\"
          TTXX$ = TTXX$ + "a su Soporte de Sistemas:\   \"
          TTXX$ = TTXX$ + "Sudebe = " + TRIM$(FORMATNUM$(SUDEBE#, "F10.4")) + "\"
          TTXX$ = TTXX$ + "Sucance = " + TRIM$(FORMATNUM$(SUCANCE#, "F10.4")) + "\"
          TTXX$ = TTXX$ + "Anti = " + TRIM$(FORMATNUM$(ANTI#, "F10.4")) + "\"
          TTXX$ = TTXX$ + "CotiMonem = " + TRIM$(FORMATNUM$(COTIMONEM#, "F10.4")) + "\"
          TTXX$ = TTXX$ + "Encuende = " + TRIM$(FORMATNUM$(ENCUENDE%, "F10.4")) + "."
          Call MENSERR(24, TTXX$)
          GoTo 49
       End If
       '
       MONEMANT% = MONEMI%
       MONEMI% = 1
       ' graba ajuste diferencia de cambio
       CoComCor$ = "AS" + Mid$(RECICORTO$, 3)
       CoComLar$ = "AS" + Mid$(RECILARGO$, 3)
       FeContab = CVDAT(FC%)
       For U5& = 1 To List2.ListCount
         DIFE1# = XVALO(Mid$(List2.LIST(U5& - 1), 241, 12))
         If Abs(DIFE1#) >= 0.005 Then
               CUECO% = CUECOCLI%(NCLIEN%)
             CuentaCo% = CUECO%
             NumeClie% = NCLIEN%
             NumeProv% = 0
             TiMovi = "AJUD"            ' AJUSTE DE SALDO DEUDORES
             OpMovi = "AS"              ' AJUSTE DE SALDO
             VaMovi = "CCORR"           ' CUENTA CORRIENTE
             Descrip = "Aj.Saldo - " + RASOCLI$(NCLIEN%)
             FechEmi = CVDAT(FC%)
             FeAcredi = CVDAT(FC%)
             NumeTal = COBRACT%
             CodiEmp% = 0
             CodiSuc% = SUCUI%
             NumeCom = NROO&
             StatInit = 0
             ImpoDeb# = DIFE1#
             Call GRACABAN
             '
             ' graba contrapartida del ajuste
             CuentaCo% = IAJU%
             TiMovi = "AJUSALD"         ' AJUSTE DE SALDO DEUDORES
             OpMovi = "AS"              ' AJUSTE DE SALDO
             VaMovi = "GRA-RC"          ' NO IMPUTADO
             Descrip = "Aj.Saldo - " + RASOCLI$(NCLIEN%)
             ImpoDeb# = (-1) * DIFE1#
             Call GRACABAN
         End If
       Next U5&
       '
       MONEMI% = MONEMANT%
       '
    End If
    '
49  OKEY% = 5
    '
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ <> "BIMONETA" Then
       If MODOMONE$ <> "MULTIMON" Then
          GoTo 880
       End If
    End If
    '
    RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCLIEN%))
    If Len(RFF$) > 0 Then
      For U& = 1 To List2.ListCount
        x$ = List2.LIST(U& - 1)
        CANCE# = CDbl(XVALO(Mid$(x$, 70, 14)))
        If CANCE# > 0.005 Then
          ABUSQUE$ = TRIM$(Left$(x$, 18))
          For JKLX& = 1 To Len(RFF$) Step 4
            RECABI& = CVS(Mid$(RFF$, JKLX&, 4))
            YCB$ = REGLEIDO$("CARBIMON", RECABI&)
            If TRIM$(Mid$(YCB$, 13, 18)) = ABUSQUE$ Or TRIM$(Mid$(YCB$, 45, 18)) = ABUSQUE$ Then
              DOCUCABE$ = TRIM$(Mid$(YCB$, 13, 18))
              For JKLY& = 1 To Len(RFF$) Step 4
                RECABI& = CVS(Mid$(RFF$, JKLY&, 4))
                YCB$ = REGLEIDO$("CARBIMON", RECABI&)
                If TRIM$(Mid$(YCB$, 13, 18)) = DOCUCABE$ Then
                  MONEMORI% = CVI(Mid$(YCB$, 31, 2))
                  TICAMBX# = COTIMONEM#     '    TICAMONE(MONEMORI%)
                  Call REPLA(YCB$, MKI$(FC%), 33, 2)
                  Call REPLA(YCB$, MKI$(NCLIEN%), 35, 2)
                  Call REPLA(YCB$, Chr$(4), 37, 1)
                  Call REPLA(YCB$, Chr$(1), 38, 1)
                  Call REPLA(YCB$, MKS$(NROO&), 39, 4)
                  Call REPLA(YCB$, MKI$(SUC%), 43, 2)
                  Call REPLA(YCB$, RECILARGO$, 45, 18)
                  Call REPLA(YCB$, MKI$(1), 63, 2) ' RECIBO SIEMPRE EN PESOS
                  Call REPLA(YCB$, MKD$(0), 65, 8)
                  Call REPLA(YCB$, MKD$(ROUND#(CANCE# * COTIMONEM#)), 73, 8)
                  Call REPLA(YCB$, MKD$(0), 81, 8)
                  Call REPLA(YCB$, MKD$(TICAMBX#), 89, 8)
                  Call REPLA(YCB$, MKD$(0), 97, 8)
                  Call REPLA(YCB$, MKD$(CANCE#), 105, 8)
                  Call REPLA(YCB$, MKD$(0), 113, 8)
                     ' ESTO ES PARA QUE SI SE APLICA DOCUMENTO EN PESOS,
                     ' SOLO CANCELE EL SALDO EN PESOS
                     If MODOMONEI% = 2 And MONEMI% <= 1 Then
                        Call REPLA(YCB$, String$(32, 0), 89, 32)
                     End If
                  'Call REGAPP("CARBIMON", YCB$)
                  '
                  DIFE1# = XVALO(Mid$(List2.LIST(U& - 1), 241, 12))
                  If DIFE1# >= 0.005 Then
                     Call REPLA(YCB$, "AS" + Mid$(RECILARGO$, 3), 45, 18)
                     Call REPLA(YCB$, MKI$(1), 63, 2) ' RECIBO SIEMPRE EN PESOS
                     Call REPLA(YCB$, MKD$(DIFE1#), 65, 8)
                     Call REPLA(YCB$, MKD$(0), 73, 8)
                     Call REPLA(YCB$, MKD$(0), 81, 8)
                     Call REPLA(YCB$, MKD$(TICAMBX#), 89, 8)
                     Call REPLA(YCB$, MKD$(0), 97, 8)
                     Call REPLA(YCB$, MKD$(0), 105, 8)
                     Call REPLA(YCB$, MKD$(0), 113, 8)
                     Call REGAPP("CARBIMON", YCB$)
                     '
                     NumeClie% = NCLIEN%
                     NumeProv% = 0
                     CoComLar = RECILARGO$
                     CoComCor$ = RECICORTO$
                     ImpApli# = DIFE1#
                     CodiEmp% = 0
                     CodiSuc% = SUC%
                     RefApli$ = "Cobranza " + RECICORTO$
                     DocOrLar$ = "AS" + Mid$(RECILARGO$, 3)
                     DocOrCor$ = "AS" + Mid$(RECICORTO$, 3)
                     TotOrig# = DIFE1#
                     IvaOrig# = 0
                     Call GRAPLICOB
                     '
                  End If
                  GoTo 835
                End If
              Next JKLY&
            End If
          Next JKLX&
        ElseIf CANCE# < 0 Then
            ' DIFERENCIA DE CAMBIO PARA CREDITOS: AJUSTES O NOTAS DE CREDITO
            DIFE1# = XVALO(Mid$(List2.LIST(U& - 1), 241, 12))
            If DIFE1# < 0 Then
                NumeClie% = NCLIEN%
                NumeProv% = 0
                CoComLar = RECILARGO$
                CoComCor$ = RECICORTO$
                ImpApli# = DIFE1#
                CodiEmp% = 0
                CodiSuc% = SUC%
                RefApli$ = "Cobranza " + RECICORTO$
                DocOrLar$ = "AS" + Mid$(RECILARGO$, 3)
                DocOrCor$ = "AS" + Mid$(RECICORTO$, 3)
                TotOrig# = DIFE1#
                IvaOrig# = 0
                Call GRAPLICOB
            End If
        End If
835   Next U&
    End If
    Call CIERRARCH("CARBIMON")
    '
    ' APLICACION AUTOMATICA DE CREDITOS
    '
    Call COPYSTRU("SELCUERE", "APLICRE"): Call TRALOCAL("APLICRE", ""): Call SETULTREG("!APLICRE", 0)
    Call COPYSTRU("SELCUERE", "APLIDEB"): Call TRALOCAL("APLIDEB", ""): Call SETULTREG("!APLIDEB", 0)
    '
    JJC& = 0
    Set RSTCREDITOS = CueCorri.OpenRecordset("SELECT * FROM APLICOBRA WHERE CODOCAP_LAR='" & RECILARGO$ & "' AND IMPOAPLI<0 ORDER BY CODOCOR_LAR DESC", dbOpenSnapshot)
    With RSTCREDITOS
        Do While Not .EOF
            DOCULAR$ = !CoDocOr_Lar
            'CajaBanc.FindFirst "CODICOM_LAR='" & DOCULAR$ & "' AND NUCLIEN=" & NCLIEN% & " AND IHABECOMP>0"
            'IMPO# = CajaBanc!IHABECOMP
            IMPO# = Abs(!ImpoApli)
            IMPOS$ = FORMATNUM$(IMPO#, "F13.2")
            JJC& = JJC& + 1
            Y$ = REGBLAN$("!APLICRE")
            Call REPLA(Y$, MKS$(JJC&), 1, 4)
            Call REPLA(Y$, DOCULAR$, 5, 18)
            Call REPLA(Y$, IMPOS$, 23, 13)
            Call GRAREG("!APLICRE", Y$, JJC&)
            .MoveNext
        Loop
    End With
    Call SETULTREG("!APLICRE", JJC&)
    '
    JJD& = 0
    Set RSTDEBITOS = CueCorri.OpenRecordset("SELECT * FROM APLICOBRA WHERE CODOCAP_LAR='" & RECILARGO$ & "' AND IMPOAPLI>0 ORDER BY CODOCOR_LAR ASC", dbOpenSnapshot)
    With RSTDEBITOS
        Do While Not .EOF
            DOCULAR$ = !CoDocOr_Lar
            'CajaBanc.FindFirst "CODICOM_LAR='" & DOCULAR$ & "' AND NUCLIEN=" & NCLIEN% & " AND IDEBECOMP>0"
            'IMPO# = CajaBanc!IDEBECOMP
            IMPO# = Abs(!ImpoApli)
            IMPOS$ = FORMATNUM$(IMPO#, "F13.2")
            JJD& = JJD& + 1
            Y$ = REGBLAN$("!APLIDEB")
            Call REPLA(Y$, MKS$(JJD&), 1, 4)
            Call REPLA(Y$, DOCULAR$, 5, 18)
            Call REPLA(Y$, IMPOS$, 23, 13)
            Call GRAREG("!APLIDEB", Y$, JJD&)
            .MoveNext
        Loop
    End With
    Call SETULTREG("!APLIDEB", JJD&)
    '
    With ApliCobra
        .FindFirst "CODOCAP_LAR='" & RECILARGO$ & "'"
841     If .NoMatch = False Then
            .Delete
            .FindNext "CODOCAP_LAR='" & RECILARGO$ & "'"
            GoTo 841
        End If
    End With
    '
    With CajaBanc
        .FindFirst "CODICOM_LAR='" & RECILARGO$ & "' AND NUCLIEN=" & NCLIEN% & " AND TipoMovim = 'RCOB'"
        IMPO# = !IHABECOMP
        IMPOS$ = FORMATNUM$(IMPO#, "F13.2")
        JJC& = ULTREG&("!APLICRE") + 1
        Y$ = REGBLAN$("!APLICRE")
        Call REPLA(Y$, MKS$(JJC&), 1, 4)
        Call REPLA(Y$, RECILARGO$, 5, 18)
        Call REPLA(Y$, IMPOS$, 23, 13)
        Call GRAREG("!APLICRE", Y$, JJC&)
        Call SETULTREG("!APLICRE", JJC&)
    End With
    '
    For ID& = 1 To JJD&
        XD$ = REGLEIDO$("!APLIDEB", ID&)
        DOCUD$ = Mid$(XD$, 5, 18)
        DOCUD_COR$ = Mid$(DOCUD$, 1, 3) & Mid$(DOCUD$, 11)
        IMPOD# = XVALO(Mid$(XD$, 23, 13))
        IMPOD_ORIG# = IMPOD#
        For IC& = 1 To JJC&
            XC$ = REGLEIDO$("!APLICRE", IC&)
            DOCUC$ = Mid$(XC$, 5, 18)
            If Mid$(DOCUC$, 1, 2) = "NC" Then
                DOCUC_COR$ = Mid$(DOCUC$, 1, 3) & Mid$(DOCUC$, 11, 8)
            Else
                DOCUC_COR$ = Mid$(DOCUC$, 1, 3) & Mid$(DOCUC$, 11, 6) & "-" & Mid$(DOCUC$, 8, 2)
            End If
            IMPOC# = XVALO(Mid$(XC$, 23, 13))
            If IMPOC# < 0.01 Then GoTo 832
            ' IMPORTE A APLICAR
            IMPAPLICAR# = IMPOD#
            If IMPOD# > IMPOC# Then IMPAPLICAR# = IMPOC#
            '
            ' GRABA EN APLICOBRA
            NumeClie% = NCLIEN%
            NumeProv% = 0
            CoComLar = DOCUC$
            CoComCor$ = DOCUC_COR$
            ImpApli# = IMPAPLICAR#
            CodiEmp% = 0
            CodiSuc% = 0
            RefApli$ = "Cobranza " + RECICORTO$
            DocOrLar$ = DOCUD$
            DocOrCor$ = DOCUD_COR$
            TotOrig# = IMPAPLICAR#
            IvaOrig# = 0
            Call GRAPLICOB
            '
            ' GRABA EN CARBIMON
            RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCLIEN%))
            If Len(RFF$) > 0 Then
              ABUSQUE$ = DOCUD$
              For JKLX& = 1 To Len(RFF$) Step 4
                RECABI& = CVS(Mid$(RFF$, JKLX&, 4))
                YCB$ = REGLEIDO$("CARBIMON", RECABI&)
                IMPOCARBI# = CVD(Mid$(YCB$, 65, 8))
                If TRIM$(Mid$(YCB$, 45, 18)) = ABUSQUE$ And (Abs(IMPOCARBI# - IMPOD_ORIG#) < 0.01 And IMPOCARBI# > 0.005) Then
                   MONEMORI% = 2
                  Call REPLA(YCB$, MKI$(FC%), 33, 2)
                  Call REPLA(YCB$, MKI$(NCLIEN%), 35, 2)
                  Call REPLA(YCB$, Chr$(4), 37, 1)
                  Call REPLA(YCB$, Chr$(1), 38, 1)
                    NOOO& = XVALO(Mid$(DOCUC$, 9, 8))
                  Call REPLA(YCB$, MKS$(NOOO&), 39, 4)
                    SUC% = XVALO(Mid$(DOCUC$, 4, 4))
                  Call REPLA(YCB$, MKI$(SUC%), 43, 2)
                  Call REPLA(YCB$, DOCUC$, 45, 18)
                    If InStr(1, YCB$, "AS-X") > 0 Then MONEMORI% = 1
                  Call REPLA(YCB$, MKI$(MONEMORI%), 63, 2) ' RECIBO SIEMPRE EN PESOS
                  Call REPLA(YCB$, MKD$(0), 65, 8)
                  Call REPLA(YCB$, MKD$(IMPAPLICAR#), 73, 8)
                  Call REPLA(YCB$, MKD$(0), 81, 8)
                    CajaBanc.FindFirst "CODICOM_LAR='" & Mid$(YCB$, 45, 18) & "' AND NUCLIEN=" & NCLIEN%
                    COTIZA# = COTIMONEM#
                    If CajaBanc.NoMatch = False Then
                        COTIZA# = XVALO(CajaBanc!Tipo_Cam)
                    End If
                  Call REPLA(YCB$, MKD$(COTIZA#), 89, 8)
                  Call REPLA(YCB$, MKD$(0), 97, 8)
                  Call REPLA(YCB$, MKD$(0), 105, 8)
                  Call REPLA(YCB$, MKD$(0), 113, 8)
                  Call REPLA(YCB$, String$(32, 0), 89, 32)
                    TOTADOLAR# = IMPAPLICAR# / COTIZA#
                    If InStr(1, YCB$, "AS-X") > 0 Then TOTADOLAR# = 0
                  Call REPLA(YCB$, MKD$(TOTADOLAR#), 105, 8)
                  Call REGAPP("CARBIMON", YCB$)
                  Exit For
                End If
              Next JKLX&
            End If
            ' ACTUALIZA REGISTRO DEBITO DESCONTANDO IMPAPLICAR
            IMPOC# = IMPOC# - IMPAPLICAR#
            IMPOS$ = FORMATNUM$(IMPOC#, "F13.2")
            Call REPLA(XC$, IMPOS$, 23, 13)
            Call GRAREG("!APLICRE", XC$, IC&)
            IMPOD# = IMPOD# - IMPAPLICAR#
            '
            ' SI NO HAY MAS CREDITO PASA AL PROXIMO CREDITO
            If IMPOD# < 0.01 Then Exit For
832     Next IC&
    Next ID&
    '
    ' RECUPERAR EL SALDO EN DOLARES
875 For U& = 1 To List2.ListCount
      x$ = List2.LIST(U& - 1)
      CANCE# = CDbl(XVALO(Mid$(x$, 70, 14)))
      If CANCE# >= 0.005 Then
        ABUSQUE$ = TRIM$(Left$(x$, 18))
        Call RECUSALDOL(ABUSQUE$, NCLIEN%)
      End If
    Next U&
    '
880 OKEY% = 6 'VOLVER A COMENTARIO(SOLO PARA PROBAR IMPRESION Y REIMPRESION DE REC. COBRANZAS)
'    '
890 Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    Exit Sub
    'Call BAJALDISCO
    '
GRANOCRED:                  ' GRABA NOTA DE CREDITO
   NFCY$ = TRIM$(Mid$(Frame8.Caption, 9))
30 CoComCor$ = Left$(NFCY$, 3) + Mid$(NFCY$, 11)
   CoComLar$ = NFCY$
   CodiEmp% = 0
   CodiSuc% = 0
   NumeClie% = GRUDIVCLI%(NCLIEN%)
   NumeProv% = 0
   FechEmi = CVDAT(FC%)
   Referen$ = RASOCLI$(NCLIEN%)
   FeContab = CVDAT(FC%)
   FeAcredi = CVDAT(FC%)
   CuentaCo% = CUECOCLI%(NCLIEN%)
   TiMovi$ = "FVEN"
   OpMovi$ = "NC"
   VaMovi$ = Mid$(NFCY$, 18, 1)
   Descrip$ = "Descuento de Caja " + NFCY$
   NumeTal = XVALO(Mid$(NFCY$, 4, 4))
   NumeCom = XVALO(Mid$(NFCY$, 9, 8))
   StatInit = 0
   ImpoDeb = (-1) * ROUND#(XVALO(DESCUEN(0)) * COTIMONEM#)
   Observa$ = RASOCLI$(NCLIEN%)
   Call GRACABAN
   '
   NEFA7# = ROUND#((COTIMONEM# * XVALO(DESCUEN(0))) / (1 + PORCEIVA(0) / 100), 2)
   NIVA8# = ROUND#(COTIMONEM# * XVALO(DESCUEN(0))) - NEFA7#
   'graba venta neta - Verificar que QQQ sea la IC de combinacion Cliente-Tipo de Comprobante
   TOTANETX# = NEFA7#
   If TOTANETX# >= 0.005 Then
     CUEDESCU$ = TRIM$(CONTROL$("", "CUEDESCU"))
     CuentaCo% = CUECOINT%(CUEDESCU)
     TiMovi$ = "VENTA"
     StatInit = 0
     ImpoDeb = TOTANETX#
     Call GRACABAN
   End If
'
   'graba iva inscripto
   TOTANETX# = NIVA8#
   If TOTANETX# >= 0.005 Then
      CUEIVX$ = CONTROL$("SVD451S", "CUEIVA")
      CuentaCo% = CUECOINT%(CUEIVX$)
      TiMovi$ = "IVA"
      OpMovi$ = "DeFi"
      VaMovi$ = TRIM$(FORMATNUM$(PORCEIVA(0), "F6.2")) + " %"
      Descrip$ = "Iva Debito Fiscal " + NFCX$
      StatInit = 0
      ImpoDeb = TOTANETX#
      BasImpo# = NEFA7#
      Call GRACABAN
   End If
   BasImpo# = 0
   '
Return                      ' ***********************************************************
'
End Sub
'
Function RENGLONEC&()
    '
    Call SETULTREG("!PRINRECI", 0)
    Call BLANARCH("!PRINRECI")
    For U& = 1 To ULTREG&("!CARMECO")
      x$ = REGLEIDO$("!CARMECO", U&)
      If x$ = REBLA$ Then GoTo 19
      NUVA$ = Mid$(x$, 19, 12)
      IMVA# = CVD(Mid$(x$, 31, 8))
      FEMII% = CVI(Mid$(x$, 39, 2))
      FEVEI% = CVI(Mid$(x$, 41, 2))
      PROTER$ = UCase$(TRIM$(Mid$(x$, 43, 2)))
      CLEA% = CVI(Mid$(x$, 45, 2))
      '
      TIPOSS$ = Space$(4)
      '
10    Z$ = REGBLAN$("!PRINRECI")
      Call REPLA(Z$, TIPOSS$, 1, 4)
      Call REPLA(Z$, NUVA$, 5, 12)
      Call REPLA(Z$, BANCO$, 17, 16)
      Call REPLA(Z$, MKD$(IMVA#), 33, 8)
      Call REGAPP("!PRINRECI", Z$)
19  Next U&
    '
    V0& = ULTREG&("!PRINRECI")
    V& = 0
    For U& = 1 To List2.ListCount
      x$ = List2.LIST(U& - 1)
      DOCU$ = Mid$(x$, 1, 18)
      CANCE# = CDbl(XVALO(Mid$(x$, 70, 14)))
      If Abs(CANCE#) >= 0.005 Then
        V& = V& + 1
        If V& <= V0& Then
              Z$ = REGLEIDO$("!PRINRECI", V&)
           Else
              Z$ = REGBLAN$("!PRINRECI")
        End If
        Call REPLA(Z$, DOCU$, 53, 18)
        Call REPLA(Z$, MKD$(CANCE#), 71, 8)
        Call GRAREG("!PRINRECI", Z$, V&)
      End If
    Next U&
    '
    RENGLONEC& = ULTREG&("!PRINRECI")
    Call CIERRARCH("!PRINRECI")
    '
End Function
'
Sub PRIRECISTA()
    '
    Dim ORCOP$(6)
    '
    On Error GoTo 0
    NC% = XVALO(Text1)
    '
    ORCOP$(1) = "ORIGINAL"
    ORCOP$(2) = "DUPLICADO"
    ORCOP$(3) = "TRIPLICADO"
    ORCOP$(4) = "CUADRUPLICADO"
    ORCOP$(5) = "QUINTUPLICADO"
    ORCOP$(6) = "SEXTUPLICADO"
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORECIST"))
    'If FORIPRE$ = "" Then
    '    Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Recibo.")
    '    Exit Sub
    'End If
    '
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Recibo de Cobranza:\  \'" + RMCC$ + "'.")
            Exit Sub
        End If
    End If
    '
    CONTA% = ULTREG&("!CARMECO")
    If CONTA% < 1 Then
        Call COMUNI("No se ha Ingresado Detalle de Valores")
        Exit Sub
    End If
    '
    AVAHOJA$ = TRIM$(CONTROL$("RECISTA", "AVAHOJA"))
    CANCELPRINT% = 0
    If CONTROL$("RECISTA", "SUPRINT") = "SI" Then CANCELPRINT% = 1
    '
    Screen.MousePointer = 11
    Call BLANARCH("!PRINRECI")
    For U& = 1 To ULTREG&("!CARMECO")
        x$ = REGLEIDO$("!CARMECO", U&)
        If x$ = REBLA$ Then GoTo 19
        MECO% = CVI(Left$(x$, 2))
        BANCO$ = TRIM$(ECOARCH$("MECOBRA", MKI$(MECO%)))
        NUVA$ = Mid$(x$, 19, 12)
        IMVA# = CVD(Mid$(x$, 31, 8))
        FEMII% = CVI(Mid$(x$, 39, 2))
        FEVEI% = CVI(Mid$(x$, 41, 2))
        PROTER$ = UCase$(TRIM$(Mid$(x$, 43, 2)))
        CLEA% = CVI(Mid$(x$, 45, 2))
        CUITLIB$ = TRIM$(Mid$(x$, 47, 16))
        '
        TIPOSS$ = Space$(4)
        If MECO% > 0 Then
            For V& = 1 To ULTREG&("MECOBRA")
                MCB$ = REGLEIDO$("MECOBRA", V&)
                If CVI(Left$(MCB$, 2)) = MECO% Then
                    TIPOSS$ = Mid$(MCB$, 57, 4)
                    GoTo 10
                End If
            Next V&
        End If
        '
10      Z$ = REGBLAN$("!PRINRECI")
        Call REPLA(Z$, TIPOSS$, 1, 4)
        Call REPLA(Z$, NUVA$, 5, 12)
        Call REPLA(Z$, BANCO$, 17, 16)
        Call REPLA(Z$, MKD$(IMVA#), 33, 8)
        FEX = FEVEI%
        If FEX > 0 Then
            Call REPLA(Z$, FECHATEX$(FEX), 41, 8)
        End If
        Call REGAPP("!PRINRECI", Z$)
        '
        If CUITLIB$ <> "" Then
          FORPRICU$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "CUIT-TRA/FORMATO")))
          If FORPRICU$ <> "" Then
              URE& = ULTREG&("!PRINRECI")
              Z$ = REGLEIDO$("!PRINRECI", URE&)
              Call REPLA(Z$, CUITLIB$, 79, 16)
              Call GRAREG("!PRINRECI", Z$, URE&)
            Else
              Z$ = REGBLAN$("!PRINRECI")
              Call REPLA(Z$, ".CUIT Librad", 5, 12)
              Call REPLA(Z$, CUITLIB$, 17, 16)
              Call REGAPP("!PRINRECI", Z$)
          End If
        End If
        '
19  Next U&
    '
    If Abs(XVALO(Text8)) >= 0.005 Then ' ajuste redondeo para que lo imprima en el formulario
       Z$ = REGBLAN$("!PRINRECI")
       Call REGAPP("!PRINRECI", Z$)
       Z$ = REGBLAN$("!PRINRECI")
       Call REPLA(Z$, "AJU", 1, 4)
       Call REPLA(Z$, "REDONDEO", 5, 12)
       Call REPLA(Z$, MKD$(XVALO(Text8)), 33, 8)
       Call REGAPP("!PRINRECI", Z$)
    End If
    '
    V0& = ULTREG&("!PRINRECI")
    '
    VACON% = 1
    List2.Visible = False
    ILISTA& = List2.ListIndex
    TOLISTA& = List2.TopIndex
    '
    V& = 0
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        x$ = List2.TEXT
        DOCU$ = Mid$(x$, 1, 18)
        CANCE# = CDbl(XVALO(Mid$(x$, 70, 14)))
        If Abs(CANCE#) >= 0.005 Then
            V& = V& + 1
            If V& <= V0& Then
                  Z$ = REGLEIDO$("!PRINRECI", V&)
               Else
                  Z$ = REGBLAN$("!PRINRECI")
            End If
            Call REPLA(Z$, DOCU$, 53, 18)
            Call REPLA(Z$, MKD$(CANCE#), 71, 8)
            Call GRAREG("!PRINRECI", Z$, V&)
        End If
    Next U&
    VACON% = 0
    '
    If XVALO(DESCUEN(0)) >= 0.005 Then
        V& = V& + 1
        If V& <= V0& Then
              Z$ = REGLEIDO$("!PRINRECI", V&)
           Else
              Z$ = REGBLAN$("!PRINRECI")
        End If
        NUNOCRE$ = TRIM$(Mid$(Frame8.Caption, 9))
        Call REPLA(Z$, NUNOCRE$, 53, 18)
        Call REPLA(Z$, MKD$((-1) * XVALO(DESCUEN(0))), 71, 8)
        Call GRAREG("!PRINRECI", Z$, V&)
    End If
    '
    V0& = ULTREG&("!PRINRECI")
    '
    TOTAVAL# = CDbl(XVALO(Label14.Caption))
    TOTACAN# = CDbl(XVALO(Label18.Caption)) + CDbl(XVALO(Label20.Caption)) - CDbl(XVALO(DESCUEN(0)))
    ITEXTO$ = NUMTEX$(TOTAVAL#)
    LOTEX% = XVALO(ATRIFORM$("", "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
    Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
    IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
    IMPLET2$ = TRIM$(NRX2$) + " " + TRIM$(NRX3$) + " ctvs. *****************************************************************************************"
    '
    ANTI# = CDbl(XVALO(Label20.Caption))
    '
    If Abs(ANTI#) >= 0.005 Then
        '
        V& = V& + 1
        If V& <= V0& Then
             Z$ = REGLEIDO$("!PRINRECI", V&)
           Else
             Z$ = REGBLAN$("!PRINRECI")
        End If
        Call REPLA(Z$, Space$(18), 53, 18)
        Call REPLA(Z$, MKD$(0), 71, 8)
        Call GRAREG("!PRINRECI", Z$, V&)
        '
        V0& = ULTREG&("!PRINRECI")
        '
        V& = V& + 1
        If V& <= V0& Then
             Z$ = REGLEIDO$("!PRINRECI", V&)
           Else
             Z$ = REGBLAN$("!PRINRECI")
        End If
        Call REPLA(Z$, "Anticipo N/Aplic.", 53, 18)
        If ANTI# < 0 Then
            Call REPLA(Z$, "Descuento de Caja", 53, 18)
        End If
        Call REPLA(Z$, MKD$(ANTI#), 71, 8)
        Call GRAREG("!PRINRECI", Z$, V&)
        '
    End If
    '
    Call CIERRARCH("!PRINRECI")
    '
30  If TRIM$(FORIPRE$) <> "" Then
        '
        CODFOR$ = TRIM$(UCase$(FORIPRE$))
        MARIZW = XVALO(ATRIFORM$(CODFOR$, "MARIZQMM"))
        MARSUW = XVALO(ATRIFORM$(CODFOR$, "MARSUPMM"))
        '
        INTERLI = XVALO(ATRIFORM$(CODFOR$, "INTERLIN"))
        If INTERLI < 1 Then
            Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
            INTERLI = 1
        End If
        '
        DESMM = XVALO(ATRIFORM$(CODFOR$, "DETDESMM"))
        HASMM = XVALO(ATRIFORM$(CODFOR$, "DETHASMM"))
        '
        CACOPI% = XVALO(ATRIFORM$(CODFOR$, "CAN-COP"))
        If CACOPI% < 1 Or CACOPI% > 6 Then
            Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
            CACOPI% = 1
        End If
        '
        ORIENTA% = 1
        PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "ORIENTA")))
        If PORI$ = "PAISAJE" Or PORI$ = "2" Then
            ORIENTA% = 2
        End If
        '
        PORTUREC$ = TRIM$(CONTROL$(IDENTER$, "PORTREC"))
        If PORTUREC$ = "" Then PORTUREC$ = CONTROL$("*", "PORTREC")
        If PORTUREC$ <> "" Then
          NEWPORT$ = PORTUREC$ + "/NOMESS"
          Call SETPRINPORT(NEWPORT$)
          If NEWPORT$ <> "" Then
            Call SETPRINPORT("RESTORE")
          End If
          PRINTERPORT$ = PORTUREC$
        End If
        '
        If TRIM$(PRINTERPORT$) <> "" Then
           Call SETPRINPORT(PRINTERPORT$)
        End If
        '
        PORIAN = Printer.Orientation  ' retiene orientacion
        '
        NRORECI$ = TRIM$(Mid$(Text11.TEXT, 4))
        '
        Screen.MousePointer = 11
        If AVAHOJA$ <> "NEWPAGE" Then
            'Printer.Orientation = 1
            Call SETSPOOL("ORIENTATION", "1")
            If ORIENTA% = 2 Then
                'Printer.Orientation = 2
                Call SETSPOOL("ORIENTATION", "2")
            End If
        End If
        '
        For COP% = 1 To CACOPI%
            '
            If AVAHOJA$ = "NEWPAGE" Then
               If PRINTERCONTROL% = 1 Then
                  'Printer.NewPage
                  Call SETSPOOL("NEWPAGE", "")
               End If
            End If
            PRINTERCONTROL% = 1
            '
            SUMAVA# = 0: SUMACAN# = 0
            Call CARFORWIN
            '
            Call PRINTFORWIN(TRIM$(Str$(NC%)), "CUEN-CLI", 0, 0)
            Call PRINTFORWIN(TRIM$(RASOCLI$(NC%)), "RASO-CLI", 0, 0)
            Call PRINTFORWIN(DOMICLI$(NC%), "DOMI-CLI", 0, 0)
            Call PRINTFORWIN(CPOSCLI$(NC%), "CPOS-CLI", 0, 0)
            Call PRINTFORWIN(LOCACLI$(NC%), "LOCA-CLI", 0, 0)
            Call PRINTFORWIN(DEPROVI$(PROVCLI$(NC%)), "PCIA-CLI", 0, 0)
            Call PRINTFORWIN(CUITCLI$(NC%), "CUIT-CLI", 0, 0)
            Call PRINTFORWIN(POSIVAL$(NC%), "PIVA-CLI", 0, 0)
            '\\\OT-08-0182-OD-21/04/8///
            Call PRINTFORWIN(Str$(VENDCLI%(NC%)), "VENDEDOR", 0, 0)
            '\\\OT-08-0182-OD-FIN///
            FF% = FECHANUM(TRIM$(Text9.TEXT))
            Call PRINTFORWIN(FETEXTO$(FF%), "FECH-EMI", 0, 0)
            Call PRINTFORWIN(NRORECI$, "NUME-COM", 0, 0)
            Call PRINTFORWIN(ORCOP$(COP%), "ORD-COPI", 0, 0)
            Call PRINTFORWIN(MONEMIS, "MONEDEMI", 0, 0)
            '
            YAC = DESMM
            '
            For YX& = 1 To ULTREG&("!PRINRECI")
               '
               YAC = YAC + INTERLI
               '
               If YAC > HASMM - 2 * INTERLI Then
                   YAC = YAC + INTERLI
                   Call PRINTFORWIN("Transporte", "BANC-CHE", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMAVA#), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                   Call PRINTFORWIN("Transporte", "DOCU-APL", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMACAN#), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                   '
                   Call PRINTFORWIN(String$(20, "*"), "TOTA-VAL", 0, 0)
                   Call PRINTFORWIN(String$(20, "*"), "TOTA-CAN", 0, 0)
                   Call PRINTFORWIN(String$(96, "*"), "IMP-LET1", 0, 0)
                   Call PRINTFORWIN(String$(96, "*"), "IMP-LET2", 0, 0)
                   '
                   If AVAHOJA$ = "NEWPAGE" Then
                        'Printer.NewPage
                        Call SETSPOOL("NEWPAGE", "")
                      Else
                        'Printer.EndDoc
                        Call SETSPOOL("ENDDOC", "")
                   End If
                   '
                   Call CARFORWIN
                   Call PRINTFORWIN(TRIM$(Str$(NC%)), "CUEN-CLI", 0, 0)
                   Call PRINTFORWIN(TRIM$(RASOCLI$(NC%)), "RASO-CLI", 0, 0)
                   Call PRINTFORWIN(DOMICLI$(NC%), "DOMI-CLI", 0, 0)
                   Call PRINTFORWIN(CPOSCLI$(NC%), "CPOS-CLI", 0, 0)
                   Call PRINTFORWIN(LOCACLI$(NC%), "LOCA-CLI", 0, 0)
                   Call PRINTFORWIN(CUITCLI$(NC%), "CUIT-CLI", 0, 0)
                   Call PRINTFORWIN(POSIVAL$(NC%), "PIVA-CLI", 0, 0)
                   FEX = FF%
                   Call PRINTFORWIN(FETEXTO$(FF%), "FECH-EMI", 0, 0)
                   Call PRINTFORWIN(NRORECI$, "NUME-COM", 0, 0)
                   '
                   YAC = DESMM + INTERLI
                   Call PRINTFORWIN("Transporte", "BANC-CHE", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMAVA#), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                   Call PRINTFORWIN("Transporte", "DOCU-APL", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMACAN#), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                   YAC = YAC + 2 * INTERLI
                   '
                End If
                '
                YY$ = REGLEIDO$("!PRINRECI", YX&)
                Call PRINTFORWIN(Left$(YY$, 4), "TIPO-VAL", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 5, 12), "NUME-CHE", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 17, 16), "BANC-CHE", 0, YAC)
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 33, 8), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                SUMAVA# = SUMAVA# + CVD(Mid$(YY$, 33, 8))
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 41, 8), 6, 8, 0, 1), "FECH-CHE", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 53, 18), "DOCU-APL", 0, YAC)
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 71, 8), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                SUMACAN# = SUMACAN# + CVD(Mid$(YY$, 71, 8))
                Call PRINTFORWIN(TRIM$(Mid$(YY$, 79, 16)), "CUIT-TRA", 0, YAC)
                '
                '//OT-8-870-09/10/08
                For i = 0 To List2.ListCount - 1
                  TXTX$ = List2.LIST(i)
                  DOCUMENTO$ = Mid$(YY$, 53, 18)
                  If Mid$(TXTX$, 1, 18) = DOCUMENTO$ Then
                    FECHADOCU$ = Mid$(TXTX$, 23, 8)
                    Call PRINTFORWIN(FECHADOCU$, "FECH-REC", 0, YAC)
                    Exit For
                  End If
                Next i
                '//OT-8-870-FIN
                '
            Next YX&
            '
            Call PRINTFORWIN(CSTRING$(MKD$(TOTAVAL#), 4, 12, 2, 2), "TOTA-VAL", 0, 0)
            Call PRINTFORWIN(CSTRING$(MKD$(TOTACAN#), 4, 12, 2, 2), "TOTA-CAN", 0, 0)
            Call PRINTFORWIN(IMPLET1$, "IMP-LET1", 0, 0)
            Call PRINTFORWIN(IMPLET2$, "IMP-LET2", 0, 0)
                   '
            If AVAHOJA$ <> "NEWPAGE" Then
                 'Printer.EndDoc
                 Call SETSPOOL("ENDDOC", "")
            End If
            '
        Next COP%
        '
        Call CIERRARCH("*.*")
        DESCRIDOC$ = "Recibo Nro. " + NRORECI$ + " - " + TRIM$(RASOCLI$(NC%))
        FEX = FF%
        DESCRIDOC$ = DESCRIDOC$ + " - " + FECHATEX$(FEX)
        TIPODOC$ = "Recibo ": NUMEDOC$ = NRORECI$
        '
        Call SAVESPOOL(DESCRIDOC$, OKX%)
        If TRIM$(PRINTERPORT$) <> "" Then
          Call SETPRINPORT("RESTORE")
        End If
        '
        Screen.MousePointer = 1
        If COMALTER%("Recibo Numero\'" + NRORECI$ + "'\Emitido y Registrado\\Repetir Impresion\Otro Recibo") = 1 Then GoTo 30
        '
        If AVAHOJA$ <> "NEWPAGE" Then
            'Printer.Orientation = PORIAN
            Call SETSPOOL("ORIENTATION", Str$(PORIAN))
        End If
        '
    End If
    Call CIERRARCH("*.*")
    '
End Sub

Sub APLIPAG(FE%, NC%, TI%, VA%, PRF%, NROF&, IMAPLI#, REFE$, DOCORI$, ITOTORI#, IVAORI#)
    '
    ARAPLI$ = "APLICOB": If NC% < 0 Then ARAPLI$ = "APLIPAG"
    PRFS$ = TRIM$(Str$(PRF%)): While Len(PRFS$) < 4: PRFS$ = "0" + PRFS$: Wend
    NROS$ = TRIM$(Str$(NROF&)): While Len(NROS$) < 8: NROS$ = "0" + NROS$: Wend
    '
    TIX% = TI% Mod 10
    TINUVA$ = "FC"                                               ' por default
    If TIX% = 0 Then TINUVA$ = "AS"
    If TIX% = 1 Then TINUVA$ = "VC"
    If TIX% = 3 Then TINUVA$ = "ND"
        If TIX% = 4 Then
        TINUVA$ = "RB"
        If NC% < 0 Then
            TINUVA$ = "OP"
            If TI% = 14 Then TINUVA$ = "FF"
        End If
    End If
    '
    If TIX% = 5 Then TINUVA$ = "DC"
    If TIX% = 6 Then TINUVA$ = "NC"
    If TIX% = 7 Then TINUVA$ = "DP"
    '
    VAD$ = Chr$(64 + VA%): If VAD$ < "A" Or VAD$ > "Z" Then VAD$ = "X"
    TINUVA$ = TINUVA$ + "-" + VAD$ + "-" + PRFS$ + "-" + NROS$
    '
    RECORD$ = MKI$(FE%)
    RECORD$ = RECORD$ + MKI$(Abs(NC%))
    '
    RECORD$ = RECORD$ + TINUVA$
    RECORD$ = RECORD$ + MKD$(IMAPLI#)
    '
    REFE$ = Left$(REFE$ + Space$(24), 24)
    RECORD$ = RECORD$ + REFE$
    '
    RECORD$ = RECORD$ + DOCORI$ + MKD$(ITOTORI#) + MKD$(IVAORI#)
    '
    RECORD$ = RECORD$ + String$(8, 0)
    '
    Call REGAPP(ARAPLI$, RECORD$)
    'Call CIERRARCH(ARAPLI$)
    '
End Sub
'

Private Sub TICAMBIS_Change()
   '
   If TICAMBIS.Enabled = True Then
      TICACH(MONEMI%) = XVALO(TICAMBIS)
   End If
   TICAMBX# = 1
   If MONEMI% > 1 Then
     TICAMBX# = XVALO(TICAMBIS)
   End If
   '
End Sub

Private Sub TICAMBIS_DblClick()
   'Call CARTAMONED
   '
   '\\\OT-07-0086-RG-06/03/07///
   If MONEMI% > 1 Then
     If DERACCE%("COTILOCA", "Ajustar Tipo de Cambio Local") >= 2 Then
        ENTRATICA.Frame4.Caption = "Cotización Local para " + MONEMIS.Caption
        ENTRATICA.Text4 = FORMATNUM(XVALO(TICAMBIS.Caption), "F10.4")
        ENTRATICA.Show 1
        If TRIM$(ENTRATICA.Text4) <> "" Then
          TICAMBIS = TRIM$(ENTRATICA.Text4)
        End If
     End If
     List1.Clear
     List2.Clear
     Label4.Caption = ""
     Label8.Caption = ""
     Label10.Caption = ""
     Label14.Caption = "": Label30 = "": Text8 = ""
     Label18.Caption = ""
     Label20.Caption = ""
     DESCUEN(0) = ""
     Text7 = ""
     Frame8.Caption = "Descuento (N./Cred)"
     BOTNEXT.Enabled = True
   End If
   '\\\OT-07-0086-RG-FIN///
   '
End Sub

Sub CARTAMONED()
   '
   If DERACCE%("TICAMBIO", "Ajustar Tipo de Cambio Central") < 2 Then
     Exit Sub
   End If
   '
   If ULTREG&("TAMONED") > 6 Then
      Call SETULTREG("TAMONED", 6)
      Call COMUNI("Tabla de Monedas Truncada\al Máximo de 6 Registros")
   End If
   '
   Call TRALOCAL("TAMONED", "")
   VTB% = VENTABU%("TAMONED/NC/TITUPAN=Monedas y Tipo de Cambio")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAMONED")
         x$ = REGLEIDO$("!TAMONED", U&)
         Call REPLA(x$, Chr$(U&), 1, 1)
         Call GRAREG("TAMONED", x$, U&)
      Next U&
      FIN& = ULTREG&("!TAMONED")
      Call SETULTREG("TAMONED", FIN&)
   End If
   Call CIERRARCH("TAMONED")
   '
End Sub

Function PRONOCRE$()
   '
   Static FLINUMERA%
   Static NUMECORR%
   Static MONUMERA$
   '
   Screen.MousePointer = 11
   If FLINUMERA% = 0 Then
      FLINUMERA% = FECHANUM("01/01/99")
      If ULTREG&("FORMFAC") > 0 Then
         FFS$ = REGLEIDO$("FORMFAC", 2&)
         FLINUMERA% = CVI(Mid$(FFS$, 19, 2))
      End If
   End If
   '
   If NUMECORR% = 0 Then
      NUMECORR% = 1
      QTRX$ = CONTROL$("NUMEFACT", "NUMECORR")
      If QTRX$ = "INDEPEN" Then
           NUMECORR% = (-1)
         ElseIf QTRX$ = "UNIFICA" Then
           NUMECORR% = (-2)
      End If
   End If
   '
   MONUMERA$ = "AUTO"
   '
   TIPDOCX% = 6
   OPMOVIX$ = "NC"
   '
   VACREDI$ = "A"
   NCLIE% = XVALO(Text1)
   PIVACRE% = PIVACLI%(NCLIE%)
   If PIVACRE% = 0 Or PIVACRE% = 3 Or PIVACRE% = 4 Or PIVACRE% = 6 Then
        VACREDI$ = "B"
      ElseIf PIVACRE% = 5 Or PIVACRE% = 7 Then
        VACREDI$ = "X"
        If TRIM$(LETREXPO$) <> "" Then
          VACREDI$ = LETREXPO$
        End If
   End If
   '
   DATEDOC% = FECHANUM(Text9.TEXT)
   If DATEDOC% < FLINUMERA% Then
      Call MENSERR(24, "Imposible Registrar en la Fecha Elegida.\  \La Fecha de Recibo Debe Ser POSTERIOR\a la Fecha de Nueva Numeración.")
      Call CIERRARCH("*.*")
      Call FINAL("")
   End If
   FLINUME = CDbl(CVDAT(FLINUMERA%))
   '
   FEXUL = DATEDOC%
   FLIDIAS = CDbl(CVDAT(DIASANTES(FEXUL, 15)))
   '//////////*******NOTA*******////////
   'Con posterioridad se detecto un Error en la Impòrtacion
   '(VARIMOVIM(A,B), SE IMPORTO EN VEZ DE LA LETRA CON EL OPCIMOVIX$(ND,NC).
   'Por tal motivo el sistema no toma en cuenta los numero
   'de NC-ND inferiores al 01/11/04.
   'Pero si lo hace a futuro. es decir yo puedo reiniciar numeración
   'al 01/11/04
   
   FLIMI = FLIDIAS: MODORAPIDO% = 1
   If FLINUME > FLIMI Then
      FLIMI = FLINUME
      MODORAPIDO% = 0
   End If
   '
22 Call APERBASDAT("CABAN")
   ASELE$ = "SELECT NumeComp FROM CajaBanc "
   '
   If OPMOVIX$ = "ND" Or OPMOVIX$ = "NC" Then
     If NUMECORR% = (-2) Then    ' Unificada debitos y creditos
       ASELE$ = ASELE$ + "WHERE ((TipoMovim = 'FVEN' AND (OpciMovim = 'ND' OR OpciMovim = 'NC')) "
       ASELE$ = ASELE$ + "OR (TipoMovim = 'FCOM' AND (OpciMovim = 'DP' OR OpciMovim = 'CP'))) "
       GoTo 25
     End If
   End If
   '
   ASELE$ = ASELE$ + "WHERE TipoMovim = 'FVEN' "
   If NUMECORR% <> 1 Then
       ASELE$ = ASELE$ + "AND OpciMovim = '" & OPMOVIX$ & "' "
     Else
       '       ASELE$ = ASELE$ + "AND (OpciMovim = 'FC' OR OpciMovim = 'ND' OR OpciMovim = 'NC') "
       '       //////****NUEVA CONDICION PARA CALCULAR PROX. Nº FAC***////
       '       ESTA CONDICION LA AGREGUE  PARA QUE TOME EN CTA. LOS Nº
       '       DE COMPROBANTES ARRUINADOS: OR OpciMovim = 'ANUL'
       ASELE$ = ASELE$ + "AND (OpciMovim = 'FC' OR OpciMovim = 'ND' OR OpciMovim = 'NC' OR OpciMovim = 'ANUL') "
   End If
   '
25 ASELE$ = ASELE$ + "AND VariMovim = '" & VACREDI$ & "' "
   ASELE$ = ASELE$ + "AND (NumeTalo = 0 OR NumeTalo = 1) "
   ASELE$ = ASELE$ + "AND FechEmisi >= " & FLIMI & " "
   ASELE$ = ASELE$ & "ORDER BY NumeComp DESC;"
   '
   Set NUFAXI = CueCorri.OpenRecordset(ASELE$, dbOpenDynaset, dbReadOnly)
   '
   NUMAXI& = 0
   On Error Resume Next
   NUFAXI.MoveFirst
   If Err < 1 Then
       On Error GoTo 0
       NUMAXI& = NUFAXI!NumeComp
     Else
       On Error GoTo 0
       If MODORAPIDO% = 1 Then
          FLIMI = FLINUME
          MODORAPIDO% = 0
          GoTo 22
       End If
   End If
   '
99 If Visible = True Then
     If NUMAXI& < 1 Then
       NUMAXI& = XVALO(InputBox$("Ingrese Número de la Próxima Nota de Crédito '" + VACREDI$ + "'", "Incicialización de Talonarios")) - 1
       If NUMAXI& < 0 Then GoTo 99
     End If
   End If
   NPP& = NUMAXI& + 1
   NFCZ$ = TRIM$(Str$(NPP&)): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
   NFCY$ = OPMOVIX$ + "-0001-" + NFCZ$ + "-" + VACREDI$

   PRONOCRE$ = TRIM$(NFCY$)
   Screen.MousePointer = 1
   '
End Function
'
Sub PRICREDESC()
    '
    CAHOJAS% = 0
    TIPOCOM$ = "(T.Libre)"
    TIFADOC% = 6 ' NOTA DE CREDITO
    NFCY$ = TRIM$(Mid$(Frame8.Caption, 9))
    VARDOCX$ = TRIM$(Mid$(NFCY$, 18, 1))
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORFAC-" + VARDOCX$))
    If VARDOCX$ = "A" Or VARDOCX$ = "C" Then
        FORIALT$ = CONTROL$("", "FORNCR-A") 'NC-A
      ElseIf UCase$(Label13.Caption) = "B" Then
        FORIALT$ = CONTROL$("", "FORNCR-B") 'NC-B
    End If
    If TRIM$(FORIALT$) <> "" Then
      FORIPRE$ = FORIALT$
    End If
    ''\\\OT-06-0038-WAR-FIN///
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Falta Definir Controlador de Impresión")
       Exit Sub
    End If
    '
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Documento.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    CODFOR$ = FORIPRE$
    For KKI% = 0 To CACOLTAB1%
      For KKJ% = 0 To CAITAB1%
        TETABU1$(KKI%, KKJ%) = ""
      Next KKJ%
    Next KKI%
    '
    Printer.TrackDefault = False
    PORTUFAC$ = TRIM$(CONTROL$(IDENTER$, "PORTFAC"))
      If PORTUFAC$ = "" Then PORTUFAC$ = CONTROL$("*", "PORTFAC")
    PORTUFAC1$ = TRIM$(PORTUFAC$)
    If PORTUFAC1$ = "" Then PORTUFAC1$ = Printer.DeviceName
    PPORT$ = REPLACAR$(PORTUFAC1$, "\", "/")
    NNCRE$ = APRORECIB.Label29
    Call COMUNI("Se Emitirá Nota de Crédito '" + NNCRE$ + "'\por Impresora '" + PPORT$ + "'.")
    '
    Screen.MousePointer = 11
    ITPORHOJA% = 32767
    '
    CANCELPRINT% = 0
    FECHDOC$ = Text9.TEXT
    NUMEDOC$ = NFCY$
    '
    TIPODOC$ = "Nota de Credito": CODOCUM$ = "CRE"
    CPAG% = 0
    DESTIDOC% = XVALO(Text1.TEXT)
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    'CODTABU1$(4) = "PRE-LIST"
    'CODTABU1$(5) = "PORDESCU"
    'CODTABU1$(6) = "PRE-NETO"
    'CODTABU1$(7) = "NET-ITEM"
    'CODTABU1$(8) = "FECH-REC"
    CODTABU1$(9) = "ORD-ITEM"
    CODTABU1$(10) = "UNIMED"
    CODTABU1$(11) = "REF-MAT1"
    CODTABU1$(12) = "REF-MAT2"
    CODTABU1$(13) = "REF-MAT3"
    CODTABU1$(14) = "N-PLANO"
    CODTABU1$(15) = "N-ORIGIN"
    CODTABU1$(16) = "N-OCOMPR"
    CODTABU1$(17) = "N-REMITO"
    CODTABU1$(18) = "N-BOLREC"
    CODTABU1$(19) = "FECH-REC"
    CODTABU1$(20) = "PRE-LIST"
    CODTABU1$(21) = "PORDESCU"
    CODTABU1$(22) = "PRE-NETO"
    CODTABU1$(23) = "NET-ITEM"
    CODTABU1$(24) = "ALICUIVA"
    CODTABU1$(25) = "IVA-ITEM"
    CODTABU1$(26) = "BRU-ITEM"
    '
    CACOLTAB1% = 26
    '
    CAITAB1% = 0
    CAITCRU% = 0
    '
    SUMATOTA# = XVALO(DESCUEN(0).TEXT)
    IDESCUEN# = 0
    '
    NroRemi$ = ""
    NUOCOMPRA$ = ""
    NROPEDIDO$ = ""
    '
    COEFIVAINCLU# = 1
    If VARDOCX$ = "B" Then
      COEFIVAINCLU# = 1 + ALIVA / 100
    End If
    '
    CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = Label13.Caption
    CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = NroRemi$
    CODPARAM$(3) = "NRO-OCOM":  DOCPARAM$(3) = NUOCOMPRA$
    CODPARAM$(4) = "NRO-AENT":  DOCPARAM$(4) = NROPEDIDO$
    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = ""
    'CODPARAM$(6) = "VENDEDOR": DOCPARAM$(6) = Eco21(2).TEXT
    'CODPARAM$(7) = "COND-PAG": DOCPARAM$(7) = Eco21(1).TEXT
    'CODPARAM$(8) = "POR-IVA": DOCPARAM$(8) = PORIVA(0).Caption
    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = ""
    'CODPARAM$(10) = "POR-IVA1": DOCPARAM$(10) = PORIVA(1).Caption
    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = ""
    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = ""
    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = ""
    'CODPARAM$(4) = "NRO-OCOM": DOCPARAM$(4) = Text10.TEXT
    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = ""
    'CODPARAM$(15) = "POR-DESC": DOCPARAM$(15) = TRIM$(PORDESCU.TEXT)
    CODPARAM$(16) = "IMP-RNI": DOCPARAM$(16) = ""
    CODPARAM$(17) = "IMP-LET1": DOCPARAM$(17) = ""
    CODPARAM$(18) = "IMP-LET2": DOCPARAM$(18) = ""
    CODPARAM$(19) = "IMP-LET3": DOCPARAM$(19) = ""
    CODPARAM$(23) = "TOTA-DOL": DOCPARAM$(23) = ""
    '
    'CODPARAM$(20) = "DOMI-TRA": DOCPARAM$(20) = LENTREGA(0).TEXT
    'CODPARAM$(21) = "DESTINO": DOCPARAM$(21) = LENTREGA(1).TEXT
    'CODPARAM$(22) = "TRANSPOR": DOCPARAM$(22) = LENTREGA(2).TEXT
    CODPARAM$(24) = "MONEDEMI": DOCPARAM$(24) = MONEMIS
    CODPARAM$(25) = "PERC-BRU": DOCPARAM$(25) = ""
    CODPARAM$(26) = "SUB-TOTA": DOCPARAM$(26) = ""
    CODPARAM$(27) = "PERC-IVA": DOCPARAM$(27) = "" 'PERCIVA
    CODPARAM$(28) = "FECH-VEN": DOCPARAM$(28) = Format$(FVENCIM, "dd/mm/yy") ' FECHA DE VENCIMIENTO
      On Error Resume Next
      FEXO = FECHANUM(FECHDOC$)
      FEXU = CVINT(FVENCIM)
      CADIVEN = DIENTRE(FEXO, FEXU)
      On Error GoTo 0
    CODPARAM$(32) = "FECH-LAN": DOCPARAM$(32) = FORMATNUM$(CADIVEN, "F5.0")
    '\\\OT-5-579-FG-16/08/05
    'CODPARAM$(29) = "TIPO-VAL": DOCPARAM$(29) = TRIM$(Text18) ' VALOR FOB
    CODPARAM$(30) = "TICAMBIO": DOCPARAM$(30) = FORMATNUM$(TICAMONE(2), "F7.4") ' TIPO DE CAMBIO
    CODPARAM$(31) = "CTAC-CLI": DOCPARAM$(31) = CONTACTCLI$(DESTIDOC%) '  CONTACTO
    '\\\OT-5-579-FG-FIN
    '\\\OT-06-0125-WAR-03/04/06///
    'CODPARAM$(33) = "NUBULTOS": DOCPARAM$(33) = TRIM$(FINFACTURA.Text2(2)) ' BULTOS
    'CODPARAM$(34) = "PES-CONJ": DOCPARAM$(34) = TRIM$(FINFACTURA.Text2(11)) + " Kg" ' PESO NETO
    'CODPARAM$(35) = "PES-COMP": DOCPARAM$(35) = TRIM$(FINFACTURA.Text2(4)) + " Kg" ' PESO BRUTO
    'CODPARAM$(36) = "BANC-CHE": DOCPARAM$(36) = TRIM$(FINFACTURA.Text2(5)) 'BANCO EXTERIOR
    'CODPARAM$(37) = "DOCU-APL": DOCPARAM$(37) = TRIM$(FINFACTURA.Text2(8)) ' BANCO LOCAL
    'CODPARAM$(38) = "NUME-CHE": DOCPARAM$(38) = TRIM$(FINFACTURA.Text2(10)) ' CREDITO
    'CODPARAM$(39) = "DEN-CUEN": DOCPARAM$(39) = TRIM$(FINFACTURA.Text2(9)) ' GUIA DE IMPORTACION
    'CODPARAM$(40) = "REF-ASIE": DOCPARAM$(40) = TRIM$(FINFACTURA.Text2(6)) ' PUERTO DE EMBARQUE
    'CODPARAM$(41) = "DET-PASE": DOCPARAM$(41) = TRIM$(FINFACTURA.Text2(7)) ' DESTINO
    'CODPARAM$(42) = "MEDIDA-1": DOCPARAM$(42) = TRIM$(IEMBALA)             ' gasto de embalaje
    CODPARAM$(43) = "ZONA-VEN": DOCPARAM$(43) = TRIM$(Str$(GRUCOCLI%(DESTIDOC%)))
    '\\\OT-06-0125-WAR-FIN///
    'AGREGUEMH04-10
    'CODPARAM$(44) = "NRO-PASE": DOCPARAM$(44) = TRIM$(Label27.Caption) 'Sucursal
    '
    CAPARDOC% = 44 '32 '\\\OT-06-0125-WAR-03/04/06///
    '
    Call SETULTREG("!OBSERVOF", 0)
    '
    PRINTERPORT$ = PORTUFAC$
    CACOPDOC% = XVALO(FINFACTURA.Text2(0).TEXT)
    If CACOPDOC% < 1 Then CACOPDOC% = 1
    '
    NOCONFIRM% = 0
    HOJASUEL% = 1
    '
    CIPRE% = 0
    CAPEX$ = 1       ' cantidad
    '
    NPEDIDD& = 0
    SUBCO$ = ""
    '
    CAITAB1% = CAITAB1% + 1
    CODDX$ = ""
    TETABU1$(1, CAITAB1%) = CODDX$
    TETABU1$(2, CAITAB1%) = "Descuento " + TRIM$(Text7) + " % S/Recibo Nro. " + TRIM$(Text11)
    TETABU1$(3, CAITAB1%) = CAPEX$
    TETABU1$(9, CAITAB1%) = Str$(U&)
    TETABU1$(10, CAITAB1%) = ""    ' unidad de medida
    TETABU1$(11, CAITAB1%) = ""    ' referencia 1
    TETABU1$(12, CAITAB1%) = ""    ' referencia 1
    TETABU1$(13, CAITAB1%) = ""    ' referencia 1
    TETABU1$(14, CAITAB1%) = ""    ' numero de plano
    TETABU1$(15, CAITAB1%) = ""    ' numero original
    TETABU1$(16, CAITAB1%) = ""    ' numero de o.compra
    TETABU1$(17, CAITAB1%) = ""    ' numero de remito
    TETABU1$(18, CAITAB1%) = ""    ' numero de doc-recepcion
    TETABU1$(19, CAITAB1%) = ""    ' fecha de recepcion
    '
    FACIVITEM = 1: If VARDOCX$ = "B" Then FACIVITEM = 1 + ALIVA / 100
    TETABU1$(20, CAITAB1%) = ""    ' precio de lista
    TETABU1$(21, CAITAB1%) = ""    ' PORCENTAJE DE DESCUENTO
    '
    TETABU1$(22, CAITAB1%) = Str$(FACIVITEM * XVALO(DESCUEN(0)) / (1 + ALIVA / 100)) ' precio unitario
    TETABU1$(23, CAITAB1%) = Str$(XVALO(DESCUEN(0)) / (1 + ALIVA / 100))       ' total neto
    TETABU1$(24, CAITAB1%) = Str$(ALIVA)                         ' alicuota iva
    TETABU1$(25, CAITAB1%) = Str$(ALIVA / 100 * XVALO(DESCUEN(0)) / (1 + ALIVA / 100)) ' importe iva del item
    TETABU1$(26, CAITAB1%) = DESCUEN(0) ' total bruto
    '
    CODPARAM$(5) = "IMP-NETO": DOCPARAM$(5) = Str$(XVALO(DESCUEN(0)) / (1 + ALIVA / 100))
    CODPARAM$(9) = "IMP-IVA": DOCPARAM$(9) = Str$(ALIVA / 100 * XVALO(DESCUEN(0)) / (1 + ALIVA / 100))
    CODPARAM$(11) = "IMP-IVA1": DOCPARAM$(11) = ""
    CODPARAM$(12) = "IMP-DESC": DOCPARAM$(12) = ""
    CODPARAM$(13) = "IMPNEDES": DOCPARAM$(13) = Str$(XVALO(DESCUEN(0)) / (1 + ALIVA / 100))
    CODPARAM$(14) = "IMP-TOTA": DOCPARAM$(14) = DESCUEN(0)
    CODPARAM$(16) = "IMP-RNI": DOCPARAM$(16) = ""
    If TICAMBIO > 0 Then
       VADOLA# = XVALO(DESCUEN(0)) / TICAMBIO
       DOCPARAM$(23) = TRIM$(CSTRING$(MKD$(VADOLA#), 4, 12, 2, 2))
    End If
    CODPARAM$(25) = "PERC-BRU": DOCPARAM$(25) = ""
    DOCPARAM$(26) = ""
    DOCPARAM$(27) = ""                          'PARA IMPRIMIR PERC-IVA
    '
    If VARDOCX$ = "B" Then
      DOCPARAM$(5) = FORMATNUM$(COEFIVAINCLU# * (XVALO(DESCUEN(0)) / (1 + ALIVA / 100)), "F13.2")
      'DOCPARAM$(12) = FORMATNUM$(COEFIVAINCLU# * XVALO(Total(1).TEXT), "F13.2")
      DOCPARAM$(13) = FORMATNUM$(COEFIVAINCLU# * XVALO(DESCUEN(0)) / (1 + ALIVA / 100), "F13.2")
      'DOCPARAM$(42) = FORMATNUM$(COEFIVAINCLU# * XVALO(IEMBALA), "F13.2")          ' gasto de embalaje
    End If
    '
    CODFOR$ = FORIPRE$
    LOIMPLET% = XVALO(ATRIFORM$(CODFOR$, "LOIMPLET"))
    If LOIMPLET% < 16 Then LOIMPLET% = 32
    TTTT# = XVALO(DESCUEN(0))
    TOTFAC$ = NUMTEX$(TTTT#)
    If MONECO$ <> "" Then
      PPXX% = InStr(TOTFAC$, "PESOS")
      If PPXX% > 0 Then
        TOTFAC$ = Left$(TOTFAC$, PPXX% - 1) + UCase$(MONECO$) + Mid$(TOTFAC$, PPXX% + 5)
      End If
    End If
    Call FRATEX(TOTFAC$, NRX1$, NRX2$, NRX3$, LOIMPLET%)
    CODPARAM$(17) = "IMP-LET1": DOCPARAM$(17) = NRX1$
    CODPARAM$(18) = "IMP-LET2": DOCPARAM$(18) = NRX2$
    CODPARAM$(19) = "IMP-LET3": DOCPARAM$(19) = NRX3$
    '
    If CACOPDOC% > 0 Then
      Call PRINTDOC("@" + FORIPRE$)
    End If
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    For KKI% = 0 To CACOLTAB1%
      For KKJ% = 0 To CAITAB1%
        TETABU1$(KKI%, KKJ%) = ""
      Next KKJ%
    Next KKI%
    '
End Sub

Sub LIMPPROM()
  Label17 = ""
  Label19 = ""
  Label23 = ""
  PROVENCI# = 0
End Sub

Private Sub Timer1_Timer()
   '
   SUVALO# = 0
   For U& = 1 To ULTREG&("!CARMECO")
      x$ = REGLEIDO$("!CARMECO", U&)
      SUVALO# = SUVALO# + ROUND#(CVD(Mid$(x$, 31, 8)))
   Next U&
   Label30 = FORMATNUM(SUVALO#, "F12.2")
   Label14 = FORMATNUM$(SUVALO# + XVALO(Text8), "F12.2")
   FALTAPACAN# = XVALO(Label8) * ((100 - XVALO(Text7)) / 100) - XVALO(Label14)
   Label29.ForeColor = RGB(0, 0, 0)
   If FALTAPACAN# < 0 Then
      Label29.ForeColor = RGB(255, 0, 0)
   End If
   Label29 = FORMATNUM$(FALTAPACAN#, "F10.2")
   Label29.Refresh
   '
End Sub

Private Sub visualizarimputaciondecreditosyanticipos_Click()
   Call RECICO04.Command11_Click
End Sub
