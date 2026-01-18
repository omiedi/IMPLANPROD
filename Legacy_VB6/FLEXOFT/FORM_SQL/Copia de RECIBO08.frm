VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form RECIBO08 
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
      Left            =   1260
      Sorted          =   -1  'True
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0F0&
      BorderStyle     =   0  'None
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
      Begin VB.PictureBox Picture999 
         Height          =   540
         Index           =   0
         Left            =   2280
         ScaleHeight     =   480
         ScaleWidth      =   1110
         TabIndex        =   98
         Top             =   420
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Neto Cobranza"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   210
         TabIndex        =   63
         Top             =   7190
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
            TabIndex        =   64
            Top             =   315
            Width           =   1485
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Aplicados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   6960
         TabIndex        =   61
         Top             =   7190
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
            TabIndex        =   62
            Top             =   315
            Width           =   1485
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Anticipos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   8880
         TabIndex        =   59
         Top             =   7190
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
            TabIndex        =   60
            Top             =   315
            Width           =   1485
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Descuento (N/Cred)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   4035
         TabIndex        =   55
         Top             =   7190
         Visible         =   0   'False
         Width           =   2820
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
            TabIndex        =   57
            Top             =   315
            Width           =   645
         End
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
            TabIndex        =   56
            Top             =   315
            Width           =   1485
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
            TabIndex        =   58
            Top             =   300
            Width           =   330
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Saldo Cancela"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   2100
         TabIndex        =   53
         Top             =   7190
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
            TabIndex        =   54
            Top             =   315
            Width           =   1485
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Prom. de Pagos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1250
         Left            =   6195
         TabIndex        =   46
         Top             =   105
         Width           =   1665
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
            TabIndex        =   52
            Top             =   315
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
            TabIndex        =   51
            Top             =   630
            Width           =   540
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
            TabIndex        =   50
            Top             =   945
            Width           =   540
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00D0F0F0&
            Caption         =   "Pr.venc:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   185
            Left            =   60
            TabIndex        =   49
            Top             =   315
            Width           =   855
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00D0F0F0&
            Caption         =   "Pr.Val:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   185
            Left            =   60
            TabIndex        =   48
            Top             =   630
            Width           =   855
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00D0F0F0&
            Caption         =   "Dias Atr:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   185
            Left            =   60
            TabIndex        =   47
            Top             =   945
            Width           =   855
         End
      End
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
         TabIndex        =   42
         Top             =   3885
         Width           =   1275
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   2355
         Left            =   240
         ScaleHeight     =   2295
         ScaleWidth      =   10395
         TabIndex        =   41
         Top             =   1440
         Width           =   10450
         Begin VB.Frame Frame17 
            Caption         =   "Frame6"
            Height          =   2900
            Left            =   10000
            TabIndex        =   82
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame16 
            Caption         =   "Frame6"
            Height          =   2900
            Left            =   8500
            TabIndex        =   81
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame15 
            Caption         =   "Frame6"
            Height          =   2900
            Left            =   7000
            TabIndex        =   80
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame14 
            Caption         =   "Frame6"
            Height          =   2900
            Left            =   5700
            TabIndex        =   79
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame13 
            Caption         =   "Frame6"
            Height          =   2900
            Left            =   5000
            TabIndex        =   78
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame6 
            Caption         =   "Frame6"
            Height          =   2900
            Left            =   4100
            TabIndex        =   77
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame4 
            Caption         =   "Frame6"
            Height          =   2900
            Left            =   2600
            TabIndex        =   76
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame3 
            Caption         =   "Frame6"
            Height          =   2900
            Left            =   600
            TabIndex        =   75
            Top             =   0
            Width           =   15
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00C0C0C0&
            Height          =   340
            Left            =   0
            ScaleHeight     =   285
            ScaleWidth      =   10395
            TabIndex        =   72
            Top             =   0
            Width           =   10455
            Begin VB.PictureBox Picture5 
               Height          =   350
               Left            =   0
               ScaleHeight     =   285
               ScaleWidth      =   10350
               TabIndex        =   73
               Top             =   0
               Width           =   10410
               Begin VB.Label Label12 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " M-P    Medio de Pago      Numero              P/T       Cl.         Emision           Vencim.               Importe"
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
                  TabIndex        =   74
                  Top             =   0
                  Width           =   10410
               End
            End
            Begin VB.Line Line5 
               X1              =   7140
               X2              =   7140
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line6 
               X1              =   630
               X2              =   630
               Y1              =   0
               Y2              =   1155
            End
            Begin VB.Line Line7 
               X1              =   2625
               X2              =   2625
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line13 
               X1              =   4200
               X2              =   4200
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line20 
               X1              =   5670
               X2              =   5670
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line22 
               X1              =   4935
               X2              =   4935
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line23 
               X1              =   8505
               X2              =   8505
               Y1              =   0
               Y2              =   1050
            End
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
            Height          =   2310
            ItemData        =   "RECIBO08.frx":0000
            Left            =   0
            List            =   "RECIBO08.frx":0002
            TabIndex        =   71
            Top             =   360
            Width           =   10410
         End
      End
      Begin VB.CommandButton Command23 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   10860
         Picture         =   "RECIBO08.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Ver e Imprimir Copia de Recibo ya Emitido"
         Top             =   6420
         Width           =   840
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   10860
         Picture         =   "RECIBO08.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Aprueba, Graba e Imprime Documento"
         Top             =   7270
         Width           =   840
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
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
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   265
         Width           =   3060
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
         Top             =   265
         Width           =   435
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00004040&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6225
         Left            =   10860
         ScaleHeight     =   6165
         ScaleWidth      =   1005
         TabIndex        =   24
         Top             =   0
         Width           =   1065
         Begin VB.CommandButton Command10 
            Caption         =   "B.Ch."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   150
            Picture         =   "RECIBO08.frx":0618
            Style           =   1  'Graphical
            TabIndex        =   101
            ToolTipText     =   "Busqueda de Cheque x Número"
            Top             =   4500
            Width           =   650
         End
         Begin VB.CommandButton Command24 
            Caption         =   "B.Ch."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   150
            Picture         =   "RECIBO08.frx":0762
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Busqueda de Cheque x Número"
            Top             =   4050
            Width           =   650
         End
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   125
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   5600
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   34
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   650
            Left            =   150
            Picture         =   "RECIBO08.frx":08AC
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Abandona la Aplicación"
            Top             =   50
            Width           =   650
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Notas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   650
            Left            =   150
            Picture         =   "RECIBO08.frx":09F6
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Notas de Clientes"
            Top             =   3390
            Width           =   650
         End
         Begin VB.CommandButton Command4 
            Caption         =   "C.Cte."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   650
            Left            =   150
            Picture         =   "RECIBO08.frx":0D00
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Resumen de Cuenta Corriente"
            Top             =   1390
            Width           =   650
         End
         Begin VB.CommandButton Command11 
            Caption         =   "New"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   650
            Left            =   150
            Picture         =   "RECIBO08.frx":100A
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Apertura de Nueva Cuenta de Cliente"
            Top             =   2720
            Width           =   650
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Setup"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   650
            Left            =   150
            Picture         =   "RECIBO08.frx":144C
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   4950
            Width           =   650
         End
         Begin VB.CommandButton BOTNEXT 
            Caption         =   "Next"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   650
            Left            =   150
            Picture         =   "RECIBO08.frx":188E
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Continúa Carga de Recibo."
            Top             =   720
            Width           =   650
         End
         Begin VB.CommandButton Command8 
            Caption         =   "More"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   650
            Left            =   150
            Picture         =   "RECIBO08.frx":1B98
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Anular, Re-Imprimir, Imputación de Créditos, Tablas"
            Top             =   2050
            Width           =   650
         End
         Begin VB.Image Image2 
            Height          =   390
            Left            =   -300
            Picture         =   "RECIBO08.frx":1CE2
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
         TabIndex        =   22
         Top             =   630
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         Height          =   2800
         Left            =   240
         ScaleHeight     =   2745
         ScaleWidth      =   10395
         TabIndex        =   16
         Top             =   4300
         Width           =   10450
         Begin VB.Frame Frame23 
            Caption         =   "Frame6"
            Height          =   2400
            Left            =   10100
            TabIndex        =   97
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame22 
            Caption         =   "Frame6"
            Height          =   2400
            Left            =   8500
            TabIndex        =   96
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame21 
            Caption         =   "Frame6"
            Height          =   2400
            Left            =   6800
            TabIndex        =   95
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame20 
            Caption         =   "Frame6"
            Height          =   2400
            Left            =   5400
            TabIndex        =   94
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame19 
            Caption         =   "Frame6"
            Height          =   2800
            Left            =   3720
            TabIndex        =   93
            Top             =   0
            Width           =   15
         End
         Begin VB.Frame Frame24 
            Caption         =   "Frame6"
            Height          =   2400
            Left            =   2500
            TabIndex        =   92
            Top             =   0
            Width           =   15
         End
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
            ItemData        =   "RECIBO08.frx":3C04
            Left            =   0
            List            =   "RECIBO08.frx":3C06
            TabIndex        =   91
            Top             =   340
            Width           =   10410
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
            Left            =   0
            TabIndex        =   90
            Top             =   360
            Width           =   10450
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5630
            Left            =   0
            TabIndex        =   84
            Top             =   360
            Width           =   10515
            Begin VB.Line Line19 
               X1              =   6720
               X2              =   6720
               Y1              =   2100
               Y2              =   2415
            End
            Begin VB.Line Line18 
               X1              =   5460
               X2              =   5460
               Y1              =   2100
               Y2              =   2415
            End
            Begin VB.Line Line17 
               X1              =   3730
               X2              =   3730
               Y1              =   2100
               Y2              =   2415
            End
            Begin VB.Line Line16 
               X1              =   2520
               X2              =   2520
               Y1              =   2100
               Y2              =   2415
            End
            Begin VB.Line Line15 
               X1              =   0
               X2              =   11295
               Y1              =   2415
               Y2              =   2415
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
               TabIndex        =   89
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   260
               Left            =   6730
               TabIndex        =   88
               Top             =   2110
               Width           =   1650
            End
            Begin VB.Line Line12 
               X1              =   1680
               X2              =   1680
               Y1              =   1575
               Y2              =   3360
            End
            Begin VB.Line Line4 
               X1              =   8400
               X2              =   8400
               Y1              =   0
               Y2              =   2415
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
               TabIndex        =   87
               Top             =   120
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   8530
               TabIndex        =   86
               Top             =   2110
               Width           =   1695
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3720
               TabIndex        =   85
               Top             =   2110
               Width           =   1695
            End
            Begin VB.Line Line14 
               BorderWidth     =   2
               X1              =   0
               X2              =   10395
               Y1              =   2085
               Y2              =   2085
            End
         End
         Begin VB.PictureBox Picture6 
            Height          =   330
            Left            =   0
            ScaleHeight     =   270
            ScaleWidth      =   10350
            TabIndex        =   67
            Top             =   0
            Width           =   10410
            Begin VB.Frame Frame18 
               Caption         =   "Frame6"
               Height          =   2100
               Left            =   0
               TabIndex        =   83
               Top             =   0
               Width           =   15
            End
            Begin VB.Label Label9 
               BorderStyle     =   1  'Fixed Single
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
               TabIndex        =   68
               Top             =   0
               Width           =   10410
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   0
            TabIndex        =   17
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
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
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
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
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
            Name            =   "Arial"
            Size            =   9
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
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   1000
         Width           =   2220
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
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   1000
         Width           =   750
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
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
         Tag             =   "TBOX"
         Text            =   " "
         Top             =   630
         Width           =   3060
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
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
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
         Top             =   630
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid GRID2 
         Height          =   765
         Left            =   240
         TabIndex        =   69
         Top             =   1920
         Visible         =   0   'False
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   1349
         _Version        =   393216
         Rows            =   23
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   8433840
         BackColorSel    =   -2147483643
         ForeColorSel    =   12582912
         BackColorBkg    =   16777215
         Redraw          =   -1  'True
         FocusRect       =   2
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid GRID1 
         Height          =   765
         Left            =   3000
         TabIndex        =   70
         Top             =   1920
         Visible         =   0   'False
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   1349
         _Version        =   393216
         Rows            =   23
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   8433840
         BackColorSel    =   -2147483643
         ForeColorSel    =   12582912
         BackColorBkg    =   16777215
         Redraw          =   -1  'True
         FocusRect       =   2
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
         Height          =   435
         Index           =   0
         Left            =   60
         OleObjectBlob   =   "RECIBO08.frx":3C08
         TabIndex        =   99
         Top             =   0
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label35 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   100
         Top             =   3960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label34 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   7875
         TabIndex        =   65
         Top             =   210
         Visible         =   0   'False
         Width           =   540
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
         TabIndex        =   45
         Top             =   3885
         Width           =   1485
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ajuste de Centavos:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3885
         TabIndex        =   44
         ToolTipText     =   "Doble Click Para Ajuste Automatico"
         Top             =   3885
         Width           =   1905
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Valores:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7455
         TabIndex        =   43
         Top             =   3945
         Width           =   1275
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
         TabIndex        =   39
         ToolTipText     =   "Doble Click para Configurar Cliente"
         Top             =   6300
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label MONEMIS 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   38
         ToolTipText     =   "T.C.: 1.0000 - Doble-Click para Cambiar Moneda"
         Top             =   6810
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Height          =   120
         Left            =   3360
         TabIndex        =   37
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
         Left            =   10815
         TabIndex        =   36
         ToolTipText     =   "Doble Click para Ajustar Tipo de Cambio"
         Top             =   6500
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label26 
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   210
         TabIndex        =   35
         Top             =   1000
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Docum. no Válido como Factura"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4515
         TabIndex        =   15
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   1005
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
         X2              =   10650
         Y1              =   3810
         Y2              =   3810
      End
      Begin VB.Line Line27 
         X1              =   10665
         X2              =   10665
         Y1              =   1480
         Y2              =   3810
      End
      Begin VB.Line Line26 
         X1              =   10665
         X2              =   10665
         Y1              =   4305
         Y2              =   7090
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8505
         TabIndex        =   21
         Top             =   630
         Visible         =   0   'False
         Width           =   2000
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7875
         TabIndex        =   20
         Top             =   705
         Visible         =   0   'False
         Width           =   540
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
         TabIndex        =   19
         Top             =   5775
         Visible         =   0   'False
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
         TabIndex        =   18
         Top             =   5880
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Numero:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7920
         TabIndex        =   13
         Top             =   1065
         Width           =   1020
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
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
         ToolTipText     =   "Doble - Click para visualzar aplicación de Créditos y Anticipos"
         Top             =   315
         Width           =   960
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "RECIBO08.frx":3C82
      Top             =   1680
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
      Visible         =   0   'False
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
      Begin VB.Menu mnuMaeCli 
         Caption         =   "Maestro de Clientes"
      End
      Begin VB.Menu mnuNoCli 
         Caption         =   "Notas de Clientes"
      End
      Begin VB.Menu php1 
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
      Begin VB.Menu mnuVerReimp 
         Caption         =   "Ver Recibos y Re-Imprimir"
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
   End
   Begin VB.Menu mnuConsuList 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuResCuCo 
         Caption         =   "Resumen de Cuenta Corriente"
      End
      Begin VB.Menu mnuLiCuCoes 
         Caption         =   "Listado de Cuentas Corrientes"
      End
      Begin VB.Menu MnuApliCredAnt 
         Caption         =   "Ver Aplicación de Créditos y Anticipos"
      End
      Begin VB.Menu php4a 
         Caption         =   "-"
      End
      Begin VB.Menu LiReReg 
         Caption         =   "Listado de Recibos Registrados"
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
         Caption         =   "Configuración General"
      End
   End
End
Attribute VB_Name = "RECIBO08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$, AUTONUME%
Dim MODORECI%
Dim MODOMONEI%
Dim ctrl As Object
Sub CARGA_ENCABEZADO_GRILLA()

    INFO_GRILLA07.MSF.Row = 1
    INFO_GRILLA07.MSF.Cols = 20
    NUCOLU% = 1: INFO_GRILLA07.TEPRUEBA = "1234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Cliente"
    NUCOLU% = 2: INFO_GRILLA07.TEPRUEBA = "1234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Cond.Pago"
    NUCOLU% = 3: INFO_GRILLA07.TEPRUEBA = "1234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Plazo"
    '
    NUCOLU% = 4: INFO_GRILLA07.TEPRUEBA = "1234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Semana1"
    NUCOLU% = 5: INFO_GRILLA07.TEPRUEBA = "1234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Importe1"
    NUCOLU% = 6: INFO_GRILLA07.TEPRUEBA = "1234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Promedio1"
    '
    NUCOLU% = 7: INFO_GRILLA07.TEPRUEBA = "1234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Semana2"
    NUCOLU% = 8: INFO_GRILLA07.TEPRUEBA = "1234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Importe2"
    NUCOLU% = 9: INFO_GRILLA07.TEPRUEBA = "1234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Promedio2"
    '
    NUCOLU% = 10: INFO_GRILLA07.TEPRUEBA = "1234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Semana3"
    NUCOLU% = 11: INFO_GRILLA07.TEPRUEBA = "1234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Importe3"
    NUCOLU% = 12: INFO_GRILLA07.TEPRUEBA = "1234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Promedio3"
    '
    NUCOLU% = 13: INFO_GRILLA07.TEPRUEBA = "1234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Semana4"
    NUCOLU% = 14: INFO_GRILLA07.TEPRUEBA = "1234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Importe4"
    NUCOLU% = 15: INFO_GRILLA07.TEPRUEBA = "1234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Promedio4"
    '
    NUCOLU% = 16: INFO_GRILLA07.TEPRUEBA = "1234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Semana5"
    NUCOLU% = 17: INFO_GRILLA07.TEPRUEBA = "1234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Importe5"
    NUCOLU% = 18: INFO_GRILLA07.TEPRUEBA = "1234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Promedio5"
    '
    NUCOLU% = 19: INFO_GRILLA07.TEPRUEBA = "1234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Num / $$"
    NUCOLU% = 20: INFO_GRILLA07.TEPRUEBA = "1234567890": INFO_GRILLA07.MSF.ColWidth(NUCOLU% - 1) = INFO_GRILLA07.TEPRUEBA.Width: INFO_GRILLA07.MSF.ColAlignment(NUCOLU% - 1) = 6: INFO_GRILLA07.MSF.TextMatrix(0, NUCOLU% - 1) = "Prome_Sem"
    '
End Sub

'
Sub ClearControls()
    For Each ctrl In Me.Controls
        If ctrl.Tag = "TBOX" Then ctrl.TEXT = ""
    Next
    Label26 = ""
    Label22 = ""
    
End Sub

Sub DIASDECOBRANZA_VENCIFACTURA()
        ' CALCULO DE DIAS DE COBRANZA A PARTIR DE LA FECHA DE VENCIMIENTO
        ' DE LAS FACTURAS CANCELADAS
        '
        ' CALCULO DE DIAS DE COBRANZA
        JJ% = 0: ACRED# = 0: TOTALVAL# = 0
        For II% = 0 To List1.ListCount - 1
             FVENCHE% = FECHANUM(Mid$(List1.List(II%), 62, 8))   ' fecha de vencimiento del cheque
             IMPOCHE# = XVALO(Mid$(List1.List(II%), 70, 14))   ' importe del cheque
             
5            If ACRED# < 0.005 Then     ' SI HAY IMPORTE DE FACTURA PENDIENTE
10               If JJ% < List2.ListCount Then ' RECORRE LOS COMPROBANTES DE LA COMPOSICION DE DEUDA
                     REGFACTURA$ = List2.List(JJ%)
                     FVENFAC% = FECHANUM(Mid$(REGFACTURA$, 47, 8))
                     ACRED# = XVALO(Mid$(REGFACTURA$, 70, 14))
                     If Abs(ACRED#) < 0.005 Then    ' VERIFICA SI EL COMPROBANTE FUE ELEGIDO
                         JJ% = JJ% + 1
                         GoTo 10
                     End If
                 Else ' SI NO HAY MAS COMPROBANTES AGREGA EL SALDO DEL ULTIMO CHEQUE CON PONDERACION 1 (UNO)
                     DIFEDIA% = 1
                     TOTALVAL# = TOTALVAL# + (DIFEDIA% * IMPOCHE#)
                     Exit For
                 End If
             End If
                 
             If IMPOCHE# >= ACRED# Then ' EL IMPORTE DEL CHEQUE ES MAYOR AL DE LA FACTURA
                 IMPOCHE# = IMPOCHE# - ACRED#
                 '
                 DIFEDIA% = DIENTRE(FVENFAC%, FVENCHE%)
                 If ACRED# < 0 Then DIFEDIA% = 1 ' SI EL COMPROBANTE ES UN CREDITO, DIFEDIA% = 1
                 '
                 TOTALVAL# = TOTALVAL# + (DIFEDIA% * ACRED#)
                 ACRED# = 0: JJ% = JJ% + 1
             Else                       ' EL IMPORTE DE LA FACTURA ES MAYOR AL DEL CHEQUE
                 DIFEDIA% = DIENTRE(FVENFAC%, FVENCHE%)
                 TOTALVAL# = TOTALVAL# + (DIFEDIA% * IMPOCHE#)
                 ACRED# = ACRED# - IMPOCHE#
                 IMPOCHE# = 0
             End If
             
             If IMPOCHE# > 0.005 Then GoTo 5 ' SI TODAVIA HAY IMPORTE DEL CHEQUE, BUSCA UN NUEVO COMPROBANTE
        Next II%
        '
        Label35 = FORMATNUM(0, "I1")
        TOTAVALO# = XVALO(Label14) ' TOTAL DE VALORES RECIBIDOS
        If TOTAVALO# > 0.005 Then Label35 = FORMATNUM((TOTALVAL# / TOTAVALO#), "I5") ' CALCULO DE DIAS PONDERADO
        '
End Sub

Function PROMEDIO_CONDI_PAGOS(CONDIPAG%)
      '
      PROMEDIO_CONDI_PAGOS = 0
      CANTI_CONDICIONES% = 0
      '
      'LEO EL REGISTRO SEGUN CONDICION DE PAGO Y TOMO LAS CONDICIONES
      'SUMO CUANTAS CONDICIONES TIENE PARA LUEGO  PODER CALCULAR EL PROMEDIO
      '
      REGISTRO1$ = FILTERGETRECORD$("CONPAG", 1, MKI$(CONDIPAG%))
      P1% = Asc(Mid$(X$, 38, 1)): If P1% > 0 Then CANTI_CONDICIONES% = CANTI_CONDICIONES% + 1
      P2% = Asc(Mid$(X$, 39, 1)): If P2% > 0 Then CANTI_CONDICIONES% = CANTI_CONDICIONES% + 1
      P3% = Asc(Mid$(X$, 40, 1)): If P3% > 0 Then CANTI_CONDICIONES% = CANTI_CONDICIONES% + 1
      P4% = Asc(Mid$(X$, 41, 1)): If P4% > 0 Then CANTI_CONDICIONES% = CANTI_CONDICIONES% + 1
      P5% = Asc(Mid$(X$, 42, 1)): If P5% > 0 Then CANTI_CONDICIONES% = CANTI_CONDICIONES% + 1
      P6% = Asc(Mid$(X$, 43, 1)): If P6% > 0 Then CANTI_CONDICIONES% = CANTI_CONDICIONES% + 1
      '
      'SUMO LOS PLAZOS DE PAGO
      SUMACONDI = P1% + P2% + P3% + P4% + P5% + P6%
      '
      'SI NO TIENE PLAZOS DEVUELVE 0 (CONTADO)
      If SUMACONDI <= 0 Then Exit Function
      '
      If CANTI_CONDICIONES% < 1 Then CANTI_CONDICIONES% = 1
      PROMEDIO_SUMACONDI = SUMACONDI / CANTI_CONDICIONES%
      '
End Function

Private Sub BOTNEXT_Click()
    '
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
   Command10.Enabled = False
   '
   Call SELMESA(DESDE%, HASTA%, CORRESP$)
'   CORRESP$ = MESANO$(MESAN$, DESDE%, HASTA%)
   If CORRESP$ = "" Then GoTo 99
   '
   Call CARGA_ENCABEZADO_GRILLA
   '
   'VEO COMO OBTENGO LAS SEMANAS
   
   PRIMER_DIA_MES% = DESDE%: PRIMER_DIA_SEMANA% = DESDE%
   '
   'OBTENGO LOS DATOS DE LA PRIMER SEMANA CONTANDO A PARTIR DEL PRIMER DIA
   '***********************************************************************
   ULTIMO_DIA_MES% = HASTA%
   FEX = PRIMER_DIA_MES%
     For I% = PRIMER_DIA_MES% To (PRIMER_DIA_MES% + 7)
       FEX = DIPOST(FEX)
       DIATEXTO$ = DSEM$(FEX)
       If DIATEXTO$ = "DOMINGO" Then
          ULTIMO_DIA_SEMANA% = FEX
          Exit For
       End If
     Next I%
DIAA = DSEM$(FEX)
   'FIN -- OBTENGO LOS DATOS DE LA PRIMER SEMANA CONTANDO A PARTIR DEL PRIMER DIA
   '******************************************************************************
   '
   'HAGO UNA CONSULTA SOBRE CAJABANC DE LOS CLIENTES QUE TIENEN RECIBOS EN EL RANGO SELECCIONADO
   '******************************************************
   
   SQLX$ = "SELECT SUM(IDEBECOMP) AS IMPODEBE, SUM(IHABECOMP)AS IMPOHABE, NUCLIEN, COUNT(*) AS CANTREG FROM CAJABANC WITH (NOLOCK) "
   SQLX$ = SQLX$ + " WHERE TiMovi$ = 'RCOB' "
   SQLX$ = SQLX$ + " GROUP BY  NUCLIEN "
   SQLX$ = SQLX$ + " ORDER BY NUCLIEN "
   
   DIAX$ = DSEM$(DESDE%)
   '
   JJ& = 0: VUELTA% = 0
   '
   Set RSBANC = READSET(SQLX$)
   With RSBANC
    Do While Not .EOF
      REG& = REG& + 1
      If VUELTA% > 0 Then FIN& = XVALO(!CANTREG): VUELTA% = 1
      '
      NC% = XVALO(!NUCLIEN)
      CONDICION_DE_PAGO% = CPAGCLI%(NC%)
      PROMEDIO_DE_PAGOS = PROMEDIO_CONDI_PAGOS(CONDICION_DE_PAGO%) 'DEVUELVE EL PROMEDIO DE LAS CONDICIONES DE PAGO
      REFECONPAGO$ = ECOARCH$("CONPAG", MKI$(CONDICION_DE_PAGO%))
      
      '
      INFO_GRILLA07.MSF.Row = INFO_GRILLA07.MSF.Row + 1
      Call TRACE(24, REG&, FIN&)
      '
      INFO_GRILLA07.MSF.TextMatrix(REG&, 1) = RASOCLI$(NC%)
      INFO_GRILLA07.MSF.TextMatrix(REG&, 2) = REFECONPAGO$
      INFO_GRILLA07.MSF.TextMatrix(REG&, 3) = PROMEDIO_DE_PAGOS
      '
      CONDI$ = "TiMovi$ = 'RCOB' AND NUCLIEN = " & NC%
      '
      ACUMULAX = 0
      For I% = 1 To 15 Step 3
         DIAX$ = DSEM$(PRIMER_DIA_SEMANA%)
         If DIAX$ <> "DOMINGO" Then
           If PRIMER_DIA_SEMANA% > HASTA% Then Exit For 'SI EL PRIMER DIA DE LA SEMANA ES MAYOR TERMINA
           If ULTIMO_DIA_SEMANA% > HASTA% Then ULTIMO_DIA_SEMANA% = HASTA% ' SI EL ULTMO ES MAYOR A FIN DE MES ASIGNO FIN DE MES
           CONDI2$ = CONDI$ + " AND FECHEMISI >= PRIMER_DIA_SEMANA% AND FECHEMISI <= ULTIMO_DIA_SEMANA%"
           IMPORTEX# = XVALO(VALOBADA("CAJABANC", "Sum(IHABECOMP)", CONDI2$, "NOMESS"))
           PROMEDIO_ATRAZO1 = XVALO(VALOBADA("CAJABANC", "AVG(AtraCheque)", CONDI2$, "NOMESS"))
           INFO_GRILLA07.MSF.TextMatrix(REG&, I% + 3) = PROMEDIO_ATRAZO1
           INFO_GRILLA07.MSF.TextMatrix(REG&, I% + 5) = FORMATNUM$((PROMEDIO_ATRAZO1 * IMPORTEX#), "F12.2")
           ACUMULAX = ACUMULAX + (PROMEDIO_ATRAZO1 * IMPORTEX#)
         End If
         PRIMER_DIA_SEMANA% = ULTIMO_DIA_SEMANA% + 1
         ULTIMO_DIA_SEMANA% = ULTIMO_DIA_SEMANA% + 7
      Next
      '
      INFO_GRILLA07.MSF.TextMatrix(REG&, 5) = FORMATNUM$(IMPORTEX#, "F14.2")
      INFO_GRILLA07.MSF.TextMatrix(REG&, 19) = FORMATNUM$(ACUMULAX, "F14.2")
      '
      .MoveNext
    Loop
    '
   End With
   RSBANC.Close
   Set RSBANC = Nothing
   INFO_GRILLA07.Show 1
   '
99 Command10.Enabled = True
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("ARCHIG07/AMACLI", "Archivos Maestros")
99  Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   For U& = 1 To ULTREG&("MECOBRA")
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
          Label14.Caption = ""
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
    Call REMPLAGRID(Me, GRID1, Picture1, Label12, List1)
    Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
End Sub

Private Sub Command23_Click()
   Command23.Enabled = False
   Call RECICO07.Command23_Click
   Command23.Enabled = True
End Sub

Private Sub Command24_Click()
   Command24.Enabled = False
   'llamada a busqueda de cheques por numero
   Call GESVAL07.BUSCACHEQ
   Command24.Enabled = True

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
    RECUEC07.Text1.TEXT = Text1.TEXT
    RECUEC07.Show 1
    Command4.Enabled = True
End Sub

Function VALIDARECIBO%()
    '
    Label34 = 1: Label34.Refresh
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
    IREDON# = XVALO(Text8)
    If Abs(IREDON#) >= 0.005 Then
       CUERED% = 0
       CUEREDO$ = TRIM$(CONTROL$("RECISTA", "CUEREDON"))
       If CUEREDO$ <> "" Then CUERED% = CUEINT%(CUEREDO$)
       If CUERED% < 1 Then
          Call MENSERR(24, "Falta Cuenta Contable Redondeo.")
          VDRX% = (-1)
          GoTo 99
       End If
    End If
    '
    ' VERIFICA QUE EL RECIBO POSEA CHEQUES
    If XVALO(Label14.Caption) < 0.01 Then
        Call COMUNI("Imposible Grabar - Falta Detalle de Valores")
        VDRX% = (-1)
        GoTo 99
    End If
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
    APRORECIB.Label21 = CODCUE$(CUECONDE%)
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
    'JANDY SQL
    
    FENUE$ = CONTROL$("", "FENURECI")
    If FENUE$ = "" Then FENUE$ = "01/01/80"
    FEXI% = FECHANUM(FENUE$)
    'FLIMI = CVDAT(FECHANUM(FENUE$))
    FLIMI = FETEXCOR1$(FEXI%)
    
    'If FLIMI > Now Then FLIMI = CVDAT(HOY(HO$))
    'FLIMI1 = CDbl(FLIMI)
     
 
    '
    SQLX$ = "SELECT NumeComp FROM CajaBanc WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + "AND (CodiCom_Cor = '" & RECICOR$ & "' "
    SQLX$ = SQLX$ + "OR CodiCom_Lar = '" & RECILAR$ & "') "
    SQLX$ = SQLX$ + "AND Status < 9 "
    'SQLX$ = SQLX$ + "AND CDBL(FechEmisi) >= " & FLIMI1 & " "
    SQLX$ = SQLX$ + "AND FechEmisi >= '" & FLIMI & "' "
    
    Dim OrpaTemp As ADODB.Recordset
    
    'Set OrpaTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    'Set OrpaTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
    'se modifica por que la rutina readset captura el error de tiempo de espera
    Set OrpaTemp = READSET(SQLX$)
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
    OrpaTemp.Close
    Set OrpaTemp = Nothing
    '
    If NEFA7# >= 0.005 Then
      If FORFASQL.VERINUFA%(NFCY$) < 1 Then
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
    Label34 = 4: Label34.Refresh
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
    Label14.Caption = ""
    Label18.Caption = ""
    Label20.Caption = ""
    '\\\OT-07-0029-RG-26/01/07///
    DESCUEN(0) = ""
    Text7 = ""
    Text8 = ""
    Call LIMPPROM
    Frame8.Caption = "Descuento (N./Cred)"
    '\\\OT-07-0029-RG-FIN///
    BOTNEXT.Enabled = True
    Command5.Enabled = False
    FF% = HOY(HO$)
    Text9.TEXT = HO$
    '
    NROPRO& = 0
    Call NURECIBO(NROPRO&)
    Call PRENURE(NROPRO&)
    '
    Text11 = "RB.      "
    Text11.SelStart = 3
    Text11.SelLength = Len(Text11.TEXT) - 3
    Text11.SetFocus
    
    If EsPorPedido& > 0 Then
      Command2_Click ' CERRAR LA VENTANA
    End If
    '
    Call REMPLAGRID(Me, GRID1, Picture1, Label12, List1)
    Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
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
    If MODONUM$ = "AUTONUME" Then TNUME$ = TRIM$(TNUME$) + TRIM$(Str$(NROPRO&))
    Text11.TEXT = TNUME$
    '
End Sub
'
Sub Command6_Click()
    Command6.Enabled = False
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then
       OPRECIB07.Show 1
       Call COMUNI("Reconfiguracion Completada\   \Re-Inicie Ahora la Aplicacion.")
       Call FINAL("")
    End If
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

Sub Command8_Click()
   Command8.Enabled = False
   RECICO07.Show 1
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

Private Sub DESCUEN_Change(Index As Integer)
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

Private Sub DESCUEN_GotFocus(Index As Integer)
    If TRIM$(Text1) = "" Then Text1.SetFocus
    NCX1 = XVALO(Text1)
    If NCX1 < 1 Or NCX1 > 32767 Then Text1.SetFocus
    NCX2% = NCX1
    If RASOCLI$(NCX2%) = "" Then Text1.SetFocus
    DESCUEN(Index) = TRIM$(DESCUEN(Index))
End Sub

Private Sub DESCUEN_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
      Call DESCUEN_LostFocus(Index)
    End If
End Sub

Private Sub DESCUEN_LostFocus(Index As Integer)
    FDESCU$ = TRIM$(FORMATNUM(XVALO(DESCUEN(Index)), "F12.2"))
    PDX = 0
    IDES# = XVALO(DESCUEN(0))
    L14# = XVALO(Label14)
    If L14# >= 1 Then
      PDX = 100 - 100 / (IDES# / L14# + 1)
    End If
    Text7 = TRIM$(FORMATNUM$(PDX, "F6.1"))
    DESCUEN(Index) = FDESCU$
End Sub

Private Sub Form_Load()
    '
    Dim BUFLE As String * 64
    '
    Call CENTRAFORM(Me)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    Call APLICACIONSKINS(Me, Picture3)

    Frame1.Refresh
    DoEvents
    '\\\OT-07-0029-RG-26/01/07///
    If CONTROL$("RECISTA", "DESCAJA") = "CREDESCU" Then
       'Frame9.Left = 5930
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
    ElseIf MODOMONE$ = "BIMONETA" Then
       MODOMONEI% = 1
    End If
    '
10  FORIPRE$ = TRIM$(CONTROL$("", "FORECIST"))
    If FORIPRE$ <> "" Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
        RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
        Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Recibo Standard de Cobranza:\  \'" + RMCC$ + "'.")
        If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
        Call CIERRARCH("*.*")
        Call FINAL("")
      End If
    End If
    '
    CUEDEUVE$ = TRIM$(CONTROL$("AMACLI", "ECUECON"))
    If CUEDEUVE$ = "" Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Correspondiente a 'Deudores por Ventas'.")
      If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
      Call CIERRARCH("*.*")
      Call FINAL("")
    End If
    If ECOARCH$("ECUECON", CUEDEUVE$) = "" Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Correspondiente a 'Deudores por Ventas'.")
      If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
      Call CIERRARCH("*.*")
      Call FINAL("")
    End If
    '
    CUECHECA$ = TRIM$(CONTROL$("", "CUECHECA"))
    If CUECHECA$ = "" Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a la Cartera de Valores.")
      If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
      Call CIERRARCH("*.*")
      Call FINAL("")
    End If
    If ECOARCH$("ECUECON", CUECHECA$) = "" Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a la Cartera de Valores.")
      If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
      Call CIERRARCH("*.*")
      Call FINAL("")
    End If
    '
    If Frame8.Visible = True Then
       CUEDESCU$ = TRIM$(CONTROL$("", "CUEDESCU"))
       If CUEDESCU$ = "" Then
         Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Descuentos de Caja.")
         If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
         Call CIERRARCH("*.*")
         Call FINAL("")
       End If
       If ECOARCH$("ECUECON", CUEDESCU$) = "" Then
         Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Descuentos de Caja.")
         If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
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
         If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
         Call CIERRARCH("*.*")
         Call FINAL("")
       ElseIf OPX% = 2 Then
         If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
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
       RECIBO08.Label22.Caption = Left$(TRIM$(Asc(Left$(XXX$, 1))) + " - " + Mid$(XXX$, 2, 30), 19)
    End If
    '
    AUTONUME% = 0
    If CONTROL$("", "NURECIST") = "AUTONUME" Then
      AUTONUME% = 1
    End If
    '
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
'    Refresh
    '
    Call REMPLAGRID(Me, GRID1, Picture1, Label12, List1)
    Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
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
   '
   If EsPorPedido& > 0 Then Unload Me: Exit Sub
   '
   Hide
End Sub

Private Sub GRID1_DblClick()
    If GRID2.Row > 0 Then
       On Error Resume Next
       List1.ListIndex = GRID2.Row - 1
       Call List1_DblClick
       On Error GoTo 0
    End If
End Sub


Private Sub GRID2_DblClick()
    If GRID2.Row > 0 Then
       On Error Resume Next
       List2.ListIndex = GRID2.Row - 1
       Call List2_Click
       On Error GoTo 0
    End If
End Sub

Private Sub Label1_DblClick()
   '\\\OT-07-0189-OD-24/04/08///
   If TRIM$(Text1) = "" Or TRIM$(Text2) = "" Then
     GoTo 99
   End If
   NC% = XVALO(Text1)
   Call RECICO07.VAPLIREC(NC%)
99
  '\\\OT-07-0189-OD-FIN///
End Sub

Private Sub Label10_Change()
   '\\\OT-08-0203-OD-29/04/08///
   PROVENCI# = 0
   FIN = List2.ListCount - 1
   For U& = 0 To FIN
     Z$ = List2.List(U&)
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
   Call REMPLAGRID(Me, GRID1, Picture1, Label12, List1)
   Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
   '\\\OT-08-0203-OD-FIN///
   '
   ' CALCULO DE DIAS DE COBRANZA A PARTIR DE LA FECHA DE VENCIMIENTO
   ' DE LAS FACTURAS CANCELADAS
   If CONTROL$("RECISTA", "DIASVEN") = "FECHAFA" Then
        Call DIASDECOBRANZA_VENCIFACTURA
   End If
   '
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
   If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
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

Private Sub Label35_Change()
    Label19 = Label35
End Sub
Private Sub Label8_Change()
    Call Timer1_Timer
End Sub

Private Sub LiReReg_Click()
    Call RECICO07.Command8_Click
End Sub

Private Sub List1_DblClick()
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
    Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
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
          Label14.Caption = ""
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
    Call REMPLAGRID(Me, GRID1, Picture1, Label12, List1)
End Sub

Private Sub mnuAjuSalDeu_Click()
    Call RECICO07.Command4_Click
End Sub

Private Sub mnuAnuRec_Click()
    Call RECICO07.Command25_Click
End Sub

Private Sub MnuApliCredAnt_Click()
    '\\\OT-07-0189-OD-24/04/08///
    Call Label1_DblClick
    '\\\OT-07-0189-OD-FIN///
End Sub

Private Sub mnuConfiGral_Click()
    Call Command6_Click
End Sub

Private Sub mnuImpCre_Click()
    Call RECICO07.Command3_Click
End Sub

Private Sub mnuLiCuCoes_Click()
    Call RECUEC07.Command3_Click
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
    Call RECICO00.Command5_Click
End Sub

Private Sub mnuSalir_Click()
    Call Command2_Click
End Sub

Private Sub mnuVerReimp_Click()
    Call RECICO07.Command23_Click
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
   Call REMPLAGRID(Me, GRID1, Picture1, Label12, List1)
   Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
   Label4.Caption = ""
   Label8.Caption = ""
   Label10.Caption = ""
   Label14.Caption = ""
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

Private Sub Text1_Change()
   'jandy SQL
    Call ClearControls

    List1.Clear
    List2.Clear
    Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
    If Picture1.Visible = True Then Call REMPLAGRID(Me, GRID1, Picture1, Label12, List1)
    Label4.Caption = ""
    Label8.Caption = ""
    Label10.Caption = ""
    Label14.Caption = ""
    Label18.Caption = ""
    Label20.Caption = ""
    '\\\OT-07-0029-RG-26/01/07///
    DESCUEN(0) = ""
    Text7 = ""
    Text8 = ""
    Call LIMPPROM
    Frame8.Caption = "Descuento (N./Cred)"
    '\\\OT-07-0029-RG-FIN///
    BOTNEXT.Enabled = True
    Command5.Enabled = False
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

             If RASOCLI$(NC%) <> "" Then
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
   If DERACCE%("STPRECIB", "Configuración de Funcionamiento Recibos") > 1 Then OPRECIB07.Show 1
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
      PORDEX = Val(Text7): If PORDEX > 99 Then PORDEX = 99
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
    If RASOCLI$(NCX2%) = "" Then Text1.SetFocus
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
    Label34 = 2: Label34.Refresh
    'ver ernesto fechas
    Static FLIMI, FENUE$, AUTONURE%
    '
    Screen.MousePointer = 11
    '
    If FENUE$ = "" Then
      FENUE$ = CONTROL$("", "FENURECI")
      If FENUE$ = "" Then FENUE$ = "01/01/80"
      FEXI% = FECHANUM(FENUE$)
      'FLIMI = CVDAT(FECHANUM(FENUE$))
      FLIMI = FETEXCOR1$(FEXI%)
      
      'If FLIMI > Now Then FLIMI = CVDAT(HOY(HO$))
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
    'FLIMI1 = CDbl(FLIMI)
    
    COBRACT% = XVALO(Label22)
    Call APERBASDAT("CABAN")
    '
    SQLX$ = "SELECT max(NumeComp) as maximo FROM CajaBanc WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + " AND CodiCom_Cor LIKE 'RB%' "
    SQLX$ = SQLX$ + " AND FechEmisi >= '" & FLIMI & "'"
    SQLX$ = SQLX$ + " AND NumeTalo = " & COBRACT%
    SQLX$ = SQLX$ + " AND Status < 9 "
    Dim MIRS As ADODB.Recordset
    'MODIFCADO PARA CAPTURAR ERROR DE TIEMPO DE ESPERA DENTRO DE FUNCION READSET
    'Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
    Set MIRS = READSET(SQLX$)
    If Not IsNull(MIRS!maximo) Then
        NROMA& = MIRS!maximo
    End If
    MIRS.Close
    Set MIRS = Nothing

    If NROPRO& < 1 Then NROPRO& = 1 + NROMA&
    Screen.MousePointer = 1
    Label34 = 3: Label34.Refresh
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
    If RASOCLI$(NC%) = "" Then
        Call MENSERR(24, "Cuenta no Figura\en Base de Datos de Clientes")
        Text1.SetFocus
        Exit Sub
    End If
    '
    NCX1X% = Int(GRUDIVCLI(NC))
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
    If DIENTRE%(FF%, HOI%) > 30 Then
      EPAC$ = TRIM$(UCase$(EMPREAC$))
      If InStr(EPAC$, "NEUMAN") < 1 Or XVALO(Left$(TRIM$(Label22.Caption), 2)) <> 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text9.SetFocus
        Exit Sub
      End If
    End If
   '
    Call CIERRARCH("*.*")
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
    'Jandy SQL
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM SADEUPEN "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NC% & " "
    SQLX$ = SQLX$ + "ORDER BY FeVencim ASC;"
    
    'Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    Dim FacPenTemp As ADODB.Recordset
    Set FacPenTemp = New ADODB.Recordset
    FacPenTemp.CursorLocation = adUseClient
25  On Error Resume Next
    FacPenTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    With FacPenTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        On Error GoTo 0
        .MoveFirst
        Do While Not .EOF
          TIX$ = UCase$(Left$(!codicom_lar, 2))
          X$ = !codicom_lar
          SG% = (-1)
          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                              If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                              If TIX$ = "NC" Then TIPOS$ = "Nota de Credito": SG% = 1
                              If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1 '\\\OT-06-0112-WAR-30/03/06///
                              If TIX$ = "RB" Then TIPOS$ = "Anticipo": SG% = 1 '\\\OT-06-0112-WAR-30/03/06///
                              '
          If TIX$ = "FC" Or TIX$ = "SC" Or TIX$ = "ND" Or TIX$ = "AS" Or TIX$ = "RB" Or TIX$ = "NC" Then
            SALDO# = !salddebe - !saldacre
            If MODOMONEI% = 2 Then
               SALDOLA# = !SaldDebe_Dola - !SaldAcre_Dola
               MONEDOCU% = !Mone_Emis
               If MONEMI% > 1 Then
                 If MONEMI% <> MONEDOCU% Then GoTo 95
                 SALDO# = SALDOLA#
               End If
            End If
            '
            'If SALDO# >= 0.005 Then
            If Abs(SALDO#) >= 0.005 Then
              IMPO# = (-1) * !ibrucomp * SG%
              SUMAFAC# = SUMAFAC# + IMPO#
              IMPO# = Abs(IMPO#)
              '
              NROX$ = !codicom_lar
              FEMI = CVINT(!fechemisi)
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
      On Error GoTo 0
    End With
    FacPenTemp.Close
    Set FacPenTemp = Nothing
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
    Call REMPLAGRID(Me, GRID1, Picture1, Label12, List1)
    Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
    Call SIGUERECI2
    '
End Sub
'

Sub SIGUERECI2()
    '
    HOI% = HOY(HO$)
    NC% = XVALO(Text1)
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
      X$ = REGLEIDO$("!CARMECO", U&)
      PROTER$ = UCase$(TRIM$(Mid$(X$, 43, 2)))
      If PROTER$ = "" Or Left$(PROTER$, 1) = "P" Then
           PROTER$ = "  "
         Else
           PROTER$ = "3 "
      End If
      If TRIM$(PROTER$) <> "" Then ATXY% = 1
    Next U&
    '
210 If ATXY% > 0 Then
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
    FALTACUIT% = 0
    '
    ' SI ESTA CONFIGURADO PARA QUE CALCULE LOS DIAS DE COBRANZA
    ' A PARTIR DEL VENCIMIENTO DE LA FACTURA, ORDENA LOS VALORES
    ' POR SU VENCIMIENTO
    If CONTROL$("RECISTA", "DIASVEN") = "FECHAFA" Then
        Call COPYSTRU("CARMECO", "CARMECF")
        Call TRACENTRAL("CARMECF", "CARMECO")
        Call FILESORT("CARMECF", "", 7, 7, "ASC")
        Call TRALOCAL("CARMECF", "CARMECO")
    End If
    '
    For U& = 1 To ULTREG&("!CARMECO")
        X$ = REGLEIDO$("!CARMECO", U&)
        If X$ = REBLA$ Then GoTo 290
        MECO% = CVI(Left$(X$, 2))
        BANCO$ = TRIM$(ECOARCH$("MECOBRA", MKI$(MECO%)))
        '
        If BANCO$ = "" Then
            Call MENSERR(24, "Medio de Pago no Válido")
            GoTo 200
        End If
        '
        '\\\OT-08-0622-OD-12/08/08///
        'MEDIOCOBRAN$ = ECOARCH$("MECOBRA", MKI$(MECO%))
        'Valida que no este marcado el medio de cobranza como no Vigiente (N/V)
        If InStr(1, BANCO$, "N/V") > 0 Then
           Call MENSERR(24, "Medio de Pago " + "'" + BANCO$ + "'" + "  No Vigente")
           GoTo 200
        End If
        '\\\OT-08-0622-OD-FIN///
        '
        NUVA$ = Mid$(X$, 19, 12)
        '
        'VERIFICA QUE SEA CHEQUE
        CUECO% = CUEMECOB%(MECO%)
        If CUECO% <> CUEICARCH% Then GoTo 19
        
        'VERIFICA SI EL CHEQUE  FUE INGRESADO
        BANCO1$ = Mid$(BANCO$, 4, 18)
        VeriChek% = YAESTACHEQUE%(BANCO1$, NUVA$)
        If VeriChek% > 0 Then
          TX1$ = "Cheque N.: " + NUVA$ + "  Banco : " + BANCO1$ + " Ya Ingresado"
          TX1$ = TX1$ + "\Ante Cualquier Duda Utilice la Opción\Consulta de Cheques 'B.Ch.'"
          Call MENSERR(24, TX1$)
          GoTo 200
        End If
        '
        'Verifica si el cheque esta repetido
        For I& = 1 To (U& - 1)
          X1$ = REGLEIDO$("!CARMECO", I&)
          If Mid$(X1$, 1, 32) = Mid$(X$, 1, 32) Then
            Call MENSERR(24, "Cheque N.: " + NUVA$ + "  Banco : " + BANCO1$ + " Repetido")
            GoTo 200
          End If
        Next I&
        '
        'BANDERA VALIDACION DE INGRESO DE CUIT
        CUITLIB$ = TRIM$(Mid$(X$, 47, 16))
        If CUITLIB$ = "" Then FALTACUIT% = 1
        '
19      IMVA# = CVD(Mid$(X$, 31, 8))
        IMVABS# = IMVA#: If UCase$(BANCO$) = "EFECTIVO" Then IMVABS# = Abs(IMVA#)
        ' LO ANTERIOR ES PARA ACEPTAR EFECTIVO EN NEGATIVO
        If IMVABS# < 0.005 Then
            Call MENSERR(24, "Falta Importe")
            GoTo 200
        End If
        '
        ' Control de fecha de acreditacion para TODOS ... no solo para cheques
        FEMI = CVI(Mid$(X$, 39, 2))
        FEVE = CVI(Mid$(X$, 41, 2))
        '
        ' ESTO ES PARA EL CASO DE QUE NO SE INGRESA FECHA DE EMISIÓN Y
        ' SI DE VENCIMIENTO - PEDIDO POR ANTONIO DE TANIT EL 16/05/2007
        If Abs(FEMI) < 0.5 And FEVE > 0.5 Then
           FEMII% = FEVE
           If FEMII% > HOI% Then FEMII% = HOI%
           Call REPLA(X$, MKI$(FEMII%), 39, 2)
           Call GRAREG("!CARMECO", X$, U&)
           FEMI = FEMII%
        End If
        '
        If Abs(FEMI) < 0.5 Then
           Call REPLA(X$, MKI$(HOI%), 39, 2)
           Call GRAREG("!CARMECO", X$, U&)
           FEMI = HOI%
        End If
        F1 = FEMI: F2 = HOI%
        If FEMI < HOI% - 32 * 13 Or FEMI > HOI% + 30 Or DIENTRE%(F1, F2) > 365 Then
           Call MENSERR(24, "Fecha de Emisión no Válida")
           GoTo 200
        End If
        '
        FEVE = CVI(Mid$(X$, 41, 2))
        If Abs(FEVE) < 0.5 Then
            Call REPLA(X$, MKI$(HOI%), 41, 2)
            Call GRAREG("!CARMECO", X$, U&)
            FEVE = HOI%
        End If
        F1 = FEVE: F2 = HOI%
        If FEVE < HOI% - 64 Or DIENTRE%(F1, F2) > 30 Or FEVE < FEMI - 0.001 Then
          Call MENSERR(24, "Fecha de Vencimiento no Válida")
          GoTo 200
        End If
        '
        PROTER$ = UCase$(TRIM$(Mid$(X$, 43, 2)))
        If PROTER$ = "" Or Left$(PROTER$, 1) = "P" Then
             PROTER$ = "  "
           Else
             PROTER$ = "3 "
        End If
        Call REPLA(X$, PROTER$, 43, 2)
        Call GRAREG("!CARMECO", X$, U&)
        '
        CLEA% = CVI(Mid$(X$, 45, 2))
        If Mid$(X$, 45, 2) = "  " Then
          Call REPLA(X$, MKI$(24), 45, 2)
          Call GRAREG("!CARMECO", X$, U&)
          CLEA% = 24
        End If
        If CLEA% < 24 Then CLEA% = 24
        If (CLEA% Mod 24) <> 0 Then
            Call MENSERR(24, "Clearing no Válido")
            GoTo 200
        End If
        Call REPLA(X$, MKI$(CLEA%), 45, 2)
        Call GRAREG("!CARMECO", X$, U&)
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
    '
    ' EN CASO QUE LOS CUIT QUE FALTEN SEAN DE CHEQUES DE TERCEROS CONSULTA
    If FALTACUIT% > 0 And ATXY% = 1 Then
      OPX% = COMALTER%("Se han Detectado Valores (Cheques de Terceros) sin Cuit del Librador\¿Desea Aplicar a Los Mismos El CUIT Del Cliente Activo? \\No, Ingreso Manual\Sí, Aplicar")
      If OPX% <= 1 Then
         GoTo 200 ' VUELVE A LA VENTANA
      End If
    End If
    
    'AGREGA LOS CUIT DEL CLIENTE AUTOMATICAMENTE
    If FALTACUIT% > 0 Then
        CARTVAL% = CUEICARCH%
        For U& = 1 To ULTREG&("!CARMECO")
          X$ = REGLEIDO$("!CARMECO", U&)
          MECO1% = CVI(Left$(X$, 2))
          '
          'VERIFICA QUE SEA CHEQUE
          CUECO1% = CUEMECOB%(MECO1%)
          If CUECO1% = CARTVAL% Then 'SI ES CHEQUE
              CUITLIB$ = TRIM$(Mid$(X$, 47, 16))
              If CUITLIB$ <> "" Then GoTo 300 ' SI YA TIENE PASA AL SIGUIENTE
              CUITLIB$ = CUITCLI$(NC%)
              Call REPLA(X$, CUITLIB$, 47, 16)
              Call GRAREG("!CARMECO", X$, U&)
         End If
300     Next U&
        Call CIERRARCH("!CARMECO")
        If ATXY% = 1 Then GoTo 210 'SI TIENE CHEQUES DE TERCERO VUELVE A LA SEGUNDA VENTANA
        
    End If
    'FIN AGREGADO DE CUIT PARA CHEQUES
    '
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
    Label14.Caption = TRIM$(CSTRING$(MKD$(SUMAVA#), 4, 14, 2, 2))
    Label14.Refresh
    '\\\OT-07-0029-RG-26/01/07///
    ANTI# = SUMAVA# + XVALO(DESCUEN(0)) - CDbl(XVALO(Label18.Caption))
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
    Call REMPLAGRID(Me, GRID1, Picture1, Label12, List1)
    '
    ' CALCULO DE DIAS DE COBRANZA A PARTIR DE LA FECHA DE VENCIMIENTO
    ' DE LAS FACTURAS CANCELADAS
    If CONTROL$("RECISTA", "DIASVEN") = "FECHAFA" Then
        Call DIASDECOBRANZA_VENCIFACTURA
    End If
    '
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
    Call REMPLAGRID(Me, GRID2, Picture2, Label9, List2)
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
    Label34 = 5: Label34.Refresh
    For U& = 1 To ULTREG&("!CARMECO")
       X$ = REGLEIDO$("!CARMECO", U&)
       Tipo% = CVI(Left$(X$, 2))
       BANCO$ = TRIM$(Mid$(X$, 3, 16))
       '
       OpMovivi$ = "CH" 'ASIGNO VALORES POR DEFECTO
       VaMovivi$ = ""
       '
       If BANCO$ = "" Then
         BANCO$ = TRIM$(ECOARCH$("MECOBRA", MKI$(Tipo%)))
       End If
       If Left$(UCase$(BANCO$), 3) = "CH." Then
         If Len(BANCO$) > 3 Then
           BANCO$ = Mid$(BANCO$, 4)
         End If
       End If
       '
       If Left$(UCase$(BANCO$), 3) = "TC." Then
        TIPOSS$ = Space$(4)
        If Len(BANCO$) > 3 Then
            MCB$ = FILTERGETRECORD$("MECOBRA", 1, MKI$(Tipo%))
            TIPOSS$ = TRIM$(Mid$(MCB$, 57, 4))
            VaMovivi$ = TIPOSS$
            OpMovivi$ = "TC" ' SE LE QUITO EL PUNTO A TC. EN  LA GRABACION
            BANCO$ = Mid$(BANCO$, 4)
        End If
       End If
       '
       NUCHE$ = TRIM$(Mid$(X$, 19, 12))
       IMCHE# = ROUND(COTIMONEM# * CVD(Mid$(X$, 31, 8)))
       FECHE% = CVI(Mid$(X$, 39, 2))
       FEVEN% = CVI(Mid$(X$, 41, 2))
       CUITLIB$ = TRIM$(Mid$(X$, 47, 16))
       '
       If Abs(IMCHE#) > 0 Then
          SUDEBE# = SUDEBE# + IMCHE#
             FEX = FEVEN%: FEVECHE$ = FECHATEX$(FEX)
          FeContab = CVDAT(FC%)
             CUECO% = CUEMECOB%(Tipo%)
          CuentaCo% = CUECO%
          NumeClie% = NCLIEN%
          NumeProv% = 0
          TiMovi$ = "CHRECI"
          OpMovi$ = OpMovivi$
          VaMovi$ = VaMovivi$
             IDEVA$ = AJUSTI$(BANCO$, 18)
             IDEVA$ = IDEVA$ + AJUSTI$(NUCHE$, 14)
             IDEVA$ = IDEVA$ + FEVECHE$
          Descrip$ = IDEVA$
             If FECHE% < 1 Then FECHE% = FC%
          FechEmi = CVDAT(FECHE%)
          FeAcredi = CVDAT(FEVEN%)
          NumeTal = COBRACT%
'          CodiEmp% = 0
'          CodiSuc% = SUCUI%
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
    Label34 = 6: Label34.Refresh
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
'       CodiEmp% = 0
'       CodiSuc% = SUCUI%
       NumeCom = NROO&
       StatInit = 0
       ImpoDeb# = IREDON#
       Observa$ = ""
       Call GRACABAN
    End If
    '
843 ' acumula anticipo entre detalle y casillero
    For U& = 1 To List2.ListCount
      Z$ = List2.List(U& - 1)
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
      FechEmi = CVDAT(FC%)
      FeVenc = CVDAT(FC%)
      ImVenc# = (-1) * ANTI#
      Call GRADEUPEN
    End If
    OKEY% = 1
    '
800 ' *********************************** aqui blanquear facturas canceladas
    CoComCor$ = RECICORTO$
    CoComLar$ = RECILARGO$
'    CodiEmp% = 0
'    CodiSuc% = SUC%
    NumeClie% = NCLIEN%
    NumeProv% = 0
    FechEmi = CVDAT(FC%)
    RefApli$ = "Cobranza " + RECICORTO$
    '
    XX% = 0
    VACON% = 1
    SUCANCE# = 0: ENCUENDE% = 0
    For U& = 1 To List2.ListCount
       Z$ = List2.List(U& - 1)
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
             'jandy sql ver order by fechavencimiento y tipo de recordset
             
             strsql = "Select * from SaDeuPen where Nuclien =" & NCLIEN%
             strsql = strsql & " And CODICOM_LAR = '" & DOCUCA$ & "'"
             If SICAN% > 0 Then
                 strsql = strsql & " And SALDDEBE > 0 "
               Else
                 strsql = strsql & " And SALDACRE > 0 "
             End If
             '
             Dim SaDeuPen As ADODB.Recordset
             Set SaDeuPen = New ADODB.Recordset
25           On Error Resume Next
             SaDeuPen.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
             If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 25
             End If
             On Error GoTo 0
             '
             With SaDeuPen
             If Not (.EOF And .BOF) Then
               .MoveFirst
               Do While Not (.EOF)
                   ENCUENDE% = 1
                   PARCAN# = ICAN#
                   If SICAN% * (!salddebe - !saldacre) < SICAN% * ICAN# Then PARCAN# = !salddebe - !saldacre
                   If PARCAN# > 0 Then
                        !salddebe = !salddebe - PARCAN#
                      Else
                        !saldacre = !saldacre + PARCAN#
                   End If
                   .Update
                   TOTAPLI# = TOTAPLI# + PARCAN#   ' SUMA APLICACIONES -  agregado por tema bimonetario 22/03
                   '
                   DocOrCor$ = !CODICOM_COR
                   TotOrig# = !ibrucomp
                   IIVAORIG# = 0
                   On Error Resume Next
                   IvaOrig# = !IIvaComp
                   On Error GoTo 0
                   '
                   ICAN# = ICAN# - PARCAN#
                   If Abs(ICAN#) >= 0.005 Then  'me indico Flavio
                      .MoveNext
                     Else
                       Exit Do
                   End If
                Loop
              End If
             End With
             SaDeuPen.Close
             Set SaDeuPen = Nothing

'             With SaDeuPen
'                .FindFirst "NuClien = " & NCLIEN% & " AND CODICOM_LAR = '" & DOCUCA$ & "' AND SALDDEBE > 0"
'850             If .NoMatch = False Then
'                   ENCUENDE% = 1
'                   PARCAN# = ICAN#: If !SaldDebe < ICAN# Then PARCAN# = !SaldDebe
'                   .Edit
'                   !SaldDebe = !SaldDebe - PARCAN#
'                   .Update
'                   TOTAPLI# = TOTAPLI# + PARCAN#   ' SUMA APLICACIONES -  agregado por tema bimonetario 22/03
'                   '
'                   DocOrCor$ = !CodiCom_Cor
'                   TotOrig# = !IBruComp
'                   IIVAORIG# = 0
'                   On Error Resume Next
'                   IvaOrig# = !IIvaComp
'                   On Error GoTo 0
'                   '
'                   ICAN# = ICAN# - PARCAN#
'                   If ICAN# >= 0.005 Then
'                     .FindNext "Nuclien = " & NumeClie% & " AND CODICOM_LAR = '" & DOCUCA$ & "' AND SALDDEBE > 0"
'                     GoTo 850
'                   End If
'                End If
'             End With
             
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
             If TOTAPLI# * SICAN% < CANCE# * SICAN% - 0.005 Then
                 ' acumula la diferencia para ajuste de saldo
                 Z1$ = Left$(Z$ + Space$(256), 256)
                 DIFAPLI$ = FORMATNUM$(CANCE# - TOTAPLI#, "F12.2")
                 Call REPLA(Z1$, DIFAPLI$, 241, 12)
                 List2.RemoveItem U& - 1
                 List2.AddItem Z1$, U& - 1
             End If
          End If
       End If
    Next U&
    
    VACON% = 0
    '
    Label34 = 7: Label34.Refresh
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
'    CodiEmp% = 0
'    CodiSuc% = SUCUI%
    NumeCom = NROO&
    StatInit = 0
    ImpoDeb# = (-1) * SUDEBE#
    AtraCance& = XVALO(Label17)
    AtraCheque& = XVALO(Label19)
    Call GRACABAN
    OKEY% = 4
    '
    Label34 = 8: Label34.Refresh
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
          TTXX$ = TTXX$ + "Sudebe = " + TRIM$(FORMATNUM$(SUDEBE#, "F10.4")) + "; "
          TTXX$ = TTXX$ + "Sucance = " + TRIM$(FORMATNUM$(SUCANCE#, "F10.4")) + "; "
          TTXX$ = TTXX$ + "Anti = " + TRIM$(FORMATNUM$(ANTI#, "F10.4")) + "; "
          TTXX$ = TTXX$ + "CotiMonem = " + TRIM$(FORMATNUM$(COTIMONEM#, "F10.4")) + "; "
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
         DIFE1# = XVALO(Mid$(List2.List(U5& - 1), 241, 12))
         If DIFE1# >= 0.005 Then
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
'             CodiEmp% = 0
'             CodiSuc% = SUCUI%
             NumeCom = NROO&
             StatInit = 0
             ImpoDeb# = DIFE1#
             Call GRACABAN
             '
             ' graba contrapartida del ajuste
             CuentaCo% = IAJU%
             TiMovi = "AJUSALD"         ' AJUSTE DE SALDO DEUDORES
             OpMovi = "AS"              ' AJUSTE DE SALDO
             VaMovi = "NO-IMP"          ' NO IMPUTADO
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
        X$ = List2.List(U& - 1)
        CANCE# = CDbl(XVALO(Mid$(X$, 70, 14)))
        If CANCE# >= 0.005 Then
          ABUSQUE$ = TRIM$(Left$(X$, 18))
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
                     'If MODOMONEI% = 2 And MONEMI% <= 1 Then
                     If (MODOMONEI% = 1 Or MODOMONEI% = 2) And MONEMI% <= 1 Then
                        Call REPLA(YCB$, String$(32, 0), 89, 32)
                     End If
                  Call REGAPP("CARBIMON", YCB$)
                  '
                  DIFE1# = XVALO(Mid$(List2.List(U& - 1), 241, 12))
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
'                     CodiEmp% = 0
'                     CodiSuc% = SUC%
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
        End If
835   Next U&
    End If
    Call CIERRARCH("CARBIMON")
    '
    ' RECUPERAR EL SALDO EN DOLARES
    For U& = 1 To List2.ListCount
      X$ = List2.List(U& - 1)
      CANCE# = CDbl(XVALO(Mid$(X$, 70, 14)))
      If CANCE# >= 0.005 Then
        ABUSQUE$ = TRIM$(Left$(X$, 18))
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
'   CodiEmp% = 0
'   CodiSuc% = 0
   NumeClie% = Int(GRUDIVCLI(NCLIEN))
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
      X$ = REGLEIDO$("!CARMECO", U&)
      If X$ = REBLA$ Then GoTo 19
      NUVA$ = Mid$(X$, 19, 12)
      IMVA# = CVD(Mid$(X$, 31, 8))
      FEMII% = CVI(Mid$(X$, 39, 2))
      FEVEI% = CVI(Mid$(X$, 41, 2))
      PROTER$ = UCase$(TRIM$(Mid$(X$, 43, 2)))
      CLEA% = CVI(Mid$(X$, 45, 2))
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
      X$ = List2.List(U& - 1)
      DOCU$ = Mid$(X$, 1, 18)
      CANCE# = CDbl(XVALO(Mid$(X$, 70, 14)))
      
      
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
        X$ = REGLEIDO$("!CARMECO", U&)
        If X$ = REBLA$ Then GoTo 19
        MECO% = CVI(Left$(X$, 2))
           
        BANCO$ = TRIM$(ECOARCH$("MECOBRA", MKI$(MECO%)))
        NUVA$ = Mid$(X$, 19, 12)
        IMVA# = CVD(Mid$(X$, 31, 8))
        FEMII% = CVI(Mid$(X$, 39, 2))
        FEVEI% = CVI(Mid$(X$, 41, 2))
        PROTER$ = UCase$(TRIM$(Mid$(X$, 43, 2)))
        CLEA% = CVI(Mid$(X$, 45, 2))
        CUITLIB$ = TRIM$(Mid$(X$, 47, 16))
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
        CUENTA$ = CODCUE$(CUEMECOB%(MECO%))
        Call REPLA(Z$, CUENTA$, 111, 12)
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
    If Abs(XVALO(Text8)) >= 0.005 Then ' ajuste redondeo
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
        X$ = List2.TEXT
        DOCU$ = Mid$(X$, 1, 18)
        CANCE# = CDbl(XVALO(Mid$(X$, 70, 14)))
        FECEMI$ = Mid(X$, 23, 8)
        FECVEN$ = Mid(X$, 47, 8)
        
        If Abs(CANCE#) >= 0.005 Then
            V& = V& + 1
            If V& <= V0& Then
                  Z$ = REGLEIDO$("!PRINRECI", V&)
               Else
                  Z$ = REGBLAN$("!PRINRECI")
            End If
            Z$ = Left$(Z$ + Space(128), 128)
            Call REPLA(Z$, DOCU$, 53, 18)
            Call REPLA(Z$, MKD$(CANCE#), 71, 8)
            Call REPLA(Z$, FECEMI$, 95, 8)
            Call REPLA(Z$, FECVEN$, 103, 8)
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
        TIPODOC$ = "Recibo ": NUMEDOC$ = NRORECI$

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
            SUMAVA# = 0: SUMACAN# = 0: NUHOJ% = 0
            Call CARFORWIN
            '
            'AGREGADO PARA QUE IMPRIMA LAS OBSERVACIONES 13/06/12 - OD
            YAC = 0
            For U& = 1 To ULTREG&("!OBSERVOF")
               TXOB$ = TRIM$(REGLEIDO$("!OBSERVOF", U&))
               Call PRINTFORWIN(TXOB$, "OBSERVA", 0, YAC)
               YAC = YAC + INTERLI
            Next U&
            Call CIERRARCH("!OBSERVOF")

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
            CANRE1% = ULTREG&("!PRINRECI")
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
                   NUHOJ% = NUHOJ% + 1
                   If NUHOJ% = 1 Then
                     On Error Resume Next
                     CANHOJ& = 1 + Int((CANRE1% / (YX& - 1)) - 0.05)
                     On Error GoTo 0
                   End If
                   Call PRINTFORWIN(TRIM$(Str$(NUHOJ%)), "HOJA-NUM", 0, 0)
                   Call PRINTFORWIN(TRIM$(Str$(CANHOJ&)), "HOJA-TOT", 0, 0)
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
                  FEMIS$ = TRIM$(Mid$(YY$, 95, 8))
                Call PRINTFORWIN(FEMIS$, "MEDIDA-1", 0, YAC) 'Fecha de Emision
                  FEMVE$ = TRIM$(Mid$(YY$, 103, 8))
                Call PRINTFORWIN(FEMVE$, "MEDIDA-2", 0, YAC) 'Fecha de Vencimiento
                CUENTA$ = TRIM$(Mid$(YY$, 111, 12))
                Call PRINTFORWIN(CUENTA$, "DEN-CUEN", 0, YAC) 'Fecha de Vencimiento
                  CUE$ = TRIM(CUEMADRE(XVALO(Text1))): DENO$ = TRIM(DENOCUE$(CUEINT(CUE$)))
                  DET$ = CUE$ & " - " & DENO$    'Cuenta Madre y denominacion de la misma
                Call PRINTFORWIN(DET$, "DET-PASE", 0, YAC)
                '
                'IMPRESION DE SALDO DE FACTURA
                NUMFACTU$ = TRIM$(Mid$(YY$, 53, 18))
                TIPX$ = UCase$(Left$(NUMFACTU$, 2))
                SALDO1# = 0
                If TIPX$ = "ND" Or TIPX$ = "FC" Then
                   CONDI$ = "CODICOM_LAR = '" & NUMFACTU$ & "'"
                   SALDO1# = XVALO(VALOBADA("SADEUPEN", "SALDDEBE", CONDI$, "NOMESS"))
                   Call PRINTFORWIN(CSTRING$(MKD$(SALDO1#), 4, 12, 2, 2), "IMPO-PEN", 0, YAC)
                End If
                '
                '//OT-8-870-09/10/08
                For I = 0 To List2.ListCount - 1
                  TXTX$ = List2.List(I)
                  DOCUMENTO$ = Mid$(YY$, 53, 18)
                  If Mid$(TXTX$, 1, 18) = DOCUMENTO$ Then
                    FECHADOCU$ = Mid$(TXTX$, 23, 8)
                    Call PRINTFORWIN(FECHADOCU$, "FECH-REC", 0, YAC)
                    Exit For
                  End If
                Next I
                '//OT-8-870-FIN
                '
            Next YX&
            '
            ENTRAOBS.Text28 = "" ' LIMPIO EL TEXTO DE OBSERVACIONES (QUEDABA CON LO ANTERIOR)

            Call PRINTFORWIN(CSTRING$(MKD$(TOTAVAL#), 4, 12, 2, 2), "TOTA-VAL", 0, 0)
            Call PRINTFORWIN(CSTRING$(MKD$(TOTACAN#), 4, 12, 2, 2), "TOTA-CAN", 0, 0)
            Call PRINTFORWIN(IMPLET1$, "IMP-LET1", 0, 0)
            Call PRINTFORWIN(IMPLET2$, "IMP-LET2", 0, 0)
            '
            NUHOJ% = NUHOJ% + 1
            If NUHOJ% = 1 Then
              On Error Resume Next
              CANHOJ& = 1 + Int((CANRE1% / (YX& - 1)) - 0.05)
              On Error GoTo 0
            End If
            Call PRINTFORWIN(TRIM$(Str$(NUHOJ%)), "HOJA-NUM", 0, 0)
            Call PRINTFORWIN(TRIM$(Str$(CANHOJ&)), "HOJA-TOT", 0, 0)
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
        
        'TIPODOC$ LO LEO ARRIBA POR QUE SI NO CUANDO PRESENTA LA VENTANA DE LAS OBSERVACIONES
        'TOMA EL DEL REMITO QUE ES EL ULTIMO DATO ASIGNADO
'        TIPODOC$ = "Recibo ": NUMEDOC$ = NRORECI$
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
     Label14.Caption = ""
     Label18.Caption = ""
     Label20.Caption = ""
     DESCUEN(0) = ""
     Text7 = ""
     Frame8.Caption = "Descuento (N./Cred)"
     BOTNEXT.Enabled = True
   End If
   Call REMPLAGRID(Me, GRID1, Picture1, Label12, List1)
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
         X$ = REGLEIDO$("!TAMONED", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG("TAMONED", X$, U&)
      Next U&
      FIN& = ULTREG&("!TAMONED")
      Call SETULTREG("TAMONED", FIN&)
   End If
   Call CIERRARCH("TAMONED")
   '
End Sub

Function PRONOCRE$()
   
   'VER ERNESTO flimi MODORAPIDO
   
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
      ElseIf PIVACRE% = 5 Then
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
   'jandy sql
   
   'FLIMI = FLIDIAS: MODORAPIDO% = 1
   'If FLINUME > FLIMI Then
     ' FLIMI = FLINUME
'    MODORAPIDO% = 0
   'End If
   '
'22
   
   Call APERBASDAT("CABAN")
   'JANDY SQL
   'SE CAMBIO LA RUTINA POR EL MAX(NUMECOMP)
   
   FLIMI = FETEXCOR1$(FLINUMERA)
  
  'ASELE$ = "SELECT NumeComp FROM CajaBanc "
   ASELE$ = "SELECT MAX(NumeComp) AS Maximo FROM CajaBanc "
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
   ASELE$ = ASELE$ + "AND FechEmisi >= '" & FLIMI & "' "
   
   'ASELE$ = ASELE$ & "ORDER BY NumeComp DESC;"
   '
   'Set NUFAXI = CueCorri.OpenRecordset(ASELE$, dbOpenDynaset, dbReadOnly)
   '
   NUMAXI& = 0
   Dim MIRS As ADODB.Recordset
   Set MIRS = FLEXCONN.Execute(FILTSQL$(ASELE$))
   If Not IsNull(MIRS!maximo) Then
        NUMAXI& = MIRS!maximo
   End If
   MIRS.Close
   Set MIRS = Nothing
   
   'On Error Resume Next
   'NUFAXI.MoveFirst
   'If Err < 1 Then
   '    On Error GoTo 0
   '    NUMAXI& = NUFAXI!NumeComp
   '  Else
   '    On Error GoTo 0
   '    If MODORAPIDO% = 1 Then
   '       FLIMI = FLINUME
   '       MODORAPIDO% = 0
   '       GoTo 22
   '    End If
   'End If
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
    If VARDOCX$ = "A" Then
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
    CACOPDOC% = XVALO(FORFINFAC.Text2(0).TEXT)
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
    TETABU1$(2, CAITAB1%) = "Desc." + TRIM$(Text7) + " % S/Rbo." + TRIM$(Text11)
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
  Label35 = ""
  PROVENCI# = 0
End Sub

Private Sub Timer1_Timer()
   '
   SUVALO# = 0
   For U& = 1 To ULTREG&("!CARMECO")
      X$ = REGLEIDO$("!CARMECO", U&)
      SUVALO# = SUVALO# + ROUND#(CVD(Mid$(X$, 31, 8)))
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


