VERSION 5.00
Begin VB.Form LISOMAQ07 
   BackColor       =   &H00DDDDC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Mantenimiento  - Opciones"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   945
      Picture         =   "LISOMAQ07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Revertir Anulaciones de Ordenes de Compra"
      Top             =   4305
      Visible         =   0   'False
      Width           =   560
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00DDDDC0&
      Caption         =   "Gestión de O/M's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3315
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   4005
      Begin VB.Frame Frame9 
         BackColor       =   &H00DDDDC0&
         Caption         =   "Ordenes de Mantenimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2715
         Left            =   240
         TabIndex        =   39
         Top             =   400
         Width           =   3585
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   960
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   40
            Top             =   1900
            Width           =   2325
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
               Picture         =   "LISOMAQ07.frx":014A
               Style           =   1  'Graphical
               TabIndex        =   41
               ToolTipText     =   "Anulación de Orden de Mantenimiento"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Anulación no Reversible"
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
               TabIndex        =   42
               Top             =   105
               Width           =   1275
            End
         End
         Begin VB.PictureBox Picture7 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   43
            Top             =   1400
            Width           =   2325
            Begin VB.CommandButton Command46 
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
               Picture         =   "LISOMAQ07.frx":0454
               Style           =   1  'Graphical
               TabIndex        =   52
               ToolTipText     =   "Cierre de Orden de Mantenimiento"
               Top             =   0
               Width           =   530
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Cierre de Orden de Mantenimiento"
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
               TabIndex        =   44
               Top             =   60
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture15 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   47
            Top             =   900
            Width           =   2325
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
               Picture         =   "LISOMAQ07.frx":075E
               Style           =   1  'Graphical
               TabIndex        =   48
               ToolTipText     =   "Modificar Orden de Mantenimiento no Anulada"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Modificación Inter-Activa"
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
               TabIndex        =   49
               Top             =   60
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture9 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   45
            Top             =   400
            Width           =   2325
            Begin VB.CommandButton Command22 
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
               Picture         =   "LISOMAQ07.frx":08A8
               Style           =   1  'Graphical
               TabIndex        =   50
               ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Ver y Re-Imprimir O/Mantenimiento"
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
               TabIndex        =   46
               Top             =   60
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Materiales Consignados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   315
         TabIndex        =   34
         Top             =   4000
         Visible         =   0   'False
         Width           =   3585
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2895
            TabIndex        =   35
            Top             =   315
            Visible         =   0   'False
            Width           =   2955
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
               Left            =   2340
               Picture         =   "LISOMAQ07.frx":0BB2
               Style           =   1  'Graphical
               TabIndex        =   37
               ToolTipText     =   "Consumo Materiales Consignados"
               Top             =   0
               Width           =   560
            End
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
               Left            =   0
               Picture         =   "LISOMAQ07.frx":0CFC
               Style           =   1  'Graphical
               TabIndex        =   36
               ToolTipText     =   "Entrega en Consignación "
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Consignación    de Materiales"
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
               TabIndex        =   38
               Top             =   105
               Width           =   1590
            End
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0FFFF&
      Caption         =   "Consultas Varias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   240
      TabIndex        =   22
      Top             =   5760
      Visible         =   0   'False
      Width           =   4005
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0FFFF&
         Caption         =   "Control de Ultimas Compras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1545
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Width           =   3585
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   30
            Top             =   735
            Width           =   2220
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
               Picture         =   "LISOMAQ07.frx":1006
               Style           =   1  'Graphical
               TabIndex        =   31
               ToolTipText     =   "Ultimas Compras por Proveedor"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Control por Proveedor"
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
               TabIndex        =   32
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture12 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   27
            Top             =   420
            Width           =   2220
            Begin VB.CommandButton Command20 
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
               Picture         =   "LISOMAQ07.frx":1310
               Style           =   1  'Graphical
               TabIndex        =   28
               ToolTipText     =   "Ultimas Compras por Material"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Por Artículo"
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
               Top             =   50
               Width           =   1590
            End
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   240
         ScaleHeight     =   585
         ScaleWidth      =   2160
         TabIndex        =   23
         Top             =   480
         Width           =   2220
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
            Picture         =   "LISOMAQ07.frx":145A
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Aprobar Ordenes de Compra"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Aprobar Orde- nes de Compra"
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
            Width           =   1485
         End
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00404000&
      Height          =   5755
      Left            =   8505
      ScaleHeight     =   5700
      ScaleWidth      =   1005
      TabIndex        =   16
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   140
         ScaleHeight     =   75
         ScaleWidth      =   495
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2775
         Width           =   555
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -105
            TabIndex        =   21
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command18 
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
         Picture         =   "LISOMAQ07.frx":1764
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   1595
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "LISOMAQ07.frx":1BA6
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   905
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "LISOMAQ07.frx":1EB0
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   215
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -255
         Picture         =   "LISOMAQ07.frx":1FFA
         Stretch         =   -1  'True
         Top             =   3015
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDC0&
      Caption         =   "Control de O/M's Pendientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3315
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   4005
      Begin VB.Frame Frame6 
         BackColor       =   &H00DDDDC0&
         Caption         =   "Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   210
         TabIndex        =   13
         Top             =   2260
         Width           =   3585
         Begin VB.OptionButton Option1 
            BackColor       =   &H00DDDDC0&
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
            Left            =   420
            TabIndex        =   15
            Top             =   390
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00DDDDC0&
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
            Left            =   1995
            TabIndex        =   14
            Top             =   390
            Width           =   1275
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00DDDDC0&
         Caption         =   "Opciones de Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1725
         Left            =   210
         TabIndex        =   1
         Top             =   400
         Width           =   3585
         Begin VB.PictureBox Picture11 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   2
            Top             =   2640
            Visible         =   0   'False
            Width           =   2325
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
               Picture         =   "LISOMAQ07.frx":3F1C
               Style           =   1  'Graphical
               TabIndex        =   3
               ToolTipText     =   "Material en Consignación"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Material en Consignación"
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
               TabIndex        =   4
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   5
            Top             =   2400
            Visible         =   0   'False
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
               Left            =   -20
               Picture         =   "LISOMAQ07.frx":4226
               Style           =   1  'Graphical
               TabIndex        =   6
               ToolTipText     =   "Pendientes por Material"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
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
               Left            =   735
               TabIndex        =   7
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   740
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   8
            Top             =   900
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
               Left            =   0
               Picture         =   "LISOMAQ07.frx":4530
               Style           =   1  'Graphical
               TabIndex        =   51
               ToolTipText     =   "Ordenes de Mantenimiento por Máquinas\Equipos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "O/M's Pendien- tes por Equipo"
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
               TabIndex        =   9
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   500
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   10
            Top             =   400
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
               Picture         =   "LISOMAQ07.frx":483A
               Style           =   1  'Graphical
               TabIndex        =   11
               ToolTipText     =   "Ordenes de Mantenimiento por Número"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Pendientes por Número de O/M"
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
               Left            =   700
               TabIndex        =   12
               Top             =   60
               Width           =   1590
            End
         End
      End
   End
End
Attribute VB_Name = "LISOMAQ07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HAYMATERIALIB%

'*****/////COMANDO OCULTO/////*****
'Se oculto en el frame8 ("Opciones de Consulta")
'el Command9("Listado de Material en Consignación)"
'*****/////**************/////*****

Private Sub command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Private Sub Command11_Click()
  '
  Command11.Enabled = False
  Call Command22_Click
  Command11.Enabled = True
  '
End Sub
Private Sub Command17_Click()
    '
    Command17.Enabled = False
    Screen.MousePointer = 11
    J& = 0: K& = 0: OCOA& = 0
    '
    For I& = 1 To ULTREG&("OCOMPRA")
         XX$ = REGLEIDO$("OCOMPRA", I&)
         '
         If CVS(Left(XX$, 4)) >= 0.5 Then    ' depura O/C's
             K& = K& + 1                     ' sin numero
             Call GRAREG("OCOMPRA", XX$, K&)
         End If
         '
         ST% = Asc(Mid$(XX$, 62, 1))
         '
         If ST% < 1 Then     '    excluye cerradas
            If CVS(Left$(XX$, 4)) <> OCOA& Then
               NP% = CVI(Mid$(XX$, 7, 2))
               If NP% < 1 Or NP% = 9999 Then
                  J& = J& + 1
                  Call GRAREG("!OCOMPRA", XX$, J&)
                  OCOA& = CVS(Left$(XX$, 4))
               End If
            End If
         End If
    Next I&
    '
    Call SETULTREG("OCOMPRA", K&)
    Call SETULTREG("!OCOMPRA", J&)
    Call CIERRARCH("!OCOMPRA")
    '
    If J& < 1 Then
          Call MENSERR(24, "No hay Pedidos de Suministro\Pendientes que Puedan Asignarse.")
          Screen.MousePointer = 1
          GoTo 99
    End If
    '
    INDI% = 1
    Screen.MousePointer = 1
    '
    XX$ = REGSEL$("!OCOMPRA")
    If Len(XX$) <= 4 Then GoTo 99
    '
    NUOCO& = CVS(Left$(XX$, 4))
    '
    YY$ = REGSEL$("PROVED1")
    If Len(YY$) <= 4 Then GoTo 99
    NP1% = CVI(Left$(YY$, 2))
    PROVEX$ = Mid$(YY$, 3, 30)
    '
    Screen.MousePointer = 11
    JJ& = 0
    FIN& = ULTREG&("OCOMPRA")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("OCOMPRA", U&)
       '
       If CVS(Left$(XX$, 4)) = NUOCO& Then
          CI0% = CVI(Mid$(XX$, 11, 2))
          CAOR = CVS(Mid$(XX$, 13, 4))
          FENTRE% = CVI(Mid$(XX$, 9, 2))
          YY$ = REGBLAN$("CARDETOC")
          Call REPLA(YY$, MKI$(CI0%), 1, 2)
          Call REPLA(YY$, UNIMED$(CI0%), 3, 4)
          Call REPLA(YY$, MKS$(CAOR), 47, 4)
          Call REPLA(YY$, MKI$(FENTRE%), 67, 2)
          Call REPLA(YY$, MKS$(U&), 71, 4)
          JJ& = JJ& + 1
          Call GRAREG("!CARDETOC", YY$, JJ&)
       End If
    Next U&
    '
    Call SETULTREG("!CARDETOC", JJ&)
    Call CIERRARCH("!CARDETOC")
    Screen.MousePointer = 1
    '
    VTB% = VENTABU%("CAROCOM2/TITUPAN=P.S.Nro " + TRIM$(Str$(NUOCO&)) + " - " + TRIM$(RASOCLI$((-1) * NP1%)))
    If VTB% < 0 Then GoTo 99
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("OCOMPRA")
    For U& = 1 To ULTREG&("!CARDETOC")
       YY$ = REGLEIDO$("!CARDETOC", U&)
       CAOR = CVS(Mid$(YY$, 47, 4))
       REOCO& = CVS(Mid$(YY$, 71, 4))
       If REOCO& > 0 Then
          If REOCO& <= FIN& Then
             XX$ = REGLEIDO$("OCOMPRA", REOCO&)
             If CVS(Left$(XX$, 4)) = NUOCO& Then
                Call REPLA(XX$, MKI$(NP1%), 7, 2)
                Call REPLA(XX$, MKS$(CAOR), 13, 4)
                Call GRAREG("OCOMPRA", XX$, REOCO&)
             End If
          End If
       End If
    Next U&
    '
    FIN& = ULTREG&("MACONSIG")
    For U& = 1 To FIN&
       XX$ = REGLEIDO$("MACONSIG", U&)
       If Val(Mid$(XX$, 24, 6)) = NUOCO& Then
          Call REPLA(XX$, MKI$(NP1%), 3, 2)
          Call GRAREG("MACONSIG", XX$, U&)
       End If
    Next U&
    '
    Call CIERRARCH("*.*")
    Call FRESHECHO("!OCOMPRA")
    Screen.MousePointer = 1
    '
99  Command17.Enabled = True
    '
End Sub

Sub Command13_Click()
   '
   Command13.Enabled = False
   Call COPYSTRU("INDIOMAN", "INPEMOMA")
   Call BLANARCH("!INPEMOMA")
   '
   Call ABREOMAN
   '
   SQLX$ = "SELECT * FROM ORDEMANT "
   SQLX$ = SQLX$ + " WHERE Status_OM >=0 "
   SQLX$ = SQLX$ + " AND Status_OM < 2" 'ni cerrada ni anulada
   SQLX$ = SQLX$ + " ORDER BY NumeOrMa_OM ASC "
   'Set OMTMP = Omantin.OpenRecordset(SQLX$, 4)
   Set OMTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With OMTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       Call MENSERR(24, "No Existen O/M's para Modificar")
       GoTo 99
     End If
     On Error GoTo 0
    '
    Y$ = REGBLAN$("INPEMOMA")
    JJ& = 0
    Do While .EOF = False
      NRO& = !NumeOrMa_OM
      Call REPLA(Y$, MKS$(NRO&), 1, 4)
      DES1$ = !codigo_equ
      DES1$ = DES1$ + " - " + Mid$(!Referen_OM, 1, 20)
      DES1$ = DES1$ + " - " + FECHATEX$(CVINT(!FechGen_OM))
      Call REPLA(Y$, DES1$, 5, 44)
      JJ& = JJ& + 1
      Call GRAREG("!INPEMOMA", Y$, JJ&)
      .MoveNext
    Loop
    '
  End With
  '
  Call SETULTREG("!INPEMOMA", JJ&)
  Call CIERRARCH("!INPEMOMA")
  '
  Call FRESHECHO("!INPEMOMA")
  YYY$ = REGSEL$("!INPEMOMA/Indice de Ordenes de Mantenimiento Pendientes para Modificación")
  '
  If Len(YYY$) <= 4 Then GoTo 99
  '
  NRO& = CVS(Left$(YYY$, 4))
  If NRO& > 0 Then
    Call MODIFOM(NRO&)
  End If
  '
99 Command13.Enabled = True
End Sub
Sub MODIFOM(NRO&)
    '
    Call ABREOMAN
    '
    Dim OrdeMant As ADODB.Recordset
    SQLX$ = "SELECT * FROM ORDEMANT WHERE NumeOrMa_OM = " & NRO& & ""
    Set OrdeMant = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With OrdeMant
      If Not (.EOF And .BOF) Then
        MOMANTEN07.Text1.TEXT = !codigo_equ
        MOMANTEN07.Text24.TEXT = FORMATNUM$(CVINT(!FechGen_OM), "D8")
        NOM1& = !NumeOrMa_OM
        MOMANTEN07.Label22 = FORMATNUM$(NOM1&, "F9.0")
        MOMANTEN07.Text28 = !Descri_OM
        MOMANTEN07.Text2 = FORMATNUM$(CVINT(!FeProCom_OM), "D8")
        MOMANTEN07.Text5 = Format$(!HoProCom_OM, "hh:mm")
        MOMANTEN07.Text4 = FORMATNUM$(CVINT(!FeProFin_OM), "D8")
        MOMANTEN07.Text6 = Format$(!HoProFin_OM, "hh:mm")
        MOMANTEN07.Text3 = !Proved_OM
        MOMANTEN07.Text7 = !Referen_OM
        COSTER1 = !CostoTer_OM
        MOMANTEN07.Text8.TEXT = FORMATNUM$(COSTER1, "F9.2")
      End If
    End With
    'CARGO LOS OPERARIOS
    MOMANTEN07.List4.Clear
    Dim RecuProp As ADODB.Recordset
    SQLX$ = "SELECT * FROM recuprop WHERE NumeOrMa_OM = " & NRO& & ""
    Set RecuProp = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With RecuProp
        Do While Not .EOF
          LEG1% = !LegaEmp_OM
          LEG2$ = FORMATNUM$(LEG1%, "F5.0")
          EMP1$ = ECOARCH$("PADRON", MKI$(LEG1%))
          'AGREGO AL LIST
          L1$ = Space(128)
          Call REPLA(L1$, LEG2$, 1, 8)
          Call REPLA(L1$, EMP1$, 10, 20)
          MOMANTEN07.List4.AddItem L1$
          .MoveNext
        Loop
    End With
    'CARGO LAS RESERVAS
    BLANARCH ("!REPUESTO")
    '
    Call ABRESERVA
    Dim Reserva As ADODB.Recordset
    SQLX$ = "SELECT * FROM Reserva WHERE SUBSTRING(IdSubOr,4,6) = " & NRO& & ""
    Set Reserva = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With Reserva
        Do While Not .EOF
          'AGREGO AL LIST
           CODI1% = !Cod_Inte
           CODI2$ = CODEXT$(CODI1%)
           DESC2$ = DESCRIT0(CODI1%)
           UM2$ = UNIMED$(CODI1%)
           USO1 = !CantRes
           USO2$ = FORMATNUM$(USO1, "F9.4")
           On Error Resume Next
           OBS2$ = !Observa
           On Error GoTo 0
           'AGREGO AL LIST
           L1$ = Space(128)
           Call REPLA(L1$, CODI2$, 1, 16)
           Call REPLA(L1$, DESC2$, 18, 40)
           Call REPLA(L1$, UM2$, 59, 2)
           Call REPLA(L1$, USO2$, 62, 10)
           Call REPLA(L1$, OBS2$, 73, 25)
           MOMANTEN07.List1.AddItem L1$
           'AGREGO AL ARCHIVO
           Y2$ = REGBLAN$("REPUESTO")
           Call REPLA(Y2$, MKI$(CODI1%), 1, 2)
           Call REPLA(Y2$, UM2$, 3, 2)
           Call REPLA(Y2$, MKS(USO1), 5, 4)
           Call REPLA(Y2$, OBS2$, 9, 25)
           JJ& = JJ& + 1
           Call GRAREG("!REPUESTO", Y2$, JJ&)
           .MoveNext
        Loop
    End With
  Call SETULTREG("!REPUESTO", JJ&)
  Call CIERRARCH("!REPUESTO")
  MOMANTEN07.Show 1
  '
End Sub
Sub CIEOMAN(NRO&)
 '
    Call ABREOMAN
    '
    Dim OrdeMant As ADODB.Recordset
    'Set OrdeMant = New ADODB.Recordset
    SQLX$ = "SELECT * FROM ORDEMANT WHERE NumeOrMa_OM = " & NRO& & ""
    'OrdeMant.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    Set OrdeMant = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With OrdeMant
        'If Not (.EOF And .BOF) Then
        Do While Not OrdeMant.EOF
            CIMANTEN07.Text1.TEXT = !codigo_equ
            CIMANTEN07.Text24.TEXT = FORMATNUM$(CVINT(!FechGen_OM), "D8")
            NOM1& = !NumeOrMa_OM
            CIMANTEN07.Label22 = FORMATNUM$(NOM1&, "F9.0")
            CIMANTEN07.Text28 = !Descri_OM
            CIMANTEN07.Text2 = FORMATNUM$(CVINT(!FeProCom_OM), "D8")
            CIMANTEN07.Text5 = Format$(!HoProCom_OM, "hh:mm")
            CIMANTEN07.Text4 = FORMATNUM$(CVINT(!FeProFin_OM), "D8")
            CIMANTEN07.Text6 = Format$(!HoProFin_OM, "hh:mm")
            CIMANTEN07.Text3 = !Proved_OM
            CIMANTEN07.Text7 = !Referen_OM
            CIMANTEN07.Text8 = FORMATNUM$(!CostoTer_OM, "F12.2")
        'End If
        OrdeMant.MoveNext
        Loop
    End With
    'CARGO LOS OPERARIOS
    CIMANTEN07.List4.Clear
  
    SQLX$ = "SELECT * FROM recuprop WHERE NumeOrMa_OM = " & NRO& & ""
    Dim RecuProp As ADODB.Recordset
    Set RecuProp = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With RecuProp
        Do While Not RecuProp.EOF
        'If Not (.EOF And .BOF) Then
          LEG1% = !LegaEmp_OM
          LEG2$ = FORMATNUM$(LEG1%, "F5.0")
          EMP1$ = ECOARCH$("PADRON", MKI$(LEG1%))
          'AGREGO AL LIST
          L1$ = Space(128)
          Call REPLA(L1$, LEG2$, 1, 8)
          Call REPLA(L1$, EMP1$, 10, 20)
          CIMANTEN07.List4.AddItem L1$
        'End If
        RecuProp.MoveNext
        Loop
    End With
    'CARGO LAS RESERVAS
    BLANARCH ("!REPUESTO")
    '
    Call ABRESERVA
    Dim Reserva As ADODB.Recordset
    'Set Reserva = New ADODB.Recordset
    SQLX$ = "SELECT * FROM Reserva WHERE SUBSTRING(IdSubOr,4,6)= " & NRO& & ""
    'Reserva.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    Set Reserva = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With Reserva
        Do While Not Reserva.EOF
        'If Not (.EOF And .BOF) Then
          'AGREGO AL LIST
           CODI1% = !Cod_Inte
           CODI2$ = CODEXT$(CODI1%)
           DESC2$ = DESCRIT0(CODI1%)
           UM2$ = UNIMED$(CODI1%)
           USO1 = !CantRes
           USO2$ = FORMATNUM$(USO1, "F9.1")
           On Error Resume Next
           OBS2$ = !Observa
           On Error GoTo 0
           'AGREGO AL LIST
           L1$ = Space(128)
           Call REPLA(L1$, CODI2$, 1, 16)
           Call REPLA(L1$, DESC2$, 18, 40)
           Call REPLA(L1$, UM2$, 59, 2)
           Call REPLA(L1$, USO2$, 62, 10)
           Call REPLA(L1$, USOIN2$, 73, 10)
           'Call REPLA(L1$, OBS2$, 73, 25)
           CIMANTEN07.List1.AddItem L1$
           'AGREGO AL ARCHIVO
           Y2$ = REGBLAN$("REPUESTO")
           Call REPLA(Y2$, MKI$(CODI1%), 1, 2)
           Call REPLA(Y2$, UM2$, 3, 2)
           Call REPLA(Y2$, MKS(USO1), 5, 4)
           Call REPLA(Y2$, OBS2$, 9, 25)
           JJ& = JJ& + 1
           Call GRAREG("!REPUESTO", Y2$, JJ&)
         'End If
         Reserva.MoveNext
         Loop
    End With
  Call SETULTREG("!REPUESTO", JJ&)
  Call CIERRARCH("!REPUESTO")
  '
  CIMANTEN07.List3.Clear
  Call CIMANTEN07.CARCONTA(CIMANTEN07.Text1.TEXT)
  '
  CIMANTEN07.Show 1
  '
End Sub

Private Sub Command18_Click()
    '
    Command18.Enabled = False
      OPOMAQ07.Show 1
    Command18.Enabled = True
    '
End Sub


Private Sub Command19_Click()
   Command19.Enabled = False
   Call FINAL("*CODORPR")
   Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    FOAPROC07.Show 1
    Command2.Enabled = True
End Sub



Private Sub Command21_Click()
   Command21.Enabled = False
   Screen.MousePointer = 11
   JJ& = 0
   FIN& = ULTREG&("OCOMPRA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("OCOMPRA", U&)
      '
      ST% = Asc(Mid$(XX$, 62, 1))
      If ST% < 1 Then              '    excluye cerradas
         NP% = CVI(Mid$(XX$, 7, 2))
         If NP% < 1 Or NP% = 9999 Then
            CAPEN = CVS(Mid$(XX$, 13, 4)) - CVS(Mid$(XX$, 58, 4))
            If CAPEN > 0 Then
               JJ& = JJ& + 1
               Call GRAREG("!OCOMPRA", XX$, JJ&)
            End If
         End If
      End If
   Next U&
   '
   Call SETULTREG("!OCOMPRA", JJ&)
   Call CIERRARCH("!OCOMPRA")
   '
   Screen.MousePointer = 1
   If JJ& < 1 Then
      Call MENSERR(24, "No Hay Pedidos de Suministro Pendientes.")
      GoTo 99
   End If
   '
   VTB% = VENTABU%("PEDSUPEN")
   '
99 Command21.Enabled = True
   '
End Sub

 Sub Command22_Click()
   '
   Command22.Enabled = False
   Call COPYSTRU("INDIOMAN", "INPEOMAN")
   Call BLANARCH("!INPEOMAN")
   '
   Call ABREOMAN
   '
   SQLX$ = "SELECT * FROM ORDEMANT "
   SQLX$ = SQLX$ + " ORDER BY NumeOrMa_OM ASC "
   'Set OMTMP = Omantin.OpenRecordset(SQLX$, 4)
   Set OMTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With OMTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       GoTo 99
     End If
     On Error GoTo 0
    '
    Y$ = REGBLAN$("INPEOMAN")
    JJ& = 0
    ANU1$ = " (A)"
    CER1$ = " (C)"
    Do While .EOF = False
      NRO& = !NumeOrMa_OM
      Call REPLA(Y$, MKS$(NRO&), 1, 4)
      DES1$ = !codigo_equ
      DES1$ = DES1$ + " - " + Mid$(!Referen_OM, 1, 20)
      DES1$ = DES1$ + " - " + FECHATEX$(CVINT(!FechGen_OM))
      If !Status_OM < 0 Then DES1$ = DES1$ + ANU1$
      If !Status_OM = 3 Then DES1$ = DES1$ + CER1$
      Call REPLA(Y$, DES1$, 5, 44)
      JJ& = JJ& + 1
      Call GRAREG("!INPEOMAN", Y$, JJ&)
      .MoveNext
    Loop
    '
  End With
  '
  Call SETULTREG("!INPEOMAN", JJ&)
  Call CIERRARCH("!INPEOMAN")
  '
  Call FRESHECHO("!INPEOMAN")
3 YYY$ = REGSEL$("!INPEOMAN/Indice General de Ordenes de Mantenimiento (Pendientes-Cumplidas-Anuladas)")
  '
  If Len(YYY$) <= 4 Then GoTo 99
  '
  NPX& = CVS(Left$(YYY$, 4))
  '
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Orden de Mantenimiento"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU$ = " " + NUDOCU$ + " "
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
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    GoTo 3
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
    FIN& = Len(TPSP$)
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7070
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7070
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
    UI& = UI& + 1
    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
    If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
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
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.HScroll1.Min = 0
    EDITFORM.HScroll1.Value = 0
    EDITFORM.HScroll1.Max = 1
    If ANCHPAGINA > EDITFORM.Frame1.Width Then
      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command1.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 3
    '
99  Command22.Enabled = True
    '
90 End Sub

Private Sub Command29_Click()
  '
  Call HELPONLINE("LISOMAN")
  '
End Sub

Sub Command3_Click()
    '
    Command3.Enabled = False
    If Option1.Value = True Then 'PANTALLA
      Call COMUNI("Opción por Pantalla. No Disponible por el Momento.")
     ElseIf Option2.Value = True Then 'IMPRESIÓN
      Call LIOMNUM
    End If
    Command3.Enabled = True
    '
End Sub
Sub LIOMEQU()
   '
   HAS% = HOY(HOS$)
   FEX = HAS%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   DES% = FECHANUM(DESS$)
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
10 OKKX% = 0
   OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
      GoTo 99
   End If
   OKKX% = 1
   DES% = CVI(Left$(OBX$, 2)): DES1 = FETEXCOR1$(DES%)
   HAS% = CVI(Mid$(OBX$, 3, 2)): HAS1 = FETEXCOR1$(HAS%)
   If HAS% < DES% Then GoTo 10
   FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   Call AMAEQUI07.INDIEQUI
   Call SELECHO("INDIEQUI/Indice de Equipos")
   MAQ$ = TRIM$(VALACT1$("INDIEQUI"))
   '
   Screen.MousePointer = 11
   '
   Call ABREOMAN
   '
   SQLX$ = "SELECT * FROM ORDEMANT "
   SQLX$ = SQLX$ + " WHERE FechGen_OM >= '" & DES1 & "' "
   SQLX$ = SQLX$ + " AND FechGen_OM <= '" & HAS1 & "'  "
   SQLX$ = SQLX$ + " AND Status_OM = 1 "
   If TRIM$(MAQ$) <> "" Then
     SQLX$ = SQLX$ + " AND Codigo_Equ = '" & MAQ$ & "' "
   End If
   SQLX$ = SQLX$ + " ORDER BY NumeOrMa_OM ASC "
   'Set OMTMP = Omantin.OpenRecordset(SQLX$, 4)
   Set OMTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With OMTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       Call MENSERR(24, "No Existen Registros que Coincidan\ con Criterio SQL.")
       GoTo 99
     End If
     On Error GoTo 0
    '
    Y$ = REGBLAN$("ORDEMANT")
    JJ& = 0
    Do While .EOF = False
      NRO& = !NumeOrMa_OM
      Call REPLA(Y$, MKS$(NRO&), 1, 4)
      Call REPLA(Y$, !Referen_OM, 5, 64)
      Call REPLA(Y$, MKI$(CVINT%(!FechGen_OM)), 69, 2)
      Call REPLA(Y$, !UsuGen_OM, 71, 25)
      Call REPLA(Y$, !codigo_equ, 96, 6)
      Call REPLA(Y$, MKI$(CVINT%(!FechCie_OM)), 102, 2)
          HOCIE$ = Format$(!HoraCie_OM, "hh:mm")
      Call REPLA(Y$, MKI$(Int(HORANUM(HOCIE$))), 104, 2)
      Call REPLA(Y$, !UsuCie_OM, 106, 25)
      ST1% = !Status_OM
      ST2$ = ""
      Select Case ST1%
        Case -1
          ST2$ = "Anulada"
        Case 1
          ST2$ = "Generada"
        Case 3
          ST2$ = "Cerrada"
      End Select
      Call REPLA(Y$, MKI$(ST1%), 131, 2)
      Call REPLA(Y$, ST2$, 133, 10)
      Call REPLA(Y$, MKI$(CVINT%(!FeProCom_OM)), 143, 2)
        HOCOPRO$ = Format$(!HoProCom_OM, "hh:mm")
      Call REPLA(Y$, MKI$(Int(HORANUM(HOCOPRO$))), 145, 2)
      Call REPLA(Y$, MKI$(CVINT%(!FeProFin_OM)), 147, 2)
        HOFIPRO$ = Format$(!HoProFin_OM, "hh:mm")
      Call REPLA(Y$, MKI$(Int(HORANUM(HOFIPRO$))), 149, 2)
        NPRO% = !Proved_OM
      Call REPLA(Y$, MKI$(NPRO%), 151, 2)
      '
      JJ& = JJ& + 1
      Call GRAREG("ORDEMANT", Y$, JJ&)
      .MoveNext
    Loop
    '
  End With
  '
  Call SETULTREG("ORDEMANT", JJ&)
  Call CIERRARCH("ORDEMANT")
  Call FILESORT("ORDEMANT", "ORDEMANT", 6, 6, "ASC")
  '
  Call HEADERS("ORDEMANT", "")
  Call HEADERS("ORDEMANT", PERIO$)
  If TRIM$(MAQ$) <> "" Then
    MAQ1$ = "Equipo: " & MAQ$ & " "
    MAQ1$ = MAQ1$ + " - " + TRIM$(ECOARCH$("INDIEQUI", MAQ$))
    Call HEADERS("ORDEMANT", MAQ1$)
  End If
  '
  Call FINAL("*LIORDEQU")
  '
99 Screen.MousePointer = 1
  '
End Sub
Sub LIOMNUM()
   '
   HAS% = HOY(HOS$)
   FEX = HAS%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   DES% = FECHANUM(DESS$)
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
10 OKKX% = 0
   OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
      GoTo 99
   End If
   OKKX% = 1
   DES% = CVI(Left$(OBX$, 2)): DES1 = FETEXCOR1$(DES%)
   HAS% = CVI(Mid$(OBX$, 3, 2)): HAS1 = FETEXCOR1$(HAS%)
   If HAS% < DES% Then GoTo 10
   FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   Screen.MousePointer = 11
   '
   Call ABREOMAN
   '
   SQLX$ = "SELECT * FROM ORDEMANT "
   SQLX$ = SQLX$ + " WHERE FechGen_OM >= '" & DES1 & "' "
   SQLX$ = SQLX$ + " AND FechGen_OM <= '" & HAS1 & "' "
   SQLX$ = SQLX$ + " AND Status_OM = 1 "
   SQLX$ = SQLX$ + " ORDER BY NumeOrMa_OM ASC "
   'Set OMTMP = Omantin.OpenRecordset(SQLX$, 4)
   Set OMTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With OMTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       Call MENSERR(24, "No Existen Registros que Coincidan\ con Criterio SQL.")
       GoTo 99
     End If
     On Error GoTo 0
    '
    Y$ = REGBLAN$("ORDEMANT")
    JJ& = 0
    Do While .EOF = False
      NRO& = !NumeOrMa_OM
      Call REPLA(Y$, MKS$(NRO&), 1, 4)
      Call REPLA(Y$, !Referen_OM, 5, 64)
      Call REPLA(Y$, MKI$(CVINT%(!FechGen_OM)), 69, 2)
      Call REPLA(Y$, !UsuGen_OM, 71, 25)
      Call REPLA(Y$, !codigo_equ, 96, 6)
      Call REPLA(Y$, MKI$(CVINT%(!FechCie_OM)), 102, 2)
        HOCIE$ = Format$(!HoraCie_OM, "hh:mm")
      Call REPLA(Y$, MKI$(Int(HORANUM(HOCIE$))), 104, 2)
      Call REPLA(Y$, !UsuCie_OM, 106, 25)
      ST1% = !Status_OM
      ST2$ = ""
      Select Case ST1%
        Case -1
          ST2$ = "Anulada"
        Case 1
          ST2$ = "Generada"
        Case 3
          ST2$ = "Cerrada"
      End Select
      Call REPLA(Y$, MKI$(ST1%), 131, 2)
      Call REPLA(Y$, ST2$, 133, 10)
      Call REPLA(Y$, MKI$(CVINT%(!FeProCom_OM)), 143, 2)
        HOCOPRO$ = Format$(!HoProCom_OM, "hh:mm")
      Call REPLA(Y$, MKI$(Int(HORANUM(HOCOPRO$))), 145, 2)
      Call REPLA(Y$, MKI$(CVINT%(!FeProFin_OM)), 147, 2)
        HOFIPRO$ = Format$(!HoProFin_OM, "hh:mm")
      Call REPLA(Y$, MKI$(Int(HORANUM(HOFIPRO$))), 149, 2)
        NPRO% = !Proved_OM
      Call REPLA(Y$, MKI$(NPRO%), 151, 2)
      '
      JJ& = JJ& + 1
      Call GRAREG("ORDEMANT", Y$, JJ&)
      .MoveNext
    Loop
    '
  End With
  '
  Call SETULTREG("ORDEMANT", JJ&)
  Call CIERRARCH("ORDEMANT")
  '
  Call HEADERS("ORDEMANT", "")
  Call HEADERS("ORDEMANT", PERIO$)
  '
  Call FINAL("*LIORDMAN")
  '
99 Screen.MousePointer = 1
  '
End Sub
Sub Command4_Click()
    Command4.Enabled = False
    '
    If Option1.Value = True Then 'PANTALLA
      Call COMUNI("Opción por Pantalla. No Disponible por el Momento.")
     ElseIf Option2.Value = True Then 'IMPRESION
      Call LIOMEQU
    End If
    '
    Command4.Enabled = True
End Sub

 Sub Command14_Click()
  '
  Command14.Enabled = False
   '
12 If COMALTER%("¿Realmente Desea\ANULAR Ordenes de Mantenimiento?\\No, Abandonar\Si, Anular O/M's") <> 2 Then
     GoTo 99
   End If
   '
   Call COPYSTRU("INDIOMAN", "INANOMAN")
   Call BLANARCH("!INANOMAN")
   '
   Call ABREOMAN
   '
   SQLX$ = "SELECT * FROM ORDEMANT "
   SQLX$ = SQLX$ + " WHERE Status_OM >= 0"
   SQLX$ = SQLX$ + " AND Status_OM < 2" 'ni cerrada ni anulada
   SQLX$ = SQLX$ + " ORDER BY NumeOrMa_OM ASC "
   'Set OMTMP = Omantin.OpenRecordset(SQLX$, 4)
   Set OMTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With OMTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       Call MENSERR(24, "No Existen O/M's para Anular")
       GoTo 99
     End If
     On Error GoTo 0
    '
    Y$ = REGBLAN$("INANOMAN")
    JJ& = 0
    Do While .EOF = False
      NRO& = !NumeOrMa_OM
      Call REPLA(Y$, MKS$(NRO&), 1, 4)
      DES1$ = !codigo_equ
      DES1$ = DES1$ + " - " + Mid$(!Referen_OM, 1, 20)
      DES1$ = DES1$ + " - " + FECHATEX$(CVINT(!FechGen_OM))
      Call REPLA(Y$, DES1$, 5, 44)
      JJ& = JJ& + 1
      Call GRAREG("!INANOMAN", Y$, JJ&)
      .MoveNext
    Loop
    '
  End With
  '
  Call SETULTREG("!INANOMAN", JJ&)
  Call CIERRARCH("!INANOMAN")
  '
  Call FRESHECHO("!INANOMAN")
  YYY$ = REGSEL$("!INANOMAN/Indice de Ordenes de Mantenimiento Pendientes para Anulación")
  '
  If Len(YYY$) <= 4 Then GoTo 99
  '
  NRO& = CVS(Left$(YYY$, 4))
  If NRO& > 0 Then
    '
    'MARCO LA OM COMO ANULADA
    Call ABREOMAN
    Call ABRESERVA
    '
    CONDI$ = " NumeOrMa_OM = " & NRO& & " "
    Call ACTUREGISTRO("ORDEMANT", "STATUS_OM", CONDI$, "-1", REG&, "NOMESS")
    Dim Reserva As ADODB.Recordset
    Set Reserva = New ADODB.Recordset
    SQLX$ = "SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + "WHERE SUBSTRING(IdSubOr,4,6) = " & NRO& & ""
    Reserva.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
    Do While Not Reserva.EOF
      Reserva!CantRes = Reserva!CantCons
      Reserva!Saldo_Entre = Reserva!CantRes - Reserva!CantCons
      Reserva.Update
      Reserva.MoveNext
    Loop
'          SQLX$ = " VAL(MID(IdSubOr,4,6)) = " & NRO& & ""
'    With OrdeMant
'      .FindFirst " NumeOrMa_OM = " & NRO& & " "
'      If .NoMatch = False Then
'        .Edit
'          !Status_OM = -1
'        .Update
'        With Reserva
'          SQLX$ = " VAL(MID(IdSubOr,4,6)) = " & NRO& & ""
'          .FindFirst SQLX$
'15        If .NoMatch = False Then
'            .Edit 'CANCELO RESERVAS
'              !CantRes = !CantCons
'              !Saldo_Entre = !CantRes - !CantCons
'            .Update
'            .FindNext SQLX$
'            GoTo 15
'          End If
'        End With
'      End If
'    End With
'    GoTo 12
    '
  End If
  '
99 Command14.Enabled = True
  '
End Sub

 Sub Command46_Click()
 '
  Command46.Enabled = False
   '
   If COMALTER%("¿Realmente Desea\CERRAR Ordenes de Mantenimiento?\\No, Abandonar\Si, Cerrar O/M's") <> 2 Then
     GoTo 99
   End If
   '
   Call COPYSTRU("INDIOMAN", "INCEOMAN")
   Call BLANARCH("!INCEOMAN")
   '
   Call ABREOMAN
   '
   SQLX$ = "SELECT * FROM ORDEMANT "
   SQLX$ = SQLX$ + " WHERE Status_OM >= 0"
   SQLX$ = SQLX$ + " AND Status_OM < 2" 'ni cerrada ni anulada
   SQLX$ = SQLX$ + " ORDER BY NumeOrMa_OM ASC "
   'Set OMTMP = Omantin.OpenRecordset(SQLX$, 4)
   Set OMTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With OMTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       Call MENSERR(24, "No Existen O/M's para Anular")
       GoTo 99
     End If
     On Error GoTo 0
    '
    Y$ = REGBLAN$("INCEOMAN")
    JJ& = 0
    Do While .EOF = False
      NRO& = !NumeOrMa_OM
      Call REPLA(Y$, MKS$(NRO&), 1, 4)
      DES1$ = !codigo_equ
      DES1$ = DES1$ + " - " + Mid$(!Referen_OM, 1, 20)
      DES1$ = DES1$ + " - " + FECHATEX$(CVINT(!FechGen_OM))
      Call REPLA(Y$, DES1$, 5, 44)
      JJ& = JJ& + 1
      Call GRAREG("!INCEOMAN", Y$, JJ&)
      .MoveNext
    Loop
    '
  End With
  '
  Call SETULTREG("!INCEOMAN", JJ&)
  Call CIERRARCH("!INCEOMAN")
  '
  Call FRESHECHO("!INCEOMAN")
  YYY$ = REGSEL$("!INCEOMAN/Indice de Ordenes de Mantenimiento Pendientes para Cierre")
  '
  If Len(YYY$) <= 4 Then GoTo 99
  '
  NRO& = CVS(Left$(YYY$, 4))
  If NRO& > 0 Then
    Call CIEOMAN(NRO&)
  End If
  '
99 Screen.MousePointer = 1
  Command46.Enabled = True
  '
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     openforms = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
End Sub
Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub


'
Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
   Call BAJALDISCO
End Sub

