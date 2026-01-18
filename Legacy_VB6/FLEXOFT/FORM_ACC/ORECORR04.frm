VERSION 5.00
Begin VB.Form ORECORR04 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Recorrida - Opciones"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFFF&
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
      Left            =   4830
      TabIndex        =   46
      Top             =   5565
      Width           =   3585
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   48
         Top             =   390
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   47
         Top             =   390
         Value           =   -1  'True
         Width           =   1380
      End
   End
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
      Picture         =   "ORECORR04.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Revertir Anulaciones de Ordenes de Compra"
      Top             =   4305
      Visible         =   0   'False
      Width           =   560
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gestión de O/R's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   160
      TabIndex        =   29
      Top             =   120
      Width           =   4005
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   59
         Top             =   2780
         Width           =   2640
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
            Height          =   580
            Left            =   0
            Picture         =   "ORECORR04.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Configuración de Procesos y Formularios"
            Top             =   0
            Width           =   650
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Configuracion de Funcionamiento"
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
            TabIndex        =   60
            Top             =   105
            Width           =   1485
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ordenes de Recorrida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2190
         Left            =   240
         TabIndex        =   35
         Top             =   400
         Width           =   3585
         Begin VB.PictureBox Picture7 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   36
            Top             =   1365
            Width           =   2325
            Begin VB.CommandButton Command51 
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
               Left            =   1720
               Picture         =   "ORECORR04.frx":058C
               Style           =   1  'Graphical
               TabIndex        =   63
               ToolTipText     =   "Anotador de Orden de Recorrida"
               Top             =   0
               Width           =   560
            End
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
               Picture         =   "ORECORR04.frx":09CE
               Style           =   1  'Graphical
               TabIndex        =   44
               ToolTipText     =   "Cierre de Orden de Mantenimiento"
               Top             =   0
               Width           =   530
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Cierre O. de Recorrida"
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
               Left            =   600
               TabIndex        =   37
               Top             =   60
               Width           =   1095
            End
         End
         Begin VB.PictureBox Picture9 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   38
            Top             =   840
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
               Picture         =   "ORECORR04.frx":0CD8
               Style           =   1  'Graphical
               TabIndex        =   42
               ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Ver O/R y Re-Imprimir"
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
               Left            =   945
               TabIndex        =   39
               Top             =   60
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture15 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   40
            Top             =   315
            Width           =   2325
            Begin VB.CommandButton Command13 
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   0
               Picture         =   "ORECORR04.frx":0FE2
               Style           =   1  'Graphical
               TabIndex        =   62
               ToolTipText     =   "Emite e Imprime Orden de Recorrida"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Emitir Orden de Recorrida"
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
               TabIndex        =   41
               Top             =   60
               Width           =   1300
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
         TabIndex        =   30
         Top             =   4000
         Visible         =   0   'False
         Width           =   3585
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2895
            TabIndex        =   31
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
               Picture         =   "ORECORR04.frx":164C
               Style           =   1  'Graphical
               TabIndex        =   33
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
               Picture         =   "ORECORR04.frx":1796
               Style           =   1  'Graphical
               TabIndex        =   32
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
               TabIndex        =   34
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
      TabIndex        =   18
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
         TabIndex        =   22
         Top             =   1560
         Width           =   3585
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   26
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
               Picture         =   "ORECORR04.frx":1AA0
               Style           =   1  'Graphical
               TabIndex        =   27
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
               TabIndex        =   28
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
            TabIndex        =   23
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
               Picture         =   "ORECORR04.frx":1DAA
               Style           =   1  'Graphical
               TabIndex        =   24
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
               TabIndex        =   25
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
         TabIndex        =   19
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
            Picture         =   "ORECORR04.frx":1EF4
            Style           =   1  'Graphical
            TabIndex        =   20
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
            TabIndex        =   21
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
      TabIndex        =   13
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   140
         ScaleHeight     =   75
         ScaleWidth      =   495
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2940
         Width           =   555
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -105
            TabIndex        =   17
            Top             =   0
            Width           =   12000
         End
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
         Picture         =   "ORECORR04.frx":21FE
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "ORECORR04.frx":2508
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   215
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -255
         Picture         =   "ORECORR04.frx":2652
         Stretch         =   -1  'True
         Top             =   3300
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Configuración de Recorridos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   4005
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Secuencia de Tareas por Recorrido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   210
         TabIndex        =   49
         Top             =   2205
         Width           =   3585
         Begin VB.PictureBox Picture17 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   500
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   56
            Top             =   400
            Width           =   2325
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
               Picture         =   "ORECORR04.frx":4574
               Style           =   1  'Graphical
               TabIndex        =   57
               ToolTipText     =   "Definición de Secuencia de Tareas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Altas, Bajas y Modificaciones"
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
               TabIndex        =   58
               Top             =   60
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture14 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   53
            Top             =   2400
            Visible         =   0   'False
            Width           =   2325
            Begin VB.CommandButton Command8 
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
               Picture         =   "ORECORR04.frx":46BE
               Style           =   1  'Graphical
               TabIndex        =   54
               ToolTipText     =   "Pendientes por Material"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
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
               TabIndex        =   55
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture13 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   50
            Top             =   2640
            Visible         =   0   'False
            Width           =   2325
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
               Picture         =   "ORECORR04.frx":49C8
               Style           =   1  'Graphical
               TabIndex        =   51
               ToolTipText     =   "Material en Consignación"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label8 
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
               TabIndex        =   52
               Top             =   105
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Tabla de Recorridos"
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
               Picture         =   "ORECORR04.frx":4CD2
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
               Picture         =   "ORECORR04.frx":4FDC
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
               Picture         =   "ORECORR04.frx":52E6
               Style           =   1  'Graphical
               TabIndex        =   43
               ToolTipText     =   "Ordenes de Mantenimiento por Máquinas\Equipos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Listado Tabla de Recorridos"
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
               Width           =   1395
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
               Picture         =   "ORECORR04.frx":55F0
               Style           =   1  'Graphical
               TabIndex        =   11
               ToolTipText     =   "Ordenes de Mantenimiento por Número"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Altas, Bajas y Modificaciones"
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
Attribute VB_Name = "ORECORR04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HAYMATERIALIB%

'*****/////COMANDO OCULTO/////*****
'Se oculto en el frame8 ("Opciones de Consulta")
'el Command9("Listado de Material en Consignación)"
'*****/////**************/////*****

Private Sub Command1_Click()
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
   Call CommSel("TARECORR", Command13, SETAREC04.Text4)
   If TRIM$(SETAREC04.Text4) <> "" Then
      Call SETAREC04.Command25_Click
   End If
   '
99 Command13.Enabled = True
   '
End Sub
Sub MODIFOM(NRO&)
    '
    Call ABREOMAN
    '
    With OrdeMant
      .FindFirst " NumeOrMa_OM = " & NRO& & ""
      If .NoMatch = False Then
        MOMANTEN04.Text1.TEXT = !Codigo_Equ
        MOMANTEN04.Text24.TEXT = FORMATNUM$(CVINT(!FechGen_OM), "D8")
        NOM1& = !NumeOrMa_OM
        MOMANTEN04.Label22 = FORMATNUM$(NOM1&, "F9.0")
        MOMANTEN04.Text28 = !Descri_OM
        MOMANTEN04.Text2 = FORMATNUM$(CVINT(!FeProCom_OM), "D8")
        MOMANTEN04.Text5 = Format$(!HoProCom_OM, "hh:mm")
        MOMANTEN04.Text4 = FORMATNUM$(CVINT(!FeProFin_OM), "D8")
        MOMANTEN04.Text6 = Format$(!HoProFin_OM, "hh:mm")
        MOMANTEN04.Text3 = !Proved_OM
        MOMANTEN04.Text7 = !Referen_OM
        COSTER1 = !CostoTer_OM
        MOMANTEN04.Text8.TEXT = FORMATNUM$(COSTER1, "F9.2")
      End If
    End With
    'CARGO LOS OPERARIOS
    MOMANTEN04.List4.Clear
    With RecuProp
      SQLX$ = " NumeOrMa_OM = " & NRO& & ""
      .FindFirst SQLX$
10    If .NoMatch = False Then
          LEG1% = !LegaEmp_OM
          LEG2$ = FORMATNUM$(LEG1%, "F5.0")
          EMP1$ = ECOARCH$("PADRON", MKI$(LEG1%))
          'AGREGO AL LIST
          L1$ = Space(128)
          Call REPLA(L1$, LEG2$, 1, 8)
          Call REPLA(L1$, EMP1$, 10, 20)
          MOMANTEN04.List4.AddItem L1$
        .FindNext SQLX$
        GoTo 10
      End If
    End With
    'CARGO LAS RESERVAS
    BLANARCH ("!REPUESTO")
    '
    Call ABRESERVA
    With Reserva
      SQLX$ = " VAL(MID(IdSubOr,4,6)) = " & NRO& & ""
      .FindFirst SQLX$
      JJ& = 0
15    If .NoMatch = False Then
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
           MOMANTEN04.List1.AddItem L1$
           'AGREGO AL ARCHIVO
           Y2$ = REGBLAN$("REPUESTO")
           Call REPLA(Y2$, MKI$(CODI1%), 1, 2)
           Call REPLA(Y2$, UM2$, 3, 2)
           Call REPLA(Y2$, MKS(USO1), 5, 4)
           Call REPLA(Y2$, OBS2$, 9, 25)
           JJ& = JJ& + 1
           Call GRAREG("!REPUESTO", Y2$, JJ&)
        .FindNext SQLX$
        GoTo 15
      End If
    End With
  Call SETULTREG("!REPUESTO", JJ&)
  Call CIERRARCH("!REPUESTO")
  MOMANTEN04.Show 1
  '
End Sub
Sub CIEOMAN(NRO&)
 '
    Call ABREOMAN
    '
    With OrdeMant
      .FindFirst " NumeOrMa_OM = " & NRO& & ""
      If .NoMatch = False Then
        CIMANTEN04.Text1.TEXT = !Codigo_Equ
        CIMANTEN04.Text24.TEXT = FORMATNUM$(CVINT(!FechGen_OM), "D8")
        NOM1& = !NumeOrMa_OM
        CIMANTEN04.Label22 = FORMATNUM$(NOM1&, "F9.0")
        CIMANTEN04.Text28 = !Descri_OM
        CIMANTEN04.Text2 = FORMATNUM$(CVINT(!FeProCom_OM), "D8")
        CIMANTEN04.Text5 = Format$(!HoProCom_OM, "hh:mm")
        CIMANTEN04.Text4 = FORMATNUM$(CVINT(!FeProFin_OM), "D8")
        CIMANTEN04.Text6 = Format$(!HoProFin_OM, "hh:mm")
        CIMANTEN04.Text3 = !Proved_OM
        CIMANTEN04.Text7 = !Referen_OM
      End If
    End With
    'CARGO LOS OPERARIOS
    CIMANTEN04.List4.Clear
    With RecuProp
      SQLX$ = " NumeOrMa_OM = " & NRO& & ""
      .FindFirst SQLX$
10    If .NoMatch = False Then
          LEG1% = !LegaEmp_OM
          LEG2$ = FORMATNUM$(LEG1%, "F5.0")
          EMP1$ = ECOARCH$("PADRON", MKI$(LEG1%))
          'AGREGO AL LIST
          L1$ = Space(128)
          Call REPLA(L1$, LEG2$, 1, 8)
          Call REPLA(L1$, EMP1$, 10, 20)
          CIMANTEN04.List4.AddItem L1$
        .FindNext SQLX$
        GoTo 10
      End If
    End With
    'CARGO LAS RESERVAS
    BLANARCH ("!REPUESTO")
    '
    Call ABRESERVA
    With Reserva
      SQLX$ = " VAL(MID(IdSubOr,4,6)) = " & NRO& & ""
      .FindFirst SQLX$
      JJ& = 0
15    If .NoMatch = False Then
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
           CIMANTEN04.List1.AddItem L1$
           'AGREGO AL ARCHIVO
           Y2$ = REGBLAN$("REPUESTO")
           Call REPLA(Y2$, MKI$(CODI1%), 1, 2)
           Call REPLA(Y2$, UM2$, 3, 2)
           Call REPLA(Y2$, MKS(USO1), 5, 4)
           Call REPLA(Y2$, OBS2$, 9, 25)
           JJ& = JJ& + 1
           Call GRAREG("!REPUESTO", Y2$, JJ&)
        .FindNext SQLX$
        GoTo 15
      End If
    End With
  Call SETULTREG("!REPUESTO", JJ&)
  Call CIERRARCH("!REPUESTO")
  '
  CIMANTEN04.List3.Clear
  Call CIMANTEN04.CARCONTA(CIMANTEN04.Text1.TEXT)
  '
  CIMANTEN04.Show 1
  '
End Sub

Private Sub Command16_Click()
   Command16.Enabled = False
   SETAREC04.Show 1
   Command16.Enabled = True
End Sub

Private Sub Command18_Click()
    '
    Command18.Enabled = False
    OPRECOR04.Show 1
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
    FOAPROC04.Show 1
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
   Call FRESHECHO("INDIOREC")
3  YYY$ = REGSEL$("INDIOREC/Indice General de Ordenes de Recorrida")
   '
   If Len(YYY$) <= 4 Then GoTo 99
   '
   NPX& = CVS(Left$(YYY$, 4))
   '
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Orden de Recorrida"
    NUDOCU$ = TRIM$(Str$(NPX&))
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
5   Call MUESTRADOC(DOCUNU&)
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
    Call TRALOCAL("TARECORR", "")
    VTB% = VENTABU%("TARECORR")
    If VTB% >= 0 Then
       Call TRACENTRAL("TARECORR", "")
       Call FILESORT("TARECORR", "", 1, 1)
    End If
    Command3.Enabled = True
    '
End Sub
'
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
   DES% = CVI(Left$(OBX$, 2)): DES1 = CDbl(CVDAT(DES%))
   HAS% = CVI(Mid$(OBX$, 3, 2)): HAS1 = CDbl(CVDAT(HAS%))
   If HAS% < DES% Then GoTo 10
   FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   Call AMAEQUI04.INDIEQUI
   Call SELECHO("INDIEQUI/Indice de Equipos")
   MAQ$ = TRIM$(VALACT1$("INDIEQUI"))
   '
   Screen.MousePointer = 11
   '
   Call ABREOMAN
   '
   SQLX$ = "SELECT * FROM ORDEMANT "
   SQLX$ = SQLX$ + " WHERE CDBL(FechGen_OM) >= " & DES1 & " "
   SQLX$ = SQLX$ + " AND CDBL(FechGen_OM) <=   " & HAS1 & "  "
   SQLX$ = SQLX$ + " AND Status_OM = 1 "
   If TRIM$(MAQ$) <> "" Then
     SQLX$ = SQLX$ + " AND Codigo_Equ = '" & MAQ$ & "' "
   End If
   SQLX$ = SQLX$ + " ORDER BY NumeOrMa_OM ASC "
   Set OMTMP = Omantin.OpenRecordset(SQLX$, 4)
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
      Call REPLA(Y$, !Codigo_Equ, 96, 6)
      Call REPLA(Y$, MKI$(CVINT%(!FechCie_OM)), 102, 2)
          HOCIE$ = Format$(!HoraCie_OM, "hh:mm")
      Call REPLA(Y$, MKI$(HORANUM(HOCIE$)), 104, 2)
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
      Call REPLA(Y$, MKI$(HORANUM(HOCOPRO$)), 145, 2)
      Call REPLA(Y$, MKI$(CVINT%(!FeProFin_OM)), 147, 2)
        HOFIPRO$ = Format$(!HoProFin_OM, "hh:mm")
      Call REPLA(Y$, MKI$(HORANUM(HOFIPRO$)), 149, 2)
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
   DES% = CVI(Left$(OBX$, 2)): DES1 = CDbl(CVDAT(DES%))
   HAS% = CVI(Mid$(OBX$, 3, 2)): HAS1 = CDbl(CVDAT(HAS%))
   If HAS% < DES% Then GoTo 10
   FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   Screen.MousePointer = 11
   '
   Call ABREOMAN
   '
   SQLX$ = "SELECT * FROM ORDEMANT "
   SQLX$ = SQLX$ + " WHERE CDBL(FechGen_OM) >= " & DES1 & " "
   SQLX$ = SQLX$ + " AND CDBL(FechGen_OM) <=   " & HAS1 & "  "
   SQLX$ = SQLX$ + " AND Status_OM = 1 "
   SQLX$ = SQLX$ + " ORDER BY NumeOrMa_OM ASC "
   Set OMTMP = Omantin.OpenRecordset(SQLX$, 4)
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
      Call REPLA(Y$, !Codigo_Equ, 96, 6)
      Call REPLA(Y$, MKI$(CVINT%(!FechCie_OM)), 102, 2)
        HOCIE$ = Format$(!HoraCie_OM, "hh:mm")
      Call REPLA(Y$, MKI$(HORANUM(HOCIE$)), 104, 2)
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
      Call REPLA(Y$, MKI$(HORANUM(HOCOPRO$)), 145, 2)
      Call REPLA(Y$, MKI$(CVINT%(!FeProFin_OM)), 147, 2)
        HOFIPRO$ = Format$(!HoProFin_OM, "hh:mm")
      Call REPLA(Y$, MKI$(HORANUM(HOFIPRO$)), 149, 2)
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
    Call CONECRUN("*LITAREC", "Listado de Tabla de Recorridas")
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
   Set OMTMP = Omantin.OpenRecordset(SQLX$, 4)
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
      DES1$ = !Codigo_Equ
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
    With OrdeMant
      .FindFirst " NumeOrMa_OM = " & NRO& & " "
      If .NoMatch = False Then
        .Edit
          !Status_OM = -1
        .Update
        With Reserva
          SQLX$ = " VAL(MID(IdSubOr,4,6)) = " & NRO& & ""
          .FindFirst SQLX$
15        If .NoMatch = False Then
            .Edit 'CANCELO RESERVAS
              !CantRes = !CantCons
              !Saldo_Entre = !CantRes - !CantCons
            .Update
            .FindNext SQLX$
            GoTo 15
          End If
        End With
      End If
    End With
    GoTo 12
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
   If COMALTER%("¿Realmente Desea\CERRAR Ordenes de Recorrida ?\\No, Abandonar\Si, Cerrar O/M's") <> 2 Then
     GoTo 99
   End If
   '
   Call FRESHECHO("INDIOREC")
3  YYY$ = REGSEL$("INDIOREC/Indice General de Ordenes de Recorrida")
   '
   If Len(YYY$) >= 4 Then
      NPX& = CVS(Left$(YYY$, 4))
      If NPX& > 0 Then
         'Call COMUNI("No Disponible")
         Call CIEORDRECO(NPX&)
         ANOTARECO04.Label7 = Str(NPX&)
         ANOTARECO04.Show 1
      End If
   End If
   '
99 Command46.Enabled = True
   '
End Sub

Private Sub Command51_Click()
'
ANOTARECO04.Label7 = ""
ANOTARECO04.Show 1
'
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
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

Sub CIEORDRECO(NUMEOR&)
'
Call ABREOMAN
'
' CREAR PRIMERA NOTA CORRESPONDIENTE AL CIERRE DE LA ORDEN DE RECORRIDA
TTX1$ = Space$(76)
Call REPLA(TTX1$, FECHATEX(HOY(HOS$)), 1, 8)
  TTXX$ = "Cierre de Orden de Recorrida Nro. " + TRIM$(FORMATNUM(NUMEOR&, "F9"))
Call REPLA(TTX1$, TTXX$, 11, 56)
Call REPLA(TTX1$, USERTER$, 69, 24)
TTX1$ = TTX1$ + Chr$(10) + Chr$(13)
TTX2$ = String$(88, "=") + Chr$(10) + Chr$(13)
ANOTA$ = TTX1$ + TTX2$

SQLX$ = "SELECT * FROM ORECORRI WHERE NumOrec= " & NUMEOR&
Set ORRE = Omantin.OpenRecordset(SQLX$, dbOpenDynaset)
With ORRE
  FCIE = CVINT(!FechCierr)
  If FCIE <> 0 Then ' VALIDA QUE LA ORDEN NO SE HAYA CERRADO PREVIAMENTE
    Call MENSERR(24, "Orden de Recorrida Nro. " & TRIM$(FORMATNUM(NUMEOR&, "F9")) & " Cerrada")
    Exit Sub
  End If
  .Edit
  !FechCierr = CVDAT(HOY(HO$))
  !ObseCierr = ANOTA$
  .Update
End With
'
End Sub
