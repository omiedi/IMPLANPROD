VERSION 5.00
Begin VB.Form MRECHLIN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sub-Sistema de Rechazos en Línea"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5970
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "MRECHLIN.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command45 
      Caption         =   "Anul."
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
      Left            =   5025
      Picture         =   "MRECHLIN.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Anulación de Rechazo en Línea"
      Top             =   3050
      Width           =   750
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
      Height          =   700
      Left            =   5025
      Picture         =   "MRECHLIN.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Configuración de Procesos y Formularios"
      Top             =   1640
      Width           =   750
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
      Left            =   5025
      Picture         =   "MRECHLIN.frx":1016
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Listado Rechazos en Línea"
      Top             =   2340
      Width           =   750
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
      Height          =   700
      Left            =   5025
      Picture         =   "MRECHLIN.frx":1320
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   940
      Width           =   750
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
      Height          =   700
      Left            =   5025
      Picture         =   "MRECHLIN.frx":162A
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   240
      Width           =   750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "IMPRESIÓN"
      Height          =   3975
      Left            =   4920
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Label Label16 
         BackColor       =   &H00C0FFC0&
         Caption         =   "IMPORTE"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFC0&
         Caption         =   "PRECIO"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "NUMERO"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "REFERENCIA"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "MOTIVO "
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "PROVEEDOR"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "DESTINO"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "CANTIDAD"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "COD MAT"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "FECHA"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command40 
      Caption         =   "Command13"
      Enabled         =   0   'False
      Height          =   330
      Left            =   9510
      TabIndex        =   29
      Top             =   4935
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   5025
      ScaleHeight     =   75
      ScaleWidth      =   705
      TabIndex        =   27
      Top             =   4470
      Width           =   765
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Archivos Maestros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   210
      TabIndex        =   9
      ToolTipText     =   "Base de Datos de Clientes en Cuenta Corriente"
      Top             =   105
      Width           =   4635
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2415
         ScaleHeight     =   585
         ScaleWidth      =   1845
         TabIndex        =   30
         Top             =   525
         Width           =   1905
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
            Picture         =   "MRECHLIN.frx":1774
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Base de Datos de Proveedores y Subcontratistas"
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Padrón de Proveedores"
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
            TabIndex        =   31
            Top             =   105
            Width           =   1380
         End
      End
      Begin VB.PictureBox Picture13 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   290
         ScaleHeight     =   585
         ScaleWidth      =   1845
         TabIndex        =   32
         Top             =   525
         Width           =   1905
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
            Picture         =   "MRECHLIN.frx":1A7E
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "A-B-M-L Base de Datos de Items de Stock"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Maestro de Artículos"
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
            TabIndex        =   33
            Top             =   105
            Width           =   1065
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
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
         TabIndex        =   10
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command35 
            Caption         =   "Inventarios"
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
            Left            =   210
            TabIndex        =   16
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command34 
            Caption         =   "Negativos"
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
            Left            =   2205
            TabIndex        =   15
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command33 
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
            Height          =   435
            Left            =   210
            TabIndex        =   14
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command30 
            Caption         =   "Críticos"
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
            Left            =   2205
            TabIndex        =   13
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command28 
            Caption         =   "Estadísticas"
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
            Left            =   210
            TabIndex        =   12
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command27 
            Caption         =   "Análisis A-B-C"
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
            Left            =   2205
            TabIndex        =   11
            Top             =   1890
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Rechazos en Línea"
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
      Left            =   210
      TabIndex        =   1
      Top             =   1575
      Width           =   4635
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2400
         ScaleHeight     =   585
         ScaleWidth      =   1845
         TabIndex        =   56
         Top             =   420
         Width           =   1905
         Begin VB.CommandButton Command24 
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
            Picture         =   "MRECHLIN.frx":1D88
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Transferencias de Stock"
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
            Left            =   3675
            Picture         =   "MRECHLIN.frx":2092
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Orden de Disposicion de Tintorería"
            Top             =   0
            Width           =   530
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Remito de Devolución"
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
            Left            =   720
            TabIndex        =   58
            Top             =   120
            Width           =   1065
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   290
         ScaleHeight     =   585
         ScaleWidth      =   1845
         TabIndex        =   17
         Top             =   420
         Width           =   1905
         Begin VB.CommandButton Command75 
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
            Picture         =   "MRECHLIN.frx":21DC
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Rechazo en Línea"
            Top             =   0
            Width           =   560
         End
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
            Left            =   3675
            Picture         =   "MRECHLIN.frx":24E6
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Orden de Disposicion de Tintorería"
            Top             =   0
            Width           =   530
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Procesa Rechazo"
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
            Left            =   720
            TabIndex        =   38
            Top             =   120
            Width           =   1065
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
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
         TabIndex        =   2
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command23 
            Caption         =   "Inventarios"
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
            Left            =   210
            TabIndex        =   8
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Negativos"
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
            Left            =   2205
            TabIndex        =   7
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command21 
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
            Height          =   435
            Left            =   210
            TabIndex        =   6
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Críticos"
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
            Left            =   2205
            TabIndex        =   5
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Estadísticas"
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
            Left            =   210
            TabIndex        =   4
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Análisis A-B-C"
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
            Left            =   2205
            TabIndex        =   3
            Top             =   1890
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0C0&
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
      Height          =   2325
      Left            =   210
      TabIndex        =   0
      Top             =   2940
      Width           =   4635
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0C0&
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
         Height          =   720
         Left            =   420
         TabIndex        =   22
         Top             =   1365
         Width           =   3795
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0E0C0&
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
            Left            =   2205
            TabIndex        =   24
            Top             =   320
            Width           =   1170
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0C0&
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
            Left            =   525
            TabIndex        =   23
            Top             =   320
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   3735
         TabIndex        =   18
         Top             =   525
         Width           =   3795
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
            Left            =   3220
            Picture         =   "MRECHLIN.frx":2630
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Ver y Re-Imprimir Copia de Rechazos Emitidos."
            Top             =   0
            Width           =   530
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
            Left            =   2020
            Picture         =   "MRECHLIN.frx":293A
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Rechazos por Proveedor"
            Top             =   0
            Width           =   600
         End
         Begin VB.CommandButton Command63 
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
            Left            =   2620
            Picture         =   "MRECHLIN.frx":2C44
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Rechazos por Pieza"
            Top             =   0
            Width           =   600
         End
         Begin VB.CommandButton Command61 
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
            Picture         =   "MRECHLIN.frx":2F4E
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Rechazos por Motivo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "B-D - Rechazos en Línea "
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
            TabIndex        =   20
            Top             =   105
            Width           =   1590
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   225
      Left            =   9660
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   225
      Left            =   8925
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4425
      Picture         =   "MRECHLIN.frx":3258
      Top             =   4755
      Width           =   2010
   End
End
Attribute VB_Name = "MRECHLIN"
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
  Command10.Enabled = False
15 If Option2 Then
     Call LISREPRO
   ElseIf Option1 Then
     Call MENSERR(24, "Listado no Disponible por Pantalla")
     Option2 = 1
     GoTo 15
  End If
99 Command10.Enabled = True
  '
End Sub


Private Sub Command13_Click()
   Command13.Enabled = False
   OPTEROPER.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command12_Click()
  '
  RTA1% = COMALTER%("Opciones de Listados:\\Por Motivo\Por Proveedor\Por Pieza")
  Select Case RTA1%
    Case 1
      Call Command61_Click
    Case 2
      Call Command10_Click
    Case 3
      Call Command63_Click
    Case Else
      Exit Sub
  End Select
  '
End Sub

Private Sub Command14_Click()
    OPRECHLIN.Show 1
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    Call CONECRUN("AMAPRO02", "Padrón de Proveedores")
    Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call CONECRUN("AMASTO02", "Maestro de Items de Stock")
    Command2.Enabled = True
End Sub


Private Sub Command24_Click()
   '
   Dim CIMOV%(32767) 'VECTOR QUE VERIFICA SI TUVIERON MOVIMIENTO DE RECHAZO ESE PROVEEDOR
   '
   '1º PASO : LECTURA DEPÓSITO ASOCIADO A MOTIVO DE DEVOLUCIÓN
10 For J& = 1 To ULTREG&("TADESTI")
     X$ = REGLEIDO$("TADESTI", J&)
     CODI1% = Asc(Mid$(X$, 1, 1))
     If CODI1% = 1 Then 'COMO ES TABLA CERRADA 1=DEVOLUCION
        DEPODEV% = Asc(Mid$(X$, 32, 1))
        Exit For
     End If
   Next J&
   '
   JJ% = 0
   Do Until JJ% = 32767
    JJ% = JJ% + 1
    CIMOV%(JJ%) = 0 'PONGO LA MARCA EN 0
   Loop
   '
   Command24.Enabled = False
   '2º PASO: SELECCIÓN DEL PROVEEDOR.
   Call SELECHO("PROVED1")
   NPRO% = Val(VALACT1$("PROVED1"))
   If NPRO% < 1 Then GoTo 99
   '3º PASO LECUTRA DE LOS RECHAZOS DE ESE PROV. Y PONE MARCA A LOS MATERIALES INVOLUCRADOS
   Call FILTERFILE("RECHLIN", "!RECHLIN", 5, MKI$(NPRO%))
   '
   If ULTREG&("!RECHLIN") <= 0 Then
    Call MENSERR(24, "No Existen Rechazos para el Proveedor Elegido.")
    GoTo 10
   End If
   '
   For J& = 1 To ULTREG&("!RECHLIN")
     X$ = REGLEIDO$("!RECHLIN", J&)
     PRORECH% = CVI(Mid$(X$, 9, 2))
     CI1% = CVI(Mid$(X$, 5, 2))
     CIMOV%(CI1%) = 1
   Next J&
   '4º PASO LEER STOCK EN DEPOSITO DE DEVOLUCIONES PARA LOS ITEMS MARCADOS Y
   'PONERLOS EN EL ARCHIVO DEL REMITO (BOLREDET)
   Call BLANARCH("!BOLREDET")
   For JJ% = 1 To 32767
    If CIMOV%(JJ%) = 1 Then 'SI TUVO MOVIMIENTOS DE RECHAZO EN LÍNEA
      NUOC& = 0
      CANTI1# = STODEPOS(JJ%, DEPODEV%)
      If CANTI1# > 0.005 Then 'SI TIENE EXISTENCIAS EN DEPOSITO DE DEVOLUCIONES
        I& = I& + 1
        Y$ = REGBLAN$("!BOLREDET")
        Call REPLA(Y$, MKI$(JJ%), 83, 2)
        Call REPLA(Y$, UNIMED$(JJ%), 85, 4)
        Call REPLA(Y$, MKD$(CANTI1#), 89, 8)
        Call REPLA(Y$, MKS$(NUOC&), 151, 4)
        Call REPLA(Y$, MKS$(I&), 189, 4)
        Call REGAPP("!BOLREDET", Y$)
      End If
    End If
    If JJ% = 32767 Then Exit For
   Next JJ%
   Call CIERRARCH("!BOLREDET")
   '

20 VTB% = VENTABU%("ECONSIG/NC/TITUPAN=Devolución - " + TRIM$(RASOCLI$(NPRO% * (-1))))
   If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
      Call BLANARCH("!BOLREDET")
      GoTo 10
   End If
   For YX& = 1 To ULTREG&("!BOLREDET")
     YY$ = REGLEIDO$("!BOLREDET", YX&)
     CI% = CVI(Mid$(YY$, 83, 2))
     CAN# = CVD(Mid$(YY$, 89, 8))
     If CI% > 0 Then
       If CI% <= NUMAS% Then
         If CAN# > 0.005 Then
           GoTo 30
         End If
       End If
     End If
   Next YX&
   Call MENSERR(24, "Devolución Cancelada !!!\Faltan Cantidades o Códigos.")
   GoTo 20
   'MUEVE STOCK (SALIDA DE LOS MATERIALES DEL DEPOSITO)
30 CODRECHA$ = "RD"
   FIN& = ULTREG&("MOVISTO")
   For U& = FIN& To 1 Step -1
       Call TRACE(20, FIN& - U& + 1, FIN&)
       XX$ = REGLEIDO$("MOVISTO", U&)
       If UCase$(Mid$(XX$, 21, 2)) = CODRECHA$ Then
         UNRORECH& = Val(Mid$(XX$, 26, 7))
         GoTo 40
       End If
   Next U&
   UNRORECH& = 0
   '
40 REMIRECH$ = CODRECHA$ + "-" + TRIM$(Str$(UNRORECH& + 1))
   '
   DEPX% = DEPODEV%
   FEX = HOY(HO$)
   FF% = FEX
   NC1% = NPRO%
   '
   For YX& = 1 To ULTREG&("!BOLREDET")
     YY$ = REGLEIDO$("!BOLREDET", YX&)
     CI% = CVI(Mid$(YY$, 83, 2))
     CAN# = CVD(Mid$(YY$, 89, 8))
     CANTI = CAN# * (-1)
     REFE$ = "Devolución de Materiales"
     If CI% > 0 Then
       If CI% <= NUMAS% Then
         If CAN# > 0.005 Then
           'MUEVE STOCK
           Call MOVISTO(FF%, CI%, LOTE$, CODRECHA$, REMIRECH$, DOSEC$, REFE$, VUNI#, MONE$, NC1%, DEPX%, CANTI)
         End If
       End If
     End If
   Next YX&
   Call CIERRARCH("MOVISTO")
   'IMPRIME DEVOLUCION
   Call REMIRECHLIN(NC1%, REMIRECH$)
   '
99 Command24.Enabled = True
   '
End Sub

Sub REMIRECHLIN(NPROVE%, REMIRECH$) 'IMPRIME DEVOLUCION
   '
   FORIPRE$ = TRIM$(CONTROL$("", "REMIRECH"))
   If FORIPRE$ = "" Then Exit Sub
   '
   If FORIPRE$ <> "" Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
         RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
         Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito por Rechazo:\  \'" + RMCC$ + "'.")
         Exit Sub
      End If
   End If
   '
   Screen.MousePointer = 11
   '
   NC% = Abs(NPROVE%)
   DESTIDOC% = (-1) * NC%
   '
   FEX = HOY(HOS$)
   FECHDOC$ = FECHATEX$(FEX)
   NUMEDOC$ = REMIRECH$
   TIPODOC$ = "Remito de Devolución"
   '
   DTP$ = OBJPLA$("TRANSPOR", DATRANSPO$)
   DATRANSPO$ = DTP$
   '
   CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(DTP$, 1, 32))
   CODPARAM$(2) = "DOMI-TRA": DOCPARAM$(2) = TRIM$(Mid$(DTP$, 33, 32))
   CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(DTP$, 65, 16))
   CODPARAM$(4) = "PATE-TRA": DOCPARAM$(4) = TRIM$(Mid$(DTP$, 81, 16))
   CODPARAM$(5) = "COND-TRA": DOCPARAM$(5) = TRIM$(Mid$(DTP$, 97, 32))
   CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(DTP$, 129, 24))
   CAPARDOC% = 6
   '
   CODTABU1$(1) = "COD-MAT"
   CODTABU1$(2) = "DES-MAT"
   CODTABU1$(3) = "CANTIDAD"
   CODTABU1$(4) = "CAKILOS"
   CACOLTAB1% = 4
   '
   CAITAB1% = 0
   For YX& = 1 To ULTREG&("!BOLREDET")
      YY$ = REGLEIDO$("!BOLREDET", YX&)
      CI% = CVI(Mid$(YY$, 83, 2))
      CAN# = CVD(Mid$(YY$, 89, 8))
      LOTE$ = Mid$(YY$, 69, 12)
      If CI% > 0 Then
         If CI% <= NUMAS% Then
            If CAN# > 0.005 Then
               UNID$ = UNIMED$(CI%)
               CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
               KILO$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 12, 3, 2)
               '
               CAITAB1% = CAITAB1% + 1
               TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
               TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
               TETABU1$(3, CAITAB1%) = CANS$
               TETABU1$(4, CAITAB1%) = KILO$
            End If
         End If
      End If
   Next YX&
   '
   Call PRINTDOC("REMIRECH")   ' Remito de Devolucion
   '
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
End Sub
'Private Sub Command3_Click()
'For KKI% = 1 To 2
'Call PRIRECHALI(KKI%)
'Next KKI%
'End Sub

'\\\OT-05-0697-WAR-27/10/05///
Private Sub Command45_Click()
   Command45.Enabled = False
    '
    If DERACCE%("ANURECH", "Anulación de Rechazo en Línea") < 2 Then
      GoTo 99
    End If
    'HACE INDICE DE RECHAZOS
    Call HACEINRECH
    '
3   Call SELECHO("!INDIRECH/Indice General de Rechazos en Línea")
    NPX& = Val(VALACT1$("!INDIRECH"))
    If NPX& < 1 Then GoTo 99
    '
    NRL$ = FORMATNUM(NPX&, "F7.0")
    TTXX$ = TTXX$ + "Confirma la Anulación del\Rechazo en Línea\Número: \'" + TRIM$(NRL$) + "'"
    If COMALTER%(TTXX$ + "\\Cancelar Operación\Confirmar Anulación") = 2 Then
      Call FILTERDELETE("RECHLIN", 1, MKS$(NPX&))
      Call FRESHECHO("!INDIRECH")
    End If
99 Command45.Enabled = True
    '
End Sub
'\\\OT-05-0697-WAR-FIN///
Private Sub Command75_Click()
    Command75.Enabled = False
    '
    VDEF$ = Space(128)
10  HOII% = HOY(HOS$)
    Call REPLA(VDEF$, MKI$(HOII%), 3, 2)
    OBJ$ = OBJPLA$("RECHALINE", VDEF$)
    PA1$ = OBJ$
    '
    VDEF$ = Space(128)
    If TRIM$(PA1$) <> "" Then
      CODI1% = CVI(Mid$(PA1$, 1, 2))
      FERECH% = CVI(Mid$(PA1$, 3, 2))
      CARECHA = CVS(Mid$(PA1$, 5, 4))
      'VALIDACIONES
      If CODI1% > NUMAS% Or CODI1% <= 0 Then
        Call MENSERR(24, "Código de Articulo no Válido")
        VDEF = PA1$
        GoTo 10
      End If
      If CARECHA <= 0.005 Then
        VDEF = PA1$
        Call MENSERR(24, "Cantidad de Rechazo no Válida")
        GoTo 10
      End If
      '\\\************************************///
      Call PROCERECHA(CODI1%, FERECH%, CARECHA)
      '\\\************************************///
      Call BAJALDISCO
      'Call RECONECTA 'para setear nuevamente la impresora predeterminada de Windows
      Call FINAL("")
      Exit Sub
    Else
      GoTo 99
    End If
99  Command75.Enabled = True
    '
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
    'HACE INDICE DE RECHAZOS
    Call HACEINRECH
    '
3   Call SELECHO("!INDIRECH/Indice General de Rechazos en Línea")
    NPX& = Val(VALACT1$("!INDIRECH"))
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Rechazo en Línea"
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
    GoTo 99
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
    EDITFORM.Command5.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 3
    '
99  Command5.Enabled = True
End Sub

Private Sub Command64_Click()
    Command64.Enabled = False
    If Option1.Value = True Then
           Call CONECRUN("*ORF-C02", "Consulta O/F's por Producto")
        Else
           J& = 0
           Screen.MousePointer = 11
           For U& = 1 To ULTREG&("ORDEFABR")
               X$ = REGLEIDO$("ORDEFABR", U&)
               STAT% = Asc(Mid$(X$, 108, 1))
               If STAT% > 0 Then
                  If STAT% < 3 Then
                     CATOT# = CVD(Mid$(X$, 89, 8))
                     CAPRA# = CVD(Mid$(X$, 119, 8))
                     CAPEN# = CATOT# - CAPRA#
                     If CAPEN# < 0 Then CAPEN# = 0
                     Call REPLA(X$, MKD$(CAPEN#), 159, 8)
                     Call REPLA(X$, MKD$(0), 167, 8)
                     J& = J& + 1
                     Call GRAREG("ORFASEL", X$, J&)
                  End If
               End If
           Next U&
           Call SETULTREG("ORFASEL", J&)
           Call CIERRARCH("*.*")
           Screen.MousePointer = 1
           '
           Call FINAL("*SORDEFA")
    End If
    Command64.Enabled = True
End Sub

Private Sub Command8_Click()
    Call ACCDIR("ORDEFABR")
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    OPRESERVA.Show 1
    Command9.Enabled = True
End Sub

Private Sub Command61_Click()
  '
  Command61.Enabled = False
15 If Option2 Then
     Call LISRECMOT
   ElseIf Option1 Then
     Call MENSERR(24, "Listado no Disponible por Pantalla")
     Option2 = 1
     GoTo 15
  End If
99 Command61.Enabled = True
  '
End Sub

Private Sub Command63_Click()
  '
  Command63.Enabled = False
15 If Option2 Then
     Call LISREART
   ElseIf Option1 Then
     Call MENSERR(24, "Listado no Disponible por Pantalla")
     Option2 = 1
     GoTo 15
  End If
  Command63.Enabled = True
  '
End Sub

Private Sub Form_Activate()
   WindowState = 0
   OPENFORMS = DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Call GRATITFOR(Caption)
    '
    Show
    '
    'TABLA DE DESTINO DE RECHAZO
    If ULTREG&("TADESTI") < 1 Then
      Call BLANARCH("TADESTI")
      Call REGAPP("TADESTI", Chr$(0) + "Retrabajo")
      Call REGAPP("TADESTI", Chr$(1) + "Devolución")
      Call REGAPP("TADESTI", Chr$(2) + "Baja")
      Call CIERRARCH("TADESTI")
    End If
    '
    If APLINVO$ <> "" Then
      MsgBox APLINVO$
      Call Command75_Click
    End If
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
   Call BAJALDISCO
   Call FINAL("")
End Sub


Sub LISREPRO() 'LISTADO DE RECHAZOS POR PROVEEDOR
    'FILTRO DE FECHA
10  FEX = HOY(HO$)
  HAS% = FEX
  HASS$ = FECHATEX$(FEX)
  DESS$ = "01" + Mid$(HASS$, 3)
  DES% = FECHANUM(DESS$)
  VDEF$ = MKI$(DES%) + MKI$(HAS%)
  '
  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
  If OBX$ = "" Then
      GoTo 99
  End If
  '
  DES% = CVI(Left$(OBX$, 2))
  HAS% = CVI(Mid$(OBX$, 3, 2))
  If DES% > HAS% Then
      GoTo 10
  End If
  FEX = DES%: PERIO$ = FECHATEX$(FEX)
  FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
  '
  JJJ& = 0
  Call BLANARCH("LIREPRO")
  FIN& = ULTREG&("RECHLIN")
  For J& = 1 To ULTREG&("RECHLIN")
    '
    Call TRACE(20, J&, FIN&)
    '
    X$ = REGLEIDO$("RECHLIN", J&)
    NROREC& = CVS(Mid$(X$, 1, 4))
    CODI3% = CVI(Mid$(X$, 5, 2))
    FEREC3% = CVI(Mid$(X$, 7, 2))
    If FEREC3% < DES% Or FEREC3% > HAS% Then GoTo 75
    PROV3% = CVI(Mid$(X$, 9, 2))
    CANTI3 = CVS(Mid$(X$, 11, 4))
    DESTI3% = Asc(Mid$(X$, 15, 1))
    MOTIV3% = Asc(Mid$(X$, 16, 1))
    USUA3$ = TRIM$(Mid$(X$, 17, 32))
    REFE3$ = TRIM$(Mid$(X$, 49, 48))
    VALOUNI3# = CVD(Mid$(X$, 97, 8))
    IMPORT3# = CVD(Mid$(X$, 105, 8))
    '
    Y$ = Space(128)
    Call REPLA(Y$, MKS$(NROREC&), 1, 4)
    Call REPLA(Y$, MKI$(FEREC3%), 5, 2)
    Call REPLA(Y$, MKI$(PROV3%), 7, 2)
    Call REPLA(Y$, MKI$(CODI3%), 9, 2)
    Call REPLA(Y$, MKS$(CANTI3), 11, 4)
    Call REPLA(Y$, Chr(DESTI3%), 15, 1)
    Call REPLA(Y$, Chr(MOTIV3%), 16, 1)
    Call REPLA(Y$, USUA3$, 17, 32)
    Call REPLA(Y$, REFE3$, 49, 48)
    Call REPLA(Y$, MKD$(VALOUNI3#), 97, 8)
    Call REPLA(Y$, MKD$(IMPORT3#), 105, 8)
    JJJ& = JJJ& + 1
    Call GRAREG("LIREPRO", Y$, JJJ&)
    '
75 Next J&
  '
  If JJJ& <= 0 Then
    Call MENSERR(24, "No Existen Rechazos para el Rango de Fechas Elegido.")
    GoTo 10
  End If
  '
  Call SETULTREG("LIREPRO", JJJ&)
  Call CIERRARCH("LIREPRO")
  Call FILESORT("LIREPRO", "LIREPRO", 3, 3, "ASC")
  Call HEADERS("LIREPRO", "")
  Call HEADERS("LIREPRO", "Periodo: " + PERIO$)
  Call FINAL("*LIREPRO")
  '
99 End Sub

Sub LISREART() 'LISTADO DE RECHAZOS POR ARTICULOS
  '
      'FILTRO DE FECHA
10  FEX = HOY(HO$)
  HAS% = FEX
  HASS$ = FECHATEX$(FEX)
  DESS$ = "01" + Mid$(HASS$, 3)
  DES% = FECHANUM(DESS$)
  VDEF$ = MKI$(DES%) + MKI$(HAS%)
  '
  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
  If OBX$ = "" Then
      GoTo 99
  End If
  '
  DES% = CVI(Left$(OBX$, 2))
  HAS% = CVI(Mid$(OBX$, 3, 2))
  If DES% > HAS% Then
      GoTo 10
  End If
  FEX = DES%: PERIO$ = FECHATEX$(FEX)
  FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
  '
  JJJ& = 0
  Call BLANARCH("LIREPRO")
  FIN& = ULTREG&("RECHLIN")
  For J& = 1 To ULTREG&("RECHLIN")
    '
    Call TRACE(20, J&, FIN&)
    '
    X$ = REGLEIDO$("RECHLIN", J&)
    NROREC& = CVS(Mid$(X$, 1, 4))
    CODI3% = CVI(Mid$(X$, 5, 2))
    FEREC3% = CVI(Mid$(X$, 7, 2))
    If FEREC3% < DES% Or FEREC3% > HAS% Then GoTo 75
    PROV3% = CVI(Mid$(X$, 9, 2))
    CANTI3 = CVS(Mid$(X$, 11, 4))
    DESTI3% = Asc(Mid$(X$, 15, 1))
    MOTIV3% = Asc(Mid$(X$, 16, 1))
    USUA3$ = TRIM$(Mid$(X$, 17, 32))
    REFE3$ = TRIM$(Mid$(X$, 49, 48))
    VALOUNI3# = CVD(Mid$(X$, 97, 8))
    IMPORT3# = CVD(Mid$(X$, 105, 8))
    '
    Y$ = Space(128)
    Call REPLA(Y$, MKS$(NROREC&), 1, 4)
    Call REPLA(Y$, MKI$(FEREC3%), 5, 2)
    Call REPLA(Y$, MKI$(PROV3%), 7, 2)
    Call REPLA(Y$, MKI$(CODI3%), 9, 2)
    Call REPLA(Y$, MKS$(CANTI3), 11, 4)
    Call REPLA(Y$, Chr(DESTI3%), 15, 1)
    Call REPLA(Y$, Chr(MOTIV3%), 16, 1)
    Call REPLA(Y$, USUA3$, 17, 32)
    Call REPLA(Y$, REFE3$, 49, 48)
    Call REPLA(Y$, MKD$(VALOUNI3#), 97, 8)
    Call REPLA(Y$, MKD$(IMPORT3#), 105, 8)
    JJJ& = JJJ& + 1
    Call GRAREG("LIREPRO", Y$, JJJ&)
    '
75 Next J&
  '
  If JJJ& <= 0 Then
    Call MENSERR(24, "No Existen Rechazos para el Rango de Fechas Elegido.")
    GoTo 10
  End If
  '
  Call SETULTREG("LIREPRO", JJJ&)
  Call CIERRARCH("LIREPRO")
  Call FILESORT("LIREPRO", "LIREPRO", 5, 5, "ASC")
  Call HEADERS("LIREPRO", "")
  Call HEADERS("LIREPRO", "Periodo: " + PERIO$)
  Call FINAL("*LIREART")
  '
99 End Sub

Sub LISRECMOT()
 '
 'FILTRO DE FECHA
10  FEX = HOY(HO$)
  HAS% = FEX
  HASS$ = FECHATEX$(FEX)
  DESS$ = "01" + Mid$(HASS$, 3)
  DES% = FECHANUM(DESS$)
  VDEF$ = MKI$(DES%) + MKI$(HAS%)
  '
  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
  If OBX$ = "" Then
      GoTo 99
  End If
  '
  DES% = CVI(Left$(OBX$, 2))
  HAS% = CVI(Mid$(OBX$, 3, 2))
  If DES% > HAS% Then
      GoTo 10
  End If
  FEX = DES%: PERIO$ = FECHATEX$(FEX)
  FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
  '
  JJJ& = 0
  Call BLANARCH("LIREPRO")
  FIN& = ULTREG&("RECHLIN")
  For J& = 1 To ULTREG&("RECHLIN")
    '
    Call TRACE(20, J&, FIN&)
    '
    X$ = REGLEIDO$("RECHLIN", J&)
    NROREC& = CVS(Mid$(X$, 1, 4))
    CODI3% = CVI(Mid$(X$, 5, 2))
    FEREC3% = CVI(Mid$(X$, 7, 2))
    If FEREC3% < DES% Or FEREC3% > HAS% Then GoTo 75
    PROV3% = CVI(Mid$(X$, 9, 2))
    CANTI3 = CVS(Mid$(X$, 11, 4))
    DESTI3% = Asc(Mid$(X$, 15, 1))
    MOTIV3% = Asc(Mid$(X$, 16, 1))
    USUA3$ = TRIM$(Mid$(X$, 17, 32))
    REFE3$ = TRIM$(Mid$(X$, 49, 48))
    VALOUNI3# = CVD(Mid$(X$, 97, 8))
    IMPORT3# = CVD(Mid$(X$, 105, 8))
    '
    Y$ = Space(128)
    Call REPLA(Y$, MKS$(NROREC&), 1, 4)
    Call REPLA(Y$, MKI$(FEREC3%), 5, 2)
    Call REPLA(Y$, MKI$(PROV3%), 7, 2)
    Call REPLA(Y$, MKI$(CODI3%), 9, 2)
    Call REPLA(Y$, MKS$(CANTI3), 11, 4)
    Call REPLA(Y$, Chr(DESTI3%), 15, 1)
    Call REPLA(Y$, Chr(MOTIV3%), 16, 1)
    Call REPLA(Y$, USUA3$, 17, 32)
    Call REPLA(Y$, REFE3$, 49, 48)
    Call REPLA(Y$, MKD$(VALOUNI3#), 97, 8)
    Call REPLA(Y$, MKD$(IMPORT3#), 105, 8)
    JJJ& = JJJ& + 1
    Call GRAREG("LIREPRO", Y$, JJJ&)
    '
75 Next J&
  '
  If JJJ& <= 0 Then
    Call MENSERR(24, "No Existen Rechazos para el Rango de Fechas Elegido.")
    GoTo 10
  End If
  '
  Call SETULTREG("LIREPRO", JJJ&)
  Call CIERRARCH("LIREPRO")
  Call FILESORT("LIREPRO", "LIREPRO", 10, 10, "ASC")
  Call HEADERS("LIREPRO", "")
  Call HEADERS("LIREPRO", "Periodo: " + PERIO$)
  Call FINAL("*LIREMOT")
  '
99 End Sub

Sub HACEINRECH()
  '
  Call BLANARCH("!INDIRECH")
  For J& = 1 To ULTREG&("RECHLIN")
    X$ = REGLEIDO$("RECHLIN", J&)
    NROREC& = CVS(Mid$(X$, 1, 4))
    NPRO% = CVI(Mid$(X$, 9, 2))
    CIXI% = CVI(Mid$(X$, 5, 2))
    FEMI1$ = FECHATEX$(CVI(Mid$(X$, 7, 2)))
    ' A ARCHIVO INDICE DE RECHAZO
    REFERE2$ = FEMI1$ + " - " + TRIM$(CODEXT$(CIXI%)) + " - " + TRIM$(DESCRIT0$(CIXI%)) ' RASOCLI$(NPRO% * (-1))
    ZZ$ = Space(64)
    Call REPLA(ZZ$, MKS$(NROREC&), 1, 4)
    Call REPLA(ZZ$, REFERE2$, 5, 60)
    Call REGAPP("!INDIRECH", ZZ$)
  Next J&
  Call CIERRARCH("!INDIRECH")
  '
End Sub

Sub REMIDEV(NC1%, DEPX1%, CI1%, RECHA1#, NURECH$, FECHRECH%, VUNI1#)
    '
    If TRIM$(CONTROL$("", "REMIRECH")) <> "" Then
        CODRECHA$ = TRIM$(CONTROL$("", "CODRRECH"))
        If CODRECHA$ = "" Then CODRECHA$ = "RD"
        FIN& = ULTREG&("MOVISTO")
        For U& = FIN& To 1 Step -1
           Call TRACE(20, FIN& - U& + 1, FIN&)
           XX$ = REGLEIDO$("MOVISTO", U&)
           If UCase$(Mid$(XX$, 21, 2)) = CODRECHA$ Then
             UNRORECH& = Val(Mid$(XX$, 26, 7))
             GoTo 11241
           End If
        Next U&
        UNRORECH& = 0
        '
11241   REMIRECH$ = CODRECHA$ + "-" + TRIM$(Str$(UNRORECH& + 1))
        '
        DOSEC$ = NURECH$
        DEPX% = DEPX1%
        RECHA# = RECHA1#
        CANTI = RECHA# * (-1)
        REFE$ = "RECHAZO - Rto." + TRIM$(REMI$)
        FF% = FECHRECH%
        VUNI# = VUNI1#
        CI% = CI1%
        Call MOVISTO(FF%, CI%, LOTE$, CODRECHA$, REMIRECH$, DOSEC$, REFE$, VUNI#, MONE$, NC1%, DEPX%, CANTI)
        '
        Call BLANARCH("!BOLREDET")
        Z$ = REGBLAN$("BOLREDET")
        Call REPLA(Z$, MKI$(CI%), 83, 2)
        Call REPLA(Z$, MKD$(RECHA1#), 171, 8)
        Call REPLA(Z$, LOTE$, 69, 12)
        Call REGAPP("!BOLREDET", Z$)
        Call CIERRARCH("!BOLREDET")
        '
        Call REMIRECHA(NC1%, DEPX%, REMIRECH$)
        '
      End If
    '
End Sub
