VERSION 5.00
Begin VB.Form MRECHLIN07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sub-Sistema de Rechazos en Línea"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6060
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "MRECHLIN07.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command40 
      Caption         =   "Command13"
      Enabled         =   0   'False
      Height          =   330
      Left            =   9510
      TabIndex        =   27
      Top             =   4935
      Visible         =   0   'False
      Width           =   750
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
      Left            =   120
      TabIndex        =   9
      Top             =   105
      Width           =   4635
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2415
         ScaleHeight     =   585
         ScaleWidth      =   1845
         TabIndex        =   28
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
            Picture         =   "MRECHLIN07.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   33
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
            TabIndex        =   29
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
         TabIndex        =   30
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
            Picture         =   "MRECHLIN07.frx":0BD4
            Style           =   1  'Graphical
            TabIndex        =   32
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
            TabIndex        =   31
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
      Left            =   120
      TabIndex        =   1
      Top             =   1575
      Width           =   4635
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2400
         ScaleHeight     =   585
         ScaleWidth      =   1845
         TabIndex        =   39
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
            Picture         =   "MRECHLIN07.frx":0EDE
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Remito de Devolución"
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
            Picture         =   "MRECHLIN07.frx":11E8
            Style           =   1  'Graphical
            TabIndex        =   40
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
            TabIndex        =   41
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
            Picture         =   "MRECHLIN07.frx":1332
            Style           =   1  'Graphical
            TabIndex        =   38
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
            Picture         =   "MRECHLIN07.frx":163C
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
            TabIndex        =   36
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
      Height          =   2220
      Left            =   120
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
            Value           =   -1  'True
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
            Picture         =   "MRECHLIN07.frx":1786
            Style           =   1  'Graphical
            TabIndex        =   37
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
            Picture         =   "MRECHLIN07.frx":1A90
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
            Picture         =   "MRECHLIN07.frx":1D9A
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
            Picture         =   "MRECHLIN07.frx":20A4
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
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00004000&
      Height          =   5895
      Left            =   4860
      ScaleHeight     =   5835
      ScaleWidth      =   1155
      TabIndex        =   43
      Top             =   0
      Width           =   1220
      Begin VB.CommandButton Command3 
         Caption         =   "Recup"
         Height          =   375
         Left            =   210
         TabIndex        =   62
         Top             =   4320
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   240
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   60
         Top             =   4080
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   61
            Top             =   105
            Width           =   12000
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
         Height          =   700
         Left            =   210
         Picture         =   "MRECHLIN07.frx":23AE
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   120
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
         Left            =   210
         Picture         =   "MRECHLIN07.frx":24F8
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   840
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
         Left            =   210
         Picture         =   "MRECHLIN07.frx":2802
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Listado Rechazos en Línea"
         Top             =   2280
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
         Left            =   210
         Picture         =   "MRECHLIN07.frx":2B0C
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   1560
         Width           =   750
      End
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
         Left            =   210
         Picture         =   "MRECHLIN07.frx":2F4E
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Anulación de Rechazo en Línea"
         Top             =   3000
         Width           =   750
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "IMPRESIÓN"
         Height          =   3975
         Left            =   5520
         TabIndex        =   44
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Label Label2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "FECHA"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "COD MAT"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0FFC0&
            Caption         =   "CANTIDAD"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0FFC0&
            Caption         =   "DESTINO"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0FFC0&
            Caption         =   "PROVEEDOR"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0FFC0&
            Caption         =   "MOTIVO "
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0FFC0&
            Caption         =   "REFERENCIA"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label Label10 
            BackColor       =   &H00C0FFC0&
            Caption         =   "NUMERO"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label15 
            BackColor       =   &H00C0FFC0&
            Caption         =   "PRECIO"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   3240
            Width           =   975
         End
         Begin VB.Label Label16 
            BackColor       =   &H00C0FFC0&
            Caption         =   "IMPORTE"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   3600
            Width           =   975
         End
      End
      Begin VB.Image Image1 
         Height          =   510
         Left            =   -390
         Picture         =   "MRECHLIN07.frx":3258
         Top             =   4755
         Width           =   2010
      End
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   225
      Left            =   9660
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   225
      Left            =   8925
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4905
      Picture         =   "MRECHLIN07.frx":6842
      Top             =   4755
      Width           =   2010
   End
End
Attribute VB_Name = "MRECHLIN07"
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
  Call LISREPRO
15 If Option2 Then
     Call FINAL("*LIREPRO")
   ElseIf Option1 Then
     Call FINAL("*RECHPRO")
'     Call MENSERR(24, "Listado no Disponible por Pantalla")
'     Option2 = 1
'     GoTo 15
  End If
99 Command10.Enabled = True
  '
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
    ACONEC$ = "ARCHIG07/AMAPRO"
    Call CONECRUN(ACONEC$, "Padron Maestro de Proveedores")

    Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    ACONEC$ = "ARCHIG07/AMASTO"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
    Command2.Enabled = True
End Sub


Private Sub Command24_Click()
   '
   Dim CIMOV%(32767) 'VECTOR QUE VERIFICA SI TUVIERON MOVIMIENTO DE RECHAZO ESE PROVEEDOR
   '
   MODOFERVI% = 0
   If InStr(1, EMPREAC$, "FERVI") > 0 Then MODOFERVI% = 1
   '
   '1º PASO : LECTURA DEPÓSITO ASOCIADO A MOTIVO DE DEVOLUCIÓN
10 For j& = 1 To ULTREG&("TADESTI")
     X$ = REGLEIDO$("TADESTI", j&)
     CODI1% = Asc(Mid$(X$, 1, 1))
     If CODI1% = 1 Then 'COMO ES TABLA CERRADA 1=DEVOLUCION
        DEPODEV% = Asc(Mid$(X$, 32, 1))
        Exit For
     End If
   Next j&
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
   For j& = 1 To ULTREG&("!RECHLIN")
     X$ = REGLEIDO$("!RECHLIN", j&)
     PRORECH% = CVI(Mid$(X$, 9, 2))
     CI1% = CVI(Mid$(X$, 5, 2))
     CIMOV%(CI1%) = 1
   Next j&
   '4º PASO LEER STOCK EN DEPOSITO DE DEVOLUCIONES PARA LOS ITEMS MARCADOS Y
   'PONERLOS EN EL ARCHIVO DEL REMITO (BOLREDET)
   Call BLANARCH("!BOLREDET")
   Call SETULTREG("!BOLREDET", 0)
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
   Call SETULTREG("!BOLREDET", I&)
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
30 If MODOFERVI% = 1 Then
        OPX% = COMALTER%("Opciones de Devolución\\Devolución Sin Regreso\Devolución En Consignación")
        If OPX% = 1 Then GoTo 40
        GoTo 70
   End If
   '
40 CODRECHA$ = "RD"
   '
   Call ABRESTOCKS
   SQLX$ = " SELECT max (CAST(SUBSTRING(DoPriCor,4,12) AS INTEGER)) as Maximo FROM MOVISTO"
   SQLX$ = SQLX$ + " WHERE CodiMovi = '" & CODRECHA$ & "'"

   UNRORECH& = 0

   Dim MIRS As ADODB.Recordset
   Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
   If Not IsNull(MIRS!maximo) Then
    UNRORECH& = XVALO(MIRS!maximo)
    GoTo 35
   End If

   UNRORECH& = 0
   '
35 REMIRECH$ = CODRECHA$ + "-" + TRIM$(Str$(UNRORECH& + 1))
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
           'Call MOVISTOK(FF%, CI%, LOTE$, CODRECHA$, REMIRECH$, REMIRECH$, DOSEC$, DOSEC$, REFE$, VUNI#, MONE$, NC1%, DEPX%, CANTI)
           Call MOVISTOK(FF%, CI%, LOTE$, CODRECHA$, REMIRECH$, REMIRECH$, DOSEC$, DOSEC$, REFE$, 0, VUNI#, 0, NC1%, DEPX%, CANTI)
           '
         End If
       End If
     End If
   Next YX&
   Call CIERRARCH("MOVISTO")
   'IMPRIME DEVOLUCION
   Call REMIRECHLIN(NC1%, REMIRECH$)
   GoTo 99
   '
70 Call GeneraRemiConsig("", NPRO%, DEPODEV%)
   
99 Command24.Enabled = True
   '
End Sub

Sub GeneraRemiConsig(REMI$, NPRO%, DEPO%, Optional NUMREMI$)
      '
      Screen.MousePointer = 11

      'GENERA MOVIMIENTO CON CODIGO 0 PARA REGISTRAR EL NUMERO DE REMITO EN CONSIGNACION
      CMO$ = TRIM$(Control$("", "CODRECON"))
      If CMO$ = "" Then CMO$ = "RC"
      '
      UNRO& = ProNroComprobante(CMO$)
      
      Call BLOQARCH("MOVISTO")
      Call BLOQARCH("TRANSFE")
      Call ABRECONEXION
      
      Screen.MousePointer = 11
      NC% = (-1) * NPRO%
      VUNI# = 0
      MONE$ = ""
      LOTE$ = ""
      FECHA% = HOY(HOS$)
      DOPRI$ = "RC-" + TRIM$(Str$(UNRO&))
      DOSEC$ = "RC-" + Mid$(HOS$, 8, 1) + Mid$(HOS$, 4, 2) + Left$(HOS$, 2)
      While Len(DOSEC$) < 10
          DOSEC$ = DOSEC$ + " "
      Wend
      NORI% = 0
      
      For YX& = 1 To ULTREG&("!BOLREDET")
          YY$ = REGLEIDO$("!BOLREDET", YX&)
          CI% = CVI(Mid$(YY$, 83, 2))
          CAN# = CVD(Mid$(YY$, 89, 8))
          CANTI = CAN# * (-1)
          REFE$ = "Remito de Traslado"
          If CI% > 0 Then
            If CI% <= NUMAS% Then
              If CAN# > 0.005 Then
                'MUEVE STOCK
                NORI% = NORI% + 1
                Call MOVISTOK(FECHA%, CI%, LOTE$, CMO$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 1, NC%, DEPO%, CANTI)
                
                    Dim MaConsig As ADODB.Recordset
                    Set MaConsig = New ADODB.Recordset
                    NUMECOMPRA$ = ""
                    SQLX$ = " select * from MACONSIG WHERE Cant_Asig < 0"
26                  On Error Resume Next
                    MaConsig.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynaset, adLockPessimistic, adCmdText
                    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                       On Error GoTo 0
                       Call CapturaErrorOpen
                       GoTo 26
                    End If
                    On Error GoTo 0
                    '
                    With MaConsig
70                         .AddNew   ' YA CONVERTIDO
                           On Error Resume Next
                           !CodiEmpr = CodiEmp%
                           On Error GoTo 0
                           !Femi_Ocom = CVDAT(FECHA%)
                           !Npro_Ocom = Abs(NC%)
                           !Raso_Prov = RASOCLI(NC%)
                            NOCOM$ = ""
                            While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
                            NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-00"
                           !Nro_Ocom = NOCOM$
                           !Cod_Inte = CI%
                           !COD_EXTE = CODEXT(CI%)
                           !NREMI_CONSIG = UNRO&
                           !Unidad = UNIMED$(CI%)
                           !cant_asig = Abs(CANTI)
                           !cant_entreg = Abs(CANTI)
                           !cant_rendid = 0
                           !cant_devu = 0
                           !FECH_ENTREG = CVDAT(FECHA%)
                           !LOTE_OCOM = ""
                           !NUORDITEM = NORI%
                           .Update
                    End With
              End If
            End If
          End If
      Next YX&
      
      'IMPRESION DE REMITO DE CONSIGNACION

      FORIPRE$ = TRIM$(Control$("", "FOREMICO"))
      If FORIPRE$ <> "" Then
            LU$ = LUDAT$
            If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
                RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
                Screen.MousePointer = 1
                Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito en Consignación:\  \'" + RMCC$ + "'.")
                Exit Sub
            End If
      End If
      '
      FEX = FECHA%
      FECHDOC$ = FECHATEX$(FEX)
      NUMEDOC$ = TRIM$(DOPRI$)
      TIPODOC$ = "Remito de Traslado"
      DTP$ = OBJPLA$("TRANSPOR", DATRANSPO$)
      DATRANSPO$ = DTP$

      CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(DTP$, 1, 32))
      CODPARAM$(2) = "DOMI-TRA": DOCPARAM$(2) = TRIM$(Mid$(DTP$, 33, 32))
      CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(DTP$, 65, 16))
      CODPARAM$(4) = "PATE-TRA": DOCPARAM$(4) = TRIM$(Mid$(DTP$, 81, 16))
      CODPARAM$(5) = "COND-TRA": DOCPARAM$(5) = TRIM$(Mid$(DTP$, 97, 32))
      CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(DTP$, 129, 24))
      CODPARAM$(7) = "LOCA-TRA": DOCPARAM$(7) = TRIM$(Mid$(DTP$, 159, 32))
      CODPARAM$(8) = "NUBULTOS": DOCPARAM$(8) = TRIM$(Mid$(DTP$, 153, 6))
      '
      CAPARDOC% = 8
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "CANTIDAD"
      CODTABU1$(4) = "CAKILOS"
      CODTABU1$(5) = "UNIMED"
      CACOLTAB1% = 5
      '
      DESTIDOC% = NPRO% * (-1)
      CAITAB1% = 0
      '
      For U& = 1 To ULTREG&("!BOLREDET")
         XX$ = REGLEIDO$("!BOLREDET", U&)
         CI% = CVI(Mid$(XX$, 83, 2))
         CAN# = CVD(Mid$(XX$, 89, 8))
         CANTI = CAN# * (-1)
         REFE$ = "Remito de Traslado"
         If CI% > 0 Then
           If CI% <= NUMAS% Then
            If CAN# > 0.005 Then
                CANS$ = FORMATNUM$(CAN#, "F12.1")
                CAITAB1% = CAITAB1% + 1
                TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
                TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
                TETABU1$(3, CAITAB1%) = CANS$
                TETABU1$(4, CAITAB1%) = CACOL1$
                TETABU1$(5, CAITAB1%) = UNIMED$(CI%)
            End If
           End If
         End If
      Next U&
      '
      NUMREMI$ = DOPRI$
      '
      Screen.MousePointer = 1
      '
      FORIPRE$ = TRIM$(Control$("", "FOREMICO"))
      If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            LU$ = LUDAT$
            If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
                RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
                Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito en Consignación:\  \'" + RMCC$ + "'.")
                Exit Sub
            End If
      End If
      Call PRINTDOC("FOREMICO") ' Remito en Consignacion
      '
End Sub

Sub REMIRECHLIN(NPROVE%, REMIRECH$) 'IMPRIME DEVOLUCION
   '
   FORIPRE$ = TRIM$(Control$("", "REMIRECH"))
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

Private Sub Command3_Click()
    '
    ' ///OT-8-950-21/10/08
    ' PROCESO DE RECUPERACION DE RECHAZO MAL VALORIZADOS
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    '
    DESDE$ = FETEXCOR1$(DES%)
    HASTA$ = FETEXCOR1$(HAS%)
    '
    ' BUSCA LOS RECHAZOS EN EL RANGO DE FECHAS ELEGIDO
10  SQLX$ = "SELECT * "
    SQLX$ = SQLX$ + "FROM MOVISTO WHERE FechMov >= '" & DESDE$ & "' AND FechMov <= '" & HASTA$ & "' AND "
    SQLX$ = SQLX$ + "DoSecLar LIKE '%RECH%' AND CantSalid > 0 ORDER BY FechMov ASC"
    '
    Call APERBASDAT("STOCKS")
    '
    Screen.MousePointer = 11
    Dim MIRS As ADODB.Recordset
    Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
    If Not (MIRS.BOF And MIRS.EOF) Then
      MIRS.MoveFirst
      ' RECORRE LOS RECHAZOS Y VERIFICA SI ESTAN BIEN VALORIZADOS
      j& = 0:  Call SETULTREG("RECURECH", 0)
      Do While Not MIRS.EOF
         CI% = MIRS!Cod_Inte
         CODIGO$ = CODEXT$(CI%)
         NPRO% = MIRS!Nume_Prov
         If NPRO% < 1 Then GoTo 25
         FECHA% = CVINT(MIRS!FECHMOV)
         Cantidad# = XVALO(MIRS!CANTSALID)
         PRECIOREGISTRADO# = XVALO(MIRS!VALOUNIT)
         SQLY$ = "COD_INTE=" & CI% & " AND NPRO_OCOM=" & NPRO% & " AND FEMI_OCOM <= '" & FETEXCOR1$(FECHA%) & "'"
         SQLY$ = SQLY$ + " ORDER BY FEMI_OCOM DESC"
         ULTIMOPRECIO# = XVALO(VALOBADA("OCOMDETA", "Pre_Unit", SQLY$)) ' OBTIENE ULTIMO PRECIO DE COMPRA
         MONEDA% = XVALO(VALOBADA("OCOMDETA", "Mone_Emis", SQLY$)) ' OBTIENE MONEDA
         PRECIOPESIFICADO# = ULTIMOPRECIO#
         If MONEDA% > 1 Then
           PRECIOPESIFICADO# = ULTIMOPRECIO# * TICAMONE(MONEDA%)
         End If
         '
         DIF1# = PRECIOPESIFICADO# - PRECIOREGISTRADO
         If DIF1# >= 0.005 Then
             RBLAN$ = REGBLAN$("RECURECH")
             Call REPLA(RBLAN$, MKI$(NPRO%), 1, 2)
             Call REPLA(RBLAN$, MKI$(CI%), 3, 2)
             Call REPLA(RBLAN$, MKS$(Cantidad#), 5, 4)
             Call REPLA(RBLAN$, MKD$(PRECIOREGISTRADO#), 9, 8)
             Call REPLA(RBLAN$, MKD$(PRECIOPESIFICADO#), 17, 8)
               DIF1# = PRECIOPESIFICADO# - PRECIOREGISTRADO
             Call REPLA(RBLAN$, MKD$(DIF1#), 25, 8)
               DIF2# = Cantidad# * DIF1#
             Call REPLA(RBLAN$, MKD$(DIF2#), 33, 8)
             Call REPLA(RBLAN$, MKI$(FECHA%), 41, 2)
             '
             j& = j& + 1
             Call GRAREG("RECURECH", RBLAN$, j&)
         End If
         '
25    MIRS.MoveNext
      Loop
      Call SETULTREG("RECURECH", j&)
      Call FILESORT("RECURECH", "", 1, 1, "ASC")
      '
      OPX% = COMALTER%("Opciones de Trabajo:\\Listado Impreso\Ajuste de Valores")
      If OPX% = 1 Then
           ' EMITE LISTADO
           Call FINAL("*RECURECH")
        ElseIf OPX% = 2 Then
           Call DIFRECHA
           GoTo 10
      End If
    End If
    Screen.MousePointer = 1
    '
    End Sub

Sub DIFRECHA()
    '
5   Screen.MousePointer = 11
    UMAX& = 0
    For U& = 1 To ULTREG&("RECURECH")
       X$ = REGLEIDO$("RECURECH", U&)
       FF% = CVI(Mid$(X$, 41, 2))
       CIXI% = CVI(Mid$(X$, 3, 2))
       CAN = CVS(Mid$(X$, 5, 4))
       NPRX = CVI(Mid$(X$, 1, 2))
       TTXX$ = FECHATEX$(FEX) + "  " + CODEXT$(CIXI%) + " (" + TRIM$(FORMATNUM$(CAN, "F12.1")) + ") - " + RASOCLI$((-1) * NPRX)
       Call GRAREG("!INDIRECH", MKS$(U&) + TTXX$, U&)
       UMAX& = U&
    Next U&
    Call SETULTREG("!INDIRECH", UMAX&)
    Call FRESHECHO("!INDIRECH")
    Screen.MousePointer = 1
    '
10  Call SELECHO("!INDIRECH")
    U& = XVALO(VALACT1$("!INDIRECH"))
    If U& > 0 Then
       X$ = REGLEIDO$("RECURECH", U&)
       FF% = CVI(Mid$(X$, 41, 2))
       NPRO% = CVI(Mid$(X$, 1, 2))
       CIXI% = CVI(Mid$(X$, 3, 2))
       CAN = CVS(Mid$(X$, 5, 4))
       PRECIORECHA# = CVD(Mid$(X$, 9, 8))
       PRECIOACTUAL# = CVD(Mid$(X$, 17, 8))
       DIFEUNIT# = CVD(Mid$(X$, 25, 8))
       DIFETOTA# = CVD(Mid$(X$, 33, 8))
       '
       If DIFETOTA# >= 0.005 Then
            ' VERIFICAR SI LO ENCUENTRA
            FECHAMOVI$ = FETEXCOR1$(FF%)
            SQLX$ = "FechMov = '" & FECHAMOVI$ & "' "
            SQLX$ = SQLX$ + " AND Cod_Inte = " & CIXI%
            SQLX$ = SQLX$ + " AND Nume_Prov = " & NPRO%
            SQLX$ = SQLX$ + " AND DoSecLar LIKE '%RECH%' "
            SQLX$ = SQLX$ + " AND CantSalid > 0 "
            CANRECHA = XVALO(VALOBADA("MOVISTO", "CANTSALID", SQLX$))
            If Abs(CANRECHA - CAN) > 0.05 Then
               Call MENSERR(24, "Imposible Procesar Ajuste Valor de este Rechazo")
               GoTo 10
            End If
            '
            SQLZ$ = "UPDATE MOVISTO SET VALOUNIT = " & TRIM$(FORMATNUM$(PRECIOACTUAL#, "F16.5")) & "WHERE " + SQLX$
            '
            FIN1& = ULTREG&("RECHLIN")
            ULTRECH& = FIN1& + 1
            If FIN1& > 0 Then
               XX1$ = REGLEIDO$("RECHLIN", FIN1&)
               ULTRECH& = 1 + CVS(Mid$(XX1$, 1, 4))
            End If
            Y$ = Space(128)
            Call REPLA(Y$, MKS$(ULTRECH&), 1, 4)
            Call REPLA(Y$, MKI$(CIXI%), 5, 2)
            Call REPLA(Y$, MKI$(FF%), 7, 2)
            Call REPLA(Y$, MKI$(NPRO%), 9, 2)
            Call REPLA(Y$, MKS$(0), 11, 4)
            Call REPLA(Y$, USERNAME$, 17, 32)
            Call REPLA(Y$, "AJUSTE DE PRECIO", 49, 48) 'FALTA REFERENCIA
            Call REPLA(Y$, MKD$(DIFEUNIT#), 97, 8) 'FALTA VALOR UNITARIO
            Call REPLA(Y$, MKD$(DIFETOTA#), 105, 8) 'FALTA IMPORTE
            Call REPLA(Y$, MKI$(0), 113, 2) 'FALTA PROCESADO FECHA
            Call REPLA(Y$, MKI$(0), 115, 2) 'FALTA NUMERO USUARIO PROCESO
            Call REGAPP("RECHLIN", Y$)
            Call CIERRARCH("RECHLIN")
            'AGREGO A ARCHIVO INDICE DE RECHAZO
            REFERE2$ = FECHATEX$(FF%)
            REFERE2$ = REFERE2$ + " - " + RASOCLI$(NPRO% * (-1))
            ZZ$ = Space(64)
            Call REPLA(ZZ$, MKS$(ULTRECH&), 1, 4)
            Call REPLA(ZZ$, REFERE2$, 5, 60)
            Call REGAPP("INDIRECH", ZZ$)
            Call CIERRARCH("INDIRECH")
            '
             FLEXCONN.Execute (FILTSQL$(SQLZ$))     ' ESTO AJUSTA EL VALOR
            '
            FEMI1$ = FECHATEX$(FF%)
            NUDO$ = ""
            REFMO$ = "AJUSTE VALOR RECHAZO"
            PREUNI1# = DIFEUNIT#
            IMPORT1# = DIFETOTA#
            NUDO$ = FORMATNUM$(ULTRECH&, "F11")
            Call PRIRECHALI(1, FEMI1$, NUDO$, CIXI%, 0, 1, NPRO%, 0, REFMO$, PREUNI1#, IMPORT1#) 'IMPRESION LOCAL
            Call PRIRECHALI(2, FEMI1$, NUDO$, CIXI%, 0, 1, NPRO%, 0, REFMO$, PREUNI1#, IMPORT1#) 'IMPRESION REMOTA
            '
            EMAI$ = TRIM$(EMAILCLI$((-1) * NPRO%))
            If EMAI$ = "" Or InStr(EMAI$, "@") < 1 Then
                Call MENSERR(24, "Imposible Enviar por Correo Electrónico.\  \Dirección de Correo del Proveedor es\Inexistente o no Válida.")
              Else
                '\\\OT-08-0802-OD-06/09/08///
                'OPCION DE CONSULTAR PRIMERO SI DESEA ENVIAR POR MAIL EL RECHAZO DE LINEA
                OPX% = COMALTER%("¿Desea Enviar Rechazo de Linea Por Email?\\No\Si")
                If OPX% = 2 Then
                     SENDBYMAIL$ = EMAI$
                     Call PRIRECHALI(1, FEMI1$, NUDO$, CIXI%, 0, 1, NPRO%, 0, REFMO$, PREUNI1#, IMPORT1#) 'IMPRESION LOCAL
                End If
                '\\\OT-08-0802-OD-FIN///
            End If
            '
            Call COMUNI("Ajuste de Rechazo Procesado")
            '
       End If
    
    End If
    '
End Sub

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
      Call PROCERECHA(CODI1%, FERECH%, CARECHA, DOCORX$)
      '\\\************************************///
      Call BAJALDISCO
      'Call RECONECTA 'para setear nuevamente la impresora predeterminada de Windows
      If InStr(1, EMPREAC$, "FERVI") < 1 Then
        Call FINAL("")
        Exit Sub
      End If
      Call COMUNI("Rechazo Procesado")
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
    NUDOC& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
    If NUDOC& > 0 Then
        Call MUESTRADOC(NUDOC&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
        GoTo 3
    Else
        ' SI NO LO ENCONTRO, VUELVE A GRABAR LA IMPRESION. ESTO POR SI POR ALGUN MOTIVO SE CORTO JUSTO
        ' AL MOMENTO DE IMPRIMIR EL RECHAZO
        For I& = 1 To ULTREG("RECHLIN")
            XX$ = REGLEIDO$("RECHLIN", I&)
            If CVS(Mid$(XX$, 1, 4)) = NPX& Then
                CI% = CVI(Mid$(XX$, 5, 2))
                FEMI1$ = FECHATEX$(CVI(Mid$(XX$, 7, 2)))
                CAN3 = CVS(Mid$(XX$, 11, 4))
                DESTI% = Asc(Mid$(15, 1))
                NPRO% = CVI(Mid$(XX$, 9, 2))
                MOTI% = Asc(Mid$(XX$, 16, 1))
                REFMO$ = Mid$(XX$, 49, 48)
                PREUNI1# = CVD(Mid$(XX$, 97, 8))
                IMPORT1# = CVD(Mid$(XX$, 105, 8))
                Exit For
            End If
        Next I&
    End If
    
    CANCELPRINT% = 1
    NUDO$ = FORMATNUM$(NPX&, "F11")
    Call PRIRECHALI(1, FEMI1$, NUDO$, CI%, CAN3, DESTI%, NPRO%, MOTI%, REFMO$, PREUNI1#, IMPORT1#) 'IMPRESION LOCAL
    CANCELPRINT% = 0
    
    TIDOCUM$ = "Rechazo en Línea"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU$ = " " + NUDOCU$ + " "
    Screen.MousePointer = 11
    NUDOC& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
    If NUDOC& > 0 Then
        Call MUESTRADOC(NUDOC&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
    End If
    GoTo 3
    
'    For UI& = ULTREG&("PDOCLST") To 1 Step -1
'      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'      If InStr(XX$, TIDOCUM$) > 4 Then
'        PPXX% = InStr(XX$, "Nro.")
'        If PPXX% > 0 Then
'          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
'        End If
'        If InStr(XX$, NUDOCU$) > 4 Then
'          DOCUNU& = CVS(Left$(XX$, 4))
'          GoTo 5
'        End If
'      End If
'    Next UI&
'    Screen.MousePointer = 1
'    Call COMUNI("Documento no Encontrado")
'    GoTo 99
'    '
'5   For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    '
'    REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'    GoTo 8
'    '
'7   Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    GoTo 99
'    '
'8   TPSP$ = ""
'    Screen.MousePointer = 11
'    While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 9
'      L& = L& - 1
'    Wend
'    '
'9   CAPAGINAS% = 0
'    For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'95  SPOOLSTRING$ = TPSP$
'    FIN& = Len(TPSP$)
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).Top = 0
'      EDITFORM.Picture1(KKI%).Refresh
'      EDITFORM.Picture1(KKI%).Height = 7070
'      EDITFORM.Picture1(KKI%).Width = 11700
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    ALTUPAGINA = 7070
'    TOPEPAGINA = 0
'    PAGINACTIVA% = 0
'    CAPAGINAS% = 0
'    '
'10  PPXX% = InStr(TPSP$, Chr$(255))
'    If PPXX% > 0 Then
'      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'      TPSP$ = Mid$(TPSP$, PPXX% + 1)
'      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
'      '
'      If Len(COPRIAT$) > 0 Then
'        COPRI% = Asc(Left$(COPRIAT$, 1))
'        ATRIP$ = Mid$(COPRIAT$, 2)
'        Call SHOWCOMMAND(COPRI%, ATRIP$)
'      End If
'      GoTo 10
'      '
'    End If
'    '
'    UI& = UI& + 1
'    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'    If InStr(XX$, TIDOCUM$) > 4 Then
'      If InStr(XX$, NUDOCU$) > 4 Then
'        DOCUNU& = CVS(Left$(XX$, 4))
'        GoTo 7
'      End If
'    End If
'    '
'    Screen.MousePointer = 1
'    PAGINACTIVA% = 0
'    For KKI% = 0 To 5
'       EDITFORM.Picture1(KKI%).Visible = False
'    Next KKI%
'    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
'    EDITFORM.Picture1(PAGINACTIVA%).Refresh
'    EDITFORM.VScroll1.Min = 0
'    EDITFORM.VScroll1.Value = 0
'    EDITFORM.VScroll1.Max = 1
'    If ALTUPAGINA > EDITFORM.Frame1.Height Then
'       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'    End If
'    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
'    EDITFORM.HScroll1.Min = 0
'    EDITFORM.HScroll1.Value = 0
'    EDITFORM.HScroll1.Max = 1
'    If ANCHPAGINA > EDITFORM.Frame1.Width Then
'      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
'    End If
'    EDITFORM.Command3.Enabled = False
'    EDITFORM.Command1.Enabled = False
'    EDITFORM.Show 1
'    '
'    GoTo 3
    '
99  Command5.Enabled = True
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
   Openforms = DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Openforms = DoEvents
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
  OPX% = COMALTER%("Opciones de Listado\\Totales\Pendientes")
  '
  JJJ& = 0
  Call BLANARCH("LIREPRO")
  FIN& = ULTREG&("RECHLIN")
  For j& = 1 To ULTREG&("RECHLIN")
    '
    Call TRACE(20, j&, FIN&)
    '
    X$ = REGLEIDO$("RECHLIN", j&)
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
    If OPX% = 2 Then
        FE% = CVI(Mid$(X$, 113, 2))
        US% = CVI(Mid$(X$, 115, 2))
        If FE% > 0 And US% > 0 Then GoTo 75
    End If
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
75 Next j&
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
  TIPOFILTRO$ = "Filtro: TOTALES"
  If OPX% = 2 Then TIPOFILTRO$ = "Filtro: PENDIENTES"
  Call HEADERS("LIREPRO", TIPOFILTRO$)
  'Call FINAL("*LIREPRO")
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
  For j& = 1 To ULTREG&("RECHLIN")
    '
    Call TRACE(20, j&, FIN&)
    '
    X$ = REGLEIDO$("RECHLIN", j&)
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
75 Next j&
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
  For j& = 1 To ULTREG&("RECHLIN")
    '
    Call TRACE(20, j&, FIN&)
    '
    X$ = REGLEIDO$("RECHLIN", j&)
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
75 Next j&
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
  For j& = 1 To ULTREG&("RECHLIN")
    X$ = REGLEIDO$("RECHLIN", j&)
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
  Next j&
  Call CIERRARCH("!INDIRECH")
  '
End Sub

'Sub REMIDEV(NC1%, DEPX1%, CI1%, RECHA1#, NURECH$, FECHRECH%, VUNI1#)
'    '
'    If TRIM$(Control$("", "REMIRECH")) <> "" Then
'        CODRECHA$ = TRIM$(Control$("", "CODRRECH"))
'        If CODRECHA$ = "" Then CODRECHA$ = "RD"
'        FIN& = ULTREG&("MOVISTO")
'        For U& = FIN& To 1 Step -1
'           Call TRACE(20, FIN& - U& + 1, FIN&)
'           XX$ = REGLEIDO$("MOVISTO", U&)
'           If UCase$(Mid$(XX$, 21, 2)) = CODRECHA$ Then
'             UNRORECH& = Val(Mid$(XX$, 26, 7))
'             GoTo 11241
'           End If
'        Next U&
'        UNRORECH& = 0
'        '
'11241   REMIRECH$ = CODRECHA$ + "-" + TRIM$(Str$(UNRORECH& + 1))
'        '
'        DOSEC$ = NURECH$
'        DEPX% = DEPX1%
'        RECHA# = RECHA1#
'        CANTI = RECHA# * (-1)
'        REFE$ = "RECHAZO - Rto." + TRIM$(REMI$)
'        FF% = FECHRECH%
'        VUNI# = VUNI1#
'        CI% = CI1%
'        Call Movisto(FF%, CI%, LOTE$, CODRECHA$, REMIRECH$, DOSEC$, REFE$, VUNI#, MONE$, NC1%, DEPX%, CANTI)
'        '
'        Call BLANARCH("!BOLREDET")
'        Z$ = REGBLAN$("BOLREDET")
'        Call REPLA(Z$, MKI$(CI%), 83, 2)
'        Call REPLA(Z$, MKD$(RECHA1#), 171, 8)
'        Call REPLA(Z$, LOTE$, 69, 12)
'        Call REGAPP("!BOLREDET", Z$)
'        Call CIERRARCH("!BOLREDET")
'        '
'        Call REMIRECHA(NC1%, DEPX%, REMIRECH$)
'        '
'      End If
'    '
'End Sub
'
Private Sub Picture1_Click()

End Sub


