VERSION 5.00
Begin VB.Form COMANO 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Mano de Obra"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   7860
   ClipControls    =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   7860
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command19 
      Caption         =   "Repara"
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
      Left            =   4200
      Picture         =   "COMANO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Gestión de Reparaciones"
      Top             =   360
      Width           =   740
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Tabla de Operarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   4200
      TabIndex        =   40
      Top             =   1160
      Width           =   3495
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   480
         ScaleHeight     =   585
         ScaleWidth      =   2340
         TabIndex        =   48
         Top             =   480
         Width           =   2400
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
            Picture         =   "COMANO.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Padron de Operarios"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "A-B-M de Opera- rios Productivos"
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
            Left            =   750
            TabIndex        =   49
            Top             =   120
            Width           =   1600
         End
      End
      Begin VB.Frame Frame2 
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
         TabIndex        =   41
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command25 
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
            TabIndex        =   47
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command24 
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
            TabIndex        =   46
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command9 
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
            TabIndex        =   45
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command5 
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
            TabIndex        =   44
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command4 
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
            TabIndex        =   43
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command3 
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
            TabIndex        =   42
            Top             =   1890
            Width           =   1590
         End
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   5160
      ScaleHeight     =   75
      ScaleWidth      =   675
      TabIndex        =   31
      Top             =   360
      Width           =   740
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   -225
         TabIndex        =   32
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Informar Operaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   19
      Top             =   2520
      Width           =   3855
      Begin VB.PictureBox Picture9 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   960
         ScaleHeight     =   585
         ScaleWidth      =   2340
         TabIndex        =   20
         Top             =   960
         Width           =   2400
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
            Left            =   0
            Picture         =   "COMANO.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Informe de Operaciones por Orden de Trabajo"
            Top             =   0
            Width           =   590
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Informes por Orden de Trabajo"
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
            TabIndex        =   21
            Top             =   120
            Width           =   1600
         End
      End
      Begin VB.Frame Frame5 
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
         TabIndex        =   24
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command18 
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
            TabIndex        =   30
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command14 
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
            TabIndex        =   29
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command13 
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
            TabIndex        =   28
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command12 
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
            TabIndex        =   27
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command11 
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
            TabIndex        =   26
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command10 
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
            TabIndex        =   25
            Top             =   420
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture10 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   600
         ScaleHeight     =   585
         ScaleWidth      =   2340
         TabIndex        =   22
         Top             =   525
         Width           =   2400
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
            Picture         =   "COMANO.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Informe de Operaciones por Fecha"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Informe por Fecha"
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
            TabIndex        =   23
            Top             =   150
            Width           =   1800
         End
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
      Height          =   720
      Left            =   6180
      Picture         =   "COMANO.frx":1388
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   360
      Width           =   740
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
      Height          =   720
      Left            =   6960
      Picture         =   "COMANO.frx":1692
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cerrar - Salir de la Aplicación"
      Top             =   360
      Width           =   740
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ordenes de Trabajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   225
      TabIndex        =   1
      Top             =   225
      Width           =   3855
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1080
         ScaleHeight     =   585
         ScaleWidth      =   2340
         TabIndex        =   35
         Top             =   1320
         Width           =   2400
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
            Picture         =   "COMANO.frx":17DC
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Cierre de Orden de Trabajo"
            Top             =   0
            Width           =   530
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Cierre de Orden de Trabajo"
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
            Top             =   110
            Width           =   1515
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   720
         ScaleHeight     =   585
         ScaleWidth      =   2340
         TabIndex        =   17
         Top             =   840
         Width           =   2400
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
            Picture         =   "COMANO.frx":1AE6
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Modificar Descripción de Orden de Trabajo Abierta "
            Top             =   0
            Width           =   530
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Modificar O/Trabajo"
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
            TabIndex        =   18
            Top             =   150
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   360
         ScaleHeight     =   585
         ScaleWidth      =   2340
         TabIndex        =   15
         Top             =   420
         Width           =   2400
         Begin VB.CommandButton Command41 
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
            Picture         =   "COMANO.frx":1DF0
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Generar Nueva Orden de Trabajo"
            Top             =   0
            Width           =   530
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar O/Trabajo"
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
            TabIndex        =   16
            Top             =   150
            Width           =   1905
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
      BackColor       =   &H00C0E0FF&
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
      Height          =   1700
      Left            =   4200
      TabIndex        =   0
      Top             =   2640
      Width           =   3525
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   720
         ScaleHeight     =   585
         ScaleWidth      =   2340
         TabIndex        =   11
         Top             =   840
         Width           =   2400
         Begin VB.CommandButton Command26 
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
            Picture         =   "COMANO.frx":20FA
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Resumen de Operaciones por Orden de Trabajo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado por Orden de Trabajo"
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
            TabIndex        =   12
            Top             =   120
            Width           =   1710
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   360
         ScaleHeight     =   585
         ScaleWidth      =   2340
         TabIndex        =   13
         Top             =   360
         Width           =   2400
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
            Left            =   0
            Picture         =   "COMANO.frx":2404
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Resumen de Operaciones por Operario"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Detalle por Operario"
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
            Top             =   50
            Width           =   1425
         End
      End
   End
   Begin VB.Label Label8 
      Height          =   225
      Left            =   5985
      TabIndex        =   55
      Top             =   105
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      Height          =   330
      Left            =   4830
      TabIndex        =   54
      Top             =   105
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10500
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4560
      Picture         =   "COMANO.frx":2D8E
      Top             =   570
      Width           =   2010
   End
   Begin VB.Menu archivo 
      Caption         =   "Informe Operaciones "
      Begin VB.Menu infop_fecha 
         Caption         =   "Informe por Fecha"
      End
      Begin VB.Menu infop_ot 
         Caption         =   "Informe por Orden de Trabajo"
      End
      Begin VB.Menu phl1 
         Caption         =   "-"
      End
      Begin VB.Menu print 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu phl2 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu ordtra 
      Caption         =   "Ordenes de Trabajo"
      Begin VB.Menu ot_gen 
         Caption         =   "Nueva Orden deTrabajo"
      End
      Begin VB.Menu ot_modif 
         Caption         =   "Modificar Existente"
      End
      Begin VB.Menu ot_cerrar 
         Caption         =   "Cierre de Orden de Trabajo"
      End
   End
   Begin VB.Menu conlis 
      Caption         =   "Listados "
      Begin VB.Menu lis_oper 
         Caption         =   "Por Operario"
      End
      Begin VB.Menu lis_ot 
         Caption         =   "Por Orden de Trabajo"
      End
   End
   Begin VB.Menu tablas 
      Caption         =   "Tablas Auxiliares"
      Begin VB.Menu tab_oper 
         Caption         =   "Padrón de Operarios"
      End
   End
   Begin VB.Menu complem 
      Caption         =   "Complementos"
      Begin VB.Menu repara 
         Caption         =   "Sub-Sistema de Reparaciones"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "COMANO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
10 FEX = HOY(HO$)
   HAS% = FEX
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   DES% = FECHANUM(DESS$)
   VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
20 OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
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
   VDEF$ = OBX$
   '
30 OPSEL% = 0
   RSX$ = REGSEL$("PADRON")
   If Len(RSX$) <= 4 Then GoTo 99
   '
   OPSEL% = CVI(Left$(RSX$, 2))
   PERIO1$ = PERIO$ + " - Operario: " + TRIM$(ECOARCH$("PADRON", MKI$(OPSEL%)))
   '
   Screen.MousePointer = 11
   JJ& = 0
   Call COPYSTRU("PARTEMOD", "PARTESEL")
   FIN& = ULTREG&("PARTEMOD")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("PARTEMOD", U&)
     If CVI(Left$(XX$, 2)) = OPSEL% Then
       FEMII% = CVI(Mid$(XX$, 3, 2))
       If FEMII% >= DES% Then
         If FEMII% <= HAS% Then
           TIEDEC = DIFHORA(CVI(Mid$(XX$, 53, 2)), CVI(Mid$(XX$, 55, 2)))
           Call REPLA(XX$, MKS$(TIEDEC), 65, 4)
           JJ& = JJ& + 1
           Call GRAREG("PARTESEL", XX$, JJ&)
         End If
       End If
     End If
   Next U&
   '
   Call SETULTREG("PARTESEL", JJ&)
   Call CIERRARCH("PARTESEL")
   '
   Call FILESORT("PARTESEL", "", 3, 3, "ASC")
   Call HEADERS("PARTESEL", "")
   Call HEADERS("PARTESEL", "Corresponde a: " + PERIO1$)
   Call FINAL("*LIPAROP")
   Screen.MousePointer = 1
   GoTo 30
   '
99 Command15.Enabled = True
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   Call CONECRUN("REPARA02", "Sub-Sistema de Reparaciones")
   Command19.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call FINAL("*PADRON")
   Command2.Enabled = True
End Sub

Private Sub Command26_Click()
    Command26.Enabled = False
    '
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
    VDEF$ = OBX$
    Label5 = CStr(DES%)
    Label8 = CStr(HAS%)
    '
    COMANOCO.Show 1
    '
99  Command26.Enabled = True
End Sub

Private Sub Command41_Click()
   Command41.Enabled = False
   DEORTRA$ = InputBox$("Descripción de la O.T.", "Generar Nueva Orden de Trabajo")
   If DEORTRA$ <> "" Then
     HOII% = HOY(HOS$)
     Call COPYSTRU("INDIOTRA", "INDIOTRB")
     NUOT& = 1 + ULTREG&("INDIOTRA")
     NOTX$ = "OT-" + TRIM$(Str$(NUOT&))
     ZZ$ = REGBLAN$("INDIOTRA")
     Call REPLA(ZZ$, NOTX$, 1, 12)
     Call REPLA(ZZ$, DEORTRA$, 13, 50)
     Call REPLA(ZZ$, MKI$(HOII%), 61, 2) ' FECHA DE APERTURA
     Call REPLA(ZZ$, MKI$(0), 63, 2) ' FECHA DE CIERRE
     Call REGAPP("INDIOTRA", ZZ$)
     Call REGAPP("INDIOTRB", ZZ$)
     Call CIERRARCH("INDIOTRA")
     Call CIERRARCH("INDIOTRB")
   End If
   Command41.Enabled = True
End Sub

Private Sub Command46_Click()
   Command46.Enabled = False
   Call COPYSTRU("INDIOTRA", "INDIOTRB")
   Call FILTERFILE("INDIOTRA", "INDIOTRB", 4, MKI$(0))
   Call SELECHO("INDIOTRB/Ordenes de Trabajo para CIERRE")
   OTSEL$ = VALACT1$("INDIOTRB")
   If OTSEL$ <> "" Then
     ZZ$ = FILTERGETRECORD("INDIOTRA", 1, OTSEL$)
     If MsgBox("Confirma Cierre de " + TRIM$(OTSEL$) + " - " + TRIM$(Mid$(ZZ$, 13, 50)) + " ? ", vbYesNoCancel, "Cierre de Orden de Trabajo") = 6 Then
       HOII% = HOY(HOS$)
       Call REPLA(ZZ$, MKI$(HOII%), 63, 2)
       Call FILTERPUTRECORD("INDIOTRA", 1, OTSEL$, ZZ$)
       Call CIERRARCH("INDIOTRA")
       Call CIERRARCH("INDIOTRB")
     End If
   End If
   Command46.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call COPYSTRU("INDIOTRA", "INDIOTRB")
   Call FILTERFILE("INDIOTRA", "INDIOTRB", 4, MKI$(0))
   Call SELECHO("INDIOTRB/Ordenes de Trabajo para MODIFICACION")
   OTSEL$ = VALACT1$("INDIOTRB")
   If OTSEL$ <> "" Then
     ZZ$ = FILTERGETRECORD("INDIOTRA", 1, OTSEL$)
     DEORTRA$ = InputBox$("Nueva Descripción de la " + TRIM$(OTSEL$) + Chr$(13) + Chr$(10) + TRIM$(Mid$(ZZ$, 13, 50)), "Modificar Orden de Trabajo", TRIM$(Mid$(ZZ$, 13, 50)))
     If DEORTRA$ <> "" Then
       Call REPLA(ZZ$, DEORTRA$, 13, 48)
       Call FILTERPUTRECORD("INDIOTRA", 1, OTSEL$, ZZ$)
       Call CIERRARCH("INDIOTRA")
     End If
   End If
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   Call GENERINDOT
5  Call SELECHO("SELFECHA")
   FESEL$ = VALACT1$("SELFECHA")
   If TRIM$(FESEL$) <> "" Then
     FESELI% = FECHANUM(FESEL$)
     Call FILTERFILE("PARTEMOD", "!PARTEMOD", 3, MKI$(FESELI%))
     ATRI$ = "TITUPAN=Parte de Operacion por Fecha: " + TRIM$(FESEL$)
10   VTB% = VENTABU%("PAROTFE/" + ATRI$)
     If VTB% >= 0 Then
       URE& = ULTREG&("!PARTEMOD")
       ATEX% = 0
       For U& = 1 To URE&
         ZZ$ = REGLEIDO$("!PARTEMOD", U&)
         NOPI% = CVI(Mid$(ZZ$, 1, 2))
         Call REPLA(ZZ$, MKI$(NOPI%), 51, 2) ' APAREA OPERARIO
         Call REPLA(ZZ$, MKI$(FESELI%), 3, 2)
         Call REPLA(ZZ$, MKI$(FESELI%), 17, 2) ' APAREA FECHA
         NORFA$ = TRIM$(Mid$(ZZ$, 5, 12))
         If NORFA$ = "" Or ECOARCH$("INDITOTT", NORFA$) = "" Then
           Call MENSERR(24, "Falta Orden de Trabajo")
           ATEX% = 1
         End If
         If CVI(Mid$(ZZ$, 55, 2)) < CVI(Mid$(ZZ$, 53, 2)) Then
           Call MENSERR(24, "Hora 'Hasta' Anterior a 'Desde'")
           ATEX% = 1
         End If
         TIEMIN% = 60 * DIFHORA(CVI(Mid$(ZZ$, 53, 2)), CVI(Mid$(ZZ$, 55, 2)))
         TMX$ = AJUSTD$(HORAMINU$(TIEMIN%), 8)
         If TMX$ <> Mid$(ZZ$, 57, 8) Then ATEX% = 1
         Call REPLA(ZZ$, TMX$, 57, 8)
         TIEDEC = DIFHORA(CVI(Mid$(ZZ$, 53, 2)), CVI(Mid$(ZZ$, 55, 2)))
         Call REPLA(ZZ$, MKS$(TIEDEC), 65, 4)
         Call GRAREG("!PARTEMOD", ZZ$, U&)
       Next U&
       Call CIERRARCH("!PARTEMOD")
       '
       If ATEX% > 0 Then GoTo 10
       '
       Call FILTERUPDATE("PARTEMOD", "", 3, MKI$(FESELI%))
       Call CIERRARCH("PARTEMOD")
       GoTo 5
     End If
   End If
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call GENERINDOT
5  Call SELECHO("INDITOTP")
   OTSEL$ = VALACT1$("INDITOTP")
   If TRIM$(OTSEL$) <> "" Then
     Call FILTERFILE("PARTEMOD", "!PARTEMOD", 4, OTSEL$)
     ATRI$ = "TITUPAN=Parte por O.T: " + TRIM$(OTSEL$) + " - " + TRIM$(ECOARCH("INDIOTRA", OTSEL$))
10   VTB% = VENTABU%("PAROTRA/" + ATRI$)
     If VTB% >= 0 Then
       URE& = ULTREG&("!PARTEMOD")
       ATEX% = 0
       For U& = 1 To URE&
         ZZ$ = REGLEIDO$("!PARTEMOD", U&)
         NOPI% = CVI(Mid$(ZZ$, 1, 2))
         Call REPLA(ZZ$, MKI$(NOPI%), 51, 2) ' APAREA OPERARIO
         FEXI% = CVI(Mid$(ZZ$, 3, 2))
         Call REPLA(ZZ$, MKI$(FEXI%), 17, 2) ' APAREA FECHA
         If FEXI < 1 Then
           Call MENSERR(24, "Falta Fecha")
           ATEX% = 1
         End If
         If CVI(Mid$(ZZ$, 55, 2)) < CVI(Mid$(ZZ$, 53, 2)) Then
           Call MENSERR(24, "Hora 'Hasta' Anterior a 'Desde'")
           ATEX% = 1
         End If
         TIEMIN% = 60 * DIFHORA(CVI(Mid$(ZZ$, 53, 2)), CVI(Mid$(ZZ$, 55, 2)))
         TMX$ = AJUSTD$(HORAMINU$(TIEMIN%), 8)
         If TMX$ <> Mid$(ZZ$, 57, 8) Then ATEX% = 1
         Call REPLA(ZZ$, TMX$, 57, 8)
         TIEDEC = DIFHORA(CVI(Mid$(ZZ$, 53, 2)), CVI(Mid$(ZZ$, 55, 2)))
         Call REPLA(ZZ$, MKS$(TIEDEC), 65, 4)
         Call GRAREG("!PARTEMOD", ZZ$, U&)
       Next U&
       Call CIERRARCH("!PARTEMOD")
       '
       If ATEX% > 0 Then GoTo 10
       '
       Call FILTERUPDATE("PARTEMOD", "", 4, OTSEL$)
       Call CIERRARCH("PARTEMOD")
       GoTo 5
     End If
   End If
   Command8.Enabled = True
End Sub

Private Sub Exit_Click()
   Call Command1_Click
End Sub

Private Sub Form_Load()
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Close: End
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   DoEvents
   '
End Sub

Private Sub ot_cerrar_Click()
   Call Command46_Click
End Sub

Private Sub ot_gen_Click()
   Call Command41_Click
End Sub

Private Sub ot_modif_Click()
   Call Command6_Click
End Sub

Private Sub repara_Click()
   Call Command19_Click
End Sub

Private Sub tab_oper_Click()
   Call Command2_Click
End Sub

Sub GENERINDOT()
   '
   Call COPYSTRU("INDIOTRA", "INDITOTT") ' INDICE DE TODAS LAS o/t'S
   Call COPYSTRU("INDIOTRA", "INDITOTP") ' INDICE DE o/t'S ABIERTAS
   '
   Screen.MousePointer = 11
   JJ1& = 0: JJ2& = 0
   REBLA$ = REGBLAN$("INDITOTT")
   FIN& = ULTREG&("PRESREPA")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("PRESREPA", U&)
     If CVS(Mid$(XX$, 125, 4)) >= 0.5 Then
       NOREPARA& = CVS(Left$(XX$, 4))
       NORFA$ = "OR-" + TRIM$(Str$(NOREPARA&))
       DORFA$ = TRIM$(Mid$(XX$, 5, 60))
       FEMI% = CVI(Mid$(XX$, 121, 2))
       '
       ZZ$ = REBLA$
       Call REPLA(ZZ$, NORFA$, 1, 12)
       Call REPLA(ZZ$, DORFA$, 13, 48)
       Call REPLA(ZZ$, MKI$(FEMI%), 61, 2)
       JJ1& = JJ1& + 1
       Call GRAREG("INDITOTT", ZZ$, JJ1&)
       '
       If CVI(Mid$(XX$, 123, 2)) < 1 Then  ' NO ESTA CERRADA
         JJ2& = JJ2& + 1
         Call GRAREG("INDITOTP", ZZ$, JJ2&)
       End If
     End If
   Next U&
   '
   FIN& = ULTREG&("INDIOTRA")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     ZZ$ = REGLEIDO$("INDIOTRA", U&)
     JJ1& = JJ1& + 1
     Call GRAREG("INDITOTT", ZZ$, JJ1&)
     '
     If CVI(Mid$(ZZ$, 63, 2)) < 1 Then  ' NO ESTA CERRADA
       JJ2& = JJ2& + 1
       Call GRAREG("INDITOTP", ZZ$, JJ2&)
     End If
   Next U&
   '
   Call SETULTREG("INDITOTT", JJ1&)
   Call SETULTREG("INDITOTP", JJ2&)
   '
   Call FILESORT("INDITOTT", "", 3, 3, "ASC")
   Call FILESORT("INDITOTP", "", 3, 3, "ASC")
   Screen.MousePointer = 1
   '
End Sub
