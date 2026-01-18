VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form INFOPOFCE 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Informe de Operaciones Sobre O/F Cerradas"
   ClientHeight    =   7428
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6552
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7428
   ScaleWidth      =   6552
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   384
      ScaleWidth      =   1128
      TabIndex        =   50
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Parte de O/F a Informar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   44
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton Command3 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5070
         TabIndex        =   2
         Top             =   1080
         Width           =   175
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3145
         TabIndex        =   0
         Top             =   240
         Width           =   1890
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5040
         TabIndex        =   1
         Top             =   240
         Width           =   175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   49
         Top             =   660
         Width           =   5010
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3405
         TabIndex        =   48
         Top             =   1080
         Width           =   1665
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1770
         TabIndex        =   47
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de O/F: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   46
         Top             =   1155
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   45
         Top             =   1440
         Width           =   5010
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Nuevos Valores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   27
      Top             =   2160
      Width           =   5415
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   28
         Top             =   360
         Width           =   630
      End
      Begin VB.CommandButton Command6 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   175
      End
      Begin VB.CommandButton Command7 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1935
         TabIndex        =   5
         Top             =   840
         Width           =   175
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   990
         TabIndex        =   4
         Top             =   840
         Width           =   960
      End
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   7
         Top             =   1440
         Width           =   175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   6
         Top             =   1440
         Width           =   480
      End
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   9
         Top             =   1920
         Width           =   175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   8
         Top             =   1920
         Width           =   480
      End
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   11
         Top             =   2400
         Width           =   175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   10
         Top             =   2400
         Width           =   480
      End
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   13
         Top             =   2880
         Width           =   175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   12
         Top             =   2880
         Width           =   480
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2535
         TabIndex        =   14
         Text            =   " "
         Top             =   3480
         Width           =   960
      End
      Begin VB.CommandButton Command5 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         TabIndex        =   15
         Top             =   3480
         Width           =   175
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4575
         TabIndex        =   16
         Text            =   " "
         Top             =   3480
         Width           =   720
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   17
         Text            =   " "
         Top             =   3885
         Width           =   1065
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         TabIndex        =   20
         Text            =   " "
         Top             =   3885
         Width           =   480
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3720
         TabIndex        =   19
         Text            =   " "
         Top             =   3885
         Width           =   480
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   18
         Text            =   " "
         Top             =   3885
         Width           =   480
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   43
         Top             =   360
         Width           =   3090
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Operación:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -1200
         TabIndex        =   42
         Top             =   420
         Width           =   2325
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Equipo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   41
         Top             =   915
         Width           =   855
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2250
         TabIndex        =   40
         Top             =   840
         Width           =   3000
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Operarios:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   80
         TabIndex        =   39
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   38
         ToolTipText     =   "Haga Doble-Click para Informe por Célula"
         Top             =   1440
         Width           =   3345
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1920
         TabIndex        =   37
         ToolTipText     =   "Haga Doble-Click para Informe por Célula"
         Top             =   1920
         Width           =   3345
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   1920
         TabIndex        =   36
         ToolTipText     =   "Haga Doble-Click para Informe por Célula"
         Top             =   2400
         Width           =   3345
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   1920
         TabIndex        =   35
         ToolTipText     =   "Haga Doble-Click para Informe por Célula"
         Top             =   2880
         Width           =   3345
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1800
         TabIndex        =   34
         Top             =   3555
         Width           =   705
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hora:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3840
         TabIndex        =   33
         Top             =   3555
         Width           =   705
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -240
         TabIndex        =   32
         Top             =   3960
         Width           =   1065
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Seg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3960
         TabIndex        =   31
         Top             =   3915
         Width           =   705
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3000
         TabIndex        =   30
         Top             =   3915
         Width           =   705
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Horas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   29
         Top             =   3915
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00004080&
      Height          =   8460
      Left            =   5640
      ScaleHeight     =   8412
      ScaleWidth      =   1020
      TabIndex        =   21
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   120
         Picture         =   "INFOPOFCE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   845
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   120
         Picture         =   "INFOPOFCE.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Salir"
         Top             =   200
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   84
         ScaleWidth      =   600
         TabIndex        =   24
         Top             =   6120
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   312
         Left            =   -288
         Picture         =   "INFOPOFCE.frx":0454
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   6480
         Width           =   1212
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "O.K."
      Top             =   6960
      Width           =   1485
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "INFOPOFCE.frx":2376
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "INFOPOFCE.frx":25AA
      TabIndex        =   51
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
End
Attribute VB_Name = "INFOPOFCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   Call SELECHO("MASTER")
   If VALACT1$("MASTER") <> "" Then
     Text10 = VALACT1$("MASTER")
   End If
   Command1.Enabled = True
End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
    Call SELECHO("TAMOTIV")
    If XVALO(VALACT1("TAMOTIV")) < 1 Then GoTo 99
    Text19 = XVALO(VALACT1("TAMOTIV"))
99  Command14.Enabled = True
End Sub

Private Sub Command2_Click()
     Command2.Enabled = False
    '
     CI% = CODINT(Text10)
     If CI% < 1 Then
        Call MENSERR%(24, "Codigo Articulo Invalido")
        GoTo 99
     End If
     '
     If Label4 = "" Then
        Call MENSERR%(24, "Debe Seleccionar la Orden de Fabricación.")
        GoTo 99
     End If
     '
     CONDI1$ = "CodIConj = " & CI% & " and  NumeOper = " & XVALO(Text3)
     If VALOBADA("SECOPER", "DescOper", CONDI1$) = "" Then
         Call MENSERR(24, "Operación Incorrecta para el Articulo Seleccionado.")
         GoTo 99
     End If
     '
     If TRIM$(ECOARCH$("PADMAQ", TRIM$(Text5))) = "" Then
         Call MENSERR(24, "Equipo Incorrecto.")
         GoTo 99
     End If
     '
     LEGA$ = ""
     OPERS% = 0
     For i% = 0 To 3
        If Text1(i%) <> "" Then
            OPERS% = 1
            LEGAJO% = XVALO(Text1(i%))
            If LEGAJO% > 0 Then
                If ECOARCH$("PADRON", MKI$(LEGAJO%)) = "" Then
                  Call MENSERR(24, "Legajo Operario no Válido")
                  GoTo 99
                Else
                  LEGA$ = LEGA$ & MKI$(LEGAJO%)
                End If
            End If
         End If
     Next i%
     If OPERS% = 0 Then
         Call MENSERR(24, "Falta Definir Operarios")
         GoTo 99
     End If
     '
     NOPE% = XVALO(Text3)
     FEOP% = FECHANUM(Text4)
     HOOP% = HORANUM(Text2)
     HOII% = HOY(HOS$)
     If FEOP = 0 Then
        Call MENSERR(24, "Fecha Incorrecta.")
        GoTo 99
     End If
     If FEOP% > HOII% Then
           Call MENSERR(24, "Fecha No Puede Ser Mayor a la Actual.")
           GoTo 99
         ElseIf FEOP% = HOII% Then
           If HOOP% > HORANUM(Time$) Then
             Call MENSERR(24, "Hora Incorrecta")
             GoTo 99
           End If
     End If
     FEX = FEOP%
     FEHOLI = CDate(FECHATEX$(FEX) + " " + HORATEX$(HOOP%))
     '
     CONDI$ = " IdSubOr = '" & Label4.Caption & "'"
     CAPROG = XVALO(VALOBADA("ORDEFABR", "CanProy", CONDI$))
     '
     CAINF# = XVALO(Text7)
     If CAINF# = 0 Then
        Call MENSERR(24, "Cantidad Invalida")
        GoTo 99
     End If
     '
     TIEINF = 60 * XVALO(Text11) + XVALO(Text9) + XVALO(Text8) / 60
     '
     If TIEINF <= 0 Then 'PARA evitar OVERFLOW
        Call COMUNI("Debe Ingresar un Tiempo Expresado en: \ Horas, Minutos y Segundos")
        GoTo 99
     End If
     '
''     If UCase$(TRIM$(Text6)) <> "N" And UCase$(TRIM$(Text6)) <> "S" Then
''        Call MENSERR(24, "Solo Valido 'S'(Re-Trabajo) o 'N'(No Re-Trabajo)")
''        GoTo 99
''     End If
     '
     
     RETRAB$ = ""
     NUMMOTIV% = 0
     Stat% = 0
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     If InStr(EPAC$, "ALEJPA") > 0 Then
        VDEF$ = ""
        Call REPLA(VDEF$, TRIM$(Label4), 1, 12)
        Call REPLA(VDEF$, MKI$(CI%), 13, 2)
        Call REPLA(VDEF$, "N", 15, 1) ' VALOR POR DEFAULT
        '
1425    VDEF$ = OBJPLA$("RETRAMO", VDEF$)
        If VDEF$ = "" Then GoTo 99
        '
        If DERACCE%("GENESCRAP", "Informe de Produccion con Motivo Scrap") > 1 Then
            RETRAB$ = TRIM$(Mid$(VDEF$, 15, 1))
            NUMMOTIV% = Asc(Mid$(VDEF$, 16, 1))
            '
            If UCase$(RETRAB$) <> "N" And UCase$(RETRAB$) <> "S" Then
               Call MENSERR(24, "Solo Valido 'S'(Re-Trabajo) o 'N'(No Re-Trabajo)")
               GoTo 1425
            End If
            '
            If UCase$(RETRAB$) = "S" And (NUMMOTIV% < 1 Or ECOARCH$("TAMOTIV", Chr$(NUMMOTIV%)) = "") Then
               Call MENSERR(24, "Ingreso de Motivo de Rechazo No Válido")
               GoTo 1425
            End If
            '
            TX$ = FILTERGETRECORD$("TAMOTIV", 1, Chr$(NUMMOTIV%))
            If Asc(Mid$(TX$, 32, 1)) > 0 Then ' SI ESTA MARCADO ESTE MOTIVO EN LA COLUMNA DE SCRAP CON 1
               Stat% = 8
            End If
            '
            If UCase$(RETRAB$) = "N" And Stat% = 8 Then
               Call MENSERR(24, "Ingreso de Motivo de Scrap\Solo si es Retrabajo")
               GoTo 1425
            End If
            '
            If Stat% = 8 Then
               Call MENSERR(24, "Imposible Grabar Motivo de Rechazo Asociada a Scrap\Orden de Fabricación Cerrada")
               GoTo 1425
            End If
     End If
     End If

''     NUMMOTIV% = XVALO(Text19)
''     If UCase$(Text6) = "S" And (NUMMOTIV% < 1 Or ECOARCH$("TAMOTIV", Chr$(NUMMOTIV%)) = "") Then
''       Call MENSERR(24, "Ingreso de Motivo de Rechazo No Válido")
''       GoTo 99
''     End If
''     '
''     TX$ = FILTERGETRECORD$("TAMOTIV", 1, Chr$(NUMMOTIV%))
''     If Asc(Mid$(TX$, 32, 1)) > 0 Then ' SI ESTA MARCADO ESTE MOTIVO EN LA COLUMNA DE SCRAP CON 1
''        Call MENSERR(24, "Imposible Grabar Motivo de Rechazo Asociada a Scrap\ Orden de Fabricación Cerrada")
''        GoTo 99
''     End If

1430 Call ABREORFAB
     SQLX$ = " SELECT * FROM INFOFAB "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & Label4 & "' "
     SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & " "
     '
     CAINFANT% = 0
     Dim INFOFABTMP As ADODB.Recordset
     Set INFOFABTMP = New ADODB.Recordset
     INFOFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

     With INFOFABTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           ' CAREAL# = CAREAL# + XVALO(!CANTREAL)
           ' TIEINF = TIEINF + XVALO(!TiemEjec)   ' deshabilitado porque si no acumulaba en el nuevo parte
           CAINFANT% = CAINFANT% + 1
         .MoveNext
         Loop
       End If
     End With
     Set INFOFABTMP = Nothing
     
     Call ABREORFAB
     SQLX$ = " SELECT * FROM InfoFab "
     Dim INFOFFFAB As ADODB.Recordset
     Set INFOFFFAB = New ADODB.Recordset
     INFOFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     With INFOFFFAB
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        !IdSubOr = TRIM$(Label4)
        !DescriOF = TRIM$(Label3)
        !NuOrInf = CAINFANT%
        !Cod_Inte = CI%
        !Cod_Exte = CODEXT$(CI%)
        !Descrip = Left$(DESCRIT$(CI%), 64)
        !NUMEOPER = NOPE%
        !AltOper = 0
        MARETRA% = 0: If RETRAB$ = "S" Then MARETRA% = 1
        !RETRABA = MARETRA%
        !DESCOPER = TRIM$(Label7)
        !CanOrPro = CAPROG
        !CANTREAL = CAINF#
        !MAQUINA = Text5
        !DescMaq = TRIM$(Label10)
        !LegaOps = LEGA$
        !FeHorLib = FEHOLI
        !Hora_Des = 0
        !Hora_Has = HOOP%
        !MaqLiber = "S"  ' LA MAQUINA ESTA LIBERADA PORQUE LA O/F ESTA CERRADA
        !TiemEjec = TIEINF
           MODO% = CAINFANT%
        !TiemStan = TIESTAN(CI%, NOPE%, CAINF#, MODO%)
        RENDI = 999.9: If !TiemStan > 0 Then RENDI = 100 * !TiemEjec / !TiemStan
        If RENDI > 999.9 Then RENDI = 999.9
        !RENDIM = RENDI
        !Observa = ""
        !Referen = " "
        !MOTIRECHA = NUMMOTIV%
       .Update
     End With
     Set INFOFFFAB = Nothing
     '
     Call MENSERR(24, "Informe de Operaciones Registrado Satisfactoriamente")
     Call BLANFORMU
     Text10 = ""
99   Command2.Enabled = True
End Sub

Private Sub Command3_Click()
    '
    CIXI% = CODINT%(TRIM$(Text10))
    If CIXI% <= 0 Then
        Call MENSERR(24, "Primero debe Seleccionar un Articulo")
        Exit Sub
    End If
    '
    CONDI$ = "COD_INTE = " & CIXI% & " AND StatuOrf = 3"
    REG& = CantRegistros&("ORDEFABR", CONDI$)
    If REG& < 1 Then
        Call MENSERR(24, "No Existen Ordenes de Fabricación\Para este Articulo")
        Exit Sub
    End If
    '
    Call SELECHO("!ECORDEF1/Ordenes de Fabricacion Cerradas")
    If VALACT1$("!ECORDEF1") <> "" Then
     Label4 = VALACT1$("!ECORDEF1")
     Label2 = VALACT2$("!ECORDEF1")
    End If
    '
End Sub

Private Sub Command4_Click(Index As Integer)
  '
  Call SELECHO("PADRON")
  VDEV$ = VALACT1$("PADRON")
  If VDEV$ <> "" Then
    Text1(Index) = TRIM$(VDEV$)
    Label5(Index) = TRIM$(VALACT2$("PADRON"))
  End If
  '
End Sub

Private Sub Command5_Click()
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
    Text4.TEXT = VDEV$
  End If
End Sub

Private Sub Command6_Click()
    '
    CIXI% = CODINT%(TRIM$(Text10))
    If CIXI% <= 0 Then
        Call MENSERR(24, "Primero debe Seleccionar un Articulo")
        Exit Sub
        If TRIM$(Label4) = "" Then
            Call MENSERR(24, "Primero debe Seleccionar la Orden de Fabricacion")
            Exit Sub
        End If
    End If
    '
    CONDI$ = "CodIConj = " & CIXI%
    REG& = CantRegistros&("SECOPER", CONDI$)
    If REG& < 1 Then
        Call MENSERR(24, "No Existen Operaciones de Fabricación\Para este Articulo")
        Exit Sub
    End If
    '
    Call SELECHO("!OPERINF1/Operaciones de Ordenes de Fabricacion")
    If TRIM$(VALACT1$("!OPERINF1")) <> "" Then
      Text3.TEXT = TRIM$(VALACT1$("!OPERINF1"))
      Label7.Caption = TRIM$(VALACT2$("!OPERINF1"))
    End If
    '
End Sub

Private Sub Command7_Click()
   '
   Call SELECHO("PADMAQ")
   VDEVU$ = VALACT1$("PADMAQ")
   If VDEVU$ <> "" Then
      Text5.TEXT = TRIM$(VDEVU$)
      Label10.Caption = TRIM$(VALACT2$("PADMAQ"))
   End If
   '
End Sub

Private Sub Command8_Click()
    Call CIERRARCH("*.*")
    Unload Me
End Sub



Private Sub Command9_Click()
   
End Sub

Private Sub Form_Load()
    '
    Text6 = "N"    ' por defecto no es Re-trabajo
    '
    Call APLICACIONSKINS(Me, Picture14)

End Sub

Private Sub Label19_Click()

End Sub

Private Sub Text1_Change(Index As Integer)
    Label5(Index) = ECOARCH$("PADRON", MKI$(XVALO(Text1(Index))))
End Sub

Private Sub Text1_DblClick(Index As Integer)
    Call Command4_Click(Index)
End Sub

Private Sub Text10_Change()
   '
   CIXXI% = CODINT%(Text10)
   Label3 = DESCRIT0$(CIXXI%)
   '
   Call BLANFORMU
   '
   Call COPYSTRU("OPERINF", "OPERINF1")
   CONDI1$ = "CodIConj = " & CIXXI%
   Call CARGALISEL("!OPERINF1", "SECOPER", "NumeOper/F2;DescOper/A32", CONDI1$)
   Call COPYSTRU("ECORDEF", "ECORDEF1")
   CONDI2$ = "COD_INTE = " & CIXXI% & " AND StatuOrf = 3"
   Call CARGALISEL("!ECORDEF1", "ORDEFABR", "IdSubOr/A12;Cod_Exte/a16;Descrip/a32", CONDI2$)
   '
End Sub

Private Sub Text10_DblClick()
    Call Command1_Click
End Sub

Private Sub Text19_Change()
    On Error Resume Next
    TXT$ = TRIM$(ECOARCH$("TAMOTIV", Chr$(Text19)))
    On Error GoTo 0
    If TXT$ <> "" Then
       Label55 = TXT$
    Else
       Label55 = ""
    End If
End Sub

Private Sub Text2_LostFocus()
  Text2 = HORATEX$(HORANUM(Text2.TEXT))
End Sub

Private Sub Text3_DblClick()
    Call Command6_Click
End Sub

Private Sub Text4_DblClick()
    Call Command5_Click
End Sub

Private Sub Text5_DblClick()
    Call Command7_Click
End Sub

Private Sub Text7_LostFocus()
    Text7 = FORMATNUM$(XVALO(Text7), "F8.1")
End Sub


Sub BLANFORMU()
   '
   Label4 = ""
   Label2 = ""
   Text3 = ""
   Label7 = ""
   Label10 = ""
   Text5 = ""
   Label7 = ""
   Text4 = ""
   Text2 = ""
   Text6 = ""
   Text7 = ""
   Text11 = ""
   Text9 = ""
   Text8 = ""
   Text19 = ""
   Label55 = ""
   For i% = 0 To 3
     Text1(i%) = ""
     Label5(i%) = ""
   Next i%
   '
End Sub
