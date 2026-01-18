VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form FISECOM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguimiento Comercial FLEXOFT"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   Icon            =   "FISECOM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   10920
      ScaleHeight     =   75
      ScaleWidth      =   675
      TabIndex        =   112
      Top             =   7560
      Width           =   740
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   113
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Command30"
      Height          =   330
      Left            =   5670
      TabIndex        =   110
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Command29"
      Height          =   330
      Left            =   3570
      TabIndex        =   109
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Erase"
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
      Left            =   10935
      Picture         =   "FISECOM.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   99
      ToolTipText     =   "Elimina Contacto Activo de la Base"
      Top             =   2100
      Width           =   750
   End
   Begin VB.Frame Frame4 
      Caption         =   "Misc."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1950
      Left            =   9660
      TabIndex        =   89
      Top             =   6300
      Width           =   1065
      Begin VB.CommandButton Command17 
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   740
         Left            =   105
         Picture         =   "FISECOM.frx":157C
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   "Documentos Conexos"
         Top             =   1155
         Width           =   850
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   740
         Left            =   105
         Picture         =   "FISECOM.frx":1886
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Abrir Anotador de Novedades"
         Top             =   315
         Width           =   850
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   4425
      Left            =   105
      TabIndex        =   44
      Top             =   105
      Width           =   10620
      Begin VB.Frame Frame7 
         Caption         =   "Representante"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   5355
         TabIndex        =   114
         Top             =   2120
         Width           =   1695
         Begin VB.TextBox Text16 
            Alignment       =   1  'Right Justify
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
            Left            =   1050
            TabIndex        =   116
            Top             =   150
            Width           =   540
         End
         Begin VB.Label Label40 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   105
            TabIndex        =   115
            Top             =   400
            Width           =   1485
         End
      End
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
         Height          =   280
         Left            =   7500
         Picture         =   "FISECOM.frx":1B90
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Marca / Quita Prioridad"
         Top             =   420
         Width           =   285
      End
      Begin VB.TextBox Text15 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   7245
         TabIndex        =   111
         Text            =   " "
         Top             =   420
         Width           =   330
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   9345
         TabIndex        =   108
         Text            =   " "
         Top             =   1050
         Width           =   1065
      End
      Begin VB.CommandButton Command28 
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
         Height          =   285
         Left            =   8925
         TabIndex        =   104
         Top             =   1050
         Width           =   150
      End
      Begin VB.TextBox Text13 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1155
         TabIndex        =   102
         Text            =   " "
         ToolTipText     =   "Doble-Click para Enviar un Mail"
         Top             =   1050
         Width           =   6630
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   8820
         TabIndex        =   98
         Top             =   4070
         Width           =   790
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   9620
         TabIndex        =   97
         Top             =   4070
         Width           =   790
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   8820
         Sorted          =   -1  'True
         TabIndex        =   95
         ToolTipText     =   "Doble Click para Actualizar Tabla"
         Top             =   3360
         Width           =   1590
      End
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
         Height          =   540
         Left            =   9450
         Picture         =   "FISECOM.frx":1FD2
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Selección Contactos Programados"
         Top             =   420
         Width           =   480
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
         Height          =   945
         Left            =   1155
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   3360
         Width           =   7260
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   9660
         TabIndex        =   17
         Text            =   " "
         Top             =   2940
         Width           =   750
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
         Height          =   280
         Left            =   1155
         TabIndex        =   15
         Text            =   " "
         Top             =   2940
         Width           =   2115
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   4830
         TabIndex        =   16
         Text            =   " "
         Top             =   2940
         Width           =   3585
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   9930
         Picture         =   "FISECOM.frx":2414
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Búsqueda Secuencial por Texto"
         Top             =   420
         Width           =   480
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1155
         TabIndex        =   5
         Text            =   " "
         ToolTipText     =   "Doble-Click para Enviar un Mail"
         Top             =   735
         Width           =   6630
      End
      Begin VB.Frame Frame5 
         Caption         =   "Nro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   7875
         TabIndex        =   65
         Top             =   315
         Width           =   1520
         Begin VB.CommandButton Command27 
            Caption         =   ">"
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
            Left            =   1280
            Picture         =   "FISECOM.frx":271E
            TabIndex        =   101
            ToolTipText     =   "Selecciona Contacto - Lista Completa"
            Top             =   260
            Width           =   160
         End
         Begin VB.CommandButton Command26 
            Caption         =   "<"
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
            Left            =   1120
            Picture         =   "FISECOM.frx":2A28
            TabIndex        =   100
            ToolTipText     =   "Selecciona Contacto - Lista Completa"
            Top             =   260
            Width           =   160
         End
         Begin VB.CommandButton Command16 
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
            Height          =   285
            Left            =   885
            TabIndex        =   0
            ToolTipText     =   "Selecciona Contacto - Lista Completa"
            Top             =   260
            Width           =   175
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   66
            Top             =   260
            Width           =   750
         End
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1155
         TabIndex        =   4
         Text            =   " "
         Top             =   420
         Width           =   6420
      End
      Begin VB.TextBox Text26 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   7875
         TabIndex        =   12
         Text            =   " "
         Top             =   1890
         Width           =   2150
      End
      Begin VB.TextBox Text24 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   7875
         TabIndex        =   14
         Text            =   " "
         Top             =   2520
         Width           =   2150
      End
      Begin VB.TextBox Text23 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1155
         TabIndex        =   10
         Text            =   " "
         Top             =   2520
         Width           =   4005
      End
      Begin VB.TextBox Text22 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   3675
         TabIndex        =   8
         Text            =   " "
         Top             =   1785
         Width           =   3375
      End
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1155
         TabIndex        =   7
         Top             =   1785
         Width           =   1485
      End
      Begin VB.TextBox Text20 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1155
         TabIndex        =   6
         Text            =   " "
         Top             =   1470
         Width           =   5895
      End
      Begin VB.TextBox Text19 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1155
         TabIndex        =   9
         Text            =   " "
         Top             =   2205
         Width           =   4005
      End
      Begin VB.TextBox Text18 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   7875
         TabIndex        =   11
         Text            =   " "
         Top             =   1470
         Width           =   2535
      End
      Begin VB.TextBox Text17 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   7875
         TabIndex        =   13
         Text            =   " "
         Top             =   2205
         Width           =   2535
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "M:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9135
         TabIndex        =   107
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label Label38 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8250
         TabIndex        =   106
         Top             =   1050
         Width           =   680
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Baja:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7875
         TabIndex        =   105
         ToolTipText     =   "Doble-Click para Re-Activar"
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         Caption         =   "Web-Site:"
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
         Left            =   105
         TabIndex        =   103
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Gr:"
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
         Left            =   7840
         TabIndex        =   96
         ToolTipText     =   "Doble Click para Actualizar Tabla"
         Top             =   3360
         Width           =   960
      End
      Begin VB.Line Line1 
         X1              =   1155
         X2              =   8400
         Y1              =   4305
         Y2              =   4305
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Observac:"
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
         Left            =   105
         TabIndex        =   71
         Top             =   3360
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Empleados:"
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
         Left            =   8505
         TabIndex        =   70
         Top             =   3045
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Ramo:"
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
         Left            =   105
         TabIndex        =   69
         Top             =   3045
         Width           =   960
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Princ.Producto"
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
         Left            =   2940
         TabIndex        =   68
         Top             =   3045
         Width           =   1800
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Dir-elec:"
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
         Left            =   105
         TabIndex        =   67
         Top             =   840
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Empresa:"
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
         Left            =   105
         TabIndex        =   64
         Top             =   525
         Width           =   960
      End
      Begin VB.Image Image5 
         Height          =   300
         Left            =   10080
         Picture         =   "FISECOM.frx":2D32
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   315
      End
      Begin VB.Image Image4 
         Height          =   300
         Left            =   10080
         Picture         =   "FISECOM.frx":303C
         Stretch         =   -1  'True
         Top             =   1890
         Width           =   315
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Calle y Nº:"
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
         Left            =   105
         TabIndex        =   52
         Top             =   1575
         Width           =   960
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Ciudad:"
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
         Left            =   2625
         TabIndex        =   51
         Top             =   1890
         Width           =   960
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "País:"
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
         Left            =   105
         TabIndex        =   50
         Top             =   2625
         Width           =   960
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Telef:"
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
         Left            =   6825
         TabIndex        =   49
         Top             =   1995
         Width           =   960
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "C.Post:"
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
         Left            =   105
         TabIndex        =   48
         Top             =   1890
         Width           =   960
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Fax:"
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
         Left            =   6825
         TabIndex        =   47
         Top             =   2625
         Width           =   960
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Provincia:"
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
         Left            =   105
         TabIndex        =   46
         Top             =   2310
         Width           =   960
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Origen:"
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
         Left            =   6825
         TabIndex        =   45
         Top             =   1575
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Novedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1950
      Left            =   105
      TabIndex        =   77
      Top             =   6300
      Width           =   9460
      Begin VB.CommandButton Command15 
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
         Height          =   280
         Left            =   8295
         TabIndex        =   87
         Top             =   1470
         Width           =   175
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   7350
         TabIndex        =   88
         Top             =   1470
         Width           =   1000
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0C0&
         Height          =   425
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   9360
         TabIndex        =   78
         Top             =   315
         Width           =   9420
         Begin VB.Label Label14 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Detalle                                                    Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   80
            Top             =   75
            Width           =   9360
         End
         Begin VB.Label Label13 
            Caption         =   "Label5"
            Height          =   15
            Left            =   0
            TabIndex        =   79
            Top             =   420
            Width           =   1380
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   1155
            Y1              =   420
            Y2              =   420
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "O.K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   745
         Left            =   8670
         Picture         =   "FISECOM.frx":3346
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1140
         Width           =   750
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1175
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   735
         Width           =   7155
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Label35"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7245
         TabIndex        =   86
         Top             =   840
         Width           =   4005
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " "
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
         Left            =   80
         TabIndex        =   85
         Top             =   800
         Width           =   1065
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Próx.Contacto:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7245
         TabIndex        =   84
         Top             =   1200
         Width           =   1380
      End
      Begin VB.Line Line4 
         X1              =   7140
         X2              =   9450
         Y1              =   1125
         Y2              =   1125
      End
      Begin VB.Label Label32 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   7140
         TabIndex        =   83
         Top             =   735
         Width           =   2310
      End
      Begin VB.Label Label31 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1165
         Left            =   0
         TabIndex        =   82
         Top             =   735
         Width           =   1170
      End
      Begin VB.Line Line5 
         X1              =   8295
         X2              =   8295
         Y1              =   735
         Y2              =   1920
      End
      Begin VB.Line Line6 
         X1              =   8295
         X2              =   10605
         Y1              =   1130
         Y2              =   1130
      End
      Begin VB.Line Line7 
         X1              =   9450
         X2              =   9450
         Y1              =   315
         Y2              =   1900
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
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
         Left            =   8400
         TabIndex        =   81
         Top             =   840
         Width           =   4000
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   11665
         Y1              =   1900
         Y2              =   1900
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contactos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1695
      Left            =   105
      TabIndex        =   72
      Top             =   4560
      Width           =   10620
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   2
         Left            =   9135
         TabIndex        =   30
         Top             =   1260
         Width           =   1275
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   2
         Left            =   7140
         TabIndex        =   29
         Top             =   1260
         Width           =   1800
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   2
         Left            =   4830
         TabIndex        =   28
         Top             =   1260
         Width           =   2115
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   2
         Left            =   1155
         TabIndex        =   27
         Top             =   1260
         Width           =   3480
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   1
         Left            =   9135
         TabIndex        =   26
         Top             =   945
         Width           =   1275
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   1
         Left            =   7140
         TabIndex        =   25
         Top             =   945
         Width           =   1800
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   1
         Left            =   4830
         TabIndex        =   24
         Top             =   945
         Width           =   2115
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   1
         Left            =   1155
         TabIndex        =   23
         Top             =   945
         Width           =   3480
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   0
         Left            =   9135
         TabIndex        =   22
         Top             =   630
         Width           =   1275
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   0
         Left            =   7140
         TabIndex        =   21
         Top             =   630
         Width           =   1800
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   0
         Left            =   4830
         TabIndex        =   20
         Top             =   630
         Width           =   2115
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   0
         Left            =   1155
         TabIndex        =   19
         Top             =   630
         Width           =   3480
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Directo / Interno:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8085
         TabIndex        =   76
         Top             =   350
         Width           =   2010
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Sector:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6510
         TabIndex        =   75
         Top             =   350
         Width           =   960
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Cargo / Función:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4305
         TabIndex        =   74
         Top             =   350
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Persona de Contacto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   73
         Top             =   350
         Width           =   1905
      End
   End
   Begin VB.CommandButton Command21 
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
      Height          =   600
      Left            =   10935
      Picture         =   "FISECOM.frx":3650
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Alta de Nueva Ficha de Seguimiento"
      Top             =   840
      Width           =   750
   End
   Begin VB.CommandButton Command3 
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
      Height          =   600
      Left            =   10935
      Picture         =   "FISECOM.frx":379A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   210
      Width           =   750
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   10815
      Sorted          =   -1  'True
      TabIndex        =   63
      Top             =   735
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Sort"
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
      Left            =   10935
      Picture         =   "FISECOM.frx":38E4
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Ordenamiento Alfabético"
      Top             =   2730
      Width           =   750
   End
   Begin VB.Frame Frame6 
      Caption         =   "Exp/Imp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   10815
      TabIndex        =   62
      Top             =   3465
      Width           =   990
      Begin VB.CommandButton Command12 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   105
         TabIndex        =   38
         ToolTipText     =   "Importa Libreta de Direcciones (de C:\FLEXOFT)"
         Top             =   650
         Width           =   330
      End
      Begin VB.CommandButton Command10 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   525
         TabIndex        =   39
         ToolTipText     =   "Exporta Direcciones para Outlook Express"
         Top             =   650
         Width           =   330
      End
      Begin VB.CommandButton Command8 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   525
         TabIndex        =   37
         ToolTipText     =   "Restaura Copia de Seguridad (Zip)"
         Top             =   315
         Width           =   330
      End
      Begin VB.CommandButton Command7 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   105
         TabIndex        =   36
         ToolTipText     =   "Confecciona Copia de Seguridad (Zip)"
         Top             =   315
         Width           =   330
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00D0FFFF&
      Height          =   2535
      Left            =   2835
      ScaleHeight     =   2475
      ScaleWidth      =   4155
      TabIndex        =   54
      Top             =   1890
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton Command6 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3885
         TabIndex        =   61
         Top             =   -10
         Width           =   270
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   1260
         TabIndex        =   60
         Top             =   2100
         Width           =   1485
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Levante el Receptor !!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   330
         Left            =   420
         TabIndex        =   59
         Top             =   1680
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   210
         TabIndex        =   58
         Top             =   1365
         Width           =   3795
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "45241284"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   420
         TabIndex        =   57
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Llamando a:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   315
         TabIndex        =   56
         Top             =   420
         Width           =   1800
      End
      Begin VB.Label Label25 
         BackColor       =   &H00800000&
         Caption         =   "Monitor de Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Width           =   4110
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H80000000&
      Caption         =   "Varios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   10815
      TabIndex        =   53
      Top             =   4620
      Width           =   990
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
         Height          =   580
         Left            =   105
         Picture         =   "FISECOM.frx":3BEE
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Administrar Mail Circular"
         Top             =   2205
         Width           =   750
      End
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
         Left            =   105
         Picture         =   "FISECOM.frx":3EF8
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "Listados por Impresora"
         Top             =   1575
         Width           =   750
      End
      Begin VB.CommandButton Command25 
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
         Left            =   105
         Picture         =   "FISECOM.frx":4202
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Imprime Ficha del Contacto"
         Top             =   945
         Width           =   750
      End
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
         Height          =   580
         Left            =   105
         Picture         =   "FISECOM.frx":486C
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   315
         Width           =   750
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
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
      Left            =   10935
      Picture         =   "FISECOM.frx":4CAE
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Guarda los Cambios Realizados"
      Top             =   1470
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10350
      Picture         =   "FISECOM.frx":4FB8
      Top             =   7730
      Width           =   2010
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Apel:"
      Height          =   225
      Left            =   210
      TabIndex        =   43
      Top             =   1785
      Width           =   960
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Apel:"
      Height          =   225
      Left            =   210
      TabIndex        =   42
      Top             =   1260
      Width           =   960
   End
End
Attribute VB_Name = "FISECOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cancelflag

Private Sub Command11_Click()
   If Val(Label26) < 1 Then
     Command16.SetFocus
     Exit Sub
   End If
   '
   Command11.Enabled = False
   NPX& = Val(Label26)
   '
   If NPX& > 0 Then
     ANOTASEG.Label26 = Label26
     ANOTASEG.Label3 = Text4
     ANOTASEG.Label7 = Text6(0)
     ANOTASEG.Label8 = Text6(1)
     ANOTASEG.Label11 = Text1
     ANOTASEG.Label13 = Text2
     ANOTASEG.Label14 = TRIM$(Text3)
     ANOTASEG.Text12 = Text12
     '
     ANOTASEG.List1.Clear
     RFF$ = RECFILT$("ANOTASEG", 1, MKS$(NPX&))
     '
     For U& = 1 To Len(RFF$) Step 4
       RENOTA& = CVS(Mid$(RFF$, U&, 4))
       XX$ = REGLEIDO$("ANOTASEG", RENOTA&)
       For KU% = 5 To 92
         If Asc(Mid$(XX$, KU%)) < 32 Then
           Mid$(XX$, KU%, 1) = " "
         End If
      Next KU%
      ANOTASEG.List1.AddItem Mid$(XX$, 5, 92)
     Next U&
     '
     TPI = ANOTASEG.List1.ListCount - 12
     If TPI < 0 Then TPI = 0
     ANOTASEG.List1.TopIndex = TPI
     '
     ANOTASEG.Show 1
   End If
   '
   Command11.Enabled = True
   '
End Sub

Private Sub Command15_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text12.TEXT = VDEV$
    End If
End Sub

Private Sub Command16_Click()
   If TRIM$(FORSELEC.Text1) = "" Then
     FORSELEC.ABUSQUE = TRIM$(Label26)
   End If
   ABUSX$ = "BSEGCOM": If USNBR% < 100 Then ABUSX$ = "BSEGALE"
   Call SELECHO(ABUSX$)
   VDEVU$ = VALACT1$(ABUSX$)
   If VDEVU$ <> "" Then
     Label26 = TRIM$(VDEVU$)
   End If
End Sub

Private Sub Command17_Click()
   If Val(Label26) < 1 Then
     Command16.SetFocus
     Exit Sub
   End If
   '
   Command17.Enabled = False
   NPX& = Val(Label26)
   '
   If NPX& > 0 Then
     DOCONEC.Label26 = Label26
     DOCONEC.Label5 = Text4
     DOCONEC.Label7 = Text6(0)
     DOCONEC.Label8 = Text6(1)
     DOCONEC.Label4 = Text1
     DOCONEC.Label13 = Text2
     DOCONEC.Label14 = TRIM$(Text3)
     '
     DOCONEC.List1.Clear
     NPXI% = NPX&
     RFF$ = RECFILT$("DOCONEC", 1, MKI$(NPXI%))
     '
     For U& = 1 To Len(RFF$) Step 4
       RENOTA& = CVS(Mid$(RFF$, U&, 4))
       XX$ = REGLEIDO$("DOCONEC", RENOTA&)
       YY$ = Space$(256)
       Call REPLA(YY$, Mid$(XX$, 3, 56), 1, 56)
       Call REPLA(YY$, Mid$(XX$, 129, 56), 61, 56)
       Call REPLA(YY$, TRIM$(Str$(RENOTA&)), 129, 6)
       DOCONEC.List1.AddItem YY$
     Next U&
     '
     DOCONEC.Show 1
   End If
   '
   Command17.Enabled = True
End Sub

Private Sub Command18_Click()
   '
   Command18.Enabled = False
   OPX% = ALTERNA%("Base de Datos Completa\Registros Activos\Prospectos en Marcha")
   Select Case OPX%
     Case 1
        Call FINAL("*BSEGCOM")
     Case 2
        Call MENSERR(24, "No Disponible por el Momento")
     Case 3
        Call SEGUIOPER
   End Select
   Command18.Enabled = True
   '
End Sub

Private Sub Command19_Click()
   RCONTA$ = AJUSTI$(Text15, 3)
   If RCONTA$ = "!!!" Then
       RCONTA$ = "   "
     Else
       RCONTA$ = "!!!"
   End If
   Text15 = TRIM$(RCONTA$)
End Sub

Private Sub Command2_Click()
   '
   Static ARGUME$
   Command2.Enabled = False
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
10 ARGUME$ = TRIM$(UCase$(InputBox$("Buscar", "Búsqueda Cliente por Texto", ARGUME$)))
   '
   If ARGUME$ <> "" Then
     Screen.MousePointer = 11
     JJ& = 0
     ABUSX$ = "BSEGCOM": If USNBR% < 100 Then ABUSX$ = "BSEGALE"
     For U& = 1 To ULTREG&(ABUSX$)
       XXX$ = REGLEIDO$(ABUSX$, U&)
       If InStr(UCase$(Mid$(XXX$, 3)), ARGUME$) > 0 Then
         JJ& = JJ& + 1
         Call GRAREG("!SECONTAC", XXX$, JJ&)
       End If
     Next U&
     '
     Call SETULTREG("!SECONTAC", JJ&)
     Call CIERRARCH("!SECONTAC")
     Call BAJALDISCO
     Call FRESHECHO("!SECONTAC")
     '
20   If JJ& > 0 Then
       Call SELECHO("!SECONTAC")
       NPX& = Val(VALACT1$("!SECONTAC"))
       If NPX& > 0 Then
         Label26 = Str$(NPX&)
       End If
     End If
     '
   End If
   Screen.MousePointer = 1
   Command2.Enabled = True
End Sub

Private Sub Command20_Click()
   If USNBR% <> 101 Then Exit Sub
   Command20.Enabled = False
   ARMAICIRC.Show 1
   '
'Call MENSERR(24, "Falta Terminar")
'GoTo 99
'   '
'   HOII% = HOY(HOS$)
'   HORAI% = HORANUM%(Time$)
'   '
'   RFF$ = RECFILT$("MAILLST", 8, MKI$(-1))
'   RGG$ = ""
'   For U& = 1 To Len(RFF$) Step 4
'     REMA& = CVS(Mid$(RFF$, U&, 4))
'     XX$ = REGLEIDO$("MAILLST", REMA&)
'     FENVI% = CVI(Mid$(XX$, 93, 2)): HENVI% = CVI(Mid$(XX$, 95, 2))
'     If HOII% > FENVI% Or (HOII% = FENVI% And HORAI% >= FENVI%) Then
'       RGG$ = RGG$ + MKS$(RENA&)
'     End If
'   Next U&
'   '
'   If RGG$ = "" Then
'     Call COMUNI("No hay Mensajes Pendientes")
'     GoTo 99
'   End If
'   '
'   For U& = 1 To Len(RFF$) Step 4
'     REMA& = CVS(Mid$(RFF$, U&, 4))
'     XX$ = REGLEIDO$("MAILLST", REMA&)
'     MSGNBR& = CVS(Left$(XX$, 4))
'     RHH$ = RECFILT$("MAILSPL", 1, MKS$(MSGNBR&))
'     For U2& = 1 To Len(RHH$) Step 4
'       REM2& = CVS(Mid$(RHH$, U2&, 4))
'       YY$ = REGLEIDO$("MAILLST", REM2&)
'Next U2&
'Next U&
       
     
99 Command20.Enabled = True
End Sub

Private Sub Command22_Click()
   If List2.ListIndex >= 0 Then
     List2.RemoveItem (List2.ListIndex)
   End If
End Sub

Private Sub Command23_Click()
   If Val(Label26) < 1 Then
     Command16.SetFocus
     Exit Sub
   End If
   '
   XX$ = REGSEL$("TAGRUCON")
   If Len(XX$) > 4 Then
     RAGRE% = Val(CVI(Left$(XX$, 2)))
     TAGRE$ = AJUSTI$(Mid$(XX$, 51, 24), 24) + AJUSTI$(Mid$(XX$, 3, 48), 48) + Str$(RAGRE%)
     For U& = 1 To List2.ListCount
       If Val(Mid$(List2.List(U& - 1), 72, 6)) = RAGRE% Then Exit Sub
     Next U&
     List2.AddItem TAGRE$
   End If
End Sub

Private Sub Command24_Click()
   If USNBR% <> 101 Then Exit Sub
   Command24.Enabled = False
   If COMALTER%("Realmente Desea Eliminar PERMANENTEMENTE\la Ficha Activa ?\\No, Cancelar\Si, Eliminar") = 2 Then
5    Screen.MousePointer = 11
     IFICHA% = Val(Label26)
     '
     ABASE$ = "BSEGCOM": GoSub BORRABASE
     ABASE$ = "BSEGALE": GoSub BORRABASE
     ABASE$ = "DOCONEC": GoSub BORRABASE
     ABASE$ = "NOTASEG": GoSub BORRABASE
     '
     ABASE$ = "ANOTASEG": GoSub BORRABASLAR
     '
     Label26 = ""
   End If
   '
   Screen.MousePointer = 1
   Command24.Enabled = True
   Exit Sub
   '
   '
BORRABASE:
   JJ& = 0
   For U& = 1 To ULTREG&(ABASE$)
     XX$ = REGLEIDO$(ABASE$, U&)
     IFICHB% = CVI(Left$(XX$, 2))
     If IFICHB% <> IFICHA% Then
       If IFICHB% > IFICHA% Then IFICHB% = IFICHB% - 1
       Call REPLA(XX$, MKI$(IFICHB%), 1, 2)
       JJ& = JJ& + 1
       Call GRAREG(ABASE$, XX$, JJ&)
     End If
   Next U&
   Call SETULTREG(ABASE$, JJ&)
   Call CIERRARCH(ABASE$)
Return
   
BORRABASLAR:
   JJ& = 0
   For U& = 1 To ULTREG&(ABASE$)
     XX$ = REGLEIDO$(ABASE$, U&)
     IFICHC& = CVS(Left$(XX$, 4))
     If IFICHC& <> IFICHA% Then
       If IFICHC& > IFICHA% Then IFICHC& = IFICHC& - 1
       Call REPLA(XX$, MKS$(IFICHC&), 1, 4)
       JJ& = JJ& + 1
       Call GRAREG(ABASE$, XX$, JJ&)
     End If
   Next U&
   Call SETULTREG(ABASE$, JJ&)
   Call CIERRARCH(ABASE$)
   '
Return
'
'
End Sub

Private Sub Command26_Click()
   ABUSX$ = "BSEGCOM": If USNBR% < 100 Then ABUSX$ = "BSEGALE"
   UKU& = REGNUM&(ABUSX$)
   If UKU& > 1 Then
     XX$ = REGLEIDO$(ABUSX$, UKU& - 1)
     Label26 = TRIM$(Str$(CVI(Left$(XX$, 2))))
   End If
End Sub

Private Sub Command27_Click()
   ABUSX$ = "BSEGCOM": If USNBR% < 100 Then ABUSX$ = "BSEGALE"
   UKU& = REGNUM&(ABUSX$)
   If UKU& < ULTREG&(ABUSX$) Then
     XX$ = REGLEIDO$(ABUSX$, UKU& + 1)
     Label26 = TRIM$(Str$(CVI(Left$(XX$, 2))))
   End If
End Sub

Private Sub Command28_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Label38 = VDEV$
    End If
End Sub

Private Sub Command29_Click()
   '
Dim ABUSQUE$(256)
   '
   MAILMESS.MAPISession1.DownLoadMail = False
   MAILMESS.MAPISession1.LogonUI = False
   MAILMESS.MAPISession1.Action = 1
   MAILMESS.MAPIMessages1.SessionID = MAILMESS.MAPISession1.SessionID
   '
   MAILMESS.MAPIMessages1.Fetch
   CAMENSA& = MAILMESS.MAPIMessages1.MsgCount
   '
   For U& = CAMENSA& To 1 Step -1
     '
10   MAILMESS.MAPIMessages1.MsgIndex = U& - 1
     ORENVI$ = TRIM$(MAILMESS.MAPIMessages1.MsgOrigDisplayName)
     MAILMESS.MAPIMessages1.Show
     'TT11$ = TRIM$(MAILMESS.MAPIMessages1.OLE)
     
     'TT12$ = TRIM$(MAILMESS.MAPIMessages1.MsgNoteText)
     MAILMESS.MAPIMessages1.Forward
     TT14$ = TRIM$(MAILMESS.MAPIMessages1.MsgSubject)
     MAILMESS.MAPIMessages1.MsgSubject = "Att: ESTO ES UNA PRUEBA"
       'MAILMESS.MAPIMessages1.MsgNoteText = " "
       'MAILMESS.MAPIMessages1.AttachmentIndex = 0
       'MAILMESS.MAPIMessages1.AttachmentPathName = "c:\PRUEBA.HTM"
       'MAILMESS.MAPIMessages1.AttachmentPosition = 1
       MAILMESS.MAPIMessages1.RecipType = 1
       'MAILMESS.MAPIMessages1.RecipDisplayName = MUESTRA$
       MAILMESS.MAPIMessages1.RecipAddress = "ebauer@flexoft.com.ar"
       MAILMESS.MAPIMessages1.MsgReceiptRequested = True
       '
       'On Error Resume Next
       MAILMESS.MAPIMessages1.Action = 3   ' vbMessageSend
  Stop
  Next U&
     MODOBO% = 0
     '
Stop
     For KKI% = 1 To CABUS%
       If Left$(UCase$(ORENVI$), Len(ABUSQUE$(KKI%))) = ABUSQUE$(KKI%) Then
         ATENTI% = 0
         MODOBO% = 0
         FORLIMPIA.Text1 = ORENVI$
         'FORLIMPIA.Text2 = Left$(MAPIMessages1.MsgNoteText, 1024)
         FORLIMPIA.Text2 = MAPIMessages1.MsgNoteText
         '
         ' busca por RFC:
         PINI% = 1
15       PPXX% = InStr(PINI%, UCase$(FORLIMPIA.Text2), "RFC822;")
         If PPXX% > 0 Then
           '
           FORLIMPIA.Text2.SelStart = PPXX% + 6
           PPYY% = InStr(PPXX%, FORLIMPIA.Text2, Chr$(13) + Chr$(10))
           If PPYY% <= PPXX% Then GoTo 18
           If PPYY% > PPXX% Then
             FORLIMPIA.Text2.SelLength = PPYY% - PPXX% - 7
           End If
           TEBUX$ = Mid$(FORLIMPIA.Text2, PPXX% + 7, FORLIMPIA.Text2.SelLength)
           If InStr(UCase$(TEBUX$), "FLEXOFT") > 0 Then
             PINI% = PPYY%
             GoTo 15
           End If
           If InStr(TEBUX$, "@") < 1 Then
             PINI% = PPYY%
             GoTo 15
           End If
           FORLIMPIA.Text3 = TEBUX$
           'MODOBO% = 1
           GoTo 50
         End If
         '
18       PINI% = 1
20       PPXX% = InStr(PINI%, FORLIMPIA.Text2, "<")
         If PPXX% > 0 Then
           FORLIMPIA.Text2.SelStart = PPXX%
           PPYY% = InStr(PPXX%, FORLIMPIA.Text2, ">")
           If PPYY% <= PPXX% Then GoTo 30
           If PPYY% > PPXX% Then
             FORLIMPIA.Text2.SelLength = PPYY% - PPXX% - 1
           End If
           TEBUX$ = Mid$(FORLIMPIA.Text2, PPXX% + 1, FORLIMPIA.Text2.SelLength)
           If InStr(UCase$(TEBUX$), "FLEXOFT") > 0 Then
             PINI% = PPYY%
             GoTo 20
           End If
           If InStr(TEBUX$, "@") < 1 Then
             PINI% = PPYY%
             GoTo 20
           End If
           FORLIMPIA.Text3 = TEBUX$
           'MODOBO% = 1
           GoTo 50
         End If
         '
30       PPYY% = InStr(FORLIMPIA.Text2, "; Failed")
         If PPYY% > 0 Then
           PPYY% = PPYY% - 1
           For PPXX% = PPYY% To 1 Step -1
             If Asc(Mid$(FORLIMPIA.Text2, PPXX%, 1)) < 45 Then
               FORLIMPIA.Text2.SelStart = PPXX%
               FORLIMPIA.Text2.SelLength = PPYY% - PPXX%
               TEBUX$ = Mid$(FORLIMPIA.Text2, PPXX% + 1, PPYY% - PPXX%)
               FORLIMPIA.Text3 = TEBUX$
               'MODOBO% = 1
               GoTo 50
             End If
           Next PPXX%
         End If
         '
50       'If MODOBO% = 1 Then
'           HOII% = HOY(HOS$)
'           If TRIM$(TEBUX$) <> "" Then
'             JJ& = 0
'             For UKK& = 1 To ULTREG&("LIDIRMAI")
'               XXX$ = REGLEIDO$("LIDIRMAI", UKK&)
'               XX1$ = TRIM$(Mid$(XXX$, 105, 60))
'               If XX1$ = TRIM$(TEBUX$) Then
'                 Call REPLA(XXX$, MKI$(HOII%), 255, 2)
'                 Call REPLA(XXX$, "MAL MAIL", 257, 8)
'                 Call GRAREG("LIDIRMAI", XXX$, UKK&)
'               End If
'             Next UKK&
'             ATENTI% = 1
'           End If
'           GoTo 60
'         End If
'         '
         FORLIMPIA.Check1.Value = 1
         FORLIMPIA.Check2.Value = 1
         FORLIMPIA.Text4 = "MAL MAIL"
         FORLIMPIA.Label7 = U&
         Screen.MousePointer = 1
         '
         If Option1.Value = True Then
           If InStr(FORLIMPIA.Text3, "@") > 0 Then
             Call FORLIMPIA.Command1_Click
             GoTo 60
           End If
         End If
         '
         FORLIMPIA.Show 1
         '
60       If ATENTI% = 1 Then
            MAPIMessages1.Delete
            If U& <= MAPIMessages1.MsgCount Then
              GoTo 10
            End If
          ElseIf ATENTI% = 2 Then
            GoTo 99
         End If
         GoTo 89
         '
       End If
     Next KKI%
     '
89 ' Next U&
   '
99 MAPISession1.Action = 2  ' !!!
   Screen.MousePointer = 1
End Sub

Private Sub Command31_Click()
    
10  On Error Resume Next
    AppActivate "Bandeja de Entrada"
    If Err < 1 Then
       On Error GoTo 0
       SendKeys "%{F4}", True
       GoTo 10
    End If
    On Error GoTo 0
    
    For U& = 1 To ULTREG&("ANOTASEG")
      XX$ = REGLEIDO$("ANOTASEG", U&)
      If CVS(Left$(XX$, 4)) = 1584 Then
         Call REPLA(XX$, MKS$(1604), 1, 4)
         Call GRAREG("ANOTASEG", XX$, U&)
      End If
    Next U&
    Call CIERRARCH("ANOTASEG")
End Sub

Private Sub Form_Load()
   Call CENTRAFORM(Me)
   Call AUTOEXEC
   USNBR% = 0
   ABC$ = UCase$(LUDAT$)
   If InStr(ABC$, " ") > 0 Then ' \ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB") > 0 Then
     LUDAT$ = "Y:\SEGUICOM\"
     MsgBox LUDAT$
   End If
   '
   FLEXLOGI.Show 1
   If USNBR% < 1 Then
      Close: End
   End If
   Label35 = NOMTER
   HOII% = HOY(HOS$)
   Label34 = HOS$
   Text12 = FECHATEX$(DIPOST(HOII%))
   ' Al establecer 0 en InputLen se indica a MSComm que
   ' lea todo el contenido del búfer de entrada cuando
   ' se use la propiedad Input.
   MSComm1.InputLen = 0
   'If Screen.Width < 12000 Then
   '   Hide
   '   FICONTAX.Show
   'End If
   '
End Sub

Private Sub Command1_Click()
   '
   If Val(Label26) < 1 Then
     Command16.SetFocus
     Exit Sub
   End If
   '
   Command1.Enabled = False
   If Val(Label26) > 0 Then
     Call GRAFICON
   End If
   Command1.Enabled = True
   '
End Sub

Private Sub Command10_Click()
   '
MsgBox "No Disponible"
Exit Sub
   '
   Command10.Enabled = False
   Call EXPORAGE
   Command10.Enabled = True
End Sub

Private Sub Command12_Click()
   '
   If USNBR% <> 101 Then Exit Sub
   Command12.Enabled = False
   IMPORCONTAC.Show 1
   'Call IMPORAGE
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   '
   Static FIDATX1#, FIDATX2#, UCONTA$
   '
   Command13.Enabled = False
   ABUSX$ = "BSEGCOM": If USNBR% < 100 Then ABUSX$ = "BSEGALE"
   If FIDATIM#(ABUSX$) = FIDAT1X# Then
     If FIDATIM#("!SECONTAC") = FIDATX2# Then
       JJ& = ULTREG&("!SECONTAC")
       GoTo 20
     End If
   End If
   '
   FIDATX1# = FIDATIM#(ABUSX$)
   JJ& = 0
   HOII% = HOY(HOS$)
   FINIPRO% = FECHANUM("30/04/06")  ' PARA EXCLUIR CONTACTOS VIEJOS
   For U& = 1 To ULTREG&(ABUSX$)
     XXX$ = REGLEIDO$(ABUSX$, U&)
     If CVI(Mid$(XXX$, 1023, 2)) <= HOII% Then
       If CVI(Mid$(XXX$, 1023, 2)) > FINIPRO% Then
         JJ& = JJ& + 1
         Call GRAREG("!SECONTAC", XXX$, JJ&)
       End If
     End If
   Next U&
   Call SETULTREG("!SECONTAC", JJ&)
   Call CIERRARCH("!SECONTAC")
   Call BAJALDISCO
   FIDATX2# = FIDATIM#("!SECONTAC")
   Call FRESHECHO("!SECONTAC")
   '
20 If JJ& > 0 Then
     FORSELEC.ABUSQUE = UCONTA$
     Call SELECHO("!SECONTAC")
     UCONTA$ = VALACT1$("!SECONTAC")
     NPX& = Val(UCONTA$)
     If NPX& > 0 Then
       Label26 = Str$(NPX&)
     End If
   End If
   '
   Command13.Enabled = True
   '
End Sub

Private Sub Command14_Click()
   If USNBR% <> 101 Then Exit Sub
   Command14.Enabled = False
   OPX% = COMALTER%("Opciones de Configuración:\\Conf.General\Particular")
   If OPX% = 1 Then
        Call CONECRUN("FLSETUP", "")
      ElseIf OPX% = 2 Then
        SETUPSGC.Show 1
   End If
   Command14.Enabled = True
End Sub

Private Sub Command21_Click()
   If USNBR% < 100 Then
      Call COMUNI("Para Registrar un Nuevo Contacto\Solicite por Mail la Correspondiente\Reserva de Cliente, Detallando:\   \- Nombre de la Empresa -\- Domicilio Completo -\- Persona de Contacto -\- Telefono / Fax / E-Mail -")
      Exit Sub
   End If
   '
   Label26 = 1 + ULTREG&("BSEGCOM")
   HOII% = HOY(HOS$)
   Label34 = HOS$
   Text12 = HOS$
   If USNBR% < 100 Then Text16 = TRIM$(Str$(USNBR%))
   On Error Resume Next
   Text4.SetFocus
   On Error GoTo 0
End Sub

Private Sub Command25_Click()
   ANCHAC = Width
   VISIAC = Picture1.Visible
   Image2.Visible = False
   Picture1.Visible = False
   Width = 10800
   On Error Resume Next
   PRINTFORM
   If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Error de Impresión de Ficha")
   End If
   On Error GoTo 0
   Image2.Visible = True
   Picture1.Visible = VISIAC
   Width = ANCHAC
End Sub

Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command4_Click()
   cancelflag = 1
End Sub

Private Sub Command5_Click()
   '
   Command5.Enabled = False
   HOII% = HOY(HOS$)
   '
   NPX& = Val(Label26)
   If NPX& > 0 Then
     If TRIM$(Text11) <> "" Then
       Call FRATEXTO(Text11, 56)
       For KKI% = 1 To CARETEX%
         TTXX$ = Space$(76)
         Call REPLA(TTXX$, RETEX$(KKI%), 11, 56)
         If KKI% = 1 Then
           Call REPLA(TTXX$, HOS$, 1, 8)
           Call REPLA(TTXX$, NOMTER$, 69, 24)
         End If
         Call REGAPP("ANOTASEG", Left$(MKS$(NPX&) + TTXX$, 96))
       Next KKI%
       Call REGAPP("ANOTASEG", MKS$(NPX&) + String$(88, "="))
       Call CIERRARCH("ANOTASEG")
       Text11 = ""
       '
       NPX1% = NPX&
       RFF$ = RECFILT$("BSEGCOM", 1, MKI$(NPX1%))
       If Len(RFF$) >= 4 Then
         NPY& = CVS(Left$(RFF$, 4))
         If NPY& <= ULTREG&("BSEGCOM") Then
           HOII% = HOY(HOS$)
           FEPROCON% = FECHANUM(Text12)
           If FEPROCON% < HOII% Then FEPROCON% = HOII%
           '
           XXX$ = REGLEIDO$("BSEGCOM", NPY&)
           If CVI(Left$(XXX$, 2)) = NPX1% Then
             Call REPLA(XXX$, MKI$(FEPROCON%), 1023, 2)
             Call GRAREG("BSEGCOM", XXX$, NPY&)
             Call CIERRARCH("BSEGCOM")
           End If
         End If
       End If
       '
       Call COPYSTRU("BSEGCOM", "BSEGALE")
       JJ& = 0
       FIN& = ULTREG&("BSEGCOM")
       For U& = 1 To FIN&
          XX$ = REGLEIDO$("BSEGCOM", U&)
          If CVI(Mid$(XX$, 971, 2)) = 3 Then
             JJ& = JJ& + 1
             Call GRAREG("BSEGALE", XX$, JJ&)
          End If
       Next U&
       Call SETULTREG("BSEGALE", JJ&)
       Call FRESHECHO("BSEGALE")
       '
     End If
   End If
   '
   Command5.Enabled = True
   '
End Sub

Private Sub Command6_Click()
   cancelflag = 1
End Sub

Private Sub Command7_Click()
   '
MsgBox "No Disponible"
Exit Sub
   '
   Command7.Enabled = False
   Call CIERRARCH("BSEGCOM")
   Call CIERRARCH("NOTASEG")
   Call CIERRARCH("*.*")
   '
   If TRIM$(UCase$(CurDir)) <> "C:\AGETEL" Then
     TTXX$ = "Su Agenda Telefónica FLEXOFT está Instalada"
     TTXX$ = TTXX$ + "\en la Carpeta " + REPLACAR$(CurDir, "\", "/") + "."
     TTXX$ = TTXX$ + "\  \Las Opciones de Respaldo y Recuperación"
     TTXX$ = TTXX$ + "\Pueden no Funcionar Correctamente en esta carpeta."
     TTXX$ = TTXX$ + "\  \Conviene Re-Instalar (Copiar) Toda la Agenda"
     TTXX$ = TTXX$ + "\a la Carpeta 'C:/AGETEL'."
     Call MENSERR(24, TTXX$)
   End If
   '
   If COMALTER%("Coloque Diskette de Respaldo\\Cancelar\Continuar") <> 2 Then
     GoTo 99
   End If
   '
   Screen.MousePointer = 11
   LDT$ = TRIM$(UCase$(LUDAT$))
   NX% = FILEOPEN%(LUCOM$ + "BKAGENTE.BAT", 2, 0)
   Print #NX%, LUBAS$ + "PKZIP " + "A:BCONTACT " + LDT$ + "BCONTACT.DAT " + LDT$ + "BCONTACT.X01 " + LDT$ + "NOTACONT.DAT " + LDT$ + "NOTACONT.X01 "
   Print #NX%, "EXIT"
   Close #NX%: NX% = 0
   Call BAJALDISCO
   '
   Shell LUCOM$ + "BKAGENTE.PIF", 2
   TINI = Timer
   While Timer < TINI + 5: Wend
   '
99 Command7.Enabled = True
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command8_Click()
   '
MsgBox "No Disponible"
Exit Sub
   '
   Command8.Enabled = False
   Call CIERRARCH("BCONTACT")
   Call CIERRARCH("*.*")
   '
   If TRIM$(UCase$(CurDir)) <> "C:\AGETEL" Then
     TTXX$ = "Su Agenda Telefónica FLEXOFT está Instalada"
     TTXX$ = TTXX$ + "\en la Carpeta " + REPLACAR$(CurDir, "\", "/") + "."
     TTXX$ = TTXX$ + "\  \Las Opciones de Respaldo y Recuperación"
     TTXX$ = TTXX$ + "\Pueden no Funcionar Correctamente en esta carpeta."
     TTXX$ = TTXX$ + "\  \Conviene Re-Instalar (Copiar) Toda la Agenda"
     TTXX$ = TTXX$ + "\a la Carpeta 'C:/AGETEL'."
     Call MENSERR(24, TTXX$)
   End If
   '
   If COMALTER%("Coloque Diskette de Respaldo\\Cancelar\Continuar") <> 2 Then
     GoTo 99
   End If
   '
   If EXISTE%("A:BCONTACT.ZIP") <> 1 Then
     Call COMUNI("Recuperación Cancelada")
     GoTo 99
   End If
   '
   Screen.MousePointer = 11
   LDT$ = TRIM$(UCase$(LUDAT$))
   NX% = FILEOPEN%(LUCOM$ + "BKAGENTE.BAT", 2, 0)
   Print #NX%, LUBAS$ + "PKUNZIP -O " + "A:BCONTACT " + LDT$
   Print #NX%, "EXIT"
   Close #NX%: NX% = 0
   Call BAJALDISCO
   '
   Shell LUCOM$ + "BKAGENTE.PIF", 2
   TINI = Timer
   While Timer < TINI + 5: Wend
   '
99 Command8.Enabled = True
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command9_Click()
   '
   If USNBR% <> 101 Then Exit Sub
   Command9.Enabled = False
   Call SORCONTAC
   Command9.Enabled = True
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Image4_Click()
   Image4.Enabled = False
   Call Dial(TRIM$(Text26.TEXT), Name)
   Image4.Enabled = True
End Sub

Private Sub Image5_Click()
   Image5.Enabled = False
   Call Dial(TRIM$(Text24.TEXT), Name)
   Image5.Enabled = True
End Sub

Private Sub Label16_DblClick()
   Call ACTAGRUP
End Sub

Private Sub Label26_Change()
   '
   Call CARFICHA
   '
End Sub

Private Sub Label37_DblClick()
   Label38 = ""
   Text14 = ""
End Sub

Private Sub List2_Click()
   List2.ToolTipText = TRIM$(Mid$(List2.TEXT, 25, 48))
End Sub

Private Sub List2_DblClick()
   Call ACTAGRUP
End Sub

Private Sub Text1_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text10_GotFocus(Index As Integer)
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text11_DblClick()
   Text11.Enabled = False
   Call Dial(TRIM$(Text11.TEXT), Name)
   Text11.Enabled = True
End Sub

Private Sub Text11_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text12_DblClick()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text12.TEXT = VDEV$
    End If
End Sub

Private Sub Text12_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text15_GotFocus()
   Text4.SetFocus
End Sub

Private Sub Text16_Change()
   Label40 = ""
   If Text16 = "1" Then
       Label40 = "Ernesto P. Bauer"
     ElseIf Text16 = "3" Then
       Label40 = "Adrian Da Encarnacao"
   End If
End Sub

Private Sub Text16_GotFocus()
   If USNBR% < 100 Then
      Text9.SetFocus
   End If
End Sub

Private Sub Text17_DblClick()
   Text17.Enabled = False
   Call Dial(TRIM$(Text17.TEXT), Name)
   Text17.Enabled = True
End Sub

Private Sub Text17_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text18_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text19_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text2_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text20_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text21_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text22_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text23_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text24_DblClick()
   Text24.Enabled = False
   Call Dial(TRIM$(Text24.TEXT), Name)
   Text24.Enabled = True
End Sub

Private Sub Text24_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text26_DblClick()
   Text26.Enabled = False
   Call Dial(TRIM$(Text26.TEXT), Name)
   Text26.Enabled = True
End Sub

'--------------------------------------------------------
'   DIALER.FRM
'   Copyright (c) 1994 Crescent Software, Inc.
'   por Carl Franklin
'
'   Actualizado por Anton de Jong
'
'   Muestra cómo marcar números telefónicos con un módem.
'
'   Para que este programa funcione, el teléfono y el
'   módem deben estar conectados a la misma línea.
'--------------------------------------------------------
'Option Explicit

'DefInt A-Z

' Este indicador se establece cuando el usuario elige Cancelar.

Sub Dial(NULLAMA$, FORNAME$)
   '
   Dim DialString$, FromModem$, dummy
    ' AT es el comando compatible Hayes ATTENTION, y es necesario para enviar comandos al módem.
    ' DT significa "marcar con tonos". El comando Dial utiliza tonos, en lugar de pulsos (DP =
    ' marcar con pulsos). Number$ es el número de teléfono a marcar.
    ' El punto y coma indica al módem que vuelva al modo de comando después de marcar
    ' (importante). Es necesario un retorno de carro, vbCr, al enviar comandos al módem.
    '
   '
   For KI% = 1 To Forms.Count
     If Forms(KI% - 1).Name = FORNAME$ Then GoTo 5
   Next KI%
   Exit Sub
   '
5  Number$ = ""
   INTERN$ = ""
   TOMALIN$ = TRIM$(CONTROL("AGETELEF", "TOMALIN"))
   '
   PPXX% = InStr(NULLAMA$, "/")
   If PPXX% < 1 Then PPXX% = 1 + Len(NULLAMA$)
   INTERN$ = TRIM$(Mid$(NULLAMA$, PPXX% + 1))
   NULLAMA$ = TRIM$(Left$(NULLAMA$, PPXX% - 1))
   '
   Number$ = TOMALIN$
   For KKI% = 1 To Len(NULLAMA$)
     CARA% = Asc(Mid$(NULLAMA$, KKI%, 1))
     If CARA% = 44 Or CARA% >= 48 Then
       If CARA% <= 57 Then
         Number$ = Number$ + Chr$(CARA%)
       End If
     End If
   Next KKI%
   If Number$ = "" Then Exit Sub
   '
10 Label28 = Number$
   Label29 = ""
   Picture1.Visible = True
   Picture1.Refresh
   cancelflag = 0
   '
   DialString$ = "ATDT" + Number$ + ";" + vbCr
    '
    ' Configuración del puerto de comunicaciones.
    ' Suponiendo que hay conectado un mouse a COM1, se establece 2 en CommPort
   SETUPSTRI$ = CONTROL("AGETELEF", "SETUPST")
   If SETUPSTRI$ = "" Then
     SETUPSTRI$ = "2     9600  N     8     1     "
   End If
   COPORT = Val(Left$(SETUPSTRI$, 6))
   On Error Resume Next
   MSComm1.CommPort = COPORT
   On Error GoTo 0
   '
   SETTI$ = ""
   For KKI% = 1 To 4
     SETTI$ = SETTI$ + TRIM$(Mid$(SETUPSTRI$, 6 * KKI% + 1, 6))
     If KKI% < 4 Then SETTI$ = SETTI$ + ","
   Next KKI%
   On Error Resume Next
   MSComm1.Settings = SETTI$
   If Err Then
      Call MENSERR(24, "Configuración de Puerto no Válida.\Corrija o Aplique Valores Predeterminados por Omisión.")
      Picture1.Visible = False
      SETUPTEL.Show 1
      Exit Sub
   End If
    '
    ' Abre el puerto de comunicaciones.
   On Error Resume Next
   MSComm1.PortOpen = True
   If Err Then
      Call MENSERR(24, "COM" + TRIM$(Str$(COPORT)) + " no está disponible.\Cambie la propiedad CommPort a otro puerto.")
      Picture1.Visible = False
      Exit Sub
   End If
    '
    ' Vuelca el búfer de entrada.
   MSComm1.InBufferCount = 0
    '
    ' Marca el número.
   MSComm1.Output = DialString$
   TFIN = Timer + 10
   Label30.Visible = False
    '
    ' Espera a que venga el "OK" del módem.
   Do
     dummy = DoEvents()
     CEVEN = MSComm1.CommEvent
     Select Case CEVEN
        Case comEvReceive
          Label29 = "Llamando ..."
          Label29.Refresh
          Label30.Visible = True
        Case comEvSend
          Label29 = "Marcando ..."
        Case Else
          'MsgBox "Error de Comunicacion"
          'exit do
     End Select
      '
      ' Si hay datos en el búfer, los lee.
     If MSComm1.InBufferCount Then
       FromModem$ = FromModem$ + MSComm1.Input
        ' Comprueba si hay "OK".
       If InStr(FromModem$, "OK") Then
             Label29 = "Llamando ..."
             Label29.Refresh
             Label30.Visible = True
             TFIN = Timer + 3
             While Timer < TFIN
               dummy = DoEvents()
             Wend
             Exit Do
          End If
       End If
        
       ' ¿El usuario ha elegido Cancelar?
       If cancelflag Then
          cancelflag = False
          Exit Do
       End If
       '
       If Timer > TFIN Then
         Picture1.Visible = False
         MsgBox "Imposible Establecer Comunicación"
         Exit Do
       End If
       '
    Loop
    '
    ' Desconecta el módem.
    MSComm1.Output = "ATH" + vbCr
    
    ' Cierra el puerto.
    MSComm1.PortOpen = False
    Picture1.Visible = False
    Call FRESHALL
    '
    If INTERN$ <> "" Then
      TFIN = Timer + 3
      While Timer < TFIN: Wend
      Number$ = INTERN$
      INTERN$ = ""
      VUELTA% = 2
      GoTo 10
    End If
    '
End Sub

Sub GRAFICON()
   '
5  IFICHA% = Val(Label26)
   REG& = REGBUS&("BSEGCOM", 1, MKI$(IFICHA%))
   If REG& < 1 Then REG& = 1 + ULTREG&("BSEGCOM")
   REGI% = REG&
   '
   If REG& <= ULTREG("BSEGCOM") Then
        RELE$ = REGLEIDO$("BSEGCOM", REG&)
      Else
        RELE$ = MKI$(REGI%) + Space$(1022)
   End If
   '
   Call REPLA(RELE$, AJUSTI$(Text4, 59) + AJUSTI$(Text15, 3), 3, 62)  ' empresa
   Call REPLA(RELE$, Text5, 65, 64)     ' DIRECCION ELECTRONICA
   Call REPLA(RELE$, Text20, 129, 48)   ' CALLE Y NUMERO
   Call REPLA(RELE$, Text21, 177, 8)    ' C.POSTAL
   Call REPLA(RELE$, Text22, 185, 32)   ' LOCALIDAD
   Call REPLA(RELE$, Text19, 217, 32)   ' PROVINCIA
   Call REPLA(RELE$, Text23, 249, 24)   ' PAIS
   '
   Call REPLA(RELE$, Text18, 273, 24)   ' ORIGEN
   Call REPLA(RELE$, Text1, 297, 24)    ' RAMO
   Call REPLA(RELE$, Text2, 321, 48)    ' PRINC.PRODUCTO
   Call REPLA(RELE$, Text3, 369, 8) ' provincia particular
   Call REPLA(RELE$, Text26, 377, 24)   ' TELEFONO
   Call REPLA(RELE$, Text17, 401, 24) ' telef.partic.
   Call REPLA(RELE$, Text24, 425, 24) ' telef2.partic
   '
   For KKI% = 0 To 2
     TTXX$ = Space$(96)
     Call REPLA(TTXX$, Text6(KKI%), 1, 32)    ' CONTACTO PERSONA 1 / 2 / 3
     Call REPLA(TTXX$, Text7(KKI%), 33, 24)
     Call REPLA(TTXX$, Text8(KKI%), 57, 24)
     Call REPLA(TTXX$, Text10(KKI%), 81, 16)
     Call REPLA(RELE$, TTXX$, 449 + 96 * KKI%, 96)
   Next KKI%
   '
   Call REPLA(RELE$, Text13, 737, 64) ' WEB-SITE
   
   PUNGRU$ = ""
   For U& = 1 To List2.ListCount
     PGR% = Val(Mid$(List2.List(U& - 1), 73, 6))
     If PGR% < 256 Then PUNGRU$ = PUNGRU$ + Chr$(PGR%)
   Next U&
   Call REPLA(RELE$, PUNGRU$ + String$(30, 0), 993, 30)
   '
   IVENDE% = XVALO(Text16)
   Call REPLA(RELE$, MKI$(IVENDE%), 971, 2)
   Call REPLA(RELE$, MKI$(FECHANUM(Label38)), 973, 2)   ' FECHA DE BAJA
   Call REPLA(RELE$, TRIM$(Text14), 975, 16)             ' MOTIVO DE BAJA
   Call REPLA(RELE$, MKI$(FECHANUM(Text12)), 1023, 2)   ' FECHA PROXIMO CONTACTO
   'HOII% = HOY(HOS$)
   'If CVI(Mid$(RELE$, 1023, 2)) < HOII% Then
   '  Call REPLA(RELE$, MKI$(HOII%), 1023, 2)
   'End If
   '
   Call GRAREG("BSEGCOM", RELE$, REG&)
   '
   Call CIERRARCH("BSEGCOM")
   Call FRESHECHO("BSEGCOM")
   '
   Call COPYSTRU("BSEGCOM", "BSEGALE")
   JJ& = 0
   FIN& = ULTREG&("BSEGCOM")
   For U& = 1 To FIN&
      XX$ = REGLEIDO$("BSEGCOM", U&)
      If CVI(Mid$(XX$, 971, 2)) = 3 Then
         JJ& = JJ& + 1
         Call GRAREG("BSEGALE", XX$, JJ&)
      End If
   Next U&
   Call SETULTREG("BSEGALE", JJ&)
   Call FRESHECHO("BSEGALE")
   '
   Screen.MousePointer = 11
   NCON% = Val(Label26)
   URN& = ULTREG&("NOTASEG")
   RNC$ = RECFILT$("NOTASEG", 1, MKI$(NCON%))
   KK& = (-3)
   '
   TTYZ$ = Text9.TEXT
   LOCADE& = Len(TTYZ$)
   For U& = LOCADE& To 1 Step -1
     If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 10
   Next U&
   U& = 0
   '
10 TEXTO$ = Left$(TTYZ$, U&)
   While Len(TEXTO$) > 0
      TTYY$ = MKI$(NCON%) + TEXTO$ + Space$(62)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG("NOTASEG", TTYY$, RENOTA&)
      TEXTO$ = Mid$(TEXTO$, 63)
   Wend
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("NOTASEG", MKI$(0) + Space$(62), RENOTA&)
   Wend
   '
   Call CIERRARCH("NOTASEG")
   Screen.MousePointer = 1

End Sub

Sub CARFICHA()
   '
5  IFICHA% = Val(Label26)
   REG& = REGBUS&("BSEGCOM", 1, MKI$(IFICHA%))
   If REG& < 1 Or REG& > ULTREG&("BSEGCOM") Then
       Text1 = ""
       Text2 = ""
       Text3 = ""
       Text4 = ""
       Text5 = ""
       Text9 = ""
       Text16 = ""
       Text15 = ""
       Text13 = ""
       Text12 = ""
       Text25 = ""
       Text14 = ""
       Text15 = ""
       Text11 = ""
       Text20 = ""
       Text21 = ""
       Text22 = ""
       Text19 = ""
       Text23 = ""
       Text26 = ""
       Text16 = "": Label40 = ""
       Text17 = ""
       Text24 = ""
       Text18 = ""
       '
       For KKI% = 0 To 2
         Text6(KKI%) = ""
         Text7(KKI%) = ""
         Text8(KKI%) = ""
         Text10(KKI%) = ""
       Next KKI%
       '
       HOII% = HOY(HOS$)
       Label34 = HOS$
       Text12 = FECHATEX$(DIPOST(HOII%))
       '
       List2.Clear
       '
     Else
       RELE$ = REGLEIDO$("BSEGCOM", REG&)
       Text4 = TRIM$(Mid$(RELE$, 3, 59))
       Text15 = TRIM$(Mid$(RELE$, 62, 3))
       Text5 = TRIM$(Mid$(RELE$, 65, 64))
       Text20 = TRIM$(Mid$(RELE$, 129, 48))
       Text21 = TRIM$(Mid$(RELE$, 177, 8))
       '
       Text22 = TRIM$(Mid$(RELE$, 185, 32))
       Text19 = TRIM$(Mid$(RELE$, 217, 32))
       Text23 = TRIM$(Mid$(RELE$, 249, 24))
       Text18 = TRIM$(Mid$(RELE$, 273, 24))
       Text1 = TRIM$(Mid$(RELE$, 297, 24))
       Text2 = TRIM$(Mid$(RELE$, 321, 48))
       Text3 = TRIM$(Mid$(RELE$, 369, 8))
       '
       Text26 = TRIM$(Mid$(RELE$, 377, 24))
       Text17 = TRIM$(Mid$(RELE$, 401, 24))
       Text24 = TRIM$(Mid$(RELE$, 425, 24))
       '
       For KKI% = 0 To 2
         TTXX$ = Mid$(RELE$, 449 + KKI% * 96, 96)
         Text6(KKI%) = Mid$(TTXX$, 1, 32)
         Text7(KKI%) = Mid$(TTXX$, 33, 24)
         Text8(KKI%) = Mid$(TTXX$, 57, 24)
         Text10(KKI%) = Mid$(TTXX$, 81, 16)
       Next KKI%
       '
       Text13 = TRIM$(Mid$(RELE$, 737, 64))
       '
       IVENDE% = CVI(Mid$(RELE$, 971, 2))
       If IVENDE% = 8224 Then IVENDE% = 0
       Text16 = TRIM$(FORMATNUM$(IVENDE%, "I4"))
       '
       List2.Clear
       PUNGRU$ = Mid$(RELE$, 993, 30)
       UTACON% = ULTREG&("TAGRUCON")
       For UKI% = 1 To Len(PUNGRU$)
         RETA& = Asc(Mid$(PUNGRU$, UKI%, 1))
         If RETA& > 0 Then
           If RETA& <= UTACON% Then
             XX$ = REGLEIDO$("TAGRUCON", RETA&)
             RAGRE% = Val(CVI(Left$(XX$, 2)))
             TAGRE$ = AJUSTI$(Mid$(XX$, 51, 24), 24) + AJUSTI$(Mid$(XX$, 3, 48), 48) + Str$(RAGRE%)
             List2.AddItem TAGRE$
           End If
         End If
       Next UKI%
       '
       FEBAJA% = CVI(Mid$(RELE$, 973, 2))                ' FECHA DE BAJA
       Label38 = ""
       If FEBAJA% > 8224 Then Label38 = FECHATEX$(FEBAJA%)
       Text14 = Mid$(RELE$, 975, 16)                      ' MOTIVO DE BAJA
       '
       Text12 = FECHATEX$(CVI(Mid$(RELE$, 1023, 2)))
       '
       Text9 = "": TTYZ$ = ""
       If IFICHA% > 0 Then
         Screen.MousePointer = 11
         RNC$ = RECFILT$("NOTASEG", 1, MKI$(IFICHA%))
         URN& = ULTREG&("NOTASEG")
         '
         For U& = 1 To Len(RNC$) Step 4
           RENOTA& = CVS(Mid$(RNC$, U&, 4))
           If RENOTA& > 0 Then
             If RENOTA& <= URN& Then
               TTYZ$ = TTYZ$ + Mid$(REGLEIDO$("NOTASEG", RENOTA&), 3)
             End If
           End If
         Next U&
         Screen.MousePointer = 1
       End If
       '
       LOCADE& = Len(TTYZ$)
       For U& = LOCADE& To 1 Step -1
         If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 30
       Next U&
       U& = 0
30     Text9 = Left$(TTYZ$, U&)
       Text9.Refresh
       '
   End If
   '
End Sub

Sub SORCONTAC()
   '
   OPXX% = COMALTER%("Elija si Criterio de Ordenamiento:\\Por Número\Alfabético\Cancelar")
   If OPXX% = 1 Then
       Call FILESORT("BSEGCOM", "", 1, 1, "ASC")
     ElseIf OPXX% = 2 Then
       Call FILESORT("BSEGCOM", "", 2, 2, "ASC")
   End If
   Call COPYSTRU("BSEGCOM", "BSEGALE")
   JJ& = 0
   FIN& = ULTREG&("BSEGCOM")
   For U& = 1 To FIN&
      XX$ = REGLEIDO$("BSEGCOM", U&)
      If CVI(Mid$(XX$, 971, 2)) = 3 Then
         JJ& = JJ& + 1
         Call GRAREG("BSEGALE", XX$, JJ&)
      End If
   Next U&
   Call SETULTREG("BSEGALE", JJ&)
   Call FRESHECHO("BSEGCOM")
   Call FRESHECHO("BSEGALE")
   Exit Sub
   '
     
   List1.Clear
   Screen.MousePointer = 11
   '
   For U& = 1 To ULTREG&("BSEGCOM")
     RELE$ = REGLEIDO$("BSEGCOM", U&)
     TX2$ = UCase$(TRIM$(Mid$(RELE$, 3, 62)))
     TTXX$ = AJUSTI$(TX2$, 64) + Str$(U&)
     List1.AddItem TTXX$
   Next U&
   '
   JJ& = 0
   Call SETULTREG("!BSEGCOM", 0)
   For U& = 1 To List1.ListCount
     TTXX$ = List1.List(U& - 1)
     REG& = Val(Mid$(TTXX$, 65))
     X$ = REGLEIDO$("BSEGCOM", REG&)
     JJ& = JJ& + 1
     Call GRAREG("!BSEGCOM", X$, JJ&)
   Next U&
   Call SETULTREG("!BSEGCOM", JJ&)
   Call CIERRARCH("*.*")
   '
   Call TRACENTRAL("BSEGCOM", "")
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
End Sub

Sub EXPORAGE()
   '
   On Error GoTo 90
   N1% = FreeFile
   Open "A:EXPORLIB.CSV" For Output As #N1%
   Print #N1%, "Nombre,Dirección electrónica"
   On Error GoTo 0
   '
   Screen.MousePointer = 11
   For U& = 1 To ULTREG&("BCONTACT")
     RELE$ = REGLEIDO$("BCONTACT", U&)
     TX2$ = TRIM$(Mid$(RELE$, 89, 56))
     TX4$ = TRIM$(Mid$(RELE$, 145, 64))
     TTXX$ = TRIM$(TX2$)
     If TRIM$(TX2$) <> "" And TRIM$(TX4$) <> "" Then
         TTXX$ = TTXX$ + " - "
       Else
         TTXX$ = TTXX$ + " "
     End If
     '
     TTYY$ = TRIM$(TTXX$ + TRIM$(TX4$))
     TTZZ$ = REPLACAR$(TTYY$, Chr$(34), "")
     TTXX$ = REPLACAR$(TTZZ$, ",", " ")
     '
     Call REPLA(RELE$, TTXX$, 3, 78)
     Call GRAREG("BCONTACT", RELE$, U&)
     '
     EMA$ = TRIM$(Mid$(RELE$, 217, 64))
     If EMA$ <> "" Then
       If Asc(Left$(EMA$, 1)) <> 39 Then
         AX$ = TTXX$ + "," + EMA$
         Print #N1%, AX$
         JJ& = JJ& + 1
       End If
     End If
   Next U&
   '
   Close #N1%
   Call CIERRARCH("*.*")
   MsgBox Str$(JJ&) + " Direcciones Exportadas"
   Screen.MousePointer = 1
   Exit Sub
   '
90 Resume 91
91 On Error GoTo 0
   '
End Sub

Sub IMPORAGE()
   '
   Screen.MousePointer = 11
   Dim RAEMP$(4096)
   For U& = 1 To ULTREG&("BCONTACT")
     X$ = REGLEIDO$("BCONTACT", U&)
     RAEMP$(U&) = UCase$(TRIM$(Mid$(X$, 145, 64)))
     CAEMP& = U&
   Next U&
   '
   FIN& = ULTREG&("!BCONTACT")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("!BCONTACT", U&)
     RASOE$ = UCase$(TRIM$(Mid$(X$, 145, 64)))
     EMAI$ = LCase$(TRIM$(Mid$(X$, 217, 64)))
     For JJ& = 1 To CAEMP&
       If RAEMP$(JJ&) = RASOE$ Then
         Y$ = REGLEIDO$("BCONTACT", JJ&)
         If InStr(Y$, EMAI$) > 0 Then GoTo 19 ' YA ESTA
         If TRIM$(Mid$(Y$, 217, 64)) = "" Then
             Call REPLA(Y$, EMAI$, 217, 64)
           ElseIf TRIM$(Mid$(Y$, 281, 64)) = "" Then
             Call REPLA(Y$, EMAI$, 281, 64)
           ElseIf TRIM$(Mid$(Y$, 345, 64)) = "" Then
             Call REPLA(Y$, EMAI$, 345, 64)
         End If
         Call GRAREG("BCONTACT", Y$, JJ&)
         GoTo 19
       End If
    Next JJ&
    '
    CAEMP& = CAEMP& + 1
    RAEMP$(CAEMP&) = UCase$(TRIM$(Mid$(X$, 145, 64)))
    KKI% = CAEMP&
    Call REPLA(X$, MKI$(KKI%), 1, 2)
    Call REGAPP("BCONTACT", X$)
19 Next U&
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
End Sub

Sub IMPOROUT()
   '
   On Error GoTo 90
   N1% = FreeFile
   Open "A:\OUTLOOK.CSV" For Input As #N1%
   Line Input #N1%, A$
   On Error GoTo 0
   '
   REG& = 0
   While Not EOF(N1%)
     REG& = REG& + 1
     REGI% = REG&
     RELE$ = MKI$(REGI%) + Space$(1022)
     '
     Line Input #N1%, TTXX$
     PPXX% = InStr(TTXX$, ",")
     If PPXX% > 0 Then
        EMPRE$ = Left$(TTXX$, PPXX% - 1)
        TTXX$ = Mid$(TTXX$, PPXX% + 1)
        PPXX% = InStr(TTXX$, ",")
        If PPXX% > 0 Then
          EMAI$ = Left$(TTXX$, PPXX% - 1)
          '
          Call REPLA(RELE$, EMPRE$, 3, 78)
          Call REPLA(RELE$, "Sr.", 81, 8)
          Call REPLA(RELE$, EMPRE$, 145, 64)
          Call REPLA(RELE$, EMAI$, 217, 64)
          Call GRAREG("BCONTACT", RELE$, REG&)
        End If
     End If
     '
   Wend
   '
   Call CIERRARCH("BCONTACT")
   Call FRESHECHO("BCONTACT")
   Exit Sub
   '
90 Resume 91
91 On Error GoTo 0
   '
End Sub

Private Sub Text26_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text3_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text4_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Sub Text5_DblClick()
   '
   NPX& = Val(Label26)
   If NPX& > 0 Then
     Call CREAMAIL(NPX&)
     '
     ATENTI% = 1
     MAILMESS.Label26 = Label26
     MAILMESS.Show 1
   End If
   '
End Sub

Private Sub Text5_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text6_GotFocus(Index As Integer)
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text7_GotFocus(Index As Integer)
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text8_GotFocus(Index As Integer)
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text9_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Sub ACTAGRUP()
   Call TRALOCAL("TAGRUCON", "")
   VTB% = VENTABU%("TAGRUCON")
   If VTB% >= 0 Then
     For U& = 1 To ULTREG&("!TAGRUCON")
       XX$ = REGLEIDO$("!TAGRUCON", U&)
       UKI% = U&
       Call REPLA(XX$, MKI$(UKI%), 1, 2)
       Call GRAREG("!TAGRUCON", XX$, U&)
     Next U&
     Call TRACENTRAL("TAGRUCON", "")
     If ULTREG&("TAGRUCON") > 255 Then
       Call SETULTREG("TAGRUCON", 255)
     End If
   End If
End Sub

Sub SEGUIOPER()
   Dim NCON%(), PROCON%(), PRIMOV&()
   ' 1.- SELECCION CONTACTOS MARCADOS
   FIN& = ULTREG&("BSEGCOM")
   CANSEL& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     RELE$ = REGLEIDO$("BSEGCOM", U&)
     If Mid$(RELE$, 62, 3) = "!!!" Then
       If TRIM$(Mid$(RELE$, 3, 59)) <> "" Then
         CANSEL& = CANSEL& + 1
         ReDim Preserve NCON%(CANSEL&), PROCON%(CANSEL&), PRIMOV&(CANSEL&)
         NCON%(CANSEL&) = CVI(Left$(RELE$, 2))
         PROCON%(CANSEL&) = CVI(Mid$(RELE$, 1023, 2))
         PRIMOV&(CANSEL&) = 0
       End If
     End If
   Next U&
   '
   ' 2.- ORDENAR POR FECHA PRÓXIMO CONTACTO
   For U1& = 1 To CANSEL&
     Call TRACE(20, U1&, CANSEL&)
     For U2& = U1& + 1 To CANSEL&
       If PROCON%(U2&) < PROCON%(U1&) Then
         NCONI% = NCON%(U2&)
         NCON%(U2&) = NCON%(U1&)
         NCON%(U1&) = NCONI%
         PROCONI% = PROCON%(U2&)
         PROCON%(U2&) = PROCON%(U1&)
         PROCON%(U1&) = PROCONI%
       End If
     Next U2&
   Next U1&
   '
   ' 3.- IDENTIFICAR ULTINA ANOTACIÓN
   For ISEL& = 1 To CANSEL&
     Call TRACE(20, ISEL&, CANSEL&)
     NPX& = NCON%(ISEL&)
     RFF$ = RECFILT$("ANOTASEG", 1, MKS$(NPX&))
     '
     For U& = Len(RFF$) To 1 Step -4
       RENOTA& = CVS(Mid$(RFF$, U& - 3, 4))
       XX$ = REGLEIDO$("ANOTASEG", RENOTA&)
       If TRIM$(Mid$(XX$, 5, 8)) <> "" Then
         If TRIM$(Mid$(XX$, 5, 8)) <> "========" Then
           PRIMOV&(ISEL&) = RENOTA&
           Exit For
         End If
       End If
     Next U&
   Next ISEL&
   '
   ' 4.- GENERAR LISTADO
   REBLA$ = REGBLAN$("SEGOPEND")
   JJ& = 0
   For ISEL& = 1 To CANSEL&
     Call TRACE(20, ISEL&, CANSEL&)
     NPX& = NCON%(ISEL&)
     YY$ = FILTERGETRECORD$("BSEGCOM", 1, MKI$(NCON%(ISEL&)))
     EMPRA$ = Mid$(YY$, 3, 59)
     NOMCON$ = Mid$(YY$, 449, 32)
     TELEF$ = Mid$(YY$, 377, 24)
     '
     RFF$ = RECFILT$("ANOTASEG", 1, MKS$(NPX&))
     KL& = 0
     '
     For U& = 1 To Len(RFF$) Step 4
       RENOTA& = CVS(Mid$(RFF$, U&, 4))
       If RENOTA& >= PRIMOV&(ISEL&) Then
         XX$ = REGLEIDO$("ANOTASEG", RENOTA&)
         ANOTA$ = Mid$(XX$, 5, 68)
         '
         ' REGISTROS DESCARTADOS
         ANOTB$ = TRIM$(ANOTA$)
         If ULTIA$ = "" Then
           If ANOTB$ = "" Then
             GoTo 29
           End If
         End If
         If Left$(ANOTB$, 10) = "==========" Then GoTo 29
         If ANOTB$ = "==" Then GoTo 29
         If ANOTB$ = "Viviana Sorrentino" Then GoTo 29
         If ANOTB$ = "Viviana" Then GoTo 29
         If ANOTB$ = "Sorrentino" Then GoTo 29
         If ANOTB$ = "Departamento" Then GoTo 29
         If ANOTB$ = "Comercial" Then GoTo 29
         If ANOTB$ = "FLEXOFT Soluciones Informáticas Integrales" Then GoTo 29
         If ANOTB$ = "Miembro de la Cámara de Empresas de Software y Servicios" Then GoTo 29
         If ANOTB$ = "Informáticos" Then GoTo 29
         If ANOTB$ = "Bauness 2170 - 5to 'D' - 1431 Buenos Aires - Argentina" Then GoTo 29
         If ANOTB$ = "Tel (00-54-11) 4524-1284 / 4541-1515 - e-mail:" Then GoTo 29
         If ANOTB$ = "info@flexoft.com.ar" Then GoTo 29
         If ANOTB$ = "Flavio Grasso." Then GoTo 29
         If ANOTB$ = "Miembro de la Cámara de Empresas de Software" Then GoTo 29
         If ANOTB$ = "y Servicios Informáticos" Then GoTo 29
         If ANOTB$ = "Tel (00-54-11) 4524-1284 / 4541-1515 / 4522-2048" Then GoTo 29
         If ANOTB$ = "e-mail: info@flexoft.com.ar" Then GoTo 29
         '
         ZZ$ = REBLA$
         If KL& = 0 Then
           Call REPLA(ZZ$, EMPRA$, 1, 32)
           Call REPLA(ZZ$, MKI$(PROCON%(ISEL&)), 137, 2)
           Call REPLA(ZZ$, NOMCON$, 139, 24)
           Call REPLA(ZZ$, TELEF$, 163, 24)
         End If
         Call REPLA(ZZ$, ANOTA$, 69, 68)
         KL& = KL& + 1
         JJ& = JJ& + 1
         Call GRAREG("SEGOPEND", ZZ$, JJ&)
         ULTIA$ = TRIM$(ANOTA$)
       End If
29   Next U&
     JJ& = JJ& + 1
     Call GRAREG("SEGOPEND", String$(256, "-"), JJ&)
   Next ISEL&
   Call SETULTREG("SEGOPEND", JJ&)
   Call CIERRARCH("SEGOPEND")
   Call FINAL("*SEGOPEN")
   '
End Sub
