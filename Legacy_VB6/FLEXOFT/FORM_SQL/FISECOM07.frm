VERSION 5.00
Begin VB.Form FISECOM07 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguimiento Comercial FLEXOFT"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FISECOM07.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11880
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8355
      Left            =   10850
      ScaleHeight     =   8325
      ScaleWidth      =   1065
      TabIndex        =   104
      Top             =   0
      Width           =   1095
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
         Height          =   600
         Left            =   160
         Picture         =   "FISECOM07.frx":1272
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "Listados por Impresora"
         Top             =   4440
         Width           =   690
      End
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   180
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   119
         Top             =   7560
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   120
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton AGenRep 
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
         Left            =   160
         Picture         =   "FISECOM07.frx":157C
         Style           =   1  'Graphical
         TabIndex        =   118
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   3720
         Width           =   690
      End
      Begin VB.CommandButton Command32 
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
         Left            =   160
         Picture         =   "FISECOM07.frx":1886
         Style           =   1  'Graphical
         TabIndex        =   116
         ToolTipText     =   "Sub-Sistema de Presupuestos de Clientes"
         Top             =   3120
         Visible         =   0   'False
         Width           =   690
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
         Left            =   160
         Picture         =   "FISECOM07.frx":1CC8
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   1320
         Width           =   690
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00624E28&
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
         ForeColor       =   &H8000000E&
         Height          =   975
         Left            =   10
         TabIndex        =   110
         Top             =   6120
         Visible         =   0   'False
         Width           =   990
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
            Left            =   160
            Picture         =   "FISECOM07.frx":1FD2
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Configuración de Funcionamiento"
            Top             =   1580
            Visible         =   0   'False
            Width           =   690
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
            Height          =   600
            Left            =   160
            Picture         =   "FISECOM07.frx":2414
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Imprime Ficha del Contacto"
            Top             =   315
            Width           =   690
         End
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
            Left            =   160
            Picture         =   "FISECOM07.frx":2A7E
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Administrar Mail Circular"
            Top             =   2205
            Visible         =   0   'False
            Width           =   690
         End
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
         Left            =   160
         Picture         =   "FISECOM07.frx":2D88
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Ordenamiento Alfabético"
         Top             =   2520
         Width           =   690
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
         Left            =   160
         Picture         =   "FISECOM07.frx":3092
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   120
         Width           =   690
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
         Left            =   160
         Picture         =   "FISECOM07.frx":31DC
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Alta de Nueva Ficha de Seguimiento"
         Top             =   720
         Width           =   690
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
         Left            =   160
         Picture         =   "FISECOM07.frx":3326
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Elimina Contacto Activo de la Base"
         Top             =   1920
         Width           =   690
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00624E28&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   1065
         Left            =   15
         TabIndex        =   105
         Top             =   3255
         Visible         =   0   'False
         Width           =   990
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
            TabIndex        =   109
            ToolTipText     =   "Confecciona Copia de Seguridad (Zip)"
            Top             =   315
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
            TabIndex        =   108
            ToolTipText     =   "Restaura Copia de Seguridad (Zip)"
            Top             =   315
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
            TabIndex        =   107
            ToolTipText     =   "Exporta Direcciones para Outlook Express"
            Top             =   650
            Width           =   330
         End
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
            TabIndex        =   106
            ToolTipText     =   "Importa Libreta de Direcciones (de C:\FLEXOFT)"
            Top             =   650
            Width           =   330
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2D3C0&
      Caption         =   "OTROS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1950
      Left            =   9660
      TabIndex        =   86
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
         Picture         =   "FISECOM07.frx":3630
         Style           =   1  'Graphical
         TabIndex        =   88
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
         Left            =   120
         Picture         =   "FISECOM07.frx":393A
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Abrir Anotador de Novedades"
         Top             =   315
         Width           =   850
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2D3C0&
      Caption         =   "IDENTIFICACION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4425
      Left            =   105
      TabIndex        =   43
      Top             =   115
      Width           =   10620
      Begin VB.CommandButton Command13 
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9790
         TabIndex        =   114
         ToolTipText     =   "Selección Contactos Programados"
         Top             =   390
         Width           =   160
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   6.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         Left            =   10020
         Picture         =   "FISECOM07.frx":3C44
         Style           =   1  'Graphical
         TabIndex        =   113
         ToolTipText     =   "Búsqueda Secuencial por Texto"
         Top             =   390
         Width           =   480
      End
      Begin VB.CommandButton Command35 
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
         Left            =   1680
         TabIndex        =   7
         ToolTipText     =   "Selecciona Contacto - Lista Completa"
         Top             =   2160
         Width           =   175
      End
      Begin VB.TextBox Text27 
         Appearance      =   0  'Flat
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
         Left            =   1155
         TabIndex        =   6
         Top             =   2160
         Width           =   555
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E2D3C0&
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
         Left            =   5320
         TabIndex        =   102
         Top             =   2120
         Width           =   1720
         Begin VB.CommandButton Command31 
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
            Height          =   235
            Left            =   1465
            TabIndex        =   11
            ToolTipText     =   "Selecciona Contacto - Lista Completa"
            Top             =   140
            Width           =   150
         End
         Begin VB.Label Label42 
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
            Left            =   1075
            TabIndex        =   117
            Top             =   150
            Width           =   405
         End
         Begin VB.Label LABEL40 
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
            Left            =   125
            TabIndex        =   103
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
         Picture         =   "FISECOM07.frx":3F4E
         Style           =   1  'Graphical
         TabIndex        =   89
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
         TabIndex        =   101
         Text            =   " "
         Top             =   420
         Width           =   330
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
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
         Left            =   9345
         TabIndex        =   100
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
         TabIndex        =   96
         Top             =   1050
         Width           =   150
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
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
         Left            =   1155
         TabIndex        =   2
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
         TabIndex        =   92
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
         TabIndex        =   91
         Top             =   4070
         Width           =   790
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   8820
         Sorted          =   -1  'True
         TabIndex        =   18
         ToolTipText     =   "Doble Click para Actualizar Tabla"
         Top             =   3360
         Width           =   1590
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
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
         TabIndex        =   17
         Top             =   3360
         Width           =   7260
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
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
         Left            =   9660
         TabIndex        =   16
         Text            =   " "
         Top             =   2940
         Width           =   750
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   1155
         TabIndex        =   9
         Text            =   " "
         Top             =   2940
         Width           =   2115
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
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
         Left            =   4830
         TabIndex        =   10
         Text            =   " "
         Top             =   2940
         Width           =   3585
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
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
         Left            =   1155
         TabIndex        =   1
         Text            =   " "
         ToolTipText     =   "Doble-Click para Enviar un Mail"
         Top             =   720
         Width           =   6630
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Nro. Prospecto / Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   850
         Left            =   7875
         TabIndex        =   62
         Top             =   150
         Width           =   2720
         Begin VB.CommandButton COMMAND16 
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
            Left            =   1400
            TabIndex        =   112
            Top             =   240
            Width           =   150
         End
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
            Left            =   1735
            Picture         =   "FISECOM07.frx":4390
            TabIndex        =   94
            ToolTipText     =   "Selecciona Contacto - Lista Completa"
            Top             =   240
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
            Left            =   1570
            Picture         =   "FISECOM07.frx":469A
            TabIndex        =   93
            ToolTipText     =   "Selecciona Contacto - Lista Completa"
            Top             =   240
            Width           =   160
         End
         Begin VB.Label Label41 
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
            Left            =   75
            TabIndex        =   115
            Top             =   525
            Width           =   1350
         End
         Begin VB.Label Label26 
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
            Left            =   75
            TabIndex        =   63
            Top             =   240
            Width           =   1350
         End
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
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
         Left            =   1155
         TabIndex        =   0
         Text            =   " "
         Top             =   420
         Width           =   6420
      End
      Begin VB.TextBox Text26 
         Appearance      =   0  'Flat
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
         Left            =   7875
         TabIndex        =   13
         Text            =   " "
         Top             =   1890
         Width           =   2150
      End
      Begin VB.TextBox Text24 
         Appearance      =   0  'Flat
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
         Left            =   7875
         TabIndex        =   15
         Text            =   " "
         Top             =   2520
         Width           =   2150
      End
      Begin VB.TextBox Text23 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   8
         Text            =   " "
         Top             =   2520
         Width           =   4005
      End
      Begin VB.TextBox Text22 
         Appearance      =   0  'Flat
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
         Left            =   3675
         TabIndex        =   5
         Text            =   " "
         Top             =   1785
         Width           =   3375
      End
      Begin VB.TextBox Text21 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   4
         Top             =   1800
         Width           =   1485
      End
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
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
         Left            =   1155
         TabIndex        =   3
         Text            =   " "
         Top             =   1470
         Width           =   5895
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
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
         Left            =   7875
         TabIndex        =   12
         Text            =   " "
         Top             =   1470
         Width           =   2535
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
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
         Left            =   7875
         TabIndex        =   14
         Text            =   " "
         Top             =   2205
         Width           =   2535
      End
      Begin VB.Label LABEL52 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   111
         Top             =   2160
         Width           =   3225
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9135
         TabIndex        =   99
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label Label38 
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
         Height          =   285
         Left            =   8250
         TabIndex        =   98
         Top             =   1050
         Width           =   680
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Baja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7875
         TabIndex        =   97
         ToolTipText     =   "Doble-Click para Re-Activar"
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Web-Site"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   95
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7820
         TabIndex        =   90
         ToolTipText     =   "Doble Click para Actualizar Tabla"
         Top             =   3360
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   68
         Top             =   3360
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Empleados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8505
         TabIndex        =   67
         Top             =   3045
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Profesión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   66
         Top             =   3045
         Width           =   960
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Especialidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2940
         TabIndex        =   65
         Top             =   3045
         Width           =   1800
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   64
         Top             =   840
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   61
         Top             =   525
         Width           =   960
      End
      Begin VB.Image Image5 
         Height          =   300
         Left            =   10080
         Picture         =   "FISECOM07.frx":49A4
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   315
      End
      Begin VB.Image Image4 
         Height          =   300
         Left            =   10080
         Picture         =   "FISECOM07.frx":4CAE
         Stretch         =   -1  'True
         Top             =   1890
         Width           =   315
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Calle y Nº"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   51
         Top             =   1575
         Width           =   960
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2625
         TabIndex        =   50
         Top             =   1890
         Width           =   960
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "País"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   49
         Top             =   2600
         Width           =   960
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Telef"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6825
         TabIndex        =   48
         Top             =   1995
         Width           =   960
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C. Postal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   47
         Top             =   1890
         Width           =   960
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6825
         TabIndex        =   46
         Top             =   2625
         Width           =   960
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   45
         Top             =   2260
         Width           =   960
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Origen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6825
         TabIndex        =   44
         Top             =   1575
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "NOVEDADES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1950
      Left            =   105
      TabIndex        =   74
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
         TabIndex        =   84
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
         TabIndex        =   85
         Top             =   1470
         Width           =   1000
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   425
         Left            =   0
         ScaleHeight     =   390
         ScaleWidth      =   9390
         TabIndex        =   75
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
            TabIndex        =   77
            Top             =   75
            Width           =   9360
         End
         Begin VB.Label Label13 
            Caption         =   "Label5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   15
            Left            =   0
            TabIndex        =   76
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
         Picture         =   "FISECOM07.frx":4FB8
         Style           =   1  'Graphical
         TabIndex        =   40
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
         TabIndex        =   83
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
         TabIndex        =   82
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
         TabIndex        =   81
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
         TabIndex        =   80
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
         TabIndex        =   79
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
         TabIndex        =   78
         Top             =   840
         Width           =   4000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "CONTACTOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   105
      TabIndex        =   69
      Top             =   4560
      Width           =   10620
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   9135
         TabIndex        =   30
         Top             =   1260
         Width           =   1275
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   7140
         TabIndex        =   29
         Top             =   1260
         Width           =   1800
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   4830
         TabIndex        =   28
         Top             =   1260
         Width           =   2115
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   1155
         TabIndex        =   27
         Top             =   1260
         Width           =   3480
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
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
         Index           =   1
         Left            =   9135
         TabIndex        =   26
         Top             =   945
         Width           =   1275
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
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
         Index           =   1
         Left            =   7140
         TabIndex        =   25
         Top             =   945
         Width           =   1800
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
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
         Index           =   1
         Left            =   4830
         TabIndex        =   24
         Top             =   945
         Width           =   2115
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
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
         Index           =   1
         Left            =   1155
         TabIndex        =   23
         Top             =   945
         Width           =   3480
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
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
         Index           =   0
         Left            =   9135
         TabIndex        =   22
         Top             =   600
         Width           =   1275
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
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
         Index           =   0
         Left            =   7140
         TabIndex        =   21
         Top             =   630
         Width           =   1800
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
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
         Index           =   0
         Left            =   4830
         TabIndex        =   20
         Top             =   630
         Width           =   2115
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
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
         Index           =   0
         Left            =   1155
         TabIndex        =   19
         Top             =   630
         Width           =   3480
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Directo / Interno"
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
         Left            =   8085
         TabIndex        =   73
         Top             =   350
         Width           =   2010
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sector"
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
         Left            =   6510
         TabIndex        =   72
         Top             =   350
         Width           =   960
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo / Función"
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
         Left            =   4305
         TabIndex        =   71
         Top             =   350
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Persona de Contacto"
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
         Left            =   210
         TabIndex        =   70
         Top             =   350
         Width           =   1905
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   10680
      Sorted          =   -1  'True
      TabIndex        =   60
      Top             =   735
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00D0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   2835
      ScaleHeight     =   2475
      ScaleWidth      =   4155
      TabIndex        =   52
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
         TabIndex        =   59
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
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
         Top             =   0
         Width           =   4110
      End
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Apel:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   42
      Top             =   1785
      Width           =   960
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Apel:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   41
      Top             =   1260
      Width           =   960
   End
End
Attribute VB_Name = "FISECOM07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cancelflag
Dim MODIFNOMBRE As Integer

Sub ActuFicha(Optional OPCI$)
    '
    If TRIM$(UCase$(OPCI$)) = "CODIGO" Then
        CONDI$ = "CODI_CLI <> '' ORDER BY CODI_CLI ASC"
      Else
        CONDI$ = "CODI_CLI <> '' ORDER BY RASO_CLI ASC"
    End If
    '
    Call CARGALISEL("INDIPROS", "BSEGCOM", "CODI_CLI/A12;RASO_CLI/A50", CONDI$, "NOMESS")

End Sub

Private Sub AGenRep_Click()
    '
    'se comento por que daba un error 5 en el procedimiento
    ' text4.gostfocus ***Queda para ver***
    AGenRep.Enabled = False '
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Command11_Click()
   '\\\OT-05-0966-OD-27/10/08///
   If TRIM$(Label26) = "" Then
     COMMAND16.SetFocus
     Exit Sub
   End If
   '
   'VALIDACION DE QUE LA FICHA ESTA GRABADA
   COCLI$ = TRIM$(Label26)
   NUFICH& = XVALO(Mid$((VALOBADA("BSEGCOM", "CODI_CLI", "CODI_CLI ='" & COCLI$ & "'")), 4))
   If NUFICH& < 1 Then COMMAND16.SetFocus: Exit Sub
   '
   Command11.Enabled = False
   If TRIM$(Label26) <> "" Then
     ANOTASEG07.Label26 = Label26
     ANOTASEG07.Label3 = Text4
     ANOTASEG07.Label7 = Text6(0)
     ANOTASEG07.Label8 = Text6(1)
     ANOTASEG07.Label11 = Text1
     ANOTASEG07.Label13 = Text2
     ANOTASEG07.Label14 = TRIM$(Text3)
     ANOTASEG07.Text12 = Text12
     '
     ANOTASEG07.List1.Clear
     Call ABRECLIEN
     ANOTACION$ = VALOBADA("BSEGCOM", "HIS_CONTAC", "CODI_CLI = '" & COCLI$ & "'")
     XX$ = ANOTACION$: LONGI& = Len(XX$)
     INICIO& = 1
     '
     Do While INICIO& < LONGI&
        FIN& = InStr(INICIO&, XX$, Chr$(10))
        If FIN& > 0 Then
             TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO&)
             ANOTASEG07.List1.AddItem TTXX$
             INICIO& = FIN& + 2
          Else
             TTXX$ = Mid$(XX$, INICIO&)
             ANOTASEG07.List1.AddItem TTXX$
             Exit Do
        End If
     Loop
     '
     TPI = ANOTASEG07.List1.ListCount - 12
     If TPI < 0 Then TPI = 0
     ANOTASEG07.List1.TopIndex = TPI
     '
     ANOTASEG07.Show 1
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
   ABUSX$ = "INDIPROS" ' : If USNBR% < 100 Then ABUSX$ = "BSEGALE"
   Call SELECHO(ABUSX$)
   VDEVU$ = VALACT1$("INDIPROS")
   If VDEVU$ <> "" Then
     Label26 = TRIM$(VDEVU$)
   End If
End Sub

Private Sub Command17_Click()
   If Val(Label26) < 1 Then
     COMMAND16.SetFocus
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
   
   OPX% = ALTERNA%("Base de Datos Completa\Alarma de Prospect")
   Select Case OPX%
     Case 1
        Screen.MousePointer = 11
        Call ActualizaArchivo(OPX%)
        Screen.MousePointer = 1

        Call FINAL("*BSEGCOM7")
     Case 2
       Call ALARMAPROSPECT
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
     SQLX$ = "SELECT * FROM BSEGCOM WITH(NOLOCK) "
     Dim BSEGTMP As ADODB.Recordset
     Set BSEGTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With BSEGTMP
       Do While Not .EOF
          For KKI% = 1 To .Fields.Count
              On Error Resume Next
              VATEX$ = ""
              VATEX$ = .Fields(KKI% - 1).Value
              On Error GoTo 0
              If InStr(VATEX$, ARGUME$) > 0 Then
                Call REPLA(EL1$, !Codi_Cli, 1, 12)
                Call REPLA(EL1$, AJUSTI$(SAFETEXT$(!raso_cli), 50), 13, 50)   ' empresa
                JJ& = JJ& + 1
                Call GRAREG("!SECONTA7", EL1$, JJ&)
                Exit For
              End If
          Next KKI%
       .MoveNext
       Loop
     End With
     '
     Call SETULTREG("!SECONTA7", JJ&)
     Call CIERRARCH("!SECONTA7")
     Call BAJALDISCO
     Call FRESHECHO("!SECONTA7")
     '
20   If JJ& > 0 Then
       Call SELECHO("!SECONTA7")
       UCONTA$ = VALACT1$("!SECONTA7")
     
        If UCONTA$ <> "" Then
          Label26 = UCONTA$
        End If
     End If
     '
   End If
   Screen.MousePointer = 1
   Command2.Enabled = True
End Sub

Private Sub Command20_Click()
   '
   Exit Sub
   '
   If USNBR% <> 101 Then Exit Sub
   Command20.Enabled = False
   ARMAICIRC.Show 1
   '
99 Command20.Enabled = True
   '
End Sub

Private Sub Command22_Click()
   If List2.ListIndex >= 0 Then
     List2.RemoveItem (List2.ListIndex)
   End If
End Sub

Private Sub Command23_Click()
   If Val(Label26) < 1 Then
     COMMAND16.SetFocus
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
   Command24.Enabled = False
   If COMALTER%("Realmente Desea Eliminar PERMANENTEMENTE\la Ficha Activa ?\\No, Cancelar\Si, Eliminar") = 2 Then
5    Screen.MousePointer = 11
     COCLI$ = TRIM$(Label26)
     '
     'borra de la tabla
     Call ABRECLIEN
     Call BORRAREGISTROS("BSEGCOM", "CODI_CLI = '" & COCLI$ & "'", REGAF&, "NOMESS")
     Call ActuFicha
     Label26 = ""
     MODOEDICION = 0
     '
   End If
   '
   Screen.MousePointer = 1
   Command24.Enabled = True
   '
End Sub

Private Sub Command26_Click()
   ABUSX$ = "INDIPROS" ': If USNBR% < 100 Then ABUSX$ = "BSEGALE"
   UKU& = REGNUM&(ABUSX$)
   If UKU& > 1 Then
     XX$ = REGLEIDO$(ABUSX$, UKU& - 1)
     Label26 = TRIM$(Left$(XX$, 12))
   End If
End Sub

Private Sub Command27_Click()
   ABUSX$ = "INDIPROS" ': If USNBR% < 100 Then ABUSX$ = "BSEGALE"
   UKU& = REGNUM&(ABUSX$)
   If UKU& > 0 Then
     XX$ = REGLEIDO$(ABUSX$, UKU& + 1)
     Label26 = TRIM$(Left$(XX$, 12))
   End If

End Sub

Private Sub Command28_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Label38 = VDEV$
    End If
End Sub

Private Sub Command31_Click()

  If DERACCE%("MODIVEN", "Modificación de Vendedor") > 1 Then ' derecho de acceso para modificar vendedor
   '
   Call SELECHO("VENDEDOR")
   VDEVU$ = VALACT1$("VENDEDOR")
   If TRIM$(VDEVU$) <> "" Then
     Label42 = VDEVU$
   End If
   '
  End If
End Sub

Private Sub Command32_Click()
    '
    CONEXI$ = "PRESUP07" ' CONECTA CON PRESUPUESTO Y LE PASA EL NUMERO DE PROSEPCTO
    If TRIM$(Label26) <> "" Then CONEXI$ = CONEXI$ + "/" + TRIM$(Label26)
    Call CONECRUN(CONEXI$, "Presupuestos de Clientes")
    '
End Sub

Sub ALARMAPROSPECT()
   '
   Call SELECHO("VENDEDOR")
   VEND% = XVALO(VALACT1$("VENDEDOR"))
   If VEND% < 1 Then
     Exit Sub
   End If
   '
   HOII% = HOY(HOS$)
   FECHOY = FETEXCOR1(HOII%)
   '
   SQLX$ = "SELECT * FROM BSEGCOM WITH(NOLOCK) WHERE "
   CONDI$ = " VEND_CLI =" & VEND%
   CONDI$ = CONDI$ + " AND FECHPROXCTO <= '" & FECHOY & " '"
   CONDI$ = CONDI$ + " AND (FECHBAJA >= '" & FECHOY & " ' OR FECHBAJA IS NULL OR FECHBAJA = '')"
   FIN& = CantRegistros("BSEGCOM", CONDI$)
   If FIN& < 1 Then
      Call COMUNI("No Se Encontraron Contactos Vencidos Para Este Vendedor")
      Exit Sub
   End If
   SQLX$ = SQLX$ + CONDI$
   '
   JJ& = 0
   JX& = 0
   TX$ = REGBLAN$("ALARPROS")
   Call SETULTREG("ALARPROS", 0)
   Dim BSEGTMP As ADODB.Recordset
   Set BSEGTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With BSEGTMP
    Do While Not .EOF
        Call TRACE(24, JX&, FIN&)
        JX& = JX& + 1
        COCLI$ = SAFETEXT$(!Codi_Cli)
        CONDI$ = "CODI_CLI =  '" & COCLI$ & "'"
        NC% = XVALO(VALOBADA("DEUDOR12", "NUME_CLI", CONDI$, "NOMESS"))
        If NC% > 0 Then GoTo 50 ' VALIDACION QUE NO ESTE DADO DE ALTA COMO CLIENTE
        Call REPLA(TX$, COCLI$, 1, 10)
        Call REPLA(TX$, SAFETEXT$(!raso_cli), 13, 64)
        TELE1$ = SAFETEXT$(!Tele_Cli)
        TELE2$ = SAFETEXT$(!TELE_Cli2)
        TELE$ = ""
        If TRIM$(TELE1$) <> "" Then TELE$ = TELE1$
        If TRIM$(TELE2$) <> "" Then TELE$ = TELE$ + "/" + TELE2$
        Call REPLA(TX$, TELE$, 77, 24)
          FEPROCTO% = CVINT(!FechProxCto)
        Call REPLA(TX$, MKI$(FEPROCTO%), 101, 2)
        notax$ = SAFETEXT$(!HIS_CONTAC)
        i& = 0
        CONTA% = 0
        While InStr(notax$, Chr$(10)) > 0
           CONTA% = CONTA% + 1
           POS2% = InStr(notax$, Chr$(10))
           NOTAX1$ = Mid$(notax$, 1, POS2% - 1)
           NOTAX1$ = REPLACAR$(NOTAX1$, Chr$(61), "") 'QUITA EL =
           NOTAX1$ = REPLACAR$(NOTAX1$, Chr$(45), "") ' QUITA EL -
           Call REPLA(TX$, SAFETEXT$(NOTAX1$), 103, 128)
           
           JJ& = JJ& + 1
           Call GRAREG("ALARPROS", TX$, JJ&)
           TX$ = REGBLAN$("ALARPROS")
           notax$ = Mid$(notax$, POS2% + 2)
        Wend
        If CONTA% < 1 Then
           JJ& = JJ& + 1
           Call REPLA(TX$, SAFETEXT$(NOTAX1$), 103, 128)
           Call GRAREG("ALARPROS", TX$, JJ&)
           TX$ = REGBLAN$("ALARPROS")
        End If
        
50  .MoveNext
    Loop
   End With
   BSEGTMP.Close
   Set BSEGTMP = Nothing
   If JJ& < 1 Then
      Call COMUNI("No Se Encontraron Contactos Vencidos Para Este Vendedor")
      Exit Sub
   End If

   Call SETULTREG("ALARPROS", JJ&)
   Call FINAL("*ALARPROS")

End Sub

Private Sub COMMAND35_Click()

   If CantRegistros("PROVINCIA", "ID > 0 ", "NOMESS") > 0 Then
       If CantRegistros("LOCALIDAD", "ID > 0 ", "NOMESS") > 0 Then
            CONDI$ = "CODFLEX = '" & Text27 & "'"
            PROVLOCA.TXTPROVNUM = VALOBADA("PROVINCIA", "ID", CONDI$, "NOMESS")
            PROVLOCA.TXTLOCALNUM = ""
            PROVLOCA.TXTNAMEPROVINCIA = LABEL52
            PROVLOCA.TXTPROVLET = Text27
            PROVLOCA.TXTLOCALIDAD = Text22
            PROVLOCA.TXTCODPOSTAL = Text21
            
            PROVLOCA.Show 1
            If PROVLOCA.APROBO > 0 Then
                LABEL52 = PROVLOCA.TXTNAMEPROVINCIA
                Text27 = PROVLOCA.TXTPROVLET
                Text22 = PROVLOCA.TXTLOCALIDAD
                Text21 = PROVLOCA.TXTCODPOSTAL
            End If
            Exit Sub
       End If
   End If

   Call SELECHO("PROVINC")
   VDEVU$ = VALACT1$("PROVINC")
   If TRIM$(VDEVU$) <> "" Then
      Text27 = VDEVU$
   End If

End Sub

Private Sub Form_Load()
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   Call CENTRAFORM(Me)
   Call TRANSLABELS(Name)
   DoEvents
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
     Caption = Caption + "  -  " + EPAC$
   End If
   Label35 = USERNAM$(USNBR%)
   HOII% = HOY(HOS$)
   Label34 = HOS$
   Text12 = FECHATEX$(DIPOST(HOII%))
   ' Al establecer 0 en InputLen se indica a MSComm que
   ' lea todo el contenido del búfer de entrada cuando
   ' se use la propiedad Input.
   'MSComm1.InputLen = 0
   'If Screen.Width < 12000 Then
   '   Hide
   '   FICONTAX.Show
   'End If
   '
   Call ABRECLIEN
   Call CREATABLA_CONTACSCOM
   '
   If EXISTE%("PRESUP07.EXE") > 0 Then Command32.Visible = True
   '
   MODIFNOMBRE = 0
   '
End Sub

Private Sub Command1_Click()
   '
   If TRIM$(Label26) = "" Then
     COMMAND16.SetFocus
     Exit Sub
   End If
   '
   Command1.Enabled = False
   If TRIM$(Label26) <> "" Then
     Screen.MousePointer = 11
      Call GRAFICON
     Screen.MousePointer = 1
    End If
   Command1.Enabled = True
   '
   If TRIM$(Text20) = "" Then
      Call COMUNI("Falta Ingresar Domicilio")
      Text20.SetFocus
      Exit Sub
   End If
   '
   If TRIM$(Text22) = "" Then
      Call COMUNI("Falta Ingresar Ciudad")
      Text22.SetFocus
      Exit Sub
   End If
   '
   If ECOARCH("provinc", Text27) = "" Then
      Call COMUNI("Falta Ingresar Provincia o Provincia No Valida")
      Text27.SetFocus
      Exit Sub
   End If
   
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
   'IMPORCONTAC.Show 1
   'Call IMPORAGE
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   '
   Static FIDATX1#, FIDATX2#, UCONTA$
   '
   Command13.Enabled = False
'   ABUSX$ = "BSEGCOM7": If USNBR% < 100 Then ABUSX$ = "BSEGALE"
'   If FIDATIM#(ABUSX$) = FIDAT1X# Then
'     If FIDATIM#("!SECONTA7") = FIDATX2# Then
'       JJ& = ULTREG&("!SECONTA7")
'       GoTo 20
'     End If
'   End If
   '
'   Call SETULTREG("SECONTA7", 0)

'   FIDATX1# = FIDATIM#(ABUSX$)
   JJ& = 0
   HOII% = HOY(HOS$)
   HOOOY = FETEXCOR1(HOII%)
   FINIPRO% = FECHANUM("30/04/06")  ' PARA EXCLUIR CONTACTOS VIEJOS
   FIPRO = FETEXCOR1(FINIPRO%)
   CONDI$ = "FechProxCto >= '" & FIPRO & "'" + " AND FechProxCto <= '" & HOOOY & "'" + " ORDER BY CAST(SUBSTRING(CODI_CLI,4,9) AS INTEGER)"
   Call CARGALISEL("SECONTA7", "BSEGCOM", "CODI_CLI/A12;RASO_CLI/A50", CONDI$)
   Call FRESHECHO("SECONTA7")
   Call SELECHO("SECONTA7")
   UCONTA$ = VALACT1$("SECONTA7")
     
   If UCONTA$ <> "" Then
     Label26 = UCONTA$
   End If

'   SQLX$ = "SELECT * FROM BSEGCOM WITH(NOLOCK) "
'   SQLX$ = SQLX$ + " WHERE FechProxCto >= '" & FIPRO & "'"
'   SQLX$ = SQLX$ + " AND FechProxCto <= '" & HOOOY & "'"
'   SQLX$ = SQLX$ + " ORDER BY CAST(SUBSTRING(CODI_CLI,4,9)AS INTEGER)"
'   Dim BSEGTMP As ADODB.Recordset
'   Set BSEGTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
'   With BSEGTMP
'    Do While Not .EOF
'        Call REPLA(EL1$, !Codi_Cli, 1, 12)
'        Call REPLA(EL1$, AJUSTI$(SAFETEXT$(!Raso_Cli), 50), 13, 50)   ' empresa
'        JJ& = JJ& + 1
'        Call GRAREG("!SECONTA7", EL1$, JJ&)
'    .MoveNext
'    Loop
'   End With
'   Call SETULTREG("!SECONTA7", JJ&)
'   '
'   Call BAJALDISCO
'   FIDATX2# = FIDATIM#("!SECONTA7")
'   Call FRESHECHO("!SECONTA7")
'   '
'20 If JJ& > 0 Then
'     FORSELEC.ABUSQUE = UCONTA$
'     Call SELECHO("!SECONTA7")
'     UCONTA$ = VALACT1$("!SECONTA7")
'
'     If UCONTA$ <> "" Then
'       Label26 = UCONTA$
'     End If
'   End If
   '
   Command13.Enabled = True
   '
End Sub

Private Sub Command14_Click()
   'If USNBR% <> 101 Then Exit Sub
   Command14.Enabled = False
   OPX% = COMALTER%("Opciones de Configuración:\\Conf.General\Particular")
   If OPX% = 1 Then
        Call CONECRUN("FSETUP04", "")
      ElseIf OPX% = 2 Then
        SETUPSGC.Show 1
   End If
   Command14.Enabled = True
End Sub

Private Sub Command21_Click()
   '
   Call ABRECLIEN
5  CONDICION$ = "CODI_CLI <> ''"
   NUMMME& = XVALO(VALOBADA("BSEGCOM", "MAX(CAST(SUBSTRING(CODI_CLI,4,9)AS integer))", CONDICION$, "NOMESS"))
   NCLI$ = "CX-" + TRIM$(Str(1 + NUMMME&))
   
   VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, NCLI$, 1, 12)
   Call REPLA(VDEF$, Space$(62), 13, 62)
10 OBX$ = OBJPLA$("ALTACLIENPROSP", VDEF$)
   If OBX$ = "" Then Label26 = "": Exit Sub
   VDEF$ = OBX$
   
   NCLI$ = TRIM$(Mid$(OBX$, 1, 12))
   If NCLI$ = "" Then
      Call MENSERR(24, "Número de Cuenta no Válido.")
      GoTo 10
   End If
   '
   CONDI$ = "Codi_cli='" & NCLI$ & "'"
   '
   NXLI$ = VALOBADA("BSEGCOM", "CODI_CLI", CONDI$, "NOMESS")
   If NXLI$ <> "" Then
      Call MENSERR(24, "Ingreso no Válido.\Código de Cuenta Pre-existente.")
      GoTo 5
   End If
   '
   RAZCL$ = TRIM$(Mid$(OBX$, 13, 62))
   If Len(RAZCL$) < 1 Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Razón Social del Cliente")
      GoTo 10
   End If
   '
   SQLX$ = "SELECT * FROM BSEGCOM"
   SQLX$ = SQLX$ + " WHERE Codi_Cli = '" & NCLI$ & "'"
   Dim BSEGTMP As ADODB.Recordset
   Set BSEGTMP = New ADODB.Recordset
   BSEGTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With BSEGTMP
       .AddNew   ' YA CONVERTIDO
       On Error Resume Next
       !CodiEmpr = CodiEmp%
       On Error GoTo 0
       !Codi_Cli = NCLI$
       !raso_cli = RAZCL$
       .Update
   End With
   BSEGTMP.Close
   Set BSEGTMP = Nothing
   '
   HOII% = HOY(HOS$)
   Label34 = HOS$
   Text12 = HOS$
   Label26 = NCLI$
   Text4.TEXT = RAZCL$
   '
   If USNBR% < 100 Then Label42 = TRIM$(Str$(USNBR%))
   '
   USNU% = USNBR% 'CARGA EL VENDEDOR SEGUN EL USUARIO
   If USNBR% > 100 Then '
     USNU% = USNBR% Mod 100
   End If
   NVEND% = 0
   For i& = 1 To ULTREG&("DATUSERS")
       X$ = REGLEIDO$("DATUSERS", i&)
       If Asc(Mid$(X$, 1, 1)) = USNU% Then
          NVEND% = Asc(Mid$(X$, 33, 2))
          Exit For
       End If
   Next i&
   Label42 = NVEND%
   '
   Call ActuFicha
   '
   MODIFNOMBRE = 0
   '
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
   COCLI$ = TRIM$(Label26)
   If COCLI$ <> "" Then
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
          TTXX1$ = TTXX1$ + TTXX$ + Chr$(10) + Chr$(13)
       Next KKI%
       Call REGAPP("ANOTASEG", MKS$(NPX&) + String$(88, "="))
       Call CIERRARCH("ANOTASEG")
       Text11 = ""
       SQLX$ = "SELECT HIS_CONTAC FROM BSEGCOM"
       SQLX$ = SQLX$ + " WHERE CODI_CLI = '" & COCLI$ & "'"
       Dim ANOTMP As New ADODB.Recordset
       Set ANOTMP = New ADODB.Recordset
       ANOTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       With ANOTMP
          If Not (.BOF And .EOF) Then
             ANOTA$ = SAFETEXT$(!HIS_CONTAC)
             ANOTA$ = ANOTA$ + TTXX1$
             ANOTA$ = ANOTA$ + String$(88, "=") + Chr$(10) + Chr$(13)
             !HIS_CONTAC = ANOTA$
             .Update
           End If
       End With
       ANOTMP.Close
       Set ANOTMP = Nothing

       '
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
   Call CIERRARCH("BSEGCOM7")
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

Private Sub Frame8_DragDrop(Source As CONTROL, X As Single, Y As Single)

End Sub

Private Sub Image4_Click()
   Image4.Enabled = False
   'Call Dial(TRIM$(Text26.TEXT), Name)
   Image4.Enabled = True
End Sub

Private Sub Image5_Click()
   Image5.Enabled = False
   'Call Dial(TRIM$(Text24.TEXT), Name)
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
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text10_GotFocus(Index As Integer)
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text11_DblClick()
   Text11.Enabled = False
   Call Dial(TRIM$(Text11.TEXT), Name)
   Text11.Enabled = True
End Sub

Private Sub Text11_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text12_DblClick()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text12.TEXT = VDEV$
    End If
End Sub

Private Sub Text12_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text15_GotFocus()
   Text4.SetFocus
End Sub

Private Sub label42_Change()
    VEND% = XVALO(Label42)
    LABEL40 = TRIM$(ECOARCH$("VENDEDOR", Chr$(VEND%)))
End Sub

Private Sub label42_DblClick()
    '
    Call Command31_Click
    '
End Sub

Private Sub label42_GotFocus()
   If USNBR% < 100 Then
      Text9.SetFocus
   End If
End Sub

Private Sub Text17_DblClick()
   Text17.Enabled = False
   'Call Dial(TRIM$(Text17.TEXT), Name)
   Text17.Enabled = True
End Sub

Private Sub Text17_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text18_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text27_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text2_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text20_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text21_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text22_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text23_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text24_DblClick()
   Text24.Enabled = False
   'Call Dial(TRIM$(Text24.TEXT), Name)
   Text24.Enabled = True
End Sub

Private Sub Text24_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text26_DblClick()
   Text26.Enabled = False
   'Call Dial(TRIM$(Text26.TEXT), Name)
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
   MSComm1.settings = SETTI$
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
    CODCLI$ = TRIM$(Label26)
    '
    'VALIDACION DEL TEXT EMPLEADO
    If XVALO(Text3) > 32766 Then
      Call MENSERR(24, "Números de Empleados Mayor al Admitido")
      GoTo 99
    End If
    Call ABRECLIEN
    SQLX$ = "SELECT * FROM BSEGCOM"
    SQLX$ = SQLX$ + " WHERE Codi_Cli = '" & CODCLI$ & "'"
    Dim FichaTmp As ADODB.Recordset
    Set FichaTmp = New ADODB.Recordset
    FichaTmp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With FichaTmp
     If .EOF Then
       .AddNew   ' YA CONVERTIDO
       On Error Resume Next
       !CodiEmpr = CodiEmp%
       On Error GoTo 0
       !Codi_Cli = CODCLI$
     End If
        !nume_cli = 0
        !raso_cli = AJUSTI$(Text4, 59) + AJUSTI$(Text15, 3) ' empresa
        !Mail_Cli = Left$(TRIM$(Text5), 64) 'DIRECCION ELECTRONICA
        !domi_cli = Left$(TRIM$(Text20), 64) ' CALLE Y NUMERO
        !Cpos_Cli = Left$(TRIM$(Text21), 16) ' C.POSTAL
        !loca_Cli = Left$(TRIM$(Text22), 48) ' LOCALIDAD
        !Prov_Cli = Left$(TRIM$(Text27), 2) ' PROVINCIA  **********Verificar
        !Pais_Cli = Left$(TRIM$(Text23), 48)    ' PAIS
        !ORIGEN = Left$(TRIM$(Text18), 64)  ' ORIGEN
        !Vend_Cli = XVALO(Label42)
        !RamoActi = Left$(TRIM$(Text1), 64) ' RAMO
        !ProduPrin = Left$(TRIM$(Text2), 128) ' PRINC.PRODUCTO
        !Empleados = XVALO(Text3) '****************Ver
        !Tele_Cli = Left$(TRIM$(Text26), 32)  ' TELEFONO
        !TELE_Cli2 = Left$(TRIM$(Text17), 32) ' telef.partic.
        !FAX_CLI = Left$(TRIM$(Text24), 24) ' telef2.partic
        !CTAC_CLI = Left$(TRIM$(Text6(0)), 48)
        !Cargo1 = Left$(TRIM$(Text7(0)), 32)
        !Sector1 = Left$(TRIM$(Text8(0)), 32)
        !Interno1 = Left$(TRIM$(Text10(0)), 32)
        !Contacto2 = Left$(TRIM$(Text6(1)), 48)
        !Cargo2 = Left$(TRIM$(Text7(1)), 32)
        !Sector2 = Left$(TRIM$(Text8(1)), 32)
        !Interno2 = Left$(TRIM$(Text10(1)), 32)
        !Contacto3 = Left$(TRIM$(Text6(2)), 48)
        !Cargo3 = Left$(TRIM$(Text7(2)), 32)
        !Sector3 = Left$(TRIM$(Text8(2)), 32)
        !Interno3 = Left$(TRIM$(Text10(2)), 32)
        !PagWeb = Left$(TRIM$(Text13), 48)
        !fechbaja = Label38
        !Mot_Baja = Left$(TRIM$(Text14), 64)
        '!FechUltEnvio =
        
        PUNGRU$ = ""
        For U& = 1 To List2.ListCount
          PGR% = XVALO(Mid$(List2.List(U& - 1), 73, 6))
          If PGR% < 256 Then PUNGRU$ = PUNGRU$ + Chr$(PGR%)
        Next U&
        !Grupos = PUNGRU$ + String$(30, 0)
        '
        FEC% = FECHANUM(Text12)
        !FechProxCto = CVDAT(FEC%)
        !nota_Cli = Text9
        .Update
   
    End With
    FichaTmp.Close
    Set FichaTmp = Nothing
    '
    If MODIFNOMBRE = 1 Then Call ActuFicha
99  Screen.MousePointer = 1
'
End Sub

Sub CARFICHA()
   '
   Screen.MousePointer = 1
   
5  COCLI$ = TRIM$(Label26)
   Call ABRECLIEN
   
   SQLX$ = "SELECT * FROM BSEGCOM WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE CODI_CLI = '" & COCLI$ & "'"
   Dim FICHTMP As ADODB.Recordset
   Set FICHTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With FICHTMP
      If .EOF Then
         Text1 = ""
         Text2 = ""
         Text3 = ""
         Text4 = ""
         Text5 = ""
         Text9 = ""
         Label42 = ""
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
         Text27 = ""
         Text23 = ""
         Text26 = ""
         Label42 = "": LABEL40 = ""
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
         'RELE$ = REGLEIDO$("BSEGCOM", REG&)
         Text4 = SAFETEXT$(TRIM$(Mid$(!raso_cli, 1, 59)))
         Text15 = SAFETEXT$(TRIM$(Mid$(!raso_cli, 60, 3)))
         Text5 = SAFETEXT$(!Mail_Cli)
         Text20 = SAFETEXT$(!domi_cli)
         Text21 = SAFETEXT$(!Cpos_Cli)
         '
         Text22 = SAFETEXT$(!loca_Cli)
         Text27 = SAFETEXT$(!Prov_Cli)
         Text23 = SAFETEXT$(!Pais_Cli)
         Text18 = SAFETEXT$(!ORIGEN)
         Text1 = SAFETEXT$(!RamoActi)
         Text2 = SAFETEXT$(!ProduPrin)
         Text3 = XVALO(!Empleados)
         '
         Text26 = SAFETEXT$(!Tele_Cli)
         Text17 = SAFETEXT$(!TELE_Cli2)
         Text24 = SAFETEXT$(!FAX_CLI)
         '
         Text6(0) = SAFETEXT$(!CTAC_CLI)
         Text7(0) = SAFETEXT$(!Cargo1)
         Text8(0) = SAFETEXT$(!Sector1)
         Text10(0) = SAFETEXT$(!Interno1)
         '
         Text6(1) = SAFETEXT$(!Contacto2)
         Text7(1) = SAFETEXT$(!Cargo2)
         Text8(1) = SAFETEXT$(!Sector2)
         Text10(1) = SAFETEXT$(!Interno2)
         '
         Text6(2) = SAFETEXT$(!Contacto3)
         Text7(2) = SAFETEXT$(!Cargo3)
         Text8(2) = SAFETEXT$(!Sector3)
         Text10(2) = SAFETEXT$(!Interno3)
         '
         Text13 = SAFETEXT$(!PagWeb)
         '
         IVENDE% = XVALO(!Vend_Cli)
         If IVENDE% = 8224 Then IVENDE% = 0
         Label42 = TRIM$(FORMATNUM$(IVENDE%, "I4"))
         '
         List2.Clear
         '
         PUNGRU$ = SAFETEXT$(!Grupos) 'Mid$(RELE$, 993, 30)
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
         FEBAJA$ = SAFETEXT$(!fechbaja)         ' FECHA DE BAJA
         Label38 = FEBAJA$
         '
         Text14 = SAFETEXT$(!Mot_Baja)                 ' MOTIVO DE BAJA
         '
         FE% = CVINT(SAFETEXT$(!FechProxCto))
         Text12 = FECHATEX(FE%)
         '
         Text9 = "" ': TTYZ$ = ""
         Text9 = SAFETEXT$(!nota_Cli)
         '
         ' BUSCA Y PRESENTA, SI EXISTE, EL NUMERO DE CLIENTE ASOCIADO AL PROSPECTO
         NCLI& = XVALO(VALOBADA("DEUDOR12", "Nume_Cli", "Codi_Cli='" & TRIM$(Label26) & "'"))
         Label41 = NCLI&
         '
      End If
   End With
   '
   FICHTMP.Close
   Set FICHTMP = Nothing
   '
   MODIFNOMBRE = 0
   '
   Screen.MousePointer = 1
End Sub

Sub SORCONTAC()
   '
   OPXX% = COMALTER%("Elija si Criterio de Ordenamiento:\\Por Codigo\Alfabético por Nombre\Cancelar")
   OPCI$ = "RSOCIAL"
   If OPXX% = 1 Then OPCI$ = "CODIGO"
   
   Call ActuFicha(OPCI$)
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
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub TEXT27_Change()
    LABEL52 = TRIM$(ECOARCH$("provinc", Text27))
End Sub

Private Sub Text3_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text4_Change()
    If TRIM$(Text4) <> "" Then
        MODIFNOMBRE = 1
    End If
End Sub

Private Sub Text4_GotFocus()
   If TRIM$(Label26) = "" Then
      On Error Resume Next ' evitar el error 5 al clickear en boton visual data
      COMMAND16.SetFocus
      On Error GoTo 0
   End If
End Sub

Public Sub Text5_DblClick()
   '
   COCLI$ = TRIM$(Label26)
   If COCLI$ <> "" Then
     '
     Call CREAMAIL(COCLI$)
     '
     ATENTI% = 1
     MAILMESS07.Label26 = Label26
     MAILMESS07.Show 1
   End If
   '
End Sub

Private Sub Text5_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text6_GotFocus(Index As Integer)
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text7_GotFocus(Index As Integer)
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text8_GotFocus(Index As Integer)
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
End Sub

Private Sub Text9_GotFocus()
   If TRIM$(Label26) = "" Then COMMAND16.SetFocus
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
   Screen.MousePointer = 11

   Dim NCON&(), PROCON%(), PRIMOV&()
   ' 1.- SELECCION CONTACTOS MARCADOS
   FIN& = ULTREG&("BSEGCOM7")
   CANSEL& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     RELE$ = REGLEIDO$("BSEGCOM7", U&)
     'If Mid$(RELE$, 62, 3) = "!!!" Then
       If TRIM$(Mid$(RELE$, 3, 59)) <> "" Then
         CANSEL& = CANSEL& + 1
         ReDim Preserve NCON&(CANSEL&), PROCON%(CANSEL&), PRIMOV&(CANSEL&)
         NCON&(CANSEL&) = XVALO(Mid$(RELE$, 4, 9))
         PROCON%(CANSEL&) = CVI(Mid$(RELE$, 1033, 2))
         PRIMOV&(CANSEL&) = 0
       End If
     'End If
    Next U&
   '
   ' 2.- ORDENAR POR FECHA PRÓXIMO CONTACTO
   For U1& = 1 To CANSEL&
     Call TRACE(20, U1&, CANSEL&)
     For U2& = U1& + 1 To CANSEL&
       If PROCON%(U2&) < PROCON%(U1&) Then
         NCONI% = NCON&(U2&)
         NCON&(U2&) = NCON&(U1&)
         NCON&(U1&) = NCONI%
         PROCONI% = PROCON%(U2&)
         PROCON%(U2&) = PROCON%(U1&)
         PROCON%(U1&) = PROCONI%
       End If
     Next U2&
   Next U1&
   '
   ' 3.- IDENTIFICAR ULTINA ANOTACIÓN
   For isel1& = 1 To CANSEL&
     Call TRACE(20, isel1&, CANSEL&)
     NPX& = NCON&(isel1&)
     RFF$ = RECFILT$("ANOTASEG", 1, MKS$(NPX&))
     '
     For U& = Len(RFF$) To 1 Step -4
       RENOTA& = CVS(Mid$(RFF$, U& - 3, 4))
       XX$ = REGLEIDO$("ANOTASEG", RENOTA&)
       If TRIM$(Mid$(XX$, 5, 8)) <> "" Then
         If TRIM$(Mid$(XX$, 5, 8)) <> "========" Then
           PRIMOV&(isel1&) = RENOTA&
           Exit For
         End If
       End If
     Next U&
   Next isel1&
   '
   ' 4.- GENERAR LISTADO
   REBLA$ = REGBLAN$("SEGOPEN7")
   JJ& = 0
   For isel1& = 1 To CANSEL&
     Call TRACE(20, isel1&, CANSEL&)
     NPX& = NCON&(isel1&)
     YY$ = FILTERGETRECORD$("BSEGCOM7", 1, MKS$(NCON&(isel1&)))
     EMPRA$ = Mid$(YY$, 3, 59)
     NOMCON$ = Mid$(YY$, 449, 32)
     TELEF$ = Mid$(YY$, 377, 24)
     '
     RFF$ = RECFILT$("ANOTASEG", 1, MKS$(NPX&))
     KL& = 0
     '
     For U& = 1 To Len(RFF$) Step 4
       RENOTA& = CVS(Mid$(RFF$, U&, 4))
       If RENOTA& >= PRIMOV&(isel1&) Then
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
           Call REPLA(ZZ$, MKI$(PROCON%(isel1&)), 137, 2)
           Call REPLA(ZZ$, NOMCON$, 139, 24)
           Call REPLA(ZZ$, TELEF$, 163, 24)
         End If
         Call REPLA(ZZ$, ANOTA$, 69, 68)
         KL& = KL& + 1
         JJ& = JJ& + 1
         Call GRAREG("SEGOPEN7", ZZ$, JJ&)
         ULTIA$ = TRIM$(ANOTA$)
       End If
29   Next U&
     JJ& = JJ& + 1
     Call GRAREG("SEGOPEN7", String$(256, "-"), JJ&)
   Next isel1&
   Screen.MousePointer = 1

   Call SETULTREG("SEGOPEN7", JJ&)
   Call CIERRARCH("SEGOPEN7")
   Call FINAL("*SEGOPEN7")
   '
End Sub
'
Sub ActualizaArchivo(OPCION%)
    Call ABRECLIEN
    '
    SQLX$ = "SELECT * FROM BSEGCOM "
    If OPCION% = 3 Then
      SQLX$ = SQLX$ + " WHERE SUBSTRING(RASO_CLI,60,3)= '!!!'"
      FIN& = CantRegistros&("BSEGCOM", "SUBSTRING(RASO_CLI,60,3)= '!!!'")
    End If
    
    SQLX$ = SQLX$ + " ORDER BY CAST(SUBSTRING(CODI_CLI,4,9)AS INTEGER)"
    If OPCION% < 3 Then
      FIN& = CantRegistros&("BSEGCOM")
    End If
    Dim BSEGTMP As ADODB.Recordset
    Set BSEGTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    Call SETULTREG("BSEGCOM7", 0)
    
    REG& = 0
    With BSEGTMP
     Do While Not .EOF
        REG& = REG& + 1
        Call TRACE(20, REG&, FIN&)
        REGI$ = SAFETEXT$(!Codi_Cli)
        RELE$ = REGI$ + Space$(1022)
        Call REPLA(RELE$, AJUSTI$(!raso_cli, 62), 3, 62)   ' empresa
        Call REPLA(RELE$, SAFETEXT(!Mail_Cli), 65, 64)     ' DIRECCION ELECTRONICA
        Call REPLA(RELE$, SAFETEXT(!domi_cli), 129, 48)   ' CALLE Y NUMERO
        Call REPLA(RELE$, SAFETEXT(!Cpos_Cli), 177, 8)    ' C.POSTAL
        Call REPLA(RELE$, SAFETEXT(!loca_Cli), 185, 32)   ' LOCALIDAD
         PROVI$ = TRIM$(SAFETEXT(!Prov_Cli))
        Call REPLA(RELE$, ECOARCH$("PROVINC", PROVI$), 217, 32)  ' PROVINCIA
        Call REPLA(RELE$, SAFETEXT(!Pais_Cli), 249, 24)   ' PAIS
        '
        Call REPLA(RELE$, SAFETEXT(!ORIGEN), 273, 24)   ' ORIGEN
        Call REPLA(RELE$, SAFETEXT(!RamoActi), 297, 24)    ' RAMO
        Call REPLA(RELE$, SAFETEXT(!ProduPrin), 321, 48)    ' PRINC.PRODUCTO
        TTYZ$ = SAFETEXT$(!nota_Cli)
        LOCADE& = Len(TTYZ$)
        For U& = LOCADE& To 1 Step -1
          If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 10
        Next U&
        U& = 0
        '
10      TEXTO1$ = Left$(TTYZ$, U&)
        While Len(TEXTO1$) > 0
           TTYY$ = MKI$(NCON%) + TEXTO1$ + Space$(62)
           KK& = KK& + 4
           If KK& <= Len(RNC$) - 3 Then
                RENOTA& = CVS(Mid$(RNC$, KK&, 4))
              Else
                URN& = URN& + 1
                RENOTA& = URN&
           End If
           Call GRAREG("NOTASEG", TTYY$, RENOTA&)
           TEXTO1$ = Mid$(TEXTO1$, 63)
        Wend
        '
        Call REPLA(RELE$, SAFETEXT(!PagWeb), 733, 64) ' WEB-SITE
        '
        IVENDE% = XVALO(!Vend_Cli)
        Call REPLA(RELE$, MKI$(IVENDE%), 900, 2)
        Call REPLA(RELE$, MKI$(FECHANUM(SAFETEXT$(!fechbaja))), 973, 2)   ' FECHA DE BAJA
        Call REPLA(RELE$, SAFETEXT$(!Mot_Baja), 975, 16)             ' MOTIVO DE BAJA
        Call REPLA(RELE$, MKI$(FECHANUM(SAFETEXT$(!FechProxCto))), 1023, 2)   ' FECHA PROXIMO CONTACTO
        HOII% = HOY(HOS$)
        If CVI(Mid$(RELE$, 1013, 2)) < HOII% Then
        Call REPLA(RELE$, MKI$(FECHANUM(SAFETEXT$(!FechUltEnvio))), 1025, 2)
        End If
        '
        Call GRAREG("BSEGCOM7", RELE$, REG&)
     .MoveNext
     Loop
    End With
    Call SETULTREG("BSEGCOM7", REG&)
    Call CIERRARCH("BSEGCOM7")
End Sub
'
Sub CREAMAIL(CODICLI$)
   '
   KOKLI$ = TRIM$(CODICLI$)
   '
   CONDICION$ = "CODI_CLI = '" & KOKLI$ & "'"
   '
   EMPRE$ = VALOBADA("BSEGCOM", "RASO_CLI", CONDICION$) 'TRIM$(Mid$(RELE$, 3, 59))
   CALNUM$ = VALOBADA("BSEGCOM", "DOMI_CLI", CONDICION$) 'TRIM$(Mid$(RELE$, 129, 48))
   CPOST$ = VALOBADA("BSEGCOM", "Cpos_Cli", CONDICION$) 'TRIM$(Mid$(RELE$, 177, 8))
   If CPOST$ <> "" Then CPOST$ = CPOST$ + " "
   LOCAX$ = CPOST$ + VALOBADA("BSEGCOM", "Loca_Cli", CONDICION$) ' CPOST$ + TRIM$(Mid$(RELE$, 185, 32))
   PROPA$ = VALOBADA("BSEGCOM", "Prov_Cli", CONDICION$) 'TRIM$(Mid$(RELE$, 217, 32))
   PRO$ = TRIM$(ECOARCH$("provinc", PROPA$))
   If PRO$ <> "" Then PRO$ = PRO$ + " - "
   
   PRO$ = PRO$ + VALOBADA("BSEGCOM", "Pais_Cli", CONDICION$) 'TRIM$(PROPA$ + TRIM$(Mid$(RELE$, 249, 24)))
   '
   CADEST% = 0
   TTXX$ = TRIM$(Mid$(RELE$, 449, 32))
   If TTXX$ <> "" Then
      CADEST% = CADEST% + 1
      DESTIX$ = TTXX$
   End If
   '
   TTXX$ = ""
   If EMPRE$ <> "" Then TTXX$ = TTXX$ + EMPRE$ + Chr$(13) + Chr$(10)
   If CALNUM$ <> "" Then TTXX$ = TTXX$ + CALNUM$ + Chr$(13) + Chr$(10)
   If LOCAX$ <> "" Then TTXX$ = TTXX$ + LOCAX$ + Chr$(13) + Chr$(10)
   If PRO$ <> "" Then TTXX$ = TTXX$ + PRO$ + Chr$(13) + Chr$(10)
   '
   MMT4$ = TRIM$(DESTIX$)
   If MMT4$ <> "" Then MMT4$ = MMT4$ + " - "
   MAILMESS07.Text4 = TRIM$(MMT4$ + EMPRE$)
     PPXX% = InStr(FISECOM07.Text5, "/")
     If PPXX% < 1 Then PPXX% = 1 + Len(FISECOM07.Text5)
   MAILMESS07.Text5 = TRIM$(Left$(FISECOM07.Text5, PPXX% - 1))
   MAILMESS07.Text20 = CONTROL$("SEGUICOM", "ASUNTO")
   '
   HOII% = HOY(HOS$)
   MAILMESS07.Label5 = "Buenos Aires, " + FETEXTO$(HOII%) + ".-"
   While MAILMESS07.Label5.Width < MAILMESS07.Text2.Width - 1500
     MAILMESS07.Label5 = " " + MAILMESS07.Label5
   Wend
   '
   TTYY$ = MAILMESS07.Label5 + Chr$(13) + Chr$(10)
   If CADEST% < 1 Then
       TTYY$ = TTYY$ + "Sres." + Chr$(13) + Chr$(10)
    Else
       TTYY$ = TTYY$ + "Sr./Sra." + Chr$(13) + Chr$(10)
       TTYY$ = TTYY$ + DESTIX$ + Chr$(13) + Chr$(10)
   End If
   '
   TTYY$ = TTYY$ + TTXX$
   SST = 4 + Len(TTYY$)
   '
   For KKII% = 1 To 5: TTYY$ = TTYY$ + Chr$(13) + Chr$(10): Next KKII%
   TTYY$ = TTYY$ + LOADFILE$(LUDAT$ + "PIEDEMAI.DAT")
   '
   MAILMESS07.Text2 = TTYY$ ' = TTYY$ + Text1
   MAILMESS07.Text2.SelStart = SST
   MAILMESS07.Text2.SelLength = 0
   '
   RFF$ = RECFILT$("ANOTASEG", 1, MKS$(NPX&))
   MAILMESS07.List1.Clear
   '
   Call ABRECLIEN
   
   ANOTACION$ = SAFETEXT$(VALOBADA("BSEGCOM", "HIS_CONTAC", CONDICION$))
   XX$ = ANOTACION$: LONGI& = Len(XX$)
   INICIO& = 1
   Do While INICIO& < LONGI&
        FIN& = InStr(INICIO&, XX$, Chr$(10))
        If FIN& > 0 Then
           TTXX$ = Mid$(XX$, INICIO&, FIN& - INICIO&)
           MAILMESS07.List1.AddItem TTXX$
           INICIO& = FIN& + 2
        Else
           TTXX$ = Mid$(XX$, INICIO&)
           MAILMESS07.List1.AddItem TTXX$
           Exit Do
        End If
   Loop
   TPU& = MAILMESS07.List1.ListCount - 7
   If TPU& < 0 Then TPU& = 0
   MAILMESS07.List1.TopIndex = TPU&
   '
End Sub




