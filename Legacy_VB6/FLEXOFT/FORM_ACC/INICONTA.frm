VERSION 5.00
Begin VB.Form INICONTA 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de Contabilidad General"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11880
   Icon            =   "INICONTA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11880
   Begin VB.PictureBox Picture9 
      Height          =   8000
      Left            =   0
      ScaleHeight     =   7935
      ScaleWidth      =   10995
      TabIndex        =   45
      Top             =   480
      Width           =   11055
      Begin VB.Image Image1 
         Height          =   8200
         Left            =   0
         Picture         =   "INICONTA.frx":1D2A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   11055
      End
   End
   Begin VB.CommandButton Command16 
      Caption         =   "EDENOR A PROVEEDORES"
      Height          =   495
      Left            =   8640
      TabIndex        =   44
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
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
      Height          =   600
      Left            =   9765
      Picture         =   "INICONTA.frx":28B00
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Informes Standard Normalizados"
      Top             =   4725
      Width           =   560
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   540
      Left            =   105
      TabIndex        =   41
      Top             =   3045
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton Command9 
      Caption         =   "CLIENTES"
      Height          =   375
      Left            =   0
      TabIndex        =   40
      Top             =   2625
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.CommandButton Command8 
      Caption         =   "BLOQUEAR"
      Height          =   375
      Left            =   0
      TabIndex        =   39
      Top             =   2205
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ListBox List2 
      Height          =   645
      Left            =   5760
      TabIndex        =   38
      Top             =   480
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00800000&
      Height          =   8600
      Left            =   11010
      ScaleHeight     =   8535
      ScaleWidth      =   855
      TabIndex        =   5
      Top             =   0
      Width           =   915
      Begin VB.PictureBox Picture16 
         Height          =   8050
         Left            =   0
         ScaleHeight     =   7995
         ScaleWidth      =   1035
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   1100
         Begin VB.Image Image3 
            Height          =   8100
            Left            =   0
            Picture         =   "INICONTA.frx":28E0A
            Stretch         =   -1  'True
            Top             =   -120
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Direct"
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
         Left            =   105
         Picture         =   "INICONTA.frx":2B1A2
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Acceso Directo a Aplicaciones Complementarias"
         Top             =   4410
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         TabIndex        =   30
         Top             =   3520
         Width           =   615
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   555
         TabIndex        =   28
         Top             =   5400
         Width           =   615
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Plan"
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
         Left            =   120
         Picture         =   "INICONTA.frx":2B4AC
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Plan General de Cuentas"
         Top             =   5835
         Width           =   615
      End
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
         Height          =   675
         Left            =   120
         Picture         =   "INICONTA.frx":2B5F6
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2850
         Width           =   615
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
         Height          =   675
         Left            =   120
         Picture         =   "INICONTA.frx":2B900
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Consultas y Listados Varios"
         Top             =   7200
         Width           =   615
      End
      Begin VB.CommandButton Command4 
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
         Height          =   675
         Left            =   120
         Picture         =   "INICONTA.frx":2BC0A
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Setup - Configuración de Funcionamiento"
         Top             =   6525
         Width           =   615
      End
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
         Height          =   675
         Left            =   120
         Picture         =   "INICONTA.frx":2C04C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1130
         Width           =   615
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
         Height          =   675
         Left            =   120
         Picture         =   "INICONTA.frx":2C356
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cierra y Abandona la Aplicacion"
         Top             =   450
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00800000&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   11835
      TabIndex        =   31
      Top             =   0
      Width           =   11895
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7440
         TabIndex        =   33
         Top             =   40
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   32
         Top             =   45
         Width           =   3735
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ejercicio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5760
         TabIndex        =   35
         Top             =   90
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -360
         TabIndex        =   34
         Top             =   90
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   1485
      Left            =   1560
      Picture         =   "INICONTA.frx":2C4A0
      ScaleHeight     =   1485
      ScaleWidth      =   4800
      TabIndex        =   26
      Top             =   3600
      Width           =   4800
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   480
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
      Height          =   600
      Left            =   5760
      Picture         =   "INICONTA.frx":44722
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6060
      Width           =   600
   End
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
      Left            =   8160
      Picture         =   "INICONTA.frx":44A2C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3600
      Width           =   600
   End
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
      Left            =   5535
      Picture         =   "INICONTA.frx":4605E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1230
      Width           =   600
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
      Height          =   600
      Left            =   1680
      Picture         =   "INICONTA.frx":461A8
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6300
      Width           =   600
   End
   Begin VB.CommandButton Command11 
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
      Left            =   1725
      Picture         =   "INICONTA.frx":464B2
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   900
      Width           =   600
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   1800
      Picture         =   "INICONTA.frx":467BC
      ScaleHeight     =   1170
      ScaleWidth      =   2235
      TabIndex        =   4
      Top             =   6540
      Width           =   2235
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Gestión de"
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
         TabIndex        =   21
         Top             =   360
         Width           =   1590
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Valores"
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
         TabIndex        =   20
         Top             =   590
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   5820
      Picture         =   "INICONTA.frx":4E182
      ScaleHeight     =   1170
      ScaleWidth      =   2235
      TabIndex        =   3
      Top             =   6330
      Width           =   2235
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Informes "
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
         TabIndex        =   23
         Top             =   365
         Width           =   1590
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Varios"
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
         TabIndex        =   22
         Top             =   590
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   8280
      Picture         =   "INICONTA.frx":55B48
      ScaleHeight     =   1170
      ScaleWidth      =   2220
      TabIndex        =   2
      Top             =   3885
      Width           =   2220
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   19
         Top             =   555
         Width           =   1800
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Opciones de "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   105
         TabIndex        =   18
         Top             =   315
         Width           =   1800
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   5640
      Picture         =   "INICONTA.frx":5D50E
      ScaleHeight     =   1155
      ScaleWidth      =   2115
      TabIndex        =   1
      Top             =   1500
      Width           =   2115
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "y Mayor"
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
         Left            =   220
         TabIndex        =   17
         Top             =   585
         Width           =   1590
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Libro Diario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   210
         TabIndex        =   16
         Top             =   340
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   1830
      Picture         =   "INICONTA.frx":65008
      ScaleHeight     =   1170
      ScaleWidth      =   2010
      TabIndex        =   0
      Top             =   1170
      Width           =   2010
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Contables"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Width           =   1590
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Asientos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   210
         TabIndex        =   24
         Top             =   345
         Width           =   1590
      End
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   36
      Top             =   480
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Line Line7 
      X1              =   3960
      X2              =   6360
      Y1              =   4440
      Y2              =   2040
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   255
      Left            =   5400
      TabIndex        =   37
      Top             =   600
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Line Line6 
      X1              =   2700
      X2              =   3960
      Y1              =   1800
      Y2              =   4320
   End
   Begin VB.Line Line5 
      X1              =   3960
      X2              =   2760
      Y1              =   4320
      Y2              =   7200
   End
   Begin VB.Line Line4 
      X1              =   3840
      X2              =   9360
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      X1              =   3960
      X2              =   6720
      Y1              =   4320
      Y2              =   7080
   End
   Begin VB.Menu Asi_Cont 
      Caption         =   "Asientos Contables"
      Begin VB.Menu Nue_Asi 
         Caption         =   "Nuevo Asiento"
         Begin VB.Menu Asi_Man 
            Caption         =   "Asiento Manual"
         End
         Begin VB.Menu plh7 
            Caption         =   "-"
         End
         Begin VB.Menu Asi_Ven 
            Caption         =   "Resumen de Ventas"
         End
         Begin VB.Menu Asi_Com 
            Caption         =   "Resumen de Compras"
         End
         Begin VB.Menu li9 
            Caption         =   "-"
         End
         Begin VB.Menu Asi_Caj 
            Caption         =   "Resumen de Cobranza"
         End
         Begin VB.Menu Asi_Pag 
            Caption         =   "Resumen de Pagos"
         End
         Begin VB.Menu li91 
            Caption         =   "-"
         End
         Begin VB.Menu Asi_CaBan 
            Caption         =   "Resumen de Caja y Bancos"
         End
         Begin VB.Menu li32 
            Caption         =   "-"
         End
         Begin VB.Menu Asi_Cie 
            Caption         =   "Asiento de Cierre"
         End
      End
      Begin VB.Menu Open_Asi 
         Caption         =   "Abrir Asiento Registrado"
      End
      Begin VB.Menu plh8 
         Caption         =   "-"
      End
      Begin VB.Menu Print1 
         Caption         =   "Imprimir"
         Enabled         =   0   'False
      End
      Begin VB.Menu plh9 
         Caption         =   "-"
      End
      Begin VB.Menu Exit1 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu Mnu_Qwery 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu Diarge 
         Caption         =   "Diario General"
      End
      Begin VB.Menu Mayor 
         Caption         =   "Mayor por Cuenta"
      End
      Begin VB.Menu Mnu_Plc 
         Caption         =   "-"
      End
      Begin VB.Menu Obalan 
         Caption         =   "Opciones de Balance"
         Begin VB.Menu Estcuen 
            Caption         =   "Estado de Cuentas"
         End
         Begin VB.Menu Resumen 
            Caption         =   "Resultados Mensuales"
         End
         Begin VB.Menu Plc1 
            Caption         =   "-"
         End
         Begin VB.Menu Susaldos 
            Caption         =   "Sumas y Saldos"
         End
      End
      Begin VB.Menu PLH20 
         Caption         =   "-"
      End
      Begin VB.Menu Con_sql 
         Caption         =   "Informes Normalizados"
      End
   End
   Begin VB.Menu Config 
      Caption         =   "Configuración"
      Begin VB.Menu mnu_emp 
         Caption         =   "Empresas"
         Begin VB.Menu Nuempre 
            Caption         =   "Nueva Empresa"
         End
         Begin VB.Menu Modempre 
            Caption         =   "Modificar Datos"
         End
         Begin VB.Menu Borrempre 
            Caption         =   "Eliminar"
         End
      End
      Begin VB.Menu plh1 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Ejer 
         Caption         =   "Ejercicios"
         Begin VB.Menu Abrejer 
            Caption         =   "Abrir Nuevo Ejercicio"
         End
         Begin VB.Menu Correjer 
            Caption         =   "Corregir Datos"
         End
      End
      Begin VB.Menu plh10 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Plan 
         Caption         =   "Plan de Cuentas"
         Begin VB.Menu Editplan 
            Caption         =   "Abrir para Editar"
         End
         Begin VB.Menu Copyplan 
            Caption         =   "Importar por Copia"
         End
      End
      Begin VB.Menu plh11 
         Caption         =   "-"
      End
      Begin VB.Menu Setup 
         Caption         =   "Configuración Avanzada"
      End
   End
   Begin VB.Menu Mnu_Setup 
      Caption         =   "Herramientas"
      Begin VB.Menu Mnu_Import 
         Caption         =   "Importar Asientos"
      End
      Begin VB.Menu Aju_Impor 
         Caption         =   "Ajuste Importación CC"
      End
   End
   Begin VB.Menu accesodirecto 
      Caption         =   "Accesos Directos"
      Begin VB.Menu tesoreria 
         Caption         =   "Tesorería"
      End
   End
End
Attribute VB_Name = "INICONTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Contablori As Database
Dim Plancuenori As Recordset
Dim NUREPLANORI As Integer

Private Sub Aju_Impor_Click()
   '
   OBX$ = OBJPLA$("AJUIMPORCC", "")
   If OBX$ <> "" Then
     FF% = CVI(Mid$(OBX$, 1, 2))
     IMAJU# = CVD(Mid$(OBX$, 3, 8))
     CuentaCo% = CUEINT%(Mid$(OBX$, 11, 12))
     If CuentaCo% > 0 Then
       If Abs(IMAJU#) >= 0.05 Then
         CoComCor$ = "AJUS-INI"
         CoComLar$ = "AJUS-INI"
         CodiEmp% = 0
         CodiSuc% = SUC%
         NumeClie% = 0
         NumeProv% = 0
         FechEmi = CVDAT(FF%)
         '
         Referen$ = "Ajuste Importacion Saldos CC"
         FeContab = CVDAT(FF%)
         FeAcredi = CVDAT(FF%)
         NumeTal = SUC%
         NumeCom = 0
         Observa$ = "Ajuste Importacion Saldos CC"
         '
         TiMovi = "AJUSTE"            ' DEPOSITO BANCARIO
         OpMovi = "INICIAL"           ' PLAZO DE ACREDITACION
         VaMovi = ""                   '
         Descrip = "Ajuste Importacion Saldos CC"
         ImpoDeb# = IMAJU#
         StatInit = 0
         Call GRACABAN
         '
         Call COMUNI("Procesado")
         Exit Sub
         '
       End If
     End If
   End If
   '
   MsgBox "Transaccion Cancelada"
   '
End Sub

Private Sub Asi_CaBan_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
   Call ASIDIAR.Asi_CaBan_Click
   ASIDIAR.Show 1
End Sub

Private Sub Asi_Caj_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
   Call ASIDIAR.Asi_Caj_Click
   ASIDIAR.Show 1
End Sub

Private Sub Asi_Cie_Click()
   '
    ACIERRE04.Show 1
   '
End Sub

Private Sub Asi_Com_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
   Call ASIDIAR.Asi_Com_Click
   ASIDIAR.Show 1
End Sub

Private Sub Asi_Man_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
   Call ASIDIAR.Asi_Man_Click
   ASIDIAR.Show 1
End Sub

Private Sub Asi_Pag_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
   Call ASIDIAR.Asi_Pag_Click
   ASIDIAR.Show 1
End Sub

Private Sub Asi_Ven_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
   Call ASIDIAR.Asi_Ven_Click
   ASIDIAR.Show 1
End Sub

Private Sub Borrempre_Click()
   Call CARMULTIEM
End Sub

Private Sub Combo1_Click()
  '
  Call CIERRARCH("*.*")
  Call BAJALDISCO
  '
  ICOMBO& = Combo1.ListIndex
  LUDA1$ = ULODATOS + TRIM$(Mid$(List1.List(ICOMBO&), 31)) + "."
  LUDA2$ = ULODATOS + TRIM$(Mid$(List1.List(ICOMBO&), 31)) + "\"
  '
  EXISZ$ = ""
  On Error Resume Next
  EXISZ$ = TRIM$(Dir(LUDA1$, 16))
  On Error GoTo 0
  '
  If EXISZ$ = "" Then
     MsgBox "No Existe Carpeta " + LUDA2$
     Call FUERASYS
  End If
  '
  EMPREAC$ = Left$(List1.List(ICOMBO&), 30)
  LUSELEC$ = TRIM$(Mid$(List1.List(ICOMBO&), 31))
  LUDAT$ = LUDA2$
  '
  FF00$ = LOADFILE$("FLEXCFG.DAT")
  Call REPLA(FF00$, LUSELEC$, 51, 45)
  Call SAVEFILE("FLEXCFG.DAT", FF00$)
  '
  RAMSTRI$ = Left$(LOADFILE("C:\FLEXOFT\FLEXRAM.DRV") + String$(512, 0), 512)
  Call REPLA(RAMSTRI$, LUDA2$, 99, 46)
  Call REPLA(RAMSTRI$, EMPREAC$, 193, 30)
  Call SAVEFILE("C:\FLEXOFT\FLEXRAM.DRV", RAMSTRI$)
  '
  VUELTA% = 0
5 EXISZ$ = ""
  On Error Resume Next
  EXISZ$ = TRIM$(Dir(LUDAT$ + "BdAccess.", 16))
  On Error GoTo 0
  If EXISZ$ = "" Then
    If VUELTA% = 0 Then
      VUELTA% = VUELTA% + 1
      MkDir LUDAT$ + "BdAccess\"
      GoTo 5
    End If
    MsgBox "No Existe Carpeta " + LUDAT$ + "BdAccess\"
    Call FUERASYS
  End If
  '
  LUACC$ = LUDAT$ + "BdAccess\"
  Call ABREPLANCUEN
  Call CARLISEJER
  '
  COMEJ% = 0: FINEJ% = 0: DENOEJ$ = "": NEJER$ = ""
  '
  Set TABEJER = Contable.OpenRecordset("SELECT * FROM TABEJCON WHERE STATUS > 0 ORDER BY FECHAPER DESC;", dbOpenDynaset, dbReadOnly)
  On Error Resume Next
  TABEJER.MoveFirst
  If Err < 1 Then
    With TABEJER
      DENOEJ$ = TRIM$(!Denoejcon)
      COMEJ% = FECHANUM(Format(!Fechaper, "DD/MM/YY"))
      FINEJ% = FECHANUM(Format(!Fechcierr, "DD/MM/YY"))
      NEJER$ = TRIM$(Str$(!CODIEJER)): While Len(NEJER$) < 2: NEJER$ = "0" + NEJER$: Wend
    End With
  End If
  TABEJER.Close
  '
  Call GRARAM(87, MKI$(COMEJ%))
  Call GRARAM(89, MKI$(FINEJ%))
  Call GRARAM(321, AJUSTI$(DENOEJ$, 48))
  Call GRARAM(369, NEJER$)
  Combo2.TEXT = DENOEJ$
  If COMEJ% < 1 Then
    Call MENSERR(24, "No Hay Ejercicios Abiertos\en la Empresa Elegida")
  End If
  '
End Sub

Private Sub Combo2_Click()
  '
  ICOMBO& = Combo2.ListIndex
  TTXX$ = List2.List(ICOMBO&)
  DENOEJ$ = TRIM$(Left$(TTXX$, 48))
  COMEJ% = Val(Mid$(TTXX$, 49, 8))
  FINEJ% = Val(Mid$(TTXX$, 57, 8))
  NEJER$ = Mid$(TTXX$, 65, 2)
  '
  Call GRARAM(87, MKI$(COMEJ%))
  Call GRARAM(89, MKI$(FINEJ%))
  Call GRARAM(321, AJUSTI$(DENOEJ$, 48))
  Call GRARAM(369, NEJER$)
  '
End Sub

Private Sub Command1_Click()
  Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
  Call CIERRARCH("*.*")
  Call FINAL("")
End Sub

Private Sub Command10_Click()
   Call ABRECABAN
   With CajaBanc
     .MoveFirst
     Do While Not .EOF
       On Error Resume Next
       AA1$ = !NuAsiDiar
       If Err Then
         .Edit
         !NuAsiDiar = ""
         .Update
       End If
       On Error GoTo 0
     .MoveNext
     Loop
   End With
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     GoTo 99
   End If
   '
   ASIDIAR.Show 1
99 Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     GoTo 99
   End If
   '
   OPX% = COMALTER%("Opciones de Trabajo:\\Diario General\Mayor por Cuenta")
   Select Case OPX%
   Case 1
      Call Diarge_Click
   Case 2
      Call Mayor_Click
   End Select
   '
99 Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   If EXISTE%("FLGEST00.EXE") > 0 Then
      Call MENSERR(24, "Imposible Acceder Desde Esta Aplicación.\  \Cierre Módulo Contable\e Ingrese por Tesorería del Escritorio Principal.")
      Call FINAL("")
      Exit Sub
   End If
   '
   If COMEJ% > 0 Then
    If DERACCE%("GESTVAL", "Gestion de Cheques y Valores") >= 2 Then
      GESVAL04.Show 1
    End If
   End If
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   INFONORM.Show 1
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     GoTo 99
   End If
   '
   OPBALAN04.Show 1
   '
99 Command15.Enabled = True
End Sub

Private Sub Command16_Click()
'   If CLACONTRA% > 0 Then
'        '
'        FEINI = Int(CDbl(CVDAT(COMEJ%)))
'        FEFIN = Int(CDbl(CVDAT(FINEJ%)))
'        '
'        Call ABRECABAN
'        CUEII% = 1
'        '
'        SQLX$ = "UPDATE CAJABANC SET CUECONTAI = 1 "
'        SQLX$ = SQLX$ + "WHERE CUECONTAI = 4 "
'        SQLX$ = SQLX$ + " AND CDBL(FECHCONTA) >= " & FEINI & " And CDbl(FECHCONTA) <= " & FEFIN
'        SQLX$ = SQLX$ + " AND (TIPOMOVIM = 'CHEMI' OR OPCIMOVIM = 'OP')"
'        SQLX$ = SQLX$ + " AND (NuAsidiar = '' or isnull(NuAsiDiar) = true OR IsEmpty(NuAsiDiar) = true) "
'
'        CueCorri.Execute (SQLX$)
'        '
'   End If
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   PopupMenu Mnu_Qwery
   Command18.Enabled = True
End Sub

Private Sub Command2_Click()
'   If EXISTE%("C:\FLEXOFT\MANUALES\" + App.EXEName + ".PDF") > 0 Then
'      On Error Resume Next
'      FORHLPDF.Pdf1.LOADFILE "C:\FLEXOFT\MANUALES\" + App.EXEName + ".PDF"
'      If Err Then
'         On Error GoTo 0
'         Call MENSERR(24, "Imposible Abrir Archivo de Ayuda.\  \Falta Instalar Control Pdf1.OCX.")
'         Exit Sub
'      End If
'      FORHLPDF.Show 1
'   End If
End Sub

Private Sub Command27_Click()
   Command27.Enabled = False
   Call ACCDIR("CONTAGEN")
   Command27.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   FPLANCUEN.Show 1
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call CONECRUN("FSETUP04", "Configuración de Funcionamiento")
   Call FINAL("")
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   'AVAMRPSB.Show 1
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call MENSERR(24, "No Disponible en la Presente Versión")
   'Call GRAJOBID
   'Call BAJRAM
   'Shell LUBAS$ + "WINTABU " + "#" + JOBID$ + "#", 1
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   '
   Command7.Enabled = False
   If COMALTER%("Está Seguro de Importar Asientos ?\\No, Cancelar\Si, Importar") <> 2 Then GoTo 99
   '
   CAREPLA& = 0
   On Error Resume Next
   Plancuen.MoveLast
   CAREPLA& = Plancuen.RecordCount
   On Error GoTo 0
   If CAREPLA& < 1 Then
     Call MENSERR(24, "Imposible Importar.\Abra y Guarde Plan General de Cuentas.\  \Luego Re-Intente Importación")
     Exit Sub
   End If
   '
   Screen.MousePointer = 11
   EJER% = (-1)
   EXTE$ = "OLD"
   '
   Call ABREPLANCUEN
   Screen.MousePointer = 11
   Set Encasien = Contable.OpenRecordset("ENCABEASIEN", dbOpenTable)
   Set Detasien = Contable.OpenRecordset("DETAASIEN", dbOpenTable)
   '
   On Error Resume Next
   Encasien.MoveLast
   CARE% = 0
   CARE% = Encasien.RecordCount
   If CARE% > 0 Then
      While Encasien.RecordCount > 0
        Encasien.MoveLast
        Encasien.Delete
      Wend
   End If
   '
5  Detasien.MoveLast
   CARE% = 0
   CARE% = Detasien.RecordCount
   If CARE% > 0 Then
      While Detasien.RecordCount > 0
        Detasien.MoveLast
        Detasien.Delete
      Wend
   End If
   On Error GoTo 0
   '
6  Encasien.Close
   Detasien.Close
   Set Encasien = Contable.OpenRecordset("ENCABEASIEN", dbOpenTable)
   Set Detasien = Contable.OpenRecordset("DETAASIEN", dbOpenTable)
   '
10 FIN& = MAXENCDIARIO%
   ORDAS% = 0: NUIDANT$ = ""
   UDIA% = 0
   For KI% = 1 To FIN&
     U& = KI%: Call TRACE(20, U&, FIN&)
     Call LEENCDIARIO(KI%)
     NUASI = CVS(E1$)
     NUASDEF& = CVI(E2$)
     FEASX = CVI(E3$)
       FECHAS$ = FECHATEX$(FEASX)
       NUIDAS$ = Mid$(FECHAS$, 7, 2) + Mid$(FECHAS$, 4, 2) + Mid$(FECHAS$, 1, 2)
       If NUIDAS$ <> NUIDANT$ Then ORDAS% = 0
       NUIDANT$ = NUIDAS$
       If Left$(NUIDAS$, 1) = "0" Then NUIDAS$ = Mid$(NUIDAS$, 2)
       ORDAS% = ORDAS% + 1
       ORDA$ = TRIM$(Str$(ORDAS%))
       While Len(ORDA$) < 3: ORDA$ = "0" + ORDA$: Wend
       NUIDAS$ = NUIDAS$ + "-" + ORDA$
     SUDEBE# = CVD(E4$)
     CONTA% = CVI(E6$)
     REFEAS$ = E7$
     TIASI% = Asc(Left$(E8$, 1))
     POEDIT% = 1: If (TIASI% > 0 And TIASI% < 4) Then POEDIT% = (-1)
     '
     
     With Encasien
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        !NUIDASIEN = NUIDAS$
        !Refeasien = REFEAS$
        !FECHASIEN = CVDAT(FECHANUM(FECHAS$))
        !OrFeAsien = ORDAS%
        !SucuAsien = 0
        !NuDfAsien = NUASDEF&
        !TIPOASIEN = TIASI%
        !StatAsien = 1  ' activo
        !EditAsien = POEDIT%
        !QPasAsien = CONTA%
        !DebeAsien = SUDEBE#
        !Observa = " "
        .Update
        .Bookmark = .LastModified
     End With
     '
     ORPASE% = 0: SUDEBE# = 0
     For K2% = UDIA% + 1 To MAXDIARIO%
       Call LEEDIARIO(K2%)
       If CVS(E12$) > NUASI + 0.01 Then
          UDIA% = K2% - 1
          GoTo 35
       End If
       If Abs(CVS(E12$) - NUASI) < 0.005 Then
          ORPASE% = ORPASE% + 1
          CUENTAINT% = CVI(E11$)
            COCUE$ = COCUENTA$(CUENTAINT%)
            CUEIII% = CUEINT%(COCUE$)
            If CUEIII% < 1 Then
              VEDECUE$ = LOADFILE$(LUDAT$ + "DENOCUE." + EXTE$)
              IDECU& = 1 + 128& * (CUENTAINT% - 1)
              DECUE$ = TRIM$(Mid$(VEDECUE$, IDECU&, 100))
              Call AGRECUENTA(COCUE$, DECUE$, CUEIII%)
            End If
          IMPOPAS# = CVD(E16$)
          '
          DETAPAS$ = ""
          If K2% <= LOF(NX12%) / 32 Then
            Get #NX12%, K2%, E20$
            DETAPAS$ = TRIM$(E20$)
          End If
          If DETAPAS$ = "" Then DETAPAS$ = DENOCUE$(CVI(E11$))
          If DETAPAS$ = "" Then DETAPAS$ = " "
          '
          IDEBE# = IMPOPAS#: IHABE# = 0
          If IMPOPAS# < 0 Then
            IDEBE# = 0: IHABE# = Abs(IMPOPAS#)
            SUDEBE# = SUDEBE# + IHABE#
          End If
          '
          AGREPLAN% = 1
          With Detasien
             .AddNew   ' YA CONVERTIDO
             On Error Resume Next
             !CodiEmpr = CodiEmp%
             On Error GoTo 0
             !NUIDASIEN = NUIDAS$
             !OrNuPase = ORPASE%
             !REFEPASE = DETAPAS$
             !FECHASIEN = CVDAT(FECHANUM(FECHAS$))
             !OrFeAsien = ORDAS%
             !SucuAsien = 0
             !NuDfAsien = NUASDEF&
             !StatPase = 1
             !CUECOINT = CUEIII%
             !CUECOEXT = COCUE$
                If !CUECOEXT = "" Then !CUECOEXT = " "
             !IDEBEPASE = IDEBE#
             !IHABEPASE = IHABE#
             .Update
             .Bookmark = .LastModified
          End With
          AGREPLAN% = 0
       End If
     Next K2%
     '
35   With Encasien
       .Edit
       !DebeAsien = SUDEBE#
       .Update
       .Bookmark = .LastModified
     End With
     '
   Next KI%
   '
   If EJER% < 0 Then
     Call CIERRACONTA
     EJER% = 1
     EXTE$ = "DAT"
     GoTo 10
   End If
   '
   Encasien.Close
   Detasien.Close
   Screen.MousePointer = 1
   '
   Call COMUNI("Importacion Completa. Debe Ahora Abrir Manualmente\los Ejercicios Contables Importados.\  \Re-Inicie Ahora la Aplicación.")
   Call CIERRARCH("*.*")
   Call FINAL("")
   '
99 Command7.Enabled = True
   '
End Sub

Function COCUENTA$(CUEII%)
   '
   Static VECUENTA$, EXTEA$, LOREL&
   '
   If EXTE$ <> EXTEA$ Then
     LOREL& = RECLEN("CUENTAS")
     EXTEA$ = EXTE$
     VECUENTA$ = LOADFILE$(LUDAT$ + "CUENTAS." + EXTE$)
   End If
   '
   COCUENTA$ = TRIM$(Mid$(VECUENTA$, 1 + LOREL& * (CUEII% - 1), 12))
   '
End Function
'
Sub AGRECUENTA(COCUE$, DECUE$, CUEIII%)
   '
      '
20 On Error Resume Next
   Set Plancuen = Contable.OpenRecordset("PLANCUEN", dbOpenDynaset)
   If Err Then
      If COMALTER%("No se Puede Abrir 'PLANCUEN'\en Modo Exclusivo.\\Reintentar\Cancelar") <> 2 Then GoTo 20
      Close: End
   End If
   '
   CAREPLA& = 0
   On Error Resume Next
   Plancuen.MoveLast
   CAREPLA& = Plancuen.RecordCount
   On Error GoTo 0
   '
   UCUE% = 0
   With Plancuen
     For U& = 1 To CAREPLA&
       If U& = 1 Then
           .MoveFirst
         Else
           .MoveNext
       End If
       If !CODINCUE > UCUE% Then
         UCUE% = !CODINCUE
       End If
     Next U&
     CUEIII% = 1 + UCUE%
     '
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CodiEmpr = CodiEmp%
     On Error GoTo 0
     !CODEXCUE = COCUE$
     !DENOCUE = DECUE$
     !CODINCUE = CUEIII%
     !STATCUE = 9   ' ANULADA
     !NIVAGRU = 4   ' ULTIMO NIVEL
     !TipoCue = "Debe"
     !POSIMPU = 0   ' NO IMPUTABLE
     .Update
   End With
   Plancuen.Close
   '
End Sub

Private Sub Command8_Click()
'   Command8.Enabled = False
'   Call BLOQTABLA("CONTAGEN/PLANCUEN")
'   MsgBox "BLOQUEADA"
'   Call LIBERTABLA("CONTAGEN/PLANCUEN")
'   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Call ABRECLIEN
End Sub

Private Sub Con_sql_Click()
   Call Command14_Click
End Sub

Private Sub Copyplan_Click()
   ORICOPLAN.Show 1
   '
   LUACCORI$ = TRIM$(ORICOPLAN.LUACCORI)
   If LUACCORI$ <> "" Then
      On Error Resume Next
      Set Contablori = OpenDatabase(LUACCORI$ + "CONTAGEN.MDB")
      If Err Then
         MsgBox "No Existe " + LUACCORI$ + "CONTAGEN.MDB"
         Exit Sub
      End If
      '
      Load FPLANCUEN
      If FPLANCUEN.List1.ListCount > 0 Then
        If COMALTER%("Desea Sobre-Escribir\el Plan de Cuentas Existente ?\\No, Cancelar\Si, Sobre-Escribir") <> 2 Then
          Exit Sub
        End If
      End If
      Set Plancuenori = Contablori.OpenRecordset("PLANCUEN", dbOpenDynaset, dbReadOnly)
      NUREPLANORI% = 0
      On Error Resume Next
      Plancuenori.MoveLast
      NUREPLANORI% = Plancuenori.RecordCount
      On Error GoTo 0
      If NUREPLANORI% < 1 Then
        Call MENSERR(24, "No Hay Plan de Cuentas\en la Empresa Origen Elegida")
        Exit Sub
      End If
      '
      Call CARPLANORI
      FPLANCUEN.Show 1
      '
   End If
End Sub

Sub CARPLANORI()
   '
   Screen.MousePointer = 11
   FPLANCUEN.List1.Clear
   '
   Call ABREPLANCUEN
   '
   FIN& = NUREPLANORI%
   With Plancuenori
     .MoveFirst
     U& = 1
     ZZ$ = Space$(256)
     '
10   COCUE$ = !CODEXCUE
     DECUE$ = !DENOCUE
     NIVEI% = !NIVAGRU
       NVX$ = Left$(String$(NIVEI%, ".") + TRIM$(Str$(NIVEI%)) + ".....", 5)
       '
     Call REPLA(ZZ$, NVX$, 1, 5)
     Call REPLA(ZZ$, COCUE$, 9, 12)
     Call REPLA(ZZ$, DECUE$, 23, 44)
     Call REPLA(ZZ$, DECUE$, 129, 128)
     '
     TICUEX$ = TRIM$(UCase$(Left$(!TipoCue, 1)))
     If TICUEX$ = "H" Then
         TICUEX$ = "Haber"
       ElseIf TICUEX$ = "I" Then
         TICUEX$ = "Ingresos"
       ElseIf TICUEX$ = "G" Then
         TICUEX$ = "Gastos"
       Else
         TICUEX$ = "Debe"
     End If
     Call REPLA(ZZ$, TICUEX$, 69, 5)
     '
     IMPUBLE$ = ""
     If !POSIMPU < 1 Then IMPUBLE$ = "No"
     Call REPLA(ZZ$, IMPUBLE$, 76, 2)
     '
     STATUX$ = ""
     If !STATCUE = 9 Then STATUX$ = "Baja"
     Call REPLA(ZZ$, STATUX$, 80, 4)
     '
     NUINT$ = AJUSTD$(Str$(!CODINCUE), 8)
     Call REPLA(ZZ$, NUINT$, 121, 8)
     '
     FPLANCUEN.List1.AddItem ZZ$
     '
     If U& < FIN& Then
       .MoveNext
       U& = U& + 1
       GoTo 10
     End If
     '
   End With
   '
90 FPLANCUEN.List1.Refresh
   Screen.MousePointer = 1
   '
End Sub
'

Private Sub Correjer_Click()
   '
   If DERACCE%("ABREJER", "Apertura de Ejercicio Contable") = 2 Then
     '
     If COMEJ% < 1 Or Val(NEJER$) < 1 Then
       Call MENSERR(24, "No Hay Ejercicios Abiertos\en la Empresa Elegida")
       Call Abrejer_Click
       Exit Sub
     End If
     '
     On Error Resume Next
     Tabejcon.Close
     On Error GoTo 0
     Set Tabejcon = Contable.OpenRecordset("TABEJCON", dbOpenDynaset)
     With Tabejcon
       .FindFirst "CODIEJER = " & Val(NEJER$)
15     DENOEJ$ = TRIM$(!Denoejcon)
       COMEJ% = FECHANUM(Format(!Fechaper, "DD/MM/YY"))
       FINEJ% = FECHANUM(Format(!Fechcierr, "DD/MM/YY"))
       VDEF$ = AJUSTI$(DENOEJ$, 28) + MKI$(COMEJ%) + MKI$(FINEJ%)
       VNUE$ = OBJPLA$("DATEJCONTA", VDEF$)
       If VNUE$ <> "" Then
         DENUE$ = Left$(VNUE$, 28)
         COMEJ% = CVI(Mid$(VNUE$, 29, 2))
         FINEJ% = CVI(Mid$(VNUE$, 31, 2))
         If FINEJ% < COMEJ% Then
           Call MENSERR(24, "Fecha de Cierre es Anterior a Apertura.")
           GoTo 15
         End If
         '
         .Edit
         !Denoejcon = DENUE$
         !Fechaper = CVDAT(COMEJ%)
         !Fechcierr = CVDAT(FINEJ%)
         .Update
         '
         Tabejcon.Close
         Call CARLISEJER
         Call GRARAM(87, MKI$(COMEJ%))
         Call GRARAM(89, MKI$(FINEJ%))
         Call GRARAM(321, AJUSTI$(DENOEJ$, 48))
         Call GRARAM(369, NEJER$)
         Combo2.TEXT = DENUE$
         '
       End If
     End With
     On Error Resume Next
     Tabejcon.Close
     On Error GoTo 0
     '
   End If
   '
End Sub

Sub Diarge_Click()
   '
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
   Set MASTERECO = Contable.OpenRecordset("SELECT * FROM PLANCUEN ORDER BY CODEXCUE ASC;", dbOpenDynaset, dbReadOnly)
   On Error Resume Next
   MASTERECO.MoveLast
   If Err Then Exit Sub
   '
   CAREC% = MASTERECO.RecordCount
   If CAREC% < 1 Then Exit Sub
   '
   With MASTERECO
      .MoveFirst
      DIARGEN.Text1 = !CODEXCUE
      .MoveLast
      DIARGEN.Text2 = !CODEXCUE
   End With
   '
   HOII% = HOY(HOS$)
   FEX = COMEJ%: If FEX > HOII% Then FEX = HOII%
   DIARGEN.Label5 = FECHATEX$(FEX)
   FEX = FINEJ%: If FEX > HOII% Then FEX = HOII%
   DIARGEN.Label6 = FECHATEX$(FEX)
   Screen.MousePointer = 1
   '
   DIARGEN.Show 1
   Exit Sub
   
   
   
   
   
   HOII% = HOY(HOS$)
   FEX = FINIPDIF%
   HAS% = HOII%
   DES% = FECHANUM("01" + Mid$(HOS$, 3))
   If HAS% > FINEJ% Then
     DES% = COMEJ%: HAS% = FINEJ%
   End If
3  VDEF$ = MKI$(DES%) + MKI$(HAS%)
   '
5  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   On Error Resume Next
   FEX = CVI(Left$(OBX$, 2))
   DESDU$ = FECHATEX$(FEX)
   DESYY = CDbl(CDate(DESDU$))
   If Err Then GoTo 3
   FEX = CVI(Mid$(OBX$, 3, 2))
   HASDU$ = FECHATEX$(FEX)
   HASDU$ = Left$(HASDU$, 6) + "20" + Mid$(HASDU$, 7)
   HASYY = CDbl(CDate(HASDU$))
   If Err Then GoTo 3
   '
8  On Error GoTo 0
   SUCUI% = 0
   'Set DIARPAN1.Data1.Recordset = Contable.OpenRecordset("SELECT * FROM DETAASIEN WHERE (CDBL(FECHASIEN) >= " & DESYY & " AND CDBL(FECHASIEN) <= " & HASYY & " AND SUCUASIen = " & SUCUI% & ") ORDER BY  FECHASIEN ASC;", dbOpenDynaset, dbReadOnly)
   SQLX$ = "SELECT FECHASIEN, NUIDASIEN, REFEPASE, CUECOEXT, IDEBEPASE, IHABEPASE "
   SQLX$ = SQLX$ + " FROM DETAASIEN WHERE "
   SQLX$ = SQLX$ + "(CDBL(FECHASIEN) >= " & DESYY & " "
   SQLX$ = SQLX$ + "AND CDBL(FECHASIEN) <= " & HASYY & " "
   SQLX$ = SQLX$ + "AND SUCUASIen = " & SUCUI% & " "
   SQLX$ = SQLX$ + "AND STATPASE > 0) "
   SQLX$ = SQLX$ + "ORDER BY NUIDASIEN, ORNUPASE;"
   Set DIARPAN.Data1.Recordset = Contable.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   On Error Resume Next
   '
   DIARPAN.Data1.Refresh
   DIARPAN.DBGrid1.Refresh
   'DIARPAN1.Data1.MoveLast
   'If Err Then
   '   Call MENSERR(24, "Sin Movimientos en el Periodo Elegido")
   '   GoTo 3
   'End If
   '
   On Error GoTo 0
   Set COLUI = DIARPAN.DBGrid1.Columns
   COLUI(0).Caption = "Fecha"
   COLUI(1).Caption = "Nro-Id"
   COLUI(2).Caption = "Referencia"
   COLUI(3).Caption = "Cuenta"
   COLUI(4).Caption = "Debe"
   COLUI(5).Caption = "Haber"
   COLUI(4).NumberFormat = "########0.00"
   COLUI(5).NumberFormat = "########0.00"
   COLUI(0).Width = 1110
   COLUI(1).Width = 1095
   COLUI(2).Width = 4425
   COLUI(3).Width = 1320
   COLUI(4).Width = 1580
   COLUI(5).Width = 1580
'DIARPAN1.Data1.Recordset = SELECASIEN
DIARPAN.Show 1
'GoTo 8
Exit Sub
   'DIARPAN.List1.Clear
   SELECASIEN.MoveFirst
   IPASE% = 0: TRANSPOR# = 0
   '
10 IPASE% = IPASE% + 1
   '
   With SELECASIEN
     Z$ = Space$(128)
       FEXA$ = Format(!FECHASIEN, "DD/MM/YY")
     Call REPLA(Z$, FEXA$, 1, 8)
     Call REPLA(Z$, !NUIDASIEN, 10, 10)
       'Call REPLA(Z$, FORMATNUM$(!NUDFASIEN, "F6.0"), 21, 6)
       DEPASE$ = !REFEPASE
       If TRIM$(DEPASE$) = "" Then DEPASE$ = DENOCUE$(CUEII%)
     Call REPLA(Z$, DEPASE$, 21, 27)
       IDEBE# = !IDEBEPASE
       IHABE# = !IHABEPASE
       If IDEBE# > 0 Then
           Call REPLA(Z$, CSTRING$(MKD$(IDEBE#), 3, 12, 2, 2), 49, 12)
         Else
           Call REPLA(Z$, CSTRING$(MKD$(IHABE#), 3, 12, 2, 2), 61, 12)
       End If
       TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
       Call REPLA(Z$, CSTRING$(MKD$(TRANSPOR#), 3, 12, 2, 2), 73, 13)
       'DIARPAN.List1.AddItem Z$
     End With
     '
     If IPASE% < CAREL& Then
       SELECASIEN.MoveNext
       GoTo 10
     End If
     '
     'If MAYORPAN.List1.ListCount > 0 Then
     '  MAYORPAN.List1.ListIndex = MAYORPAN.List1.ListCount - 1
     'End If
     '
     'MAYORPAN.Ltext2 = Text11
     'MAYORPAN.Label4 = DENOCUE$(CUEII%)
     DIARPAN.Show 1
     '
   GoTo 3
End Sub

Private Sub Editplan_Click()
   Call Command3_Click
End Sub

Private Sub Estcuen_Click()
   OPBALAN04.Option1(2).Value = 1
   On Error Resume Next
   OPBALAN04.Command1.SetFocus
   On Error GoTo 0
   OPBALAN04.Show 1
End Sub

Private Sub Exit_Click()
   Call Command1_Click
End Sub

Private Sub Exit1_Click()
   Call Command1_Click
End Sub

Private Sub Form_Load()
   '
   CCOLO% = 193
   BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture1.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture2.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture3.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture4.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture5.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture6.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Picture8.BackColor = &HFCD7B6
      MARCOBARRA.Top = 7700
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If

   If MAXENCDIARIO% < 1 Then
      Command7.Enabled = False
   End If
   '
   Call CENTRAFORM(Me)
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
     Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   Call CARLISEMP
   RAMSTRI$ = Left$(LOADFILE(LUCOM$ + "FLEXRAM.DRV") + String$(512, 0), 512)
   LUDAT$ = TRIM$(Mid$(RAMSTRI$, 99, 46))
   
   Label13 = LUDAT$
   Combo1.TEXT = TRIM$(Mid$(RAMSTRI$, 193, 30))
   '
   DENOEJ$ = TRIM$(Mid$(RAMSTRI$, 321, 48))
   COMEJ% = CVI(Mid$(RAMSTRI$, 87, 2))
   FINEJ% = CVI(Mid$(RAMSTRI$, 89, 2))
   NEJER$ = Mid$(RAMSTRI$, 369, 2)
   Combo2.TEXT = DENOEJ$
   '
   VUELTA% = 0
5  EXISZ$ = ""
   On Error Resume Next
   EXISZ$ = TRIM$(Dir(LUDAT$ + "BdAccess.", 16))
   On Error GoTo 0
   If EXISZ$ = "" Then
     If VUELTA% = 0 Then
       VUELTA% = VUELTA% + 1
       MkDir LUDAT$ + "BdAccess\"
       GoTo 5
     End If
     MsgBox "No Existe Carpeta " + LUDAT$ + "BdAccess\"
     Call FUERASYS
   End If
   '
   LUACC$ = LUDAT$ + "BdAccess\"
   '
   Call ABREPLANCUEN
   Call CARLISEJER
   '
   COMEJ% = 0: FINEJ% = 0: DENOEJ$ = "": NEJER$ = ""
   '
   Set TABEJER = Contable.OpenRecordset("SELECT * FROM TABEJCON WHERE STATUS > 0 ORDER BY FECHAPER DESC;", dbOpenDynaset, dbReadOnly)
   On Error Resume Next
   TABEJER.MoveFirst
   If Err < 1 Then
     With TABEJER
       DENOEJ$ = TRIM$(!Denoejcon)
       COMEJ% = FECHANUM(Format(!Fechaper, "DD/MM/YY"))
       FINEJ% = FECHANUM(Format(!Fechcierr, "DD/MM/YY"))
       NEJER$ = TRIM$(Str$(!CODIEJER)): While Len(NEJER$) < 2: NEJER$ = "0" + NEJER$: Wend
     End With
   End If
   TABEJER.Close
   '
   Call GRARAM(87, MKI$(COMEJ%))
   Call GRARAM(89, MKI$(FINEJ%))
   Call GRARAM(321, AJUSTI$(DENOEJ$, 48))
   Call GRARAM(369, NEJER$)
   Combo2.TEXT = DENOEJ$
   '
   If APLINVO$ = "PLANCUEN" Then
     FPLANCUEN.Show 1
     Call FINAL("")
   End If
   '
   If COMEJ% < 1 Then
     Call MENSERR(24, "No Hay Ejercicios Abiertos\en la Empresa Elegida")
   End If
   '
'If HOY(HOS$) > FECHANUM("01/12/07") Then
'   AA1$ = LOADFILE$("FLOGON.EXE")
'   Call SAVEFILE("FLOCONT.EXE", AA1$)
'End If
'If EXISTE%("FLOCONT.EXE") > 0 Then
'   Call COMUNI("Plazo de Demostración Vencido.")
'   Call FINAL("")
'End If
''
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
   Call CIERRARCH("*.*")
   Cancel = 0
   Call FINAL("")
End Sub

Sub CARLISEMP()
   '
   Combo1.Clear: List1.Clear
   If EXISTE%(CurDir + "\FLEXMUL.DRV") = 1 Then
     LIEMP$ = LOADFILE$(CurDir + "\FLEXMUL.DRV")
     Dim FMX As String * 79
     '
     For i% = 1 To Len(LIEMP$) Step 79
       FMX$ = Mid$(LIEMP$, i%, 79)
       SUCAR% = 0
       For J% = 1 To 77
         SUCAR% = SUCAR% + Asc(Mid$(FMX$, J%, 1))
       Next J%
       If CVI(Mid$(FMX$, 78, 2)) = SUCAR% Then
         RASO$ = DESENCRI$(Left$(FMX$, 30))
         PACCE$ = DESENCRI$(Mid$(FMX$, 32, 46))
         Z$ = Space$(96)
         Call REPLA(Z$, RASO$, 1, 30)
         Call REPLA(Z$, PACCE$, 51, 46)
         Combo1.AddItem TRIM$(Left$(Z$, 30))
         List1.AddItem Z$
       End If
     Next i%
     '
   End If
   '
End Sub

Sub CARLISEJER()
   '
   Combo2.Clear: List2.Clear
   Set TABEJER = Contable.OpenRecordset("SELECT * FROM TABEJCON WHERE STATUS > 0 ORDER BY FECHAPER DESC;", dbOpenDynaset, dbReadOnly)
   CAEJA% = 0
   On Error Resume Next
   TABEJER.MoveLast
   CAEJA% = TABEJER.RecordCount
   On Error GoTo 0
   '
   If CAEJA% > 0 Then
     With TABEJER
       .MoveFirst
       IEJU% = 1
10     DENOEJI$ = TRIM$(!Denoejcon)
       COMEJI% = FECHANUM(Format(!Fechaper, "DD/MM/YY"))
       FINEJI% = FECHANUM(Format(!Fechcierr, "DD/MM/YY"))
       NEJERI$ = TRIM$(Str$(!CODIEJER)): While Len(NEJERI$) < 2: NEJERI$ = "0" + NEJERI$: Wend
       Combo2.AddItem DENOEJI$
       TTXX$ = Space$(80)
       Call REPLA(TTXX$, DENOEJI$, 1, 48)
       Call REPLA(TTXX$, TRIM$(Str$(COMEJI%)), 49, 8)
       Call REPLA(TTXX$, TRIM$(Str$(FINEJI%)), 57, 8)
       Call REPLA(TTXX$, NEJERI$, 65, 2)
       List2.AddItem TTXX$
       '
       If IEJU% < CAEJA% Then
         IEJU% = IEJU% + 1
         .MoveNext
         GoTo 10
       End If
     End With
   End If
   TABEJER.Close
   '
End Sub


Private Sub Abrejer_Click()
   '
   If DERACCE%("ABREJER", "Apertura de Ejercicio Contable") = 2 Then
     '
     VDEF$ = String$(80, 0)
     Call REPLA(VDEF$, MKI$(COMEJ%), 1, 2)
     Call REPLA(VDEF$, DENOEJ$, 3, 32)
     FEX = FINEJ%
     CMJ1% = DIPOST(FEX)
     Call REPLA(VDEF$, MKI$(CMJ1%), 41, 2)
     Call REPLA(VDEF$, "", 43, 32)
     '
10   OBX$ = OBJPLA$("ABREJER", VDEF$)
     If OBX$ = "" Then Exit Sub
     '
     VDEF$ = OBX$
     DENUE$ = TRIM$(Mid$(OBX$, 43, 32))
     If DENUE$ = "" Then
       Call MENSERR(24, "Falta Denominación del Nuevo Ejercicio")
       GoTo 10
     End If
     '
     CMJACT% = FINEJ%
     APNUE% = CVI(Mid$(OBX$, 41, 2))
     'If APNUE% < CMJACT% Then
     '  Call MENSERR(24, "Nueva Apertura es Anterior\al Cierre del Ejercicio Actual")
     '  GoTo 10
     'End If
     '
     FEX = APNUE%
     APNU$ = FECHATEX$(FEX)
     FEAPNU$ = Left$(APNU$, 6) + "20" + Mid$(APNU$, 7)
     If Val(Mid$(APNU$, 7)) > 80 Then FEAPNU$ = Left$(APNU$, 6) + "19" + Mid$(APNU$, 7)
     ANX% = Val(Mid$(APNU$, 7, 2)) + 1: If ANX% > 99 Then ANX% = ANX% - 100
     AN$ = TRIM$(Str$(ANX%)): While Len(AN$) < 2: AN$ = "0" + AN$: Wend
     FE$ = Left$(APNU$, 6) + AN$
     FEX = FECHANUM(FE$)
     CIENUE% = DIANTE(FEX)
     FEXX = CIENUE%
     FE$ = FECHATEX$(FEXX)
     FECINU$ = Left$(FE$, 6) + "20" + Mid$(FE$, 7)
     If Val(Mid$(FE$, 7)) > 80 Then FECINU$ = Left$(FE$, 6) + "19" + Mid$(FE$, 7)
     '
     On Error Resume Next
     Tabejcon.Close
     On Error GoTo 0
     Set Tabejcon = Contable.OpenRecordset("TABEJCON", dbOpenTable)
     
     With Tabejcon
       UREI% = 0
       On Error Resume Next
       .MoveLast
       UREI% = .RecordCount
       On Error GoTo 0
       .AddNew   ' YA CONVERTIDO
       On Error Resume Next
       !CodiEmpr = CodiEmp%
       On Error GoTo 0
       !CODIEJER = 1 + UREI%
       !Denoejcon = DENUE$
       !Fechaper = CDate(FEAPNU$)
       !Fechcierr = CDate(FECINU$)
       !Status = 1
       !Diverif = 0
       !Marcacti = 1
       .Update
       .Bookmark = .LastModified
     End With
     Tabejcon.Close
     '
     DENOEJ$ = TRIM$(DENUE$)
     COMEJ% = APNUE%
     FINEJ% = CIENUE%
     NEJER$ = TRIM$(Str$(1 + UREI%)): While Len(NEJER$) < 2: NEJER$ = "0" + NEJER$: Wend
     '
     Call GRARAM(87, MKI$(COMEJ%))
     Call GRARAM(89, MKI$(FINEJ%))
     Call GRARAM(321, AJUSTI$(DENOEJ$, 48))
     Call GRARAM(369, NEJER$)
     Combo2.TEXT = DENOEJ$
     '
     Call COMUNI("Nuevo Ejercicio O.K.")
     '
   End If
   '
End Sub

Sub Mayor_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
10 RG1$ = REGSEL$("ECUECOT")
   If Len(RG1$) > 4 Then
     COCUE$ = Left$(RG1$, 12)
     CUEII% = CUEINT%(COCUE$)
     If CUEII% > 0 Then
       Call LMAYORPAN(CUEII%)
       GoTo 10
     End If
   End If
End Sub

Private Sub Mnu_Import_Click()
   Call Command7_Click
End Sub

Private Sub Modempre_Click()
   Call CARMULTIEM
End Sub

Private Sub Nuempre_Click()
   Call CARMULTIEM
End Sub

Private Sub Open_Asi_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
   Call ASIDIAR.Open_Asi_Click
   ASIDIAR.Show 1
End Sub

Private Sub Resumen_Click()
   OPBALAN04.Option1(1).Value = 1
   On Error Resume Next
   OPBALAN04.Command1.SetFocus
   On Error GoTo 0
   OPBALAN04.Show 1
End Sub

Private Sub setup_Click()
   Call Command4_Click
End Sub

Private Sub Susaldos_Click()
   OPBALAN04.Option1(4).Value = 1
   On Error Resume Next
   OPBALAN04.Command1.SetFocus
   On Error GoTo 0
   OPBALAN04.Show 1
End Sub

Sub LMAYORPAN(CUEII%)
   '
'   FEINI = Int(CDbl(CVDAT(COMEJ%)))
'   FEFIN = Int(CDbl(CVDAT(FINEJ%)))
   FEINI = Int(CDbl(CVDAT(COMEJ%)))
   FEFIN = Int(CDbl(CVDAT(FINEJ%)))
   '
   Call BLOQARCH("LIMAYOAC")
   Screen.MousePointer = 11
   '
   SQLX$ = "SELECT * FROM DetaAsien "
   SQLX$ = SQLX$ + "WHERE (CueCoInt=" & CUEII%
   SQLX$ = SQLX$ + " AND SucuAsien = " & SUCUI%
   'SQLX$ = SQLX$ + " AND CDBL(Fechasien) >= " & FEINI
   SQLX$ = SQLX$ + " AND Fechasien <= " & FEFIN
   SQLX$ = SQLX$ + " AND statpase >0 ) "
   SQLX$ = SQLX$ + "ORDER BY FECHASIEN, ORFEASIEN ASC;"
   Set PASEASIEN = Contable.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   '
   JJ& = 0
   TRANSPOR# = 0
   REBLA$ = REGBLAN$("LIMAYOAC")
10 With PASEASIEN
     On Error Resume Next
     .MoveFirst
     If Err Then GoTo 30
     On Error GoTo 0
     Do While Not .EOF
       IDEBE# = 0
       IHABE# = 0
       MACON$ = ""
       DEPASE$ = ""
       CENCO$ = ""
       On Error Resume Next
       IDEBE# = (Int(100 * !IDEBEPASE + 0.5)) / 100
       IHABE# = (Int(100 * !IHABEPASE + 0.5)) / 100
       If Abs(IDEBE#) < 0.005 And Abs(IHABE#) < 0.005 Then GoTo 19 ' ELIMINA REGISTROS EN BLANCO
       '
       If Int(CDbl(!FECHASIEN)) < FEINI Then
           TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
         ElseIf Int(CDbl(!FECHASIEN)) <= FEFIN Then
           Z$ = REBLA$
           Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 1, 2)
           Call REPLA(Z$, !NUIDASIEN, 3, 12)
             MACON$ = TRIM$(!MARCONSI)
           Call REPLA(Z$, MACON$, 15, 1)
           Call REPLA(Z$, "D", 16, 1)
             DEPASE$ = !REFEPASE
           Call REPLA(Z$, DEPASE$, 17, 32)
           Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
           Call REPLA(Z$, MKD$(IHABE#), 57, 8)
             CENCO$ = !CenCos
           On Error GoTo 0
           Call REPLA(Z$, CENCO$, 65, 8)
           JJ& = JJ& + 1
           Call GRAREG("LIMAYOAC", Z$, JJ&)
       End If
19     On Error GoTo 0
     .MoveNext
     Loop
   End With
   '
30 On Error GoTo 0
   Call ABRECABAN
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where (CueContaI = " & CUEII% & " "
   SQLX$ = SQLX$ + " AND (NuAsidiar = '' or isnull(NuAsiDiar) = true OR IsEmpty(NuAsiDiar) = true)) "
   'SQLX$ = SQLX$ + " AND CDBL(FechConta) >= " & FEINI
   SQLX$ = SQLX$ + " AND FechConta <= " & FEFIN & " "
   SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC, FEHOREAL ASC;"
   '
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   With CABATEMP
     On Error Resume Next
     .MoveLast
     If Err Then
        GoTo 40
     End If
     On Error GoTo 0
     FIN& = .RecordCount
     .MoveFirst
     JJ1& = 0
     Do While Not .EOF
       JJ1& = JJ1& + 1
       Call TRACE(20, JJ1&, FIN&)
       IDEBE# = 0
       IHABE# = 0
       MACON$ = ""
       DEPASE$ = ""
       DE2PASE$ = ""
       CENCO$ = ""
       On Error Resume Next
         IDEBE# = (Int(100 * !IDEBECOMP + 0.5)) / 100
         IHABE# = (Int(100 * !IHABECOMP + 0.5)) / 100
         If Abs(IDEBE#) < 0.005 And Abs(IHABE#) < 0.005 Then GoTo 39 ' ELIMINA REGISTROS EN BLANCO
       If Int(CDbl(!FechConta)) < FEINI Then
           TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
         ElseIf Int(CDbl(!FechConta)) <= FEFIN Then
           Z$ = REBLA$
           Call REPLA(Z$, MKI$(CVINT(!FechConta)), 1, 2)
           Call REPLA(Z$, !CodiCom_Cor, 3, 12)
             MACON$ = TRIM$(!MARCONSI)
           Call REPLA(Z$, MACON$, 15, 1)
           Call REPLA(Z$, "C", 16, 1)
             DEPASE$ = !RefeCom
             DE2PASE$ = !DescriMov
           Call REPLA(Z$, DEPASE$, 17, 32)
           Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
           Call REPLA(Z$, MKD$(IHABE#), 57, 8)
             CENCO$ = !CenCosto
           Call REPLA(Z$, CENCO$, 65, 8)
           Call REPLA(Z$, DE2PASE$, 73, 56)
           JJ& = JJ& + 1
           Call GRAREG("LIMAYOAC", Z$, JJ&)
       End If
39   On Error GoTo 0
     .MoveNext
     Loop
   End With
   '
40 Call SETULTREG("LIMAYOAC", JJ&)
   Call FILESORT("LIMAYOAC", "", 1, 1, "ASC")
   '
   MAYORPAN.List1.Clear
   Z$ = Space$(128 + 64)
     FEX = DIANTE(COMEJ%)
     FEXA$ = FECHATEX$(FEX)
   Call REPLA(Z$, FEXA$, 1, 8)
     DEPASE$ = "TRANSPORTE .............................."
   Call REPLA(Z$, DEPASE$, 35, 30)
   Call REPLA(Z$, CSTRING$(MKD$(TRANSPOR#), 3, 14, 2, 2), 94, 14)
   MAYORPAN.List1.AddItem Z$
   MAYORPAN.List1.AddItem Space$(95) + " =========== "
   '
   SUDEBE1# = 0: SUHABE1# = 0
   For U& = 1 To ULTREG("LIMAYOAC")
     X$ = REGLEIDO$("LIMAYOAC", U&)
     Z$ = Space$(128 + 64 + 8)
       FEX = CVI(Left$(X$, 2))
       FEXA$ = FECHATEX$(FEX)
     Call REPLA(Z$, FEXA$, 1, 8)
       COCOM$ = TRIM$(Mid$(X$, 3, 12))
     Call REPLA(Z$, COCOM$, 10, 14)
       MACON$ = TRIM$(Mid$(X$, 15, 1))
     Call REPLA(Z$, MACON$, 24, 2)
       ORIMO$ = TRIM$(Mid$(X$, 16, 1))
     Call REPLA(Z$, ORIMO$, 128, 1)
       DEPASE$ = TRIM$(Mid$(X$, 17, 32))
       DE2PASE$ = TRIM$(Mid$(X$, 73, 56))
     Call REPLA(Z$, DEPASE$, 35, 30)
     Call REPLA(Z$, DE2PASE$, 129, 64)
       CENCO$ = TRIM$(Mid$(X$, 65, 8))
     Call REPLA(Z$, CENCO$, 193, 8)
       IDEBE# = CVD(Mid$(X$, 49, 8))
       IHABE# = CVD(Mid$(X$, 57, 8))
     If IDEBE# > 0 Then
         Call REPLA(Z$, CSTRING$(MKD$(IDEBE#), 3, 14, 2, 2), 66, 14)
         SUDEBE1# = SUDEBE1# + IDEBE#
       Else
         Call REPLA(Z$, CSTRING$(MKD$(IHABE#), 3, 14, 2, 2), 80, 14)
         SUHABE1# = SUHABE1# + IHABE#
     End If
       TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
     Call REPLA(Z$, CSTRING$(MKD$(TRANSPOR#), 3, 14, 2, 2), 94, 14)
     MAYORPAN.List1.AddItem Z$
   Next U&
   Call CIERRARCH("LIMAYOAC")
   '
   On Error GoTo 0
   If MAYORPAN.List1.ListCount > 0 Then
     MAYORPAN.List1.ListIndex = MAYORPAN.List1.ListCount - 1
   End If
   MAYORPAN.List1.ToolTipText = ""
   '
   MAYORPAN.Ltext2 = CODCUE$(CUEII%)
   MAYORPAN.Label4 = DENOCUE$(CUEII%)
   MAYORPAN.Label21 = TRIM$(FORMATNUM$(SUDEBE1#, "F11.2"))
   MAYORPAN.Label22 = TRIM$(FORMATNUM$(SUHABE1#, "F11.2"))
   Screen.MousePointer = 1
   MAYORPAN.Show 1
   '
End Sub

Sub CARMULTIEM()
   '
   Call BLANARCH("!CARMULTI")
   NX% = FILEOPEN("FLEXMUL.DRV", 0, 79)
   Dim FMX As String * 79
   '
   For i% = 1 To LOF(NX%) / 79
     Get #NX%, i%, FMX$
     If Left$(FMX$, 77) <> String$(77, 0) Then
       SUCAR% = 0
       For J% = 1 To 77
         SUCAR% = SUCAR% + Asc(Mid$(FMX$, J%, 1))
       Next J%
       If CVI(Mid$(FMX$, 78, 2)) = SUCAR% Then
         RASO$ = DESENCRI$(Left$(FMX$, 30))
         PACCE$ = DESENCRI$(Mid$(FMX$, 32, 46))
         Z$ = Space$(96)
         Call REPLA(Z$, RASO$, 1, 30)
         Call REPLA(Z$, PACCE$, 31, 46)
         Call REGAPP("!CARMULTI", Z$)
       End If
     End If
   Next i%
   '
   Close #NX%
   Call CIERRARCH("!CARMULTI")
   '
   VTB% = VENTABU%("CARMULTI")
   If VTB% >= 0 Then
     '
     NX% = FILEOPEN("FLEXMUL.DRV", 0, 79)
     Dim FMY As String * 79
     '
     For U& = 1 To ULTREG&("!CARMULTI")
       Z$ = REGLEIDO$("!CARMULTI", U&)
       RASO$ = AJUSTI$(Left$(Z$, 30), 30)
       PACCE$ = AJUSTI$(Mid$(Z$, 31, 46), 46)
       If TRIM$(RASO$) <> "" Then
         If TRIM$(PACCE$) <> "" Then
           Z1$ = String$(79, 0)
           Call REPLA(Z1$, ENCRIPTA$(RASO$), 1, 30)
           Call REPLA(Z1$, ENCRIPTA$(PACCE$), 32, 46)
           SUCAR% = 0
           For J% = 1 To 77
             SUCAR% = SUCAR% + Asc(Mid$(Z1$, J%, 1))
           Next J%
           Call REPLA(Z1$, MKI$(SUCAR%), 78, 2)
           V& = V& + 1
           LSet FMY$ = Z1$
           Put #NX%, V&, FMY$
         End If
       End If
     Next U&
     '
     For U& = V& + 1 To LOF(NX%) / 79
       LSet FMY$ = String$(79, 0)
       Put #NX%, U&, FMY$
     Next U&
     '
     Close #NX%
     Call CIERRARCH("!CARMULTI")
     '
   End If
   '
End Sub
'

Private Sub tesoreria_Click()
   Call Command13_Click
End Sub
