VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form INICONTA07 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de Contabilidad General"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11880
   Icon            =   "INICONTA07.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11880
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   48
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command10 
      Caption         =   " C.P."
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
      Left            =   7560
      Picture         =   "INICONTA07.frx":1D2A
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Control Presupuestario"
      Top             =   6825
      Width           =   600
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   1680
      TabIndex        =   43
      Top             =   2835
      Visible         =   0   'False
      Width           =   6315
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
      Picture         =   "INICONTA07.frx":1E74
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Informes Standard Normalizados"
      Top             =   4725
      Width           =   560
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
         Height          =   675
         Left            =   105
         Picture         =   "INICONTA07.frx":217E
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   4440
         Width           =   615
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
         Picture         =   "INICONTA07.frx":2488
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Acceso Directo a Aplicaciones Complementarias"
         Top             =   3690
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
         Top             =   2805
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
         Picture         =   "INICONTA07.frx":2792
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
         Picture         =   "INICONTA07.frx":28DC
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2130
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
         Picture         =   "INICONTA07.frx":2BE6
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
         Picture         =   "INICONTA07.frx":2EF0
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
         Picture         =   "INICONTA07.frx":3332
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
         Picture         =   "INICONTA07.frx":363C
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
      Begin VB.ComboBox Combo3 
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
         Left            =   5880
         Sorted          =   -1  'True
         TabIndex        =   44
         Top             =   45
         Visible         =   0   'False
         Width           =   630
      End
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
         Left            =   7755
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
         Left            =   1110
         Sorted          =   -1  'True
         TabIndex        =   32
         Top             =   45
         Width           =   3735
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "U-Neg:"
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
         Left            =   4200
         TabIndex        =   45
         Top             =   90
         Visible         =   0   'False
         Width           =   1575
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
         Left            =   6075
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
         Left            =   -570
         TabIndex        =   34
         Top             =   90
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   1485
      Left            =   1560
      Picture         =   "INICONTA07.frx":3786
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
      Picture         =   "INICONTA07.frx":1BA08
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
      Picture         =   "INICONTA07.frx":1BD12
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
      Picture         =   "INICONTA07.frx":1D344
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
      Picture         =   "INICONTA07.frx":1D48E
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
      Picture         =   "INICONTA07.frx":1D798
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
      Picture         =   "INICONTA07.frx":1DAA2
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
      Picture         =   "INICONTA07.frx":25468
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
      Picture         =   "INICONTA07.frx":2CE2E
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
      Picture         =   "INICONTA07.frx":347F4
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
      Picture         =   "INICONTA07.frx":3C2EE
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
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "INICONTA07.frx":43F80
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "INICONTA07.frx":441B4
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
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
      Begin VB.Menu Uni_Nego 
         Caption         =   "Unidades de Negocios"
         Begin VB.Menu Nue_UNeg 
            Caption         =   "Nueva Unidad de Negocios"
         End
         Begin VB.Menu Mod_UNeg 
            Caption         =   "Modificar Datos"
         End
         Begin VB.Menu Eli_Uneg 
            Caption         =   "Eliminar Unidad de Negocios"
         End
      End
      Begin VB.Menu ph2 
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
      Begin VB.Menu FACPRO 
         Caption         =   "Fac.Proveedores"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "INICONTA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Contablori As Database
Dim Plancuenori As Recordset
Dim NUREPLANORI As Integer
Dim FECHAINI As String
Dim FECHAFIN As String
Dim DENU$

Sub CONTROLPRESU()
    '
    'SELECCION DE MES
    Call CARSELMESA
15  Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    CORRESP$ = MESANO(MESASEL$, DES%, HAS%)
    '
    If CORRESP$ = "" Then Exit Sub
    '
    DES1$ = FETEXCOR1$(DES%)
    HAS1$ = FETEXCOR1$(HAS%)
    'SELECCIONAR OPCION SI INGRESOS O GASTOS
30  OPX% = COMALTER%("Opciones de Trabajo:\\Ingresos\Gastos")
    If OPX% < 1 Or OPX% > 2 Then GoTo 15
    '
    MODOSEL$ = "I"
    If OPX% = 2 Then MODOSEL$ = "G"
    
    'BORRO LOS REGISTROS SEGUN OPCION PARA EL RANGO DE FECHAS CON TIPO 'C'
    CONDI$ = "FECHA >= '" & DES1$ & "'"
    CONDI$ = CONDI$ + " AND FECHA <= '" & HAS1$ & "'"
    CONDI$ = CONDI$ + " AND MARBORRA = 0 "
    CONDI1$ = " AND TIPO = 'C' "
    '
    Call BORRAREGISTROS("PLACOPRE", CONDI$ + CONDI1$, REGAF&, "NOMESS")
    'FIN BORRA REGISTROS CON TIPO C EN RANGO DE FECHAS ELEGIDOS
    '
    'RECORRER TABLA PLACOPRE Y OBTENER DEL CAJABANC Y DETASIEN LOS IMPORTES COINCIDENTES
    'CON CENTRO DE COSTOS Y CUENTA CONTABLE
    SQLX$ = "SELECT * FROM PLACOPRE WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    Dim PLATMP As ADODB.Recordset
    Set PLATMP = READSET(SQLX$)
    '
    With PLATMP
     Do While Not .EOF
        CTAINT% = XVALO(!CueContaI)
        CCSTO$ = TRIM$(SAFETEXT$(!CenCosto))
        FECHAULT = FETEXCOR1$(HAS%)
        CONCEP$ = SAFETEXT$(!CONCEPTO)
        INGR# = XVALO(!INGRESO)
        GAST# = XVALO(!GASTOS)
        COEXCUEN$ = SAFETEXT$(!CODCUENTA)
        CORRESP$ = SAFETEXT$(!MESANO)
        '
        Dim PLATMP1 As ADODB.Recordset
        '
        'GRABACION POR DETALLE DE INGRESOS
        IMPO1# = 0: IMPO2# = 0: IMPOREAL# = 0
        If MODOSEL$ = "I" And INGR# > 0.005 Then ' INGRESOS
           '
           'RECORRE CAJABANC Y ACTUALIZA
           CONDI3$ = " CueContaI = " & CTAINT% & " AND CENCOSTO = '" & CCSTO$ & "' AND FECHEMISI>='" & DES1$ & "' AND FECHEMISI<='" & HAS1$ & "'"
           SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE " & CONDI3$
           Set PLATMP1 = READSET(SQLX$)
           With PLATMP1
           Do While Not .EOF
                IMPO1# = XVALO(!IHABECOMP) - XVALO(!IDEBECOMP)
                FECHAULT = FETEXCOR1$(CVINT(!FechEmisi))
                CONCEP$ = SAFETEXT$(!ObserGen)
                If Abs(IMPO1#) > 0.005 Then
100                 RNDX$ = RANDSTRING$(-8)
                    CONDI$ = "CENCOSTO = '" & RNDX$ & "'"
                    If CantRegistros("PLACOPRE", CONDI$, "NOMESS") > 0 Then GoTo 100
                    '
                    Call CreaRegistro("PLACOPRE", "CENCOSTO", RNDX$, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "CUECONTAI", CONDI$, CTAINT%, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "FECHA", CONDI$, FECHAULT, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "CONCEPTO", CONDI$, CONCEP$, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "INGRESO", CONDI$, 0, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "GASTOS", CONDI$, 0, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "MARBORRA", CONDI$, 0, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "CODCUENTA", CONDI$, COEXCUEN$, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "MESANO", CONDI$, CORRESP$, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "TIPO", CONDI$, "C", REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "Imp_Real", CONDI$, FORMATNUM$(IMPO1#, "F16.5"), REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "Ing_Gas", CONDI$, MODOSEL$, REGAF&, "NOMESS")
                    '
                    Call ACTUREGISTRO("PLACOPRE", "CENCOSTO", CONDI$, CCSTO$, REGAF&, "NOMESS")
                End If
                .MoveNext
           Loop
           End With
            '
           CONDI4$ = " CueCoInt = " & CTAINT% & " AND CenCos = '" & CCSTO$ & "' AND FECHASIEN>='" & DES1$ & "' AND FECHASIEN<='" & HAS1$ & "'"
           SQLX$ = "SELECT * FROM DETAASIEN WITH(NOLOCK) WHERE " & CONDI4$
           Set PLATMP1 = READSET(SQLX$)
           With PLATMP1
           Do While Not .EOF
                NUIDAS = SAFETEXT$(!NUIDASIEN)
                TIPOAS = XVALO(VALOBADA("ENCABEASIEN", "TIPOASIEN", "NUIDASIEN='" & NUIDAS & "'", "NOMESS"))
                IMPO1# = XVALO(!IHabePase) - XVALO(!IDebePase)
                FECHAULT = FETEXCOR1$(CVINT(!FECHASIEN))
                CONCEP$ = SAFETEXT$(!RefePase)
                If Abs(IMPO1#) > 0.005 And TIPOAS < 1 Then 'ENTRA SI EL MOV. TIENE VALOR Y NO ES UN ASIENTO AUTOMATICO
110                 RNDX$ = RANDSTRING$(-8)
                    CONDI$ = "CENCOSTO = '" & RNDX$ & "'"
                    If CantRegistros("PLACOPRE", CONDI$, "NOMESS") > 0 Then GoTo 110
                    '
                    Call CreaRegistro("PLACOPRE", "CENCOSTO", RNDX$, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "CUECONTAI", CONDI$, CTAINT%, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "FECHA", CONDI$, FECHAULT, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "CONCEPTO", CONDI$, CONCEP$, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "INGRESO", CONDI$, 0, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "GASTOS", CONDI$, 0, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "MARBORRA", CONDI$, 0, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "CODCUENTA", CONDI$, COEXCUEN$, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "MESANO", CONDI$, CORRESP$, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "TIPO", CONDI$, "C", REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "Imp_Real", CONDI$, FORMATNUM$(IMPO1#, "F16.5"), REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "Ing_Gas", CONDI$, MODOSEL$, REGAF&, "NOMESS")
                    '
                    Call ACTUREGISTRO("PLACOPRE", "CENCOSTO", CONDI$, CCSTO$, REGAF&, "NOMESS")
                End If
                .MoveNext
           Loop
           End With
        ElseIf MODOSEL$ = "G" And GAST# > 0.005 Then ' GASTOS
           CONDI3$ = " CueContaI = " & CTAINT% & " AND CENCOSTO = '" & CCSTO$ & "' AND FECHEMISI>='" & DES1$ & "' AND FECHEMISI<='" & HAS1$ & "'"
           SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE " & CONDI3$
           Set PLATMP1 = READSET(SQLX$)
           With PLATMP1
           Do While Not .EOF
                IMPO1# = XVALO(!IDEBECOMP) - XVALO(!IHABECOMP)
                FECHAULT = FETEXCOR1$(CVINT(!FechEmisi))
                CONCEP$ = SAFETEXT$(!ObserGen)
                If Abs(IMPO1#) > 0.005 Then
120                RNDX$ = RANDSTRING$(-8)
                    CONDI$ = "CENCOSTO = '" & RNDX$ & "'"
                    If CantRegistros("PLACOPRE", CONDI$, "NOMESS") > 0 Then GoTo 120
                    '
                    Call CreaRegistro("PLACOPRE", "CENCOSTO", RNDX$, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "CUECONTAI", CONDI$, CTAINT%, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "FECHA", CONDI$, FECHAULT, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "CONCEPTO", CONDI$, CONCEP$, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "INGRESO", CONDI$, 0, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "GASTOS", CONDI$, 0, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "MARBORRA", CONDI$, 0, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "CODCUENTA", CONDI$, COEXCUEN$, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "MESANO", CONDI$, CORRESP$, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "TIPO", CONDI$, "C", REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "Imp_Real", CONDI$, FORMATNUM$(IMPO1#, "F16.5"), REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "Ing_Gas", CONDI$, MODOSEL$, REGAF&, "NOMESS")
                    '
                    Call ACTUREGISTRO("PLACOPRE", "CENCOSTO", CONDI$, CCSTO$, REGAF&, "NOMESS")
                End If
                .MoveNext
           Loop
           End With
           CONDI4$ = " CueCoInt = " & CTAINT% & " AND CenCos = '" & CCSTO$ & "' AND FECHASIEN>='" & DES1$ & "' AND FECHASIEN<='" & HAS1$ & "'"
           SQLX$ = "SELECT * FROM DETAASIEN WITH(NOLOCK) WHERE " & CONDI4$
           Set PLATMP1 = READSET(SQLX$)
           With PLATMP1
           Do While Not .EOF
                NUIDAS = SAFETEXT$(!NUIDASIEN)
                TIPOAS = XVALO(VALOBADA("ENCABEASIEN", "TIPOASIEN", "NUIDASIEN='" & NUIDAS & "'", "NOMESS"))
                IMPO1# = XVALO(!IDebePase) - XVALO(!IHabePase)
                FECHAULT = FETEXCOR1$(CVINT(!FECHASIEN))
                CONCEP$ = SAFETEXT$(!RefePase)
                If Abs(IMPO1#) > 0.005 And TIPOAS < 1 Then 'ENTRA SI EL MOV. TIENE VALOR Y NO ES UN ASIENTO AUTOMATICO
130                 RNDX$ = RANDSTRING$(-8)
                    CONDI$ = "CENCOSTO = '" & RNDX$ & "'"
                    If CantRegistros("PLACOPRE", CONDI$, "NOMESS") > 0 Then GoTo 130
                    '
                    Call CreaRegistro("PLACOPRE", "CENCOSTO", RNDX$, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "CUECONTAI", CONDI$, CTAINT%, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "FECHA", CONDI$, FECHAULT, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "CONCEPTO", CONDI$, CONCEP$, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "INGRESO", CONDI$, 0, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "GASTOS", CONDI$, 0, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "MARBORRA", CONDI$, 0, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "CODCUENTA", CONDI$, COEXCUEN$, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "MESANO", CONDI$, CORRESP$, REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "TIPO", CONDI$, "C", REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "Imp_Real", CONDI$, FORMATNUM$(IMPO1#, "F16.5"), REGAF&, "NOMESS")
                    Call ACTUREGISTRO("PLACOPRE", "Ing_Gas", CONDI$, MODOSEL$, REGAF&, "NOMESS")
                    '
                    Call ACTUREGISTRO("PLACOPRE", "CENCOSTO", CONDI$, CCSTO$, REGAF&, "NOMESS")
                End If
                .MoveNext
           Loop
           End With
        End If
        
        '
        .MoveNext
     Loop
    End With
    PLATMP.Close
    Set PLATMP = Nothing
    '
    If MODOSEL$ = "I" Then
        OPX% = COMALTER%("Opciones de Salida:\\Detalle\Totales")
        If OPX% < 1 Or OPX% > 2 Then Exit Sub
        If OPX% = 1 Then Call FINAL("?CONPREING(" & CORRESP$ & ")")
        If OPX% = 2 Then Call FINAL("?CONPREINGT(" & CORRESP$ & ")")
    Else
        OPX% = COMALTER%("Opciones de Salida:\\Detalle\Totales")
        If OPX% < 1 Or OPX% > 2 Then Exit Sub
        If OPX% = 1 Then Call FINAL("?CONPREGAS(" & CORRESP$ & ")")
        If OPX% = 2 Then Call FINAL("?CONPREGAST(" & CORRESP$ & ")")
    End If

End Sub

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
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
   Call ASIDIAR07.Asi_CaBan_Click
   ASIDIAR07.Show 1
End Sub

Private Sub Asi_Caj_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
   Call ASIDIAR07.Asi_Caj_Click
   ASIDIAR07.Show 1
End Sub

Private Sub Asi_Cie_Click()
   '
    ACIERRE07.Text10 = Left$(REGLEIDO$("ECUECOT", 1), 12)
    ACIERRE07.Text2 = Left$(REGLEIDO$("ECUECOT", ULTREG&("ECUECOT")), 12)
    ACIERRE07.Label4.Caption = FECHATEX$(FINEJ%)
    ACIERRE07.Show 1
   '
End Sub

Private Sub Asi_Com_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
   Call ASIDIAR07.Asi_Com_Click
   ASIDIAR07.Show 1
End Sub

Private Sub Asi_Man_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
   Call ASIDIAR07.Asi_Man_Click
   ASIDIAR07.Show 1
End Sub

Private Sub Asi_Pag_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
   Call ASIDIAR07.Asi_Pag_Click
   ASIDIAR07.Show 1
End Sub

Private Sub Asi_Ven_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
   Call ASIDIAR07.Asi_Ven_Click
   ASIDIAR07.Show 1
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
  LUDA1$ = TRIM$(Mid$(List1.List(ICOMBO&), 31)) + "."
    If Mid$(LUDA1$, 2, 1) <> ":" Then LUDA1$ = ULODATOS + LUDA1$
  LUDA2$ = TRIM$(Mid$(List1.List(ICOMBO&), 31)) + "\"
    If Mid$(LUDA2$, 2, 1) <> ":" Then LUDA2$ = ULODATOS + LUDA2$
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
  On Error Resume Next
  FLEXCONN.Close
  On Error GoTo 0
  '
  Call ABRECONEXION
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
  'Set TABEJER = Contable.OpenRecordset("SELECT * FROM TABEJCON WHERE STATUS > 0 ORDER BY FECHAPER DESC;", dbOpenDynaset, dbReadOnly)
  SQLX$ = "SELECT * FROM TABEJCON "
  SQLX$ = SQLX$ + "WHERE STATUS > 0 "
  SQLX$ = SQLX$ + "ORDER BY FECHAPER DESC "
  Dim TABEJER As ADODB.Recordset
  Set TABEJER = FLEXCONN.Execute(FILTSQL$(SQLX$))
  With TABEJER
    On Error Resume Next
    'TABEJER.MoveFirst
    .MoveFirst
    If Err < 1 Then
      On Error GoTo 0
      'With TABEJER
        DENOEJ$ = TRIM$(!Denoejcon)
        COMEJ% = FECHANUM(Format(!FECHAPER, "DD/MM/YY"))
        FINEJ% = FECHANUM(Format(!Fechcierr, "DD/MM/YY"))
        NEJER$ = TRIM$(Str$(!CodiEjer)): While Len(NEJER$) < 2: NEJER$ = "0" + NEJER$: Wend
    End If
    On Error GoTo 0
  End With
  TABEJER.Close
  Set TABEJER = Nothing
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

Private Sub Combo3_Change()
   '
   EPAC1$ = TRIM$(LEERAM$(193, 30))
   UNINEGO% = XVALO(Combo3)
   NUNE% = UNINEGO%
   DENU$ = TRIM$(VALOBADA("UNEGOCIOS", "Deno_Uneg", "Nume_Uneg = " & NUNE%))
   '
   EMPREAC$ = TRIM$(LEERAM$(193, 30)) + " (" + DENU$ + ") "
   If EMPREAC$ <> "" Then EMPREAC$ = EMPREAC$ + " - "
   EMPREAC$ = EMPREAC$ + VERNUM$ + " - " + USERNAME$
   '
   Caption = "Sistema de Contabilidad General - " + EMPREAC$
   '
End Sub

Private Sub Combo3_Click()
   Call Combo3_Change
End Sub

Private Sub Command1_Click()
  Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
  Call CIERRARCH("*.*")
  Call FINAL("")
End Sub

Sub CONTROBADAT()
   Call ABRECABAN
   '
   SQLX$ = "SELECT CodiCom_Lar, CueContaI,iDEBEcOMP,iHABECOMP FROM CAJABANC "
   SQLX$ = SQLX$ + " WHERE (NuAsidiar = '' or NuAsiDiar IS NULL) "
   SQLX$ = SQLX$ + " AND (CueContaI = 455 OR CueContaI = 464 OR CueContaI = 468 OR CueContaI = 469 OR CueContaI = 470)"
   Dim CASEL As ADODB.Recordset
   Set CASEL = New ADODB.Recordset
   CASEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   '
   With CASEL
     On Error Resume Next
     .MoveFirst
     Do While Not .EOF
       On Error Resume Next
       TTXX$ = AJUSTI$(!CodiCom_Lar, 20)
       TTXX$ = TTXX$ + FORMATNUM$(!CueContaI, "F6.0")
       TTXX$ = TTXX$ + FORMATNUM$((!IDEBECOMP - !IHABECOMP), "F12.2")
       SQLY$ = "select FechConta, Nuclien, Nuprove from CAJABANC "
       SQLY$ = SQLY$ + " WHERE CodiCom_Lar = '" & !CodiCom_Lar & "' "
       Dim Busca As ADODB.Recordset
       Set Busca = New ADODB.Recordset
       Set Busca = FLEXCONN.Execute(FILTSQL$(SQLY$))
       TTXX$ = TTXX$ + "   " + FECHATEX$(CVINT(Busca!FECHCONTA))
       TTXX$ = TTXX$ + FORMATNUM$(Busca!NUCLIEN, "F6.0")
       TTXX$ = TTXX$ + FORMATNUM$(Busca!NuProve, "F6.0")
       Busca.Close
       List3.AddItem TTXX$
       List3.Refresh
     .MoveNext
     Loop
   End With
Exit Sub
   
   SQLX$ = "SELECT CodInCue FROM PLANCUEN ORDER BY CODINCUE"
   Dim PLAsel As ADODB.Recordset
   Set PLAsel = New ADODB.Recordset
   PLAsel.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   With PLAsel
     On Error Resume Next
     .MoveFirst
     Do While Not .EOF
       TTXX$ = TTXX$ + FORMATNUM$(!CODINCUE, "I4")
     .MoveNext
     Loop
   End With
   For KKI% = 1 To Len(TTXX$) Step 256
   MsgBox Mid$(TTXX$, KKI%, 256)
   Next KKI%
      
End Sub

Private Sub Command10_Click()
    'RESUMEN
    'PRESENTAR OPCIONES PLANIFICACION Y CONTROL
    'PLANIFICACION:
    '              LEER TABLA PLACOPRE Y PRESENTAR VENTANA PARA EDITAR
    'CONTROL:
    
    Command10.Enabled = False

10  OPX% = COMALTER%("Opciones de Trabajo:\\Planificación\Control")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    '
    If OPX% = 1 Then 'PLANIFICACION
        Call PLANIFICACION
        GoTo 10
    ElseIf OPX% = 2 Then
        Call CONTROLPRESU 'CONTROL
        GoTo 10
    End If
    
99  Command10.Enabled = True
End Sub
Sub PLANIFICACION()
    '
    'SELECCION DE MES
    Call CARSELMESA
15  Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    CORRESP$ = MESANO(MESASEL$, DES%, HAS%)
    '
    If CORRESP$ = "" Then Exit Sub
    '
    DES1$ = FETEXCOR1$(DES%)
    HAS1$ = FETEXCOR1$(HAS%)
    'CONSULTA PARA CARGAR LA VENTANA
    SQLX$ = "SELECT * FROM PLACOPRE WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE FECHA >= '" & DES1$ & "'"
    SQLX$ = SQLX$ + " AND FECHA <= '" & HAS1$ & "'"
    SQLX$ = SQLX$ + " AND MARBORRA = 0 "
    SQLX$ = SQLX$ + " AND TIPO = 'P'"
    Dim PLATMP As ADODB.Recordset
    Set PLATMP = READSET(SQLX$)
    '
    Call SETULTREG("!PLACOPRE", 0)
    With PLATMP
    JJ& = 0
    Do While Not .EOF
       X$ = REGBLAN$("PLACOPRE")
       CCTAI$ = CODCUE$(XVALO(!CueContaI))
       CCSTO$ = TRIM$(SAFETEXT$(!CenCosto))
       FECH% = CVINT(!FECHA)
       CONCEP$ = TRIM$(SAFETEXT$(!CONCEPTO))
       INGR# = XVALO(!INGRESO)
       GAST# = XVALO(!GASTOS)
       
       Call REPLA(X$, CCTAI$, 1, 12)
       Call REPLA(X$, CCSTO$, 13, 8)
       Call REPLA(X$, MKI$(FECH%), 21, 2)
       Call REPLA(X$, CONCEP$, 23, 48)
       Call REPLA(X$, MKD$(INGR#), 71, 8)
       Call REPLA(X$, MKD$(GAST#), 79, 8)
       
       JJ& = JJ& + 1
       Call GRAREG("!PLACOPRE", X$, JJ&)
       .MoveNext
       Loop
    End With
    PLATMP.Close
    Set PLATMP = Nothing
         '

20  VTB% = VENTABU%("PLACOPRE/NC/TITUPAN=Planificación Presupuestaria - Período: " & CORRESP$)
    If VTB% < 0 Then GoTo 15
    '
    For U& = 1 To ULTREG&("!PLACOPRE")
       X$ = REGLEIDO$("!PLACOPRE", U&)
       CCTAI$ = TRIM$(Mid$(X$, 1, 12))
       CCSTO$ = TRIM$(Mid(X$, 13, 8))
       FECH% = CVI(Mid$(X$, 21, 2))
       CONCEP$ = TRIM$(Mid(X$, 23, 48))
       INGR# = CVD(Mid$(X$, 71, 8))
       GAST# = CVD(Mid$(X$, 79, 8))
       '
       'VALIDO
       If CCTAI$ = "" Then
         Call COMUNI("Falta Ingresar Cta. Contable en Fila: " & U&)
         GoTo 20
       End If
       '
       If XVALO(FECH%) < 1 Then
         Call COMUNI("Falta Ingresar Fecha en Fila: " & U&)
         GoTo 20
       End If
       '
       If FECH% < DES% Or FECH% > HAS% Then
         Call COMUNI("Fecha Fuera de Rango en Fila: " & U&)
         GoTo 20
       End If

       If CONCEP$ = "" Then
         Call COMUNI("Falta Ingresar Concepto en Fila: " & U&)
         GoTo 20
       End If
       '
       If INGR# < 0.005 And GAST# < 0.005 Then
          Call COMUNI("Falta Ingresar Importe en Fila: " & U&)
          GoTo 20
       End If
       '
       If INGR# > 0.005 And GAST# > 0.005 Then
          Call COMUNI("Debe Ingresar Ingresos o Gasos en Fila: " & U&)
          GoTo 20
       End If
    Next U&
    'FIN VALIDACIONES
    '
    'MARBORRA = 1 PARA TODOS
    CONDI$ = "FECHA >= '" & DES1$ & "' AND FECHA <= '" & HAS1$ & "'"

    Call ACTUREGISTRO("PLACOPRE", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
    '
    ' RECORRER LA VENTANA Y GRABAR EN B.DATOS
    For JJ& = 1 To ULTREG&("!PLACOPRE")
     'LEO
     X$ = REGLEIDO$("!PLACOPRE", JJ&)
     CCTAI$ = TRIM$(Mid$(X$, 1, 12))
     CCSTO$ = TRIM$(Mid(X$, 13, 8))
     FECH% = CVI(Mid$(X$, 21, 2))
     CONCEP$ = TRIM$(Mid(X$, 23, 48))
     INGR# = CVD(Mid$(X$, 71, 8))
     GAST# = CVD(Mid$(X$, 79, 8))
     CUINT% = CUEINT%(CCTAI$)
     INGR_GAST$ = "I" ' INGRESO
     If GAST# > 0.005 Then INGR_GAST$ = "G" 'GASTO
     '
     'ALTA DE REGISTROS
100  RNDX$ = RANDSTRING$(-8)
     CONDI$ = "CENCOSTO = '" & RNDX$ & "'"
     If CantRegistros("PLACOPRE", CONDI$, "NOMESS") > 0 Then GoTo 100
     '
     Call CreaRegistro("PLACOPRE", "CENCOSTO", RNDX$, "NOMESS")
     Call ACTUREGISTRO("PLACOPRE", "CUECONTAI", CONDI$, CUINT%, REGAF&, "NOMESS")
     Call ACTUREGISTRO("PLACOPRE", "FECHA", CONDI$, FETEXCOR1$(FECH%), REGAF&, "NOMESS")
     Call ACTUREGISTRO("PLACOPRE", "CONCEPTO", CONDI$, CONCEP$, REGAF&, "NOMESS")
     Call ACTUREGISTRO("PLACOPRE", "INGRESO", CONDI$, FORMATNUM$(INGR#, "F16.5"), REGAF&, "NOMESS")
     Call ACTUREGISTRO("PLACOPRE", "GASTOS", CONDI$, FORMATNUM$(GAST#, "F16.5"), REGAF&, "NOMESS")
     Call ACTUREGISTRO("PLACOPRE", "MARBORRA", CONDI$, 0, REGAF&, "NOMESS")
     Call ACTUREGISTRO("PLACOPRE", "CODCUENTA", CONDI$, CCTAI$, REGAF&, "NOMESS")
     Call ACTUREGISTRO("PLACOPRE", "MESANO", CONDI$, CORRESP$, REGAF&, "NOMESS")
     Call ACTUREGISTRO("PLACOPRE", "TIPO", CONDI$, "P", REGAF&, "NOMESS")
     Call ACTUREGISTRO("PLACOPRE", "Ing_Gas", CONDI$, INGR_GAST$, REGAF&, "NOMESS")

     '
     Call ACTUREGISTRO("PLACOPRE", "CENCOSTO", CONDI$, CCSTO$, REGAF&, "NOMESS")
     '
    Next JJ&
    '
    'BORRO LOS MARBORRA = 1 MARCADOS ANTERIORMENTE
    CONDI$ = "FECHA >= '" & DES1$ & "' AND FECHA <= '" & HAS1$ & "' AND MARBORRA = 1"
    Call BORRAREGISTROS("PLACOPRE", CONDI$, REG&)
    '
    Call COMUNI("Datos Ingresados")
    '
End Sub




Sub Command11_Click()
   Command11.Enabled = False
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     GoTo 99
   End If
   '
   ASIDIAR07.Show 1
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
    If DERACCE%("GESVALO", "Tesorería y Gestion de Valores") >= 2 Then
      GESVAL07.Show 1
    End If
   End If
   Command13.Enabled = True
End Sub

Sub Command14_Click()
   Command14.Enabled = False
   INFONOR07.Show 1
   Command14.Enabled = True
End Sub

Sub Command15_Click()
   Command15.Enabled = False
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     GoTo 99
   End If
   '
   OPBALAN07.Show 1
   '
99 Command15.Enabled = True
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   PopupMenu Mnu_Qwery
   Command18.Enabled = True
End Sub

Private Sub Command2_Click()
   If EXISTE%("C:\FLEXOFT\MANUALES\" + App.EXEName + ".PDF") > 0 Then
      On Error Resume Next
      FORHLPDF.Pdf1.LOADFILE "C:\FLEXOFT\MANUALES\" + App.EXEName + ".PDF"
      If Err Then
         On Error GoTo 0
         Call MENSERR(24, "Imposible Abrir Archivo de Ayuda.\  \Falta Instalar Control Pdf1.OCX.")
         Exit Sub
      End If
      FORHLPDF.Show 1
   End If
End Sub

Private Sub Command27_Click()
   Command27.Enabled = False
   OPX% = ALTERNA%("Control de Base de Datos\Otros Accesos Directos")
   If OPX% = 1 Then
        Call CONTROBADAT
      ElseIf OPX% = 2 Then
        Call ACCDIR("CONTAGEN")
   End If
   Command27.Enabled = True
End Sub

 Sub Command3_Click()
    '
    Command3.Enabled = False
    If DERACCE%("MOPLACUE", "Modificar Plan General de Cuentas") < 1 Then
        GoTo 99
    End If
    FPLANCUEN07.Show 1
99  Command3.Enabled = True
    '
End Sub

Sub Command4_Click()
   Command4.Enabled = False
   If DERACCE%("STPCONTA", "Configuración de Funcionamiento de Contabilidad") = 2 Then
      OPASCO07.Show 1
      Call FINAL("")
   End If
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
   'Set Encasien = Contable.OpenRecordset("ENCABEASIEN", dbOpenTable)
   SQLX$ = "DELETE FROM ENCABEASIEN "
     FLEXCONN.Execute (FILTSQL$(SQLX$))
   '
'   On Error Resume Next
'   Encasien.MoveLast
'   CARE% = 0
'   CARE% = Encasien.RecordCount
'   If CARE% > 0 Then
'      While Encasien.RecordCount > 0
'        Encasien.MoveLast
'        Encasien.Delete
'      Wend
'   End If
'
   'Set Detasien = Contable.OpenRecordset("DETAASIEN", dbOpenTable)
   SQLX$ = "DELETE FROM DETAASIEN "
    FLEXCONN.Execute (FILTSQL$(SQLX$))

'5  Detasien.MoveLast
'   CARE% = 0
'   CARE% = Detasien.RecordCount
'   If CARE% > 0 Then
'      While Detasien.RecordCount > 0
'        Detasien.MoveLast
'        Detasien.Delete
'      Wend
'   End If
'   On Error GoTo 0

   '
'6  Encasien.Close
'   Detasien.Close
'   Set Encasien = Contable.OpenRecordset("ENCABEASIEN", dbOpenTable)
   '
   SQLX$ = "SELECT * FROM ENCABEASIEN "
   Dim Encasien As ADODB.Recordset
   Set Encasien = New ADODB.Recordset
   Encasien.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

'   Set Detasien = Contable.OpenRecordset("DETAASIEN", dbOpenTable)
   SQLX$ = "SELECT * FROM DETAASIEN "
   Dim Detasien As ADODB.Recordset
   Set Detasien = New ADODB.Recordset
   Detasien.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText


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
             !RefePase = DETAPAS$
             !FECHASIEN = CVDAT(FECHANUM(FECHAS$))
             !OrFeAsien = ORDAS%
             !SucuAsien = 0
             !NuDfAsien = NUASDEF&
             !StatPase = 1
             !CUECOINT = CUEIII%
             !CUECOEXT = COCUE$
                If !CUECOEXT = "" Then !CUECOEXT = " "
             !IDebePase = IDEBE#
             !IHabePase = IHABE#
             .Update
          End With
          AGREPLAN% = 0
       End If
     Next K2%
     '
35   With Encasien
       '.Edit
       !DebeAsien = SUDEBE#
       .Update
'       .Bookmark = .LastModified
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
   Set Encasien = Nothing
   Detasien.Close
   Set Detasien = Nothing
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
   'Set Plancuen = Contable.OpenRecordset("PLANCUEN", dbOpenDynaset)
   SQLX$ = "SELECT * FROM PLANCUEN"
   Dim PLANCCCUEN As ADODB.Recordset
   Set PLANCCCUEN = New ADODB.Recordset
   PLANCCCUEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   If Err Then
      If COMALTER%("No se Puede Abrir 'PLANCUEN'\en Modo Exclusivo.\\Reintentar\Cancelar") <> 2 Then GoTo 20
      Close: End
   End If
   '
   CAREPLA& = 0
'   'Plancuen.MoveLast
'   .MoveLast
'   CAREPLA& = .RecordCount
'   On Error GoTo 0
'   '
'   UCUE% = 0
'   With Plancuen
'     For U& = 1 To CAREPLA&
'       If U& = 1 Then
'           .MoveFirst
'         Else
'           .MoveNext
'       End If
'       If !CODINCUE > UCUE% Then
'         UCUE% = !CODINCUE
'       End If
'     Next U&
   With PLANCCCUEN
   On Error Resume Next
   .MoveFirst
   If Err < 1 Then
     On Error GoTo 0
     UCUE% = 0
     Do While Not .EOF
       If !CODINCUE > UCUE% Then
         UCUE% = !CODINCUE
       End If
     Loop
   End If
    On Error GoTo 0
    CUEIII% = 1 + UCUE%
     '
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CodiEmpr = CodiEmp%
     On Error GoTo 0
     !CodExCue = COCUE$
     !DENOCUE = DECUE$
     !CODINCUE = CUEIII%
     !STATCUE = 9   ' ANULADA
     !NIVAGRU = 4   ' ULTIMO NIVEL
     !TipoCue = "Debe"
     !POSIMPU = 0   ' NO IMPUTABLE
     .Update
   End With
   PLANCCCUEN.Close
   Set PLANCCCUEN = Nothing
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
   Call COMUNI(" No Disponible en esta Version")
   Exit Sub
   
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
      Load FPLANCUEN07
      If FPLANCUEN07.List1.ListCount > 0 Then
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
      FPLANCUEN07.Show 1
      '
   End If
End Sub

Sub CARPLANORI()
   '
   Screen.MousePointer = 11
   FPLANCUEN07.List1.Clear
   '
   Call ABREPLANCUEN
   '
   FIN& = NUREPLANORI%
   With Plancuenori
     .MoveFirst
     U& = 1
     ZZ$ = Space$(256)
     '
10   COCUE$ = !CodExCue
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
     FPLANCUEN07.List1.AddItem ZZ$
     '
     If U& < FIN& Then
       .MoveNext
       U& = U& + 1
       GoTo 10
     End If
     '
   End With
   '
90 FPLANCUEN07.List1.Refresh
   Screen.MousePointer = 1
   '
End Sub
'

Private Sub Correjer_Click()
   '
   If DERACCE%("ABREJER", "Apertura de Ejercicio Contable") = 2 Then
     '
     If COMEJ% < 1 Or Val(NEJER$) < 0 Then
       Call MENSERR(24, "No Hay Ejercicios Abiertos\en la Empresa Elegida")
       Call Abrejer_Click
       Exit Sub
     End If
     '
     NEJER$ = TRIM$(Mid$(FECHATEX$(COMEJ%), 7, 2))
     NEJU& = XVALO(NEJER$)
     COMEJU = FETEXCOR1$(COMEJ%)
     strsql = "SELECT * FROM TABEJCON WHERE STATUS > 0 "
     strsql = strsql + "AND  FechAper = '" & COMEJU & "' "
     Dim TABEJCON1 As ADODB.Recordset
     Set TABEJCON1 = New ADODB.Recordset
     TABEJCON1.Open FILTSQL$(strsql), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     
     With TABEJCON1
       If (.EOF And .BOF) Then
          Call MENSERR(24, "No Hay Ejercicios Abiertos\en la Empresa Elegida")
          Exit Sub
       End If
       '
15     DENOEJ$ = TRIM$(!Denoejcon)
       COMEJ% = FECHANUM(Format(!FECHAPER, "DD/MM/YY"))
       FINEJ% = FECHANUM(Format(!Fechcierr, "DD/MM/YY"))
       VDEF$ = AJUSTI$(DENOEJ$, 28) + MKI$(COMEJ%) + MKI$(FINEJ%)
       VNUE$ = OBJPLA$("DATEJCONTA", VDEF$)
       If VNUE$ <> "" Then
         DENUE$ = Left$(VNUE$, 28)
         COMEJ% = CVI(Mid$(VNUE$, 29, 2))
         FINEJ% = CVI(Mid$(VNUE$, 31, 2))
         NEJER$ = Mid$(FECHATEX$(COMEJ%), 7, 2)
         If FINEJ% < COMEJ% Then
           Call MENSERR(24, "Fecha de Cierre es Anterior a Apertura.")
           GoTo 15
         End If
         '
         !Denoejcon = DENUE$
         !FECHAPER = CVDAT(COMEJ%)
         !Fechcierr = CVDAT(FINEJ%)
         !CodiEjer = NEJER$
         .Update
         '
         Call CARLISEJER
         Call GRARAM(87, MKI$(COMEJ%))
         Call GRARAM(89, MKI$(FINEJ%))
         Call GRARAM(321, AJUSTI$(DENOEJ$, 48))
         Call GRARAM(369, NEJER$)
         Combo2.TEXT = DENUE$
         '
       End If
     End With
     '
     TABEJCON1.Close
     Set TABEJCON1 = Nothing
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
   Dim MIRS As ADODB.Recordset
   strsql = "SELECT min(CodExCue) as Minimo, "
   strsql = strsql & "max(CodExCue) as Maximo  FROM PLANCUEN "
   Set MIRS = FLEXCONN.Execute(FILTSQL$(strsql))
   With MIRS
    If (.EOF And .BOF) Then
       Exit Sub
    End If
       DIARGEN07.Text1 = !Minimo
       DIARGEN07.Text2 = !maximo
   End With
   MIRS.Close
   Set MIRS = Nothing
   '
   HOII% = HOY(HOS$)
   FEX = COMEJ%: If FEX > HOII% Then FEX = HOII%
   DIARGEN07.Label5 = FECHATEX$(FEX)
   FEX = FINEJ%: If FEX > HOII% Then FEX = HOII%
   DIARGEN07.Label6 = FECHATEX$(FEX)
   Screen.MousePointer = 1
   '
   DIARGEN07.Show 1
   '
End Sub

Private Sub Editplan_Click()
   Call Command3_Click
End Sub

Private Sub Eli_Uneg_Click()
   Call ABRECONEXION
   '
   EPAC1$ = TRIM$(LEERAM$(193, 30))
   NUNE% = XVALO(Combo3)
   DENE$ = VALOBADA("UNEGOCIOS", "Deno_Uneg", "Nume_Uneg = " & NUNE%)
   MENSA$ = "La Eliminación de Una Unidad de Negocios\Puede Dañar el Funcionamiento del Sistema.\   \¿ Está Seguro de Eliminar la Unidad de Negocios\'" + TRIM$(DENE0$) + "' ?"
   If COMALTER%(MENSA$ + "\\No, Cancelar Petición\Si, Eliminar U.N.") = 2 Then
       SQLX$ = "DELETE UNEGOCIOS WHERE Nume_Uneg = " & NUNE%
       FLEXCONN.Execute SQLX$
       Call COMUNI("Eliminada Unidad de Negocios\'" + TRIM$(DENE$) + "'\Bajo el Número " & NUNE% & ".\   \Re-Inicie Ahora la Aplicación.")
       Call FINAL("")
   End If
   '
End Sub

Private Sub Estcuen_Click()
   OPBALAN07.Option1(2).Value = 1
   On Error Resume Next
   OPBALAN07.Command1.SetFocus
   On Error GoTo 0
   OPBALAN07.Show 1
End Sub

Private Sub Exit_Click()
   Call Command1_Click
End Sub

Private Sub Exit1_Click()
   Call Command1_Click
End Sub

Private Sub FacProv_Click()
    FORFAC09.Show 1
End Sub

Private Sub FACPRO_Click()
   FORFAC09.Show 1

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
   Call APLICACIONSKINS(Me)
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
   Screen.MousePointer = 11
   SQLX$ = "UPDATE ENCABEASIEN SET NIV_CLASI = 0 WHERE NIV_CLASI IS NULL"
   SQLY$ = "UPDATE DETAASIEN SET NIV_CLASI = 0 WHERE NIV_CLASI IS NULL"
   SQLZ$ = "UPDATE CAJABANC SET NIV_CLASI = 0 WHERE NIV_CLASI IS NULL"
   SQLT$ = "UPDATE SACREPEN SET NIV_CLASI = 0 WHERE NIV_CLASI IS NULL"
   If CantRegistros("ENCABEASIEN") > 0 Then
      On Error Resume Next
      FLEXCONN.Execute SQLX$
      If Err Then
         On Error GoTo 0
         Call COMUNI("Imposible Completar Comando SQL\" + SQLX$ + "\   \Puede Continuar Ejecución Normalmente.\Se Recomienda Informar a FLEXOFT a la Brevedad.")
      End If
      On Error GoTo 0
   End If
   If CantRegistros("DETAASIEN") > 0 Then
      On Error Resume Next
      FLEXCONN.Execute SQLY$
      If Err Then
         On Error GoTo 0
         Call COMUNI("Imposible Completar Comando SQL\" + SQLY$ + "\   \Puede Continuar Ejecución Normalmente.\Se Recomienda Informar a FLEXOFT a la Brevedad.")
      End If
      On Error GoTo 0
   End If
   If CantRegistros("CAJABANC") > 0 Then
      On Error Resume Next
      FLEXCONN.Execute SQLZ$
      If Err Then
         On Error GoTo 0
         Call COMUNI("Imposible Completar Comando SQL\" + SQLZ$ + "\   \Puede Continuar Ejecución Normalmente.\Se Recomienda Informar a FLEXOFT a la Brevedad.")
      End If
      On Error GoTo 0
   End If
   If CantRegistros("SACREPEN") > 0 Then
      On Error Resume Next
      FLEXCONN.Execute SQLT$
      If Err Then
         On Error GoTo 0
         Call COMUNI("Imposible Completar Comando SQL\" + SQLT$ + "\   \Puede Continuar Ejecución Normalmente.\Se Recomienda Informar a FLEXOFT a la Brevedad.")
      End If
      On Error GoTo 0
   End If
   Screen.MousePointer = 1
   '
   COMEJ% = 0: FINEJ% = 0: DENOEJ$ = "": NEJER$ = ""
   '
   Dim TABEJER As ADODB.Recordset
   strsql = "SELECT * FROM TABEJCON WHERE STATUS > 0 "
   strsql = strsql & "ORDER BY FECHAPER DESC"
   Set TABEJER = FLEXCONN.Execute(FILTSQL$(strsql))
   With TABEJER
     If (.EOF And .BOF) Then
         Call MENSERR(24, "No Hay Ejercicios Abiertos\en la Empresa Elegida")
     Else
       .MoveFirst
       DENOEJ$ = TRIM$(!Denoejcon)
       COMEJ% = FECHANUM(Format(!FECHAPER, "DD/MM/YY"))
       FINEJ% = FECHANUM(Format(!Fechcierr, "DD/MM/YY"))
       NEJER$ = TRIM$(Str$(!CodiEjer)): While Len(NEJER$) < 2: NEJER$ = "0" + NEJER$: Wend
       '.MoveNext
       'Loop
     End If
   End With
   
   TABEJER.Close
   Set TABEJER = Nothing
   
   '
   Call GRARAM(87, MKI$(COMEJ%))
   Call GRARAM(89, MKI$(FINEJ%))
   Call GRARAM(321, AJUSTI$(DENOEJ$, 48))
   Call GRARAM(369, NEJER$)
   Combo2.TEXT = DENOEJ$
   '
   If APLINVO$ = "PLANCUEN" Then
     Call Command3_Click
     Call FINAL("")
   End If
   '
   If COMEJ% < 1 Then
     Call MENSERR(24, "No Hay Ejercicios Abiertos\en la Empresa Elegida")
   End If
   '
   UNINEGO% = 0
   Call CREACAMPO("CONTAGEN", "UNEGOCIOS", "Nume_Uneg", 3, 0)
   Call CREACAMPO("CONTAGEN", "UNEGOCIOS", "Deno_Uneg", 10, 48)
   CANUNE% = CantRegistros&("UNEGOCIOS")
   If CANUNE% > 0 Then
      Combo3.Clear
      SQLX$ = "SELECT Nume_Uneg FROM UNEGOCIOS ORDER BY Nume_Uneg ASC"
      Set UNEGO = FLEXCONN.Execute(FILTSQL$(SQLX$))
      Do While Not UNEGO.EOF
         Combo3.AddItem TRIM$(Str$(UNEGO!Nume_Uneg))
         UNEGO.MoveNext
      Loop
      UNEGO.Close
      Set UNEGO = Nothing
      '
      NUNE% = XVALO(VALOBADA("UNEGOCIOS", "Nume_Uneg", "Nume_Uneg >= 0 ORDER BY Nume_Uneg ASC"))
      Label14.Visible = True
      Combo3.Visible = True
      Combo3.TEXT = TRIM$(Str$(NUNE%))
   End If
   '
   Call CONAPLI(APLINVO$, TERMINA%) ' Busca la aplicacion a mostrar
   If TERMINA% > 0 Then Call FINAL("")

   '
   'UNINEGO% = CodiEmp%
   '
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
   If EXISTE%(App.Path + "\FLEXMUL.DRV") = 1 Then
     LIEMP$ = LOADFILE$(App.Path + "\FLEXMUL.DRV")
     Dim FMX As String * 79
     '
     For I% = 1 To Len(LIEMP$) Step 79
       FMX$ = Mid$(LIEMP$, I%, 79)
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
     Next I%
     '
   End If
   '
End Sub

Sub CARLISEJER()
   '
   Combo2.Clear: List2.Clear
   'Set TABEJER = Contable.OpenRecordset("SELECT * FROM TABEJCON WHERE STATUS > 0 ORDER BY FECHAPER DESC;", dbOpenDynaset, dbReadOnly)
   'jandy sql
   
   'CAEJA% = 0
   'On Error Resume Next
   'TABEJER.MoveLast
   'CAEJA% = TABEJER.RecordCount
   'On Error GoTo 0
   '
   Dim TABEJER As ADODB.Recordset
   strsql = "SELECT * FROM TABEJCON WHERE STATUS > 0 "
   strsql = strsql & "ORDER BY FECHAPER DESC"
   Set TABEJER = FLEXCONN.Execute(FILTSQL$(strsql))
   
   'If CAEJA% > 0 Then
   With TABEJER
    If Not (.EOF And .BOF) Then
       .MoveFirst
       Do While Not .EOF
       'IEJU% = 1
10     DENOEJI$ = TRIM$(!Denoejcon)
       COMEJI% = FECHANUM(Format(!FECHAPER, "DD/MM/YY"))
       FINEJI% = FECHANUM(Format(!Fechcierr, "DD/MM/YY"))
       NEJERI$ = TRIM$(Str$(!CodiEjer)): While Len(NEJERI$) < 2: NEJERI$ = "0" + NEJERI$: Wend
       Combo2.AddItem DENOEJI$
       TTXX$ = Space$(80)
       Call REPLA(TTXX$, DENOEJI$, 1, 48)
       Call REPLA(TTXX$, TRIM$(Str$(COMEJI%)), 49, 8)
       Call REPLA(TTXX$, TRIM$(Str$(FINEJI%)), 57, 8)
       Call REPLA(TTXX$, NEJERI$, 65, 2)
       List2.AddItem TTXX$
       '
       'If IEJU% < CAEJA% Then
       '  IEJU% = IEJU% + 1
         .MoveNext
       '  GoTo 10
       Loop
       End If
     End With
 '  End If
   TABEJER.Close
   Set TABEJER = Nothing
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
     
     'Set Tabejcon = Contable.OpenRecordset("TABEJCON", dbOpenTable)
     SQLX$ = "SELECT * FROM TABEJCON"
     Set TABEJCON1 = New ADODB.Recordset
     TABEJCON1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

     With TABEJCON1
'       UREI% = 0
'       On Error Resume Next
'       .MoveFirst
'       On Error GoTo 0
'       Do While Not .EOF
'
'       UREI% = .RecordCount
'       On Error GoTo 0
       .AddNew   ' YA CONVERTIDO
       On Error Resume Next
       !CodiEmpr = CodiEmp%
       On Error GoTo 0
       !CodiEjer = Mid$(APNU$, 7, 2)
       !Denoejcon = DENUE$
       !FECHAPER = CDate(FEAPNU$)
       !Fechcierr = CDate(FECINU$)
       !Status = 1
       !Diverif = 0
       !Marcacti = 1
       .Update
       '.Bookmark = .LastModified
     End With
     TABEJCON1.Close
     Set TABEJCON1 = Nothing

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
     Call COMUNI("Nuevo Ejercicio O.K.\   \Re-Inicie Ahora la Aplicación.")
     Call FINAL("")
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
   On Error Resume Next
   Unload INIMAY07
   On Error GoTo 0
End Sub

Private Sub Mnu_Import_Click()
   Call Command7_Click
End Sub

Private Sub Mod_UNeg_Click()
   Call ABRECONEXION
   '
   EPAC1$ = TRIM$(LEERAM$(193, 30))
   NUNE% = XVALO(Combo3)
   DENE0$ = VALOBADA("UNEGOCIOS", "Deno_Uneg", "Nume_Uneg = " & NUNE%)
   DENE$ = InputBox$("Denominación de la Unidad de Negocios Nro." & NUNE%, "Nueva Unidad de Negocios - " & EPAC1$, DENE0$)
   If TRIM$(DENE$) <> "" Then
       SQLX$ = "SELECT * FROM UNEGOCIOS WHERE Nume_Uneg = " & NUNE%
       Set UNEGO = New ADODB.Recordset
       UNEGO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
       With UNEGO
          !Deno_Uneg = Left$(DENE$, 48)
          .Update
       End With
       UNEGO.Close
       Set UNEGO = Nothing
       '
       Call COMUNI("Modificada Unidad de Negocios\'" + TRIM$(DENE$) + "'\Bajo el Número " & NUNE% & ".\   \Re-Inicie Ahora la Aplicación.")
       Call FINAL("")
   End If
   '

End Sub

Private Sub Modempre_Click()
   Call CARMULTIEM
End Sub

Private Sub Nue_UNeg_Click()
   Call ABRECONEXION
   '
   EPAC1$ = TRIM$(LEERAM$(193, 30))
   NUNE% = 0
10 If CantRegistros&("UNEGOCIOS", "Nume_Uneg = " & NUNE%) > 0 Then
      NUNE% = NUNE% + 1
      GoTo 10
   End If
   '
   DENE$ = InputBox$("Denominación de la Unidad de Negocios Nro." & NUNE%, "Nueva Unidad de Negocios - " & EPAC1$)
   If TRIM$(DENE$) <> "" Then
       SQLX$ = "SELECT * FROM UNEGOCIOS"
       Set UNEGO = New ADODB.Recordset
       UNEGO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
       With UNEGO
          .AddNew   ' YA CONVERTIDO
          On Error Resume Next
          !CodiEmpr = CodiEmp%
          On Error GoTo 0
          !Nume_Uneg = NUNE%
          !Deno_Uneg = Left$(DENE$, 48)
          .Update
       End With
       UNEGO.Close
       Set UNEGO = Nothing
       '
       Call COMUNI("Creada Unidad de Negocios\'" + TRIM$(DENE$) + "'\Como Número " & NUNE% & ".\   \Re-Inicie Ahora la Aplicación.")
       Call FINAL("")
   End If
   '
End Sub

Private Sub Nuempre_Click()
   NUEMPRE09.Show 1
   'Call CARMULTIEM
End Sub

Private Sub Open_Asi_Click()
   If COMEJ% < 1 Then
     Call MENSERR(24, "Falta Abrir Ejercicio Contable")
     Exit Sub
   End If
   '
   Call ASIDIAR07.Open_Asi_Click
   ASIDIAR07.Show 1
End Sub

Private Sub Resumen_Click()
   OPBALAN07.Option1(1).Value = 1
   On Error Resume Next
   OPBALAN07.Command1.SetFocus
   On Error GoTo 0
   OPBALAN07.Show 1
End Sub

Private Sub Setup_Click()
   Call Command4_Click
End Sub

Private Sub Susaldos_Click()
   OPBALAN07.Option1(4).Value = 1
   On Error Resume Next
   OPBALAN07.Command1.SetFocus
   On Error GoTo 0
   OPBALAN07.Show 1
End Sub

Sub LMAYORPAN(CUEII%)
   '
   FEINI = FETEXCOR1$(COMEJ%)
   FEFIN = FETEXCOR1$(FINEJ%)
   Call BLOQARCH("LIMAYOAC")
   '
   Dim PASEASIEN  As ADODB.Recordset
   Set PASEASIEN = New ADODB.Recordset
   '
   INIMAY07.Ltext2 = CODCUE$(CUEII%)
   INIMAY07.Label4 = DENOCUE$(CUEII%)
   INIMAY07.Label8 = ""
   INIMAY07.Show 1
   If INIMAY07.Label8 <> "OK" Then Exit Sub
   '
   Screen.MousePointer = 11
   SQLY$ = "UPDATE DETAASIEN SET NIV_CLASI = 0 WHERE NIV_CLASI IS NULL and CueCoInt = " & CUEII%
   SQLZ$ = "UPDATE CAJABANC SET NIV_CLASI = 0 WHERE NIV_CLASI IS NULL and CueContaI = " & CUEII%
   On Error Resume Next
   FLEXCONN.Execute SQLY$
   FLEXCONN.Execute SQLZ$
   On Error GoTo 0
   '
   NICONFLIM% = NIVCONFI%
   FEX1% = FECHANUM(INIMAY07.Label5): If FEX1% < COMEJ% Then FEX1% = COMEJ%
   FEX2% = FECHANUM(INIMAY07.Label6): If FEX2% > FINEJ% Then FEX2% = FINEJ%
   FEINI = FETEXCOR1$(FEX1%)
   FEFIN = FETEXCOR1$(FEX2%)
   CONSOLI% = 0
   ' SUCUI% = UNINEGO%
   ' If INIMAY07.Check1.Value = 1 Then CONSOLI% = 1
   PORDETALLE% = 0: If INIMAY07.Option2 = True Then PORDETALLE% = 1
   '
   TRANSPOR# = 0
   SQLX$ = "SELECT sum(idebepase) AS TRADEBE, SUM(ihabepase) AS TRAHABE FROM DetaAsien "
   SQLX$ = SQLX$ + "WHERE (CueCoInt=" & CUEII%
   ' If CONSOLI% < 1 Then SQLX$ = SQLX$ + " AND SucuAsien = " & SUCUI%
   SQLX$ = SQLX$ + " AND Fechasien < '" & FEINI & "'"
   SQLX$ = SQLX$ + " AND statpase > 0 ) "
   SQLX$ = SQLX$ + " AND Niv_Clasi <= " & NICONFLIM%
   Set PASEASIEN = FLEXCONN.Execute(FILTSQL$(SQLX$))
   TRANSPOR# = XVALO(PASEASIEN!TRADEBE) - XVALO(PASEASIEN!TRAHABE)
   PASEASIEN.Close
   '
   Set PASEASIEN = New ADODB.Recordset
   PASEASIEN.CursorLocation = adUseClient
   SQLX$ = "SELECT * FROM DetaAsien "
   SQLX$ = SQLX$ + "INNER JOIN EncabeAsien ON DetaAsien.NuIdAsien = EncabeAsien.NuIdAsien "
   SQLX$ = SQLX$ + "WHERE DetaAsien.CueCoInt=" & CUEII%
   ' If CONSOLI% < 1 Then SQLX$ = SQLX$ + " AND DetaAsien.SucuAsien = " & SUCUI%
   SQLX$ = SQLX$ + " AND DetaAsien.Fechasien >= '" & FEINI & "'"
   SQLX$ = SQLX$ + " AND DetaAsien.Fechasien <= '" & FEFIN & "'"
   SQLX$ = SQLX$ + " AND DetaAsien.statpase > 0  "
   SQLX$ = SQLX$ + " AND DetaAsien.CodiEmpr = " & CodiEmp% & "  "
   SQLX$ = SQLX$ + " AND ENCABEASIEN.statASIEN > 0  "
   SQLX$ = SQLX$ + " AND EncabeAsien.CodiEmpr = " & CodiEmp% & "  "
   SQLX$ = SQLX$ + " AND DetaAsien.Niv_Clasi <= " & NICONFLIM% & " "
   If PORDETALLE% > 0 Then SQLX$ = SQLX$ + "AND ENCABEASIEN.TIPOASIEN < 1 "
   SQLX$ = SQLX$ + "ORDER BY DetaAsien.FECHASIEN, DetaAsien.ORFEASIEN ASC;"
   PASEASIEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   '
   JJ& = 0
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
       IDEBE# = (Int(100 * !IDebePase + 0.5)) / 100
       IHABE# = (Int(100 * !IHabePase + 0.5)) / 100
       If Abs(IDEBE#) < 0.005 And Abs(IHABE#) < 0.005 Then GoTo 19 ' ELIMINA REGISTROS EN BLANCO
       '
       Z$ = REBLA$
       Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 1, 2)
         NUASIX$ = !NUIDASIEN
       Call REPLA(Z$, NUASIX$, 3, 12)
         MACON$ = TRIM$(!MARCONSI)
       Call REPLA(Z$, MACON$, 15, 1)
       Call REPLA(Z$, "D", 16, 1)
         'DEPASE$ = !REFEPASE
         DEPASE = !Refeasien
       Call REPLA(Z$, DEPASE$, 17, 32)
       Call REPLA(Z$, MKD$(IDEBE#), 49, 8)
       Call REPLA(Z$, MKD$(IHABE#), 57, 8)
         CENCO$ = !CenCos
       On Error GoTo 0
       Call REPLA(Z$, CENCO$, 65, 8)
       JJ& = JJ& + 1
       Call GRAREG("LIMAYOAC", Z$, JJ&)
19     On Error GoTo 0
     .MoveNext
     Loop
   End With
   PASEASIEN.Close
   Set PASEASIEN = Nothing
   '
30 On Error GoTo 0
   
   Call ABRECABAN
   Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = New ADODB.Recordset
   '
   SQLX$ = "SELECT sum(IDEBECOMP) AS TRADEBE, SUM(IHABECOMP) AS TRAHABE FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE (CueContaI = " & CUEII%
   SQLX$ = SQLX$ + " AND (NuAsidiar = '' or NuAsiDiar IS NULL)) "
   SQLX$ = SQLX$ + " AND FechConta < '" & FEINI & "' "
   SQLX$ = SQLX$ + " AND Niv_Clasi <= " & NICONFLIM% & " "
   'If CONSOLI% < 1 Then SQLX$ = SQLX$ + " AND CODIEMPR = " & CodiEmp% & " "
   Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   'If Not IsNull(CABATEMP!TRANSPORTE) Then
   TRANSPOR# = TRANSPOR# + XVALO(CABATEMP!TRADEBE) - XVALO(CABATEMP!TRAHABE)
   'End If
   CABATEMP.Close
   '
   Set CABATEMP = New ADODB.Recordset
   CABATEMP.CursorLocation = adUseClient
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where CueContaI = " & CUEII% & " "
   If PORDETALLE% = 0 Then SQLX$ = SQLX$ + " AND (NuAsidiar = '' or NuAsiDiar IS NULL OR NuAsiDiar = '') "
   SQLX$ = SQLX$ + " AND FechConta >= '" & FEINI & "' "
   SQLX$ = SQLX$ + " AND FechConta <= '" & FEFIN & "' "
   SQLX$ = SQLX$ + " AND Niv_Clasi <= " & NICONFLIM% & " "
   'If CONSOLI% < 1 Then SQLX$ = SQLX$ + " AND CODIEMPR = " & SUCUI% & " "
   SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC, FEHOREAL ASC;"
   '
   CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
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
         'If Int(CDbl(!FechConta)) < int(cdbl(FEINI)) Then
         If CVINT(!FECHCONTA) < COMEJ% Then
            TRANSPOR# = TRANSPOR# + IDEBE# - IHABE#
         'ElseIf Int(CDbl(!FechConta)) <= FEFIN Then
         ElseIf CVINT(!FECHCONTA) <= FINEJ% Then
           Z$ = REBLA$
           Call REPLA(Z$, MKI$(CVINT(!FECHCONTA)), 1, 2)
           Call REPLA(Z$, !CodiCom_Cor, 3, 12)
             MACON$ = TRIM$(!MARCONSI)
           Call REPLA(Z$, MACON$, 15, 1)
           Call REPLA(Z$, "C", 16, 1)
             DEPASE$ = !RefeCom
             REFE2$ = TRIM$(SAFETEXT$(!ObserGen))
             If TRIM$(UCase$(DEPASE$)) = TRIM$(UCase$(INIMAY07.Label4)) Then
                If REFE2$ <> "" Then
                  DEPASE$ = REFE2$
                End If
             End If
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
   CABATEMP.Close
   Set CABATEMP = Nothing
   
   '
40 Call SETULTREG("LIMAYOAC", JJ&)
   If InStr(EMPREAC$, "DOSIVA") > 0 Then Call FILESORT("LIMAYOAC", "", 2, 2, "ASC")
   Call FILESORT("LIMAYOAC", "", 1, 1, "ASC")
   '
   MAYORPAN07.List1.Clear
   Z$ = Space$(128 + 64)
     FEX = DIANTE(FEX1%)
     FEXA$ = FECHATEX$(FEX)
   Call REPLA(Z$, FEXA$, 1, 8)
     DEPASE$ = "TRANSPORTE .............................."
   Call REPLA(Z$, DEPASE$, 35, 30)
   Call REPLA(Z$, CSTRING$(MKD$(TRANSPOR#), 3, 14, 2, 2), 94, 14)
   MAYORPAN07.List1.AddItem Z$
   MAYORPAN07.List1.AddItem Space$(95) + " =========== "
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
     MAYORPAN07.List1.AddItem Z$
   Next U&
   Call CIERRARCH("LIMAYOAC")
   '
   On Error GoTo 0
   If MAYORPAN07.List1.ListCount > 0 Then
     MAYORPAN07.List1.ListIndex = MAYORPAN07.List1.ListCount - 1
   End If
   MAYORPAN07.List1.ToolTipText = ""
   '
   MAYORPAN07.Ltext2 = CODCUE$(CUEII%)
   MAYORPAN07.Label4 = DENOCUE$(CUEII%)
   MAYORPAN07.Label7 = INIMAY07.Label7
   MAYORPAN07.Label21 = TRIM$(FORMATNUM$(SUDEBE1#, "F11.2"))
   MAYORPAN07.Label22 = TRIM$(FORMATNUM$(SUHABE1#, "F11.2"))
   Screen.MousePointer = 1
   MAYORPAN07.Show 1
   '
End Sub

Sub CARMULTIEM()
   '
   Call BLANARCH("!CARMULTI")
   NX% = FILEOPEN(App.Path + "\FLEXMUL.DRV", 0, 79)
   Dim FMX As String * 79
   '
   For I% = 1 To LOF(NX%) / 79
     Get #NX%, I%, FMX$
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
   Next I%
   '
   Close #NX%
   Call CIERRARCH("!CARMULTI")
   '
   VTB% = VENTABU%("CARMULTI")
   If VTB% >= 0 Then
     '
     NX% = FILEOPEN(App.Path + "\FLEXMUL.DRV", 0, 79)
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

Sub CONAPLI(APLINVO$, TERMINA%)
   
   CODCARPE$ = TRIM$(BUSCAVALOR(PARAPLI$, "CODCARPE =")) ' codigo de carpeta de import / export
   CODCENCO$ = COCENCOS$(CODCARPE$)
   '
'''   ' LO QUE SIGUE SE TRASLADO A CASE ASIEDIAR POR PROLIJIDAD. ADEMÁS VER LOS ERRORES SEÑALADOS
'''   ' LECTURA DEL CENTRO DE COSTOS PASADO DESDE CARPETA DE IMPORTACION
'''   If Mid$(APLINVO$, 1, 8) = "ASIEDIAR" Then
'''        If CODCARPE$ <> "" Then
'''           CONDI$ = "CodCenCos ='" & CODCARPE$ & "' AND STATUS < 9"   ' LA CONDICION ESTA MAL
'''           If CantRegistros("CARPEIMP", CONDI$) > 0 Then
'''                NPRO$ = VALOBADA("CARPEIMP", "CodCenCos", CONDI$, "NOMESS")
'''                CCIMP$ = VALOBADA("CARPEIMP", "CodCarpImp", CONDI$, "NOMESS")   ' ESTO ES UNA TAUTOLOGIA
'''                Label15 = CCIMP$                  ' ESTO NO HACE FALTA  ... ES CODCARPE$
'''                ASIDIAR07.Text111 = NPRO$
'''                Call Command11_Click
'''                If Label15 <> "" Then
'''                   Command99.Enabled = False      ' FALTA ASIDIAR07.COMMAND99
'''                   Text111.Enabled = False        ' FALTA ASIDIAR09.TEXT111
'''                End If
'''                Call FINAL("")
'''            End If
'''        End If
'''   End If
'''   '
   TERMINA% = 1
   Select Case APLINVO$
       '
       'TRANSACCIONES
       '
       Case "ASIEDIAR"
           Label15 = CODCARPE$
           If CODCARPE$ <> "" Then
              CONDI$ = "CodCarpImp ='" & CODCARPE$ & "' AND STATUS < 9"
              CODCENCO$ = VALOBADA("CARPEIMP", "CodCenCos", CONDI$, "NOMESS")
              ASIDIAR07.Text111 = CODCENCO$
              ASIDIAR07.Text111.Enabled = False      ' PARA QUE NO SE PUEDA CAMBIAR
              ASIDIAR07.Command99.Enabled = False
           End If
           Call Command11_Click                      ' ASIENTO DE DIARIO
           If CODCARPE$ <> "" Then
              Call FINAL("")                         ' PARA QUE CIERRE Y SALGA DE LA APLICACION
           End If
       Case "ASIEREVEN"
           Call ASIDIAR07.Asi_Ven_Click              ' ASIENTO RESUMEN VENTAS
           Call Command11_Click
       Case "ASIERECOM"
           Call ASIDIAR07.Asi_Com_Click              ' ASIENTO RESUMEN COMPRAS
           Call Command11_Click
       Case "ASIERECOB"
           Call ASIDIAR07.Asi_Caj_Click              ' ASIENTO RESUMEN COBRANZAS
           Call Command11_Click
       Case "ASIEREPAG"
           Call ASIDIAR07.Asi_Pag_Click              ' ASIENTO RESUMEN PAGOS
           Call Command11_Click
       Case "ASIECAJBANC"
           Call ASIDIAR07.Asi_CaBan_Click            ' ASIENTO CAJA Y BANCOS
           Call Command11_Click
       Case "ASIEAPER"                               ' ESTA COMENTADO PORQUE FALTA AGREGAR LA OPCION AL MENU
           Call ASIDIAR07.Asi_Aper_Click             ' ASIENTO DE APERTURA
           Call Command11_Click
       Case "ASIECIER"
           Call ASIDIAR07.Asi_Cie_Click              ' ASIENTO DE CIERRE
       Case "BUASIEXIS"
           Call ASIDIAR07.Command25_Click            ' BUSCA ASIENTOS ACTIVOS
           Call Command11_Click
       Case "PLANCUE"
           Call Command3_Click                       ' PLAN GENERAL DE CUENTAS
       '
       'CONSULTAS Y LISTADOS
       '
       Case "DIAGENER"
           Call Diarge_Click                         ' DIARIO GENERAL
       Case "MAYPORCUE"
           Call Mayor_Click                          ' MAYOR POR CUENTA
       Case "BALSUMSAL"
           Call Command15_Click                      ' BALANCE DE SUMAS Y SALDOS
       Case "INFONORM"
           Call Command14_Click                      ' INFORMES NORMALIZADOS
       Case "IMPLACUE"
           Call FPLANCUEN07.Command6_Click           ' IMPRIME PLAN CUENTAS
       '
       'CONFIGURACION CONTABILIDAD
       '
       Case "CONGENCONT"
           Call Command4_Click                       ' CONFIGURACION GENERAL DE CONTABILIDAD
      
       Case Else
           TERMINA% = 0
   End Select
End Sub
