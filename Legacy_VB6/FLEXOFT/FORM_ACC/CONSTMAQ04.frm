VERSION 5.00
Begin VB.Form CONSTMAQ04 
   BackColor       =   &H00D0F0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stocks - Existencias y Disponibilidad"
   ClientHeight    =   7395
   ClientLeft      =   1050
   ClientTop       =   1785
   ClientWidth     =   9270
   ClipControls    =   0   'False
   Icon            =   "CONSTMAQ04.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Histórico de Movimientos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2050
      Left            =   120
      TabIndex        =   30
      Top             =   5320
      Width           =   8140
      Begin VB.ListBox List2 
         Height          =   255
         Left            =   2880
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ListBox List5 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   31
         Top             =   675
         Width           =   7875
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha    Doc.Nro.        Referencia             D   Entrada       Salida      Saldo"
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
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   7875
      End
      Begin VB.Line Line6 
         X1              =   120
         X2              =   8000
         Y1              =   1890
         Y2              =   1890
      End
      Begin VB.Line Line5 
         X1              =   8000
         X2              =   8000
         Y1              =   360
         Y2              =   1890
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Situación de Stocks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   210
      TabIndex        =   19
      Top             =   2520
      Width           =   7890
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No Reserv:"
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
         Left            =   1995
         TabIndex        =   29
         Top             =   380
         Width           =   1380
      End
      Begin VB.Label Label404 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2900
         TabIndex        =   28
         Top             =   670
         Width           =   1170
      End
      Begin VB.Label Label401 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1050
         TabIndex        =   25
         Top             =   670
         Width           =   1170
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Neto:"
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
         Left            =   105
         TabIndex        =   24
         Top             =   380
         Width           =   1380
      End
      Begin VB.Label Label402 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4700
         TabIndex        =   23
         Top             =   670
         Width           =   1170
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible:"
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
         Left            =   3885
         TabIndex        =   22
         Top             =   380
         Width           =   1275
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Faltantes:"
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
         Left            =   5670
         TabIndex        =   21
         Top             =   380
         Width           =   1170
      End
      Begin VB.Label Label403 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   6510
         TabIndex        =   20
         Top             =   670
         Width           =   1170
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D0F0F0&
      Caption         =   "A Ingresar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   820
      Left            =   5460
      TabIndex        =   13
      Top             =   1575
      Width           =   2640
      Begin VB.Label Label301 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1260
         TabIndex        =   15
         Top             =   315
         Width           =   1170
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Compras:"
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
         Left            =   315
         TabIndex        =   14
         Top             =   420
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Comprometido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   820
      Left            =   2730
      TabIndex        =   10
      Top             =   1575
      Width           =   2640
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reservas:"
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
         Left            =   115
         TabIndex        =   12
         Top             =   420
         Width           =   1065
      End
      Begin VB.Label Label202 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1260
         TabIndex        =   11
         Top             =   315
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   820
      Left            =   210
      TabIndex        =   8
      Top             =   1575
      Width           =   2430
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Propio:"
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
         TabIndex        =   18
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1050
         TabIndex        =   9
         Top             =   315
         Width           =   1170
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   210
      TabIndex        =   5
      Top             =   4095
      Width           =   3795
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2730
      TabIndex        =   0
      Top             =   210
      Width           =   2430
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7455
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   645
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   735
      Width           =   5160
   End
   Begin VB.CommandButton Command2 
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
      Height          =   360
      Left            =   5145
      TabIndex        =   1
      Top             =   210
      Width           =   175
   End
   Begin VB.ListBox LIUBIC 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   4200
      TabIndex        =   2
      Top             =   4095
      Width           =   3900
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004040&
      Height          =   7800
      Left            =   8390
      ScaleHeight     =   7740
      ScaleWidth      =   1005
      TabIndex        =   33
      Top             =   -120
      Width           =   1065
      Begin VB.CommandButton Command28 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "CONSTMAQ04.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Recepción de Proveedores Externos"
         Top             =   4725
         Width           =   650
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
         Height          =   640
         Left            =   105
         Picture         =   "CONSTMAQ04.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Ficha Kardex - Historico de Movimientos"
         Top             =   2760
         Width           =   650
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "CONSTMAQ04.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Movimientos de Stocks - Consultas y Listados"
         Top             =   3400
         Width           =   650
      End
      Begin VB.CommandButton Command27 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "CONSTMAQ04.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1780
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "CONSTMAQ04.frx":0CEA
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Exit"
         Top             =   210
         Width           =   650
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
         Height          =   630
         Left            =   105
         Picture         =   "CONSTMAQ04.frx":0E34
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   850
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -255
         Picture         =   "CONSTMAQ04.frx":113E
         Stretch         =   -1  'True
         Top             =   6920
         Width           =   1410
      End
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fabricación:"
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
      Left            =   7440
      TabIndex        =   53
      Top             =   8025
      Width           =   1065
   End
   Begin VB.Label Label302 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   8475
      TabIndex        =   52
      Top             =   7920
      Width           =   1170
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Ingresar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7680
      TabIndex        =   51
      Top             =   8445
      Width           =   855
   End
   Begin VB.Label Label303 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   8475
      TabIndex        =   50
      Top             =   8340
      Width           =   1170
   End
   Begin VB.Label Label201 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6345
      TabIndex        =   49
      Top             =   7680
      Width           =   1170
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pedidos:"
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
      Left            =   5400
      TabIndex        =   48
      Top             =   7785
      Width           =   855
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Comprometido:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5040
      TabIndex        =   47
      Top             =   8385
      Width           =   1335
   End
   Begin VB.Label Label203 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6345
      TabIndex        =   46
      Top             =   8520
      Width           =   1170
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Stock:"
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
      Left            =   1200
      TabIndex        =   45
      Top             =   8535
      Width           =   2535
   End
   Begin VB.Label Label103 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3825
      TabIndex        =   44
      Top             =   8430
      Width           =   1170
   End
   Begin VB.Label Label102 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1425
      TabIndex        =   43
      Top             =   8400
      Width           =   1170
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Remoto:"
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
      Left            =   480
      TabIndex        =   42
      Top             =   8505
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Lote Economico"
      Height          =   225
      Left            =   105
      TabIndex        =   41
      Top             =   8085
      Width           =   1275
   End
   Begin VB.Label LECONOMI 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   285
      Left            =   1470
      TabIndex        =   40
      Top             =   8085
      Width           =   1065
   End
   Begin VB.Line Line7 
      X1              =   0
      X2              =   9999
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   8100
      X2              =   8100
      Y1              =   3780
      Y2              =   5280
   End
   Begin VB.Line Line2 
      X1              =   4005
      X2              =   4005
      Y1              =   3780
      Y2              =   5290
   End
   Begin VB.Line Line4 
      X1              =   210
      X2              =   4000
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line3 
      X1              =   4200
      X2              =   8120
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label STOMI 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   285
      Left            =   7035
      TabIndex        =   27
      Top             =   1050
      Width           =   1065
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mínimo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5670
      TabIndex        =   26
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Ubicaciones"
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
      Left            =   4200
      TabIndex        =   17
      ToolTipText     =   "Doble-Click para Actualizar Ubicaciones"
      Top             =   3780
      Width           =   3900
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Existencias"
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
      Left            =   210
      TabIndex        =   16
      Top             =   3780
      Width           =   3795
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Artículo:"
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
      Left            =   420
      TabIndex        =   7
      Top             =   315
      Width           =   2220
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U-M:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6510
      TabIndex        =   6
      Top             =   585
      Width           =   855
   End
   Begin VB.Menu MnuArch 
      Caption         =   "Gestión"
      Begin VB.Menu MnuOpe 
         Caption         =   "Situación de Stocks"
      End
      Begin VB.Menu php0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSal 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu MnuMov 
      Caption         =   "Movimientos"
      Begin VB.Menu MnuReMa 
         Caption         =   "Recepción de Materiales"
      End
      Begin VB.Menu MnuVaCo 
         Caption         =   "Vale de Consumo"
      End
      Begin VB.Menu php2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAju 
         Caption         =   "Ajustes de Stock"
      End
      Begin VB.Menu MnuTra 
         Caption         =   "Transferencias"
      End
   End
   Begin VB.Menu MnuCyL 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu MnuExF 
         Caption         =   "Existencias Actuales"
         Begin VB.Menu MnuPan1 
            Caption         =   "Por Pantalla"
         End
         Begin VB.Menu MnuImp1 
            Caption         =   "Por Impresión"
         End
      End
      Begin VB.Menu MnuExN 
         Caption         =   "Control de Negativos"
         Begin VB.Menu MnuPan2 
            Caption         =   "Por Pantalla"
         End
         Begin VB.Menu MnuImp2 
            Caption         =   "Por Impresión"
         End
      End
      Begin VB.Menu MnuMin 
         Caption         =   "Items Sub-Mínimo"
         Begin VB.Menu MnuPan3 
            Caption         =   "Por Pantalla"
         End
         Begin VB.Menu MnuImp3 
            Caption         =   "Por Impresión"
         End
      End
      Begin VB.Menu MnuGeRe 
         Caption         =   "Situación de Reservas"
         Begin VB.Menu MnuPan0 
            Caption         =   "Por Pantalla"
         End
         Begin VB.Menu MnuImp0 
            Caption         =   "Por Impresión"
         End
      End
   End
   Begin VB.Menu MnuSet 
      Caption         =   "Configuración"
      Begin VB.Menu MnuTab 
         Caption         =   "Tablas"
         Begin VB.Menu MnuDepAlm 
            Caption         =   "Dépositos y Almacenes"
         End
      End
      Begin VB.Menu MnuForImp 
         Caption         =   "Formularios de Impresión"
      End
   End
   Begin VB.Menu MnuAcce 
      Caption         =   "Accesos Directos"
      Begin VB.Menu MnuOtAc 
         Caption         =   "Otros Accesos Directos"
      End
   End
   Begin VB.Menu MnuAyu 
      Caption         =   "Ayuda"
      Begin VB.Menu MnuAyu1 
         Caption         =   "Ayuda FLEXOFT en Línea"
      End
   End
End
Attribute VB_Name = "CONSTMAQ04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CONTROCLIKK%
Dim MOCONPED%
Dim ACUTOT#
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Private Sub Command2_Click()
    '
    ASEL$ = "MASTER"
    '
    Call SELECHO(ASEL$)
    CODD$ = VALACT1$(ASEL$)
    If CODD$ = "" Then
        Exit Sub
    End If
    COD$ = CODD$
    Text1.TEXT = COD$
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = UNIMED$(CI%)
    '
End Sub

Private Sub Command22_Click()
    '
    Command22.Enabled = False
    LIMOSTK04.Show 1
    Command22.Enabled = True
    '
End Sub

Private Sub Command27_Click()
    Command27.Enabled = False
    'Call MENSERR(24, "No hay Opciones de Configuración\en la Presente Versión.")
    OPMOMAQ04.Show 1
    Command27.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    If NUEWINTA% = 1 Then
        Call CONECRUN("*KARDEX/" + Text1, "Ficha Kardex de Movimientos")
      Else
        Call GENKARDEX(Text1)
    End If
    Command3.Enabled = True
End Sub

Private Sub Command28_Click()
   '
   Command28.Enabled = False
     FOREMAQ04.Show 1
   Command28.Enabled = True
   '
End Sub

Private Sub Command4_Click()
  Command4.Enabled = False
    '
    If CODINT%(Text1) < 1 Then
      Call Command2_Click
    End If
    '
    If CODINT%(Text1) > 0 Then
      '
      Call LIKARDEX(Text1.TEXT)
      Call CONECRUN("*COKARDEX/" + Text1.TEXT, "Ficha Kardex Indirecta") ' + "/WAIT")
      '
'      Call GENKARDEX(Text1)
      AppActivate Caption
      SendKeys Chr$(1)
      DoEvents
      DoEvents
      Call FRESHALL
    End If
    '
99  Command4.Enabled = True
End Sub

Private Sub Command5_Click()
  '
  Call HELPONLINE("CONSMAQ")
  '
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   MOVIST00.Show 1
   If CODINT%(Text1.TEXT) > 0 Then
      COD$ = TRIM$(Text1.TEXT)
      CI% = CODINT%(COD$)
      Text2.TEXT = UNIMED$(CI%)
      Text3.TEXT = DESCRIT$(CI%)
      Call MUEDIS(CI%)                ' recalcular existencias
   End If
   Command6.Enabled = True
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    LISTOK00.Show 1
    Command9.Enabled = True
End Sub

Private Sub Form_Activate()
   Call GRATITFOR(Caption)
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
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    '
    MOCONPED% = 1
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
   Hide
End Sub



Private Sub Label13_DblClick()
    CIII% = CODINT%(Text1.TEXT)
    Call CARUBIC(CIII%)
    Screen.MousePointer = 11
    Text1.TEXT = ""
    Text1.TEXT = CODEXT$(CIII%)
    Screen.MousePointer = 1
End Sub

Private Sub Label202_DblClick()
   CONTROCLIKK% = 1
   TOPIN = List2.TopIndex
   For U& = 1 To List2.ListCount
     List2.Selected(U& - 1) = False
     If Left$(List2.List(U& - 1), 2) = "OD" Then
       List2.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   List2.TopIndex = TOPIN
   On Error GoTo 0
   CONTROCLIKK% = 0
End Sub



Private Sub LIUBIC_DblClick()
    CIII% = CODINT%(Text1.TEXT)
    Call CARUBIC(CIII%)
    Screen.MousePointer = 11
    Text1.TEXT = ""
    Text1.TEXT = CODEXT$(CIII%)
    Screen.MousePointer = 1
End Sub

Private Sub MnuAju_Click()
   Call LIMOSTK04.Command25_Click
End Sub

Private Sub MnuAyu1_Click()
  '
  Call Command5_Click
  '
End Sub

Private Sub MnuDepAlm_Click()
  '
  Call OPMOMAQ04.Command16_Click
  '
End Sub

Private Sub MnuForImp_Click()
  '
  Call Command27_Click
  '
End Sub

Private Sub MnuImp0_Click()
  '
  LIMOSTK04.Option4.Value = True
  Call LIMOSTK04.Command14_Click
  '
End Sub

Private Sub MnuImp1_Click()
  '
  LIMOSTK04.Option4.Value = True
  Call LIMOSTK04.Command19_Click
  '
End Sub

Private Sub MnuImp2_Click()
  '
  LIMOSTK04.Option4.Value = True
  Call LIMOSTK04.Command20_Click
  '
End Sub

Private Sub MnuImp3_Click()
  '
  LIMOSTK04.Option4.Value = True
  Call LIMOSTK04.Command30_Click
  '
End Sub
Private Sub MnuOpe_Click()
  '
  Call Command2_Click
  '
End Sub

Private Sub MnuOtAc_Click()
  '
  
  '
End Sub

Private Sub MnuPan0_Click()
  '
  LIMOSTK04.Option5.Value = True
  Call LIMOSTK04.Command14_Click
  '
End Sub

Private Sub MnuPan1_Click()
  '
  LIMOSTK04.Option5.Value = True
  Call LIMOSTK04.Command19_Click
  '
End Sub

Private Sub MnuPan2_Click()
  '
  LIMOSTK04.Option5.Value = True
  Call LIMOSTK04.Command20_Click
  '
End Sub

Private Sub MnuPan3_Click()
  '
  LIMOSTK04.Option5.Value = True
  Call LIMOSTK04.Command30_Click
  '
End Sub

Private Sub MnuReMa_Click()
    Call LIMOSTK04.Command28_Click
End Sub

Private Sub MnuSal_Click()
  '
  Call Command1_Click
  '
End Sub





Private Sub MnuTra_Click()
    Call LIMOSTK04.Command24_Click
End Sub

Private Sub MnuVaCo_Click()
    Call LIMOSTK04.Command21_Click
End Sub

Private Sub Text1_Change()
    '
    Text2.TEXT = ""
    Text3.TEXT = ""
    Label101.Caption = ""
    Label202.Caption = ""
    Label301.Caption = ""
    Label401.Caption = ""
    Label404.Caption = ""
    Label402.Caption = ""
    Label403.Caption = ""
    STOMI = "" 'label de stock minimo
    List1.Clear
    List5.Clear 'lista de mov. kardex
    LIUBIC.Clear
    If CODINT%(Text1.TEXT) > 0 Then
        COD$ = TRIM$(Text1.TEXT)
        CI% = CODINT%(COD$)
        Text2.TEXT = UNIMED$(CI%)
        Text3.TEXT = DESCRIT$(CI%)
        OPENFORMS = DoEvents
        Call MUEDIS(CI%) ' recalcular existencias
        CODT1$ = CODEXT$(CI%)
        Call LIKARDEX(CODT1$)   'CARGA LISTA DE MOVIMIENTO KARDEX
    End If
    '
    Screen.MousePointer = 1
    '
End Sub
'
Sub LIKARDEX(CODL$)
 '
   Static DEPXX%, PRIDEP%
   If PRIDEP% < 1 Then
     Call CARDEPOS
     PRIDEP% = COSU%(1)
   End If
   '
   LU$ = LUDAT$
   If EXISTE%(LU$ + "COKARDEX.RFS") < 1 Then
     Call COPYSTRU("MOVISTO", "COKARDEX")
   End If
   REBLA$ = REGBLAN$("COKARDEX")
   '
   'OBJETO PARA FILTRO DE FECHA, DEPOSITO, MOVIMIENTO VIEJOS
5  CIXX% = CODINT%(CODL$)
  '
  'LIMPIO EL KARDEX ANTERIOR
  List5.Clear
  '
  ACUTOT# = 0
   CIDD% = CIXX%
   JJ& = 0
   If CIDD% > 0 Then
     Call ABRESTOCKS
     SQLX$ = " SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIDD% & " "
     SQLX$ = SQLX$ + " AND CodiMovi <> 'KU' "
'     SQLX$ = SQLX$ + "AND Cdbl(FechMov) >= " & DESDE & " "
'     SQLX$ = SQLX$ + "AND INT(Cdbl(FechMov)) <= " & HASTA & " "
'     If DEPXX% > 0 Then
'       If DEPXX% = PRIDEP% Then
'           SQLX$ = SQLX$ + "AND (DepoMovi = " & DEPXX% & " OR DepoMovi = 0 OR ISNULL(DepoMovi) = True OR ISEMPTY(DepoMovi) = True) "
'         Else
'           SQLX$ = SQLX$ + "AND DepoMovi = " & DEPXX% & " "
'       End If
'     End If
     SQLX$ = SQLX$ + "ORDER BY FechMov ASC, FeReMovi ASC; "
     '
     Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     With MoviTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           YY$ = Space(256)
           '
           Call REPLA(YY$, FECHATEX$(CVINT(!FechMov)), 1, 8)
          ' Call REPLA(YY$, !CodiMovi, 21, 2)
           Call REPLA(YY$, !DoPriCor, 10, 12)
           'Call REPLA(YY$, !DoSecCor, 21, 10)
           Call REPLA(YY$, !ReferCor, 23, 17)
           Call REPLA(YY$, !DepoMovi, 41, 2)
           '
           CAIN = !CantIngre
           CAIN1$ = FORMATNUM$(CAIN, "F8.1")
           CASA = !CantSalid
           CASA1$ = FORMATNUM$(CASA, "F8.1")
           SALD = CAIN - CASAL
           SALD1$ = FORMATNUM$(SALD, "F8.1")
           '
           If CASA > 0 Then
               ACUTOT# = ACUTOT# - CASA
             ElseIf CAIN > 0 Then
               ACUTOT# = ACUTOT# + CAIN
           End If
           ACUTOT1$ = FORMATNUM$(ACUTOT#, "F8.1")
           '
           Call REPLA(YY$, CAIN1$, 44, 9)
           Call REPLA(YY$, CASA1$, 54, 9)
           Call REPLA(YY$, ACUTOT1$, 64, 9)
           '''
           ZZ$ = REBLA$
           Call REPLA(ZZ$, MKI$(CVINT(!FechMov)), 1, 2)
           Call REPLA(ZZ$, MKI$(!Cod_Inte), 3, 2)
           Call REPLA(ZZ$, !IdenLote, 5, 16)
           Call REPLA(ZZ$, !CodiMovi, 21, 2)
           Call REPLA(ZZ$, !DoPriCor, 23, 10)
           Call REPLA(ZZ$, !DoSecCor, 33, 10)
           Call REPLA(ZZ$, !ReferCor, 43, 30)
           Call REPLA(ZZ$, MKD$(!ValoUnit), 73, 8)
             MONEX$ = "$": If !MoneVal = 2 Then MONEX$ = "U$"
           Call REPLA(ZZ$, MONEX$, 81, 2)
           Call REPLA(ZZ$, Chr$(!DepoMovi), 83, 1)
           Call REPLA(ZZ$, MKI$(!Nume_Clie), 84, 2)
           Call REPLA(ZZ$, MKI$(!Nume_Prov), 86, 2)
           Call REPLA(ZZ$, MKD$(!CantIngre), 88, 8)
           Call REPLA(ZZ$, MKD$(!CantSalid), 96, 8)
           Call REPLA(ZZ$, MKD$(0), 104, 8)
           Call REPLA(ZZ$, CVINT(!FeReMovi), 122, 2)
           Call REPLA(ZZ$, Chr$(!NumUsuar), 126, 1)
           JJ& = JJ& + 1
           Call GRAREG("COKARDEX", ZZ$, JJ&)
           '''
           List5.AddItem YY$
         .MoveNext
         Loop
       End If
     End With
     '''
     Call SETULTREG("COKARDEX", JJ&)
     Call CIERRARCH("COKARDEX")
     '''
     Set MoviTemp = Nothing
     Movisto.Close
     '
     End If
End Sub
Private Sub Text1_DblClick()
    '
    ASEL$ = "MASTER"
    Text1.SelLength = 0
    Call SELECHO(ASEL$)
    Text1.SelLength = 0
    CODD$ = VALACT1$(ASEL$)
    If CODD$ = "" Then
        Exit Sub
    End If
    COD$ = CODD$
    Text1.TEXT = COD$
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = UNIMED$(CI%)
    '
End Sub
'

Private Sub Text1_LostFocus()
    COD$ = TRIM$(Text1.TEXT)
    CI% = CODINT%(COD$)
    Text3.TEXT = DESCRIT$(CI%)
    Text2.TEXT = UNIMED$(CI%)
End Sub

Private Sub Text2_GotFocus()
  On Error Resume Next
    Text1.SetFocus
  On Error GoTo 0
End Sub

Private Sub Text3_Click()
    Text1.SetFocus
End Sub

Sub SHOWPEDIPEN()
     '
     Static UFEPEDET#
     If CDbl(FileDateTime(LUDAT$ + "PEDDETA.DAT")) <> UFEPEDET# Then
        UFEPEDET# = CDbl(FileDateTime(LUDAT$ + "PEDDETA.DAT"))
        '
        JJ& = 0
        UREPRE& = ULTREG&("PEDDETA")
        Screen.MousePointer = 11
        JJ& = 0
        NROPEDA& = 0: CAPED% = 0
        FINF = 999999: FSUP = 0
        For REPRESU& = 1 To UREPRE&
          Call TRACE(20, REPRESU&, UREPRE&)
          X$ = REGLEIDO$("PEDDETA", REPRESU&)
          If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
            CAN = CVS(Mid$(X$, 79, 4))
            CADESPACH = CVS(Mid$(X$, 95, 4))
            CAPENDENT = CAN - CADESPACH
            If CAPENDENT > 0.5 Then
              NRODES& = CVS(Left$(X$, 4))
              If NRODES& <> NRODESA& Then
                PRIREDETA& = REPRESU&
                NRODESA& = NRODES&
                TTXX$ = MKS$(NRODES&)
                NOCO$ = TRIM$(Mid$(X$, 9, 16))
                If NOCO$ <> "" Then
                  TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
                End If
                FEX = CVI(Mid$(X$, 5, 2))
                TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
                NCLIE% = CVI(Mid$(X$, 7, 2))
                TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
                REGENCA& = CVS(Mid$(X$, 109, 4))
                JJ& = JJ& + 1
                Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
              End If
            End If
          End If
        Next REPRESU&
        Call SETULTREG("INDIPEP", JJ&)
        Call CIERRARCH("INDIPEP")
        '
        Screen.MousePointer = 1
        YAINDI% = 1
     End If
     '
2112 If ULTREG&("INDIPEP") < 1 Then
        Call COMUNI("No hay O/V's Pendientes\que Puedan Consultarse")
        GoTo 9999
     End If
     '
2110 X$ = REGSEL$("INDIPEP/Ordenes de Venta Pendientes de Entrega")
     If Len(X$) <= 4 Then
        GoTo 9999
     End If
     '
     NPX& = CVS(Left$(X$, 4))
     If NPX& < 1 Then GoTo 9999
     '
     TIDOCUM$ = "Orden de Despacho"
     NUDOCU$ = TRIM$(Str$(NPX&))
     While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
     NUDOCU$ = NUDOCU$ + " "
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
     Exit Sub
     '
5    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Cls
       EDITFORM.Picture1(KKI%).DrawWidth = 3
       EDITFORM.Picture1(KKI%).DrawStyle = 0
       EDITFORM.Picture1(KKI%).FillStyle = 1
     Next KKI%
     '
     REDOCU$ = REGACT$("PDOCLST")
     URE& = ULTREG&("PDOCSPL")
     LI& = 0: LS& = URE&
6    E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
     L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
     DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
     If DCC& < DOCUNU& Then LI& = L&: GoTo 6
     If DCC& > DOCUNU& Then LS& = L&: GoTo 6
     GoTo 8
     '
7    Screen.MousePointer = 1
     Call MENSERR(24, "Documento no Registrado")
     GoTo 9999
     '
8    TPSP$ = ""
     Screen.MousePointer = 11
     While L& > 1
       DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
       If DCC& < DOCUNU& Then GoTo 9
       L& = L& - 1
     Wend
     '
9    CAPAGINAS% = 0
     For KKL& = L& To URE&
       TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
       If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
       If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
         TPSP$ = TPSP$ + Mid$(TBUF$, 5)
       End If
     Next KKL&
     '
95   SPOOLSTRING$ = TPSP$
     For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Cls
       EDITFORM.Picture1(KKI%).Top = 0
       EDITFORM.Picture1(KKI%).Refresh
       EDITFORM.Picture1(KKI%).Height = 7170
       EDITFORM.Picture1(KKI%).Width = 11700
       EDITFORM.Picture1(KKI%).DrawWidth = 3
       EDITFORM.Picture1(KKI%).DrawStyle = 0
       EDITFORM.Picture1(KKI%).FillStyle = 1
     Next KKI%
     ALTUPAGINA = 7170
     TOPEPAGINA = 0
     PAGINACTIVA% = 0
     CAPAGINAS% = 0
     '
10   PPXX% = InStr(TPSP$, Chr$(255))
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
     EDITFORM.Command3.Enabled = False
     EDITFORM.Command1.Enabled = False
     EDITFORM.Show 1
     '
     GoTo 2110
     '
9999 End Sub

Sub MUEDIS(CI%)
     '
     Static CTX%, VECOPEN$, CMPED$
     '
     Screen.MousePointer = 11
     If CTX% <> 1 Then
        IDENSEDE$ = CONTROL$("", "IDENSEDE")
        VECOPEN$ = PUNOCPEN$
        CMPED$ = "PC"
        CMO$ = TRIM$(CONTROL$("", "CODPECLI"))
        If CMO$ <> "" Then CMPED$ = CMO$
        CTX% = 1
     End If
     '
     List1.Clear
     STOPROP = 0
     STOREMO = 0
     TOPEDID = 0
     TOSUMIN = 0
     TOFABRI = 0
     TORESER = 0
     STOMINI = STOMIN(CI%)
     LECONOM = LOTESTAN(CI%)
     GoSub PRESEDIS
     '
     Call CARDEPOS
     '
     Screen.MousePointer = 11
     Call VESARRAS(CI%)
     '
     Screen.MousePointer = 11
     STOPROP = 0
     If CI% > 0 Then
       For U& = 1 To NUSU%
         CAPAR = STODEPOS(CI%, COSU%(U&)) '        XXZ$ = Mid$(SALCA$, 4 * U& - 3, 4)
         If Abs(CAPAR) >= 0.05 Then
           XXX$ = ADEPO$(U&) + CSTRING$(MKS$(CAPAR), 4, 8, 1, 2)
           If DISTO(U&) < 0 Then XXX$ = XXX$ + "(x)"
           List1.AddItem XXX$
           If DISTO%(U&) >= 0 Then
             STOPROP = STOPROP + CAPAR
           End If
         End If
       Next U&
       GoSub PRESEDIS
     End If
     '
10   List2.Clear
     HOII% = HOY(HO$)
     LI2TEX$ = Space$(80)
     Call REPLA(LI2TEX$, "Stock Propio al " + HO$, 15, 30)
     Call REPLA(LI2TEX$, CSTRING$(MKS$(STOPROP), 4, 9, 1, 2), 44, 9)
     Call REPLA(LI2TEX$, HO$, 65, 9)
     Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
     List2.AddItem LI2TEX$
     SALSTO = STOTOT
     '
12   TORECEP = 0
     Call ABRESTOCKS
     SQLX$ = " SELECT * FROM BOLRECEP "
     SQLX$ = SQLX$ + " WHERE Status = 0 "
     SQLX$ = SQLX$ + "AND Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + " ORDER BY NumeBoRe "
     '
     Set COBRECEPTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With COBRECEPTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         JJ& = 0
         K& = 0
         Do While (.EOF = False)
            CARECEP = !CantNom - !CanToApro
            If CARECEP > 0 Then
              NCI% = !NUMPROV
              REMIPRO$ = !NumRem_Prov
              NROBOL& = !NumeBore
              FEX = CVINT(!FechRecep)
              LI2TEX$ = Space$(80)
              Call REPLA(LI2TEX$, "BR." + TRIM$(CSTRING$(MKS$(NROBOL&), 3, 6, 0, 2)), 1, 9)
              Call REPLA(LI2TEX$, "RT." + REMIPRO$ + " - " + RASOCLI$((-1) * NCI%), 15, 30)
              Call REPLA(LI2TEX$, CSTRING$(MKS$(CARECEP), 4, 9, 1, 2), 44, 9)
              Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 9)
              Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
              List2.AddItem LI2TEX$
              TOSUMIN = TOSUMIN + CARECEP
              TORECEP = TORECEP + CARECEP
            End If
         .MoveNext
         Loop
       End If
       On Error GoTo 0
     End With
     '
     If TORECEP > 0.05 Then
        XXX$ = "RECEPC. " + CSTRING$(MKS$(TORECEP), 4, 8, 1, 2)
        List1.AddItem XXX$
     End If
     GoSub PRESEDIS
     '
     Call ABREPEDCLI
     SQLX$ = " SELECT * FROM PEDDETA "
     SQLX$ = SQLX$ + "WHERE CodiInt = " & CI% & " "
     SQLX$ = SQLX$ + "AND Status < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND CantEnt < (CanPed - 0.005) "
     SQLX$ = SQLX$ + "AND VenRep = 'V' "
     SQLX$ = SQLX$ + "ORDER BY NroPed ASC "
     '
     Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     With PEDTEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           NROX& = !NroPed
           NCI% = !NroClie
           CAPEN = !CANPED - !CantEnt
           FEX = CVINT(!FeSolEnt)
           If CAPEN < 0 Then CAPEN = 0
           '
           If CAPEN > 0 Then
             TOPEDID = TOPEDID + CAPEN
             LI2TEX$ = Space$(80)
             Call REPLA(LI2TEX$, CMPED$ + "." + TRIM$(CSTRING$(MKS$(NROX&), 3, 6, 0, 2)), 1, 9)
             Call REPLA(LI2TEX$, RASOCLI$(NCI%), 15, 30)
             Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 53, 9)
             Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 9)
             Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
             List2.AddItem LI2TEX$
           End If
         .MoveNext
         Loop
       End If
     End With
     GoSub PRESEDIS
     GoTo 15
     '
     '
15   Call ABRECOMPRAS
     SQLX$ = " SELECT * FROM OCOMDETA "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + "AND Stat_Ocom < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND  Stat_Ocom <> -1 " 'Para que no  hayan O/C's sin aprobar
     SQLX$ = SQLX$ + "AND Cant_Rec < (Cant_Ocom - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY Nume_Ocom ASC "
     '
     Set OcoTemp = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     With OcoTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           NUORSE& = !N_OrServ
           If NUORSE& < 1 Then
             COEFCAN = !Coef_Unids
             If COEFCAN < 0.005 Then COEFCAN = 1
             CACOM = !Cant_Ocom * COEFCAN
             CAREC = !Cant_rec * COEFCAN
             CAPEN = CACOM - CAREC
             If CAPEN > 0 Then
               NCI% = !Npro_Ocom
               NROX& = !Nume_Ocom
               FEX = CVINT(!Fentre_Sol)
               LI2TEX$ = Space$(80)
               Call REPLA(LI2TEX$, "OC." + TRIM$(CSTRING$(MKS$(NROX&), 3, 6, 0, 2)), 1, 9)
               Call REPLA(LI2TEX$, RASOCLI$((-1) * NCI%), 15, 30)
               Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 44, 9)
               Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 9)
               Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
               List2.AddItem LI2TEX$
               TOSUMIN = TOSUMIN + CAPEN
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     GoSub PRESEDIS
     '
25   Call ABREORFAB
     SQLX$ = " SELECT * FROM ORDEFABR "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + "AND StatuOrf < 3 "   ' para que no este anulada/cumplida/cerrada
     SQLX$ = SQLX$ + "AND CanApro < (CanProy - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY NumeOrFa ASC "
     '
     Set OfaTemp = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     With OfaTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CACOM = !CanProy
           CAREC = !CanApro
           CAPEN = CACOM - CAREC
           If CAPEN > 0 Then
             NroPed& = !NUMEoRFA
             FEX = CVINT(!FechEnSol)
             REFE$ = !Referen
             LI2TEX$ = Space$(80)
             Call REPLA(LI2TEX$, "OF." + TRIM$(CSTRING$(MKS$(NroPed&), 3, 6, 0, 2)), 1, 9)
             Call REPLA(LI2TEX$, REFE$, 15, 30)
             Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 44, 9)
             Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 9)
             Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
             List2.AddItem LI2TEX$
             TOFABRI = TOFABRI + CAPEN
           End If
         .MoveNext
         Loop
       End If
     End With
     GoSub PRESEDIS
     '
35   Call ABRESERVA
     SQLX$ = " SELECT * FROM RESERVA "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + "AND CantCons < (CantRes - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY IdSubOr ASC "
     '
     Set ReserTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     With ReserTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           TORE# = !CantRes
           CACO# = !CantCons
           '
           If TRIM$(UCase$(UNIMED$(CI%))) = "GK" Then
              TORE# = TORE# / 1000
              CACO# = CACO# / 1000
           End If
           '
           SIGNO% = Sgn(CVS(Mid$(X$, 75, 4)))
           SARE = TORE# - CACO#
           If SARE * SIGNO% < 0 Then
             SARE = 0
           End If
           CAPEN = SARE
           If Abs(CAPEN) > 0 Then
             NORFA$ = TRIM$(!IdSubOr)
             FEX = CVINT(!FechRes)
             If Left$(NORFA$, 2) = "OR" Then
               REFE$ = "Reserva p/Reparación"
               GoTo 38
             End If
             '
             CICON% = 0
             Call ABREORFAB
             OrdeFabr.FindFirst "IdSubOr = '" & NORFA$ & "'"
             If OrdeFabr.NoMatch = False Then
               CICON% = OrdeFabr!Cod_Inte
             End If
             DESTIN$ = NORFA$
             If CICON% > 0 Then
               DESTIN$ = TRIM$(CODEXT$(CICON%)) + " - " + TRIM$(DESCRIT0$(CICON%))
             End If
             '
             REFE$ = "Reserva p/" + DESTIN$
             '
38            LI2TEX$ = Space$(80)
             Call REPLA(LI2TEX$, Left$(NORFA$, 12), 1, 12)
             Call REPLA(LI2TEX$, REFE$, 15, 30)
             Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN), 4, 9, 1, 2), 53, 9)
             Call REPLA(LI2TEX$, FECHATEX$(FEX), 65, 9)
             Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
             List2.AddItem LI2TEX$
             TORESER = TORESER + CAPEN
           End If
         .MoveNext
         Loop
       End If
     End With
     '
     GoSub PRESEDIS
     '
20   LIUBIC.Clear
     RGXX$ = RECFILT$("UBISTOCK", 1, MKI$(CI%))
     For U& = 1 To Len(RGXX$) Step 4
        REG& = CVS(Mid$(RGXX$, U&, 4))
        X$ = REGLEIDO$("UBISTOCK", REG&)
        If CVI(Left$(X$, 2)) <> CI% Then
           X$ = REGLEIDO$("UBISTOCK", 1)
           Call GRAREG("UBISTOCK", X$, 1)
           Call CIERRARCH("UBISTOCK")
           GoTo 20
        End If
        '
        If TRIM$(Mid$(X$, 3, 24)) <> "" Then
           Y$ = Mid$(X$, 3, 24)
           LIUBIC.AddItem Y$
        End If
     Next U&
     '
     Screen.MousePointer = 1
     'Call GRAPHDIS
     '
     Exit Sub
     '
PRESEDIS:
     '
     STOMI.Caption = TRIM$(CSTRING$(MKS$(STOMINI), 4, 12, 1, 2))
     LECONOMI.Caption = TRIM$(CSTRING$(MKS$(LECONOM), 4, 12, 1, 2))
     '
     Label101.Caption = TRIM$(CSTRING$(MKS$(STOPROP), 4, 12, 1, 2))
     Label102.Caption = TRIM$(CSTRING$(MKS$(STOREMO), 4, 12, 1, 2))
     STOTOT = STOPROP + STOREMO
     Label103.Caption = TRIM$(CSTRING$(MKS$(STOTOT), 4, 12, 1, 2))
     '
     Label201.Caption = TRIM$(CSTRING$(MKS$(TOPEDID), 4, 12, 1, 2))
     Label202.Caption = TRIM$(CSTRING$(MKS$(TORESER), 4, 12, 1, 2))
     TOCOMPR = TOPEDID + TORESER
     Label203.Caption = TRIM$(CSTRING$(MKS$(TOCOMPR), 4, 12, 1, 2))
     '
     Label301.Caption = TRIM$(CSTRING$(MKS$(TOSUMIN), 4, 12, 1, 2))
     Label302.Caption = TRIM$(CSTRING$(MKS$(TOFABRI), 4, 12, 1, 2))
     TOINPRE = TOSUMIN + TOFABRI
     Label303.Caption = TRIM$(CSTRING$(MKS$(TOINPRE), 4, 12, 1, 2))
     '
     STONET = STOTOT - TOCOMPR
     STODIS = STONET - STOMINI + TOINPRE
     STOFAL = 0
     If STODIS < 0 Then
        STOFAL = Abs(STODIS)
        STODIS = 0
     End If
     '
     Label401.Caption = TRIM$(CSTRING$(MKS$(STONET), 4, 12, 1, 2))
     Label402.Caption = TRIM$(CSTRING$(MKS$(STODIS), 4, 12, 1, 2))
     Label403.Caption = TRIM$(CSTRING$(MKS$(STOFAL), 4, 12, 1, 2))
     Label404.Caption = TRIM$(CSTRING$(MKS$(Val(Label103) - Val(Label202)), 4, 12, 1, 2))
     '
     DoEvents
     '
Return
'
End Sub
'
Private Sub Text3_GotFocus()
  '
  On Error Resume Next
  Text1.SetFocus
  On Error GoTo 0
  '
End Sub
