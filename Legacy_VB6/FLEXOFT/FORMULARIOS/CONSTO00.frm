VERSION 5.00
Begin VB.Form CONSTO00 
   BackColor       =   &H00D0F0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stocks - Existencias y Disponibilidad"
   ClientHeight    =   6285
   ClientLeft      =   1050
   ClientTop       =   1500
   ClientWidth     =   9435
   ClipControls    =   0   'False
   Icon            =   "CONSTO00.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Queries"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   8295
      TabIndex        =   46
      Top             =   2775
      Width           =   1000
      Begin VB.CommandButton COMMAND3 
         Caption         =   "Kardex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   770
         Left            =   105
         Picture         =   "CONSTO00.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   265
         Width           =   750
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Move"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   770
         Left            =   105
         Picture         =   "CONSTO00.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Realizar Movimientos de Stock"
         Top             =   1800
         Width           =   750
      End
      Begin VB.CommandButton Command9 
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
         Height          =   760
         Left            =   105
         Picture         =   "CONSTO00.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Acceso a Consultas y Listados"
         Top             =   1040
         Width           =   750
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2590
      Left            =   8295
      TabIndex        =   42
      Top             =   105
      Width           =   1000
      Begin VB.CommandButton Command27 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   740
         Left            =   105
         Picture         =   "CONSTO00.frx":06E8
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1735
         Width           =   750
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   740
         Left            =   105
         Picture         =   "CONSTO00.frx":0B2A
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1000
         Width           =   750
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   740
         Left            =   105
         Picture         =   "CONSTO00.frx":0E34
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   265
         Width           =   750
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
      TabIndex        =   31
      Top             =   3360
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
         TabIndex        =   50
         Top             =   380
         Width           =   1380
      End
      Begin VB.Label Label404 
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
         Left            =   2900
         TabIndex        =   49
         Top             =   670
         Width           =   1170
      End
      Begin VB.Label Label401 
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
         Left            =   1050
         TabIndex        =   37
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
         TabIndex        =   36
         Top             =   380
         Width           =   1380
      End
      Begin VB.Label Label402 
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
         Left            =   4700
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
         Top             =   380
         Width           =   1170
      End
      Begin VB.Label Label403 
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
         Left            =   6510
         TabIndex        =   32
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
      Height          =   1590
      Left            =   5460
      TabIndex        =   15
      Top             =   1575
      Width           =   2640
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
         Left            =   1260
         TabIndex        =   30
         Top             =   1155
         Width           =   1170
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         TabIndex        =   29
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label301 
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
         Left            =   1260
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   420
         Width           =   855
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
         Left            =   1260
         TabIndex        =   17
         Top             =   735
         Width           =   1170
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
         Left            =   105
         TabIndex        =   16
         Top             =   840
         Width           =   1065
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
      Height          =   1590
      Left            =   2730
      TabIndex        =   10
      Top             =   1575
      Width           =   2640
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
         Left            =   1260
         TabIndex        =   28
         Top             =   1155
         Width           =   1170
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         TabIndex        =   27
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Asignados y Reservas:"
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
         Left            =   105
         TabIndex        =   14
         Top             =   660
         Width           =   1065
      End
      Begin VB.Label Label202 
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
         Left            =   1260
         TabIndex        =   13
         Top             =   735
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
         Left            =   315
         TabIndex        =   12
         Top             =   420
         Width           =   855
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
      Height          =   1590
      Left            =   210
      TabIndex        =   8
      Top             =   1575
      Width           =   2430
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
         Left            =   1050
         TabIndex        =   26
         Top             =   1155
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         TabIndex        =   25
         Top             =   1260
         Width           =   855
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
         Left            =   105
         TabIndex        =   24
         Top             =   840
         Width           =   855
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
         Left            =   1050
         TabIndex        =   23
         Top             =   735
         Width           =   1170
      End
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
         TabIndex        =   22
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label101 
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
      Top             =   4935
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
      Top             =   210
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
      Top             =   4935
      Width           =   3900
   End
   Begin VB.Line Line1 
      X1              =   8100
      X2              =   8100
      Y1              =   4620
      Y2              =   6120
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   7800
      Picture         =   "CONSTO00.frx":0F7E
      Top             =   5610
      Width           =   2010
   End
   Begin VB.Line Line2 
      X1              =   4000
      X2              =   4000
      Y1              =   4620
      Y2              =   6120
   End
   Begin VB.Line Line4 
      X1              =   210
      X2              =   3990
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line3 
      X1              =   4200
      X2              =   9600
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Economico:"
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
      TabIndex        =   41
      Top             =   1075
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
      Left            =   7035
      TabIndex        =   40
      Top             =   1050
      Width           =   1065
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
      TabIndex        =   39
      Top             =   735
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
      TabIndex        =   38
      Top             =   760
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
      TabIndex        =   21
      ToolTipText     =   "Doble-Click para Actualizar Ubicaciones"
      Top             =   4620
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
      TabIndex        =   20
      Top             =   4620
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
      Top             =   315
      Width           =   855
   End
End
Attribute VB_Name = "CONSTO00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CONTROCLIKK%
Dim MOCONPED%

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

Private Sub Command27_Click()
    Command27.Enabled = False
    Call MENSERR(24, "No hay Opciones de Configuración\en la Presente Versión.")
    Command27.Enabled = True
End Sub

Private Sub Command3_Click()
    COMMAND3.Enabled = False
    If NUEWINTA% = 1 Then
        Call CONECRUN("*KARDEX/" + Text1, "Ficha Kardex de Movimientos")
      Else
        Call GENKARDEX(Text1)
    End If
    COMMAND3.Enabled = True
End Sub

Private Sub Command5_Click()
    Call MENSERR(24, "No se Encuentra Ayuda Específica")
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

Private Sub Text1_Change()
    '
    Text2.TEXT = ""
    Text3.TEXT = ""
    Label101.Caption = ""
    Label102.Caption = ""
    Label103.Caption = ""
    Label201.Caption = ""
    Label202.Caption = ""
    Label203.Caption = ""
    Label301.Caption = ""
    Label302.Caption = ""
    Label303.Caption = ""
    Label401.Caption = ""
    Label404.Caption = ""
    Label402.Caption = ""
    Label403.Caption = ""
    List1.Clear
    LIUBIC.Clear
    If CODINT%(Text1.TEXT) > 0 Then
        COD$ = TRIM$(Text1.TEXT)
        CI% = CODINT%(COD$)
        Text2.TEXT = UNIMED$(CI%)
        Text3.TEXT = DESCRIT$(CI%)
        OPENFORMS = DoEvents
        Call MUEDIS(CI%)                ' recalcular existencias
    End If
    '
End Sub
'

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
    CONEXI.Text1.SetFocus
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
     EDITFORM.COMMAND3.Enabled = False
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
     Call VERISAL(CI%)
     Call VESARRAS(CI%)
     '
     Screen.MousePointer = 11
     REG& = CI% + 1
     If REG& > 0 Then
        If REG& <= NUMAS% + 1 Then
           SALCA$ = REGLEIDO$("SALSTOCK", REG&) + Mid$(REGLEIDO$("SALDEPOS", REG&), 41)
           For U& = 1 To NUSU%
               XXZ$ = Mid$(SALCA$, 4 * U& - 3, 4)
               XXX$ = SUCU$(U&) + CSTRING$(XXZ$, 4, 8, 1, 2)
               If DISTO(U&) < 0 Then XXX$ = XXX$ + "(x)"
               If Abs(CVS(XXZ$)) >= 0.05 Then
                  List1.AddItem XXX$
                  CAPAR = CVS(XXZ$)
                  If DISTO%(U&) >= 0 Then
                     STOPROP = STOPROP + CVS(XXZ$)
                  End If
               End If
           Next U&
           GoSub PRESEDIS
       End If
     End If
     '
10   'List2.Clear
     HOII% = HOY(HO$)
     'LI2TEX$ = Space$(80)
     'Call REPLA(LI2TEX$, "Stock Propio al " + HO$, 15, 30)
     'Call REPLA(LI2TEX$, CSTRING$(MKS$(STOPROP), 4, 9, 1, 2), 44, 9)
     'Call REPLA(LI2TEX$, HO$, 65, 9)
     'Call REPLA(LI2TEX$, IDENSEDE$, 75, 8)
     'List2.AddItem LI2TEX$
     SALSTO = STOTOT
     '
12   TORECEP = 0
     RFF$ = RECFILT$("BOLREPEN", 5, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("BOLREPEN", U&)
        CIIX% = CVI(Mid$(X$, 83, 2))
        If CIIX% <> CI% Then
           X$ = REGLEIDO$("BOLREPEN", 1)
           Call GRAREG("BOLREPEN", X$, 1)
           Call CIERRARCH("BOLREPEN")
           GoTo 12
        End If
        '
        CARECEP = CVD(Mid$(X$, 89, 8))
        If CARECEP > 0 Then
           TOSUMIN = TOSUMIN + CARECEP
           TORECEP = TORECEP + CARECEP
        End If
     Next UK&
     '
     If TORECEP > 0.05 Then
        XXX$ = "Recepción no Aprob. " + CSTRING$(MKS$(TORECEP), 4, 8, 1, 2)
        List1.AddItem XXX$
     End If
     '
     GoSub PRESEDIS
     '
     EPAC$ = EMPREAC$
     '
     If MOCONPED% <> 2 Then
       '
       UREOD& = ULTREG&("MAVENDID")
       RGXX$ = RECFILT$("MAVENDID", 20, MKI$(CI%))
       For U& = 1 To Len(RGXX$) Step 4
         REG& = CVS(Mid$(RGXX$, U&, 4))
         If REG& > 0 Then
           If REG& <= UREOD& Then
             X$ = REGLEIDO$("MAVENDID", REG&)
             If Asc(Mid$(X$, 63, 1)) <= 1 Then  ' STATUS 1
               NRODES& = CVS(Mid$(X$, 87, 4))
               NROPED& = CVS(Mid$(X$, 93, 4))
               NCI% = CVI(Mid$(X$, 1, 2))
               CAPEN = CVS(Mid$(X$, 177, 4)) - CVS(Mid$(X$, 225, 4))
               FEX = CVI(Mid$(X$, 125, 2))
               If CAPEN < 0 Then CAPEN = 0
               '
               If CAPEN > 0 Then
                 TORESER = TORESER + CAPEN
               End If
             End If
           End If
         End If
       Next U&
       '
       ABUSQUE$ = "PEDDETP"
14     UREPE& = ULTREG&(ABUSQUE$)
       RGXX$ = RECFILT$(ABUSQUE$, 12, MKI$(CI%))
       For U& = 1 To Len(RGXX$) Step 4
         REG& = CVS(Mid$(RGXX$, U&, 4))
         If REG& > 0 Then
           If REG& <= UREPE& Then
             X$ = REGLEIDO$(ABUSQUE$, REG&)
             '
             If ABUSQUE$ = "PEDDETP" Then
               REGPEDA& = CVS(Mid$(X$, 109, 4))
               If Left$(REGLEIDO$("PEDDETA", REGPEDA&), 108) <> Left$(X$, 108) Then
                 ABUSQUE$ = "PEDDETA"
                 GoTo 14
               End If
             End If
             '
             If Asc(Mid$(X$, 27, 1)) <= 2 Then
               NROPED& = CVS(Left$(X$, 4))
               NCI% = CVI(Mid$(X$, 7, 2))
               CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
               CAPEN = CAPEN - CVS(Mid$(X$, 75, 4)) 'descuenta asignados
               FEX = CVI(Mid$(X$, 25, 2))
               If CAPEN < 0 Then CAPEN = 0
               '
               If CAPEN > 0 Then
                 TOPEDID = TOPEDID + CAPEN
               End If
             End If
           End If
         End If
       Next U&
       '
       GoSub PRESEDIS
       GoTo 15
       '
     End If
     '
15   RFF$ = RECFILT$("OCOMPRA", 6, MKI$(CI%))
     ' esta bien - despues tambien lee de OCOMDETA - 20/11/01
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("OCOMPRA", U&)
        CIIX% = CVI(Mid$(X$, 11, 2))
        If CIIX% <> CI% Then
           X$ = REGLEIDO$("OCOMPRA", 1)
           Call GRAREG("OCOMPRA", X$, 1)
           Call CIERRARCH("OCOMPRA")
           GoTo 15
        End If
        STAT% = Asc(Mid$(X$, 62, 1))
        If STAT% <= 0 Then
           CACOM = CVS(Mid$(X$, 13, 4))
           CAREC = CVS(Mid$(X$, 58, 4))
           CAPEN = CACOM - CAREC
           If CAPEN > 0 Then
              TOSUMIN = TOSUMIN + CAPEN
           End If
        End If
     Next UK&
     '
18   RFF$ = RECFILT$("OCOMDETA", 5, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("OCOMDETA", U&)
        CIIX% = CVI(Mid$(X$, 9, 2))
        If CIIX% <> CI% Then
           X$ = REGLEIDO$("OCOMDETA", 1)
           Call GRAREG("OCOMDETA", X$, 1)
           Call CIERRARCH("OCOMDETA")
           GoTo 18
        End If
        STAT% = Asc(Mid$(X$, 124, 1))
        If STAT% < 9 Then      ' no esta anulada
          NUORSE& = CVS(Mid$(X$, 117, 4))
          If NUORSE& < 1 Then
            COEFCAN = CVS(Mid$(X$, 63, 4))
            If COEFCAN < 0.005 Then COEFCAN = 1
            CACOM = CVS(Mid$(X$, 55, 4)) * COEFCAN
            CAREC = CVS(Mid$(X$, 99, 4)) * COEFCAN
            CAPEN = CACOM - CAREC
            If CAPEN > 0 Then
              TOSUMIN = TOSUMIN + CAPEN
            End If
          End If
        End If
     Next UK&
     '
     GoSub PRESEDIS
     '
25   RFF$ = RECFILT$("ORDEFABR", 5, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("ORDEFABR", U&)
        CIIX% = CVI(Mid$(X$, 83, 2))
        If CIIX% <> CI% Then
           X$ = REGLEIDO$("ORDEFABR", 1)
           Call GRAREG("ORDEFABR", X$, 1)
           Call CIERRARCH("ORDEFABR")
           GoTo 25
        End If
        STAT% = Asc(Mid$(X$, 108, 1))
        If STAT% <= 2 Then
           CACOM = CVD(Mid$(X$, 89, 8))
           CAREC = CVD(Mid$(X$, 119, 8))
           CAPEN = CACOM - CAREC
           If CAPEN > 0 Then
              TOFABRI = TOFABRI + CAPEN
           End If
        End If
     Next UK&
     '
     GoSub PRESEDIS
     '
35   RFF$ = RECFILT$("RESERVA", 1, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("RESERVA", U&)
        CIIX% = CVI(Mid$(X$, 1, 2))
        If CIIX% <> CI% Then
           X$ = REGLEIDO$("RESERVA", 1)
           Call GRAREG("RESERVA", X$, 1)
           Call CIERRARCH("RESERVA")
           GoTo 35
        End If
        TORE# = CVD(Mid$(X$, 19, 8))
        CACO# = CVD(Mid$(X$, 27, 8))
        '
        If TRIM$(UCase$(UNIMED$(CIIX%))) = "GK" Then
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
           NORFA$ = Mid$(X$, 3, 12)
           FEX = CVI(Mid$(X$, 17, 2))
           CICON% = CVI(Mid$(X$, 79, 2))
           If CICON% < 1 Or CICON% > NUMAS% Then
             REBUL& = REGBUS&("ORDEFABR", 3, NORFA$)
             If REBUL& > 0 Then
               XY$ = REGLEIDO$("ORDEFABR", REBUL&)
               CICON% = CVI(Mid$(XY$, 83, 2))
               Call REPLA(X$, MKI$(CICON%), 79, 2)
               Call GRAREG("RESERVA", X$, U&)
             End If
           End If
           TORESER = TORESER + CAPEN
        End If
     Next UK&
     '
     GoSub PRESEDIS
     '
40   RFF$ = RECFILT$("MACONSIG", 8, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("MACONSIG", U&)
        CIIX% = CVI(Mid$(X$, 33, 2))
        If CIIX% <> CI% Then
           X$ = REGLEIDO$("MACONSIG", 1)
           Call GRAREG("MACONSIG", X$, 1)
           Call CIERRARCH("MACONSIG")
           GoTo 40
        End If
        TORE# = CVD(Mid$(X$, 55, 8))
        CACO# = CVD(Mid$(X$, 63, 8))
        '
        If TRIM$(UCase$(UNIMED$(CIIX%))) = "GK" Then
           TORE# = TORE# / 1000
           CACO# = CACO# / 1000
        End If
        '
        SARE = TORE# - CACO#
        If SARE < 0 Then
          SARE = 0
        End If
        CAPEN = SARE
        If Abs(CAPEN) > 0 Then
           NORFA$ = TRIM$(Mid$(X$, 21, 12))
           NUOCO& = Val(Mid$(NORFA$, 4, 6))
           For TYU& = 1 To Len(VECOPEN$) Step 4
             If Abs(CVS(Mid$(VECOPEN$, TYU&, 4)) - NUOCO&) < 0.1 Then GoTo 42
           Next TYU&
           GoTo 44  ' FILTRA PARA O/C's PENDIENTES
           '
42         FEX = CVI(Mid$(X$, 1, 2))
           TORESER = TORESER + CAPEN
        End If
44   Next UK&
     '
     GoSub PRESEDIS
     '
45   RFF$ = RECFILT$("STOREMOT", 1, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("STOREMOT", U&)
        CIIX% = CVI(Mid$(X$, 1, 2))
        If CIIX% <> CI% Then
           X$ = REGLEIDO$("STOREMOT", 1)
           Call GRAREG("STOREMOT", X$, 1)
           Call CIERRARCH("STOREMOT")
           GoTo 45
        End If
        NORFA$ = Mid$(X$, 3, 12)
        TIMO$ = UCase$(Mid$(X$, 15, 2))
        REFE$ = Mid$(X$, 17, 50)
        TORE# = CVD(Mid$(X$, 67, 8))
        CACO# = CVD(Mid$(X$, 75, 8))
        FEX = CVI(Mid$(X$, 83, 2))
        SDOR$ = Mid$(X$, 85, 8)
        CAPEN = TORE# - CACO#
        '
        If Abs(CAPEN) > 0 Then
           If TIMO$ = "ST" Then
               STOREMO = STOREMO + CAPEN
               XXX$ = AJUSTD$(SDOR$, 20) + CSTRING$(MKS$(CAPEN), 4, 8, 1, 2)
               List1.AddItem XXX$
           End If
           If TIMO$ = "PC" Or TIMO$ = CMPED$ Then TOPEDID = TOPEDID - CAPEN
           If TIMO$ = "OF" Then TOFABRI = TOFABRI + CAPEN
           If TIMO$ = "RF" Then TORESER = TORESER - CAPEN
           If TIMO$ = "OC" Then TOSUMIN = TOSUMIN + CAPEN
           If TIMO$ = "BR" Then TOSUMIN = TOSUMIN + CAPEN
        End If
     Next UK&
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
     DUPLI% = 0
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
Return
'
End Sub
'
