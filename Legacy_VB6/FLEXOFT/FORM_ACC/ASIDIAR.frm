VERSION 5.00
Begin VB.Form ASIDIAR 
   BackColor       =   &H00E0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asientos Contables"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture16 
      Height          =   8300
      Left            =   10800
      ScaleHeight     =   8235
      ScaleWidth      =   1035
      TabIndex        =   68
      Top             =   -240
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   8400
         Left            =   120
         Picture         =   "ASIDIAR.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command29 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9500
      TabIndex        =   2
      Top             =   720
      Width           =   255
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00E0E0C0&
      Caption         =   "Auxil."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   9900
      TabIndex        =   57
      Top             =   4080
      Width           =   1020
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
         Height          =   580
         Left            =   180
         Picture         =   "ASIDIAR.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   360
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         Height          =   560
         Left            =   180
         Picture         =   "ASIDIAR.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Activa Calculadora del Sistema"
         Top             =   960
         Width           =   650
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00E0E0C0&
      Caption         =   "A-B-M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3850
      Left            =   9900
      TabIndex        =   50
      Top             =   120
      Width           =   1020
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
         Height          =   560
         Left            =   180
         Picture         =   "ASIDIAR.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Imprime Comprobante del Asiento"
         Top             =   3160
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   180
         Picture         =   "ASIDIAR.frx":3016
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Restaura Valores Anteriores del Asiento"
         Top             =   2600
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
         Height          =   560
         Left            =   180
         Picture         =   "ASIDIAR.frx":3320
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Aprueba y Graba el Asiento"
         Top             =   2040
         Width           =   650
      End
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
         Height          =   560
         Left            =   180
         Picture         =   "ASIDIAR.frx":362A
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Borra Asiento Activo"
         Top             =   1480
         Width           =   650
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
         Height          =   560
         Left            =   180
         Picture         =   "ASIDIAR.frx":3934
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Abre Asiento Existente"
         Top             =   920
         Width           =   650
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
         Height          =   560
         Left            =   180
         Picture         =   "ASIDIAR.frx":3C3E
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Nuevo Asiento Contable"
         Top             =   360
         Width           =   650
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Frame10"
      Height          =   3975
      Left            =   6100
      TabIndex        =   47
      Top             =   1300
      Width           =   15
   End
   Begin VB.Frame Frame9 
      Caption         =   "Frame9"
      Height          =   3975
      Left            =   9480
      TabIndex        =   46
      Top             =   1300
      Width           =   15
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   3975
      Left            =   7780
      TabIndex        =   45
      Top             =   1300
      Width           =   15
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3975
      Left            =   1640
      TabIndex        =   44
      Top             =   1300
      Width           =   15
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   5340
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "ASIDIAR.frx":3D88
      Top             =   720
      Width           =   4110
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00800000&
      Height          =   8600
      Left            =   11040
      ScaleHeight     =   8535
      ScaleWidth      =   1215
      TabIndex        =   32
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   120
         Picture         =   "ASIDIAR.frx":3D8A
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Restaura Asiento Anulado"
         Top             =   1680
         Width           =   615
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   555
         TabIndex        =   60
         Top             =   4560
         Width           =   615
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   61
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command10 
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
         Picture         =   "ASIDIAR.frx":4094
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Plan General de Cuentas"
         Top             =   5040
         Width           =   615
      End
      Begin VB.CommandButton Command19 
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
         Picture         =   "ASIDIAR.frx":41DE
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command18 
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
         Height          =   675
         Left            =   120
         Picture         =   "ASIDIAR.frx":4328
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command17 
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
         Picture         =   "ASIDIAR.frx":4632
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   6480
         Width           =   615
      End
      Begin VB.CommandButton Command16 
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
         Left            =   120
         Picture         =   "ASIDIAR.frx":4A74
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5760
         Width           =   615
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Util."
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
         Picture         =   "ASIDIAR.frx":4D7E
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   7200
         Width           =   615
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   310
      Left            =   9540
      TabIndex        =   29
      Top             =   210
      Value           =   16384
      Width           =   175
   End
   Begin VB.CommandButton Command21 
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
      Height          =   300
      Left            =   6480
      TabIndex        =   0
      Top             =   210
      Width           =   175
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0C0&
      Caption         =   "Registrar Pase - Movimiento"
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
      Left            =   120
      TabIndex        =   16
      Top             =   5790
      Width           =   10800
      Begin VB.CommandButton Command99 
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
         Left            =   2205
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   1575
         Width           =   175
      End
      Begin VB.TextBox Text111 
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
         Left            =   1125
         TabIndex        =   5
         Top             =   1575
         Width           =   1080
      End
      Begin VB.TextBox Text12 
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
         Left            =   1125
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1155
         Width           =   4365
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Ver Mayor"
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
         Left            =   3960
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "Muestra el Mayor de la Cuenta Activa"
         Top             =   315
         Width           =   1500
      End
      Begin VB.CommandButton Command11 
         Caption         =   "App"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   9840
         Picture         =   "ASIDIAR.frx":5088
         TabIndex        =   10
         ToolTipText     =   "Agrega Pase al Final del Asiento"
         Top             =   300
         Width           =   765
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Insert"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   9840
         Picture         =   "ASIDIAR.frx":5392
         TabIndex        =   11
         ToolTipText     =   "Intercala Pase Encima del Activo"
         Top             =   560
         Width           =   765
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Del"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   9840
         Picture         =   "ASIDIAR.frx":569C
         TabIndex        =   12
         ToolTipText     =   "Elimina el Pase Activo del Asiento"
         Top             =   810
         Width           =   765
      End
      Begin VB.CommandButton Command8 
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
         Height          =   250
         Left            =   9840
         Picture         =   "ASIDIAR.frx":59A6
         TabIndex        =   13
         ToolTipText     =   "Eldición el Pase Activo para su Modificación"
         Top             =   1070
         Width           =   765
      End
      Begin VB.CommandButton Command6 
         Height          =   585
         Left            =   9840
         Picture         =   "ASIDIAR.frx":5CB0
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Agrega Pase al Asiento"
         Top             =   1400
         Width           =   765
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Balance / Diferencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   6015
         TabIndex        =   22
         Top             =   1150
         Width           =   3440
         Begin VB.Label Label10 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1800
            TabIndex        =   24
            ToolTipText     =   "Doble Click para Balancear Asiento"
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label8 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   105
            TabIndex        =   23
            ToolTipText     =   "Doble Click para Balancear Asiento"
            Top             =   360
            Width           =   1500
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Haber"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   7740
         TabIndex        =   21
         Top             =   210
         Width           =   1710
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   105
            MaxLength       =   11
            TabIndex        =   7
            Text            =   " "
            Top             =   360
            Width           =   1500
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Debe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   6020
         TabIndex        =   20
         Top             =   210
         Width           =   1710
         Begin VB.TextBox Text14 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   105
            MaxLength       =   11
            TabIndex        =   6
            Top             =   360
            Width           =   1500
         End
      End
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
         Height          =   285
         Left            =   2910
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text11 
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
         Left            =   1125
         TabIndex        =   3
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label Label13 
         Height          =   330
         Left            =   5565
         TabIndex        =   66
         Top             =   1575
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label Label12 
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
         Height          =   285
         Left            =   2415
         TabIndex        =   65
         Top             =   1575
         Width           =   3060
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C. Costos:"
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
         Left            =   75
         TabIndex        =   64
         Top             =   1635
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle:"
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
         Left            =   75
         TabIndex        =   62
         Top             =   1220
         Width           =   960
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   9630
         X2              =   9630
         Y1              =   120
         Y2              =   2140
      End
      Begin VB.Label Label4 
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
         Height          =   285
         Left            =   210
         TabIndex        =   19
         Top             =   750
         Width           =   5280
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         Left            =   75
         TabIndex        =   17
         Top             =   375
         Width           =   960
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   4545
      Left            =   120
      TabIndex        =   15
      Top             =   1400
      Width           =   9735
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   0
         ScaleHeight     =   390
         ScaleWidth      =   9570
         TabIndex        =   39
         Top             =   0
         Width           =   9630
         Begin VB.Label Label27 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cuenta              Detalle / Referencia                                                       Debe                 Haber"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   30
            TabIndex        =   40
            Top             =   105
            Width           =   11355
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
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
         Height          =   3405
         Left            =   0
         MultiSelect     =   2  'Extended
         TabIndex        =   38
         Top             =   450
         Width           =   9650
      End
      Begin VB.PictureBox Picture1 
         Height          =   460
         Left            =   0
         ScaleHeight     =   405
         ScaleWidth      =   9570
         TabIndex        =   26
         Top             =   3840
         Width           =   9630
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7770
            TabIndex        =   28
            Top             =   60
            Width           =   1500
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6090
            TabIndex        =   27
            Top             =   60
            Width           =   1500
         End
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9630
         Y1              =   4300
         Y2              =   4300
      End
      Begin VB.Line Line2 
         X1              =   9630
         X2              =   9630
         Y1              =   0
         Y2              =   4320
      End
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   31
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   11100
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Ltext2 
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8200
      TabIndex        =   43
      Top             =   220
      Width           =   1350
   End
   Begin VB.Label Ltext1 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5370
      TabIndex        =   42
      Top             =   220
      Width           =   1110
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Referencia:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3420
      TabIndex        =   41
      Top             =   765
      Width           =   1875
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
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
      Left            =   -120
      TabIndex        =   30
      Top             =   960
      Width           =   4110
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Asiento de Diario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   540
      TabIndex        =   25
      Top             =   240
      Width           =   5475
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Número:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6765
      TabIndex        =   18
      Top             =   270
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3720
      TabIndex        =   14
      Top             =   270
      Width           =   1275
   End
   Begin VB.Menu arch_01 
      Caption         =   "Asientos Contables"
      Begin VB.Menu nu_asi 
         Caption         =   "Nuevo Asiento"
         Begin VB.Menu Asi_man 
            Caption         =   "Asiento Manual"
         End
         Begin VB.Menu li3 
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
            Caption         =   "Resumen de Caja y Bancos "
         End
         Begin VB.Menu li32 
            Caption         =   "-"
         End
         Begin VB.Menu Asi_Aper 
            Caption         =   "Asiento de Apertura"
         End
         Begin VB.Menu li33 
            Caption         =   "-"
         End
         Begin VB.Menu Asi_Cie 
            Caption         =   "Asiento de Cierre"
         End
      End
      Begin VB.Menu Open_asi 
         Caption         =   "Abrir Asiento Registrado"
      End
      Begin VB.Menu Borr_asi 
         Caption         =   "Borrar Asiento"
      End
      Begin VB.Menu Restore 
         Caption         =   "Restaurar Anterior"
      End
      Begin VB.Menu li1 
         Caption         =   "-"
      End
      Begin VB.Menu asi_save 
         Caption         =   "Guardar"
      End
      Begin VB.Menu li2 
         Caption         =   "-"
      End
      Begin VB.Menu Print 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu li4 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Edit_01 
      Caption         =   "Edicion"
      Begin VB.Menu ed_cut 
         Caption         =   "Cortar"
      End
      Begin VB.Menu ed_copy 
         Caption         =   "Copiar"
      End
      Begin VB.Menu Ed_paste 
         Caption         =   "Pegar"
      End
      Begin VB.Menu Ed_selec 
         Caption         =   "Seleccionar Todo"
      End
      Begin VB.Menu li55 
         Caption         =   "-"
      End
      Begin VB.Menu Ed_Invert 
         Caption         =   "Invertir"
      End
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu restaurarasientoanulado 
         Caption         =   "Restaura Asiento Anulado"
      End
   End
   Begin VB.Menu Reports 
      Caption         =   "Consultas"
      Begin VB.Menu Diario 
         Caption         =   "Libro Diario General"
      End
      Begin VB.Menu Mayor 
         Caption         =   "Mayor por Cuenta"
      End
      Begin VB.Menu li5 
         Caption         =   "-"
      End
      Begin VB.Menu opbalan 
         Caption         =   "Opciones de Balance"
      End
   End
   Begin VB.Menu tablas 
      Caption         =   "Tablas Auxiliares"
      Begin VB.Menu plancue 
         Caption         =   "Plan de Cuentas"
      End
      Begin VB.Menu Cue_Banc 
         Caption         =   "Cuentas Bancarias"
      End
   End
   Begin VB.Menu tools 
      Caption         =   "Herramientas"
      Begin VB.Menu calcu 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu setup 
         Caption         =   "Configurar"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "ASIDIAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FEAS%
Dim SUDEBE#, SUHABE#
Dim ACTEDIT%
Dim PRIMO%, PRIMOL&
Dim HUBOCAMBIOS%
Dim MODOECO%
Dim HAYTACENCO%

Private Sub Asi_Aper_Click()
   Call Command14_Click
   Ltext1 = FECHATEX$(COMEJ%)
   PPXX% = InStr(NUPROA$, "-"): If PPXX% < 1 Then PPXX% = 1 + Len(NUPROA$)
   NUASIAP$ = Left$(NUPROA$, PPXX% - 1) + "-001"
   Call APERBASDAT("PLANCUEN")
   '
   ' VERIFICA SI YA EXISTE EL ASIENTO
   ACTICARGA% = 1
   SQLX$ = "SELECT * FROM DetaAsien WHERE NuidAsien='" & TRIM$(NUASIAP$) & "' AND SucuAsien = " & SUCUI% & " AND Statpase > 0 "
   Set PASEASIEN = Contable.OpenRecordset(SQLX$, dbOpenDynaset, dbReadOnly)
   On Error Resume Next
   PASEASIEN.MoveLast
   If Err Then
      ACTICARGA% = 0
   End If
   On Error GoTo 0
   '
   Ltext2 = NUASIAP$
   TIASI% = 8
   If TRIM$(Text4) = "" Then
      Text4 = "APERTURA DE CUENTAS AL " + FECHATEX$(COMEJ%)
   End If
   Text11.SetFocus
   '
End Sub

Sub Asi_CaBan_Click()
   Call Command14_Click    ' NUEVO ASIENTO - LIMPIA FORMULARIO
   TIASI% = 3
   Label3.Caption = "(Resumen Mensual de Caja - Bancos)"
   Call ASICABAN
   If List1.ListCount < 1 Then
     Call MENSERR(24, "No hay Movimientos\en el Período Elegido")
     HUBOCAMBIOS% = 0
   End If
End Sub

Sub Asi_Caj_Click()
   If EXISTE%("FLGEST00.EXE") > 0 Then
      Call MENSERR(24, "No Compatible con Presente Versión.\  \Debe Generar Asiento Resumen\de Caja y Bancos.")
      Exit Sub
   End If
   '
   Call Command14_Click    ' NUEVO ASIENTO - LIMPIA FORMULARIO
   TIASI% = 4
   Label3.Caption = "(Resumen Mensual de Cobranza)"
   Call ASICOBRA
   If List1.ListCount < 1 Then
     Call MENSERR(24, "No hay Movimientos\en el Período Elegido")
     HUBOCAMBIOS% = 0
   End If
End Sub

Private Sub Asi_Cie_Click()
   '
   On Error Resume Next
   ACIERRE04.Show 1
   If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Imposible Generar Asiento de Cierre\en Entorno Actual.")
   End If
   On Error GoTo 0
   '
End Sub

Sub Asi_Com_Click()
   Call Command14_Click    ' NUEVO ASIENTO - LIMPIA FORMULARIO
   TIASI% = 2
   Label3.Caption = "(Resumen Mensual de Compras)"
   Call ASICOMPRAS
   If List1.ListCount < 1 Then
     Call MENSERR(24, "No hay Movimientos\en el Período Elegido")
     HUBOCAMBIOS% = 0
   End If
End Sub

Sub Asi_Man_Click()
   Call Command14_Click
End Sub

Sub Asi_Pag_Click()
   If EXISTE%("FLGEST00.EXE") > 0 Then
      Call MENSERR(24, "No Compatible con Presente Versión.\  \Debe Generar Asiento Resumen\de Caja y Bancos.")
      Exit Sub
   End If
   '
   Call Command14_Click    ' NUEVO ASIENTO - LIMPIA FORMULARIO
   TIASI% = 5
   Label3.Caption = "(Resumen Mensual de Pagos)"
   Call ASIPAGOS
   If List1.ListCount < 1 Then
     Call MENSERR(24, "No hay Movimientos\en el Período Elegido")
     HUBOCAMBIOS% = 0
   End If
End Sub

Private Sub asi_save_Click()
   Call Command1_Click
End Sub

Sub Asi_Ven_Click()
   Call Command14_Click    ' NUEVO ASIENTO - LIMPIA FORMULARIO
   TIASI% = 1
   Label3.Caption = "(Resumen Mensual de Ventas)"
   Call ASIVENTAS
   If List1.ListCount < 1 Then
     Call MENSERR(24, "No hay Movimientos\en el Período Elegido")
     HUBOCAMBIOS% = 0
   End If
End Sub

Private Sub Borr_asi_Click()
   Call Command24_Click
End Sub

'
Private Sub Command1_Click()
   Command1.Enabled = False
   '
   If POSEDI% < 0 Then
     Call MENSERR(24, "Asiento Bloqueado para Modificación")
     GoTo 90
   End If
   '
   If COMEJ% < 1 Then
      Call MENSERR(24, "Falta Abrir Ejercicio Contable")
      GoTo 90
   End If
   '
   If CONTROL$("PROPSCG", "CONTAGEN") = "GUEST" Then
      Call MENSERR(24, "Sistema Distribuido - Invitado en Contabilidad General\Imposible Grabar Este Asiento.")
      GoTo 90
   End If
   '
   If TRIM$(Ltext2) = "" Then
     Call MENSERR(24, "Número de Asiento No Válido.")
     GoTo 90
   End If
   '
   Screen.MousePointer = 11
10 For U& = 1 To List1.ListCount
     Z$ = List1.List(U& - 1)
     CUEII% = CUEINT%(Left$(Z$, 12))
     If CUEII% < 1 Then
       Call MENSERR(24, "Imposible Grabar este Asiento.\  \Número de Cuenta Incorrecto (" + TRIM$(Left$(Z$, 12)) + ")\en Renglon " + TRIM$(Str$(U&)) + ".")
       GoTo 90
     End If
     IMPA# = Val(Mid$(Z$, 52, 13)) - Val(Mid$(Z$, 66, 13))
     SUMATOT# = SUMATOT# + IMPA#
   Next U&
   '
   If Abs(SUMATOT#) >= 0.005 Then
      Call MENSERR(24, "Imposible Grabar este Asiento.\No Balancea.")
      GoTo 90
   End If
   '
   Call COPYSTRU("INDIASIE", "BLOQASIE")
   Call BLOQARCH("BLOQASIE")
   '
   ' BUSCAR POR NUMERO DE ASIENTO Y MARCAR statasien=(-1)
20 On Error Resume Next
   Set EncasiGra = Contable.OpenRecordset("ENCABEASIEN", dbOpenDynaset)
   If Err Then
     If COMALTER%("No se Puede Abrir 'ENCABEASIEN'\en Modo Exclusivo.\\Reintentar\Cancelar") <> 2 Then GoTo 20
     GoTo 90
   End If
   '
21 Set DetasiGra = Contable.OpenRecordset("DETAASIEN", dbOpenDynaset)
   If Err Then
     If COMALTER%("No se Puede Abrir 'DETAASIEN'\en Modo Exclusivo.\\Reintentar\Cancelar") <> 2 Then GoTo 21
     GoTo 90
   End If
   '
   NUASI$ = Ltext2
   '
   With EncasiGra
     .FindFirst "NUIDASIEN = '" & NUASI$ & "'"
     GoTo 31
30   .FindNext "NUIDASIEN = '" & NUASI$ & "'"
31   If .NoMatch = False Then
       If !StatAsien <> (-1) Then
         If DERACCE%("CONTRASI/NOMESS", "Abrir Asientos de Otros Usuarios") < 2 Then
           If IsNull(!Nume_Usuar) = True Or !Nume_Usuar <> USNBR% Then
             Call MENSERR(24, "Imposible Grabar Asiento '" + TRIM$(Ltext2) + "'.\Usuario sin Derecho de Control no Propietario.")
             GoTo 90
           End If
         End If
       End If
       '
       .Edit
       !StatAsien = (-1)
       .Update
       GoTo 30
     End If
   End With
   '
   With DetasiGra
     .FindFirst "NUIDASIEN = '" & NUASI$ & "'"
     GoTo 36
35   .FindNext "NUIDASIEN = '" & NUASI$ & "'"
36   If .NoMatch = False Then
       .Edit
       !StatPase = (-1)
       .Update
       GoTo 35
     End If
   End With
   '
   If List1.ListCount > 0 Then
        ' lee de pantalla y graba
        On Error GoTo 0
        With EncasiGra
          .AddNew
          !NUIDASIEN = NUASI$
          If Text4 = "" Then Text4 = " "
          !Refeasien = AJUSTI$(Text4, 64)
             FECHAU% = FECHANUM(Ltext1)
          !FECHASIEN = CVDAT(FECHAU%)
             PPXX% = InStr(NUASI$, "-")
          !OrFeAsien = Val(Mid$(NUASI$, PPXX% + 1))
          !SucuAsien = 0
          !NuDfAsien = 0
          !TIPOASIEN = TIASI%
          !StatAsien = 1  ' activo
          !EditAsien = POSEDI%
          !QPasAsien = List1.ListCount
          !DebeAsien = Val(Label21)
          !Observa = " "
          !Nume_Usuar = USNBR%
          .Update
          .Bookmark = .LastModified
        End With
        '
        ORPASE% = 0
        For K2% = 1 To List1.ListCount
          TTXX$ = List1.List(K2% - 1)
          COCUE$ = Mid$(TTXX$, 1, 12)
          DEPASE$ = Mid$(TTXX$, 14, 36)
          IDEBE# = Val(Mid$(TTXX$, 52, 13))
          IHABE# = Val(Mid$(TTXX$, 66, 13))
          MACON$ = Mid$(TTXX$, 120, 1)
          CENCOS1$ = Mid$(TTXX$, 121, 8) '\\\OT-06-0244-WAR-16/06/06///
          '
          With DetasiGra
            .AddNew
            !NUIDASIEN = NUASI$
              ORPASE% = ORPASE% + 1
            !OrNuPase = ORPASE%
            !REFEPASE = DEPASE$
              FECHAU% = FECHANUM(Ltext1)
            !FECHASIEN = CVDAT(FECHAU%)
              PPXX% = InStr(NUASI$, "-")
            !OrFeAsien = Val(Mid$(NUASI$, PPXX% + 1))
            !SucuAsien = 0
            !NuDfAsien = 0
            !StatPase = 1
            !CUECOEXT = COCUE$
            !CUECOINT = CUEINT%(COCUE$)
            !IDEBEPASE = IDEBE#
            !IHABEPASE = IHABE#
            !MARCONSI = MACON$
            !CenCos = CENCOS1$ '\\\OT-06-0244-WAR-16/06/06///
            .Update
            .Bookmark = .LastModified
          End With
        Next K2%
   End If
   '
   EncasiGra.Close
   DetasiGra.Close
   '
   If TIASI% = 1 Then
         Call GENASIVEN(NUASI$, 0, 0)
      ElseIf TIASI% = 2 Then
         Call GENASICOM(NUASI$, 0, 0)
      ElseIf TIASI% = 4 Then
         Call GENASICOB(NUASI$, 0, 0)
      ElseIf TIASI% = 5 Then
         Call GENASIPAG(NUASI$, 0, 0)
    End If
   ' re-genera el indice de asientos
   Call GENINDIAS
   '
60 On Error GoTo 0
   HUBOCAMBIOS% = 0
   '
   If TIASI% = 3 Then
      Call MARCASCAJA
   End If
   '
   If CONTROL$("ASICONTA", "IMPRECOM") <> "NO" Then
      Call PRINASIDIAR
   End If
   '
   On Error Resume Next
   EncasiGra.Close
   DetasiGra.Close
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call LIBARCH("BLOQASIE")
   '
   Call Command14_Click
   If TIASI% > 0 Then
      Unload ASIDIAR
   End If
   '
90 Screen.MousePointer = 1
   Command1.Enabled = True
End Sub

Sub GENINDIAS()
   '
   SUCUI% = 0
   LIMINF = CDbl(CVDAT(COMEJ%))
   LIMSUP = CDbl(CVDAT(FINEJ%))
   SQLX$ = "SELECT Nuidasien, RefeAsien FROM EncabeAsien "
   SQLX$ = SQLX$ + "WHERE ("
   SQLX$ = SQLX$ + "StatAsien > 0 "
   SQLX$ = SQLX$ + "AND SucuAsien = " & SUCUI% & " "
   SQLX$ = SQLX$ + "AND CDBL(FECHASIEN) >= " & LIMINF & " "
   SQLX$ = SQLX$ + "AND CDBL(FECHASIEN) <= " & LIMSUP & " "
   SQLX$ = SQLX$ + ") "
   SQLX$ = SQLX$ + "ORDER by NuidAsien ASC;"
   Set INDIASIEN = Contable.OpenRecordset(SQLX$, dbOpenDynaset, dbReadOnly)
   '
   JJ& = 0
   On Error Resume Next
   INDIASIEN.MoveLast
   If Err Then GoTo 60
   '
   CAREC% = INDIASIEN.RecordCount
   If CAREC% < 1 Then GoTo 60
   '
   FIN& = CAREC%
   REBLA$ = REGBLAN$("INDIASIE")
   Call COPYSTRU("INDIASIE", "IASDIA" + NEJER$)
   JJ& = 0
   INDIASIEN.MoveFirst
   '
50 Call TRACE(20, JJ&, FIN&)
   EL1$ = REBLA$
   Call REPLA(EL1$, INDIASIEN!NUIDASIEN, 1, 16)
   Call REPLA(EL1$, INDIASIEN!Refeasien, 17, 46)
   JJ& = JJ& + 1
   Call GRAREG("IASDIA" + NEJER$, EL1$, JJ&)
   '
   If JJ& < CAREC% Then
     INDIASIEN.MoveNext
     GoTo 50
   End If
   '
60 INDIASIEN.Close
   Call SETULTREG("IASDIA" + NEJER$, JJ&)
   Call CIERRARCH("IASDIA" + NEJER$)
   '
End Sub

Sub GENINDIASBAJA()
   '
   Call COPYSTRU("IASDIA", "IASDIB")
   '
   SUCUI% = 0
   LIMINF = CDbl(CVDAT(COMEJ%))
   LIMSUP = CDbl(CVDAT(FINEJ%))
   SQLX$ = "SELECT Nuidasien, RefeAsien FROM EncabeAsien "
   SQLX$ = SQLX$ + "WHERE ("
   SQLX$ = SQLX$ + "StatAsien < 0 "
   SQLX$ = SQLX$ + "AND SucuAsien = " & SUCUI% & " "
   SQLX$ = SQLX$ + "AND CDBL(FECHASIEN) >= " & LIMINF & " "
   SQLX$ = SQLX$ + "AND CDBL(FECHASIEN) <= " & LIMSUP & " "
   SQLX$ = SQLX$ + ") "
   SQLX$ = SQLX$ + "ORDER by NuidAsien ASC;"
   Set INDIASIEN = Contable.OpenRecordset(SQLX$, dbOpenDynaset, dbReadOnly)
   '
   JJ& = 0
   On Error Resume Next
   INDIASIEN.MoveLast
   If Err Then GoTo 60
   '
   CAREC% = INDIASIEN.RecordCount
   If CAREC% < 1 Then GoTo 60
   '
   FIN& = CAREC%
   REBLA$ = REGBLAN$("INDIASIE")
   Call COPYSTRU("INDIASIE", "IASDIB" + NEJER$)
   JJ& = 0
   INDIASIEN.MoveFirst
   '
50 Call TRACE(20, JJ&, FIN&)
   EL1$ = REBLA$
   Call REPLA(EL1$, INDIASIEN!NUIDASIEN, 1, 16)
   Call REPLA(EL1$, INDIASIEN!Refeasien, 17, 46)
   JJ& = JJ& + 1
   Call GRAREG("IASDIB" + NEJER$, EL1$, JJ&)
   '
   If JJ& < CAREC% Then
     INDIASIEN.MoveNext
     GoTo 50
   End If
   '
60 INDIASIEN.Close
   Call SETULTREG("IASDIB" + NEJER$, JJ&)
   Call CIERRARCH("IASDIB" + NEJER$)
   '
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   FPLANCUEN.Show 1
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Call AGREPASE(1)
End Sub

Private Sub Command12_Click()
   Call AGREPASE(2)
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   '
   If HUBOCAMBIOS% > 0 Then
     If COMALTER%("Desea Descartar los Cambios Realizados ?\\No, Conservar\Si, Descartar") <> 2 Then
       GoTo 99
     End If
   End If
   '
   List1.Clear
   Text4 = ""
   Label21 = ""
   Label22 = ""
   Text11 = ""
   Label4 = ""
   Text12 = ""
   Text14 = ""
   Text15 = ""
   Label8 = ""
   Label10 = ""
   Text111 = "" '\\\OT-06-0244-WAR-16/06/06///
   Label13 = ""
   Command4.Enabled = True
   Command8.Enabled = True
   ACTEDIT% = 0
   TIASI% = 0
   HUBOCAMBIOS% = 0
   POSEDI% = 1
   '
   Ltext1 = FECHATEX$(FULTASDIA%)
   Ltext2 = NUPROA$
   Label3.Caption = "(Asiento Manual)"
   On Error Resume Next
   Command21.SetFocus
   On Error GoTo 0
   '
99 Command14.Enabled = True
End Sub

Private Sub Command17_Click()
   OPASCO04.Show 1
End Sub

Private Sub Command18_Click()
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

Private Sub Command19_Click()
   '
   Command19.Enabled = False
   If HUBOCAMBIOS% > 0 Then
     If COMALTER%("Desea Descartar los Cambios Realizados ?\\No, Conservar\Si, Descartar") <> 2 Then
       GoTo 99
     End If
     HUBOCAMBIOS% = 0
   End If
   Call CIERRARCH("*.*")
   Unload Me
   Exit Sub
   '
99 Command19.Enabled = True
End Sub



Private Sub Command21_GotFocus()
   Ltext1.ForeColor = RGB(255, 255, 255)
   Ltext1.BackColor = RGB(0, 0, 120)
End Sub

Private Sub Command21_LostFocus()
   Ltext1.ForeColor = RGB(0, 0, 0)
   Ltext1.BackColor = RGB(225, 225, 225)
End Sub

Private Sub Command24_Click()
   Command24.Enabled = False
   If List1.ListCount < 1 Then
        Call MENSERR(24, "Imposible Borrar Asiento no Cargado")
      Else
        If COMALTER%("Confirma Borrar Asiento " + Ltext2 + "\\No, Conservar\Si, Borrar Asiento") = 2 Then
            ' BUSCAR POR NUMERO DE ASIENTO Y MARCAR statasien=(-1)
            On Error Resume Next
            Encasien.Close
            Detasien.Close
            '
20          On Error Resume Next
            Set Encasien = Contable.OpenRecordset("ENCABEASIEN", dbOpenDynaset)
            If Err Then
              If COMALTER%("No se Puede Abrir 'ENCABEASIEN'\en Modo Exclusivo.\\Reintentar\Cancelar") <> 2 Then GoTo 20
              GoTo 90
            End If
            '
21          Set Detasien = Contable.OpenRecordset("DETAASIEN", dbOpenDynaset)
            If Err Then
              If COMALTER%("No se Puede Abrir 'DETAASIEN'\en Modo Exclusivo.\\Reintentar\Cancelar") <> 2 Then GoTo 21
              GoTo 90
            End If
            '
            Screen.MousePointer = 11
            NUASI$ = Ltext2
            '
            With Encasien
              .FindFirst "NUIDASIEN = '" & NUASI$ & "'"
              GoTo 31
30            .FindNext "NUIDASIEN = '" & NUASI$ & "'"
31            If .NoMatch = False Then
                If !StatAsien <> (-1) Then
                  If DERACCE%("CONTRASI/NOMESS", "Abrir Asientos de Otros Usuarios") < 2 Then
                    If IsNull(!Nume_Usuar) = True Or !Nume_Usuar <> USNBR% Then
                      Call MENSERR(24, "Imposible Borrar Asiento '" + TRIM$(Ltext2) + "'.\Usuario sin Derecho de Control no Propietario.")
                      GoTo 90
                    End If
                  End If
                End If
                '
                .Edit
                !StatAsien = (-1)
                .Update
                GoTo 30
              End If
            End With
            '
            With Detasien
              .FindFirst "NUIDASIEN = '" & NUASI$ & "'"
              GoTo 36
35            .FindNext "NUIDASIEN = '" & NUASI$ & "'"
36            If .NoMatch = False Then
                .Edit
                !StatPase = (-1)
                .Update
                GoTo 35
              End If
            End With
            '
            Encasien.Close
            Detasien.Close
            '
            Call ABRECABAN
            With CajaBanc
              .FindFirst "NuAsiDiar = '" & NUASI$ & "'"
              GoTo 46
45            .FindNext "NuAsiDiar = '" & NUASI$ & "'"
46            If .NoMatch = False Then
                .Edit
                !NuAsiDiar = ""
                .Update
                GoTo 45
              End If
            End With
            Screen.MousePointer = 1
            '
            Call Command14_Click    ' LIMPIA PANTALLA
            '
            ' re-genera el indice de asientos
            Call GENINDIAS
            Call CIERRARCH("*.*")
            Call BAJALDISCO
            '
60          On Error GoTo 0
            HUBOCAMBIOS% = 0
            '
        End If
   End If
90 Command24.Enabled = True
End Sub

Private Sub Command25_Click()
   Command25.Enabled = False
   '
   If HUBOCAMBIOS% > 0 Then
     If COMALTER%("Desea Descartar los Cambios Realizados ?\\No, Conservar\Si, Descartar") <> 2 Then
       GoTo 99
     End If
     HUBOCAMBIOS% = 0
   End If
   '
   If EXISTE%(LUDAT$ + "IASDIA" + NEJER$ + ".RFS") < 1 Then
     Call COPYSTRU("INDIASIE", "IASDIA" + NEJER$)
   End If
   '
   'If ULTREG&("IASDIA" + NEJER$) < 1 Then
     Call GENINDIAS
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   'End If
   '
   FORSELEC.ABUSQUE = Ltext2
   Call FRESHECHO("IASDIA" + NEJER$)
   Call SELECHO("IASDIA" + NEJER$)
   VDV$ = VALACT1$("IASDIA" + NEJER$)
   If VDV$ <> "" Then
     ACTICARGA% = 1
     Ltext2 = VDV$
   End If
   '
99 Command25.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   COEXT$ = TRIM$(Text11)
   '
   If COEXT$ = "" Or CUEINT%(COEXT$) < 1 Then
     COEXT$ = TRIM$(Left$(List1.TEXT, 12))
   End If
   '
   If COEXT$ = "" Then
     COEXT$ = TRIM$(Left$(REGSEL$("ECUECOT"), 12))
   End If
   '
   CUEII% = CUEINT%(COEXT)
   If CUEII% > 0 Then
     Call INICONTA.LMAYORPAN(CUEII%)
   End If
99 Command22.Enabled = True
End Sub

Private Sub Command29_Click()
   Command29.Enabled = False
   If TIASI% > 0 And TIASI% <> 4 And TIASI% <> 8 Then Command1.SetFocus: GoTo 99
   If POSEDI% < 0 Then Command1.SetFocus: GoTo 99
   If VALINUMAS% <> 1 Then GoTo 99
   '
   Call SELECHO("TAREFASI")
   VDEVU$ = VALACT1$("TAREFASI")
   If TRIM$(VDEVU$) <> "" Then
     Text4 = TRIM$(VDEVU$)
     Text11.SetFocus
   End If
   '
99 Command29.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   If HUBOCAMBIOS% > 0 Then
     If COMALTER%("Desea Descartar los Cambios Realizados ?\\No, Conservar\Si, Descartar") <> 2 Then
       GoTo 99
     End If
     HUBOCAMBIOS% = 0
   End If
   '
   RETENUAS$ = Ltext2
   Call Command14_Click     ' LIMPIA FORMULARIO
   ACTICARGA% = 1
   Ltext2 = RETENUAS$
   '
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Call BORRAPASE
End Sub

Private Sub Command5_Click()
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   'PRINTFORM
   Call PRINASIDIAR
End Sub

Private Sub Command6_Click()
   Call AGREPASE(0)
End Sub

Private Sub Command7_Click()
   If EXISTE%("C:\WINDOWS\CALC.EXE") > 0 Then
        Shell "C:\WINDOWS\CALC.EXE", 4
      Else
        Call MENSERR(24, "No se Encuentra Utilitario\'CALC.EXE'\en Carpeta 'WINDOWS' del Disco 'C:'.")
   End If
End Sub

Private Sub Command8_Click()
   Call EDITPASE
End Sub

Private Sub Command9_Click()

   Call GENINDIASBAJA
   Call SELECHO("IASDIB" + NEJER$)
   NUASI$ = VALACT1$("IASDIB" + NEJER$)
   
   ' BUSCAR POR NUMERO DE ASIENTO Y MARCAR statasien=(-1)
20 On Error Resume Next
   Set EncasiGra = Contable.OpenRecordset("ENCABEASIEN", dbOpenDynaset)
   If Err Then
     If COMALTER%("No se Puede Abrir 'ENCABEASIEN'\en Modo Exclusivo.\\Reintentar\Cancelar") <> 2 Then GoTo 20
     GoTo 90
   End If
   '
21 Set DetasiGra = Contable.OpenRecordset("DETAASIEN", dbOpenDynaset)
   If Err Then
     If COMALTER%("No se Puede Abrir 'DETAASIEN'\en Modo Exclusivo.\\Reintentar\Cancelar") <> 2 Then GoTo 21
     GoTo 90
   End If
   '
   With EncasiGra
     .FindFirst "NUIDASIEN = '" & NUASI$ & "'"
     If .NoMatch = False Then
       If !StatAsien = (-1) Then
        .Edit
        !StatAsien = 1
        .Update
       End If
     End If
   End With

90 Exit Sub
End Sub

Private Sub Command99_Click()
   '\\\OT-06-0244-WAR-16/06/06///
   If TIASI% > 0 And TIASI% <> 4 And TIASI% <> 8 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
   On Error Resume Next
    Text111.SetFocus 'aca daba el error 5 llamada no valida
   On Error GoTo 0
   '
   If ULTREG&("TACENCOS") < 1 Then
      Call MENSERR(24, "Falta Definir Centros de Costos")
      Exit Sub
   End If
   '
   RG1$ = REGSEL$("TACENCOS")
   If Len(RG1$) > 4 Then
      Text111.TEXT = Left$(RG1$, 8)
      Label12.Caption = ECOARCH$("TACENCOS", Text111.TEXT)
      Text14.SetFocus
   End If
   '\\\OT-06-0244-WAR-FIN///
End Sub

Private Sub Diario_Click()
   Call INICONTA.Diarge_Click
End Sub

Private Sub ed_copy_Click()
   Screen.MousePointer = 11
   TTXX$ = ""
   If Text4.SelLength > 0 Then
     TTXX$ = TTXX$ + "Referencia: " + Text4.SelText + "|"
   End If
   '
   For U& = 1 To List1.ListCount
     If List1.Selected(U& - 1) = True Then
       LL1T$ = List1.List(U& - 1)
       Call REPLA(LL1T$, " ", 120, 1)                                    ' QUITA LA MARCA DE CONCILIACION
       TTXX$ = TTXX$ + "Pase Nro." + TRIM$(Str$(U&)) + ": " + LL1T$ + "|"
     End If
   Next U&
   '
   Clipboard.Clear
   Clipboard.SetText TTXX$
   Screen.MousePointer = 1
   '
End Sub

Private Sub ed_cut_Click()
   TPI = List1.TopIndex
10 For U& = 1 To List1.ListCount
     If List1.Selected(U& - 1) = True Then
       List1.RemoveItem (U& - 1)
       GoTo 10
     End If
   Next U&
   '
   List1.ListIndex = (-1)
   Call SUMASIEN
   List1.TopIndex = TPI
End Sub

Private Sub Ed_Invert_Click()
   For KU& = 1 To List1.ListCount
     Z$ = List1.List(KU& - 1)
     IDE$ = Mid$(Z$, 52, 13)
     IHA$ = Mid$(Z$, 66, 13)
     Call REPLA(Z$, IHA$, 52, 13)
     Call REPLA(Z$, IDE$, 66, 13)
     List1.RemoveItem (KU& - 1)
     List1.AddItem Z$, KU& - 1
   Next KU&
   List1.Refresh
   Call SUMASIEN
   HUBOCAMBIOS% = 1
End Sub

Private Sub Ed_paste_Click()
   TTXY$ = Clipboard.GetText
   While Len(TTXY$) > 0
     PPXY% = InStr(TTXY$, "|")
     If PPXY% < 1 Then PPXY% = 1 + Len(TTXY$)
     TTZZ$ = Left$(TTXY$, PPXY% - 1)
     TTXY$ = Mid$(TTXY$, PPXY% + 1)
     '
     If Len(TTZZ$) > 0 Then
       If Left$(TTZZ$, 11) = "Referencia:" Then
           Text4 = TRIM$(Mid$(TTZZ$, 12))
         ElseIf Left$(TTZZ$, 4) = "Pase" Then
           PPYY% = InStr(TTZZ$, ":")
           List1.AddItem TRIM$(Mid$(TTZZ$, PPYY% + 1))
       End If
     End If
     '
   Wend
   Call SUMASIEN
   '
End Sub

Private Sub Ed_selec_Click()
   '
   Text4.SelStart = 0
   Text4.SelLength = Len(Text4)
   Text4.SetFocus
   '
   For U& = 1 To List1.ListCount
     List1.Selected(U& - 1) = True
   Next U&
   '
End Sub

Private Sub Exit_Click()
   Call Command19_Click
End Sub

'
Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Frame8.BackColor = &HFCD7B6
      Me.Frame9.BackColor = &HFCD7B6
      Me.Frame10.BackColor = &HFCD7B6
      Me.Frame11.Visible = False
      Me.Frame12.Visible = False
      
      MARCOBARRA.Top = 8000
      MARCOBARRA.Left = 200
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If

    Caption = Caption + " - " + TRIM$(EJACT$)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    MODOECO% = 1
    If CONTROL("ASICONTA", "DETAPASE") = "REFASIEN" Then MODOECO% = 2
    '
    If COMEJ% < 1 Then
       Call MENSERR(24, "Falta Abrir Ejercicios Contables")
       Hide
       Exit Sub
    End If
    '
    If ACTICARGA% > 0 Then Exit Sub
    '
    If TIASI% = 0 Then
         Command14_Click
      ElseIf TIASI% = 1 Then
         Label3.Caption = "(Resumen Mensual de Ventas)"
         Call ASIVENTAS
         If List1.ListCount < 1 Then
            Call MENSERR(24, "No hay Movimientos\en el Período Elegido")
            Text1.TEXT = "********"
         End If
      ElseIf TIASI% = 2 Then
         Label3.Caption = "(Resumen Mensual de Compras)"
         Call ASICOMPRAS
         If List1.ListCount < 1 Then
            Call MENSERR(24, "No hay Movimientos\en el Período Elegido")
            Text1.TEXT = "********"
         End If
      ElseIf TIASI% = 3 Then
         Label3.Caption = "(Resumen Mensual de Caja-Bancos)"
         Call RECAJABAN
         If List1.ListCount < 1 Then
            Call MENSERR(24, "No hay Movimientos\en el Período Elegido")
            Text1.TEXT = "********"
         End If
      ElseIf TIASI% = 4 Then
         Label3.Caption = "(Asiento de Apertura del Ejercicio)"
         If COMEJ% < 1 Then
            Call MENSERR(24, "Falta Abrir Ejercicio Contable")
            Text1.TEXT = "********"
            Exit Sub
         End If
         FEX = COMEJ%: Text1.TEXT = FECHATEX$(FEX)
         Text2.TEXT = "1.0"
         Call CARGASIDIAR
    End If
    '
    HAYTACENCOS% = 1
    If ULTREG&("TACENCOS") < 1 Then
      HAYTACENCOS% = 0
      Text111.Enabled = False
      Command99.Enabled = False
    End If
    '
    ACTEDIT% = 0
    '
End Sub


Private Sub Form_Unload(Cancel As Integer)
    TIASI% = 0
    HUBOCAMBIOS% = 0
    Call CIERRACONTA
    Call LIBARCH("ASICONT")
End Sub

Private Sub Label10_DblClick()
   If CUEINT%(Text11) > 0 Then
     If Val(Label10) >= 0.005 Then
       Text15 = Label10
       Call Command6_Click
     End If
   End If
End Sub

Private Sub Label11_Click()
   'OPASICONT.Show 1'Se comento x Error 424
End Sub

Private Sub Label21_Change()
   HUBOCAMBIOS% = 1
End Sub

Private Sub Label22_Change()
   HUBOCAMBIOS% = 1
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label8_Click()
   If CUEINT%(Text11) > 0 Then
     If Val(Label8) >= 0.005 Then
       Text14 = Label8
       Call Command6_Click
     End If
   End If
End Sub

Private Sub List1_DblClick()
   Call EDITPASE
End Sub
'
Private Sub LText1_DblClick()

   If TIASI% > 0 And TIASI% <> 4 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
   If FECHANUM(Ltext1) < 1 Then
      HOII% = HOY(HOS$)
      Ltext1 = HOS$
   End If
   If List1.ListCount > 0 Then
      Call MENSERR(24, "No Valido Cambiar Fecha en este Contexto")
      Command1.SetFocus
      Exit Sub
   End If
   On Error Resume Next
   FORSELEC.ABUSQUE = Ltext1
   On Error GoTo 0
   FE1% = FECHANUM(Ltext1)
   Call REPLA(VDEF$, MKI$(FE1%), 1, 2)
   
   'PRESENTA OBJETO PARA SELECCIONAR FECHA
10 OBX$ = OBJPLA$("ENTRAFECHA1", VDEF$)
   If OBX$ = "" Then GoTo 99
   '
   FE% = CVI(Mid$(OBX$, 1, 2))
   '
   FESELEC$ = FECHATEX$(FE%)
   If FESELEC$ <> "" Then
     Ltext1 = FESELEC$
     If VALIFECHAS% > 0 Then
         Call Command14_Click
         If List1.ListCount > 0 Then Exit Sub
         '
         Ltext1 = FESELEC$
         Ltext2 = NUPROA$
         On Error Resume Next
         Text4.SetFocus
         On Error GoTo 0
       Else
         Call MENSERR(24, "Fecha no Válida")
         VDEF$ = OBX$
         GoTo 10
     End If
   End If
99
End Sub
'
Private Sub Command21_Click()
   If TIASI% > 0 And TIASI% <> 4 And TIASI% <> 8 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
   If FECHANUM(Ltext1) < 1 Then
      HOII% = HOY(HOS$)
      Ltext1 = HOS$
   End If
   On Error Resume Next
   FORSELEC.ABUSQUE = Ltext1
   On Error GoTo 0
   Call SELECHO("SELFECHA")
   FESELEC$ = VALACT1$("SELFECHA")
   '
   If FESELEC$ <> "" Then
     Ltext1 = FESELEC$
     If VALIFECHAS% > 0 Then
         Call Command14_Click
         If List1.ListCount > 0 Then Exit Sub
         '
         Ltext1 = FESELEC$
         Ltext2 = NUPROA$
         On Error Resume Next
         Text4.SetFocus
         On Error GoTo 0
       Else
         Call MENSERR(24, "Fecha no Válida")
     End If
   End If
End Sub
'
Private Sub Ltext2_Change()
   If ACTICARGA% > 0 Then
     Call CARGASIDIAR
     ACTICARGA% = 0
   End If
End Sub

Private Sub Mayor_Click()
   Call INICONTA.Mayor_Click
End Sub

Sub Open_Asi_Click()
   Call Command25_Click
End Sub

Private Sub plancue_Click()
   FPLANCUEN.Show 1
End Sub

Private Sub Print_Click()
   Call Command5_Click
End Sub

Private Sub restaurarasientoanulado_Click()
   Call Command9_Click
End Sub

Private Sub Restore_Click()
   Call Command3_Click
End Sub

Private Sub setup_Click()
   Call Command17_Click
End Sub

Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 113 Then ' CORRESPONDE TECLA <F2>
     Call Command31_Click
   End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 27 Then
      If List1.ListCount > 0 Then
         Call SUMASIEN
         If Abs(SUDEBE# - SUHABE#) < 0.005 Then
            If SUDEBE# >= 0.005 Then
               Command1.SetFocus
               Exit Sub
            End If
         End If
      End If
      Command14.SetFocus   ' BOTON APRUEBA Y GRABA
   End If
   '
   If KeyAscii% = 13 Then
      Text12.SetFocus
   End If
   '
End Sub


Private Sub Text111_Change()
  Label12.Caption = ECOARCH$("TACENCOS", Text111.TEXT) '\\\OT-06-0244-WAR-16/06/06///
End Sub
Private Sub TEXT111_LOSTFOCUS()
  '\\\OT-06-0244-WAR-20/06/06///
  If TRIM$(Text111.TEXT) <> "" Then
    If TRIM$(ECOARCH$("TACENCOS", Text111.TEXT)) = "" Then
      Call MENSERR(24, "Centro de Costos no Válido")
      On Error Resume Next
      Text111.SetFocus
      On Error GoTo 0
      Exit Sub
    End If
    FECAS1% = FECHANUM(Ltext1.Caption)
    If VALIPROY%(Text111.TEXT, FECAS1%) = (-1) Then
      Call MENSERR(24, "Fecha de Asiento Contable fuera del Límite\de Vigencia del Centro de Costos.")
      On Error Resume Next
      Text111.SetFocus
      On Error GoTo 0
    End If
  End If
  '\\\OT-06-0244-WAR-FIN///
End Sub
Private Sub Text111_DblClick()
  Call Command99_Click '\\\OT-06-0244-WAR-16/06/06///
End Sub

Private Sub Text111_GotFocus()
   '\\\OT-06-0244-WAR-16/06/06///
   If TIASI% > 0 And TIASI% <> 4 And TIASI% <> 8 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   '
   If VALICUENTA% <> 1 Then Exit Sub
   '
   Text111.SelStart = 0
   Text111.SelLength = Len(Text111)
   '\\\OT-06-0244-WAR-FIN///
End Sub

Private Sub Text111_KeyPress(KeyAscii As Integer)
'\\\OT-06-0244-WAR-16/06/06///
 '
   If KeyAscii = 27 Then
      If List1.ListCount > 0 Then
         Call SUMASIEN
         If Abs(SUDEBE# - SUHABE#) < 0.005 Then
            If SUDEBE# >= 0.005 Then
               Command1.SetFocus
            End If
         End If
      End If
      Command19.SetFocus
      Exit Sub
   End If
   '
   If KeyAscii% = 13 Then
      Text14.SetFocus
      Exit Sub
   End If
   '
   If KeyAscii = 8 Then
     If Text111 = "" Then
       Text12.SetFocus 'Text11.SetFocus\\\OT-06-0244-WAR-16/06/06///
       Exit Sub
     End If
   End If
   '
   '\\\OT-06-0244-WAR-FIN///
End Sub
Private Sub Text111_KeyDown(KeyCode As Integer, Shift As Integer)
   '\\\OT-06-0244-WAR-16/06/06///
   If KeyCode = 38 Or (KeyCode = 37 And Text111 = "") Then
     Text12.SetFocus
   End If
   If KeyCode = 113 Then ' CORRESPONDE TECLA <F2>
     Call Command99_Click
   End If
   '\\\OT-06-0244-WAR-FIN///
End Sub

Private Sub Text12_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 38 Or (KeyCode = 37 And Text12 = "") Then
     Text11.SetFocus
   End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 27 Then
      If List1.ListCount > 0 Then
         Call SUMASIEN
         If Abs(SUDEBE# - SUHABE#) < 0.005 Then
            If SUDEBE# >= 0.005 Then
               Command1.SetFocus
            End If
         End If
      End If
      Command19.SetFocus
      Exit Sub
   End If
   '
   If KeyAscii% = 13 Then
      If HAYTACENCOS% > 0 Then
          Text111.SetFocus 'Text14.SetFocus'\\\OT-06-0244-WAR-16/06/06///
        Else
          Text14.SetFocus
      End If
      Exit Sub
   End If
   '
   If KeyAscii = 8 Then
     If Text12 = "" Then
       Text11.SetFocus
       Exit Sub
     End If
   End If
   '
End Sub

Private Sub Text14_LostFocus()
   Text14.TEXT = CSTRING$(MKD$(Val(Text14.TEXT)), 3, 11, 2, 2)
End Sub

Private Sub Text15_LostFocus()
   Text15.TEXT = CSTRING$(MKD$(Val(Text15.TEXT)), 3, 11, 2, 2)
End Sub

Private Sub Text14_GotFocus()
   If TIASI% > 0 And TIASI% <> 4 And TIASI% <> 8 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
   If VALICUENTA% <> 1 Then Exit Sub
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   If Abs(Val(Text14.TEXT)) < 0.005 Then
      Text14.TEXT = "       0.00"
   End If
   Text14.SelStart = 0
   Text14.SelLength = Len(Text14.TEXT)
End Sub

Private Sub Text15_GotFocus()
   If TIASI% > 0 And TIASI% <> 4 And TIASI% <> 8 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
   If VALICUENTA% <> 1 Then Exit Sub
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   If Abs(Val(Text15.TEXT)) < 0.005 Then
      Text15.TEXT = "       0.00"
   End If
   Text15.SelStart = 0
   Text15.SelLength = Len(Text14.TEXT)
End Sub
Private Sub Text14_Change()
   If Abs(Val(Text14.TEXT)) >= 0.005 Then
      Text15.TEXT = "       0.00"
   End If
End Sub

Private Sub Text15_Change()
   If Abs(Val(Text15.TEXT)) >= 0.005 Then
      Text14.TEXT = "       0.00"
   End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 27 Then
      If List1.ListCount > 0 Then
         Call SUMASIEN
         If Abs(SUDEBE# - SUHABE#) < 0.005 Then
            If SUDEBE# >= 0.005 Then
               Command1.SetFocus
            End If
         End If
      End If
      Command19.SetFocus
   End If
   '
   If KeyAscii% = 13 Then
      If Abs(Val(Text14.TEXT)) > 0 Then
           Call AGREPASE(0)
         Else
           Text15.SetFocus
      End If
   End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 27 Then
      If List1.ListCount > 0 Then
         Call SUMASIEN
         If Abs(SUDEBE# - SUHABE#) < 0.005 Then
            If SUDEBE# >= 0.005 Then
               Command1.SetFocus
            End If
         End If
      End If
      Command19.SetFocus
   End If
   '
   If KeyAscii% = 13 Then
      If Abs(Val(Text15.TEXT)) > 0 Then
           Call AGREPASE(0)
         Else
           Text14.SetFocus
      End If
   End If
End Sub

Private Sub Command2_Click()
    Call HELPONLINE("ASIDIAR")
End Sub
'
Private Sub Text4_DblClick()
   Text4 = ULOCASE$(Text4)
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
     Text11.SetFocus
   End If
End Sub

Private Sub Text4_GOTFOCUS()
   If TIASI% > 0 And TIASI% <> 4 And TIASI% < 8 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
   If VALINUMAS% <> 1 Then Exit Sub
   '
   Text4.SelStart = 0
   Text4.SelLength = Len(Text4)
End Sub

Private Sub Text11_DblClick()
   Call Command31_Click
End Sub
'
Private Sub Text11_Change()
   'If VALICUENTA% <> 1 Then Exit Sub
   If Text11 = "" Then Label4 = ""
End Sub
Private Sub Command31_Click()
   If TIASI% > 0 And TIASI% <> 4 And TIASI% <> 8 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
   Text11.SetFocus
   '
   If ULTREG&("ECUECON") < 1 Then
      Call MENSERR(24, "Falta Definir Plan de Cuentas Contables")
      FPLANCUEN.Show 1
      Exit Sub
   End If
   '
   RG1$ = REGSEL$("ECUECON")
   If Len(RG1$) > 4 Then
      Text11.TEXT = Left$(RG1$, 12)
      Label4.Caption = DENOCUE$(CUEINT%(Text11))
      Text12.SetFocus
   End If
End Sub
'
Private Sub Text11_GotFocus()
   If TIASI% > 0 And TIASI% <> 4 And TIASI% <> 8 Then
      On Error Resume Next
      Command1.SetFocus
      Exit Sub
   End If
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
   If VALINUMAS% <> 1 Then
      'Ltext2.SetFocus
      Exit Sub
   End If
   '
   Text11.SelStart = 0
   Text11.SelLength = 12
End Sub
'
Private Sub Text12_GotFocus()
   If TIASI% > 0 And TIASI% <> 4 And TIASI% <> 8 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   '
   If VALICUENTA% <> 1 Then Exit Sub
   '
   If Text12.TEXT = "" Then
     If (MODOECO% <> 2 Or TRIM$(Text4) = "") Then
         Text12.TEXT = TRIM$(Label4.Caption)
       Else
         Text12.TEXT = TRIM$(Text4.TEXT)
     End If
   End If
   '
   Text12.SelStart = 0
   Text12.SelLength = Len(Text12)
End Sub
'

Private Sub VScroll1_Change()
   If TIASI% > 0 Then Exit Sub
   Static VALANT%
   If VALANT% = 0 Then VALANT% = 16384
   NROAC$ = TRIM$(Ltext2)
   '
   If EXISTE%(LUDAT$ + "IASDIA" + NEJER$ + ".RFS") < 1 Then
     Call COPYSTRU("INDIASIE", "IASDIA" + NEJER$)
   End If
   '
   If ULTREG&("IASDIA" + NEJER$) < 1 Then
     Call GENINDIAS
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
   '
   URE& = ULTREG&("IASDIA" + NEJER$)
   If URE < 1 Then
     Exit Sub
   End If
   '
   PNUAS$ = TRIM$(Left$(REGLEIDO$("IASDIA" + NEJER$, 1), 12))
   UNUAS$ = TRIM$(Left$(REGLEIDO$("IASDIA" + NEJER$, URE&), 12))
   If NROAC$ = "" Or NROAC$ < PNUAS$ Then
        RGENC& = 0
        GoTo 85
      ElseIf NROAC$ > UNUAS$ Then
        RGENC& = 1 + ULTREG&("IASDIA" + NEJER$)
        GoTo 85
   End If
   '
   FIN& = ULTREG&("IASDIA" + NEJER$)
   LI& = 0: LS& = FIN& + 1
82 E& = Int((LS& - LI&) / 2)
   If E& < 1 Then GoTo 83
   L& = LI& + E&
   If L& < 1 Then RGENC& = 1: GoTo 85
   If L& > FIN& Then RGENC& = FIN& + 1: GoTo 85
   XX$ = REGLEIDO$("IASDIA" + NEJER$, L&)
   If TRIM$(Left$(XX$, 16)) < NROAC$ Then LI& = L&: GoTo 82
   If TRIM$(Left$(XX$, 16)) > NROAC$ Then LS& = L&: GoTo 82
83 RGENC& = L&
   '
85 If VScroll1.Value < VALANT% Then
        RGENC& = RGENC& + 1
      ElseIf VScroll1.Value > VALANT% Then
        RGENC& = RGENC& - 1
   End If
   VALANT% = VScroll1.Value
   If RGENC& > 0 Then
     If RGENC& <= ULTREG&("IASDIA" + NEJER$) Then
       XX$ = REGLEIDO$("IASDIA" + NEJER$, RGENC&)
       ACTICARGA% = 1
       Ltext2 = TRIM$(Left$(XX$, 12))
'           Call LEENCDIARIO(RGENC%)
'           NROAS = CVS(E1$)
'           Ltext2 = CSTRING$(MKS$(NROAS), 3, 8, 1, 2)
'           Call CARGASIDIAR
'         Else
'           Ltext1 = ""
'           VV% = VALIFECHAS%
'           Text4.SetFocus
      End If
   End If
End Sub
'
Function VALIFECHAS%()
   '
   FEAS% = FECHANUM(Ltext1)
   FEASA% = FEAS%
   If FEAS% < COMEJ% Or FEAS% > FINEJ% Then
      FEX = FEPROASDIA%
      FEAS% = FEX
      Ltext1 = FECHATEX$(FEX)
      Ltext2 = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
   End If
   '
   If FEAS% = FEASA% Then
       VALIFECHAS% = 1
     Else
       VALIFECHAS% = 0
   End If
   '
End Function

Function VALINUMAS%()
   '
   VALINUMAS% = 0
   If Ltext2 <> "" Then
     VALINUMAS% = 1
   End If
   '
End Function
'
Function VALICUENTA%()
   '
   CUE$ = TRIM$(Text11.TEXT)
   If CUE$ <> "" Then
     CUEII% = CUEINT%(CUE$)
     If CUEII% > 0 Then
       Label4.Caption = DENOCUE$(CUEII%)
       VALICUENTA% = 1
       Exit Function
     End If
   End If
   Label4.Caption = ""
   VALICUENTA% = 0
   On Error Resume Next 'Se atajo xq provocaba un error 5 llamada no valida
     Text11.SetFocus
   On Error GoTo 0
   '
End Function
'
Function REGENCAS%()
   '
   REGENCAS% = 0
   VV% = VALINUMAS%
   NRO = Val(Ltext2)
   LI% = 0: LS% = MAXENCDIARIO% + 1
   L% = MAXENCDIARIO%: ERRO% = 0
   If NRO > ULTASDIA + 0.01 Then
      FEASA = FECHANUM(Text1.TEXT)
      If FEASA < FULTASDIA% Then
         FEASA = FULTASDIA%
      End If
   GoTo 99
   End If
   '
82 E% = Int((LS% - LI%) / 2)
   If E% < 1 Then GoTo 98
   L% = LI% + E%
   If L% < 1 Or L% > MAXENCDIARIO% Then GoTo 98
   Call LEENCDIARIO(L%)
   If CVS(E1$) < NRO - 0.01 Then LI% = L%: GoTo 82
   If CVS(E1$) > NRO + 0.01 Then LS% = L%: GoTo 82
   REGENCAS% = L%
   Exit Function
   '
98 FEASA = FECHANUM(Text1.TEXT)
   If FEASA < COMEJ% Then FEASA = COMEJ%
   If LI% > 0 Then
      Call LEENCDIARIO(LI%)
      If FEASA < CVI(E3$) Then FEASA = CVI(E3$)
   End If
   If LS% < LI% + 1 Then LS% = LI% + 1
   If LS% <= MAXENCDIARIO% Then
      Call LEENCDIARIO(LS%)
      If FEASA > CVI(E3$) Then FEASA = CVI(E3$)
   End If
   '
99 Text1.TEXT = FECHATEX$(FEASA)
   '
End Function

Function NUPROA$()
   '
   NUPROA$ = ""
   If VALIFECHAS% <> 1 Then GoTo 20
   '
   FEASX = FECHANUM(Ltext1)
   FECHAS$ = FECHATEX$(FEASX)
   FECHAU% = FEASX
   FEASYYYY = CVDAT(FECHAU%)
   NUIDAS$ = Mid$(FECHAS$, 7, 2) + Mid$(FECHAS$, 4, 2) + Mid$(FECHAS$, 1, 2)
   If Left$(NUIDAS$, 1) = "0" Then NUIDAS$ = Mid$(NUIDAS$, 2)
   ORDAS% = 0
   '
   SUCUI% = 0
   Set SELECASIEN = Contable.OpenRecordset("SELECT * FROM EncaBEAsien WHERE (STR$(FECHASIEN) = '" & FEASYYYY & "' and SucuAsien = " & SUCUI% & ") ORDER BY  ORFEASIEN DESC;", dbOpenDynaset, dbReadOnly)
   On Error Resume Next
   SELECASIEN.MoveFirst
   If Err Then GoTo 20
   ORDAS% = SELECASIEN!OrFeAsien
   '
20 On Error GoTo 0
   ORDAS% = ORDAS% + 1
   ORDA$ = TRIM$(Str$(ORDAS%))
   While Len(ORDA$) < 3: ORDA$ = "0" + ORDA$: Wend
   NUPROA$ = NUIDAS$ + "-" + ORDA$
   '
End Function

Sub RENUMERAS(FEASPE)
   Screen.MousePointer = 11
   Call ABRENCDIARIO
   FINALIX% = MAXENCDIARIO%
   For UKI% = 1 To FINALIX%
     Get #NX9%, UKI%, E0$
     NUASAN = CVS(Mid$(E0$, 1, 4))
     FEANTE% = CVI(Mid$(E0$, 7, 2))
     If FEANTE% > FEASPE + 0.1 Then
       NUASNUE = (10 + Int(10 * NUASAN + 0.5)) / 10
       Mid$(E0$, 1, 4) = MKS$(NUASNUE)
       Put #NX9%, UKI%, E0$
     End If
   Next UKI%
   '
   Call ABREDIARIO
   FINALIX% = MAXDIARIO%
   For UKI% = 1 To FINALIX%
     Get #NX10%, UKI%, E10$
     NUASAN = CVS(Mid$(E10$, 3, 4))
     FEANTE% = CVI(Mid$(E10$, 11, 2))
     If FEANTE% > FEASPE + 0.1 Then
       NUASNUE = (10 + Int(10 * NUASAN + 0.5)) / 10
       Mid$(E10$, 3, 4) = MKS$(NUASNUE)
       Put #NX10%, UKI%, E10$
     End If
   Next UKI%
   '
   Close #NX9%: NX9% = 0
   Close #NX10%: NX10% = 0
   Screen.MousePointer = 1
   Call COMUNI("Re-Numeración Completada.\Re-Inicie Ahora el Sistema de Contabilidad")
   Call FINAL("")
   '
End Sub
'
Sub CARGASIDIAR()
   '
   SUCUI% = 0
   Set INDIASIEN = Contable.OpenRecordset("SELECT * FROM EncabeAsien WHERE (NuidAsien='" & TRIM$(Ltext2) & "' AND SucuAsien = " & SUCUI% & ");", dbOpenDynaset, dbReadOnly)
   On Error Resume Next
   INDIASIEN.MoveLast
   CAREC% = INDIASIEN.RecordCount
   If CAREC% < 1 Then
     Call MENSERR(24, "Imposible Cargar Asiento " + TRIM$(Ltext2))
     GoTo 99
   End If
   On Error GoTo 0
   '
   With INDIASIEN
     Ltext1 = Format(!FECHASIEN, "dd/mm/yy")
     Text4 = !Refeasien
     If DERACCE%("CONTRASI/NOMESS", "Abrir Asientos de Otros Usuarios") < 1 Then
       If IsNull(!Nume_Usuar) = True Or !Nume_Usuar <> USNBR% Then
         List1.Clear
         Text11 = ""
         Text12 = ""
         Text111 = ""
         Label13 = ""
         Call MENSERR(24, "Imposible Abrir Asiento '" + TRIM$(Ltext2) + "'.\Propietario: " + USERNAM$(!Nume_Usuar) + ".\   \Usuario Actual sin Derecho de Control no Propietario.")
         POSEDI% = (-1)
         GoTo 99
       End If
     End If
     '
     POSEDI% = !EditAsien
     TIASI% = !TIPOASIEN
   End With
   '
83 List1.Clear
   ACTEDIT% = 0
   Text11 = ""
   Text12 = ""
   Text111 = "" '\\\OT-06-0244-WAR-16/06/06///
   Label13 = ""
   Command4.Enabled = True
   Command8.Enabled = True
   VUELTA% = 0
85 SQLX$ = "SELECT * FROM DetaAsien WHERE NuidAsien='" & TRIM$(Ltext2) & "' AND SucuAsien = " & SUCUI% & " "
   If VUELTA% < 1 Then
     SQLX$ = SQLX$ & " AND Statpase > 0 "
   End If
   SQLX$ = SQLX$ + " ORDER BY OrNuPase ASC "
   Set PASEASIEN = Contable.OpenRecordset(SQLX$, dbOpenDynaset, dbReadOnly)
   On Error Resume Next
   PASEASIEN.MoveLast
   CAPAS% = PASEASIEN.RecordCount
   If CAPAS% < 1 Then
     If VUELTA% < 1 Then
       TTXX$ = "Imposible Cargar Asiento " + TRIM$(Ltext2) + ".\Puede Intentar la Reconstruccion Automatica\de este Asiento Aparentemente Dañado.\   \Nota: La Reconstrucción se Basa en Recuperar\Versiones Anteriores del Asiento. Puede\Aparecer Modificado o con Pases Duplicados.\Se Recomienda una Cuidadosa Revisión.\\Intentar Reconstruir\Abandonar"
       OPXX% = COMALTER%(TTXX$)
       If OPXX% = 1 Then
           VUELTA% = 1
           GoTo 85
         Else
           GoTo 99
       End If
     End If
     Call MENSERR(24, "Imposible Reconstruir Asiento " + TRIM$(Ltext2) + "\Consulte a su Soporte de Sistemas FLEXOFT")
     GoTo 99
   End If
   '
   PASEASIEN.MoveFirst
   IPASE% = 0
   '
10 IPASE% = IPASE% + 1
   With PASEASIEN
     Z$ = Space$(128)
     CUEII% = !CUECOINT
        Call REPLA(Z$, CODCUE$(CUEII%), 1, 12)
     DEPASE$ = !REFEPASE
     If TRIM$(DEPASE$) = "" Then DEPASE$ = DENOCUE$(CUEII%)
        Call REPLA(Z$, DEPASE$, 14, 36)
     IDEBE# = !IDEBEPASE
     IHABE# = !IHABEPASE
     Call REPLA(Z$, CSTRING$(MKD$(IDEBE#), 3, 12, 2, 2), 52, 13)
     Call REPLA(Z$, CSTRING$(MKD$(IHABE#), 3, 12, 2, 2), 66, 13)
     '\\\OT-06-0244-WAR-16/06/06///
     CENCOS1$ = "": MACON$ = " "
     On Error Resume Next
      CENCOS1$ = !CenCos
      MACON$ = !MARCONSI
     On Error GoTo 0
     If MACON$ = "" Then MACON$ = " "
     Call REPLA(Z$, MACON$, 120, 1)
     Call REPLA(Z$, CENCOS1$, 121, 8)
     '\\\OT-06-0244-WAR-FIN///
     List1.AddItem Z$
   End With
   '
   If IPASE% < CAPAS% Then
     PASEASIEN.MoveNext
     GoTo 10
   End If
   '
   List1.ListIndex = (-1)
   If TIASI% = 4 Then POSEDI% = 0
   Call SUMASIEN
   '
   If TIASI% = 8 Then
        Label3 = "(Asiento de Apertura)"
      ElseIf TIASI% = 9 Then
        Label3 = "(Asiento de Cierre)"
   End If
   '
   HUBOCAMBIOS% = 0
   Screen.MousePointer = 1
   '
99 End Sub
'
Sub CLEARASIDIAR()
   '
Exit Sub
   FEX = FEPROASDIA%
   FEAS% = FEX
   Text1.TEXT = FECHATEX$(FEX)
   Text2.TEXT = CSTRING$(MKS$(NUPROAS), 3, 8, 1, 2)
   List1.Clear
   YAESTABAS% = 0
   Text11.TEXT = ""
   Label4.Caption = ""
   Text12.TEXT = ""
   Text4.TEXT = Space$(40)
   Text14.TEXT = ""
   Text15.TEXT = ""
   Call SUMASIEN
   ACTEDIT% = 0
   '
   Text1.SetFocus
   '
End Sub

Sub AGREPASE(MODO%)
   '
   If TIASI% > 0 And TIASI% <> 4 And TIASI% <> 8 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
   If VALICUENTA% <> 1 Then
      Exit Sub
   End If
   '
   If VALINUMAS% <> 1 Then
      Text2.SetFocus
      Exit Sub
   End If
   '
   If Val(Text14.TEXT) < 0.005 Then
      If Val(Text15.TEXT) < 0.005 Then
         Text14.SetFocus
         Exit Sub
      End If
   End If
   '
   REAC% = List1.ListIndex
   If REAC% < 0 Or REAC% >= List1.ListCount Then
      MODO% = 1
   End If
   If MODO% = 0 Then
      If REAC% >= 0 Then
         If REAC% < List1.ListCount Then
            MODO% = 2
         End If
      End If
   End If
   '
   Z$ = Space$(128)
   Call REPLA(Z$, Text11.TEXT, 1, 12)
   DEPASE$ = Text12.TEXT: If TRIM$(DEPASE$) = "" Then DEPASE$ = DENOCUE$(CUEINT%(Text11.TEXT))
   Call REPLA(Z$, DEPASE$, 14, 36)
   Call REPLA(Z$, CSTRING$(MKD$(Val(Text14.TEXT)), 3, 12, 2, 2), 52, 13)
   Call REPLA(Z$, CSTRING$(MKD$(Val(Text15.TEXT)), 3, 12, 2, 2), 66, 13)
   '\\\OT-06-0244-WAR-16/06/06///
   'Los ultimos 8 caracteres de Z$ es el C. Costos
   Call REPLA(Z$, Label13, 120, 1)      ' MARCA DE CONCILIACION
   Call REPLA(Z$, Text111.TEXT, 121, 8)
   '\\\OT-06-0244-WAR-FIN///
   '
   If MODO% = 1 Then
        List1.AddItem Z$
        List1.ListIndex = (-1)
        If List1.TopIndex < List1.ListCount - 9 Then
           List1.TopIndex = List1.ListCount - 9
        End If
        REAC% = List1.ListCount - 1
      ElseIf MODO% = 2 Then
        List1.AddItem Z$, REAC%
   End If
   For KU& = 1 To List1.ListCount
      On Error Resume Next
      List1.Selected(KU& - 1) = False
      On Error GoTo 0
   Next KU&
   On Error Resume Next
   List1.ListIndex = REAC% + 1
   List1.Selected(REAC% + 1) = True
   On Error GoTo 0
   '
   Call SUMASIEN
   Text14.TEXT = Label8.Caption
   Text15.TEXT = Label10.Caption
   ACTEDIT% = 0
   Command4.Enabled = True
   Command8.Enabled = True
   '
   'If InStr(EMPREAC$, "DOSIVA") > 0 Then
     Text12.TEXT = ""
     Text111.TEXT = "" '\\\OT-06-0244-WAR-16/06/06///
     Label13 = " "
   'End If
   '
End Sub

Sub EDITPASE()
   If TIASI% > 0 And TIASI% <> 4 And TIASI% <> 8 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then Command1.SetFocus: Exit Sub
   If ACTEDIT% > 0 Then
      Exit Sub
   End If
   '
   U& = List1.ListIndex
   If U& >= 0 Then
      If U& < List1.ListCount Then
         Z$ = List1.List(U&)
         Text11.TEXT = Left$(Z$, 12)
         VV% = VALICUENTA%
         Text12.TEXT = Mid$(Z$, 14, 36)
         IDEBE# = Val(Mid$(Z$, 52, 13))
         IHABE# = Val(Mid$(Z$, 66, 13))
         Text14.TEXT = CSTRING$(MKD$(IDEBE#), 3, 11, 2, 2)
         Text15.TEXT = CSTRING$(MKD$(IHABE#), 3, 11, 2, 2)
         '\\\OT-06-0244-WAR-16/06/06///
         Text111 = Mid$(Z$, 121, 8) 'C. Costos
         Label13 = Mid$(Z$, 120, 1) ' MARCA DE CONCILIACION
         '\\\OT-06-0244-WAR-FIN///
         List1.RemoveItem U&
         If U& < List1.ListCount Then
              List1.ListIndex = U&
            Else
              List1.ListIndex = (-1)
         End If
         ACTEDIT% = 1
         Command4.Enabled = False
         Command8.Enabled = False
         Call SUMASIEN
      End If
   End If
End Sub

Sub BORRAPASE()
   If TIASI% > 0 And TIASI% <> 4 And TIASI% <> 8 Then Command1.SetFocus: Exit Sub
   If POSEDI% < 0 Then
      TIASORI% = Asc(Left$(E8$, 1))
      If TIASORI% = 1 Or TIASORI% = 2 Then
         If COMALTER%("Realmente Quiere Borrar Este Asiento ?\\No, Continuar\Si, Borrar") = 2 Then
            List1.Clear
            Label21.Caption = ""
            Label22.Caption = ""
            List1.Refresh
            Label21.Refresh
            Label22.Refresh
            Exit Sub
         End If
      End If
      Command1.SetFocus: Exit Sub
   End If
   U& = List1.ListIndex
   If U& >= 0 Then
      If U& < List1.ListCount Then
         List1.RemoveItem U&
         If U& < List1.ListCount Then
              List1.ListIndex = U&
              List1.Selected(U&) = True
              List1.Refresh
            Else
              List1.ListIndex = (-1)
         End If
         Call SUMASIEN
      End If
   End If
End Sub
'
Sub SUMASIEN()
   '
   SUDEBE# = 0: SUHABE# = 0
   For U& = 1 To List1.ListCount
      Z$ = List1.List(U& - 1)
      SUDEBE# = SUDEBE# + Val(Mid$(Z$, 52, 13))
      SUHABE# = SUHABE# + Val(Mid$(Z$, 66, 13))
   Next U&
   Label21.Caption = CSTRING$(MKD$(SUDEBE#), 3, 12, 2, 2)
   Label22.Caption = CSTRING$(MKD$(SUHABE#), 3, 12, 2, 2)
   '
   Label8.Caption = ""
   Label10.Caption = ""
   If SUDEBE# >= SUHABE# + 0.005 Then
        Label10.Caption = CSTRING$(MKD$(SUDEBE# - SUHABE#), 3, 11, 2, 2)
      ElseIf SUHABE# >= SUDEBE# + 0.005 Then
        Label8.Caption = CSTRING$(MKD$(SUHABE# - SUDEBE#), 3, 11, 2, 2)
   End If
   '
   On Error Resume Next
   Text11.SetFocus
   On Error GoTo 0
   '
End Sub

Sub ASIVENTAS()
   '
   TIASI% = 1
   Call FRESHALL
   Call CARSELMESA
10 Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) = "" Then
      Exit Sub
   End If
   '
   MMXX0$ = TRIM$(MESASEL$)
   MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
   If TRIM$(MMXX1$) = "" Then
      Exit Sub
   End If
   MESASEL$ = MMXX1$
   '
   If DES% < COMEJ% Or HAS% > FINEJ% Then
      Call MENSERR(24, "Período Fuera de Límites\del Ejercicio Activo.")
      GoTo 10
   End If
   '
   FEAS% = HAS%
   FEX = FEAS%: Ltext1 = FECHATEX$(FEX)
   FF$ = Ltext1
   If VALIFECHAS% > 0 Then
         NPA = NUPROA$
      Else
         Call MENSERR(24, "Fecha no Válida")
         GoTo 10
   End If
   '
   REFE$ = TRIM$("RESUMEN MENSUAL DE VENTAS DE " + MESASEL$)
   SUCUI% = 0
   SQLX$ = "SELECT * FROM EncabeAsien "
   SQLX$ = SQLX$ + "WHERE RefeAsien = '" & REFE$ & "' "
   SQLX$ = SQLX$ + "AND TIPOASIEN = 1 "
   SQLX$ = SQLX$ + "AND SucuAsien = " & SUCUI% & " "
   SQLX$ = SQLX$ + "AND STATASIEN = 1 ;"
   Set INDIASIEN = Contable.OpenRecordset(SQLX$, dbOpenDynaset, dbReadOnly)
   'Set INDIASIEN = Contable.OpenRecordset("SELECT * FROM EncabeAsien WHERE (SucuAsien = " & SUCUI% & " AND STATASIEN = 1 AND REFEASIEN = '" & REFE$ & "');", dbOpenDynaset, dbReadOnly)
   On Error Resume Next
   INDIASIEN.MoveLast
   CAREC% = INDIASIEN.RecordCount
   If CAREC% > 0 Then
       INDIASIEN.MoveFirst
       ENCUAS% = 1
       NROASE$ = INDIASIEN!NUIDASIEN
       Ltext1 = FF$
       Ltext2 = NROASE$
       Text4 = REFE$
       VV% = VALINUMAS%
       YAESTABAS% = 1
       Screen.MousePointer = 1
       CMDX% = COMALTER%("Asiento de Ventas de " + MESASEL$ + "\ya Registrado.\\Mantener\Refrescar")
       If CMDX% <> 2 Then
           Call CARGASIDIAR
           Exit Sub
       End If
     Else
       Ltext2 = NUPROA$
       Text4.TEXT = REFE$
   End If
   '
90 Screen.MousePointer = 11
   Call GENASIVEN("", DES%, HAS%)
   '
   Text4.TEXT = "RESUMEN MENSUAL DE VENTAS DE " + MESASEL$
   CONTA% = 0: SUMA# = 0
   List1.Clear: List2.Clear
   FIN& = ULTREG&("ASIDIAR")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     ASX$ = REGLEIDO$("ASIDIAR", U&)
     If Left$(ASX$, 4) = Chr$(1) + Chr$(TIASI%) + MKI$(FEAS%) Then
       If CVI(Mid$(ASX$, 5, 2)) > 0 Then
         CONTA% = CONTA% + 1
         IPASE# = CVD(Mid$(ASX$, 39, 8))
         SUMA# = SUMA# + (Int(100 * IPASE#) + 0.5) / 100
         '
         Z$ = Space$(128)
         Call REPLA(Z$, CODCUE$(CVI(Mid$(ASX$, 5, 2))), 1, 12)
         DEPASE$ = DENOCUE$(CVI(Mid$(ASX$, 5, 2)))
         Call REPLA(Z$, DEPASE$, 14, 36)
         If IPASE# > 0 Then
              Call REPLA(Z$, CSTRING$(MKD$(IPASE#), 3, 12, 2, 2), 52, 13)
            Else
              Call REPLA(Z$, CSTRING$(MKD$(Abs(IPASE#)), 3, 12, 2, 2), 66, 13)
         End If
         List2.AddItem "A" + Z$
       End If
     End If
   Next U&
   '
   For U& = 1 To List2.ListCount
      Z$ = Mid$(List2.List(U& - 1), 2)
      List1.AddItem Z$
   Next U&
   '
   Call SUMASIEN
   Screen.MousePointer = 1
   '
End Sub

Sub ASICOMPRAS()
   '
   TIASI% = 2
   Call FRESHALL
   Call CARSELMESA
10 Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) = "" Then
      Exit Sub
   End If
   '
   MMXX0$ = TRIM$(MESASEL$)
   MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
   If TRIM$(MMXX1$) = "" Then
      Exit Sub
   End If
   MESASEL$ = MMXX1$
   '
   If DES% < COMEJ% Or HAS% > FINEJ% Then
      Call MENSERR(24, "Período Fuera de Límites\del Ejercicio Activo.")
      GoTo 10
   End If
   '
   FEAS% = HAS%
   FEX = FEAS%: Ltext1 = FECHATEX$(FEX)
   FF$ = Ltext1
   If VALIFECHAS% > 0 Then
         NPA = NUPROA$
      Else
         Call MENSERR(24, "Fecha no Válida")
         GoTo 10
   End If
   '
   REFE$ = TRIM$("RESUMEN MENSUAL DE COMPRAS DE " + MESASEL$)
   SUCUI% = 0
   SQLX$ = "SELECT * FROM EncabeAsien "
   SQLX$ = SQLX$ + "WHERE RefeAsien = '" & REFE$ & "' "
   SQLX$ = SQLX$ + "AND TIPOASIEN = 2 "
   SQLX$ = SQLX$ + "AND SucuAsien = " & SUCUI% & " "
   SQLX$ = SQLX$ + "AND STATASIEN = 1 ;"
   Set INDIASIEN = Contable.OpenRecordset(SQLX$, dbOpenDynaset, dbReadOnly)
   'Set INDIASIEN = Contable.OpenRecordset("SELECT * FROM EncabeAsien WHERE (SucuAsien = " & SUCUI% & " AND STATASIEN = 1 AND REFEASIEN = '" & REFE$ & "');", dbOpenDynaset, dbReadOnly)
   On Error Resume Next
   INDIASIEN.MoveLast
   CAREC% = INDIASIEN.RecordCount
   If CAREC% > 0 Then
       INDIASIEN.MoveFirst
       ENCUAS% = 1
       NROASE$ = INDIASIEN!NUIDASIEN
       Ltext1 = FF$
       Ltext2 = NROASE$
       Text4 = REFE$
       VV% = VALINUMAS%
       YAESTABAS% = 1
       CMDX% = COMALTER%("Asiento de Compras de " + MESASEL$ + "\ya Registrado.\\Mantener\Refrescar")
       If CMDX% <> 2 Then
           Call CARGASIDIAR
           Exit Sub
       End If
     Else
       Ltext2 = NUPROA$
       Text4.TEXT = REFE$
   End If
   '
90 Screen.MousePointer = 11
   Call GENASICOM("", DES%, HAS%)
   '
   Text4.TEXT = "RESUMEN MENSUAL DE COMPRAS DE " + MESASEL$
   CONTA% = 0: SUMA# = 0
   List1.Clear: List2.Clear
   FIN& = ULTREG&("ASIDIAR")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     ASX$ = REGLEIDO$("ASIDIAR", U&)
     If Left$(ASX$, 4) = Chr$(1) + Chr$(TIASI%) + MKI$(FEAS%) Then
       If CVI(Mid$(ASX$, 5, 2)) > 0 Then
         CONTA% = CONTA% + 1
         IPASE# = CVD(Mid$(ASX$, 39, 8))
         SUMA# = SUMA# + (Int(100 * IPASE#) + 0.5) / 100
         '
         Z$ = Space$(128)
         Call REPLA(Z$, CODCUE$(CVI(Mid$(ASX$, 5, 2))), 1, 12)
         DEPASE$ = DENOCUE$(CVI(Mid$(ASX$, 5, 2)))
         Call REPLA(Z$, DEPASE$, 14, 36)
         If IPASE# > 0 Then
              Call REPLA(Z$, CSTRING$(MKD$(IPASE#), 3, 12, 2, 2), 52, 13)
            Else
              Call REPLA(Z$, CSTRING$(MKD$(Abs(IPASE#)), 3, 12, 2, 2), 66, 13)
         End If
         List2.AddItem "A" + Z$
       End If
     End If
   Next U&
   '
   For U& = 1 To List2.ListCount
      Z$ = Mid$(List2.List(U& - 1), 2)
      List1.AddItem Z$
   Next U&
   '
   Call SUMASIEN
   Screen.MousePointer = 1
   '
End Sub

Sub ASICOBRA()
   '
   TIASI% = 4
   Call FRESHALL
   Call CARSELMESA
10 Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) = "" Then
      Exit Sub
   End If
   '
   MMXX0$ = TRIM$(MESASEL$)
   MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
   If TRIM$(MMXX1$) = "" Then
      Exit Sub
   End If
   MESASEL$ = MMXX1$
   '
   If DES% < COMEJ% Or HAS% > FINEJ% Then
      Call MENSERR(24, "Período Fuera de Límites\del Ejercicio Activo.")
      GoTo 10
   End If
   '
   FEAS% = HAS%
   FEX = FEAS%: Ltext1 = FECHATEX$(FEX)
   FF$ = Ltext1
   If VALIFECHAS% > 0 Then
         NPA = NUPROA$
      Else
         Call MENSERR(24, "Fecha no Válida")
         GoTo 10
   End If
   '
   REFE$ = TRIM$("RESUMEN MENSUAL COBRANZA DE " + MESASEL$)
   SUCUI% = 0
   SQLX$ = "SELECT * FROM EncabeAsien "
   SQLX$ = SQLX$ + "WHERE RefeAsien = '" & REFE$ & "' "
   SQLX$ = SQLX$ + "AND TIPOASIEN = 4 "
   SQLX$ = SQLX$ + "AND SucuAsien = " & SUCUI% & " "
   SQLX$ = SQLX$ + "AND STATASIEN = 1 ;"
   Set INDIASIEN = Contable.OpenRecordset(SQLX$, dbOpenDynaset, dbReadOnly)
   'Set INDIASIEN = Contable.OpenRecordset("SELECT * FROM EncabeAsien WHERE (SucuAsien = " & SUCUI% & " AND STATASIEN = 1 AND REFEASIEN = '" & REFE$ & "');", dbOpenDynaset, dbReadOnly)
   On Error Resume Next
   INDIASIEN.MoveLast
   CAREC% = INDIASIEN.RecordCount
   If CAREC% > 0 Then
       INDIASIEN.MoveFirst
       ENCUAS% = 1
       NROASE$ = INDIASIEN!NUIDASIEN
       Ltext1 = FF$
       Ltext2 = NROASE$
       Text4 = REFE$
       VV% = VALINUMAS%
       YAESTABAS% = 1
       CMDX% = COMALTER%("Asiento de Cobranza de " + MESASEL$ + "\ya Registrado.\\Mantener\Refrescar")
       If CMDX% <> 2 Then
           Call CARGASIDIAR
           Exit Sub
       End If
     Else
       Ltext2 = NUPROA$
       Text4.TEXT = REFE$
   End If
   '
90 Screen.MousePointer = 11
   On Error GoTo 0
   Call GENASICOB("", DES%, HAS%)
   '
   Text4.TEXT = "RESUMEN MENSUAL COBRANZA DE " + MESASEL$
   CONTA% = 0: SUMA# = 0
   List1.Clear: List2.Clear
   FIN& = ULTREG&("ASIDIAR")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     ASX$ = REGLEIDO$("ASIDIAR", U&)
     If Left$(ASX$, 4) = Chr$(1) + Chr$(TIASI%) + MKI$(FEAS%) Then
       If CVI(Mid$(ASX$, 5, 2)) > 0 Then
         CONTA% = CONTA% + 1
         IPASE# = CVD(Mid$(ASX$, 39, 8))
         SUMA# = SUMA# + (Int(100 * IPASE#) + 0.5) / 100
         '
         Z$ = Space$(128)
         Call REPLA(Z$, CODCUE$(CVI(Mid$(ASX$, 5, 2))), 1, 12)
         DEPASE$ = DENOCUE$(CVI(Mid$(ASX$, 5, 2)))
         Call REPLA(Z$, DEPASE$, 14, 36)
         If IPASE# > 0 Then
              Call REPLA(Z$, CSTRING$(MKD$(IPASE#), 3, 12, 2, 2), 52, 13)
            Else
              Call REPLA(Z$, CSTRING$(MKD$(Abs(IPASE#)), 3, 12, 2, 2), 66, 13)
         End If
         List2.AddItem "A" + Z$
       End If
     End If
   Next U&
   '
   For U& = 1 To List2.ListCount
      Z$ = Mid$(List2.List(U& - 1), 2)
      List1.AddItem Z$
   Next U&
   '
   Call SUMASIEN
   Screen.MousePointer = 1
   '
End Sub

Sub ASIPAGOS()
   '
   TIASI% = 5
   Call FRESHALL
   Call CARSELMESA
10 Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) = "" Then
      Exit Sub
   End If
   '
   MMXX0$ = TRIM$(MESASEL$)
   MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
   If TRIM$(MMXX1$) = "" Then
      Exit Sub
   End If
   MESASEL$ = MMXX1$
   '
   If DES% < COMEJ% Or HAS% > FINEJ% Then
      Call MENSERR(24, "Período Fuera de Límites\del Ejercicio Activo.")
      GoTo 10
   End If
   '
   FEAS% = HAS%
   FEX = FEAS%: Ltext1 = FECHATEX$(FEX)
   FF$ = Ltext1
   If VALIFECHAS% > 0 Then
         NPA = NUPROA$
      Else
         Call MENSERR(24, "Fecha no Válida")
         GoTo 10
   End If
   '
   REFE$ = TRIM$("RESUMEN MENSUAL PAGOS DE " + MESASEL$)
   SUCUI% = 0
   SQLX$ = "SELECT * FROM EncabeAsien "
   SQLX$ = SQLX$ + "WHERE RefeAsien = '" & REFE$ & "' "
   SQLX$ = SQLX$ + "AND TIPOASIEN = 5 "
   SQLX$ = SQLX$ + "AND SucuAsien = " & SUCUI% & " "
   SQLX$ = SQLX$ + "AND STATASIEN = 1 ;"
   Set INDIASIEN = Contable.OpenRecordset(SQLX$, dbOpenDynaset, dbReadOnly)
   'Set INDIASIEN = Contable.OpenRecordset("SELECT * FROM EncabeAsien WHERE (SucuAsien = " & SUCUI% & " AND STATASIEN = 1 AND REFEASIEN = '" & REFE$ & "');", dbOpenDynaset, dbReadOnly)
   On Error Resume Next
   INDIASIEN.MoveLast
   CAREC% = INDIASIEN.RecordCount
   If CAREC% > 0 Then
       INDIASIEN.MoveFirst
       ENCUAS% = 1
       NROASE$ = INDIASIEN!NUIDASIEN
       Ltext1 = FF$
       Ltext2 = NROASE$
       Text4 = REFE$
       VV% = VALINUMAS%
       YAESTABAS% = 1
       CMDX% = COMALTER%("Asiento de Pagos de " + MESASEL$ + "\ya Registrado.\\Mantener\Refrescar")
       If CMDX% <> 2 Then
           Call CARGASIDIAR
           Exit Sub
       End If
     Else
       Ltext2 = NUPROA$
       Text4.TEXT = REFE$
   End If
   '
90 Screen.MousePointer = 11
   Call GENASIPAG("", DES%, HAS%)
   '
   Text4.TEXT = "RESUMEN MENSUAL PAGOS DE " + MESASEL$
   CONTA% = 0: SUMA# = 0
   List1.Clear: List2.Clear
   FIN& = ULTREG&("ASIDIAR")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     ASX$ = REGLEIDO$("ASIDIAR", U&)
     If Left$(ASX$, 4) = Chr$(1) + Chr$(TIASI%) + MKI$(FEAS%) Then
       If CVI(Mid$(ASX$, 5, 2)) > 0 Then
         CONTA% = CONTA% + 1
         IPASE# = CVD(Mid$(ASX$, 39, 8))
         SUMA# = SUMA# + (Int(100 * IPASE#) + 0.5) / 100
         '
         Z$ = Space$(128)
         Call REPLA(Z$, CODCUE$(CVI(Mid$(ASX$, 5, 2))), 1, 12)
         DEPASE$ = DENOCUE$(CVI(Mid$(ASX$, 5, 2)))
         Call REPLA(Z$, DEPASE$, 14, 36)
         If IPASE# > 0 Then
              Call REPLA(Z$, CSTRING$(MKD$(IPASE#), 3, 12, 2, 2), 52, 13)
            Else
              Call REPLA(Z$, CSTRING$(MKD$(Abs(IPASE#)), 3, 12, 2, 2), 66, 13)
         End If
         List2.AddItem "A" + Z$
       End If
     End If
   Next U&
   '
   For U& = 1 To List2.ListCount
      Z$ = Mid$(List2.List(U& - 1), 2)
      List1.AddItem Z$
   Next U&
   '
   Call SUMASIEN
   Screen.MousePointer = 1
   '
End Sub

Sub ASICABAN()
   '
   TIASI% = 3
   Call FRESHALL
   Call CARSELMESA
10 Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) = "" Then
      Exit Sub
   End If
   '
   MMXX0$ = TRIM$(MESASEL$)
   MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
   If TRIM$(MMXX1$) = "" Then
      Exit Sub
   End If
   MESASEL$ = MMXX1$
   MMXX2$ = Left$(MMXX1$, 3) + "-" + Right$(MMXX1$, 2)
   '
   If DES% < COMEJ% Or HAS% > FINEJ% Then
      Call MENSERR(24, "Período Fuera de Límites\del Ejercicio Activo.")
      GoTo 10
   End If
   '
   FEAS% = HAS%
   FEX = FEAS%: Ltext1 = FECHATEX$(FEX)
   FF$ = Ltext1
   If VALIFECHAS% > 0 Then
         NPA = NUPROA$
      Else
         Call MENSERR(24, "Fecha no Válida")
         GoTo 10
   End If
   '
   REFE$ = TRIM$("RESUMEN MENSUAL CAJA-BANCOS " + MMXX2$)
   SUCUI% = 0
   SQLX$ = "SELECT * FROM EncabeAsien "
   SQLX$ = SQLX$ + "WHERE RefeAsien = '" & REFE$ & "' "
   SQLX$ = SQLX$ + "AND TIPOASIEN = 3 "
   SQLX$ = SQLX$ + "AND SucuAsien = " & SUCUI% & " "
   SQLX$ = SQLX$ + "AND STATASIEN = 1 ;"
   Set INDIASIEN = Contable.OpenRecordset(SQLX$, dbOpenDynaset, dbReadOnly)
   'Set INDIASIEN = Contable.OpenRecordset("SELECT * FROM EncabeAsien WHERE (SucuAsien = " & SUCUI% & " AND STATASIEN = 1 AND REFEASIEN = '" & REFE$ & "');", dbOpenDynaset, dbReadOnly)
   On Error Resume Next
   INDIASIEN.MoveLast
   CAREC% = INDIASIEN.RecordCount
   If CAREC% > 0 Then
       INDIASIEN.MoveFirst
       ENCUAS% = 1
       NROASE$ = INDIASIEN!NUIDASIEN
       Ltext1 = FF$
       Ltext2 = NROASE$
       Text4 = REFE$
       VV% = VALINUMAS%
       YAESTABAS% = 1
       CMDX% = COMALTER%("Asiento de Caja-Bancos de " + MESASEL$ + "\ya Registrado.\\Mantener\Refrescar")
       If CMDX% <> 2 Then
           Call CARGASIDIAR
           Exit Sub
       End If
     Else
       Ltext2 = NUPROA$
       Text4.TEXT = REFE$
   End If
   '
90 Screen.MousePointer = 11
   Call GENASICABA("", DES%, HAS%)
   '
   Text4.TEXT = "RESUMEN MENSUAL CAJA-BANCOS " + MMXX2$
   CONTA% = 0: SUMA# = 0
   List1.Clear: List2.Clear
   FIN& = ULTREG&("ASIDIAR")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     ASX$ = REGLEIDO$("ASIDIAR", U&)
     If Left$(ASX$, 4) = Chr$(1) + Chr$(TIASI%) + MKI$(FEAS%) Then
       If CVI(Mid$(ASX$, 5, 2)) > 0 Then
         CONTA% = CONTA% + 1
         IPASE# = CVD(Mid$(ASX$, 39, 8))
         SUMA# = SUMA# + (Int(100 * IPASE#) + 0.5) / 100
         '
         Z$ = Space$(128)
         Call REPLA(Z$, CODCUE$(CVI(Mid$(ASX$, 5, 2))), 1, 12)
         DEPASE$ = DENOCUE$(CVI(Mid$(ASX$, 5, 2)))
         Call REPLA(Z$, DEPASE$, 14, 36)
         If IPASE# > 0 Then
              Call REPLA(Z$, CSTRING$(MKD$(IPASE#), 3, 12, 2, 2), 52, 13)
            Else
              Call REPLA(Z$, CSTRING$(MKD$(Abs(IPASE#)), 3, 12, 2, 2), 66, 13)
         End If
         List2.AddItem "A" + Z$
       End If
     End If
   Next U&
   '
   For U& = 1 To List2.ListCount
      Z$ = Mid$(List2.List(U& - 1), 2)
      List1.AddItem Z$
   Next U&
   '
   Call SUMASIEN
   Screen.MousePointer = 1
   '
End Sub

Sub RECAJABAN()
   '
   Dim CTA%(2048), IPAS#(2048)
10 Call SELECHO("SELFECHA")
   FF$ = VALACT1$("SELFECHA")
   If FF$ = "" Then
      Text1.TEXT = "********"
      Exit Sub
   End If
   '
   FEAS% = FECHANUM(FF$)
   If FEAS% < COMEJ% Or FEAS% > FINEJ% Or FEAS% > HOY(HOS$) Then
      Call MENSERR(24, "Fecha no Válida")
      Text1.TEXT = "********"
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   List1.Clear: List2.Clear
   Text1.TEXT = FF$
   '
   If VALIFECHAS% > 0 Then
         NPA = NUPROAS
      Else
         Call MENSERR(24, "Fecha no Válida")
         GoTo 10
   End If
   REFE$ = "RESUMEN DE CAJA-BANCOS AL " + FF$ + Space$(40)
   Text4.TEXT = REFE$
   '
   PRIMOL& = 0: HAYMO% = 0: ASHAS = 999999
   FIN& = MAXDIA&
   For IL& = 1 To MAXDIA&
      U& = IL&
      Call TRACE(20, U&, FIN&)
      Call LEESUCAJA(IL&)
      If CVI(D11$) > 0 Then
         If CVI(D15$) <= FEAS% Then
            If CVS(D19$) < 0.05 Then
               If PRIMOL& = 0 Then PRIMOL& = IL&
               CUE% = CVI(D11$)
               IM# = CVD(D16$)
               For J% = 1 To CONTA%
                  If CTA%(J%) = CUE% Then
                     IPAS#(J%) = IPAS#(J%) + (Int(100 * IM# + 0.5)) / 100
                     GoTo 75
                  End If
               Next J%
               CONTA% = CONTA% + 1
               CTA%(CONTA%) = CUE%
               IPAS#(CONTA%) = (Int(100 * IM# + 0.5)) / 100
            End If
         End If
      End If
75 Next IL&
   '
   K% = 0
   For i% = 1 To CONTA%
      If Abs(IPAS#(i%)) >= 0.005 Then
         K% = K% + 1
         CTA%(K%) = CTA%(i%)
         IPAS#(K%) = IPAS#(i%)
      End If
   Next i%
   CONTA% = K%
   '
   For i% = 1 To CONTA%
      Z$ = Space$(128)
      Call REPLA(Z$, CODCUE$(CTA%(i%)), 1, 12)
      DEPASE$ = DENOCUE$(CTA%(i%))
      Call REPLA(Z$, DEPASE$, 14, 36)
      If IPAS#(i%) > 0 Then
           Call REPLA(Z$, CSTRING$(MKD$(IPAS#(i%)), 3, 12, 2, 2), 52, 13)
         Else
           Call REPLA(Z$, CSTRING$(MKD$(Abs(IPAS#(i%))), 3, 12, 2, 2), 66, 13)
      End If
      List2.AddItem "A" + Z$
   Next i%
   '
   For U& = 1 To List2.ListCount
      Z$ = Mid$(List2.List(U& - 1), 2)
      List1.AddItem Z$
   Next U&
   '
   Call SUMASIEN
   Screen.MousePointer = 1
   '
End Sub

Sub MARCASCAJA()
   '
Exit Sub
   Screen.MousePointer = 11
   NROAS = Val(Text2.TEXT)
   FEAS% = FECHANUM(Text1.TEXT)
   If PRIMOL& < 1 Then PRIMOL& = 1
   FIN& = MAXDIA&
   For IL& = PRIMOL& To MAXDIA&
      U& = IL&
      Call TRACE(20, U&, FIN&)
      Call LEESUCAJA(IL&)
      If CVI(D11$) > 0 Then
         If CVI(D15$) <= FEAS% Then
            If CVS(D19$) < 0.05 Then
               DXX$ = D10$
               Call REPLA(DXX$, MKS$(NROAS), 55, 4)
               LSet D10$ = DXX$
               Put #NX6%, IL&, D10$
            End If
         End If
      End If
   Next IL&
   '
   Close #NX6%: NX6% = 0
   Screen.MousePointer = 1
   '
End Sub

Sub PRINASIDIAR()
   '
'  Exit Sub
   FORIPRE$ = CONTROL$("", "ASICONTA")
   '
   If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      FECHDOC$ = Ltext1.Caption
      NUMEDOC$ = TRIM$(Ltext2.Caption)
      TIPODOC$ = "Comprobante de Diario"
      '
      CODPARAM$(1) = "REF-ASIE"
      CODPARAM$(2) = "TOT-DEBE"
      CODPARAM$(3) = "TOT-HABE"
      CAPARDOC% = 3
      '
      DOCPARAM$(1) = TRIM$(Text4.TEXT)
      DOCPARAM$(2) = Label21.Caption
      DOCPARAM$(3) = Label22.Caption
      '
      CODTABU1$(1) = "NRO-PASE"
      CODTABU1$(2) = "COD-CUEN"
      CODTABU1$(3) = "DEN-CUEN"
      CODTABU1$(4) = "DET-PASE"
      CODTABU1$(5) = "IMP-DEBE"
      CODTABU1$(6) = "IMP-HABE"
      CODTABU1$(7) = "NRO-PASE"
      CACOLTAB1% = 7
      '
      CAITAB1% = 0
      For U& = 1 To List1.ListCount
         Z$ = List1.List(U& - 1)
         CAITAB1% = CAITAB1% + 1
         TETABU1$(1, CAITAB1%) = "." ' Str$(CAITAB1%)
         TETABU1$(2, CAITAB1%) = Left$(Z$, 12)
         CUE% = CUEINT%(Left$(Z$, 12))
         TETABU1$(3, CAITAB1%) = DENOCUE$(CUE%)
         TETABU1$(4, CAITAB1%) = Mid$(Z$, 14, 36)
         TETABU1$(5, CAITAB1%) = Mid$(Z$, 52, 13)
         TETABU1$(6, CAITAB1%) = Mid$(Z$, 66, 13)
         TETABU1$(7, CAITAB1%) = Str$(CAITAB1%)
      Next U&
      '
      Call PRINTDOC("ASICONTA")   ' asiento contable
      Screen.MousePointer = 1
      '
   End If
   '
End Sub

Sub GENASICOM(ASIDIA$, DES%, HAS%)
    '
    Static DESDAT, HASDAT
    '
    Screen.MousePointer = 11
    '
    Dim IMASI#(32767)
    CUEMX% = 0
    '
    If DES% > 0 Then DESDAT = CDbl(CVDAT(DES%))
    If HAS% > 0 Then HASDAT = CDbl(CVDAT(HAS%))
    '
    If IsEmpty(DESDAT) = True Or IsNull(DESDAT) = True Or IsEmpty(HASDAT) = True Or IsNull(HASDAT) = True Then
      Call MENSERR(24, "Imposible Registrar")
      Exit Sub
    End If
    '
    Call ABRECABAN
    SQLX0$ = "SELECT * FROM CAJABANC "
    SQLX0$ = SQLX0$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX0$ = SQLX0$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    If TRIM$(ASIDIA$) <> "" Then
        Set CABATEMP = CueCorri.OpenRecordset(SQLX0$, dbOpenDynaset)
      Else
        Set CABATEMP = CueCorri.OpenRecordset(SQLX0$, dbOpenSnapshot)
    End If
    '
    'SQLX$ = "SELECT * FROM CAJABANC "
    SQLX1$ = SQLX1$ + "INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX1$ = SQLX1$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX1$ = SQLX1$ + "AND (TipoMovim = 'FCOM' "
    SQLX1$ = SQLX1$ + "OR TipoMovim = 'AJUP') "
    'SQLX$ = SQLX$ + "Order by FechConta ASC "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    'JJ& = 0: REBLA$ = REGBLAN$("SUDICOM")
    'J2& = 0: REBL2$ = REGBLAN$("ANAICOM")
    
    HAYCAJA% = 0
    With CABATEMP    ' CajaBanc
      .FindFirst SQLX1$
15    SUMASI# = 0
      If .NoMatch = False Then
        '
        HAYCAJA% = 1
        MARCAREG = .Bookmark
        '
        NUMEFA$ = !CodiCom_Lar
        NCLIE% = !NuProve
        FECHAD = Int(CDbl(!FechConta))
        MOBASI$ = !TipoMovim
        CUEII% = !CueContaI
        '
        If CUEII% > 0 Then
          If TRIM$(ASIDIA$) <> "" Then
              .Edit
              !NuAsiDiar = ASIDIA$
              .Update
            Else
              ' ESTA CONDICION SE PUSO PARA UNIFICAR LOS AJUSTES DE SALDO
              ' DISTINTOS GRABADOS CON EL MISMO NÚMERO DE AJUSTE
              If MOBASI$ <> "AJUP" Then
                 IMASI#(CUEII%) = IMASI#(CUEII%) + (!IDEBECOMP - !IHABECOMP)
                 SUMASI# = SUMASI# + (!IDEBECOMP - !IHABECOMP)
              End If
              If CUEII% > CUEMX% Then CUEMX% = CUEII%
          End If
        End If
        '
        SQLX2$ = "INT(CDBL(FechConta)) = " & FECHAD & " "
        SQLX2$ = SQLX2$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
        SQLX2$ = SQLX2$ + "AND NuProve = " & NCLIE% & " "
        '       SQLX2$ = SQLX2$ + "AND TipoMovim <> '" & MOBASI$ & "' "
        ' ESTA CONDICION SE PUSO PARA UNIFICAR LOS AJUSTES DE SALDO
        ' DISTINTOS GRABADOS CON EL MISMO NÚMERO DE AJUSTE
        If MOBASI$ <> "AJUP" Then SQLX2$ = SQLX2$ + "AND TipoMovim <> '" & MOBASI$ & "' "
        '
        .FindFirst SQLX2$
25      If .NoMatch = False Then
          '
          CUEII% = !CueContaI
          If CUEII% > 0 Then
            If TRIM$(ASIDIA$) <> "" Then
                .Edit
                !NuAsiDiar = ASIDIA$
                .Update
              Else
                IMASI#(CUEII%) = IMASI#(CUEII%) + (!IDEBECOMP - !IHABECOMP)
                SUMASI# = SUMASI# + (!IDEBECOMP - !IHABECOMP)
                If CUEII% > CUEMX% Then CUEMX% = CUEII%
            End If
          End If
          '
          .FindNext SQLX2$
          GoTo 25
        End If
        '
        If Abs(SUMASI#) >= 0.005 Then
          Call MENSERR(24, "Comprobante " + NUMEFA$ + "\del " + FECHATEX$(CVINT%(FECHAD)) + "\del Proveedor '" + TRIM$(RASOCLI$((-1) * NCLIE%)) + "'\no Balancea.")
        End If
        '
        .Bookmark = MARCAREG
        .FindNext SQLX1$
        GoTo 15
      End If
    End With
    CABATEMP.Close
    '
    If HAYCAJA% < 1 Then
      FIN& = ULTREG&("ASIDIAR")
      TTXX$ = Chr$(1) + Chr$(2) + MKI$(HAS%)
      For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        ASX$ = REGLEIDO$("ASIDIAR", U&)
        If Left$(ASX$, 4) = TTXX$ Then
           CUEXI% = CVI(Mid$(ASX$, 5, 2))
           If CUEXI% > 0 Then
             IMASI#(CUEXI%) = CVD(Mid$(ASX$, 39, 8))
             If CUEXI% > CUEMX% Then CUEMX% = CUEXI%
           End If
        End If
      Next U&
    End If
    '
    TTXX$ = Chr$(1) + Chr$(2) + MKI$(HAS%)
    V& = 0: FIN& = ULTREG&("ASIDIAR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("ASIDIAR", U&)
      If Left$(X$, 4) <> TTXX$ Then
        V& = V& + 1
        Call GRAREG("ASIDIAR", X$, V&)
      End If
    Next U&
    '
    AGRECU% = 0
    For K% = 1 To CUEMX%
      If Abs(IMASI#(K%)) >= 0.005 Then
        X$ = REGBLAN$("ASIDIAR")
        Call REPLA(X$, TTXX$, 1, 4)
        Call REPLA(X$, MKI$(K%), 5, 2)
        Call REPLA(X$, MKD$(IMASI#(K%)), 39, 8)
        V& = V& + 1
        Call GRAREG("ASIDIAR", X$, V&)
      End If
    Next K%
    Call SETULTREG("ASIDIAR", V&)
    '
End Sub

Sub GENASIVEN(ASIDIA$, DES%, HAS%)
    '
    Static DESDAT, HASDAT
    '
    Screen.MousePointer = 11
    '
    Dim IMASI#(32767)
    CUEMX% = 0
    '
    If DES% > 0 Then DESDAT = CDbl(CVDAT(DES%))
    If HAS% > 0 Then HASDAT = CDbl(CVDAT(HAS%))
    '
    If IsEmpty(DESDAT) = True Or IsNull(DESDAT) = True Or IsEmpty(HASDAT) = True Or IsNull(HASDAT) = True Then
      Call MENSERR(24, "Imposible Registrar")
      Exit Sub
    End If
    '
    Call ABRECABAN
    '\\\OT-05-0416-WAR-23/06/05///
    SQLX0$ = "SELECT * FROM CAJABANC "
    SQLX0$ = SQLX0$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX0$ = SQLX0$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    If TRIM$(ASIDIA$) <> "" Then
        Set CABATEMP = CueCorri.OpenRecordset(SQLX0$, dbOpenDynaset)
      Else
        Set CABATEMP = CueCorri.OpenRecordset(SQLX0$, dbOpenSnapshot)
    End If
    '\\\OT-05-0416-WAR-FIN///
    '
    'SQLX$ = "SELECT * FROM CAJABANC "
    SQLX1$ = SQLX1$ + "INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX1$ = SQLX1$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX1$ = SQLX1$ + "AND (TipoMovim = 'FVEN' "
    SQLX1$ = SQLX1$ + "OR TipoMovim = 'AJUD') "
    'SQLX$ = SQLX$ + "Order by FechConta ASC "
    '
    HAYCAJA% = 0
    With CABATEMP 'CajaBanc
      .FindFirst SQLX1$
15    SUMASI# = 0
      If .NoMatch = False Then
        '
        HAYCAJA% = 1
        MARCAREG = .Bookmark
        '
        NUMEFA$ = !CodiCom_Lar
        NCLIE% = !NuClien
        FECHAD = Int(CDbl(!FechConta))
        MOBASI$ = !TipoMovim
        CUEII% = !CueContaI
        '
        If CUEII% > 0 Then
          If TRIM$(ASIDIA$) <> "" Then
              .Edit
              !NuAsiDiar = ASIDIA$
              .Update
            Else
              ' ESTA CONDICION SE PUSO PARA UNIFICAR LOS AJUSTES DE SALDO
              ' DISTINTOS GRABADOS CON EL MISMO NÚMERO DE AJUSTE
              If MOBASI$ <> "AJUD" Then
                 IMASI#(CUEII%) = IMASI#(CUEII%) + ROUND(!IDEBECOMP) - ROUND(!IHABECOMP)
                 SUMASI# = SUMASI# + ROUND(!IDEBECOMP) - ROUND(!IHABECOMP)
              End If
              If CUEII% > CUEMX% Then CUEMX% = CUEII%
          End If
        End If
        '
        SQLX2$ = "INT(CDBL(FechConta)) = " & FECHAD & " "
        SQLX2$ = SQLX2$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
        SQLX2$ = SQLX2$ + "AND (NuClien = " & NCLIE% & " OR (NuClien = 0 AND NuProve > 0)) "
        'SQLX2$ = SQLX2$ + "AND TipoMovim <> '" & MOBASI$ & "' "
        ' ESTA CONDICION SE PUSO PARA UNIFICAR LOS AJUSTES DE SALDO
        ' DISTINTOS GRABADOS CON EL MISMO NÚMERO DE AJUSTE
        If MOBASI$ <> "AJUD" Then SQLX2$ = SQLX2$ + "AND TipoMovim <> '" & MOBASI$ & "' "
        '
        .FindFirst SQLX2$
25      If .NoMatch = False Then
          '
          CUEII% = !CueContaI
          If CUEII% > 0 Then
            If TRIM$(ASIDIA$) <> "" Then
                .Edit
                !NuAsiDiar = ASIDIA$
                .Update
              Else
                IMASI#(CUEII%) = IMASI#(CUEII%) + ROUND(!IDEBECOMP) - ROUND(!IHABECOMP)
                SUMASI# = SUMASI# + ROUND(!IDEBECOMP) - ROUND(!IHABECOMP)
                If CUEII% > CUEMX% Then CUEMX% = CUEII%
            End If
          End If
          '
          .FindNext SQLX2$
          GoTo 25
        End If
        '
        If Abs(SUMASI#) >= 0.005 Then
          Call MENSERR(24, "Comprobante " + NUMEFA$ + " no Balancea.")
        End If
        '
        .Bookmark = MARCAREG
        .FindNext SQLX1$
        GoTo 15
      End If
    End With
    CABATEMP.Close
    '
    If HAYCAJA% < 1 Then
      FIN& = ULTREG&("ASIDIAR")
      TTXX$ = Chr$(1) + Chr$(1) + MKI$(HAS%)
      For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        ASX$ = REGLEIDO$("ASIDIAR", U&)
        If Left$(ASX$, 4) = TTXX$ Then
           CUEXI% = CVI(Mid$(ASX$, 5, 2))
           If CUEXI% > 0 Then
             IMASI#(CUEXI%) = CVD(Mid$(ASX$, 39, 8))
             If CUEXI% > CUEMX% Then CUEMX% = CUEXI%
           End If
        End If
      Next U&
    End If
    '
    TTXX$ = Chr$(1) + Chr$(1) + MKI$(HAS%)
    V& = 0: FIN& = ULTREG&("ASIDIAR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("ASIDIAR", U&)
      If Left$(X$, 4) <> TTXX$ Then
        V& = V& + 1
        Call GRAREG("ASIDIAR", X$, V&)
      End If
    Next U&
    '
    AGRECU% = 0
    For K% = 1 To CUEMX%
      If Abs(IMASI#(K%)) >= 0.005 Then
        X$ = REGBLAN$("ASIDIAR")
        Call REPLA(X$, TTXX$, 1, 4)
        Call REPLA(X$, MKI$(K%), 5, 2)
        Call REPLA(X$, MKD$(IMASI#(K%)), 39, 8)
        V& = V& + 1
        Call GRAREG("ASIDIAR", X$, V&)
      End If
    Next K%
    Call SETULTREG("ASIDIAR", V&)
    '
End Sub

Sub GENASICOB(ASIDIA$, DES%, HAS%)
    '
    Static DESDAT, HASDAT
    '
    Screen.MousePointer = 11
    '
    Dim IMASI#(32767)
    CUEMX% = 0
    '
    If DES% > 0 Then DESDAT = CDbl(CVDAT(DES%))
    If HAS% > 0 Then HASDAT = CDbl(CVDAT(HAS%))
    '
    If IsEmpty(DESDAT) = True Or IsNull(DESDAT) = True Or IsEmpty(HASDAT) = True Or IsNull(HASDAT) = True Then
      Call MENSERR(24, "Imposible Registrar")
      Exit Sub
    End If
    '
    Call ABRECABAN
    '\\\OT-05-0416-WAR-23/06/05///
    SQLX0$ = "SELECT * FROM CAJABANC "
    SQLX0$ = SQLX0$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX0$ = SQLX0$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    Set CABATEMP = CueCorri.OpenRecordset(SQLX0$, dbOpenDynaset)
    '\\\OT-05-0416-WAR-FIN///
    'SQLX$ = "SELECT * FROM CAJABANC "
    SQLX1$ = "INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX1$ = SQLX1$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX1$ = SQLX1$ + "AND TipoMovim = 'RCOB' "
    'SQLX$ = SQLX$ + "Order by FechConta ASC "
    '
    With CABATEMP 'CajaBanc
      .FindFirst SQLX1$
15    SUMASI# = 0
      Screen.MousePointer = 11
      If .NoMatch = False Then
        '
        MARCAREG = .Bookmark
        '
        NUMEFA$ = !CodiCom_Lar
        NCLIE% = !NuClien
        NPROE% = !NuProve
        FECHAD = Int(CDbl(!FechConta))
        MOBASI$ = !TipoMovim
        CUEII% = !CueContaI
        '
        If CUEII% > 0 Then
          If TRIM$(ASIDIA$) <> "" Then
              .Edit
              !NuAsiDiar = ASIDIA$
              .Update
            Else
              IMASI#(CUEII%) = IMASI#(CUEII%) + (!IDEBECOMP - !IHABECOMP)
              SUMASI# = SUMASI# + (!IDEBECOMP - !IHABECOMP)
              If CUEII% > CUEMX% Then CUEMX% = CUEII%
          End If
        End If
        '
        SQLX2$ = "INT(CDBL(FechConta)) = " & FECHAD & " "
        SQLX2$ = SQLX2$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
        SQLX2$ = SQLX2$ + "AND NuClien = " & NCLIE% & " "
        SQLX2$ = SQLX2$ + "AND NuProve = " & NPROE% & " "
        SQLX2$ = SQLX2$ + "AND TipoMovim <> '" & MOBASI$ & "' "
        '
        .FindFirst SQLX2$
25      If .NoMatch = False Then
          '
          CUEII% = !CueContaI
          If CUEII% > 0 Then
            If TRIM$(ASIDIA$) <> "" Then
                .Edit
                !NuAsiDiar = ASIDIA$
                .Update
              Else
                IMASI#(CUEII%) = IMASI#(CUEII%) + (!IDEBECOMP - !IHABECOMP)
                SUMASI# = SUMASI# + (!IDEBECOMP - !IHABECOMP)
                If CUEII% > CUEMX% Then CUEMX% = CUEII%
            End If
          End If
          '
          .FindNext SQLX2$
          GoTo 25
        End If
        '
        If Abs(SUMASI#) >= 0.005 Then
          Call MENSERR(24, "Comprobante " + NUMEFA$ + " no Balancea.")
        End If
        '
        .Bookmark = MARCAREG
        .FindNext SQLX1$
        GoTo 15
      End If
    End With
    '
    TTXX$ = Chr$(1) + Chr$(4) + MKI$(HAS%)
    V& = 0: FIN& = ULTREG&("ASIDIAR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("ASIDIAR", U&)
      If Left$(X$, 4) <> TTXX$ Then
        V& = V& + 1
        Call GRAREG("ASIDIAR", X$, V&)
      End If
    Next U&
    '
    AGRECU% = 0
    For K% = 1 To CUEMX%
      If Abs(IMASI#(K%)) >= 0.005 Then
        X$ = REGBLAN$("ASIDIAR")
        Call REPLA(X$, TTXX$, 1, 4)
        Call REPLA(X$, MKI$(K%), 5, 2)
        Call REPLA(X$, MKD$(IMASI#(K%)), 39, 8)
        V& = V& + 1
        Call GRAREG("ASIDIAR", X$, V&)
      End If
    Next K%
    Call SETULTREG("ASIDIAR", V&)
    '
End Sub

Sub GENASIPAG(ASIDIA$, DES%, HAS%)
    '
    Static DESDAT, HASDAT
    '
    Screen.MousePointer = 11
    '
    Dim IMASI#(32767)
    CUEMX% = 0
    '
    If DES% > 0 Then DESDAT = CDbl(CVDAT(DES%))
    If HAS% > 0 Then HASDAT = CDbl(CVDAT(HAS%))
    '
    If IsEmpty(DESDAT) = True Or IsNull(DESDAT) = True Or IsEmpty(HASDAT) = True Or IsNull(HASDAT) = True Then
      Call MENSERR(24, "Imposible Registrar")
      Exit Sub
    End If
    '
    Call ABRECABAN
    '\\\OT-05-0416-WAR-23/06/05///
    SQLX0$ = "SELECT * FROM CAJABANC "
    SQLX0$ = SQLX0$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX0$ = SQLX0$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX0$ = SQLX0$ + "ORDER BY CodiCom_Lar ASC "
    Set CABATEMP = CueCorri.OpenRecordset(SQLX0$, dbOpenDynaset)
    '\\\OT-05-0416-WAR-FIN///
    'SQLX$ = "SELECT * FROM CAJABANC "
    SQLX1$ = "INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX1$ = SQLX1$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX1$ = SQLX1$ + "AND TipoMovim = 'OPAG' "
    'SQLX$ = SQLX$ + "Order by FechConta ASC "
    '
    With CABATEMP 'CajaBanc
      .FindFirst SQLX1$
15    SUMASI# = 0
      Screen.MousePointer = 11
      If .NoMatch = False Then
        '
        If !CodiCom_Lar = NUMEFA$ Then GoTo 29
        
        MARCAREG = .Bookmark
        '
        NUMEFA$ = !CodiCom_Lar
        NCLIE% = !NuClien
        NPROE% = !NuProve
        FECHAD = Int(CDbl(!FechConta))
        MOBASI$ = !TipoMovim
        OPCIMOV$ = !OPCIMOVIM
        CUEII% = !CueContaI
        '
        If CUEII% = 26 Then
        AA = BB
        End If
        If CUEII% > 0 Then
          If TRIM$(ASIDIA$) <> "" Then
              .Edit
              !NuAsiDiar = ASIDIA$
              .Update
            Else
'              IMASI#(CUEII%) = IMASI#(CUEII%) + (!IDEBECOMP - !IHABECOMP)
'              SUMASI# = SUMASI# + (!IDEBECOMP - !IHABECOMP)
'              If CUEII% > CUEMX% Then CUEMX% = CUEII%
          End If
        End If
        '
        SQLX2$ = "INT(CDBL(FechConta)) = " & FECHAD & " "
        SQLX2$ = SQLX2$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
        SQLX2$ = SQLX2$ + "AND NuClien = " & NCLIE% & " "
        SQLX2$ = SQLX2$ + "AND NuProve = " & NPROE% & " "
        If OPCIMOV$ = "PC" Then SQLX2$ = SQLX2$ + " AND OPCIMOVIM<>'FC' AND OPCIMOVIM<>'TF' AND OPCIMOVIM<>'CrFi' AND OPCIMOVIM<>'IB-C'"
        'SQLX2$ = SQLX2$ + "AND TipoMovim <> '" & MOBASI$ & "' "
        '
        .FindFirst SQLX2$
25      If .NoMatch = False Then
          '
          CUEII% = !CueContaI
          If CUEII% > 0 Then
            If TRIM$(ASIDIA$) <> "" Then
                .Edit
                !NuAsiDiar = ASIDIA$
                .Update
              Else
                IMASI#(CUEII%) = IMASI#(CUEII%) + (!IDEBECOMP - !IHABECOMP)
                SUMASI# = SUMASI# + (!IDEBECOMP - !IHABECOMP)
                If CUEII% > CUEMX% Then CUEMX% = CUEII%
            End If
          End If
          '
          .FindNext SQLX2$
          GoTo 25
        End If
        '
        If Abs(SUMASI#) >= 0.005 Then
          Call MENSERR(24, "Comprobante " + NUMEFA$ + " no Balancea.")
        End If
        '
        .Bookmark = MARCAREG
29      .FindNext SQLX1$
        GoTo 15
      End If
    End With
    '
    TTXX$ = Chr$(1) + Chr$(5) + MKI$(HAS%)
    V& = 0: FIN& = ULTREG&("ASIDIAR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("ASIDIAR", U&)
      If Left$(X$, 4) <> TTXX$ Then
        V& = V& + 1
        Call GRAREG("ASIDIAR", X$, V&)
      End If
    Next U&
    '
    AGRECU% = 0
    For K% = 1 To CUEMX%
      If Abs(IMASI#(K%)) >= 0.005 Then
        X$ = REGBLAN$("ASIDIAR")
        Call REPLA(X$, TTXX$, 1, 4)
        Call REPLA(X$, MKI$(K%), 5, 2)
        Call REPLA(X$, MKD$(IMASI#(K%)), 39, 8)
        V& = V& + 1
        Call GRAREG("ASIDIAR", X$, V&)
      End If
    Next K%
    Call SETULTREG("ASIDIAR", V&)
    '
End Sub

Sub GENASICABA(ASIDIA$, DES%, HAS%)
    '
    Static DESDAT, HASDAT
    '
    Screen.MousePointer = 11
    '
    Dim IMASI#(32767)
    CUEMX% = 0
    '
    If DES% > 0 Then DESDAT = CDbl(CVDAT(DES%))
    If HAS% > 0 Then HASDAT = CDbl(CVDAT(HAS%))
    '
    If IsEmpty(DESDAT) = True Or IsNull(DESDAT) = True Or IsEmpty(HASDAT) = True Or IsNull(HASDAT) = True Then
      Call MENSERR(24, "Imposible Registrar")
      Exit Sub
    End If
    '
End Sub

Sub GENASIPAGVIE(ASIDIA$, DES%, HAS%)
    '
    Static DESDAT, HASDAT
    '
    Screen.MousePointer = 11
    '
    Dim IMASI#(32767)
    CUEMX% = 0
    '
    If DES% > 0 Then DESDAT = CDbl(CVDAT(DES%))
    If HAS% > 0 Then HASDAT = CDbl(CVDAT(HAS%))
    '
    Call ABRECABAN
    'SQLX$ = "SELECT * FROM CAJABANC "
    SQLX1$ = SQLX1$ + "INT(CDBL(FechConta)) >= " & DESDAT & " "
    SQLX1$ = SQLX1$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    SQLX1$ = SQLX1$ + "AND TipoMovim = 'OPAG' "
    'SQLX$ = SQLX$ + "Order by FechConta ASC "
    '
    With CajaBanc
      .FindFirst SQLX1$
15    SUMASI# = 0
      Screen.MousePointer = 11
      If .NoMatch = False Then
        '
        MARCAREG = .Bookmark
        '
        NUMEFA$ = !CodiCom_Lar
        NCLIE% = !NuClien
        NPROE% = !NuProve
        FECHAD = Int(CDbl(!FechConta))
        MOBASI$ = !TipoMovim
        CUEII% = !CueContaI
        '
        If CUEII% > 0 Then
          If TRIM$(ASIDIA$) <> "" Then
              .Edit
              !NuAsiDiar = ASIDIA$
              .Update
            Else
              IMASI#(CUEII%) = IMASI#(CUEII%) + (!IDEBECOMP - !IHABECOMP)
              SUMASI# = SUMASI# + (!IDEBECOMP - !IHABECOMP)
              If CUEII% > CUEMX% Then CUEMX% = CUEII%
          End If
        End If
        '
        SQLX2$ = "INT(CDBL(FechConta)) = " & FECHAD & " "
        SQLX2$ = SQLX2$ + "AND CodiCom_Lar = '" & NUMEFA$ & "' "
        'SQLX2$ = SQLX2$ + "AND NuClien = " & NCLIE% & " "
        SQLX2$ = SQLX2$ + "AND NuProve = " & NPROE% & " "
        SQLX2$ = SQLX2$ + "AND TipoMovim <> '" & MOBASI$ & "' "
        '
        .FindFirst SQLX2$
25      If .NoMatch = False Then
          '
          CUEII% = !CueContaI
          If CUEII% > 0 Then
            If TRIM$(ASIDIA$) <> "" Then
                .Edit
                !NuAsiDiar = ASIDIA$
                .Update
              Else
                IMASI#(CUEII%) = IMASI#(CUEII%) + (!IDEBECOMP - !IHABECOMP)
                SUMASI# = SUMASI# + (!IDEBECOMP - !IHABECOMP)
                If CUEII% > CUEMX% Then CUEMX% = CUEII%
            End If
          End If
          '
          .FindNext SQLX2$
          GoTo 25
        End If
        '
        If Abs(SUMASI#) >= 0.005 Then
          Call MENSERR(24, "Comprobante " + NUMEFA$ + " no Balancea.")
        End If
        '
        .Bookmark = MARCAREG
        .FindNext SQLX1$
        GoTo 15
      End If
    End With
    '
    TTXX$ = Chr$(1) + Chr$(5) + MKI$(HAS%)
    V& = 0: FIN& = ULTREG&("ASIDIAR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("ASIDIAR", U&)
      If Left$(X$, 4) <> TTXX$ Then
        V& = V& + 1
        Call GRAREG("ASIDIAR", X$, V&)
      End If
    Next U&
    '
    AGRECU% = 0
    For K% = 1 To CUEMX%
      If Abs(IMASI#(K%)) >= 0.005 Then
        X$ = REGBLAN$("ASIDIAR")
        Call REPLA(X$, TTXX$, 1, 4)
        Call REPLA(X$, MKI$(K%), 5, 2)
        Call REPLA(X$, MKD$(IMASI#(K%)), 39, 8)
        V& = V& + 1
        Call GRAREG("ASIDIAR", X$, V&)
      End If
    Next K%
    Call SETULTREG("ASIDIAR", V&)
    '
End Sub


