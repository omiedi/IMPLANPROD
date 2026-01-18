VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CIERRAOF14 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre de Orden de Fabricación con Trazabilidad"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   13080
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      BackColor       =   &H00CDDADA&
      Caption         =   "# "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   9080
      TabIndex        =   47
      Top             =   720
      Width           =   3060
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
         Height          =   285
         Left            =   630
         TabIndex        =   49
         Top             =   210
         Width           =   810
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
         Height          =   285
         Left            =   2100
         TabIndex        =   48
         Top             =   210
         Width           =   810
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PNL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -750
         TabIndex        =   51
         Top             =   260
         Width           =   1335
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PLC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   700
         TabIndex        =   50
         Top             =   260
         Width           =   1335
      End
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
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
      Left            =   945
      TabIndex        =   41
      Top             =   630
      Width           =   435
   End
   Begin VB.CommandButton Command7 
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
      Left            =   1365
      TabIndex        =   40
      Top             =   630
      Width           =   175
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
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
      Left            =   5175
      TabIndex        =   39
      Top             =   630
      Width           =   435
   End
   Begin VB.CommandButton Command9 
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
      Left            =   5595
      TabIndex        =   38
      Top             =   630
      Width           =   175
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
      Height          =   285
      Left            =   11115
      TabIndex        =   35
      Top             =   210
      Width           =   960
   End
   Begin VB.CommandButton Command4 
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
      Left            =   9915
      TabIndex        =   32
      Top             =   210
      Width           =   175
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      Caption         =   "DATOS DE LA O-F ACTIVA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5950
      Left            =   105
      TabIndex        =   6
      Top             =   1365
      Width           =   12030
      Begin MSFlexGridLib.MSFlexGrid mfgReserva 
         Height          =   4215
         Left            =   105
         TabIndex        =   55
         Top             =   1680
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   7435
         _Version        =   393216
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Roboto Mono Medium"
            Size            =   8.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00CDDADA&
         BorderStyle     =   0  'None
         Caption         =   "Materiales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4200
         Left            =   105
         TabIndex        =   24
         Top             =   1640
         Width           =   9360
         Begin VB.Frame Frame9 
            Caption         =   "Frame4"
            Height          =   4050
            Left            =   8700
            TabIndex        =   52
            Top             =   120
            Width           =   15
         End
         Begin VB.Frame Frame11 
            Caption         =   "Frame3"
            Height          =   4050
            Left            =   6510
            TabIndex        =   29
            Top             =   120
            Width           =   15
         End
         Begin VB.Frame Frame7 
            Caption         =   "Frame3"
            Height          =   4050
            Left            =   1680
            TabIndex        =   28
            Top             =   120
            Width           =   15
         End
         Begin VB.Frame Frame4 
            Caption         =   "Frame4"
            Height          =   4050
            Left            =   7665
            TabIndex        =   27
            Top             =   120
            Width           =   15
         End
         Begin VB.Frame Frame5 
            Caption         =   "Frame5"
            Height          =   4050
            Left            =   4620
            TabIndex        =   26
            Top             =   120
            Width           =   15
         End
         Begin VB.Frame Frame6 
            Caption         =   "Frame6"
            Height          =   4050
            Left            =   5565
            TabIndex        =   25
            Top             =   120
            Width           =   15
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Roboto Mono Medium"
               Size            =   8.25
               Charset         =   0
               Weight          =   500
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3855
            ItemData        =   "CIERRAOF14.frx":0000
            Left            =   0
            List            =   "CIERRAOF14.frx":0002
            TabIndex        =   30
            Top             =   350
            Width           =   9360
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Codigo          Descripción                 Uso Un.   Neces.    Trazado     Falta  Dp    "
            BeginProperty Font 
               Name            =   "Roboto Mono Medium"
               Size            =   8.25
               Charset         =   0
               Weight          =   500
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   0
            TabIndex        =   31
            Top             =   105
            Width           =   9360
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00CDDADA&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   50
         TabIndex        =   7
         Top             =   240
         Width           =   11850
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. O/F"
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
            TabIndex        =   22
            Top             =   270
            Width           =   960
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Height          =   255
            Left            =   1155
            TabIndex        =   21
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
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
            Left            =   4815
            TabIndex        =   20
            Top             =   270
            Width           =   960
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Height          =   255
            Left            =   5865
            TabIndex        =   19
            Top             =   210
            Width           =   330
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Dias Atraso"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   9960
            TabIndex        =   18
            Top             =   1080
            Width           =   960
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 12"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   465
            Left            =   11010
            TabIndex        =   17
            Top             =   810
            Width           =   750
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Solicitada"
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
            Left            =   8805
            TabIndex        =   16
            Top             =   240
            Width           =   1905
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Height          =   255
            Left            =   10800
            TabIndex        =   15
            Top             =   180
            Width           =   960
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Producto"
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
            TabIndex        =   14
            Top             =   570
            Width           =   960
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Height          =   255
            Left            =   1155
            TabIndex        =   13
            Top             =   525
            Width           =   2010
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Height          =   465
            Left            =   1155
            TabIndex        =   12
            Top             =   840
            Width           =   7080
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad"
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
            Left            =   5595
            TabIndex        =   11
            Top             =   270
            Width           =   1485
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Height          =   255
            Left            =   7170
            TabIndex        =   10
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo"
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
            Left            =   5595
            TabIndex        =   9
            Top             =   600
            Width           =   1485
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Height          =   255
            Left            =   7170
            TabIndex        =   8
            Top             =   525
            Width           =   1065
         End
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2730
      TabIndex        =   3
      Top             =   210
      Width           =   175
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   8460
      Left            =   12240
      ScaleHeight     =   8430
      ScaleWidth      =   1350
      TabIndex        =   2
      Top             =   0
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   53
         Top             =   6600
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   54
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Kdx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   105
         Picture         =   "CIERRAOF14.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   2320
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
         Height          =   640
         Left            =   105
         Picture         =   "CIERRAOF14.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Acceso a Trazabilidad de Reservas"
         Top             =   1680
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   105
         Picture         =   "CIERRAOF14.frx":0458
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Aprueba Cierre de Orden de Fabricación"
         Top             =   5355
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
         Picture         =   "CIERRAOF14.frx":0762
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Salir"
         Top             =   200
         Width           =   650
      End
      Begin VB.CommandButton Command29 
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
         Picture         =   "CIERRAOF14.frx":08AC
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   845
         Width           =   650
      End
   End
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   945
      TabIndex        =   45
      Top             =   960
      Width           =   2445
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Entrada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   -30
      TabIndex        =   44
      Top             =   710
      Width           =   915
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5175
      TabIndex        =   43
      Top             =   945
      Width           =   2310
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Consumo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3990
      TabIndex        =   42
      Top             =   705
      Width           =   1125
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9920
      TabIndex        =   36
      Top             =   285
      Width           =   1125
   End
   Begin VB.Label Label24 
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
      Left            =   8970
      TabIndex        =   34
      Top             =   210
      Width           =   1020
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Cierre"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   7230
      TabIndex        =   33
      Top             =   300
      Width           =   1650
   End
   Begin VB.Label Label117 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "O/F Nro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -440
      TabIndex        =   5
      Top             =   300
      Width           =   1335
   End
   Begin VB.Label Label144 
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
      Left            =   960
      TabIndex        =   4
      Top             =   210
      Width           =   1770
   End
   Begin VB.Menu Transacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu DesvioPNLPLCvsTrazado 
         Caption         =   "Desvío de Consumo (PNL-PLC vs Trazado)"
      End
      Begin VB.Menu DevioConsumoOfCerrada 
         Caption         =   "Scrap de Proceso (OF Cerradas)"
      End
   End
   Begin VB.Menu Calidad 
      Caption         =   "Calidad"
      Begin VB.Menu ControlCalidadComponentesOfActiva 
         Caption         =   "Control de Calidad de Componentes (OF Activa)"
      End
   End
   Begin VB.Menu ConsultayListados 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu InformeTecnico 
         Caption         =   "Informe Tecnico en Word"
      End
      Begin VB.Menu mnuReportesVarios 
         Caption         =   "Reportes Varios"
      End
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuracion"
      Begin VB.Menu mnuAsociarReporte 
         Caption         =   "Vincular Reportes al Formulario"
      End
   End
End
Attribute VB_Name = "CIERRAOF14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ExigeCalidadCompleto As Boolean

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    Call TRAZORF14.Command10_Click
    Command10.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    '
    If XVALO(Text3) = 99 Then
       If COMALTER%("Atención: Verifique Depósitos Predeterminados de Consumo !\\Continuar\Cancelar") <> 1 Then
          GoTo 99
       End If
    End If
    '
    Dim CODIN%(4096), LOTRE$(4096), FEALT%(4096), SALFIS(4096), PP1$(4096), PP2$(4096)
    Dim CANECE(32767), CANECE2(32767)
    '
    HOII% = HOY(HO$)
    FECHA% = FECHANUM(Label24)
    If FECHA% = 0 Or FECHA% > HOII% Then
       Call MENSERR(24, "Fecha No Válida")
       GoTo 99
    End If
     '
    If XVALO(Text1) < 1 Then
       Call MENSERR(24, "Imposible Cerrar - Falta Cantidad")
       GoTo 99
    End If
    '
    DEPO% = XVALO(Text2)
    If DEPO% < 1 Or TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPO%))) = "" Then
       Call MENSERR(24, "Falta Depósito de Entrada")
       GoTo 99
    End If
    '
    DP2% = XVALO(Text3)
    If DP2% < 1 Or TRIM$(ECOARCH$("ECODEP", Chr$(DP2%))) = "" Then
       Call MENSERR(24, "Falta Depósito de Consumo")
       GoTo 99
    End If
    '
    For KKU% = 1 To List1.ListCount
      If XVALO(Mid$(List1.List(KKU% - 1), 72, 10)) >= 0.05 Then
        Call MENSERR(24, "Cierre de O/F Imposible.\   \Stock Trazado Insuficiente\en Código '" + TRIM$(Left$(List1.List(KKU% - 1), 16)) + "'.")
        GoTo 99
      End If
      DPXX1% = XVALO(Mid$(List1.List(KKU% - 1), 82, 5))
      If DPXX1% < 1 Or DPXX1% > 255 Then
        Call MENSERR(24, "Depósito de Consumo Inexistente o No Válido\en Código '" + TRIM$(Left$(List1.List(KKU% - 1), 16)) + "'.")
        GoTo 99
      End If
    Next KKU%
    '
    Call BLOQARCH("RESERVOF")
    ' VERIFICAR NEGATIVOS
    CAITMO% = 0
    For KKI% = 1 To NUMAS%
       CANECE(KKI%) = 0
    Next KKI%
    '
    For KKI% = 0 To List1.ListCount
       TTXX$ = List1.List(KKI%)
       code$ = TRIM$(Left$(TTXX$, 16))
       CIXI% = CODINT%(code$)
       CACONSU = XVALO(Mid$(TTXX$, 52, 10))
       If CACONSU > 0.005 Then
          CAITMO% = CAITMO% + 1
          CIMOVIM%(CAITMO%) = CIXI%
          CAMOVIM(CAITMO) = (-1) * CACONSU
          CANECE(CIXI%) = CANECE(CIXI%) + CACONSU
          
'          ' SOLO CONTROLA POR PNL Y PLC, AQUELLOS COMPONENTES SOBRE LOS CUALES SE INFORMARON DICHAS CANTIDADES
'          CantidadTrazadaReal# = XVALO(mfgReserva.TextMatrix(KKI% + 1, 7))
'          If CantidadTrazadaReal# > 0.05 Then
'            CANECE2(CIXI%) = CANECE2(CIXI%) + CACONSU - CantidadTrazadaReal#
'          End If
       End If
    Next KKI%
    '
'    ' VERIFICA QUE LA CANTIDAD INFORMADA EN PNL Y PLC, SEA SUFICIENTE PARA LA NECESIDAD DEL COMPONENTE EN ESTE CIERRE
'    For KKI% = 1 To NUMAS%
'       If CANECE2(KKI%) >= 0.05 Then
'          Call MENSERR(24, "Cierre de O/F Imposible.\ \Diferencia entre Lotes Trazados\en Código '" + CODEXT$(KKI%) + "'\y los Valores de PNL y PLC")
'          GoTo 99
'       End If
'    Next KKI%
    '
    If PUEDEMOVI%(DP2%) < 0 Then
       Call COMUNI("Transacción Cancelada por Control de Negativos.")
       GoTo 99
    End If
    '
    ' CARGAR UN VECTOR CON LOS LOTES DISPONIBLES
    NORFA$ = TRIM$(Label2)
    Call APERBASDAT("RESERVA")
    CONDI$ = "IdSubOr = '" & NORFA$ & "' AND CantRes - CantCons > 0 "
    FIN& = CantRegistros&("RESERVA", CONDI$, "NOMESS")
    
    SQLX$ = " SELECT * FROM RESERVA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    'SQLX$ = SQLX$ + " AND CantRes - CantCons > 0 "
    'Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    Set RESERTMP = READSET(SQLX$)
    k& = 0
    With RESERTMP
       If .EOF Then
          OPXX% = COMALTER%("No Existen Reservas Vigentes\para la O/F Elegida.\  \¿ Realmente Confirma el Cierre\Sin Consumo de Materiales Componentes ?\\No, Cancelar Cierre\Si Confirmarlo")
          If OPXX% = 2 Then GoTo 20   ' CONFIRMA CIERRE
          GoTo 99
       End If
'       FIN& = .RecordCount
'       .MoveFirst
'       On Error GoTo 0
       JJ& = 0
       Do While (.EOF = False)
          k& = k& + 1
          Call TRACE(20, k&, FIN&)
          '
          CIXI% = XVALO(!Cod_Inte)
          IDLO$ = TRIM$(SAFETEXT$(!IDENLOTE))
          If IDLO$ <> "" Then
            If TRAZORF14.TieneRegistroCalidad(IDLO$, observaciones$) = False And ExigeCalidadCompleto = True Then
                Call COMUNI("Imposible Cerrar O-Fabricación\ \Lote '" & IDLO$ & "'\con Registro de Calidad Incompleto.\ \Corrija desde el Menu 'Calidad'")
                GoTo 99
            End If
            
            CANRES# = XVALO(!CANTRES) - XVALO(!CantCons)
            JJ& = JJ& + 1
            CODIN%(JJ&) = CIXI%
            LOTRE$(JJ&) = IDLO$
              Call ACSALOTE(CIXI%, IDLO$, "NOMESS")         ' ACTUALIZA SALDO DEL LOTE
            FEALT%(JJ&) = FECREALOT%(IDLO$)
              SALNOCO# = SALNOCONS(IDLO$, CIXI%)
              If CANRES# > SALNOCO# + 0.005 Then
                 Call MENSERR(24, "Error de Trazado en Código '" + CODEXT$(CIXI%) + "'.\Imposible Cerrar esta O/F.\   \Vuelva a Trazar las Reservas y Re-Intente.")
                 GoTo 99
              End If
            SALFIS(JJ&) = CANRES#                           ' CANTIDAD RESERVA RESTANTE (MENOS LO CONSUMIDO)
            PP1$(JJ&) = SAFETEXT$(!PNL_DES)
            PP2$(JJ&) = SAFETEXT$(!PLC_HAS)
            CANECE(CIXI%) = CANECE(CIXI%) - SALNOCO#        ' NECESIDAD MENOS EL SALDO DEL LOTE RESERVADO
          End If
19     .MoveNext
       Loop
    End With
    '
20  Screen.MousePointer = 1
    '
    ' VERIFICA SI LA CANTIDAD NECESARIA PARA EL CIERRE FUE CUBIERTA POR LOS LOTES RESERVADOS PARA CADA COMPONENTE
    For KKI% = 1 To NUMAS%
       If CANECE(KKI%) >= 0.05 Then
          Call MENSERR(24, "Cierre de O/F Imposible.\Error de Lotes Trazados en\Código '" + CODEXT$(KKI%) + "'.\   \Revise Lotes y Situación de Stocks.")
          GoTo 99
       End If
    Next KKI%
    '
    ' ordenar por fecha de alta
    For U1& = 1 To JJ&
      For U2& = U1& + 1 To JJ&
        If FEALT%(U1&) > FEALT%(U2&) Then
          SW = SWAP(FEALT%(U1&), FEALT%(U2&))
          SW = SWAP(CODIN%(U1&), CODIN%(U2&))
          SW = SWAP(LOTRE$(U1&), LOTRE$(U2&))
          SW = SWAP(SALFIS(U1&), SALFIS(U2&))
          SW = SWAP(PP1$(U1&), PP1$(U2&))
          SW = SWAP(PP2$(U1&), PP2$(U2&))
        End If
      Next U2&
    Next U1&
    '
    Call APERBASDAT("STOCKS")
    Call APERBASDAT("RESERVA")
    '
    CMOV$ = "CF"
    CODICON$ = Label10
    DOPRI$ = TRIM$(Label2)
    DOSEC$ = "PF-" + TRIM$(Mid$(Label2, 4))
    REFE$ = "CONSUMO " + CODICON$
    NORIT% = 0
    '
    ' GRABAR MOVIMIENTO DE ENTRADA
    CICON% = CODINT%(CODICON$)
    CODCON$ = CODEXT$(CICON%)
    ILOTE$ = Label2
       PDES$ = Text5
       PHAS$ = Text6
    CMOV$ = "OF"
    DOPRI$ = TRIM$(Label2)
    DOSEC$ = "IF-" + TRIM$(Mid$(DOPRI$, 4))
    REFE$ = "Fabricación Propia " + CODICON$
    VUNI# = 0
    MONEX% = 1
    CANPRO = XVALO(Text1)
    GoSub GRAPRODUC
    '
    ' GRABAR CONSUMO Y CANCELA RESERVA
    SQLT$ = "SELECT * FROM MOVISTO "
    SQLT$ = SQLT$ + " WHERE CodiMovi = '" & CMOV$ & "' "
    SQLT$ = SQLT$ + " AND DoPriCor = '" & TRIM$(DOPRI$) & "' "
    'SQLT$ = SQLT$ + " AND IdenMovi = '" & IDMO$ & "' "
    'Set MACONSUM = Stocks.OpenRecordset(SQLT$, dbOpenDynaset)
    Dim MACONSUM As New ADODB.Recordset
25  On Error Resume Next
    MACONSUM.Open FILTSQL$(SQLT$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    '
    For KKI% = 0 To List1.ListCount
       TTXX$ = List1.List(KKI%)
       code$ = TRIM$(Left$(TTXX$, 16))
       CIXI% = CODINT%(code$)
       CACONSU = XVALO(Mid$(TTXX$, 52, 10))
       DPXX1% = XVALO(Mid$(TTXX$, 82, 5))
       '
       Do While CACONSU >= 0.05
          HUBOCON% = 0
          For KKJ& = 1 To JJ&
            If CODIN%(KKJ&) = CIXI% Then
               CACON = CACONSU
               If CACON > SALFIS(KKJ&) Then CACON = SALFIS(KKJ&)
               If CACON > 0 Then
                  LOTE$ = LOTRE$(KKJ&)
                  CANTIX = CACON
                      PCOM$ = PP1$(KKJ&)
                      PFIN$ = PP2$(KKJ&)
                  NORIT% = NORIT% + 1
                  CMOV$ = "CF"
                  GoSub GRACONSUM
                  GoSub CANCELARES
                  Call ACSALOTE(CIXI%, LOTE$, "NOMESS")
                  SALFIS(KKJ&) = SALFIS(KKJ&) - CACON
                  CACONSU = CACONSU - CACON
                  HUBOCON% = 1
               End If
            End If
          Next KKJ&
          If HUBOCON% < 1 Then
             Call MENSERR(24, "Imposible Completar Cierre Parcial\de O/F '" + NORFA$ + "'")
             GoTo 99
          End If
       Loop
       '
    Next KKI%
    '
    MACONSUM.Close
    Set MACONSUM = Nothing
    
    ' ACTUALIZAR PENDIENTE DE O/F
    GoSub ACUPENOF
    If CIERRA% = 1 Then Call ANULARES(NORFA$)
    Label144 = ""
    GoTo 99

'*********************************************************************************
    '
ACUPENOF:
    ACUPRODU = 0
    Call APERBASDAT("STOCKS")
    SQLX$ = "SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE CodiMovi = 'OF' "
    SQLX$ = SQLX$ + " AND DoPriCor = '" & NORFA$ & "' "
    SQLX$ = SQLX$ + " AND Cod_Inte = " & CICON% & " "
    'Set ORYACIE = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    Dim ORYACIE As New ADODB.Recordset
35  On Error Resume Next
    ORYACIE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 35
    End If
    On Error GoTo 0
    With ORYACIE
         Do While Not .EOF
            ACUPRODU = ACUPRODU + XVALO(!CantIngre)
         .MoveNext
         Loop
    End With
    '
    CIERRA% = 0
    If ACUPRODU >= XVALO(Label13) - 0.05 Then
        CIERRA% = 1
      Else
        MDXX% = COMALTER%("ATENCION: Esta O/F queda con\un Saldo Pendiente.\\Mantener Abierta\Cerrar")
        If MDXX% = 2 Then CIERRA% = 1
    End If
    '
    Call APERBASDAT("ORFAB")
    SQLX$ = "SELECT CanApro,CanRech,Statuorf,FechCierr FROM OrdeFabr  WHERE IdSubOr = '" & NORFA$ & "'"
45  On Error Resume Next
    Dim rs As New ADODB.Recordset
    rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 45
    End If
    On Error GoTo 0
    With rs
          !CanApro = ACUPRODU
          !CanRech = XVALO(!CanRech) + SCRAP
          If CIERRA% = 1 Then
             !Statuorf = 3         ' MARCA COMO CERRADA
           Else
             If XVALO(!Statuorf) > 2 Then !Statuorf = 1
         End If
         !FechCierr = CVDAT(FECHANUM(Label24))
         .Update
    End With
    rs.Close
    Set rs = Nothing

'''    With OrdeFabr
'''       .FindFirst "IdSubOr = '" & NORFA$ & "'"
'''       If .NoMatch = False Then
'''         .Edit
'''         !CanApro = ACUPRODU
'''         !CanRech = !CanRech + SCRAP
'''         If CIERRA% = 1 Then
'''             !Statuorf = 3         ' MARCA COMO CERRADA
'''           Else
'''             If !Statuorf > 2 Then !Statuorf = 1
'''         End If
'''         !FechCierr = CVDAT(FECHANUM(Label24))
'''         .Update
'''       End If
'''    End With
    '
    '25/04/07 (OT-07-0161)'PARA GUARDAR EL DEPOSITO DE ENTRADA Y EL DEPOSITO
    If CICON% > 0 Then    'DE CONSUMO EN EL MAESTRO DE ARTICULOS
        CONDI$ = "CODINT = " & CICON%
        
        Call ACTUREGISTRO("MASTER", "DEPOENTRA", CONDI$, XVALO(Text2), REGAF&, "NOMESS")
        Call ACTUREGISTRO("MASTER", "DEPOCON", CONDI$, XVALO(Text3), REGAF&, "NOMESS")
    End If
    '
Return
    '
GRAPRODUC:
    SQLX$ = "SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE (CodiMovi = 'OF') "
    SQLX$ = SQLX$ + " AND DoPriCor = '" & TRIM$(NORFA$) & "'"
    'SQLX$ = SQLX$ + " AND IdenMovi = '" & IDMO$ & "'"
    'Set PRODORFA = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    '
'55  On Error Resume Next
    'Dim PRODORFA As ADODB.Recordset
    Dim PRODORFA As New ADODB.Recordset
    PRODORFA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'       On Error GoTo 0
'       Call CapturaErrorOpen
'       GoTo 55
'    End If
'    On Error GoTo 0
    With PRODORFA
      .AddNew
      !Fechmov = CVDAT(FECHA%)
      !Cod_Inte = CICON%
      !COD_EXTE = CODCON$
      !DESCRIP = DESCRIT0$(CICON%)
      !IDENLOTE = TRIM$(ILOTE$)
      !PNL_DES = PDES$
      !PLC_HAS = PHAS$
      !CodiMovi = CMOV$
      !DoPriCor = DOPRI$
      !DoPriLar = DOPRI$
      !DoSecCor = DOSEC$
      !DoSecLar = DOSEC$
      !IdenMovi = IDMO$
      !NuOrItem = 0
      !NUPEDCLI = 0
      !ReferCor = REFE$
      !ValoUnit = VUNI#
      !MoneVal = MONEX%
      !TIPOCAM = TICAMONE(MONEX%)
         DEPOI% = DEPO%: If DEPO% = 99 Then DEPOI% = DEPOPRE%(CICON%)
      !DepoMovi = DEPOI%
      !Nume_Clie = 0
      !Nume_Prov = 0
      !CantIngre = CANPRO
      !Cantsalid = 0
      !CantSaldo = 0
      !FeReMovi = Now
      !NumUsuar = USERN%
      .Update
    End With
    '
    'Call CREALOTIN(FECHA%, CICON%, ILOTE$, CDbl(CANPRO))
    Call CREALOTIN(FECHA%, 0, 0, CICON%, ILOTE$, "", "", XVALO(CANPRO))
    '
    PRODORFA.Close
    Set PRODORFA = Nothing
Return
'
'
GRACONSUM:
    With MACONSUM
      .AddNew
      !Fechmov = CVDAT(FECHA%)
      !Cod_Inte = CIXI%
      !COD_EXTE = code$
      !DESCRIP = DESCRIT0$(CIXI%)
      !IDENLOTE = TRIM$(LOTE$)
      !PNL_DES = PCOM$
      !PLC_HAS = PFIN$
      !CodiMovi = CMOV$
      !DoPriCor = DOPRI$
      !DoPriLar = DOPRI$
      !DoSecCor = DOSEC$
      !DoSecLar = DOSEC$
      !IdenMovi = IDMO$
         LOTECON$ = AJUSTI$(LOTE$, 24) + FORMATNUM$(CANTIX, "F8.1")
      !LotConsum = LOTECON$
      !NuOrItem = NORIT%
      !NUPEDCLI = 0
      !ReferCor = REFE$
      !ValoUnit = 0
      !MoneVal = 1
      !TIPOCAM = TICAMONE(1)
         DEPOI% = DPXX1%: If DPXX1% = 99 Then DEPOI% = DEPOPRE%(CIXI%)
      !DepoMovi = DEPOI%
      !Nume_Clie = 0
      !Nume_Prov = 0
      !CantIngre = 0
      !Cantsalid = Abs(CANTIX)
      !CantSaldo = 0
      !FeReMovi = Now
      !NumUsuar = USERN%
      .Update
      '
    End With
Return
'
CANCELARES:
    CANTLIM = CANTIX - 0.05
    SQLX$ = " SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "'"
    SQLX$ = SQLX$ + "AND IdenLote = '" & LOTE$ & "' "
    SQLX$ = SQLX$ + " AND CantRes - CantCons >= " & TRIM$(FORMATNUM$(CANTLIM, "F12.2"))
    '
    'Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    
65  On Error Resume Next
    Dim RESERTMP2 As New ADODB.Recordset
    RESERTMP2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 65
    End If
    On Error GoTo 0
    With RESERTMP2
'       On Error Resume Next
'       .MoveFirst
       If .EOF Then
          Call MENSERR(24, "No Existen Reservas Vigentes\para la O/F Elegida. ")
          GoTo 99
       End If
       !CANTRES = XVALO(!CANTRES) - CANTIX
    
'       PLCHasta& = XVALO(!PLC_HAS)
'       If PLCHasta& > 0 Then
'            !PNL_DES = 1
'            !PLC_HAS = SAFETEXT$(!CANTRES)
'       End If
       
       .Update
       
       If XVALO(!CANTRES) < 0.05 Then
          .Delete
       End If
    End With
    '
    RESERTMP2.Close
    Set RESERTMP2 = Nothing
Return
'
'
99 Call LIBARCH("RESERVOF")
   Command2.Enabled = True
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command29_Click()
  '
  Call OPNOIN("Ayuda FLEXOFT en Línea")
  '
End Sub

Private Sub Command3_Click()
   '
   Call ACTECOR(2)
   NROOF$ = ""
   If ULTREG&("ECORDEP") < 1 Then
     Call MENSERR(24, "No Hay Ordenes de Fabricación en Curso\para el Código Seleccionado")
     GoTo 90
   End If
   '
   Call SELECHO("ECORDEP/Ordenes de Trabajo en Curso")
   NROOF$ = AJUSTI$(VALACT1$("ECORDEP"), 12)
   '
90 Label144 = NROOF$
   Screen.MousePointer = 1
   '
End Sub


Private Sub Command4_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = TRIM$(VALACT1$("SELFECHA"))
       If VDEV$ <> "" Then
       Label24 = VDEV$
    End If
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   TRAZORF14.Label32 = Label144
   TRAZORF14.Show 1
   Call Text1_Change
   Command5.Enabled = True
End Sub


Private Sub ControlCalidadComponentesOfActiva_Click()

    Dim RST As ADODB.Recordset
    
    OrdenFabricacion$ = Label144
    If OrdenFabricacion$ = "" Then
        Call COMUNI("Falta Elegir una Orden de Fabricación")
        GoTo 99
    End If
    
    TituloLista$ = "Lotes con Registro de Calidad Incompleto"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Codigo/A16;Descripcion/A32;Lote/A18;CodigoInterno/A0", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Codigo/A16;Descripcion/A32;Lote/A18;CodigoInterno/A0", FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$

    ConsultaSql$ = "SELECT COD_EXTE, DESCRIP, IDENLOTE, COD_INTE FROM RESERVA LEFT JOIN " _
                    & " (SELECT INFOCALIDAD.DOCUMENTOORIGEN, INFOCALIDAD.EMBALAJE, INFOCALIDAD.ENCAPSULADO, INFOCALIDAD.CERTIFICADO, INFOCALIDAD.PATHFOTOCALIDAD, INFOCALIDAD.STATUS, BOLRECEP.NUMEPARTE FROM INFOCALIDAD INNER JOIN BOLRECEP ON INFOCALIDAD.DOCUMENTOORIGEN = BOLRECEP.IDENLOTE) AS CALIDAD " _
                    & " ON RESERVA.IDENLOTE = CALIDAD.DOCUMENTOORIGEN WHERE RESERVA.IDSUBOR='" & TRIM$(OrdenFabricacion$) & "' AND RESERVA.IDENLOTE<>''" _
                    & " AND (CALIDAD.EMBALAJE='' OR CALIDAD.ENCAPSULADO='' OR CALIDAD.CERTIFICADO='' OR (CALIDAD.PATHFOTOCALIDAD='' OR CALIDAD.PATHFOTOCALIDAD LIKE '@CAPTURA%') OR CALIDAD.STATUS='' OR CALIDAD.NUMEPARTE='')"
    Set RST = FLEXCONN.Execute(ConsultaSql$)
    Call Carga_MSF_Recordset(RST, "Codigo/A16;Descripcion/A32;Lote/A18;CodigoInterno/A0", FRMZELECHO.msf2, 1)
    On Error Resume Next
    RST.Close
    Set RST = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 10000
    FRMZELECHO.Show 1
    
    LoteRecepcion$ = TRIM(RESP_ZELE_VECT(3))

    If LoteRecepcion$ <> "" Then
        codigoArticulo$ = TRIM(RESP_ZELE_VECT(1))
        CodigoInterno% = TRIM(RESP_ZELE_VECT(4))
        descripcionArticulo$ = TRIM(RESP_ZELE_VECT(2))
        
        opcionCalidad% = COMALTER%("Opciones de Calidad (" & LoteRecepcion$ & "):\\Informar Registro\Imprimir Etiqueta")
        If opcionCalidad% < 1 Or opcionCalidad% > 2 Then Exit Sub
        
        If opcionCalidad% = 1 Then
            Call FILTERFILE("TAMARCAS", "!TAMARCAS", 4, MKI$(CodigoInterno%))
            
            FILTX$ = LoteRecepcion$ & ";" & codigoArticulo$ & ";" & descripcionArticulo$
            formularioCalidad$ = TRIM$(Control("TRAZABI", "REGICAL"))
            Call STDFORM(formularioCalidad$, FILTX$)
        Else
            CantidadLote# = XVALO(VALOBADA("LOTINGRE", "CANALTA", "IDENLOTE='" & LoteRecepcion$ & "'"))
            
            Call PRINSTIKAPRO(CodigoInterno%, LoteRecepcion$, 1, CantidadLote#)
        End If
    End If
        
99
    
End Sub

Private Sub Command7_Click()
    Call CommSel("TADEPOSI", Command7, Text2)
End Sub

Private Sub Command9_Click()
    Call CommSel("ECODEP", Command9, Text3)
End Sub

Private Sub DevioConsumoOfCerrada_Click()

    Dim RST As ADODB.Recordset
    Dim RsLotesConsumidos As ADODB.Recordset
    Dim RsSaldoPorDeposito As ADODB.Recordset
    
20  TituloLista$ = "Desvío de Consumo en O-Fabricación Cerrada"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "O-Fabricacion/A18;Producto/A48;Orden/I0", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "O-Fabricacion/A18;Producto/A48;Orden/I0", FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$

    ConsultaSql$ = "SELECT IDSUBOR, REFEREN, NUMEORFA FROM ORDEFABR WITH(NOLOCK) WHERE STATUORF = 3 ORDER BY NUMEORFA DESC "
    Set RST = FLEXCONN.Execute(ConsultaSql$)
    Call Carga_MSF_Recordset(RST, "O-Fabricación/A18;Producto/A48;Orden/I0", FRMZELECHO.msf2, 1)
    On Error Resume Next
    RST.Close
    Set RST = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 10000
    FRMZELECHO.Show 1
    
    OrdenDeFabricacion$ = TRIM(RESP_ZELE_VECT(1))
    
    If OrdenDeFabricacion$ <> "" Then
    
30      TituloLista$ = "Lista de Componentes y Lotes Consumidos para la " & OrdenDeFabricacion$
        Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Código/A20;Descripción/A32;Lote/A16;Consumo/F10.1", FRMZELECHO)
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Código/A20;Descripción/A32;Lote/A16;Consumo/F10.1", FRMZELECHO)
        FRMZELECHO.Caption = TituloLista$
    
        ConsultaSql$ = "SELECT COD_EXTE, DESCRIP, IDENLOTE, SUM(CANTSALID) AS CONSUMO FROM MOVISTO WITH(NOLOCK) WHERE DOPRILAR='" & OrdenDeFabricacion$ & "' AND CODIMOVI='CF' GROUP BY COD_EXTE, DESCRIP, IDENLOTE"
        Set RsLotesConsumidos = FLEXCONN.Execute(ConsultaSql$)
        Call Carga_MSF_Recordset(RsLotesConsumidos, "Código/A20;Descripción/A32;Lote/A16;Consumo/F10.1", FRMZELECHO.msf2, 1)
        On Error Resume Next
        RsLotesConsumidos.Close
        Set RsLotesConsumidos = Nothing
        On Error GoTo 0
            
        FRMZELECHO.Width = 10000
        FRMZELECHO.Show 1
        
        CodigoComponente$ = TRIM(RESP_ZELE_VECT(1))
        LoteComponente$ = TRIM(RESP_ZELE_VECT(3))
        
        If LoteComponente$ <> "" Then
        
            ' MUESTRA LOS DEPOSITO DONDE EL LOTE TIENE STOCK
            TituloLista$ = "Saldo por Depósito"
            Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Nro/I5;Depósito/A24;Saldo/F12.1", FRMZELECHO)
            Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Nro/I5;Depósito/A24;Saldo/F12.1", FRMZELECHO)
            FRMZELECHO.Caption = TituloLista$
        
            ConsultaSql$ = "SELECT DEPOMOVI, COD_INTE, (SUM(CANTINGRE) - SUM(CANTSALID)) AS SALDO FROM MOVISTO WITH(NOLOCK) WHERE IDENLOTE='" & LoteComponente$ & "' GROUP BY DEPOMOVI, COD_INTE ORDER BY DEPOMOVI"
            Set RsSaldoPorDeposito = READSET(ConsultaSql$)
            Call Carga_MSF_RecordsetConEco(RsSaldoPorDeposito, "Nro/I5;Depósito(DAT|TADEPOSI|BYTE)/A24;Saldo/F12.1", FRMZELECHO.msf2, 1)
            On Error Resume Next
            RsSaldoPorDeposito.Close
            Set RsSaldoPorDeposito = Nothing
            On Error GoTo 0
                
            FRMZELECHO.Width = 8000
            FRMZELECHO.TXTBUSCAR = ""
            FRMZELECHO.Show 1
            
            Deposito% = XVALO(RESP_ZELE_VECT(1))
            SaldoDeposito# = XVALO(RESP_ZELE_VECT(3))
            If Deposito% = 0 Then GoTo 30
            If SaldoDeposito# < 0.05 Then
                Call MENSERR(24, "Lote sin Saldo para el Depósito Elegido")
                GoTo 30
            End If
            
            SaldoDelLote# = XVALO(InputBox("Saldo del Lote", "Desvío en Consumo de O-Fabricación", TRIM$(FORMATNUM$(SaldoDeposito#, "F10.1"))))
            If SaldoDelLote# < 0.05 Then GoTo 55
            CantidadScrap# = SaldoDeposito# - SaldoDelLote#
            
            If SaldoDeposito# < CantidadScrap# Or CantidadScrap# < 0.05 Then
                Call MENSERR(24, "La Cantidad Ingresada Supera el Saldo del Lote\o la Misma es Incorrecta")
                GoTo 30
            End If
            
            CodigoInterno% = CODINT%(CodigoComponente$)
            
            CMO$ = TRIM$(Control$("", "CODSCRAP"))
            If CMO$ = "" Then CMO$ = "SC"
            NROVAL& = ProNroComprobante(CMO$)
            Comprobante$ = CMO$ + "." + TRIM$(Str$(NROVAL&))
            
            ' GUARDA EL MOVIMIENTO DE SCRAP
            Call MOVISTOK(HOY(HO$), CodigoInterno%, LoteComponente$, CMO$, Comprobante$, Comprobante$, OrdenDeFabricacion$, OrdenDeFabricacion$, "Desvío de Consumo OF Cerrada", 0, 0, 0, 0, Deposito%, CantidadScrap# * (-1))
            
55          Screen.MousePointer = 1
            GoTo 30
        End If
    
        Screen.MousePointer = 1
        GoTo 20
    End If

    Screen.MousePointer = 1
99  Exit Sub

End Sub

Sub Form_Load()
  '
  Screen.MousePointer = 11
  Call CENTRAFORM(Me)
  HOII% = HOY(HOS$)
  Call BLANARCH("!OPERINF")
  Call SETULTREG("!OPERINF", 0)
  Call FRESHECHO("!OPERINF")
  'AGREGUE15/09/06MH
  '
  ExigeCalidadCompleto = False
  If Control("TRAZABI", "EXICOCAL") = "SI" Then ExigeCalidadCompleto = True
  '
  Call CARGA_ENCABEZADO(mfgReserva, "Codigo/A16;Descripcion/A25;Uso/F10.5;Necesidad/F9.1;Trazado/F9.1;Falta/F9.1;PNL-PLC/F9.1;Dep./I3", FRMZELECHO)
  For i = 1 To 8
    mfgReserva.COL = i: mfgReserva.CellFontBold = False
  Next i
  '
  Screen.MousePointer = 1
  '
End Sub

Private Sub InformeTecnico_Click()

    OrdenFabricacion$ = TRIM$(Label144)
    If OrdenFabricacion$ = "" Then
        Call COMUNI("Falta Elegir Orden de Fabricación")
    Else
        TRAZORF14.InformeTecnicaOFEnWord (OrdenFabricacion$)
    End If
    
End Sub

Private Sub Label144_Change()
    '
    Label2 = Label144
    '
End Sub


Private Sub Label15_Change()
   Text1 = Label15
End Sub

Private Sub Label2_Change()
   Screen.MousePointer = 11
   NORFA$ = Label2
   Label4.Caption = ""
   Label6.Caption = ""
   Label8.Caption = ""
   Label10.Caption = ""
   Label11.Caption = ""
   Label13.Caption = ""
   Label15.Caption = ""
   List1.Clear
   mfgReserva.Rows = 1
   '
   Call ABREORFAB
   '
   SQLX$ = " SELECT * FROM Ordefabr WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "' "
   'Set ORFATEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
   Set ORFATEMP = READSET(SQLX$)
   '
   With ORFATEMP
      '
      CIIX% = XVALO(!Cod_Inte)
      Label10.Caption = CODEXT$(CIIX%)
      Label11.Caption = DESCRIT$(CIIX%)
      Label13.Caption = FORMATNUM$(XVALO(!CanProy), "F9.1")
        FESOLENT = CVINT(!FechEnSol)
      Label8.Caption = FECHATEX$(FESOLENT)
        DIATRA& = DIENTRE(FESOLENT, FECALC): If DIATRA < 1 Then DIATRA = 0
      Label6 = TRIM$(Str$(DIATRA&))
      Label4.Caption = XVALO(!Statuorf)
        SALDORF = XVALO(!CanProy) - XVALO(!CanApro): If SALDORF < 0 Then SALDORF = 0
      Label15.Caption = FORMATNUM$(SALDORF, "F9.1")
   End With
   '
   HOII% = HOY(HOS$)
   Label24 = HOS$
   Text2 = Str(DEPENART%(CIIX%))  'PARA MOSTRAR EL DEPOSITO DE ENTRADA Y DE CONSUMO
   Text3 = Str(DEPCONART%(CIIX%)) 'ASIGNADO AL ARTICULO
   '
   Call CARGACOTRA
   '
   Screen.MousePointer = 1
End Sub

Sub CARGACOTRA()
    '
    Static ULTIORFA$
    '
    Dim CODIN%(2048), CATRAZ(2048), USOUNI(2048), CantidadDePlacas(2048)
    '
    CANFA = XVALO(Text1)
    NORFA$ = TRIM$(Label2)
    If NORFA$ = "" Then Exit Sub
    '
    ControlPNLPLC$ = Control("TRAZABI", "EXPNLPLC")
    '
    CONDI$ = "IdSubOr = '" & NORFA$ & "' AND CantRes - CantCons > 0 "
    FIN& = CantRegistros&("RESERVA", CONDI$, "NOMESS")
    
    Call APERBASDAT("RESERVA")
    SQLX$ = " SELECT * FROM RESERVA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    SQLX$ = SQLX$ + " ORDER BY NuOrItem ASC"
    '
'    Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    Set RESERTMP = READSET(SQLX$)
    k& = 0
    With RESERTMP
      If .EOF Then
         If Label144 <> ULTIORFA$ Then     ' PARA QUE NO REPITA EL MENSAJE
            Call MENSERR(24, "No Existen Reservas Vigentes\para la O/F Elegida. ")
            ULTIORFA$ = Label144
         End If
         GoTo 99
      End If
      JJ& = 0
      Do While (.EOF = False)
        k& = k& + 1
        Call TRACE(20, k&, FIN&)
        '
        CIXI% = XVALO(!Cod_Inte)
        IDLO$ = TRIM$(SAFETEXT$(!IDENLOTE))
        CANRES = XVALO(!CANTRES) - XVALO(!CantCons)
        CANTRA = 0
            CantidadPlacasDesde& = 0
            CantidadPlacasHasta& = 0
        If IDLO$ <> "" Then
            CANTRA = CANRES
            CantidadPlacasDesde& = XVALO(!PNL_DES)
            CantidadPlacasHasta& = XVALO(!PLC_HAS)
        End If
        '
        For JJ& = 1 To CANMO&
           If CODIN%(JJ&) = CIXI% Then
              CATRAZ(JJ&) = CATRAZ(JJ&) + CANTRA
              CantidadDePlacas(JJ&) = CantidadDePlacas(JJ&) + (CantidadPlacasHasta& - CantidadPlacasDesde& + 1)
              GoTo 19
           End If
        Next JJ&
        CANMO& = CANMO& + 1
        CODIN%(CANMO&) = CIXI%
        CATRAZ(CANMO&) = CANTRA
        USOUNI(CANMO&) = XVALO(!UsoUnit)
        CantidadDePlacas(CANMO&) = CantidadPlacasHasta& - CantidadPlacasDesde& + 1
        '
19    .MoveNext
      Loop
    End With
    '
    List1.Clear: mfgReserva.Rows = 1
    For JJ& = 1 To CANMO&
       TTXX$ = Space$(128)
       CIXI% = CODIN%(JJ&)
            Codigo$ = CODEXT$(CIXI%)
       Call REPLA(TTXX$, Codigo$, 1, 16)
            Descripcion$ = DESCRIT$(CIXI%)
       Call REPLA(TTXX$, Descripcion$, 17, 25)
            Uso$ = FORMATNUM(USOUNI(JJ&), "F10.4")
       Call REPLA(TTXX$, Uso$, 42, 10)
            CANCO = USOUNI(JJ&) * CANFA
            CantidadNecesaria$ = FORMATNUM(CANCO, "F10.1")
       Call REPLA(TTXX$, CantidadNecesaria$, 52, 10)
            CANTRA = CATRAZ(JJ)
            CantidadTrazada$ = FORMATNUM(CANTRA, "F10.1")
       Call REPLA(TTXX$, CantidadTrazada$, 62, 10)
            FALTA = CANCO - CATRAZ(JJ): If FALTA < 0 Then FALTA = 0
            CantidadFaltante$ = FORMATNUM(FALTA, "F10.1")
       Call REPLA(TTXX$, CantidadFaltante$, 72, 10)
            DPXX% = XVALO(Text3)
            If XVALO(Text3) = 99 Then DPXX% = DEPOPRE%(CIXI%)
            Deposito$ = FORMATNUM$(DPXX%, "F5.0")
       Call REPLA(TTXX$, Deposito$, 82, 5)
       List1.AddItem TTXX$
       '
       ' COMPLETA GRILLA DE RESERVAS
       mfgReserva.Rows = mfgReserva.Rows + 1
       FilaGrilla% = mfgReserva.Rows - 1
       mfgReserva.TextMatrix(FilaGrilla%, 1) = Codigo$
       mfgReserva.TextMatrix(FilaGrilla%, 2) = Descripcion$
       mfgReserva.TextMatrix(FilaGrilla%, 3) = Uso$
       mfgReserva.TextMatrix(FilaGrilla%, 4) = CantidadNecesaria$
       mfgReserva.TextMatrix(FilaGrilla%, 5) = CantidadTrazada$
       mfgReserva.TextMatrix(FilaGrilla%, 6) = CantidadFaltante$
            TotalDePlacas$ = FORMATNUM(CantidadDePlacas(JJ&), "I10")
       mfgReserva.TextMatrix(FilaGrilla%, 7) = TotalDePlacas$
       mfgReserva.TextMatrix(FilaGrilla%, 8) = Deposito$
       '
    Next JJ&
99 End Sub
'
Private Sub Label24_Change()
   FEX1 = FECHANUM(Label8)
   FEX2 = FECHANUM(Label24)
   ATRASO = DIENTRE(FEX1, FEX2)
   Label6 = FORMATNUM$(ATRASO, "I3")
End Sub

Private Sub Label24_DblClick()
   Call Command4_Click
End Sub

Private Sub Label8_Change()
   Call Label24_Change
End Sub

Private Sub DesvioPNLPLCvsTrazado_Click()

    Dim RST As ADODB.Recordset
    Dim RsSaldoPorDeposito As ADODB.Recordset
    
    OrdenFabricacion$ = Label144
    If OrdenFabricacion$ = "" Then
        Call COMUNI("Falta Elegir una Orden de Fabricación")
        GoTo 99
    End If
    
    ' MUESTRA LA DIFERENCIA ENTRE LO TRAZADO Y LAS CANTIDAD INGRESADAS EN PNL_DES Y PLC_HAS. DICHA DIFERENCIA ES SCRAP EN LINEA DE PRODUCCION
20  TituloLista$ = "Desvío PNL-PLC vs Cantidad Trazada"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Componente/A16;Descripcion/A32;Lote/A16;Reserva/F8.1;Desde/A8;Hasta/A8;Scrap/F8.1", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Componente/A16;Descripcion/A32;Lote/A16;Reserva/F8.1;Desde/A8;Hasta/A8;Scrap/F8.1", FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$

    ConsultaSql$ = "SELECT COD_EXTE, DESCRIP, IDENLOTE, CANTRES, PNL_DES, PLC_HAS, (CAST(PLC_HAS AS INT) - CANTRES) AS SCRAP FROM RESERVA WITH(NOLOCK) WHERE IDSUBOR='" & OrdenFabricacion$ & "' AND IDENLOTE <> '' AND PNL_DES <> '' AND PLC_HAS <> '' ORDER BY COD_EXTE "
    Set RST = FLEXCONN.Execute(ConsultaSql$)
    Call Carga_MSF_Recordset(RST, "Componente/A16;Descripcion/A32;Lote/A16;Reserva/F8.1;Desde/A8;Hasta/A8;Scrap/F8.1", FRMZELECHO.msf2, 1)
    On Error Resume Next
    RST.Close
    Set RST = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 12000
    FRMZELECHO.Show 1
    
    CodigoComponente$ = TRIM(RESP_ZELE_VECT(1))
    LoteComponente$ = TRIM(RESP_ZELE_VECT(3))
    CantidadScrap# = Abs(XVALO(RESP_ZELE_VECT(7)))
    
    If CodigoComponente$ <> "" Then
    
        ' MUESTRA LOS DEPOSITO DONDE EL LOTE TIENE STOCK
        TituloLista$ = "Saldo por Depósito"
        Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Nro/I5;Depósito/A24;Saldo/F12.1", FRMZELECHO)
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Nro/I5;Depósito/A24;Saldo/F12.1", FRMZELECHO)
        FRMZELECHO.Caption = TituloLista$
    
        ConsultaSql$ = "SELECT DEPOMOVI, COD_INTE, (SUM(CANTINGRE) - SUM(CANTSALID)) AS SALDO FROM MOVISTO WITH(NOLOCK) WHERE IDENLOTE='" & LoteComponente$ & "' GROUP BY DEPOMOVI, COD_INTE ORDER BY DEPOMOVI"
        Set RsSaldoPorDeposito = READSET(ConsultaSql$)
        Call Carga_MSF_RecordsetConEco(RsSaldoPorDeposito, "Nro/I5;Depósito(DAT|TADEPOSI|BYTE)/A24;Saldo/F12.1", FRMZELECHO.msf2, 1)
        On Error Resume Next
        RsSaldoPorDeposito.Close
        Set RsSaldoPorDeposito = Nothing
        On Error GoTo 0
            
        FRMZELECHO.Width = 8000
        FRMZELECHO.TXTBUSCAR = ""
        FRMZELECHO.Show 1
        
        Deposito% = XVALO(RESP_ZELE_VECT(1))
        SaldoDeposito# = XVALO(RESP_ZELE_VECT(3))
        If Deposito% = 0 Then GoTo 20
        If SaldoDeposito# < CantidadScrap# Or CantidadScrap# < 0.05 Then
            Call MENSERR(24, "La Cantidad Ingresada Supera el Saldo del Lote\o la Misma es Incorrecta")
            GoTo 20
        End If
        
        CodigoInterno% = CODINT%(CodigoComponente$)
        '
        CMO$ = TRIM$(Control$("", "CODSCRAP"))
        If CMO$ = "" Then CMO$ = "SC"
        NROVAL& = ProNroComprobante(CMO$)
        Comprobante$ = CMO$ + "." + TRIM$(Str$(NROVAL&))
        
        ' GUARDA EL MOVIMIENTO DE SCRAP
        Call MOVISTOK(HOY(HO$), CodigoInterno%, LoteComponente$, CMO$, Comprobante$, Comprobante$, OrdenDeFabricacion$, OrdenDeFabricacion$, "Desvío OF PNL-PLC", 0, 0, 0, 0, Deposito%, CantidadScrap# * (-1))
            
        ' ACTUALIZA LA CANTIDAD RESERVADA, DISMINUYENDO LA RESERVA EN LA CANTIDAD DE SCRAP ( RESERVA = RESERVA - CantidadScrap )
            ConsultaSql$ = "UPDATE RESERVA SET CANTRES=CANTRES-" & TRIM$(FORMATNUM$(CantidadScrap#, "F10.1")) & " WHERE IDSUBOR='" & OrdenFabricacion$ & "' AND COD_INTE=" & CStr(CodigoInterno%) & " AND IDENLOTE='" & LoteComponente$ & "'"
        FLEXCONN.Execute ConsultaSql$
        
        Screen.MousePointer = 1
        GoTo 20
    End If

99  Exit Sub
End Sub

Private Sub mnuAsociarReporte_Click()
    Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
End Sub

Private Sub mnuReportesVarios_Click()
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
End Sub

Private Sub Text1_Change()
   Call CARGACOTRA
End Sub

Private Sub Text2_Change()
  DEPO% = XVALO(Text2)
  Label22.Caption = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPO%)))
End Sub

Private Sub Text2_DblClick()
  Call Command7_Click
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text3.SetFocus
   End If
End Sub

Private Sub Text3_Change()
  DEPO% = XVALO(Text3)
  Label20.Caption = TRIM$(ECOARCH$("ECODEP", Chr$(DEPO%)))
  Call CARGACOTRA
End Sub

Private Sub Text3_DblClick()
  Call Command9_Click
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 13 Then
       List1.SetFocus
    End If
End Sub


