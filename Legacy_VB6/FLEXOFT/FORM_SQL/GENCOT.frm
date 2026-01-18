VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form GENCOT 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   0  'None
   ClientHeight    =   8775
   ClientLeft      =   1845
   ClientTop       =   3525
   ClientWidth     =   13755
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleMode       =   0  'User
   ScaleWidth      =   12976.42
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command7 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   13395
      TabIndex        =   64
      ToolTipText     =   "Selección  de Remitos a inclluir en C.O.T."
      Top             =   120
      Width           =   175
   End
   Begin VB.TextBox Text17 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Text            =   "Remitos del mes"
      Top             =   120
      Width           =   1770
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11640
      TabIndex        =   62
      TabStop         =   0   'False
      ToolTipText     =   "Doble Click Para Borrar Mes Activo"
      Top             =   120
      Width           =   1725
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6720
      TabIndex        =   61
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Text            =   "Soilicitud de C.O.T.  N°:"
      Top             =   120
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6240
      Top             =   4200
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   8640
      TabIndex        =   48
      Top             =   5760
      Width           =   4875
      Begin VB.Frame Frame9 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Fecha Salida"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         TabIndex        =   66
         Top             =   720
         Width           =   2040
         Begin VB.CommandButton Command9 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1755
            TabIndex        =   68
            Top             =   255
            Width           =   175
         End
         Begin VB.TextBox Text20 
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
            Left            =   480
            TabIndex        =   67
            Top             =   240
            Width           =   1230
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00E2D3C0&
         Height          =   720
         Left            =   120
         TabIndex        =   54
         Top             =   1320
         Width           =   4650
         Begin VB.ComboBox COMBO1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   285
            Width           =   3255
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00D8BD05&
            BackStyle       =   0  'Transparent
            Caption         =   "Kilometros"
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
            Index           =   0
            Left            =   0
            TabIndex        =   56
            Top             =   360
            Width           =   1065
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "GENERAR C.O.T."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   2280
         Width           =   4665
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Rto. Seleccionado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   51
         Top             =   120
         Width           =   1935
         Begin VB.Label lblRemSel 
            Alignment       =   1  'Right Justify
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
            Height          =   270
            Left            =   240
            TabIndex        =   52
            Top             =   240
            Width           =   1485
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Importe Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         TabIndex        =   49
         Top             =   120
         Width           =   2040
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
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
            Height          =   270
            Left            =   450
            TabIndex        =   50
            Top             =   240
            Width           =   1455
         End
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   8640
      TabIndex        =   43
      Top             =   4800
      Width           =   4875
      Begin VB.Frame Frame17 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Cant. de Items"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   1695
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
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
            Height          =   270
            Left            =   90
            TabIndex        =   47
            Top             =   240
            Width           =   1485
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Cant. de Remitos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   44
         Top             =   240
         Width           =   1650
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
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
            Height          =   270
            Left            =   90
            TabIndex        =   45
            Top             =   240
            Width           =   1450
         End
      End
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Text            =   "Remito de Entrega"
      Top             =   840
      Width           =   3315
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar no visible"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   720
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Datos de Entrega"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   120
      TabIndex        =   29
      Top             =   4800
      Width           =   8355
      Begin VB.CommandButton Command10 
         Caption         =   "Dom. - Entregas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6720
         TabIndex        =   69
         ToolTipText     =   "Lista de Domicilios de Entrega de Cliente Activo"
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Act.Ficha Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6720
         TabIndex        =   65
         ToolTipText     =   "Actualiza los datos de Entrega en la Ficha del Cliente"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox LENTREGA 
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
         Index           =   1
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   5415
      End
      Begin VB.TextBox LENTREGA 
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
         Index           =   0
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   5415
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "Localidad:"
         Top             =   600
         Width           =   1050
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "Domicilio:"
         Top             =   240
         Width           =   1050
      End
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Text            =   "Tipo de Remito (Entrega o Devolución)"
      Top             =   840
      Width           =   3360
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   "E"
      ToolTipText     =   "Doble Click Para Cambiar a Devolución o Entrega"
      Top             =   840
      Width           =   1395
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   480
      Width           =   2925
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "Fecha:"
      Top             =   720
      Width           =   810
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   12450
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   720
      Width           =   1113
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8880
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1920
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   120
      Width           =   1125
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "Numero de Remito:"
      Top             =   480
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "Numero de Cliente:"
      Top             =   120
      Width           =   1770
   End
   Begin VB.CommandButton Command5 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10635
      TabIndex        =   17
      ToolTipText     =   "Selección  de Remitos a inclluir en C.O.T."
      Top             =   480
      Visible         =   0   'False
      Width           =   175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3020
      Left            =   120
      TabIndex        =   13
      Top             =   5760
      Width           =   8355
      Begin VB.TextBox TXTPCIA 
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
         Left            =   1800
         TabIndex        =   42
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Index           =   3
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "Pcia.:"
         Top             =   720
         Width           =   810
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   720
         Width           =   440
      End
      Begin VB.CommandButton Command6 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   40
         ToolTipText     =   "Selección de Provincia"
         Top             =   720
         Width           =   175
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "C.Postal:"
         Top             =   330
         Width           =   810
      End
      Begin VB.TextBox TXTCPOSCLI 
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
         Left            =   1080
         TabIndex        =   2
         Top             =   330
         Width           =   2895
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Transportista"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         Left            =   4440
         TabIndex        =   15
         Top             =   210
         Width           =   3690
         Begin VB.CommandButton Command2 
            Caption         =   "Actualizar Datos de Entrega"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   2280
            Width           =   3480
         End
         Begin VB.CommandButton Command18 
            Caption         =   "Editar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Modificar Datos del Transporte Activo"
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Nuevo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Agrega Nuevo Transporte a la Base de Datos"
            Top             =   240
            Width           =   1050
         End
         Begin VB.TextBox Text45 
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
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   750
         End
         Begin VB.CommandButton Command19 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   920
            TabIndex        =   33
            Top             =   260
            Width           =   175
         End
         Begin VB.TextBox LENTREGA 
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
            Index           =   2
            Left            =   120
            TabIndex        =   8
            Top             =   675
            Width           =   3495
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Domicilio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   105
            TabIndex        =   16
            Top             =   1080
            Width           =   3480
            Begin VB.CommandButton Command3 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   3120
               TabIndex        =   36
               ToolTipText     =   "Aplicar Domicilio de Transporte Como Domicilio de Entrega"
               Top             =   240
               Width           =   300
            End
            Begin VB.TextBox DOMITRAN 
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
               Index           =   4
               Left            =   120
               TabIndex        =   32
               Top             =   600
               Width           =   2980
            End
            Begin VB.TextBox DOMITRAN 
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
               Index           =   3
               Left            =   120
               TabIndex        =   9
               Top             =   240
               Width           =   2980
            End
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Conductor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1700
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   3885
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
            Height          =   285
            Left            =   960
            TabIndex        =   5
            Top             =   720
            Width           =   2800
         End
         Begin VB.TextBox Text12 
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
            Left            =   960
            TabIndex        =   4
            Top             =   360
            Width           =   2800
         End
         Begin VB.TextBox Text19 
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
            Left            =   960
            TabIndex        =   6
            Top             =   1065
            Width           =   2800
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Patente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   120
            TabIndex        =   59
            Top             =   795
            Width           =   675
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   120
            TabIndex        =   58
            Top             =   1140
            Width           =   675
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   360
            TabIndex        =   57
            Top             =   435
            Width           =   435
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   720
      OleObjectBlob   =   "GENCOT.frx":0000
      Top             =   3000
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   3285
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "Botón Derecho Opciones "
      Top             =   1320
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   5794
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   855
   End
   Begin VB.Menu mnuarchivo1 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Visible         =   0   'False
      Begin VB.Menu mnucopiar 
         Caption         =   " Copiar Celda Seleccionada"
         Visible         =   0   'False
      End
      Begin VB.Menu QuitarRemi 
         Caption         =   "Quitar Remito Seleccionado"
         Visible         =   0   'False
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuver 
      Caption         =   "Ver"
      Begin VB.Menu mnuversolcot 
         Caption         =   "Ver Remitos de Solicitud de C.O.T."
         Begin VB.Menu mnusolaprobcot 
            Caption         =   "Solicitudes de C.O.T. Aprobadas"
         End
         Begin VB.Menu mnuReimprimirCot 
            Caption         =   "Re-Imprimir C.O.T."
         End
      End
   End
   Begin VB.Menu mnuremitdev 
      Caption         =   "Remitos de Devolución"
      Begin VB.Menu mnuGenRemDev 
         Caption         =   "Generar Remito de Devolución"
      End
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuración"
      Begin VB.Menu Config 
         Caption         =   "Configuraciones C.O.T."
      End
   End
End
Attribute VB_Name = "GENCOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GLOBALFRM_strContenidoCelda$
Dim GLOBALFRM_FILA&
Dim COT$
Dim IDCOT$
Dim ID_COT&
Dim NAMEARCHIVO$


Sub CARGAR_DATOS_ENTREGA(NCLIE, NroPed&, REMI$)
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    

    If NroPed& > 0 Then
       SQLX$ = "SELECT DOMIENT,LOCAENT,TranspEnt,NROOCOM,FLETE_TEX, NTranspo FROM PEDENCA WITH(NOLOCK)"
       SQLX$ = SQLX$ + " WHERE NROPED = " & NroPed&
       Set rs = READSET(SQLX$)
       With rs
         If Not rs.EOF Then
           If TRIM$(SAFETEXT$(!DOMIENT)) <> "" Then
                LENTREGA(0) = SAFETEXT(!DOMIENT)
                LENTREGA(1) = SAFETEXT(!LOCAENT)
                NTRAN% = XVALO(!NTranspo)
                If NTRAN% > 0 Then ' SI VIENE CON CODIGO DE TRANSPORTE COMPLETA AUTOMATICAMENTE
                  Text45 = NTRAN%
                Else
                  LENTREGA(2) = SAFETEXT(!TranspEnt)
                End If
                
           End If
         End If
       End With
20     rs.Close
       Set rs = Nothing
       TXTCPOSCLI.TEXT = CPOSCLI$(NCLIE)
       
       If InStr(EPAC$, "ENVAPLAST") > 0 Then 'BUSCA EL DATO DE COD.POSTAL  SI ESTA ENTREPARENTESIS EN LA CAJA LOCALIDAD
        If TRIM$(LENTREGA(1)) <> "" Then
           Call TEXTOLOTES$(TRIM$(LENTREGA(1)), "(", ")")
           On Error Resume Next
           TXTCPOSCLI.TEXT = CADENATEX$(1)
        End If
       End If
       
       If LENTREGA(0) = "" Then
        Text45 = TRIM$(Str$(NUTRANCLI%(NCLIE)))
        If TRIM$(Text45) = "0" Then Text45 = ""
          LENTREGA(0) = TRIM$(DOMIENT$(NCLIE))
          LENTREGA(1) = TRIM$(LOCAENT$(NCLIE))
          LENTREGA(2) = NUTRANCLI%(NCLIE)
       End If
    End If
    
    Text8.TEXT = PROVCLI$(NCLIE)
    Call AgregarDatosdeEntregaAGrilla(REMI$)
End Sub
Sub AgregarDatosdeEntregaAGrilla(REMI$)

    For i& = 1 To GRID.Rows - 1
       If REMI$ = GRID.TextMatrix(i&, 8) Then
          
          GRID.TextMatrix(i&, 10) = LENTREGA(0)
          GRID.TextMatrix(i&, 11) = LENTREGA(1)
'          GRID.TextMatrix(i&, 12) = LENTREGA(2)
          GRID.TextMatrix(i&, 12) = TXTCPOSCLI.TEXT
          GRID.TextMatrix(i&, 13) = "" 'Text12 'CUIT TRANSPORTE
          GRID.TextMatrix(i&, 14) = Text8
          GRID.TextMatrix(i&, 15) = DOMITRAN(3) 'DOMICILIO TRANSPORTE
          GRID.TextMatrix(i&, 16) = DOMITRAN(4)  'LOCALIDAD TRANSPORTE
          GRID.TextMatrix(i&, 17) = "" ' Text18 'PATENTE
          GRID.TextMatrix(i&, 18) = Text2 'CLIENTE
          'GRID.TextMatrix(i&, 19) = Text13 'NUMERO DNI CLIENTE SI ES POSICI
       End If
    Next i&
End Sub
Sub AgregarDatosDeGrillaaText(REMI$)
    For i& = 1 To GRID.Rows - 1
       If REMI$ = GRID.TextMatrix(i&, 8) Then
          LENTREGA(0) = GRID.TextMatrix(i&, 10)
          LENTREGA(1) = GRID.TextMatrix(i&, 11)
'          LENTREGA(2) = GRID.TextMatrix(i&, 12)
          TXTCPOSCLI.TEXT = GRID.TextMatrix(i&, 12)
          'Text12 = GRID.TextMatrix(i&, 13)  'CUIT TRANSPORTE
          Text8 = GRID.TextMatrix(i&, 14)
          DOMITRAN(3) = GRID.TextMatrix(i&, 15)    'DOMICILIO TRANSPORTE
          DOMITRAN(4) = GRID.TextMatrix(i&, 16)     'LOCALIDAD TRANSPORTE
          'Text18 = GRID.TextMatrix(i&, 17) 'PATENTE
          Text2 = GRID.TextMatrix(i&, 18) 'CLIENTE
          'Text13 = GRID.TextMatrix(i&, 19)
          Text6 = rasocli$(XVALO(Text2))
          'Text13 = GRID.TextMatrix(i&, 19)
          Exit For
       End If
    Next i&
    Label18 = GRID.TextMatrix(GRID.Rows - 1, 2)

End Sub
Sub COMPLETAR_CASILLEROS(IDCOT&)
        Condi1$ = "ID_ENCA =" & IDCOT&
        PATENTE$ = TRIM$(VALOBADA("DETACOT", "PATENTE", Condi1$, "NOMESS"))
        Cuitran$ = VALOBADA("DETACOT", "CUITTRANSP", Condi1$, "NOMESS")
        CONDI2$ = "ENCACOT_ID =" & IDCOT&
        FE% = CVINT(VALOBADA("ENCACOT", "FESALIDA", CONDI2$, "NOMESS"))
        Text12.TEXT = Cuitran$
        Text18.TEXT = PATENTE$
        Text20.TEXT = FECHATEX$(FE%)
End Sub

Sub Generar_Archivo_Cot(IDENCA&, ret_DATOSCOT$, ret_NAMEARCH$)
    Dim oDE As New DatoEmpresa 'instancio cla
    CUIT1$ = TRIM$(oDE.Cuit)
    'REEEMPLAZO GUIONES POR VACIO Y POR LAS DUDAS PUNTO O GUION BAJO
    CUIT1$ = REPLACAR$(CUIT1$, "-", ""): CUIT1$ = REPLACAR$(CUIT1$, ".", ""): CUIT1$ = REPLACAR$(CUIT1$, "_", "")
'   '
    NREMI1$ = "": NREMI_ANT$ = ""
    TX1$ = ""
'    ANO1$ = Year(Date)
'    MES1$ = Month(Date)
'    DI1$ = Left$(FECHATEX(HOY(HO$)), 2)
'    FECHA1$ = ANO1$ & MES1$ & DI1$
     FECHA1$ = FECHAFORMATOINVERSO$(HOY(HO$))
    '
    'FECHA ENTREGA
'    ANOENT$ = Year(Text20)
'    MESENT$ = Month(Text20)
'    DIENT$ = Left$(Text20, 2)
    FECHAENT$ = FECHAFORMATOINVERSO$(FECHANUM(Text20))
    
''  NAMEARCH$ = "TB_30714269026_003002_20171714_000005" & FECHA1$ & NEXTIDCOT
    NAMEARCH$ = "TB_" & TRIM$(CUIT1$) & "_000001_" & FECHA1$ & "_" & FORMATOCON0$(IDENCA&, 6)
    CARPETOK$ = LUCOM$ & NAMEARCH$ & ".TXT"
    NARCHI2% = FILEOPEN%(CARPETOK$, 2, 128)  ' ABRE EL ARCHIVO DE ESCRITURA
    TX1$ = "01" & "|"
    TX1$ = TX1$ & CUIT1$ & vbCrLf
    Debug.Print TX1$
'    TX1$ = TX1$ & "02" & "|"
'    TX1$ = TX1$ & FECHA1$ & "|" '4
    
    For i& = 1 To GRID.Rows - 1
      NREMI1$ = TRIM$(GRID.TextMatrix(i&, 8))
      If NREMI1$ <> NREMI_ANT$ Then
        TX1$ = TX1$ & "02" & "|"
        FEREM% = FECHANUM(TRIM$(GRID.TextMatrix(i&, 9)))
        FECHREMI$ = FECHAFORMATOINVERSO$(FEREM%)
        TX1$ = TX1$ & FECHREMI$ & "|" '4
        ImporteRemito$ = CalcutotalRemito$(NREMI1$)
        NC = XVALO(GRID.TextMatrix(i&, 18))
        PREFIJO$ = TRIM$(Left(GRID.TextMatrix(i&, 8), 4))
        While Len(PREFIJO$) < 5: PREFIJO$ = "0" & PREFIJO$: Wend
        NREMI$ = TRIM$(Mid$(GRID.TextMatrix(i&, 8), 6))
        ' TX1$ = TX1$ & "91 R" & PREFIJO$ & NREMI$ & "|" '5
        TX1$ = TX1$ & "091" & PREFIJO$ & NREMI$ & "|" '5
        TX1$ = TX1$ & FECHAENT$ & "|" '6 F.SALIDA MERCADERIA
        HORAX$ = Format(Time, "HH:MM")
        HORAX$ = REPLACAR(HORAX$, ":", "")
        TX1$ = TX1$ & HORAX$ & "|" '7
        LETRA$ = TRIM$(Text7.TEXT) 'E POR QUE ES ENVIO
        TX1$ = TX1$ & LETRA$ & "|" '8
        If LETRA$ = "E" Then
          VentaoDevolucion$ = "0" 'venta
          ImporteRemito$ = TRIM$(Label3)
          CUITX$ = TRIM$(CuitCli(NC))
          RazSocSale$ = TRIM$(oDE.razonSocial)
        Else
          VentaoDevolucion$ = "1" 'devolucion
          ImporteRemito$ = "0"
          CUITX$ = TRIM$(oDE.Cuit) ' CUIT DE EMPRESA
          RazSocSale$ = TRIM$(rasocli$(NC))

        End If
        
        If PIVACLI(NC) < 1 Then
          DESTINATARIO$ = "1"
          TIPODOCUMENTA$ = "DNI"
          NDOCUM$ = TRIM$(GRID.TextMatrix(i&, 19))
        Else
          DESTINATARIO$ = "0"
          TIPODOCUMENTA$ = ""
          NombreEmpresa$ = ""
          NDOCUM$ = ""
        End If
        TX1$ = TX1$ & DESTINATARIO$ & "|" '9
        
        'TIPODOCUMENTA$ = "DNI" SOLO SI ES CONSUMIDOR FINAL
        TIPODOCUMENTA$ = ""
        TX1$ = TX1$ & TIPODOCUMENTA$ & "|" '10
        '
        TX1$ = TX1$ & NDOCUM$ & "|" '11
        '
        CUITX$ = REPLACAR$(CUITX$, "-", ""): CUITX$ = REPLACAR$(CUITX$, ".", ""): CUITX$ = REPLACAR$(CUITX$, "_", "")

        TX1$ = TX1$ & CUITX$ & "|" '12
        NombreEmpresa$ = TRIM$(rasocli$(NC))

        TX1$ = TX1$ & Left(TRIM$(NombreEmpresa$), 50) & "|" '13
        TX1$ = TX1$ & DESTINATARIO$ & "|" '14
          DomicilioEntrega$ = Left(TRIM$(GRID.TextMatrix(i&, 10)), 40)
        TX1$ = TX1$ & DomicilioEntrega$ & "|" '15
        TX1$ = TX1$ & "|S/N"  '16
        TX1$ = TX1$ & "|"  '17
        TX1$ = TX1$ & "|"  '18
        TX1$ = TX1$ & "|"  '19
        TX1$ = TX1$ & "|"  '20
        CODIPOSTAL$ = Left$(GRID.TextMatrix(i&, 12), 8)
        TX1$ = TX1$ & CODIPOSTAL$ & "|"  '21
         LocalidadEntrega$ = TRIM$(GRID.TextMatrix(i&, 11))
        TX1$ = TX1$ & LocalidadEntrega$ & "|"  '22
         PCIA$ = TRIM$(LetraProvArbaSegunLetraFlexoft$(TRIM$(GRID.TextMatrix(i&, 14))))
        TX1$ = TX1$ & PCIA$ & "|"  '23
        TX1$ = TX1$ & "|"  '24
        TX1$ = TX1$ & "NO|"  '25
        '
        TX1$ = TX1$ & CUIT1$ & "|"  '26
        TX1$ = TX1$ & Left(RazSocSale$, 50) & "|" '27
        TX1$ = TX1$ & "0|"   '28
        If LETRA$ = "E" Then
          TX1$ = TX1$ & Left(TRIM$(oDE.CalleNumero), 40) & "|"  '29
        Else
          TX1$ = TX1$ & Left(DomicilioEntrega$, 40) & "|"  '29 DEVOLUCION
        End If
        TX1$ = TX1$ & "|S/N"  '30
        TX1$ = TX1$ & "|"  '31
        TX1$ = TX1$ & "|"  '32
        TX1$ = TX1$ & "|"  '33
        TX1$ = TX1$ & "|"  '34
        '
        If LETRA$ = "E" Then
           CodPostSale$ = Left(TRIM$(oDE.CodPostal), 8)
           LocalSale$ = Left(TRIM$(oDE.Localidad), 50)
           LetraSale$ = Left(TRIM$(LetraProvArbaSegunLetraFlexoft$(oDE.LetProvincia)), 1)
        Else
           CodPostSale$ = Left(TRIM$(CODIPOSTAL$), 8)
           LocalSale$ = Left(TRIM$(LocalidadEntrega$), 50)
           LetraSale$ = Left(TRIM$(PCIA$), 1)
        End If
        '
        TX1$ = TX1$ & CodPostSale$ & "|"  '35
        TX1$ = TX1$ & LocalSale$ & "|"  '36
        TX1$ = TX1$ & LetraSale$ & "|"  '37
        '
        CuitTransportista$ = TRIM$(Text12.TEXT)
        CuitTransportista$ = REPLACAR$(CuitTransportista$, "-", ""): CuitTransportista$ = REPLACAR$(CuitTransportista$, ".", ""): CuitTransportista$ = REPLACAR$(CuitTransportista$, "_", "")
        TX1$ = TX1$ & CuitTransportista$ & "|"  '38
        
        TipoCaminoRecorrido$ = "U"
        TX1$ = TX1$ & TipoCaminoRecorrido$ & "|"  '39
        '
        TX1$ = TX1$ & "|"  '40
        TX1$ = TX1$ & "|"  '41
        TX1$ = TX1$ & "|"  '42
        '
        PATENTE$ = Text18.TEXT
        PATENTE$ = REPLACAR$(PATENTE$, " ", "")
        TX1$ = TX1$ & PATENTE$ & "|"  '43
        '
        TX1$ = TX1$ & "|"  '44
        TX1$ = TX1$ & VentaoDevolucion$ & "|"  '45
        ImporteRemito$ = REPLACAR$(ImporteRemito$, ",", ""): ImporteRemito$ = REPLACAR$(ImporteRemito$, ".", "")
        TX1$ = TX1$ & ImporteRemito$ & vbCrLf  '46
        NREMI_ANT$ = NREMI1$
      End If
      TX1$ = TX1$ & "03|" '47
      
      CI1% = XVALO(GRID.TextMatrix(i&, 1))
      CoArba$ = CodigoArba$(CI1%)
      If TRIM$(CoArba$) = "" Then
         CoArba$ = CONTROL("REMITO", "CODIARBA")
      End If
      TX1$ = TX1$ & CoArba$ & "|" '48
      
      'TX1$ = TX1$ & "23891" & "|" '48
      X1X = UnidMedidaArbaSegunUnidFlexoft$(Unimed$(CI1%), RETNUM%)
      TX1$ = TX1$ & SAFETEXT$(RETNUM%) & "|" '49
      CANTI1$ = TRIM$(FORMATNUM$(XVALO(GRID.TextMatrix(i&, 4)), "F16.2"))
      CANTI1$ = REPLACAR$(CANTI1$, ",", ""): CANTI1$ = REPLACAR$(CANTI1$, ".", "")
      
      TX1$ = TX1$ & SAFETEXT$(CANTI1$) & "|" '50
      CODI$ = CODEXT$(CI1%)
      TX1$ = TX1$ & CODI$ & "|" '51
      DESCRICION$ = TRIM$(Left$(GRID.TextMatrix(i&, 3), 40))
      TX1$ = TX1$ & DESCRICION$ & "|" '52
      TX1$ = TX1$ & TRIM$(ECOARCH$("UMEDIDA", Unimed$(CI1%))) & "|" '53
      TX1$ = TX1$ & SAFETEXT$(CANTI1$) & vbCrLf '54 ' SE REPITE LA CANTI1$DAD.
      Debug.Print TX1$
    Next i&
    '
    TX1$ = TX1$ & "04|"
    TX1$ = TX1$ & SAFETEXT$(XVALO(Label2.Caption))
    Print #NARCHI2%, TX1$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
    Close #NARCHI2%
    
    ModalidadCOT$ = CONTROL("REMELEC", "MODOCOT")
    If ModalidadCOT$ = "ARBA" Then
        ret_DATOSCOT$ = TRIM$(CUIT1$) & " " & CONTROL$("REMITO", "PASSCOT") & " " & Mid$(NAMEARCH$, 4)
    ElseIf ModalidadCOT$ = "AGIP" Then
        Token$ = TRIM$(CONTROL("REMELEC", "TOKEAGIP"))
        Sign$ = TRIM$(CONTROL("REMELEC", "SIGNAGIP"))
        ret_DATOSCOT$ = Token$ & " " & Sign$ & " " & Mid$(NAMEARCH$, 4)
    End If
    
    ret_NAMEARCH$ = Mid$(NAMEARCH$, 4)
    Set oDE = Nothing '

End Sub

Function GRABAR_DATOS_SOLICITUD_COT()
'1.-VALIDAR  SI LOS REMITOS YA FUERON GRABADOS
'2.-VER EL DETALLE DE LOS GUARDADOS
'3.-VALIDAR SI ESTA GRABADO SI OBTUVO COT
    'ESTA FUNCION LLAMA A UN SP QUE LE DEVUELVE EL ID DEL REGISTRO GENERADO
    FE% = FECHANUM(Text20) ' fecha de pantalla
    CANTIDIAS% = COMBO1.ListIndex
    If CANTIDIAS% <= 0 Then
       CANTIDIAS% = 1
    Else
       CANTIDIAS% = COMBO1.ListIndex + 1
    End If
    
    FE% = DIASPOST(FE%, CANTIDIAS%)

    Dim OBJCMD As ADODB.Command 'El objeto
    Set OBJCMD = New ADODB.Command
    OBJCMD.CommandType = adCmdStoredProc
    OBJCMD.CommandText = "NUEVA_SOLIC_COT_2"
    Set OBJCMD.ActiveConnection = FLEXCONN
    Dim valor_NuevoID As Long

    'Los parametros
    OBJCMD.Parameters(1) = CODIEMPR%
    OBJCMD.Parameters(2) = USNBR%
    OBJCMD.Parameters(3) = Date
    OBJCMD.Parameters(4) = FECHAFORMATOINVERSO$(HOY(HO$))
    OBJCMD.Parameters(7) = CVDAT(FE%)
    
    '
    'Mando a ejecutar el sp
    OBJCMD.Execute
    '
    'Capturo el valor de retorno
    valor_NuevoID = OBJCMD.Parameters(0).Value
    GRABAR_DATOS_SOLICITUD_COT = valor_NuevoID

End Function

Sub GRABAR_DETALLE_SOLICITUD_COT(CI1%, Codigo$, DESCRICION$, VALUNIT#, NUMCLI&, CANTI, REMI$, IDENCAX&, _
    DOMENTRE$, LOCALIDADENTR$, COPOS$, CUITTRANSP$, PCIA$, DOMITRANSP$, LOCALITRANSP$, PATENTE$, DNITRANS$, FE_ORIG_REMITO%)
    FE% = HOY(HO$)
    PVTA& = XVALO(Left$(REMI$, 4))
    NUREM& = XVALO(Mid$(REMI$, 6))
    '
    'OBTENGO EL TIPO DE CAMBIO Y LA MONEDA SEGUN EL NRO DE REMITO PARA GUARDAR EN LA TABLA DETACOT
    x = MonedaY_TipoCambioDeRemitosSegunDoprinum(NUREM&, Moneda%, TIP#)
    MONEX% = Moneda%
    TIPOCAMBIO# = TIP#
    '''''''''''''''''''''''''
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    SQLX$ = "Select * from DETACOT"
25  On Error Resume Next

    rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
    End If
    With rs
       .AddNew
       !CODIEMPR = CODIEMPR%
       !FECHMOV = CVDAT(FE%)
       !FeFormCot = FECHAFORMATOINVERSO$(HOY(HO$))
       !COD_INTE = CI1%
       !Cod_Exte = Codigo$
       !Descrip = Left$(DESCRICION$, 128)
       !VALOUNIT = ROUND(VALUNIT#, 2) 'EN ENVAPLAST AHORA SE ES EN DOLARES SE GUARDA CALCULADO A PESOS (SE CALCULA EN PANTALLA CUANDO CARGA LA INFO DEL REMITO)
       !Nume_Clie = NUMCLI&
       !CANTINGRE = 0
       !CANTSALID = CANTI
       !NUMPVTA = PVTA&
       !DOPRINUM = NUREM&
       !DOPRILAR = Left("RT-X-" & REMI$, 24)
       !ID_ENCA = IDENCAX&
       !DOM_ENTRE = Left(DOMENTRE$, 128)
       !LOCALIDAD_ENTR = Left(LOCALIDADENTR$, 64)
       !COPOSTAL = Left(COPOS$, 24)
       !CUITTRANSP = Left(CUITTRANSP$, 24)
       !PROVINCIA = Left(TRIM$(PCIA$), 24)
       !DOMITRANSP = Left(DOMITRANSP$, 24)
       !LOCALITRANSP = Left(LOCALITRANSP$, 64)
       !PATENTE = Left(PATENTE$, 24)
       !DNITRANS = DNITRANS$
       !FEREMITO = CVDAT(FE_ORIG_REMITO%)
       !CODUNICO = Left("91 R" & REMI$, 24)
       !TIPOCAM = TIPOCAMBIO#
       !MoneVal = MONEX%
       .Update
    End With
    rs.Close
    Set rs = Nothing
    
    ''''''''''''''''''''''''''''
    'NRO DE REMITO QUE DABA ERROR EN TANIT : 256198 CLIENTE 113

'    Dim OBJCMD As ADODB.Command 'El objeto
'    Set OBJCMD = New ADODB.Command
'    OBJCMD.CommandType = adCmdStoredProc
'    OBJCMD.CommandText = "AGREGA_DETACOT_2"
'    Set OBJCMD.ActiveConnection = FLEXCONN
'    Dim valor_NuevoID As Integer
'    'Los parametros
'    OBJCMD.Parameters(1) = CodiEmpr% '@CodiEmpr smallint
'    OBJCMD.Parameters(2) = FETEXCOR1$(CVINT(Date)) '@FechMov smalldatetime
'    OBJCMD.Parameters(3) = FECHAFORMATOINVERSO$(HOY(HO$)) '@FeFormCot nvarchar  (24)
'    OBJCMD.Parameters(4) = CI1% '@Cod_Inte smallint
'    OBJCMD.Parameters(5) = Codigo$ '@Cod_Exte nvarchar  (48)
'    OBJCMD.Parameters(6) = Left$(DESCRICION$, 64) '@Descrip nvarchar  (128)
'    OBJCMD.Parameters(7) = ROUND(VALUNIT#, 2) '@ValoUnit float  (8)
'    OBJCMD.Parameters(8) = NUMCLI& '@Nume_Clie int
'    OBJCMD.Parameters(9) = 0 '@CantIngre float  (8),
'    OBJCMD.Parameters(10) = CANTI ' @CantSalid float  (8),
'    OBJCMD.Parameters(11) = PVTA& '@NUMPVTA smallint, @
'    OBJCMD.Parameters(12) = NUREM& 'DOPRINUM int,
'    OBJCMD.Parameters(13) = "RT-X-" & REMI$ '@DoPriLar nvarchar  (48),
'    OBJCMD.Parameters(14) = IDENCAX& '@ID_ENCA int,
'    OBJCMD.Parameters(15) = DOMENTRE$
'    OBJCMD.Parameters(16) = LOCALIDADENTR$
'    OBJCMD.Parameters(17) = COPOS$
'    OBJCMD.Parameters(18) = CUITTRANSP$
'    OBJCMD.Parameters(19) = TRIM$(PCIA$)
'    OBJCMD.Parameters(20) = DOMITRANSP$
'    OBJCMD.Parameters(21) = LOCALITRANSP$
'    OBJCMD.Parameters(22) = PATENTE$
'    OBJCMD.Parameters(23) = DNITRANS$
'    OBJCMD.Parameters(24) = CVDAT(FE_ORIG_REMITO%) 'FECHA ORIGINAL DEL REMITO
'    OBJCMD.Parameters(25) = "91 R" & REMI$
'    OBJCMD.Execute
'    On Error GoTo 0
'
'    'ALTER  PROCEDURE [dbo].[AGREGA_DETACOT_2] ( , , , , , , , ,     @Code         int = NULL output )
'    'Capturo el valor de retorno
'    valor_NuevoID = OBJCMD.Parameters(0).Value
'    Set OBJCMD = Nothing
    
End Sub

Sub limpiarDatoConductor()
    Text12 = ""
    Text18 = ""
    Text19 = ""
'    Text13 = ""

End Sub

Sub LISTA_DOMI_ENTREGA(NCLIEN)
    
    If NCLIEN < 1 Then
      Call COMUNI("Debe Seleccionar un  Remito, Para Poder Identificar el Cliente")
      GoTo 99
    End If
    Campos$ = "N°SUCUR./F10;DENO_SUC/A32;DOMI_SUC/A32;LOCA_SUC/A32;Cód.Postal/A10"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Width = 13000
    i& = 0
    
    i& = i& + 1
    FRMZELECHO.MSF2.Rows = i& + 1
    FRMZELECHO.MSF2.TextMatrix(i&, 1) = "0"
    FRMZELECHO.MSF2.TextMatrix(i&, 2) = "DOM.ENTREGA DEL CLIENTE"
    Domicilio$ = TRIM$(DOMIENT$(NCLIEN)): If Domicilio$ = "" Then Domicilio$ = "--SIN CARGAR--"
    FRMZELECHO.MSF2.TextMatrix(i&, 3) = TRIM$(DOMIENT$(NCLIEN))
    LOCALENT$ = TRIM$(LOCAENT$(NCLIEN)): If LOCALENT$ = "" Then LOCALENT$ = "--SIN CARGAR--"
    FRMZELECHO.MSF2.TextMatrix(i&, 4) = TRIM$(LOCAENT$(NCLIEN))
    FRMZELECHO.MSF2.TextMatrix(i&, 5) = TRIM$(CodPostal$(NCLIEN))
    
'    FRMZELECHO.msf2.TextMatrix(i&, 5) = TRIM$(TRANENT$(NCLIE))
    FRMZELECHO.Caption = "Seleccione Cual Domicilio de Entrega Modificar"
    SQLX$ = "SELECT* FROM SUCENTRE with(nolock) WHERE Nume_Cli =" & NCLIEN
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        i& = i& + 1
        FRMZELECHO.MSF2.Rows = i& + 1
        FRMZELECHO.MSF2.TextMatrix(i&, 1) = XVALO(!NUME_SUC)
        FRMZELECHO.MSF2.TextMatrix(i&, 2) = "SUC: " & SAFETEXT$(!DENO_SUC)
        FRMZELECHO.MSF2.TextMatrix(i&, 3) = SAFETEXT$(!DOMI_SUC)
        FRMZELECHO.MSF2.TextMatrix(i&, 4) = SAFETEXT$(!LOCA_SUC)
        FRMZELECHO.MSF2.TextMatrix(i&, 4) = SAFETEXT$(!CodPostal_Suc)
'        FRMZELECHO.msf2.TextMatrix(i&, 5) = SAFETEXT$(!Lentre2_Cli)
        .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    
    Erase CADENATEX$
    FRMZELECHO.Show 1
99
End Sub

Sub MENU_CELDA(GRID As msFlexGrid, REG&, COL%)
        If GRID.Rows <= 1 Then Exit Sub
        
'        Call COLOREAR_GRILLA_SOLACELDA(GRID, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL
        GLOBALFRM_strContenidoCelda$ = TRIM$(GRID.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        GLOBALFRM_FILA& = REG&

        PopupMenu mnuArchivo
        Call INVICTRL
        'Call COLOREAR_GRILLA_SOLACELDA(GRID, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA
End Sub

Sub INVICTRL()
  '
  mnucopiar.Visible = Not (mnucopiar.Visible)
  QuitarRemi.Visible = Not (QuitarRemi.Visible)
  '
End Sub

Sub MOSTRAR_SOLICITUD_DE_COT(NROSOLCOT&)

  Call LIMPIA_CAJA_TEXTO
  GRID.Rows = 1
  Dim obj As New RECORXET
  Set RS1 = obj.RS_COTSLECCIONADO(NROSOLCOT&)
'  Call Carga_MSF_Recordset(RS1, "ID Solicitud/F12;Fecha Solicitud", FRMZELECHO.msf2)
  Set obj = Nothing
  With RS1
    Do While Not .EOF

             J& = GRID.Rows
             GRID.Rows = J& + 1
             GRID.TextMatrix(J&, 1) = XVALO(!COD_INTE)
             GRID.TextMatrix(J&, 2) = J&
             GRID.TextMatrix(J&, 3) = SAFETEXT$(!Descrip)
             GRID.TextMatrix(J&, 4) = TRIM$(FORMATNUM$(XVALO(!CANTSALID), "F9.1"))
             GRID.TextMatrix(J&, 5) = TRIM$(FORMATNUM$(XVALO(!VALOUNIT), "F12.2"))
             IMPORTEX = (XVALO(!CANTSALID)) * XVALO(!VALOUNIT)
             ACUMULATOTAL = ACUMULATOTAL + IMPORTEX
             GRID.TextMatrix(J&, 6) = 0
             GRID.TextMatrix(J&, 7) = ""
             GRID.TextMatrix(J&, 8) = Mid$(SAFETEXT$(!DOPRILAR), 6)
             GRID.TextMatrix(J&, 9) = CVDAT(CVINT(!FEREMITO))
             GRID.TextMatrix(J&, 10) = SAFETEXT$(!DOM_ENTRE)
             GRID.TextMatrix(J&, 11) = SAFETEXT$(!LOCALIDAD_ENTR)
              GRID.TextMatrix(J&, 12) = SAFETEXT$(!COPOSTAL)
             GRID.TextMatrix(J&, 13) = "" 'SAFETEXT$(!CUITTRANSP) 'CUIT TRANSPORTE
             GRID.TextMatrix(J&, 14) = SAFETEXT$(!PROVINCIA)
             GRID.TextMatrix(J&, 15) = SAFETEXT$(!DOMITRANSP) 'DOMICILIO TRANSPORTE
             GRID.TextMatrix(J&, 16) = SAFETEXT$(!LOCALITRANSP)  'LOCALIDAD TRANSPORTE
             GRID.TextMatrix(J&, 17) = "" 'SAFETEXT$(!PATENTE) 'PATENTE
             GRID.TextMatrix(J&, 18) = XVALO(!Nume_Clie) 'CLIENTE
             GRID.TextMatrix(J&, 19) = SAFETEXT$(!DNITRANS)  'DNI CLIENTE
             .MoveNext
    Loop
  End With
  RS1.Close
  Set RS1 = Nothing
  Call COMPLETAR_CASILLEROS(NROSOLCOT&)
  Call CalcutotalSeleccionado
  Label18 = GRID.TextMatrix(GRID.Rows - 1, 2)

End Sub
Public Function VALIDAR_DATOS%()
'falta validar datos de la empresa+++++ +++++++**********
    MENSAERROR$ = ""
    
    FALTANCODIGOARBA% = 0
    VALIDAR_DATOS% = 1
    A = rasocli$(1) ' para provocar que cambie el numero de cliente, si no algunas funciones si se repite no devuelven los datos
    Dim oDE As New DatoEmpresa 'instancio cla
    If TRIM$(oDE.CalleNumero) = "" Then MENSAERROR$ = "Falta Nro. De Calle Empresa Activa ": GoTo 20
    If TRIM$(oDE.CodPostal) = "" Then MENSAERROR$ = "Falta Código Postal Empresa Activa ": GoTo 20
    If TRIM$(oDE.Cuit) = "" Then MENSAERROR$ = "Falta Nro. De C.U.I.T. Empresa Activa ": GoTo 20
    If TRIM$(oDE.IngBruto) = "" Then MENSAERROR$ = "Falta Nro. De Ingresos Bruto Empresa Activa ": GoTo 20
    If TRIM$(oDE.LetProvincia) = "" Then MENSAERROR$ = "Falta Indicar Provincia de Empresa Activa ": GoTo 20
    If TRIM$(oDE.Localidad) = "" Then MENSAERROR$ = "Falta Localidad Empresa Activa ": GoTo 20
    If TRIM$(oDE.NombreEmpresa) = "" Then MENSAERROR$ = "Falta Nombre de Empresa Activa ": GoTo 20
    If TRIM$(oDE.razonSocial) = "" Then MENSAERROR$ = "Falta Razon Social Empresa Activa ": GoTo 20
    '
    If TRIM$(CONTROL$("REMITO", "PASSCOT")) = "" Then
      Call COMUNI("" & MENSAERROR$ & "\Falta Password Arba de Obtención de C.O.T.\Estos Datos se Cargan Desde Esta Pantalla > Menú Configuración > Configuraciones C.O.T.")
      VALIDAR_DATOS% = -1
      Exit Function
    End If
    
20  If MENSAERROR$ <> "" Then
      Call COMUNI("" & MENSAERROR$ & "\Estos Datos se Cargan Desde la Pantalla Inicial del Sistema\Botón Setup\Botón Empresa")
      VALIDAR_DATOS% = -1
      Exit Function
    End If
    '
    CoArba$ = TRIM$(CONTROL("REMITO", "CODIARBA"))
    If CoArba$ = "" Then
       MENSAERROR$ = "Falta Código Genérico Arba\Agregar de Esta Pantalla > Menú > Configuración > Configuración C.O.T.":
       Call COMUNI(MENSAERROR$)
       VALIDAR_DATOS% = -1
       Exit Function
    End If
    
    For i& = 1 To GRID.Rows - 1
      NREMI1$ = TRIM$(GRID.TextMatrix(i&, 8))
      ENFILAX$ = "En Fila: " & i&
      ENRemito$ = " En Remito: " & NREMI1$
      If NREMI1$ = "" Then MENSAERROR$ = "Falta Nro. De Remito ": GoTo 80
      If NREMI1$ <> NREMI_ANT$ Then
        TX1$ = TX1$ & "02" & "|"
        FEREM% = FECHANUM(TRIM$(GRID.TextMatrix(i&, 9)))
        If FEREM% < 34 Then MENSAERROR$ = "Falta Fecha Remito ": GoTo 80
        FECHREMI$ = FECHAFORMATOINVERSO$(FEREM%)
        ImporteRemito$ = TRIM$(CalcutotalRemito$(NREMI1$))
        If XVALO(ImporteRemito$) < 0.05 Then MENSAERROR$ = "Falta Importe ": GoTo 80
        NC = XVALO(GRID.TextMatrix(i&, 18))
        If NC < 1 Then MENSAERROR$ = "Falta Número de Cliente ": GoTo 80
        PREFIJO$ = Left(GRID.TextMatrix(i&, 8), 4)
        If XVALO(PREFIJO$) < 1 Then MENSAERROR$ = "Falta Prefijo de Remito ": GoTo 80
        NREMI$ = Mid$(GRID.TextMatrix(i&, 8), 6)
        If XVALO(NREMI$) < 1 Then MENSAERROR$ = "Falta Nro. de Remito ": GoTo 80
        HORAX$ = Format(Time, "HH:MM")
        LETRA$ = TRIM$(Text7.TEXT) 'E POR QUE ES ENVIO
        If LETRA$ = "E" Then
          VentaoDevolucion$ = "0" 'venta
          ImporteRemito$ = TRIM$(Label3)
          CUITX$ = TRIM$(CuitCli$(NC))
          RazSocSale$ = TRIM$(oDE.razonSocial)
        Else
          VentaoDevolucion$ = "1" 'devolucion
          ImporteRemito$ = "0"
          CUITX$ = TRIM$(oDE.Cuit) ' CUIT DE EMPRESA
          RazSocSale$ = TRIM$(rasocli$(NC))

        End If
        If CUITX$ = "" Then MENSAERROR$ = "Falta Nro. de C.U.I.T. De Cliente": GoTo 80
        If PIVACLI(NC) < 1 Then
          DESTINATARIO$ = "1"
          NombreEmpresa$ = TRIM$(rasocli$(NC))
          NDOCUM$ = TRIM$(GRID.TextMatrix(i&, 19))
          If NDOCUM$ = "" Then
                MENSAERROR$ = "Falta Nro. de Documento Del Cliente \Cliente Consumidor Final "
                Call COMUNI(MENSAERROR$)
                VALIDAR_DATOS% = -1
                Exit Function
          End If

        Else
          DESTINATARIO$ = "0"
          NombreEmpresa$ = ""
        End If
        
        TIPODOCUMENTA$ = "DNI"
        '
        
        '
          DomicilioEntrega$ = TRIM$(GRID.TextMatrix(i&, 10))
        If DomicilioEntrega$ = "" Then MENSAERROR$ = "Falta Domicilio de Entrega ": GoTo 80
         CODIPOSTAL$ = TRIM$(GRID.TextMatrix(i&, 12))
         If CODIPOSTAL$ = "" Then MENSAERROR$ = "Falta Código Postal de Entrega ": GoTo 80
         LocalidadEntrega$ = TRIM$(GRID.TextMatrix(i&, 11))
         If LocalidadEntrega$ = "" Then MENSAERROR$ = "Falta Localidad de Entrega ": GoTo 80
         PCIA$ = TRIM$(LetraProvArbaSegunLetraFlexoft$(TRIM$(GRID.TextMatrix(i&, 14))))
         If PCIA$ = "" Then MENSAERROR$ = "Falta Provincia de Entrega ": GoTo 80
        '
        If LETRA$ = "E" Then
           CodPostSale$ = TRIM$(oDE.CodPostal)
           LocalSale$ = TRIM$(oDE.Localidad)
           LetraSale$ = TRIM$(LetraProvArbaSegunLetraFlexoft$(oDE.LetProvincia))
        Else
           CodPostSale$ = TRIM$(CODIPOSTAL$)
           LocalSale$ = TRIM$(LocalidadEntrega$)
           LetraSale$ = TRIM$(PCIA$)
        End If
        '
        CuitTransportista$ = Text12.TEXT 'TRIM$(GRID.TextMatrix(i&, 13))
        If CuitTransportista$ = "" Then
           MENSAERROR$ = "Falta C.U.I.T. Transportista "
           Call COMUNI(MENSAERROR$)
           VALIDAR_DATOS% = -1
           Exit Function
        End If
        
        PATENTE$ = Text18.TEXT 'TRIM$(GRID.TextMatrix(i&, 17))
        If PATENTE$ = "" Then
           MENSAERROR$ = "Falta Patente Transporte "
           Call COMUNI(MENSAERROR$)
           VALIDAR_DATOS% = -1
           Exit Function
        End If
        
        NREMI_ANT$ = NREMI1$
      End If
      CI1% = XVALO(GRID.TextMatrix(i&, 1))
      If CI1% < 1 Then MENSAERROR$ = "Falta Código de Producto ": GoTo 80
      '
      If CodigoArba$(CI1%) = "" Then FALTANCODIGOARBA% = 1
      '
      X1X = UnidMedidaArbaSegunUnidFlexoft$(Unimed$(CI1%), RETNUM%)
      CANTI = TRIM$(FORMATNUM$(XVALO(GRID.TextMatrix(i&, 4)), "F16.2"))
      TX1$ = TX1$ & SAFETEXT$(CANTI) & "|" '50
      CODI$ = CODEXT$(CI1%)
      TX1$ = TX1$ & CODI$ & "|" '51
      DESCRICION$ = TRIM$(Left$(GRID.TextMatrix(i&, 3), 64))
    Next i&
    Set oDE = Nothing '
    If FALTANCODIGOARBA% > 0 Then
       Call COMUNI("Se han Detectado Códigos Sin Asociación al Código Arba\Se Reemplazaran estos Por el Código Genérico")
    End If
    Exit Function
    '********************************************************
    
80  Call COMUNI("" & MENSAERROR$ & ENFILAX$ & ENRemito$)
    VALIDAR_DATOS% = -1
    Set oDE = Nothing '
End Function

Private Sub Command10_Click()
    Command10.Enabled = False
    NCLIEN = XVALO(Text2)
    Call LISTA_DOMI_ENTREGA(NCLIEN)
    On Error Resume Next
    SUCU1$ = RESP_ZELE_VECT(1)
    If SUCU1$ <> "" Then
      LENTREGA(0) = RESP_ZELE_VECT(3)
      LENTREGA(1) = RESP_ZELE_VECT(4)
      TXTCPOSCLI = RESP_ZELE_VECT(5)
    End If
    On Error GoTo 0
99  Command10.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    If GRID.Rows <= 1 Then
      Call COMUNI("No se Ha Seleccionado Ningún Remito")
      GoTo 99
    End If
    IDDETCOT& = XVALO(Text15) 'VEO SI YA FUE SOLICITADO
    If IDDETCOT& > 0 Then
      CONDI2$ = "ENCACOT_ID =" & IDDETCOT& 'VEO SI OBTUVO COT
      COTX$ = TRIM$(VALOBADA("ENCACOT", "COT", CONDI2$, "NOMESS"))
      If COTX$ <> "" Then
         Call COMUNI("Atención !!!! \Remito: " & NUREMIX$ & " Ya Fué Procesado la Solicitud y Obtuvo C.O.T.")
         GoTo 99
      End If
    End If
    
    FE% = HOY(HO$)
    If FECHANUM(Text20) < 1 Then
      Call COMUNI("Fecha de Salida Inválida")
      GoTo 99
    End If
    If FECHANUM(Text20) < FE% Then
      Call COMUNI("Fecha de Salida No Puede Ser Inferior a la Fecha Actual")
      GoTo 99
    End If
    If FECHANUM(Text20) > DIASPOST(FE%, 7) Then
      OPX% = COMALTER%("Fecha de Salida Supera en Más de 7 Días a la Fecha Actual\\Cancelar\Continuar")
      If OPX% <= 1 Then
         GoTo 99
      End If
    End If
    A = COMBO1.ListIndex
    '0 SE VALIDAN LOS DATOS SI NO NO PASA
    '1 SE INICIA LA TRANSACCION EN SUSPENSO
    '2 SE OBTIENE EL PROXIMO VALOR DE ID DE LA TABLA CABECERA
    '3 SE PASAN LOS DATOS DETALLE A LA TABLA DETACOT CON EL MISMO IDENTIFICADOR DE LA TABLA CABECERA (ENCACOT)
    '4 SE GENERA EL ARCHIVO TXT EN C:\FLEXOFT
    '5
  
    If VALIDAR_DATOS% < 1 Then GoTo 99
    
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans
    IDENCAX& = GRABAR_DATOS_SOLICITUD_COT ' ESTA FUNCION LLAMA A UN SP QUE LE DEVUELVE EL ID DEL REGISTRO GENERADO
    If IDENCAX& > 0 Then
      For J& = 1 To GRID.Rows - 1
             'GRABO EL DETALLE
             CI1% = XVALO(GRID.TextMatrix(J&, 1))
             Codigo$ = CODEXT$(CI1%)
             DESCRICION$ = Left$(GRID.TextMatrix(J&, 3), 64)
             CANTI = XVALO(GRID.TextMatrix(J&, 4))
             VALUNIT# = XVALO(GRID.TextMatrix(J&, 5))
             FEREM% = FECHANUM(GRID.TextMatrix(J&, 9))
             REMI$ = GRID.TextMatrix(J&, 8)
             NUMCLI& = XVALO(GRID.TextMatrix(J&, 18))
             DOMENTRE$ = GRID.TextMatrix(J&, 10)
             LOCALIDADENTR$ = GRID.TextMatrix(J&, 11)
             COPOS$ = GRID.TextMatrix(J&, 12)
             CUITTRANSP$ = Text12.TEXT 'GRID.TextMatrix(J&, 13)
             PCIA$ = GRID.TextMatrix(J&, 14)
             DOMITRANSP$ = GRID.TextMatrix(J&, 15) 'DOMICILIO TRANSPORTE
             LOCALITRANSP$ = GRID.TextMatrix(J&, 16)  'LOCALIDAD TRANSPORTE
             PATENTE$ = Text18.TEXT 'GRID.TextMatrix(J&, 17)  'PATENTE
             DNITRANS$ = GRID.TextMatrix(J&, 19)   'DNI TRANSPORTISTA
             Call GRABAR_DETALLE_SOLICITUD_COT(CI1%, Codigo$, DESCRICION$, VALUNIT#, NUMCLI&, CANTI, REMI$, IDENCAX& _
             , DOMENTRE$, LOCALIDADENTR$, COPOS$, CUITTRANSP$, PCIA$, DOMITRANSP$, LOCALITRANSP$, PATENTE$, DNITRANS$, FEREM%)
      Next J&
    Else
      Call COMUNI("No se ha Podido Grabar el Encabezado de la Tabla de Registros del C.O.T.")
      GoTo 99
    End If
    '
    
    ID_COT& = IDENCAX& ' VARIABLE A NIVEL DE FRM formtato numerico
    Call Generar_Archivo_Cot(IDENCAX&, ret_DATOSCOT$, ret_NAMEARCH$) 'ret_CARPETOK$ retorna la ruta y nombre del archivo creado.
    IDCOT$ = ret_DATOSCOT$ ' VARIABLE A NIVEL DE FRM para pasar los parametros a .net
    NAMEARCHIVO$ = "COT_" + ret_NAMEARCH$ 'variable con el nombre del archivo que voy a buscar que genera el .net
    ModalidadCOT$ = CONTROL("REMELEC", "MODOCOT")
    XXX = Shell("C:\Flexoft\COT\" + "COT.exe " & ret_DATOSCOT$ & " " & ModalidadCOT$, 4)
    
    Timer1.Enabled = True
    For R& = 1 To 40
      Call T_Espera(1 / 2)
      If Timer1.Enabled = False Then
         GoTo 50
      End If
    Next R&
50  If ID_COT& = 0 Then 'SE GENERO EL COT
       FLEXCONN.CommitTrans
    Else 'NO SE GENERO EL COT
       FLEXCONN.RollbackTrans
    End If
    GoTo 99
    
ERROR_GUARDAR:

    If Err.Number <> 0 Then
       MsgBox Err.Description
           FLEXCONN.RollbackTrans
           Call COMUNI("Se Produjo un Error al Grabar Tabla de Registros del C.O.T.\No se Ha Realizado Ninguna Transacción")
           On Error Resume Next
           GoTo 99
        Else
           FLEXCONN.CommitTrans
        End If
99      Command4.Enabled = True


End Sub
Private Sub Command6_Click()
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "Id/A10;Provincia/A30", FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Id/A10;Provincia/A30", FRMZELECHO)
   FRMZELECHO.Caption = "Listado de Provincias"
   For i& = 1 To ULTREG&("PROVINC")
     x$ = REGLEIDO$("PROVINC", i&)
     LETRA$ = Left$(x$, 1)
     PROVINICIA$ = TRIM$(Mid$(x$, 2))
     FRMZELECHO.MSF2.Rows = i& + 1
     FRMZELECHO.MSF2.TextMatrix(i&, 1) = LETRA$
     FRMZELECHO.MSF2.TextMatrix(i&, 2) = PROVINICIA$
   Next i&
   Call CIERRARCH("PROVINC")
   
   
   FRMZELECHO.Width = 6500
   FRMZELECHO.Show 1
   LETR$ = RESP_ZELE_VECT(1)
   '
   If LETR$ = "" Then Exit Sub
   
'    Call SELECHO("PROVINC")
'    VDEVU$ = VALACT1$("PROVINC")
    If TRIM$(LETR$) <> "" Then
       Text8.TEXT = LETR$
    End If
End Sub


Private Sub Command7_Click()

     If TRIM$(Text16) = "" Then Command5_Click: GoTo 99 ' SI NO TIENE MES SELECCIONADO VA AL BOTON DE ARRIBA DONDE PRESENTA A PARTIR DEL MES
     Command7.Enabled = False
     '
     MESYANO$ = TRIM$(Text16)
     CORRESQ$ = MESANO$(MESYANO$, Des%, Has%)
     '
     Desde = FETEXCOR1$(Des%)
     Hasta = FETEXCOR1$(Has%)
     '
     CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
     If Len(CMO$) <> 2 Then
        CMO$ = "RT"
     End If
     CONDI$ = "FechMov >= '" & Desde & "'"
     CONDI$ = CONDI$ + "AND FechMov <= '" & Hasta & "'"
     CONDI$ = CONDI$ + " AND substring(REFERCOR,1,4) <> 'ANUL' " 'PARA EVITAR COMPROBANTES ARRUINADOS
     CONDI$ = CONDI$ + " AND CodiMovi =  '" & CMO$ & "'"
     CONDI$ = CONDI$ + " AND substring(DOPRILAR,1,2) = 'RT' "
'     CONDI$ = CONDI$ + " AND Nume_Clie = " & XVALO(Text2)
     '
     Screen.MousePointer = 11
     '
     Call COPYSTRU("SELDOCAP", "SELREMIT")
12   Call CARGALISEL("SELREMIT", "MOVISTO", "DOPRILAR/A18;FechMov/D8;REFERCOR/A24", CONDI$, "DISTINCT")
     '
     If ULTREG&("SELREMIT") < 1 Then
        Call MENSERR(24, "No se Encontraron Remitos en el Rango Seleccionado.")
        GoTo 99
     End If

15   Call SELECHO("SELREMIT")
     '
     Screen.MousePointer = 1
     '
     NUDOCU$ = TRIM$(Mid$(VALACT1$("SELREMIT"), 6))
     If NUDOCU$ = "" Then GoTo 99
     MEXAGE$ = "Remito Existente en Grilla"
     
     If VALIDA_COD_GRID%(NUDOCU$, Me.GRID, 8, MEXAGE$) < 1 Then
        Text3 = NUDOCU$
     End If
   
99  Command7.Enabled = True
End Sub

Private Sub Command8_Click()

    Command8.Enabled = False
    NCLIEN = XVALO(Text2)
    Call LISTA_DOMI_ENTREGA(NCLIEN)

    NSUC$ = TRIM$(RESP_ZELE_VECT(1))
    If NSUC$ = "" Then GoTo 99
    NUMESUC% = XVALO(NSUC$)
    If NUMESUC% > 0 Then
       CONDI$ = "NUME_SUC = " & NUMESUC% & " AND Nume_Cli = " & NCLIEN
       Call ACTUREGISTRO("SUCENTRE", "DOMI_SUC", CONDI$, LENTREGA(0).TEXT, REGAF&, "NOMESS")
       Call ACTUREGISTRO("SUCENTRE", "LOCA_SUC", CONDI$, LENTREGA(1).TEXT, REGAF&, "NOMESS")
       Call ACTUREGISTRO("SUCENTRE", "CodPostal_Suc", CONDI$, Left$(TRIM$(TXTCPOSCLI.TEXT), 10), REGAF&, "NOMESS")
    End If
    If NSUC$ = "0" Then 'VARIABLE DE TEXTO PARA DIFERENCIAR SI ELIGIO EL NRO DE SUCURSAL 0 O SALIO SIN SELECCIONAR
        strSQL = "Update Deudor12 set Lentre0_Cli ='" & TRIM$(LENTREGA(0).TEXT) & "'"
        strSQL = strSQL & ", Lentre1_Cli ='" & TRIM$(LENTREGA(1).TEXT) & "'"
        strSQL = strSQL & ", Lentre2_Cli ='" & TRIM$(LENTREGA(2).TEXT) & "'"
        strSQL = strSQL & ", Lentre3_Cli ='" & Left$(TRIM$(TXTCPOSCLI.TEXT), 10) & "'"
        strSQL = strSQL & " where Nume_Cli = " & NCLIEN
        FLEXCONN.Execute (FILTSQL$(strSQL))
    End If
    '
99  Command8.Enabled = True
End Sub

Private Sub Config_Click()
   STPCOT07.Show 1
End Sub

Private Sub GRID_Click()
    i& = GRID.MouseRow
    J& = GRID.MouseCol
    If GRID.Rows = 1 Then Exit Sub
    If i& = 0 Or i& > GRID.Rows Then Exit Sub
    '
    On Error Resume Next
    GRID.Row = i&
    GRID.COL = J&
    On Error Resume Next
    
    REG& = i&
    If REG& < 1 Then
        Exit Sub
    End If
    '
    If Me.lblRemSel.Caption = GRID.TextMatrix(i&, 8) Then
       Call lblRemSel_Change
    Else
      Me.lblRemSel.Caption = GRID.TextMatrix(i&, 8)
    End If
    
End Sub

Private Sub Label3_DblClick()
'NO HICE ESTA MODIFICACION POR QUE EL IMPORTE QUE PASA AL COT LO CALCULA CON LOS DATOS DE LA GRILLA
'     IMPORTE$ = XVALO(Label3.Caption)
'     IMPO# = XVALO(InputBox("Ingrese Importe", "Importe", IMPORTE$))
'     IMPORTE$ = FORMATNUM$(XVALO(IMPORTE$), "F12.2")
'     Label3 = IMPORTE$
End Sub

Private Sub lblRemSel_Change()
  Call LIMPIA_CAJA_TEXTO
  Call AgregarDatosDeGrillaaText(Me.lblRemSel)
  Call CalcutotalSeleccionado
'  Call AgregarDatosdeEntregaAGrilla(Me.lblRemSel)
End Sub

Private Sub mnucopiar_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub
Private Sub Command1_Click()
     Command1.Enabled = False
     Text15 = 0: Text14.Visible = False: Text15.Visible = False ' ENCABEZADO QUE SE USA PARA MOSTRAR DATOS DE UN COT EN ESPECIAL
'     If TRIM$(Text3) = "" Then
'       Call COMUNI("Imposible Agregar Sin Remito Seleccionado Previamente")
'       GoTo 99
'     End If
     
     Call CARSELMESA
     
     '
10   Screen.MousePointer = 1
     Call SELECHO("SELMESA")
     MESASEL$ = VALACT1$("SELMESA")
     If TRIM$(MESASEL$) = "" Then
        GoTo 99
     End If
     Call ABRESTOCKS
     '
     CORRESQ$ = MESANO$(MESASEL$, Des%, Has%)
     '
     Text16 = MESASEL$
     Desde = FETEXCOR1$(Des%)
     Hasta = FETEXCOR1$(Has%)
     '
     CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
     If Len(CMO$) <> 2 Then
        CMO$ = "RT"
     End If
     CONDI$ = "FechMov >= '" & Desde & "'"
     CONDI$ = CONDI$ + "AND FechMov <= '" & Hasta & "'"
     CONDI$ = CONDI$ + " AND substring(REFERCOR,1,4) <> 'ANUL' " 'PARA EVITAR COMPROBANTES ARRUINADOS
     CONDI$ = CONDI$ + " AND CodiMovi =  '" & CMO$ & "'"
     CONDI$ = CONDI$ + " AND substring(DOPRILAR,1,2) = 'RT' "
'     CONDI$ = CONDI$ + " AND Nume_Clie = " & XVALO(Text2)
     '
     Screen.MousePointer = 11
     '
     Call COPYSTRU("SELDOCAP", "SELREMIT")
12   Call CARGALISEL("SELREMIT", "MOVISTO", "DOPRILAR/A18;FechMov/D8;REFERCOR/A24", CONDI$, "DISTINCT")
     '
     If ULTREG&("SELREMIT") < 1 Then
        Call MENSERR(24, "No se Encontraron Remitos en el Rango Seleccionado.")
        GoTo 10
     End If

15   Call SELECHO("SELREMIT")
     '
     Screen.MousePointer = 1
     '
     NUDOCU$ = TRIM$(Mid$(VALACT1$("SELREMIT"), 6))
     If NUDOCU$ = "" Then GoTo 10
     MEXAGE$ = "Remito Existente en Grilla"
     If VALIDA_COD_GRID%(NUDOCU$, Me.GRID, 9, MEXAGE$) < 1 Then
        Text3 = NUDOCU$
     End If
     
    
99   Command1.Enabled = True

End Sub

Private Sub Command16_Click()
   Command16.Enabled = False
   '
   If DERACCE%("ATRANSPO", "Actualizar Base de Datos de Transportes") < 2 Then
      GoTo 99
   End If
   '
   VDEF$ = REGBLAN$("TRANSPOR")
   CONDI$ = "CODI_TRA > 0"
10 PROXI% = 1 + XVALO(VALOBADA("TRANSPORTES", "MAX(CODI_TRA)", CONDI$, "NOMESS"))
   Call REPLA(VDEF$, MKI$(PROXI%), 1, 2)
   '
   OBX$ = OBJPLA$("CARGA-TRANSPOR", VDEF$)
   If OBX$ <> "" Then
     If TRIM$(Mid$(OBX$, 3, 32)) = "" Then
       Call MENSERR(24, "Falta Nombre del Transporte")
       VDEF$ = OBX$
       GoTo 10
     End If
     '
     Call BLOQARCH("TRANSPOR")
     PROXI% = 1 + XVALO(VALOBADA("TRANSPORTES", "MAX(CODI_TRA)", CONDI$, "NOMESS"))
     Call REPLA(OBX$, MKI$(PROXI%), 1, 2)
     Call REGAPP("TRANSPOR", OBX$)
     Call CIERRARCH("TRANSPOR")
     ' Graba en Tabla SQL Transportes.
     COD1% = PROXI%
     DENO$ = TRIM$(Mid$(OBX$, 3, 32))
     RASO$ = TRIM$(Mid$(OBX$, 35, 48))
     DOMI$ = TRIM$(Mid$(OBX$, 83, 48))
     LOCA$ = TRIM$(Mid$(OBX$, 131, 48))
     Cuit$ = TRIM$(Mid$(OBX$, 179, 15))
     PIVA% = Asc(Mid$(OBX$, 194, 1))
     TELE$ = TRIM$(Mid$(OBX$, 195, 48))
     FAXI$ = TRIM$(Mid$(OBX$, 243, 48))
     mail$ = TRIM$(Mid$(OBX$, 291, 48))
     CTAC$ = TRIM$(Mid$(OBX$, 339, 48))
     Dim TRANSPOR As ADODB.Recordset
     Set TRANSPOR = New ADODB.Recordset
     SQLX$ = "Select * from TRANSPORTES"
25   On Error Resume Next
     TRANSPOR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
     End If
     With TRANSPOR
          If COD1% > 0 Then
                 .AddNew
                 !CODI_TRA = COD1%
                 !DENO_TRA = DENO$
                 !RASO_TRA = RASO$
                 !DOMI_TRA = DOMI$
                 !LOCA_TRA = LOCA$
                 !CUIT_TRA = Cuit$
                 !PIVA_TRA = PIVA%
                 !TELE_TRA = TELE$
                 !FAXI_TRA = FAXI$
                 !MAIL_TRA = mail$
                 !CTAC_TRA = CTAC$
                 .Update
          End If
     End With
     '
   End If
99 Command16.Enabled = True

End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   '
   If DERACCE%("ATRANSPO", "Actualizar Base de Datos de Transportes") < 2 Then
      GoTo 99
   End If
   '
   NTRA = XVALO(Text45)
   If NTRA > 0 Then
     If NTRA < 32767 Then
       NTRAI% = NTRA
       'If ECOARCH$("TRANSPOR", MKI$(NTRAI%)) <> "" Then
       
       'MODIFICADO POR QUE EN GABAL SI GRABA EN UNA EMPRESA
       'NO LE DEJA EDITAR POR QUE NO ENCUENTRA EL ECOARCH
       CONDI$ = "CODI_TRA = " & NTRAI%
       If CantRegistros("TRANSPORTES", CONDI$) > 0 Then
         'ASIGNO CON LA FUNICION SI NO ENCUENTRA EN SQL TRAE DE DAT
         '*********************************************************
         CUITRA$ = CUITTRANS$(NTRAI%)
         DENOTRA$ = DENOTRANS$(NTRAI%)
         RASOTRA$ = RASOTRANS$(NTRAI%)
         DOMITRA$ = DOMITRANS$(NTRAI%)
         LOCATRA$ = LOCATRANS$(NTRAI%)
         TELETRA$ = TELETRANS$(NTRAI%)
         FAXITRA$ = FAXITRANS$(NTRAI%)
         MAILTRA$ = MAILTRANS$(NTRAI%)
         CTACTRA$ = CTACTRANS$(NTRAI%)
         PIVATRA% = PIVATRANS%(NTRAI%)
         '
         'ARMO EL STRING
         '**************
         Call REPLA(VDEF$, MKI$(NTRAI%), 1, 2)
         Call REPLA(VDEF$, DENOTRA$, 3, 32)
         Call REPLA(VDEF$, RASOTRA$, 35, 48)
         Call REPLA(VDEF$, DOMITRA$, 83, 48)
         Call REPLA(VDEF$, LOCATRA$, 131, 48)
         Call REPLA(VDEF$, CUITRA$, 179, 15)
         Call REPLA(VDEF$, Chr$(PIVATRA%), 194, 1)
         Call REPLA(VDEF$, TELETRA$, 195, 48)
         Call REPLA(VDEF$, FAXITRA$, 243, 48)
         Call REPLA(VDEF$, MAILTRA$, 291, 48)
         Call REPLA(VDEF$, CTACTRA$, 339, 48)

         'VDEF$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRAI%))
10       OBX$ = OBJPLA$("CARGA-TRANSPOR", VDEF$)
         If OBX$ <> "" Then
           If TRIM$(Mid$(OBX$, 3, 32)) = "" Then
             Call MENSERR(24, "Falta Nombre del Transporte")
             VDEF$ = OBX$
             GoTo 10
           End If
           '
           Call BLOQARCH("TRANSPOR")
           Call FILTERPUTRECORD("TRANSPOR", 1, MKI$(NTRAI%), OBX$)
           Call CIERRARCH("TRANSPOR")
           ' Edita y Graba en Tabla SQL Transportes.
           COD1% = NTRAI%
           DENO$ = TRIM$(Mid$(OBX$, 3, 32))
           RASO$ = TRIM$(Mid$(OBX$, 35, 48))
           DOMI$ = TRIM$(Mid$(OBX$, 83, 48))
           LOCA$ = TRIM$(Mid$(OBX$, 131, 48))
           Cuit$ = TRIM$(Mid$(OBX$, 179, 15))
           PIVA% = Asc(Mid$(OBX$, 194, 1))
           TELE$ = TRIM$(Mid$(OBX$, 195, 48))
           FAXI$ = TRIM$(Mid$(OBX$, 243, 48))
           mail$ = TRIM$(Mid$(OBX$, 291, 48))
           CTAC$ = TRIM$(Mid$(OBX$, 339, 48))
'           SQLX$ = "Codi_Tra = " & COD1% & ""
           
          Dim RS1 As ADODB.Recordset
          Set RS1 = New ADODB.Recordset
          SQLX$ = "Select * from TRANSPORTES"
          SQLX$ = SQLX$ + " WHERE Codi_Tra = " & COD1%
30        On Error Resume Next
          RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
            On Error GoTo 0
            Call CapturaErrorOpen
            GoTo 30
          End If
          On Error GoTo 0
        
          '
          With RS1
           Do While Not .EOF
              !DENO_TRA = Left$(TRIM$(DENO$), 128)
              !RASO_TRA = Left$(TRIM$(RASO$), 128)
              !DOMI_TRA = Left$(TRIM$(DOMI$), 128)
              !LOCA_TRA = Left$(TRIM$(LOCA$), 128)
              !CUIT_TRA = Left$(TRIM$(Cuit$), 32)
              !PIVA_TRA = PIVA%
              
              
              !TELE_TRA = Left$(TRIM$(TELE$), 64)
              !FAXI_TRA = Left$(TRIM$(FAXI$), 64)
              !MAIL_TRA = Left$(TRIM$(mail$), 64)
'              !MAIL_TRA = Left$(TRIM$(Text10(3)), 128)
              !CTAC_TRA = Left$(TRIM$(CTAC$), 128)
              .Update
              .MoveNext
           Loop
          End With
          RS1.Close
          Set RS1 = Nothing
           '
           Call Text45_Change
         End If
         GoTo 99
       End If
     End If
   End If
   '
   Call MENSERR(24, "Imposible Editar Datos Transportista")
   '
99 Command18.Enabled = True

End Sub


Private Sub Command19_Click()
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "Código Transp./F14;Denominación/A40", FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Código Transp./F14;Denominación/A40", FRMZELECHO)
   FRMZELECHO.Caption = "Listado de Transportistas"
   For i& = 1 To ULTREG&("TRANSPOR")
     x$ = REGLEIDO$("TRANSPOR", i&)
     CODTRAN% = CVI(Left$(x$, 2))
     DENOTRANS1$ = TRIM$(Mid$(x$, 3, 32))
     FRMZELECHO.MSF2.Rows = i& + 1
     FRMZELECHO.MSF2.TextMatrix(i&, 1) = CODTRAN%
     FRMZELECHO.MSF2.TextMatrix(i&, 2) = DENOTRANS1$
   Next i&
   Call CIERRARCH("TRANSPOR")
   
   
   FRMZELECHO.Width = 7200
   FRMZELECHO.Show 1
   NTRAN% = XVALO(RESP_ZELE_VECT(1))
   '
   If NTRAN% > 0 Then
        Text45 = TRIM$(Str$(NTRAN%))
   End If

End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call AgregarDatosdeEntregaAGrilla(Me.lblRemSel)
   Command2.Enabled = True
End Sub

Private Sub Command3_Click()
   Me.LENTREGA(0) = Me.DOMITRAN(3)
   Me.LENTREGA(1) = Me.DOMITRAN(4)

End Sub

Private Sub Command5_Click()
     Command5.Enabled = False
     Call Command1_Click
     Command5.Enabled = True
     Exit Sub
     
     Call CARSELMESA
     '
10   Screen.MousePointer = 1
     Call SELECHO("SELMESA")
     MESASEL$ = VALACT1$("SELMESA")
     If TRIM$(MESASEL$) = "" Then
        GoTo 99
     End If
     Call ABRESTOCKS
     '
     CORRESQ$ = MESANO$(MESASEL$, Des%, Has%)
     '
     Desde = FETEXCOR1$(Des%)
     Hasta = FETEXCOR1$(Has%)
     '
     CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
     If Len(CMO$) <> 2 Then
        CMO$ = "RT"
     End If

     CONDI$ = "FechMov >= '" & Desde & "'"
     CONDI$ = CONDI$ + "AND FechMov <= '" & Hasta & "'"
     CONDI$ = CONDI$ + " AND substring(REFERCOR,1,4) <> 'ANUL' " 'PARA EVITAR COMPROBANTES ARRUINADOS
     CONDI$ = CONDI$ + " AND CodiMovi =  '" & CMO$ & "'"
     CONDI$ = CONDI$ + " AND substring(DOPRILAR,1,2) = 'RT' "
     '
     Screen.MousePointer = 11
     '
     Call COPYSTRU("SELDOCAP", "SELREMIT")
12   Call CARGALISEL("SELREMIT", "MOVISTO", "DOPRILAR/A18;FechMov/D8;REFERCOR/A24", CONDI$, "DISTINCT")
     '
     If ULTREG&("SELREMIT") < 1 Then
        Call MENSERR(24, "No se Encontraron Remitos en el Rango Seleccionado.")
        GoTo 10
     End If

15   Call SELECHO("SELREMIT")
     '
     Screen.MousePointer = 1
     '
     NUDOCU$ = TRIM$(Mid$(VALACT1$("SELREMIT"), 6))
     If NUDOCU$ = "" Then GoTo 10
     Me.Text3 = NUDOCU$
    
    
99   Command5.Enabled = True
End Sub

Private Sub Form_Load()
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    Campos$ = "CI/F0;Ord./F4;Descripción/A32;Salida/F9.1;P.Unit./F12.2;Nro.Ped./F7;"
    Campos$ = Campos$ + "N°Serie/A0;Remito/A16;F.Remito/D8;Domicilio/A20;Localidad/A20;C.P./A6;"
    Campos$ = Campos$ + "CUIT/A0;PCIA/A20;DOMITRNAS/A0;LOCATRANS/A0;PATENTE/A0;Cte/F8;Dni/A12"
   Call CARGA_ENCABEZADO(Me.GRID, Campos$, Me)
    
    ModalidadCOT$ = CONTROL("REMELEC", "MODOCOT")
    
    COMBO1.AddItem "Menos de 500"
    COMBO1.AddItem "Entre 500 y menos de 1000"
    COMBO1.AddItem "Igual o más de 1000"
'    Combo1.TEXT = "Menos de 500"
    COMBO1.ListIndex = 0
    'VALIDO SI EXISTE LA CARPETA SI NO LA CREO
    CARPEACTUAL$ = App.Path + "\"
    CARPEDESTIN$ = "C:\FLEXOFT\COT\"
    On Error Resume Next
    MkDir "C:\FLEXOFT"
    MkDir "C:\FLEXOFT\COT"
    On Error GoTo 0
    If EXISTE%("C:\FLEXOFT\COT\.") < 1 Then
       Call MENSERR(24, "Imposible crear carpeta 'C:/FLEXOFT/COT'")
       GoTo 99
    End If
    '
    'Eliminar archivos
    If EXISTE%(CARPEDESTIN$ + "COT.EXE") > 0 Then
       FILEKILL (CARPEDESTIN$ + "COT.EXE")
    End If
    '
    If EXISTE%(CARPEACTUAL$ + "COT.exe") < 1 Then
       Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPEACTUAL$, "\", "/") & "COT.exe\ Imposible Generar C.O.T. ")
       GoTo 99
       
    End If
    Call FileCopy(CARPEACTUAL$ + "COT.EXE", CARPEDESTIN$ + "COT.EXE")
    If ModalidadCOT$ = "AGIP" Then
        VencimientoTokenAGIP$ = TRIM$(CONTROL("REMELEC", "EXPIAGIP"))
        If VencimientoTokenAGIP$ = "" Then VencimientoTokenAGIP$ = "01/01/1980 00:00"
        
        If InStr(EPAC$, "ENVAPLAST") > 0 Then 'modifico por que en la pc de vicki daba error 13
            FEVEN% = CVINT(VencimientoTokenAGIP$)  'aparentemente formato de fechas
            If FEVEN% < HOY(HO$) Then
                    Call COMUNI("El Ultimo Token Obtenido se Encuentra Vencido.\ \Recuerde Actualizarlo Para Poder Obtener el COT.")
                    GoTo 99
            End If
        End If
        
        FechaExpiraToken = CDate(VencimientoTokenAGIP$)
        If FechaExpiraToken < Now Then
               Call COMUNI("El Ultimo Token Obtenido se Encuentra Vencido.\ \Recuerde Actualizarlo Para Poder Obtener el COT.")
        End If
    End If
99
End Sub

Private Sub Command9_Click()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text20.TEXT = VDEVV$
        FF% = FECHANUM(Text20.TEXT)
    End If
End Sub

Private Sub GRID_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
     REG& = GRID.MouseRow
     COL% = GRID.MouseCol
     Call MENU_CELDA(GRID, REG&, COL%)
     Call CalcutotalSeleccionado 'RECALCULO POR SI ELIMINA
     Me.lblRemSel = ""
  End If
End Sub
Private Sub GRID_DblClick()
'    i& = GRID.MouseRow
'    J& = GRID.MouseCol
'    If i& = 0 Or i& > GRID.Rows Then Exit Sub
'    '
'    On Error Resume Next
'    GRID.Row = i&
'    GRID.COL = J&
'    On Error Resume Next
'
'    REG& = i&
'    If REG& < 1 Then
'        Exit Sub
'    End If
'    Call COLOREAR_GRILLA_SOLAFILA(Me.GRID, &HFFFF80, REG&)
End Sub

Private Sub LENTREGA_DblClick(Index As Integer)
    If Index = 0 Then
       CAMPO$ = "Domicilio de Entrega": CAMPOFILTRO$ = "DomiEnt"
    ElseIf Index = 1 Then
       CAMPO$ = "Localidad de Entrega": CAMPOFILTRO$ = "LocaEnt"
    ElseIf Index = 2 And XVALO(Text19) < 1 Then: CAMPOFILTRO$ = "TranspEnt"
        CAMPO$ = "Transporte"
    Else
        Exit Sub
    End If
    
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPO$ & "/A99;CONSERPARADOR/A0", FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, CAMPO$ & "/A99;CONSERPARADOR/A0", FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = ANTOT + 1100
    FRMZELECHO.Caption = CAMPO$
    FRMZELECHO.MSF1.Rows = 1
    
    SQLX$ = "SELECT DISTINCT  " & CAMPOFILTRO$ & " AS VALOR1"
    If Index < 2 Then SQLX$ = SQLX$ + " , LocaEnt  AS VALOR2"
    SQLX$ = SQLX$ + " FROM PEDENCA WITH(NOLOCK) WHERE NroClie = " & XVALO(Text2)
    SQLX$ = SQLX$ + " ORDER BY " & CAMPOFILTRO$
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        i& = i& + 1
        FRMZELECHO.MSF2.Rows = i& + 1
        FRMZELECHO.MSF2.TextMatrix(i&, 1) = SAFETEXT$(!VALOR1)
        If Index <= 1 Then
          FRMZELECHO.MSF2.TextMatrix(i&, 1) = SAFETEXT$(!VALOR1) & " " & SAFETEXT$(!VALOR2)
          If TRIM$(SAFETEXT$(!VALOR2)) <> "" Then
             FRMZELECHO.MSF2.TextMatrix(i&, 2) = SAFETEXT$(!VALOR1) & "|" & SAFETEXT$(!VALOR2)
          End If
        End If
        .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    FRMZELECHO.Show 1
    '
    LECTURAX$ = TRIM$(RESP_ZELE_VECT(2))
    If LECTURAX$ = "" Then LECTURAX$ = TRIM$(RESP_ZELE_VECT(1))
    If LECTURAX$ <> "" Then
     If Index <= 1 Then
       Call TEXTOLOTES$(LECTURAX$, "|")
       On Error Resume Next
       LECTURA1$ = CADENATEX$(1)
       LECTURA2$ = CADENATEX$(2)
       On Error GoTo 0
       If Index = 0 Then LENTREGA(0).TEXT = Left$(LECTURA1$, 48)
       LENTREGA(1).TEXT = Left$(LECTURA2$, 48)
     Else
       LENTREGA(Index).TEXT = Left$(LECTURAX$, 48)
     End If
    End If

End Sub


Private Sub mnuimprimir_Click()
     Call ReimpresionDeRemito(Text3, OK%)
     
End Sub

Private Sub mnupendicot_Click()
'   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "ID Solicitud/F12;Fecha Solicitud", FRMZELECHO)
'   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID Solicitud/F12;Fecha Solicitud", FRMZELECHO)
'   FRMZELECHO.Caption = "SOLICITUD DE COT PENDIENTES"
'
'   Dim obj As New RECORXET
'   Set RS1 = obj.RS_Solicitud_COT_Fallidas
'   Call Carga_MSF_Recordset(RS1, "ID Solicitud/F12;Fecha Solicitud", FRMZELECHO.msf2, 1)
'   Set obj = Nothing
'   RS1.Close
'   Set RS1 = Nothing
'
'   FRMZELECHO.Width = 6500
'   FRMZELECHO.Show 1
'   NRO& = XVALO(RESP_ZELE_VECT(1))
'   '
'   If NRO& < 1 Then Exit Sub
'   Call MOSTRAR_SOLICITUD_DE_COT(NRO&)

End Sub
      
Private Sub mnuGenRemDev_Click()
   GENDEVCOT.Show 1
End Sub

Private Sub mnuReimprimirCot_Click()
   Campos$ = "ID C.O.T./F10;Fecha Solicitud/D8;Cta/F8;R.Social/A24;Remito/A20;C.O.T./A12"
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = "SOLICITUD DE COT APROBADAS"

   Dim obj As New RECORXET
   Set RS1 = obj.RS_Solicitud_COT_OK("E")
   Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.MSF2, 1)
   Set obj = Nothing
   RS1.Close
   Set RS1 = Nothing


   FRMZELECHO.Width = 11000
   FRMZELECHO.Show 1
   NRO& = XVALO(RESP_ZELE_VECT(1))
   If NRO& < 1 Then Exit Sub
   '
   Condi1$ = "ID_ENCA =" & NRO&
   PATENTE$ = TRIM$(VALOBADA("DETACOT", "PATENTE", Condi1$, "NOMESS"))
   CODINTEGRIDAD$ = TRIM$(VALOBADA("DETACOT", "CODINTEGRIDAD", Condi1$, "NOMESS"))

   CONDI2$ = "ENCACOT_ID =" & NRO&
   ValorProcesado$ = TRIM$(VALOBADA("ENCACOT", "COT", CONDI2$, "NOMESS"))
   CODINTEGRIDAD$ = TRIM$(VALOBADA("ENCACOT", "CODINTEGRIDAD", CONDI2$, "NOMESS"))

   FE% = CVINT(VALOBADA("ENCACOT", "FESALIDA", CONDI2$, "NOMESS"))
   

   FILTX$ = TRIM$(ValorProcesado$) & ";" & FECHATEX$(FE%) & ";" & PATENTE$ & ";" & CODINTEGRIDAD$ & ";" & TRIM$(Str$(NRO&))
   Call STDFORM("COTARBA", FILTX$) 'MUESTRA EL FORMULARIO PARA IMPRMIR

End Sub

Private Sub mnusolaprobcot_Click()
   Campos$ = "ID C.O.T./F10;Fecha Solicitud/D8;Cta/F8;R.Social/A24;Remito/A20;C.O.T./A12"
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = "SOLICITUD DE COT APROBADAS"

   Dim obj As New RECORXET
   Set RS1 = obj.RS_Solicitud_COT_OK("E")
   Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.MSF2, 1)
   Set obj = Nothing
   RS1.Close
   Set RS1 = Nothing

   FRMZELECHO.Width = 11000
   FRMZELECHO.Show 1
   NRO& = XVALO(RESP_ZELE_VECT(1))
   Text15 = NRO&: Text14.Visible = True: Text15.Visible = True

   If NRO& < 1 Then Exit Sub
   '
   Call MOSTRAR_SOLICITUD_DE_COT(NRO&)
   

End Sub

Private Sub QuitarRemi_Click()
   NROREMI$ = TRIM$(GRID.TextMatrix(GLOBALFRM_FILA&, 8))
   If NROREMI$ = "" Then Call COMUNI("Falta Nro. de Remito"): Exit Sub
   '
   i& = 0
   Do While i& < GRID.Rows - 1
      i& = i& + 1
      If TRIM$(GRID.TextMatrix(i&, 8)) = NROREMI$ Then
       Call cmdEliminaItem(GRID, i&)
       i& = i& - 1
      End If
   Loop
   If GRID.Rows < 2 Then
      Text3 = ""
   Else
      Call numerargrilla(GRID, 1)
   End If
   
End Sub

Sub ReimpresionDeRemito(rto$, OK%)
'    If Me.GRID.Rows < 2 Then Exit Sub
'
'    PRINTERPORT$ = TRIM$(CONTROL$(IDENTER$, "PRIREEDI"))
'    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
'
'    REPLAORI% = 0
'    If CONTROL$("REMSTAN", "REPLAORI") = "SI" Then
'       REPLAORI% = 1
'    End If
'
'    OK% = 0
'
'    HAYSELECCION% = 0
'    For i& = 1 To GRID.Rows - 1
'         If TRIM$(GRID.TextMatrix(i&, 0)) = "*" Then
'            HAYSELECCION% = HAYSELECCION% + 1
'            Exit For
'         End If
'    Next i&
'    '
'    If HAYSELECCION% < 1 Then
'       Call COMUNI("Atención No se Han Seleccionado Items Para Imprimir\Puede Seleccionar Todos desde el Menú > Edición > Revertir Selección\O Realizando Doble Click Sobre los Items a Imprimir")
'       GoTo 99
'    End If
'    '
'    If HAYSELECCION% > 0 And HAYSELECCION% < GRID.Rows - 1 Then
'       OPX% = COMALTER%("Atención Solo se Imprimiran los Items Seleccionados\\Cancelar\Imprimir Items Seleccionados")
'       If OPX% < 2 Then GoTo 99
'    End If
'
'    If TRIM$(FORIPRE$) <> "" Then
'      '
'      Call LIMPIATETAB
'      '
'      FF% = CVINT(Text4)
'      NCLI = XVALO(Text2)
'      NROREMIT$ = rto$
'
'      '
'      FEX = FF%
'      FECHDOC$ = FECHATEX$(FEX)
'      NUMEDOC$ = NROREMIT$
'      TIPODOC$ = "Remito"
'      '
'      CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(LENTREGA(2))
'      CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = TRIM$(Text15.TEXT)
'      CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Text12.TEXT)
'      CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = TRIM$(Text17.TEXT)
'      CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = TRIM$(DOMITRAN(3))
'      CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Text13.TEXT)
''      CODPARAM$(7) = "DEPOSAL ": DOCPARAM$(7) = TRIM$(DEPOSAL$)
'      CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = TRIM$(rto$)
'      CODPARAM$(9) = "NRO-OCOM": DOCPARAM$(9) = TRIM$(Text20.TEXT)
'      CODPARAM$(10) = "PATE-TRA": DOCPARAM$(10) = TRIM$(Text18.TEXT)
''      CODPARAM$(11) = "N-AUTENT": DOCPARAM$(11) = TRIM$(NORVENTA)
'      CODPARAM$(12) = "COND-TRA": DOCPARAM$(12) = TRIM$(Text19)
'      CODPARAM$(13) = "LOCA-TRA": DOCPARAM$(13) = TRIM$(DOMITRAN(4))
'      CODPARAM$(14) = "NRO-AENT": DOCPARAM$(14) = TRIM$(Text22)
'
''      HORA$ = HORAENT$(NCLI)
'      CODPARAM$(15) = "Z-PLANO": DOCPARAM$(15) = ""
'      CODPARAM$(16) = "IMP-TOTA": DOCPARAM$(16) = ""
'      CODPARAM$(17) = "DOCU-APL": DOCPARAM$(17) = ""
'      CODPARAM$(18) = "DESTINO": DOCPARAM$(18) = TRIM$(TXTFLETE)
'      CODPARAM$(19) = "OBSERVA": DOCPARAM$(19) = TRIM$(Text16(0)) & " " & TRIM$(Text16(1))
'      CODPARAM$(20) = "LAR-CUTE": DOCPARAM$(20) = TRIM$(LENTREGA(0))
'      CODPARAM$(21) = "LAR-MATE": DOCPARAM$(21) = TRIM$(LENTREGA(1))
'      CODPARAM$(22) = "TELE-CLI": DOCPARAM$(22) = TEL1CLI$(NCLI)
'      CODPARAM$(23) = "COND-PAG": DOCPARAM$(23) = ECOARCH$("CONPAG", Chr$(CPAGCLI%(NCLI)))
'
'      CAPARDOC% = 23
'      '
'      CODTABU1$(1) = "COD-MAT"
'      CODTABU1$(2) = "DES-MAT"
'      CODTABU1$(3) = "CANTIDAD"
'      CODTABU1$(4) = "NRO-LOTE"
'      CODTABU1$(5) = "UNIMED"
'      CODTABU1$(6) = "N-OCOMPR"
'      CODTABU1$(7) = "ORD-ITEM"
'      CODTABU1$(8) = "REF-MAT1"
'      CODTABU1$(9) = "REF-MAT2"
'      CODTABU1$(10) = "PRE-NETO"
'      CODTABU1$(11) = "NET-ITEM"
'      CODTABU1$(12) = "REF-MAT6"
'      CODTABU1$(13) = "REF-MAT5"
'      CACOLTAB1% = 13
'      '
'      DESTIDOCNUE = NCLI
'      CAITAB1% = 0
'      '
'      JJ& = 0
'      '1 "CI/F0;
'      '2 Ord./F4;
'      '3 Descripción/A32
'      '4;Salida/F9.1;
'      '5Ingreso/f9.1;
'      '6 Nro.Ped./F7;
'      '7 P.Unit./F12.2
'      '8;N° Serie/A16
'      '9;Remito/A16"
'      YAPASO% = 0: YAPASO% = 0
'      For i& = 1 To GRID.Rows - 1
'         If TRIM$(GRID.TextMatrix(i&, 0)) <> "*" Then GoTo 50
'         REMI$ = TRIM$(GRID.TextMatrix(i&, 9))
'         If REMI_ANT$ <> "" Then
'            If REMI_ANT$ <> REMI$ And YAPASO% = 0 Then
'               CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = "Varios"
'               YAPASO% = 1
'            End If
'         End If
'         'GENERO EL STRING CON LOS NUMEROS DE PEDIDOS RELACIONADOS
'         NROPED& = XVALO(GRID.TextMatrix(i&, 6))
'            If (NROPED_ANT& <> NROPED&) And NROPED& > 0 Then
'               PEDIDOSX$ = PEDIDOSX$ & NROPED& & "-"
'            End If
'         CI1% = XVALO(GRID.TextMatrix(i&, 1))
'         JJ& = JJ& + 1
'          CAITAB1% = CAITAB1% + 1
'          CODDX$ = CODEXT$(XVALO(GRID.TextMatrix(i&, 1)))
'          If REPLAORI% = 1 Then
'             CODDX$ = CODORIG$(CI1%, NCLI)
'          End If
'          NPDX& = XVALO(GRID.TextMatrix(i&, 6))
'          CANSAL = XVALO(GRID.TextMatrix(i&, 4))
'          CANING = XVALO(GRID.TextMatrix(i&, 5))
'          CANS$ = FORMATNUM$(CANSAL - CANING, "F10." & CNTDC$ & "")
'          TETABU1$(1, CAITAB1%) = CODDX$
'          TETABU1$(2, CAITAB1%) = GRID.TextMatrix(i&, 3)
'          TETABU1$(3, CAITAB1%) = CANS$
'          NSERIE$ = TRIM$(GRID.TextMatrix(i&, 8))
'          If NSERIE$ <> "" And NSERIE$ <> "0" Then NSERIE$ = "(N°.Serie: " & NSERIE$ & ")"
'          TETABU1$(4, CAITAB1%) = NSERIE$
'          TETABU1$(5, CAITAB1%) = UNIMED$(CI1%)
''          TETABU1$(6, CAITAB1%) = NUOCO$
'          TETABU1$(7, CAITAB1%) = Str$(CAITAB1%)
'          TETABU1$(8, CAITAB1%) = NUOCOCLI$(NPDX&, CI1%)
'          TETABU1$(9, CAITAB1%) = NUPLANO$(CI1%)
''          TETABU1$(10, CAITAB1%) = FORMATNUM$(IMPOUNI#, "F12.2")
''          TETABU1$(11, CAITAB1%) = FORMATNUM$(IMPORTEX#, "F12.2")
''          CORIG$ = TRIM$(CODORIG$(CI%, NCLI))
''          If CORIG$ = CODEXT$(CI1%) Then CORIG$ = ""
''          TETABU1$(12, CAITAB1%) = CORIG$
''          'FIN ALEJPA
'
'          TETABU1$(13, CAITAB1%) = XVALO(GRID.TextMatrix(i&, 2)) ' NUMERO DE ORDEN INGRESADO EN VENTANA DE ARCON NEUQUEN.
'          'IMPRESION DE ESPECIFICACION
'          If CONTROL("FACTUREM", "INCLUSPE") = "SI" Then ' IMPRIME ESPECIFICACIONES DEL ARTICULO
'            XX$ = ESPECIFIT$(CI1%)
'            Call PRINTEXT(XX$, 2, "DES-MAT", FORIPRE$)
'          End If
'          Call CARLINSEP         'Carga linea de separacion entre codigo y codigo
'          REMI_ANT$ = REMI$
'          NROPED_ANT& = NROPED&
'50    Next i&
'      CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = TRIM$(Text17.TEXT)
'      If InStrRev(PEDIDOSX$, "-") > 0 Then PEDIDOSX$ = Mid$(PEDIDOSX$, 1, Len(PEDIDOSX$) - 1)
'      CODPARAM$(14) = "NRO-AENT": DOCPARAM$(14) = PEDIDOSX$
'      Call PRINTDOC("FOREMIFA")   ' Remito de Facturacion
'      OK% = 1
'      '
'      If TRIM$(PRINTERPORT$) <> "" Then
'         Call SETPRINPORT("RESTORE")
'      End If
'      '
'    End If
'
'99
End Sub


Private Sub mnuRevertirSeleccion_Click()
'   mnuRevertirSeleccion.Enabled = False
'   Screen.MousePointer = 11
'   FIN& = GRID.Rows - 1
'   GRID.Redraw = False
'   For i& = 1 To FIN&
'       Call TRACE(20, i&, FIN&)
'       Call COLOREAR_GRILLA_SOLAFILA(Me.GRID, &HFFFF80, i&)
'   Next i&
'   Call CalcutotalSeleccionado
'   GRID.Redraw = True
'   Screen.MousePointer = 1
'   mnuRevertirSeleccion.Enabled = True
End Sub


Private Sub Option1_Click()
   If Option1.Value = True Then
     TXTFLETE = "Flete A Cargo del Cliente"
   Else
    TXTFLETE = ""
   End If
End Sub

Private Sub Option2_Click()
  If Option2.Value = True Then
     EMPRESAX$ = EMPREAC$
     If Len(EMPREAC$) > 6 Then
       POS1% = InStr(EMPREAC$, " ")
       EMPRESAX$ = Mid$(EMPREAC$, 1, POS1%)
     End If
     TXTFLETE = "Flete A Cargo de " & EMPRESAX$
   Else
     TXTFLETE = ""
   End If

End Sub


Private Sub Text15_Change()
   Text12 = ""
   Text18 = ""
   If XVALO(Text15) < 1 Then Exit Sub
   
   CONDI$ = "ID_ENCA =" & XVALO(Text15)
   SQLX$ = "SELECT CUITTRANSP, PATENTE,DNITRANS FROM DETACOT WITH(NOLOCK) WHERE " & CONDI$
   Set rs = READSET(SQLX$)
   With rs
         Text12 = SAFETEXT$(!CUITTRANSP)
         Text18 = SAFETEXT$(!PATENTE)
'         Text13 = SAFETEXT$(!DNITRANS)
   End With
   rs.Close
   Set rs = Nothing
End Sub

Private Sub Text16_DblClick()
   Text16 = ""
End Sub

Private Sub Text3_Change()
    '
    If Command1.Enabled = True Then
'      Label18 = ""
    End If
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    '
    If XVALO(Text15) > 0 Then GRID.Rows = 1
    
    Text15 = "": Text14.Visible = False: Text15.Visible = False
    NUREMIX$ = "RT-X-" & Text3
    CONDI$ = "DoPriLar = '" & NUREMIX$ & "'"
    IDDETCOT& = XVALO(VALOBADA("DETACOT", "ID_ENCA", CONDI$, "NOMESS")) 'VEO SI YA FUE SOLICITADO
    If IDDETCOT& > 0 Then
      CONDI2$ = "ENCACOT_ID =" & IDDETCOT& 'VEO SI OBTUVO COT
      COTX$ = TRIM$(VALOBADA("ENCACOT", "COT", CONDI2$, "NOMESS"))
      If COTX$ <> "" Then
         Call COMUNI("Atención !!!! \Remito: " & NUREMIX$ & " Ya Fué Procesado la Solicitud y Obtuvo C.O.T.")
      Else
         Call COMUNI("Atención !!!! \Remito: " & NUREMIX$ & " Ya Fué Procesado la Solicitud Está Pendiente de Obtener C.O.T.\Identificador de Solicitud de C.O.T.: " & IDDETCOT&)
      End If
      Text3 = ""
      Exit Sub
    End If
    '
    Me.lblRemSel = Text3
    
    Call LIMPIA_CAJA_TEXTO
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If CMO$ = "" Then CMO$ = "RT"
    '
    VUELTA% = 0: NPED& = 0
    Call COLOREAR_GRILLASQL_TODAS_MISMO_COLOR(GRID, vbWhite)
    If Me.Text3 <> "" Then
        CONDI$ = "(CodiMovi = '" & CMO$ & "' OR CodiMovi = 'DV') "
        CONDI$ = CONDI$ + " AND DoPriLAR LIKE '%" & Text3 & "%' "
        VUELTA% = 0: J& = 0
        SQLX$ = "SELECT COD_INTE,NuorItem,DESCRI_LAR, CantSalid, cantingre, NUPEDCLI, FECHMOV,"
        SQLX$ = SQLX$ + " Nume_Clie, FECHMOV,IDENLOTE,ValoUnit,MoneVal,TIPOCAM FROM MOVISTO WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE " & CONDI$ & "  ORDER BY NuOrItem"
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
             J& = GRID.Rows
             GRID.Rows = J& + 1

             If VUELTA% = 0 Then
               VUELTA% = 1
               NCLIE = XVALO(!Nume_Clie)
               Text2 = NCLIE
               Text6 = rasocli$(XVALO(Text2))
               Text4 = FECHATEX$(CVINT(!FECHMOV))
             End If
             If NCLIE_ANT <> NCLIE Then
                DOCUMENTOX$ = ""
                If PIVACLI%(NCLIE) < 1 Then
                    DOCUMENTOX$ = CuitCli(NCLIE)
                End If
             End If

             'UN REMITO PUEDE TENER VARIOS PEDIDOS SI ALGUNO TIENE PEDIDO LE PASO EL NRO PARA QUE MUESTRE LOS D.ENTR DEL PEDIDO.
             If XVALO(!NUPEDCLI) > 0 Then NPED& = XVALO(!NUPEDCLI)
             GRID.TextMatrix(J&, 1) = XVALO(!COD_INTE)
             GRID.TextMatrix(J&, 2) = J&
             GRID.TextMatrix(J&, 3) = SAFETEXT$(!DESCRI_LAR)
             GRID.TextMatrix(J&, 4) = TRIM$(FORMATNUM$(XVALO(!CANTSALID), "F9.1"))
             VALORUNITARIO# = XVALO(!VALOUNIT)
             'GRID.TextMatrix(j&, 5) = TRIM$(FORMATNUM$(VALORUNITARIO#, "F12.2"))
             If XVALO(!MoneVal) > 1 And InStr(1, EPAC$, "ENVAPLAST") > 0 Then   'SI ES ENVAPLAST Y MONEDA > 1
                VALORUNITARIO# = ROUND(VALORUNITARIO# * XVALO(!TIPOCAM), 2) 'LO PASA PESOS
                GRID.TextMatrix(J&, 5) = TRIM$(FORMATNUM$(VALORUNITARIO#, "F12.2"))
                Call COLOREAR_FUENTE_SOLO_UNA_CELDA(Me.GRID, vbRed, J&, 5)
             Else
                GRID.TextMatrix(J&, 5) = TRIM$(FORMATNUM$(VALORUNITARIO#, "F12.2"))
             End If
             'IMPORTEX = (XVALO(!CANTSALID)) * XVALO(!VALOUNITIVA)
             IMPORTEX = (XVALO(!CANTSALID)) * XVALO(VALORUNITARIO#) 'SE CAMBIO A VALOUNIT PARA CONSIDERAR SOLO EL NETO, EN A Y B GUARDA SIEMPRE EL NETO
             ACUMULATOTAL = ACUMULATOTAL + IMPORTEX
             GRID.TextMatrix(J&, 6) = XVALO(!NUPEDCLI)
             Call COLOREAR_GRILLA_SOLO_UNA_AFILA(GRID, &HE0E0E0, J&)
             GRID.TextMatrix(J&, 7) = SAFETEXT$(!idenlote)
             GRID.TextMatrix(J&, 8) = Text3
             GRID.TextMatrix(J&, 9) = FECHATEX$(CVINT(!FECHMOV))
             GRID.TextMatrix(J&, 19) = DOCUMENTOX$
             NCLIE_ANT = NCLIE

            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
        'OBTENGO LOS DATOS DE ENTREGA DEL PEDIDO SI VA SIN NRO DE PEDIDO O EL PEDIDO NO TIENE DATOS LE ASIGNA LOS DATOS
        'DE ENTREGA DE LA SOLAPA ENTREGA .
        Call CARGAR_DATOS_ENTREGA(XVALO(Text2), NPED&, Text3)
        Call CalcutotalSeleccionado
        Label18 = GRID.TextMatrix(GRID.Rows - 1, 2)
        '

        '
'        'MOSTRAR DATOS DE ENTREGA
'        For i& = 1 To GRID.Rows - 1
'           NPDX& = XVALO(GRID.TextMatrix(i&, 6))
'           If NPDX& > 0 Then
'             SQLX$ = "SELECT DOMIENT,LOCAENT,TranspEnt,NROOCOM,FLETE_TEX, NTranspo FROM PEDENCA WITH(NOLOCK)"
'             SQLX$ = SQLX$ + " WHERE NROPED = " & NPDX&
'             Set RS = READSET(SQLX$)
'             With RS
'             Do While Not .EOF
'               If TRIM$(SAFETEXT$(!DOMIENT)) <> "" Then
'                 LENTREGA(0) = SAFETEXT(!DOMIENT)
'                 LENTREGA(1) = SAFETEXT(!LOCAENT)
'                 NTRAN% = XVALO(!NTranspo)
'                 If NTRAN% > 0 Then ' SI VIENE CON CODIGO DE TRANSPORTE COMPLETA AUTOMATICAMENTE
'                   Text45 = NTRAN%
'                 Else
'                   LENTREGA(2) = SAFETEXT(!TranspEnt)
'                 End If
'                 Text20 = SAFETEXT(!NROOCOM)
'                 TXTFLETE = SAFETEXT(!FLETE_TEX)
'                 GoTo 20
'               End If
'               .MoveNext
'             Loop
'             End With
'20           RS.Close
'             Set RS = Nothing
'           End If
'        Next i&
    End If

End Sub

Sub CalcutotalSeleccionado()
        Label3 = "": ACUMULATOTAL = 0
        For i& = 1 To GRID.Rows - 1
          NROREMI$ = GRID.TextMatrix(i&, 8)
          If NROREMI$ <> NROREMI_ANT$ Then CANTREM = CANTREM + 1: Label2.Caption = FORMATNUM$(CANTREM, "F8.1")
          IMPORTEX = (XVALO(GRID.TextMatrix(i&, 4)) * XVALO(GRID.TextMatrix(i&, 5)))
          ACUMULATOTAL = ACUMULATOTAL + IMPORTEX
          NROREMI_ANT$ = NROREMI$
        Next i&
        Label3.Caption = FORMATNUM$(ACUMULATOTAL, "F12.2")
End Sub

Function CalcutotalRemito$(Nremito$)
        For i& = 1 To GRID.Rows - 1
          NROREMIGRILLA$ = GRID.TextMatrix(i&, 8)
          If Nremito$ = NROREMIGRILLA$ Then
             IMPORTEX = (XVALO(GRID.TextMatrix(i&, 4)) * XVALO(GRID.TextMatrix(i&, 5)))
             ACUMULATOTAL = ACUMULATOTAL + IMPORTEX
          End If
        Next i&
        CalcutotalRemito$ = FORMATNUM$(ACUMULATOTAL, "F14.2")
End Function

Sub LIMPIARTEXT(FORMULARIO As Form, LIMPIOTEXT%)
  Dim CONTROLX As CONTROL
  For Each CONTROLX In FORMULARIO.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
        If TypeOf CONTROLX Is TextBox And LIMPIOTEXT% > 0 Then
           CONTROLX.TEXT = ""
        End If
  Next
End Sub

Private Sub Text45_Change()
'  Me.LENTREGA(2) = ""
'  Text12 = ""
'  Text18 = ""
'  Me.DOMITRAN(3) = ""
'  Me.DOMITRAN(4) = ""
  '
  NTRA = XVALO(Text45)
  If NTRA > 0 Then
    If NTRA < 32767 Then
      NTRAI% = NTRA
      CONDI$ = "CODI_TRA = " & NTRAI%
       If CantRegistros("TRANSPORTES", CONDI$) > 0 Then
'        Text12 = CUITTRANS$(NTRAI%) 'LO COMENTE POR QUE SI PRIMERO INGRESAS EL CUIT Y LUEGO EL TRANSPORTISTA TE BORRA LO QUE ESCRIBISTE.
        LENTREGA(2) = DENOTRANS$(NTRAI%)
        DOMITRAN(3) = DOMITRANS$(NTRAI%)
        DOMITRAN(4) = LOCATRANS$(NTRAI%)
      End If
    End If
  End If

End Sub

Sub LIMPIA_CAJA_TEXTO()
    Text2 = ""
'    Text3 = ""
    Text4 = ""
    Text6 = ""
    Text8 = ""
    Text21 = ""
    Text23 = ""
    Me.LENTREGA(0) = ""
    Me.LENTREGA(1) = ""
    Me.LENTREGA(2) = ""
    Me.DOMITRAN(3) = ""
    Me.DOMITRAN(4) = ""
    Me.Text45 = ""
    Label18.Caption = ""
    TXTCPOSCLI = ""
    Label2.Caption = ""
    Label3.Caption = ""
    
End Sub

Private Sub Text7_DblClick()
  If Text7.TEXT = "E" Then
    Text7.TEXT = "D"
    Text7.BackColor = vbBlack
    Text7.ForeColor = vbRed
    Text11.BackColor = vbBlack
    Text11.ForeColor = vbRed
    Text11.TEXT = "Remito de Devolución"
  Else
    Text7.TEXT = "E"
    Text7.BackColor = &HE0E0E0
    Text7.ForeColor = &H80000008
    Text11.BackColor = &HE0E0E0
    Text11.ForeColor = &H80000008
    Text11.TEXT = "Remito de Entrega"
  End If
End Sub


Private Sub Text8_Change()
    If TRIM$(Text8.TEXT) = "" Then TXTPCIA = ""
    DPX1$ = UCase$(TRIM$(Text8.TEXT))
    If DPX1$ <> "" Then
        Text8.TEXT = DPX1$
        TXTPCIA = ECOARCH$("PROVINC", DPX1$)
    End If
End Sub


Private Sub Timer1_Timer()
   Static CONTA%
   CONTA% = CONTA% + 1 ' SI SUPERA LOS 40 SEGUNDOS DE DEMORA EN TRAER EL CAE PRESENTA MENSAJE
   If CONTA% > 40 Then
     Call COMUNI("No Se Ha Podido Obtener el C.O.T.")
     CONTA% = 0
   End If
   '
   COT$ = LEECOT$    '  (IDENTCAE$)
  
End Sub

Function LEECOT$() ' (IDENTCAE$)
    '
    CARPETOK$ = "C:\FLEXOFT\COT\"
    LEECOT$ = ""
    COT$ = ""
    TX1$ = ""
    J& = 0: L& = 0
    Dim DESCRIPCIONERROR$(), PROCESADOOK%(), CODIGOERR%(), ValorProcesado$()
    ReDim Preserve DESCRIPCIONERROR$(1)
    ReDim Preserve PROCESADOOK%(1)
    ReDim Preserve CODIGOERR%(1)
    ReDim Preserve ValorProcesado$(1)

'    PROCESADOOK% = 0
    '
'    X = COPYFILE%(CARPETOK$ + "COT_" + "123456.xml", CARPETOK$ + "COT_" + "123456.txt")
'** RECORRO EL ARCHIVO SI YA SE GENERO ****
'    IDCOT$ = IDCOT$
    If EXISTE%(CARPETOK$ + NAMEARCHIVO$ + ".TXT") > 0 Then
        NARCHIVO% = FILEOPEN%(CARPETOK$ + NAMEARCHIVO$ + ".TXT", 1, 256)  'ABRE EL ARCHIVO DE LECTURA
        Do While Not EOF(NARCHIVO%) ' RECORRE
          Line Input #NARCHIVO%, XXXXX$ ' LEE CADA REGISTRO
          TX1$ = TX1$ + XXXXX$ + vbCrLf
          If InStr(UCase$(XXXXX$), UCase$("<numeroComprobante>")) > 0 Then '92683690
             POS1% = InStr(UCase$(XXXXX$), UCase$("<numeroComprobante>")) + 19
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</numeroComprobante>")
             COT$ = Mid$(TXT$, 1, POS2% - 1)
          End If
          If InStr(UCase$(XXXXX$), UCase$("<codigoIntegridad>")) > 0 Then '31b9e82ef398b85d755b72d074f41bc7
             POS1% = InStr(UCase$(XXXXX$), UCase$("<codigoIntegridad>")) + 19
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</codigoIntegridad>")
             CODINTEGRIDAD$ = Mid$(TXT$, 1, POS2% - 1)
          End If
          If InStr(UCase$(XXXXX$), UCase$("<numeroUnico>")) > 0 Then '91 R000100030875
             POS1% = InStr(UCase$(XXXXX$), UCase$("<numeroUnico>")) + 13
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</numeroUnico>")
             NUMEROUNICO$ = Mid$(TXT$, 1, POS2% - 1)
             
             J& = J& + 1
'             ReDim Preserve DESCRIPCIONERROR$(j&)
             ReDim Preserve PROCESADOOK%(J&)
             ReDim Preserve CODIGOERR%(J&)
             ReDim Preserve ValorProcesado$(J&)
             
          End If
          If InStr(UCase$(XXXXX$), UCase$("<procesado>")) > 0 Then 'NO o SI
             POS1% = InStr(UCase$(XXXXX$), UCase$("<procesado>")) + 11
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</procesado>")
             ValorProcesado$(J&) = Mid$(TXT$, 1, POS2% - 1)
             If TRIM$(UCase$(ValorProcesado$(J&))) = "SI" Then PROCESADOOK%(J&) = 1
          End If
          
          If InStr(UCase$(XXXXX$), UCase$("<codigo>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<codigo>")) + 8
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</codigo>")
             CODIGOERR%(J&) = XVALO(Mid$(TXT$, 1, POS2% - 1))
          End If
          '
          'EN CASO QUE VENGA con PROCESADO NO PERO SEA CODIGO DIFERENTE A 17 BORRA EL REGISTRO
          'SI ES 17 SE GRABA IGUAL COMO SI OBTUVO COT , POR QUE EN REALIDAD OBTUVO Y NO QUEDO INFORMADO EN EL SISTEMA.
          If TRIM$(UCase$(ValorProcesado$(J&))) = "NO" And CODIGOERR%(J&) > 0 And CODIGOERR%(J&) <> 17 Then
             'BORRO EL/LOS REGISTRO DETACOT QUE COINCIDAN CON EL NUMERO QUE REBOTARON
             SQLX$ = "DELETE FROM DETACOT WHERE CODUNICO = '" & NUMEROUNICO$ & "' AND CantSalid > 0.005 "
             FLEXCONN.Execute (SQLX$)
             CODIGOERR%(J&) = 0
          ElseIf TRIM$(UCase$(ValorProcesado$(J&))) = "NO" And CODIGOERR%(J&) = 17 Then
             ValorProcesado$(J&) = "SI"
             CODIGOERR%(J&) = 0
             PROCESADOOK%(J&) = 1
          End If
          '
          If InStr(UCase$(XXXXX$), UCase$("<descripcion>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<descripcion>")) + 13
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</descripcion>")
             L& = L& + 1
             ReDim Preserve DESCRIPCIONERROR$(L&)
             DESCRIPCIONERROR$(L&) = Mid$(TXT$, 1, POS2% - 1) & " (" & NUMEROUNICO$ & ")"
          End If
          
        Loop
        Close NARCHIVO%
        
        '** DESHABILITO EL TIMER ****
        '** DE ACUERDO AL RESULTADO IMPRIMO SI FUE TODO BIEN ****
        'If TRIM$(UCase$(ValorProcesado$)) = "SI" Then
         On Error Resume Next
          '** MUESTRO UNA LISTA DE ERORRES DEVUELTOS. ****
         If UBound(DESCRIPCIONERROR$) > 0 Then
             Campos$ = "Descripciónes de Errores Devueltos x ARBA/A84"
             Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , 12000)
             Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO, , , 12000)
             FRMZELECHO.Caption = "Detalle de Errores Devueltos"
             FRMZELECHO.Width = 12000
             For J& = 1 To UBound(DESCRIPCIONERROR$)
                 If TRIM$(DESCRIPCIONERROR$(J&)) <> "" Then
                   FRMZELECHO.MSF2.Rows = J& + 1
                   FRMZELECHO.MSF2.TextMatrix(J&, 1) = DESCRIPCIONERROR$(J&)
                 End If
             Next J&
             If FRMZELECHO.MSF2.Rows > 1 Then
                FRMZELECHO.Show 1
             End If
'             Timer1.Enabled = False
         End If
         On Error GoTo 0
         For J& = 1 To UBound(PROCESADOOK%)
            If PROCESADOOK%(J&) > 0 Then
              FE% = FECHANUM(Text20) ' fecha de pantalla
              CANTIDIAS% = COMBO1.ListIndex + 1
              If CANTIDIAS% < 0 Then CANTIDIAS% = 1
              FE% = DIASPOST(FE%, CANTIDIAS%)
              CONDI$ = "ENCACOT_ID = " & ID_COT&
              Call ACTUREGISTRO("ENCACOT", "COT", CONDI$, COT$, REGAF&, "NOMESS")
              Call ACTUREGISTRO("ENCACOT", "CODINTEGRIDAD", CONDI$, CODINTEGRIDAD$, REGAF&, "NOMESS")
              FILTX$ = TRIM$(COT$) & ";" & FECHATEX$(FE%) & ";" & TRIM$(Text18) & ";" & CODINTEGRIDAD$ & ";" & TRIM$(Str$(ID_COT&))
              Call STDFORM("COTARBA", FILTX$) 'MUESTRA EL FORMULARIO PARA IMPRMIR
              Call LIMPIA_CAJA_TEXTO: Text3 = "": Call limpiarDatoConductor
              GRID.Rows = 1
              ID_COT& = 0 'SI AL MENOS UNO OBTUVO COT ASIGNO ID_COT& = 0 P Q ES LA BANDERA PARA QUE GRABE
              Exit For
            End If
         Next J&
'         If PROCESADOOK%(J&) > 0 Then ID_COT& = 0 'SI AL MENOS UNO OBTUVO COT ASIGNO ID_COT& = 0 P Q ES LA BANDERA PARA QUE GRABE
         
         Timer1.Enabled = False

    End If
    '


End Function
Function LEECOT_ant$() ' (IDENTCAE$)
    '
'''    CARPETOK$ = "C:\FLEXOFT\COT\"
'''    LEECOT$ = ""
'''    COT$ = ""
'''    TX1$ = ""
'''    J& = 0
'''    Dim DESCRIPCIONERROR$()
'''    PROCESADOOK% = 0
'''    '
''''    X = COPYFILE%(CARPETOK$ + "COT_" + "123456.xml", CARPETOK$ + "COT_" + "123456.txt")
''''** RECORRO EL ARCHIVO SI YA SE GENERO ****
''''    IDCOT$ = IDCOT$
'''    If EXISTE%(CARPETOK$ + NAMEARCHIVO$ + ".TXT") > 0 Then
'''        NARCHIVO% = FILEOPEN%(CARPETOK$ + NAMEARCHIVO$ + ".TXT", 1, 256)  'ABRE EL ARCHIVO DE LECTURA
'''        Do While Not EOF(NARCHIVO%) ' RECORRE
'''          Line Input #NARCHIVO%, XXXXX$ ' LEE CADA REGISTRO
'''          TX1$ = TX1$ + XXXXX$ + vbCrLf
'''          If InStr(UCase$(XXXXX$), UCase$("<numeroComprobante>")) > 0 Then '92683690
'''             POS1% = InStr(UCase$(XXXXX$), UCase$("<numeroComprobante>")) + 19
'''             TXT$ = Mid$(XXXXX$, POS1%)
'''             POS2% = InStr(TXT$, "</numeroComprobante>")
'''             COT$ = Mid$(TXT$, 1, POS2% - 1)
'''          End If
'''          If InStr(UCase$(XXXXX$), UCase$("<codigoIntegridad>")) > 0 Then '31b9e82ef398b85d755b72d074f41bc7
'''             POS1% = InStr(UCase$(XXXXX$), UCase$("<codigoIntegridad>")) + 19
'''             TXT$ = Mid$(XXXXX$, POS1%)
'''             POS2% = InStr(TXT$, "</codigoIntegridad>")
'''             CODINTEGRIDAD$ = Mid$(TXT$, 1, POS2% - 1)
'''          End If
'''          If InStr(UCase$(XXXXX$), UCase$("<numeroUnico>")) > 0 Then '91 R000100030875
'''             POS1% = InStr(UCase$(XXXXX$), UCase$("<numeroUnico>")) + 13
'''             TXT$ = Mid$(XXXXX$, POS1%)
'''             POS2% = InStr(TXT$, "</numeroUnico>")
'''             NUMEROUNICO$ = Mid$(TXT$, 1, POS2% - 1)
'''          End If
'''          If InStr(UCase$(XXXXX$), UCase$("<procesado>")) > 0 Then 'NO o SI
'''             POS1% = InStr(UCase$(XXXXX$), UCase$("<procesado>")) + 11
'''             TXT$ = Mid$(XXXXX$, POS1%)
'''             POS2% = InStr(TXT$, "</procesado>")
'''             ValorProcesado$ = Mid$(TXT$, 1, POS2% - 1)
'''             If TRIM$(UCase$(ValorProcesado$)) = "SI" Then PROCESADOOK% = 1
'''          End If
'''
'''          If InStr(UCase$(XXXXX$), UCase$("<codigo>")) > 0 Then
'''             POS1% = InStr(UCase$(XXXXX$), UCase$("<codigo>")) + 8
'''             TXT$ = Mid$(XXXXX$, POS1%)
'''             POS2% = InStr(TXT$, "</codigo>")
'''             CODIGOERR% = XVALO(Mid$(TXT$, 1, POS2% - 1))
'''          End If
'''          '
'''          'EN CASO QUE VENGA con PROCESADO NO PERO SEA CODIGO DIFERENTE A 17 BORRA EL REGISTRO
'''          'SI ES 17 SE GRABA IGUAL COMO SI OBTUVO COT , POR QUE EN REALIDAD OBTUVO Y NO QUEDO INFORMADO EN EL SISTEMA.
'''          If TRIM$(UCase$(ValorProcesado$)) = "NO" And CODIGOERR% > 0 And CODIGOERR% <> 17 Then
'''             'BORRO EL/LOS REGISTRO DETACOT QUE COINCIDAN CON EL NUMERO QUE REBOTARON
'''             SQLX$ = "DELETE FROM DETACOT WHERE CODUNICO = '" & NUMEROUNICO$ & "' AND CantSalid > 0.005 "
'''             FLEXCONN.Execute (SQLX$)
'''             CODIGOERR% = 0
'''             ValorProcesado$ = ""
'''          ElseIf TRIM$(UCase$(ValorProcesado$)) = "NO" And CODIGOERR% = 17 Then
'''             ValorProcesado$ = ""
'''             CODIGOERR% = 0
'''             PROCESADOOK% = 1
'''          End If
'''          '
'''          If InStr(UCase$(XXXXX$), UCase$("<descripcion>")) > 0 Then
'''             POS1% = InStr(UCase$(XXXXX$), UCase$("<descripcion>")) + 13
'''             TXT$ = Mid$(XXXXX$, POS1%)
'''             POS2% = InStr(TXT$, "</descripcion>")
'''             J& = J& + 1
'''             ReDim Preserve DESCRIPCIONERROR$(J&)
'''             DESCRIPCIONERROR$(J&) = Mid$(TXT$, 1, POS2% - 1) & " (" & NUMEROUNICO$ & ")"
'''          End If
'''
'''        Loop
'''        Close NARCHIVO%
'''
'''        '** DESHABILITO EL TIMER ****
'''        '** DE ACUERDO AL RESULTADO IMPRIMO SI FUE TODO BIEN ****
'''        'If TRIM$(UCase$(ValorProcesado$)) = "SI" Then
'''         On Error Resume Next
'''          '** MUESTRO UNA LISTA DE ERORRES DEVUELTOS. ****
'''         If UBound(DESCRIPCIONERROR$) > 0 Then
'''             CAMPOS$ = "Descripciónes de Errores Devueltos x ARBA/A84"
'''             Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO, , , 12000)
'''             Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO, , , 12000)
'''             FRMZELECHO.Caption = "Detalle de Errores Devueltos"
'''             FRMZELECHO.Width = 12000
'''             For J& = 1 To UBound(DESCRIPCIONERROR$)
'''                 FRMZELECHO.msf2.Rows = J& + 1
'''                 FRMZELECHO.msf2.TextMatrix(J&, 1) = DESCRIPCIONERROR$(J&)
'''             Next J&
'''             FRMZELECHO.Show 1
''''             Timer1.Enabled = False
'''         End If
'''         On Error GoTo 0
'''
'''        If PROCESADOOK% > 0 Then
'''           FE% = FECHANUM(Text20) ' fecha de pantalla
'''           CANTIDIAS% = COMBO1.ListIndex + 1
'''           If CANTIDIAS% < 0 Then CANTIDIAS% = 1
'''           FE% = DIASPOST(FE%, CANTIDIAS%)
'''           CONDI$ = "ENCACOT_ID = " & ID_COT&
'''           Call ACTUREGISTRO("ENCACOT", "COT", CONDI$, COT$, REGAF&, "NOMESS")
'''           Call ACTUREGISTRO("ENCACOT", "CODINTEGRIDAD", CONDI$, CODINTEGRIDAD$, REGAF&, "NOMESS")
'''           FILTX$ = TRIM$(COT$) & ";" & FECHATEX$(FE%) & ";" & TRIM$(Text18) & ";" & CODINTEGRIDAD$ & ";" & TRIM$(Str$(ID_COT&))
'''           Call STDFORM("COTARBA", FILTX$) 'MUESTRA EL FORMULARIO PARA IMPRMIR
'''           Call LIMPIA_CAJA_TEXTO: Text3 = "": Call limpiarDatoConductor
'''           GRID.Rows = 1
'''           ID_COT& = 0
''''           Timer1.Enabled = False
'''         End If
''''        Else
'''         If PROCESADOOK% > 0 Then ID_COT& = 0 'SI AL MENOS UNO OBTUVO COT ASIGNO ID_COT& = 0 P Q ES LA BANDERA PARA QUE GRABE
'''         Timer1.Enabled = False
'''
'''    End If
    '


End Function


