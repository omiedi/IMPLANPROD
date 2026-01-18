VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form GENDEVCOT 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   9270
   ClientLeft      =   1845
   ClientTop       =   3525
   ClientWidth     =   13590
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleMode       =   0  'User
   ScaleWidth      =   12820.76
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
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
      Left            =   5400
      TabIndex        =   48
      Top             =   3600
      Width           =   615
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
      Height          =   330
      Index           =   1
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "Filas Pintadas de Amarillo ya Procesaron C.O.T. De Devolución"
      Top             =   3600
      Width           =   5415
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
      Height          =   330
      Index           =   0
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   "Selección de Detalle a Devolver"
      Top             =   3600
      Width           =   2895
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
      Height          =   330
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   45
      Text            =   "Detalle de Solicitud de Cot "
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PROCESO FECHA"
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
      TabIndex        =   40
      Top             =   8880
      Visible         =   0   'False
      Width           =   3480
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
      Height          =   435
      Left            =   3890
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   8800
      Width           =   4665
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00C0C0C0&
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
      Height          =   615
      Left            =   120
      TabIndex        =   32
      Top             =   8160
      Width           =   8445
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
         Height          =   330
         Left            =   5450
         TabIndex        =   35
         Top             =   225
         Width           =   2925
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
         Left            =   720
         TabIndex        =   34
         Top             =   240
         Width           =   1605
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
         Height          =   285
         Left            =   3240
         TabIndex        =   33
         Top             =   240
         Width           =   1125
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
         Left            =   120
         TabIndex        =   38
         Top             =   315
         Width           =   435
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
         Left            =   4640
         TabIndex        =   37
         Top             =   300
         Width           =   675
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
         Left            =   2400
         TabIndex        =   36
         Top             =   315
         Width           =   675
      End
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
      Left            =   3120
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "Soilicitud de C.O.T.  N°:"
      Top             =   120
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6240
      Top             =   4200
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2025
      Left            =   8640
      TabIndex        =   13
      Top             =   7200
      Width           =   4995
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0C0C0&
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
         Left            =   120
         TabIndex        =   41
         Top             =   1320
         Width           =   1920
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
            Left            =   1515
            TabIndex        =   43
            Top             =   180
            Width           =   175
         End
         Begin VB.TextBox Text20 
            Alignment       =   1  'Right Justify
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
            Left            =   240
            TabIndex        =   42
            Top             =   180
            Width           =   1230
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Kilómetros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   30
         Top             =   1320
         Width           =   2130
         Begin VB.ComboBox COMBO1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   200
            Width           =   1920
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0C0C0&
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
         Left            =   2760
         TabIndex        =   28
         Top             =   720
         Width           =   2160
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
            Left            =   210
            TabIndex        =   29
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
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
         Left            =   120
         TabIndex        =   26
         Top             =   720
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
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1485
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   16
         Top             =   120
         Width           =   1935
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
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1485
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
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
         Left            =   2760
         TabIndex        =   14
         Top             =   120
         Width           =   2130
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
            Left            =   210
            TabIndex        =   15
            Top             =   240
            Width           =   1815
         End
      End
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "Remito de Devolución"
      Top             =   120
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
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Empresa Destinataria"
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
      TabIndex        =   8
      Top             =   7200
      Width           =   8475
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
         Left            =   6000
         TabIndex        =   25
         ToolTipText     =   "Selección de Provincia"
         Top             =   600
         Visible         =   0   'False
         Width           =   175
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   600
         Width           =   440
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
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "Pcia.:"
         Top             =   600
         Width           =   570
      End
      Begin VB.TextBox TXTPCIA 
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
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox TXTCPOSCLI 
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
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   2415
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
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "C.Postal:"
         Top             =   240
         Width           =   1050
      End
      Begin VB.TextBox LENTREGA 
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox LENTREGA 
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
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   3615
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
         TabIndex        =   10
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
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "Domicilio:"
         Top             =   240
         Width           =   1050
      End
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "D"
      ToolTipText     =   "Doble Click Para Cambiar a Devolución o Entrega"
      Top             =   120
      Width           =   915
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
      Left            =   11160
      TabIndex        =   4
      ToolTipText     =   "Selección  de Remitos a inclluir en C.O.T."
      Top             =   360
      Visible         =   0   'False
      Width           =   175
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   720
      OleObjectBlob   =   "GENDEVCOT.frx":0000
      Top             =   3000
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   3165
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "Botón Derecho Opciones "
      Top             =   3900
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   5583
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
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   3045
      Left            =   0
      TabIndex        =   44
      ToolTipText     =   "Botón Derecho Opciones - Doble Click Para Bajar de a un Item"
      Top             =   480
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   5371
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
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
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
   Begin VB.Menu mnuRemitoDevolucion 
      Caption         =   "Remito de Devolución"
      Begin VB.Menu mnuremdevol 
         Caption         =   "Selección de Solicitudes de C.O.T.  Aprobadas"
      End
      Begin VB.Menu mnuVerRemDev 
         Caption         =   "Ver y Re-Imprimir Remitos de Devolución"
      End
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuración"
      Begin VB.Menu Config 
         Caption         =   "Configuraciones C.O.T."
      End
   End
End
Attribute VB_Name = "GENDEVCOT"
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


Sub ACTUALIZARFECHAREMITO_EN_DETACOT()
   
   
     Dim rs As ADODB.Recordset
     Set rs = New ADODB.Recordset
     SQLX$ = "SELECT DOPRILAR, FEREMITO FROM DETACOT  "
25   On Error Resume Next
     rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
     End If
     With rs
      Do While Not .EOF
        REMI$ = SAFETEXT$(!doprilar)
        FEREMI1% = CVINT(VALOBADA("MOVISTO", "FECHMOV", "DoPriLar = '" & REMI$ & "'"))
        If FEREMI1% > 0 Then
           !FEREMITO = CVDAT(FEREMI1%)
        .Update
        End If
        .MoveNext
      Loop
     End With
     rs.Close
     Set rs = Nothing
     Call COMUNI("PROCESO TERMINADO")
End Sub

Sub CARGAR_DATOS_ENTREGA(NCLIE, NROPED&, REMI$)
    If NROPED& > 0 Then
       SQLX$ = "SELECT DOMIENT,LOCAENT,TranspEnt,NROOCOM,FLETE_TEX, NTranspo FROM PEDENCA WITH(NOLOCK)"
       SQLX$ = SQLX$ + " WHERE NROPED = " & NROPED&
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

'    For i& = 1 To GRID.Rows - 1
'       If REMI$ = GRID.TextMatrix(i&, 8) Then
'
'          GRID.TextMatrix(i&, 10) = LENTREGA(0)
'          GRID.TextMatrix(i&, 11) = LENTREGA(1)
''          GRID.TextMatrix(i&, 12) = LENTREGA(2)
'          GRID.TextMatrix(i&, 12) = TXTCPOSCLI.TEXT
'          GRID.TextMatrix(i&, 13) = "" 'Text12 'CUIT TRANSPORTE
'          GRID.TextMatrix(i&, 14) = Text8
'          GRID.TextMatrix(i&, 15) = DOMITRAN(3) 'DOMICILIO TRANSPORTE
'          GRID.TextMatrix(i&, 16) = DOMITRAN(4)  'LOCALIDAD TRANSPORTE
'          GRID.TextMatrix(i&, 17) = "" ' Text18 'PATENTE
'          GRID.TextMatrix(i&, 18) = Text2 'CLIENTE
'          'GRID.TextMatrix(i&, 19) = Text13 'NUMERO DNI CLIENTE SI ES POSICI
'       End If
'    Next i&
End Sub
Sub AgregarDatosDeGrillaaText(REMI$)
'    For i& = 1 To GRID.Rows - 1
'       If REMI$ = GRID.TextMatrix(i&, 8) Then
'          LENTREGA(0) = GRID.TextMatrix(i&, 10)
'          LENTREGA(1) = GRID.TextMatrix(i&, 11)
''          LENTREGA(2) = GRID.TextMatrix(i&, 12)
'          TXTCPOSCLI.TEXT = GRID.TextMatrix(i&, 12)
'          'Text12 = GRID.TextMatrix(i&, 13)  'CUIT TRANSPORTE
'          Text8 = GRID.TextMatrix(i&, 14)
'          'Text18 = GRID.TextMatrix(i&, 17) 'PATENTE
''          Text2 = GRID.TextMatrix(i&, 18) 'CLIENTE
'          'Text13 = GRID.TextMatrix(i&, 19)
'          Text6 = rasocli$(XVALO(Text2))
'          'Text13 = GRID.TextMatrix(i&, 19)
'          Exit For
'       End If
'    Next i&
'    Label18 = GRID.TextMatrix(GRID.Rows - 1, 2)

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

Sub Generar_Archivo_Cot_Dev(IDENCA&, ret_DATOSCOT$, ret_NAMEARCH$)
    Dim oDE As New DatoEmpresa 'instancio cla
    CUIT1$ = TRIM$(oDE.Cuit)
    'REEEMPLAZO GUIONES POR VACIO Y POR LAS DUDAS PUNTO O GUION BAJO
    CUIT1$ = REPLACAR$(CUIT1$, "-", ""): CUIT1$ = REPLACAR$(CUIT1$, ".", ""): CUIT1$ = REPLACAR$(CUIT1$, "_", "")
'   '
    NREMI1$ = "": NREMI_ANT$ = ""
    TX1$ = ""

    FECHA1$ = FECHAFORMATOINVERSO$(HOY(HO$))
    '
    FECHAENT$ = FECHAFORMATOINVERSO$(FECHANUM(Text20))
    
''  NAMEARCH$ = "TB_30714269026_003002_20171714_000005" & FECHA1$ & NEXTIDCOT
    NAMEARCH$ = "TB_" & TRIM$(CUIT1$) & "_000001_" & FECHA1$ & "_" & FORMATOCON0$(IDENCA&, 6)
    CARPETOK$ = LUCOM$ & NAMEARCH$ & ".TXT"
    NARCHI2% = FILEOPEN%(CARPETOK$, 2, 128)  ' ABRE EL ARCHIVO DE ESCRITURA
    TX1$ = "01" & "|"
    TX1$ = TX1$ & CUIT1$ & vbCrLf
    
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
        NREMI$ = TRIM$(Mid$(GRID.TextMatrix(i&, 8), 6))
        TX1$ = TX1$ & "91 R" & PREFIJO$ & NREMI$ & "|" '5
        TX1$ = TX1$ & FECHAENT$ & "|" '6 F.SALIDA MERCADERIA
        HORAX$ = Format(Time, "HH:MM")
        HORAX$ = REPLACAR(HORAX$, ":", "")
        TX1$ = TX1$ & HORAX$ & "|" '7
        LETRA$ = TRIM$(Text7.TEXT) 'E POR QUE ES ENVIO
        TX1$ = TX1$ & LETRA$ & "|" '8
'        If LETRA$ = "E" Then
'          VentaoDevolucion$ = "0" 'venta
'          ImporteRemito$ = TRIM$(Label3)
'          CUITX$ = TRIM$(CUITCLI(NC))
'          RazSocSale$ = TRIM$(oDE.razonSocial)
'        Else
          VentaoDevolucion$ = "1" 'devolucion
          ImporteRemito$ = "0"
          CUITX$ = TRIM$(oDE.Cuit) ' CUIT DE EMPRESA
          RazSocSale$ = TRIM$(rasocli$(NC))

'        End If
        
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
        NombreEmpresa$ = Left(oDE.razonSocial, 50)

        TX1$ = TX1$ & Left(TRIM$(NombreEmpresa$), 50) & "|" '13
        TX1$ = TX1$ & DESTINATARIO$ & "|" '14
          DomicilioEntrega$ = LENTREGA(0)
        TX1$ = TX1$ & DomicilioEntrega$ & "|" '15
        TX1$ = TX1$ & "|S/N"  '16
        TX1$ = TX1$ & "|"  '17
        TX1$ = TX1$ & "|"  '18
        TX1$ = TX1$ & "|"  '19
        TX1$ = TX1$ & "|"  '20
         CODIPOSTAL$ = Left$(TRIM$(TXTCPOSCLI), 8)
        TX1$ = TX1$ & CODIPOSTAL$ & "|"  '21
         LocalidadEntrega$ = Left$(TRIM$(LENTREGA(1)), 50)
        TX1$ = TX1$ & LocalidadEntrega$ & "|"  '22
         PCIA$ = TRIM$(LetraProvArbaSegunLetraFlexoft$(Me.Text8))
        TX1$ = TX1$ & PCIA$ & "|"  '23
        TX1$ = TX1$ & "|"  '24
        TX1$ = TX1$ & "NO|"  '25
        '
        CUITCLIENTE$ = CuitCli(NC)
        CUITCLIENTE$ = REPLACAR$(CUITCLIENTE$, "-", ""): CUITCLIENTE$ = REPLACAR$(CUITCLIENTE$, ".", ""): CUITCLIENTE$ = REPLACAR$(CUITCLIENTE$, "_", "")
        TX1$ = TX1$ & Left(CUITCLIENTE$, 11) & "|" '26
        TX1$ = TX1$ & Left(TRIM$(rasocli$(NC)), 50) & "|"  '27
        TX1$ = TX1$ & "0|"   '28
'        If LETRA$ = "E" Then
'          TX1$ = TX1$ & Left(TRIM$(oDE.CalleNumero), 40) & "|"  '29
'        Else
         DomicilioEntrega$ = TRIM$(GRID.TextMatrix(i&, 10))
          TX1$ = TX1$ & Left(DomicilioEntrega$, 40) & "|"  '29 DEVOLUCION
'        End If
        TX1$ = TX1$ & "|S/N"  '30
        TX1$ = TX1$ & "|"  '31
        TX1$ = TX1$ & "|"  '32
        TX1$ = TX1$ & "|"  '33
        TX1$ = TX1$ & "|"  '34
        '
'        If LETRA$ = "E" Then
'           CodPostSale$ = Left(TRIM$(oDE.CodPostal), 8)
'           LocalSale$ = Left(TRIM$(oDE.Localidad), 50)
'           LetraSale$ = Left(TRIM$(LetraProvArbaSegunLetraFlexoft$(oDE.LetProvincia)), 1)
'        Else
           CodPostSale$ = Left$(TRIM$(GRID.TextMatrix(i&, 12)), 8)
           LocalSale$ = Left(TRIM$(TRIM$(GRID.TextMatrix(i&, 11))), 50)
           LetraSale$ = Left(TRIM$(LetraProvArbaSegunLetraFlexoft$(TRIM$(GRID.TextMatrix(i&, 14)))), 1)
'        End If
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
      
      X1X = TRIM$(UnidMedidaArbaSegunUnidFlexoft$(Unimed$(CI1%), RETNUM%))
      If X1X = "" Then X1X = "UNIDADES"
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
'      Debug.Print TX1$
    Next i&
    '
    TX1$ = TX1$ & "04|"
    TX1$ = TX1$ & SAFETEXT$(XVALO(Label2.Caption))
    Print #NARCHI2%, TX1$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
    Close #NARCHI2%
    ret_DATOSCOT$ = TRIM$(CUIT1$) & " " & CONTROL$("REMITO", "PASSCOT") & " " & Mid$(NAMEARCH$, 4)
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
    Dim valor_NuevoID As Integer

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
 
    PVTA& = XVALO(Left$(REMI$, 4))
    NUREM& = XVALO(Mid$(REMI$, 6))
    Dim OBJCMD As ADODB.Command 'El objeto
    Set OBJCMD = New ADODB.Command
    OBJCMD.CommandType = adCmdStoredProc
    OBJCMD.CommandText = "AGREGA_DETACOT_2"
    Set OBJCMD.ActiveConnection = FLEXCONN
    Dim valor_NuevoID As Integer
    'Los parametros
    OBJCMD.Parameters(1) = CODIEMPR% '@CodiEmpr smallint
    OBJCMD.Parameters(2) = FETEXCOR1$(CVINT(Date)) '@FechMov smalldatetime
    OBJCMD.Parameters(3) = FECHAFORMATOINVERSO$(HOY(HO$)) '@FeFormCot nvarchar  (24)
    OBJCMD.Parameters(4) = CI1% '@Cod_Inte smallint
    OBJCMD.Parameters(5) = Codigo$ '@Cod_Exte nvarchar  (48)
    OBJCMD.Parameters(6) = Left$(DESCRICION$, 64) '@Descrip nvarchar  (128)
    OBJCMD.Parameters(7) = ROUND(VALUNIT#, 2) '@ValoUnit float  (8)
    OBJCMD.Parameters(8) = NUMCLI& '@Nume_Clie int
    If CANTI > 0 Then
      OBJCMD.Parameters(9) = 0 '@CantIngre float  (8),
      OBJCMD.Parameters(10) = CANTI ' @CantSalid float  (8),
    Else
      OBJCMD.Parameters(9) = Abs(CANTI) '@CantIngre float  (8),
      OBJCMD.Parameters(10) = 0 ' @CantSalid float  (8),
    End If
    OBJCMD.Parameters(11) = PVTA& '@NUMPVTA smallint, @
    OBJCMD.Parameters(12) = NUREM& 'DOPRINUM int,
    OBJCMD.Parameters(13) = "RT-X-" & REMI$ '@DoPriLar nvarchar  (48),
    OBJCMD.Parameters(14) = IDENCAX& '@ID_ENCA int,
    OBJCMD.Parameters(15) = DOMENTRE$
    OBJCMD.Parameters(16) = LOCALIDADENTR$
    OBJCMD.Parameters(17) = COPOS$
    OBJCMD.Parameters(18) = CUITTRANSP$
    OBJCMD.Parameters(19) = TRIM$(PCIA$)
    OBJCMD.Parameters(20) = DOMITRANSP$
    OBJCMD.Parameters(21) = LOCALITRANSP$
    OBJCMD.Parameters(22) = PATENTE$
    OBJCMD.Parameters(23) = DNITRANS$
    OBJCMD.Parameters(24) = CVDAT(FE_ORIG_REMITO%) 'FECHA ORIGINAL DEL REMITO
    OBJCMD.Parameters(25) = "91 R" & REMI$
    OBJCMD.Execute
    
    'ALTER  PROCEDURE [dbo].[AGREGA_DETACOT_2] ( , , , , , , , ,     @Code         int = NULL output )
    'Capturo el valor de retorno
    valor_NuevoID = OBJCMD.Parameters(0).Value
    Set OBJCMD = Nothing
End Sub

Sub limpiarDatoConductor()
    Text12 = ""
    Text18 = ""
    Text19 = ""
'    Text13 = ""

End Sub

Sub MENU_CELDA(MSF As MSFlexGrid, REG&, COL%)
        If MSF.Rows <= 1 Then Exit Sub
        
'        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
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

  Dim obj As New RECORXET
  Set rs1 = obj.RS_COTSLECCIONADO(NROSOLCOT&)
  Set obj = Nothing
  With rs1
    Do While Not .EOF

             J& = MSF.Rows
             MSF.Rows = J& + 1
             MSF.TextMatrix(J&, 1) = XVALO(!COD_INTE)
             MSF.TextMatrix(J&, 2) = J&
             MSF.TextMatrix(J&, 3) = SAFETEXT$(!Descrip)
             
             MSF.TextMatrix(J&, 4) = TRIM$(FORMATNUM$(XVALO(!CANTSALID), "F9.1"))
             MSF.TextMatrix(J&, 5) = TRIM$(FORMATNUM$(XVALO(!VALOUNIT), "F12.2"))
             IMPORTEX = (XVALO(!CANTSALID)) * XVALO(!VALOUNIT)
             ACUMULATOTAL = ACUMULATOTAL + IMPORTEX
             CONDI$ = "DOPRILAR = '" & SAFETEXT$(!doprilar) & "' AND COD_INTE = " & XVALO(!COD_INTE)
             NROPED& = XVALO(VALOBADA("MOVISTO", "NUPEDCLI", CONDI$, ""))
             MSF.TextMatrix(J&, 6) = NROPED&
             MSF.TextMatrix(J&, 7) = ""
             MSF.TextMatrix(J&, 8) = Mid$(SAFETEXT$(!doprilar), 6)
             MSF.TextMatrix(J&, 9) = FECHATEX$(CVINT(!FEREMITO))
             MSF.TextMatrix(J&, 10) = SAFETEXT$(!DOM_ENTRE)
             MSF.TextMatrix(J&, 11) = SAFETEXT$(!LOCALIDAD_ENTR)
              MSF.TextMatrix(J&, 12) = SAFETEXT$(!COPOSTAL)
             MSF.TextMatrix(J&, 13) = "" 'SAFETEXT$(!CUITTRANSP) 'CUIT TRANSPORTE
             MSF.TextMatrix(J&, 14) = SAFETEXT$(!PROVINCIA)
             MSF.TextMatrix(J&, 15) = SAFETEXT$(!DOMITRANSP) 'DOMICILIO TRANSPORTE
             MSF.TextMatrix(J&, 16) = SAFETEXT$(!LOCALITRANSP)  'LOCALIDAD TRANSPORTE
             MSF.TextMatrix(J&, 17) = "" 'SAFETEXT$(!PATENTE) 'PATENTE
             MSF.TextMatrix(J&, 18) = XVALO(!Nume_Clie) 'CLIENTE
             MSF.TextMatrix(J&, 19) = SAFETEXT$(!DNITRANS)  'DNI CLIENTE
             .MoveNext
    Loop
  End With
  rs1.Close
  Set rs1 = Nothing
  'PINTA LAS FILAS DONDE YA SE HAN REALIZADO DEVOLUCIONES.
  For J& = 1 To MSF.Rows - 1
     CI1% = XVALO(MSF.TextMatrix(J&, 1))
     REMI$ = TRIM$(MSF.TextMatrix(J&, 8))
     REMI1& = XVALO(Mid$(MSF.TextMatrix(J&, 8), 6))
     NCLIE = XVALO(MSF.TextMatrix(J&, 18))
     CONDI$ = " COD_INTE = " & CI1% & " AND NUME_CLIE= " & NCLIE & " AND DOPRINUM = " & REMI1& & " AND CANTINGRE > 0"
     If CantRegistros&("DETACOT", CONDI$, "NOMESS") > 0 Then
        Call COLOREAR_GRILLA_SOLAFILA_2(MSF, &HFFFF&, J&, , , 1) ' VERSION 2 MODIFICADO
     End If
  Next J&
  
  Call COMPLETAR_CASILLEROS(NROSOLCOT&)
  Call CalcutotalSeleccionado
  Label18 = MSF.TextMatrix(MSF.Rows - 1, 2)

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
    If TRIM$(oDE.localidad) = "" Then MENSAERROR$ = "Falta Localidad Empresa Activa ": GoTo 20
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
           LocalSale$ = TRIM$(oDE.localidad)
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

Private Sub Command4_Click()
    Command4.Enabled = False
    If GRID.Rows <= 1 Then
      Call COMUNI("No se Ha Seleccionado Ningún Remito")
      GoTo 99
    End If
    '
'***************************************************************************************************
    'COMENTADO POR QUE SE PODIA DAR QUE EXISTA UN ITEM REPETIDO EN UN MISMO REMITO Y CON ESTO NO LO DEJABA BAJAR
'    IDDETCOT& = XVALO(Text15) 'VEO SI YA FUE SOLICITADO
'    If IDDETCOT& > 0 Then
'      CONDI2$ = "ENCACOT_ID =" & IDDETCOT& 'VEO SI OBTUVO COT
'      COTX$ = TRIM$(VALOBADA("ENCACOT", "COT", CONDI2$, "NOMESS"))
'      If COTX$ <> "" Then
'         Call COMUNI("Atención !!!! \Remito: " & NUREMIX$ & " Ya Fué Procesado la Solicitud y Obtuvo C.O.T.")
'         GoTo 99
'      End If
'    End If
'***************************************************************************************************
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
      'GRABO EL DETALLE
      For J& = 1 To GRID.Rows - 1
             CI1% = XVALO(GRID.TextMatrix(J&, 1))
             Codigo$ = CODEXT$(CI1%)
             DESCRICION$ = Left$(GRID.TextMatrix(J&, 3), 64)
             CANTI = (-1) * XVALO(GRID.TextMatrix(J&, 4))
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
    Call Generar_Archivo_Cot_Dev(IDENCAX&, ret_DATOSCOT$, ret_NAMEARCH$) 'ret_CARPETOK$ retorna la ruta y nombre del archivo creado.
    IDCOT$ = ret_DATOSCOT$ ' VARIABLE A NIVEL DE FRM para pasar los parametros a .net
    NAMEARCHIVO$ = "COT_" + ret_NAMEARCH$ 'variable con el nombre del archivo que voy a buscar que genera el .net
    XXX = Shell("C:\Flexoft\COT\" + "COT.exe " & ret_DATOSCOT$, 4)
    Timer1.Enabled = True
    For R& = 1 To 40
      Call T_Espera(1 / 2)
      If Timer1.Enabled = False Then
         'SI EL TIMER SE PUSO EN FALSE ES POR QUE RECIBIO LA DEVOLUCION DE ARBA
         GoTo 50
      End If
    Next R&
    
50  If ID_COT& = 0 Then 'SE GENERO EL COT, CUANDO OBTIENE COT SATISFACTORIO LA VARIABLE GLOBAL ID_COT& SE PONE EN 0
       FLEXCONN.CommitTrans 'CONFIRMA LA TRANSACCION
    Else 'NO SE GENERO EL COT, SE VUELVE PARA ATRAS LA TRANSACCION.
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
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Id/A10;Provincia/A30", FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Id/A10;Provincia/A30", FRMZELECHO)
   FRMZELECHO.Caption = "Listado de Provincias"
   For i& = 1 To ULTREG&("PROVINC")
     x$ = REGLEIDO$("PROVINC", i&)
     LETRA$ = Left$(x$, 1)
     PROVINICIA$ = TRIM$(Mid$(x$, 2))
     FRMZELECHO.msf2.Rows = i& + 1
     FRMZELECHO.msf2.TextMatrix(i&, 1) = LETRA$
     FRMZELECHO.msf2.TextMatrix(i&, 2) = PROVINICIA$
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
     If VALIDA_COD_GRID%(NUDOCU$, Me.GRID, 9, MEXAGE$) < 1 Then
        Text3 = NUDOCU$
     End If
   
99  Command7.Enabled = True
End Sub

Private Sub Command8_Click()

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
'    If Me.lblRemSel.Caption = GRID.TextMatrix(i&, 8) Then
'       Call lblRemSel_Change
'    Else
'      Me.lblRemSel.Caption = GRID.TextMatrix(i&, 8)
'    End If
    
End Sub

Private Sub GRID_DblClick()
   If (GRID.MouseRow = 0) Then Exit Sub
   REG& = GRID.MouseRow
   If REG& < 1 Or GRID.Rows <= 1 Then Exit Sub
   '

   MSF.Rows = MSF.Rows + 1
   MSF.Cols = GRID.Cols
   MSF.TextMatrix(MSF.Rows - 1, 1) = GRID.TextMatrix(REG&, 1)
   MSF.TextMatrix(MSF.Rows - 1, 2) = GRID.TextMatrix(REG&, 2)
   MSF.TextMatrix(MSF.Rows - 1, 3) = GRID.TextMatrix(REG&, 3)
   MSF.TextMatrix(MSF.Rows - 1, 4) = GRID.TextMatrix(REG&, 4)
   MSF.TextMatrix(MSF.Rows - 1, 5) = GRID.TextMatrix(REG&, 5)
   MSF.TextMatrix(MSF.Rows - 1, 6) = GRID.TextMatrix(REG&, 6)
   MSF.TextMatrix(MSF.Rows - 1, 7) = GRID.TextMatrix(REG&, 7)
   MSF.TextMatrix(MSF.Rows - 1, 8) = GRID.TextMatrix(REG&, 8)
   MSF.TextMatrix(MSF.Rows - 1, 9) = GRID.TextMatrix(REG&, 9)
   MSF.TextMatrix(MSF.Rows - 1, 10) = GRID.TextMatrix(REG&, 10)
   MSF.TextMatrix(MSF.Rows - 1, 11) = GRID.TextMatrix(REG&, 11)
   MSF.TextMatrix(MSF.Rows - 1, 12) = GRID.TextMatrix(REG&, 12)
   MSF.TextMatrix(MSF.Rows - 1, 13) = GRID.TextMatrix(REG&, 13)
   MSF.TextMatrix(MSF.Rows - 1, 14) = GRID.TextMatrix(REG&, 14)
   MSF.TextMatrix(MSF.Rows - 1, 15) = GRID.TextMatrix(REG&, 15)
   MSF.TextMatrix(MSF.Rows - 1, 16) = GRID.TextMatrix(REG&, 16)
   MSF.TextMatrix(MSF.Rows - 1, 17) = GRID.TextMatrix(REG&, 17)
   MSF.TextMatrix(MSF.Rows - 1, 18) = GRID.TextMatrix(REG&, 18)
   MSF.TopRow = MSF.Rows - 1
   MSF.Row = MSF.Rows - 1
   MSF.ColSel = 1
   MSF.SetFocus
   '
   'ORDENO LA GRILLA SUPERIOR CUANDO SUBE
   MSF.COL = 6 'xx Columan por la q quieres ordenar
   MSF.Sort = 1 'Aplicar orden al grid

80  Call cmdEliminaItem(Me.GRID, GRID.Row)
   Call CalcutotalSeleccionado 'VER LA ACTUALIZACION DEL IMPORTE.
   Label18 = GRID.Rows - 1

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


Private Sub Command2_Click()
   Command2.Enabled = False
   Call ACTUALIZARFECHAREMITO_EN_DETACOT
   'Call AgregarDatosdeEntregaAGrilla(Me.lblRemSel)
   Command2.Enabled = True
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
'     Me.Text3 = NUDOCU$
    
    
99   Command5.Enabled = True
End Sub

Private Sub Form_Load()
    Campos$ = "CI/F0;Ord./F4;Descripción/A32;Salida/F9.1;P.Unit./F12.2;Nro.Ped./F7;"
    Campos$ = Campos$ + "N°Serie/A0;Remito/A16;F.Remito/D8;Domicilio/A20;Localidad/A20;C.P./A6;"
    Campos$ = Campos$ + "CUIT/A0;PCIA/A20;DOMITRNAS/A0;LOCATRANS/A0;PATENTE/A0;Cte/F8;Dni/A12"
    
    Call CARGA_ENCABEZADO(Me.GRID, Campos$, Me)
    Call CARGA_ENCABEZADO(Me.MSF, Campos$, Me)
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
    
    Dim oDE As New DatoEmpresa 'instancio cla
    CUIT1$ = TRIM$(oDE.Cuit)
    'REEEMPLAZO GUIONES POR VACIO Y POR LAS DUDAS PUNTO O GUION BAJO
    CUIT1$ = REPLACAR$(CUIT1$, "-", ""): CUIT1$ = REPLACAR$(CUIT1$, ".", ""): CUIT1$ = REPLACAR$(CUIT1$, "_", "")
    LENTREGA(0) = oDE.CalleNumero
    LENTREGA(1) = oDE.localidad
    TXTCPOSCLI = oDE.CodPostal
    Text8 = oDE.LetProvincia
    
    

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
      
Private Sub mnuReimprimirCot_Click()
'   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "ID Solicitud/F12;Fecha Solicitud;C.O.T./A14", FRMZELECHO)
'   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID Solicitud/F12;Fecha Solicitud;C.O.T./A14", FRMZELECHO)
'   FRMZELECHO.Caption = "SOLICITUD DE COT APROBADAS"
'
'   Dim obj As New RECORXET
'   Set RS1 = obj.RS_Solicitud_COT_OK
'   Call Carga_MSF_Recordset(RS1, "ID Solicitud/F12;Fecha Solicitud;C.O.T./A14", FRMZELECHO.msf2, 1)
'   Set obj = Nothing
'   RS1.Close
'   Set RS1 = Nothing
'
'   FRMZELECHO.Width = 6500
'   FRMZELECHO.Show 1
'   NRO& = XVALO(RESP_ZELE_VECT(1))
'   If NRO& < 1 Then Exit Sub
'   '
'   CONDI1$ = "ID_ENCA =" & NRO&
'   PATENTE$ = TRIM$(VALOBADA("DETACOT", "PATENTE", CONDI1$, "NOMESS"))
'   CODINTEGRIDAD$ = TRIM$(VALOBADA("DETACOT", "CODINTEGRIDAD", CONDI1$, "NOMESS"))
'
'   CONDI2$ = "ENCACOT_ID =" & NRO&
'   ValorProcesado$ = TRIM$(VALOBADA("ENCACOT", "COT", CONDI2$, "NOMESS"))
'   CODINTEGRIDAD$ = TRIM$(VALOBADA("ENCACOT", "CODINTEGRIDAD", CONDI2$, "NOMESS"))
'
'   FE% = CVINT(VALOBADA("ENCACOT", "FESALIDA", CONDI2$, "NOMESS"))
'
'
'   FILTX$ = TRIM$(ValorProcesado$) & ";" & FECHATEX$(FE%) & ";" & PATENTE$ & ";" & CODINTEGRIDAD$ & ";" & TRIM$(Str$(NRO&))
'   Call STDFORM("COTARBA", FILTX$) 'MUESTRA EL FORMULARIO PARA IMPRMIR

End Sub

Private Sub mnuremdevol_Click()
   
   Campos$ = "ID C.O.T./F10;Fecha Solicitud/D8;Cta/F8;R.Social/A24;Remito/A20;C.O.T./A12"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = "SOLICITUD DE COT APROBADAS"

   Dim obj As New RECORXET
   Set rs1 = obj.RS_Solicitud_COT_OK
   Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.msf2, 1)
   Set obj = Nothing
   rs1.Close
   Set rs1 = Nothing

   FRMZELECHO.Width = 11000
   FRMZELECHO.Show 1
   NRO& = XVALO(RESP_ZELE_VECT(1))
   REMI$ = Mid$(RESP_ZELE_VECT(5), 6)
   
   'VALIDO SI EL REMITO SELECCIONADO YA ESTA EN LA FILA SUPERIOR
   EnFila& = VALIDA_COD_GRID%(REMI$, Me.MSF, 8, "Este Remito Ya Fue Seleccionado")
   If EnFila& > 0 Then
     Call SELECCIONARFILA(Me.MSF, EnFila&)
     Exit Sub
   End If
   
   'VALIDO SI EL REMITO SELECCIONADO YA ESTA EN LA FILA INFERIOR
   EnFila& = VALIDA_COD_GRID%(REMI$, Me.GRID, 8, "Este Remito Ya Fue Seleccionado")
   If EnFila& > 0 Then
     Call SELECCIONARFILA(Me.GRID, EnFila&)
     Exit Sub
   End If
   
   Text15 = NRO&: Text14.Visible = True: Text15.Visible = True
   
   If NRO& < 1 Then Exit Sub
   If MSF.Rows > 1 Or GRID.Rows > 1 Then
      OPX% = COMALTER%("Ya Hay Cargada Información en Pantalla\\Agregar al Detalle\Limpiar Grillas y Cargar Detalle de Cot Seleccionado")
   End If
   '
   If OPX% > 1 Then
    MSF.Rows = 1
    GRID.Rows = 1
    Call LIMPIA_CAJA_TEXTO: Text3 = "": Call limpiarDatoConductor
   End If
   '

   Call MOSTRAR_SOLICITUD_DE_COT(NRO&)

End Sub

Private Sub mnusolaprobcot_Click()
'   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "ID Solicitud/F12;Fecha Solicitud/D8;C.O.T./A12", FRMZELECHO)
'   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID Solicitud/F12;Fecha Solicitud/D8;C.O.T./A12", FRMZELECHO)
'   FRMZELECHO.Caption = "SOLICITUD DE COT APROBADAS"
'
'   Dim obj As New RECORXET
'   Set RS1 = obj.RS_Solicitud_COT_OK
'   Call Carga_MSF_Recordset(RS1, "ID Solicitud/F12;Fecha Solicitud/D8;C.O.T./A12", FRMZELECHO.msf2, 1)
'   Set obj = Nothing
'   RS1.Close
'   Set RS1 = Nothing
'
'   FRMZELECHO.Width = 6500
'   FRMZELECHO.Show 1
'   NRO& = XVALO(RESP_ZELE_VECT(1))
'   Text15 = NRO&: Text14.Visible = True: Text15.Visible = True
'
'   If NRO& < 1 Then Exit Sub
'   '
'   Call MOSTRAR_SOLICITUD_DE_COT(NRO&)
   

End Sub

Private Sub mnuVerRemDev_Click()
   Campos$ = "ID C.O.T./F10;Fecha Solicitud/D8;Cta/F8;R.Social/A24;Remito/A20;C.O.T./A12"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = "SOLICITUD DE COT APROBADAS"

   Dim obj As New RECORXET
   Set rs1 = obj.RS_Solicitud_COT_OK("D")
   Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.msf2, 1)
   Set obj = Nothing
   rs1.Close
   Set rs1 = Nothing


   FRMZELECHO.Width = 11000
20 FRMZELECHO.Show 1
   NRO& = XVALO(RESP_ZELE_VECT(1))
   If NRO& < 1 Then Exit Sub
   '
   Condi1$ = "ID_ENCA =" & NRO&
   PATENTE$ = TRIM$(VALOBADA("DETACOT", "PATENTE", Condi1$, "NOMESS"))
'   CODINTEGRIDAD$ = TRIM$(VALOBADA("DETACOT", "CODINTEGRIDAD", CONDI1$, "NOMESS"))

   CONDI2$ = "ENCACOT_ID =" & NRO&
   ValorProcesado$ = TRIM$(VALOBADA("ENCACOT", "COT", CONDI2$, "NOMESS"))
   CODINTEGRIDAD$ = TRIM$(VALOBADA("ENCACOT", "CODINTEGRIDAD", CONDI2$, "NOMESS"))

   FE% = CVINT(VALOBADA("ENCACOT", "FESALIDA", CONDI2$, "NOMESS"))
   

   FILTX$ = TRIM$(ValorProcesado$) & ";" & FECHATEX$(FE%) & ";" & PATENTE$ & ";" & CODINTEGRIDAD$ & ";" & TRIM$(Str$(NRO&))
   Call STDFORM("COTARBD", FILTX$) 'MUESTRA EL FORMULARIO PARA IMPRMIR

   GoTo 20
End Sub

Private Sub MSF_DblClick()
   If (MSF.MouseRow = 0) Then Exit Sub
   REG& = MSF.MouseRow
   If REG& < 1 Or MSF.Rows <= 1 Then Exit Sub
   '
   If TRIM$(MSF.TextMatrix(REG&, 0)) = "*" Then
      Call COMUNI("Atención!!!\Este Remito Yá Posee Realizada Devoluciones\Recuerde Que Si Yá Se Han Realizado Devoluciones Parciales\No Podra Volver a Solicitar Otro C.O.T.")
      Exit Sub
   End If
   '
   'VALIDO QUE EL MISMO NUMERO DE REMITO NO HAYA TENIDO ALGUNA DEVOLUCION PREVIA
   REMI$ = TRIM$(MSF.TextMatrix(REG&, 8))
   For i& = 1 To MSF.Rows - 1
     If TRIM$(MSF.TextMatrix(i&, 0)) = "*" And REG& <> i& Then 'PARA QUE NO SE VALIDE A SI MISMO
        If REMI$ = MSF.TextMatrix(i&, 8) Then
          Call COMUNI("Atención!!!\Este Remito Yá Posee Realizada Devoluciones Parciales\Recuerde Que Si Yá Se Han Realizado Devoluciones\No Podra Volver a Solicitar Otro C.O.T.\(A.R.B.A. Valida Por Nro. de Remito  - No por Item)")
          Exit Sub
        End If
     End If
   Next i&
   
'   RESULTA& = 0
'     REG_ARRIBA$ = STRINGFILA$(MSF, REG&)
'     For J& = 1 To GRID.Rows - 1
'        REG_ABAJO$ = STRINGFILA$(GRID, J&)
'        If REG_ARRIBA$ = REG_ABAJO$ Then
'           Call COMUNI("Esta Fila Yá Fué Seleccionda")
'           Call SELECCIONARFILA(MSF, REG&)
'           Call SELECCIONARFILA(GRID, J&)
'           Exit Sub
'        End If
'     Next J&

    GRID.Rows = GRID.Rows + 1
    GRID.Cols = MSF.Cols
    GRID.TextMatrix(GRID.Rows - 1, 1) = MSF.TextMatrix(REG&, 1)
    GRID.TextMatrix(GRID.Rows - 1, 2) = MSF.TextMatrix(REG&, 2)
    GRID.TextMatrix(GRID.Rows - 1, 3) = MSF.TextMatrix(REG&, 3)
    GRID.TextMatrix(GRID.Rows - 1, 4) = MSF.TextMatrix(REG&, 4)
    GRID.TextMatrix(GRID.Rows - 1, 5) = MSF.TextMatrix(REG&, 5)
    GRID.TextMatrix(GRID.Rows - 1, 6) = MSF.TextMatrix(REG&, 6)
    GRID.TextMatrix(GRID.Rows - 1, 7) = MSF.TextMatrix(REG&, 7)
    GRID.TextMatrix(GRID.Rows - 1, 8) = MSF.TextMatrix(REG&, 8)
    GRID.TextMatrix(GRID.Rows - 1, 9) = MSF.TextMatrix(REG&, 9)
    GRID.TextMatrix(GRID.Rows - 1, 10) = MSF.TextMatrix(REG&, 10)
    GRID.TextMatrix(GRID.Rows - 1, 11) = MSF.TextMatrix(REG&, 11)
    GRID.TextMatrix(GRID.Rows - 1, 12) = MSF.TextMatrix(REG&, 12)
    GRID.TextMatrix(GRID.Rows - 1, 13) = MSF.TextMatrix(REG&, 13)
    GRID.TextMatrix(GRID.Rows - 1, 14) = MSF.TextMatrix(REG&, 14)
    GRID.TextMatrix(GRID.Rows - 1, 15) = MSF.TextMatrix(REG&, 15)
    GRID.TextMatrix(GRID.Rows - 1, 16) = MSF.TextMatrix(REG&, 16)
    GRID.TextMatrix(GRID.Rows - 1, 17) = MSF.TextMatrix(REG&, 17)
    GRID.TextMatrix(GRID.Rows - 1, 18) = MSF.TextMatrix(REG&, 18)
    GRID.TextMatrix(GRID.Rows - 1, 19) = MSF.TextMatrix(REG&, 19)
    GRID.TopRow = GRID.Rows - 1
    GRID.Row = GRID.Rows - 1
    GRID.ColSel = 1
    GRID.SetFocus
   '
    Call numerargrilla(Me.GRID, 2, 0)
80  Call cmdEliminaItem(Me.MSF, MSF.Row)
    Call numerargrilla(Me.MSF, 2, 0)

    Call CalcutotalSeleccionado
    Label18 = GRID.Rows - 1
End Sub


Private Sub MSF_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
     REG& = MSF.MouseRow
     COL% = MSF.MouseCol
     Call MENU_CELDA(MSF, REG&, COL%)
     Call CalcutotalSeleccionado 'RECALCULO POR SI ELIMINA
     Me.lblRemSel = ""
  End If

End Sub


Private Sub QuitarRemi_Click()
   NroRemi$ = TRIM$(MSF.TextMatrix(GLOBALFRM_FILA&, 8))
   If NroRemi$ = "" Then Call COMUNI("Falta Nro. de Remito"): Exit Sub
   '
   i& = 0
   Do While i& < MSF.Rows - 1
      i& = i& + 1
      If TRIM$(MSF.TextMatrix(i&, 8)) = NroRemi$ Then
       Call cmdEliminaItem(MSF, i&)
       i& = i& - 1
      End If
   Loop
   If MSF.Rows < 2 Then
      Text3 = ""
   Else
      Call numerargrilla(MSF, 1)
   End If
   
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
'''    If Command1.Enabled = True Then
''''      Label18 = ""
'''    End If
'''    '
'''    '
'''    If XVALO(Text15) > 0 Then GRID.Rows = 1
'''
'''    Text15 = "": Text14.Visible = False: Text15.Visible = False
'''    NUREMIX$ = "RT-X-" & Text3
'''    CONDI$ = "DoPriLar = '" & NUREMIX$ & "'"
'''    IDDETCOT& = XVALO(VALOBADA("DETACOT", "ID_ENCA", CONDI$, "NOMESS")) 'VEO SI YA FUE SOLICITADO
'''    If IDDETCOT& > 0 Then
'''      CONDI2$ = "ENCACOT_ID =" & IDDETCOT& 'VEO SI OBTUVO COT
'''      COTX$ = TRIM$(VALOBADA("ENCACOT", "COT", CONDI2$, "NOMESS"))
'''      If COTX$ <> "" Then
'''         Call COMUNI("Atención !!!! \Remito: " & NUREMIX$ & " Ya Fué Procesado la Solicitud y Obtuvo C.O.T.")
'''      Else
'''         Call COMUNI("Atención !!!! \Remito: " & NUREMIX$ & " Ya Fué Procesado la Solicitud Está Pendiente de Obtener C.O.T.\Identificador de Solicitud de C.O.T.: " & IDDETCOT&)
'''      End If
'''      Text3 = ""
'''      Exit Sub
'''    End If
'''    '
'''    Me.lblRemSel = Text3
'''
'''    Call LIMPIA_CAJA_TEXTO
'''    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
'''    If CMO$ = "" Then CMO$ = "RT"
'''    '
'''    VUELTA% = 0: NPED& = 0
'''    Call COLOREAR_GRILLASQL_TODAS_MISMO_COLOR(GRID, vbWhite)
'''    If Me.Text3 <> "" Then
'''        CONDI$ = "(CodiMovi = '" & CMO$ & "' OR CodiMovi = 'DV') "
'''        CONDI$ = CONDI$ + " AND DoPriLAR LIKE '%" & Text3 & "%' "
'''        VUELTA% = 0: J& = 0
'''        SQLX$ = "SELECT COD_INTE,NuorItem,DESCRI_LAR, CantSalid, cantingre, NUPEDCLI, FECHMOV,"
'''        SQLX$ = SQLX$ + " Nume_Clie, FECHMOV,IDENLOTE,ValoUnitIva FROM MOVISTO WITH(NOLOCK)"
'''        SQLX$ = SQLX$ + " WHERE " & CONDI$ & "  ORDER BY NuOrItem"
'''        Set RS = READSET(SQLX$)
'''        With RS
'''          Do While Not .EOF
'''             J& = GRID.Rows
'''             GRID.Rows = J& + 1
'''
'''             If VUELTA% = 0 Then
'''               VUELTA% = 1
'''               NCLIE = XVALO(!NUME_CLIE)
'''               Text2 = NCLIE
'''               Text6 = rasocli$(XVALO(Text2))
'''               Text4 = FECHATEX$(CVINT(!FECHMOV))
'''             End If
'''             If NCLIE_ANT <> NCLIE Then
'''                DOCUMENTOX$ = ""
'''                If PIVACLI%(NCLIE) < 1 Then
'''                    DOCUMENTOX$ = CUITCLI(NCLIE)
'''                End If
'''             End If
'''
'''             'UN REMITO PUEDE TENER VARIOS PEDIDOS SI ALGUNO TIENE PEDIDO LE PASO EL NRO PARA QUE MUESTRE LOS D.ENTR DEL PEDIDO.
'''             If XVALO(!NUPEDCLI) > 0 Then NPED& = XVALO(!NUPEDCLI)
'''             GRID.TextMatrix(J&, 1) = XVALO(!COD_INTE)
'''             GRID.TextMatrix(J&, 2) = J&
'''             GRID.TextMatrix(J&, 3) = SAFETEXT$(!DESCRI_LAR)
'''             GRID.TextMatrix(J&, 4) = TRIM$(FORMATNUM$(XVALO(!CANTSALID), "F9.1"))
'''             GRID.TextMatrix(J&, 5) = TRIM$(FORMATNUM$(XVALO(!VALOUNITIVA), "F12.2"))
'''             IMPORTEX = (XVALO(!CANTSALID)) * XVALO(!VALOUNITIVA)
'''             ACUMULATOTAL = ACUMULATOTAL + IMPORTEX
'''             GRID.TextMatrix(J&, 6) = XVALO(!NUPEDCLI)
'''             Call COLOREAR_GRILLA_SOLO_UNA_AFILA(GRID, &HE0E0E0, J&)
'''             GRID.TextMatrix(J&, 7) = SAFETEXT$(!IDENLOTE)
'''             GRID.TextMatrix(J&, 8) = Text3
'''             GRID.TextMatrix(J&, 9) = FECHATEX$(CVINT(!FECHMOV))
'''             GRID.TextMatrix(J&, 19) = DOCUMENTOX$
'''             NCLIE_ANT = NCLIE
'''
'''            .MoveNext
'''          Loop
'''        End With
'''        RS.Close
'''        Set RS = Nothing
'''        'OBTENGO LOS DATOS DE ENTREGA DEL PEDIDO SI VA SIN NRO DE PEDIDO O EL PEDIDO NO TIENE DATOS LE ASIGNA LOS DATOS
'''        'DE ENTREGA DE LA SOLAPA ENTREGA .
'''        Call CARGAR_DATOS_ENTREGA(XVALO(Text2), NPED&, Text3)
'''        Call CalcutotalSeleccionado
'''        Label18 = GRID.TextMatrix(GRID.Rows - 1, 2)
'''        '
'''
'''        '
''''        'MOSTRAR DATOS DE ENTREGA
''''        For i& = 1 To GRID.Rows - 1
''''           NPDX& = XVALO(GRID.TextMatrix(i&, 6))
''''           If NPDX& > 0 Then
''''             SQLX$ = "SELECT DOMIENT,LOCAENT,TranspEnt,NROOCOM,FLETE_TEX, NTranspo FROM PEDENCA WITH(NOLOCK)"
''''             SQLX$ = SQLX$ + " WHERE NROPED = " & NPDX&
''''             Set RS = READSET(SQLX$)
''''             With RS
''''             Do While Not .EOF
''''               If TRIM$(SAFETEXT$(!DOMIENT)) <> "" Then
''''                 LENTREGA(0) = SAFETEXT(!DOMIENT)
''''                 LENTREGA(1) = SAFETEXT(!LOCAENT)
''''                 NTRAN% = XVALO(!NTranspo)
''''                 If NTRAN% > 0 Then ' SI VIENE CON CODIGO DE TRANSPORTE COMPLETA AUTOMATICAMENTE
''''                   Text45 = NTRAN%
''''                 Else
''''                   LENTREGA(2) = SAFETEXT(!TranspEnt)
''''                 End If
''''                 Text20 = SAFETEXT(!NROOCOM)
''''                 TXTFLETE = SAFETEXT(!FLETE_TEX)
''''                 GoTo 20
''''               End If
''''               .MoveNext
''''             Loop
''''             End With
''''20           RS.Close
''''             Set RS = Nothing
''''           End If
''''        Next i&
'''    End If

End Sub

Sub CalcutotalSeleccionado()
        GRID.COL = 8
        GRID.Sort = 1
        Label3 = "": ACUMULATOTAL = 0
        For i& = 1 To GRID.Rows - 1
          NroRemi$ = GRID.TextMatrix(i&, 8)
          If NroRemi$ <> NROREMI_ANT$ Then CANTREM = CANTREM + 1: Label2.Caption = FORMATNUM$(CANTREM, "F8.1")
          IMPORTEX = (XVALO(GRID.TextMatrix(i&, 4)) * XVALO(GRID.TextMatrix(i&, 5)))
          ACUMULATOTAL = ACUMULATOTAL + IMPORTEX
          NROREMI_ANT$ = NroRemi$
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
        CalcutotalRemito$ = FORMATNUM$(ACUMULATOTAL, "F12.2")
End Function

Sub LIMPIARTEXT(FORMULARIO As Form, LIMPIOTEXT%)
  Dim CONTROLX As CONTROL
  For Each CONTROLX In FORMULARIO.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
        If TypeOf CONTROLX Is TextBox And LIMPIOTEXT% > 0 Then
           CONTROLX.TEXT = ""
        End If
  Next
End Sub



Sub LIMPIA_CAJA_TEXTO()
'    Text2 = ""
'    Text3 = ""
    Text4 = ""
    Text6 = ""
'    Text8 = ""
    Text21 = ""
    Text23 = ""
'    Me.LENTREGA(0) = ""
'    Me.LENTREGA(1) = ""
    Label18.Caption = ""
'    TXTCPOSCLI = ""
    Label2.Caption = ""
    Label3.Caption = ""
    
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
    J& = 0
    Dim DESCRIPCIONERROR$()
    '
'    X = COPYFILE%(CARPETOK$ + "COT_" + "123456.xml", CARPETOK$ + "COT_" + "123456.txt")
'** RECORRO EL ARCHIVO SI YA SE GENERO ****
'    IDCOT$ = IDCOT$
    If EXISTE%(CARPETOK$ + NAMEARCHIVO$ + ".TXT") > 0 Then
        NARCHIVO% = FILEOPEN%(CARPETOK$ + NAMEARCHIVO$ + ".TXT", 1, 256)  'ABRE EL ARCHIVO DE LECTURA
        Do While Not EOF(NARCHIVO%) ' RECORRE
          Line Input #NARCHIVO%, XXXXX$ ' LEE CADA REGISTRO
          TX1$ = TX1$ + XXXXX$ + vbCrLf
          If InStr(UCase$(XXXXX$), UCase$("<numeroComprobante>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<numeroComprobante>")) + 19
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</numeroComprobante>")
             COT$ = Mid$(TXT$, 1, POS2% - 1)
             'caso de punto de venta 4
          End If
          If InStr(UCase$(XXXXX$), UCase$("<codigoIntegridad>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<codigoIntegridad>")) + 19
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</codigoIntegridad>")
             CODINTEGRIDAD$ = Mid$(TXT$, 1, POS2% - 1)
          End If
          If InStr(UCase$(XXXXX$), UCase$("<procesado>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<procesado>")) + 11
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</procesado>")
             ValorProcesado$ = Mid$(TXT$, 1, POS2% - 1)
          End If
          If InStr(UCase$(XXXXX$), UCase$("<descripcion>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<descripcion>")) + 13
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</descripcion>")
             J& = J& + 1
             ReDim Preserve DESCRIPCIONERROR$(J&)
             DESCRIPCIONERROR$(J&) = Mid$(TXT$, 1, POS2% - 1)
          End If
          
        Loop
        Close NARCHIVO%
        
        '** DESHABILITO EL TIMER ****
        '** DE ACUERDO AL RESULTADO IMPRIMO SI FUE TODO BIEN ****
        If TRIM$(UCase$(ValorProcesado$)) = "SI" Then
           FE% = FECHANUM(Text20) ' fecha de pantalla
           CANTIDIAS% = COMBO1.ListIndex + 1
           If CANTIDIAS% < 0 Then CANTIDIAS% = 1
           FE% = DIASPOST(FE%, CANTIDIAS%)
           CONDI$ = "ENCACOT_ID = " & ID_COT&
           Call ACTUREGISTRO("ENCACOT", "COT", CONDI$, COT$, REGAF&, "NOMESS")
           Call ACTUREGISTRO("ENCACOT", "CODINTEGRIDAD", CONDI$, CODINTEGRIDAD$, REGAF&, "NOMESS")
           FILTX$ = TRIM$(COT$) & ";" & FECHATEX$(FE%) & ";" & TRIM$(Text18) & ";" & CODINTEGRIDAD$ & ";" & TRIM$(Str$(ID_COT&))
           Call STDFORM("COTARBD", FILTX$) 'MUESTRA EL FORMULARIO PARA IMPRMIR
           Call LIMPIA_CAJA_TEXTO: Text3 = "": Call limpiarDatoConductor
           GRID.Rows = 1
           ID_COT& = 0
           Timer1.Enabled = False

        Else
         '** CASO CONTRARIO MUESTRO UNA LISTA DE ERORRES DEVUELTOS. ****
           On Error Resume Next
           If UBound(DESCRIPCIONERROR$) > 0 Then
             Campos$ = "Descripciónes de Errores Devueltos x ARBA/A64"
             Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , 9000)
             Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO, , , 9000)
             FRMZELECHO.Caption = "Detalle de Errores Devueltos"
             FRMZELECHO.Width = 8000
             For J& = 1 To UBound(DESCRIPCIONERROR$)
                 FRMZELECHO.msf2.Rows = J& + 1
                 FRMZELECHO.msf2.TextMatrix(J&, 1) = DESCRIPCIONERROR$(J&)
             Next J&
             FRMZELECHO.Show 1
             Timer1.Enabled = False
           End If
           On Error GoTo 0
        End If
    End If
    '


End Function

