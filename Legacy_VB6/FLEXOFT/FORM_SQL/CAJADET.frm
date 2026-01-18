VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form CAJADET 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos de Caja"
   ClientHeight    =   8370
   ClientLeft      =   -15
   ClientTop       =   675
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   14610
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   600
      Left            =   0
      TabIndex        =   42
      Top             =   6140
      Width           =   5655
      Begin VB.CommandButton Command5 
         Caption         =   "Quitar Selección"
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
         Left            =   3480
         TabIndex        =   45
         Top             =   200
         Width           =   2055
      End
      Begin VB.TextBox TXTBUSCAR 
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
         Left            =   720
         TabIndex        =   44
         ToolTipText     =   "Escriba el Texto a Buscar y Presione Enter Para Seleccionar las Coincidencias"
         Top             =   160
         Width           =   2655
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "CAJADET.frx":0000
         TabIndex        =   43
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton Command19 
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
      Height          =   315
      Left            =   2700
      TabIndex        =   41
      Top             =   6840
      Width           =   255
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Index           =   7
      Left            =   12000
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Text            =   "Otros Medios: "
      Top             =   8040
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Index           =   5
      Left            =   13440
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1095
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Index           =   4
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1095
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Index           =   6
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Text            =   "Total Efectivo:"
      Top             =   8040
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   9720
      TabIndex        =   36
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtingreso 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   315
      Left            =   7320
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   7605
      Width           =   975
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Index           =   5
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   7605
      Width           =   3615
   End
   Begin VB.TextBox txtgastos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   315
      Left            =   7320
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Index           =   4
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6840
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Actualizar Información"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   18
      Top             =   1320
      Width           =   3015
   End
   Begin VB.PictureBox CommonDialog1 
      Height          =   480
      Left            =   6120
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   21
      Top             =   3960
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox Skin1 
      Height          =   480
      Left            =   8520
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Imprimir"
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
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7680
      Width           =   2415
   End
   Begin VB.TextBox Text20 
      Appearance      =   0  'Flat
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
      Height          =   315
      Index           =   1
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7605
      Width           =   4455
   End
   Begin VB.TextBox Text20 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Index           =   0
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6840
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar Retiro"
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
      Index           =   0
      Left            =   5040
      TabIndex        =   17
      Top             =   7200
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   315
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7605
      Width           =   975
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "Retiro de Caja:"
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "Ingreso de Caja:"
      Top             =   7605
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar Ingreso Manual"
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
      Index           =   1
      Left            =   5040
      TabIndex        =   14
      Top             =   8040
      Width           =   2175
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Index           =   8
      Left            =   12120
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "Salidas:"
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Index           =   2
      Left            =   12120
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "Resultado:"
      Top             =   7260
      Width           =   1095
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Index           =   0
      Left            =   13440
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6420
      Width           =   1095
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Index           =   1
      Left            =   12120
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Ingreso:"
      Top             =   6420
      Width           =   1095
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Index           =   1
      Left            =   13440
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Index           =   2
      Left            =   13440
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7260
      Width           =   1095
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Index           =   3
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "DETALLE DE CAJA"
      Top             =   0
      Width           =   14595
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   0
      TabIndex        =   5
      Top             =   620
      Width           =   14535
      Begin VB.CommandButton Command3 
         Caption         =   "Seleccionar M. Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Doble Click Para Limpiar"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox TXTDESCRIUSUARIO 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox TXTNUMUSUARIO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Doble Clic Para limpiar Casillero y ver Movimientos de Todos los Usuarios"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Selección de Usuarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Reg.:"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TXTCANTIFILAS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker Desde 
         Height          =   330
         Left            =   11280
         TabIndex        =   24
         Top             =   240
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   130023425
         CurrentDate     =   41123
      End
      Begin MSComCtl2.DTPicker Hasta 
         Height          =   330
         Left            =   12960
         TabIndex        =   25
         Top             =   240
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   130023425
         CurrentDate     =   41123
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   4215
      Left            =   0
      TabIndex        =   23
      Top             =   1920
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   7435
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuTRansfxUsuario 
         Caption         =   "Transferencias de Efectivo de  caja entre usuarios"
      End
   End
End
Attribute VB_Name = "CAJADET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
   If Index = 0 Then
     If DERACCE%("RETICAJA", "GENERAR RETIRO DE CAJA") < 2 Then Exit Sub
   End If
   If Index = 1 Then
     If DERACCE%("INGRCAJA", "GENERAR INGRESO DE CAJA") < 2 Then Exit Sub
   End If
   
   Command1(Index).Enabled = False
'   USNX% = USNBR% Mod 100
'   If USNX% <> 1 Then
'      Call COMUNI("Solo el Usuario Supervisor Puede Realizar Retiros o Ingresos a Caja")
'      GoTo 99
'   End If
   
   If TRIM$(Text20(Index)) = "" Then
        MsgBox "FALTA INGRESAR REFERENCIA"
        GoTo 99
   End If
   '
   If CUECOINT(txtgastos) < 1 Then
      Call COMUNI("Atención !! Falta Cuenta Contrapartida de Cuenta de Gastos")
      GoTo 99
   End If
   
   If CUECOINT(Me.txtingreso) < 1 Then
      Call COMUNI("Atención !! Falta Cuenta Contrapartida de Cuenta de Gastos")
      GoTo 99
   End If

   RFF$ = RECFILT$("MECOBRA", 1, MKI$(1)) ' SE DEBE CARGAR LA CUENTA DE EFECTIVO PRIMERA EN LOS  MEDIOS DE COBRANZA
   If Len(RFF$) >= 4 Then
     REME& = CVS(Mid$(RFF$, 1, 4))
     If REME& > 0 Then
       CUCU$ = Mid$(REGLEIDO$("MECOBRA", REME&), 45, 12)
       CMC% = CUECOINT%(TRIM$(CUCU$))
     End If
   End If
   If CMC% < 1 Then
     Call MENSERR(24, "Falta Cargar Medio de Cobranza Efectivo (Recuerde Agregar Esta en Primer Lugar")
     GoTo 99
   End If
   '
   
   
   Select Case (Index)
     Case 0
        If XVALO(Text2(Index)) > XVALO(Text15(4)) Then
           Call COMUNI("Imposible Registrar Retiro Mayor al Total de Efectivo de Esta Caja")
           GoTo 99
        End If
        If XVALO(Text2(Index)) > 0 Then Call GRABA_MOVI_CAJA(XVALO(Text2(0)), 0, CMC%)
     Case 1
        If XVALO(Text2(Index)) > 0 Then Call GRABA_MOVI_CAJA(0, XVALO(Text2(1)), CMC%)
   End Select
'   MsgBox "HECHO"
   Call Command2_Click
   Text2(0) = ""
   Text2(1) = ""
   Text20(0) = ""
   Text20(1) = ""
99 Command1(Index).Enabled = True

End Sub
Sub GRABA_MOVI_CAJA(Sali#, INGR#, CMC%)
    
    SALIDA# = XVALO(Sali#)
    INGRESO# = XVALO(INGR#)
    If SALIDA# > 0 Then
       TiMovi$ = "SALECAJA"
       VARIMOVIX$ = "SALIDA"
       REFE$ = "(SALIDA) " & TRIM$(Text20(0))
       NUMEDOCU& = 1 + MAXNUMCAMPO&("CAJABANC", "NUMECOMP", "VARIMOVIM = 'SALIDA DE CAJA' AND TIPOMOVIM = 'SALIDA DE CAJA'")
       CueConta% = CUECOINT(Me.txtgastos)
       ImpoDeb# = Sali#
    ElseIf INGRESO# > 0 Then
       TiMovi$ = "INGRCAJA"
       VARIMOVIX$ = "INGRESO"
       REFE$ = TRIM$(Text20(1))
       NUMEDOCU& = 1 + MAXNUMCAMPO&("CAJABANC", "NUMECOMP", "VARIMOVIM = 'INGRESO MANUAL' AND TIPOMOVIM = 'INGRESO MANUAL'")
       CueConta% = CUECOINT(Me.txtingreso)
       ImpoDeb# = -1 * INGR#
    End If
    
   '
   US1% = XVALO(TXTNUMUSUARIO)
   If Abs(ImpoDeb#) < 0.005 Then
     Call MENSERR(24, "Falta Importe")
     GoTo 99
   End If
   '
   If US1% < 1 Then
     Call MENSERR(24, "Falta Usuario ")
     GoTo 99
   End If
   
   If TRIM$(REFE$) = "" Then
     Call MENSERR(24, "Falta Referencia")
     GoTo 99
   End If
   '
   FEDEP% = HOYSERVER%
   Screen.MousePointer = 11
   FETX$ = FECHATEX$(HOY(HO$))
   NRECIB$ = "TF-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
   NRECIBX$ = NRECIB$ + "__"
   ORMA% = (-1)
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE CodiCom_Cor LIKE '" & NRECIBX$ & "'"
   Dim CAJABTEMP As ADODB.Recordset
   Set CAJABTEMP = READSET(SQLX$)

   On Error Resume Next
   With CAJABTEMP
30   'If .NoMatch = False Then
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
          If Val(Mid$(!CodiCom_Cor, 11)) > ORMA% Then
            ORMA% = Val(Mid$(!CodiCom_Cor, 11))
          End If
          .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   Set CAJABTEMP = Nothing
   ORMX$ = TRIM$(Str$(ORMA% + 1))
   While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
   NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
   '
   CoComCor$ = NRECIB$
   CoComLar$ = NRECIB$
   NumeClie = 0
   NumeProv% = 0
   FechEmi = CVDAT(FEDEP%)
   '
   Referen$ = REFE$
   FeContab = CVDAT(FEDEP%)
   FeAcredi = CVDAT(FEDEP%)
   NumeTal = 1
   NumeCom = Val(Mid$(NRECIB$, 4))
   Observa$ = "Transferencia de Fondos"
   '
   CUDEB1% = CMC% 'ver si esto esta bien
   CuentaCo% = CUDEB1%
   TiMovi = TiMovi$          ' TRANSFERENCIA DE FONDOS
   OpMovi = ""                  '
   VaMovi = VARIMOVIX$                  '
   Descrip = "TRANSFERENCIA MANUAL REALIZADA POR USUARIO: " & US1%
   'ImpoDeb# = ITRAN#
   StatInit = 0

   Call ABRECABAN
25 On Error Resume Next
   Dim CAJABANC As ADODB.Recordset
   Set CAJABANC = New ADODB.Recordset
   strsql = "SELECT * FROM CAJABANC where numecomp < 1 "
   CAJABANC.Open strsql, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      MsgBox Err.Description
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   With CAJABANC
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CodiEmpr = CodiEmp%
     On Error GoTo 0
      !CodiCom_Cor = Left$(CoComCor$, 12)
      !CODICOM_LAR = Left$(CoComLar$, 24)
      !REFECOM = Left$(Referen$, 48)
      !FECHCONTA = SafeDate(FeContab)
      !fechacred = SafeDate(FeAcredi)
      !CUECONTAI = CuentaCo%
      !TipoMovim = Left$(TiMovi$, 8)
      !DescriMov = Left$(Descrip$, 128)
      !FECHEMISI = FechEmi
      !CodiEmpr = CodiEmp%
      !CodiSucu = CodiSuc%
      !NumeTalo = NumeTal
      !NUMECOMP = NumeCom
      !Status = StatInit
        On Error Resume Next
      !Mone_Emis = MONEMI%
      !ValAbsComp = Abs(ImpoDeb#)
      If ImpoDeb# > 0 Then
          !IDEBECOMP = 0
          !IHABECOMP = ImpoDeb#
        Else
          !IDEBECOMP = Abs(ImpoDeb#)
          !IHABECOMP = 0
      End If
      !NROCAI = TRIM$(Left$(NumCai$, 24))
      !OBSERGEN = Left$(Observa$, 64)
      !FeHoReal = Now()
      !NuAsiDiar = ""
      '
      !CUECONTAI = CuentaCo%
      USNU% = US1%
      !USERNUM = USNU%
      '
     .Update
   End With
   CAJABANC.Close
   Set CAJABANC = Nothing
   '
   CuentaCo% = CueConta%
   Descrip = "TRANSFERENCIA MANUAL REALIZADA POR USUARIO: " & US1%
   ImpoDeb# = (-1) * ImpoDeb#
   If YAVOLVIO% < 1 Then
     YAVOLVIO% = 1
     TiMovi = ""          ' TRANSFERENCIA DE FONDOS
     VaMovi = ""                  '
     GoTo 25
   End If
   Call COMUNI(TiMovi$ & " Completa y Verificada")
   '
99 Screen.MousePointer = 1
   Exit Sub
   
End Sub

Private Sub Command19_Click()
      CAMPOS$ = "CÓDIGO/F12;MOTIVO/A48"
      Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
      Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
      FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"
    
      Dim obj As New RECORXET
      Set RS1 = obj.RS_MOTIVOS()
      Call Carga_MSF_Recordset(RS1, CAMPOS$, FRMZELECHO.msf2)
      Set obj = Nothing
      RS1.Close
      Set RS1 = Nothing
      '
    
      FRMZELECHO.Width = 8500
      FRMZELECHO.Show 1
      Text20(0) = RESP_ZELE_VECT(2)

End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call CARGA_GRILLA_CAJA
   
   Command2.Enabled = True
End Sub

Private Sub Command25_Click()
   Command25.Enabled = False
   'valido
   If MSF.Rows <= 1 Then
      MsgBox "GRILLA DE IMPRESIÓN VACIA"
      GoTo 99
   End If
   '
   FILA& = MSF.Rows
   'LE AGREGO LOS TOTALES A LA GRILLA PARA QUE SE VEA EN LA IMPRESION
    MSF.Rows = FILA& + 6
    MSF.TextMatrix(FILA& + 2, 5) = "Ingreso:": MSF.TextMatrix(FILA& + 2, 6) = Text15(0)
    MSF.TextMatrix(FILA& + 3, 5) = "Retiro:":  MSF.TextMatrix(FILA& + 3, 7) = Text15(1)
    MSF.TextMatrix(FILA& + 4, 5) = "Resultado:":  MSF.TextMatrix(FILA& + 4, 7) = Text15(2)

   Set GRILLA_IMPRESION = MSF
   '
   PREIMPRE.txtEncabezado2.Text = ""
   '
   ENCABEZADO1$ = "Reporte de Movimientos de Caja"
   
   PREIMPRE.txtEncabezado1.Text = ENCABEZADO1$
   PREIMPRE.txtPiePagina = "Fin de Reporte"
   PREIMPRE.Combo1.Text = "Vertical"
   PREIMPRE.Show 1
   '
   'ESTO ES PARA VOLVER A CORTAR LA GRILLA DONDE ESTABA.
   MSF.Rows = FILA&
99 Command25.Enabled = True

End Sub



Private Sub Command3_Click()
    Command3.Enabled = False
    '
    Call CIERRARCH("!MECOBRA")
    Screen.MousePointer = 1

    Call SELECHO("MECOBRA")
    If XVALO(VALACT1$("MECOBRA")) > 0 Then
       Text3 = VALACT1$("MECOBRA")
       Text1 = VALACT2$("MECOBRA")
    Else
       Text3 = ""
       Text1 = ""
    End If
    '
    Call Command2_Click
99  Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    If DERACCE%("CAMBUSUAR", "VISUALIZAR MOVIMIENTOS DE CAJA DE OTROS USUARIOS") < 2 Then GoTo 99
    Screen.MousePointer = 11
    Call BLOQARCH("FLEXUSR")
    FUSER$ = LOADFILE$("USER.DAT")
    JJI& = 0: KKI% = 0
    For i% = 1 To Len(FUSER$) Step 64
       F1$ = Mid$(FUSER$, i%, 64)
       If Asc(F1$) = 0 Then GoTo 9
       NOM$ = Left$(F1$, 30)
       If Left$(NOM$, 1) = "@" Then NOM$ = DESENCRI$(Mid$(F1$, 2, 29))
       '
       KKI% = KKI% + 1
       X$ = REGBLAN$("FLEXUSR")
       Call REPLA$(X$, MKI$(KKI%), 1, 2)
       Call REPLA$(X$, NOM$, 3, 30)
       JJI& = JJI& + 1
       Call GRAREG("FLEXUSR", X$, JJI&)
    Next i%
    '
9   Call SETULTREG("FLEXUSR", JJI&)
10  Call CIERRARCH("FLEXUSR")
    '
    Call SELECHO("FLEXUSR")
    Screen.MousePointer = 1
    If VALACT1$("FLEXUSR") = "" Then GoTo 99
    '
    USEL% = XVALO(TRIM$(VALACT1$("FLEXUSR")))
'    DEUSER$ = VALACT2$("FLEXUSR")
    Me.TXTNUMUSUARIO = USEL%
'    Me.TXTDESCRIUSUARIO = DEUSER$
99  Command4.Enabled = True
End Sub

Private Sub Command5_Click()
  TXTBUSCAR.Text = ""
  For i& = 1 To MSF.Rows - 1
     Call COLOREAR_GRILLA_SOLAFILA(MSF, vbWhite, i&, 0, 1)
  Next i&
End Sub

Private Sub Desde_Click()
    If KeyAscii = 13 Then
       Call Command2_Click
    End If

End Sub

Private Sub Form_Load()
'    Call APLICAR_SKIN(Me, App.Path & "\SKINS\SteelBlue.skn")
    Call ENCABEZA_GRILLA
    Desde.Value = Date
    Hasta.Value = Date
    'Call CARGA_GRILLA_CAJA
    USNX% = USNBR% Mod 100
    Me.TXTNUMUSUARIO = USNX%
    Me.txtgastos = TRIM$(CONTROL$("CEMEDESC", "CUEGASTO"))
    Me.txtingreso = TRIM$(CONTROL$("CEMEDESC", "CUENTAIN"))
'    Call CARGAVECTOR_MES
    '
End Sub


Sub CARGA_GRILLA_CAJA()
   '
   MSF.Rows = 1
' Call LIMPIASUBTOTALES
   '
   CMC% = PrimerCtaMedioCobranza%
   If CMC% < 1 Then
     Call MENSERR(24, "Falta Cargar Medio de Cobranza Efectivo (Recuerde Agregar Esta en Primer Lugar")
     GoTo 10
   End If
   
   Screen.MousePointer = 11
   DES% = CVINT(Desde.Value)
   HAS% = CVINT(Hasta.Value)
   DESDEX = FETEXCOR1$(DES%)
   HASTAX = FETEXCOR1$(HAS%)
   FILA& = 0
   'CARGA EL ENCABEZADO DEL TRANSPORTE
   SQLX$ = "select * from cajabanc where "
   CONDI$ = "(tipomovim = 'chreci' or TipoMovim = 'INGRCAJA' OR TipoMovim = 'SALECAJA' OR TipoMovim = 'TRAFON')"
   CONDI$ = CONDI$ + "  AND FECHEMISI < '" & DESDEX & "'"
   If XVALO(TXTNUMUSUARIO) > 0 Then CONDI$ = CONDI$ + "  AND USERNUM = " & XVALO(Me.TXTNUMUSUARIO)

   If XVALO(Text4) > 0 Then CONDI$ = CONDI$ + "  AND CUECONTAI = " & XVALO(Text4) ' SI TIENE FILTRO POR CUENTA CONTABLE
   If CantRegistros&("CAJABANC", CONDI$, "NOMESS") < 1 Then GoTo 10
   SQLX$ = "SELECT SUM(idebecomp) AS DEBEX, SUM(ihabecomp)AS HABEX"
   SQLX$ = SQLX$ + " FROM CAJABANC WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE " + CONDI$
'   SQLX$ = SQLX$ + " group by CUENTA_CONT"

   Set RECITMP0 = FLEXCONN.Execute(SQLX$)
   With RECITMP0
      Do While Not .EOF
           RETIRO# = XVALO(!HABEX)
           SUPAGO# = XVALO(!DEBEX)
           SUINGREACUM# = SUINGREACUM# + SUPAGO#
           RETIROACUM# = RETIROACUM# + RETIRO#
           '
           FILA& = FILA& + 1
           TXTCANTIFILAS = FILA&: TXTCANTIFILAS.Refresh
           MSF.Rows = FILA& + 1
           MSF.TextMatrix(FILA&, 2) = "-------------------"
           MSF.TextMatrix(FILA&, 3) = "-------------------"
           MSF.TextMatrix(FILA&, 4) = "-------------------"
           MSF.TextMatrix(FILA&, 5) = "-------------------"
           MSF.TextMatrix(FILA&, 6) = "Transporte................"
           'MSF.TextMatrix(FILA&, 6) = "-------------------"
           MSF.TextMatrix(FILA&, 7) = TRIM$(FORMATNUM$(SUPAGO#, "F12.2"))
           MSF.TextMatrix(FILA&, 8) = TRIM$(FORMATNUM$(RETIRO#, "F12.2"))
           
           
           .MoveNext
      Loop
    End With
    RECITMP0.Close
    Set RECITMP0 = Nothing
    FILA& = FILA& + 1
    '
    'BLANQEO POR QUE PIDIERON QUE SOLO ACUMULE EL TOTAL DEL RANGO SELECCIONADO.
'    SUINGREACUM# = 0
'    RETIROACUM# = 0

10 CONDI$ = "(tipomovim = 'chreci' or TipoMovim = 'INGRCAJA' OR TipoMovim = 'SALECAJA' OR TipoMovim = 'TRAFON')"
   CONDI$ = CONDI$ + " AND FECHEMISI >= '" & DESDEX & "' AND FECHEMISI <= '" & HASTAX & "'"
   If XVALO(TXTNUMUSUARIO) > 0 Then CONDI$ = CONDI$ + "  AND USERNUM = " & XVALO(Me.TXTNUMUSUARIO)
   If XVALO(Text4) > 0 Then CONDI$ = CONDI$ + "  AND CUECONTAI = " & XVALO(Text4)
   '
   If CantRegistros&("CAJABANC", CONDI$, "NOMESS") < 1 Then GoTo 90
   '
   'MSF.Redraw = False
   '
   Dim CLS As New clientes
   CLS.cargarvector
   SQLX$ = "SELECT IDEBECOMP,TipoMovim, IHABECOMP,FECHEMISI,DescriMov,NUCLIEN,CODICOM_LAR,REFECOM,CUECONTAI,USERNUM  FROM CAJABANC WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " + CONDI$
   SQLX$ = SQLX$ + " ORDER BY FeHoReal"
   Set RECITMP = FLEXCONN.Execute(SQLX$)
   With RECITMP
      Do While Not .EOF
           RETIRO# = XVALO(!IHABECOMP)
           SUINGRE# = XVALO(!IDEBECOMP)
           FEMI% = CVINT(!FECHEMISI)
           REFE$ = SAFETEXT$(!DescriMov)
           If SAFETEXT$(!TipoMovim) = "INGRCAJA" Or SAFETEXT$(!TipoMovim) = "SALECAJA" Then REFE$ = SAFETEXT$(!REFECOM)
           LGJO& = XVALO(!NUCLIEN)
           DOCUM$ = SAFETEXT$(!CODICOM_LAR)
           SUINGREACUM# = SUINGREACUM# + SUINGRE#
           RETIROACUM# = RETIROACUM# + RETIRO#
           US% = XVALO(!USERNUM)
           '
           FILA& = FILA& + 1
           TXTCANTIFILAS = FILA&: TXTCANTIFILAS.Refresh
           MSF.Rows = FILA& + 1
           MSF.TextMatrix(FILA&, 1) = NOMBRE_USUARIO$(US%)
           MSF.TextMatrix(FILA&, 2) = FECHATEX$(FEMI%)
           MSF.TextMatrix(FILA&, 3) = DOCUM$
           MSF.TextMatrix(FILA&, 4) = REFE$
           MSF.TextMatrix(FILA&, 5) = TRIM$(Str$(LGJO&))
           MSF.TextMatrix(FILA&, 6) = CLS.nombre(LGJO&)
           MSF.TextMatrix(FILA&, 7) = TRIM$(FORMATNUM$(SUINGRE#, "F12.2"))
           MSF.TextMatrix(FILA&, 8) = TRIM$(FORMATNUM$(RETIRO#, "F12.2"))
           MSF.TextMatrix(FILA&, 9) = TRIM$(!CUECONTAI)
           .MoveNext
      Loop
    End With
    RECITMP.Close
    Set RECITMP = Nothing
    '
90  '

    Call SUMATOTAL(SUINGREACUM#, RETIROACUM#)
    '
    'ACA CALCULAR EL TOTAL EFECTIVO (TRANSPORTE Y FECHAS SELECCIONADAS
   CONDI$ = "(tipomovim = 'chreci' or TipoMovim = 'INGRCAJA' OR TipoMovim = 'SALECAJA' OR TipoMovim = 'TRAFON')"
   CONDI$ = CONDI$ + " AND  FECHEMISI <= '" & HASTAX & "'"
   If XVALO(TXTNUMUSUARIO) > 0 Then CONDI$ = CONDI$ + "  AND USERNUM = " & XVALO(Me.TXTNUMUSUARIO)
   CONDI$ = CONDI$ + "  AND CUECONTAI = " & CMC% 'CTA DE EFECTIVO
   '
   Text15(4) = FORMATNUM$(XVALO(VALOBADA("CAJABANC", "Sum(IDEBECOMP - IHABECOMP)", CONDI$, "NOMESS")), "F12.2")
   Text15(5) = FORMATNUM$(XVALO(Text15(2)) - XVALO(Text15(4)), "F12.2")
   '
    '
    '
    Screen.MousePointer = 1

End Sub
Sub SUMATOTAL(SUPAGOACUM#, RETIROACUM#)
    Text15(0) = FORMATNUM$(SUPAGOACUM#, "F14.2")
    Text15(1) = FORMATNUM$(RETIROACUM#, "F14.2")
    Text15(2) = FORMATNUM$(SUPAGOACUM# - RETIROACUM#, "F14.2")
    
End Sub
Sub ENCABEZA_GRILLA()
    Call CARGA_ENCABEZADO(MSF, "Usuario/A18;Fecha Em./D8;Documento/A18;Referencia/A36;Lgjo/f7;Nombre y Apellido/A24;Ingreso/F12.2;Retiro/F12.2;C.C/F2", Me)
    '
'    MSF.MergeCells = flexMergeRestrictColumns
'    MSF.MergeCol(5) = True

End Sub

Sub LIMPIASUBTOTALES()

   For i% = 0 To 2
      Text15(i%) = ""
   Next i%
End Sub


Private Sub mnuTRansfxUsuario_Click()
     
   FEX = HOY(HOS$): Call REPLA(VDEF$, MKI$(FEX), 1, 2)
   NUMUS% = USNBR% Mod 100
   Call REPLA(VDEF$, SAFETEXT$(NUMUS%), 15, 4)
   '
   OBJETOX$ = "TRANSFE-USER"
   If NUMUS% = 1 Then OBJETOX$ = "Transferencias D"
10 DTP$ = OBJPLA$(OBJETOX$, VDEF$)
   If TRIM$(DTP$) = "" Then Exit Sub
   VDEF$ = DTP$
   
   FEX = CVI(Left(VDEF$, 2))
   ImpoDeb# = CVD(Mid$(VDEF$, 3, 8))
   US1% = XVALO(Mid$(VDEF$, 11, 4))
   US2% = XVALO(Mid$(VDEF$, 15, 4))
   REFERENCIA$ = Mid$(VDEF$, 19, 20)
   '
   
   If FEX < 1 Then
     Call MENSERR(24, "Fecha no Válida")
     GoTo 10
   End If
   '
   If ImpoDeb# < 0.005 Then
     Call MENSERR(24, "Falta Importe")
     GoTo 10
   End If
   '
   If US1% < 1 Then
     Call MENSERR(24, "Falta Usuario Desde")
     GoTo 10
   End If
   If US2% < 1 Then
     Call MENSERR(24, "Falta Usuario Hasta")
     GoTo 10
   End If
   
   If TRIM$(REFERENCIA) = "" Then
     Call MENSERR(24, "Falta Referencia")
     GoTo 10
   End If
   '
   CMC% = PrimerCtaMedioCobranza%
   If CMC% < 1 Then
     Call MENSERR(24, "Falta Cargar Medio de Cobranza Efectivo (Recuerde Agregar Esta en Primer Lugar")
     GoTo 10
   End If
   '
   'ACA CALCULAR EL TOTAL EFECTIVO (TRANSPORTE Y FECHAS SELECCIONADAS
   HASTAX = FETEXCOR1$(HOY(HO$))
   CONDI$ = "(tipomovim = 'chreci' or TipoMovim = 'INGRCAJA' OR TipoMovim = 'SALECAJA' OR TipoMovim = 'TRAFON')"
   CONDI$ = CONDI$ + " AND  FECHEMISI <= '" & HASTAX & "'"
   CONDI$ = CONDI$ + "  AND USERNUM = " & US1%
   CONDI$ = CONDI$ + "  AND CUECONTAI = " & CMC% 'CTA DE EFECTIVO
   '
   SALDO_EF_ENCAJA = ROUND(XVALO(VALOBADA("CAJABANC", "Sum(IDEBECOMP - IHABECOMP)", CONDI$, "NOMESS")), 2)

   If ImpoDeb# > SALDO_EF_ENCAJA Then
      Call COMUNI("Imposible Registrar Retiro Mayor al Total de Efectivo de la Caja del Usuario Seleccionado\Saldo de Usuario: " & NOMBRE_USUARIO(US1%) & " en Efectivo : " & SALDO_EF_ENCAJA)
      GoTo 10
   End If

   REFE$ = TRIM$(Mid$(OBX$, 35, 40)) + " " + FECHATEX$(FEX)
   '
   FEDEP% = FEX
   Screen.MousePointer = 11
   FETX$ = FECHATEX$(HOY(HO$))
   NRECIB$ = "TF-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
   NRECIBX$ = NRECIB$ + "__"
   ORMA% = (-1)
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE CodiCom_Cor LIKE '" & NRECIBX$ & "'"
   Dim CAJABTEMP As ADODB.Recordset
   Set CAJABTEMP = READSET(SQLX$)

   On Error Resume Next
   With CAJABTEMP
30   'If .NoMatch = False Then
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
          If Val(Mid$(!CodiCom_Cor, 11)) > ORMA% Then
            ORMA% = Val(Mid$(!CodiCom_Cor, 11))
          End If
          .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   Set CAJABTEMP = Nothing
   ORMX$ = TRIM$(Str$(ORMA% + 1))
   While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
   NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
   '
   CoComCor$ = NRECIB$
   CoComLar$ = NRECIB$
   NumeClie = 0
   NumeProv% = 0
   FechEmi = CVDAT(FEDEP%)
   '
   Referen$ = REFE$
   FeContab = CVDAT(FEDEP%)
   FeAcredi = CVDAT(FEDEP%)
   NumeTal = Val(Mid$(NRECIB$, 11, 2))
   NumeCom = Val(Mid$(NRECIB$, 4, 6))
   Observa$ = "Transferencia de Fondos"
   '
   CUDEB1% = CMC% 'cuenta por defecto efectivo
   CuentaCo% = CUDEB1%
   TiMovi = "TRAFON"            ' TRANSFERENCIA DE FONDOS
   OpMovi = ""                  '
   VaMovi = ""                  '
   Descrip = REFERENCIA$ + "(Destino: " & NOMBRE_USUARIO(US2%) & ")"
   'ImpoDeb# = ITRAN#
   StatInit = 0

   Call ABRECABAN
25 On Error Resume Next
   Dim CAJABANC As ADODB.Recordset
   Set CAJABANC = New ADODB.Recordset

   
   strsql = "SELECT * FROM CAJABANC "
   CAJABANC.Open strsql, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      MsgBox Err.Description
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   With CAJABANC
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CodiEmpr = CodiEmp%
     On Error GoTo 0
      !CodiCom_Cor = Left$(CoComCor$, 12)
      !CODICOM_LAR = Left$(CoComLar$, 24)
      !REFECOM = Left$(Referen$, 48)
      !FECHCONTA = SafeDate(FeContab)
      !fechacred = SafeDate(FeAcredi)
      !CUECONTAI = CuentaCo%
      !TipoMovim = Left$(TiMovi$, 8)
      !DescriMov = Left$(Descrip$, 128)
      !FECHEMISI = FechEmi
      !CodiEmpr = CodiEmp%
      !CodiSucu = CodiSuc%
      !NumeTalo = NumeTal
      !NUMECOMP = NumeCom
      !Status = StatInit
        On Error Resume Next
      !Mone_Emis = MONEMI%
      !ValAbsComp = Abs(ImpoDeb#)
      If ImpoDeb# > 0 Then
          !IDEBECOMP = 0
          !IHABECOMP = ImpoDeb#
        Else
          !IDEBECOMP = Abs(ImpoDeb#)
          !IHABECOMP = 0
      End If
      !NROCAI = TRIM$(Left$(NumCai$, 24))
      !OBSERGEN = Left$(Observa$, 64)
      !FeHoReal = Now()
      !NuAsiDiar = ""
      '
      USNU% = US1%
      !USERNUM = USNU%
      '
     .Update
   End With
   CAJABANC.Close
   Set CAJABANC = Nothing
   '
   
   If YAVOLVIO% < 1 Then
     YAVOLVIO% = 1
     CuentaCo% = CUDEB1%
     ImpoDeb# = (-1) * ImpoDeb#

     Descrip = REFERENCIA$ + "(Origen: " & NOMBRE_USUARIO(US1%) & ")"
     US1% = US2%
     GoTo 25
   End If
   Call COMUNI("Transferencia Completa y Verificada")
   '
   Screen.MousePointer = 1
End Sub





Private Sub MSF_DblClick()
        REG& = MSF.MouseRow
        If REG& < 1 Then Exit Sub
        
        TIDOCUM$ = MSF.TextMatrix(REG&, 2)
        If InStr(TIDOCUM$, "RB-") Then TIDOCUM$ = "Recibo"
        FE% = FECHANUM(MSF.TextMatrix(REG&, 1))
        NUDOCU$ = " " + CStr(XVALO(Mid$(MSF.TextMatrix(REG&, 2), 11, 8))) + " "
        NUDOCU1$ = Str$(XVALO(NUDOCU$))
        NUDOCU2$ = NUDOCU$
        
        TIPOCOM% = 3
        CONDI$ = "NUMECOMP= " & XVALO(NUDOCU$) & " AND fechemisi = '" & FETEXCOR1$(FE%) & "' and TIPOMOVIM = 'RCOB'"
        
        FE% = CVINT(VALOBADA("CAJABANC", "FeHoReal", CONDI$, "NOMESS"))
        NC = XVALO(VALOBADA("CAJABANC", "NUCLIEN", CONDI$, "NOMESS"))
        NOMB$ = rasocli$(NC)
        
        
        If TIDOCUM$ = "Recibo" Then VARCONTROL$ = NOMB$
        DOCUNU& = DOCID&(TIDOCUM$, NUDOCU1$, NUDOCU2$, VARCONTROL$)
        If DOCUNU& < 1 Then
         Call COMUNI("Documento No Encontrado")
        Else
          OPPXX$ = ""
        If NUDOCU$ = "00000032" Then OPPXX$ = "RECUP"
          Call MUESTRADOC(DOCUNU&, OPPXX$, TIDOCUM$, NUDOCU1$, NUDOCU2$)
        End If

End Sub


Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call PINTATEXT(Text20(Index))
    End If
End Sub

Private Sub Text2_LOSTFOCUS(Index As Integer)
    Text2(Index) = FORMATNUM$(XVALO(Text2(Index)), "F14.2")
End Sub

Private Sub Text20_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call PINTATEXT(Text2(1))
    End If
End Sub

Private Sub Text3_Change()
   TIPOVAL% = XVALO(Text3)
   RFF$ = RECFILT$("MECOBRA", 1, MKI$(TIPOVAL%))
   If Len(RFF$) >= 4 Then
     REME& = CVS(Mid$(RFF$, 1, 4))
     If REME& > 0 Then
       CUCU$ = Mid$(REGLEIDO$("MECOBRA", REME&), 45, 12)
       CMC% = CUECOINT%(TRIM$(CUCU$))
       If CMC% < 1 Then CMC% = CMC0%
     End If
   End If
   Text4 = CMC%
   '
'   CUEMECOB% = CMC%

End Sub

Private Sub Text3_DblClick()
    Text3 = ""
    Text1 = ""
End Sub


Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If TRIM$(TXTBUSCAR.Text) <> "" Then
         BUSCO$ = TRIM$(TXTBUSCAR.Text)
         For i& = 1 To MSF.Rows - 1
            Call COLOREAR_GRILLA_SOLAFILA(MSF, vbWhite, i&, 0, 1)
            For j& = 1 To MSF.Cols - 1
                If (InStr(UCase$(MSF.TextMatrix(i&, j&)), UCase$(BUSCO$))) Then
                   Call COLOREAR_GRILLA_SOLAFILA(MSF, vbYellow, i&, 0, 1, 0)
                   Exit For
                End If
            Next j&
         Next i&
      End If
   End If
End Sub


Private Sub txtgastos_Change()
 Me.Text16(4) = DENOCUE$(CUECOINT(Me.txtgastos))
End Sub

Private Sub txtgastos_DblClick()
           USNX% = USNBR% Mod 100
           If USNX% <> 1 Then
              Call COMUNI("Solo el Usuario Supervisor Puede Realizar Cambios en las Cuentas Contables")
              Exit Sub
           End If
        
           CuentaContable$ = TRIM$(CONTROL$("CEMEDESC", "CUEGASTO"))
    
           Call SELECHO("ECUECON")
            
           CueConta$ = TRIM$(VALACT1$("ECUECON"))
           If CueConta$ <> "" Then
                Me.txtgastos = CueConta$
                Call GRACONTROL("CEMEDESC", "CUEGASTO", CueConta$)
                Call COMUNI("Se Ha Actualizdo la Cuenta Contable Para Retiros en Efectivo")
           End If
End Sub


Private Sub txtingreso_Change()
    Me.Text16(5) = DENOCUE$(CUECOINT(Me.txtingreso))
End Sub


Private Sub txtingreso_DblClick()
            CuentaContable$ = TRIM$(CONTROL$("CEMEDESC", "CUENTAIN"))
            Call SELECHO("ECUECON")
            CueConta$ = TRIM$(VALACT1$("ECUECON"))
            If CueConta$ <> "" Then
                Call GRACONTROL("CEMEDESC", "CUENTAIN", CueConta$)
                Me.txtingreso = CueConta$
                Call COMUNI("Se Ha Actualizdo la Cuenta Contable Para Ingresos de Efectivo")
            End If

End Sub


Private Sub TXTNUMUSUARIO_Change()
   TXTDESCRIUSUARIO = ""
   US% = XVALO(Me.TXTNUMUSUARIO)
   If US% > 0 Then
    RFF$ = RECFILT$("FLEXUSR", 1, MKI$(US%)) ' SE DEBE CARGAR LA CUENTA DE EFECTIVO PRIMERA EN LOS  MEDIOS DE COBRANZA
    If Len(RFF$) >= 4 Then
      REME& = CVS(Mid$(RFF$, 1, 4))
      If REME& > 0 Then
        NAMEUSUARIO$ = Mid$(REGLEIDO$("FLEXUSR", REME&), 3, 30)
      End If
    End If
   End If
   Me.TXTDESCRIUSUARIO = NAMEUSUARIO$
   Call Command2_Click
End Sub

Private Sub TXTNUMUSUARIO_DblClick()

       If DERACCE%("CAMBUSUAR", "VISUALIZAR MOVIMIENTOS DE CAJA DE OTROS USUARIOS") < 2 Then GoTo 99
       TXTNUMUSUARIO.Text = ""
       
99

End Sub

Function PrimerCtaMedioCobranza%()
   RFF$ = RECFILT$("MECOBRA", 1, MKI$(1)) ' SE DEBE CARGAR LA CUENTA DE EFECTIVO PRIMERA EN LOS  MEDIOS DE COBRANZA
   If Len(RFF$) >= 4 Then
     REME& = CVS(Mid$(RFF$, 1, 4))
     If REME& > 0 Then
       CUCU$ = Mid$(REGLEIDO$("MECOBRA", REME&), 45, 12)
       CMC% = CUECOINT%(TRIM$(CUCU$))
     End If
   End If
   
   PrimerCtaMedioCobranza% = CMC%
End Function
