VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form INGTEL 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe de Asignación de Telas para Bolsas "
   ClientHeight    =   9135
   ClientLeft      =   1050
   ClientTop       =   1155
   ClientWidth     =   15495
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   15495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   65
      Top             =   5220
      Width           =   15495
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "<<<<<<    Lotes Asignados a La Orden de Fabricación Activa    >>>>>>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   1
         Left            =   4000
         TabIndex        =   66
         Top             =   120
         Width           =   8970
      End
   End
   Begin VB.Frame Frame3 
      Height          =   495
      Left            =   0
      TabIndex        =   63
      Top             =   1320
      Width           =   15495
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "<<<<<<     Asignación De telas para Bolsas     >>>>>>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   0
         Left            =   4000
         TabIndex        =   64
         Top             =   120
         Width           =   9810
      End
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
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
      Height          =   2250
      Index           =   1
      Left            =   5280
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "INGTEL.frx":0000
      Top             =   2280
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Materiales Transferidos  a Depósito Para Bolsas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      Index           =   1
      Left            =   0
      TabIndex        =   31
      Top             =   1800
      Width           =   15495
      Begin VB.CommandButton Command40 
         Appearance      =   0  'Flat
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
         Left            =   12390
         TabIndex        =   68
         Top             =   1320
         Width           =   200
      End
      Begin VB.TextBox TXTNUMOPERARIO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11760
         TabIndex        =   67
         Top             =   1320
         Width           =   600
      End
      Begin VB.TextBox TXTELEMSELEC 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   960
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox TXTKGS 
         Alignment       =   1  'Right Justify
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
         Left            =   12360
         Locked          =   -1  'True
         TabIndex        =   60
         Text            =   " "
         Top             =   3000
         Width           =   780
      End
      Begin VB.TextBox TXTCORTESINFO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11205
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   " "
         Top             =   2160
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox TXTCORTXBSON 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   14400
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   " "
         Top             =   2160
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox TXTCODINT 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   12240
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   2160
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox TXTLOTE2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   14160
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   " "
         Top             =   1800
         Width           =   1020
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Grabar"
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
         Left            =   13920
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "O.K."
         Top             =   3000
         Width           =   1245
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Registrar Peso"
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
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Lee el Peso de La Balanza"
         Top             =   3000
         Width           =   1485
      End
      Begin VB.TextBox TXTCORTE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   14400
         Locked          =   -1  'True
         TabIndex        =   49
         Text            =   " "
         Top             =   2640
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox TXTCANTCORTE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   13005
         TabIndex        =   47
         Text            =   " "
         Top             =   2640
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox TXTNECESIDADCORTES 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11205
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   " "
         Top             =   1800
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox TXTRESTAN 
         Alignment       =   1  'Right Justify
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
         Left            =   11205
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   " "
         Top             =   2520
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
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
         Left            =   12390
         TabIndex        =   40
         Top             =   960
         Visible         =   0   'False
         Width           =   200
      End
      Begin VB.TextBox TXTNOP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11760
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   960
         Width           =   600
      End
      Begin VB.TextBox TXTCANPROY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   14235
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   " "
         Top             =   240
         Width           =   955
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   7.5
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   13080
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   175
      End
      Begin MSFlexGridLib.MSFlexGrid MSF0 
         Height          =   3015
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "Doble Clic Para Asignar a Orden de Fabricación/Clic Botón Derecho Eliminar Registro"
         Top             =   360
         Width           =   9765
         _ExtentX        =   17224
         _ExtentY        =   5318
         _Version        =   393216
         Rows            =   26
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Operario:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   10800
         TabIndex        =   70
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label LBLNOMBREOPERARIO 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   12600
         TabIndex        =   69
         Top             =   1320
         Width           =   2590
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kgs:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   11760
         TabIndex        =   61
         Top             =   3120
         Width           =   405
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cortes Informados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   10200
         TabIndex        =   59
         Top             =   2160
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cortes x Bolsa: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   13680
         TabIndex        =   56
         Top             =   2115
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   13215
         TabIndex        =   54
         Top             =   1800
         Width           =   885
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Restan:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   10200
         TabIndex        =   50
         Top             =   2640
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de Cortes:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   12000
         TabIndex        =   48
         Top             =   2595
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Corte m.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   13440
         TabIndex        =   46
         Top             =   2640
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label LABEL12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Necesidad de Cortes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   5
         Left            =   10200
         TabIndex        =   43
         Top             =   1755
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Operación:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10845
         TabIndex        =   42
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label DESCRIOP 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   12600
         TabIndex        =   41
         Top             =   960
         Width           =   2590
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.Proy.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   13260
         TabIndex        =   38
         Top             =   315
         Width           =   885
      End
      Begin VB.Label LBLDESCRIOF 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   10305
         TabIndex        =   36
         Top             =   600
         Width           =   4885
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de O/F: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10200
         TabIndex        =   35
         Top             =   315
         Width           =   855
      End
      Begin VB.Label LBLOF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   11070
         TabIndex        =   34
         ToolTipText     =   "Seleccione Orden de Fabricación Asociada"
         Top             =   240
         Width           =   1890
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Transferir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   14040
      TabIndex        =   14
      Top             =   840
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1305
      Left            =   360
      TabIndex        =   4
      Top             =   0
      Width           =   15495
      Begin VB.TextBox txtcodintelem 
         Alignment       =   1  'Right Justify
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
         Left            =   11895
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Código Interno"
         Top             =   915
         Width           =   675
      End
      Begin VB.CommandButton Command4 
         Caption         =   "+"
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
         Left            =   7200
         Picture         =   "INGTEL.frx":00CC
         TabIndex        =   27
         Top             =   900
         Width           =   300
      End
      Begin VB.TextBox TXTPESOORIG 
         Alignment       =   1  'Right Justify
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
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   915
         Width           =   1365
      End
      Begin VB.TextBox TXTMETROS 
         Alignment       =   1  'Right Justify
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
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   915
         Width           =   1125
      End
      Begin VB.CommandButton command15 
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
         Height          =   330
         Left            =   1710
         Picture         =   "INGTEL.frx":03D6
         TabIndex        =   15
         Top             =   920
         Width           =   175
      End
      Begin VB.TextBox TXTPESO 
         Alignment       =   1  'Right Justify
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
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   520
         Width           =   1125
      End
      Begin VB.TextBox TXTLOTE 
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
         Left            =   11175
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   520
         Width           =   1395
      End
      Begin VB.TextBox TXTDESCRIELEM 
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
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   150
         Width           =   4170
      End
      Begin VB.TextBox TXTCODIGOELEM 
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
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   150
         Width           =   1530
      End
      Begin VB.TextBox TXTLECTURA 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   0
         Top             =   120
         Width           =   2355
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CI:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   11595
         TabIndex        =   29
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Peso Original:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   4375
         TabIndex        =   26
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Height          =   330
         Left            =   1920
         TabIndex        =   23
         Top             =   555
         Width           =   2325
      End
      Begin VB.Label lbldepogeneral 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   1200
         TabIndex        =   22
         Top             =   555
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep. General"
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
         Left            =   -740
         TabIndex        =   21
         Top             =   600
         Width           =   1890
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Metros:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   7650
         TabIndex        =   20
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label6 
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
         Height          =   330
         Left            =   1920
         TabIndex        =   18
         Top             =   915
         Width           =   2325
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep. Corte  de Eslingas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -120
         TabIndex        =   17
         Top             =   840
         Width           =   1170
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   1200
         TabIndex        =   16
         Top             =   915
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Peso Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   7230
         TabIndex        =   13
         Top             =   525
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   10680
         TabIndex        =   11
         Top             =   525
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   7245
         TabIndex        =   9
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código Elemento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   3480
         TabIndex        =   7
         Top             =   150
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   ">>>>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   156
         Width           =   612
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1440
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "INGTEL.frx":06E0
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   3240
      OleObjectBlob   =   "INGTEL.frx":0914
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   1410
      Left            =   0
      TabIndex        =   30
      Top             =   5715
      Width           =   15520
      _ExtentX        =   27384
      _ExtentY        =   2487
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   1650
      Left            =   0
      TabIndex        =   71
      Top             =   7560
      Width           =   15465
      _ExtentX        =   27279
      _ExtentY        =   2910
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   72
      Top             =   7100
      Width           =   15495
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "<<<<<<    Detalle de Lotes Asignados a La Orden de Fabricación Activa    >>>>>>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   2
         Left            =   4000
         TabIndex        =   73
         Top             =   120
         Width           =   8970
      End
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuedit 
      Caption         =   "Edición"
   End
   Begin VB.Menu mnuCons 
      Caption         =   "Consultas y Listados"
   End
   Begin VB.Menu mnuut 
      Caption         =   "Utilidades"
   End
   Begin VB.Menu mnuaccesodirecto 
      Caption         =   "Accesos Directos"
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuración"
      Begin VB.Menu mnuDepoCinESl 
         Caption         =   "Depósito de Cintilla y Eslingas y Bolsas"
      End
   End
End
Attribute VB_Name = "INGTEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub blanformu1()
        TXTLECTURA = ""
        TXTCODIGOELEM = ""
        TXTPESOORIG = ""
        TXTDESCRIELEM = ""
        TXTPESO = ""
        TXTMETROS = ""
        TXTLOTE = ""
        txtcodintelem = ""
        TXTKGS = ""
End Sub

Private Sub Command1_Click()
    LBLOF = OrdenFabricacionAbiertas$
End Sub

Private Sub mnuDepoCinESl_Click()
    ConfigDepo.Show 1
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    CI1% = XVALO(TXTELEMSELEC)
    If CI1% < 1 Then
      Call COMUNI("Código No Válido")
      GoTo 99
    End If
    
    Dim ctp As New Conteplast
    PESOX# = ctp.PesoLeidoBalanza(CI1%, 2)
    If PESOX# < 0 Then PESOX# = 0
    TXTKGS = FORMATNUM$(PESOX#, "f12.2")
99  Command11.Enabled = True

End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    FECHA% = HOY(HO$)
    CMOV$ = "TR"
    
    DEPOMATBOL% = XVALO(Me.Label13)

    If DEPOMATBOL% < 1 Then
       Call MENSERR(24, "Falta Depósito de Materiales para Bolsas")
       GoTo 99
    End If
    If XVALO(TXTELEMSELEC) < 1 Then
       Call MENSERR(24, "Código Elemento No Válido")
       GoTo 99
    End If
    If TRIM$(Me.TXTLOTE2) = "" Then
       Call MENSERR(24, "Lote No Válido")
       GoTo 99
    End If
    If XVALO(TXTCANTCORTE) <= 0 Then
       Call MENSERR(24, "Cantidad No Válida")
       GoTo 99
    End If
    If NUMEOPERACION% < 1 Then
       Call MENSERR(24, "Operación No Válida")
       GoTo 99
    End If
    '
'    If XVALO(TXTKGS) <= 0 Then
'       Call MENSERR(24, "Pesaje Informado No Válido")
'       GoTo 99
'    End If
    
'    LOTE$ = UCase$(TXTLOTE2)
'    coefi# = CoeficientePesoPorMetroSegunLote#(LOTE$)
'    If coefi# = 0 Then
'       Call MENSERR(24, "Falta Registrar Peso para Este Lote")
'       GoTo 99
'    End If

    CI1% = TXTELEMSELEC
    CantdeCortes = XVALO(TXTCANTCORTE)
    NCOMP& = ProNroComprobante(CMOV$)
    DOPRI$ = "TR." + TRIM$(str$(NCOMP&))
    DOSEC$ = TRIM$(LBLOF)
    NUMEOPERACION% = XVALO(TXTNOP)
    DESCRIDEPO$ = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPOMATBOL%)))
    REFE$ = "Tr.de " & TRIM$(DESCRIDEPO$)
    'SE TRANFIEREN METROS SEGUN EL KG REGISTRADOS
    CANTI = MetrosSegunPesoSegúnLote#(LOTE$, XVALO(Me.TXTKGS))
    Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC%, DEPOMATBOL%, (-1) * CANTI, , , , , , , , , , NUMEOPERACION%)   'GUARDO EL NUMERO DE OPERACION DE CADA ITEMS TRANSFERIDO
    
    DESCRIDEPO$ = ECOARCH$("TADEPOSI", Chr$(DEPOMATBOL%))
    REFE$ = "Tr.A " & TRIM$(DESCRIDEPO$)
    NCOMP& = ProNroComprobante(CMOV$)
    
    Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC%, DEPOMATBOL%, CANTI, , , , , , , , , , NUMEOPERACION%)  'GUARDO EL NUMERO DE OPERACION DE CADA ITEMS TRANSFERIDO
    Call ACSALOTE(CI1%, LOTE$, "NOMESS")

    Dim ctp As New Conteplast
    '
    NORFA$ = TRIM$(LBLOF)
    CIFAB% = CODIFAB%(NORFA$)
    NUMEOPERARIO% = XVALO(TXTNUMOPERARIO)
    CORTEENMETROS = XVALO(TXTCORTE)
    PESOKG = XVALO(TXTKGS)

    Call ctp.GrabarOperacion(NORFA$, CODIEMPR%, CIFAB%, NUMEOPERACION%, CantdeCortes, LOTE$, CI1%, XVALO(TXTCORTXBSON), NUMEOPERARIO%, CORTEENMETROS, PESOKG, RET_NEWORDEN%, Ret_NuevoIDGenerado&)
    Call ctp.CargaMatTranfXOrden(MSF1, LBLOF, DEPOMATBOL%, NUMEOPERACION%)
    
    Call ctp.TransfIngresada(MSF0, DEPOMATBOL%, "")

    Call ctp.CargaOperacionesInformadas(NORFA$, CIFAB%, MSF2, 0, "", 99)
    Set ctp = Nothing

'
'    CIELEM% = CI1% 'material
'    CANTI = CantdeCortes 'cantidad de cortes)
'    NOP% = NUMEOPERACION%
'    NROOF$ = NORFA$
'    NROPED& = NroPedidoSegunOf&(NORFA$)
'    NC = cliente_SEGUN_PEDIDO(NROPED&)
'    ORDENINFO = RET_NEWORDEN%
'    UBICACION$ = TRIM$(MSF2.TextMatrix(REG&, 10))
'    IDUBICACION = XVALO(MSF2.TextMatrix(REG&, 12))
'     DESCRIOPERACION$ = MSF2.TextMatrix(REG&, 7)
'    CODIGOINTFAB = CIFAB%
'     Call PRINTETIQPALLET(CODEXT$(CIELEM%), LOTE$, CANTI, NOP%, NC, NROPED&, NROOF$, ORDENINFO, UBICACION$, IDUBICACION, DESCRIOPERACION$, CODIGOINTFAB)
     '+++++++++++++++++++++++++++++++++++++++++++++++
    '++++++++++++++++++++++++++++++++++++++++++++++++++++
     'BUSCO EL REGISTRO QUE SE CREO segun CAINFANT%
     CODIGOEXT$ = CODEXT$(CI1%)
     For i& = 1 To MSF2.Rows - 1
        If XVALO(MSF2.TextMatrix(i&, 11)) = RET_NEWORDEN% Then
          If MSF2.TextMatrix(i&, 9) = TRIM$(CODIGOEXT$) Then
            If TRIM$(MSF2.TextMatrix(i&, 7)) = TRIM$(DESCRIOP) Then
               REG& = i&
               Exit For
            End If
          End If
        End If
     Next i&
     
     CIELEM% = CODINT(MSF2.TextMatrix(REG&, 9)) 'material
     LOTE$ = TRIM$(MSF2.TextMatrix(REG&, 3))
     CANTI = XVALO(MSF2.TextMatrix(REG&, 4)) 'cantidad de cortes)
     NOP% = XVALO(MSF2.TextMatrix(REG&, 6))
     NROPED& = NroPedidoSegunOf&(NORFA$)
     NC = cliente_SEGUN_PEDIDO(NROPED&)
     NROOF$ = NORFA$
     ORDENINFO = XVALO(MSF2.TextMatrix(REG&, 11))
     UBICACION$ = TRIM$(MSF2.TextMatrix(REG&, 10))
     IDUBICACION = XVALO(MSF2.TextMatrix(REG&, 12))
     DESCRIOPERACION$ = MSF2.TextMatrix(REG&, 7)
     CODIGOINTFAB = CIFAB%
     IDINFOFAB& = XVALO(MSF2.TextMatrix(REG&, 16))

     Call INFOCORTE.PRINTETIQPALLET(CODEXT$(CIELEM%), LOTE$, CANTI, NOP%, NC, NROPED&, NROOF$, ORDENINFO, UBICACION$, IDUBICACION, DESCRIOPERACION$, CODIGOINTFAB, IDINFOFAB&)
     '++++++++++++++++++++++++++++++++++++++++++++++
    Call BLANFORMU2
99 Screen.MousePointer = 1
   Command2.Enabled = True
   
End Sub
Sub BLANFORMU2()
    'LBLOF = ""
    TXTCANPROY = ""
    LBLDESCRIOF = ""
    TXTELEMSELEC = ""
    TXTNOP = ""
    DESCRIOP = ""
    LBLNOMBREOPERARIO = ""
    TXTNUMOPERARIO = ""
    TXTNECESIDADCORTES = ""
    TXTCORTESINFO = ""
    TXTRESTAN = ""
    TXTCANTCORTE = ""
    TXTCODINT = ""
    TXTKGS = ""
    TXTCORTXBSON = ""
    TXTCORTE = ""
End Sub
Private Sub Command3_Click()
    Command3.Enabled = False
    CI1% = CODINT(TXTCODIGOELEM)
    If CI1% < 1 Then
       Call COMUNI("Falta Seleccionar Código")
       GoTo 99
    End If
    DEPOALMACEN% = XVALO(Me.lbldepogeneral)
    If DEPOALMACEN% < 1 Then
      Call COMUNI("Falta Depósito General")
      GoTo 99
    End If
    
    DEPOMATBOL% = Me.Label13
    If DEPOMATBOL% < 1 Then
      Call COMUNI("Falta Depósito de Corte ")
      GoTo 99
    End If
    If XVALO(TXTMETROS) <= 0 Then
       Call COMUNI("Imposible Transferir Cantidad de Metros 0")
       GoTo 99
    End If
    
    FECHA% = HOY(HO$)
    CMOV$ = "TR"
    
    CANTI = XVALO(TXTMETROS)
    LOTE$ = TRIM$(TXTLOTE)
    NCOMP& = ProNroComprobante(CMOV$)
    DOPRI$ = "TR." + TRIM$(str$(NCOMP&))
    DOSEC$ = LOTE$
    DESCRIDEPO$ = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPOALMACEN%)))
    REFE$ = "Tr.de " & TRIM$(DESCRIDEPO$)
    Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC%, DEPOALMACEN%, (-1) * CANTI)    'GUARDO EL NUMERO DE OPERACION DE CADA ITEMS TRANSFERIDO
    
    DESCRIDEPO$ = ECOARCH$("TADEPOSI", Chr$(DEPOMATBOL%))
    REFE$ = "Tr.A " & TRIM$(DESCRIDEPO$)
    NCOMP& = ProNroComprobante(CMOV$)
    
    Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC%, DEPOMATBOL%, CANTI)  'GUARDO EL NUMERO DE OPERACION DE CADA ITEMS TRANSFERIDO
    Call ACSALOTE(CI1%, LOTE$, "NOMESS")
    Dim ctp As New Conteplast
    '
    Call ctp.TransfIngresada(MSF0, DEPOMATBOL%, "")
    '
99  Command3.Enabled = True
    Screen.MousePointer = 1
    Call blanformu1
End Sub
Private Sub Command4_Click()
   IDLOT$ = TRIM$(TXTLOTE.TEXT)
   If IDLOT$ <> "" Then
     On Error Resume Next
     CONSUMLOT07.Label4 = TXTCODIGOELEM
     If Err Then Exit Sub
     On Error GoTo 0
     CONSUMLOT07.Label1 = IDLOT$
     CONSUMLOT07.Show 1
   End If
End Sub

Private Sub Command40_Click()
  Call SELECHO("PADRON")
  VDEV$ = VALACT1$("PADRON")
  If VDEV$ <> "" Then
    Me.TXTNUMOPERARIO = TRIM$(VDEV$)
    Me.LBLNOMBREOPERARIO = TRIM$(VALACT2$("PADRON"))
  End If
End Sub


Private Sub Form_Load()
    Campos$ = "Código/A14;Descripción/A26;Lote/A12;Kgs/F10.2;Saldo Mts./f10.2;CI/a0;DOPRILAR/A0"
    Call CARGA_ENCABEZADO(Me.MSF0, Campos$, Me)

    DEPOMATBOL% = XVALO(Control("CONTEPLA", "DEMATBOL"))
    If DEPOMATBOL% < 1 Then
       Call COMUNI("Debe Configurar Depósito de Bolsas")
       Exit Sub
    End If
    Label13 = XVALO(XVALO(Control("CONTEPLA", "DEMATBOL")))
    lbldepogeneral = Control$("", "DEPOCEN")
    
    Dim ctp As New Conteplast
    
    Call ctp.TransfIngresada(MSF0, DEPOMATBOL%, "")
    Call CARGA_ENCABEZADO(MSF1, "CODIGO/A20;Descripción/A40;LOTE/A16;Cons.Mts./f10.1;CI/f6;Op.N°/F6;Kg.Transf/F9.2", Me)
    Call CARGA_ENCABEZADO(MSF2, "Código/A16;Descripción/A28;Lote/A8;Cant.Cortes/F10.1;CI/a0;N°Op./F5;Descrip/A12;Hora/A0;Material/A16;Ubicación/A12;NroInf/a0;IDUbic/a0;N°Operario/F9;Nombre/A16;Peso/f8.2;IdReg/a0", Me)
    


End Sub

Private Sub Label13_Change()
    Label6 = TRIM$(ECOARCH$("TADEPOSI", Chr$(XVALO(Label13))))
End Sub

Private Sub Label5_DblClick(Index As Integer)
   If Me.Height >= 9600 Then
      Me.Height = 7860
   ElseIf Me.Height <= 8500 Then
      Me.Height = 9915
   End If
End Sub


Private Sub lbldepogeneral_Change()
   DEPO% = XVALO(lbldepogeneral)
   Label3 = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPO%)))
End Sub

Private Sub LBLOF_Change()
   NORFA$ = TRIM$(Me.LBLOF)
   If NORFA$ = "" Then MSF1.Rows = 1: Exit Sub
   LBLDESCRIOF = RESP_ZELE_VECT(2)
   CIFAB% = CODIFAB%(NORFA$)
   NombreOperacion$ = "FABRICACIÓN DE BOLSA"
     
   Dim ctp As New Conteplast
   NOPER% = ctp.NumeOperXNombre%(CIFAB%, NombreOperacion$)
   Set ctp = Nothing
   If NOPER% < 1 Then
     Call COMUNI("Atención !! La Orden de Fabricación Seleccionada No Registra Operacion de FABRICACIÓN DE BOLSA")
     Exit Sub
   End If
   
   TXTCANPROY = CantiProyectadaOF(NORFA$)
   TXTNOP = NOPER%
   DESCRIOP = "FABRICACIÓN DE BOLSA"
   
   DEPOMATBOL% = XVALO(Control("CONTEPLA", "DEMATBOL"))
   Call ctp.CargaMatTranfXOrden(Me.MSF1, NORFA$, DEPOMATBOL%, NOPER%)
   Call ctp.CargaOperacionesInformadas(NORFA$, CIFAB%, MSF2, 0, "", 99)
   Set ctp = Nothing
   
   '    TXTCODINT = MSF1.TextMatrix(REG&, 6)
'    CODIGOFAB = CODINT%(Text10)
'
'   Condi1$ = "CodIConj = " & CODIGOFAB & " and  NumeOper = " & NOP%
'   Label7 = VALOBADA("SECOPER", "DescOper", Condi1$)
'
'   Condi$ = "CODICONJ = " & CODIGOFAB
'   Condi$ = Condi$ + " AND (CODIELEM = " & CIELEM% & "  or CODIELEM = " & CODIGORANGO&(CIELEM%) & ")"
'   DESOP$ = TRIM$(Label7)
'   If DESOP$ <> "" Then
'    Condi$ = Condi$ + " AND COMPONENTE = '" & DESOP$ & "'"
'   End If
'   'LO PASO A MTS 'EL CORTE EN LA TABLA FORMULAS ESTA EN CM POR QUE ASI LO CARGAN EN LA COTIZACION
'   TXTCORTE = XVALO(VALOBADA("FORMULAS", "CORTE", Condi$, "NOMESS")) / 100
'
'   Me.TXTCANTICORTES = XVALO(VALOBADA("FORMULAS", "MODUXBB", Condi$, "NOMESS"))
'   NORFA$ = TRIM$(Label4)
'   NOPER% = XVALO(Text3)
'   LOTE$ = TRIM$(TXTLOTE)
'   'CANTIDAD DE CORTES DE LA OF X LOTE Y OPERACION
'   Condi$ = "IDSUBOR = '" & NORFA$ & "' AND IDENLOTE = '" & LOTE$ & "' AND NUMEOPER = " & NOPER%
'   Me.TXTCORTESOPYLOTE = XVALO(VALOBADA("INFOFAB", "Sum(CANTREAL)", Condi$, "NOMESS"))
'   'CANTIDAD DE CORTES DE LA OF X  OPERACION
'   Condi$ = "IDSUBOR = '" & NORFA$ & "' AND NUMEOPER = " & NOPER%
'   Me.TXTCORTESOPERACION = XVALO(VALOBADA("INFOFAB", "Sum(CANTREAL)", Condi$, "NOMESS"))
'   Me.TXTNECESIDADCORTES = FORMATNUM$(XVALO(TXTCANTICORTES) * XVALO(Text14), "F8")
'   TXTRESTAN = FORMATNUM$(XVALO(TXTNECESIDADCORTES) - XVALO(TXTCORTESOPERACION), "F8")
'    End If
   
'   Call ctp.CargaOperacionesInformadas(NORFA$, CIFAB%, MSF1, 0, "", 99)
'   Call CargaMatTranfXOrdenSoloIngresos


End Sub



Private Sub MSF0_DblClick()
    REG& = MSF0.MouseRow
    COL& = MSF0.MouseCol
    If REG& = 0 Or REG& >= MSF0.Rows Then Exit Sub
    
    Call SELECCIONARFILA(MSF0, REG&)
    Me.TXTLOTE2 = MSF0.TextMatrix(REG&, 3)
    NORFA$ = OrdenFabricacionAbiertas$
    If NORFA$ = "" Then Exit Sub
    CODIGOFAB% = CODIFAB%(NORFA$)
    LBLOF = NORFA$ 'ACA EN EL EVENTO CHANGE CARGA DATOS DE LA OF
    NOPER% = XVALO(TXTNOP)
    CIELEM% = XVALO(MSF0.TextMatrix(REG&, 6))
    Dim ctp As New Conteplast
    If ctp.ValidarCodigosOF(CODIGOFAB%, CIELEM%) <= 0 Then
        Call COMUNI("Atención El Código Seleccionado No Corresponde a la Orden de Fabriación Seleccionada")
        Call BLANFORMU2
        LBLOF = ""
        Exit Sub
    End If
    '
    TXTELEMSELEC = CIELEM%
    CONDI$ = "CODICONJ = " & CODIGOFAB
    'Condi$ = Condi$ + " AND (CODIELEM = " & CIELEM% & "  or CODIELEM = " & CODIGORANGO&(CIELEM%) & ")"
    DESOP$ = TRIM$(DESCRIOP)
    If DESOP$ <> "" Then
    CONDI$ = CONDI$ + " AND COMPONENTE = '" & DESOP$ & "'"
    End If
    'LO PASO A MTS 'EL CORTE EN LA TABLA FORMULAS ESTA EN CM POR QUE ASI LO CARGAN EN LA COTIZACION
    TXTCORTE = XVALO(VALOBADA("FORMULAS", "CORTE", CONDI$, "NOMESS")) / 100

    Me.TXTCORTXBSON = XVALO(VALOBADA("FORMULAS", "MODUXBB", CONDI$, "NOMESS"))
    LOTE$ = TRIM$(TXTLOTE2)
    
    TXTNECESIDADCORTES = ROUND(XVALO(TXTCANPROY) * XVALO(TXTCORTXBSON), 0)
    'CANTIDAD DE CORTES DE LA OF X LOTE Y OPERACION
   CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND IDENLOTE = '" & LOTE$ & "' AND NUMEOPER = " & NOPER%
   Me.TXTCORTESINFO = XVALO(VALOBADA("INFOFAB", "Sum(CANTREAL)", CONDI$, "NOMESS"))
   'CANTIDAD DE CORTES DE LA OF X  OPERACION
   CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND NUMEOPER = " & NOPER%
   TXTRESTAN = FORMATNUM$(XVALO(TXTNECESIDADCORTES) - XVALO(TXTCORTESINFO), "F8")
    
End Sub


Private Sub MSF0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    REG& = MSF0.MouseRow
    COL& = MSF0.MouseCol
    If Button = 2 Then
        Dim ctp As New Conteplast
    
        If REG& = 0 Or REG& >= MSF0.Rows Then Exit Sub
        OPX% = COMALTER%("Confirmar ELIMINACIÓN  de Registro de Modo Permanente\\Cancelar\Eliminar Registro Seleccionado")
        If OPX% < 2 Then Exit Sub
        If OPX% = 2 Then
           DEPOMATBOL% = Control("CONTEPLA", "DEMATBOL")
           LOTE$ = MSF0.TextMatrix(REG&, 3)
           PesoOriginal# = PesoOriginalSegúnLote#(LOTE$)
           CI1% = XVALO(MSF0.TextMatrix(REG&, 6))
           NROTRANSF$ = MSF0.TextMatrix(REG&, 7)
           CONDI$ = " IDENLOTE = '" & LOTE$ & "'   AND CODIMOVI = 'TR'  AND DOPRILAR <> '" & NROTRANSF$ & "'"
           If CantRegistros&("MOVISTO", CONDI$, "NOMESS") < 1 Then
             CONDI$ = "CODIMOVI = 'TR' AND DOPRILAR = '" & NROTRANSF$ & "' AND IDENLOTE = '" & LOTE$ & "' AND COD_INTE = " & CI1%
             Call BORRAREGISTROS("MOVISTO", CONDI$, REGAF&, "NOMESS")
             Call ctp.TransfIngresada(MSF0, DEPOMATBOL%, "")
           Else
             Call COMUNI("Imosible Eliminar Este Registro\ El mismo Posee Transferencias e Infomes de Asignación a Orden de Fabricación")
             Exit Sub
           End If
        End If

    End If
End Sub

Private Sub TXTLECTURA_DblClick()
'   TXTLECTURA = IngresoManualCodigoLote$
End Sub

Private Sub TXTLECTURA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     LecturaEtiqueta$ = TRIM$(TXTLECTURA)
     LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, "'", "-")
     LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, "]", "|")
     If LecturaEtiqueta$ = "" Then Exit Sub
     n% = InStr(1, LecturaEtiqueta$, "|")
     If n% < 1 Then Exit Sub
     Dim ctp As New Conteplast

     Dim strArray() As String
     LECTURA$ = LecturaEtiqueta$
     On Error Resume Next
     strArray = Split(LECTURA$, "|")
     TXTCODIGOELEM = strArray(0)
     LOTE$ = TRIM$(strArray(1))
     On Error GoTo 0
     '
     CIELEM% = CODINT%(TXTCODIGOELEM)
     'ESTO ES POR EL KILOMBO QUE HICIERON QUE CAMBIARON CODIGOS DESPUES DE ESTAR GENERADAS LAS ETIQUETAS
     If CIELEM% < 1 Then
       CIELEM% = XVALO(VALOBADA("BOLRECEP", "COD_INTE", "IDENLOTE = '" & LOTE$ & "'", "NOMESS"))
       TXTCODIGOELEM = CODEXT$(CIELEM%)
     End If

     Me.TXTDESCRIELEM = DESCRIT0$(CIELEM%)

     'CICONJ% = Me.txtcodintelem
     Me.txtcodintelem = CIELEM%
     DEPOPAÑOL% = XVALO(lbldepogeneral)
     CANTMETROSBOBINA = SaldoLote(CIELEM%, LOTE$, DEPOPAÑOL%)
     Me.TXTMETROS = CANTMETROSBOBINA
'     PesoPorMetro# = PESMETRO(CIELEM%)
     PesoPorMetro# = CoeficientePesoPorMetroSegunLote#(LOTE$)
     TXTPESOORIG = FORMATNUM$(PesoOriginalSegúnLote#(LOTE$), "F10.2")

     PESOSALDOBOBINA = PesoPorMetro# * CANTMETROSBOBINA
     '
     Me.TXTLOTE = LOTE$
     Me.TXTPESO = FORMATNUM$(PESOSALDOBOBINA, "F10.2")
     On Error GoTo 0
   End If

End Sub

Private Sub TXTNUMOPERARIO_Change()
   Me.LBLNOMBREOPERARIO = ECOARCH$("PADRON", MKI$(XVALO(TXTNUMOPERARIO)))
End Sub


