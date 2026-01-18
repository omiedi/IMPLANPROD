VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CARCOTI07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de Cotizaciones Recibidas"
   ClientHeight    =   7755
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid msf2 
      Bindings        =   "CARCOTI07.frx":0000
      Height          =   2385
      Left            =   240
      TabIndex        =   55
      Top             =   1800
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   4207
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   8421504
      BackColorSel    =   14737632
      ForeColorSel    =   8421504
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox ORDITEM 
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
      Left            =   7320
      TabIndex        =   10
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00CDDADA&
      Caption         =   "Detalle de Oferta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3550
      Left            =   240
      TabIndex        =   33
      Top             =   4200
      Width           =   8055
      Begin VB.TextBox Pordescu 
         Alignment       =   1  'Right Justify
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
         Left            =   4320
         TabIndex        =   51
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Left            =   7320
         TabIndex        =   48
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton Command9 
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
         Height          =   285
         Left            =   7680
         TabIndex        =   47
         Top             =   1800
         Width           =   135
      End
      Begin VB.CommandButton Command8 
         Caption         =   "O/C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   525
         TabIndex        =   22
         Top             =   2340
         Width           =   645
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Registrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5985
         TabIndex        =   21
         Top             =   3075
         Width           =   1815
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
         Height          =   285
         Left            =   7680
         TabIndex        =   20
         Top             =   2760
         Width           =   135
      End
      Begin VB.TextBox Text9 
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
         Left            =   6720
         TabIndex        =   19
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
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
         Left            =   6720
         TabIndex        =   18
         Top             =   1320
         Width           =   1095
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
         Height          =   285
         Left            =   7680
         TabIndex        =   14
         Top             =   360
         Width           =   135
      End
      Begin VB.TextBox Text7 
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
         Left            =   6720
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command4 
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
         Height          =   285
         Left            =   4080
         TabIndex        =   17
         Top             =   840
         Width           =   135
      End
      Begin VB.TextBox Text6 
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
         Left            =   3720
         TabIndex        =   16
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Left            =   1320
         TabIndex        =   15
         Top             =   840
         Width           =   1455
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
         Height          =   1425
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   2040
         Width           =   4335
      End
      Begin VB.CommandButton Command3 
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
         Height          =   285
         Left            =   2040
         TabIndex        =   12
         Top             =   360
         Width           =   135
      End
      Begin VB.TextBox Text3 
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
         Left            =   1320
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblConDescuento 
         Alignment       =   1  'Right Justify
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
         Left            =   4320
         TabIndex        =   54
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Precio C/Descuento:"
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
         Left            =   2160
         TabIndex        =   53
         Top             =   1785
         Width           =   2175
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento:"
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
         Left            =   3120
         TabIndex        =   52
         Top             =   1425
         Width           =   1215
      End
      Begin VB.Label Label22 
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
         Height          =   465
         Left            =   5760
         TabIndex        =   50
         Top             =   2190
         Width           =   2055
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Con.Pago:"
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
         Left            =   5280
         TabIndex        =   49
         Top             =   1905
         Width           =   1335
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
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
         Left            =   6600
         TabIndex        =   45
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label19 
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
         Left            =   4320
         TabIndex        =   44
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Validez:"
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
         Left            =   5400
         TabIndex        =   43
         Top             =   2820
         Width           =   1215
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "P.Entrega:"
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
         Left            =   5450
         TabIndex        =   42
         Top             =   1425
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Importe:"
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
         Left            =   5280
         TabIndex        =   41
         Top             =   945
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Left            =   5400
         TabIndex        =   40
         Top             =   465
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda:"
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
         Left            =   2400
         TabIndex        =   39
         Top             =   945
         Width           =   1215
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Precio:"
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
         Left            =   0
         TabIndex        =   38
         Top             =   940
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
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
         Left            =   240
         TabIndex        =   36
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label11 
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
         Left            =   2280
         TabIndex        =   35
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
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
         Left            =   0
         TabIndex        =   34
         Top             =   460
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame1"
      Height          =   2535
      Left            =   7280
      TabIndex        =   32
      Top             =   1680
      Width           =   15
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame1"
      Height          =   2415
      Left            =   5880
      TabIndex        =   31
      Top             =   1800
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame1"
      Height          =   2500
      Left            =   5400
      TabIndex        =   29
      Top             =   1680
      Width           =   10
      Begin VB.Frame Frame3 
         Caption         =   "Frame1"
         Height          =   2295
         Left            =   300
         TabIndex        =   30
         Top             =   100
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2535
      Left            =   840
      TabIndex        =   28
      Top             =   1680
      Width           =   15
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00404040&
      Height          =   8520
      Left            =   8450
      ScaleHeight     =   8460
      ScaleWidth      =   1005
      TabIndex        =   25
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Height          =   650
         Left            =   120
         Picture         =   "CARCOTI07.frx":0014
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Listado de Sollicitudes por  Rango de SIM"
         Top             =   2280
         Width           =   600
      End
      Begin VB.CommandButton Command22 
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
         Height          =   650
         Left            =   130
         Picture         =   "CARCOTI07.frx":031E
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   100
         Width           =   600
      End
      Begin VB.CommandButton Command13 
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
         Height          =   650
         Left            =   130
         Picture         =   "CARCOTI07.frx":0468
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   750
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "CARCOTI07.frx":0772
         Top             =   6840
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      ItemData        =   "CARCOTI07.frx":2694
      Left            =   240
      List            =   "CARCOTI07.frx":2696
      TabIndex        =   23
      Top             =   2160
      Width           =   8055
   End
   Begin VB.CommandButton Command1 
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
      Height          =   285
      Left            =   3000
      TabIndex        =   9
      Top             =   240
      Width           =   135
   End
   Begin VB.TextBox Text1 
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
      Left            =   1920
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Prov   Proveedor                                                                M.        Pr.Unitario  P.Entrega"
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
      Left            =   240
      TabIndex        =   24
      Top             =   1815
      Width           =   8055
   End
   Begin VB.Label Label9 
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
      Left            =   1560
      TabIndex        =   7
      Top             =   1200
      Width           =   6735
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
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
      Left            =   7200
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label6 
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
      Left            =   1560
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
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
      Left            =   5880
      TabIndex        =   4
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Orden:"
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
      Left            =   5160
      TabIndex        =   3
      Top             =   345
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion: "
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
      Left            =   360
      TabIndex        =   2
      Top             =   1300
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
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
      Left            =   240
      TabIndex        =   1
      Top             =   820
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de SIM:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   340
      Width           =   1575
   End
End
Attribute VB_Name = "CARCOTI07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Command1.Enabled = False
    'Call PEDCOT07.GENINSOL
    'Call FRESHECHO("insolint")
10  If ULTREG&("INSOLINT") > 0 Then ' SI TIENE INFORMACION MUESTRA EL INDICE
        Call SELECHO("INSOLINT/Indice General de Solicitud de Materiales")
        Text1.TEXT = VALACT1$("INSOLINT")
    Else
        GoTo 99
    End If
    '
    'Verifica estado de la sim
    RESS% = XVALO(VALOBADA("ENCASOLINT", "STATUS", "NUME_SOLINT = " & XVALO(Text1.TEXT) & " AND Status = 9"))
    If RESS% >= 8 Then
        Call MENSERR(24, "Imposible Cargar Cotización\S.I.M. Anulada")
        GoTo 10
    End If
99  Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call FINAL("?LISNUMSIM")
    Command2.Enabled = True
End Sub

Private Sub Command22_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Sub BUSCARESIM()
    '
    Label6.Caption = ""
    Label8.Caption = ""
    Label9.Caption = ""
    msf2.Rows = 1
    Text3.TEXT = ""
    Pordescu = ""
    Me.lblConDescuento = ""
    Call LIMPIACASI
    '
    If XVALO(Text1.TEXT) > 9999999 Then
      Call MENSERR(24, "Numero de Solicitud Muy Grande")
      Exit Sub
    End If
    If XVALO(ORDITEM.TEXT) > 20 Then
      Call MENSERR(24, "Numero de Orden Muy Grande")
      Exit Sub
    End If
    '
    CONDI$ = "NUME_SOLINT = " & XVALO(Text1.TEXT) & " AND NUORD_ITEM = " & XVALO(ORDITEM.TEXT)
    CODD% = XVALO(VALOBADA("DETASOLINT", "COD_INTE", CONDI$))
    Label6.Caption = CODEXT(CODD%)
    DESC = VALOBADA("DETASOLINT", "DESCRITEM", CONDI$)
    Label9.Caption = DESC
    CANTT = XVALO(VALOBADA("DETASOLINT", "CANT_SOLINT", CONDI$))
    Label8.Caption = FORMATNUM$(CANTT, "F9.1")
    '
'    Call CARGALISTA(List1, "COTIRECI", "NPRO_OCOM/I5> + PROVED/A42;MONE_EMIS/I3;PRE_UNIT/F13.4>;PLAZ_ENTRE/I4>", CONDI$)
    Call cargagridCotiReci(CONDI$)
    
'    SQLX$ = "SELECT * FROM COTIRECI WHERE " & CONDI$
'    Set COTITMP = FLEXCONN.Execute(SQLX$)
'    '
'    With COTITMP
'       'CARGA EL LISTADO CON LAS COTIZACIONES YA INGRESADAS
'       J% = 0
'          Do While Not .EOF
'            PROVV% = XVALO(!NPRO_OCOM)
'            MONE = XVALO(!MONE_EMIS)
'            PREUN = XVALO(!PRE_UNIT)
'            ENTRE = XVALO(!PLAZ_ENTRE)
'            TXTX$ = AJUSTD$(Str(PROVV%), 5) + " " + AJUSTI$(RASOCLI((-1) * PROVV%), 42) + "   " + AJUSTI$(Str(MONE), 2) + "  " + AJUSTD$(Str(PREUN), 8) + " " + AJUSTI$(Str(ENTRE), 8)
'            List1.List(J%) = TXTX$
'            J% = J% + 1
'            .MoveNext
'          Loop
'        'End If
'    End With
'    COTITMP.Close
'    Set COTITMP = Nothing
'    '
End Sub

Sub cargagridCotiReci(CONDI$)

    SQLX$ = "SELECT Npro_Ocom, Mone_Emis,Pre_Unit,Plaz_Entre, Descuento "
    SQLX$ = SQLX$ + " FROM COTIRECI WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
         j& = j& + 1
         msf2.Rows = j& + 1
         msf2.TextMatrix(j&, 1) = XVALO(!Npro_Ocom)
         msf2.TextMatrix(j&, 2) = rasocli$(-1 * XVALO(!Npro_Ocom))
         msf2.TextMatrix(j&, 3) = XVALO(!Mone_Emis)
         msf2.TextMatrix(j&, 4) = FORMATNUM$(XVALO(!Pre_Unit), "f13.4")
         msf2.TextMatrix(j&, 5) = XVALO(!Plaz_Entre)
         DESCUTEX$ = SAFETEXT$(!Descuento)
         msf2.TextMatrix(j&, 6) = DESCUTEX$
         'le paso el importe y el descuento para que considere el resultante
         msf2.TextMatrix(j&, 7) = FORMATNUM$(ImporteConDescuento#(XVALO(msf2.TextMatrix(j&, 4)), DESCUTEX$), "F13.4")
        .MoveNext
      Loop
    End With
    Me.Pordescu = DESCUTEX$
    RS.Close
    Set RS = Nothing
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("PROVED1")
    NC% = XVALO(TRIM$(VALACT1$("PROVED1")))
    If NC% > 0 Then
        Text3.TEXT = TRIM$(Str$(NC%))
        Label11.Caption = rasocli$((-1) * NC%)
    End If
    '
End Sub

Private Sub Command4_Click()
    Call SELECHO("TAMONED")
    MN% = XVALO(TRIM$(VALACT1$("TAMONED")))
    If MN% > 0 Then
        Text6.TEXT = VALACT1$("TAMONED")
        Label19 = VALACT2("TAMONED")
    End If
End Sub

Private Sub Command5_Click()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text7.TEXT = VDEVV$
        FF% = FECHANUM(Text7.TEXT)
    End If
End Sub

Private Sub Command6_Click()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
    End If
End Sub

Private Sub Command7_Click()
    '
    Command7.Enabled = False
    '
    If Label6 = "" And Label9 = "" Then
        Call MENSERR(24, "Falta Articulo a Cotizar")
        GoTo 99
    End If
    '
    If XVALO(Text3.TEXT) < 1 Then
        Call MENSERR(24, "Debe Seleccionar un Proveedor")
        GoTo 99
    End If
    '
    FF1% = FECHANUM(Text9) ' VALIDEZ
    FF2% = FECHANUM(Text7) ' FECHA
    If FF1% = 0 Or FF2% = 0 Then ' VALIDA QUE EXISTA FECHA Y VALIDEZ
        Call MENSERR(24, "Debe Ingresar Validez y Fecha de Cotización")
        GoTo 99
    End If
    '
    If FF1% < FF2% Then ' VALIDA QUE VALIDEZ SEA MAYOR O IGUAL A LA FECHA DE COTIZACION
        Call MENSERR(24, "Validez Debe Ser Mayor a la Fecha de Cotización")
        GoTo 99
    End If
    '
    COPAGI% = XVALO(Text2)
    On Error Resume Next
    If ECOARCH$("CONPAG", Chr$(COPAGI%)) = "" Then
        Call MENSERR(24, "Concición de Pago No Válida")
        GoTo 99
    End If
    On Error GoTo 0
    '
    '
    'BUSCO EL SECTOR PRODUCTIVO DE LA SOLIICTUD
    CONDI1$ = "NUME_SOLINT = " & XVALO(Text1)
    SECPR$ = VALOBADA("ENCASOLINT", "Sect_Solint", CONDI1$, "NOMESS")
    MONE% = XVALO(Text6)
    If MONE% < 1 Then MONE% = 1
    PREPESOS = XVALO(Text5) * TICAMONE(MONE%)

    
    
10  CONDI$ = "NUME_SOLINT = " & XVALO(Text1.TEXT) & " AND NUORD_ITEM = " & XVALO(ORDITEM.TEXT) & " AND NPRO_OCOM = " & XVALO(Text3.TEXT)
    XX& = CantRegistros&("COTIRECI", CONDI$)
    If XX& < 1 Then
        Call ABRECOMPRAS
        SQLX$ = "SELECT * FROM COTIRECI "
        Set COTIRETMP = New ADODB.Recordset
        COTIRETMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
        With COTIRETMP
          .AddNew        ' SI NO EXISTE AGREGA LOS DATOS A LA TABLA DE COTIZACIONES   ' YA CONVERTIDO
          On Error Resume Next
          !CodiEmpr = CodiEmp%
          On Error GoTo 0
          !NUME_SOLINT = XVALO(Text1.TEXT)
          !NUORD_ITEM = XVALO(ORDITEM.TEXT)
          !Npro_Ocom = XVALO(Text3.TEXT)
          !COD_INTE = CODINT(Label6.Caption)
          !Cod_Exte = SAFETEXT(Label6.Caption)
          !DESCRITEM = SAFETEXT(Label9.Caption)
          !CANT_SOLINT = XVALO(Label8.Caption)
          !UNI_SOL = VALOBADA("DETASOLINT", "UNI_SOL", " NUME_SOLINT = " & XVALO(Text1.TEXT) & " AND NUORD_ITEM = " & XVALO(ORDITEM.TEXT))
          !Sect_Solint = SECPR$
          !PRE_UNIT_PES = PREPESOS
          !Descuento = TRIM$(Left$(Pordescu, 64))
          !CONDPAG = XVALO(Text2)
          .Update
       End With
       COTIRETMP.Close
       '
    End If
    
    'REALIZA LA ACTULIZACION DE LAS COTIZACIONES
    
    SQL1$ = "SELECT * FROM COTIRECI WHERE " & CONDI$
    Dim RSTCOTIRECI As ADODB.Recordset
    Set RSTCOTIRECI = New ADODB.Recordset
    RSTCOTIRECI.Open FILTSQL$(SQL1$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With RSTCOTIRECI
        If Not (.EOF And .BOF) Then
                PRE$ = TRIM$(FORMATNUM(Text5, "F15.4")): If PRE$ = "" Then PRE$ = "0" ' VALOR POR DEFAULT CUANDO NO INGRESA NADA
            !Pre_Unit = XVALO(PRE$)
                MON$ = TRIM$(FORMATNUM(Text6, "I3")): If MON$ = "" Then MON$ = "0" ' VALOR POR DEFAULT CUANDO NO INGRESA NADA
            !Mone_Emis = XVALO(MON$)
            !FECH_COTI = CVDAT(FECHANUM(Text7.TEXT))
            !VENCI_OFER = CVDAT(FECHANUM(Text9.TEXT))
                PLAENT$ = TRIM$(FORMATNUM(Text10, "I5")): If PLAENT$ = "" Then PLAENT$ = 0 ' VALOR POR DEFAULT CUANDO NO INGRESA NADA
            !Plaz_Entre = XVALO(PLAENT$)
                IMPO$ = TRIM$(FORMATNUM$(Label20, "F15.4")): If IMPO$ = "" Then IMPO$ = "0" ' VALOR POR DEFAULT CUANDO NO INGRESA NADA
            !INET_ITEM = XVALO(IMPO$)
            !OBSERVA = Text4.TEXT
            !CONDPAG = XVALO(Text2)
            !Sect_Solint = SECPR$
            !PRE_UNIT_PES = XVALO(PREPESOS)
            !Descuento = TRIM$(Left$(Pordescu, 64))
            .Update
        End If
    End With
    RSTCOTIRECI.Close
    Set RSTCOTIRECI = Nothing
    
'      PRE$ = TRIM$(FORMATNUM(Text5, "F15.4")): If PRE$ = "" Then PRE$ = "0" ' VALOR POR DEFAULT CUANDO NO INGRESA NADA
'    Call ACTUREGISTRO("COTIRECI", "PRE_UNIT", CONDI$, PRE$, XYX&)
'      MON$ = TRIM$(FORMATNUM(Text6, "I3")): If MON$ = "" Then MON$ = "0" ' VALOR POR DEFAULT CUANDO NO INGRESA NADA
'    Call ACTUREGISTRO("COTIRECI", "MONE_EMIS", CONDI$, MON$, XYX&)
'    Call ACTUREGISTRO("COTIRECI", "FECH_COTI", CONDI$, CVDAT(FECHANUM(Text7.TEXT)), XYX&)
'    Call ACTUREGISTRO("COTIRECI", "VENCI_OFER", CONDI$, CVDAT(FECHANUM(Text9.TEXT)), XYX&)
'      PLAENT$ = TRIM$(FORMATNUM(Text10, "I5")): If PLAENT$ = "" Then PLAENT$ = 0 ' VALOR POR DEFAULT CUANDO NO INGRESA NADA
'    Call ACTUREGISTRO("COTIRECI", "PLAZ_ENTRE", CONDI$, PLAENT$, XYX&)
'      IMPO$ = TRIM$(FORMATNUM$(Label20, "F15.4")): If IMPO$ = "" Then IMPO$ = "0" ' VALOR POR DEFAULT CUANDO NO INGRESA NADA
'    Call ACTUREGISTRO("COTIRECI", "INET_ITEM", CONDI$, IMPO$, XYX&)
'    Call ACTUREGISTRO("COTIRECI", "OBSERVA", CONDI$, Text4.TEXT, XX&)
'    Call ACTUREGISTRO("COTIRECI", "CONDPAG", CONDI$, XVALO(Text2), XX&) ' CONDICION DE PAGO
'    Call ACTUREGISTRO("COTIRECI", "Sect_Solint", CONDI$, SECPR$, XX&) ' CONDICION DE PAGO
'    Call ACTUREGISTRO("COTIRECI", "PRE_UNIT_PES", CONDI$, PREPESOS, XX&) ' CONDICION DE PAGO

    '
    TTXX$ = "" 'Graba en el anotador de solicitudes el proveedor al que se le cargo cotizacion
    TTXX$ = "Cotiza " & rasocli((-1) * XVALO(Text3.TEXT)) & CRLF$
    TTXX$ = TTXX$ & "Item Numero: " & (i& + 1) & CRLF$
    NSLI& = XVALO(Text1.TEXT)
    Call ANOTASIM07.ANOTASIM(NSLI&, TTXX$)
    '
    Call BUSCARESIM
    '
99  Command7.Enabled = True

End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    OPX% = COMALTER%("Modo de Generación:\\Item Unico\Más de un Item")
    If OPX% = 1 Then
         GoTo 5
       ElseIf OPX% = 2 Then
         Call OCOMUL
    End If
    GoTo 99
    '
5   If XVALO(Text1) < 1 Then
        Call MENSERR(24, "Falta Numero de SIM")
        GoTo 99
    End If
    If XVALO(ORDITEM) < 1 Then
        Call MENSERR(24, "Falta Numero de Orden")
        GoTo 99
    End If
    If XVALO(Text3) < 1 Then
        Call MENSERR(24, "Falta Proveedor Adjudicado")
        GoTo 99
    End If
    If XVALO(Text5) < 0.00005 Then
        Call MENSERR(24, "Falta Precio")
        GoTo 99
    End If
    '
10  CONDI$ = "NUME_SOLINT = " & XVALO(Text1.TEXT) & " AND NUORD_ITEM = " & XVALO(ORDITEM.TEXT) & " AND NPRO_OCOM = " & XVALO(Text3.TEXT)
    XX& = CantRegistros&("COTIRECI", CONDI$)
    If XX& < 1 Then
        Call MENSERR(24, "Cotizacion no Registrada.")
        GoTo 99
    End If
    APLI$ = "CREC" + AJUSTI$(Text1, 8) + AJUSTI$(ORDITEM, 4) + AJUSTI$(Text3, 6)
    Call CONECRUN("ORCOMP07/" + APLI$, "Generacion de Ordenes de Compra")
99  Command8.Enabled = True
End Sub

Sub OCOMUL()
    '
    Call PoneEnCeroNull("COTIRECI", "NUME_OCOM", RAFE&)
    '
    SQLX$ = "SELECT * FROM COTIRECI WHERE nume_ocom is null or nume_ocom = 0 order by nume_solint asc, nuord_item asc"
    Dim COTRECI As ADODB.Recordset
    Set COTRECI = New ADODB.Recordset
    COTRECI.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    '
    Screen.MousePointer = 11
    With COTRECI
       Do While Not .EOF
          NUSOLI& = XVALO(!NUME_SOLINT)
          NORIT% = XVALO(!NUORD_ITEM)
          CONDI1$ = "nume_solint = " & NUSOLI& & " AND nuord_item = " & NORIT%
          If CantRegistros("COTIRECI", CONDI1$ & " AND NUME_OCOM > 0") > 0 Then
             CONDI2$ = "NUME_OCOM = 0 And " & CONDI1$
             Call ACTUREGISTRO("COTIRECI", "NUME_OCOM", CONDI2$, "-1", RAFE&)
          End If
       .MoveNext
       Loop
    End With
    Screen.MousePointer = 1
    '
    Call COPYSTRU("PROVED1", "PROVEDX")
    CONDI$ = "NUME_OCOM = 0"
    Call CARGALISEL("PROVEDX", "COTIRECI+PROVED12/COTIRECI.NPRO_OCOM = PROVED12.NUME_CLI", "COTIRECI.NPRO_OCOM/F6;PROVED12.RASO_CLI/A36", CONDI$, "DISTINCT")
    Call SELECHO("PROVEDX")
    '
    NPROX& = XVALO(VALACT1$("PROVEDX"))
    If NPROX& > 0 Then
        NroProveedor$ = CStr(NPROX&)
        NroProveedor$ = AJUSTD$(NroProveedor$, 6)
        If XVALO(Text1) > 0 Then
            NroSolicitud$ = TRIM$(Text1)
            NroSolicitud$ = AJUSTD$(NroSolicitud$, 6)
        End If
        APLI$ = "CREM" & NroProveedor$ & NroSolicitud$
        Call CONECRUN("ORCOMP07/" + APLI$, "Generacion de Ordenes de Compra")
    End If
    
End Sub
'
Private Sub Command9_Click()
   '
   Call SELECHO("CONPAG")
   VDEVU$ = VALACT1$("CONPAG")
   If TRIM$(VDEVU$) <> "" Then
      Text2 = VDEVU$
   End If
   '
End Sub

Private Sub Form_Load()
   Call CARGA_ENCABEZADO(Me.msf2, "Cta/F3;R.Social Prov/A18;M/A2;Pr.Unit/F8.2;P.Entrega/A10;Desc/A8;C-Desc./F8.2", Me)
End Sub

Private Sub lblConDescuento_Change()
   lblConDescuento = FORMATNUM$(lblConDescuento, "F13.4")
End Sub

Private Sub List1_Click()
   'CARGA EL DETALLE DE LA OFERTA
   If msf2.Rows < 2 Then Exit Sub
   REG& = msf2.MouseRow
   For XX% = 0 To List1.ListCount - 1
       If List1.Selected(XX%) = True Then
          TTXX$ = List1.LIST(XX%)
          Text3.TEXT = msf2.TextMatrix(REG&, 1)
          Text5.TEXT = msf2.TextMatrix(REG&, 4)
          'Text5.TEXT = TRIM$(Mid$(TTXX$, 55, 12))
          'Text6.TEXT = TRIM$(Mid$(TTXX$, 51, 2))
          Text6.TEXT = msf2.TextMatrix(REG&, 3)
          'Text10.TEXT = TRIM$(Mid$(TTXX$, 68, 8))
          Text10.TEXT = msf2.TextMatrix(REG&, 5)
       End If
    Next XX%
    
    'PRESENTAR LA CONDICION DE PAGO
    CONDI$ = "NUME_SOLINT = " & XVALO(Text1.TEXT) & " AND NUORD_ITEM = " & XVALO(ORDITEM.TEXT) & " AND NPRO_OCOM = " & XVALO(Text3.TEXT)
    CONPA% = XVALO(VALOBADA("COTIRECI", "CONDPAG", CONDI$, "NOMESS"))
    Text2 = CONPA%
End Sub

Private Sub msf2_DblClick()
    REG& = msf2.MouseRow
    
    If REG& = 0 Then Exit Sub
    If REG& > msf2.Rows - 1 Then Exit Sub
    Text3.TEXT = msf2.TextMatrix(REG&, 1)
    Text5.TEXT = msf2.TextMatrix(REG&, 4)
    Text6.TEXT = msf2.TextMatrix(REG&, 3)
    Text10.TEXT = msf2.TextMatrix(REG&, 5)
    
    'PRESENTAR LA CONDICION DE PAGO
    CONDI$ = "NUME_SOLINT = " & XVALO(Text1.TEXT) & " AND NUORD_ITEM = " & XVALO(ORDITEM.TEXT) & " AND NPRO_OCOM = " & XVALO(Text3.TEXT)
    CONPA% = XVALO(VALOBADA("COTIRECI", "CONDPAG", CONDI$, "NOMESS"))
    Text2 = CONPA%
End Sub


Private Sub ORDITEM_Change()
    Call BUSCARESIM
End Sub

Private Sub ORDITEM_Click()
    Call BUSCARESIM
End Sub

Private Sub Pordescu_Change()
 Me.lblConDescuento = FORMATNUM$(ImporteConDescuento#(XVALO(Text5), Pordescu), "F13.4")
End Sub

Private Sub Text1_Change()

    ORDITEM.Clear   'LIMPIA COMBO PARA LIMPIAR FORMULARIO
    Call BUSCARESIM 'Carga codigo del articulo, descripcion y cantidad segun solicitud
    '
    XREG& = XVALO(VALOBADA("DETASOLINT", "COUNT(NUORD_ITEM)", "NUME_SOLINT = '" & Text1.TEXT & "'"))
    For JJ& = 1 To XREG&        'CARGA LOS NUMEROS DE LOS ITEMS SEGUN XREG%(REGISTROS QUE CUMPLEN CONDICION)
        ORDITEM.AddItem (JJ&)
    Next JJ&
End Sub

Private Sub Text1_DblClick()
    Call Command1_Click
End Sub


Sub LIMPIACASI()
    '
    Label11.Caption = ""
    Label19.Caption = ""
    Label20.Caption = ""
    Text4.TEXT = ""
    Text5.TEXT = ""
    Text6.TEXT = ""
    Text7.TEXT = ""
    Text9.TEXT = ""
    Text10.TEXT = ""
    Text2 = ""
    Me.lblConDescuento = ""
    Me.Pordescu = ""
    
    '
End Sub

Private Sub Text2_Change()
    COPAGI% = XVALO(Text2)
    On Error Resume Next
    If COPAGI% > 0 And ECOARCH$("CONPAG", Chr$(COPAGI%)) Then
       Label22 = ECOARCH$("CONPAG", Chr$(COPAGI%))
    Else
       Label22 = ""
    End If
    On Error GoTo 0

End Sub

Private Sub Text3_Change()
    '
    Call LIMPIACASI 'LIMPIA LOS CASILLEROS DEL DETALLE DE OFERTA
    '
    NPRO& = XVALO(Text3)
    Label11.Caption = rasocli$((-1) * NPRO&) 'carga razon social del proveedor
    '
    CONDI$ = "NUME_SOLINT = " & XVALO(Text1.TEXT) & " AND NUORD_ITEM = " & XVALO(ORDITEM.TEXT) & " AND NPRO_OCOM = " & NPRO&
    CONDI2$ = "NPRO_OCOM = " & NPRO& & " AND COD_INTE = " & CODINT%(Label6.Caption) & " ORDER BY FECH_COTI DESC"
    '
    Text5 = FORMATNUM$(XVALO(VALOBADA("COTIRECI", "PRE_UNIT", CONDI$)), "F10.4") 'carga precio unitario del articulo
    Text6 = XVALO(VALOBADA("COTIRECI", "MONE_EMIS", CONDI$)) ' carga la moneda de emision de la cotizacion
    '
    If TRIM$(Text5.TEXT) = "" And Label6.Caption <> "" Then
        Text5 = FORMATNUM$(XVALO(VALOBADA("COTIRECI", "PRE_UNIT", CONDI2$)), "F10.4") ' CARGA PRECIO ULTIMA COTIZACION
        Text6 = XVALO(TRIM$((VALOBADA("COTIRECI", "MONE_EMIS", CONDI2$))))
    End If
    '
    'Text6 = XVALO(VALOBADA("COTIRECI", "MONE_EMIS", CONDI$)) ' carga la moneda de emision de la cotizacion
    Text7 = FECHATEX$(CVINT(VALOBADA("COTIRECI", "FECH_COTI", CONDI$))) ' carga facha cotizacion
    'IMPO = XVALO(VALOBADA("COTIRECI", "INET_ITEM", CONDI$)) ' carga importe de la cotizacion
    '
    Text9 = FECHATEX$(CVINT(VALOBADA("COTIRECI", "VENCI_OFER", CONDI$))) ' carga la fecha de vencimiento de la cotizacion
    Text10 = XVALO(VALOBADA("COTIRECI", "PLAZ_ENTRE", CONDI$)) ' carga el plazo de vigencia de la cotizacion
    Text4 = VALOBADA("COTIRECI", "OBSERVA", CONDI$)
    Pordescu = VALOBADA("COTIRECI", "DESCUENTO", CONDI$)
    
End Sub

Private Sub Text3_DblClick()
    Call Command3_Click
End Sub

Private Sub Text5_Change()
    '
    IMPTO = XVALO(Text5.TEXT) * XVALO(Label8.Caption) 'importe = cantidad * Preciounitario
    Label20.Caption = FORMATNUM$(IMPTO, "F11.2")
    lblConDescuento = ImporteConDescuento#(XVALO(Text5), Pordescu)
    '
End Sub

Private Sub Text5_LostFocus()
    Text5 = FORMATNUM$(Text5.TEXT, "F13.4")
End Sub

Private Sub Text6_Change()
    Label19 = ""
    MONE% = XVALO(Text6)
    If MONE% > 0 Then Label19.Caption = ECOARCH$("TAMONED", Chr$(FORMATNUM$(MONE%, "I3")))
    Text6 = FORMATNUM$(MONE%, "I3")
End Sub

Private Sub Text6_DblClick()
    Call Command4_Click
End Sub

Private Sub Text7_DblClick()
    Call Command5_Click
End Sub

Private Sub Text8_Change()

End Sub

Private Sub Text9_DblClick()
    Call Command6_Click
End Sub

