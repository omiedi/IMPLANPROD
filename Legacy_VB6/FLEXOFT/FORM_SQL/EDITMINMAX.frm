VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form EDITMINMAX 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vista y Edicion de Stocks Mínimos y Máximos de Artículos"
   ClientHeight    =   10425
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   14580
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10425
   ScaleWidth      =   14580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   540
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   " "
      Top             =   9840
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Planilla Origen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   0
      TabIndex        =   37
      Top             =   5360
      Visible         =   0   'False
      Width           =   14600
      Begin VB.CommandButton Command6 
         Caption         =   "Salir"
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
         Left            =   11400
         Picture         =   "EDITMINMAX.frx":0000
         TabIndex        =   43
         Top             =   600
         Width           =   3015
      End
      Begin VB.CommandButton Command7 
         Height          =   375
         Left            =   13800
         Picture         =   "EDITMINMAX.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Descargar Datos de Planilla Excel"
         Top             =   195
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IntegralHeight  =   0   'False
         Left            =   11400
         TabIndex        =   38
         Text            =   " "
         Top             =   195
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
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
         Left            =   960
         TabIndex        =   42
         Top             =   240
         Width           =   10275
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Label9"
         Height          =   195
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exportar a Excel"
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
      Left            =   7560
      Picture         =   "EDITMINMAX.frx":0614
      TabIndex        =   36
      ToolTipText     =   "Guardar "
      Top             =   5700
      Width           =   2385
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Bajar Items Seleccionados a Grilla de Envio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   680
      Left            =   9960
      Picture         =   "EDITMINMAX.frx":091E
      TabIndex        =   35
      ToolTipText     =   "Guardar "
      Top             =   5350
      Width           =   4545
   End
   Begin VB.TextBox TXTMOVIL2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Left            =   480
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Costo segun el proveedor y la via de entrega"
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
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
      Left            =   9880
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Text            =   "Cta."
      Top             =   0
      Width           =   470
   End
   Begin VB.TextBox TXTCTADECALCULO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Left            =   9880
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   240
      Width           =   470
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
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
      TabIndex        =   30
      TabStop         =   0   'False
      Text            =   "Cte. Selec. de Cálculo"
      Top             =   0
      Width           =   2655
   End
   Begin VB.TextBox TXTCLIENTECALCULADO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exportar a Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   10800
      TabIndex        =   29
      Top             =   9960
      Width           =   1605
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   13095
      ScaleHeight     =   75
      ScaleWidth      =   1425
      TabIndex        =   25
      Top             =   40
      Width           =   1485
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
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
      Left            =   7125
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "Fecha Cálculo"
      Top             =   0
      Width           =   2775
   End
   Begin VB.TextBox TXTFECHACALCULO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Left            =   7125
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   240
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
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
      Left            =   5295
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "Usuario"
      Top             =   0
      Width           =   1845
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "Último Envío"
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox TXTUSUARIO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   240
      Width           =   1860
   End
   Begin VB.TextBox TXTULTIMOENVIO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Grabar Vista Actual"
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
      Left            =   7560
      Picture         =   "EDITMINMAX.frx":0C28
      TabIndex        =   18
      ToolTipText     =   "Guardar "
      Top             =   5350
      Width           =   2385
   End
   Begin VB.TextBox TXTMOVIL 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Left            =   360
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Costo segun el proveedor y la via de entrega"
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
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
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "Cliente"
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox TXTCLIENTE 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text5 
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
      Left            =   1095
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "R.Social"
      Top             =   0
      Width           =   3045
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   20
      TabIndex        =   8
      Top             =   5280
      Width           =   7455
      Begin VB.TextBox TXTBUSCAR 
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
         Left            =   740
         TabIndex        =   10
         Top             =   160
         Width           =   2535
      End
      Begin VB.ComboBox CMBFILTRO 
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
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   160
         Width           =   3135
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   160
         Width           =   585
      End
      Begin VB.Label Label71 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Campo:"
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
         Left            =   3360
         TabIndex        =   11
         Top             =   160
         Width           =   705
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enviar Actualización de St.Min. y Max."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   12480
      Picture         =   "EDITMINMAX.frx":0F32
      TabIndex        =   7
      ToolTipText     =   "Guardar "
      Top             =   9960
      Width           =   2025
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   9840
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "EDITMINMAX.frx":123C
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   8520
      OleObjectBlob   =   "EDITMINMAX.frx":1470
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   4725
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar"
      Top             =   600
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   8334
      _Version        =   393216
      Cols            =   4
      BackColor       =   15724527
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
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
   Begin MSFlexGridLib.MSFlexGrid MSF_2 
      Height          =   4725
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
      Top             =   600
      Visible         =   0   'False
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   8334
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   8421631
      BackColorBkg    =   12632256
      FillStyle       =   1
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
   Begin VB.TextBox TXTRASOCLI 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Left            =   1095
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   240
      Width           =   3045
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   3765
      Left            =   0
      TabIndex        =   27
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Superior"
      Top             =   6120
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   6641
      _Version        =   393216
      Cols            =   4
      BackColor       =   15724527
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
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
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Items Seleccionados Para Enviar Actualización"
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
      Left            =   120
      TabIndex        =   28
      Top             =   5900
      Width           =   3855
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   13080
      Picture         =   "EDITMINMAX.frx":14CE
      ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
      Top             =   195
      Width           =   1515
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   8520
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   240
      Left            =   8400
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   10080
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Visible         =   0   'False
      Begin VB.Menu mnucopiar 
         Caption         =   "Copiar Celda Seleccionada"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuabrirgrafico 
         Caption         =   "Abrir Estadístico de Consumo"
         Visible         =   0   'False
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuArchivo2 
      Caption         =   "Archivo2"
      Visible         =   0   'False
      Begin VB.Menu mnucopiar2 
         Caption         =   "Copiar Celda Seleccionada"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuabrirgrafico2 
         Caption         =   "Abrir Estadístico de Consumo"
         Visible         =   0   'False
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuEditar 
      Caption         =   "Editar"
      Begin VB.Menu Refrescalc 
         Caption         =   "Refrescar Desde Cálculo Estadisticos"
      End
      Begin VB.Menu revertirseleccion 
         Caption         =   "Revertir Selección"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
      Begin VB.Menu ultimavistaguardada 
         Caption         =   "Último Vista Guardada"
      End
   End
   Begin VB.Menu mnuimpo 
      Caption         =   "Importar"
      Begin VB.Menu impostockminmax 
         Caption         =   "Importar desde Excel Stocks Min. y Max."
      End
   End
End
Attribute VB_Name = "EDITMINMAX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FrmLeft, FrmTop, FrmWidth, FrmHeight As Long
Dim GridTop, GridLeft, GridWidth, GridHeight As Long
Dim FRAMETop, FRAMELeft, FRAMEWidth, FRAMEHeight As Long
Dim Com1Top, Com2Top, Com2Left As Long
Dim Com1LEFT As Long
Dim GLOBALFRM_strContenidoCelda$
Dim GLOBALFRM_FILA&
Dim LecturaDatosOriginales As Byte


Private objexcel As Excel.Application
Private Libro As Excel.Workbook
Private Hoja As Excel.Worksheet

Private Sub BOTEXIT_Click()
  Unload Me
End Sub

Sub GENERAR_ARCHIVO_TRANSFERENCIA_STOCK_MIN_MAX(NCLIE)
   Ruta$ = ""
   NC_CONFIG1 = XVALO(CONTROL("RETRASLA", "SUC1CLIE"))
   NC_CONFIG2 = XVALO(CONTROL("RETRASLA", "SUC2CLIE"))
   If NC_CONFIG1 = NCLIE Then
        Ruta$ = CONTROL("RETRASLA", "SUC1RUTA")
    ElseIf NC_CONFIG2 = NCLIE Then
        Ruta$ = CONTROL("RETRASLA", "SUC2RUTA")
   End If
   '
   If Ruta$ = "" Then
      Call COMUNI("Atención !!!!! Sucursal Seleccinada No es Válida\Verifique Configuración de Remito a Sucursales\No se Generó Archivo de Envio de Stock Sugeridos")
      Exit Sub
   End If
   
   'CARGA LOS TITULOS
   A$ = Date
   A$ = REPLACAR(A$, "/", "")
   B$ = Time
   B$ = REPLACAR(B$, ":", "")
   B$ = REPLACAR(B$, ".", "")
   NARCHI2% = FILEOPEN%(Ruta$ + "\ARCHIVO_TRANSF_MIN_MAX " + A$ + " " + B$ + ".TXT", 2, 256)    ' ABRE EL ARCHIVO DE ESCRITURA
    TX$ = AJUSTI$("NCLIENTE", 8): TX$ = TX$ + ";" 'NCLIENTE
    TX$ = TX$ + AJUSTI$("CODIGO", 8): TX$ = TX$ + ";"  'CODIGO
    TX$ = TX$ + AJUSTI$("ULT.ENV.", 8): TX$ = TX$ + ";" 'FECHA ULTIMO ENVIO
    TX$ = TX$ + AJUSTI$("USUARIO", 8): TX$ = TX$ + ";" 'USUARIO
    TX$ = TX$ + AJUSTI$("CALCULO", 8): TX$ = TX$ + ";" 'FECHA DE CALCULO
    TX$ = TX$ + AJUSTI$("EMISION", 8): TX$ = TX$ + ";" 'FECHA DE EMISION
    TX$ = TX$ + AJUSTI$("MINSUGER", 10): TX$ = TX$ + ";" 'StockMinimoSugerido
    TX$ = TX$ + AJUSTI$("MAXSUGER", 10) 'StockMaximoSugerido
   Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO STRING DE CONEXION
   FE% = HOY(HO$)
   POS1% = InStr(TXTFECHACALCULO, " ")
   NUSER1% = USNBR% Mod 100
   FECHACALCULO% = CVINT(Mid$(TXTFECHACALCULO, POS1% + 1, 8))

   For I& = 1 To msf2.Rows - 1
         TX$ = AJUSTI$(Me.TXTCLIENTE, 8): TX$ = TX$ + ";" 'NCLIENTE
         TX$ = TX$ + AJUSTI$(CODINT%(msf2.TextMatrix(I&, 2)), 8): TX$ = TX$ + ";"   'CODIGO
         TX$ = TX$ + AJUSTI$(FECHATEX$(FE%), 8): TX$ = TX$ + ";" 'FECHA ULTIMO ENVIO
         TX$ = TX$ + AJUSTI$(SAFETEXT$(NUSER1%), 8): TX$ = TX$ + ";" 'USUARIO
         TX$ = TX$ + AJUSTI$(FECHATEX$(FECHACALCULO%), 8): TX$ = TX$ + ";" 'FECHA DE CALCULO
         TX$ = TX$ + AJUSTI$(FECHATEX$(CVINT(FE%)), 8): TX$ = TX$ + ";" 'FECHA DE EMISION
         TX$ = TX$ + AJUSTI$(msf2.TextMatrix(I&, 6), 10): TX$ = TX$ + ";" 'StockMinimoSugerido
         TX$ = TX$ + AJUSTI$(msf2.TextMatrix(I&, 7), 10): TX$ = TX$ + ";" 'StockMaximoSugerido
         Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO STRING DE CONEXION
    Next I&
    Close #NARCHI2%

End Sub

Function VALIDARREPETICION%(CI1%)
   VALIDARREPETICION% = 0
   For I& = 1 To msf2.Rows - 1
      CILEIDO% = msf2.TextMatrix(I&, 1)
      If CILEIDO% = CI1% Then
         VALIDARREPETICION% = 1
      End If
   Next I&
End Function

Private Sub CMBFILTRO_Change()
   TXTBUSCAR_Change
End Sub

Private Sub Combo1_Click()
   CANOJ% = Libro.Sheets.Count
   For KKI% = 1 To CANOJ%
      If Combo1.List(KKI% - 1) = Libro.Sheets(KKI%).Name Then
        Set Hoja = Libro.Sheets(KKI%)
        Exit Sub
      End If
   Next KKI%
   Combo1.TEXT = Libro.Sheets(1).Name
   Set Hoja = Libro.Sheets(1)
   msf2.Clear
   msf2.Rows = 1

End Sub


Private Sub Command1_Click()
   Command1.Enabled = False
   '
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   AREXPO$ = UCase$(CommonDialog1.FileName)
   If TRIM$(AREXPO$) <> "" Then
      Call GENEXCEL(AREXPO$, Me.msf2)
   End If
   '

99 Command1.Enabled = True

End Sub
Sub GENEXCEL(NombreDeArchivo$, GRID As MSFlexGrid)
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim VALOR As Double
   Dim TEXTO As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   TIPODATO% = 0: COLUM& = 0
   For I& = 1 To GRID.Rows
       RENG1& = I&: COLUM& = 0
       For COLU1& = 1 To GRID.Cols
          If GRID.ColWidth(COLU1& - 1) = 0 Then GoTo 10 ' OBVIA LAS COLUMNAS OCULTAS
          COLUM& = COLUM& + 1
          '
          VALORX = GRID.TextMatrix(I& - 1, COLU1 - 1)
          If Len(VALORX) = 8 And Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
             FEXU$ = VALORX
             VALORX = FETEXCOR1$(Int(FECHANUM(FEXU$)))
             TIPODATO% = 1
          End If
          ' ESTO ES PARA QUE EXPORTE COMO CERO Y NO ESPACIOS
          FFCC$ = UCase$(Left$(FORCAMPO$(COLU1&), 1))
          If FFCC$ = "I" Or FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
             If TRIM$(CStr(VALORX)) = "" Then
                VALORX = 0
             End If
             TIPODATO% = 2
          End If
          If FFCC$ = "C" Or FFCC$ = "A" Then VALORX = "'" & VALORX
          '
          'PARA QUE NO REPITA ENCABEZADO
'          If RENG1& = 1 Then
'            If VALORX = VALORX_ANT And TRIM$(VALORX) <> "" Then
'              If COLUM& <= 17 Then
'               LETRAX1$ = Chr$(64 + COLUM& - 1)
'               LETRAX2$ = Chr$(64 + COLUM&)
'               RANXO$ = LETRAX1$ & TRIM$(Str$(RENG1&)) & ":" & LETRAX2$ & TRIM$(Str$(RENG1&))
'               Range(RANXO$).Select
'
'               XLSHEET.Cells.HorizontalAlignment = xlCenter
'               XLSHEET.Cells.HorizontalAlignment = xlCenter
'               XLSHEET.Cells.VerticalAlignment = xlCenter
'               XLSHEET.Cells.WrapText = False
'               XLSHEET.Cells.Orientation = 0
'               XLSHEET.Cells.AddIndent = False
'               XLSHEET.Cells.IndentLevel = 0
'               XLSHEET.Cells.ShrinkToFit = False
'               XLSHEET.Cells.ReadingOrder = xlContext
'               XLSHEET.Cells.MergeCells = True
'               GoTo 10
'             End If
'            End If
'          End If
          '
          XLSHEET.Cells(RENG1&, COLUM&) = VALORX
          If TIPODATO% = 1 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yyyy"
          'If TIPODATO% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#,##0.00"
          If TIPODATO% = 0 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
          TIPODATO% = 0
          VALORX_ANT = VALORX
10     Next COLU1&
   Next I&
   '
   Close #N1%
   On Error GoTo 30
   XLSHEET.SaveAs NombreDeArchivo$
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   Screen.MousePointer = 1
   Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NombreDeArchivo$, "\", "/"))
   '
Exit Sub
   '
   '
20 Resume 21
21 On Error GoTo 0
   MsgBox "Utilitario EXCEL no Instalado"
   On Error GoTo 0
   Exit Sub
   '
30 Resume 31
31 On Error GoTo 0
   Screen.MousePointer = 1
   '
   ' Cierra el objeto Book, el cual es utilizado al crear
   ' una nueva hoja en el archivo Excel.
   ' El valor False es para que no pregunte si quiere
   ' guardar el archivo.
   XLBOOK.Close (False)
   XLAPP.Quit ' Cierra Excel
   Set XLAPP = Nothing
   Set XLBOOK = Nothing
   Set XLSHEET = Nothing
   '
   Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
   '
End Sub



Private Sub Command2_Click()
    Command2.Enabled = False
    If Me.msf2.Rows <= 1 Then
      Call COMUNI("Grilla de Items Seleccionados Seleccionados Para Enviar Vacía")
      GoTo 99
    End If
    If XVALO(Me.TXTCLIENTE) < 1 Then
      Call COMUNI("No se ha Seleccionado la Sucursal")
      GoTo 99
    End If

    Screen.MousePointer = 11
    Call GRABAR_Y_ENVIAR
    
    'REFRESCO LA VISTA
    Call ultimavistaguardada_Click
    '
    
99  Screen.MousePointer = 1
    Command2.Enabled = True
End Sub

Sub GRABAR_Y_ENVIAR()

   '
   Call BLOQUEAR_CONTROLES
   
   CONDI$ = " nuclien = " & XVALO(Me.TXTCLIENTE)
   CONDI$ = CONDI$ + " AND MARBORRA = 0 "
   Call PoneEnCeroNull("StockSugXSuc", "MARBORRA", REGAF&, "NOMESS")
   Call PoneEnCeroNull("StockSugXSuc", "NUCLIEN", REGAF&, "NOMESS")
'   Call ACTUREGISTRO("StockSugXSuc", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
   '
   '
   '
   NC = XVALO(Me.TXTCLIENTE)
   FIN& = msf2.Rows - 1
'   Dim RS As ADODB.Recordset
'   Set RS = New ADODB.Recordset
   FLEXCONN.BeginTrans
   On Error GoTo ERROR_GUARDAR
   NUSER1% = USNBR% Mod 100
   For I& = 1 To msf2.Rows - 1
      CI1% = XVALO(msf2.TextMatrix(I&, 1))
      SQLX$ = "Select * from StockSugXSuc  "
      SQLX$ = SQLX$ + " WHERE NUCLIEN = " & NC
      SQLX$ = SQLX$ + " AND CODINT = " & CI1%
      Dim RS As ADODB.Recordset
      Set RS = New ADODB.Recordset

      RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      '
      With RS
         If .EOF Then
           .AddNew
         End If
         
         !FechaUltEnvio = FETEXCOR1$(HOY(HO$))
         !USUARIO = NUSER1%
         !FECHACALCULO = FETEXCOR1$(FECHANUM(Me.TXTFECHACALCULO)) 'GUARDO CON QUE FECHA DE CALCULO SE GUARDO
         !FechaEmision = FETEXCOR1$(HOY(HO$))
         !StockMinimoActualEnSuc = XVALO(msf2.TextMatrix(I&, 6)) 'XVALO(MSF2.TextMatrix(I&, 6))
         !StockMaximoActualEnSuc = XVALO(msf2.TextMatrix(I&, 7)) 'XVALO(MSF2.TextMatrix(I&, 7))
         !StockMinimoSugerido = XVALO(msf2.TextMatrix(I&, 6))
         !StockMaximoSugerido = XVALO(msf2.TextMatrix(I&, 7))
         'VALIDACION PARA GUARDAR LA ULTIMA FECHA DE MODIFICACION.
         If Abs(XVALO(msf2.TextMatrix(I&, 4)) - XVALO(msf2.TextMatrix(I&, 6)) > 0.05 Or XVALO(msf2.TextMatrix(I&, 5)) - XVALO(msf2.TextMatrix(I&, 7))) > 0.05 Then
            !FechaModifItem = FETEXCOR1$(HOY(HO$))
         Else
            If FECHANUM(msf2.TextMatrix(I&, 8)) > 0 Then
              !FechaModifItem = FETEXCOR1$(FECHANUM(msf2.TextMatrix(I&, 8)))
            Else
              !FechaModifItem = FETEXCOR1$(HOY(HO$))
            End If
         End If
         !MARBORRA = 0
         .Update
         .MoveNext
     End With
     RS.Close
     Set RS = Nothing
   Next I&
   '
   
ERROR_GUARDAR:
    If Err <> 0 Then
      FLEXCONN.RollbackTrans
      Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
      Command3.Enabled = True
      Call DESBLOQUEAR_CONTROLES
      Exit Sub
    Else
      Call GENERAR_ARCHIVO_TRANSFERENCIA_STOCK_MIN_MAX(XVALO(Me.TXTCLIENTE))
      FLEXCONN.CommitTrans
      Screen.MousePointer = 1
      Call COMUNI("Proceso Completo")
      Command3.Enabled = True
      Call DESBLOQUEAR_CONTROLES
    End If
   Screen.MousePointer = 1
   Call DESBLOQUEAR_CONTROLES
End Sub

Sub DESBLOQUEAR_CONTROLES()
    revertirseleccion.Enabled = True
    ultimavistaguardada.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Command2.Enabled = True
    Command1.Enabled = True
End Sub
Sub BLOQUEAR_CONTROLES()
    revertirseleccion.Enabled = False
    ultimavistaguardada.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command2.Enabled = False
    Command1.Enabled = False
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Screen.MousePointer = 11
   Call BLOQUEAR_CONTROLES
   Call GRABASTOCKSUGERIDO
   Command3.Enabled = True
   Screen.MousePointer = 1
   Call DESBLOQUEAR_CONTROLES
End Sub

Sub GRABASTOCKSUGERIDO()

   '
   If GRID.Visible = False Then
      Call COMUNI("Imposible Grabar Sobre Grilla Filtrada Vuelva a Grilla Completa y Vuelva a Intentar")
      Exit Sub
   End If
   
   CONDI$ = " nuclien = " & XVALO(Me.TXTCLIENTE)
   CONDI$ = CONDI$ + " AND MARBORRA = 0 "
   Call PoneEnCeroNull("StockSugXSuc", "MARBORRA", REGAF&, "NOMESS")
   Call PoneEnCeroNull("StockSugXSuc", "NUCLIEN", REGAF&, "NOMESS")
   Call ACTUREGISTRO("StockSugXSuc", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
   '
   NUSER1% = USNBR% Mod 100

   FLEXCONN.BeginTrans ' NO LO PONGO MAS ARRIBA SI NO CUANDO COMPARA PARA BORRAR LOS REGISTROS QUE ESTEN EN 1 NO LOS ENCUENTRA POR QUE NO SE REALIZO LA TRANSACCION
   FLEXCONN.CommitTrans
   FLEXCONN.BeginTrans
   On Error GoTo ERROR_GUARDAR
   
   FIN& = GRID.Rows - 1
   SQLX$ = "Select * from StockSugXSuc  WHERE NUCLIEN < 1"
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset
   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   I& = 0
   For I& = 1 To GRID.Rows - 1
     With RS
         .AddNew
         !NUCLIEN = XVALO(Me.TXTCLIENTE)
         !FechaUltEnvio = CVDAT(CVINT(Me.TXTULTIMOENVIO))
         !USUARIO = NUSER1%
         !FECHACALCULO = CVDAT(CVINT(Me.TXTFECHACALCULO))
         !FechaEmision = HOY(HO$)
         !CODINT = XVALO(GRID.TextMatrix(I&, 1))
         !COD_EXTE = GRID.TextMatrix(I&, 2)
         !Descripcion = GRID.TextMatrix(I&, 3)
         !StockMinimoActualEnSuc = XVALO(GRID.TextMatrix(I&, 4))
         !StockMaximoActualEnSuc = XVALO(GRID.TextMatrix(I&, 5))
         !StockMinimoSugerido = XVALO(GRID.TextMatrix(I&, 6))
         !StockMaximoSugerido = XVALO(GRID.TextMatrix(I&, 7))
         .Update
     End With
   Next I&
   '
   Screen.MousePointer = 1
   CONDI$ = " NUCLIEN = " & XVALO(Me.TXTCLIENTE)
   CONDI$ = CONDI$ + " AND MARBORRA = 1"
   Call BORRAREGISTROS("StockSugXSuc", CONDI$, REGAF&, "NOMESS")

ERROR_GUARDAR:
    If Err <> 0 Then
      FLEXCONN.RollbackTrans
      MsgBox "ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
      Command3.Enabled = True
      Exit Sub
    Else
      FLEXCONN.CommitTrans
      RS.Close
      Set RS = Nothing
      Call COMUNI("Proceso Completo")
      Command3.Enabled = True
    End If


End Sub

Private Sub EDITMINMAX_Click()

End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call BLOQUEAR_CONTROLES
    'SETEO EL MSF QUE ESTE VISIBLE
    msf2.Redraw = False
    Dim MSFACTIVO As MSFlexGrid
    Dim MSFOCULTO As MSFlexGrid
    If GRID.Visible = True Then
      Set MSFACTIVO = GRID
      Set MSFOCULTO = MSF_2
    Else
      Set MSFACTIVO = MSF_2
      Set MSFOCULTO = GRID
    End If
    Screen.MousePointer = 11
    FIN& = MSFACTIVO.Rows - 1
'    Do While MSFACTIVO.Rows > 1
    For REG& = 1 To MSFACTIVO.Rows - 1
'      REG& = REG& + 1
      If MSFACTIVO.Rows = REG& Then Exit For
      k& = k& + 1
      If MSFACTIVO.Rows <= 1 Then Exit For
      Call TRACE(20, k&, FIN&)
      If MSFACTIVO.TextMatrix(REG&, 0) = "*" Then
        
        If VALIDARREPETICION(MSFACTIVO.TextMatrix(REG&, 1)) > 0 Then
           GoTo 50
        End If
        msf2.Rows = msf2.Rows + 1
        msf2.TextMatrix(msf2.Rows - 1, 1) = MSFACTIVO.TextMatrix(REG&, 1)
        msf2.TextMatrix(msf2.Rows - 1, 2) = MSFACTIVO.TextMatrix(REG&, 2)
        msf2.TextMatrix(msf2.Rows - 1, 3) = MSFACTIVO.TextMatrix(REG&, 3)
        msf2.TextMatrix(msf2.Rows - 1, 4) = MSFACTIVO.TextMatrix(REG&, 4)
        msf2.TextMatrix(msf2.Rows - 1, 5) = MSFACTIVO.TextMatrix(REG&, 5)
        msf2.TextMatrix(msf2.Rows - 1, 6) = MSFACTIVO.TextMatrix(REG&, 6): msf2.Row = msf2.Rows - 1: msf2.COL = 6: msf2.CellBackColor = &HFFFFFF
        msf2.TextMatrix(msf2.Rows - 1, 7) = MSFACTIVO.TextMatrix(REG&, 7): msf2.Row = msf2.Rows - 1: msf2.COL = 7: msf2.CellBackColor = &HFFFFFF
        msf2.TextMatrix(msf2.Rows - 1, 8) = MSFACTIVO.TextMatrix(REG&, 8)
50      'borro siempre primero el oculto
        If MSFOCULTO.Rows > 2 Then
          Call BORRARFILA(MSFACTIVO.TextMatrix(REG&, 1), MSFOCULTO) 'CUANDO ENTRA ACA EL MSFOCULTO SI TIENE MAS DE UN REGISTRO LO BORRO
        Else
           MSFOCULTO.Rows = 1
        End If
        'luego borro el activo
        If MSFACTIVO.Rows > 2 Then
          Call MSFACTIVO.RemoveItem(REG&)
        Else
          MSFACTIVO.Rows = 1 ' SI SE USA REMOVEITEM CON SOLO UNA FILA DA ERROR
        End If
        REG& = REG& - 1 ' HAY QUE RESTAR UNO POR QUE SE ELIMINA UN REGISTRO SI NO TERMINA SALTEANDO EL SIGUIENTE.
      End If
'    Loop
     
    Next REG&
    msf2.Redraw = True
    Screen.MousePointer = 1
    Command4.Enabled = True
    Me.TXTMOVIL.ToolTipText = ""
    Me.TXTMOVIL2.ToolTipText = ""
    Call DESBLOQUEAR_CONTROLES

End Sub
Sub BORRARFILA(CODIGO$, MSFOCULTO As MSFlexGrid)
    CI1% = XVALO(CODIGO$)
    For I& = 1 To MSFOCULTO.Rows - 1
       SD = XVALO(MSFOCULTO.TextMatrix(I&, 1))
       If CI1% = XVALO(MSFOCULTO.TextMatrix(I&, 1)) Then
         Call MSFOCULTO.RemoveItem(I&)
         Exit For
       End If
    Next I&
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   '
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   AREXPO$ = UCase$(CommonDialog1.FileName)
   If TRIM$(AREXPO$) <> "" Then
      If GRID.Visible = True Then
       Call GENEXCEL(AREXPO$, Me.GRID)
      Else
       Call GENEXCEL(AREXPO$, Me.MSF_2)
      End If
   End If
   '
99 Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Frame4.Visible = False
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   Screen.MousePointer = 11
   Command6.Enabled = False
   CANFIL& = Range("A" & Rows.Count).End(xlUp).Row
   FILA& = 0
   If msf2.Rows > 1 Then
      OPX% = COMALTER%("Grilla Inferior Posee Inforamación\\Cancelar\Sobre-Escribir\Agregar al final")
      If OPX% < 2 Then GoTo 99
      If OPX% = 3 Then
        FILA& = msf2.Rows - 1
      End If
   End If
   On Error Resume Next
   Call Combo1_Click
   On Error GoTo 0

   CANCOL& = 5   ' Hoja.Columns
   '
   '
   msf2.Redraw = False

   j& = 1
   For I& = 1 To CANFIL&
         VALOR = Hoja.Cells(I&, 1)
         valor1$ = SAFETEXT$(VALOR)
         If CODINT%(TRIM$(valor1$)) > 0 Then
            FILA& = FILA& + 1
            For j& = 1 To CANCOL&
              msf2.Rows = FILA& + 1

              valor1$ = Hoja.Cells(I&, j&)
              Label9.Font = msf2.Font
              Label9.FontSize = msf2.CellFontSize
              Label9.FontBold = msf2.CellFontBold
              Label9.Caption = valor1$

              Select Case j&
                Case 1, 2
                 msf2.TextMatrix(FILA&, j& + 1) = valor1$
                 If Label9.Width > msf2.ColWidth(j& + 1) Then
                  msf2.ColWidth(j& + 1) = Label9.Width
                 End If
                Case 3
                 msf2.TextMatrix(FILA&, 6) = valor1$
                 If Label9.Width > msf2.ColWidth(6) Then
                  msf2.ColWidth(6) = Label9.Width
                 End If

                Case 4
                 msf2.TextMatrix(FILA&, 7) = valor1$
                 If Label9.Width > msf2.ColWidth(7) Then
                  msf2.ColWidth(7) = Label9.Width
                 End If

              End Select
              
              

            Next j&
         End If
     Next I&

99 Screen.MousePointer = 1
   '
   Call COMUNI("Se Han Descargado Los Datos de La Planilla Excel ")
   '
   On Error Resume Next
   Set Hoja = Nothing
   Libro.Close: Set Libro = Nothing
   objexcel.Close: Set objexcel = Nothing
   On Error GoTo 0
   msf2.Redraw = True

   Frame4.Visible = False
   Command6.Enabled = True
   Command7.Enabled = True
End Sub

Private Sub Form_Load()
   Call CARGA_ENCABEZADO(EDITMINMAX.GRID, "CODINT/F0;Código/A20;Descripción/A40;St.Min.Suc./F10.1;St.Max.Suc/F10.1;St.Min.Sug/F10.1;St.Max.Sug./F10.1;Ult.Mod./D8", EDITMINMAX, 1, , ANTOT, , 1)
   Call CARGA_ENCABEZADO(EDITMINMAX.msf2, "CODINT/F0;Código/A20;Descripción/A40;St.Min.Suc./F10.1;St.Max.Suc/F10.1;St.Min.Sug/F10.1;St.Max.Sug./F10.1;Ult.Mod./D8", EDITMINMAX, , , ANTOT)
   Call CARGA_ENCABEZADO(EDITMINMAX.MSF_2, "CODINT/F0;Código/A20;Descripción/A40;St.Min.Suc./F10.1;St.Max.Suc/F10.1;St.Min.Sug/F10.1;St.Max.Sug./F10.1;Ult.Mod./D8", EDITMINMAX, , , ANTOT)

'   For i = 1 To 100
'      GRID.Rows = i + 1
'      GRID.TextMatrix(i, 1) = "AAAA" & i
'      GRID.TextMatrix(i, 2) = "BBBB" & i
'      GRID.TextMatrix(i, 3) = "CCCC" & i
'      GRID.TextMatrix(i, 4) = "DDDD" & i
'   Next i
   
'   'EN LA FUNCION DE CARGAR LOS ENCABEZADOS HABRIA QUE ASIGNAR LOS ANCHOS DE LABEL 5 Y ASIGNARSELO A LABEL8
'   For U& = 1 To GRID.Cols
'          JI& = JI& + 1
'          ' PROVISORIAMENTE A VER SI FUNCIONA
'          'AQUI AGREGAR A LABEL5 LA CANTIDAD DE CARACTERES
'          ANCH& = 20
'          Label5.Caption = FORMATNUM$("9999999.9999999", "A25" + "*")
'          GRID.ColWidth(JI& - 1) = Label5.Width
'25        If JI& >= Label8.Count Then
'             NUPIC% = Label8.Count
'             Load Label8(NUPIC%)
'             GoTo 25
'          End If
'          Label8(U&) = Label5.Width
'   Next U&
'   '
'   Label7 = 100
   '
End Sub

Private Sub GRID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
     REG& = GRID.MouseRow
     COL% = GRID.MouseCol
     Call MENU_CELDA(GRID, REG&, COL%)
  End If
End Sub

Private Sub impostockminmax_Click()
   '
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents

   AREXPO$ = UCase$(CommonDialog1.FileName)
   
   Frame4.Visible = True
   Frame4.Height = 1050
   Frame4.Left = 50
   Frame4.Top = 5280
   Frame4.Width = 145005
   Label3.Caption = AREXPO$
   Frame4.Visible = True
   
   
   '
99

End Sub
Sub IMPORTAEXCEL(AREXPO$)
   
End Sub

Private Sub Label3_Change()
  msf2.Clear
  msf2.Rows = 1
  Combo1.Clear
  Combo1.TEXT = ""
  '
  If Label3.Caption <> "" Then
     'Dim objexcel As Excel.Application
     Set objexcel = New Excel.Application
     Set Libro = objexcel.Workbooks.Open(Label3.Caption)
     '
     CANOJ% = Libro.Sheets.Count
     For KKI% = 1 To CANOJ%
        Combo1.AddItem Libro.Sheets(KKI%).Name
        If KKI% = 1 Then
           Combo1.TEXT = Libro.Sheets(KKI%).Name
        End If
     Next KKI%
     On Error Resume Next
     Call Combo1_Click
     On Error GoTo 0
  End If

End Sub

Private Sub revertirseleccion_Click()
        'SETEO EL MSF QUE ESTE VISIBLE
        Dim MSFACTIVO As MSFlexGrid
        If GRID.Visible = True Then
          Set MSFACTIVO = GRID
        Else
          Set MSFACTIVO = MSF_2
        End If
        '
        Call BLOQUEAR_CONTROLES
        HaySeleccion% = 0
        Screen.MousePointer = 11
        FIN& = MSFACTIVO.Rows - 1
        For I& = 1 To MSFACTIVO.Rows - 1
           Call TRACE(20, I&, FIN&)
           Call COLOREAR_GRILLA_SOLAFILA(MSFACTIVO, &HFFFF80, I&, 0, 0, 0)
        Next I&
        Screen.MousePointer = 1
        Call DESBLOQUEAR_CONTROLES
        '
End Sub

Private Sub TXTCLIENTE_Change()
   MSF_2.Rows = 1
End Sub

Private Sub TXTMOVIL2_LostFocus()
     If TXTMOVIL2.Visible = True Then
       X = TEXTOLOTES$(TXTMOVIL2.ToolTipText, ";")
       On Error Resume Next
       REG& = XVALO(CADENATEX$(1))
       COL& = XVALO(CADENATEX$(2))
       On Error GoTo 0

       REG& = XVALO(TXTMOVIL.ToolTipText)
       GRID.TextMatrix(REG&, COL&) = FORMATNUM$(XVALO(TXTMOVIL2), "F10.1")
       TXTMOVIL2.Visible = False
     End If
End Sub


Private Sub GRID_DblClick()
'   If (GRID.MouseRow = 0) Then Exit Sub
'   REG& = GRID.MouseRow
'   If GRID.Rows <= 1 Then Exit Sub
'   MSF2.Rows = MSF2.Rows + 1
'   MSF2.TextMatrix(MSF2.Rows - 1, 1) = GRID.TextMatrix(REG&, 1)
'   MSF2.TextMatrix(MSF2.Rows - 1, 2) = GRID.TextMatrix(REG&, 2)
'   MSF2.TextMatrix(MSF2.Rows - 1, 3) = GRID.TextMatrix(REG&, 3)
'   MSF2.TextMatrix(MSF2.Rows - 1, 4) = GRID.TextMatrix(REG&, 4)
'   MSF2.TextMatrix(MSF2.Rows - 1, 5) = GRID.TextMatrix(REG&, 5)
'   MSF2.TextMatrix(MSF2.Rows - 1, 6) = GRID.TextMatrix(REG&, 6)
'      If REG& > 0 Then MSF2.Row = MSF2.Rows - 1: MSF2.COL = 6: MSF2.CellBackColor = &HFFFFFF
'   MSF2.TextMatrix(MSF2.Rows - 1, 7) = GRID.TextMatrix(REG&, 7)
'      If REG& > 0 Then MSF2.Row = MSF2.Rows - 1: MSF2.COL = 7: MSF2.CellBackColor = &HFFFFFF
'   MSF2.TextMatrix(MSF2.Rows - 1, 8) = GRID.TextMatrix(REG&, 8)
'   'MSF2.TextMatrix(MSF2.Rows - 1, 9) = GRID.TextMatrix(REG&, 9)
'50 If GRID.Rows > 2 Then
'          Call GRID.RemoveItem(REG&)
'   Else
'          GRID.Rows = 1
'   End If
'   '
'   If MSF2.Rows > 1 Then
'    MSF2.COL = 7
'    MSF2.Row = MSF2.Rows - 1
'    With TXTMOVIL2
'      .ZOrder 0
'      .Width = MSF2.CellWidth
'      .Left = MSF2.CellLeft + MSF2.Left
'      .Top = MSF2.CellTop + MSF2.Top
'      .TEXT = MSF2.TEXT
'      .Visible = True
'      .ToolTipText = MSF2.Row
'      .SetFocus
'    End With
'   End If


End Sub
Private Sub MSF_2_DblClick()
'   If (MSF_2.MouseRow = 0) Then Exit Sub
'   REG& = MSF_2.MouseRow
'
'   MSF2.Rows = MSF2.Rows + 1
'   MSF2.TextMatrix(MSF2.Rows - 1, 1) = MSF_2.TextMatrix(REG&, 1)
'   MSF2.TextMatrix(MSF2.Rows - 1, 2) = MSF_2.TextMatrix(REG&, 2)
'   MSF2.TextMatrix(MSF2.Rows - 1, 3) = MSF_2.TextMatrix(REG&, 3)
'   MSF2.TextMatrix(MSF2.Rows - 1, 4) = MSF_2.TextMatrix(REG&, 4)
'   MSF2.TextMatrix(MSF2.Rows - 1, 5) = MSF_2.TextMatrix(REG&, 5)
'   MSF2.TextMatrix(MSF2.Rows - 1, 6) = MSF_2.TextMatrix(REG&, 6): MSF2.Row = MSF2.Rows - 1: MSF2.COL = 6: MSF2.CellBackColor = &HFFFFFF
'   MSF2.TextMatrix(MSF2.Rows - 1, 7) = MSF_2.TextMatrix(REG&, 7): MSF2.Row = MSF2.Rows - 1: MSF2.COL = 7: MSF2.CellBackColor = &HFFFFFF
'   MSF2.TextMatrix(MSF2.Rows - 1, 8) = MSF_2.TextMatrix(REG&, 8)
'   Call MSF_2.RemoveItem(REG&)
'
'   '
'   If MSF2.Rows > 1 Then
'    MSF2.COL = 7
'    MSF2.Row = MSF2.Rows - 1
'    With TXTMOVIL2
'      .ZOrder 0
'      .Width = MSF2.CellWidth
'      .Left = MSF2.CellLeft + MSF2.Left
'      .Top = MSF2.CellTop + MSF2.Top
'      .TEXT = MSF2.TEXT
'      .Visible = True
'      .ToolTipText = MSF2.Row
'      .SetFocus
'    End With
'   End If


End Sub


Private Sub MSF2_DblClick()
   If (msf2.MouseRow = 0) Then Exit Sub
   If GRID.Visible = False Then
      Call COMUNI("Debe Quitar Filtro Para Quitar La Fila Seleccionada")
      Exit Sub
   End If
   
   
   REG& = msf2.Row
   GRID.Rows = GRID.Rows + 1
   GRID.Cols = msf2.Cols
   GRID.TextMatrix(GRID.Rows - 1, 1) = msf2.TextMatrix(REG&, 1)
   GRID.TextMatrix(GRID.Rows - 1, 2) = msf2.TextMatrix(REG&, 2)
   GRID.TextMatrix(GRID.Rows - 1, 3) = msf2.TextMatrix(REG&, 3)
   GRID.TextMatrix(GRID.Rows - 1, 4) = msf2.TextMatrix(REG&, 4)
   GRID.TextMatrix(GRID.Rows - 1, 5) = msf2.TextMatrix(REG&, 5)
   GRID.TextMatrix(GRID.Rows - 1, 6) = msf2.TextMatrix(REG&, 6): GRID.Row = GRID.Rows - 1: GRID.COL = 6: GRID.CellBackColor = &HFFFFFF
   GRID.TextMatrix(GRID.Rows - 1, 7) = msf2.TextMatrix(REG&, 7): GRID.Row = GRID.Rows - 1: GRID.COL = 7: GRID.CellBackColor = &HFFFFFF
   GRID.TextMatrix(GRID.Rows - 1, 8) = msf2.TextMatrix(REG&, 8)
   Call borrafila(Me.msf2.Row, Me, Me.msf2)

End Sub

Private Sub TXTMOVIL_LostFocus()
        Dim MSFGRID As MSFlexGrid
        If GRID.Visible = True Then
          Set MSFGRID = GRID
        Else
          Set MSFGRID = MSF_2
        End If

        If TXTMOVIL.Visible = True Then
          X = TEXTOLOTES$(TXTMOVIL.ToolTipText, ";")
          On Error Resume Next
          REG& = XVALO(CADENATEX$(1))
          COL& = XVALO(CADENATEX$(2))
          On Error GoTo 0
    
          REG& = XVALO(TXTMOVIL.ToolTipText)
          MSFGRID.TextMatrix(REG&, COL&) = FORMATNUM$(XVALO(TXTMOVIL), "F10.1")
          TXTMOVIL.Visible = False
        End If
End Sub
Private Sub TXTMOVIL_KeyPress(KeyAscii As Integer)
   Dim MSFGRID As MSFlexGrid
   If GRID.Visible = True Then
     Set MSFGRID = GRID
   Else
     Set MSFGRID = MSF_2
   End If
   If KeyAscii = 13 Then
      X = TEXTOLOTES$(TXTMOVIL.ToolTipText, ";")
      On Error Resume Next
      REG& = XVALO(CADENATEX$(1))
      COL& = XVALO(CADENATEX$(2))
      MSFGRID.COL = COL&
      MSFGRID.Row = REG&
      On Error GoTo 0
      MSFGRID.TextMatrix(REG&, COL&) = FORMATNUM$(XVALO(TXTMOVIL), "F10.1")
      TXTMOVIL.Visible = False
'      If COL& = 6 Then
'        MSFGRID.COL = COL& + 1
''        TXTMOVIL.ToolTipText = TRIM$(Str$(REG&)) & ";" & TRIM$(Str$(COL& + 1))
'      ElseIf COL& = 7 Then
'        MSFGRID.Row = REG& + 1
''        TXTMOVIL.ToolTipText = TRIM$(Str$(REG& + 1)) & ";" & TRIM$(Str$(COL&))
'      End If
      TXTBUSCAR.SetFocus
'            TXTMOVIL.Top = GRID.CellTop + GRID.Top
'            TXTMOVIL.Left = GRID.CellLeft + GRID.Left
'            TXTMOVIL.Height = GRID.CellHeight - 10
'            TXTMOVIL.Width = GRID.CellWidth
'            TXTMOVIL.TEXT = TRIM$(GRID.TextMatrix(GRID.Row, GRID.COL))
'            TXTMOVIL.Visible = True
'            TXTMOVIL.ZOrder 0
'            TXTMOVIL.SetFocus
'''      With TXTMOVIL
''      .ZOrder 0
''      .Width = GRID.CellWidth
''      .Left = GRID.CellLeft + GRID.Left
''      .Top = GRID.CellTop + GRID.Top
''      .TEXT = GRID.TEXT
''      .Visible = True
''       REG& = GRID.Row
''      .ToolTipText = TRIM$(Str$(REG&)) & ";" & TRIM$(Str$(COL&))
''      End With
''      GRID.SetFocus
''
''      Call GRID_Click
''      On Error Resume Next
''      Command1.SetFocus
''      On Error GoTo 0
   End If
End Sub

Private Sub TXTMOVIL2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      X = TEXTOLOTES$(TXTMOVIL2.ToolTipText, ";")
      On Error Resume Next
      REG& = XVALO(CADENATEX$(1))
      COL& = XVALO(CADENATEX$(2))
      On Error GoTo 0
      msf2.TextMatrix(REG&, COL&) = FORMATNUM$(XVALO(TXTMOVIL2), "F10.1")
      TXTMOVIL2.Visible = False
''      If COL& = 6 Then
''        MSF2.COL = COL& + 1
'''        TXTMOVIL2.ToolTipText = TRIM$(Str$(REG&)) & ";" & TRIM$(Str$(COL& + 1))
''      ElseIf COL& = 7 Then
''        MSF2.Row = REG& + 1
'''        TXTMOVIL2.ToolTipText = TRIM$(Str$(REG& + 1)) & ";" & TRIM$(Str$(COL&))
''      End If
      TXTBUSCAR.SetFocus
'      With TXTMOVIL2
'      .ZOrder 0
'      .Width = GRID.CellWidth
'      .Left = GRID.CellLeft + GRID.Left
'      .Top = GRID.CellTop + GRID.Top
'      .TEXT = GRID.TEXT
'      .Visible = True
'       REG& = GRID.Row
'      .ToolTipText = TRIM$(Str$(REG&)) & ";" & TRIM$(Str$(COL&))
'      End With
'      GRID.SetFocus
'
'      Call GRID_Click
'      On Error Resume Next
'      Command1.SetFocus
'      On Error GoTo 0
   End If
End Sub

Private Sub TXTMOVIL_GotFocus()
     Call PINTATEXT(TXTMOVIL)
End Sub

Private Sub TXTMOVIL2_GotFocus()
     Call PINTATEXT(TXTMOVIL2)
End Sub

Private Sub Label13_Change()
   Call ACTUTOTA2
End Sub

Private Sub Label5_Click()
   Call ACTUTOTA2
End Sub


Private Sub Label9_Change()
    Call ACTUTOTA1
End Sub
Sub ACTUTOTA1()
   Static COTIHOY#
   If COTIHOY# = 0 Then
      COTIHOY# = TICAMONE(2)
   End If
   If COTIHOY# = 0 Then COTIHOY# = 1
   Label14 = FORMATNUM$((XVALO(Label9) * COTIHOY#) + XVALO(Label8), "F12.2")
   Label17 = FORMATNUM$((XVALO(Label13) * COTIHOY#) + XVALO(Label5), "F12.2")
End Sub
Sub ACTUTOTA2()
   Static COTIHOY#
   If COTIHOY# = 0 Then
      COTIHOY# = TICAMONE(2)
   End If
   If COTIHOY# = 0 Then COTIHOY# = 1
   Label17 = FORMATNUM$((XVALO(Label13) * COTIHOY#) - XVALO(Label5), "F12.2")
   Label14 = FORMATNUM$((XVALO(Label9) * COTIHOY#) + XVALO(Label8), "F12.2")
End Sub

Private Sub GRID_Click()
    Static MODO  As Boolean
    REG& = GRID.MouseRow
    
    If (REG& = 0) Then
        ' Ordena en forma ascendente
        If MODO Then
            GRID.COL = GRID.MouseCol
            GRID.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            GRID.COL = GRID.MouseCol
            GRID.Sort = 1
            MODO = True
        End If
        Exit Sub
    End If
    
    On Error Resume Next
    If REG& > 0 And GRID.Rows > 1 Then
        COLUM = GRID.MouseCol
        If COLUM <> 6 And COLUM <> 7 Then
          TXTMOVIL.Visible = False
          Call COLOREAR_GRILLA_SOLAFILA(GRID, &HFFFF80, GRID.Row, 0, 0, 0)
          GRID.Row = REG&: GRID.COL = 6: GRID.CellBackColor = &HFFFFFF
          GRID.Row = REG&: GRID.COL = 7: GRID.CellBackColor = &HFFFFFF
          Exit Sub
        End If
        GRID.COL = COLUM
        With TXTMOVIL
          .ZOrder 0
          .Width = GRID.CellWidth
          .Left = GRID.CellLeft + GRID.Left
          .Top = GRID.CellTop + GRID.Top
          .TEXT = GRID.TEXT
          .Visible = True
          REG& = GRID.Row
          .ToolTipText = TRIM$(Str$(REG&)) & ";" & TRIM$(Str$(COLUM))
          
          .SetFocus
        End With
   End If
   On Error GoTo 0
End Sub



Private Sub MSF_2_Click()
    Static MODO  As Boolean
    If (MSF_2.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            MSF_2.COL = MSF_2.MouseCol
            MSF_2.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSF_2.COL = MSF_2.MouseCol
            MSF_2.Sort = 1
            MODO = True
        End If
    End If
    
    COLUM = MSF_2.MouseCol
    REG& = MSF_2.MouseRow
    If COLUM <> 6 And COLUM <> 7 Then
      TXTMOVIL.Visible = False
      Call COLOREAR_GRILLA_SOLAFILA(MSF_2, &HFFFF80, REG&, 0, 0, 0)
      MSF_2.Row = REG&: MSF_2.COL = 6: MSF_2.CellBackColor = &HFFFFFF
      MSF_2.Row = REG&: MSF_2.COL = 7: MSF_2.CellBackColor = &HFFFFFF
      Exit Sub
    End If
    MSF_2.COL = COLUM
    With TXTMOVIL
      .ZOrder 0
      .Width = MSF_2.CellWidth
      .Left = MSF_2.CellLeft + MSF_2.Left
      .Top = MSF_2.CellTop + MSF_2.Top
      .TEXT = MSF_2.TEXT
      .Visible = True
      REG& = MSF_2.Row
      .ToolTipText = TRIM$(Str$(REG&)) & ";" & TRIM$(Str$(COLUM))
      
      .SetFocus
    End With

    
    
End Sub

Private Sub MSF_2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    '
    REG& = MSF_2.MouseRow
    COL% = MSF_2.MouseCol
    If REG& = 0 Or REG& > MSF_2.Rows Then Exit Sub
    
    Call MENU_CELDA2(MSF_2, REG&, COL%)
    '
  End If
End Sub

Private Sub msf2_Click()
    Static MODO  As Boolean
    If (msf2.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            msf2.COL = msf2.MouseCol
            msf2.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            msf2.COL = msf2.MouseCol
            msf2.Sort = 1
            MODO = True
        End If
    End If
    COLUM = msf2.MouseCol
    If COLUM <> 6 And COLUM <> 7 Then
      TXTMOVIL2.Visible = False
      Exit Sub
    End If
    
    msf2.COL = COLUM
    With TXTMOVIL2
      .ZOrder 0
      .Width = msf2.CellWidth
      .Left = msf2.CellLeft + msf2.Left
      .Top = msf2.CellTop + msf2.Top
      .TEXT = msf2.TEXT
      .Visible = True
      REG& = msf2.Row
      .ToolTipText = TRIM$(Str$(REG&)) & ";" & TRIM$(Str$(COLUM))
      .SetFocus
    End With

End Sub

Private Sub Refrescalc_Click()
   NC = XVALO(Me.TXTCTADECALCULO)
   If XVALO(Me.TXTCLIENTE) <> NC Then
      Call COMUNI("Imposible Refrescar Grilla - Último Cálculo Generado a Cliente: " & RASOCLI$(NC))
      Exit Sub
   End If
   
   Call BLOQUEAR_CONTROLES
   'LEO EL ARCHIVO Y ASIGNO A UN VECTOR
   'VERIFICO SI EL ULTIMO CALCULO ES DE LA SUCURSAL ACTIVA
   Screen.MousePointer = 11
   FIN& = ULTREG&("STATSTOK")
   Dim STOCKMINIMOREC(32767)
   Dim STOCKMAXIMOREC(32767)
   For I& = 1 To ULTREG&("STATSTOK")
        REG$ = REGLEIDO$("STATSTOK", I&)
        CI1% = CODINT%(Left$(REG$, 15))
        RVEN$ = Mid$(REG$, 225, 16)
        RFAB$ = Mid$(REG$, 353, 16)
        RTOT$ = Mid$(REG$, 65, 16)
        '
        MINIMO = FORMATNUM$(CVS(Mid$(RTOT$, 1, 4)), "F10.1")
        MAXIMO = FORMATNUM$(CVS(Mid$(RTOT$, 5, 4)), "F10.1")
        PROMEDIO = FORMATNUM$(CVS(Mid$(RTOT$, 9, 4)), "F10.1")
        ESTIMADO = FORMATNUM$(CVS(Mid$(RTOT$, 13, 4)), "F10.1")
        '
        STOCKMINIMOREC(CI1%) = PROMEDIO
        STOCKMAXIMOREC(CI1%) = MAXIMO
   Next I&
   '
   JJ& = 0
   For I& = 1 To GRID.Rows - 1
         CI1% = XVALO(GRID.TextMatrix(I&, 1))
         STMI = STOCKMINIMOREC(CI1%)
         STMA = STOCKMAXIMOREC(CI1%)
         GRID.TextMatrix(I&, 6) = STMI
         GRID.TextMatrix(I&, 7) = STMA
    Next I&
   Screen.MousePointer = 1
   Call DESBLOQUEAR_CONTROLES
 End Sub
 


Private Sub TXTBUSCAR_Change()
    '
    If TRIM$(TXTBUSCAR) = "" Then
      MSF_2.Visible = False
      GRID.Visible = True
      Exit Sub
    Else
      MSF_2.Visible = True
      GRID.Visible = False
    End If
    '
    MSF_2.Rows = 2
    For I& = 1 To MSF_2.Cols - 1
       MSF_2.TextMatrix(1, I&) = ""
    Next I&
    REG& = 0
    MSF_2.Visible = True
    MSF_2.ZOrder 0
    MSF_2.Cols = GRID.Cols
    MSF_2.Redraw = False
    GRID.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For I& = 1 To GRID.Rows - 1
      For j& = 1 To GRID.Cols - 1
        If (InStr(UCase$(GRID.TextMatrix(I&, j&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(GRID.TextMatrix(0, j&))) Then
            REG& = REG& + 1
            MSF_2.Rows = REG& + 1
            For H& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(REG&, H&) = GRID.TextMatrix(I&, H&)
            Next H&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next j&
    Next I&
    '
    MSF_2.Redraw = True
    
    '
End Sub

Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      TXTBUSCAR.TEXT = ""
   End If
End Sub

Private Sub ultimavistaguardada_Click()
   'ESTE DATO SE CONFIGURA EN MODULO DE REMITO DE TRASLADO DE ARCON
   'DONDE SE ASOCIA EL CLIENTE Y EL DEPOSITO EJEMPLO CLIENTE NEUQUEN Y DEPOSITO ALMACEN NEUQUEN
   CLIENTE1 = CONTROL("RETRASLA", "SUC1CLIE")
   CLIENTE2 = CONTROL("RETRASLA", "SUC2CLIE")
   RASOCLIENTE1$ = TRIM$(RASOCLI$(XVALO(CLIENTE1)))
   RASOCLIENTE2$ = TRIM$(RASOCLI$(XVALO(CLIENTE2)))
   If XVALO(CLIENTE1) + XVALO(CLIENTE2) < 1 Then GoTo 10
   DEPORELA% = 0: CLIENTESEL = 0
   OPX% = COMALTER%("Modalidad de Cálculo: \\Todos \" & Left$(RASOCLIENTE1$, 16) & "\" & Left$(RASOCLIENTE2$, 16) & "")
   If OPX% < 1 Or OPX% > 3 Then Exit Sub
   '
   'COMPLETA CASILLEROS DE ENCABEZADO
   If OPX% = 1 Then Me.TXTCLIENTE = 0: Me.TXTRASOCLI = "TODOS": CLIENTESEL = 0
   If OPX% = 2 Then CLIENTESEL = XVALO(CLIENTE1): Me.TXTCLIENTE = CLIENTE1: Me.TXTRASOCLI = RASOCLIENTE1$
   If OPX% = 3 Then CLIENTESEL = XVALO(CLIENTE2): Me.TXTCLIENTE = CLIENTE2: Me.TXTRASOCLI = RASOCLIENTE2$
   '*****************************************************************************************************
   '
   NC = XVALO(TXTCLIENTE)
   'RECORRO EL MASTER Y MUESTRO A LA DERECHA LOS DATOS GUARDADOS EN LA TABLA StockSugXSuc QUE CONTIENE LAS CANTIDADES DE STOCK MIN Y MAXIMO
   'DE ESTE MODO SIEMPRE MUESTRA TODO EL MASTER Y ACTUALIZA CADA VEZ QUE GRABA LA LISTA COMPLETA DE ARTICULOS EN StockSugXSuc
10 Screen.MousePointer = 11
   Call BLOQUEAR_CONTROLES
   '
   FIN& = CantRegistros&("MASTER")
   JJ& = 0: VUELTA% = 0: I& = 0
   Dim CODIGOLEIDO%(32767)
   'LEER PRIMERO TODOS LOS DATOS DE TABLA STOCKSUGERIDO
   'PONER EN UNO EL VECTOR CODIGOLEIDO%(CI1%) 'TODO ESTO ES LO QUE SE GRABO CON DATOS
   'LEER EL MASTER Y SI CODIGOLEIDO < 1 AGREGAR 'DE ESTE MODO SIEMPRE ESTA TODO  EL MASTER EN LA GRILLA
   'ORDENAR POR DESCRIPCION
   
   'LIMPIO LA GRILLA INFERIOR
   msf2.Rows = 1
   '
   NUSER1% = USNBR% Mod 100
   SQLX$ = "SELECT M.CODINT AS CODMASTER ,M.CODIGO,M.Descripcion,S.StockMinimoActualEnSuc,S.StockMaximoActualEnSuc, S.StockMinimoSugerido,"
   SQLX$ = SQLX$ + " S.StockMaximoSugerido, S.FechaModifItem,S.CODINT AS CODMSTSUG, S.FechaUltEnvio,S.FechaEmision, S.Usuario"
   SQLX$ = SQLX$ + " FROM StockSugXSuc as S WITH(NOLOCK) INNER JOIN MASTER as M ON S.CODINT = M.CODINT"
   SQLX$ = SQLX$ + " WHERE M.CODINT > 0 AND NUCLIEN = " & NC
   SQLX$ = SQLX$ + "  ORDER BY M.DESCRIPCION "
   Set RS = READSET(SQLX$)
   GRID.Redraw = False
   With RS
     Do While Not .EOF
        Call TRACE(20, I&, FIN&)
        CI1% = XVALO(!CODMASTER) 'CODIGO DE TABLA MASTER
        If VUELTA% < 1 Then
              'SE LEE SOLO UNA VEZ POR QUE ES PARA EL ENCABEZADO
              Me.TXTULTIMOENVIO = !FechaUltEnvio
              FEMI% = CVINT(!FechaEmision)
              US% = (XVALO(!USUARIO))
              Me.TXTUSUARIO = USERNAM$(NUSER1%)
              VUELTA% = 1
        End If
            
        CI1% = XVALO(!CODMASTER)
        CODIGOLEIDO%(CI1%) = 1
        CODI$ = TRIM$(SAFETEXT$(!CODIGO))
        DESCR$ = TRIM$(SAFETEXT$(!Descripcion))
        STOCKMINACTSUC = XVALO(!StockMinimoActualEnSuc)
        STOCKMAXACTSUC = XVALO(!StockMaximoActualEnSuc)
        STOCKMINSUGERIDO = XVALO(!StockMinimoSugerido)
        STOCKMAXSUGERIDO = XVALO(!StockMaximoSugerido)
        ULTIMAMODIFICACION$ = FECHATEX$(CVINT(!FechaModifItem)): If Right(ULTIMAMODIFICACION$, 2) = "80" Then ULTIMAMODIFICACION$ = ""
        
        I& = I& + 1
        GRID.Rows = I& + 1
        GRID.TextMatrix(I&, 1) = CI1%
        GRID.TextMatrix(I&, 2) = CODI$
        GRID.TextMatrix(I&, 3) = DESCR$
        GRID.TextMatrix(I&, 4) = FORMATNUM$(STOCKMINACTSUC, "F10.1")
        GRID.TextMatrix(I&, 5) = FORMATNUM$(STOCKMAXACTSUC, "F10.1")
        GRID.TextMatrix(I&, 6) = FORMATNUM$(STOCKMINSUGERIDO, "F10.1")
        GRID.TextMatrix(I&, 7) = FORMATNUM$(STOCKMAXSUGERIDO, "F10.1")
        GRID.TextMatrix(I&, 8) = ULTIMAMODIFICACION$
        
        .MoveNext
            
     Loop
   End With
   RS.Close
   Set RS = Nothing
   '
   'LEO EL MASTER Y AGREGO LOS CODIGOS QUE NO SE AGREGARON EN LA CONSULTA ANTERIOR (CONTROLO CON EL VECTOR CODIGOLEIDO)
   SQLX$ = "SELECT CODINT,CODIGO, DESCRIPCION"
   SQLX$ = SQLX$ + "  FROM MASTER M WITH(NOLOCK)  "
   SQLX$ = SQLX$ + " WHERE CODINT > 0 "
   SQLX$ = SQLX$ + "  ORDER BY DESCRIPCION "
   Set RS = READSET(SQLX$)
   '
   GRID.Redraw = False
   With RS
      Do While Not .EOF
        Call TRACE(20, I&, FIN&)
             '
        CI1% = XVALO(!CODINT)
        If CODIGOLEIDO%(CI1%) < 1 Then
            CODI$ = TRIM$(SAFETEXT$(!CODIGO))
            DESCR$ = TRIM$(SAFETEXT$(!Descripcion))
            STOCKMINACTSUC = 0
            STOCKMAXACTSUC = 0
            STOCKMINSUGERIDO = 0
            STOCKMAXSUGERIDO = 0
            ULTIMAMODIFICACION$ = ""
            '
            I& = I& + 1
            GRID.Rows = I& + 1
            GRID.TextMatrix(I&, 1) = CI1%
            GRID.TextMatrix(I&, 2) = CODI$
            GRID.TextMatrix(I&, 3) = DESCR$
            GRID.TextMatrix(I&, 4) = FORMATNUM$(STOCKMINACTSUC, "F10.1")
            GRID.TextMatrix(I&, 5) = FORMATNUM$(STOCKMAXACTSUC, "F10.1")
            GRID.TextMatrix(I&, 6) = FORMATNUM$(STOCKMINSUGERIDO, "F10.1")
            GRID.TextMatrix(I&, 7) = FORMATNUM$(STOCKMAXSUGERIDO, "F10.1")
            GRID.TextMatrix(I&, 8) = ULTIMAMODIFICACION$
         End If
        .MoveNext
      Loop
   End With
   RS.Close
   Set RS = Nothing
   Erase CODIGOLEIDO%
 
   
   'PONE EN COLOR BLANCO LAS COLUMNAS EDITABLES
   For I& = 1 To GRID.Rows - 1
    GRID.Row = I&: GRID.COL = 6: GRID.CellBackColor = &HFFFFFF
    GRID.Row = I&: GRID.COL = 7: GRID.CellBackColor = &HFFFFFF
   Next I&
   
   'ORDENO TODO EL GRID
   GRID.COL = 3
   GRID.Sort = 1
   GRID.Redraw = True
   '
   Screen.MousePointer = 1
   Call DESBLOQUEAR_CONTROLES
   
'   SQLX$ = "SELECT M.CODINT"
'   SQLX$ = SQLX$ + "  FROM MASTER M WITH(NOLOCK)  LEFT JOIN  StockSugXSuc S ON M.CODINT = S.CODINT "
'
'   SQLX$ = "SELECT DISTINCT(M.CODINT), M.CODIGO, M.DESCRIPCION, S.nuclien, S.StockMinimoActual, S.StockMaximoActual, S.StockMinimoActualEnSuc,S.CODINT AS CODINTSTOCK"
'   SQLX$ = SQLX$ + " ,S.StockMaximoActualEnSuc, S.StockMinimoSugerido, S.StockMaximoSugerido, S.FechaModifItem,"
'   SQLX$ = SQLX$ + " S.FechaEmision,S.FechaUltEnvio, S.USUARIO, S.FechaCalculo  FROM MASTER M WITH(NOLOCK)  LEFT JOIN  StockSugXSuc  S ON M.CODINT = S.CODINT "
'   SQLX$ = SQLX$ + " WHERE M.CODINT > 0"
'   SQLX$ = SQLX$ + "  ORDER BY NUCLIEN,M.DESCRIPCION "
'   Set RS = READSET(SQLX$)
'   '
'   GRID.Redraw = False
'   With RS
'     Do While Not .EOF
'       Call TRACE(20, I&, FIN&)
'       CI1% = XVALO(!CODINT) 'CODIGO DE TABLA MASTER
'       CI2% = XVALO(!CODINTSTOCK) 'CODIGO DE TABLA StockSugXSuc
'       NC = XVALO(!NUCLIEN)
'       If NC <> NCANTE Then
'          A = 1
'       End If
'       If CLIENTESEL = NC Then
'            STOCKMINIMOREC = 0
'            STOCKMAXIMO = 0
'            If VUELTA% < 1 Then
'              'SE LEE SOLO UNA VEZ POR QUE ES PARA EL ENCABEZADO
'              Me.TXTULTIMOENVIO = !FechaUltEnvio
'              FEMI% = CVINT(!FechaEmision)
'              US% = (XVALO(!Usuario))
'              Me.TXTUSUARIO = USERNAM$(US%)
'              VUELTA% = 1
'            End If
'
'            CI1% = XVALO(!CODINT)
'            CODI$ = TRIM$(SAFETEXT$(!CODIGO))
'            DESCR$ = TRIM$(SAFETEXT$(!Descripcion))
'            STOCKMINACTSUC = XVALO(!StockMinimoActualEnSuc)
'            STOCKMAXACTSUC = XVALO(!StockMaximoActualEnSuc)
'            STOCKMINSUGERIDO = XVALO(!StockMinimoSugerido)
'            STOCKMAXSUGERIDO = XVALO(!StockMaximoSugerido)
'            ULTIMAMODIFICACION$ = FECHATEX$(CVINT(!FechaModifItem)): If Right(ULTIMAMODIFICACION$, 2) = "80" Then ULTIMAMODIFICACION$ = ""
'
'            I& = I& + 1
'            GRID.Rows = I& + 1
'            GRID.TextMatrix(I&, 1) = CI1%
'            GRID.TextMatrix(I&, 2) = CODI$
'            GRID.TextMatrix(I&, 3) = DESCR$
'            GRID.TextMatrix(I&, 4) = FORMATNUM$(STOCKMINACTSUC, "F10.1")
'            GRID.TextMatrix(I&, 5) = FORMATNUM$(STOCKMAXACTSUC, "F10.1")
'            GRID.TextMatrix(I&, 6) = FORMATNUM$(STOCKMINSUGERIDO, "F10.1")
'            GRID.TextMatrix(I&, 7) = FORMATNUM$(STOCKMAXSUGERIDO, "F10.1")
'            GRID.TextMatrix(I&, 8) = ULTIMAMODIFICACION$
'        End If
'
'       .MoveNext
'
'     Loop
'   End With
'   RS.Close
'   Set RS = Nothing
'   'SI NO TENIA VISTA GUARDADA MUESTRA TODOS LOS DATOS
'   If I& = 0 Then
'        SQLX$ = "SELECT CODINT,CODIGO, DESCRIPCION"
'        SQLX$ = SQLX$ + "  FROM MASTER M WITH(NOLOCK)  "
'        SQLX$ = SQLX$ + " WHERE CODINT > 0 "
'        SQLX$ = SQLX$ + "  ORDER BY DESCRIPCION "
'        Set RS = READSET(SQLX$)
'        '
'        GRID.Redraw = False
'        With RS
'          Do While Not .EOF
'            Call TRACE(20, I&, FIN&)
'                 '
'            CI1% = XVALO(!CODINT)
'            CODI$ = TRIM$(SAFETEXT$(!CODIGO))
'            DESCR$ = TRIM$(SAFETEXT$(!Descripcion))
'            STOCKMINACTSUC = 0
'            STOCKMAXACTSUC = 0
'            STOCKMINSUGERIDO = 0
'            STOCKMAXSUGERIDO = 0
'            ULTIMAMODIFICACION$ = ""
'            '
'            I& = I& + 1
'            GRID.Rows = I& + 1
'            GRID.TextMatrix(I&, 1) = CI1%
'            GRID.TextMatrix(I&, 2) = CODI$
'            GRID.TextMatrix(I&, 3) = DESCR$
'            GRID.TextMatrix(I&, 4) = FORMATNUM$(STOCKMINACTSUC, "F10.1")
'            GRID.TextMatrix(I&, 5) = FORMATNUM$(STOCKMAXACTSUC, "F10.1")
'            GRID.TextMatrix(I&, 6) = FORMATNUM$(STOCKMINSUGERIDO, "F10.1")
'            GRID.TextMatrix(I&, 7) = FORMATNUM$(STOCKMAXSUGERIDO, "F10.1")
'            GRID.TextMatrix(I&, 8) = ULTIMAMODIFICACION$
'
'            .MoveNext
'          Loop
'        End With
'        RS.Close
'        Set RS = Nothing
'   End If
'   Screen.MousePointer = 1
'
'   'PONE EN COLOR BLANCO LAS COLUMNAS EDITABLES
'   For I& = 1 To GRID.Rows - 1
'    GRID.Row = I&: GRID.COL = 6: GRID.CellBackColor = &HFFFFFF
'    GRID.Row = I&: GRID.COL = 7: GRID.CellBackColor = &HFFFFFF
'   Next I&
'   GRID.Redraw = True
   '       1CODINT/F0;
'       2Código/F20;
'       3Descripción/A40;
'       4St.Min.Act./F9.1;
'       5St.Max.Act./F10.1;
'       6St.Min.Suc./F10.1;
'       7St.Max.Suc/F10.1;
'       8St.Min.Sug/F10.1;
'       9St.Max.Sug./F10.1;
'       10F.Envio/D8;
'       11F.Emsion/F0;
'       12Usuario/F0;
'       13FechaCalculo/F0


End Sub
Sub MENU_CELDA(MSF As MSFlexGrid, REG&, COL%)
        If MSF.Rows <= 1 Then Exit Sub
        
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        GLOBALFRM_FILA& = REG&  'CELDA SELECCIONADA
        PopupMenu mnuArchivo
        Call INVICTRL
        Call COLOREAR_GRILLA_SOLACELDA(MSF, MSF.BackColor, REG&, COL%)      ' LA PINTA DE AMARILLA
End Sub

Sub INVICTRL()
  '
  mnucopiar.Visible = Not (mnucopiar.Visible)
  mnuabrirgrafico.Visible = Not (mnuabrirgrafico.Visible)
  '
End Sub

Private Sub mnucopiar_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub
Private Sub mnuabrirgrafico_Click()
       REG& = GLOBALFRM_FILA&
       Screen.MousePointer = 11
       CODIGO$ = GRID.TextMatrix(REG&, 2)
       ANESTO07.MESAN = CODIGO$
       Screen.MousePointer = 1
       ANESTO07.Show 1
End Sub

Sub MENU_CELDA2(MSF As MSFlexGrid, REG&, COL%)
        If MSF.Rows <= 1 Then Exit Sub
        
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL2
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        GLOBALFRM_FILA& = REG&  'CELDA SELECCIONADA
        PopupMenu mnuArchivo2
        Call INVICTRL2
        Call COLOREAR_GRILLA_SOLACELDA(MSF, MSF.BackColor, REG&, COL%)      ' LA PINTA DE AMARILLA
End Sub

Sub INVICTRL2()
  '
  mnucopiar2.Visible = Not (mnucopiar.Visible)
  mnuabrirgrafico2.Visible = Not (mnuabrirgrafico2.Visible)
  '
End Sub

Private Sub mnucopiar2_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub
Private Sub mnuabrirgrafico2_Click()
       REG& = GLOBALFRM_FILA&
       Screen.MousePointer = 11
       CODIGO$ = MSF_2.TextMatrix(REG&, 2)
       ANESTO07.MESAN = CODIGO$
       Screen.MousePointer = 1
       ANESTO07.Show 1
End Sub

