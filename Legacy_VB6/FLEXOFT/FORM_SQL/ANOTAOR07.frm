VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ANOTAOR07 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anotador Ordenes de Fabricación"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Novedades"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Index           =   1
      Left            =   120
      TabIndex        =   28
      Top             =   6300
      Width           =   10635
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00C0C0C0&
         Height          =   425
         Index           =   0
         Left            =   30
         ScaleHeight     =   360
         ScaleWidth      =   10230
         TabIndex        =   31
         Top             =   315
         Width           =   10290
         Begin VB.Label Label44 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   10270
         End
         Begin VB.Label Label23 
            Caption         =   "Label5"
            Height          =   15
            Index           =   0
            Left            =   0
            TabIndex        =   32
            Top             =   420
            Width           =   1380
         End
         Begin VB.Line Line2 
            Index           =   0
            X1              =   0
            X2              =   1155
            Y1              =   420
            Y2              =   420
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1175
         Left            =   1260
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   720
         Width           =   6840
      End
      Begin VB.CommandButton Command30 
         Caption         =   "O.K"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   8240
         Picture         =   "ANOTAOR07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Agrega el Texo al Anotador"
         Top             =   1210
         Width           =   750
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   495
         Index           =   1
         Left            =   9000
         OleObjectBlob   =   "ANOTAOR07.frx":030A
         TabIndex        =   34
         Top             =   1320
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   300
         Index           =   0
         Left            =   8230
         OleObjectBlob   =   "ANOTAOR07.frx":0386
         TabIndex        =   35
         Top             =   800
         Width           =   2100
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   300
         Index           =   2
         Left            =   45
         OleObjectBlob   =   "ANOTAOR07.frx":03DC
         TabIndex        =   36
         Top             =   800
         Width           =   1155
      End
      Begin VB.Line Line25 
         X1              =   8180
         X2              =   8180
         Y1              =   735
         Y2              =   1900
      End
      Begin VB.Line Line22 
         X1              =   0
         X2              =   10665
         Y1              =   1900
         Y2              =   1900
      End
      Begin VB.Label Label57 
         BackColor       =   &H00000000&
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
         Height          =   90
         Index           =   0
         Left            =   60
         TabIndex        =   38
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label10"
         Height          =   105
         Left            =   8250
         TabIndex        =   37
         Top             =   1080
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   26
      Top             =   720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command6 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   9870
      Picture         =   "ANOTAOR07.frx":0432
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Consultas y Re-Impresión"
      Top             =   780
      Width           =   750
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Orden de Fabricación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1785
      Left            =   105
      TabIndex        =   6
      Top             =   -30
      Width           =   9570
      Begin VB.CommandButton Command22 
         Caption         =   "Desvincular Pedido de O.F. Anulada"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   480
         TabIndex        =   40
         Top             =   1400
         Width           =   3075
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   600
         Left            =   6960
         ScaleHeight     =   540
         ScaleWidth      =   2340
         TabIndex        =   23
         Top             =   720
         Width           =   2400
         Begin VB.CommandButton Command33 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   1800
            Picture         =   "ANOTAOR07.frx":073C
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Des-Asignar Armador (Sin Orden de Compra)"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command1 
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
            Height          =   600
            Left            =   0
            Picture         =   "ANOTAOR07.frx":0A46
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Asignar Armador Externo"
            Top             =   -10
            Width           =   560
         End
         Begin VB.Label Label19 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Asig. Armador Externo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   555
            TabIndex        =   25
            Top             =   0
            Width           =   1260
         End
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
         Height          =   300
         Left            =   2550
         TabIndex        =   9
         Top             =   315
         Width           =   195
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido:"
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
         Left            =   6200
         TabIndex        =   46
         Top             =   1380
         Width           =   675
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6980
         TabIndex        =   45
         ToolTipText     =   "Doble Click para Restaurar Situación de Preparacion y Despacho"
         Top             =   1320
         Width           =   1050
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Orden:"
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
         Left            =   8040
         TabIndex        =   42
         Top             =   1380
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8805
         TabIndex        =   41
         ToolTipText     =   "Doble Click para Restaurar Situación de Preparacion y Despacho"
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1155
         TabIndex        =   20
         ToolTipText     =   "Doble  Click Para Marcar Como Cerrada O.Fabr. Activa"
         Top             =   1050
         Width           =   330
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
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
         Left            =   -210
         TabIndex        =   19
         Top             =   1110
         Width           =   1275
      End
      Begin VB.Label Label16 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5460
         TabIndex        =   18
         Top             =   300
         Width           =   1170
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4815
         TabIndex        =   17
         Top             =   345
         Width           =   555
      End
      Begin VB.Label Label14 
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
         ForeColor       =   &H00000000&
         Height          =   600
         Left            =   1920
         TabIndex        =   15
         Top             =   735
         Width           =   4725
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6750
         TabIndex        =   14
         Top             =   345
         Width           =   720
      End
      Begin VB.Label Label11 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7560
         TabIndex        =   13
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2805
         TabIndex        =   12
         Top             =   375
         Width           =   675
      End
      Begin VB.Label Label7 
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1155
         TabIndex        =   11
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3570
         TabIndex        =   10
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
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
         Left            =   -105
         TabIndex        =   8
         Top             =   735
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
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
         Index           =   0
         Left            =   -840
         TabIndex        =   7
         Top             =   380
         Width           =   1905
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anotaciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4450
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   10620
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   4000
         Index           =   0
         Left            =   10290
         TabIndex        =   22
         Top             =   700
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   4000
         Left            =   8085
         TabIndex        =   5
         Top             =   700
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   4000
         Left            =   1155
         TabIndex        =   4
         Top             =   700
         Width           =   15
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3435
         Left            =   0
         TabIndex        =   3
         Top             =   735
         Width           =   10580
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   10515
         TabIndex        =   2
         Top             =   315
         Width           =   10580
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   10560
         End
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   10605
         Y1              =   4410
         Y2              =   4410
      End
      Begin VB.Line Line1 
         X1              =   10580
         X2              =   10580
         Y1              =   320
         Y2              =   6225
      End
   End
   Begin VB.CommandButton Command9 
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
      Height          =   720
      Left            =   9870
      Picture         =   "ANOTAOR07.frx":0D50
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Sale de la Aplicación"
      Top             =   60
      Width           =   750
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "ANOTAOR07.frx":0E9A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "ANOTAOR07.frx":10CE
      TabIndex        =   27
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
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
      Left            =   0
      TabIndex        =   44
      Top             =   60
      Width           =   1275
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1365
      TabIndex        =   43
      ToolTipText     =   "Doble Click para Restaurar Situación de Preparacion y Despacho"
      Top             =   0
      Width           =   330
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuMarcarCierreOf 
         Caption         =   "Marcar Cierre de Orden de Fabricaciòn Activa"
      End
   End
   Begin VB.Menu mnuCtalis 
      Caption         =   "Consultas y Listados"
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
   End
   Begin VB.Menu mnuUt 
      Caption         =   "Utilidades"
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuración"
   End
   Begin VB.Menu mnuacces 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuIngresoMatCorte 
         Caption         =   "Ingreso de Bobinas a Corte"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTransfLinea 
         Caption         =   "Transferencia a Linea"
         Visible         =   0   'False
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSitCierrOF 
         Caption         =   "Situación de Consumos de Cierre de Orden de Fabricación"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuinformarRetiCorte 
         Caption         =   "Informar Retiro de Cortes"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRepDiario 
         Caption         =   "Reporte Diario"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "ANOTAOR07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
       Command1.Enabled = False
       If InStr(1, EMPREAC$, "GABAL") > 0 Then
          Call COMUNI("Opción No Disponible en la Presente Versión")
       End If
       '
       If TRIM$(Label7) = "" Then
          Call COMUNI("Falta Seleccionar O.Fabricacion")
          GoTo 99
       End If
       '
       NORFA$ = TRIM$(Label7)
       '
       EPAC$ = TRIM$(UCase$(EMPREAC$))
       MODOGABAL% = 0: If InStr(EPAC$, "GABAL") > 0 Then MODOGABAL% = 1
       If MODOGABAL% > 0 Then
          Call GRABAARMADOR_GABAL(NORFA$, 0)
       Else
          Call GRABAARMADOR(NORFA$)
       End If
       '
       Label7 = "": Label7 = NORFA$

99     Command1.Enabled = True
End Sub



Private Sub Command2_Click()
    Command2.Enabled = False
    ' CODIGO REALIZADO PARA MEDITEA PARA QUE ACTUALICE LA FECHA DE ENTREGA SOLICITADA
    ' DE LAS ORDENES DE COMPRA POR ARMADO DE EQUIPO, EN BASE AL LEAD TIME DE LOS
    ' EQUIPOS
    Call ABRECONEXION
    SQLX$ = "SELECT COD_INTE, Femi_Ocom, Fentre_Sol FROM OCOMDETA WHERE STAT_OCOM<3 AND DOC_ORIG LIKE 'OF-%'"
    Dim Rst As ADODB.Recordset
    Set Rst = New ADODB.Recordset
25  On Error Resume Next
    Rst.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 25
    End If
    On Error GoTo 0
    With Rst
        Do While Not .EOF
            CI% = XVALO(!cod_inte)
            FEMI% = CVINT(!Femi_Ocom)
            LT% = LEAD_TIME%(CI%)
            !Fentre_Sol = CVDAT(DIASPOST(FEMI%, LT%))
            .Update
            '
            .MoveNext
        Loop
    End With
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    Call COMUNI("Proceso Finalizado")
    '
    Command2.Enabled = True
End Sub

Private Sub Command3_Click()
   '
   '\\\OT-08-0631-OD-12/08/08///
   '
   Command3.Enabled = False
   NumOrFab$ = Label7
   NPX& = XVALO(Mid$(Label7, 4, 6))
   If NPX& > 0 Then
     If TRIM$(Text1) <> "" Then
       Call FRATEXTO(Text1, 56)
       For KKI% = 1 To CARETEX%
         TTX1$ = Space$(76)
         Call REPLA(TTX1$, RETEX$(KKI%), 11, 56)
         If KKI% = 1 Then
           Call REPLA(TTX1$, HOS$, 1, 8)
           Call REPLA(TTX1$, USERTER$, 69, 24)
         End If
         TTXX$ = TTXX$ + TTX1$ + Chr$(10) + Chr$(13)
       Next KKI%
       
      Call APERBASDAT("PEDCLI")
      SQLX$ = "SELECT AnotaOrf FROM OrdeFabr"
      SQLX$ = SQLX$ + " WHERE NumeOrfa = " & NPX& & ""
      Dim OrdeFabTemp As ADODB.Recordset
      Set OrdeFabTemp = New ADODB.Recordset
      OrdeFabTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      On Error Resume Next
      With OrdeFabTemp
        .MoveFirst
        If Err < 1 Then
            On Error GoTo 0
            ANOTA$ = SAFETEXT(!AnotaOrf)
            ANOTA$ = ANOTA$ + TTXX$
            ANOTA$ = ANOTA$ + String$(88, "=") + Chr$(10) + Chr$(13)
            !AnotaOrf = ANOTA$
            .Update
        End If
      End With
      OrdeFabTemp.Close
      Set OrdeFabTemp = Nothing
      '
      Text1 = ""
      '
      List1.Clear
      'REFRESCO LA PANTALLA
      Label7 = "": Label7 = NumOrFab$
      '
     End If
   End If
   Command3.Enabled = True
   '\\\OT-08-0631-OD-FIN///
End Sub

Private Sub Command22_Click()
    Command22.Enabled = False
    NORFA$ = TRIM$(Label7)
    If NORFA$ = "" Then
       Call COMUNI("Falta Seleccionar Una Orden de Fabricación Activa")
       GoTo 99
    End If
    
    STAT% = XVALO(Label18)
    If STAT% <> 9 Then
       Call COMUNI("Debe Primero Anular La Orden de Fabricación")
       GoTo 99
    End If
    OPX% = COMALTER%("¿Esta Seguro Que Desea Desvincular La Orden '" & NORFA$ & "' del Pedido N°: " & Label9 & " Orden N°: " & Label1 & " ?\\No, Cancelar\Si, Continuar")
    If OPX% < 2 Then GoTo 99
   
    Call ABRECONEXION
    SQLX$ = "SELECT NuOrNroPed, NroPed FROM ORDEFABR WHERE IDSUBOR = '" & NORFA$ & "'"
    Dim Rst As ADODB.Recordset
    Set Rst = New ADODB.Recordset
25  On Error Resume Next
    Rst.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 25
    End If
    On Error GoTo 0
    With Rst
        Do While Not .EOF
            !NuOrNroPed = 0
            !NROPED = 0
            .Update
            .MoveNext
        Loop
    End With
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    Call COMUNI("Proceso Finalizado")
    NORFA$ = TRIM$(Label7)
    Label7 = ""
    Label7 = NORFA$
    '
99  Command22.Enabled = True
End Sub

Private Sub Command33_Click()
    Command33.Enabled = False
    NORFA$ = TRIM$(Label7)
    If NORFA$ = "" Then
       Call COMUNI("Falta Seleccionar Orden de Fabricación")
       GoTo 99
    End If
    '
    CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
    ARMAD% = XVALO(VALOBADA("ORDEFABR", "ARMADOR", CONDI$, "NOMESS"))
    If ARMAD% < 1 Then
       Call COMUNI("Orden de Fabricación Seleccionada No Tiene Asignado Armador")
       GoTo 99
    End If
    '
    CONDI$ = "Doc_Orig = '" & NORFA$ & "'"
    If CantRegistros&("OCOMDETA", CONDI$, "NOMESS") > 0 Then
       Call COMUNI("Imposible Quitar Armador !! \Orden de Fabricación Seleccionada \ Ya Posee Orden de Compra Asociada")
       GoTo 99
    End If
    
    SQLX$ = "SELECT Armador ,AnotaOrf FROM ORDEFABR "
    SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "'"
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
20  On Error Resume Next
    rs1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 20
    End If
    On Error GoTo 0
    
    CUERPO$ = "Des-Asignación de Armador Externo" & vbCrLf
    CUERPO$ = CUERPO$ + " (" & ARMAD% & " " & TRIM$(ECOARCH$("ARMADORS", MKI$(ARMAD%))) & " )"
    HISTO$ = FORMATEA_ANOTADOR$(CUERPO$)
    With rs1
      Do While Not .EOF
            !ARMADOR = 0
            !AnotaOrf = SAFETEXT$(!AnotaOrf) & HISTO$
            .Update
            .MoveNext
      Loop
    End With
    rs1.Close
    Set rs1 = Nothing
    Label7 = "": Label7 = NORFA$
     
99  Command33.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
3   Call SELECHO("ECORDEF/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
    NORFA$ = TRIM$(VALACT1$("ECORDEF"))
    If NORFA$ = "" Then GoTo 99
    '
    Label7 = TRIM$(NORFA$)
    '
99  Command4.Enabled = True
End Sub

Private Sub Command6_Click()
  
   '\\\OT-08-0631-OD-12/08/08///
   Command6.Enabled = False
   NPX& = Mid$(Label7, 4, 6)
   NORFA$ = TRIM$(Label7)
   '
   If NPX& > 0 Then
      TIDOCUM$ = "Orden de Fabricación " + NORFA$
'      If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
'        TIDOCUM$ = "Orden de Despacho + "
'      End If
      NUDOCU$ = TRIM$(str$(NPX&))
      While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
      NUDOCU$ = NUDOCU$ + " "
      Screen.MousePointer = 11
      '
      For UI& = ULTREG&("PDOCLST") To 1 Step -1
       XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
       If InStr(XX$, TIDOCUM$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 5
       End If
      Next UI&
      '
      Screen.MousePointer = 1
      Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
      GoTo 99
      '
5     Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
      
    End If
    '
99  Command6.Enabled = True
    '\\\OT-08-0631-OD-FIN///
End Sub

Private Sub Command9_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub Form_Load()
    Call CENTRAFORM(Me)
    '
    DoEvents
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    
    'fecha y usuario solapa de cobranza
    SkinLabel1(0).Caption = Left$(USERNAME, 17)
    HOII% = HOY(HOS$)
    SkinLabel1(2).Caption = " " + HOS$
    
    EPAC$ = TRIM$(UCase$(EMPREAC$))
   
    If InStr(EPAC$, "CONTEPLAST") > 0 Then
        mnuIngresoMatCorte.Visible = True
        mnuTransfLinea.Visible = True
        S.Visible = True
        mnuinformarRetiCorte.Visible = True
        mnuRepDiario.Visible = True
        mnuSitCierrOF.Visible = True
    End If
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    Screen.MousePointer = 1
'    HOII% = HOY(HOS$)
'    Label9 = " " + HOS$
'    Label12 = Left$(USERNAME, 17)
    '
End Sub


Private Sub Label18_Click()
   If XVALO(Label18.Caption) >= 3 Then
      Call COMUNI("La Orden de Fabricación Activa Yá se Encuentra Cerrada")
   Else
      NORFAX$ = Label18.Caption
      EPAC$ = UCase$(TRIM$(EMPREAC$))

      If NORFAX$ <> "" Then
        CONDI$ = "IdSubOr = '" & NORFAX$ & "'"
        FECHAHOY = FETEXCOR1$(HOY(HO$))
        SQLX$ = "UPDATE ORDEFABR SET  StatuOrf = 3 ,  FechCierr = '" & FECHAHOY & "'"
        SQLX$ = SQLX$ + " WHERE " & CONDI$
        FLEXCONN.Execute (SQLX$)
        If InStr(EPAC$, "CONTEPLAST") < 1 Or InStr(EPAC$, "CAMBACERES") < 1 Then 'en conteplast no borro las operaciones
           ' Cancela Operaciones Programadas   'por las dudas para que no queden huerfanas
           Call BORRAREGISTROS("OperFab", CONDI$, RAFE&, "NOMESS")
        End If
        ' Cancela Reservas de Componentes
        Call BORRAREGISTROS("Reserva", CONDI$, RAFE&, "NOMESS")
        ' Cancela Orden en Indice de Ordenes Pendientes
        NORFAX$ = AJUSTI$(NORFAX$, 12)
        Call FILTERDELETE("ECORDEP", 1, NORFAX$)
        Call FRESHECHO("ECORDEP")
      End If
      Call CIERRARCH("*.*")
      Call LIBARCH("ECORDEP")
      Call CIERRARCH("*.*")
      
      
      TTXYZ$ = Space$(76)
      HOII% = HOY(HOS$)
      Call REPLA(TTXYZ$, HOS$, 1, 8)
      Call REPLA(TTXYZ$, "Marca de Cierre Manual (Anotador)", 11, 56)
      Call REPLA(TTXYZ$, USERTER$, 69, 24)
      Call AGREGA_ANOTAORFAB(NORFAX$, TTXYZ$)

   End If
End Sub

Private Sub mnuinformarRetiCorte_Click()
    NORFA$ = TRIM$(Label7)
    If NORFA$ = "" Then Exit Sub

    INFOCORTE.Label4 = NORFA$

    INFOCORTE.Show 1

End Sub

Private Sub mnuIngresoMatCorte_Click()
    NORFA$ = TRIM$(Label7)
    If NORFA$ = "" Then Exit Sub

    INGREBOB.TXTOF = NORFA$
    INGREBOB.Show 1
End Sub


Private Sub mnuMarcarCierreOf_Click()
    If TRIM$(Label7) = "" Then Exit Sub
    If XVALO(Label18) = 3 Then
       Call COMUNI("La Orden de Fabricaciòn Yà se Encuentra Cerrada"): Exit Sub
    End If
    '
    If DERACCE%("CIEMASOF", "Cierre Automático de Ordenes de Fabricación") < 2 Then Exit Sub
    '
    Call CIEMASOF.MarcaCierreOFBorraReserva(TRIM$(Label7))
    TTXYZ$ = Space$(76)
    HOII% = HOY(HOS$)
    Call REPLA(TTXYZ$, HOS$, 1, 8)
    Call REPLA(TTXYZ$, " Marca de Cierre Manual", 11, 56)
    Call REPLA(TTXYZ$, USERTER$, 69, 24)
    Call AGREGA_ANOTAORFAB(Label7, TTXYZ$)
    Call Label7_Change

End Sub

Private Sub mnuRepDiario_Click()
   InfDiaSec.Label4 = TRIM$(Label7)
   InfDiaSec.Show 1

End Sub

Private Sub mnuSitCierrOF_Click()
   
   DetaCierreCtp.TXTACERRAR = 1
   
   'ESTE BLOQUE SIEMPRE AL FINAL
   DetaCierreCtp.Text10 = Label7
   DetaCierreCtp.Show 1

End Sub

Private Sub mnuTransfLinea_Click()
  TRAFALINEA.TXTOT = Label7
  TRAFALINEA.Show 1

End Sub

Private Sub Label7_Change()
     List1.Clear
     If TRIM$(Label7) = "" Then Exit Sub
     '
     NORFA$ = Mid$(Label7, 4, 6)
     '
     '\\\OT-08-0631-OD-12/08/08///

     Screen.MousePointer = 11
     Call ABREORFAB
     
     SQLX$ = " SELECT FechGen, CanProy,COD_INTE, StatuOrf, AnotaOrf,NuOrNroPed,NROPED FROM ORDEFABR WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE NumeOrFa = '" & NORFA$ & "' "
    ' SQLX$ = SQLX$ + " AND StatuOrf <> 9 "
    '
'     Dim OrdeFabTemp As ADODB.Recordset
'     Set OrdeFabTemp = New ADODB.Recordset
'     OrdeFabTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     Set OrdeFabTemp = READSET(SQLX$)
     With OrdeFabTemp
         On Error Resume Next
         .MoveFirst
         If Err Then
           On Error GoTo 0
           Screen.MousePointer = 1
           Call MENSERR(24, "Ningun Registro Coincide con Registro SQL")
           Exit Sub
         End If
        '
         i& = 0
         Do While Not .EOF
           FEGEN = CVINT(!FECHGEN)
           CATOT# = XVALO(!CANPROY)
           CIXI% = XVALO(!cod_inte)
           STAT% = XVALO(!STATUORF)
           NUORPED& = XVALO(!NuOrNroPed)
           NPED& = XVALO(!NROPED)
           XX$ = SAFETEXT$(!AnotaOrf): LONGI& = Len(XX$)
           .MoveNext
         Loop
        '
    End With
    Set OrdeFabTemp = Nothing
    OrdeFabTemp.Close
    '
    
1420 Label8 = FECHATEX$(FEGEN)
     Label16 = FORMATNUM(CATOT#, "F8.1")
     Label11 = CODEXT$(CIXI%)
     Label14 = DESCRIT$(CIXI%)
     Label18 = TRIM$(str$(STAT%))
     Label9 = NUORPED&
     Label1 = NPED&

     Inicio& = 1
     Do While Inicio& < LONGI&
       FIN& = InStr(Inicio&, XX$, Chr$(10))
       If FIN& > 0 Then
          TTXX$ = Mid$(XX$, Inicio&, FIN& - Inicio&)
          List1.AddItem TTXX$
          Inicio& = FIN& + 2
       Else
          TTXX$ = Mid$(XX$, Inicio&)
          List1.AddItem TTXX$
          Exit Do
       End If
     Loop
     Screen.MousePointer = 1
         '
    '\\\OT-08-0631-OD-12/08/08///


     '
'20   NPX& = Val(Mid$(NORFA$, 4, 6))
'     RFF$ = RECFILT$("ANOTAORF", 1, MKS$(NPX&))
'     For U& = 1 To Len(RFF$) Step 4
'       RENOTA& = CVS(Mid$(RFF$, U&, 4))
'       XX$ = REGLEIDO$("ANOTAORF", RENOTA&)
'       List1.AddItem Mid$(XX$, 5, 92)
'     Next U&
     '
End Sub

