VERSION 5.00
Begin VB.Form TRAZAUTO 
   BackColor       =   &H00FFFFD0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Trazabilidad"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
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
      Height          =   750
      Left            =   8505
      Picture         =   "TRAZAUTO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Graba - Guardar Modificaciones"
      Top             =   5190
      Width           =   800
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame3"
      Height          =   1170
      Left            =   4725
      TabIndex        =   33
      Top             =   1470
      Width           =   15
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame3"
      Height          =   1170
      Left            =   1870
      TabIndex        =   32
      Top             =   1470
      Width           =   15
   End
   Begin VB.TextBox Text1 
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
      Left            =   1365
      TabIndex        =   25
      Top             =   315
      Width           =   960
   End
   Begin VB.CommandButton Command5 
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
      Left            =   2310
      TabIndex        =   24
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFD0&
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Left            =   210
      TabIndex        =   21
      Top             =   1155
      Width           =   8100
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   2730
         TabIndex        =   31
         Top             =   315
         Width           =   15
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   210
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   29
         Top             =   580
         Width           =   7680
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "IdenLote        Cantid.  Código           Descripción"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   30
         Top             =   315
         Width           =   7680
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   5000
      Left            =   8505
      ScaleHeight     =   4935
      ScaleWidth      =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command14 
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
         Height          =   650
         Left            =   105
         Picture         =   "TRAZAUTO.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   3360
         Width           =   650
      End
      Begin VB.CommandButton Command8 
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
         Height          =   690
         Left            =   105
         Picture         =   "TRAZAUTO.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   2625
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Height          =   690
         Left            =   105
         Picture         =   "TRAZAUTO.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   945
         Width           =   650
      End
      Begin VB.CommandButton Command16 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "TRAZAUTO.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Listado de Informe de Trazabilidad"
         Top             =   1890
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   2
         Top             =   4200
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   12000
         End
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
         Height          =   690
         Left            =   105
         Picture         =   "TRAZAUTO.frx":106A
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Salir"
         Top             =   200
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "TRAZAUTO.frx":11B4
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4560
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFD0&
      Caption         =   "Trazabilidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   210
      TabIndex        =   5
      Top             =   2900
      Width           =   8100
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFD0&
         Caption         =   "Fechas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2080
         Left            =   6195
         TabIndex        =   41
         Top             =   315
         Width           =   1695
         Begin VB.TextBox Text12 
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
            Left            =   315
            TabIndex        =   13
            Top             =   1575
            Width           =   1065
         End
         Begin VB.CommandButton Command6 
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
            TabIndex        =   14
            Top             =   1575
            Width           =   175
         End
         Begin VB.TextBox Text8 
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
            Left            =   315
            TabIndex        =   11
            Top             =   1050
            Width           =   1065
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
            Left            =   1365
            TabIndex        =   12
            Top             =   1050
            Width           =   175
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Vencimiento:"
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
            Left            =   105
            TabIndex        =   45
            Top             =   1365
            Width           =   1335
         End
         Begin VB.Label Label5 
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
            Height          =   300
            Left            =   105
            TabIndex        =   44
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Ingreso:"
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
            Left            =   105
            TabIndex        =   43
            Top             =   315
            Width           =   1335
         End
         Begin VB.Label Label17 
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
            Left            =   315
            TabIndex        =   42
            Top             =   525
            Width           =   1020
         End
      End
      Begin VB.TextBox Text17 
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
         IMEMode         =   3  'DISABLE
         Left            =   945
         TabIndex        =   10
         Top             =   2100
         Width           =   5055
      End
      Begin VB.TextBox Text14 
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
         Left            =   3465
         TabIndex        =   9
         Top             =   1470
         Width           =   2535
      End
      Begin VB.TextBox Text5 
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
         Left            =   945
         TabIndex        =   8
         Top             =   1470
         Width           =   2220
      End
      Begin VB.Label Label29 
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
         Height          =   300
         Left            =   3675
         TabIndex        =   40
         Top             =   525
         Width           =   1335
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
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
         Left            =   5040
         TabIndex        =   37
         ToolTipText     =   "Cantidad nominal"
         Top             =   420
         Width           =   1020
      End
      Begin VB.Label Label6 
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
         Left            =   945
         TabIndex        =   36
         ToolTipText     =   "Descripción"
         Top             =   735
         Width           =   5115
      End
      Begin VB.Label Label27 
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
         Left            =   945
         TabIndex        =   35
         ToolTipText     =   "Código"
         Top             =   420
         Width           =   2175
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Left            =   -420
         TabIndex        =   34
         Top             =   525
         Width           =   1335
      End
      Begin VB.Label Label23 
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
         Left            =   6195
         TabIndex        =   22
         Top             =   2625
         Width           =   1650
      End
      Begin VB.Label Label20 
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
         Left            =   945
         TabIndex        =   20
         Top             =   2625
         Width           =   4380
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicacion Física:"
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
         Left            =   0
         TabIndex        =   17
         Top             =   1890
         Width           =   1755
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N.C.R.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4200
         TabIndex        =   16
         Top             =   2685
         Width           =   1965
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote de Fabricación:"
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
         Left            =   2625
         TabIndex        =   15
         Top             =   1215
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
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
         Left            =   -420
         TabIndex        =   7
         Top             =   2685
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Marca:"
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
         Left            =   -105
         TabIndex        =   6
         Top             =   1215
         Width           =   1335
      End
   End
   Begin VB.Label Label31 
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
      Height          =   300
      Left            =   105
      TabIndex        =   47
      Top             =   800
      Width           =   1125
   End
   Begin VB.Label Label30 
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
      Left            =   1365
      TabIndex        =   46
      Top             =   735
      Width           =   4590
   End
   Begin VB.Label Label28 
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
      Left            =   6405
      TabIndex        =   39
      Top             =   315
      Width           =   1860
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Remito Nro:"
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
      Left            =   4725
      TabIndex        =   38
      Top             =   380
      Width           =   1545
   End
   Begin VB.Label Label14 
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
      Height          =   300
      Left            =   5775
      TabIndex        =   28
      Top             =   800
      Width           =   1335
   End
   Begin VB.Label Label24 
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
      Left            =   7245
      TabIndex        =   27
      Top             =   735
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "B.R. Nº:"
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
      Left            =   105
      TabIndex        =   26
      Top             =   380
      Width           =   1125
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
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
      Height          =   300
      Left            =   1470
      TabIndex        =   19
      Top             =   4830
      Width           =   1335
   End
   Begin VB.Label Label18 
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
      Height          =   300
      Left            =   1470
      TabIndex        =   18
      Top             =   4830
      Width           =   1335
   End
End
Attribute VB_Name = "TRAZAUTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BRTEMP As Recordset

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Hide
   '
End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
    Call TRAZORF04.Command14_Click
    Command14.Enabled = True
End Sub

Private Sub Command16_Click()
    Command16.Enabled = False
    ' LISTADO DE CONSUMO POR LOTE DE RECEPCIÓN
    Call TRAZABI04.LICONLOT
    Command16.Enabled = True
End Sub

Private Sub Command2_Click()

   Command2.Enabled = False
   '
   NUBORE& = XVALO(Text1)
   If NUBORE& < 1 Then
      Call MENSERR(24, "Número de Boleta de Recepción no Válido.")
      GoTo 99
   End If
   '
   If ECOARCH$("INDIBOL", MKS$(NUBORE&)) = "" Then
      Call MENSERR(24, "Número de Boleta de Recepción no Válido.")
      GoTo 99
   End If
   '
   CIXI% = CODINT%(Label27)
   If CIXI% < 1 Then
      Call MENSERR(24, "Codigo no Válido")
      GoTo 99
   End If
   '
30 'If FECHANUM(Text8.TEXT) < 1 Then
   '   Call MENSERR(24, "Fecha de Fabricación no Válida.")
   '   GoTo 99
   'End If
   '
'   If TRIM$(Text9.TEXT) = "" Then
'    Call MENSERR(24, "Trazabilidad de Cliente no Válida.")
'    GoTo 99
'   End If
   '
'   If TRIM$(Text14.TEXT) = "" Then
'    Call MENSERR(24, "Código de Cliente no Válido.")
'    GoTo 99
'   End If
   '
   'If Text12.TEXT = "" Or FECHANUM(Text12) < 1 Then
   '   Call MENSERR(24, "Fecha de Vencimiento no Válida.")
   '   GoTo 99
   'End If
   '
   If FECHANUM(Text8) > FECHANUM(Label17) Then
      Call MENSERR(24, "La Fecha de Fabricación no Puede Ser\Posterior a la Fecha de Ingreso.")
      GoTo 99
   End If
   '
   If FECHANUM(Text12) < FECHANUM(Text8) Then
      Call MENSERR(24, "La Fecha de Vencimiento no Puede Ser\Anterior a la Fecha de Fabricación.")
      GoTo 99
   End If
   '
   If FECHANUM(Text12) < HOY(HOS$) Then
      Call MENSERR(24, "Este Material ya Está Vencido\al Día de Hoy.")
   End If
   '
   If TRIM$(Text17.TEXT) = "" Then
    Call MENSERR(24, "Ubicacion Física no Válida.")
    GoTo 99
   End If
   '
   IDREGIS$ = TRIM$(Label23)
   BRTEMP.FindFirst "IdenLote = '" & IDREGIS$ & "'"
   If BRTEMP.NoMatch = False Then
      BRTEMP.Edit
      BRTEMP!MARCA = Left$(Text5, 32)
      BRTEMP!NUMEPARTE = Left$(Text3, 32)
      BRTEMP!FECHFAB = CVDAT(FECHANUM(Left$(Text8, 8)))
      BRTEMP!TRAZ_CLI = Left$(Text9, 32)
      BRTEMP!COD_CLI = Left$(Text14, 32)
      BRTEMP!TIPO = Left$(Text11, 32)
      BRTEMP!CARAC = Left$(Text16, 32)
      BRTEMP!PLAZO_USO = CVDAT(FECHANUM(Left$(Text12, 8)))
      BRTEMP!UBICFIS = Left$(Text17, 32)
      BRTEMP.Update
      '
      FF% = FECHANUM(Left$(Label17, 8))
      CANOM# = XVALO(Label21)
      Call CREALOTIN(FF%, CIXI%, IDREGIS$, CANOM#)
      '
    Else
      Call MENSERR(24, "Registro no encontrado.")
      '  MENSAje de que no encontro el registro
   End If
   '
99 Screen.MousePointer = 1
   Command2.Enabled = True
   List1.SetFocus
   '
End Sub

Private Sub Command3_Click()
   'FORHLPDF.Pdf1.LOADFILE "C:\Mis Documentos\contab04.pdf"
   Command3.Enabled = False
   CUDI$ = UCase$(CurDir)
   PPXX% = InStr(CUDI$, "\WINDOWS")
   If PPXX% > 0 Then
      rutax$ = Left$(CUDI$, PPXX%) + "MANUALES\"
      If EXISTE%(rutax$ + "TRAZAB04.PDF") > 0 Then
         FORHLPDF.Pdf1.setShowToolbar (True)
         On Error Resume Next
         FORHLPDF.Pdf1.LOADFILE rutax$ + "TRAZAB04.pdf"
         If Err Then GoTo 20
         On Error GoTo 0
         FORHLPDF.Show 1
         GoTo 99
      End If
   End If
   '
20 Call MENSERR(24, "Imposible Abrir Archivo de Ayuda")
   '
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()

   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text8.TEXT = VDEV$
   End If
   
End Sub

Private Sub Command5_Click()

   Command5.Enabled = False
3  Call SELECHO("INDIBOL")
   NPX& = Val(VALACT1$("INDIBOL"))
   If NPX& > 0 Then
      Text1.TEXT = NPX&
   End If
99 Command5.Enabled = True

End Sub

Private Sub Command6_Click()

   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text12.TEXT = VDEV$
   End If
   
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call TRAZORF04.Command10_Click
   Command8.Enabled = True
End Sub
'
Function CARPRODOF$(NROOF$)
  '
  CARPRODOF$ = ""
  Call ABREORFAB
  '
  SQLX$ = " SELECT * FROM OPERFAB "
  SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
  Set OPFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
  '
  With OPFABTMP
    '
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      Exit Function
    End If
    On Error GoTo 0
    '
    CI% = !Cod_Inte
    CARPRODOF$ = CODEXT$(CI%)
  End With
  '
End Function

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If Left$(APLINVO$, 7) = "BORECEP" Then ' Or INTERPRE% = 1 Then
       TRAZAUTO.Text1 = Mid$(APLINVO$, 8)
       OPENFORMS = DoEvents
    End If
    '
End Sub

Sub VERIPERPRUE()
    '
    If CONTROL$("TRAZABI", "FINICIO") = "" Then
       HOII% = HOY(HOS$)
       Call GRACONTROL("TRAZABI", "FINICIO", HOS$)
    End If
    '
    FINICIO = FECHANUM(CONTROL$("TRAZABI", "FINICIO"))
    FECHACT = CVINT(FileDateTime(LUDAT$ + "MONIEXEC.DAT"))
    '
    Dim BB As String * 128
    NX1% = FILEOPEN(LUDAT$ + "MONIEXEC.DAT", 0, 128)
    FINI& = LOF(NX1%) / 128
    COMI& = FINI& - 128: If COMI& < 1 Then COMI& = 1
    '
    For U& = COMI& To FINI&
       Get #NX1%, U&, BB$
       FEJECU = FECHANUM(Mid$(BB$, 69, 8))
       If FEJECU > FECHACT Then
          FECHACT = FEJECU
       End If
    Next U&
    '
    If DIENTRE%(FINICIO, FECHACT) > 7 Then
       Call MENSERR(24, "Plazo de Verificación Vencido.")
       Call FINAL("")
    End If
    '
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Label24_Click()

    Label17.Caption = Label24.Caption
    
End Sub

Private Sub List1_Click()
    
    TTXX$ = List1.TEXT
    NCR$ = Mid(TTXX$, 1, 12)
    COD$ = Mid(TTXX$, 26, 16)
    CI% = CODINT%(COD$)
    CantNom$ = Mid(TTXX$, 14, 10)
    DESC$ = Mid(TTXX$, 43, 30)
    Label23.Caption = NCR$
    Label27.Caption = COD$
    Label6.Caption = DESC$
    Label21.Caption = CantNom$
    Call LIMPTEXT
    '
    IDREGIS$ = TRIM$(NCR$)
    BRTEMP.FindFirst "IdenLote = '" & IDREGIS$ & "'"
    If BRTEMP.NoMatch = False Then
        On Error Resume Next
        Text5 = BRTEMP!MARCA
        Text3 = BRTEMP!NUMEPARTE
        Text8 = TRIM$(FECHATEX$(CVINT(BRTEMP!FECHFAB)))
        Text9 = BRTEMP!TRAZ_CLI
        Text14 = BRTEMP!COD_CLI
        Text12 = TRIM$(FECHATEX$(CVINT(BRTEMP!PLAZO_USO)))
        Text17 = BRTEMP!UBICFIS
          If TRIM$(Text17) = "" Then
             UBICA$ = UBISTO$(CI%) + ","
             PPXX% = InStr(UBICA$, ",")
             Text17 = TRIM$(Left$(UBICA$, PPXX% - 1))
          End If
        Label20.Caption = BRTEMP!USUAPRO
        Text6.TEXT = BRTEMP!RASO_PROV
        List1.SetFocus
        On Error GoTo 0
      Else
        Label20.Caption = BRTEMP!USUAPRO
        Text6.TEXT = BRTEMP!RASO_PROV
        Text5.SetFocus
    End If

    Call FILTERFILE("TAMARCAS", "!TAMARCAS", 4, MKI$(CI%))
    Call FRESHECHO("!TAMARCAS")
    Call FILTERFILE("TANUMPAR", "!TANUMPAR", 4, MKI$(CI%))
    Call FRESHECHO("!TANUMPAR")
    '
End Sub

Private Sub List1_DblClick()
    '
    If DERACCE%("FRALOTIN", "Fraccionamiento de Lotes de Ingreso") < 2 Then Exit Sub
    '
    '\\\OT-07-0066-RG-21/02/07///
    TTXX$ = List1.TEXT
    IDLOTX$ = TRIM$(Mid(TTXX$, 1, 12))
    SQLX$ = "SELECT * FROM LOTINGRE "
    SQLX$ = SQLX$ + "WHERE IdenLote = '" & IDLOTX$ & "'"
    Set LOTINSEL = Stocks.OpenRecordset(SQLX$, dbOpenDynaset) ' para poder ajustar cantidad saldo
    With LOTINSEL
       On Error Resume Next
       .MoveFirst
       If Err Then
          On Error GoTo 0
          Call MENSERR(24, "Lote '" + IDLOTX$ + "' no Encontrado.\Probable Falta Trazar Ingreso.")
          Exit Sub
       End If
       On Error GoTo 0
       '
       SUINGRE# = !CanAlta
       SUCONSU# = 0
       SURESER# = 0
       If IsNull(!CanConsu) = True Or IsEmpty(!CanConsu) = True Then GoTo 22
       SUCONSU# = !CanConsu
22     If IsNull(!CanReser) = True Or IsEmpty(!CanReser) = True Then GoTo 23
       SURESER# = !CanReser
23     If SUCONSU# >= 0.05 Or SURESER# >= 0.05 Then
          Call MENSERR(24, "Imposible Fraccionar Lote\con Reservas / Consumos.")
          Exit Sub
       End If
    End With
    '
    TTXX$ = List1.TEXT
    NCR$ = TRIM$(Mid(TTXX$, 1, 12))
    PPXX% = InStr(NCR$, "/")
    If PPXX% > 0 Then
       NCR$ = Left$(NCR$, PPXX% - 1)
    End If
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = "SELECT * FROM BOLRECEP"
    SQLX$ = SQLX$ + " WHERE NUMEBORE = " & Int(XVALO(Text1.TEXT))
    Set BRTEMP = Stocks.OpenRecordset(SQLX, dbOpenDynaset)
    '
    For KKII% = 65 To 90
       NCR1$ = NCR$ + "/" + Chr$(KKII%)
       BRTEMP.FindFirst "IdenLote = '" & NCR1$ & "'"
       If BRTEMP.NoMatch = True Then GoTo 20
    Next KKII%
    Call MENSERR(24, "Imposible Fraccionar Este Lote.")
    Exit Sub
    '
20  FRACLOTIN04.Label9 = Text1
    FRACLOTIN04.Label27 = TRIM$(Label27.Caption)
    FRACLOTIN04.Label6 = TRIM$(Label6.Caption)
    FRACLOTIN04.Label4 = TRIM$(Label21.Caption)
    FRACLOTIN04.Label7 = NCR1$
    FRACLOTIN04.Label21 = TRIM$(Label21.Caption)
    FRACLOTIN04.Text1 = ""
    FRACLOTIN04.Label2 = TRIM$(Label23.Caption)
    FRACLOTIN04.Label24 = Label24
    Screen.MousePointer = 1
    FRACLOTIN04.Show 1
    '
    Call Text1_Change
    '
End Sub
'
Private Sub Text1_Change()
    Call LIMPLAB
    Call LIMPTEXT
    List1.Clear
    Set BRTEMP = Nothing
    If Text1.TEXT <> "" Then
      List1.Clear
      Label30 = ""
      Label28 = ""
      Label24 = ""
      Label17 = ""
      '
      Call APERBASDAT("STOCKS")
      SQLX$ = "SELECT * FROM BOLRECEP"
      SQLX$ = SQLX$ + " WHERE NUMEBORE = " & Int(XVALO(Text1.TEXT))
      Set BRTEMP = Stocks.OpenRecordset(SQLX, dbOpenDynaset)
      With BRTEMP
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
            On Error GoTo 0
            NPX% = !NumProv
            Label30 = RASOCLI$((-1) * NPX%)
            Label28 = !NumRem_Prov
            Label24 = FECHATEX$(CVINT(!FechRecep))
            Label17 = FECHATEX$(CVINT(!FechRecep))
            '
            Do While Not .EOF
               If CVINT(!FeHoApro) > 0 Then
                 TTXX$ = Space$(128)
                 Call REPLA(TTXX$, TRIM$(!IdenLote), 1, 12)
                 Call REPLA(TTXX$, FORMATNUM$(!CantNom, "F10.1"), 14, 10)
                 Call REPLA(TTXX$, TRIM$(!COD_EXTE), 26, 16)
                 Call REPLA(TTXX$, TRIM$(!DESCRIPITEM), 43, 48)
                 List1.AddItem TTXX$
               End If
            .MoveNext
            Loop
         End If
      End With
      '
      If List1.ListCount < 1 Then
         Call MENSERR(24, "La BR Elegida no Registra Ingresos\Aprobados que Puedan Procesarse.")
         Text1 = ""
         Label30 = ""
         Label28 = ""
         Label24 = ""
         Label17 = ""
      End If
    End If
    Screen.MousePointer = 1
    
End Sub

Private Sub Text1_DblClick()

    Command5.Enabled = False
3   Call SELECHO("INDIBOL")
    NPX& = Val(VALACT1$("INDIBOL"))
    If NPX& > 0 Then
       Text1.TEXT = NPX&
    End If
99  Command5.Enabled = True

End Sub

Sub LIMPLAB()

    Label27.Caption = ""
    Label6.Caption = ""
    Label21.Caption = ""
    Label23.Caption = ""
    Label20.Caption = ""
    Label28.Caption = ""
    Label30.Caption = ""
    Label24.Caption = ""
    Label17.Caption = ""
  
End Sub

Sub LIMPTEXT()

    Text5 = ""
    Text3 = ""
    Text6 = ""
    Text8 = ""
    Text9 = ""
    Text14 = ""
    Text11 = ""
    Text16 = ""
    Text12 = ""
    Text17 = ""
  
End Sub

Private Sub Text12_DblClick()

    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
       Text12.TEXT = VDEV$
    End If
    Text17.SetFocus
    
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Command2.SetFocus
   End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       Text17.SetFocus
    End If
    
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       Text8.SetFocus
    End If

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text14.SetFocus
   End If
End Sub

Private Sub Text8_DblClick()

    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
       Text8.TEXT = VDEV$
    End If

End Sub

Sub LICONLOT()
    ' LISTADO DE CONSUMO POR LOTE DE RECEPCIÓN
    '
    Dim VALOU(128)
    '
10  Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    '
    DESDE = CDbl(CVDAT(DES%))
    HASTA = CDbl(CVDAT(HAS%))
    REBLA$ = REGBLAN$("LICONLOT")
    JJ& = 0
    '
    OPXX% = COMALTER%("Formato de salida:\\Detalle\Resumen")
    If OPXX% < 1 Or OPXX% > 2 Then Exit Sub
    '
    Call VERILOTCO   ' ESTA RUTINA TOMA TODOS LOS MOVIMIENTOS DE MOVISTO QUE TIENEN IdenLote = ''
                     ' Y LOTCONSUM CON INFORMACIÓN Y GENERA LOS MOVIMIENTOS COMPLEMENTARIOS
                     ' CORRESPONDIENTES
    Call APERBASDAT("STOCKS")
    Screen.MousePointer = 1
20  SQLX$ = "SELECT * FROM LOTINGRE "
    SQLX$ = SQLX$ + "WHERE FechAlta >= " & DESDE & " "
    SQLX$ = SQLX$ + "AND FechAlta <= " & HASTA & " "
    SQLX$ = SQLX$ + "ORDER BY FechAlta ASC, IdenLote ASC "
    '
    Set LOTINSEL = Stocks.OpenRecordset(SQLX$, dbOpenDynaset) ' para poder ajustar cantidad saldo
    With LOTINSEL
       On Error Resume Next
       .MoveLast
       If Err Then
          On Error GoTo 0
          Call MENSERR(24, "No Hay Lotes Ingresados\en Rango de Fechas Especificado.")
          GoTo 10
       End If
       On Error GoTo 0
       FIN& = .RecordCount
       '
       .MoveFirst
       KU& = 0: IDLOA$ = ""
       Do While Not .EOF
          KU& = KU& + 1
          Call TRACE(20, KU&, FIN&)
          FALTA% = CVINT(!FechAlta)
          TIPO$ = UCase$(Left$(!IdenLote, 2))
          IDLO$ = TRIM$(!IdenLote)
          TRACLI$ = ""                ' FALTA TRAZABILIDAD CLIENTE
          PROVE$ = ""
          CIXI% = !Cod_Inte
          CANTIN# = 0: CSALDO# = 0: CACONRE# = 0: CARESRE# = 0
          '
          If IDLO$ <> IDLOA$ Then
               SUBLOT% = 0
             Else
               SUBLOT% = SUBLOT% + 1
          End If
          SULOTX$ = TRIM$(FORMATNUM$(SUBLOT%, "I3"))
          If SULOTX$ <> "" Then SULOTX$ = "/" + SULOTX$
          IDLOA$ = IDLO$
          '
          Call ACSALOT(CIXI%, IDLO$, CANTIN#, CACONRE#, CARESRE#, CSALDO#, NPROVX%)
             ' ESTA RUTINA HACE LO SIGUIENTE:
             ' - VERIFICA LAS CANTIDADES DEL MOVISTO CONTRA EL LOTE

          If CAMOVILO% < 1 Or CANTIN# < 0.05 Then
25           TTXX$ = "No se Ha Encontrado Movimiento de Ingreso\Correspondiente al Lote '" + IDLO$ + "'"
             OPX% = COMALTER%(TTXX$ + "\\Eliminar Lote\Continuar")
             If OPX% = 1 Then
                 LOTINSEL.Delete
                 GoTo 20
               ElseIf OPX% = 2 Then
                 GoTo 30
               Else
                 GoTo 25
            End If
          End If
          '
30        PROVE$ = FANTCLI$((-1) * NPROVX%)
          '
          XX0$ = REBLA$
          JJ0& = JJ&
          SUCONSU# = 0: SURESER# = 0
          Call REPLA(XX0$, MKI$(FALTA%), 1, 2)
          Call REPLA(XX0$, TIPO$, 3, 2)
          Call REPLA(XX0$, IDLO$ + SULOTX$, 5, 16)
          Call REPLA(XX0$, TRACLI$, 21, 32)
          Call REPLA(XX0$, PROVE$, 53, 16)
          Call REPLA(XX0$, MKI$(CIXI%), 69, 2)
          Call REPLA(XX0$, MKD$(CANTIN#), 71, 8)
          Call REPLA(XX0$, MKD$(CSALDO#), 79, 8)
          '
          For KKI% = 1 To CAMOVILO%
             CCONSU# = CANSALLO#(KKI%)
             If Abs(CCONSU#) >= 0.05 Then
                XX$ = XX0$
                Call REPLA(XX$, MKI$(FEMOVILO%(KKI%)), 95, 2)
                Call REPLA(XX$, DEMOVILO$(KKI%), 97, 24)
                Call REPLA(XX$, MKD$(CCONSU#), 153, 8)
                SUCONSU# = SUCONSU# + CCONSU#
                JJ& = JJ& + 1
                Call GRAREG("LICONLOT", XX$, JJ&)
             End If
          Next KKI%
          '
          For KKI% = 1 To CAMOVILO%
             CCONSU# = ROUND(CANRESLO#(KKI%), 1)
             If Abs(CCONSU#) >= 0.05 Then
                XX$ = XX0$
                Call REPLA(XX$, MKI$(FEMOVILO%(KKI%)), 95, 2)
                Call REPLA(XX$, DEMOVILO$(KKI%), 97, 24)
                    '    ' aca faltan los numeros PNL/PLC
                Call REPLA(XX$, MKD$(CCONSU#), 169, 8)
                SURESER# = SURESER# + CCONSU#
                JJ& = JJ& + 1
                Call GRAREG("LICONLOT", XX$, JJ&)
             End If
          Next KKI%
          '
          If JJ& <= JJ0& Then
             XX$ = XX0$
             JJ& = JJ& + 1
             Call GRAREG("LICONLOT", XX$, JJ&)
          End If
          '
       .MoveNext
       Loop
    End With
    '
    Call SETULTREG("LICONLOT", JJ&)
    Screen.MousePointer = 1
    '
    If OPXX% = 1 Then
         Call HEADERS("LICONLOT", "")
         Call HEADERS("LICONLOT", "Ingresos Correspondientes a: " + PERIO$)
         Call CONECRUN("*LICONLOT", "Listado Detalle de Consumos por Lote.")
      Else
         JJ1& = 0
         For KKU& = 1 To JJ&
            XX$ = REGLEIDO$("LICONLOT", KKU&)
            If Left$(XX$, 70) = Left$(ZZ$, 70) Then
                SUCONSU# = CVD(Mid$(XX$, 153, 8)) + CVD(Mid$(ZZ$, 153, 8))
                SURESER# = CVD(Mid$(XX$, 169, 8)) + CVD(Mid$(ZZ$, 169, 8))
                Call REPLA(ZZ$, MKD$(SUCONSU#), 153, 8)
                Call REPLA(ZZ$, MKD$(SURESER#), 169, 8)
              Else
                ZZ$ = XX$
                JJ1& = JJ1& + 1
            End If
            Call GRAREG("RECONLOT", ZZ$, JJ1&)
         Next KKU&
         Call SETULTREG("RECONLOT", JJ1&)
         '
         'ESTO ES PARA TRASLADAR EL CÓDIGO DE ARTICULO A LA PRIMERA COLUMNA
         For U& = 1 To JJ1&
           XX$ = REGLEIDO$("RECONLOT", U&)
           XX1$ = Mid$(XX$, 69, 2) + Left$(XX$, 2)
           Call REPLA(XX$, XX1$, 1, 4)
           Call GRAREG("RECONLOT", XX$, U&)
         Next U&
         Call CIERRARCH("RECONLOT")
         '
         Call FILESORT("RECONLOT", "", 1, 3)
         Call HEADERS("RECONLOT", "")
         Call HEADERS("RECONLOT", "Ingresos Correspondientes a: " + PERIO$)
         Call CONECRUN("*RECONLOT", "Listado Resumen de Consumos por Lote.")
    End If
    '
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text12.SetFocus
   End If
End Sub
