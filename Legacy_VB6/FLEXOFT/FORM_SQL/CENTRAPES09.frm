VERSION 5.00
Begin VB.Form CENTRAPES09 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Central de Fraccionamiento - Consumos"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5940
      TabIndex        =   74
      Top             =   5880
      Width           =   175
   End
   Begin VB.CommandButton Command8 
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
      Left            =   3195
      TabIndex        =   73
      Top             =   5880
      Width           =   175
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   11025
      ScaleHeight     =   2355
      ScaleWidth      =   1320
      TabIndex        =   57
      Top             =   5460
      Width           =   1380
      Begin VB.CommandButton Command15 
         Height          =   600
         Left            =   105
         Picture         =   "CENTRAPES09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Ficha Kardex de Movimientos"
         Top             =   105
         Width           =   645
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   59
         Top             =   1245
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   -135
            TabIndex        =   60
            Top             =   -60
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Save"
         Height          =   780
         Left            =   105
         Picture         =   "CENTRAPES09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Registra Movimientos de Stock"
         Top             =   1485
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "CENTRAPES09.frx":0614
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   795
         Width           =   1515
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Fraccionamiento"
      Height          =   1860
      Left            =   6300
      TabIndex        =   48
      Top             =   5880
      Width           =   4575
      Begin VB.CommandButton Command3 
         Caption         =   "O.K."
         Height          =   285
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Genera Etiquetas Material Consumido y Remanente"
         Top             =   1470
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   840
         TabIndex        =   52
         Top             =   1470
         Width           =   1400
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   840
         TabIndex        =   56
         ToolTipText     =   "Doble Click para Consumir Saldo Lote"
         Top             =   345
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   840
         TabIndex        =   54
         ToolTipText     =   "Doble Click para Consumir Saldo Asignacion"
         Top             =   705
         Width           =   1395
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   2380
         X2              =   4515
         Y1              =   1300
         Y2              =   1300
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MAX:"
         Height          =   270
         Left            =   2145
         TabIndex        =   68
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "min:"
         Height          =   270
         Left            =   1470
         TabIndex        =   67
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2955
         TabIndex        =   66
         ToolTipText     =   "Código"
         Top             =   915
         Width           =   1395
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2955
         TabIndex        =   65
         ToolTipText     =   "Código"
         Top             =   570
         Width           =   1395
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Prec:"
         Height          =   270
         Left            =   2415
         TabIndex        =   64
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2955
         TabIndex        =   63
         ToolTipText     =   "Código"
         Top             =   210
         Width           =   1395
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   2380
         X2              =   2380
         Y1              =   105
         Y2              =   1300
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
         Height          =   270
         Left            =   -645
         TabIndex        =   55
         Top             =   390
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   840
         TabIndex        =   53
         Top             =   1050
         Width           =   1395
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fracc:"
         Height          =   270
         Left            =   -645
         TabIndex        =   51
         Top             =   1510
         Width           =   1455
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Asigna:"
         Height          =   270
         Left            =   -645
         TabIndex        =   50
         Top             =   735
         Width           =   1455
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Consum:"
         Height          =   270
         Left            =   -645
         TabIndex        =   49
         Top             =   1095
         Width           =   1455
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   945
      TabIndex        =   44
      Top             =   5460
      Width           =   5170
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
      Left            =   10605
      TabIndex        =   35
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ubicacion del Remanente"
      Height          =   1020
      Left            =   945
      TabIndex        =   27
      Top             =   6720
      Width           =   5205
      Begin VB.CommandButton Command2 
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
         Left            =   4830
         TabIndex        =   34
         Top             =   270
         Width           =   175
      End
      Begin VB.CommandButton Command1 
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
         Left            =   1530
         TabIndex        =   33
         Top             =   630
         Width           =   175
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
         Height          =   270
         Left            =   -500
         TabIndex        =   70
         Top             =   320
         Width           =   1455
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1035
         TabIndex        =   69
         ToolTipText     =   "Código"
         Top             =   270
         Width           =   1395
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ubic:"
         Height          =   270
         Left            =   2205
         TabIndex        =   32
         Top             =   320
         Width           =   1035
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Deposito:"
         Height          =   270
         Left            =   105
         TabIndex        =   31
         Top             =   675
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1035
         TabIndex        =   30
         Top             =   630
         Width           =   510
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3360
         TabIndex        =   29
         Top             =   270
         Width           =   1500
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1890
         TabIndex        =   28
         Top             =   630
         Width           =   3090
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   11025
      ScaleHeight     =   1515
      ScaleWidth      =   1320
      TabIndex        =   25
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command7 
         Height          =   600
         Left            =   120
         Picture         =   "CENTRAPES09.frx":2536
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Fraccionamiento Multiple"
         Top             =   765
         Width           =   645
      End
      Begin VB.CommandButton Command4 
         Height          =   600
         Left            =   135
         Picture         =   "CENTRAPES09.frx":2840
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Salir"
         Top             =   90
         Width           =   650
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Material Asignado a Fraccionar"
      Height          =   3570
      Left            =   105
      TabIndex        =   14
      Top             =   1680
      Width           =   11685
      Begin VB.Frame Frame7 
         Caption         =   "Frame3"
         Height          =   3040
         Left            =   11240
         TabIndex        =   40
         Top             =   375
         Width           =   30
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame3"
         Height          =   3040
         Left            =   7455
         TabIndex        =   39
         Top             =   375
         Width           =   30
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame3"
         Height          =   3040
         Left            =   9975
         TabIndex        =   38
         Top             =   375
         Width           =   30
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   3040
         Left            =   2100
         TabIndex        =   18
         Top             =   375
         Width           =   30
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame3"
         Height          =   3040
         Left            =   6930
         TabIndex        =   17
         Top             =   375
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame3"
         Height          =   3040
         Left            =   8715
         TabIndex        =   16
         Top             =   375
         Width           =   30
      End
      Begin VB.Frame Frame9 
         Caption         =   "Frame3"
         Height          =   3040
         Left            =   5040
         TabIndex        =   15
         Top             =   375
         Width           =   30
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
         Height          =   2760
         ItemData        =   "CENTRAPES09.frx":298A
         Left            =   225
         List            =   "CENTRAPES09.frx":298C
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   645
         Width           =   11325
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"CENTRAPES09.frx":298E
         Height          =   270
         Left            =   225
         TabIndex        =   20
         Top             =   375
         Width           =   11310
      End
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
      Left            =   3225
      TabIndex        =   1
      Top             =   315
      Width           =   175
   End
   Begin VB.CommandButton Command30 
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
      Left            =   10605
      TabIndex        =   0
      Top             =   735
      Width           =   175
   End
   Begin VB.Label Label16 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   10080
      TabIndex        =   61
      Top             =   735
      Width           =   510
   End
   Begin VB.Label Label33 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7785
      TabIndex        =   47
      ToolTipText     =   "Código"
      Top             =   5460
      Width           =   3105
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Id-Reque:"
      Height          =   270
      Left            =   6405
      TabIndex        =   46
      Top             =   5505
      Width           =   1305
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ">>>>>"
      Height          =   225
      Left            =   -3045
      TabIndex        =   45
      Top             =   5505
      Width           =   3900
   End
   Begin VB.Label Label21 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   945
      TabIndex        =   43
      Top             =   6300
      Width           =   5160
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   945
      TabIndex        =   42
      Top             =   5880
      Width           =   2280
   End
   Begin VB.Label Label19 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4170
      TabIndex        =   41
      Top             =   5880
      Width           =   1800
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cant.Proyectada:"
      Height          =   300
      Left            =   4200
      TabIndex        =   37
      Top             =   790
      Width           =   1545
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5850
      TabIndex        =   36
      Top             =   750
      Width           =   1395
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      Height          =   270
      Left            =   4620
      TabIndex        =   24
      Top             =   6405
      Width           =   900
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Id.Bulto:"
      Height          =   270
      Left            =   3330
      TabIndex        =   23
      Top             =   5925
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descrip:"
      Height          =   285
      Left            =   -300
      TabIndex        =   22
      Top             =   6345
      Width           =   1185
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      Height          =   270
      Left            =   210
      TabIndex        =   21
      Top             =   5925
      Width           =   675
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Movimiento:"
      Height          =   510
      Left            =   8400
      TabIndex        =   13
      Top             =   210
      Width           =   1155
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
      Height          =   300
      Left            =   -150
      TabIndex        =   12
      Top             =   790
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "O/Prod. Nº:"
      Height          =   300
      Left            =   180
      TabIndex        =   11
      Top             =   355
      Width           =   1125
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   9615
      TabIndex        =   10
      Top             =   315
      Width           =   1020
   End
   Begin VB.Label Label28 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   750
      Width           =   1965
   End
   Begin VB.Label Label30 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   1155
      Width           =   5790
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      Height          =   300
      Left            =   210
      TabIndex        =   7
      Top             =   1195
      Width           =   1125
   End
   Begin VB.Label Label32 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      ToolTipText     =   "Doble-Click para Blanquear"
      Top             =   315
      Width           =   1800
   End
   Begin VB.Label Label26 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5460
      TabIndex        =   5
      Top             =   315
      Width           =   1800
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro S/F:"
      Height          =   300
      Left            =   4200
      TabIndex        =   4
      Top             =   365
      Width           =   1125
   End
   Begin VB.Label Label38 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7455
      TabIndex        =   3
      Top             =   1155
      Width           =   3300
   End
   Begin VB.Label Label57 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Depósito Pesada:"
      Height          =   225
      Left            =   8295
      TabIndex        =   2
      Top             =   780
      Width           =   1695
   End
End
Attribute VB_Name = "CENTRAPES09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private IndexGuia As Integer

Private Sub Command1_Click()
   Call SELECHO("TADEPOSI")
   NUMDEP% = XVALO(VALACT1$("TADEPOSI"))
   Label9 = VALACT1$("TADEPOSI")
End Sub

Private Sub Command13_Click()
   '
   Command13.Enabled = False
   DEPOTRANS% = XVALO(CONTROL$("BOLRECEP", "DEPTRANS")) ' DEPOSITO MATERIAL EN TRANSITO
   If XVALO(Label16) = DEPOTRANS% Then
        COMUNI ("Depósito de Pesadas No Válido\Configurado Como Material en Transito ")
        Label16 = ""
        GoTo 99
   End If

   Screen.MousePointer = 11
   
   NDEP% = XVALO(Label16)
   FF% = FECHANUM(Label24)
   NORI% = 0
   For UI& = 0 To List1.ListCount
      LINEA$ = List1.List(UI&)
      CODIGO$ = TRIM$(Mid(LINEA$, 1, 16))
      CIXI% = CODINT(CODIGO$)
      IDLOT$ = TRIM$(Mid$(LINEA$, 47, 16))
      CANCON# = XVALO(Mid$(LINEA$, 93, 12))
      '
      If CANCON# >= FRACMIN(CIXI%) Then
        CMO$ = "CF"
        DOPRI$ = "OP-" + Label32
        DOSEC$ = Label26
        REFE$ = "Fracc.Central de Pesadas"
        IDREQ$ = TRIM$(Mid$(LINEA$, 151, 32))
        NORI% = NORI% + 1
        Call MOVISTOK(FF%, CIXI%, IDLOT$, CMO$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, 0, 0, 0, NDEP%, (-1) * CANCON#, , , , , IDREQ$)
        '
        CONDI$ = "Cod_Inte = " & CIXI%
        CONDI$ = CONDI$ + " AND IdenLote = '" & IDLOT$ & "'"
        CONDI$ = CONDI$ + " AND IdenMovi = '" & IDREQ$ & "'"
        CANYACON# = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", CONDI$, "NOMESS"))
        CANYACON# = CANYACON# - XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", CONDI$, "NOMESS"))
        If CANYACON# < 0 Then CANYACON# = 0
        '
        CONDI$ = "Cod_Inte = " & CIXI%
        CONDI$ = CONDI$ + " AND IdenLote = '" & IDLOT$ & "'"
        CONDI$ = CONDI$ + " AND IdenReque = '" & IDREQ$ & "'"
        CONDI$ = CONDI$ + " AND Kilo_Asig > 0"
        Call ACTUREGISTRO("BALAREQMO", "Kilo_Consu", CONDI$, CANYACON, RAFE&, "NOMESS")
      End If
      '
      REMANEN# = SaldoLote(CIXI%, IDLOT$, NDEP%) - CANCON
      If REMANEN# >= FRACMIN(CIXI%) / 2 Then
        DEPREMA% = XVALO(Label9)
        UBIREMA$ = TRIM$(Label7)
        If DEPREMA% > 0 And DEPREMA% <> NDEP% Then
           ' REALIZA EL MOVIMIENTO DE STOCK
           CMOV$ = TRIM$(CONTROL$("", "CODTRAST"))
           If CMOV$ = "" Then CMOV$ = "TR"
           CAMPOVA$ = "max(CAST(substring(DoPriCor, 4,12) AS INTEGER)"
           CONDI$ = "CodiMovi = '" & CMOV$ & "'"
           NRO& = 1 + XVALO(VALOBADA("MOVISTO", CAMPOVA$, CONDI$))
           DOPRI$ = "TR." + TRIM$(Str$(NRO&))
           Call MOVISTOK(FF%, CIXI%, IDLOT$, "TR", DOPRI$, DOPRI$, DOSEC$, DOSEC$, "Remanente C.Pesada", NORI%, 0, 0, 0, NDEP%, (-1) * REMANEN#)
           Call MOVISTOK(FF%, CIXI%, IDLOT$, "TR", DOPRI$, DOPRI$, DOSEC$, DOSEC$, "Transferencia de C.Pesada", NORI%, 0, 0, 0, DEPREMA%, REMANEN#)
        End If
        If UBIREMA$ <> TRIM$(UBILOTE$(CIXI%, IDLOT$, 1)) Then
           ' PONE LA NUEVA UBICACION
           CONDI$ = "COD_INTE = " & CIXI%
           CONDI$ = CONDI$ & " AND IDENLOTE = '" & IDLOT$ & "'"
           CONDI$ = CONDI$ & " AND DEPOUBIC = " & DEPREMA%
           Call ACTUREGISTRO("LOTINGRE", "UBICFIS", CONDI$, UBIREMA$, REGG&, "NOMESS")
        End If
      End If
   Next UI&
   '
   Label32 = ""
   List1.Clear
99 Screen.MousePointer = 1
   Command13.Enabled = True
   '
End Sub

Private Sub Command15_Click()
    FIKARDEX09.Text1 = Label20
    FIKARDEX09.Show 1
End Sub

Private Sub Command2_Click()
    DEPO% = XVALO(Label9)
    If ECOARCH$("TADEPOSI", Chr$(DEPO%)) = "" Then
        Call MENSERR(24, "Primero Debe Seleccionar el Deposito")
        Exit Sub
    End If
    Call GENTAUBI(DEPO%)
    Call SELECHO("!TAUBICA")
    
    UBISELE$ = VALACT1$("!TAUBICA")
    If UBISELE$ <> "" Then
       Label7 = UBISELE$
    End If
End Sub

Private Sub Command3_Click()
   '
   Command3.Enabled = False
   '
   CIXI% = CODINT%(Label20)
   If CIXI% < 1 Then GoTo 99
   '
   If XVALO(Text1) < 0 Then
      COMUNI ("La Cant. Fraccionada Debe ser Mayor a Cero")
      GoTo 99
   End If
   '
   If XVALO(Text1) > XVALO(Label4) Then
      Call COMUNI("La Cant. Fraccionada No Puede Ser Mayor a la Cant. en Saldo")
      GoTo 99
   End If
   '
   DEPENTRA% = XVALO(Label9)
   '
   DEPOTRANS% = XVALO(CONTROL$("BOLRECEP", "DEPTRANS")) ' DEPOSITO MATERIAL EN TRANSITO
   If DEPENTRA% = DEPOTRANS% Then
       Call COMUNI("Depósito de Remanente No Válido\Configurado como Material en Transito ")
       GoTo 99
   End If
   '
   If Label44 > FRACMIN(CIXI%) Then  ' HAY REMANENTE
      If DEPENTRA% < 1 Or ECOARCH$("TADEPOSI", Chr$(DEPENTRA%)) = "" Then
         Call MENSERR(24, "Deposito de Entrada de Remanente no Válido")
         GoTo 99
      End If
      '
   End If
   '
   If Label33 = "" Then
      Call MENSERR(24, "El Fraccionado no Corresponde a Requerimiento Alguno.")
      GoTo 99
   End If
   '
20 IDLO$ = TRIM$(Label19)
   IDREQ$ = TRIM$(Label33)
   For JJ& = 1 To List1.ListCount
     LINEA$ = List1.List(JJ& - 1)
     If TRIM$(Left$(LINEA$, 16)) = TRIM$(Label20) Then
       If TRIM$(Mid$(LINEA$, 47, 16)) = IDLO$ Then
         If TRIM$(Mid$(LINEA$, 151, 32)) = IDREQ$ Then
           IndexGuia = JJ& - 1
           GoTo 30
         End If
       End If
     End If
   Next JJ&
   '
   'AGREGAR UNA LINEA
   Call REPLA(LINEA$, Label20, 1, 16)
   Call REPLA(LINEA$, Label21, 19, 26)
   Call REPLA(LINEA$, IDLO$, 47, 16)
   Call REPLA(LINEA$, UNIMED$(CIXI%), 65, 4)
      CX$ = FORMATSTAN$(0#, "F12." & CNTDC$, CIXI%)
   Call REPLA(LINEA$, CX$, 69, 12)
      CX$ = FORMATSTAN$(0#, "F12." & CNTDC$, CIXI%)
   Call REPLA(LINEA$, CX$, 81, 12)
      CX$ = FORMATSTAN$(0, "F12." & CNTDC$, CIXI%)
   Call REPLA(LINEA$, CX$, 93, 12)
   Call REPLA(LINEA$, Label33, 151, 32)
   List1.AddItem LINEA$
   GoTo 20
   '
30 FRAC$ = FORMATSTAN(XVALO(Text1), "F12." & CNTDC$, CIXI%)
   DEPO$ = FORMATNUM$(XVALO(Label9), "I4")
   UBICA$ = Label7
   '
   Call REPLA(LINEA$, FRAC$, 93, 12)
   Call REPLA(LINEA$, DEPO$, 183, 4)
   Call REPLA(LINEA$, UBICA$, 187, 24)
   '
   List1.List(IndexGuia) = LINEA$
   '
   SALD# = XVALO(Label44)
   If SALD# > 0 Then
      CANTX# = XVALO(Text1)
      DEP% = XVALO(Label16)
      LOTE$ = Label19
      'Imprime Etiqueta por Parte Fraccionada
      Call PRINSTIKAPRO1(CIXI%, LOTE$, DEP%, CANTX#)
      'Imprime Etiqueta por Remanente en caso de Haberlo
      Call PRINSTIKAPRO1(CIXI%, LOTE$, DEPENTRA%, SALD#)
   End If
  
   
   
 
  
   '
99 Command3.Enabled = True
   '
End Sub

Private Sub Command30_Click()
   Call SELECHO("TADEPOSI")
   NUMDEP% = XVALO(VALACT1$("TADEPOSI"))
   Label16 = VALACT1$("TADEPOSI")
End Sub

Private Sub Command4_Click()
   '
   For UI& = 0 To List1.ListCount
      LINEA$ = List1.List(UI&)
      CODIGO$ = TRIM$(Mid(LINEA$, 1, 16))
      CIXI% = CODINT(CODIGO$)
      CANCON# = XVALO(Mid$(LINEA$, 93, 12))
      If CANCON# >= FRACMIN(CIXI%) Then
         Call MENSERR(24, "Imposible Cerrar. Debe Registrar Movimientos\con Click sobre Botón 'Save'")
         Exit Sub
      End If
   Next UI&
   '
10 Call CIERRARCH("*.*")
   Hide
   '
End Sub

Private Sub Command5_Click()
    '
    CONDI$ = "StatuOrf < 3 "   ' para que no este anulada ni cumplida
    CONDI$ = CONDI$ + "AND CanApro < (CanProy - 0.005) "
    CONDI$ = CONDI$ + " and NuOrProd <> '' "
    CONDI$ = CONDI$ + " ORDER BY NuOrProd "
    '
    Call CARGALISEL("ORDPEN", "ORDEFABR", "NUORPROD/A16;REFEREN/A48", CONDI$, "DISTINCT")
    Call CommSel("ORDPEN/Indice de Ordenes de Produccion Pendientes", Command5, Label32)
    
End Sub

Private Sub Command6_Click()
    Call CommSel("SELFECHA", Command6, Label24)
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   FRACMULT09.Show 1
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Text2 = ""
    CODAN$ = Label20
    Call CommSel("MASTER", Command8, Label20)
    If Label20 <> LOTAN$ Then
       Label19 = ""
    End If
End Sub

Private Sub Command9_Click()
    Text2 = ""
    CIXI% = CODINT(Label20)
    If CIXI% < 1 Then
       Call COMUNI("Codigo No Valido")
       GoTo 99
    End If
    '
    DEPP% = XVALO(Label16)
    If DEPP% < 1 Then
       Call COMUNI("Deposito No Valido")
       GoTo 99
    End If
    '
    Call ARMALILOT(CIXI%, DEPP%, , , 1, "APRO")
    If ULTREG&("!LILOTSE") < 1 Then
       Call MENSERR(24, "No Hay Lotes Aprobados Disponibles\para el Código '" + TRIM$(Label20) + "'\en el Deposito Numero '" & DEPP% & ".")
       GoTo 99
    End If
    '
    Call CommSel("!LILOTSE", Command9, Label19)
    '
99 End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    HOII% = HOY(HOS$)
    Label24 = HOS$
    Label16.Caption = CONTROL$("STOCKS", "DEPCENPE")
    '
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   '
   For UI& = 0 To List1.ListCount
      LINEA$ = List1.List(UI&)
      CODIGO$ = TRIM$(Mid(LINEA$, 1, 16))
      CIXI% = CODINT(CODIGO$)
      CANCON# = XVALO(Mid$(LINEA$, 93, 12))
      If CANCON# >= FRACMIN(CIXI%) Then
         Cancel = 1
         Call MENSERR(24, "Imposible Cerrar. Debe Registrar Movimientos\con Click sobre Botón 'Save'")
         Exit Sub
      End If
   Next UI&
   Cancel = 0
   '
End Sub

Private Sub Label16_DblClick()
    Call Command30_Click
End Sub

Private Sub Label19_Change()
   '
   IDLO$ = TRIM$(Label19)
   '
   IDRE% = 0
   If IDLO$ <> "" Then
     For JJ& = 1 To List1.ListCount
       LINEA$ = List1.List(JJ& - 1)
       If TRIM$(Left$(LINEA$, 16)) = TRIM$(Label20) Then
         Label33 = TRIM$(Mid$(LINEA$, 151, 32))
         IDRE% = 1
         If TRIM$(Mid$(List1.List(JJ& - 1), 47, 16)) = IDLO$ Then
           List1.ListIndex = JJ& - 1
           Call List1_Click
           Call Text1_Change
           Exit Sub
         End If
       End If
     Next JJ&
     Call MENSERR(24, "Lote no Asignado")
     If IDRE% = 0 Then Label33 = ""
     Call PROPUFRAC
   End If
   '
End Sub

Private Sub Label20_Change()
   CI% = CODINT(Label20)
   Label21 = DESCRIT0(CI%)
   '
   If CI% > 0 Then
      For JJ& = 1 To List1.ListCount
         If TRIM$(Left$(List1.List(JJ& - 1), 16)) = TRIM$(Label20) Then
            Exit Sub
         End If
      Next JJ&
      Call MENSERR(24, "Material no Requerido")
   End If
   '
End Sub


Private Sub Label3_Click()
    CIXI% = CODINT%(Label20)
    VALAFRAC = MINIVAL(XVALO(Label3), XVALO(Label4) - XVALO(Label5))
    VALAFRAC = MAXIVAL(VALAFRAC, 0)
    Text1 = FORMATNUM$(VALAFRAC, "F13." + CNTDC(CIXI%))
End Sub

Private Sub Label32_Change()
    '
    ' CASO 1 - PONE EL CAMPO EN BLANCO
    NORFA$ = Label32
       NVM$ = TRIM$(ProNroComprobante("CF"))
       While Len(NVM$) < 6: NVM$ = "0" + NVM$: Wend
    Label51 = "CF-" + NVM$
    HOII% = HOY(HOS$): Label24 = HOS$
    
    '
    ' PONER EN BLANCO TODOS LOS RESTANTES CAMPOS Y LIST'
    Text1 = "": Text2 = ""
    Label17 = "": Label28 = "": Label30 = "": Label26 = ""
    Label19 = "": Label20 = "": Label21 = ""
    Label44 = "": Label7 = "": Label9 = ""
    Label3 = "": Label4 = "": Label5 = "": Label33 = ""
    List1.Clear
    '
    If NORFA$ <> "" Then
       ' presenta datos de la orden de fabricación
       CONDI$ = "NuOrProd = '" & NORFA$ & "'"
       CIXI% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", CONDI$, "NOMESS"))
       Label28 = CODEXT$(CIXI%)
       Label30 = DESCRIT0$(CIXI%)
       Label26 = SAFETEXT(VALOBADA("ORDEFABR", "IdSubOr", CONDI$, "NOMESS"))
       Label17 = FORMATNUM(XVALO(VALOBADA("ORDEFABR", "CANPROY", CONDI$, "NOMESS")), "F13." + CNTDC$(CIXI%))
       Call CargarLista
    End If
    '
99  Screen.MousePointer = 1
    '
End Sub

Private Sub Label32_DblClick()
    Call Command5_Click
End Sub

Private Sub Label4_DblClick()
   Text1 = Label4
End Sub

Private Sub Label9_Change()
    Label6 = ECOARCH$("TADEPOSI", Chr$(XVALO(Label9)))
End Sub

Private Sub List1_Click()
   '
   NDEP% = XVALO(Label16)
   LINEA$ = List1.List(List1.ListIndex)
   CODIGO$ = TRIM$(Mid(LINEA$, 1, 16))
   CIXI% = CODINT%(CODIGO$)
   IDLOT$ = TRIM$(Mid$(LINEA$, 47, 16))
   IDREQ$ = TRIM$(Mid$(LINEA$, 151, 32))
   CANASIG$ = TRIM$(Mid$(LINEA$, 69, 12))
   CANCONSU$ = TRIM$(Mid$(LINEA$, 81, 12))
   CANAFRAC# = XVALO(Mid$(LINEA$, 93, 12))
   If CANAFRAC# < FRACMIN(CIXI%) Then
      CANAFRAC# = XVALO(CANASIG$) - XVALO(CANCONSU$)
   End If
   '
   'VALIDACION DE DEPOSITO SALDO
   If XVALO(Label4) < FRACMIN(CIXI%) Then
      Label16 = TRIM$(Str$(DEPOLOT%(CIXI%, IDLOT$)))
   End If
   '
   Label20 = CODIGO$
   Label21 = DESCRIT0(CIXI%)
   Label19 = IDLOT$
   Label33 = IDREQ$
   '
   Label3 = FORMATNUM(CANASIG$, "F13." + CNTDC$(CIXI%))
   Label5 = FORMATNUM(CANCONSU$, "F13." + CNTDC$(CIXI%))
   '
   NDEP% = XVALO(Label16)
   SALENDEP = SaldoLote(CIXI%, IDLOT$, NDEP%)
   '
   Call PROPUFRAC
   '
   If SALENDEP < XVALO(Label3) - XVALO(Label5) Then
      Call MENSERR(24, "Saldo Insuficiente del Lote\en el Depósito Indicado.")
   End If
   '
End Sub
'
Sub PROPUFRAC()
   '
   CIXI% = CODINT%(Label20)
   IDLOT$ = TRIM$(Label19)
   IDREQ$ = TRIM$(Label33)
   NDEP% = XVALO(Label16)
   '
   Label4 = FORMATNUM$(MAXIVAL(SaldoLote(CIXI%, IDLOT$, NDEP%), 0), "F13." + CNTDC$(CIXI%))
      If CANAFRAC# > XVALO(Label4) Then CANAFRAC# = XVALO(Label4)
   Text1 = FORMATNUM$(CANAFRAC#, "F13." + CNTDC$(CIXI%))
   '
   'Muestra max min prec
   COCONJ% = CODINT(Label28)
   COELE% = CODINT(Label20)
   CONDI$ = "CODICONJ = " & COCONJ% & " AND CODIELEM= " & COELE%
   MINIMO# = XVALO(VALOBADA("FORMUPAT", "Acepta_Min", CONDI$))
   MAXIMO# = XVALO(VALOBADA("FORMUPAT", "Acepta_Max", CONDI$))
   PREC# = XVALO(VALOBADA("FORMUPAT", "Acepta_Prec", CONDI$))
   '
   MINIMO# = MINIMO# + XVALO(Label3)
   MAXIMO# = MAXIMO# + XVALO(Label3)
   '
   Label40 = FORMATNUM(MINIMO#, "F10." & CNTDC$)
   Label41 = FORMATNUM(MAXIMO#, "F10." & CNTDC$)
   Label37 = FORMATNUM(PREC#, "F10." & CNTDC$)
   '
   'Muestra datos del Remanente
   Label9 = Label16
   Label7 = ""
   CONDI$ = "IDENLOTE ='" & IDLOT$ & "'"
   DEUBI% = XVALO(VALOBADA("LOTINGRE", "DEPOUBIC", CONDI$))
   If DEUBI% = XVALO(Label9) Then
      Label7 = SAFETEXT(VALOBADA("LOTINGRE", "UBICFIS", CONDI$)) 'Mid(List1.List(IndexGuia), 187, 24)  '
   End If
   '
End Sub

Private Sub Text1_Change()
    CIXI% = CODINT(Label20)
    REMAN = XVALO(Label4) - XVALO(Text1)
    If REMAN < 0 Then REMAN = 0
    Label44 = FORMATNUM$(REMAN, "F13." + CNTDC$(CIXI%))
End Sub

Private Sub Text1_LostFocus()
   CIXI% = CODINT(Label20)
   Text1 = FORMATNUM(XVALO(Text1), "F13." & CNTDC$(CIXI%))
End Sub

Private Sub label16_Change()

   Label38 = ECOARCH$("TADEPOSI", Chr$(XVALO(Label16)))
   List1.Clear
   Call Label32_Change
   
End Sub
'
Sub CargarLista()
    '
    Dim Reque As ADODB.Recordset
    Set Reque = New ADODB.Recordset
    DESTINO$ = Label26
    SQLX$ = "SELECT * FROM BALAREQMO "
    SQLX$ = SQLX$ + " WHERE DESTINO ='" & DESTINO$ & "' "
    SQLX$ = SQLX$ + " AND KILO_ASIG > 0 "
    '
    Set Reque = FLEXCONN.Execute(SQLX$)
    List1.Clear

    With Reque
       Do While Not (.EOF)
          LINEA$ = Space$(256)
          CIXI% = XVALO(!Cod_Inte)
          CODIGO$ = CODEXT$(CIXI%)
          DESC$ = DESCRIT0(CIXI%)
          UNI$ = UNIMED$(CIXI%)
          IDLOT$ = SAFETEXT(!idenLote)
          CANTASI# = XVALO(!Kilo_Asig)
          IDREQUE$ = SAFETEXT$(!IdenReque)
          '
          CONDI$ = "Cod_Inte = " & CIXI%
          CONDI$ = CONDI$ + " AND IdenLote = '" & IDLOT$ & "'"
          CONDI$ = CONDI$ + " AND IdenMovi = '" & IDREQUE$ & "'"
          CANYACON# = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", CONDI$, "NOMESS"))
          CANYACON# = CANYACON# - XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", CONDI$, "NOMESS"))
          CANPEN# = CANTASI# - CANYACON#
          If CANPEN# < 0 Then CANPEN# = 0
          '
          Call REPLA(LINEA$, CODIGO$, 1, 16)
          Call REPLA(LINEA$, DESC$, 19, 26)
          Call REPLA(LINEA$, IDLOT$, 47, 16)
          Call REPLA(LINEA$, UNI$, 65, 4)
            CX$ = FORMATSTAN$(CANTASI#, "F12." & CNTDC$, CIXI%)
          Call REPLA(LINEA$, CX$, 69, 12)
            CX$ = FORMATSTAN$(CANYACON#, "F12." & CNTDC$, CIXI%)
          Call REPLA(LINEA$, CX$, 81, 12)
            CX$ = FORMATSTAN$(0, "F12." & CNTDC$, CIXI%)
          Call REPLA(LINEA$, CX$, 93, 12)
          Call REPLA(LINEA$, IDREQUE$, 151, 32)
          List1.AddItem LINEA$
       .MoveNext
      Loop
    End With
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call FRACLECBAR(Text2, Label20, Label19, Label44)
       Call Label19_Change
    End If
End Sub
