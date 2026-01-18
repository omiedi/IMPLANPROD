VERSION 5.00
Begin VB.Form REUBILO09 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Ubicación "
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
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
      Height          =   810
      Left            =   8400
      Picture         =   "REUBILO09.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Graba - Aprueba la Transacción"
      Top             =   5640
      Width           =   830
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
      Left            =   7980
      TabIndex        =   37
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Re-Ubicacion de Lotes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   120
      TabIndex        =   29
      Top             =   5520
      Width           =   8205
      Begin VB.CommandButton Command1 
         Caption         =   "!"
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
         Left            =   7575
         TabIndex        =   52
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text12 
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
         Left            =   1455
         MaxLength       =   16
         TabIndex        =   50
         Top             =   420
         Width           =   1785
      End
      Begin VB.TextBox Text3 
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
         Left            =   5325
         TabIndex        =   44
         Top             =   420
         Width           =   2055
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
         Left            =   7380
         TabIndex        =   43
         Top             =   420
         Width           =   175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Precinto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   51
         Top             =   465
         Width           =   1110
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Nueva Ubicación:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3675
         TabIndex        =   42
         Top             =   465
         Width           =   1770
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   5985
      TabIndex        =   21
      Top             =   3045
      Width           =   2330
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicacion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   35
         Top             =   1575
         Width           =   1035
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   34
         Top             =   980
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Bultos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   33
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Label41 
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
         Left            =   630
         TabIndex        =   24
         Top             =   1260
         Width           =   1275
      End
      Begin VB.Label Label40 
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
         Left            =   630
         TabIndex        =   23
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label39 
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
         Left            =   210
         TabIndex        =   22
         Top             =   1890
         Width           =   1905
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   3570
      TabIndex        =   17
      Top             =   3045
      Width           =   2325
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   32
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricación:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   31
         Top             =   980
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimiento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   105
         TabIndex        =   30
         Top             =   1575
         Width           =   1335
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
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
         Left            =   630
         TabIndex        =   20
         Top             =   1890
         Width           =   1275
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
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
         Left            =   630
         TabIndex        =   19
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
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
         Left            =   630
         TabIndex        =   18
         Top             =   1260
         Width           =   1275
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   105
      TabIndex        =   12
      Top             =   3045
      Width           =   3375
      Begin VB.Label Label1 
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
         Left            =   1995
         TabIndex        =   49
         Top             =   1365
         Width           =   1020
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   840
         TabIndex        =   40
         Top             =   1080
         Width           =   1545
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
         Left            =   735
         TabIndex        =   16
         Top             =   1890
         Width           =   2325
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Envase:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   15
         Top             =   1575
         Width           =   810
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Id.Bulto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   420
         TabIndex        =   14
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Label2 
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
         Left            =   1470
         TabIndex        =   13
         Top             =   630
         Width           =   1545
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
      Left            =   3600
      TabIndex        =   11
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
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
      Height          =   1710
      Left            =   105
      TabIndex        =   3
      Top             =   1260
      Width           =   8205
      Begin VB.Frame Frame1 
         Caption         =   "Frame3"
         Height          =   1220
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame3"
         Height          =   1220
         Left            =   3930
         TabIndex        =   36
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame14 
         Caption         =   "Frame14"
         Height          =   1220
         Left            =   2880
         TabIndex        =   28
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame13 
         Caption         =   "Frame3"
         Height          =   1185
         Left            =   5415
         TabIndex        =   27
         Top             =   405
         Width           =   15
      End
      Begin VB.Frame Frame12 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame11 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1155
         Left            =   5760
         TabIndex        =   6
         Top             =   405
         Width           =   15
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   600
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   6945
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Id.Bulto              F.Venc.   Ubicacion      D    Cantidad"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   600
         TabIndex        =   5
         Top             =   390
         Width           =   6945
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   5490
      Left            =   8400
      ScaleHeight     =   5430
      ScaleWidth      =   1320
      TabIndex        =   0
      Top             =   -20
      Width           =   1380
      Begin VB.CommandButton Command11 
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
         Left            =   120
         Picture         =   "REUBILO09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   690
      End
      Begin VB.CommandButton Command10 
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
         Left            =   120
         Picture         =   "REUBILO09.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   870
         Width           =   690
      End
      Begin VB.CommandButton Command7 
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
         Left            =   120
         Picture         =   "REUBILO09.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   1605
         Width           =   690
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   1
         Top             =   4725
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -285
         Picture         =   "REUBILO09.frx":0A68
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5040
         Width           =   1515
      End
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
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5355
      TabIndex        =   39
      Top             =   375
      Width           =   1545
   End
   Begin VB.Label Label13 
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
      Left            =   7035
      TabIndex        =   38
      Top             =   315
      Width           =   1020
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   210
      TabIndex        =   10
      Top             =   795
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
      Left            =   1470
      TabIndex        =   9
      Top             =   735
      Width           =   6690
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
      Left            =   1470
      TabIndex        =   8
      Top             =   315
      Width           =   2220
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -210
      TabIndex        =   7
      Top             =   375
      Width           =   1545
   End
End
Attribute VB_Name = "REUBILO09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   CIXI% = CODINT(Label28)
   If XVALO(CIXI%) > 0 Then
      If TRIM$(Label2) <> "" Then
         'BUSCA EL DEPOSITO DE ORIGEN DEL LOTE SELECCIONADO
         CONDI$ = "COD_INTE = '" & CIXI% & "' AND IDENLOTE = '" & TRIM$(Label2) & "'"
         DEPP% = XVALO(VALOBADA("LOTINGRE", "DepoUbic", CONDI$, "NOMESS"))
         '
         Call UBIRECOM(CIXI%, TRIM$(Label2), DEPP%)
         Call CommSel("!TAUBICA", Command1, Text3)
      End If
   End If
   '
End Sub

Private Sub Command10_Click()
 '
   Command10.Enabled = False
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
99 Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command2_Click()

   Command2.Enabled = False
   '
   'VALIDACIONES
   FECHA% = FECHANUM(Label13)
   FEFA% = FECHANUM(Label36)
   FEVE% = FECHANUM(Label38)
   CIXI% = CODINT%(Label28)
   IDLOTX$ = TRIM$(Label2)
   UBIANT$ = TRIM$(Label39.Caption)
   '
   If UBIANT$ = "" Then
      Call COMUNI("ATENCION: Lote Sin Ubicación Original.")
   End If
   '
   If CIXI% < 1 Then
      Call MENSERR(24, "Codigo no Válido.")
      GoTo 99
   End If
   '
   If TRIM$(Text3) = "" Then
      COMUNI ("Falta Ingreso de Ubicación")
      Text3.SetFocus
      GoTo 99
   End If
   '
   UBICAIN$ = Mid$(TRIM$(UCase(Text3)), 1, 16)
   If UBICAIN$ <> "" Then
      If TRIM$(VALOBADA("UBIPRE", "Codiubic", "Codiubic = '" & UBICAIN$ & "'")) = "" Then
        Call MENSERR(24, "Ubicacion " + TRIM$(UBICAIN$) + " Inexistente")
        GoTo 99
      End If
   End If
   '
   COD$ = TRIM$(Label28)              'Codigo ext
   CI% = CODINT(COD$)                 'Cod Interno
   POSI$ = TRIM$(Text3)               'Ubicación Propuesta
   CONDI$ = "COD_INTE = " & CI% & " AND IDENLOTE = '" & IDLOTX$ & "'"
   Destino$ = SAFETEXT$(VALOBADA("LOTINGRE", "COD_DESTI", CONDI$))                    'Destino
   '
   If PERMITEUBI%(CI%, Destino$, POSI$) <> 1 Then
      Text3.SetFocus
      GoTo 99
   End If
   '
   DPRI$ = "RU-" + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Mid$(HOS$, 1, 2) & "-00"
   DSEC$ = DPRI$
   '
   CONDI$ = "CODIMOVI = 'RU' AND DOPRILAR = '" & DPRI$ & "'"
   NORIT% = 1 + XVALO(VALOBADA("MOVISTO", "MAX(NUORITEM)", CONDI$))
   '
   CONDI$ = "COD_INTE = " & CI% & " AND IDENLOTE = '" & IDLOTX$ & "'"
   'SE MODIFICA PARA QUE TOME EL DEPOSITO DESDE EL LOTINGRE
'25 DEPX% = XVALO(VALOBADA("MOVISTO", "DEPOMOVI", CONDI$, "NOMESS"))
   DEPX% = XVALO(VALOBADA("LOTINGRE", "DEPOUBIC", CONDI$, "NOMESS"))
   '
   Call MOVISTOK(FECHA%, CIXI%, TRIM$(Label2), "RU", DPRI$, DPRI$, DSEC$, DSEC$, "Re-Ubica en " + Text3, NORIT%, 0, 0, 0, DEPX%, 0)
   '
'   SQLX$ = "Select * from LOTINGRE where IdenLote = '" & IDLOTX$ & "'" 'Selecciona el registro correspondiente al LOTE
'   Dim LOTINGTMP As ADODB.Recordset
'   Set LOTINGTMP = New ADODB.Recordset
'   LOTINGTMP.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'   '
'   With LOTINGTMP
'      !UBICFIS = UBICAIN$
'      .Update
'   End With
'   Set LOTINGTMP = Nothing
'   '
'REEMPLAZA LO ARRIBA COMENTADO
   CONDI$ = "IdenLote = '" & IDLOTX$ & "'"
   Call ACTUREGISTRO("LOTINGRE", "UBICFIS", CONDI$, UBICAIN$, REG&, "NOMESS")
   '
   Call ACTUBIPRE(UBIANT$, UBICAIN$, DEPX%)
   
   '
   Call GENTAUBI
   Call FRESHECHO("TAUBICA")
   '
   LOTE$ = TRIM$(Label2)
   CONDI$ = "IdenLote ='" & LOTE$ & "'"
   Call ACTUREGISTRO("LOTINGRE", "Precinto", CONDI$, Text12, REG&)
   '
   Call COMUNI("Procesado")
   '
   Call LIMPLAB
   Call LIMPTEXT
   Call Label28_Change
 
99 Screen.MousePointer = 1
   Command2.Enabled = True
   '
End Sub


Private Sub Command4_Click()
    '
    If TRIM$(Label2) = "" Then
        Call MENSERR(24, "Debe Seleccionar un Lote")
        Exit Sub
    End If
    '
    'BUSCA EL DEPOSITO DE ORIGEN DEL LOTE SELECCIONADO
    CONDI$ = "COD_INTE = '" & CODINT%(TRIM$(Label28)) & "' AND IDENLOTE = '" & TRIM$(Label2) & "'"
    DEPP% = XVALO(VALOBADA("LOTINGRE", "DepoUbic", CONDI$, "NOMESS"))
    '
    Call GENTAUBI(DEPP%)
    Call SELECHO("!TAUBICA")
    '
    UBISELE$ = VALACT1$("!TAUBICA")
    If UBISELE$ <> "" Then
       Text3 = UBISELE$
    End If
    '
End Sub

Private Sub Command6_Click()

   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text12.TEXT = VDEV$
   End If
   
End Sub

Private Sub Command5_Click()
    Call SELECHO("MASTER")
    If VALACT1$("MASTER") <> "" Then
       Label28 = VALACT1$("MASTER")
    End If
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   FIKARDEX09.Text1 = TRIM$(Label28)
   FIKARDEX09.Show 1
   Command7.Enabled = True
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
    Call VERJOBID(OKEY%)
    
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    Call ABRECONEXION
    HOII% = HOY(HOS$)
    Label13 = HOS$
    '
    If Left$(APLINVO$, 7) = "BORECEP" Then ' Or INTERPRE% = 1 Then
       Text1 = Mid$(APLINVO$, 8)
       openforms = DoEvents
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

Private Sub Label28_Change()
    CIXI% = CODINT%(Label28)
    Label30 = DESCRIT0$(CIXI%)
    Call LIMPLAB
    Call LIMPTEXT
    Call ARMALILOTES(CIXI%)
End Sub

Sub ARMALILOTES(CIXI%)
    '
    CIXI% = CODINT(Label28)
    '
    FLEXCONN.Execute (FILTSQL$("SET DATEFORMAT dmy"))
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM LOTINGRE WITH(NOLOCK)"
'    SQLX$ = SQLX$ + " INNER JOIN BOLRECEP"
'    SQLX$ = SQLX$ + " ON LOTINGRE.IDENLOTE = BOLRECEP.IDENLOTE"
    SQLX$ = SQLX$ + " WHERE LOTINGRE.Cod_Inte = " & CIXI%
    SQLX$ = SQLX$ + " AND (CanAlta - CanConsu) > 0"    ' no trae los consumidos
    SQLX$ = SQLX$ + " ORDER BY FechAlta ASC, LOTINGRE.IdenLote ASC "
    '
    Dim LOTINGTEMP As ADODB.Recordset
    Set LOTINGTEMP = New ADODB.Recordset
    Set LOTINGTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    REBLA$ = Space$(256)
    JJ& = 0: List1.Clear
    With LOTINGTEMP
25     On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Call MENSERR(24, "No Figuran Lotes Disponibles\para el Material Elegido.")
         GoTo 99
       End If
       '
       On Error GoTo 0
       Do While Not .EOF
         IDLOT$ = TRIM$(!idenLote)
         '
         TTXX$ = REBLA$
         Call REPLA(TTXX$, IDLOT$, 1, 20) 'lote
         FEX = CVINT(!FECHVENC)
         Call REPLA(TTXX$, FECHATEX$(FEX), 23, 8)                          ' Fecha vencimiento
         Call REPLA(TTXX$, SAFETEXT(!UBICFIS), 33, 14)                      ' Ubicacion
         Call REPLA(TTXX$, FORMATNUM$(XVALO(!CANALTA - !CanConsu), "F13." & CNTDC$(CIXI%)), 48, 13)  ' Kilos
         Call REPLA(TTXX$, XVALO(!DEPOUBIC), 47, 2)                         ' Deposito
         Call REPLA(TTXX$, SAFETEXT(!CodEnvase), 72, 25)                    ' Envase
         FEXALTA = CVINT(!FECHALTA)
         Call REPLA(TTXX$, FECHATEX$(FEXALTA), 97, 8)                       ' fecha Ingreso
         FEXFAB = CVINT(!FechFab)
         Call REPLA(TTXX$, FECHATEX$(FEXFAB), 105, 8)                       ' Fecha fabricación
            COEFI = 1: If !CANALTA > 0 Then COEFI = !Cantidad / !CANALTA
         Call REPLA(TTXX$, FORMATNUM(XVALO(COEFI * (!CANALTA - !CanConsu)), "F13." & CNTDC$), 113, 13) ' Cantidad
         '
         JJ& = JJ& + 1
         List1.AddItem TTXX$
       .MoveNext
       Loop
    End With
    '
99  LOTINGTEMP.Close
    Set LOTINGTEMP = Nothing
    '
End Sub

Private Sub Label28_DblClick()
    '
10  TTXX$ = InputBox$("Pinche sobre Barras", "Lectura Código de Barras")
    If Len(TTXX$) > 16 Then
       FRABOBI09.Label28 = TRIM$(Left$(TTXX$, 16))
       For KKI% = 1 To List1.ListCount
          LOTSE$ = TRIM$(Mid$(TTXX$, 17, 12))
          If TRIM$(Left$(List1.List(KKI% - 1), 12)) = LOTSE$ Then
             List1.Selected(KKI% - 1) = True
             Call List1_Click
             Exit Sub
          End If
       Next KKI%
       Call MENSERR(24, "Lectura no Valida")
       GoTo 10
    End If
    '
End Sub
'
Private Sub List1_Click()
    
    Call LIMPTEXT
    '
    TTXX$ = List1.TEXT
    '
    NCR$ = Mid(TTXX$, 1, 20)
    COD$ = Label28
    CI% = CODINT%(COD$)
    Label2.Caption = NCR$
    Label24.Caption = TRIM$(Mid$(TTXX$, 72, 25)) ' Envase
    Label37.Caption = TRIM$(Mid$(TTXX$, 97, 8))   ' F. Ingreso
    Label36.Caption = TRIM$(Mid$(TTXX$, 105, 8))  ' F. Fabricación
    Label38.Caption = TRIM$(Mid$(TTXX$, 23, 8))   ' F. Vencimiento
    Label40.Caption = TRIM$(Mid$(TTXX$, 113, 13))  ' Cantidad
    Label41.Caption = TRIM$(Mid$(TTXX$, 48, 13))  ' Kilos
    Label39.Caption = TRIM$(Mid$(TTXX$, 33, 14))  ' Ubicación
    Label34 = Label39.Caption
    Label33 = Label40
    '
    CONDI$ = "IdenLote = '" & NCR$ & "'"
    Label1 = TRIM$(VALOBADA("LOTINGRE", "Libera", CONDI$))
    Text12 = TRIM$(VALOBADA("LOTINGRE", "Precinto", CONDI$))
    '
    NCR$ = TRIM$(NCR$)
    PPXX% = InStr(NCR$, "/")
    If PPXX% > 0 Then
       NCR$ = Left$(NCR$, PPXX% - 1) ' VERIFICA LA EXISTENCIA DE LA BARRA IDENTIFICADORA DE SUBLOTE
    End If
    '
   
End Sub

Sub LIMPLAB()
    Label2.Caption = ""
    Label24.Caption = ""
    Label36.Caption = ""
    Label37.Caption = ""
    Label38.Caption = ""
    Label39.Caption = ""
    Label40.Caption = ""
    Label41.Caption = ""
    Label1.Caption = ""
End Sub

Sub LIMPTEXT()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    Text8 = ""
    Text9 = ""
    Text14 = ""
    Text11 = ""
    Text16 = ""
    Text12 = ""
    Text17 = ""
    Call BLANARCH("!CARFRABO")
End Sub

Private Sub Text12_GotFocus()
    Text12.SelStart = 0
    Text12.SelLength = Text12.MaxLength
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text3.SetFocus
    End If
End Sub

Private Sub Text3_DblClick()
    Call Command4_Click
End Sub

Sub UBIRECOM1(CIXI%, LOTE$)
    '
    UBIPOSI$ = ""
    '
    DESTI$ = TRIM$(VALOBADA("LOTINGRE", "COD_DESTI", "IDENLOTE ='" & TRIM$(Label2) & "'"))
    If DESTI$ = "" Then GoTo 40
    
    'Ubicaciones por Mismo Destino
    CONDI$ = "SELECT * FROM LOTINGRE WHERE COD_DESTI ='" & DESTI$ & "'"
    Dim UBICA As ADODB.Recordset
    Set UBICA = FLEXCONN.Execute(FILTSQL$(CONDI$))
    Do While Not (UBICA.EOF)
       'GENERACION DE UBICACIONES
       UBIPO$ = AJUSTI$(UBICA!UBICFIS, 12)
       If TRIM$(UBIPO$) <> "" Then
          If InStr(UBIPOSI$, UBIPO$) < 1 Then
             UBIPOSI$ = UBIPOSI$ + UBIPO$
          End If
       End If
       UBICA.MoveNext
    Loop
    '
    UBICA.Close
    Set UBICA = Nothing
    
40  'Ubicaciones por Mismo Material
    CONDI$ = "SELECT DISTINCT(UBICFIS)AS UBICACION  from lotingre where COD_INTE = " & CIXI% & " AND UBICFIS <> ''"
    Set UBICA = FLEXCONN.Execute(FILTSQL$(CONDI$))
    Do While Not (UBICA.EOF)
       'GENERACION DE UBICACIONES
       UBIPO$ = AJUSTI$(UBICA!UBICACION, 12)
       If InStr(UBIPOSI$, UBIPO$) < 1 Then
          UBIPOSI$ = UBIPOSI$ + UBIPO$
       End If
    UBICA.MoveNext
    Loop
    '
    UBICA.Close
    Set UBICA = Nothing
    '
    
    '1era Ubicacion Libre de cada Referencia
    REFANT$ = ""
    SQLX$ = "SELECT * FROM UBIPRE ORDER BY REFERENCIA"
    Dim UBIPRETMP As ADODB.Recordset
    Set UBIPRETMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With UBIPRETMP
      Do While Not .EOF
        REFACT$ = TRIM$(!Referencia)
        If REFACT$ <> REFANT$ Then
           UBIX$ = !CODIUBIC
           If CantRegistros("LOTINGRE", "UBICFIS = '" & UBIX$ & "'") < 1 Then
              REFANT$ = REFACT$
              UBIPO$ = AJUSTI$(UBIX$, 12)
              If InStr(UBIPOSI$, UBIPO$) < 1 Then
                 UBIPOSI$ = UBIPOSI$ + UBIPO$
              End If
           End If
        End If
        .MoveNext
      Loop
    End With
    '
    'Se Genera el Archivo TAUBICA
    JJ& = 0
    For U& = 1 To ULTREG&("TAUBICA")
       XX$ = REGLEIDO$("TAUBICA", U&)
       CODU$ = AJUSTI$(XX$, 12)
       If InStr(UBIPOSI$, CODU$) > 0 Then
          JJ& = JJ& + 1
          Call GRAREG("!TAUBICA", XX$, JJ&)
       End If
    Next U&
    Call SETULTREG("!TAUBICA", JJ&)
    Call FRESHECHO("!taubica")
    '
End Sub


