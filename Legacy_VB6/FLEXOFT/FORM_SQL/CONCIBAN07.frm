VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.ocx"
Begin VB.Form CONCIBAN07 
   BackColor       =   &H0095C5E1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conciliación Bancaria"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   495
      Left            =   9975
      TabIndex        =   14
      Top             =   0
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   873
      ButtonWidth     =   794
      ButtonHeight    =   714
      Appearance      =   1
      ImageList       =   "imlToolbarPics"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Search"
            Object.ToolTipText     =   "Busqueda por Texto"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.ToolTipText     =   "Listado por Rango de Fechas"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "EXIT"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   44
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
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
      Left            =   10395
      TabIndex        =   39
      ToolTipText     =   "Procesa y Marca los Movimientos Conciliados"
      Top             =   5460
      Width           =   1380
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H0095C5E1&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   3690
      Left            =   105
      TabIndex        =   16
      Top             =   5460
      Width           =   11670
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   10530
         ScaleHeight     =   75
         ScaleWidth      =   840
         TabIndex        =   42
         Top             =   420
         Width           =   900
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H0095C5E1&
         Caption         =   "BUSCAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   10290
         TabIndex        =   38
         Top             =   735
         Width           =   1380
         Begin VB.CommandButton Command2 
            Caption         =   "Buscar"
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
            Left            =   105
            TabIndex        =   41
            ToolTipText     =   "Busca por Texto en Pendientes de Conciliación"
            Top             =   770
            Width           =   1170
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   105
            TabIndex        =   40
            Top             =   360
            Width           =   1170
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Frame1"
         Height          =   2280
         Left            =   2700
         TabIndex        =   25
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame15 
         Caption         =   "Frame8"
         Height          =   2280
         Left            =   6825
         TabIndex        =   24
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame14 
         Caption         =   "Frame10"
         Height          =   2280
         Left            =   3850
         TabIndex        =   23
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame13 
         Caption         =   "Frame9"
         Height          =   2280
         Left            =   8400
         TabIndex        =   22
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame11 
         Caption         =   "Frame9"
         Height          =   2280
         Left            =   9870
         TabIndex        =   21
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame3"
         Height          =   2280
         Left            =   1010
         TabIndex        =   20
         Top             =   315
         Width           =   15
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
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
         Height          =   2190
         ItemData        =   "CONCIBAN07.frx":0000
         Left            =   0
         List            =   "CONCIBAN07.frx":0002
         TabIndex        =   17
         Top             =   420
         Width           =   10135
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   0
         ScaleHeight     =   420
         ScaleWidth      =   10110
         TabIndex        =   18
         Top             =   0
         Width           =   10135
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   $"CONCIBAN07.frx":0004
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
            Left            =   30
            TabIndex        =   19
            Top             =   105
            Width           =   11355
         End
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Pendiente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10290
         TabIndex        =   37
         Top             =   2100
         Width           =   2850
      End
      Begin VB.Label Label15 
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
         Height          =   285
         Left            =   10290
         TabIndex        =   36
         Top             =   2330
         Width           =   1350
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame9"
      Height          =   2800
      Left            =   11520
      TabIndex        =   15
      Top             =   2030
      Width           =   15
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   2800
      Left            =   1100
      TabIndex        =   13
      Top             =   2030
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   2800
      Left            =   2785
      TabIndex        =   12
      Top             =   2030
      Width           =   15
   End
   Begin VB.Frame Frame9 
      Caption         =   "Frame9"
      Height          =   2800
      Left            =   9960
      TabIndex        =   8
      Top             =   2030
      Width           =   15
   End
   Begin VB.Frame Frame10 
      Caption         =   "Frame10"
      Height          =   2800
      Left            =   6915
      TabIndex        =   9
      Top             =   2030
      Width           =   15
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   2800
      Left            =   8490
      TabIndex        =   7
      Top             =   2030
      Width           =   15
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2800
      Left            =   3940
      TabIndex        =   6
      Top             =   2030
      Width           =   15
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00D0E0D0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   3165
      Left            =   105
      TabIndex        =   0
      Top             =   1680
      Width           =   11670
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
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
         Height          =   2730
         ItemData        =   "CONCIBAN07.frx":00BE
         Left            =   0
         List            =   "CONCIBAN07.frx":00C0
         TabIndex        =   2
         Top             =   420
         Width           =   11670
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   0
         ScaleHeight     =   420
         ScaleWidth      =   11640
         TabIndex        =   3
         Top             =   0
         Width           =   11670
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   $"CONCIBAN07.frx":00C2
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
            Height          =   375
            Left            =   30
            TabIndex        =   4
            Top             =   100
            Width           =   11595
         End
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   11760
         Y1              =   3130
         Y2              =   3130
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "CONCIBAN07.frx":0181
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CONCIBAN07.frx":03B5
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label14 
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
      Height          =   285
      Left            =   9960
      TabIndex        =   35
      Top             =   4935
      Width           =   1570
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo Conciliado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8250
      TabIndex        =   34
      Top             =   5020
      Width           =   1650
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "PENDIENTES DE CONCILIACIÓN"
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
      Left            =   105
      TabIndex        =   33
      Top             =   5145
      Width           =   4005
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "MOVIMIENTOS CONCILIADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   105
      TabIndex        =   32
      Top             =   1365
      Width           =   4005
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
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
      Left            =   3480
      TabIndex        =   31
      Top             =   1120
      Width           =   1065
   End
   Begin VB.Label Label8 
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
      Height          =   285
      Left            =   4605
      TabIndex        =   30
      Top             =   1050
      Width           =   1035
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta"
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
      Left            =   5850
      TabIndex        =   29
      Top             =   1120
      Width           =   705
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
      Height          =   285
      Left            =   6600
      TabIndex        =   28
      Top             =   1050
      Width           =   1035
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo Inicial"
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
      Left            =   8600
      TabIndex        =   27
      Top             =   1120
      Width           =   1320
   End
   Begin VB.Label Label3 
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
      Height          =   285
      Left            =   9960
      TabIndex        =   26
      Top             =   1050
      Width           =   1515
   End
   Begin ComctlLib.ImageList imlToolbarPics 
      Left            =   8925
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483634
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483644
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   14
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONCIBAN07.frx":0413
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONCIBAN07.frx":072D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONCIBAN07.frx":0A47
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONCIBAN07.frx":0D61
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONCIBAN07.frx":107B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONCIBAN07.frx":1395
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONCIBAN07.frx":16AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONCIBAN07.frx":19C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONCIBAN07.frx":1CE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONCIBAN07.frx":1FFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONCIBAN07.frx":2317
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONCIBAN07.frx":24F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONCIBAN07.frx":2603
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "CONCIBAN07.frx":291D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Conciliación Bancaria"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   4275
   End
   Begin VB.Label Label4 
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
      Height          =   285
      Left            =   6615
      TabIndex        =   10
      Top             =   630
      Width           =   4860
   End
   Begin VB.Label Ltext2 
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
      Height          =   285
      Left            =   6615
      TabIndex        =   5
      Top             =   225
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
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
      Left            =   5280
      TabIndex        =   1
      Top             =   300
      Width           =   1275
   End
End
Attribute VB_Name = "CONCIBAN07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   CUEBA$ = Ltext2
   CUEII% = CUEINT%(CUEBA$)
   '
   Call APERBASDAT("CABAN")
   Call ABREPLANCUEN
   '
   FIN& = List1.ListCount
   HAYDESCON% = 0
   'MsgBox CStr(FIN&)
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     List1.ListIndex = U& - 1
     TTXX$ = List1.List(U& - 1)
     If TRIM$(Mid$(TTXX$, 128, 1)) <> "" Then    ' de una conciliacion anterior
       If Mid$(TTXX$, 27, 8) <> " ****** " Then  ' tampoco es des-conciliación
         GoTo 29
       End If
     End If
     '
     FEXI% = FECHANUM(Left$(TTXX$, 8))
     FECONTA1 = FETEXCOR1$(FEXI%)
     IDASIE$ = TRIM$(Mid$(TTXX$, 11, 15))
     If Mid$(TTXX$, 27, 8) = " ****** " Then
          FEACRED = 0
          HAYDESCON% = 1
        Else
          FEACRED = CVDAT(FECHANUM(Mid$(TTXX$, 27, 8)))
     End If
     REFEMOV$ = Mid$(TTXX$, 38, 28)
     IDEBE# = Val(Mid$(TTXX$, 66, 14))
     IHABE# = Val(Mid$(TTXX$, 80, 14))
     ORIMO$ = UCase$(TRIM$(Mid$(TTXX$, 121, 6)))
     '
     If ORIMO$ = "CABAN" Then
          SQLX$ = "SELECT * FROM CAJABANC "
          SQLX$ = SQLX$ + "WHERE FechConta = '" & FECONTA1 & "' "
          SQLX$ = SQLX$ + "AND CodiCom_Cor = '" & IDASIE$ & "' "
          SQLX$ = SQLX$ + "AND CueContaI = " & CUEII% & " "
          SQLX$ = SQLX$ + "AND LEFT(DescriMov, 28) = '" & REFEMOV$ & "' "
          '
          Dim CAJJJABANC As ADODB.Recordset
          Set CAJJJABANC = New ADODB.Recordset
27        On Error Resume Next
          CAJJJABANC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 27
          End If
          On Error GoTo 0
          With CAJJJABANC    ' CajaBanc
            On Error Resume Next
            .MoveFirst
25          If Err < 1 Then
              On Error GoTo 0
              CAMOCO% = 0
              Do While Not .EOF
                If Abs(IDEBE# - !idebecomp) < 0.005 Then
                  If Abs(IHABE# - !ihabecomp) < 0.005 Then
                    If FEACRED > 0 Then
                        !MARCONSI = "*"
                        !fechacred = FEACRED
                        CAMOCO% = CAMOCO% + 1
                        If CAMOCO% > 1 Then
                           Call COMUNI("Movimientos Adicionales Conciliados\con Igual Importe e Identificador.")
                        End If
                      Else
                        !MARCONSI = " "
                        CAMOCO% = CAMOCO% + 1
                        If CAMOCO% > 1 Then
                           Call COMUNI("Movimientos Adicionales Des-Conciliados\con Igual Importe e Identificador.")
                        End If
                    End If
                    .Update
                    If Err > 0 Then
                        Call MENSERR(24, "Problemas al Procesar '" & IDASIE$ & "'.\" & Err.Description)
                    End If
                    'GoTo 29    ' suprimido 11/05/09 por problema guidi
                    '           ' de tener dos cheques con el mismo
                    '           ' numero e importe para conciliar
                  End If
                End If
                .MoveNext
              Loop
            End If
            On Error GoTo 0
          End With
          
       ElseIf ORIMO$ = "DIAR" Then
          VUELTA% = 0
          SQLX$ = "SELECT * FROM DetaAsien "
          SQLX$ = SQLX$ + "WHERE FechAsien = '" & FECONTA1 & "' "
          ' SQLX$ = SQLX$ + "AND SucuAsien = " & SUCUI% & ""
          SQLX$ = SQLX$ + "AND NuIdAsien = '" & IDASIE$ & "' "
          SQLX$ = SQLX$ + "AND CueCoInt = " & CUEII% & " "
          SQLX$ = SQLX$ + "AND LEFT(RefePase, 28) = '" & REFEMOV$ & "' "
          SQLX$ = SQLX$ + "AND StatPase > 0 "
          '
          Dim DETTTASIEN As ADODB.Recordset
          Set DETTTASIEN = New ADODB.Recordset
28        On Error Resume Next
          DETTTASIEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 28
          End If
          On Error GoTo 0
          With DETTTASIEN
            On Error Resume Next
            .MoveFirst
            If Err < 1 Then
              CAMOCO% = 0

              Do While Not .EOF
                On Error GoTo 0
                If Abs(IDEBE# - !idebepase) < 0.005 Then
                  If Abs(IHABE# - !ihabepase) < 0.005 Then
                    If FEACRED > 0 Then
                        !MARCONSI = "*"
                        CAMOCO% = CAMOCO% + 1
                        If CAMOCO% > 1 Then
                           Call COMUNI("Movimientos Adicionales Conciliados\con Igual Importe e Identificador.")
                        End If
                      Else
                        !MARCONSI = " "
                        CAMOCO% = CAMOCO% + 1
                        If CAMOCO% > 1 Then
                           Call COMUNI("Movimientos Adicionales Des-Conciliados\con Igual Importe e Identificador.")
                        End If
                    End If
                    .Update
                    If Err > 0 Then
                        Call MENSERR(24, "Problemas al Procesar '" & IDASIE$ & "'.\" & Err.Description)
                    End If
                  End If
                End If
                .MoveNext
              Loop
              GoTo 29
            End If
            On Error GoTo 0
          End With
    End If
    'Text1 = CStr(U&)
     '
29 Next U&
          
''''       ElseIf ORIMO$ = "DIAR" Then
''''          SQLX$ = "SELECT * FROM DetaAsien "
''''          SQLX$ = SQLX$ + "WHERE FechAsien = '" & FECONTA1 & "' "
''''          ' SQLX$ = SQLX$ + "AND SucuAsien = " & SUCUI% & ""
''''          SQLX$ = SQLX$ + "AND NuIdAsien = '" & IDASIE$ & "' "
''''          SQLX$ = SQLX$ + "AND CueCoInt = " & CUEII% & " "
''''          SQLX$ = SQLX$ + "AND LEFT(RefePase, 28) = '" & REFEMOV$ & "' "
''''          SQLX$ = SQLX$ + "AND StatPase > 0 "
''''          '
''''          Dim DETTTASIEN As ADODB.Recordset
''''          Set DETTTASIEN = New ADODB.Recordset
''''28        On Error Resume Next
''''          DETTTASIEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
''''          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
''''             On Error GoTo 0
''''             Call CapturaErrorOpen
''''             GoTo 28
''''          End If
''''          On Error GoTo 0
''''          With DETTTASIEN
''''            On Error Resume Next
''''            .MoveFirst
''''            If Err < 1 Then
''''              Do While Not .EOF
''''                On Error GoTo 0
''''                If Abs(IDEBE# - !IDebePase) < 0.005 Then
''''                  If Abs(IHABE# - !IHabePase) < 0.005 Then
''''                    If FEACRED > 0 Then
''''                        !MARCONSI = "*"
''''                      Else
''''                        !MARCONSI = " "
''''                    End If
''''                    .Update
''''                    GoTo 29
''''                  End If
''''                End If
''''                .MoveNext
''''              Loop
''''            End If
''''            On Error GoTo 0
''''          End With
''''    End If
''''     '
''''29 Next U&
   '
   'MsgBox "REMUEVE ITEMS"
   Set CAJJJABANC = Nothing
   Set DETTTASIEN = Nothing
   If HAYDESCON% > 0 Then
35   FIN& = List1.ListCount
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       TTXX$ = List1.List(U& - 1)
       If Mid$(TTXX$, 27, 8) = " ****** " Then
         List1.RemoveItem (U& - 1)
         GoTo 35
       End If
       'Text1 = CStr(FIN&)
     Next U&
   End If
   '
   Call COMUNI("Se Ha Completado la Conciliacion Bancaria")
99 Command1.Enabled = True
   '
End Sub

Private Sub Command2_Click()
   ABUS$ = UCase$(TRIM$(Text1))
   If ABUS$ = "" Then
     Call COMUNI("Introduzca Argumento de Búsqueda")
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
   '
   If ABUS$ <> "" Then
     INI& = List2.ListIndex
10   For KU& = INI& + 1 To List2.ListCount - 1
       If InStr(UCase$(List2.List(KU&)), ABUS$) > 0 Then
         List2.ListIndex = KU&
         Exit Sub
       End If
     Next KU&
     If INI& > (-1) Then
       INI& = (-1)
       GoTo 10
     End If
     Call COMUNI("Argumento no Encontrado")
   End If
End Sub


Private Sub Form_Load()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
End Sub

Private Sub List1_DblClick()
   REAC& = List1.ListIndex
   TTXX$ = List1.TEXT
   If Mid$(TTXX$, 27, 8) = " ****** " Then Exit Sub
   HAYDESCON% = 0
   If TRIM$(Mid$(TTXX$, 128, 1)) <> "" Then
     'Call MENSERR(24, "Movimiento Previamente Conciliado.\Imposible Revertir Conciliación.")
     IDMO$ = TRIM$(Mid$(TTXX$, 11, 16))
     TTYY$ = "Movimiento '" + IDMO$ + "' Previamente Conciliado.\   \Confirma Revertir la Conciliación Previa\de Este Movimiento ?\\No, Conservar\Si, Revertir"
     If COMALTER%(TTYY$) = 2 Then
         List1.RemoveItem REAC&
         TTZZ$ = TTXX$
         Call REPLA(TTZZ$, " ****** ", 27, 8)
         List1.AddItem TTZZ$, REAC&
         HAYDESCON% = 1
         GoTo 35
       Else
         Exit Sub
     End If
   End If
   '
   List1.RemoveItem REAC&
35 FACRE% = FECHANUM(Mid$(TTXX$, 27, 8))
   For KU& = 1 To List2.ListCount
     If FECHANUM(Mid$(List2.List(KU& - 1), 27, 8)) > FACRE% Then
       List2.AddItem TTXX$, KU& - 1
       GoTo 40
     End If
   Next KU&
   List2.AddItem TTXX$
   '
40 TRANSPOR# = Val(Label3)
   For KU& = 1 To List1.ListCount
     TTXX$ = List1.List(KU& - 1)
     If Mid$(TTXX$, 27, 8) <> " ****** " Then
       TRANSPOR# = TRANSPOR# + Val(Mid$(TTXX$, 66, 14)) - Val(Mid$(TTXX$, 80, 14))
     End If
     Call REPLA(TTXX$, FORMATNUM$(TRANSPOR#, "F14.2"), 94, 14)
     List1.RemoveItem (KU& - 1)
     List1.AddItem TTXX$, KU& - 1
   Next KU&
   Label14 = FORMATNUM(TRANSPOR#, "F12.2")
   '
   TRANSPOR# = 0
   For KU& = 1 To List2.ListCount
     TTXX$ = List2.List(KU& - 1)
     TRANSPOR# = TRANSPOR# + Val(Mid$(TTXX$, 66, 14)) - Val(Mid$(TTXX$, 80, 14))
   Next KU&
   Label15 = FORMATNUM(TRANSPOR#, "F12.2")
   '
   If HAYDESCON% > 0 Then
     If COMALTER%("Se Recomienda Ahora\Procesar la Conciliación.\\Procesar Ahora\Continuar") <> 2 Then
       Call Command1_Click
     End If
   End If
   '
End Sub

Private Sub List2_DblClick()
   TTXX$ = List2.TEXT
   REAC& = List2.ListIndex
   '
   VDEF$ = Space$(44) + String$(20, 0)
   Call REPLA(VDEF$, Mid$(TTXX$, 11, 15), 1, 16)
   Call REPLA(VDEF$, Mid$(TTXX$, 38, 28), 17, 28)
   Call REPLA(VDEF$, MKD$(Val(Mid$(TTXX$, 66, 14))), 45, 8)
   Call REPLA(VDEF$, MKD$(Val(Mid$(TTXX$, 80, 14))), 53, 8)
   Call REPLA(VDEF$, MKI$(FECHANUM(Mid$(TTXX$, 1, 8))), 61, 2)
   Call REPLA(VDEF$, MKI$(FECHANUM(Mid$(TTXX$, 27, 8))), 63, 2)
   FACREA% = CVI(Mid$(VDEF$, 63, 2))
   '
20 OBX$ = OBJPLA$("AJUFACRE", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   ' AQUI VALIDACIONES
   '
   DES% = FECHANUM(Label8)
   Has% = FECHANUM(Label6)
   FACRE% = CVI(Mid$(OBX$, 63, 2))
   If FACRE% < DES% Or FACRE% > Has% Then
     Call MENSERR(24, "La Fecha de Acreditación Debe Estar\Entre " + Label8 + " y " + Label6)
     GoTo 20
   End If
   '
   If UCase$(TRIM$(Mid$(TTXX$, 121, 6))) = "DIAR" Then
     If FACRE% <> FACREA% Then
       Call MENSERR(24, "No Válido Cambiar Fecha de Acreditación\en Asientos de Diario.")
       GoTo 20
     End If
   End If
   '
   FEX = FACRE%
   Call REPLA(TTXX$, FECHATEX$(FEX), 27, 8)
   For KU& = List1.ListCount To 1 Step -1
     If FECHANUM(Mid$(List1.List(KU& - 1), 27, 8)) <= FACRE% Then
       List1.AddItem TTXX$, KU&
       REINSE& = KU&
       GoTo 40
     End If
   Next KU&
   List1.AddItem TTXX$, 0
   REINSE& = 0
   '
40 If REINSE& < 1 Then
       TRANSPOR# = Val(Label3)
     Else
       TRANSPOR# = Val(Mid$(List1.List(REINSE& - 1), 94, 14))
   End If
   '
   For KU& = REINSE& + 1 To List1.ListCount
     TTXX$ = List1.List(KU& - 1)
     TRANSPOR# = TRANSPOR# + Val(Mid$(TTXX$, 66, 14)) - Val(Mid$(TTXX$, 80, 14))
     Call REPLA(TTXX$, FORMATNUM$(TRANSPOR#, "F14.2"), 94, 14)
     List1.RemoveItem (KU& - 1)
     List1.AddItem TTXX$, KU& - 1
   Next KU&
   Label14 = FORMATNUM(TRANSPOR#, "F12.2")
   On Error Resume Next
   List1.ListIndex = REINSE&
   On Error GoTo 0
   '
   List2.RemoveItem REAC&
   '
   TRANSPOR# = 0
   For KU& = 1 To List2.ListCount
     TTXX$ = List2.List(KU& - 1)
     TRANSPOR# = TRANSPOR# + Val(Mid$(TTXX$, 66, 14)) - Val(Mid$(TTXX$, 80, 14))
   Next KU&
   Label15 = FORMATNUM(TRANSPOR#, "F12.2")
   '
End Sub

Private Sub Ltext2_Change()
   If Ltext2 = TRIM$(CUECARCHE$) Then
     Command1.Visible = True   ' PARA HABILITAR CONCILIACIóN DE CARTERA
   End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call Command2_Click
   End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
   '
   Select Case UCase$(Button.Key)
      Case "SEARCH"
        Call Command2_Click
      Case "PRINT"
        Call LICONCIPRIN
      Case "EXIT"
        Unload Me
   End Select
   '
End Sub
'
Sub LICONCIPRIN()
  '\\\OT-05-0498-13/07/05///
  DES% = FECHANUM(Label8)
  Has% = FECHANUM(Label6)
  If Has% < DES% Then
    Call MENSERR(24, "Rango de Fechas no procesable.")
    GoTo 99
  End If
  CORRESP$ = FECHATEX$(DES%) + " al " + FECHATEX$(Has%) + " - " + TRIM$(Label4) + " (" + TRIM$(Ltext2) + ")"
  '
15 JJ& = 0
  YATRAN% = 0
  TRANSPO# = Val(Label3)
  '
  For U& = 1 To List1.ListCount
    TTXX$ = List1.List(U& - 1)
    FEMO% = FECHANUM(Left$(TTXX$, 8))
    IDPAS$ = TRIM(Mid$(TTXX$, 9, 12))
    'FA1$ = Mid$(TTXX$, 27, 8)
    FEAC% = FECHANUM(Mid$(TTXX$, 27, 8))
  If FEAC% <= 0 Then GoTo 29
    IDEBE# = Val(Mid$(TTXX$, 66, 14))
    IHABE# = Val(Mid$(TTXX$, 80, 14))
   If U& = 1 Then
     If TRIM$(Mid$(TTXX$, 66, 28)) = "" Then
       IDEBE# = Val(Mid$(TTXX$, 94, 14))
       If IDEBE# < 0 Then
         IHABE# = Abs(IDEBE#)
         IDEBE# = 0
       End If
     End If
   End If
   '
   'If FEMO% < DES% Then
   '    TRANSPO# = TRANSPO# + IDEBE# - IHABE#
   '    GoTo 29
   '  ElseIf FEMO% <= HAS% Then
       If YATRAN% < 1 Then
         If Abs(TRANSPO#) >= 0.005 Then
           XX$ = REGBLAN$("CONCIBAF")
            Desde = DES%
           Call REPLA(XX$, MKI$(DIANTE(Desde)), 19, 2)
           Call REPLA(XX$, "Saldo Anterior", 21, 17)
           Call REPLA(XX$, MKD$(TRANSPO#), 85, 8)
           JJ& = JJ& + 1
           Call GRAREG("CONCIBAF", XX$, JJ&)
           '
           XX$ = REGBLAN$("CONCIBAF")
           JJ& = JJ& + 1
           Call GRAREG("CONCIBAF", XX$, JJ&)
           YATRAN% = 1
         End If
       End If
         '
         FFAX$ = Left$(TTXX$, 8)
       Call REPLA(XX$, MKI(FECHANUM(FFAX$)), 1, 2)
         COCOM$ = Mid$(TTXX$, 10, 16)
       Call REPLA(XX$, COCOM$, 3, 16)
        FFAC$ = Mid$(TTXX$, 27, 8)
       Call REPLA(XX$, MKI(FECHANUM(FFAC$)), 19, 2)
         REFE$ = Mid$(TTXX$, 35, 30)
       Call REPLA(XX$, REFE$, 21, 48)
       Call REPLA(XX$, MKD$(IDEBE#), 69, 8)
       Call REPLA(XX$, MKD$(IHABE#), 77, 8)
         ITRAN# = Val(Mid$(TTXX$, 94, 14))
       Call REPLA(XX$, MKD$(ITRAN#), 85, 8)
       JJ& = JJ& + 1
       Call GRAREG("CONCIBAF", XX$, JJ&)
   'End If
29 Next U&
  'OT-05-0498-WAR-08/08/05
   XX$ = REGBLAN$("CONCIBAF")
   JJ& = JJ& + 1
   Call GRAREG("CONCIBAF", XX$, JJ&)
   For J1& = 1 To List2.ListCount
    TTXX$ = List2.List(J1& - 1)
    FEMO% = FECHANUM(Left$(TTXX$, 8))
    IDPAS$ = TRIM(Mid$(TTXX$, 9, 12))
    FEAC% = FECHANUM(Mid$(TTXX$, 27, 8))
  If FEAC% <= 0 Then GoTo 30
    IDEBE# = Val(Mid$(TTXX$, 66, 14))
    IHABE# = Val(Mid$(TTXX$, 80, 14))
   If J1& = 1 Then
     If TRIM$(Mid$(TTXX$, 66, 28)) = "" Then
       IDEBE# = Val(Mid$(TTXX$, 94, 14))
       If IDEBE# < 0 Then
         IHABE# = Abs(IDEBE#)
         IDEBE# = 0
       End If
     End If
   End If
    '
    XX$ = REGBLAN$("CONCIBAF")
    '
      FFAX$ = Left$(TTXX$, 8)
    Call REPLA(XX$, MKI(FECHANUM(FFAX$)), 1, 2)
      COCOM$ = Mid$(TTXX$, 10, 16)
    Call REPLA(XX$, COCOM$, 3, 16)
     FFAC$ = Mid$(TTXX$, 27, 8)
    Call REPLA(XX$, MKI(FECHANUM(FFAC$)), 19, 2)
      REFE$ = Mid$(TTXX$, 35, 30)
    Call REPLA(XX$, REFE$, 21, 48)
    Call REPLA(XX$, MKD$(IDEBE#), 69, 8)
    Call REPLA(XX$, MKD$(IHABE#), 77, 8)
    Call REPLA(XX$, "Pend. Conciliación", 93, 24)
    JJ& = JJ& + 1
    Call GRAREG("CONCIBAF", XX$, JJ&)
    '
30 Next J1&
  'OT-05-0498-WAR-FIN
  Call SETULTREG("CONCIBAF", JJ&)
  Call CIERRARCH("CONCIBAF")
  '
  Call HEADERS("CONCIBAF", "")
  Call HEADERS("CONCIBAF", "Corresponde a: " + CORRESP$)
  '
  Screen.MousePointer = 1
  Call CONECRUN("*LCONBAF", "Listado Movimientos Conciliados")
  '\\\OT-05-0498-FIN///
99 End Sub

Sub CARGA_TABLA_GASEXTRA()
   '
   If ULTREG&("CUEBANC") < 1 Then
      Call MENSERR(24, "Falta Definir Cuentas Bancarias")
      Call CARCUEBANC
      Exit Sub
   End If
   '
   Call TRALOCAL("GASEXTRA", "")
   J& = 0
   VTB% = VENTABU%("GASEXTRA")
   If VTB% >= 0 Then
     For KU& = 1 To ULTREG&("!GASEXTRA")
       XX$ = REGLEIDO$("!GASEXTRA", KU&)
       CUENTACON$ = Left$(XX$, 12)
       If CUEINT%(CUENTACON$) > 0 Then
         J& = J& + 1
         Call GRAREG("!GASEXTRA", XX$, J&)
       End If
     Next KU&
     Call CIERRARCH("!GASEXTRA")
     Call TRACENTRAL("GASEXTRA", "")
   End If

End Sub
Sub CARGA_TABLA_CFGCONCI()
   If ULTREG&("CUEBANC") < 1 Then
      Call MENSERR(24, "Falta Definir Cuentas Bancarias")
      Call CARCUEBANC
      Exit Sub
   End If
   '
   Call TRALOCAL("CFGCONCI", "")
   J& = 0
   VTB% = VENTABU%("CFGCONCI")
   If VTB% >= 0 Then
     For KU& = 1 To ULTREG&("!CFGCONCI")
       XX$ = REGLEIDO$("!CFGCONCI", KU&)
       CUENTACON$ = Left$(XX$, 12)
       If CUEINT%(CUENTACON$) > 0 Then
         J& = J& + 1
         Call GRAREG("!CFGCONCI", XX$, J&)
       End If
     Next KU&
     Call CIERRARCH("!CFGCONCI")
     Call TRACENTRAL("CFGCONCI", "")
   End If

End Sub

