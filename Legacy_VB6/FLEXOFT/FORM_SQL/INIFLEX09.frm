VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.ocx"
Begin VB.Form INIFLEX09 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Software de Gestión ERP - IMPLANSOFT (r)"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11910
   Icon            =   "INIFLEX09.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   11910
   Begin VB.PictureBox Picture999 
      Height          =   495
      Index           =   0
      Left            =   4200
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   31
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   255
      Index           =   0
      Left            =   3480
      OleObjectBlob   =   "INIFLEX09.frx":1D2A
      TabIndex        =   30
      Top             =   6240
      Visible         =   0   'False
      Width           =   495
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3480
      OleObjectBlob   =   "INIFLEX09.frx":1DA4
      Top             =   5520
   End
   Begin VB.Timer Timer4 
      Interval        =   60000
      Left            =   4320
      Top             =   3600
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   3465
      Top             =   3780
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4620
      Top             =   1785
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00624E28&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8600
      Left            =   11025
      ScaleHeight     =   8535
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   0
      Width           =   3075
      Begin VB.Frame Frame4 
         BackColor       =   &H00624E28&
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   0
         TabIndex        =   32
         Top             =   6240
         Width           =   855
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   90
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   36
            Top             =   960
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   37
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Miercoles"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   255
            Left            =   15
            TabIndex        =   35
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "00/00/00"
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
            Left            =   30
            TabIndex        =   34
            Top             =   600
            Width           =   750
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "00:00:00"
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
            Left            =   30
            TabIndex        =   33
            Top             =   1560
            Width           =   750
         End
      End
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "INIFLEX09.frx":1FD8
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Diseñador de Formularios"
         Top             =   4680
         Width           =   600
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Clear"
         Height          =   255
         Left            =   105
         TabIndex        =   28
         Top             =   3480
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00624E28&
         BorderStyle     =   0  'None
         Height          =   2115
         Left            =   0
         ScaleHeight     =   2115
         ScaleWidth      =   1275
         TabIndex        =   26
         Top             =   6195
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "INIFLEX09.frx":241A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Generador de Reportes - Consultas SQL"
         Top             =   4080
         Width           =   600
      End
      Begin VB.CommandButton Command16 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "INIFLEX09.frx":2724
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Cerrar Sesión - Cambiar Usuario / Entorno"
         Top             =   700
         Width           =   600
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Picture         =   "INIFLEX09.frx":286E
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2810
         Width           =   600
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         Picture         =   "INIFLEX09.frx":2CB0
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cambio de Clave Personal de Acceso"
         Top             =   2205
         Width           =   600
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "INIFLEX09.frx":2FBA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Manuales de Uso y Operacion"
         Top             =   1600
         Width           =   600
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "INIFLEX09.frx":32C4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Terminar - Cerrar Sesion de Trabajo"
         Top             =   105
         Width           =   600
      End
   End
   Begin VB.PictureBox Picture10 
      BackColor       =   &H80000005&
      Height          =   8500
      Left            =   -10
      ScaleHeight     =   8445
      ScaleWidth      =   3090
      TabIndex        =   17
      Top             =   -50
      Width           =   3150
      Begin ComctlLib.TreeView TreeView2 
         Height          =   4710
         Left            =   0
         TabIndex        =   25
         Top             =   5460
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   8308
         _Version        =   327682
         Indentation     =   441
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imlSmallIcons"
         BorderStyle     =   1
         Appearance      =   0
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
      Begin ComctlLib.TreeView Treeview1 
         Height          =   2400
         Left            =   -30
         TabIndex        =   18
         Top             =   1995
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   4233
         _Version        =   327682
         Indentation     =   441
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imlSmallIcons"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         Caption         =   "Frame1"
         Height          =   8415
         Left            =   3045
         MousePointer    =   9  'Size W E
         TabIndex        =   19
         Top             =   -80
         Width           =   75
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00404040&
         Height          =   75
         Left            =   -20
         TabIndex        =   40
         Top             =   850
         Width           =   15000
      End
      Begin VB.PictureBox Picture19 
         BackColor       =   &H00000000&
         Height          =   855
         Left            =   0
         ScaleHeight     =   795
         ScaleWidth      =   3000
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   3060
         Begin VB.Image Image11 
            Height          =   1590
            Left            =   0
            Top             =   0
            Width           =   2010
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Frame1"
         ForeColor       =   &H00404040&
         Height          =   75
         Left            =   -1800
         TabIndex        =   20
         Top             =   4920
         Width           =   15000
      End
      Begin ACTIVESKINLibCtl.SkinLabel Label7 
         Height          =   255
         Left            =   0
         OleObjectBlob   =   "INIFLEX09.frx":340E
         TabIndex        =   38
         Top             =   5040
         Width           =   3615
      End
      Begin ACTIVESKINLibCtl.SkinLabel Label6 
         Height          =   255
         Left            =   0
         OleObjectBlob   =   "INIFLEX09.frx":3476
         TabIndex        =   39
         Top             =   900
         Width           =   3615
      End
      Begin ComctlLib.ImageList imlIcons 
         Left            =   600
         Top             =   4185
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   1
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "INIFLEX09.frx":34E8
               Key             =   "book"
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0080C0FF&
      Height          =   2000
      Left            =   3150
      ScaleHeight     =   1935
      ScaleWidth      =   15945
      TabIndex        =   24
      Top             =   7245
      Width           =   16000
      Begin VB.Image Image3 
         Height          =   480
         Index           =   0
         Left            =   210
         Picture         =   "INIFLEX09.frx":3802
         Top             =   210
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00C0C0C0&
      Height          =   1300
      Left            =   8925
      Picture         =   "INIFLEX09.frx":DB62
      ScaleHeight     =   1245
      ScaleWidth      =   1770
      TabIndex        =   1
      Top             =   105
      Visible         =   0   'False
      Width           =   1830
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   165
         TabIndex        =   7
         Top             =   420
         Width           =   225
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   780
         TabIndex        =   6
         Top             =   900
         Width           =   225
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1460
         TabIndex        =   5
         Top             =   470
         Width           =   225
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   800
         TabIndex        =   4
         Top             =   60
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E1E3&
      Height          =   1380
      Left            =   3000
      ScaleHeight     =   1320
      ScaleWidth      =   2370
      TabIndex        =   21
      ToolTipText     =   "Click con Boton Derecho para Cambiar la Imagen"
      Top             =   360
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.CommandButton Command17 
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
      Left            =   2520
      Picture         =   "INIFLEX09.frx":152F4
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Depósito Caotico SAPORITI"
      Top             =   6195
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   2625
      Pattern         =   "*.EXK"
      TabIndex        =   14
      Top             =   210
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   420
      Top             =   0
   End
   Begin VB.PictureBox Picture9 
      Height          =   2010
      Left            =   420
      ScaleHeight     =   1950
      ScaleWidth      =   2055
      TabIndex        =   12
      Top             =   210
      Visible         =   0   'False
      Width           =   2115
      Begin VB.Image Image1 
         Height          =   1590
         Left            =   0
         Top             =   0
         Width           =   2010
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5500
      Top             =   0
   End
   Begin VB.CommandButton Command12 
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
      Left            =   1785
      Picture         =   "INIFLEX09.frx":1543E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Control y Movimientos de Stocks"
      Top             =   5040
      Width           =   600
   End
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
      Height          =   600
      Left            =   630
      Picture         =   "INIFLEX09.frx":15748
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Actualización y Consulta de Archivos Maestros"
      Top             =   3360
      Width           =   600
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frame1"
      ForeColor       =   &H00404040&
      Height          =   15
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   23
      Top             =   7140
      Width           =   15000
   End
   Begin VB.Label Label2 
      Caption         =   "http://www.smn.gov.ar/"
      Height          =   330
      Left            =   3360
      TabIndex        =   22
      Top             =   2625
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.Image Image2 
      Height          =   19425
      Left            =   6615
      Picture         =   "INIFLEX09.frx":15A52
      Top             =   200
      Visible         =   0   'False
      Width           =   35790
   End
   Begin ComctlLib.ImageList imlSmallIcons 
      Left            =   5000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "INIFLEX09.frx":42CFF
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "INIFLEX09.frx":43221
            Key             =   "cylinder"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "INIFLEX09.frx":43743
            Key             =   "open1"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "INIFLEX09.frx":45295
            Key             =   "leaf"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "INIFLEX09.frx":457B7
            Key             =   "open"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "INIFLEX09.frx":45CD9
            Key             =   "smlBook"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "INIFLEX09.frx":4633B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "INIFLEX09.frx":4699D
            Key             =   "closed1"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "INIFLEX09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARRABOR%

'
Private Sub Command1_Click()
   '
   Call GRAVALCON(USNBR%, "POSIFORPRIN", "TOP", FORMATNUM$(Top, "F10.0"))
   Call GRAVALCON(USNBR%, "POSIFORPRIN", "LEFT", FORMATNUM$(Left, "F10.0"))
   If WindowState < 1 Then Call GRAVALCON(USNBR%, "POSIFORPRIN", "HEIGHT", FORMATNUM$(Height, "F10.0"))
   If WindowState < 1 Then Call GRAVALCON(USNBR%, "POSIFORPRIN", "WIDTH", FORMATNUM$(Width, "F10.0"))
   Call GRAVALCON(USNBR%, "POSIFORPRIN", "WINDOWSTATE", FORMATNUM$(WindowState, "F10.0"))
   Call GRAVALCON(USNBR%, "POSIFORPRIN", "FRAME1_LEFT", FORMATNUM$(Frame1.Left, "F10.0"))
   Call GRAVALCON(USNBR%, "POSIFORPRIN", "FRAME2_TOP", FORMATNUM$(Frame2.Top, "F10.0"))
   '
   Call COMPACTABD
   '
   Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   If PROCACTI% < 1 Then
      Call GENBACKUP
      Call CANCELUCOM
      Call FINAL("")
   End If
   '
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   If EXISTE%("IFABRI07.EXE") > 0 Then
     Call CONECRUN("IFABRI07", "Informes de Fabricación")
   Else
     Call COMUNI("Falta módulo IFABRI07")
   End If
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   If EXISTE%("ARCHIG07.EXE") > 0 Then
     Call CONECRUN("ARCHIG07", "Archivos Maestros")
   Else
     Call COMUNI("Falta módulo ARCHIG07")
   End If
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   If EXISTE%("CONSTO07.EXE") > 0 Then
     Call CONECRUN("CONSTO07", "Control de Stocks")
   Else
     Call COMUNI("Falta módulo CONSTO07")
   End If
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   If EXISTE%("PROMRP07.EXE") > 0 Then
     Call CONECRUN("PROMRP07", "Programación de Necesidades MRP")
   Else
     Call COMUNI("Falta módulo PROMRP07")
   End If
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call CONECRUN("SOLIMA07", "Solicitud Interna de Materiales")
   Command14.Enabled = True
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
'   If EXISTE%("ORFABR04.EXE") > 0 Then
'     Call CONECRUN("ORFABR04", "Sistema de Ordenes de Fabricación")
'   Else
'     Call COMUNI("Falta módulo ORFABR04")
'   End If
  '\\\OT-05-0415-WAR-22/06/05///
   If EXISTE%("ORCONF07.EXE") > 0 Then
       Call CONECRUN("ORCONF07", "Sistema de Ordenes de Fabricación y Envios de Prendas")
     ElseIf EXISTE%("ORFABR07.EXE") > 0 Then
       Call CONECRUN("ORFABR07", "Sistema de Ordenes de Fabricación")
     ElseIf EXISTE%("MINORF07.EXE") > 0 Then
       Call CONECRUN("MINORF07", "Sistema Basico de Fabricación")
     Else
       Call COMUNI("Falta módulo ORFABR07")
   End If
   '\\\OT-05-0415-WAR-FIN///
   Command15.Enabled = True
End Sub

Private Sub Command16_Click()
   Command16.Enabled = False
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   If PROCACTI% < 1 Then
      Call CANCELUCOM
      ' verifica si es multiempresa
      Call RECONN_EMPRE(OKXX%)
      If OKXX% < 1 Then GoTo 80
      TCONFI$ = LOADFILE$(CurDir + "\" + "FLEXCFG.DAT")
      If TCONFI$ = "" Then Exit Sub
      '
      LUDAT$ = TRIM$(Mid$(TCONFI$, 51, 45))
      If Right$(LUDAT$, 1) <> "\" Then
         LUDAT$ = LUDAT$ + "\"
         Mid$(TCONFI$, 51, 45) = Left$(LUDAT$ + Space$(45), 45)
         Call SAVEFILE("FLEXCFG.DAT", TCONFI$)
      End If
      If Mid$(LUDAT$, 2, 1) <> ":" Then
         ULO$ = Left$(TRIM$(LOADFILE$(LUCOM$ + "ULOGDAT.DRV")), 2)
         If ULO$ = "" Then ULO$ = Left$(CurDir$, 2)
         LUDAT$ = ULO$ + LUDAT$
      End If
      '
      RSTRI$ = LOADFILE$(LUCOM$ + "FLEXRAM.DRV") + String$(512, 0)
      EMPREAC$ = TRIM$(Mid$(RSTRI$, 193, 30))
         Call GRARAM(99, AJUSTI$(LUDAT$, 46))
         Call GRARAM(193, AJUSTI$(EMPREAC$, 30))
         Call GRARAM(191, Chr$(CodiEmp%))
         Call BAJRAM
      EMPREAC$ = EMPREAC$ + " (" + TRIM$(Str$(CodiEmp%)) + ") "
      EMPREAC$ = EMPREAC$ + VERNUM$ + " - " + USERNAME$
      Screen.MousePointer = 1
      If OKEY% < 1 Then Call FINAL("")
      EPAC$ = TRIM$(EMPREAC$)
      If EPAC$ <> "" Then
         Caption = "Sistema Preconfigurado FLEXOFT" + " - " + EPAC$
         Me.Refresh
      End If
      Call GRATITFOR(Caption)
      Call PONELOGO
      Call UPDATES
GoTo 99

80    ' conecta con FLOGON.EXE
      On Error Resume Next
      RETOR$ = TRIM$(UCase$(App.EXEName))
      Shell "FLOGON.EXE " + RETOR$, vbNormalFocus
      If Err Then
        Call MERRFINAL("Falta Utilitario 'FLOGON.EXE'.")
      End If
      Call FINAL("")
   End If
   '
99 Command16.Enabled = True
End Sub

Sub UPDATES()
   '
   Screen.MousePointer = 11
   Label2 = 171: DoEvents
'   Call LEEDATSYS
   '
   YAMENSA% = 0
   INSTALAC% = 0
   '
   DIREPRO$ = App.Path
   If EXISTE%("F:\FLEXOFT\NOVEDAD\.") > 0 Then
       DIRENEW$ = "F:\FLEXOFT\NOVEDAD"
     ElseIf EXISTE%("C:\FLEXOFT\NOVEDAD\.") > 0 Then
       DIRENEW$ = "C:\FLEXOFT\NOVEDAD"
     Else
       DIRENEW$ = PARENTFOLDER$(DIREPRO$) + "\NOVEDAD"
   End If
   DIRNU$ = DIRENEW$
10 On Error Resume Next
   File1.Path = DIRNU$
   If Err < 1 Then
     On Error GoTo 0
     File1.Pattern = "*.*"
     '
     For KU& = 1 To File1.ListCount
       ANAME$ = TRIM$(File1.List(KU& - 1))
       PPXX% = InStr(ANAME$, ".")
       If PPXX% < 1 Then PPXX% = 1 + Len(ANAME$)
       exten$ = TRIM$(UCase$(Mid$(ANAME$, PPXX% + 1)))
       '
       If Len(exten$) = 3 Then
         '
         If exten$ = "RFS" Or exten$ = "FRM" Then
           FECHANUE# = 0
           On Error Resume Next
           FECHANUE# = FileDateTime(DIRENEW$ + "\" + ANAME$)
           On Error GoTo 0
           FECHAVIE# = 0
           If EXISTE%(LUDAT$ + ANAME$) > 0 Then
             FECHAVIE# = FileDateTime(LUDAT$ + ANAME$)
           End If
           '
           If FECHANUE# > FECHAVIE# + 0.0001 Then  ' ESO SON APROX 6 SEGUNDOS
             On Error Resume Next
             FileCopy DIRENEW$ + "\" + ANAME$, LUDAT$ + ANAME$
             If Err Then
               MsgBox "Imposible Instalar (3) " + ANAME$
             End If
             On Error GoTo 0
           End If
           GoTo 29         ' SALE DE INSTALACION DE DATOS
         End If
         '
         RUTACOM$ = App.Path + "\"
         If EXISTE%(RUTACOM$ + ANAME$) > 0 Or exten$ = "PRF" Then
           On Error Resume Next
           FECHANUE# = FileDateTime(DIRENEW$ + "\" + ANAME$)
           If Err Then
              On Error GoTo 0
              GoTo 29
           End If
           On Error GoTo 0
           FECHAVIE# = 0
           If EXISTE%(RUTACOM$ + ANAME$) > 0 Then
             FECHAVIE# = FileDateTime(RUTACOM$ + ANAME$)
           End If
           If FECHANUE# > FECHAVIE# + 0.0001 Then  ' ESO SON APROX 6 SEGUNDOS
             If YAMENSA% = 0 Then
               MESSA$ = "Hay Actualizaciones Listas para Instalarse, des-" + Chr$(10) + Chr$(13)
               MESSA$ = MESSA$ + "de la Carpeta " + DIRENEW$ + "." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
               MESSA$ = MESSA$ + "Para Instalar las Actualizaciones Disponibles" + Chr$(10) + Chr$(13)
               MESSA$ = MESSA$ + "Cierre las Aplicaciones FLEXOFT en TODAS las" + Chr$(10) + Chr$(13)
               MESSA$ = MESSA$ + "Estaciones de Trabajo." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
               MESSA$ = MESSA$ + "¿ Desea Ahora Actualizar su Sistema FLEXOFT ?"
               RESPU% = MsgBox(MESSA$, vbYesNoCancel)
               If RESPU% <> 6 Then
                 Screen.MousePointer = 1
                 Exit Sub     ' 6 es = SI
               End If
               YAMENSA% = 1
             End If
             '
             EXTECO$ = Left$(EXTE$, 2)
             ANAMCO$ = Left$(ANAME$, Len(ANAME$) - 1)
             On Error Resume Next
             Kill RUTACOM$ + ANAMCO$ + "0"
             Kill DIREVIE$ + ANAMCO$ + "9"
             On Error GoTo 0
             If EXISTE%(DIREVIE$ + ANAMCO$ + "9") > 0 Then
               'MsgBox "Imposible Instalar (1) " + ANAME$
               MsgBox "Imposible Instalar. No Pudieron Borrarse " + RUTACOM$ + ANAMCO$ + "0 y " + DIREVIE$ + ANAMCO$ + "9"
               On Error GoTo 0
               GoTo 29
             End If
             '
             On Error Resume Next
             For KI% = 8 To 0 Step -1
               Name DIREVIE$ + ANAMCO$ + TRIM$(Str$(KI%)) As DIREVIE$ + ANAMCO$ + TRIM$(Str$(KI% + 1))
             Next KI%
             On Error GoTo 0
             If EXISTE%(DIREVIE$ + ANAMCO$ + "0") > 0 Then
               'MsgBox "Imposible Instalar (2) " + ANAME$
               MsgBox "Imposible Instalar. No Pudo Renombrarse " + DIREVIE$ + ANAMCO$ + "0 como " + DIREVIE$ + ANAMCO$ + "1"
               On Error GoTo 0
               GoTo 29
             End If
             '
             If InStr(ANAME$, "MSTDRV04") < 1 Then  ' PARA QUE NO RESPALDE RECLAMO
               If EXISTE%(RUTACOM$ + ANAME$) < 1 Then GoTo 24
               On Error Resume Next
               Name RUTACOM$ + ANAME$ As RUTACOM$ + ANAMCO$ + "0"
               If Err Then
                 ' MsgBox "Imposible Instalar (4) " + ANAME$
                 MsgBox "Imposible Instalar. No Pudo Renombrarse " + RUTACOM$ + ANAME$ + " como " + RUTACOM$ + ANAMCO$ + "0"
                 On Error GoTo 0
                 GoTo 29
               End If
               On Error Resume Next
               Name RUTACOM$ + ANAMCO$ + "0" As DIREVIE$ + ANAMCO$ + "0"
               If Err Then
                 ' MsgBox "Imposible Instalar (5) " + ANAME$
                 MsgBox "Imposible Instalar. No Pudo Renombrarse " + RUTACOM$ + ANAMCO$ + "0 como " + DIREVIE$ + ANAMCO$ + "0"
                 On Error GoTo 0
                 GoTo 29
               End If
             End If
             '
24           On Error Resume Next
             FileCopy DIRENEW$ + "\" + ANAME$, RUTACOM$ + ANAME$
             If Err Then
               MsgBox "Imposible Instalar. No Pudo Copiarse " + DIRENEW$ + "\" + ANAME$ + " como " + RUTACOM$ + ANAME$
               On Error GoTo 0
               GoTo 29
             End If
             INSTALAC% = 1
             '
           End If
         End If
       End If
29   Next KU&
     '
   End If
   '
   ' esto es para explorar tambien el \FLEXOFT\NOVEDAD de la carpeta de trabajo
   On Error GoTo 0
   If DIRNU$ = "F:\FLEXOFT\NOVEDAD" Then
     DIRNU$ = PARENTFOLDER$(DIREPRO$) + "\NOVEDAD"
     If DIRNU$ <> "F:\FLEXOFT\NOVEDAD" Then
       GoTo 10
     End If
   End If
   '
   If INSTALAC% > 0 Then MsgBox "Actualizaciones Instaladas"
   Screen.MousePointer = 1
   '
End Sub

Sub RECONN_EMPRE(OKXX%)
   '
   OKXX% = 0
   If EXISTE%("FLEXMUL.DRV") = 1 Then
      SELUDAT2.List2.Clear
      SELUDAT2.List2.Visible = True
      '
      NX% = FILEOPEN("FLEXMUL.DRV", 0, 79)
      Dim FMX As String * 79
      '
      For i% = 1 To LOF(NX%) / 79
         Get #NX%, i%, FMX$
         SUCAR% = 0
         For J% = 1 To 77
            SUCAR% = SUCAR% + Asc(Mid$(FMX$, J%, 1))
         Next J%
         If CVI(Mid$(FMX$, 78, 2)) = SUCAR% Then
            RASO$ = DESENCRI$(Left$(FMX$, 30))
            PACCE$ = DESENCRI$(Mid$(FMX$, 32, 46))
            If EXISTE%(TRIM$(PACCE$) + "\.") > 0 Then
               Z$ = Space$(96)
               Call REPLA(Z$, RASO$, 1, 30)
               Call REPLA(Z$, PACCE$, 51, 46)
               SELUDAT2.List2.AddItem Z$
            End If
         End If
      Next i%
      Close #NX%
      '
      If SELUDAT2.List2.ListCount > 0 Then
         OKXX% = 1
         SELUDAT2.List2.ListIndex = 0
         Screen.MousePointer = 1
         If SELUDAT2.List2.ListCount = 1 Then
            Call SELUDAT2.List2_DblClick
            GoTo 25
         End If
         SELUDAT2.Show 1
         Screen.MousePointer = 11
         '
         If SELUDAT2.List2.ListCount > 1 Then
           If TRIM$(EMPREAC$) <> TRIM$(Left$(SELUDAT2.List2.List(0), 30)) Then
             If DERACCE%(EMPREAC$, "Acceso a " + EMPREAC$) < 2 Then
               Call FINAL("")                 ' NO ACCEDE
             End If
             Call validaMultiEmpXUsuario
           End If
         End If
         '
      End If
   End If
   '
25
End Sub

Private Sub Command17_Click()
   Command17.Visible = False
   ACONEC$ = "TRANSA07"
   Call CONECRUN(ACONEC$, "Deposito Caotico SAPORITI")
   Command17.Visible = True
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   If EXISTE%("ORCOMP07.EXE") > 0 Then
     Call CONECRUN("ORCOMP07", "Sistema de Ordenes de Compra")
   Else
     Call COMUNI("Falta módulo ORCOMP07")
   End If
   Command18.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call MENSERR(24, "Ayuda en Linea no Disponible\en Esta Localización.")
   Command2.Enabled = True
End Sub

Private Sub Command23_Click()
    Command23.Enabled = False
    MODU$ = "REMITO07"
    Call CONECRUN(MODU$, "Emision de Remito de Facturacion")
    Command23.Enabled = True
End Sub

Private Sub Command28_Click()
   Command28.Visible = False
   ACONEC$ = "BRECEP07"
   Call CONECRUN(ACONEC$, "Recepcion de Materiales")
   Command28.Visible = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call CONECRUN("FSETUP04/CLAVUSER", "Cambio de Clave Personal de Acceso")
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   If DERACCE%("FSETUP07/", "CONFIGURACION GENERAL") < 2 Then GoTo 99
   FSETUPNEW.Show 1
   Call FINAL("")
99 Command4.Enabled = True
'   Call CONECRUN("FSETUP04", "Configuración de Funcionamiento")
'   Call FINAL("")
'   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
     NOMECOR$ = "SQLCTRL"
     On Error Resume Next
     BaDaConsul.Close
     On Error GoTo 0
     On Error Resume Next
     Kill App.Path + "\" + NOMECOR$ + "1.MDB"
     On Error GoTo 0
60   On Error Resume Next
     CompactDatabase App.Path + "\" + NOMECOR$ + ".MDB", App.Path + "\" + NOMECOR$ + "1.MDB"
     If Err Then GoTo 75
     FileCopy App.Path + "\" + NOMECOR$ + ".MDB", App.Path + "\" + NOMECOR$ + ".OLD"
     If Err Then GoTo 75
     FileCopy App.Path + "\" + NOMECOR$ + "1.MDB", App.Path + "\" + NOMECOR$ + ".MDB"
     If Err Then
75     TTXX$ = "Error (" & Err.Number & ") al Compactar " + NOMECOR$
       If Err.Number = 3343 Then
          On Error GoTo 0
          TTXX$ = TTXX$ + "\  \Esta Base de Datos está Dañada."
          TTXX$ = TTXX$ + "\¿ Desea Que FLEXOFT Intente Repararla ?"
          TTXX$ = TTXX$ + "\\Si, Reparar\No, Cancelar"
          If COMALTER%(TTXX$) = 1 Then
            On Error Resume Next
            DBEngine.RepairDatabase LUACC$ + NOMECOR$ + ".MDB"
            GoTo 60
          End If
          GoTo 80
       End If
       MsgBox TTXX$
     End If
80   On Error GoTo 0
     On Error Resume Next
     Kill App.Path + "\" + NOMECOR$ + "1.MDB"
     On Error GoTo 0
     MsgBox "FIN"
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call CONECRUN("GENREP07", "Generador de Reportes SQL")
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   Call CONECRUN("DSGFRM09", "Diseñador de Formularios")
   Command7.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   If EXISTE%("CALCAP07.EXE") > 0 Then
     Call CONECRUN("CALCAP07", "Cálculo de Capacidad")
   Else
     Call COMUNI("Falta módulo CALCAP07")
   End If
   Command9.Enabled = True
End Sub

Private Sub Form_Load()

   
     
   Screen.MousePointer = 11
   Visible = False
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   Screen.MousePointer = 11
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
     Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   Screen.MousePointer = 11
   Call ABRECONEXION
   Call CREATABLAS_SQLCTRL
   Picture4.BackColor = Picture7.BackColor
   '
   ' LEE LOS DATOS GRABADOS DE LA POSICION ANTERIOR
   Screen.MousePointer = 11
   TOPX$ = VALCONTROL$(USNBR%, "POSIFORPRIN", "TOP")
   Screen.MousePointer = 11
   LEFTX$ = VALCONTROL$(USNBR%, "POSIFORPRIN", "LEFT")
   HEIX$ = VALCONTROL$(USNBR%, "POSIFORPRIN", "HEIGHT")
   WIDX$ = VALCONTROL$(USNBR%, "POSIFORPRIN", "WIDTH")
   WINSTX$ = VALCONTROL$(USNBR%, "POSIFORPRIN", "WINDOWSTATE")
   FR1LFT$ = VALCONTROL$(USNBR%, "POSIFORPRIN", "FRAME1_LEFT")
   FR2TOP$ = VALCONTROL$(USNBR%, "POSIFORPRIN", "FRAME2_TOP")
   FR3TOP$ = VALCONTROL$(USNBR%, "POSIFORPRIN", "FRAME3_TOP")
   '
   ' LEE LOS DATOS GRABADOS DE LA IMAGEN PERSONAL
   Screen.MousePointer = 11
   ARIMAGX$ = VALCONTROL$(USNBR%, "IMAGEN", "ARCHIMAGEN")
   '
   'APLICA LOS DATOS DE LA POSICION Y LA IMAGEN
   Screen.MousePointer = 11
   On Error Resume Next
     If WINSTX$ <> "" Then WindowState = XVALO(WINSTX$)
     If TOPX$ <> "" Then Top = XVALO(TOPX$)
     If LEFTX$ <> "" Then Left = XVALO(LEFTX$)
     If WindowState < 1 Then
         If HEIX$ <> "" Then Height = XVALO(HEIX$)
         If WIDX$ <> "" Then Width = XVALO(WIDX$)
     End If
     If FR1LFT$ <> "" Then Frame1.Left = XVALO(FR1LFT$)
     If FR2TOP$ <> "" Then Frame2.Top = XVALO(FR2TOP$)
        If Frame2.Top > Height - 1000 Then Frame2.Top = Height - 1000
     If FR3TOP$ <> "" Then DINFE& = XVALO(FR3TOP$)
     If ARIMAGX$ <> "" Then Image2.Picture = LoadPicture(ARIMAGX$)
   On Error GoTo 0
   '
   If DINFE& < 750 Or DINFE& > Height Then DINFE& = 1500
   Frame3.Top = Height - DINFE&
   '
   On Error Resume Next
   Picture3.Top = Frame3.Top + Frame3.Height
   Picture3.Height = Height - Picture3.Top
   On Error GoTo 0
   Screen.MousePointer = 11
   Call ACOMODATODO
   '
   ' PONE LOS ACCESOS DIRECTOS
   Screen.MousePointer = 11
   Call CARGACCDIR
   '
   Treeview1.Style = tvwTreelinesPlusMinusPictureText
   On Error Resume Next
   imlSmallIcons.ListImages.Add , "nueva", LoadPicture("C:\WINDOWS\ICONS\BOOKS05.ICO")
   On Error GoTo 0
   Set nodX = Treeview1.Nodes.Add(, , "ROOT", "Sistema Preconfigurado", "closed1", "closed1")
   Set nodX = Treeview1.Nodes.Add("ROOT", tvwChild, "FLGEST", "Gestión Administrativa", "closed1", "closed1")
   Set nodX = Treeview1.Nodes.Add("ROOT", tvwChild, "FLIND", "Gestion Industrial", "closed1", "closed1")
   '
   Screen.MousePointer = 11
   Call CARGAMENU
   '
   Visible = True
'   UTILIZA_SKIN% = 1
'   Call APLICACIONSKINS(Me)
   Frame5.Top = Label6.Top - Frame5.Height
   Screen.MousePointer = 1
   '
End Sub

Sub CARGAMENU()
    '
    Dim CLACLA$(4096), Veces%(4096)
    CANCLA% = 0
    '
    Treeview1.Nodes.Clear
    Set nodX = Treeview1.Nodes.Add(, , "ROOT", "Sistema Preconfigurado", "closed1", "closed1")
    USNU% = USNBR% Mod 100
    '
    U& = 1
    Do While U& <= Treeview1.Nodes.Count
       CLAPA$ = Treeview1.Nodes(U&).key
10     SQLX$ = "SELECT * FROM USER_MENUS WITH(NOLOCK) "
       SQLX$ = SQLX$ + " WHERE USUARIO = 0 " ' & USNU%
       SQLX$ = SQLX$ + " AND CodigoPadre = '" & CLAPA$ & "' "
       SQLX$ = SQLX$ + " ORDER BY NumeOrden ASC"
       'Set MENUTEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenSnapshot)
       Set MENUTEMP = READSET(SQLX$)
       '
       With MENUTEMP
         If Not (.BOF And .EOF) Then
            Do While Not .EOF
              CODPA$ = SAFETEXT$(!CodigoPadre)
              CODIJ$ = SAFETEXT$(!CodigoHijo)
              DESIJ$ = SAFETEXT$(!Descripcion)
              IMAGX$ = SAFETEXT$(!CodigoImagen)
              If CODPA$ = CODIJ$ Then GoTo 19
              '
              For KKII% = 1 To CANCLA%
                 If CLACLA$(KKII%) = CODIJ$ Then
                    Veces%(KKII%) = Veces%(KKII%) + 1
                    CODIJ$ = AJUSTI$(CODIJ$, 64) + "/" + TRIM$(Str$(Veces%(KKII%)))
                    GoTo 17
                 End If
              Next KKII%
              CANCLA% = CANCLA% + 1
              CLACLA$(CANCLA%) = CODIJ$
              '
17            On Error Resume Next
              Set nodX = Treeview1.Nodes.Add(CODPA$, tvwChild, CODIJ$, DESIJ$, "closed1", "closed1")
              If TRIM$(CODPA$) = "" Or TRIM$(CODPA$) = "ROOT" Then
                 nodX.EnsureVisible ' Muestra todos los nodos.
              End If
              On Error GoTo 0
19          .MoveNext
            Loop
         Else
            If USNU% > 0 Then
               MENUTEMP.Close
               USNU% = 0
               GoTo 10      ' REPITE PARA EL MENU GENERAL
            End If
         End If
       End With
       MENUTEMP.Close
       Set MENUTEMP = Nothing
       U& = U& + 1
    Loop
    '
    Call CARGAPREFERIDOS
    '
End Sub

Sub CARGAPREFERIDOS()
    '
    TreeView2.Nodes.Clear
    USNU% = USNBR% Mod 100
    
    '
20  SQLX$ = "SELECT TOP 10 * FROM USER_MENUS WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE USUARIO = " & USNU%
    SQLX$ = SQLX$ + " AND CantiAcce > 0"
    SQLX$ = SQLX$ + " ORDER BY CantiAcce DESC"
    'Set MENUTEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenDynaset)
    Set MENUTEMP = READSET(SQLX$)
    '
    With MENUTEMP
      If Not (.BOF And .EOF) Then
            Do While Not .EOF
              CODIJ$ = SAFETEXT$(!CodigoHijo)
              DESIJ$ = TRIM$(SAFETEXT$(!Descripcion)) + " (" + TRIM$(Str$(XVALO(!CantiAcce))) + ")"
              IMAGX$ = SAFETEXT$(!CodigoImagen)
              On Error Resume Next
              Set nodY = TreeView2.Nodes.Add(, , CODIJ$, DESIJ$, "closed1", "closed1")
              nodY.EnsureVisible ' Muestra todos los nodos.
              On Error GoTo 0
            .MoveNext
            Loop
         Else
            If USNU% > 0 Then
              MENUTEMP.Close
              USNU% = 0
              GoTo 20      ' REPITE PARA EL MENU GENERAL
            End If
      End If
    End With
    MENUTEMP.Close
    '
End Sub
'
Sub CUENTACCESOS(CLAVIJO$)
    '
    USNU% = 0
    CANTACC = 0
    '
    'LA  PRIMERA VEZ TOMA LOS DATOS DEL REGISTRO CON EL USUARIO 0 QUE ES CON EL QUE SE CREA PARA MANTENER EL CRITERIO ANTERIOR
    'SI NO EXISTE LO CREA SI NO LO EDITA
10  SQLX$ = "SELECT * FROM USER_MENUS "
    SQLX$ = SQLX$ + " WHERE USUARIO = " & USNU%
    SQLX$ = SQLX$ + " AND CodigoHijo = '" & CLAVIJO$ & "'"
    'Set MENUTEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenDynaset)
    Dim MENUTEMP As New ADODB.Recordset
    MENUTEMP.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText    '
    With MENUTEMP
      If Not (.BOF And .EOF) Then
        USU1% = XVALO(!USUARIO)
        COPADRE$ = SAFETEXT$(!CodigoPadre)
        COHIJO$ = SAFETEXT$(!CodigoHijo)
        DESCRI$ = SAFETEXT$(!Descripcion)
        COIMAGE$ = SAFETEXT$(!CodigoImagen)
        NUOR% = XVALO(!NumeOrden)
      Else
        .AddNew
      End If
         !CantiAcce = XVALO(!CantiAcce) + CANTACC ' LA PRIMERA VEZ CARGA EN 0
         !USUARIO = USNU%
         !CodigoPadre = COPADRE$
         !CodigoHijo = COHIJO$
         !Descripcion = DESCRI$
         !CodigoImagen = COIMAGE$
         !NumeOrden = NUOR%
          .Update
    End With
    MENUTEMP.Close
    Set MENUTEMP = Nothing
    '
    If USNU% > 0 Then GoTo 50 'SE VA
    
    USNU% = USNBR% Mod 100
    CANTACC = 1
    GoTo 10
50  'MENUTEMP.Close
    '
    Call CARGAPREFERIDOS
    '
End Sub
'
Sub CARGACCDIR()
   CANADIR = XVALO(VALCONTROL(USNBR%, "ADIRECT", "CANADIR"))
   If CANADIR > 255 Then CANADIR = 255
   '
   BULTIMA = 0    ' Image3(0).Left + Image3(0).Width
   For i& = 1 To CANADIR
      If i& > Image3.Count Then
         Load Image3(i& - 1)
      End If
      '
      ADIR$ = "ACCDIR-" + TRIM$(Str$(i& - 1))
      ARCHICON$ = VALCONTROL(USNBR%, ADIR$, "ARCHICON")
      TOOTIPTE$ = VALCONTROL(USNBR%, ADIR$, "TOOTIPTE")
      Image3(i& - 1).ToolTipText = TOOTIPTE$
      Image3(i& - 1).Left = 0
      '
      If ARCHICON$ <> "" Then
        On Error Resume Next
        Image3(i& - 1).Picture = LoadPicture(ARCHICON$)
        If Err Then
             On Error GoTo 0
           Else
             On Error GoTo 0
             Image3(i& - 1).Top = Image3(0).Top + (Image3(0).Height - Image3(i& - 1).Height) / 2
             Image3(i& - 1).Left = BULTIMA + 150
             BULTIMA = Image3(i& - 1).Left + Image3(i& - 1).Width
             Image3(i& - 1).Visible = True
             Image3(i& - 1).Refresh
        End If
      End If
      '
   Next i&
   '
   For i& = 1 To CANADIR
     If Image3(i& - 1).Left = 0 Then
       ADIR$ = "ACCDIR-" + TRIM$(Str$(i& - 1))
       ARCHICON$ = VALCONTROL(USNBR%, ADIR$, "ARCHICON")
       If ARCHICON$ <> "" Then
         Image3(i& - 1).Top = Image3(0).Top + (Image3(0).Height - Image3(i& - 1).Height) / 2
         Image3(i& - 1).Left = BULTIMA + 150
         BULTIMA = Image3(i& - 1).Left + Image3(i& - 1).Width
         Image3(i& - 1).Visible = True
         Image3(i& - 1).Refresh
       End If
     End If
   Next i&
   '
End Sub

Sub PONELOGO()
   '
   RUTARGRA$ = TRIM$(Control$("", "RUTARGRA"))
   If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
   If Right$(RUTARGRA$, 1) <> "\" Then
      RUTARGRA$ = RUTARGRA$ + "\"
   End If
   '
   ISOLOGO$ = TRIM$(Control$("", "ISOLOGO"))
   If ISOLOGO$ <> "" Then
      ARCHIMAG$ = RUTARGRA$ + ISOLOGO$
      If EXISTE%(ARCHIMAG$) = 1 Then
        Image11.Picture = LoadPicture("")
        On Error Resume Next
        Image11.Picture = LoadPicture(ARCHIMAG$)
        If Err Then
          On Error GoTo 0
          ARXX$ = REPLACAR$(ARCHIMAG$, "\", "/")
          Call MENSERR(24, "Archivo Imagen Isologo\" + ARXX$ + "\No Válido o Inexistente.")
        End If
        On Error GoTo 0
        Picture19.Visible = True
        IMAGEFILE$ = ARCHIMAG$     ' LUDAT$ + "LOGOEMPR.BMP"
        '
        FORMALTER.Image3.Picture = LoadPicture("")
        On Error Resume Next
        FORMALTER.Image3.Picture = LoadPicture(IMAGEFILE$)
        On Error GoTo 0
        '
12      HIMAGEN = FORMALTER.Image3.Picture.Width
        VIMAGEN = FORMALTER.Image3.Picture.Height
        '
        'Picture19.Width = Frame1.Left
        'Picture19.Height = 1000
        'Picture19.Left = 0
        HCUADRO = Frame1.Left   'Picture19.Width
        VCUADRO = 1000 ' Picture19.Height
        '
        REDUH = HIMAGEN / HCUADRO
        REDUV = VIMAGEN / VCUADRO
        REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
        If REDUC < 1 Then REDUC = 1
        '
        With Picture19
          .Height = VCUADRO * REDUV / REDUC
          .Width = HCUADRO * REDUH / REDUC
          '.Left = .Left + (HCUADRO - Picture19.Width) / 2
          .Left = (HCUADRO - Picture19.Width) / 2
            If .Left < 0 Then .Left = 0
          .Top = .Top - (VCUADRO - Picture19.Height) / 2
            If .Top < 0 Then .Top = 0
        End With
        '
        Image11.Width = Picture19.Width
        Image11.Height = Picture19.Height
        Image11.Stretch = True
        On Error Resume Next
        Image11.Picture = LoadPicture(IMAGEFILE) ' FORMALTER.IMAGE11.Picture ' , PH1 + BORIZQ, PV1 + BORSUP, ANCHCUA, ALTUCUA
        On Error GoTo 0
        '
        FORMALTER.Image3.Picture = LoadPicture("")
        '
      End If
   End If
   '
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   '
   Cancel = False
   If UnloadMode = 0 Then     ' desde el botón x
      Cancel = True           ' cancela la descarga
      MsgBox "Debe Utilizar el Botón de Salida para Cerrar Correctamente la Sesión de Trabajo", vbExclamation
   End If
   '
End Sub
'
Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
   On Error Resume Next
   Picture10.Height = ScaleHeight
   Picture4.Top = ScaleHeight - Picture4.Height
   Frame4.Top = ScaleHeight - Frame4.Height
   Command7.Top = Picture4.Top - Command7.Height
   Command6.Top = Command7.Top - Command6.Height
   Command6.Visible = True
   Command7.Visible = True
   Picture4.Visible = True
   If Picture4.Top < Command7.Top + Command7.Height Then Command7.Top = Picture4.Top - Command6.Height
   If Command7.Top < Command4.Top + Command4.Height + 200 Then Command7.Visible = False
   If Command6.Top < Command4.Top + Command4.Height + 200 Then Command6.Visible = False
   If Picture4.Top < Command4.Top + Command4.Height Then Picture4.Visible = False
   Frame1.Height = ScaleHeight + 250
   Picture7.Height = ScaleHeight
   Picture7.Left = ScaleWidth - 840
      On Error Resume Next
      DINFE& = XVALO(VALCONTROL$(USNBR%, "POSIFORPRIN", "FRAME3_TOP"))
      On Error GoTo 0
      If DINFE& < 750 Or DINFE& > Height Then DINFE& = 1500
   Frame3.Top = Height - DINFE&
   '
   Picture3.Top = Frame3.Top + Frame3.Height
   Picture3.Height = Height - Picture3.Top
      TreeView2.Top = Frame2.Top + Frame2.Height + Label7.Height - 50
      On Error Resume Next
      TreeView2.Height = Height - TreeView2.Top - 450
      On Error GoTo 0
   Call CARGAIMAGEN
   Picture7.Refresh
   Picture8.Refresh
   Picture10.Refresh
   On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call FINAL("")
End Sub


Private Sub frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 1
End Sub

Private Sub frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Frame1.MousePointer = 9
   If ARRABOR% = 1 Then
      If Frame1.Left + X > 20 Then
          Frame1.Left = Frame1.Left + X
          Call ACOMODATODO
      End If
   End If
End Sub

Private Sub frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 0
End Sub

Sub ACOMODATODO()
    On Error Resume Next
    Picture10.Width = Frame1.Left + Frame1.Width + 30
    Treeview1.Width = Frame1.Left + 50
    TreeView2.Width = Frame1.Left + 50
    Label6.Width = Frame1.Left + 50
    Label7.Width = Frame1.Left + 50
    Frame5.Top = Label6.Top - Frame5.Height
    On Error GoTo 0
    Call CARGAIMAGEN
    Call PONELOGO
    Label6.Top = Picture19.Height
    Treeview1.Top = Label6.Top + Label6.Height - 10
    Treeview1.Height = TreeView2.Top - Label7.Height - Treeview1.Top
    Picture7.Refresh
    Picture8.Refresh
    Picture10.Refresh
    '
    On Error Resume Next
    Label7.Top = Frame2.Top + Frame2.Height - 30
    TreeView2.Top = Frame2.Top + Frame2.Height + Label7.Height - 50
    TreeView2.Height = Height - TreeView2.Top - 450
    Treeview1.Height = TreeView2.Top - Label7.Height - Treeview1.Top
    On Error GoTo 0
    Picture10.Refresh
End Sub

Private Sub frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 1
End Sub

Private Sub frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Frame2.MousePointer = 7
   If ARRABOR% = 1 Then
      Frame2.Top = Frame2.Top + Y
      Call ACOMODATODO
   End If
End Sub

Private Sub frame2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 0
End Sub

Private Sub frame3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 1
End Sub

Private Sub frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Frame3.MousePointer = 7
   If ARRABOR% = 1 Then
      Frame3.Top = Frame3.Top + Y
      On Error Resume Next
      Picture2.Height = Frame3.Top
      Picture3.Top = Frame3.Top + Frame3.Height
      Picture3.Height = Height - Picture3.Top
      On Error GoTo 0
      Picture3.Refresh
      Call CARGAIMAGEN
   End If
End Sub

Private Sub frame3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 0
   DINFE& = Height - Frame3.Top
   Call GRAVALCON(USNBR%, "POSIFORPRIN", "FRAME3_TOP", FORMATNUM$(DINFE&, "F10.0"))
End Sub


Private Sub Image3_DblClick(Index As Integer)
   AREXE$ = VALCONTROL$(USNBR%, "ACCDIR-" + TRIM$(Str$(Index)), "DESTAPLI")
   LINCOMMA$ = VALCONTROL$(USNBR%, "ACCDIR-" + TRIM$(Str$(Index)), "LINCOMMA")
   On Error Resume Next
   NX1 = Shell(AREXE$ + " " + LINCOMMA$, vbNormalFocus)
   If Err Then
      MsgBox "Imposible Ejecutar '" + AREXE$ + " " + LINCOMMA$ + "'"
   End If
   On Error GoTo 0
End Sub

Private Sub Image3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      FORCONSHK.Label7 = ""
      FORCONSHK.Label6 = TRIM$(Str$(Index))
      ADIR$ = "ACCDIR-" + TRIM$(Str$(Index))
      FORCONSHK.Text1 = VALCONTROL(USNBR%, ADIR$, "DESTAPLI")
      FORCONSHK.Text2 = VALCONTROL(USNBR%, ADIR$, "LINCOMMA")
      FORCONSHK.Text3 = VALCONTROL(USNBR%, ADIR$, "ARCHICON")
      FORCONSHK.Text4 = VALCONTROL(USNBR%, ADIR$, "TOOTIPTE")
      '
      FORCONSHK.Show 1
      If FORCONSHK.Label7 = "OK" Then
         DESTAPLI$ = FORCONSHK.Text1
         LINCOMMA$ = FORCONSHK.Text2
         ARCHICON$ = FORCONSHK.Text3
         TOOTIPTE$ = FORCONSHK.Text4
         '
         If FORCONSHK.Label6 = "" Then
             INDIA& = Image3.Count
             Call GRAVALCON(USNBR%, "ADIRECT", "CANADIR", TRIM$(Str$(1 + INDIA&)))
           Else
             INDIA& = XVALO(FORCONSHK.Label6)
         End If
         '
         ADIR$ = "ACCDIR-" + TRIM$(Str$(INDIA&))
         Call GRAVALCON(USNBR%, ADIR$, "DESTAPLI", DESTAPLI$)
         Call GRAVALCON(USNBR%, ADIR$, "LINCOMMA", LINCOMMA$)
         Call GRAVALCON(USNBR%, ADIR$, "ARCHICON", ARCHICON$)
         Call GRAVALCON(USNBR%, ADIR$, "TOOTIPTE", TOOTIPTE$)
      End If
      '
      Call CARGACCDIR
   End If
End Sub


Private Sub Label5_Click()
    Call CARGAMENU
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
    If DERACCE%("LOGONIMG", "CAMBIAR IMAGEN PROYECTO INICIAL") < 2 Then GoTo 99
    '
    If Button = 2 Then
10      CommonDialog1.CancelError = True
        CommonDialog1.DialogTitle = "Archivo Origen de Imagen"
            ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
        CommonDialog1.FILTER = "Todos los archivos (*.*)|*.*|Mapas de Bit" & _
        "(*.BMP)|*.BMP|Archivos Graficos JPEG (*.JPG)|*.JPG"
        ' Especificar el filtro predeterminado
        CommonDialog1.FilterIndex = 2
        On Error GoTo 99
        CommonDialog1.ShowOpen
        On Error GoTo 0
        '
        DoEvents
        ARIMAG$ = UCase$(CommonDialog1.FileName)
        '
        If ARIMAG$ <> "" Then
           On Error Resume Next
           Image2.Picture = LoadPicture(ARIMAG$)
           If Not (Err) Then
              On Error GoTo 0
              Call GRAVALCON(USNBR%, "IMAGEN", "ARCHIMAGEN", ARIMAG$)
              Call CARGAIMAGEN
              GoTo 80
           End If
           On Error GoTo 0
           Call MENSERR(24, "Imposible Aplicar la Imagen Elegida.")
           GoTo 10
        End If
        '
80      Call COMUNI("Cambio Aplicado.\Re-Inicie Ahora la Aplicación.")
        Call FINAL("")
        '
    End If
99  On Error GoTo 0
    '
End Sub

Private Sub Picture3_DblClick()
'URLopen "http://www.smn.gov.ar/"
'Inet1.Password = "I(3Lei#4"
'Inet1.USERNAME = "Juan Soto"
'Text1 = Inet1.OpenURL
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      FORCONSHK.Label7 = ""
      FORCONSHK.Show 1
      If FORCONSHK.Label7 = "OK" Then
         DESTAPLI$ = FORCONSHK.Text1
         LINCOMMA$ = FORCONSHK.Text2
         ARCHICON$ = FORCONSHK.Text3
         TOOTIPTE$ = FORCONSHK.Text4
         '
         If FORCONSHK.Label6 = "" Then
             INDIA& = Image3.Count
             Call GRAVALCON(USNBR%, "ADIRECT", "CANADIR", TRIM$(Str$(1 + INDIA&)))
           Else
             INDIA& = XVALO(FORCONSHK.Label6)
         End If
         '
         ADIR$ = "ACCDIR-" + TRIM$(Str$(INDIA&))
         Call GRAVALCON(USNBR%, ADIR$, "DESTAPLI", DESTAPLI$)
         Call GRAVALCON(USNBR%, ADIR$, "LINCOMMA", LINCOMMA$)
         Call GRAVALCON(USNBR%, ADIR$, "ARCHICON", ARCHICON$)
         Call GRAVALCON(USNBR%, ADIR$, "TOOTIPTE", TOOTIPTE$)
      End If
      '
      Call CARGACCDIR
   End If
End Sub

Private Sub Picture7_Click()
  RUTAEJECUT_VBNET$ = App.Path + "\"
  XXX = Shell(RUTAEJECUT_VBNET$ + "IMPLANET.exe " & LUDAT$, 4)
   
End Sub

Private Sub Timer1_Timer()
   Static LACTIVA%
   For KKI% = 0 To 3
     Label1(KKI%).FontSize = 12
     Label1(KKI%).FontBold = False
     Label1(KKI%).ForeColor = RGB(0, 0, 0)
     If KKI% = LACTIVA% Then
       Label1(KKI%).FontSize = 12
       Label1(KKI%).FontBold = True
       Label1(KKI%).ForeColor = RGB(255, 0, 0)
     End If
   Next KKI%
   DoEvents
   LACTIVA% = (LACTIVA% + 1) Mod 4
   '
   
End Sub

Private Sub Timer2_Timer()
Exit Sub
   If EXISTE%("C:\FLEXOFT\PRUEBA.PDF") > 0 Then
     ACROL$ = FILEFIND$("C:", "ACRORD32.EXE", 0)
     If ACROL$ <> "" Then
       RUTOR$ = ACROL$ + " "
       TEXACT$ = "" '  "Adobe"
       RUTAR$ = "C:\FLEXOFT\"
       ANAME$ = RUTOR$ + RUTAR$ + "PRUEBA.PDF"
       RETURNVALUE = Shell(ANAME$, 1)
18     On Error Resume Next
       AppActivate TEXACT$ + " Reader - [PRUEBA.pdf]" ' Adobe Reader"
       'If Err Then
         'Resume 20
20       'On Error GoTo 0
         'TINI = Timer
         While Timer < TINI + 1: DoEvents: Wend
         'GoTo 18
       'End If
       On Error GoTo 0
       '
     End If
   End If
End Sub

Function PROCACTI%()
   '
   File1.Path = LUCOM$
   File1.Pattern = "*.EXK"
   File1.Refresh
   '
   For U& = 1 To File1.ListCount
      X$ = TRIM$(UCase$(File1.List(U& - 1)))
      If X$ = "ADFLESQL.EXK" Then GoTo 9         ' EXCEPCION
      If X$ = "FLEX_ADM.EXK" Then GoTo 9         ' EXCEPCION
      If Left$(X$, 6) = "SEGCOM" Then GoTo 9     ' EXCEPCION CRM
      If X$ <> "" Then
        If X$ <> App.EXEName + ".EXK" Then
          If FileLen(LUCOM$ + X$) > 0 Then
            On Error GoTo 10
            Kill LUCOM$ + X$
            On Error GoTo 0
          End If
        End If
      End If
9  Next U&
   PROCACTI% = 0
   Exit Function
   '
10 Resume 11
11 On Error GoTo 0
   PROCACTI% = 1
   X1$ = Left$(X$, Len(X$) - 1) + "E"
   TTXX$ = "Imposible Cerrar Normalmente la Sesión de Trabajo.\   \"
   TTXX$ = TTXX$ + "Se Encuentra Abierto y Ejecutando el Proceso\"
   TTXX$ = TTXX$ + "'" + X1$ + "'\  \"
   TTXX$ = TTXX$ + "Cierre este Proceso Antes de Cerrar Sesión."
   Call COMUNI(TTXX$)
   '
End Function

Private Sub Timer3_Timer()
    HOII = HOY(HOS$)
    Label4 = HOS$
    DIASE$ = DSEM$(HOII)
    Label5 = UCase$(Left$(DIASE$, 1)) + LCase$(Mid$(DIASE$, 2))
    Label3 = Left$(Time, 8)
End Sub

Private Sub Timer4_Timer()
    
    If EXISTE%("ESTOCK07.EXE") < 1 Then Timer4.Enabled = False: Exit Sub
    '
    DS% = DIASEM(HOY(HO$))
    DIADEHOY$ = DSEM$(HOY(HO$))
    FE% = HOY(HO$)
    'valida si ya grabo el mismo idea
    If FE% = XVALO(Control("GENEROAR", "FECHAGEN")) Then Timer4.Enabled = False: Exit Sub
    DIACONFIG$ = UCase$(TRIM$(Control("ESTOCK07", "DIAPROXE")))  'DIA CONFIGURADO EN ESTOCK07 (ESTADISTICO DE CONSUMO)
    If DIACONFIG$ = "" Then Timer4.Enabled = False: Exit Sub
    '
    'si es el dia configurado ingresa
    If DIADEHOY$ = DIACONFIG$ Then
       If EXISTE%(Control("ESTOCK07", "RUTASUMI")) > 0 Then
          On Error Resume Next
          Kill (Control("ESTOCK07", "RUTASUMI"))
          On Error GoTo 0
        End If
        Call CONECRUN("ESTOCK07/CONMIN", "CONTROL ESTADISTICO DE CONSUMO")
        Call GRACONTROL("GENEROAR", "FECHAGEN", SAFETEXT$(FE%))
    End If
    

End Sub

Private Sub Treeview1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      With FORMENUDY
         .Combo1.Clear
         TTX2$ = ""
         For U& = 1 To Treeview1.Nodes.Count
            TTX1$ = AJUSTI$(Treeview1.Nodes(U&).TEXT, 46) + "  " + TRIM$(Left$(Treeview1.Nodes(U&).key, 64))
            .Combo1.AddItem TTX1$
            If Treeview1.Nodes(U&).Selected = True Then
               TTX2$ = AJUSTI$(Treeview1.Nodes(U&).TEXT, 46) + "  " + TRIM$(Left$(Treeview1.Nodes(U&).key, 64))
            End If
         Next U&
         If TTX2$ <> "" Then
            .Combo1.TEXT = TTX2$
            Call .Combo1_Click
         End If
         .Show 1
      End With
      Call CARGAMENU
   End If
   '
End Sub

Private Sub Treeview1_NodeClick(ByVal Node As ComctlLib.Node)
   '
   CLANODO$ = TRIM$(Left$(Node.key, 64))
   PPXX% = InStr(CLANODO$, "/"): If PPXX% < 1 Then PPXX% = 1 + Len(CLANODO$)
   EXECONN$ = Left$(CLANODO$, PPXX% - 1)
   '
   TTXX$ = Node.TEXT
   TICONE$ = TRIM$(Left$(CLANODO$, 1))
   If TICONE$ = "?" Or TICONE$ = "*" Or EXISTE%(EXECONN$ + ".EXE") > 0 Then
      Call CUENTACCESOS(CLANODO$)
      Call CONECRUN(CLANODO$, TTXX$)
      Exit Sub
   End If
   '
End Sub

Sub CARGAIMAGEN()
    Picture2.Top = 0
    Picture2.Left = Picture10.Width
    On Error Resume Next
    Picture2.Width = Picture7.Left - Picture10.Width
    Picture2.Height = Frame3.Top
    On Error GoTo 0
    Picture2.Refresh
    Picture3.Top = Frame3.Top + Frame3.Height
    Picture3.Left = Frame1.Left + Frame1.Width
    '
    HIMAGEN = Image2.Picture.Width
    VIMAGEN = Image2.Picture.Height
    '
    HCUADRO = Picture2.Width
    VCUADRO = Picture2.Height
    '
    REDUH = HIMAGEN / HCUADRO
    REDUV = VIMAGEN / VCUADRO
    REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
    '
    X = 0: Y = 0
    NUEVA_HIMAGEN = HIMAGEN / REDUC
    If NUEVA_HIMAGEN < HCUADRO Then
        X = (HCUADRO - NUEVA_HIMAGEN) / 2
    End If
    NUEVA_VIMAGEN = VIMAGEN / REDUC
    If NUEVA_VIMAGEN < VCUADRO Then
        Y = (VCUADRO - NUEVA_VIMAGEN) / 2
    End If
    '
    Picture2.Cls
    Picture2.PaintPicture Image2.Picture, X, Y, HIMAGEN / REDUC, VIMAGEN / REDUC ' Picture2.Height, Picture2.Height
    Picture2.Visible = True
    Picture2.Refresh

End Sub

Private Sub TreeView2_NodeClick(ByVal Node As ComctlLib.Node)
   '
   CLANODO$ = TRIM$(Node.key)
   PPXX% = InStr(CLANODO$, "/"): If PPXX% < 1 Then PPXX% = 1 + Len(CLANODO$)
   EXECONN$ = Left$(CLANODO$, PPXX% - 1)
   '
   TTXX$ = Node.TEXT
   TICONE$ = TRIM$(Left$(CLANODO$, 1))
   If TICONE$ = "?" Or TICONE$ = "*" Or EXISTE%(EXECONN$ + ".EXE") > 0 Then
      Call CUENTACCESOS(CLANODO$)
      Call CONECRUN(CLANODO$, TTXX$)
      Exit Sub
   End If
   
   '
End Sub
