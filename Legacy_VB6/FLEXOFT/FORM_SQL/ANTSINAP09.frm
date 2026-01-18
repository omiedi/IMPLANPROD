VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ANTSINAP09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aplicar Factura Anticipo"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10140
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   5520
      Left            =   9240
      ScaleHeight     =   5460
      ScaleWidth      =   1320
      TabIndex        =   40
      Top             =   120
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   43
         Top             =   4680
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   44
            Top             =   0
            Width           =   12000
         End
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
         Height          =   690
         Left            =   120
         Picture         =   "ANTSINAP09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Salir"
         Top             =   225
         Width           =   690
      End
      Begin VB.CommandButton Command11 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "ANTSINAP09.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Aplicar Anticipo a la Factura Generada"
         Top             =   3840
         Width           =   690
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -285
         Picture         =   "ANTSINAP09.frx":0454
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4920
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Factura/s Anticipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5040
      TabIndex        =   25
      Top             =   5640
      Width           =   4935
      Begin VB.TextBox Total 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Index           =   1
         Left            =   3255
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   " "
         Top             =   975
         Width           =   1650
      End
      Begin VB.TextBox Total 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Index           =   0
         Left            =   3255
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   " "
         Top             =   615
         Width           =   1650
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   3240
         TabIndex        =   29
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox BASIVA 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Index           =   3
         Left            =   1800
         TabIndex        =   28
         Top             =   615
         Width           =   1395
      End
      Begin VB.TextBox BASIVA 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Index           =   2
         Left            =   1800
         TabIndex        =   27
         Top             =   975
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   3240
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1320
         Width           =   1650
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Importe "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1200
         TabIndex        =   39
         Top             =   1440
         Width           =   1980
      End
      Begin VB.Label PORIVA 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   720
         TabIndex        =   38
         Top             =   720
         Width           =   645
      End
      Begin VB.Label PORIVA 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   720
         TabIndex        =   37
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IVA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   -360
         TabIndex        =   36
         Top             =   1065
         Width           =   1065
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IVA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   -120
         TabIndex        =   35
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Ltota 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   1425
         TabIndex        =   34
         Top             =   720
         Width           =   225
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IVA No Gravado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1560
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label Ltota 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   1425
         TabIndex        =   32
         Top             =   1065
         Width           =   225
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Factura "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   10
      Top             =   5640
      Width           =   4935
      Begin VB.TextBox LABEL8 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   3240
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1320
         Width           =   1650
      End
      Begin VB.TextBox BASIVA 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Index           =   1
         Left            =   1800
         TabIndex        =   16
         Top             =   600
         Width           =   1395
      End
      Begin VB.TextBox BASIVA 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Index           =   0
         Left            =   1800
         TabIndex        =   15
         Top             =   960
         Width           =   1395
      End
      Begin VB.TextBox TLABEL28 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   3240
         TabIndex        =   14
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox Total 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Index           =   3
         Left            =   3255
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   " "
         Top             =   960
         Width           =   1650
      End
      Begin VB.TextBox Total 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Index           =   4
         Left            =   3240
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   " "
         Top             =   600
         Width           =   1650
      End
      Begin VB.Label Ltota 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   1395
         TabIndex        =   24
         Top             =   735
         Width           =   225
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IVA No Gravado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1440
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label Ltota 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   1395
         TabIndex        =   22
         Top             =   1095
         Width           =   225
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IVA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   -120
         TabIndex        =   21
         Top             =   1095
         Width           =   855
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IVA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   -360
         TabIndex        =   20
         Top             =   735
         Width           =   1065
      End
      Begin VB.Label PORIVA 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   720
         TabIndex        =   19
         Top             =   735
         Width           =   645
      End
      Begin VB.Label PORIVA 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   720
         TabIndex        =   18
         Top             =   1095
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Importe Original "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1200
         TabIndex        =   11
         Top             =   1440
         Width           =   1980
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   3
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   2205
      Left            =   45
      TabIndex        =   0
      Top             =   120
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   3889
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      AllowUserResizing=   1
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "ANTSINAP09.frx":2376
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "ANTSINAP09.frx":25AA
      TabIndex        =   4
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   2805
      Left            =   0
      TabIndex        =   5
      Top             =   2760
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   4948
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      AllowUserResizing=   1
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Facturas Anticipo: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5640
      TabIndex        =   46
      Top             =   2520
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label Label7 
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
      Height          =   315
      Left            =   7755
      TabIndex        =   45
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6840
      TabIndex        =   9
      Top             =   8400
      Width           =   1500
   End
   Begin VB.Label Label4 
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
      Height          =   315
      Left            =   8400
      TabIndex        =   8
      Top             =   8280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Factura/s Anticipo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   960
      TabIndex        =   7
      Top             =   8400
      Width           =   1980
   End
   Begin VB.Label Label9 
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
      Height          =   315
      Left            =   3000
      TabIndex        =   6
      Top             =   8280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
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
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   360
      TabIndex        =   1
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "ANTSINAP09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Sub ACTUTOTALANTI()
   TOTAL1# = XVALO(BASIVA(2)) + XVALO(BASIVA(3)) + XVALO(Total(0)) + XVALO(Total(1)) + XVALO(Text2)
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Command11_Click()
        Command11.Enabled = False
        If MSF2.Rows <= 1 Then
           OPPX% = COMALTER%("No se Ha Seleccionado Ninguna Factura Anticipo:\\Continuar\Volver para Aplicar")
           If OPPX% < 1 Or OPPX% > 2 Then Unload Me
           If OPPX% = 2 Then GoTo 99
        End If
        '
        If XVALO(BASIVA(3)) > XVALO(BASIVA(1)) + 0.005 Then
           Call COMUNI("Base Iva 10.5% No puede ser mayor que el de la factura")
           GoTo 99
         End If
        If XVALO(BASIVA(2)) > XVALO(BASIVA(0)) + 0.005 Then
           Call COMUNI("Base Iva 21.0% No puede ser mayor que el de la factura")
          GoTo 99
        End If
        '
        If XVALO(Text2) > XVALO(TLABEL28) + 0.005 Then
          Call COMUNI("Importe No Gravado a Aplicar No puede ser Mayor que el de la factura")
          GoTo 99
        End If

        If XVALO(Text1) >= XVALO(Label8) + 0.005 Then
          Call COMUNI("El Importe a Aplicar No puede ser Mayor ni igual a la Factura Generada")
          GoTo 99
        End If
        '
        FIN& = ULTREG&("!CARDEFAC")
        X2$ = REGLEIDO$("!CARDEFAC", FIN&)
        
        DOPRIC$ = TRIM$(Mid$(X2$, 387, 16))
        NORIT% = 1 + CVI(Mid$(X2$, 503, 2))
        For REG& = 1 To MSF2.Rows - 1
           NFACT$ = MSF2.TextMatrix(REG&, 2)
           IMPOR105# = XVALO(MSF2.TextMatrix(REG&, 6)) * (-1)
           IMPOR210# = XVALO(MSF2.TextMatrix(REG&, 8)) * (-1)
           IMPOEXENTO# = XVALO(MSF2.TextMatrix(REG&, 9)) * (-1)
           If Abs(IMPOR105#) > 0.005 Then
            TXT1$ = "Ant." & NFACT$ + " (Iva al 10.50%) "
            Call REPLA(TX$, MKS$(1), 1, 4)
            Call REPLA(TX$, MKS$(0), 5, 2)
            Call REPLA(TX$, TXT1$, 7, 64)
            Call REPLA(TX$, DOPRIC$, 387, 16)
            Call REPLA(TX$, MKS(IMPOR105#), 429, 4)
            Call REPLA(TX$, MKD(IMPOR105#), 433, 8)
            Call REPLA(TX$, MKS$(XVALO(FORFASQL.PORIVA(1))), 441, 4)
            Call REPLA(TX$, MKI$(NORIT%), 503, 2)
            Call REGAPP("!CARDEFAC", TX$)
            W2 = XVALO(FORFASQL.PORIVA(1))
           End If
           TXT1$ = "Factura Anticipo " & NFACT$
           If Abs(IMPOR210#) > 0.005 Then
            TXT1$ = "Ant." & NFACT$ + " (Iva al 21.00%) "
            Call REPLA(TX$, MKS$(1), 1, 4)
            Call REPLA(TX$, MKS$(0), 5, 2)
            Call REPLA(TX$, TXT1$, 7, 64)
            Call REPLA(TX$, DOPRIC$, 387, 16)
            Call REPLA(TX$, MKS(IMPOR210#), 429, 4)
            Call REPLA(TX$, MKD(IMPOR210#), 433, 8)
            Call REPLA(TX$, MKS$(XVALO(FORFASQL.PORIVA(0))), 441, 4)
            Call REPLA(TX$, MKI$(NORIT%), 503, 2)
            Call REGAPP("!CARDEFAC", TX$)
           End If
           If Abs(IMPOEXENTO#) > 0.005 Then
            TXT1$ = "Ant." & NFACT$ + " (Iva Exento) "
            Call REPLA(TX$, MKS$(1), 1, 4)
            Call REPLA(TX$, MKS$(0), 5, 2)
            Call REPLA(TX$, TXT1$, 7, 64)
            Call REPLA(TX$, DOPRIC$, 387, 16)
            Call REPLA(TX$, MKS(IMPOEXENTO#), 429, 4)
            Call REPLA(TX$, MKD(IMPOEXENTO#), 433, 8)
            Call REPLA(TX$, MKI$(NORIT%), 503, 2)
            'Call REPLA(TX$, MKS$(XVALO(PORIVA(0))), 441, 4)
            Call REGAPP("!CARDEFAC", TX$)
           End If

        Next REG&
        '
        'ESTO SE AGREGA PARA QUE CUANDO LA FACTURA A LA CUAL SE APLICA UN ANTICIPO TENGA DESCUENTO  -
        'HACE QUE TOME EL DESCUENTO Y LO PONGA EN EL DETALLE Y LO BORRE DE PORDESCU -
        'PARA QUE NO SE RECALCULE EL VALOR Y QUEDE MAL LA APLICACION DE LA FACTURA ANTICIPO.
        '
        If TRIM$(FORFASQL.PORDESCU) <> "" Then
           DESCUEX$ = FORFASQL.PORDESCU.TEXT
           IMPODESCUENX# = XVALO(FORFASQL.Total(1))
           TXT1$ = "Desc.: " & DESCUEX$ & " %"
            Call REPLA(TX$, MKS$(1), 1, 4)
            Call REPLA(TX$, MKS$(0), 5, 2)
            Call REPLA(TX$, TXT1$, 7, 64)
            Call REPLA(TX$, DOPRIC$, 387, 16)
            Call REPLA(TX$, MKS(IMPODESCUENX# * (-1)), 429, 4)
            Call REPLA(TX$, MKD(IMPODESCUENX# * (-1)), 433, 8)
            Call REPLA(TX$, MKS$(0), 441, 4)
            Call REPLA(TX$, MKI$(NORIT% + 1), 503, 2)
            Call REGAPP("!CARDEFAC", TX$)
            FORFASQL.PORDESCU = ""
        End If
        '
        'FIN CONSIDERACION DE DESCUENTO
        '
        Call FORFASQL.CARLISFAC
        Call FORFASQL.CALTOFAC
        Call CIERRARCH("!CARDEFAC")
        
        Hide
        
99 Command11.Enabled = True
        
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Load()
    
    Call APLICACIONSKINS(Me)
    
End Sub



Private Sub MSF_DblClick()
   If (MSF.MouseRow = 0) Then Exit Sub
   REG& = MSF.Row
   '
   NUFACTU$ = MSF.TextMatrix(REG&, 2)

   Condi$ = "DOCU_ORIG = '" & NUFACTU$ & "' AND TIPOMOVIM = 'FVEN' "
   TIPOMON% = XVALO(VALOBADA("FACTUANTI", "Mone_Emis", Condi$, "NOMESS"))
   If TIPOMON% < 1 Then TIPOMON% = 1
   '
   If TIPOMON% <> MONEMI% Then
     Call COMUNI("Imposible Aplicar F.Anticipo Con Diferente Moneda\Emita Factura en Misma Moneda")
     Exit Sub
   End If
   
   'If REG& < 1 Then Exit Sub
   MSF2.Rows = MSF2.Rows + 1
   MSF2.TextMatrix(MSF2.Rows - 1, 1) = MSF.TextMatrix(REG&, 1)
   MSF2.TextMatrix(MSF2.Rows - 1, 2) = MSF.TextMatrix(REG&, 2)
   MSF2.TextMatrix(MSF2.Rows - 1, 3) = MSF.TextMatrix(REG&, 3)
   MSF2.TextMatrix(MSF2.Rows - 1, 3) = MSF.TextMatrix(REG&, 3)
   Label4.ForeColor = RGB(0, 0, 0)
   If XVALO(Label4) < 0.005 Then Label4.ForeColor = RGB(255, 0, 0)
   '
   'NETO
   Condi$ = "TIPOMOVIM = 'VENTA'  AND DOCU_ORIG = '" & NUFACTU$ & "'"  'TOMAR EL IMPODEBE
   IMPOTOTAL = XVALO(VALOBADA("FACTUANTI", "SUM(IMPODEBE - IMPOHABE)", Condi$, "NOMESS"))
   TIPOCAMB = XVALO(VALOBADA("FACTUANTI", "COTIZA_DOLAR", Condi$, "NOMESS")): If TIPOCAMB <= 0 Then TIPOCAMB = 1
   'IVA1
   Condi$ = "TIPOMOVIM = 'IVA'  AND DOCU_ORIG = '" & NUFACTU$ & "' AND VARIMOVIM LIKE '" & PORIVA(1) & "%'"
   IMPOIVA1 = XVALO(VALOBADA("FACTUANTI", "SUM(IMPODEBE - IMPOHABE)", Condi$, "NOMESS"))
   BASEIMPO1 = XVALO(VALOBADA("FACTUANTI", "SUM(BASEIMPODEBE - BASEIMPOHABE)", Condi$, "NOMESS"))
   BASIVA(3) = TRIM$(FORMATNUM$(XVALO(BASIVA(3)) + BASEIMPO1, "F12.2"))
   Total(0) = TRIM$(FORMATNUM$(XVALO(Total(0)) + IMPOIVA1, "F12.2"))
   'IVA2
   Condi$ = "TIPOMOVIM = 'IVA'  AND DOCU_ORIG = '" & NUFACTU$ & "' AND VARIMOVIM LIKE '" & PORIVA(0) & "%'"
   IMPOIVA2 = XVALO(VALOBADA("FACTUANTI", "SUM(IMPODEBE - IMPOHABE)", Condi$, "NOMESS"))
   BASEIMPO2 = XVALO(VALOBADA("FACTUANTI", "SUM(BASEIMPODEBE - BASEIMPOHABE)", Condi$, "NOMESS"))
   BASIVA(2) = TRIM$(FORMATNUM$(XVALO(BASIVA(2)) + BASEIMPO2, "F12.2"))
   Total(1) = TRIM$(FORMATNUM$(XVALO(Total(1)) + IMPOIVA2, "F12.2"))
   '
   Condi$ = "TIPOMOVIM = 'VENTA'  AND DOCU_ORIG = '" & NUFACTU$ & "' AND VARIMOVIM = 'EXENTO'"
   IMPOEXENTO = XVALO(VALOBADA("FACTUANTI", "SUM(BASEIMPODEBE - BASEIMPOHABE)", Condi$, "NOMESS"))
   Text2 = FORMATNUM$(XVALO(Text2) + IMPOEXENTO, "F12.2")
  ' msf2.Rows = msf2.Rows + 1
   TOTAL_AJUS = BASEIMPO1 + BASEIMPO2 + IMPOEXENTO


   MSF2.TextMatrix(MSF2.Rows - 1, 1) = MSF.TextMatrix(REG&, 1)
   MSF2.TextMatrix(MSF2.Rows - 1, 2) = MSF.TextMatrix(REG&, 2)
   MSF2.TextMatrix(MSF2.Rows - 1, 3) = MSF.TextMatrix(REG&, 3)
   MSF2.TextMatrix(MSF2.Rows - 1, 4) = FORMATNUM$(TOTAL_AJUS, "F12.2")

   MSF2.TextMatrix(MSF2.Rows - 1, 5) = FORMATNUM$(IMPOIVA1, "F12.2")
   MSF2.TextMatrix(MSF2.Rows - 1, 6) = FORMATNUM$(BASEIMPO1, "F12.2")
   MSF2.TextMatrix(MSF2.Rows - 1, 7) = FORMATNUM$(IMPOIVA2, "F12.2")
   MSF2.TextMatrix(MSF2.Rows - 1, 8) = FORMATNUM$(BASEIMPO2, "F12.2")
   MSF2.TextMatrix(MSF2.Rows - 1, 9) = FORMATNUM$(IMPOEXENTO, "F12.2")
   '
   Text1 = TRIM$(FORMATNUM$(XVALO(BASIVA(3)) + XVALO(BASIVA(2)) + XVALO(Total(0)) + XVALO(Total(1)) + IMPOEXENTO, "F12.2"))
   Call borrafila(ANTSINAP09.MSF.Row, ANTSINAP09, ANTSINAP09.MSF)
   
End Sub
Private Sub MSF2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 2 Then
    '
    i& = MSF2.MouseRow
    J& = MSF2.MouseCol
    If i& = 0 Or i& > MSF2.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF2.Row = i&
    MSF2.COL = J&
    On Error Resume Next

    NUFACTU$ = TRIM$(MSF2.TextMatrix(i&, 2))
    If NUFACTU$ = "" Then
        Exit Sub
    End If
    '
    
    GRAV105# = XVALO(MSF2.TextMatrix(i&, 6))
    GRAV210# = XVALO(MSF2.TextMatrix(i&, 8))
    NOGRAV# = XVALO(MSF2.TextMatrix(i&, 9))
    '
    'COMPARO LOS VALORES DE LA FACTURA QUE SE ESTA REALIZANDO CONTRA EL SALDO DE LAS FACTURAS ANTICIPO
    FactGRAV105# = XVALO(BASIVA(1)) ' + XVALO(BASIVA(3))
    FactGRAV210# = XVALO(BASIVA(0)) '+ XVALO(BASIVA(2))
    FactNOGRAV# = XVALO(TLABEL28)
    '
    'SI EL DE LA FACTURA ANTICIPO ES MAYOR QUE EL DE LA FACTURA PRESENTA EL DE LA FACTURA - 0.01
    If NOGRAV# >= FactNOGRAV# + 0.005 Then FactNOGRAV# = FactNOGRAV# - 0.01
    If GRAV105# >= FactGRAV105# + 0.005 Then FactGRAV105# = FactGRAV105# - 0.01
    If GRAV210# >= FactGRAV210# + 0.005 Then FactGRAV210# = FactGRAV210# - 0.01
    '
    'CASO CONTRARIO PRESENTA EL TOTAL DEL SALDO DE LA FACTURA ANTICIPO
    If NOGRAV# < FactNOGRAV# Then FactNOGRAV# = NOGRAV#
    If GRAV105# < FactGRAV105# Then FactGRAV105# = GRAV105#
    If GRAV210# < FactGRAV210# Then FactGRAV210# = GRAV210#
    '
    Call REPLA(VDEF$, NUFACTU$, 1, 18)
    Call REPLA(VDEF$, MKD$(FactGRAV105#), 19, 8)
    Call REPLA(VDEF$, MKD$(FactGRAV210#), 27, 8)
    Call REPLA(VDEF$, MKD$(FactNOGRAV#), 35, 8)
    '
20  OBX$ = OBJPLA$("ANTIFACTU", VDEF$)
    If OBX$ = "" Then Exit Sub
    '
    OBJGRAV105 = CVD(Mid$(OBX$, 19, 8))
    OBJGRAV210 = CVD(Mid$(OBX$, 27, 8))
    OBJNOGRAVADO# = CVD(Mid$(OBX$, 35, 8))
    '
    If OBJGRAV105 > GRAV105# Then
       Call COMUNI("Imposible Ajustar Importe Gravado al 10.5% a Valor Mayor")
       VDEF$ = OBX$
       GoTo 20
    End If
    '
    If OBJGRAV210 > GRAV210# Then
       Call COMUNI("Imposible Ajustar Importe Gravado al 21.0% a Valor Mayor")
       VDEF$ = OBX$
       GoTo 20
    End If
    '
    If OBJNOGRAVADO# > NOGRAV# Then
       Call COMUNI("Imposible Ajustar Importe Exento a Valor Mayor")
       VDEF$ = OBX$
       GoTo 20
    End If

    TASA = XVALO(PORIVA(3))
    Iva = OBJGRAV105 * (TASA / 100)
    '
    MSF2.TextMatrix(i&, 5) = FORMATNUM$(Iva, "F12.2")
    MSF2.TextMatrix(i&, 6) = FORMATNUM$(OBJGRAV105, "F12.2")
    TASA = XVALO(PORIVA(2))
    Iva = OBJGRAV210 * (TASA / 100)
    '
    MSF2.TextMatrix(i&, 7) = FORMATNUM$(Iva, "F12.2")
    '
    MSF2.TextMatrix(i&, 8) = FORMATNUM$(OBJGRAV210, "F12.2")
    MSF2.TextMatrix(i&, 9) = FORMATNUM$(OBJNOGRAVADO#, "F12.2")
    MSF2.TextMatrix(i&, 4) = FORMATNUM$(XVALO(MSF2.TextMatrix(i&, 5)) + XVALO(MSF2.TextMatrix(i&, 6)) + XVALO(MSF2.TextMatrix(i&, 7)) + XVALO(MSF2.TextMatrix(i&, 8)) + XVALO(MSF2.TextMatrix(i&, 9)), "F12.2")
    '
    Call ACTU_MSF2
    '
  End If
End Sub

Private Sub MSF2_DblClick()
   i& = MSF2.MouseRow
   If i& = 0 Then Exit Sub
   
   REG& = MSF2.Row
'   If REG& < 1 Then Exit Sub
   MSF.Rows = MSF.Rows + 1
   MSF.TextMatrix(MSF.Rows - 1, 1) = MSF2.TextMatrix(REG&, 1)
   MSF.TextMatrix(MSF.Rows - 1, 2) = MSF2.TextMatrix(REG&, 2)
   MSF.TextMatrix(MSF.Rows - 1, 3) = MSF2.TextMatrix(REG&, 3)
   NUFACTU$ = MSF2.TextMatrix(REG&, 2)
   'NETO
  
'   CONDI$ = "TIPOMOVIM = 'VENTA'  AND DOCU_ORIG = '" & NUFACTU$ & "'"  'TOMAR EL IMPODEBE
'   IMPOTOTAL = FORMATNUM$(XVALO(VALOBADA("FACTUANTI", "SUM(IMPODEBE - IMPOHABE)", CONDI$, "NOMESS")), "F12.2")
'
'   'IVA1
'   CONDI$ = "TIPOMOVIM = 'IVA'  AND DOCU_ORIG = '" & NUFACTU$ & "' AND VARIMOVIM LIKE '" & PORIVA(0) & "%'"
'   IMPOIVA1 = XVALO(VALOBADA("FACTUANTI", "SUM(IMPODEBE - IMPOHABE)", CONDI$, "NOMESS"))
'   BASEIMPO1 = XVALO(VALOBADA("FACTUANTI", "SUM(BASEIMPODEBE - BASEIMPOHABE)", CONDI$, "NOMESS"))
'   BASIVA(3) = XVALO(BASIVA(3)) - BASEIMPO1
'   Total(0) = FORMATNUM$(XVALO(Total(0)) - IMPOIVA1, "F12.2")
'   'IVA2
'   CONDI$ = "TIPOMOVIM = 'IVA'  AND DOCU_ORIG = '" & NUFACTU$ & "' AND VARIMOVIM LIKE '" & PORIVA(1) & "%'"
'   IMPOIVA2 = XVALO(VALOBADA("FACTUANTI", "SUM(IMPODEBE - IMPOHABE)", CONDI$, "NOMESS"))
'   BASEIMPO2 = XVALO(VALOBADA("FACTUANTI", "SUM(BASEIMPODEBE - BASEIMPOHABE)", CONDI$, "NOMESS"))
'   BASIVA(2) = FORMATNUM$(XVALO(BASIVA(2)) - BASEIMPO2, "F12.2")
'   Total(1) = FORMATNUM$(XVALO(Total(1)) - IMPOIVA2, "F12.2")

   

   Call borrafila(ANTSINAP09.MSF2.Row, ANTSINAP09, ANTSINAP09.MSF2)
   Call ACTU_MSF2
'   Call TXTBUSCAR_Change

End Sub
Sub ACTU_MSF2()
   BASIVA(3) = "": BASIVA(2) = ""
   Total(0) = 0: Total(1) = ""
   For i& = 1 To MSF2.Rows - 1
    IMPOIVA1 = XVALO(MSF2.TextMatrix(i&, 5))
    BASEIMPO1 = XVALO(MSF2.TextMatrix(i&, 6))
    BASIVA(3) = TRIM$(FORMATNUM$(XVALO(BASIVA(3)) + BASEIMPO1, "F12.2"))
    Total(0) = TRIM$(FORMATNUM$(XVALO(Total(0)) + IMPOIVA1, "F12.2"))
    'IVA2
    IMPOIVA2 = XVALO(MSF2.TextMatrix(i&, 7))
    BASEIMPO2 = XVALO(MSF2.TextMatrix(i&, 8))
    BASIVA(2) = TRIM$(FORMATNUM$(XVALO(BASIVA(2)) + BASEIMPO2, "F12.2"))
    Total(1) = TRIM$(FORMATNUM$(XVALO(Total(1)) + IMPOIVA2, "F12.2"))
    'EXENTO
    IMPOEXENTO# = XVALO(MSF2.TextMatrix(i&, 9))
    Text2 = FORMATNUM$(IMPOEXENTO#, "F12.2")
   Next i&
   Text1 = TRIM$(FORMATNUM$(XVALO(BASIVA(3)) + XVALO(BASIVA(2)) + XVALO(Total(0)) + XVALO(Total(1)) + IMPOEXENTO#, "F12.2"))

End Sub

