VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form PROVLOCA 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tabla de Provincias, Localidades y Código Postal"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2160
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   19
      Top             =   4440
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Cust."
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
      Left            =   1890
      Picture         =   "PROVLOCA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Acceso a Maestro de Clientes"
      Top             =   5040
      Width           =   650
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   5685
      Left            =   6840
      ScaleHeight     =   5655
      ScaleWidth      =   1560
      TabIndex        =   15
      Top             =   0
      Width           =   1590
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   90
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   21
         Top             =   1755
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   22
            Top             =   0
            Width           =   12000
         End
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
         Height          =   650
         Left            =   90
         Picture         =   "PROVLOCA.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   1020
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
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
         Left            =   90
         Picture         =   "PROVLOCA.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   170
         Width           =   650
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Deuda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5355
      Picture         =   "PROVLOCA.frx":0896
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Composición de Deuda a Fecha"
      Top             =   6195
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   2625
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   6195
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Formato de Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   525
      TabIndex        =   9
      Top             =   5985
      Width           =   4005
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Listado Tabulado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   420
         TabIndex        =   12
         Top             =   420
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Plan Financiero de Cobranza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   420
         TabIndex        =   11
         Top             =   735
         Width           =   3270
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Gráfico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   420
         TabIndex        =   10
         Top             =   1155
         Width           =   1380
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   3990
         Y1              =   1050
         Y2              =   1050
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6765
      Begin VB.TextBox TXTPROVLET 
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
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   " "
         Top             =   420
         Width           =   435
      End
      Begin VB.TextBox TXTCODPOSTAL 
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
         Height          =   360
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1440
         Width           =   4890
      End
      Begin VB.TextBox TXTLOCALIDAD 
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
         Height          =   360
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   930
         Width           =   3930
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
         Height          =   360
         Left            =   2280
         TabIndex        =   6
         Top             =   930
         Width           =   175
      End
      Begin VB.TextBox TXTLOCALNUM 
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
         Height          =   360
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   " "
         Top             =   930
         Width           =   675
      End
      Begin VB.TextBox TXTNAMEPROVINCIA 
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
         Height          =   360
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   420
         Width           =   3930
      End
      Begin VB.CommandButton Command3 
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
         Height          =   360
         Left            =   2280
         TabIndex        =   2
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox TXTPROVNUM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   1
         Text            =   " "
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código Postal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -840
         TabIndex        =   23
         Top             =   1600
         Width           =   2325
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Localidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -855
         TabIndex        =   8
         Top             =   1080
         Width           =   2325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -840
         TabIndex        =   4
         Top             =   580
         Width           =   2325
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "PROVLOCA.frx":09E0
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "PROVLOCA.frx":0C14
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "PROVLOCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Public APROBO%
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
            APROBO% = 1
            Hide
End Sub

Private Sub Command3_Click()
    IDPROV% = ListaSeleccionProvincia
    If IDPROV% < 1 Then Exit Sub
    TXTPROVNUM = IDPROV%
End Sub

Private Sub Command4_Click()
    IDLOCALNUM& = ListaSeleccionLocalidades(XVALO(Me.TXTPROVNUM))
    Me.TXTLOCALNUM = IDLOCALNUM&
    Me.TXTCODPOSTAL = TRIM$(RESP_ZELE_VECT(3))
End Sub

Private Sub Form_Activate()
   APROBO% = 0
End Sub

Private Sub Form_Load()

  APROBO% = 0
End Sub

Private Sub TXTLOCALNUM_Change()
  If XVALO(TXTLOCALNUM) > 0 Then
      Condi$ = "ID = " & XVALO(TXTLOCALNUM)
      TXTLOCALIDAD = VALOBADA("LOCALIDAD", "NOMBRE", Condi$, "NOMESS")
   Else
      TXTLOCALIDAD = ""
   End If
End Sub

Private Sub TXTPROVNUM_Change()
   Me.TXTCODPOSTAL = ""
   Me.TXTLOCALIDAD = ""
   Me.TXTLOCALNUM = ""
   If XVALO(TXTPROVNUM) > 0 Then
      Condi$ = "ID = " & XVALO(TXTPROVNUM)
      TXTNAMEPROVINCIA = VALOBADA("PROVINCIA", "NOMBRE", Condi$, "NOMESS")
      TXTPROVLET = VALOBADA("PROVINCIA", "CODFLEX", Condi$, "NOMESS")
   Else
      TXTNAMEPROVINCIA = ""
   End If
End Sub
