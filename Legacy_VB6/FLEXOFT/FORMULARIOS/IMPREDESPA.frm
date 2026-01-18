VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form IMPREDESPA 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Puerto de Impresión"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   300
      Index           =   0
      Left            =   1800
      ScaleHeight     =   240
      ScaleWidth      =   1110
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command14 
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
      Height          =   1065
      Left            =   5640
      Picture         =   "IMPREDESPA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Aprueba, Graba e Imprime Pedido"
      Top             =   480
      Width           =   735
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E2D3C0&
      Caption         =   "IMPRESIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5475
      Begin VB.ComboBox Combo1 
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
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   3225
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   960
         Width           =   3225
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Puerto Orden de Venta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   420
         TabIndex        =   9
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Puerto Orden de Despacho"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2175
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "IMPREDESPA.frx":030A
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "IMPREDESPA.frx":053E
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      ToolTipText     =   "Identer"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Height          =   435
      Left            =   105
      TabIndex        =   1
      Top             =   1890
      Width           =   1380
   End
End
Attribute VB_Name = "IMPREDESPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command14_Click()
    Command14.Enabled = False
    
    Call GRACONTROL(IDENTER$, "PUERDESP", Combo2.TEXT)
    Call GRACONTROL("*", "PUERDESP", Combo2.TEXT)
    CONPRIPE.Label2 = "Pto. Orden de Despacho: " & Combo2.TEXT
    '
    Call GRACONTROL(IDENTER$, "PUERORVE", Combo1.TEXT)
    Call GRACONTROL("*", "PUERORVE", Combo1.TEXT)
    CONPRIPE.Label2 = "Pto. Orden de Despacho: " & Combo2.TEXT
    CONPRIPE.Label3 = "Pto. Orden de Venta: " & Combo1.TEXT
    Command14.Enabled = True
    Unload Me
End Sub

Private Sub Form_Load()
    '
    
    PDES$ = Control$(IDENTER$, "PUERORVE")
    If PDES$ = "" Then PDES$ = Control$("*", "PUERORVE")
    '
    Combo1.Clear
    Dim PRX As Printer
    On Error Resume Next
    Combo1.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo1.AddItem PRX.DeviceName
      If PRX.DeviceName = PDES$ Then
        Combo1.TEXT = PDES$
      End If
    Next
    PDES$ = Control$(IDENTER$, "PUERDESP")
    If PDES$ = "" Then PDES$ = Control$("*", "PUERDESP")
    '
    Combo2.Clear
    Dim PRX1 As Printer
    On Error Resume Next
    Combo2.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX1 In Printers
      Combo2.AddItem PRX1.DeviceName
      If PRX1.DeviceName = PDES$ Then
        Combo2.TEXT = PDES$
      End If
    Next
    Label3 = IDENTER$
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
End Sub

