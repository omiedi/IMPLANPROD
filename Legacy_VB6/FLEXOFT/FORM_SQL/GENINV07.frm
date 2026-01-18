VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form GENINV07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generación de Inventarios a Fecha"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6555
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
   ScaleHeight     =   5385
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   210
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      Caption         =   "SELECCIÓN POR DEPÓSITO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   2640
      Width           =   5370
      Begin VB.CheckBox Check5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0D0D0&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   480
         TabIndex        =   26
         Top             =   360
         Width           =   225
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Incluir Existencia Depósitos No Disponibles"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   840
         TabIndex        =   34
         Top             =   360
         Width           =   3825
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      Caption         =   "VALORIZACIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   20
      Top             =   1720
      Width           =   5370
      Begin VB.CheckBox Check6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0D0D0&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3600
         TabIndex        =   38
         Top             =   360
         Width           =   225
      End
      Begin VB.TextBox Text2 
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
         Height          =   300
         Left            =   705
         TabIndex        =   22
         Top             =   405
         Width           =   540
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
         Left            =   1245
         TabIndex        =   21
         Top             =   405
         Width           =   175
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sólo artículos valorizados en la moneda elegida"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   610
         Left            =   3950
         TabIndex        =   37
         Top             =   150
         Width           =   1665
      End
      Begin VB.Label Label2 
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
         Height          =   300
         Left            =   1455
         TabIndex        =   24
         Top             =   405
         Width           =   1740
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -1245
         TabIndex        =   23
         Top             =   510
         Width           =   1905
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   5715
      Left            =   5670
      ScaleHeight     =   5685
      ScaleWidth      =   1035
      TabIndex        =   16
      Top             =   -105
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   35
         Top             =   5040
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   240
         Top             =   2880
      End
      Begin VB.CommandButton command1 
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
         Picture         =   "GENINV07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Exit"
         Top             =   315
         Width           =   640
      End
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
         Height          =   650
         Left            =   105
         Picture         =   "GENINV07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Start"
         Top             =   1995
         Width           =   640
      End
      Begin VB.CommandButton Command29 
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
         Picture         =   "GENINV07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Help"
         Top             =   965
         Width           =   640
      End
      Begin VB.Label Label10 
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   3840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label9 
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3480
         Visible         =   0   'False
         Width           =   615
      End
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
      Height          =   300
      Left            =   5300
      TabIndex        =   1
      Top             =   105
      Width           =   175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "AVANCE DE PROCESO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   5370
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0D0D0&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   525
         TabIndex        =   13
         Top             =   1365
         Width           =   225
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0D0D0&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   525
         TabIndex        =   11
         Top             =   1050
         Width           =   225
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0D0D0&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   525
         TabIndex        =   9
         Top             =   735
         Width           =   225
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0D0D0&
         Caption         =   "Check1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   525
         TabIndex        =   7
         Top             =   420
         Width           =   225
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Generación Inventario Analitico y Global"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1050
         TabIndex        =   14
         Top             =   1365
         Width           =   3585
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicialización de Archivos de Inventarios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1050
         TabIndex        =   12
         Top             =   1050
         Width           =   3585
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Recálculo de Existencias"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1050
         TabIndex        =   10
         Top             =   735
         Width           =   3585
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicialización Archivos de Transferencia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1050
         TabIndex        =   8
         Top             =   420
         Width           =   3585
      End
   End
   Begin VB.TextBox Text1 
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
      Height          =   300
      Left            =   4230
      TabIndex        =   0
      Text            =   " "
      Top             =   105
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "SELECCIÓN DE COSTOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   120
      TabIndex        =   2
      Top             =   630
      Width           =   5370
      Begin VB.OptionButton Option7 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "C. Calculado (s/ precio de compra)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   3650
         TabIndex        =   39
         ToolTipText     =   "Busca Primero en lista vigente si devulve 0 busca costo de reposición"
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Suministros"
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
         Left            =   5760
         TabIndex        =   33
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton Option5 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Lista Vigente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   1700
      End
      Begin VB.OptionButton Option6 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Costo Informado / Lista Vigente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   1900
         TabIndex        =   31
         ToolTipText     =   "Busca Primero en lista vigente si devulve 0 busca costo de reposición"
         Top             =   580
         Width           =   1695
      End
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Precio de Venta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3650
         TabIndex        =   15
         Top             =   260
         Width           =   1590
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Costo Calculado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1900
         TabIndex        =   4
         Top             =   260
         Width           =   1700
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Costo Informado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   250
         Value           =   -1  'True
         Width           =   1700
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1080
      OleObjectBlob   =   "GENINV07.frx":075E
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "GENINV07.frx":0992
      TabIndex        =   30
      Top             =   420
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Inventario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2290
      TabIndex        =   5
      Top             =   200
      Width           =   1785
   End
End
Attribute VB_Name = "GENINV07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Command2.Enabled = True
    Hide
    Call FRESHALL
End Sub
'

Sub Command2_Click()
    Command2.Enabled = False
    Call GENINV2
    Command2.Enabled = True
End Sub
'

Private Sub Command29_Click()
   Command29.Enabled = False
   Call HELPONLINE("INVENTAR")
   Command29.Enabled = True
End Sub

Private Sub Command3_Click()
    Call SELECHO("SELFECHA")
    FECHINVE$ = VALACT1$("SELFECHA")
    If FECHANUM(FECHINVE$) > 0 Then
        Text1.TEXT = FECHINVE$
        Command2.SetFocus
      Else
        Text1.TEXT = HOS$
    End If
End Sub
'

Private Sub Command4_Click()
    '
    '\\\OT-07-0018-RG-23/01/07///
    Call SELECHO("TAMONED")
    MON$ = VALACT1$("TAMONED")
    If MON$ <> "" Then
       Text2.TEXT = MON$
    End If
    '\\\OT-07-0018-RG-FIN///
    '
End Sub


Private Sub Form_Load()
    HO% = HOY(HOS$)
    Text1.TEXT = HOS$
    '
    '\\\OT-07-0018-RG-23/01/07///
    MON% = XVALO(Text2)
    Label2.Caption = ECOARCH$("TAMONED", Chr(MON%))
    '\\\OT-07-0018-RG-FIN///
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    If InStr(1, EMPREAC$, "FERVI") > 0 Then
        Option1.Visible = True
    End If
    '
End Sub '

Private Sub Label9_Change()
   KKI& = XVALO(Label9)
   FIN& = XVALO(Label10)
   Call TRACE(20, KKI& Mod FIN&, FIN&)
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("SELFECHA")
    FECHINVE$ = VALACT1$("SELFECHA")
    If FECHANUM(FECHINVE$) > 0 Then
        Text1.TEXT = FECHINVE$
        Command2.SetFocus
      Else
        Text1.TEXT = HOS$
    End If
End Sub

Sub GENINV2()
    '
    Dim VESAL$()
    ReDim VESAL$(NUMAS%)
    '
    HOII% = HOY(HOS$)
    FECHINVE$ = Text1
    If FECHANUM(FECHINVE$) < 1 Then
        Text1 = HOS$
        Text1.SetFocus
        Exit Sub
    End If
    '
    '\\\OT-07-0018-RG-23/01/07///
    If Text2 = "" Then
       Call MENSERR(24, "Falta Ingresar el Tipo de Moneda")
       Text2.SetFocus
       Exit Sub
    Else
       MUN% = XVALO(Text2)
       MONDESC$ = ECOARCH$("TAMONED", Chr(MUN%))
       If MUN% <= 0 Or MONDESC$ = "" Then
          Call MENSERR(24, "Tipo de Moneda No Valido")
          Text2 = "1"
          Text2.SetFocus
          Exit Sub
       End If
    End If
    '\\\OT-07-0018-RG-FIN///
    '
    itemsValorizadosMonedaElegida% = 0
    If Check6.Value = 1 Then
        itemsValorizadosMonedaElegida% = XVALO(Text2)
    End If
    '
    Call GRACONTROL("INVENTAR", "FECHAINV", FECHINVE$)
    Text1 = FECHINVE$
    Text1.Refresh
    Call GRACONTROL("INVENTAR", "FEGENINV", "Generado el " + HOS$ + " - " + Left$(Time$, 5) + " hs.")
    FF% = FECHANUM(FECHINVE$)
    Call GRACONTROL("INVENTAR", "MONEDINV", Text2)
    VALO$ = "NO": If itemsValorizadosMonedaElegida% > 0 Then VALO$ = "SI"
    Call GRACONTROL("INVENTAR", "ITMONELE", VALO$)
    '
    Dim COSU%(62), SUCU$(62), DISTKK%(62)
    NUSU% = 1
    For i& = 1 To ULTREG&("SUCURSAL")
      X$ = REGLEIDO$("SUCURSAL", i&)
      If i& <= 62 Then
        If Asc(Left$(X$, 1)) > 32 Then
          NUSU% = i&
          COSU%(i&) = Asc(Mid$(X$, 64, 1))
          SUCU$(i&) = Left$(X$, 20)
          DISTKK%(i&) = 1
          '\\\OT-08-0778-OD-04/09/08/// VALIDA SI ESTA TILDADA LA OPCION DE _
          "INCLUIR EXISTENCIAS EN DEPOSITOS NO DISPONIBLES "
          If Check5.Value <> 1 Then
            If UCase$(Mid$(X$, 62, 2) = "NO") Then DISTKK%(i&) = 0
          End If
          '\\\OT-08-0778-OD-FIN///
        End If
      End If
    Next i&
    Call CIERRARCH("SUCURSAL")
    '
100 GENINV07.Caption = "Generacion de Inventarios al " + FECHINVE$
    '
    Call COPYSTRU("SALDEPOS", "SALDEPO")
    '
    VALORI% = 1
    PERIODO$ = "": NIVELCOSTO% = 0
    TICO$ = "Costos: Informados"
    If GENINV07.Option3 = True Then
        VALORI% = 2: TICOS$ = "Costos: Calculados"
        '
101     Call SELECHO("SELMESA")
        PERIODO$ = TRIM$(VALACT1$("SELMESA"))
        If PERIODO$ <> "" Then
            If CantRegistros("HISTOCOSTOS", "PERIODO='" & PERIODO$ & "'") = 0 Then
                Call COMUNI("No Se Encuentra Información de Costos\para el Período Elegido.")
                GoTo 101
            End If
            NIVELCO% = COMALTER%("Nivel de Valorización:\\Costo de Suministros\Costo Directo\Costo Standard")
            If NIVELCO% = 1 Then
                NIVELCOSTO% = 0
            ElseIf NIVELCO% = 2 Then
                NIVELCOSTO% = 1
            ElseIf NIVELCO% = 3 Then
                NIVELCOSTO% = 2
            Else
                NIVELCOSTO% = 1
            End If
        Else
            Call COMUNI("Se Considera Para la Valorización del Inventario,\al Último Cálculo de Costos Realizado")
        End If
      ElseIf GENINV07.Option4 = True Then
        VALORI% = 3: TICOS$ = "Costos: S/Precios de Venta"
        Dim APRELI$(128), MONELI%(128)
         'SE CAMBIO ESTO POR QUE MONELI NO SE ESTA CARGANDO EN LA GRABACION DEL ARCHIVO SE CAMBIO A QUE LEA SQL 13/11/17
'        CALIS% = ULTREG&("LISTAS"): If CALIS% > 64 Then CALIS% = 64
'        For KKI% = 1 To CALIS%
'          ULK& = KKI%
'          XX$ = REGLEIDO$("LISTAS", ULK&)
'          MONELI%(KKI%) = Asc(Mid$(XX$, 50, 1))
'          APRELI$(KKI%) = TRIM$(Str$(KKI%))
'          While Len(APRELI$(KKI%)) < 3: APRELI$(KKI%) = "0" + APRELI$(KKI%): Wend
'          APRELI$(KKI%) = "LIPRE" + APRELI$(KKI%)
'        Next KKI%
         ULK& = 0
         CALIS% = CantRegistros("LISTAPRE", "Cod_Lista > 0 ", "NOMESS")
         SQLX$ = "SELECT * FROM LISTAPRE WITH(NOLOCK) WHERE COD_LISTA > 0 ORDER BY COD_LISTA"
         Set rs = READSET(SQLX$)
         With rs
           Do While Not .EOF
             ULK& = ULK& + 1
             MONELI%(ULK&) = XVALO(!Moneda)
             APRELI$(ULK&) = XVALO(!Cod_Lista)
             While Len(APRELI$(ULK&)) < 3: APRELI$(ULK&) = "0" + APRELI$(ULK&): Wend
             APRELI$(KKI%) = "LIPRE" + APRELI$(ULK&)
             .MoveNext
           Loop
         End With
         rs.Close: Set rs = Nothing
         'FIN CAMBIO 13/11/17
      ElseIf GENINV07.Option5 = True Then
        VALORI% = 4: TICOS$ = "Costos: S/Lista de Precios Vigente"
      ElseIf GENINV07.Option6 = True Then
        VALORI% = 5: TICOS$ = "Costos: S/Lista de Precios Vigente / Costos Informados"
      ElseIf GENINV07.Option1 = True Then
        VALORI% = 6: TICOS$ = "Costos: Suministros"
      ElseIf GENINV07.Option7 = True Then
        ' VALORIZA EL INVENTARIO CONSIDERANDO EL COSTO DE COMPONENTES = ULTIMO PRECIO DE COMPRA
        VALORI% = 7: TICOS$ = "Costos: Calculados (s/ Precio de Compra)"
        
        NIVELCO% = COMALTER%("Nivel de Valorización:\\Costo de Suministros\Costo Directo")
        If NIVELCO% = 1 Then
            NIVELCOSTO% = 0
        ElseIf NIVELCO% = 2 Then
            NIVELCOSTO% = 1
        Else
            NIVELCOSTO% = 1
        End If
        
        If NIVELCO% = 1 Then
            TICOS$ = TICOS$ & vbCrLf & "Modo: Costo de Suministros"
        ElseIf NIVELCO% = 2 Then
            TICOS$ = TICOS$ & vbCrLf & "Modo: Costo Directo"
        End If
        
        ' CARGA TABLA DE COTIZACIONES PARA QUE EL USUARIO INGRESE EL TIPO DE CAMBIO
        ' A LA FECHA DE INVENTARIO. ESTE SERA CONSIDERADO EN LOS PRECIOS EN MONEDA EXTRANJERA
        ' AL MOMENTO DE PESIFICARLOS.
        Dim TablaCotizaciones(256) As Single
        
        VENTANA.Inicializar = 0
        VENTANA.CampEditables = "3"
        VENTANA.AgregaRegistro = 0
        VENTANA.UtilizaFormula = 0
        VENTANA.HabilitaBorrar = 0
        VENTANA.HabilitaInsertar = 0
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = ""
        VENTANA.Inicializar = 1
        VENTANA.Campox = "Id/I2;Moneda/A24;Cotizacion/F12.4"
        
        VENTABNEW.MSF.Rows = ULTREG&("TAMONED") + 1
        For i& = 1 To ULTREG&("TAMONED")
            XX$ = REGLEIDO$("TAMONED", i&)
            ID% = Asc(Mid$(XX$, 1, 1))
            Descripcion$ = Mid$(XX$, 2, 24)
            Cotizacion# = CVS(Mid$(XX$, 29, 4))
            
            VENTABNEW.MSF.TextMatrix(i&, 1) = CStr(ID%)
            VENTABNEW.MSF.TextMatrix(i&, 2) = Descripcion$
            VENTABNEW.MSF.TextMatrix(i&, 3) = FORMATNUM$(Cotizacion#, "F12.4")
            
            TablaCotizaciones(ID%) = Cotizacion#
        Next i&
        
        VENTABNEW.Show 1
        
        If VENTABNEW.APROBADO > 0 Then
            For i& = 1 To ULTREG&("TAMONED")
                ID% = XVALO(VENTABNEW.MSF.TextMatrix(i&, 1))
                TablaCotizaciones(ID%) = XVALO(VENTABNEW.MSF.TextMatrix(i&, 3))
            Next i&
        End If
        
        Set VENTANA = Nothing

        ' GUARDA EN COSUMAT Y PRHABI, EL ULTIMO PRECIO DE COMPRA INMEDIATAMENTE ANTES DE LA FECHA DE INVENTARIO
        ' Y EL PROVEEDOR HABITUAL DEL ARTICULO
        Screen.MousePointer = 11
        
        FechaInventario% = FECHANUM(FECHINVE$)
        FIN& = NUMAS
        Dim COSUMAT(), FECHACOMPRA%(), PRHABI%(), DP%()
        ReDim COSUMAT(FIN&), FECHACOMPRA%(FIN&), PRHABI%(FIN&), DP%(FIN&)
        For i& = 1 To FIN&
          Call TRACE(24, i&, FIN&)
          
          CII% = i&
          COSUMAT(CII%) = 0: PRHABI%(CII%) = 0
          SQLX$ = "SELECT OCOMDETA.PRUN_NETO, OCOMDETA.MONE_EMIS, OCOMDETA.FEMI_OCOM, MASTER.PROVHABI, MASTER.PORNACIO, MASTER.COSUNI, MASTER.MONECOS FROM OCOMDETA WITH(NOLOCK) INNER JOIN MASTER WITH(NOLOCK) ON OCOMDETA.COD_INTE = MASTER.CODINT WHERE COD_INTE=" & CII%
          SQLX$ = SQLX$ & " AND FEMI_OCOM <= '" & FETEXCOR1$(FechaInventario%) & "' ORDER BY FEMI_OCOM DESC"
          Set Rst = READSET(SQLX$)
          With Rst
            If .EOF = False Then
                MONEI% = XVALO(!Mone_Emis)
                COSUDOL = XVALO(!prun_neto)
                FEBAS% = CVINT(!FEMI_OCOM)
                If MONEI% > 1 Then
                     COSUMAT(CII%) = COSUDOL * TablaCotizaciones(MONEI%) * (1 + XVALO(!PORNACIO) / 100)
                Else
                     COSUN = COSUDOL
                     COSUMAT(CII%) = COSUN
                End If
                PRHABI%(CII%) = XVALO(!PROVHABI)
                FECHACOMPRA%(CII%) = FEBAS%
            End If
            
            If COSUMAT(CII%) < 0.005 Then
                MONEI% = MONECOSTO%(CII%)
                COSUMAT(CII%) = COSUNI(CII%) * TablaCotizaciones(MONEI%) * (1 + PORNACIO(CII%) / 100)
            End If
            
            On Error Resume Next
            Rst.Close
            Set Rst = Nothing
            On Error GoTo 0
            
          End With
          
          CAIT1% = CANCOMPO%(CII%)
          If CAIT1% > 0 Then
            DP%(CII%) = 1: FECHACOMPRA%(CII%) = 0
          End If
          
        Next i&
        
        PERIO$ = METEXTO$(FechaInventario%, 1)
        
        Screen.MousePointer = 1
        
    End If
    
    Screen.MousePointer = 11
    
    Call GRACONTROL("INVENTAR", "MODOINVE", CStr(VALORI%))
        
    ' ********************* INICIALIZANDO ARCHIVO DE TRANSFERENCIA
    '
    GENINV07.Check1.Value = 2
    FIN& = NUMAS% + 1
    '
    'Call BLANARCH("SALDEPO")
    'N2% = FILEOPEN%("SALDEPO", 0, 256)
    'Dim SSQ As String * 256
    '
    FIN& = NUMAS% + 1
    For i& = 1 To FIN&
        Call TRACE(20, i&, FIN&)
        SSQ$ = String$(255, 0) + Chr$(1)
        Call GRAREG("SALDEPO", SSQ$, i&)
    Next i&
    Call SETULTREG("SALDEPO", FIN&)
    GENINV07.Check1.Value = 1
    '
    ' ******************************* RE-CALCULANDO STOCK AL  FFS$
    '
    GENINV07.Check2.Value = 2
    '
    Call ABRESTOCKS
    '
    FECHA = FETEXCOR1$(FF%)
    '
    CONDI$ = "FechMov <= '" & FECHA & "' "
    '
    TOT& = CantRegistros&("MOVISTO", CONDI$, "NOMESS")
    If TOT& < 1 Then
      Call MENSERR(24, "No Existen Movimientos Anteriores la la fecha: \ " & CVDAT(FF%) & "")
      Exit Sub
    End If
    '
    'SQLX$ = "SELECT Cod_Inte, DepoMovi, SUM(CantIngre - CantSalid) AS SALDEPOI "
    SQLX$ = "SELECT Cod_Inte, DepoMovi, SUM(CantIngre) as SUIN,SUM(CantSalid) AS SUSAL "
    SQLX$ = SQLX$ + "FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE " & CONDI$
    SQLX$ = SQLX$ + " GROUP BY Cod_Inte, DepoMovi"
    '
    Label10 = 120
    Label9 = 0
    Timer1.Enabled = True
    Set GENIVITMP = READSET(SQLX$, 120)        ' PROVISORIO AL 29 MAR 2011 POR TIMEOUT DE TANIT
    Timer1.Enabled = False
    TOT& = 0
    With GENIVITMP
      Do While .EOF = False
        TOT& = TOT& + 1
      .MoveNext
      Loop
    End With
    GENIVITMP.Close
    '
    Set GENIVITMP = READSET(SQLX$, 120)        ' PROVISORIO AL 29 MAR 2011 POR TIMEOUT DE TANIT
    '
    With GENIVITMP
      CONTI& = 0
      Do While .EOF = False
       '
       Call TRACE(20, CONTI&, TOT&)
       '
       CI% = XVALO(!cod_inte)
       If itemsValorizadosMonedaElegida% > 0 Then
            If MONECOSTO%(CI%) <> itemsValorizadosMonedaElegida% Then GoTo 115
       End If
       DEP% = XVALO(!depomovi)
       LOT$ = "" '  (REVISAR ESTO) :If IsNull(!IdenLote) = False Then LOT$ = !IdenLote
105    SUCX% = 1
       For U% = 1 To NUSU%
          If COSU%(U%) = DEP% Then SUCX% = U%: GoTo 110
       Next U%
' 110    CANTI = !SALDEPOI * DISTKK%(SUCX%)
110    CANTI = (XVALO(!SUIN) - XVALO(!SUSAL)) * DISTKK%(SUCX%)
       RESAL& = CI% + 1
       If CI% > 0 Then
         If CI% <= NUMAS% Then
            If Len(VESAL$(CI%)) < 256 Then
               VESAL$(CI%) = Left$(VESAL$(CI%) + String$(256, 0), 256)
            End If
            SSQ$ = VESAL$(CI%)
            If Len(VESAL$(CI%)) > 0 Then
              CANUE = CANTI + CVS(Mid$(SSQ$, 4 * SUCX% - 3, 4))  ' modificado 27/05/06
              Mid$(SSQ$, 4 * SUCX% - 3, 4) = MKS$(CANUE)
              VESAL$(CI%) = SSQ$
            End If
          End If
       End If
       CONTI& = CONTI& + 1
115    .MoveNext
      Loop
    End With
    '
    ' guardar en saldepo completando el vesal$ con chr$(0)'s
    FIN& = NUMAS%
    For CIXO% = 1 To NUMAS%
      RESAL1& = CIXO% + 1
      Call TRACE(20, RESAL1&, FIN&)
      If VESAL$(CIXO%) <> "" Then
        CIXS& = CIXO%
        SSQ1$ = REGLEIDO("SALDEPO", CIXS&)
        Call GRAREG("SALDEPO", VESAL$(CIXO%) + String$(256, 0), RESAL1&)
      End If
    Next CIXO%
    GENIVITMP.Close
    Set GENIVITMP = Nothing
    '
    Check2.Value = 1
    Call CIERRARCH("SALDEPO")
    '
' ********************** INICIALIZANDO ARCHIVOS DE INVENTARIOS

    GENINV07.Check3.Value = 2
    '
    Call HEADERS("INVEDEQ", "")
    Call HEADERS("INVEDEQ", "Fecha de Inventario: " + FECHINVE$)
    Call HEADERS("INVEDEQ", TICOS$)
    If TRIM$(LOTAUT$) <> "" Then Call HEADERS("INVEDEQ", "Lote / Partida: " + LOTAUT$)
    '
    Call HEADERS("INVEGEN", "")
    Call HEADERS("INVEGEN", "Fecha de Inventario: " + FECHINVE$)
    Call HEADERS("INVEGEN", TICOS$)
    If TRIM$(LOTAUT$) <> "" Then Call HEADERS("INVEGEN", "Lote / Partida: " + LOTAUT$)
    '
    Call HEADERS("INVANAL", "")
    Call HEADERS("INVANAL", "Fecha de Inventario: " + FECHINVE$)
    Call HEADERS("INVANAL", TICOS$)
    If TRIM$(LOTAUT$) <> "" Then Call HEADERS("INVANAL", "Lote / Partida: " + LOTAUT$)

120 JP& = 0 ' Call BLANARCH("INVEDEP")
130 JN& = 0 ' Call BLANARCH("INVEGEN")
140 JL& = 0 ' Call BLANARCH("INVANAL")
   '
    Check3.Value = 1
    '
' ******************** GENERANDO INVENTARIO ANALITICO Y GLOBAL
    '
    GENINV07.Check4.Value = 2
    '
    FIN& = NUINV%
    '
    J& = 0: TTI% = 0
    URE& = ULTREG&("SALDEPO")
    '
    For i& = 2 To URE&
      Call TRACE(20, i&, URE&)
      '
      CI% = i& - 1
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          '
          REG& = CI% + 1
          SSQ$ = REGLEIDO$("SALDEPO", REG&)
          '
          SALTOTU = 0
          For U% = 1 To NUSU%
            SALTOTU = SALTOTU + Abs(CVS(Mid$(SSQ$, 4 * U% - 3, 4)))
          Next U%
          '
          If Abs(SALTOTU) >= 0.05 Then
            Total = 0
            COU = COSUNI(CI%) * TICAMONE(MONECOSTO%(CI%))
            '\\\OT-07-0018-RG-23/01/07///
            On Error Resume Next
            COU = COU / TICAMONE(MUN%)
            On Error GoTo 0
            '\\\OT-07-0018-RG-FIN///
            If VALORI% = 2 Then
                'COU = COSCALCU(CI%, MODOCO%)
                COU = COSCALCUHISTO(CI%, MODOCO%, PERIODO$, NIVELCOSTO%)
              ElseIf VALORI% = 3 Then
                COU = 0
                For KKI% = 1 To CALIS%
                  COU = PRELISTA#(APRELI$(KKI%), CI%)
                  If COU >= 0.00005 Then
                    'COU = COU * TICAMONE(MONECOSTO%(MONELI%(KKI%)))
                    COU = COU * TICAMONE(MONELI%(KKI%))
                    COU = COU / TICAMONE(MUN%) 'CALCULA SEGUN LA MONEDA ELEGIDA QUE SE DEBE PRESENTAR EL REPORTE
                    'COU = COU * TICAMONE(MONECOSTO%(CI%)) 'deberia ser asi on todo caso TICAMONE(MONELI%(KKI%)) SI MONELI SE CARGA CORRECTAMENTE
                    Exit For
                  End If
                Next KKI%
              ElseIf VALORI% = 4 Then 'LISTA DE PRECIOS VIGENTE
                  COU = PRELISTACT#(CI%, TIMONE%, 1)
                  If COU >= 0.00005 Then
                    COU = COU * TICAMONE(TIMONE%)
                  End If
                  '
              ElseIf VALORI% = 5 Then 'LISTA DE PRECIOS VIGENTE SI NO COSTOS INFORMADO
                  COU = PRELISTACT#(CI%, TIMONE%, 0)
                  If COU >= 0.00005 Then
                    COU = COU * TICAMONE(TIMONE%)
                  End If
                  '
              ElseIf VALORI% = 6 Then
                  COU = COSTOSUMI(CI%, MODOCO%)
                  '
              ElseIf VALORI% = 7 Then
                    '
                    Call EXPLOMUL(CI%)
                    COU = 0
                    For KKI% = 1 To CONTAMUL
                      If DP%(CIK%(KKI%)) < 1 Then   ' NO TIENE ESTRUCTURA
                        COSUN = COSUMAT(CIK%(KKI%))
                        COU = COU + CANT(KKI%) * COSUN
                      Else
                        If COSUMAT(CIK%(KKI%)) > 0 Then
                          PHB% = PRHABI%(CIK%(KKI%))
                          If PHB% > 0 Then
                            If PROVINT%(PHB%) = 0 Then
                              COSUN = COSUMAT(CIK%(KKI%))
                              COU = COU + CANT(KKI%) * COSUN
                            End If
                          End If
                        End If
                      End If
                    Next KKI%
                    '
                    If NIVELCOSTO% = 1 Then
                        CondicionSql$ = "PERIODO='" & PERIO$ & "' AND COD_INTE=" & CStr(CI%) & " ORDER BY FECOSCAL DESC"
                        CostoManoObra = XVALO(VALOBADA("HISTOCOSTOS", "TOTELABORACION", CondicionSql$))
                        COU = COU + CostoManoObra
                    End If
            End If
            '
            TMT% = TIMATE%(CI%)
            '
            For U% = 1 To NUSU%
              SALDE = CVS(Mid$(SSQ$, 4 * U% - 3, 4))
              If Abs(SALDE) >= 0.05 Then
                '
                Total = Total + SALDE
                '
                Z$ = REGBLAN$("INVANAL")
                Call REPLA(Z$, MKI$(CI%), 1, 2)
                Call REPLA(Z$, Chr$(COSU%(U%)), 3, 1)
                Call REPLA(Z$, Chr$(TMT%), 4, 1)
                Call REPLA(Z$, MKS$(SALDE), 5, 4)
                Call REPLA(Z$, MKS$(0), 9, 4)
                Call REPLA(Z$, MKS$(0), 13, 4)
                Call REPLA(Z$, MKD$(0), 17, 8)
                JL& = JL& + 1: Call GRAREG("INVANAL", Z$, JL&)
                '
                X$ = Left$(Z$, 8) + String$(12, 0)
                Call REPLA(X$, Chr$(TMT%), 4, 1)
                Call REPLA(X$, MKS$(COU), 9, 4)
                Call REPLA(X$, MKD$(CDbl(SALDE * COU)), 13, 8)
                JP& = JP& + 1: Call GRAREG("INVEDEP", X$, JP&)
                '
              End If
            Next U%
            '
            If Abs(Total) >= 0.05 Then
              Call REPLA(Z$, MKS$(Total), 9, 4)
              Call REPLA(Z$, MKS$(COU), 13, 4)
              Call REPLA(Z$, MKD$(COU * Total), 17, 8)
              Call GRAREG("INVANAL", Z$, JL&)
              '
              Y$ = REGBLAN$("INVEGEN")
              Call REPLA(Y$, MKI$(CI%), 1, 2)
              Call REPLA(Y$, Chr$(0), 3, 1)
              Call REPLA(Y$, Chr$(TMT%), 4, 1)
              If VALORI% = 7 Then Call REPLA(Y$, MKI$(FECHACOMPRA(CI%)), 3, 2)
              Call REPLA(Y$, MKS$(Total), 5, 4)
              Call REPLA(Y$, MKS$(COU), 9, 4)
              Call REPLA(Y$, MKD$(CDbl(Total * COU)), 13, 8)
              JN& = JN& + 1: Call GRAREG("INVEGEN", Y$, JN&)
              '
            End If
          End If
        End If
      End If
199 Next i&
    '
    Call SETULTREG("INVEDEP", JP&)
    Call SETULTREG("INVEGEN", JN&)
    Call SETULTREG("INVANAL", JL&)
    Call FILESORT("INVANAL", "", 1, 1, "ASC")
    Call FILESORT("INVEGEN", "", 1, 1, "ASC")
    Call FILESORT("INVEDEP", "", 1, 1, "ASC")
    Check4.Value = 1
    
999 Call CIERRARCH("*.*")
    FiNVEN07.Text1 = FECHINVE$
    FiNVEN07.Text2 = TRIM$(Control$("INVENTAR", "FEGENINV"))
    Screen.MousePointer = 1
    Call COMUNI("Inventario Generado Exitosamente")
    Hide
    Unload Me
    '
End Sub



Private Sub Text2_Change()
  '
  '\\\OT-07-0018-RG-23/01/07///
  If Text2 = "" Then Label2.Caption = ""
  On Error Resume Next
  If Text2 <> "" Then
     MUN% = XVALO(Text2)
     Text2.TEXT = MUN%
     Label2.Caption = Mid$(ECOARCH$("TAMONED", Chr(MUN%)), 1, 14)
  End If
  '\\\OT-07-0018-RG-FIN///
  '
End Sub

Private Sub Text2_DblClick()
   '
   '\\\OT-07-0018-RG-23/01/07///
   Call Command4_Click
   '\\\OT-07-0018-RG-FIN///
   '
End Sub


