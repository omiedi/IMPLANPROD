VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form SELCODLOT07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Lote por Codigo"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   30
      Top             =   720
      Visible         =   0   'False
      Width           =   1170
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
      Height          =   285
      Left            =   3390
      MaxLength       =   16
      TabIndex        =   17
      Top             =   210
      Width           =   2205
   End
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
      Left            =   5595
      TabIndex        =   16
      Top             =   210
      Width           =   175
   End
   Begin VB.Frame Frame9 
      Caption         =   "Fechas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   120
      TabIndex        =   9
      Top             =   2295
      Width           =   5700
      Begin VB.Label Label36 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   2190
         TabIndex        =   15
         Top             =   555
         Width           =   1275
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   390
         TabIndex        =   14
         Top             =   555
         Width           =   1275
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   4110
         TabIndex        =   13
         Top             =   555
         Width           =   1275
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimiento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   3885
         TabIndex        =   12
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricación:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   1980
         TabIndex        =   11
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   210
         TabIndex        =   10
         Top             =   315
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004040&
      Height          =   6000
      Left            =   6000
      ScaleHeight     =   5940
      ScaleWidth      =   2160
      TabIndex        =   6
      Top             =   0
      Width           =   2220
      Begin VB.CommandButton Command29 
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
         Height          =   795
         Left            =   90
         Picture         =   "SELCODLOT07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Aprueba la Selección"
         Top             =   2415
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
         Height          =   690
         Left            =   90
         Picture         =   "SELCODLOT07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   240
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -330
         Picture         =   "SELCODLOT07.frx":0454
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   1680
         Width           =   1515
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
      Left            =   1770
      TabIndex        =   1
      Top             =   1245
      Width           =   175
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "SELCODLOT07.frx":2376
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   1200
      OleObjectBlob   =   "SELCODLOT07.frx":25AA
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Destino:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Left            =   4300
      TabIndex        =   28
      Top             =   1680
      Width           =   1020
   End
   Begin VB.Label Label8 
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
      Height          =   285
      Left            =   4560
      TabIndex        =   27
      Top             =   1935
      Width           =   1155
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Lote Prov."
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
      Left            =   1860
      TabIndex        =   26
      Top             =   1005
      Width           =   1035
   End
   Begin VB.Label Label4 
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
      Height          =   285
      Left            =   2070
      TabIndex        =   25
      Top             =   1245
      Width           =   1680
   End
   Begin VB.Label Label1 
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
      Height          =   285
      Left            =   3825
      TabIndex        =   24
      Top             =   1245
      Width           =   1935
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicacion:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1560
      TabIndex        =   23
      Top             =   1650
      Width           =   1035
   End
   Begin VB.Label Label16 
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
      Height          =   285
      Left            =   1800
      TabIndex        =   22
      Top             =   1905
      Width           =   1275
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   405
      TabIndex        =   21
      Top             =   1905
      Width           =   1275
   End
   Begin VB.Label Label43 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo N/Res."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   -510
      TabIndex        =   20
      Top             =   1650
      Width           =   1965
   End
   Begin VB.Label Label9 
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
      Height          =   285
      Left            =   3180
      TabIndex        =   19
      Top             =   1920
      Width           =   1275
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Disponible:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Left            =   3060
      TabIndex        =   18
      Top             =   1665
      Width           =   960
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "I-R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   165
      TabIndex        =   5
      Top             =   1005
      Width           =   390
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   -660
      TabIndex        =   4
      Top             =   375
      Width           =   1965
   End
   Begin VB.Label Label2 
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
      Height          =   285
      Left            =   420
      TabIndex        =   3
      Top             =   630
      Width           =   5325
   End
   Begin VB.Label Label28 
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
      Height          =   285
      Left            =   405
      TabIndex        =   2
      Top             =   1245
      Width           =   1365
   End
   Begin VB.Label Label44 
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
      Height          =   285
      Left            =   2235
      TabIndex        =   0
      Top             =   255
      Width           =   1125
   End
End
Attribute VB_Name = "SELCODLOT07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DEPOX%
Sub LIMPIAR()
    Label4 = ""
    Label28 = ""
    Label5 = ""
    Text1 = ""
    Label16 = ""
    Label9 = ""
    Label36 = ""
    Label37 = ""
    Label38 = ""

End Sub


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Call SELECHO("MASTER")
    COD$ = VALACT1("MASTER")
    If COD$ <> "" Then
       Text1 = COD$
    End If
End Sub

Private Sub Command29_Click()
    '
    Command29.Enabled = False
    CICOMPO% = CODINT%(Text1)
    If CICOMPO% < 1 Then
       Call MENSERR(24, "Código Inexistente o no Válido")
       GoTo 99
    End If
    '
    'Exige lote solamente si es trazable
    If ESTRAZABLE%(CICOMPO%) > 0 Then
       If TRIM$(Label28) = "" Then
          Call MENSERR(24, "No se Ha Elegido el Lote")
          GoTo 99
        End If
    End If
    '
    'Si no es trazable lote seleccionado  =""
    If ESTRAZABLE%(CICOMPO%) < 1 Then
       Label28 = ""
    End If
    '
90  Command29.Enabled = True
    Hide
    Exit Sub
    '
99  Command29.Enabled = True
'
End Sub

Private Sub Command5_Click()
    CIXI% = CODINT%(Text1)
    ' Variable sin utilizar: El Label4 contiene el lote del proeevedor, no el IR.
'   DEPX% = XVALO(Label4)
    '
    DEPX% = 0
    If XVALO(Label15) <> 0 Then DEPX% = XVALO(Label15)
    Call ARMALILOT(CIXI%, DEPX%, 0, 0)
    Call SELECHO("!LILOTSE")
    VDEVU$ = VALACT1$("!LILOTSE")
    Label28 = VDEVU$
End Sub

Private Sub Form_Load()
    Call APLICACIONSKINS(Me, Picture2)

End Sub

Private Sub Text1_Change()
    CICOMPO% = CODINT%(TRIM(Text1))
    Label2 = DESCRIT0$(CICOMPO%)
    If CICOMPO% > 0 Then
       Select Case ESTRAZABLE%(CICOMPO%)
         Case Is = 0
           Label1 = "(No Trazable)"
         Case Is = 1
           Label1 = "(Trazable por Lote)"
         Case Is = 2
           Label1 = "(Trazable por Serie)"
       End Select
     Else
       Label1 = ""
    End If
    
    '
    Call BLANCFORMU
End Sub

Private Sub Label28_Change()
    'Completa datos
    CI% = CODINT%(Text1)
    LOTE$ = TRIM$(Label28)
    If CI% > 0 Then
       If LOTE$ <> "" Then
          CONDI$ = "Cod_Inte =" & CI% & " AND IdenLote ='" & LOTE$ & "'"
          '
          FEC1 = CVINT(VALOBADA("LOTINGRE", "FechAlta", CONDI$))
          Label37 = FECHATEX$(FEC1)
          '
          FEC2 = CVINT(VALOBADA("LOTINGRE", "FechFab", CONDI$))
          Label36 = FECHATEX$(FEC2)
          '
          FEC3 = CVINT(VALOBADA("LOTINGRE", "FechVenc", CONDI$))
          Label38 = FECHATEX$(FEC3)
          '
          CANTX# = VALOBADA("LOTINGRE", "CanSaldo", CONDI$)
          Label5 = FORMATNUM(CANTX#, "F13." & CNTDC$)
          Label16 = VALOBADA("LOTINGRE", "UbicFis", CONDI$)
          Label9 = VALOBADA("LOTINGRE", "Libera", CONDI$)
          Label4 = VALOBADA("LOTINGRE", "lotecol_prov", CONDI$)
          Label8 = VALOBADA("LOTINGRE", "cod_desti", CONDI$)
        End If
    End If
End Sub

Private Sub Text1_DblClick()
   Call Command2_Click
End Sub
Private Sub BLANCFORMU()
    Label37 = ""
    Label36 = ""
    Label38 = ""
    Label5 = ""
    Label16 = ""
    Label9 = ""
    'Label28 = ""
End Sub
