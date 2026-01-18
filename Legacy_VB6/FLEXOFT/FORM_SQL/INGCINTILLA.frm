VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form INGCINTILLA 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe de Ingreso de Material a Depósito de Cintillas"
   ClientHeight    =   4875
   ClientLeft      =   1050
   ClientTop       =   1155
   ClientWidth     =   13005
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   13005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1305
      Left            =   0
      TabIndex        =   8
      Top             =   -120
      Width           =   12975
      Begin VB.CommandButton Command3 
         Caption         =   "Transferir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   11160
         TabIndex        =   34
         Top             =   915
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Transferir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   12360
         TabIndex        =   33
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox TXTCODIGOELEM 
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
         Height          =   330
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   150
         Width           =   1530
      End
      Begin VB.TextBox TXTDESCRIELEM 
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
         Height          =   330
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   150
         Width           =   4650
      End
      Begin VB.TextBox TXTLOTE 
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
         Height          =   330
         Left            =   9615
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   520
         Width           =   1395
      End
      Begin VB.TextBox TXTPESO 
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
         Height          =   330
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   520
         Width           =   1125
      End
      Begin VB.CommandButton command15 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1710
         Picture         =   "INGCINTILLA.frx":0000
         TabIndex        =   14
         Top             =   920
         Width           =   175
      End
      Begin VB.TextBox TXTMETROS 
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
         Height          =   330
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   915
         Width           =   1125
      End
      Begin VB.TextBox TXTPESOORIG 
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
         Height          =   330
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   915
         Width           =   1365
      End
      Begin VB.CommandButton Command4 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6720
         Picture         =   "INGCINTILLA.frx":030A
         TabIndex        =   11
         Top             =   900
         Width           =   300
      End
      Begin VB.TextBox txtcodintelem 
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
         Height          =   330
         Left            =   10335
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Código Interno"
         Top             =   915
         Width           =   675
      End
      Begin VB.TextBox TXTLECTURA 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   9
         Top             =   120
         Width           =   2355
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   ">>>>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   156
         Width           =   612
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código Elemento:"
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
         Index           =   5
         Left            =   3480
         TabIndex        =   31
         Top             =   150
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
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
         Index           =   9
         Left            =   6765
         TabIndex        =   30
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
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
         Index           =   2
         Left            =   9120
         TabIndex        =   29
         Top             =   525
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Peso Actual:"
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
         Index           =   3
         Left            =   6750
         TabIndex        =   28
         Top             =   525
         Width           =   1095
      End
      Begin VB.Label Label13 
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
         Height          =   330
         Left            =   1200
         TabIndex        =   27
         Top             =   915
         Width           =   495
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep. Corte  deCintillas:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -120
         TabIndex        =   26
         Top             =   840
         Width           =   1170
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
         Height          =   330
         Left            =   1920
         TabIndex        =   25
         Top             =   915
         Width           =   2325
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Metros:"
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
         Index           =   10
         Left            =   7170
         TabIndex        =   24
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep. General"
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
         Left            =   -740
         TabIndex        =   23
         Top             =   600
         Width           =   1890
      End
      Begin VB.Label lbldepogeneral 
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
         Height          =   330
         Left            =   1200
         TabIndex        =   22
         Top             =   555
         Width           =   495
      End
      Begin VB.Label Label3 
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
         Height          =   330
         Left            =   1920
         TabIndex        =   21
         Top             =   555
         Width           =   2325
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Peso Original:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   4375
         TabIndex        =   20
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CI:"
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
         Index           =   6
         Left            =   10035
         TabIndex        =   19
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Materiales Transferidos  a Depósito de Cintillas (Solo con Saldo)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   12975
      Begin MSFlexGridLib.MSFlexGrid MSF0 
         Height          =   3015
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Width           =   12885
         _ExtentX        =   22728
         _ExtentY        =   5318
         _Version        =   393216
         Rows            =   26
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1440
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "INGCINTILLA.frx":0614
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   3240
      OleObjectBlob   =   "INGCINTILLA.frx":0848
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   1650
      Left            =   0
      TabIndex        =   4
      Top             =   7560
      Width           =   15465
      _ExtentX        =   27279
      _ExtentY        =   2910
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame4 
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   7100
      Width           =   15495
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "<<<<<<    Detalle de Lotes Asignados a La Orden de Fabricación Activa    >>>>>>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   2
         Left            =   4000
         TabIndex        =   6
         Top             =   120
         Width           =   8970
      End
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuedit 
      Caption         =   "Edición"
   End
   Begin VB.Menu mnuCons 
      Caption         =   "Consultas y Listados"
   End
   Begin VB.Menu mnuut 
      Caption         =   "Utilidades"
   End
   Begin VB.Menu mnuaccesodirecto 
      Caption         =   "Accesos Directos"
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuración"
      Begin VB.Menu mnuDepoCinESl 
         Caption         =   "Depósito de Cintilla y Eslingas"
      End
   End
End
Attribute VB_Name = "INGCINTILLA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub blanformu1()
        TXTLECTURA = ""
        TXTCODIGOELEM = ""
        TXTPESOORIG = ""
        TXTDESCRIELEM = ""
        TXTPESO = ""
        TXTMETROS = ""
        TXTLOTE = ""
        txtcodintelem = ""
        TXTKGS = ""
End Sub






Private Sub Command3_Click()
    Command3.Enabled = False
    CI1% = CODINT(TXTCODIGOELEM)
    If CI1% < 1 Then
       Call COMUNI("Falta Seleccionar Código")
       GoTo 99
    End If
    DEPOALMACEN% = XVALO(Me.lbldepogeneral)
    If DEPOALMACEN% < 1 Then
      Call COMUNI("Falta Depósito General")
      GoTo 99
    End If
    
    DEPOCINTILLA% = Me.Label13
    If DEPOCINTILLA% < 1 Then
      Call COMUNI("Falta Depósito de Corte de Cintillas")
      GoTo 99
    End If
    If XVALO(TXTMETROS) <= 0 Then
       Call COMUNI("Imposible Transferir Cantidad de Metros 0")
       GoTo 99
    End If
    
    FECHA% = HOY(HO$)
    CMOV$ = "TR"
    
    CANTI = XVALO(TXTMETROS)
    LOTE$ = TRIM$(TXTLOTE)
    NCOMP& = ProNroComprobante(CMOV$)
    DOPRI$ = "TR." + TRIM$(str$(NCOMP&))
    DOSEC$ = LOTE$
    DESCRIDEPO$ = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPOALMACEN%)))
    REFE$ = "Tr.de " & TRIM$(DESCRIDEPO$)
    Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC%, DEPOALMACEN%, (-1) * CANTI)    'GUARDO EL NUMERO DE OPERACION DE CADA ITEMS TRANSFERIDO
    
    DESCRIDEPO$ = ECOARCH$("TADEPOSI", Chr$(DEPOCINTILLA%))
    REFE$ = "Tr.A " & TRIM$(DESCRIDEPO$)
    NCOMP& = ProNroComprobante(CMOV$)
    
    Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC%, DEPOCINTILLA%, CANTI)  'GUARDO EL NUMERO DE OPERACION DE CADA ITEMS TRANSFERIDO
    Call ACSALOTE(CI1%, LOTE$, "NOMESS")
    Dim CTP As New Conteplast
    Call CTP.TransfIngresada(MSF0, DEPOCINTILLA%)
    
99  Command3.Enabled = True
    Screen.MousePointer = 1
    Call blanformu1
End Sub
Private Sub Command4_Click()
   IDLOT$ = TRIM$(TXTLOTE.TEXT)
   If IDLOT$ <> "" Then
     On Error Resume Next
     CONSUMLOT07.Label4 = TXTCODIGOELEM
     If Err Then Exit Sub
     On Error GoTo 0
     CONSUMLOT07.Label1 = IDLOT$
     CONSUMLOT07.Show 1
   End If
End Sub




Private Sub Command40_Click()

End Sub

Private Sub Command6_Click()

End Sub


Private Sub Form_Load()
'   NO MUESTRO LOS KG POR QUE LOS MISMOS NO SE ACTUALIZAN
    Campos$ = "Código/A14;Descripción/A26;Lote/A12;Kgs/A0;Saldo Mts./f10.2;CI/a0;DOPRILAR/A0"
    Call CARGA_ENCABEZADO(Me.MSF0, Campos$, Me)

    DEPOCINTILLA% = XVALO(CONTROL("CONTEPLA", "DEPOCINT"))
    If DEPOCINTILLA% < 1 Then
       Call COMUNI("Debe Configurar Depósito de Corte de Cintillas")
       Exit Sub
    End If
    Label13 = XVALO(CONTROL("CONTEPLA", "DEPOCINT"))
    lbldepogeneral = CONTROL$("", "DEPOCEN")
    
    Dim CTP As New Conteplast
    Call CTP.TransfIngresada(MSF0, DEPOCINTILLA%)
    Call CARGA_ENCABEZADO(MSF2, "Código/A16;Descripción/A28;Lote/A8;Cant.Cortes/F10.1;CI/a0;N°Op./F5;Descrip/A12;Hora/A0;Material/A16;Ubicación/A12;NroInf/a0;IDUbic/a0;N°Operario/F9;Nombre/A16;Peso/f8.2;IdReg/a0", Me)
    


End Sub

Private Sub Label13_Change()
    Label6 = TRIM$(ECOARCH$("TADEPOSI", Chr$(XVALO(Label13))))
End Sub

Private Sub Label5_DblClick(Index As Integer)
   If Me.Height >= 9600 Then
      Me.Height = 7860
   ElseIf Me.Height <= 8500 Then
      Me.Height = 9915
   End If
End Sub


Private Sub lbldepogeneral_Change()
   DEPO% = XVALO(lbldepogeneral)
   Label3 = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPO%)))
End Sub



Private Sub mnuDepoCinESl_Click()
    ConfigDepo.Show 1
End Sub

Private Sub MSF0_DblClick()
    REG& = MSF0.MouseRow
    COL& = MSF0.MouseCol
    If REG& = 0 Or REG& >= MSF0.Rows Then Exit Sub
    If COL& > 0 Then Exit Sub
    IDLOT$ = TRIM$(MSF0.TextMatrix(REG&, 3))
    If IDLOT$ <> "" Then
     On Error Resume Next
     CONSUMLOT07.Label4 = TRIM$(MSF0.TextMatrix(REG&, 1))
     If Err Then Exit Sub
     On Error GoTo 0
     CONSUMLOT07.Label1 = IDLOT$
     CONSUMLOT07.Show 1
    End If
    
    
End Sub


Private Sub MSF0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    REG& = MSF0.MouseRow
    COL& = MSF0.MouseCol
    If Button = 2 Then
        Dim CTP As New Conteplast
    
        If REG& = 0 Or REG& >= MSF0.Rows Then Exit Sub
        OPX% = COMALTER%("Confirmar ELIMINACIÓN  de Registro de Modo Permanente\\Cancelar\Eliminar Registro Seleccionado")
        If OPX% < 2 Then Exit Sub
        If OPX% = 2 Then
           DEPOCINTILLA% = XVALO(CONTROL("CONTEPLA", "DEPOCINT"))
           LOTE$ = MSF0.TextMatrix(REG&, 3)
           PesoOriginal# = PesoOriginalSegúnLote#(LOTE$)
           CI1% = XVALO(MSF0.TextMatrix(REG&, 6))
           NROTRANSF$ = MSF0.TextMatrix(REG&, 7)
           CONDI$ = " IDENLOTE = '" & LOTE$ & "'   AND CODIMOVI = 'TR'  AND DOPRILAR <> '" & NROTRANSF$ & "'"
           If CantRegistros&("MOVISTO", CONDI$, "NOMESS") < 1 Then
             CONDI$ = "CODIMOVI = 'TR' AND DOPRILAR = '" & NROTRANSF$ & "' AND IDENLOTE = '" & LOTE$ & "' AND COD_INTE = " & CI1%
             Call BORRAREGISTROS("MOVISTO", CONDI$, REGAF&, "NOMESS")
             Call CTP.TransfIngresada(MSF0, DEPOCINTILLA%)
           Else
             Call COMUNI("Imosible Eliminar Este Registro\ El mismo Posee Transferencias e Infomes de Asignación a Orden de Fabricación")
             Exit Sub
           End If
        End If

    End If
End Sub

Private Sub MSF1_Click()

End Sub

Private Sub TXTLECTURA_DblClick()
'   TXTLECTURA = IngresoManualCodigoLote$
End Sub

Private Sub TXTLECTURA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     LecturaEtiqueta$ = TRIM$(TXTLECTURA)
     LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, "'", "-")
     LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, "]", "|")
     If LecturaEtiqueta$ = "" Then Exit Sub
     N% = InStr(1, LecturaEtiqueta$, "|")
     If N% < 1 Then Exit Sub
     Dim CTP As New Conteplast

     Dim strArray() As String
     LECTURA$ = LecturaEtiqueta$
     On Error Resume Next
     strArray = Split(LECTURA$, "|")
     TXTCODIGOELEM = strArray(0)
     LOTE$ = TRIM$(strArray(1))
     On Error GoTo 0
     '
     CIELEM% = CODINT%(TXTCODIGOELEM)
     
     'ESTO ES POR EL KILOMBO QUE HICIERON QUE CAMBIARON CODIGOS DESPUES DE ESTAR GENERADAS LAS ETIQUETAS
     If CIELEM% < 1 Then
       CIELEM% = XVALO(VALOBADA("BOLRECEP", "COD_INTE", "IDENLOTE = '" & LOTE$ & "'", "NOMESS"))
       TXTCODIGOELEM = CODEXT$(CIELEM%)
     End If
     Me.TXTDESCRIELEM = DESCRIT0$(CIELEM%)

     'CICONJ% = Me.txtcodintelem
     Me.txtcodintelem = CIELEM%
     DEPOPAÑOL% = XVALO(lbldepogeneral)
     CANTMETROSBOBINA = SaldoLote(CIELEM%, LOTE$, DEPOPAÑOL%)
     Me.TXTMETROS = CANTMETROSBOBINA
'     PesoPorMetro# = PESMETRO(CIELEM%)
     PesoPorMetro# = CoeficientePesoPorMetroSegunLote#(LOTE$)
     TXTPESOORIG = FORMATNUM$(PesoOriginalSegúnLote#(LOTE$), "F10.2")

     PESOSALDOBOBINA = PesoPorMetro# * CANTMETROSBOBINA
     '
     Me.TXTLOTE = LOTE$
     Me.TXTPESO = FORMATNUM$(PESOSALDOBOBINA, "F10.2")
     On Error GoTo 0
   End If

End Sub

Private Sub TXTNUMOPERARIO_Change()
End Sub


