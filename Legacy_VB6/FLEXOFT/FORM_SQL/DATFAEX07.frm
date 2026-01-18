VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form DATFAEX07 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos Complementarios de Factura de Exportacion"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2520
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.TextBox Text3 
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
      Height          =   300
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   27
      ToolTipText     =   "Doble click para Desplegar Lista de Selección"
      Top             =   2220
      Width           =   4610
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
      Height          =   300
      Left            =   5800
      Picture         =   "DATFAEX07.frx":0000
      TabIndex        =   26
      ToolTipText     =   "Ingresa el Permiso en la Ventana Inferior"
      Top             =   2220
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
      Left            =   5800
      Picture         =   "DATFAEX07.frx":030A
      TabIndex        =   25
      ToolTipText     =   "Ingresa el Permiso en la Ventana Inferior"
      Top             =   1800
      Width           =   175
   End
   Begin VB.TextBox Text2 
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
      Height          =   300
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   24
      ToolTipText     =   "Doble click para Desplegar Lista de Selección"
      Top             =   1800
      Width           =   4610
   End
   Begin VB.ComboBox Combo6 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   21
      ToolTipText     =   "Click Para Desplegar"
      Top             =   3030
      Width           =   4905
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   7665
      Left            =   6360
      ScaleHeight     =   7635
      ScaleWidth      =   1665
      TabIndex        =   20
      Top             =   0
      Width           =   1695
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
         Height          =   690
         Left            =   120
         Picture         =   "DATFAEX07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Aprueba la Información"
         Top             =   6600
         Width           =   650
      End
      Begin VB.CommandButton Command13 
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
         Left            =   150
         Picture         =   "DATFAEX07.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1080
         Width           =   650
      End
      Begin VB.CommandButton Command5 
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
         Left            =   150
         Picture         =   "DATFAEX07.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir - Cancelar"
         Top             =   240
         Width           =   650
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "PERMISOS DE EMBARQUE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   120
      TabIndex        =   8
      Top             =   5480
      Width           =   6120
      Begin VB.CommandButton Command3 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5490
         TabIndex        =   19
         ToolTipText     =   "Elimina el Permiso Seleccionado"
         Top             =   1480
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5090
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ingresa el Permiso en la Ventana Inferior"
         Top             =   1480
         Width           =   375
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
         Left            =   240
         MaxLength       =   16
         TabIndex        =   2
         ToolTipText     =   "Ingreso el Permiso"
         Top             =   1560
         Width           =   4785
      End
      Begin VB.ListBox List2 
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
         Height          =   810
         ItemData        =   "DATFAEX07.frx":0D72
         Left            =   240
         List            =   "DATFAEX07.frx":0D74
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   5625
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lista de Permisos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   5625
      End
   End
   Begin VB.ComboBox Combo5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Click Para Desplegar"
      Top             =   2610
      Width           =   4905
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "COMPROBANTES ASOCIADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   6120
      Begin VB.Frame Frame4 
         Height          =   1100
         Left            =   5520
         TabIndex        =   17
         Top             =   550
         Width           =   15
      End
      Begin VB.Frame Frame14 
         Height          =   1100
         Left            =   3960
         TabIndex        =   16
         Top             =   550
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "                "
         Height          =   1100
         Left            =   2520
         TabIndex        =   15
         Top             =   550
         Width           =   15
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         ItemData        =   "DATFAEX07.frx":0D76
         Left            =   240
         List            =   "DATFAEX07.frx":0D78
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   550
         Width           =   5625
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Comprobante                 Punto Venta          Número"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   35
         ToolTipText     =   "Doble click Para Editar"
         Top             =   310
         Width           =   5625
      End
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Click Para Desplegar"
      Top             =   1365
      Width           =   4905
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   960
      OleObjectBlob   =   "DATFAEX07.frx":0D7A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   1440
      OleObjectBlob   =   "DATFAEX07.frx":0FAE
      TabIndex        =   23
      Top             =   -120
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comprobante:"
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
      Left            =   2640
      TabIndex        =   33
      Top             =   870
      Width           =   1230
   End
   Begin VB.Label Label14 
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
      Left            =   3960
      TabIndex        =   32
      Top             =   780
      Width           =   2100
   End
   Begin VB.Label Label10 
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
      Height          =   300
      Left            =   4845
      TabIndex        =   31
      Top             =   375
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
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
      Left            =   4560
      TabIndex        =   30
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label9 
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
      Left            =   120
      TabIndex        =   29
      Top             =   375
      Width           =   4500
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
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
      Left            =   -120
      TabIndex        =   28
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Idioma:"
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
      Left            =   525
      TabIndex        =   22
      Top             =   3150
      Width           =   675
   End
   Begin VB.Label Label17 
      Caption         =   "Label17"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Incoterms:"
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
      Left            =   270
      TabIndex        =   13
      Top             =   2745
      Width           =   960
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuit Pais:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   345
      TabIndex        =   12
      Top             =   2340
      Width           =   900
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Destino"
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
      Left            =   180
      TabIndex        =   11
      Top             =   1935
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Exportación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   0
      TabIndex        =   5
      Top             =   1275
      Width           =   960
   End
End
Attribute VB_Name = "DATFAEX07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub BLANFORMU()
  
   Label9 = ""
   Label10 = ""
   Label14 = ""
   List1.Clear
   List2.Clear
   Text1 = ""
End Sub

Private Sub Command1_Click()
    FRMZELEOEX.MSF1.Rows = 1
    FRMZELEOEX.MSF2.Rows = 1
    FRMZELEOEX.TXTBUSCAR = ""
    Call CARGACOMBO2
    FRMZELEOEX.Caption = "Selección de Destino"
    FRMZELEOEX.Show 1
    Text2 = FRMZELEOEX.RESP_ZELE

End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   '
   'VALIDACIONES
   If TRIM$(Combo1.TEXT) = "" Then
      Call COMUNI("Falta Ingresar Tipo de Exportación")
      GoTo 99
   End If
   '
   If Left$((DATFAEX07.Combo1.TEXT), 1) <> "1" Then
      If List2.ListCount > 0 And Left$(Label14, 2) = "FC" Then
        Call COMUNI("Tipo de Exportación (" & TIPO_EXPO$ & ") No Válido Con Ingreso de Exportación")
        GoTo 99
      End If
   End If
   '
   If TRIM$(Text2.TEXT) = "" Then
      Call COMUNI("Falta Ingresar Destino")
      GoTo 99
   End If
   '
   If TRIM$(Combo5.TEXT) = "" And XVALO(Mid$(Combo1.TEXT, 1, 1)) = 1 Then
      Call COMUNI("Falta Ingresar Incoterms")
      GoTo 99
   End If
   '
   If TRIM$(Combo6.TEXT) = "" Then
      Call COMUNI("Falta Ingresar Idioma")
      GoTo 99
   End If
   'FIN VALIDACIONES
   '
   If Left$(Label14, 2) = "FC" And List1.ListCount > 0 Then
      Call COMUNI("Comprobantes Asociados Solo para Nota de Crédito y Nota de Débito")
      GoTo 99
   End If
   '
   If TRIM$(Combo1.TEXT) = "" Then
      Call COMUNI("Falta Ingresar Idioma")
      GoTo 99
   End If
   '
   Label17 = "OK"
   Hide
   
99 Command2.Enabled = True

End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    'borra el permiso seleccionado
    If List2.ListIndex >= 0 Then
      List2.RemoveItem (List2.ListIndex)
      Text1.SetFocus
    End If
    '
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   If TRIM$(Text1) <> "" Then
      List2.AddItem Text1
      Text1 = ""
   End If
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Label17 = ""
   Screen.MousePointer = 1
   Command5.Enabled = True
   Hide
End Sub

Private Sub Command6_Click()
    '
    FRMZELEOEX.MSF1.Rows = 1
    FRMZELEOEX.MSF2.Rows = 1
    FRMZELEOEX.TXTBUSCAR = ""
    Call CARGACOMBO4
    FRMZELEOEX.Caption = "Selección de Cuit Pais"
    FRMZELEOEX.Show 1
    Text3 = FRMZELEOEX.RESP_ZELE
    '
End Sub

Private Sub Form_Activate()
   If Label17 = "OK" Then Call ACTUFRM
End Sub

Private Sub Form_Load()
'  Label17 = ""
'  'Call BLANFORMU
'  Screen.MousePointer = 11
'  '
'  TIFACT$ = "FC-"
'  If FORFASQL.Label7.Caption = "N.Débito:" Then TIFACT$ = "ND-"
'  If FORFASQL.Label7.Caption = "N.Crédito:" Then TIFACT$ = "NC-"
'  DATFAEX07.Label9 = RASOCLI$(XVALO(FORFASQL.Text1))
'  DATFAEX07.Label10 = FORFASQL.Text9
'  NUM$ = FORFASQL.Text11
'  While Len(NUM$) < 8: NUM$ = "0" & NUM$: Wend
'  DATFAEX07.Label14 = TIFACT$ & FORFASQL.Label12 & "-" & NUM$
'  '
'  Call CARGACOMBO1(Combo1)
'  Call CARGACOMBO2(Combo2)
'  Call CARGACOMBO4(Combo4)
'  Call CARGACOMBO5(Combo5)
'  Call CARGACOMBO6(Combo6)
'  Label12 = CONTROL("FACTURA", "PUVEFAEX")
'  Combo1.TEXT = Combo1.List(1)
'  Combo5.TEXT = Combo5.List(4)
'  '
'  'CARGO TABLA DE MOTIVOS SI NO ESTABA CARGADA PREVIAMENTE
'  If EXISTE(LUDAT$ + "\TATICOMP.RFS") > 0 Then
'     If ULTREG&("TATICOMP") < 3 Then
'            Call GRAREG("TATICOMP", Chr$(19) + "Factura", 1)
'            Call GRAREG("TATICOMP", Chr$(20) + "Nota de Credito", 2)
'            Call GRAREG("TATICOMP", Chr$(21) + "Nota de Debito", 3)
'            Call CIERRARCH("TATICOMP")
'     End If
'  End If
  Call ACTUFRM
  Call APLICACIONSKINS(Me, Picture6)
  
  Screen.MousePointer = 1
End Sub

Sub ACTUFRM()

  Screen.MousePointer = 11
  '
  Label17 = ""
  TIFACT$ = "FC-"
  If FORFASQL.Label7.Caption = "N.Débito:" Then TIFACT$ = "ND-"
  If FORFASQL.Label7.Caption = "N.Crédito:" Then TIFACT$ = "NC-"
  DATFAEX07.Label9 = rasocli$(XVALO(FORFASQL.Text1))
  DATFAEX07.Label10 = FORFASQL.Text9
  NUM$ = FORFASQL.Text11
  While Len(NUM$) < 8: NUM$ = "0" & NUM$: Wend
  DATFAEX07.Label14 = TIFACT$ & FORFASQL.Label12 & "-" & NUM$
  '
  Combo1.Clear
  'Combo2.Clear
  'Combo4.Clear
  Combo5.Clear
  Combo6.Clear
  Text2 = ""
  Text3 = ""
  '
  Call CARGACOMBO1(Combo1)
  'Call CARGACOMBO2(Combo2)
  'Call CARGACOMBO4(Combo4)
  Call CARGACOMBO5(Combo5)
  Call CARGACOMBO6(Combo6)
  Label12 = Control("FACTURA", "PUVEFAEX")
  Combo1.TEXT = Combo1.List(1)
  Combo5.TEXT = Combo5.List(4)
  '
  'CARGO TABLA DE MOTIVOS SI NO ESTABA CARGADA PREVIAMENTE
  If EXISTE(LUDAT$ + "\TATICOMP.RFS") > 0 Then
     If ULTREG&("TATICOMP") < 3 Then
            Call GRAREG("TATICOMP", Chr$(19) + "Factura", 1)
            Call GRAREG("TATICOMP", Chr$(20) + "Nota de Credito", 2)
            Call GRAREG("TATICOMP", Chr$(21) + "Nota de Debito", 3)
            Call CIERRARCH("TATICOMP")
     End If
  End If
  Screen.MousePointer = 1
End Sub
Sub CARGACOMBO2()
    '
    FRMZELEOEX.MSF2.FixedCols = 0
    FRMZELEOEX.MSF2.Cols = 1
    FRMZELEOEX.MSF1.FixedCols = 0
    FRMZELEOEX.MSF1.Cols = 1
    FRMZELEOEX.MSF2.ColAlignment(0) = 1
    FRMZELEOEX.MSF1.ColAlignment(0) = 1
    FRMZELEOEX.MSF1.Rows = 0
    FRMZELEOEX.MSF2.Rows = 0


    FRMZELEOEX.MSF2.AddItem "101    BURKINA FASO"
    FRMZELEOEX.MSF2.AddItem "102    ARGELIA"
    FRMZELEOEX.MSF2.AddItem "103    BOTSWANA"
    FRMZELEOEX.MSF2.AddItem "104    BURUNDI"
    FRMZELEOEX.MSF2.AddItem "105    CAMERUN"
    FRMZELEOEX.MSF2.AddItem "107    REP. CENTROAFRICANA."
    FRMZELEOEX.MSF2.AddItem "108    CONGO"
    FRMZELEOEX.MSF2.AddItem "109    REP.DEMOCRAT.DEL CONGO EX ZAIRE"
    FRMZELEOEX.MSF2.AddItem "110    COSTA DE MARFIL"
    FRMZELEOEX.MSF2.AddItem "111    CHAD"
    FRMZELEOEX.MSF2.AddItem "112    BENIN"
    FRMZELEOEX.MSF2.AddItem "113    EGIPTO"
    FRMZELEOEX.MSF2.AddItem "115    GABON"
    FRMZELEOEX.MSF2.AddItem "116    GAMBIA"
    FRMZELEOEX.MSF2.AddItem "117    GHANA"
    FRMZELEOEX.MSF2.AddItem "118    GUINEA"
    FRMZELEOEX.MSF2.AddItem "119    GUINEA ECUATORIAL"
    FRMZELEOEX.MSF2.AddItem "120    KENYA"
    FRMZELEOEX.MSF2.AddItem "121    LESOTHO"
    FRMZELEOEX.MSF2.AddItem "122    LIBERIA"
    FRMZELEOEX.MSF2.AddItem "123    LIBIA"
    FRMZELEOEX.MSF2.AddItem "124    MADAGASCAR"
    FRMZELEOEX.MSF2.AddItem "125    MALAWI"
    FRMZELEOEX.MSF2.AddItem "126    MALI"
    FRMZELEOEX.MSF2.AddItem "127    MARRUECOS"
    FRMZELEOEX.MSF2.AddItem "128    MAURICIO , ISLAS"
    FRMZELEOEX.MSF2.AddItem "129    MAURITANIA"
    FRMZELEOEX.MSF2.AddItem "130    NIGER"
    FRMZELEOEX.MSF2.AddItem "131    NIGERIA"
    FRMZELEOEX.MSF2.AddItem "132    ZIMBABWE"
    FRMZELEOEX.MSF2.AddItem "133    RWANDA"
    FRMZELEOEX.MSF2.AddItem "134    SENEGAL"
    FRMZELEOEX.MSF2.AddItem "135    SIERRA LEONA"
    FRMZELEOEX.MSF2.AddItem "136    SOMALIA"
    FRMZELEOEX.MSF2.AddItem "137    SWAZILANDIA"
    FRMZELEOEX.MSF2.AddItem "138    SUDAN"
    FRMZELEOEX.MSF2.AddItem "139    TANZANIA"
    FRMZELEOEX.MSF2.AddItem "140    TOGO"
    FRMZELEOEX.MSF2.AddItem "141    TUNEZ"
    FRMZELEOEX.MSF2.AddItem "142    UGANDA"
    FRMZELEOEX.MSF2.AddItem "144    ZAMBIA"
    FRMZELEOEX.MSF2.AddItem "145    TERRIT.VINCULADOS AL R UNIDO"
    FRMZELEOEX.MSF2.AddItem "146    TERRIT.VINCULADOS A ESPAÑA"
    FRMZELEOEX.MSF2.AddItem "147    TERRIT.VINCULADOS A FRANCIA"
    FRMZELEOEX.MSF2.AddItem "149    ANGOLA"
    FRMZELEOEX.MSF2.AddItem "150    CABO VERDE"
    FRMZELEOEX.MSF2.AddItem "151    MOZAMBIQUE"
    FRMZELEOEX.MSF2.AddItem "152    SEYCHELLES"
    FRMZELEOEX.MSF2.AddItem "153    DJIBOUTI"
    FRMZELEOEX.MSF2.AddItem "155    COMORAS"
    FRMZELEOEX.MSF2.AddItem "156    GUINEA BISSAU"
    FRMZELEOEX.MSF2.AddItem "157    STO.TOME Y PRINCIPE"
    FRMZELEOEX.MSF2.AddItem "158    NAMIBIA"
    FRMZELEOEX.MSF2.AddItem "159    SUDAFRICA"
    FRMZELEOEX.MSF2.AddItem "160    ERITREA"
    FRMZELEOEX.MSF2.AddItem "161    ETIOPIA"
    FRMZELEOEX.MSF2.AddItem "197    RESTO (AFRICA)"
    FRMZELEOEX.MSF2.AddItem "198    INDETERMINADO (AFRICA)"
    FRMZELEOEX.MSF2.AddItem "200    ARGENTINA"
    FRMZELEOEX.MSF2.AddItem "201    BARBADOS"
    FRMZELEOEX.MSF2.AddItem "202    BOLIVIA"
    FRMZELEOEX.MSF2.AddItem "203    BRASIL"
    FRMZELEOEX.MSF2.AddItem "204    CANADA"
    FRMZELEOEX.MSF2.AddItem "205    COLOMBIA"
    FRMZELEOEX.MSF2.AddItem "206    COSTA RICA"
    FRMZELEOEX.MSF2.AddItem "207    CUBA"
    FRMZELEOEX.MSF2.AddItem "208    CHILE"
    FRMZELEOEX.MSF2.AddItem "209    REPÚBLICA DOMINICANA"
    FRMZELEOEX.MSF2.AddItem "210    ECUADOR"
    FRMZELEOEX.MSF2.AddItem "211    EL SALVADOR"
    FRMZELEOEX.MSF2.AddItem "212    ESTADOS UNIDOS"
    FRMZELEOEX.MSF2.AddItem "213    GUATEMALA"
    FRMZELEOEX.MSF2.AddItem "214    GUYANA"
    FRMZELEOEX.MSF2.AddItem "215    HAITI"
    FRMZELEOEX.MSF2.AddItem "216    HONDURAS"
    FRMZELEOEX.MSF2.AddItem "217    JAMAICA"
    FRMZELEOEX.MSF2.AddItem "218    MEXICO"
    FRMZELEOEX.MSF2.AddItem "219    NICARAGUA"
    FRMZELEOEX.MSF2.AddItem "220    PANAMA"
    FRMZELEOEX.MSF2.AddItem "221    PARAGUAY"
    FRMZELEOEX.MSF2.AddItem "222    PERU"
    FRMZELEOEX.MSF2.AddItem "223    Puerto RICO"
    FRMZELEOEX.MSF2.AddItem "224    TRINIDAD Y TOBAGO"
    FRMZELEOEX.MSF2.AddItem "225    URUGUAY"
    FRMZELEOEX.MSF2.AddItem "226    VENEZUELA"
    FRMZELEOEX.MSF2.AddItem "227    TERRIT.VINCULADO AL R.UNIDO"
    FRMZELEOEX.MSF2.AddItem "228    TER.VINCULADOS A DINAMARCA"
    FRMZELEOEX.MSF2.AddItem "229    TERRIT.VINCULADOS A FRANCIA AMERIC."
    FRMZELEOEX.MSF2.AddItem "230    TERRIT.HOLANDESES"
    FRMZELEOEX.MSF2.AddItem "231    TER.VINCULADOS A ESTADOS UNIDOS"
    FRMZELEOEX.MSF2.AddItem "232    SURINAME"
    FRMZELEOEX.MSF2.AddItem "233    DOMINICA"
    FRMZELEOEX.MSF2.AddItem "234    SANTA LUCIA"
    FRMZELEOEX.MSF2.AddItem "235    SAN VICENTE Y LAS GRANADINAS"
    FRMZELEOEX.MSF2.AddItem "236    BELICE"
    FRMZELEOEX.MSF2.AddItem "237    ANTIGUA Y BARBUDA"
    FRMZELEOEX.MSF2.AddItem "238    S.CRISTOBAL Y NEVIS"
    FRMZELEOEX.MSF2.AddItem "239    BAHAMAS"
    FRMZELEOEX.MSF2.AddItem "240    GRENADA"
    FRMZELEOEX.MSF2.AddItem "241    ANTILLAS HOLANDESAS"
    FRMZELEOEX.MSF2.AddItem "250    AAE Tierra del Fuego - ARGENTINA"
    FRMZELEOEX.MSF2.AddItem "251    ZF La Plata - ARGENTINA"
    FRMZELEOEX.MSF2.AddItem "252    ZF Justo Daract - ARGENTINA"
    FRMZELEOEX.MSF2.AddItem "253    ZF Río Gallegos - ARGENTINA"
    FRMZELEOEX.MSF2.AddItem "254    ISLAS Malvinas - ARGENTINA"
    FRMZELEOEX.MSF2.AddItem "255    ZF Tucumán - ARGENTINA"
    FRMZELEOEX.MSF2.AddItem "256    ZF Córdoba - ARGENTINA"
    FRMZELEOEX.MSF2.AddItem "257    ZF Mendoza - ARGENTINA"
    FRMZELEOEX.MSF2.AddItem "258    ZF General Pico - ARGENTINA"
    FRMZELEOEX.MSF2.AddItem " 259    ZF Comodoro Rivadavia - ARGENTINA"
    FRMZELEOEX.MSF2.AddItem "260    ZF Iquique"
    FRMZELEOEX.MSF2.AddItem "261    ZF Punta Arenas"
    FRMZELEOEX.MSF2.AddItem "262    ZF Salta - ARGENTINA"
    FRMZELEOEX.MSF2.AddItem "263    ZF Paso de los Libres - ARGENTINA"
    FRMZELEOEX.MSF2.AddItem "264    ZF Puerto Iguazú - ARGENTINA"
    FRMZELEOEX.MSF2.AddItem "265    SECTOR ANTARTICO ARG."
    FRMZELEOEX.MSF2.AddItem "270    ZF Colón - REPÚBLICA DE PANAMÁ"
    FRMZELEOEX.MSF2.AddItem "271    ZF Winner (Sta. C. de la Sierra) - BOLIVIA"
    FRMZELEOEX.MSF2.AddItem "280    ZF Colonia - URUGUAY"
    FRMZELEOEX.MSF2.AddItem "281    ZF Florida - URUGUAY"
    FRMZELEOEX.MSF2.AddItem "283    ZF Zonamerica - URUGUAY"
    FRMZELEOEX.MSF2.AddItem "284    ZF Nueva Helvecia - URUGUAY"
    FRMZELEOEX.MSF2.AddItem "285    ZF Nueva Palmira - URUGUAY"
    FRMZELEOEX.MSF2.AddItem "286    ZF Río Negro - URUGUAY"
    FRMZELEOEX.MSF2.AddItem "287    ZF Rivera - URUGUAY"
    FRMZELEOEX.MSF2.AddItem "288    ZF San José - URUGUAY"
    FRMZELEOEX.MSF2.AddItem "291    ZF Manaos - BRASIL"
    FRMZELEOEX.MSF2.AddItem "295    MAR ARG ZONA ECO.EX"
    FRMZELEOEX.MSF2.AddItem "296    RIOS ARG NAVEG INTER"
    FRMZELEOEX.MSF2.AddItem "297    RESTO AMERICA"
    FRMZELEOEX.MSF2.AddItem "298    INDETERMINADO (AMERICA)"
    FRMZELEOEX.MSF2.AddItem "301    AFGANISTAN"
    FRMZELEOEX.MSF2.AddItem "302    ARABIA SAUDITA"
    FRMZELEOEX.MSF2.AddItem "303    BAHREIN"
    FRMZELEOEX.MSF2.AddItem "304    MYANMAR (EX - BIRMANIA)"
    FRMZELEOEX.MSF2.AddItem "305    BUTAN"
    FRMZELEOEX.MSF2.AddItem "306    CAMBODYA (EX - KAMPUCHE)"
    FRMZELEOEX.MSF2.AddItem "307    SRI LANKA"
    FRMZELEOEX.MSF2.AddItem "308    COREA DEMOCRATICA"
    FRMZELEOEX.MSF2.AddItem "309    COREA REPUBLICANA"
    FRMZELEOEX.MSF2.AddItem "310    CHINA"
    FRMZELEOEX.MSF2.AddItem "312    FILIPINAS"
    FRMZELEOEX.MSF2.AddItem "313    TAIWAN"
    FRMZELEOEX.MSF2.AddItem "315    INDIA"
    FRMZELEOEX.MSF2.AddItem "316    INDONESIA"
    FRMZELEOEX.MSF2.AddItem "317    IRAK"
    FRMZELEOEX.MSF2.AddItem "318    IRAN"
    FRMZELEOEX.MSF2.AddItem "319    ISRAEL"
    FRMZELEOEX.MSF2.AddItem "320    JAPON"
    FRMZELEOEX.MSF2.AddItem "321    JORDANIA"
    FRMZELEOEX.MSF2.AddItem "322    QATAR"
    FRMZELEOEX.MSF2.AddItem "323    KUWAIT"
    FRMZELEOEX.MSF2.AddItem "324    LAOS"
    FRMZELEOEX.MSF2.AddItem "325    LIBANO"
    FRMZELEOEX.MSF2.AddItem "326    MALASIA"
    FRMZELEOEX.MSF2.AddItem "327    MALDIVAS ISLAS"
    FRMZELEOEX.MSF2.AddItem "328    OMAN"
    FRMZELEOEX.MSF2.AddItem "329    MONGOLIA"
    FRMZELEOEX.MSF2.AddItem "330    NEPAL"
    FRMZELEOEX.MSF2.AddItem "331    EMIRATOS ARABES UNIDOS"
    FRMZELEOEX.MSF2.AddItem "332    PAKISTÁN"
    FRMZELEOEX.MSF2.AddItem "333    SINGAPUR"
    FRMZELEOEX.MSF2.AddItem "334    SIRIA"
    FRMZELEOEX.MSF2.AddItem "335    THAILANDIA"
    FRMZELEOEX.MSF2.AddItem "337    VIETNAM"
    FRMZELEOEX.MSF2.AddItem "341    HONG KONG"
    FRMZELEOEX.MSF2.AddItem "344    MACAO"
    FRMZELEOEX.MSF2.AddItem "345    BANGLADESH"
    FRMZELEOEX.MSF2.AddItem "346    BRUNEI"
    FRMZELEOEX.MSF2.AddItem "348    REPUBLICA DE YEMEN"
    FRMZELEOEX.MSF2.AddItem "349    ARMENIA"
    FRMZELEOEX.MSF2.AddItem "350    AZERBAIJAN"
    FRMZELEOEX.MSF2.AddItem "351    GEORGIA"
    FRMZELEOEX.MSF2.AddItem "352    KAZAJSTAN"
    FRMZELEOEX.MSF2.AddItem "353    KIRGUIZISTAN"
    FRMZELEOEX.MSF2.AddItem "354    TAYIKISTAN"
    FRMZELEOEX.MSF2.AddItem "355    TURKMENISTAN"
    FRMZELEOEX.MSF2.AddItem "356    UZBEKISTAN"
    FRMZELEOEX.MSF2.AddItem "357    TERR.AU.PALESTINOS"
    FRMZELEOEX.MSF2.AddItem "397    RESTO DE ASIA"
    FRMZELEOEX.MSF2.AddItem "398    INDET.(ASIA)"
    FRMZELEOEX.MSF2.AddItem "401    ALBANIA"
    FRMZELEOEX.MSF2.AddItem "404    ANDORRA"
    FRMZELEOEX.MSF2.AddItem "405    AUSTRIA"
    FRMZELEOEX.MSF2.AddItem "406    BELGICA"
    FRMZELEOEX.MSF2.AddItem "407    BULGARIA"
    FRMZELEOEX.MSF2.AddItem "409    DINAMARCA"
    FRMZELEOEX.MSF2.AddItem "410    ESPAÑA"
    FRMZELEOEX.MSF2.AddItem "411    FINLANDIA"
    FRMZELEOEX.MSF2.AddItem "412    FRANCIA"
    FRMZELEOEX.MSF2.AddItem "413    GRECIA"
    FRMZELEOEX.MSF2.AddItem "414    HUNGRIA"
    FRMZELEOEX.MSF2.AddItem "415    IRLANDA"
    FRMZELEOEX.MSF2.AddItem "416    ISLANDIA"
    FRMZELEOEX.MSF2.AddItem "417    ITALIA"
    FRMZELEOEX.MSF2.AddItem "418    LIECHTENSTEIN"
    FRMZELEOEX.MSF2.AddItem "419    LUXEMBURGO"
    FRMZELEOEX.MSF2.AddItem "420    MALTA"
    FRMZELEOEX.MSF2.AddItem "421    MONACO"
    FRMZELEOEX.MSF2.AddItem "422    NORUEGA"
    FRMZELEOEX.MSF2.AddItem "423    PAISES BAJOS"
    FRMZELEOEX.MSF2.AddItem "424    POLONIA"
    FRMZELEOEX.MSF2.AddItem "425    PORTUGAL"
    FRMZELEOEX.MSF2.AddItem "426    REINO UNIDO"
    FRMZELEOEX.MSF2.AddItem "427    RUMANIA"
    FRMZELEOEX.MSF2.AddItem "428    SAN MARINO"
    FRMZELEOEX.MSF2.AddItem "429    SUECIA"
    FRMZELEOEX.MSF2.AddItem "430    SUIZA"
    FRMZELEOEX.MSF2.AddItem "431    VATICANO(SANTA SEDE)"
    FRMZELEOEX.MSF2.AddItem "433    POS.BRIT.(EUROPA)"
    FRMZELEOEX.MSF2.AddItem "435    CHIPRE"
    FRMZELEOEX.MSF2.AddItem "436    TURQUIA"
    FRMZELEOEX.MSF2.AddItem "438    ALEMANIA,REP.FED."
    FRMZELEOEX.MSF2.AddItem "439    BIELORRUSIA"
    FRMZELEOEX.MSF2.AddItem "440    ESTONIA"
    FRMZELEOEX.MSF2.AddItem "441    LETONIA"
    FRMZELEOEX.MSF2.AddItem "442    LITUANIA"
    FRMZELEOEX.MSF2.AddItem "443    MOLDAVIA"
    FRMZELEOEX.MSF2.AddItem "444    RUSIA"
    FRMZELEOEX.MSF2.AddItem "445    UCRANIA"
    FRMZELEOEX.MSF2.AddItem "446    BOSNIA HERZEGOVINA"
    FRMZELEOEX.MSF2.AddItem "447    CROACIA"
    FRMZELEOEX.MSF2.AddItem "448    ESLOVAQUIA"
    FRMZELEOEX.MSF2.AddItem "449    ESLOVENIA"
    FRMZELEOEX.MSF2.AddItem "450    MACEDONIA"
    FRMZELEOEX.MSF2.AddItem "451    REP.CHECA"
    FRMZELEOEX.MSF2.AddItem "453    MONTENEGRO"
    FRMZELEOEX.MSF2.AddItem "454    SERBIA"
    FRMZELEOEX.MSF2.AddItem "497    RESTO EUROPA"
    FRMZELEOEX.MSF2.AddItem "498    INDET.(EUROPA)"
    FRMZELEOEX.MSF2.AddItem "501    AUSTRALIA"
    FRMZELEOEX.MSF2.AddItem "503    NAURU"
    FRMZELEOEX.MSF2.AddItem "504    NUEVA ZELANDIA"
    FRMZELEOEX.MSF2.AddItem "505    VANATU"
    FRMZELEOEX.MSF2.AddItem "506    SAMOA OCCIDENTAL"
    FRMZELEOEX.MSF2.AddItem "507    TERRITORIO VINCULADOS A AUSTRALIA"
    FRMZELEOEX.MSF2.AddItem "508    TERRITORIOS VINCULADOS AL R. UNIDO"
    FRMZELEOEX.MSF2.AddItem "509    TERRITORIOS VINCULADOS A FRANCIA"
    FRMZELEOEX.MSF2.AddItem "510    TER VINCULADOS A NUEVA. ZELANDA"
    FRMZELEOEX.MSF2.AddItem "511    TER. VINCULADOS A ESTADOS UNIDOS"
    FRMZELEOEX.MSF2.AddItem "512    FIJI , ISLAS"
    FRMZELEOEX.MSF2.AddItem "513    PAPUA NUEVA GUINEA"
    FRMZELEOEX.MSF2.AddItem "514    KIRIBATI , ISLAS"
    FRMZELEOEX.MSF2.AddItem "515    MICRONESIA , EST.FEDER"
    FRMZELEOEX.MSF2.AddItem "516    PALAU"
    FRMZELEOEX.MSF2.AddItem "517    TUVALU"
    FRMZELEOEX.MSF2.AddItem "518    SALOMON , ISLAS"
    FRMZELEOEX.MSF2.AddItem "519    TONGA"
    FRMZELEOEX.MSF2.AddItem "520    MARSHALL , ISLAS"
    FRMZELEOEX.MSF2.AddItem "521    MARIANAS , ISLAS"
    FRMZELEOEX.MSF2.AddItem "597    RESTO OCEANIA"
    FRMZELEOEX.MSF2.AddItem "598    INDET.(OCEANIA)"
    FRMZELEOEX.MSF2.AddItem "997    RESTO CONTINENTE"
    FRMZELEOEX.MSF2.AddItem "998    INDET.(CONTINENTE)"

End Sub
Sub CARGACOMBO1(COMBOX As comboBox)
   
    COMBOX.AddItem ""
    COMBOX.AddItem "1  Exportacion definitiva de bienes"
    COMBOX.AddItem "2  Servicios"
    COMBOX.AddItem "4  Otros"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Label17 = ""
   Screen.MousePointer = 1

End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Text10_LostFocus()
   While Len(Text2) < 8: Text2 = "0" & Text2: Wend
End Sub

Sub CARGACOMBO4()
    FRMZELEOEX.MSF2.FixedCols = 0
    FRMZELEOEX.MSF2.Cols = 1
    FRMZELEOEX.MSF1.FixedCols = 0
    FRMZELEOEX.MSF1.Cols = 1
    FRMZELEOEX.MSF2.ColAlignment(0) = 1
    FRMZELEOEX.MSF1.ColAlignment(0) = 1
    FRMZELEOEX.MSF1.Rows = 0
    FRMZELEOEX.MSF2.Rows = 0

    FRMZELEOEX.MSF2.AddItem "50000000016         URUGUAY - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000000024         PARAGUAY - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000000032         CHILE - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000000040         BOLIVIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000000059         BRASIL - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001012         BURKINA FASO - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001020         ARGELIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001039         BOTSWANA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001047         BURUNDI - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001055         CAMERUN - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001071         CENTRO AFRICANO, REP. - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001101         COSTA DE MARFIL - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001136         EGIPTO - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001152         GABON - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001160         GAMBIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001179         GHANA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001187         GUINEA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001195         GUINEA ECUATORIAL - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001209         KENIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001217         LESOTHO - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001225         REPUBLICA DE LIBERIA (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001233         LIBIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001241         MADAGASCAR - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001276         MARRUECOS - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001284         REPUBLICA DE MAURICIO - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001292         MAURITANIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001306         NIGER - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001314         NIGERIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001322         ZIMBABWE - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001330         RUANDA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001349         SENEGAL - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001357         SIERRA LEONA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001365         SOMALIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001373         REINO DE SWAZILANDIA (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001381         SUDAN - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001403         TOGO - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001411         REPUBLICA TUNECINA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001446         ZAMBIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001454         POS.BRITANICA (AFRICA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001462         POS.ESPAÑOLA (AFRICA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001470         POS.FRANCESA (AFRICA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001489         POS.PORTUGUESA (AFRICA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001497         REPUBLICA DE ANGOLA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001500         REPUBLICA DE CABO VERDE (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001519         MOZAMBIQUE - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001527         CONGO REP.POPULAR - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001535         CHAD - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001543         MALAWI - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001551         TANZANIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001586         COSTA RICA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001616         ZAIRE - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001624         BENIN - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001632         MALI - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001705         UGANDA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001713         SUDAFRICA, REP. DE - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001810         REPUBLICA DE SEYCHELLES (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001829         SANTO TOME Y PRINCIPE - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001837         NAMIBIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001845         GUINEA BISSAU - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001853         ERITREA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001861         REPUBLICA DE DJIBUTI (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001896         COMORAS - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000001985         INDETERMINADO (AFRICA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002019         BARBADOS (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002043         CANADA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002051         COLOMBIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002094         DOMINICANA, REPUBLICA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002116         EL SALVADOR - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002124         ESTADOS UNIDOS - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002132         GUATEMALA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002140         REPUBLICA COOPERATIVA DE GUYANA (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002159         HAITI - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002167         HONDURAS - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002175         JAMAICA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002183         MEXICO - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002191         NICARAGUA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002205         REPUBLICA DE PANAMA (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002213         ESTADO LIBRE ASOCIADO DE PUERTO RICO (Estado asoc. a EEUU) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002221         PERU - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002256         ANTIGUA Y BARBUDA (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002264         VENEZUELA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002272         POS.BRITANICA (AMERICA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002280         POS.DANESA (AMERICA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002299         POS.FRANCESA (AMERICA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002302         POS.PAISES BAJOS (AMERICA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002310         POS.E.E.U.U. (AMERICA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002329         SURINAME - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002337         EL COMMONWEALTH DE DOMINICA (Estado Asociado) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002345         SANTA LUCIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002353         SAN VICENTE Y LAS GRANADINAS (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002361         BELICE (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002396         CUBA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002426         ECUADOR - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002434         REPUBLICA DE TRINIDAD Y TOBAGO - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002825         BUTAN - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002841         MYANMAR (EX BIRMANIA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002876         ISRAEL - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002882         ESTADO ASOCIADO DE GRANADA (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002892         FEDERACION DE SAN CRISTOBAL (Islas Saint Kitts and Nevis) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002906         COMUNIDAD DE LAS BAHAMAS (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002914         TAILANDIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002922         INDETERMINADO (AMERICA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002930         IRAN - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000002981         ESTADO DE QATAR (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003007         REINO HACHEMITA DE JORDANIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003015         AFGANISTAN - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003023         ARABIA SAUDITA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003031         ESTADO DE BAHREIN (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003066         CAMBOYA (EX KAMPUCHEA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003074         REPUBLICA DEMOCRATICA SOCIALISTA DE SRI LANKA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003082         COREA DEMOCRATICA  - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003090         COREA REPUBLICANA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003104         CHINA REP.POPULAR - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003112         REPUBLICA DE CHIPRE (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003120         FILIPINAS - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003139         TAIWAN - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003147         GAZA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003155         INDIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003163         INDONESIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003171         IRAK - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003201         JAPON - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003236         ESTADO DE KUWAIT (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003244         LAOS - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003252         LIBANO - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003260         MALASIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003279         REPUBLICA DE MALDIVAS (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003287         SULTANATO DE OMAN - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003295         MONGOLIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003309         NEPAL - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003317         EMIRATOS ARABES UNIDOS (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003325         PAKISTAN - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003333         SINGAPUR - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003341         SIRIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003376         VIETNAM - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003392         REPUBLICA DEL YEMEN - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003414         POS.BRITANICA (HONG KONG) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003422         POS.JAPONESA (ASIA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003449         MACAO - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003457         BANGLADESH - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003503         TURQUIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003546         ITALIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003554         TURKMENISTAN - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003562         UZBEKISTAN - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003570         TERRITORIOS AUTONOMOS PALESTINOS - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003813         ISLANDIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003880         GEORGIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003899         TAYIKISTAN - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003902         AZERBAIDZHAN - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003910         BRUNEI DARUSSALAM (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003929         KAZAJSTAN - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003937         KIRGUISTAN - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000003961         INDETERMINADO (ASIA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004011         REPUBLICA DE ALBANIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004046         PRINCIPADO DEL VALLE DE ANDORRA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004054         AUSTRIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004062         BELGICA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004070         BULGARIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004097         DINAMARCA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004100         ESPAÑA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004119         FINLANDIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004127         FRANCIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004135         GRECIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004143         HUNGRIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004151         IRLANDA (EIRE) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004186         PRINCIPADO DE LIECHTENSTEIN (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004194         GRAN DUCADO DE LUXEMBURGO - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004216         PRINCIPADO DE MONACO - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004224         NORUEGA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004232         PAISES BAJOS - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004240         POLONIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004259         PORTUGAL - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004267         REINO UNIDO - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004275         RUMANIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004283         SERENISIMA REPUBLICA DE SAN MARINO (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004291         SUECIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004305         SUIZA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004313         SANTA SEDE (VATICANO) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004321         YUGOSLAVIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004364         REPUBLICA DE MALTA (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004380         ALEMANIA, REP. FED. - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004399         BIELORUSIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004402         ESTONIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004410         LETONIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004429         LITUANIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004437         MOLDOVA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004461         BOSNIA HERZEGOVINA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004496         ESLOVENIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004909         MACEDONIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004917         POS.BRITANICA (EUROPA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004984         INDETERMINADO (EUROPA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000004992         AUSTRALIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000005034         REPUBLICA DE NAURU (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000005042         NUEVA ZELANDA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000005050         REPUBLICA DE VANUATU - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000005069         SAMOA OCCIDENTAL - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000005077         POS.AUSTRALIANA (OCEANIA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000005085         POS.BRITANICA (OCEANIA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000005093         POS.FRANCESA (OCEANIA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000005107         POS.NEOCELANDESA (OCEANIA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000005115         POS.E.E.U.U. (OCEANIA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000005123         FIJI, ISLAS - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000005131         PAPUA, ISLAS - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000005166         KIRIBATI - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000005174         TUVALU - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000005182         ISLAS SALOMON - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000005190         REINO DE TONGA (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000005204         REPUBLICA DE LAS ISLAS MARSHALL (Estado independiente) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000005212         ISLAS MARIANAS - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000005905         MICRONESIA ESTADOS FED. - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000005913         PALAU - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000005980         INDETERMINADO (OCEANIA) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006014         RUSA, FEDERACION - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006022         ARMENIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006030         CROACIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006049         UCRANIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006057         CHECA, REPUBLICA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006065         ESLOVACA, REPUBLICA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006529         ANGUILA (Territorio no autónomo del Reino Unido) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006537         ARUBA (Territorio de Países Bajos) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006545         ISLAS DE COOK (Territorio autónomo asociado a Nueva Zelanda) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006553         PATAU - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006561         POLINESIA FRANCESA (Territorio de Ultramar de Francia) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006596         ANTILLAS HOLANDESAS (Territorio de Países Bajos) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006626         ASCENCION - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006634         BERMUDAS (Territorio no autónomo del Reino Unido) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006642         CAMPIONE D@ITALIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006650         COLONIA DE GIBRALTAR - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006669         GROENLANDIA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006677         GUAM (Territorio no autónomo de los EEUU) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006685         HONK KONG (Territorio de China) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006693         ISLAS AZORES - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006707         ISLAS DEL CANAL:Guernesey,Jersey,Alderney,G.Stark,L.Sark,etc - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006715         ISLAS CAIMAN (Territorio no autónomo del Reino Unido) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006723         ISLA CHRISTMAS - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006731         ISLA DE COCOS O KEELING - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006766         ISLA DE MAN (Territorio del Reino Unido) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006774         ISLA DE NORFOLK - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006782         ISLAS TURKAS Y CAICOS (Territorio no autónomo del R. Unido) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006790         ISLAS PACIFICO - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006804         ISLA DE SAN PEDRO Y MIGUELON - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006812         ISLA QESHM - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006820         ISLAS VIRGENES BRITANICAS(Territorio no autónomo de R.UNIDO) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006839         ISLAS VIRGENES DE ESTADOS UNIDOS DE AMERICA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006847         LABUAN - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006855         MADEIRA (Territorio de Portugal) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006863         MONTSERRAT (Territorio no autónomo del Reino Unido) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006871         NIUE - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006901         PITCAIRN - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006936         REGIMEN APLICABLE A LAS SA FINANCIERAS(ley 11.073 de la ROU) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006944         SANTA ELENA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006952         SAMOA AMERICANA (Territorio no autónomo de los EEUU) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006960         ARCHIPIELAGO DE SVBALBARD - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006979         TRISTAN DA CUNHA - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006987         TRIESTE (Italia) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000006995         TOKELAU - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000007002         ZONA LIBRE DE OSTRAVA (ciudad de la antigua Checoeslovaquia) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000009986         PARA PERSONAS FISICAS DE INDETERMINADO (CONTINENTE) - Persona Física"
    FRMZELEOEX.MSF2.AddItem "50000009994         PARA PERSONAS FISICAS DE OTROS PAISES - Persona Física"
    FRMZELEOEX.MSF2.AddItem "51600000016         URUGUAY - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600000024         PARAGUAY - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600000032         CHILE - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600000040         BOLIVIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600000059         BRASIL - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001012         BURKINA FASO - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001020         ARGELIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001039         BOTSWANA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001047         BURUNDI - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001055         CAMERUN - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001071         CENTRO AFRICANO, REP. - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001101         COSTA DE MARFIL - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001136         EGIPTO - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001144         ETIOPIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001152         GABON - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001160         GAMBIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001179         GHANA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001187         GUINEA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001195         GUINEA ECUATORIAL - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001209         KENIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001217         LESOTHO - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001225         REPUBLICA DE LIBERIA (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001233         LIBIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001241         MADAGASCAR - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001276         MARRUECOS - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001284         REPUBLICA DE MAURICIO - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001292         MAURITANIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001306         NIGER - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001314         NIGERIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001322         ZIMBABWE - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001330         RUANDA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001349         SENEGAL - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001357         SIERRA LEONA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001365         SOMALIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001373         REINO DE SWAZILANDIA (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001381         SUDAN - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001403         TOGO - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001411         REPUBLICA TUNECINA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001446         ZAMBIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001454         POS.BRITANICA (AFRICA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001462         POS.ESPAÑOLA (AFRICA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001470         POS.FRANCESA (AFRICA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001489         POS.PORTUGUESA (AFRICA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001497         REPUBLICA DE ANGOLA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001500         REPUBLICA DE CABO VERDE (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001519         MOZAMBIQUE - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001527         CONGO REP.POPULAR - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001535         CHAD - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001543         MALAWI - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001551         TANZANIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001586         COSTA RICA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001616         ZAIRE - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001624         BENIN - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001632         MALI - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001705         UGANDA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001713         SUDAFRICA, REP. DE - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001810         REPUBLICA DE SEYCHELLES (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001829         SANTO TOME Y PRINCIPE - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001837         NAMIBIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001845         GUINEA BISSAU - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001853         ERITREA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001861         REPUBLICA DE DJIBUTI (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001896         COMORAS - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600001985         INDETERMINADO (AFRICA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002019         BARBADOS (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002043         CANADA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002051         COLOMBIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002094         DOMINICANA, REPUBLICA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002116         EL SALVADOR - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002124         ESTADOS UNIDOS - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002132         GUATEMALA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002140         REPUBLICA COOPERATIVA DE GUYANA (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002159         HAITI - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002167         HONDURAS - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002175         JAMAICA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002183         MEXICO - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002191         NICARAGUA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002205         REPUBLICA DE PANAMA (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002213         ESTADO LIBRE ASOCIADO DE PUERTO RICO (Estado asoc. a EEUU) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002221         PERU - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002256         ANTIGUA Y BARBUDA (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002264         VENEZUELA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002272         POS.BRITANICA (AMERICA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002280         POS.DANESA (AMERICA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002299         POS.FRANCESA (AMERICA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002302         POS.PAISES BAJOS (AMERICA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002310         POS.E.E.U.U. (AMERICA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002329         SURINAME - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002337         EL COMMONWEALTH DE DOMINICA (Estado Asociado) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002345         SANTA LUCIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002353         SAN VICENTE Y LAS GRANADINAS (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002361         BELICE (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002396         CUBA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002426         ECUADOR - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002434         REPUBLICA DE TRINIDAD Y TOBAGO - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002825         BUTAN - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002841         MYANMAR (EX BIRMANIA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002876         ISRAEL - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002884         ESTADO ASOCIADO DE GRANADA (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002892         FEDERACION DE SAN CRISTOBAL (Islas Saint Kitts and Nevis) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002906         COMUNIDAD DE LAS BAHAMAS (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002914         TAILANDIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002922         INDETERMINADO (AMERICA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002930         IRAN - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600002981         ESTADO DE QATAR (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003007         REINO HACHEMITA DE JORDANIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003015         AFGANISTAN - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003023         ARABIA SAUDITA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003031         ESTADO DE BAHREIN (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003066         CAMBOYA (EX KAMPUCHEA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003074         REPUBLICA DEMOCRATICA SOCIALISTA DE SRI LANKA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003082         COREA DEMOCRATICA  - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003090         COREA REPUBLICANA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003104         CHINA REP.POPULAR - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003112         REPUBLICA DE CHIPRE (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003120         FILIPINAS - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003139         TAIWAN - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003147         GAZA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003155         INDIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003163         INDONESIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003171         IRAK - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003201         JAPON - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003236         ESTADO DE KUWAIT (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003244         LAOS - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003252         LIBANO - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003260         MALASIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003279         REPUBLICA DE MALDIVAS (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003287         SULTANATO DE OMAN - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003295         MONGOLIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003309         NEPAL - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003317         EMIRATOS ARABES UNIDOS (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003325         PAKISTAN - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003333         SINGAPUR - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003341         SIRIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003376         VIETNAM - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003392         REPUBLICA DEL YEMEN - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003414         POS.BRITANICA (HONG KONG) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003422         POS.JAPONESA (ASIA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003449         MACAO - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003457         BANGLADESH - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003503         TURQUIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003546         ITALIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003554         TURKMENISTAN - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003562         UZBEKISTAN - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003570         TERRITORIOS AUTONOMOS PALESTINOS - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003813         ISLANDIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003880         GEORGIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003899         TAYIKISTAN - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003902         AZERBAIDZHAN - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003910         BRUNEI DARUSSALAM (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003929         KAZAJSTAN - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003937         KIRGUISTAN - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600003961         INDETERMINADO (ASIA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004011         REPUBLICA DE ALBANIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004046         PRINCIPADO DEL VALLE DE ANDORRA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004054         AUSTRIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004062         BELGICA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004070         BULGARIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004097         DINAMARCA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004100         ESPAÑA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004119         FINLANDIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004127         FRANCIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004135         GRECIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004143         HUNGRIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004151         IRLANDA (EIRE) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004186         PRINCIPADO DE LIECHTENSTEIN (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004194         GRAN DUCADO DE LUXEMBURGO - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004216         PRINCIPADO DE MONACO - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004224         NORUEGA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004232         PAISES BAJOS - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004240         POLONIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004259         PORTUGAL - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004267         REINO UNIDO - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004275         RUMANIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004283         SERENISIMA REPUBLICA DE SAN MARINO (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004291         SUECIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004305         SUIZA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004313         SANTA SEDE (VATICANO) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004321         YUGOSLAVIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004364         REPUBLICA DE MALTA (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004380         ALEMANIA, REP. FED. - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004399         BIELORUSIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004402         ESTONIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004410         LETONIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004429         LITUANIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004437         MOLDOVA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004461         BOSNIA HERZEGOVINA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004496         ESLOVENIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004909         MACEDONIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004917         POS.BRITANICA (EUROPA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004984         INDETERMINADO (EUROPA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600004992         AUSTRALIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600005034         REPUBLICA DE NAURU (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600005042         NUEVA ZELANDA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600005050         REPUBLICA DE VANUATU - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600005069         SAMOA OCCIDENTAL - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600005077         POS.AUSTRALIANA (OCEANIA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600005085         POS.BRITANICA (OCEANIA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600005093         POS.FRANCESA (OCEANIA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600005107         POS.NEOCELANDESA (OCEANIA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600005115         POS.E.E.U.U. (OCEANIA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600005123         FIJI, ISLAS - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600005131         PAPUA, ISLAS - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600005166         KIRIBATI - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600005174         TUVALU - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600005182         ISLAS SALOMON - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600005190         REINO DE TONGA (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600005204         REPUBLICA DE LAS ISLAS MARSHALL (Estado independiente) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600005212         ISLAS MARIANAS - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600005905         MICRONESIA ESTADOS FEDERADOS - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600005913         PALAU - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600005980         INDETERMINADO (OCEANIA) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006014         RUSA, FEDERACION - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006022         ARMENIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006030         CROACIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006049         UCRANIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006057         CHECA, REPUBLICA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006065         ESLOVACA, REPUBLICA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006529         ANGUILA (Territorio no autónomo del Reino Unido) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006537         ARUBA (Territorio de Países Bajos) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006545         ISLAS DE COOK (Territorio autónomo asociado a Nueva Zelanda) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006553         PATAU - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006561         POLINESIA FRANCESA (Territorio de Ultramar de Francia) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006596         ANTILLAS HOLANDESAS (Territorio de Países Bajos) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006626         ASCENCION - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006634         BERMUDAS (Territorio no autónomo del Reino Unido) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006642         CAMPIONE D@ITALIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006650         COLONIA DE GIBRALTAR - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006669         GROENLANDIA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006677         GUAM (Territorio no autónomo de los EEUU) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006685         HONK KONG (Territorio de China) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006693         ISLAS AZORES - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006707         ISLAS DEL CANAL:Guernesey,Jersey,Alderney,G.Stark,L.Sark,etc - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006715         ISLAS CAIMAN (Territorio no autónomo del Reino Unido) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006723         ISLA CHRISTMAS - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006731         ISLA DE COCOS O KEELING - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006766         ISLA DE MAN (Territorio del Reino Unido) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006774         ISLA DE NORFOLK - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006782         ISLAS TURKAS Y CAICOS (Territorio no autónomo del R. Unido) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006790         ISLAS PACIFICO - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006804         ISLA DE SAN PEDRO Y MIGUELON - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006812         ISLA QESHM - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006820         ISLAS VIRGENES BRITANICAS(Territorio no autónomo de R.UNIDO) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006839         ISLAS VIRGENES DE ESTADOS UNIDOS DE AMERICA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006847         LABUAN - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006855         MADEIRA (Territorio de Portugal) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006863         MONTSERRAT (Territorio no autónomo del Reino Unido) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006871         NIUE - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006901         PITCAIRN - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006936         REGIMEN APLICABLE A LAS SA FINANCIERAS(ley 11.073 de la ROU) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006944         SANTA ELENA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006952         SAMOA AMERICANA (Territorio no autónomo de los EEUU) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006960         ARCHIPIELAGO DE SVBALBARD - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006979         TRISTAN DA CUNHA - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006987         TRIESTE (Italia) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600006995         TOKELAU - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600007002         ZONA LIBRE DE OSTRAVA (ciudad de la antigua Checoeslovaquia) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600009986         PARA PERSONAS FISICAS DE INDETERMINADO (CONTINENTE) - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "51600009994         PARA PERSONAS FISICAS DE OTROS PAISES - Otro tipo de Entidad"
    FRMZELEOEX.MSF2.AddItem "55000000018         URUGUAY - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000000026         PARAGUAY - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000000034         CHILE - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000000042         BOLIVIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000000050         BRASIL - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001014         BURKINA FASO - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001022         ARGELIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001030         BOTSWANA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001049         BURUNDI - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001057         CAMERUN - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001073         CENTRO AFRICANO, REP. - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001103         COSTA DE MARFIL - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001138         EGIPTO - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001146         ETIOPIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001154         GABON - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001162         GAMBIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001170         GHANA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001189         GUINEA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001197         GUINEA ECUATORIAL - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001200         KENIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001219         LESOTHO - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001227         REPUBLICA DE LIBERIA (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001235         LIBIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001243         MADAGASCAR - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001278         MARRUECOS - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001286         REPUBLICA DE MAURICIO - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001294         MAURITANIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001308         NIGER - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001316         NIGERIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001324         ZIMBABWE - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001332         RUANDA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001340         SENEGAL - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001359         SIERRA LEONA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001367         SOMALIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001375         REINO DE SWAZILANDIA (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001383         SUDAN - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001405         TOGO - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001413         REPUBLICA TUNECINA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001448         ZAMBIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001456         POS.BRITANICA (AFRICA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001464         POS.ESPAÑOLA (AFRICA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001472         POS.FRANCESA (AFRICA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001480         POS.PORTUGUESA (AFRICA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001499         REPUBLICA DE ANGOLA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001502         REPUBLICA DE CABO VERDE (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001510         MOZAMBIQUE - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001529         CONGO REP.POPULAR - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001537         CHAD - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001545         MALAWI - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001553         TANZANIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001588         COSTA RICA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001618         ZAIRE - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001626         BENIN - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001634         MALI - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001707         UGANDA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001715         SUDAFRICA, REP. DE - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001812         REPUBLICA DE SEYCHELLES (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001820         SANTO TOME Y PRINCIPE - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001839         NAMIBIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001847         GUINEA BISSAU - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001855         ERITREA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001863         REPUBLICA DE DJIBUTI (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001898         COMORAS - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000001987         INDETERMINADO (AFRICA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002010         BARBADOS (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002045         CANADA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002053         COLOMBIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002096         DOMINICANA, REPUBLICA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002118         EL SALVADOR - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002126         ESTADOS UNIDOS - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002134         GUATEMALA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002142         REPUBLICA COOPERATIVA DE GUYANA (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002150         HAITI - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002169         HONDURAS - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002177         JAMAICA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002185         MEXICO - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002193         NICARAGUA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002207         REPUBLICA DE PANAMA (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002215         ESTADO LIBRE ASOCIADO DE PUERTO RICO (Estado asoc. a EEUU) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002223         PERU - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002258         ANTIGUA Y BARBUDA (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002266         VENEZUELA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002274         POS.BRITANICA (AMERICA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002282         POS.DANESA (AMERICA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002290         POS.FRANCESA (AMERICA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002304         POS.PAISES BAJOS (AMERICA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002312         POS.E.E.U.U. (AMERICA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002320         SURINAME - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002339         EL COMMONWEALTH DE DOMINICA (Estado Asociado) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002347         SANTA LUCIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002355         SAN VICENTE Y LAS GRANADINAS (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002363         BELICE (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002398         CUBA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002428         ECUADOR - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002436         REPUBLICA DE TRINIDAD Y TOBAGO - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002827         BUTAN - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002843         MYANMAR (EX BIRMANIA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002878         ISRAEL - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002884         ESTADO ASOCIADO DE GRANADA (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002894         FEDERACION DE SAN CRISTOBAL (Islas Saint Kitts and Nevis) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002908         COMUNIDAD DE LAS BAHAMAS (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002916         TAILANDIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002924         INDETERMINADO (AMERICA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002932         IRAN - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000002983         ESTADO DE QATAR (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003009         REINO HACHEMITA DE JORDANIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003017         AFGANISTAN - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003025         ARABIA SAUDITA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003033         ESTADO DE BAHREIN (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003068         CAMBOYA (EX KAMPUCHEA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003076         REPUBLICA DEMOCRATICA SOCIALISTA DE SRI LANKA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003084         COREA DEMOCRATICA  - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003092         COREA REPUBLICANA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003106         CHINA REP.POPULAR - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003114         REPUBLICA DE CHIPRE (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003122         FILIPINAS - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003130         TAIWAN - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003149         GAZA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003157         INDIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003165         INDONESIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003173         IRAK - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003203         JAPON - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003238         ESTADO DE KUWAIT (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003246         LAOS - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003254         LIBANO - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003262         MALASIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003270         REPUBLICA DE MALDIVAS (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003289         SULTANATO DE OMAN - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003297         MONGOLIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003300         NEPAL - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003319         EMIRATOS ARABES UNIDOS (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003327         PAKISTAN - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003335         SINGAPUR - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003343         SIRIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003378         VIETNAM - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003394         REPUBLICA DEL YEMEN - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003416         POS.BRITANICA (HONG KONG) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003424         POS.JAPONESA (ASIA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003440         MACAO - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003459         BANGLADESH - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003505         TURQUIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003548         ITALIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003556         TURKMENISTAN - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003564         UZBEKISTAN - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003572         TERRITORIOS AUTONOMOS PALESTINOS - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003815         ISLANDIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003882         GEORGIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003890         TAYIKISTAN - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003904         AZERBAIDZHAN - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003912         BRUNEI DARUSSALAM (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003920         KAZAJSTAN - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003939         KIRGUISTAN - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000003963         INDETERMINADO (ASIA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004013         REPUBLICA DE ALBANIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004048         PRINCIPADO DEL VALLE DE ANDORRA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004056         AUSTRIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004064         BELGICA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004072         BULGARIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004099         DINAMARCA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004102         ESPAÑA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004110         FINLANDIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004129         FRANCIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004137         GRECIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004145         HUNGRIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004153         IRLANDA (EIRE) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004188         PRINCIPADO DE LIECHTENSTEIN (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004196         GRAN DUCADO DE LUXEMBURGO - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004218         PRINCIPADO DE MONACO - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004226         NORUEGA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004234         PAISES BAJOS - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004242         POLONIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004250         PORTUGAL - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004269         REINO UNIDO - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004277         RUMANIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004285         SERENISIMA REPUBLICA DE SAN MARINO (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004293         SUECIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004307         SUIZA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004315         SANTA SEDE (VATICANO) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004323         YUGOSLAVIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004366         REPUBLICA DE MALTA (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004382         ALEMANIA, REP. FED. - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004390         BIELORUSIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004404         ESTONIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004412         LETONIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004420         LITUANIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004439         MOLDOVA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004463         BOSNIA HERZEGOVINA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004498         ESLOVENIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004900         MACEDONIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004919         POS.BRITANICA (EUROPA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004986         INDETERMINADO (EUROPA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000004994         AUSTRALIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000005036         REPUBLICA DE NAURU (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000005044         NUEVA ZELANDA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000005052         REPUBLICA DE VANUATU - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000005069         SAMOA OCCIDENTAL - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000005079         POS.AUSTRALIANA (OCEANIA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000005087         POS.BRITANICA (OCEANIA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000005095         POS.FRANCESA (OCEANIA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000005109         POS.NEOCELANDESA (OCEANIA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000005117         POS.E.E.U.U. (OCEANIA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000005125         FIJI, ISLAS - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000005133         PAPUA, ISLAS - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000005168         KIRIBATI - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000005176         TUVALU - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000005184         ISLAS SALOMON - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000005192         REINO DE TONGA (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000005206         REPUBLICA DE LAS ISLAS MARSHALL (Estado independiente) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000005214         ISLAS MARIANAS - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000005907         MICRONESIA ESTADOS FED. - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000005915         PALAU - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000005982         INDETERMINADO (OCEANIA) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006016         RUSA, FEDERACION - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006024         ARMENIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006032         CROACIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006040         UCRANIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006067         ESLOVACA, REPUBLICA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006520         ANGUILA (Territorio no autónomo del Reino Unido) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006539         ARUBA (Territorio de Países Bajos) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006547         ISLAS DE COOK (Territorio autónomo asociado a Nueva Zelanda) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006555         PATAU - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006563         POLINESIA FRANCESA (Territorio de Ultramar de Francia) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006598         ANTILLAS HOLANDESAS (Territorio de Países Bajos) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006628         ASCENCION - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006636         BERMUDAS (Territorio no autónomo del Reino Unido) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006644         CAMPIONE D@ITALIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006652         COLONIA DE GIBRALTAR - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006660         GROENLANDIA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006679         GUAM (Territorio no autónomo de los EEUU) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006687         HONK KONG (Territorio de China) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006695         ISLAS AZORES - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006709         ISLAS DEL CANAL:Guernesey,Jersey,Alderney,G.Stark,L.Sark,etc - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006717         ISLAS CAIMAN (Territorio no autónomo del Reino Unido) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006725         ISLA CHRISTMAS - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006733         ISLA DE COCOS O KEELING - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006768         ISLA DE MAN (Territorio del Reino Unido) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006776         ISLA DE NORFOLK - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006784         ISLAS TURKAS Y CAICOS (Territorio no autónomo del R. Unido) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006792         ISLAS PACIFICO - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006806         ISLA DE SAN PEDRO Y MIGUELON - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006814         ISLA QESHM - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006822         ISLAS VIRGENES BRITANICAS(Territorio no autónomo de R.UNIDO) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006830         ISLAS VIRGENES DE ESTADOS UNIDOS DE AMERICA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006849         LABUAN - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006857         MADEIRA (Territorio de Portugal) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006865         MONTSERRAT (Territorio no autónomo del Reino Unido) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006873         NIUE - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006903         PITCAIRN - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006938         REGIMEN APLICABLE A LAS SA FINANCIERAS(ley 11.073 de la ROU) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006946         SANTA ELENA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006954         SAMOA AMERICANA (Territorio no autónomo de los EEUU) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006962         ARCHIPIELAGO DE SVBALBARD - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006970         TRISTAN DA CUNHA - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006989         TRIESTE (Italia) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000006997         TOKELAU - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000007004         ZONA LIBRE DE OSTRAVA (ciudad de la antigua Checoeslovaquia) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000009988         PARA PERSONAS FISICAS DE INDETERMINADO (CONTINENTE) - Persona Jurídica"
    FRMZELEOEX.MSF2.AddItem "55000009996         PARA PERSONAS FISICAS DE OTROS PAISES - Persona Jurídica"


End Sub

Sub CARGACOMBO5(COMBOX As comboBox)
    COMBOX.AddItem ""
    COMBOX.AddItem "EXW"
    COMBOX.AddItem "FCA"
    COMBOX.AddItem "FAS"
    COMBOX.AddItem "FOB"
    COMBOX.AddItem "CFR"
    COMBOX.AddItem "CIF"
    COMBOX.AddItem "CPT"
    COMBOX.AddItem "CIP"
    COMBOX.AddItem "DAF"
    COMBOX.AddItem "DES"
    COMBOX.AddItem "DEQ"
    COMBOX.AddItem "DDU"
    COMBOX.AddItem "DDP"
    COMBOX.AddItem "SVC"
    COMBOX.AddItem "DAP"
End Sub
Sub CARGACOMBO6(COMBOX As comboBox)

    COMBOX.AddItem ""
    COMBOX.AddItem "1 Español"
    COMBOX.AddItem "2 Inglés"
    COMBOX.AddItem "3 Portugués"
    
End Sub
Private Sub Label25_DblClick()
   
    Call SETULTREG&("!COMPASOC", 0)
    JJ& = 0
    For JX% = 0 To List1.ListCount - 1
      L1$ = List1.List(JX%)
      TIPCO% = Left$(TRIM$(L1$), 2)
      PVTA$ = TRIM$(Mid$(L1$, 20, 4))
      NCBTE$ = TRIM$(Mid$(L1$, 32, 8))
      '
      YY$ = REGBLAN$("!COMPASOC")
      Call REPLA(YY$, MKI$(TIPCO%), 1, 2)
      Call REPLA(YY$, PVTA$, 3, 4)
      Call REPLA(YY$, NCBTE$, 7, 8)
      JJ& = JJ& + 1
      Call GRAREG("!COMPASOC", YY$, JJ&)
      '
    Next JX%

   
10 VTB% = VENTABU%("COMPASOC")
   If VTB% < 0 Then Exit Sub
   '
   'VALIDO
   JJ& = 0
   List1.Clear
   
   For U& = 1 To ULTREG&("!COMPASOC")
         XX$ = REGLEIDO$("!COMPASOC", U&)
         TIPCO% = CVI(Mid$(XX$, 1, 2))
         DESCRI$ = TRIM$(ECOARCH$("TATICOMP", Chr$(TIPCO%)))
         PVTA$ = TRIM$(Mid$(XX$, 3, 4))
         NCBTE$ = TRIM$(Mid$(XX$, 7, 8))
         If TIPCO% > 0 Then
            If XVALO(PVTA$) < 1 Then Call COMUNI("Punto de Venta No Válido"): GoTo 10
            If XVALO(NCBTE$) < 1 Then Call COMUNI("Número de Comprobante No Válido"): GoTo 10
         End If
    Next U&
    '
    List1.Clear
    For U& = 1 To ULTREG&("!COMPASOC")
         XX$ = REGLEIDO$("!COMPASOC", U&)
         TIPCO% = CVI(Mid$(XX$, 1, 2))
         DESCRI$ = TRIM$(ECOARCH$("TATICOMP", Chr$(TIPCO%)))
         PVTA$ = TRIM$(Mid$(XX$, 3, 4))
         While Len(PVTA$) < 4: PVTA$ = "0" & PVTA$: Wend
         NCBTE$ = TRIM$(Mid$(XX$, 7, 8))
         While Len(NCBTE$) < 8: NCBTE$ = "0" & NCBTE$: Wend
         Call REPLA(TX$, TRIM$(Str$(TIPCO%)) & " " & DESCRI$, 1, 18)
         Call REPLA(TX$, PVTA$, 20, 4)
         Call REPLA(TX$, NCBTE$, 32, 8)
         List1.AddItem TX$
     Next U&
     '
     Call CIERRARCH("!COMPASOC")
     Call CIERRARCH("COMPASOC")

      
End Sub

Private Sub LIST1_Click()
   If List1.ListCount > 0 Then
      Call Label25_DblClick
   End If
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call Command4_Click
   End If
End Sub


Private Sub Text2_DblClick()
   Call Command1_Click
End Sub


Private Sub Text3_DBLCLICK()
   Call Command6_Click
End Sub


