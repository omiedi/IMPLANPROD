VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form STDECRTE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saldos de Materiales Transferidos al Depòsito de Corte"
   ClientHeight    =   7890
   ClientLeft      =   1050
   ClientTop       =   855
   ClientWidth     =   15720
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   15720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   14280
      TabIndex        =   4
      Top             =   4440
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSF0 
      Height          =   4125
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15705
      _ExtentX        =   27702
      _ExtentY        =   7276
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
      OleObjectBlob   =   "STDECRTE.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   3240
      OleObjectBlob   =   "STDECRTE.frx":0234
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2445
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   15465
      _ExtentX        =   27279
      _ExtentY        =   4313
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      Redraw          =   -1  'True
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Consumos  (O.F.Activa)"
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
      Index           =   7
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "STDECRTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
'      Campos$ = "Código/A16;Descripción/A32;Lote/A16;Peso/F10.2;Neces.M./F12.2;Bobina M./f12.2;NADA/a0"
'      Call CARGA_ENCABEZADO(Me.MSF0, Campos$, Me)
'      Call CARGA_ENCABEZADO(MSF1, "CODIGO/A20;Descripción/A40;LOTE/A16;Saldo.Transf.Mts./f18.1;CI/f6;Op.N°/F6;Kg.Transf.(Orig.)/F16.2", Me)
'      lbldepogeneral = CONTROL$("", "DEPOCEN")
'       CI1% = XVALO(!COD_INTE)
'       MSF1.TextMatrix(REG&, 1) = SAFETEXT$(!COD_EXTE)
'       MSF1.TextMatrix(REG&, 2) = SAFETEXT$(!Descrip)
'        CANTMETROSBOBINA = XVALO(!Resultado)
'        PesoPorMetro# = CoeficientePesoPorMetroSegunLote#(LOTE$)
'        PESOSALDOBOBINA = CANTMETROSBOBINA * PesoPorMetro#
'       MSF1.TextMatrix(REG&, 3) = ""
'       MSF1.TextMatrix(REG&, 4) = FORMATNUM$(CANTMETROSBOBINA, "F10.1")
'       MSF1.TextMatrix(REG&, 5) = XVALO(!COD_INTE)
'       MSF1.TextMatrix(REG&, 6) = 0
'       MSF1.TextMatrix(REG&, 7) = FORMATNUM$(XVALO(PESOSALDOBOBINA), "F10.2")
'       If AGREGARCODIGORANGO% > 0 Then
'         CIRANGO% = CODIGORANGO&(CI1%)
'         If CIRANGO% < 1 Then
'           CODIGOX$ = "========"
'           DESCRIXO$ = "======================"
'         Else
'           CODIGOX$ = CODEXT$(CIRANGO%)
'           DESCRIXO$ = DESCRIT0$(CIRANGO%)
'         End If
'         MSF1.TextMatrix(REG&, 8) = CODIGOX$
'         MSF1.TextMatrix(REG&, 9) = DESCRIXO$

End Sub



