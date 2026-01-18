VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ASIGLOPRM 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Situación de Stocks - Considera Tiempos de Secado."
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   14505
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
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
      Height          =   450
      Left            =   13680
      Picture         =   "ASIGLOPRM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Acepta Ventana  -  Continuar"
      Top             =   4750
      Width           =   810
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   12720
      Picture         =   "ASIGLOPRM.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Refrescar las Reservas"
      Top             =   4750
      Width           =   885
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7800
      OleObjectBlob   =   "ASIGLOPRM.frx":0454
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   4440
      OleObjectBlob   =   "ASIGLOPRM.frx":0688
      TabIndex        =   0
      Top             =   -120
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   4455
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Seleccione con Doble Click la Fila Con el Item a Seleccionar Reserva"
      Top             =   240
      Width           =   14475
      _ExtentX        =   25532
      _ExtentY        =   7858
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle de Factura:"
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
      Index           =   0
      Left            =   30
      TabIndex        =   3
      Top             =   0
      Width           =   2745
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuArchivosegundo 
      Caption         =   "Archivo"
      Begin VB.Menu mnurefrescar 
         Caption         =   "Refrescar "
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuaceptar 
         Caption         =   "Aceptar"
      End
   End
   Begin VB.Menu mnuacc 
      Caption         =   "Acceso Directos"
      Begin VB.Menu mnuFikardex 
         Caption         =   "Ficha Kardex Código Activo"
      End
   End
End
Attribute VB_Name = "ASIGLOPRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GLOBALFRM_strContenidoCelda$
Dim GLOBALFRM_REG&
Dim GLOBALFRM_COL&




Private Sub Command1_Click()
  Command1.Enabled = False
  Call ASIGNALOTES_PREMOTEC("!CARDEFAC", RESUL$)
  Command1.Enabled = True
End Sub

Private Sub Command6_Click()
  Command6.Enabled = False
  VALIDA = "APRO"
  For i& = 1 To MSF.Rows - 1
       LOTE$ = MSF.TextMatrix(i&, 6)
       If LOTE$ = "" Then
            Call MENSERR(24, "Todos los Articulos Deben Tener un Lote Asignado")
            Exit Sub
       End If
   Next i&
   
  'RECORRO LA GRILLA LEO EL LOTE Y LO VUELVA A RECORRER ACUMULANDO PARA VER SI NO SUPERA EL SALDO DEL LOTE
  ACUMULASTRINGLOTE$ = ""
   For i& = 1 To MSF.Rows - 1
     CI1% = XVALO(MSF.TextMatrix(i&, 1))
     LOTE$ = MSF.TextMatrix(i&, 6)
     If InStr(ACUMULASTRINGLOTE$, LOTE$) < 1 Then ' TODAVIA NO LO CONTROLÓ (SI YA LO LEYO NO INGRESA)
          ACUMULOSUMA = SUMA_CANT_SEGUNDATO(ASIGLOPRM.MSF, LOTE$, 6, 7)
          a1 = SaldoLote(CI1%, LOTE$)
          If ACUMULOSUMA <= SaldoLote(CI1%, LOTE$) Then
            ACUMULASTRINGLOTE$ = ACUMULASTRINGLOTE$ + LOTE$ + ";"
          Else
            Call COMUNI("Inconsistencia Entre Reservas y Saldos de O.Fabriación Generadas\Controlar !!!")
            Exit Sub
          End If
      End If
    Next i&
    '
    
    Hide
    Command6.Enabled = True

End Sub

Private Sub Form_Load()
    Call CREAR_TABLA_RESERPED

    Campos$ = "CODIGOINTERNO/F0;Nro.Ped./F8;Código/A20;Descripción/A32;Cant.Ped/f10.1;Lote/A12;Cant.Res./F10.1"
    Call CARGA_ENCABEZADO(ASIGLOPRM.MSF, Campos$, ASIGLOPRM)
'    'CAMPOS$ = "Fecha Reserva/d8;Cant.Reservada/F14.1;Cant.Cons./F14.1;Cant.Pend/f14.1;Ord.Fabricación/A15"
'    CAMPOS$ = "Cant.Reservada/F14.1;Fabricación/A15;F.Secado/D8;Alta/A0"
'    Call CARGA_ENCABEZADO(RESEPEDI.MSF1, CAMPOS$, RESEPEDI)
End Sub



Private Sub mnuaceptar_Click()
   Command6_Click
End Sub

Private Sub mnurefrescar_Click()
   Command1_Click
End Sub

Private Sub MSF_DblClick()
  REG& = MSF.MouseRow
  RESEPEDI.Command1.Enabled = False
  RESEPEDI.TXTNROPED = MSF.TextMatrix(REG&, 2)
    Campos$ = "CODIGOINTERNO/F0;Código/A20;Descripción/A48;Cant.Pend/f11.1"
  RESEPEDI.MSF.Rows = 2
  RESEPEDI.MSF.TextMatrix(1, 1) = MSF.TextMatrix(REG&, 1)
  RESEPEDI.MSF.TextMatrix(1, 2) = MSF.TextMatrix(REG&, 3)
  RESEPEDI.MSF.TextMatrix(1, 3) = MSF.TextMatrix(REG&, 4)
  RESEPEDI.MSF.TextMatrix(1, 4) = MSF.TextMatrix(REG&, 5)
  RESEPEDI.Show 1

End Sub







Private Sub TXTNORFA_Change()
    Command2.Enabled = False
    If TRIM$(TXTNORFA) <> "" Then
       Command2.Enabled = True
    End If
End Sub

