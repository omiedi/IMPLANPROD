VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form MOVISTK07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stocks - Movimientos"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5670
   Icon            =   "MOVISTK07.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   1
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   30
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   29
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command8 
      Enabled         =   0   'False
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
      Left            =   945
      Picture         =   "MOVISTK07.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Remito de Traslado"
      Top             =   5460
      Width           =   560
   End
   Begin VB.CommandButton Command10 
      Enabled         =   0   'False
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
      Left            =   105
      Picture         =   "MOVISTK07.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Lista de Empaque"
      Top             =   5460
      Width           =   560
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   5370
      Left            =   4800
      ScaleHeight     =   5340
      ScaleWidth      =   1035
      TabIndex        =   20
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command2 
         Caption         =   "Depo."
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
         Left            =   120
         Picture         =   "MOVISTK07.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Definición de Depósitos y Sucursales"
         Top             =   3840
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   32
         Top             =   4560
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Clear"
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
         Left            =   120
         Picture         =   "MOVISTK07.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Depurar Movimientos Viejos de B.Datos Movimientos de Stock"
         Top             =   2400
         Width           =   650
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Direct"
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
         Left            =   105
         Picture         =   "MOVISTK07.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1680
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
         Left            =   105
         Picture         =   "MOVISTK07.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   200
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Help"
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
         Left            =   105
         Picture         =   "MOVISTK07.frx":1386
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   890
         Width           =   650
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Setup"
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
         Left            =   105
         Picture         =   "MOVISTK07.frx":1690
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3145
         Width           =   650
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "MOVIMIENTOS DE STOCK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4600
      Left            =   120
      TabIndex        =   4
      Top             =   105
      Width           =   4530
      Begin VB.Frame Frame8 
         BackColor       =   &H00CDDADA&
         Caption         =   "EXPEDICIÓN Y DESPACHO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   210
         TabIndex        =   7
         Top             =   1785
         Width           =   4110
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   525
            ScaleHeight     =   615
            ScaleWidth      =   3120
            TabIndex        =   12
            Top             =   405
            Width           =   3150
            Begin VB.CommandButton Command13 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   2590
               Picture         =   "MOVISTK07.frx":1AD2
               Style           =   1  'Graphical
               TabIndex        =   27
               ToolTipText     =   "Re-Imprimir Materiales en Consignacion"
               Top             =   0
               Width           =   540
            End
            Begin VB.CommandButton Command23 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   0
               Picture         =   "MOVISTK07.frx":1DDC
               Style           =   1  'Graphical
               TabIndex        =   2
               ToolTipText     =   "Remito de Ventas - Clientes"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command22 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   2020
               Picture         =   "MOVISTK07.frx":20E6
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "Remito en Consignacion - Proveedores de Servicios"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Comprobantes de  Despacho"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   19
               Top             =   105
               Width           =   1215
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00CDDADA&
         Caption         =   "FÁBRICA - AJUSTES - TRANSFERENCIAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   210
         TabIndex        =   6
         Top             =   3150
         Width           =   4110
         Begin VB.PictureBox Picture10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   210
            ScaleHeight     =   615
            ScaleWidth      =   3660
            TabIndex        =   13
            Top             =   420
            Width           =   3690
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
               Height          =   645
               Left            =   3120
               Picture         =   "MOVISTK07.frx":23F0
               Style           =   1  'Graphical
               TabIndex        =   17
               ToolTipText     =   "Descartes de Stock"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command24 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   560
               Picture         =   "MOVISTK07.frx":26FA
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Transferencias de Stock"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command25 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   0
               Picture         =   "MOVISTK07.frx":2A04
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "Ajuste de Existencias"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command21 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   2560
               Picture         =   "MOVISTK07.frx":2B4E
               Style           =   1  'Graphical
               TabIndex        =   3
               ToolTipText     =   "Vale de Consumo de Materiales"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Movimientos     Varios"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   1260
               TabIndex        =   14
               Top             =   105
               Width           =   1110
            End
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00CDDADA&
         Caption         =   "ENTRADAS A DEPÓSITO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   210
         TabIndex        =   5
         Top             =   420
         Width           =   4110
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   2205
            ScaleHeight     =   615
            ScaleWidth      =   1665
            TabIndex        =   10
            Top             =   420
            Width           =   1695
            Begin VB.CommandButton Command18 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   0
               Picture         =   "MOVISTK07.frx":2C98
               Style           =   1  'Graphical
               TabIndex        =   1
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Propia Producción"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   11
               Top             =   105
               Width           =   1110
            End
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   210
            ScaleHeight     =   615
            ScaleWidth      =   1770
            TabIndex        =   8
            Top             =   420
            Width           =   1800
            Begin VB.CommandButton Command28 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   0
               Picture         =   "MOVISTK07.frx":2FA2
               Style           =   1  'Graphical
               TabIndex        =   0
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Recepción Proveedores"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   9
               Top             =   105
               Width           =   1110
            End
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "MOVISTK07.frx":32AC
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "MOVISTK07.frx":34E0
      TabIndex        =   31
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
End
Attribute VB_Name = "MOVISTK07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    Call CONECRUN("LISEMPAQ", "Generación de Listas de Empaque")
    Command10.Enabled = True
End Sub


Private Sub Command13_Click()
     Command17.Enabled = False
     '
     FF% = FECHANUM("01/06/11")
     If InStr(1, EMPREAC$, "FERVI") > 0 Then
        Call CARGALISEL("PRINTRCO", "PDOC_TBL", "NUMEBUSQUE/A10;REFERENCIA/A64", "TIPODOCU='Remito de Traslado' AND REFERENCIA LIKE '%Remito de Traslado%' AND CodiEmpr=" & CodiEmp%, "DISTINCT")
     Else
        Call CARGALISEL("PRINTRCO", "PDOC_TBL", "NUMEBUSQUE/A10;REFERENCIA/A64", "TIPODOCU='Remito de Traslado' AND REFERENCIA LIKE '%Remito de Traslado%' AND CodiEmpr=" & CodiEmp% & " AND FECHAYHORA>'" & FETEXCOR1$(FF%) & "'", "DISTINCT")
     End If
     '
5    Call SELECHO("PRINTRCO/Remitos en Consignacion")
     RESUL& = XVALO(VALACT1$("PRINTRCO"))
     If RESUL& > 0 Then
        CMO$ = TRIM$(Control$("", "CODRECON"))
        If CMO$ = "" Then CMO$ = "RC"
        REMI$ = CMO$ + "-" + TRIM$(Str$(RESUL&))
        Call MUESTRADOC(0, "", "Remito de Traslado", Str$(RESUL&), REMI$)
        GoTo 5
     End If
     '
     Command17.Enabled = True
     Exit Sub
     '
     '
     Screen.MousePointer = 11
     CMO$ = TRIM$(Control$("", "CODREFAC"))
     If Len(CMO$) <> 2 Then
        CMO$ = "RT"
     End If
     '
     Call CARSELMESA
     '
     Screen.MousePointer = 1
    
     Call ABRESTOCKS
     
'     CONDI$ = "Nremi_consig <> '' AND CODIEMPR = " & CodiEmp% & " order by Fecha_Entreg asc, Nremi_consig asc"
'     '
'     Call COPYSTRU("SELDOCAP", "SELREMIT")
'     Call CIERRARCH("!SELREMIT")
'     Call CARGALISEL("!SELREMIT", "Maconsig", "Nremi_Consig/A16;Fecha_Entreg/D8;raso_prov/A24", CONDI$, "DISTINCT")
     CONDI$ = "CODIMOVI='RC' ORDER BY FECHMOV, DOPRICOR"
     '
     Call COPYSTRU("SELDOCAP", "SELREMIT")
     Call CIERRARCH("!SELREMIT")
     Call CARGALISEL("!SELREMIT", "MOVISTO+PROVED12/MOVISTO.NUME_PROV=PROVED12.NUME_CLI", "MOVISTO.DOPRICOR/A12;MOVISTO.FECHMOV/D8;PROVED12.RASO_CLI/A24", CONDI$, "DISTINCT")
     '
     If ULTREG&("!SELREMIT") < 1 Then
        Call MENSERR(24, "No se Encontraron Remitos en el Rango Seleccionado.")
        GoTo 99
     End If

15   Call SELECHO("!SELREMIT")
     '
     Screen.MousePointer = 1
     '
     'MODIFICADO POR QUE ANTES SOLO GUARDABA EL NUMERO
     PREFIJO$ = ""
     If InStr(VALACT1$("!SELREMIT"), "RC-") < 1 Then PREFIJO$ = "RC-"
     NUDOCU$ = PREFIJO$ + TRIM$(VALACT1$("!SELREMIT")) + " "
     '
     If VALACT1$("!SELREMIT") = "" Then GoTo 99
     TIDOCUM$ = "Remito de Traslado"
      
     Screen.MousePointer = 11
      '
     For UI& = ULTREG&("PDOCLST") To 1 Step -1
        XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
        '\\\OT-08-0646-OD-20/08/08///
         If InStr(XX$, TIDOCUM$) > 4 Then
            If InStr(XX$, NUDOCU$) > 4 Then
                PPXX% = InStr(XX$, "Nro.")
                If PPXX% > 0 Then
                 XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
                End If
                If InStr(XX$, NUDOCU$) > 4 Then
                 DOCUNU& = CVS(Left$(XX$, 4))
                 GoTo 50
                End If
            End If
         End If
     Next UI&
     Screen.MousePointer = 1
40   Call COMUNI("No se Encontró el Documento Impreso")
     GoTo 15
     '
50   Call MUESTRADOC(DOCUNU&)
     GoTo 15
     '
99   Command17.Enabled = True
     '
End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
    '
    If CLACONTRA% = 0 Then GoTo 99
    
    If DERACCE%("DESCAMOV", "Descarga de Movimientos de Stock") < 2 Then
       GoTo 99
    End If
    '
'    If USNBR% > 100 Then
'       Call COMUNI("Este Proceso no Puede ser Ejecutado por el Supervisor.\   \Como Supervisor Asigne Derechos a un Usuario Común \y Remuevalos una Vez Completado el Mismo.")
'    End If
    '
10  OPX% = COMALTER%("Opciones de Depuración\\Depuración de Movimientos\Eliminación de Movimiento Viejos")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    If OPX% = 1 Then
       SOLONOTRAZABLES% = 0
       OPX2% = COMALTER%("Opciones de Depuración\\Depurar Solo ´NO´ Trazables\Depurar Todos")
       If OPX2% < 1 Or OPX2% > 2 Then GoTo 10
       If OPX2% = 1 Then SOLONOTRAZABLES% = 1
    End If
    '
    If OPX% = 2 Then GoTo 40
'    If OPX% = 3 Then GoTo 50
    '
    HOII% = HOY(HOS$)
13  FECHINVE$ = InputBox$("Ingrese la Fecha de Corte", "Depuración de los Movimientos de Stocks")
    If TRIM$(FECHINVE$) = "" Then GoTo 99
    FF% = FECHANUM(FECHINVE$)
    If FF% < 1 Or FF% >= HOY(HOS$) Then
       Call MENSERR(24, "Fecha no Válida")
       GoTo 13
    End If
    FEX = FF%
    FECHINVE$ = FECHATEX$(FEX)
    FECHINVE1$ = FETEXCOR1$(FF%)
    '
    If COMALTER%("Está ABSOLUTAMENTE SEGURO de Depurar\los Movimientos de Stock Anteriores al " + FECHINVE$ + "\\No, Conservarlos\Si, Depurar") <> 2 Then
       GoTo 99
    End If
    '
    Screen.MousePointer = 11
    '
    ' CREACION DE LA TABLA MOVIVIE
    Call COMUNI("FLEXOFT Debe Ahora Crear la Tabla de Respaldo.\  \Este Proceso Puede Demandar Varios Minutos\y NO DEBE INTERRUMPIRSE Aunque el Equipo Parezca Haberse 'Colgado' !!!")
    Screen.MousePointer = 11
    TIMUT = FLEXCONN.CommandTimeout
    FLEXCONN.CommandTimeout = 999
    Call CREACAMPO("STOCKS", "MOVISTO", "Marborra", 3, 0, 1)
    Call COPIA_STRUC_MOVISTO
    FLEXCONN.CommandTimeout = TIMUT
    Screen.MousePointer = 1
    Call COMUNI("Tabla de Respaldo Creada OK.")
    '
    FEX = FF%: FEC1$ = FETEXCOR1$(DIPOST(FEX))
    CONDI$ = "FechMov < '" & FEC1$ & "' AND CodiMovi <> 'KU' "
    Fin& = CantRegistros("MOVISTO", CONDI$)
    If Fin& < 1 Then
        Call MENSERR(24, "No Existen Movimientos Anteriores\a la Fecha: " + FECHATEX$(FF%) + ".")
        Screen.MousePointer = 1
        Command14.Enabled = True
        Exit Sub
    End If
    '
    SQLX$ = " SELECT * FROM MOVISTO WHERE " & CONDI$
    Dim MOVIAC As ADODB.Recordset
    Set MOVIAC = New ADODB.Recordset
25  On Error Resume Next
    MOVIAC.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    '
    SQLY$ = " SELECT * FROM MOVIVIE"
    Dim MOVIVI As ADODB.Recordset
    Set MOVIVI = New ADODB.Recordset
35  On Error Resume Next
    MOVIVI.Open (SQLY$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 35
    End If
    On Error GoTo 0
    '
    ' COPIA LOS REGISTROS A MOVIVIE
    Screen.MousePointer = 11
    U& = 0
    With MOVIAC
      Do While Not .EOF
         U& = U& + 1
         Call TRACE(20, U&, Fin&)
         '
         If SOLONOTRAZABLES% > 0 Then
            CI1% = XVALO(!cod_inte)
            If ESTRAZABLE%(CI1%) > 0 Then GoTo 29
         End If
         
         For KKI% = 1 To .Fields.Count
            NOCAM$ = UCase$(.Fields(KKI% - 1).Name)
            VALAC = .Fields(KKI% - 1).Value
            '
            With MOVIVI
               If KKI% = 1 Then
                  .AddNew
               End If
               For KKJ% = 1 To .Fields.Count
                  If UCase$(.Fields(KKJ% - 1).Name) = NOCAM$ Then
                     .Fields(KKJ% - 1).Value = VALAC
                  End If
               Next KKJ%
            End With
         Next KKI%
         MOVIVI.Update
         '
         !MARBORRA = 1
         .Update
29    .MoveNext

      Loop
    End With
    '
     MOVIAC.Close: Set MOVIAC = Nothing
    MOVIVI.Close: Set MOVIVI = Nothing
    '
    GoTo 10
    
     'BORRA LOS REGISTROS MARCADOS CON MARBORRA Y LOS DE ACUMULADOS
40   TIMEANT% = FLEXCONN.CommandTimeout
     FLEXCONN.CommandTimeout = 999
     Call BORRAREGISTROS("MOVISTO", "Marborra = 1", RAFE&)
     Call BORRAREGISTROS("MOVISTO", "CodiMovi = 'KU'", RAFE&)
     FLEXCONN.CommandTimeout = TIMEANT%
     
'
'40  FECHINVE$ = InputBox$("Ingrese la Fecha de Corte", "Depuración de los Movimientos de Stocks")
'    If TRIM$(FECHINVE$) = "" Then GoTo 99
'    FF% = FECHANUM(FECHINVE$)
'    If FF% < 1 Or FF% >= HOY(HOS$) Then
'       Call MENSERR(24, "Fecha no Válida")
'       GoTo 13
'    End If
'    FEX = FF%
'    FFFF% = Int(FEX)
'    HASTAX = FETEXCOR1$(DIPOST(FEX))
'    TIMEANT% = FLEXCONN.CommandTimeout
'    FLEXCONN.CommandTimeout = 999
'45  FFFF% = DIANTE(FFFF%, 100)
'    If FFFF% > 0 Then
'      If CantRegistros("MOVISTO", "MARBORRA = 1", "NOMESS") > 0 Then
'        CONDI8$ = " Marborra = 1 AND FechMov >= '" & FETEXCOR1$(FFFF%) & "'  AND FechMov < '" & HASTAX & "'"
'        CONDI9$ = " CodiMovi = 'KU' AND FechMov >= '" & FETEXCOR1$(FFFF%) & "'  AND FechMov < '" & HASTAX & "'"
''        Call BORRAREGISTROS("MOVISTO", "Marborra = 1", RAFE&)
''        Call BORRAREGISTROS("MOVISTO", "CodiMovi = 'KU'", RAFE&)
'         Call BORRAREGISTROS("MOVISTO", CONDI8$, RAFE&)
'         Call BORRAREGISTROS("MOVISTO", CONDI9$, RAFE&)
'      Else
'        MsgBox "No paso cantregistro"
'      End If
'      GoTo 45
'      MsgBox "Paso"
'    End If
'    '
'    FLEXCONN.CommandTimeout = TIMEANT%
'   GoTo 10
    
    ' GRABA LOS REGISTROS CON LOS TOTALES
50  FECHINVE$ = InputBox$("Ingrese la Fecha de Corte", "Depuración de los Movimientos de Stocks", FECHINVE$)
    If TRIM$(FECHINVE$) = "" Then GoTo 99
    FF% = FECHANUM(FECHINVE$)
    If FF% < 1 Or FF% >= HOY(HOS$) Then
       Call MENSERR(24, "Fecha no Válida")
       GoTo 13
    End If
    '
    FEX = FF%: FEC1$ = FETEXCOR1$(DIPOST(FEX))
    CONDI$ = "FechMov < '" & FEC1$ & "' AND CodiMovi <> 'KU' "
    SQLX$ = "SELECT Cod_Inte, DepoMovi, SUM(CantIngre) as SUIN,SUM(CantSalid) AS SUSAL "
    SQLX$ = SQLX$ + " FROM MOVIVIE WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    SQLX$ = SQLX$ + " GROUP BY Cod_Inte, DepoMovi"
    '
    TIMUT = FLEXCONN.CommandTimeout
    Set GENIVITMP = READSET(SQLX$, 999)        ' PROVISORIO AL 29 MAR 2011 POR TIMEOUT DE TANIT
    FLEXCONN.CommandTimeout = TIMUT
    '
    SQLX$ = " SELECT * FROM MOVISTO WHERE CODIMOVI = 'KU'"
    Set MOVIAC = New ADODB.Recordset
55  On Error Resume Next
    MOVIAC.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 55
    End If
    On Error GoTo 0
    '
    With GENIVITMP
      CONTI& = 0
      Do While .EOF = False
         CIXI% = XVALO(!cod_inte)
         DEPX% = XVALO(!DepoMovi)
         CANSA = XVALO(!SUIN) - XVALO(!SUSAL)
         '
         Call TRACE(20, CONTI&, TOT&)
         If Abs(CANSA) > 0.00005 Then
           With MOVIAC
             .AddNew
             !FechMov = CVDAT(FF%)
             !cod_inte = CIXI%
             !Cod_Exte = CODEXT$(CIXI%)
             !Descrip = Left$(DESCRIT$(CIXI%), 64)
             !DepoMovi = DEPX
             If CANSA < 0 Then
                 !CANTSALID = Abs(CANSA)
                 !CANTINGRE = 0
               Else
                 !CANTINGRE = CANSA
                 !CANTSALID = 0
             End If
             !CODIMOVI = "KU" 'MOVIMIENTO ACUMULADOs
             !ReferCor = "Acumulados al " & CVDAT(FF%) & " "
             !idenlote = ""
             !dopricor = "QM" + FECHATEX$(FF%)
             !DOPRILAR = "QM" + FECHATEX$(FF%)
             !DoSecCor = "QM" + FECHATEX$(FF%)
             !DOSECLAR = "QM" + FECHATEX$(FF%)
                NORIT& = NORIT& + 1
             !NuOrItem = NORIT& Mod 10000
             !NUPEDCLI = 0
             !VALOUNIT = 0
             !MoneVal = 0 '"" cambie a 0 xq es campo integer
             !NUME_CLIE = 0
             !Nume_Prov = 0
             !FeReMovi = Now
             !NumUsuar = USNBR%
             .Update
           End With
         End If
      .MoveNext
      Loop
    End With
    '
    GENIVITMP.Close
    Set GENIVITMP = Nothing
    '
    'FLEXCONN.CommandTimeout = TIMEANT%
    Screen.MousePointer = 1
    Call COMUNI("Proceso Completo")
    '
99  Screen.MousePointer = 1
    Command14.Enabled = True
    '
End Sub

Sub Command17_Click()
    Command17.Enabled = False
    If DERACCE%("DESCRAP", "Descartes de Stock") >= 2 Then
      Call DESCRAP
    End If
99  Command17.Enabled = True
End Sub

Sub Command18_Click()
    Command18.Enabled = False
    Call RECPROP
    Command18.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   If CLACONTRA% > 0 Then
     Call CARTADEPOSI
   End If
99 Command2.Enabled = True
End Sub

Sub Command21_Click()
    Command21.Enabled = False
2   OPX% = COMALTER%("Opciones:\\Generacion del Vale de Materiales\Re-impresión de Vale")
    EPAC$ = TRIM$(UCase$(EMPREAC$))

    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    If OPX% = 1 Then
        If DERACCE%("VALEMAT", "Vale de Consumo de Materiales") >= 2 Then
          If InStr(EPAC$, "ENVAPLAST") > 0 Then '
           Call VALEMAT_CONTEPLAST
          Else
           Call VALEMAT
          End If
        End If
        GoTo 2
    Else
        Call CARGALISEL("PRIVALMA", "PDOC_TBL", "NUMEBUSQUE/A10;REFERENCIA/A64", "TIPODOCU='Vale de Consumo de Materiales'", "DISTINCT")
'        '
5       Call SELECHO("PRIVALMA/Vale de Consumo de Materiales")
        RESUL& = XVALO(VALACT1$("PRIVALMA"))
        If RESUL& > 0 Then
           CMO$ = TRIM$(Control$("", "CODCONSU"))
           If CMO$ = "" Then CMO$ = "VM"
           Call MUESTRADOC(0, "", "Vale de Consumo de Materiales", Str$(RESUL&), CMO$ & "." & CStr(RESUL&))
           GoTo 5
        End If
        GoTo 2
    End If
99  Command21.Enabled = True
End Sub

Private Sub Command22_Click()
    'Call NUENTRECONSI(NPROVE%)
    Command22.Enabled = False
10  OPX% = COMALTER%("Consignación\\Por Proveedor\Por Proveedor y Orden de Compra\Cancelar")
    If OPX% < 1 Or OPX% > 2 Then
      GoTo 99
    End If
    '
    If OPX% = 1 Then
       'Por Proveedor
       Call NUENTRECONSI(0)
       GoTo 10
       
    ElseIf OPX% = 2 Then
        'Por Proveedor y Orden de Compra
20      EBOL$ = OBJPLA$("SELPROVE", "")
        NPRO% = CVI(EBOL$)
        If NPRO% < 1 Then GoTo 10
        '
        'VALIDACION QUE TENGA ORDENES DE COMPRA PENDIENTES DE CONSIGNACION.
        CONDI$ = "Cant_Asig > Cant_Entreg AND Cant_Asig >= 0.05 AND Npro_Ocom = " & NPRO%
        If CantRegistros&("MACONSIG", CONDI$, "NOMESS") < 1 Then
           Call COMUNI("Este Proveedor No Posee Consignaciones Pendientes")
           GoTo 20
        End If
        '
        'PRESENTA LISTA DE ORDENES DE COMPRA CON MATERIALES PENDIENTES DE CONSIGNACION.
        'INNER JOIN CON OCOMENCA PARA QUE NO PRESENTE CONSIGNACIONES DE O.C. ANULADAS.
        CONDI$ = "MACONSIG.Cant_Asig > MACONSIG.Cant_Entreg AND MACONSIG.Cant_Asig >= 0.05 AND MACONSIG.Npro_Ocom = " & NPRO% & " AND OCOMENCA.Stat_Ocom < 8 "
50      Call CARGALISEL("ORCOXPROV", "MACONSIG+OCOMENCA/CAST(SUBSTRING(MACONSIG.NRO_OCOM,4,6)AS INT)=OCOMENCA.Nume_Ocom", "MACONSIG.Nro_Ocom/A12;MACONSIG.Npro_Ocom/F6+proved12/A32;MACONSIG.Femi_Ocom/D8", CONDI$, "NOMESS/DISTINCT/ORDER BY MACONSIG.NPRO_OCOM")
        Call SELECHO("ORCOXPROV/Ordenes de Compras con Consignaciones Pendientes")
        NROCOMX$ = TRIM$(VALACT1$("ORCOXPROV"))
        If NROCOMX$ = "" Then GoTo 20
        Call NUENTRECONSI(NPRO%, NROCOMX$)
        GoTo 50
        '
    End If
99  Command22.Enabled = True
End Sub

Private Sub Command23_Click()
    Command23.Enabled = False
    MODU$ = "REMITO07"
    Call CONECRUN(MODU$, "Emision de Remito de Facturacion")
    Command23.Enabled = True
End Sub

 Sub Command24_Click()
   Command24.Enabled = False
   If DERACCE%("TRASTOCK", "Transferencias de Stock") >= 2 Then
      TRANSTK07.Show 1
   End If
   Command24.Enabled = True
End Sub

Sub Command25_Click()
    Command25.Enabled = False
    If DERACCE%("AJUSTOCK", "Ajustes de Stock") >= 2 Then
       PREPA% = 0
       FORMAJU% = 0
       AJUSTK07.Show 1
    End If
    Command25.Enabled = True
End Sub

Sub Command26_Click()
   If DERACCE%("STPSTOCK", "Configuración de Funcionamiento de Stock") = 2 Then
     Call CONECRUN("FSETUP04/OPTRAJUST", "Configuración de Funcionamiento")
     Call FINAL("")
   End If
End Sub

Private Sub Command27_Click()
    Command28.Enabled = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "MEDITEA") > 0 Then
         REINEQUI07.Show 1    ' provisorio - falta emprolijar (EPB - 5/2/2009)
    ElseIf InStr(EPAC$, "ACHERNAR") > 0 Then
10      OPX% = ALTERNA%("Fraccionamiento de Bobinas\Preparación de Tintas\Modificar Ubicación\Despachos Cargados por Colectora\Otros Accesos Directos")
        If OPX% = 1 Then
            Call FRABOBIN
          ElseIf OPX% = 2 Then
            Call PREPATIN
          ElseIf OPX% = 3 Then
            REUBIC07.Show 1
            GoTo 10
          ElseIf OPX% = 4 Then
            DESCOL09.Show 1
            GoTo 10
          ElseIf OPX% = 5 Then
            Call ACCDIR("CONSTOCK")
        End If
    ElseIf InStr(EPAC$, "NICAR") > 0 Then
        INGLOT07.Show 1
      Else
        Call ACCDIR("CONSTOCK")
    End If
    Command28.Enabled = True
End Sub

Sub FRABOBIN()
    If EXISTE%("TRAZAB07.EXE") > 0 Then
       ACONEC$ = "TRAZAB07/FRABOBIN"
       Call CONECRUN(ACONEC$, "Fraccionamiento de Bobinas")
    End If
End Sub
'
Sub PREPATIN()
    If EXISTE%("TRAZAB07.EXE") > 0 Then
       ACONEC$ = "TRAZAB07/PREPATIN"
       Call CONECRUN(ACONEC$, "Preparacion de Tintas")
    End If
End Sub
'
Private Sub Command28_Click()
    Command28.Visible = False
    ACONEC$ = "BRECEP07"
    Call CONECRUN(ACONEC$, "Recepcion de Materiales")
    Command28.Visible = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("CONSTOCK")
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    If EXISTE%("RETRAS02.EXE") > 0 Then
        Call CONECRUN("RETRAS02", "Emision de Remito de Traslado")
      Else
        Call CONECRUN("REMITRAS", "Emision de Remito de Traslado")
    End If
    Command8.Enabled = True
End Sub

Private Sub Form_Activate()
    TEXPE$ = "Remito de Ventas - Clientes"
    If Control$("", "MODOEXPE") = "ODESPACH" Then
        TEXPE$ = "Orden de Despacho"
    End If
    Command23.ToolTipText = TEXPE$
    Call FRESHALL
End Sub

'
Sub VALEMAT_CONTEPLAST()
     '\\\OT-06-0244-WAR-27/06/06///
     EPAC$ = TRIM$(UCase$(EMPREAC$))

     RTA1% = ALTERNA%("Registrar Vale de Consumo \Anular Vale de Consumo")
     Select Case RTA1%
       Case 1
         GoTo 500
       Case 2
         Call ANUVALEMAT
         GoTo 9999
       Case Else
        GoTo 9999
     End Select
     '\\\OT-06-0244-WAR-FIN///
500  DepOrig% = XVALO(Control$("", "DEPOCEN"))
     DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DepOrig%))
     '
1100 Screen.MousePointer = 11
     CMO$ = TRIM$(Control$("", "CODCONSU"))
     If CMO$ = "" Then CMO$ = "VM"
     '
     Call APERBASDAT("STOCKS")
     'SQLX$ = "SELECT * FROM MOVISTO "
     'SQLX$ = SQLX$ + "WHERE CodiMovi = '" & CMO$ & "' "
     'SQLX$ = SQLX$ + "ORDER BY VAL(MID$(DoPriCor,4,7)) DESC"
     '
     'Set MOVISTEMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
         
     UNRO& = 0
'     'SQLX$ = " SELECT max(substring(DoPriCor, 4,12)) as Maximo FROM MOVISTO" se comento esta linea por que le faltaba el cast abajo agregado(por ese motivo devolvia mal el valoar maximo)
'     SQLX$ = " SELECT max(cast(substring(DoPriCor, 4,12)as int)) as Maximo FROM MOVISTO"
'     SQLX$ = SQLX$ + " WHERE CodiMovi = '" & CMO$ & "' "
'
'     Dim MIRS As ADODB.Recordset
'     Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$), Cantidad)
'     If Not IsNull(MIRS!maximo) Then
'          UNRO& = XVALO(MIRS!maximo)
'     End If
     
     'On Error Resume Next
     'With MOVISTEMP
     '  .MoveFirst
     '  If Err < 1 Then
     '    UNRO& = XVALO(Mid$(!DoPriCor, 4, 7))
     '  End If
     '  On Error GoTo 0
     'End With
     '
     NVM$ = TRIM$(ProNroComprobante(CMO$))
     UNRO& = XVALO(NVM$)
     '
1002 NROVAL& = UNRO&
'     MIRS.Close
'     Set MIRS = Nothing
 
     Call FRESHALL
     '
1105 Call CIERRARCH("*.*")
     '
1111 HOII% = HOY(HOS$)
     VDEF$ = MKS$(NROVAL&) + MKI$(HOII%) + Chr$(DepOrig%) + Chr$(0) + Space$(32)
     PEDY$ = OBJPLA$("VALEMATER", VDEF$)
     If PEDY$ = "" Then GoTo 9999
     FF% = CVI(Mid$(PEDY$, 5, 2))
     DepOrig% = Asc(Mid$(PEDY$, 7, 1))
     DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DepOrig%))
     REFE$ = Mid$(PEDY$, 9, 32)
     CCOS1$ = Mid$(PEDY$, 41, 8)
     NC1 = 0
     If InStr(EPAC$, "ENVAPLAST") > 0 Then 'ENJUAGUE PARA GEENERAR UN VALE DE CONSUMO HACIA CONTEPLAST
       DESCRICOSTO$ = TRIM$(ECOARCH("TACENCOS", TRIM$(CCOS1$)))
       strArray = Split(DESCRICOSTO$, "|")
       NC1 = XVALO(strArray(0)) 'EJEMPLO CODIGO CENTRO DE COSTO: 10
       RASO$ = TRIM$(UCase$(strArray(1))) 'EJEMPLO CODIGO CENTRO DE COSTO: 10
       If NC1 > 0 Then          'EJEMPLO REFERENCIA: 3930|CONTEPLAST
       
          If RASO$ = "CONTEPLAST" Or RASO$ = "PEDRO DE MENDOZA" Or RASO$ = "CAMBACERES" Or InStr(RASO$, "CLIENTE GENERICO") > 0 Then
             CCOS1$ = "" 'LO BORRO PARA QUE NO LO VALIDE
          End If
       End If
     End If

     '\\\OT-06-0244-WAR-28/06/06///
     If TRIM$(CCOS1$) <> "" And VALIPROY%(CCOS1$, FF%) = (-1) Then
        Call MENSERR(24, "Centro de Costos No Vigente\Según Fecha Seleccionada.")
        GoTo 1111
     End If
     '\\\OT-06-0244-WAR-FIN///
    If TRIM$(REFE$) = "" Then
        Call MENSERR(24, "Referencia o Motivo es Mandatoria")
        GoTo 1111
     End If
     If DepOrig% < 1 Or DepOrig% >= 99 Or TRIM$(DEPOSI$) = "" Then
        Call MENSERR(24, "Depósito de Salida no Válido")
        GoTo 1111
     End If
     '
1112 Call SETULTREG("!BOLREDET", 0)
     Call CIERRARCH("*.*")
     '
     Screen.MousePointer = 11
     '
     fff$ = REPLACAR$(FECHATEX$(FF%), "/", "-")
     AVENTA$ = "BOLREDET"
     If EXISTE%("TRAZAB07.EXE") > 0 Or EXISTE%("TRAZAB08.EXE") > 0 Then AVENTA$ = "MOVITRAZ"
1113 VTB% = VENTABU%(AVENTA$ + "/NC/TITUPAN=Vale de Materiales - " + TRIM$(DEPOSI$) + " - " + fff$)
     If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
         GoTo 1105
     End If
     '
     CONTA% = ULTREG&("!BOLREDET")
     '
1115 If CONTA% > 0 Then
       Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
       '
       '\\\OT-05-0730-WAR-10/11/05///
       CAITMO% = 0
       For YX& = 1 To ULTREG&("!BOLREDET")
           YY$ = REGLEIDO$("!BOLREDET", YX&)
           CI% = CVI(Mid$(YY$, 83, 2))
           CAN# = CVD(Mid$(YY$, 89, 8))
           If CI% > 0 Then
             If CI% <= NUMAS% Then
               If CAN# > 0.005 Then
                 CAITMO% = CAITMO% + 1
                 CIMOVIM%(CAITMO%) = CI%
                 CAMOVIM(CAITMO) = (-1) * CAN#
               End If
             End If
           End If
       Next YX&
       If PUEDEMOVI%(DepOrig%) < 0 Then GoTo 1113 ' CANCELA SALIDA
       '\\\OT-05-0730-WAR-FIN///
       '
       
       ' VALIDA LOS DATOS INGRESADOS
       HUBOSELOT% = 0
       For YX& = 1 To ULTREG&("!BOLREDET")
          YY$ = REGLEIDO$("!BOLREDET", YX&)
          CI% = CVI(Mid$(YY$, 83, 2))
          If ESTRAZABLE%(CI%) > o Then
              ' OBLIGA A INGRESAR EL LOTE
              ' COMPARA CANTIDAD DESPACHADA CON SALDO DEL LOTE
              CANTSA = CVD(Mid$(YY$, 89, 8))
              DEPX% = DepOrig%
              LOTEX$ = TRIM$(Mid$(YY$, 107, 12))
              If LOTEX$ <> "" Then
                 If SaldoLote(CI%, LOTEX$, DEPX%) >= CANTSA Then
                    GoTo 19
                 End If
              End If
              LOTEX$ = LOTESELEC$(CI%, DEPX%, CANTSA)
              HUBOSELOT% = 1
              If LOTEX$ = "" Then
                 GoTo 1113
              End If
19             Call REPLA(YY$, LOTEX$, 69, 12)
              Call REPLA(YY$, LOTEX$, 107, 12)
              Call GRAREG("!BOLREDET", YY$, YX&)
          End If
       Next YX&
       Call CIERRARCH("!BOLREDET")
       If HUBOSELOT% = 1 Then GoTo 1113
       '
       Call SETULTREG("!OBSERVOF", 0)
       Call CIERRARCH("!OBSERVOF")
       VTB% = VENTABU%("OBSERVOF/TITUPAN=Vale de Materiales - " + TRIM$(DEPOSI$) + " - " + fff$)
       '
       AINDI% = 0
       NUORIT% = (-1)
       For YX& = 1 To ULTREG&("!BOLREDET")
         YY$ = REGLEIDO$("!BOLREDET", YX&)
         CI% = CVI(Mid$(YY$, 83, 2))
         CAN# = CVD(Mid$(YY$, 89, 8))
         LOTE$ = Mid$(YY$, 69, 12)
           If TRIM$(LOTE$) = "" Then LOTE$ = Mid$(YY$, 107, 12)
         REMI$ = CMO$ + "." + TRIM$(Str$(NROVAL&))
         If CI% > 0 Then
           If CI% <= NUMAS% Then
             If CAN# > 0.005 Then
               DOSEC$ = ""
1120           CANTI = (-1) * CAN#
               NUORIT% = NUORIT% + 1
               '\\\OT-06-0244-WAR-20/06/06///
               'Si Tiene Centro de Costos => Valorizo
               VUNI# = 0
               If CCOS1$ <> "" Then
                 VUNI# = COSPESOS(CI%)
                 DOSEC$ = "CC-" + CCOS1$
               End If
               '\\\OT-06-0244-WAR-FIN///
               Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMI$, REMI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, 0, NC1, DepOrig%, CANTI)   '\\\OT-06-0244-WAR-20/06/06///
             End If
           End If
         End If
       Next YX&
       Call CIERRARCH("MOVISTO")
       '
       If NC1 > 0 Then 'CASO ENVAPLAST
         Call PrintValeConsuEnvaplast(NC1, REMI$, FF%, REFE$)
         GoTo 9999
       End If
       
       FORIPRE$ = TRIM$(Control$("", "FORCONSU"))
       If FORIPRE$ <> "" Then
         LU$ = LUDAT$
         If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
           RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
           Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Vale de Consumo de Materiales:\  \'" + RMCC$ + "'.")
           GoTo 9999
         End If
       End If
       '
       If TRIM$(FORIPRE$) <> "" Then
         '
         'VALIDACION DE MODO DE IMPRESION
         MODOIMPRE07.Check7.Caption = "Imprime Vale de Consumo"
         MODOIMPRE07.Label1.Caption = "" 'LIMPIA EL LABEL DONDE PONE OK, SI ACEPTA
         MODOIMPRE07.Check8.Value = 0
         Screen.MousePointer = 1
         MODOIMPRE07.Show 1
         CANCELPRINT% = 0
         If MODOIMPRE07.Label1.Caption = "OK" Then
          If MODOIMPRE07.Check8.Value > 0 Then CANCELPRINT% = 1
         End If
         Screen.MousePointer = 11
         '
         FEX = FF%
         FECHDOC$ = FECHATEX$(FEX)
         NUMEDOC$ = REMI$
         TIPODOC$ = "Vale de Consumo de Materiales"
         '
         CODPARAM$(1) = "REF-MAT1": DOCPARAM$(1) = REFE$
         CODPARAM$(2) = "DEPOSITO": DOCPARAM$(2) = DEPOSI$
         ' Muestra el Centro de Costo en el formulario de impresion
         CODPARAM$(3) = "REF-MAT2": DOCPARAM$(3) = TRIM$(CCOS1$) + " - " + TRIM$(ECOARCH("TACENCOS", TRIM$(CCOS1$)))
         CAPARDOC% = 3
         '
         CODTABU1$(1) = "COD-MPRI"
         CODTABU1$(2) = "DES-MPRI"
         CODTABU1$(3) = "CAN-MPRI"
         CODTABU1$(4) = "CAKILOS"
         CODTABU1$(5) = "UNI-MPRI"
         CODTABU1$(6) = "UBI-MPRI"
         CACOLTAB1% = 6
         '
         DESTIDOC% = 0
         CAITAB1% = 0
         '
         For YX& = 1 To ULTREG&("!BOLREDET")
           YY$ = REGLEIDO$("!BOLREDET", YX&)
           CI% = CVI(Mid$(YY$, 83, 2))
           CAN# = CVD(Mid$(YY$, 89, 8))
           If CI% > 0 Then
             If CI% <= NUMAS% Then
               If CAN# > 0.005 Then
                 UNID$ = TRIM$(Mid$(YY$, 85, 4))
                 If UNID$ = "" Then UNID$ = Unimed$(CI%)
                 CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
                 KILO$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 12, 3, 2)
                 NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
                 NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
                 '
                 CAITAB1% = CAITAB1% + 1
                 TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
                 TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
                 TETABU1$(3, CAITAB1%) = CANS$
                 TETABU1$(4, CAITAB1%) = KILO$
                 TETABU1$(5, CAITAB1%) = UNID$
                 TETABU1$(6, CAITAB1%) = UBISTO$(CI%) '
               End If
             End If
           End If
         Next YX&
         '
         Call PRINTDOC("FORCONSU")   ' Remito en Consignacion
       End If
       '
1199 End If
     Screen.MousePointer = 1
     GoTo 1100
     '
9999 Screen.MousePointer = 1
     Call CIERRARCH("*.*")
     '
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Openforms = DoEvents
   Wend
End Sub
Sub PrintValeConsuEnvaplast(NCLI, REMI$, FF%, REFE$)
      DatoExtRem.Text2 = NCLI
      DatoExtRem.Text6 = rasocli$(NCLI)
      DatoExtRem.Text4 = FECHATEX$(FF%)
      DatoExtRem.Text16(0) = REFE$
      DatoExtRem.Text3 = REMI$
      DatoExtRem.Show 1
      Exit Sub
'    NCLI = Val(Text1.TEXT)
'    REMI$ = TRIM$(Text11.TEXT)
    NROO& = Val(Mid$(REMI$, 4))
'    FF% = FECHANUM(Text9.TEXT)
'    DEPO% = Val(Text10.TEXT)
'    PREFIREM$ = TRIM(Label10)
'    NROREMIT$ = TRIM$(Mid$(REMI$, 4))
'    While Len(NROREMIT$) < 8: NROREMIT$ = "0" + NROREMIT$: Wend
'    NROREMIT$ = CMO$ + "-X-" + PREFIREM$ + "-" + NROREMIT$
    '
    HO% = HOY(HOS$)
    HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
    User% = USNBR% Mod 100
    CTRXX$ = MKI$(HO%) + MKI$(Int(HP)) + Chr$(User%) + String$(2, 0)
    '
    YXI% = 0
    NUORIT% = 0
    Call APERBASDAT("STOCKS")
    Call APERBASDAT("PEDCLI")
    NOCLOSE = 1: NOARRAS% = 1

    STRNPEDI$ = ""

    '
    '
    GoSub PRIDOCREMI
    '
GoTo 999
    '
    '
    '
    
PRIDOCREMI:
    '
    PRINTERPORT$ = TRIM$(Control$(IDENTER$, "PORTREM"))
 'MsgBox "printerport = *" + PRINTERPORT$ + "*"
    FORIPRE$ = TRIM$(Control$("", "FOREMIFA"))
    'ASIGNO EL FORMULARIO CONFIGURADO - SI NO TIENE CONFIGURADO FRM LE VUELVO A ASIGNAR EL GENERAL
    
    If REMITOSINCARGO% > 0 Then
        FORIPRE$ = TRIM$(Control$("", "FORREMSC"))
        If FORIPRE$ = "" Then
          FORIPRE$ = TRIM$(Control$("", "FOREMIFA"))
        End If
    End If
    
 'MsgBox "foripre = *" + FORIPRE$ + "*"
    If TRIM$(FORIPRE$) <> "" Then
      '
'      PREFIREM$ = TRIM(Label10)
'      If PREFIREM$ = "" Then PREFIREM$ = "0001"
'      NROREMIT$ = TRIM$(Mid$(REMI$, 4))
'      While Len(NROREMIT$) < 8: NROREMIT$ = "0" + NROREMIT$: Wend
'      NROREMIT$ = PREFIREM$ + "-" + NROREMIT$
      '
      Call LIMPIATETAB
      FEX = FF%
      FECHDOC$ = FECHATEX$(FEX)
      NUMEDOC$ = NROREMIT$
      TIPODOC$ = "Remito"
      
'      If InStr(EPAC$, "ENVAPLAST") > 0 Then NCLIX& = NCLI: NROPROV_PROPIO = env.NROPROV_ENVAPLAST_WALDOBTT(NCLIX&)
'      If Option2.Value = True Then
'        TIPODOC$ = "Remito (en Consignación)"
'      End If
'      '
'      CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Text8.TEXT)
'      CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = TRIM$(Text15.TEXT)
'      CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Text12.TEXT)
'      CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = TRIM$(Text17.TEXT)
'      If InStr(EPAC$, "ENVAPLAST") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Then
'        If TRIM$(Text8.TEXT) <> "" Then TEXTOENTREGA$ = TRIM$(Text8.TEXT)
'        If TRIM$(Text14.TEXT) <> "" Then TEXTOENTREGA$ = TEXTOENTREGA$ & "-" & TRIM$(Text14.TEXT)
'        If TRIM$(Text21.TEXT) <> "" Then TEXTOENTREGA$ = TEXTOENTREGA$ & "-" & TRIM$(Text21.TEXT)
'        If TRIM$(TXTCODPOSTAL.TEXT) <> "" Then TEXTOENTREGA$ = TEXTOENTREGA$ & "-" & TRIM$(TXTCODPOSTAL)
'        CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = TEXTOENTREGA$
'      End If
      
      
      CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Text13.TEXT)
      CODPARAM$(7) = "DEPOSAL ": DOCPARAM$(7) = TRIM$(DEPOSAL$)
      'CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = TRIM$(NROREMIT$)
      CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = Mid$(NROREMIT$, 6) ' SE IMPRIME FORMATO 0000-00000000
      CODPARAM$(9) = "NRO-OCOM": DOCPARAM$(9) = TRIM$(Text20.TEXT)
      CODPARAM$(10) = "PATE-TRA": DOCPARAM$(10) = TRIM$(Text18.TEXT)
      CODPARAM$(11) = "N-AUTENT": DOCPARAM$(11) = TRIM$(NORVENTA)
      CODPARAM$(12) = "COND-TRA": DOCPARAM$(12) = TRIM$(Text19)
      CODPARAM$(13) = "LOCA-TRA": DOCPARAM$(13) = TRIM$(Text21)
      CODPARAM$(14) = "NRO-AENT": DOCPARAM$(14) = TRIM$(Text22)

      '\\\OT-07-0007-RG-08/01/07///
      HORA$ = HORAENT$(NCLI)
      CODPARAM$(15) = "Z-PLANO": DOCPARAM$(15) = TRIM$(HORA$)
      CODPARAM$(16) = "IMP-TOTA": DOCPARAM$(16) = TRIM$(FORMATNUM$(ACUMLAIMPOX#, "F14.2"))
      CODPARAM$(17) = "DOCU-APL": DOCPARAM$(17) = TRIM$(STRNPEDI$)
      CODPARAM$(18) = "DESTINO": DOCPARAM$(18) = TRIM$(rasocli$(NCLI)) + " - Almacén " + Label11
      
      CODPARAM$(19) = "PROV-CLI": DOCPARAM$(19) = NROPROV_PROPIO
      '\\\OT-07-0007-RG-FIN///
      CAPARDOC% = 19
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "CANTIDAD"
      CODTABU1$(4) = "NRO-LOTE"
      CODTABU1$(5) = "UNIMED"
      CODTABU1$(6) = "N-OCOMPR"
      CODTABU1$(7) = "ORD-ITEM"
      CODTABU1$(8) = "REF-MAT1"
      CODTABU1$(9) = "REF-MAT2"
      CODTABU1$(10) = "PRE-NETO"
      CODTABU1$(11) = "NET-ITEM"
      CODTABU1$(12) = "REF-MAT6"
      CODTABU1$(13) = "REF-MAT5"
      CACOLTAB1% = 13
      '
      DESTIDOCNUE = NCLI
      CAITAB1% = 0
      '
      For YX& = 1 To ULTREG&("!BOLREDET")
        YY$ = REGLEIDO$("!BOLREDET", YX&)
        CI% = CVI(Mid$(YY$, 83, 2))
        CAN# = CVD(Mid$(YY$, 89, 8))
        NPEDI& = CVS(Mid$(YY$, 151, 4))
        CAC# = 0
        If CI% > 0 Or Frame15.Visible = True Then
          If CI% <= NUMAS% Or Frame15.Visible = True Then
            If CAN# + CAC# > 0.005 Or Frame15.Visible = True Then
              '
              UNID$ = Unimed$(CI%)
              If UNID$ = "" Then UNID$ = Unimed$(CI%)
              CANS$ = CSTRING$(MKD$(CAN#), 3, 12, 1, 2)
              If CAC# >= 0.05 Then
                CANS$ = CSTRING$(MKD$(CAC#), 3, 12, 1, 2)
                Call REPLA(CANS$, "(*)", 1, 3)
              End If
              NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
              NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
              ORD& = XVALO(Mid$(YY$, 241, 5)) ' NUMERO DE ORDEN INGRESADO EN VENTANA DE ARCON NEUQUEN.
              '
              SERINU$ = TRIM$(Mid$(YY$, 209, 16)) + TRIM$(Mid$(YY$, 225, 8))
              If SERINU$ <> "" Then
                If TRIM$(Mid$(YY$, 233, 8)) <> "" Then
                  SERINU$ = SERINU$ + " - "
                End If
              End If
              SERINU$ = SERINU$ + TRIM$(Mid$(YY$, 233, 8))
              '
              CAITAB1% = CAITAB1% + 1
              CODDX$ = CODEXT$(CI%)
              If TRIM$(Mid$(YY$, 97, 8)) <> "" Then 'VALIDACION POR QUE SI NO QUEDA EL GUION SOLO
                CODDX$ = CODEXT$(CI%) + "-" + TRIM$(Mid$(YY$, 97, 8)) ' agrega subcodigo
              End If
              If REPLAORI% = 1 Then
                CODDX$ = CODORIG$(CI%, NCLI)
              End If
              '
              TETABU1$(1, CAITAB1%) = CODDX$
              TETABU1$(2, CAITAB1%) = DESCRIT0$(CI%)
              If Frame15.Visible = True Then
                TETABU1$(2, CAITAB1%) = Mid$(YY$, 5, 64)
              End If
              '
              TETABU1$(3, CAITAB1%) = CANS$
              TETABU1$(4, CAITAB1%) = NULOTE$
              TETABU1$(5, CAITAB1%) = UNID$
              TETABU1$(6, CAITAB1%) = NUOCO$
              TETABU1$(7, CAITAB1%) = Str$(CAITAB1%)
              TETABU1$(8, CAITAB1%) = NUOCOCLI$(Val(NUOCO$), CI%)
              TETABU1$(9, CAITAB1%) = NUPLANO$(CI%) '
              '
              'VALORIZACION ALEJPA SEGUN LISTA VIGENTE
              IMPOUNI# = 0: IMPORTEX# = 0
              If InStr(UCase$(EMPREAC$), "ALEJPA") > 0 Then
                IMPOUNI# = PRELISTACT#(CI%, TIMONE%)
                IMPOUNI# = IMPOUNI# * TICAMONE(TIMONE%) 'VALOR UNITARIO EN PESOS
                IMPORTEX# = CAN# * IMPOUNI#             'IMPORTE EN PESOS
                ACUMLAIMPOX# = ACUMLAIMPOX# + IMPORTEX# 'IMPORTE ACUMULADO EN PESOS
              End If

              TETABU1$(10, CAITAB1%) = FORMATNUM$(IMPOUNI#, "F12.2")
              TETABU1$(11, CAITAB1%) = FORMATNUM$(IMPORTEX#, "F12.2")
              If InStr(EPAC$, "ENVAPLAST") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Then
                 DES2$ = TRIM$(DESCRIT2$(CI%))
                 If DES2$ <> "" Then
                    Call PRINTEXT(DESCRIT2$(CI%), 2, "DES-MAT", FORIPRE$)
                 End If
              End If

                CORIG$ = TRIM$(CODORIG$(CI%, NCLI))
                'If CORIG$ = TRIM$(CODDX$) Then CORIG$ = ""
                If CORIG$ = CODEXT$(CI%) Then CORIG$ = ""
                'If CORIG$ <> CODEXT$(CI%) Then CORIG$ = CODEXT$(CI%) 'ESTA LINEA HABRIA QUE QUITAR
              If InStr(EPAC$, "ENVAPLAST") > 0 Then
'                 If CORIG$ <> "" Then 'al final dijo Enrique que no quiere que se imprima
'                   Call PRINTEXT(CORIG$, 1, "COD-MAT", FORIPRE$)
'                 End If
              Else
                 TETABU1$(12, CAITAB1%) = CORIG$
              End If
              'FIN ALEJPA
              '
              TETABU1$(13, CAITAB1%) = SAFETEXT$(ORD&) ' NUMERO DE ORDEN INGRESADO EN VENTANA DE ARCON NEUQUEN.
              
              If EsReteplast% = 1 Then
                NumeroRemito$ = NUMEDOC$
                CodigoInternoProducto% = CI%
                
                ConsultaSql$ = "select count(*) bultos from bultosfabric b with(nolock)" _
                & " inner join ordefabr o with(nolock) on (b.ordenfabricacion=o.idsubor)" _
                & " where b.remito='" & NumeroRemito$ & "' and o.cod_inte=" & CStr(CodigoInternoProducto%)
                Set Rst = READSET(ConsultaSql$)
                If Not (Rst.BOF And Rst.EOF) Then
                    CantidadBultosDespachados& = XVALO(Rst!bultos)
                End If
                On Error Resume Next
                Rst.Close
                Set Rst = Nothing
                On Error GoTo 0
                TETABU1$(13, CAITAB1%) = CStr(CantidadBultosDespachados&)
              End If
              
              'IMPRESION DE ESPECIFICACION
              If Control("FACTUREM", "INCLUSPE") = "SI" Then ' IMPRIME ESPECIFICACIONES DEL ARTICULO
                XX$ = ESPECIFIT$(CI%)
                Call PRINTEXT(XX$, 2, "DES-MAT", FORIPRE$)
              End If
              '
              If Control("REMSTAN", "PRIOBSPE") = "SI" Then 'OBSERVACIONES DE PEDIDOS
                NUPEDIORI& = CVS(Mid$(YY$, 151, 4))
                NUORITPE% = CVS(Mid$(YY$, 189, 4))
                OBSPE$ = OBSERVAPEDI$(CI%, NCLI, NUPEDIORI&, NUORITPE%)
                Call PRINTEXT(OBSPE$, 2, "DES-MAT", FORIPRE$)
              End If
              If InStr(EPAC$, "ENVAPLAST") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Then  'IMPRIME EL LOTE DEBAJO DE LA DESCRIPCION Y/O OBSERVACION DEL PEDIDO
                 Call PRINTEXT(NULOTE$, 2, "DES-MAT", FORIPRE$)
                 Call CARLINSEP("INVISIBLE")         'Carga linea de separacion entre codigo y codigo EN ESTE CASO SIN LINEA
              Else
                 Call CARLINSEP
              End If
              
              If SERINU$ <> "" Then
                CAITAB1% = CAITAB1% + 1
                For KI1% = 1 To 20
                  TETABU1$(KI1%, CAITAB1%) = ""
                Next KI1%
                TETABU1$(2, CAITAB1%) = "S.No." + SERINU$
              End If
              '
              Call DETLOTREM(CI%, YX&) 'SI NO LO AGREGA A LA DESCRIPCION
              '
              'NCIY% = NCLI
              CIIY% = CI%
              'Call DATADI(NCIY%, CIIY%, ORCO$, NORI$, NPLA$)
            End If
          End If
        End If
      Next YX&
      '
 'MsgBox "antes de PRINTDOC(FOREMIFA)"
      CantidadCopiasRemito% = XVALO(Text23)
      If CantidadCopiasRemito% > 0 Then
        CACOPDOC% = CantidadCopiasRemito%
        Call PRINTDOC("@" + FORIPRE$)   ' Remito de Facturacion
        CACOPDOC% = 0
      Else
        Call PRINTDOC("@" + FORIPRE$)   ' Remito de Facturacion
      End If
      '
    End If
    '
Return
    '
ANOTAPEDI:                ' ANOTADOR DE PEDIDOS
   '
   NPEDIDX& = NUPEDIORI&
   If NPEDIDX& <> PEDIANT& Then
      If PEDIANT& <> 0 Then
          Call AGREGA_ANOTAPED(PEDIANT&, Space$(10) + String$(56, "-"))
          Call AGREGA_ANOTAPED(PEDIANT&, String$(88, "="))
      End If
      PEDIANT& = NPEDIDX&
      TTXX$ = "Remito " + TRIM$(Mid$(Text11, 4)) + ":"
      '
      TTXYZ$ = Space$(76)
      HOII% = HOY(HOS$)
      Call REPLA(TTXYZ$, HOS$, 1, 8)
      Call REPLA(TTXYZ$, TTXX$, 11, 56)
      Call REPLA(TTXYZ$, USERTER$, 69, 24)
      Call AGREGA_ANOTAPED(NPEDIDX&, TTXYZ$)
      Call REPLA(TTXYZ$, Space$(88), 1, 88)
      Call AGREGA_ANOTAPED(NPEDIDX&, TTXYZ$)
      '
   End If
   '
   TTXX$ = Space$(76)
   Call REPLA(TTXX$, "- " + CODEXT$(CI%), 11, 18)
   Call REPLA(TTXX$, DESCRIT0$(CI%), 29, 38)
   Call REPLA(TTXX$, FORMATNUM$(CAN#, "F10.1"), 57, 10)
   Call AGREGA_ANOTAPED(NPEDIDX&, TTXX$)
   '
Return

999


End Sub


Private Sub Form_Load()
    '
    Call TRANSLABELS(Name)
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture2)
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If EXISTE%("REMITO07.EXE") > 0 Then Command23.Enabled = True
    If EXISTE%("ORCOMP07.EXE") > 0 Then Command22.Enabled = True
    If EXISTE%("RETRAS07.EXE") > 0 Then Command8.Enabled = True
    If EXISTE%("LISEMPAQ.EXE") > 0 Then Command10.Enabled = True
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
      Call CIERRARCH("*.*")
      Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Hide
End Sub

Sub RECPROP()
    '
    If PRECOCIE% = 0 Then
       PRECOCIE% = (-1)
       If Control$("", "PRECOCIE") = "SI" Then
          PRECOCIE% = 1
       End If
    End If
    '
10  HOOI% = HOY(HO$)
    If ULTREG&("ORDEFABR") < 1 Then   ' por numero de O/F
        GoTo 100
    End If
    '
    MODO% = COMALTER%("Puede Declarar Entradas a Stock:\  \a) por Cierre Parcial o Total de Orden de Fabricación\o\b) en Forma Directa por Armado de Conjunto\\Cierre de O./Fabricación\Armado de Conjunto\Aprobación\Cancelar")
    On MODO% GoTo 20, 100, 50
    Exit Sub
    '
20  Call ORFCIE
    Exit Sub
    '
50  If InStr(1, EMPREAC$, "ARCON") > 0 Then
       If DERACCE%("APCIEOF", "Aprobación Cierre Parcial de O/F") >= 2 Then
          Call APCIEOF
          GoTo 10
       End If
    Else
       Call COMUNI("Opcion No Disponible en la Presente Versión")
    End If
    Exit Sub
    '
100 Screen.MousePointer = 11
    For U& = ULTREG&("RECPROP") To 1 Step -1
        XX$ = REGLEIDO$("RECPROP", U&)
        If CVS(Mid$(XX$, 59, 4)) > 0 Then
            UPARTE& = CVS(Mid$(XX$, 59, 4))
            GoTo 105
        End If
    Next U&
    UPARTE& = 0
    '
105 Screen.MousePointer = 1
    Call BLANARCH("!CAREPRO")
    '
    VDEF$ = String$(88, 0)
    Call REPLA(VDEF$, MKI$(HOOI%), 1, 2)
    Call REPLA(VDEF$, MKS$(UPARTE& + 1), 85, 4)
    Call REPLA(VDEF$, Chr$(0), 37, 1)
    Call REPLA(VDEF$, Chr$(0), 38, 1)
    '
110 XX$ = OBJPLA$("RECPROP", VDEF$)
    If XX$ = "" Then Exit Sub
    '
160 VDEF$ = XX$
    FECHA% = CVI(Left$(XX$, 2))
    If FECHA% > HOOI% Or FECHA% < 1 Then
        Call MENSERR(24, "Fecha Incorrecta")
        GoTo 110
    End If
    '
'    CI% = CVI(Mid$(XX$, 3, 2))
'    If TRIM$(CODEXT$(CI%)) = "" Then
'        Call MENSERR(24, "Codigo Inexistente.")
'        GoTo 110
'    End If
'    CODCON$ = CODEXT$(CI%)
'    '
    DP1% = Asc(Mid$(XX$, 37, 1))
    DP2% = Asc(Mid$(XX$, 38, 1))
    If TRIM$(ECOARCH$("ECODEP", Chr$(DP1%))) = "" Then
        Call MENSERR(24, "Falta Depósito de Entrada")
        GoTo 110
    End If
    '
    If TRIM$(ECOARCH$("ECODEP", Chr$(DP2%))) = "" Then
        Call MENSERR(24, "Falta Depósito de Consumo")
        GoTo 110
    End If
    '
120 U3& = 0
    For U1& = 1 To ULTREG&("!CAREPRO")
      XY$ = REGLEIDO$("!CAREPRO", U1&)
      CANTI = CVD(Mid$(XY$, 23, 8))
      If Abs(CANTI) >= 0.05 Then
        U3& = U3& + 1
        Call GRAREG("!CAREPRO", XY$, U3&)
      End If
    Next U1&
    Call SETULTREG("!CAREPRO", U3&)
    '
    VTB% = VENTABU%("CAREPRO")
    If VTB% < 0 Then GoTo 110
    '
    CARE& = ULTREG&("!CAREPRO")
    If CARE& < 1 Then GoTo 120
    '
    NUORIT% = 0
    For U& = 1 To CARE&
       XY$ = REGLEIDO$("!CAREPRO", U&)
       CANTI = CVD(Mid$(XY$, 23, 8))
       Scrap = 0
       CI% = CVI(Left$(XY$, 2))
       CODICONJ$ = CODEXT$(CI%)
       CICONJ% = CI%
       '
       '\\\OT-06-0165-WAR-08/05/06///
       'If CANTI < 0 Or SCRAP < 0 Or CANTI + SCRAP < 0.05 Then
       If Scrap < 0 Then
       '\\\OT-06-0165-WAR-FIN///
          Call MENSERR(24, "Cantidades Incorrectas en Código " + TRIM$(CODICONJ$))
          GoTo 120
       End If
       '
       NORFABRI$ = ""
       CACONJU# = CDbl(CANTI + Scrap)
       VACONTROL% = 1
       '
       Call SETULTREG("!MATEROF", 0)
       Call CIERRARCH("!MATEROF")
       Call CARCOMPO(CODICONJ$, NORFABRI$, CACONJU#, VACONTROL%, 2)
       '
       If VACONTROL% < 0 Then GoTo 120           ' cancelacion
       '
       '\\\OT-05-0730-WAR-10/11/05///
       CAITMO% = 0
         For YX& = 1 To ULTREG&("!MATEROF")
            YY$ = REGLEIDO$("!MATEROF", YX&)
            CI% = CVI(Mid$(YY$, 1, 2))
            CAN# = CVS(Mid$(YY$, 5, 4))
             If CI% > 0 Then
               If CI% <= NUMAS% Then
                 If CAN# > 0.005 Then
                   CAITMO% = CAITMO% + 1
                   CIMOVIM%(CAITMO%) = CI%
                   CAMOVIM(CAITMO) = (-1) * CAN#
                 End If
               End If
             End If
         Next YX&
       If PUEDEMOVI%(DP2%) < 0 Then GoTo 110 ' CANCELA SALIDA
       '\\\OT-05-0730-WAR-FIN///
       '
       NRO& = CVS(Mid$(XX$, 85, 4))
       NROOF& = CVS(Mid$(XY$, 19, 4))
       If NROOF& < 1 Then NROOF& = NRO&
       DOPRI$ = "OF-" + TRIM$(Str$(NROOF&))
       DOSEC$ = "PF-" + TRIM$(Str$(NRO&))
       LOTE$ = Mid$(XY$, 3, 16)
       If LOTE$ = "" Then
          LOTE$ = DOPRI$
       End If
       REFE$ = Mid$(XX$, 55, 30)
       If TRIM$(REFE$) = "" Then
          REFE$ = "Informe de Produccion"
       End If
       If CANTI < (-0.005) Then
          REFE$ = "Desarmado de Conjunto"
       End If
       '
       NC% = 0
       VUNI# = 0
       MONE$ = ""
       MONEII% = 0
       CMOV$ = "PF"
       '
       NUORIT% = 0
       DEPX% = DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(CICONJ%)
       Call MOVISTOK(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CANTI)
       '
       YY$ = REGBLAN$("RECPROP")
       Call REPLA(YY$, MKI$(FECHA%), 1, 2)
       Call REPLA(YY$, Chr$(DP1%), 3, 1)
       Call REPLA(YY$, Chr$(DP2%), 4, 1)
       Call REPLA(YY$, REFE$, 5, 32)
       Call REPLA(YY$, MKI$(CI%), 37, 2)
       Call REPLA(YY$, MKS$(CANTI), 39, 4)
       Call REPLA(YY$, LOTE$, 43, 16)
       Call REPLA(YY$, MKS$(NRO&), 59, 4)
       Call REPLA(YY$, MKI$(-1), 63, 2)
       Call REGAPP("RECPROP", YY$)
    
       COD$ = CODEXT$(CICONJ%)
       CODCON$ = COD$
       '
       For U2& = 1 To ULTREG&("!MATEROF")
          ZZ$ = REGLEIDO$("!MATEROF", U2&)
          CI2% = CVI(Left$(ZZ$, 2))
          CAN = -1 * CVS(Mid$(ZZ$, 5, 4))
          '
          If TRIM$(UCase$(Unimed$(CI2%))) = "GK" Then
             CAN = CAN / 1000
          End If
          '
          If CI2% > 0 Then
             If CI2% <= NUMAS% Then
                If Abs(CAN) >= 0.05 Then
                   LOTE$ = ""
                   CMOV$ = "CF"
                   DOPRI$ = "OF-" + TRIM$(Str$(NROOF&))
                   DOSEC$ = "PF-" + TRIM$(Str$(NRO&))
                   REFE$ = "CONSUMO S/DESP." + CODCON$
                   '\\\OT-06-0165-WAR-08/05/06///
                   If CAN > 0 Then
                    REFE$ = "DESARMADO CONJ." + CODCON$
                   End If
                   '\\\OT-06-0165-WAR-FIN///
                   DEPX% = DP2%: If DP2% = 99 Then DEPX% = DEPOPRE%(CI2%)
                   '
                   If STODEPOS(CI2%, DEPX%) < (-1) * CAN Then
                     If InStr(EPAC$, "YALTRE") > 0 Then
                       Call CARDEPOS
                       For KKII% = 1 To NUSU%
                         If STODEPOS(CI2%, COSU%(KKII%)) >= (-1) * CAN Then
                           DEPX% = COSU%(KKII%)
                           GoTo 124
                         End If
                       Next KKII%
                     End If
                   End If
                   '
124                NUORIT% = NUORIT% + 1
                   Call MOVISTOK(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CAN)
                End If
             End If
          End If
       Next U2&
       Call CIERRARCH("!MATEROF")
       '
       XY$ = REGLEIDO$("!CAREPRO", U&)
       Call REPLA(XY$, MKD$(0), 23, 8)
       Call GRAREG("!CAREPRO", XY$, U&)
       '
129 Next U&
    Call CIERRARCH("MOVISTO")
    Call CIERRARCH("RECPROP")
    '
180 Call CIERRARCH("*.*")
    Call BLANARCH("!CAREPRO")
    Screen.MousePointer = 1
    GoTo 100
    '
End Sub




Private Sub Timer1_Timer()
    '
End Sub

