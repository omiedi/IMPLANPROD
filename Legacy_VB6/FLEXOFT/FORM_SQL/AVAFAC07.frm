VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form AVAFAC07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Avance de Facturación y Cobranza"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   25
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2D3C0&
      Caption         =   "FILTRO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   3900
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
         Height          =   300
         Left            =   1575
         TabIndex        =   19
         Top             =   280
         Width           =   175
      End
      Begin VB.TextBox TEXT2 
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
         Left            =   1080
         TabIndex        =   20
         Top             =   280
         Width           =   495
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor"
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
         Left            =   -750
         TabIndex        =   22
         Top             =   315
         Width           =   1710
      End
      Begin VB.Label Label23 
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
         Left            =   1920
         TabIndex        =   21
         Top             =   285
         Width           =   1830
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "AGRUPAMIENTO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   120
      TabIndex        =   15
      Top             =   3960
      Width           =   3900
      Begin VB.OptionButton Option6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Por Cliente"
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
         Left            =   2205
         TabIndex        =   17
         Top             =   315
         Width           =   1275
      End
      Begin VB.OptionButton Option5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Por Fecha"
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
         Left            =   630
         TabIndex        =   16
         Top             =   315
         Value           =   -1  'True
         Width           =   1380
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   6500
      Left            =   4200
      ScaleHeight     =   6465
      ScaleWidth      =   930
      TabIndex        =   12
      Top             =   0
      Width           =   960
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   100
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   27
         Top             =   5000
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command4 
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
         Picture         =   "AVAFAC07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Detalle de Facturación por Rango de Fechas."
         Top             =   1560
         Width           =   650
      End
      Begin VB.CommandButton Command8 
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
         Height          =   690
         Left            =   120
         Picture         =   "AVAFAC07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Resúmen Operativo"
         Top             =   890
         Width           =   650
      End
      Begin VB.CommandButton BOTREC 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "AVAFAC07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Genera Reporte Solicitado"
         Top             =   4080
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
         Left            =   120
         Picture         =   "AVAFAC07.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   200
         Width           =   650
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "OPCIONES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      Left            =   120
      TabIndex        =   6
      Top             =   740
      Width           =   3900
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Notas de Crédito"
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
         Index           =   2
         Left            =   630
         TabIndex        =   11
         Top             =   1405
         Value           =   1  'Checked
         Width           =   3060
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Notas de Débito"
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
         Index           =   1
         Left            =   630
         TabIndex        =   10
         Top             =   1095
         Value           =   1  'Checked
         Width           =   2955
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Facturas (Contado y C.Corriente)"
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
         Index           =   0
         Left            =   630
         TabIndex        =   9
         Top             =   775
         Value           =   1  'Checked
         Width           =   3165
      End
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Cobranza por Recibo"
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
         Left            =   315
         TabIndex        =   8
         Top             =   1850
         Width           =   2850
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Facturación"
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
         Left            =   315
         TabIndex        =   7
         Top             =   400
         Value           =   -1  'True
         Width           =   1380
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E2D3C0&
      Caption         =   "SALIDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   3900
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Pantalla"
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
         Left            =   630
         TabIndex        =   5
         Top             =   315
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E2D3C0&
         Caption         =   "Impresión"
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
         Left            =   2205
         TabIndex        =   4
         Top             =   315
         Width           =   1275
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
      Left            =   3840
      TabIndex        =   1
      Top             =   195
      Width           =   175
   End
   Begin VB.TextBox MESAN 
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
      Left            =   1245
      TabIndex        =   0
      Text            =   " "
      ToolTipText     =   "Doble Click para Seleccionar Período"
      Top             =   195
      Width           =   2580
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "AVAFAC07.frx":0A68
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   2040
      OleObjectBlob   =   "AVAFAC07.frx":0C9C
      TabIndex        =   26
      Top             =   -120
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mes y año"
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
      Index           =   10
      Left            =   240
      TabIndex        =   2
      Top             =   320
      Width           =   900
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "Archivo "
      Begin VB.Menu mnuImpresiónComprobantes 
         Caption         =   "Impresión de Comprobantes"
      End
   End
End
Attribute VB_Name = "AVAFAC07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DESDE%, HASTA%, CORRESP$
Dim DESDAT As Double
Dim HASDAT As Double
'
Sub RESUOPERnuevo()
      '
      Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
      If PERIO$ = "" Then Exit Sub
      '
      Dim PERIOD$(), DESDEI%(), HASTAI%(), VENTA#(), COMPRA#(), COBRA#(), PAGOS#()
      '
      'TODO ESTO VA A VER QUE DESCOMENTAR PARA JYM
      'AGREGADO PAA ACTUALIZAR URE& POR QUE SI NO EL VECTOR SOLO SE CARGABA CON 48 Y SI EL RANGO ERA MAS GRANDE NO TRAIA LAS FECHAS MAS VIEJAS
      FEDES = FECHATEX$(DESDEX%)
      FEHAS = FECHATEX$(HASTAX%)
      CANTIMESES% = 1 + DateDiff("M", FEDES, FEHAS)  ' CANTIDAD DE MESES COMPRENIDOS EN EL RANGO
      If CANTIMESES% < 1 Then Call COMUNI("Error de Ingreso de Rango de Fechas"): Exit Sub
'      FIN TODO ESTO VA A VER QUE DESCOMENTAR PARA JYM

      Call CARSELMESA
      URE& = ULTREG&("SELMESA")
      If URE& < CANTIMESES% Then URE& = CANTIMESES% ' CODIGO PARA EL CAMBIO EN JYM

      ReDim PERIOD$(URE&)
      ReDim DESDEI%(URE&)
      ReDim HASTAI%(URE&)
      ReDim VENTA#(URE&)
      ReDim COMPRA#(URE&)
      ReDim COBRA#(URE&)
      ReDim PAGOS#(URE&)
      '
      i& = 0
'      For i& = 1 To URE&
'          RELE$ = REGLEIDO$("SELMESA", i&)
'          PERIOD$(i&) = MESANO$(RELE$, Des%, Has%)
'          DESDEI%(i&) = Des%
'          HASTAI%(i&) = Has%
'      Next i&
      '
'===================================
      For i& = 1 To URE&
          If i& = 1 Then
             MESINICIO = Month(FEDES)
             ANO = Year(FEDES)
             NOMBREMES$ = UCase$(MonthName(MESINICIO)) 'nombre del mes
             MESX$ = FORMATOCON0$(SAFETEXT$(MESINICIO), 2)
             ANOX$ = SAFETEXT$(Right(ANO, 2))
             Des% = FECHANUM("01/" & MESX$ & "/" & ANOX$)
             Has% = FECHANUM(FinDelMes(Des%))
             If Des% < 1 Or Has% < 1 Then
               A = 1
             End If
          Else
             MESINICIO = MESINICIO + 1
             If MESINICIO > 12 Then
               MESINICIO = 1
               ANO = ANO + 1
             End If
             NOMBREMES$ = UCase$(MonthName(MESINICIO))
             MESX$ = FORMATOCON0$(SAFETEXT$(MESINICIO), 2)
             ANOX$ = SAFETEXT$(Right(ANO, 2))
             Des% = FECHANUM("01/" & MESX$ & "/" & ANOX$)
             Has% = FECHANUM(FinDelMes(Des%))
               
          End If
          
          PERIOD$(i&) = NOMBREMES$ & " " & SAFETEXT$(Right(ANO, 2))
          DESDEI%(i&) = Des%
          HASTAI%(i&) = Has%
          
          If DESDEX% > Des% And DESDEX% < Has% Then
             PERIOD$(i&) = PERIOD$(i&) + " (Inc.)"
             DESDEI%(i&) = DESDEX%
          End If
          If HASTAX% > Des% And HASTAX% < Has% Then
             PERIOD$(i&) = PERIOD$(i&) + " (Inc.)"
             HASTAI%(i&) = HASTAX%
          End If
          '
      Next i&
      '

'===================================
      Call APERBASDAT("CABAN")
      '
      'PARA OBTENER VENTAS NETAS
      SQLX$ = "SELECT IDEBECOMP, IHABECOMP, FECHCONTA FROM CAJABANC WITH(NOLOCK) "
      SQLX$ = SQLX$ + " WHERE TipoMovim = 'FVEN'"
      SQLX$ = SQLX$ + " AND  NuClien > 0 "
      '
'      'Set AVANCE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'      Dim AVANCE As ADODB.Recordset
'      Set AVANCE = New ADODB.Recordset
'25    On Error Resume Next
'      AVANCE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'           On Error GoTo 0
'           Call CapturaErrorOpen
'           GoTo 25
'      End If
'      On Error GoTo 0
      '
      Set AVANCE = READSET(SQLX$)
      With AVANCE
          On Error Resume Next
          If (.EOF And .BOF) Then
             On Error GoTo 0
             Screen.MousePointer = 1
             GoTo 10
          End If
          On Error GoTo 0
          '
          Do While Not .EOF
               IMPO# = XVALO(!IDEBECOMP) - XVALO(!IHABECOMP)
               FFI% = CVINT%(!fechconta)
               '
               J& = 0
               For J& = 1 To URE&
                   If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                      VENTA#(J&) = VENTA#(J&) + IMPO#
                      Exit For
                   End If
               Next J&
          .MoveNext
          Loop
          '
      End With
      
      '
      'PARA OBTENER COBRANZAS NETAS
10     SQLX$ = "SELECT IHABECOMP, IDEBECOMP, FECHCONTA FROM CAJABANC WITH(NOLOCK)"
       SQLX$ = SQLX$ + "WHERE TipoMovim = 'RCOB'"
       SQLX$ = SQLX$ + " AND  NuClien > 0 "
       '
'       'Set AVANCE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'       Set AVANCE = New ADODB.Recordset
'26     On Error Resume Next
'       AVANCE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'           On Error GoTo 0
'           Call CapturaErrorOpen
'           GoTo 26
'       End If
'       On Error GoTo 0
       Set AVANCE = READSET(SQLX$)

       With AVANCE
           On Error Resume Next
           If (.EOF And .BOF) Then
              On Error GoTo 0
              Screen.MousePointer = 1
              GoTo 20
           End If
           On Error GoTo 0
          '
           Do While Not .EOF
                IMPO# = XVALO(!IHABECOMP) - XVALO(!IDEBECOMP)
                FFI% = CVINT%(!fechconta)
                '
                J& = 0
                For J& = 1 To URE&
                    If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                       COBRA#(J&) = COBRA#(J&) + IMPO#
                       Exit For
                    End If
                Next J&
           .MoveNext
           Loop
           '
       End With
       '
       'PARA OBTENER COMPRAS NETAS
20     SQLX$ = "SELECT IHABECOMP, IDEBECOMP, FECHCONTA FROM CAJABANC WITH(NOLOCK)"
       SQLX$ = SQLX$ + "WHERE TipoMovim = 'FCOM'"
       SQLX$ = SQLX$ + "AND  NuProve > 0 "
       '
'       'Set AVANCE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'       Set AVANCE = New ADODB.Recordset
'27     On Error Resume Next
'       AVANCE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'           On Error GoTo 0
'           Call CapturaErrorOpen
'           GoTo 27
'       End If
'       On Error GoTo 0
       Set AVANCE = READSET(SQLX$)
       With AVANCE
           On Error Resume Next
           .MoveFirst
           If Err Then
              On Error GoTo 0
              Screen.MousePointer = 1
              GoTo 30
           End If
           '
           Do While Not .EOF
                IMPO# = XVALO(!IHABECOMP) - XVALO(!IDEBECOMP)
                FFI% = CVINT%(!fechconta)
                '
                J& = 0
                For J& = 1 To URE&
                    If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                       COMPRA#(J&) = COMPRA#(J&) + IMPO#
                       Exit For
                    End If
                Next J&
           .MoveNext
           Loop
           '
       End With
       '
       'PARA OBTENER PAGOS NETOS
30     SQLX$ = "SELECT IDEBECOMP, IHABECOMP, FECHCONTA FROM CAJABANC WITH(NOLOCK)"
       SQLX$ = SQLX$ + "WHERE TipoMovim = 'OPAG'"
       SQLX$ = SQLX$ + "AND  NuProve > 0 "
       '
'       'Set AVANCE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'       Set AVANCE = New ADODB.Recordset
'28     On Error Resume Next
'       AVANCE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'           On Error GoTo 0
'           Call CapturaErrorOpen
'           GoTo 28
'       End If
'       On Error GoTo 0
       '
       Set AVANCE = READSET(SQLX$)
       With AVANCE
           On Error Resume Next
           .MoveFirst
           If Err Then
              On Error GoTo 0
              Screen.MousePointer = 1
              GoTo 40
           End If
          '
           Do While Not .EOF
                IMPO# = XVALO(!IDEBECOMP) - XVALO(!IHABECOMP)
                FFI% = CVINT%(!fechconta)
                '
                J& = 0
                For J& = 1 To URE&
                    If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                       PAGOS#(J&) = PAGOS#(J&) + IMPO#
                       Exit For
                    End If
              Next J&
           .MoveNext
           '
           Loop
           '
       End With
       '
40     Call BLANARCH("!EVOLMEN")
       Y$ = REGBLAN$("!EVOLMEN")
       J& = 0
       FEX = DESDEX%
       Des% = COMIENFE%(FEX)
       Has% = HASTAX%
       For k& = 1 To URE&
          '
          If DESDEI%(k&) >= Des% And DESDEI%(k&) <= Has% Then
             PER$ = PERIOD$(k&)
             TVEN# = VENTA#(k&)
             TCOM# = COMPRA#(k&)
             TCOB# = COBRA#(k&)
             TPAG# = PAGOS#(k&)
             RES1# = 0: RES2# = 0
             '
             Call REPLA(Y$, PER$, 1, 16)
             Call REPLA(Y$, MKD$(TVEN#), 17, 8)
             Call REPLA(Y$, MKD$(TCOM#), 25, 8)
             RES1# = TVEN# - TCOM#
             Call REPLA(Y$, MKD$(RES1#), 33, 8)
             Call REPLA(Y$, MKD$(TCOB#), 41, 8)
             Call REPLA(Y$, MKD$(TPAG#), 49, 8)
             RES2# = TCOB# - TPAG#
             Call REPLA(Y$, MKD$(RES2#), 57, 8)
             Call REPLA(Y$, MKI$(DESDEI%(k&)), 65, 2)
             J& = J& + 1
             Call GRAREG("!EVOLMEN", Y$, J&)
          End If
          '
       Next k&
       '
       Call SETULTREG("!EVOLMEN", J&)
       Call TRACENTRAL("EVOLMEN", "")
       '
       Call FILESORT("EVOLMEN", "", 8, 8, "ASC")
       Call HEADERS("EVOLMEN", "")
       '
       Call FINAL("*EVOLMEN")       '
End Sub

Sub RESUOPER2()
      '
      Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
      If PERIO$ = "" Then Exit Sub
      '
      Dim PERIOD$(), DESDEI%(), HASTAI%(), VENTA#(), COMPRA#(), COBRA#(), PAGOS#()
      '
      'TODO ESTO VA A VER QUE DESCOMENTAR PARA JYM
      'AGREGADO PAA ACTUALIZAR URE& POR QUE SI NO EL VECTOR SOLO SE CARGABA CON 48 Y SI EL RANGO ERA MAS GRANDE NO TRAIA LAS FECHAS MAS VIEJAS
      FEDES = FECHATEX$(DESDEX%)
      FEHAS = FECHATEX$(HASTAX%)
      CANTIMESES% = 1 + DateDiff("M", FEDES, FEHAS)  ' CANTIDAD DE MESES COMPRENIDOS EN EL RANGO
      If CANTIMESES% < 1 Then Call COMUNI("Error de Ingreso de Rango de Fechas"): Exit Sub
'      FIN TODO ESTO VA A VER QUE DESCOMENTAR PARA JYM
      
      Call CARSELMESA
      URE& = ULTREG&("SELMESA")
      If URE& < CANTIMESES% Then URE& = CANTIMESES% ' CODIGO PARA EL CAMBIO EN JYM

      ReDim PERIOD$(URE&)
      ReDim DESDEI%(URE&)
      ReDim HASTAI%(URE&)
      ReDim VENTA#(URE&)
      ReDim COMPRA#(URE&)
      ReDim COBRA#(URE&)
      ReDim PAGOS#(URE&)
      '
      i& = 0
'      For i& = 1 To URE&
'          RELE$ = REGLEIDO$("SELMESA", i&)
'          PERIOD$(i&) = MESANO$(RELE$, Des%, Has%)
'          DESDEI%(i&) = Des%
'          HASTAI%(i&) = Has%
'      Next i&
      '
'===================================
      For i& = 1 To URE&
          If i& = 1 Then
             MESINICIO = Month(FEDES)
             ANO = Year(FEDES)
             NOMBREMES$ = UCase$(MonthName(MESINICIO)) 'nombre del mes
             MESX$ = FORMATOCON0$(SAFETEXT$(MESINICIO), 2)
             ANOX$ = SAFETEXT$(Right(ANO, 2))
             Des% = FECHANUM("01/" & MESX$ & "/" & ANOX$)
             Has% = FECHANUM("& NumeroFinDelMes%(des%) & " & MESX$ & "/" & ANOX$)
             'FECHANUM(FinDelMes(Des%))NumeroFinDelMes%(des%)
             If Des% < 1 Or Has% < 1 Then
               A = 1
             End If
          Else
             MESINICIO = MESINICIO + 1
             If MESINICIO > 12 Then
               MESINICIO = 1
               ANO = ANO + 1
             End If
             NOMBREMES$ = UCase$(MonthName(MESINICIO))
             MESX$ = FORMATOCON0$(SAFETEXT$(MESINICIO), 2)
             ANOX$ = SAFETEXT$(Right(ANO, 2))
             Des% = FECHANUM("01/" & MESX$ & "/" & ANOX$)
             Has% = FECHANUM("& NumeroFinDelMes%(des%) & " & MESX$ & "/" & ANOX$)
             'Has% = FECHANUM(FinDelMes(Des%))
               
          End If
          
          PERIOD$(i&) = NOMBREMES$ & " " & SAFETEXT$(Right(ANO, 2))
          DESDEI%(i&) = Des%
          HASTAI%(i&) = Has%
          
'          If DESDEX% > Des% And DESDEX% < Has% Then
'             PERIOD$(i&) = PERIOD$(i&) + " (Inc.)"
'             DESDEI%(i&) = DESDEX%
'          End If
'          If HASTAX% > Des% And HASTAX% < Has% Then
'             PERIOD$(i&) = PERIOD$(i&) + " (Inc.)"
'             HASTAI%(i&) = HASTAX%
'          End If
          '
      Next i&
      '

'===================================      Call APERBASDAT("CABAN")
      '
      'PARA OBTENER VENTAS NETAS
      SQLX$ = "SELECT IDEBECOMP, IHABECOMP, FECHCONTA FROM CAJABANC WITH(NOLOCK) "
      SQLX$ = SQLX$ + " WHERE TipoMovim = 'FVEN'"
      SQLX$ = SQLX$ + " AND  NuClien > 0 "
      '
'      'Set AVANCE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'      Dim AVANCE As ADODB.Recordset
'      Set AVANCE = New ADODB.Recordset
'25    On Error Resume Next
'      AVANCE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'           On Error GoTo 0
'           Call CapturaErrorOpen
'           GoTo 25
'      End If
'      On Error GoTo 0
      '
      Set AVANCE = READSET(SQLX$)
      With AVANCE
          On Error Resume Next
          If (.EOF And .BOF) Then
             On Error GoTo 0
             Screen.MousePointer = 1
             GoTo 10
          End If
          On Error GoTo 0
          '
          Do While Not .EOF
               IMPO# = XVALO(!IDEBECOMP) - XVALO(!IHABECOMP)
               FFI% = CVINT%(!fechconta)
               '
               J& = 0
               For J& = 1 To URE&
                   If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                      VENTA#(J&) = VENTA#(J&) + IMPO#
                      Exit For
                   End If
               Next J&
          .MoveNext
          Loop
          '
      End With
      
      '
      'PARA OBTENER COBRANZAS NETAS
10     SQLX$ = "SELECT IHABECOMP, IDEBECOMP, FECHCONTA FROM CAJABANC WITH(NOLOCK)"
       SQLX$ = SQLX$ + "WHERE TipoMovim = 'RCOB'"
       SQLX$ = SQLX$ + " AND  NuClien > 0 "
       '
'       'Set AVANCE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'       Set AVANCE = New ADODB.Recordset
'26     On Error Resume Next
'       AVANCE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'           On Error GoTo 0
'           Call CapturaErrorOpen
'           GoTo 26
'       End If
'       On Error GoTo 0
       Set AVANCE = READSET(SQLX$)

       With AVANCE
           On Error Resume Next
           If (.EOF And .BOF) Then
              On Error GoTo 0
              Screen.MousePointer = 1
              GoTo 20
           End If
           On Error GoTo 0
          '
           Do While Not .EOF
                IMPO# = XVALO(!IHABECOMP) - XVALO(!IDEBECOMP)
                FFI% = CVINT%(!fechconta)
                '
                J& = 0
                For J& = 1 To URE&
                    If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                       COBRA#(J&) = COBRA#(J&) + IMPO#
                       Exit For
                    End If
                Next J&
           .MoveNext
           Loop
           '
       End With
       '
       'PARA OBTENER COMPRAS NETAS
20     SQLX$ = "SELECT IHABECOMP, IDEBECOMP, FECHCONTA FROM CAJABANC WITH(NOLOCK)"
       SQLX$ = SQLX$ + "WHERE TipoMovim = 'FCOM'"
       SQLX$ = SQLX$ + "AND  NuProve > 0 "
       '
'       'Set AVANCE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'       Set AVANCE = New ADODB.Recordset
'27     On Error Resume Next
'       AVANCE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'           On Error GoTo 0
'           Call CapturaErrorOpen
'           GoTo 27
'       End If
'       On Error GoTo 0
       Set AVANCE = READSET(SQLX$)
       With AVANCE
           On Error Resume Next
           .MoveFirst
           If Err Then
              On Error GoTo 0
              Screen.MousePointer = 1
              GoTo 30
           End If
           '
           Do While Not .EOF
                IMPO# = XVALO(!IHABECOMP) - XVALO(!IDEBECOMP)
                FFI% = CVINT%(!fechconta)
                '
                J& = 0
                For J& = 1 To URE&
                    If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                       COMPRA#(J&) = COMPRA#(J&) + IMPO#
                       Exit For
                    End If
                Next J&
           .MoveNext
           Loop
           '
       End With
       '
       'PARA OBTENER PAGOS NETOS
30     SQLX$ = "SELECT IDEBECOMP, IHABECOMP, FECHCONTA FROM CAJABANC WITH(NOLOCK)"
       SQLX$ = SQLX$ + "WHERE TipoMovim = 'OPAG'"
       SQLX$ = SQLX$ + "AND  NuProve > 0 "
       '
'       'Set AVANCE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'       Set AVANCE = New ADODB.Recordset
'28     On Error Resume Next
'       AVANCE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'           On Error GoTo 0
'           Call CapturaErrorOpen
'           GoTo 28
'       End If
'       On Error GoTo 0
       '
       Set AVANCE = READSET(SQLX$)
       With AVANCE
           On Error Resume Next
           .MoveFirst
           If Err Then
              On Error GoTo 0
              Screen.MousePointer = 1
              GoTo 40
           End If
          '
           Do While Not .EOF
                IMPO# = XVALO(!IDEBECOMP) - XVALO(!IHABECOMP)
                FFI% = CVINT%(!fechconta)
                '
                J& = 0
                For J& = 1 To URE&
                    If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                       PAGOS#(J&) = PAGOS#(J&) + IMPO#
                       Exit For
                    End If
              Next J&
           .MoveNext
           '
           Loop
           '
       End With
       '
40     Call BLANARCH("!EVOLMEN")
       Y$ = REGBLAN$("!EVOLMEN")
       J& = 0
       FEX = DESDEX%
       Des% = COMIENFE%(FEX)
       Has% = HASTAX%
       For k& = 1 To URE&
          '
          If DESDEI%(k&) >= Des% And DESDEI%(k&) <= Has% Then
             PER$ = PERIOD$(k&)
             TVEN# = VENTA#(k&)
             TCOM# = COMPRA#(k&)
             TCOB# = COBRA#(k&)
             TPAG# = PAGOS#(k&)
             RES1# = 0: RES2# = 0
             '
             Call REPLA(Y$, PER$, 1, 16)
             Call REPLA(Y$, MKD$(TVEN#), 17, 8)
             Call REPLA(Y$, MKD$(TCOM#), 25, 8)
             RES1# = TVEN# - TCOM#
             Call REPLA(Y$, MKD$(RES1#), 33, 8)
             Call REPLA(Y$, MKD$(TCOB#), 41, 8)
             Call REPLA(Y$, MKD$(TPAG#), 49, 8)
             RES2# = TCOB# - TPAG#
             Call REPLA(Y$, MKD$(RES2#), 57, 8)
             Call REPLA(Y$, MKI$(DESDEI%(k&)), 65, 2)
             J& = J& + 1
             Call GRAREG("!EVOLMEN", Y$, J&)
          End If
          '
       Next k&
       '
       Call SETULTREG("!EVOLMEN", J&)
       Call TRACENTRAL("EVOLMEN", "")
       '
       Call FILESORT("EVOLMEN", "", 8, 8, "ASC")
       Call HEADERS("EVOLMEN", "")
       '
       Call FINAL("*EVOLMEN")       '
End Sub




Sub RESOPEGASNuevo()
      '
      'RESUMEN OPERATIVO INCLUYE GASTOS (J Y M )
      '
      Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
      If PERIO$ = "" Then Exit Sub
      '
      Dim PERIOD$(), DESDEI%(), HASTAI%(), VENTA#(), COMPRA#(), COBRA#(), PAGOS#(), GASTOS#()
      '
      'TODO ESTO VA A VER QUE DESCOMENTAR PARA JYM
      'AGREGADO PAA ACTUALIZAR URE& POR QUE SI NO EL VECTOR SOLO SE CARGABA CON 48 Y SI EL RANGO ERA MAS GRANDE NO TRAIA LAS FECHAS MAS VIEJAS
      FEDES = FECHATEX$(DESDEX%)
      FEHAS = FECHATEX$(HASTAX%)
      CANTIMESES% = 1 + DateDiff("M", FEDES, FEHAS)  ' CANTIDAD DE MESES COMPRENIDOS EN EL RANGO
      If CANTIMESES% < 1 Then Call COMUNI("Error de Ingreso de Rango de Fechas"): Exit Sub
'      FIN TODO ESTO VA A VER QUE DESCOMENTAR PARA JYM

      Call CARSELMESA
      URE& = ULTREG&("SELMESA")
      If URE& < CANTIMESES% Then URE& = CANTIMESES% ' CODIGO PARA EL CAMBIO EN JYM
      ReDim PERIOD$(URE&)
      ReDim DESDEI%(URE&)
      ReDim HASTAI%(URE&)
      ReDim VENTA#(URE&)
      ReDim COMPRA#(URE&)
      ReDim COBRA#(URE&)
      ReDim PAGOS#(URE&)
      ReDim GASTOS#(URE&)
      '
      i& = 0
      Screen.MousePointer = 11
      
'      For i& = 1 To URE&
'          RELE$ = REGLEIDO$("SELMESA", i&)
'          PERIOD$(i&) = MESANO$(RELE$, Des%, Has%)
'          DESDEI%(i&) = Des%
'          HASTAI%(i&) = Has%
'          '
'          If DESDEX% > Des% And DESDEX% < Has% Then
'             PERIOD$(i&) = MESANO$(RELE$, Des%, Has%) + " (Inc.)"
'             DESDEI%(i&) = DESDEX%
'          End If
'          If HASTAX% > Des% And HASTAX% < Has% Then
'             PERIOD$(i&) = MESANO$(RELE$, Des%, Has%) + " (Inc.)"
'             HASTAI%(i&) = HASTAX%
'          End If
'          '
'      Next i&
      
'===================================
      For i& = 1 To URE&
          If i& = 1 Then
             MESINICIO = Month(FEDES)
             ANO = Year(FEDES)
             NOMBREMES$ = UCase$(MonthName(MESINICIO)) 'nombre del mes
             MESX$ = FORMATOCON0$(SAFETEXT$(MESINICIO), 2)
             ANOX$ = SAFETEXT$(Right(ANO, 2))
             Des% = FECHANUM("01/" & MESX$ & "/" & ANOX$)
             'Has% = FECHANUM(FinDelMes(Des%))
             Has% = FECHANUM("& NumeroFinDelMes%(des%) & " & MESX$ & "/" & ANOX$)
          Else
             MESINICIO = MESINICIO + 1
             If MESINICIO > 12 Then
               MESINICIO = 1
               ANO = ANO + 1
             End If
             NOMBREMES$ = UCase$(MonthName(MESINICIO))
             MESX$ = FORMATOCON0$(SAFETEXT$(MESINICIO), 2)
             ANOX$ = SAFETEXT$(Right(ANO, 2))
             Des% = FECHANUM("01/" & MESX$ & "/" & ANOX$)
             'Has% = FECHANUM(FinDelMes(Des%))
             Has% = FECHANUM("& NumeroFinDelMes%(des%) & " & MESX$ & "/" & ANOX$)
               
          End If
          
          PERIOD$(i&) = NOMBREMES$ & " " & SAFETEXT$(Right(ANO, 2))
          DESDEI%(i&) = Des%
          HASTAI%(i&) = Has%
          
          If DESDEX% > Des% And DESDEX% < Has% Then
             PERIOD$(i&) = PERIOD$(i&) + " (Inc.)"
             DESDEI%(i&) = DESDEX%
          End If
          If HASTAX% > Des% And HASTAX% < Has% Then
             PERIOD$(i&) = PERIOD$(i&) + " (Inc.)"
             HASTAI%(i&) = HASTAX%
          End If
          '
      Next i&
      '

'===================================
      Call APERBASDAT("CABAN")
      '
      'PARA OBTENER VENTAS NETAS
      SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
      SQLX$ = SQLX$ + "WHERE TipoMovim = 'FVEN'"
      SQLX$ = SQLX$ + "AND  NuClien > 0 "
      '
      Dim AVANCE  As New ADODB.Recordset
      Set AVANCE = FLEXCONN.Execute(FILTSQL$(SQLX$))
      '
      With AVANCE
          '
          Do While Not .EOF
               IMPO# = XVALO(!IDEBECOMP) - XVALO(!IHABECOMP)
               FFI% = CVINT%(!fechconta)
               '
               J& = 0
               For J& = 1 To URE&
                   If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                      VENTA#(J&) = VENTA#(J&) + IMPO#
                      Exit For
                   End If
               Next J&
          .MoveNext
          Loop
          '
      End With
      '
      AVANCE.Close
      Set AVANCE = Nothing
      'PARA OBTENER COBRANZAS NETAS
10     SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
       SQLX$ = SQLX$ + "WHERE TipoMovim = 'RCOB'"
       SQLX$ = SQLX$ + "AND  NuClien > 0 "
       '
       Dim AVANCE1  As New ADODB.Recordset
       Set AVANCE1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
        
       With AVANCE1
          '
           Do While Not .EOF
                IMPO# = XVALO(!IHABECOMP) - XVALO(!IDEBECOMP)
                FFI% = CVINT%(!fechconta)
                '
                J& = 0
                For J& = 1 To URE&
                    If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                       COBRA#(J&) = COBRA#(J&) + IMPO#
                       Exit For
                    End If
                Next J&
           .MoveNext
           Loop
           '
       End With
       '
       AVANCE1.Close
       Set AVANCE1 = Nothing

       'PARA OBTENER COMPRAS NETAS
20     SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
       SQLX$ = SQLX$ + "WHERE TipoMovim = 'FCOM'"
       SQLX$ = SQLX$ + "AND  NuProve > 0 "
       '
       Dim AVANCE2  As New ADODB.Recordset
       Set AVANCE2 = FLEXCONN.Execute(FILTSQL$(SQLX$))

       With AVANCE2
           '
           Do While Not .EOF
                IMPO# = XVALO(!IHABECOMP) - XVALO(!IDEBECOMP)
                FFI% = CVINT%(!fechconta)
                '
                J& = 0
                For J& = 1 To URE&
                    If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                       COMPRA#(J&) = COMPRA#(J&) + IMPO#
                       Exit For
                    End If
                Next J&
           .MoveNext
           Loop
           '
       End With
       '
       AVANCE2.Close
       Set AVANCE2 = Nothing

       'PARA OBTENER PAGOS NETOS
30     SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
       SQLX$ = SQLX$ + "WHERE TipoMovim = 'OPAG'"
       SQLX$ = SQLX$ + "AND  NuProve > 0 "
       '
       Dim AVANCE3  As New ADODB.Recordset
       Set AVANCE3 = FLEXCONN.Execute(FILTSQL$(SQLX$))
       '
       With AVANCE3
          '
           Do While Not .EOF
            IMPO# = XVALO(!IDEBECOMP) - XVALO(!IHABECOMP)
            FFI% = CVINT%(!fechconta)
            '
            J& = 0
            For J& = 1 To URE&
                 If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                    If TICUEPRO%((!NuProve) * (-1)) = 0 Then
                        PAGOS#(J&) = PAGOS#(J&) + IMPO#
                      Else
                        GASTOS#(J&) = GASTOS#(J&) + IMPO#
                    End If
                    Exit For
                 End If
            Next J&
           .MoveNext
           '
           Loop
           '
       End With
       '
       AVANCE3.Close
       Set AVANCE3 = Nothing
       
       Call BLANARCH("!RESUOPER")
       Y$ = REGBLAN$("!RESUOPER")
       J& = 0
       FEX = DESDEX%
       Des% = COMIENFE%(FEX)
       Has% = HASTAX%
       '
       For k& = 1 To URE&
          '
          '
          If DESDEI%(k&) >= Des% And DESDEI%(k&) <= Has% Then
             PER$ = PERIOD$(k&)
             TVEN# = VENTA#(k&)
             TCOM# = COMPRA#(k&)
             TCOB# = COBRA#(k&)
             TPAG# = PAGOS#(k&)
             TGAS# = GASTOS#(k&)
             RES1# = 0: RES2# = 0
             '
             'INCLUYE GASTOS (RESUOPER)
             Call REPLA(Y$, PER$, 1, 16)
             Call REPLA(Y$, MKD$(TVEN#), 17, 8)
             Call REPLA(Y$, MKD$(TCOM#), 25, 8)
             Call REPLA(Y$, MKD$(TGAS#), 33, 8)
             RES1# = TVEN# - TCOM# - TGAS#
             Call REPLA(Y$, MKD$(RES1#), 41, 8)
             Call REPLA(Y$, MKD$(TCOB#), 49, 8)
             Call REPLA(Y$, MKD$(TPAG#), 57, 8)
             'Call REPLA(Y$, MKD$(TGAS#), 65, 8) ' SE QUITA ESTA COLUMNA A PEDIDO DE J Y M
             RES2# = TCOB# - TPAG# '- TGAS#
             Call REPLA(Y$, MKD$(RES2#), 73, 8)
             Call REPLA(Y$, MKI$(DESDEI%(k&)), 81, 2)
             J& = J& + 1
             Call GRAREG("!RESUOPER", Y$, J&)
          End If
          '
       Next k&
       '
       Call SETULTREG("!RESUOPER", J&)
       Call TRACENTRAL("RESUOPER", "")
       Call FILESORT("RESUOPER", "", 10, 10, "DESC")
       Call CIERRARCH("RESUOPER")

       Call HEADERS("RESUOPER", "")
       Call HEADERS("RESUOPER", "Corresponde a: " + PERIO$)
       '
       Screen.MousePointer = 1
       '
       Call FINAL("*RESUOPER")
       '
End Sub

Sub RESUOPER()
      '
      Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
      If PERIO$ = "" Then Exit Sub
      '
      Dim PERIOD$(), DESDEI%(), HASTAI%(), VENTA#(), COMPRA#(), COBRA#(), PAGOS#()
      '
      Call CARSELMESA
      URE& = ULTREG&("SELMESA")
      ReDim PERIOD$(URE&)
      ReDim DESDEI%(URE&)
      ReDim HASTAI%(URE&)
      ReDim VENTA#(URE&)
      ReDim COMPRA#(URE&)
      ReDim COBRA#(URE&)
      ReDim PAGOS#(URE&)
      '
      i& = 0
      For i& = 1 To URE&
          RELE$ = REGLEIDO$("SELMESA", i&)
          PERIOD$(i&) = MESANO$(RELE$, Des%, Has%)
          DESDEI%(i&) = Des%
          HASTAI%(i&) = Has%
      Next i&
      '
      Call APERBASDAT("CABAN")
      '
      'PARA OBTENER VENTAS NETAS
      SQLX$ = "SELECT IDEBECOMP, IHABECOMP, FECHCONTA FROM CAJABANC WITH(NOLOCK) "
      SQLX$ = SQLX$ + " WHERE TipoMovim = 'FVEN'"
      SQLX$ = SQLX$ + " AND  NuClien > 0 "
      '
'      'Set AVANCE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'      Dim AVANCE As ADODB.Recordset
'      Set AVANCE = New ADODB.Recordset
'25    On Error Resume Next
'      AVANCE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'           On Error GoTo 0
'           Call CapturaErrorOpen
'           GoTo 25
'      End If
'      On Error GoTo 0
      '
      Set AVANCE = READSET(SQLX$)
      With AVANCE
          On Error Resume Next
          If (.EOF And .BOF) Then
             On Error GoTo 0
             Screen.MousePointer = 1
             GoTo 10
          End If
          On Error GoTo 0
          '
          Do While Not .EOF
               IMPO# = XVALO(!IDEBECOMP) - XVALO(!IHABECOMP)
               FFI% = CVINT%(!fechconta)
               '
               J& = 0
               For J& = 1 To URE&
                   If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                      VENTA#(J&) = VENTA#(J&) + IMPO#
                      Exit For
                   End If
               Next J&
          .MoveNext
          Loop
          '
      End With
      
      '
      'PARA OBTENER COBRANZAS NETAS
10     SQLX$ = "SELECT IHABECOMP, IDEBECOMP, FECHCONTA FROM CAJABANC WITH(NOLOCK)"
       SQLX$ = SQLX$ + "WHERE TipoMovim = 'RCOB'"
       SQLX$ = SQLX$ + " AND  NuClien > 0 "
       '
'       'Set AVANCE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'       Set AVANCE = New ADODB.Recordset
'26     On Error Resume Next
'       AVANCE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'           On Error GoTo 0
'           Call CapturaErrorOpen
'           GoTo 26
'       End If
'       On Error GoTo 0
       Set AVANCE = READSET(SQLX$)

       With AVANCE
           On Error Resume Next
           If (.EOF And .BOF) Then
              On Error GoTo 0
              Screen.MousePointer = 1
              GoTo 20
           End If
           On Error GoTo 0
          '
           Do While Not .EOF
                IMPO# = XVALO(!IHABECOMP) - XVALO(!IDEBECOMP)
                FFI% = CVINT%(!fechconta)
                '
                J& = 0
                For J& = 1 To URE&
                    If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                       COBRA#(J&) = COBRA#(J&) + IMPO#
                       Exit For
                    End If
                Next J&
           .MoveNext
           Loop
           '
       End With
       '
       'PARA OBTENER COMPRAS NETAS
20     SQLX$ = "SELECT IHABECOMP, IDEBECOMP, FECHCONTA FROM CAJABANC WITH(NOLOCK)"
       SQLX$ = SQLX$ + "WHERE TipoMovim = 'FCOM'"
       SQLX$ = SQLX$ + "AND  NuProve > 0 "
       '
'       'Set AVANCE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'       Set AVANCE = New ADODB.Recordset
'27     On Error Resume Next
'       AVANCE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'           On Error GoTo 0
'           Call CapturaErrorOpen
'           GoTo 27
'       End If
'       On Error GoTo 0
       Set AVANCE = READSET(SQLX$)
       With AVANCE
           On Error Resume Next
           .MoveFirst
           If Err Then
              On Error GoTo 0
              Screen.MousePointer = 1
              GoTo 30
           End If
           '
           Do While Not .EOF
                IMPO# = XVALO(!IHABECOMP) - XVALO(!IDEBECOMP)
                FFI% = CVINT%(!fechconta)
                '
                J& = 0
                For J& = 1 To URE&
                    If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                       COMPRA#(J&) = COMPRA#(J&) + IMPO#
                       Exit For
                    End If
                Next J&
           .MoveNext
           Loop
           '
       End With
       '
       'PARA OBTENER PAGOS NETOS
30     SQLX$ = "SELECT IDEBECOMP, IHABECOMP, FECHCONTA FROM CAJABANC WITH(NOLOCK)"
       SQLX$ = SQLX$ + "WHERE TipoMovim = 'OPAG'"
       SQLX$ = SQLX$ + "AND  NuProve > 0 "
       '
'       'Set AVANCE = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'       Set AVANCE = New ADODB.Recordset
'28     On Error Resume Next
'       AVANCE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'           On Error GoTo 0
'           Call CapturaErrorOpen
'           GoTo 28
'       End If
'       On Error GoTo 0
       '
       Set AVANCE = READSET(SQLX$)
       With AVANCE
           On Error Resume Next
           .MoveFirst
           If Err Then
              On Error GoTo 0
              Screen.MousePointer = 1
              GoTo 40
           End If
          '
           Do While Not .EOF
                IMPO# = XVALO(!IDEBECOMP) - XVALO(!IHABECOMP)
                FFI% = CVINT%(!fechconta)
                '
                J& = 0
                For J& = 1 To URE&
                    If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                       PAGOS#(J&) = PAGOS#(J&) + IMPO#
                       Exit For
                    End If
              Next J&
           .MoveNext
           '
           Loop
           '
       End With
       '
40     Call BLANARCH("!EVOLMEN")
       Y$ = REGBLAN$("!EVOLMEN")
       J& = 0
       FEX = DESDEX%
       Des% = COMIENFE%(FEX)
       Has% = HASTAX%
       For k& = 1 To URE&
          '
          If DESDEI%(k&) >= Des% And DESDEI%(k&) <= Has% Then
             PER$ = PERIOD$(k&)
             TVEN# = VENTA#(k&)
             TCOM# = COMPRA#(k&)
             TCOB# = COBRA#(k&)
             TPAG# = PAGOS#(k&)
             RES1# = 0: RES2# = 0
             '
             Call REPLA(Y$, PER$, 1, 16)
             Call REPLA(Y$, MKD$(TVEN#), 17, 8)
             Call REPLA(Y$, MKD$(TCOM#), 25, 8)
             RES1# = TVEN# - TCOM#
             Call REPLA(Y$, MKD$(RES1#), 33, 8)
             Call REPLA(Y$, MKD$(TCOB#), 41, 8)
             Call REPLA(Y$, MKD$(TPAG#), 49, 8)
             RES2# = TCOB# - TPAG#
             Call REPLA(Y$, MKD$(RES2#), 57, 8)
             Call REPLA(Y$, MKI$(DESDEI%(k&)), 65, 2)
             J& = J& + 1
             Call GRAREG("!EVOLMEN", Y$, J&)
          End If
          '
       Next k&
       '
       Call SETULTREG("!EVOLMEN", J&)
       Call TRACENTRAL("EVOLMEN", "")
       '
       Call FILESORT("EVOLMEN", "", 8, 8, "ASC")
       Call HEADERS("EVOLMEN", "")
       '
       Call FINAL("*EVOLMEN")       '
End Sub

Sub RESOPEGAS()
      '
      'RESUMEN OPERATIVO INCLUYE GASTOS (J Y M )
      '
      Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
      If PERIO$ = "" Then Exit Sub
      '
      Dim PERIOD$(), DESDEI%(), HASTAI%(), VENTA#(), COMPRA#(), COBRA#(), PAGOS#(), GASTOS#()
      '
      Call CARSELMESA
      URE& = ULTREG&("SELMESA")
      
      ReDim PERIOD$(URE&)
      ReDim DESDEI%(URE&)
      ReDim HASTAI%(URE&)
      ReDim VENTA#(URE&)
      ReDim COMPRA#(URE&)
      ReDim COBRA#(URE&)
      ReDim PAGOS#(URE&)
      ReDim GASTOS#(URE&)
      '
      i& = 0
      Screen.MousePointer = 11
      For i& = 1 To URE&
          RELE$ = REGLEIDO$("SELMESA", i&)
          PERIOD$(i&) = MESANO$(RELE$, Des%, Has%)
          DESDEI%(i&) = Des%
          HASTAI%(i&) = Has%
          '
          If DESDEX% > Des% And DESDEX% < Has% Then
             PERIOD$(i&) = MESANO$(RELE$, Des%, Has%) + " (Inc.)"
             DESDEI%(i&) = DESDEX%
          End If
          If HASTAX% > Des% And HASTAX% < Has% Then
             PERIOD$(i&) = MESANO$(RELE$, Des%, Has%) + " (Inc.)"
             HASTAI%(i&) = HASTAX%
          End If
          '
      Next i&
      '
      Call APERBASDAT("CABAN")
      '
      'PARA OBTENER VENTAS NETAS
      SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
      SQLX$ = SQLX$ + "WHERE TipoMovim = 'FVEN'"
      SQLX$ = SQLX$ + "AND  NuClien > 0 "
      '
      Dim AVANCE  As New ADODB.Recordset
      Set AVANCE = FLEXCONN.Execute(FILTSQL$(SQLX$))
      '
      With AVANCE
          '
          Do While Not .EOF
               IMPO# = XVALO(!IDEBECOMP) - XVALO(!IHABECOMP)
               FFI% = CVINT%(!fechconta)
               '
               J& = 0
               For J& = 1 To URE&
                   If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                      VENTA#(J&) = VENTA#(J&) + IMPO#
                      Exit For
                   End If
               Next J&
          .MoveNext
          Loop
          '
      End With
      '
      AVANCE.Close
      Set AVANCE = Nothing
      'PARA OBTENER COBRANZAS NETAS
10     SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
       SQLX$ = SQLX$ + "WHERE TipoMovim = 'RCOB'"
       SQLX$ = SQLX$ + "AND  NuClien > 0 "
       '
       Dim AVANCE1  As New ADODB.Recordset
       Set AVANCE1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
        
       With AVANCE1
          '
           Do While Not .EOF
                IMPO# = XVALO(!IHABECOMP) - XVALO(!IDEBECOMP)
                FFI% = CVINT%(!fechconta)
                '
                J& = 0
                For J& = 1 To URE&
                    If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                       COBRA#(J&) = COBRA#(J&) + IMPO#
                       Exit For
                    End If
                Next J&
           .MoveNext
           Loop
           '
       End With
       '
       AVANCE1.Close
       Set AVANCE1 = Nothing

       'PARA OBTENER COMPRAS NETAS
20     SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
       SQLX$ = SQLX$ + "WHERE TipoMovim = 'FCOM'"
       SQLX$ = SQLX$ + "AND  NuProve > 0 "
       '
       Dim AVANCE2  As New ADODB.Recordset
       Set AVANCE2 = FLEXCONN.Execute(FILTSQL$(SQLX$))

       With AVANCE2
           '
           Do While Not .EOF
                IMPO# = XVALO(!IHABECOMP) - XVALO(!IDEBECOMP)
                FFI% = CVINT%(!fechconta)
                '
                J& = 0
                For J& = 1 To URE&
                    If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                       COMPRA#(J&) = COMPRA#(J&) + IMPO#
                       Exit For
                    End If
                Next J&
           .MoveNext
           Loop
           '
       End With
       '
       AVANCE2.Close
       Set AVANCE2 = Nothing

       'PARA OBTENER PAGOS NETOS
30     SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
       SQLX$ = SQLX$ + "WHERE TipoMovim = 'OPAG'"
       SQLX$ = SQLX$ + "AND  NuProve > 0 "
       '
       Dim AVANCE3  As New ADODB.Recordset
       Set AVANCE3 = FLEXCONN.Execute(FILTSQL$(SQLX$))
       '
       With AVANCE3
          '
           Do While Not .EOF
            IMPO# = XVALO(!IDEBECOMP) - XVALO(!IHABECOMP)
            FFI% = CVINT%(!fechconta)
            '
            J& = 0
            For J& = 1 To URE&
                 If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
                    If TICUEPRO%((!NuProve) * (-1)) = 0 Then
                        PAGOS#(J&) = PAGOS#(J&) + IMPO#
                      Else
                        GASTOS#(J&) = GASTOS#(J&) + IMPO#
                    End If
                    Exit For
                 End If
            Next J&
           .MoveNext
           '
           Loop
           '
       End With
       '
       AVANCE3.Close
       Set AVANCE3 = Nothing
       
       Call BLANARCH("!RESUOPER")
       Y$ = REGBLAN$("!RESUOPER")
       J& = 0
       FEX = DESDEX%
       Des% = COMIENFE%(FEX)
       Has% = HASTAX%
       '
       For k& = 1 To URE&
          '
          If DESDEI%(k&) >= Des% And DESDEI%(k&) <= Has% Then
             PER$ = PERIOD$(k&)
             TVEN# = VENTA#(k&)
             TCOM# = COMPRA#(k&)
             TCOB# = COBRA#(k&)
             TPAG# = PAGOS#(k&)
             TGAS# = GASTOS#(k&)
             RES1# = 0: RES2# = 0
             '
             'INCLUYE GASTOS (RESUOPER)
             Call REPLA(Y$, PER$, 1, 16)
             Call REPLA(Y$, MKD$(TVEN#), 17, 8)
             Call REPLA(Y$, MKD$(TCOM#), 25, 8)
             Call REPLA(Y$, MKD$(TGAS#), 33, 8)
             RES1# = TVEN# - TCOM# - TGAS#
             Call REPLA(Y$, MKD$(RES1#), 41, 8)
             Call REPLA(Y$, MKD$(TCOB#), 49, 8)
             Call REPLA(Y$, MKD$(TPAG#), 57, 8)
             'Call REPLA(Y$, MKD$(TGAS#), 65, 8) ' SE QUITA ESTA COLUMNA A PEDIDO DE J Y M
             RES2# = TCOB# - TPAG# '- TGAS#
             Call REPLA(Y$, MKD$(RES2#), 73, 8)
             Call REPLA(Y$, MKI$(DESDEI%(k&)), 81, 2)
             J& = J& + 1
             Call GRAREG("!RESUOPER", Y$, J&)
          End If
          '
       Next k&
       '
       Call SETULTREG("!RESUOPER", J&)
       Call TRACENTRAL("RESUOPER", "")
       Call FILESORT("RESUOPER", "", 10, 10, "DESC")
       Call CIERRARCH("RESUOPER")

       Call HEADERS("RESUOPER", "")
       Call HEADERS("RESUOPER", "Corresponde a: " + PERIO$)
       '
       Screen.MousePointer = 1
       '
       Call FINAL("*RESUOPER")
       '
End Sub

Sub RESOPEGAS2()
      '
      'RESUMEN OPERATIVO INCLUYE GASTOS (J Y M )
      '
      Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
      If PERIO$ = "" Then Exit Sub
      '
      Dim PERIOD$(), DESDEI%(), HASTAI%(), VENTA#(), COMPRA#(), COBRA#(), PAGOS#(), GASTOS#()
      '
      'TODO ESTO VA A VER QUE DESCOMENTAR PARA JYM
      'AGREGADO PAA ACTUALIZAR URE& POR QUE SI NO EL VECTOR SOLO SE CARGABA CON 48 Y SI EL RANGO ERA MAS GRANDE NO TRAIA LAS FECHAS MAS VIEJAS
      FEDES = FECHATEX$(DESDEX%)
      FEHAS = FECHATEX$(HASTAX%)
      CANTIMESES% = 1 + DateDiff("M", FEDES, FEHAS)  ' CANTIDAD DE MESES COMPRENIDOS EN EL RANGO
      If CANTIMESES% < 1 Then Call COMUNI("Error de Ingreso de Rango de Fechas"): Exit Sub
'      FIN TODO ESTO VA A VER QUE DESCOMENTAR PARA JYM
      
      'Call CARSELMESA
      URE& = CANTIMESES%
      'If URE& < CANTIMESES% Then URE& = CANTIMESES% ' CODIGO PARA EL CAMBIO EN JYM

      ReDim PERIOD$(URE&)
      ReDim DESDEI%(URE&)
      ReDim HASTAI%(URE&)
      ReDim VENTA#(URE&)
      ReDim COMPRA#(URE&)
      ReDim COBRA#(URE&)
      ReDim PAGOS#(URE&)
      ReDim GASTOS#(URE&)
      '
      i& = 0
      Screen.MousePointer = 11
'      For i& = 1 To URE&
'          RELE$ = REGLEIDO$("SELMESA", i&)
'          PERIOD$(i&) = MESANO$(RELE$, Des%, Has%)
'          DESDEI%(i&) = Des%
'          HASTAI%(i&) = Has%
'          '
'          If DESDEX% > Des% And DESDEX% < Has% Then
'             PERIOD$(i&) = MESANO$(RELE$, Des%, Has%) + " (Inc.)"
'             DESDEI%(i&) = DESDEX%
'          End If
'          If HASTAX% > Des% And HASTAX% < Has% Then
'             PERIOD$(i&) = MESANO$(RELE$, Des%, Has%) + " (Inc.)"
'             HASTAI%(i&) = HASTAX%
'          End If
'          '
'      Next i&
      '
'===================================
      For i& = 1 To URE&
          If i& = 1 Then
            
             MESINICIO = Month(FEDES)
             ANO = Year(FEDES)
             NOMBREMES$ = UCase$(MonthName(MESINICIO)) 'nombre del mes
             MESX$ = FORMATOCON0$(SAFETEXT$(MESINICIO), 2)
             ANOX$ = SAFETEXT$(Right(ANO, 2))
             Des% = FECHANUM("01/" & MESX$ & "/" & ANOX$)
             DIAFIN$ = SAFETEXT$(NumeroFinDelMes%(Des%))
             FECHAFIN$ = DIAFIN$ & "/" & MESX$ & "/" & ANOX$
             Has% = FECHANUM(FECHAFIN$)
          Else
             MESINICIO = MESINICIO + 1
             If MESINICIO > 12 Then
               MESINICIO = 1
               ANO = ANO + 1
             End If
             NOMBREMES$ = UCase$(MonthName(MESINICIO))
             MESX$ = FORMATOCON0$(SAFETEXT$(MESINICIO), 2)
             ANOX$ = SAFETEXT$(Right(ANO, 2))
             Des% = FECHANUM("01/" & MESX$ & "/" & ANOX$)
             DIAFIN$ = SAFETEXT$(NumeroFinDelMes%(Des%))
             FECHAFIN$ = DIAFIN$ & "/" & MESX$ & "/" & ANOX$
             Has% = FECHANUM(FECHAFIN$)
          End If
          
          PERIOD$(i&) = NOMBREMES$ & " " & SAFETEXT$(Right(ANO, 2))
          DESDEI%(i&) = Des%
          HASTAI%(i&) = Has%
          
          If DESDEX% > Des% And DESDEX% < Has% Then
             PERIOD$(i&) = PERIOD$(i&) + " (Inc.)"
             DESDEI%(i&) = DESDEX%
          End If
          If HASTAX% > Des% And HASTAX% < Has% Then
             PERIOD$(i&) = PERIOD$(i&) + " (Inc.)"
             HASTAI%(i&) = HASTAX%
          End If
          '
          If i& = 1 Then 'PRIMER MES
            Condi$ = " AND DAY(fechconta) >=  " & XVALO(Left$(FETEXCOR1$(DESDEX%), 2))
            Condi$ = Condi$ + " AND DAY(fechconta) <=  " & XVALO(Left$(FETEXCOR1$(Has%), 2))
            Condi$ = Condi$ + " AND  MONTH(fechconta) = " & XVALO(MESINICIO)
            Condi$ = Condi$ + " AND  YEAR(fechconta) = " & XVALO(ANO)
'            Condi$ = " AND  fechconta >= '" & FETEXCOR1$(DESDEX%) & "'"
'            Condi$ = Condi$ + " AND  fechconta <= '" & FETEXCOR1$(Has%) & "'"
          ElseIf i& = CANTIMESES% Then 'ULTIMO MES
'            Condi$ = " AND  fechconta >= '" & FETEXCOR1$(Des%) & "'"
'            Condi$ = Condi$ + " AND  fechconta <= '" & FETEXCOR1$(Has%) & "'"
            
            Condi$ = " AND DAY(fechconta) >=  " & XVALO(Left$(FETEXCOR1$(Des%), 2))
            Condi$ = Condi$ + " AND DAY(fechconta) <=  " & XVALO(Left$(FETEXCOR1$(HASTAX%), 2))
            Condi$ = Condi$ + " AND  MONTH(fechconta) = " & XVALO(MESINICIO)
            Condi$ = Condi$ + " AND  YEAR(fechconta) = " & XVALO(ANO)
            
          Else 'TODOS LOS MESES INTERMEDIOS
            Condi$ = " AND  MONTH(fechconta) = " & XVALO(MESINICIO)
            Condi$ = Condi$ + " AND  YEAR(fechconta) = " & XVALO(ANO)
          End If
          VENTA#(i&) = 0
          COMPRA#(i&) = 0
          COBRA#(i&) = 0
          PAGOS#(i&) = 0
          GASTOS#(i&) = 0
          
          'PARA OBTENER VENTAS NETAS
          SQLX$ = "SELECT SUM(IDEBECOMP) AS IDEBE ,SUM (IHABECOMP) AS IHABE FROM CAJABANC WITH(NOLOCK) "
          SQLX$ = SQLX$ + " WHERE TipoMovim = 'FVEN'"
          SQLX$ = SQLX$ + " AND  NuClien > 0 "
          SQLX$ = SQLX$ + Condi$
          Set AVANCE = READSET(SQLX$)
          With AVANCE
             IMPO# = XVALO(!IDEBE) - XVALO(!IHABE)
             VENTA#(i&) = IMPO#
          End With
          AVANCE.Close
          Set AVANCE = Nothing
          
          'PARA OBTENER COBRANZAS NETAS
          SQLX$ = "SELECT SUM(IDEBECOMP) AS IDEBE ,SUM (IHABECOMP) AS IHABE FROM CAJABANC WITH(NOLOCK) "
          SQLX$ = SQLX$ + " WHERE TipoMovim = 'RCOB'"
          SQLX$ = SQLX$ + " AND  NuClien > 0 "
          SQLX$ = SQLX$ + Condi$
          Set AVANCE1 = READSET(SQLX$)
          With AVANCE1
             IMPO# = XVALO(!IHABE) - XVALO(!IDEBE)
             COBRA#(i&) = IMPO#
          End With
          AVANCE1.Close
          Set AVANCE1 = Nothing
'MsgBox SQLX$
         'PARA OBTENER COMPRAS NETAS
          SQLX$ = "SELECT SUM(IDEBECOMP) AS IDEBE ,SUM (IHABECOMP) AS IHABE FROM CAJABANC WITH(NOLOCK) "
          SQLX$ = SQLX$ + " WHERE TipoMovim = 'FCOM'"
          SQLX$ = SQLX$ + " AND  NuProve  > 0 "
          SQLX$ = SQLX$ + Condi$
          Set AVANCE2 = READSET(SQLX$)
          With AVANCE2
             IMPO# = XVALO(!IHABE) - XVALO(!IDEBE)
             COMPRA#(i&) = IMPO#
          End With
          AVANCE2.Close
          Set AVANCE2 = Nothing
          
         'PARA OBTENER GASTOS NETOS
          SQLX$ = "SELECT SUM(C.IDEBECOMP) AS IDEBE ,SUM (C.IHABECOMP) AS IHABE FROM CAJABANC C WITH(NOLOCK) "
          SQLX$ = SQLX$ + " INNER JOIN PROVED12 P ON C.NuProve = P.Nume_Cli"
          SQLX$ = SQLX$ + " WHERE C.TipoMovim = 'OPAG'"
          SQLX$ = SQLX$ + " AND C.NuProve  > 0 AND P.Tcpr_Cli > 0 "
          SQLX$ = SQLX$ + Condi$
          Set AVANCE3 = READSET(SQLX$)
          With AVANCE3
             IMPO# = XVALO(!IDEBE) - XVALO(!IHABE)
             GASTOS#(i&) = IMPO#
             
          End With
          AVANCE3.Close
          Set AVANCE3 = Nothing
          '
         'PARA OBTENER PAGOS NETOS
          SQLX$ = "SELECT SUM(C.IDEBECOMP) AS IDEBE ,SUM (C.IHABECOMP) AS IHABE FROM CAJABANC C WITH(NOLOCK) "
          SQLX$ = SQLX$ + " INNER JOIN PROVED12 P ON C.NuProve = P.Nume_Cli"
          SQLX$ = SQLX$ + " WHERE C.TipoMovim = 'OPAG'"
          SQLX$ = SQLX$ + " AND C.NuProve  > 0 AND P.Tcpr_Cli < 1 "
          SQLX$ = SQLX$ + Condi$
          Set AVANCE4 = READSET(SQLX$)
          With AVANCE4
             IMPO# = XVALO(!IDEBE) - XVALO(!IHABE)
             PAGOS#(i&) = IMPO#
          End With
          AVANCE4.Close
          Set AVANCE4 = Nothing
      Next i&
      GoTo 150
'       Call BLANARCH("!RESUOPER")
'       Y$ = REGBLAN$("!RESUOPER")
'       J& = 0
'       FEX = DESDEX%
'       Des% = COMIENFE%(FEX)
'       Has% = HASTAX%
'       '
'       For k& = 1 To URE&
'          '
'          If DESDEI%(k&) >= Des% And DESDEI%(k&) <= Has% Then
'             PER$ = PERIOD$(k&)
'             TVEN# = VENTA#(k&)
'             TCOM# = COMPRA#(k&)
'             TCOB# = COBRA#(k&)
'             TPAG# = PAGOS#(k&)
'             TGAS# = GASTOS#(k&)
'             RES1# = 0: RES2# = 0
'             '
'             'INCLUYE GASTOS (RESUOPER)
'             Call REPLA(Y$, PER$, 1, 16)
'             Call REPLA(Y$, MKD$(TVEN#), 17, 8)
'             Call REPLA(Y$, MKD$(TCOM#), 25, 8)
'             Call REPLA(Y$, MKD$(TGAS#), 33, 8)
'             RES1# = TVEN# - TCOM# - TGAS#
'             Call REPLA(Y$, MKD$(RES1#), 41, 8)
'             Call REPLA(Y$, MKD$(TCOB#), 49, 8)
'             Call REPLA(Y$, MKD$(TPAG#), 57, 8)
'             'Call REPLA(Y$, MKD$(TGAS#), 65, 8) ' SE QUITA ESTA COLUMNA A PEDIDO DE J Y M
'             RES2# = TCOB# - TPAG# '- TGAS#
'             Call REPLA(Y$, MKD$(RES2#), 73, 8)
'             Call REPLA(Y$, MKI$(DESDEI%(k&)), 81, 2)
'             J& = J& + 1
'             Call GRAREG("!RESUOPER", Y$, J&)
'          End If
'          '
'       Next k&
'       '
'       Call SETULTREG("!RESUOPER", J&)
'       Call TRACENTRAL("RESUOPER", "")
'       Call FILESORT("RESUOPER", "", 10, 10, "DESC")
'       Call CIERRARCH("RESUOPER")
'
'       Call HEADERS("RESUOPER", "")
'       Call HEADERS("RESUOPER", "Corresponde a: " + PERIO$)
'       '
'       Screen.MousePointer = 1
'       '
'       Call FINAL("*RESUOPER")
'       '
'Exit Sub
'===================================
'      Call APERBASDAT("CABAN")
'      '
'      'PARA OBTENER VENTAS NETAS
'      SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
'      SQLX$ = SQLX$ + "WHERE TipoMovim = 'FVEN'"
'      SQLX$ = SQLX$ + "AND  NuClien > 0 "
'      '
''      Dim AVANCE  As New ADODB.Recordset
''      Set AVANCE = FLEXCONN.Execute(FILTSQL$(SQLX$))
'      '
'      With AVANCE
'          '
'          Do While Not .EOF
'               IMPO# = XVALO(!IDEBECOMP) - XVALO(!IHABECOMP)
'               FFI% = CVINT%(!fechconta)
'               '
'               J& = 0
'               For J& = 1 To URE&
'                   If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
'                      VENTA#(J&) = VENTA#(J&) + IMPO#
'                      Exit For
'                   End If
'               Next J&
'          .MoveNext
'          Loop
'          '
'      End With
'      '
'      AVANCE.Close
'      Set AVANCE = Nothing
'      'PARA OBTENER COBRANZAS NETAS
'10     SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
'       SQLX$ = SQLX$ + "WHERE TipoMovim = 'RCOB'"
'       SQLX$ = SQLX$ + "AND  NuClien > 0 "
'       '
''       Dim AVANCE1  As New ADODB.Recordset
'       Set AVANCE1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
'
'       With AVANCE1
'          '
'           Do While Not .EOF
'                IMPO# = XVALO(!IHABECOMP) - XVALO(!IDEBECOMP)
'                FFI% = CVINT%(!fechconta)
'                '
'                J& = 0
'                For J& = 1 To URE&
'                    If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
'                       COBRA#(J&) = COBRA#(J&) + IMPO#
'                       Exit For
'                    End If
'                Next J&
'           .MoveNext
'           Loop
'           '
'       End With
'       '
'       AVANCE1.Close
'       Set AVANCE1 = Nothing
'
'       'PARA OBTENER COMPRAS NETAS
'20     SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
'       SQLX$ = SQLX$ + "WHERE TipoMovim = 'FCOM'"
'       SQLX$ = SQLX$ + "AND  NuProve > 0 "
'       '
''       Dim AVANCE2  As New ADODB.Recordset
'       Set AVANCE2 = FLEXCONN.Execute(FILTSQL$(SQLX$))
'
'       With AVANCE2
'           '
'           Do While Not .EOF
'                IMPO# = XVALO(!IHABECOMP) - XVALO(!IDEBECOMP)
'                FFI% = CVINT%(!fechconta)
'                '
'                J& = 0
'                For J& = 1 To URE&
'                    If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
'                       COMPRA#(J&) = COMPRA#(J&) + IMPO#
'                       Exit For
'                    End If
'                Next J&
'           .MoveNext
'           Loop
'           '
'       End With
'       '
'       AVANCE2.Close
'       Set AVANCE2 = Nothing
'
'       'PARA OBTENER PAGOS NETOS
'30     SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
'       SQLX$ = SQLX$ + "WHERE TipoMovim = 'OPAG'"
'       SQLX$ = SQLX$ + "AND  NuProve > 0 "
'       '
''       Dim AVANCE3  As New ADODB.Recordset
'       Set AVANCE3 = FLEXCONN.Execute(FILTSQL$(SQLX$))
'       '
'       With AVANCE3
'          '
'           Do While Not .EOF
'            IMPO# = XVALO(!IDEBECOMP) - XVALO(!IHABECOMP)
'            FFI% = CVINT%(!fechconta)
'            '
'            J& = 0
'            For J& = 1 To URE&
'                 If FFI% >= DESDEI%(J&) And FFI% <= HASTAI%(J&) Then
'                    If TICUEPRO%((!NuProve) * (-1)) = 0 Then
'                        PAGOS#(J&) = PAGOS#(J&) + IMPO#
'                      Else
'                        GASTOS#(J&) = GASTOS#(J&) + IMPO#
'                    End If
'                    Exit For
'                 End If
'            Next J&
'           .MoveNext
'           '
'           Loop
'           '
'       End With
'       '
'       AVANCE3.Close
'       Set AVANCE3 = Nothing
'
150    Call BLANARCH("!RESUOPER")
       Y$ = REGBLAN$("!RESUOPER")
       J& = 0
       FEX = DESDEX%
       Des% = COMIENFE%(FEX)
       Has% = HASTAX%
       '
       For k& = 1 To URE&
          '
'          If DESDEI%(k&) >= Des% And DESDEI%(k&) <= Has% Then
             PER$ = PERIOD$(k&)
             TVEN# = VENTA#(k&)
             TCOM# = COMPRA#(k&)
             TCOB# = COBRA#(k&)
             TPAG# = PAGOS#(k&)
             TGAS# = GASTOS#(k&)
             RES1# = 0: RES2# = 0
             '
             'INCLUYE GASTOS (RESUOPER)
             Call REPLA(Y$, PER$, 1, 16)
             Call REPLA(Y$, MKD$(TVEN#), 17, 8)
             Call REPLA(Y$, MKD$(TCOM#), 25, 8)
             Call REPLA(Y$, MKD$(TGAS#), 33, 8)
             RES1# = TVEN# - TCOM# - TGAS#
             Call REPLA(Y$, MKD$(RES1#), 41, 8)
             Call REPLA(Y$, MKD$(TCOB#), 49, 8)
             Call REPLA(Y$, MKD$(TPAG#), 57, 8)
             'Call REPLA(Y$, MKD$(TGAS#), 65, 8) ' SE QUITA ESTA COLUMNA A PEDIDO DE J Y M
             RES2# = TCOB# - TPAG# '- TGAS#
             Call REPLA(Y$, MKD$(RES2#), 73, 8)
             Call REPLA(Y$, MKI$(DESDEI%(k&)), 81, 2)
             J& = J& + 1
             Call GRAREG("!RESUOPER", Y$, J&)
'          End If
          '
       Next k&
       '
       Call SETULTREG("!RESUOPER", J&)
       Call TRACENTRAL("RESUOPER", "")
       Call FILESORT("RESUOPER", "", 10, 10, "ASC")
       Call CIERRARCH("RESUOPER")

       Call HEADERS("RESUOPER", "")
       Call HEADERS("RESUOPER", "Corresponde a: " + PERIO$)
       '
       Screen.MousePointer = 1
       '
       Call FINAL("*RESUOPER")
       '
End Sub

Private Sub BOTREC_Click()
    '
    If VALIMESAN% < 1 Then GoTo 99
    
    Call BLANARCH("AVAFACB/NOERR")
    Call BLANARCH("AVAFACC/NOERR")
    Screen.MousePointer = 11
    
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    
    DESDA = FETEXCOR1$(DESDE%)
    HASDA = FETEXCOR1$(HASTA%)
    ELEVEN% = XVALO(TEXT2)
    SQLX$ = "SELECT TipoMovim, NuClien, NUMECOMP, ValAbsComp, OpciMovim, codicom_lar, Mone_Emis, Tipo_Cam, FECHCONTA, ModoOpciMovim, DEUDOR12.RASO_CLI, DEUDOR12.VEND_CLI "
    SQLX$ = SQLX$ + " FROM CAJABANC WITH(NOLOCK) INNER JOIN DEUDOR12 WITH(NOLOCK) ON CAJABANC.NUCLIEN=DEUDOR12.NUME_CLI "
    SQLX$ = SQLX$ + "where CAJABANC.FechConta >= '" & DESDA & "' "
    SQLX$ = SQLX$ + "AND CAJABANC.FechConta <= '" & HASDA & "' "
    SQLX$ = SQLX$ + "AND CAJABANC.CODIEMPR = " & CodiEmp%
    If ELEVEN% <> 0 Then SQLX$ = SQLX$ + " AND DEUDOR12.VEND_CLI=" & CStr(ELEVEN%)
    
    If Option3.Value = True Then
        TipoComprobante$ = ""
        If Check1(0).Value = 1 Then TipoComprobante$ = TipoComprobante$ + " CAJABANC.OPCIMOVIM='FC' OR"
        If Check1(1).Value = 1 Then TipoComprobante$ = TipoComprobante$ + " CAJABANC.OPCIMOVIM='ND' OR"
        If Check1(2).Value = 1 Then TipoComprobante$ = TipoComprobante$ + " CAJABANC.OPCIMOVIM='NC' OR"
        If TipoComprobante$ <> "" Then
            TipoComprobante$ = "(" + TRIM$(Mid$(TipoComprobante$, 1, Len(TipoComprobante$) - 3)) + ")"
            SQLX$ = SQLX$ + " AND " + TipoComprobante$
        End If
    End If
    If Option3.Value = True Then
        OP% = COMALTER%("Opciones de Trabajo:\\Importes Brutos\Importes Netos (Sin Iva)")
        If OP% = 1 Then
            SQLX$ = SQLX$ + " AND CAJABANC.TipoMovim = 'FVEN' "
        ElseIf OP% = 2 Then
            SQLX$ = SQLX$ + " AND CAJABANC.TipoMovim = 'VENTA' "
        Else
            GoTo 99
        End If
      ElseIf Option4.Value = True Then
         SQLX$ = SQLX$ + " AND CAJABANC.TipoMovim = 'RCOB' "
    End If
    Screen.MousePointer = 11
    CantidadRegistros& = CantRegistros("(" & SQLX$ & ") T ", "")
    SQLX$ = SQLX$ + " ORDER BY CAJABANC.FechEmisi ASC, CAJABANC.NumeComp ASC "

    Dim AVANCETMP As ADODB.Recordset
    Set AVANCETMP = READSET(SQLX$)
    With AVANCETMP
        JJ& = 0
        REBLAY$ = REGBLAN$("AVAFACB")
        REBLAX$ = REGBLAN$("AVAFACC")
        Do While Not .EOF
            '
            Call TRACE(24, JJ&, CantidadRegistros&)
            '
            TipoComprobante$ = SAFETEXT$(!OpciMovim)
            Cliente& = XVALO(!NUCLIEN)
            razonSocial$ = SAFETEXT$(!raso_cli)
            IMPORTE# = XVALO(!ValAbsComp)
            If !modoopcimovim = "COMPENSA" Then IMPORTE# = 0 ' ES UN DEBITO POR COMPENSACION CLIENTE-PROVEEDOR
              If TipoComprobante$ = "NC" Then IMPORTE# = (-1) * IMPORTE#
            Comprobante$ = SAFETEXT$(!codicom_lar)
            Moneda% = 0: TipoCambio# = 1
            Moneda% = XVALO(!Mone_Emis)
            If Moneda% = 0 Then Moneda% = 1
            TipoCambio# = XVALO(!Tipo_Cam)
            If TipoCambio# = 0 Then TipoCambio# = 1
            '
            Y$ = REBLAY$
            FechaEmision% = CVINT(!fechconta)
            Call REPLA(Y$, MKI$(FechaEmision%), 1, 2)
            Call REPLA(Y$, razonSocial$, 3, 28)
            Call REPLA(Y$, Comprobante$, 31, 18)
            Call REPLA(Y$, MKD$(IMPORTE#), 49, 8)
            If Moneda% > 1 Then
              If TipoCambio# > 0.1 Then
                Call REPLA(Y$, SIMBOLPES$(Moneda%), 65, 4)
                Call REPLA(Y$, MKD$(IMPORTE# / TipoCambio#), 69, 8)
              End If
            End If
            '
            JJ& = JJ& + 1
            Call GRAREG("AVAFACB", Y$, JJ&)
            '
            XX$ = REBLAX$
            Call REPLA(XX$, MKS(Cliente&), 1, 4)
            Call REPLA(XX$, razonSocial$, 5, 28)
            Call REPLA(XX$, MKI$(FechaEmision%), 33, 2)
            Call REPLA(XX$, Comprobante$, 35, 24)
            Call REPLA(XX$, MKD$(IMPORTE#), 59, 8)
            '
            If Moneda% > 1 Then
              If TipoCambio# > 0.1 Then
                Call REPLA(XX$, SIMBOLPES$(Moneda%), 75, 4)
                Call REPLA(XX$, MKD$(IMPORTE# / TipoCambio#), 79, 8)
              End If
            End If
            '
            Call GRAREG("AVAFACC", XX$, JJ&)
49      .MoveNext
        Loop
      End With
      AVANCETMP.Close
      Set AVANCETMP = Nothing
      '
      Call SETULTREG("AVAFACB", JJ&)
      Call SETULTREG("AVAFACC", JJ&)
      Call CIERRARCH("*.*")
      '
      Screen.MousePointer = 1
      If JJ& = 0 Then
        Call MENSERR(24, "No Se Ha Encontrado Información para el Período Elegido.")
        Exit Sub
      End If
      ' SE AGREGA ORDENAMIENTO PORQUE APARENTEMENTE EL SQL NO LO
      ' ESTA HACIENDO CORRECTAMENTE. SE DETECTA EN MEDITEA
      'Call FILESORT("AVAFACB", "AVAFACB", 3, 3, "ASC") ' POR COMPROBANTE
      'Call FILESORT("AVAFACB", "AVAFACB", 1, 1, "ASC") ' POR FECHA
      '
      VEND1$ = "Todos"
      ECOVEN1$ = TRIM$(Label23)
      If ELEVEN% > 0 And ECOVEN1$ <> "" Then
        VEND1$ = ECOVEN1$
      End If
      '
      Screen.MousePointer = 1
      If Option6 Then
        If Option2 Then
          Call FILESORT("AVAFACC", "AVAFACC", 3, 3, "ASC")
          Call FILESORT("AVAFACC", "AVAFACC", 1, 1, "ASC")
          Call HEADERS("AVAFACC", "")
          Call HEADERS("AVAFACC", "Corresponde a: " & CORRESP$ & " - " & VEND1$)
          Call FINAL("*LIAVAFC")
        ElseIf Option1 Then
          Call FINAL("*COVECLIE")
        End If
        GoTo 50
      End If
      If Option1.Value = True Then
          Call FINAL("*COAVAFB")
        ElseIf Option2.Value = True Then
          Call HEADERS("AVAFACB", "")
          Call HEADERS("AVAFACB", "Corresponde a: " & CORRESP$ & " - " & VEND1$)
          Call FINAL("*LIAVAFB")
      End If
      '
50    Hide
'HASTA ACA NUEVO
   '//////*****RUTINA ANTIGUA*****///////
    '
'    MMXX0$ = TRIM$(MESAN.TEXT)
'    MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
'    If TRIM$(MMXX1$) = "" Then
'        Call MENSERR(24, "Ingrese el Mes a Listar")
'        Exit Sub
'    End If
'    '
'    Call BLANARCH("AVAFACB/NOERR")
'    REAVA& = 0
'    Screen.MousePointer = 11
'    '
'    LTX$ = LETREXPO$
'    If Option3.Value = True Then
'      FIN& = ULTREG&("FACTUR")
'      For IL& = 1 To FIN&
'        Call TRACE(20, IL&, FIN&)
'        X$ = REGLEIDO$("FACTUR", IL&)
'        FF% = CVI(Mid$(X$, 7, 2))
'        If FF% >= DES% Then
'            If FF% <= HAS% Then
'                NCLI = CVI(Left$(X$, 2))
'                CLIE$ = ""
'                If NCLI > 0 Then
'                      CLIE$ = RASOCLI$(NCLI)
'                   Else
'                      REGCU& = Abs(NCLI)
'                      If REGCU& > 0 Then
'                         If REGCU& <= ULTREG&("CUITS") Then
'                            Z$ = REGLEIDO$("CUITS", REGCU&)
'                            CLIE$ = Left$(Z$, 30)
'                         End If
'                      End If
'                End If
'                NRO& = CVS(Mid$(X$, 3, 4))
'                '
'                TI% = Asc(Mid$(X$, 13, 1)): If TI% < 0 Or TI% > 7 Then TI% = 2 ' default factura
'                If TI% = 2 And Check1(0).Value <> 1 Then GoTo 19
'                If TI% = 3 And Check1(1).Value <> 1 Then GoTo 19
'                If TI% = 6 And Check1(2).Value <> 1 Then GoTo 19
'                '
'                VA% = Asc(Mid$(X$, 14, 1)): If VA% < 1 Or VA% > 3 Then VA% = 1 ' default factura
'                SIG% = 1: If TI% > 4 Then SIG% = (-1)
'                NETO# = Abs((Int(100 * (CVD(Mid$(X$, 15, 8)) + CVD(Mid$(X$, 49, 8))) + 0.5)) / 100)
'                NETO# = NETO# * SIG%
'                NUSUCUVE% = CVI(Mid$(X$, 47, 2))
'                If TI% > 0 And TI% < 7 And TI% <> 4 Then
'                    NNX$ = FORMDOC$(NRO&, TI%, VA%, NUSUCUVE%)
'                    ' por VDR que ponga letra 'E' a exportaciones
'                    PVCLX1% = PIVACLI%(NCLI)
'                    If PVCLX1% = 5 Or PVCLX1% = 7 Then
'                      Call REPLA(NNX$, PREFEXPO$, 4, 4)
'                      Call REPLA(NNX$, LETREXPO$, 18, 1)
'                    End If
'                    '
'                    Y$ = REGBLAN$("AVAFACB")
'                    Call REPLA(Y$, MKI$(FF%), 1, 2)
'                    Call REPLA(Y$, CLIE$, 3, 28)
'                    Call REPLA(Y$, NNX$, 31, 18)
'                    Call REPLA(Y$, MKD$(NETO#), 49, 8)
'                    REAVA& = REAVA& + 1
'                    Call GRAREG("AVAFACB", Y$, REAVA&)
'                End If
'            End If
'        End If
'19    Next IL&
'      '
'    ElseIf Option4.Value = True Then ' Avance de Cobranza
'      '
'      FIN& = ULTREG&("COBRAN")
'      For IL& = 1 To FIN&
'        Call TRACE(20, IL&, FIN&)
'        X$ = REGLEIDO$("COBRAN", IL&)
'        FF% = CVI(Mid$(X$, 7, 2))
'        If FF% >= DES% Then
'            If FF% <= HAS% Then
'                NCLI = CVI(Left$(X$, 2))
'                CLIE$ = RASOCLI$(NCLI)
'                '
'                NUSUCUVE% = CVI(Mid$(X$, 23, 2))
'                NRO& = CVS(Mid$(X$, 3, 4))
'                NETO# = Abs((Int(100 * (CVD(Mid$(X$, 15, 8))) + 0.5)) / 100)
'                PFX$ = TRIM$(Str$(NUSUCUVE%)): While Len(PFX$) < 2: PFX$ = "0" + PFX$: Wend
'                NNX$ = TRIM$(Str$(NRO&)): While Len(NNX$) < 6: NNX$ = "0" + NNX$: Wend
'                NNX$ = "RB-" + PFX$ + "-" + NNX$
'                '
'                Y$ = REGBLAN$("AVAFACB")
'                Call REPLA(Y$, MKI$(FF%), 1, 2)
'                Call REPLA(Y$, CLIE$, 3, 28)
'                Call REPLA(Y$, NNX$, 31, 18)
'                Call REPLA(Y$, MKD$(NETO#), 49, 8)
'                REAVA& = REAVA& + 1
'                Call GRAREG("AVAFACB", Y$, REAVA&)
'            End If
'        End If
'      Next IL&
'      '
'    End If
'    '
'    Call SETULTREG("AVAFACB", REAVA&)
'    Call CIERRARCH("*.*")
'    Screen.MousePointer = 1
'    '
'    Call FILESORT("AVAFACB", "", 1, 1)
'    '
'    If Option1.Value = True Then
'        Call FINAL("*COAVAFB")
'      ElseIf Option2.Value = True Then
'        Call FINAL("*LIAVAFB")
'    End If
'    '
'    Hide
'    '
99 End Sub

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command2_Click()
  '
  Call SELECHO("VENDEDOR")
  VDEVU$ = VALACT1$("VENDEDOR")
  If TRIM$(VDEVU$) <> "" Then
    TEXT2.TEXT = VDEVU$
  End If
  '
End Sub

Private Sub Command3_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
        Exit Sub
    End If
    MESAN.TEXT = MESASEL$
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call REMEVE07.Command15_Click
   Command4.Enabled = True

End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
   opt% = ALTERNA%("Resumen Operativo\Otros Accesos Directos")
   '
   If opt% = 0 Then GoTo 99
   If opt% = 1 Then
      If InStr(EMPREAC$, "J Y M") > 0 Then
         OPX% = COMALTER%("Formato de Salida:\\Excluir Gastos\Incluir Gastos")
         If OPX% = 1 Then
            'OPX1% = COMALTER%("Selección\\Excluir Gastos Tradicional\Excluir Gastos Nuevo Modelo")
               'If OPX1% = 1 Then
                  Call RESUOPER
               'ElseIf OPX1% = 2 Then
               '   Call RESUOPER2
               'End If
            GoTo 99
         ElseIf OPX% = 2 Then
            OPX1% = COMALTER%("Selección\\Incluir Gastos Tradicional\Incluir Gastos Nuevo Modelo")
              If OPX1% = 1 Then
                Call RESOPEGAS ' RESUMEN OPERATIVO INCLUTE GASTO
               ElseIf OPX1% = 2 Then
                Call RESOPEGAS2
               End If
            GoTo 99
         End If
      End If
      '
   Call RESUOPER
   '
   End If
   '
99 Command8.Enabled = True
End Sub

Private Sub Form_Load()
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me, Picture1)
End Sub

Private Sub MESAN_DblClick()
   '
   Call DESHASFECHA(Des%, Has%, PERIO$)
   MESAN = PERIO$
   '
End Sub

Private Sub mnuImpresiónComprobantes_Click()
    '
    '\\\OT-05-0668-WAR-12/10/05///
    Call DESHASFECHA(Des%, Has%, PERIO$)
    If PERIO$ = "" Then GoTo 99
    '
    Screen.MousePointer = 11
    PRINFACT.GRID.Redraw = False
    DESDA = FETEXCOR1$(Des%)
    HASDA = FETEXCOR1$(Has%)
    SQLX$ = "SELECT C.OpciMovim, C.TipoMovim, C.NuClien, C.NUMECOMP, C.ValAbsComp, "
    SQLX$ = SQLX$ + " C.OpciMovim, C.CodiSucu, C.codicom_lar, C.Mone_Emis, C.Tipo_Cam, C.FECHCONTA, D.Vend_Cli, D.Raso_Cli "
    SQLX$ = SQLX$ + " FROM CAJABANC C WITH(NOLOCK) inner join DEUDOR12 D ON D.Nume_Cli = C.NuClien"
    SQLX$ = SQLX$ + " where C.FechConta >= '" & DESDA & "' "
    SQLX$ = SQLX$ + " AND C.FechConta <= '" & HASDA & "' "

    OP% = COMALTER%("Opciones de Trabajo:\\Importes Brutos\Importes Netos (Sin Iva)")
    If OP% = 1 Then
      SQLX$ = SQLX$ + "AND TipoMovim = 'FVEN' "
    ElseIf OP% = 2 Then
      SQLX$ = SQLX$ + "AND TipoMovim = 'VENTA' "
    Else
      Screen.MousePointer = 1
      GoTo 99
    End If
    Screen.MousePointer = 1
    SQLX$ = SQLX$ + "Order by FechEmisi ASC, NumeComp ASC "
    
    '
'    PRINFACT
    Campos$ = "Fecha/D8;Cta./F8;Cliente/A32;M.Orig./A7;Documento/A24;Importe $/F14.2;NUMECOMP/F10;PARAORDENAR/A0"
    Call CARGA_ENCABEZADO(PRINFACT.GRID, Campos$, PRINFACT, 1)
    Call CARGA_ENCABEZADO(PRINFACT.MSF_2, Campos$, PRINFACT, 1)

    Set rs = READSET(SQLX$)
      With rs
        If .EOF Then
          Call MENSERR(24, "Ningun Registro Coincide con el Filtro SQL")
          Exit Sub
        End If
        JJ& = 0
        Do While Not .EOF
          OPCIXM$ = SAFETEXT$(!OpciMovim)
          If OPCIXM$ = "FC" And Check1(0).Value <> 1 Then GoTo 49
          If OPCIXM$ = "ND" And Check1(1).Value <> 1 Then GoTo 49
          If OPCIXM$ = "NC" And Check1(2).Value <> 1 Then GoTo 49
          If SAFETEXT$(!TIPOMOVIM) = "RCOB" And Option4.Value = False Then GoTo 49
          '
          NCLI = XVALO(!NUCLIEN)
          CLIEX$ = ""
          If NCLI > 0 Then
             '\\\OT-06-0302-WAR-19/07/06///
             NUMVEN% = XVALO(!VEND_CLI)
             ELEVEN% = XVALO(TEXT2) 'VENDEDOR ELEGIDO
             If ELEVEN% = 0 Or ELEVEN% = NUMVEN% Then
             '\\\OT-06-0302-WAR-FIN///
                CLIEX$ = SAFETEXT$(!raso_cli)
                NRO& = XVALO(!NumeComp)
                NETO# = XVALO(!ValAbsComp) 'Abs((Int(100 * (CVD(Mid$(X$, 15, 8)) + CVD(Mid$(X$, 49, 8))) + 0.5)) / 100)
                  If OPCIXM$ = "NC" Then NETO# = (-1) * NETO#
                NUSUCUVE% = XVALO(!CodiSucu)
                NNX$ = SAFETEXT$(!codicom_lar)
                MONEM% = 0: TICAM = 1
                On Error Resume Next
                MONEM% = XVALO(!Mone_Emis)
                On Error GoTo 0
                '
                TICAM = 1
                On Error Resume Next
                If IsNull(!Tipo_Cam) = False Then
                  TICAM = !Tipo_Cam
                End If
                On Error GoTo 0
                '
                SIM$ = "$"
'                If MONEM% > 1 Then
                  If TICAM > 0.1 Then
                    SIM$ = SIMBOLPES$(MONEM%)
'                    NETO# = NETO# / TICAM
                  End If
'                End If

                J& = J& + 1
                PRINFACT.GRID.Rows = J& + 1
                PRINFACT.GRID.TextMatrix(J&, 1) = FECHATEX(CVINT(!fechconta))
                PRINFACT.GRID.TextMatrix(J&, 2) = NCLI
                PRINFACT.GRID.TextMatrix(J&, 3) = CLIEX$
                PRINFACT.GRID.TextMatrix(J&, 4) = SIM$
                PRINFACT.GRID.TextMatrix(J&, 5) = NNX$
                PRINFACT.GRID.TextMatrix(J&, 6) = FORMATNUM$(NETO#, "F14.2")
                PRINFACT.GRID.TextMatrix(J&, 7) = XVALO(!NumeComp)
                PRINFACT.GRID.TextMatrix(J&, 8) = Left$(NNX$, 7) & Right$(NNX$, 1) & FECHALETRA$(XVALO(!NumeComp))
                '
             End If
          End If
49      .MoveNext
        Loop
      End With
      rs.Close
      Set rs = Nothing
      PRINFACT.GRID.Redraw = True
      PRINFACT.Show 1
      '
      Screen.MousePointer = 1

99
End Sub

Private Sub Option3_Click()
   For KI% = 0 To 2
     Check1(KI%).Enabled = True
     Check1(KI%).Value = 1
   Next KI%
End Sub

Private Sub Option4_Click()
   For KI% = 0 To 2
     Check1(KI%).Value = 0
     Check1(KI%).Enabled = False
   Next KI%
End Sub
'\\\OT-05-0668-WAR-12/10/05///
Function VALIMESAN%()
   VMA% = 0
   '
   If Len(MESAN) = 19 Then
     DESDE% = FECHANUM(Left$(MESAN, 8))
     If DESDE% > 0 Then
       HASTA% = FECHANUM(Mid$(MESAN, 12, 8))
       If HASTA% >= DESDE% Then
         VMA% = 1
         'DESDAT = Int(CDbl(CVDAT(DESDE%)))
         'HASDAT = Int(CDbl(CVDAT(HASTA%)))
         DESDAT = DESDE%
         HASDAT = HASTA%
         CORRESP$ = MESAN
         GoTo 98
       End If
     End If
   End If
   '
   CORRESP$ = MESANO$(MESAN, DESDE%, HASTA%)
   If CORRESP$ <> "" Then
       VMA% = 1
       'DESDAT = Int(CDbl(CVDAT(DESDE%)))
       'HASDAT = Int(CDbl(CVDAT(HASTA%)))
       DESDAT = DESDE%
       HASDAT = HASTA%
     Else
       Call MENSERR(24, "Falta Elegir Período Mensual")
   End If
   '
98 VALIMESAN% = VMA%
   '
End Function

'\\\OT-05-0668-WAR-FIN///
Private Sub Text2_Change()
  '
  Label23 = TRIM$(ECOARCH$("VENDEDOR", MKI$(XVALO(TEXT2.TEXT))))
  '
End Sub
