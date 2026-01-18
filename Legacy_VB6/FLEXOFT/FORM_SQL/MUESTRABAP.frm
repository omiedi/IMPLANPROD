VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form MUESTRABAP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Reportes"
   ClientHeight    =   7470
   ClientLeft      =   120
   ClientTop       =   675
   ClientWidth     =   15735
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   15735
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRIDH 
      Height          =   6165
      Left            =   120
      TabIndex        =   20
      Top             =   720
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   10874
      _Version        =   393216
      Rows            =   3
      RowHeightMin    =   1
      BackColorSel    =   12632256
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      TextStyleFixed  =   1
      FillStyle       =   1
      SelectionMode   =   1
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   2
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ejecutar Consulta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   280
      Width           =   7815
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   120
      TabIndex        =   12
      Top             =   6840
      Width           =   8350
      Begin VB.CommandButton Command4 
         Caption         =   "Filtrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   7440
         Picture         =   "MUESTRABAP.frx":0000
         TabIndex        =   22
         ToolTipText     =   "Emite Listado"
         Top             =   130
         Width           =   700
      End
      Begin VB.TextBox TXTBUSCAR 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   750
         TabIndex        =   14
         Top             =   160
         Width           =   2535
      End
      Begin VB.ComboBox CMBFILTRO 
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
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   160
         Width           =   3135
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   160
         Width           =   585
      End
      Begin VB.Label Label71 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Campo:"
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
         Left            =   3360
         TabIndex        =   15
         Top             =   160
         Width           =   705
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exportar a Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   12960
      TabIndex        =   11
      Top             =   6960
      Width           =   2685
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Maximizar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   10560
      Picture         =   "MUESTRABAP.frx":030A
      TabIndex        =   10
      ToolTipText     =   "Emite Listado"
      Top             =   6960
      Visible         =   0   'False
      Width           =   2025
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1440
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2640
      OleObjectBlob   =   "MUESTRABAP.frx":0614
      Top             =   -120
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "MUESTRABAP.frx":0848
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   6165
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   10874
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF_2 
      Height          =   6165
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   10874
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   8421631
      BackColorBkg    =   12632256
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   7575
      Begin MSComCtl2.DTPicker dthorades 
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   116981762
         CurrentDate     =   42549
      End
      Begin MSComCtl2.DTPicker dthorahas 
         Height          =   375
         Left            =   5880
         TabIndex        =   3
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   116981762
         CurrentDate     =   42549
      End
      Begin MSComCtl2.DTPicker DTPFechades 
         Height          =   375
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Width           =   1400
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   116981761
         CurrentDate     =   42549
      End
      Begin MSComCtl2.DTPicker DTPFechahas 
         Height          =   375
         Left            =   4440
         TabIndex        =   2
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   116981761
         CurrentDate     =   42549
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   19
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   735
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRIDH2 
      Height          =   6165
      Left            =   120
      TabIndex        =   21
      Top             =   720
      Visible         =   0   'False
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   10874
      _Version        =   393216
      BackColorFixed  =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7800
      TabIndex        =   23
      Top             =   30
      Width           =   6975
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuReportes 
      Caption         =   "Reportes"
      Begin VB.Menu mnuReporte1 
         Caption         =   "Reporte General Reservas x Fechas"
      End
      Begin VB.Menu MnuParticularesInstitucionesSinReservas 
         Caption         =   "Particulares e Instituciones Sin Reservas  (Filtrado x Fecha de Reserva)"
      End
      Begin VB.Menu mnuReporteHistorico 
         Caption         =   "Reporte Histórico"
      End
   End
End
Attribute VB_Name = "MUESTRABAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FrmLeft, FrmTop, FrmWidth, FrmHeight As Long
Dim GridTop, GridLeft, GridWidth, GridHeight As Long
Dim FRAMETop, FRAMELeft, FRAMEWidth, FRAMEHeight As Long
Dim Com1Top, Com2Top, Com2Left As Long
Dim Com1LEFT As Long

Dim LecturaDatosOriginales As Byte


Private Sub BOTEXIT_Click()
  Unload Me
End Sub

Sub BUSCARHERICAL()
    If TRIM$(TXTBUSCAR) = "" Then
      GRIDH2.Visible = False
      GRIDH.Visible = True
      Exit Sub
    End If
    '
    GRIDH2.Rows = 2
    For i& = 1 To GRIDH2.Cols - 1
       GRIDH2.TextMatrix(1, i&) = ""
    Next i&
    
    REG& = 0
    GRIDH2.Visible = True
    GRIDH2.ZOrder 0
    GRIDH2.Cols = GRIDH2.Cols
    GRIDH2.Redraw = False
    GRIDH.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For i& = 1 To GRIDH.Rows - 1
      For J& = 0 To GRIDH.Cols - 1
        If (InStr(UCase$(GRIDH.TextMatrix(i&, J&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(GRIDH.TextMatrix(0, J&))) Then
            REG& = REG& + 1
            GRIDH2.Rows = REG& + 1
            For H& = 0 To GRIDH2.Cols - 1
              GRIDH2.TextMatrix(REG&, H&) = GRIDH.TextMatrix(i&, H&)
            Next H&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next J&
    Next i&
    '
    GRIDH2.Redraw = True
    
    '


End Sub

Sub filtrar()
    If GRIDH.Visible = True Or GRIDH2.Visible = True Then
       Call BUSCARHERICAL
       Exit Sub
    End If
    
    If TRIM$(TXTBUSCAR) = "" Then
      MSF_2.Visible = False
      GRID.Visible = True
      Exit Sub
    End If
    '
    
    MSF_2.Rows = 2
    For i& = 1 To MSF_2.Cols - 1
       MSF_2.TextMatrix(1, i&) = ""
    Next i&
    REG& = 0
    MSF_2.Visible = True
    MSF_2.ZOrder 0
    MSF_2.Cols = GRID.Cols
    MSF_2.Redraw = False
    GRID.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For i& = 1 To GRID.Rows - 1
      For J& = 1 To GRID.Cols - 1
        If (InStr(UCase$(GRID.TextMatrix(i&, J&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(GRID.TextMatrix(0, J&))) Then
            REG& = REG& + 1
            MSF_2.Rows = REG& + 1
            For H& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(REG&, H&) = GRID.TextMatrix(i&, H&)
            Next H&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next J&
    Next i&
    '
    MSF_2.Redraw = True
    
End Sub

Sub RESERVASSEGUNAÑO(DES, HAS)
   Dim ESTAXO$(5)
   ESTAXO$(0) = "CANCELADA"
   ESTAXO$(1) = "EN ESPERA"
   ESTAXO$(2) = "CONFIRMADA"
   ESTAXO$(3) = "RECONFIRMADA"
   '
   ANO$ = Year(DES)
   SQLX$ = "select b.FechaReserva as freserva, F.fecha as ffuncion , f.obra,I.SCHOOLS AS NAMEINSTITUCION"
   SQLX$ = SQLX$ + ", f.idfuncion, b.IDFuncion,I.Telephones AS TELEFUN,i.Email AS MAILINST,P.name,P.Telephones AS TELEFPART"
   SQLX$ = SQLX$ + ", P.Email AS MAILPART,B.LASTUPDATE AS BLASTUPDATE"
   SQLX$ = SQLX$ + ", B.Cantidad ,  F.Precio_A, F.Precio_B, F.Precio_C, B.Presentes"
   SQLX$ = SQLX$ + ", B.ESTADO, B.IDParticular, B.IDInstitucion,f.Fecha,B.PRECIO AS BPRECIO"
   SQLX$ = SQLX$ + ", P.SCHOOLS AS NAMEPARTICULAR, F.LUGAR AS FLUGAR, I.address as IDomi"
   SQLX$ = SQLX$ + ", I.town as ICiudad, I.province as IProv, I.[Zip Code] as ICP "
   SQLX$ = SQLX$ + ", P.address as PDomi, P.town as PCiudad, P.province as PProv, P.[Zip Code] as PCP "
   SQLX$ = SQLX$ + " from bookings b WITH(NOLOCK) inner join funciones f on b.idfuncion = f.idfuncion "
   SQLX$ = SQLX$ + " left join instituciones I on  b.IDInstitucion =  I.id"
   SQLX$ = SQLX$ + "  inner join Particulares P on b.idparticular = P.id"
   SQLX$ = SQLX$ + " WHERE b.idreserva > 0 "
   SQLX$ = SQLX$ + " AND f.Fecha  >= '" & DES & "'"
   SQLX$ = SQLX$ + " AND f.Fecha  <= '" & HAS & "'"

   J& = MUESTRABAP.GRIDH.Rows
   MUESTRABAP.GRIDH.ColWidth(0) = 1000
   MUESTRABAP.GRIDH2.ColWidth(0) = 1000
   MUESTRABAP.GRIDH.TextMatrix(0, 0) = "Año"
   MUESTRABAP.GRIDH2.TextMatrix(0, 0) = "Año"
   
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF

       MUESTRABAP.GRIDH.Rows = J& + 1
       MUESTRABAP.GRIDH.TextMatrix(J&, 0) = ANO$
       MUESTRABAP.GRIDH.TextMatrix(J&, 1) = XVALO(!IDINSTITUCION) 'SAFETEXT$(!ffuncion)
       MUESTRABAP.GRIDH.TextMatrix(J&, 2) = SAFETEXT$(!NAMEINSTITUCION) 'SAFETEXT$(!freserva)
       MUESTRABAP.GRIDH.TextMatrix(J&, 3) = SAFETEXT$(!TELEFUN) 'SAFETEXT$(!OBRA)
       MUESTRABAP.GRIDH.TextMatrix(J&, 4) = SAFETEXT$(!MAILINST) 'SAFETEXT$(!NAMEINSTITUCION)
       MUESTRABAP.GRIDH.TextMatrix(J&, 5) = SAFETEXT$(!IDomi) 'SAFETEXT$(!TELEFUN)
       MUESTRABAP.GRIDH.TextMatrix(J&, 6) = SAFETEXT$(!ICiudad) 'SAFETEXT$(!MAILINST)
       MUESTRABAP.GRIDH.TextMatrix(J&, 7) = SAFETEXT$(!IProv) 'SAFETEXT$(!IDomi)
       MUESTRABAP.GRIDH.TextMatrix(J&, 8) = SAFETEXT$(!ICP) 'SAFETEXT$(!ICiudad)
       
       MUESTRABAP.GRIDH.TextMatrix(J&, 9) = XVALO(!IDPARTICULAR)
       MUESTRABAP.GRIDH.TextMatrix(J&, 10) = SAFETEXT$(!NAMEPARTICULAR) 'SAFETEXT$(!IProv)
       MUESTRABAP.GRIDH.TextMatrix(J&, 11) = SAFETEXT$(!TELEFPART) 'SAFETEXT$(!ICP)
       MUESTRABAP.GRIDH.TextMatrix(J&, 12) = SAFETEXT$(!MAILPART) 'SAFETEXT$(!NAMEPARTICULAR)
       MUESTRABAP.GRIDH.TextMatrix(J&, 13) = SAFETEXT$(!PDomi) 'SAFETEXT$(!TELEFPART)
       MUESTRABAP.GRIDH.TextMatrix(J&, 14) = SAFETEXT$(!PCiudad) 'SAFETEXT$(!MAILPART)
       MUESTRABAP.GRIDH.TextMatrix(J&, 15) = SAFETEXT$(!PProv) 'SAFETEXT$(!PDomi)
       MUESTRABAP.GRIDH.TextMatrix(J&, 16) = SAFETEXT$(!PCP) 'SAFETEXT$(!PCiudad)
       
       EST% = XVALO(!ESTADO)
       MUESTRABAP.GRIDH.TextMatrix(J&, 17) = SAFETEXT$(!BLASTUPDATE) 'SAFETEXT$(!PProv)
       MUESTRABAP.GRIDH.TextMatrix(J&, 18) = ESTAXO$(EST% + 1) 'SAFETEXT$(!PCP)
       MUESTRABAP.GRIDH.TextMatrix(J&, 19) = SAFETEXT$(!freserva) 'XVALO(!Cantidad)
       
       MUESTRABAP.GRIDH.TextMatrix(J&, 20) = SAFETEXT$(!ffuncion) 'XVALO(!Precio_A)
       MUESTRABAP.GRIDH.TextMatrix(J&, 21) = SAFETEXT$(!OBRA) 'XVALO(!Precio_B)
       MUESTRABAP.GRIDH.TextMatrix(J&, 22) = SAFETEXT$(!FLUGAR)
       
       MUESTRABAP.GRIDH.TextMatrix(J&, 23) = XVALO(!Cantidad) 'XVALO(!Precio_C)
       MUESTRABAP.GRIDH.TextMatrix(J&, 24) = XVALO(!BPRECIO)
        MUESTRABAP.GRIDH.TextMatrix(J&, 25) = XVALO(!Precio_A) 'XVALO(!PRESENTES)
'        EST% = XVALO(!ESTADO)
       MUESTRABAP.GRIDH.TextMatrix(J&, 26) = XVALO(!Precio_B) 'ESTAXO$(EST% + 1) 'POR QUE EN LA TABLA GUARDA -1, 0,1,2
       MUESTRABAP.GRIDH.TextMatrix(J&, 27) = XVALO(!Precio_C) 'CVINT(!ffuncion)
       MUESTRABAP.GRIDH.TextMatrix(J&, 28) = XVALO(!PRESENTES) 'CVINT(!freserva)
        'SAFETEXT$(!BLASTUPDATE)
       MUESTRABAP.GRIDH.TextMatrix(J&, 29) = FechaHoraNum(SAFETEXT$(!ffuncion))
       MUESTRABAP.GRIDH.TextMatrix(J&, 30) = FechaHoraNum(SAFETEXT$(!freserva))
       J& = J& + 1
       Label3.Caption = "Registros Procesados: " & J&
       Label3.Refresh
       .MoveNext
     Loop
   End With
   RS.Close
   Set RS = Nothing
99 Screen.MousePointer = 1

End Sub

Private Sub CMBFILTRO_Change()
   TXTBUSCAR_Change
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   '
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   AREXPO$ = UCase$(CommonDialog1.FileName)
   If TRIM$(AREXPO$) <> "" Then
     If MSF_2.Visible = True Then
      Call GENEXCEL(AREXPO$, MUESTRABAP.MSF_2)
     ElseIf Me.GRID.Visible = True Then
      Call GENEXCEL(AREXPO$, MUESTRABAP.GRID)
     ElseIf Me.GRIDH.Visible = True Then
      Call GENEXCEL_MSFH(AREXPO$, MUESTRABAP.GRIDH)
     ElseIf Me.GRIDH2.Visible = True Then
      Call GENEXCEL_MSFH(AREXPO$, MUESTRABAP.GRIDH2)
     End If
   End If
   '

99 Command1.Enabled = True

End Sub
Private Sub Command2_Click()
    '
    If Command2.Caption = "Maximizar" Then
        If LecturaDatosOriginales = 0 Then
            ' GUARDA DIMENSIONES Y POSICION DEL FORMULARIO, GRID Y BOTONES
            FrmLeft = Me.Left: FrmTop = Me.Top: FrmWidth = Me.Width: FrmHeight = Me.Height
            GridTop = GRID.Top: GridLeft = GRID.Left: GridWidth = GRID.Width: GridHeight = GRID.Height
            Com1Top = Command1.Top: Com1Top = Command1.Top: Com1LEFT = Command1.Left
            Com2Top = Command2.Top: Com2Top = Command2.Top: Com2Left = Command2.Left

            FRAMETop = Frame1.Top: FRAMELeft = Frame1.Left: FRAMEWidth = Frame1.Width: FRAMEHeight = Frame1.Height
            LecturaDatosOriginales = 1
        End If
    
        QCOL% = GRID.Cols
        ANCHOCOL& = 0
        For COLUM% = 1 To QCOL% - 1
            ANCHOCOL& = ANCHOCOL& + GRID.ColWidth(COLUM%)
        Next COLUM%
        ANCHOSCR& = Screen.Width
        If ANCHOCOL& > ANCHOSCR& Then
            Me.Left = 0: Me.Width = ANCHOSCR&
            GRID.Width = ANCHOSCR& - 150
        Else
            LEFTFORM& = Int((ANCHOSCR& - ANCHOCOL&) / 2)
            Me.Left = LEFTFORM&: Me.Width = ANCHOCOL&
            GRID.Width = ANCHOCOL& - 150
        End If
        '
        QFILAS% = GRID.Rows
        ALTOFILA& = 0
        For FILA% = 1 To QFILAS% - 1
            ALTOFILA& = ALTOFILA& + GRID.RowHeight(FILA%)
        Next FILA%
        ALTOSCR& = Screen.Height
        If ALTOFILA& + 1600 > ALTOSCR& Then
            Me.Top = 0: Me.Height = ALTOSCR& - 600
            GRID.Height = ALTOSCR& - 1600
        Else
            TOPFORM& = Int((ALTOSCR& - (ALTOFILA& + 1600)) / 2)
            Me.Top = TOPFORM&: Me.Height = ALTOFILA& + 1600
            GRID.Height = ALTOFILA& + 600
        End If
        '
        BORDEINF& = GRID.Height + 100
        Command2.Top = BORDEINF&
        Command1.Top = BORDEINF&
        Frame1.Top = BORDEINF&
        Command1.Left = Me.Width - 1800
        Command2.Caption = "Minimizar"
        
    Else
        Me.Left = FrmLeft: Me.Top = FrmTop: Me.Width = FrmWidth: Me.Height = FrmHeight
        GRID.Top = GridTop: GRID.Left = GridLeft: GRID.Width = GridWidth: GRID.Height = GridHeight
        Frame1.Top = FRAMETop: Frame1.Left = FRAMELeft: Frame1.Width = FRAMEWidth: Frame1.Height = FRAMEHeight
        Command2.Top = Com2Top: Command2.Top = Com2Top: Command2.Left = Com2Left
        Command1.Top = Com1Top: Command1.Top = Com1Top: Command1.Left = Com1LEFT
        Command2.Caption = "Maximizar"
    End If
    MSF_2.Height = GRID.Height
    MSF_2.Width = GRID.Width
    MSF_2.Top = GRID.Top
    MSF_2.Left = GRID.Left
    
    '
End Sub

Sub GENEXCEL(NombreDeArchivo$, GRID As MSFlexGrid)
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim valor As Double
   Dim TEXTO As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   TIPODATO% = 0: COLUM& = 0
   For i& = 1 To GRID.Rows
       RENG1& = i&: COLUM& = 0
       For COLU1& = 1 To GRID.Cols
          If GRID.ColWidth(COLU1& - 1) = 0 Then GoTo 10 ' OBVIA LAS COLUMNAS OCULTAS
          COLUM& = COLUM& + 1
          '
          VALORX = GRID.TextMatrix(i& - 1, COLU1 - 1)
          
'          If Len(VALORX) = 8 And Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
'             FEXU$ = VALORX
'             VALORX = FETEXCOR1$(Int(FECHANUM(FEXU$)))
'             TIPODATO% = 1
'          End If
          'ESTO MODIFIQUE PARA QUE MUESTRE TODO COMPLETO LA FECHA Y LA  HORA
          If IsDate(VALORX) Then
              If Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
                 If Len(VALORX) = 8 Then
                   VALORX = FETEXCOR1$(CVINT(VALORX))
                   TIPODATO% = 1
                 Else
                   TIPODATO% = 0
                   VALORX = "A " & VALORX
                 End If
              End If
          End If

          ' ESTO ES PARA QUE EXPORTE COMO CERO Y NO ESPACIOS
          FFCC$ = UCase$(Left$(FORCAMPO$(COLU1&), 1))
          If FFCC$ = "I" Or FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
             If TRIM$(CStr(VALORX)) = "" Then
                VALORX = 0
             End If
             TIPODATO% = 2
          End If
          If FFCC$ = "C" Or FFCC$ = "A" Then VALORX = "'" & VALORX
          '
          'PARA QUE NO REPITA ENCABEZADO
'          If RENG1& = 1 Then
'            If VALORX = VALORX_ANT And TRIM$(VALORX) <> "" Then
'              If COLUM& <= 17 Then
'               LETRAX1$ = Chr$(64 + COLUM& - 1)
'               LETRAX2$ = Chr$(64 + COLUM&)
'               RANXO$ = LETRAX1$ & TRIM$(Str$(RENG1&)) & ":" & LETRAX2$ & TRIM$(Str$(RENG1&))
'               Range(RANXO$).Select
'
'               XLSHEET.Cells.HorizontalAlignment = xlCenter
'               XLSHEET.Cells.HorizontalAlignment = xlCenter
'               XLSHEET.Cells.VerticalAlignment = xlCenter
'               XLSHEET.Cells.WrapText = False
'               XLSHEET.Cells.Orientation = 0
'               XLSHEET.Cells.AddIndent = False
'               XLSHEET.Cells.IndentLevel = 0
'               XLSHEET.Cells.ShrinkToFit = False
'               XLSHEET.Cells.ReadingOrder = xlContext
'               XLSHEET.Cells.MergeCells = True
'               GoTo 10
'             End If
'            End If
'          End If
          '
          XLSHEET.Cells(RENG1&, COLUM&) = VALORX
          If TIPODATO% = 1 Then
             XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yy"
'             If IsDate(DatoCelda) Then
'                    .NumberFormat = "dd/mm/yy"
          End If
          'If TIPODATO% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#,##0.00"
          If TIPODATO% = 0 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
          'If TIPODATO% = 3 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "[$-2C0A] dd/mm/aa hh:mm:ss AM/PM;@"
          
          TIPODATO% = 0
          VALORX_ANT = VALORX
10     Next COLU1&
       Command1.Caption = "Registros Exportados: " & i&
   Next i&
   '
   Close #N1%
   On Error GoTo 30
   XLSHEET.SaveAs NombreDeArchivo$
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   Screen.MousePointer = 1
   MsgBox "Planilla Excel Creada Exitosamente: " & NombreDeArchivo$
   'Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NombreDeArchivo$, "\", "/"))
   '
Exit Sub
   '
   '
20 Resume 21
21 On Error GoTo 0
   MsgBox "Utilitario EXCEL no Instalado"
   On Error GoTo 0
   Exit Sub
   '
30 Resume 31
31 On Error GoTo 0
   Screen.MousePointer = 1
   '
   ' Cierra el objeto Book, el cual es utilizado al crear
   ' una nueva hoja en el archivo Excel.
   ' El valor False es para que no pregunte si quiere
   ' guardar el archivo.
   XLBOOK.Close (False)
   XLAPP.Quit ' Cierra Excel
   Set XLAPP = Nothing
   Set XLBOOK = Nothing
   Set XLSHEET = Nothing
   '
   Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
   '
End Sub



Sub GENEXCEL_MSFH(NombreDeArchivo$, GRID As MSHFlexGrid)
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim valor As Double
   Dim TEXTO As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   TIPODATO% = 0: COLUM& = 0
   For i& = 1 To GRID.Rows
       RENG1& = i&: COLUM& = 0
       For COLU1& = 1 To GRID.Cols
          If GRID.ColWidth(COLU1& - 1) = 0 Then GoTo 10 ' OBVIA LAS COLUMNAS OCULTAS
          COLUM& = COLUM& + 1
          '
          VALORX = GRID.TextMatrix(i& - 1, COLU1 - 1)
'          If Len(VALORX) = 8 And Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
'             FEXU$ = VALORX
'             VALORX = FETEXCOR1$(Int(FECHANUM(FEXU$)))
'             TIPODATO% = 1
'          End If
          If IsDate(VALORX) Then
              If Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
                 If Len(VALORX) = 8 Then
                   VALORX = FETEXCOR1$(CVINT(VALORX))
                   TIPODATO% = 1
                 Else
                   TIPODATO% = 0
                   VALORX = "A " & VALORX
                 End If
              End If
          End If

          ' ESTO ES PARA QUE EXPORTE COMO CERO Y NO ESPACIOS
          FFCC$ = UCase$(Left$(FORCAMPO$(COLU1&), 1))
          If FFCC$ = "I" Or FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
             If TRIM$(CStr(VALORX)) = "" Then
                VALORX = 0
             End If
             TIPODATO% = 2
          End If
          If FFCC$ = "C" Or FFCC$ = "A" Then VALORX = "'" & VALORX
          '
          'PARA QUE NO REPITA ENCABEZADO
'          If RENG1& = 1 Then
'            If VALORX = VALORX_ANT And TRIM$(VALORX) <> "" Then
'              If COLUM& <= 17 Then
'               LETRAX1$ = Chr$(64 + COLUM& - 1)
'               LETRAX2$ = Chr$(64 + COLUM&)
'               RANXO$ = LETRAX1$ & TRIM$(Str$(RENG1&)) & ":" & LETRAX2$ & TRIM$(Str$(RENG1&))
'               Range(RANXO$).Select
'
'               XLSHEET.Cells.HorizontalAlignment = xlCenter
'               XLSHEET.Cells.HorizontalAlignment = xlCenter
'               XLSHEET.Cells.VerticalAlignment = xlCenter
'               XLSHEET.Cells.WrapText = False
'               XLSHEET.Cells.Orientation = 0
'               XLSHEET.Cells.AddIndent = False
'               XLSHEET.Cells.IndentLevel = 0
'               XLSHEET.Cells.ShrinkToFit = False
'               XLSHEET.Cells.ReadingOrder = xlContext
'               XLSHEET.Cells.MergeCells = True
'               GoTo 10
'             End If
'            End If
'          End If
          '
          XLSHEET.Cells(RENG1&, COLUM&) = VALORX
          If TIPODATO% = 1 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yyyy"
          'If TIPODATO% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#,##0.00"
          If TIPODATO% = 0 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
'          If TIPODATO% = 3 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "DD/MM/AA HH:MM:SS"
'          If TIPODATO% = 3 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "[$-2C0A] dd/mm/aa hh:mm:ss AM/PM;@"
          
          TIPODATO% = 0
          VALORX_ANT = VALORX
10     Next COLU1&
       Command1.Caption = "Registros: " & i&
   Next i&
   '
   Close #N1%
   On Error GoTo 30
   XLSHEET.SaveAs NombreDeArchivo$
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   Screen.MousePointer = 1
   MsgBox "Planilla Excel Creada Exitosamente: " & NombreDeArchivo$
   'Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NombreDeArchivo$, "\", "/"))
   '
Exit Sub
   '
   '
20 Resume 21
21 On Error GoTo 0
   MsgBox "Utilitario EXCEL no Instalado"
   On Error GoTo 0
   Exit Sub
   '
30 Resume 31
31 On Error GoTo 0
   Screen.MousePointer = 1
   '
   ' Cierra el objeto Book, el cual es utilizado al crear
   ' una nueva hoja en el archivo Excel.
   ' El valor False es para que no pregunte si quiere
   ' guardar el archivo.
   XLBOOK.Close (False)
   XLAPP.Quit ' Cierra Excel
   Set XLAPP = Nothing
   Set XLBOOK = Nothing
   Set XLSHEET = Nothing
   '
   Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
   '
End Sub





Private Sub Command3_Click()
   Screen.MousePointer = 11
   Me.GRID.Rows = 1: Me.MSF_2.Rows = 1
   Me.GRID.Visible = True
   Me.MSF_2.Visible = False
   GRIDH.Visible = False
   GRIDH2.Visible = False
   
   If MnuParticularesInstitucionesSinReservas.Caption = Command3.Caption Then Call PartSinReservas: GoTo 99
   If mnuReporteHistorico.Caption = Command3.Caption Then Call HISTORICO: GoTo 99
'   If mnuInstitucionesQueNoTuvieronReservas.Caption = Command3.Caption Then Call InstSinReservas: GoTo 99
   FE1% = 0: FE2% = 0
   Call ABRECONEXION
   CAMPOS$ = "F.FUNCION/D24;FECHA RESERVA/D24;OBRA/A4;INSTITUCION/A32;I-TELEFONO/A18;I-MAIL/A32"
   CAMPOS$ = CAMPOS$ + " ;I-Domicilio/A32;I-Ciudad/A18;I-Provincia/A16;I-C.Postal/A12"
   CAMPOS$ = CAMPOS$ + " ;PARTICULAR/A32;P-TELEFONO/A18;P-MAIL/A32"
   CAMPOS$ = CAMPOS$ + " ;P-Domicilio/A32;P-Ciudad/A18;P-Provincia/A16;P-C.Postal/A12"
   CAMPOS$ = CAMPOS$ + ";CANTIDAD/f8;Pr.Res/F7;PRECIO 1/F7;PRECIO 2/F7;PRECIO 3/F7;Presentes/F12"
   CAMPOS$ = CAMPOS$ + ";Status/A12;Orden1/F0;Orden2/F0;LastUpdate/A24;Locación/A10"

   Call CARGA_ENCABEZADO(Me.GRID, CAMPOS$, Me, 1)
   Call CARGA_ENCABEZADO(Me.MSF_2, CAMPOS$, Me, 1)
   
   Dim ESTAXO$(5)
   ESTAXO$(0) = "CANCELADA"
   ESTAXO$(1) = "EN ESPERA"
   ESTAXO$(2) = "CONFIRMADA"
   ESTAXO$(3) = "RECONFIRMADA"
   
   If IsDate(Me.DTPFechades) Then FE1% = CVINT(Me.DTPFechades.Value): If FE1% > 0 Then DES = FETEXCOR1$(FE1%)
   If IsDate(Me.DTPFechahas) Then FE2% = CVINT(Me.DTPFechahas.Value): If FE2% > 0 Then HAS = FETEXCOR1$(FE2%)
   SQLX$ = "select b.FechaReserva as freserva, F.fecha as ffuncion , f.obra,I.SCHOOLS AS NAMEINSTITUCION"
   SQLX$ = SQLX$ + ", f.idfuncion, b.IDFuncion,I.Telephones AS TELEFUN,i.Email AS MAILINST,P.name,P.Telephones AS TELEFPART"
   SQLX$ = SQLX$ + ", P.Email AS MAILPART,B.LASTUPDATE AS BLASTUPDATE"
   SQLX$ = SQLX$ + ", B.Cantidad ,  F.Precio_A, F.Precio_B, F.Precio_C, B.Presentes"
   SQLX$ = SQLX$ + ", B.ESTADO, B.IDParticular, B.IDInstitucion,f.Fecha,B.precio AS PRECIORESERVA"
   SQLX$ = SQLX$ + ", P.SCHOOLS AS NAMEPARTICULAR, F.LUGAR AS FLUGAR, I.address as IDomi"
   SQLX$ = SQLX$ + ", I.town as ICiudad, I.province as IProv, I.[Zip Code] as ICP "
   SQLX$ = SQLX$ + ", P.address as PDomi, P.town as PCiudad, P.province as PProv, P.[Zip Code] as PCP "
   SQLX$ = SQLX$ + " from bookings b WITH(NOLOCK) inner join funciones f on b.idfuncion = f.idfuncion "
   SQLX$ = SQLX$ + " left join instituciones I on  b.IDInstitucion =  I.id"
   SQLX$ = SQLX$ + "  inner join Particulares P on b.idparticular = P.id"
   SQLX$ = SQLX$ + " WHERE b.idreserva > 0 "
   TARDE% = 0
   If Len(dthorades.Value) > 10 Then
      'SI ES PM LE SUMA 12
      'VALIDO QUE SI ES PM QUE NO SEA >  QUE 12 POR QUE DA ERROR
      AMOPM$ = UCase(Right(dthorades.Value, 4))
      If AMOPM$ = "P.M." Then
         If XVALO(Mid(Me.dthorades.Value, 12, 2)) >= 12 Then
            If XVALO(Mid(Me.dthorades.Value, 15, 2)) > 0 Then
             MsgBox "SI EL HORARIO ES P.M. EL MISMO NO PUEDE SER MENOR QUE 12"
             GoTo 99
            End If
         End If
         TARDE% = 12
      End If
      HORADESDE$ = SAFETEXT$(XVALO(Mid(Me.dthorades.Value, 12, 2)) + TARDE%)
      fechayhorades = DTPFechades.Value & " " & HORADESDE$ & Mid(Me.dthorades.Value, 14, 6)
      SQLX$ = SQLX$ + " AND f.Fecha   >= '" & fechayhorades & "'"
   Else
      If FE1% > 0 Then SQLX$ = SQLX$ + " AND f.Fecha  >= '" & DES & "'"
   End If
   
   TARDE% = 0
   If Len(dthorahas.Value) > 10 Then
      AMOPM$ = UCase(Right(dthorahas.Value, 4))
      If AMOPM$ = "P.M." Then
        If XVALO(Mid(Me.dthorahas.Value, 12, 2)) >= 12 Then
           If XVALO(Mid(Me.dthorahas.Value, 15, 2)) > 0 Then
              MsgBox "SI EL HORARIO ES P.M. EL MISMO NO PUEDE SER MENOR QUE 12"
              GoTo 99
            End If
         End If
         TARDE% = 12
      End If
      
      HORAHASTA$ = SAFETEXT$(XVALO(Mid(Me.dthorahas.Value, 12, 2)) + TARDE%)
      fechayhorahas = DTPFechahas.Value & " " & HORAHASTA$ & Mid(Me.dthorahas.Value, 14, 6)
      'fechayhorahas = DTPFechahas.Value & " " & Mid(Me.dthorahas.Value, 12, 8)
      SQLX$ = SQLX$ + " AND f.Fecha <= '" & fechayhorahas & "'"
   Else
      If FE2% > 0 Then SQLX$ = SQLX$ + " AND f.Fecha  <= '" & HAS & "'"
   End If
   
   SQLX$ = SQLX$ + " ORDER BY f.Fecha ASC"
   J& = 0
   
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF
       J& = J& + 1
       Label3.Caption = "Registros Procesados: " & J&
       Label3.Refresh
       MUESTRABAP.GRID.Rows = J& + 1
       MUESTRABAP.GRID.TextMatrix(J&, 1) = SAFETEXT$(!ffuncion)
       MUESTRABAP.GRID.TextMatrix(J&, 2) = SAFETEXT$(!freserva)
       MUESTRABAP.GRID.TextMatrix(J&, 3) = SAFETEXT$(!OBRA)
       MUESTRABAP.GRID.TextMatrix(J&, 4) = SAFETEXT$(!NAMEINSTITUCION)
       MUESTRABAP.GRID.TextMatrix(J&, 5) = SAFETEXT$(!TELEFUN)
       MUESTRABAP.GRID.TextMatrix(J&, 6) = SAFETEXT$(!MAILINST)
       
       MUESTRABAP.GRID.TextMatrix(J&, 7) = SAFETEXT$(!IDomi)
       MUESTRABAP.GRID.TextMatrix(J&, 8) = SAFETEXT$(!ICiudad)
       MUESTRABAP.GRID.TextMatrix(J&, 9) = SAFETEXT$(!IProv)
       MUESTRABAP.GRID.TextMatrix(J&, 10) = SAFETEXT$(!ICP)
       
       MUESTRABAP.GRID.TextMatrix(J&, 11) = SAFETEXT$(!NAMEPARTICULAR)
       MUESTRABAP.GRID.TextMatrix(J&, 12) = SAFETEXT$(!TELEFPART)
       MUESTRABAP.GRID.TextMatrix(J&, 13) = SAFETEXT$(!MAILPART)
       
       MUESTRABAP.GRID.TextMatrix(J&, 14) = SAFETEXT$(!PDomi)
       MUESTRABAP.GRID.TextMatrix(J&, 15) = SAFETEXT$(!PCiudad)
       MUESTRABAP.GRID.TextMatrix(J&, 16) = SAFETEXT$(!PProv)
       MUESTRABAP.GRID.TextMatrix(J&, 17) = SAFETEXT$(!PCP)
       
       MUESTRABAP.GRID.TextMatrix(J&, 18) = XVALO(!Cantidad)
       MUESTRABAP.GRID.TextMatrix(J&, 19) = XVALO(!PRECIORESERVA)
       MUESTRABAP.GRID.TextMatrix(J&, 20) = XVALO(!Precio_A)
       MUESTRABAP.GRID.TextMatrix(J&, 21) = XVALO(!Precio_B)
       MUESTRABAP.GRID.TextMatrix(J&, 22) = XVALO(!Precio_C)
        MUESTRABAP.GRID.TextMatrix(J&, 23) = XVALO(!PRESENTES)
        EST% = XVALO(!ESTADO)
       MUESTRABAP.GRID.TextMatrix(J&, 24) = ESTAXO$(EST% + 1) 'POR QUE EN LA TABLA GUARDA -1, 0,1,2
       
       MUESTRABAP.GRID.TextMatrix(J&, 25) = FechaHoraNum(!ffuncion)
       MUESTRABAP.GRID.TextMatrix(J&, 26) = FechaHoraNum(!freserva)
       MUESTRABAP.GRID.TextMatrix(J&, 27) = SAFETEXT$(!BLASTUPDATE)
       MUESTRABAP.GRID.TextMatrix(J&, 28) = SAFETEXT$(!FLUGAR)

       .MoveNext
     Loop
   End With
   RS.Close
   Set RS = Nothing
99 Screen.MousePointer = 1
End Sub

Sub PartSinReservas()
     Screen.MousePointer = 11
     Me.GRID.Visible = True
     Me.MSF_2.Visible = False
   
     GRIDH.Visible = False
     GRIDH2.Visible = False

     Me.GRID.Rows = 1: Me.MSF_2.Rows = 1
'     Frame2.Visible = False
     MUESTRABAP.GRID.Redraw = False
     Call ABRECONEXION
     CAMPOS$ = "ID/f10;Particular/A36;Institución/A36,Teléfono/A32;Mail/A36;Last Update/A28;Orden L.U./F0;Domicilio/A20;Ciudad/A20;Provincia/A20;C.Postal/A12"
     Call CARGA_ENCABEZADO(Me.GRID, CAMPOS$, Me, 1)
     Call CARGA_ENCABEZADO(Me.MSF_2, CAMPOS$, Me, 1)
     'CARGO EL ECO DE LA INSTITUCION
     Dim INSTITU$()
     SQLX$ = "SELECT ID, Schools FROM INSTITUCIONES WITH(NOLOCK) ORDER BY ID "
     Set RS0 = READSET(SQLX$)
     With RS0
     Do While Not .EOF
        ID& = XVALO(!ID)
        SCHOLS$ = SAFETEXT$(!SCHOOLS)
        ReDim Preserve INSTITU$(ID&)
        INSTITU$(ID&) = SCHOLS$
        .MoveNext
     Loop
     End With
     RS0.Close
     Set RS0 = Nothing
     'FIN ECO DE LA INSTITUCION
     
     SQLX$ = "Select ID, NAME, SCHOOLS,School, Telephones, Email,LastUpdate ,ADDRESS,TOWN, PROVINCE,[ZIP CODE] "
     SQLX$ = SQLX$ + " from particulares  WITH(NOLOCK)"
     SQLX$ = SQLX$ + " where PARTICULARES.ID >= 0  AND  not exists (select 1 from BOOKINGS WHERE BOOKINGS.IDPARTICULAR = particulares.ID "
     
     
     'CONDICION DE FECHA
     If IsDate(Me.DTPFechades) Then FE1% = CVINT(Me.DTPFechades.Value): If FE1% > 0 Then DES = FETEXCOR1$(FE1%)
     If IsDate(Me.DTPFechahas) Then FE2% = CVINT(Me.DTPFechahas.Value): If FE2% > 0 Then HAS = FETEXCOR1$(FE2%)
     
     If Len(dthorades.Value) > 10 Then
       'SI ES PM LE SUMA 12
       AMOPM$ = UCase(Right(dthorades.Value, 4))
       If AMOPM$ = "P.M." Then TARDE% = 12
       HORADESDE$ = SAFETEXT$(XVALO(Mid(Me.dthorades.Value, 12, 2)) + TARDE%)
       fechayhorades = DTPFechades.Value & " " & HORADESDE$ & Mid(Me.dthorades.Value, 14, 6)
       SQLX$ = SQLX$ + " AND fechaReserva   >= '" & fechayhorades & "'"
      Else
       If FE1% > 0 Then SQLX$ = SQLX$ + " AND fechaReserva  >= '" & DES & "'"
     End If
     TARDE% = 0
     If Len(dthorahas.Value) > 10 Then
       AMOPM$ = UCase(Right(dthorahas.Value, 4))
       If AMOPM$ = "P.M." Then TARDE% = 12
       HORAHASTA$ = SAFETEXT$(XVALO(Mid(Me.dthorahas.Value, 12, 2)) + TARDE%)
       fechayhorahas = DTPFechahas.Value & " " & HORAHASTA$ & Mid(Me.dthorahas.Value, 14, 6)
       SQLX$ = SQLX$ + " AND fechaReserva <= '" & fechayhorahas & "'"
     Else
       If FE2% > 0 Then SQLX$ = SQLX$ + " AND fechaReserva  <= '" & HAS & "'"
     End If
     'CIERRO EL PARENTESIS
     SQLX$ = SQLX$ + ")"
     
     Set RS = READSET(SQLX$)
     J& = 0
     With RS
     Do While Not .EOF
       J& = J& + 1
       Label3.Caption = "Registros Procesados: " & J&
       Label3.Refresh

       MUESTRABAP.GRID.Rows = J& + 1
       MUESTRABAP.GRID.TextMatrix(J&, 1) = XVALO(!ID)
       MUESTRABAP.GRID.TextMatrix(J&, 2) = SAFETEXT$(!SCHOOLS)
       MUESTRABAP.GRID.TextMatrix(J&, 3) = INSTITU$(XVALO(!School))
       MUESTRABAP.GRID.TextMatrix(J&, 4) = SAFETEXT$(!Telephones)
       MUESTRABAP.GRID.TextMatrix(J&, 5) = SAFETEXT$(!Email)
'       FE% = CVINT(!LASTUPDATE)
       MUESTRABAP.GRID.TextMatrix(J&, 6) = SAFETEXT$(!LASTUPDATE)
       MUESTRABAP.GRID.TextMatrix(J&, 7) = FechaHoraNum(SAFETEXT$(!LASTUPDATE))
       
       MUESTRABAP.GRID.TextMatrix(J&, 8) = SAFETEXT$(!Address)
       MUESTRABAP.GRID.TextMatrix(J&, 9) = SAFETEXT$(!TOWN)
       MUESTRABAP.GRID.TextMatrix(J&, 10) = SAFETEXT$(!PROVINCE)
       MUESTRABAP.GRID.TextMatrix(J&, 11) = SAFETEXT$(![ZIP CODE])
      
       
       .MoveNext
     Loop
     End With
     RS.Close
     Set RS = Nothing
     GRID.COL = 1
     GRID.Sort = 3
     'CARGA BAJO LAS MISMAS CONDICIONES LOS INSTITUOS QUE NO TUVIERON RESERVAS.
     Call InstSinReservas
99   Screen.MousePointer = 1
     MUESTRABAP.GRID.Redraw = True
End Sub


Sub InstSinReservas()
     Screen.MousePointer = 11
     MUESTRABAP.GRID.Redraw = False
     Call ABRECONEXION
     'CAMPOS$ = "ID/f10;Institución/A24;Teléfono/A24;Mail/A32;Last Update/A40;Orden L.U./F0"
'     Call CARGA_ENCABEZADO(Me.GRID, CAMPOS$, Me, 1)
'     Call CARGA_ENCABEZADO(Me.MSF_2, CAMPOS$, Me, 1)
     SQLX$ = "Select ID, NAME, SCHOOLS, Telephones, Email ,LastUpdate ,ADDRESS,TOWN, PROVINCE,[ZIP CODE] "
     SQLX$ = SQLX$ + " from INSTITUCIONES  WITH(NOLOCK)"
     SQLX$ = SQLX$ + " where ID >= 0 AND  not exists (select 1 from BOOKINGS WHERE BOOKINGS.IDInstitucion = INSTITUCIONES.ID"
     
     'CONDICION DE FECHA
     If IsDate(Me.DTPFechades) Then FE1% = CVINT(Me.DTPFechades.Value): If FE1% > 0 Then DES = FETEXCOR1$(FE1%)
     If IsDate(Me.DTPFechahas) Then FE2% = CVINT(Me.DTPFechahas.Value): If FE2% > 0 Then HAS = FETEXCOR1$(FE2%)
     
     If Len(dthorades.Value) > 10 Then
       'SI ES PM LE SUMA 12
       AMOPM$ = UCase(Right(dthorades.Value, 4))
       If AMOPM$ = "P.M." Then TARDE% = 12
       HORADESDE$ = SAFETEXT$(XVALO(Mid(Me.dthorades.Value, 12, 2)) + TARDE%)
       fechayhorades = DTPFechades.Value & " " & HORADESDE$ & Mid(Me.dthorades.Value, 14, 6)
       SQLX$ = SQLX$ + " AND fechaReserva   >= '" & fechayhorades & "'"
      Else
       If FE1% > 0 Then SQLX$ = SQLX$ + " AND fechaReserva  >= '" & DES & "'"
     End If
     TARDE% = 0
     If Len(dthorahas.Value) > 10 Then
       AMOPM$ = UCase(Right(dthorahas.Value, 4))
       If AMOPM$ = "P.M." Then TARDE% = 12
       HORAHASTA$ = SAFETEXT$(XVALO(Mid(Me.dthorahas.Value, 12, 2)) + TARDE%)
       fechayhorahas = DTPFechahas.Value & " " & HORAHASTA$ & Mid(Me.dthorahas.Value, 14, 6)
       SQLX$ = SQLX$ + " AND fechaReserva <= '" & fechayhorahas & "'"
     Else
       If FE2% > 0 Then SQLX$ = SQLX$ + " AND fechaReserva  <= '" & HAS & "'"
     End If
     'CIERRO EL PARENTESIS
     SQLX$ = SQLX$ + ")"

     Set RS = READSET(SQLX$)
     J& = Me.GRID.Rows
     With RS
     Do While Not .EOF
       J& = J& + 1
       MUESTRABAP.GRID.Rows = J& + 1
       MUESTRABAP.GRID.TextMatrix(J&, 1) = XVALO(!ID)
       MUESTRABAP.GRID.TextMatrix(J&, 2) = ""
       MUESTRABAP.GRID.TextMatrix(J&, 3) = SAFETEXT$(!SCHOOLS)
       MUESTRABAP.GRID.TextMatrix(J&, 4) = SAFETEXT$(!Telephones)
       MUESTRABAP.GRID.TextMatrix(J&, 5) = SAFETEXT$(!Email)
       'FE% = CVINT(!LASTUPDATE)
       MUESTRABAP.GRID.TextMatrix(J&, 6) = SAFETEXT$(!LASTUPDATE)
       MUESTRABAP.GRID.TextMatrix(J&, 7) = FechaHoraNum(SAFETEXT$(!LASTUPDATE))
       MUESTRABAP.GRID.TextMatrix(J&, 8) = SAFETEXT$(!Address)
       MUESTRABAP.GRID.TextMatrix(J&, 9) = SAFETEXT$(!TOWN)
       MUESTRABAP.GRID.TextMatrix(J&, 10) = SAFETEXT$(!PROVINCE)
       MUESTRABAP.GRID.TextMatrix(J&, 11) = SAFETEXT$(![ZIP CODE])
       
       .MoveNext
     Loop
     End With
     RS.Close
     Set RS = Nothing
'     GRID.COL = 1
'     GRID.Sort = 3
     
99   Screen.MousePointer = 1
     MUESTRABAP.GRID.Redraw = True
End Sub


Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Shift = 1 Then
     If Button = 2 Then
       TXTO$ = "Reporte Histórico" & vbCrLf
       TXTO$ = TXTO$ + "Tener en Cuenta: " + vbCrLf
       TXTO$ = TXTO$ + "QUE EN EL RANGO SELECCIONADO SE CONSIDERA LA FECHA DE LA FUNCION " + vbCrLf
       TXTO$ = TXTO$ + "Y A LA VEZ LA FECHA DE RESERVA " + vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf

       TXTO$ = TXTO$ + "CASO DE EJEMPLO:  " + vbCrLf
       TXTO$ = TXTO$ + " ELIJO EL MES DE AGOSTO Y EN LA MISMA TENGO UNA FUNCION CON FECHA AGOSTO LA MISMA SE VA  A MOSTRAR"
       TXTO$ = TXTO$ + " PERO SI LA RESERVA PARA DICHA FUNCION SE REALIZO EN OTRO MES ANTERIR A AGOSTO "
       TXTO$ = TXTO$ + " TAMBIEN VA A MOSTRAR OTRO REGISTRO INDICANDO QUE NO REALIZO RESERVAS"
       TXTO$ = TXTO$ + " YA QUE DENTRO DEL MES DE AGOSTO NO REALIZO NINGUNA RESERVA"
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       MsgBox TXTO$
     End If
   End If

End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call filtrar
   Command4.Enabled = True
End Sub

Private Sub Form_Load()
'   GRID.Cols = 5
'   For i = 1 To 100
'      GRID.Rows = i + 1
'      GRID.TextMatrix(i, 1) = "AAAA" & i
'      GRID.TextMatrix(i, 2) = "BBBB" & i
'      GRID.TextMatrix(i, 3) = "CCCC" & i
'      GRID.TextMatrix(i, 4) = "DDDD" & i
'   Next i
   
'   'EN LA FUNCION DE CARGAR LOS ENCABEZADOS HABRIA QUE ASIGNAR LOS ANCHOS DE LABEL 5 Y ASIGNARSELO A LABEL8
'   For U& = 1 To GRID.Cols
'          JI& = JI& + 1
'          ' PROVISORIAMENTE A VER SI FUNCIONA
'          'AQUI AGREGAR A LABEL5 LA CANTIDAD DE CARACTERES
'          ANCH& = 20
'          Label5.Caption = FORMATNUM$("9999999.9999999", "A25" + "*")
'          GRID.ColWidth(JI& - 1) = Label5.Width
'25        If JI& >= Label8.Count Then
'             NUPIC% = Label8.Count
'             Load Label8(NUPIC%)
'             GoTo 25
'          End If
'          Label8(U&) = Label5.Width
'   Next U&
'   '
'   Label7 = 100
   '
End Sub






Private Sub Form_Resize()
    On Error Resume Next
    Me.MSF_2.Height = Me.Height - Frame2.Height - (Me.Command1.Height + 1000)
    Me.MSF_2.Width = Me.Width - 550
    Me.GRID.Height = Me.Height - Frame2.Height - (Me.Command1.Height + 1000)
    Me.GRID.Width = Me.Width - 550
    Me.GRIDH.Height = Me.Height - Frame2.Height - (Me.Command1.Height + 1000)
    Me.GRIDH.Width = Me.Width - 550
    Me.GRIDH2.Height = Me.Height - Frame2.Height - (Me.Command1.Height + 1000)
    Me.GRIDH2.Width = Me.Width - 550

    On Error GoTo 0
    Frame2.Visible = True
'    Frame2.Top = Me.Height - (Frame2.Height + 400)
    Frame1.Visible = True
    Frame1.Top = Me.Height - (Frame2.Height + 700)
    Me.Command1.Top = Me.Height - (Me.Command1.Height + 900)
'    Me.Command2.Top = Me.Height - Me.Command2.Height + 700

End Sub


Private Sub GRID_Click()
    Static MODO  As Boolean
    
    If (GRID.MouseRow = 0) Then
        COLUX% = GRID.MouseCol
        If TRIM$(mnuReporte1.Caption) = "Reporte General Reservas x Fechas" Then
           'en este caso ordena por la columna que se cargo como cvint de la fecha
           If GRID.MouseCol = 1 Then COLUX% = 25
           If GRID.MouseCol = 2 Then COLUX% = 26
        End If
        If MnuParticularesInstitucionesSinReservas.Caption = Command3.Caption Then
           'en este caso ordena por la columna que se cargo como cvint de la fecha
           If GRID.MouseCol = 6 Then COLUX% = 7
        End If
'        If mnuInstitucionesQueNoTuvieronReservas.Caption = Command3.Caption Then
'           'en este caso ordena por la columna que se cargo como cvint de la fecha
'           If GRID.MouseCol = 6 Then COLUX% = 7
'        End If

        ' Ordena en forma ascendente
        If MODO Then
            GRID.COL = COLUX%
            GRID.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            GRID.COL = COLUX%
            GRID.Sort = 1
            MODO = True
        End If
    End If
End Sub

Private Sub mnuInstitucionesQueNoTuvieronReservas_Click()
   Command3.Caption = mnuInstitucionesQueNoTuvieronReservas.Caption
   Frame2.Visible = False
End Sub

Private Sub MnuParticularesInstitucionesSinReservas_Click()
    Command3.Caption = MnuParticularesInstitucionesSinReservas.Caption
'    Frame2.Visible = False
End Sub

Private Sub mnuReporte1_Click()
   If mnuReporte1.Caption <> Command3.Caption Then
      GRID.Rows = 1
      MSF_2.Rows = 1
      Command3.Caption = mnuReporte1.Caption
     
   End If
   '
   Frame2.Visible = True
   DTPFechades.ZOrder 0
   DTPFechahas.ZOrder 0
   Me.dthorades.ZOrder 0
   Me.dthorahas.ZOrder 0
   '
End Sub

Private Sub mnuReporteHistorico_Click()
   Command3.Caption = mnuReporteHistorico.Caption
End Sub

Private Sub mnuSalir_Click()
  Unload Me
End Sub

Private Sub MSF_2_Click()
    Static MODO  As Boolean
    If (MSF_2.MouseRow = 0) Then
        COLUX% = MSF_2.MouseCol
        If TRIM$(mnuReporte1.Caption) = "Reporte General Reservas x Fechas" Then
           'en este caso ordena por la columna que se cargo como cvint de la fecha
           If MSF_2.MouseCol = 1 Then COLUX% = 25
           If MSF_2.MouseCol = 2 Then COLUX% = 26
        End If
        If MnuParticularesInstitucionesSinReservas.Caption = Command3.Caption Then
           'en este caso ordena por la columna que se cargo como cvint de la fecha
           If MSF_2.MouseCol = 5 Then COLUX% = 6
        End If
'        If mnuInstitucionesQueNoTuvieronReservas.Caption = Command3.Caption Then
'           'en este caso ordena por la columna que se cargo como cvint de la fecha
'           If MSF_2.MouseCol = 5 Then COLUX% = 6
'        End If
  
        ' Ordena en forma ascendente
        If MODO Then
            MSF_2.COL = COLUX%
            MSF_2.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSF_2.COL = COLUX%
            MSF_2.Sort = 1
            MODO = True
        End If
    End If
End Sub
Private Sub GRID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
'    '
'    I& = GRID.MouseRow
'    j& = GRID.MouseCol
'    If I& = 0 Or I& > GRID.Rows Then Exit Sub
'    '
'    On Error Resume Next
'    GRID.Row = I&
'    GRID.COL = j&
'    On Error Resume Next
'
'    NPX& = TRIM$(GRID.TextMatrix(I&, 1))
'    If NPX& < 1 Then
'        Exit Sub
'    End If
'
'    If Me.Caption = "Pedidos" Then
'        '
'        NUDOCU$ = TRIM$(Str$(NPX&))
'        TIDOCUM$ = "Pedido de Cliente"
'        While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
'        NUDOCU$ = NUDOCU$ + " "
'        Screen.MousePointer = 11
'        '
'        'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
'        DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
''        If DOCUNU& < 1 Then
''             Screen.MousePointer = 1
''             Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a la " & DOCUMENTO_ELEGIDO$ & " .\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
''             On Error Resume Next
''             OPPED07.Command46.SetFocus
''             OPPED07.Show 1
''             If Err Then
''              On Error GoTo 0
''             End If
''             Exit Sub
''        End If
'        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
'    Else
'        TIDOCUM$ = "Nota de Pedido"
'        NUDOCU$ = TRIM$(Str$(NPX&))
'        DOCUNU& = 0
'        '
'        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$)
'    End If
'    Screen.MousePointer = 1
  End If
End Sub
Private Sub MSF_2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    '
    i& = MSF_2.MouseRow
    J& = MSF_2.MouseCol
    If i& = 0 Or i& > MSF_2.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF_2.Row = i&
    MSF_2.COL = J&
    On Error Resume Next

    NPX& = TRIM$(MSF_2.TextMatrix(i&, 1))
    If NPX& < 1 Then
        Exit Sub
    End If

    If Me.Caption = "Pedidos" Then
        '
        NUDOCU$ = TRIM$(Str$(NPX&))
        TIDOCUM$ = "Pedido de Cliente"
        While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
        NUDOCU$ = NUDOCU$ + " "
        Screen.MousePointer = 11
        '
        'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
        DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
'        If DOCUNU& < 1 Then
'             Screen.MousePointer = 1
'             Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a la " & DOCUMENTO_ELEGIDO$ & " .\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
'             On Error Resume Next
'             OPPED07.Command46.SetFocus
'             OPPED07.Show 1
'             If Err Then
'              On Error GoTo 0
'             End If
'             Exit Sub
'        End If
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
    Else
        TIDOCUM$ = "Nota de Pedido"
        NUDOCU$ = TRIM$(Str$(NPX&))
        DOCUNU& = 0
        '
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$)
    End If
    Screen.MousePointer = 1
  End If
End Sub

Private Sub msf2_Click()
    Static MODO  As Boolean
    If (msf2.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            msf2.COL = msf2.MouseCol
            msf2.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            msf2.COL = msf2.MouseCol
            msf2.Sort = 1
            MODO = True
        End If
    End If

End Sub

Private Sub MSF2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    '
    i& = msf2.MouseRow
    J& = msf2.MouseCol
    If i& = 0 Or i& > msf2.Rows Then Exit Sub
    '
    On Error Resume Next
    msf2.Row = i&
    msf2.COL = J&
    On Error Resume Next

    NPX& = TRIM$(msf2.TextMatrix(i&, 1))
    If NPX& < 1 Then
        Exit Sub
    End If

    If Me.Caption = "Pedidos" Then
        '
        NUDOCU$ = TRIM$(Str$(NPX&))
        TIDOCUM$ = "Pedido de Cliente"
        While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
        NUDOCU$ = NUDOCU$ + " "
        Screen.MousePointer = 11
        '
        'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
        DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
'        If DOCUNU& < 1 Then
'             Screen.MousePointer = 1
'             Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a la " & DOCUMENTO_ELEGIDO$ & " .\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
'             On Error Resume Next
'             OPPED07.Command46.SetFocus
'             OPPED07.Show 1
'             If Err Then
'              On Error GoTo 0
'             End If
'             Exit Sub
'        End If
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
    Else
        TIDOCUM$ = "Nota de Pedido"
        NUDOCU$ = TRIM$(Str$(NPX&))
        DOCUNU& = 0
        '
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$)
    End If
    Screen.MousePointer = 1
  End If
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      On Error Resume Next
      Command3.SetFocus
      On Error GoTo 0
   End If

End Sub

Private Sub TXTBUSCAR_Change()
    '
'    If GRIDH.Visible = True Then
'       Dim GRID As MSHFlexGrid
'       Dim MSF_2 As MSHFlexGrid
'       Set GRID = GRIDH
'       Set MSF_2 = GRIDH2
'    End If
 
     If TRIM$(TXTBUSCAR) = "" Then
      If GRIDH.Visible Or GRIDH2.Visible Then
        GRIDH2.Visible = False
        GRIDH.Visible = True
        Call PINTATEXT(TXTBUSCAR)
        Exit Sub
      End If
     End If
     
     If TRIM$(TXTBUSCAR) = "" Then
      MSF_2.Visible = False
      GRID.Visible = True
      Call PINTATEXT(TXTBUSCAR)
      Exit Sub
     End If

    '
End Sub

Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      TXTBUSCAR.TEXT = ""
      Call PINTATEXT(TXTBUSCAR)
   End If
   
   If KeyAscii = 13 Then
     Call Command4_Click
     Call PINTATEXT(TXTBUSCAR)
   End If
End Sub
Sub HISTORICO()
   Screen.MousePointer = 11
   'Me.GRIDH.Rows = 1: Me.MSF_2.Rows = 1
   Me.GRID.Visible = False
   Me.MSF_2.Visible = False
   
   FE1% = 0: FE2% = 0
'   Call ABRECONEXION
   GRIDH.Visible = True
   GRIDH2.Visible = False
   CAMPOS$ = "I-ID/F5;Institución/A24;I-Teléfono/A20;I-Mail/A32;I-Domicilio/A24;I-Ciudad/A32"
   CAMPOS$ = CAMPOS$ + " ;I-Provincia/A32;I-C.Postal/A12"
   CAMPOS$ = CAMPOS$ + " ;ID/F7;PARTICULAR/A32;P-TELEFONO/A18;P-MAIL/A32"
   CAMPOS$ = CAMPOS$ + " ;P-Domicilio/A32;P-Ciudad/A18;P-Provincia/A16;P-C.Postal/A12"
   CAMPOS$ = CAMPOS$ + " ;LastUpdate/A24;Status/A12;F.Reserva/A20;F.Función/A20;Obra/A14;Locación/A20"
   CAMPOS$ = CAMPOS$ + ";CANTIDAD/f8;PRECIO/F7;PRECIO 1/F7;PRECIO 2/F7;PRECIO 3/F7;Presentes/F12;OFUNCION/F0;ORESERVA/F0"
   Call CARGA_ENCABEZADOHERICAL(Me.GRIDH, CAMPOS$, Me, 1)
   Call CARGA_ENCABEZADOHERICAL(Me.GRIDH2, CAMPOS$, Me, 1)
   GRIDH.Rows = 2
   GRIDH2.Rows = 2
   GRIDH.FixedRows = 1
   GRIDH2.FixedRows = 1
   If IsDate(Me.DTPFechades) Then FE1% = CVINT(Me.DTPFechades.Value): If FE1% > 0 Then DES = FETEXCOR1$(FE1%, "HASTA")
   If IsDate(Me.DTPFechahas) Then FE2% = CVINT(Me.DTPFechahas.Value): If FE2% > 0 Then HAS = FETEXCOR1$(FE2%, "HASTA")
   
   If FE1% < 1 Then
       MsgBox "Debe Seleccionar una Fecha Desde"
       GoTo 99
   End If
   If FE2% < 1 Then
      MsgBox "Debe Seleccionar una Fecha Desde"
      GoTo 99
   End If
   'CRITERIO DEL FOR
   'RECORRE DESDE EL AÑO DESDE HASTA EL AÑO HASTA
   'ASIGNANDO AL PRIMERO LA FECHA REAL AL IGUAL QUE AL ULTIMO
   'EJEMPLO RANGO DE FECHAS DESDE: 15/07/12 AL 30/07/14
   'PRIMERO VA A SER DESDE EL 15/07/12 HASTA EL 31/12/12
   'EL SIGUIENTE DESDE EL 01/01/13 HASTA EL 31/12/13
   'Y EL ULTIMO DEL 01/01/14 HASTA EL 30/07/14
   DES = FETEXCOR1$(FE1%)
   HAS = FETEXCOR1$(FE2%)
   ANODES& = XVALO(Right(DES, 4))
   ANOHAS& = XVALO(Right(HAS, 4))
   For i& = ANODES To ANOHAS
       FEX1% = FECHANUM("01/01/" & Right(SAFETEXT$(i&), 2)): DESDEX = FETEXCOR1$(FEX1%)
       FEX2% = FECHANUM("31/12/" & Right(SAFETEXT$(i&), 2)): HASTAX = FETEXCOR1$(FEX2%, "HASTA")
       If i& = ANODES& Then
          DESDEX = DES
       End If
       If i& = ANOHAS& Then
          HASTAX = HAS
       End If
       Label3.Caption = "Procesando Reservas Año: " & Year(DESDEX)
       Call RESERVASSEGUNAÑO(DESDEX, HASTAX)
       Label3.Caption = "Procesando Particulares Sin Reserva Año: " & Year(DESDEX)
       Call PARTICULARSINRESERVASEGUNAÑO(DESDEX, HASTAX)
       Label3.Caption = "Procesando Instituciones Sin Reserva Año: " & Year(DESDEX)
       Call INSTITUCIONESSINRESERVA(DESDEX, HASTAX)
   Next i&
   '
99 Screen.MousePointer = 1
End Sub
Sub PARTICULARSINRESERVASEGUNAÑO(DES, HAS)
     Screen.MousePointer = 11
     MUESTRABAP.GRIDH.Redraw = False
     Call ABRECONEXION
     'CARGO EL ECO DE LA INSTITUCION
     ANO$ = Year(DES)
     Dim INSTITU$(), DOMI$(), CIUDAD$(), CPOS(), PROV$(), TELEFONO$(), MAIL$()
     SQLX$ = "SELECT ID, Schools,[zip code] AS ICP,Town ,Address, PROVINCE, TELEPHONES, EMAIL"
     SQLX$ = SQLX$ + " FROM INSTITUCIONES WITH(NOLOCK) ORDER BY ID "
     Set RS0 = READSET(SQLX$)
     With RS0
     Do While Not .EOF
        ID& = XVALO(!ID)
        SCHOLS$ = SAFETEXT$(!SCHOOLS)
        ReDim Preserve INSTITU$(ID&), DOMI$(ID&), CIUDAD$(ID&), CPOS(ID&), PROV$(ID&), TELEFONO$(ID&), MAIL$(ID&)
        INSTITU$(ID&) = SCHOLS$
        DOMI$(ID&) = SAFETEXT$(!Address)
        CIUDAD$(ID&) = SAFETEXT$(!TOWN)
        CPOS(ID&) = SAFETEXT$(!ICP)
        PROV$(ID&) = SAFETEXT$(!PROVINCE)
        TELEFONO$(ID&) = SAFETEXT$(!Telephones)
        MAIL$(ID&) = SAFETEXT$(!Email)
        .MoveNext
     Loop
     End With
     RS0.Close
     Set RS0 = Nothing
     'FIN ECO DE LA INSTITUCION
     
     SQLX$ = "Select ID, NAME, SCHOOLS,School, Telephones, Email,LastUpdate, Address,town,PROVINCE, [ZIP CODE]  from particulares  WITH(NOLOCK)"
     SQLX$ = SQLX$ + " where PARTICULARES.ID >= 0  AND  not exists (select 1 from BOOKINGS WHERE BOOKINGS.IDPARTICULAR = particulares.ID "
     SQLX$ = SQLX$ + " AND fechaReserva  >= '" & DES & "'"
     SQLX$ = SQLX$ + " AND fechaReserva  <= '" & HAS & "'"
     SQLX$ = SQLX$ + ")"
     
     Set RS = READSET(SQLX$)
     J& = MUESTRABAP.GRIDH.Rows
     With RS
     Do While Not .EOF
       MUESTRABAP.GRIDH.Rows = J& + 1
          IND& = XVALO(!School)

       MUESTRABAP.GRIDH.TextMatrix(J&, 0) = ANO$
       MUESTRABAP.GRIDH.TextMatrix(J&, 1) = IND& 'SAFETEXT$(!ffuncion)
       MUESTRABAP.GRIDH.TextMatrix(J&, 2) = INSTITU$(IND&)
       MUESTRABAP.GRIDH.TextMatrix(J&, 3) = TELEFONO$(IND&)
       MUESTRABAP.GRIDH.TextMatrix(J&, 4) = MAIL$(IND&)
       MUESTRABAP.GRIDH.TextMatrix(J&, 5) = DOMI$(IND&)
       MUESTRABAP.GRIDH.TextMatrix(J&, 6) = CIUDAD$(IND&)
       MUESTRABAP.GRIDH.TextMatrix(J&, 7) = PROV$(IND&)
       MUESTRABAP.GRIDH.TextMatrix(J&, 8) = CPOS(IND&)
       
       MUESTRABAP.GRIDH.TextMatrix(J&, 9) = XVALO(!ID)
       MUESTRABAP.GRIDH.TextMatrix(J&, 10) = SAFETEXT$(!SCHOOLS)
       MUESTRABAP.GRIDH.TextMatrix(J&, 11) = SAFETEXT$(!Telephones) 'SAFETEXT$(!ICP)
       
       MUESTRABAP.GRIDH.TextMatrix(J&, 12) = SAFETEXT$(!Email)
       MUESTRABAP.GRIDH.TextMatrix(J&, 13) = SAFETEXT$(!Address)
       MUESTRABAP.GRIDH.TextMatrix(J&, 14) = SAFETEXT$(!TOWN) 'SAFETEXT$(!MAILPART)
       
       MUESTRABAP.GRIDH.TextMatrix(J&, 15) = SAFETEXT$(!PROVINCE) 'SAFETEXT$(!PDomi)
       MUESTRABAP.GRIDH.TextMatrix(J&, 16) = SAFETEXT$(![ZIP CODE]) 'SAFETEXT$(!PCiudad)
       MUESTRABAP.GRIDH.TextMatrix(J&, 17) = SAFETEXT$(!LASTUPDATE) 'SAFETEXT$(!PProv)
       MUESTRABAP.GRIDH.TextMatrix(J&, 18) = "NO RESERVÓ" 'SAFETEXT$(!PCP)
       
       MUESTRABAP.GRIDH.TextMatrix(J&, 19) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 20) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 21) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 22) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 23) = ""
        MUESTRABAP.GRIDH.TextMatrix(J&, 24) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 25) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 26) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 27) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 28) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 29) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 30) = ""
       
       J& = J& + 1
       Label3.Caption = "Registros Procesados: " & J&
       Label3.Refresh

       .MoveNext
     Loop
     End With
     RS.Close
     Set RS = Nothing
     
'     GRIDH.COL = 1
'     GRIDH.Sort = 3
     'CARGA BAJO LAS MISMAS CONDICIONES LOS INSTITUOS QUE NO TUVIERON RESERVAS.
99   Screen.MousePointer = 1
     MUESTRABAP.GRIDH.Redraw = True

End Sub

Sub INSTITUCIONESSINRESERVA(DES, HAS)
     Screen.MousePointer = 11
     MUESTRABAP.GRIDH.Redraw = False
     Call ABRECONEXION
     'CARGO EL ECO DE LA INSTITUCION
     ANO$ = Year(DES)
'     Dim INSTITU$(), DOMI$(), CIUDAD$(), CPOS(), PROV$(), TELEFONO$(), MAIL$()
'     SQLX$ = "SELECT ID, Schools,[zip code] AS ICP,Town ,Address, PROVINCE, TELEPHONES, EMAIL"
'     SQLX$ = SQLX$ + " FROM INSTITUCIONES WITH(NOLOCK) ORDER BY ID "
'     Set RS0 = READSET(SQLX$)
'     With RS0
'     Do While Not .EOF
'        ID& = XVALO(!ID)
'        SCHOLS$ = SAFETEXT$(!Schools)
'        ReDim Preserve INSTITU$(ID&), DOMI$(ID&), CIUDAD$(ID&), CPOS(ID&), PROV$(ID&), TELEFONO$(ID&), MAIL$(ID&)
'        INSTITU$(ID&) = SCHOLS$
'        DOMI$(ID&) = SAFETEXT$(!Address)
'        CIUDAD$(ID&) = SAFETEXT$(!town)
'        CPOS(ID&) = SAFETEXT$(!ICP)
'        PROV$(ID&) = SAFETEXT$(!province)
'        TELEFONO$(ID&) = SAFETEXT$(!Telephones)
'        MAIL$(ID&) = SAFETEXT$(!EMAIL)
'        .MoveNext
'     Loop
'     End With
'     RS0.Close
'     Set RS0 = Nothing
'     'FIN ECO DE LA INSTITUCION
     SQLX$ = "Select ID, ADDRESS, TOWN, SCHOOLS, EMAIL, Telephones, Email ,LastUpdate,PROVINCE,[ZIP CODE] from INSTITUCIONES  WITH(NOLOCK)"
     SQLX$ = SQLX$ + "where ID >= 0 AND  not exists (select 1 from BOOKINGS WHERE BOOKINGS.IDInstitucion = INSTITUCIONES.ID"
     SQLX$ = SQLX$ + " AND fechaReserva  >= '" & DES & "'"
     SQLX$ = SQLX$ + " AND fechaReserva  <= '" & HAS & "'"
     SQLX$ = SQLX$ + ")"
     
     Set RS = READSET(SQLX$)
     J& = MUESTRABAP.GRIDH.Rows
     With RS
     Do While Not .EOF
       MUESTRABAP.GRIDH.Rows = J& + 1

       MUESTRABAP.GRIDH.TextMatrix(J&, 0) = ANO$
       MUESTRABAP.GRIDH.TextMatrix(J&, 1) = XVALO(!ID)
       MUESTRABAP.GRIDH.TextMatrix(J&, 2) = SAFETEXT$(!SCHOOLS)
       MUESTRABAP.GRIDH.TextMatrix(J&, 3) = SAFETEXT$(!Telephones)
       MUESTRABAP.GRIDH.TextMatrix(J&, 4) = SAFETEXT$(!Email)
       MUESTRABAP.GRIDH.TextMatrix(J&, 5) = SAFETEXT$(!Address)
       MUESTRABAP.GRIDH.TextMatrix(J&, 6) = SAFETEXT$(!TOWN)
       MUESTRABAP.GRIDH.TextMatrix(J&, 7) = SAFETEXT$(!PROVINCE)
       MUESTRABAP.GRIDH.TextMatrix(J&, 8) = SAFETEXT$(![ZIP CODE])
       
       MUESTRABAP.GRIDH.TextMatrix(J&, 9) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 10) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 11) = ""
       
       MUESTRABAP.GRIDH.TextMatrix(J&, 12) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 13) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 14) = ""
       
       MUESTRABAP.GRIDH.TextMatrix(J&, 15) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 16) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 17) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 18) = "NO RESERVÓ" 'SAFETEXT$(!PCP)
       
       MUESTRABAP.GRIDH.TextMatrix(J&, 19) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 20) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 21) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 22) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 23) = ""
        MUESTRABAP.GRIDH.TextMatrix(J&, 24) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 25) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 26) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 27) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 28) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 29) = ""
       MUESTRABAP.GRIDH.TextMatrix(J&, 30) = ""

       J& = J& + 1
       Label3.Caption = "Registros Procesados: " & J&
       Label3.Refresh
       .MoveNext
     Loop
     End With
     RS.Close
     Set RS = Nothing
     
'     GRIDH.COL = 1
'     GRIDH.Sort = 3
     'CARGA BAJO LAS MISMAS CONDICIONES LOS INSTITUOS QUE NO TUVIERON RESERVAS.
99   Screen.MousePointer = 1
     MUESTRABAP.GRIDH.Redraw = True

End Sub

Sub CARGA_ENCABEZADOHERICAL(MSF1 As MSHFlexGrid, LISTADECAMPOS$, FORM1 As Form, Optional FiltroCombo%, Optional NOREPIT%, Optional ANCHO_TOTAL, Optional ANCHO_COLUMNA_0%, Optional PINTAFILACLICK%)
    '
    
    Dim NOCAMPO$(), FORMX$()
    ReDim NOCAMPO$(1), FORMX$(1)
    '
    ANCHO_TOTAL = 0
    On Error Resume Next
    If FORM1.Label8.Count > 1 Then
      Erase FORM1.Label8
    End If
    On Error GoTo 0
    If FiltroCombo% > 0 Then
       FORM1.CMBFILTRO.AddItem ""
    End If
    NUPIC% = 0
    MODOSE$ = UCase$(SAFETEXT$(MODOSEL$))
    LSTX$ = TRIM$(LISTADECAMPOS$)
    CANCA% = 0
    ATOTA% = 1
    While LSTX$ <> ""
       PPXX% = InStr(LSTX$, ";")
       PPXX2% = InStr(LSTX$, ",")
       If PPXX2% > 0 And PPXX2% < PPXX% Then PPXX% = PPXX2%
       If PPXX% < 1 Then PPXX% = 1 + Len(LSTX$)
       '
       CANCA% = CANCA% + 1
       If CANCA% >= UBound(NOCAMPO$) Then
          ReDim Preserve NOCAMPO$(CANCA%), FORMX$(CANCA%)
       End If
       NOCX$ = Left$(LSTX$, PPXX% - 1)
       LSTX$ = Mid$(LSTX$, PPXX% + 1)
       '
       PPXX% = InStr(NOCX$, "/")
       If PPXX% < 1 Then PPXX% = 1 + Len(NOCX$)
       NOCAMPO$(CANCA%) = Left$(NOCX$, PPXX% - 1)
       FORMX$(CANCA%) = Mid$(NOCX$, PPXX% + 1)
       '
       'AGREGO EL LOS CAMPOS AL FILTRO
       If FiltroCombo% > 0 Then
          FORM1.CMBFILTRO.AddItem NOCAMPO$(CANCA%)
       End If

       If FORMX$(CANCA%) = "" Then FORMX$(CANCA%) = "A16"
       '
       ATOTA% = ATOTA% + XVALO(Mid$(FORMX$(CANCA%), 2)) + 2
       '
'       TECO$ = ""
'       PPXX% = InStr(FORMX$(CANCA%), "+")
'       If PPXX% > 0 Then
'          TECO$ = UCase$(TRIM$(Mid$(FORMX$(CANCA%), PPXX% + 1)))
'          PPXX% = InStr(TECO$, "/")
'          If PPXX% > 0 Then
'             FORMATX$ = TRIM$(Mid$(TECO$, PPXX% + 1))
'             ATOTA% = ATOTA% + Len(TMASCARA$(FORMATX$))
'          End If
'       End If
       
       If CANCA% = 1 Then ACOL1% = XVALO(Mid$(FORMX$(CANCA%), 2))
       '
    Wend

    MSF1.Cols = CANCA% + 1
    MSF1.Rows = 1
    For i% = 1 To CANCA%
       'PRIMER COLUMNA ASIGNA POR DEFECTO ANCHO = 2

       If i% = 0 Then
         If UCase$(FORM1.Name) = "SHOWREP07" Then GoTo 80

         If ANCHO_COLUMNA_0% > 0 Then
           ancho% = ANCHO_COLUMNA_0% 'SI  SE PASA EL ANCHO DE LA COLUMNA 0 COMO PARAMETRO TOMA ESTE
         Else
           ancho% = 3  'EN CASO CONTRARIO ASIGNA VALOR POR DEFECTO
         End If
         NUCOLU% = i%:   FORM1.TEPRUEBA = String$(ancho%, "A"): MSF1.ColWidth(NUCOLU%) = FORM1.TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
'         GoTo 50
       End If
       
30     LETRX$ = "A"
       ancho% = XVALO(Mid$(FORMX$(i%), 2))
       If XVALO(Mid$(FORMX$(i%), 2)) > 0 Then
          ancho% = ancho% + 2
       ElseIf LETRX$ = "Z" Then
            
       Else
          ancho% = 0 'COLUMNAS OCULTAS PASADAS CON F0.0
       End If
       LETR$ = UCase$(Left$(FORMX$(i%), 1))
       ALINEA% = 1
       If LETR$ = "I" Or LETR$ = "G" Or LETR$ = "F" Then
          ALINEA% = 6
       ElseIf LETR$ = "Z" Then
          ALINEA% = 3
       End If
       '
       If ancho% > 0 Then
          NUCOLU% = i%:   FORM1.TEPRUEBA = String$(ancho%, "A"): MSF1.ColWidth(NUCOLU%) = FORM1.TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
       Else ' PARA QUE NO SE VEAN LAS COLUMNAS OCULTAS
          NUCOLU% = i%:   FORM1.TEPRUEBA = String$(ancho%, "A"): MSF1.ColWidth(NUCOLU%) = 0: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
       End If
       MSF1.COL = NUCOLU% - 1
       MSF1.CellFontBold = True
       
       ANCHO_TOTAL = ANCHO_TOTAL + FORM1.TEPRUEBA.Width
       
       'A PARTIR DE AQUI CODIGO PARA CARGAR ENCABEZADO DE GENREP07

50     If UCase$(FORM1.Name) = "SHOWREP07" Then
          FORM1.Label5.FontName = FORM1.GRID.CellFontName
          FORM1.Label5.FontSize = FORM1.GRID.CellFontSize
          FORM1.Label5.FontBold = FORM1.GRID.CellFontBold
          ' Label5.Caption = TEXTBIN$("1.234567", FORMAECO$(I&)) + "*"
          FORM1.Label5.Caption = String$(ancho%, "A")    ' PROVISORIAMENTE A VER SI FUNCIONA
          FORM1.GRID.COL = i%
          FORM1.GRID.ColWidth(JI& - 1) = SHOWREP07.Label5.Width
25        If i% >= FORM1.Label8.Count Then
             NUPIC% = FORM1.Label8.Count
             Load FORM1.Label8(NUPIC%)
             GoTo 25
          End If
          FORM1.Label8(i%) = Str$(Len(FORM1.Label5) - 1)
          If i% = 0 Then FORM1.Label8(i%) = 0
       End If
80  Next i%
    '
    If NUPIC% = i% - 1 Then
      If UCase$(FORM1.Name) = "SHOWREP07" Then Load FORM1.Label8(NUPIC% + 1): Load FORM1.Label8(NUPIC% + 2)
    End If
    If NOREPIT% > 0 Then Call Duplicados(FORM1.CMBFILTRO)
    
End Sub


