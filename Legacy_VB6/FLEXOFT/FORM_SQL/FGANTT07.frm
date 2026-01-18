VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FGANTT07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0D0D0&
   Caption         =   "Secuencia de Carga por Equipo Productivo"
   ClientHeight    =   9375
   ClientLeft      =   435
   ClientTop       =   3780
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   13485
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C7D9CC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7305
      Left            =   -400
      ScaleHeight     =   7305
      ScaleWidth      =   2790
      TabIndex        =   24
      Top             =   0
      Width           =   2790
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0D0D0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2175
      ScaleWidth      =   13605
      TabIndex        =   4
      Top             =   7530
      Width           =   13600
      Begin VB.ListBox List1 
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
         Height          =   285
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   1905
      End
      Begin MSFlexGridLib.MSFlexGrid dgvFabricacion 
         Height          =   1335
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2355
         _Version        =   393216
         Rows            =   0
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
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
      Begin VB.ListBox List2 
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
         Height          =   240
         IntegralHeight  =   0   'False
         Left            =   2160
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0D0D0&
         Caption         =   "Sector Productivo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   7200
         TabIndex        =   18
         Top             =   960
         Width           =   2055
         Begin VB.ComboBox Combo1 
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
            Left            =   75
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   300
            Width           =   1905
         End
      End
      Begin VB.ListBox List3 
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
         Height          =   240
         IntegralHeight  =   0   'False
         Left            =   4680
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   7200
         ScaleHeight     =   540
         ScaleWidth      =   3135
         TabIndex        =   12
         Top             =   360
         Width           =   3165
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Mensual"
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
            Height          =   225
            Index           =   3
            Left            =   1750
            TabIndex        =   16
            Top             =   280
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Quincenal"
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
            Height          =   225
            Index           =   2
            Left            =   1600
            TabIndex        =   15
            Top             =   50
            Width           =   1395
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Semanal"
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
            Height          =   225
            Index           =   1
            Left            =   250
            TabIndex        =   14
            Top             =   280
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Diario"
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
            Height          =   225
            Index           =   0
            Left            =   100
            TabIndex        =   13
            Top             =   50
            Width           =   1035
         End
      End
      Begin VB.CommandButton Command8 
         Height          =   590
         Left            =   11160
         Picture         =   "FGANTT07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Consultas por Pantalla"
         Top             =   1080
         Width           =   550
      End
      Begin VB.CommandButton Command7 
         Height          =   590
         Left            =   11160
         Picture         =   "FGANTT07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Listados por Impresora"
         Top             =   360
         Width           =   550
      End
      Begin VB.CommandButton Command6 
         Height          =   590
         Left            =   10560
         Picture         =   "FGANTT07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Acceso Directo a Aplicaciones Complementarias"
         Top             =   1080
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.CommandButton Command3 
         Height          =   590
         Left            =   11760
         Picture         =   "FGANTT07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   1080
         Width           =   550
      End
      Begin VB.CommandButton Command5 
         Height          =   590
         Left            =   10560
         Picture         =   "FGANTT07.frx":0BA0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Impresión Gráfico de Carga"
         Top             =   360
         Width           =   550
      End
      Begin VB.CommandButton Command2 
         Height          =   590
         Left            =   11760
         Picture         =   "FGANTT07.frx":120A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Re-Cálcule General de Carga de Máquinas"
         Top             =   360
         Width           =   550
      End
      Begin MSFlexGridLib.MSFlexGrid dgvReferencia 
         Height          =   1335
         Left            =   2160
         TabIndex        =   30
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   2355
         _Version        =   393216
         Rows            =   0
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
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
      Begin MSFlexGridLib.MSFlexGrid dgvArticulos 
         Height          =   1335
         Left            =   4680
         TabIndex        =   31
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   2355
         _Version        =   393216
         Rows            =   0
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
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
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Artículos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   28
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia | O/Compra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "O/Fabricación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Período"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   22
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7550
      Left            =   1560
      ScaleHeight     =   7545
      ScaleWidth      =   11955
      TabIndex        =   0
      Top             =   0
      Width           =   11950
      Begin VB.PictureBox Picture999 
         Height          =   435
         Index           =   1
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   1215
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         Left            =   840
         TabIndex        =   3
         Top             =   7300
         Width           =   10800
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   6730
         Left            =   11655
         TabIndex        =   2
         Top             =   540
         Width           =   225
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   4020
         OleObjectBlob   =   "FGANTT07.frx":1514
         Top             =   2100
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
         Height          =   435
         Index           =   0
         Left            =   4080
         OleObjectBlob   =   "FGANTT07.frx":1748
         TabIndex        =   21
         Top             =   420
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0D0D0&
         Height          =   255
         Left            =   -650
         TabIndex        =   25
         Top             =   7310
         Width           =   1455
      End
      Begin VB.Line Line4 
         X1              =   11640
         X2              =   11640
         Y1              =   8085
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   0
         X2              =   12000
         Y1              =   7290
         Y2              =   7290
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   0
         X2              =   12000
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   840
         X2              =   840
         Y1              =   0
         Y2              =   8295
      End
   End
   Begin VB.Label TEPRUEBA2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   32
      Top             =   7320
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label TEPRUEBA 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuRecalculo 
         Caption         =   "Re-Cálculo de Carga de Máquinas"
      End
      Begin VB.Menu S 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnusimulacion 
      Caption         =   "Simulación"
      Begin VB.Menu mnuGenerarSimulacion 
         Caption         =   "Generar Simulación de Fábrica"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
      Begin VB.Menu mnuopcionesdevista 
         Caption         =   "Opciones de Vista"
         Begin VB.Menu mnuordenfabricacionreales 
            Caption         =   "Vista de Ord. Fabric. Reales"
         End
         Begin VB.Menu mnuincluyesimulacion 
            Caption         =   "Vista Incluyendo Simulación"
         End
      End
   End
   Begin VB.Menu mnureportes 
      Caption         =   "Reportes"
      Begin VB.Menu mnulistadoimpresora 
         Caption         =   "Listados por impresora"
      End
      Begin VB.Menu mnuconsultaxpantalla 
         Caption         =   "Consulta por Pantalla"
      End
      Begin VB.Menu mnuimpresiongrafico 
         Caption         =   "Impresión de Gráfico de Carga"
      End
   End
   Begin VB.Menu mnuaccesodirecto 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuArchivoMaestro 
         Caption         =   "Archivos Maestros"
      End
      Begin VB.Menu mnuaccesoaplicacioncomplementaria 
         Caption         =   "Acceso Directo a Aplicaciones Complementarias"
      End
   End
   Begin VB.Menu mnuconfig 
      Caption         =   "Configuración"
      Begin VB.Menu mnuconfiguracion 
         Caption         =   "Configuraciones de Funcionamiento"
      End
   End
End
Attribute VB_Name = "FGANTT07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XMAXI
Dim MATIE$(1024)
Dim CAOFA%, CACODACT%
Dim TBAJACLI
Dim NODIBUJA%
Public CONSIMULACION%
Public CARGAFORMULARIO%
Public AltoFilaMaquinas%
Public AltoBarra%
Private Sub Combo1_Click()
    'CADA VEZ QUE SE CAMBIA EL COMBO SE ACTUALIZA LA IMAGEN
    For i% = 1 To Option1.UBound
        If Option1(i%).Value = True Then
           Call Option1_Click(i%)
        End If
    Next i%
End Sub




Sub Command2_Click()
    '
    Command2.Enabled = False
    '
    Call GENORDEFDAT(1)     ' solo pendientes
    COMPLECALCU% = 0
    CALCAR07.Show 1
    '
    If COMPLECALCU% = 1 Then
      Call CIERRARCH("*.*")
      Call LIBARCH("*.*")
      Call FILESORT("OPERFAP", "", 20, 21, "ASC")
      Call COMUNI("Proceso Completo")
      If EXISTE%(App.EXEName + ".EXE") > 0 Then
        Call BLOQEXE("*")
        Call CONECRUN(App.EXEName, "Programación de Carga de Máquinas")
      End If
      Call FINAL("")
    End If
    '
    Command2.Enabled = True
    '
End Sub

Sub Command3_Click()
   Command3.Enabled = False
   If DERACCE%("STPCRP", "Configuración de Funcionamiento Carga de Maquinas") = 2 Then
     CFGCRP07.Show 1
   End If
   Command3.Enabled = True
End Sub


Private Sub Command5_Click()
   Command5.Enabled = False
   ANCHAN = Width
   Call Form_Resize
   Width = Printer.ScaleWidth
   PRINTFORM
   Width = ANCHAN
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   FIN& = ULTREG&("OPERFAT")
   
   For II& = 1 To FIN&
    REG = REGLEIDO$("OPERFAT", II&)
    NORFA$ = Mid$(REG, 1, 12)
    OPERA% = CVI(Mid$(REG, 15, 2))
    NIVEL& = CVS(Mid$(REG, 199, 4))
    
    FECHAI% = CVI(Mid$(REG, 135, 2))
    HORAI$ = CVI(Mid$(REG, 137, 2))
    FECHAF% = CVI(Mid$(REG, 139, 2))
    HORAF$ = CStr(CVI(Mid$(REG, 141, 2)))
    
    Debug.Print NORFA$ & "-" & CStr(OPERA%) & "-" & CStr(NIVEL&) & "-" & FECHATEX$(FECHAI%) & "-" & HORAI$ & "--" & FECHATEX$(FECHAF%) & "-" & HORAF$
    
   Next II&
   
   Exit Sub
   ADIRCA07.Show 1
   '
   If ATENTI% > 0 Then ' SE INFORMARON OPERACIONES
     If COMALTER%("Debe Ahora Re-Calcular la Carga de Máquinas\\Si, Re-calcular\No, Continuar") = 1 Then
       '
       COMPLECALCU% = 0
       On Error Resume Next
       CALCUCAR.Show 1
       If Err Then
         CALCAR07.Show 1
       End If
       '
       If COMPLECALCU% = 1 Then
         Call CIERRARCH("*.*")
         Call LIBARCH("*.*")
         Call FILESORT("OPERFAP", "", 20, 21, "ASC")
         Call COMUNI("Proceso Completo")
         If EXISTE%(App.EXEName + ".EXE") > 0 Then
           Call BLOQEXE("*")
           Call CONECRUN(App.EXEName, "Programación de Carga de Máquinas")
           Call FINAL("")
         End If
       End If
       '
     End If
   End If
   '
   Command6.Enabled = True
End Sub

Sub Command7_Click()
   Command7.Enabled = False
   LCARMA07.Show 1
   Command7.Enabled = True
End Sub

Sub Command8_Click()
   Command8.Enabled = False
   CONCAR07.Show 1
   Command8.Enabled = True
End Sub

Private Sub dgvArticulos_Click()

    Screen.MousePointer = 11

    ' *** GUARDA LA POSICION SELECCIONADA DEL GRID
    IdFila& = dgvArticulos.Row
    If IdFila& < 0 Then Exit Sub
    
    ' *** LIMPIA EL GRID DE FABRICACION Y EL DE REFERENCIAS
    For dgvFab_fila& = 0 To dgvFabricacion.Rows - 1
        dgvFabricacion.Row = dgvFab_fila&
        dgvFabricacion.CellForeColor = &H0&
        dgvFabricacion.CellBackColor = &H80000005
    Next dgvFab_fila&
    For dgvRef_fila& = 0 To dgvReferencia.Rows - 1
        dgvReferencia.Row = dgvRef_fila&
        dgvReferencia.CellForeColor = &H0&
        dgvReferencia.CellBackColor = &H80000005
    Next dgvRef_fila&
    For dgvArt_fila& = 0 To dgvArticulos.Rows - 1
        dgvArticulos.Row = dgvArt_fila&
        dgvArticulos.CellForeColor = &H0&
        dgvArticulos.CellBackColor = &H80000005
    Next dgvArt_fila&
    dgvFabricacion.Refresh
    dgvReferencia.Refresh
    
    ' *** SE POSICIONA EN LA FILA ELEGIDA PARA CAMBIARLE EL COLOR
    dgvArticulos.Row = IdFila&
    dgvArticulos.CellForeColor = &H80000005
    dgvArticulos.CellBackColor = &H808080
    dgvArticulos.Refresh

    ' *** UTILIZA EL LIST PARA DESPLEGAR LA LISTA DE OPERACIONES
    List3.ListIndex = IdFila&
    Call List3_Click
    
    Screen.MousePointer = 1
    
End Sub

Private Sub dgvFabricacion_Click()

'    CONTROCLICK% = (-1)
'    IdFila& = dgvFabricacion.Row
'    List1.ListIndex = IdFila& + 1
'
'    Call PREGANTT
'    NORFA$ = TRIM$(Left$(List1.TEXT, 24))
'    Call MUEDATORF(NORFA$)
'    CONTROCLICK% = 0
    
End Sub

Private Sub dgvFabricacion_DblClick()

    Screen.MousePointer = 11
    ' *** GUARDA LA POSICION SELECCIONADA DEL GRID
    IdFila& = dgvFabricacion.Row
    
    If IdFila& < 0 Then Exit Sub
    
    ' *** LIMPIA EL GRID
    For dgvFab_fila& = 0 To dgvFabricacion.Rows - 1
        dgvFabricacion.Row = dgvFab_fila&
        dgvFabricacion.CellForeColor = &H0&
        dgvFabricacion.CellBackColor = &H80000005
    Next dgvFab_fila&
    
    ' *** SE POSICIONA EN LA FILA ELEGIDA PARA CAMBIARLE EL COLOR
    dgvFabricacion.Row = IdFila&
    dgvFabricacion.CellForeColor = &H80000005
    dgvFabricacion.CellBackColor = &H808080
    dgvFabricacion.Refresh
    
    ' *** UTILIZA EL LIST PARA DESPLEGAR LA LISTA DE OPERACIONES
    List1.ListIndex = IdFila&
    Call List1_Click
    
    Screen.MousePointer = 1
    
End Sub

Private Sub dgvReferencia_Click()

    Screen.MousePointer = 11
    ' *** GUARDA LA POSICION SELECCIONADA DEL GRID
    IdFila& = dgvReferencia.Row
    
    If IdFila& < 0 Then Exit Sub
    
    ' *** LIMPIA EL GRID DE FABRICACION Y EL DE REFERENCIAS
    For dgvFab_fila& = 0 To dgvFabricacion.Rows - 1
        dgvFabricacion.Row = dgvFab_fila&
        dgvFabricacion.CellForeColor = &H0&
        dgvFabricacion.CellBackColor = &H80000005
    Next dgvFab_fila&
    For dgvArt_fila& = 0 To dgvArticulos.Rows - 1
        dgvArticulos.Row = dgvArt_fila&
        dgvArticulos.CellForeColor = &H0&
        dgvArticulos.CellBackColor = &H80000005
    Next dgvArt_fila&
    For dgvRef_fila& = 0 To dgvReferencia.Rows - 1
        dgvReferencia.Row = dgvRef_fila&
        dgvReferencia.CellForeColor = &H0&
        dgvReferencia.CellBackColor = &H80000005
    Next dgvRef_fila&
    dgvFabricacion.Refresh
    dgvArticulos.Refresh
    
    ' *** SE POSICIONA EN LA FILA ELEGIDA PARA CAMBIARLE EL COLOR
    dgvReferencia.Row = IdFila&
    dgvReferencia.CellForeColor = &H80000005
    dgvReferencia.CellBackColor = &H808080
    dgvReferencia.Refresh
    
    ' *** UTILIZA EL LIST PARA DESPLEGAR LA LISTA DE OPERACIONES
    List2.ListIndex = IdFila&
    Call List2_Click
    
    Screen.MousePointer = 1

End Sub

Private Sub Form_Load()
   '
   CARGAFORMULARIO% = 1
   
   AltoFilaMaquinas% = TEPRUEBA.Height
   AltoBarra% = 145
  
   VERANTER$ = "CARGMAQ"
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
''   EPAC$ = TRIM$(EMPREAC$)
''   If EPAC$ <> "" Then
''     '
''10   Caption = Caption + "  -  " + EPAC$
''     'APPX$ = App.EXEName
''     'If EXISTE%(APPX$ + ".KKK") > 0 Then GoTo 10
''     'FMDX = FECHMODU%
''     'BB1 = HOY(HOS$)
''     'If DIENTRE%(FMDX, HOY(HOS$)) > 100 Then
''     '  Call SAVEFILE(APPX$ + ".KKK", String$(128, 0))
''     'End If
''     '
''   End If
10  EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
   Call GRATITFOR(Caption)
   '
   'AL MENU SE LE AGREGA EL DATO DE LA SIMULACION ACTIVA.
   mnuordenfabricacionreales.Checked = True
   mnuincluyesimulacion.Checked = False
   UlitmoCalculoSimulacion& = XVALO(CONTROL("CARMAQUI", "ULTISIMU"))
   If UlitmoCalculoSimulacion& > 0 Then mnuincluyesimulacion.Caption = mnuincluyesimulacion.Caption & " (" & UlitmoCalculoSimulacion& & ")"
   '
   If EXISTE_TABLA("SIMU_ORDEFABR") > 0 Then
     If CantRegistros("SIMU_ORDEFABR", , "NOMESS") > 0 Then
       OPX% = COMALTER%("Opciones de Trabajo\\Cálculo Normal\Cálculo Considerando Simulación")
       If OPX% > 1 Then
         CONSIMULACION% = 1
         mnuordenfabricacionreales.Checked = False
         mnuincluyesimulacion.Checked = True
       End If
     End If
   End If
   FEIPER = HOY(HO$)
   Screen.MousePointer = 11
   For U& = 1 To ULTREG&("OPERFAP")
     XX$ = REGLEIDO$("OPERFAP", U&)
     FECOMI = CVI(Mid$(XX$, 135, 2))
     If FECOMI >= 1 Then
       If FECOMI < FEIPER Then
         FEIPER = FECOMI
       End If
     End If
   Next U&
   Screen.MousePointer = 1
   'While DIASEM%(FEIPER) <> 1
   '   FEIPER = DIANTE(FEIPER)
   'Wend
   '
   Show
   Refresh
   NODIBUJA% = 1
   HScroll1.Min = FEIPER - 32 * 13
   HScroll1.Max = FEIPER + 32 * 13
   HScroll1.Value = FEIPER
   NODIBUJA% = 0
   
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me, Picture2)
   '
   Call DIBULINVER
   Call CALCUFECHA
   Call DIBUFECHA
   Call CARLIORF
   Call DIBUEQUIP
   Call PREGANTT
'
   'If List1.ListIndex >= 0 Then
   '   NORFA$ = TRIM$(Left$(List1.TEXT, 24))
   '   Call MUEDATORF(NORFA$)
   'End If
   '
   ' CARGA LISTA DE EQUIPOS
   Combo1.Clear
   Combo1.AddItem "     "
   For U& = 1 To ULTREG&("SECPROD")
        X$ = REGLEIDO$("SECPROD", U&)
        SECTOR$ = FORMATNUM$(CVS(Mid$(X$, 1, 4)), "F7.3")
        DESCRISEC$ = Mid$(X$, 5, 32)
        TTXX$ = SECTOR$ & " " & DESCRISEC$
        Combo1.AddItem TTXX$
   Next U&
   '
   Call CONAPLI(APLINVO$, TERMINA%) ' Busca la aplicacion a mostrar
   If TERMINA% > 0 Then Call FINAL("")
   '
   YACARGANTT% = 1
   '
   CARGAFORMULARIO% = 0
   '
End Sub

Private Sub Form_Resize()

    If CARGAFORMULARIO% = 0 Then
        AnchoFormulario% = Me.Width - 100
        
        INCREY = TEPRUEBA.Height * 2
        
        Picture1.Width = AnchoFormulario% - Picture1.Left
        Line1.X2 = Picture1.Width - Line1.X1
        Line4.X1 = Picture1.Width - VScroll1.Width - 90: Line4.X2 = Line4.X1
        VScroll1.Left = Line4.X1 + 90
        HScroll1.Width = Picture1.Width - HScroll1.Left - VScroll1.Width
        
        Picture2.Top = Me.Height - Picture2.Height - 400
        Picture1.Height = Picture2.Top
        Picture4.Height = Picture1.Height
        HScroll1.Top = Picture1.Height - HScroll1.Height
        Line3.Y1 = HScroll1.Top - 50: Line3.Y2 = Line3.Y1: Line3.X2 = Line1.X2
        Line2.Y2 = Line3.Y1
        Line4.Y1 = Line2.Y2
        VScroll1.Height = Line1.Y1 + (Line3.Y1 - Line1.Y1) - 570
        
        If AltoFilaMaquinas% = TEPRUEBA2.Height Then
            TEPRUEBA2.Height = ((Line3.Y1 - Line1.Y1) / 30) / 1.2
            AltoBarra% = TEPRUEBA2.Height * 0.6
        Else
            TEPRUEBA2.Height = AltoFilaMaquinas%
            AltoBarra% = 145
        End If
        
        Picture1.Cls
        Picture4.Cls
        
        Call DIBULINVER
        Call CALCUFECHA
        Call DIBUFECHA
        Call DIBUEQUIP
        Call PREGANTT
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub


Private Sub HScroll1_Change()
   If YACARGANTT% = 1 Then
      FEIPERANT = FEIPER
      FEIPER = HScroll1.Value
      If FEIPER < FEIPERANT Then
           SENTI% = -1
         Else
           SENTI% = 1
      End If
10    FEX = FEIPER
      FEXU$ = FECHATEX$(FEX)
      If Val(Left$(FEXU$, 2)) < 1 Or Val(Mid$(FEXU$, 4, 2)) < 1 Or FEX <> FEIPER Then
         If SENTI% > 0 Then
             FEIPER = DIPOST(FEIPER)
           Else
             FEIPER = DIANTE(FEIPER)
         End If
         GoTo 10
      End If
   End If
   '
   If FEIPER >= HScroll1.Min Then
     If FEIPER <= HScroll1.Max Then
       HScroll1.Value = FEIPER
     End If
   End If
   Picture1.Cls
   Picture4.Cls
   Call DIBULINVER
   Call CALCUFECHA
   Call DIBUFECHA
   Call DIBUEQUIP
   Call PREGANTT
End Sub

Private Sub List1_Click()
   If CONTROCLICK% = 0 Then
        ' *** LIMPIA LAS GRILLAS DE REFERENCIA Y ARTICULOS
        For dgvArt_fila& = 0 To dgvArticulos.Rows - 1
            dgvArticulos.Row = dgvArt_fila&
            dgvArticulos.CellForeColor = &H0&
            dgvArticulos.CellBackColor = &H80000005
        Next dgvArt_fila&
        For dgvRef_fila& = 0 To dgvReferencia.Rows - 1
            dgvReferencia.Row = dgvRef_fila&
            dgvReferencia.CellForeColor = &H0&
            dgvReferencia.CellBackColor = &H80000005
        Next dgvRef_fila&
    
        Call PREGANTT
        NORFA$ = TRIM$(Left$(List1.TEXT, 24))
        Call MUEDATORF(NORFA$)
   End If
End Sub

Private Sub List2_Click()

   REFEOF$ = TRIM$(List2.List(List2.ListIndex))
   ULTSELE& = 0
   CONTROCLICK% = (-1)
   For U& = 1 To CAOFA%
        If OFAREFE$(U&) = REFEOF$ Then
            For i& = 0 To List1.ListCount - 1
                NORFA$ = TRIM$(Mid$(List1.List(i&), 1, 12))
                If NORFB(U&) = NORFA$ Then
                    List1.Selected(i&) = True
                    IdFila& = i&
                    dgvFabricacion.Row = IdFila&
                    dgvFabricacion.CellForeColor = &H80000008
                    dgvFabricacion.CellBackColor = &HE2D3C0
                    ULTSELE& = U&
                End If
            Next i&
        End If
   Next U&
   CONTROCLICK% = 0
   List1.Refresh
   Call PREGANTT
   CONTROCLICK% = 0

End Sub

Private Sub List3_Click()
   CIIX% = CODINT%(List3.TEXT)
   CONTROCLICK% = (-1)
   OFACT& = 0
   
   For U& = 1 To CAOFA%
        If CIOFA%(U&) = CIIX% Then
            For i& = 0 To List1.ListCount - 1
                NORFA$ = TRIM$(Mid$(List1.List(i&), 1, 12))
                If NORFB(U&) = NORFA$ Then
                    List1.Selected(i&) = True
                    IdFila& = i&
                    dgvFabricacion.Row = IdFila&
                    dgvFabricacion.CellForeColor = &H80000008
                    dgvFabricacion.CellBackColor = &HE2D3C0
                    OFACT& = U&
                End If
            Next i&
        End If
   Next U&
   
   CONTROCLICK% = 0
   List1.Refresh
   Call PREGANTT

End Sub

Private Sub List3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
     Call SELECHO("!CODIORFA")
     CDXX$ = TRIM$(VALACT1$("!CODIORFA"))
     If CDXX$ <> "" Then
       For UK& = 1 To List3.ListCount
         If TRIM$(List3.List(UK& - 1)) = CDXX$ Then
           List3.ListIndex = (UK& - 1)
           Exit For
         End If
       Next UK&
     End If
   End If
End Sub

Private Sub mnuaccesoaplicacioncomplementaria_Click()
   Call Command6_Click
End Sub

Private Sub mnuArchivoMaestro_Click()
   'Call Command4_Click
End Sub

Private Sub mnuconfiguracion_Click()
    Call Command3_Click
End Sub

Private Sub mnuconsultaxpantalla_Click()
   Call Command8_Click
End Sub

Private Sub mnuGenerarSimulacion_Click()
    SIMUCARM.Show 1
End Sub

Private Sub mnuimpresiongrafico_Click()
   Call Command5_Click
End Sub

Private Sub mnuincluyesimulacion_Click()
   CONSIMULACION% = 1
   Screen.MousePointer = 11
   '
   mnuordenfabricacionreales.Checked = False
   mnuincluyesimulacion.Checked = True

   Call DIBULINVER
   Call CALCUFECHA
   Call DIBUFECHA
   Call CARLIORF
   Call DIBUEQUIP
   Call PREGANTT
   Combo1.Clear
   Combo1.AddItem "     "
   For U& = 1 To ULTREG&("SECPROD")
        X$ = REGLEIDO$("SECPROD", U&)
        SECTOR$ = FORMATNUM$(CVS(Mid$(X$, 1, 4)), "F7.3")
        DESCRISEC$ = Mid$(X$, 5, 32)
        TTXX$ = SECTOR$ & " " & DESCRISEC$
        Combo1.AddItem TTXX$
   Next U&

End Sub


Private Sub mnulistadoimpresora_Click()
   Call Command7_Click
End Sub

Private Sub mnuordenfabricacionreales_Click()
   CONSIMULACION% = 0
   Screen.MousePointer = 11
   '
   mnuordenfabricacionreales.Checked = True
   mnuincluyesimulacion.Checked = False
   Call DIBULINVER
   Call CALCUFECHA
   Call DIBUFECHA
   Call CARLIORF
   Call DIBUEQUIP
   Call PREGANTT
   Combo1.Clear
   Combo1.AddItem "     "
   For U& = 1 To ULTREG&("SECPROD")
        X$ = REGLEIDO$("SECPROD", U&)
        SECTOR$ = FORMATNUM$(CVS(Mid$(X$, 1, 4)), "F7.3")
        DESCRISEC$ = Mid$(X$, 5, 32)
        TTXX$ = SECTOR$ & " " & DESCRISEC$
        Combo1.AddItem TTXX$
   Next U&
   '

End Sub

Private Sub mnuRecalculo_Click()
  Call Command2_Click
End Sub

Private Sub mnusalir_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")

End Sub

Private Sub Option1_Click(Index As Integer)
   Picture1.Cls
   Call DIBULINVER
   Call CALCUFECHA
   Call DIBUFECHA
   Call DIBUEQUIP
   Call PREGANTT
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   XMUOFA = MUORFACT.Left
   YMUOFA = MUORFACT.Top
   MUORFACT.Hide
   DoEvents
   '
   If X < Line2.X1 Then
     For UI% = PRIEQVIS% To CAEQ%
       If Y >= YEQU(UI%) Then
         If Y <= YEQU(UI%) + 250 Then
           COEQUIPX$ = CODEQ$(UI%)
           Picture1.ToolTipText = TRIM$(ECOARCH$("PADMAQ", COEQUIPX$))
           TBAJACLI = Timer
         End If
       End If
     Next UI%
     Exit Sub
   End If
   '
   For UI% = 1 To cabarr%
     If X >= XINI1(UI%) Then
       If X <= XFIN1(UI%) Then
         If Y >= YINI1(UI%) Then
           If Y <= YINI1(UI%) + 145 Then
             NORFAX$ = NUOFA$(UI%)
             For U& = 1 To List1.ListCount
               V& = Val(Mid$(List1.List(U& - 1), 25))
               If NORFB$(V&) = NORFAX$ Then
                 CONTROCLICK% = -1
                 List1.ListIndex = U& - 1
                 CONTROCLICK% = 0
                 OFACT& = U&
                 Call PREGANTT
                 Call MUEDATORF(NORFAX$)
                 Exit Sub
               End If
             Next U&
           End If
         End If
       End If
     End If
   Next UI%
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Timer > TBAJACLI + 1 Then
      Picture1.ToolTipText = ""
   End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Timer > TBAJACLI + 1 Then
      Picture1.ToolTipText = ""
   End If
End Sub

Private Sub VScroll1_Change()
   PRIEQVIS% = VScroll1.Value + 1
   Picture1.Cls
   Picture4.Cls
   Call DIBULINVER
   Call CALCUFECHA
   Call DIBUFECHA
   Call DIBUEQUIP
   Call PREGANTT
End Sub

Sub DIBULINVER()
   CAPERIO% = 7
   CALINHOR% = 12: INTERLIN% = 3
   If Option1(0) = True Then
        CAPERIO% = 2
        CALINHOR% = 24: INTERLIN% = 4
      ElseIf Option1(2) = True Then
        CAPERIO% = 15
        CALINHOR% = 6: INTERLIN% = 3
      ElseIf Option1(3) = True Then
        CAPERIO% = 31
        CALINHOR% = 2: INTERLIN% = 2
   End If
   Call CALCUFECHA
   APERIO = (Picture1.Width - Line2.X1 - VScroll1.Width - 80) / (CAPERIO%)
   DrawWidth = 2
   X1 = Line2.X1
   Y2 = Picture1.Height
   Y3 = Line1.Y1
   For II% = 1 To CAPERIO% + 1
      DSXX% = DIASEM%(DESPERIO(II%))
      If II% <= CAPERIO% Then
        If DSXX% > 5 Or DSXX% < 1 Then
          Picture1.Line (X1 + 20, YH1)-(X1 + APERIO, Y2), RGB(160, 160, 160), BF
        End If
      End If
      XINI(II%) = X1
      X2 = X1
      X1 = X1 + APERIO
      If II% = CAPERIO% Then XMAXI = X1 - 20
      Picture1.Line (X1, 0)-(X1, Y2), QBColor(0)
      '
      If II% <= CAPERIO% Then
        For JJ% = 1 To CALINHOR% - 1
          X3 = X2 + JJ% * APERIO / CALINHOR%
          COLOLIN = RGB(200, 200, 200)
          If JJ% Mod INTERLIN% = 0 Then COLOLIN = RGB(140, 140, 140)
          Picture1.Line (X3, Y3)-(X3, Y2), COLOLIN
        Next JJ%
      End If
   Next II%
   Picture1.Refresh
End Sub
'
Sub DIBUFECHA()
   '
   Picture1.Font = TEPRUEBA.Font
   TEPRUEBA.FontBold = True
   If Option1(2).Value = True Then
      TEPRUEBA.FontBold = False
   End If
   Picture1.FontSize = TEPRUEBA.FontSize
   Picture1.FontBold = TEPRUEBA.FontBold
   FORMAFE% = 0: If CAPERIO% > 15 Then FORMAFE% = 1
   For II% = 1 To CAPERIO%
      TTXX$ = FECHATEX$(DESPERIO(II%))
      If CAPERIO% > 7 Then
         TTXX$ = Left$(TTXX$, 5)
      End If
      If CAPERIO% > 15 Then
         TTYY$ = Mid$(TTXX$, 4, 2)
         TTXX$ = Left$(TTXX$, 2)
         If Left$(TTYY$, 1) = "0" Then TTYY$ = Mid$(TTYY$, 2)
      End If
      TEPRUEBA.Caption = TTXX$
      Picture1.CurrentX = Line2.X1 + APERIO * II% - APERIO / 2 - TEPRUEBA.Width / 2
      Picture1.CurrentY = 120 - 60 * FORMAFE%
      Picture1.Print TTXX$
      If FORMAFE% = 1 Then
         TEPRUEBA.Caption = "---"
         Picture1.CurrentX = Line2.X1 + APERIO * II% - APERIO / 2 - TEPRUEBA.Width / 2
         Picture1.CurrentY = 165
         Picture1.Print "---"
         TEPRUEBA.Caption = TTYY$
         Picture1.CurrentX = Line2.X1 + APERIO * II% - APERIO / 2 - TEPRUEBA.Width / 2
         Picture1.CurrentY = 300
         Picture1.Print TTYY$
        Else
         TTYY$ = DSEM$(DESPERIO(II%))
         If CAPERIO% > 7 Then TTYY$ = Left$(TTYY$, 2)
         TTYY$ = Left$(TTYY$, 1) + LCase$(Mid$(TTYY$, 2))
         TTYY$ = "(" + TTYY$ + ")"
         TEPRUEBA.Caption = TTYY$
         Picture1.CurrentX = Line2.X1 + APERIO * II% - APERIO / 2 - TEPRUEBA.Width / 2
         Picture1.CurrentY = 300
         Picture1.Print TTYY$
      End If
   Next II%
   TEPRUEBA.FontBold = True
End Sub
'
Sub DIBUEQUIP()
   Static MAXSCROLL%
   '
   Picture1.Font = TEPRUEBA.Font
   Picture1.FontSize = TEPRUEBA.FontSize
   Picture1.FontBold = TEPRUEBA.FontBold
   INCREY = TEPRUEBA2.Height * 1.2
   Picture1.CurrentY = Line1.Y1 - INCREY / 2
   '
   Picture4.Font = Picture1.Font
   Picture4.FontSize = Picture1.FontSize
   Picture4.FontBold = Picture1.FontBold
   Picture4.CurrentY = Picture1.CurrentY
   '
   SECTOR = 0
   If Combo1.ListIndex > -1 Then SECTOR = XVALO(Mid$(Combo1.List(Combo1.ListIndex), 1, 7))
   '
   If PRIEQVIS% < 1 Then PRIEQVIS% = 1
   For U& = PRIEQVIS% To CAEQ%
      Picture1.CurrentY = Picture1.CurrentY + INCREY
      Picture4.CurrentY = Picture1.CurrentY
      YINI(U&) = Picture1.CurrentY
      ' FILTRA EQUIPOS POR SECTOR PRODUCTIVO
      If SECTORMAQUINA(CODEQ$(U&)) = SECTOR Or SECTOR = 0 Then
        If YINI(U&) <= Picture1.Height - HScroll1.Height - INCREY Then
           Picture1.CurrentX = 80
           Picture1.Print CODEQ$(U&);
           
           Picture4.CurrentX = Picture1.CurrentX
                EQUI$ = TRIM$(ECOARCH$("PADMAQ", CODEQ$(U&)))
           Picture4.Print EQUI$;
           
           YEQU(U&) = YINI(U&)
        End If
      End If
   Next U&
   '
20 If CAEQVIS% = 0 Then
      CAEQVIS% = (Picture1.Height - Line1.Y1 - HScroll1.Height - INCREY) / INCREY
   End If
   '
   If MAXSCROLL% <= 0 Then
      VScroll1.Max = 0
      If CAEQVIS% < CAEQ% Then
         VScroll1.Max = CAEQ% - CAEQVIS%
      End If
      MAXSCROLL% = VScroll1.Max
   End If
   VScroll1.Max = MAXSCROLL%
   VScroll1.Value = PRIEQVIS% - 1
   '
End Sub
'

Sub PREGANTT()
   '
   If NODIBUJA% = 1 Then Exit Sub
   '
   CONTROCLICK = (-1)
   'TOPE = List1.TopIndex
   'ACTI = List1.ListIndex
   cabarr% = 0
   '
   SECTOR = 0
   If Combo1.ListIndex > -1 Then SECTOR = XVALO(Mid$(Combo1.List(Combo1.ListIndex), 1, 7))
   '
   NOFASE$ = TRIM$(Left$(List1.TEXT, 24))
   For U& = 1 To List1.ListCount
      NOFA$ = TRIM$(Left$(List1.List(U& - 1), 24))
      'NOFA$ = NORFB$(U&)
      ' COLO$ = Chr$(0) + Chr$(230) + Chr$(64)
      COLO$ = Chr$(225) + Chr$(197) + Chr$(149) ' NARANJA
      '
      If List1.Selected(U& - 1) = True Then
         'COLO$ = Chr$(255) + Chr$(0) + Chr$(0)
         COLO$ = Chr$(192) + Chr$(211) + Chr$(226) ' CELESTE
      End If
      '
      If List1.ListIndex = U& - 1 Then
           'COLO$ = Chr$(128) + Chr$(0) + Chr$(128)
           COLO$ = Chr$(90) + Chr$(118) + Chr$(80)   ' AZUL
        ElseIf Left$(NOFA$, 9) = Left$(NOFASE$, 9) Then
           'COLO$ = Chr$(200) + Chr$(0) + Chr$(200)   ' PARA COLOREAR LAS ASOCIADAS
         COLO$ = Chr$(204) + Chr$(217) + Chr$(199) ' VERDE CLARO
      End If
      '
      '
      COLOPRE$ = Chr$(220) + Chr$(220) + Chr$(0)
      V& = Val(Mid$(List1.List(U& - 1), 25))
      LONGMATIE& = Len(MATIE$(V&))
      For JJ& = 1 To LONGMATIE& Step 36
         MAQUI$ = Mid$(MATIE$(V&), JJ&, 24)
         If TRIM$(MAQUI$) = "" Then Exit For
         '
         FEHOR$ = Mid$(MATIE$(V&), JJ& + 24, 12)
         FEINI% = CVI(Mid$(FEHOR$, 1, 2))
         HOINI% = CVI(Mid$(FEHOR$, 3, 2))
         FEFIN% = CVI(Mid$(FEHOR$, 5, 2))
         HOFIN% = CVI(Mid$(FEHOR$, 7, 2))
         FEFINPREP% = CVI(Mid$(FEHOR$, 9, 2))
         HOFINPREP% = CVI(Mid$(FEHOR$, 11, 2))
         If FEINI% > FEFIN% Then
            FEINI% = 0
            HOINI% = 0
            FEFIN% = 0
            HOFIN% = 0
            FEFINPREP% = 0
            HOFINPREP% = 0
         End If
         '
         ' *** VERIFICA, ANTES DE IR A DIBUJAR, QUE EL INICIO ESTE ANTES DE LA FECHA HASTA DE LA GRAFICA
         ' *** Y EL FIN DESPUES DE LA FECHA DESDE DE LA GRAFICA
         If FEINI% > DESPERIO(CAPERIO% + 1) Then GoTo 40
         If FEFIN% < DESPERIO(1) Then GoTo 40
         '
         For K% = 1 To 24 Step 6
            EQUI$ = Mid$(MAQUI$, K%, 6)
            If TRIM$(EQUI$) = "" Then Exit For
            ' FILTRA EQUIPOS POR SECTOR PRODUCTIVO
            If SECTORMAQUINA(EQUI$) = SECTOR Or SECTOR = 0 Then
                CBX% = cabarr%
                                
                Call DIBUBARRA(FEINI%, HOINI%, FEFIN%, HOFIN%, EQUI$, COLO$, 0)
                Call DIBUBARRA(FEINI%, HOINI%, FEFINPREP%, HOFINPREP%, EQUI$, COLOPRE$, 1)
                If cabarr% > CBX% Then
                   NUOFA$(cabarr%) = NOFA$
                End If
            End If
         Next K%
40    Next JJ&
      '
50 Next U&
   '
   HScroll1.LargeChange = CAPERIO%
   CONTROCLICK% = 0
   YACARGANTT% = 1
   '
End Sub
'

Sub CARLIORF()
   '
   List1.Clear
   List2.Clear
   List3.Clear: JJ& = 0
   Call COPYSTRU("ECOMAST", "CODIORFA")
   CAOFA% = 0
   CAEQ% = 0
   For U& = 1 To ULTREG&("PADMAQ")
      ZZ$ = REGLEIDO$("PADMAQ", U&)
      If TRIM$(Left$(ZZ$, 6)) <> "" Then
         CAEQ% = CAEQ% + 1
         CODEQ$(CAEQ%) = Left$(ZZ$, 6)
         STATEQ%(CAEQ%) = 0
      End If
   Next U&
   '
   ARCHOPERFAP$ = "OPERFAP"
   If CONSIMULACION% > 0 Then ' SIMULACION
      Call COPYSTRU("OPERFAP", "SOPERFAP")
      ARCHOPERFAP$ = "SOPERFAP"
   End If

   For U& = 1 To ULTREG&(ARCHOPERFAP$)
      '
      XX$ = REGLEIDO$(ARCHOPERFAP$, U&)
      NOFA$ = TRIM$(Left$(XX$, 12))
      If CONSIMULACION% < 1 Then ' QUE NO MUESTRE LAS SIMULACIONES.
        If Left$(TRIM$(NOFA$), 3) = "SF-" Then GoTo 30
      End If
      MAQUI$ = Mid$(XX$, 73, 6) + Mid$(XX$, 203, 18)

      'ITIEM$ = Mid$(XX$, 135, 8)
      ITIEM$ = Mid$(XX$, 135, 12)
      CIIX% = CVI(Mid$(XX$, 13, 2))
      '
      REFEOF$ = TRIM$(SAFETEXT$(VALOBADA("ORDEFABR", "ID_ORFA", "IDSUBOR='" & NOFA$ & "' AND COD_INTE='" & CStr(CIIX%) & "'")))
      '
      YACODI% = 0: YAREFE% = 0
      For K% = 1 To CAOFA%
         If CIOFA%(K%) = CIIX% Then YACODI% = 1
         If OFAREFE$(K%) = REFEOF$ And REFEOF$ <> "" Then
            YAREFE% = 1
         End If
         If NORFB$(K%) = NOFA$ Then
            GoTo 10
        End If
      Next K%
      CAOFA% = CAOFA% + 1
      K% = CAOFA%
      NORFB$(CAOFA%) = NOFA$
      CIOFA%(CAOFA%) = CIIX%
      OFAREFE$(CAOFA%) = REFEOF$
      MATIE$(CAOFA%) = ""
      TTXX$ = Space$(32)
      Call REPLA(TTXX$, NOFA$, 1, 24)
      Call REPLA(TTXX$, Str$(CAOFA%), 25, 8)
      List1.AddItem TTXX$
      If YACODI% = 0 Then
         List3.AddItem CODEXT$(CIIX%)
         TTYZ$ = Space$(64)
         Call REPLA(TTYZ$, CODEXT$(CIIX%), 1, 16)
         Call REPLA(TTYZ$, DESCRIT$(CIIX%), 17, 48)
         JJ& = JJ& + 1
         Call GRAREG("CODIORFA", TTYZ$, JJ&)
      End If
      '
      If YAREFE% = 0 And REFEOF$ <> "" Then
         List2.AddItem REFEOF$
      End If
      '
10    MATIE$(K%) = MATIE$(K%) + MAQUI$ + ITIEM$
      '
      For K% = 1 To 24 Step 6
         EQUIPO$ = Mid$(MAQUI$, K%, 6)
         If TRIM$(EQUIPO$) <> "" Then
            For J% = 1 To CAEQ%
               If CODEQ$(J%) = EQUIPO$ Then
                  STATEQ%(J%) = 1
                  GoTo 20
               End If
            Next J%
            Call MENSERR(24, "Equipo '" + TRIM$(EQUIPO$) + "'no Figura en Tablas\en O/F Número " + TRIM$(NOFA$) + "para Código " + TRIM$(CODEXT$(CIIX%)) + ".")
         End If
20    Next K%
      '
30 Next U&
   Call SETULTREG("CODIORFA", JJ&)
   Call FILESORT("CODIORFA", "", 1, 1, "ASC")
   Call TRALOCAL("CODIORFA", "")
   '
   J% = 0
   For i% = 1 To CAEQ%
      If STATEQ%(i%) = 1 Then
         J% = J% + 1
         CODEQ$(J%) = CODEQ$(i%)
         STATEQ%(J%) = 1
      End If
   Next i%
   CAEQ% = J%
   '
   ' *** CARGA GRILLA DE ORDENES DE FABRICACION
   For II& = 0 To List1.ListCount - 1
        TTXX$ = List1.List(II&)
        dgvFabricacion.Rows = dgvFabricacion.Rows + 1
        dgvFabricacion.TextMatrix(dgvFabricacion.Rows - 1, 0) = TTXX$
   Next II&
   dgvFabricacion.ColWidth(0) = dgvFabricacion.ColWidth(0) * 2
   
   ' *** CARGA GRILLA DE REFERENCIAS / O-COMPRA
   For II& = 0 To List2.ListCount - 1
        TTXX$ = List2.List(II&)
        dgvReferencia.Rows = dgvReferencia.Rows + 1
        dgvReferencia.TextMatrix(dgvReferencia.Rows - 1, 0) = TTXX$
        dgvReferencia.Row = dgvReferencia.Rows - 1: dgvReferencia.COL = 0
        dgvReferencia.CellAlignment = 1
   Next II&
   dgvReferencia.ColWidth(0) = dgvReferencia.ColWidth(0) * 2
   
   ' *** CARGA GRILLA DE REFERENCIAS / O-COMPRA
   For II& = 0 To List3.ListCount - 1
        TTXX$ = List3.List(II&)
        dgvArticulos.Rows = dgvArticulos.Rows + 1
        dgvArticulos.TextMatrix(dgvArticulos.Rows - 1, 0) = TTXX$
        dgvArticulos.Row = dgvArticulos.Rows - 1: dgvArticulos.COL = 0
        dgvArticulos.CellAlignment = 1
   Next II&
   dgvArticulos.ColWidth(0) = dgvArticulos.ColWidth(0) * 2
   
End Sub
'
Sub DIBUBARRA(FEINI%, HOINI%, FEFIN%, HOFIN%, EQUI$, COLO$, PREPAR%)
   '
   Static DHR0%, TUNOR
   If TUNOR < 1 Then
      'DHR0% = 2400: TUNOR = 0
      'For U& = 1 To ULTREG&("TURGEN")
      '   X$ = REGLEIDO$("TURGEN", U&)
      '   DSX% = Asc(Left$(X$, 1))
      '   If DSX% > 0 Then
      '      DHR% = CVI(Mid$(X$, 2, 2))
      '      HHR% = CVI(Mid$(X$, 4, 2))
      '      If HHR% > DHR% Then
      '         If DHR% < DHR0% Then DHR0% = DHR%
      '         TT1 = 60 * Int(DHR% / 100) + (DHR% Mod 100)
      '         TT2 = 60 * Int(HHR% / 100) + (HHR% Mod 100)
      '         If (TT2 - TT1) > TUNOR Then TUNOR = TT2 - TT1
      '      End If
      '   End If
      'Next U&
      'DHR0% = 60 * Int(DHR0% / 100) + (DHR0% Mod 100)
   End If
   TUNOR = 24 * 60: DHR0% = 0
   '
   If FEINI% > DESPERIO(CAPERIO% + 1) Then
      Exit Sub
   End If
   '
   If FEFIN% < DESPERIO(1) Then
      Exit Sub
   End If
   '
   If FEINI% < DESPERIO(1) Then
      X1 = XINI(1)
      GoTo 10
   End If
   '
   For II% = 1 To CAPERIO%
      If FEINI% = DESPERIO(II%) Then
         MINU1% = 60 * Int(HOINI% / 100) + (HOINI% Mod 100)
         X1 = XINI(II%) + APERIO * (MINU1% - DHR0%) / TUNOR
         GoTo 10
      End If
   Next II%
   Exit Sub
   '
10 If FEFIN% >= DESPERIO(CAPERIO% + 1) Then
      X2 = XINI(CAPERIO% + 1)
      GoTo 20
   End If
   '
   For II% = 1 To CAPERIO%
      If FEFIN% = DESPERIO(II%) Then
         MINU2% = 60 * Int(HOFIN% / 100) + (HOFIN% Mod 100)
         X2 = XINI(II%) + APERIO * (MINU2% - DHR0%) / TUNOR
         GoTo 20
      End If
   Next II%
   '
20 For II% = PRIEQVIS% To CAEQ%
      If CODEQ$(II%) = EQUI$ Then
         Y1 = YINI(II%)
         GoTo 30
      End If
   Next II%
   Exit Sub
   '
30 If Y1 < 100 Then Exit Sub
   If X2 > XMAXI Then X2 = XMAXI
   ROJO% = Asc(Left$(COLO$, 1))
   VERDE% = Asc(Mid$(COLO$, 2, 1))
   AZUL% = Asc(Mid$(COLO$, 3, 1))
   Picture1.Line (X1, Y1 + PREPAR% * 60)-(X2, Y1 + AltoBarra%), RGB(ROJO%, VERDE%, AZUL%), BF
   Picture1.Line (X1 + 12, Y1 + AltoBarra%)-(X2, Y1 + AltoBarra% + 18), QBColor(0), BF
   Picture1.Line (X2 - 12, Y1 + 25)-(X2, Y1 + AltoBarra% + 18), QBColor(0), BF
   '
   If PREPAR% = 0 Then
     cabarr% = cabarr% + 1
     XINI1(cabarr%) = X1
     XFIN1(cabarr%) = X2
     YINI1(cabarr%) = Y1
   End If
   '
End Sub
'

Sub MUEDATORF(NORFA$)
   '
   MUORFACT.Hide
   DoEvents
   If CONSIMULACION% > 0 Then
      MUORFACT.CONSIMULACION% = CONSIMULACION%
   End If
   
   MUORFACT.Label2 = NORFA$
   MUORFACT.Left = XMUOFA
   MUORFACT.Top = YMUOFA
   MUORFACT.Show ' 1
   '
End Sub


Sub CONAPLI(APLINVO$, TERMINA%)
   TERMINA% = 1
   Select Case APLINVO$
           '
           'TRANSACCIONES
           '
        Case "REGENVARMAQ"
           Call Command2_Click                  ' RECALCULO GENERAL DE MAQUINAS
        Case "OPEPROYOF"
           Call CONCAR07.Command7_Click         ' OPERACIONES PROYECTADAS POR O/F
           '
           'CONSULTAS Y LISTADOS
           '
        Case "CARMAPAN"
           Call Command8_Click                  ' CONSULTAS Y LISTADOS POR PANTALLA
        Case "OPEPENOF"
           Call CONCAR07.Command26_Click        ' OPERACIONES PENDIENTES POR O/F
        Case "OPEPENMAQ"
           Call CONCAR07.Command15_Click        ' OPERACIONES PENDIENTES POR MAQUINA
        Case "CARMAIMP"
           Call Command7_Click                  ' CONSULTAS Y LISTADOS POR IMPRESORA
        Case "PLANICAR"
            Call LCARMA07.Command28_Click       ' PLANILLA DE CARGA
        Case "FENTPROY"
            Call LCARMA07.Command32_Click       ' FECHA DE ENTREGA PROYECTADA
        Case "PROCARMAQ"
            Call LCARMA07.Command7_Click        ' PROGRAMA DE CARGA DE MAQUINA
        '
        'CONFIGURACION GENERAL CARGA DE MAQUINAS
        '
        Case "CONGENCARMAQ"
            Call Command3_Click                 ' CONFIGURACION GENERAL CARGA DE MAQUINAS
        Case "GENPLANTA"
            Call Command3_Click
        Case "MAQUIEQUI"
            Call Command3_Click
        Case "RECUGENE"
            Call Command3_Click
        Case "TURNPROD"
            Call CFGCRP07.Command20_Click       ' TURNOS DIARIOS
        Case "CALINFERI"
            Call CFGCRP07.Command17_Click       ' CALENDARIO DE FERIADOS
        Case "HORESEQUI"
            Call CFGCRP07.Command6_Click        ' HORARIOS ESPECIALES POR EQUIPO
        Case "PARAEQUIPRO"
            Call CFGCRP07.Command99_Click       ' PARADAS DE EQUIPO PROGRAMADAS
        Case "PLAUNIPROD"
            Call CFGCRP07.Command1_Click        ' PLANTAS Y UNIDADES PRODUCTIVAS
        Case "SECENTPROD"
            Call CFGCRP07.Command2_Click        ' SECCIONES Y CENTROS DE PRODUCCION
        Case "MAQEQUIPROD"
            Call CFGCRP07.Command3_Click        ' MAQUINASY EQUIPOS DE PRODUCCION
        Case "DEPOALMA"
            Call CFGCRP07.Command18_Click       ' DEPOSITOS Y ALMACENES
        Case "ROTPERSEC"
            Call CFGCRP07.Command24_Click       ' ROTACION DE PERSONAL POR SECCION
        Case Else
           TERMINA% = 0
   End Select
End Sub

Function SECTORMAQUINA(CODMAQ$)
   '
   CONDI$ = "CODIGO_EQU = '" & CODMAQ$ & "'"
   SECT = XVALO(VALOBADA("EQUIPOS", "UBICA_EQU", CONDI$, "N0MESS"))
   If SECT = 0 Then
     FGR$ = FILTERGETRECORD$("PADMAQ", 1, AJUSTI$(MAQX$, 6))
     SECT = CVS(Mid$(FGR$, 37, 4))
   End If
   SECTORMAQUINA = SECT
   '
End Function





