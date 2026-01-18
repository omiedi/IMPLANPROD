VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FIKARDEXAM 
   BackColor       =   &H00FBCA9D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha de Movimientos de Material (Ficha Kardex)"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   4920
      TabIndex        =   29
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   26
      Top             =   600
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ComboBox Combo1 
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
      ForeColor       =   &H00000000&
      Height          =   345
      ItemData        =   "FIKARDEXAM.frx":0000
      Left            =   7200
      List            =   "FIKARDEXAM.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   180
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FBCA9D&
      Caption         =   "DESCRIPCIÓN"
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
      Left            =   105
      TabIndex        =   19
      Top             =   630
      Width           =   4755
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
         Height          =   735
         Left            =   75
         TabIndex        =   21
         Top             =   240
         Width           =   4605
      End
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   3
      Left            =   3570
      Picture         =   "FIKARDEXAM.frx":0004
      TabIndex        =   17
      Top             =   215
      Width           =   175
   End
   Begin VB.CommandButton BOTSEL 
      Caption         =   "..."
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
      Index           =   2
      Left            =   5820
      Picture         =   "FIKARDEXAM.frx":030E
      TabIndex        =   16
      Top             =   210
      Width           =   175
   End
   Begin VB.TextBox Text4 
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
      Left            =   5295
      TabIndex        =   15
      Text            =   " "
      Top             =   210
      Width           =   540
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FBCA9D&
      Caption         =   "RANGO DE FECHAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   5280
      TabIndex        =   8
      Top             =   840
      Width           =   4260
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FBCA9D&
         Caption         =   "Incluir Antiguos"
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
         Left            =   2520
         TabIndex        =   20
         Top             =   675
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton BOTSEL 
         Caption         =   "..."
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
         Index           =   1
         Left            =   3945
         Picture         =   "FIKARDEXAM.frx":0618
         TabIndex        =   12
         Top             =   360
         Width           =   175
      End
      Begin VB.CommandButton BOTSEL 
         Caption         =   "..."
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
         Index           =   0
         Left            =   2030
         Picture         =   "FIKARDEXAM.frx":0922
         TabIndex        =   11
         Top             =   360
         Width           =   175
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
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
         Left            =   960
         TabIndex        =   10
         Text            =   " "
         Top             =   360
         Width           =   1065
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
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
         Left            =   2865
         TabIndex        =   9
         Text            =   " "
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
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
         Left            =   0
         TabIndex        =   14
         Top             =   435
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
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
         Left            =   2040
         TabIndex        =   13
         Top             =   435
         Width           =   750
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
      Height          =   285
      Left            =   1020
      TabIndex        =   3
      Top             =   210
      Width           =   2535
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      FillColor       =   &H00400000&
      ForeColor       =   &H80000008&
      Height          =   1755
      Left            =   9600
      ScaleHeight     =   1725
      ScaleWidth      =   2985
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   240
         ScaleHeight     =   165
         ScaleWidth      =   1575
         TabIndex        =   30
         Top             =   960
         Width           =   1605
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   31
            Top             =   -240
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command2 
         Height          =   615
         Left            =   50
         Picture         =   "FIKARDEXAM.frx":0C2C
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Búsqueda por Texto "
         Top             =   150
         Width           =   600
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
         Height          =   620
         Left            =   1500
         Picture         =   "FIKARDEXAM.frx":1F92
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   150
         Width           =   600
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   780
         Picture         =   "FIKARDEXAM.frx":20DC
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprime Ficha de Movimientos de Material (Ficha Kardex)"
         Top             =   150
         Width           =   600
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Un Momento Por Favor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   0
         TabIndex        =   32
         Top             =   1320
         Visible         =   0   'False
         Width           =   2040
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   5625
      Left            =   105
      TabIndex        =   24
      Top             =   1800
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   9922
      _Version        =   393216
      Rows            =   23
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   8433840
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      Redraw          =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3480
      OleObjectBlob   =   "FIKARDEXAM.frx":2746
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   -240
      OleObjectBlob   =   "FIKARDEXAM.frx":297A
      TabIndex        =   27
      Top             =   840
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   195
      Left            =   720
      TabIndex        =   25
      Top             =   16120
      Width           =   870
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   9195
      Picture         =   "FIKARDEXAM.frx":29F4
      Stretch         =   -1  'True
      ToolTipText     =   "Detalle Movimientos por Lote Activo"
      Top             =   135
      Width           =   360
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Partidas"
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
      Left            =   6060
      TabIndex        =   23
      Top             =   290
      Width           =   1065
   End
   Begin VB.Label Label5 
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
      Height          =   315
      Left            =   4395
      TabIndex        =   18
      Top             =   210
      Width           =   405
   End
   Begin VB.Label Label4 
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
      Height          =   285
      Left            =   5280
      TabIndex        =   7
      Top             =   555
      Width           =   4260
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dep"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4800
      TabIndex        =   6
      Top             =   290
      Width           =   420
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U-M"
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
      Left            =   3830
      TabIndex        =   5
      Top             =   290
      Width           =   510
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
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
      Left            =   0
      TabIndex        =   4
      Top             =   285
      Width           =   960
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuDepMov 
         Caption         =   "Depurar Tabla de Movimientos de Stocks"
      End
   End
   Begin VB.Menu edicion 
      Caption         =   "Edición"
   End
End
Attribute VB_Name = "FIKARDEXAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARBUS$, ABUSQUE$



Sub DepurarMovistoAmbest()
    REINGRESAFECHA% = 0
    If DERACCE%("DESCAMOV", "Descarga de Movimientos de Stock") < 2 Then
       GoTo 99
    End If
    '
'10  OPX% = COMALTER%("Confirmar Depuración\\Si, Iniciar Depuración\No, Cancelar")
'    If OPX% < 1 Or OPX% > 2 Then GoTo 99
'    If OPX% = 1 Then
       SOLONOTRAZABLES% = 0
       OPX2% = COMALTER%("Opciones de Depuración\\Depurar Solo ´NO´ Trazables\Depurar Todos")
       If OPX2% < 1 Or OPX2% > 2 Then GoTo 99
       If OPX2% = 1 Then SOLONOTRAZABLES% = 1
'    End If
    '
'    If OPX% = 2 Then GoTo 40
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
    ' CREACION DE LA TABLA MOVIPRU
    If REINGRESAFECHA% > 0 Then GoTo 20 ' SI ES PARA REINGRESAR LA FECHA NO VUELVE A CREAR LA TABLA DE RESPALDO
    Call COMUNI("FLEXOFT Debe Ahora Crear la Tabla de Respaldo.\  \Este Proceso Puede Demandar Varios Minutos\y NO DEBE INTERRUMPIRSE Aunque el Equipo Parezca Haberse 'Colgado' !!!")
    Screen.MousePointer = 11
    TIMUT = FLEXCONN.CommandTimeout
    FLEXCONN.CommandTimeout = 999
    Call CREACAMPO("STOCKS", "MOVISTO", "Marborra", 3, 0, 1)
'    Call COPIA_STRUC_MOVISTO
    Call COPIA_STRUC_MOVISTO_A_MOVIPRU

    FLEXCONN.CommandTimeout = TIMUT
    Screen.MousePointer = 1
    Call COMUNI("Tabla de Respaldo Creada OK.")
    '
20  FEX = FF%: FEC1$ = FETEXCOR1$(DIPOST(FEX))
    CONDI$ = "FechMov < '" & FEC1$ & "' AND CodiMovi <> 'KU' "
    FIN& = CantRegistros("MOVISTO", CONDI$)
    If FIN& < 1 Then
        Call MENSERR(24, "No Existen Movimientos Anteriores\a la Fecha: " + FECHATEX$(FF%) + ".")
        Screen.MousePointer = 1
        GoTo 99
    End If
    '
    
    If FIN& > 10000 Then
        Call MENSERR(24, "Cantidad de Movimientos a Transferir (" & FIN& & ") Por Vez\Supera la Cantidad Recomendada de 10.000\Ingrese una Fecha Menor a la yá Ingresada")
        Screen.MousePointer = 1
        REINGRESAFECHA% = 1
        GoTo 13
    End If
    SQLX$ = " SELECT * FROM MOVISTO WHERE " & CONDI$
    SQLX$ = SQLX$ + " ORDER BY FEREMOVI"
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
    SQLY$ = " SELECT * FROM MOVIPRU"
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
    ' COPIA LOS REGISTROS A MOVIPRU
    Screen.MousePointer = 11
    U& = 0
    With MOVIAC
      Do While Not .EOF
         U& = U& + 1
         Call TRACE(20, U&, FIN&)
         '
         If SOLONOTRAZABLES% > 0 Then
            CI1% = XVALO(!Cod_Inte)
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
         MOVIVI!YAMIGRADO = 0
         MOVIVI.Update
         '
         !Marborra = 1
         .Update
29    .MoveNext
      Loop
    End With
    '
    MOVIAC.Close: Set MOVIAC = Nothing
    MOVIVI.Close: Set MOVIVI = Nothing
    '
    'GOTO 10
    
     'BORRA LOS REGISTROS MARCADOS CON MARBORRA Y LOS DE ACUMULADOS
40   TIMEANT% = FLEXCONN.CommandTimeout
     FLEXCONN.CommandTimeout = 999
     '
     'ELIMINO LOS REGISTROS MARBORRA = 1
     PASO% = 0
43   SQLX1$ = " DELETE TOP(100) FROM MOVISTO WHERE MARBORRA = 1"
     PASO% = PASO% + 1
     FLEXCONN.Execute (SQLX1$)
     If CantRegistros("MOVISTO", "MARBORRA = 1", "NOMESS") > 0 And PASO% < 120 Then
       GoTo 43
     End If
     If CantRegistros("MOVISTO", "MARBORRA = 1", "NOMESS") > 0 Then
       Call COMUNI("No se Han Podido Eliminar los Registros De La Tabla Movisto (Condición: MARBORRA = 1)")
     End If
     '
     'ELIMINO LOS REGISTROS CodiMovi = 'KU'
     PASO% = 0
44   SQLX2$ = " DELETE TOP(100) FROM MOVISTO WHERE CodiMovi = 'KU'"
     PASO% = PASO% + 1
     FLEXCONN.Execute (SQLX2$)
     If CantRegistros("MOVISTO", "CodiMovi = 'KU'", "NOMESS") > 0 And PASO% < 120 Then
       GoTo 44
     End If
     If CantRegistros("MOVISTO", "CodiMovi = 'KU'", "NOMESS") > 0 Then
            Call COMUNI("No se Han Podido Eliminar los Registros De La Tabla Movisto (Condición: CodiMovi = 'KU')")
     End If

     
'     Call BORRAREGISTROS("MOVISTO", "Marborra = 1", RAFE&)
'     Call BORRAREGISTROS("MOVISTO", "CodiMovi = 'KU'", RAFE&)
     FLEXCONN.CommandTimeout = TIMEANT%
'     If CantRegistros("MOVISTO", "MARBORRA = 1", "NOMESS") > 0 And PASO% < 10 Then GoTo 40
'     If CantRegistros("MOVISTO", "CodiMovi = 'KU'", "NOMESS") > 0 And PASO% < 10 Then GoTo 40
    ' GRABA LOS REGISTROS CON LOS TOTALES
50  'FECHINVE$ = InputBox$("Ingrese la Fecha de Corte", "Depuración de los Movimientos de Stocks")
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
    SQLX$ = SQLX$ + " FROM MOVIPRU WITH(NOLOCK)"
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
         CIXI% = XVALO(!Cod_Inte)
         DEPX% = XVALO(!DEPOMOVI)
         CANSA = XVALO(!SUIN) - XVALO(!SUSAL)
         '
         Call TRACE(20, CONTI&, TOT&)
         If Abs(CANSA) > 0.00005 Then
           With MOVIAC
             .AddNew
             !FechMov = CVDAT(FF%)
             !Cod_Inte = CIXI%
             !COD_EXTE = CODEXT$(CIXI%)
             !Descrip = Left$(DESCRIT$(CIXI%), 64)
             !DEPOMOVI = DEPX
             If CANSA < 0 Then
                 !CantSalid = Abs(CANSA)
                 !CantIngre = 0
               Else
                 !CantIngre = CANSA
                 !CantSalid = 0
             End If
             !CODIMOVI = "KU" 'MOVIMIENTO ACUMULADOs
             !ReferCor = "Acumulados al " & CVDAT(FF%) & " "
             !IDENLOTE = ""
             !DoPriCor = "QM" + FECHATEX$(FF%)
             !DoPriLar = "QM" + FECHATEX$(FF%)
             !DoSecCor = "QM" + FECHATEX$(FF%)
             !DoSecLar = "QM" + FECHATEX$(FF%)
                NORIT& = NORIT& + 1
             !NuorItem = NORIT& Mod 10000
             !NUPEDCLI = 0
             !ValoUnit = 0
             !MoneVal = 0 '"" cambie a 0 xq es campo integer
             !Nume_Clie = 0
             !Nume_Prov = 0
             !FEREMOVI = Now
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
    
    Call TRANSFERIR_MOVIPRU_A_MOVISTOMDB
    Call COMUNI("Proceso Completo")
    '
99  Screen.MousePointer = 1

End Sub
Sub EXPORTAR_MOVIPRU_A_MOVISTOMDB(OK%)

    SQLX$ = " SELECT * FROM MOVIPRU WHERE YAMIGRADO < 1 OR YAMIGRADO IS NULL"
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
    'MDB
    SQLY$ = " SELECT * FROM MOVISTO"
35  On Error Resume Next
    Set MOVIVI = BaDaMoviExt.OpenRecordset(SQLY$, dbOpenDynaset)

    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       TXTX$ = "Se ha Producido un Error " + vbCrLf
       TXTX$ = TXTX$ + "Descripción: " & Err & vbCrLf
       TXTX$ = TXTX$ + "Comuniquese Con su Soporte de Sistemas."
       OK% = -1
       Exit Sub
    End If
    On Error GoTo 0
    '
    ' COPIA LOS REGISTROS A MOVIPRU
    Screen.MousePointer = 11
    U& = 0
    With MOVIAC
      Do While Not .EOF
         U& = U& + 1
         Call TRACE(20, U&, FIN&)
         '
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
         !YAMIGRADO = 1
         .Update
29    .MoveNext
      Loop
    End With
    '
    MOVIAC.Close: Set MOVIAC = Nothing
    MOVIVI.Close: Set MOVIVI = Nothing
    OK% = 1
End Sub

Sub TRANSFERIR_MOVIPRU_A_MOVISTOMDB()
   
   '1.-TRANSFIERE A UNA TABLA MOVIPRU DE RESPALDO Y LUEGO BORRA LA TABLA MOVIVIE (SIEMPRE AGREGA INFORMACION)
   '2.-TRANSFIERE A UNA TABLA MDB (SIEMPRE AGREGA INFORMACION)
   '3.-GENERA LA EN LA RUTA DEL PENDRIVE
   '4.-Preguntar si desea pasar la informacion a la unidad externa
'    Call EXPORTAR_MOVIVIE_A_MOVIPRU
    Call EXPORTAR_MOVIPRU_A_MOVISTOMDB(OK%)
    If OK% < 0 Then Exit Sub
'    Call BORRAREGISTROS("MOVIVIE", "Marborra = 1", RAFE&)

End Sub
Sub EXPORTAR_MOVIVIE_A_MOVIPRU()

    SQLX$ = " SELECT * FROM MOVIVIE "
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
    SQLY$ = " SELECT * FROM MOVIPRU"
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
    ' COPIA LOS REGISTROS A MOVIPRU
    Screen.MousePointer = 11
    U& = 0
    With MOVIAC
      Do While Not .EOF
         U& = U& + 1
         Call TRACE(20, U&, FIN&)
         '
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
         MOVIVI!YAMIGRADO = 0
         MOVIVI.Update
         '
         !Marborra = 1
         .Update
29    .MoveNext
      Loop
    End With
    '
    MOVIAC.Close: Set MOVIAC = Nothing
    MOVIVI.Close: Set MOVIVI = Nothing
    '
End Sub
Private Sub BOTSEL_Click(Index As Integer)
    If Index < 2 Then
         Call SELECHO("SELFECHA")
         If VALACT1$("SELFECHA") <> "" Then
           If Index = 0 Then
                Text11.TEXT = VALACT1$("SELFECHA")
              Else
                Text12.TEXT = VALACT1$("SELFECHA")
           End If
         End If
      ElseIf Index = 2 Then
        Call SELECHO("TADEPOSI")
        If VALACT1$("TADEPOSI") <> "" Then
          Text4 = TRIM$(VALACT1$("TADEPOSI"))
        End If
      ElseIf Index = 3 Then
        Call SELECHO("MASTER")
        If VALACT1$("MASTER") <> "" Then
          Text1 = VALACT1$("MASTER")
        End If
    End If
End Sub

Private Sub Check1_Click()
   Call GENKARDEX
End Sub

Private Sub Combo1_Click()
  'SeleccionEnList
'  For I = 0 To List1.ListCount - 1
'     LOTMOV$ = Mid$(List1.List(I), 129, 16)
'     If LOTMOV$ = Combo1.TEXT Then
'          List1.Selected(I) = True
'        Else
'          List1.Selected(I) = False
'     End If
'   Next I
  For I = 1 To GRID.Rows - 1
     LOTMOV$ = TRIM$(GRID.TextMatrix(I, 4))
     GRID.Row = I
     For j& = 0 To GRID.Cols - 1
        GRID.COL = j
        If LOTMOV$ = TRIM$(Combo1.TEXT) And TRIM$(Combo1.TEXT) <> "" Then
             GRID.CellBackColor = RGB(0, 0, 255)
             GRID.CellForeColor = RGB(255, 255, 255)
           Else
             GRID.CellBackColor = RGB(255, 255, 255)
             GRID.CellForeColor = RGB(0, 0, 0)
        End If
     Next j&
   Next I
End Sub

Private Sub Command1_Click()
  Call FINAL("")
End Sub

Private Sub Command2_Click()
   
   ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", , ARBUS$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     For U& = 1 To GRID.Rows - 1
      For I& = 0 To GRID.Cols - 1
       If InStr(UCase$(GRID.TextMatrix(U&, I&)), ABUSQUE$) > 0 Then
        GRID.Row = U&: GRID.COL = I&
        If GRID.TopRow > GRID.Row Then GRID.TopRow = GRID.Row
        On Error Resume Next
        VISIROWS = (GRID.Height / GRID.RowHeight(j&)) - 4
        If Err > 0 Then
         VISIROWS = 1
        End If
        On Error GoTo 0
         If VISIROWS > 0 Then
            If GRID.TopRow < GRID.Row - VISIROWS Then
              GRID.TopRow = GRID.Row - VISIROWS
            End If
         End If
        GRID.SetFocus
        Exit Sub
       End If
       Next I&
     Next U&
     '
     Call MENSERR(24, "Argumento no Encontrado.")
   End If

End Sub

Private Sub Command3_Click()
'    '
'    Call ABRECONEXION
'    '
'    For ii% = 1 To NUMAS%
'        CIXI% = ii%
'        If STATUIT(CIXI%) < 0 Then GoTo 80
'        If CantRegistros("movisto", "cod_inte=" & CIXI% & " and refercor='' and cantsalid < 0.05 and cantingre < 0.05") > 1 Then
'            OPX% = COMALTER%("Codigo " & CODEXT$(CIXI%) & " a Depurar\\No, Cancelar\Si, Continuar")
'            If OPX% = 1 Then GoTo 80
'        End If
'        '
'        Call BORRAREGISTROS("movisto", "cod_inte=" & CIXI% & " and refercor='' and cantsalid < 0.05 and cantingre < 0.05", REGAF&)
'        '
'80  Next ii%

'    Call ABRECONEXION
'    '
'    For ii% = 1 To NUMAS%
'        CIXI% = ii%
'        If STATUIT(CIXI%) < 0 Then GoTo 80
'        If CantRegistros("movisto", "cod_inte=" & CIXI% & " and refercor='' and (cantsalid > 0.05 or cantingre > 0.05)") > 1 Then
'            OPX% = COMALTER%("Codigo " & CODEXT$(CIXI%) & " a Depurar\\No, Cancelar\Si, Continuar")
'            If OPX% = 1 Then GoTo 80
'        End If
'        sqlx$ = "select * from movisto where cod_inte=" & CIXI% & " and refercor='' and (cantsalid > 0.05 or cantingre > 0.05)"
'        '
'        Dim rst As ADODB.Recordset
'        Set rst = New ADODB.Recordset
'        rst.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'        '
'        With rst
'            Do While Not .EOF
'                dopricor$ = TRIM$(SAFETEXT$(!dopricor))
'                DEPO% = XVALO(!depomovi)
'                CANTI# = XVALO(!cantsalid)
'                '
'                CONDI$ = "cod_inte=" & CIXI% & " and depomovi=" & DEPO% & " and dopricor='" & dopricor$ & "'"
'                careg& = CantRegistros("movisto", CONDI$)
'                '
'                If careg& > 1 Then
'                    !cantsalid = 0
'                    !cantingre = 0
'                    .Update
'                     MsgBox "entro " & careg& & vbCrLf & CONDI$
'                End If
'                .MoveNext
'            Loop
'        End With
'        On Error Resume Next
'        rst.Close
'        Set rst = Nothing
'        On Error GoTo 0
'80  Next ii%
    '
    Call COMUNI("proceso finalizado")
    '
End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
   ZBLA$ = REGBLAN$("COKARDEX")
   JJ& = 0
   FIN& = GRID.Rows - 1 ' List1.ListCount
   '\\\OT-06-0390-MH-13/09/06///
   TODO% = 0
'   For KU& = 1 To FIN&
'     If List1.Selected(KU& - 1) = True Then GoTo 10
'   Next KU&
   TODO% = 1
   GoTo 20
   '
10 OPXX% = COMALTER%("Seleccione Tipo de Listado\\Texto Seleccionado\Todo\Cancelar")
   TODO% = OPXX% - 1
   If TODO% > 1 Then GoTo 99
   '
20 For J1% = 1 To GRID.Rows - 1   'List1.ListCount
     '
     J2& = J1% + 1
     Call TRACE(24, J2&, FIN&)
     '
'     If TODO% = 1 Or List1.Selected(J1% - 1) Then
        '
        'X$ = List1.List(J1% - 1)
        FEC1% = FECHANUM(GRID.TextMatrix(J1%, 0))  ' Left$(X$, 8))
        CM1$ = GRID.TextMatrix(J1%, 1)             ' Mid$(X$, 11, 2)
        DOPRI1$ = GRID.TextMatrix(J1%, 2)          ' Mid$(X$, 15, 12)
        DOSEC1$ = GRID.TextMatrix(J1%, 3)          ' Mid$(X$, 29, 12)
        REFE1$ = GRID.TextMatrix(J1%, 5)           ' Mid$(X$, 43, 26)
        dep1% = XVALO(GRID.TextMatrix(J1%, 6))     ' Mid$(X$, 70, 3))
        ENTRA1# = XVALO(GRID.TextMatrix(J1%, 7))   ' Mid$(X$, 75, 10))
        SALE1# = XVALO(GRID.TextMatrix(J1%, 8))    ' Mid$(X$, 85, 10))
        SALD1# = XVALO(GRID.TextMatrix(J1%, 9))    ' Mid$(X$, 95, 12))
        '
        YY$ = ZBLA$
        Call REPLA(YY$, MKI$(FEC1%), 1, 2)
        CI1% = CODINT%(Text1.TEXT)
        Call REPLA(YY$, MKI$(CI1%), 3, 2)
        Call REPLA(YY$, CM1$, 21, 2)
        Call REPLA(YY$, DOPRI1$, 23, 10)
        Call REPLA(YY$, DOSEC1$, 33, 10)
        Call REPLA(YY$, REFE1$, 43, 30)
        Call REPLA(YY$, Chr$(dep1%), 83, 1)
        Call REPLA(YY$, MKD$(ENTRA1#), 88, 8)
        Call REPLA(YY$, MKD$(SALE1#), 96, 8)
        Call REPLA(YY$, MKD$(SALD1#), 104, 8)
        JJ& = JJ& + 1
        Call GRAREG("COKARDEX", YY$, JJ&)
        '
'     End If
     '
   Next J1%
   Call SETULTREG("COKARDEX", JJ&)
  
   '\\\OT-06-0390-MH-FIN///
   '
   Call HEADERS("COKARDEX", "")
   If Val(Text4) > 0 Or ECOARCH$("TADEPOSI", Chr$(Val(Text4))) <> "" Then
      Call HEADERS("COKARDEX", "Deposito: " & TRIM$(Text4) & " - " & ECOARCH$("TADEPOSI", Chr$(Val(Text4))))
   End If
   Call HEADERS("COKARDEX", "Artículo: " + Text1.TEXT + " - " + Label6.Caption)
   Call HEADERS("COKARDEX", "Período: " + Text11.TEXT + " - " + Text12.TEXT)
   Call FINAL("*LICARDEX")
   '
99 Command8.Enabled = True
   '
 End Sub

Private Sub Filtroporproveedor_Click()
    CI1% = CODINT%(Text1)
    If CI1% < 1 Then Exit Sub
    
    CAMPOS$ = "Cta.Prov/F9;R.Social/A40"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO, , , ANTOT)
'    If ANTOT < 5500 Then ANTOT = 5500
'    FRMZELECHO.Width = ANTOT + 1500

    SQLX$ = " EXECUTE ProveedorSegunCodigo  " & CI1%
    Dim RS As New ADODB.Recordset
    Set RS = FLEXCONN.Execute(SQLX$)
    FRMZELECHO.MSF1.Rows = 1
    j& = 0
    '
    With RS
      Do While Not .EOF
        NPROV = XVALO(!NUMPROV)
        RAZONPROVEEDOR$ = rasocli$(-1 * NPROV)
        '
        j& = j& + 1
        FRMZELECHO.msf2.Rows = j& + 1
        FRMZELECHO.msf2.TextMatrix(j&, 1) = NPROV
        FRMZELECHO.msf2.TextMatrix(j&, 2) = RAZONPROVEEDOR$
        .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    If j& = 0 Then
       Call COMUNI("No se Detectaron Entregas Ingresos Con Proveedor Para el Código Activo")
       Exit Sub
    End If
    
    FRMZELECHO.Caption = "Proveedores de Código: " & CODEXT$(CI1%)
    FRMZELECHO.Show 1
    
    PROVED = XVALO(RESP_ZELE_VECT(1))
    If PROVED > O Then
       Label8 = PROVED
       Label9 = rasocli$(-1 * PROVED)
    Else
       Label8 = ""
    End If

End Sub

Private Sub Form_Load()
   '
    
   Call APLICACIONSKINS(Me, Picture3)
   '
    UTILIZA_SKIN% = 0
    '
     EPAC$ = TRIM$(EMPREAC$)
'    VERANTER$ = "CONSTOCK"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    Call GRATITFOR(Caption)
    '
    LU$ = LUDAT$
    'BUSCA  LA RUTA
    'DEBE SELECCIONAR LA DEL ARCHIVO EN CUESTION
    'SI ELIGE ESTE ESTA VA A SER LA RUTA DONDE VA A GENERAR EL ARCHIVO MDB
    RutaArchivo$ = SELECCION_Archivo$("txt", CommonDialog1) '    EPAC$ = TRIM$(EMPREAC$)
    'aca leer solo el final del archivo Ambest4280.TXT
    POS1% = InStrRev(RutaArchivo$, "\")
    RUTAX$ = Mid$(RutaArchivo$, 1, POS1%)
    ARRCH$ = Mid$(RutaArchivo$, Len(RUTAX$) + 1, POS1%)
    If ARRCH$ = ("AMBEST4280") Or ARRCH$ = ("AMBEST4280.TXT") Then
        POS1% = InStrRev(RutaArchivo$, "\")
        RUTAX$ = Mid$(RutaArchivo$, 1, POS1%)
        If EXISTE%(RUTAX$ + "BDACCESS\STOCKS.MDB") < 1 Then
           Call COMUNI("Atención Se Va A Generar la Tabla de Respaldo Movil\En " & REPLACAR(RUTAX$ + "BDACCESS", "\", "/") & "")
        End If
        Call ABREBASE_MDB(RUTAX$)
        Screen.MousePointer = 1
       'mensaje se generara el archivio de respaldo en esta ruta.
       'validar si esta creada la base de datos
    End If
    Screen.MousePointer = 1
'    VERANTER$ = "CONSTOCK"
'    Call VERJOBID(OKEY%)
'    If OKEY% < 1 Then Call FINAL("")
'    '
'
'    If EPAC$ <> "" Then
'       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
'       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
'    End If
'    Call GRATITFOR(Caption)
'    '
'    LU$ = LUDAT$
    On Error Resume Next
'    Show
    On Error GoTo 0
    Openforms = DoEvents
    '
    Call ABRESTOCKS
    Call ABREPEDCLI
    Call ABRECOMPRAS
    Call ABREORFAB
    Call ABRESERVA
    '
    Call CREAR_STORED_PROCEDURE
   
    Text11 = "01/01/80"
    FEINIKAR$ = TRIM$(CONTROL("", "FEINIKAR"))
    FF% = FECHANUM(FEINIKAR$)
    If FF% > 0 Then Text11 = FECHATEX$(FF%)
    HOII% = HOY(HOS$)
    Text12 = HOS$
    Call Text4_Change
    '
    TEPRUEBA.Font = GRID.Font
    TEPRUEBA.FontSize = GRID.FontSize
    GRID.Cols = 10
    GRID.Row = 0
    NUCOLU% = 1: TEPRUEBA = "99/99/99M": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 1: GRID.TextMatrix(0, NUCOLU% - 1) = "Fecha"
    NUCOLU% = 2: TEPRUEBA = "OOM": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 1: GRID.TextMatrix(0, NUCOLU% - 1) = "CM"
    NUCOLU% = 3: TEPRUEBA = "MM-0000-00000000M": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 1: GRID.TextMatrix(0, NUCOLU% - 1) = "Transaccion"
    NUCOLU% = 4: TEPRUEBA = "MM-0000-00000000M": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 1: GRID.TextMatrix(0, NUCOLU% - 1) = "Doc.Numero"
    NUCOLU% = 6: TEPRUEBA = "ABCDEFGH1JKLMNOPQR": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 1: GRID.TextMatrix(0, NUCOLU% - 1) = "Referencia"
    NUCOLU% = 7: TEPRUEBA = "99M": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 6: GRID.TextMatrix(0, NUCOLU% - 1) = "Dp"
    NUCOLU% = 8: TEPRUEBA = "999999.99999M": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 6: GRID.TextMatrix(0, NUCOLU% - 1) = "Entrada"
    NUCOLU% = 9: TEPRUEBA = "999999.99999M": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 6: GRID.TextMatrix(0, NUCOLU% - 1) = "Salida"
    NUCOLU% = 10: TEPRUEBA = "999999.99999M": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 6: GRID.TextMatrix(0, NUCOLU% - 1) = "Saldo"
    NUCOLU% = 5: TEPRUEBA = "MM-999999-99": GRID.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: GRID.ColAlignment(NUCOLU% - 1) = 1: GRID.TextMatrix(0, NUCOLU% - 1) = "ID-Bulto"
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")

End Sub


Private Sub GRID_DblClick()
    '
    iRow& = GRID.Row
    CANTING = XVALO(GRID.TextMatrix(iRow&, 7))
    CANTSAL = XVALO(GRID.TextMatrix(iRow&, 8))
    COMO$ = TRIM$(GRID.TextMatrix(iRow&, 1))
    BRE$ = TRIM$(GRID.TextMatrix(iRow&, 2))
    CI1% = CODINT%(TRIM$(Text1))
    
    '
    If CANTING > 0 Then
       IDLOT$ = TRIM$(GRID.TextMatrix(iRow&, 4))
       If IDLOT$ <> "" Then
         OPX% = COMALTER("Opciones de Trabajo\\Consulta de Consumos Por Lote\Ver Datos Relacionados Según Tipo de Movimiento")
         If OPX% < 1 Then Exit Sub
         If OPX% = 1 Then
             On Error Resume Next
             CONSUMLOT07.Label4 = Text1
             If Err Then Exit Sub
             On Error GoTo 0
             CONSUMLOT07.Label1 = IDLOT$
             CONSUMLOT07.Show 1
             Exit Sub
         End If
       End If
     ElseIf CANTSAL > 0 Then
       On Error Resume Next
       LOCONSUM09.Label4 = TRIM$(GRID.TextMatrix(iRow&, 2))
       If Err Then Exit Sub
       On Error GoTo 0
       OPX% = COMALTER("Opciones de Trabajo\\Consulta de Consumos Por Lote\Ver Datos Relacionados Según Tipo de Movimiento")
       If OPX% < 1 Then Exit Sub
       If OPX% = 1 Then
        LOCONSUM09.Label5 = TRIM$(GRID.TextMatrix(iRow&, 0))
        LOCONSUM09.Label6 = TRIM$(GRID.TextMatrix(iRow&, 4))
        LOCONSUM09.Show 1
        Exit Sub
       End If
    End If
    '
    Select Case TRIM$(UCase(COMO$))
      Case "BR"
        CAMPOS$ = "O.Compra/F9;Cta./F8;Proveedor/A32;Id.Lote/A16"
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO, , , ANTOT)
        Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO, , , ANTOT)
        FRMZELECHO.Caption = "Orden de Compra Según Recepcion: " & BRE$
        FRMZELECHO.Width = 8600
        '
        NBR& = XVALO(Mid$(BRE$, 4))
        SQLX$ = "SELECT * FROM BOLRECEP WITH(NOLOCK) WHERE NUMEBORE = " & NBR& & " AND COD_INTE = " & CI1%
        Set RS = READSET(SQLX$): K& = 0
        Do While Not RS.EOF
          K& = K& + 1
          FRMZELECHO.msf2.Rows = K& + 1
          FRMZELECHO.msf2.TextMatrix(K&, 1) = XVALO(RS!NUMOC)
          FRMZELECHO.msf2.TextMatrix(K&, 2) = XVALO(RS!NUMPROV)
          FRMZELECHO.msf2.TextMatrix(K&, 3) = SAFETEXT$(RS!RASO_PROV)
          FRMZELECHO.msf2.TextMatrix(K&, 4) = SAFETEXT$(RS!IDENLOTE)
          RS.MoveNext
        Loop
        RS.Close
        Set RS = Nothing
        FRMZELECHO.Show 1
        NUOC& = XVALO(RESP_ZELE_VECT(1))
        DOCUNU& = NUOC&
        TIDOCU$ = "NOTA DE PEDIDO"
        NUDOCU$ = TRIM$(Str$(DOCUNU&))
        Call MUESTRADOC(0, "", TIDOCU$, NUDOCU$, NUDOCU$)

    End Select
    
End Sub

Private Sub GRID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   I& = GRID.MouseRow
   j& = GRID.MouseCol
   If I& = 0 Or I& > GRID.Rows Then Exit Sub
   
   If Button = 2 Then
       '
       GRID.Row = I&
       GRID.COL = j&
       GRID.CellBackColor = &H80FFFF   ' LA PINTA DE AMARILLA

       EPAC$ = TRIM$(UCase$(EMPREAC$))
       If InStr(EPAC$, "ASEMBLI") > 0 Then
            FEMOX1$ = GRID.TextMatrix(I&, 0)
            TIMOI1$ = GRID.TextMatrix(I&, 1)
            DOPRIX1$ = GRID.TextMatrix(I&, 2)
            LOTEA1$ = GRID.TextMatrix(I&, 4)
            CANTING1 = GRID.TextMatrix(I&, 7)
            CANTSAL1 = GRID.TextMatrix(I&, 8)
            Call AJUASEMBLI(CODINT%(Text1), CANTING1, CANTSAL1, TIMOI1$, DOPRIX1$, FEMOX1$, LOTEA1$)
            GRID.CellBackColor = RGB(255, 255, 255)
            Exit Sub
        End If
       '
       FECH% = FECHANUM(GRID.TextMatrix(I&, 0))
       If FECH% < 1 Then Exit Sub
'       FECHAX = FETEXCOR1$(FECH%)
       COMOVI$ = TRIM$(GRID.TextMatrix(I&, 1))
       DOPRI$ = TRIM$(GRID.TextMatrix(I&, 2))
       DOSEC$ = TRIM$(GRID.TextMatrix(I&, 3))
       IDLOT$ = TRIM$(GRID.TextMatrix(I&, 4))
       FECHAX# = CDbl(CVDAT(FECH%))
       CONDI$ = "INT(CDBL(FechMov)) = " & FECHAX# & " AND CODIMOVI = '" & COMOVI$ & "' AND DOPRILAR = '" & DOPRI$ & "' AND DOSECLAR = '" & DOSEC$ & "' And IDENLOTE = '" & IDLOT$ & "'"

       'FECHAYHORA = VALOBADA("MOVISTO", "FEREMOVI", CONDI$, "NOMESS")
       SQLX$ = "SELECT FEREMOVI FROM MOVISTO  WHERE " & CONDI$
       Dim RS As Recordset
       Set RS = BaDaMoviExt.OpenRecordset(SQLX$, dbOpenDynaset)
       FECHAYHORA = SAFETEXT$(RS!FEREMOVI)
       MsgBox "Fecha y Hora Real del Movimiento  " & vbCrLf & vbCrLf & FECHAYHORA, , "FLEXOFT"
       GRID.CellBackColor = RGB(255, 255, 255)
   End If
   
End Sub

Sub AJUASEMBLI(CIXI%, CANTING, CANTSAL, TIMOI$, DOPRIX$, FEMOX$, LOTEA$)
'         TTXX$ = List1.TEXT
         CANTING = XVALO(CANTING)
         CANTSAL = XVALO(CANTSAL)
         '
         If CIXI% < 1 Then Exit Sub
         '
         If CANTSAL > 0 Or TIMOI$ = "AJ" Or TIMOI$ = "TR" Then
            Call ARMALILOT(CIXI%)
            Call FRESHECHO("!LILOTSE")
            Screen.MousePointer = 1
            If ULTREG("!LILOTSE") < 1 Then
              Call MENSERR(24, "No Hay Lotes Disponibles\para el Ajuste Propuesto.")
              Exit Sub
            End If
            '
            Call SELECHO("!LILOTSE")
            VDEVU$ = VALACT1$("!LILOTSE")
            VDEV2$ = VALACT2$("!LILOTSE")
            If TRIM$(VDEVU$) = "" Then Exit Sub
            '
            CAMAXI = XVALO(Mid$(VDEV2$, 21, 12))
            If CANTSAL > CAMAXI Then
               Call MENSERR(24, "Cantidad del Movimiento Supera\Saldo Fisico del Lote")
               Exit Sub
            End If
            '
            FEMOI% = FECHANUM(FEMOX$)
            TTZZ$ = "Aprueba Aplicar el Lote '" + VDEVU$ + "'\al Comprobante '" + TIMOI$ + "' - '" + DOPRIX$ + "'"
            TTZZ$ = TTZZ$ + "\de Fecha '" + FEMOX$ + "' ?\\No, Cancelar\Si, Aplicar"
            If COMALTER%(TTZZ$) = 2 Then
               SQLX$ = "Cod_Inte = " & CIXI%
               SQLX$ = SQLX$ + " AND CodiMovi = '" & TIMOI$ & "'"
               SQLX$ = SQLX$ + " AND DoPrilar = '" & DOPRIX$ & "'"
               SQLX$ = SQLX$ + " AND FechMov= '" & FETEXCOR1$(FEMOI%) & "'"
               If CANTSAL >= 0.05 Then SQLX$ = SQLX$ + " AND CantSalid = " & TRIM$(FORMATNUM(CANTSAL, "F12.4"))
               SQLX$ = SQLX$ + " AND IdenLote = '" & LOTEA$ & "' "
               '
               Screen.MousePointer = 11
               Call ACTUREGISTRO("MOVISTO", "IdenLote", SQLX$, VDEVU$, REGAF&, "NOMESS")
               If REGAF& > 0 Then
                    Call ACSALOTE(CIXI%, VDEVU$, "NOMESS")
                    Call GENKARDEX
               Else
                    Call MENSERR(24, "Actualización Imposible")
               End If
               Screen.MousePointer = 1
            End If
         End If
End Sub
Private Sub Image1_Click()
   IDLOT$ = TRIM$(Combo1.TEXT)
   If IDLOT$ <> "" Then
     On Error Resume Next
     CONSUMLOT07.Label4 = Text1
     If Err Then Exit Sub
     On Error GoTo 0
     CONSUMLOT07.Label1 = IDLOT$
     CONSUMLOT07.Show 1
   End If
End Sub


Private Sub List1_DblClick()
   TTXX$ = List1.TEXT
   CANTING = XVALO(Mid$(TTXX$, 75, 10))
   CANTSAL = XVALO(Mid$(TTXX$, 85, 10))
   '
   If CANTING > 0 Then
       IDLOT$ = TRIM$(Mid$(TTXX$, 129, 16))
       If IDLOT$ <> "" Then
         On Error Resume Next
         CONSUMLOT07.Label4 = Text1
         If Err Then Exit Sub
         On Error GoTo 0
         CONSUMLOT07.Label1 = IDLOT$
         CONSUMLOT07.Show 1
       End If
     ElseIf CANTSAL > 0 Then
       On Error Resume Next
       LOCONSUM09.Label4 = TRIM$(Mid$(TTXX$, 15, 14))
       If Err Then Exit Sub
       On Error GoTo 0
       LOCONSUM09.Label5 = Mid$(TTXX$, 1, 8)
       LOCONSUM09.Label6 = TRIM$(Mid$(TTXX$, 43, 26))
       LOCONSUM09.Show 1
   End If
   '
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      '
      LINEA$ = List1.LIST(List1.ListIndex)
      FECHA1$ = Mid$(LINEA$, 1, 8)
      FF% = FECHANUM(FECHA1$)
      FECHA2$ = FETEXCOR1$(FF%)
      LOTE$ = Mid(LINEA$, 15, 12)
      '
      CONDI$ = "DOPRICOR ='" & TRIM$(LOTE$) & "' AND FECHMOV='" & FECHA2$ & "'"
      '
      CODUSU% = XVALO(VALOBADA("MOVISTO", "NUMUSUAR", CONDI$))
      USU$ = USERNAM$(CODUSU%)
      FechaReal$ = Mid(SAFETEXT$(VALOBADA("MOVISTO", "FEREMOVI", CONDI$)), 1, 10)
      HORA$ = Mid(SAFETEXT$(VALOBADA("MOVISTO", "FEREMOVI", CONDI$)), 12, 8)
      '
      Call REPLA(VDEF$, USU$, 1, 20)
      Call REPLA(VDEF$, FechaReal$, 21, 10)
      Call REPLA(VDEF$, HORA$, 31, 8)
      DATOS$ = OBJPLA("DATOSMOV", VDEF$)
   End If
   '
End Sub


Private Sub Label8_Change()
   Call GENKARDEX
End Sub

Private Sub Label8_DblClick()
   Label8 = ""
   Label9 = ""
End Sub

Private Sub mnuDepMov_Click()
    Label8.Visible = True
    Call DepurarMovistoAmbest
    Label8.Visible = False
    
End Sub

Private Sub Text1_Change()
   '
   CI1% = CODINT%(Text1.TEXT)
   Label6.Caption = DESCRIT0$(CI1%)
   Label5.Caption = UNIMED$(CI1%)
   Call FRESHALL
   '
   Call GENKARDEX
End Sub

Private Sub Text1_DblClick()
   Call BOTSEL_Click(3)
End Sub

Private Sub Text11_Change()
   'List1.Clear
   If FECHANUM(Text11) <> 0 Then
     Call GENKARDEX
   End If
End Sub

Private Sub Text11_DblClick()
   Call BOTSEL_Click(0)
End Sub

Private Sub Text12_Change()
   'List1.Clear
   If FECHANUM(Text12) <> 0 Then
     Call GENKARDEX
   End If
End Sub

Private Sub Text12_DblClick()
   Call BOTSEL_Click(0)
End Sub

Private Sub Text3_Change()

End Sub

Private Sub Text4_Change()
   'List1.Clear
   If Val(Text4) < 256 Then
     If Val(Text4) < 1 Or ECOARCH$("TADEPOSI", Chr$(Val(Text4))) = "" Then
         Label4 = "Todos los Depósitos"
       Else
         Label4 = ECOARCH$("TADEPOSI", Chr$(Val(Text4)))
     End If
     Call GENKARDEX
   End If
End Sub

Private Sub Text4_DblClick()
   Call BOTSEL_Click(2)
End Sub

Sub GENKARDEX()
   '
Call GENKARDEX2
Exit Sub
   Static DEPXX%, PRIDEP%
   If PRIDEP% < 1 Then
     Call CARDEPOS
     PRIDEP% = COSU%(1)
   End If
   '
   Combo1.Clear
   VECOMBO$ = Space$(1024)
   PINI& = 1
   '
   LU$ = LUDAT$
   If EXISTE%(LU$ + "COKARDEX.RFS") < 1 Then
     Call COPYSTRU("MOVISTO", "COKARDEX")
   End If
   REBLA$ = REGBLAN$("COKARDEX")
   '
   'List1.Clear
   CODL$ = Text1
5  CIXX% = CODINT%(CODL$)
   If CIXX% < 1 Then Exit Sub
   '
   DEPXX% = Val(Text4)
   If DEPXX% < 0 Then DEPXX% = 0
   '
   DESFE% = FECHANUM(Text11)
   desde = CDbl(CVDAT(DESFE%))
   HASFE% = FECHANUM(Text12)
   Hasta = CDbl(CVDAT(HASFE%))
   If desde > Hasta Then
     Call MENSERR(24, "Rango de Fechas no Válido")
     Exit Sub
   End If
   '
   INCLUVIE% = 0
   If Check1.Value = 1 Then INCLUVIE% = 1
   '
6  CIDD% = CIXX%
   JJ& = 0
   SARRAS# = 0
   INCLUPOS% = 0
   GRID.Redraw = False
   If CIDD% > 0 Then
     Call ABRESTOCKS
'     Call ABRESTOCKVIE
     iRow& = 0: GRID.Rows = 23
     For I& = 1 To GRID.Rows - 1
        For j& = 0 To GRID.Cols - 1
           GRID.TextMatrix(I&, j&) = ""
        Next j&
     Next I&
     If INCLUVIE% = 1 Then
     '
        SQLX1$ = " SELECT FECHMOV,  CANTINGRE, CANTSALID, COD_INTE, IdenLote, CodiMovi, DoPriCor, DoSecCor, ReferCor, VALOUNIT, MoneVal, "
        SQLX1$ = SQLX1$ + " DEPOMOVI,  Nume_Clie, Nume_Prov, FeReMovi,  NumUsuar "

        SQLX1$ = SQLX1$ + " FROM MOVIVIE WITH(NOLOCK) "
        SQLX1$ = SQLX1$ + "WHERE Cod_Inte = " & CIDD% & " "
        If DEPXX% > 0 Then
          If DEPXX% = PRIDEP% Then
              SQLX1$ = SQLX1$ + "AND (DepoMovi = 1 OR DepoMovi = 0 OR DepoMovi IS NULL)"
            Else
              SQLX1$ = SQLX1$ + "AND DepoMovi = " & DEPXX% & " "
          End If
        End If
        SQLX1$ = SQLX1$ + "AND CodiMovi <> 'KU' "
        SQLX1$ = SQLX1$ + "ORDER BY FechMov , FeReMovi ASC "
        '
'        Dim MOVIVIETMP As ADODB.Recordset
'        Set MOVIVIETMP = FLEXCONN.Execute(FILTSQL$(SQLX1$))
        Set MOVIVIETMP = READSET(SQLX1$)
        With MOVIVIETMP
            Do While Not .EOF
              If CVINT(!FechMov) > HASFE% Then
                If INCLUPOS% = 0 Then
                  OPXX% = COMALTER%("Se Han Detectado Movimientos POSTERIORES\a la Fecha Límite Seleccionada.\\No Considerarlos\Incluirlos")
                  If OPXX% < 2 Then INCLUPOS% = (-1)
                  If OPXX% > 1 Then INCLUPOS% = 1
                End If
              End If
              '
              If CVINT(!FechMov) < DESFE% Then
                  SARRAS# = SARRAS# + XVALO(!CantIngre) - XVALO(!CantSalid)
                ElseIf CVINT(!FechMov) <= HASFE% Or INCLUPOS% = 1 Then
                  YY$ = REBLA$
                  Call REPLA(YY$, MKI$(CVINT(!FechMov)), 1, 2)
                  Call REPLA(YY$, MKI$(XVALO(!Cod_Inte)), 3, 2)
                  Call REPLA(YY$, SAFETEXT$(!IDENLOTE), 5, 16)
                  Call REPLA(YY$, SAFETEXT$(!CODIMOVI), 21, 2)
                  Call REPLA(YY$, SAFETEXT$(!DoPriCor), 23, 10)
                  Call REPLA(YY$, SAFETEXT$(!DoSecCor), 33, 10)
                  Call REPLA(YY$, SAFETEXT$(!ReferCor), 43, 30)
                  Call REPLA(YY$, MKD$(XVALO(!ValoUnit)), 73, 8)
                    MONEX$ = "$": If XVALO(!MoneVal) = 2 Then MONEX$ = "U$"
                  Call REPLA(YY$, MONEX$, 81, 2)
                  Call REPLA(YY$, Chr$(XVALO(!DEPOMOVI)), 83, 1)
                  Call REPLA(YY$, MKI$(XVALO(!Nume_Clie)), 84, 2)
                  Call REPLA(YY$, MKI$(XVALO(!Nume_Prov)), 86, 2)
                  Call REPLA(YY$, MKD$(XVALO(!CantIngre)), 88, 8)
                  Call REPLA(YY$, MKD$(XVALO(!CantSalid)), 96, 8)
                  Call REPLA(YY$, MKD$(0), 104, 8)
                  Call REPLA(YY$, Mid$(SAFETEXT$(!DoPriCor), 11, 4), 112, 4)
                  Call REPLA(YY$, Mid$(SAFETEXT$(!DoSecCor), 11, 4), 116, 4)
                  Call REPLA(YY$, CVINT(!FEREMOVI), 122, 2)
                  Call REPLA(YY$, Chr$(XVALO(!NumUsuar)), 126, 1)
                  JJ& = JJ& + 1
                  Call GRAREG("#COKARDEX", YY$, JJ&)
              End If
            .MoveNext
            Loop
        End With
        MOVIVIETMP.Close
        Set MOVIVIETMP = Nothing
     End If
     
     SQLX$ = " SELECT FECHMOV,  CANTINGRE, CANTSALID, COD_INTE, IdenLote, CodiMovi, DoPriCor, DoSecCor, ReferCor, VALOUNIT, MoneVal, "
     SQLX$ = SQLX$ + " DEPOMOVI,  Nume_Clie, Nume_Prov, FeReMovi,  NumUsuar "
     SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIDD% & " AND (VERIFICADO <> -1 OR VERIFICADO IS NULL) "
     If DEPXX% > 0 Then
       If DEPXX% = PRIDEP% Then
           SQLX$ = SQLX$ + "AND (DepoMovi = 1 OR DepoMovi = 0 OR DepoMovi IS NULL)"
         Else
           SQLX$ = SQLX$ + "AND DepoMovi = " & DEPXX% & " "
       End If
     End If
     If INCLUVIE% = 1 Then
       SQLX$ = SQLX$ + "AND CodiMovi <> 'KU' "
     End If
     SQLX$ = SQLX$ + "ORDER BY FechMov , FeReMovi ASC "

     '
'     Dim MoviTemp As ADODB.Recordset
'     Set MoviTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
     Set MoviTemp = READSET(SQLX$)
     With MoviTemp
         Do While Not .EOF
           If CVINT(!FechMov) > HASFE% Then
             If INCLUPOS% = 0 Then
               OPXX% = COMALTER%("Se Han Detectado Movimientos POSTERIORES\a la Fecha Límite Seleccionada.\\No Considerarlos\Incluirlos")
               If OPXX% < 2 Then INCLUPOS% = (-1)
               If OPXX% > 1 Then INCLUPOS% = 1
             End If
           End If
           '
           If CVINT(!FechMov) < DESFE% Then
               SARRAS# = SARRAS# + XVALO(!CantIngre) - XVALO(!CantSalid)
             ElseIf CVINT(!FechMov) <= HASFE% Or INCLUPOS% = 1 Then
               YY$ = REBLA$
               Call REPLA(YY$, MKI$(CVINT(!FechMov)), 1, 2)
               Call REPLA(YY$, MKI$(XVALO(!Cod_Inte)), 3, 2)
               Call REPLA(YY$, SAFETEXT$(!IDENLOTE), 5, 16)
               Call REPLA(YY$, SAFETEXT$(!CODIMOVI), 21, 2)
               Call REPLA(YY$, SAFETEXT$(!DoPriCor), 23, 10)
               Call REPLA(YY$, SAFETEXT$(!DoSecCor), 33, 10)
               Call REPLA(YY$, SAFETEXT$(!ReferCor), 43, 30)
               Call REPLA(YY$, MKD$(XVALO(!ValoUnit)), 73, 8)
                 MONEX$ = "$": If XVALO(!MoneVal) = 2 Then MONEX$ = "U$"
               Call REPLA(YY$, MONEX$, 81, 2)
               Call REPLA(YY$, Chr$(XVALO(!DEPOMOVI)), 83, 1)
               Call REPLA(YY$, MKI$(XVALO(!Nume_Clie)), 84, 2)
               Call REPLA(YY$, MKI$(XVALO(!Nume_Prov)), 86, 2)
               Call REPLA(YY$, MKD$(XVALO(!CantIngre)), 88, 8)
               Call REPLA(YY$, MKD$(XVALO(!CantSalid)), 96, 8)
               Call REPLA(YY$, MKD$(0), 104, 8)
               Call REPLA(YY$, Mid$(SAFETEXT$(!DoPriCor), 11, 4), 112, 4)
               Call REPLA(YY$, Mid$(SAFETEXT$(!DoSecCor), 11, 4), 116, 4)
               Call REPLA(YY$, CVINT(!FEREMOVI), 122, 2)
               Call REPLA(YY$, Chr$(XVALO(!NumUsuar)), 126, 1)
               JJ& = JJ& + 1
               ' Call GRAREG("!COKARDEX", YY$, JJ&)
               Call GRAREG("#COKARDEX", YY$, JJ&)
           End If
         .MoveNext
         Loop
     End With
     MoviTemp.Close
     Set MoviTemp = Nothing
     'Call SETULTREG("!COKARDEX", JJ&)
     Call SETULTREG("#COKARDEX", JJ&)
     Screen.MousePointer = 11
     '
     If INCLUVIE% = 1 Then
       Call BLOQARCH("COKARDEX")
       'Call TRACENTRAL("COKARDEX", "")
       Call FILEDUMP("COKARDEX")
       Call FILESORT("COKARDEX", "", 1, 1, "ASC")
       'Call TRALOCAL("COKARDEX", "")
       For KKUL& = 1 To JJ&
         CKX$ = REGLEIDO$("COKARDEX", KKUL&)
         Call GRAREG("#COKARDEX", CKX$, KKUL&)
       Next KKUL&
       Call CIERRARCH("COKARDEX")
     End If
     '
     iRow& = 0: GRID.Rows = 23
     For I& = 1 To GRID.Rows - 1
        For j& = 0 To GRID.Cols - 1
           GRID.TextMatrix(I&, j&) = ""
        Next j&
     Next I&
     Combo1.TEXT = "": Call Combo1_Click
     ' FIN& = ULTREG&("!COKARDEX")
     FIN& = ULTREG&("#COKARDEX")
     CASALDO# = 0
     If Abs(SARRAS#) >= 0.05 Then
        ' TTXX$ = Space$(256)
        FEX = DIANTE(DESFE%)
        ' Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
        ' IROW& = IROW& + 1
        ' GRID.Row = IROW&
        ' GRID.Col = 0
        GRID.TextMatrix(iRow, 0) = FECHATEX$(FEX)
        REFEM$ = "Saldo Anterior al " + FECHATEX$(DESFE%)
        ' Call REPLA(TTXX$, REFEM$, 43, 36)
        ' GRID.Col = 5: GRID.TEXT = REFEM$
        GRID.TextMatrix(iRow, 5) = REFEM$
        '
        CASALDO# = SARRAS#
        ' Call REPLA(TTXX$, FORMATNUM$(CASALDO#, "F12.1"), 95, 12)
        ' Call REPLA(TTXX$, FORMATNUM$(CASALDO#, "F12." & CNTDC$(CIXX%) & ""), 95, 12)
        ' GRID.Col = 9: GRID.TEXT = TRIM$(FORMATNUM$(CASALDO#, "F12." & CNTDC$(CIXX%) & ""))
        GRID.TextMatrix(iRow, 9) = TRIM$(FORMATNUM$(CASALDO#, "F12." & CNTDC$(CIXX%) & ""))
        ' List1.AddItem TTXX$
     End If
     '
     GRID.Rows = FIN& + 1
     For KU& = 1 To FIN&
       Call TRACE(20, KU&, FIN&)
       ' YY$ = REGLEIDO$("!COKARDEX", KU&)
       YY$ = REGLEIDO$("#COKARDEX", KU&)
       ' TTXX$ = Space$(256)
       FEX = CVI(Left$(YY$, 2))
       ' Call REPLA(TTXX$, FECHATEX$(FEX), 1, 8)
       iRow& = iRow& + 1
       If iRow > GRID.Rows - 1 Then GRID.Rows = iRow& + 1
       ' GRID.Row = IROW&
       ' GRID.Col = 0:
       GRID.TextMatrix(iRow, 0) = FECHATEX$(FEX)
         CMOVI$ = Mid$(YY$, 21, 2)
       GRID.TextMatrix(iRow&, 1) = CMOVI$
         DOPRI$ = Mid$(YY$, 23, 10) + Mid$(YY$, 112, 4)
       GRID.TextMatrix(iRow&, 2) = DOPRI$
         DOSEC$ = Mid$(YY$, 33, 10) + Mid$(YY$, 116, 4)
       GRID.TextMatrix(iRow&, 3) = DOSEC$
         REFEM$ = Mid$(YY$, 43, 30)
       GRID.TextMatrix(iRow&, 5) = REFEM$
       ' Call REPLA(TTXX$, REFEM$, 43, 26)
         DEPOM$ = FORMATNUM$(Asc(Mid$(YY$, 83, 1)), "I3")
       GRID.TextMatrix(iRow&, 6) = TRIM$(DEPOM$)
         CAINGRE# = CVD(Mid$(YY$, 88, 8))
       GRID.TextMatrix(iRow&, 7) = TRIM$(FORMATNUM$(CAINGRE#, "F12." & CNTDC$(CIXX%) & ""))
         CASALI# = CVD(Mid$(YY$, 96, 8))
       GRID.TextMatrix(iRow&, 8) = TRIM$(FORMATNUM$(CASALI#, "F12." & CNTDC$(CIXX%) & ""))
         CASALDO# = CASALDO# + CAINGRE# - CASALI#
       GRID.TextMatrix(iRow&, 9) = TRIM$(FORMATNUM$(CASALDO#, "F12." & CNTDC$(CIXX%) & ""))
         IDLOT$ = Mid$(YY$, 5, 16)
       GRID.TextMatrix(iRow&, 4) = TRIM$(IDLOT$)
       '
       If TRIM$(IDLOT$) <> "" Then
          If InStr(VECOMBO$, IDLOT$) < 1 Then
             If PINI& > Len(VECOMBO$) Then VECOMBO$ = VECOMBO$ + Space$(1024)
             Mid$(VECOMBO$, PINI&, 16) = IDLOT$
             PINI& = PINI& + 16
          End If
       End If
       '
       Call CompletaCombo(IDLOT$)
'       List1.AddItem TTXX$
     Next KU&
     Call FREEMEM("COKARDEX")
     '
     GRID.Redraw = True
     
     For KU& = 1 To PINI& - 1 Step 16
        Combo1.AddItem Mid$(VECOMBO$, KU&, 16)
     Next KU&
     '
     ' ESTO ES PARA POSICIONARSE AL FINAL DE LA LISTA
     'TPXI& = List1.ListCount - 19
     'If TPXI& < 0 Then TPXI& = 0
     'List1.TopIndex = TPXI&
     '
     TPXI& = GRID.Rows - 22: If TPXI& < 1 Then TPXI& = 1
     On Error Resume Next
     GRID.TopRow = TPXI&
     On Error GoTo 0
     '
     Screen.MousePointer = 1
     '
Exit Sub
'       Call SETULTREG("COKARDEX", JJ&)
'       Call CIERRARCH("COKARDEX")
'       Screen.MousePointer = 1
'       Call CONECRUN("*COKARDEX/" + CODD$, "Ficha Kardex Indirecta") ' + "/WAIT")
'       Screen.MousePointer = 1
'    Else
'       Call MENSERR(24, "Codigo: " + TRIM$(CODD$) + "\sin Movimientos Según Condición SQL.")
   End If
   '
   Screen.MousePointer = 1
   GoTo 5
   '
End Sub

Sub CompletaCombo(IDLOT$)
    '
    If TRIM$(IDLOT$) <> "" Then
       IDLOT1$ = TRIM$(IDLOT$)
       Dim EXISTE As Boolean
       EXISTE = False
       '
       For I = 0 To Combo1.ListCount - 1
          If Combo1.LIST(I) = IDLOT1$ Then
             EXISTE = True
             Exit Sub
          End If
       Next I
       '
       If EXISTE = False Then
          Combo1.AddItem IDLOT1$
       End If
    End If
    '
End Sub

Sub GENKARDEX2()
   '
   Call ABRECONEXION
   '
   
   Static DEPXX%, PRIDEP%
   If PRIDEP% < 1 Then
      Call CARDEPOS
      PRIDEP% = COSU%(1)
   End If
   '
   Combo1.Clear
   VECOMBO$ = Space$(1024)
   PINI& = 1
   '
   LU$ = LUDAT$
   If EXISTE%(LU$ + "COKARDEX.RFS") < 1 Then
     Call COPYSTRU("MOVISTO", "COKARDEX")
   End If
   REBLA$ = REGBLAN$("COKARDEX")

'''   List1.Clear
   'PARA LIMPIAR LA GRILLA POR MAS QUE EL CODIGO NO SEA VALIDO
   
   iRow& = 0: GRID.Rows = 23
   For I& = 1 To GRID.Rows - 1
     GRID.Row = I
      For j& = 0 To GRID.Cols - 1
          GRID.COL = j
         GRID.TextMatrix(I&, j&) = ""
         GRID.CellBackColor = RGB(255, 255, 255)
         GRID.CellForeColor = RGB(0, 0, 0)
      Next j&
   Next I&
   '
   CODL$ = Text1
5  CIXX% = CODINT%(CODL$)
   If CIXX% < 1 Then Exit Sub
   '
   DEPXX% = XVALO(Text4)
   If DEPXX% < 0 Then DEPXX% = 0
   '
   DESFE% = FECHANUM(Text11): ' desde = FETEXCOR1$(DESFE%)
    desde = CDbl(CVDAT(DESFE%))
   HASFE% = FECHANUM(Text12): 'Hasta = FETEXCOR1$(HASFE%)
    Hasta = CDbl(CVDAT(HASFE%))
   If DESFE% > HASFE% Then
     Call MENSERR(24, "Rango de Fechas no Válido")
     Exit Sub
   
   End If
   '
   INCLUVIE% = 0:   If Check1.Value = 1 Then INCLUVIE% = 1
   '
6  CIDD% = CIXX%
   CONDI0$ = "Cod_Inte = " & CIDD% & " "
   If DEPXX% > 0 Then
      If DEPXX% = PRIDEP% Then
          CONDI0$ = CONDI0$ + "AND (DepoMovi = 1 OR DepoMovi = 0 OR DepoMovi IS NULL)"
        Else
          CONDI0$ = CONDI0$ + "AND DepoMovi = " & DEPXX% & " "
      End If
   End If
   '
   Screen.MousePointer = 11

   Call ABRESTOCKS
   iRow& = 0: GRID.Rows = 23
   For I& = 1 To GRID.Rows - 1
      For j& = 0 To GRID.Cols - 1
         GRID.TextMatrix(I&, j&) = ""
      Next j&
   Next I&
'   ACUIN = 0: ACUSA = 0
'   If INCLUVIE% = 1 Then
'      CONDI0$ = CONDI0$ + "AND CodiMovi <> 'KU' "
'      ACUIN = XVALO(VALOBADA("MOVIVIE", "SUM(CANTINGRE)", "FECHMOV < '" & desde & "' AND " & CONDI0$, "NOMESS"))
'      ACUSA = XVALO(VALOBADA("MOVIVIE", "SUM(CANTSALID)", "FECHMOV < '" & desde & "' AND " & CONDI0$, "NOMESS"))
'   End If
'   ACUIN = ACUIN + XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", "FECHMOV < '" & desde & "' AND " & CONDI0$, "NOMESS"))
'   ACUSA = ACUSA + XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", "FECHMOV < '" & desde & "' AND " & CONDI0$, "NOMESS"))
'   '
'
'   SARRAS# = ACUIN - ACUSA
'   If Abs(SARRAS#) >= 0.05 Then
'       iRow& = iRow& + 1
'       FEX = DIANTE(DESFE%)
'       GRID.TextMatrix(iRow&, 0) = FECHATEX$(FEX)
'       REFEM$ = "Saldo Anterior al " + FECHATEX$(DESFE%)
'       GRID.TextMatrix(iRow, 5) = REFEM$
'       CASALDO# = SARRAS#
'       GRID.TextMatrix(iRow, 9) = TRIM$(FORMATNUM$(CASALDO#, "F12." & CNTDC$(CIXX%) & ""))
'       iRow& = iRow& + 1  ' PARA QUE DEJE UNO LIBRE
'   End If

   JJ& = 0
   ' SARRAS# = 0
   INCLUPOS% = 0
'   If CIDD% > 0 Then
'     Call ABRESTOCKS
'     Call ABRESTOCKVIE
'     If INCLUVIE% = 1 Then
'     '
'        SQLX1$ = " SELECT * FROM MOVIVIE WITH(NOLOCK) "
'        SQLX1$ = SQLX1$ + "WHERE Cod_Inte = " & CIDD% & " "
'        If DEPXX% > 0 Then
'          If DEPXX% = PRIDEP% Then
'              SQLX1$ = SQLX1$ + "AND (DepoMovi = 1 OR DepoMovi = 0 OR DepoMovi IS NULL)"
'            Else
'              SQLX1$ = SQLX1$ + "AND DepoMovi = " & DEPXX% & " "
'          End If
'        End If
'        SQLX1$ = SQLX1$ + "AND CodiMovi <> 'KU' "
'        sqlx1$ = sqlx1$ + "ORDER BY FechMov , FeReMovi ASC "
        '
'   BASEAC$ = "MOVISTO": If INCLUVIE% = 1 Then BASEAC$ = "MOVIVIE"
   '
   GRID.Redraw = False
'20 sqlx1$ = "SELECT FECHMOV, CodiMovi, DOPRILAR,  DoSecLar, IdenLote, ReferCor,  DEPOMOVI,  CANTINGRE, CANTSALID "
'   sqlx1$ = sqlx1$ + " FROM " + BASEAC$
'   sqlx1$ = sqlx1$ + " WHERE " + CONDI0$
'   sqlx1$ = sqlx1$ + " WHERE FECHMOV >= '" & desde & "' "
'   sqlx1$ = sqlx1$ + " AND FECHMOV <= '" & Hasta & "' "
'   sqlx1$ = sqlx1$ + " ORDER BY FechMov , FeReMovi ASC "
'   Set MOVIVIETMP = BaDaMoviExt.OpenRecordset(sqlx1$, dbOpenSnapshot)
'
''   Dim MOVIVIETMP As ADODB.Recordset
''   Set MOVIVIETMP = FLEXCONN.Execute(FILTSQL$(SQLX1$))
'   Set MOVIVIETMP = READSET(sqlx1$)
20   SQLX$ = " SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIDD% & " "
     SQLX$ = SQLX$ + "AND Cdbl(FechMov) >= " & desde & " "
     SQLX$ = SQLX$ + "AND INT(Cdbl(FechMov)) <= " & Hasta & " "
     If DEPXX% > 0 Then
       If DEPXX% = PRIDEP% Then
           SQLX$ = SQLX$ + "AND (DepoMovi = " & DEPXX% & " OR DepoMovi = 0 OR ISNULL(DepoMovi) = True OR ISEMPTY(DepoMovi) = True) "
         Else
           SQLX$ = SQLX$ + "AND DepoMovi = " & DEPXX% & " "
       End If
     End If
     If INCLUVIE% = 1 Then
       SQLX$ = SQLX$ + "AND CodiMovi <> 'KU' "
     End If
     SQLX$ = SQLX$ + "ORDER BY FechMov ASC, FeReMovi ASC; "
     '
     Set MoviTemp = BaDaMoviExt.OpenRecordset(SQLX$, dbOpenSnapshot)
     With MoviTemp

      Do While Not .EOF
              If CVINT(!FechMov) > HASFE% Then
                If INCLUPOS% = 0 Then
                  OPXX% = COMALTER%("Se Han Detectado Movimientos POSTERIORES\a la Fecha Límite Seleccionada.\\No Considerarlos\Incluirlos")
                  If OPXX% < 2 Then INCLUPOS% = (-1)
                  If OPXX% > 1 Then INCLUPOS% = 1
                End If
              End If
         IDLOT$ = SAFETEXT$(!IDENLOTE)
         If XVALO(Label8) > 0 Then
            If PROVELOTE%(IDLOT$) <> XVALO(Label8) Then GoTo 65
         End If
             '
              
         iRow& = iRow& + 1
         If iRow& > GRID.Rows - 1 Then
            GRID.Rows = iRow& + 1
         End If
         GRID.TextMatrix(iRow, 0) = FECHATEX$(CVINT(!FechMov))
         GRID.TextMatrix(iRow&, 1) = SAFETEXT$(!CODIMOVI)
         GRID.TextMatrix(iRow&, 2) = SAFETEXT$(!DoPriLar)
         GRID.TextMatrix(iRow&, 3) = SAFETEXT$(!DoSecLar)
            IDLOT$ = SAFETEXT$(!IDENLOTE)
         GRID.TextMatrix(iRow&, 4) = IDLOT$
         GRID.TextMatrix(iRow&, 5) = SAFETEXT$(!ReferCor)
         GRID.TextMatrix(iRow&, 6) = XVALO(!DEPOMOVI)
            CAINGRE# = XVALO(!CantIngre)
            CASALI# = XVALO(!CantSalid)
            CASALDO# = CASALDO# + XVALO(!CantIngre) - XVALO(!CantSalid)
         GRID.TextMatrix(iRow&, 7) = TRIM$(FORMATNUM$(CAINGRE#, "F12." & CNTDC$(CIXX%) & ""))
         GRID.TextMatrix(iRow&, 8) = TRIM$(FORMATNUM$(CASALI#, "F12." & CNTDC$(CIXX%) & ""))
         GRID.TextMatrix(iRow&, 9) = TRIM$(FORMATNUM$(CASALDO#, "F12." & CNTDC$(CIXX%) & ""))
            Call CompletaCombo(IDLOT$)
65    .MoveNext
      Loop
   End With
   MoviTemp.Close
   Set MoviTemp = Nothing
'   If BASEAC$ = "MOVIVIE" Then
'      BASEAC$ = "MOVISTO"
'      GoTo 20
'   End If
   '
   GRID.Redraw = True
   ' ESTO ES PARA POSICIONARSE AL FINAL DE LA LISTA
   TPXI& = GRID.Rows - 22: If TPXI& < 1 Then TPXI& = 1
   On Error Resume Next
   GRID.TopRow = TPXI&
   On Error GoTo 0
   '
   Screen.MousePointer = 1
   '
End Sub



Private Sub GRID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then ' BUSQUEDA  POR TEXTO CON F3
       If ARBUS$ <> "" Then
          VUELTA% = 0
          j& = GRID.Row

          H& = 0: If GRID.COL < GRID.Cols Then H& = GRID.COL + 1
10        For U& = j& To GRID.Rows - 1
             HI& = H&: If U& > j& Then HI& = 0
             For I& = HI& To GRID.Cols - 1
              If InStr(UCase$(GRID.TextMatrix(U&, I&)), ABUSQUE$) > 0 Then
                GRID.Row = U&: GRID.COL = I&
                On Error Resume Next
                 VISIROWS = (GRID.Height / GRID.RowHeight(j&)) - 4
                If Err > 0 Then
                 VISIROWS = 1
                End If
                On Error GoTo 0
                 If VISIROWS > 0 Then
                    If GRID.TopRow < GRID.Row - VISIROWS Then
                      GRID.TopRow = GRID.Row - VISIROWS
                    End If
                 End If
                GRID.SetFocus
                Exit Sub
              End If
              Next I&
          Next U&
          '
          Call MENSERR(24, "Argumento no Encontrado.")
       End If
   End If

End Sub

