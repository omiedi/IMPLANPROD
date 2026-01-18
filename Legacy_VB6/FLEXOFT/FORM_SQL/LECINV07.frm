VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form LECINV07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lectura de Inventarios"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5205
   ClipControls    =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture13 
      BackColor       =   &H00404080&
      Height          =   6795
      Left            =   4320
      ScaleHeight     =   6735
      ScaleWidth      =   1110
      TabIndex        =   12
      Top             =   -120
      Width           =   1170
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
         Height          =   640
         Left            =   105
         Picture         =   "LECINV07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Cerrar - Salir de la Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "LECINV07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   1050
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   13
         Top             =   4200
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "LECINV07.frx":0454
         Top             =   4560
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Left            =   160
      TabIndex        =   1
      Top             =   2160
      Width           =   3975
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1440
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   2
         Top             =   1905
         Width           =   2010
         Begin VB.CommandButton Command8 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "LECINV07.frx":2376
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Mensajes de Error - Registros No Procesables (Ultima Lectura)"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Registros No Procesables"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   720
            TabIndex        =   3
            Top             =   120
            Width           =   1185
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1080
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   4
         Top             =   1395
         Width           =   2010
         Begin VB.CommandButton Command23 
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
            Left            =   0
            Picture         =   "LECINV07.frx":2680
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Listado de Códigos No Leidos con Saldo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "No Leidos con Saldo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   720
            TabIndex        =   5
            Top             =   45
            Width           =   1125
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   720
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   6
         Top             =   900
         Width           =   2010
         Begin VB.CommandButton Command15 
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
            Left            =   0
            Picture         =   "LECINV07.frx":298A
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Listado de Códigos Leidos Sin Saldo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Registros Leí- dos sin Saldo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   650
            TabIndex        =   7
            Top             =   45
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   400
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   8
         Top             =   405
         Width           =   2010
         Begin VB.CommandButton Command26 
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
            Left            =   0
            Picture         =   "LECINV07.frx":2C94
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Listado de Códigos Leidos Con Saldo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Registros Leí- dos con Saldo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   650
            TabIndex        =   9
            Top             =   45
            Width           =   1305
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Inventario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Left            =   160
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   675
         ScaleHeight     =   585
         ScaleWidth      =   2670
         TabIndex        =   18
         Top             =   1080
         Width           =   2730
         Begin VB.CommandButton Command5 
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
            Left            =   0
            Picture         =   "LECINV07.frx":2F9E
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Lectura de Archivo Generado con Colectora de Datos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Lectura de Inventario (Colectora de Datos)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   720
            TabIndex        =   19
            Top             =   105
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   675
         ScaleHeight     =   585
         ScaleWidth      =   2670
         TabIndex        =   10
         Top             =   375
         Width           =   2730
         Begin VB.CommandButton Command3 
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
            Left            =   0
            Picture         =   "LECINV07.frx":32A8
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Inicialización de Inventario"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Inicialización de Ar- chivos de Inventario"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   720
            TabIndex        =   11
            Top             =   120
            Width           =   1935
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11025
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "LECINV07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   Screen.MousePointer = 11
   Call FINAL("?MARSINSAL") ' MARCADOS SIN SALDOS
   Screen.MousePointer = 1
   Command15.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Screen.MousePointer = 11
    '
    If ULTREG&("INVMENSA") > 0 Then ' LISTADO DE ERRORES
        Call FINAL("*INVMENSA")
    Else
        Call COMUNI("No se han Encontrado Errores en la Lectura del Archivo")
    End If
    '
    Screen.MousePointer = 1
    Command8.Enabled = True
End Sub

Private Sub Command23_Click()
   Command23.Enabled = False
   Screen.MousePointer = 11
   Call FINAL("?NOMARCNSAL") ' LISTADO NO MARCADOS CON SALDO
   Screen.MousePointer = 1
   Command23.Enabled = True
End Sub

Private Sub Command26_Click()
  Command26.Enabled = False
  Screen.MousePointer = 11
  Call FINAL("?MARCONSAL") 'MARCADOS CON SALDO
  Screen.MousePointer = 1
  Command26.Enabled = True
  
End Sub

Private Sub Command3_Click()
   'INICIALIZACION DE INVENTARIO (PONER EN 0 CAMPO MARINVE TODOS LOS CAMPOS)
    Command3.Enabled = False
    Screen.MousePointer = 11
    '
    Call APERBASDAT("STOCKS")
    CONDI$ = "COD_INTE > 0 "
    Call ACTUREGISTRO("LOTINGRE", "MARINVE", CONDI$, 0, REGAF&, "NOMESS")
    Screen.MousePointer = 1
    Call SETULTREG("INVMENSA", 0)
    Call COMUNI("Proceso Terminado")
    '
    Command3.Enabled = True
End Sub

Private Sub Command5_Click()

   'LECTURA DEL ARCHIVO CARGADO CON LA LECTORA
    
    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "CAPTURA DE INVENTARIOS"
    CommonDialog1.FILTER = "Archivo de Texto(*.TXT)|*.TXT|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
    '
    DoEvents
    Screen.MousePointer = 11
    ARCH$ = UCase$(CommonDialog1.FileName)
    '
    NARCHIVO% = FILEOPEN%(ARCH$, 1, 128) 'ABRE EL ARCHIVO DE LECTURA
        '
    Call SETULTREG("INVMENSA", 0)
    U& = 0
    Do While Not EOF(NARCHIVO%) ' RECORRE
          Line Input #NARCHIVO%, TXTO$ ' LEE CADA REGISTRO
          MENSA$ = ""
          If TRIM$(TXTO$) = "" Then GoTo 50 ' VALIDO QUE EL REGISTRO NO ESTE VACIO
          '
          CODIEXT$ = TRIM$(Left$(TXTO$, 16))  '        CODIGO EXTERNO
          LOTEX$ = Mid$(TXTO$, 17, 12) '        LOTE
          CI% = CODINT(CODIEXT$) '              CODIGO INTERNO
          '
          If CI% < 1 Then MENSA$ = "Codigo: " & CODIEXT$ & " Inexistente o Dado de Baja": GoTo 50
          '
          CONDI$ = "Cod_Inte = " & CI% & "  AND IdenLote = '" & LOTEX$ & "'"
          CANRE& = CantRegistros&("LOTINGRE", CONDI$)
          If CANRE& > 0 Then
             Call ACTUREGISTRO("LOTINGRE", "MARINVE", CONDI$, 1, REGAF&, "NOMESS")
          Else
            MENSA$ = "Lote: " & LOTEX$ & " No Válido Para Código: " & CODIEXT$
          End If
50        If TRIM$(MENSA$) <> "" Then
            U& = U& + 1
            Call REPLA(TX$, MENSA$, 1, 96) ' GRABA MENSAJE DE ERROR
            Call GRAREG("INVMENSA", TX$, U&)
          End If
    Loop
    Call SETULTREG("INVMENSA", U&)
    Close NARCHIVO%
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    Call APARTIMA ' ACTUALIZACION DE MASTER A LOTINGRE DE TIPO DE MATERIAL
    '
    Call COMUNI("Lectura Realizada. Debe Reiniciar el Módulo")
    Call FINAL("") 'Nota: Se  realiza un call final del formulario luego de realizar
    'la lectura debido a que cuando se utiliza el CommonDialog cambia el path
    ' y al intentar ejecutar el listado de errores (Dat)devolvia un mensaje
    ' que no encontraba el archivo MSTDRV04.EXE (FALTA UTILITARIO MSTDRV04.EXE)
99
   
End Sub

