VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form SITORFAB09 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Situación de Stocks Por Orden de Fabricación  o Simulación por Código y Cantidad."
   ClientHeight    =   9135
   ClientLeft      =   75
   ClientTop       =   285
   ClientWidth     =   19140
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9135
   ScaleWidth      =   19140
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9000
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   19095
      Begin VB.Frame Frame5 
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
         Height          =   1065
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   13020
         Begin VB.CommandButton Command1 
            Caption         =   "Selec.Manual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   24
            Top             =   285
            Width           =   1500
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FBCA9D&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   "Cant.Proyectada"
            Top             =   270
            Width           =   1500
         End
         Begin VB.TextBox TXTCANTPROY 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   540
            Width           =   1500
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BackColor       =   &H00FBCA9D&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9500
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "Cant.Realizada"
            Top             =   270
            Width           =   1500
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FBCA9D&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   10980
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "Saldo a Cerrar"
            Top             =   270
            Width           =   1380
         End
         Begin VB.TextBox TXTSALDOOF 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   10980
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   540
            Width           =   1380
         End
         Begin VB.TextBox TXTCANTREALI 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   9500
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   540
            Width           =   1500
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   12360
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   540
            Width           =   420
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FBCA9D&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   12360
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   "Niv."
            Top             =   270
            Width           =   420
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FBCA9D&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "Descripción"
            Top             =   270
            Width           =   4100
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   540
            Width           =   4100
         End
         Begin VB.TextBox Text10 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   540
            Width           =   1630
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   540
            Width           =   2235
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            BackColor       =   &H00FBCA9D&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Text            =   "Código"
            Top             =   270
            Width           =   2235
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H00FBCA9D&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "O.Fabricación"
            Top             =   270
            Width           =   1630
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   1065
         Left            =   12840
         TabIndex        =   5
         Top             =   0
         Width           =   6200
         Begin VB.CommandButton Command4 
            Caption         =   "Multinivel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   3240
            TabIndex        =   27
            Top             =   600
            Width           =   1620
         End
         Begin VB.CommandButton CMDOF 
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
            Height          =   765
            Left            =   5100
            Picture         =   "SITORFAB09.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Ordenes de  Fabricación"
            Top             =   120
            Width           =   850
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Exportar a Excel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   3240
            TabIndex        =   25
            Top             =   160
            Width           =   1620
         End
         Begin VB.CommandButton Command2 
            Caption         =   "O.F."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   120
            TabIndex        =   8
            Top             =   150
            Width           =   850
         End
         Begin VB.CommandButton Command41 
            Height          =   765
            Left            =   1140
            Picture         =   "SITORFAB09.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Sube al Nivel Anterior"
            Top             =   150
            Width           =   850
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
            Height          =   765
            Left            =   2160
            Picture         =   "SITORFAB09.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Imprime Grilla Activa"
            Top             =   150
            Width           =   850
         End
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   4920
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   9720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "SITORFAB09.frx":0C7E
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "SITORFAB09.frx":0EB2
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   7980
      Left            =   60
      TabIndex        =   3
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar"
      Top             =   1080
      Width           =   19035
      _ExtentX        =   33576
      _ExtentY        =   14076
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
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
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Menu AGREGAR_RUTA 
      Caption         =   "Agregar Ruta"
      Visible         =   0   'False
      Begin VB.Menu mnuedi 
         Caption         =   "Editar"
      End
   End
End
Attribute VB_Name = "SITORFAB09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private objexcel As Excel.Application
Private Libro As Excel.Workbook
Private Hoja As Excel.Worksheet
Private Declare Function ShellExecute Lib "shell32.dll" _
                                    Alias "ShellExecuteA" _
                                    (ByVal hwnd As Long, _
                                    ByVal lpOperation As String, _
                                    ByVal lpFile As String, _
                                    ByVal lpParameters As String, _
                                    ByVal lpDirectory As String, _
                                    ByVal nShowCmd As Long) _
                                    As Long
Public CAMBIO%
Public ID_LOTENIVELANTERIOR%
Dim LOTENIVELANTERIOR$(15)



Sub CARGA_GRILLA_CARATULA()
   '
   '
   'VERIFICO SI LA ORDEN DE FABRICACION SELECCIONADA TIENE CERTIFICADOS APLICADOS
   SQLX0$ = "SELECT * FROM CERTIFICADOS WITH(NOLOCK)" '**TABLA CERTIFICADOS***QUE SE GRABA EN ABMCERTIFICADOS
   SQLX0$ = SQLX0$ + " WHERE IDENLOTE ='" & Text10 & "'"
   SQLX0$ = SQLX0$ + " AND NUMEROSERIE = '" & Text1 & "'"
   Set RST0 = READSET(SQLX0$)
   With RST0
     Do While Not .EOF
        LOTE$ = SAFETEXT$(!idenlote)
        RUT$ = SAFETEXT$(!RUT_ARCHIVO)
        'COFA$ = SAFETEXT$(!FAMILIA)
        COFA$ = TRIM$(CODFAMIL$(CODINT%(Text11)))
        FAMI$ = TRIM$(ECOARCH$("TAFAMIL", COFA$))
        If FAMI$ = "" Then GoTo 10
        TIT% = XVALO(!TIT_CERTIFICADO)
        TITU$ = ECOARCH$("TITUCERT", MKI$(TIT%))
        HOJAX% = XVALO(!HOJAS)
        If FAMI_ANT$ <> FAMI$ Then
           INDI% = INDI% + 1
           ACUMULA_HOJAX% = 0
           HOJAINI% = 0
        End If
        GoSub CARGAREGIS
10      .MoveNext
     Loop
   End With
   
   SQLX$ = " SELECT * FROM CERTI_ESTRUC WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE ORD_Fabric_Gral = '" & Text10 & "'"
   SQLX$ = SQLX$ + " AND N_Serie_Gral = '" & Text1 & "'"
   SQLX$ = SQLX$ + " order by  FAMILIA "
   
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF
        CI1% = XVALO(!CODINT)
        
        LOTE$ = SAFETEXT$(!idenlote)
        NUMSE$ = SAFETEXT$(!NUMEROSERIE)
        'COFA$ = SAFETEXT$(!FAMILIA)
        COFA$ = TRIM$(CODFAMIL$(CI1%))
        FAMI$ = TRIM$(ECOARCH$("TAFAMIL", COFA$))
        TIT% = XVALO(!TIT_CERTIFICADO)
        TITU$ = ECOARCH$("TITUCERT", MKI$(TIT%))
        HOJAX% = XVALO(!HOJAS)
        '
        If LOTE$ = "" And FAMI$ = "" Then GoTo 100
        If LOTE$ <> "" Then
           '
           CONDI2$ = "CODINT = " & CI1%
           CONDI2$ = CONDI2$ + " AND IDENLOTE = '" & LOTE$ & "'"
            If NUMSE$ <> "" Then
               CONDI2$ = CONDI2$ + " AND NUMEROSERIE = '" & NUMSE$ & "'"
            Else
               CONDI2$ = CONDI2$ + " AND (NUMEROSERIE = '' OR NUMEROSERIE IS NULL) "
            End If
            '
            If LOTE$ = LOTE2$ Then GoTo 100 'POR QUE SI ES EL MISMO LOTE DUPLICA
                '                           'INFORMACION
            sqlx2$ = "SELECT * FROM CERTIFICADOS WITH(NOLOCK)" '**TABLA CERTIFICADOS***QUE SE GRABA EN ABMCERTIFICADOS
            sqlx2$ = sqlx2$ + " WHERE " & CONDI2$
            Set Rst = READSET(sqlx2$)
            With Rst
              Do While Not .EOF
                LOTE2$ = SAFETEXT$(!idenlote)

                RUT$ = SAFETEXT$(!RUT_ARCHIVO)
                'COFA$ = SAFETEXT$(!FAMILIA)
                COFA$ = TRIM$(CODFAMIL$(CI1%))
                FAMI$ = TRIM$(ECOARCH$("TAFAMIL", COFA$))
                TIT% = XVALO(!TIT_CERTIFICADO)
                TITU$ = ECOARCH$("TITUCERT", MKI$(TIT%))
                HOJAX% = XVALO(!HOJAS)
                '
                If FAMI_ANT$ <> FAMI$ Then
                   INDI% = INDI% + 1
                   ACUMULA_HOJAX% = 0
                   HOJAINI% = 0
                End If
                '
                GoSub CARGAREGIS
                '
30              .MoveNext
              Loop
            End With
          Else
                If FAMI_ANT$ <> FAMI$ Then
                   INDI% = INDI% + 1
                   ACUMULA_HOJAX% = 0
                   HOJAINI% = 0
                End If
                GoSub CARGAREGIS
          End If
100    .MoveNext
     Loop
   End With
   On Error Resume Next
   RS.Close
   Rst.Close
   RST0.Close
   Set RST0 = Nothing
   Set RS = Nothing
   Set Rst = Nothing
   On Error GoTo 0
   
Exit Sub
    'CARGA LOS REGISTROS EN LA GRILLA
    'FAMILIA (BOMBA)
    'TITULO DE CERTIFICADO
    'HOJAS DESDE HASTA (1.1 - 1.3)
CARGAREGIS:
    If INDI% < 1 Then INDI% = 1
    If HOJAX% < 1 Then HOJAX% = 1
    ACUMULA_HOJAX% = ACUMULA_HOJAX% + HOJAX%
    HOJAINI% = HOJAINI%: If HOJAINI% < 1 Then HOJAINI% = 1
    If ACUMULA_HOJAX% < 1 Then ACUMULA_HOJAX% = 1
    HOJAFIN% = ACUMULA_HOJAX%
    INICIA$ = SAFETEXT$(INDI%) & "." & HOJAINI%
    FINALIZA$ = SAFETEXT$(INDI%) & "." & HOJAFIN%
    TX$ = INICIA$ & " - " & FINALIZA$
    '
    k& = k& + 1
    CARATULA.MSF.Rows = k& + 1
    CARATULA.MSF.TextMatrix(k&, 1) = FAMI$
    CARATULA.MSF.TextMatrix(k&, 2) = TITU$
    CARATULA.MSF.TextMatrix(k&, 3) = TX$
    FAMI_ANT$ = FAMI$
    HOJAINI% = HOJAFIN% + 1
    
Return

End Sub




Private Sub CMDOF_Click()
   If TRIM$(Text11) = "" Then Exit Sub
   CODE$ = Text11
   CAN = XVALO(TXTSALDOOF)
   NCLIE = 0
   NUMPEDI% = 0
   ORCOCLI$ = ""
   
   ACONEC1$ = Chr$(124) + CODE$ + Chr$(124) + TRIM$(Str$(CAN)) + Chr$(124) + TRIM$(Str$(NCLIE)) + Chr$(124) + TRIM$(Str$(NUMPEDI%)) + Chr$(124) + ORCOCLI$
   ACONEC1$ = REPLACAR$(ACONEC1$, "/", "ß") 'PARA EVITAR CODIGOS CON BARRA POR QUE EN  EL PASAJE LLEGA TRUNCADO POR ESTE CARACTER (JYM) (ALT-225)
   ACONEC$ = TRIM$("ORFABR07/ORFNUE" + ACONEC1$)
   'ORFABR07/ORFNUE|000004700|2|999|3697|
   Call CONECRUN(ACONEC$, "Ordenes de Fabricación")

End Sub

Private Sub Command1_Click()
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   
   If COD$ <> "" Then
     Text11 = ""
5    CANTI = XVALO(InputBox$("Cantidad a Simular", "Ingrese Cantidad a Simular", CANTI))
     If CANTI < 1 Then Text11.TEXT = "": Exit Sub
     Me.TXTCANTPROY = CANTI
     Me.TXTCANTREALI = "0"
     Me.TXTSALDOOF = CANTI
     Me.Text10 = ""
     Me.Text6 = ""
     Text11.TEXT = COD$
     CMDOF.Enabled = True
   End If

End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call SELECHO("ECORDEP")
    NORFA$ = TRIM$(VALACT1$("ECORDEP"))
    If NORFA$ = "" Then GoTo 99
    '
    Screen.MousePointer = 11
    '
    CMDOF.Enabled = False
    CI1% = CODIFAB%(NORFA$)
    '
    ID_LOTENIVELANTERIOR% = 0
    LOTENIVELANTERIOR$(ID_LOTENIVELANTERIOR%) = NORFA$
'    NUSERIE$ = TRIM$(Left$(VALACT2$("INDIFANU"), 16))
'    Text1 = NUSERIE$
    Text10 = NORFA$
    Text6 = 0
    '
    CANPROYE& = XVALO(VALOBADA("ORDEFABR", "CanProy", "IdSubOr = '" & NORFA$ & "'")): Me.TXTCANTPROY = FORMATNUM$(CANPROYE&, "f9.1")
    CANREALE& = XVALO(VALOBADA("ORDEFABR", "CanApro", "IdSubOr = '" & NORFA$ & "'")): Me.TXTCANTREALI = FORMATNUM$(CANREALE&, "f9.1")
    SALDOOF = CANPROYE& - CANREALE&: If SALDOOF < 0 Then SALDOOF = 0
    Me.TXTSALDOOF = FORMATNUM$(SALDOOF, "f9.1")
    
    Text11 = CODEXT(CI1%)

99  Command2.Enabled = True
End Sub

Private Sub Command4_Click()
 
   Command4.Enabled = False
   CODD$ = Text11.TEXT
   CI% = CODINT%(CODD$)
   If CI% < 1 Then Exit Sub
   
   Campos$ = "CodXEelem/A16;Descripción/A52;U.M./A5;Uso/F10.4"
   Call CARGA_ENCABEZADO(MSF, Campos$, SITORFAB09)


   CICON% = CI%
   '
   Screen.MousePointer = 11
   Call EXPLOMUL(CI%)
   If CONTAMUL% < 2 Then Exit Sub
   '
   IMPLOSI07.Top = Top + (Height - ScaleHeight - 50) + SSTab1.Top + MARCOFICHA(2).Top + Picture5.Top
   IMPLOSI07.Left = Left + (Width - ScaleWidth) / 2 + SSTab1.Left + MARCOFICHA(2).Left + Picture5.Left
   IMPLOSI07.List1.Clear
   IMPLOSI07.Caption = "Explosión Multinivel " + DESCRIT$(CI%)
   IMPLOSI07.Show
   '
   COEFCA = 1: If PORMILLAR% = 1 Then COEFCA = 1000
   FIN& = CONTAMUL
   For KKI% = 2 To CONTAMUL
     Z$ = Space$(64)
     RESUL = NIVE%(KKI%)
     SP$ = Space$(2 * (NIVE%(KKI%) - 1))
     SP$ = String$(RESUL, ">")
     MSF.Rows = KKI% + 1
     MSF.TextMatrix(KKI%, 1) = CODEXT$(CIK%(KKI%))
     MSF.TextMatrix(KKI%, 2) = SP$ & DESCRIT0$(CIK%(KKI%))
     MSF.TextMatrix(KKI%, 3) = COEFCA * CANT(KKI%)
   Next KKI%
   Screen.MousePointer = 1
'   Set RS = READSET(SQLX$)
'   With RS
'      Do While Not .EOF
'         j& = j& + 1
'         MSF.Rows = j& + 1
'         Conjunto = XVALO(!Conjunto)
'         Subconjunto = XVALO(!Subconjunto)
'         NIVEL = XVALO(!NIVEL)
'         MSF.TextMatrix(j&, 1) = Conjunto
'         MSF.TextMatrix(j&, 2) = Subconjunto
'         MSF.TextMatrix(j&, 3) = SAFETEXT(!XConjunto)
'         MSF.TextMatrix(j&, 4) = SAFETEXT(!XSubconjunto)
'         MSF.TextMatrix(j&, 5) = XVALO(!xUsoBruto)
'         MSF.TextMatrix(j&, 6) = NIVEL
'         MSF.TextMatrix(j&, 7) = SAFETEXT(!UMED_FORM)
'       .MoveNext
'      Loop

99 Command4.Enabled = True
End Sub



Function MoverYEliminarFila(ByRef grid As msFlexGrid, ByVal origen As Integer, ByVal destino As Integer)
    Dim col As Integer
    Dim totalCols As Integer
    totalCols = grid.Cols - 1  ' Obtener el número total de columnas
    
    ' Insertar la fila debajo de la fila indicada
    grid.Rows = grid.Rows + 1
    For col = 0 To totalCols
        grid.TextMatrix(grid.Rows - 1, col) = grid.TextMatrix(origen, col)
    Next col
    
    ' Eliminar la fila de su lugar original
    If destino < origen Then
        For col = 0 To totalCols
            grid.TextMatrix(origen, col) = grid.TextMatrix(origen + 1, col)
        Next col
    Else
        For col = 0 To totalCols
            grid.TextMatrix(origen, col) = grid.TextMatrix(origen - 1, col)
        Next col
    End If
    grid.Rows = grid.Rows - 1
End Function

Private Sub Command41_Click()
  Command41.Enabled = False
  
  If IUCOD% > 0 Then

    ID_LOTENIVELANTERIOR% = ID_LOTENIVELANTERIOR% - 1
    IUCOD% = IUCOD% - 1
    Text6 = SAFETEXT$(XVALO(Text6) - 1)

    Text11.TEXT = UCOD$(IUCOD% + 1)
  End If
  Command41.Enabled = True
End Sub


Private Sub Command5_Click()
End Sub


Private Sub Command8_Click()
   Command8.Enabled = False
   '
   ENCA$ = "Código/A16;Descripción/A32;U.M./A3;Cant./F11.1;Dep.Centr./f11.1;Otros Disp./f21.1;Mat.Res./f11.1;Otr. Reser./f12.1;"
   ENCA$ = ENCA$ + "Disponib./f11.1;Faltan/f11.1;Compras/f11.1;Fab.Pend./f11.1;Dispondrá/F11.1;Faltarán/f11.1"
   Call ENCABEZADO_GENREP(SHOWREP07.grid, ENCA$, SHOWREP07)
   Call CopiarMsfAOtroMsf(MSF, SHOWREP07.grid)
   SHOWREP07.Label3 = SITORFAB09.Caption
   SHOWREP07.Label7 = MSF.Rows - 1
   'SHOWREP07.Show 1
   Call SHOWREP07.IMPRIMIR
   '
99 Command8.Enabled = True

End Sub

Private Sub Form_Load()
   Call VERJOBID(OKEY%)

   If OKEY% < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
10   Caption = Caption + "  -  " + EPAC$
     Call GRATITFOR(Caption)
   End If
   Openforms = DoEvents
   ENCA$ = "Código/A16;Descripción/A32;U.M./A3;Cant./F11.1;Dep.Centr./f11.1;Otros Depos./f14.1;Mat.Res./f11.1;Otr. Reser./f13.1;"
   ENCA$ = ENCA$ + "Total Faltan/f12;T/f0;Compras/f11.1;Fab.Pend./f11.1;Dispondrá/F11.1;Faltarán/f11.1"
   Call CARGA_ENCABEZADO(MSF, ENCA$, SITORFAB09)
   'Call CARGA_ENCABEZADO(MSF2, ENCA$, SITORFAB09)
   MSF.MergeCells = flexMergeRestrictColumns
   MSF.MergeCol(1) = True
   Call ABRECONEXION
   DEPO% = XVALO(Control$("", "DEPOCEN"))
   If DEPO% < 1 Then Call COMUNI("Falta Definir Depósito Central")
   
   CAMBIO% = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  End: Close
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    MSF.Height = Me.Height - 1500
    MSF.Width = Me.Width - 250
    If MSF.Visible = True Then
        MSF.Height = Me.Height - 1700
        MSF.Width = Me.Width - 250
        
    End If
    On Error GoTo 0
End Sub

Private Sub MSF_DblClick()

   If IUCOD% < 128 Then
     If TRIM$(Text11) <> "" Then
       If Text11 <> UCOD$(IUCOD%) Then
         IUCOD% = IUCOD% + 1: UCOD$(IUCOD%) = Text11.TEXT
                              ICOMPO%(IUCOD%) = MSF.Row
       End If
     End If
   End If
   '
   'PONE EL CODIGO EN LA CAJA DE TEXTO  Y SE EJECUTA EL EVENTO CHANGE
   COD$ = TRIM$(MSF.TextMatrix(MSF.Row, 1))
   If COD$ <> "" Then
     ID_LOTENIVELANTERIOR% = ID_LOTENIVELANTERIOR% + 1
     LOTENIVELANTERIOR$(ID_LOTENIVELANTERIOR%) = TRIM$(MSF.TextMatrix(MSF.Row, 5))
     Text11.TEXT = COD$
   End If
   '
   'SUBE UN NIVEL
   Text6 = SAFETEXT$(XVALO(Text6) + 1)
   '
   'SI NO TIENE ESTRUCTURA PRESENTAMENSAJE Y VUELVE PARA ATRAS
   If CAIT% < 1 Then
     Call COMUNI("Código Seleccionado No Posee Estructura")
     'CODIGONIVELANTERIOR% = CODINT%(UCOD$(IUCOD%))
     Call Command41_Click 'vuelve para atras
   End If
   Call COLOREAR_GRILLASQL(SITORFAB09.MSF, &HFBCA9D)

End Sub

Private Sub MSF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
      strContenidoCelda$ = MSF.TextMatrix(MSF.MouseRow, MSF.MouseCol)
      If MsgBox(strContenidoCelda$, vbYesNo, "Copiar Contenido de Celda Seleccionada") = vbYes Then
       Clipboard.Clear
       Clipboard.SetText strContenidoCelda
      End If
    End If

'  If Button = 2 Then
'    '
'    I& = MSF.MouseRow
'    j& = MSF.MouseCol
'    If I& = 0 Or I& > MSF.Rows Then Exit Sub
'    '
'    On Error Resume Next
'    MSF.Row = I&
'    MSF.COL = j&
'    On Error Resume Next
'
'    COD$ = TRIM$(MSF.TextMatrix(I&, 1))
'    If COD$ = "" Then
'        Exit Sub
'    End If
'  End If
  
'End Sub

End Sub

Sub INVICTRL()
  '
  AGREGAR_RUTA.Visible = Not (AGREGAR_RUTA.Visible)
  '
End Sub




Private Sub Text10_Change()
    Call Text11_Change
End Sub

Private Sub Text11_Change()
      '
      CODD$ = Text11.TEXT                       ' CODIGO DEL CONJUNTO
      CI% = CODINT%(CODD$)                      ' CODIGO INTERNO DEL CONJUNTO
      NORFA$ = TRIM$(Text10)                    ' NUMERO DE ORDEN DE FABRICACION
      
'      If CI% < 1 Or NORFA$ = "" Then Exit Sub
      If CI% < 1 Then Exit Sub
      Text4 = DESCRIT0$(CI%)
      Screen.MousePointer = 11
      
      PORMILLAR% = 0
      ' LEE LA ESTRUCTURA DEL PRODUCTO
      Call LEEEXPLO(CODD$, 1, 1)
      '
      MSF.Rows = 1
      MSF.BackColor = vbWhite
      '
      CODI_SUP$ = CODD$
      COSUPE% = CI%                     ' CODIGO INTERNO DEL CONJUNTO
      NIVEX% = XVALO(Text6)             ' NIVEL
      i% = 0
      
      ' RECORRE LA ESTRUCTURA DEL PRODUCTO
      For j% = 1 To CAIT%
            'BUSCO RUTA EN TABLA CERTIFICADOS
            CODI$ = TRIM$(Mid$(EXPLOLIN$(j%), 1, 16))
            CI1% = CODINT%(CODI$)
            '
            ' SI FALTA COMPLETAR CON MAS REGISTROS CON CANTIDAD UNITARIA DEL COMPONENTE ACTUAL
            ' SE UTILIZA EL CODIGO DENTRO DEL IF
            i% = i% + 1
            MSF.Rows = i% + 1
            MSF.TextMatrix(i%, 1) = Mid$(EXPLOLIN$(j%), 1, 16)   'codigo
            MSF.TextMatrix(i%, 2) = Mid$(EXPLOLIN$(j%), 17, 28)  'descripcion
            MSF.TextMatrix(i%, 3) = Mid$(EXPLOLIN$(j%), 47, 2)   'u.m
            'MSF.TextMatrix(I%, 4) = Mid$(EXPLOLIN$(j%), 49, 12)  'uso
             'CANTI = SaldoOFab(NORFA$)
             CANTI = XVALO(TXTSALDOOF)
             CANTI = CANTI * XVALO(Mid$(EXPLOLIN$(j%), 49, 12))
             MSF.TextMatrix(i%, 4) = FORMATNUM$(CANTI, "F9.1") 'uso * CANTIDD
              StockTodosLosDepositos = STOCKACT(CI1%): If StockTodosLosDepositos <= 0 Then StockTodosLosDepositos = 0
              DEPO% = XVALO(Control$("", "DEPOCEN"))
              StockDepositoCentral = STODEPOS(CI1%, DEPO%): If StockDepositoCentral <= 0 Then StockDepositoCentral = 0
              
            MSF.TextMatrix(i%, 5) = FORMATNUM$(StockDepositoCentral, "F9.1") 'DEPOSITO CENTRAL
            MSF.TextMatrix(i%, 6) = FORMATNUM$(StockTodosLosDepositos - StockDepositoCentral, "F9.1") ' resto Depositos disponibles
             StResEstaOF = STOCKRES_POR_OFABRICACION(NORFA$, CI1%)
            MSF.TextMatrix(i%, 7) = FORMATNUM$(StResEstaOF, "F9.1") 'Mat.Res.PARA ESTA ORDEN
             StResTotal = STOCKRES(CI1%)
             StResResto = FORMATNUM$(StResTotal - StResEstaOF, "F9.1"): If StResResto < 0.05 Then StResResto = 0
            MSF.TextMatrix(i%, 8) = StResResto 'Otr. Reser
            DISPONIB = CANTI - (StockTodosLosDepositos + StockDepositoCentral)
            DISPONIB = DISPONIB + StockDepositoCentral + XVALO(MSF.TextMatrix(i%, 8))
            'DISPONIBLE SE CAMBIA PARA QUE APAREZCA COMO EL TITULO FALTAN
            If DISPONIB < 0 Then DISPONIB = 0
            MSF.TextMatrix(i%, 9) = FORMATNUM$(DISPONIB, "F9.1") 'Disponib
'              FALTAN = DISPONIB - CANTI
'              If FALTAN < 0 Then
'                 FALTAN = Abs(FALTAN)
'              Else
'                 FALTAN = 0
'              End If
'            MSF.TextMatrix(I%, 10) = FORMATNUM$(FALTAN, "F9.1") 'Faltan
             CanCompPen = CANOCPEN(CI1%)
             CANTIPENRECEP = CANT_RECEP_PENDIENTE_APROB(0, CI1%) 'LA CANTIDAD RECEPCIONADA  PENDIENTE DE APROBACION
            MSF.TextMatrix(i%, 11) = FORMATNUM$(CanCompPen + CANTIPENRECEP, "F9.1") 'Compras
             EnFabricacion = XVALO(FabricacionPendiente(CI1%))
            MSF.TextMatrix(i%, 12) = FORMATNUM$(EnFabricacion, "F9.1") 'Fab.Pend
            MSF.TextMatrix(i%, 13) = FORMATNUM$(EnFabricacion + CanCompPen + CANTIPENRECEP, "F9.1") 'A Desponer
            MSF.TextMatrix(i%, 14) = "0" 'A faltar
            AFALTAR = XVALO(MSF.TextMatrix(i%, 9))  'FALTA''''
            
            If AFALTAR > 0.05 Then
              AFALTAR = AFALTAR - (EnFabricacion + CanCompPen + CANTIPENRECEP)
              If AFALTAR < 0.05 Then AFALTAR = 0
              MSF.TextMatrix(i%, 14) = FORMATNUM$(AFALTAR, "F9.1") 'A faltar
            End If
     Next j%
     On Error Resume Next
     Rst.Close
     Set Rst = Nothing
     RS.Close
     Set RS = Nothing
     On Error GoTo 0
     '
'     Call ORDENAR_MSF(CERTIFICADOS.MSF, 1, 6)
     '
     Call COLOREAR_GRILLASQL(SITORFAB09.MSF, &HFBCA9D)
     INDEXI% = 0
     If IUCOD% >= 0 Then
        If ICOMPO%(IUCOD% + 1) >= 0 Then
          If ICOMPO%(IUCOD% + 1) < MSF.Rows Then
            INDEXI% = ICOMPO%(IUCOD% + 1)
          End If
        End If
     End If
     If INDEXI% > 0 Then
         MSF.Row = INDEXI%
         Call COLOREAR_GRILLA_SOLAFILA(SITORFAB09.MSF, &H8000000F, MSF.Row)
         If MSF.TopRow < MSF.Row - 8 Then
           MSF.TopRow = MSF.Row - 8
         End If
     End If
      
     Screen.MousePointer = 1
'     Me.Caption = Me.Caption & " (Cantidad de Registros: " & I% & ")"
      '
End Sub
Private Sub Text6_Change()
   Call Text11_Change
End Sub
Private Sub IMPRIMIR(ByVal HwndForm As Long, ByVal Ruta As String)
    Dim lRet    As Long
    lRet = ShellExecute(HwndForm, "Print", Ruta, ByVal 0&, 0&, 7)
End Sub


Private Sub Command3_Click()
   Command3.Enabled = False
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
      Call GENEXCEL(AREXPO$, SITORFAB09.MSF)
'     If MSF_2.Visible = True Then
'      Call GENEXCEL(AREXPO$, MUESTRADATOS.MSF_2)
'     Else
'      Call GENEXCEL(AREXPO$, MUESTRADATOS.GRID)
    
'     End If
   End If
   '

99 Command3.Enabled = True

End Sub

Sub GENEXCEL(NombreDeArchivo$, grid As msFlexGrid)
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim valor As Double
   Dim Texto As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   TipoDato% = 0: COLUM& = 0
   For i& = 1 To grid.Rows
       RENG1& = i&: COLUM& = 0
       For COLU1& = 1 To grid.Cols
          If grid.ColWidth(COLU1& - 1) = 0 Then GoTo 10 ' OBVIA LAS COLUMNAS OCULTAS
          COLUM& = COLUM& + 1
          '
          VALORX = grid.TextMatrix(i& - 1, COLU1 - 1)
          If Len(VALORX) = 8 And Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
             FEXU$ = VALORX
             VALORX = FETEXCOR1$(Int(FECHANUM(FEXU$)))
             TipoDato% = 1
          End If
          ' ESTO ES PARA QUE EXPORTE COMO CERO Y NO ESPACIOS
          FFCC$ = UCase$(Left$(FORCAMPO$(COLU1&), 1))
          If FFCC$ = "I" Or FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
             If TRIM$(CStr(VALORX)) = "" Then
                VALORX = 0
             End If
             TipoDato% = 2
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
          If TipoDato% = 1 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yyyy"
          'If TIPODATO% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#,##0.00"
          If TipoDato% = 0 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
          TipoDato% = 0
          VALORX_ANT = VALORX
10     Next COLU1&
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
   Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NombreDeArchivo$, "\", "/"))
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




