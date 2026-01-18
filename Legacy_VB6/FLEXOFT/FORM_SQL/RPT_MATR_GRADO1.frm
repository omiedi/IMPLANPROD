VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form RPT_MATR_GRADO1 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13110
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   13110
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   1920
      TabIndex        =   32
      Top             =   7800
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6360
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exp.Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8160
      TabIndex        =   31
      ToolTipText     =   "Exportar a Excel"
      Top             =   7350
      Width           =   1815
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FBCA9D&
      Height          =   600
      Index           =   3
      Left            =   40
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Text            =   "Matrículas Cobradas (Todas o Según Filtro)"
      Top             =   0
      Width           =   13080
   End
   Begin VB.TextBox Text17 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Index           =   0
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   "Cobrado:"
      Top             =   7320
      Width           =   885
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   50
      TabIndex        =   23
      Top             =   720
      Width           =   1575
      Begin VB.CheckBox Check1 
         Caption         =   "Incluir Alumnos de Baja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1575
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   225
      Left            =   120
      OleObjectBlob   =   "RPT_MATR_GRADO1.frx":0000
      TabIndex        =   22
      Top             =   7440
      Width           =   1560
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   11880
      Picture         =   "RPT_MATR_GRADO1.frx":005E
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "iMPRIME REPORTE ACTIVO"
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   10140
      Begin VB.TextBox txtxombo1 
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
         Height          =   315
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   1030
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   2
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Año de Matrícula:"
         Top             =   240
         Width           =   1605
      End
      Begin VB.ComboBox Combo20 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   240
         Width           =   1320
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   1
         Left            =   165
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "Sala de:"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   6
         Left            =   165
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Grado:"
         Top             =   255
         Width           =   1815
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   8
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Año:"
         Top             =   240
         Width           =   405
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   9
         Left            =   3180
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "Turno:"
         Top             =   615
         Width           =   645
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   5
         Left            =   3165
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "Letra:"
         Top             =   240
         Width           =   645
      End
      Begin VB.TextBox TXCOMBO5 
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
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   255
         Width           =   850
      End
      Begin VB.TextBox TXCOMBO7 
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
         Height          =   315
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   600
         Width           =   850
      End
      Begin VB.TextBox TXCOMBO3 
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
         Height          =   315
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   255
         Width           =   540
      End
      Begin VB.TextBox TXCOMBO2 
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
         Height          =   315
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   615
         Width           =   2460
      End
      Begin VB.TextBox TXCOMBO6 
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
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   255
         Width           =   1030
      End
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2010
         Sorted          =   -1  'True
         TabIndex        =   12
         Text            =   "Combo5"
         Top             =   240
         Width           =   1140
      End
      Begin VB.ComboBox Combo7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2010
         Sorted          =   -1  'True
         TabIndex        =   10
         Text            =   "Combo7"
         Top             =   585
         Width           =   1140
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3885
         Sorted          =   -1  'True
         TabIndex        =   9
         Text            =   "Combo3"
         Top             =   240
         Width           =   740
      End
      Begin VB.ComboBox Combo6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5280
         Sorted          =   -1  'True
         TabIndex        =   11
         Text            =   "Combo6"
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3840
         Sorted          =   -1  'True
         TabIndex        =   13
         Text            =   "Combo2"
         Top             =   600
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker DESDE 
         Height          =   375
         Left            =   6840
         TabIndex        =   28
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   14737632
         CalendarTitleBackColor=   14737632
         Format          =   129957889
         CurrentDate     =   40875
      End
      Begin MSComCtl2.DTPicker HASTA 
         Height          =   375
         Left            =   8640
         TabIndex        =   29
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   14737632
         CalendarTitleBackColor=   14737632
         Format          =   129957889
         CurrentDate     =   40875
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   5325
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Doble Click Para Ver Recibo"
      Top             =   1920
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   9393
      _Version        =   393216
      Cols            =   4
      BackColor       =   14737632
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
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
      Left            =   0
      OleObjectBlob   =   "RPT_MATR_GRADO1.frx":5C70
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   225
      Left            =   120
      OleObjectBlob   =   "RPT_MATR_GRADO1.frx":5EA4
      TabIndex        =   25
      Top             =   7800
      Width           =   1560
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "RPT_MATR_GRADO1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Command1.Enabled = False
    ''
    Call ABRECONEXION
    '
    Call CARGA_INDICE_ALUMNOS
    '
    If RESPUESTA_BUSQUEDA$ = "" Then GoTo 99
    '
    Text1 = RESPUESTA_BUSQUEDA$
    
99  Command1.Enabled = True

End Sub

Private Sub Command2_Click()
'   Text1 = ""
   Call LIMPIARFRM(Me)
   Call CARGA_INFO
End Sub
   Sub LIMPIARFRM(frm As Form)
    Dim CONTROLX As CONTROL ' declaro el objeto control
    For Each CONTROLX In frm.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
        If TypeOf CONTROLX Is TextBox Then
          If CONTROLX.BackColor = &HE0E0E0 Then GoTo 50
            CONTROLX = ""
          End If
50  Next

End Sub

 Private Sub Command3_Click()
   Command3.Enabled = False
   
   'valido que haya informacion en el list
   If MSF.Rows <= 1 Then Exit Sub
   '
   'presenta cuadro de dialogo
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
   ARCHIEX$ = UCase$(CommonDialog1.FileName)
   
   'si selecciona archivo genera excel
   If TRIM$(ARCHIEX$) <> "" Then
      Call GENERAXCEL(RPT_MATR_GRADO.MSF, ARCHIEX$)
   End If
   
99 Command3.Enabled = True

End Sub


Private Sub DESDE_Change()
   Call CARGA_INFO
End Sub

Private Sub DESDE_Click()
   Call CARGA_INFO
End Sub

Private Sub Form_Load()
    Call APLICAR_SKIN(Me, App.Path & "\SKINS\SteelBlue.skn")
    
    Combo2.AddItem ""
    Combo2.AddItem "MAÑANA"
    Combo2.AddItem "TARDE"
    Combo2.AddItem "NOCHE"
    '
    'GRADO
    Combo5.AddItem ""
    Combo5.AddItem "1º"
    Combo5.AddItem "2º"
    Combo5.AddItem "3º"
    Combo5.AddItem "4º"
    Combo5.AddItem "5º"
    Combo5.AddItem "6º"
    Combo5.AddItem "7º"
    '
    'SALA DE
    Combo7.AddItem ""
    Combo7.AddItem "3"
    Combo7.AddItem "4"
    Combo7.AddItem "5"
    '
    'AÑO
    Combo6.AddItem ""
    Combo6.AddItem "1ro"
    Combo6.AddItem "2do"
    Combo6.AddItem "3ro"
    Combo6.AddItem "4to"
    Combo6.AddItem "5to"
    Combo6.AddItem "6to"
    '
'    'COLOR
'    Combo8.AddItem ""
'    Combo8.AddItem "BLANCO"
'    Combo8.AddItem "ROSA"
'    Combo8.AddItem "ROJA"
'    Combo8.AddItem "VERDE"
'    Combo8.AddItem "AZUL"
'    Combo8.AddItem "AMARILLA"
'    Combo8.AddItem "TURQUESA"
'    Combo8.AddItem "CELESTE"
    '
    'LETRA
    Combo3.AddItem ""
    Combo3.AddItem "A"
    Combo3.AddItem "B"
    Combo3.AddItem "C"
    '
    Call CARGAAÑOS(RPT_MATR_GRADO.Combo20)
    '
    DESDE.Value = Date
    HASTA.Value = Date
    Call ENCABEZA_GRILLA
    '
    On Error Resume Next
    '
    'ELIMINO LA VISTA
    SQLX$ = "DROP VIEW MATRI_PAGAS"
    BASSSEX.Execute SQLX$
    '
    'GENERO LA VISTA
    SQLX$ = "CREATE VIEW [MATRI_PAGAS] AS "
    SQLX$ = SQLX$ + " SELECT NOMBRE_APELLIDO, VARIMOVI, ALUMNOS.LEGAJO AS LEGAJO, TIPOMOVIM, DEBE, "
    SQLX$ = SQLX$ + " HABER, CODICOM_LAR, ALUMNOS.GRADO, ALUMNOS.AÑO, ALUMNOS.TURNO, ALUMNOS.LETRA, "
    SQLX$ = SQLX$ + " ALUMNOS.SALA_DE, FECHEMISI, MES_CUOTA, ANO_CUOTA, CUOTA_ORIGINAL, BONIFICACION, RECARGO, ALUMNOS.STATUS AS STATUS"
    SQLX$ = SQLX$ + " FROM ALUMNOS LEFT JOIN CAJABANC_COL ON ALUMNOS.LEGAJO=CAJABANC_COL.LEGAJO "
    SQLX$ = SQLX$ + " Where ALUMNOS.LEGAJO > 0"
    BASSSEX.Execute SQLX$
    '
    On Error GoTo 0
    '
End Sub

Private Sub Command25_Click()
   Command25.Enabled = False
   'valido
   If MSF.Rows <= 1 Then
      MsgBox "GRILLA DE IMPRESIÓN VACIA"
      GoTo 99
   End If
   '
   Set GRILLA_IMPRESION = MSF
   '
   PREIMPRE.txtEncabezado2.Text = ""

   '
   
   ENCABEZADO1$ = "Situación de Matrículas Generadas"
   
   PREIMPRE.txtEncabezado1.Text = ENCABEZADO1$
   PREIMPRE.txtPiePagina = "Fin de Reporte"
   PREIMPRE.Combo1.Text = "Vertical"
   PREIMPRE.Show 1
   '
99 Command25.Enabled = True

End Sub

Private Sub HASTA_Change()
   Call CARGA_INFO
End Sub

Private Sub HASTA_Click()
  Call CARGA_INFO
End Sub

Private Sub MSF_DblClick()
   LGJO& = XVALO(MSF.TextMatrix(MSF.Row, 3))
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   i& = MSF.MouseRow
   j& = MSF.MouseCol
   '
   If i& < 1 Then Exit Sub
   'VALIDO
   If i& = 0 Or i& > MSF.Rows Then Exit Sub
   If j& > MSF.Cols Then Exit Sub
   '
   'RESTAURO AL COLOR ORIGINAL LA COLUMNA 0
   MSF.COL = 0
   For H& = 0 To MSF.Rows - 1
     MSF.Row = H&
     MSF.CellBackColor = &H8000000F
   Next H&
   '
   'PINTO DE AMARILLO LA SELECCIONADA
   MSF.Row = i&
   MSF.CellBackColor = &H80FFFF   ' LA PINTA DE AMARILLA
   '
   If LGJO& < 1 Then Exit Sub
   '
   DOCU$ = TRIM$(MSF.TextMatrix(MSF.Row, 9))
   VARMOV$ = TRIM$(MSF.TextMatrix(MSF.Row, 12))
   NURECI& = XVALO(Mid$(DOCU$, 5))
   '
   If VARMOV$ = "COBMATRI" Then
     Call PRIMATRICULA(LGJO&, NURECI&, DOCU$, "COBMATRI")
   ElseIf VARMOV$ = "RCOB" Then
     Call PRIRECIBO(LGJO&, NURECI&, DOCU$, "RCOB")
   End If
   '
End Sub

Private Sub Option2_Click(Index As Integer)
    Call CARGA_INFO
End Sub

Private Sub Text1_Change()
'     Call LIMPIAR
     TEXT2 = VALOADMIN("ALUMNOS", "NOMBRE_APELLIDO", "LEGAJO = " & XVALO(Text1), "NOMESS")
     Call CARGA_INFO
     

End Sub
Sub CARGA_INFO()
   '
   MSF.Rows = 1
   MSF.Cols = 18
   '
   ACUMULA_ACOBRAR# = 0
   ACUMULA_COBRADO# = 0
   
   '
   CONDI$ = " LEGAJO > 0 "
   '
   'BAJA
   If Check1.Value = 0 Then CONDI$ = CONDI$ + " AND STATUS >= 0"
   '
   'FECHAS
   DES% = CVINT(DESDE.Value)
   HAS% = CVINT(HASTA.Value)
   DESDEX = FETEXCOR1$(DES%)
   HASTAX = FETEXCOR1$(HAS%)
   FILA& = 0

   GRADOX% = XVALO(TXCOMBO5)
   SALADEX% = XVALO(TXCOMBO7)
   LETRX$ = TRIM$(TXCOMBO3)
   ANOX% = XVALO(TXCOMBO6)
   ANOXMATRICULA% = XVALO(txtxombo1.Text)
   TURNOX$ = TRIM$(TXCOMBO2)
   
   REG& = CantRegistros&("MATRI_PAGAS", CONDI$, "NOMESS")
   If REG& < 1 Then Exit Sub
   '
   MSF.Redraw = False
   '
      'CARGA EL ENCABEZADO DEL TRANSPORTE
   condi0$ = "(CUENTA_CONT = 2 OR CUENTA_CONT = 3)"
   condi0$ = condi0$ + " AND FECHEMISI < #" & DESDEX & "#"

   SQLX$ = "SELECT * FROM MATRI_PAGAS "
   SQLX$ = SQLX$ + " WHERE " + CONDI$
   SQLX$ = SQLX$ + " AND HABER > 0.05 "
   SQLX$ = SQLX$ + " AND (VARIMOVI = 'COBMATRI' OR VARIMOVI = 'MATRICULA')"
   If GRADOX% > 0 Then SQLX$ = SQLX$ + " AND GRADO = " & GRADOX%
   If SALADEX% > 0 Then SQLX$ = SQLX$ + " AND SALA_DE = " & SALADEX%
   If LETRX$ <> "" Then SQLX$ = SQLX$ + " AND LETRA = '" & LETRX$ & "'"
   If ANOX% > 0 Then SQLX$ = SQLX$ + " AND AÑO = " & ANOX%
   If ANOXMATRICULA% > 0 Then SQLX$ = SQLX$ + " AND ANO_CUOTA = " & ANOXMATRICULA%
   If TURNOX$ <> "" Then SQLX$ = SQLX$ + " AND TURNO = '" & TURNOX$ & "'"
   If DES% > 0 Then SQLX$ = SQLX$ + " AND FECHEMISI >= '" & DESDEX & "'"
   If HAS% > 0 Then SQLX$ = SQLX$ + " AND FECHEMISI <= '" & HASTAX & "'"
   
'   If SINMATRICULA% > 0 Then SQLX$ = SQLX$ + " AND (VARIMOVI = '' OR VARIMOVI IS NULL)"
'   If CONMATRICULA% > 0 Then SQLX$ = SQLX$ + " AND VARIMOVI <> ''"
   SQLX$ = SQLX$ + " ORDER BY  FECHEMISI,GRADO,SALA_DE,LETRA,AÑO,TURNO,ANO_CUOTA,LEGAJO "
   '
   ProgressBar1.Min = 0
   ProgressBar1.Value = 0
   ProgressBar1.Max = REG&
   '
   On Error Resume Next
   PASO1% = 0
50 'Set RECITMP = BaDaABRE_ADMIN_ALQ.OpenRecordset(SQLX$, dbOpenSnapshot)
   Set RECITMP = BASSSEX.Execute(SQLX$)
   If Err > 0 And PASO1% < 5 Then
      Call pausa
      PASO1% = PASO1% + 1
      On Error GoTo 0
      GoTo 50
   End If
   On Error GoTo 0
   '
   With RECITMP
      Do While Not .EOF
           '
           PROGRE% = PROGRE% + 1
           ProgressBar1.Visible = True
           ProgressBar1.Value = PROGRE%
           SkinLabel2 = SAFETEXT$(CLng((ProgressBar1.Value * 100) / ProgressBar1.Max) & " %")
           SkinLabel1 = Str$(PROGRE%) + " /" + Str(REG&)
           FEMI% = CVINT(!FECHEMISI)
           NombreAPeX$ = SAFETEXT$(!NOMBRE_APELLIDO)
           GRADOX1% = SAFETEXT$(XVALO(!GRADO))
           SALADEX1% = SAFETEXT$(XVALO(!SALA_DE))
           LETRAX1$ = SAFETEXT$(!LETRA)
           ANOX1% = SAFETEXT$(XVALO(!AÑO))
           TURNOX$ = SAFETEXT$(!TURNO)
           '
           DOCUM$ = SAFETEXT$(!CODICOM_LAR)
           MEX% = XVALO(!MES_CUOTA)
           ANOX% = XVALO(!ANO_CUOTA)
           CONCEP = SAFETEXT$(!VARIMOVI)
           CUO_ORIGINALX# = XVALO(!DEBE)
           BONIFICACIONX# = XVALO(!BONIFICACION)
           RECARGOX# = XVALO(!RECARGO)
           SUPAGO# = XVALO(!HABER)
           LGJO& = XVALO(!LEGAJO)
          '
           FILA& = FILA& + 1
           MSF.Rows = FILA& + 1
           MSF.TextMatrix(FILA&, 1) = FECHATEX$(FEMI%)
           
           MSF.TextMatrix(FILA&, 2) = NombreAPeX$
           MSF.TextMatrix(FILA&, 3) = SAFETEXT$(LGJO&)
           MSF.TextMatrix(FILA&, 4) = SAFETEXT$(SALADEX1%) 'SALA
           MSF.TextMatrix(FILA&, 5) = SAFETEXT$(GRADOX1%) 'GRADOX1%
           MSF.TextMatrix(FILA&, 6) = SAFETEXT$(ANOX1%) 'AÑO
           MSF.TextMatrix(FILA&, 7) = SAFETEXT$(LETRAX1$) 'LETRA
           MSF.TextMatrix(FILA&, 8) = Left$(SAFETEXT$(TURNOX$), 1) 'TURNO)
           MSF.TextMatrix(FILA&, 9) = DOCUM$
             MESXX$ = "00" & TRIM$(Str(MEX%))
             MESXX$ = Right(MESXX$, 2)
           MSF.TextMatrix(FILA&, 10) = MESXX$
           MSF.TextMatrix(FILA&, 11) = SAFETEXT$(ANOX%)
           MSF.TextMatrix(FILA&, 12) = CONCEP
'           MSF.TextMatrix(FILA&, 13) = TRIM$(FORMATNUM$(CUO_ORIGINALX#, "F12.2"))
'           MSF.TextMatrix(FILA&, 14) = TRIM$(FORMATNUM$(BONIFICACIONX#, "F12.2"))
'           MSF.TextMatrix(FILA&, 15) = TRIM$(FORMATNUM$(RECARGOX#, "F12.2"))
           MSF.TextMatrix(FILA&, 13) = TRIM$(FORMATNUM$(SUPAGO#, "F12.2"))
'           ACUMULA_ACOBRAR# = ACUMULA_ACOBRAR# + CUO_ORIGINALX#
           ACUMULA_COBRADO# = ACUMULA_COBRADO# + SUPAGO#
           .MoveNext
      Loop
    End With
    RECITMP.Close
    Set RECITMP = Nothing
    '
    Call ENCABEZA_GRILLA
    MSF.Redraw = True
    
    Text3 = FORMATNUM$(ACUMULA_COBRADO#, "f14.2") '
'    Text4 = FORMATNUM$(ACUMULA_ACOBRAR#, "f14.2")
    '
        'CARGO A PIE DE LA GRILLA EL TOTOL
'    MSF.Rows = FILA& + 7
    
'    MSF.TextMatrix(FILA& + 2, 12) = "A Cobrar: ": Call COLOREAR_GRILLA_SOLO_UNA_AFILA(REPORTE_GENERAL.MSF, &HFCE4A7, FILA& + 2)
'    MSF.TextMatrix(FILA& + 2, 15) = "Cobrado: "
'
'    MSF.TextMatrix(FILA& + 2, 13) = Text4
'    MSF.TextMatrix(FILA& + 2, 16) = Text3

End Sub
Sub ENCABEZA_GRILLA()
    MSF.Cols = 17

    NUCOLU% = 1: TEPRUEBA = "10": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = ""
    NUCOLU% = 2: TEPRUEBA = "10/10/20120": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Fecha Emis."
    NUCOLU% = 3: TEPRUEBA = "ABCDEFGHIJKLLLMNÑOPQR": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Nombre Y Apellido"
    NUCOLU% = 4: TEPRUEBA = "10/10/201": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Lgjo."
    NUCOLU% = 5: TEPRUEBA = "ABC": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Sl"
    NUCOLU% = 6: TEPRUEBA = "ABC": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Gd"
    NUCOLU% = 7: TEPRUEBA = "ABC": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Añ"
    NUCOLU% = 8: TEPRUEBA = "ABC": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Lt"
    NUCOLU% = 9: TEPRUEBA = "ABC": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Tr"
    NUCOLU% = 10: TEPRUEBA = "12345678900000000": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Documento"
    NUCOLU% = 11: TEPRUEBA = "ABCD": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Mes"
    NUCOLU% = 12: TEPRUEBA = "CHAPA": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Año"
    NUCOLU% = 13: TEPRUEBA = "123456789012": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Concepto"
'    NUCOLU% = 14: TEPRUEBA = "123456789012": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Imp.Original"
'    NUCOLU% = 15: TEPRUEBA = "123456789012": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Bonificación"
'    NUCOLU% = 16: TEPRUEBA = "123456789012": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Recargo"
    NUCOLU% = 14: TEPRUEBA = "123456789012": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Su Pago"
    '
    MSF.MergeCells = flexMergeRestrictColumns
    MSF.MergeCol(2) = True

End Sub

Private Sub TXCOMBO2_Change()
      Call CARGA_INFO
End Sub

Private Sub TXCOMBO3_Change()
   TXTCOMBO3 = Combo3.Text
   Call CARGA_INFO

End Sub

Private Sub TXCOMBO5_Change()
   Call CARGA_INFO
End Sub

Private Sub TXCOMBO6_Change()
  Call CARGA_INFO
End Sub

Private Sub TXCOMBO7_Change()
      Call CARGA_INFO
End Sub

Private Sub txtxombo1_Change()
     Call CARGA_INFO

End Sub

Private Sub Check1_Click()
 Call CARGA_INFO
End Sub

Private Sub Combo2_Change()
   TXCOMBO2 = Combo2.Text
'   Call CARGA_INFO

End Sub

Private Sub Combo2_Click()
   TXCOMBO2 = Combo2.Text
   Call CARGA_INFO
End Sub

Private Sub Combo20_Change()
   txtxombo1 = Combo20.Text
'   Call CARGA_INFO

End Sub

Private Sub Combo20_Click()
   txtxombo1 = Combo20.Text
   Call CARGA_INFO

End Sub

Private Sub Combo3_Change()
   TXCOMBO3 = Combo3.Text
'   Call CARGA_INFO

End Sub

Private Sub Combo3_Click()
   TXCOMBO3 = Combo3.Text
   Call CARGA_INFO

End Sub

Private Sub Combo5_Change()
   TXCOMBO5 = Combo5.Text
'   Call CARGA_INFO

End Sub

Private Sub Combo5_Click()
   TXCOMBO5 = Combo5.Text
   Call CARGA_INFO

End Sub

Private Sub Combo6_Change()
   TXCOMBO6 = Combo6.Text
'   Call CARGA_INFO

End Sub

Private Sub Combo6_Click()
   TXCOMBO6 = Combo6.Text
'   Call CARGA_INFO
End Sub

Private Sub Combo7_Change()
     TXCOMBO7 = Combo7.Text
'     Call CARGA_INFO

End Sub

Private Sub Combo7_Click()
   TXCOMBO7 = Combo7.Text
   Call CARGA_INFO

End Sub

Private Sub pausa() ' este el procedimiento creado por nosotros

    Dim comenzar
    
    Dim controlar
    
    comenzar = Timer
    
    Do Until controlar >= comenzar + 1  ' este es el bucle de espera
    
    controlar = Timer
    
    DoEvents ' esta sentencia evita un bucle egoista
    
    Loop

End Sub
