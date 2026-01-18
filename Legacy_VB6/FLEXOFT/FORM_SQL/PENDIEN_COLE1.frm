VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form PENDIEN_COLE1 
   Caption         =   "Pendientes"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   12555
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   3600
      TabIndex        =   32
      Top             =   7800
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6000
      Top             =   3960
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
      Height          =   495
      Left            =   6680
      TabIndex        =   31
      ToolTipText     =   "Exportar a Excel"
      Top             =   7720
      Width           =   1815
   End
   Begin VB.TextBox Text4 
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
      Left            =   11040
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   705
      Width           =   1155
   End
   Begin MSComCtl2.DTPicker Fecha 
      Height          =   360
      Left            =   11040
      TabIndex        =   29
      Top             =   690
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
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
      Format          =   131334145
      CurrentDate     =   41434
   End
   Begin VB.TextBox c 
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
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Text            =   "Vencimiento a Fecha:"
      Top             =   705
      Width           =   1845
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   12540
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
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   255
         Width           =   1030
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
         Left            =   9700
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   255
         Width           =   2480
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
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   255
         Width           =   540
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
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   255
         Width           =   850
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
         Height          =   360
         Index           =   5
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Letra:"
         Top             =   240
         Width           =   645
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
         Height          =   360
         Index           =   9
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "Turno:"
         Top             =   240
         Width           =   645
      End
      Begin VB.TextBox c 
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
         Height          =   360
         Index           =   8
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Año:"
         Top             =   240
         Width           =   525
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
         Height          =   360
         Index           =   6
         Left            =   165
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Grado:"
         Top             =   240
         Width           =   1815
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
         Height          =   360
         Index           =   1
         Left            =   6860
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Sala de:"
         Top             =   240
         Width           =   855
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
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   255
         Width           =   850
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
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   21
         Text            =   "Combo5"
         Top             =   240
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
         Left            =   3960
         Sorted          =   -1  'True
         TabIndex        =   23
         Text            =   "Combo3"
         Top             =   240
         Width           =   760
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
         Left            =   5380
         Sorted          =   -1  'True
         TabIndex        =   24
         Text            =   "Combo6"
         Top             =   240
         Width           =   1335
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
         Left            =   7770
         Sorted          =   -1  'True
         TabIndex        =   22
         Text            =   "Combo7"
         Top             =   240
         Width           =   1140
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
         Left            =   9680
         Sorted          =   -1  'True
         TabIndex        =   25
         Text            =   "Combo2"
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7720
      Width           =   2055
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Imprimir"
      Height          =   690
      Left            =   0
      Picture         =   "PENDIEN_COLE1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7560
      Width           =   1695
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
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Cuotas y Matrículas Vencidas"
      Top             =   0
      Width           =   12705
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ver Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   6
      Top             =   7720
      Width           =   1815
   End
   Begin VB.TextBox TEXT2 
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
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   705
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Left            =   1305
      TabIndex        =   3
      ToolTipText     =   "Ingrese Nª Legajo Para Filtrar"
      Top             =   705
      Width           =   1755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   2
      Top             =   705
      Width           =   255
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
      Index           =   6
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Legajo:"
      Top             =   705
      Width           =   1005
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   5565
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   12510
      _ExtentX        =   22066
      _ExtentY        =   9816
      _Version        =   393216
      Cols            =   4
      BackColor       =   14737632
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2880
      OleObjectBlob   =   "PENDIEN_COLE1.frx":5C12
      Top             =   8160
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   225
      Left            =   1800
      OleObjectBlob   =   "PENDIEN_COLE1.frx":5E46
      TabIndex        =   26
      Top             =   7680
      Width           =   1560
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   225
      Left            =   1800
      OleObjectBlob   =   "PENDIEN_COLE1.frx":5EA4
      TabIndex        =   27
      Top             =   8040
      Width           =   1560
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   9480
      TabIndex        =   4
      Top             =   8160
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PENDIEN_COLE1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim VECTRECI#()
    Dim VECTMATRICULA#()

Sub PROCESO_ACTUALIZA_FECHA_VENCIMIENTO()
   'CON LA COLUMNA MES_CUOTA Y ANO_CUOTA GENERA LA FECHA DE VENCIMIENTO CON DIA 20 POR DEFECTO
   'ESTO SE REALIZO PARA PODER OBTENER REPORTE DE CUOTAS VENCIDAS A LA FECHA QUE SE EMITA EL REPORTE.
   'EL REPORTE ES EL DEL FRM PENDIEN_COLE
   ''
   'POR AHORA SE GENERA SOLO EN ESTE FRM HABRIA SI SE USA EN OTRO LADO HABRIA QUE CORRER ESTE PROCESO POR LAS DUDAS
   'O BIEN GUARDAR ESTA INFO CUANDO SE GENERAN LAS CUOTAS.
   FECHAX = FETEXCOR1$(FECHANUM("01/01/2011"))
    SQLX$ = "SELECT FECHVENC, ANO_CUOTA, MES_CUOTA FROM CAJABANC_COL"
    SQLX$ = SQLX$ + " WHERE (FECHVENC IS NULL OR FECHVENC < '" & FECHAX & "')"
    Set AUX = New ADODB.Recordset
    AUX.Open (SQLX$), BASSSEX, adOpenKeyset, adLockPessimistic, adCmdText
    With AUX
      Do While Not .EOF
       ANO1$ = TRIM$(!ANO_CUOTA)
       MEX$ = TRIM$(!MES_CUOTA)
       While Len(MEX$) < 2:  MEX$ = "0" & MEX$: Wend
       VENCI = "20/" & MEX$ & "/" & ANO1$ ' FECHA 20 DE VENCIMIENTO POR DEFECTO
       
       VENCI1% = CVINT(VENCI)
       VENCI2 = FETEXCOR1$(VENCI1%)
       !FECHVENC = VENCI2
      .Update
      .MoveNext
      Loop
    
    
    End With

End Sub

Private Sub Combo2_Change()
   TXCOMBO2 = Combo2.Text
End Sub

Private Sub Combo2_Click()
   TXCOMBO2 = Combo2.Text
End Sub

Private Sub Combo3_Change()
   TXCOMBO3 = Combo3.Text
End Sub

Private Sub Combo3_Click()
   TXCOMBO3 = Combo3.Text
End Sub

Private Sub Combo5_Change()
   TXCOMBO5 = Combo5.Text
End Sub

Private Sub Combo5_Click()
   TXCOMBO5 = Combo5.Text
End Sub

Private Sub Combo6_Change()
   TXCOMBO6 = Combo6.Text
End Sub

Private Sub Combo6_Click()
 TXCOMBO6 = Combo6.Text
End Sub


Private Sub Combo7_Change()
   TXCOMBO7 = Combo7.Text

End Sub

Private Sub Combo7_Click()
   TXCOMBO7 = Combo7.Text
End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
    '
    '
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
   Command2.Enabled = False
   Text1 = ""
   Text3 = ""
   Call CARGA_GRILLA_CTA_CTE("TODOS")
   '
   Call ENCABEZA_GRILLA
   Command2.Enabled = 1
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
   LGJO& = XVALO(Text1)
   GRADOX% = XVALO(TXCOMBO5)
   SALADEX% = XVALO(TXCOMBO7)
   LETRX$ = TRIM$(TXCOMBO3)
   ANOX% = XVALO(TXCOMBO6)
   TURNOX$ = TRIM$(TXCOMBO2)
   '
   If LGJO& > 0 Then FILTRO$ = "Legajo: " & LGJO&
   If GRADOX% > 0 Then FILTRO$ = FILTRO$ + " Grado: " & GRADOX%
   If SALADEX% > 0 Then FILTRO$ = FILTRO$ + " Sala de: " & SALADEX%
   If LETRX$ <> "" Then FILTRO$ = FILTRO$ + " Letra:" & LETRX$
   If ANOX% > 0 Then FILTRO$ = FILTRO$ + " Año: " & ANOX%
   If TURNOX$ <> "" Then FILTRO$ = FILTRO$ + " Turno: " & TURNOX$

   PREIMPRE.txtEncabezado2.Text = "Filtro: " & FILTRO$
   '
   ENCABEZADO1$ = "Cuotas y/o Matrículas Pendientes"
   
   PREIMPRE.txtEncabezado1.Text = ENCABEZADO1$
   PREIMPRE.txtPiePagina = "Fin de Reporte"
   PREIMPRE.Combo1.Text = "Vertical"
   PREIMPRE.Show 1
   '
99 Command25.Enabled = True

End Sub


Private Sub Fecha_CloseUp()
     Text4 = Fecha.Value

End Sub

Private Sub Form_Load()
'HASTA ACA LLEGUE 19/04/15

    Call APLICAR_SKIN(Me, App.Path & "\SKINS\SteelBlue.skn")
    Call ENCABEZA_GRILLA
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
    'LETRA
    Combo3.AddItem ""
    Combo3.AddItem "A"
    Combo3.AddItem "B"
    Combo3.AddItem "C"
    '
    Fecha.Value = Date
    Call PROCESO_ACTUALIZA_FECHA_VENCIMIENTO
    'ELIMINO LA VISTA
    SQLX$ = "DROP VIEW ALUM_CAJAB"
    On Error Resume Next
    BASSSEX.Execute SQLX$
    On Error GoTo 0
    '
    'GENERO LA VISTA
    SQLX$ = "CREATE VIEW [ALUM_CAJAB] AS "
    SQLX$ = SQLX$ + " SELECT NOMBRE_APELLIDO, VARIMOVI, ALUMNOS.LEGAJO AS LEGAJO, TIPOMOVIM, DEBE, NUMEDOC,FECHVENC, "
    SQLX$ = SQLX$ + " HABER, CODICOM_LAR, ALUMNOS.GRADO, ALUMNOS.AÑO, ALUMNOS.TURNO, ALUMNOS.LETRA, "
    SQLX$ = SQLX$ + " ALUMNOS.SALA_DE, FECHEMISI, MES_CUOTA, ANO_CUOTA, CUOTA_ORIGINAL, BONIFICACION, RECARGO, ALUMNOS.STATUS AS STATUS"
    SQLX$ = SQLX$ + " FROM ALUMNOS LEFT JOIN CAJABANC_COL ON ALUMNOS.LEGAJO=CAJABANC_COL.LEGAJO "
    SQLX$ = SQLX$ + " Where ALUMNOS.LEGAJO > 0"
    BASSSEX.Execute SQLX$
    '
    On Error GoTo 0

    Call CARGA_VECTOR_SUPAGO
End Sub
Sub CARGA_VECTOR_SUPAGO()
    '
     Erase VECTRECI#
    Erase VECTMATRICULA#
   SQLX$ = " SELECT MAX(NUMEDOC) AS MAXIMOX FROM CAJABANC_COL"
    
    Set RS = BASSSEX.Execute(SQLX$)
    MAXIMO1& = RS!MAXIMOX
    RS.Close
    Set RS = Nothing
    '
    ReDim VECTRECI#(2, 2)
    ReDim VECTRECI#(1, MAXIMO1&)
    ReDim VECTRECI#(2, MAXIMO1&)

    
    SQLX$ = " SELECT SUM (HABER) AS HABE, SUM(BONIFICACION)AS BONIFICAX ,SUM(RECARGO) AS RECARGOX ,docu_apli FROM CAJABANC_COL "
    SQLX$ = SQLX$ + "  WHERE( VARIMOVI = 'RCOB' or VARIMOVI = 'COBMATRI')"
    SQLX$ = SQLX$ + "  AND TIPOMOVIM = 'RECIBO COBRANZA'"
    SQLX$ = SQLX$ + "  group by docu_apli"
    SQLX$ = SQLX$ + " ORDER BY docu_apli "
    Set SUPATMP = BASSSEX.Execute(SQLX$)
    With SUPATMP
      Do While Not .EOF
         HABEX# = XVALO(!HABE)
         BONIFICA1# = Abs(XVALO(!BONIFICAX))
         RECARGO1# = -1 * XVALO(!RECARGOX)
         SUPAGOX1# = HABEX# + BONIFICA1# + RECARGO1#
         
         
         If Left(!DOCU_APLI, 3) = "CTA" Then
           VECTRECI#(1, XVALO(Right(!DOCU_APLI, 8))) = SUPAGOX1#
         ElseIf Left(!DOCU_APLI, 3) = "MTR" Then
           VECTRECI#(2, XVALO(Right(!DOCU_APLI, 8))) = SUPAGOX1#
         End If
         
         .MoveNext
      Loop
   End With
   '
   SUPATMP.Close
   Set SUPATMP = Nothing
   '
    SQLX$ = " SELECT SUM (HABER) AS HABE, SUM(BONIFICACION)AS BONIFICAX ,SUM(RECARGO) AS RECARGOX,NUMEDOC  FROM CAJABANC_COL "
    SQLX$ = SQLX$ + "  WHERE VARIMOVI = 'COBMATRI' "
    SQLX$ = SQLX$ + "  AND TIPOMOVIM = 'RECIBO COBRANZA'"
    SQLX$ = SQLX$ + "  group by docu_apli, NUMEDOC"
    SQLX$ = SQLX$ + " ORDER BY NUMEDOC "
    Set SUPATMP = BASSSEX.Execute(SQLX$)
    With SUPATMP
      Do While Not .EOF
         HABEX# = XVALO(!HABE)
         BONIFICA1# = Abs(XVALO(!BONIFICAX))
         RECARGO1# = -1 * XVALO(!RECARGOX)
         SUPAGOX1# = HABEX# + BONIFICA1# + RECARGO1#
         '
         IND1& = XVALO(!NUMEDOC)
         ReDim Preserve VECTMATRICULA#(IND1&)
         VECTMATRICULA#(IND1&) = SUPAGOX1#
         .MoveNext
      Loop
   End With
   '
   SUPATMP.Close
   Set SUPATMP = Nothing

   'OR VARIMOVI = 'COBMATRI')"
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
      Call GENERAXCEL(PENDIEN_COLE.MSF, ARCHIEX$)
   End If
   
99 Command3.Enabled = True

End Sub

Private Sub Text1_Change()

     MSF.Rows = 1
     Text2 = ""
     Text3 = ""
     LGJO& = XVALO(Text1)
     If LGJO& < 1 Then Exit Sub
     Text2 = VALOADMIN("ALUMNOS", "NOMBRE_APELLIDO", "LEGAJO = " & XVALO(Text1), "NOMESS")
     
     Z& = NURECIBO&(FORMATEADO$)
     Text24 = FORMATEADO$
     '
     Call CARGA_GRILLA_CTA_CTE
End Sub
Sub CARGA_RECIBOS()
   '
   MSF.Rows = 1
   MSF.Cols = 11
   '
   LEGAJO& = XVALO(Text1)
   
   CONDI$ = " (VARIMOVI = 'RCOB' OR VARIMOVI = 'COBMATRI')"
   CONDI$ = CONDI$ + "  AND TIPOMOVIM = 'RECIBO COBRANZA'"
   If LEGAJO& > 0 Then CONDI$ = CONDI$ + " AND LEGAJO = " & LEGAJO& & ""
   '
   If CantRegistros&("CAJABANC_COL", CONDI$, "NOMESS") < 1 Then Exit Sub
   '
   MSF.Redraw = False
   '
   SQLX$ = "SELECT * FROM CAJABANC_COL "
   SQLX$ = SQLX$ + " WHERE " + CONDI$
   Set RECITMP = BaDaABRE_ADMIN_ALQ.OpenRecordset(SQLX$, dbOpenSnapshot)
   With RECITMP
      Do While Not .EOF
           FEMI% = CVINT(!FECHEMISI)
           DOCUM$ = SAFETEXT$(!CODICOM_LAR)
           MEX% = XVALO(!MES_CUOTA)
           ANOX% = XVALO(!ANO_CUOTA)
           CONCEP = SAFETEXT$(!VARIMOVI)
           CUO_ORIGINALX# = XVALO(!CUOTA_ORIGINAL)
           BONIFICACIONX# = XVALO(!BONIFICACION)
           RECARGOX# = XVALO(!RECARGO)
           SUPAGO# = XVALO(!HABER)
           LGJO& = XVALO(!LEGAJO)
          '
           FILA& = FILA& + 1
           MSF.Rows = FILA& + 1
           MSF.TextMatrix(FILA&, 1) = FECHATEX$(FEMI%)
           MSF.TextMatrix(FILA&, 2) = DOCUM$
           MESXX$ = "00" & TRIM$(Str(MEX%))
           MESXX$ = Right(MESXX$, 2)
           MSF.TextMatrix(FILA&, 3) = MESXX$
           MSF.TextMatrix(FILA&, 4) = SAFETEXT$(ANOX%)
           MSF.TextMatrix(FILA&, 5) = CONCEP
           MSF.TextMatrix(FILA&, 6) = TRIM$(FORMATNUM$(CUO_ORIGINALX#, "F12.2"))
           MSF.TextMatrix(FILA&, 7) = TRIM$(FORMATNUM$(BONIFICACIONX#, "F12.2"))
           MSF.TextMatrix(FILA&, 8) = TRIM$(FORMATNUM$(RECARGOX#, "F12.2"))
           MSF.TextMatrix(FILA&, 9) = TRIM$(FORMATNUM$(SUPAGO#, "F12.2"))
           MSF.TextMatrix(FILA&, 10) = TRIM$(Str$(LGJO&))
           .MoveNext
      Loop
    End With
    RECITMP.Close
    Set RECITMP = Nothing
    '
    Call ENCABEZA_GRILLA
    MSF.Redraw = True



End Sub
Sub ENCABEZA_GRILLA()
     '
     MSF.Cols = 10
 
     NUCOLU% = 1: TEPRUEBA = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Nombre y Apellido"
     NUCOLU% = 2: TEPRUEBA = "10/10/20120": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Fecha Emis."
     NUCOLU% = 3: TEPRUEBA = "12345678900000000000": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Documento"
     NUCOLU% = 4: TEPRUEBA = "ABCD": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Mes"
     NUCOLU% = 5: TEPRUEBA = "CHAPA": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Año"
     NUCOLU% = 6: TEPRUEBA = "123456789012": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Concepto"
     NUCOLU% = 7: TEPRUEBA = "1234567890": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Debe"
     NUCOLU% = 8: TEPRUEBA = "1234567890": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Haber"
     NUCOLU% = 9: TEPRUEBA = "1234567890": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Saldo"
     NUCOLU% = 10: TEPRUEBA = "1234589": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "LGJO."

End Sub

Sub CARGA_GRILLA_CTA_CTE(Optional MODO$)
     
     '
     
     LGJO& = XVALO(Text1)
     
     GRADOX% = XVALO(TXCOMBO5)
     SALADEX% = XVALO(TXCOMBO7)
     LETRX$ = TRIM$(TXCOMBO3)
     ANOX% = XVALO(TXCOMBO6)
     TURNOX$ = TRIM$(TXCOMBO2)
     '

     'VALIDA SI NO HAY NINGUN OBJETO DE FILTRO CON INFORMACION SE VA
     If MODO$ = "" Then 'SI MODO ES TODOS CONTINUA SI ESTA VACIO VALIDA QUE ALGUNO DE LOS OTROS FILTROS ESTEN ACTIVOS CON INFORMACION
       If LGJO& < 1 And GRADOX% < 1 And SALADEX% < 1 And LETRX$ = "" And TURNOX$ = "" And ANOX% < 1 And IsDate(CVDAT(CVINT(CDate(Fecha.Value)))) = False Then MSF.Rows = 1: Exit Sub
     End If
     '
     'MANEJO DE FECHA DE VENCIMIENTO

     MSF.Rows = 1
     MSF.Cols = 10
     MSF.Redraw = False
     '
     'formateo fecha para consulta
     HASTAX = FETEXCOR1$(CVINT(Fecha.Value))
     
     CONDI$ = "(TIPOMOVIM = 'CUOTA' OR  TIPOMOVIM = 'MATRICULA')"
     If LEGAJO& > 0 Then CONDI$ = CONDI$ + " AND LEGAJO = " & LGJO&
     CONDI$ = CONDI$ + " AND (VARIMOVI = 'CUOTA' OR VARIMOVI = 'MATRICULA')"
     REG& = CantRegistros&("CAJABANC_COL", CONDI$, "NOMESS")
     '
     ProgressBar1.Min = 0
     ProgressBar1.Value = 0
     ProgressBar1.Max = REG&
     PROGRE& = 0
''''     SQLX$ = " SELECT A.NOMBRE_APELLIDO, C.VARIMOVI, A.LEGAJO AS LEGAJO, C.TIPOMOVIM, C.DEBE, C.NUMEDOC, C.FECHVENC, "
''''     SQLX$ = SQLX$ + " C.HABER, C.CODICOM_LAR, A.GRADO, A.AÑO, A.TURNO, A.LETRA, "
''''     SQLX$ = SQLX$ + " A.SALA_DE, C.FECHEMISI, A.MES_CUOTA, A.ANO_CUOTA, C.CUOTA_ORIGINAL, C.BONIFICACION, C.RECARGO, A.STATUS AS STATUS"
''''     SQLX$ = SQLX$ + " FROM ALUMNOS A INNER JOIN CAJABANC_COL C ON A.LEGAJO= C.LEGAJO "
''''     SQLX$ = SQLX$ + " Where A.LEGAJO > 0"
''''
''''     SQLX$ = SQLX$ + " AND (C.TIPOMOVIM = 'CUOTA' OR  C.TIPOMOVIM = 'MATRICULA')"
''''     If XVALO(Text1) > 0 Then SQLX$ = SQLX$ + " AND A.LEGAJO = " & LGJO&
''''     SQLX$ = SQLX$ + " AND (C.VARIMOVI = 'CUOTA' OR C.VARIMOVI = 'MATRICULA')"
     
 
     SQLX$ = "SELECT DEBE , HABER , CODICOM_LAR, "
     SQLX$ = SQLX$ + " FECHEMISI,MES_CUOTA, ANO_CUOTA, VARIMOVI, NUMEDOC, LEGAJO, NUMEDOC"
     SQLX$ = SQLX$ + " FROM ALUM_CAJAB"
     SQLX$ = SQLX$ + " WHERE (TIPOMOVIM = 'CUOTA' OR  TIPOMOVIM = 'MATRICULA')"
     If XVALO(Text1) > 0 Then SQLX$ = SQLX$ + " AND LEGAJO = " & LGJO&
     SQLX$ = SQLX$ + " AND (VARIMOVI = 'CUOTA' OR VARIMOVI = 'MATRICULA')"
     '
     'NUEVO
     If GRADOX% > 0 Then SQLX$ = SQLX$ + " AND GRADO = " & GRADOX%
     If SALADEX% > 0 Then SQLX$ = SQLX$ + " AND SALA_DE = " & SALADEX%
     If LETRX$ <> "" Then SQLX$ = SQLX$ + " AND LETRA = '" & LETRX$ & "'"
     If ANOX% > 0 Then SQLX$ = SQLX$ + " AND AÑO = " & ANOX%
     If ANOXMATRICULA% > 0 Then SQLX$ = SQLX$ + " AND ANO_CUOTA = " & ANOXMATRICULA%
     If TURNOX$ <> "" Then SQLX$ = SQLX$ + " AND TURNO = '" & TURNOX$ & "'"
     '
     FILA& = 0

     SQLX$ = SQLX$ + " AND FECHVENC <=  '" & HASTAX & "'"


     SQLX$ = SQLX$ + " order by ANO_CUOTA, MES_CUOTA, CODICOM_LAR "
     Screen.MousePointer = 11
     Set CTA_CTETMP = BASSSEX.Execute(SQLX$)
     With CTA_CTETMP
        Do While Not .EOF
         'BARRA DE PROGRESO
           PROGRE& = PROGRE& + 1
           ProgressBar1.Visible = True
           ProgressBar1.Value = PROGRE&
           Command2.Caption = SAFETEXT$(CLng((ProgressBar1.Value * 100) / ProgressBar1.Max) & " %")
           SkinLabel2 = SAFETEXT$(CLng((ProgressBar1.Value * 100) / ProgressBar1.Max) & " %")
           SkinLabel1 = Str$(PROGRE&) + " /" + Str(REG&)

           FEMI% = CVINT(!FECHEMISI)
           MEX% = XVALO(!MES_CUOTA)
           DOCUM$ = SAFETEXT$(!CODICOM_LAR)
           ANOX% = XVALO(!ANO_CUOTA)
           CONCEP = SAFETEXT$(!VARIMOVI)
           DEBEX = XVALO(!DEBE)
           PAGOALUMNO# = 0
           If Left$(DOCUM$, 3) = "CTA" Then
                PAGOALUMNO# = VECTRECI#(1, XVALO(Right(DOCUM$, 8)))
           ElseIf Left$(DOCUM$, 3) = "MTR" Then
                PAGOALUMNO# = VECTRECI#(2, XVALO(Right(DOCUM$, 8)))
           Else
             A = 1
           End If
          'PAGOALUMNO# = SUPAGOX#(DOCUM$) ' DEVUELVE LOS PAGOS REFERENCIADOS AL DOCUMENTO
          HABEX = PAGOALUMNO#
           SALDO = DEBEX - HABEX
           LGJO& = XVALO(!LEGAJO)
           If SALDO < 0.05 Then GoTo 50
           '
           FILA& = FILA& + 1
           MSF.Rows = FILA& + 1
           MSF.TextMatrix(FILA&, 0) = NombreAPe$(LGJO&)
           MSF.TextMatrix(FILA&, 1) = FECHATEX$(FEMI%)
           MSF.TextMatrix(FILA&, 2) = DOCUM$
           MESXX$ = "00" & TRIM$(Str(MEX%))
           MESXX$ = Right(MESXX$, 2)
           MSF.TextMatrix(FILA&, 3) = MESXX$
           MSF.TextMatrix(FILA&, 4) = SAFETEXT$(ANOX%)
           MSF.TextMatrix(FILA&, 5) = CONCEP
           MSF.TextMatrix(FILA&, 6) = FORMATNUM$(DEBEX, "F10.2")
           MSF.TextMatrix(FILA&, 7) = FORMATNUM$(HABEX, "F10.2")
           MSF.TextMatrix(FILA&, 8) = FORMATNUM$(SALDO, "F10.2")
           MSF.TextMatrix(FILA&, 9) = TRIM$(SAFETEXT$(LGJO&))
           '
           ACUMULA# = ACUMULA# + SALDO

50      .MoveNext
        Loop
     End With
     CTA_CTETMP.Close
     Set CTA_CTETMP = Nothing
     '
     'CARGO A PIE DE LA GRILLA EL TOTOL
     MSF.Rows = FILA& + 3
     MSF.TextMatrix(FILA& + 2, 7) = "TOTAL:"
     MSF.TextMatrix(FILA& + 2, 8) = FORMATNUM$(ACUMULA#, "F10.2")
     '
     'TOTAL A PIE DE FRM
     Text3 = FORMATNUM$(ACUMULA#, "F15.2")
     '
     Screen.MousePointer = 1
     Command2.Caption = "VER TODOS"
     MSF.Redraw = True
End Sub

Private Sub Text4_Change()
   Call CARGA_GRILLA_CTA_CTE
End Sub

Private Sub TXCOMBO2_Change()
   Call CARGA_GRILLA_CTA_CTE
End Sub

Private Sub TXCOMBO3_Change()
    Call CARGA_GRILLA_CTA_CTE
End Sub

Private Sub TXCOMBO5_Change()
   Call CARGA_GRILLA_CTA_CTE
End Sub

Private Sub TXCOMBO6_Change()
    Call CARGA_GRILLA_CTA_CTE
End Sub

Private Sub TXCOMBO7_Change()
    Call CARGA_GRILLA_CTA_CTE
End Sub
