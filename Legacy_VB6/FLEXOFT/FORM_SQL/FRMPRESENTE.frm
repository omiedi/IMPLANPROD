VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FRMPRESENTE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PRESENTES"
   ClientHeight    =   8685
   ClientLeft      =   -15
   ClientTop       =   675
   ClientWidth     =   15570
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   15570
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8640
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15555
      Begin VB.TextBox TXTAULA 
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
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox TXTDESDE 
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
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Grabar Presentes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12240
         TabIndex        =   24
         ToolTipText     =   "Grabar Presentes"
         Top             =   8160
         Width           =   3015
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
         Index           =   1
         Left            =   11040
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Fecha:"
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox TXTMOVIL 
         BackColor       =   &H0080FFFF&
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
         Left            =   0
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Costo segun el proveedor y la via de entrega"
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TXTCODAREA 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   600
         Width           =   675
      End
      Begin VB.TextBox TXTDESCRICURSO 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   960
         Width           =   7515
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
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Descirpción de Curso:"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox TXTDESCRIAREA 
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   600
         Width           =   6795
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Width           =   2595
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
         Left            =   4800
         TabIndex        =   12
         Top             =   240
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Código de Curso:"
         Top             =   240
         Width           =   1965
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "Area:"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox TXTMASID 
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
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame Frame2 
         Height          =   600
         Left            =   120
         TabIndex        =   4
         Top             =   3240
         Width           =   10695
         Begin VB.CommandButton Command43 
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
            Left            =   10200
            Picture         =   "FRMPRESENTE.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Refresh"
            Top             =   160
            Width           =   390
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
            Index           =   2
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Text            =   "Código Interno:"
            Top             =   170
            Width           =   1515
         End
         Begin VB.TextBox TXTIDBUSCAR 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   9360
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   170
            Width           =   765
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
            Left            =   720
            TabIndex        =   7
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
            Left            =   4440
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   160
            Width           =   3135
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FRMPRESENTE.frx":014A
            TabIndex        =   5
            Top             =   240
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   3480
            OleObjectBlob   =   "FRMPRESENTE.frx":01B8
            TabIndex        =   8
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox TXTCODINT 
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
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   1395
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
         Index           =   3
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "Código Interno:"
         Top             =   240
         Width           =   1515
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   1875
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Click Para Seleccionar hacia la Grilla Inferior"
         Top             =   1320
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   3307
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
      Begin MSFlexGridLib.MSFlexGrid MSF_2 
         Height          =   1755
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
         Top             =   1320
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   3096
         _Version        =   393216
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   12632319
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
      Begin MSFlexGridLib.MSFlexGrid MSF3 
         Height          =   4275
         Left            =   120
         TabIndex        =   20
         Top             =   3840
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   7541
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
      Begin MSComCtl2.DTPicker DTP1 
         Height          =   300
         Left            =   11865
         TabIndex        =   23
         Top             =   3360
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
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
         Format          =   131596289
         CurrentDate     =   41123
      End
   End
   Begin MSComDlg.CommonDialog CMD1 
      Left            =   7200
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   10815
      OleObjectBlob   =   "FRMPRESENTE.frx":0224
      Top             =   105
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuReportes 
      Caption         =   "Reportes"
      Begin VB.Menu mnuPresentes 
         Caption         =   "Listados de Presentes"
      End
   End
End
Attribute VB_Name = "FRMPRESENTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub GRABAR_PRESENTES()
   Dim DSX$(6)
   DSX$(0) = "DOMINGO"
   DSX$(1) = "LUNES"
   DSX$(2) = "MARTES"
   DSX$(3) = "MIERCOLES"
   DSX$(4) = "JUEVES"
   DSX$(5) = "VIERNES"
   DSX$(6) = "SABADO"

   US% = USNBR% Mod 100
   On Error GoTo ERROR_GUARDAR
   FLEXCONN.BeginTrans
   CONDI$ = "FECHLIS = '" & FETEXCOR1$(CVINT(DTP1.Value)) & "'"
   CONDI$ = CONDI$ + " AND CODINT = " & XVALO(Me.TXTMASID)
   Call BORRAREGISTROS("PRESENTES", CONDI$, REGAF&, "NOMESS")
   For i& = 1 To MSF3.Rows - 1
        NC& = XVALO(MSF3.TextMatrix(i&, 2))
        PRES = Asc(MSF3.TextMatrix(i&, 4))
        Observa$ = MSF3.TextMatrix(i&, 12)
        PRETE% = 0
        If PRES = 254 Then PRETE% = 1
        SQLX$ = "INSERT INTO PRESENTES "
        SQLX$ = SQLX$ + " (USERNUM,CODIEMPR,PRESENTE,CODINT,NUME_CLI,ANO,MESTEX,DIAATEX,FECHLIS,OBSERVACIONES,PRESTE,DESCRICURSO)"
        SQLX$ = SQLX$ + " VALUES(" & US%
        SQLX$ = SQLX$ + " ," & CodiEmp%
        SQLX$ = SQLX$ + " ," & PRES
        SQLX$ = SQLX$ + " ," & XVALO(Me.TXTMASID)
        SQLX$ = SQLX$ + " ," & NC&
        SQLX$ = SQLX$ + " ," & Year(Me.DTP1.Value)
        SQLX$ = SQLX$ + " ,'" & METEXTO$(CVINT(Me.DTP1.Value)) & "'"
        A1% = DIASEM%(CVINT(Me.DTP1.Value)): DIAX$ = DSX$(A1%)
        SQLX$ = SQLX$ + " ,'" & DIAX$ & "'"
        SQLX$ = SQLX$ + " ,'" & CVDAT(CVINT(Me.DTP1.Value)) & "'"
        SQLX$ = SQLX$ + " ,'" & Left$(Observa$, 256) & "'"
        SQLX$ = SQLX$ + " ," & PRETE% ' GUARDO PARA LOS LISTADOS 0 O 1 (AUSENTE O PRESENTE)
        SQLX$ = SQLX$ + " ,'" & Left$(TXTDESCRICURSO.Text, 256) & "'"
        SQLX$ = SQLX$ + ")"
        FLEXCONN.Execute (SQLX$)
        If PRETE% > 0 Then
             TXT1$ = " MARCADO COMO PRESENTE: " & FECHLIS & " - " & DIAATEX
             Call ANOTAX_ALUMNO(NC&, TXT1$)
        Else
        
        End If
   Next i&
   FLEXCONN.CommitTrans
   Exit Sub
ERROR_GUARDAR:
    If Err <> 0 Then
      FLEXCONN.RollbackTrans
      MsgBox "ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
      Screen.MousePointer = 1
      Exit Sub
    End If
End Sub

Private Sub CMBFILTRO_Change()
   Call TXTBUSCAR_Change
End Sub

Private Sub CMBFILTRO_Click()
   Call TXTBUSCAR_Change

End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
    '
    '
    Call ABRECONEXION
    '
    If CURSOSEL > 0 Then
      Text1 = XVALO(RESP_ZELE_VECT(1))
    End If
    '
99  Command1.Enabled = True
    '
    
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   If HOY(HO$) <> CVINT(Me.DTP1.Value) Then
      'NO PERMITE GUARDAR PRESENTES FUERA DE FECHA ACCION SOLO ADMITIDA CON DERECHOS.
      If DERACCE%("GrabaPres", "Modificación de Presentes Fuera de Rango") < 2 Then GoTo 99
   End If
   
   T1X$ = "Confirma Guardar la Lista de Presentes\Para el Curso: " & TXTDESCRICURSO.Text
   T1X$ = T1X$ + "\Horario: " & Me.TXTDESDE
   T1X$ = T1X$ + "\Aula: " & Me.TXTAULA
   T1X$ = T1X$ + "\\Cancelar\Guardar"
   OPX% = COMALTER%(T1X$)
   If OPX% < 2 Then GoTo 99
   '
   Call GRABAR_PRESENTES
   
   
99 Command3.Enabled = True
End Sub

Private Sub Command43_Click()
   Call TXTIDBUSCAR_Change
End Sub

Private Sub DTP1_Change()
        MSF3.Rows = 1
        TXTMOVIL(0) = ""
        TXTMOVIL(0).Visible = False
End Sub

Private Sub Form_Load()
    Campox$ = "F.Inicio/D8;F.Fin/D8;C/F1;Días/A10;C/F1;Días/A10;H.Desde/H10;H.Hasta/H10;Id Prof/F8;Profesor/A32;C.A./f4;Aula/A18;Val.Curso/f10.2;Val.Matrícula/F11.2;Cant.Meses/F10;Capacidad/F10;Ord/F0;CODINT/F7"
    Call CARGA_ENCABEZADO(Me.MSF, Campox$, Me, 1)
    Call CARGA_ENCABEZADO(Me.MSF_2, Campox$, Me, 0)

    CAMPOS$ = "CI1/F0;Cta./F5;Nombre/A24;Presente/f10;Dias/A20;Inicia/A10;Finaliza/A10;Profesor/A14;Celular/A14;Mail/A20;SALDO/F10.2;Observaciones/a128"
    MSF3.Font.Bold = True
    MSF3.Font.Name = "ARIAL"
    MSF3.Font.Size = 8
    MSF3.ScrollTrack = True
    MSF3.ScrollBars = flexScrollBarBoth
    MSF3.MergeCells = flexMergeFree
    MSF3.MergeCol(1) = True
    '
    Call CARGA_ENCABEZADO(MSF3, CAMPOS$, Me)
    Call INICIALIZAR_CHECKLIST_ENGRILLA(MSF3, 4, 0, 0)
    DTP1.Value = Date
    Me.Height = Screen.Height - 2000
    Me.Width = Screen.Width - 1000

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   MSF.Width = Me.Width - 300
   MSF_2.Width = Me.Width - 300
   
   MSF.Height = Me.Height - 1400
   MSF_2.Height = Me.Height - 1400
   
   MSF3.Width = Me.Width - 300
   MSF3.Height = Me.Height - 1400
   
   Me.Frame1.Width = Me.Width - 300


End Sub

Private Sub Form_Resize()
   Frame1.Width = Me.Width - 300
   Frame1.Height = Me.Height - 1000
   Me.MSF.Width = Me.Width - 500
   Me.MSF3.Width = Me.Width - 500
   Me.MSF_2.Width = Me.Width - 500
   Me.MSF3.Height = Me.Height - 5500
   Me.Command3.Top = MSF3.Top + MSF3.Height - 100
   AA = Me.Height
   
End Sub

Private Sub mnuPresentes_Click()
    If CURSOSEL > 0 Then
      CURSOPADRE& = XVALO(RESP_ZELE_VECT(1))
    End If

    Campox$ = "Curso/A64;F.Inicio/D8;F.Fin/D8;H.Desde/H12;H.Hasta/H12;C./F2"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campox$, Me)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campox$, Me)
    
    Dim obj As New RECORXET
    Set RS1 = obj.RS_CURSOS_ACT_2(CURSOPADRE&, 0, 1)
    REG& = 0
    
    Call Carga_MSF_Recordset(RS1, Campox$, FRMZELECHO.msf2, 1)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    FRMZELECHO.Width = 15000
    FRMZELECHO.Show 1

    CI1 = XVALO(RESP_ZELE_VECT(6))

   Call FINAL("?PRESENTES 2(" & CI1 & ")")
End Sub

Private Sub MSF_Click()
  If (MSF.MouseRow = 0) Or MSF.Rows <= 1 Then Exit Sub
  REG& = MSF.MouseRow
  If REG& < 1 Or MSF.Rows <= 1 Then Exit Sub
  MSF3.Rows = 1
  TXTMOVIL(0).Text = ""
  CI1 = XVALO(MSF.TextMatrix(REG&, 20))
  Desde$ = MSF.TextMatrix(REG&, 9)
  Hasta$ = MSF.TextMatrix(REG&, 10)
  LOSDIAS$ = TRIM$(MSF.TextMatrix(REG&, 4))
  DIA2$ = TRIM$(MSF.TextMatrix(REG&, 6))
  DIA3$ = TRIM$(MSF.TextMatrix(REG&, 8))
  
'  If DIAS2$ <> "" Then DIAS$ = DIAS$ + " Y " & DIAS2$
  LOSDIAS$ = DIA1$: If TRIM$(DIA2$) <> "" Then LOSDIAS$ = LOSDIAS$ & " Y " & DIA2$
  If TRIM$(DIA3$) <> "" Then LOSDIAS$ = DIA1$: If TRIM$(DIA2$) <> "" Then LOSDIAS$ = LOSDIAS$ & " , " & DIA2$ & " Y " & DIA3$
  
  PROFE$ = MSF.TextMatrix(REG&, 12)
  AULA1$ = MSF.TextMatrix(REG&, 14)

  TXTMASID = CI1
  TXTAULA = AULA1$
  TXTDESDE = Desde$
  TXTIDBUSCAR = ""
  SQLX$ = "SELECT D.Nume_Cli, D.Raso_Cli ,D.MAIL_CLI ,A.CURSOS_ID ,A.Descripcion,"
  SQLX$ = SQLX$ + "  D.CELU_CLI FROM DEUDOR12 D WITH(NOLOCK)"
  SQLX$ = SQLX$ + "  LEFT JOIN  ALUMNO_CURSO A WITH(NOLOCK) ON D.Nume_Cli = A.NUME_CLI"
  SQLX$ = SQLX$ + "  WHERE A.CURSOS_ID = " & CI1
  'SQLX$ = SQLX$ + "  AND P.FECHLIS = '" & FETEXCOR1$(CVINT(DTP1.Value)) & "'"
  SQLX$ = SQLX$ + "  ORDER BY D.NUME_CLI"
  'VER SACANDO LA FECHA Y TRAYENDO EL RESULTADO EN EL CAMPO CON UN VALOBADA
  Set rs = READSET(SQLX$)
  
  With rs
  Do While Not .EOF
    j& = j& + 1
    MSF3.Rows = j& + 1
    NC = XVALO(!NUME_CLI)
    MSF3.TextMatrix(j&, 1) = CI1
    MSF3.TextMatrix(j&, 2) = NC
    MSF3.TextMatrix(j&, 3) = SAFETEXT$(!RASO_CLI)
    EST% = PRESENTEALDIA(NC, CI1, CVINT(DTP1.Value)) 'ACA EL CHECK
    MSF3.TextMatrix(j&, 4) = Chr$(EST%) 'ACA EL CHECK
    If MSF3.TextMatrix(j&, 4) = Chr(254) Then
        Call COLOREAR_GRILLA_SOLACELDA(MSF3, &HC0FFC0, j&, 4)
        ESTATE% = 1
    Else
        ESTATE% = 0
    End If
    MSF3.TextMatrix(j&, 5) = DIAS$
    MSF3.TextMatrix(j&, 6) = Desde$
    MSF3.TextMatrix(j&, 7) = Hasta$
    MSF3.TextMatrix(j&, 8) = PROFE$
    MSF3.TextMatrix(j&, 9) = SAFETEXT$(!CELU_CLI)
    MSF3.TextMatrix(j&, 10) = SAFETEXT$(!MAIL_CLI)
    MSF3.TextMatrix(j&, 11) = FORMATNUM$(SALDOAFECHAX#(NC, HOY(HO$)), "G10.2") 'es el saldo a fecha
    MSF3.TextMatrix(j&, 12) = OBSERVACION_PRESENTES$(NC, CI1, CVINT(DTP1.Value))
    Call INICIALIZAR_CHECKLIST_ENGRILLA(MSF3, 4, j&, ESTATE%) '0 FALSE , 1 CHECK
    
    .MoveNext
  Loop
  End With
  rs.Close
  Set rs = Nothing
'  Call INICIALIZAR_CHECKLIST_ENGRILLA(MSF3, 4, 0, 0)

End Sub


Private Sub MSF3_DblClick()
  If MSF3.Rows <= 1 Then Exit Sub
  
  With MSF3
    ' verifica que el MSF3lexgrid contiene filas
    If .MouseRow > 0 And .MouseCol >= 0 Then
         LEGAJO& = MSF3.TextMatrix(.MouseRow, 2)
         ACONEC$ = "RECIBO10/RECICUTA" + SAFETEXT$(LEGAJO&) ' NUME
         Call CONECRUN(ACONEC$, "Generación de Recibo")
    End If
 End With
End Sub

Private Sub TXTBUSCAR_Change()
    If TRIM$(TXTBUSCAR) = "" Then
      MSF_2.Visible = False
      MSF.Visible = True
      Exit Sub
    End If
    
    TXTIDBUSCAR = ""
    '
    MSF_2.Rows = 2
    For i& = 1 To MSF_2.Cols - 1
       MSF_2.TextMatrix(1, i&) = ""
    Next i&
    REG& = 0
    MSF_2.Visible = True
    MSF_2.ZOrder 0
    MSF_2.Cols = MSF.Cols
    MSF_2.Redraw = False
    MSF.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For i& = 1 To MSF.Rows - 1
      For j& = 1 To MSF.Cols - 1
        If (InStr(UCase$(MSF.TextMatrix(i&, j&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.Text)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(MSF.TextMatrix(0, j&))) Then
            REG& = REG& + 1
            MSF_2.Rows = REG& + 1
            For H& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(REG&, H&) = MSF.TextMatrix(i&, H&)
            Next H&
          End If
          If (TRIM$(CMBFILTRO.Text)) = "" Then Exit For
        End If
      Next j&
    Next i&
    '
    MSF_2.Redraw = True

End Sub

Private Sub TXTIDBUSCAR_Change()
  Dim MS As MSFlexGrid
  Set MS = MSF
  If TRIM$(TXTBUSCAR.Text) <> "" Then
     Set MS = Me.MSF_2
  End If
  
  MASID& = XVALO(Me.TXTIDBUSCAR)
  For i& = 1 To MS.Rows - 1
     'PRIMERO DESAPLICA LA SELECCION (POR SI ESTA SELECCIONADO OTRA FILA)
     MS.Row = i&
     MS.RowSel = False
     'FIN DESAPLICACION DE SELECCION
     
     IDGRILLA& = XVALO((MS.TextMatrix(i&, 18)))
     If IDGRILLA& = MASID& Then 'SI LO ENCUENTRA LO SELECCIONA.
          Call ResaltarFila(MS, i&)
'          Exit For COMENTADO PARA QUE TERMINE DE LIMPIAR SI  HABIA ALGUNO SELECCIONADO
     End If
 Next i&
 TXTIDBUSCAR.SetFocus

End Sub

Private Sub TXTMOVIL_LostFocus(Index As Integer)
     Call TEXTOLOTES$(TXTMOVIL(Index).ToolTipText, ",")
     On Error Resume Next
     REG& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
     If TXTMOVIL(Index).Visible Then
         MSF3.TextMatrix(REG&, COLU1&) = TXTMOVIL(Index)
     End If

End Sub
Private Sub TXTMOVIL_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call TEXTOLOTES$(TXTMOVIL(Index).ToolTipText, ",")
     On Error Resume Next
     REG& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
     If TXTMOVIL(Index).Visible Then
        MSF3.TextMatrix(REG&, COLU1&) = TXTMOVIL(Index)
         TXTMOVIL(Index).Visible = False
     End If
     If COLU1& < MSF3.Cols - 1 Then
       COLU1& = COLU1& + 1
     End If
     If REG& < MSF3.Rows - 1 Then
          REG& = REG& + 1
       Else
          REG& = 1
     End If
     MSF3.Row = REG&
     MSF3.COL = COLU1&
     TXTMOVIL(Index).ToolTipText = REG& & "," & COLU1&

     '
     With MSF3
            TXTMOVIL(Index).Move .Left + .CellLeft, .Top + .CellTop, .CellWidth - 8, .CellHeight - 8
            TXTMOVIL(Index).Visible = True
            TXTMOVIL(Index).ZOrder 0
            TXTMOVIL(Index).Text = .TextMatrix(REG&, COLU1&)
            TXTMOVIL(Index).SetFocus
     End With

   End If
   
End Sub



Private Sub MSF3_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, Y As Single)
  If MSF3.Rows <= 1 Then Exit Sub
  
  With MSF3
    ' verifica que el MSF3lexgrid contiene filas
    If .MouseRow > 0 And .MouseCol >= 0 Then
      If .MouseCol <> 4 And .MouseCol < .Cols - 1 Then ' COLUMNA 4 ES LA DEL CHECK Y QUE SEA MENOR A LA ULTIMA P.Q. ES OBSERVAVIONES
        ' le asigna el dato con TextMatrix
        .ToolTipText = "Doble Click Para Pasar a Pantalla de Recibos"
      Else
        .ToolTipText = vbNullString
      End If
    Else
        ' No muestra ningún ToolTip
        .ToolTipText = vbNullString
          
    End If
  End With
    
End Sub

Private Sub MSF3_Click()
    If MSF3.Rows < 2 Then Exit Sub
    REG& = MSF3.MouseRow
    If REG& < 1 Then Exit Sub
    COL% = MSF3.MouseCol
    If MSF3.Row > MSF3.Rows - 1 Then Exit Sub
    If COL% = 4 Then
      If HOYSERVER% <> CVINT(Me.DTP1.Value) Then
          'NO PERMITE GUARDAR PRESENTES FUERA DE FECHA ACCION SOLO ADMITIDA CON DERECHOS.
          If DERACCE%("GrabaPres", "Modificación de Presentes Fuera de Rango") < 2 Then Exit Sub
      End If

      With MSF3
        .COL = COL%
        .Row = REG&
        .CellFontName = "Wingdings"
        .CellFontSize = 14
        .CellAlignment = flexAlignCenterCenter
        ' edita la celda
        If .TextMatrix(REG&, COL%) = Chr(254) Then
            .TextMatrix(REG&, COL%) = Chr(168) ' false
            Call COLOREAR_GRILLA_SOLACELDA(MSF3, vbWhite, REG&, COL%)
        Else
            .TextMatrix(REG&, COL%) = Chr(254) ' true
            Call COLOREAR_GRILLA_SOLACELDA(MSF3, &HC0FFC0, REG&, COL%)
        End If
        Call GRABAR_PRESENTES
        Exit Sub
      End With
    End If
    '
    If COL% = 12 Then
        TXTMOVIL(Index).ToolTipText = REG& & "," & COL%
        With MSF3
            .COL = COL%
            TXTMOVIL(Index).Move .Left + .CellLeft, .Top + .CellTop, .CellWidth - 8, .CellHeight - 8
            TXTMOVIL(Index).Visible = True
            TXTMOVIL(Index).ZOrder 0
            TXTMOVIL(Index).Text = .TextMatrix(REG&, COL%)
            TXTMOVIL(Index).SetFocus
        End With
        Exit Sub
    End If

End Sub
Private Sub MSF3_Scroll()
   If TXTMOVIL(Index).Visible = True Then
     REG& = XVALO(TXTMOVIL(Index).ToolTipText)
     MSF3.Text = TXTMOVIL(Index)
     TXTMOVIL(Index).Visible = False
   End If

End Sub


Private Sub Text1_Change()
        Call LIMPIAR
        Call CARGA_DATOS_CURSOS
        Call CARGAR_GRILLA_CURSOS

End Sub
Sub LIMPIAR()
            Me.TXTCODAREA = ""
            Me.TXTDESCRIAREA = ""
            Me.TXTDESCRICURSO = ""
            Me.MSF.Rows = 1
            Me.MSF3.Rows = 1
            Me.TXTAULA = ""
            Me.TXTMASID = ""
            Me.TXTDESDE = ""
            Call LIMPIAR_MSF_SIN_BORRAR_REGISTROS(MSF, 0)
            Me.TXTIDBUSCAR = ""
End Sub
Sub CARGA_DATOS_CURSOS()

   Curso& = XVALO(Text1)
   If Curso& <= 0 Then Exit Sub
   CONDI$ = "CURSOS_ID = " & Curso&
   If CantRegistros&("CURSOS", "CURSOS_ID = " & Curso&, "NOMESS") < 1 Then Exit Sub
   '
   SQLX$ = "SELECT * FROM CURSOS WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " + CONDI$
   Set rs = FLEXCONN.Execute(SQLX$)
   With rs
        Me.TXTCODAREA = SAFETEXT$(!COD_AREA)
        Me.TXTDESCRIAREA = SAFETEXT$(!AREA)
        Me.TXTDESCRICURSO = SAFETEXT$(!Descripcion)
    End With
    rs.Close
    Set rs = Nothing
End Sub

Sub CARGAR_GRILLA_CURSOS()
    
    Campox$ = "F.Inicio/D8;F.Fin/D8;C/A0;Días/A12;C/A0;Días/A12;C/A0;Días/A12;H.Desde/H12;H.Hasta/H12;Id Prof/F8;Profesor/A32;C.A./f4;Aula/A18;Val.Curso/f10.2;Val.Matrícula/F11.2;Cant.Meses/F10;Capacidad/F10;Ord/F0;CODINT/F7"
    Call CARGA_ENCABEZADO(Me.MSF, Campox$, Me)
    Dim obj As New RECORXET
    Set RS1 = obj.RS_CURSOS_ACT(XVALO(Text1), 0, 1)
    REG& = 0
    With RS1
    Do While Not .EOF
       REG& = REG& + 1
       MSF.Rows = REG& + 1
       MSF.TextMatrix(REG&, 1) = FECHATEX$(CVINT(!feinicurs))
       MSF.TextMatrix(REG&, 2) = FECHATEX$(CVINT(!FeFinCurs))
       MSF.TextMatrix(REG&, 3) = XVALO(!COD_DIA1)
       MSF.TextMatrix(REG&, 4) = SAFETEXT$(!desc_dia1)
       MSF.TextMatrix(REG&, 5) = XVALO(!COD_DIA2)
       MSF.TextMatrix(REG&, 6) = SAFETEXT$(!desc_dia2)
       MSF.TextMatrix(REG&, 7) = XVALO(!COD_DIA3)
       MSF.TextMatrix(REG&, 8) = SAFETEXT$(!desc_dia3)
        'ooooooo
       MSF.TextMatrix(REG&, 9) = SAFETEXT$(!H_DESDE)
       MSF.TextMatrix(REG&, 10) = SAFETEXT$(!H_HASTA)
       MSF.TextMatrix(REG&, 11) = XVALO(!ID_PROFESOR)
       MSF.TextMatrix(REG&, 12) = VALOBADA("PROFESORES", "RASO_CLI", "NUME_CLI = " & XVALO(!ID_PROFESOR))
       MSF.TextMatrix(REG&, 13) = XVALO(!ID_AULA)
       CONDI$ = "LISTA = 'AULAS' And CAMPO1 = '" & FORMATOCON0$(XVALO(!ID_AULA), 4) & "'"
       MSF.TextMatrix(REG&, 14) = VALOBADA("DATASQL", "CAMPO2", CONDI$, "MOMES")
       MSF.TextMatrix(REG&, 15) = FORMATNUM$(XVALO(TXTCURSOTOT), "F12.2")
       MSF.TextMatrix(REG&, 16) = FORMATNUM$(XVALO(TXTMATRICULA2), "F12.2")
       MSF.TextMatrix(REG&, 17) = XVALO(!CantMeses)
       MSF.TextMatrix(REG&, 18) = XVALO(!CAPACIDAD)
       MSF.TextMatrix(REG&, 19) = XVALO(!MASTER_ID)
       MSF.TextMatrix(REG&, 20) = XVALO(!CODINT)
      .MoveNext
    Loop
    End With
    RS1.Close
    Set RS1 = Nothing

End Sub
Function CURSOSEL()
    Campox$ = "C.Curso/F6;C.Area/F5;Area/A30;Desc.Curso/A58;Curso Cmp./F10.2;Matrícula/F10.2;Der. Examen/F10.2"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campox$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campox$, FRMZELECHO)
    Dim obj As New RECORXET
    Set RS1 = obj.RS_CURSOS()
    Call Carga_MSF_Recordset(RS1, Campox, FRMZELECHO.msf2, 1)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    FRMZELECHO.Width = 15000
    FRMZELECHO.Show 1

    NumeroCliente$ = TRIM(RESP_ZELE_VECT(1))
    CURSOSEL = XVALO(NumeroCliente$)

End Function
    Public Function RS_CURSOS() As ADODB.Recordset
      Dim RS_FUN As ADODB.Recordset
      Set RS_FUN = New ADODB.Recordset
      SQLX$ = "SELECT CURSOS_ID,COD_AREA, AREA,DESCRIPCION,ValCurso,ValMatricula,ValDerExamen"
      SQLX$ = SQLX$ + " FROM CURSOS WITH(NOLOCK) ORDER BY COD_AREA,DESCRIPCION ASC"
      Set RS_FUN = READSET(SQLX$)
      Set RS_CURSOS = RS_FUN
      Set RS_FUN = Nothing
    End Function

