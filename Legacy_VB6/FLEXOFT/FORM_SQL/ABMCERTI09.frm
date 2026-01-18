VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ABMCERTI09 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ABM de Certificados x Orden de Fabricación o Boleta de Recepción"
   ClientHeight    =   8790
   ClientLeft      =   75
   ClientTop       =   285
   ClientWidth     =   12330
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8790
   ScaleWidth      =   12330
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4920
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "A.B.M de Certificados x Boleta de Recepción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   12
      Top             =   4440
      Width           =   12135
      Begin VB.CommandButton Command62 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3480
         Picture         =   "ABMCERTI09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Tabla de Títulos de Selección"
         Top             =   240
         Width           =   550
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Bindings        =   "ABMCERTI09.frx":014A
         Height          =   3050
         Left            =   165
         TabIndex        =   13
         Top             =   1200
         Width           =   11820
         _ExtentX        =   20849
         _ExtentY        =   5371
         _Version        =   393216
         BackColor       =   16777215
         FocusRect       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton Command8 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   11325
         Picture         =   "ABMCERTI09.frx":015E
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Cargar Items"
         Top             =   240
         Width           =   645
      End
      Begin VB.CommandButton Command5 
         Caption         =   "B.R."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2760
         TabIndex        =   17
         Top             =   240
         Width           =   520
      End
      Begin VB.Frame Frame5 
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
         Height          =   825
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2475
         Begin VB.TextBox Text12 
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
            Height          =   285
            Left            =   105
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   "B.Recepción"
            Top             =   150
            Width           =   2235
         End
         Begin VB.TextBox Text6 
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
            Left            =   105
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   2235
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "A.B.M. de Certificados x Orden de Fabricación "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   12135
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
         Height          =   795
         Left            =   8640
         Picture         =   "ABMCERTI09.frx":0468
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Actualizar Indice de Selección de O.Fab."
         Top             =   240
         Width           =   525
      End
      Begin VB.CommandButton COMMAND3 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   11205
         Picture         =   "ABMCERTI09.frx":05B2
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Cargar Items"
         Top             =   240
         Width           =   735
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
         Height          =   795
         Left            =   7920
         TabIndex        =   11
         Top             =   240
         Width           =   520
      End
      Begin VB.Frame Frame5 
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
         Height          =   825
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   200
         Width           =   7695
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "Descripción"
            Top             =   150
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
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
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
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   420
            Width           =   1140
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
            TabIndex        =   7
            Top             =   420
            Width           =   2235
         End
         Begin VB.TextBox Text13 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Text            =   "Código"
            Top             =   150
            Width           =   2235
         End
         Begin VB.TextBox Text14 
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
            Height          =   285
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Text            =   "Ord.Fabric."
            Top             =   150
            Width           =   1140
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Bindings        =   "ABMCERTI09.frx":08BC
         Height          =   3050
         Left            =   165
         TabIndex        =   20
         Top             =   1080
         Width           =   11820
         _ExtentX        =   20849
         _ExtentY        =   5371
         _Version        =   393216
         BackColor       =   16777215
         FocusRect       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   0
      Top             =   0
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
      OleObjectBlob   =   "ABMCERTI09.frx":08D0
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "ABMCERTI09.frx":0B04
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
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
End
Attribute VB_Name = "ABMCERTI09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click(Index As Integer)
      
End Sub

Private Sub Command3_Click()

       Command3.Enabled = True
       Screen.MousePointer = 11
       CODI$ = Text11
       CI1% = CODINT%(CODI$)
       DESCRI$ = Text4
       LOTE$ = Text10
       NUMERSE$ = Text15
       '
       'MARCO MARBORRA 1
       CONDI$ = " CODINT = " & CI1%
       CONDI$ = CONDI$ + " AND IDENLOTE = '" & LOTE$ & "'"
       Call ACTUREGISTRO("CERTIFICADOS", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
       '
       SQLX$ = "SELECT * FROM CERTIFICADOS "
       SQLX$ = SQLX$ + " WHERE CODINT < 1 "
       
       Dim RS As ADODB.Recordset
       Set RS = New ADODB.Recordset
25     On Error Resume Next
       RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText     ' SE MODIFICO EL adOpenDynamic POR EL adOpenKeyset
       '
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
       End If
       On Error GoTo 0
       '
       'ALTA EN TABLA CERTIFICADOS
       With RS
         For i& = 1 To MSF.Rows - 1
            .AddNew
            !CODINT = CI1%
            !Cod_Exte = CODI$
            !Descripcion = DESCRI$
            !Status = 0
            !idenlote = LOTE$
            !NUMEROSERIE = MSF.TextMatrix(i&, 4)
            !RUT_ARCHIVO = MSF.TextMatrix(i&, 5)
            !TIT_CERTIFICADO = XVALO(MSF.TextMatrix(i&, 6))
            !HOJAS = XVALO(MSF.TextMatrix(i&, 7))
            !MARBORRA = 0
            .Update
          Next i&
        End With
        '
        CONDI$ = CONDI$ + " AND MARBORRA = 1"
        Call BORRAREGISTROS("CERTIFICADOS", CONDI$, REGAF&, "NOMESS")
        '
        '
        'RECORRO LA GRILLA
        'SI YA TIENE REGISTRO EN EL TANUMSE Y CAMBIO EL N.SERIE ACTUALIZA EL MISMO
        'SI NO EXISTE LA PRIMERA VEZ VA A SER EL PROVISORIO(OF-012345-1-1) EN ESE CASO AGREGA
        'UN NUEVO REGISTRO Y TOMA EL NUMERO DE ORDEN LUEGO DEL ULTIMO GUION
        For i& = 1 To MSF.Rows - 1
            NUMERO_SERIE$ = MSF.TextMatrix(i&, 4)
            NUMERO_SERIE_ANT$ = MSF.TextMatrix(i&, 8)
            CONDI$ = "NUMEROSERIE = '" & NUMERO_SERIE_ANT$ & "' AND IDSUBOR = '" & LOTE$ & "'"
            CANTREG& = CantRegistros&("TANUMSE", CONDI$, "NOMESS")
            If CANTREG& > 0 Then
                If NUMERO_SERIE_ANT$ <> NUMERO_SERIE$ Then
                   Call ACTUREGISTRO("TANUMSE", "NUMEROSERIE", CONDI$, NUMERO_SERIE$, REGAF&, "NOMESS")
                   'ACTUALIZO TABLA CERTI_ESTRUC
                   Condi1$ = "IDENLOTE = '" & NUMERO_SERIE_ANT$ & "' AND ORD_FABRIC_GRAL = '" & LOTE$ & "'"
                   Call ACTUREGISTRO("CERTI_ESTRUC", "IDENLOTE", Condi1$, NUMERO_SERIE$, REGAF&, "NOMESS")
                End If
            Else
                CONDI$ = "NUMEROSERIE = '" & NUMERO_SERIE$ & "' AND IDSUBOR = '" & LOTE$ & "'"
                CANTREG& = CantRegistros&("TANUMSE", CONDI$, "NOMESS")
                If CANTREG& < 1 Then
                    Dim rs1 As ADODB.Recordset
                    Set rs1 = New ADODB.Recordset
26                  On Error Resume Next
                    SQLX$ = "SELECT * FROM TANUMSE WHERE COD_INTE < 1 "
                    rs1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText     ' SE MODIFICO EL adOpenDynamic POR EL adOpenKeyset
                    '
                    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                       On Error GoTo 0
                       Call CapturaErrorOpen
                       GoTo 26
                    End If
                    On Error GoTo 0
                    '
                    'VER AHORA EL TEMA DEL NUMERO DE ORDEN SI ESTA BIEN ASI
                    With rs1
                        .AddNew
                        If Left$(NUMERO_SERIE_ANT$, 3) = "OF-" Then
                          POS1% = InStrRev(NUMERO_SERIE_ANT$, "-")
                          NumOrden% = Mid$(NUMERO_SERIE_ANT$, POS1% + 1)
                        End If
                        .AddNew   ' YA CONVERTIDO
                        On Error Resume Next
                        !CODIEMPR = CodiEmp%
                        On Error GoTo 0
                        !idsubor = LOTE$
                           HOII% = HOY(HOS$)
                           FF% = HOII%
                        !FechAlta = CVDAT(FF%)
                        !cod_inte = CI1%
                        !Cod_Exte = CODI$
                        !Descrip = DESCRI$
                        !MESANIO = ""
                        !Ten_Ensa = 0
                            NUEVONUMSERIE$ = NORFA$ & "-" & SAFETEXT$(NumOrden%)
                        !NUMEROSERIE = NUMERO_SERIE$
                        !NUMEORD = NumOrden%
                        !DEPONUMSE = (-1)
                        'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
                        CUERPO$ = "Alta N/Serie a Partir de: " & NORFA$
                        HISTO$ = FORMATEA_ANOTADOR$(CUERPO$)
                        !Historial = HISTO$
                        .Update
                    End With
                    rs1.Close
                    Set rs1 = Nothing
                End If
            End If
          Next i&

99      Command3.Enabled = True
        Screen.MousePointer = 1
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    '
    
    Call SELECHO("OFABNOAN/Indice General de Ordenes de Compra (Pendientes-Cumplidas)")
    NORFA$ = TRIM$(VALACT1$("OFABNOAN"))
    If NORFA$ = "" Then GoTo 99
    '
    CI1% = CODIFAB%(NORFA$)
    
'    If ESTRAZABLE%(CI1%) < 2 Then
'    If TRIM$(CODFAMIL$(CI1%)) = "" Then
'       Call MENSERR(24, "Código: " & CODEXT(CI1%) & " No Tiene Asignado Familia")
'    End If
    '
    Text10 = NORFA$
    Text11 = CODEXT(CI1%)
    Text4 = DESCRIT$(CI1%)
    '
'    Call CARGAMSNSER(NORFA$)
    
99  Screen.MousePointer = 1
 
    Command2.Enabled = True
End Sub
Sub ACTUIN_OFABNOAN()
    CONDI$ = "IDSUBOR <> '' AND StatuOrf <> 9 ORDER BY NUMEORFA"
    Call CARGALISEL("OFABNOAN", "ORDEFABR", "IDSUBOR/A12;COD_EXTE/A16;DESCRIP/A48;CanProy/F10.1", CONDI$, "NOMESS")
End Sub

Sub CARGAMSNSER(NORFA$)
    '
    MSF.Rows = 1
    MSF.Redraw = False
    '
    SQLX$ = "SELECT IDSUBOR,CANPROY,COD_INTE,Cod_Exte, DESCRIP FROM ORDEFABR WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "'"
    '
    Set RS = READSET(SQLX$)
    '
    i& = 0
    With RS
      For ORD% = 1 To XVALO(!CANPROY)
'        Do While Not .EOF
          i& = i& + 1
            CI1% = XVALO(!cod_inte)

            MSF.Rows = MSF.Rows + 1
            MSF.TextMatrix(i&, 1) = CODEXT$(CI1%)
            MSF.TextMatrix(i&, 2) = DESCRIT0$(CI1%)
               NORFA$ = SAFETEXT$(!idsubor)
            MSF.TextMatrix(i&, 3) = NORFA$
                CONDI$ = "IDSUBOR= '" & NORFA$ & "' And NUMEORD = " & ORD%
                NUSE$ = TRIM$(VALOBADA("TANUMSE", "NUMEROSERIE", CONDI$, "NOMESS"))
                If NUSE$ = "" Then NUSE$ = NORFA$ & "-" & SAFETEXT$(ORD%)
            MSF.TextMatrix(i&, 4) = NUSE$
            RUT$ = "" 'POR DEFECTO LA RUTA LA ESCRIBE VACIA
            MSF.TextMatrix(i&, 5) = SAFETEXT$(RUT$)
            MSF.TextMatrix(i&, 8) = NUSE$
            If NUSEX_ANT$ = NUSE$ Then GoTo 80

            SQLX$ = "SELECT * FROM CERTIFICADOS WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE IDENLOTE = '" & NORFA$ & "' AND NUMEROSERIE = '" & NUSE$ & "'"
            INGRERUTA% = 0
            'CONTROLA SI TIENEN ASIGNADA RUTA EN TABLA CERTIFICADOS
            Set rs1 = READSET(SQLX$)
            With rs1
               Do While Not .EOF
                 RUT$ = TRIM$(SAFETEXT$(!RUT_ARCHIVO))
                 ORDENX% = XVALO(TRIM$(SAFETEXT$(!TIT_CERTIFICADO)))
                 NUSEX$ = TRIM$(SAFETEXT$(!NUMEROSERIE))
                 If RUT$ <> "" Or ORDENX% > 0 Then
                    'SI TIENE RUTA O TITULO DE CERTIFICADO AL PRIMERO LE AGREGA LA RUTA
                    If INGRERUTA% < 1 Then
                       MSF.TextMatrix(i&, 5) = SAFETEXT$(RUT$)
                       MSF.TextMatrix(i&, 6) = SAFETEXT$(!TIT_CERTIFICADO)
                       MSF.TextMatrix(i&, 7) = SAFETEXT$(!HOJAS)
                       INGRERUTA% = 1
                    Else
                       'SI ENTRA ACA HAY MAS DE UNO Y VA GENERANDO NUVOS REGISTROS.
                       i& = i& + 1
                       MSF.Rows = MSF.Rows + 1
                       MSF.TextMatrix(i&, 1) = SAFETEXT$(!Cod_Exte)
                       MSF.TextMatrix(i&, 2) = SAFETEXT$(!Descripcion)
                       MSF.TextMatrix(i&, 3) = SAFETEXT$(!idenlote)
                       MSF.TextMatrix(i&, 4) = SAFETEXT$(!NUMEROSERIE)
                       MSF.TextMatrix(i&, 5) = SAFETEXT$(RUT$)
                       MSF.TextMatrix(i&, 6) = SAFETEXT$(!TIT_CERTIFICADO)
                       MSF.TextMatrix(i&, 7) = SAFETEXT$(!HOJAS)
                    End If
                 End If
               .MoveNext
               Loop
            End With
'            .MoveNext
'        Loop
       NUSEX_ANT$ = NUSE$

80     Next ORD%
    End With
          
    
    On Error Resume Next
    RS.Close
    Set RS = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
    '
''''    CANPRONORFA& = XVALO(VALOBADA("ORDEFABR", "CANPROY", "IDSUBOR = '" & NORFA$ & "'", "NOMESS"))
''''    '
''''    If CANNUSE& < CANPRONORFA& Then
''''          I& = I& + 1
''''          MSF.Rows = MSF.Rows + 1
''''          CI1% = XVALO(VALOBADA("ORDEFABR", "COD_INTE", CONDI$, "NOMESS"))
''''          MSF.TextMatrix(I&, 1) = CODEXT$(CI1%)
''''          MSF.TextMatrix(I&, 2) = DESCRIT0$(CI1%)
''''          MSF.TextMatrix(I&, 3) = NORFA$
''''          NUSE$ = ""
''''          MSF.TextMatrix(I&, 4) = NUSE$
''''          '
''''          RUT$ = "" 'POR DEFECTO LA RUTA LA ESCRIBE VACIA
''''          MSF.TextMatrix(I&, 5) = RUT$
''''
''''          SQLX$ = "SELECT * FROM CERTIFICADOS WITH(NOLOCK)"
''''          SQLX$ = SQLX$ + " WHERE IDENLOTE = '" & NORFA$ & "' "
''''          INGRERUTA% = 0
''''          'CONTROLA SI TIENEN ASIGNADA RUTA EN TABLA CERTIFICADOS
''''          Set RS1 = READSET(SQLX$)
''''          With RS1
''''             Do While Not .EOF
''''               RUT$ = TRIM$(SAFETEXT$(!RUT_ARCHIVO))
''''               ORDENX% = XVALO(TRIM$(SAFETEXT$(!TIT_CERTIFICADO)))
''''               NUSEX$ = TRIM$(SAFETEXT$(!NUMEROSERIE))
''''
''''               If RUT$ <> "" Or ORDENX% > 0 Then
''''                  'SI TIENE RUTA O TITULO DE CERTIFICADO AL PRIMERO LE AGREGA LA RUTA
''''                  If INGRERUTA% < 1 Then
''''                     MSF.TextMatrix(I&, 5) = SAFETEXT$(RUT$)
''''                     MSF.TextMatrix(I&, 6) = SAFETEXT$(!TIT_CERTIFICADO)
''''                     MSF.TextMatrix(I&, 7) = SAFETEXT$(!HOJAS)
''''                     INGRERUTA% = 1
''''                  Else
''''                     'SI ENTRA ACA HAY MAS DE UNO Y VA GENERANDO NUVOS REGISTROS.
''''                     I& = I& + 1
''''                     MSF.Rows = MSF.Rows + 1
''''                     MSF.TextMatrix(I&, 1) = SAFETEXT$(!COD_EXTE)
''''                     MSF.TextMatrix(I&, 2) = SAFETEXT$(!Descripcion)
''''                     MSF.TextMatrix(I&, 3) = SAFETEXT$(!IDENLOTE)
''''                     MSF.TextMatrix(I&, 4) = SAFETEXT$(!NUMEROSERIE)
''''                     MSF.TextMatrix(I&, 5) = SAFETEXT$(RUT$)
''''                     MSF.TextMatrix(I&, 6) = SAFETEXT$(!TIT_CERTIFICADO)
''''                     MSF.TextMatrix(I&, 7) = SAFETEXT$(!HOJAS)
''''                  End If
''''               End If
''''               .MoveNext
''''             Loop
''''          End With
''''    End If
    MSF.COL = 4
    MSF.Sort = 1

    MSF.Redraw = True

99  Command2.Enabled = True
End Sub

Sub CARGAMSFBR(BOLRE)
    MSF1.Rows = 1
    MSF1.Redraw = False
    POS1% = InStr(BOLRE, "-")
    BRE = Mid$(BOLRE, POS1% + 1)
    If InStr(BRE, "-") > 0 Then
      LOTE$ = BOLRE
      CONDI$ = " IDENLOTE  = '" & LOTE$ & "'"
    Else
      LOTE$ = BOLRE & "%"
      CONDI$ = " IDENLOTE  LIKE '" & LOTE$ & "'"
    End If
    '
    SQLX$ = "SELECT COD_EXTE, DESCRIP, IDENLOTE,COD_INTE  FROM BOLRECEP WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE   " & CONDI$
    SQLX$ = SQLX$ + " ORDER BY IDENLOTE"
    Set RS = READSET(SQLX$)
    
    i& = 0
    'MODIFICAR LOS CAMPOS SEGUN
    With RS
        Do While Not .EOF
          i& = i& + 1
          MSF1.Rows = MSF1.Rows + 1
          MSF1.TextMatrix(i&, 1) = SAFETEXT$(!Cod_Exte)
          MSF1.TextMatrix(i&, 2) = SAFETEXT$(!Descrip)
          LOTE$ = TRIM$(SAFETEXT$(!idenlote))
          MSF1.TextMatrix(i&, 3) = LOTE$
          RUT$ = "" 'POR DEFECTO LA RUTA LA ESCRIBE VACIA
          MSF1.TextMatrix(i&, 4) = RUT$
          CI1% = XVALO(!cod_inte)
          SQLX$ = "SELECT * FROM CERTIFICADOS WITH(NOLOCK)"
          SQLX$ = SQLX$ + " WHERE IDENLOTE = '" & LOTE$ & "' AND CODINT = " & CI1%
          INGRERUTA% = 0
          'CONTROLA SI TIENEN ASIGNADA RUTA EN TABLA CERTIFICADOS
          Set rs1 = READSET(SQLX$)
          With rs1
             Do While Not .EOF
               RUT$ = TRIM$(SAFETEXT$(!RUT_ARCHIVO))
               If RUT$ <> "" Then
                  'SI TIENE RUTA AL PRIMERO LE AGREGA LA RUTA
                  If INGRERUTA% < 1 Then
                     MSF1.TextMatrix(i&, 4) = SAFETEXT$(RUT$)
                     MSF1.TextMatrix(i&, 5) = SAFETEXT$(!TIT_CERTIFICADO)
                     MSF1.TextMatrix(i&, 6) = SAFETEXT$(!HOJAS)

                     INGRERUTA% = 1
                  Else
                     'SI ENTRA ACA HAY MAS DE UNO Y VA GENERANDO NUVOS REGISTROS.
                     i& = i& + 1
                     MSF1.Rows = MSF1.Rows + 1
                     MSF1.TextMatrix(i&, 1) = SAFETEXT$(!Cod_Exte)
                     MSF1.TextMatrix(i&, 2) = SAFETEXT$(!Descripcion)
                     MSF1.TextMatrix(i&, 3) = SAFETEXT$(!idenlote)
                     MSF1.TextMatrix(i&, 4) = SAFETEXT$(RUT$)
                     MSF1.TextMatrix(i&, 5) = SAFETEXT$(!TIT_CERTIFICADO)
                     MSF1.TextMatrix(i&, 6) = SAFETEXT$(!HOJAS)

                  End If
               End If
               .MoveNext
             Loop
          End With
          
          .MoveNext
        Loop
    End With
    On Error Resume Next
    RS.Close
    Set RS = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
    Call ORDENAR_MSF(ABMCERTI09.MSF, 3, 1)
    MSF1.Redraw = True

99  Command2.Enabled = True


End Sub

Private Sub Command4_Click()
   Text1 = ""
End Sub

Sub Command43_Click()
   Command43.Enabled = False
   Call ACTUIN_OFABNOAN
   Command43.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
3  Call SELECHO("INDIBOL")
   BOLRE = Val(VALACT1$("INDIBOL"))
   If BOLRE < 1 Then GoTo 99
   '
   Text6 = "BR-" & BOLRE

   Screen.MousePointer = 11
99 Screen.MousePointer = 1
   Command5.Enabled = True

End Sub




Private Sub Command62_Click()
    Command62.Enabled = False
    '
    Call TRALOCAL("TITUCERT", "")
    TITUPA$ = "Carga de Títulos de Certificados"
    VTB% = VENTABU%("TITUCERT/TITUPAN=" + TITUPA$)
    If VTB% < 0 Then GoTo 99
    '
    J& = 0
    For U& = 1 To ULTREG&("!TITUCERT")
        X$ = REGLEIDO$("!TITUCERT", U&)
        ORDEX% = CVI(Left$(X$, 2))
        If ORDEX% >= 1 Then
            J& = J& + 1
            Call GRAREG("TITUCERT", X$, J&)
        End If
    Next U&
    '
    Call SETULTREG("TITUCERT", J&)
    Call CIERRARCH("TITUCERT")
    Call CIERRARCH("!TITUCERT")
    Call FRESHECHO("TITUCERT")
    '
99  Command62.Enabled = True

End Sub

Private Sub Command8_Click()
      Command8.Enabled = True
      Screen.MousePointer = 11
      LOTE$ = Text6
      LOTE$ = LOTE$ & "%"
      '
      CONDI$ = " idenlote LIKE '" & LOTE$ & "'"
      Call ACTUREGISTRO("CERTIFICADOS", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
      
      SQLX$ = " SELECT * FROM CERTIFICADOS WHERE CODINT < 1"
      Screen.MousePointer = 11
      Dim RS As ADODB.Recordset
      Set RS = New ADODB.Recordset
25    On Error Resume Next
      RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText     ' SE MODIFICO EL adOpenDynamic POR EL adOpenKeyset
      '
      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
      End If
      On Error GoTo 0
      '
      With RS
         For i& = 1 To MSF1.Rows - 1
            .AddNew
            CI1% = CODINT%(MSF1.TextMatrix(i&, 1))
            !CODINT = CI1%
            !Cod_Exte = MSF1.TextMatrix(i&, 1)
            !Descripcion = MSF1.TextMatrix(i&, 2)
            !Status = 0
            !idenlote = MSF1.TextMatrix(i&, 3)
            !NUMEROSERIE = ""
            !RUT_ARCHIVO = MSF1.TextMatrix(i&, 4)
            !TIT_CERTIFICADO = XVALO(MSF1.TextMatrix(i&, 5))
            !familia = TRIM$(CODFAMIL$(CI1%))
            !HOJAS = XVALO(MSF1.TextMatrix(i&, 6))
            !MARBORRA = 0
            .Update
          Next i&
      End With

      RS.Close
      Set RS = Nothing
              
      CONDI$ = CONDI$ + " AND MARBORRA = 1"
      Call BORRAREGISTROS("CERTIFICADOS", CONDI$, REGAF&, "NOMESS")

99    Command8.Enabled = True
      Screen.MousePointer = 1
End Sub

Private Sub Form_Load()
        
        Call ABRECONEXION
        Call CREACAMPO("", "CERTIFICADOS", "CODINT", 3, 0, 1)
        Call CREACAMPO("", "CERTIFICADOS", "COD_EXTE", 10, 20)
        Call CREACAMPO("", "CERTIFICADOS", "DESCRIPCION", 10, 64)
        Call CREACAMPO("", "CERTIFICADOS", "STATUS", 3, 0)
        Call CREACAMPO("", "CERTIFICADOS", "IDENLOTE", 10, 32, 1)
        Call CREACAMPO("", "CERTIFICADOS", "NUMEROSERIE", 10, 32, 1)
        Call CREACAMPO("", "CERTIFICADOS", "RUT_ARCHIVO", 12, 0)
        Call CREACAMPO("", "CERTIFICADOS", "MARBORRA", 3, 0, 1)
        Call CREACAMPO("", "CERTIFICADOS", "TIT_CERTIFICADO", 3, 0, 1)
        Call CREACAMPO("", "CERTIFICADOS", "Familia", 10, 5)
        Call CREACAMPO("", "CERTIFICADOS", "HOJAS", 10, 5)
        
        Call CARGA_ENCABEZADO(MSF, "Código/A16;Descripción/A40;O.Fabricación./A14;N.Serie/A16;Link Certificado/A48;C.T./I5;Hojas/A5;NUSEANT/A-2")
        Call CARGA_ENCABEZADO(MSF1, "Código/A16;Descripción/A40;Lote;Link Certificado/A48;C.T./I5;Hojas/A5")
        MSF.MergeCells = flexMergeRestrictColumns
        MSF.MergeCol(1) = True
        MSF.MergeCol(4) = True
        MSF1.MergeCells = flexMergeRestrictColumns
        MSF1.MergeCol(1) = True
        MSF1.MergeCol(3) = True
        '
        '
        End Sub

Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame6_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub MSF_DblClick()
       i& = MSF.MouseRow
       J& = MSF.MouseCol
       If i& = 0 Or i& > MSF.Rows Then Exit Sub
       '
       On Error Resume Next
       MSF.Row = i&
       MSF.COL = J&
       On Error Resume Next

       LOTE$ = TRIM$(MSF.TextMatrix(i&, 3))
       If LOTE$ = "" Then
           Exit Sub
       Else
           ASIGNARUTA.Text14 = "O.Fab."
           ASIGNARUTA.Text9(5) = "N.Serie"
           ASIGNARUTA.Text15 = "Cantidad"
           ASIGNARUTA.Text15.Alignment = 0
           ASIGNARUTA.Text10.Alignment = 0
           ASIGNARUTA.Text11 = MSF.TextMatrix(i&, 1)
           ASIGNARUTA.Text4 = MSF.TextMatrix(i&, 2)
           ASIGNARUTA.Text10 = MSF.TextMatrix(i&, 3)
           ASIGNARUTA.Text15 = MSF.TextMatrix(i&, 4)
           ASIGNARUTA.Text1 = MSF.TextMatrix(i&, 5)
           ASIGNARUTA.Text3 = ECOARCH$("TITUCERT", MKI$(XVALO(MSF.TextMatrix(i&, 6))))
           ASIGNARUTA.Text5 = XVALO(MSF.TextMatrix(i&, 7))
           ORDEX% = 0
           ASIGNARUTA.Show 1
       End If
End Sub


Private Sub MSF1_DblClick()
      i& = MSF1.MouseRow
      J& = MSF1.MouseCol
      If i& = 0 Or i& > MSF1.Rows Then Exit Sub
      '
      
      On Error Resume Next
      MSF1.Row = i&
      MSF1.COL = J&
      On Error GoTo 0
      LOTE$ = TRIM$(MSF1.TextMatrix(i&, 3))
      If LOTE$ = "" Then
           Exit Sub
      Else
           ASIGNARUTA.Text14 = "Lote"
           ASIGNARUTA.Text9(5).Visible = False
           ASIGNARUTA.Text15.Visible = False
           ASIGNARUTA.Text10.Alignment = 0
           ASIGNARUTA.Text11 = MSF1.TextMatrix(i&, 1)
           ASIGNARUTA.Text4 = MSF1.TextMatrix(i&, 2)
           ASIGNARUTA.Text10 = MSF1.TextMatrix(i&, 3)
           ASIGNARUTA.Text1 = MSF1.TextMatrix(i&, 4)
           ASIGNARUTA.Text3 = ECOARCH$("TITUCERT", MKI$(XVALO(MSF1.TextMatrix(i&, 5))))
           ASIGNARUTA.ORDEX% = XVALO(MSF1.TextMatrix(i&, 5))
           ASIGNARUTA.Text5 = XVALO(MSF1.TextMatrix(i&, 6)) ' cantidad de hojas
           ASIGNARUTA.Show 1
      End If
      
       
End Sub
'Private Sub MSF1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'          '
'          If MSF1.Col <> 3 Then
'            Command4.Visible = False
'            Exit Sub
'          End If
'          MSF1.Col = 3
'          With Command4
'            .ZOrder 0
''            .Width = MSF1.CellWidth
'            .Left = (MSF1.CellLeft + MSF1.Left + MSF1.CellWidth) - .Width
'            .Top = MSF1.CellTop + MSF1.Top
''            .TEXT = MSF1.TEXT
'            .Visible = True
'          End With
'          '
'End Sub
'      Private Sub MSF1_MouseUp(Button As Integer, _
'         Shift As Integer, x As Single, y As Single)
''          Static CurrentWidth As Single
''          If MSF1.Rows <= 1 Then Exit Sub
''          If MSF1.CellWidth <> CurrentWidth Then
''              Command4.Width = MSF1.CellWidth
''              CurrentWidth = MSF1.CellWidth
''          End If
'      End Sub

Private Sub Option1_Click()

  If Option1.Value = True Then
     Command2.Caption = "Recepción"
  Else
     Command2.Caption = "O.Fabricación"
  End If
  
End Sub

Private Sub Option2_Click()
  If Option2.Value = True Then
     Command1.Caption = "Recepción"
  Else
     Command1.Caption = "O.Fabricación"
  End If

End Sub

Sub CARGA_ENCABEZADO(MSF1 As msFlexGrid, LISTADECAMPOS$)
    '
    Dim NOCAMPO$(), FORMX$()
    ReDim NOCAMPO$(1), FORMX$(1)
    '
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
    For i% = 0 To CANCA%
      If i% = 0 Then
       ANCHO% = 1
       NUCOLU% = i%:   TEPRUEBA = String$(ANCHO%, "A"): MSF1.ColWidth(NUCOLU%) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU%) = 1: MSF1.TextMatrix(0, NUCOLU%) = ""
       GoTo 50
      Else
       ANCHO% = XVALO(Mid$(FORMX$(i%), 2)) + 2
      End If
      LETR$ = UCase$(Left$(FORMX$(i%), 1))
       ALINEA% = 1
       If LETR$ = "I" Or LETR$ = "G" Or LETR$ = "F" Then ALINEA% = 6
       NUCOLU% = i%:   TEPRUEBA = String$(ANCHO%, "A"): MSF1.ColWidth(NUCOLU%) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
50    Next i%
End Sub



Private Sub Text10_Change()
    CARGAMSNSER (Text10)
End Sub

Private Sub Text6_Change()
      Call CARGAMSFBR(Text6)
End Sub
Sub LIMPIAR()

    Text11 = ""
    Text4 = ""
    Text10 = ""
    MSF.Rows = 1
    '
    Text6 = ""
    MSF1.Rows = 1

End Sub
Sub test()

'SELECT EST_PROD.DS_PROD,  EST_PROD.CD_PROD,  EST_PROD.VLR_CUSTO_GRP,

'EST_PROD_TABLOIDE_CPL.CD_EMP,
'EST_PROD_TABLOIDE_CPL.CD_TABLOIDE,
'EST_PROD_TABLOIDE_CPL.VLR_GRP_PROD,
'EST_PROD_TABLOIDE_CPL.ATUALIZADO,
'EST_PROD_TABLOIDE_CPL.FLAG_PRECO_SUBST,
'EST_PROD_TABLOIDE_CPL.CD_USU,
'EST_PROD_TABLOIDE_CPL.DT_ALT_ULT,
'EST_PROD_TABLOIDE_CPL.STS_PROD,ISNULL(EST_PROD_PRECO.VLR_TABELA,0) AS VLR_TABELA,

'EST_PROD_FABRIC.NM_FABRIC

'From EST_PROD_TABLOIDE_CPL

'INNER JOIN EST_PROD ON     EST_PROD_TABLOIDE_CPL.CD_EMP = EST_PROD.CD_EMP
'AND EST_PROD_TABLOIDE_CPL.CD_PROD = EST_PROD.CD_PROD

'INNER JOIN EST_PROD_FABRIC      ON EST_PROD_FABRIC.CD_FABRIC = EST_PROD.CD_FABRIC

' LEFT JOIN EST_PROD_PRECO ON      EST_PROD_PRECO.CD_EMP = EST_PROD.CD_EMP
'AND EST_PROD_PRECO.CD_PROD = EST_PROD.CD_PROD

'WHERE
'EST_PROD_TABLOIDE_CPL.CD_EMP = 1     AND
'EST_PROD_TABLOIDE_CPL.CD_TABLOIDE = 162 AND ISNULL(
'EST_PROD_TABLOIDE_CPL.STS_PROD, 0) = 0
'
'AND ISNULL(
'EST_PROD_PRECO.CD_FILIAL,0) IN (0,1)
'ORDER BY DS_PROD

'EST_PROD ,
'EST_PROD_TABLOIDE_CPL ,
'EST_PROD_FABRIC ,
'EST_PROD_PRECO

End Sub
