VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
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
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Bindings        =   "ABMCERI09.frx":0000
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
         Left            =   11450
         Picture         =   "ABMCERI09.frx":0014
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Cargar Items"
         Top             =   240
         Width           =   520
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
         Left            =   10800
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
         Height          =   555
         Left            =   8040
         Picture         =   "ABMCERI09.frx":031E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Actualizar Indice de Selección de O.Fab."
         Top             =   480
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
         Picture         =   "ABMCERI09.frx":0468
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
         Height          =   555
         Left            =   8640
         TabIndex        =   11
         Top             =   480
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
         Bindings        =   "ABMCERI09.frx":0772
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
      OleObjectBlob   =   "ABMCERI09.frx":0786
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "ABMCERI09.frx":09BA
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

Private Sub COMMAND3_Click()

       COMMAND3.Enabled = True
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
       With RS
         For I& = 1 To MSF.Rows - 1
            .AddNew
            !CODINT = CI1%
            !COD_EXTE = CODI$
            !Descripcion = DESCRI$
            !Status = 0
            !IDENLOTE = LOTE$
            !NUMEROSERIE = MSF.TextMatrix(I&, 4)
            !RUT_ARCHIVO = MSF.TextMatrix(I&, 5)
            !TIT_CERTIFICADO = XVALO(MSF.TextMatrix(I&, 6))
            !HOJAS = XVALO(MSF.TextMatrix(I&, 7))
            !MARBORRA = 0
            .Update
          Next I&
        End With
        '
        CONDI$ = CONDI$ + " AND MARBORRA = 1"
        Call BORRAREGISTROS("CERTIFICADOS", CONDI$, REGAF&, "NOMESS")
        '
99      COMMAND3.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    

    Call SELECHO("OFABNOAN/Indice General de Ordenes de Compra (Pendientes-Cumplidas)")
    NORFA$ = TRIM$(VALACT1$("OFABNOAN"))
    If NORFA$ = "" Then GoTo 99
    '
    CI1% = CODIFAB%(NORFA$)
    
'    If ESTRAZABLE%(CI1%) < 2 Then
    If TRIM$(CODFAMIL$(CI1%)) = "" Then
       Call MENSERR(24, "Código: " & CODEXT(CI1%) & " No Tiene Asignado Familia")
    End If
    '
    Text10 = NORFA$
    Text11 = CODEXT(CI1%)
    Text4 = DESCRIT$(CI1%)
    '
    Call CARGAMSNSER(NORFA$)
    
99  Screen.MousePointer = 1
 
    Command2.Enabled = True
End Sub
Sub ACTUIN_OFABNOAN()
    CONDI$ = "IDSUBOR <> '' AND StatuOrf <> 9"
    Call CARGALISEL("OFABNOAN", "ORDEFABR", "IDSUBOR/A12;COD_INTE/C16+MASTER/A48;CanProy/F10.1", CONDI$, "NOMESS")
End Sub

Sub CARGAMSNSER(NORFA$)
    '
    MSF.Rows = 1
    MSF.Redraw = False
    '
    CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
    CANNUSE& = CantRegistros&("TANUMSE", CONDI$)
    '
    
    SQLX$ = "SELECT COD_EXTE, DESCRIP, IDSUBOR, NUMEROSERIE FROM TANUMSE WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE  IDSUBOR = '" & NORFA$ & "'"
    SQLX$ = SQLX$ + " ORDER BY NUMEROSERIE"
    Set RS = READSET(SQLX$)

    I& = 0
    'MODIFICAR LOS CAMPOS SEGUN
    With RS
        Do While Not .EOF
          I& = I& + 1
          MSF.Rows = MSF.Rows + 1
          MSF.TextMatrix(I&, 1) = SAFETEXT$(!COD_EXTE)
          MSF.TextMatrix(I&, 2) = SAFETEXT$(!DESCRIP)
          MSF.TextMatrix(I&, 3) = SAFETEXT$(!IDSUBOR)
          NUSE$ = TRIM$(SAFETEXT$(!NUMEROSERIE))
          MSF.TextMatrix(I&, 4) = NUSE$
          RUT$ = "" 'POR DEFECTO LA RUTA LA ESCRIBE VACIA
          MSF.TextMatrix(I&, 5) = SAFETEXT$(RUT$)
          
          SQLX$ = "SELECT * FROM CERTIFICADOS WITH(NOLOCK)"
          SQLX$ = SQLX$ + " WHERE IDENLOTE = '" & NORFA$ & "' AND NUMEROSERIE = '" & NUSE$ & "'"
          INGRERUTA% = 0
          'CONTROLA SI TIENEN ASIGNADA RUTA EN TABLA CERTIFICADOS
          Set RS1 = READSET(SQLX$)
          With RS1
             Do While Not .EOF
               RUT$ = TRIM$(SAFETEXT$(!RUT_ARCHIVO))
               ORDENX% = XVALO(TRIM$(SAFETEXT$(!TIT_CERTIFICADO)))
               NUSEX$ = TRIM$(SAFETEXT$(!NUMEROSERIE))
               
               If RUT$ <> "" Or ORDENX% > 0 Then
                  'SI TIENE RUTA O TITULO DE CERTIFICADO AL PRIMERO LE AGREGA LA RUTA
                  If INGRERUTA% < 1 Then
                     MSF.TextMatrix(I&, 5) = SAFETEXT$(RUT$)
                     MSF.TextMatrix(I&, 6) = SAFETEXT$(!TIT_CERTIFICADO)
                     MSF.TextMatrix(I&, 7) = SAFETEXT$(!HOJAS)
                     INGRERUTA% = 1
                  Else
                     'SI ENTRA ACA HAY MAS DE UNO Y VA GENERANDO NUVOS REGISTROS.
                     I& = I& + 1
                     MSF.Rows = MSF.Rows + 1
                     MSF.TextMatrix(I&, 1) = SAFETEXT$(!COD_EXTE)
                     MSF.TextMatrix(I&, 2) = SAFETEXT$(!Descripcion)
                     MSF.TextMatrix(I&, 3) = SAFETEXT$(!IDENLOTE)
                     MSF.TextMatrix(I&, 4) = SAFETEXT$(!NUMEROSERIE)
                     MSF.TextMatrix(I&, 5) = SAFETEXT$(RUT$)
                     MSF.TextMatrix(I&, 6) = SAFETEXT$(!TIT_CERTIFICADO)
                     MSF.TextMatrix(I&, 7) = SAFETEXT$(!HOJAS)
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
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
    '
    CANPRONORFA& = XVALO(VALOBADA("ORDEFABR", "CANPROY", "IDSUBOR = '" & NORFA$ & "'", "NOMESS"))
    '
    If CANNUSE& < CANPRONORFA& Then
          I& = I& + 1
          MSF.Rows = MSF.Rows + 1
          CI1% = XVALO(VALOBADA("ORDEFABR", "COD_INTE", CONDI$, "NOMESS"))
          MSF.TextMatrix(I&, 1) = CODEXT$(CI1%)
          MSF.TextMatrix(I&, 2) = DESCRIT0$(CI1%)
          MSF.TextMatrix(I&, 3) = NORFA$
          NUSE$ = ""
          MSF.TextMatrix(I&, 4) = NUSE$
          '
          RUT$ = "" 'POR DEFECTO LA RUTA LA ESCRIBE VACIA
          MSF.TextMatrix(I&, 5) = RUT$
          
          SQLX$ = "SELECT * FROM CERTIFICADOS WITH(NOLOCK)"
          SQLX$ = SQLX$ + " WHERE IDENLOTE = '" & NORFA$ & "' "
          INGRERUTA% = 0
          'CONTROLA SI TIENEN ASIGNADA RUTA EN TABLA CERTIFICADOS
          Set RS1 = READSET(SQLX$)
          With RS1
             Do While Not .EOF
               RUT$ = TRIM$(SAFETEXT$(!RUT_ARCHIVO))
               ORDENX% = XVALO(TRIM$(SAFETEXT$(!TIT_CERTIFICADO)))
               NUSEX$ = TRIM$(SAFETEXT$(!NUMEROSERIE))
               
               If RUT$ <> "" Or ORDENX% > 0 Then
                  'SI TIENE RUTA O TITULO DE CERTIFICADO AL PRIMERO LE AGREGA LA RUTA
                  If INGRERUTA% < 1 Then
                     MSF.TextMatrix(I&, 5) = SAFETEXT$(RUT$)
                     MSF.TextMatrix(I&, 6) = SAFETEXT$(!TIT_CERTIFICADO)
                     MSF.TextMatrix(I&, 7) = SAFETEXT$(!HOJAS)
                     INGRERUTA% = 1
                  Else
                     'SI ENTRA ACA HAY MAS DE UNO Y VA GENERANDO NUVOS REGISTROS.
                     I& = I& + 1
                     MSF.Rows = MSF.Rows + 1
                     MSF.TextMatrix(I&, 1) = SAFETEXT$(!COD_EXTE)
                     MSF.TextMatrix(I&, 2) = SAFETEXT$(!Descripcion)
                     MSF.TextMatrix(I&, 3) = SAFETEXT$(!IDENLOTE)
                     MSF.TextMatrix(I&, 4) = SAFETEXT$(!NUMEROSERIE)
                     MSF.TextMatrix(I&, 5) = SAFETEXT$(RUT$)
                     MSF.TextMatrix(I&, 6) = SAFETEXT$(!TIT_CERTIFICADO)
                     MSF.TextMatrix(I&, 7) = SAFETEXT$(!HOJAS)
                  End If
               End If
               .MoveNext
             Loop
          End With
    End If

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
    
    I& = 0
    'MODIFICAR LOS CAMPOS SEGUN
    With RS
        Do While Not .EOF
          I& = I& + 1
          MSF1.Rows = MSF1.Rows + 1
          MSF1.TextMatrix(I&, 1) = SAFETEXT$(!COD_EXTE)
          MSF1.TextMatrix(I&, 2) = SAFETEXT$(!DESCRIP)
          LOTE$ = TRIM$(SAFETEXT$(!IDENLOTE))
          MSF1.TextMatrix(I&, 3) = LOTE$
          RUT$ = "" 'POR DEFECTO LA RUTA LA ESCRIBE VACIA
          MSF1.TextMatrix(I&, 4) = RUT$
          CI1% = XVALO(!COD_INTE)
          SQLX$ = "SELECT * FROM CERTIFICADOS WITH(NOLOCK)"
          SQLX$ = SQLX$ + " WHERE IDENLOTE = '" & LOTE$ & "' AND CODINT = " & CI1%
          INGRERUTA% = 0
          'CONTROLA SI TIENEN ASIGNADA RUTA EN TABLA CERTIFICADOS
          Set RS1 = READSET(SQLX$)
          With RS1
             Do While Not .EOF
               RUT$ = TRIM$(SAFETEXT$(!RUT_ARCHIVO))
               If RUT$ <> "" Then
                  'SI TIENE RUTA AL PRIMERO LE AGREGA LA RUTA
                  If INGRERUTA% < 1 Then
                     MSF1.TextMatrix(I&, 4) = SAFETEXT$(RUT$)
                     MSF1.TextMatrix(I&, 5) = SAFETEXT$(!TIT_CERTIFICADO)
                     MSF1.TextMatrix(I&, 6) = SAFETEXT$(!HOJAS)

                     INGRERUTA% = 1
                  Else
                     'SI ENTRA ACA HAY MAS DE UNO Y VA GENERANDO NUVOS REGISTROS.
                     I& = I& + 1
                     MSF1.Rows = MSF1.Rows + 1
                     MSF1.TextMatrix(I&, 1) = SAFETEXT$(!COD_EXTE)
                     MSF1.TextMatrix(I&, 2) = SAFETEXT$(!Descripcion)
                     MSF1.TextMatrix(I&, 3) = SAFETEXT$(!IDENLOTE)
                     MSF1.TextMatrix(I&, 4) = SAFETEXT$(RUT$)
                     MSF1.TextMatrix(I&, 5) = SAFETEXT$(!TIT_CERTIFICADO)
                     MSF1.TextMatrix(I&, 6) = SAFETEXT$(!HOJAS)

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
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
    Call ORDENAR_MSF(ABMCERTI09.MSF, 3, 1)
    MSF1.Redraw = True

99  Command2.Enabled = True


End Sub

Private Sub Command4_Click()
   Text1 = ""
End Sub

Sub Command43_Click()

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
         For I& = 1 To MSF1.Rows - 1
            .AddNew
            CI1% = CODINT%(MSF1.TextMatrix(I&, 1))
            !CODINT = CI1%
            !COD_EXTE = MSF1.TextMatrix(I&, 1)
            !Descripcion = MSF1.TextMatrix(I&, 2)
            !Status = 0
            !IDENLOTE = MSF1.TextMatrix(I&, 3)
            !NUMEROSERIE = ""
            !RUT_ARCHIVO = MSF1.TextMatrix(I&, 4)
            !TIT_CERTIFICADO = XVALO(MSF1.TextMatrix(I&, 5))
            !FAMILIA = TRIM$(CODFAMIL$(CI1%))
            !HOJAS = XVALO(MSF1.TextMatrix(I&, 6))
            !MARBORRA = 0
            .Update
          Next I&
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
        
        Call CARGA_ENCABEZADO(MSF, "Código/A16;Descripción/A40;O.Fabricación./A14;N.Serie/A16;Link Certificado/A48;C.T./I5;Hojas/A5")
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

Private Sub Frame3_DragDrop(Source As CONTROL, x As Single, y As Single)

End Sub

Private Sub Frame6_DragDrop(Source As CONTROL, x As Single, y As Single)

End Sub

Private Sub MSF_DblClick()
       I& = MSF.MouseRow
       j& = MSF.MouseCol
       If I& = 0 Or I& > MSF.Rows Then Exit Sub
       '
       On Error Resume Next
       MSF.Row = I&
       MSF.Col = j&
       On Error Resume Next

       LOTE$ = TRIM$(MSF.TextMatrix(I&, 3))
       If LOTE$ = "" Then
           Exit Sub
       Else
           ASIGNARUTA.Text14 = "O.Fab."
           ASIGNARUTA.Text9(5) = "N.Serie"
           ASIGNARUTA.Text15 = "Cantidad"
           ASIGNARUTA.Text15.Alignment = 0
           ASIGNARUTA.Text10.Alignment = 0
           ASIGNARUTA.Text11 = MSF.TextMatrix(I&, 1)
           ASIGNARUTA.Text4 = MSF.TextMatrix(I&, 2)
           ASIGNARUTA.Text10 = MSF.TextMatrix(I&, 3)
           ASIGNARUTA.Text15 = MSF.TextMatrix(I&, 4)
           ASIGNARUTA.Text1 = MSF.TextMatrix(I&, 5)
           ASIGNARUTA.Text3 = ECOARCH$("TITUCERT", MKI$(XVALO(MSF.TextMatrix(I&, 6))))
           ASIGNARUTA.Text5 = XVALO(MSF.TextMatrix(I&, 7))
           ASIGNARUTA.Show 1
       End If
End Sub


Private Sub MSF1_DblClick()
      I& = MSF1.MouseRow
      j& = MSF1.MouseCol
      If I& = 0 Or I& > MSF1.Rows Then Exit Sub
      '
      
      On Error Resume Next
      MSF1.Row = I&
      MSF1.Col = j&
      On Error GoTo 0
      LOTE$ = TRIM$(MSF1.TextMatrix(I&, 3))
      If LOTE$ = "" Then
           Exit Sub
      Else
           ASIGNARUTA.Text14 = "Lote"
           ASIGNARUTA.Text9(5).Visible = False
           ASIGNARUTA.Text15.Visible = False
           ASIGNARUTA.Text10.Alignment = 0
           ASIGNARUTA.Text11 = MSF1.TextMatrix(I&, 1)
           ASIGNARUTA.Text4 = MSF1.TextMatrix(I&, 2)
           ASIGNARUTA.Text10 = MSF1.TextMatrix(I&, 3)
           ASIGNARUTA.Text1 = MSF1.TextMatrix(I&, 4)
           ASIGNARUTA.Text3 = ECOARCH$("TITUCERT", MKI$(XVALO(MSF1.TextMatrix(I&, 5))))
           ASIGNARUTA.ORDEX% = XVALO(MSF1.TextMatrix(I&, 5))
           ASIGNARUTA.Text5 = XVALO(MSF1.TextMatrix(I&, 6)) ' cantidad de hojas
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

Sub CARGA_ENCABEZADO(MSF1 As MSFlexGrid, LISTADECAMPOS$)
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
    For I% = 0 To CANCA%
      If I% = 0 Then
       ancho% = 1
       NUCOLU% = I%:   TEPRUEBA = String$(ancho%, "A"): MSF1.ColWidth(NUCOLU%) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU%) = 1: MSF1.TextMatrix(0, NUCOLU%) = ""
       GoTo 50
      Else
       ancho% = XVALO(Mid$(FORMX$(I%), 2)) + 2
      End If
      LETR$ = UCase$(Left$(FORMX$(I%), 1))
       ALINEA% = 1
       If LETR$ = "I" Or LETR$ = "G" Or LETR$ = "F" Then ALINEA% = 6
       NUCOLU% = I%:   TEPRUEBA = String$(ancho%, "A"): MSF1.ColWidth(NUCOLU%) = TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(I%)
50    Next I%
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
