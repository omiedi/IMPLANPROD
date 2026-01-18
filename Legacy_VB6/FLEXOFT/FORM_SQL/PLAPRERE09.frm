VERSION 5.00
Begin VB.Form PLAPRERE09 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilla de Presupuestación por Reparaciones"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11565
   Icon            =   "PLAPRERE09.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   11565
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "O.K"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   10730
      Picture         =   "PLAPRERE09.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   765
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "TRABAJOS A EFECTUAR"
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
      Height          =   1740
      Left            =   80
      TabIndex        =   20
      Top             =   5040
      Width           =   10540
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Left            =   9240
         TabIndex        =   1
         Top             =   390
         Width           =   1095
      End
      Begin VB.TextBox Text3 
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
         Height          =   1185
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   390
         Width           =   7470
      End
      Begin VB.Label Label2 
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
         Height          =   300
         Left            =   8040
         TabIndex        =   25
         Top             =   1260
         Width           =   2310
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Efectuó:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7560
         TabIndex        =   22
         Top             =   945
         Width           =   960
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Estim.Horas:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7860
         TabIndex        =   21
         Top             =   420
         Width           =   1320
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      Caption         =   "REPUESTOS"
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
      Height          =   2520
      Left            =   80
      TabIndex        =   17
      Top             =   2445
      Width           =   10540
      Begin VB.Frame Frame2 
         Caption         =   "                                                                                                                  "
         Height          =   2120
         Left            =   9000
         TabIndex        =   24
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "                                                                                                                  "
         Height          =   2120
         Left            =   1950
         TabIndex        =   23
         Top             =   315
         Width           =   15
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1830
         ItemData        =   "PLAPRERE09.frx":0614
         Left            =   70
         List            =   "PLAPRERE09.frx":0616
         TabIndex        =   18
         Top             =   600
         Width           =   10400
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Código          Descripción                                                     Cant.   "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   75
         TabIndex        =   19
         ToolTipText     =   "Doble Click Para Editar Repuestos"
         Top             =   360
         Width           =   10395
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00CDDADA&
      Caption         =   "CODIGO - MODELO"
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
      Height          =   2355
      Left            =   80
      TabIndex        =   6
      Top             =   30
      Width           =   10545
      Begin VB.Frame Frame9 
         BackColor       =   &H00CDDADA&
         Caption         =   "Técnico Externo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   1
         Left            =   1005
         TabIndex        =   26
         Top             =   240
         Width           =   5175
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   210
            Width           =   375
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Talón Nro"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   2880
            TabIndex        =   31
            Top             =   225
            Width           =   720
         End
         Begin VB.Label Label11 
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
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   3675
            TabIndex        =   30
            Top             =   210
            Width           =   1335
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Técnico:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   -40
            TabIndex        =   29
            Top             =   225
            Width           =   720
         End
         Begin VB.Label Label13 
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
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   1155
            TabIndex        =   28
            Top             =   210
            Width           =   1575
         End
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Accesorios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4935
         TabIndex        =   16
         Top             =   975
         Width           =   1245
      End
      Begin VB.Label Label28 
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
         Height          =   1005
         Left            =   5655
         TabIndex        =   15
         Top             =   1230
         Width           =   4740
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Serie"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6090
         TabIndex        =   14
         Top             =   380
         Width           =   1245
      End
      Begin VB.Label Label26 
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
         Height          =   300
         Left            =   6930
         TabIndex        =   13
         Top             =   735
         Width           =   3480
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   9255
         TabIndex        =   12
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Talón Nro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7350
         TabIndex        =   11
         Top             =   360
         Width           =   1800
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Defectos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   -90
         TabIndex        =   10
         Top             =   1260
         Width           =   1005
      End
      Begin VB.Label Label22 
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
         Height          =   300
         Left            =   1005
         TabIndex        =   9
         Top             =   855
         Width           =   4005
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Equipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   -90
         TabIndex        =   8
         Top             =   915
         Width           =   1005
      End
      Begin VB.Label Label20 
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
         Height          =   1005
         Left            =   1005
         TabIndex        =   7
         Top             =   1230
         Width           =   4005
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   5715
      Left            =   10680
      ScaleHeight     =   5685
      ScaleWidth      =   915
      TabIndex        =   4
      Top             =   0
      Width           =   945
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "PLAPRERE09.frx":0618
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   80
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "PLAPRERE09.frx":0762
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   800
         Width           =   650
      End
   End
End
Attribute VB_Name = "PLAPRERE09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Command1.Enabled = False
    Screen.MousePointer = 11
    '
    NUREP& = XVALO(Label25)
    'GRABA ENCABEZADO
    'VALIDA SI VIENE DE UN STATUS MENOR A 1 LO GRABA SI NO LO MANTIENE
    STAT% = XVALO(VALOBADA("REPARA", "STATUS", "NUMREPA = " & NUREP&, "NOMESS"))
    CONDI$ = "NUMREPA = " & NUREP&
    ESTGAR_REP% = XVALO(VALOBADA("REPARA", "Status_Gar", CONDI$, "NOMESS"))

'    If STAT% < 1 Then Call ACTUREGISTRO("REPARA", "STATUS", "NUMREPA = " & NUREP& & "", "1", REGAF&, "NOMESS")
'    '
'    Call ACTUREGISTRO("REPARA", "Usuario_Repu", "NUMREPA = " & NUREP& & "", Label2, REGAF&, "NOMESS")
'    Call ACTUREGISTRO("REPARA", "Estima_Horas", "NUMREPA = " & NUREP& & "", TRIM$(Text2), REGAF&, "NOMESS")
'     Call ACTUREGISTRO("REPARA", "Trab_Efect", "NUMREPA = " & NUREP& & "", Text3, REGAF&, "NOMESS")
    SQLX$ = "SELECT * FROM REPARA "
    SQLX$ = SQLX$ + " WHERE  NUMREPA = " & NUREP&
    Dim RS2 As ADODB.Recordset
    Set RS2 = New ADODB.Recordset
27  On Error Resume Next
    RS2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 27
    End If
    On Error GoTo 0

    With RS2
      Do While Not .EOF
         'If STAT% < 1 Then !Status = 1
         ESTADO% = XVALO(RS2!Status)
         If ESTADO% < 1 Then RS2!Status = 1
         
         !Usuario_Repu = Label2
         !Estima_Horas = XVALO(Text2)
         !TRAB_EFECT = Text3
         If TRIM$(SAFETEXT$(!REPARACION)) = "" Then !REPARACION = Text3
         
         .MoveNext
      Loop
    End With
    RS2.Close
    Set RS2 = Nothing
    '
    Call ACTUREGISTRO("REPUREPA", "MARBORRA", "NUMREPA = " & NUREP&, 1, REGAF&, "NOMESS")
    '
    'TOMA EL NUMERO DE LISTA SELECCIONADA
    LPRE% = XVALO(Control("REPARA", "LIPREREP"))
    ARCHIPRE$ = TRIM$(Str$(LPRE%))
'
'    If List1.ListCount > 0 Then ' GRABACION DE REPUESTOS DEL LIST
'      For I& = 1 To (List1.ListCount)
'         TX$ = List1.List(I& - 1)
'         CI% = CODINT%(Mid$(TX$, 1, 16))
'         CANTI = XVALO(Mid$(TX$, 75, 8))
'         Call CreaRegistro("REPUREPA", "NUMREPA", NUREP&)
'         Call ACTUREGISTRO("REPUREPA", "NUORD_ITEM", "NUMREPA = " & NUREP& & " AND NUORD_ITEM IS NULL", I&, REGAF&, "NOMESS")
'         '
'         Call ACTUREGISTRO("REPUREPA", "COD_INTE", "NUMREPA = " & NUREP& & " AND NUORD_ITEM = " & I& & "", CI%, REGAF&, "NOMESS")
'         Call ACTUREGISTRO("REPUREPA", "CANT_SOL", "NUMREPA = " & NUREP& & " AND NUORD_ITEM = " & I& & "", FORMATNUM(CANTI, "F8.1"), REGAF&, "NOMESS")
'         COEFI = Val(CONTROL$("REPARA", "COEFPREA"))
'         If COEFI < 0.05 Then COEFI = 1
'         If LPRE% > 0 Then 'SI TIENE CONFIGURADA LISTA TOMA LISTA CASO CONTRARIO TOMA EL COSTO
'            IMPUNI# = PRELISTA#(ARCHIPRE$, CI%)
'         Else
'            IMPUNI# = COSUNI(CI%) * COEFI
'         End If
'         '
'         Call ACTUREGISTRO("REPUREPA", "PRUN_NETO", "NUMREPA = " & NUREP& & " AND NUORD_ITEM = " & I& & "", IMPUNI#, REGAF&, "NOMESS")
'         IMPNET# = XVALO(CANTI) * IMPUNI#
'         Call ACTUREGISTRO("REPUREPA", "IMPO_NETO", "NUMREPA = " & NUREP& & " AND NUORD_ITEM = " & I& & "", IMPNET#, REGAF&, "NOMESS")
'
'      Next I&

    SQLX$ = "SELECT * FROM REPUREPA "
    SQLX$ = SQLX$ + " WHERE  NUMREPA = " & NUREP&
    Dim RS3 As ADODB.Recordset
    Set RS3 = New ADODB.Recordset
28  On Error Resume Next
    RS3.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     '
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 28
     End If
    On Error GoTo 0
    COEFI = Val(Control$("REPARA", "COEFPREA"))
    With RS3
        If List1.ListCount > 0 Then ' GRABACION DE REPUESTOS DEL LIST
          For i& = 1 To (List1.ListCount)
             TX$ = List1.List(i& - 1)
             CI% = CODINT%(Mid$(TX$, 1, 16))
             CANTI = XVALO(Mid$(TX$, 75, 8))
             If ESTGAR_REP% < 1 Then ' SI NO ESTA EN GARANTIA CALCULA  LOS PRECIOS
                If COEFI < 0.05 Then COEFI = 1
                If LPRE% > 0 Then 'SI TIENE CONFIGURADA LISTA TOMA LISTA CASO CONTRARIO TOMA EL COSTO
                   IMPUNI# = PRELISTA#(ARCHIPRE$, CI%)
                Else
                   IMPUNI# = COSUNI(CI%) * COEFI
                End If
                IMPNET# = XVALO(CANTI) * IMPUNI#
             Else
                IMPUNI# = 0: IMPNET# = 0 'SI VIEN AQUI ES POR QUE ESTA EN GARANTIA ENTONCES VALORES 0
             End If
             '
             ' SI SE ESTA INTENTANDO MODIFICAR LA LISTA DE REPUESTOS DE UNA REPARACION QUE YA FUE CERRAR (STATUS = 4) ,
             ' SE DIFERENCIA LOS ITEMS NUEVOS DE LOS PRECARGADO A PARTIR DEL CAMPO MARCAITEMAGREGADO, EL CUAL ES UTILIZADO
             ' EN PRESUPUESTOS PARA DETECTAR LA MODIFICACION DE LOS REPUESTOS Y EN EL CIERRE PARA CONSUMIR UNICAMENTE
             ' LOS NUEVOS REPUETOS.
             MarcaItemAgregado% = 0
             If STAT% = 4 Then
                CondicionSql$ = "NUMREPA = " & NUREP& & " AND COD_INTE=" & CStr(CI%) & " AND MARBORRA = 1 AND MarcaItemAgregado = 0"
                ArticuloPreexistente% = CantRegistros("REPUREPA", CondicionSql$)
                If ArticuloPreexistente% = 0 Then MarcaItemAgregado% = 1
             End If
             '
             .AddNew
             !NumRepa = NUREP&
             !NuOrd_Item = i&
             !COD_INTE = CI%
             !Cant_Sol = XVALO(CANTI)
             !prun_neto = XVALO(IMPUNI#)
             !IMPO_NETO = XVALO(IMPNET#)
             !MarcaItemAgregado = MarcaItemAgregado%
             .Update
          Next i&
        End If
    End With
    RS3.Close
    Set RS3 = Nothing
    '
    CONDI$ = "NUMREPA = " & NUREP& & " and MARBORRA = 1"
    Call BORRAREGISTROS("REPUREPA", CONDI$, REGAF&, "NOMESS")
    '
    NUMEREPA$ = "OR-" & TRIM$(Str$(NUREP&))
    If STAT% > 2 Then
        Call BORRAREGISTROS("RESERVA", "IDSUBOR = '" & NUMEREPA$ & "'", REGAF&, "NOMESS")
    
        SQLX$ = "SELECT * FROM REPUREPA WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE NUMREPA = " & NUREP&
        SQLX$ = SQLX$ + " ORDER BY NUORD_ITEM"
        '
        Dim REPUTMP As ADODB.Recordset
        Set REPUTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With REPUTMP
        Do While Not .EOF
             SQLX1$ = "SELECT * FROM RESERVA"
             SQLX1$ = SQLX1$ + " WHERE COD_INTE < 1 "
             Dim RESETMP As ADODB.Recordset
             Set RESETMP = New ADODB.Recordset
             RESETMP.Open FILTSQL$(SQLX1$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
             With RESETMP
               .AddNew   ' YA CONVERTIDO
               On Error Resume Next
               !CodiEmpr = CodiEmp%
               On Error GoTo 0
               CI% = XVALO(REPUTMP!COD_INTE)
               CANTI = XVALO(REPUTMP!Cant_Sol)
               !COD_INTE = CI%
               !COD_EXTE = CODEXT$(CI%)
               !Descrip = Left$(DESCRIT$(CI%), 64) 'modificado por que devolvia un error de datos
               !IDSUBOR = NUMEREPA$
               !FechRes = Now
               !CANTRES = CANTI
               !CANTCONS = 0
               !Saldo_Entre = CANTI
               !IDENLOTE = ""
               !Deposito = 0
               !USOUNIT = CANTI
               !NuOrItem = XVALO(REPUTMP!NuOrd_Item)
               .Update
             End With
             RESETMP.Close
             Set RESETMP = Nothing
    
        .MoveNext
        Loop
        End With
        
        REPUTMP.Close
        Set REPUTMP = Nothing
    End If
    '
    NUMESER$ = TRIM$(Label26)
    RACLI$ = VALOBADA("REPARA", "RASO_CLI", "NUMREPA = " & NUREP& & "", "NOMESS")
    CUERPO$ = "Planilla de Repuestos de Reparación" & vbCrLf
    CUERPO$ = CUERPO$ + "Número de Reparación: " & NUREP& & vbCrLf
    CUERPO$ = CUERPO$ + "Cliente: " & RACLI$
    Call ANOTANUSE(NUMESER$, CUERPO$, , "NOMESS")

    Screen.MousePointer = 1
    Command1.Enabled = True
    
    ' ** GENERACION AUTOMATICA DEL PRESUPUESTO POR IMPORTES MINIMOS
    
    If ESTGAR_REP% = 0 Then ' ** SIGUE SOLO SI LA REPARACION 'NO ESTA' EN GARANTIA
    
        If STAT% < 2 Then   ' ** SIGUE SOLO SI LA REPARACION NO TIENE PRESUPUESTO
        
            ImporteMinimoGenerarPresupuesto# = XVALO(Control$("REPARA", "GENAUMIN"))
            If ImporteMinimoGenerarPresupuesto# > 0.005 Then ' ** SIGUE SI ESTA CONFIGURADO EL IMPORTE MINIMO
                PRESURE09.Label25 = Label25
                ImportePresupuesto# = XVALO(PRESURE09.Label6)
                If ImportePresupuesto# < ImporteMinimoGenerarPresupuesto# Then
                    PRESURE09.ModoAutomaticoPorImporteMinimo% = 1
                    Call PRESURE09.Command1_Click
                    '
                    MREPARA09.ModoAutomaticoPorImporteMinimo% = 1
                    MREPARA09.ModoAutomaticoPorImporteMinimo_Reparacion = XVALO(Label25)
                    Call MREPARA09.Command4_Click
                Else
                    Call COMUNI("ATENCION !!!\ \Orden de Reparación PENDIENTE de PRESUPUESTAR.")
                End If
            End If
        
        End If
        
    Else
        
        MREPARA09.ModoAutomaticoPorImporteMinimo% = 1
        MREPARA09.ModoAutomaticoPorImporteMinimo_Reparacion = XVALO(Label25)
        Call MREPARA09.Command4_Click
        MREPARA09.ModoAutomaticoPorImporteMinimo% = 0
    End If
    
    Call Command2_Click ' SE CIERRA
    
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Unload Me
 End Sub



Private Sub Frame4_DblClick()
    Call Label9_DblClick
End Sub

Private Sub Label25_Change()
    NUREP& = XVALO(Label25)
    CONDI$ = "NUMREPA = " & NUREP& & ""
    Label22 = VALOBADA("REPARA", "DESCRIPCION", CONDI$, "NOMESS")
    Label26 = VALOBADA("REPARA", "NUMEROSERIE", CONDI$, "NOMESS")
    Label20 = VALOBADA("REPARA", "DEFEC_INFOR", CONDI$, "NOMESS")
    Label28 = VALOBADA("REPARA", "ACCESORIOS", CONDI$, "NOMESS")
    ESTEMHS# = XVALO(VALOBADA("REPARA", "Estima_Horas", CONDI$, "NOMESS"))
    Text2 = FORMATNUM$(ESTEMHS#, "F5.2")
    
    Text3 = VALOBADA("REPARA", "TRAB_EFECT", CONDI$, "NOMESS")
    Text5 = XVALO(VALOBADA("REPARA", "TEC_EXT", CONDI$, "NOMESS"))
    Label11.Caption = VALOBADA("REPARA", "TALON_TEC", CONDI$, "NOMESS")

    Label2 = USERNAME$
    '
    List1.Clear
    '
    SQLX$ = "SELECT * FROM REPUREPA  WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NUMREPA = " & NUREP&
    SQLX$ = SQLX$ + " ORDER BY NUORD_ITEM "
    Dim REPURTMP As ADODB.Recordset
    Set REPURTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With REPURTMP
      Do While Not .EOF
       CI% = XVALO(!COD_INTE)
       CANTI = XVALO(!Cant_Sol)
       
       Call REPLA(TTXX$, CODEXT$(CI%), 1, 16)
       Call REPLA(TTXX$, DESCRIT$(CI%), 17, 56)
       Call REPLA(TTXX$, FORMATNUM(CANTI, "F8.1"), 76, 8)
       
       List1.AddItem TTXX$
       .MoveNext
      Loop
    End With
    REPURTMP.Close
    Set REPURTMP = Nothing
    
    CONDI$ = "NUMREPA = " & NUREP& & ""
    EstadoDeReparacion% = VALOBADA("REPARA", "STATUS", CONDI$, "NOMESS")
    OrdenCerradaConItemsAgregados& = CantRegistros("REPUREPA", CONDI$ & " AND MARCAITEMAGREGADO = 1")
    If EstadoDeReparacion% = 4 Or OrdenCerradaConItemsAgregados& > 0 Then
        Mensaje$ = "La Orden Elegida Se Encuentra Actualmente Cerrada!!!\ \Tener en Cuenta Que Sólo Pueden Agregarse Nuevos Items.\ \La Modificación y/o Eliminación de un Código Ingresado con Anterioridad al Cierre\No Está Contemplado Como Funcionalidad en la Presente Pantalla."
        Call COMUNI(Mensaje$)
    End If
    
End Sub

Private Sub Label9_DblClick()
   Call SETULTREG("!MATEROF", 0)
   
   If List1.ListCount > 0 Then
      For i& = 1 To (List1.ListCount)
         TX$ = List1.List(i& - 1)
         CI% = CODINT%(Mid$(TX$, 1, 16))
         CANTI = XVALO(Mid$(TX$, 76, 8))
         Call REPLA(ZZ$, MKI$(CI%), 1, 2)
         Call REPLA(ZZ$, MKS$(CANTI), 5, 4)
         Call GRAREG("!MATEROF", ZZ$, i&)
      Next i&
   End If
   ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=Repuestos para Reparación"
      
10 VTB% = VENTABU%("EDITREP1" + ATRI$)
   If VTB% < 0 Then
     Exit Sub
   End If
   '
   For U& = 1 To ULTREG&("!MATEROF")
     ZZ$ = REGLEIDO$("!MATEROF", U&)
     CII% = CVI(Left$(ZZ$, 2))
     CANTI = CVS(Mid$(ZZ$, 5, 4))
     If XVALO(CANTI) < 0.005 Then
        Call COMUNI("Código: " & CODEXT$(CII%) & " Sin Cantidad")
        GoTo 10
     End If
   Next U&

   List1.Clear
   For U& = 1 To ULTREG&("!MATEROF")
     ZZ$ = REGLEIDO$("!MATEROF", U&)
     TTXX$ = Space$(128)
     CII% = CVI(Left$(ZZ$, 2))
     Call REPLA(TTXX$, CODEXT$(CII%), 1, 16)
     Call REPLA(TTXX$, DESCRIT$(CII%), 17, 56)
     CANTI = CVS(Mid$(ZZ$, 5, 4))
     Call REPLA(TTXX$, FORMATNUM(CANTI, "F8.1"), 76, 8)
     List1.AddItem TTXX$
   Next U&
   '
   Call CIERRARCH("!MATEROF")
   Screen.MousePointer = 1

End Sub


Private Sub List1_DblClick()
  If List1.ListCount > 0 Then
    Call Label9_DblClick
  End If
End Sub

Private Sub text2_LostFocus()
   Text2 = FORMATNUM$(Text2, "F5.2")
End Sub

Private Sub Text5_Change()
   If XVALO(Text5) > 0 Then
     Frame9(1).Visible = True
     Label13 = NOMBRE_TEC_EXT$(XVALO(Text5))
     
   Else
     Frame9(1).Visible = False
   End If

End Sub
