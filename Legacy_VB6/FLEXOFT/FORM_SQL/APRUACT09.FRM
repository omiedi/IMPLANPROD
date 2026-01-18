VERSION 5.00
Begin VB.Form APRUACT09 
   Caption         =   "Aprobacion de Actividad"
   ClientHeight    =   5670
   ClientLeft      =   3360
   ClientTop       =   3225
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   9075
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1710
      Left            =   105
      TabIndex        =   11
      Top             =   3885
      Width           =   8850
      Begin VB.CommandButton Command7 
         Caption         =   "Exit"
         Height          =   570
         Left            =   8085
         Picture         =   "APRUACT09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Cierra Ventana de Aprobacion"
         Top             =   315
         Width           =   640
      End
      Begin VB.CommandButton Command5 
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
         Height          =   570
         Left            =   8085
         Picture         =   "APRUACT09.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Guardar"
         Top             =   945
         Width           =   640
      End
      Begin VB.Frame Frame11 
         Caption         =   "Contacto el:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2625
         TabIndex        =   15
         Top             =   195
         Width           =   1575
         Begin VB.CommandButton Command15 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   3.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1305
            TabIndex        =   17
            Top             =   260
            Width           =   150
         End
         Begin VB.TextBox Text12 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   270
            TabIndex        =   16
            Top             =   260
            Width           =   1030
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Con:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   4305
         TabIndex        =   13
         Top             =   210
         Width           =   3705
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   14
            Text            =   "Combo1"
            Top             =   260
            Width           =   3465
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Asuntos a Tratar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   210
         TabIndex        =   12
         Top             =   855
         Width           =   7815
         Begin VB.CommandButton Command11 
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   560
            Left            =   7280
            Picture         =   "APRUACT09.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Guarda Actividad Programada"
            Top             =   110
            Width           =   520
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   135
            TabIndex        =   1
            Top             =   240
            Width           =   6975
         End
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Proxima Actividad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   840
         TabIndex        =   26
         Top             =   525
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   3750
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   8850
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   1260
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   2625
         Width           =   7365
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Contacto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5925
         TabIndex        =   25
         Top             =   705
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6810
         TabIndex        =   24
         Top             =   645
         Width           =   1800
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   1260
         TabIndex        =   22
         Top             =   1470
         Width           =   7350
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1260
         TabIndex        =   21
         Top             =   1095
         Width           =   7335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Comentario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   20
         Top             =   2640
         Width           =   1065
      End
      Begin VB.Label LABEL52 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   10
         Top             =   630
         Width           =   4650
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         Caption         =   "Referencia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   9
         Top             =   1155
         Width           =   1065
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripcion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   8
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Prospecto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3255
         TabIndex        =   7
         Top             =   345
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Nro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   6
         Top             =   345
         Width           =   960
      End
      Begin VB.Label Label41 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1260
         TabIndex        =   5
         Top             =   315
         Width           =   990
      End
      Begin VB.Label Label26 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4350
         TabIndex        =   4
         Top             =   315
         Width           =   1575
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Oport:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6030
         TabIndex        =   3
         Top             =   345
         Width           =   750
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6810
         TabIndex        =   2
         Top             =   315
         Width           =   1800
      End
   End
End
Attribute VB_Name = "APRUACT09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command15_Click()
    '
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text12.TEXT = VDEV$
    End If
    '
End Sub

Private Sub Command5_Click()
     '
     Command5.Enabled = False
     COCLI$ = TRIM$(Label26)
     CODOP$ = TRIM$(Label10)
     If COCLI$ <> "" Then
         PERSCON$ = Combo1.TEXT
         REFACT$ = TRIM$(Text1)
         FEPROG% = FECHANUM(Text12)
         CODI$ = SAFETEXT$(Label26)
         NUMACT$ = TRIM$(Label41)
         COMEN$ = SAFETEXT$(Text2)
         '
         If COMEN$ = "" Then
            Call MENSERR(24, "Falta Ingresar Comentario de Cierre")
            GoTo 99
         Else
            OPX1% = COMALTER%("¿Desea Cerrar la presente Actividad y Registrar una Nueva?\\No, Salir\Si, Realizar")
            If OPX1% = 2 Then
               Call NUEVAACT(COCLI$, CODOP$, PERSCON$, REFACT$, FEPROG%) ' Proxima Actividad
               CONDI$ = " Nume_Acti = '" & NUMACT$ & "'"
               CONDI$ = CONDI$ + " AND Codi_Cli <= '" & CODI$ & " '"
               Call ACTUREGISTRO("ACTIPROY", "STAT_ACTI", CONDI$, 3, REG&, "NOMESS") ' Cierre de Actividad
               Call ACTUREGISTRO("ACTIPROY", "COMENTAR", CONDI$, COMEN$, REG&, "NOMESS")   ' COMENTARIO DE REALIZACION
               Call COMUNI("Proceso Finalizado")
            End If
         End If
     End If
     Unload Me
     '
     Call CONACT09.Command1_Click
99   Screen.MousePointer = 1
     Command5.Enabled = True
     '
End Sub

Private Sub Command7_Click()
     Unload Me
End Sub

Private Sub Form_Load()
     '
     Label41 = CONACT09.Label41
     FEX = DIPOST(DIPOST(HOY(HOS$)))
     Text12 = FECHATEX$(FEX)
     '
End Sub

Private Sub Label26_Change()
    '
    LABEL52 = VALOBADA("BSEGCOM", "RASO_CLI", "CODI_CLI = '" & Label26 & "'")
    '
    Combo1.Clear
    CODICLIX$ = TRIM$(Label26)
    CONDI$ = "CodiCli = '" & CODICLIX$ & "'"
    '
    SQLX$ = "SELECT * FROM CONTACSCOM WHERE " + CONDI$
    Dim Fichanew As ADODB.Recordset
    Set Fichanew = New ADODB.Recordset
    Set Fichanew = FLEXCONN.Execute(FILTSQL$(SQLX$, "NOMESS"))
    '
    With Fichanew
       Do While Not .EOF
       NOCON$ = TRIM$(SAFETEXT$(!NOME_CTAC))
       If NOCON$ <> "" Then
          Combo1.AddItem NOCON$
       End If
       .MoveNext
       Loop
    End With
    Fichanew.Close
    Set Fichanew = Nothing
    '
End Sub

Private Sub Label41_Change()
       '
       CONDI$ = "NUME_ACTI = '" & Label41 & "'"
       Label26 = VALOBADA("ACTIPROY", "Codi_Cli", CONDI$, "NOMESS")     ' CODIGO DE PROSPECTO
       Label10 = VALOBADA("ACTIPROY", "CODOPORT", CONDI$, "NOMESS")     ' CODIGO DE OPORTUNIDAD
         NOMECON$ = VALOBADA("ACTIPROY", "CONTACTO", CONDI$, "NOMESS")  ' PERSONA DE CONTACTO
       Call FISECOM09.CARLICONTAC(Combo1, Label26, NOMECON$)
       Label2 = VALOBADA("ACTIPROY", "REFE_ACTI", CONDI$, "NOMESS")     ' REFERENCIA DE ACTIVIDAD
       Label3 = VALOBADA("ACTIPROY", "DESC_ACTI", CONDI$, "NOMESS")     ' DESCRIPCION LARGA DE ACTIVIDAD
       Label4 = AJUSTI$(CONACT09.Combo1.TEXT, 24)
       '
End Sub

Sub NUEVAACT(COCLI$, COPOR$, PERSCON$, REFACT$, FEPROG%)
    '
    ' VERIFICAR SI NO ESTA YA GRABADA
    CONDI$ = "Codi_Cli = '" & COCLI$ & "' "                              ' CODIGO DE PROSPECTO
    CONDI$ = CONDI$ & " AND CODOPORT = '" & COPOR$ & "' "                ' CODIGO DE OPORTUNIDAD
    CONDI$ = CONDI$ & " AND CONTACTO = '" & PERSCON$ & "' "              ' PERSONA DE CONTACTO
    CONDI$ = CONDI$ & " AND REFE_ACTI = '" & REFACT$ & "' "              ' REFERENCIA DE ACTIVIDAD
    CONDI$ = CONDI$ & " AND FHOR_PROG = '" & FETEXCOR1$(FEPROG%) & "' "  ' FECHA Y HORA PROGRAMADO
    If CantRegistros("ACTIPROY", CONDI$) > 0 Then GoTo 99                ' YA ESTABA REGISTRADA
    '
    PROACT$ = INICRM09.PROXIACTI$
    Call CreaRegistro("ACTIPROY", "NUME_ACTI", PROACT$)
    '
    CONDI$ = "NUME_ACTI = " & NTAR&
    Call ACTUREGISTRO("ACTIPROY", "Codi_Cli", CONDI$, COCLI$, RAFE&, "NOMESS")               ' CODIGO DE PROSPECTO
    Call ACTUREGISTRO("ACTIPROY", "CODOPORT", CONDI$, COPOR$, RAFE&, "NOMESS")               ' CODIGO DE OPORTUNIDAD
    Call ACTUREGISTRO("ACTIPROY", "CONTACTO", CONDI$, PERSCON$, RAFE&, "NOMESS")             ' PERSONA DE CONTACTO
    Call ACTUREGISTRO("ACTIPROY", "REFE_ACTI", CONDI$, REFACT$, RAFE&, "NOMESS")             ' REFERENCIA DE ACTIVIDAD
    Call ACTUREGISTRO("ACTIPROY", "DESC_ACTI", CONDI$, "", RAFE&, "NOMESS")                  ' DESCRIPCION LARGA DE ACTIVIDAD
    Call ACTUREGISTRO("ACTIPROY", "USUA_PROG", CONDI$, USNBR% Mod 100, RAFE&, "NOMESS")      ' USUARIO PROGRAMADO
    Call ACTUREGISTRO("ACTIPROY", "USUA_ASIG", CONDI$, USNBR% Mod 100, RAFE&, "NOMESS")      ' USUARIO ASIGNADO
    Call ACTUREGISTRO("ACTIPROY", "FHOR_PROG", CONDI$, FETEXCOR1$(FEPROG%), RAFE&, "NOMESS") ' FECHA Y HORA PROGRAMADO
    Call ACTUREGISTRO("ACTIPROY", "USUA_REAL", CONDI$, 0, RAFE&, "NOMESS")                   ' USUARIO QUE REALIZO
    Call ACTUREGISTRO("ACTIPROY", "FHOR_REAL", CONDI$, 0, RAFE&, "NOMESS")                   ' FECHA DE REALIZACION
    Call ACTUREGISTRO("ACTIPROY", "COMENTAR", CONDI$, "", RAFE&, "NOMESS")                   ' COMENTARIO DE REALIZACION
    Call ACTUREGISTRO("ACTIPROY", "USUA_CTRL", CONDI$, 0, RAFE&, "NOMESS")                   ' USUARIO DE CONTROL Y APROBACION
    Call ACTUREGISTRO("ACTIPROY", "FHOR_CTRL", CONDI$, 0, RAFE&, "NOMESS")                   ' FECHA DE CONTROL
    Call ACTUREGISTRO("ACTIPROY", "STAT_ACTI", CONDI$, 0, RAFE&, "NOMESS")                   ' STATUS (0 = ACTIVA ; 3 = CERRADA)
    '
99  Screen.MousePointer = SCMA
    '
End Sub

