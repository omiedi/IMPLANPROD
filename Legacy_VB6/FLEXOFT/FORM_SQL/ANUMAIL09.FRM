VERSION 5.00
Begin VB.Form ANUMAIL09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Marca de No Responder"
   ClientHeight    =   2955
   ClientLeft      =   3285
   ClientTop       =   3150
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Prospect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   6825
      TabIndex        =   18
      Top             =   105
      Width           =   1485
      Begin VB.Label Label26 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   105
         TabIndex        =   19
         Top             =   525
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   4920
      Left            =   8400
      ScaleHeight     =   4860
      ScaleWidth      =   945
      TabIndex        =   12
      Top             =   -105
      Width           =   1000
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
         Height          =   1060
         Left            =   105
         Picture         =   "ANUMAIL09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Guardar"
         Top             =   1050
         Width           =   640
      End
      Begin VB.CommandButton Command1 
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
         Height          =   675
         Left            =   105
         Picture         =   "ANUMAIL09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   285
         Width           =   640
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   90
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   13
         Top             =   2205
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   14
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   105
         TabIndex        =   17
         Top             =   1785
         Width           =   645
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "ANUMAIL09.frx":0454
         Top             =   2490
         Width           =   1515
      End
   End
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
      TabIndex        =   2
      Top             =   1155
      Width           =   8220
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
         TabIndex        =   6
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
            TabIndex        =   8
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
            TabIndex        =   7
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
         TabIndex        =   4
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
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   255
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
         TabIndex        =   3
         Top             =   855
         Width           =   7815
         Begin VB.CommandButton Command11 
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   6.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   560
            Left            =   7280
            Picture         =   "ANUMAIL09.frx":2376
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Guarda Actividad Programada"
            Top             =   110
            Width           =   520
         End
         Begin VB.TextBox Text8 
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
            ToolTipText     =   "Introduzca 'NUNCA MAS' para Cancelar"
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
         TabIndex        =   10
         Top             =   525
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Motivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1020
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   6645
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
         Height          =   525
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   315
         Width           =   6315
      End
   End
End
Attribute VB_Name = "ANUMAIL09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Hide
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   COCLI$ = TRIM$(Label26)
   CODOP$ = ""
   If COCLI$ <> "" Then
      If FISECOM09.VALIPROCON%(Text12, Text8) > 0 Then
         PERSCON$ = Combo1.TEXT
         REFACT$ = TRIM$(Text8)
         FEPROG% = FECHANUM(Text12)
         Call FISECOM09.PROGRACTI(COCLI$, CODOP$, PERSCON$, REFACT$, FEPROG%)
      End If
   End If
   Command11.Enabled = True
End Sub

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
     If UCase$(TRIM$(Text8)) <> "NUNCA MAS" Then
        If TRIM$(Text2) = "" Then
           Call MENSERR(24, "Mandatorio Indicar Motivo")
           GoTo 99
        End If
     End If
     '
     If FECHANUM(Text12) < HOY(HO$) And UCase$(TRIM$(Text8)) <> "NUNCA MAS" Then
        Call MENSERR(24, "Mandatorio Fecha Futura de Proximo Contacto")
        GoTo 99
     End If
     '
     If TRIM$(Text8) = "" Then
        Call MENSERR(24, "Mandatorio Indicar Tema del proximo Contacto")
        GoTo 99
     End If
     '
     ' PONE LA PROXIMA ACTIVIDAD PROYECTADA
     If TRIM$(Text12) = "" Then
        If TRIM$(UCase$(Text8)) = "NUNCA MAS" Then
           GoTo 20                   ' SUPRIME PROXIMA ACTIVIDAD
        End If
     End If
     '
     Call Command11_Click
     '
20   NORIMAI& = XVALO(Label13)
     If NORIMAI& > 0 Then
        CONDI$ = "NUME_MAIL = " & NORIMAI&
        ' marca como respondido
        Call ACTUREGISTRO("BANDENTRA", "CONF_RECEP", CONDI$, AHORA, RAFE&, "NOMESS")
        ' pone el numero de mail de respuesta = (-1)
        Call ACTUREGISTRO("BANDENTRA", "NUMAIRESPU", CONDI$, (-1), RAFE&, "NOMESS")
        ' pone el motivo de anulacion
        Call CREACAMPO("", "BANDENTRA", "MOTI_ANUL", 12, 0)
        Call ACTUREGISTRO("BANDENTRA", "MOTI_ANUL", CONDI$, Text2, RAFE&)
        '
        ' PONE EL USUARIO QUE ANULÓ
        Call CREACAMPO("", "BANDENTRA", "USUA_ANUL", 3, 2)
        Call ACTUREGISTRO("BANDENTRA", "USUA_ANUL", CONDI$, USNBR% Mod 100, RAFE&)
        ' saca de la lista
        If CONACT09.Visible = True Then
           ILIS& = CONACT09.List3.ListIndex
           If XVALO(Left$(CONACT09.List3.TEXT, 5)) = NORIMAI& Then
                CONACT09.List3.RemoveItem (ILIS&)
                If ILIS& >= 0 Then
                  If ILIS& < CONACT09.List3.ListCount Then
                    CONACT09.List3.Selected(ILIS&) = True
                  End If
                End If
              Else
                Call CONACT09.Command1_Click
           End If
        End If
     End If
     '
     Label13 = "ANUL"
     '
99   Screen.MousePointer = 1
     Command5.Enabled = True
     Hide
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

Private Sub Label11_DblClick()
   Text12 = ""
   Text8 = "Nunca Mas"
End Sub

Private Sub Label26_Change()
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
    CONDI$ = "NUME_ACTI = '" & PROACT$ & "'"
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

