VERSION 5.00
Begin VB.Form ANULACT09 
   Caption         =   "Anulacion de Actividad"
   ClientHeight    =   3930
   ClientLeft      =   3360
   ClientTop       =   3225
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   8955
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
      Width           =   8745
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
         Height          =   935
         Left            =   7960
         Picture         =   "ANULACT09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Confirma la Anulacion"
         Top             =   2625
         Width           =   640
      End
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
         TabIndex        =   15
         Top             =   2625
         Width           =   6630
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   1095
         Width           =   7335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Motivo o Comentario:"
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
         Left            =   180
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   315
         Width           =   1800
      End
   End
End
Attribute VB_Name = "ANULACT09"
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
     NUMACT$ = TRIM$(Label41)
     COMEN$ = SAFETEXT$(Text2)
     '
     If NUMACT$ = "" Then GoTo 99
     '
     If COMEN$ = "" Then
         Call MENSERR(24, "Falta Ingresar Motivo o Comentario")
         GoTo 99
       Else
         OPX1% = COMALTER%("¿Desea Anular la presente Actividad SIN Registrar una Nueva?\\No, Salir\Si, Anular")
         If OPX1% = 2 Then
            CONDI$ = "NUME_ACTI = '" & Label41 & "'"
            Call ACTUREGISTRO("ACTIPROY", "STAT_ACTI", CONDI$, 9, REG&, "NOMESS") ' Anulacion de Actividad
            Call ACTUREGISTRO("ACTIPROY", "COMENTAR", CONDI$, COMEN$, REG&, "NOMESS")   ' COMENTARIO DE REALIZACION
            Call COMUNI("Anulada")
            Call CONACT09.Command1_Click
         End If
     End If
     Unload Me
     '
99   Screen.MousePointer = 1
     Command5.Enabled = True
     '
End Sub

Private Sub Label26_Change()
    LABEL52 = VALOBADA("BSEGCOM", "RASO_CLI", "CODI_CLI = '" & Label26 & "'")
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

