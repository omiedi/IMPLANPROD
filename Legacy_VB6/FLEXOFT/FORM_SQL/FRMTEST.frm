VERSION 5.00
Begin VB.Form FRMTEST 
   Caption         =   "Form1"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3480
      TabIndex        =   45
      Top             =   3480
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "CONSULTA SQL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   5040
      TabIndex        =   8
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton Command1 
         Caption         =   "Pedidos Pendientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   37
         Top             =   3720
         Width           =   4095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "N° Remito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   29
         Top             =   2400
         Width           =   4095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "N° Recibo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Label RECIBO1 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3600
         TabIndex        =   44
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N° Rec."
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
         Index           =   5
         Left            =   240
         TabIndex        =   43
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label RECIBO 
         BorderStyle     =   1  'Fixed Single
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
         Index           =   5
         Left            =   960
         TabIndex        =   42
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Label FINAL1 
         BorderStyle     =   1  'Fixed Single
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
         Index           =   5
         Left            =   960
         TabIndex        =   41
         Top             =   3120
         Width           =   3375
      End
      Begin VB.Label inicio 
         BorderStyle     =   1  'Fixed Single
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
         Index           =   5
         Left            =   960
         TabIndex        =   40
         Top             =   2880
         Width           =   3375
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Final"
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
         Index           =   5
         Left            =   240
         TabIndex        =   39
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inicio"
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
         Index           =   5
         Left            =   240
         TabIndex        =   38
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inicio"
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
         Index           =   3
         Left            =   240
         TabIndex        =   28
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Final"
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
         Index           =   3
         Left            =   240
         TabIndex        =   27
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label inicio 
         BorderStyle     =   1  'Fixed Single
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
         Index           =   3
         Left            =   960
         TabIndex        =   26
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label FINAL1 
         BorderStyle     =   1  'Fixed Single
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
         Index           =   3
         Left            =   960
         TabIndex        =   25
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label RECIBO 
         BorderStyle     =   1  'Fixed Single
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
         Index           =   3
         Left            =   960
         TabIndex        =   24
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N° Rec."
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
         Index           =   3
         Left            =   240
         TabIndex        =   23
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N° Rec."
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
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
      Begin VB.Label RECIBO 
         BorderStyle     =   1  'Fixed Single
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
         Index           =   1
         Left            =   960
         TabIndex        =   14
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label FINAL1 
         BorderStyle     =   1  'Fixed Single
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
         Index           =   1
         Left            =   960
         TabIndex        =   13
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label inicio 
         BorderStyle     =   1  'Fixed Single
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
         Index           =   1
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Final"
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
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inicio"
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
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "STORED PROCEDURE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton Command1 
         Caption         =   "Pedidos Pendientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   30
         Top             =   3720
         Width           =   4095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "N° Remito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   22
         Top             =   2400
         Width           =   4095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "N° Recibo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N° Rec."
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
         Index           =   4
         Left            =   240
         TabIndex        =   36
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label RECIBO 
         BorderStyle     =   1  'Fixed Single
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
         Index           =   4
         Left            =   960
         TabIndex        =   35
         Top             =   3360
         Width           =   2415
      End
      Begin VB.Label FINAL1 
         BorderStyle     =   1  'Fixed Single
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
         Index           =   4
         Left            =   960
         TabIndex        =   34
         Top             =   3120
         Width           =   3375
      End
      Begin VB.Label inicio 
         BorderStyle     =   1  'Fixed Single
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
         Index           =   4
         Left            =   960
         TabIndex        =   33
         Top             =   2880
         Width           =   3375
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Final"
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
         Index           =   4
         Left            =   240
         TabIndex        =   32
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inicio"
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
         Index           =   4
         Left            =   240
         TabIndex        =   31
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inicio"
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
         Index           =   2
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Final"
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
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label inicio 
         BorderStyle     =   1  'Fixed Single
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
         Index           =   2
         Left            =   960
         TabIndex        =   19
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label FINAL1 
         BorderStyle     =   1  'Fixed Single
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
         Index           =   2
         Left            =   960
         TabIndex        =   18
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label RECIBO 
         BorderStyle     =   1  'Fixed Single
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
         Index           =   2
         Left            =   960
         TabIndex        =   17
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N° Rec."
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
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inicio"
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
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Final"
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
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin VB.Label inicio 
         BorderStyle     =   1  'Fixed Single
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
         Index           =   0
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label FINAL1 
         BorderStyle     =   1  'Fixed Single
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
         Index           =   0
         Left            =   960
         TabIndex        =   4
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label RECIBO 
         BorderStyle     =   1  'Fixed Single
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
         Index           =   0
         Left            =   960
         TabIndex        =   3
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N° Rec."
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
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
   End
End
Attribute VB_Name = "FRMTEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
    Static CMO$, FLIM%, YARECU%

    FENUE$ = Control$("", "FENURECI")
      If FENUE$ = "" Then FENUE$ = "01/01/80"
      FEXI% = FECHANUM(FENUE$)
      'FLIMI = CVDAT(FECHANUM(FENUE$))
      FLIMI = FETEXCOR1$(FEXI%)

Select Case Index

   Case 0
    'INTERVENCION SP
    inicio(Index) = Now
    inicio(Index).Refresh
    SQLX$ = SQLX$ + " EXECUTE PROXIMO_N_RECIBO " & COBRACT%
    SQLX$ = SQLX$ + ",'" & FLIMI & "'"
    Set RS = FLEXCONN.Execute(SQLX$)
    NROPRO& = 1 + XVALO(RS.Fields(0))
    RECIBO(Index) = NROPRO&
    FINAL1(Index) = Now
    FINAL1(Index).Refresh
    'FIN INTERVENCION SP
Case 1
    'CODIGO ANTERIOR
    inicio(Index) = Now
    inicio(Index).Refresh

    SQLX$ = "SELECT NumeComp as maximo FROM CajaBanc WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + " AND CodiCom_Cor LIKE 'RB%' "
    SQLX$ = SQLX$ + " AND FechEmisi >= '" & FLIMI & "'"
    SQLX$ = SQLX$ + " AND NumeTalo = " & COBRACT%
    SQLX$ = SQLX$ + " AND Status < 9 "
    SQLX$ = SQLX$ + " ORDER BY NUMECOMP DESC"
    Dim MIRS As ADODB.Recordset
    'MODIFCADO PARA CAPTURAR ERROR DE TIEMPO DE ESPERA DENTRO DE FUNCION READSET
    'Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
    Set MIRS = READSET(SQLX$)
    If Not IsNull(MIRS!maximo) Then
        NROPRO& = 1 + XVALO(MIRS!maximo)
    End If
    MIRS.Close
    Set MIRS = Nothing
    RECIBO(Index) = NROPRO&
    FINAL1(Index) = Now
    FINAL1(Index).Refresh
Case 2

    Screen.MousePointer = 11
    '
    If CMO$ = "" Then
      CMO$ = TRIM$(Control$("", "CODREFAC"))
      If CMO$ = "" Then CMO$ = "RT"
    End If
    '
    If FLIM% < 1 Then
      FENUE$ = Control$("", "FENUREFA")
      If FENUE$ = "" Then FENUE$ = "18/04/03"
      FLIM% = FECHANUM(FENUE$)
      If FLIM% > HOY(HO$) Then FLIM% = HOY(HO$)
    End If
    FLID = FETEXCOR1$(FLIM%)
    '
    inicio(Index) = Now
    inicio(Index).Refresh
   
    SQLX$ = " EXECUTE MAXIMO_N_REMITO  '" & CMO$ & "'"
    SQLX$ = SQLX$ + ",  '" & FLID & "'"
    SQLX$ = SQLX$ + ",4 "
    SQLX$ = SQLX$ + "," & CodiEmp%
    Set RS = FLEXCONN.Execute(SQLX$)
    NROPRO& = XVALO(RS.Fields(0))
    RECIBO(Index) = NROPRO&
    FINAL1(Index) = Now
    FINAL1(Index).Refresh
    Screen.MousePointer = 1
Case 3
    Screen.MousePointer = 11
    '
    If CMO$ = "" Then
      CMO$ = TRIM$(Control$("", "CODREFAC"))
      If CMO$ = "" Then CMO$ = "RT"
    End If
    '
    If FLIM% < 1 Then
      FENUE$ = Control$("", "FENUREFA")
      If FENUE$ = "" Then FENUE$ = "18/04/03"
      FLIM% = FECHANUM(FENUE$)
      If FLIM% > HOY(HO$) Then FLIM% = HOY(HO$)
    End If
    FLID = FETEXCOR1$(FLIM%)
    '
    inicio(Index) = Now
    inicio(Index).Refresh

20  NUPU = 4 'XVALO(NPVX$)
    CONDI$ = "FechMov >= '" & FLID & "'"
    CONDI$ = CONDI$ + " AND CodiMovi = '" & CMO$ & "' "
    CONDI$ = CONDI$ + " AND NUMPVTA = " & NUPU
    CONDI$ = CONDI$ & " AND Codiempr =" & CodiEmp%
    
    REMITONUM& = XVALO(VALOBADA("MOVISTO", "MAX(DOPRINUM)", CONDI$, ""))
    RECIBO(Index) = REMITONUM&
    FINAL1(Index) = Now
    FINAL1(Index).Refresh
    Screen.MousePointer = 1
Case 4
    Screen.MousePointer = 11
    inicio(Index) = Now
    inicio(Index).Refresh

    SQLX$ = " EXECUTE PEDIDOS_PENDIENTES " & XVALO(Text1)
    SQLX$ = SQLX$ + "," & CodiEmp%
    SQLX$ = SQLX$ + ", 0"
    Set RS = FLEXCONN.Execute(SQLX$)
    With RS
    Do While Not .EOF
'        A = RS!NroClie
'        B = RS!NROPED
'        c = RS!FECHEMIS
'        D = RS!NROOCOM
        CANTI = CANTI + 1
        .MoveNext
    Loop
    End With
    FINAL1(Index) = Now
    FINAL1(Index).Refresh
    Command1(Index).Caption = " " & TRIM$(Str$(CANTI)) & " PEDIDOS"
    Screen.MousePointer = 1
Case 5
    Screen.MousePointer = 11
    inicio(Index) = Now
    inicio(Index).Refresh


   SQLX$ = " SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom "
   SQLX$ = SQLX$ + " FROM PEDENCA "
   SQLX$ = SQLX$ + " INNER JOIN  PEDDETA "
   SQLX$ = SQLX$ + " ON (PEDENCA.NroPed = PEDDETA.NroPed) "
   SQLX$ = SQLX$ + " WHERE PEDENCA.NroClie = " & Text1 & " "
   SQLX$ = SQLX$ + " AND PEDDETA.Status < 8 "   ' para que no este anulada ni cumplida
   SQLX$ = SQLX$ + " AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005) "
   SQLX$ = SQLX$ + " AND PEDENCA.Status < 9 "
   SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
   SQLX$ = SQLX$ + " AND PEDDETA.CODIEMPR = " & CodiEmp% & " "
   If EsPorPedido& > 0 Then
      SQLX$ = SQLX$ + " AND PEDDETA.NroPed = " & EsPorPedido&
'      GoTo 251
   End If

   SQLX$ = SQLX$ + " ORDER BY PEDENCA.NroPed ASC "
   'JANDY SQL
   Dim PediModifTemp As ADODB.Recordset
   Set PediModifTemp = New ADODB.Recordset
   'Set PediModifTemp = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   PediModifTemp.CursorLocation = adUseClient
   PediModifTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    With PediModifTemp
    Do While Not .EOF
        CANTI = CANTI + 1
        .MoveNext
    Loop
    End With
    FINAL1(Index) = Now
    FINAL1(Index).Refresh
    Command1(Index).Caption = " " & TRIM$(Str$(CANTI)) & " PEDIDOS"

    Screen.MousePointer = 1
End Select
End Sub

Private Sub Command2_Click()
    Static CMO$, FLIM%, YARECU%
    Screen.MousePointer = 11
    '
    If CMO$ = "" Then
      CMO$ = TRIM$(Control$("", "CODREFAC"))
      If CMO$ = "" Then CMO$ = "RT"
    End If
    '
    If FLIM% < 1 Then
      FENUE$ = Control$("", "FENUREFA")
      If FENUE$ = "" Then FENUE$ = "18/04/03"
      FLIM% = FECHANUM(FENUE$)
      If FLIM% > HOY(HO$) Then FLIM% = HOY(HO$)
    End If
    FLID = FETEXCOR1$(FLIM%)
    '
    SQLX$ = SQLX$ + " EXECUTE MAXIMO_N_REMITO  '" & FLID & "'"
    SQLX$ = SQLX$ + ",'" & CMO$ & "'"
    SQLX$ = SQLX$ + ", 4"
    SQLX$ = SQLX$ + "," & CodiEmp%
    Set RS = FLEXCONN.Execute(SQLX$)
    NROPRO& = 1 + XVALO(RS.Fields(0))
    RECIBO(2) = NROPRO&
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    Call ABRECONEXION
    Call CREAR_STORED_PROCEDURE
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Call final("")
End Sub

Private Sub Text1_Change()
  RECIBO1 = Text1
End Sub


