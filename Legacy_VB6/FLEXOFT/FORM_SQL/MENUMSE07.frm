VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MENUMSE07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trazabilidad por Número de Serie"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8370
   ClipControls    =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   8370
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CMD 
      Left            =   3960
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   375
      Left            =   6480
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      Caption         =   "Impresion de Etiquetas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   225
      TabIndex        =   23
      Top             =   1420
      Width           =   3240
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   320
         ScaleHeight     =   585
         ScaleWidth      =   2505
         TabIndex        =   24
         Top             =   375
         Width           =   2565
         Begin VB.CommandButton Command10 
            Height          =   650
            Left            =   1960
            Picture         =   "MENUMSE07.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Impresión de Etiquetas"
            Top             =   0
            Visible         =   0   'False
            Width           =   560
         End
         Begin VB.CommandButton Command5 
            Height          =   650
            Left            =   0
            Picture         =   "MENUMSE07.frx":066A
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Impresión de Etiquetas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Impresión de Etiquetas"
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
            Left            =   720
            TabIndex        =   25
            Top             =   105
            Width           =   1785
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "Ensayo de Equipos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   225
      TabIndex        =   19
      Top             =   2700
      Width           =   3240
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2145
         TabIndex        =   31
         Top             =   375
         Width           =   2205
         Begin VB.CommandButton Command2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "MENUMSE07.frx":0CD4
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Ensayo de Equipos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Registros de Control"
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
            Left            =   720
            TabIndex        =   32
            Top             =   105
            Width           =   1140
         End
      End
   End
   Begin VB.PictureBox Picture13 
      BackColor       =   &H00404080&
      Height          =   6795
      Left            =   7440
      ScaleHeight     =   6735
      ScaleWidth      =   1110
      TabIndex        =   14
      Top             =   -120
      Width           =   1170
      Begin VB.CommandButton Command9 
         Caption         =   "Diag"
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
         TabIndex        =   40
         Top             =   2520
         Visible         =   0   'False
         Width           =   640
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Recu- pero"
         Height          =   315
         Left            =   120
         TabIndex        =   38
         Top             =   3210
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command20 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   120
         Picture         =   "MENUMSE07.frx":0FDE
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Configuración de Formulario"
         Top             =   1800
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   120
         Picture         =   "MENUMSE07.frx":1420
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Cerrar - Salir de la Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   120
         Picture         =   "MENUMSE07.frx":156A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   1050
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   15
         Top             =   3480
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "MENUMSE07.frx":1874
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   3720
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3810
      Left            =   3675
      TabIndex        =   1
      Top             =   105
      Width           =   3555
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1260
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   35
         Top             =   2285
         Width           =   2010
         Begin VB.CommandButton Command8 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   -15
            Picture         =   "MENUMSE07.frx":3796
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Equipos Aprobados Con O.Fab. Abierta"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Equipos Aprob. C/O.F. Abierta"
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
            Left            =   600
            TabIndex        =   36
            Top             =   90
            Width           =   1305
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   725
         ScaleHeight     =   585
         ScaleWidth      =   2040
         TabIndex        =   28
         Top             =   3080
         Width           =   2100
         Begin VB.CommandButton Command4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "MENUMSE07.frx":3AA0
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Historial de Equipos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Historial de Equipos"
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
            Left            =   735
            TabIndex        =   30
            Top             =   105
            Width           =   1500
         End
      End
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   2
         Top             =   1785
         Width           =   2010
         Begin VB.CommandButton Command18 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "MENUMSE07.frx":3DAA
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Listado de Números de Serie por Cliente"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Cliente"
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
            Left            =   720
            TabIndex        =   3
            Top             =   200
            Width           =   1185
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   4
         Top             =   1275
         Width           =   2010
         Begin VB.CommandButton Command23 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "MENUMSE07.frx":41EC
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Listado de Números de Serie por Remito"
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Número     de Remito"
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
            Left            =   735
            TabIndex        =   5
            Top             =   50
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   6
         Top             =   780
         Width           =   2010
         Begin VB.CommandButton Command15 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "MENUMSE07.frx":44F6
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Listado de Números de Serie por OF"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Orden de Fabricación"
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
            Left            =   720
            TabIndex        =   8
            Top             =   50
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   9
         Top             =   285
         Width           =   2010
         Begin VB.CommandButton Command26 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "MENUMSE07.frx":4800
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Listado por Número de Serie"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Por Número              de Serie"
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
            Left            =   735
            TabIndex        =   11
            Top             =   50
            Width           =   1905
         End
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         X1              =   0
         X2              =   3570
         Y1              =   3000
         Y2              =   3000
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "Transacciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   225
      TabIndex        =   0
      Top             =   120
      Width           =   3270
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2190
         TabIndex        =   12
         Top             =   375
         Width           =   2250
         Begin VB.CommandButton Command17 
            Caption         =   "Tr"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   17.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   1630
            Picture         =   "MENUMSE07.frx":494A
            TabIndex        =   34
            ToolTipText     =   "Trazabilidad de Componentes"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "MENUMSE07.frx":4C54
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Generador de Número de Serie"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Números de Serie"
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
            Left            =   720
            TabIndex        =   13
            Top             =   105
            Width           =   975
         End
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11025
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "MENUMSE07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command10_Click()
    '
    Command10.Enabled = False
    
    Call ABREORFAB
5   OPX% = ALTERNA%("Selección x O.Fabricación\Selección x N.Serie")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    If OPX% = 1 Then
        Call GENUMSE07.ACTLISNUMSERXOF
        '
10      Call SELECHO("INNUSEOF")
        NORFA$ = TRIM$(VALACT1$("INNUSEOF"))
        If NORFA$ = "" Then GoTo 5
        '
        SQLX$ = "SELECT NUMEROSERIE,COD_INTE FROM TANUMSE WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "'"
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
             NUSE$ = SAFETEXT$(!NUMEROSERIE)
             CI1% = XVALO(!COD_INTE)
             Call PRINTETIQ(CI1%, NUSE$, , 1)
             .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
        GoTo 10
    Else
20      Call SELECHO("INDINUSE")
        TTXX$ = TRIM$(VALACT1$("INDINUSE"))
        If TTXX$ = "" Then GoTo 5
        
        RRXX$ = TRIM$(VALACT2$("INDINUSE"))
        POS% = InStr(RRXX$, " ")
        Codigo% = CODINT(TRIM$(Mid$(RRXX$, 1, POS% - 1)))
        NUSERIE$ = TTXX$
        '
        Call PRINTETIQ(Codigo%, NUSERIE$, , 1)
        GoTo 20
            '
    End If
    '
99  Command10.Enabled = True

End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    ' numeros de serie faltantes
    MESAN$ = "1009"
    Dim ACTI%(32767)
    MAXI% = 0: MINI% = 32767
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    'Set RS = FLEXCONN.Execute("select NUMEROSERIE FROM TANUMSE")
    SQLX$ = "select NUMEROSERIE FROM TANUMSE WITH(NOLOCK)"
    Set Rst = READSET(SQLX$)
    With rs
      Do While Not .EOF
        NS$ = SAFETEXT$(!NUMEROSERIE)
        If Mid$(NS$, 5, 4) = MESAN$ Then
           INUM% = XVALO(Mid$(NS$, 9))
           If INUM% < MINI% Then MINI% = INUM%
           If INUM% > MAXI% Then MAXI% = INUM%
           ACTI%(INUM%) = ACTI%(INUM%) + 1
        End If
      .MoveNext
      Loop
    End With
    '
    For i% = 1 To MAXI%
       If ACTI%(i%) <> 1 Then
          AAA = BBB
       End If
    Next i%
    
    Command6.Enabled = True
End Sub

Private Sub Command15_Click()
    '\\\OT-08-0929-OD-06/10/08///
    Command15.Enabled = False
    '
    Call ACTUREGISTRO("TANUMSE", "DEPONUMSE", "DEPONUMSE IS NULL", "-1", REG&) ' ESTO DEJA A LOS NUMEROS DE SERIE EN -1 CUANDO TIENEN NULL
    '
    RTA1% = ALTERNA%("Numero de Serie por O/F\Equipos en Proceso de Fabricacion")
    If RTA1% < 1 Or RTA1% > 2 Then GoTo 99
    If RTA1% = 1 Then Call FINAL("?LNUSEOF")
    If RTA1% = 2 Then Call FINAL("?EQUIPROFA")
    '
99  Command15.Enabled = True
    '\\\OT-08-0929-OD-FIN///
End Sub

Private Sub Command17_Click()
    Command17.Enabled = False
    If CONTROL("CONFTRAZ", "ASIGLOTE") = "SI" Then
       LOTEXSQR.Show 1
       GoTo 99
    End If
    
    ACONEC$ = "REGTRA07/REGTRA07"
    Call CONECRUN(ACONEC$, "Registro de Trazabilidad")
99  Command17.Enabled = True
End Sub

Private Sub Command18_Click()
    '\\\OT-08-0929-OD-06/10/08///
    '
    Command18.Enabled = False
    Call FINAL("?LNUSECLI")
    '
   Command18.Enabled = True
    '\\\OT-08-0929-OD-FIN///

End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False

10  TTXX$ = "Registro de Control Final\Informe de Conformidad de Puesta a Tierra\"
    TTXX$ = TTXX$ + "Informe de Corriente de Fuga a Tierra\Registro de Medición\Ensayo De Seguridad Eléctrica"
    OPXX% = ALTERNA%(TTXX$)
    '
    If OPXX% < 1 Or OPXX% > 5 Then GoTo 99
    '
    If OPXX% = 1 Then
        REGICON07.Show 1: GoTo 10
      ElseIf OPXX% = 2 Then
        COPUTIE07.Show 1: GoTo 10
      ElseIf OPXX% = 3 Then
         COFUTIE07.Show 1: GoTo 10
      ElseIf OPXX% = 4 Then
        REGMEDI07.Show 1: GoTo 10
     ElseIf OPXX% = 5 Then
        SEGELE07.Show 1: GoTo 10
    End If
    '
99  Command2.Enabled = True
    If APLINVO$ <> "" Then
       Call FINAL("")
    End If
    '
End Sub
Private Sub Command20_Click()
   If DERACCE("CFGNUSE", "Configuracion de Formulario - N/Serie") > 1 Then
     CONETIEQUI07.Show 1
   End If
End Sub

Private Sub Command23_Click()
    '
    Command23.Enabled = False
    Call FINAL("?LNUSEREM")
    Command23.Enabled = True
    '
End Sub

Private Sub Command26_Click()
    '
    Command26.Enabled = False
    RTA1% = ALTERNA%(" Filtro por Rango de Fecha \Filtro por Número de Serie")
    If RTA1% < 1 Or RTA1% > 2 Then GoTo 99
    If RTA1% = 1 Then Call FINAL("?LISNUMSE")
    If RTA1% = 2 Then Call FINAL("?LISNUMSE1")
99  Command26.Enabled = True
    '
End Sub

Private Sub Command3_Click()
   If DERACCE("GENUMSER", "Generar Números de Serie") > 1 Then
     GENUMSE07.Show 1
   End If
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    CONUSE07.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    '
    Command5.Enabled = False
    
    Call ABREORFAB
5   OPX% = ALTERNA%("Selección x O.Fabricación\Selección x N.Serie")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    If OPX% = 1 Then
        Call GENUMSE07.ACTLISNUMSERXOF
        '
10      Call SELECHO("INNUSEOF")
        NORFA$ = TRIM$(VALACT1$("INNUSEOF"))
        If NORFA$ = "" Then GoTo 5
        '
        SQLX$ = "SELECT NUMEROSERIE,COD_INTE FROM TANUMSE WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "'"
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
             NUSE$ = SAFETEXT$(!NUMEROSERIE)
             CI1% = XVALO(!COD_INTE)
             Call PRINTETIQ(CI1%, NUSE$)
             .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
        GoTo 10
    Else
20      Call SELECHO("INDINUSE")
        TTXX$ = TRIM$(VALACT1$("INDINUSE"))
        If TTXX$ = "" Then GoTo 5
        
        RRXX$ = TRIM$(VALACT2$("INDINUSE"))
        POS% = InStr(RRXX$, " ")
        Codigo% = CODINT(TRIM$(Mid$(RRXX$, 1, POS% - 1)))
        NUSERIE$ = TTXX$
        '
        Call PRINTETIQ(Codigo%, NUSERIE$, 1)
        GoTo 20
            '
    End If
    '
99  Command5.Enabled = True
End Sub

Private Sub Command7_Click()
    '
    If DERACCE%("RECUHIS", "Recuperación de Historial de Números de Serie") > 1 Then
    
        Command7.Enabled = False
        Screen.MousePointer = 11
        '
        ' LIMPIA TODO EL HISTORIAL
        Call ACTUREGISTRO("TANUMSE", "HISTORIAL", "NUMEROSERIE <> ''", "", RAFE&)
        '
        FIN& = CantRegistros("TANUMSE")
        JJ& = 0
        SQLX$ = "SELECT * FROM TANUMSE WITH(NOLOCK)"
        Dim TANUTMP As ADODB.Recordset
        'Set TANUTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        Set TANUTMP = READSET(SQLX$)
        With TANUTMP
          Do While Not .EOF
            JJ& = JJ& + 1
            Call TRACE(20, JJ&, FIN&)
            NSERIE$ = SAFETEXT(!NUMEROSERIE)
            ORFA$ = SAFETEXT(!IDSUBOR)
            FF% = CVINT(!FECHALTA)
            FFENTREGA% = CVINT(!FechEntrega)
            NUMEFACTURA$ = SAFETEXT$(!NUMFACTURA)
            NUMEREMITO$ = SAFETEXT$(!NumRemito)
            ' CARGA DE ALTA DE NUMERO DE SERIE
            CUERPO$ = "Alta N/Serie a partir de "
            If TRIM$(ORFA$) <> "" Then ' si tiene numero de of
               Call ANOTANUSE(NSERIE$, CUERPO$ & ORFA$, FF%)
            Else         ' caso contrario lo anota como carga manual
               CUERPO$ = "Alta N/Serie en Forma Manual"
               Call ANOTANUSE(NSERIE$, CUERPO$, FF%)
            End If
            '
            ' ENSAYO DE PUESTA A TIERRA
            '
            CONDI$ = "NUMEROSERIE = '" & NSERIE$ & "'" ' CONDICION  QUE SE UTILIZA PARA TODAS
            PUESTATIERRA$ = VALOBADA("ENSAR070", "NUMEROSERIE", CONDI$, "NOMESS") ' VERIFICO SI EXISTE EN LA TABLA
            '
            If TRIM$(PUESTATIERRA$) <> "" Then ' SI EXISTE EL NUMERO DE SERIE
               '
               TCNCO$ = VALOBADA("ENSAR070", "TECNICO", CONDI$, "NOMESS") ' TRAIGO EL TECNICO
               RTDO$ = VALOBADA("ENSAR070", "VerGeneral", CONDI$, "NOMESS") 'TRAIGO EL RESULTADO
               FF% = CVINT(VALOBADA("ENSAR070", "Fech_Ensa", CONDI$, "NOMESS")) ' TRAIGO LA FECHA
               CUERPO1$ = "Conformidad Puesta a Tierra" & vbCrLf ' ARMO EL CUERPO DE LA ANOTACION
               CUERPO1$ = CUERPO1$ & "Técnico: " & TCNCO$
               CUERPO1$ = CUERPO1$ & " - Resultado: " & RTDO$
               '
               Call ANOTANUSE(NSERIE$, CUERPO1$, FF%)
               '
            End If
            '
            'ENSAYO DE CORRIENTE DE FUGA
            
            CORRIENTEFUGA$ = VALOBADA("ENSAR072", "NUMEROSERIE", CONDI$, "NOMESS") ' VERIFICO SI EXISTE EN LA TABLA
            '
            If TRIM$(CORRIENTEFUGA$) <> "" Then ' SI EXISTE EL NUMERO DE SERIE
               '
               TCNCO$ = VALOBADA("ENSAR072", "TECNICO", CONDI$, "NOMESS") ' TRAIGO EL TECNICO
               RTDO$ = VALOBADA("ENSAR072", "VerGeneral", CONDI$, "NOMESS") 'TRAIGO EL RESULTADO
               FF% = CVINT(VALOBADA("ENSAR072", "Fech_Ensa", CONDI$, "NOMESS")) ' TRAIGO LA FECHA
               CUERPO2$ = "Corriente de Fuga a Tierra" & vbCrLf ' ARMO EL CUERPO DE LA ANOTACION
               CUERPO2$ = CUERPO2$ & "Técnico: " & TCNCO$
               CUERPO2$ = CUERPO2$ & " - Resultado: " & RTDO$
               '
               Call ANOTANUSE(NSERIE$, CUERPO2$, FF%)
               '
            End If
            '
            'REGISTRO DE MEDICION
            '
            REGMEDICION$ = VALOBADA("ENSAR074", "NUMEROSERIE", CONDI$, "NOMESS") ' VERIFICO SI EXISTE EN LA TABLA
            '
            If TRIM$(REGMEDICION$) <> "" Then ' SI EXISTE EL NUMERO DE SERIE
               '
               TCNCO$ = VALOBADA("ENSAR074", "TECNICO", CONDI$, "NOMESS") ' TRAIGO EL TECNICO
               RTDO$ = VALOBADA("ENSAR074", "VerGeneral", CONDI$, "NOMESS") 'TRAIGO EL RESULTADO
               FF% = CVINT(VALOBADA("ENSAR074", "Fech_Ensa", CONDI$, "NOMESS")) ' TRAIGO LA FECHA
               CUERPO3$ = "Registro de Medición" & vbCrLf ' ARMO EL CUERPO DE LA ANOTACION
               CUERPO3$ = CUERPO3$ & "Técnico: " & TCNCO$
               CUERPO3$ = CUERPO3$ & " - Resultado: " & RTDO$
               '
               Call ANOTANUSE(NSERIE$, CUERPO3$, FF%)
               '
            End If
            '
            'REGISTRO DE CONTROL FINAL
            '
            REGCTROLFINAL$ = VALOBADA("CONTROLFIN", "NUMEROSERIE", CONDI$, "NOMESS") ' VERIFICO SI EXISTE EN LA TABLA
            '
            If TRIM$(REGCTROLFINAL$) <> "" Then ' SI EXISTE EL NUMERO DE SERIE
               '
               TCNCO$ = VALOBADA("TANUMSE", "TEC_ENSA", CONDI$, "NOMESS") ' TRAIGO EL TECNICO DEL TANUMSE POR QUE EN CONTROL FIN NO TIENE
               RTDO$ = VALOBADA("TANUMSE", "DIC_ENSA", CONDI$, "NOMESS") 'TRAIGO EL RESULTADO
               FF% = CVINT(VALOBADA("CONTROLFIN", "Fech_Ensa", CONDI$, "NOMESS")) ' TRAIGO LA FECHA
               CUERPO4$ = "Registro de Control Final" & vbCrLf ' ARMO EL CUERPO DE LA ANOTACION
               CUERPO4$ = CUERPO4$ & "Técnico: " & TCNCO$
               CUERPO4$ = CUERPO4$ & " - Resultado: " & RTDO$
               '
               Call ANOTANUSE(NSERIE$, CUERPO4$, FF%)
               '
            End If
            '
            'PEDDETA (VERIFICAR SI EN EL PEDETA EXISTE EL NUMERO DE SERIE)
            '
            NUSPEDDETA$ = VALOBADA("PEDDETA", "NUMEROSERIE", CONDI$, "NOMESS") ' VERIFICO SI EXISTE EN LA TABLA
            '
            If TRIM$(NUSPEDDETA$) <> "" Then ' SI EXISTE EL NUMERO DE SERIE
               '
               NUPED& = XVALO(VALOBADA("PEDDETA", "NROPED", CONDI$, "NOMESS")) ' TRAIGO EL NUMERO DE PEDIDO
               NC% = XVALO(VALOBADA("PEDDETA", "NROCLIE", CONDI$, "NOMESS")) 'TRAIGO EL RESULTADO
               razonSocial$ = rasocli$(NC%)
               FF% = CVINT(VALOBADA("PEDDETA", "fechemis", CONDI$, "NOMESS")) ' TRAIGO LA FECHA
               CUERPO5$ = "Asignado a Pedido: " & NUPED& & vbCrLf ' ARMO EL CUERPO DE LA ANOTACION
               CUERPO5$ = CUERPO5$ & "Cliente: " & NC%
               CUERPO5$ = CUERPO5$ & " - Razon Social: " & razonSocial$
               '
               Call ANOTANUSE(NSERIE$, CUERPO5$, FF%)
               '
            End If
            '
            'MOVISTO (MOVIMIENTO EN EL MOVISTO)
            '
            SQLX$ = "SELECT DEPOMOVI,FECHMOV FROM MOVISTO WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE IDENLOTE = '" & NSERIE$ & "'"
            SQLX$ = SQLX$ + " AND CODIMOVI = 'TR'"
            SQLX$ = SQLX$ + " AND CANTINGRE > 0"
            Dim MOVTMP As ADODB.Recordset
            'Set MOVTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
            Set MOVTMP = READSET(SQLX$)
            With MOVTMP
              Do While Not .EOF
                DEP% = XVALO(!DepoMovi)
                FF% = CVINT(!Fechmov)
                CUERPO6$ = "Transferencia de Equipo a Depósito: " & DEP%
                Call ANOTANUSE(NSERIE$, CUERPO6$, FF%)
                .MoveNext
              Loop
            End With
            MOVTMP.Close
            Set MOVTMP = Nothing
            '
            'NUMERO DE FACTURA Y REMITO
            '
            ' LA INFORMACION LA TRAIGO DEL RECORDSET DEL TANUMSE
            If TRIM$(NUMEFACTURA$) <> "" Or TRIM$(NUMEREMITO$) <> "" Then ' LA INFORMACION LA TRAIGO DEL RECORDSET
                CUERPO7$ = "Venta de Equipo S/Factura: " & NUMEFACTURA$ & vbCrLf
                CUERPO7$ = CUERPO7$ & "Remito: " & NUMEREMITO$
                Call ANOTANUSE(NSERIE$, CUERPO7$, FFENTREGA%) ' FECHA TAMBIEN VIENE DEL RECORDSET DEL TANUMSE
            
            End If
            
            .MoveNext
            
          Loop
        End With
        TANUTMP.Close
        Set TANUTMP = Nothing
        Screen.MousePointer = 1
        Command7.Enabled = True
    End If
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call FINAL("?EQUIPENCIE") 'Reporte de Equipos Aprobados con O.F. Abierta
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   ' DIAGNOSTICO DE COMPARACION SI EN TABLA MOVISTO DIFIERE DE LAS ENTREGAS DEL TANUMSE
   
   Command9.Enabled = False
   Screen.MousePointer = 11
   '
   Call DESHASFECHA(Des%, Has%, PERIO$)
   If PERIO$ = "" Then GoTo 99
   '
   Desde = FETEXCOR1$(Des%)
   Hasta = FETEXCOR1$(Has%)

   Call ABRESTOCKS
   '
   SQLX$ = "SELECT SUM(CANTSALID) AS MOV_CANTSAL, COD_INTE, DOPRILAR"
   SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE CODIMOVI = 'RT'"
   SQLX$ = SQLX$ + " AND CANTSALID > 0"
   SQLX$ = SQLX$ + " AND FECHMOV >= '" & Desde & "'"
   SQLX$ = SQLX$ + " AND FECHMOV <= '" & Hasta & "'"
   SQLX$ = SQLX$ + " GROUP BY COD_INTE,DOPRILAR"

   NARCHI2% = FILEOPEN%(LUDAT$ + "MOVISTO_TANUMSE.txt", 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA

   Dim MOVTMP As New ADODB.Recordset
   Set MOVTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With MOVTMP
     Do While Not .EOF
      CI% = XVALO(!COD_INTE)
      If ESTRAZABLE%(CI%) = 2 Then
        MOV_NUREMITO$ = SAFETEXT$(!DoPriLar)
        MOV_CANTSALIDA& = XVALO(!MOV_CANTSAL)
        CONDI$ = "NUMREMITO = '" & MOV_NUREMITO$ & "' AND COD_INTE = " & CI% & ""
        TAN_CANTSALIDA& = CantRegistros&("TANUMSE", CONDI$)
        If TAN_CANTSALIDA& <> MOV_CANTSALIDA& Then
           TX$ = "Diferencia en Remito: " & MOV_NUREMITO$ & " -Código: " & CODEXT$(CI%) & " -Movisto: " & MOV_CANTSALIDA& & " -Tanumse: " & TAN_CANTSALIDA&
           Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
           'Call COMUNI("Diferencia en Remito: " & MOV_NUREMITO$ & "- Código: " & CODEXT$(CI%) & " \Movisto: " & MOV_CANTSALIDA& & "\Tanumse: " & TAN_CANTSALIDA&)
        End If
      End If
      .MoveNext
     Loop
   End With
   Close #NARCHI2%

99 Screen.MousePointer = 1
   Command9.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    Call ABRECONEXION
    Call CREATABLA_EQUIPOS
    Call CREAR_STORED_PROCEDURE
    '
    If APLINVO$ = "ENSAEQ" Then
      'ESTO VIENE DE ORDEN DE FABRICACION CON LA OPCION 'Ensayos y Numeros de Serie'
      If CONTROL("CONFTRAZ", "ASIGLOTE") <> "SI" Then
       Call Command2_Click
      Else
       '==============================================================================================
       'EYM
       'si esta configurado que asigna los lotes via lectora
       LOTEXSQR.Show 1:
       Call CIERRARCH("*.*")
       Call FINAL("")
       '==============================================================================================
      End If
    End If
    '
    If Left$(APLINVO$, 7) = "GENSEOF" Then ' VIENE DE ORDEN DE FABRICACION
        If InStr(Mid$(APLINVO$, 8), "OF-") > 0 Then 'CUANDO SE GENERA O.F. DE TRAZABLE=2
          GENUMSE07.Label13 = TRIM$(Mid$(APLINVO$, 8))
          If TRIM$(CONTROL$(IDENTER$, "IMPAUTOM")) = "SI" Then
              GENUMSE07.Command1_Click
          Else
              GENUMSE07.Show 1
          End If
          'CUANDO VIENE DE OF IMPRIME AUTOMATICAMENTE
          Call FINAL("")
          Exit Sub
        End If
    End If
    '
    If CONTROL("CONFTRAZ", "ASIGLOTE") = "SI" Then
       '==============================================================================================
       'EYM TEMPORAL PARA QUE HABRA DIRECTEMANTE
       'si esta configurado que asigna los lotes via lectora
       LOTEXSQR.Show 1
       Call CIERRARCH("*.*")
       Call FINAL("")

       '==============================================================================================
      End If
    'PARA PODER IMPRIMIR LAS ETIQUETAS DESDE AFUERA
    If Left$(APLINVO$, 8) = "IMPRIETI" Then ' VIENE DE ORDEN DE FABRICACION
       NUSERIE$ = TRIM$(Mid$(APLINVO$, 9))
       CIXI% = XVALO(VALOBADA("TANUMSE", "COD_INTE", "NUMEROSERIE='" & NUSERIE$ & "'"))
       Call PRINTETIQ(CIXI%, NUSERIE$)
       Call FINAL("")
       Exit Sub
    End If
    '
    If HOY(HOS$) > FECHANUM("05/08/09") Then ' BOTON DE RECUPERACION DE HISTORIAL PARA MEDITEA
       Command7.Visible = False
    End If
    
    If InStr(1, EMPREAC$, "MEDITEA") > 0 Then
        Command10.Visible = True
    End If
End Sub

Sub PRINTETIQ(CIXI%, NUSERIE$, Optional REIMPRE%, Optional MODO%)
    '
    Static STAT_FORIPRE$
    Static CONSUMO$
    Static STAT_OF$

    CODIX$ = CODEXT$(CIXI%)
    If REIMPRE% > 0 Then 'SI ES REIMPRESION SIEMPRE PREGUNTA CON QUE FRM IMPRIME
      STAT_FORIPRE$ = "" 'EXCEPTO QUE ESTE CONFIGURADO TAMAÑO DE ETIQUETA
      CONSUMO$ = ""      'PERO SI ES 110 O 220 SIEMPRE PREGUNTA EN LA REIMPRESION
      STAT_OF$ = ""
    End If
    '
    condi0$ = "NUMEROSERIE = '" & NUSERIE$ & "'"
    NORFA$ = VALOBADA("TANUMSE", "IDSUBOR", condi0$, "NOMESS")

    If NORFA$ <> STAT_OF$ Then  ' SI CAMBIA EL  CODIGO  PREGUNTA
       STAT_OF$ = NORFA$        ' EN QUE TAMAÑO IMPRIME LA ETIQUETA
    Else
       NOMENSATENSION% = 1
       GoTo 30                   'SI EL CODIGO ES EL MISMO NO PREGUNTA
    End If
    '
    JJ% = 0 ' MODIFICACION PARA PODER IMPRIMIR ETIQUETAS DOS MEDIDAS
    FORIPR1$ = TRIM$(CONTROL$("", "STIKAPRO")) 'ETIQUETA STANDARD
    FORIPR2$ = TRIM$(CONTROL$("", "STIKAPR2")) 'ETIQUETA PEQUEÑA
    FORIPR_IMP$ = TRIM$(CONTROL("", "STIKIMPO"))
    
    '
    If FORIPR1$ <> "" Then JJ% = JJ% + 1: FORIPRE$ = FORIPR1$ 'POR SI HAY 1 SOLO FRM CONFIGURADO
    If FORIPR2$ <> "" Then JJ% = JJ% + 1: FORIPRE$ = FORIPR2$ 'POR SI HAY 1 SOLO FRM CONFIGURADO
    
    '
    If JJ% < 1 Then
        Call COMUNI("Falta Configurar Formulario de Impresión de Etiquetas")
        Exit Sub ' SI NO HAY NINGUN FRM CONFIGURADO
    End If
    '

    If JJ% > 1 Then ' SI ESTAN LOS 2 CONFIGURADOS
       'VERIFICA SI TIENE CONFIGURADO TAMAÑO EN A.B.M
       CONDI$ = "Codint = " & CIXI%
       TipoEti% = XVALO(VALOBADA("MASTER", "Tamano_Etiqueta", CONDI$, "NOMESS"))
       'SI NO TIENE CONFIGURACION DE TAMAÑO PREGUNTA
       If TipoEti% < 1 Or TipoEti% > 2 Then
         OPX% = ALTERNA%("Etiqueta Standard\Etiqueta Pequeña")
         If OPX% > 1 Then
           FORIPRE$ = FORIPR2$ ' SI ELIGE LA PEQUEÑA
         Else
           FORIPRE$ = FORIPR1$ ' SI ELIGE STANDARD O CANCELA IMPRIME STANDARD
         End If
         'SI TIENE CONFIGURADO TAMAÑO TOMA EL QUE ESTA CONFIGURADO SIN PREGUNTAR
       Else
         FORIPRE$ = FORIPR1$: If TipoEti% = 1 Then FORIPRE$ = FORIPR2$
       End If
    End If
    '
    STAT_FORIPRE$ = FORIPRE$
    
    'SI ES IMPORTADO IMPRIME CON OTRA ETIQUETA
    Dim obj As New Art_Importados
    Call obj.VeriImportado(CODIX$) 'VERIFICA QUE NO SEA IMPORTADO
    If obj.EsImportado > 0 Then  'SI ES IMPORTADO
       STAT_FORIPRE$ = FORIPR_IMP$
    End If

    If MODO% = 1 Then
        STAT_FORIPRE$ = TRIM$(STAT_FORIPRE$)
        STAT_FORIPRE$ = Mid(STAT_FORIPRE$, 1, Len(STAT_FORIPRE$) - 1) + "X"
    End If
    
    'OPCION DE MODIFICAR TENSION
'    If NOMENSATENSION% < 1 Then
      TXX1$ = ""
      If TENSION_ENSAYO > 0 Then
          TXX1$ = "Equipo Registrado Con Tensión de : " & TENSION_ENSAYO & "\"
      End If
      '
      If Left$(TRIM$(CODIX$), 2) = "11" Then
        CONSUMO$ = TALLE$(CIXI%)
      Else
        CONSUMO$ = COLOR$(CIXI%)
      End If
      
'      TXX1$ = TXX1$ + "Opciones de Impresión"
'      OPX1% = COMALTER%(TXX1$ & "\\Impresión a 110 V.\Impresión a 220 V.")
'      If OPX1% < 1 Or OPX1% > 2 Then Exit Sub
'27    If OPX1% = 1 Then CONSUMO$ = VALOBADA("MASTER", "TALLE", CONDI$, "NOMESS")
'      If OPX1% = 2 Then CONSUMO$ = VALOBADA("MASTER", "COLOR", CONDI$, "NOMESS")
''    End If


30  PFAC$ = CONTROL$(IDENTER$, "PORTSTIK")
    If PFAC$ = "" Then PFAC$ = CONTROL$("*", "PORTSTIK")
    If PFAC$ = "" Then Exit Sub ' POR SI NOHAY NINGUNA IMPRESORA CONFIGURADA
    ImpresoraPredeterminada$ = Printer.DeviceName
    Call SETPRINPORT2(PFAC$)
    Call APERBASDAT("STOCKS")
    '
'    SQLX$ = " SELECT * FROM TANUMSE "
'    SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CIXI% & " "
'    SQLX$ = SQLX$ + "AND NumeroSerie = '" & NUSERIE$ & "' "
'    '
'    Dim LOTINGTEMP As ADODB.Recordset
'    Set LOTINGTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    Codigo$ = CODIX$
    DESXRIP$ = DESCRIT0$(CIXI%)
    LOTEE$ = NUSERIE$
    CONDI$ = "NUMEROSERIE = '" & NUSERIE$ & "'"
    TENSION_ENSAYO = XVALO(VALOBADA("TANUMSE", "TEN_ENSA", CONDI$, "NOMESS"))
    '
    If CONTROL("BOLRECEP", "FORMATO") = "TXT" Then
        TXTX$ = TRIM$(FORIPRE$)
        If TRIM$(FORIPRE$) <> "" Then
            GoSub PRINZEBRA
          Else
            Call MENSERR(24, "Falta Definir Archivo Transferencia de Impresión")
        End If
        GoTo 90
    End If
    '
    TIPODOC$ = "Etiqueta Identificador de Lote"
    CODPARAM$(1) = "COD-MAT": DOCPARAM$(1) = Codigo$
    CODPARAM$(2) = "NRO-LOTE": DOCPARAM$(2) = LOTEE$
    CODPARAM$(3) = "DES-MAT": DOCPARAM$(3) = DESXRIP$
    CODPARAM$(4) = "TAL-CONJ": DOCPARAM$(4) = CONSUMO$
    
    CAPARDOC = 4
    '
    Call PRINTDOC("@" + STAT_FORIPRE$)
    '
90  Call SETPRINPORT2("RESTORE")

    Exit Sub
    '
PRINZEBRA:
    '
    TXTX$ = LOADFILE(LUDAT$ + FORIPRE$ + ".TXT")
    Call REPLATEXT(TXTX$, Chr$(13) + Chr$(10), Chr$(10))
    Call REPLATEXT(TXTX$, "@CODI-MAT", Codigo$)
    Call REPLATEXT(TXTX$, "@NRO-LOTE", LOTEE$)
    Call REPLATEXT(TXTX$, "@DEPOSITO", Str$(DEPOSITO%))
    Call REPLATEXT(TXTX$, "@UBI-MPRI", UBIFISICA$)
    Call REPLATEXT(TXTX$, "@MEDIDA-1", FORMATNUM(Largo#, "F10.1"))
    Call REPLATEXT(TXTX$, "@MEDIDA-2", FORMATNUM(ANCHO#, "F10.1"))
    Call REPLATEXT(TXTX$, "@CANTIDAD", FORMATNUM(Cantidad#, "F10.1"))
    Call REPLATEXT(TXTX$, "@DESC-MAT", DESXRIP$)
    Call REPLATEXT(TXTX$, "@REF-MAT1", AJUSTI$(Codigo$, 16) + AJUSTI$(LOTEE$, 12))
    '
    TXTX1$ = TXTX$
    ARCHOUT$ = LUCOM$ + FORIPRE$ + ".PRN"
    Call SAVEFILE(ARCHOUT$, TXTX$)
    'PFAC$ = TRIM$(CONTROL$(IDENTER$, "PORTSTIK"))
    'Call DOSCOMMAND("PRINT " + ARCHOUT$)
    '
75  On Error Resume Next
    Printer.Print TXTX1$
    If Err Then
       On Error GoTo 0
       OPX% = COMALTER%("Error de la Impresora\\Re-Intentar\Abandonar")
       If OPX% = 1 Then GoTo 75
    End If
    On Error GoTo 0
    '
Return
'

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call CIERRARCH("*.*")
    Call FINAL("")

End Sub

Private Sub Frame5_DragDrop(Source As CONTROL, X As Single, Y As Single)

End Sub


