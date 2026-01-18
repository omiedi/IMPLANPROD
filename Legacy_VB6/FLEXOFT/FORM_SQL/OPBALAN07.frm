VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPBALAN07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0095C5E1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Balance General - Sumas y Saldos"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H0095C5E1&
      Caption         =   "SELECCIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1550
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   8325
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Cuentas con Movimiento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   3
         Left            =   360
         TabIndex        =   22
         Top             =   480
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Todas las Cuentas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Index           =   0
         Left            =   360
         TabIndex        =   21
         Top             =   960
         Width           =   1260
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7845
         TabIndex        =   20
         Top             =   465
         Width           =   175
      End
      Begin VB.TextBox Text8 
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
         Left            =   6375
         TabIndex        =   19
         Text            =   " "
         Top             =   465
         Width           =   1485
      End
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4605
         TabIndex        =   18
         Top             =   480
         Width           =   175
      End
      Begin VB.TextBox Text10 
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
         Left            =   3135
         TabIndex        =   17
         Top             =   480
         Width           =   1485
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5400
         TabIndex        =   26
         Top             =   570
         Width           =   885
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2160
         TabIndex        =   25
         Top             =   585
         Width           =   885
      End
      Begin VB.Label Label1 
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
         Height          =   495
         Left            =   5640
         TabIndex        =   24
         Top             =   870
         Width           =   2415
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
         Height          =   495
         Left            =   2400
         TabIndex        =   23
         Top             =   870
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   3480
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H0095C5E1&
      Caption         =   "PERÍODO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   3060
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   420
         TabIndex        =   14
         Top             =   420
         Width           =   2100
      End
      Begin VB.CommandButton Command6 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2520
         TabIndex        =   13
         Top             =   420
         Width           =   175
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Asiento Apertura"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1000
         TabIndex        =   12
         Top             =   880
         Width           =   1770
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Asientos de Cierre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   840
         TabIndex        =   11
         Top             =   1155
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Excluir:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   910
         Width           =   795
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H0095C5E1&
      Caption         =   "FORMATO DE SALIDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   6510
      TabIndex        =   5
      Top             =   240
      Width           =   2805
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Completo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   360
         TabIndex        =   9
         Top             =   300
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Condensado Nivel 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   360
         TabIndex        =   8
         Top             =   570
         Width           =   1875
      End
      Begin VB.OptionButton Option4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Nivel 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1320
         TabIndex        =   7
         Top             =   850
         Width           =   920
      End
      Begin VB.OptionButton Option5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Nivel 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1320
         TabIndex        =   6
         Top             =   1150
         Width           =   920
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0095C5E1&
      Caption         =   "TIPO DE BALANCE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   3255
      TabIndex        =   1
      Top             =   240
      Width           =   3165
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Saldos - Estado de Cuentas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   780
         Width           =   2700
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Resultados Mensuales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   2370
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Sumas y Saldos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Value           =   -1  'True
         Width           =   1860
      End
   End
   Begin VB.CommandButton Command1 
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
      Height          =   675
      Left            =   8580
      Picture         =   "OPBALAN07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2670
      Width           =   750
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "OPBALAN07.frx":030A
      Top             =   120
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPBALAN07.frx":053E
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "OPBALAN07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   Call BALSUSAL
   Command1.Enabled = True
End Sub

Private Sub Command6_Click()
   Call FRESHALL
   Call CARSELMESA
10 Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   Text6.TEXT = MESASEL$
End Sub

Private Sub Command8_Click()
    Call SELECHO("ECUECOT")
    PRICUEX$ = VALACT1$("ECUECOT")
    If PRICUEX$ <> "" Then
        Text8.TEXT = PRICUEX$
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("ECUECOT")
    PRICUEX$ = VALACT1$("ECUECOT")
    If PRICUEX$ <> "" Then
        Text10.TEXT = PRICUEX$
    End If
End Sub

Private Sub Form_Load()
    '
'    EPAC$ = TRIM$(EMPREAC$)
'    If EPAC$ <> "" Then
'        Caption = Caption + "  -  " + EPAC$
'    End If
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    '
    Text10 = Left$(REGLEIDO$("ECUECOT", 1), 12)
    Text8 = Left$(REGLEIDO$("ECUECOT", ULTREG&("ECUECOT")), 12)
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub
'
Sub BALSUSAL()  ' balance de sumas y saldos
   '
   Dim NUMECO$(), SUMASIT#(), FECO%()
   ReDim NUMECO$(1024), SUMASIT#(1024), FECO%(1024)
   CANCOM& = 0
   '
   MMXX0$ = TRIM$(Text6.TEXT)
   MMXX1$ = MESANO$(MMXX0$, Des%, Has%)
   If MMXX1$ <> "" Then GoTo 12
   '
   If Option1(4).Value = True Then
5     VDEF$ = MKI$(COMEJ%) + MKI$(FINEJ%)
      OBXX$ = OBJPLA$("DESHASFECHA", VDEF$)
      If OBXX$ = "" Then Exit Sub
      Des% = CVI(Left$(OBXX$, 2))
      Has% = CVI(Mid$(OBXX$, 3, 2))
      If Has% < Des% Then GoTo 5
      If Des% < COMEJ% Or Has% > FINEJ% Then GoTo 5
      FEX = Des%: MMXX1$ = FECHATEX$(FEX)
      FEX = Has%: MMXX1$ = MMXX1$ + " - " + FECHATEX$(FEX)
      Text6.TEXT = MMXX1$
      Text6.Refresh
      GoTo 13
   End If
   '
   
   Call FRESHALL
   Call CARSELMESA
9 Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) = "" Then
      Exit Sub
   End If
   '
   MMXX0$ = TRIM$(MESASEL$)
   MMXX1$ = MESANO$(MMXX0$, Des%, Has%)
   If TRIM$(MMXX1$) = "" Then
      Exit Sub
   End If
   '
12 MESASEL$ = MMXX1$
   If Des% < COMEJ% Or Has% > FINEJ% Then
      Call MENSERR(24, "Período Fuera de Límites\del Ejercicio Activo.")
      GoTo 9
   End If
   '

   Text6.TEXT = MESASEL$
   Text6.Refresh
   '
13 Todas% = 0: If Option1(0).Value = True Then Todas% = 1
   NIVCOND% = 4
   If Option3.Value = True Then
         NIVCOND% = 3
      ElseIf Option4.Value = True Then
         NIVCOND% = 2
      ElseIf Option5.Value = True Then
         NIVCOND% = 1
   End If
   TIBA% = 1                           'BALANCE MENSUAL
   TIBALAN$ = "Balance Mensual de " + MESASEL$
   If Option1(2) = True Then           'ESTADO DE CUENTAS
      TIBA% = 2
      TIBALAN$ = "Estado de Cuentas a " + MESASEL$
   End If
   If Option1(4) = True Then
      TIBA% = 3 'SUMAS Y SALDOS
      TIBALAN$ = "Balance de Sumas y Saldos de " + MESASEL$
   End If
   '
   Screen.MousePointer = 11
   Dim SALINI1#(16384)
   Dim SUDEBE1#(16384)
   Dim SUHABE1#(16384)
   '
   'Acumuladores para suma total de columnas y muestra
   SUMAINI# = 0
   SUMADEB# = 0
   SUMAHAB# = 0
   RESU# = 0
   SALD# = 0
   NCUEMAX% = 0
   NICONFLIM% = NIVCONFI%
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   '
   ' ***** OJO QUE SI SE MODIFICA EL CRITERIO ACÁ, TAMBIÉN EN ****
   ' ***** LA RUTINA DE GENERACIÓN DEL ASIENTO DE CIERRE !!!! ****
    '
   'DES% = DIANTE(14817, 5)
'   HAS% = DIANTE(HAS%, 1)

10  SQLX$ = "SELECT cuecoint, SUM(ROUND(IDEBEPASE,2)) AS SALDOIND, SUM(ROUND(IHABEPASE,2)) AS SALDOINH FROM DETAASIEN "
    SQLX$ = SQLX$ + " WHERE FECHASIEN  < '" & FETEXCOR1$(Des%) & "' "
    ' SQLX$ = SQLX$ + "AND SUCUASIen = " & SUCUI% & " "
    SQLX$ = SQLX$ + "AND STATPASE > 0 "
    SQLX$ = SQLX$ + "AND Niv_Clasi <= " & NICONFLIM% & " "
    SQLX$ = SQLX$ + "GROUP BY CUECOINT "
    '
    Dim MODIASEL As ADODB.Recordset
    Set MODIASEL = New ADODB.Recordset
    On Error Resume Next
    MODIASEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then
       On Error GoTo 0
       GoTo 15
    End If
    On Error GoTo 0
    '
    With MODIASEL
       If Not .BOF And Not .EOF Then
         Do While Not .EOF
           NCLIE% = !CUECOINT
           If NCLIE% > 0 Then
               If NCLIE% > NCUEMAX% Then NCUEMAX% = NCLIE%
               SALINI1#(NCLIE%) = SALINI1#(NCLIE%) + !SALDOIND - !SALDOINH
           End If
         .MoveNext
         Loop
       End If
    End With
    MODIASEL.Close
    '
15  SQLX$ = "SELECT cuecoint, SUM(ROUND(IDEBEPASE,2)) AS sumadebe, SUM(ROUND(IHABEPASE,2)) AS SUMAHABE "
    SQLX$ = SQLX$ + " FROM DETAASIEN "
    SQLX$ = SQLX$ + " WHERE (FECHASIEN  >= '" & FETEXCOR1$(Des%) & "' AND FECHASIEN <= '" & FETEXCOR1$(Has%, "HASTA") & "') "
    ' SQLX$ = SQLX$ + "AND SUCUASIen = " & SUCUI% & " "
    SQLX$ = SQLX$ + "AND STATPASE > 0 "
    SQLX$ = SQLX$ + "AND Niv_Clasi <= " & NICONFLIM% & " "
    SQLX$ = SQLX$ + "GROUP BY CUECOINT "
    '
    Set MODIASEL = New ADODB.Recordset
    On Error Resume Next
    MODIASEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then
       On Error GoTo 0
       GoTo 20
    End If
    On Error GoTo 0
    '
    With MODIASEL
       If Not .BOF And Not .EOF Then
         Do While Not .EOF
           NCLIE% = !CUECOINT
           If NCLIE% > 0 Then
               SUDEBE1#(NCLIE%) = SUDEBE1#(NCLIE%) + ROUND(!SUMADEBE)
               SUHABE1#(NCLIE%) = SUHABE1#(NCLIE%) + ROUND(!sumahabe)
           End If
         .MoveNext
         Loop
       End If
    End With
    MODIASEL.Close
    '
    ' OPCION EXCLUIR ASIENTO DE APERTURA
20  If Check2.Value = 1 Then
       If Des% <= COMEJ% Then
          SQLX$ = "SELECT CUECOINT, SUM(ROUND(IDEBEPASE,2)) AS SUMADEBE, SUM(ROUND(IHABEPASE,2)) AS SUMAHABE "
          SQLX$ = SQLX$ + "FROM DETAASIEN "
          SQLX$ = SQLX$ + "INNER JOIN ENCABEASIEN ON ENCABEASIEN.NUIDASIEN = DETAASIEN.NUIDASIEN "
          SQLX$ = SQLX$ + "WHERE ENCABEASIEN.STATASIEN > 0 "
          SQLX$ = SQLX$ + "AND ENCABEASIEN.TIPOASIEN = 8 "
          SQLX$ = SQLX$ + "AND ENCABEASIEN.FECHASIEN  = '" & FETEXCOR1$(COMEJ%) & "' "
          ' SQLX$ = SQLX$ + "AND ENCABEASIEN.SUCUASIEN = " & SUCUI% & " "
          SQLX$ = SQLX$ + "AND DETAASIEN.STATPASE > 0 "
          SQLX$ = SQLX$ + "AND DETAASIEN.Niv_Clasi <= " & NICONFLIM% & " "
          SQLX$ = SQLX$ + "GROUP BY CUECOINT "
          '
          Set MODIASEL = New ADODB.Recordset
          On Error Resume Next
          MODIASEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
          If Err Then
             On Error GoTo 0
             GoTo 30
          End If
          On Error GoTo 0
          
          With MODIASEL
             If Not .BOF And Not .EOF Then
               Do While Not .EOF
                 NCLIE% = !CUECOINT
                 If NCLIE% > 0 Then
                     If NCLIE% > NCUEMAX% Then NCUEMAX% = NCLIE%
                     SALINI1#(NCLIE%) = SALINI1(NCLIE%) + ROUND(!SUMADEBE) - ROUND(!sumahabe)
                     SUDEBE1#(NCLIE%) = SUDEBE1#(NCLIE%) - ROUND(!SUMADEBE)
                     SUHABE1#(NCLIE%) = SUHABE1#(NCLIE%) - ROUND(!sumahabe)
                 End If
               .MoveNext
               Loop
             End If
          End With
          MODIASEL.Close
       End If
    End If
    '
    '
    ' OPCION EXCLUIR ASIENTO DE CIERRE
30  If Check1.Value = 1 Then
       If Has% >= FINEJ% Then
          SQLX$ = "SELECT CUECOINT, SUM(ROUND(IDEBEPASE,2)) as SUMADEBE, SUM(ROUND(IHABEPASE,2)) AS SUMAHABE "
          SQLX$ = SQLX$ + "FROM DETAASIEN "
          SQLX$ = SQLX$ + "INNER JOIN ENCABEASIEN ON ENCABEASIEN.NUIDASIEN = DETAASIEN.NUIDASIEN "
          SQLX$ = SQLX$ + "WHERE ENCABEASIEN.STATASIEN > 0 "
          SQLX$ = SQLX$ + "AND ENCABEASIEN.TIPOASIEN = 9 "
          SQLX$ = SQLX$ + "AND ENCABEASIEN.FECHASIEN  = '" & FETEXCOR1$(FINEJ%) & "' "
          ' SQLX$ = SQLX$ + "AND ENCABEASIEN.SUCUASIEN = " & SUCUI% & " "
          SQLX$ = SQLX$ + "AND DETAASIEN.STATPASE > 0 "
          SQLX$ = SQLX$ + "AND DETAASIEN.Niv_Clasi <= " & NICONFLIM% & " "
          SQLX$ = SQLX$ + "GROUP BY CUECOINT "
          '
          Set MODIASEL = New ADODB.Recordset
          On Error Resume Next
          MODIASEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
          If Err Then
             On Error GoTo 0
             GoTo 40
          End If
          On Error GoTo 0
          
          With MODIASEL
             If Not .BOF And Not .EOF Then
               Do While Not .EOF
                 NCLIE% = !CUECOINT
                 If NCLIE% > 0 Then
                     SUDEBE1#(NCLIE%) = SUDEBE1#(NCLIE%) - ROUND(!SUMADEBE)
                     SUHABE1#(NCLIE%) = SUHABE1#(NCLIE%) - ROUND(!sumahabe)
                 End If
               .MoveNext
               Loop
             End If
          End With
          MODIASEL.Close
       End If
    End If
    '
    '
'   DESDE AQUI SE CARGAN LOS MOVIMIENTOS DE CAJA Y BANCOS
'   queda pendiente la verificacion del balance de los asientos
40  'SQLX$ = "SELECT CueContaI, SUM(IDEBECOMP - IHABECOMP) AS SALDOIN FROM CAJABANC "
    SQLX$ = "SELECT CueContaI, SUM(ROUND(IDEBECOMP,2)) as sumadebe, SUM(ROUND(IHABECOMP,2)) AS sumahabe FROM CAJABANC "
    SQLX$ = SQLX$ + " WHERE FechConta  < '" & FETEXCOR1$(Des%) & "' "
    SQLX$ = SQLX$ + " AND (NuAsidiar = '' or NuAsiDiar IS NULL)  "
    SQLX$ = SQLX$ + " AND Niv_Clasi <= " & NICONFLIM% & " "
    SQLX$ = SQLX$ + " GROUP BY CueContaI "
    '
    Dim CABTEMP As ADODB.Recordset
    Set CABTEMP = New ADODB.Recordset
    On Error Resume Next
    CABTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then
       On Error GoTo 0
       GoTo 90
    End If
    On Error GoTo 0
    '
    With CABTEMP
       If Not .BOF And Not .EOF Then
         Do While Not .EOF
           CUE% = !CUECONTAI
           If CUE% > 0 Then
               If CUE% > NCUEMAX% Then NCUEMAX% = CUE%
               SALINI1#(CUE%) = SALINI1#(CUE%) + (ROUND(!SUMADEBE) - ROUND(!sumahabe))
           End If
         .MoveNext
         Loop
       End If
    End With
    CABTEMP.Close
    '
    SQLX$ = "SELECT CueContaI, SUM(ROUND(IDEBECOMP,2)) as sumadebe, SUM(ROUND(IHABECOMP,2)) AS SUMAHABE FROM CAJABANC "
    SQLX$ = SQLX$ + " WHERE (FechConta  >= '" & FETEXCOR1$(Des%) & "' AND FechConta <= '" & FETEXCOR1$(Has%, "HASTA") & "') "
    SQLX$ = SQLX$ + " AND (NuAsidiar = '' or NuAsiDiar IS NULL)  "
    SQLX$ = SQLX$ + " AND Niv_Clasi <= " & NICONFLIM% & " "
    SQLX$ = SQLX$ + " GROUP BY CueContaI"
    Set CABTEMP = New ADODB.Recordset
    On Error Resume Next
    CABTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then
       On Error GoTo 0
       GoTo 90
    End If
    On Error GoTo 0
    '
    With CABTEMP
       If Not .BOF And Not .EOF Then
         Do While Not .EOF
           CUE% = !CUECONTAI
           If CUE% > 0 Then
               SUDEBE1#(CUE%) = SUDEBE1#(CUE%) + !SUMADEBE
               SUHABE1#(CUE%) = SUHABE1#(CUE%) + !sumahabe
               If CUE% = 363 Then
               A = B
               End If
           End If
         .MoveNext
         Loop
       End If
    End With
   CABTEMP.Close
   Set CABATEMP = Nothing
   '
90 On TIBA% GoTo 100, 100, 300
   Call CIERRARCH("*.*")
   Exit Sub
   '
100 ' BALANCE MENSUAL *******************************
   '
'esta parte resolverla similar a 300 en adelante

   On Error Resume Next
   'Contable.Execute "DROP TABLE BALASUSA"
   'Contable.Execute "SELECT codexcue, denocue, codincue INTO Balasusa FROM Plancuen;"
   'JANDY SQL
   strsql = "IF EXISTS (SELECT name FROM sysdatabases WHERE name = 'BALASUSA')"
   strsql = strsql & " DROP TABLE BALASUSA"
   strsql = strsql & " SELECT codexcue, denocue, codincue INTO Balasusa FROM Plancuen"
    FLEXCONN.Execute (FILTSQL$(strsql))
   On Error GoTo 0
   
   '
   Dim BALAPRIM As ADODB.Recordset
   Set BALAPRIM = New ADODB.Recordset
   SQLX$ = "SELECT NIVAGRU, CODEXCUE, DENOCUE, CODINCUE, SALINI, SUDEBE, SUHABE, RESULT, SALFIN, NuOrdIt "
   '\\\OT-5-549-FG-29/07/05
   SQLX$ = SQLX$ + " FROM PLANCUEN ORDER BY NuOrdIt ASC"
   '\\\OT-5-549-FG-FIN
   
   'Set BALAPRIM = Contable.OpenRecordset(SQLX$)
   BALAPRIM.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   
   '
   CACUEN& = 0
   On Error Resume Next
   With BALAPRIM
     If Not (.EOF And .BOF) Then
     'FIN& = .RecordCount
     'For U& = 1 To FIN&
       'If U& = 1 Then
            '.MoveFirst
          'Else
            '.MoveNext
       Do While (Not .EOF)
        CUEI% = !CODINCUE
        '.Edit
        If CUEI% > NCUEMAX% Then NCUEMAX% = CUEI%
        !SALINI = SALINI1#(CUEI%)
        !SUDEBE = SUDEBE1#(CUEI%)
        !SUHABE = SUHABE1#(CUEI%)
        !Result = SUDEBE1#(CUEI%) - SUHABE1#(CUEI%)
        !SALFIN = SALINI1#(CUEI%) + SUDEBE1#(CUEI%) - SUHABE1#(CUEI%)
        .Update
        .MoveNext
       Loop
     'Next U&
     End If
   End With
   BALAPRIM.Close
   Set BALAPRIM = Nothing
   '
   Call ABREPLANCUEN
   SQLX$ = "SELECT NIVAGRU, CODEXCUE, DENOCUE, CODINCUE, SALINI, SUDEBE, SUHABE, RESULT, SALFIN "
   '\\\OT-5-549-FG-29/07/05
   SQLX$ = SQLX$ + " FROM PLANCUEN "
   '\\\OT-5-549-FG-FIN
   ACONDI% = 0
   '
   If TRIM$(Text10) <> "" Then
     If ACONDI% > 0 Then
         SQLX$ = SQLX$ + "AND "
       Else
         SQLX$ = SQLX$ + "WHERE "
     End If
     SQLX$ = SQLX$ + "CODEXCUE >= '" & TRIM$(Text10) & "' "
     ACONDI% = 1
   End If
   '
   If TRIM$(Text8) <> "" Then
     If ACONDI% > 0 Then
         SQLX$ = SQLX$ + "AND "
       Else
         SQLX$ = SQLX$ + "WHERE "
     End If
     SQLX$ = SQLX$ + "CODEXCUE <= '" & TRIM$(Text8) & "' "
   End If
   SQLX$ = SQLX$ + "ORDER BY NuOrdIt ASC"
   '
   Dim Balasec As ADODB.Recordset
   Set Balasec = New ADODB.Recordset
   
   'Set BALASEC = Contable.OpenRecordset(SQLX$)
   Balasec.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   
   Balasec.MoveLast
   If Balasec.RecordCount < 1 Then
     Call MENSERR(24, "Sin Registros Activos")
     Screen.MousePointer = 1
     Balasec.Close
     Set Balasec = Nothing
     Exit Sub
   End If
   '
   'If NXP% < 1 Then
   '   NXP% = FILEOPEN%(LUDAT$ + "PLAGRAL" + "." + EXTE$, 0, 4)
   'End If
   '
   JJ& = 0: KLJ& = 0
   Dim IVP As String * 4
   Dim SUNIVEL#(5), DENIVEL$(5)
   '
   With Balasec
     FIN& = .RecordCount
     .MoveFirst
     UU& = 1
     Do While Not .EOF
       Call TRACE(20, UU&, FIN&)
       CUE% = !CODINCUE         'CVI(Mid$(IVP$, 1, 2))
       NIVEI% = !NIVAGRU        'CVI(Mid$(IVP$, 3, 2))
       PRONI% = NIVEI%
       If Not .EOF Then
         .MoveNext
         PRONI% = !NIVAGRU
         .MovePrevious
       End If
       '
       If CUE% > 0 Then
         If CUE% <= 16384 Then
           IMPO# = SUDEBE1#(CUE%) - SUHABE1#(CUE%)
           If TIBA% = 2 Then IMPO# = IMPO# + SALINI1#(CUE%)
           If NIVEI% > NIVANT% Then
               If NIVEI% <= NIVCOND% Then
                 ZZ$ = Space$(128)
                 JJ& = JJ& + 1
                 Call GRAREG("BALANCE", ZZ$, JJ&)
               End If
             ElseIf NIVEI% < NIVANT% Then
               For KQ% = NIVANT% - 1 To NIVEI% Step -1
                 If KQ% <= NIVCOND% Then
                    If Todas% = 1 Or Abs(SUNIVEL#(KQ%)) > 0.005 Then
                      ZZ$ = Space$(128)
                      JJ& = JJ& + 1
                      Call GRAREG("BALANCE", ZZ$, JJ&)
                      ZZ$ = "." + Space$(127)
                      QQ$ = Space$(88) + String$(40, 0)
                      TXTX$ = "TOTAL " + DENIVEL$(KQ%)
                      Call REPLA(QQ$, TXTX$, 15, 74)
                      If Abs(SUNIVEL#(KQ%)) >= 0.005 Then TXTX$ = TXTX$ + " " + String$(128, ".")
                      ITXT$ = CSTRING$(MKD$(SUNIVEL#(KQ%)), 4, 12, 2, 2) + Space$(128)
                      Call REPLA(ZZ$, TXTX$, 1 + 5 * KQ%, 128)
                      Call REPLA(ZZ$, ITXT$, 117 - 10 * KQ%, 128)
                      JJ& = JJ& + 1
                      Call GRAREG("BALANCE", ZZ$, JJ&)
                      '
                      Call REPLA(QQ$, MKD$(SUNIVEL#(KQ%)), 121 - 8 * KQ%, 8)
                      KLJ& = KLJ& + 1
                      Call GRAREG("BALANCF", QQ$, KLJ&)
                    End If
                  End If
                  SUNIVEL#(KQ%) = 0
                  DENIVEL$(KQ%) = ""
                Next KQ%
                If NIVEI% <= NIVCOND% Then
                  ZZ$ = Space$(128)
                  JJ& = JJ& + 1
                  Call GRAREG("BALANCE", ZZ$, JJ&)
                End If
           End If
           '
           If NIVEI% <= NIVCOND% Then
             If Todas% = 1 Or Abs(IMPO#) >= 0.005 Or PRONI% > NIVEI% Then
                ZZ$ = "." + Space$(127)
                QQ$ = Space$(88) + String$(40, 0)
                TXTX$ = TRIM$(CODCUE$(CUE%)) + " " + TRIM$(DENOCUE$(CUE%))
                Call REPLA(QQ$, CODCUE$(CUE%), 1, 14)
                Call REPLA(QQ$, DENOCUE$(CUE%), 15, 74)
                If Abs(IMPO#) >= 0.005 Then TXTX$ = TXTX$ + " " + String$(128, ".")
                ITXT$ = CSTRING$(MKD$(IMPO#), 4, 12, 2, 2) + Space$(128)
                Call REPLA(ZZ$, TXTX$, 1 + 5 * NIVEI%, 128)
                Call REPLA(ZZ$, ITXT$, 117 - 10 * NIVEI%, 128)
                JJ& = JJ& + 1
                Call GRAREG("BALANCE", ZZ$, JJ&)
                '
                Call REPLA(QQ$, MKD$(IMPO#), 121 - 8 * NIVEI%, 8)
                KLJ& = KLJ& + 1
                Call GRAREG("BALANCF", QQ$, KLJ&)
             End If
           End If
           '
           For KQ% = 0 To NIVEI% - 1
             SUNIVEL#(KQ%) = SUNIVEL#(KQ%) + IMPO#
           Next KQ%
           SUNIVEL#(NIVEI%) = IMPO#
           DENIVEL$(NIVEI%) = TRIM$(DENOCUE$(CUE%))
           NIVANT% = NIVEI%
           '
         End If
       End If
       '
       .MoveNext
       UU& = UU& + 1
     Loop
   End With
   Balasec.Close
   Set Balasec = Nothing
   '
   'JJ& = 0
   For KQ% = NIVANT% - 1 To 0 Step -1   'NIVEI% Step -1
      If KQ% <= NIVCOND% Then
         If Todas% = 1 Or Abs(SUNIVEL#(KQ%)) >= 0.005 Then
            ZZ$ = Space$(128)
            JJ& = JJ& + 1
            Call GRAREG("BALANCE", ZZ$, JJ&)
            ZZ$ = "." + Space$(127)
            QQ$ = Space$(88) + String$(40, 0)
            TXTX$ = "TOTAL " + DENIVEL$(KQ%)
            Call REPLA(QQ$, TXTX$, 15, 74)
            If Abs(SUNIVEL#(KQ%)) >= 0.005 Then TXTX$ = TXTX$ + " " + String$(128, ".")
            ITXT$ = CSTRING$(MKD$(SUNIVEL#(KQ%)), 4, 12, 2, 2) + Space$(128)
            Call REPLA(ZZ$, TXTX$, 1 + 5 * KQ%, 128)
            Call REPLA(ZZ$, ITXT$, 117 - 10 * KQ%, 128)
            JJ& = JJ& + 1
            Call GRAREG("BALANCE", ZZ$, JJ&)
            '
            Call REPLA(QQ$, MKD$(SUNIVEL#(KQ%)), 121 - 8 * KQ%, 8)
            KLJ& = KLJ& + 1
            Call GRAREG("BALANCF", QQ$, KLJ&)
         End If
      End If
      SUNIVEL#(KQ%) = 0
      DENIVEL$(KQ%) = ""
   Next KQ%
   ZZ$ = Space$(128)
   JJ& = JJ& + 1
   Call GRAREG("BALANCE", ZZ$, JJ&)
   '
   Call SETULTREG("BALANCE", JJ&)
   Call CIERRARCH("BALANCE")
   '
   Call SETULTREG("BALANCF", KLJ&)
   Call CIERRARCH("BALANCF")
   '
   Call HEADERS("BALANCE", "")
   Call HEADERS("BALANCE", "Corresponde a: " + MESASEL$)
   '
   Call HEADERS("BALANCF", "")
   Call HEADERS("BALANCF", "Corresponde a: " + MESASEL$)
   '
   Close #NXP%: NXP% = 0
   Screen.MousePointer = 1
   '
   OPX1% = COMALTER%("Elija Formato de Salida:\\Planilla Plana\Exportable a Excel")
   If OPX1% = 1 Then
       Call CONECRUN("*BALANCE", "Balance General")
     ElseIf OPX1% = 2 Then
       Call CONECRUN("*BALANCF", "Balance General")
   End If
   Exit Sub
   '
   '
300 ' BALANCE DE SUMAS Y SALDOS *********************
   '
   For KKI% = 1 To NCUEMAX%
      If Abs(SALINI1#(KKI%)) >= 0.005 Or Abs(SUDEBE1#(KKI%)) >= 0.005 Or Abs(SUHABE1#(KKI%)) >= 0.005 Then
        CDXX$ = CODCUE$(KKI%)
        If CDXX$ = "" Then
            'Call COMUNI("Falta Cuenta " & KKI%)
            Call FaltaCuenta(KKI%)
        End If
      End If
   Next KKI%
   '
   SQLX$ = "SELECT CODEXCUE, DENOCUE, CODINCUE, SALINI, SUDEBE, SUHABE, RESULT, SALFIN "
   SQLX$ = SQLX$ + " FROM PLANCUEN "
   '
   Dim BALAPRIM1 As ADODB.Recordset
   Set BALAPRIM1 = New ADODB.Recordset
   BALAPRIM1.CursorType = adOpenStatic
   BALAPRIM1.LockType = adLockPessimistic
   BALAPRIM1.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
   '
   CACUEN& = 0
   With BALAPRIM1
     Do While Not .EOF
       CUEI% = !CODINCUE
       !SALINI = SALINI1#(CUEI%)
       !SUDEBE = SUDEBE1#(CUEI%)
       !SUHABE = SUHABE1#(CUEI%)
       !Result = SUDEBE1#(CUEI%) - SUHABE1#(CUEI%)
       !SALFIN = SALINI1#(CUEI%) + SUDEBE1#(CUEI%) - SUHABE1#(CUEI%)
       .Update
       SUMAINI# = SUMAINI# + SALINI1#(CUEI%)
       SUMADEB# = SUMADEB# + SUDEBE1#(CUEI%)
       SUMAHAB# = SUMAHAB# + SUHABE1#(CUEI%)
     .MoveNext
     Loop
   End With
   BALAPRIM1.Close
   Set BALAPRIM1 = Nothing
   '
   '
   SQLX$ = "SELECT CODEXCUE, DENOCUE, CODINCUE, SALINI, SUDEBE, SUHABE, RESULT, SALFIN ,NuOrdIt"
   SQLX$ = SQLX$ + " FROM PLANCUEN "
   ACONDI% = 0
   If Option1(3) = True Then
      SQLX$ = SQLX$ + "WHERE (SALINI <> 0 OR SUDEBE <> 0 OR SUHABE <> 0 "
      SQLX$ = SQLX$ + "OR RESULT > 0 OR SALFIN > 0 ) "
      ACONDI% = 1
   End If
   '
   If TRIM$(Text10) <> "" Then
     If ACONDI% > 0 Then
         SQLX$ = SQLX$ + "AND "
       Else
         SQLX$ = SQLX$ + "WHERE "
     End If
     SQLX$ = SQLX$ + "CODEXCUE >= '" & TRIM$(Text10) & "' "
     ACONDI% = 1
   End If
   '
   If TRIM$(Text8) <> "" Then
     If ACONDI% > 0 Then
         SQLX$ = SQLX$ + "AND "
       Else
         SQLX$ = SQLX$ + "WHERE "
     End If
     SQLX$ = SQLX$ + "CODEXCUE <= '" & TRIM$(Text8) & "' "
   End If
   SQLX$ = SQLX$ + " ORDER BY NuOrdIt ASC "
   '
   '
   Dim BALASEC1 As ADODB.Recordset
   Set BALASEC1 = New ADODB.Recordset
   BALASEC1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   '
   JJ& = 0
   With BALASEC1
      Do While Not .EOF
         TTXYZ$ = AJUSTI$(!CodExCue, 13)
         TTXYZ$ = TTXYZ$ + AJUSTI$(!DENOCUE, 30)
         TTXYZ$ = TTXYZ$ + FORMATNUM$(!SALINI, "F13.2")
         TTXYZ$ = TTXYZ$ + FORMATNUM$(!SUDEBE, "F13.2")
         TTXYZ$ = TTXYZ$ + FORMATNUM$(!SUHABE, "F13.2")
         TTXYZ$ = TTXYZ$ + FORMATNUM$(!Result, "F13.2")
         TTXYZ$ = TTXYZ$ + FORMATNUM$(!SALFIN, "F13.2")
         ' en lugar de cargarlo a la grilla se agrega al list
         BASUMSAL07.List1.AddItem TTXYZ$
         '
         ZZ$ = REGBLAN$("BASUSAL")
         Call REPLA(ZZ$, .Fields(0).Value, 1, 12)
         Call REPLA(ZZ$, .Fields(1).Value, 13, 52)
         Call REPLA(ZZ$, MKD$(.Fields(3).Value), 65, 8)
         Call REPLA(ZZ$, MKD$(.Fields(4).Value), 73, 8)
         Call REPLA(ZZ$, MKD$(.Fields(5).Value), 81, 8)
         Call REPLA(ZZ$, MKD$(.Fields(6).Value), 89, 8)
         Call REPLA(ZZ$, MKD$(.Fields(7).Value), 97, 8)
         JJ& = JJ& + 1
         Call GRAREG("!BASUSAL", ZZ$, JJ&)
     .MoveNext
     Loop
   End With
   '
   Call SETULTREG("!BASUSAL", JJ&)
   Call CIERRARCH("!BASUSAL")
   Call TRACENTRAL("BASUSAL", "")
   '
   ' poner los encabezados de las columnas
   Screen.MousePointer = 1
   '
   BASUMSAL07.Label2 = Text6
   BASUMSAL07.Label3.Caption = FORMATNUM$(SUMAINI#, "F12.2")
   BASUMSAL07.Label4.Caption = FORMATNUM$(SUMADEB#, "F12.2")
   BASUMSAL07.Label5.Caption = FORMATNUM$(SUMAHAB#, "F12.2")
   RESU# = SUMADEB# - SUMAHAB#
   SALD# = SUMAINI# + RESU#
   BASUMSAL07.Label6.Caption = FORMATNUM$(RESU#, "F12.2")
   BASUMSAL07.Label7.Caption = FORMATNUM$(SALD#, "f12.2")
   
   BASUMSAL07.Show 1
   Exit Sub
   '
End Sub

Private Sub Text10_Change()
    Label2 = ECOARCH("ECUECOT", Text10.TEXT)
End Sub

Private Sub Text6_DblClick()
   Call FRESHALL
   Call CARSELMESA
10 Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   Text6.TEXT = MESASEL$
End Sub

Private Sub Text8_Change()
    Label1 = ECOARCH("ECUECOT", Text8.TEXT)
End Sub

Sub FaltaCuenta(NCUEN%)
10  DENOMICUEN$ = InputBox$("Ingrese Descripción de la cuenta Nro." & NCUEN%, "Alta Automática de Cuenta Faltante", "Cuenta Numero " & NCUEN%)
    If DENOMICUEN$ <> "" Then
       'CONSIGO EL VALOR MAXIMO DE NUORDIT
        SQLX$ = "SELECT MAX(NUORDIT)AS MAXORD"
        SQLX$ = SQLX$ + " FROM PLANCUEN"
        Dim MAXXXIMO As ADODB.Recordset
        Set MAXXXIMO = FLEXCONN.Execute(FILTSQL$(SQLX$))
        VALMAX& = MAXXXIMO!MAXORD + 1
        '
        'AGREGO LA CUENTA
        SQLX$ = "SELECT * FROM PLANCUEN "
        SQLX$ = SQLX$ + " Where NUORDIT = " & VALMAX& & " "
        Dim PLLLANCUEN As ADODB.Recordset
        Set PLLLANCUEN = New ADODB.Recordset
        PLLLANCUEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
        With PLLLANCUEN
           .AddNew   ' YA CONVERTIDO
           On Error Resume Next
           !CodiEmpr = CodiEmp%
           On Error GoTo 0
           !NIVAGRU = 0
           !DENOCUE = DENOMICUEN$
           !CodExCue = "CREC-" & NCUEN%
           !CODINCUE = NCUEN%
           !NUORDIT = VALMAX&
           !STATCUE = 1
           !TipoCue = "Debe"
           !POSIMPU = 1
           .Update
        End With
      Else
        Call COMUNI("Imposible Continuar !!!\Falta Ingresar Descripción de Cuenta")
        Call FINAL("")
    End If
    Call COMUNI("Alta Automática Completada.\  \Debe Volver a Pedir el Balance")
    Call FINAL("")
End Sub

