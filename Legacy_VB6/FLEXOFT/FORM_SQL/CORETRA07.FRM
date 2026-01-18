VERSION 5.00
Begin VB.Form CORETRA07 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Trazabilidad de Componentes"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   2880
      TabIndex        =   24
      Top             =   180
      Width           =   165
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Equipos que Comparten Mismo Componente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   19
      Top             =   4200
      Width           =   8760
      Begin VB.Frame Frame5 
         Caption         =   "                "
         Height          =   2520
         Left            =   2040
         TabIndex        =   23
         Top             =   480
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "                                                                                                                  "
         Height          =   2520
         Left            =   6480
         TabIndex        =   22
         Top             =   480
         Width           =   15
      End
      Begin VB.ListBox List2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         ItemData        =   "CORETRA07.frx":0000
         Left            =   240
         List            =   "CORETRA07.frx":0002
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   600
         Width           =   8265
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Código                    Descripción                                                         Nº Serie"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   8265
      End
   End
   Begin VB.TextBox Text1 
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
      Left            =   840
      TabIndex        =   18
      Top             =   180
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Consumos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   135
      TabIndex        =   5
      Top             =   1410
      Width           =   8730
      Begin VB.Frame Frame1 
         Caption         =   "                                                                                                                  "
         Height          =   2280
         Left            =   6600
         TabIndex        =   14
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame14 
         Caption         =   "                "
         Height          =   2280
         Left            =   2000
         TabIndex        =   9
         Top             =   360
         Width           =   15
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         ItemData        =   "CORETRA07.frx":0004
         Left            =   240
         List            =   "CORETRA07.frx":0006
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   600
         Width           =   8265
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Código                    Descripción                                                             Lote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   8265
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00000040&
      Height          =   8400
      Left            =   9000
      ScaleHeight     =   8340
      ScaleWidth      =   960
      TabIndex        =   1
      Top             =   0
      Width           =   1020
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
         Height          =   690
         Left            =   120
         Picture         =   "CORETRA07.frx":0008
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Trazabilidad Inversa"
         Top             =   3240
         Width           =   690
      End
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
         Height          =   690
         Left            =   120
         Picture         =   "CORETRA07.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Acceso a Maestro de Items de Stock"
         Top             =   1800
         Width           =   690
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Kdx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "CORETRA07.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   2520
         Width           =   690
      End
      Begin VB.CommandButton Command11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "CORETRA07.frx":0766
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   225
         Width           =   690
      End
      Begin VB.CommandButton Command10 
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
         Height          =   690
         Left            =   120
         Picture         =   "CORETRA07.frx":08B0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   960
         Width           =   690
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   3
         Top             =   6480
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -240
         Picture         =   "CORETRA07.frx":0BBA
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   6840
         Width           =   1515
      End
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4080
      TabIndex        =   31
      Top             =   975
      Width           =   465
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fact"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6440
      TabIndex        =   30
      Top             =   960
      Width           =   465
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   7000
      TabIndex        =   29
      Top             =   960
      Width           =   1890
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   4560
      TabIndex        =   28
      Top             =   960
      Width           =   1875
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   840
      TabIndex        =   27
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -840
      TabIndex        =   26
      Top             =   975
      Width           =   1545
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   1560
      TabIndex        =   25
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "N/Serie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -760
      TabIndex        =   17
      Top             =   225
      Width           =   1545
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   3000
      TabIndex        =   13
      Top             =   585
      Width           =   5895
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -840
      TabIndex        =   12
      Top             =   600
      Width           =   1545
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   840
      TabIndex        =   11
      Top             =   585
      Width           =   2055
   End
   Begin VB.Label Label28 
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   7215
      TabIndex        =   8
      Top             =   180
      Width           =   1665
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Número de O/F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5520
      TabIndex        =   7
      Top             =   225
      Width           =   1545
   End
End
Attribute VB_Name = "CORETRA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   Call SELECHO("INDINUSE")
   TTXX$ = VALACT1$("INDINUSE")
   If TTXX$ <> "" Then Text1.TEXT = TTXX$
   '
End Sub

Private Sub Command11_Click()
    Hide
End Sub

Private Sub Command2_Click()
   '
   Command2.Enabled = False
   TTXX$ = List2.List(List2.ListIndex)
   CODILLA$ = TRIM$(Mid$(TTXX$, 1, 16))
   If CODILLA$ = "" Then CODILLA$ = TRIM$(Label11)
   FIKARDEX07.Text1 = CODILLA$
   FIKARDEX07.Show 1
   Command2.Enabled = True
   '
End Sub

Private Sub Command3_Click()
   TRAZINVER09.Show 1
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    If DERACCE%("AMASTO07", "Maestro de Items de Stock") > 1 Then
       AMASTOK07.Text1 = Label11
       AMASTOK07.Show 1
    End If
    Command4.Enabled = True
End Sub

Private Sub Form_Load()
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    Call ABRECONEXION
    '
End Sub





Private Sub Label28_Change()
   Call LIMPIAR
   NORFA$ = Label28
   Ci% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", "IdSubOr = '" & NORFA$ & "'"))
   If Ci% > 0 Then
     Label11 = CODEXT(Ci%)
     Label13 = DESCRIT$(Ci%)
     Call CARGACLIEN ' CARGA LOS DATOS DEL CLIENTE , REMITO Y FACTURA SEGUN EL NUMERO DE SERIE DEL ARTICULO
     CONDI$ = " DopriCor = '" & NORFA$ & "' AND CodiMovi = 'CF'"
     Call ENUMERA("MOVISTO", "OrdConsum", CONDI$)
     Call CARGALISTA(List1, "MOVISTO", "Cod_Exte/A16; Descrip/A43; IdenLote/A16; OrdConsum/F12", CONDI$)
     '
10   ' ELIMINA LOS ITEMS NO TRAZABLES
     Screen.MousePointer = 11
     i& = 1
     Do While i& <= List1.ListCount
        CIXI% = CODINT%(TRIM$(Left$(List1.List(i& - 1), 16)))
        If ESTRAZABLE%(CIXI%) < 1 Then
           List1.RemoveItem (i& - 1)
         Else
           i& = i& + 1
        End If
     Loop
     Screen.MousePointer = 1
     '
     'RECOREEMOS PARA QUE NO QUEDEN REPETIDOS EN EL LIST
     CIXIA% = 0: LOTA$ = "":  i& = 1
     Do While i& <= List1.ListCount
       LOT$ = TRIM$(Mid$(List1.List(i& - 1), 59, 16))
       CIXI% = CODINT%(TRIM$(Left$(List1.List(i& - 1), 16)))
       If CIXI% <> CIXIA% And (LOT$ = "" Or LOT$ <> LOTA$) Then
           i& = i& + 1
           CIXIA% = CIXI%
           LOTA$ = LOT$
         Else
           List1.RemoveItem (i& - 1)
       End If
     Loop
     '
   End If

End Sub

Sub LIMPIAR()

     Label11 = ""
     Label13 = ""
     Label5 = ""
     Label3 = ""
     Label6 = ""
     Label7 = ""
     List1.Clear
     List2.Clear
    
End Sub


Private Sub List1_DblClick()
   '
   List2.Clear
   '
   Call APERBASDAT("STOCKS")
   TTXX$ = List1.List(List1.ListIndex)
   LOTE$ = TRIM$(Mid$(TTXX$, 62, 16))
   If LOTE$ = "" Then
      Call MENSERR(24, "Componente no Trazado")
      Exit Sub
   End If
   SQLX$ = " SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + "WHERE IdenLote = '" & LOTE$ & "' "
   SQLX$ = SQLX$ + "AND CODIMOVI = 'CF'"
   '
   Dim LOTETEMP As ADODB.Recordset
   Set LOTETEMP = New ADODB.Recordset
   LOTETEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockPessimistic, adCmdText
   '
   With LOTETEMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       Do While Not .EOF
       NUORFA$ = TRIM$(SAFETEXT$(!DOPRILAR))
       SQLX$ = "SELECT * FROM TANUMSE "
       SQLX$ = SQLX$ + "WHERE IDSUBOR = '" & NUORFA$ & "'"
       '
       Dim ORFATEMP As ADODB.Recordset
       Set ORFATEMP = New ADODB.Recordset
       ORFATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockPessimistic, adCmdText
       '
       With ORFATEMP
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
           Do While Not .EOF
              CIXI% = XVALO(!COD_INTE)
              NUMSER$ = SAFETEXT(!NUMEROSERIE)
              ' VERIFICA QUE EL NUMERO DE SERIE NO ESTE PREVIAMENTE CARGADO
              For i = 1 To List2.ListCount
                NSER$ = Mid$(List2.List(i - 1), 60, 20)
                If TRIM$(NUMSER$) = TRIM$(NSER$) Then
                  GoTo 30
                End If
              Next i
              '
              TXTX$ = Space(128)
              Call REPLA(TXTX$, CODEXT$(CIXI%), 1, 16)
              Call REPLA(TXTX$, DESCRIT$(CIXI%), 19, 40)
              Call REPLA(TXTX$, NUMSER$, 60, 20)
              List2.AddItem (TXTX$)
30            .MoveNext
           Loop
         End If
       End With
       .MoveNext
       Loop
     End If
   End With
   ORFATEMP.Close
   Set ORFATEMP = Nothing
   LOTETEMP.Close
   Set LOTETEMP = Nothing
    
End Sub

Private Sub List2_DblClick()
   TTXX$ = List2.List(List2.ListIndex)
   List1.Clear
   List2.Clear
   Text1.TEXT = TRIM$(Mid$(TTXX$, 60, 20))
End Sub

Private Sub Text1_Change()
    '
    OF$ = VALOBADA("TANUMSE", "IDSUBOR", "NUMEROSERIE= '" & TRIM$(Text1.TEXT) & "'")
    If TRIM$(OF$) = "" Then
       Label28 = ""
       Call LIMPIAR
       Exit Sub
    End If
    '
    Label28.Caption = OF$
    Call Label28_Change
    'Si no hay consumos Traz. muestra mensaje
    If List1.ListCount = 0 Then
       Call COMUNI("No Hay Consumos Trazables Para Este Equipo")
    End If
    '
    If Label28.Caption = "" Then
       CIXI% = XVALO(VALOBADA("TANUMSE", "cod_Inte", "NUMEROSERIE= '" & TRIM$(Text1.TEXT) & "'"))
       If CIXI > 0 Then
         Label11.Caption = CODEXT$(CIXI%)
         Label13.Caption = DESCRIT$(CIXI%)
         Call CARGACLIEN ' CARGA LOS DATOS DEL CLIENTE, FACTURA Y REMITO SEGUN NUMERO DE SERIE
         Call COMUNI("Número de Serie Probablemente Cargado en Forma Manual")
       End If
    End If
    '
End Sub

Private Sub Text1_DblClick()
   '
   Call Command1_Click
   '
End Sub

Sub CARGACLIEN()
    '
    CONDI$ = "NUMEROSERIE = '" & TRIM$(Text1.TEXT) & "'"
    NC = XVALO(VALOBADA("TANUMSE", "NUMECLI", CONDI$))
    DEPO% = XVALO(VALOBADA("TANUMSE", "DEPONUMSE", CONDI$))
    If DEPO% = -1 Then
       Label5.Caption = ""
       Label3.Caption = "En Producción"
    ElseIf NC > 0 Then
       Label5.Caption = NC
       Label3.Caption = TRIM$(rasocli$(NC))
    End If
    Label6.Caption = SAFETEXT(VALOBADA("TANUMSE", "NUMREMITO", "NUMEROSERIE = '" & TRIM$(Text1.TEXT) & "'"))
    Label7.Caption = SAFETEXT(VALOBADA("TANUMSE", "NUMFACTURA", "NUMEROSERIE = '" & TRIM$(Text1.TEXT) & "'"))
    If TRIM(Label6) = "" And DEPO% > 0 Then
       DESDEPO$ = TRIM(ECOARCH("TADEPOSI", SAFETEXT(DEPO%)))
       Label6.Caption = DEPO% & "  (" & DESDEPO$ & ")"
    End If
    '
End Sub
