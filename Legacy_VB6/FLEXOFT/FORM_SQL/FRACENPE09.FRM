VERSION 5.00
Begin VB.Form FRACENPE09 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Listados de existencias en déposito  Central de  Pesadas"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fraccionamiento de Materiales "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   60
      TabIndex        =   24
      Top             =   5865
      Width           =   10770
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9240
         TabIndex        =   37
         Top             =   345
         Width           =   1365
      End
      Begin VB.TextBox Text2 
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
         Left            =   3960
         TabIndex        =   36
         Top             =   1080
         Width           =   3975
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
         Height          =   300
         Left            =   9225
         TabIndex        =   34
         Top             =   705
         Width           =   1365
      End
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
         Height          =   390
         Left            =   8355
         Picture         =   "FRACENPE09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Graba - Aprueba la Transacción"
         Top             =   1455
         Width           =   2250
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8190
         TabIndex        =   38
         Top             =   435
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Verifique Cód. de Barras >>>>>>>>:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   35
         Top             =   1080
         Width           =   5160
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8295
         TabIndex        =   32
         Top             =   1170
         Width           =   840
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fraccionar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8175
         TabIndex        =   31
         Top             =   795
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Id.Bulto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5055
         TabIndex        =   30
         Top             =   450
         Width           =   1080
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
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   29
         Top             =   420
         Width           =   870
      End
      Begin VB.Label Label10 
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
         Height          =   300
         Left            =   300
         TabIndex        =   28
         Top             =   705
         Width           =   7635
      End
      Begin VB.Label Label9 
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
         Height          =   300
         Left            =   9225
         TabIndex        =   27
         Top             =   1095
         Width           =   1365
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
         Height          =   300
         Left            =   6240
         TabIndex        =   26
         Top             =   360
         Width           =   1710
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
         Height          =   300
         Left            =   1005
         TabIndex        =   25
         Top             =   345
         Width           =   2505
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame3"
      Height          =   1950
      Left            =   9345
      TabIndex        =   17
      Top             =   3660
      Width           =   15
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame3"
      Height          =   1980
      Left            =   6795
      TabIndex        =   16
      Top             =   3645
      Width           =   15
   End
   Begin VB.CommandButton Command1 
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
      Left            =   2850
      TabIndex        =   9
      Top             =   225
      Width           =   175
   End
   Begin VB.CommandButton Command5 
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
      Height          =   285
      Left            =   10590
      TabIndex        =   5
      Top             =   240
      Width           =   175
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Materiales Asignados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   75
      TabIndex        =   3
      Top             =   3315
      Width           =   10770
      Begin VB.Frame Frame11 
         Caption         =   "Frame3"
         Height          =   1935
         Left            =   7995
         TabIndex        =   19
         Top             =   390
         Width           =   15
      End
      Begin VB.Frame Frame10 
         Caption         =   "Frame3"
         Height          =   1965
         Left            =   2430
         TabIndex        =   18
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame3"
         Height          =   1965
         Left            =   5475
         TabIndex        =   15
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame3"
         Height          =   1980
         Left            =   570
         TabIndex        =   14
         Top             =   330
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
         Height          =   1680
         Left            =   135
         MousePointer    =   1  'Arrow
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   645
         Width           =   10410
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"FRACENPE09.frx":030A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   135
         TabIndex        =   13
         Top             =   315
         Width           =   10410
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Materiales en Deposito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   75
      TabIndex        =   0
      Top             =   930
      Width           =   10740
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1845
         Left            =   9240
         TabIndex        =   23
         Top             =   390
         Width           =   15
      End
      Begin VB.Frame Frame13 
         Caption         =   "Frame3"
         Height          =   1905
         Left            =   1965
         TabIndex        =   22
         Top             =   330
         Width           =   15
      End
      Begin VB.Frame Frame9 
         Caption         =   "Frame3"
         Height          =   1845
         Left            =   7740
         TabIndex        =   20
         Top             =   390
         Width           =   15
         Begin VB.Frame Frame12 
            Caption         =   "Frame3"
            Height          =   2200
            Left            =   -2835
            TabIndex        =   21
            Top             =   1260
            Width           =   15
         End
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
         Height          =   1680
         Left            =   105
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Top             =   560
         Width           =   10440
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"FRACENPE09.frx":03C5
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   315
         Width           =   10425
      End
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nº O/Fabric."
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
      Left            =   7200
      TabIndex        =   12
      Top             =   300
      Width           =   1320
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Déposito C. Pesadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   165
      TabIndex        =   11
      Top             =   300
      Width           =   1935
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
      Left            =   285
      TabIndex        =   10
      Top             =   555
      Width           =   2745
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Left            =   2250
      TabIndex        =   8
      Top             =   225
      Width           =   600
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
      Left            =   8580
      TabIndex        =   7
      Top             =   240
      Width           =   2010
   End
   Begin VB.Label Label30 
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
      Left            =   5415
      TabIndex        =   6
      Top             =   570
      Width           =   5370
   End
End
Attribute VB_Name = "FRACENPE09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Lista_Asignados()
   List2.Clear
   Dim Asignados As ADODB.Recordset
   Set Asignados = New ADODB.Recordset
   
   ORD$ = Label28
   MITAD1$ = Left(Label28, 10)
   MITAD2$ = XVALO(Mid(Label28, 11, 3))
   ORFA$ = MITAD1$ & MITAD2$
  
   FILA$ = Space$(120)
   
   
   SQLX$ = "select * from balareqmo  where kilo_asig > kilo_consu AND  Destino = '" & ORFA$ & "' order by idenreque"
   Set Asignados = FLEXCONN.Execute(FILTSQL$(SQLX$))
   
   If (Asignados.EOF And Asignados.BOF) Then GoTo 99
   Asignados.MoveFirst
   
   Do While Not (Asignados.EOF)
     CODX% = XVALO(Asignados!COD_INTE)
     '
     NRORD$ = SAFETEXT(Asignados!NUORITEM)
     Codigo$ = CODEXT(CODX%)
     Descripcion$ = VALOBADA("MASTER", "DESCRIPCION", "CODINT = " & CODX%)
     LOTEACT$ = SAFETEXT(Asignados!idenlote)
     '
     CANTASIG$ = SAFETEXT(Asignados!Kilo_Asig)
     If XVALO(CANTASIG$) = 0 Then
        CANTASIG$ = "0"
     End If
     CANTASIG$ = FORMATNUM$(CANTASIG$, "F12." & CNTDC$)
     '
     CantFrac$ = SAFETEXT(Asignados!Kilo_Consu)
     If XVALO(CantFrac$) = 0 Then
        CantFrac$ = "0"
     End If
     CantFrac$ = FORMATNUM$(CantFrac$, "F12." & CNTDC$)
     '
     CantSaldo$ = SAFETEXT(Asignados!KILO_SALDO)
     If XVALO(CantSaldo$) = 0 Then
        CantSaldo$ = "0"
     End If
     CantSaldo$ = FORMATNUM$(CantSaldo$, "F12." & CNTDC$)
    
     '
     FILA$ = ""
     
     Call REPLA(FILA, NRORD$, 1, 3)
     Call REPLA(FILA, Codigo$, 6, 16)
     Call REPLA(FILA, Descripcion$, 23, 27)
     Call REPLA(FILA, LOTEACT$, 52, 10)
     Call REPLA(FILA, CANTASIG$, 65, 9)
     Call REPLA(FILA, CantFrac$, 81, 9)
     Call REPLA(FILA, CantSaldo$, 93, 9)
     '
     List2.AddItem (FILA$)
     Asignados.MoveNext
   Loop
99
End Sub
'
Sub Lista_Deposito()
   List1.Clear
   Dim Deposito As ADODB.Recordset
   Set Deposito = New ADODB.Recordset
   '
   DEPO$ = Label2
   FILA$ = Space$(120)
   SQLX$ = "SELECT Cod_Inte, IdenLote,  (CanAlta - CanConsu) AS SALDO FROM LOTINGRE WHERE DepoUbic ='" & DEPO$ & "' AND SalNeto > '0'"
   Set Deposito = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   If (Deposito.EOF And Deposito.BOF) Then GoTo 99
   Deposito.MoveFirst
   '
   Do While Not (Deposito.EOF)
     CODX% = XVALO(Deposito!COD_INTE)
     Codigo$ = CODEXT(CODX%)
     Descripcion$ = VALOBADA("MASTER", "DESCRIPCION", "CODINT = " & CODX%)
     LOTEACT$ = SAFETEXT(Deposito!idenlote)
     SALDEP$ = SAFETEXT(Deposito!SALDO): SALDEP$ = FORMATNUM$(SALDEP$, "F12." & CNTDC$)
     '
     Call REPLA(FILA, Codigo$, 1, 16)
     Call REPLA(FILA, Descripcion$, 19, 57)
     Call REPLA(FILA, LOTEACT$, 74, 15)
     Call REPLA(FILA, SALDEP$, 90, 9)
     '
     List1.AddItem (FILA$)
     Deposito.MoveNext
   Loop
99
End Sub

Private Sub Command1_Click()
    Call SELECHO("TADEPOSI")
    If VALACT1$("TADEPOSI") <> "" Then
       Label2 = VALACT1$("TADEPOSI")
       Label3 = VALACT2$("TADEPOSI")
    End If
    
    If Label2 <> "" And Label28 <> "" Then
      Lista_Asignados
      Lista_Deposito
    End If
End Sub

    
Private Sub Command5_Click()
   Call SELECHO("INDIREQU")
    
   If TRIM$(VALACT1$("INDIREQU")) <> "" Then
      Label28 = TRIM$(VALACT1$("INDIREQU"))
      Label30 = TRIM$(VALACT2$("INDIREQU"))
   End If
   If Label2 <> "" And Label28 <> "" Then
      Lista_Asignados ' Completa List 2
      Lista_Deposito 'Completa List 1
   End If

End Sub


 

Private Sub Label2_Change()
   Lista_Deposito
End Sub



Private Sub List1_Click()
   List1.ListIndex = -1
End Sub

Private Sub List2_DblClick()
    '
    FILA$ = List2.TEXT
    Text2 = Mid$(FILA$, 6, 16) + Mid$(FILA$, 52, 12)
    
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    '
    ' LIMPIA LOS CAMPOS
    Label7 = ""
    Label6 = ""
    Label10 = ""
    Label9 = ""
    Text1 = ""
    '
    If KeyAscii > 13 Then
      Exit Sub
    End If
    '
    TENTRA$ = TRIM$(Text2.TEXT)
    For U& = 1 To Len(TENTRA$)
      If Asc(Mid$(TENTRA$, U&, 1)) = 39 Then
        Mid$(TENTRA$, U&, 1) = Chr$(45)
      End If
    Next U&
    '
    CODE$ = TRIM$(Left$(TENTRA$, 16))
    LOTE$ = TRIM$(Mid$(TENTRA$, 17, 12))
    '
    'BUSCA CODIGO Y LOTE EN LIST2 (ASIGNADOS)
    CODENCOT = False
    For I = 0 To List2.ListCount - 1
       CODD$ = TRIM$(Mid$(List2.List(I), 6, 16))
       If CODD$ = CODE$ Then
          CODENCOT = True
          LOTT$ = TRIM$(Mid$(List2.List(I), 52, 12))
          If LOTT$ = LOTE$ Then
             ' LO ENCONTRO
             LOTEENCOT = True
             CIXI% = CODINT%(CODE$)
             Label7 = CODE$
             Label6 = LOTE$
             Label10 = DESCRIT$(CIXI%)
             Text1 = Mid$(List2.List(I), 65, 9)
             GoTo 30
         End If
       End If
    Next I
    '
    ' MENSAJE DE ERROR NO LO ENCONTRO
    If CODENCOT = False Then ' Si no encontro el codigo muestra mensaje
       COMUNI ("CODIGO " & CODE$ & " NO SE ENCUENTRA EN DEPOSITO")
     Else ' de lo contrario no se encontro el lote
       COMUNI ("LOTE NO CORRESPONDE CON EL ASIGNADO EN DEPOSITO")
    End If
    Exit Sub
    '
30  'BUSCA LOTE EN DEPOSITO (LIST1) EL NUMERO DEL LOTE
    For I = 0 To List1.ListCount - 1
       LOTEDEPO$ = Mid$(List1.List(I), 74, 15)
       If TRIM$(LOTEDEPO$) = TRIM$(LOTE$) Then
          List1.ListIndex = I
          CANTDEPO$ = Mid$(List1.List(I), 90, 9)
          Text3 = FORMATNUM$(CANTDEPO$, "F13." & CNTDC$)   'DISPONIBLE
       End If
    Next I
    
    'CALCULA EL SALDO
    If (XVALO(Text3) - XVALO(Text1)) < 0 Then
       COMUNI ("Cantidad a Fraccionar No Puede ser Mayor que la Cantidad Disponible")
       Exit Sub
    End If
    '
    SAL# = XVALO(Text3) - XVALO(Text1)
    If SAL# = 0 Then
       Label9 = "0.000"
     Else
       Label9 = FORMATNUM(SAL#, "F13." & CNTDC$)         'Saldo
    End If

End Sub

