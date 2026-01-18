VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FORDERAC 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Derechos de Acceso de Usuarios"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10830
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
   ScaleHeight     =   6030
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   4365
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   7699
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   16
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      ItemData        =   "FORDERAC.frx":0000
      Left            =   240
      List            =   "FORDERAC.frx":0002
      TabIndex        =   15
      Top             =   6960
      Visible         =   0   'False
      Width           =   10545
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Nuevo Derecho"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6720
      TabIndex        =   14
      Top             =   1080
      Width           =   2025
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
      Left            =   8880
      Picture         =   "FORDERAC.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   195
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   6720
      TabIndex        =   11
      Top             =   90
      Width           =   2025
      Begin VB.TextBox Text1 
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
         Height          =   360
         Left            =   195
         TabIndex        =   0
         ToolTipText     =   "<ENTER> para Búsqueda de Texto - F3 para Seguir Buscando + Coincidencias"
         Top             =   330
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3795
      Left            =   210
      ScaleHeight     =   3735
      ScaleWidth      =   10395
      TabIndex        =   8
      Top             =   1680
      Width           =   10455
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3435
         ItemData        =   "FORDERAC.frx":066E
         Left            =   -10
         List            =   "FORDERAC.frx":0670
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   10545
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   10580
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Aplicación                                                                     R    W      Módulo"
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
            Left            =   105
            TabIndex        =   10
            Top             =   60
            Width           =   7470
         End
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   0
         X2              =   6300
         Y1              =   3990
         Y2              =   3990
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Usuario FLEXOFT Seleccionado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1310
      Left            =   90
      TabIndex        =   4
      Top             =   120
      Width           =   6525
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   2205
         TabIndex        =   7
         Top             =   510
         Width           =   4080
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   1470
         TabIndex        =   6
         Top             =   510
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Número de Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   210
         TabIndex        =   5
         Top             =   510
         Width           =   1065
      End
   End
   Begin VB.CommandButton Command7 
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
      Left            =   8880
      Picture         =   "FORDERAC.frx":0672
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   810
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
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
      Left            =   9840
      Picture         =   "FORDERAC.frx":097C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   810
      Width           =   855
   End
   Begin VB.CommandButton command1 
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
      Height          =   600
      Left            =   9840
      Picture         =   "FORDERAC.frx":0C86
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   195
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "FORDERAC.frx":0DD0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "FORDERAC.frx":0E4A
      Top             =   1200
   End
   Begin ACTIVESKINLibCtl.Skin Skin2 
      Left            =   3600
      OleObjectBlob   =   "FORDERAC.frx":107E
      Top             =   1200
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   840
   End
End
Attribute VB_Name = "FORDERAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VACON%
Dim ARBUS$, ABUSQUE$

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command2_Click()
   '
   Screen.MousePointer = 11
   USEL% = Val(Label2.Caption)
   V& = 0
   Call BLOQARCH("FLEXACC")
   '
'   VACON% = 1
'   FIN& = ULTREG&("FLEXACC")
'   For U& = 1 To ULTREG&("FLEXACC")
'      Call TRACE(20, U&, FIN&)
'      X$ = REGLEIDO$("FLEXACC", U&)
'      If CVI(Left$(X$, 2)) <> USEL% Then
'          V& = V& + 1
'          Call GRAREG("FLEXACC", X$, V&)
'      End If
'   Next U&
   '
   'A PARTIR DE ACA CODIGO NUEVO
   FIN& = MSF.Rows - 1
   Dim RST As ADODB.Recordset
   Set RST = New ADODB.Recordset

   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
'      If Left(MSF.TextMatrix(U&, 2), 1) = "S" Or Left(MSF.TextMatrix(U&, 3), 1) = "S" Then
         CAPLI$ = TRIM$(MSF.TextMatrix(U&, 4))
         DRX1% = 0: If Left(MSF.TextMatrix(U&, 2), 1) = "S" Then DRX1% = 1
         DRX2% = 0: If Left(MSF.TextMatrix(U&, 3), 1) = "S" Then DRX2% = 1
         SQLX$ = "SELECT * FROM DER_ACCE "
         SQLX$ = SQLX$ + " WHERE NUM_USER = " & USEL%
         SQLX$ = SQLX$ + " AND COD_APLI = '" & CAPLI$ & "'"
         RST.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
         
'         Set RST = BaDaConsul.OpenRecordset(SQLX$, dbOpenDynaset)
         With RST
          If .EOF Then
           If DRX1% < 1 And DRX2% < 1 Then GoTo 19
           .AddNew
          End If
         !NUM_USER = USEL%
         !COD_APLI = CAPLI$
         !DER_READ = DRX1%
         !DER_WRITE = DRX2%
         .Update
      End With
19    RST.Close
   Next U&

   On Error GoTo ERROR_MDB
   Call ABRETACON
   For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
      'If Left(MSF.TextMatrix(U&, 2), 1) = "S" Or Left(MSF.TextMatrix(U&, 3), 1) = "S" Then
         CAPLI$ = TRIM$(MSF.TextMatrix(U&, 4))
         DRX1% = 0: If Left(MSF.TextMatrix(U&, 2), 1) = "S" Then DRX1% = 1
         DRX2% = 0: If Left(MSF.TextMatrix(U&, 3), 1) = "S" Then DRX2% = 1
         SQLX$ = "SELECT * FROM DER_ACCE "
         SQLX$ = SQLX$ + " WHERE NUM_USER = " & USEL%
         SQLX$ = SQLX$ + " AND COD_APLI = '" & CAPLI$ & "'"
         Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenDynaset)
         With PARATEMP
          If .EOF Then
           If DRX1% < 1 And DRX2% < 1 Then GoTo 20
           .AddNew
          Else
           .Edit
          End If
         !NUM_USER = USEL%
         !COD_APLI = CAPLI$
         !DER_READ = DRX1%
         !DER_WRITE = DRX2%
         .Update
      End With
20    PARATEMP.Close
   Next U&
   On Error GoTo 0
   Call LIBARCH("FLEXACC")

   GoTo 90
   
   
   
   'FIN CODIGO NUEVO
'   FIN& = List1.ListCount
'   Call ABRETACON
'   For U& = 1 To List1.ListCount
'      Call TRACE(20, U&, FIN&)
'      Z$ = List1.List(U& - 1)
'      If Mid$(Z$, 45, 1) = "S" Or Mid$(Z$, 49, 1) = "S" Then
'         CAPLI$ = TRIM$(Mid$(Z$, 53, 32))
'         If Len(CAPLI$) <= 8 Then
'            Z1$ = MKI$(0) + Space$(46)
'            Call REPLA(Z1$, MKI$(USEL%), 1, 2)
'            Call REPLA$(Z1$, Mid$(Z$, 53, 8), 3, 8)
'            Call REPLA(Z1$, Mid$(Z$, 45, 1), 11, 1)
'            Call REPLA(Z1$, Mid$(Z$, 49, 1), 12, 1)
'            V& = V& + 1
'            Call GRAREG("FLEXACC", Z1$, V&)
'         End If
'      End If
'      '
'      CAPLI$ = TRIM$(Mid$(Z$, 53, 32))
'      DRX1% = 0: If Mid$(Z$, 45, 1) = "S" Then DRX1% = 1
'      DRX2% = 0: If Mid$(Z$, 49, 1) = "S" Then DRX2% = 1
'      SQLX$ = "SELECT * FROM DER_ACCE "
'      SQLX$ = SQLX$ + " WHERE NUM_USER = " & USEL%
'      SQLX$ = SQLX$ + " AND COD_APLI = '" & CAPLI$ & "'"
'      Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenDynaset)
'      With PARATEMP
'        On Error Resume Next
'        .MoveFirst
'        If Err Then
'           On Error GoTo 0
'           If DRX1% < 1 And DRX2% < 1 Then GoTo 19
'           .AddNew
'         Else
'           On Error GoTo 0
'           .Edit
'        End If
'        !NUM_USER = USEL%
'        !COD_APLI = CAPLI$
'        !DER_READ = DRX1%
'        !DER_WRITE = DRX2%
'        .Update
'      End With
'19    PARATEMP.Close
'   Next U&
'   '
'   Call SETULTREG("FLEXACC", V&)
'   Call CIERRARCH("FLEXACC")
   Call LIBARCH("FLEXACC")
'   VACON% = 0
   '
90 Screen.MousePointer = 1
   Hide
   Exit Sub
ERROR_MDB:
  Call COMUNI("Se ha Producido Un Error En Tabla SQLCTRL-DERACCE\Consulte a su Soporte de Sistema")
  Screen.MousePointer = 1: Hide

   '
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   OPPX% = COMALTER%("Seleccione el Formato del Listado:\\    Completo    \Solo Derechos Asignados")
   Call HEADERS("DERUSER", "")
   Call HEADERS("DERUSER", "Usuario: " + TRIM$(Label4) + " (" + TRIM$(Label2) + ")")
   '
   JJ& = 0
   For UI% = 1 To MSF.Rows - 1
        If OPPX% = 1 Then
           ZZZ$ = Space$(64)
           Call REPLA(ZZZ$, TRIM$(MSF.TextMatrix(UI%, 4)), 1, 12)
           Call REPLA(ZZZ$, TRIM$(MSF.TextMatrix(UI%, 1)), 13, 48)
           Call REPLA(ZZZ$, TRIM$(MSF.TextMatrix(UI%, 2)), 61, 2)
           Call REPLA(ZZZ$, TRIM$(MSF.TextMatrix(UI%, 3)), 63, 2)
           JJ& = JJ& + 1
           Call GRAREG("DERUSER", ZZZ$, JJ&)
        ElseIf OPPX% = 2 Then
           If TRIM$(MSF.TextMatrix(UI%, 2)) <> "" Or TRIM$(MSF.TextMatrix(UI%, 3)) <> "" Then
                ZZZ$ = Space$(64)
                Call REPLA(ZZZ$, TRIM$(MSF.TextMatrix(UI%, 4)), 1, 12)
                Call REPLA(ZZZ$, TRIM$(MSF.TextMatrix(UI%, 1)), 13, 48)
                Call REPLA(ZZZ$, TRIM$(MSF.TextMatrix(UI%, 2)), 61, 2)
                Call REPLA(ZZZ$, TRIM$(MSF.TextMatrix(UI%, 3)), 63, 2)
                JJ& = JJ& + 1
                Call GRAREG("DERUSER", ZZZ$, JJ&)
           End If
        Else
            Exit Sub
        End If
   Next UI%
  
   
   
'   JJ& = 0
'   For UI% = 1 To List1.ListCount
'     XXX$ = List1.List(UI% - 1)
'        If OPPX% = 1 Then
'           ZZZ$ = Space$(64)
'           Call REPLA(ZZZ$, Mid$(XXX$, 53, 12), 1, 12)
'           Call REPLA(ZZZ$, Mid$(XXX$, 1, 36), 13, 48)
'           Call REPLA(ZZZ$, Mid$(XXX$, 45, 2), 61, 2)
'           Call REPLA(ZZZ$, Mid$(XXX$, 49, 2), 63, 2)
'           JJ& = JJ& + 1
'           Call GRAREG("DERUSER", ZZZ$, JJ&)
'        ElseIf OPPX% = 2 Then
'           If TRIM$(Mid$(XXX$, 45, 6)) <> "" Then
'              ZZZ$ = Space$(64)
'              Call REPLA(ZZZ$, Mid$(XXX$, 53, 12), 1, 12)
'              Call REPLA(ZZZ$, Mid$(XXX$, 1, 36), 13, 48)
'              Call REPLA(ZZZ$, Mid$(XXX$, 45, 2), 61, 2)
'              Call REPLA(ZZZ$, Mid$(XXX$, 49, 2), 63, 2)
'              JJ& = JJ& + 1
'              Call GRAREG("DERUSER", ZZZ$, JJ&)
'           End If
'        Else
'            Exit Sub
'        End If
'   Next UI%
   Call SETULTREG("DERUSER", JJ&)
   Call CIERRARCH("DERUSER")
   '
   Call CONECRUN("*DERUSER", "Listado Derechos de Acceso por Usuario")
   '
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
10 OBX$ = OBJPLA$("NUEVDERE", VDEF$)
   If OBX$ = "" Then
        GoTo 99
   End If
   
   CODI$ = UCase$(TRIM$(Mid$(OBX$, 1, 16)))
   DESCR$ = UCase$(TRIM$(Mid$(OBX$, 17, 24)))
   '
   'VALIDO
      If TRIM$(CODI$) = "" Then
     Call COMUNI("Falta Codigo de Derecho")
     VDEF$ = OBX$: GoTo 10
   End If
   '
   If TRIM$(DESCR$) = "" Then
     Call COMUNI("Falta Descripción de Derecho")
     VDEF$ = OBX$: GoTo 10
   End If
   '
   'PASO LOS PARAMETROS A LA FUNCION DERACCE
   If DERACCE%(CODI$, DESCR$) > 0 Then
   End If
   '
   'ACTUALIZO EL LIST
   USEL% = XVALO(Label2)
   Call ACLISDER(USEL%, CODI$)
   
   Text1 = CODI$
   Text1.SetFocus
   
   
99 Command4.Enabled = True

End Sub


Private Sub Command7_Click()
    Call HELPONLINE("DERACCE")
End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      ' &H8000000F
     End If

    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    Call CARGA_ENCABEZADO(Me.MSF, "Aplicación/A48;Lect./A4;Escrit./A4;Módulo/A28", Me)
    
 
End Sub

Private Sub List1_Click()
   If VACON% = 1 Then
       Exit Sub
   End If
   ILISTA& = List1.ListIndex
   Z$ = List1.TEXT
   If Mid$(Z$, 49, 2) = "Si" Then
        Call REPLA(Z$, "       ", 45, 7)
      ElseIf Mid$(Z$, 45, 2) = "Si" Then
        Call REPLA(Z$, "Si", 49, 2)
      Else
        Call REPLA(Z$, "Si     ", 45, 7)
   End If
   List1.RemoveItem ILISTA&
   List1.AddItem Z$, ILISTA&
   VACON% = 1
   List1.ListIndex = ILISTA&
   VACON% = 0
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      CODAP$ = TRIM$(Mid$(List1.TEXT, 53, 32))
      CODAP12$ = TRIM$(Left$(CODAP$, 12))
      JJ& = 0
      For U& = 1 To ULTREG&("INDIAPLI")
         XX$ = REGLEIDO$("INDIAPLI", U&)
         If TRIM$(Left$(XX$, 12)) <> CODAP12$ Then
           JJ& = JJ& + 1
           Call GRAREG("INDIAPLI", XX$, JJ&)
         End If
      Next U&
      Call SETULTREG("INDIAPLI", JJ&)
      '
      Call ABRETACON
      SQLX$ = "DELETE FROM INDI_APLI WHERE COD_APLI = '" & CODAP$ & "'"
      On Error Resume Next
      BaDaConsul.Execute SQLX$
      If Err Then
          On Error GoTo 0
          Exit Sub
        Else
          ILISTA& = List1.ListIndex
          List1.RemoveItem ILISTA&
      End If
      On Error GoTo 0
   End If
End Sub

Private Sub MSF_DblClick()
    i& = MSF.MouseRow
    J& = MSF.MouseCol
    If J& = 2 Or J& = 3 Then
       If MSF.TextMatrix(i&, J&) = "Si" Then
          MSF.TextMatrix(i&, J&) = ""
       Else
          MSF.TextMatrix(i&, J&) = "Si"
       End If
    End If
  
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii% <> 13 Then Exit Sub
   VACON% = 1
   ABUSQUE$ = UCase$(TRIM$(Text1.TEXT))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     For U& = 1 To MSF.Rows - 1
      For i& = 0 To MSF.Cols - 1
       If InStr(UCase$(MSF.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
        MSF.Row = U&: MSF.COL = i&
        If MSF.TopRow > MSF.Row Then MSF.TopRow = MSF.Row
        On Error Resume Next
        VISIROWS = (MSF.Height / MSF.RowHeight(J&)) - 4
        If Err > 0 Then
         VISIROWS = 1
        End If
        On Error GoTo 0
         If VISIROWS > 0 Then
            If MSF.TopRow < MSF.Row - VISIROWS Then
              MSF.TopRow = MSF.Row - VISIROWS
            End If
         End If
        Call SELECCIONARFILA(Me.MSF, MSF.Row)
        MSF.SetFocus
        Exit Sub
       End If
       Next i&
     Next U&
     '
     Call MENSERR(24, "Argumento no Encontrado.")
   End If
'   If KeyAscii% = 13 Then
'     VACON% = 1
'     INIBUS& = List1.ListIndex + 1
'     TEBUS$ = TRIM$(UCase$(Text1.TEXT))
'     For KU& = INIBUS& + 1 To List1.ListCount
'       If InStr(UCase$(List1.List(KU& - 1)), TEBUS$) > 0 Then
'         List1.ListIndex = KU& - 1
'         VACON% = 0
'         Exit Sub
'       End If
'     Next KU&
'     For KU& = 1 To INIBUS&
'       If InStr(UCase$(List1.List(KU& - 1)), TEBUS$) > 0 Then
'         List1.ListIndex = KU& - 1
'         VACON% = 0
'         Exit Sub
'       End If
'     Next KU&
'     VACON% = 0
'   End If
End Sub

Private Sub MSF_KeyDown(KeyCode As Integer, Shift As Integer)
   REG& = MSF.MouseRow
   COL& = MSF.MouseCol

   If KeyCode = 114 Then ' BUSQUEDA  POR TEXTO CON F3
       If ARBUS$ <> "" Then
          VUELTA% = 0
          J& = MSF.Row + 1

          h& = 0: If MSF.COL < MSF.Cols Then h& = MSF.COL + 1
10        For U& = J& To MSF.Rows - 1
             HI& = h&: If U& > J& Then HI& = 0
             For i& = HI& To MSF.Cols - 1
              If InStr(UCase$(MSF.TextMatrix(U&, i&)), ABUSQUE$) > 0 Then
                MSF.Row = U&: MSF.COL = i&
                On Error Resume Next
                 VISIROWS = (MSF.Height / MSF.RowHeight(J&)) - 4
                If Err > 0 Then
                 VISIROWS = 1
                End If
                On Error GoTo 0
                 If VISIROWS > 0 Then
                    If MSF.TopRow < MSF.Row - VISIROWS Then
                      MSF.TopRow = MSF.Row - VISIROWS
                    End If
                 End If
                MSF.SetFocus
                Call SELECCIONARFILA(Me.MSF, MSF.Row)
                Exit Sub
              End If
              Next i&
          Next U&
          '
          Call MENSERR(24, "Argumento no Encontrado.")
       End If
   End If
   
   
   
   If KeyCode = 46 Then
      User% = USNBR% Mod 100
      If User% <> 1 Then Exit Sub
      OPX% = COMALTER%("El Derecho Selecccionado Se Borrará De Manera Irreversible\\Cancelar\Borrar")
      If OPX% < 2 Then Exit Sub

      CODAP$ = MSF.TextMatrix(REG&, 4)
      JJ& = 0
      '
      'BORRA EN SQL
      SQLX$ = "DELETE FROM INDI_APLI WHERE COD_APLI = '" & CODAP$ & "'"
      On Error Resume Next
      FLEXCONN.Execute SQLX$
      
      Call cmdEliminaItem(Me.MSF, REG&)
      On Error GoTo 0
      On Error GoTo ERROR_MDB
      Call ABRETACON
      SQLX$ = "DELETE FROM INDI_APLI WHERE COD_APLI = '" & CODAP$ & "'"
      On Error Resume Next
      BaDaConsul.Execute SQLX$
      If Err Then
          On Error GoTo 0
          Exit Sub
      End If
      On Error GoTo 0
   End If
   Exit Sub
   
ERROR_MDB:
    Call COMUNI("No se Ha Podido Guardar el Derecho en Tabla SQLCTRL.MDB\Consulte a su Soporte de Sistema")
    On Error GoTo 0
End Sub

