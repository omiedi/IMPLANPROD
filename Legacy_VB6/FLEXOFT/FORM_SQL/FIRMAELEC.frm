VERSION 5.00
Begin VB.Form FIRMAELEC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingresa Clave"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Height          =   720
      Left            =   2880
      Picture         =   "FIRMAELEC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   550
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
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   645
      Width           =   1575
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
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -120
      TabIndex        =   4
      Top             =   690
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -120
      TabIndex        =   3
      Top             =   285
      Width           =   1215
   End
End
Attribute VB_Name = "FIRMAELEC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ESTADO
'
Private Sub Command3_Click()
    ESTADO = 0
'    IDENT$ =

    IDENT1$ = UCase$(TRIM$(Text1))
    CLAV1$ = UCase$(TRIM$(Text2))
    If IDENT1$ = "" And CLAV1$ = "" Then GoTo 40
    
    Dim obj As New RECORXET
    Set rs1 = obj.RS_USUSQLConClave(IDENT1$, CLAV1$)
    Ident$ = SAFETEXT$(rs1!IDENTIFICADOR)
    CLAV$ = SAFETEXT$(rs1!CLAVE)
    NUS% = XVALO(rs1!NUSER)
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    
    If IDENT1$ = TRIM$(UCase$(Ident$)) Then
      If CLAV1$ = TRIM$(UCase$(CLAV$)) Then
        ESTADO = 1
        Label1 = NUS%
        GoTo 99
        
      End If
    End If
    
'---------------------------------------------------------------------------------------------
    '
40  Call MENSERR(24, "Usuario / Clave Incorrecta")
    Text1 = "*": Text2 = ""
    '
    
99  If ESTADO > 0 Then
       Hide
    Else
       Unload Me
    End If
    '
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
End Sub

Private Sub Form_Load()
  Text1 = ""
  Text2 = ""
  Label1 = ""
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ESTADO = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         Text2.SetFocus
       ElseIf KeyAscii = 27 Then
         Unload Me
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         Command3.SetFocus
       ElseIf KeyAscii = 27 Then
         Unload Me
    End If
End Sub

Sub APROENPASS()
    '
    STDSCREEN.Text99(INDEXTEX%).BackColor = RGB(255, 255, 255)
    STDSCREEN.Text99(INDEXTEX%).TEXT = ""
    '
    TANTE = Timer
    VDEVUX$ = TRIM$(UCase$(ENTRAPASS.Text1))
    If VDEVUX$ = "" Then
       On Error Resume Next
       STDSCREEN.Text99(INDEXTEX%) = ""
       STDSCREEN.Text99(INDEXTEX%).SetFocus
       On Error GoTo 0
       Text1 = ""
       Exit Sub
    End If
    '
    ORITE% = ORITEXT99%(INDEXTEX%)
    If ORITE% > 0 And ORITE% <= CANOBJ% Then
       FOREP$ = FOREPLA$(ORITE%)
       ACNAM$ = UCase$(TRIM$(ONAME$(ORITE%)))
    End If
    '
    If VDEVUX$ <> "" Then
       '
       ' VALIDACION DE LA CATEGORIA
       ID1$ = VDEVUX$
       CCAT$ = CCAT_OPER$(ID1$)
       CATI$ = CCAT$
       VALIMCAT = 0
       RESTOPA$ = TRIM$(Mid$(FOREP$, 10))
       If Left$(RESTOPA$, 1) = ">" Then
             VALIMCAT = XVALO(Mid$(RESTOPA$, 2))
             If VALIMCAT > 0 Then
                If XVALO(CCAT$) < VALIMCAT Then
                   If CATI$ <> "" Then
                     Call MENSERR(24, "Categoría Informada del Operador (" & CATI$ & ") Menor a la Requerida (" & VALIMCAT & ").")
                     Text1 = "*"
                     Exit Sub
                   End If
                End If
             End If
           ElseIf Left$(RESTOPA$, 1) = "=" Then
             RESTOPA$ = TRIM$(Mid$(RESTOPA$, 2))
             RESTOPA$ = REPLACAR$(RESTOPA$, ",", ";")
             RESTOPA$ = REPLACAR$(RESTOPA$, "/", ";")
             RESTOPA$ = ";" + RESTOPA$
             If InStr(RESTOPA$ + ";", ";" + CATI$ + ";") < 1 Then
                Call MENSERR(24, "Categoría Informada del Operador (" & CATI$ & ") No Admisible Para Firma en esta Posición (" & Mid$(RESTOPA$, 2) & ").")
                Text1 = "*"
                Exit Sub
             End If
       End If
       ' FIN DE VALIDACION DE CATEGORIA
       '
       NOPXX$ = TRIM$(NOME_OPER$(VDEVUX$))
       If NOPXX$ = "" Then
           TTXX$ = "El Identificador de Usuario '" & VDEVUX$ + "'"
           TTXX$ = TTXX$ + "\No Está Asociado a Ningún Usuario del Sistema"
           TTXX$ = TTXX$ + "\o Empleado del Padrón."
           TTXX$ = TTXX$ + "\   \Realice la Asignación Correspondiente."
           TTXX$ = TTXX$ + "\y Repita la Transaccion."
           Call MENSERR(24, TTXX$)
       End If
       '
       If TRIM$(NOPXX$) <> "" Then   ' QUE SOLO BLOQUEE SI ENCUENTRA EL EMPLEADO
          Texto(ORITE%) = VDEVUX$                 ' ESTE TIENE QUE IR ANTES DEL VALIRANGO
          If STDSCREEN.VALIRANGO%(ORITE%) < 1 Then
              Call MENSERR(24, "Usuario / Clave no Válido\para la Firma Electrónica Requerida.")
              Text1 = "*"
              Exit Sub
          End If
          '
  ' POSIBLE ERROR 15
          STDSCREEN.Text99(INDEXTEX%).TEXT = VDEVUX$            ' ESTE VA DESPUES PARA NO ACTUALIZAR NOMBRE DE USUARIO SI NO VALIDA
          On Error Resume Next
  '                If TRIM$(VDEVUX$) <> "" Then STDSCREEN.TEXT99(INDEXTEX%).Enabled = False
          On Error GoTo 0
          STDSCREEN.Text99(INDEXTEX%).BackColor = RGB(255, 255, 255)
          '
          ' aqui bloquea los campos asociados
          For KI% = 1 To STDSCREEN.Text99.Count
            ORITE1% = ORITEXT99%(KI%)
            If ORITE1% > 0 And ORITE1% <= CANOBJ% Then
              LOCKFI1$ = UCase$(TRIM$(LOCKFI$(ORITE1%)))
              If LOCKFI1$ <> "" Then
                FOREP1$ = UCase$(FOREPLA$(ORITE1%))
                If LOCKFI1$ = ACNAM$ Or InStr(FOREP1$, "{" + ACNAM$ + "}") > 0 Then
                  If Left$(UCase$(FOREP1$), 9) = "@PASSWORD" Then TREP = VDEVUX$: GoTo 90
                  If UCase$(FOREP1$) = "@USERNAME" Then TREP = USERNAME$: GoTo 90
                  If UCase$(FOREP1$) = "@AHORA" Then TREP = Time: GoTo 90
                  If UCase$(FOREP1$) = "@HOY" Then TREP = FECHATEX$(HOY(HOS$)): GoTo 90
                  '
                  If KI% <= STDSCREEN.Text99.Count Then
                    If TRIM$(STDSCREEN.Text99(KI%).TEXT) = "" Or Left$(FOREP1$, 7) = "APEYNOM" Or Left$(FOREP1$, 6) = "NOMBRE" Or Left$(FOREP1$, 6) = "APENOM" Or Left$(FOREP1$, 6) = "APYNOM" Then
                      If InStr(FOREP1$, "/") < 1 Then
                        TREP = TEXREPLA$(FOREP1$)
90                              TREPP = FORMATNUM(TREP, FORPA$(ORITE1%))
                        If TREPP <> TMASCARA(FORPA$(ORITE1%)) Then   ' si no se ha ingresado valor, mantiene vacio
                          '                                         '         y habilitado para ingresar
                          If TRIM$(CStr(TREPP)) <> "" Then                      ' para que no pise los ingresos manuales
                            Texto(ORITE1%) = TREPP
                            ' POSIBLE ERROR 16 - CAPTURADO
                            On Error Resume Next
                            STDSCREEN.Text99(KI%) = TREPP
                            On Error GoTo 0
                          End If
                        End If
                      End If
                    End If
                    On Error Resume Next
                    If TRIM$(STDSCREEN.Text99(KI%).TEXT) <> "" Then STDSCREEN.Text99(KI%).Enabled = False
                    On Error GoTo 0
                  End If
                End If
              End If
            End If
            On Error Resume Next
            STDSCREEN.Text99(KI% - 1).BackColor = RGB(255, 255, 255)
            On Error GoTo 0
          Next KI%
       End If
    End If
    '
End Sub
