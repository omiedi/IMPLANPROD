VERSION 5.00
Begin VB.Form CANJECOMP 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7995
   ClientLeft      =   165
   ClientTop       =   5535
   ClientWidth     =   9825
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton command1 
      Caption         =   "O.K."
      Height          =   810
      Left            =   9030
      Picture         =   "CANJECOMP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7100
      Width           =   750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   105
      TabIndex        =   12
      Top             =   5040
      Width           =   8700
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
         Height          =   2040
         ItemData        =   "CANJECOMP.frx":030A
         Left            =   165
         List            =   "CANJECOMP.frx":030C
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   630
         Width           =   8370
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Código                    Descripción                                                     Cant.  "
         Height          =   270
         Left            =   165
         TabIndex        =   17
         Top             =   315
         Width           =   8370
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Configuración de Tipo de Canje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   8700
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   945
         TabIndex        =   19
         Top             =   315
         Width           =   2010
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "N/Serie:"
         Height          =   225
         Left            =   105
         TabIndex        =   18
         Top             =   375
         Width           =   750
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         Height          =   225
         Left            =   60
         TabIndex        =   11
         Top             =   1365
         Width           =   750
      End
      Begin VB.Label Label2 
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
         Height          =   285
         Left            =   1995
         TabIndex        =   10
         Top             =   1320
         Width           =   6570
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   945
         TabIndex        =   9
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   225
         Left            =   105
         TabIndex        =   6
         Top             =   855
         Width           =   750
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   945
         TabIndex        =   2
         Top             =   795
         Width           =   2010
      End
      Begin VB.Label Label8 
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
         Height          =   285
         Left            =   3045
         TabIndex        =   1
         Top             =   795
         Width           =   5520
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00062A70&
      Height          =   6990
      Left            =   8925
      ScaleHeight     =   6930
      ScaleWidth      =   1320
      TabIndex        =   3
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command3 
         Caption         =   "Salir"
         Height          =   640
         Left            =   120
         Picture         =   "CANJECOMP.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   200
         Width           =   640
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   4
         Top             =   6195
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   5
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "CANJECOMP.frx":0458
         Top             =   6510
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quitar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   105
      TabIndex        =   7
      Top             =   2100
      Width           =   8700
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
         ItemData        =   "CANJECOMP.frx":237A
         Left            =   165
         List            =   "CANJECOMP.frx":2381
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   630
         Width           =   8370
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Código                    Descripción                                                     Cant.  "
         Height          =   270
         Left            =   165
         TabIndex        =   15
         Top             =   315
         Width           =   8370
      End
   End
End
Attribute VB_Name = "CANJECOMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    command1.Enabled = False
    Screen.MousePointer = 11

    'VALIDO INFORAMACION DE ENCABEZADO
    CI% = CODINT%(XVALO(Label11))
    TIPOCANJ% = XVALO(Label1)
    If CI% < 1 Or TIPOCANJ% < 1 Then GoTo 99
    '
    'MARCO LOS CAMPOS PARA BORRAR LUEGO
    CONDI$ = "CODICONJ = " & CI% & " AND TIPOCANJE  = " & TIPOCANJ%
    Call ACTUREGISTRO("CANJPREDEF", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
    '
    For JX% = 0 To List1.ListCount - 1
          L1$ = List1.List(JX%)
          CODI0$ = Mid$(L1$, 1, 16)
          CIELEM% = CODINT%(CODI0$)
          QUIT = XVALO(Mid$(L1$, 54, 10))
          'GRABO EN TABLA SQL
          Call GRABOCANJE(CI%, CIELEM%, TIPOCANJ%, 0, QUIT)
          '
    Next JX%
    '
    For JX% = 0 To List2.ListCount - 1
        L1$ = List2.List(JX%)
        CODI0$ = Mid$(L1$, 1, 16)
        CIELEM% = CODINT%(CODI0$)
        AGRE = XVALO(Mid$(L1$, 54, 10))
        'GRABO EN TABLA SQL
        Call GRABOCANJE(CI%, CIELEM%, TIPOCANJ%, AGRE, 0)
    Next JX%
    '
    'BORRO LOS REGISTROS MARCADOS
    CONDI$ = "CODICONJ = " & CI% & " AND TIPOCANJE  = " & TIPOCANJ% & " And MARBORRA = 1"
    Call BORRAREGISTROS("CANJPREDEF", CONDI$, REGAF&, "NOMESS")
    '
    Call COMUNI("Grabación Realizada")
    '
99  Screen.MousePointer = 1
    command1.Enabled = True
    '
End Sub
Sub GRABOCANJE(CI%, CIELEM%, TIPOCANJ%, AGRE, QUIT)

10 RNDX$ = RANDSTRING$(-8)
   If CantRegistros("CANJPREDEF", "AGREGAR = '" & RNDX$ & "'", "NOMESS") > 0 Then GoTo 10
   Call CreaRegistro("CANJPREDEF", "AGREGAR", RNDX$, "NOMESS")
   CONDI$ = "AGREGAR = '" & RNDX$ & "'"
   Call ACTUREGISTRO("CANJPREDEF", "CODICONJ", CONDI$, CI%, REGAF&, "NOMESS")
   Call ACTUREGISTRO("CANJPREDEF", "TIPOCANJE", CONDI$, TIPOCANJ%, REGAF&, "NOMESS")
   Call ACTUREGISTRO("CANJPREDEF", "CODIELEM", CONDI$, CIELEM%, REGAF&, "NOMESS")
   Call ACTUREGISTRO("CANJPREDEF", "MARBORRA", CONDI$, 0, REGAF&, "NOMESS")
   Call ACTUREGISTRO("CANJPREDEF", "QUITAR", CONDI$, FORMATNUM$(QUIT, "F16.5"), REGAF&, "NOMESS")
   '
   Call ACTUREGISTRO("CANJPREDEF", "AGREGAR", CONDI$, FORMATNUM$(AGRE, "F16.5"), REGAF&, "NOMESS")
   '
End Sub

Private Sub Command3_Click()
    Call CIERRARCH("*.*")
    Unload Me
   
End Sub

Private Sub label1_Change()
    '
    TIPOCANJ% = XVALO(Label1)
    Label2 = TRIM$(ECOARCH$("TATICANJ", MKI$(TIPOCANJ%)))
    CI% = CODINT%(Label11)
    '
    'VALIDO
    If CI% < 1 Then
       Call COMUNI("Código Inexistente, Verifique !!!")
       Exit Sub
    End If
    '
    If TIPOCANJ% < 1 Then
       Call COMUNI("Tipo de Canje Inexistente, Verifique !!!")
       Exit Sub
    End If
    '
    CONDI$ = "CODICONJ = " & CI% & " AND TIPOCANJE  = " & TIPOCANJ%
    If CantRegistros&("CANJPREDEF", CONDI$, "NOMESS") > 0 Then
       CONDI1$ = CONDI$ + " AND QUITAR > 0"
       CONDI2$ = CONDI$ + " AND AGREGAR > 0"
       Call CARGALISTA(List1, "CANJPREDEF", "CODIELEM/C16+MASTER/A35;QUITAR/F9.1", CONDI1$, "")
       Call CARGALISTA(List2, "CANJPREDEF", "CODIELEM/C16+MASTER/A35;AGREGAR/F9.1", CONDI2$, "")
    End If
    '
 
End Sub

Private Sub Label11_Change()
    CI% = CODINT%(XVALO(Label11))
    Label8 = DESCRIT0$(CI%)
End Sub


Private Sub Label14_DblClick()
   Call CARGAVENTANA(List1)
End Sub
Private Sub Label3_DblClick()
   Call CARGAVENTANA(List2)
End Sub
Sub CARGAVENTANA(OBJ_LIST As ListBox)

    Call BLANARCH("!CANJCOMP")
    'LEO EL LIST
    JJ& = 0
    For JX% = 0 To OBJ_LIST.ListCount - 1
      L1$ = OBJ_LIST.List(JX%)
      CODI0$ = Mid$(L1$, 1, 16)
      CI% = CODINT%(CODI0$)
      CAN1 = XVALO(Mid$(L1$, 54, 10))
      '
      'ESCRIBO EN LA VENTANA
      YY$ = REGBLAN$("CANJCOMP")
      Call REPLA(YY$, MKI$(CI%), 1, 2)
      Call REPLA(YY$, MKS$(CAN1), 3, 4)
      JJ& = JJ& + 1
      Call GRAREG("!CANJCOMP", YY$, JJ&)
      '
    Next JX%
    Call SETULTREG("!CANJCOMP", JJ&)
    
50  ATRI$ = " Tabla de Canje de Componentes"
    VTB% = VENTABU%("CANJCOMP/" + ATRI$)
   
    If VTB% < 0 Then Exit Sub
    '
    'VALIDO
    FIN& = ULTREG&("!CANJCOMP")
    For J& = 1 To FIN&
      X$ = REGLEIDO$("!CANJCOMP", J&)
      CI% = CVI(Mid$(X$, 1, 2))
      CAN1 = CVS(Mid$(X$, 3, 4))
      If CI% < 1 Then Call COMUNI("Código No Válido en Fila: " & J&): GoTo 50
      If CAN1 < 0.05 Then Call COMUNI("Falta Cantidad en Fila: " & J&): GoTo 50
    Next J&

    '
    OBJ_LIST.Clear
    '
    'LEO LA VENTANA
    FIN& = ULTREG&("!CANJCOMP")
    For J& = 1 To FIN&
      X$ = REGLEIDO$("!CANJCOMP", J&)
      CI% = CVI(Mid$(X$, 1, 2))
      CODI2$ = CODEXT$(CI%)
      DESC2$ = DESCRIT0(CI%)
      CAN1 = CVS(Mid$(X$, 3, 4))
      CAN2$ = FORMATNUM$(CAN1, "F9.1")
      'AGREGO AL LIST
      L1$ = Space(128)
      Call REPLA(L1$, CODI2$, 1, 16)
      Call REPLA(L1$, DESC2$, 18, 35)
      Call REPLA(L1$, CAN2$, 54, 10)
      OBJ_LIST.AddItem L1$
    Next J&
    '

End Sub

Private Sub List1_DblClick()
   If List1.ListCount > 0 Then
      Call CARGAVENTANA(List1)
   End If
End Sub

Private Sub List2_Click()
   
   If List2.ListCount > 0 Then
      Call CARGAVENTANA(List2)
   End If

End Sub
