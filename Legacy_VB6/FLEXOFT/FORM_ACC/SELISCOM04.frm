VERSION 5.00
Begin VB.Form SELISCOM04 
   BackColor       =   &H00EEF0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión Compras de Materiales"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEF0F0&
      Caption         =   "Salida o Transacción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   210
      TabIndex        =   8
      Top             =   1470
      Width           =   3480
      Begin VB.OptionButton Option4 
         BackColor       =   &H00EEF0F0&
         Caption         =   "Generar Ordenes de Compra"
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
         Left            =   420
         TabIndex        =   10
         Top             =   460
         Width           =   2850
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00EEF0F0&
         Caption         =   "Pedido de Materiales (interno)"
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
         Left            =   420
         TabIndex        =   9
         Top             =   735
         Value           =   -1  'True
         Width           =   2955
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   3090
      Left            =   3885
      ScaleHeight     =   3030
      ScaleWidth      =   1020
      TabIndex        =   1
      Top             =   -105
      Width           =   1080
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   130
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   4
         Top             =   1800
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   12000
         End
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
         Height          =   600
         Left            =   130
         Picture         =   "SELISCOM04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancela la Transacción"
         Top             =   400
         Width           =   600
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
         Height          =   600
         Left            =   130
         Picture         =   "SELISCOM04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Ejecuta Transacción Seleccionada"
         Top             =   1000
         Width           =   600
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "SELISCOM04.frx":0614
         Top             =   2310
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00EEF0F0&
      Caption         =   "Tipo de Material"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   3480
      Begin VB.OptionButton Option2 
         BackColor       =   &H00EEF0F0&
         Caption         =   "Materiales Libres (s/Cod)"
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
         Left            =   420
         TabIndex        =   7
         Top             =   735
         Width           =   2745
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EEF0F0&
         Caption         =   "Material Codificado"
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
         Left            =   420
         TabIndex        =   6
         Top             =   460
         Value           =   -1  'True
         Width           =   2220
      End
   End
End
Attribute VB_Name = "SELISCOM04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    '
    Call CIERRARCH("*.*")
    Call FINAL("")
    '
End Sub


Private Sub Command2_Click()
    '
    If SELISCOM04.Option4.Value = True Then
        Call GENEOCOM
    Else
        MODO% = 1
        Call LISCOPEN(MODO%)
    End If
    '
End Sub

Private Sub Form_Load()
    Call CENTRAFORM(Me)
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FUERASYS
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
End Sub

Sub GENEOCOM()
   ' Genera Ordenes de Compra
   '
   HOII% = HOY(HOS$)
   '
'   Call ABREORTRA
'   SQLX$ = "SELECT * FROM LISTAMAT "
'   SQLX$ = SQLX$ + " WHERE (Nro_Ocom = '' or isnull(Nro_Ocom) = true OR IsEmpty(Nro_Ocom) = true)" 'CONDICION PENDIENTE DE COMPRA
'   SQLX$ = SQLX$ + " ORDER BY Nro_Otrab, Nro_Grup, Nuorden "
'   Set LMATPEN = OTrabajo.OpenRecordset(SQLX$, 4)
'   '
   MODO% = 2
   Call LISCOPEN(MODO%)
   If MODO% < 0 Then  ' cancelado
     Exit Sub
   End If
   '
   If SELISCOM04.Option1.Value = True Then
       ' MATERIAL CODIFICADO
       ITECOM04.List1.Clear
       JJ& = 0
       FIN& = ULTREG&("LIMAPECO")
       For U& = 1 To FIN&
         XX$ = REGLEIDO$("LIMAPECO", U&)
         CIXI% = CVI(Left$(XX$, 2))
         CDEX$ = CODEXT$(CIXI%)
         NOTRA$ = TRIM$(Mid$(XX$, 3, 16))
         NUGRU$ = TRIM$(Mid$(XX$, 19, 10))
         CANEPAR = CVS(Mid$(XX$, 29, 4))
         FENECESI% = CVI(Mid$(XX$, 37, 2))
         NORIT% = CVI(Mid$(XX$, 39, 2))
         If FENECESI% < HOII% Then FENECECI% = HOII%
         FEX = FENECESI%: FENECE$ = FECHATEX$(FEX)
         '
         For KKI% = 1 To ITECOM04.List1.ListCount
           TXT1$ = ITECOM04.List1.List(KKI% - 1)
           If TRIM$(Left$(TXT1$, 16)) = CDEX$ Then  ' COINCIDE CODIGO
             If Mid$(TXT1$, 79, 8) = FENECE$ Then   ' TAMBIEN COINCIDE SEMANA
               CANUE = Val(Mid$(TXT1$, 65, 12)) + CANEPAR
               Call REPLA(TXT1$, FORMATNUM$(CANUE, "F12.3"), 65, 12)
               TXT1$ = TXT1$ + Chr$(124) + AJUSTI$(NOTRA$, 24) + AJUSTI$(NUGRU$, 16) + FORMATNUM(NORIT%, "I7")
               ITECOM04.List1.RemoveItem (KKI% - 1)
               ITECOM04.List1.AddItem TXT1$, KKI% - 1
               GoTo 19
             End If
           End If
         Next KKI%
         '
         TXT1$ = AJUSTI$(CDEX$, 16) + AJUSTI$(DESCRIT0$(CIXI%), 48) + Space$(26) ' TOTAL = 90 CAR
         Call REPLA(TXT1$, FORMATNUM$(CANEPAR, "F12.3"), 65, 12)
         Call REPLA(TXT1$, FENECE$, 79, 8)
         TXT1$ = TXT1$ + Chr$(124) + AJUSTI$(NOTRA$, 24) + AJUSTI$(NUGRU$, 16) + FORMATNUM(NORIT%, "I7")
         ITECOM04.List1.AddItem TXT1$
         '
19        Next U&
       '
       If ITECOM04.List1.ListCount < 1 Then
           Call MENSERR(24, "No hay Materiales Codificados\Pendientes de Compra\Según Listas de Materiales.")
           Exit Sub
         Else
           On Error Resume Next
           ITECOM04.Show 1
           ITECOM04.Refresh
           DoEvents
           On Error GoTo 0
       End If
       '
     Else
       'MATERIAL DE TEXTO LIBRE
        COMLIB04.List1.Clear
        JJ& = 0
        FIN& = ULTREG&("LIMAPELI")
        For U& = 1 To FIN&
          XX$ = REGLEIDO$("LIMAPELI", U&)
          NOTRA$ = TRIM$(Left$(XX$, 16))
          NUGRU$ = TRIM$(Mid$(XX$, 17, 10))
          NORIT% = CVI(Mid$(XX$, 27, 2))
          CANIT = CVD(Mid$(XX$, 29, 8))
          DECOR$ = TRIM$(Mid$(XX$, 37, 48))
          FEX = CVI(Mid$(XX$, 183, 2))
          '
          XXY$ = Space$(100)
          Call REPLA(XXY$, NOTRA$, 1, 16)
          Call REPLA(XXY$, NUGRU$, 17, 10)
          Call REPLA(XXY$, FORMATNUM$(NORIT%, "I4"), 27, 4)
          Call REPLA(XXY$, FORMATNUM$(CANIT, "F10.3"), 31, 10)
          Call REPLA(XXY$, DECOR$, 43, 48)
          Call REPLA(XXY$, FECHATEX$(FEX), 93, 8)
           '
          COMLIB04.List1.AddItem XXY$
        Next U&
        '
        If COMLIB04.List1.ListCount < 1 Then
            Call MENSERR(24, "No hay Materiales Libres\Pendientes de Compra\Según Listas de Materiales.")
            Exit Sub
          Else
            COMLIB04.Show 1
        End If
        '
   End If
   Screen.MousePointer = 1
   '
End Sub
   '
Sub LISCOPEN(MODO%)     ' MODO% = 1 --> Emite Listado // Modo% = 2 -- > Genera Lista de Compra
   '
   Call SETULTREG("LIMAPECO", 0)
   Call SETULTREG("LIMAPELI", 0)
   '
10 OBX$ = OBJPLA$("SELEVCLAVE", VDEF$)
   If OBX$ = "" Then
      MODO% = (-1)  ' SEÑAL DE CANCELACION
      Exit Sub
   End If
   '
   VDEF$ = OBX$
   NOTRA$ = TRIM$(Left$(OBX$, 16))
   If NOTRA$ <> "" Then
     If ECOARCH$("INDIOTRA", NOTRA$) = "" Then
       Call MENSERR(24, "Numero de E.V. no Válido")
       GoTo 10
     End If
   End If
   
   Call ABREORTRA
   SQLX$ = "SELECT * FROM LISTAMAT "
   SQLX$ = SQLX$ + " WHERE (Nro_Ocom = '' or isnull(Nro_Ocom) = true OR IsEmpty(Nro_Ocom) = true) " 'CONDICION PENDIENTE DE COMPRA
   'ACA SI HAY FILTRO DE EV
   If NOTRA$ <> "" Then
      SQLX$ = SQLX$ + " AND Nro_Otrab = '" & NOTRA$ & "'"
   End If
   '
   LISCLA$ = TRIM$(Mid$(OBX$, 17, 36))
   If LISCLA$ <> "" Then
     'ACA SI HAY FILTRO DE LLAVES
     If Left$(LISCLA$, 1) <> "+" Then
       If Left$(LISCLA$, 1) <> "-" Then
         LISCLA$ = "+" + LISCLA$
       End If
     End If
     '
     CONDI$ = ""
     While Len(LISCLA$) > 0
       VALA$ = Left$(LISCLA$, 3)
       LLA$ = Mid(VALA$, 2, 2)
      
       If Left$(VALA$, 1) = "+" Then
           If CONDI$ <> "" Then CONDI$ = CONDI$ + " OR "
           CONDI$ = CONDI$ + "Clave = '" & LLA$ & "' "
         Else
            If CONDI$ <> "" Then CONDI$ = CONDI$ + " AND "
           CONDI$ = CONDI$ + "Clave <> '" & LLA$ & "' "
       End If
       LISCLA$ = Mid$(LISCLA$, 4)
     Wend
     '
     SQLX$ = SQLX$ + " AND (" + CONDI$ + ") "
   End If
   '
   SQLX$ = SQLX$ + " ORDER BY Nro_Otrab, Nro_Grup "
   '
   Set LMATPEN = OTrabajo.OpenRecordset(SQLX$, 4)
   '
   With LMATPEN
     On Error Resume Next
     'Valida que la tabla no este vacia
     .MoveFirst
     If Err Then
        Call MENSERR(24, "No Hay Materiales Pendientes de Compra")
        On Error GoTo 0
        Exit Sub
     End If
     On Error GoTo 0
     '
     If SELISCOM04.Option1.Value = True Then
          ' MATERIAL CODIFICADO
          JJ& = 0
          Do While (.EOF = False)
            XXY$ = REGBLAN$("LIMAPECO")
            '
            CIXI% = CODINT(!CODEXTE)
            If CIXI% > 0 Then
              '
              ' calcula fecha de necesidad
              NOTRA$ = !Nro_Otrab
              NUGRU$ = !Nro_Grup
              NORIT% = !NuOrden
              CLAVX% = Val(!CLAVE) Mod 10
              '
              FENECESI% = 0
              If CLAVX% = 5 Or CLAVX% = 7 Then
                GrupOtra.FindFirst "Nro_Otrab = '" & NOTRA$ & "'AND Nro_Grup = '" & NUGRU$ & "' "
                If GrupOtra.NoMatch = False Then
                  PLAX$ = "0000"
                  On Error Resume Next
                  PLAX$ = GrupOtra!PLA05_GRUP
                  If CLAVX% = 7 Then PLAX$ = GrupOtra!PLA07_GRUP
                  On Error GoTo 0
                  FENECESI% = FECHSEMAN(PLAX$)
                End If
              End If
              '
              Call REPLA(XXY$, MKI(CIXI%), 1, 2)
              Call REPLA(XXY$, !Nro_Otrab, 3, 16)
              Call REPLA(XXY$, !Nro_Grup, 19, 10)
              Call REPLA(XXY$, MKS(!CANTID), 29, 4)
              Call REPLA(XXY$, MKI$(FENECESI%), 37, 2)
              Call REPLA(XXY$, MKI$(NORIT%), 39, 2)
              JJ& = JJ& + 1
              Call GRAREG("LIMAPECO", XXY$, JJ&)
            End If
          .MoveNext
          Loop
          '
          Call SETULTREG("LIMAPECO", JJ&)
          Call CIERRARCH("LIMAPECO")
          Screen.MousePointer = 1
          Call FILESORT("LIMAPECO", "LIMAPECO", 7, 7, "ASC")
          Call FILESORT("LIMAPECO", "LIMAPECO", 1, 1, "ASC")
          If MODO% = 1 Then
            Call FINAL("*LIMAPECO")
          End If
       Else
          'MATERIAL DE TEXTO LIBRE
          JJ& = 0
          Do While (.EOF = False)
            XXY$ = REGBLAN$("LIMAPELI")
            '
            CIXI% = CODINT(!CODEXTE)
            If CIXI% <= 0 Then    ' PARA EXCLUIR MATERIAL CODIFICADO
              '
              ' calcula fecha de necesidad
              NOTRA$ = !Nro_Otrab
              NUGRU$ = !Nro_Grup
              CLAVX% = Val(!CLAVE) Mod 10
              '
              FENECESI% = 0
              If CLAVX% = 5 Or CLAVX% = 7 Then
                GrupOtra.FindFirst "Nro_Otrab = '" & NOTRA$ & "'AND Nro_Grup = '" & NUGRU$ & "' "
                If GrupOtra.NoMatch = False Then
                  PLAX$ = "0000"
                  On Error Resume Next
                  PLAX$ = GrupOtra!PLA05_GRUP
                  If CLAVX% = 7 Then PLAX$ = GrupOtra!PLA07_GRUP
                  On Error GoTo 0
                  FENECESI% = FECHSEMAN(PLAX$)
                End If
              End If
              '
              Call REPLA(XXY$, !Nro_Otrab, 1, 16)
              Call REPLA(XXY$, !Nro_Grup, 17, 10)
              Call REPLA(XXY$, MKI(!NuOrden), 27, 2)
              Call REPLA(XXY$, MKD(!CANTID), 29, 8)
              Call REPLA(XXY$, !Descri_Cor, 37, 48)
              Call REPLA(XXY$, !NPlano, 85, 24)
              Call REPLA(XXY$, !Posic, 109, 16)
              Call REPLA(XXY$, !MATERIAL, 125, 24)
              Call REPLA(XXY$, !CODEXTE, 149, 24)
              Call REPLA(XXY$, !CLAVE, 173, 10)
              Call REPLA(XXY$, MKI$(FENECESI%), 183, 2)
              JJ& = JJ& + 1
              Call GRAREG("LIMAPELI", XXY$, JJ&)
            End If
          .MoveNext
          Loop
          '
          Call SETULTREG("LIMAPELI", JJ&)
          Call CIERRARCH("LIMAPELI")
          Screen.MousePointer = 1
          Call FILESORT("LIMAPELI", "LIMAPELI", 1, 3, "ASC")
          If MODO% = 1 Then
            Call FINAL("*LIMAPELI")
          End If
     End If
   End With
   '
End Sub


