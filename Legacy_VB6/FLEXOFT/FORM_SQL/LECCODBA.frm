VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form LECCODBA 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lectura de Código de Barra"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   435
      Left            =   105
      TabIndex        =   9
      Top             =   105
      Width           =   9990
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"LECCODBA.frx":0000
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   50
         TabIndex        =   10
         Top             =   120
         Width           =   9900
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   7
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton BOTNEXT 
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
      Height          =   660
      Left            =   10185
      Picture         =   "LECCODBA.frx":00A2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cargar Items"
      Top             =   4410
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4110
      TabIndex        =   0
      ToolTipText     =   "Ingreso Manual o por Lectora"
      Top             =   4680
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   4440
      Left            =   10185
      ScaleHeight     =   4380
      ScaleWidth      =   1110
      TabIndex        =   2
      Top             =   -105
      Width           =   1170
      Begin VB.CommandButton Command26 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "LECCODBA.frx":03AC
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Configuracion de Lectura de Código de Barra"
         Top             =   2100
         Width           =   650
      End
      Begin VB.CommandButton Command24 
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
         Left            =   105
         Picture         =   "LECCODBA.frx":07EE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Quitar  Producto de la Lista"
         Top             =   1260
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "LECCODBA.frx":0AF8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "LECCODBA.frx":0C42
         Top             =   3990
         Width           =   1515
      End
   End
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
      Height          =   3660
      ItemData        =   "LECCODBA.frx":2B64
      Left            =   105
      List            =   "LECCODBA.frx":2B66
      TabIndex        =   1
      Top             =   525
      Width           =   9990
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4620
      OleObjectBlob   =   "LECCODBA.frx":2B68
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "LECCODBA.frx":2D9C
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Barras (Artículo) >>>>>>>>:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   960
      TabIndex        =   6
      Top             =   4755
      Width           =   3150
   End
End
Attribute VB_Name = "LECCODBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BOTNEXT_Click()
        '
        BOTNEXT.Enabled = False
        If List1.ListCount < 1 Then
           Call COMUNI("No Se Han Ingresado Artículos")
           GoTo 99
        End If
        '
        Hide
99      BOTNEXT.Enabled = True


End Sub

Private Sub Command2_Click()
  '
  Command2.Enabled = False
  List1.Clear
  Hide
  Command2.Enabled = True
  '
End Sub
Private Sub Command24_Click()
    '
    If List1.ListIndex >= 0 Then
      List1.RemoveItem (List1.ListIndex)
      Text1.SetFocus
    End If
    '
End Sub

Private Sub Command26_Click()
   Command26.Enabled = False
   CNFGBARRA.Show 1
   Command26.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call APLICACIONSKINS(Me, Picture1)
    '
End Sub

Private Sub List1_Click()
    '
    Text1.SetFocus
    '
End Sub

Private Sub Text1_Change()

    If Len(TRIM$(Text1)) = 25 Then
       Call Text1_KeyPress(13)
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
        '
        'VALIDACIONES
        '
        If KeyAscii <> 13 Then Exit Sub
        If TRIM$(Text1) = "" Then Exit Sub

        'LECTURA DE POSICIONES CONFIGURADAS
        If TRIM$(Control("LECCODBA", "OPLECTUR")) = "LONGFIJA" Then
            PCOD_IN% = XVALO(Control("LECCODBA", "CODIINIC"))
            PCOD_FIN% = XVALO(Control("LECCODBA", "CODIGFIN"))
            PCANT_IN% = XVALO(Control("LECCODBA", "CANTINIC"))
            PCANT_FIN% = XVALO(Control("LECCODBA", "CANTIFIN"))
            If PCOD_IN% < 1 Or PCOD_FIN% < 1 Or PCANT_IN% < 1 Or PCANT_FIN% < 1 Then
               Call COMUNI("Falta Información de Posiciones de Lectura")
               CNFGBARRA.Show 1
               GoTo 60
            End If
        End If
        '
        If TRIM$(Control("LECCODBA", "OPLECTUR")) = "CARACTER" Then
           CARCTR$ = TRIM$(Control("LECCODBA", "CARASEPA"))
           If TRIM$(CARCTR$) = "" Then Call COMUNI("Falta Ingresar Caracter Separador"): CNFGBARRA.Show 1: GoTo 60
        End If
        '
        'OBTENGO CODIGO INTERNO DE EXTERNO O UPC
        CODI$ = LEECODIGOBARRA$(Text1, PCOD_IN%, PCOD_FIN%)
        CANTI = LEECANTIBARRA(Text1, PCANT_IN%, PCANT_FIN%)
        '
        CI% = CODINT%(CODI$) ' CODIGO INTERNO DE CODIGO EXTERNO
        If CI% < 1 Then
          CI% = CODINTBARRA%(CODI$) ' CODIGO INTERNO DE UPC
        End If
        '
        'VALIDO QUE EL CODIGO SEA VALIDO
        If CI% < 1 Then
           Call MENSERR(24, "Código Inexistente o No Válido")
           Text1.SelStart = 0
           Text1.SelLength = Len(Text1)
           GoTo 60
        End If
        '
        Call REPLA(TX$, CODI$, 1, 24)
        Call REPLA(TX$, Left(DESCRIT$(CI%), 30), 26, 36)
        Call REPLA(TX$, FORMATNUM$(CANTI, "F16.1"), 64, 16)
        '
        'COMPARAR SI EXISTE EN LA LISTA, EN ESE CASO ACUMULO SOBRE EL MISMO
        For U& = 1 To List1.ListCount
          If TRIM$(Left$(List1.List(U& - 1), 24)) = CODI$ Then
             CANTLIST = XVALO(Mid$(List1.List(U& - 1), 64, 16))
             Call REPLA(TX$, FORMATNUM$((CANTLIST + CANTI), "F16.1"), 64, 16)
             List1.RemoveItem U& - 1
             List1.AddItem TX$, U& - 1
             GoTo 50
          End If
        Next U&
        '
        List1.AddItem TX$
        '
50      Text1 = ""
60      Text1.SetFocus
    '
   End Sub
Function LEECODIGOBARRA$(TXT1$, PINI%, PFIN%)
   TXT$ = TRIM$(TXT1$)
   If TXT$ = "" Then GoTo 99
   '
   
   If Control("LECCODBA", "OPLECTUR") = "CARACTER" Then
    CARCTR$ = TRIM$(Control("LECCODBA", "CARASEPA"))
    If CARCTR$ <> "" Then
      PFIN% = InStr(TXT$, CARCTR$) - 1
      PINI% = 1
    Else
      TXT$ = ""
      GoTo 99
    End If
   End If

   If PINI% < 1 Then TXT$ = "": GoTo 99
   If PFIN% < PINI% Then TXT$ = "": GoTo 99
   TXT$ = Mid$(TXT$, PINI%, PFIN% - PINI% + 1)
   '
99 LEECODIGOBARRA$ = TXT$
End Function

Function LEECANTIBARRA(TXT1$, PINI%, PFIN%)
   TXT$ = TRIM$(TXT1$)
   If TXT$ = "" Then GoTo 99
   '
   If Control("LECCODBA", "OPLECTUR") = "CARACTER" Then
    CARCTR$ = TRIM$(Control("LECCODBA", "CARASEPA"))
    If CARCTR$ <> "" Then
      PINI% = InStr(TXT$, CARCTR$) + 1
      PFIN% = Len(TXT$)
    Else
      TXT$ = ""
      GoTo 99
    End If
   End If
   '
   If PINI% < 1 Then CANTI = 0: GoTo 99
   If PFIN% < PINI% Then CANTI = 0: GoTo 99
   CANTI = XVALO(Mid$(TXT$, PINI%, PFIN% - PINI% + 1))
   '
99 LEECANTIBARRA = CANTI
   '
End Function

Function LECT_CODIGO%(List As ListBox, Fila%)
     CI% = CODINT(TRIM$(Left$(List.List(Fila%), 24)))
     LECT_CODIGO% = CI%
End Function

Function LECT_CANTIDAD(List As ListBox, Fila%)
     CANTI = XVALO(Mid$(List.List(Fila%), 64, 16))
     LECT_CANTIDAD = CANTI
End Function


