VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form BCODE09 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lectura de Código de Barra"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   12
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
      Height          =   690
      Left            =   9360
      Picture         =   "BCODE09.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Continúa Carga de Remito."
      Top             =   4440
      Width           =   650
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   105
      TabIndex        =   6
      Top             =   185
      Width           =   9150
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
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   9150
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   4545
      Left            =   9240
      ScaleHeight     =   4485
      ScaleWidth      =   1005
      TabIndex        =   2
      Top             =   -210
      Width           =   1065
      Begin VB.Frame Frame17 
         BackColor       =   &H00004080&
         Caption         =   "Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   630
         Left            =   90
         TabIndex        =   8
         Top             =   3360
         Width           =   630
         Begin VB.Label Label18 
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
            Height          =   270
            Left            =   105
            TabIndex        =   9
            Top             =   270
            Width           =   435
         End
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Help"
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
         Picture         =   "BCODE09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   1005
         Width           =   650
      End
      Begin VB.CommandButton Command24 
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
         Picture         =   "BCODE09.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Eliminar Producto de la Lista"
         Top             =   2580
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
         Picture         =   "BCODE09.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.CommandButton Command13 
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
         Left            =   105
         Picture         =   "BCODE09.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Configuracion de Funcionamiento"
         Top             =   1890
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "BCODE09.frx":0EAA
         Top             =   4080
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
      Height          =   4110
      Left            =   105
      TabIndex        =   1
      Top             =   450
      Width           =   9150
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "BCODE09.frx":2DCC
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "BCODE09.frx":3000
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label VALI_BCODE09 
      Caption         =   "VALI_BCODE09"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "BCODE09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TOTITFO&

Private Sub BOTNEXT_Click()
        '
        BOTNEXT.Enabled = False
        For U& = 1 To List1.ListCount
            CI% = CODINT(TRIM$(Left$(List1.List(U& - 1), 16)))
            CANTSA = XVALO(Mid$(List1.List(U& - 1), 49, 8))
            LOTE$ = TRIM$(Mid$(List1.List(U& - 1), 61, 12))
            DEPX% = XVALO(FOREMIT07.Text10)
            '
            CANTILOTE = SaldoLote(CI%, LOTE$, DEPX%)
            If CANTILOTE < CANTSA Then
               DEPOS$ = TRIM$(ECOARCH$("ECODEP", MKI$(DEPX%)))
               Call COMUNI("Lote Insuficiente Para Código: " & CODEXT$(CI%) & " - Lote: " & LOTE$ & "\En Depósito: " & DEPOS$ & " - Saldo: " & CANTILOTE)
               GoTo 99
            End If
            '
        Next U&
        '
        Label18 = ""
        VALI_BCODE09 = "APROB"
        Hide
99      BOTNEXT.Enabled = True
End Sub

Private Sub Command13_Click()
    '
    Command13.Enabled = False
    If DERACCE%("STPCOBAR", "Configuración de Funcionamiento de Codigos de Barra") = 2 Then
        STPKBN04.Show 1
    End If
    Command13.Enabled = True
    '
End Sub

Private Sub Command2_Click()
  '
  Command2.Enabled = False
  List1.Clear
  VALI_BCODE09 = ""
  Call CIERRARCH("*.*")
  Screen.MousePointer = 1
  Command2.Enabled = True
  Hide 'Unload Me
  '
End Sub

Private Sub Command24_Click()
    '
    If List1.ListIndex >= 0 Then
      List1.RemoveItem (List1.ListIndex)
      Text1.SetFocus
    End If
    '
    Label18 = List1.ListCount
End Sub

Private Sub Form_Load()
    '
    XXTITU$ = Space$(0)
    '
    VALI_BCODE09 = ""
    XXCA$ = "Codigo"
    XXTC$ = "   Cant."
    XXTD$ = "Lote"
    Descr$ = "Descripción"
    Call REPLA(XXTITU$, XXCA$, 1, 17)
    Call REPLA(XXTITU$, Descr$, 18, 30)
    Call REPLA(XXTITU$, XXTC$, 49, 8)
    Call REPLA(XXTITU$, XXTD$, 61, 4)
    '
    List2.AddItem XXTITU$
    '
    FORIPRE$ = TRIM$(Control$("", "FOREMIFA"))
    If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Remito.")
        OPREMIT07.Show 1
        Call FINAL("")
    End If
    TOTITFO& = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
    '
    Call APLICACIONSKINS(Me, Picture1)
    '
End Sub
Private Sub List1_Click()
    '
    Text1.SetFocus
    '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
      If Text1 <> "" Then
        '
        XXARTIDES$ = Space$(64)
        '
        XXCODI$ = Left$((Text1), 16)
        XXCANT1$ = Mid(Text1, 29, 8)
        
        XXCANT$ = FORMATNUM$(XVALO(XXCANT1$), "F8.0")
        XXTARD1$ = TRIM$(Mid(Text1, 17, 12))
        'VALIDACIONES
        CI% = CODINT%(XXCODI$)
        DESCRIX$ = DESCRIT$(CI%)
        If XVALO(XXCANT1$) < FRACMIN(CI%) Then
          Call MENSERR(24, "Cantidad no Válida")
          Text1.SelStart = 0
          Text1.SelLength = Len(Text1)
          GoTo 99
        End If
        If CODINT%(XXCODI$) <= 0 Then
          Call MENSERR(24, "Código no Válido")
          Text1.SelStart = 0
          Text1.SelLength = Len(Text1)
          GoTo 99
        End If
        '
        Call REPLA(XXARTIDES$, XXCODI$, 1, 16)
        Call REPLA(XXARTIDES$, DESCRIX$, 18, 30)
        Call REPLA(XXARTIDES$, XXCANT$, 49, 8)
        Call REPLA(XXARTIDES$, XXTARD1$, 61, 12) ' LOTE
        '
        If XXTARD1$ <> "" Then 'SI EXISTE EL NUMERO DE TRAJETA, COMPARA QUE NO ESTE REPETIDO EN LA LISTA
             For U& = 1 To List1.ListCount
               If TRIM$(Left$(List1.List(U& - 1), 16)) = TRIM$(XXCODI$) And TRIM$(Mid$(List1.List(U& - 1), 61, 12)) = TRIM$(XXTARD1$) Then ' COINCIDE EL CÓDIGO Y LOTE
                 '
                 ' ACUMULA EN EL REGISTRO
                 RELISTA$ = List1.List(U& - 1)
                 CANTOTA = XVALO(Mid$(RELISTA$, 49, 8)) + XVALO(XXCANT$)
                 XXCANT$ = FORMATNUM$(CANTOTA, "F8.0")
                 Call REPLA(RELISTA$, XXCANT$, 49, 8)
                 List1.RemoveItem U& - 1
                 List1.AddItem RELISTA$, U& - 1
                 GoTo 90
                 '
               End If
             Next U&
             '
        End If
        '
        List1.AddItem XXARTIDES$
        Label18 = List1.ListCount
        '
90      Text1 = ""
        Text1.SetFocus
        '
      End If
    End If
    '
99 End Sub



