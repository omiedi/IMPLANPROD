VERSION 5.00
Begin VB.Form LITMAOP 
   BackColor       =   &H00C0FFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copia de Lista de Materiales"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   3090
      Left            =   6405
      ScaleHeight     =   3030
      ScaleWidth      =   1020
      TabIndex        =   7
      Top             =   -105
      Width           =   1080
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   615
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   10
         Top             =   2280
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   11
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
         Picture         =   "LITMAOP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   360
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
         Picture         =   "LITMAOP.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Ejecuta  la Copia"
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Coef."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   215
         TabIndex        =   17
         Top             =   1680
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "LITMAOP.frx":0614
         Top             =   2480
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Rango de Especificaciones de Seleccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   6210
      Begin VB.ComboBox Combo2 
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
         Left            =   4095
         TabIndex        =   13
         Text            =   "Grupos"
         ToolTipText     =   "Grupos"
         Top             =   1725
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   4095
         TabIndex        =   12
         Text            =   "Grupos"
         ToolTipText     =   "Grupos"
         Top             =   525
         Width           =   1935
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
         Height          =   315
         Left            =   3150
         TabIndex        =   4
         Top             =   1725
         Width           =   175
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   3
         Text            =   " "
         Top             =   1725
         Width           =   1890
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
         Height          =   315
         Left            =   3130
         TabIndex        =   2
         Top             =   525
         Width           =   175
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   1
         Top             =   525
         Width           =   1890
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
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
         Height          =   315
         Left            =   1260
         TabIndex        =   15
         Top             =   2145
         Width           =   4740
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   315
         Left            =   1260
         TabIndex        =   14
         Top             =   945
         Width           =   4740
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   0
         X2              =   6240
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hacia EV:"
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
         Left            =   -1140
         TabIndex        =   6
         Top             =   1830
         Width           =   2325
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde EV:"
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
         Left            =   -1140
         TabIndex        =   5
         Top             =   630
         Width           =   2325
      End
   End
End
Attribute VB_Name = "LITMAOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    '
    Unload Me
    '
End Sub

Private Sub Command2_Click()
    '
    If TRIM$(Text10) = "" Or ECOT$(Text10) = "" Then
       Call MENSERR(24, "E.V. 'Desde' no Válida o Inexistente")
       Exit Sub
    End If
    '
    If Combo1 = "" Then
        Call COMUNI("Debe Seleccionar el Grupo 'Desde E.V.'")
        Exit Sub
    End If
    '
    If TRIM$(Text8) = "" Or ECOT$(Text8) = "" Then
       Call MENSERR(24, "E.V. 'Hacia' no Válida o Inexistente")
       Exit Sub
    End If
    '
    If Combo2 = "" Then
        Call COMUNI("Debe Seleccionar el Grupo 'Hacia E.V.'")
        Exit Sub
    End If
    '
    If TRIM$(Text10) = TRIM$(Text8) Then
      If Combo1 = Combo2 Then
        Call MENSERR(24, "Copia Imposible !!!\Coinciden Origen y Destino de Copia.")
        Exit Sub
      End If
    End If
    '
    If TRIM$(Text1) = "" Or Val(Text1) < 0.5 Or Val(Text1) > 99.9 Then
        Call COMUNI("Coeficiente no Válido.")
        Text1.SelStart = 0
        Text1.SelLength = Len(TRIM(Text1))
        Text1.SetFocus
        Exit Sub
    End If
    '
    Call GRALITMA
    '
End Sub
'
Sub GRALITMA()
    '
    COEF1# = Val(Text1)
    '
    Call ABREORTRA
    On Error GoTo 0
    SQLX$ = " SELECT * FROM LISTAMAT "
    SQLX$ = SQLX$ + " WHERE TRIM$(Nro_Otrab) = '" & TRIM$(Text10) & "' "
    SQLX$ = SQLX$ + " AND TRIM$(Nro_Grup) = '" & TRIM$(Combo1) & "' "
    Set GRALITMP = OTrabajo.OpenRecordset(SQLX$, 4)
    With GRALITMP
        On Error Resume Next
        .MoveFirst
        If Err Then
            Call COMUNI("Copia Imposible !!!\Grupo Origen sin Lista de Materiales.")
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
        Do While .EOF = False
           With ListaMat
            .AddNew
             !Nro_Otrab = TRIM$(Text8)
             !Nro_Grup = TRIM$(Combo2)
             !NuOrden = GRALITMP.Fields("NUORDEN")
             !CANTID = (GRALITMP.Fields("Cantid") * COEF1#)
             !DESCRI_COR = GRALITMP.Fields("Descri_Cor")
             !NPlano = GRALITMP.Fields("NPlano")
             !Posic = GRALITMP.Fields("Posic")
             !Material = GRALITMP.Fields("Material")
             !CODEXTE = GRALITMP.Fields("CodExte")
             !Kilos = GRALITMP.Fields("Kilos")
             !CLAVE = GRALITMP.Fields("Clave")
             !Descri_Lar = GRALITMP.Fields("Descri_Lar")
            .Update
           End With
          .MoveNext
        Loop
        Call COMUNI("Copia Completa")
    End With
    Call Command1_Click
    '
End Sub

Private Sub Command8_Click()
    Call SELECHO("INDIOTRA")
    EV2$ = VALACT1$("INDIOTRA")
    If EV2$ <> "" Then
        Text8.TEXT = EV2$
        Label2 = ECOT$(EV2$)
    End If
End Sub

Private Sub Command9_Click()
    '
    Call SELECHO("INDIOTRA")
    EV1$ = VALACT1$("INDIOTRA")
    If EV1$ <> "" Then
        Text10.TEXT = EV1$
        Label1 = ECOT(EV1$)
    End If
    '
End Sub

Private Function ECOT$(EV$)
   '
   ECOT$ = ""
   Call ABREORTRA
   SQLX$ = "SELECT * FROM OrdeTRAB "
   SQLX$ = SQLX$ + " WHERE NRO_OTRAB = '" & EV$ & " ' "
   Set ECOTMP = OTrabajo.OpenRecordset(SQLX$, 4)
   '
   With ECOTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       ECOT$ = ""
       Exit Function
     Else
       ECOT$ = !Refe_Otrab
     End If
   End With
   On Error GoTo 0
   '
End Function

Private Sub Form_Load()
    '
    Text1 = 1
    Text1 = FORMATNUM(Text1, "F4.1")
    '
End Sub

Private Sub Text10_GotFocus()
   Text10.SelStart = 0
   Text10.SelLength = Len(Text10)
End Sub

Private Sub Text8_Change()
    Call ECOGRUP2(Text8)
End Sub



Private Sub Text8_GotFocus()
   Text8.SelStart = 0
   Text8.SelLength = Len(Text8)
End Sub

Private Sub Text9_Click()
    Text10.SetFocus
End Sub
Private Sub Text10_Change()
    Call ECOGRUP1(Text10)
End Sub

Sub ECOGRUP1(EV$)
   '
   Call ABREORTRA
   Combo1.Clear
   SQLX$ = "SELECT * FROM GRUPOTRA "
   SQLX$ = SQLX$ + " WHERE NRO_OTRAB = '" & EV$ & " ' "
   Set ECOTMP = OTrabajo.OpenRecordset(SQLX$, 4)
   '
   With ECOTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       Combo1 = ""
       Label1 = ""
       Exit Sub
     Else
       Label1 = ECOT(EV$)
       Do While .EOF = False
         Combo1.AddItem !Nro_Grup
         .MoveNext
       Loop
     End If
   End With
   On Error Resume Next
   Combo1 = Combo1.List(0)
   On Error GoTo 0
   '
End Sub

Sub ECOGRUP2(EV$)
   Call ABREORTRA
   Combo2.Clear
   SQLX$ = "SELECT * FROM GRUPOTRA "
   SQLX$ = SQLX$ + " WHERE NRO_OTRAB = '" & EV$ & " ' "
   Set ECOTMP = OTrabajo.OpenRecordset(SQLX$, 4)
   '
   With ECOTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       Combo2 = ""
       Label2 = ""
       Exit Sub
     Else
       Label2 = ECOT(EV$)
       Do While .EOF = False
         Combo2.AddItem !Nro_Grup
         .MoveNext
       Loop
     End If
   End With
   On Error Resume Next
   Combo2 = Combo2.List(0)
   On Error GoTo 0
   '
End Sub
