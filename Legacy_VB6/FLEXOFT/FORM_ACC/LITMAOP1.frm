VERSION 5.00
Begin VB.Form SELISCOM04 
   BackColor       =   &H00EEF0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copia de Lista de Materiales"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
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
      Left            =   105
      TabIndex        =   8
      Top             =   1470
      Width           =   3060
      Begin VB.OptionButton Option4 
         BackColor       =   &H00EEF0F0&
         Caption         =   "Generar Pedidos"
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
         Width           =   2220
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00EEF0F0&
         Caption         =   "Listado de Pendientes"
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
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   3090
      Left            =   3255
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
         Picture         =   "LITMAOP1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Abandona la Aplicación"
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
         Picture         =   "LITMAOP1.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Ejecuta  la Copia"
         Top             =   1000
         Width           =   600
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "LITMAOP1.frx":0614
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
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   3060
      Begin VB.OptionButton Option2 
         BackColor       =   &H00EEF0F0&
         Caption         =   "Texto Libre"
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
         Width           =   2115
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
    Call GRALITMA
    '
End Sub
'
Sub GRALITMA()
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
             !CANTID = GRALITMP.Fields("Cantid")
             !Descri_Cor = GRALITMP.Fields("Descri_Cor")
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

Private Sub Text10_GotFocus()
   Text10.SelStart = 0
   Text10.SelLength = Len(Text10)
End Sub

Private Sub Text8_Change()
    Call ECOGRUP2(Text8)
End Sub
Private Sub Text8_DblClick()
    Call SELECHO("DEUDOR1")
    NCMAX% = Val(VALACT1$("DEUDOR1"))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = RASOCLI$(NCMAX%)
    End If
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
       Exit Sub
     Else
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
       Exit Sub
     Else
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
