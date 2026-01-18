VERSION 5.00
Begin VB.Form DOCONEC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vinculación Documentos Relacionados"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   10830
   Visible         =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Acciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   8715
      TabIndex        =   27
      Top             =   6930
      Width           =   2010
      Begin VB.CommandButton Command5 
         Height          =   420
         Left            =   1050
         Picture         =   "DOCONEC.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Crea un Nuevo Documento Vinculado en la Carpeta Especificada"
         Top             =   260
         Width           =   750
      End
      Begin VB.CommandButton Command4 
         Height          =   420
         Left            =   1060
         Picture         =   "DOCONEC.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Desconectar Documento"
         Top             =   735
         Width           =   750
      End
      Begin VB.CommandButton Command3 
         Height          =   420
         Left            =   210
         Picture         =   "DOCONEC.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Visualizar Documento Activo"
         Top             =   260
         Width           =   750
      End
      Begin VB.CommandButton Command1 
         Height          =   420
         Left            =   210
         Picture         =   "DOCONEC.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Conectar Documento"
         Top             =   735
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Documentos Vinculados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   2535
      Left            =   105
      TabIndex        =   25
      Top             =   1890
      Width           =   10620
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   210
         TabIndex        =   26
         Top             =   420
         Width           =   10200
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1590
      Left            =   105
      TabIndex        =   9
      Top             =   210
      Width           =   10620
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   6.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   9660
         Picture         =   "DOCONEC.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   390
         Width           =   750
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Nro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   730
         Left            =   8085
         TabIndex        =   10
         Top             =   315
         Width           =   1380
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   210
            TabIndex        =   11
            Top             =   315
            Width           =   960
         End
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Observac:"
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
         Left            =   105
         TabIndex        =   24
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Empleados:"
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
         Left            =   8505
         TabIndex        =   23
         Top             =   1230
         Width           =   1065
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ramo:"
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
         Left            =   105
         TabIndex        =   22
         Top             =   1230
         Width           =   960
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Princ.Producto:"
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
         Left            =   2940
         TabIndex        =   21
         Top             =   1230
         Width           =   1800
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contactos:"
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
         Left            =   105
         TabIndex        =   20
         Top             =   800
         Width           =   960
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa:"
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
         Left            =   105
         TabIndex        =   19
         Top             =   500
         Width           =   960
      End
      Begin VB.Label Label5 
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
         Height          =   285
         Left            =   1155
         TabIndex        =   18
         Top             =   420
         Width           =   6735
      End
      Begin VB.Label Label7 
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
         Height          =   285
         Left            =   1155
         TabIndex        =   17
         Top             =   735
         Width           =   3270
      End
      Begin VB.Label Label8 
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
         Height          =   285
         Left            =   4620
         TabIndex        =   16
         Top             =   735
         Width           =   3270
      End
      Begin VB.Label Label4 
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
         Height          =   285
         Left            =   1155
         TabIndex        =   15
         Top             =   1155
         Width           =   2115
      End
      Begin VB.Label Label13 
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
         Height          =   285
         Left            =   4830
         TabIndex        =   14
         Top             =   1155
         Width           =   3585
      End
      Begin VB.Label Label14 
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
         Height          =   285
         Left            =   9660
         TabIndex        =   13
         Top             =   1155
         Width           =   750
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Archivo de Conexión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3585
      Left            =   105
      TabIndex        =   0
      Top             =   4620
      Width           =   9465
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   3780
         TabIndex        =   6
         Top             =   420
         Width           =   5475
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   315
         TabIndex        =   2
         Top             =   840
         Width           =   3375
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   1470
         TabIndex        =   1
         Top             =   420
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   8
         Top             =   470
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   315
         TabIndex        =   7
         Top             =   2580
         Width           =   8205
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -210
         TabIndex        =   5
         Top             =   2350
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   315
         TabIndex        =   4
         Top             =   3150
         Width           =   8205
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Carpeta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -210
         TabIndex        =   3
         Top             =   2920
         Width           =   1065
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   9240
      Picture         =   "DOCONEC.frx":09F2
      Top             =   6405
      Width           =   2010
   End
End
Attribute VB_Name = "DOCONEC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   Call CONECTADOC
   Command1.Enabled = True
   '
   End Sub


Private Sub Command2_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command3_Click()
    '
    Command3.Enabled = False
    Call MUESTRADOC2
    Command3.Enabled = True
    '
End Sub

Private Sub Command4_Click()
   '
   Command4.Enabled = False
   XX$ = REGBLAN$("DOCONEC")
   Call REPLA(XX$, MKI$(Val(Label26)), 1, 2)
   Call REPLA(XX$, TRIM$(UCase$(Label1)), 3, 126)
   Call REPLA(XX$, TRIM$(UCase$(Label11)), 129, 128)
   '
   For U& = 1 To ULTREG&("DOCONEC")
     YY$ = REGLEIDO$("DOCONEC", U&)
     If Left$(YY$, 2) + UCase$(Mid$(YY$, 3, 126)) = Left$(XX$, 128) Then
       Call GRAREG("DOCONEC", String$(1024, 0), U&)
       GoTo 19
     End If
   Next U&
19 Call CIERRARCH("DOCONEC")
   '
   NPX& = Val(Label26)
   DOCONEC.List1.Clear
   NPXI% = NPX&
   RFF$ = RECFILT$("DOCONEC", 1, MKI$(NPXI%))
   '
   For U& = 1 To Len(RFF$) Step 4
     RENOTA& = CVS(Mid$(RFF$, U&, 4))
     XX$ = REGLEIDO$("DOCONEC", RENOTA&)
     YY$ = Space$(256)
     Call REPLA(YY$, Mid$(XX$, 3, 56), 1, 56)
     Call REPLA(YY$, Mid$(XX$, 129, 56), 61, 56)
     Call REPLA(YY$, TRIM$(Str$(RENOTA&)), 129, 6)
     DOCONEC.List1.AddItem YY$
   Next U&
   List1.Refresh
   DoEvents
   Command4.Enabled = True
   '
End Sub

Private Sub Command5_Click()
   '
   Command5.Enabled = False
   For KU% = 0 To 32767
     If EXISTE%(Label11 + "\Documento" + TRIM$(Str$(KU%)) + ".doc") < 1 Then GoTo 10
   Next KU%
   MsgBox "Imposible Abrir en Esta Carpeta"
   GoTo 99
   '
10 Label1 = "Documento" + TRIM$(Str$(KU%)) + ".doc"
   On Error Resume Next
   FileCopy LUDAT$ + "HOJENBLA.DOC", Label11 + "\Documento" + TRIM$(Str$(KU%)) + ".doc"
   If Err > 0 Then
     MsgBox "Error (" + TRIM$(Str$(Err)) + ") al Crear Nuevo Documento."
     On Error GoTo 0
     GoTo 99
   End If
   On Error GoTo 0
   '
   NUEVODOC% = 1
   DOCUORIGI$ = UCase$(TRIM$(Label11 + "\Documento" + TRIM$(Str$(KU%)) + ".doc"))
   Call MUESTRADOC2
   'MsgBox DOCUORIGI$ + "*" & "   " & UCase$(TRIM$(Label11 + "\Documento" + TRIM$(Str$(KU%)) + ".doc")) + "*"
   If DOCUORIGI$ <> UCase$(TRIM$(Label11 + "\" + Label1)) Then
     On Error Resume Next
     Kill DOCUORIGI$
     On Error GoTo 0
   End If
   NUEVODOC% = 0
   Call CONECTADOC
   '
99 Command5.Enabled = True
   '
End Sub

Private Sub Dir1_Change()
   Label11.Caption = Dir1.Path
   File1.Path = Dir1.Path
   Label1 = ""
End Sub

Private Sub Drive1_Change()
   On Error GoTo 20
   Dir1.Path = Drive1.Drive
   GoTo 99
   '
20 Resume 21
21 On Error GoTo 0
   Call MENSERR(24, "Unidad no Disponible.")
   Drive1.Drive = Left$(CurDir$, 2)
   '
99 On Error GoTo 0
End Sub

Private Sub file1_click()
  Label1 = File1.FileName
End Sub

Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    '
    EPAC$ = EMPREAC$
    If EPAC$ <> "" Then
      Caption = Caption + " - " + EMPREAC$
    End If
    '
    Label11.Caption = Dir1.Path
    File1.Path = Dir1.Path
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub


Private Sub List1_Click()
   RENOTA& = Val(Mid$(List1.TEXT, 129, 6))
   DRXX$ = Left$(LUDAT$, 2)
   XX$ = REGLEIDO$("DOCONEC", RENOTA&)
   DRYY$ = UCase$(Mid$(XX$, 129, 2))
   If DRYY$ >= "A:" And DRYY$ <= "C:" Then DRXX$ = DRYY$
   On Error Resume Next
   Drive1.Drive = DRXX$ ' Mid$(XX$, 129, 2)
   Dir1.Path = DRXX$ + TRIM$(Mid$(XX$, 131, 126))
   Label1 = TRIM$(Mid$(XX$, 3, 126))
   '
   For U& = 1 To File1.ListCount
     If UCase$(File1.List(U& - 1)) = UCase$(Label1) Then
       File1.Selected(U& - 1) = True
       Exit For
     End If
   Next U&
   On Error GoTo 0
   '
End Sub

Private Sub List1_DblClick()
   '
   RENOTA& = Val(Mid$(List1.TEXT, 129, 6))
   DRXX$ = Left$(LUDAT$, 2)
   XX$ = REGLEIDO$("DOCONEC", RENOTA&)
   DRYY$ = UCase$(Mid$(XX$, 129, 2))
   If DRYY$ >= "A:" And DRYY$ <= "C:" Then DRXX$ = DRYY$
   Label11 = DRXX$ + TRIM$(Mid$(XX$, 131, 128))
   Label1 = TRIM$(Mid$(XX$, 3, 126))
   '
   Call MUESTRADOC2
   '
End Sub

Sub MUESTRADOC2()
    '
    On Error Resume Next
    WORDDOCS.OLE1.SourceDoc = Label11 + "\" + Label1
    TIDOCU$ = "WORD"
    If UCase$(Right$(Label1, 4)) = ".EML" Then
        TIDOCU$ = "MAIL"
      ElseIf UCase$(Right$(Label1, 4)) = ".XLS" Then
        TIDOCU$ = "EXCL"
    End If
    '
    UNILOG$ = "": DOCUNE$ = WORDDOCS.OLE1.SourceDoc
    If Mid$(DOCUNE$, 2, 1) = ":" Then
      UNILOG$ = Left$(DOCUNE$, 2)
      DOCUNE$ = Mid$(DOCUNE$, 3)
    End If
    '
    WORDDOCS.OLE1.Action = 1
    '
    WORDDOCS.OLE1.Action = 6
    WORDDOCS.VScroll1.Min = 0
    WORDDOCS.VScroll1.Value = 0
    ABARRA = WORDDOCS.OLE1.Height - WORDDOCS.Picture1.Height
    If ABARRA < 1 Then ABARRA = 1
    WORDDOCS.VScroll1.Max = ABARRA
    WORDDOCS.VScroll1.SmallChange = ABARRA / 100
    WORDDOCS.VScroll1.LargeChange = ABARRA / 10
    WORDDOCS.OLE1.Top = 0
    WORDDOCS.OLE1.DoVerb (-1)
    If TIDOCU$ = "WORD" Then
        SendKeys "%{F10}", True  ' esto para documentos en word
      ElseIf TIDOCU$ = "MAIL" Then
        SendKeys Chr$(13) + "%{F10}", True ' esto para mails
      ElseIf TIDOCU$ = "EXCL" Then
        SendKeys Chr$(13) & "%V" & (DOWN8) & "P", True ' ESTO PARA EXCEL
    End If
    WORDDOCS.Show 1
    If Left$(DOCUACTIVO$, 2) = "\\" Then
      For U& = 1 To Drive1.ListCount
        ABC$ = Drive1.List(U& - 1)
        If Mid$(ABC$, 5, 2) = "\\" Then
          DEF$ = TRIM$(Mid$(ABC$, 5))
          DEF$ = Left$(DEF$, Len(DEF$) - 1)
          If Left$(DOCUACTIVO$, Len(DEF$)) = DEF$ Then
            ULOGI$ = Left$(ABC$, 2)
            DOCUACTIVO$ = ULOGI$ + Mid$(DOCUACTIVO$, 1 + Len(DEF$))
            GoTo 30
          End If
        End If
      Next U&
    End If
    '
30  RUTA$ = DOCUACTIVO$
    ARCHIX$ = ""
    For KKI% = Len(RUTA$) To 1 Step -1
      If Mid$(RUTA$, KKI%, 1) = "\" Then
        ARCHIX$ = Mid$(RUTA$, KKI% + 1)
        RUTA$ = Left$(RUTA$, KKI% - 1)
        GoTo 40
      End If
    Next KKI%
    '
40 Drive1.Drive = ULOGI$
   Dir1.Path = RUTA$
   Label1 = ARCHIX$
   For U& = 1 To File1.ListCount
     If TRIM$(UCase$(File1.List(U& - 1))) = TRIM$(UCase$(Label1)) Then
       File1.Selected(U& - 1) = True
       Exit For
     End If
   Next U&
   '
   On Error GoTo 0
   '
End Sub

Sub CONECTADOC()
   '
   XX$ = REGBLAN$("DOCONEC")
   Call REPLA(XX$, MKI$(Val(Label26)), 1, 2)
   Call REPLA(XX$, TRIM$(Label1), 3, 126)
   Call REPLA(XX$, TRIM$(Label11), 129, 128)
   Call REGAPP("DOCONEC", XX$)
   Call CIERRARCH("DOCONEC")
   '
   NPX& = Val(Label26)
   DOCONEC.List1.Clear
   NPXI% = NPX&
   RFF$ = RECFILT$("DOCONEC", 1, MKI$(NPXI%))
   '
   For U& = 1 To Len(RFF$) Step 4
     RENOTA& = CVS(Mid$(RFF$, U&, 4))
     XX$ = REGLEIDO$("DOCONEC", RENOTA&)
     YY$ = Space$(256)
     Call REPLA(YY$, Mid$(XX$, 3, 56), 1, 56)
     Call REPLA(YY$, Mid$(XX$, 129, 56), 61, 56)
     Call REPLA(YY$, TRIM$(Str$(RENOTA&)), 129, 6)
     DOCONEC.List1.AddItem YY$
   Next U&
   List1.Refresh
   DoEvents
   '
End Sub
