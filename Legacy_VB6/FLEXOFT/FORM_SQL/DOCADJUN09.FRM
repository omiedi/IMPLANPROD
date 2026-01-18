VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form DOCADJUN09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjuntar Archivos"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   9675
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Formato"
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
      Height          =   1170
      Left            =   105
      TabIndex        =   13
      Top             =   6510
      Width           =   9465
      Begin VB.CommandButton Command3 
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
         Left            =   8120
         TabIndex        =   22
         Top             =   315
         Width           =   170
      End
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
         Left            =   4725
         TabIndex        =   21
         Top             =   735
         Width           =   3555
      End
      Begin VB.CommandButton Command2 
         Height          =   740
         Left            =   8520
         Picture         =   "DOCADJUN09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Adjuntar / Procesar"
         Top             =   290
         Width           =   750
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
         Height          =   300
         Left            =   4725
         TabIndex        =   19
         Top             =   315
         Width           =   3555
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Zip Auto-Extraible a Carpeta:"
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
         Left            =   1890
         TabIndex        =   18
         Top             =   735
         Width           =   2850
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Zip-File (*.zip) ...................."
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
         Left            =   1890
         TabIndex        =   17
         Top             =   315
         Width           =   2955
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Abierto"
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
         Left            =   420
         TabIndex        =   16
         Top             =   420
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Lista de Archivos Adjuntos"
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
      Height          =   1800
      Left            =   105
      TabIndex        =   9
      Top             =   105
      Width           =   9465
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
         Height          =   1185
         Left            =   210
         TabIndex        =   10
         Top             =   420
         Width           =   9045
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Selección Archivo Adjunto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Left            =   105
      TabIndex        =   0
      Top             =   1995
      Width           =   9465
      Begin VB.CommandButton Command1 
         Height          =   420
         Left            =   8505
         Picture         =   "DOCADJUN09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Incluir Documento"
         Top             =   3255
         Width           =   750
      End
      Begin VB.CommandButton Command4 
         Height          =   420
         Left            =   8505
         Picture         =   "DOCADJUN09.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Excluir Documento Activo"
         Top             =   3780
         Width           =   750
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
         Left            =   315
         TabIndex        =   12
         Top             =   2730
         Width           =   3375
      End
      Begin VB.FileListBox File1 
         Height          =   2625
         Left            =   3780
         TabIndex        =   6
         Top             =   420
         Width           =   5475
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   315
         TabIndex        =   2
         Top             =   945
         Width           =   3375
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   1470
         TabIndex        =   1
         Top             =   420
         Width           =   2220
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   8400
         X2              =   8400
         Y1              =   3045
         Y2              =   4410
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro:"
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
         Left            =   -420
         TabIndex        =   11
         Top             =   2500
         Width           =   1065
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
         Top             =   3315
         Width           =   7995
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
         Top             =   3090
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
         Top             =   3885
         Width           =   7995
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
         Top             =   3660
         Width           =   1065
      End
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
End
Attribute VB_Name = "DOCADJUN09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
   '
   Command1.Enabled = False
   Call CONECTADOC
   Command1.Enabled = True
   '
End Sub


Private Sub Command2_Click()
   Command2.Enabled = False
   '
   If List1.ListCount < 1 Then
      Call MENSERR(24, "No se han Seleccionado Archivos a Adjuntar")
      GoTo 99
   End If
   '
   'If Option3.Value = True Then Call MENSERR(24, "Opcion Autoextraible no Funciona por Ahora")
   '
   If Option1.Value = False Then
      AZIP$ = TRIM$(Text2)
      If AZIP$ = "" Or InStr(AZIP$, ".") > 0 Then
        Call MENSERR(24, "Archivo no Válido")
        Text2.SetFocus
        GoTo 99
      End If
   End If
   '
   If Option1.Value = True Then
       Hide
       Exit Sub
   End If
   '
Call MENSERR(24, "Imposible Enviar Archivos Comprimidos.")
Hide

   If Option2.Value = True Or Option3.Value = True Then
      '
      ' aqui ver si se puede usar el generador de winzip
      AZIP$ = Text2
      AZIPSH$ = AZIP$
      While InStr(AZIPSH$, "\") > 0
          AZIPSH$ = Mid$(AZIPSH$, 2)
      Wend
      '
      On Error Resume Next
      Kill AZIP$ + ".ZIP"
      If Option3.Value = True Then
        Kill AZIP$ + ".EXE"
      End If
      On Error GoTo 0
      '
      On Error Resume Next
      RETURNVALUE = Shell("C:\Archivos de programa\WinZip\WINZIP32.EXE", 4)
      If Err Then
         On Error GoTo 0
         MsgBox "No se ha Encontrado 'C:\Archivos de programa\WinZip\WINZIP32.EXE'"
         GoTo 99
      End If
      '
      AppActivate RETURNVALUE   ' : AppActivate RETURNVALUE
      SendKeys "^N", True
      SendKeys AZIP$ + Chr$(13), True
      '
      TTXX$ = ""
      For U& = 1 To List1.ListCount
        TTXX$ = TTXX$ + Chr$(34) + ReplaSkey$(List1.List(U& - 1)) + Chr$(34) + Chr$(32)
      Next U&
      '
      SendKeys "%A", True
      SendKeys TRIM$(TTXX$) + Chr$(13), True
      '
      If Option3.Value = True Then
        SendKeys "%AK", True
        SendKeys "%D" + TRIM$(UCase$(Text3)) + Chr$(13), True
        On Error Resume Next
        AppActivate "WinZip Self-Extractor"
        On Error GoTo 0
        SendKeys "%N", True
        SendKeys "%C", True
      End If
      '
      AppActivate "WinZip - " + AZIPSH$ + ".ZIP"
      AppActivate "WinZip - " + AZIPSH$ + ".ZIP"
      'DoEvents
16 On Error Resume Next
'DoEvents
AppActivate "WINZIP"
If Err Then
    Resume 20
  Else
    SendKeys "%{F4}", True
    GoTo 16
End If
'
      
20    FIFUNA$ = AZIP$ + ".ZIP"
      'If Option3.Value = True Then
      '  FIFUNA$ = AZIP$ + ".EXE"
      'End If
      FISHNA$ = FIFUNA$
      While InStr(FISHNA$, "\") > 0
        FISHNA$ = Mid$(FISHNA$, 2)
      Wend
      If TRIM$(FISHNA$) <> "" Then
        MAILMESS07.List3.AddItem AJUSTI$(FISHNA$, 128) + FIFUNA$
      End If
      Command2.Enabled = True
      List1.Clear
      Hide
   End If
   '
99 Command2.Enabled = True
End Sub

Private Sub Command3_Click()
   '
   Command3.Enabled = False
   CommonDialog1.CancelError = True
   On Error GoTo 99
   '
10 CommonDialog1.DialogTitle = "Crear Archivo de Compresión"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.Filter = "Zip-Files (*.zip)|*.zip|"
   CommonDialog1.FilterIndex = 1
   CommonDialog1.ShowOpen
   '
   ' Presentar el nombre del archivo seleccionado
   TTXX$ = UCase$(CommonDialog1.FileName)
   PPXX% = InStr(TTXX$, ".")
   If PPXX% > 0 Then
     If Mid$(TTXX$, PPXX% + 1, 3) <> "ZIP" Then
       Call MENSERR(24, "Archivo no Válido - Debe ser '*.ZIP'")
       GoTo 10
     End If
     TTXX$ = Left$(TTXX$, PPXX% - 1)
   End If
   Text2 = TTXX$
   '
98 Resume 99
99 On Error GoTo 0
   Command3.Enabled = True
   '
End Sub

Private Sub Command4_Click()
   '
   Command4.Enabled = False
   For U& = 1 To List1.ListCount
     If List1.Selected(U& - 1) = True Then
       List1.RemoveItem (U& - 1)
       Exit For
     End If
   Next U&
   List1.Refresh
   DoEvents
   Command4.Enabled = True
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

Private Sub File1_DblClick()
   Call CONECTADOC
End Sub

Private Sub Form_Load()
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
   On Error Resume Next
   Hide
   On Error GoTo 0
End Sub


Private Sub Label11_Change()
  If Right$(Label11, 1) = "\" Then
    Label11 = Left$(Label11, Len(Label11) - 1)
  End If
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
   ARDOC$ = TRIM$(List1.TEXT)
'   RENOTA& = Val(Mid$(List1.TEXT, 129, 6))
'   DRXX$ = Left$(LUDAT$, 2)
'   XX$ = REGLEIDO$("DOCONEC", RENOTA&)
'   DRYY$ = UCase$(Mid$(XX$, 129, 2))
'   If DRYY$ >= "A:" And DRYY$ <= "C:" Then DRXX$ = DRYY$
'   Label11 = DRXX$ + TRIM$(Mid$(XX$, 131, 128))
'   Label1 = TRIM$(Mid$(XX$, 3, 126))
   '
   Call MUESTRADOC1(ARDOC$)
   '
End Sub

Sub MUESTRADOC1(ARDOC$)
    '
    Call ShellExecute(Me.hwnd, "Open", ARDOC$, "", "", 1)
    '
End Sub

Sub CONECTADOC()
   '
   If TRIM$(Label1) <> "" Then
     XX$ = TRIM$(Label11) + "\" + TRIM$(Label1)
     If EXISTE%(XX$) > 0 Then
          For U& = 1 To List1.ListCount
            If List1.List(U& - 1) = XX$ Then Exit Sub
          Next U&
          List1.AddItem XX$
          List1.Refresh
          DoEvents
        Else
          MsgBox "Imposible Conectar Archivo '" + XX$ + "' - Inexistente."
     End If
   End If
   '
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
     If Y > 0 Then
       YACT% = Int(5 * Y / List1.Height)
       On Error Resume Next
       List1.ListIndex = List1.TopIndex + YACT%
       TTXX$ = List1.TEXT + Chr$(13) + Chr$(10)
       TTXX$ = TTXX$ + "Actualizado: " + Format$(FileDateTime(List1.TEXT), "dd/mm/yyyy - hh:mm") + " hs" + Chr$(13) + Chr$(10)
       TTXX$ = TTXX$ + "Tamaño: " + TRIM$(Str$(FileLen(List1.TEXT))) + " bytes."
       If Err = 0 Then
         MsgBox TTXX$
       End If
       On Error GoTo 0
     End If
   End If
End Sub

Private Sub Text1_Change()
   File1.Pattern = Text1
End Sub
