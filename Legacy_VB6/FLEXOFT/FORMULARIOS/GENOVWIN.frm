VERSION 5.00
Begin VB.Form GENOVWIN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Diskettes (Disquetera 'A')"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8175
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
   ScaleHeight     =   3705
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   7140
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2800
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   210
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   210
      Width           =   6735
   End
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
      Height          =   2760
      Left            =   210
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   2625
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.CommandButton Command29 
      Caption         =   "O.K."
      Height          =   800
      Left            =   7140
      Picture         =   "GENOVWIN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Genera Juego de Diskettes"
      Top             =   1060
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   800
      Left            =   7140
      Picture         =   "GENOVWIN.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   800
      Left            =   7140
      Picture         =   "GENOVWIN.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancela la Selección"
      Top             =   1910
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   6615
      Picture         =   "GENOVWIN.frx":075E
      Top             =   3030
      Width           =   2010
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
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
      Height          =   330
      Left            =   5040
      TabIndex        =   6
      Top             =   210
      Width           =   1905
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Número de Serie del Volumen:"
      Height          =   225
      Left            =   1785
      TabIndex        =   5
      Top             =   315
      Width           =   3060
   End
End
Attribute VB_Name = "GENOVWIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Call CIERRARCH("*.*")
  Call BAJALDISCO
  Hide
End Sub

Function DISKREADY%()
    '
    On Error Resume Next
    File1.Path = "A:\"
    File1.Pattern = "*.*"
    File1.Refresh
    DISKREADY% = 1
    If Err Then
       MsgBox "Coloque Diskette"
       DISKREADY% = 0
       List2.Clear
       Label2 = ""
    End If
    On Error GoTo 0
    '
End Function

Private Sub Command2_Click()
   Command2.Enabled = False
   If DISKREADY% > 0 Then
     On Error Resume Next
     N1% = FreeFile
     Open "A:\FLEXKEY.KEY" For Random Access Read Write Shared As #N1% Len = 127
     If Err Then
       MsgBox "Diskette Protegido - Quite Protección."
       GoTo 99
     End If
     '
     Close #N1%
     On Error GoTo 0
     Kill "A:\*.*"
     If File1.ListCount > 0 Then
       MsgBox "Operación no Completada."
     End If
     List2.Clear
     '
   End If
99 On Error GoTo 0
   Close #N1%
   Command2.Enabled = True
End Sub

Private Sub Command29_Click()
   Command29.Enabled = False
   Screen.MousePointer = 11
   Kill ULODATOS$ + "\ANOVE\*.*"
   ARCOPI$ = "OBJNAME.DRV": GoSub COPIARCH
   ARCOPI$ = "OBJDETA.DRV": GoSub COPIARCH
   ARCOPI$ = "FLEXPRFS.FLX": GoSub COPIARCH
   ARCOPI$ = "FLEXPRF.ZIP": GoSub COPIARCH
   ARCOPI$ = "FLEXRFS.ZIP": GoSub COPIARCH
   '
   For U& = 1 To List2.ListCount
     If List2.Selected(U& - 1) Then
       ARCOPI$ = TRIM$(Left$(List2.List(U& - 1), 16))
       GoSub COPIARCH
     End If
   Next U&
   Screen.MousePointer = 1
   '
   NX% = FILEOPEN%("C:\FLEXOFT\DOSCOMM.BAT", 2, 0)
   Print #NX%, "@ECHO OFF"
   Print #NX%, "\FLEXOWIN\ARJ A A:\NOVEWIN " + ULODATOS$ + "\ANOVE\*.* -VA"
   Print #NX%, "EXIT"
   Close #NX%: NX% = 0
   '
   Shell "C:\FLEXOFT\DOSEXEC.PIF", vbNormalFocus
   '
   Command29.Enabled = True
   Exit Sub
   '
COPIARCH:
   If EXISTE%(ULODATOS$ + "\FLEXOWIN\" + ARCOPI$) > 0 Then
     FileCopy ULODATOS$ + "\FLEXOWIN\" + ARCOPI$, ULODATOS$ + "\ANOVE\" + ARCOPI$
   End If
Return
'
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   If DISKREADY% > 0 Then
     On Error Resume Next
     N1% = FreeFile
     Open "A:\FLEXKEY.KEY" For Random Access Read Write Shared As #N1% Len = 127
     If Err Then
       MsgBox "Diskette Protegido - Quite Protección."
       GoTo 99
     End If
     '
     Close #N1%
     On Error GoTo 0
     '
     NX% = FILEOPEN%("C:\FLEXOFT\DOSCOMM.BAT", 2, 0)
     Print #NX%, "@ECHO OFF"
     Print #NX%, "FORMAT A:"
     Print #NX%, "EXIT"
     Close #NX%: NX% = 0
     '
     Shell "C:\FLEXOFT\DOSEXEC.PIF", vbNormalFocus
     '
     List2.Clear: Label2 = ""
     '
   End If
99 On Error GoTo 0
   Close #N1%
   Command3.Enabled = True
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   Refresh
   DoEvents
   File1.Path = ULODATOS$ + "\FLEXOWIN"
   List1.Clear
   '
   File1.Pattern = "*.EXE": FLIMIT# = 0: GoSub LEEDATAR
   File1.Pattern = "*.PRF": FLIMIT# = FileDateTime(ULODATOS$ + "\FLEXOWIN\FLEXPRFS.FLX"): GoSub LEEDATAR
   '
   List2.Clear
   For U& = List1.ListCount To 1 Step -1
     BNAME$ = Mid$(List1.List(U& - 1), 21)
     List2.AddItem BNAME$
   Next U&
   Screen.MousePointer = 1
   Exit Sub
   '
LEEDATAR:
'
For U& = 1 To File1.ListCount
   ANAME$ = ULODATOS$ + "\FLEXOWIN\" + UCase$(File1.List(U& - 1))
   If FileDateTime(ANAME$) >= FLIMIT# Then
      BNAME$ = Space$(20) + Mid$(ANAME$, 13) + Space$(108)
      FARCH$ = Format(FileDateTime(ANAME$), "#######0.000000")
      While Len(FARCH$) < 16: FARCH$ = "0" + FARCH$: Wend
      Call REPLA(BNAME$, FARCH$, 1, 16)
      FARCH$ = Format(FileDateTime(ANAME$), "dd/mm/yy   hh:mm")
      TARCH$ = AJUSTD$(Format$(FileLen(ANAME$) / 1000, "######0.0"), 9)
      Call REPLA(BNAME$, FARCH$, 40, 16)
      Call REPLA(BNAME$, AJUSTD$(TARCH$, 10) + " kB", 58, 13)
      List1.AddItem BNAME$
      DoEvents
   End If
Next U&
'
Return

End Sub
