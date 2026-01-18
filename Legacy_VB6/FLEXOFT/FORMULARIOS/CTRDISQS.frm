VERSION 5.00
Begin VB.Form CTRDISQS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Diskettes (Disquetera 'A')"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8055
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
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Format"
      Height          =   800
      Left            =   7140
      Picture         =   "CTRDISQS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Formatear Diskette Colocado"
      Top             =   2730
      Width           =   855
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Control"
      Height          =   800
      Left            =   7140
      Picture         =   "CTRDISQS.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Controlar un Diskette"
      Top             =   1050
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   800
      Left            =   7140
      Picture         =   "CTRDISQS.frx":2AAC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Limpiar"
      Height          =   800
      Left            =   7140
      Picture         =   "CTRDISQS.frx":2BF6
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Borrar todos los Archivos del Diskette"
      Top             =   1890
      Width           =   855
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
      Height          =   2760
      Left            =   210
      TabIndex        =   1
      Top             =   735
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
      Left            =   420
      TabIndex        =   0
      Top             =   1260
      Width           =   2430
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
      TabIndex        =   7
      Top             =   210
      Width           =   1905
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Número de Serie del Volumen:"
      Height          =   225
      Left            =   1785
      TabIndex        =   6
      Top             =   315
      Width           =   3060
   End
End
Attribute VB_Name = "CTRDISQS"
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
     File1.Refresh
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
   If DISKREADY% > 0 Then
     List2.Clear
     Label2 = SERIVOL$("A:")
     For U& = 1 To File1.ListCount
       ANAME$ = "A:\" + File1.List(U& - 1)
       BNAME$ = Mid$(ANAME$, 4) + Space$(128)
       FARCH$ = Format(FileDateTime(ANAME$), "dd/mm/yy   hh:mm")
       Call REPLA(BNAME$, FARCH$, 33, 16)
       TARCH$ = AJUSTD$(Format$(FileLen(ANAME$) / 1000, "######0.0"), 9)
       Call REPLA(BNAME$, TARCH$ + " Kb", 49, 12)
       List2.AddItem BNAME$
     Next U&
   End If
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
