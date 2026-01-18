VERSION 5.00
Begin VB.Form FLEXLOGI 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema Preconfigurado FLEXOFT (r)"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3840
   Icon            =   "FLEXLOGI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File2 
      Height          =   870
      Left            =   105
      Pattern         =   "*.EXE"
      TabIndex        =   10
      Top             =   5565
      Visible         =   0   'False
      Width           =   4530
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   105
      TabIndex        =   9
      Top             =   3150
      Width           =   4530
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Acceso al Sistema"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2640
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3585
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2205
         TabIndex        =   2
         Top             =   2020
         Width           =   1170
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2205
         TabIndex        =   1
         Top             =   600
         Width           =   1170
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
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
         Height          =   360
         Left            =   210
         TabIndex        =   6
         Top             =   1155
         Width           =   3165
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clave de"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1350
         TabIndex        =   5
         Top             =   1995
         Width           =   765
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1395
         TabIndex        =   4
         Top             =   735
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   585
         Left            =   315
         Picture         =   "FLEXLOGI.frx":030A
         Stretch         =   -1  'True
         Top             =   420
         Width           =   690
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acceso:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1425
         TabIndex        =   3
         Top             =   2205
         Width           =   705
      End
      Begin VB.Image Image2 
         Height          =   525
         Left            =   210
         Picture         =   "FLEXLOGI.frx":0614
         Stretch         =   -1  'True
         Top             =   1890
         Width           =   585
      End
      Begin VB.Image Image4 
         Height          =   525
         Left            =   210
         Picture         =   "FLEXLOGI.frx":091E
         Stretch         =   -1  'True
         Top             =   1890
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Image Image5 
         Height          =   540
         Left            =   210
         Picture         =   "FLEXLOGI.frx":0C28
         Stretch         =   -1  'True
         Top             =   1890
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   2520
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label99 
      Caption         =   "Microsoft Vísual Basic 5.O ................................"
      Height          =   225
      Left            =   525
      TabIndex        =   8
      Top             =   315
      Visible         =   0   'False
      Width           =   2745
   End
End
Attribute VB_Name = "FLEXLOGI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FF0 As String * 128
Dim CMD$

Dim CLA$, CLB$, CLAV$
Dim CLAVE$, CONTRA$
Dim QWERTY%
Dim NUSER%
Dim USERNA$
Dim EMPREAC$
Dim MULTIEM%
Dim YACAR%
'
Private Sub Form_Load()
    '
    Text1.Enabled = True
    Text3.Enabled = False
    '
    Label2 = 30: DoEvents
    CMD$ = Command$
    '
    Screen.MousePointer = 1
    On Error Resume Next
    Text1.SetFocus
    '
End Sub
'
Private Sub Text1_KeyPress(KeyAscii As Integer)
   '
   A% = KeyAscii: KeyAscii = 0
   If A% = 13 Then
        Call BUSER
     ElseIf A% = 27 Then
        Close: End
     ElseIf A% < 48 Then
        Exit Sub
     Else
        If A% > 96 Then A% = A% - 32
        If Len(Text1.TEXT) < 12 Then
           CLB$ = CLB$ + "*"
           Text1.TEXT = CLB$
           CLA$ = CLA$ + Chr$(A%)
           CLAVE$ = Left$(CLA$ + Space$(12), 12)
        End If
   End If
   '
End Sub

Sub BUSER()
    '
    If UCase$(CLA$) = "ERNIE" Then
         NUSER% = 1
         Label1 = "ERNESTO P. BAUER"
       ElseIf UCase$(CLA$) = "FLOR" Then
         NUSER% = 2
         Label1 = "FLORENCIA RICCI"
       ElseIf UCase$(CLA$) = "ADRIAN" Then
         NUSER% = 3
         Label1 = "ADRIAN DA ENCARNACAO"
       ElseIf UCase$(CLA$) = "FLAVIO" Then
         NUSER% = 4
         Label1 = "FLAVIO GRASSO"
    End If
    If NUSER% > 0 Then
       Label1.Refresh
       Text3.Enabled = True
       On Error Resume Next
       Text3.SetFocus
       On Error GoTo 0
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   '
   If Text3.TEXT = "" Then
      CLA$ = "": CLB$ = ""
   End If
   A% = KeyAscii
   KeyAscii = 0
   Text3.TEXT = CLB$
   If A% = 13 Then
      CONTRA$ = Left$(CLA$ + Space$(12), 12)
      Call BUCLA
      Exit Sub
   End If
   '
   If A% = 27 Then Close: End
   If A% < 48 Then Exit Sub
   If A% > 96 Then A% = A% - 32
   If Len(Text3.TEXT) < 12 Then
      CLB$ = CLB$ + "*"
      Text3.TEXT = CLB$
      CLA$ = CLA$ + Chr$(A%)
      CONTRA$ = Left$(CLA$ + Space$(12), 12)
   End If
          '
End Sub

Sub BUCLA()
   '
   Static CAIN%
   '
   If NUSER% = 1 Then CLAV$ = AJUSTI$("MONICA50", 12)
   If NUSER% = 2 Then CLAV$ = AJUSTI$("FLOR82", 12)
   If NUSER% = 3 Then CLAV$ = AJUSTI$("ADRIAN", 12)
   If NUSER% = 4 Then CLAV$ = AJUSTI$("1234", 12)
   If Left$(CLAV$, 12) = Left$(CONTRA$, 12) Then
       USNBR% = NUSER%
       USERNAME$ = Label1
       If USNBR% = 1 Then USNBR% = 101
       If USNBR% = 2 Then USNBR% = 102
       If USNBR% = 4 Then USNBR% = 104
       Hide
     Else
       CAIN% = CAIN% + 1
       If CAIN% > 3 Then
          MsgBox "Superado Limite Permitido Intentos de Acceso !!!"
          Close: End
       End If
       '
       MsgBox "Contraseña Incorrecta"
       Text3 = ""
       Text3.SetFocus
   End If
   '
End Sub
'
