VERSION 5.00
Begin VB.Form INI_CD 
   AutoRedraw      =   -1  'True
   Caption         =   "FLEXOFT - Sistema Demo Activa"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   5325
   Icon            =   "INI_CD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   1170
      Left            =   0
      ScaleHeight     =   1110
      ScaleWidth      =   5310
      TabIndex        =   2
      Top             =   3150
      Width           =   5370
      Begin VB.CommandButton Command1 
         Height          =   750
         Left            =   4305
         Picture         =   "INI_CD.frx":1D2A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   210
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "FLEXOFT Modular Integrado"
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
         Left            =   1050
         TabIndex        =   5
         Top             =   700
         Width           =   3060
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sistema Power-Up Gestión"
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
         Left            =   630
         TabIndex        =   4
         Top             =   450
         Value           =   -1  'True
         Width           =   3060
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sistema Power-Up Industrial"
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
         Left            =   210
         TabIndex        =   3
         Top             =   210
         Width           =   3060
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   2  'Horizontal Line
      Height          =   3210
      Left            =   0
      Picture         =   "INI_CD.frx":2034
      ScaleHeight     =   3210
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   2725.888
      TabIndex        =   0
      Top             =   0
      Width           =   5370
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Soluciones Informáticas Integrales "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   2835
         TabIndex        =   1
         Top             =   105
         Width           =   2415
      End
   End
End
Attribute VB_Name = "INI_CD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Screen.MousePointer = 11
   If Option1.Value = True Then
       If EXISTE%("FX_PROGRAMS\FLEXO00.EXE") > 0 Then
         ChDir "\FX_Programs"
         Shell "FLEXO00.EXE", vbNormalFocus
         Close
         End
       End If
       If EXISTE%("FX_PROGRAMS\FLEXO04.EXE") > 0 Then
         ChDir "\FX_Programs"
         Shell "FLEXO04.EXE", vbNormalFocus
         Close
         End
       End If
     ElseIf Option2.Value = True Then
       If EXISTE%("FX_PROGRAMS\FLGEST00.EXE") > 0 Then
         ChDir "\FX_Programs"
         Shell "FLGEST00.EXE", vbNormalFocus
         Close
         End
       End If
       If EXISTE%("FX_PROGRAMS\FLGEST04.EXE") > 0 Then
         ChDir "\FX_Programs"
         Shell "FLGEST04.EXE", vbNormalFocus
         Close
         End
       End If
     ElseIf Option3.Value = True Then
       If EXISTE%("FX_PROGRAMS\FLEXOFT.EXE") > 0 Then
         ChDir "\FX_Programs"
         Shell "FLEXOFT.EXE", vbNormalFocus
         Close
         End
       End If
   End If
End Sub

Private Sub Form_Load()
   Show
   Refresh
End Sub

Private Sub Option1_DblClick()
   If EXISTE%("FX_PROGRAMS\FLEXO00.EXE") > 0 Then
     ChDir "\FX_Programs"
     Shell "FLEXO00.EXE", vbNormalFocus
     Close: End
   End If
   If EXISTE%("FX_PROGRAMS\FLEXO04.EXE") > 0 Then
     ChDir "\FX_Programs"
     Shell "FLEXO04.EXE", vbNormalFocus
     Close: End
   End If
End Sub

Private Sub Option2_DblClick()
   If EXISTE%("FX_PROGRAMS\FLGEST00.EXE") > 0 Then
     ChDir "\FX_Programs"
     Shell "FLGEST00.EXE", vbNormalFocus
     Close: End
   End If
   If EXISTE%("FX_PROGRAMS\FLGEST04.EXE") > 0 Then
     ChDir "\FX_Programs"
     Shell "FLGEST04.EXE", vbNormalFocus
     Close: End
   End If
End Sub

Private Sub Option3_DblClick()
   If EXISTE%("FX_PROGRAMS\FLEXOFT.EXE") > 0 Then
     ChDir "\FX_Programs"
     Shell "FLEXOFT.EXE", vbNormalFocus
     Close: End
   End If
End Sub

Function EXISTE%(ARCHX$)
    '
    ARXX$ = Trim$(ARCHX$)
    If InStr(ARXX$, ".") < 1 Then ARXX$ = ARXX$ + ".DAT"
    If Right$(ARXX$, 1) = "." Then ARXX$ = Left$(ARXX$, Len(ARXX$) - 1)
    '
    EXI% = 0
    If Trim$(ARXX$) <> "" Then
      On Error Resume Next
      EXISZ$ = Trim$(Dir(ARXX$))
      If Err < 1 Then
        If EXISZ$ <> "" Then
          EXI% = 1
        End If
      End If
    End If
    '
    EXISTE% = EXI%
    On Error GoTo 0
    '
End Function

