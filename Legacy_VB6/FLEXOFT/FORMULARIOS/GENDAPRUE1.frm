VERSION 5.00
Begin VB.Form GENDAPRUE1 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Instalador de Controladores de Proceso"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   3360
      TabIndex        =   7
      Top             =   105
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   5145
      Top             =   105
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   360
      ScaleHeight     =   75
      ScaleWidth      =   4875
      TabIndex        =   4
      Top             =   1560
      Width           =   4935
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
   Begin VB.Label Label4 
      Height          =   330
      Left            =   1260
      TabIndex        =   8
      Top             =   1890
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   315
      Picture         =   "GENDAPRUE1.frx":0000
      Top             =   735
      Width           =   480
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   315
      TabIndex        =   6
      Top             =   735
      Width           =   4950
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   2120
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Instalando Controlador ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   280
      Width           =   4935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   1
      Top             =   1725
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1725
      Width           =   495
   End
End
Attribute VB_Name = "GENDAPRUE1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
   '
   ARCHIN$ = Trim$(Label18.Caption)
   '
   If Trim$(ARCHIN$) = "Datos de Prueba" Then
     File1.Path = "\FX_DATPRUE"
     File1.Pattern = "*.*"
     FIN& = File1.ListCount
     For II& = 1 To FIN&
       Call TRACE(20, II&, FIN&)
       DoEvents
       ORIGI$ = "\FX_DATPRUE\" + Trim$(UCase$(File1.List(II& - 1)))
       DESTI$ = Label4 + "\" + Trim$(UCase$(File1.List(II& - 1)))
       On Error Resume Next
       FileCopy ORIGI$, DESTI$
     Next II&
     '
     On Error Resume Next
     MkDir Label4 + "\BDACCESS"
     File1.Path = "\FX_DATPRUE\BDACCESS"
     File1.Pattern = "*.*"
     FIN& = File1.ListCount
     For II& = 1 To FIN&
       Call TRACE(20, II&, FIN&)
       DoEvents
       ORIGI$ = "\FX_DATPRUE\BDACCESS\" + Trim$(UCase$(File1.List(II& - 1)))
       DESTI$ = Label4 + "\BDACCESS\" + Trim$(UCase$(File1.List(II& - 1)))
       On Error Resume Next
       FileCopy ORIGI$, DESTI$
     Next II&
     '
     Screen.MousePointer = 1
     On Error GoTo 0
   End If
   '
90 On Error Resume Next
   Hide
   On Error GoTo 0
   'Call FRESHALL
   DoEvents
   '
End Sub


Private Sub Timer1_Timer()
   If Image1.Visible = True Then
       Image1.Visible = False
     Else
       Image1.Visible = True
   End If
   DoEvents
End Sub

Sub TRACE(RG%, I&, FIN&)
    '
    Static FINA&, COEF, TANT, TANTUI, VISI%, VERTI%, TINI, IFORMU&, IBARRA&, IMARCO&
    '
    If FIN& <> FINA& Then
        FINA& = FIN&
        VERTI% = 0
        If FIN& > 0 Then
            '
            IFORMU& = -1: IBARRA& = -1: IMARCO& = (-1)
            For TTT& = 1 To Forms.Count
               If Forms(TTT& - 1).Visible = True Then
                  If Forms(TTT& - 1).Enabled = True Then
                    For U& = 0 To Forms(TTT& - 1).Controls.Count - 1
                        NM$ = Forms(TTT& - 1).Controls(U&).Name
                        If NM$ = "BARRATRAZA" Then
                            IFORMU& = TTT& - 1
                            IBARRA& = U&
                          ElseIf NM$ = "MARCOBARRA" Then
                            IFORMU& = TTT& - 1
                            IMARCO& = U&
                        End If
                    Next U&
                  End If
               End If
               If IFORMU& >= 0 Then GoTo 10
            Next TTT&
            Exit Sub
            '
10          If IFORMU& < 0 Or IMARCO& < 0 Or IBARRA& < 0 Then
                Exit Sub
            End If
            '
            COEF = Forms(IFORMU&).Controls(IMARCO&).Width / FIN&
            If Forms(IFORMU&).Controls(IMARCO&).Height > Forms(IFORMU&).Controls(IMARCO&).Width Then
               VERTI% = 1
               COEF = Forms(IFORMU&).Controls(IMARCO&).Height / FIN&
            End If
            TINI = Timer
        End If
        On Error Resume Next
        Forms(IFORMU&).Controls(IBARRA&).Left = 0
        Forms(IFORMU&).Controls(IBARRA&).Top = 0
        Forms(IFORMU&).Controls(IBARRA&).Width = Forms(IFORMU&).Controls(IMARCO&).Width
        Forms(IFORMU&).Controls(IBARRA&).Height = Forms(IFORMU&).Controls(IMARCO&).Height
        On Error GoTo 0
    End If
    '
    If Abs(Timer - TANTUI) < 0.25 Then
       Exit Sub
    End If
    TANTUI = Timer
    '
    If RG% = 1 Then
      If Int(Timer) = TANT Then
        If I& < FIN& Then
          Exit Sub
        End If
      End If
      TANT = Int(Timer)
    End If
    '
    On Error Resume Next
    If IFORMU& >= 0 Then
       If IMARCO& >= 0 Then
          If IBARRA& >= 0 Then
             If VERTI% = 0 Then
                  Forms(IFORMU&).Controls(IBARRA&).Left = COEF * I&
                Else
                  Forms(IFORMU&).Controls(IBARRA&).Top = COEF * (-I&)
             End If
             '
             If I& >= FIN& Then
                FINA& = 0
                Forms(IFORMU&).Controls(IBARRA&).Left = 0
                Forms(IFORMU&).Controls(IBARRA&).Top = 0
             End If
             '
             Forms(IFORMU&).Controls(IMARCO&).Refresh
             Forms(IFORMU&).Controls(IBARRA&).Refresh
          End If
       End If
    End If
    On Error GoTo 0
    '
End Sub
'

