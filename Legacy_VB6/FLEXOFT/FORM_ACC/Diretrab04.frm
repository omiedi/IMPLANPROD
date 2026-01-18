VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form DIRETRAB04 
   BackColor       =   &H00C0F0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carpetas de Trabajo"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   25
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0F0C0&
      Caption         =   "Servicios SQL-Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   105
      TabIndex        =   21
      Top             =   2415
      Width           =   6420
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   5670
         Picture         =   "Diretrab04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Configurar Conexión a SQL-SERVER"
         Top             =   315
         Width           =   600
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Servidor:"
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
         Left            =   -120
         TabIndex        =   33
         Top             =   360
         Width           =   2430
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Base de Datos Inicial:"
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
         Left            =   -120
         TabIndex        =   32
         Top             =   780
         Width           =   2430
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   2415
         TabIndex        =   24
         Top             =   735
         Width           =   3165
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
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
         Height          =   280
         Left            =   2415
         TabIndex        =   23
         Top             =   315
         Width           =   3165
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   8520
      Left            =   6660
      ScaleHeight     =   8460
      ScaleWidth      =   1845
      TabIndex        =   14
      Top             =   0
      Width           =   1905
      Begin VB.CommandButton Command2 
         Caption         =   "Rest."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "Diretrab04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2805
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "Diretrab04.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   870
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "Diretrab04.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2115
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "Diretrab04.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   180
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   15
         Top             =   3885
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   16
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "Diretrab04.frx":0D72
         Top             =   4230
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0F0C0&
      Caption         =   "Idioma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   105
      TabIndex        =   7
      Top             =   3675
      Width           =   6420
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0F0C0&
         Caption         =   "Portugués"
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
         Index           =   5
         Left            =   4620
         TabIndex        =   13
         Top             =   600
         Width           =   1590
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0F0C0&
         Caption         =   "Italiano"
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
         Index           =   4
         Left            =   4620
         TabIndex        =   12
         Top             =   315
         Width           =   1590
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0F0C0&
         Caption         =   "Francais"
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
         Index           =   3
         Left            =   2625
         TabIndex        =   11
         Top             =   600
         Width           =   1590
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0F0C0&
         Caption         =   "Deutsch"
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
         Index           =   2
         Left            =   2625
         TabIndex        =   10
         Top             =   315
         Width           =   1590
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0F0C0&
         Caption         =   "English"
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
         Index           =   1
         Left            =   840
         TabIndex        =   9
         Top             =   600
         Width           =   1590
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0F0C0&
         Caption         =   "Español"
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
         Index           =   0
         Left            =   840
         TabIndex        =   8
         Top             =   315
         Value           =   -1  'True
         Width           =   1590
      End
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   6090
      TabIndex        =   5
      Text            =   " "
      Top             =   315
      Width           =   435
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   2520
      TabIndex        =   1
      Text            =   " "
      Top             =   1995
      Width           =   4005
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   2520
      TabIndex        =   0
      Text            =   " "
      Top             =   1575
      Width           =   4005
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
      Height          =   280
      Left            =   2520
      TabIndex        =   4
      Text            =   " "
      Top             =   1155
      Width           =   4005
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
      Height          =   280
      Left            =   2520
      TabIndex        =   3
      Text            =   " "
      Top             =   735
      Width           =   4005
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
      Height          =   280
      Left            =   2520
      TabIndex        =   2
      Text            =   " "
      Top             =   315
      Width           =   435
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "Diretrab04.frx":2C94
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "Diretrab04.frx":2D0E
      Top             =   1200
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Utilitarios Flexoft:"
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
      Left            =   -45
      TabIndex        =   31
      Top             =   2040
      Width           =   2430
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Base de Datos Central:"
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
      Left            =   -45
      TabIndex        =   30
      Top             =   1620
      Width           =   2430
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Carpeta de Trabajo Local:"
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
      Left            =   -360
      TabIndex        =   29
      Top             =   1200
      Width           =   2745
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema de Aplicación:"
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
      Left            =   -45
      TabIndex        =   28
      Top             =   780
      Width           =   2430
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Disco de 'boot':"
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
      Left            =   -45
      TabIndex        =   27
      Top             =   360
      Width           =   2430
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Disco de trabajo:"
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
      Left            =   3570
      TabIndex        =   6
      Top             =   420
      Width           =   2430
   End
End
Attribute VB_Name = "DIRETRAB04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GRABA%
Dim DIAPLI As String
Dim LUCO1 As String
Dim LUDA1 As String
Dim LUBA1 As String
Dim LUDAX As String
Dim LUBAX As String
'
Private Sub Command1_Click()
    Call GRADIRE
    If GRABA% <> 1 Then Exit Sub
    'Call VERIFLIM(1)
    DIRETRAB04.Hide
End Sub

Private Sub Command2_Click()
    Call DIRETRA
End Sub

Private Sub Command3_Click()
    Call CIERRARCH("*.*")
    DIRETRAB04.Hide
End Sub

Private Sub Command4_Click()
    Call HELPONLINE("CARPETRA")
End Sub

Private Sub Command6_Click()

End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    LUDAT$ = Text6 + Text4 + "\"
    CONN_SQLS.Label1 = LUDAT$
    CONN_SQLS.Show 1
    Call CONNECSQL
    Command7.Enabled = True
End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
      Me.Image2.Visible = False
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Option1(0).BackColor = &HFCD7B6
      Me.Option1(1).BackColor = &HFCD7B6
      Me.Option1(2).BackColor = &HFCD7B6
      Me.Option1(3).BackColor = &HFCD7B6
      Me.Option1(4).BackColor = &HFCD7B6
      Me.Option1(5).BackColor = &HFCD7B6
      
      
    End If

    Caption = Caption + " - " + EMPREAC$
    Call DIRETRA
    Call APLICACIONSKINS(Me)
    On Error Resume Next
    Option1(LANGUA%).Value = True
    On Error GoTo 0
End Sub

Private Sub Image3_Click()

End Sub

Private Sub Text1_GotFocus()
    DIRETRAB04.Text4.SetFocus
End Sub
Private Sub Text2_GotFocus()
    DIRETRAB04.Text4.SetFocus
End Sub
Private Sub text3_gotfocus()
    DIRETRAB04.Text4.SetFocus
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text5.SetFocus
    End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1.SetFocus
    End If
End Sub

Sub DIRETRA()
    '
    DIRETRAB04.MousePointer = 11
    '
    LCB$ = LUBOOT$ + ":"
    DITRA$ = ULODATOS$
    DIAPLI$ = UCase$(Mid$(CurDir, 3))
    LUCO1$ = UCase$(TRIM$(LUCOM$))
    LUDA1$ = UCase$(TRIM$(LUDAT$))
    ULDA$ = ""
    If Mid$(LUDA1$, 2, 1) = ":" Then
       ULDA$ = Left$(LUDA1$, 2)
       LUDA1$ = Mid$(LUDA1$, 3)
    End If
    '
    LUBA1$ = UCase$(TRIM$(LUBAS$))
    If Mid$(LUBA1$, 2, 1) = ":" Then
       LUBA1$ = Mid$(LUBA1$, 3)
    End If
    '
    If Len(LUCO1$) > 1 Then
       If Right$(LUCO1$, 1) = "\" Then
          LUCO1$ = Left$(LUCO1$, Len(LUCO1$) - 1)
       End If
    End If
    '
    If Len(LUDA1$) > 1 Then
       If Right$(LUDA1$, 1) = "\" Then
          LUDA1$ = Left$(LUDA1$, Len(LUDA1$) - 1)
       End If
    End If
    '
    If LUDA1$ = "" Then LUDA1$ = DIAPLI$
    On Error GoTo CAPERR1
    NX% = FreeFile
    Open ULDA$ + LUDA1$ + "\MASTER.DAT" For Random Access Read Write Shared As #NX% Len = 128
    On Error GoTo 0
    Close #NX% = NX% = 0
    '
    If Len(LUBA1$) > 1 Then
       If Right$(LUBA1$, 1) = "\" Then
          LUBA1$ = Left$(LUBA1$, Len(LUBA1$) - 1)
       End If
    End If
    '
    If LUBA1$ = "" Then
       LUBA1$ = TRIM$(UCase$(Mid$(CurDir, 3)))
       LDDX$ = AJUSTI$(TRIM$(LUBA1$) + "\", 46)
    End If
    On Error GoTo CAPERR2
    NX% = FreeFile
    Open LUBA1$ + "\OBJNAME.DRV" For Random Access Read Write Shared As #NX% Len = 128
    On Error GoTo 0
    Close #NX%: NX% = 0
    '
    If EXISTE%(DIAPLI$ + "\LANGUAGE.DRV") > 0 Then
      On Error Resume Next
      LANGUA% = Val(LOADFILE$(DIAPLI$ + "\LANGUAGE.DRV"))
      On Error GoTo 0
    End If
    If LANGUA% < 0 Or LANGUA% > 5 Then LANGUA% = 0

    DIRETRAB04.Text1.TEXT = LCB$
    DIRETRAB04.Text2.TEXT = DIAPLI$
    DIRETRAB04.Text3.TEXT = LUCO1$
    DIRETRAB04.Text4.TEXT = LUDA1$
    DIRETRAB04.Text5.TEXT = LUBA1$
    DIRETRAB04.Text6.TEXT = DITRA$
      If ULDA$ <> "" Then
         DIRETRAB04.Text6.TEXT = ULDA$
      End If
    DIRETRAB04.Option1(LANGUA%).Value = True
    '
    Call CONNECSQL    ' intenta conectar a sqlserver
    '
    DIRETRAB04.MousePointer = 1
    On Error Resume Next
    DIRETRAB04.Text4.SetFocus
    On Error GoTo 0
    GoTo FINSUB
    
CAPERR1:
    LUDA1$ = TRIM$(UCase$(Mid$(CurDir, 3)))
    Resume Next
    '
CAPERR2:
    LUBA1$ = Mid$(CurDir, 3)
    Resume Next
    '
FINSUB:
End Sub

Sub GRADIRE()
    '
    GRABA% = 0
    LUDAX$ = TRIM$(DIRETRAB04.Text4.TEXT)
    On Error GoTo CAPERR3
    NX% = FreeFile
    Open LUDAX$ + "\MASTER.DAT" For Random Access Read Write Shared As #NX% Len = 128
    On Error GoTo 0
    Close #NX% = NX% = 0
    If LUDAX$ <> TRIM$(DIRETRAB04.Text4.TEXT) Then
        Call MENSERR(24, "Carpeta no Válida")
        DIRETRAB04.Text4.TEXT = LUDA1$
        DIRETRAB04.Text4.SetFocus
        Exit Sub
      Else
        LUDA1$ = LUDAX$
    End If
    '
    LUBAX$ = TRIM$(DIRETRAB04.Text5.TEXT)
    On Error GoTo CAPERR4
    NX% = FreeFile
    Open LUBAX$ + "\OBJNAME.DRV" For Random Access Read Write Shared As #NX% Len = 128
    On Error GoTo 0
    Close #NX% = NX% = 0
    If LUBAX$ <> TRIM$(DIRETRAB04.Text5.TEXT) Then
        Call MENSERR(24, "Carpeta no Válida")
        DIRETRAB04.Text5.TEXT = LUBA1$
        DIRETRAB04.Text5.SetFocus
        Exit Sub
      Else
        LUBA1$ = LUBAX$
    End If
    '
grabadi:
    '
    NX% = FreeFile
    Open "FLEXCFG.DAT" For Random Access Read Write Shared As #NX% Len = 128
    Dim FF0 As String * 128
    '
    Get #NX%, 1, FF0$
    XXX$ = FF0$
    Call REPLA(XXX$, DIAPLI$, 3, 46)
    Call REPLA(XXX$, LUDA1$, 51, 46)
    FF0$ = XXX$
    Put #NX%, 1, FF0$
    '
    LDDX$ = AJUSTI$(TRIM$(LUDA1$) + "\", 46)
    AA11$ = TRIM$(DIRETRAB04.Text6.TEXT)
    Call GRARAM(99, AA11$ + LDDX$)
    LUDAT$ = AA11$ + TRIM$(LDDX$)
    '
    Get #NX%, 3, FF0$
    XXX$ = FF0$
    Call REPLA(XXX$, LUBA1$, 1, 128)
    FF0$ = XXX$
    Put #NX%, 3, FF0$
    LDDX$ = AJUSTI$(TRIM$(LUBA1$) + "\", 46)
    Call GRARAM(145, LDDX$)
    LUBAS$ = TRIM$(LDDX$)
    If Mid$(LUBAS$, 2, 1) <> ":" Then
        ULOAC$ = Left$(CurDir, 2)
        LUBAS$ = ULOAC$ + LUBAS$
    End If
    '
    Close #NX%: NX% = 0
    '
    AA11$ = TRIM$(DIRETRAB04.Text6.TEXT)
    On Error Resume Next
    NN11% = FreeFile
    Open "C:\FLEXOFT\ULOGDAT.DRV" For Output As #NN11%
    Print #NN11%, AA11$
    Close #NN11%
    On Error GoTo 0
    '
    IDTX$ = IDENTER$
    ULXX$ = LOADFILE$("ULODATOS.DRV")
    For KKI% = 1 To Len(ULXX$) Step 16
      ULYY$ = Mid$(ULXX$, KKI%, 16)
      If TRIM$(Left$(ULYY$, 14)) = IDTX$ Then
        Mid$(ULYY$, 15, 2) = AA11$
        Mid$(ULXX$, KKI%, 16) = ULYY$
        GoTo 100
      End If
    Next KKI%
    ULXX$ = ULXX$ + AJUSTI$(IDTX$, 14) + AA11$
100 Call SAVEFILE("ULODATOS.DRV", ULXX$)
    '
    '
    LANGUA% = 0
    For KKKI% = 0 To 5
      If DIRETRAB04.Option1(KKKI%).Value = True Then
        LANGUA% = KKKI%
      End If
    Next KKKI%
    Call SAVEFILE(DIAPLI$ + "\LANGUAGE.DRV", TRIM$(Str$(LANGUA%)))
    Call GRARAM(305, TRIM$(Str$(LANGUA%)))
    Call BAJRAM
    '
    GRABA% = 1
  Exit Sub
  '
CAPERR3:
    LUDAX$ = LUDA1$
    Resume Next
    '
CAPERR4:
    LUBAX$ = LUBA1$
    Resume Next
    '
End Sub
'
Sub CONNECSQL()
    '
    Label11 = ""
    Label12 = ""
    LUDATX$ = Text6 + Text4 + "\"
20  CONNSTRI$ = TRIM$(LOADFILE$(LUDATX$ + "CONNSQLS.DRV"))
    If CONNSTRI$ <> "" Then
       Set MiConn = New ADODB.Connection
       MiConn.ConnectionString = CONNSTRI$
       On Error Resume Next
       MiConn.Open
       '
       If Err Then
           Label11 = ""
           Label12 = ""
           Call MENSERR(24, "Imposible Conectar con\" + CONNSTRI$)
           MiConn.Close
           Exit Sub
         Else
           TXT$ = CONNSTRI$
           PINI% = InStr(TXT$, "Data Source=") + 12
           TXT1$ = Mid$(TXT$, PINI%)
           PFIN% = InStr(TXT1$, ";")
           Label11 = Left$(TXT1$, PFIN% - 1)
           PINI% = InStr(TXT$, "Initial Catalog=") + 16
           TXT1$ = Mid$(TXT$, PINI%)
           PFIN% = InStr(TXT1$, ";")
           Label12 = Left$(TXT1$, PFIN% - 1)
       End If
    End If
    '
End Sub

