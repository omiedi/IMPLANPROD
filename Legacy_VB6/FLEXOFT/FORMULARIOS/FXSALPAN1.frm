VERSION 5.00
Begin VB.Form FXSALPAN1 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema Preconfigurado FLEXOFT - Demostración Interactiva"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FXSALPAN1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   11910
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   420
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Interval        =   80
      Left            =   840
      Top             =   0
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
      Height          =   2430
      Left            =   5460
      TabIndex        =   11
      Top             =   630
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Terminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10710
      Picture         =   "FXSALPAN1.frx":1D2A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   315
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   790
      Left            =   975
      TabIndex        =   2
      Top             =   450
      Visible         =   0   'False
      Width           =   2790
      Begin VB.Image Image9 
         BorderStyle     =   1  'Fixed Single
         Height          =   915
         Left            =   -210
         Picture         =   "FXSALPAN1.frx":1E74
         Top             =   -75
         Width           =   3120
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "al Servicio de la Producción "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1275
      Left            =   525
      TabIndex        =   1
      Top             =   3045
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Un Universo de Herramientas Informáticas "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1590
      Left            =   630
      TabIndex        =   0
      Top             =   1470
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.Label MODULO 
      BackStyle       =   0  'Transparent
      Caption         =   "- Planeamiento Financiero"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Index           =   5
      Left            =   5565
      TabIndex        =   8
      Top             =   6720
      Visible         =   0   'False
      Width           =   4950
   End
   Begin VB.Label MODULO 
      BackStyle       =   0  'Transparent
      Caption         =   "- Presupuestación de Producto"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Index           =   4
      Left            =   5565
      TabIndex        =   7
      Top             =   6090
      Visible         =   0   'False
      Width           =   4950
   End
   Begin VB.Label MODULO 
      BackStyle       =   0  'Transparent
      Caption         =   "- Rastreabilidad según ISO 9000"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Index           =   3
      Left            =   5565
      TabIndex        =   6
      Top             =   5460
      Visible         =   0   'False
      Width           =   6420
   End
   Begin VB.Label MODULO 
      BackStyle       =   0  'Transparent
      Caption         =   "- Carga de Máquinas (MRP-II)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Index           =   2
      Left            =   5565
      TabIndex        =   5
      Top             =   4830
      Visible         =   0   'False
      Width           =   6420
   End
   Begin VB.Label MODULO 
      BackStyle       =   0  'Transparent
      Caption         =   "- Programación de Materiales (MRP-I) "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Index           =   1
      Left            =   5565
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   6945
   End
   Begin VB.Label MODULO 
      BackStyle       =   0  'Transparent
      Caption         =   "- Control de Stocks"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Index           =   0
      Left            =   5565
      TabIndex        =   3
      Top             =   3570
      Visible         =   0   'False
      Width           =   4950
   End
   Begin VB.Image Image11 
      Height          =   855
      Left            =   1680
      Top             =   7245
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image Image10 
      Height          =   750
      Left            =   3780
      Stretch         =   -1  'True
      Top             =   7350
      Width           =   1590
   End
   Begin VB.Label MODULO 
      BackStyle       =   0  'Transparent
      Caption         =   "y mucho más ..."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Index           =   6
      Left            =   9135
      TabIndex        =   9
      Top             =   7560
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   630
      Picture         =   "FXSALPAN1.frx":A6FA
      Top             =   5355
      Width           =   765
   End
   Begin VB.Image Image2 
      Height          =   765
      Left            =   1410
      Picture         =   "FXSALPAN1.frx":C650
      Top             =   5355
      Width           =   765
   End
   Begin VB.Image Image3 
      Height          =   765
      Left            =   2190
      Picture         =   "FXSALPAN1.frx":E5A6
      Top             =   5355
      Width           =   765
   End
   Begin VB.Image Image4 
      Height          =   765
      Left            =   2970
      Picture         =   "FXSALPAN1.frx":104FC
      Top             =   5355
      Width           =   765
   End
   Begin VB.Image Image5 
      Height          =   765
      Left            =   630
      Picture         =   "FXSALPAN1.frx":12452
      Top             =   6195
      Width           =   765
   End
   Begin VB.Image Image6 
      Height          =   765
      Left            =   1410
      Picture         =   "FXSALPAN1.frx":143A8
      Top             =   6195
      Width           =   765
   End
   Begin VB.Image Image7 
      Height          =   765
      Left            =   2190
      Picture         =   "FXSALPAN1.frx":162FE
      Top             =   6195
      Width           =   765
   End
   Begin VB.Image Image8 
      Height          =   765
      Left            =   2970
      Picture         =   "FXSALPAN1.frx":18254
      Top             =   6195
      Width           =   765
   End
End
Attribute VB_Name = "FXSALPAN1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VECES%
'

Private Sub Command1_Click()
   FINDEMO.Show 1
End Sub

Private Sub Command11_Click()
    'Call SECUDEMO("ARCHIMA")
End Sub

Private Sub Command12_Click()
    'Call SECUDEMO("STOCK")
End Sub

Private Sub Command13_Click()
    'Call SECUDEMO("PROGMAT")
End Sub

Private Sub Command18_Click()
   'Call SECUDEMO("ANACOS")
End Sub

Private Sub Form_Activate()
   If YAPRELICEN% > 0 Then Command1.Visible = True
End Sub

Private Sub Form_Click()
   Exit Sub
   If YAPRELICEN% > 0 Then OPCIDEMO.Show 1
End Sub

Private Sub Form_DblClick()
   Close: End
   If VECES% < 70 Then VECES% = 70
End Sub

Private Sub Form_Deactivate()
   Command1.Visible = False
End Sub

Private Sub Image9_Click()
   'If YAPRELICEN% > 0 Then OPCIDEMO.Show 1
End Sub

Private Sub Label1_Click()
   'If YAPRELICEN% > 0 Then OPCIDEMO.Show 1
End Sub

Private Sub Label2_Click()
   'If YAPRELICEN% > 0 Then OPCIDEMO.Show 1
End Sub

Private Sub MODULO_Click(Index As Integer)
   'If YAPRELICEN% > 0 Then OPCIDEMO.Show 1
End Sub

Private Sub Timer1_Timer()
    '
    If VECES% >= 70 Then VECES% = 50
    '
    If VECES% < 1 Then Command1.Visible = False
    VECES% = VECES% + 1
    If VECES% > 6 Then Frame5.Visible = True
    If VECES% > 12 Then Label1.Visible = True
    If VECES% > 18 Then Label2.Visible = True
    '
    idx = Int((VECES% - 24) / 5)
    If idx >= 0 Then
      If idx <= 6 Then
        MODULO(idx).Visible = True
      End If
    End If
    '
    If VECES% < 70 Then Exit Sub
    '
10  Frame5.Visible = True
    Label1.Visible = True
    Label2.Visible = True
    For idx = 0 To 6: MODULO(idx).Visible = True: Next
    '
    Timer1.Enabled = False
    Timer2.Enabled = False
    Image1.Visible = False
    Image2.Visible = False
    Image3.Visible = False
    Image4.Visible = False
    Image5.Visible = False
    Image6.Visible = False
    Image7.Visible = False
    Image8.Visible = False
    Command1.Visible = True
    '
End Sub
'
Private Sub Form_Load()
    '
    If Screen.Width < 12000 Then
       Timer1.Enabled = False
       Timer2.Enabled = False
       CONFIPAN.Show 1
       Close: End
    End If
    '
    'Call CENTRAFORM(Me)
    If VERNUM$ <> "" Then
      Caption = Caption + " - " + VERNUM$
    End If
    '
    Image1.Visible = False
    Image1.Stretch = True
    Image2.Visible = False
    Image2.Stretch = True
    Image3.Visible = False
    Image3.Stretch = True
    Image4.Visible = False
    Image4.Stretch = True
    Image5.Visible = False
    Image5.Stretch = True
    Image6.Visible = False
    Image6.Stretch = True
    Image7.Visible = False
    Image7.Stretch = True
    Image8.Visible = False
    Image8.Stretch = True
    '
    veripan% = 0
    Randomize
    '
    DEMOPATH$ = CurDir
    If DEMOPATH$ = "C:\Archivos de programa\DevStudio\VB" Then
       DEMOPATH$ = "\DEMOFLEX"
    End If
    If Right$(DEMOPATH$, 1) = "\" Then
       DEMOPATH$ = Left$(DEMOPATH$, Len(DEMOPATH$) - 1)
    End If
    '
    File1.Path = "C:\Logos Clientes"
End Sub

Private Sub Timer2_Timer()
    '
    Static VXX, px, py, tamano, INCREX, INCREY
    If VXX = 0 Then
        tamano = 240
        px = (ScaleWidth - tamano) / 2
        py = (ScaleHeight - tamano) / 2
        INCREX = 0.5 - Rnd
        INCREY = 0.5 - Rnd
        While (INCREX ^ 2 + INCREY ^ 2) ^ 0.5 < 0.36
        
            INCREX = INCREX * 1.2
            INCREY = INCREY * 1.2
        Wend
        SIGNO = 1
    End If
    '
    VXX = VXX + 1
    If VXX = 9 Then
        VXX = 1
    End If
    '
    px = px + 100 * INCREX * tamano / 300
    py = py + 100 * INCREY * tamano / 300
    
    tamano = tamano + 40
    If px + tamano < 0 Then
        VXX = 0: Exit Sub
    End If
    '
    If px > ScaleWidth Then
        VXX = 0: Exit Sub
    End If
    '
    If py + tamano < 0 Then
        VXX = 0: Exit Sub
    End If
    '
    If py > ScaleHeight Then
        VXX = 0: Exit Sub
    End If
    '
    If tamano >= ScaleWidth Then VXX = 0
    '
    Select Case VXX
        '
        Case 1
            Image8.Visible = False
            Image1.Height = tamano
            Image1.Width = tamano
            Image1.Move px, py
            Image1.Visible = True
        Case 2
            Image1.Visible = False
            Image2.Height = tamano
            Image2.Width = tamano
            Image2.Move px, py
            Image2.Visible = True
        Case 3
            Image2.Visible = False
            Image3.Height = tamano
            Image3.Width = tamano
            Image3.Move px, py
            Image3.Visible = True
        Case 4
            Image3.Visible = False
            Image4.Height = tamano
            Image4.Width = tamano
            Image4.Move px, py
            Image4.Visible = True
        Case 5
            Image4.Visible = False
            Image5.Height = tamano
            Image5.Width = tamano
            Image5.Move px, py
            Image5.Visible = True
        Case 6
            Image5.Visible = False
            Image6.Height = tamano
            Image6.Width = tamano
            Image6.Move px, py
            Image6.Visible = True
        Case 7
            Image6.Visible = False
            Image7.Height = tamano
            Image7.Width = tamano
            Image7.Move px, py
            Image7.Visible = True
        Case 8
            Image7.Visible = False
            Image8.Height = tamano
            Image8.Width = tamano
            Image8.Move px, py
            Image8.Visible = True
    End Select
    '
End Sub

Function VERNUM$()
   '
   YY$ = ""
3  On Error Resume Next
   XX$ = Trim$(FileDateTime(App.EXEName + ".EXE"))
   On Error GoTo 0
   '
   If XX$ <> "" Then
      For KKI% = 1 To Len(XX$)
         ZZ$ = Mid$(XX$, KKI%, 1)
         KKJ% = Asc(ZZ$)
         If KKJ% <= 32 Then GoTo 5
         If KKJ% > 90 Then GoTo 5
         YY$ = YY$ + ZZ$
      Next KKI%
      '
5     PPXX% = InStr(YY$, "200")
      If PPXX% > 0 Then
         YY$ = Left$(YY$, PPXX% - 1) + Mid$(YY$, PPXX% + 2)
      End If
      '
      For KKI% = 1 To Len(YY$)
        If Mid$(YY$, KKI%, 1) = "/" Then
          Mid$(YY$, KKI%, 1) = "."
        End If
      Next KKI%
      YY$ = "V." + YY$
   End If
   '
   VERNUM$ = YY$
   '
End Function

Private Sub Timer3_Timer()
    Static VXX, px, py, tamano, INCREX, INCREY, IFILE&, RELAVER
    If VXX = 0 Then
        tamano = 240
        px = (ScaleWidth - tamano) / 2
        py = (ScaleHeight - tamano) / 2
        INCREX = 0.5 - Rnd
        INCREY = 0.5 - Rnd
        While (INCREX ^ 2 + INCREY ^ 2) ^ 0.5 < 0.36
        
            INCREX = INCREX * 1.2
            INCREY = INCREY * 1.2
        Wend
        SIGNO = 1
        '
10      On Error GoTo 0
        IFILE& = IFILE& + 1
        If IFILE > File1.ListCount Then IFILE& = 1
        ARXYZ$ = File1.Path + "\" + File1.List(IFILE& - 1)
        Set Image11.Picture = LoadPicture("")
        On Error Resume Next
        Set Image11.Picture = LoadPicture(ARXYZ$)
        If Err > 0 Then GoTo 10
        On Error GoTo 0
        RELAVER = Image11.Height / Image11.Width
        Set Image10.Picture = LoadPicture("")
        Set Image10.Picture = LoadPicture(ARXYZ$)
    End If
    '
    VXX = VXX + 1
    If VXX = 9 Then
        VXX = 1
    End If
    '
    px = px + 100 * INCREX * tamano / 300
    py = py + 100 * INCREY * tamano / 300
    
    tamano = tamano + 40
    If px + tamano < 0 Then
        VXX = 0: Exit Sub
    End If
    '
    If px > ScaleWidth Then
        VXX = 0: Exit Sub
    End If
    '
    If py + tamano < 0 Then
        VXX = 0: Exit Sub
    End If
    '
    If py > ScaleHeight Then
        VXX = 0: Exit Sub
    End If
    '
    If tamano >= ScaleWidth Then VXX = 0
    '
    Image10.Visible = False
    Image10.Height = 2 * tamano * RELAVER
    Image10.Width = 2 * tamano
    Image10.Move px, py
    Image10.Refresh
    Image10.Visible = True
    DoEvents

End Sub

