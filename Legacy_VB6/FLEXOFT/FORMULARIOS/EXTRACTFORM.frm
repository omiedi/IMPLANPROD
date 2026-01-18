VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form EXTRACTFORM 
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
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.FileListBox File1 
      Height          =   1065
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
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "EXTRACTFORM.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "EXTRACTFORM.frx":0234
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   315
      Picture         =   "EXTRACTFORM.frx":0292
      Top             =   735
      Width           =   480
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
Attribute VB_Name = "EXTRACTFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
   '
   Call TRANSLABELS(Name)
   '
   Call APLICACIONSKINS(Me)

   ARCHIN$ = TRIM$(Label18.Caption)
   '
   If TRIM$(ARCHIN$) = "Controladores" Then
     File1.Path = "\RFSX"
     File1.Pattern = "*.RFS"
     FIN& = File1.ListCount
     For II& = 1 To FIN&
       Call TRACE(20, II&, FIN&)
       DoEvents
       ORIGI$ = "\RFSX\" + TRIM$(UCase$(File1.List(II& - 1)))
       DESTI$ = LUCOM$ + "DATOS\" + TRIM$(UCase$(File1.List(II& - 1)))
       On Error Resume Next
       FileCopy ORIGI$, DESTI$
     Next II&
     Screen.MousePointer = 1
     On Error GoTo 0
     GoTo 90
   End If
   '
   If TRIM$(ARCHIN$) = "Datos de Prueba" Then
     File1.Path = "\FX_DATPRUE"
     File1.Pattern = "*.*"
     FIN& = File1.ListCount
     For II& = 1 To FIN&
       Call TRACE(20, II&, FIN&)
       DoEvents
       ORIGI$ = "\FX_DATPRUE\" + TRIM$(UCase$(File1.List(II& - 1)))
       DESTI$ = LUCOM$ + "DATOS\" + TRIM$(UCase$(File1.List(II& - 1)))
       On Error Resume Next
       FileCopy ORIGI$, DESTI$
     Next II&
     Screen.MousePointer = 1
     On Error GoTo 0
     GoTo 90
   End If
   '
   If ARCHIN$ <> "" Then
     '
     Label18.Visible = True
     Label18.Refresh
     DoEvents
     Dim BUFGRA As String * 128
     Dim BUFLE As String * 160
     NX% = FreeFile
     Open LUBAS$ + "FLEXPRFS.FLX" For Random Access Read Shared As #NX% Len = 160
     FIN& = LOF(NX%) / 160
     '
     LUDESTI$ = LUDAT$
     If InStr(ARCHIN$, ".PRF") > 0 Then LUDESTI$ = CurDir
     If Right$(LUDESTI$, 1) <> "\" Then LUDESTI$ = LUDESTI$ + "\"
     '
     n1% = FreeFile
     Open LUDESTI$ + ARCHIN$ For Random Access Read Write Shared As #n1% Len = 128
     JJ& = 0
     '
     LI& = 0: LS& = FIN& + 1
     INI& = 0
     ABRE$ = TRIM$(UCase$(ARCHIN$))
1510 E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 1580
     L& = LI& + E&: If L& < 1 Or L& > FIN& Then GoTo 1580
     Get #NX%, L&, BUFLE$
     I1$ = TRIM$(UCase$(Left$(BUFLE$, 32)))
     If I1$ < ABRE$ Then LI& = L&: GoTo 1510
     If I1$ > ABRE$ Then LS& = L&: GoTo 1510
     INI& = L& - 128: If INI& < 1 Then INI& = 1
     '
1580 If INI& < 1 Then
        Call MENSERR(24, "Imposible Extraer Controlador " + ABRE$)
        GoTo 80
     End If
     '
     For II& = INI& To FIN&
       Call TRACE(20, II&, FIN&)
       DoEvents
       Get #NX%, II&, BUFLE$
       I1$ = TRIM$(UCase$(Left$(BUFLE$, 32)))
       If I1$ > ARCHIN$ Then GoTo 80
       If I1$ = ARCHIN$ Then
         BUFGRA$ = Mid$(BUFLE$, 33)
         JJ& = JJ& + 1
         Put #n1%, JJ&, BUFGRA$
       End If
     Next II&
     '
80   Close #n1%
     Close #NX%
     '
90   On Error Resume Next
     Hide
     On Error GoTo 0
     Call FRESHALL
     DoEvents
     '
   End If
End Sub


Private Sub Timer1_Timer()
   If Image1.Visible = True Then
       Image1.Visible = False
     Else
       Image1.Visible = True
   End If
   DoEvents
End Sub

