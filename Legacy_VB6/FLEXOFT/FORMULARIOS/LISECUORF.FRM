VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form LISECUORF 
   BackColor       =   &H00C0D0D0&
   Caption         =   "Secuenciacion de Ordenes de Fabricacion"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   3960
      Left            =   8640
      ScaleHeight     =   3930
      ScaleWidth      =   1005
      TabIndex        =   3
      Top             =   0
      Width           =   1035
      Begin VB.CommandButton BOTEXIT 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   120
         Picture         =   "LISECUORF.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Sale de la Aplicación"
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BOTSAVE 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   120
         Picture         =   "LISECUORF.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Graba las Modificaciones"
         Top             =   1080
         Width           =   750
      End
      Begin VB.CommandButton BOTCUT 
         Caption         =   "Cortar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   5
         Top             =   2940
         Width           =   750
      End
      Begin VB.CommandButton BOTPASTE 
         Caption         =   "Pegar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   3330
         Width           =   750
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   1
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.ListBox LISECU 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4020
      OleObjectBlob   =   "LISECUORF.frx":0454
      Top             =   2100
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   4080
      OleObjectBlob   =   "LISECUORF.frx":0688
      TabIndex        =   2
      Top             =   420
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "LISECUORF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CONSIMULACION%
Dim REGSE$

Private Sub BOTCUT_Click()
    If LISECU.ListIndex < 0 Or LISECU.ListIndex >= LISECU.ListCount Then
        Call MENSERR(24, "Falta Seleccionar Orden de fabricacion")
        Exit Sub
    End If
    ILISTA& = LISECU.ListIndex
    REGSE$ = LISECU.TEXT
    LISECU.RemoveItem ILISTA&
    BOTPASTE.Enabled = True
End Sub

Private Sub BOTEXIT_Click()
    Call CIERRARCH("SECUORF")
    Hide
End Sub

Private Sub BOTPASTE_Click()
    If LISECU.ListIndex < 0 Or LISECU.ListIndex >= LISECU.ListCount Then
        Call MENSERR(24, "Falta Seleccionar Punto de Insersion")
        Exit Sub
    End If
    ILISTA& = LISECU.ListIndex
    LISECU.AddItem REGSE$, ILISTA&
    LISECU.ListIndex = ILISTA&
    BOTCUT.Enabled = True
    BOTPASTE.Enabled = False
End Sub

Private Sub BOTSAVE_Click()
    Screen.MousePointer = 11
    For I& = 0 To LISECU.ListCount - 1
      NOFA$ = Left$(LISECU.LIST(I&), 12)
      IOFA% = Val(Mid$(LISECU.LIST(I&), 11, 4))
      For j& = I& + 1 To LISECU.ListCount - 1
        NOFB$ = Left$(LISECU.LIST(j&), 12)
        IOFB% = Val(Mid$(LISECU.LIST(j&), 11, 4))
        If Left$(NOFB$, 10) = Left$(NOFA$, 10) Then
          If IOFB% > IOFA% Then
            Screen.MousePointer = 1
            Call MENSERR(24, "Secuencia no Procesable Entre " + NOFA$ + " y " + NOFB$ + ".\ \El Orden Debe Ser DECRECIENTE entre Sub-Ordenes\Pertenecientes a la Misma Orden Principal.")
            Exit Sub
          End If
        End If
      Next j&
    Next I&
    
    j& = 0
    
    ARCHSECUORF$ = "SECUORF"
    If CONSIMULACION% > 0 Then
       ARCHSECUORF$ = "SSECUORF"
    End If
    For I& = 0 To LISECU.ListCount - 1
        LISECU.ListIndex = I&
        REGSE$ = Left$(LISECU.TEXT, 12) + Mid$(LISECU.TEXT, 19)
        j& = j& + 1
        Call GRAREG(ARCHSECUORF$, REGSE$, j&)
    Next I&
    Call SETULTREG&(ARCHSECUORF$, j&)
    Call CIERRARCH(ARCHSECUORF$)
    Screen.MousePointer = 1
    Hide
End Sub

Private Sub Form_Load()
   Call APLICACIONSKINS(Me)
End Sub

Private Sub LISECU_DblClick()
    If LISECU.ListIndex < 0 Or LISECU.ListIndex >= LISECU.ListCount Then
        Call MENSERR(24, "Falta Seleccionar Orden de fabricacion")
        Exit Sub
    End If
    ILISTA& = LISECU.ListIndex
    REGSE$ = LISECU.TEXT
    LISECU.RemoveItem ILISTA&
    BOTPASTE.Enabled = True
End Sub
