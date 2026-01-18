VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FRMREVUPD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   16275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6900
      Left            =   7320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Top             =   0
      Width           =   7980
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Copiar"
      Height          =   255
      Left            =   14280
      TabIndex        =   11
      Top             =   6960
      Width           =   1000
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   3840
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00404040&
      Height          =   8520
      Left            =   15360
      ScaleHeight     =   8460
      ScaleWidth      =   1005
      TabIndex        =   4
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command14 
         Height          =   700
         Left            =   80
         Picture         =   "FRMREVUPD.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Revertir Texto"
         Top             =   6480
         Width           =   700
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   80
         Picture         =   "FRMREVUPD.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cancela Observaciones"
         Top             =   50
         Width           =   700
      End
   End
   Begin VB.TextBox Text28 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6420
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   7260
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   795
      Left            =   105
      ScaleHeight     =   735
      ScaleWidth      =   6255
      TabIndex        =   1
      Top             =   -30
      Width           =   6320
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Revierte un update por un Mostar Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   6300
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FRMREVUPD.frx":0294
      Top             =   600
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "FRMREVUPD.frx":04C8
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Caracteres Disponibles:"
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
      Left            =   120
      TabIndex        =   10
      Top             =   6960
      Width           =   2025
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "512"
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
      Left            =   2280
      TabIndex        =   9
      Top             =   6960
      Width           =   330
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   210
      TabIndex        =   6
      Top             =   4725
      Width           =   750
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   210
      TabIndex        =   3
      Top             =   5250
      Width           =   795
   End
End
Attribute VB_Name = "FRMREVUPD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command14_Click()
    Command14.Enabled = False
    Dim strTest As String
    Dim strArray() As String
    Dim intCount As Integer
   
    strArray = Split(Text28, Chr(13))
    Text1 = ""
    For i& = LBound(strArray) To UBound(strArray)
      
       A = strArray(i&)
       On Error Resume Next
       strArray2 = Split(A, "=")
       If UBound(strArray2) <> 1 Then
         Text1 = Text1 & A & "   *************" & vbCrLf
         GoTo 50 'NEXT
       End If
       Text1 = Text1 + TRIM$(strArray2(1)) & " = " & TRIM$(strArray2(0)) & vbCrLf
     
50   Next i&

    
    
    
    
    Command14.Enabled = True

End Sub

Private Sub Command22_Click()
   Unload Me
End Sub

Private Sub Command17_Click()
   Text1.SelStart = 0
   Text1.SelLength = Len(Text28.TEXT)
   Text1.SetFocus
   Clipboard.SetText Mid$(Text1.TEXT, Text1.SelStart + 1, Text1.SelLength)
End Sub
'
Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
End Sub
'
Private Sub Command6_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Clipboard.GetText + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub
'
Private Sub Command5_Click()
   Text28.SelStart = 0
   Text28.SelLength = Len(Text28.TEXT)
   Text28.SetFocus
End Sub
'

Private Sub Form_Activate()
   Call PINTATEXT(Text28)
End Sub

Private Sub Form_Load()
   '
   
   'Call APLICACIONSKINS(Me, Picture5)
   '
End Sub

Private Sub Text28_Change()
   
   Label2 = Text28.MaxLength - Len(Text28)
End Sub
