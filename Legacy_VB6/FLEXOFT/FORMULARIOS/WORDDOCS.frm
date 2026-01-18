VERSION 5.00
Begin VB.Form WORDDOCS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Documentos en Word 6.0"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   6.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   11025
      Picture         =   "WORDDOCS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   105
      Width           =   750
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   8310
      Left            =   10660
      TabIndex        =   2
      Top             =   0
      Width           =   225
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      Height          =   8310
      Left            =   0
      ScaleHeight     =   8250
      ScaleWidth      =   10560
      TabIndex        =   1
      Top             =   0
      Width           =   10620
      Begin VB.OLE OLE1 
         BorderStyle     =   0  'None
         Height          =   1065
         Left            =   0
         SizeMode        =   2  'AutoSize
         TabIndex        =   0
         Top             =   0
         Width           =   3270
      End
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Line Line1 
      X1              =   10650
      X2              =   10650
      Y1              =   0
      Y2              =   8295
   End
End
Attribute VB_Name = "WORDDOCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command2_Click()
   ' MsgBox WORDDOCS.OLE1.SourceDoc
   DOCUACTIVO$ = TRIM$(UCase$(OLE1.SourceDoc))
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Form_Activate()
   Call CENTRAFORM(Me)
   If NUEVODOC% = 1 Then
      OLE1.Action = 7
      NUEVODOC% = 0
    Else
      On Error Resume Next
      OLE1.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   DOCUACTIVO$ = TRIM$(UCase$(OLE1.SourceDoc))
End Sub

Private Sub VScroll1_Change()
   OLE1.Top = (-1) * VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
   OLE1.Top = (-1) * VScroll1.Value
End Sub
