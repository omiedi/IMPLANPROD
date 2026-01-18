VERSION 5.00
Begin VB.Form EXCLDOCS 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hoja de Cálculo Excel"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11880
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin VB.HScrollBar HScroll1 
      Height          =   220
      Left            =   0
      TabIndex        =   5
      Top             =   8030
      Width           =   10650
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   11025
      Picture         =   "EXCLDOCS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   105
      Width           =   750
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7995
      Left            =   10660
      TabIndex        =   2
      Top             =   0
      Width           =   225
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7995
      Left            =   0
      ScaleHeight     =   7935
      ScaleWidth      =   10560
      TabIndex        =   1
      Top             =   0
      Width           =   10620
      Begin VB.OLE OLE1 
         BorderStyle     =   0  'None
         Class           =   "Excel.Sheet.8"
         Height          =   1065
         Left            =   0
         SizeMode        =   2  'AutoSize
         TabIndex        =   0
         Top             =   0
         Width           =   3270
      End
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   10640
      Y1              =   8000
      Y2              =   8000
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
Attribute VB_Name = "EXCLDOCS"
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
'Exit Sub
On Error Resume Next
OLE1.DoVerb (-1)
If Err Then
  On Error GoTo 0
  Call CONECRUN("PROMRP02", "")
  Close: End
End If
On Error GoTo 0
Hide
Exit Sub
'   If NUEVODOC% = 1 Then
'      OLE1.Action = 7
'      NUEVODOC% = 0
'    Else
'      OLE1.SetFocus
'   End If
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

Private Sub HScroll1_Change()
   OLE1.Left = (-1) * HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
   OLE1.Left = (-1) * HScroll1.Value
End Sub

