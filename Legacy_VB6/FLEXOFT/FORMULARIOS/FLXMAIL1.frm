VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form FLXMAIL1 
   Caption         =   "Form1"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   330
      Left            =   5040
      TabIndex        =   6
      Top             =   210
      Width           =   1590
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   435
      Left            =   7350
      TabIndex        =   5
      Top             =   3255
      Width           =   435
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2115
      Left            =   105
      TabIndex        =   4
      Top             =   2940
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   3731
      _Version        =   327680
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"FLXMAIL1.frx":0000
   End
   Begin VB.TextBox Text1 
      Height          =   540
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "FLXMAIL1.frx":00D1
      Top             =   2310
      Width           =   7050
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enviar"
      Height          =   540
      Left            =   3150
      TabIndex        =   2
      Top             =   105
      Width           =   1170
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   105
      TabIndex        =   1
      Top             =   840
      Width           =   7575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   540
      Left            =   1470
      TabIndex        =   0
      Top             =   105
      Width           =   1065
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   735
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   105
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "FLXMAIL1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
MAPISession1.DownLoadMail = False
MAPISession1.LogonUI = True
MAPISession1.Action = 1
MAPIMessages1.SessionID = MAPISession1.SessionID
'BBB = MAPISession1.UserName
'
For kki% = 1 To 3
  MAPIMessages1.MsgIndex = -1
  MAPIMessages1.MsgSubject = "cuarta prueba" + Str$(kki%)
  MAPIMessages1.RecipAddress = "asistec@flexoft.com.ar"
  MAPIMessages1.MsgNoteText = "este es un TERCER mensaje de prueba NUMERO" + Str$(kki%) 'RichTextBox1.TextRTF   '
  'MAPIMessages1.AttachmentIndex = 0
  'MAPIMessages1.AttachmentPathName = "c:\WINDOWS\ICONS\IMAGENES FLEXOFT\LOGO FLEXOFT CHICO.BMP"
  'MAPIMessages1.AttachmentPosition = 0
  'MAPIMessages1.Action = 3   ' vbMessageSend
  'MAPIMessages1.MsgCount = KKI%
  'MAPIMessages1.MsgIndex = KKI% - 1
  'MAPIMessages1.Action = 4   ' vbMessageSAVE
MAPIMessages1.SessionID = MAPISession1.SessionID
  If kki% < 3 Then
       MAPIMessages1.Action = 3   ' vbMessageSend
     Else
       MAPIMessages1.Action = 4   ' vbMessageSAVE
  End If
    MAPIMessages1.SessionID = 0
    DoEvents
'MAPIMessages1.AttachmentPosition
Next kki%
'
'ReadFromFile
    'VBMail.MapiMess.MsgNoteText = txtNoteText
    'VBMail.MapiMess.MsgReceiptRequested = ReturnRequest
'Call CopyNamestoMsgBuffer(Me, True)
                  
'    On Error Resume Next
'    VBMail.MapiMess.Action = vbMessageSend
'    If Err Then
'        MsgBox "Error durante el envío: " + Str$(Err)
'    Else
'        Unload Me
'    End If
    MAPIMessages1.SessionID = 0
    MAPISession1.Action = 2
    '
    MAPISession1.Action = 1
    MAPIMessages1.SessionID = MAPISession1.SessionID
  MAPIMessages1.Action = 3   ' vbMessageSend
    MAPIMessages1.SessionID = 0
    MAPISession1.Action = 2
    Unload Me
    

End Sub

Private Sub Command1_Click()
MAPISession1.DownLoadMail = False
MAPISession1.LogonUI = True
MAPISession1.Action = 1
MAPIMessages1.SessionID = MAPISession1.SessionID
BBB = MAPISession1.UserName
MAPIMessages1.Action = 1
CCC = MAPIMessages1.MsgCount
'
List1.Clear
For i = 0 To CCC - 1
        MAPIMessages1.MsgIndex = i
        A$ = ""
        'If Not mailctl.MsgRead Then
        '    a$ = conUnreadMessage + " "
        '    If UnRead = 0 Then
        '        StartIndex = i  ' Posición de inicio en la lista de correo.
        '    End If
        '    UnRead = UnRead + 1
        'Else
        '    a$ = "  "
        'End If
        A$ = A$ + MAPIMessages1.MsgOrigDisplayName
        b$ = MAPIMessages1.MsgSubject
        c$ = MAPIMessages1.MsgDateReceived
        List1.AddItem A$ + Chr$(9) + b$ + Chr$(9) + c$
    Next i
    'MAPIMessages1.SessionID = 0
    'MAPISession1.Action = 2
    'Unload Me
End Sub



Private Sub Command3_Click()
RichTextBox1.SelColor = RGB(255, 0, 0)
'ttxx$ = RichTextBox1.SelText
'ttyy$ = "": For u& = 1 To Len(ttxx$): ttyy$ = ttyy$ + Str$(Asc(Mid$(ttxx$, u&, 1))): Next
'MsgBox ttyy$
End Sub


Private Sub List1_Click()
MAPIMessages1.MsgIndex = List1.ListIndex
TTX1$ = MAPIMessages1.MsgNoteText
TTX2$ = ""
For KKU& = 1 To Len(TTX1$)
CARA% = Asc(Mid$(TTX1$, KKU&, 1))
'If CARA% >= 32 Then
  If CARA% < 128 Then
    TTX2$ = TTX2$ + Chr$(CARA%)
  End If
'End If
Next KKU&
RichTextBox1.TextRTF = TTX2$
End Sub
