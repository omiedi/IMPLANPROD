VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form SENDMAIANT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample Client Application for vbSendMail Component"
   ClientHeight    =   7590
   ClientLeft      =   1755
   ClientTop       =   1710
   ClientWidth     =   7875
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   7875
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   6000
      Begin VB.TextBox txtBcc 
         Height          =   285
         Left            =   1605
         TabIndex        =   12
         Top             =   2280
         Width           =   4200
      End
      Begin VB.TextBox txtCcName 
         Height          =   285
         Left            =   1605
         TabIndex        =   11
         Top             =   1560
         Width           =   4200
      End
      Begin VB.TextBox txtCc 
         Height          =   285
         Left            =   1605
         TabIndex        =   10
         Top             =   1920
         Width           =   4200
      End
      Begin VB.TextBox txtTo 
         Height          =   285
         Left            =   1605
         TabIndex        =   9
         Top             =   1200
         Width           =   4200
      End
      Begin VB.TextBox txtToName 
         Height          =   285
         Left            =   1605
         TabIndex        =   8
         Top             =   840
         Width           =   4200
      End
      Begin VB.Label lblBcc 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bcc: Email"
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
         Left            =   0
         TabIndex        =   17
         Top             =   2340
         Width           =   915
      End
      Begin VB.Label lblCcName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cc: Name"
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
         Left            =   165
         TabIndex        =   16
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label lblCC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cc: Email"
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
         Left            =   105
         TabIndex        =   15
         Top             =   1920
         Width           =   810
      End
      Begin VB.Label lblTo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recipient Email"
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
         Left            =   45
         TabIndex        =   14
         Top             =   1260
         Width           =   1335
      End
      Begin VB.Label lblToName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recipient Name"
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
         Left            =   45
         TabIndex        =   13
         Top             =   900
         Width           =   1365
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6960
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   315
      Left            =   6420
      TabIndex        =   6
      Top             =   1140
      Width           =   1275
   End
   Begin VB.ListBox lstStatus 
      BackColor       =   &H8000000F&
      Height          =   1035
      Left            =   1980
      TabIndex        =   4
      Top             =   5760
      Width           =   4200
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   315
      Left            =   6420
      TabIndex        =   0
      Top             =   5760
      Width           =   1275
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   315
      Left            =   6420
      TabIndex        =   2
      Top             =   600
      Width           =   1275
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   315
      Left            =   6420
      TabIndex        =   1
      Top             =   180
      Width           =   1275
   End
   Begin MSWinsockLib.Winsock WinSck 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Left            =   975
      TabIndex        =   5
      Top             =   5820
      Width           =   555
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Progress  "
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
      Left            =   3780
      TabIndex        =   3
      Top             =   6960
      Width           =   870
   End
End
Attribute VB_Name = "SENDMAIANT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Option Compare Text


' *****************************************************************************
' Required declaration of the vbSendMail component (withevents is optional)
' You also need a reference to the vbSendMail component in the Project References
' *****************************************************************************
'Public WithEvents posendmail As vbSendMail.clsSendMail
Dim posendmail As New clsSendMail
'Set posendmail = New clsSendMail

' misc local vars
Dim bAuthLogin      As Boolean
Dim bPopLogin       As Boolean
Dim bHtml           As Boolean
Dim MyEncodeType    As ENCODE_METHOD
Dim etPriority      As MAIL_PRIORITY
Dim bReceipt        As Boolean



Sub cmdSend_Click()

    ' *****************************************************************************
    ' This is where all of the Components Properties are set / Methods called
    ' *****************************************************************************

    cmdSend.Enabled = False
    lstStatus.Clear
    Screen.MousePointer = vbHourglass
    
    With posendmail

        ' **************************************************************************
        ' Optional properties for sending email, but these should be set first
        ' if you are going to use them
        ' **************************************************************************
 
        .SMTPHostValidation = VALIDATE_NONE         ' Optional, default = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = ";"                            ' Optional, default = ";" (semicolon)

        ' **************************************************************************
        ' Basic properties for sending email
        ' **************************************************************************
        .SMTPHost = TRIM(ARMANEW09.txtServer.TEXT)                  ' Required the fist time, optional thereafter
        .from = ARMANEW09.txtFrom.TEXT                        ' Required the fist time, optional thereafter
        .FromDisplayName = ARMANEW09.txtFromName.TEXT         ' Optional, saved after first use
        .Recipient = "landryfj@yahoo.com.ar" 'txtTo.TEXT                     ' Required, separate multiple entries with delimiter character
        .RecipientDisplayName = txtToName.TEXT      ' Optional, separate multiple entries with delimiter character
        .CcRecipient = txtCc                        ' Optional, separate multiple entries with delimiter character
        .CcDisplayName = txtCcName                  ' Optional, separate multiple entries with delimiter character
        .BccRecipient = txtBcc                      ' Optional, separate multiple entries with delimiter character
        .ReplyToAddress = ARMANEW09.txtFrom.TEXT              ' Optional, used when different than 'From' address
        .Subject = ARMANEW09.txtSubject.TEXT                  ' Optional
        .Message = ARMANEW09.txtMsg.TEXT                      ' Optional
        .Attachment = TRIM(ARMANEW09.txtAttach.TEXT)          ' Optional, separate multiple entries with delimiter character

        ' **************************************************************************
        ' Additional Optional properties, use as required by your application / environment
        ' **************************************************************************
        .AsHTML = bHtml                             ' Optional, default = FALSE, send mail as html or plain text
        .ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
        .EncodeType = MyEncodeType                  ' Optional, default = MIME_ENCODE
        .Priority = etPriority                      ' Optional, default = PRIORITY_NORMAL
        .Receipt = bReceipt                         ' Optional, default = FALSE
        .UseAuthentication = bAuthLogin             ' Optional, default = FALSE
        .UsePopAuthentication = bPopLogin           ' Optional, default = FALSE
        .Username = txtUserName                     ' Optional, default = Null String
        .Password = txtPassword                     ' Optional, default = Null String, value is NOT saved
        .POP3Host = txtPopServer
        .MaxRecipients = 100                        ' Optional, default = 100, recipient count before error is raised
        
        ' **************************************************************************
        ' Advanced Properties, change only if you have a good reason to do so.
        ' **************************************************************************
        ' .ConnectTimeout = 10                      ' Optional, default = 10
        ' .ConnectRetry = 5                         ' Optional, default = 5
        ' .MessageTimeout = 60                      ' Optional, default = 60
        ' .PersistentSettings = True                ' Optional, default = TRUE
        ' .SMTPPort = 25                            ' Optional, default = 25

        ' **************************************************************************
        ' OK, all of the properties are set, send the email...
        ' **************************************************************************
        ' .Connect                                  ' Optional, use when sending bulk mail
        
Stop
        .Send                                       ' Required
        ' .Disconnect                               ' Optional, use when sending bulk mail
        ARMANEW09.txtServer.TEXT = .SMTPHost                  ' Optional, re-populate the Host in case
                                                    ' MX look up was used to find a host    End With
    End With
    Screen.MousePointer = vbDefault
    cmdSend.Enabled = True

End Sub

Private Sub Command1_Click()
    '
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Archivo Origen de Datos"
    CommonDialog1.Filter = "Archivos Html(*.*)|*.*|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
   
    DoEvents
    Screen.MousePointer = 11
    NOMARCH$ = UCase$(CommonDialog1.FileName)
    TXTX$ = LOADFILE(NOMARCH$)
    
99

End Sub

' *****************************************************************************
' The following four Subs capture the Events fired by the vbSendMail component
' *****************************************************************************

Private Sub poSendMail_Progress(lPercentCompete As Long)

    ' vbSendMail 'Progress Event'
    lblProgress = lPercentCompete & "% complete"

End Sub

Private Sub poSendMail_SendFailed(Explanation As String)

    ' vbSendMail 'SendFailed Event
    MsgBox ("Your attempt to send mail failed for the following reason(s): " & vbCrLf & Explanation)
    lblProgress = ""
    Screen.MousePointer = vbDefault
    cmdSend.Enabled = True
    
End Sub

Private Sub poSendMail_SendSuccesful()

    ' vbSendMail 'SendSuccesful Event'
    MsgBox "Send Successful!"
    lblProgress = ""

End Sub

Private Sub poSendMail_Status(Status As String)

    ' vbSendMail 'Status Event'
    lstStatus.AddItem Status
    lstStatus.ListIndex = lstStatus.ListCount - 1
    lstStatus.ListIndex = -1

End Sub

Sub Form_Load()
    ' *****************************************************************************
    ' Required to activate the vbSendMail component.
    ' *****************************************************************************

'    With Me
'        .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
'        .fraOptions.Height = 2475
'        .lblProgress = ""
'    End With

'    cboPriority.AddItem "Normal"
'    cboPriority.AddItem "High"
'    cboPriority.AddItem "Low"
'    cboPriority.ListIndex = 0
'
'    CenterControlsVertical 100, False, txtServer, txtPopServer, txtFromName, txtFrom, txtToName, txtTo, txtCcName, txtCc, txtBcc, txtSubject, txtMsg, txtAttach, lstStatus, lblProgress
'    AlignControlsTop False, cmdSend 'lblServer, 'txtServer
'    CenterControlsHorizontal 300, False, lblServer, txtServer, cmdSend
'    AlignControlsLeft False, lblToName, lblTo, lblCcName, lblCC, lblBcc, lblSubject, lblMsg, lstStatus, lblAttach, lblStatus    'lblServer,lblPopServer,lblFromName,lblFrom,

    'CenterControlRelativeVertical , txtServer 'lblServer
    'CenterControlRelativeVertical txtPopServer 'lblPopServer
    'CenterControlRelativeVertical cmdSend ', txtServer
    'CenterControlRelativeVertical lblFromName, txtFromName
    'CenterControlRelativeVertical cmdReset, txtPopServer
    'CenterControlRelativeVertical lblFrom, txtFrom
    'CenterControlRelativeVertical lblToName, txtToName
    'CenterControlRelativeVertical cmdExit, txtFrom
    'CenterControlRelativeVertical lblTo, txtTo
    'CenterControlRelativeVertical lblCcName, txtCcName
    'CenterControlRelativeVertical lblCC, txtCc
    'CenterControlRelativeVertical lblBcc, txtBcc
    'CenterControlRelativeVertical lblSubject, txtSubject
    'CenterControlRelativeVertical lblAttach, txtAttach
    'CenterControlRelativeVertical cmdBrowse, txtAttach
    'AlignControlsTop False, txtMsg, lblMsg
    'AlignControlsTop False, lstStatus, lblStatus

'    fraOptions.Top = txtTo.Top - 135
'
'    'AlignControlsLeft True, txtServer, txtPopServer, txtFromName, txtFrom, txtToName, txtTo, txtCcName, txtCc, txtBcc, txtSubject, txtMsg, lstStatus, txtAttach, lblProgress
'    AlignControlsLeft True, cmdSend, cmdReset, cmdExit, cmdBrowse, fraOptions
'
'    lblPopServer.Visible = False
'    txtPopServer.Visible = False

    'Me.Show 1

'    RetrieveSavedValues

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' *****************************************************************************
    ' Unload the component before quiting.
    ' *****************************************************************************

    Set posendmail = Nothing

End Sub

Private Sub RetrieveSavedValues()

    ' *****************************************************************************
    ' Retrieve saved values by reading the components 'Persistent' properties
    ' *****************************************************************************
    posendmail.PersistentSettings = True
    txtServer.TEXT = posendmail.SMTPHost
    txtPopServer.TEXT = posendmail.POP3Host
    txtFrom.TEXT = posendmail.from
    txtFromName.TEXT = posendmail.FromDisplayName
    txtUserName = posendmail.Username
    'optEncodeType(poSendMail.EncodeType).Value = True
    If posendmail.UseAuthentication Then ckLogin = vbChecked Else ckLogin = vbUnchecked

End Sub

Private Sub optEncodeType_Click(Index As Integer)

'    If optEncodeType(0).Value = True Then
'        MyEncodeType = MIME_ENCODE
'        cboPriority.Enabled = True
'        ckHtml.Enabled = True
'        ckReceipt.Enabled = True
'        ckLogin.Enabled = True
'    Else
'        MyEncodeType = UU_ENCODE
'        ckHtml.Value = vbUnchecked
'        ckReceipt.Value = vbUnchecked
'        ckLogin.Value = vbUnchecked
'        cboPriority.Enabled = False
'        ckHtml.Enabled = False
'        ckReceipt.Enabled = False
'        ckLogin.Enabled = False
'    End If

End Sub

Private Sub cboPriority_Click()

    Select Case cboPriority.ListIndex

        Case 0: etPriority = NORMAL_PRIORITY
        Case 1: etPriority = HIGH_PRIORITY
        Case 2: etPriority = LOW_PRIORITY

    End Select

End Sub

Private Sub cboPriority_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode

        Case 38, 40

        Case Else: KeyCode = 0

    End Select

End Sub

Private Sub cboPriority_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub ckHtml_Click()

    If ckHtml.Value = vbChecked Then bHtml = True Else bHtml = False

End Sub

Private Sub ckLogin_Click()

    If ckLogin.Value = vbChecked Then
        bAuthLogin = True
        fraOptions.Height = 3555
    Else
        bAuthLogin = False
        If ckPopLogin.Value = vbUnchecked Then fraOptions.Height = 2475
    End If

End Sub

Private Sub ckPopLogin_Click()

    If ckPopLogin.Value = vbChecked Then
        bPopLogin = True
        lblPopServer.Visible = True
        txtPopServer.Visible = True
        fraOptions.Height = 3555
    Else
        bPopLogin = False
        lblPopServer.Visible = False
        txtPopServer.Visible = False
        If ckLogin.Value = vbUnchecked Then fraOptions.Height = 2475
    End If

End Sub

Private Sub ckReceipt_Click()

    If ckReceipt.Value = vbChecked Then bReceipt = True Else bReceipt = False

End Sub

Private Sub cmdBrowse_Click()

    Dim sFilenames()    As String
    Dim i               As Integer
    
    On Local Error GoTo Err_Cancel
  
    With cmDialog
        .FileName = ""
        .CancelError = True
        .Filter = "All Files (*.*)|*.*|HTML Files (*.htm;*.html;*.shtml)|*.htm;*.html;*.shtml|Images (*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif"
        .FilterIndex = 1
        .DialogTitle = "Select File Attachment(s)"
        .MaxFileSize = &H7FFF
        .flags = &H4 Or &H800 Or &H40000 Or &H200 Or &H80000
        .ShowOpen
        ' get the selected name(s)
        sFilenames = Split(.FileName, vbNullChar)
    End With
    
    If UBound(sFilenames) = 0 Then
        If txtAttach.TEXT = "" Then
            txtAttach.TEXT = sFilenames(0)
        Else
            txtAttach.TEXT = txtAttach.TEXT & ";" & sFilenames(0)
        End If
    ElseIf UBound(sFilenames) > 0 Then
        If Right$(sFilenames(0), 1) <> "\" Then sFilenames(0) = sFilenames(0) & "\"
        For i = 1 To UBound(sFilenames)
            If txtAttach.TEXT = "" Then
                txtAttach.TEXT = sFilenames(0) & sFilenames(i)
            Else
                txtAttach.TEXT = txtAttach.TEXT & ";" & sFilenames(0) & sFilenames(i)
            End If
        Next
    Else
        Exit Sub
    End If
    
Err_Cancel:

End Sub

Private Sub cmdExit_Click()

Dim frm As Form

For Each frm In Forms
    Unload frm
    Set frm = Nothing
Next

End

End Sub

Private Sub cmdReset_Click()

    ClearTextBoxesOnForm
    lstStatus.Clear
    lblProgress = ""
    RetrieveSavedValues

End Sub

Private Sub AlignControlsLeft(StandardizeWidth As Boolean, base As Object, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com
    On Error Resume Next

    Dim i As Integer
    For i = 0 To UBound(cnts)
        cnts(i).Left = base.Left
        If StandardizeWidth Then cnts(i).Width = base.Width
    Next

End Sub

Private Sub CenterControlsVertical(space As Single, AlignLeft As Boolean, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    Dim sngTotalSpace As Single
    Dim i As Integer
    Dim sngBaseLeft As Single

    Dim sngParentHeight As Single

    sngParentHeight = Me.ScaleHeight

    For i = 0 To UBound(cnts)
        sngTotalSpace = sngTotalSpace + cnts(i).Height
    Next

    sngTotalSpace = sngTotalSpace + (space * (UBound(cnts)))
    cnts(0).Top = (sngParentHeight - sngTotalSpace) / 2

    sngBaseLeft = cnts(0).Left

    For i = 1 To UBound(cnts)
        cnts(i).Top = cnts(i - 1).Top + cnts(i - 1).Height + space
        If AlignLeft Then cnts(i).Left = sngBaseLeft
    Next

End Sub

Private Sub CenterControlHorizontal(child As Object)

    child.Left = (Me.ScaleWidth - child.Width) / 2

End Sub

Public Sub CenterControlsHorizontal(space As Single, AlignTop As Boolean, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    Dim sngTotalSpace As Single
    Dim i As Integer
    Dim sngBaseTop As Single
    Dim sngParentWidth As Single

    sngParentWidth = Me.ScaleWidth

    For i = 0 To UBound(cnts)
        sngTotalSpace = sngTotalSpace + cnts(i).Width
    Next

    sngTotalSpace = sngTotalSpace + (space * (UBound(cnts)))

    cnts(0).Left = (sngParentWidth - sngTotalSpace) / 2
    sngBaseTop = cnts(0).Top

    For i = 1 To UBound(cnts)
        cnts(i).Left = cnts(i - 1).Left + cnts(i - 1).Width + space
        If AlignTop Then cnts(i).Top = sngBaseTop
    Next

End Sub

Public Sub AlignControlsTop(StandardizeHeight As Boolean, base As Object, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    On Error Resume Next
    Dim i As Integer
    For i = 0 To UBound(cnts)
        cnts(i).Top = base.Top
        If StandardizeHeight Then cnts(i).Height = base.Height
    Next

End Sub

Public Sub CenterControlRelativeVertical(ctl As Object, RelativeTo As Object)

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    On Error Resume Next
    ctl.Top = RelativeTo.Top + ((RelativeTo.Height - ctl.Height) / 2)

End Sub

Public Sub SetHorizontalDistance(distance As Single, StandardizeWidth As Boolean, AlignTop As Boolean, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    On Error Resume Next
    Dim i As Integer
    For i = 1 To UBound(cnts)
        If StandardizeWidth Then cnts(i).Width = cnts(i - 1).Width
        cnts(i).Left = cnts(i - 1).Left + cnts(i - 1).Width + distance
        If AlignTop Then cnts(i).Top = cnts(i - 1).Top
    Next

End Sub

Public Sub CenterControlsRelativeHorizontal(RelativeTo As Object, space As Single, ParamArray cnts())

    ' This is a modified version of a function in the SmartForm class,
    ' which is available on http://www.freevbcode.com

    On Error Resume Next
    Dim sngTotalWidth As Single
    Dim i As Integer
    For i = 0 To UBound(cnts)
        sngTotalWidth = sngTotalWidth + cnts(i).Width
        If i < UBound(cnts) Then sngTotalWidth = sngTotalWidth + space
    Next

    cnts(0).Left = RelativeTo.Left + ((RelativeTo.Width - sngTotalWidth) / 2)

    For i = 1 To UBound(cnts)
        cnts(i).Left = cnts(i - 1).Left + cnts(i - 1).Width + space
        cnts(i).Top = cnts(0).Top
    Next

End Sub

Public Sub ClearTextBoxesOnForm()

    ' Snippet Taken From http://www.freevbcode.com

    Dim ctl As CONTROL
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.TEXT = ""
        End If
    Next

End Sub

