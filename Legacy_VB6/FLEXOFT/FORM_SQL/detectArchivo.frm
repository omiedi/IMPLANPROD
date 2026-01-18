VERSION 5.00
Begin VB.Form DetectArchivo 
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   3690
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "DetectArchivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetLogicalDriveStrings Lib "kernel32" _
Alias "GetLogicalDriveStringsA" _
(ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Function GetDriveStrings() As String
    Dim result As Long
    Dim strDrives As String
    Dim lenStrDrives As Long
    
    result = GetLogicalDriveStrings(0, strDrives)
    
    strDrives = String(result, 0)
    lenStrDrives = result
    
    result = GetLogicalDriveStrings(lenStrDrives, strDrives)
    
    If result = 0 Then
    GetDriveStrings = ""
    Else
    GetDriveStrings = strDrives
    End If
End Function

Private Sub Command1_Click()
    Dim strDrives As String
    strDrives = GetDriveStrings()
    Screen.MousePointer = 1
    Command1.Enabled = False
    If strDrives = "" Then
      MsgBox "No se han encontrado unidades", vbCritical
      Command1.Enabled = True
    Else
      DisplayDriveTypes strDrives
    End If
    If List1.ListCount > 0 Then
        A = List1.LIST(0)
        Screen.MousePointer = 1
         Call ABREBASE_MDB(List1.LIST(0))
    End If
'    Unload Me
End Sub

Private Sub DisplayDriveTypes(drives As String)
    Dim pos As Long
    Dim drive As String
    
    List1.Clear
    pos = 1
    Do While Not Mid$(drives, pos, 1) = Chr(0)
        drive = Mid$(drives, pos, 3)
        MsgBox drive
        pos = pos + 4
        On Error Resume Next
        FindInDirectory drive, "AMBEST\Ambest4280.TXT"
        On Error GoTo 0
    Loop
End Sub
Sub FindInDirectory(ByVal Path As String, ByVal Find As String)
    Dim FileName As String
    FileName = Dir(Path & Find, vbArchive)
    If EXISTE%(Path & Find) > 0 Then
      FDFD = 1
    End If
    'Buscar dentro de la carpeta'
    Do While FileName <> ""
        List1.AddItem Path
        List1.AddItem Path & FileName
        FileName = Dir
    Loop
End Sub

Private Sub Form_Load()
    EPAC$ = TRIM$(EMPREAC$)
'    VERANTER$ = "CONSTOCK"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    Call GRATITFOR(Caption)
    '
    LU$ = LUDAT$

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub
