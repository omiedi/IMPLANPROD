VERSION 5.00
Begin VB.Form DOCONEC09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vinculación Documentos Relacionados"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1180
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   8700
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto o Prospecto:"
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
         Left            =   -210
         TabIndex        =   24
         Top             =   495
         Width           =   2400
      End
      Begin VB.Label Label3 
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
         Height          =   540
         Left            =   3255
         TabIndex        =   23
         Top             =   480
         Width           =   5280
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
         Left            =   1470
         TabIndex        =   22
         Top             =   735
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   3500
      Left            =   9000
      ScaleHeight     =   3435
      ScaleWidth      =   945
      TabIndex        =   17
      Top             =   0
      Width           =   1000
      Begin VB.CommandButton Command2 
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
         Height          =   675
         Left            =   105
         Picture         =   "DOCONEC09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   200
         Width           =   640
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   18
         Top             =   2700
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   19
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "DOCONEC09.frx":014A
         Top             =   3030
         Width           =   1515
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Documentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3800
      Left            =   120
      TabIndex        =   4
      Top             =   3550
      Width           =   9650
      Begin VB.CommandButton Command6 
         Height          =   360
         Left            =   8685
         Picture         =   "DOCONEC09.frx":206C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Agregar Documentos"
         Top             =   3285
         Width           =   750
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1170
         TabIndex        =   9
         Top             =   3315
         Width           =   7440
      End
      Begin VB.CommandButton Command7 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   8685
         Picture         =   "DOCONEC09.frx":21B6
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Visualizar Documento Activo"
         Top             =   2205
         Width           =   750
      End
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   3900
         TabIndex        =   7
         Top             =   285
         Width           =   5520
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   195
         TabIndex        =   6
         Top             =   675
         Width           =   3495
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   1470
         TabIndex        =   5
         Top             =   255
         Width           =   2220
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Conectar Como:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   225
         TabIndex        =   16
         Top             =   3210
         Width           =   810
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   330
         TabIndex        =   15
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Label27 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1170
         TabIndex        =   14
         Top             =   2235
         Width           =   7440
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         TabIndex        =   13
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label Label29 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1170
         TabIndex        =   12
         Top             =   2595
         Width           =   7440
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Carpeta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         TabIndex        =   11
         Top             =   2580
         Width           =   810
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Documentos Conectados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   120
      TabIndex        =   0
      Top             =   1350
      Width           =   8730
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   285
         Width           =   7665
      End
      Begin VB.CommandButton Command11 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   7900
         Picture         =   "DOCONEC09.frx":24C0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Visualizar Documento Adjunto"
         Top             =   1170
         Width           =   700
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Del"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   7900
         Picture         =   "DOCONEC09.frx":27CA
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Baja del Archivo Adjunto"
         Top             =   285
         Width           =   700
      End
   End
End
Attribute VB_Name = "DOCONEC09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RUTATMP$
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command11_Click()

    Command11.Enabled = False
    If List2.ListIndex > -1 Then
        NOMBRE$ = TRIM$(Mid$(List2.List(List2.ListIndex), 1, 64))
        NOMBREORIG$ = TRIM$(Mid$(List2.List(List2.ListIndex), 65))
        CONDI$ = "codi_cli = '" & Label26 & "' AND NOMASIGNADO='" & NOMBRE$ & "'"
        XXX$ = VALOBADA("DOCADCRM", "FILE_DATA", CONDI$)
        ARCHIVOTMP$ = RUTATMP$ & "\" & NOMBREORIG$
        Call SAVEFILE(ARCHIVOTMP$, XXX$)
        Call MUESTRADOC3(ARCHIVOTMP$)
    End If
    Command11.Enabled = True
    
End Sub

Private Sub Command2_Click()
   Unload Me
   On Error Resume Next
   Hide
   On Error GoTo 0
End Sub

Private Sub Command24_Click()
   '
   Command24.Enabled = False
   If List2.ListIndex > -1 Then
        NOMBRE$ = TRIM$(Mid$(List2.List(List2.ListIndex), 1, 64))
        CONDI$ = "codi_cli = '" & Label26 & "' AND NOMASIGNADO='" & NOMBRE$ & "'"
        Call BORRAREGISTROS("DOCADCRM", CONDI$, REG&)
        List2.RemoveItem (List2.ListIndex)
        List2.Refresh
   End If
   Command24.Enabled = True
    '
End Sub

Private Sub Command6_Click()
    '
    Command6.Enabled = False
    If TRIM(Label26) <> "" Then
      F$ = File1.Path + "\" + Label27
      F2$ = SAFETEXT$(Label27)
      F3$ = TRIM$(Text2)
      XXX$ = LOADFILE$(F$)
      If XXX$ <> "" And F3$ <> "" Then
10       NUMAPROX$ = RANDSTRING$(-12)
         If CantRegistros("DOCADCRM", "CODI_CLI = '" & NUMAPROX$ & "'", "NOMESS") > 0 Then GoTo 10
         Call CreaRegistro("DOCADCRM", "CODI_CLI", NUMAPROX$, "NOMESS")
         '
         CONDI$ = "CODI_CLI = '" & NUMAPROX$ & "'"
         Call ACTUREGISTRO("DOCADCRM", "NomAsignado", "CODI_CLI = '" & NUMAPROX$ & "'", F3$, RAFE&, "NOMESS")
         Call ACTUREGISTRO("DOCADCRM", "FILE_NAME", "CODI_CLI = '" & NUMAPROX$ & "'", F2$, RAFE&, "NOMESS")
         Call ACTUREGISTRO("DOCADCRM", "FILE_DATA", "CODI_CLI = '" & NUMAPROX$ & "'", XXX$, RAFE&, "NOMESS")
         Call ACTUREGISTRO("DOCADCRM", "CODI_CLI", "CODI_CLI = '" & NUMAPROX$ & "'", Label26, RAFE&, "NOMESS")
         '
         XX$ = Space$(256)
         Call REPLA(XX$, AJUSTI$(Text2, 64), 1, 64)
         Call REPLA(XX$, TRIM$(Label27), 65, 128)
         List2.AddItem XX$
         List2.Refresh
       End If
   End If
   '
   DoEvents
   Command6.Enabled = True
   '
End Sub

Private Sub Command7_Click()
    '
    Command7.Enabled = False
    ARDOC$ = Label29 + "\" + Label27
    Call MUESTRADOC3(ARDOC$)
    Command7.Enabled = True
   
End Sub


Private Sub Dir1_Change()
     '
     Label29.Caption = Dir1.Path
     File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
     '
     On Error GoTo 20
     Dir1.Path = Drive1.Drive
     GoTo 99
     '
20   Resume 21
21   On Error GoTo 0
     Call MENSERR(24, "Unidad no Disponible.")
     Drive1.Drive = Left$(CurDir$, 2)
     '
99   On Error GoTo 0
End Sub

Private Sub File1_Click()
     '
     Label27 = File1.filename
     Label29.Caption = Dir1.Path
     '
End Sub

Private Sub Form_Load()
     '
     TTXX$ = "CARPETMP"
     On Error Resume Next
     GetAttr (LUCOM$ & TTXX$)
     If Err Then
        MkDir (LUCOM$ & TTXX$)
     End If
     On Error GoTo 0
     RUTATMP$ = LUCOM$ & TTXX$
     '
End Sub

Private Sub Label26_Change()
    Label3 = "": List2.Clear
    Label3 = VALOBADA("BSEGCOM", "Raso_Cli", "CODI_CLI = '" & Label26 & "'")
    CONDI$ = "Codi_Cli = '" & Label26 & "'"
    Call CARGALISTA(List2, "DOCADCRM", "NOMASIGNADO/A64;FILE_NAME/A192", CONDI$, "NOMESS")
End Sub

Sub MUESTRADOC3(ARDOC$)
    '
    Call ShellExecute(Me.hwnd, "Open", ARDOC$, "", "", 1)
    '
End Sub



Private Sub List2_DblClick()
    Call Command11_Click
End Sub
