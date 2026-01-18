VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   2595
      Left            =   210
      TabIndex        =   4
      Top             =   3360
      Width           =   6630
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   5565
      TabIndex        =   3
      Top             =   315
      Width           =   1275
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   210
      TabIndex        =   2
      Top             =   2520
      Width           =   5160
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   3045
      TabIndex        =   1
      Top             =   315
      Width           =   2325
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   210
      TabIndex        =   0
      Top             =   315
      Width           =   2325
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
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
      Left            =   5565
      TabIndex        =   5
      Top             =   2835
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Dim NOMARCH$(4096), FEARCH#(4096), RUTARCH$(4096)
   List1.Clear
   List1.AddItem "A:\"
   ILISTA& = 0
   CARCHI% = 0
   '
'GoTo 99
10 Dir1.Path = List1.List(ILISTA&)
   For KKI% = 1 To Dir1.ListCount
     TTXX$ = Dir1.List(KKI% - 1) + "\"
     List1.AddItem TTXX$
     List1.ListIndex = List1.ListCount - 1
   Next KKI%
   '
   If ILISTA& < List1.ListCount - 1 Then
     ILISTA& = ILISTA& + 1
     GoTo 10
   End If
   '
20 For ILISTA& = 1 To List1.ListCount
     File1.Path = List1.List(ILISTA& - 1)
     DoEvents
     File1.Pattern = "*.*"
     For KKI% = 1 To File1.ListCount
       FINAME$ = Trim$(UCase$(File1.List(KKI% - 1)))
       If FINAME$ <> "" Then
         ACOMPLE$ = List1.List(ILISTA& - 1) + FINAME$
         If FileLen(ACOMPLE$) > 0 Then
           LARCHI& = FileLen(ACOMPLE$)
           CARCHI% = CARCHI% + 1
           NOMARCH$(CARCHI%) = FINAME$
           RUTARCH$(CARCHI%) = List1.List(ILISTA& - 1)
         End If
       End If
29   Next KKI%
   Next ILISTA&
   '
   Dim BUFGRA As String * 1024
   NX% = FreeFile
   Open "C:\ANOVE\TODOS.FLX" For Random Access Read Write As #NX% Len = 1024
   Dim BUFLE As String * 960
   For KKI% = 1 To CARCHI%
     IDENTI$ = Mid$(RUTARCH$(KKI%) + NOMARCH$(KKI%), 4)
     N1% = FreeFile
     Open RUTARCH$(KKI%) + NOMARCH$(KKI%) For Random Access Read As #N1% Len = 960
     CARELE& = 1 + Int((FileLen(RUTARCH$(KKI%) + NOMARCH$(KKI%)) - 1) / 960)
     For JJ& = 1 To CARELE&
       Get #N1%, JJ&, BUFLE$
       TTXX$ = String$(1024, 0)
       Mid$(TTXX$, 64) = IDENTI$
       Mid$(TTXX$, 65, 960) = BUFLE$
       BUFGRA$ = TTXX$
       KK& = KK& + 1
       Put #NX%, KK&, BUFGRA$
     Next JJ&
     Close #N1%
   Next KKI%
   Close #NX%
   End
   '
End Sub

