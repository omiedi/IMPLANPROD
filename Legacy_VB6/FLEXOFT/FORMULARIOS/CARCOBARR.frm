VERSION 5.00
Begin VB.Form CARCOBARR 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturacion por Lista de Empaque"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   2085
      Left            =   210
      Sorted          =   -1  'True
      TabIndex        =   29
      Top             =   6825
      Width           =   8415
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   210
      TabIndex        =   0
      Top             =   3570
      Width           =   7365
   End
   Begin VB.PictureBox Picture1 
      Height          =   435
      Left            =   210
      ScaleHeight     =   375
      ScaleWidth      =   7305
      TabIndex        =   11
      Top             =   3255
      Width           =   7365
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant"
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
         Left            =   5590
         TabIndex        =   14
         Top             =   50
         Width           =   1275
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
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
         Left            =   2060
         TabIndex        =   13
         Top             =   50
         Width           =   2010
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Left            =   50
         TabIndex        =   12
         Top             =   50
         Width           =   1275
      End
   End
   Begin VB.CommandButton Command3 
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
      Height          =   855
      Left            =   7790
      Picture         =   "CARCOBARR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command29 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7770
      Picture         =   "CARCOBARR.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7770
      Picture         =   "CARCOBARR.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4095
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3570
      TabIndex        =   3
      Top             =   2205
      Width           =   175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6300
      MaxLength       =   4
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   2010
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1050
      TabIndex        =   2
      Text            =   " "
      Top             =   2205
      Width           =   2535
   End
   Begin VB.Label Label19 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   360
      Left            =   1050
      TabIndex        =   28
      Top             =   2625
      Width           =   6525
   End
   Begin VB.Label Label18 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   6510
      TabIndex        =   27
      Top             =   1155
      Width           =   1110
   End
   Begin VB.Label Label17 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   6510
      TabIndex        =   26
      Top             =   210
      Width           =   1110
   End
   Begin VB.Label Label16 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   1470
      TabIndex        =   25
      Top             =   1155
      Width           =   3735
   End
   Begin VB.Label Label15 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   2310
      TabIndex        =   24
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   1470
      TabIndex        =   23
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   1470
      TabIndex        =   22
      Top             =   525
      Width           =   3735
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   1470
      TabIndex        =   21
      Top             =   210
      Width           =   3735
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0001"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   420
      TabIndex        =   20
      Top             =   1155
      Width           =   900
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   330
      Left            =   5670
      TabIndex        =   19
      Top             =   255
      Width           =   750
   End
   Begin VB.Label Label12 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   5790
      TabIndex        =   18
      Top             =   1155
      Width           =   585
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6360
      TabIndex        =   17
      Top             =   1155
      Width           =   150
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Factura:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   5040
      TabIndex        =   16
      Top             =   765
      Width           =   1590
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   330
      Left            =   210
      TabIndex        =   15
      Top             =   915
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   7245
      Picture         =   "CARCOBARR.frx":075E
      Top             =   6030
      Width           =   2010
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U-M:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3675
      TabIndex        =   7
      Top             =   2310
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
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
      Left            =   0
      TabIndex        =   6
      Top             =   2310
      Width           =   960
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cant:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5355
      TabIndex        =   5
      Top             =   2310
      Width           =   855
   End
   Begin VB.Label UMEDI 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4620
      TabIndex        =   4
      Top             =   2205
      Width           =   645
   End
End
Attribute VB_Name = "CARCOBARR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
    Call SELECHO("MASTER")
    CODD$ = VALACT1$("MASTER")
    COD$ = CODD$
    Text1.TEXT = COD$
    '
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
End Sub

Private Sub Command27_Click()
   If List1.ListCount > 0 Then
     If COMALTER%("Cancela Sin Registrar ?\\No - Restaurar\Si - Cancelar") = 2 Then
       List1.Clear
     End If
   End If
End Sub

Private Sub Command29_Click()
   Hide
End Sub

Private Sub Command3_Click()
   If List1.ListCount > 0 Then
     If COMALTER%("Cancela Sin Registrar ?\\No - Restaurar\Si - Cancelar") <> 2 Then
       Exit Sub
     End If
   End If
   List1.Clear
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Form_Load()
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
   End If
   '
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
   '
End Sub

Private Sub List1_DblClick()
   REAC = List1.ListIndex
   If REAC >= 0 Then
     If REAC < List1.ListCount Then
       List1.RemoveItem REAC
       If REAC < List1.ListCount Then
           List1.ListIndex = REAC
         Else
           List1.ListIndex = List1.ListCount - 1
       End If
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
     End If
   End If
End Sub

Private Sub Text1_Change()
   If TRIM$(Text1) = "" Then
       Label19.Caption = ""
       UMEDI.Caption = ""
     Else
       COD$ = TRIM$(Text1.TEXT)
       CI% = CODINT%(COD$)
       UMEDI.Caption = UNIMED$(CI%)
       Label19 = DESCRIT$(CI%)
   End If
End Sub

Private Sub Text1_KEYPRESS(KeyAscii As Integer)
    '
    If KeyAscii% = 13 Then
      Label19.Caption = ""
      UMEDI.Caption = ""
      '
      If Len(Text1.TEXT) >= 3 Then
        TTXX$ = UCase$(Text1.TEXT)
        If CODINT%(TTXX$) > 0 Then
          COD$ = TRIM$(TTXX$)
          CI% = CODINT%(COD$)
          UMEDI.Caption = UNIMED$(CI%)
          Label19 = DESCRIT$(CI%)
          CAN = Val(Text2.TEXT)
          For U& = 1 To List1.ListCount
            If TRIM$(Left$(List1.List(U& - 1), 16)) = COD$ Then
              TTXX$ = List1.List(U& - 1)
              CAN = CAN + Val(Mid$(List1.List(U& - 1), 52, 6))
              List1.RemoveItem (U& - 1)
              Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 52, 6)
              List1.AddItem TTXX$, U& - 1
              List1.ListIndex = U& - 1
              GoTo 20
            End If
          Next U&
          '
          TTXX$ = Space$(64)
          Call REPLA(TTXX$, AJUSTI$(COD$, 16), 1, 16)
          Call REPLA(TTXX$, DESCRIT$(CI%), 18, 33)
          Call REPLA(TTXX$, CSTRING$(MKS$(CAN), 3, 6, 0, 2), 52, 6)
          List1.AddItem TTXX$
          List1.ListIndex = List1.ListCount - 1
20        Text1.TEXT = ""
          Text2.TEXT = 1
        End If
      End If
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
    End If
    '
End Sub
'

Private Sub Text1_DblClick()
    Call SELECHO("MASTER")
    CODD$ = VALACT1$("MASTER")
    COD$ = CODD$
    Text1.TEXT = COD$
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
   End If
End Sub
