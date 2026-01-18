VERSION 5.00
Object = "{CA1CF9C1-EEDB-4CC8-A95A-CC2E923FC6D7}#1.0#0"; "DataMatrixAx.dll"
Begin VB.Form BIOPRO11 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Expedicion"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   7995
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta DM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   3600
      TabIndex        =   10
      Top             =   120
      Width           =   3375
      Begin DataMatrixAxLibCtl.DataMatrixCtrl DataMatrixCtrl1 
         Height          =   2595
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   2835
         _cx             =   5001
         _cy             =   4577
         DataToEncode    =   "1234567"
         PreferredFormat =   0
         EncodingMode    =   4
         Orientation     =   0
         Transparent     =   0   'False
         BackColor       =   16777215
         ForeColor       =   0
         QuietZone       =   -1  'True
         AlignH          =   1
         AlignV          =   1
         StructuredAppend=   0   'False
         StructuredAppendPos=   1
         StructuredAppendTotal=   2
         StructuredAppendID=   1
         QuietZoneSize   =   1
         StretchBarcode  =   -1  'True
      End
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      ItemData        =   "BIOPRO11.frx":0000
      Left            =   465
      List            =   "BIOPRO11.frx":0007
      TabIndex        =   5
      Top             =   185
      Width           =   2910
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   465
      TabIndex        =   0
      Top             =   3480
      Width           =   2910
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   6945
      Left            =   7080
      ScaleHeight     =   6885
      ScaleWidth      =   1125
      TabIndex        =   2
      Top             =   -240
      Width           =   1185
      Begin VB.Frame Frame17 
         BackColor       =   &H00808080&
         Caption         =   "Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   630
         Left            =   90
         TabIndex        =   7
         Top             =   2880
         Width           =   630
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   105
            TabIndex        =   8
            Top             =   270
            Width           =   435
         End
      End
      Begin VB.CommandButton Command29 
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
         Height          =   690
         Left            =   105
         Picture         =   "BIOPRO11.frx":0018
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Borra Item Activo"
         Top             =   1115
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Height          =   690
         Left            =   105
         Picture         =   "BIOPRO11.frx":0322
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Limpia la Lista"
         Top             =   315
         Width           =   650
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "BIOPRO11.frx":062C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Imprime Etiqueta Resumen"
         Top             =   1915
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "BIOPRO11.frx":0C96
         Top             =   3600
         Width           =   1515
      End
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
      Height          =   2760
      Left            =   480
      TabIndex        =   1
      Top             =   450
      Width           =   2910
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   ">>>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   40
      TabIndex        =   9
      Top             =   3520
      Width           =   495
   End
End
Attribute VB_Name = "BIOPRO11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command13_Click()
    '
    Command13.Enabled = False
    Printer.CurrentX = 2048
    Printer.Print "Etiqueta DataMatrix"
    'DataMatrix-ActiveX will use the Visual Basic Printer object
    Call DataMatrixCtrl1.SetPrinterHDC(Printer.hDC)

    'print the barcodes
    Call DataMatrixCtrl1.DrawMatrixToSize(40, 10, 30, 30, dmMM)
    
    '
    Printer.CurrentY = 5000
    Printer.Print
    TTXX$ = ""
    For I% = 1 To List1.ListCount
       TTXX$ = TTXX$ + List1.List(I% - 1) + vbCrLf
    Next I%
    On Error Resume Next
    Printer.Print TTXX$
    Printer.EndDoc
    If Err Then
       On Error GoTo 0
       MsgBox "Error de Conexion con Impresora - Revise y Reintente"
    End If
    On Error GoTo 0
    Command13.Enabled = True
    '
End Sub

Private Sub Command2_Click()
  '
  Command2.Enabled = False
  List1.Clear
  Command2.Enabled = True '
End Sub

Private Sub Command29_Click()
    '
    If List1.ListIndex >= 0 Then
      List1.RemoveItem (List1.ListIndex)
      Label18 = List1.ListCount
      Call GENERADM
      Text1.SetFocus
    End If
    '
End Sub

Private Sub Form_Load()
    '
    On Error Resume Next
    Call GENERADM
    Text1.SetFocus
    On Error GoTo 0
    '
End Sub

Private Sub Label18_Change()
    If Label18 >= 5 Then
       MsgBox "Limite Alcanzado de " & Label18 & " Unidades - Imprima Etiqueta Resumen"
    End If
End Sub

Private Sub List1_Click()
    '
    Text1.SetFocus
    '
End Sub

Private Sub Text1_change()
    '
    If Len(Text1) >= 19 Then
       List1.AddItem Left$(Text1, 19)
       Label18 = List1.ListCount
       Call GENERADM
       List1.Refresh
       Text1 = ""
       Text1.SetFocus
    End If
    '
99 End Sub

Sub GENERADM()
    '
    TTXX$ = ""
    For I% = 1 To List1.ListCount
       TTXX$ = TTXX$ + List1.List(I% - 1) + vbCrLf
    Next I%
    '
    DataMatrixCtrl1.DataToEncode = TTXX$

End Sub

