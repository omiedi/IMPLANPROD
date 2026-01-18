VERSION 5.00
Begin VB.Form DATFAEX07 
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Direccionamiento Impresiones"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2625
      TabIndex        =   1
      Top             =   2205
      Width           =   1485
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
      Height          =   330
      Left            =   525
      TabIndex        =   0
      Top             =   2205
      Width           =   1485
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   525
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1575
      Width           =   3585
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   525
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   630
      Width           =   3585
   End
   Begin VB.Label Label3 
      Height          =   435
      Left            =   3675
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Remitos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   5
      Top             =   1260
      Width           =   1485
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Facturas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   4
      Top             =   315
      Width           =   1485
   End
End
Attribute VB_Name = "DATFAEX07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Label3 = "OK"
    Hide
End Sub

Private Sub Command2_Click()
    Label3 = ""
    Hide
End Sub

Private Sub Form_Load()
    Combo1.Clear
    Dim PRX As Printer
    On Error Resume Next
    Combo1.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo1.AddItem PRX.DeviceName
      If PRX.DeviceName = PORTUFAC$ Then
        Combo1.TEXT = PORTUFAC$
      End If
    Next
    '
    Combo2.Clear
    On Error Resume Next
    Combo2.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo2.AddItem PRX.DeviceName
      If PRX.DeviceName = PORTUREM$ Then
        Combo2.TEXT = PORTUREM$
      End If
    Next
    
End Sub
