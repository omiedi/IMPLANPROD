VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form DIREPRIN 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Direccionamiento de Impresiones"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "SELECCION DE IMPRESORAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3855
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   3585
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1455
         Width           =   3585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Facturas, Notas de Débito y Crédito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   2685
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Remitos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   3480
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2280
      TabIndex        =   1
      Top             =   2240
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   2240
      Width           =   1485
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "DIREPRIN.frx":0000
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   2760
      OleObjectBlob   =   "DIREPRIN.frx":0234
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label3 
      Height          =   435
      Left            =   3675
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "DIREPRIN"
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

    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
    End If

    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
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
