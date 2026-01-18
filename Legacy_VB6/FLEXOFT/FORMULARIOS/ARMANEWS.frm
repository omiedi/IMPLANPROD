VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form ARMANEWS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Armar Mail Circular"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   3360
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   2625
      Top             =   0
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pausa"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   640
      Left            =   8000
      TabIndex        =   39
      Top             =   3990
      Width           =   750
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   40
         Text            =   "10"
         Top             =   250
         Width           =   435
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   330
      Left            =   210
      TabIndex        =   33
      Top             =   4095
      Width           =   6105
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Flexoft"
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
         Left            =   5145
         TabIndex        =   38
         Top             =   105
         Width           =   960
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Clientes"
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
         Left            =   3885
         TabIndex        =   37
         Top             =   105
         Width           =   1065
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ag.Telef."
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
         Left            =   2520
         TabIndex        =   36
         Top             =   105
         Width           =   1170
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Seg.Com."
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
         Left            =   1155
         TabIndex        =   35
         Top             =   105
         Width           =   1170
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Todos"
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
         Left            =   0
         TabIndex        =   34
         Top             =   105
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin VB.CommandButton Command3 
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
      Height          =   810
      Left            =   7980
      Picture         =   "ARMANEWS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Limpiar Marcas de Envío"
      Top             =   5490
      Width           =   750
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Envios Anteriores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1400
      Left            =   4095
      TabIndex        =   27
      Top             =   6405
      Width           =   3690
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   -20
         TabIndex        =   28
         Top             =   525
         Width           =   3700
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   1575
         TabIndex        =   30
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label14 
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
         Height          =   255
         Left            =   2730
         TabIndex        =   29
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Suprimidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1400
      Left            =   210
      TabIndex        =   23
      Top             =   6405
      Width           =   3795
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   0
         TabIndex        =   24
         Top             =   525
         Width           =   3795
      End
      Begin VB.Label Label9 
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
         Height          =   255
         Left            =   2835
         TabIndex        =   26
         Top             =   210
         Width           =   750
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   1680
         TabIndex        =   25
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Enviados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1780
      Left            =   4095
      TabIndex        =   17
      Top             =   4515
      Width           =   3690
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   0
         TabIndex        =   18
         Top             =   525
         Width           =   3700
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   1575
         TabIndex        =   22
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Left            =   2730
         TabIndex        =   21
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pendientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1780
      Left            =   210
      TabIndex        =   15
      Top             =   4515
      Width           =   3795
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   -20
         TabIndex        =   16
         Top             =   525
         Width           =   3795
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   1680
         TabIndex        =   20
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   2835
         TabIndex        =   19
         Top             =   210
         Width           =   750
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   7980
      Picture         =   "ARMANEWS.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Interrumpe Transmisión"
      Top             =   6310
      Width           =   750
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   7980
      Picture         =   "ARMANEWS.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Iniciar / Continuar Transmisión"
      Top             =   4660
      Width           =   750
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Selecc.Dest."
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
      Left            =   6405
      TabIndex        =   12
      Top             =   4095
      Width           =   1380
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Selección Archivo Adjunto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   8520
      Begin VB.TextBox Text1 
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
         Left            =   5145
         TabIndex        =   10
         Top             =   420
         Width           =   2850
      End
      Begin VB.FileListBox File1 
         Height          =   2040
         Left            =   4410
         TabIndex        =   6
         Top             =   945
         Width           =   3900
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   315
         TabIndex        =   2
         Top             =   945
         Width           =   3900
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   1470
         TabIndex        =   1
         Top             =   420
         Width           =   2745
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro:"
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
         Left            =   3990
         TabIndex        =   9
         Top             =   470
         Width           =   1065
      End
      Begin VB.Label Label2 
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
         Height          =   330
         Left            =   315
         TabIndex        =   8
         Top             =   470
         Width           =   1065
      End
      Begin VB.Label Label1 
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
         Height          =   405
         Left            =   4410
         TabIndex        =   7
         Top             =   3255
         Width           =   3900
      End
      Begin VB.Label Label10 
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
         Left            =   3885
         TabIndex        =   5
         Top             =   3030
         Width           =   1065
      End
      Begin VB.Label Label11 
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
         Height          =   1035
         Left            =   315
         TabIndex        =   4
         Top             =   2625
         Width           =   3900
      End
      Begin VB.Label Label12 
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
         Left            =   -210
         TabIndex        =   3
         Top             =   2400
         Width           =   1065
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   630
      Top             =   0
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
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Label Label15 
      Height          =   225
      Left            =   7560
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   7380
      Picture         =   "ARMANEWS.frx":0E56
      Top             =   7280
      Width           =   2010
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
End
Attribute VB_Name = "ARMANEWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   Command1.Enabled = True
End Sub

Private Sub Command2_Click()
   Label15.Caption = "STOP"
End Sub

Private Sub Command3_Click()
   '
   Command3.Enabled = False
   If COMALTER%("Está Seguro de Blanquear Lista de Envíos ?\\No, Cancelar\Si, Refrescar") = 2 Then
     Screen.MousePointer = 11
     For REMAI& = 1 To ULTREG&("LIDIRMAI")
       XX$ = REGLEIDO$("LIDIRMAI", REMAI&)
       Call REPLA(XX$, MKI$(0), 511, 2)
       Call GRAREG("LIDIRMAI", XX$, REMAI&)
     Next REMAI&
     Call CIERRARCH("LIDIRMAI")
     Screen.MousePointer = 1
   End If
   Command3.Enabled = True
   Unload Me
   '
End Sub

Private Sub Command4_Click()
   'Timer1.Enabled = True
   Timer2.Enabled = True
End Sub

Private Sub Dir1_Change()
   Label11.Caption = Dir1.Path
   File1.Path = Dir1.Path
   Label1 = ""
End Sub

Private Sub Drive1_Change()
   On Error GoTo 20
   Dir1.Path = Drive1.Drive
   GoTo 99
   '
20 Resume 21
21 On Error GoTo 0
   Call MENSERR(24, "Unidad no Disponible.")
   Drive1.Drive = Left$(CurDir$, 2)
   '
99 On Error GoTo 0
End Sub

Private Sub File1_Click()
  Label1 = File1.filename
End Sub

Private Sub Form_Load()
    '
    EPAC$ = EMPREAC$
    If EPAC$ <> "" Then
      Caption = Caption + " - " + EMPREAC$
    End If
    '
    Refresh
    DoEvents
    '
    Label11.Caption = Dir1.Path
    File1.Path = Dir1.Path
    '
    Screen.MousePointer = 11
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    '
    For U& = 1 To ULTREG&("LIDIRMAI")
      XX$ = REGLEIDO$("LIDIRMAI", U&)
      TTXX$ = AJUSTI$(Mid$(XX$, 5, 60), 80) + Str$(U&)
      If CVI(Mid$(XX$, 255, 2)) > 0 Or TRIM$(Mid$(XX$, 257, 8)) <> "" Then
         ' esta marcado de baja
         List3.AddItem TTXX$
         List3.ListIndex = List3.ListCount - 1
         Label9 = TRIM$(Str$(List3.ListCount))
       ElseIf CVI(Mid$(XX$, 511, 2)) < 0 Then
         ' esta marcado suprimido
         List3.AddItem TTXX$
         List3.ListIndex = List3.ListCount - 1
         Label9 = TRIM$(Str$(List3.ListCount))
       ElseIf CVI(Mid$(XX$, 511, 2)) > 0 Then
         ' ya fue enviado
         List4.AddItem TTXX$
         List4.ListIndex = List4.ListCount - 1
         Label14 = TRIM$(Str$(List4.ListCount))
       Else
         ' agregar a lists de envios
         List1.AddItem TTXX$
         List1.ListIndex = List1.ListCount - 1
         Label4 = TRIM$(Str$(List1.ListCount))
      End If
      DoEvents
    Next U&
    '
    On Error Resume Next
    List1.TopIndex = 0
    List1.ListIndex = -1
    List2.TopIndex = 0
    List2.ListIndex = -1
    List3.TopIndex = 0
    List3.ListIndex = -1
    List4.TopIndex = 0
    List4.ListIndex = -1
    On Error GoTo 0
    '
    Screen.MousePointer = 1

End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Hide
End Sub


Private Sub Label11_Change()
  If Right$(Label11, 1) = "\" Then
    Label11 = Left$(Label11, Len(Label11) - 1)
  End If
End Sub

Private Sub List1_DblClick()
   '
   U& = List1.ListIndex
   REMAI& = Val(Mid$(List1.List(U&), 81))
   '
   XX$ = REGLEIDO$("LIDIRMAI", REMAI&)
   Call REPLA(XX$, MKI$(-1), 511, 2)
   Call GRAREG("LIDIRMAI", XX$, REMAI&)
   Call CIERRARCH("LIDIRMAI")
   '
   List3.AddItem List1.List(U&)
   List1.RemoveItem U&
   Label9 = TRIM$(Str$(List3.ListCount))
   '
End Sub

Private Sub List4_DblClick()
    KU& = List4.ListIndex
    TTXX$ = List4.TEXT
    List4.RemoveItem KU&
    List1.AddItem TTXX$, 0
End Sub

Private Sub Text1_Change()
   File1.Pattern = Text1
End Sub


Private Sub Timer1_Timer()
   Static U&
   '
   If U& < 1 Then
     U& = 0
   End If
   '
   HOII% = HOY(HOS$)
   U& = U& + 1
   If U& > List1.ListCount Then
      Call COMUNI("Fin de Transmisión")
      Timer1.Enabled = False
      Exit Sub
   End If
   '
     List1.ListIndex = U& - 1
     REMAI& = Val(Mid$(List1.List(U& - 1), 81))
     XX$ = REGLEIDO$("LIDIRMAI", REMAI&)
     APENO$ = TRIM$(Mid$(XX$, 65, 40))
     EMPRA$ = TRIM$(Mid$(XX$, 173, 64))
     If InStr(EMPRA$, "!") > 0 Then
       EMPRA$ = REPLACAR$(EMPRA$, "!", " ")
       Call REPLA(XX$, EMPRA$, 173, 64)
       Call GRAREG("LIDIRMAI", XX$, REMAI&)
     End If
     DIREL$ = REPLACEN$(TRIM$(Mid$(XX$, 105, 60)))
     '
     HC% = 0
     If Right$(DIREL$, 1) = "." Then
       DIREL$ = Left$(DIREL$, Len(DIREL$) - 1)
       HC% = 1
     End If
     If Left$(UCase$(DIREL$), 5) = "EMAIL" Then
       DIREL$ = Mid$(DIREL$, 6)
       HC% = 1
     End If
     If Left$(DIREL$, 1) = ";" Then
       DIREL$ = Mid$(DIREL$, 2)
       HC% = 1
     End If
     If HC% = 1 Then
       Call REPLA(XX$, TRIM$(DIREL$), 105, 60)
       Call GRAREG("LIDIRMAI", XX$, REMAI&)
     End If
     '
     MUESTRA$ = ""
     If TRIM$(APENO$) <> "" Then MUESTRA$ = MUESTRA$ + " " + TRIM$(APENO$)
     If TRIM$(EMPRA$) <> "" Then
       If MUESTRA$ <> "" Then
         MUESTRA$ = MUESTRA$ + " -"
       End If
       MUESTRA$ = MUESTRA$ + " " + EMPRA$
     End If
     MUESTRA$ = REPLACEN$(MUESTRA$)
     '
20   TINI = Timer
     RETVAL = Shell("C:\Archivos de Programa\Outlook Express\msimn.exe", 1)
     'abc1% = Shell("C:\Program Files\Outlook Express1\msimn.exe", 1) - PARA EN INGLES
     On Error Resume Next
     AppActivate RETVAL
     If Err > 0 Then
        On Error GoTo 0
        Call RECUERROR(OKEY%)
        If OKEY% > 0 Then GoTo 20
        If COMALTER%("Se ha Producido Error " + TRIM$(Str$(Err)) + "\\Continuar\Terminar") = 1 Then
          On Error GoTo 0
          GoTo 20
        End If
        Call FINAL("")
     End If
     On Error GoTo 0
     '
     DIREL2$ = ""
     PPXX% = InStr(DIREL$, "/")
     If PPXX% > 0 Then
        DIREL2$ = TRIM$(Mid$(DIREL$, PPXX% + 1))
        DIREL$ = TRIM$(Left$(DIREL$, PPXX% - 1))
     End If
     '
     REPITE% = 0
     ESCA$ = ""
     For KKII% = 1 To 5: ESCA$ = ESCA$ + "{ESC}": Next KKII%
     SendKeys "{ESC}", True
     SendKeys "{ESC}", True
     SendKeys "{ESC}", True
     ' ESTO SE SUPRIME. LO QUE HAY QUE HACER ES LO SIGUIENTE:
     ' 1.- OUTLOOK EXPRESS - HERRAMIENTAS - OPCIONES - REDACTAR
     ' 2.- TILDAR CORREO - SELECCIONAR ARCHIVO DEL NEWSLETTER
     ' 3.- CLICK SOBRE ACEPTAR
     'SendKeys "%MUD\\FLEXSERVER\DISCOD\DOCU_COM\NEWSLETTER\ORIGINALES\NEWSLETTER_020707.HTM" + Chr$(13), True ' "{TAB}{TAB}"
     'SendKeys "%MUD\\SERVERNEW\DISCOF\DOCU_COM\NEWSLETTER\ORIGINALES\NEWSLETTER_021007.HTM" + Chr$(13), True ' "{TAB}{TAB}"
'     SendKeys "%ANM" + DIREL$ + "{TAB}{TAB}", True
'        MUESTRA$ = REPLACAR$(MUESTRA$, "(", " ")
'        MUESTRA$ = REPLACAR$(MUESTRA$, ")", " ")
'     SendKeys "Att: " + MUESTRA$ + "%AE" + "%" + "{F4}", True
     TINI = Timer: While Timer < TINI + 3: Wend
     MUESTRA$ = REPLACAR$(MUESTRA$, "(", " ")
     MUESTRA$ = REPLACAR$(MUESTRA$, ")", " ")
     SK$ = "%ANM" + DIREL$ + "{TAB}{TAB}+Att: " + MUESTRA$ + "%AE" + "%" + "{F4}"
     SendKeys SK$, True
     '
     If DIREL2$ <> "" Then GoTo 20
     '
30   XX$ = REGLEIDO$("LIDIRMAI", REMAI&)
     Call REPLA(XX$, MKI$(HOII%), 511, 2)
     Call GRAREG("LIDIRMAI", XX$, REMAI&)
     '
     Label4 = TRIM$(Str$(Val(Label4) - 1))
     List2.AddItem List1.TEXT
     List2.ListIndex = List2.ListCount - 1
     Label6.Caption = TRIM$(Str$(List2.ListCount))
     '
     If Val(Text2) < 99 Then
       If List2.ListCount Mod Val(Text2) = 0 Then
         OPX% = COMALTER%("Opciones de Procesamiento:\\Continuar\Interrumpir")
         If OPX% = 2 Then Label15 = "STOP"
       End If
     End If
     '
     If Label15.Caption = "STOP" Then GoTo 98
     '
     While Timer < TINI + 5
     Wend
     '
   Exit Sub
   '
98 Call CIERRARCH("LIDIRMAI")
   Call BAJALDISCO
   Timer1.Enabled = False
   '
End Sub

Sub RECUERROR(OKX%)
   '
   ' trata de apagar Outlook Express
   OKX% = 0
   TINI = Timer
3   On Error Resume Next
   AppActivate "Outlook Express", False
   If Err Then GoTo 5
   On Error GoTo 0
   SK$ = "%" + "C"
   On Error Resume Next
   SendKeys Chr$(13), True
   On Error GoTo 0
   On Error Resume Next
   SendKeys SK$, True
   On Error GoTo 0
   If Timer < TINI + 10 Then GoTo 3
   '
   ' CIERRA LA BANDEJA DE ENTRADA
5  OKX% = 0
   TINI = Timer
10 On Error Resume Next
   AppActivate "Bandeja de Entrada", False
   If Err Then
      If Timer < TINI + 60 Then GoTo 10
      GoTo 90
   End If
   '
   On Error GoTo 0
   SK$ = "%" + "{F4}"
   On Error Resume Next
   SendKeys SK$, True
   If Err Then
      On Error GoTo 0
      GoTo 90
   End If
   '
   On Error Resume Next
   AppActivate "Bandeja de Entrada", False
   If Err Then
      On Error GoTo 0
      OKX% = 1
      GoTo 90
   End If
   If Timer < TINI + 60 Then GoTo 10
   '
90 End Sub

Private Sub Timer2_Timer()
     '
     Static U&, RETVAL, VUELTA%
     '
     If (IsNull(RETVAL) = True Or IsEmpty(RETVAL) = True) Then GoTo 10 ' ES LA PRIMERA PASADA
     '
     ' VERIFICA SI TODAVÍA SE ESTÁ EJECUTANDO OUTLOOK EXPRESS
     On Error Resume Next
     AppActivate RETVAL
     If Err Then
        ' YA SE CERRO EL OUTLOOK EXPRESS ANTERIOR
        On Error GoTo 0
        GoTo 10
     End If
     If VUELTA% < 3 Then
        SK$ = "%{F4}{ENTER}{ENTER}"     ' para terminar de enviar
        SendKeys SK$, False            ' con false para que lo haga inmediatamente
        VUELTA% = VUELTA% + 1
        Exit Sub
     End If
     '
10   VUELTA% = 0
     '
11   On Error Resume Next
     AppActivate "Mensaje nuevo", False
     If Err Then
        On Error GoTo 0
        GoTo 12
     End If
     On Error GoTo 0
     SK$ = "Att: " + MUESTRA$ + "%AE" + "%" + "{F4}"  ' para terminar de enviar el anterior
     SendKeys SK$, True             ' con false para que lo haga inmediatamente
     If VUEL2% < 3 Then
        VUEL2% = VUEL2% + 1
        GoTo 11
     End If
     '
12   If U& < 1 Then U& = 0
     HOII% = HOY(HOS$)
     U& = U& + 1
     If U& > List1.ListCount Then
        Call COMUNI("Fin de Transmisión")
        Timer2.Enabled = False
        Exit Sub
     End If
     '
     GoSub LEEDATOS
     '
20   TINI = Timer
     RETVAL = Shell("C:\Archivos de Programa\Outlook Express\msimn.exe", 1)
     On Error Resume Next
     AppActivate RETVAL
     If Err > 0 Then
        NERRO = Err.Number
        DERRO$ = Err.Description
        On Error GoTo 0
        If COMALTER%("Se ha Producido Error " + TRIM$(Str$(NERRO)) + "\'" + DERRO$ + "'\\Continuar\Terminar") = 1 Then
          GoTo 20
        End If
        Call FINAL("")
     End If
     On Error GoTo 0
     '
     While Timer < TINI + 2: Openforms = DoEvents: Wend
     SendKeys "{ESC}", True
     SendKeys "{ESC}", True
     SendKeys "{ESC}", True
     '
     DIREL2$ = ""
     PPXX% = InStr(DIREL$, "/")
     If PPXX% > 0 Then
        DIREL2$ = TRIM$(Mid$(DIREL$, PPXX% + 1))
        DIREL$ = TRIM$(Left$(DIREL$, PPXX% - 1))
     End If
     '
     ' 1.- OUTLOOK EXPRESS - HERRAMIENTAS - OPCIONES - REDACTAR
     ' 2.- TILDAR CORREO - SELECCIONAR ARCHIVO DEL NEWSLETTER
     ' 3.- CLICK SOBRE ACEPTAR
     MUESTRA$ = REPLACAR$(MUESTRA$, "(", " ")
     MUESTRA$ = REPLACAR$(MUESTRA$, ")", " ")
     SK$ = "%ANM" + DIREL$ + "{TAB}{TAB}+Att: " + MUESTRA$ + "%AE" + "%" + "{F4}"
     SendKeys SK$, True
     While Timer < TINI + 1: Openforms = DoEvents: Wend
     '
30   XX$ = REGLEIDO$("LIDIRMAI", REMAI&)
     Call REPLA(XX$, MKI$(HOII%), 511, 2)
     Call GRAREG("LIDIRMAI", XX$, REMAI&)
     '
     Label4 = TRIM$(Str$(Val(Label4) - 1))
     List2.AddItem List1.TEXT
     List2.ListIndex = List2.ListCount - 1
     Label6.Caption = TRIM$(Str$(List2.ListCount))
     '
     If Val(Text2) < 99 Then
       If List2.ListCount Mod Val(Text2) = 0 Then
         OPX% = COMALTER%("Opciones de Procesamiento:\\Continuar\Interrumpir")
         If OPX% = 2 Then Label15 = "STOP"
       End If
     End If
     '
     If Label15.Caption = "STOP" Then GoTo 98
     '
   Exit Sub
   '
98 Call CIERRARCH("LIDIRMAI")
   Call BAJALDISCO
   Timer2.Enabled = False
   Exit Sub
   '
   '
LEEDATOS:
    List1.ListIndex = U& - 1
    REMAI& = Val(Mid$(List1.List(U& - 1), 81))
    XX$ = REGLEIDO$("LIDIRMAI", REMAI&)
    APENO$ = TRIM$(Mid$(XX$, 65, 40))
    EMPRA$ = TRIM$(Mid$(XX$, 173, 64))
    If InStr(EMPRA$, "!") > 0 Then
      EMPRA$ = REPLACAR$(EMPRA$, "!", " ")
      Call REPLA(XX$, EMPRA$, 173, 64)
      Call GRAREG("LIDIRMAI", XX$, REMAI&)
    End If
    DIREL$ = REPLACEN$(TRIM$(Mid$(XX$, 105, 60)))
    '
    HC% = 0
    If Right$(DIREL$, 1) = "." Then
      DIREL$ = Left$(DIREL$, Len(DIREL$) - 1)
      HC% = 1
    End If
    If Left$(UCase$(DIREL$), 5) = "EMAIL" Then
      DIREL$ = Mid$(DIREL$, 6)
      HC% = 1
    End If
    If Left$(DIREL$, 1) = ";" Then
      DIREL$ = Mid$(DIREL$, 2)
      HC% = 1
    End If
    If HC% = 1 Then
      Call REPLA(XX$, TRIM$(DIREL$), 105, 60)
      Call GRAREG("LIDIRMAI", XX$, REMAI&)
    End If
    '
    MUESTRA$ = ""
    If TRIM$(APENO$) <> "" Then MUESTRA$ = MUESTRA$ + " " + TRIM$(APENO$)
    If TRIM$(EMPRA$) <> "" Then
      If MUESTRA$ <> "" Then
        MUESTRA$ = MUESTRA$ + " -"
      End If
      MUESTRA$ = MUESTRA$ + " " + EMPRA$
    End If
    MUESTRA$ = REPLACEN$(MUESTRA$)
Return
'
End Sub
