VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FASPECTO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Aspecto"
   ClientHeight    =   2340
   ClientLeft      =   1755
   ClientTop       =   1710
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   2625
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   9
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   3045
      OleObjectBlob   =   "FASPECTO.frx":0000
      TabIndex        =   8
      Top             =   945
      Visible         =   0   'False
      Width           =   1485
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3465
      OleObjectBlob   =   "FASPECTO.frx":0060
      Top             =   945
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   315
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "FASPECTO.frx":0294
      Top             =   840
      Width           =   6000
   End
   Begin VB.CommandButton Command2 
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
      Height          =   675
      Left            =   6720
      Picture         =   "FASPECTO.frx":0296
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guardar "
      Top             =   945
      Width           =   705
   End
   Begin VB.Frame Frame1 
      Caption         =   "Configuración de Aspecto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   105
      TabIndex        =   2
      Top             =   210
      Width           =   6375
      Begin VB.CommandButton Command5 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5880
         TabIndex        =   7
         ToolTipText     =   "Limpiar Selección"
         Top             =   210
         Width           =   330
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Examinar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4305
         Picture         =   "FASPECTO.frx":05A0
         TabIndex        =   5
         ToolTipText     =   "Selección de Ruta de Archivo "
         Top             =   210
         Width           =   1380
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   5070
      Left            =   6600
      ScaleHeight     =   5010
      ScaleWidth      =   1965
      TabIndex        =   0
      Top             =   -360
      Visible         =   0   'False
      Width           =   2025
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   3
         Top             =   2100
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   4
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "FASPECTO.frx":08AA
         Top             =   2310
         Width           =   1515
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FASPECTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command2_Click()
    Command2.Enabled = False
    'VALIDO QUE LA RUTA EXIST
    Ruta$ = TRIM$(Text1)
    If EXISTE%(Ruta$) < 1 And Ruta$ <> "" Then
       Call COMUNI("La Ruta Elegida No Es Válida")
       GoTo 99
    End If
    '
    'VALIDA QUE LA RUTA NO SUPERE LOS 512 CARACTERES
    If Len(Ruta$) > 512 Then
       Call COMUNI("La Ruta Seleccionada Excede la Longitud Permitida")
       GoTo 99
    End If
    '
    Call REPLA(TX$, Ruta$, 1, 512)
    '
    'GUARDO EN WINDOWS
    NARCHI1% = FILEOPEN%(LUBAS + "FLEX_SKIN.CFG", 2, 512) ' ABRE EL ARCHIVO DE ESCRITURA
    Print #NARCHI1%, TX$ ' ESCRIBE EN EL ARCHIVO
    Close #NARCHI1%
    '
    'GUARDO EN DATOS
    NARCHI2% = FILEOPEN%(LUDAT$ + "FLEX_SKIN.CFG", 2, 512) ' ABRE EL ARCHIVO DE ESCRITURA
    Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO
    Close #NARCHI2%
    

99  Command2.Enabled = True
End Sub

Private Sub Command4_Click()
    '
    Command4.Enabled = False
    CD1.FILTER = "(*.SKN)|*.SKN|"
    CD1.ShowOpen
    Ruta$ = TRIM$(CD1.FileName)
    Text1.TEXT = Ruta$
    '

99  Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    
    Command5.Enabled = False
    Text1.TEXT = ""
    Command5.Enabled = True

End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
      Me.Image2.Visible = False
      Me.Frame1.BackColor = &HFCD7B6
    End If
    '
    RUTXX$ = LUBAS$
    If INTERPRE% > 0 Then RUTXX$ = LUDAT$
    '
    'LEO EL ARCHIVO DE CONFIGURACION
    ARCHIVOSKIN$ = TRIM$(Left$(LOADFILE$(RUTXX$ + "FLEX_SKIN.CFG"), 512))

    If EXISTE%(RUTXX$ + "FLEX_SKIN.CFG") < 1 Then Exit Sub
    
    Text1 = ARCHIVOSKIN$
    '
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Text1_Change()
   
   If TRIM$(Text1) = "" Then Exit Sub
   
   RUTXX$ = LUBAS$
   If INTERPRE% > 0 Then RUTXX$ = LUDAT$
   If EXISTE%(RUTXX$ + "FLEX_SKIN.CFG") < 1 Then Exit Sub
   '
   Command2_Click
   Call APLICACIONSKINS(Me)
End Sub
