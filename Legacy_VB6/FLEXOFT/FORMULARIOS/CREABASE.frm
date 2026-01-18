VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form CREABASE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear Base de Datos"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   9
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3360
      Top             =   4155
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
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
      Height          =   300
      Left            =   6960
      TabIndex        =   8
      Top             =   705
      Width           =   180
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Crear la Base"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4680
      TabIndex        =   7
      Top             =   3885
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
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
      Height          =   345
      Left            =   675
      TabIndex        =   6
      Top             =   3900
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Advertencia"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   180
      TabIndex        =   4
      Top             =   1200
      Width           =   7185
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   195
         TabIndex        =   5
         Top             =   345
         Width           =   6825
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2655
      TabIndex        =   1
      Top             =   240
      Width           =   4470
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "CREABASE.frx":0000
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "CREABASE.frx":0234
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1245
      TabIndex        =   3
      Top             =   705
      Width           =   5730
   End
   Begin VB.Label Label2 
      Caption         =   "Ruta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   2
      Top             =   795
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de la Base"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   345
      TabIndex        =   0
      Top             =   300
      Width           =   1815
   End
End
Attribute VB_Name = "CREABASE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    '
    If TRIM$(Text1) = "" Then
       MsgBox "Falta Ingresar el nombre de la B.Datos"
       Exit Sub
    End If
    '
    If TRIM$(Label3) = "" Then
       MsgBox "Falta Ingresar la Ruta para la B.Datos"
       Exit Sub
    End If
    '
    Hide
    '
End Sub

Private Sub Command3_Click()
  '
  CD1.InitDir = LUDAT$
  CD1.ShowOpen
  '
  If CD1.FileName <> "" Then
     ARCH% = Len(CD1.FileTitle)
     Total% = Len(CD1.FileName)
     Ruta% = Total% - ARCH%
     '
     CADENA$ = Mid(CD1.FileName, 1, Ruta%)
     Label3 = CADENA$
  End If
  
  '
End Sub

Private Sub Form_Load()
  '
  Label4 = "1- Solo un usuario con nivel de Administrador puede crear una Base de Datos. " + Chr(13) + Chr(13)
  Label4 = Label4 + "2- El proceso de creacion de una Base de Datos debe ejecutarse en el equipo servidor SQL." + Chr(13) + Chr(13)
  Label4 = Label4 + "3- La Ruta debe referenciar una carpeta valida en el equipo donde reside el motor SQL SERVER." + Chr(13)
  Label4 = Label4 + "No son Validas unidades de red mapeadas"
  '
  'Call APLICACIONSKINS(Me)
   '
End Sub
