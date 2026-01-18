VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SELDOC09 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de Documento"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
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
      Left            =   4365
      TabIndex        =   7
      Top             =   690
      Width           =   165
   End
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
      Height          =   315
      Left            =   1230
      TabIndex        =   4
      Top             =   165
      Width           =   5910
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   7200
      Picture         =   "SELDOC09.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Transfiere a Central de Pesadas"
      Top             =   165
      Width           =   675
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "Nombre:"
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
      TabIndex        =   6
      Top             =   720
      Width           =   840
   End
   Begin VB.Label Label5 
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
      Height          =   300
      Left            =   1230
      TabIndex        =   5
      Top             =   690
      Width           =   3135
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000004&
      Caption         =   "Ruta:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   840
   End
   Begin VB.Label Label4 
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
      Height          =   300
      Left            =   1215
      TabIndex        =   1
      Top             =   1035
      Width           =   5925
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000004&
      Caption         =   "Referencia:"
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
      TabIndex        =   0
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "SELDOC09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub eio()
  CD1.ShowOpen
  Label5 = ""
  Label4 = ""
  Label5 = CD1.FileTitle
  Label4 = CD1.filename
  
  'Agregar Documento

'
'
'   If Text1 = "" Then
'      MsgBox ("Falta Seleccionar Archivo")
'      Command1.SetFocus
'      Exit Sub
'   End If
'   '
'   Dim REGISTRO$
'   REGISTRO$ = Space(256)
'   '
'   URE& = ULTREG&("DOCUMEN")
'   XX$ = REGLEIDO("DOCUMEN", URE&)
'   COD1% = CVI(Mid$(XX$, 1, 2))
'   COD1% = COD1% + 1
'
'   URE& = URE& + 1
'
'   Call REPLA(REGISTRO$, MKI(COD1%), 1, 2) 'CODDOCU
'   Call REPLA(REGISTRO$, Text1, 3, 64) 'DENOMINACION
'   Call REPLA(REGISTRO$, Label4, 68, 128) 'UBICACION
'   Call REGAPP("DOCUMEN", REGISTRO$)
'   '
'   FILA$ = Space(128)
'   Call REPLA(FILA$, SAFETEXT(COD1%), 200, 10) 'CODDOCU
'   Call REPLA(FILA$, Text1, 1, 64) 'DENOMINACION
'   List1.AddItem FILA$
'   List2.AddItem FILA$
'   Text1 = ""
'   Label4 = ""
End Sub

Private Sub Command1_Click()
  CD1.ShowOpen
  Label5 = CD1.FileTitle
  Label4 = CD1.filename
End Sub

Private Sub Command2_Click()
  'Agregar Documento



   If Text1 = "" Then
      MsgBox ("Falta Ingresar Referencia de Archivo")
      Command1.SetFocus
      Exit Sub
   End If
   '
   Dim REGISTRO$
   REGISTRO$ = Space(256)
   '
   URE& = ULTREG&("DOCUMEN")
   XX$ = REGLEIDO("DOCUMEN", URE&)
   COD1% = CVI(Mid$(XX$, 1, 2))
   COD1% = COD1% + 1

   URE& = URE& + 1

   Call REPLA(REGISTRO$, MKI(COD1%), 1, 2) 'CODDOCU
   Call REPLA(REGISTRO$, Text1, 3, 64) 'DENOMINACION
   Call REPLA(REGISTRO$, Label4, 68, 128) 'UBICACION
   Call REGAPP("DOCUMEN", REGISTRO$)
   '
   FILA$ = Space(128)
   Call REPLA(FILA$, SAFETEXT(COD1%), 200, 10) 'CODDOCU
   Call REPLA(FILA$, Text1, 1, 64) 'DENOMINACION
   CARDOC09.List1.AddItem FILA$
   
   Text1 = ""
   Label4 = ""
   Label5 = ""
   Hide
End Sub

