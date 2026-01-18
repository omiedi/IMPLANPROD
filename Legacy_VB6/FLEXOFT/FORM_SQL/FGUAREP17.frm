VERSION 5.00
Begin VB.Form FGUAREP17 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guardar Consulta SQL"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Guardar como Archivo SQF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   105
      TabIndex        =   6
      Top             =   1155
      Width           =   9150
      Begin VB.TextBox Text4 
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
         Left            =   5355
         TabIndex        =   14
         Text            =   " "
         Top             =   420
         Width           =   2745
      End
      Begin VB.CommandButton Command5 
         Caption         =   "."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8085
         TabIndex        =   13
         Top             =   420
         Width           =   175
      End
      Begin VB.CommandButton Command1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8400
         Picture         =   "FGUAREP17.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Actualizar Lista"
         Top             =   420
         Width           =   540
      End
      Begin VB.TextBox Text3 
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
         Left            =   2100
         TabIndex        =   8
         Text            =   " "
         ToolTipText     =   "Doble Click para Guardar como Archivo"
         Top             =   390
         Width           =   2430
      End
      Begin VB.CommandButton Command10 
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
         Left            =   4335
         Picture         =   "FGUAREP17.frx":014A
         TabIndex        =   7
         ToolTipText     =   "Click para Seleccionar Carpeta de Salida"
         Top             =   735
         Width           =   175
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Envio x Mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4725
         TabIndex        =   16
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label18 
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
         Left            =   4725
         TabIndex        =   15
         Top             =   765
         Width           =   4215
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
         TabIndex        =   11
         Top             =   525
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
         Height          =   300
         Left            =   420
         TabIndex        =   10
         ToolTipText     =   "Doble Click para Seleccionar Carpeta de Salida"
         Top             =   765
         Width           =   4110
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
         Left            =   945
         TabIndex        =   9
         Top             =   465
         Width           =   1065
      End
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
      Height          =   280
      Left            =   6300
      TabIndex        =   2
      Text            =   " "
      Top             =   210
      Width           =   2010
   End
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
      Height          =   280
      Left            =   1575
      TabIndex        =   1
      Top             =   630
      Width           =   6735
   End
   Begin VB.CommandButton Command6 
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
      Height          =   735
      Left            =   8505
      Picture         =   "FGUAREP17.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   750
   End
   Begin VB.Label Label1 
      Height          =   225
      Left            =   210
      TabIndex        =   5
      Top             =   210
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Codigo:"
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
      Left            =   5355
      TabIndex        =   4
      Top             =   255
      Width           =   795
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Descripcion:"
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
      Left            =   210
      TabIndex        =   3
      Top             =   705
      Width           =   1275
   End
End
Attribute VB_Name = "FGUAREP17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command10_Click()
   Command10.Enabled = False
   Call Label11_DblClick
   Command10.Enabled = True
End Sub

Private Sub Command6_Click()
    If TRIM$(Text1) = "" Then
        Call MENSERR(24, "Falta Código de Consulta")
        Exit Sub
     ElseIf Len(Text1) > 12 Then
        Call MENSERR(24, "Longitud Excesiva Código de Consulta - Max 12 Pos")
        Exit Sub
    End If
    '
    If TRIM$(Text2) = "" Then
        Call MENSERR(24, "Falta Descripcion de Consulta")
        Exit Sub
     ElseIf Len(Text2) > 128 Then
        Call MENSERR(24, "Longitud Excesiva Descripción de Consulta - Max 128 Pos")
        Exit Sub
    End If
    '
    Label1 = "OK"
    Hide
End Sub

Private Sub Form_Load()
    Label1 = "CANCEL"
End Sub

Private Sub Label11_DblClick()
   If TRIM$(Text3) <> "" Then
     SELCARP07.Show 1
     Label11 = SELCARP07.Label1
   End If
End Sub


Private Sub Text3_DblClick()
    Text3 = Text1
End Sub

