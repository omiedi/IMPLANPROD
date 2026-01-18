VERSION 5.00
Begin VB.Form FORMDOWN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Instalacion de Novedades"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Continuar Sin Actualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   3150
      TabIndex        =   4
      Top             =   4095
      Width           =   2955
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Instalar Ahora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   105
      TabIndex        =   2
      Top             =   4095
      Width           =   2955
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      IntegralHeight  =   0   'False
      ItemData        =   "FORMDOWN.frx":0000
      Left            =   105
      List            =   "FORMDOWN.frx":0002
      TabIndex        =   1
      Top             =   840
      Width           =   6000
   End
   Begin VB.Label Label4 
      Height          =   330
      Left            =   3780
      TabIndex        =   6
      Top             =   4305
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   330
      Left            =   1260
      TabIndex        =   5
      Top             =   4300
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"FORMDOWN.frx":0004
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   105
      TabIndex        =   3
      Top             =   3360
      Width           =   6000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"FORMDOWN.frx":00D0
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   6000
   End
End
Attribute VB_Name = "FORMDOWN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'   Label3 = "BAJAR"
'   Hide
End Sub

Private Sub Command2_Click()
'   Unload Me
End Sub

'Private Sub List3_DblClick()
'    ARNAME$ = UCase$(TRIM$(Left$(List3.TEXT, 32)))
'    If InStr(ARNAME$, ".PDF") > 0 Then
'       On Error Resume Next
'       SHOWPDF.Pdf1.LOADFILE$ (Label4 + ARNAME$)
'       If Err Then
'           MsgBox "ERROR " & Err.Number & " - Imposible Mostrar Archivo '" & Label4 + ARNAME$ & "'."
'           On Error GoTo 0
'         Else
'           On Error GoTo 0
'           SHOWPDF.Show 1
'       End If
'    End If
'End Sub
