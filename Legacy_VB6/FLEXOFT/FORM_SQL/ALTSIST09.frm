VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ALTSIST09 
   BackColor       =   &H00EEEEDD&
   Caption         =   "Alta de Sistemas / Lineas"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004040&
      Height          =   4635
      Left            =   6360
      ScaleHeight     =   4575
      ScaleWidth      =   1320
      TabIndex        =   13
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command21 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   90
         Picture         =   "ALTSIST09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Visualizar Archivo POE en Formato PDF"
         Top             =   720
         Width           =   600
      End
      Begin VB.CommandButton Command10 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   90
         Picture         =   "ALTSIST09.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Guarda Sistema / Linea"
         Top             =   1920
         Width           =   600
      End
      Begin VB.CommandButton command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   90
         Picture         =   "ALTSIST09.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   120
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "ALTSIST09.frx":059E
         Top             =   1440
         Width           =   1515
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00EEEEDD&
      Caption         =   "Sistema / Linea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2505
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6195
      Begin VB.Frame Frame1 
         BackColor       =   &H00EEEEDD&
         Caption         =   "Fecha de Inicio"
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
         Left            =   4320
         TabIndex        =   14
         Top             =   360
         Width           =   1695
         Begin VB.TextBox Text14 
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
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton Command7 
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
            Left            =   1290
            TabIndex        =   3
            Top             =   240
            Width           =   175
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
         Height          =   300
         Left            =   1560
         MaxLength       =   48
         TabIndex        =   6
         Top             =   1800
         Width           =   4455
      End
      Begin VB.TextBox Text6 
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
         Left            =   4305
         MaxLength       =   48
         TabIndex        =   5
         Top             =   1260
         Width           =   1695
      End
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
         Height          =   300
         Index           =   1
         Left            =   1550
         MaxLength       =   48
         TabIndex        =   4
         Top             =   1260
         Width           =   1710
      End
      Begin VB.TextBox Text23 
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
         Left            =   1550
         MaxLength       =   48
         TabIndex        =   1
         Top             =   600
         Width           =   2550
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "P.O.E.:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1845
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Revision:"
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
         Index           =   1
         Left            =   2880
         TabIndex        =   11
         Top             =   1305
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Criticidad:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1305
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Denominacion:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   645
         Width           =   1335
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   1155
      OleObjectBlob   =   "ALTSIST09.frx":24C0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2730
      OleObjectBlob   =   "ALTSIST09.frx":253A
      Top             =   0
   End
End
Attribute VB_Name = "ALTSIST09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command21_Click()
10  POEPDF$ = TRIM$(InputBox("Ingrese Nº de P.O.E.", "Busqueda de P.O.E."))
    If POEPDF$ = "" Then Exit Sub
    On Error Resume Next
    UBICCC$ = LUDAT$ + POEPDF$ + ".pdf"
    PPOOEE% = ShellExecute(hwnd, "open", UBICCC$, vbNullString, vbNullString, SW_SHOWNORMAL)
    If PPOOEE% <> 42 Then
        Call MENSERR(24, "Nombre de Archivo no Encontrado.\Verifique que el Nombre Ingresado sea el Correcto.")
        On Error GoTo 0
        GoTo 10
    End If
    On Error GoTo 0
    '
End Sub

Private Sub Command7_Click()
Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text14.TEXT = VALACT1$("SELFECHA")
   End If
End Sub


Private Sub Form_Load()
   Call APLICACIONSKINS(Me)

End Sub
