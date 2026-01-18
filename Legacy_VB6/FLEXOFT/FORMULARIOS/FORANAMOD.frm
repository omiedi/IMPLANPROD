VERSION 5.00
Begin VB.Form FORANAMOD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Analisis de Costos de Mano de Obra"
   ClientHeight    =   5145
   ClientLeft      =   150
   ClientTop       =   405
   ClientWidth     =   11880
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll1 
      Height          =   225
      Left            =   0
      TabIndex        =   1
      Top             =   4935
      Width           =   11800
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11900
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "FORANAMOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub HScroll1_Change()
   FORANAMOD.List1.Left = (-1) * HScroll1.Value
   FORANAMOD.List1.Width = HScroll1.Value + FORANAMOD.Width
End Sub

Private Sub List1_Click()
   Static INCREFON
   If INCREFON = 0 Then INCREFON = (-1)
   If FORANAMOD.List1.FontSize <= 7 Then INCREFON = 1
   If FORANAMOD.List1.FontSize >= 10 Then INCREFON = -1
   '
   FORANAMOD.List1.FontSize = FORANAMOD.List1.FontSize + INCREFON
   FORANAMOD.List1.Height = FORANAMOD.HScroll1.Top - 50
   If FORANAMOD.List1.FontSize < 7 Then
       FORANAMOD.List1.FontBold = True
     Else
       FORANAMOD.List1.FontBold = False
   End If
   '
   FORANAMOD.List1.Left = 0
   FORANAMOD.HScroll1.Min = 0
   FORANAMOD.HScroll1.Value = 0
   FORANAMOD.TEPRUEBA.Font = FORANAMOD.List1.Font
   FORANAMOD.TEPRUEBA.FontSize = FORANAMOD.List1.FontSize
   LOMAX& = FORANAMOD.TEPRUEBA.Width
   If LOMAX > FORANAMOD.Width - 250 Then
        FORANAMOD.HScroll1.Max = LOMAX + 250 - FORANAMOD.Width
      Else
        FORANAMOD.HScroll1.Max = 1
   End If
   On Error Resume Next
   FORANAMOD.HScroll1.SmallChange = FORANAMOD.HScroll1.Max / 80
   FORANAMOD.HScroll1.LargeChange = FORANAMOD.HScroll1.Max / 8
   On Error GoTo 0
   FORANAMOD.List1.Refresh
   FORANAMOD.HScroll1.Refresh
End Sub
