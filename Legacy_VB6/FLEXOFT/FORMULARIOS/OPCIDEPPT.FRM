VERSION 5.00
Begin VB.Form OPCIDEPPT 
   BackColor       =   &H00E0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demostración FLEXOFT - Opciones"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
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
      Left            =   6735
      Picture         =   "OPCIDEPPT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   210
      Width           =   680
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFC0&
      Height          =   3480
      Left            =   210
      TabIndex        =   0
      Top             =   105
      Width           =   6210
      Begin VB.Line Line16 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   725
         X2              =   725
         Y1              =   2730
         Y2              =   3255
      End
      Begin VB.Line Line15 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   735
         X2              =   1260
         Y1              =   3260
         Y2              =   3260
      End
      Begin VB.Line Line14 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   1270
         X2              =   1270
         Y1              =   2730
         Y2              =   3255
      End
      Begin VB.Line Line13 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   735
         X2              =   1260
         Y1              =   2730
         Y2              =   2730
      End
      Begin VB.Line Line12 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   735
         X2              =   1260
         Y1              =   1890
         Y2              =   1890
      End
      Begin VB.Line Line11 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   1270
         X2              =   1270
         Y1              =   1890
         Y2              =   2415
      End
      Begin VB.Line Line10 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   1260
         X2              =   735
         Y1              =   2415
         Y2              =   2415
      End
      Begin VB.Line Line9 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   725
         X2              =   725
         Y1              =   1890
         Y2              =   2415
      End
      Begin VB.Line Line8 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   735
         X2              =   1260
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Line Line7 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   1270
         X2              =   1270
         Y1              =   1155
         Y2              =   1680
      End
      Begin VB.Line Line6 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   1260
         X2              =   735
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line5 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   725
         X2              =   725
         Y1              =   1155
         Y2              =   1680
      End
      Begin VB.Line Line4 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   735
         X2              =   1260
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line Line3 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   1270
         X2              =   1270
         Y1              =   420
         Y2              =   945
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   735
         X2              =   1260
         Y1              =   950
         Y2              =   950
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   725
         X2              =   725
         Y1              =   420
         Y2              =   945
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   750
         Picture         =   "OPCIDEPPT.frx":014A
         Top             =   435
         Width           =   240
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   755
         Picture         =   "OPCIDEPPT.frx":1E44
         Stretch         =   -1  'True
         Top             =   1900
         Width           =   480
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Información Comercial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   330
         Left            =   1785
         TabIndex        =   4
         Top             =   2835
         Width           =   3690
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Instrucciones para Demo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   330
         Left            =   1785
         TabIndex        =   3
         Top             =   525
         Width           =   3795
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ejecutar Demostración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   330
         Left            =   1785
         TabIndex        =   2
         Top             =   1260
         Width           =   3900
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Acerca de FLEXOFT ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   330
         Left            =   1785
         TabIndex        =   1
         Top             =   1995
         Width           =   3900
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   755
         Picture         =   "OPCIDEPPT.frx":214E
         Top             =   2740
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   755
         Picture         =   "OPCIDEPPT.frx":2458
         Stretch         =   -1  'True
         Top             =   1165
         Width           =   480
      End
   End
   Begin VB.Image Image6 
      Height          =   510
      Left            =   6075
      Picture         =   "OPCIDEPPT.frx":2762
      Top             =   3045
      Width           =   2010
   End
End
Attribute VB_Name = "OPCIDEPPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Command9_Click()
   DEMOSPPT.Command1.Visible = True
   Hide
   MASINPPT.Show 1
End Sub

Private Sub Form_Deactivate()
   Label4.ForeColor = &H8080&
   Label1.ForeColor = &H8080&
   Label3.ForeColor = &H8080&
   Label5.ForeColor = &H8080&
   Label1.FontUnderline = False
   Label3.FontUnderline = False
   Label4.FontUnderline = False
   Label5.FontUnderline = False
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Label4.ForeColor = &H8080&
   Label1.ForeColor = &H8080&
   Label3.ForeColor = &H8080&
   Label5.ForeColor = &H8080&
   Label1.FontUnderline = False
   Label3.FontUnderline = False
   Label4.FontUnderline = False
   Label5.FontUnderline = False
   Line1.Visible = False
   Line2.Visible = False
   Line3.Visible = False
   Line4.Visible = False
   Line5.Visible = False
   Line6.Visible = False
   Line7.Visible = False
   Line8.Visible = False
   Line9.Visible = False
   Line10.Visible = False
   Line11.Visible = False
   Line12.Visible = False
   Line13.Visible = False
   Line14.Visible = False
   Line15.Visible = False
   Line16.Visible = False
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Line13.Visible = True
   Line14.Visible = True
   Line15.Visible = True
   Line16.Visible = True
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then Image3.BorderStyle = 1
   Hide
   MENDEPPT.Show
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Line5.Visible = True
   Line6.Visible = True
   Line7.Visible = True
   Line8.Visible = True
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then Image3.BorderStyle = 0
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Line1.Visible = True
   Line2.Visible = True
   Line3.Visible = True
   Line4.Visible = True
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Line17.Visible = True
   Line18.Visible = True
   Line19.Visible = True
   Line20.Visible = True
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then Image4.BorderStyle = 0
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then Image5.BorderStyle = 1
   MsgBox "FLEXOFT Soluciones Informáticas" & vbCrLf & "Bauness 2170 - 5º Piso - Of. 'D'" & vbCrLf & "TeleFax: 00-54-11-4524-1284 / 4524-0796 / 4522-2048" & vbCrLf & "e-mail: info@flexoft.com.ar / web: www.flexoft.com.ar"
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Line9.Visible = True
   Line10.Visible = True
   Line11.Visible = True
   Line12.Visible = True
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then Image5.BorderStyle = 0
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then Image1.BorderStyle = 1
   MASINPPT.Show 1
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then Image1.BorderStyle = 0
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then Label2.BorderStyle = 1
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then Label2.BorderStyle = 0
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   If Button = 1 Then Image2.BorderStyle = 1
'   Hide
'   Call SECUDEMO("INSTRU")
    Call ShellExecute(Me.hwnd, "Open", "Instrucciones demo.pps", "", "", 1)
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then Image2.BorderStyle = 0
End Sub

Private Sub Label1_DblClick()
    MsgBox "FLEXOFT Soluciones Informáticas" & vbCrLf & "Bauness 2170 - 5º Piso - Of. 'D'" & vbCrLf & "TeleFax: 00-54-11-4524-1284 / 4524-0796 / 4522-2048" & vbCrLf & "e-mail: info@flexoft.com.ar / web: www.flexoft.com.ar"
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Label1.ForeColor = &HC0&
   Label1.FontUnderline = True
End Sub

Private Sub Label3_DblClick()
   Hide
   MENDEPPT.Show
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Label3.ForeColor = &HC0&
   Label3.FontUnderline = True
End Sub

Private Sub Label4_DblClick()
    Call ShellExecute(Me.hwnd, "Open", "Instrucciones demo.pps", "", "", 1)
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Label4.ForeColor = &HC0&
   Label4.FontUnderline = True
End Sub

Private Sub Label5_DblClick()
   If Button = 1 Then Image1.BorderStyle = 1
   MASINPPT.Show 1
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Label5.ForeColor = &HC0&
   Label5.FontUnderline = True
End Sub
