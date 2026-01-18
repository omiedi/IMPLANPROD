VERSION 5.00
Begin VB.Form OPCHEQ00 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Gestión de Valores"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      Caption         =   "Cuentas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   4830
      Picture         =   "OPCHEQ00.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Archivo Maestro de Cuentas Contables"
      Top             =   1965
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cheques Rechazados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2640
      Left            =   210
      TabIndex        =   4
      Top             =   1575
      Width           =   4425
      Begin VB.CommandButton Command14 
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
         Height          =   285
         Left            =   3990
         TabIndex        =   8
         Top             =   525
         Width           =   175
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Procesamiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   210
         TabIndex        =   5
         Top             =   1365
         Width           =   4005
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Por Ajuste de Saldo Deudor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   210
            TabIndex        =   7
            Top             =   420
            Value           =   -1  'True
            Width           =   1800
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Por Nota de Debito al Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1995
            TabIndex        =   6
            Top             =   420
            Width           =   1905
         End
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Cheques Rechazados:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   105
         TabIndex        =   11
         Top             =   420
         Width           =   2325
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   10
         Top             =   525
         Width           =   1485
      End
      Begin VB.Label Label12 
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
         Left            =   375
         TabIndex        =   9
         Top             =   900
         Width           =   3795
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Bancos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   4830
      Picture         =   "OPCHEQ00.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Acceso a Tabla de Cuentas Bancarias"
      Top             =   2790
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gestion de Valores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   4425
      Begin VB.CommandButton Command13 
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
         Height          =   285
         Left            =   3990
         TabIndex        =   12
         Top             =   420
         Width           =   175
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cartera de Valores:"
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
         Left            =   105
         TabIndex        =   15
         Top             =   480
         Width           =   2325
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   14
         Top             =   420
         Width           =   1485
      End
      Begin VB.Label Label9 
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
         Left            =   375
         TabIndex        =   13
         Top             =   795
         Width           =   3795
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   4830
      Picture         =   "OPCHEQ00.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1155
      Width           =   855
   End
   Begin VB.CommandButton command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   4830
      Picture         =   "OPCHEQ00.frx":075E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   345
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   4620
      X2              =   5880
      Y1              =   4190
      Y2              =   4190
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4305
      Picture         =   "OPCHEQ00.frx":08A8
      Top             =   3680
      Width           =   2010
   End
End
Attribute VB_Name = "OPCHEQ00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CARCUEINGAS
   Command11.Enabled = True
End Sub

Private Sub Command14_Click()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label11 = VDEVU$
        Label12 = VALACT2$("ECUECON")
    End If
End Sub
'
Private Sub Command13_Click()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label8 = VDEVU$
        Label9 = VALACT2$("ECUECON")
    End If
End Sub
'
Private Sub Command2_Click()
    '
    Call GRACONTROL("", "CUECHECA", Label8)
    Call GRACONTROL("", "CUECHERE", Label11)
    '
    MONUME$ = "AJUSALDO"
    If Option4.Value = True Then MONUME$ = "NOTADEBI"
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'

Private Sub Command4_Click()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call CONECRUN("PRIFORMS/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call CARCUEBANC
   Command6.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Label8 = CONTROL$("", "CUECHECA")
    Label9 = ECOARCH$("ECUECON", Label8)
    Label11 = CONTROL$("", "CUECHERE")
    Label12 = ECOARCH$("ECUECON", Label11)
    '
    If CONTROL$("", "CHERECHA") = "NOTADEBI" Then Option4.Value = True
    '
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Label11_DblClick()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label11 = VDEVU$
        Label12 = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Label8_DblClick()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label8 = VDEVU$
        Label9 = VALACT2$("ECUECON")
    End If
End Sub
