VERSION 5.00
Begin VB.Form OPASICONT 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comprobantes de Asiento Contable - Caja y Bancos"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6660
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
   ScaleHeight     =   6120
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command8 
      Caption         =   "Config."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   5600
      Picture         =   "OPASICONT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Configuracion Avanzada "
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   5600
      Picture         =   "OPASICONT.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Verificar y Corregir Fechas Apertura y Cierre"
      Top             =   3600
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
      Height          =   1065
      Left            =   210
      TabIndex        =   17
      Top             =   4935
      Width           =   5160
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
         Left            =   2835
         TabIndex        =   19
         Top             =   420
         Width           =   1905
      End
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
         Left            =   525
         TabIndex        =   18
         Top             =   420
         Value           =   -1  'True
         Width           =   1800
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
      Height          =   800
      Left            =   5600
      Picture         =   "OPASICONT.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Acceso a Tabla de Cuentas Bancarias"
      Top             =   4440
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
      Height          =   1695
      Left            =   210
      TabIndex        =   11
      Top             =   3150
      Width           =   5160
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   4750
      End
      Begin VB.CommandButton Command4 
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
         Height          =   360
         Left            =   4780
         TabIndex        =   13
         Top             =   630
         Width           =   175
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3255
         TabIndex        =   12
         Text            =   " "
         Top             =   630
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Contable Asociada a la Cartera de Valores:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   210
         TabIndex        =   15
         Top             =   525
         Width           =   2955
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Asientos Contables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   210
      TabIndex        =   2
      Top             =   315
      Width           =   5160
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2415
         TabIndex        =   10
         Text            =   " "
         Top             =   420
         Width           =   1380
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
         Height          =   360
         Left            =   3780
         TabIndex        =   9
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   945
         Width           =   3750
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   880
         Left            =   4095
         Picture         =   "OPASICONT.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   420
         Width           =   855
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Comprobante Impreso"
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
         TabIndex        =   3
         Top             =   1470
         Width           =   4750
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Imprime Comprobante"
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
            Left            =   735
            TabIndex        =   5
            Top             =   420
            Value           =   -1  'True
            Width           =   1800
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Suprime Impresión"
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
            Left            =   2730
            TabIndex        =   4
            Top             =   420
            Width           =   1695
         End
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   6
         Top             =   525
         Width           =   2325
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
      Height          =   800
      Left            =   5600
      Picture         =   "OPASICONT.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
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
      Height          =   800
      Left            =   5600
      Picture         =   "OPASICONT.frx":106A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   5040
      Picture         =   "OPASICONT.frx":11B4
      Top             =   5480
      Width           =   2010
   End
End
Attribute VB_Name = "OPASICONT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "ASICONTA", Text1.TEXT)
    Call GRACONTROL("", "CUECHECA", Text3.TEXT)
    '
    MONUME$ = "SI": If Option2.Value = True Then MONUME$ = "NO"
    Call GRACONTROL("ASICONTA", "IMPRECOM", MONUME$)
    '
    MONUME$ = "AJUSALDO": If Option4.Value = True Then MONUME$ = "NOTADEBI"
    Call GRACONTROL("", "CHERECHA", MONUME$)
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

Private Sub Command7_Click()
   '
   Command7.Enabled = False
   OPX% = COMALTER%("Verificacion Fechas de Apertura\y Cierre de Ejercicio Contable:\\Ejercicio Actual\Anterior\Cancelar")
   If OPX% = 1 Then
        EXTU$ = "DAT"
      ElseIf OPX% = 2 Then
        EXTU$ = "OLD"
      Else
        GoTo 99
   End If
   '
   LOFI$ = LOADFILE$(LUDAT$ + "VARIAB." + EXTU$)
   If Len(LOFI$) >= 64 Then
     VDEF$ = Mid$(LOFI$, 33, 32)
     VNUE$ = OBJPLA$("DATEJCONTA", VDEF$)
     If VNUE$ <> "" Then
       Call REPLA(LOFI$, VNUE$, 33, 32)
       Call SAVEFILE(LUDAT$ + "VARIAB." + EXTU$, LOFI$)
       Call CIERRARCH("*.*")
       Call COMUNI("Actualización Completa !!!\Cierre Todo y Re-Inicie FLEXOFT")
     End If
   End If
99 Command7.Enabled = True
   '
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    TTXX$ = "Esta Opcion Solo Debe Ser Utilizada\por Implementadores FLEXOFT Debidamente Calificados.\"
    TTXX$ = TTXX$ + "  \Su Utilizacion Indebida Puede Producir Daños al Sistema\NO CUBIERTOS POR GARANTIA."
    If COMALTER%(TTXX$ + "\\Cancelar\Continuar") = 2 Then
      If CLACONTRA% = 1 Then
        CFGSCGPLUS.Show 1
      End If
    End If
    Command8.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Text1.TEXT = CONTROL$("", "ASICONTA")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    '
    Text3.TEXT = CONTROL$("", "CUECHECA")
    Text4.TEXT = ECOARCH$("ECUECON", Text3.TEXT)
    '
    If CONTROL$("ASICONTA", "IMPRECOM") = "NO" Then Option2.Value = True
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

Private Sub Text3_DblClick()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("ECUECON")
    End If
End Sub

