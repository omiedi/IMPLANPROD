VERSION 5.00
Begin VB.Form STPPMOD04 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes de Producción - Configuración"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5490
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
   ScaleHeight     =   5070
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Registro de Autocontrol"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   12
      Top             =   3675
      Width           =   4215
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
         Height          =   330
         Left            =   1680
         TabIndex        =   15
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.CommandButton Command3 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3045
         TabIndex        =   14
         Top             =   315
         Width           =   175
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
         Height          =   750
         Left            =   3360
         Picture         =   "STPPMOD04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   315
         Width           =   750
      End
      Begin VB.Label Label2 
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
         Height          =   330
         Left            =   105
         TabIndex        =   17
         Top             =   735
         Width           =   3135
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   525
         TabIndex        =   16
         Top             =   420
         Width           =   1065
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Opciones Operativas"
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
      Left            =   210
      TabIndex        =   9
      Top             =   2415
      Width           =   4215
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Discriminación Piezas No Conformes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   525
         TabIndex        =   11
         Top             =   735
         Width           =   3585
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Informe de Horas de Descanso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   525
         TabIndex        =   10
         Top             =   420
         Width           =   3060
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Límite Rendimiento"
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
      Left            =   210
      TabIndex        =   6
      Top             =   1575
      Width           =   4215
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Left            =   2730
         TabIndex        =   7
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Máximo Calculado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -630
         TabIndex        =   8
         Top             =   360
         Width           =   3270
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000080&
      Height          =   5055
      Left            =   4620
      ScaleHeight     =   4995
      ScaleWidth      =   1005
      TabIndex        =   3
      Top             =   -105
      Width           =   1065
      Begin VB.CommandButton Command14 
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
         Height          =   600
         Left            =   105
         Picture         =   "STPPMOD04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Aprueba, Graba e Imprime Presupuesto"
         Top             =   1050
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
         Left            =   105
         Picture         =   "STPPMOD04.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Exit"
         Top             =   420
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "STPPMOD04.frx":075E
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4515
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Configuración Horaria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4215
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Minutos Centesimales"
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
         Left            =   525
         TabIndex        =   2
         Top             =   735
         Width           =   3165
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Normal (Horas, Minutos y Segundos)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   525
         TabIndex        =   1
         Top             =   315
         Value           =   -1  'True
         Width           =   3480
      End
   End
End
Attribute VB_Name = "STPPMOD04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub
'

Private Sub Command14_Click()
   '\\\OT-06-0309-MH-28/08/06///
    OPT$ = "HORANORM"
    If Option3.Value = True Then
       OPT$ = "MINUCENT"
    End If
    RENMAX$ = TRIM$(FORMATNUM$(Text1.TEXT, "F7.1"))
    Call GRACONTROL("PADIAMOD", "CONFHORA", OPT$)
    Call GRACONTROL("PADIAMOD", "RENDIMAX", RENMAX$)
    '\\\OT-06-0309-MH-FIN///
    INFDESCA$ = "NO"
    If Check1.Value = 1 Then INFDESCA$ = "SI"
    Call GRACONTROL("PADIAMOD", "INFDESCA", INFDESCA$)
    '
    PINOCONF$ = "NO"
    If Check2.Value = 1 Then PINOCONF$ = "SI"
    Call GRACONTROL("PADIAMOD", "PINOCONF", PINOCONF$)
    '
    '\\\OT-07-0087-RG-06/03/07///
    Call GRACONTROL("IFABRICA", "FAUTCONT", Text4.TEXT)
    '\\\OT-07-0087-RG-FIN///
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub



Private Sub Command3_Click()
    '
    '\\\OT-07-0087-RG-06/03/07///
    Call CommSel("INDIFRM", Command3, Text4)
    '\\\OT-07-0087-RG-FIN///
    '
End Sub

Private Sub Command5_Click()
    '
    '\\\OT-07-0087-RG-06/03/07///
    Command5.Enabled = False
    Call CONECRUN("FLFORMS/" + TRIM$(Text4.TEXT), "Formularios de Impresion")
    Command5.Enabled = True
    '\\\OT-07-0087-RG-FIN///
    '
End Sub

Private Sub Form_Load()
    '\\\OT-06-0309-MH-28/08/06///
    VALORCONTROL$ = UCase$(TRIM$(CONTROL$("PADIAMOD", "CONFHORA")))
    If VALORCONTROL$ = "MINUCENT" Then
       Option3.Value = True
    End If
    RENMAX$ = CONTROL("PADIAMOD", "RENDIMAX")
    Text1.TEXT = RENMAX$
    '
    If CONTROL$("PADIAMOD", "INFDESCA") = "SI" Then Check1.Value = 1
    If CONTROL$("PADIAMOD", "PINOCONF") = "SI" Then Check2.Value = 1
    '\\\OT-06-0309-MH-FIN///
    '
    '\\\OT-07-0087-RG-06/03/07///
    Text4.TEXT = CONTROL$("IFABRICA", "FAUTCONT")
    '\\\OT-07-0087-RG-FIN///
    '
End Sub

Private Sub Text1_LostFocus()
    RENMAX$ = TRIM$(FORMATNUM$(Text1.TEXT, "F7.1"))
    Text1 = RENMAX$
End Sub

Private Sub Text4_Change()
    '
    '\\\OT-07-0087-RG-06/03/07///
    Label2.Caption = ECOARCH$("INDIFRM", Text4)
    '\\\OT-07-0087-RG-FIN///
    '
End Sub

Private Sub Text4_DblClick()
    '
    '\\\OT-07-0087-RG-06/03/07///
    Call Command3_Click
    '\\\OT-07-0087-RG-FIN///
    '
End Sub
