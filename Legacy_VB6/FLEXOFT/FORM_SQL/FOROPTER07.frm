VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FOROPTER 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Documentos de Impresion"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6585
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
   ScaleHeight     =   6000
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Remito de Traslado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   240
      TabIndex        =   22
      Top             =   4080
      Width           =   5160
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   840
         MaxLength       =   4
         TabIndex        =   29
         Top             =   360
         Width           =   470
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2415
         TabIndex        =   27
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.CommandButton Command9 
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
         TabIndex        =   26
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   3750
      End
      Begin VB.CommandButton Command7 
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
         Picture         =   "FOROPTER07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   315
         Width           =   855
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0FFFF&
         Caption         =   " Suprime  la  Impresión  al  Generar  el  Documento"
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
         Left            =   210
         TabIndex        =   23
         Top             =   1365
         Width           =   4740
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "P.Vta.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   30
         Top             =   420
         Width           =   525
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1320
         TabIndex        =   28
         Top             =   420
         Width           =   1000
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Orden de Servicios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   210
      TabIndex        =   13
      Top             =   240
      Width           =   5160
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
         Left            =   4080
         Picture         =   "FOROPTER07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   320
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   3750
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
         TabIndex        =   16
         Top             =   315
         Width           =   175
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
         Height          =   360
         Left            =   2415
         TabIndex        =   15
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   " Suprime  la  Impresión  al  Generar  el  Documento"
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
         Left            =   210
         TabIndex        =   14
         Top             =   1365
         Width           =   4740
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   18
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00004040&
      Height          =   6840
      Left            =   5520
      ScaleHeight     =   6780
      ScaleWidth      =   2160
      TabIndex        =   7
      Top             =   0
      Width           =   2220
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   200
         Picture         =   "FOROPTER07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1800
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   200
         Picture         =   "FOROPTER07.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1080
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   200
         Picture         =   "FOROPTER07.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   210
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   8
         Top             =   5280
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -225
         Picture         =   "FOROPTER07.frx":0D72
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5595
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Orden de Compra / Pedido de Suministro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   210
      TabIndex        =   0
      Top             =   2150
      Width           =   5160
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFFF&
         Caption         =   " Suprime  la  Impresión  al  Generar  el  Documento"
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
         Left            =   210
         TabIndex        =   6
         Top             =   1365
         Width           =   4740
      End
      Begin VB.CommandButton Command6 
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
         Picture         =   "FOROPTER07.frx":2C94
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   315
         Width           =   855
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   3750
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
         Left            =   3780
         TabIndex        =   2
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2415
         TabIndex        =   1
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   5
         Top             =   420
         Width           =   2325
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "FOROPTER07.frx":2F9E
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "FOROPTER07.frx":31D2
      TabIndex        =   21
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
End
Attribute VB_Name = "FOROPTER"
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
    Call GRACONTROL("", "FORSERVI", Text1.TEXT)
    Call GRACONTROL("", "FORMOCOM", Text3.TEXT)
    Call GRACONTROL("", "REMISERV", Text6.TEXT)
    Call GRACONTROL("", "PVTARTSE", Text7.TEXT)
    '
    SUPRINT$ = "NO": CANCELPRINT% = 0
    If Check1.Value = 1 Then SUPRINT$ = "SI": CANCELPRINT% = 1
    Call GRACONTROL("ORSERVI", "SUPRINT", SUPRINT$)
    '
    SUPRINT$ = "NO": CANCELPRINT% = 0
    If Check2.Value = 1 Then SUPRINT$ = "SI": CANCELPRINT% = 1
    Call GRACONTROL("OCOMPRA", "SUPRINT", SUPRINT$)
    '
    SUPRINT$ = "NO": CANCELPRINT% = 0
    If Check3.Value = 1 Then SUPRINT$ = "SI": CANCELPRINT% = 1
    Call GRACONTROL("REMITOSE", "SUPRINT", SUPRINT$)

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
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command5_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command6_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text3.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    Call CONECRUN("FLFORMS/" + TRIM$(Text6.TEXT), "Formularios de Impresion")
    Command7.Enabled = True
End Sub

Private Sub Command9_Click()

    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text5.TEXT = VALACT2$("INDIFRM")
    End If

End Sub

Private Sub Form_Load()
    '
    Text1.TEXT = Control$("", "FORSERVI")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text3.TEXT = Control$("", "FORMOCOM")
    Text4.TEXT = ECOARCH$("INDIFRM", Text3.TEXT)
    Text6.TEXT = Control$("", "REMISERV")
    Text5.TEXT = ECOARCH$("INDIFRM", Text6.TEXT)
    Text7.TEXT = Control$("", "PVTARTSE") ' PUNTO DE VENTA DE REMITO DE SERVICIO
    
    '
    If Control$("ORSERVI", "SUPRINT") = "SI" Then Check1.Value = 1
    If Control$("OCOMPRA", "SUPRINT") = "SI" Then Check2.Value = 1
    If Control$("REMITOSE", "SUPRINT") = "SI" Then Check3.Value = 1
    '
    Call APLICACIONSKINS(Me, Picture14)
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
'
Private Sub Text3_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text4_CLICK()
    Text3.SetFocus
End Sub

Private Sub Text5_Change()
   On Error Resume Next
   Text6.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text6_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text5.TEXT = VALACT2$("INDIFRM")
    End If

End Sub
