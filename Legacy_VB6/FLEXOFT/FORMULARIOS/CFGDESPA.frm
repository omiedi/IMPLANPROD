VERSION 5.00
Begin VB.Form CFGDESPA 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Comprobante de Despacho"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6525
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
   ScaleHeight     =   7035
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Etiqueta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2200
      Left            =   210
      TabIndex        =   18
      Top             =   2500
      Width           =   5160
      Begin VB.TextBox Text8 
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
         TabIndex        =   23
         Text            =   " "
         Top             =   325
         Width           =   1380
      End
      Begin VB.CommandButton Command8 
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
         TabIndex        =   22
         Top             =   325
         Width           =   175
      End
      Begin VB.TextBox Text5 
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
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   " "
         Top             =   750
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
         Height          =   780
         Left            =   4095
         Picture         =   "CFGDESPA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   325
         Width           =   855
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Puerto Impresión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   210
         TabIndex        =   19
         Top             =   1200
         Width           =   4800
         Begin VB.ComboBox Combo1 
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
            Left            =   315
            TabIndex        =   26
            Top             =   300
            Width           =   4320
         End
      End
      Begin VB.Label Label3 
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
         Left            =   0
         TabIndex        =   24
         Top             =   450
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Depósitos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   210
      TabIndex        =   9
      Top             =   4800
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
         Left            =   4200
         TabIndex        =   16
         Text            =   " "
         Top             =   1150
         Width           =   540
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
         Left            =   4725
         TabIndex        =   15
         Top             =   1150
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
         Left            =   210
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1550
         Width           =   4695
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
         Left            =   4200
         TabIndex        =   12
         Text            =   " "
         Top             =   250
         Width           =   540
      End
      Begin VB.CommandButton Command5 
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
         Left            =   4725
         TabIndex        =   11
         Top             =   250
         Width           =   175
      End
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
         Left            =   210
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   650
         Width           =   4695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Expedición:"
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
         Left            =   840
         TabIndex        =   17
         Top             =   1250
         Width           =   3270
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Productos Terminados:"
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
         Left            =   840
         TabIndex        =   13
         Top             =   350
         Width           =   3270
      End
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
      Height          =   855
      Left            =   5520
      Picture         =   "CFGDESPA.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   315
      Width           =   855
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
      Height          =   855
      Left            =   5520
      Picture         =   "CFGDESPA.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1260
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Comprobante de Despacho"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2200
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   5160
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Impresión"
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
         TabIndex        =   8
         Top             =   1200
         Width           =   4800
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Suprime Impresión al Generar el Comprobante"
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
            Left            =   315
            TabIndex        =   25
            Top             =   215
            Width           =   4320
         End
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
         Height          =   780
         Left            =   4095
         Picture         =   "CFGDESPA.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   315
         Width           =   855
      End
      Begin VB.TextBox Text7 
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
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   750
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
         TabIndex        =   4
         Top             =   325
         Width           =   175
      End
      Begin VB.TextBox Text6 
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
         TabIndex        =   3
         Text            =   " "
         Top             =   325
         Width           =   1380
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
         Left            =   0
         TabIndex        =   7
         Top             =   450
         Width           =   2325
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4960
      Picture         =   "CFGDESPA.frx":0A68
      Top             =   6370
      Width           =   2010
   End
End
Attribute VB_Name = "CFGDESPA"
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
    Call GRACONTROL("", "CDESPACH", Text6.TEXT)
    Call GRACONTROL("", "DEPOTERM", Text2.TEXT)
    Call GRACONTROL("", "DEPOEXPE", Text4.TEXT)
    '\\\OT-07-0035-RG-31/01/07///
    Call GRACONTROL("", "ETIDESP", Text8.TEXT)
    Call GRACONTROL(IDENTER$, "PUERDES", Combo1.TEXT)
    Call GRACONTROL("*", "PUERDES", Combo1.TEXT)
    '\\\OT-07-0035-RG-FIN///
    SUPRINT$ = "NO": CANCELPRINT% = 0
    If Check2.Value = 1 Then SUPRINT$ = "SI": CANCELPRINT% = 1
    Call GRACONTROL("CDESPACH", "SUPRINT", SUPRINT$)
    '
    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completa.\Re-Inicie Ahora la Aplicación.")
    Call BAJALDISCO
    Call FINAL("")
    Hide
    '
End Sub


Private Sub Command3_Click()
   Call SELECHO("TADEPOSI")
   VDEVU$ = TRIM$(VALACT1$("TADEPOSI"))
   If VDEVU$ <> "" Then
     Text4 = VDEVU$
   End If
End Sub

Private Sub Command4_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command5_Click()
   Call SELECHO("TADEPOSI")
   VDEVU$ = TRIM$(VALACT1$("TADEPOSI"))
   If VDEVU$ <> "" Then
     Text2 = VDEVU$
   End If
End Sub

Private Sub Command6_Click()
   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text6.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command7_Click()
    '\\\OT-07-0035-RG-31/01/07///
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text8.TEXT), "Formularios de Impresion")
    '\\\OT-07-0035-RG-FIN///
End Sub

Private Sub Command8_Click()
    '\\\OT-07-0035-RG-31/01/07///
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text8.TEXT = VDEVU$
        Text5.TEXT = VALACT2$("INDIFRM")
    End If
    '\\\OT-07-0035-RG-FIN///
End Sub

Private Sub Form_Load()
    '
    Text6.TEXT = CONTROL$("", "CDESPACH")
    Text7.TEXT = ECOARCH$("INDIFRM", Text6.TEXT)
    '\\\OT-07-0035-RG-31/01/07///
    Text8.TEXT = CONTROL$("", "ETIDESP")
    Text5.TEXT = ECOARCH$("INDIFRM", Text8.TEXT)
    '\\\OT-07-0035-RG-FIN///
    Text2 = CONTROL("", "DEPOTERM")
    Text4 = CONTROL("", "DEPOEXPE")
    '
    If CONTROL$("CDESPACH", "SUPRINT") = "SI" Then
       Check2.Value = 1
       CANCELPRINT% = 1
    End If
    '\\\OT-07-0035-RG-31/01/07///
    Combo1.Clear
    Dim PRX As Printer
    For Each PRX In Printers
      Combo1.AddItem PRX.DeviceName
    Next
    Combo1.TEXT = TRIM$(CONTROL$(IDENTER$, "PUERDES"))
      If Combo1.TEXT = "" Then Combo1.TEXT = CONTROL$("*", "PUERDES")
      If Combo1.TEXT = "" Then Combo1.TEXT = Printer.DeviceName
    '\\\OT-07-0035-RG-FIN///
End Sub

Private Sub Text1_GotFocus()
   Text2.SetFocus
End Sub

Private Sub Text2_Change()
   Text1 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text2)))
End Sub

Private Sub Text2_DblClick()
   Call Command5_Click
End Sub

Private Sub Text3_GotFocus()
   Text4.SetFocus
End Sub

Private Sub Text4_Change()
   Text3 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text4)))
End Sub

Private Sub Text4_DblClick()
   Call Command3_Click
End Sub

Private Sub Text5_Click()
    '\\\OT-07-0035-RG-31/01/07///
    Text8.SetFocus
    '\\\OT-07-0035-RG-FIN///
End Sub

Private Sub Text7_CLICK()
    Text6.SetFocus
End Sub


Private Sub Text6_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text6.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text8_DblClick()
    '\\\OT-07-0035-RG-31/01/07///
    Call Command8_Click
    '\\\OT-07-0035-RG-FIN///
End Sub
