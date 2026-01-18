VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form CFGDES07 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Comprobante de Despacho"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6480
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
   ScaleHeight     =   8115
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CMD1 
      Left            =   3000
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Reporte Despachos Diarios en Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Index           =   1
      Left            =   240
      TabIndex        =   32
      Top             =   6120
      Width           =   5160
      Begin VB.CommandButton Command10 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   37
         Top             =   960
         Width           =   3015
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Puerto Impresión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   1
         Left            =   210
         TabIndex        =   35
         Top             =   1200
         Width           =   4800
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   210
            TabIndex        =   36
            Top             =   300
            Width           =   4425
         End
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   210
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "CFGDES07.frx":0000
         Top             =   240
         Width           =   4470
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
         Height          =   300
         Left            =   4740
         TabIndex        =   33
         Top             =   600
         Width           =   175
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2730
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   30
      Top             =   630
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   8130
      Left            =   5565
      ScaleHeight     =   8070
      ScaleWidth      =   945
      TabIndex        =   25
      Top             =   0
      Width           =   1000
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   28
         Top             =   7335
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   12000
         End
      End
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
         Height          =   650
         Left            =   105
         Picture         =   "CFGDES07.frx":0002
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1050
         Width           =   650
      End
      Begin VB.CommandButton command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "CFGDES07.frx":030C
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   315
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "CFGDES07.frx":0456
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   7650
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Etiqueta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   2020
      Width           =   5160
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2415
         TabIndex        =   21
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
         Height          =   300
         Left            =   3780
         TabIndex        =   20
         Top             =   325
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
         Height          =   300
         Left            =   210
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   " "
         Top             =   750
         Width           =   3750
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   740
         Left            =   4095
         Picture         =   "CFGDES07.frx":2378
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   325
         Width           =   855
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Puerto Impresión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   0
         Left            =   210
         TabIndex        =   17
         Top             =   1200
         Width           =   4800
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   210
            TabIndex        =   24
            Top             =   300
            Width           =   4425
         End
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
         Index           =   0
         Left            =   0
         TabIndex        =   22
         Top             =   450
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Depósitos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   240
      TabIndex        =   7
      Top             =   4065
      Width           =   5160
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
         Height          =   300
         Left            =   4200
         TabIndex        =   14
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
         Height          =   300
         Left            =   4725
         TabIndex        =   13
         Top             =   1150
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
         Height          =   300
         Left            =   210
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1550
         Width           =   4695
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
         Height          =   300
         Left            =   4200
         TabIndex        =   10
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
         Height          =   300
         Left            =   4725
         TabIndex        =   9
         Top             =   250
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
         Height          =   300
         Left            =   210
         TabIndex        =   8
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   840
         TabIndex        =   15
         Top             =   1250
         Width           =   3270
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Productos Terminados:"
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
         Left            =   840
         TabIndex        =   11
         Top             =   350
         Width           =   3270
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Comprobante de Despacho"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2060
      Left            =   210
      TabIndex        =   0
      Top             =   -30
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
         TabIndex        =   6
         Top             =   1155
         Width           =   4800
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Suprime Impresión al Generar el Comprobante"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   315
            TabIndex        =   23
            Top             =   215
            Width           =   4320
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   740
         Left            =   4095
         Picture         =   "CFGDES07.frx":2682
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   315
         Width           =   855
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   210
         TabIndex        =   3
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
         Height          =   300
         Left            =   3780
         TabIndex        =   2
         Top             =   325
         Width           =   175
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
         Height          =   300
         Left            =   2415
         TabIndex        =   1
         Text            =   " "
         Top             =   325
         Width           =   1380
      End
      Begin VB.Label Label4 
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
         Top             =   450
         Width           =   2325
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "CFGDES07.frx":298C
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "CFGDES07.frx":2BC0
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "CFGDES07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
   If CLACONTRA% < 1 Then Exit Sub

   If TRIM$(Control("BLOQEXCE", "BL0LISDE")) = "SI" Then
      Call GRACONTROL("BLOQEXCE", "BL0LISDE", "NO")
   Else
      Call GRACONTROL("BLOQEXCE", "BL0LISDE", "SI")
   End If
   Command10.Caption = "Bloqueado: " & TRIM$(Control("BLOQEXCE", "BL0LISDE"))
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
    
    Call GRACONTROL(IDENTER$, "PULISDES", Combo2.TEXT)
    Call GRACONTROL("*", "PULISDES", Combo2.TEXT)
    Call GRACONTROL("", "RULISDES", Text10.TEXT)
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

Private Sub Command9_Click()
   Command9.Enabled = False
   RUTAX$ = TRIM$(SELECCION_Archivo$("*.*", CMD1))
   If RUTAX$ <> "" Then Text10 = RUTAX$
   Command9.Enabled = True
   Screen.MousePointer = 1


End Sub

Private Sub Form_Load()
    '
    Call APLICACIONSKINS(Me, Picture1)
    
    Text6.TEXT = Control$("", "CDESPACH")
    Text7.TEXT = ECOARCH$("INDIFRM", Text6.TEXT)
    '\\\OT-07-0035-RG-31/01/07///
    Text8.TEXT = Control$("", "ETIDESP")
    Text5.TEXT = ECOARCH$("INDIFRM", Text8.TEXT)
    '\\\OT-07-0035-RG-FIN///
    Text2 = Control("", "DEPOTERM")
    Text4 = Control("", "DEPOEXPE")
    '
    If Control$("CDESPACH", "SUPRINT") = "SI" Then
       Check2.Value = 1
       CANCELPRINT% = 1
    End If
    Text10.TEXT = Control("", "RULISDES") 'RUTA LISTA DE DESPACHOS AHP
    '\\\OT-07-0035-RG-31/01/07///
    Combo1.Clear
    Dim PRX As Printer
    For Each PRX In Printers
      Combo1.AddItem PRX.DeviceName
    Next
    Combo1.TEXT = TRIM$(Control$(IDENTER$, "PUERDES"))
      If Combo1.TEXT = "" Then Combo1.TEXT = Control$("*", "PUERDES")
      If Combo1.TEXT = "" Then Combo1.TEXT = Printer.DeviceName
      
    Combo2.Clear
    For Each PRX In Printers
      Combo2.AddItem PRX.DeviceName
    Next
    Combo2.TEXT = TRIM$(Control$(IDENTER$, "PULISDES"))
    Command10.Caption = "Bloqueado: " & TRIM$(Control("BLOQEXCE", "BL0LISDE"))
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
