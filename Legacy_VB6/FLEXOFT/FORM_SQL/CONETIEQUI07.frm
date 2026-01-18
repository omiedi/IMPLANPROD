VERSION 5.00
Begin VB.Form CONETIEQUI07 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Formularios"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6225
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
   ScaleHeight     =   5340
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Formato de N.Serie"
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
      Left            =   210
      TabIndex        =   26
      Top             =   2400
      Width           =   4950
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
         Left            =   360
         MaxLength       =   24
         TabIndex        =   27
         ToolTipText     =   "INDIQUE EL FORMATO X PARA LOS CARACTERES REEMPLAZABLES E INTERCALE EN SU POSICIÓN LOS SEPARADORES CORRESPONDIENTES"
         Top             =   240
         Width           =   4305
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Puerto de Impresion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   210
      TabIndex        =   15
      Top             =   3120
      Width           =   4950
      Begin VB.ComboBox Combo1 
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
         Left            =   975
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   360
         Width           =   3795
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Puerto:"
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
         Left            =   -1440
         TabIndex        =   17
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Impresión Alta / Baja Número de Serie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   240
      TabIndex        =   9
      Top             =   4080
      Width           =   4920
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
         Height          =   650
         Left            =   4060
         Picture         =   "CONETIEQUI07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   645
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
         Height          =   300
         Left            =   240
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   " "
         Top             =   720
         Width           =   3630
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
         Left            =   3720
         TabIndex        =   12
         Top             =   360
         Width           =   180
      End
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
         Height          =   300
         Left            =   2400
         TabIndex        =   11
         Text            =   " "
         Top             =   360
         Width           =   1380
      End
      Begin VB.Label Label1 
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
         TabIndex        =   10
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Impresion de Etiqueta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      Left            =   210
      TabIndex        =   3
      Top             =   120
      Width           =   4950
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Impresión Autómatica Cuando se Genera La O.F."
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
         Left            =   200
         TabIndex        =   25
         Top             =   1920
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
         Height          =   650
         Left            =   4095
         Picture         =   "CONETIEQUI07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1200
         Width           =   645
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
         Height          =   285
         Left            =   210
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1560
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
         TabIndex        =   19
         Top             =   1200
         Width           =   180
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
         Height          =   300
         Left            =   2415
         TabIndex        =   18
         Text            =   " "
         Top             =   1200
         Width           =   1380
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   7
         Text            =   " "
         Top             =   345
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
         Height          =   300
         Left            =   3780
         TabIndex        =   6
         Top             =   345
         Width           =   180
      End
      Begin VB.TextBox Text9 
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
         Left            =   210
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   705
         Width           =   3750
      End
      Begin VB.CommandButton Command8 
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
         Height          =   650
         Left            =   4095
         Picture         =   "CONETIEQUI07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   345
         Width           =   645
      End
      Begin VB.Label Label2 
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
         TabIndex        =   22
         Top             =   1260
         Width           =   2325
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
         TabIndex        =   8
         Top             =   405
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   7140
      Left            =   5355
      ScaleHeight     =   7080
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   0
      Width           =   1200
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   100
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   23
         Top             =   3960
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   120
            TabIndex        =   24
            Top             =   120
            Width           =   12000
         End
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
         Height          =   500
         Left            =   105
         Picture         =   "CONETIEQUI07.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   315
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   105
         Picture         =   "CONETIEQUI07.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Guardar"
         Top             =   840
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "CONETIEQUI07.frx":0D72
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4200
         Width           =   1515
      End
   End
End
Attribute VB_Name = "CONETIEQUI07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub
'\\\OT-05-0723-WAR-22/11/05///

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "STIKAPRO", Text10.TEXT)
    Call GRACONTROL("", "STIKAPR2", Text1.TEXT)
    '
    Call GRACONTROL(IDENTER$, "PORTSTIK", Combo1.TEXT)
    Call GRACONTROL("*", "PORTSTIK", Combo1.TEXT)
    
    Call GRACONTROL("", "ALBANSE", Text4.TEXT)
    '
    IMPRAUTO$ = "NO"
    If Check4.Value = 1 Then IMPRAUTO$ = "SI"
    Call GRACONTROL(IDENTER$, "IMPAUTOM", IMPRAUTO$)
    '
    Call GRACONTROL("", "FORMATNS", Text3.TEXT)
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text4.TEXT = VDEVU$
        Text5.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'

Private Sub Command4_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
    End If
End Sub

Private Sub Command5_Click()
     Call CONECRUN("FLFORMS/" + TRIM$(Text4.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command6_Click()
   Call CONECRUN("FLFORMS/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command8_Click()
     Call CONECRUN("FLFORMS/" + TRIM$(Text10.TEXT), "Formularios de Impresion")
End Sub
Private Sub Command9_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text10.TEXT = VDEVU$
    End If
End Sub

Private Sub Form_Load()
    '
    '
    Text10.TEXT = CONTROL$("", "STIKAPRO")
    Text9.TEXT = ECOARCH$("INDIFRM", Text10.TEXT)
    Text1.TEXT = CONTROL$("", "STIKAPR2")
    Text3.TEXT = CONTROL("", "FORMATNS")
    '
    '
    PFAC$ = CONTROL$(IDENTER$, "PORTSTIK")
    If PFAC$ = "" Then PFAC$ = CONTROL$("*", "PORTSTIK")
    Combo1.Clear
    Dim PRX As Printer
    On Error Resume Next
    Combo1.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo1.AddItem PRX.DeviceName
      If PRX.DeviceName = PFAC$ Then
        Combo1.TEXT = PFAC$
      End If
    Next
    PFAC$ = TRIM$(CONTROL$(IDENTER$, "PORTSTIK"))
    If PFAC$ <> "" Then Combo1.TEXT = PFAC$
    '
    Text4.TEXT = CONTROL$("", "ALBANSE")
    Text5.TEXT = ECOARCH$("INDIFRM", Text4.TEXT)
    
    Check4.Value = 0
    IMPR$ = TRIM$(CONTROL$(IDENTER$, "IMPAUTOM"))
    If IMPR$ = "SI" Then Check4.Value = 1
    '
End Sub

Private Sub Text1_Change()
   Text2.TEXT = ECOARCH$("INDIFRM", UCase$(Text1.TEXT))
End Sub

Private Sub Text1_DblClick()
  Call Command4_Click
End Sub

Private Sub Text10_Change()
    Text9.TEXT = ECOARCH$("INDIFRM", UCase$(Text10.TEXT))
End Sub

Private Sub Text10_DblClick()
  Call Command9_Click
End Sub
