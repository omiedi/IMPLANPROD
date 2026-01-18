VERSION 5.00
Begin VB.Form CFGETIDE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Etiqueta de Despacho"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6165
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5040
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
         Height          =   720
         Left            =   4080
         Picture         =   "CFGETIDE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   320
         Width           =   750
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
         Height          =   735
         Left            =   180
         TabIndex        =   9
         Top             =   1250
         Width           =   4680
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
            Left            =   160
            TabIndex        =   10
            Top             =   300
            Width           =   4320
         End
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   750
         Width           =   3750
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
         Height          =   315
         Left            =   3780
         TabIndex        =   7
         Top             =   325
         Width           =   175
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2415
         TabIndex        =   6
         Text            =   " "
         Top             =   325
         Width           =   1380
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
         TabIndex        =   11
         Top             =   410
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   5775
      Left            =   5280
      ScaleHeight     =   5715
      ScaleWidth      =   1455
      TabIndex        =   1
      Top             =   0
      Width           =   1515
      Begin VB.CommandButton Command3 
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
         Height          =   650
         Left            =   120
         Picture         =   "CFGETIDE.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Cerrar y Salir"
         Top             =   120
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
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
         Left            =   120
         Picture         =   "CFGETIDE.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Guarda la Configuracion"
         Top             =   840
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   2
         Top             =   1560
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "CFGETIDE.frx":075E
         Top             =   1800
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   10200
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   5400
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   7440
      TabIndex        =   4
      Top             =   1140
      Width           =   1695
   End
End
Attribute VB_Name = "CFGETIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
    Call GRACONTROL("", "ETIDESP", Text8.TEXT)
    Call GRACONTROL(IDENTER$, "PUERDES", Combo1.TEXT)
    Call GRACONTROL("*", "PUERDES", Combo1.TEXT)
    '
    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completa.\Re-Inicie Ahora la Aplicación.")
    Call BAJALDISCO
    Call FINAL("")
    Hide
End Sub

Private Sub Command3_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command7_Click()
    ACONEC$ = "FLFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text8.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command8_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text8.TEXT = VDEVU$
        Text5.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Form_Load()
    Text8.TEXT = CONTROL$("", "ETIDESP")
    Text5.TEXT = ECOARCH$("INDIFRM", Text8.TEXT)
    Dim PRX As Printer
    For Each PRX In Printers
      Combo1.AddItem PRX.DeviceName
    Next
    Combo1.TEXT = TRIM$(CONTROL$(IDENTER$, "PUERDES"))
    If Combo1.TEXT = "" Then Combo1.TEXT = CONTROL$("*", "PUERDES")
    If Combo1.TEXT = "" Then Combo1.TEXT = Printer.DeviceName
End Sub

Private Sub Text5_GotFocus()
    Text8.SetFocus
End Sub

Private Sub Text8_DblClick()
    Call Command8_Click
End Sub
