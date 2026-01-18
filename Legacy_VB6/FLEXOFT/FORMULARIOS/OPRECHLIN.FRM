VERSION 5.00
Begin VB.Form OPRECHLIN 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rechazos en Línea - Configuración"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ruta de Impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   2880
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   140
         TabIndex        =   14
         Text            =   " "
         Top             =   360
         Width           =   2595
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Tablas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   3120
      TabIndex        =   11
      Top             =   120
      Width           =   2040
      Begin VB.CommandButton Command3 
         Caption         =   "Destino"
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
         Left            =   1080
         Picture         =   "OPRECHLIN.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Actualizar Tabla de Destinos de Rechazo."
         Top             =   230
         Width           =   780
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Motivo"
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
         Left            =   150
         Picture         =   "OPRECHLIN.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Actualizar Tabla de Motivos de Rechazo."
         Top             =   230
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Impresión de Rechazo"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   5040
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
         Height          =   780
         Left            =   4095
         Picture         =   "OPRECHLIN.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   315
         Width           =   750
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
         Left            =   210
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   3750
      End
      Begin VB.CommandButton Command7 
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
         TabIndex        =   7
         Top             =   315
         Width           =   175
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
         Left            =   2415
         TabIndex        =   6
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label5 
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008080&
      Height          =   4635
      Left            =   5310
      ScaleHeight     =   4575
      ScaleWidth      =   1320
      TabIndex        =   0
      Top             =   -150
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   3
         Top             =   1800
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   4
            Top             =   105
            Width           =   12000
         End
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
         Height          =   690
         Left            =   105
         Picture         =   "OPRECHLIN.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   1015
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "OPRECHLIN.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Guarda - Cierra - Abandona la Aplicación"
         Top             =   330
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "OPRECHLIN.frx":09F2
         Top             =   2040
         Width           =   1515
      End
   End
End
Attribute VB_Name = "OPRECHLIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  '
  Call CIERRARCH("*.*")
  Unload Me
  '
End Sub

Private Sub Command2_Click()
  '
  Call GRACONTROL("", "RECHALIN", Text7.TEXT)
  Call GRACONTROL("", "RUIMPREC", Combo1.TEXT)  'RUTA DE IMPRESION DE RECHAZOS
  '
End Sub

Private Sub Command3_Click()
  '
  If DERACCE%("EDIMOTIV", "Edición de Tabla de Motivos de Rechazo") >= 2 Then
    Call TRALOCAL("TADESTI", "TADESTI")
    '
    ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 1400))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
    VTB% = VENTABU%("TADESTI/" + ATRI$)
    If VTB% >= 0 Then
       JJ& = 0
      For U& = 1 To ULTREG&("!TADESTI")
         X$ = REGLEIDO$("!TADESTI", U&)
         JJ& = JJ& + 1
         Call GRAREG("TADESTI", X$, JJ&)
      Next U&
      Call SETULTREG("TADESTI", JJ&)
      Call CIERRARCH("TADESTI")
    End If
  End If
  '
End Sub

Private Sub Command6_Click()
  '
  If DERACCE%("EDIMOTIV", "Edición de Tabla de Motivos de Rechazo") >= 2 Then
    Call TRALOCAL("TAMOTIV", "TAMOTIV")
    '
    ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 1400))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
    VTB% = VENTABU%("TAMOTIV/" + ATRI$)
    If VTB% >= 0 Then
       JJ& = 0
      For U& = 1 To ULTREG&("!TAMOTIV")
         X$ = REGLEIDO$("!TAMOTIV", U&)
         JJ& = JJ& + 1
         Call GRAREG("TAMOTIV", X$, JJ&)
      Next U&
      Call SETULTREG("TAMOTIV", JJ&)
      Call CIERRARCH("TAMOTIV")
    End If
  End If
  '
End Sub

Private Sub Command7_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Text6.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command8_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text7.TEXT), "Formularios de Impresion")
End Sub

Private Sub Form_Load()
  '
  Call CENTRAFORM(Me)
  OPRECHLIN.Text7.TEXT = CONTROL$("", "RECHALIN")  'RECHAZOS EN LINEA
  '
  Combo1.Clear
  Dim PRX As Printer
  For Each PRX In Printers
      Combo1.AddItem PRX.DeviceName
      ' La define como predeterminada del sistema.
      'Set Printer = prx
  Next
  Combo1.TEXT = Printer.DeviceName
  Combo1.TEXT = CONTROL$("", "RUIMPREC") 'RUTA IMPRESIÓN DE RECHAZOS
  '
End Sub

Private Sub Text7_Change()
    '
    VDEVU$ = Text7
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Text6.TEXT = ECOARCH$("INDIFRM", Text7.TEXT)
    End If
    '
End Sub

Private Sub Combo1_LostFocus()
    '
    Dim PRX As Printer
    For I% = 1 To Printers.Count
        If Printers(I% - 1).DeviceName = Combo1.TEXT Then
            ' La define como predeterminada del sistema.
            Set Printer = Printers(I% - 1)
            GoTo 10
        End If
    Next
    '
10  Combo1.TEXT = Printer.DeviceName
    '
End Sub
