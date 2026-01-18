VERSION 5.00
Begin VB.Form CFGISO9K 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración General Sub-Sistema de Calidad"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6345
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
   ScaleHeight     =   6480
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Comprobante Despachos Internos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   210
      TabIndex        =   26
      Top             =   4830
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
         Height          =   300
         Left            =   2415
         TabIndex        =   30
         Text            =   " "
         Top             =   325
         Width           =   1380
      End
      Begin VB.CommandButton Command10 
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
         TabIndex        =   29
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
         Height          =   300
         Left            =   210
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   3750
      End
      Begin VB.CommandButton Command9 
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
         Height          =   740
         Left            =   4095
         Picture         =   "CFGISO9K.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   315
         Width           =   855
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
         TabIndex        =   31
         Top             =   450
         Width           =   2325
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Comprobante Interno de Entrega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   210
      TabIndex        =   20
      Top             =   3255
      Width           =   5160
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
         Height          =   740
         Left            =   4095
         Picture         =   "CFGISO9K.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   315
         Width           =   855
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
         Left            =   210
         TabIndex        =   23
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
         Height          =   300
         Left            =   3780
         TabIndex        =   22
         Top             =   325
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
         Height          =   300
         Left            =   2415
         TabIndex        =   21
         Text            =   " "
         Top             =   325
         Width           =   1380
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
         TabIndex        =   25
         Top             =   450
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Comprobante Demanda de Acción Correctiva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   210
      TabIndex        =   14
      Top             =   1680
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
         Height          =   740
         Left            =   4095
         Picture         =   "CFGISO9K.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   315
         Width           =   855
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
         Height          =   300
         Left            =   210
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   " "
         Top             =   750
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
         Height          =   300
         Left            =   3780
         TabIndex        =   16
         Top             =   325
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
         Height          =   300
         Left            =   2415
         TabIndex        =   15
         Text            =   " "
         Top             =   325
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
         TabIndex        =   19
         Top             =   450
         Width           =   2325
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Comprobante de Orden de Fabricación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   210
      TabIndex        =   8
      Top             =   105
      Width           =   5160
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
         Height          =   300
         Left            =   2415
         TabIndex        =   12
         Text            =   " "
         Top             =   325
         Width           =   1380
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
         TabIndex        =   11
         Top             =   325
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
         Height          =   300
         Left            =   210
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   750
         Width           =   3750
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
         Height          =   740
         Left            =   4095
         Picture         =   "CFGISO9K.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   315
         Width           =   855
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
         TabIndex        =   13
         Top             =   450
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   6840
      Left            =   5460
      ScaleHeight     =   6780
      ScaleWidth      =   1110
      TabIndex        =   4
      Top             =   0
      Width           =   1170
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   32
         Top             =   5670
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command25 
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
         Picture         =   "CFGISO9K.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Acceso A Tablas Causas\Destinos\Categorias"
         Top             =   4725
         Width           =   645
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
         Height          =   600
         Left            =   105
         Picture         =   "CFGISO9K.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   840
         Width           =   645
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
         Height          =   610
         Left            =   105
         Picture         =   "CFGISO9K.frx":107C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   190
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "CFGISO9K.frx":11C6
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5985
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Presentación de Costo y Precios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   0
      Top             =   7770
      Width           =   5160
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Calculado"
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
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Precio Base (C.Rep. + 15 %)"
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
         Left            =   3360
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Informado"
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
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1905
      End
   End
End
Attribute VB_Name = "CFGISO9K"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
    FORSELEC.ABUSQUE = Text8
    Call SELECHO("INDISQR/Formularios de Impresión")
    CODFOR$ = VALACT1$("INDISQR")
    If TRIM$(CODFOR$) = "" Then
        Text8.SetFocus
      Else
        Text8.TEXT = CODFOR$
    End If

End Sub

Private Sub Command2_Click()

   Call GRACONTROL("", "FORMOFAB", Text6.TEXT)
   Call GRACONTROL("", "DEMANACC", Text1.TEXT)
   Call GRACONTROL("", "COMPINTE", Text3.TEXT)
   Call GRACONTROL("", "DESPINTE", Text8.TEXT)
   
End Sub

Private Sub Command25_Click()

    Command25.Enabled = False
    OPX% = ALTERNA%("Tabla de Causas - Motivos\Tabla de Destinos\Tabla de Categorias")
     
    If OPX% = 1 Then
        Call TABLACAUSA
      ElseIf OPX% = 2 Then
        Call TABLADESTINO
      ElseIf OPX% = 3 Then
        Call TABLACATEGORIA
    End If
    
99  Command25.Enabled = True
End Sub

Sub TABLACAUSA()

    Call TRALOCAL("TAMOTI1", "TAMOTI1")
    '
    ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 1400))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
    VTB% = VENTABU%("TAMOTI1/" + ATRI$)
    If VTB% >= 0 Then
       JJ& = 0
      For U& = 1 To ULTREG&("!TAMOTI1")
         X$ = REGLEIDO$("!TAMOTI1", U&)
         JJ& = JJ& + 1
         Call GRAREG("TAMOTI1", X$, JJ&)
      Next U&
      Call SETULTREG("TAMOTI1", JJ&)
      Call CIERRARCH("TAMOTI1")
    End If
    
End Sub
Sub TABLADESTINO()
       '
  'TABLA DE DESTINO DE RECHAZO
  If ULTREG&("TADESTI") < 1 Then
    Call BLANARCH("TADESTI")
    Call REGAPP("TADESTI", Chr$(0) + "Retrabajo")
    Call REGAPP("TADESTI", Chr$(1) + "Baja")
    Call REGAPP("TADESTI", Chr$(2) + "Devolución")
    Call CIERRARCH("TADESTI")
  End If
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

   
End Sub
Sub TABLACATEGORIA()
    Call COPYSTRU("TAMOTIV", "TACATEGO")
    
    Call TRALOCAL("TACATEGO", "")
    '
    ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 1400))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
    VTB% = VENTABU%("TACATEGO/" + ATRI$)
    If VTB% >= 0 Then
       JJ& = 0
      For U& = 1 To ULTREG&("!TACATEGO")
         X$ = REGLEIDO$("!TACATEGO", U&)
         JJ& = JJ& + 1
         Call GRAREG("TACATEGO", X$, JJ&)
      Next U&
      Call SETULTREG("TACATEGO", JJ&)
      Call CIERRARCH("TACATEGO")
    End If
    

End Sub

Private Sub Command3_Click()

    FORSELEC.ABUSQUE = Text1
    Call SELECHO("INDISQR/Formularios de Impresión")
    CODFOR$ = VALACT1$("INDISQR")
    If TRIM$(CODFOR$) = "" Then
        Text1.SetFocus
      Else
        Text1.TEXT = CODFOR$
    End If

End Sub

Private Sub Command4_Click()

    FORSELEC.ABUSQUE = Text6
    Call SELECHO("INDISQR/Formularios de Impresión")
    CODFOR$ = VALACT1$("INDISQR")
    If TRIM$(CODFOR$) = "" Then
        Text6.SetFocus
      Else
        Text6.TEXT = CODFOR$
    End If

End Sub

Private Sub Command5_Click()
   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1.TEXT), "Formularios de Impresion")

End Sub

Private Sub Command6_Click()

   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text6.TEXT), "Formularios de Impresion")

End Sub

Private Sub Command7_Click()
    FORSELEC.ABUSQUE = Text3
    Call SELECHO("INDISQR/Formularios de Impresión")
    CODFOR$ = VALACT1$("INDISQR")
    If TRIM$(CODFOR$) = "" Then
        Text1.SetFocus
      Else
        Text3.TEXT = CODFOR$
    End If

End Sub

Private Sub Command8_Click()
   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text3.TEXT), "Formularios de Impresion")

End Sub

Private Sub Command9_Click()
   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text8.TEXT), "Formularios de Impresion")

End Sub

Private Sub Form_Load()
   Text6.TEXT = CONTROL$("", "FORMOFAB")
   Text1.TEXT = CONTROL$("", "DEMANACC")
   Text3.TEXT = CONTROL$("", "COMPINTE")
   Text8.TEXT = CONTROL$("", "DESPINTE")
End Sub

Private Sub Text1_Change()

    CODFOR$ = TRIM$(Text1)
    Text2 = ECOARCH$("INDISQR", CODFOR$)

End Sub

Private Sub Text3_Change()
    CODFOR$ = TRIM$(Text3)
    Text4 = ECOARCH$("INDISQR", CODFOR$)

End Sub

Private Sub Text6_Change()

    CODFOR$ = TRIM$(Text6)
    Text7 = ECOARCH$("INDISQR", CODFOR$)

End Sub

Private Sub Text8_Change()
    CODFOR$ = TRIM$(Text8)
    Text5 = ECOARCH$("INDISQR", CODFOR$)

End Sub
