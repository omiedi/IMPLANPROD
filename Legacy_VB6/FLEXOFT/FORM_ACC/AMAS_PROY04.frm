VERSION 5.00
Begin VB.Form AMAS_PROY04 
   BackColor       =   &H00D0F0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proyectos - A-B-M-L "
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   ControlBox      =   0   'False
   Icon            =   "AMAS_PROY04.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Index           =   1
      Left            =   105
      TabIndex        =   21
      Top             =   210
      Width           =   10740
      Begin VB.TextBox Text5 
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
         Height          =   300
         Left            =   2415
         TabIndex        =   34
         Text            =   " "
         Top             =   1320
         Width           =   4455
      End
      Begin VB.CommandButton Command10 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2100
         TabIndex        =   33
         Top             =   1320
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.TextBox Text4 
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
         Height          =   300
         Left            =   1470
         TabIndex        =   2
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text2 
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
         Height          =   285
         Left            =   9700
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   900
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
         Height          =   285
         Left            =   3045
         Picture         =   "AMAS_PROY04.frx":030A
         TabIndex        =   22
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         MaxLength       =   8
         TabIndex        =   0
         Top             =   420
         Width           =   1620
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1470
         TabIndex        =   1
         Top             =   840
         Width           =   5415
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7035
         TabIndex        =   39
         Top             =   1360
         Width           =   1380
      End
      Begin VB.Label Label6 
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
         Height          =   285
         Left            =   8610
         TabIndex        =   38
         Top             =   1320
         Width           =   2010
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "15/06/06"
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
         Left            =   9700
         TabIndex        =   37
         Top             =   420
         Width           =   900
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         TabIndex        =   35
         Top             =   1365
         Width           =   1380
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Cierre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7980
         TabIndex        =   29
         Top             =   920
         Width           =   1380
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Alta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7980
         TabIndex        =   27
         Top             =   480
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   26
         Top             =   465
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Denominación:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         TabIndex        =   25
         Top             =   920
         Width           =   1380
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   6440
         TabIndex        =   24
         Top             =   420
         Width           =   435
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5250
         TabIndex        =   23
         Top             =   465
         Width           =   1005
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004040&
      Height          =   8100
      Left            =   11025
      ScaleHeight     =   8040
      ScaleWidth      =   1005
      TabIndex        =   18
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command20 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "AMAS_PROY04.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Búsqueda por Denominación, Descripción General - Especificación, etc"
         Top             =   4425
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Undo"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "AMAS_PROY04.frx":197A
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Restablece los Valores Anteriores"
         Top             =   3795
         Width           =   650
      End
      Begin VB.CommandButton Command21 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "AMAS_PROY04.frx":1C84
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Alta de Nuevo Proyecto"
         Top             =   2505
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "More"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "AMAS_PROY04.frx":1DCE
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Opciones Varias"
         Top             =   5985
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   30
         Top             =   7035
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   31
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "AMAS_PROY04.frx":1F18
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   3150
         Width           =   650
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
         Height          =   640
         Left            =   105
         Picture         =   "AMAS_PROY04.frx":2222
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   8080
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command24 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "AMAS_PROY04.frx":252C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Baja Lógica Individual"
         Top             =   9345
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command23 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "AMAS_PROY04.frx":2836
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Sustitución de Código de Repuesto"
         Top             =   8700
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command22 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "AMAS_PROY04.frx":2980
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Alta de Nuevo Repuesto  por Copia Completa de Datos de Uno Existente"
         Top             =   8055
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "AMAS_PROY04.frx":2ACA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Ayuda en Línea"
         Top             =   900
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Index           =   4
         Left            =   105
         Picture         =   "AMAS_PROY04.frx":2DD4
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Exit"
         Top             =   210
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -255
         Picture         =   "AMAS_PROY04.frx":2F1E
         Stretch         =   -1  'True
         Top             =   7350
         Width           =   1410
      End
   End
   Begin VB.Frame MARCOFICHA 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5715
      Index           =   0
      Left            =   105
      TabIndex        =   9
      Top             =   2205
      Width           =   10740
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   2310
         ScaleHeight     =   375
         ScaleWidth      =   8145
         TabIndex        =   15
         Top             =   630
         Width           =   8205
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción General - Especificaciones"
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
            TabIndex        =   16
            Top             =   105
            Width           =   7680
         End
      End
      Begin VB.CommandButton Command7 
         Height          =   260
         Left            =   2310
         TabIndex        =   14
         Top             =   5280
         Width           =   3240
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Left            =   9180
         TabIndex        =   13
         Top             =   5280
         Width           =   1320
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Pegar"
         Height          =   260
         Left            =   7905
         TabIndex        =   12
         Top             =   5280
         Width           =   1275
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Copiar"
         Height          =   260
         Left            =   6735
         TabIndex        =   11
         Top             =   5280
         Width           =   1170
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Cortar"
         Height          =   260
         Left            =   5550
         TabIndex        =   10
         Top             =   5280
         Width           =   1170
      End
      Begin VB.TextBox Text28 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4230
         Left            =   2310
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1050
         Width           =   8205
      End
      Begin VB.Line Line1 
         X1              =   10515
         X2              =   10515
         Y1              =   630
         Y2              =   5535
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción y Especificaciones:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   210
         TabIndex        =   17
         Top             =   600
         Width           =   1905
      End
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   12000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu archiv 
      Caption         =   "Archivos"
      Begin VB.Menu Abri 
         Caption         =   "Abrir"
      End
      Begin VB.Menu php0 
         Caption         =   "-"
      End
      Begin VB.Menu nueite 
         Caption         =   "Nuevo Proyecto"
      End
      Begin VB.Menu altcop 
         Caption         =   "Alta por Copia"
         Visible         =   0   'False
      End
      Begin VB.Menu PHP2 
         Caption         =   "-"
      End
      Begin VB.Menu guar 
         Caption         =   "Guardar"
      End
      Begin VB.Menu php1 
         Caption         =   "-"
      End
      Begin VB.Menu Sali 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Edi 
      Caption         =   "Edición"
      Begin VB.Menu deshcam 
         Caption         =   "Deshacer Cambios"
      End
      Begin VB.Menu PHP202 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSearch 
         Caption         =   "Buscar"
      End
      Begin VB.Menu php3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu SuCoAr 
         Caption         =   "Sustitución de Código de Repuesto"
         Visible         =   0   'False
      End
      Begin VB.Menu BaLoIn 
         Caption         =   "Baja Lógica Individual"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu CoyLi 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu LiCaProDet 
         Caption         =   "Listado de Cargos por Proyecto"
         Begin VB.Menu MnuLiDet 
            Caption         =   "Opción Detalle"
         End
         Begin VB.Menu LiAcum 
            Caption         =   "Opción Acumulado"
         End
      End
   End
   Begin VB.Menu Config 
      Caption         =   "Accesos Directos"
      Begin VB.Menu MnuClien 
         Caption         =   "Maestro de Clientes"
      End
      Begin VB.Menu php5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOrfabri 
         Caption         =   "Ordenes de Fabricación"
      End
   End
   Begin VB.Menu Ayu 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "AMAS_PROY04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CIACTI%(15)
Dim ARCHEX$(15)
Dim GRABATODO%
Dim YACAR%
Dim MODIFIC%

Private Sub Abri_Click()
  Call Command9_Click
End Sub
Private Sub Ayu_Click()
  '
  Call Command4_Click
  '
End Sub

Private Sub Command1_Click()
  '
  Command1.Enabled = False
    LISPROY04.Show 1
  Command1.Enabled = True
  '
End Sub

Private Sub Command10_Click()
   Call SELECHO("DEUDOR1")
   NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE% > 0 Then
     If REGICLI&(NCLIE%) > 0 Then
       Text4.TEXT = TRIM$(Str$(NCLIE%))
       Text5.TEXT = RASOCLI$(NCLIE%)
     End If
   End If
End Sub

Private Sub Command11_Click()
   Call SELECHO("TATIMAT")
   VALONUE$ = VALACT1$("TATIMAT")
   If VALONUE$ <> "" Then
      Text6.TEXT = VALONUE$
   End If
End Sub

Private Sub Command12_Click()
  Call GRABAPROY
End Sub
Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command2_Click()
   '
   If COMALTER%("Realmente Desea Descartar los Cambios ?\\No, Continuar\Si, Descartar Cambios") <> 2 Then
      Exit Sub
   End If
   '
   YACAR% = 0
   MODIFIC% = 0
   '
   NUPROY1$ = Text1.TEXT
   Text1.TEXT = ""
   Call BLANFORMU
   Text1.TEXT = NUPROY1$
   '
End Sub

Private Sub Command20_Click()
   Static ARGU$
   Command20.Enabled = False
10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda de Proyecto por Texto", ARGU$)))
   If ARGU$ <> "" Then
     Screen.MousePointer = 11
     JJ& = 0
     Call COPYSTRU("INDIPROY", "SELPROY")
     '
     Call ABREPLANCUEN
     On Error Resume Next
     With TaProyec
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           If !Stat_Proy < 0 Then GoTo 15 'no presenta proyectos de baja
           For KKI% = 1 To .Fields.Count
             On Error Resume Next
             VATEX$ = ""
             VATEX$ = UCase(.Fields(KKI% - 1).Value) 'UCASE!!!
             On Error GoTo 0
             If InStr(VATEX$, ARGU$) > 0 Then
               Call REPLA(XX$, !Cod_Proy, 1, 8)
               Call REPLA(XX$, TRIM$(FANTCLI$(!Clien_Proy)) & " - " & !Deno_Proy, 9, 48)
               '
               JJ& = JJ& + 1
               Call GRAREG("!SELPROY", XX$, JJ&)
               Exit For
             End If
           Next KKI%
15       .MoveNext
         Loop
       End If
     End With
     Call SETULTREG("!SELPROY", JJ&)
     '
     Screen.MousePointer = 1
     If JJ& < 1 Then
         Call MENSERR(24, "No se Encontraron Coincidencias")
         GoTo 10
       ElseIf JJ& = 1 Then
         NUPROY1$ = TRIM$(Mid$(REGLEIDO$("!SELPROY", 1), 1, 8))
         Text1.TEXT = NUPROY1$
       Else
         Call FRESHECHO("!SELPROY")
         Call SELECHO("!SELPROY")
         NPROYX$ = VALACT1$("!SELPROY")
         If NPROYX$ <> "" Then
           Text1 = NPROYX$
         End If
     End If
   End If
   Command20.Enabled = True
End Sub

Private Sub Command21_Click()
   '
   Command21.Enabled = False
   '
   If DERACCE%("ALTANUPRO", "Alta de Nuevo Proyecto") < 2 Then
      Exit Sub
   End If
   '
   Screen.MousePointer = 1
   '
   VDEF$ = Space$(48) + String$(11, 0)
   HOII% = HOY(HOS$)
   Call REPLA(VDEF$, MKI$(HOII%), 59, 2)
   '
9  OBJ$ = OBJPLA$("NUEPROY", VDEF$)
   If OBJ$ = "" Then GoTo 99
   VDEF$ = OBJ$
   '
   NUPROY$ = TRIM$(Mid$(OBJ$, 1, 8))
   DESCRIPROY$ = TRIM$(Mid$(OBJ$, 9, 48))
   NUCLIPROY% = CVI(Mid$(OBJ$, 57, 2))
   RASOCLIPROY$ = RASOCLI$(NUCLIPROY%)
   FECPROY% = CVI(Mid$(OBJ$, 59, 2))
   '
   '//////////////VALIDACIONES////////////
   If TRIM$(NUPROY$) = "" Then
      Call COMUNI("Código de Proyecto no válido")
      GoTo 9
   End If
   '
   Call ABREPLANCUEN
   TaProyec.FindFirst "Cod_Proy = '" & NUPROY$ & "' "
   If TaProyec.NoMatch = False Then
      Call COMUNI("Código de Proyecto no Válido o Pre-Existente")
      GoTo 9
   End If
   '\\\Me fijo si el Proyecto ya existe como Centro de Costos
   EXISPRO$ = FILTERGETRECORD$("TACENCOS", 1, NUPROY$)
   If TRIM$(EXISPRO$) <> "" Then
      Call COMUNI("Código de Proyecto no Válido o Pre-Existente")
      GoTo 9
   End If
   '
   If TRIM$(DESCRIPROY$) = "" Then
      Call COMUNI("Debe Ingresar una Descripción\ para el Proyecto")
      GoTo 9
   End If
   '
   If NUCLIPROY% <= 0 Then
      Call COMUNI("Cliente no Válido.")
      GoTo 9
   End If
   '
   If TRIM$(RASOCLI$(NUCLIPROY%)) = "" Then
      Call COMUNI("Cliente no Válido.")
      GoTo 9
   End If
   '
   JJJ = CVINT(Date)
   If FECPROY% > JJJ Then
      Call COMUNI("La Fecha de Generación\ no puede ser Superior\ al dia de la Fecha.")
      GoTo 9
   End If
   '
   Call ABREPLANCUEN
   With TaProyec
      .AddNew
        !Cod_Proy = NUPROY$
        !Deno_Proy = DESCRIPROY$
        !Clien_Proy = NUCLIPROY%
        !FeAlt_Proy = CVDAT(FECPROY%)
        !Stat_Proy = 0 'GENERADA
        !Descrip_Proy = ""
      .Update
   End With
   'Agrego en Tacencos el Proyecto como Centro de Costos
   XXZ$ = Space(48)
   Call REPLA(XXZ$, NUPROY$, 1, 8)
   Call REPLA(XXZ$, Left$(DESCRIPROY$, 35), 9, 35)
   Call REPLA(XXZ$, MKI$(FECPROY%), 45, 2) 'Fecha Vigente 'Desde'
   Call REGAPP("TACENCOS", XXZ$)
   Call CIERRARCH("TACENCOS")
   '
   Call INDIPROY
   Text1 = TRIM$(NUPROY$)
   '
99 Screen.MousePointer = 1
   Command21.Enabled = True
   '
End Sub
Private Sub Command3_Click(Index As Integer)
   Call CIERRARCH("*.*")
   'Call BAJALDISCO
   Call FINAL("")
End Sub
Private Sub Command39_Click()
   Command39.Enabled = False
   PRINTFORM
   Command39.Enabled = True
End Sub

Private Sub Command38_Click()
  '
  Command38.Enabled = False
    LISPROY04.Show 1
  Command38.Enabled = True
  '
End Sub

Private Sub Command4_Click()
  '
  Command4.Enabled = False
    Call HELPONLINE("AMAPROY")
  Command4.Enabled = True
  '
End Sub

Private Sub Command5_Click()
   Text28.SelStart = 0
   Text28.SelLength = Len(Text28.TEXT)
   Text28.SetFocus
End Sub
Private Sub Command6_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Clipboard.GetText + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
End Sub

Private Sub Command9_Click()
   '
   Call SELECHO("INDIPROY")
   NUOPROY$ = TRIM$(VALACT1$("INDIPROY"))
   Text1.TEXT = TRIM$(NUOPROY$)
   '
End Sub
Private Sub DeshCam_Click()
  Call Command2_Click
End Sub
Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    Screen.MousePointer = 11
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FUERASYS
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    HOII% = HOY(HOS$)
    Label5.Caption = HOS$
    Label6 = USERNAME
    '
    Refresh
    '
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
    Screen.MousePointer = 1
    '
End Sub


Sub BLANFORMU()
   '
   Text3.TEXT = ""
   Label28.Caption = ""
   Text4.TEXT = ""
   Text28.TEXT = ""
   Text31 = ""
   Text32 = ""
   Text27 = ""
   Label40 = ""
   Label41 = ""
   Label43 = ""
   Text11 = ""
   YACAR% = 0
   '
End Sub

Private Sub guar_Click()
  Call Command12_Click
End Sub

Private Sub impri_Click()
Call Command39_Click
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label5.ToolTipText = Label5
End Sub

Private Sub MayEq_Click()
  SETGEMA04.Option1(5).Value = True
  Call SETGEMA04.Command5_Click
End Sub

Private Sub MnuPlant_Click()
  SETGEMA04.Option1(1).Value = True
  Call SETGEMA04.Command5_Click
End Sub

Private Sub LiAcum_Click()
  '
  Call LISPROY04.Command3_Click
  '
End Sub

Private Sub MnuClien_Click()
  '
  Call CIERRARCH("*.*")
  ACONEC$ = "ARCHIG04/AMACLI"
  Call CONECRUN(ACONEC$, "Padron Maestro de Clientes")
  '
End Sub

Private Sub MnuLiDet_Click()
  '
  Call LISPROY04.Command8_Click
  '
End Sub

Private Sub MnuOrfabri_Click()
  '
  If EXISTE%("ORFABR04.EXE") > 0 Then
    Call CONECRUN("ORFABR04", "Ordenes de Fabricación")
  End If
  '
End Sub

Private Sub MnuSearch_Click()
  '
  Call Command20_Click
  '
End Sub

Private Sub nueite_Click()
  Call Command21_Click
End Sub

Private Sub Sali_Click()
  Call Command3_Click(4)
End Sub
Private Sub Text1_Change()
   '
   Call BLANFORMU
   If TRIM$(Text1.TEXT) <> "" Then
      NUPROY$ = TRIM$(Text1.TEXT)
      Call CARGOPROY
   End If
   '
End Sub

Private Sub Text1_DblClick()
   '
   Call Command9_Click
   '
End Sub
Private Sub Text25_DblClick()
   Call SELECHO("SELFECHA")
   VALONUE$ = VALACT1$("SELFECHA")
   If VALONUE$ <> "" Then
      Text25.TEXT = VALONUE$
   End If
End Sub

Private Sub Text2_gotfocus()
  On Error Resume Next
    Text1.SetFocus
  On Error GoTo 0
End Sub

Private Sub Text28_Change()
  MODIFIC% = 1
End Sub

Private Sub Text3_Change()
  MODIFIC% = 1
End Sub

Private Sub Text4_Change()
   NC1% = XVALO(Text4.TEXT)
   Text5.TEXT = RASOCLI$(NC1%)
   MODIFIC% = 1
End Sub

Private Sub Text4_GotFocus()
  '
  On Error Resume Next
    Text1.SetFocus
  On Error GoTo 0
  '
End Sub

Private Sub Text5_GOTFOCUS()
  On Error Resume Next
    Text1.SetFocus
  On Error GoTo 0
End Sub
   '
Function VALICLIE%()
   VCC% = 0
   If XVALO(Text4.TEXT) > 0 Then
     If XVALO(Text4.TEXT) <= 32767 Then
       NCLIE% = XVALO(Text4.TEXT)
       If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
         VCC% = 1
       End If
     End If
   End If
   '
99 VALICLIE% = VCC%
   '
End Function
Sub INDIPROY() 'GENERA EL INDICE DE PROYECTOS
    Screen.MousePointer = 11
    Call ABREPLANCUEN
    Set PROYTMP = Contable.OpenRecordset("SELECT * FROM TAPROYEC  ORDER BY DENO_PROY ASC;", dbOpenDynaset, dbReadOnly)
    On Error Resume Next
    PROYTMP.MoveLast
    If Err Then Exit Sub
    '
    CAREC% = PROYTMP.RecordCount
    If CAREC% < 1 Then Exit Sub
    '
    FIN& = CAREC%
    REBLB$ = REGBLAN$("INDIPROY")
    JJ& = 0
    With PROYTMP
      .MoveFirst
      '
10    Call TRACE(20, JJ&, FIN&)
      JJ& = JJ& + 1
      NUCLIPROY% = !Clien_Proy
      EL2$ = REBLB$
      Call REPLA(EL2$, !Cod_Proy, 1, 8)
      Call REPLA(EL2$, TRIM$(FANTCLI$(NUCLIPROY%)) & " - " & !Deno_Proy, 9, 48)
      Call GRAREG("INDIPROY", EL2$, JJ&)
      '
      If JJ& < CAREC% Then
        PROYTMP.MoveNext
        GoTo 10
      End If
    End With
    '
    PROYTMP.Close
    Call SETULTREG("INDIPROY", JJ&)
    Call FILESORT("INDIPROY", "", 1, 1, "ASC")
    Call FRESHECHO("INDIPROY")
    Call BAJALDISCO
    Screen.MousePointer = 1
End Sub
Sub CARGOPROY() 'CARGA LA FICHA DE UN PROYECTO DETERMINADO
   '
   NUPROY$ = TRIM$(Text1)
   If NUPROY$ = "" Or ECOARCH("INDIPROY", Text1) = "" Then Exit Sub
   '
   Call ABREPLANCUEN
   SQLX$ = "SELECT * FROM TAPROYEC"
   Set OTTMP = Contable.OpenRecordset(SQLX$, 4)
   With OTTMP
     .FindFirst "  Cod_Proy = '" & NUPROY$ & "' "
     If .NoMatch = True Then
        On Error GoTo 0
        Exit Sub
     End If
     On Error GoTo 0
     '
     Text3 = !Deno_Proy
     '
     NCLIE% = !Clien_Proy
     Text4 = NCLIE%
     If NCLIE% > 0 Then
       Text5 = RASOCLI(NCLIE%)
     End If
     '
     Label5 = FECHATEX(CVINT(!FeAlt_Proy))
     Text2.TEXT = FECHATEX(CVINT(!FeCie_Proy))
     Label28 = CStr(!Stat_Proy)
     Text28 = !Descrip_Proy
     MODIFIC% = 0
     YACAR% = 1
   End With
   '
End Sub
Sub GRABAPROY()
   '
   If YACAR% < 1 Or MODIFIC% < 1 Then
     Exit Sub
   End If
   '
   If DERACCE%("GRABAPROY", "Grabar Proyectos") < 2 Then
       Exit Sub
   End If
   'VALIDACIONES
   '
   If TRIM$(Text1) = "" Then
      Call COMUNI("Número de Proyecto no válido.")
      Exit Sub
   End If
   '
   If ECOARCH$("INDIPROY", Text1) = "" Then
      Call COMUNI("Número de Proyecto no válido.")
      Exit Sub
   End If
   '
   If TRIM$(Text3) = "" Then
      Call COMUNI("Falta Denominación del Proyecto.")
      Exit Sub
   End If
   '
   If TRIM$(Text28) = "" Then
     Call COMUNI("Falta Especificación del Proyecto.")
     Exit Sub
   End If
   '
   NUPROY$ = TRIM$(Text1.TEXT)
   If NUPROY$ = "" Then Exit Sub
   '
   Call ABREPLANCUEN
   SQLX$ = "SELECT * FROM TAPROYEC"
   Set OTTMP = Contable.OpenRecordset(SQLX$, 2)
   With OTTMP
     .FindFirst "  Cod_Proy = '" & NUPROY$ & "' "
     If .NoMatch = True Then
        On Error GoTo 0
        Exit Sub
     End If
     On Error GoTo 0
     '
     .Edit
        !Descrip_Proy = Text28.TEXT
        !Deno_Proy = Text3.TEXT
          NUCLIEN1% = XVALO(Text4.TEXT)
        !Clien_Proy = NUCLIEN1%
     .Update
   End With
   '
   MODIFIC% = 0
   Call INDIPROY
   '
End Sub
