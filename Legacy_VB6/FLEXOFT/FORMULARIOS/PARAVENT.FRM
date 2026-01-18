VERSION 5.00
Begin VB.Form PARAVENT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DEFINICION FORMULARIO DE VENTANA"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   " Operacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   4560
      TabIndex        =   22
      Top             =   2280
      Width           =   1935
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "Insertar y Borrar:"
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
         Left            =   360
         TabIndex        =   41
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Agregar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Posicion y Operacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   4335
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
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
         Left            =   3300
         MaxLength       =   8
         TabIndex        =   38
         Top             =   1200
         Width           =   900
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
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
         Left            =   1215
         MaxLength       =   8
         TabIndex        =   36
         Top             =   1200
         Width           =   900
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
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
         Left            =   3300
         MaxLength       =   8
         TabIndex        =   34
         Top             =   840
         Width           =   900
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
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
         Left            =   1215
         MaxLength       =   8
         TabIndex        =   32
         Top             =   840
         Width           =   900
      End
      Begin VB.TextBox Text7 
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
         Left            =   3780
         MaxLength       =   8
         TabIndex        =   20
         Top             =   360
         Width           =   420
      End
      Begin VB.TextBox Text6 
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
         Left            =   2295
         MaxLength       =   8
         TabIndex        =   18
         Top             =   360
         Width           =   420
      End
      Begin VB.TextBox Text9 
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
         Left            =   885
         MaxLength       =   8
         TabIndex        =   16
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ancho:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2535
         TabIndex        =   39
         Top             =   1230
         Width           =   615
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Altura:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   540
         TabIndex        =   37
         Top             =   1230
         Width           =   570
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B.Izq:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   35
         Top             =   870
         Width           =   510
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B.Super:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   33
         Top             =   870
         Width           =   750
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Col-IZQ:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2955
         TabIndex        =   21
         Top             =   390
         Width           =   720
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rg-FIN:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1515
         TabIndex        =   19
         Top             =   390
         Width           =   675
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rg-INI:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Top             =   390
         Width           =   630
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   3120
      Left            =   6600
      ScaleHeight     =   3060
      ScaleWidth      =   1965
      TabIndex        =   8
      Top             =   0
      Width           =   2025
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
         Height          =   690
         Left            =   120
         Picture         =   "PARAVENT.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   150
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   9
         Top             =   2430
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   10
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "PARAVENT.frx":014A
         Top             =   2655
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6375
      Begin VB.TextBox Text13 
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
         Left            =   4845
         TabIndex        =   30
         Top             =   1305
         Width           =   1380
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
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
         Left            =   3675
         MaxLength       =   8
         TabIndex        =   28
         Top             =   1305
         Width           =   420
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
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
         Left            =   5775
         MaxLength       =   8
         TabIndex        =   26
         Top             =   1680
         Width           =   420
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
         Height          =   285
         Left            =   3240
         TabIndex        =   24
         Top             =   360
         Width           =   175
      End
      Begin VB.CommandButton Command18 
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
         Left            =   3420
         Picture         =   "PARAVENT.frx":206C
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Refresca Lista de Seleccion"
         Top             =   360
         Width           =   300
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
         Left            =   4850
         TabIndex        =   13
         Top             =   360
         Width           =   1380
      End
      Begin VB.TextBox Text4 
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
         Left            =   1245
         TabIndex        =   12
         Top             =   1680
         Width           =   3690
      End
      Begin VB.TextBox Text3 
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
         Left            =   1245
         TabIndex        =   4
         Top             =   705
         Width           =   4980
      End
      Begin VB.TextBox Text1 
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
         Left            =   1245
         MaxLength       =   8
         TabIndex        =   3
         Top             =   360
         Width           =   1380
      End
      Begin VB.TextBox Text2 
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
         Left            =   1245
         TabIndex        =   2
         Top             =   1305
         Width           =   1380
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   6360
         Y1              =   1120
         Y2              =   1120
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EOF$:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4260
         TabIndex        =   31
         Top             =   1365
         Width           =   540
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PEOF%:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2880
         TabIndex        =   29
         Top             =   1365
         Width           =   690
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L.O.R:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5100
         TabIndex        =   27
         Top             =   1710
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salida:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4200
         TabIndex        =   14
         Top             =   420
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".PRF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2640
         TabIndex        =   11
         Top             =   390
         Width           =   525
      End
      Begin VB.Label lblFrom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   765
         Width           =   1080
      End
      Begin VB.Label lblFromName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   420
         TabIndex        =   6
         Top             =   1365
         Width           =   720
      End
      Begin VB.Label lblServer 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   420
         TabIndex        =   5
         Top             =   390
         Width           =   720
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   6600
      Picture         =   "PARAVENT.frx":21B6
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Guarda Configuración de Correo Electrónico"
      Top             =   3240
      Width           =   850
   End
End
Attribute VB_Name = "PARAVENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public APA$

Private Sub Command1_Click()
   Unload Me
End Sub

'
Private Sub Command10_Click()
   Call COPYSTRU("DIRFS", "DIPRF")
   Call SELECHO("DIPRF/DIRECTORIO ARCHIVOS TIPO PRF")
   ASEL$ = VALACT1$("DIPRF")
   If ASEL$ <> "" Then Text1 = ASEL$
End Sub

Private Sub Command18_Click()
   Call ARMALISTARCH(APA$, "PRF", MKI$(6))    ' EL 6 ES POR VENTANA TABULADA
End Sub

Private Sub Command2_Click()
   If Text1 <> "" Then
     If Text2 <> "" Then
       PARAVEN2.Label2 = Text1
       PARAVEN2.Label3 = Text2
       PARAVEN2.Show 1
     End If
   End If
End Sub

Private Sub Form_Load()
   APA$ = App.Path
   APA$ = "C:\SYS_CLIEN\BICICLETA\FLEXOFT\WINDOWS"
   '
End Sub

Sub CREASTRUVEN()
   '
   Dim RGNX$(128)
   RGNX$(0) = String$(128, 0)
   TTXX$ = AJUSTI$("ESTRUCTURA TABLA RANDOM", 47) + Chr$(0) + Chr$(128)
   Call REPLA(RGNX$(0), TTXX$, 1, 49)
   '
   RGNX$(1) = ARMAREN("NOMBRE_DEL_CAMPO", 6, 0, 1, 32, 0)
   RGNX$(2) = ARMAREN("TIPO DE VARIABLE", 1, 0, 33, 1, 0)
   RGNX$(3) = ARMAREN("LONGITUD DECIMAL", 1, 0, 34, 1, 0)
   RGNX$(4) = ARMAREN("POSICION INICIAL", 1, 0, 35, 1, 0)
   RGNX$(5) = ARMAREN("LONGITUD CAMPO", 1, 0, 36, 1, 0)
   RGNX$(6) = ARMAREN("TEXTO ENCABEZADO", 6, 0, 37, 32, 0)
   RGNX$(7) = ARMAREN("ANCHO FISICO COLUMNA", 1, 0, 69, 1, 0)
   RGNX$(8) = ARMAREN("ANCHO LOGICO COLUMNA", 1, 0, 70, 1, 0)
   RGNX$(9) = ARMAREN("CARACTERES MARGEN", 1, 0, 71, 1, 0)
   RGNX$(10) = ARMAREN("JUSTIFICACION (1/2)", 1, 0, 72, 1, 0)
   RGNX$(11) = ARMAREN("EDITABLE (0/1/2)", 1, 0, 73, 1, 0)
   RGNX$(12) = ARMAREN("ARCHIVO ECO", 6, 0, 74, 8, 0)
   RGNX$(13) = ARMAREN("COLUMNA ECO", 1, 0, 82, 1, 0)
   '
   TTXX$ = ""
   For KKII% = 0 To 13
      TTXX$ = TTXX$ + RGNX$(KKII%)
   Next KKII%
   Call SAVEFILE(LUDAT$ + "STRUVEN.RFS", TTXX$)
   '
End Sub

Function ARMAREN$(NOCA$, TIVAX%, LODEC%, PINIX%, LOCAM%, Index%)
   '
   ARE$ = String$(128, 0)
   Call REPLA(ARE$, UCase$(AJUSTI$(NOCA$, 32)), 1, 32)
   Call REPLA(ARE$, Chr$(TIVAX%), 33, 1)
   Call REPLA(ARE$, Chr$(LODEC%), 34, 1)
   Call REPLA(ARE$, Chr$(PINIX%), 35, 1)
   Call REPLA(ARE$, Chr$(LOCAM%), 36, 1)
   Call REPLA(ARE$, Chr$(Index%), 37, 1)
   ARMAREN$ = ARE$
   '
End Function

Private Sub Text1_Change()
   PROCE$ = TRIM$(Text1) + ".PRF"
   XXX$ = LOADFILE$(APA$ + "\" + PROCE$)
   YYY$ = Left$(XXX$, 128)
   '
   Text3 = UCase$(TRIM$(Mid$(YYY$, 1, 48)))
   Text5 = UCase$(TRIM$(Mid$(YYY$, 49, 8)))
   Text2 = UCase$(TRIM$(Mid$(YYY$, 57, 8)))
   '
   LOR1% = Asc(Mid$(YYY$, 106, 1)): If LOR1% > 16 Then LOR1% = 0
   LOR% = 256 * LOR1% + Asc(Mid$(YYY$, 107))
   TexT8 = TRIM$(Str$(LOR%))
   '
       CAMAX& = CVI(Mid$(YYY$, 108))
   PEOF% = 1 + Asc(Mid$(YYY$, 110))
   Text12 = TRIM$(Str$(PEOF%))
   EOFS$ = Mid$(YYY$, 111, 16)
   Text13 = EOFS$
   '
   COD$ = Mid$(YYY$, 81)
   RGI& = Asc(Mid$(COD$, 1)): Text9 = TRIM$(Str$(RGI&))
   RGF& = Asc(Mid$(COD$, 2)): Text6 = TRIM$(Str$(RGF&))
   LEFTPOS% = Asc(Mid$(COD$, 3)): Text7 = TRIM$(Str$(LEFTPOS%))
   '
   '  BLP% = Asc(Mid$(COD$, 4))
   '  VIV% = Asc(Mid$(COD$, 5))
   APPI% = Asc(Mid$(COD$, 6)): Check1.Value = APPI%
   IDL% = Asc(Mid$(COD$, 7)): Check2.Value = IDL%
   '  HDR& = Asc(Mid$(COD$, 8))
   '  RFT% = Asc(Mid$(COD$, 9))
   '  PRX% = Asc(Mid$(COD$, 10))
   '
   DPOSI$ = Mid$(YYY$, 73, 8): If DPOSI$ = Space$(8) Then DPOSI$ = String$(8, 0)
   BSUPER = CVI(Mid$(DPOSI$, 1, 2)): Text14 = TRIM$(Str$(BSUPER))
   BIZQUI = CVI(Mid$(DPOSI$, 3, 2)): Text15 = TRIM$(Str$(BIZQUI))
   BALTUR = CVI(Mid$(DPOSI$, 5, 2)): Text16 = TRIM$(Str$(BALTUR))
   BANCHO = CVI(Mid$(DPOSI$, 7, 2)): Text17 = TRIM$(Str$(BANCHO))
   '
End Sub

Private Sub Text2_Change()
   Text4 = ""
   If Text2 <> "" Then
      ZZZ$ = LOADFILE$(LUDAT$ + Text2 + ".RFS")
      Text4 = UCase$(TRIM$(Left$(ZZZ$, 48)))
   End If
End Sub
