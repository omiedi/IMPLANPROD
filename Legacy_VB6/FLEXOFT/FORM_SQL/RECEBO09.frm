VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form RECEBO09 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción Bombas para Evaluar Reparación"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   615
   ClientWidth     =   10860
   Icon            =   "RECEBO09.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   10860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12000
      Begin VB.CommandButton Command6 
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
         Left            =   2000
         TabIndex        =   35
         Top             =   1440
         Width           =   195
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00400000&
         Height          =   2895
         Left            =   9840
         ScaleHeight     =   2835
         ScaleWidth      =   1065
         TabIndex        =   29
         Top             =   5520
         Width           =   1125
         Begin VB.CommandButton Command3 
            Caption         =   "Help"
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
            Index           =   0
            Left            =   180
            Picture         =   "RECEBO09.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   840
            Width           =   650
         End
         Begin VB.CommandButton Command2 
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
            Left            =   180
            Picture         =   "RECEBO09.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   120
            Width           =   650
         End
         Begin VB.CommandButton Command14 
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
            Height          =   650
            Left            =   180
            Picture         =   "RECEBO09.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Grabar Recepción"
            Top             =   1560
            Width           =   650
         End
         Begin VB.Image Image3 
            Height          =   390
            Left            =   -240
            Picture         =   "RECEBO09.frx":0DC8
            Top             =   2220
            Width           =   1515
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Sitio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   1560
         Left            =   7980
         TabIndex        =   27
         Top             =   3730
         Width           =   2745
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
            Height          =   1140
            Left            =   210
            MultiLine       =   -1  'True
            TabIndex        =   28
            Text            =   "RECEBO09.frx":2CEA
            Top             =   300
            Width           =   2325
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Remito Número"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   1095
         Left            =   7980
         TabIndex        =   25
         Top             =   2600
         Width           =   2745
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
            Height          =   315
            Left            =   210
            TabIndex        =   26
            Text            =   " "
            Top             =   510
            Width           =   2325
         End
      End
      Begin VB.CommandButton Command1 
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
         Left            =   1095
         TabIndex        =   24
         Top             =   1095
         Width           =   195
      End
      Begin VB.TextBox Text10 
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
         Left            =   8595
         TabIndex        =   17
         Text            =   " "
         Top             =   1095
         Width           =   2130
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Serie Número"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   900
         Left            =   7980
         TabIndex        =   14
         Top             =   1575
         Width           =   2745
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
            Height          =   315
            Left            =   210
            MultiLine       =   -1  'True
            TabIndex        =   0
            Text            =   "RECEBO09.frx":2CEC
            Top             =   360
            Width           =   2325
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   3480
         Left            =   90
         TabIndex        =   7
         Top             =   1815
         Width           =   7680
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
            Height          =   2985
            Left            =   315
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Top             =   360
            Width           =   7155
         End
      End
      Begin VB.TextBox Text9 
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
         Left            =   9615
         TabIndex        =   1
         Text            =   " "
         Top             =   315
         Width           =   960
      End
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10560
         TabIndex        =   2
         Top             =   315
         Width           =   180
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2685
         Left            =   120
         TabIndex        =   30
         ToolTipText     =   "Doble Click Para Editar o Agregar"
         Top             =   5520
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   4736
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   16501405
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   195
         TabIndex        =   37
         Top             =   1440
         Width           =   1830
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   2205
         TabIndex        =   36
         Top             =   1440
         Width           =   5535
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Materiales Suministrados por el Cliente:"
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
         Left            =   120
         TabIndex        =   34
         Top             =   5280
         Width           =   3495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   5220
         TabIndex        =   22
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   7200
         TabIndex        =   21
         Top             =   1095
         Width           =   1380
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   5220
         TabIndex        =   20
         Top             =   735
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   5745
         TabIndex        =   19
         Top             =   315
         Width           =   2220
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   5745
         TabIndex        =   18
         Top             =   705
         Width           =   2220
      End
      Begin VB.Label REPRESUP 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   225
         Left            =   6510
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   4935
         TabIndex        =   15
         Top             =   1785
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   9615
         TabIndex        =   13
         Top             =   705
         Width           =   1110
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   2205
         TabIndex        =   12
         Top             =   1095
         Width           =   2850
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   1365
         TabIndex        =   11
         Top             =   1080
         Width           =   750
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   1365
         TabIndex        =   10
         Top             =   705
         Width           =   3690
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   1365
         TabIndex        =   9
         Top             =   315
         Width           =   3690
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   200
         TabIndex        =   8
         Top             =   1095
         Width           =   870
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   420
         Picture         =   "RECEBO09.frx":2CEE
         ToolTipText     =   "Click Aquí para cambiar Modalidad de Facturación"
         Top             =   105
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Talon Nro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   7770
         TabIndex        =   6
         Top             =   735
         Width           =   1800
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   8820
         TabIndex        =   5
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   330
         Left            =   105
         TabIndex        =   4
         Top             =   735
         Width           =   960
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuver 
      Caption         =   "Ver"
      Begin VB.Menu mnuverrec 
         Caption         =   "Visualizar Recepción"
      End
      Begin VB.Menu mnuVisPlanReparacion 
         Caption         =   "Visualizar Planilla de Reparación"
      End
   End
End
Attribute VB_Name = "RECEBO09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   Call SELECHO("DEUDOR1")
   NCLIE% = Val(VALACT1$("DEUDOR1"))
   If NCLIE% > 0 Then
      Label11 = TRIM$(Str$(NCLIE%))
   End If
   Command1.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   CODD$ = Label9
   
   If TRIM$(CODD$) = "" Then 'NECESITA EL CODIGO PARA IMPRIMIR LA PLANILLA DE REPARACIÓN
     'SI NO LA GENERA AHORA NO SE IMPRIME SOLO LA PODRA GENERAR CUANDO SE GRABE EL PRESUPUESTO.
      OPX% = COMALTER%("No Se Ha Seleccionado Un Código de Reparación\Si Continúa No Se Generará la Planilla de Reparación\\Cancelar\Continuar")
      If OPX% < 2 Then GoTo 99
   End If
   
   Call GRAPRESU(OKX%)
   If OKX% > 0 Then
     CANCELPRINT% = 0
     Call GRABAR_MATSUMINISTRADOS
     Call PRIPRESU
     Call IMPRIME_Operaciones_Reparacion(XVALO(Label18), Label9, Text1.TEXT, Text10.TEXT, XVALO(Label11))    'IMPRESION FANTASMA
     Call BLANFORMU
     Label18 = ""
     Command14.Enabled = True
     Label11 = ""
     Label18 = TRIM$(Str$(NUPRESUP&))
     Label10 = ""
     CONDI$ = "NUMREPA >= 0 ORDER BY NUMREPA"
     Call CARGALISEL("PREPAEMI", "REPARA", "NumRepa/I10;Referen/A128", CONDI$)
     Call FRESHECHO("PREPAEMI")
     Hide
   End If
99 Command14.Enabled = True
End Sub

Sub GRABAR_MATSUMINISTRADOS()
   
    NEWID& = PROXID_DETAMASU
    VUELTA% = 0
    For i& = 1 To MSF1.Rows - 1
       CIXI% = XVALO(MSF1.TextMatrix(i&, 1))
       Codigo$ = MSF1.TextMatrix(i&, 2)
       DESCRI$ = MSF1.TextMatrix(i&, 3)
       CANTI = XVALO(MSF1.TextMatrix(i&, 4))
       Call MOVI_SUMINISTRADOS(NEWID&, CIXI%, "BR", CANTI, XVALO(Label11), XVALO(Label18), HOY(HO$), "Recepción de Materiales Suministrados X Cliente")
       If VUELTA% < 1 Then
          TTXYZ$ = Space$(92)
          NROREP& = XVALO(Label18)
          HOII% = HOY(HOS$)
          Call REPLA(TTXYZ$, HOS$, 1, 8)
          Call REPLA(TTXYZ$, USERTER$, 69, 24)
          Call REPLA(TTXYZ$, "Detalle de Material Suministrados por el Cliente:", 11, 56)
          Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
          
          TTXYZ$ = Space$(92)
          Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
          
          Call REPLA(TTXYZ$, AJUSTD(SAFETEXT$("Cant."), 6), 11, 6)
          Call REPLA(TTXYZ$, "Código", 18, 18)
          Call REPLA(TTXYZ$, "Descripción", 37, 20)
          Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
          VUELTA% = 1
        End If
        
        TTXYZ$ = Space$(92)
        Call REPLA(TTXYZ$, AJUSTD(SAFETEXT$(CANTI), 6), 11, 6)
        Call REPLA(TTXYZ$, Codigo$, 18, 18)
        Call REPLA(TTXYZ$, DESCRI$, 37, 20)
        Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
    Next i&
    Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, String$(88, "="))
    '

    FILTX$ = TRIM$(Str$(NROREP&)) & ";" & TRIM$(NEWID&)
    Call STDFORM("BRE1-ARC", FILTX$)

End Sub
Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Unload Me
    Hide
End Sub

Sub IMPRIME_Operaciones_Reparacion(NROPRES&, CODD$, Nserie$, Referencia$, NC)
    '
    FORIPRE$ = TRIM$(CONTROL$("FOREPAR", "DOCUREP0"))
    If FORIPRE$ = "" Then Exit Sub
    
    TIPODOC$ = "Operaciones de Reparación"
    DESTIDOC% = NC
    NUMEDOC$ = "PR-" + TRIM$(Label18) '"7" n.presupuesto
    '
    CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = CODD$ 'codigo externo
    CODPARAM$(2) = "REF-MAT1": DOCPARAM$(2) = Nserie$ 'n.serie
    CODPARAM$(3) = "MEDIDA-3": DOCPARAM$(3) = Referencia$ 'referencia
    CODPARAM$(4) = "DES-MPRI": DOCPARAM$(4) = Label8 ' ver vacio
    CODPARAM$(5) = "TOT-TIEM": DOCPARAM$(5) = "" 'HORMINSE$(3600 * Val(Label19)) 'estimado de horas es el label19
    CAPARDOC% = 5
    '
    CODTABU1$(1) = "NOP-OPER"
    CODTABU1$(2) = "OBS-OPER"
    CODTABU1$(3) = "TIE-PROD"
    CACOLTAB1% = 3
    '
    CAITAB1% = 0
    RFFX$ = RECFILT$("SECOPREP", 1, CODD$)
    URN& = ULTREG&("OBSOPREP")
    For I1& = 1 To Len(RFFX$) Step 4
      i& = CVS(Mid$(RFFX$, I1&, 4))
      X$ = REGLEIDO$("SECOPREP", i&)
      NOPE% = CVI(Mid$(X$, 17, 2))
      DOPE$ = Mid$(X$, 19, 48)
      TIEX = CVS(Mid$(X$, 73, 4))
      CAOPS = CVS(Mid$(X$, 97, 4)): If CAOPS < 1 Then CAOPS = 1
      '
      RNC$ = RECFILT$("OBSOPREP", 1, CODD$)
      '
      For U& = 1 To Len(RNC$) Step 4
        RENOTA& = CVS(Mid$(RNC$, U&, 4))
        If RENOTA& > 0 Then
          If RENOTA& <= URN& Then
            TTZZ$ = REGLEIDO$("OBSOPREP", RENOTA&)
            NOPY% = CVI(Mid$(TTZZ$, 17, 2))
            If NOPY% = NOPE% Then
              DOPE$ = DOPE$ + Mid$(TTZZ$, 19)
            End If
          End If
        End If
      Next U&
      '
      CAITAB1% = CAITAB1% + 1
      TETABU1$(1, CAITAB1%) = Str$(NOPE%)
      TETABU1$(2, CAITAB1%) = DOPE$
      TETABU1$(3, CAITAB1%) = HORMINSE$(3600 * TIEX)
      '
      Call CARLINSEP
    Next I1&
    Call CIERRARCH("*.*")
    '
    Call PRINTDOC("@" + FORIPRE$)
    
End Sub



Private Sub Command6_Click()
   Command6.Enabled = False
   Call SELECHO("INDIREP")
   COD$ = TRIM$(VALACT1$("INDIREP"))
   If COD$ <> "" Then
     Label9 = COD$
   End If
   Screen.MousePointer = 1
99 Command6.Enabled = True
End Sub

Private Sub Command9_Click()
   Call SELECHO("SELFECHA")
   Text9.TEXT = VALACT1$("SELFECHA")
   FF% = FECHANUM(Text9.TEXT)
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$
    '
    CUHOR = Val(CONTROL$("REPARA", "VALHORA"))
    COEFI = Val(CONTROL$("REPARA", "COEFPREA"))
    Label20 = CSTRING$(MKS$(CUHOR * COEFI), 4, 8, 2, 2)
    '
    Label18 = TRIM$(Str$(NUPRESUP&))
    '
    Call CARGA_ENCABEZADO(Me.MSF1, "CI/F0;Código/A14;Descripción/A20;Cantidad/F5.1", Me)
    Call CREATABLA_DETAMASU
End Sub
'

Private Sub Form_Resize()
   If WindowState = 1 Then
      Call CIERRARCH("*.*")
      Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Label11 = ""
   Label18 = "": Label10 = ""
   Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Openforms = DoEvents
   Wend
End Sub

Private Sub Label11_Change()
   '
   NC% = Val(Label11)
   Label12 = ""
   Label13 = ""
   Label14 = ""
   Label15 = ""
   Label16 = ""
   Label17 = ""
   '
   If NC% > 0 Then
     If rasocli$(NC%) <> "" Then
       Label12 = rasocli$(NC%): Label12.Refresh
       Label13 = DOMICLI$(NC%)
       Label14 = CPOSCLI$(NC%)
       Label15 = LOCACLI$(NC%)
       Label16 = TEL1CLI$(NC%)
       Label17 = FAXCLI$(NC%)
       DoEvents
     End If
   End If
   '
End Sub

Private Sub Label18_Change()
   '
   NPX& = Val(Label18)
   Call BLANFORMU
'   For UI& = ULTREG&("PRESREPA") To 1 Step -1
'     XX$ = REGLEIDO$("PRESREPA", UI&)
'     If CVS(Left$(XX$, 4)) = NPX& Then
'       REPRESUP = TRIM$(Str$(UI&))
'       NCLIE% = CVI(Mid$(XX$, 67, 2))
'       Label11 = TRIM$(Str$(NCLIE%))
       '
        CONDI$ = "NumRepa=" & NPX& & " "
        Label11 = VALOBADA("REPARA", "Nume_Cli", "CONDI$", "NOMESS")
'       CII% = CVI(Mid$(XX$, 69, 2))
'       Label10 = CODEXT$(CII%)
'       TEPRUEBA = DESCRIT$(CII%)
'       Text1 = Mid$(XX$, 71, 15)
        Text1 = VALOBADA("REPARA", "NumeroSerie", "CONDI$", "NOMESS")
'       Text10 = Mid$(XX$, 86, 9)
        Text10 = VALOBADA("REPARA", "Referen2", "CONDI$", "NOMESS")
'       TALON = Mid$(XX$, 95, 6)
'       Text28 = TRIM$(Mid$(XX$, 129, 368))
        Text28 = VALOBADA("REPARA", "Observa", "CONDI$", "NOMESS")
'       Label19 = CSTRING$(Mid$(XX$, 101, 4), 4, 8, 2, 2)
'       If CVS(Mid$(XX$, 105, 4)) > 0.5 Then
'         Label20 = CSTRING$(Mid$(XX$, 105, 4), 4, 8, 2, 2)
'       End If
       '
       'Call CAREPUES
       'Call PREREPUES
       '
       On Error Resume Next
       Text10.SetFocus
       On Error GoTo 0
       Openforms = DoEvents
       Exit Sub
'     End If
'   Next UI&
   '
End Sub


Private Sub Label9_Change()
    CODD$ = AJUSTI$(Label9, 16)
    REGRE& = REGBUS&("INDIREP", 1, CODD$)
    '
    Screen.MousePointer = 11
    XX$ = REGLEIDO$("INDIREP", REGRE&)
    Label8 = TRIM$(Mid$(XX$, 17, 64))

End Sub

Private Sub Label9_DblClick()
     Label9 = ""
End Sub


Private Sub mnuverrec_Click()
   Call verRecepcion
End Sub
Sub verRecepcion(Optional TALON&)
   Screen.MousePointer = 11
   If TALON& < 1 Then
        
        CONDI$ = "NuOrdRep >= 0 ORDER BY NUMREPA"
        Call CARGALISEL("PREPAEMI", "REPARA", "NumRepa/I10;Referen/A128", CONDI$)
10      Call SELECHO("PREPAEMI")
         Screen.MousePointer = 1
        '
        TALONX& = XVALO(VALACT1$("PREPAEMI"))
   Else
        TALONX& = TALON&
   End If
   If TALONX& < 1 Then Exit Sub
   '
   'AVERIGUO EL PRIMER ID DE LA RECEPCION SEGUN EL TALON EN LA TABLA DETAMASU
   CONDI$ = "TALON = " & TALONX& & " ORDER BY ID ASC"
   PRIMERID& = XVALO(VALOBADA("DETAMASU", "ID", CONDI$, "NOMESS"))

   FILTX$ = TRIM$(Str$(TALONX&)) & ";" & TRIM$(PRIMERID&)
   Call STDFORM("BRE1-ARC", FILTX$)
   
   'SI ES LLAMADA SIN NUMERO DE TALON VUELVE A LA LISTA DE SELECCION
   If TALON& < 1 Then GoTo 10
   Screen.MousePointer = 1
End Sub

Private Sub mnuVisPlanReparacion_Click()
3   Call SELECHO("PREPAEMI/Indice Presupuestos de Reparación")
    NPX& = Val(VALACT1$("PREPAEMI"))

    If NPX& < 1 Then GoTo 99

    '
4   NUDOCU$ = TRIM$(Str$(NPX&))
    Screen.MousePointer = 11
    TIDOCUM$ = "Presupuesto de Reparación " + NUDOCU$
    TIDOCUM$ = "Operaciones de Reparación PR-" + NUDOCU$
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 5
      End If
    Next UI&
    '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
    On Error Resume Next
    GoTo 99
    '
5 Call MUESTRADOC(DOCUNU&)
99
End Sub

Private Sub MSF1_Click()
   '
End Sub

Private Sub MSF1_DblClick()
   Call SETULTREG("!LIPECOTI", 0)
   For i& = 1 To MSF1.Rows - 1
      CIXI% = XVALO(MSF1.TextMatrix(i&, 1))
      Codigo$ = MSF1.TextMatrix(i&, 2)
      DESCRI$ = MSF1.TextMatrix(i&, 3)
      CANTI = XVALO(MSF1.TextMatrix(i&, 4))
      Observa$ = MSF1.TextMatrix(i&, 5)
     '
      YY$ = REGBLAN$("LIPECOTI")
      Call REPLA(YY$, MKI$(CIXI%), 1, 2)
      Call REPLA(YY$, MKS$(CANTI), 3, 4)
      Call REPLA(YY$, MKI$(0), 7, 2)
      Call REPLA(YY$, Observa$, 9, 55)
      JJ& = JJ& + 1
      Call GRAREG("!LIPECOTI", YY$, JJ&)
      '
    Next i&
    Call SETULTREG("!LIPECOTI", JJ&)
    Call CIERRARCH("!LIPECOTI")
    '
    PRF$ = "MATESUMI"
    ATRI$ = "/TITUPAN=MATERIALES SUMINISTRADOS POR EL CLIENTE"
    VTB% = VENTABU%("MATESUMI" + ATRI$)
    If VTB% >= 0 Then
      '
      
      FIN& = ULTREG&("!LIPECOTI")
      For J& = 1 To FIN&
        X$ = REGLEIDO$("!LIPECOTI", J&)
        CODI1% = CVI(Mid$(X$, 1, 2))
        If CODI1% < 1 Then
           Call COMUNI("Código en Fila: " & J& & " Inválido o Inexistente")
           Exit Sub
        End If
        
        CAN1 = CVS(Mid$(X$, 3, 4))
        If CAN1 <= 0 Then
           Call COMUNI("Falta Cantidad en Fila: " & J&)
           Exit Sub
        End If
      Next J&
      '
      MSF1.Rows = 1
      FIN& = ULTREG&("!LIPECOTI")
      For J& = 1 To FIN&
        X$ = REGLEIDO$("!LIPECOTI", J&)
        CODI1% = CVI(Mid$(X$, 1, 2))
        CODI2$ = CODEXT$(CODI1%)
        DESC2$ = DESCRIT0(CODI1%)
        CAN1 = CVS(Mid$(X$, 3, 4))
        CAN2$ = FORMATNUM$(CAN1, "F9.1")
        OBS2$ = Mid$(X$, 9, 55)
        
        'AGREGO A LA GRILLA
        MSF1.Rows = J& + 1
        MSF1.TextMatrix(J&, 1) = CODI1%
        MSF1.TextMatrix(J&, 2) = CODI2$
        MSF1.TextMatrix(J&, 3) = DESC2$
        MSF1.TextMatrix(J&, 4) = CAN2$
      Next J&
      '
    End If

End Sub


Private Sub TEPRUEBA_Change()
   'DESCRI$ = TEPRUEBA.Caption
   'If TEPRUEBA.Width < Label8.Width Then
   '  DESCRI$ = Chr$(10) + Chr$(13) + DESCRI$
   'End If
   'Label8.Caption = DESCRI$
End Sub

Function VALICLIE%()
   VCC% = 0
   If Val(Label11) > 0 Then
     If Val(Label11) <= 32767 Then
       NCLIE% = Val(Label11)
       If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
         VCC% = 1
       End If
     End If
   End If
   '
99 VALICLIE% = VCC%
   '
End Function
'
Sub CAREPUES()
   '
   Call SETULTREG("!MATEROF", 0)
   NPX& = Val(Label18)
   NPI% = NPX& Mod 10000
   RFF$ = RECFILT$("REPUREPA", 2, MKI$(NPI%))
   '
   URE& = ULTREG&("REPUREPA")
30 For KK& = 1 To Len(RFF$) Step 4
     REG& = CVS(Mid$(RFF$, KK&, 4))
     If REG& > 0 Then
       If REG& <= URE& Then
         XX$ = REGLEIDO$("REPUREPA", REG&)
         If CVS(Left$(XX$, 4)) = NPX& Then
           CII% = CVI(Mid$(XX$, 7, 2))
           CANTI = CVS(Mid$(XX$, 9, 4))
           '
           zz$ = REGBLAN$("MATEROF")
           Call REPLA(zz$, MKI$(CII%), 1, 2)
           Call REPLA(zz$, UNIMED$(CII%), 3, 2)
           Call REPLA(zz$, MKS$(CANTI), 5, 4)
           Call REPLA(zz$, MKD$(CDbl(CANTI)), 9, 8)
           Call REPLA(zz$, MKS$(0), 17, 4)
           Call REGAPP("!MATEROF", zz$)
           '
         End If
       End If
     End If
   Next KK&
   Call CIERRARCH("!MATEROF")
   '
End Sub
'
Sub PREREPUES()
   '
   List1.Clear
   SUMA# = 0
   SUMAREP = ""
   COEFI = Val(CONTROL$("REPARA", "COEFPREA"))
   For U& = 1 To ULTREG&("!MATEROF")
     zz$ = REGLEIDO$("!MATEROF", U&)
     TTXX$ = Space$(128)
     CII% = CVI(Left$(zz$, 2))
     Call REPLA(TTXX$, CODEXT$(CII%), 1, 16)
     Call REPLA(TTXX$, DESCRIT$(CII%), 17, 38)
     NEBRUX$ = Mid$(zz$, 5, 4)
     Call REPLA(TTXX$, CSTRING$(NEBRUX$, 4, 10, 1, 2), 54, 10)
     COSUX$ = MKS$(COEFI * COSUNI(CII%))
     Call REPLA(TTXX$, CSTRING$(COSUX$, 4, 10, 2, 2), 64, 10)
     ITOX$ = MKS$(CVS(NEBRUX$) * CVS(COSUX$))
     Call REPLA(TTXX$, CSTRING$(ITOX$, 4, 12, 2, 2), 74, 12)
     SUMA# = SUMA# + Val(Mid$(TTXX$, 74, 12))
     List1.AddItem TTXX$
   Next U&
   SUMAREP = CSTRING$(MKD$(SUMA#), 4, 11, 2, 2)
'  '
End Sub
'
Sub BLANFORMU()
   '
   Label8 = ""
   Label9 = ""
   Label10 = ""
   Label12 = ""
   Label13 = ""
   Label14 = ""
   Label15 = ""
   Label16 = ""
   Label17 = ""
   Text1 = ""
   Text10 = ""
   TALON = ""
   TEPRUEBA = ""
   Text28 = ""
   Label19 = ""
   Label20 = ""
   Text2 = ""
   Text3 = ""
   '
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   '
   Me.MSF1.Rows = 1
End Sub
'

Private Sub Text9_GotFocus()
   On Error Resume Next
   Command9.SetFocus
   On Error GoTo 0
End Sub


Sub PRIPRESU()
    '

    Exit Sub
    
    FORIPRE$ = TRIM$(CONTROL$("FOREPAR", "DOCUREP5"))
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Imposible Imprimir Recepción de Reparación.\Falta Definir Formulario Electrónico.")
       Exit Sub
    End If
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Recepción de Reparación.\No existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
10  NC% = Val(Label11)
    TIPODOC$ = "Recepción de Reparación"
    DESTIDOC% = NC%
    NUMEDOC$ = TRIM$(Label18)
    FECHDOC$ = TRIM$(Text9.TEXT)
    '
    CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = DEPOSAL$
    CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = Text2
    CODPARAM$(3) = "N-AUTENT": DOCPARAM$(3) = Text3
    CAPARDOC% = 3
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "COLOR"
    CODTABU1$(4) = "UNIMED"
    CODTABU1$(5) = "NRO-LOTE"
    CODTABU1$(6) = "CANTIDAD"
    CODTABU1$(7) = "N-OCOMPR"
    CODTABU1$(8) = "CAKILOS"
    CODTABU1$(9) = "F-PLANO"
    CODTABU1$(10) = "FECH-REC"
    CACOLTAB1% = 10
    '
    CAITAB1% = 1
    TETABU1$(2, CAITAB1%) = "S/Nro." + TRIM$(Text1) + " - Ref." + TRIM$(Text10)
    TETABU1$(6, CAITAB1%) = "1"
    '
    Call FRATEXTO(Text28, 46)
    For KKI% = 1 To CARETEX%
      CAITAB1% = CAITAB1% + 1
      TETABU1$(2, CAITAB1%) = RETEX$(KKI%)
    Next KKI%                           '
    '
    '
    Call PRINTDOC("@" + FORIPRE$)
    '
End Sub

Sub GRAPRESU(OKX%)
   '
   OKX% = 0
'   Call BLOQARCH("PRESREPA")
   '
'   REG& = Val(REPRESUP)
'   If REG& < 1 Then REG& = 1 + ULTREG&("PRESREPA")
'   If REG& < 1 Or REG& > 1 + ULTREG&("PRESREPA") Then
'      Call COMUNI("Imposible Registrar - Error de Indices")
'      GoTo 98
'   End If
   '

   Screen.MousePointer = 11
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLIE% = Val(Label11)
   '
   If TRIM$(Text1) = "" Then
      MERRO$ = "Falta Número de Serie\o Identificación del Equipo."
      GoTo 98
   End If
   '
   If Len(TRIM$(Text1)) > 15 Then
        MERRO$ = "Exceso Longitud Número de Serie"
        Text1 = Left$(TRIM$(Text1), 15)
        GoTo 98
   End If
   '
   If Len(TRIM$(Text10)) > 24 Then
        MERRO$ = "Exceso Longitud Referencia"
        Text10 = Left$(TRIM$(Text10), 9)
        GoTo 98
   End If
   '
   If TRIM$(Text3) = "" Then
      MERRO$ = "Falta Sitio de Ingreso de la Bomba."
      GoTo 98
   End If
   '
   HOII% = HOY(HOS$)
   FEMI% = FECHANUM(Text9.TEXT)
   If FEMI% < 1 Or FEMI% > HOII% Then
      MERRO$ = "Fecha de Emisión no Válida"
      GoTo 98
   End If
   '
'   XX$ = REGBLAN$("PRESREPA")
   REFEPRE$ = "S/Nro." + Text1 + " - " + Left$(TRIM$(Label12), 30) + " - R: " + TRIM$(Text10) + " - " + TRIM$(Text9)
   NPLAL& = Val(Label18)
   ' Grabacion en Tabla SQL
   SQLX$ = "SELECT * FROM REPARA"
   SQLX$ = SQLX$ + " WHERE Nume_Cli < 0"
   Dim PRESUPUESTO As ADODB.Recordset
   Set PRESUPUESTO = New ADODB.Recordset
   PRESUPUESTO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With PRESUPUESTO
     .AddNew
     !CodiEmpr = CodiEmp%
     !NUMREPA = Val(Label18)
     !Referen = Left$(REFEPRE$, 128)
     !nume_cli = Val(Label11)
     !raso_cli = Left$(Label12, 64)
     !domi_cli = Left$(Label13, 64)
     !loca_Cli = Left$(Label15, 64)
     !Tele_Cli = Left$(Label16, 24)
     !Pais_Cli = ""
      FF% = FECHANUM("01/01/80")
     !FechEnSol = CVDAT(FF%)
     !Mail_Cli = ""
     !Accesorios = ""
     !Defec_Infor = ""
     !Trab_Efect = ""
     !PresuEsti = 0
     !Status = 0
     !Usuario_Repu = ""
     !Estima_Horas = 0
     !FOr_Compra = CVDAT(FF%)
     !FECHENT = CVDAT(FF%)
     !Nom_Tra = ""
     !Cuit_Cli = ""
     !Piva_Cli = ""
     !Conductor = ""
     !Patente = 0
     !Lentre_Cli = ""
     !Status_Gar = 0
     !CausaFalla = 0
     !DeCauFalla = ""
     !TipoFalla = 0
     !DeTipFalla = ""
     !cod_inte = 0
     !NumRemito = Left$(Text2, 18)
     !Or_Compra = ""
     !Descripcion = ""
     !NUMEROSERIE = Left$(Text1, 32)
      FE% = FECHANUM(Text9)
     !FechRecep = CVDAT(FE%)
     !Referen2 = Text10
     !HoraMObra = 0
     !ValMObra = 0
     !Imp_Mobra = 0
     !Imp_Repu = 0
     !Imp_Tot = 0
     !FechPres = CVDAT(FF%)
     !FechApro = CVDAT(FF%)
     !NROPED = 0
     !Observa = Text28
     !AnotaRep = ""
     !NUORDREP = 0
     !RefordRep = ""
     !FechOrdRep = CVDAT(FF%)
     !CodRepa = ""
     !Sitio = Left$(Text3, 50)
     .Update
   End With
   PRESUPUESTO.Close
   Set PRESUPUESTO = Nothing
   '
   TTXYZ$ = Space$(92)
   NROREP& = Val(Label18)
   HOII% = HOY(HOS$)
   Call REPLA(TTXYZ$, HOS$, 1, 8)
   Call REPLA(TTXYZ$, USERTER$, 69, 24)
   Call REPLA(TTXYZ$, "Recepción Material para Presupuestación de Reparación:", 11, 56)
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   TTXYZ$ = Space$(92)
   Call REPLA(TTXYZ$, Space$(56), 11, 56)
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   Call REPLA(TTXYZ$, "Cliente: " + Label12, 11, 56)
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   Call REPLA(TTXYZ$, "    Tel: " + TRIM$(Label16), 11, 56)
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   Call REPLA(TTXYZ$, "    Fax: " + TRIM$(Label17), 11, 56)
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   '
   Call REPLA(TTXYZ$, Space$(56), 11, 56)
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   Call REPLA(TTXYZ$, "  S/Nro: " + Text1, 11, 56)
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   Call REPLA(TTXYZ$, "Referen: " + Text10, 11, 56)
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   Call REPLA(TTXYZ$, Space$(56), 11, 56)
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   '
   Call FRATEXTO(Text3, 46)
   For KK% = 1 To CARETEX%
     TTXYZ$ = Space$(92)
     Call REPLA(TTXYZ$, RETEX$(KK%), 20, 46)
     If KK% = 1 Then Call REPLA(TTXYZ$, "  Sitio: ", 11, 9)
     Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   Next KK%
   '
   Call REPLA(TTXYZ$, Space$(56), 11, 56)
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   '
   Call FRATEXTO(Text28, 46)
   For KKI% = 1 To CARETEX%
     TTXYZ$ = Space$(92)
     Call REPLA(TTXYZ$, RETEX$(KKI%), 20, 46)
     If KKI% = 1 Then Call REPLA(TTXYZ$, " Observ: ", 11, 9)
     Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, TTXYZ$)
   Next KKI%
   '
   Call PRESRE09.AGREGA_ANOTAREPARA(NROREP&, String$(88, "="))
'   Call REPLA(XX$, MKS$(NPLAL&), 1, 4)
'   Call REPLA(XX$, REFEPRE$, 5, 60)
'   Call REPLA(XX$, MKI$(NCLIE%), 67, 2)
'   Call REPLA(XX$, TRIM$(Text1), 71, 15)
'   Call REPLA(XX$, TRIM$(Text10), 86, 9)
'   Call REPLA(XX$, TRIM$(TALONX$), 95, 6)
'   Call REPLA(XX$, MKS$(0), 101, 4)
'   Call REPLA(XX$, MKS$(0), 105, 4)
'   Call REPLA(XX$, MKS$(0), 109, 4)
'   Call REPLA(XX$, MKS$(0), 113, 4)
'   Call REPLA(XX$, MKS$(0), 117, 4)
'   Call REPLA(XX$, MKI$(0), 121, 2)   ' NO SE EMITIO PRESUPUESTO
'   Call REPLA(XX$, Text28 + Space$(368), 129, 368)
'   Call REPLA(XX$, "", 497, 16)
'   Call GRAREG("PRESREPA", XX$, REG&)
'   Call CIERRARCH("PRESREPA")
'   '
'   NPLAL& = Val(Label18)
'   NPLAI% = NPLAL& Mod 10000
'   '
'   TTXYZ$ = Space$(92)
'   NROREP& = Val(Label18)
'   HOII% = HOY(HOS$)
'   Call REPLA(TTXYZ$, HOS$, 1, 8)
'   Call REPLA(TTXYZ$, USERTER$, 69, 24)
'   Call REPLA(TTXYZ$, "Recepción Material para Presupuestación de Reparación:", 11, 56)
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
'   '
'   TTXYZ$ = Space$(92)
'   Call REPLA(TTXYZ$, Space$(56), 11, 56)
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
'   Call REPLA(TTXYZ$, "Cliente: " + Label12, 11, 56)
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
'   Call REPLA(TTXYZ$, "    Tel: " + TRIM$(Label16), 11, 56)
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
'   Call REPLA(TTXYZ$, "    Fax: " + TRIM$(Label17), 11, 56)
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
'   '
'   Call REPLA(TTXYZ$, Space$(56), 11, 56)
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
'   Call REPLA(TTXYZ$, "  S/Nro: " + Text1, 11, 56)
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
'   Call REPLA(TTXYZ$, "Referen: " + Text10, 11, 56)
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
'   Call REPLA(TTXYZ$, Space$(56), 11, 56)
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + TTXYZ$)
   '
'   Call FRATEXTO(Text28, 46)
'   For KKI% = 1 To CARETEX%
'     TTXX$ = Space$(92)
'     Call REPLA(TTXX$, RETEX$(KKI%), 10, 46)
'     If KKI% = 1 Then Call REPLA(TTXX$, " Observ: ", 1, 9)
'     Call REGAPP("ANOTAREP", Left$(MKS$(NROREP&) + Space$(10) + TTXX$, 96))
'   Next KKI%
'   '
'   Call REGAPP("ANOTAREP", MKS$(NROREP&) + String$(88, "="))
'   Call CIERRARCH("ANOTAREP")
   '
   
   Call FRESHECHO("PRESREPA")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   OKX% = 1
   Screen.MousePointer = 1
   Exit Sub
   '
98 Call CIERRARCH("PRESREPA")
   Call CIERRARCH("ANOTAREP")
   '
   Screen.MousePointer = 1
   Call MENSERR(24, MERRO$)
   '
End Sub

Function NUPRESUP&()
'   '
'   NPP& = 1
'   '
'   FIN& = ULTREG&("PRESREPA")
'   If FIN& > 0 Then
'      For U& = FIN& To FIN& - 128 Step -1
'         If U& < 1 Then GoTo 99
'         XX$ = REGLEIDO$("PRESREPA", U&)
'         NPPX& = CVS(Mid$(XX$, 1, 4))
'         If NPPX& >= NPP& Then
'            NPP& = 1 + NPPX&
'         End If
'      Next U&
'   End If
'   '
   NPP& = 1 + XVALO(VALOBADA("REPARA", "Max(NumRepa)", "NumRepa > 0", "NOMESS"))
   '
99 NUPRESUP& = NPP&
   '
End Function
'

