VERSION 5.00
Begin VB.Form PRINREP07 
   BackColor       =   &H00DAE4EB&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Impresion"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00DAE4EB&
      Caption         =   "CONTROL DE IMPRESIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4090
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4530
      Begin VB.Frame Frame5 
         BackColor       =   &H00DAE4EB&
         Caption         =   "ORIENTACIÓN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   210
         TabIndex        =   22
         Top             =   1845
         Width           =   1365
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00DAE4EB&
            Caption         =   "Paisaje"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   150
            TabIndex        =   23
            Top             =   510
            Width           =   960
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00DAE4EB&
            Caption         =   "Retrato"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   24
            Top             =   240
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DAE4EB&
         Caption         =   "COPIAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   3480
         TabIndex        =   20
         Top             =   1845
         Width           =   855
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   200
            TabIndex        =   21
            Text            =   " "
            Top             =   315
            Width           =   435
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DAE4EB&
         Caption         =   "ÁREA DE IMPRESIÓN (MM)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   210
         TabIndex        =   11
         Top             =   2730
         Width           =   4110
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   15
            Text            =   " "
            Top             =   420
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   14
            Text            =   " "
            Top             =   735
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   3150
            TabIndex        =   13
            Text            =   " "
            Top             =   420
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   3150
            TabIndex        =   12
            Text            =   " "
            Top             =   735
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ancho"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   -105
            TabIndex        =   19
            Top             =   490
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Largo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   -105
            TabIndex        =   18
            Top             =   810
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Margen Izquierdo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   1485
            TabIndex        =   17
            Top             =   810
            Width           =   1575
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Margen Superior"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   1540
            TabIndex        =   16
            Top             =   490
            Width           =   1530
         End
      End
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
         TabIndex        =   10
         Text            =   " "
         Top             =   300
         Width           =   4110
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00DAE4EB&
         Caption         =   "TIPO DE LETRA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   210
         TabIndex        =   5
         Top             =   720
         Width           =   3270
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   75
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   " "
            Top             =   630
            Width           =   2985
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   6
            Text            =   " "
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fuente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -270
            TabIndex        =   9
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tamaño"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1560
            TabIndex        =   8
            Top             =   325
            Width           =   855
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00DAE4EB&
         Caption         =   "VARIOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   1620
         TabIndex        =   3
         Top             =   1845
         Width           =   1800
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00DAE4EB&
            Caption         =   "Suprime fecha"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   120
            TabIndex        =   31
            Top             =   480
            Width           =   1620
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H00DAE4EB&
            Caption         =   "Renumera"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1140
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00DAE4EB&
         Caption         =   "ANCHO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   3570
         TabIndex        =   1
         Top             =   720
         Width           =   750
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   150
            TabIndex        =   2
            Top             =   525
            Width           =   450
         End
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   1275
      TabIndex        =   28
      ToolTipText     =   "Click con Boton Derecho para Cambiar la Imagen"
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H002C3D46&
      ForeColor       =   &H80000008&
      Height          =   4530
      Left            =   4800
      ScaleHeight     =   4500
      ScaleWidth      =   2295
      TabIndex        =   25
      Top             =   0
      Width           =   2325
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   29
         Top             =   3600
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
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
         Picture         =   "PRINREP07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Emite Listado"
         Top             =   2835
         Width           =   700
      End
      Begin VB.CommandButton BOTEXIT 
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
         Picture         =   "PRINREP07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Abandona el Listador"
         Top             =   280
         Width           =   700
      End
   End
End
Attribute VB_Name = "PRINREP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IMPREDET$
Dim VXA$(64)

Public TIPRO%
Public ARCHISAL$, ASAL%, ATOT%
Dim SUBTI$(16)
Dim XX As String * 128

Public PRINTERSET%, LARGOJA, ANCHOJA, ALTUREN, TALEBASI, TAFUESEL
Public TOTALI%, POSITOT%, RESUB$, RETRA$, ULTIHOJ%
Dim HAYTOTA%
Public CXC$
Public NUHOJ%, NUHOK%
Dim DATLISTOK%, SOBRESCRI%
Dim SRY As String * 255
Dim GU As String * 255
Dim LII$(10)
Dim CCC$(10)
'
Dim RGFJ%(32), TFF%(32)
Dim VFMIN$(32), VFMAX$(32)
'
Dim VDS$(32, 16), VDX#(32, 16), VD%(32, 16), VFS$(32), VFSA$(32), VFX#(32), VF%(32), FI$(32)
Dim TVAL%(20), CHE$(20), BAN$(20), IMPCHE#(20), FECHE%(20)
Dim VDP%(32), VDPX#(32), VDPS$(32), VDA$(32), VDAA%(32), VDAB#(32)
Dim SUMA#(32, 2), SUPAR#(32)
Dim LEYE$(16)
Dim SUMACOL#(256)
'
Dim CALIMI As Integer
Dim CALIMS As Integer
Dim COEFH As Single
'
Dim MIZQ$
Dim PORIAN
'
Dim YALIS%(48)
Dim YAPRE%
'
Dim COLUAC%
Dim MODIFI%
'
Dim PPZZ As String * 128
Dim FUFIJA$(128)
'
Dim ANCHOPAP, LARGOPAP, MARGENI, MARGENS, ALTULOGO
Dim CACOPIAS%
'
Dim SALEXCEL%
'
Dim TIFORMED%
'
Private Sub BOTEXIT_Click()
   Hide
End Sub

Private Sub Command1_Click()
    '
    Call SELECPRINTER(OKEY%)
    If OKEY% <> 1 Then Exit Sub
    '
    CACOPIAS% = XVALO(Text6)
    If CACOPIAS% < 1 Then CACOPIAS% = 1
    CACOP% = CACOPIAS%
    '
    INTERLIX = 0
    If REHOJX% > 96 Then INTERLIX = Printer.ScaleHeight / REHOJX%
    
'    Call GRADATIMPRE
'    Call CARDATIMPRE
    '
'    Call VERDATLIS
    '
'         Call SETPRINTER
'         Call CARHEADERS
'         Call ARMAENCA
    On Error GoTo 199
    For COP% = 1 To CACOP%
            If TNC% < 1 Then
                'Call ENCAHOJA
             End If
       Call LISTABU
    Next COP%
    On Error GoTo 0
    '
    Call CIERRARCH("*.*")
    PRINTERSET% = 0
    If ARCHISAL$ = "" Then Printer.Orientation = PORIAN
    If COMMA$ <> "" Then
       Call BLOQEXE("*")
    End If
    '
    Hide
    '
    Exit Sub
    '
199 ERNU$ = TRIM$(Str$(Err.Number))
    MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") en la Impresión de\"
    MENSA$ = MENSA$ + "este Listado.\  \"
    '
    MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
    MENSA$ = MENSA$ + "   1.- Impresora Apagada o fuera de Línea.\"
    MENSA$ = MENSA$ + "   2.- Falta Papel o Papel Atorado.\"
    MENSA$ = MENSA$ + "   3.- Problemas en Cable de Señal de Impresora.\"
    MENSA$ = MENSA$ + "   4.- Problemas de Red - No hay Acceso a Impre-\"
    MENSA$ = MENSA$ + "        sora Remota.\"
    MENSA$ = MENSA$ + "   5.- Impresora de Matriz de Puntos Configurada\"
    MENSA$ = MENSA$ + "        para 'Calidad de Imprenta'. Corrija la Con-\"
    MENSA$ = MENSA$ + "        figuración para 'Borrador'.\"
    MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
    '
    c$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(c$, 53, "Error de Impresión de Listado")
    If RESPU% = 4 Then
        Resume
    End If
    Call CIERRARCH("*.*")
    Call FINAL("")
    Close: End
    '
End Sub

Sub INITOTA()
    '
    RENULIS% = 0
    If PRINREP07.Check2.Value = 1 Then RENULIS% = 1
    '
    For K3% = 1 To NCL%
        For K4% = 0 To 2
           SUMA#(K3%, K4%) = 0
        Next K4%
    Next K3%
    '
    NUHOJ% = 0 ' : NUHOK% = 0
    PPXX% = InStr(ATRIBU$, "PRIHOJA=")
    If PPXX% > 0 Then
       NUHOJ% = Val(Mid$(ATRIBU$, PPXX% + 8)) - 1
    End If
    '
    If RENULIS% = 1 Then
       NUHOJ% = Val(InputBox$("Numera Desde Hoja Número", "Numeración de Hojas", "1")) - 1
    End If
    '
End Sub
Sub LISTABU()
    '
'VER COMENTADO SHOWPLA
'    If SHOWPLA07.Visible = True Then
'        Call LISTABUPLA
'        Exit Sub
'    End If
    '
    For K2% = 1 To 256: SUMACOL#(K2%) = 0: Next K2%
    NUHOJ% = 0: NUHOK% = 0
    '
    SUPRIFE% = 0
    If Check1.Value = 1 Then SUPRIFE% = 1
    '
    Call CARDATIMPRE
    '
    ATOT% = 0
    For J& = 1 To SHOWREP07.GRID.Cols
       rr = SHOWREP07.Label8.Count
       ACOLUM%(J&) = XVALO(SHOWREP07.Label8(J&).Caption)
       If SHOWREP07.GRID.ColWidth(J& - 1) < 1 Then
          ACOLUM%(J&) = 0
       End If
       ATOT% = ATOT% + ACOLUM%(J&) + 1
       VXA$(J&) = ""
    Next J&
    LII$(1) = String$(ATOT%, "-")
    LII$(2) = String$(ATOT%, "-")
    LII$(3) = String$(ATOT%, "-")
    '
    Call SETPRINTER
    Call INITOTA
    Call ARMAENCA
    Call ENCAHOJA
    '
    LARGOMAX = LARGOJA
    If ORIENTA% = 1 Then
         If LARGOMAX > LARGOPAP * 56.7 Then
             LARGOMAX = LARGOPAP * 56.7
         End If
      ElseIf ORIENTA% = 2 Then
         If LARGOMAX > ANCHOPAP * 56.7 Then
             LARGOMAX = ANCHOPAP * 56.7
         End If
    End If
    '
    FIN& = XVALO(SHOWREP07.Label7)
    HAYTOTA% = 0
    For KI& = 1 To SHOWREP07.Label4.Count
      If SHOWREP07.Label4(KI& - 1).Visible = True Then TOTALI% = 1
    Next KI&
    '
    For U& = 1 To FIN&
       TTXX$ = "": FINSUBTO% = 0: HUBOCA% = 0
       CURRY = Printer.CurrentY
       Printer.Print MIZQ$ + Space$(OFS%);     ' PRE-IMPRIME MARGEN IZQUIERDO
       ALTUREMAX = 0
       For J& = 1 To SHOWREP07.GRID.Cols
          If SHOWREP07.GRID.Rows < U& + 1 Then SHOWREP07.GRID.Rows = U& + 1
          SHOWREP07.GRID.Row = U&
          SHOWREP07.GRID.COL = J& - 1
          VLXX$ = Left$(SHOWREP07.GRID.TEXT + Space$(ACOLUM%(J&)), ACOLUM%(J&))
          'INTERVENCION PARA CUANDO SE GENERA LA INFO POR FUERA
          If UCase$(App.EXEName) <> "GENREP07" Then
            LETR$ = FORCAMPO$(J& - 1)
            If LETR$ = "F" Or LETR$ = "I" Or LETR$ = "$" Then
               VLXX$ = AJUSTD$(VLXX$, ACOLUM%(J&))
            End If
          End If
          'FIN DE INTERVENCION PARA LA IMPRESION DESDE OTRA GRILLA
          
          SUMACOL#(J&) = SUMACOL#(J&) + XVALO(VLXX$)
          '
          If VLXX$ <> VXA$(J&) Then HUBOCA% = 1
          FCP$ = UCase$(Left$(FORCOLUM$(J&), 1))
          If FCP$ = "$" Then HUBOCA% = 1
          If InStr(FORCOLUM$(J&), "/R") > 0 Then HUBOCA% = 1
          VXA$(J&) = VLXX$
          
          If HUBOCA% = 0 Then VLXX$ = ""
          '
          If TRIM$(VLXX$) = "" Then VLXX$ = Space$(ACOLUM%(J&))
          If TTXX$ <> "" Then VLXX$ = " " + VLXX$
          TTXX$ = TTXX$ + VLXX$
          If TRIM$(VLXX$) <> "" Then
             If SHOWREP07.GRID.CellFontUnderline = True Then
                FINSUBTO% = Len(TTXX$)
             End If
          End If
          '
          If Left$(FORCOLUM$(J&), 3) = "JPG" Then
               Call SHOWREP07.IMPRIMELAIMAGEN(SHOWREP07.GRID, U&, J& - 1)
               FINSUBTO% = ATOT%
             ElseIf Left$(FORCOLUM$(J&), 1) = "X" Then
               Call SHOWREP07.IMPRIMEMULIN(SHOWREP07.GRID, U&, J& - 1)
               FINSUBTO% = ATOT%
             Else
               Printer.Print VLXX$;      ' AGREGADO EPB 17/01/2011 PARA IMPRIMIR POR CAMPO
          End If
          '
19     Next J&
       '
       If Printer.CurrentY > LARGOMAX - (7 + 6 * TOTALI%) * ALTUREN Then
           ULTIHOJ% = 0
           Call LITRANSPO
           '
           Printer.Print MIZQ$ + Space$(OFS%) + LII$(3)
           Call EJECT
           For K2% = 1 To NCL%: VFSA$(K2%) = "": Next K2%
           Printer.NewPage
           Call ENCAHOJA
       End If
       '
       ' Printer.Print MIZQ$ + Space$(OFS%) + TTXX$
       Printer.Print ""                    ' para forzar avance de renglon - EPB
       '
       If ALTUREMAX > Printer.CurrentY - CURRY Then
          Printer.CurrentY = CURRY + ALTUREMAX
       End If
       '
       If U& < FIN& Then
          If FINSUBTO% > 0 Then
             Printer.Print MIZQ$ + Space$(OFS%) + String$(FINSUBTO%, "-")
          End If
       End If
    Next U&
    '
    Call LITOTA
    Printer.Print MIZQ$ + Space$(OFS%) + LII$(3)
    Call EJECT
    Printer.EndDoc
End Sub
'
Sub LISTABUPLA()
    '
    For K2% = 1 To 256: SUMACOL#(K2%) = 0: Next K2%
    NUHOJ% = 0: NUHOK% = 0
    '
    Call CARDATIMPRE
    '
    ATOT% = 0
    For J& = 1 To SHOWPLA07.GRID2.Cols
       ACOLUM%(J&) = XVALO(SHOWPLA07.Label8(J&).Caption)
       ATOT% = ATOT% + ACOLUM%(J&) + 1
       VXA$(J&) = ""
    Next J&
    LII$(1) = String$(ATOT%, "-")
    LII$(2) = String$(ATOT%, "-")
    LII$(3) = String$(ATOT%, "-")
    '
    Call SETPRINTER
    Call ARMAENCA
    Call ENCAHOJA
    '
    LARGOMAX = LARGOJA
    If ORIENTA% = 1 Then
         If LARGOMAX > LARGOPAP * 56.7 Then
             LARGOMAX = LARGOPAP * 56.7
         End If
      ElseIf ORIENTA% = 2 Then
         If LARGOMAX > ANCHOPAP * 56.7 Then
             LARGOMAX = ANCHOPAP * 56.7
         End If
    End If
    '
    FIN& = XVALO(SHOWPLA07.Label7)
    HAYTOTA% = 0
    For KI& = 1 To SHOWPLA07.Label4.Count
      If SHOWPLA07.Label4(KI& - 1).Visible = True Then TOTALI% = 1
    Next KI&
    '
    For U& = 1 To FIN&
       TTXX$ = "": FINSUBTO% = 0: HUBOCA% = 0
       For J& = 1 To SHOWPLA07.GRID2.Cols
          SHOWPLA07.GRID2.Row = U&
          SHOWPLA07.GRID2.COL = J& - 1
          VLXX$ = Left$(SHOWPLA07.GRID2.TEXT + Space$(ACOLUM%(J&)), ACOLUM%(J&))
          SUMACOL#(J&) = SUMACOL#(J&) + XVALO(VLXX$)
          '
          If VLXX$ <> VXA$(J&) Then HUBOCA% = 1
          FCP$ = UCase$(Left$(FORCAMPO$(J&), 1))
          If FCP$ = "$" Then HUBOCA% = 1
          If InStr(FORCAMPO$(J&), "/R") > 0 Then HUBOCA% = 1
          VXA$(J&) = VLXX$
          If HUBOCA% = 0 Then VLXX$ = ""
          '
          If TRIM$(VLXX$) = "" Then VLXX$ = Space$(ACOLUM%(J&))
          If TTXX$ <> "" Then TTXX$ = TTXX$ + " "
          TTXX$ = TTXX$ + VLXX$
          If TRIM$(VLXX$) <> "" Then
             If SHOWPLA07.GRID2.CellFontUnderline = True Then
                FINSUBTO% = Len(TTXX$)
             End If
          End If
       Next J&
       '
       If Printer.CurrentY > LARGOMAX - (7 + 6 * TOTALI%) * ALTUREN Then
           ULTIHOJ% = 0
           Call LITRANSPO
           '
           Printer.Print MIZQ$ + Space$(OFS%) + LII$(3)
           Call EJECT
           For K2% = 1 To NCL%: VFSA$(K2%) = "": Next K2%
           Printer.NewPage
           Call ENCAHOJA
       End If
       '
       Printer.Print MIZQ$ + Space$(OFS%) + TTXX$
       If U& < FIN& Then
          If FINSUBTO% > 0 Then
             Printer.Print MIZQ$ + Space$(OFS%) + String$(FINSUBTO%, "-")
          End If
       End If
    Next U&
    '
    Call LITOTA
    Printer.Print MIZQ$ + Space$(OFS%) + LII$(3)
    Call EJECT
    Printer.EndDoc
End Sub

Sub LITRANSPO()
    '
'VER COMENTADO SHOWPLA
'    If SHOWPLA07.Visible = True Then
'       Call LITRANSPOPLA
'       Exit Sub
'    End If
    '
    If TOTALI% > 0 Then
        '
        TTXX$ = ""
        For K2% = 1 To SHOWREP07.GRID.Cols
           If SHOWREP07.Label4(K2% - 1).Visible = True Then
                ' VALOXX$ = TRIM$(FORMATNUM$(SUMACOL#(K2%), "F16.2"))
                VALOXX$ = TRIM$(FORMATNUM$(SUMACOL#(K2%), FORCAMPO$(COLUVALOR + 1)))
                VLXX$ = AJUSTD$(VALOXX$, ACOLUM%(K2%))
             Else
                VLXX$ = Space$(ACOLUM%(K2%))
           End If
           If TTXX$ <> "" Then TTXX$ = TTXX$ + " "
           TTXX$ = TTXX$ + VLXX$
        Next K2%
        '
        LEYTOT$ = "Transporte  "
        Call REPLA(TTXX$, LEYTOT$, 1, Len(LEYTOT$))
        '
        Printer.Print MIZQ$ + Space$(OFS%) + LII$(3)
        Printer.Print MIZQ$ + Space$(OFS%) + TTXX$
        '
    End If
    '
End Sub
'
Sub LITRANSPOPLA()
    '
    If TOTALI% > 0 Then
        '
        TTXX$ = ""
        For K2% = 1 To SHOWPLA07.GRID2.Cols
           If SHOWPLA07.Label4(K2% - 1).Visible = True Then
                VALOXX$ = TRIM$(FORMATNUM$(SUMACOL#(K2%), "F16.2"))
                VLXX$ = AJUSTD$(VALOXX$, ACOLUM%(K2%))
             Else
                VLXX$ = Space$(ACOLUM%(K2%))
           End If
           If TTXX$ <> "" Then TTXX$ = TTXX$ + " "
           TTXX$ = TTXX$ + VLXX$
        Next K2%
        '
        LEYTOT$ = "Transporte  "
        Call REPLA(TTXX$, LEYTOT$, 1, Len(LEYTOT$))
        '
        Printer.Print MIZQ$ + Space$(OFS%) + LII$(3)
        Printer.Print MIZQ$ + Space$(OFS%) + TTXX$
        '
    End If
    '

End Sub

Sub LITOTA()
    '
'VER COMENTADO SHOWPLA
'    If SHOWPLA07.Visible = True Then
'       Call LITOTAPLA
'       Exit Sub
'    End If
    '
    If TOTALI% > 0 Then
        '
        TTXX$ = ""
        For K2% = 1 To SHOWREP07.GRID.Cols
           If SHOWREP07.Label4(K2% - 1).Visible = True Then
                VLXX$ = AJUSTD$(SHOWREP07.Label4(K2% - 1), ACOLUM%(K2%))
             Else
                VLXX$ = Space$(ACOLUM%(K2%))
           End If
           If TTXX$ <> "" Then TTXX$ = TTXX$ + " "
           TTXX$ = TTXX$ + VLXX$
        Next K2%
        '
        LEYTOT$ = "T O T A L   G E N E R A L   "
        Call REPLA(TTXX$, LEYTOT$, 1, Len(LEYTOT$))
        '
        Printer.Print MIZQ$ + Space$(OFS%) + LII$(3)
        Printer.Print
        Printer.Print MIZQ$ + Space$(OFS%) + TTXX$
        Printer.Print
        '
    End If
    '
End Sub
'
Sub LITOTAPLA()
    '
    If TOTALI% > 0 Then
        '
        TTXX$ = ""
        For K2% = 1 To SHOWPLA07.GRID2.Cols
           If SHOWPLA07.Label4(K2% - 1).Visible = True Then
                VLXX$ = AJUSTD$(SHOWPLA07.Label4(K2% - 1), ACOLUM%(K2%))
             Else
                VLXX$ = Space$(ACOLUM%(K2%))
           End If
           If TTXX$ <> "" Then TTXX$ = TTXX$ + " "
           TTXX$ = TTXX$ + VLXX$
        Next K2%
        '
        LEYTOT$ = "T O T A L   G E N E R A L   "
        Call REPLA(TTXX$, LEYTOT$, 1, Len(LEYTOT$))
        '
        Printer.Print MIZQ$ + Space$(OFS%) + LII$(3)
        Printer.Print
        Printer.Print MIZQ$ + Space$(OFS%) + TTXX$
        Printer.Print
        '
    End If

End Sub

Sub EJECT()
    '
10  On Error GoTo 99
    While Printer.CurrentY < Printer.ScaleHeight - 5 * ALTUREN
        Printer.Print " "
        If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
    Wend
    RNGL$ = MIZQ$ + Space$(OFS%) + String$(ATOT%, "-")
    ANCHOJB = Printer.TextWidth(RNGL$)
    If ULTIHOJ% > 0 Or TOTALI% = 0 Then
        Printer.Print RNGL$
        If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
    End If
    Printer.FontSize = 6
    '
    TXT$ = "Sistema FLEXOFT s/n " + SERINUM$
    TXT$ = TXT$ + " licenciado a " + EMPRELI$
    'If UCase$(FECHALI$) <> "S/LIMITE" Then
    '    TXT$ = TXT$ + " hasta " + FECHALI$
    'End If
    TXT$ = TXT$ + " - Derechos Reservados L.11723"
    TXT$ = MIZQ$ + Space$(OFS%) + TXT$
    '
    While Printer.TextWidth(TXT$) < ANCHOJB
       TXT$ = " " + TXT$ + " "
    Wend
    Printer.Print TXT$
    If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
    Printer.FontSize = TAFUESEL
    On Error GoTo 0
    Exit Sub
    '
99  ERNU$ = TRIM$(Str$(Err.Number))
100 MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") en la Impresión de\"
    MENSA$ = MENSA$ + "este Listado.\  \"
    '
    MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
    MENSA$ = MENSA$ + "   1.- Impresora Apagada o fuera de Línea.\"
    MENSA$ = MENSA$ + "   2.- Falta Papel o Papel Atorado.\"
    MENSA$ = MENSA$ + "   3.- Problemas en Cable de Señal de Impresora.\"
    MENSA$ = MENSA$ + "   4.- Problemas de Red - No hay Acceso a Impre-\"
    MENSA$ = MENSA$ + "        sora Remota.\"
    MENSA$ = MENSA$ + "   5.- Impresora de Matriz de Puntos Configurada\"
    MENSA$ = MENSA$ + "        para 'Calidad de Imprenta'. Corrija la Con-\"
    MENSA$ = MENSA$ + "        figuración para 'Borrador'.\"
    MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
    '
    c$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(c$, 53, "Error de Impresión de Listado")
    If RESPU% = 4 Then
        Resume 10
    End If
    Call CIERRARCH("*.*")
    Call FINAL("")
    Close: End
    '
End Sub
'
Sub SETPRINTER()
    '
    Dim FUENTESEL
    Dim FUENPRUE
    Dim NADAFUEN
    Static FUFIJA$(512), CAFUFIJ%
    '
    If IDL% = 32767 Then
        For U% = 1 To 17
           CCI$(U%) = ""
        Next U%
        Exit Sub
    End If
    '
    SRY$ = String$(255, "*")
    GU$ = String$(255, "-")
    CXC$ = "": CC$ = ""
       If BLP% = 1 Then CXC$ = "¦": CC$ = "="
    '
    If PRINTERSET% = 0 Then
        '
        If Printer.CurrentX > 0 Then
            Printer.NewPage
        End If
        '
10      NOMIMPRE$ = TRIM$(Printer.DeviceName)
        If NOMIMPRE$ = "" Then
            Call MENSERR(24, "Falta Seleccionar Impresora\en la Configuración de Windows")
            Call FINAL("")
        End If
        '
        YAESTA% = 0
        CAFUFIJ% = 0
        NX% = FILEOPEN(LUCOM$ + "PRINTERS.DRV", 0, 128)
        FIN& = LOF(NX%) / 128
        Dim PPZZ As String * 128
        For IPRIN& = 1 To FIN&
            Get #NX%, IPRIN&, PPZZ$
            PPXX$ = PPZZ$
            If TRIM$(Left$(PPXX$, 48)) = Left$(NOMIMPRE$, 48) Then
                YAESTA% = 1
                If UCase$(Mid$(PPXX$, 125, 2)) = "SI" Then
                    CAFUFIJ% = CAFUFIJ% + 1
                    NOFUEN$ = UCase$(TRIM$(Mid$(PPZZ$, 65, 48)))
                    ' ESTO SE AGREGA PARA SUPRIMIR FUENTES NO USABLES
                    If InStr(NOFUEN$, "COURIER") > 0 Then GoTo 14
                    If InStr(NOFUEN$, "LUCIDA") > 0 Then GoTo 14
                    If InStr(NOFUEN$, "TERMINAL") > 0 Then GoTo 14
                    If InStr(NOFUEN$, "DRAFT") > 0 Then GoTo 14
                    If InStr(NOFUEN$, "ROMAN") > 0 Then GoTo 14
                    GoTo 19
                    '
14                  FUFIJA$(CAFUFIJ%) = PPZZ$
                End If
            End If
19      Next IPRIN&
        '
        Close #NX%
        NOMIMP$ = REPLACAR$(NOMIMPRE$, "\", "/")
        If YAESTA% < 1 Then
            OPX% = COMALTER%("Se ha Detectado un Modelo de Impresora\Desconocido para el Sistema Aplicativo\  \'" + NOMIMP$ + "'\  \\Configurar\Cancelar")
            If OPX% <> 1 Then Call FINAL("")
            Call SETUPRINTER
            GoTo 10
        End If
        '
        If CAFUFIJ% < 1 Then
            OPX% = COMALTER%("No hay Fuentes Utilizables para Impresora\  \'" + NOMIMP$ + "'\  \\Configurar\Cancelar")
            If OPX% <> 1 Then Call FINAL("")
            Call SETUPRINTER
            GoTo 10
        End If
        '
      End If
      
        On Error Resume Next
        PORIAN = Printer.Orientation
        If Option2(1).Value = True Then
            Printer.Orientation = 2
            ORIENTA% = 2
          Else
            Printer.Orientation = 1
            ORIENTA% = 1
        End If
        '
        Printer.ScaleMode = 1
        ANCHOJA = 56.7 * (ANCHOPAP - MARGENI)
        If ORIENTA% = 2 Then ANCHOJA = 56.7 * (LARGOPAP - MARGENI)
        'ANCHOMAX = Int(56.7 * CVS(Mid$(FUFIJA$(1), 49, 4)))
        'If ANCHOMAX < ANCHOJA Then ANCHOJA = ANCHOMAX
        '
        ATEXMAX = 0
        TAFUESEL = 0
        '
        TALEBASI = 24
        ATOT% = 0
'VER COMENTADO SHOWPLA
'        If SHOWPLA07.Visible = True Then
'             For j& = 1 To SHOWPLA07.GRID2.Cols
'               ATOT% = ATOT% + XVALO(SHOWPLA07.Label8(j&).Caption) + 1
'             Next j&
'           Else
             For J& = 1 To SHOWREP07.GRID.Cols
               ATOT% = ATOT% + XVALO(SHOWREP07.Label8(J&).Caption) + 1
             Next J&
'        End If
        '
100     TEPRUEBA$ = String$(ATOT%, "*")
        Printer.FontBold = False
        Printer.FontItalic = False
        For i% = 1 To CAFUFIJ%
            '
            Printer.FontName = TRIM$(Mid$(FUFIJA$(i%), 65, 48))
            Printer.FontBold = False
            Printer.FontItalic = False
            '
            FUENPRUE = TRIM$(Mid$(FUFIJA$(i%), 65, 48))
            If Printer.FontName <> FUENPRUE Then
                GoTo 109   ' modificado 20/11/01 para que busque la fuente
                           ' disponible
                Call COMUNI("Se ha Detectado un Error\en la Configuración Local\de la Impresora Predeterminada.\  \Imposible Emitir el Listado Solicitado.")
                Call CIERRARCH("*.*")
                Close
                Call FILEKILL(LUCOM$ + "PRINTERS.DRV")
                Call FINAL("")
            End If
            '
            FUENTEMIN = CVS(Mid$(FUFIJA$(i%), 113, 4))
            FUENTEMAX = CVS(Mid$(FUFIJA$(i%), 117, 4))
            '
            TALEBA = FUENTEMAX + 0.4
            While TALEBA > FUENTEMIN
               TALEBA = TALEBA - 0.4
               Printer.FontSize = TALEBA
               If Printer.FontSize > TALEBA Then GoTo 105
               '
               On Error Resume Next
               ATEXTO = Printer.TextWidth(TEPRUEBA$)
               If Err Then
                 On Error GoTo 0
                 GoTo 105
               End If
               On Error GoTo 0
               If ATEXTO <= ANCHOJA Then
                  If ATEXTO > ATEXMAX Then
                    FUENTESEL = FUENPRUE    ' Printer.FontName
                    TAFUESEL = Printer.FontSize ' TALEBASI
                    ATEXMAX = ATEXTO
                  End If
                  GoTo 109
               End If
               TALEBA = Printer.FontSize
105         Wend
        '
109     Next i%
        '
        If FUENTESEL = NADAFUEN Or TAFUESEL < 1 Then
            Call MENSERR(24, "No hay Fuentes Adecuadas\para la Impresion de Este Listado")
            Call FINAL("")
        End If
        '
        Printer.FontName = FUENTESEL
        Printer.FontSize = TAFUESEL
        Text5.TEXT = FUENTESEL
        Text5.Refresh
        Text7.TEXT = CSTRING$(MKS$(TAFUESEL), 3, 6, 1, 1)
        Text7.Refresh
       '
        If FUENTESEL = "" Or TAFUESEL < 0.4 Then
            Call MENSERR(24, "No puede Establecerse\el Formato de Impresion.\Consulte al Soporte de Sistemas.")
            Call FINAL("")
        End If
        '
110     Printer.FontName = FUENTESEL
        Printer.FontSize = TAFUESEL
        LARGOJA = Val(Printer.ScaleHeight)
        'ANCHOJA = Printer.ScaleWidth
        TEPRUEBB$ = TEPRUEBA$ ' : If LX810 = 1 Then TEPRUEBB$ = TEPRUEBA$ + "*****"
        ANCHOJB = Val(Printer.TextWidth(TEPRUEBB$))
        TALEBASI = Printer.FontSize
        ALTUREN = Printer.TextHeight(TEPRUEBA$)
        Printer.FontBold = False
        Printer.FontItalic = False
        MIZMAX = ANCHOJA - ANCHOJB
        If MARGENI > 0 Then
            MIZMAX = 56.7 * MARGENI
        End If
        MIZQ$ = ""
        While Printer.TextWidth(MIZQ$ + "*") <= MIZMAX
           MIZQ$ = MIZQ$ + " "
        Wend
        MIZQ$ = Mid$(MIZQ$, 2)
        PRINTERSET% = 1
        '
    ' End If
    '
End Sub
'
Sub CARDATIMPRE()
    '
    NOMIMPRE$ = TRIM$(Printer.DeviceName)
    On Error Resume Next
    Printer.Orientation = 1
    On Error GoTo 0
    Printer.ScaleMode = 1
    ANCHOPAP = Int(Printer.ScaleWidth / 56.7)
    LARGOPAP = Int(Printer.ScaleHeight / 56.7)
    MARGENI = 0
    MARGENS = 0
    ORIENTA% = 1
    CACOPIAS% = 1
    RENULIS% = 0
    '
    Dim PPZZ As String * 128
    NX% = FILEOPEN(LUCOM$ + "LISTSET.DRV", 0, 128)
    FIN& = LOF(NX%) / 128
    For IPRIN& = 1 To FIN&
         Get #NX%, IPRIN&, PPZZ$
         PPXX$ = PPZZ$
         If TRIM$(Left$(PPXX$, 48)) = NOMIMPRE$ Then
           If TRIM$(UCase$(Mid$(PPXX$, 57, 8))) = TRIM$(UCase$(PROPRIN$)) Then
             ANCHOPAP1 = CVS(Mid$(PPXX$, 49, 4))
             LARGOPAP1 = CVS(Mid$(PPXX$, 53, 4))
             '
             If ANCHOPAP1 > 0 Then
                If ANCHOPAP1 < ANCHOPAP Then
                   ANCHOPAP = ANCHOPAP1
                End If
             End If
             '
             If LARGOPAP1 > 0 Then
                If LARGOPAP1 < LARGOPAP Then
                   LARGOPAP = LARGOPAP1
                End If
             End If
             '
             GoTo 10
             '
           End If
         End If
    Next IPRIN&
    '
10  Close #NX%
    NX% = FILEOPEN%(LUCOM$ + "LISTABU.DRV", 0, 128)
    FIN& = LOF(NX%) / 128
    For IPRIN& = 1 To FIN&
         Get #NX%, IPRIN&, PPZZ$
         PPXX$ = PPZZ$
         If TRIM$(Left$(PPXX$, 8)) = TRIM$(PROPRIN$) Then
             MARGENI1 = CVS(Mid$(PPXX$, 9, 4))
             MARGENS1 = CVS(Mid$(PPXX$, 13, 4))
             ORIENTA1% = Asc(Mid$(PPXX$, 17, 1))
             CACOPIAS1% = Asc(Mid$(PPXX$, 18, 1))
             RENULIS% = Asc(Mid$(PPXX$, 19, 1))
             '
             If MARGENI1 > MARGENI Then
                 MARGENI = MARGENI1
             End If
             '
             If MARGENS1 > MARGENS Then
                   MARGENS = MARGENS1
             End If
             '
             If ORIENTA1% > 0 Then
                If ORIENTA1% <= 2 Then
                    ORIENTA% = ORIENTA1%
                End If
             End If
             '
             If CACOPIAS1% > 0 Then
                If CACOPIAS1% <= 32 Then
                    CACOPIAS% = CACOPIAS1%
                End If
             End If
             '
             GoTo 20
             '
         End If
    Next IPRIN&
    '
20 ' If ORIENTA% = 1 Then
'         Option2(0).Value = True        ' RETRATO
'       Else
'         Option2(1).Value = True        ' PAISAJE
'    End If
    Text4(0).TEXT = CSTRING$(MKS$(ANCHOPAP), 3, 6, 1, 1)
    Text4(1).TEXT = CSTRING$(MKS$(LARGOPAP), 3, 6, 1, 1)
    Text4(2).TEXT = CSTRING$(MKS$(MARGENS), 3, 6, 1, 1)
    Text4(3).TEXT = CSTRING$(MKS$(MARGENI), 3, 6, 1, 1)
    Text6.TEXT = CACOPIAS%
    'Check2.Value = 0
    'If RENULIS% = 1 Then
    '  Check2.Value = 1
    'End If
    '
End Sub

Sub ENCAHOJA()
    '
'VER COMENTADO SHOWPLA
'    If SHOWPLA07.Visible = True Then
'       Call ENCAHOJAPLA
'       Exit Sub
'    End If
    '
    Dim HDRX As String * 256
    Dim HDRZ$(48)
    Dim ENCATAB$(48)
    '
    HOII% = HOY(HOS$)
    '
    NUHOJ% = NUHOJ% + 1
    NH# = CDbl(NUHOJ%)
    '
    For K3% = 1 To NCL%
         VDAA%(K3%) = 0
         VDAB#(K3%) = 0
         VDA$(K3%) = ""
    Next K3%
    '
    If NUHOJ% >= PRIHOJ% Then
        Call MEMBRETE
        If TNC% < 1 Then
            Printer.CurrentY = MARGENS * 56.7 + ALTULOGO
            CURRY = Printer.CurrentY
            For RI% = 1 To RGI&
               Printer.Print , " "
               If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
            Next RI%
        End If
    End If
    '
    s1$ = SHOWREP07.Label3
    ESPAMED% = ATOT% - Len(s1$) - 43
    If ESPAMED% < 3 Then ESPAMED% = 3
    S2$ = Left$(SRY$, Len(s1$))
    '
    TEFE$ = "Fecha: " + HOS$ + " - " + Left$(Time$, 5): If SUPRIFE% = 1 Then TEFE$ = "                       "
    SUFE$ = "******": If SUPRIFE% = 1 Then SUFE$ = "      "
    HDRZ$(1) = s1$ + Space$(ESPAMED%) + "   " + TEFE$ + "    Hoja No:#####"
    HDRZ$(2) = S2$ + Space$(ESPAMED%) + "   " + SUFE$ + "                     ********"
    CAENCA% = 2
    '
    PPXX% = InStr(HDRZ$(1), "#####")
    If PPXX% > 0 Then Call REPLA(HDRZ$(1), NUMSTRI$(NH#, 5, 0, 0), PPXX%, 5)
    '
    Printer.Print MIZQ$ + Space$(OFS%) + HDRZ$(1)
    Printer.Print MIZQ$ + Space$(OFS%) + HDRZ$(2)
    '
    For RI% = 1 To SHOWREP07.Combo1.Count
      If TRIM$(SHOWREP07.Combo1(RI% - 1).TEXT) <> "" Then
        LEYEN$ = TRIM$(SHOWREP07.Label1(RI% - 1)) + ": "
        LEYEN$ = LEYEN$ + SHOWREP07.Combo1(RI% - 1).TEXT
        On Error Resume Next
          ACLA$ = TRIM$(SHOWREP07.FiltEco1(RI% - 1).Caption)
        On Error GoTo 0
        If ACLA$ <> "" Then
          LEYEN$ = LEYEN$ + " - " + ACLA$
        End If
        '
        If LEYEN$ <> "" Then
          If Left$(LEYEN$, 1) <> ":" Then
             PRI% = InStr(LEYEN$, ":")
             If PRI% = 0 Then PRI% = Len(LEYEN$) + 1
             SUBRAX$ = String$(PRI%, "*")
             While Len(SUBRAX$) < Len(s1$)
                LEYEN$ = " " + LEYEN$
                SUBRAX$ = " " + SUBRAX$
             Wend
             '
             If ARCHISAL$ = "" Then
                Printer.Print " "
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
             End If
             '
             LEYF$ = ""
             If Len(LEYEN$) > ATOT% Then
                PPXY% = InStr(UCase$(LEYEN$), "HASTA:")
                If PPXY% > 0 Then
                   LEYF$ = Mid$(LEYEN$, PPXY%)
                   LEYEN$ = Left$(LEYEN$, PPXY% - 1)
                End If
             End If
                       '
             If ARCHISAL$ = "" Then
                  Printer.Print MIZQ$ + Space$(OFS%) + Left$(LEYEN$ + Space$(ATOT%), ATOT%)
                  If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                  Printer.Print MIZQ$ + Space$(OFS%) + Left$(SUBRAX$ + Space$(ATOT%), ATOT%)
                  If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
             End If
             '
             If Len(LEYF$) > 0 Then
                If ARCHISAL$ = "" Then
                    Printer.Print MIZQ$ + Space$(OFS%) + Left$(Space$(1 + Len(SUBRAX$)) + LEYF$, ATOT%)
                    If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                End If
             End If
             '
          End If
        End If
      End If
      '
    Next RI%
    '
30  Printer.Print "  ": Printer.Print " "
    '
    Printer.Print MIZQ$ + LII$(1)
    If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
    If CAENCOL% < 1 Then
        Printer.Print MIZQ$ + ENCA$
        If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
      Else
        For KKCOL% = 1 To CAENCOL%
          Printer.Print MIZQ$ + ENCATAB$(KKCOL%)
          If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
        Next KKCOL%
    End If
    Printer.Print MIZQ$ + LII$(2)
    If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
    '
    '
    For K3% = 1 To NCL%
       SUMA#(K3%, 1) = 0
    Next K3%
    '
End Sub
'
Sub ENCAHOJAPLA()
    Dim HDRX As String * 256
    Dim HDRZ$(48)
    Dim ENCATAB$(48)
    '
    HOII% = HOY(HOS$)
    '
    NUHOJ% = NUHOJ% + 1
    NH# = CDbl(NUHOJ%)
    '
    For K3% = 1 To NCL%
         VDAA%(K3%) = 0
         VDAB#(K3%) = 0
         VDA$(K3%) = ""
    Next K3%
    '
    If NUHOJ% >= PRIHOJ% Then
        Call MEMBRETE
        If TNC% < 1 Then
            Printer.CurrentY = MARGENS * 56.7 + ALTULOGO
            CURRY = Printer.CurrentY
            For RI% = 1 To RGI&
               Printer.Print , " "
               If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
            Next RI%
        End If
    End If
    '
    s1$ = SHOWPLA07.Label3
    ESPAMED% = ATOT% - Len(s1$) - 43
    If ESPAMED% < 3 Then ESPAMED% = 3
    S2$ = Left$(SRY$, Len(s1$))
    '
    TEFE$ = "Fecha: " + HOS$ + " - " + Left$(Time$, 5): If SUPRIFE% = 1 Then TEFE$ = "                       "
    SUFE$ = "******": If SUPRIFE% = 1 Then SUFE$ = "      "
    HDRZ$(1) = s1$ + Space$(ESPAMED%) + "   " + TEFE$ + "    Hoja No:#####"
    HDRZ$(2) = S2$ + Space$(ESPAMED%) + "   " + SUFE$ + "                     ********"
    CAENCA% = 2
    '
    PPXX% = InStr(HDRZ$(1), "#####")
    If PPXX% > 0 Then Call REPLA(HDRZ$(1), NUMSTRI$(NH#, 5, 0, 0), PPXX%, 5)
    '
    Printer.Print MIZQ$ + Space$(OFS%) + HDRZ$(1)
    Printer.Print MIZQ$ + Space$(OFS%) + HDRZ$(2)
    '
    For RI% = 1 To SHOWPLA07.Combo1.Count
      If TRIM$(SHOWPLA07.Combo1(RI% - 1).TEXT) <> "" Then
        LEYEN$ = TRIM$(SHOWPLA07.Label1(RI% - 1)) + ": "
        LEYEN$ = LEYEN$ + SHOWPLA07.Combo1(RI% - 1).TEXT
'        On Error Resume Next
'          ACLA$ = TRIM$(SHOWPLA07.FiltEco1(RI% - 1).Caption)
'        On Error GoTo 0
        If ACLA$ <> "" Then
          LEYEN$ = LEYEN$ + " - " + ACLA$
        End If
        '
        If LEYEN$ <> "" Then
          If Left$(LEYEN$, 1) <> ":" Then
             PRI% = InStr(LEYEN$, ":")
             If PRI% = 0 Then PRI% = Len(LEYEN$) + 1
             SUBRAY$ = String$(PRI%, "*")
             While Len(SUBRAY$) < Len(s1$)
                LEYEN$ = " " + LEYEN$
                SUBRAY$ = " " + SUBRAY$
             Wend
             '
             If ARCHISAL$ = "" Then
                Printer.Print " "
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
             End If
             '
             LEYF$ = ""
             If Len(LEYEN$) > ATOT% Then
                PPXY% = InStr(UCase$(LEYEN$), "HASTA:")
                If PPXY% > 0 Then
                   LEYF$ = Mid$(LEYEN$, PPXY%)
                   LEYEN$ = Left$(LEYEN$, PPXY% - 1)
                End If
             End If
                       '
             If ARCHISAL$ = "" Then
                  Printer.Print MIZQ$ + Space$(OFS%) + Left$(LEYEN$ + Space$(ATOT%), ATOT%)
                  If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                  Printer.Print MIZQ$ + Space$(OFS%) + Left$(SUBRAY$ + Space$(ATOT%), ATOT%)
                  If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
             End If
             '
             If Len(LEYF$) > 0 Then
                If ARCHISAL$ = "" Then
                    Printer.Print MIZQ$ + Space$(OFS%) + Left$(Space$(1 + Len(SUBRAY$)) + LEYF$, ATOT%)
                    If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                End If
             End If
             '
          End If
        End If
      End If
      '
    Next RI%
    '
30  Printer.Print "  ": Printer.Print " "
    '
    Printer.Print MIZQ$ + LII$(1)
    If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
    If CAENCOL% < 1 Then
        Printer.Print MIZQ$ + ENCA$
        If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
      Else
        For KKCOL% = 1 To CAENCOL%
          Printer.Print MIZQ$ + ENCATAB$(KKCOL%)
          If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
        Next KKCOL%
    End If
    Printer.Print MIZQ$ + LII$(2)
    If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
    '
    '
    For K3% = 1 To NCL%
       SUMA#(K3%, 1) = 0
    Next K3%
    '
End Sub
'
Sub ENCAHOJA1(NCXX$)
    '
    NRENCA% = 0
    Dim HDRZ$(10)
    '
    If ARCHISAL$ = "" Then
        If TNC% > 0 Then
            Printer.CurrentY = 56.7 * MARGENS
            CURRY = Printer.CurrentY
            For RI% = 1 To RGI&
                 Printer.Print , " "
                 If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
            Next RI%
            If Printer.CurrentY < 12 * 56.7 Then   ' PARA DEJAR ESPACIO PARA EL LOGO
               Printer.CurrentY = 12 * 56.7
               CURRY = Printer.CurrentY
            End If
        End If
    End If
    '
    HOII% = HOY(HOS$)
    ESPAMED% = ATOT% - Len(s1$) - 43
    S2$ = Left$(SRY$, Len(s1$))
    TEFE$ = "Fecha: " + HOS$ + " - " + Left$(Time$, 5): If SUPRIFE% = 1 Then TEFE$ = "                       "
    SUFE$ = "******": If SUPRIFE% = 1 Then SUFE$ = "      "
    HDRZ$(1) = s1$ + Space$(ESPAMED%) + "   " + TEFE$ + "    Hoja No:#####"
    HDRZ$(2) = S2$ + Space$(ESPAMED%) + "   " + SUFE$ + "                     ********"
    HDRZ$(3) = Space$(10)
    'HDRZ$(4) = Space$(10)
    'HDRZ$(5) = Space$(10)
    CAENCA% = 3
    '
    NUHOK% = NUHOK% + 1
    PPXX% = InStr(HDRZ$(1), "#####")
    If PPXX% > 0 Then
        Call REPLA(HDRZ$(1), FORMATNUM$(NUHOK%, "I5"), PPXX%, 5)
    End If
    For U& = 1 To CAENCA%
       If ARCHISAL$ = "" Then
            If TNC% > 0 Then Printer.Print MIZQ$ + Space$(OFS%) + HDRZ$(U&) Else Printer.Print
            If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
          ElseIf ARCHISAL$ = "SCREEN" Then
            PREVIEW07.List4.AddItem Space$(OFS%) + HDRZ$(U&)
          Else
            Print #ASAL%, Space$(OFS%);
            Print #ASAL%, HDRZ$(U&)
       End If
       NRENCA% = NRENCA% + 1
    Next U&
    '
    CLIK% = 0
    For CLIJ% = 1 To TNC%
        TTXI$ = Mid$(NCXX$, PR%(CLNC%(CLIJ%)), LF%(CLNC%(CLIJ%)))
        If TVL%(CLNC%(CLIJ%)) = 6 Then
            'Esto es para imprimir fecha y hora en listados con destinatario
            'If SUPRIFE% <> 1 Then
            '  CLIK% = CLIK% + 1
            '  If CLIK% <= 2 Then
            '    While Len(TTXI$) < ATOT% - 28: TTXI$ = TTXI$ + " ": Wend
            '    If CLIK% = 1 Then
            '        HOII% = HOY(HO$)
            '        TTXI$ = TTXI$ + "Emitido el: " + HOS$ + " - " + Left$(Time$, 5)
            '      ElseIf CLIK% = 2 Then
            '        TTXI$ = TTXI$ + "***********"
            '    End If
            '  End If
            'End If
            '
            TTXI1$ = TRIM$(TTXI$)
            SUBRA1$ = ""
            PRI% = InStr(TTXI1$, ":")
            If PRI% > 0 Then
               SUBRA1$ = String$(PRI%, "*")
               While Len(SUBRA1$) < Len(s1$)
                 TTXI1$ = " " + TTXI1$
                 SUBRA1$ = " " + SUBRA1$
               Wend
            End If        '
            '
            If ARCHISAL$ = "" Then
                Printer.Print MIZQ$ + TTXI1$
                If SUBRA1$ <> "" Then Printer.Print MIZQ$ + SUBRA1$
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
              ElseIf ARCHISAL$ = "SCREEN" Then
                PREVIEW07.List4.AddItem TTXI1$
                If SUBRA1$ <> "" Then PREVIEW07.List4.AddItem SUBRA1$
              Else
                Print #ASAL%, TTXI1$
                If SUBRA1$ <> "" Then Print #ASAL%, SUBRA1$
            End If
            NRENCA% = NRENCA% + 1
        End If
    Next CLIJ%
    '
    '
End Sub
'
Sub MEMBRETE()
    '
    ALTULOGO = 0
    '
    RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
    If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
    If Right$(RUTARGRA$, 1) <> "\" Then
       RUTARGRA$ = RUTARGRA$ + "\"
    End If
    '
    ISOLOGO$ = TRIM$(CONTROL$("", "ISOLOGO"))
    If ISOLOGO$ <> "" Then
      ARCHIMAG$ = RUTARGRA$ + ISOLOGO$
      If EXISTE%(ARCHIMAG$) = 1 Then
        '
        FORMALTER.Image3.Picture = LoadPicture("")
        FORMALTER.Image3.Picture = LoadPicture(ARCHIMAG$)
        '
        HIMAGEN = FORMALTER.Image3.Picture.Width
        VIMAGEN = FORMALTER.Image3.Picture.Height
        '
        ph1 = 0: PH2 = 100 * 56.7
        pv1 = 0: PV2 = 12 * 56.7
        ALTULOGO = PV2
        HCUADRO = PH2 - ph1
        VCUADRO = PV2 - pv1
        '
        REDUH = HIMAGEN / HCUADRO
        REDUV = VIMAGEN / VCUADRO
        REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
        '
        altucua = VCUADRO * REDUV / REDUC
        anchcua = HCUADRO * REDUH / REDUC
        borizq = ANCHOJA - anchcua - 2 * 56.7
        borsup = 0
        '
        Printer.PaintPicture FORMALTER.Image3.Picture, ph1 + borizq, pv1 + borsup, anchcua, altucua
      End If
    End If
    '
    If EXISTE%(LUDAT$ + PROPRIN$ + ".PRL") Then
      FNA = Printer.FontName
      FNB = Printer.FontItalic
      FNC = Printer.FontSize
      FND = Printer.FontBold
      '
      MARIZQUIE = Printer.TextWidth(MIZQ$)
      '
      PRLXX$ = LOADFILE$(LUDAT$ + PROPRIN$ + ".PRL")
      For IL& = 1 To Len(PRLXX$) Step 128
        RF$ = Mid$(PRLXX$, IL&, 128)
        FUENX$ = TRIM$(Mid$(RF$, 87, 32))
          Printer.FontName = FUENX$
          Printer.FontItalic = False
        TAFUENX = CVS(Mid$(RF$, 119, 4))
        If TAFUENX > 2 Then
          Printer.FontSize = TAFUENX
        End If
          Printer.FontBold = False
        If Asc(Mid$(RF$, 123, 1)) > 32 Then
          Printer.FontBold = True
        End If
          Printer.CurrentX = CVS(Mid$(RF$, 57, 4)) * 56.7 + MARIZQUIE
          Printer.CurrentY = CVS(Mid$(RF$, 61, 4)) * 56.7 + MARSUPIE
          Printer.Print TRIM$(Mid$(RF$, 9, 48))
          If Printer.CurrentY > ALTULOGO Then
            ALTULOGO = Printer.CurrentY
          End If
      Next IL&
      '
      Printer.FontName = FNA
      Printer.FontItalic = FNB
      Printer.FontSize = FNC
      Printer.FontBold = FND
    End If
    '
End Sub

Sub ARMAENCA()
    '
    ENCA$ = ""
    '
'VER COMENTADO SHOWPLA
'    If SHOWPLA07.Visible = True Then
'          For j& = 1 To SHOWPLA07.GRID2.Cols
'             SHOWPLA07.GRID2.Row = 0
'             SHOWPLA07.GRID2.COL = j& - 1
'             ALI = SHOWPLA07.GRID2.ColAlignment(j& - 1)
'             '
'             VLXY$ = SHOWPLA07.GRID2.TEXT
'             If ALI = 6 Or j& > 2 Then
'                  VLXX$ = AJUSTD$(VLXY$, ACOLUM%(j&))
'                Else
'                  VLXX$ = AJUSTI$(VLXY$, ACOLUM%(j&))
'             End If
'             If TRIM$(VLXX$) = "" Then VLXX$ = Space$(ACOLUM%(j&))
'             If ENCA$ <> "" Then ENCA$ = ENCA$ + " "
'             ENCA$ = ENCA$ + VLXX$
'          Next j&
'       Else
          For J& = 1 To SHOWREP07.GRID.Cols
             SHOWREP07.GRID.Row = 0
             SHOWREP07.GRID.COL = J& - 1
             ALI = SHOWREP07.GRID.ColAlignment(J& - 1)
             '
             VLXY$ = SHOWREP07.GRID.TEXT
             If ALI >= 6 Then
                  VLXX$ = AJUSTD$(VLXY$, ACOLUM%(J&))
                Else
                  VLXX$ = AJUSTI$(VLXY$, ACOLUM%(J&))
             End If
             If TRIM$(VLXX$) = "" Then VLXX$ = Space$(ACOLUM%(J&))
             If ENCA$ <> "" Then ENCA$ = ENCA$ + " "
             ENCA$ = ENCA$ + VLXX$
          Next J&
'    End If
    '
End Sub
'
Sub CARGALIMPRE()
    '
    Combo1.Clear
    Dim PRX As Printer
    For Each PRX In Printers
       Combo1.AddItem PRX.DeviceName
    Next
    Combo1.TEXT = Printer.DeviceName
    IMPREDET$ = Printer.DeviceName
    '
    If IMPREDET$ = "" Then
       Call MENSERR(24, "No Hay Impresora Predeterminada.\Imposible Imprimir.")
       Call FINAL("")
    End If
End Sub

Sub SELECPRINTER(OKEY%)
    '
    If TRIM$(IMPREDET$) <> TRIM$(Combo1.TEXT) Then
        OKEY% = 0
        Printer.TrackDefault = False
        NPPXX$ = TRIM$(Combo1.TEXT)
        For Each PRX9 In Printers
          If PRX9.DeviceName = NPPXX$ Then
            On Error Resume Next
            Set Printer = PRX9
            On Error GoTo 0
            GoTo 10
          End If
        Next PRX9
        '
10      If Printer.DeviceName <> Combo1.TEXT Then
          MENSA$ = "Imposible Aplicar Impresora Seleccionada:\"
          MENSA$ = MENSA$ + "'" + TRIM$(REPLACAR$(Combo1.TEXT, "\", "/")) + "'\   \"
          MENSA$ = MENSA$ + "FLEXOFT ha Restaurado la Impresora Predeterminada:\"
          MENSA$ = MENSA$ + "'" + TRIM$(REPLACAR$(IMPREDET$, "\", "/")) + "'\  \"
          MENSA$ = MENSA$ + "Para Utilizar la Elegida, Cambie la Impresora\"
          MENSA$ = MENSA$ + "Predeterminada para Windows."
          Combo1.TEXT = IMPREDET$
          Call COMUNI(MENSA$)
          Combo1.SetFocus
          Exit Sub
        End If
    End If
    OKEY% = 1
    '
End Sub

Sub SETUPRINTER()
    '
    ' WLITAB07.Hide
    '
    NOMIMPRE$ = TRIM$(Printer.DeviceName)
    If NOMIMPRE$ = "" Then
        Call MENSERR(24, "Falta Seleccionar Impresora\en la Configuración de Windows")
        Call FINAL("")
    End If
    '
    YAESTA% = 0
    Call BLANARCH("!PRINTERS")
    NOMI$ = AJUSTI$(NOMIMPRE$, 48)
    NX% = FILEOPEN(LUCOM$ + "PRINTERS.DRV", 0, 128)
    For IPRIN& = 1 To LOF(NX%) / 128
        Get #NX%, IPRIN&, PPZZ$
        PPXX$ = PPZZ$
        If TRIM$(Left$(PPXX$, 48)) = NOMIMPRE$ Then
            Call REGAPP("!PRINTERS", PPXX$)
            YAESTA% = 1
            GoTo 10
        End If
    Next IPRIN&
    '
10  Close #NX%
    If YAESTA% = 1 Then
        PPXX$ = Left$(PPXX$, 52)
        'GoTo 50
    End If
    '
    Printer.ScaleMode = 1
    ANCHOJMM = Int(Printer.ScaleWidth / 56.7)
    Call REPLA(PPXX$, MKS$(ANCHOJMM), 49, 4)
    '
    CAFUFIJ% = 0
    For i% = 1 To Printer.FontCount
        Printer.FontName = Printer.Fonts(i% - 1)
        NOFUEN$ = TRIM$(UCase$(Printer.FontName))
        If InStr(NOFUEN$, "COURIER") > 0 Then GoTo 11
        If InStr(NOFUEN$, "LUCIDA") > 0 Then GoTo 11
        If InStr(NOFUEN$, "TERMINAL") > 0 Then GoTo 11
        If InStr(NOFUEN$, "DRAFT") > 0 Then GoTo 11
        If InStr(NOFUEN$, "ROMAN") > 0 Then GoTo 11
        GoTo 12
        '
11      a1 = 0: A2 = 0
        On Error Resume Next
        a1 = Printer.TextWidth(String$(80, "l"))
        A2 = Printer.TextWidth(String$(80, "M"))
        On Error GoTo 0
        If a1 > 0 And A2 > 0 Then
          If a1 = A2 Then
            If CAFUFIJ% < 128 Then
              CAFUFIJ% = CAFUFIJ% + 1
              FUFIJA$(CAFUFIJ%) = Printer.Fonts(i% - 1)
            End If
          End If
        End If
12  Next i%
    '
    If CAFUFIJ% < 1 Then
       Call MENSERR(24, "No se Encontraron Fuentes de Ancho Fijo\para la Impresora Seleccionada.\  \No es Posible Emitir Listado por Impresora.\  \Consulte a su Soporte de Sistemas.")
       Call FINAL("")
    End If
    '
    Call BLANARCH("!PRINTERS")
    TALEBASI = 24
    TEPRUEBA$ = String$(128, "*")
    Printer.FontBold = False
    Printer.FontItalic = False
    For i% = 1 To CAFUFIJ%
        Printer.FontName = FUFIJA$(i%)
        If Printer.FontName <> FUFIJA$(i%) Then GoTo 19
        '
        Printer.FontSize = 24
        FUENTEMAX = Printer.FontSize
        Printer.FontSize = 2
        FUENTEMIN = Printer.FontSize
        '
14      If FUENTEMIN <= FUENTEMAX Then
            XX$ = REGBLAN$("!PRINTERS")
            Call REPLA(XX$, NOMIMPRE$, 1, 48)
            Call REPLA(XX$, MKS$(ANCHOJMM), 49, 4)
            Call REPLA(XX$, FUFIJA$(i%), 65, 48)
            Call REPLA(XX$, MKS$(FUENTEMIN), 113, 4)
            Call REPLA(XX$, MKS$(FUENTEMAX), 117, 4)
            Call REPLA(XX$, "SI", 125, 2)
            Call REGAPP("!PRINTERS", XX$)
        End If
        '
19  Next i%
    Call CIERRARCH("!PRINTERS")
    '
    PPXX$ = AJUSTI$(NOMIMPRE$, 48) + MKS$(ANCHOJMM) + String$(12, 0)
    '
50  PPYY$ = OBJPLA$("DATIMPRE", PPXX$)
    If PPYY$ = "" Then
        Call CIERRARCH("PRINTERS")
        GoTo 99
    End If
    '
    ANCHOJMM = CVS(Mid$(PPYY$, 49, 8))
    If ANCHOJMM > Int(Printer.ScaleWidth / 56.7) Or ANCHOJMM < 48 Then
        PPXX$ = PPYY$
        GoTo 50
    End If
    '
    VTB% = VENTABU%("PRINSET/" + AJUSTI$(NOMIMPRE$, 48))
    If VTB% >= 0 Then
        NX% = FILEOPEN%(LUCOM$ + "PRINTERS.DRV", 0, 128)
        NOMI$ = AJUSTI$(NOMIMPRE$, 48)
        jj& = 0
        For IPRIN& = 1 To LOF(NX%) / 128
            Get #NX%, IPRIN&, PPZZ$
            If Left$(PPZZ$, 48) <> NOMI$ Then
               jj& = jj& + 1
               Put #NX%, jj&, PPZZ$
            End If
        Next IPRIN&
        For IPRIN& = 1 To ULTREG&("!PRINTERS")
            PPXX$ = REGLEIDO$("!PRINTERS", IPRIN&)
            Call REPLA(PPXX$, NOMI$, 1, 48)
            Call REPLA(PPXX$, MKS$(ANCHOJMM), 49, 4)
            If UCase$(Mid$(PPXX$, 125, 2)) <> "SI" Then
                Call REPLA(PPXX$, "  ", 125, 2)
            End If
            PPZZ$ = PPXX$
            jj& = jj& + 1
            Put #NX%, jj&, PPZZ$
        Next IPRIN&
        For IPRIN& = jj& + 1 To LOF(NX%) / 128
            PPZZ$ = String$(128, 0)
            Put #NX%, IPRIN&, PPZZ$
        Next IPRIN&
        Close #NX%
    End If
    Call CIERRARCH("!PRINTERS")
    '
99  FORMA$ = FORMALI$
    If TRIM$(FORMA$) <> "" Then
        Call WLITAB07.CARDATLIS
        Call CARDATIMPRE
      Else
        Call COMUNI("Para Efectivizar la nueva Configuración\Debe Volver a pedir el Listado.")
        Call FINAL("")
    End If
    '
End Sub
'

Private Sub Form_Load()
    Call CARDATIMPRE
    Call SETPRINTER
End Sub

Private Sub Option2_Click(Index As Integer)
    Call SETPRINTER
End Sub

Private Sub Text7_Change()
    If XVALO(Text7) < 6 Then
       If Option2(0).Value = True Then
          Option2(1).Value = True
          Call Option2_Click(1)
       End If
    End If
End Sub
