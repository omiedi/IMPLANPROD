VERSION 5.00
Begin VB.Form SETAREC04 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Secuencia de Tareas por Recorrido"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404080&
      Height          =   7000
      Left            =   9660
      ScaleHeight     =   6945
      ScaleWidth      =   1110
      TabIndex        =   12
      Top             =   0
      Width           =   1170
      Begin VB.CommandButton Command25 
         Caption         =   "Print"
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
         Picture         =   "SETAREC04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Emite e Imprime Orden de Recorrida"
         Top             =   4410
         Width           =   650
      End
      Begin VB.CommandButton Command13 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   105
         Picture         =   "SETAREC04.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Agregar Nueva Tarea"
         Top             =   1890
         Width           =   650
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
         Height          =   620
         Left            =   105
         Picture         =   "SETAREC04.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
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
         Height          =   620
         Left            =   105
         Picture         =   "SETAREC04.frx":08FE
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   975
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   13
         Top             =   5235
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   14
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "SETAREC04.frx":0C08
         Top             =   5475
         Width           =   1515
      End
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
      Height          =   330
      Left            =   315
      TabIndex        =   11
      Text            =   " "
      Top             =   345
      Width           =   1380
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
      Height          =   330
      Left            =   1680
      TabIndex        =   10
      Top             =   345
      Width           =   175
   End
   Begin VB.PictureBox Picture3 
      Height          =   5130
      Left            =   200
      ScaleHeight     =   5070
      ScaleWidth      =   9240
      TabIndex        =   0
      Top             =   800
      Width           =   9300
      Begin VB.VScrollBar VScroll1 
         Height          =   5070
         Left            =   9000
         TabIndex        =   4
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   99000
         Left            =   0
         ScaleHeight     =   99000
         ScaleWidth      =   9195
         TabIndex        =   1
         Top             =   0
         Width           =   9200
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
            Height          =   285
            Index           =   1
            Left            =   3150
            TabIndex        =   22
            Top             =   2730
            Width           =   175
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
            Height          =   285
            Index           =   0
            Left            =   3150
            TabIndex        =   19
            Top             =   315
            Width           =   175
         End
         Begin VB.TextBox Text99 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1605
            Index           =   1
            Left            =   1980
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   3120
            Width           =   6885
         End
         Begin VB.TextBox Text99 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1605
            Index           =   0
            Left            =   1980
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   720
            Width           =   6885
         End
         Begin VB.Label label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "Equipo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   840
            TabIndex        =   25
            Top             =   2800
            Width           =   1110
         End
         Begin VB.Label label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "Equipo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   840
            TabIndex        =   24
            Top             =   380
            Width           =   1110
         End
         Begin VB.Label Label4 
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
            Index           =   1
            Left            =   3465
            TabIndex        =   23
            Top             =   2730
            Width           =   5370
         End
         Begin VB.Label Label3 
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
            Height          =   285
            Index           =   1
            Left            =   1995
            TabIndex        =   21
            Top             =   2730
            Width           =   1170
         End
         Begin VB.Label Label4 
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
            Index           =   0
            Left            =   3465
            TabIndex        =   20
            Top             =   315
            Width           =   5370
         End
         Begin VB.Label Label3 
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
            Height          =   285
            Index           =   0
            Left            =   1995
            TabIndex        =   18
            Top             =   315
            Width           =   1170
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Tarea Nro. 2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   105
            TabIndex        =   6
            Top             =   3150
            Width           =   1800
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Tarea Nro. 1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   105
            TabIndex        =   5
            Top             =   735
            Width           =   1800
         End
      End
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2310
      TabIndex        =   9
      Top             =   345
      Width           =   7155
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
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
      Index           =   39
      Left            =   0
      TabIndex        =   8
      Top             =   105
      Width           =   795
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
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
      Index           =   38
      Left            =   1680
      TabIndex        =   7
      Top             =   105
      Width           =   1380
   End
End
Attribute VB_Name = "SETAREC04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim INTERVER
Dim MODIFIC%

Private Sub Command1_Click()
   If MODIFIC% > 0 Then
     If COMALTER%("¿ Descarta los Cambios ?\\No, Conservarlos\Si, Descartar") <> 2 Then
        Exit Sub
     End If
   End If
   MODIFIC% = 0
   Unload Me
End Sub

Private Sub Command13_Click()
  '
  Command13.Enabled = False
  If TRIM$(Text4.TEXT) <> "" And ECOARCH$("TARECORR", Text4) <> "" Then
    Call CARGABLOTAR      'AGREGA COMBO DE TAREA
    On Error Resume Next
      'SE POSICIONA EN LA ULTIMA TAREA
       VScroll1.Value = VScroll1.Max
       NUPIC% = Label3.Count
       Command4(NUPIC%).SetFocus
    On Error GoTo 0
  End If
  Command13.Enabled = True
  '
End Sub

Private Sub Command2_Click()
    '
    If MODIFIC% < 1 Then Exit Sub
    '
    Command2.Enabled = False
    Screen.MousePointer = 11
    CODISEC$ = Left$(UCase$(TRIM$(Text4)), 8)
    DESEC$ = Label2
    If CODISEC$ <> "" Then
       '
       Call ABREOMAN
       ' borra las tareas anteriores
       SQLX$ = "DELETE FROM SETARECO WHERE Cod_Reco = '" & CODISEC$ & "'"
       Omantin.Execute SQLX$
       '
       ' agrega las nuevas tareas
       NUPIC% = Label3.Count
       '
       JI% = 0
       For KKI% = 1 To NUPIC%
          MAQUI$ = Left$(Label3(KKI% - 1), 8)
          DELAR$ = Text99(KKI% - 1)
          If TRIM$(DECOR$) <> "" Or TRIM$(DELAR$) <> "" Then
             With SeTaReco
                .AddNew
                JI% = JI% + 1
                !Cod_Reco = CODISEC$
                !Des_Reco = DESEC$
                !Norit_Tar = JI%
                !MaEquip = MAQUI$
                !Des_Larga = DELAR$
                .Update
             End With
          End If
       Next KKI%
       '
    End If
    Screen.MousePointer = 1
    MODIFIC% = 0
    Command2.Enabled = True
End Sub

Sub Command25_Click()

    If TRIM$(Text4.TEXT) = "" Then Exit Sub
    Call EMIORECO
    Text4.TEXT = ""
End Sub

Private Sub Command3_Click()
   Call CommSel("TARECORR", Command3, Text4)
End Sub

Private Sub Command4_Click(Index As Integer)
   Call CommSel("INDIEQUI", Command4(Index), Label3(Index))
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If MODIFIC% > 0 Then
     If COMALTER%("¿ Descarta los Cambios ?\\No, Conservarlos\Si, Descartar") <> 2 Then
        Cancel = 1
        Exit Sub
     End If
   End If
End Sub

Private Sub label3_Change(Index As Integer)
    Label4(Index) = ECOARCH$("INDIEQUI", Label3(Index))
    MODIFIC% = 1
End Sub

Private Sub Label3_DblClick(Index As Integer)
   Call Command4_Click(Index)
End Sub

Private Sub Text99_Change(Index As Integer)
   MODIFIC% = 1
End Sub

Private Sub VScroll1_Change()
   On Error Resume Next 'POR ERROR '6' DESBORDAMIENTO
     Picture4.Top = (-10) * VScroll1.Value
   On Error GoTo 0
End Sub

Private Sub VScroll1_Scroll()
  On Error Resume Next 'POR ERROR '6' DESBORDAMIENTO
   Picture4.Top = (-10) * VScroll1.Value
  On Error GoTo 0
End Sub

Private Sub Form_Load()
   '
   Static TOPEINI, IZQUINI, ANCHINI, ALTUINI
   '
   If INTERVER < 1 Then
     TOPEINI = Label3(0).Top
     IZQUINI = Label3(0).Left
     ANCHINI = Label3(0).Width
     ALTUINI = Label3(0).Height
     INTERVER = Label3(1).Top - Label3(0).Top
   End If
   '
   Picture4.Height = Picture3.Height
   VScroll1.Min = 0
   VScroll1.Max = (Picture4.Height - Picture3.Height) / 10 + 1
   VScroll1.Value = 0
   '
End Sub

Private Sub Text4_Change()
   Label2 = ECOARCH$("TARECORR", Text4)
   Call BLANFORMU
   Call CARGASECU
End Sub

Private Sub Text4_DblClick()
   Call Command3_Click
End Sub

Sub CARGABLOTAR() 'CARGA BLOQUE DE TAREA
  '
  NUPIC% = Label3.Count
  '
  Load Label1(NUPIC%)
  Label1(NUPIC%).Top = Label1(NUPIC% - 1).Top + INTERVER
  Label1(NUPIC%).Visible = True
  Label1(NUPIC%).Caption = "Tarea Nro. " & (NUPIC% + 1)
  '
  Load Label4(NUPIC%)
  Label4(NUPIC%).Top = Label4(NUPIC% - 1).Top + INTERVER
  Label4(NUPIC%).Visible = True
  Label4(NUPIC%).Caption = ""
  '
  Load Label3(NUPIC%)
  Label3(NUPIC%).Top = Label3(NUPIC% - 1).Top + INTERVER
  Label3(NUPIC%).Visible = True
  Label3(NUPIC%).Caption = ""
  '
  Load label5(NUPIC%)
  label5(NUPIC%).Top = label5(NUPIC% - 1).Top + INTERVER
  label5(NUPIC%).Visible = True
  label5(NUPIC%).Caption = "Equipo:"
  '
  Load Command4(NUPIC%)
  Command4(NUPIC%).Top = Command4(NUPIC% - 1).Top + INTERVER
  Command4(NUPIC%).Visible = True
  '
  Load Text99(NUPIC%)
  Text99(NUPIC%).Top = Text99(NUPIC% - 1).Top + INTERVER
  Text99(NUPIC%).Visible = True
  Text99(NUPIC%).TEXT = ""
  '
  If Text99(NUPIC%).Top + Text99(NUPIC%).Height > Picture4.Height - 500 Then
     Picture4.Height = Text99(NUPIC%).Top + Text99(NUPIC%).Height + 500
  End If
  '
  VScroll1.Max = (Picture4.Height - Picture3.Height) / 10 + 1
  '
End Sub

Sub BLANFORMU()
   '
   NUPIC% = Label3.Count
   For KKI% = 2 To NUPIC% - 1
      Unload Label3(KKI%)
      Unload Text99(KKI%)
      Unload Label1(KKI%)
      Unload Label4(KKI%)
      Unload label5(KKI%)
      Unload Command4(KKI%)
   Next KKI%
   '
   For KKI% = 0 To 1
      Label3(KKI%) = ""
      Text99(KKI%) = ""
   Next KKI%
   '
   Picture4.Height = Picture3.Height
   VScroll1.Min = 0
   VScroll1.Max = (Picture4.Height - Picture3.Height) / 10 + 1
   VScroll1.Value = 0
   MODIFIC% = 0
   '
End Sub

Sub CARGASECU()
    Screen.MousePointer = 11
    CODISEC$ = Left$(UCase$(TRIM$(Text4)), 8)
    '
    If CODISEC$ <> "" Then
       Call ABREOMAN
       SQLX$ = "SELECT * FROM SETARECO "
       SQLX$ = SQLX$ + "WHERE Cod_Reco = '" & CODISEC$ & "' "
       SQLX$ = SQLX$ + "ORDER BY Norit_Tar ASC "
       Set TARESEL = Omantin.OpenRecordset(SQLX$, dbOpenSnapshot)
       '
       With TARESEL
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
             On Error GoTo 0
             ITAR% = -1
             Do While Not .EOF
                ITAR% = ITAR% + 1
                If ITAR% > Label3.Count - 1 Then
                   Call CARGABLOTAR
                End If
                Label3(ITAR%) = SAFETEXT(!MaEquip)
                Text99(ITAR%) = !Des_Larga
             .MoveNext
             Loop
          End If
          On Error GoTo 0
       End With
       MODIFIC% = 0
    TARESEL.Close
    End If
    
    MODIFIC% = 0
    Screen.MousePointer = 1
    '
End Sub

Sub EMIORECO()
    '
    Call ABREOMAN
    SQLX$ = "SELECT MAX(NumORec) AS UNUM FROM ORECORRI "
    UNUI& = 0
    On Error Resume Next
    Set PRONU = Omantin.OpenRecordset(SQLX$, dbOpenSnapshot)
    If Not Err Then
       PRONU.MoveFirst
       UNUI& = PRONU!UNUM
    End If
    On Error GoTo 0
    NUOREC& = UNUI& + 1
    VUELTA% = 0
    '
10  FORIPRE$ = TRIM$(CONTROL$("", "FORMOREC"))
    If FORIPRE$ = "" Then FORIPRE$ = "OREC-MNT"  ' Impresion de la  Orden de Recorrida
    '
    TIPODOC$ = "Orden de Recorrida"
    NUMEDOC$ = TRIM$(Str(NUOREC&))
    DESTIDOC% = 0
    
    HOII% = HOY(HOS$)
    CODPARAM$(1) = "FECH-EMI": DOCPARAM$(1) = HOS$
    CODPARAM$(2) = "DES-OPER": DOCPARAM$(2) = Label2.Caption
    CAPARDOC% = 2
       
    CODTABU1$(1) = "ORD-ITEM"
    CODTABU1$(2) = "EQUIPO1"
    CODTABU1$(3) = "REF-MAT1"
    CACOLTAB1% = 3
    
    FIN& = Label3.Count
    CAITAB1% = 0
    For KKII& = 1 To FIN&
      If TRIM$(Label3(KKII& - 1)) <> "" Or TRIM$(Text99(KKII& - 1)) <> "" Then
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = TRIM$(XVALO(KKII&))
        TETABU1$(2, CAITAB1%) = Label3(KKII& - 1)
        TETABU1$(3, CAITAB1%) = Label4(KKII& - 1)
    
        CAITAB1% = CAITAB1% + 1
    
        FORTELIB$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "REF-MAT1/FORMATO")))
        ATELIB% = Val(Mid$(FORTELIB$, 2))
        Call FRATEXTO(Text99(KKII& - 1), ATELIB%)
        TETABU1$(3, CAITAB1%) = RETEX$(1)
        For KKU1% = 2 To CARETEX%
            CAITAB1% = CAITAB1% + 1
            TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
        Next KKU1%
        CAITAB1% = CAITAB1% + 2
      End If
    Next KKII&
    
    If FORIPRE$ <> "" Then
      Call PRINTDOC("@" + FORIPRE$) ' ORDEN DE RECORRIDA
    End If
    '
    If VUELTA% = 0 Then
        ' GRABA LA ORDEN DE RECORRIDA
        SQLX$ = "SELECT * FROM ORECORRI WHERE NumOrec = " & NUOREC&
        Set ORRE = Omantin.OpenRecordset(SQLX$, dbOpenDynaset)
        With ORRE
           .AddNew
           !NumORec = NUOREC
           !DesORec = SETAREC04.Label2
           !FechEmis = CVDAT(HOY(HOS$))
           !Fechcierr = CVDAT(FECHANUM("01/01/80"))
              TXOB$ = ENTRAOBS.Text28   ' ""
           !ObseEmis = TXOB$
           !ObseCierr = ""
           .Update
        End With
        '
        ' agrega al indice de ordenes de recorrida
        Call COPYSTRU("INDIOMAN", "INDIOREC")
        TTXX$ = FECHATEX$(HOY(HO$)) + " - " + SETAREC04.Label2
        Call FILTERPUTRECORD("INDIOREC", 1, MKS$(NUOREC&), MKS$(NUOREC&) + TTXX$)
        Call CIERRARCH("INDIOMAN")
    End If
    VUELTA% = VUELTA% + 1
    '
    If COMALTER%("Se ha Emitido por Impresora\La Orden de Recorrida Nº: " & NUOREC& & "\\Repetir Impresion\Continuar") = 1 Then GoTo 10
    '
End Sub

