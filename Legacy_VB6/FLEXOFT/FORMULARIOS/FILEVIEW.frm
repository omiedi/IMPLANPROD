VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FILEVIEW 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EEEEDD&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11010
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   Begin VB.FileListBox File1 
      Height          =   1665
      Left            =   240
      TabIndex        =   18
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Frame FRAMEBOT2 
      BackColor       =   &H00EEEEDD&
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   7665
      TabIndex        =   11
      Top             =   2730
      Width           =   750
      Begin VB.CommandButton BOTERASE 
         Caption         =   "Del."
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
         Left            =   105
         TabIndex        =   14
         ToolTipText     =   "Elimina Registro Activo"
         Top             =   630
         Width           =   540
      End
      Begin VB.CommandButton BOTBUSCA 
         Caption         =   "Find"
         Enabled         =   0   'False
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
         Left            =   85
         TabIndex        =   13
         ToolTipText     =   "Buscar Texto por Campo"
         Top             =   945
         Width           =   540
      End
      Begin VB.CommandButton BOTINSE 
         Caption         =   "Ins."
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
         Left            =   85
         TabIndex        =   12
         ToolTipText     =   "Intercala Registro Antes del Activo"
         Top             =   315
         Width           =   540
      End
   End
   Begin VB.TextBox Text2 
      Height          =   435
      Left            =   840
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   5460
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   105
      Top             =   4200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controles de Prueba no Visibles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2070
      Left            =   5460
      TabIndex        =   8
      Top             =   4515
      Visible         =   0   'False
      Width           =   4320
      Begin VB.ListBox LIPRUEBA 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         ItemData        =   "FILEVIEW.frx":0000
         Left            =   210
         List            =   "FILEVIEW.frx":0007
         TabIndex        =   9
         Top             =   315
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.Label PROFILE 
         Caption         =   "PROFILE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   17
         Top             =   1560
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label TEPRUTEX 
         AutoSize        =   -1  'True
         Caption         =   "TEPRUTEX"
         Height          =   225
         Left            =   2640
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
         Caption         =   "TEPRUEBA"
         Height          =   225
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label TEPRUANCH 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "TEPRUANCH"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1320
         TabIndex        =   10
         Top             =   1080
         Visible         =   0   'False
         Width           =   1080
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   9660
      ScaleHeight     =   75
      ScaleWidth      =   675
      TabIndex        =   6
      Top             =   105
      Width           =   735
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   3255
      TabIndex        =   2
      Top             =   4830
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   4110
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   7250
      _Version        =   327680
      Rows            =   16
      Cols            =   32
      FixedCols       =   0
      BackColorFixed  =   8433840
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      FocusRect       =   2
      AllowUserResizing=   1
   End
   Begin VB.Frame FRAMEBOT1 
      BackColor       =   &H00EEEEDD&
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1520
      Left            =   7665
      TabIndex        =   0
      Top             =   315
      Width           =   750
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
         Height          =   525
         Left            =   85
         Picture         =   "FILEVIEW.frx":0021
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cancela los Datos Ingresados"
         Top             =   315
         Width           =   540
      End
      Begin VB.CommandButton BOTREC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   85
         Picture         =   "FILEVIEW.frx":032B
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Graba - Guardar Modificaciones"
         Top             =   870
         Width           =   540
      End
   End
   Begin VB.Label ANAME 
      Height          =   375
      Left            =   5520
      TabIndex        =   20
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label LUDATU 
      Caption         =   "F:\CLIENTES\DOSIVAC\DATOS\"
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   4200
      Width           =   4455
   End
End
Attribute VB_Name = "FILEVIEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim YACARDA%
Dim GRABACT%      ' GRABABLE=1 - NO GRABA=0
Dim BIZW&, BSUP&  ' POSICION DEL FORMULARIO
Dim LEFTPOS%      ' espacios libres a izquierda del formulario
Dim PROFIL$
Dim ARCHITRA$, ARCHILOC$
Dim ATRIBU$
Dim YYY As String * 128
Dim CTTT%
Dim CL%(32), LC%(128)
Dim NUCOLUM%
Dim PRICOLED%           ' primera columna editable
Dim ULTCOLED%           ' ultima columna editable
Dim MAIZQ, MAINT, MADER
Dim BLABUFCA As String
Dim UROW, UCOL
Dim COLTEBO%, ROWTEBO%  ' COLUMNA Y RENGLON CAJA TEXTO
Dim VISIROWS&           ' RENGLONES VISIBLES
Dim INSEROW&
Dim RECIBECLICK%
Dim MODITEX%
Dim TEXTOINI$
Dim LIMPIAGR%

Dim ASTRUC$
Dim TVL%(128)
Dim LDL%(128)
Dim PR%(128)
Dim LF%(128)
Dim TT%(128)
Dim RTT%(128)
Dim ACL%(128)
Dim LK%(128)
Dim PINSE%(128)
Dim ECL$(128)
Dim MG%(128)
Dim PE%(128)
Dim JT%(128)
Dim AE$(128)
Dim CE%(128)
Dim CL0%(128)
Dim CFF%(128)
Dim LCF%(128)
Dim RGFL&(128)
Dim CLF%(128)
Dim ECLF$(128)
Dim LKF%(128)
Dim LEF%(128)
Dim CEF%(128)
Dim AEF$(128)


Private Sub ANAME_Change()
   ASTRUC$ = LUDATU + ANAME + ".RFS"
   Call CARDATVEN
   Call DIBUVEN
   SMP = Screen.MousePointer
   Screen.MousePointer = 1
End Sub

Private Sub BOTBUSCA_Click()
   ' Call BULISVEN
End Sub

Private Sub BOTERASE_Click()
  If GRID.Row <= UREPOS& Then
    Screen.MousePointer = 11
    GRID.RemoveItem GRID.Row
    If GRID.Rows < VISIROWS& Then GRID.Rows = VISIROWS&
    UREAR& = ULTREG&(ARCHILOC$)
    For U& = GRID.Row + 1 To UREAR&
      XX$ = REGLEIDO$(ARCHILOC$, U&)
      Call GRAREG(ARCHILOC$, XX$, U& - 1)
    Next U&
    If GRID.Row <= UREAR& Then
      Call SETULTREG(ARCHILOC$, UREAR& - 1)
    End If
    If GRID.Row > UREAR& Then
      GRID.Col = 0
      Call SETNEXTCOL
    End If
    Call MUETEBOX
    Screen.MousePointer = 1
 End If
    MODIFI% = 1
End Sub

Private Sub BOTINSE_Click()
  '
  If APROENTRA% <> 1 Then Exit Sub
  If GRID.Row < UREPOS& Then
    '
    Screen.MousePointer = 11
    UREAR& = ULTREG&(ARCHILOC$)
    For U& = UREAR& To GRID.Row Step -1
      XX$ = REGLEIDO$(ARCHILOC$, U&)
      Call GRAREG(ARCHILOC$, XX$, U& + 1)
    Next U&
    '
    XX$ = REGBLAN$(ARCHILOC$)
    Call GRAREG(ARCHILOC$, XX$, GRID.Row)
    '
    GRID.AddItem BLABUFCA$, GRID.Row
    GRID.Col = 0
    INSEROW& = GRID.Row
    Call SETNEXTCOL
    Call MUETEBOX
    MODIFI% = 1
    Screen.MousePointer = 1
 End If
'
End Sub

Private Sub BOTREC_Click()
    '
    If GRABACT% > 0 Then
      If APROENTRA% <> 1 Then Exit Sub
    End If
    YACARDA% = 0
    MODIFI% = 0
    VTB% = 0
    Hide
    Call LIMPIAGRID
    '
    Call FRESHALL
    Call CIERRARCH("*.*")
    '
End Sub

Private Sub BOTEXIT_Click()
    '
    YACARDA% = 0
    'If MODIFI% > 0 Then
    '    RESPU = MsgBox("Abandona sin Guardar los Cambios ?", 4371, "Operacion Critica")
    '    If RESPU <> 6 Then
    '        BOTREC.SetFocus
    '        Exit Sub
    '    End If
    '    MODIFI% = 0
    'End If
    '
    VTB% = (-1)
    'Call CIERRARCH("*.*")
    Unload Me
    '
End Sub
'

Private Sub File1_DblClick()
   FINAM$ = File1.filename
   PPXX% = InStr(FINAM$, ".")
   ANAME = Left$(FINAM$, PPXX% - 1)
End Sub

Private Sub Form_DblClick()
   Hide
End Sub

Private Sub Form_Load()
   '
   'Call TRANSLABELS(Name)
   Show
   Refresh
   LUDAT$ = LUDATU.Caption
   File1.Path = LUDAT$
   File1.Pattern = "*.DAT"
   '
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Shift = 1 Then
     If Button = 2 Then
       MsgBox "Controlador: " + Trim$(PROFIL$)
     End If
   End If
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
    YACARDA% = 0
    If MODIFI% > 0 Then
        RESPU = MsgBox("Abandona sin Guardar los Cambios ?", 4371, "Operacion Critica")
        If RESPU <> 6 Then
            Cancel = 1
            BOTREC.SetFocus
            Exit Sub
        End If
        MODIFI% = 0
    End If
    '
    VTB% = (-1)
    On Error Resume Next
    Hide
    On Error GoTo 0
    Call LIMPIAGRID
    Call FRESHALL
    Call CIERRARCH("*.*")
    '
End Sub

Private Sub GRID_Click()
   '
   ' linea agregada el 10/12/02 para corregir que al cambiar
   ' de renglon arrastraba el valor del texto anterior de
   ' la caja azul de texto
   'If GRID.Row >= UREPOS& Then Text1 = ""
   '
   GRID.RowSel = GRID.Row
   GRID.ColSel = GRID.Col
   If GRID.Row < 1 Then GRID.Row = 1
   If GRID.Row > UREPOS& Then
      GRID.Row = UREPOS&
   End If
   If GRID.Row > ULTREG&(ARCHILOC$) Then
      GRID.Col = 0
   End If
   If PE%(GRID.Col + 1) <> 1 Then
     If GRID.Row <> INSEROW& Then
       If PRICOLED% < 1 Then
         ROWAC& = GRID.Row
         GRID.Row = UREPOS&
         If GRID.Row <= ULTREG&(ARCHILOC$) Then
           GRID.Row = ROWAC&
           Exit Sub
         End If
       End If
       Call SETNEXTCOL
     End If
   End If
   If RECIBECLICK% = 0 Then Call MUETEBOX
   '
End Sub

Private Sub GRID_EnterCell()
   If YACARDA% < 1 Then Exit Sub
End Sub


Private Sub GRID_LeaveCell()
   '
   If LIMPIAGR% > 0 Then Exit Sub
   If YACARDA% < 1 Then Exit Sub
   If RECIBECLICK% < 0 Then Exit Sub
   If ACL%(GRID.Col + 1) < 1 Then Exit Sub
   AAA = UREPOS&
   If GRID.Row < UREPOS& Then
     If PE%(GRID.Col + 1) <> 1 Then Exit Sub
   End If
   If APROENTRA% < 1 Then
     Exit Sub
   End If
   '
End Sub

Private Sub GRID_RowColChange()
   'Text1 = ""
   ' esto es a partir del 10/12/2003 por el error señalado
   ' mas abajo
   If GRID.Row > UREPOS& Then
     Exit Sub
   End If
   '
   ' lo que sigue estaba hasta el 10/12/2003 y producia un
   ' error 28 al hacer click sobre una columna eco (p.ej
   ' descripción de un item) y varios renglones más abajo
   ' del último.
   '
   'If LIMPIAGR% <> 1 Then
   '  If GRID.Row > UREPOS& Then
   '    GRID.Row = UREPOS&
   '    If PE%(GRID.Col + 1) <> 1 Then Call SETNEXTCOL
   '  End If
   'End If
End Sub

Private Sub GRID_Scroll()
   Call MUETEBOX
End Sub

Private Sub PROFILE_Change()
   '
   TIFORMED% = 1
   PX% = InStr(PROFILE, "/")
   If PX% < 1 Then PX% = 1 + Len(PROFILE)
   '
   PROFIL$ = Left$(PROFILE, PX% - 1) + ".PRF"
   ATRIBU$ = Mid$(PROFILE, PX%)
   '
   Call CARDATVEN
   Call DIBUVEN
   SMP = Screen.MousePointer
   Screen.MousePointer = 1

End Sub

Sub CARDATVEN()
     '
85 FINU% = FILEOPEN%(ASTRUC$, 0, 128)
   Dim XYZ As String * 128
   '
   AAA = LOF(FINU%)
   If LOF(FINU%) <= 128 Then
     XYZ$ = String$(128, 0)
     Put #FINU%, 1, XYZ$
     Close #FINU%: FINU% = 0
     Kill ASTRUC$
     MsgBox "Falta Estructura de Archivo"
     Stop
     '
'     If vuelta% = 1 Then
'        ARCHEX$ = FISTRUNAME$(FINDEX%)
'86      PPXX% = InStr(ARCHEX$, "\")
'        If PPXX% > 0 Then ARCHEX$ = Mid$(ARCHEX$, PPXX% + 1): GoTo 86
'        PPXX% = InStr(UCase$(ARCHEX$), ".RFS")
'        If PPXX% > 0 Then ARCHEX$ = Left$(ARCHEX$, PPXX% - 1)
'        Call EXTRACT(ARCHEX$)
'        vuelta% = 2
'        GoTo 85
'     End If
'     '
'     ARX1$ = Trim$(REPLACAR$(FIFULLNAME$(FINDEX%), "\", "/"))
'     Call COMUNI("Proceso Interrumpido en Rutina 'FILEMGR'.\  \Falta Controlador de Estructura del Archivo\'" + ARX1$ + "'.")
'     Close: Reset: End
   End If
   '
   Get #FINU%, 1, XYZ$
   LOR% = Asc(Mid$(XYZ$, 49, 1))
   LOR1% = Asc(Mid$(XYZ$, 48, 1)): If LOR1% > 16 Then LOR1% = 0
   LOR% = 256 * LOR1% + LOR%
   If LOR% < 1 Then
     ARX1$ = Trim$(REPLACAR$(ASTRUC$, "\", "/"))
     MsgBox "Proceso Interrumpido en Rutina 'FILEMGR'.\  \Error en Longitud de Registro del Archivo\'" + ARX1$ + "'."
     Close: Reset: End
   End If
   '
   TITU$ = Left$(XYZ$, 47)
   'RBL$ = String$(LOR%, 0)
   'IXX$ = "": VXX$ = "": PXX$ = "": LXX$ = "": DXX$ = ""
   '
   'ARX$ = FISHORNAME$(FINDEX%)
   PSSA% = 0
   For I% = 1 To 128
     If I% + 1 <= LOF(FINU%) / 128 Then
       Get #FINU%, I% + 1, XYZ$
       If XYZ$ <> String$(128, 0) Then
         CAFI% = I%: NUCOLUM% = CAFI%
         TVL%(CAFI%) = Asc(Mid$(XYZ$, 33)) ': VXX$ = VXX$ + Chr$(TV%)
         'DC% = Asc(Mid$(XYZ$, 34)): DXX$ = DXX$ + Chr$(DC%)
         'PS% = Asc(Mid$(XYZ$, 35))
         '  While PS% < PSSA%: PS% = PS% + 256: Wend
         '  PXX$ = PXX$ + MKI$(PS%)
         '  PSSA% = PS%
         LC%(CAFI%) = Asc(Mid$(XYZ$, 36)) ': LXX$ = LXX$ + Chr$(LG%)
         ACL%(CAFI%) = LC%(CAFI%)
         If TVL(CAFI%) <= 2 Then
              ACL%(CAFI%) = 5
            ElseIf TVL%(CAFI%) <= 4 Then
              ACL%(CAFI%) = 12
            ElseIf TVL%(CAFI%) = 5 Then
              ACL%(CAFI%) = 8
            ElseIf TVL%(CAFI%) = 7 Then
              ACL%(CAFI%) = 15
         End If
         '
         ECL$(NUCOLUM%) = Trim$(Left$(XYZ$, 32)) 'REPLACEN$(Trim$(Mid$(YYY$, 37, 32)))
         ECL$(NUCOLUM%) = Left$(ECL$(NUCOLUM%), ACL%(NUCOLUM%))
         TEPRUEBA.Caption = String$(1 + ACL%(NUCOLUM%), "0")
         GRID.ColWidth(NUCOLUM% - 1) = TEPRUEBA.Width
         GRID.Col = NUCOLUM% - 1
         GRID.Row = 0
         GRID.RowHeight(0) = 500
         GRID.CellFontSize = GRID.FontSize
          If Left$(ECL$(NUCOLUM%), 1) = "<" Then
              ECL$(NUCOLUM%) = Mid$(ECL$(NUCOLUM%), 2)
              GRID.CellAlignment = 1
            ElseIf Left$(ECL$(NUCOLUM%), 1) = ">" Then
              ECL$(NUCOLUM%) = Mid$(ECL$(NUCOLUM%), 2)
              GRID.CellAlignment = 7
            Else
              GRID.CellAlignment = 4
          End If
         GRID.Clip = Trim$(ECL$(NUCOLUM%))
         'XC% = Asc(Mid$(XYZ$, 37)): If XC% <> 1 Then XC% = 0
         'If Left$(ARX$, 1) <> "!" Then
         '  IXX$ = IXX$ + Chr$(XC%)
         'End If
         'If TV% = 6 Then
         '  If PS% > 0 Then
         '    If LG% > 0 Then
         '      If PS% + LG% <= LOR% + 1 Then
         '        Call REPLA(RBL$, Space$(LG%), PS%, LG%)
         '      End If
         '    End If
         '  End If
         'End If
       End If
     End If
   Next I%
   '
   Close #FINU%: FINU% = 0
   Call DIBUVEN
Exit Sub
   '
'   BLANREG$(FINDEX%) = RBL$
'   LOREGIS%(FINDEX%) = LOR%
'   RELEIDO$(FINDEX%) = String$(LOR%, 0)
'   FILETIT$(FINDEX%) = TITU$
'   TVASTRI$(FINDEX%) = VXX$
'   DECSTRI$(FINDEX%) = DXX$
'   POSSTRI$(FINDEX%) = PXX$
'   LONSTRI$(FINDEX%) = LXX$
'   IDXSTRI$(FINDEX%) = IXX$
   '
5800 GRABACT% = 0
     BSUW& = 0: BIZW& = 0
     '
     If EXISTE%(PROFIL$) < 1 Then
       Call EXTRAPROF
     End If
     '
     GRID.Clear
     GRID.Cols = 64
     '
10   N15% = FILEOPEN%(PROFIL$, 10, 128)
     Get #N15%, 1, YYY$
     ARCHITRA$ = Trim$(Mid$(YYY$, 57, 8))
     ARCHILOC$ = "!" + ARCHITRA$
     PROCE$ = Trim$(Mid$(YYY$, 65, 8))
     LOR1% = Asc(Mid$(YYY$, 106, 1)): If LOR1% > 16 Then LOR1% = 0
     LOR% = 256 * LOR1% + Asc(Mid$(YYY$, 107))
     CAMAX& = CVI(Mid$(YYY$, 108))
     PEOF% = 1 + Asc(Mid$(YYY$, 110))
     EOFS$ = Mid$(YYY$, 111, 16)
5803 Caption = Left$(YYY$, 48)
     SALI$ = Mid$(YYY$, 49, 8)
     COD$ = Mid$(YYY$, 81)
     RGI& = Asc(Mid$(COD$, 1))
     RGF& = Asc(Mid$(COD$, 2))
     LEFTPOS% = Asc(Mid$(COD$, 3))
5804 BLP% = Asc(Mid$(COD$, 4))
     VIV% = Asc(Mid$(COD$, 5))
     APPI% = Asc(Mid$(COD$, 6))
     IDL% = Asc(Mid$(COD$, 7))
        BOTINSE.Enabled = True
        BOTERASE.Enabled = True
        If IDL% < 1 Then
            BOTINSE.Enabled = False
            BOTERASE.Enabled = False
        End If
     HDR& = Asc(Mid$(COD$, 8))
     RFT% = Asc(Mid$(COD$, 9))
     If IDL% > 0 Or APPI% > 0 Then GRABACT% = 1
     PRX% = Asc(Mid$(COD$, 10))
     NUCOLUM% = 0
     PRICOLED% = 0
     ULTCOLED% = 0
     NFT% = 0
     NCLF% = 0
     HAYTOTA% = 0
     RG% = 1
     BLABUFCA$ = ""
     '
5805 RG% = RG% + 1
     If RG% > LOF(N15%) / 128 Or RG% > 33 Then GoTo 5820
5806 Get #N15%, RG%, YYY$
     If YYY$ = String$(128, 0) Then GoTo 5820
     '
5807 NUCOLUM% = NUCOLUM% + 1
     TVL%(NUCOLUM%) = Asc(Mid$(YYY$, 33))
     LDL%(NUCOLUM%) = Asc(Mid$(YYY$, 34))
         If RG% = 2 Then TVL%(NUCOLUM%) = 3: LDL%(NUCOLUM%) = 0
5808 PR%(NUCOLUM%) = Asc(Mid$(YYY$, 35))
     While PR%(NUCOLUM%) < PR%(NUCOLUM% - 1)
        PR%(NUCOLUM%) = PR%(NUCOLUM%) + 256
     Wend
     LF%(NUCOLUM%) = Asc(Mid$(YYY$, 36))
         If RG% = 2 Then
           PR%(NUCOLUM%) = 0
           LF%(NUCOLUM%) = 2
         End If
5809 TT%(NUCOLUM%) = Asc(Mid$(YYY$, 88))
     If TT%(NUCOLUM%) > 0 Then
        HAYTOTA% = 1
        If TT%(NUCOLUM%) <> NUCOLUM% Then
            RTT%(TT%(NUCOLUM%) Mod 100) = 1
        End If
     End If
     '
5811 GRID.ColWidth(NUCOLUM% - 1) = 0
     ACL%(NUCOLUM%) = Asc(Mid$(YYY$, 69))
     LK%(NUCOLUM%) = Asc(Mid$(YYY$, 70))
     PINSE%(NUCOLUM%) = 1 + Len(BLABUFCA$)
     BLABUFCA$ = BLABUFCA$ + String$(1 + ACL%(NUCOLUM%), " ") + Chr$(9)
     If Asc(Mid$(YYY$, 69)) > 0 Then
       ECL$(NUCOLUM%) = "" 'REPLACEN$(Trim$(Mid$(YYY$, 37, 32)))
       TEPRUEBA.Caption = String$(1 + ACL%(NUCOLUM%), "0")
       GRID.ColWidth(NUCOLUM% - 1) = TEPRUEBA.Width
       GRID.Col = NUCOLUM% - 1
       GRID.Row = 0
       GRID.RowHeight(0) = 500
       GRID.CellFontSize = GRID.FontSize
          If Left$(ECL$(NUCOLUM%), 1) = "<" Then
              ECL$(NUCOLUM%) = Mid$(ECL$(NUCOLUM%), 2)
              GRID.CellAlignment = 1
            ElseIf Left$(ECL$(NUCOLUM%), 1) = ">" Then
              ECL$(NUCOLUM%) = Mid$(ECL$(NUCOLUM%), 2)
              GRID.CellAlignment = 7
            Else
              GRID.CellAlignment = 4
          End If
       GRID.Clip = Trim$(ECL$(NUCOLUM%))
       MG%(NUCOLUM%) = Asc(Mid$(YYY$, 71))
       PE%(NUCOLUM%) = Asc(Mid$(YYY$, 73))
          If RG% = 2 Then PE%(NUCOLUM%) = 2
       If PE%(NUCOLUM%) = 1 Then
          GRABACT% = 1
       End If
       JT%(NUCOLUM%) = Asc(Mid$(YYY$, 72))
       AE$(NUCOLUM%) = Trim$(Mid$(YYY$, 74, 8))
       CE%(NUCOLUM%) = Asc(Mid$(YYY$, 82))
       If TVL%(NUCOLUM%) = 5 Then
         AE$(NUCOLUM%) = "SELFECHA"
       End If
    End If
    '
    GoTo 5805
5817 '
5820 ATOT% = NUCOLUM% - 1 - NCLF%
     BLABUFCA$ = Left$(BLABUFCA$, Len(BLABUFCA$) - 1) + Chr$(13)
     GRID.Cols = NUCOLUM%
     CLII% = LEFTPOS% + 2
     For KCL% = 1 To NUCOLUM%
        LC%(KCL%) = ACL%(KCL%)
        If Left$(ECL$(KCL%), 1) <> "@" Then
            LC%(KCL%) = ACL%(KCL%)
            ATOT% = ATOT% + ACL%(KCL%)
            CL%(KCL%) = CLII%
            CLII% = CLII% + ACL%(KCL%) + 1
        End If
     Next KCL%
     '
5821 For KCL% = 1 To NUCOLUM%
        If CE%(KCL%) <= 0 Then GoTo 5824
5822    For KCL1% = 1 To NUCOLUM%
           If CL0%(KCL1%) = CE%(KCL%) Then
               CE%(KCL%) = KCL1%: PE%(KCL1%) = 3
               GoTo 5824
           End If
5823    Next KCL1%
5824 Next KCL%
     '
5830 NFT% = 0
     RG% = 1
     RG0% = RGI&
     RGFI% = 25: RGFS% = 1
     '
5835 RG% = RG% + 1
     If RG% > LOF(N15%) / 128 Or RG% > 33 Then GoTo 5850
5840 Get #N15%, RG%, YYY$
     If YYY$ = String$(128, 0) Then GoTo 5850
5841 If Asc(Mid$(YYY$, 83)) < 1 Then GoTo 5847 ' SUPRIME RG-FILTRO=0
5842 NFT% = NFT% + 1
     CFF%(NFT%) = RG% - 1
     LCF%(NFT%) = Asc(Mid$(YYY$, 83))
     RGFL&(NFT%) = Asc(Mid$(YYY$, 84))
     CLF%(NFT%) = Asc(Mid$(YYY$, 85))
5843 If RGFL&(NFT%) - 1 < RG0% Then
         RG0% = RGFL&(NFT%) - 1
     End If
     If RGFL&(NFT%) < RGFI% Then RGFI% = RGFL&(NFT%)
     If RGFL&(NFT%) > RGFS% Then RGFS% = RGFL&(NFT%)
     '
5844 ECLF$(NFT%) = RTrim$(Mid$(YYY$, 37, 32))
     If Len(ECLF$(NFT%)) > 0 Then
         ECLF$(NFT%) = ECLF$(NFT%) + ":"
     End If
5845 LKF%(NFT%) = Asc(Mid$(YYY$, 70))
     LEF%(NFT%) = Asc(Mid$(YYY$, 86))
     CEF%(NFT%) = Asc(Mid$(YYY$, 87))
5846 AEF$(NFT%) = Mid$(YYY$, 74, 8)
5847 GoTo 5835
 '
5850 Close #N15%: N15% = 0
     '
5700 PX% = InStr(UCase$(ATRIBU$), "TITUPAN=")
     If PX% > 0 Then
         XX$ = Mid$(ATRIBU$, PX% + 8)
         PY% = InStr(XX$, "/"): If PY% < 1 Then PY% = 1 + Len(XX$)
         XX$ = Left$(XX$, PY% - 1)
         If Trim$(XX$) <> "" Then
             Caption = XX$
         End If
     End If
     '
5705 PX% = InStr(UCase$(ATRIBU$), "ARCHIVO=")
     If PX% > 0 Then
         XX$ = Mid$(ATRIBU$, PX% + 8)
         PY% = InStr(XX$, "/"): If PY% < 1 Then PY% = 1 + Len(XX$)
         XX$ = Left$(XX$, PY% - 1)
         If Trim$(XX$) <> "" Then ARCHITRA$ = Trim$(XX$)
     End If
     '
5710 PX% = InStr(UCase$(ATRIBU$), "BORDESUP=")
     If PX% > 0 Then
         XX$ = Mid$(ATRIBU$, PX% + 9)
         PY% = InStr(XX$, "/"): If PY% < 1 Then PY% = 1 + Len(XX$)
         XX$ = Left$(XX$, PY% - 1)
         If Left$(XX$, 1) = "W" Then
              BSUW& = Val(Mid$(XX$, 2))
            Else
              RGI1& = Val(XX$)
              If RGI1& > 0 Then If RGI1& <= 22 Then RGI& = RGI1&
         End If
     End If
     '
5720 PX% = InStr(UCase$(ATRIBU$), "BORDEINF=")
     If PX% > 0 Then
         XX$ = Mid$(ATRIBU$, PX% + 9)
         PY% = InStr(XX$, "/"): If PY% < 1 Then PY% = 1 + Len(XX$)
         XX$ = Left$(XX$, PY% - 1)
         RGF1& = Val(XX$)
         If RGF1& > RGI& Then If RGF1& <= 22 Then RGF& = RGF1&
     End If
     If (RGF& - RGI&) < 12 Then RGF& = RGI& + 12
     '
5730 PX% = InStr(UCase$(ATRIBU$), "BORDEIZQ=")
     If PX% > 0 Then
         XX$ = Mid$(ATRIBU$, PX% + 9)
         PY% = InStr(XX$, "/"): If PY% < 1 Then PY% = 1 + Len(XX$)
         XX$ = Left$(XX$, PY% - 1)
         If Left$(XX$, 1) = "W" Then
               BIZW& = Val(Mid$(XX$, 2))
             Else
               LEFTPOX% = Val(XX$)
               If LEFTPOX% > 0 Then If LEFTPOX% <= 76 Then LEFTPOS% = LEFTPOX%
         End If
         '
         CLII% = LEFTPOS% + 2
         For KCL% = 1 To NUCOLUM%
            LC%(KCL%) = ACL%(KCL%)
            If Left$(ECL$(KCL%), 1) <> "@" Then
               LC%(KCL%) = ACL%(KCL%)
               ATOT% = ATOT% + ACL%(KCL%)
               CL%(KCL%) = CLII%
               CLII% = CLII% + ACL%(KCL%) + 1
            End If
         Next KCL%
         '
     End If
     '
5740 PX% = InStr(UCase$(ATRIBU$), "COLORPAN=")
     If PX% > 0 Then
         XX$ = Mid$(ATRIBU$, PX% + 9, 3)
         If Mid$(XX$, 2, 1) = "-" Then COLPAN$ = XX$
     End If
     '
5750 NOCIE% = 0
     If InStr(UCase$(ATRIBU$), "/NC") > 0 Then NOCIE% = 1
     '
5760 PREGU% = 0
     If InStr(UCase$(ATRIBU$), "/PS") > 0 Then PREGU% = 1
     '
5770 PX% = InStr(UCase$(ATRIBU$), "MODO=")
     MODOS$ = ""
     If PX% > 0 Then
         MODOS$ = Mid$(ATRIBU$, PX% + 5, 3)
     End If
     '
5780 NOFORM% = 0
     If InStr(UCase$(ATRIBU$), "/NOFORM") > 0 Then NOFORM% = 1
     '
5790 JK% = 0
     For UU% = 1 To NUCOLUM%
        If ACL%(UU%) > 0 Then JK% = JK% + 1
        ARGATRI$ = "ARCHECO(" + Trim$(Str$(JK%)) + ")="
        PX% = InStr(UCase$(ATRIBU$), ARGATRI$)
        If PX% > 0 Then
           XX$ = Mid$(ATRIBU$, PX% + Len(ARGATRI$))
           PY% = InStr(XX$, "/"): If PY% < 1 Then PY% = 1 + Len(XX$)
           XX$ = Left$(XX$, PY% - 1)
           If Trim$(XX$) <> "" Then AE$(UU%) = Trim$(XX$)
        End If
        If PE%(UU%) = 1 Then
          If ACL%(UU%) > 0 Then
            If PRICOLED% = 0 Then PRICOLED% = UU%
            ULTCOLED% = UU%
          End If
        End If
     Next UU%
     '
End Sub
'

Private Sub Text1_Change()
   '
   COLAC% = GRID.Col + 1
   Text1.ToolTipText = ""
   If AE$(COLAC%) <> "" Then
     ARCHEZ$ = Trim$(UCase$(AE$(COLAC%)))
     If ARCHEZ$ = "UNICOM" Then ARCHEZ$ = "UNIMED"
     If ARCHEZ$ = "MASTER" Or ARCHEZ$ = "OCOMEMI" Or ARCHEZ$ = "ECOMAST" Or Left$(ARCHEZ$, 5) = "LIPRE" Then
         TTTXT$ = DESCRIT0$(CODINT%(Text1))
       Else
         If Mid$(ARCHEZ$, 6, 3) = "###" Then
           NORDEX$ = Trim$(Str$(GRID.Row))
           While Len(NORDEX$) < 3: NORDEX$ = "0" + NORDEX$: Wend
           ARCHEZ$ = Left$(ARCHEZ$, 5) + NORDEX$
         End If
         TTTXT$ = Trim$(ECOARCH$(ARCHEZ$, Text1))
     End If
     Text1.ToolTipText = TTTXT$
   End If
   '
   If MODITEX% <> 1 Then Exit Sub
   Call MUESTRAECOS
   '
End Sub
'
Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
     COLAC% = GRID.Col + 1
     If AE$(COLAC%) <> "" Then
       ARCHEZ$ = Trim$(UCase$(AE$(COLAC%)))
       ARGUBUS$ = Trim$(Text1)
       If ARCHEZ$ = "MASTER" Or ARCHEZ$ = "ECOMAST" Or ARCHEZ$ = "OCOMEMI" Or Left$(ARCHEZ$, 5) = "LIPRE" Then
           Text1.Enabled = False
           Call MUESTRADATCOD(ARGUBUS$)
           Text1.Enabled = True
           Exit Sub
         ElseIf ARCHEZ$ = "PROVED1" Or ARCHEZ$ = "ACREDOR" Then
           Text1.Enabled = False
           Call MUESTRADATPRO(ARGUBUS$)
           Text1.Enabled = True
           Exit Sub
       End If
     End If
   End If
   '
End Sub

Private Sub Text1_DblClick()
   COLAC% = GRID.Col + 1
   If AE$(COLAC%) <> "" Then
     FORSELEC.SINI = Str$(Top + (Height - ScaleHeight) + Text1.Top + Text1.Height)
     FORSELEC.LINI = Str$(Left + Text1.Left)
     FORSELEC.ABUSQUE = Trim$(Text1)
     '
     ARCHEK$ = AE$(COLAC%)
     '
     If ARCHEK$ = "UNICOM" Then ARCHEK$ = "UNIMED"
     '
     If Mid$(ARCHEK$, 6, 3) = "###" Then
       NORDEX$ = Trim$(Str$(GRID.Row))
       While Len(NORDEX$) < 3: NORDEX$ = "0" + NORDEX$: Wend
       ARCHEK$ = Left$(AE$(COLAC%), 5) + NORDEX$
     End If
     '
     If ARCHEK$ = "LOTECOD" Then
       For KKI% = COLAC% To 1 Step -1
         If TVL%(KKI%) = 7 Then
           GRID.Col = KKI% - 1
           CIXI% = CODINT%(GRID.Text)
           GRID.Col = COLAC%
           VDEVU$ = LOTESEL$(CIXI%)
           GoTo 10
         End If
       Next KKI%
       Exit Sub
     End If
     '
     DELISECO$ = ""
     If Left$(UCase$(ARCHEK$), 5) = "LIPRE" Then
       DELISECO$ = DENOLISTA$(Val(Mid$(ARCHEK$, 6, 3)))
     End If
     Call SELECHO(ARCHEK$ + "/" + DELISECO$)
     VDEVU$ = VALACT1$(ARCHEK$)
     '
10   If VDEVU$ <> "" Then
       If JT%(GRID.Col + 1) = 2 Then
         While Len(VDEVU$) < ACL%(GRID.Col + 1)
           VDEVU$ = " " + VDEVU$
         Wend
       End If
       If Trim$(VDEVU$) <> Trim$(Text1) Then
         MODITEX% = 1
       End If
       Text1 = VDEVU$
       GRID.Text = VDEVU$
       Call GRAVACOLRO
       Call MUESTRAECOS
       RECIBECLICK% = (-1)
       Call SETNEXTCOL
       RECIBECLICK% = 0
     End If
   End If
End Sub

Private Sub Text1_GotFocus()
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1)
  TEXTOINI$ = Trim$(Text1.Text)
  GRID.Col = COLTEBO% - 1
  GRID.Row = ROWTEBO%
  Text1.Refresh
End Sub

Function APROENTRA%()
   '
   RECIBECLICK% = -1
   APROENTRA% = 1
   If Trim$(Text1) <> TEXTOINI$ Then 'TRIM$(GRID.TEXT) Then
     '
     COLAC% = GRID.Col + 1
   
     VDEVU$ = Text1
     VABIN$ = CBIN$(VDEVU$, TVL%(COLAC%), LF%(COLAC%))
     If AE$(COLAC%) <> "" Then
       '
       ARCHEK$ = AE$(COLAC%)
       If ARCHEK$ = "UNICOM" Then ARCHEK$ = "UNIMED"
       '
       If Mid$(ARCHEK$, 6, 3) = "###" Then
         NORDEX$ = Trim$(Str$(GRID.Row))
         While Len(NORDEX$) < 3: NORDEX$ = "0" + NORDEX$: Wend
         ARCHEK$ = Left$(AE$(COLAC%), 5) + NORDEX$
       End If
       '
       If AE$(COLAC%) = "FECHA" Or AE$(COLAC%) = "SELFECHA" Then
           If CVI(VABIN$) < 1 Then
             Call MENSERR(24, "Ingreso no Válido")
             MODITEX% = 1
             GoTo 90
           End If
         ElseIf ECOARCH(ARCHEK$, VABIN$) = "" Then
           Call MENSERR(24, "Ingreso no Válido")
           MODITEX% = 1
           GoTo 90
       End If
       '
     End If
     '
10   ROWAC% = GRID.Row
     If ROWAC% > 0 Then
       VANUTE$ = CSTRING$(VABIN$, TVL%(COLAC%), ACL%(COLAC%), LDL%(COLAC%), JT%(COLAC%))
       If PE%(COLAC%) = 0 Or PE%(COLAC%) = 2 Then
         RECIBECLICK% = -1
         RINSERO& = INSEROW&
         For U& = 1 To ULTREG&(ARCHILOC$)
           If U& <> ROWAC% Then
             GRID.Row = U&
             If Left$(GRID.Text, ACL%(COLAC%)) = VANUTE$ Then
               RECIBECLICK% = 0
               Call SETNEXTCOL
               GoTo 90
             End If
           End If
         Next U&
         RECIBECLICK% = 0
       End If
       GRID.Row = ROWAC%
       GRID.Col = COLAC% - 1
       INSEROW& = RINSERO&
       GRID.Text = VANUTE$
       Call GRAVACOLRO
       Call MUESTRAECOS
     End If
     '
     APROENTRA% = 1
     RECIBECLICK% = 0
     Exit Function
     '
90   APROENTRA% = 0
     Call MUETEBOX
     If GRID.Row <= ULTREG&(ARCHILOC$) Then
       Call MUESTRAECOS
     End If
   End If
   RECIBECLICK% = 0
   '
End Function

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
MODITEX% = 0
If KeyCode = 40 Then
     If APROENTRA% <> 1 Then Exit Sub
     If GRID.Row < UREPOS& Then
       If GRID.Row >= ULTREG&(ARCHILOC$) Then
         GRID.Col = 0
       End If
       GRID.Row = GRID.Row + 1
     End If
     If GRID.Row > ULTREG&(ARCHILOC$) Then
       GRID.Col = 0
     End If
     If PE%(GRID.Col + 1) <> 1 Then Call SETNEXTCOL
     Call MUETEBOX
   ElseIf KeyCode = 38 Then
     If APROENTRA% <> 1 Then Exit Sub
     If GRID.Row > 1 Then
       GRID.Row = GRID.Row - 1
       If PE%(GRID.Col + 1) <> 1 Then Call SETNEXTCOL
     End If
     If GRID.TopRow > GRID.Row Then
       GRID.TopRow = GRID.Row
     End If
     Call MUETEBOX
   ElseIf KeyCode = 39 Then
     If Text1.SelStart >= Len(RTrim(Text1)) Or Text1.SelText = Text1 Then
       If APROENTRA% <> 1 Then Exit Sub
       If GRID.Row > ULTREG&(ARCHILOC$) Then GRID.Col = 0
       Call SETNEXTCOL
     End If
   ElseIf KeyCode = 37 Then
     If Text1.SelText <> Text1 Or Text1 = "" Then
       If Text1.SelStart < 1 Then
         If APROENTRA% <> 1 Then Exit Sub
         Call SETPREVCOL
       End If
     End If
   ElseIf KeyCode = vbKeyF2 Then
     COLAC% = GRID.Col + 1
     If AE$(COLAC%) <> "" Then
       FORSELEC.SINI = Str$(Top + (Height - ScaleHeight) + Text1.Top + Text1.Height)
       FORSELEC.LINI = Str$(Left + Text1.Left)
       FORSELEC.ABUSQUE = Trim$(Text1)
       '
       ARCHEK$ = AE$(COLAC%)
       '
       If ARCHEK$ = "UNICOM" Then ARCHEK$ = "UNIMED"
       '
       If Mid$(ARCHEK$, 6, 3) = "###" Then
         NORDEX$ = Trim$(Str$(GRID.Row))
         While Len(NORDEX$) < 3: NORDEX$ = "0" + NORDEX$: Wend
         ARCHEK$ = Left$(AE$(COLAC%), 5) + NORDEX$
       End If
       '
       Call SELECHO(ARCHEK$)
       VDEVU$ = VALACT1$(ARCHEK$)
       If VDEVU$ <> "" Then
         If JT%(GRID.Col + 1) = 2 Then
           While Len(VDEVU$) < ACL%(GRID.Col + 1)
             VDEVU$ = " " + VDEVU$
           Wend
         End If
         If Trim$(VDEVU$) <> Trim$(Text1) Then
           MODITEX% = 1
         End If
         Text1 = VDEVU$
         GRID.Text = VDEVU$   'Text1
         Call GRAVACOLRO
         Call MUESTRAECOS
         RECIBECLICK% = (-1)
         Call SETNEXTCOL
         RECIBECLICK% = 0
       End If
     End If
   ElseIf KeyCode = 36 Then    ' home
     If APROENTRA% <> 1 Then Exit Sub
     If GRID.Col > PRICOLED% - 1 Then
         GRID.Col = PRICOLED% - 1
       Else
         GRID.Row = 1
         If GRID.TopRow > GRID.Row Then
           GRID.TopRow = GRID.Row
         End If
     End If
     Call MUETEBOX
   ElseIf KeyCode = 35 Then    ' END
     If APROENTRA% <> 1 Then Exit Sub
     If GRID.Col < ULTCOLED% - 1 Then
         GRID.Col = ULTCOLED% - 1
       Else
         GRID.Row = UREPOS&
         If GRID.Row > ULTREG&(ARCHILOC$) Then
           GRID.Col = 0
           Call SETNEXTCOL
         End If
         If GRID.TopRow < GRID.Row - VISIROWS& + 3 Then
           GRID.TopRow = GRID.Row - VISIROWS& + 3
         End If
     End If
     Call MUETEBOX
   ElseIf KeyCode = 8 Then    ' backspace
     If Text1.SelText = Text1.Text Then
       If APROENTRA% <> 1 Then Exit Sub
       Call SETPREVCOL
       Call MUETEBOX
     End If
   ElseIf KeyCode = 27 Then    ' ESCAPE
     If APROENTRA% <> 1 Then Exit Sub
     On Error Resume Next
     Text1.Visible = False
     BOTREC.SetFocus
     On Error GoTo 0
   ElseIf KeyCode = 33 Then
     If APROENTRA% <> 1 Then Exit Sub
     NUEROW& = GRID.Row - VISIROWS + 2
     If NUEROW& < 1 Then NUEROW& = 1
     DESPLARO& = GRID.Row - NUEROW&
     GRID.Row = NUEROW&
     If GRID.TopRow - DESPLARO& > 0 Then
         GRID.TopRow = GRID.TopRow - DESPLARO&
       Else
         GRID.TopRow = 1
     End If
     If PE%(GRID.Col + 1) <> 1 Then Call SETNEXTCOL
     Call MUETEBOX
   ElseIf KeyCode = 34 Then
     If APROENTRA% <> 1 Then Exit Sub
     NUEROW& = GRID.Row + VISIROWS - 2
     If NUEROW& > UREPOS& Then NUEROW& = UREPOS&
     DESPLARO& = NUEROW& - GRID.Row
     GRID.Row = NUEROW&
     If GRID.TopRow < GRID.Row - VISIROWS + 2 Then
       GRID.TopRow = GRID.TopRow + DESPLARO&
     End If
     If PE%(GRID.Col + 1) <> 1 Then Call SETNEXTCOL
     If GRID.Row > ULTREG&(ARCHILOC$) Then
       GRID.Col = 0
       Call SETNEXTCOL
     End If
     Call MUETEBOX
End If
'
If KeyCode% = 8 Or KeyCode = 32 Or KeyCode > 40 Then
  MODITEX% = 1
End If
'
End Sub

'
Private Sub Text1_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 13 Then
     If APROENTRA% = 1 Then
       If GRID.Row > ULTREG&(ARCHILOC$) Then GRID.Col = 0
       RECIBECLICK% = (-1)
       Call SETNEXTCOL
       RECIBECLICK% = 0
     End If
   End If
   '
End Sub

Sub GRAVACOLRO()
   '
   COLAC% = GRID.Col + 1
   REGAC& = GRID.Row
   If REGAC& > 0 Then
     If REGAC& <= UREPOS& Then
       VDEVU$ = GRID.Text
       VABIN$ = CBIN$(VDEVU$, TVL%(COLAC%), LF%(COLAC%))
       If REGAC& <= ULTREG&(ARCHILOC$) Then
           XX$ = REGLEIDO$(ARCHILOC$, REGAC&)
         Else
           XX$ = REGBLAN$(ARCHILOC$)
       End If
       Call REPLA(XX$, VABIN$, PR%(COLAC%), LF%(COLAC%))
       Call GRAREG(ARCHILOC$, XX$, REGAC&)
     End If
   End If
   '
End Sub
'
Sub SETNEXTCOL()
   '
   If YACARDA% < 1 Then Exit Sub
5  COLNUE = GRID.Col + 1
   ROWNUE = GRID.Row
   '
10 If ROWNUE > UREPOS& Then If APPI% < 1 Then Exit Sub
   If COLNUE >= NUCOLUM% Then
     COLNUE = 0
     If ROWNUE < UREPOS& Then
         ROWNUE = ROWNUE + 1
       Else
         If APPI% < 1 Then Exit Sub
     End If
   End If
   If ROWNUE <= ULTREG&(ARCHILOC$) Then
     If ROWNUE <> INSEROW& Then
       If PE%(COLNUE + 1) <> 1 Or ACL%(COLNUE + 1) < 1 Then
         COLNUE = COLNUE + 1
         GoTo 10
       End If
     End If
   End If
   GRID.Col = COLNUE
   GRID.Row = ROWNUE
   If GRID.CellWidth < 0 Then GoTo 5
   Call MUETEBOX
   '
End Sub

Sub SETPREVCOL()
   '
   COLNUE = GRID.Col - 1
   '
10 If COLNUE < 0 Then
     If GRID.Row > 1 Then
       GRID.Row = GRID.Row - 1
       If GRID.TopRow > GRID.Row Then
         GRID.TopRow = GRID.Row
       End If
     End If
     COLNUE = NUCOLUM% - 1
   End If
   '
   If PE%(COLNUE + 1) <> 1 Or ACL%(COLNUE + 1) < 1 Then
     COLNUE = COLNUE - 1
     GoTo 10
   End If
   GRID.Col = COLNUE
   Call MUETEBOX
   '
End Sub

Private Sub Text2_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Timer1_Timer()
   'Static ACOLANT
   '
   'If Visible = False Then Exit Sub
   '
   'ACOLTOT = 380 ' 350
   'For KKI% = 1 To GRID.Cols
   '  ACOLTOT = ACOLTOT + GRID.ColWidth(KKI% - 1)
   'Next KKI%
   'If ACOLTOT <> ACOLANT Or YACARDA% = 0 Then
   '  Call FRESHALL
   '  WD1 = ACOLTOT + 750 + MAINT + MADER
   '  'If WD1 > 12000 - Left Then WD1 = 12000 - Left
   '  If WD1 > Screen.Width - Left Then WD1 = Screen.Width - Left
   '  Width = WD1
   '  FRAMEBOT1.Left = Width - 750 - MADER
   '  FRAMEBOT2.Left = Width - 750 - MADER
   '  MARCOBARRA.Left = Width - 750 - MADER
   '  ACOLANT = ACOLTOT
   '  GRID.Width = Width - 750 - MAINT - MADER
   '  Call FRESHALL
   '  '
   '  If PRICOLED% < 1 Then
   '    If YACARDA% > 0 Then
   '      GRID.Row = UREPOS&
   '    End If
   '  End If
   '  Call FRESHALL
   '  Call MUETEBOX
   '  Call FRESHALL
   'End If
   ''
   'If YACARDA% = 0 Then
   '  If Visible = True Then
   '    Call CARLIVEN
   '    YACARDA% = 1
   '    GRID.Row = 1
   '    GRID.Col = 0
   '    If PRICOLED% < 1 Then
   '      GRID.Row = UREPOS&
   '      If GRID.Row <= ULTREG&(ARCHILOC$) Then
   '        GoTo 10
   '      End If
   '    End If
   '    Call SETNEXTCOL
   '    '
10 '    Call MUETEBOX
   '  End If
   'End If
   '
End Sub

Sub DIBUVEN()     ' Rutina de dibujo pantalla tabulada
   '
RGF& = 25: RGI& = 3
   HREN% = ScaleHeight / 25
   MAIZQ = 0
   MAINT = 150
   MADER = 200
   '
   ATOT% = 1
   For KCL% = 1 To NUCOLUM%
     If ACL%(KCL%) > 0 Then
       ATOT% = ATOT% + ACL%(KCL%) + 1
     End If
   Next KCL%
   '
   ' para aparear con vieja version
   Height = Screen.Height / 25 * (RGF& - RGI& - 1)
   LIPRUEBA.Height = ScaleHeight - 768
   Height = LIPRUEBA.Height + 768
   '
   ' para aparear con vieja version
   Top = Screen.Height / 25 * (RGI&)
   If BSUW& <> 0 Then Top = BSUW&
   HBAS = Screen.Height / 25 * (RGF& - RGI& - 1)
   Top = Top + (HBAS - Height) / 2
   '
   ' para aparear con vieja version
   ANCHOTOT% = 4
   For KKI% = 1 To NUCOLUM%
     If ACL%(KKI%) > 0 Then
       ANCHOTOT% = ANCHOTOT% + ACL%(KKI%) + 1
     End If
   Next KKI%
   TEPRUANCH.Caption = String$(ANCHOTOT%, "*")
   Width = TEPRUANCH.Width + 750 + MAINT + MADER
   '
   ' para aparear con vieja version
   If BIZW& <> 0 Then
       Left = BIZW&
     Else
       If LEFTPOS% < 1 Then LEFTPOS% = 1 + Int((Screen.Width - Width) / 300 + 0.5)
       If LEFTPOS% < 1 Then LEFTPOS% = 1
       Left = Screen.Width / 80 * (LEFTPOS% - 1)
   End If
   '
   GRID.Top = 0
   GRID.Left = MAIZQ
   HGRID = GRID.RowHeight(0)
   While HGRID + GRID.RowHeight(1) < ScaleHeight - 80
     HGRID = HGRID + GRID.RowHeight(1) + 8
   Wend
   GRID.Height = HGRID
   DIBORINF = ScaleHeight - GRID.Height
   If DIBORINF > 125 Then Height = Height + 125 - DIBORINF
   '
   TEPRUANCH.Caption = String$(ANCHOTOT% - 3, "*")
   ACOLTO = 0
   For KKI% = 1 To GRID.Cols
     ACOLTO = ACOLTO + GRID.ColWidth(KKI% - 1)
   Next KKI%
   COEFAN = TEPRUANCH.Width / ACOLTO
   For KKI% = 1 To GRID.Cols
     GRID.ColWidth(KKI% - 1) = COEFAN * GRID.ColWidth(KKI% - 1)
   Next KKI%
   '
   ALIBRE = GRID.Height - GRID.RowHeight(0)
   BOTEXIT.Visible = True
   BOTREC.Top = 870
   FRAMEBOT1.Height = BOTREC.Top + BOTREC.Height + 100
   '
   If FRAMEBOT1.Height + FRAMEBOT2.Height + 120 > ALIBRE Then
     BOTEXIT.Visible = False
     BOTREC.Top = BOTEXIT.Top
   End If
   FRAMEBOT1.Height = BOTREC.Top + BOTREC.Height + 100
   FRAMEBOT1.Top = GRID.Top + GRID.RowHeight(0) - 90
   FRAMEBOT1.Left = MAIZQ + GRID.Width + MAINT
   FRAMEBOT2.Top = GRID.Top + GRID.Height - FRAMEBOT2.Height + 20
   FRAMEBOT2.Left = MAIZQ + GRID.Width + MAINT
   '
   MARCOBARRA.Top = 180
   MARCOBARRA.Left = MAIZQ + GRID.Width + MAINT + 20
   '
End Sub
'

Sub CARLIVEN()
    '
    Screen.MousePointer = 11
    TINI = Timer
    FIN& = ULTREG&(ARCHILOC$)
    'FINTE& = GRID.Rows - 1
    'If FINTE& > FIN& Then
    '  YACA1% = YACARDA%
    '  YACARDA% = 1
    '  Call LIMPIAGRID
    '  YACARDA% = YACA1%
    'End If
    '
    GRID.Rows = 2 + (GRID.Height - GRID.RowHeight(0)) / GRID.RowHeight(1)
    VISIROWS& = GRID.Rows
    If FIN& > GRID.Rows - 1 Then GRID.Rows = FIN& + 1
    TETUX$ = String$(FIN& * Len(BLABUFCA$), 0)
    '
    POIN& = 1
    For KKI% = 1 To GRID.Cols
      If GRID.ColWidth(KKI% - 1) > 0 Then
        JUCOL% = 1: If JT%(KKI%) = 1 Then JUCOL% = 0
        GRID.ColAlignment(KKI% - 1) = JUCOL%
      End If
    Next KKI%
    '
    GRID.FillStyle = 1
    INIK% = 1
12  For KKI% = INIK% To GRID.Cols
      If GRID.ColWidth(KKI% - 1) > 0 Then
        If FIN& > 0 Then
          If PE%(KKI%) <> 1 Then
            FINK% = KKI%
            '
15          If FINK% < NUCOLUM% Then
              If GRID.ColWidth(FINK%) < 1 Or PE%(FINK% + 1) <> 1 Then
                FINK% = FINK% + 1
                GoTo 15
              End If
            End If
            '
            GRID.Col = KKI% - 1
            GRID.Row = 1
            GRID.ColSel = FINK% - 1
            GRID.RowSel = FIN&
            GRID.CellBackColor = RGB(230, 230, 230)
            INIK% = FINK% + 1
            GoTo 12
            '
          End If
        End If
      End If
    Next KKI%
    '
    GRID.FillStyle = 0
    GRID.Row = 1
    '
    For U& = 1 To FIN&
      '
      Call TRACE(20, U&, FIN&)
      BUFCARGA$ = BLABUFCA$
701   XXX$ = REGLEIDO$(ARCHILOC$, U&)
      RELIB$ = Space$(ATOT% + 2)
      '
      For K2% = 1 To NUCOLUM%
        If CFT%(K2%) <> 1 Then
          If RTT%(K2%) < 1 Then
            If LF%(K2%) > 0 Then
              If PR%(K2%) > 0 Then
                If ACL%(K2%) > MG%(K2%) Then
                  XXZ$ = Mid$(XXX$, PR%(K2%), LF%(K2%))
                  ACLX% = ACL%(K2%)
                  TXT$ = CSTRING$(XXZ$, TVL%(K2%), ACLX%, LDL%(K2%), JT%(K2%))
                  Mid$(BUFCARGA$, PINSE%(K2%), ACL%(K2%)) = TXT$
                 '
                  If Trim$(AE$(K2%)) <> "" Then
                    COLOG% = K2%
                    CLE% = CE%(COLOG%)
                    If CLE% > 0 Then
                      LCE% = LC%(CE%(COLOG%))
                      If LCE% > 0 Then
                        '
                        ARCHEK$ = AE$(K2%)
                        If ARCHEK$ = "UNICOM" Then ARCHEK$ = "UNIMED"
                        '
                        If Left$(ARCHEK$, 5) = "LIPRE" Then
                          ARCHEK$ = "MASTER"
                        End If
                        '
                        If Mid$(ARCHEK$, 6, 3) = "###" Then
                          NORDEX$ = Trim$(Str$(U&))
                          While Len(NORDEX$) < 3: NORDEX$ = "0" + NORDEX$: Wend
                          ARCHEK$ = Left$(AE$(K2%), 5) + NORDEX$
                        End If
                        '
                        XXE$ = ECOARCH$(ARCHEK$, XXZ$)
                        JTE% = JT%(CE%(COLOG%))
                        TXT$ = CSTRING$(XXE$, 6, Len(XXE$), 0, JTE%)
                        Mid$(BUFCARGA$, PINSE%(CE%(COLOG%)), ACL%(CE%(COLOG%))) = TXT$
                      End If
                    End If
                  End If
                  '
90                TTT% = TT%(CFT%(K2%))
                  If TTT% > 0 Then
                    If (TTT% Mod 100) <= NUCOLUM% Then
                      TTX#(K2%) = TTX#(K2%) + Val(TXT$)
                      If TTT% < 100 Then TXX#(TTT%) = TXX#(TTT%) + Val(TXT$) Else TXX#(TTT% Mod 100) = TXX#(TTT% Mod 100) - Val(TXT$)
                      If (TTT% Mod 100) <> CFT%(K2%) Then
                        COLOG% = CLT%(TTT% Mod 100)
                        CLE% = CL%(COLOG%)
                        LCE% = LC%(COLOG%)
                        JTE% = JT%(COLOG%)
                        TXT$ = CSTRING$(MKD$(TXX#(TTT% Mod 100)), TVL%(CFT%(COLOG%)), LCE%, LDL%(CFT%(COLOG%)), JTE%)
                        If CLE% > 0 Then Mid$(RELIB$, CLE% - LEFTPOS%, LCE%) = TXT$
                      End If
                    End If
                  End If
                  '
                End If
              End If
            End If
          End If
        End If
      Next K2%
      '
      Mid$(TETUX$, POIN&, Len(BUFCARGA$)) = BUFCARGA$
      POIN& = POIN& + Len(BUFCARGA$)
      '
989 Next U&
    '
    While POIN& < Len(TETUX$)
      Mid$(TETUX$, POIN&, Len(BLABUFCA$)) = BLABUFCA$
      POIN& = POIN& + Len(BLABUFCA$)
    Wend
    TETUX$ = Left$(TETUX$, POIN&)
    GRID.Col = 0
    GRID.Row = 1
    GRID.ColSel = NUCOLUM% - 1
    GRID.RowSel = FIN&
    GRID.Clip = TETUX$
    GRID.FillStyle = 0
    '
    GRID.Row = 1
    GRID.Col = 0
    GRID.TopRow = 1
    GRID.Refresh
    If PRICOLED% < 1 Then
      GRID.Row = UREPOS&
      If GRID.Row <= ULTREG&(ARCHILOC$) Then
        Screen.MousePointer = 1
        Exit Sub
      End If
    End If
    Call SETNEXTCOL
    Call MUETEBOX
    '
    Screen.MousePointer = 1
    '
End Sub

Sub EXTRAPROF()
   '
   Timer1.Enabled = False
   'Call EXTRACT(PROFIL$)
   Timer1.Enabled = True
   '
   If EXISTE%(PROFIL$) < 1 Then
     FORMB$ = UCase$(REPLACAR$(PROFIL$, "\", "/"))
     Call MENSERR(24, "Proceso Interrumpido.\Falta Controlador de Procesos '" + FORMB$ + "'.")
     Close
     Call CIERRARCH("*.*")
     Call FINAL("")
   End If
   '
End Sub

Sub MUETEBOX()
   '
   Text1.Visible = False
   If YACARDA% = 0 Then Exit Sub
   '
   If GRID.Row >= UREPOS& Then Text1 = ""
   If GRID.Row > UREPOS& Then Exit Sub
   '
   If GRID.RowIsVisible(GRID.Row) = True Then
     If GRID.Row > 0 Then
     If PE%(GRID.Col + 1) = 1 Or GRID.Row > ULTREG&(ARCHILOC$) Or GRID.Row = INSEROW& Then
     If GRID.CellWidth > 0 Then
       Text1.Top = GRID.CellTop + GRID.Top - 20
       Text1.Left = GRID.CellLeft
       Text1.Height = GRID.CellHeight - 10
       Text1.Width = GRID.CellWidth
       TEPRUTEX = Trim$(GRID.Text)
       If JT%(GRID.Col + 1) = 2 Then
         While TEPRUTEX.Width < Text1.Width
           TEPRUTEX = " " + TEPRUTEX
         Wend
         TEPRUTEX = Mid$(TEPRUTEX, 2)
       End If
       MODITEX% = 0
       Text1 = TEPRUTEX
       Text1.Visible = True
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
       Text1.SelStart = 0
       MODITEX% = 1
       COLTEBO% = GRID.Col + 1
       ROWTEBO% = GRID.Row
     End If
     End If
     End If
   End If
   '
End Sub

Function UREPOS&()
   '
   'UREPOS& = ULTREG&(ARCHILOC$)
   If APPI% = 1 Then UREPOS& = UREPOS& + 1
   If GRID.Rows < UREPOS& + 1 Then
     GRID.Rows = UREPOS& + 1
   End If
   '
End Function

Sub MUESTRAECOS()
   '
   If GRID.Row < 1 Then Exit Sub
   '
   COLAC% = GRID.Col + 1
   If AE$(COLAC%) <> "" Then
     CLE% = CE%(COLAC%)
     If CLE% > 0 Then
       LCE% = LC%(CE%(COLAC%))
       If LCE% > 0 Then
         '
         ARCHEK$ = AE$(COLAC%)
         If ARCHEK$ = "UNICOM" Then ARCHEK$ = "UNIMED"
         '
         If Left$(ARCHEK$, 5) = "LIPRE" Then
           ARCHEK$ = "MASTER"
         End If
         '
         If Mid$(ARCHEK$, 6, 3) = "###" Then
           NORDEX$ = Trim$(Str$(GRID.Row))
           While Len(NORDEX$) < 3: NORDEX$ = "0" + NORDEX$: Wend
           ARCHEK$ = Left$(AE$(COLAC%), 5) + NORDEX$
         End If
         '
         XXE$ = ECOARCH$(ARCHEK$, CBIN$(Text1, TVL%(COLAC%), LF%(COLAC%)))
         JTE% = JT%(CE%(COLAC%))
         TXT$ = CSTRING$(XXE$, 6, Len(XXE$), 0, 1)
         RECIBECLICK% = -1
         GRID.Col = CLE% - 1
         GRID.Text = TXT$
         If ACL%(CLE%) > 0 Then
           If GRID.Row <= ULTREG&(ARCHILOC$) Then
             Call GRAVACOLRO
           End If
         End If
         GRID.Col = COLAC% - 1
         RECIBECLICK% = 0
         '
         If ARCHEK$ = "MASTER" Or ARCHEK$ = "MATPRIM" Then
           CIKI% = CODINT%(Text1)
           If CIKI% > 0 Then
             For COLEC% = COLAC% + 1 To GRID.Cols
               If AE$(COLEC%) = "UNIMED" Then
                    JTE% = 1
                    RECIBECLICK% = (-1)
                    GRID.Col = COLEC% - 1
                    GRID.Text = UNIMED$(CIKI%)
                    If ACL%(COLEC%) > 0 Then
                      If GRID.Row <= ULTREG&(ARCHILOC$) Then
                        Call GRAVACOLRO
                      End If
                    End If
                    GRID.Col = COLAC% - 1
                    RECIBECLICK% = 0
                  ElseIf AE$(COLEC%) = "UNICOM" Then
                    JTE% = 1
                    RECIBECLICK% = (-1)
                    GRID.Col = COLEC% - 1
                    GRID.Text = UNICOM$(CIKI%)
                    If ACL%(COLEC%) > 0 Then
                      If GRID.Row <= ULTREG&(ARCHILOC$) Then
                        Call GRAVACOLRO
                      End If
                    End If
                    GRID.Col = COLAC% - 1
                    RECIBECLICK% = 0
               End If
             Next COLEC%
           End If
         End If
         '
         If ECOPRECIO% >= 0 Then
           If Left$(AE$(COLAC%), 5) = "LIPRE" Then
             ARCHEC$ = AE$(COLAC%)
             NLIS% = Val(Mid$(ARCHEC$, 6, 3))
             If NLIS% > 0 Then
               RECIBECLICK% = (-1)
               For KK1% = CE%(COLAC%) + 1 To NUCOLUM%
                 If ACL%(KK1%) > 0 Then
                   PRELISMO# = PRELISTA#(ARCHEC$, CVI(CBIN$(Text1, TVL%(COLAC%), LF%(COLAC%))))
                   PRELIX$ = Trim$(CSTRING$(MKD$(PRELISMO#), 3, 12, LDL%(KK1%), 2))
                   GRID.Col = KK1% - 1
                   JTE% = JT%(KK1%)
                   TXT$ = CSTRING$(PRELIX$, 6, ACL%(KK1%), 0, JTE%)
                   GRID.Text = TXT$
                   If GRID.Row <= ULTREG&(ARCHILOC$) Then
                     Call GRAVACOLRO
                   End If
                   GRID.Col = COLAC% - 1
                   GoTo 90
                 End If
               Next KK1%
             End If
           End If
         End If
         '
90       RECIBECLICK% = 0
         '
       End If
     End If
   End If
   '
End Sub

Sub LIMPIAGRID()
   '
   'If YACARDA% < 1 Then Exit Sub
   '
   FINTE& = GRID.Rows - 1
   TETUX$ = String$(FINTE& * Len(BLABUFCA$), 0)
   POIN& = 1
   While POIN& < Len(TETUX$)
     Mid$(TETUX$, POIN&, Len(BLABUFCA$)) = BLABUFCA$
     POIN& = POIN& + Len(BLABUFCA$)
   Wend
   '
   On Error Resume Next
   LIMPIAGR% = 1
   GRID.FillStyle = 1
   GRID.Col = 0
   GRID.Row = 1
   GRID.ColSel = NUCOLUM% - 1
   GRID.RowSel = FINTE&
   GRID.Clip = TETUX$
   GRID.FillStyle = 0
   LIMPIAGR% = 0
   On Error GoTo 0
   '
End Sub
