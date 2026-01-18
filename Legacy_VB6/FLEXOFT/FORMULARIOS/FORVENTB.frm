VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FORVENTB 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11010
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MARCOBARRA 
      Appearance      =   0  'Flat
      BackColor       =   &H002B9973&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   7665
      ScaleHeight     =   165
      ScaleWidth      =   630
      TabIndex        =   23
      Top             =   120
      Width           =   660
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H0048CC9D&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   8760
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   18
      Top             =   1560
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EEEEDD&
      BorderStyle     =   0  'None
      Height          =   100
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   105
      ScaleWidth      =   15000
      TabIndex        =   17
      Top             =   7875
      Width           =   15000
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
      Left            =   4725
      TabIndex        =   16
      Top             =   4830
      Width           =   175
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4F9FD&
      Height          =   1095
      Left            =   3255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   5880
      Width           =   1485
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7545
      TabIndex        =   13
      Top             =   1920
      Visible         =   0   'False
      Width           =   900
      Begin VB.CommandButton Command3 
         Caption         =   "Orden"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   30
         TabIndex        =   14
         ToolTipText     =   "Ordenamiento por Columnas"
         Top             =   200
         Width           =   825
      End
   End
   Begin VB.Frame FRAMEBOT2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   7560
      TabIndex        =   12
      Top             =   2640
      Width           =   900
      Begin VB.CommandButton BOTINSE 
         Caption         =   "Insertar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   30
         TabIndex        =   22
         ToolTipText     =   "Intercala Registro Antes del Activo"
         Top             =   200
         Width           =   825
      End
      Begin VB.CommandButton BOTBUSCA 
         Caption         =   "Buscar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   30
         TabIndex        =   21
         ToolTipText     =   "Buscar Texto por Campo"
         Top             =   850
         Width           =   825
      End
      Begin VB.CommandButton BOTERASE 
         Caption         =   "Borrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   30
         TabIndex        =   20
         ToolTipText     =   "Elimina Registro Activo"
         Top             =   520
         Width           =   825
      End
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   840
      TabIndex        =   5
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
      Height          =   1590
      Left            =   5460
      TabIndex        =   9
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
         Height          =   300
         ItemData        =   "FORVENTB.frx":0000
         Left            =   210
         List            =   "FORVENTB.frx":0007
         TabIndex        =   10
         Top             =   315
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.Label TEPRUANCH 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "TEPRUANCH"
         Height          =   225
         Left            =   210
         TabIndex        =   11
         Top             =   1050
         Visible         =   0   'False
         Width           =   1080
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4F9FD&
      Height          =   360
      Left            =   3255
      TabIndex        =   4
      Top             =   4800
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
      _Version        =   393216
      Rows            =   16
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   8433840
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      FocusRect       =   2
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Frame FRAMEBOT1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1520
      Left            =   7545
      TabIndex        =   0
      Top             =   315
      Width           =   900
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
         Left            =   180
         Picture         =   "FORVENTB.frx":0021
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Cancela los Datos Ingresados"
         Top             =   240
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
         Left            =   180
         Picture         =   "FORVENTB.frx":032B
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Graba - Guardar Modificaciones"
         Top             =   840
         Width           =   540
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FORVENTB.frx":0635
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "FORVENTB.frx":0869
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image Image1 
      Height          =   295
      Left            =   9765
      Picture         =   "FORVENTB.frx":08E3
      Stretch         =   -1  'True
      ToolTipText     =   "Doble-Click para Bloquear Posición"
      Top             =   630
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label TEPRUTEX 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   225
      Left            =   3255
      TabIndex        =   8
      Top             =   5355
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   225
      Left            =   2415
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label PROFILE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   315
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   1380
   End
End
Attribute VB_Name = "FORVENTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim YACARDA%
Dim GRABACT%      ' GRABABLE=1 - NO GRABA=0
Dim BIZW&, BSUP&  ' POSICION DEL FORMULARIO
Dim BSUPER, BIZQUI, BALTUR, BANCHO
Dim LEFTPOS%      ' espacios libres a izquierda del formulario
Dim PROFIL$
Dim ARCHITRA$, ARCHILIB$
Dim ATRIBU$
Dim YYY As String * 128
Dim CTTT%
Dim cl%(32), CE%(32), lC%(32), PE%(32)
Dim NUCOLUM%
Dim PRICOLED%           ' primera columna editable
Dim ULTCOLED%           ' ultima columna editable
Dim MAIZQ, MAINT, MADER
Dim BLABUFCA As String
Dim PINSE%(64)
Dim PUEDINSLIN%
Dim UROW, UCOL
Dim COLTEBO%, ROWTEBO%  ' COLUMNA Y RENGLON CAJA TEXTO
Dim VISIROWS&           ' RENGLONES VISIBLES
Dim INSEROW&
Dim RECIBECLICK%
Dim MODITEX%
Dim TEXTOINI$
Dim MAXCARS%
Dim LIMPIAGR%
Dim COLUSOR%
Dim LIMPIATEX%
Dim ARRABOR%
Dim FAMOLDE$ ' FABRICACION POR MOLDE PARA MOSTRAR CUANDO SE HACE CLIC CON BOTON DERECHO LA PROYECCION DE FABRICACION

'
Private Sub BOTBUSCA_Click()
   ' Call BULISVEN
End Sub

Private Sub BOTERASE_Click()
   If BOTERASE.Enabled = True Then
     If GRID.Row <= UREPOS& Then
       RELIMI& = GRID.Row
       Screen.MousePointer = 11
       
       For COL% = 1 To GRID.Cols
         If ae$(COL%) = "CAMARA" Then
            GRID.COL = COL% - 1
            ArchivoFoto$ = GRID.TEXT
         End If
       Next COL%
              
       GRID.RemoveItem GRID.Row
       If GRID.Rows < VISIROWS& Then GRID.Rows = VISIROWS&
       UREAR& = ULTREG&(ARCHILOC$)
       For U& = RELIMI& + 1 To UREAR&
         XX$ = REGLEIDO$(ARCHILOC$, U&)
         Call GRAREG(ARCHILOC$, XX$, U& - 1)
         If COLTELIB% > 0 Then
           XX$ = REGLEIDO$(ARCHILIB$, U&)
           Call GRAREG(ARCHILIB$, XX$, U& - 1)
         End If
       Next U&
       If RELIMI& <= UREAR& Then
         Call SETULTREG(ARCHILOC$, UREAR& - 1)
         Call CIERRARCH(ARCHILOC$)
         If COLTELIB% > 0 Then
           Call SETULTREG(ARCHILIB$, UREAR& - 1)
           Call CIERRARCH(ARCHILIB$)
         End If
         UREAR& = UREAR& - 1
       End If
       If RELIMI& > UREAR& Then
         GRID.COL = 0
         Call SETNEXTCOL
       End If
       Call MUETEBOX
       Call Text1_Change
       
       If ArchivoFoto$ <> "" Then
            RutaGuardaCaptura$ = TRIM$(Control("", "RUTARGRA"))
            If RutaGuardaCaptura$ = "" Then RutaGuardaCaptura$ = LUDAT$
            On Error Resume Next
            Call FILEKILL(RutaGuardaCaptura$ & "\" & ArchivoFoto$ & ".JPG")
            On Error GoTo 0
       End If
       
       Screen.MousePointer = 1
     End If
     MODIFI% = 1
  End If
End Sub

Private Sub BOTINSE_Click()
  '
  If APROENTRA% <> 1 Then Exit Sub
  If PUEDINSLIN% < 0 Then Exit Sub    ' NO PUEDEN INSERTARSE RENGLONES
  If GRID.Row < UREPOS& Then
    '
    Screen.MousePointer = 11
    UREAR& = ULTREG&(ARCHILOC$)
    For U& = UREAR& To GRID.Row Step -1
      XX$ = REGLEIDO$(ARCHILOC$, U&)
      Call GRAREG(ARCHILOC$, XX$, U& + 1)
      If COLTELIB% > 0 Then
        XX$ = REGLEIDO$(ARCHILIB$, U&)
        Call GRAREG(ARCHILIB$, XX$, U& + 1)
      End If
    Next U&
    '
    XX$ = REGBLAN$(ARCHILOC$)
    Call GRAREG(ARCHILOC$, XX$, GRID.Row)
    If COLTELIB% > 0 Then
      XX$ = REGBLAN$(ARCHILIB$)
      Call GRAREG(ARCHILIB$, XX$, GRID.Row)
    End If
    '
    GRID.AddItem BLABUFCA$, GRID.Row
    GRID.COL = 0
    INSEROW& = GRID.Row
    Call SETNEXTCOL
    Call MUETEBOX
    MODIFI% = 1
    Screen.MousePointer = 1
 End If
'
End Sub

Sub BOTREC_Click()
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
    Call CIERRARCH("*.*")
    Unload Me
    '
End Sub
'

Private Sub Command1_Click()
   Call Text1_DblClick
End Sub

Private Sub Command3_Click()
   Static UCOLSOR%, UMODSOR$
   If COLTELIB% > 0 Then
      Call MENSERR(24, "Imposible Ordenar esta Tabla")
      Exit Sub
   End If
   '
   If FORVENTB.PROFILE.Caption = "FLEXUSR" Then
      Call MENSERR(24, "Imposible Ordenar esta Tabla")
      Exit Sub
   End If
   '
   Command3.Enabled = False
   If UREPOS& > 0 Then
     TTXX$ = "Elija Columna para Ordenamiento ASC / DESC\"
     TTYY$ = ""
     GRID.Row = 0
     For KKI% = 1 To GRID.Cols
       If GRID.ColWidth(KKI% - 1) > 0 Then
         GRID.COL = KKI% - 1
         TTXX$ = TTXX$ + "\" + TRIM$(GRID.TEXT)
         TTYY$ = TTYY$ + Chr$(KKI%)
       End If
     Next KKI%
   End If
   '
   CCX% = COMALTER%(TTXX$)
   '
   If CCX% > 0 Then
     Screen.MousePointer = 11
     Call CIERRARCH(ARCHILOC$)
     '
     INTERFI$ = "TRAN" + RANDSTRING$(-4)   ' negativo para que sean numeros
     ALOCDA$ = Mid$(ARCHILOC$, 2)
     Call COPYSTRU(ALOCDA$, INTERFI$)
     XY1% = COPYFILE(LUCOM$ + ALOCDA$ + ".WKF", LUDAT$ + INTERFI$ + ".DAT")
     '
     YACANT% = YACARDA%
     YACARDA% = 0
     COLUSOR% = Asc(Mid$(TTYY$, CCX%)) - 1
     YACARDA% = YACANT%
     '
     MODOX$ = "ASC"
     If COLUSOR% = UCOLSOR% Then
       If UMODSOR$ <> "ASC" Then
           MODOX$ = "ASC"
         Else
           MODOX$ = "DESC"
       End If
     End If
     UCOLSOR% = COLUSOR%: UMODSOR$ = MODOX$
     '
     Call FILESORT(INTERFI$, "", COLUSOR%, COLUSOR%, MODOX$)
     XY1% = COPYFILE(LUDAT$ + INTERFI$ + ".DAT", LUCOM$ + ALOCDA$ + ".WKF")
     '
     On Error Resume Next
     Kill (LUDAT$ + INTERFI$ + ".DAT")
     Kill (LUDAT$ + INTERFI$ + ".RFS")
     On Error GoTo 0
     '
     Call BAJALDISCO
     Call CARLIVEN
     Screen.MousePointer = 1
   End If
   Command3.Enabled = True
End Sub

Private Sub Form_Activate()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
End Sub

Private Sub FRAMEBOT1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image1.Visible = False
   MARCOBARRA.Visible = True
End Sub

Private Sub FRAMEBOT2_DblClick()
    Call Image1_DblClick
End Sub

Private Sub Image1_DblClick()
   BSUPERI% = Top
   BIZQUII% = Left
   BALTURI% = Height
   BANCHOI% = Width
   '
   YZT$ = LOADFILE$(PROFIL$)
   TTXX$ = Left$(YZT$, 128)
   TTZZ$ = MKI$(BSUPERI%) + MKI$(BIZQUII%) + MKI$(BALTURI%) + MKI$(BANCHOI%)
   Call REPLA(YZT$, TTZZ$, 73, 8)
   '
   ' GUARDA EL ANCHO DE CADA COLUMNA
   For KKU& = 129 To Len(YZT$) Step 128
      TTXX$ = Mid$(YZT$, KKU&, 128)
      NUCOLUI% = (KKU& - 1) / 128
      If NUCOLUI% <= NUCOLUM% Then
        APIX%(NUCOLUI%) = GRID.ColWidth(NUCOLUI% - 1)
           ACL0% = ACL%(NUCOLUI%) Mod 100
           VCIEN% = 0: If Asc(Mid$(TTXX$, 69, 1)) > 99 Then VCIEN% = 100
        TEPRUEBA.Caption = String$(1 + ACL0%, "0")
        ACL%(NUCOLUI%) = Int(ACL0% * (APIX%(NUCOLUI%) / TEPRUEBA.Width) + 0.5)
        Mid$(TTXX$, 69, 1) = Chr$(VCIEN% + ACL%(NUCOLUI%))
        Mid$(TTXX$, 89, 2) = MKI$(GRID.ColWidth(NUCOLUI% - 1))
        Mid$(YZT$, KKU&, 128) = TTXX$
      End If
   Next KKU&
   '
   Call SAVEFILE(PROFIL$, YZT$)
   Call CARDATVEN
   Call DIBUVEN
   YACARDA% = 0
   Call CARLIVEN
   '
End Sub

Private Sub Form_Load()

'   If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
'      Me.BackColor = &HFCD7B6
'      Me.Frame1.BackColor = &HFCD7B6
'      Me.Frame2.BackColor = &HFCD7B6
'      Me.FRAMEBOT1.BackColor = &HFCD7B6
'      Me.FRAMEBOT2.BackColor = &HFCD7B6
'    End If
   '
   GRID.BackColorFixed = RGB(221, 230, 235)
    
   Call TRANSLABELS(Name)
   '
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Shift = 1 Then
     If Button = 2 Then
       TTXX$ = "Controlador: " + TRIM$(PROFIL$)
       TTXX$ = TTXX$ + "\¿ Desea Refrescar el Controlador ?\\No, Conservarlo\Si, Refrescarlo"
       If COMALTER%(TTXX$) = 2 Then
          Call EXTRACT(PROFIL$)
       End If
     End If
   End If
   '
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image1.Visible = False
   MARCOBARRA.Visible = True
   If X >= Image1.Left Then
     If X <= Image1.Left + Image1.Width Then
       If Y >= Image1.Top Then
         If Y <= Image1.Top + Image1.Height Then
            MARCOBARRA.Visible = False
            Image1.Visible = True
         End If
       End If
     End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
    Cancel = 0
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
   MODITEX% = 0
   '
   RBCL% = RECIBECLICK%
   RECIBECLICK% = (-1)
   GRID.RowSel = GRID.Row
   GRID.ColSel = GRID.COL
   If GRID.Row < 1 Then GRID.Row = 1
   If GRID.Row = 1 Then COLUSOR% = GRID.COL
   If GRID.Row > UREPOS& Then
      GRID.Row = UREPOS&
   End If
   If GRID.Row > ULTREG&(ARCHILOC$) Then
      GRID.COL = 0
   End If
   If PE%(GRID.COL + 1) <> 1 Then
     Text1 = ""      ' agregado EPB el 30/11/2011 por el arrastre de valor a otro campo en MEDITEA - BRecepcion
     If GRID.Row <> INSEROW& Then
       If PRICOLED% < 1 Then
         ROWAC& = GRID.Row
         GRID.Row = UREPOS&
         If GRID.Row <= ULTREG&(ARCHILOC$) Then
           GRID.Row = ROWAC&
           RECIBECLICK% = RBCL%
           Exit Sub
         End If
       End If
       RECIBECLICK% = RBCL%
       Call SETNEXTCOL
     End If
   End If
   RECIBECLICK% = RBCL%
   If RECIBECLICK% = 0 Then Call MUETEBOX
   '
End Sub

Private Sub GRID_DblClick()
   Static UCOLSOR%, UMODSOR$
   If GRID.Row <= 1 Then   ' solo si es el encabezado
     If COLTELIB% > 0 Then
        Call MENSERR(24, "Imposible Ordenar esta Tabla")
        Exit Sub
     End If
     '
     If FORVENTB.PROFILE.Caption = "FLEXUSR" Then
        Call MENSERR(24, "Imposible Ordenar esta Tabla")
        Exit Sub
     End If
     '
     Screen.MousePointer = 11
     Call CIERRARCH(ARCHILOC$)
     INTERFI$ = "TRAN" + RANDSTRING$(-4)   ' negativo para que sean numeros
     ALOCDA$ = Mid$(ARCHILOC$, 2)
     Call COPYSTRU(ALOCDA$, INTERFI$)
     XY1% = COPYFILE(LUCOM$ + ALOCDA$ + ".WKF", LUDAT$ + INTERFI$ + ".DAT")
     '
     YACANT% = YACARDA%
     YACARDA% = 0
     COLUSOR% = GRID.COL
     YACARDA% = YACANT%
     '
     MODOX$ = "ASC"
     If COLUSOR% = UCOLSOR% Then
       If UMODSOR$ <> "ASC" Then
           MODOX$ = "ASC"
         Else
           MODOX$ = "DESC"
       End If
     End If
     UCOLSOR% = COLUSOR%: UMODSOR$ = MODOX$
     '
     Call FILESORT(INTERFI$, "", COLUSOR%, COLUSOR%, MODOX$)
     XY1% = COPYFILE(LUDAT$ + INTERFI$ + ".DAT", LUCOM$ + ALOCDA$ + ".WKF")
     '
     On Error Resume Next
     Kill (LUDAT$ + INTERFI$ + ".DAT")
     Kill (LUDAT$ + INTERFI$ + ".RFS")
     On Error GoTo 0
     '
     Call BAJALDISCO
     Call CARLIVEN
     Screen.MousePointer = 1
   End If
End Sub

Private Sub GRID_EnterCell()
   If YACARDA% < 1 Then Exit Sub
End Sub


Private Sub GRID_LeaveCell()
   '
   If LIMPIAGR% > 0 Then Exit Sub
   If YACARDA% < 1 Then Exit Sub
   If RECIBECLICK% < 0 Then Exit Sub
   If ACL%(GRID.COL + 1) < 1 Then Exit Sub
   AAA = UREPOS&
   If GRID.Row < UREPOS& Then
     If PE%(GRID.COL + 1) <> 1 Then Exit Sub
   End If
   If APROENTRA% < 1 Then
     Exit Sub
   End If
   '
End Sub

Private Sub GRID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ' linea agregada el 10/12/02 para corregir que al cambiar
   ' de renglon arrastraba el valor del texto anterior de
   ' la caja azul de texto
   'If GRID.Row >= UREPOS& Then Text1 = ""
   If GRID.Row >= UREPOS& Then LIMPIATEX% = 1
   '
   If Button = 2 Then
     For KU% = 1 To NUCOLUM%
       If TVL%(KU%) = 7 Then
       GRID.COL = KU% - 1
         ARGUBUS$ = GRID.TEXT
         Call MUESTRADATCOD(ARGUBUS$)
       End If
     Next KU%
   End If
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

Private Sub Picture1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
   Picture1.MousePointer = 7
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 1
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If ARRABOR% = 1 Then
      Height = Picture1.Top + Y + 450
      Picture1.Top = Height - 450
   End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 0
End Sub

Private Sub PROFILE_Change()
   '
   TIFORMED% = 1
   PX% = InStr(PROFILE, "/")
   If PX% < 1 Then PX% = 1 + Len(PROFILE)
   '
   PROFIL$ = Left$(PROFILE, PX% - 1) + ".PRF"
   ATRIBU$ = Mid$(PROFILE, PX%)
   'If VERDEMO% = 1 Then
   '  PROFIL$ = "\FX_PROFILS\" + PROFIL$
   'End If
   '
   Call CARDATVEN
   Call DIBUVEN
   SMP = Screen.MousePointer
   Screen.MousePointer = 1

End Sub

Sub CARDATVEN()
     '
     ARZZ$ = UCase$(PROFIL$)
     Call GRATABCONTRO(ARZZ$)
     '
5800 GRABACT% = 0
     BSUW& = 0: BIZW& = 0
     COLTELIB% = (-1)
     PUEDINSLIN% = 0
     '
     If EXISTE%(PROFIL$) < 1 Then
       Call EXTRAPROF
     End If
     '
     GRID.Clear
     '
10   N15% = FILEOPEN%(PROFIL$, 10, 128)
     '
     'DETERMINA CANTIDAD DE COLUMNAS
     For KKKI% = 1 To 64
        Get #N15%, KKKI% + 1, YYY$
        If YYY$ = String$(128, 0) Then GoTo 5802
     Next KKKI%
5802 GRID.Clear
     GRID.Cols = KKKI% - 1
     '
     Get #N15%, 1, YYY$
     ARCHITRA$ = TRIM$(Mid$(YYY$, 57, 8))
     If ARCHITRA$ = "" Then
         TTXX$ = "Aparente Daño no Recuperable\en Controlador '" + PROFIL$ + "'.\   \Instale Versión Actualizada de ese Controlador."
         Call MENSERR(24, TTXX$)
         Call FINAL("")
     End If
     '
     ARCHILOC$ = "!" + ARCHITRA$
     ARCHILIB$ = ""
     PROCE$ = TRIM$(Mid$(YYY$, 65, 8))
     LOR1% = Asc(Mid$(YYY$, 106, 1)): If LOR1% > 16 Then LOR1% = 0
     LOR% = 256 * LOR1% + Asc(Mid$(YYY$, 107))
     CAMAX& = CVI(Mid$(YYY$, 108))
     PEOF% = 1 + Asc(Mid$(YYY$, 110))
     EOFS$ = Mid$(YYY$, 111, 16)
5803 Caption = Left$(YYY$, 48)
     sali$ = Mid$(YYY$, 49, 8)
     COD$ = Mid$(YYY$, 81)
     RGI& = Asc(Mid$(COD$, 1))
     RGF& = Asc(Mid$(COD$, 2))
     LEFTPOS% = Asc(Mid$(COD$, 3))
5804 BLP% = Asc(Mid$(COD$, 4))
     VIV% = Asc(Mid$(COD$, 5))
     APPI% = Asc(Mid$(COD$, 6))
     IDL% = Asc(Mid$(COD$, 7))
          If InStr(UCase$(ATRIBU$), "/POINDEL=NO") > 0 Then IDL% = 0
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
     DPOSI$ = Mid$(YYY$, 73, 8)
        If DPOSI$ = Space$(8) Then DPOSI$ = String$(8, 0)
     BSUPER = CVI(Mid$(DPOSI$, 1, 2))
     BIZQUI = CVI(Mid$(DPOSI$, 3, 2))
     BALTUR = CVI(Mid$(DPOSI$, 5, 2))
     BANCHO = CVI(Mid$(DPOSI$, 7, 2))
     '
     YADECPRE% = 0
     '
5805 RG% = RG% + 1
     If RG% > LOF(N15%) / 128 Or RG% > 33 Then GoTo 5820
5806 Get #N15%, RG%, YYY$
     If YYY$ = String$(128, 0) Then GoTo 5820
     '
5807 NUCOLUM% = NUCOLUM% + 1
     CRF%(NUCOLUM%) = RG% - 1       ' NUMERO DE COLUMNA REAL DEL ARCHIVO
     TVL%(NUCOLUM%) = Asc(Mid$(YYY$, 33))
     LDL%(NUCOLUM%) = Asc(Mid$(YYY$, 34))
         If RG% = 2 Then TVL%(NUCOLUM%) = 3: LDL%(NUCOLUM%) = 0
         ' AQUI PONE LA CANTIDAD DE DECIMALES
         If Left$(App.EXEName, 6) = "LISPRE" Or Left$(App.EXEName, 6) = "VENGES" Or Left$(App.EXEName, 6) = "PEDIDO" Then
           If Left$(ARZZ$, 8) = "LIPREVEN" Or Left$(ARZZ$, 3) = "FAC" Or Left$(ARZZ$, 6) = "CARDET" Then
             If LDL%(NUCOLUM%) >= 2 Then       ' CANTIDAD DE DECIMALES
               If YADECPRE% = 0 Then
                 CADECI% = 0
                 On Error Resume Next
                 CADECI% = XVALO(LISPREACC.Text3)
                 On Error GoTo 0
                 If CADECI% < 1 Then
                    NLIS% = XVALO(Mid$(ARCHIPRE$, 6, 3))
                    If NLIS% < 1 Then NLIS% = LIPRACTI%
                    CADECI% = DECILISTA%(NLIS%)
                 End If
                 If CADECI% > 0 Then
                    LDL%(NUCOLUM%) = CADECI%
                 End If
                 YADECPRE% = 1
               End If
             End If
           End If
         End If
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
     ATELIB% = 0
     ACL%(NUCOLUM%) = Asc(Mid$(YYY$, 69))
     If ACL%(NUCOLUM%) > 100 Then
       ATELIB% = 1
       ACL%(NUCOLUM%) = ACL(NUCOLUM%) - 100
     End If
     APIX%(NUCOLUM%) = CVI(Mid$(YYY$, 89, 2))
     LK%(NUCOLUM%) = Asc(Mid$(YYY$, 70))
     PINSE%(NUCOLUM%) = 1 + Len(BLABUFCA$)
     BLABUFCA$ = BLABUFCA$ + String$(1 + ACL%(NUCOLUM%), " ") + Chr$(9)
     If Asc(Mid$(YYY$, 69)) > 0 Then
       ECL$(NUCOLUM%) = REPLACEN$(TRIM$(Mid$(YYY$, 37, 32)))
       TEPRUEBA.Caption = String$(1 + ACL%(NUCOLUM%), "0")
       GRID.ColWidth(NUCOLUM% - 1) = TEPRUEBA.Width
       If APIX%(NUCOLUM%) > 0 Then
           GRID.ColWidth(NUCOLUM% - 1) = APIX%(NUCOLUM%)
           'ACL%(NUCOLUM%) = ACL%(NUCOLUM%) * (APIX%(NUCOLUM%) / TEPRUEBA.Width)
       End If
       GRID.COL = NUCOLUM% - 1
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
       GRID.Clip = TRIM$(ECL$(NUCOLUM%))
       MG%(NUCOLUM%) = Asc(Mid$(YYY$, 71))
       PE%(NUCOLUM%) = Asc(Mid$(YYY$, 73))
          If RG% = 2 Then PE%(NUCOLUM%) = 2
       If PE%(NUCOLUM%) = 1 Then
          GRABACT% = 1
       End If
       If PE%(NUCOLUM%) = 2 Then
          PUEDINSLIN% = (-1)  ' NO PUEDEN INSERTARSE RENGLONES
       End If
       JT%(NUCOLUM%) = Asc(Mid$(YYY$, 72))
       ae$(NUCOLUM%) = TRIM$(Mid$(YYY$, 74, 8))
       CE%(NUCOLUM%) = Asc(Mid$(YYY$, 82))
       If TVL%(NUCOLUM%) = 5 Then
         ae$(NUCOLUM%) = "SELFECHA"
       End If
       If COLTELIB% < 0 Then
         If TVL%(NUCOLUM%) = 6 Then ' SOLO SI TEXTO
           If PE%(NUCOLUM%) = 1 Then ' SOLO SI EDITABLE
             If ATELIB% > 0 Then ' MARCA DE RENGLONES MULTIPLES
               COLTELIB% = NUCOLUM%
             End If
           End If
         End If
       End If
       '
    End If
    '
    GoTo 5805
5817 '
5820 If COLTELIB% > 0 Then
       LORELIB& = RECLEN("ARTELIB")
       ARCHILIB$ = "1" + Left$(ARCHITRA$, 7)
       Call COPYSTRU("ARTELIB", ARCHILIB$)
       ARCHILIB$ = "!" + ARCHILIB$
     End If
     '
     ATOT% = NUCOLUM% - 1 - NCLF%
     BLABUFCA$ = Left$(BLABUFCA$, Len(BLABUFCA$) - 1) + Chr$(13)
     GRID.Cols = NUCOLUM%
     CLII% = LEFTPOS% + 2
     For KCL% = 1 To NUCOLUM%
        lC%(KCL%) = ACL%(KCL%)
        If Left$(ECL$(KCL%), 1) <> "@" Then
            lC%(KCL%) = ACL%(KCL%)
            ATOT% = ATOT% + ACL%(KCL%)
            cl%(KCL%) = CLII%
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
         If TRIM$(XX$) <> "" Then
             Caption = XX$
         End If
     End If
     '
5705 PX% = InStr(UCase$(ATRIBU$), "ARCHIVO=")
     If PX% > 0 Then
         XX$ = Mid$(ATRIBU$, PX% + 8)
         PY% = InStr(XX$, "/"): If PY% < 1 Then PY% = 1 + Len(XX$)
         XX$ = Left$(XX$, PY% - 1)
         If TRIM$(XX$) <> "" Then ARCHITRA$ = TRIM$(XX$)
     End If
     '
5710 PX% = InStr(UCase$(ATRIBU$), "BORDESUP=")
     If PX% > 0 Then
         XX$ = Mid$(ATRIBU$, PX% + 9)
         PY% = InStr(XX$, "/"): If PY% < 1 Then PY% = 1 + Len(XX$)
         XX$ = Left$(XX$, PY% - 1)
         If Left$(XX$, 1) = "W" Then
              BSUW& = XVALO(Mid$(XX$, 2))
            Else
              RGI1& = XVALO(XX$)
              If RGI1& > 0 Then If RGI1& <= 22 Then RGI& = RGI1&
         End If
     End If
     '
5720 PX% = InStr(UCase$(ATRIBU$), "BORDEINF=")
     If PX% > 0 Then
         XX$ = Mid$(ATRIBU$, PX% + 9)
         PY% = InStr(XX$, "/"): If PY% < 1 Then PY% = 1 + Len(XX$)
         XX$ = Left$(XX$, PY% - 1)
         RGF1& = XVALO(XX$)
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
               BIZW& = XVALO(Mid$(XX$, 2))
             Else
               LEFTPOX% = XVALO(XX$)
               If LEFTPOX% > 0 Then If LEFTPOX% <= 76 Then LEFTPOS% = LEFTPOX%
         End If
         '
         CLII% = LEFTPOS% + 2
         For KCL% = 1 To NUCOLUM%
            lC%(KCL%) = ACL%(KCL%)
            If Left$(ECL$(KCL%), 1) <> "@" Then
               lC%(KCL%) = ACL%(KCL%)
               ATOT% = ATOT% + ACL%(KCL%)
               cl%(KCL%) = CLII%
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
        ARGATRI$ = "ARCHECO(" + TRIM$(str$(JK%)) + ")="
        PX% = InStr(UCase$(ATRIBU$), ARGATRI$)
        If PX% > 0 Then
           XX$ = Mid$(ATRIBU$, PX% + Len(ARGATRI$))
           PY% = InStr(XX$, "/"): If PY% < 1 Then PY% = 1 + Len(XX$)
           XX$ = Left$(XX$, PY% - 1)
           If TRIM$(XX$) <> "" Then ae$(UU%) = TRIM$(XX$)
        End If
        ' PERMITE DEFINIR POSIBILIDAD DE EDICION DE COLUMNAS
        ARGATRI$ = "POSEDIT(" + TRIM$(str$(JK%)) + ")="
        PX% = InStr(UCase$(ATRIBU$), ARGATRI$)
        If PX% > 0 Then
           XX$ = Mid$(ATRIBU$, PX% + Len(ARGATRI$))
           PY% = InStr(XX$, "/"): If PY% < 1 Then PY% = 1 + Len(XX$)
           XX$ = Left$(XX$, PY% - 1)
           If TRIM$(XX$) <> "" Then
              On Error Resume Next
              PE%(UU%) = XVALO(XX$)
              On Error GoTo 0
           End If
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
   COLAC% = GRID.COL + 1
   Text1.ToolTipText = ""
   
   ' SE EXCLUYE DEL ECO A LA CAPTURA DE IMAGEN POR CAMARA
   If ae$(COLAC%) <> "" And ae$(COLAC%) <> "CAMARA" And ae$(COLAC%) <> "BALANZA" Then
     ARCHEZ$ = TRIM$(UCase$(ae$(COLAC%)))
     If ARCHEZ$ = "UNICOM" Then ARCHEZ$ = "UNIMED"
     If ARCHEZ$ = "MASTER" Or ARCHEZ$ = "OCOMEMI" Or ARCHEZ$ = "ECOMAST" Or Left$(ARCHEZ$, 5) = "LIPRE" Then
         If COLTELIB% < 1 Then
             TTTXT$ = DESCRIT$(CODINT%(Text1))
           Else
             TTTXT$ = DESCRIT0$(CODINT%(Text1))
         End If
       Else
         If Mid$(ARCHEZ$, 6, 3) = "###" Then
           NORDEX$ = TRIM$(str$(GRID.Row))
           While Len(NORDEX$) < 3: NORDEX$ = "0" + NORDEX$: Wend
           ARCHEZ$ = Left$(ARCHEZ$, 5) + NORDEX$
         End If
         TTTXT$ = TRIM$(ECOARCH$(ARCHEZ$, Text1))
     End If
     Text1.ToolTipText = TTTXT$
   End If
   '
   If MODITEX% <> 1 Then Exit Sub
   Call MUESTRAECOS
   '
End Sub
'
Private Sub Text11_Change()

End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
     COLAC% = GRID.COL + 1
     If ae$(COLAC%) <> "" Then
       ARCHEZ$ = TRIM$(UCase$(ae$(COLAC%)))
       ARGUBUS$ = TRIM$(Text1)
       If ARCHEZ$ = "MASTER" Or ARCHEZ$ = "ECOMAST" Or ARCHEZ$ = "OCOMEMI" Or Left$(ARCHEZ$, 5) = "LIPRE" Then
           Text1.Enabled = False
           Call MUESTRADATCOD(ARGUBUS$)
           Text1.Enabled = True
           Exit Sub
         ElseIf ARCHEZ$ = "PROVED1" Or ARCHEZ$ = "ACREDOR" Or ARCHEZ$ = "PROVEPLA" Then
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
   COLAC% = GRID.COL + 1
   
   ' PROCESA EL ECO DEL CAMPO SETEADO EN EL PRF COMO CAMARA = CAMARA WEB
   If ae$(COLAC%) = "CAMARA" Then
        On Error GoTo Err_FalloCapturaFoto
        RutaGuardaCaptura$ = TRIM$(Control("", "RUTARGRA"))
        If RutaGuardaCaptura$ = "" Then RutaGuardaCaptura$ = LUDAT$
        ' ESTABLECE UN NUMERO DE 4 DIGITOS ALEATORIOS PARA EL NOMBRE DEL ARCHIVO
5       ArchivoFoto$ = CStr(RANDSTRING$(-4))
        ' VERIFICA QUE EL NUMERO NO ESTE OCUPADO
        For F% = 1 To GRID.Rows - 1
            If TRIM$(GRID.TextMatrix(F%, COLAC%)) = ArchivoFoto$ Then
                GoTo 5
            End If
        Next F%
        WEBCAM.Text1 = ArchivoFoto$
        WEBCAM.Text2 = RutaGuardaCaptura$
        WEBCAM.Show 1
        If EXISTE%(RutaGuardaCaptura$ & "\" & ArchivoFoto$ & ".JPG") > 0 Then
            GRID.TEXT = ArchivoFoto$: Text1 = ArchivoFoto$
            Call GRAVACOLRO
        End If
        On Error GoTo 0
        Exit Sub
   End If
   
   ' PROCESA EL ECO DEL CAMPO SETEADO EN EL PRF COMO CAMARA = CAMARA WEB
   If ae$(COLAC%) = "BALANZA" Then
        BALANZA.CodigoProducto = GRID.TextMatrix(GRID.Row, 5)
        BALANZA.TXTDESCRIPCION = GRID.TextMatrix(GRID.Row, 6)
        BALANZA.Show 1
        
        If BALANZA.Respuesta = "APROBO" Then
            GRID.TEXT = BALANZA.TXTPESO
            Text1 = BALANZA.TXTPESO
            Call Text1_KeyPress(13)
        End If
        Unload BALANZA
        
        Exit Sub
   End If
   
   If ae$(COLAC%) <> "" Then
     FORSELEC.SINI = str$(Top + (Height - ScaleHeight) + Text1.Top + Text1.Height)
     FORSELEC.LINI = str$(Left + Text1.Left)
     FORSELEC.ABUSQUE = TRIM$(Text1)
     '
     ARCHEK$ = ae$(COLAC%)
     If ARCHEK$ = "DESCOMBI" Then Exit Sub
     '
     If ARCHEK$ = "UNICOM" Then ARCHEK$ = "UNIMED"
     '
     If Mid$(ARCHEK$, 6, 3) = "###" Then
       NORDEX$ = TRIM$(str$(GRID.Row))
       While Len(NORDEX$) < 3: NORDEX$ = "0" + NORDEX$: Wend
       ARCHEK$ = Left$(ae$(COLAC%), 5) + NORDEX$
     End If
     '
     If ARCHEK$ = "LOTECOD" Then
       For KKI% = COLAC% To 1 Step -1
         If TVL%(KKI%) = 7 Then
           GRID.COL = KKI% - 1
           CIXI% = CODINT%(GRID.TEXT)
           GRID.COL = COLAC%
           VDEVU$ = TRIM$(Left$(LOTESEL$(CIXI%), 16))
           GoTo 10
         End If
       Next KKI%
       Exit Sub
     End If
     '
     DELISECO$ = ""
     If Left$(UCase$(ARCHEK$), 5) = "LIPRE" Then
       DELISECO$ = DENOLISTA$(XVALO(Mid$(ARCHEK$, 6, 3)))
     End If
     Call SELECHO(ARCHEK$ + "/" + DELISECO$)
     VDEVU$ = VALACT1$(ARCHEK$)
     '
10   If VDEVU$ <> "" Then
       If JT%(GRID.COL + 1) = 2 Then
         While Len(VDEVU$) < ACL%(GRID.COL + 1)
           VDEVU$ = " " + VDEVU$
         Wend
       End If
       If TRIM$(VDEVU$) <> TRIM$(Text1) Then
         MODITEX% = 1
       End If
       Text1 = VDEVU$
       If VDEVU$ <> GRID.TEXT Then CAMBIOVALGRILLA% = 1
       GRID.TEXT = VDEVU$
       Call GRAVACOLRO
       Call MUESTRAECOS
       RECIBECLICK% = (-1)
       Call SETNEXTCOL
       RECIBECLICK% = 0
     End If
   End If
   
99 Exit Sub
   
   
Err_FalloCapturaFoto:
      
   GoTo 99
   
End Sub

Private Sub Text1_GotFocus()
  Text1.SelStart = 0
  Text1.SelLength = Len(Text1)
  TEXTOINI$ = TRIM$(Text1.TEXT)
  GRID.COL = COLTEBO% - 1
  GRID.Row = ROWTEBO%
  Text1.Refresh
  ' BLOQUEA TEXT1 CUANDO CORRESPONDE A LA CAPTURA DE UNA IMAGEN VIA WEB CAM.
  ' EN ESTOS CASOS, EL TEXT1 TIENE UN NUMERO QUE REPRESENTA EL NOMBRE DEL ARCHIVO
  ' QUE CONTIENE LA IMAGEN Y NO PUEDE SER MODIFICADO PORQUE SINO SE PERDERIA
  ' EL VINCULO CON EL REGISTRO AL CUAL CORRESPONDE.
  Text1.Locked = False
  COLAC% = GRID.COL + 1
  If ae$(COLAC%) = "CAMARA" Or ae$(COLAC%) = "BALANZA" Then Text1.Locked = True
End Sub

Private Sub Text11_GotFocus()
  Text11.SelStart = 1 + Len(Text11)
  Text11.SelLength = Len(Text11)
  TEXTOINI$ = Text11.TEXT
  GRID.COL = COLTEBO% - 1
  GRID.Row = ROWTEBO%
  Text11.Refresh
End Sub

Function APROENTRA%()
   '
   RECIBECLICK% = -1
   APROENTRA% = 1
   '
   If GRID.COL + 1 = COLTELIB% Then GoTo 11
   '
   If Text1 <> "" And TRIM$(Text1) <> TEXTOINI$ Then 'TRIM$(GRID.TEXT) Then
     '
     COLAC% = GRID.COL + 1
   
     VDEVU$ = Text1
     VABIN$ = CBIN$(VDEVU$, TVL%(COLAC%), LF%(COLAC%))
     UCA_VABIN$ = VABIN$
     If TVL%(COLAC%) = 6 Then UCA_VABIN$ = UCase$(VABIN$)
     If ae$(COLAC%) <> "" Then
       '
       ARCHEK$ = ae$(COLAC%)
       If ARCHEK$ = "UNICOM" Then ARCHEK$ = "UNIMED"
       '
       If Mid$(ARCHEK$, 6, 3) = "###" Then
         NORDEX$ = TRIM$(str$(GRID.Row))
         While Len(NORDEX$) < 3: NORDEX$ = "0" + NORDEX$: Wend
         ARCHEK$ = Left$(ae$(COLAC%), 5) + NORDEX$
       End If
       '
       If ae$(COLAC%) = "FECHA" Or ae$(COLAC%) = "SELFECHA" Then
           If CVI(VABIN$) = 0 Then
             Call MENSERR(24, "Ingreso no Válido")
             MODITEX% = 1
             GoTo 90
           End If
         ' SE EXCLUYE DEL ECO A LA CAPTURA DE IMAGEN POR CAMARA
         ElseIf ECOARCH(ARCHEK$, UCA_VABIN$) = "" And ae$(COLAC%) <> "CAMARA" And ae$(COLAC%) <> "BALANZA" Then
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
             If Left$(GRID.TEXT, ACL%(COLAC%)) = VANUTE$ Then
               RECIBECLICK% = 0
               Call SETNEXTCOL
               GoTo 90
             End If
           End If
         Next U&
         RECIBECLICK% = 0
       End If
       GRID.Row = ROWAC%
       GRID.COL = COLAC% - 1
       INSEROW& = RINSERO&
       If VANUTE$ <> GRID.TEXT Then CAMBIOVALGRILLA% = 1
       GRID.TEXT = VANUTE$
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
   Exit Function
   '
   '
   ' LO QUE SIGUE UNICAMENTE PARA LA COLUMN A DE TEXTO LIBRE
11 If TRIM$(Text11) <> TEXTOINI$ Then 'TRIM$(GRID.TEXT) Then
     '
     COLAC% = GRID.COL + 1
   
     VDEVU$ = Text11
     VABIN$ = VDEVU$
     '
20   ROWAC% = GRID.Row
     If ROWAC% > 0 Then
       If ROWAC% <= ULTREG&(ARCHILOC$) Then
         VANUTE$ = VDEVU$
           'PPXX% = InStr(VANUTE$, Chr$(13))
           'If PPXX% > 0 Then VANUTE$ = Left$(VANUTE$, PPXX% - 1)
         GRID.Row = ROWAC%
         GRID.COL = COLAC% - 1
         INSEROW& = RINSERO&
         If VANUTE$ <> GRID.TEXT Then CAMBIOVALGRILLA% = 1
         GRID.TEXT = VANUTE$
         Call GRAVACOLRO
         Call MUESTRAECOS
       End If
     End If
     '
     APROENTRA% = 1
     RECIBECLICK% = 0
     Exit Function
     '
91   APROENTRA% = 0
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
         GRID.COL = 0
       End If
       GRID.Row = GRID.Row + 1
     End If
     If GRID.Row > ULTREG&(ARCHILOC$) Then
       GRID.COL = 0
     End If
     If PE%(GRID.COL + 1) <> 1 Then Call SETNEXTCOL
     Call MUETEBOX
   ElseIf KeyCode = 38 Then
     If APROENTRA% <> 1 Then Exit Sub
     If GRID.Row > 1 Then
       GRID.Row = GRID.Row - 1
       If PE%(GRID.COL + 1) <> 1 Then Call SETNEXTCOL
     End If
     If GRID.TopRow > GRID.Row Then
       GRID.TopRow = GRID.Row
     End If
     Call MUETEBOX
   ElseIf KeyCode = 39 Then
     If Text1.SelStart >= Len(RTrim(Text1)) Or Text1.SelText = Text1 Then
       If APROENTRA% <> 1 Then Exit Sub
       If GRID.Row > ULTREG&(ARCHILOC$) Then GRID.COL = 0
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
     COLAC% = GRID.COL + 1
     Call Text1_DblClick
   ElseIf KeyCode = 36 Then    ' home
     If APROENTRA% <> 1 Then Exit Sub
     If GRID.COL > PRICOLED% - 1 Then
         GRID.COL = PRICOLED% - 1
       Else
         GRID.Row = 1
         If GRID.TopRow > GRID.Row Then
           GRID.TopRow = GRID.Row
         End If
     End If
     Call MUETEBOX
   ElseIf KeyCode = 35 Then    ' END
     If APROENTRA% <> 1 Then Exit Sub
     If GRID.COL < ULTCOLED% - 1 Then
         GRID.COL = ULTCOLED% - 1
       Else
         GRID.Row = UREPOS&
         If GRID.Row > ULTREG&(ARCHILOC$) Then
           GRID.COL = 0
           Call SETNEXTCOL
         End If
         If GRID.TopRow < GRID.Row - VISIROWS& + 3 Then
           GRID.TopRow = GRID.Row - VISIROWS& + 3
         End If
     End If
     Call MUETEBOX
   ElseIf KeyCode = 8 Then    ' backspace
     If Text1.SelText = Text1.TEXT Then
       If APROENTRA% <> 1 Then Exit Sub
       Call SETPREVCOL
       Call MUETEBOX
     End If
   ElseIf KeyCode = 27 Then    ' ESCAPE
     If APROENTRA% <> 1 Then Exit Sub
     On Error Resume Next
     Text1.Visible = False
     Command1.Visible = False
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
     If PE%(GRID.COL + 1) <> 1 Then Call SETNEXTCOL
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
     If PE%(GRID.COL + 1) <> 1 Then Call SETNEXTCOL
     If GRID.Row > ULTREG&(ARCHILOC$) Then
       GRID.COL = 0
       Call SETNEXTCOL
     End If
     Call MUETEBOX
   ElseIf KeyCode = 46 Then
     If Text1.SelLength = Len(Text1) Then
       KeyCode = 0
       Call BOTERASE_Click
     End If
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
   If Text1.SelLength >= MAXCARS% Then
      Text1 = ""
      Exit Sub
   End If
   '
   If KeyAscii% >= 48 Then
     If Len(TRIM$(Text1)) >= MAXCARS% Then '  - 1 Then
       KeyAscii% = 0    ' 13
     End If
   End If
   '
   If KeyAscii = 13 Then
     If APROENTRA% = 1 Then
       If GRID.Row > ULTREG&(ARCHILOC$) Then GRID.COL = 0
       RECIBECLICK% = (-1)
       Call SETNEXTCOL
       RECIBECLICK% = 0
     End If
   End If
   '
   If Len(TRIM$(Text1)) > MAXCARS% Then
      MsgBox ("Texto Ingresado Excede Tamaño del Campo")
      Text1 = Left$(TRIM$(Text1), MAXCARS%)
      Text1.SelStart = 1 + Len(Text1)
   End If
End Sub

Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 39 Then
     If Text11.SelStart >= Len(RTrim(Text11)) Or Text11.SelText = Text1 Then
       If APROENTRA% <> 1 Then Exit Sub
       If GRID.Row > ULTREG&(ARCHILOC$) Then GRID.COL = 0
       Call SETNEXTCOL
     End If
   ElseIf KeyCode = 37 Then
     If Text11.SelText <> Text11 Or Text11 = "" Then
       If Text11.SelStart < 1 Then
         If APROENTRA% <> 1 Then Exit Sub
         Call SETPREVCOL
       End If
     End If
   End If
End Sub

Sub GRAVACOLRO()
   '
   COLAC% = GRID.COL + 1
   REGAC& = GRID.Row
   If REGAC& > 0 Then
     If REGAC& <= UREPOS& Then
       VDEVU$ = GRID.TEXT
       VABIN$ = CBIN$(VDEVU$, TVL%(COLAC%), LF%(COLAC%))
       If REGAC& <= ULTREG&(ARCHILOC$) Then
           XX$ = REGLEIDO$(ARCHILOC$, REGAC&)
         Else
           XX$ = REGBLAN$(ARCHILOC$)
       End If
       Call REPLA(XX$, VABIN$, PR%(COLAC%), LF%(COLAC%))
       Call GRAREG(ARCHILOC$, XX$, REGAC&)
       If COLAC% = COLTELIB% Then
         If REGAC& <= ULTREG&(ARCHILOC$) Then
             XX$ = REGLEIDO$(ARCHILIB$, REGAC&)
           Else
             XX$ = REGBLAN$(ARCHILIB$)
         End If
         Call REPLA(XX$, VDEVU$, 1, 2048)
         Call GRAREG(ARCHILIB$, XX$, REGAC&)
       End If
     End If
   End If
   '
End Sub
'
Sub SETNEXTCOL()
   '
   If YACARDA% < 1 Then Exit Sub
   COLVIE% = GRID.COL
   ROWVIE% = GRID.Row
   '
5  COLNUE = GRID.COL + 1
   ROWNUE = GRID.Row
   UREPOX& = UREPOS&
   '
10 If ROWNUE > UREPOX& Then If APPI% < 1 Then GoTo 99
   If COLNUE >= NUCOLUM% Then
     COLNUE = 0
     If ROWNUE < UREPOX& Then
         ROWNUE = ROWNUE + 1
       Else
         If APPI% < 1 Then GoTo 99
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
   ' Text1 = ""         ' agregado EPB el 14/11/2011 por el arrastre de valor a otro campo en MEDITEA - BRecepcion
   GRID.COL = COLNUE
   GRID.Row = ROWNUE
   If GRID.CellWidth < 0 Then GoTo 5
   Call MUETEBOX
   Exit Sub
   '
99 Exit Sub
End Sub

Sub SETPREVCOL()
   '
   COLNUE = GRID.COL - 1
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
   GRID.COL = COLNUE
   Call MUETEBOX
   '
End Sub

Private Sub Text2_gotfocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Timer1_Timer()
   Static ACOLANT
   '
   If Visible = False Then Exit Sub
   '
   ACOLTOT = 380 ' 350
   For KKI% = 1 To GRID.Cols
     ACOLTOT = ACOLTOT + GRID.ColWidth(KKI% - 1)
   Next KKI%
   If ACOLTOT <> ACOLANT Or YACARDA% = 0 Then
     Call FRESHALL
     WD1 = ACOLTOT + 750 + MAINT + MADER
     'If WD1 > 12000 - Left Then WD1 = 12000 - Left
     If WD1 > Screen.Width - Left Then WD1 = Screen.Width - Left
     Width = WD1
     FRAMEBOT1.Left = Width - 750 - MADER - 70
     FRAMEBOT2.Left = Width - 750 - MADER - 70
     Frame2.Left = Width - 750 - MADER - 70
     'MARCOBARRA.Left = Width - 750 - MADER
     MARCOBARRA.Left = FRAMEBOT1.Left + ((FRAMEBOT1.Width - MARCOBARRA.Width) / 2)
     Image1.Left = MARCOBARRA.Left + 120
     ACOLANT = ACOLTOT
     GRID.Width = Width - 750 - MAINT - MADER
     Call FRESHALL
     '
     If PRICOLED% < 1 Then
       If YACARDA% > 0 Then
         GRID.Row = UREPOS&
       End If
     End If
     Call FRESHALL
     Call MUETEBOX
     Call FRESHALL
   End If
   '
   If YACARDA% = 0 Then
     If Visible = True Then
       Call CARLIVEN
       YACARDA% = 1
       GRID.Row = 1
       GRID.COL = 0
       If PRICOLED% < 1 Then
         GRID.Row = UREPOS&
         If GRID.Row <= ULTREG&(ARCHILOC$) Then
           GoTo 10
         End If
       End If
       Call SETNEXTCOL
       '
10     Call MUETEBOX
     End If
   End If
   '
End Sub

Sub DIBUVEN()     ' Rutina de dibujo pantalla tabulada
   '
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
  TEPRUANCH.Font = "Fixedsys"
  GRID.Font = "Fixedsys"
  Text1.Font = "Fixedsys"
  Text11.Font = "Fixedsys"
  Text2.Font = "Fixedsys"
  TEPRUEBA.Font = "Fixedsys"
  TEPRUTEX.Font = "Fixedsys"
  '
If ATOT% > 85 Then
  TEPRUANCH.Font = "Lucida Console"
  TEPRUANCH.FontSize = 9
  TEPRUANCH.FontBold = False
  '
  GRID.Font = "Lucida Console"
  GRID.FontSize = 9
  GRID.FontBold = False
  '
  Text1.Font = "Lucida Console"
  Text1.FontSize = 9
  Text1.FontBold = False
  '
  Text11.Font = "Lucida Console"
  Text11.FontSize = 9
  Text11.FontBold = False
  '
  Text2.Font = "Lucida Console"
  Text2.FontSize = 9
  Text2.FontBold = False
  '
  TEPRUEBA.Font = "Lucida Console"
  TEPRUEBA.FontSize = 9
  TEPRUEBA.FontBold = False
  '
  TEPRUTEX.Font = "Lucida Console"
  TEPRUTEX.FontSize = 9
  TEPRUTEX.FontBold = False
  '
End If
   ' para aparear con vieja version
   Height = Screen.Height / 25 * (RGF& - RGI& - 1)
   LIPRUEBA.Height = ScaleHeight - 768
   Height = LIPRUEBA.Height + 768
   Picture1.Top = Height - 450
   '
   ' para aparear con vieja version
   Top = Screen.Height / 25 * (RGI&)
   If BSUW& <> 0 Then Top = BSUW&
   HBAS = Screen.Height / 25 * (RGF& - RGI& - 1)
   Top = Top + (HBAS - Height) / 2
   '
   ' para aparear con vieja version
   ANCHOTOT% = 4: SUMAPIX = 0
   For KKI% = 1 To NUCOLUM%
     If ACL%(KKI%) > 0 Then
       ANCHOTOT% = ANCHOTOT% + ACL%(KKI%) + 1
     End If
     SUMAPIX = SUMAPIX + APIX(KKI%)
   Next KKI%
   If SUMAPIX < 1 Then
       TEPRUANCH.Caption = String$(ANCHOTOT%, "*")
       Width = TEPRUANCH.Width + 750 + MAINT + MADER
     Else
       TEPRUANCH.Caption = String$(4 + NUCOLUM%, "*")
       Width = TEPRUANCH.Width + SUMAPIX + 750 + MAINT + MADER
   End If
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
   ' POSICION ANTERIOR GRABADA
   If BSUPER > 0 Then Top = BSUPER
   If BIZQUI > 0 Then Left = BIZQUI
   If BALTUR > 0 Then Height = BALTUR
   If BANCHO > 0 Then Width = BANCHO
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
   Picture1.Top = Height - 450
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
   If FRAMEBOT2.Top - (FRAMEBOT1.Top + FRAMEBOT1.Height) > Frame2.Height + 100 Then
     Frame2.Visible = True
     Frame2.Top = FRAMEBOT2.Top - 50 - Frame2.Height
   End If
   MARCOBARRA.Top = 180
   MARCOBARRA.Left = MAIZQ + GRID.Width + MAINT + 20
   Image1.Top = 100
   Image1.Left = MARCOBARRA.Left + 120
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
            GRID.COL = KKI% - 1
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
                  If TRIM$(ae$(K2%)) <> "" Then
                    COLOG% = K2%
                    CLE% = CE%(COLOG%)
                    If CLE% > 0 Then
                      LCE% = lC%(CE%(COLOG%))
                      If LCE% > 0 Then
                        '
                        ARCHEK$ = ae$(K2%)
                        If ARCHEK$ = "UNICOM" Then ARCHEK$ = "UNIMED"
                        '
                        If Left$(ARCHEK$, 5) = "LIPRE" Then
                          ARCHEK$ = "MASTER"
                        End If
                        '
                        If Mid$(ARCHEK$, 6, 3) = "###" Then
                          NORDEX$ = TRIM$(str$(U&))
                          While Len(NORDEX$) < 3: NORDEX$ = "0" + NORDEX$: Wend
                          ARCHEK$ = Left$(ae$(K2%), 5) + NORDEX$
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
                      TTX#(K2%) = TTX#(K2%) + XVALO(TXT$)
                      If TTT% < 100 Then TXX#(TTT%) = TXX#(TTT%) + XVALO(TXT$) Else TXX#(TTT% Mod 100) = TXX#(TTT% Mod 100) - XVALO(TXT$)
                      If (TTT% Mod 100) <> CFT%(K2%) Then
                        COLOG% = CLT%(TTT% Mod 100)
                        CLE% = cl%(COLOG%)
                        LCE% = lC%(COLOG%)
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
    GRID.COL = 0
    GRID.Row = 1
    GRID.ColSel = NUCOLUM% - 1
    GRID.RowSel = FIN&
    GRID.Clip = TETUX$
    GRID.FillStyle = 0
    '
    If COLTELIB% > 0 Then
      For U& = 1 To FIN&
        XXX$ = RTrim(REGLEIDO$(ARCHILIB$, U&))
        GRID.COL = COLTELIB% - 1
        GRID.Row = U&
        If XXX$ <> GRID.TEXT Then CAMBIOVALGRILLA% = 1
        GRID.TEXT = XXX$
      Next U&
    End If
    '
    CAMBIOVALGRILLA% = 1
    '
    GRID.Row = 1
    GRID.COL = 0
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
   Call EXTRACT(PROFIL$)
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
   Command1.Visible = False
   Text11 = "": Text11.Visible = False
   If YACARDA% = 0 Then Exit Sub
   '
   If GRID.COL + 1 = COLTELIB% Then
     If LIMPIATEX% = 1 Then Text1 = ""
     LIMPIATEX% = 0
     GoTo 50
   End If
   '
   If GRID.Row >= UREPOS& Then Text1 = ""
   If GRID.Row > UREPOS& Then
     If LIMPIATEX% = 1 Then Text1 = ""
     LIMPIATEX% = 0
     Exit Sub
   End If
   '
   LIMPIATEX% = 0
   If GRID.RowIsVisible(GRID.Row) = True Then
     If GRID.Row > 0 Then
     If PE%(GRID.COL + 1) = 1 Or GRID.Row > ULTREG&(ARCHILOC$) Or GRID.Row = INSEROW& Then
     If GRID.CellWidth > 0 Then
       ' esto que sigue es para poner en blanco el texto al cambiar de celda
       ' implementado 5/5/2006 - epb
       If (Text1.Top <> GRID.CellTop + GRID.Top - 20) Or (Text1.Left <> GRID.CellLeft) Then
         Text1 = GRID.TEXT
       End If
       '
       Text1.Top = GRID.CellTop + GRID.Top - 20
       Text1.Left = GRID.CellLeft
       Text1.Height = GRID.CellHeight - 10
       Text1.Width = GRID.CellWidth
       MAXCARS% = LK%(GRID.COL + 1)
       If ACL%(GRID.COL + 1) > MAXCARS% Then MAXCARS% = ACL%(GRID.COL + 1)
       If TVL%(GRID.COL + 1) <> 6 Then MAXCARS% = 999
       TEPRUTEX = TRIM$(GRID.TEXT)
       If JT%(GRID.COL + 1) = 2 Then
         While TEPRUTEX.Width < Text1.Width
           TEPRUTEX = " " + TEPRUTEX
         Wend
         TEPRUTEX = Mid$(TEPRUTEX, 2)
       End If
       MODITEX% = 0
       Text1 = TEPRUTEX
       Text1.Visible = True
       '
       COLAC% = GRID.COL + 1
       If ae$(COLAC%) <> "" Then
          Command1.Top = GRID.CellTop + GRID.Top - 20
          Command1.Left = Text1.Left + Text1.Width - Command1.Width
          FSIZE = Command1.FontSize
          Command1.FontSize = 1
          Command1.Height = Text1.Height
          Command1.FontSize = FSIZE
          Command1.Visible = True
       End If
       '
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
       Text1.SelStart = 0
       MODITEX% = 1
       COLTEBO% = GRID.COL + 1
       ROWTEBO% = GRID.Row
     End If
     End If
     End If
   End If
   Exit Sub
   '
50 If GRID.Row >= UREPOS& Then Text11 = ""
   If GRID.Row > UREPOS& Then Exit Sub
   '
   If GRID.RowIsVisible(GRID.Row) = True Then
     If GRID.Row > 0 Then
       If PE%(GRID.COL + 1) = 1 Or GRID.Row > ULTREG&(ARCHILOC$) Or GRID.Row = INSEROW& Then
        If GRID.CellWidth > 0 Then
          Text11.Top = GRID.CellTop + GRID.Top - 20
          Text11.Left = GRID.CellLeft
          Text11.Height = 4 * GRID.CellHeight - 10
          Text11.Width = GRID.CellWidth
          '
          TEPRUTEX = RTrim(REGLEIDO$(ARCHILIB$, GRID.Row))
          MODITEX% = 0
          Text11 = TEPRUTEX
          Text11.Visible = True
          On Error Resume Next
          Text11.SetFocus
          On Error GoTo 0
          Text11.SelStart = 0
          MODITEX% = 1
          COLTEBO% = GRID.COL + 1
          ROWTEBO% = GRID.Row
        End If
       End If
     End If
   End If
   Exit Sub
   '
 
End Sub

Function UREPOS&()
   '
   UREPOS& = ULTREG&(ARCHILOC$)
   If APPI% = 1 Then UREPOS& = UREPOS& + 1
   If GRID.Rows < UREPOS& + 1 Then
     On Error Resume Next
     GRID.Rows = UREPOS& + 1
     If Err Then
        On Error GoTo 0
        Call MENSERR(24, "Imposible Cargar Grilla con " + TRIM$(str$(UREPOS&)) + " Renglones - Consulte.")
        Call FINAL("")
     End If
     On Error GoTo 0
   End If
   '
End Function

Sub MUESTRAECOS()
   '
   If GRID.Row < 1 Then Exit Sub
   '
   COLAC% = GRID.COL + 1
   If ae$(COLAC%) <> "" Then
     CLE% = CE%(COLAC%)
     If CLE% > 0 Then
       LCE% = lC%(CE%(COLAC%))
       If LCE% > 0 Then
         '
         ARCHEK$ = ae$(COLAC%)
         If ARCHEK$ = "UNICOM" Then ARCHEK$ = "UNIMED"
         '
         If Left$(ARCHEK$, 5) = "LIPRE" Then
           ARCHEK$ = "MASTER"
         End If
         '
         If Mid$(ARCHEK$, 6, 3) = "###" Then
           NORDEX$ = TRIM$(str$(GRID.Row))
           While Len(NORDEX$) < 3: NORDEX$ = "0" + NORDEX$: Wend
           ARCHEK$ = Left$(ae$(COLAC%), 5) + NORDEX$
         End If
         '
         XXE$ = ECOARCH$(ARCHEK$, CBIN$(Text1, TVL%(COLAC%), LF%(COLAC%)))
         '
         ' LO QUE SIGUE ES PARA CORREGIR ERROR DE QUE NO PRESENTA ECO EN ULTIMO RENGLON DE reimputa.prf
         If TRIM$(XXE$) = "" Then
           If TVL%(COLAC%) = 6 Then
             If TRIM$(Text1) = "" Then
               XXE$ = ECOARCH$(ARCHEK$, CBIN$(GRID.TEXT, TVL%(COLAC%), LF%(COLAC%)))
             End If
           End If
         End If
         '
         JTE% = JT%(CE%(COLAC%))
         TXT$ = CSTRING$(XXE$, 6, Len(XXE$), 0, 1)
         RECIBECLICK% = -1
         GRID.COL = CLE% - 1
         '
         ' ESTO SE PUSO EL 9 DE OCTUBRE 2006 POR PROBLEMA TELEBIT DE
         ' NO GUARDAR MODIFICACION DE TEXTO EN O/COMPRA
         ' el 30/08/07 se REEMPLAZO POR LA SIGUENTE CONDICION
         If CLE% = COLTELIB% Then
            If UCase$(XXE$) = "MATERIAL NO CODIFICADO" Then
               If TRIM$(GRID.TEXT) <> "" Then
                  GoTo 29
               End If
            End If
         End If
         If TXT$ <> GRID.TEXT Then CAMBIOVALGRILLA% = 1
         GRID.TEXT = TXT$
         '
29       If ACL%(CLE%) > 0 Then
           If GRID.Row <= ULTREG&(ARCHILOC$) Then
             Call GRAVACOLRO
           End If
         End If
         GRID.COL = COLAC% - 1
         RECIBECLICK% = 0
         '
         If ARCHEK$ = "MASTER" Or ARCHEK$ = "MATPRIM" Then
           CAMBICODI% = 0
           If TRIM$(Text1) <> TRIM$(GRID.TEXT) Then CAMBICODI% = 1
           CIKI% = CODINT%(Text1)
           If CIKI% > 0 Then
             For COLEC% = COLAC% + 1 To GRID.Cols
               If ae$(COLEC%) = "UNIMED" Then
                    JTE% = 1
                    RECIBECLICK% = (-1)
                    GRID.COL = COLEC% - 1
                    If (TRIM$(GRID.TEXT) = "" Or CAMBICODI% > 0) Then
                       GRID.TEXT = Unimed$(CIKI%)   ' SI CAMBIO EL CODIGO PONE UNIMED POR DEFAULT
                    End If
                    If ACL%(COLEC%) > 0 Then
                      If GRID.Row <= ULTREG&(ARCHILOC$) Then
                        Call GRAVACOLRO
                      End If
                    End If
                    GRID.COL = COLAC% - 1
                    RECIBECLICK% = 0
                  ElseIf ae$(COLEC%) = "UNICOM" Then
                    JTE% = 1
                    RECIBECLICK% = (-1)
                    GRID.COL = COLEC% - 1
                    If (TRIM$(GRID.TEXT) = "" Or CAMBICODI% > 0) Then
                       GRID.TEXT = UNICOM$(CIKI%)
                    End If
                    If ACL%(COLEC%) > 0 Then
                      If GRID.Row <= ULTREG&(ARCHILOC$) Then
                        Call GRAVACOLRO
                      End If
                    End If
                    GRID.COL = COLAC% - 1
                    RECIBECLICK% = 0
                  ElseIf ae$(COLEC%) = "DESCOMBI" Then
                    JTE% = 1
                    RECIBECLICK% = (-1)
                    GRID.COL = COLEC% - 1
                    If (TRIM$(GRID.TEXT) = "" Or CAMBICODI% > 0) Then
                      DESCUXX$ = DESCOMBINA$(Val(CLIPROACTIVO), CIKI%)
                       If TVL%(COLEC%) < 5 Then
                          DESCUXX$ = TRIM$(FORMATNUM$(PORCEDESCU(DESCUXX$), "F10.2"))
                       End If
                       GRID.TEXT = DESCUXX$
                    End If
                    If ACL%(COLEC%) > 0 Then
                      If GRID.Row <= ULTREG&(ARCHILOC$) Then
                        Call GRAVACOLRO
                      End If
                    End If
                    GRID.COL = COLAC% - 1
                    RECIBECLICK% = 0
               End If
             Next COLEC%
           End If
         End If
         '
         If ECOPRECIO% >= 0 Then
           If Left$(ae$(COLAC%), 5) = "LIPRE" Then
             ARCHEC$ = ae$(COLAC%)
             NLIS% = XVALO(Mid$(ARCHEC$, 6, 3))
             If NLIS% > 0 Then
               RECIBECLICK% = (-1)
               For KK1% = CE%(COLAC%) + 1 To NUCOLUM%
                 If ACL%(KK1%) > 0 Then
                   If TVL%(KK1%) < 3 Or TVL%(KK1%) > 4 Then GoTo 90 ' SI EL PRIMER CASILLERO LIBRE NO ES NUMERO SALTA AFUERA
                   PRELISMO# = PRELISTA#(ARCHEC$, CVI(CBIN$(Text1, TVL%(COLAC%), LF%(COLAC%))))
                     CANDECI% = DECILISTA%(NLIS%)
                   PRELIX$ = TRIM$(CSTRING$(MKD$(PRELISMO#), 3, 12, LDL%(KK1%), 2))
                   GRID.COL = KK1% - 1
                   JTE% = JT%(KK1%)
                   TXT$ = CSTRING$(PRELIX$, 6, ACL%(KK1%), 0, JTE%)
                   ' ESTA CONDICION SE PUSO EL 18/10/2006 POR PROBLEMA DOSIVAC DE QUE AL
                   ' LLAMAR UN PEDIDO POR PRESUPUESTO PREVIO EN LA SEGUNDA CONSULTA NO TRAÍA
                   ' EL PRECIO DEL PRESUPUESTO SINO EL DE LA LISTA DE PRECIOS
                   'If TRIM$(GRID.TEXT) = "" Or NOFRESHPRE% < 1 Or CAMBICODI% > 0 Then
                   'If TRIM$(GRID.TEXT) = "" Or NOFRESHPRE% < 1 Then
                   ' SE AGREGA EL CAMBICODI PORQUE CUANDO SE APRUEBA EL DETALLE POR PRIMERA VEZ
                   ' Y LUEGO SE EDITA PARA MODIFICAR, SI SE CAMBIABA EL CODIGO EL PRECIO NO SE ACTUALIZABA.
                   If TRIM$(GRID.TEXT) = "" Or NOFRESHPRE% < 1 Or CAMBICODI% > 0 Then
                      If TXT$ <> GRID.TEXT Then CAMBIOVALGRILLA% = 1
                      GRID.TEXT = TXT$
                   End If
                   ' FIN MODIFICA DEL 18/10/2006
                   If GRID.Row <= ULTREG&(ARCHILOC$) Then
                     Call GRAVACOLRO
                   End If
                   GRID.COL = COLAC% - 1
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
   GRID.COL = 0
   GRID.Row = 1
   GRID.ColSel = NUCOLUM% - 1
   GRID.RowSel = FINTE&
   GRID.Clip = TETUX$
   GRID.FillStyle = 0
   LIMPIAGR% = 0
   On Error GoTo 0
   '
End Sub
