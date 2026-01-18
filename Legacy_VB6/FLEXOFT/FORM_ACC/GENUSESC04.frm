VERSION 5.00
Begin VB.Form GENUSESC04 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generador de Números de Serie"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   6840
      Width           =   8175
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Gas Natural"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   120
         TabIndex        =   20
         Top             =   470
         Width           =   3135
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   195
         Width           =   3135
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Left            =   6000
         TabIndex        =   12
         Top             =   300
         Width           =   1995
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "N. de Serie Generado:"
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
         Left            =   3600
         TabIndex        =   11
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00004080&
      Height          =   7965
      Left            =   8640
      ScaleHeight     =   7905
      ScaleWidth      =   885
      TabIndex        =   1
      Top             =   0
      Width           =   945
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
         Height          =   640
         Left            =   105
         Picture         =   "GENUSESC04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Genera y graba N.Serie"
         Top             =   6840
         Width           =   650
      End
      Begin VB.CommandButton Command14 
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
         Picture         =   "GENUSESC04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   5520
         Width           =   650
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
         Height          =   640
         Left            =   105
         Picture         =   "GENUSESC04.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir"
         Top             =   480
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   2
         Top             =   6360
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
   End
   Begin VB.Frame Frame40 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Generador de Números de Serie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6740
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Seleccionar Opción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   360
         TabIndex        =   21
         Top             =   6120
         Width           =   7455
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Gas Envadado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   4320
            TabIndex        =   23
            Top             =   150
            Width           =   2055
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Gas Natural"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1800
            TabIndex        =   22
            Top             =   150
            Width           =   2535
         End
      End
      Begin VB.CommandButton Command5 
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
         Left            =   7680
         TabIndex        =   10
         Top             =   240
         Width           =   175
      End
      Begin VB.Image Image1 
         Height          =   2055
         Index           =   5
         Left            =   5520
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   2055
         Index           =   4
         Left            =   3000
         Picture         =   "GENUSESC04.frx":0896
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   2055
         Index           =   3
         Left            =   360
         Picture         =   "GENUSESC04.frx":4A252
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   2055
         Index           =   2
         Left            =   5520
         Picture         =   "GENUSESC04.frx":81D49
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   2055
         Index           =   1
         Left            =   3000
         Picture         =   "GENUSESC04.frx":C145A
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   2055
         Index           =   0
         Left            =   360
         Picture         =   "GENUSESC04.frx":C5092
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label40 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   5
         Left            =   5520
         TabIndex        =   18
         Top             =   5520
         Width           =   2325
      End
      Begin VB.Label Label40 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   4
         Left            =   3000
         TabIndex        =   17
         Top             =   5520
         Width           =   2325
      End
      Begin VB.Label Label40 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   3
         Left            =   360
         TabIndex        =   16
         Top             =   5520
         Width           =   2325
      End
      Begin VB.Label Label40 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   2
         Left            =   5520
         TabIndex        =   15
         Top             =   2760
         Width           =   2325
      End
      Begin VB.Label Label40 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   1
         Left            =   3000
         TabIndex        =   14
         Top             =   2760
         Width           =   2325
      End
      Begin VB.Label Label40 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   2760
         Width           =   2325
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   6600
         TabIndex        =   9
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Left            =   5160
         TabIndex        =   8
         Top             =   320
         Width           =   1335
      End
   End
   Begin VB.Label Label7 
      Height          =   135
      Left            =   1320
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      Height          =   135
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "GENUSESC04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TOPEINI(5), IZQUINI(5), ANCHINI(5), ALTUINI(5)
Dim KKODDGN$(5) ' VECTOR DE CODIGO
Dim KKODDGE$(5)


Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Call FINAL("")
   '
End Sub


Private Sub Command14_Click()
    '
    Command14.Enabled = False
    '
    FSTPFET04.Show 1
    '
    Command14.Enabled = True
    
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    '
    If TRIM$(Label2) = "" Then ' validacion de que se haya seleccionado una imagen
     Call COMUNI("Debe Seleccionar Una Imagen Para Generar el Nº Serie")
     GoTo 99
    End If
    '
    If Option1(0).Value = True Then
     CODEXXXT$ = TRIM$(Label6) ' valida que el codigo sea existente
    Else
     CODEXXXT$ = TRIM$(Label7)
    End If
    '
    CI% = CODINT%(CODEXXXT$)
    If CI% < 1 Then
     Call COMUNI("Código Inexistente")
     GoTo 99
    End If
    '
    If TRIM$(Label5) = "" Then
     Call COMUNI("Debe Seleccionar Gas Natural o Gas Envasado")
     GoTo 99
    End If
    '
    Screen.MousePointer = 11
    '
    FE% = FECHANUM(Label3)
    DESCRIPCION$ = DESCRIT(CI%)
    MAÑO$ = REPLACAR$(Label3, "/", "")
    MESYANO$ = Mid$(TRIM$(MAÑO$), 3, 4)
    NumSeriDes$ = TRIM$(Label2)
    NumOrden% = 1 + MaxNumOr%(MESYANO$)
    '
    SQLX$ = "SELECT * FROM TANUMSE"
    Set TANUMTMP = BDEQUIPOS.OpenRecordset(SQLX$, dbOpenDynaset)
    With TANUMTMP
        .AddNew
        !FechAlta = CVDAT(FE%)
        !Cod_Inte = CI%
        !Cod_Exte = CODEXT$(CI%)
        !Descrip = DESCRIPCION$
        !MESANIO = MESYANO$
        !NumeroSerie = NumSeriDes$
        !NUMEORD = NumOrden%
        .Update
    End With
    Call PrintEtqEsc
    Screen.MousePointer = 1
    Call COMUNI("Alta de N.Serie Realizada Exitosamente")
    '
    Label2 = ""
    Label4 = ""
    Label5 = ""
    Label6 = ""
    Label7 = ""
    '
99  Command2.Enabled = True
End Sub



Private Sub Command5_Click()
  '
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
    Label3 = VDEV$
  End If
  '
End Sub

Private Sub Form_Load()
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    Call ABREEQUIPOS
    Call CREATABLA_EQUIPOS
    FF% = HOY(HO$)
    Label3 = HO$ ' ASIGNACION DE FECHA DEL DIA
    '
    Option1(0).Value = True
    
'    RUTARGRA$ = TRIM$(Control$("", "RUTARGRA")) ' RUTA DEL LOGO
'    If RUTARGRA$ = "" Then RUTARGRA$ = LUDAT$
'
'    For kk% = 0 To 4
'      TOPEINI(kk%) = Picture1(kk%).Top
'      IZQUINI(kk%) = Picture1(kk%).Left
'      ANCHINI(kk%) = Picture1(kk%).Width
'      ALTUINI(kk%) = Picture1(kk%).Height
'    Next kk%
'
'    For kk% = 0 To 4
'      'Label34(KK%).Caption = ""
'      Picture1(kk%).Top = TOPEINI(kk%)
'      Picture1(kk%).Left = IZQUINI(kk%)
'      Picture1(kk%).Width = ANCHINI(kk%)
'      Picture1(kk%).Height = ALTUINI(kk%)
'      Picture1(kk%).Picture = LoadPicture("")
'   Next kk%
'   Index = 0 ' PARA PROBAR
   For Index = 0 To 4
'    Image1(Index).Picture = LoadPicture("")
'    Picture1(Index).Cls
      Select Case Index
        Case 0
'          COPLANO$ = RUTARGRA$ + "\MontajePalaceAcero.jpg"
          Label40(Index).Caption = "Palace INOX"
          KKODDGN$(Index) = "01.01-00170"
          KKODDGE$(Index) = "01.01-00180"
        Case 1
'          COPLANO$ = RUTARGRA$ + "\MontajeCandorBlanca.jpg"
          Label40(Index).Caption = "Candor Cocina Blanca"
          KKODDGN$(Index) = "01.01-00050"
          KKODDGE$(Index) = "01.01-00040"

        Case 2
'          COPLANO$ = RUTARGRA$ + "\MontajePalaceBlanca.jpg"
          Label40(Index).Caption = "Palace Cristal"
          KKODDGN$(Index) = "01.01-00150"
          KKODDGE$(Index) = "01.01-00160"
        Case 3
'          COPLANO$ = RUTARGRA$ + "\MASTER BLANCA.jpg"
          Label40(Index).Caption = "Master Cocina Blanca"
          KKODDGN$(Index) = "01.01-14010"
          KKODDGE$(Index) = "01.01-14020"
        Case 4
'          COPLANO$ = RUTARGRA$ + "\MASTER Inox.jpg"
          Label40(Index).Caption = "Master Inox"
          KKODDGN$(Index) = "01.01-14070"
          KKODDGE$(Index) = "01.01-14080"

      End Select
'        If COPLANO$ <> "" Then
'           If EXISTE%(COPLANO$) > 0 Then
'              On Error GoTo 90
'              Image1(Index).Picture = LoadPicture(COPLANO$)
'              Picture1(Index).Top = TOPEINI(Index)
'              Picture1(Index).Left = IZQUINI(Index)
'              Picture1(Index).Width = ANCHINI(Index)
'              Picture1(Index).Height = ALTUINI(Index)
'              Picture1(Index).Picture = LoadPicture("")
'              Picture1(Index).Refresh
'              '
'              HIMAGEN = Image1(Index).Picture.Width
'              VIMAGEN = Image1(Index).Picture.Height
'
'              HCUADRO = Picture1(Index).Width
'              VCUADRO = Picture1(Index).Height
'              '
'              REDUH = HIMAGEN / HCUADRO
'              REDUV = VIMAGEN / VCUADRO
'              REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
'              Picture1(Index).Height = VCUADRO * REDUV / REDUC
'              Picture1(Index).Width = HCUADRO * REDUH / REDUC
'              Picture1(Index).Refresh
'              '
'              HCUADRO = Picture1(Index).Width
'              VCUADRO = Picture1(Index).Height
'              Picture1(Index).Left = (Frame4(Index).Width - HCUADRO) / 2
'              '
'              Picture1(Index).PaintPicture Image1(Index).Picture, 0, 0, HCUADRO, VCUADRO
'              Picture1(Index).Visible = True
'              Picture1(Index).Refresh
'            Else
'           End If
'        End If
   Next Index
'   GoTo 98
   '
'90 On Error GoTo 0
'   Call MENSERR(24, "Imagen no Válida")
   '
98 Screen.MousePointer = 1
End Sub
 Function MaxNumOr%(MESAÑO$)
    MaxNXX% = 0
    SQLX$ = "SELECT MAX(NumeOrd) AS MAXNUM FROM TANUMSE "    ' TRAE EL VALOR MAXIMO DEL NUMERO DE ORDEN DEL MESAÑO PASADO COMO ARGUMENTO
    SQLX$ = SQLX$ + " WHERE MESANIO = '" & MESAÑO$ & "' or MESANIO = '" & XVALO(MESAÑO$) & "'"
    Set TANUMTMP = BDEQUIPOS.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With TANUMTMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         MaxNXX% = XVALO(!MAXNUM)
      End If
      On Error GoTo 0
   End With
   MaxNumOr% = MaxNXX%

 End Function

Private Sub Frame4_DragDrop(Index As Integer, Source As CONTROL, X As Single, Y As Single)

End Sub

Private Sub Image1_DblClick(Index As Integer)
    
    If Index = 5 Then Exit Sub
    
    VDEVV$ = TRIM$(Label3)
    '
    If FECHANUM(VDEVV$) > 0 Then ' SI LA FECHA ES VALIDA
         MAÑO$ = REPLACAR$(Label3, "/", "")
         MESYANO$ = Mid$(TRIM$(MAÑO$), 3, 4)
         NROORDEN$ = TRIM$(Str(1 + MaxNumOr%(MESYANO$))) 'TOMA EL VALOR MAXIMO DE NUMERO DE ORDEN + 1
         '
         While Len(NROORDEN$) < 4: NROORDEN$ = "0" & NROORDEN$: Wend ' FORMATEA
         
         Label6 = KKODDGN$(Index) 'NUMERO DE CODIGO EXTERNO DE IMAGEN CLICKEADA
         Label7 = KKODDGE$(Index)
         If Option1(0).Value = True Then
             KODI$ = Right(Label6, 5)
         Else
             KODI$ = Right(Label7, 5)
         End If
         '
         Label2 = KODI$ & MESYANO$ & NROORDEN$ ' CONCATENA
         Label4 = TRIM$(Label40(Index))
         '
     Else
         Call COMUNI("Fecha Inválida")
         Label2 = ""
    End If
    
    If Option1(0).Value = True Then
      Label5 = "Gas Natural"
    Else
      Label5 = "Gas Envasado"
    End If
    

End Sub

Private Sub Option1_Click(Index As Integer)
     Label5 = TRIM$(Option1(Index).Caption)
     If Option1(0).Value = True Then
         KOD$ = Right(Label6, 5)
         Label2 = KOD$ + Mid$(Label2, 6)
         Label5 = "Gas Natural"
     Else
        KOD$ = Right(Label7, 5)
        Label2 = KOD$ + Mid$(Label2, 6)
        Label5 = "Gas Envasado"
     End If

End Sub

Private Sub Picture1_DblClick(Index As Integer)

End Sub
Sub PrintEtqEsc()

    TIPODOC$ = "Etiqueta de Numero de Serie" ' guardar el documento immpreso
    NUMEDOC$ = TRIM$(Label2)
    '
    CODPARAM$(1) = "DES-MAT": DOCPARAM$(1) = Label4 'DESCRIPCION
    CODPARAM$(2) = "NRO-LOTE": DOCPARAM$(2) = TRIM$(Label2) 'NUMERO DE SERIE
    CODPARAM$(3) = "REF-MAT1": DOCPARAM$(3) = TRIM$(Label2) 'NUMERO DE SERIE
    CODPARAM$(4) = "REF-MAT2": DOCPARAM$(4) = TRIM$(Label2) 'NUMERO DE SERIE
    CAPARDOC% = 4
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FUENTIE"))
    Call PRINTDOC("@" + FORIPRE$)
    '
End Sub
