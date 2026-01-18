VERSION 5.00
Begin VB.Form NUEING09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFD0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesamiento de Ingresos"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFD0&
      Caption         =   "Lote Activo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   135
      TabIndex        =   17
      Top             =   1890
      Width           =   8100
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
         Left            =   6405
         TabIndex        =   27
         Top             =   2940
         Width           =   175
      End
      Begin VB.CommandButton BOTINSE 
         Caption         =   "Aprobar"
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
         Left            =   6825
         TabIndex        =   23
         ToolTipText     =   "Asigna Ubicación al Lote Activo"
         Top             =   2940
         Width           =   1035
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
         Height          =   285
         Left            =   3645
         TabIndex        =   21
         Top             =   2970
         Width           =   2895
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFD0&
         Caption         =   "Origen:"
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
         Left            =   105
         TabIndex        =   41
         Top             =   2070
         Width           =   795
      End
      Begin VB.Label Label23 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   990
         TabIndex        =   40
         Top             =   1995
         Width           =   4695
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFD0&
         Caption         =   "Remito:"
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
         Left            =   2520
         TabIndex        =   39
         Top             =   2490
         Width           =   795
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3405
         TabIndex        =   38
         Top             =   2415
         Width           =   2280
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFD0&
         Caption         =   "Fecha:"
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
         Left            =   105
         TabIndex        =   37
         Top             =   2490
         Width           =   795
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   990
         TabIndex        =   36
         Top             =   2415
         Width           =   1230
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFD0&
         Caption         =   "Destino:"
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
         Left            =   5880
         TabIndex        =   35
         Top             =   2160
         Width           =   795
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         TabIndex        =   34
         Top             =   2415
         Width           =   1650
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFD0&
         Caption         =   "Lote:"
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
         Left            =   4935
         TabIndex        =   33
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label14 
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
         Height          =   285
         Left            =   5790
         TabIndex        =   32
         Top             =   420
         Width           =   2100
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFD0&
         Caption         =   "Envase:"
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
         Left            =   2520
         TabIndex        =   31
         Top             =   1470
         Width           =   795
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3405
         TabIndex        =   30
         Top             =   1395
         Width           =   2280
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFD0&
         Caption         =   "Cant:"
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
         Left            =   105
         TabIndex        =   29
         Top             =   1470
         Width           =   795
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   990
         TabIndex        =   28
         Top             =   1395
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6555
         TabIndex        =   26
         Top             =   1395
         Width           =   1335
      End
      Begin VB.Label Label6 
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
         Height          =   285
         Left            =   2205
         TabIndex        =   25
         Top             =   420
         Width           =   2415
      End
      Begin VB.Label Label5 
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
         Height          =   285
         Left            =   465
         TabIndex        =   24
         Top             =   810
         Width           =   7440
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFD0&
         Caption         =   "Kilos:"
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
         Left            =   5670
         TabIndex        =   22
         Top             =   1470
         Width           =   795
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFD0&
         Caption         =   "Descripción:"
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
         Left            =   105
         TabIndex        =   20
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFD0&
         Caption         =   "Ubicación:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2685
         TabIndex        =   19
         Top             =   3030
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFD0&
         Caption         =   "Código:"
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
         Left            =   1515
         TabIndex        =   18
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFD0&
      Caption         =   "Medidas Papel:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   315
      TabIndex        =   12
      Top             =   6930
      Width           =   5055
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Ancho (mm):"
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
         Left            =   210
         TabIndex        =   16
         Top             =   465
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Largo (Mt):"
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
         Left            =   2835
         TabIndex        =   15
         Top             =   465
         Width           =   1335
      End
      Begin VB.Label Label15 
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
         Left            =   1300
         TabIndex        =   14
         ToolTipText     =   "Cantidad nominal"
         Top             =   420
         Width           =   1125
      End
      Begin VB.Label Label32 
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
         Left            =   3820
         TabIndex        =   13
         ToolTipText     =   "Cantidad nominal"
         Top             =   420
         Width           =   1125
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
      Left            =   8415
      Picture         =   "NUEING09.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Aprueba y Procesa los Ingresos"
      Top             =   4620
      Width           =   800
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame3"
      Height          =   1170
      Left            =   5560
      TabIndex        =   10
      Top             =   480
      Width           =   15
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame3"
      Height          =   1170
      Left            =   4560
      TabIndex        =   9
      Top             =   480
      Width           =   15
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFD0&
      Caption         =   "Nuevos Ingresos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Left            =   135
      TabIndex        =   4
      Top             =   120
      Width           =   8100
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   2520
         TabIndex        =   8
         Top             =   360
         Width           =   15
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   210
         MousePointer    =   1  'Arrow
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   7680
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo                Lote                Kilos   Ubicación"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   7
         Top             =   315
         Width           =   7680
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   4485
      Left            =   8400
      ScaleHeight     =   4425
      ScaleWidth      =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   1380
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
         Left            =   105
         Picture         =   "NUEING09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   850
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   2
         Top             =   3750
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
         Height          =   650
         Left            =   105
         Picture         =   "NUEING09.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Salir"
         Top             =   200
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "NUEING09.frx":075E
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4050
         Width           =   1515
      End
   End
End
Attribute VB_Name = "NUEING09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTINSE_Click()
   '
   UBICAIN$ = Mid$(TRIM$(UCase(Text3.TEXT)), 1, 16)
   If UBICAIN$ <> "" Then
      If TRIM$(VALOBADA("UBIPRE", "Codiubic", "Codiubic = '" & UBICAIN$ & "'")) = "" Then
        Call MENSERR(24, "Ubicacion " + TRIM$(UBICAIN$) + " Inexistente")
        Exit Sub
      End If
      '
      ' Revisa Asignaciones Previas
      For KKI% = 1 To List1.ListCount
         If TRIM$(Mid$(List1.List(KKI% - 1), 51, 16)) = UBICAIN$ Then
            CODEPRO$ = TRIM$(Left$(List1.List(KKI% - 1), 20))
            OPX% = COMALTER%("La Ubicacion " + UBICAIN$ + " ya ha sido Asignada\al Código '" + CODEPRO$ + "'.\  \¿Confirma la Posicion Elegida?\\No, Cambiar\Si, Continuar")
            '
            If OPX% = 1 Then
                 Text3.SetFocus
                 Exit Sub
               ElseIf OPX% = 2 Then
                 GoTo 20
            End If
         End If
      Next KKI%
      '
20    CODIPROI% = XVALO(VALOBADA("LOTINGRE", "Cod_Inte", "UbicFis = '" & UBICAIN$ & "' AND CANALTA > (CANCONSU + 0.05) ORDER BY FECHALTA ASC "))
      If CODIPROI% > 0 Then
          CODEPRO$ = CODEXT$(CODIPROI%)
          OPX% = COMALTER%("La Ubicacion " + UBICAIN$ + " se Encuentra Ocupada\por Código '" + CODEPRO$ + "'.\  \¿Confirma la Posicion Elegida?\\No, Cambiar\Si, Continuar")
          '
          If OPX% = 1 Then
               Text3.SetFocus
               Exit Sub
             ElseIf OPX% = 2 Then
              TTPP$ = List1.TEXT
              Call REPLA(TTPP$, UBICAIN$, 51, 16)
              List1.List(List1.ListIndex) = TTPP$
              List1.Refresh
          End If
      Else    ' esta libre
          TTPP$ = List1.TEXT
          Call REPLA(TTPP$, UBICAIN$, 51, 16)
          List1.List(List1.ListIndex) = TTPP$
          List1.Refresh
          '
          Call FILTERPUTRECORD("TAUBICA", 1, UBICAIN$, AJUSTI$(UBICAIN$, 16) + Label6)
          Call FRESHECHO("TAUBICA")
          FORSELEC.ABUSQUE = UBICAIN$
      End If
   End If
   '
End Sub

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub


Private Sub Command2_Click()
    '
    Call ABRESTOCKS 'Abre base Stock
    For II& = 0 To List1.ListCount - 1 'Recorre el list1
        TTXX$ = List1.List(II&)
        DOCUINGR$ = Mid$(TRIM$(TTXX$), 81, 16) 'asigna el documento de ingreso
        UBICINGR$ = TRIM$(Mid$(TRIM$(TTXX$), 51, 16)) 'asigan la ubicacion
        If UBICINGR$ = "" Then GoTo 59 ' NO SE ADJUDICO UBICACION
        '
        SQLX$ = "Select * from TRANRECEP where DocuIngre = '" & DOCUINGR$ & "'" 'Selecciona el registro correspondiente al documento de ingreso
        Dim TRASRETPM As ADODB.Recordset
        Set TRASRETPM =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
        '
        With TRASRETPM
            DOCINGRE$ = SAFETEXT(!DocuIngre)
            CCIIN% = CODINT(SAFETEXT(!Cod_Exte))
            LOTEINGR$ = SAFETEXT(!idenLote)
            CANTII# = XVALO(!CanKilos)
            CANTIDA = XVALO(!Cantidad)
            FECHALTT% = CVINT(!FechMov)
            FECHFABR% = CVINT(!FechFab)
            FECHAVEN% = CVINT(!FechVenc)
            MARCAX$ = ""
            ORIGENIN$ = SAFETEXT$(!Origen)
            DEPX% = 0
            REMIPROX$ = SAFETEXT$(!NumRemProv)
            ENVASX$ = TRIM$(!CodEnvase)
            '
            APRX$ = UCase$(Left$(TRIM$(SAFETEXT$(!AproRech)), 1))
            DESTIX$ = "Cuarentena": STALIB$ = "CUAR"
            If APRX$ = "A" Or APRX$ = "1" Then DESTIX$ = "Aprobado": STALIB$ = "APRO"
            If APRX$ = "R" Or APRX$ = "3" Then DESTIX$ = "Rechazado": STALIB$ = "RECH"
            If APRX$ = "P" Or APRX$ = "4" Then DESTIX$ = "Procesa Aprob": STALIB$ = "PRAP"
            '
            IDLOT$ = LOTEINGR$
            If DESTIX$ = "Cuarentena" Then IDLOT$ = LOTEINGR$ + "-Q"
            If DESTIX$ = "Rechazado" Then IDLOT$ = LOTEINGR$ + "-R"
            '
            EXISLOT$ = VALOBADA("LOTINGRE", "IdenLote", "IdenLote = '" & IDLOT$ & "'")
            If EXISLOT$ <> "" Then     ' YA EXISTE EL LOTE DE INGRESO
               For KKII% = 65 To 90
                  IDLOT1$ = IDLOT$ + "/" + Chr$(KKII%) ' ARMA EL NUEVO SUBLOTE
                  EXISLOT1$ = VALOBADA("LOTINGRE", "IdenLote", "IdenLote = '" & IDLOT1$ & "'")
                  If EXISLOT1$ = "" Then
                     IDLOT$ = IDLOT1$
                     GoTo 30
                  End If
               Next KKII%
               Call MENSERR(24, "Imposible Ingresar Lote '" + LOTEINGR$ + "'.")
               GoTo 59
            End If
            '
30          STALIB_IN$ = STALIB$
            If STALIB$ = "PRAP" Then STALIB_IN$ = "APRO"
            Call CREALOTIN(FECHALTT%, FECHFABR%, FECHAVEN%, CCIIN%, IDLOT$, _
                 MARCAX$, LOTEINGR$, CANTII#, , , DEDPX%, UBICINGR$, ORIGENIN$, _
                 REMIPROX$, ENVASX$, STALIB_IN$, CANTIDA)
                 '
            ' PROCESA POR MOVISTO EL MOVIMIENTO DE ENTRADA
            Call MOVISTOK(FECHALTT%, CCIIN%, IDLOT$, "BR", DOCINGRE$, DOCINGRE$, REMIPROX$, REMIPROX$, "Recep. " + ORIGENIN$, 0, 0, 0, 0, 0, CANTII)
            Call ACSALOTE(CCIIN%, IDLOT$, "NOMESS")
            '
            ' SI ES APROBACION O RECHAZO, DESCUENTA DEL LOTE -Q
            CANTODES = CANTII
            If STALIB$ = "PRAP" Then
               SQLX$ = "SELECT * FROM LOTINGRE "
               SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CCIIN% & " "
               SQLX$ = SQLX$ + " AND LoteCol_Prov = '" & LOTEINGR$ & "' "
               SQLX$ = SQLX$ + " AND Libera = 'CUAR' "
               SQLX$ = SQLX$ + " ORDER BY FECHALTA ASC "
               '
               Dim LOTSEL As ADODB.Recordset
               Set LOTSEL =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
               '
               With LOTSEL
                  If Not (.BOF And .EOF) Then
                     Do While Not .EOF
                        If CANTODES < 0.05 Then GoTo 45
                        CANDES = CANTODES
                        If CANDES > XVALO(!CanAlta) - XVALO(!CanConsu) Then
                           CANDES = XVALO(!CanAlta) - XVALO(!CanConsu)
                        End If
                        IDLOTOR$ = SAFETEXT$(!idenLote)
                        Call MOVISTOK(FECHALTT%, CCIIN%, IDLOTOR$, "BR", _
                             DOCINGRE$, DOCINGRE$, REMIPROX$, REMIPROX$, _
                             DESTIX$ + " " + ORIGENIN$, 0, 0, 0, 0, 0, (-1) * CANDES)
                        Call ACSALOTE(CCIIN%, IDLOTOR$, "NOMESS")
                        CANTODES = CANTODES - CANDES
                     .MoveNext
                     Loop
                  End If
               End With
45             LOTSEL.Close
               Set LOTSEL = Nothing
            End If
            '
        End With
        '
        SQLX$ = "Update TRANRECEP Set Ingresado = 1 where DocuIngre = '" & DOCUINGR$ & "'" 'Realiza el update del campo 'Ingresado'
        Dim TRANRECTMP As ADODB.Recordset
        Set TRANRECTMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$)) 'ejecuta la consulta SQL
59  Next II&
    '
    Screen.MousePointer = 1
    Call COMUNI("Proceso Completo")
    Unload Me
    '
End Sub

Private Sub Command4_Click()
    Call SELECHO("TAUBICA")
    UBISELE$ = VALACT1$("TAUBICA")
    If UBISELE$ <> "" Then
       Text3 = UBISELE$
    End If
End Sub

Private Sub Form_Load()
    '
    '\\\OT-08-1071-FL-24/11/08/// Carga la lista
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    Call ABRESTOCKS
    CONDI$ = "STATUS = 2 AND (INGRESADO < 1 OR INGRESADO IS NULL)"
    FIN& = CantRegistros("TRANRECEP", CONDI$)
    '
    SQLX$ = "SELECT * FROM TRANRECEP WHERE STATUS = 2 AND (INGRESADO < 1 OR INGRESADO IS NULL)"
    Dim TRANRETMP As ADODB.Recordset
    Set TRANRETMP = New ADODB.Recordset
    TRANRETMP.CursorType = adOpenDynamic
    TRANRETMP.LockType = adLockPessimistic
    TRANRETMP.OPEN FILTSQL$ (SQLX$), FLEXCONN, , , adCmdText
    '
    Screen.MousePointer = 11
    JJ& = 0
    With TRANRETMP
    'RECORRO TRANRECEP
      Do While Not .EOF
        JJ& = JJ& + 1
        Call TRACE(20, JJ&, FIN&)
        DOCUI$ = TRIM$(SAFETEXT$(!CodiCompro))
        NUCO$ = TRIM$(Str$(XVALO(!NumeCompro))): While Len(NUCO$) < 6: NUCO$ = "0" + NUCO$: Wend
        NORI$ = TRIM$(Str$(XVALO(!NuOrItem))): While Len(NORI$) < 3: NORI$ = "0" + NORI$: Wend
        DOCING$ = DOCUI$ + "-" + NUCO$ + "-" + NORI$
        If Len(DOCUI$) < 1 Or Len(DOCUI$) > 5 Or Len(NUCO$) > 6 Or Len(NORI$) > 3 Then
           Call MENSERR(24, "Imposible Procesar Comprobante '" + DOCING$ + "'.")
           !Ingresado = 9            ' MARCA NO PROCESADO
           .Update
           GoTo 19
        End If
        '
        TTXX$ = Space$(128)
        CODIGO$ = AJUSTI$(SAFETEXT$(!Cod_Exte), 20)
        Call REPLA(TTXX$, CODIGO$, 1, 20)
        CIXI% = CODINT%(TRIM$(CODIGO$))
        If CIXI% < 1 Then
           !Ingresado = 9            ' MARCA NO PROCESADO
           .Update
           GoTo 19                   ' NO PUEDE PROCESAR SIN CODIGO
        End If
        '
        LOTE$ = AJUSTI$(SAFETEXT$(!idenLote), 16)
        Call REPLA(TTXX$, LOTE$, 23, 16)
        '
        CANKILS$ = FORMATNUM(XVALO(!CanKilos), "F10.1")
        Call REPLA(TTXX$, CANKILS$, 39, 10)
        '
        Call REPLA(TTXX$, DOCING$, 81, 16)
        '
        ' SI FALTA LOTE NO INGRESA
        If TRIM$(LOTE$) = "" Or TRIM$(UCase$(LOTE$)) = "NULL" Then
            !Ingresado = 9            ' MARCA NO PROCESADO
            .Update
            GoTo 19
        End If
        '
        ' NO PERMITE ENTRAR YA INGRESADOS
        CONDI$ = "Cod_Inte = " & CIXI%
        CONDI$ = CONDI$ + " AND CodiMovi = 'BR' "
        CONDI$ = CONDI$ + " AND DoPriLar = '" & DOCING$ & "' "
        DPL$ = VALOBADA("MOVISTO", "DoPriLar", CONDI$)
        If DPL$ <> "" Then
           !Ingresado = 9            ' MARCA NO PROCESADO
           .Update
           GoTo 19
        End If
        '
        For KKI% = 1 To List1.ListCount
           If TRIM$(Mid$(List1.List(KKI% - 1), 81, 16)) = DOCING$ Then
              ' AGREGA A REGISTROS NO PROCESADOS
              ' ESTO ES PARA NO PROCESAR DOS VECES EL MISMO INGRESO
              !Ingresado = 9            ' MARCA NO PROCESADO
              .Update
              GoTo 19
           End If
        Next KKI%
        !DocuIngre = DOCING$
        .Update
        '
        List1.AddItem TTXX$
19    .MoveNext
      Loop
    End With
    TRANRETMP.Close
    Set TRANRETMP = Nothing
    Screen.MousePointer = 1
    '\\\OT-08-1071-FL-///
    '
End Sub

Private Sub List1_Click()
    '
    '\\\OT-08-1071-FL-24/11/08/// Carga el formulario con los datos tomados de la lista
    TTXX$ = List1.TEXT
    DOCING$ = TRIM$(Mid$(TTXX$, 81, 16))
    Label6 = TRIM$(Mid$(TTXX$, 1, 20))
    CIXI% = CODINT%(Label6)
    Label5 = DESCRIT0$(CIXI%)
    Label7 = TRIM(Mid$(TTXX$, 39, 10))     ' CANTIDAD EN KILOS
    Label14 = TRIM$(Mid$(TTXX$, 23, 16))
    Text3.TEXT = Mid$(List1.TEXT, 51, 16)  ' UBICACION
    '
    CANOMI = XVALO(VALOBADA("TRANRECEP", "Cantidad", "DocuIngre = '" & DOCING$ & "'"))
    Label8 = FORMATNUM$(CANOMI, "F10.1")
    '
    ENVASX$ = TRIM$(VALOBADA("TRANRECEP", "CodEnvase", "DocuIngre = '" & DOCING$ & "'"))
    Label12 = ENVASX$
    '
    ORIGEX$ = TRIM$(VALOBADA("TRANRECEP", "Origen", "DocuIngre = '" & DOCING$ & "'"))
    Label23 = ORIGEX$
    '
    FECHMO = CVINT%(VALOBADA("TRANRECEP", "FechMov", "DocuIngre = '" & DOCING$ & "'"))
    Label19 = FECHATEX$(FECHMO)
    '
    REMIPROX$ = TRIM$(VALOBADA("TRANRECEP", "NumRemProv", "DocuIngre = '" & DOCING$ & "'"))
    Label21 = REMIPROX$
    '
    APRX$ = TRIM$(VALOBADA("TRANRECEP", "AproRech", "DocuIngre = '" & DOCING$ & "'"))
    APRX$ = UCase$(Left$(APRX$, 1))
    DESTIX$ = "Cuarentena"
    If APRX$ = "A" Or APRX$ = "1" Then DESTIX$ = "Aprobado"
    If APRX$ = "R" Or APRX$ = "3" Then DESTIX$ = "Rechazado"
    If APRX$ = "P" Or APRX$ = "4" Then DESTIX$ = "Procesa Aprob"
    Label17 = DESTIX$
    '
End Sub



