VERSION 5.00
Begin VB.Form FORINVAJU 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventarios de Materiales - Generación y Consulta"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   225
      Left            =   3465
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Scrap"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   7770
      Picture         =   "FORINVAJU.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Tabla de Materiales Obsoletos"
      Top             =   3605
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Pr.Prod."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   7770
      Picture         =   "FORINVAJU.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Tabla Valores Mensuales Productos Propia Producción"
      Top             =   1955
      Width           =   855
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Selección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1470
      Left            =   210
      TabIndex        =   11
      Top             =   2930
      Width           =   7365
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   6405
         ScaleHeight     =   75
         ScaleWidth      =   810
         TabIndex        =   29
         Top             =   315
         Width           =   865
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   960
         Left            =   3465
         TabIndex        =   26
         Top             =   315
         Width           =   2640
         Begin VB.OptionButton Option12 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Items Obsoletos"
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
            Height          =   225
            Left            =   105
            TabIndex        =   28
            Top             =   525
            Width           =   2115
         End
         Begin VB.OptionButton Option11 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Stock Total Inventario"
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
            Left            =   105
            TabIndex        =   27
            Top             =   210
            Value           =   -1  'True
            Width           =   2640
         End
      End
      Begin VB.CommandButton Command3 
         Height          =   690
         Left            =   6405
         Picture         =   "FORINVAJU.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   630
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Propia Producción"
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
         Left            =   420
         TabIndex        =   14
         Top             =   1050
         Width           =   2010
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Reventa (Importados)"
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
         Left            =   420
         TabIndex        =   13
         Top             =   735
         Width           =   2325
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Listado General"
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
         Left            =   420
         TabIndex        =   12
         Top             =   420
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   6310
         X2              =   6310
         Y1              =   105
         Y2              =   1470
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   6300
         X2              =   6300
         Y1              =   105
         Y2              =   1440
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   3160
         X2              =   3160
         Y1              =   105
         Y2              =   1470
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   3150
         X2              =   3150
         Y1              =   105
         Y2              =   1440
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2640
      Left            =   3990
      TabIndex        =   10
      Top             =   210
      Width           =   3585
      Begin VB.OptionButton Option10 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Matriz de Valores Mensuales"
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
         Left            =   420
         TabIndex        =   23
         Top             =   2310
         Width           =   2850
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Listado General Integrado"
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
         Left            =   420
         TabIndex        =   21
         Top             =   1995
         Value           =   -1  'True
         Width           =   2850
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Valor Neto de Realización"
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
         Left            =   420
         TabIndex        =   19
         Top             =   1680
         Width           =   2850
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Costo de Reposición"
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
         Left            =   420
         TabIndex        =   18
         Top             =   1365
         Width           =   2850
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Precio Ultima Compra"
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
         Left            =   420
         TabIndex        =   17
         Top             =   1050
         Width           =   2850
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Historico Actualizado p/Partida"
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
         Left            =   420
         TabIndex        =   16
         Top             =   735
         Width           =   3060
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Valor Historico por Partida"
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
         Left            =   420
         TabIndex        =   15
         Top             =   420
         Width           =   2850
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cálculo y Generación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2640
      Left            =   210
      TabIndex        =   3
      Top             =   210
      Width           =   3585
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2475
         TabIndex        =   6
         Top             =   1785
         Width           =   2535
         Begin VB.CommandButton Command7 
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
            Left            =   0
            Picture         =   "FORINVAJU.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Re-Calcular Exis- tencias a Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   8
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
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
         Left            =   2205
         TabIndex        =   5
         Text            =   " "
         Top             =   420
         Width           =   1065
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
         Height          =   330
         Left            =   315
         TabIndex        =   4
         Text            =   " "
         Top             =   840
         Width           =   2955
      End
      Begin VB.Label Label3 
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
         Height          =   330
         Left            =   1785
         TabIndex        =   25
         Top             =   1260
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valores Pr.Prod:"
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
         Left            =   -210
         TabIndex        =   24
         Top             =   1340
         Width           =   1905
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Inventario:"
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
         Left            =   210
         TabIndex        =   9
         Top             =   525
         Width           =   1905
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cotiz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   7770
      Picture         =   "FORINVAJU.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Tablas Cotizaciones Dolar, Indice de Ajuste y Costos Inventario Anterior"
      Top             =   2775
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "P.Impor."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   7770
      Picture         =   "FORINVAJU.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Productos Importados (Material de Reventa)"
      Top             =   1135
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Height          =   780
      Left            =   7770
      Picture         =   "FORINVAJU.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   315
      Width           =   855
   End
End
Attribute VB_Name = "FORINVAJU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command10_Click()
    '
    If Option1.Value = True Then
        Call FINAL("*COMATCO")
      ElseIf Option2.Value = True Then
        Call FINAL("*LIMATCO")
    End If
    '
End Sub


Private Sub Command2_Click()
   Command2.Enabled = False
   INVEIMPO.Show 1
   Command2.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
FFXX1% = FECHANUM(Text1)
'If FFXX1% < FECHANUM("01/01/02") Or FFXX1% > FECHANUM("31/12/02") Then
'  Call COMUNI("Fecha de Inventario\Fuera de Límites de Diseño\de esta Aplicación !!!\  \Los Valores Ajustados son Condicionales.")
'End If
    If Option1.Value = True Then       ' LISTADO GENERAL
        ORIMATE$ = "General"
        ORIMAT% = 0
      ElseIf Option2.Value = True Then ' IMPORTADOS REVENTA
        ORIMATE$ = "Importados (Reventa)"
        ORIMAT% = 2
      ElseIf Option3.Value = True Then ' PROPIA PRODUCCION
        ORIMATE$ = "Nacional (Prop.Producción)"
        ORIMAT% = 1
    End If
    '
    Call COPYSTRU("PARTIDOS", "COREPACT")
    AREPACT$ = "CORACT" + Mid$(Text1, 7, 2)
    Call COPYSTRU("COREPACT", AREPACT$)
    '
    INTEGRA% = 0
    If Option4.Value = True Then
       TIVALUA$ = "Valor Histórico por Partida"
    ElseIf Option5.Value = True Then
       TIVALUA$ = "Histórico Actualizado por Partida"
    ElseIf Option6.Value = True Then
       INTEGRA% = 1: TIVALUA$ = "Precio Ultima Compra"
    ElseIf Option7.Value = True Then
       INTEGRA% = 1: TIVALUA$ = "Costo de Reposición"
    ElseIf Option8.Value = True Then
       INTEGRA% = 1: TIVALUA$ = "Valor Neto Realización"
    ElseIf Option9.Value = True Then
       TIVALUA$ = "Integrado"
    ElseIf Option10.Value = True Then
       Call MATVAMEN
       GoTo 99
    End If
    '
    TOBSOLE$ = ""
    ARENTRA$ = "PARTIDIS"
    If Option12.Value = True Then
      TOBSOLE$ = " (Obsoletos)"
      ARENTRA$ = "PORTIDIS"
    End If
    '
    Screen.MousePointer = 11
    Call SETULTREG("PARTIDOS", 0)
    FIN& = ULTREG&(ARENTRA$)
    CIANT% = 0: JJ& = 0
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$(ARENTRA$, U&)
      CI% = CVI(Left$(XX$, 2))
      GRUORI% = Asc(Mid$(XX$, 3, 1))
      If ORIMAT% = 0 Or GRUORI% = ORIMAT% Then
        If INTEGRA% = 1 Then
          If CI% = CIANT% Then
            YY$ = REGLEIDO$("PARTIDOS", JJ&)
            CANTE = CVS(Mid$(XX$, 49, 4)) + CVS(Mid$(YY$, 49, 4))
            Call REPLA(XX$, MKS$(CANTE), 49, 4)
            CANTE = CVS(Mid$(XX$, 53, 4)) + CVS(Mid$(YY$, 53, 4))
            Call REPLA(XX$, MKS$(CANTE), 53, 4)
            Call REPLA(XX$, MKS$(CANTE), 57, 4)
            CANTED# = CVD(Mid$(XX$, 113, 8)) + CVD(Mid$(YY$, 113, 8))
            Call REPLA(XX$, MKD$(CANTED#), 113, 8)
            Call REPLA(XX$, MKD$(CANTED#), 121, 8)
            CANTED# = CVD(Mid$(XX$, 137, 8)) + CVD(Mid$(YY$, 137, 8))
            Call REPLA(XX$, MKD$(CANTED#), 137, 8)
            Call REPLA(XX$, MKD$(CANTED#), 145, 8)
            CANTED# = CVD(Mid$(XX$, 185, 8)) + CVD(Mid$(YY$, 185, 8))
            Call REPLA(XX$, MKD$(CANTED#), 185, 8)
            Call REPLA(XX$, MKD$(CANTED#), 193, 8)
            Call GRAREG("PARTIDOS", XX$, JJ&)
            GoTo 19
          End If
        End If
        JJ& = JJ& + 1
        Call GRAREG("PARTIDOS", XX$, JJ&)
        CIANT% = CI%
      End If
19  Next U&
    Call SETULTREG("PARTIDOS", JJ&)
    Call HEADERS("PARTIDOS", "")
    Call HEADERS("PARTIDOS", "Corresponde a: " + Text1 + " - " + ORIMATE$ + " - " + TIVALUA$ + TOBSOLE$)
    '
    ' ESTO QUE SIGUE ES PARA RESPALDAR EL COSTO DE REPOSICIóN ACTUALIZADO
    If ORIMAT% = 0 Then
      If INTEGRA% = 1 Then
        FIN& = ULTREG&("PARTIDOS")
        For U& = 1 To FIN&
          XX$ = REGLEIDO$("PARTIDOS", U&)
          Call GRAREG(AREPACT$, XX$, U&)
        Next U&
        Call SETULTREG(AREPACT$, FIN&)
        Call CIERRARCH(AREPACT$)
      End If
    End If
    ' *******************************************************************
    '
    If Option4.Value = True Then
         Call FINAL("*LIPARTI1")
      ElseIf Option5.Value = True Then
         Call FINAL("*LIPARTI2")
      ElseIf Option6.Value = True Then
         Call FINAL("*LIPARTI3")
      ElseIf Option7.Value = True Then
         Call FINAL("*LIPARTI4")
      ElseIf Option8.Value = True Then
         Call FINAL("*LIPARTI5")
      ElseIf Option9.Value = True Then
         Call FINAL("*LIPARTID")
    End If
    Screen.MousePointer = 1
    '
99  Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    '
    Command4.Enabled = False
10  ANOX$ = TRIM$(InputBox$("Año a Listar / Actualizar", "Período de Actualización"))
    If Len(ANOX$) < 2 Then GoTo 99
    If Len(ANOX$) < 2 Then GoTo 10
    ANOX$ = Right$(ANOX$, 2)
    ATRI$ = "Año 20" + ANOX$
    '
    ARESAN$ = LUDAT$ + "CPROP_" + ANOX$ + ".DAT"
    If EXISTE%(ARESAN$) < 1 Then
      Call MENSERR(24, "No Existe " + REPLACAR$(ARESAN$, "\", "/"))
      GoTo 10
    End If
    '
    TIVAIMP% = COMALTER%("Tipo de Valores a Presentar:\\Históricos\Pre-Ajustados")
    If TIVAIMP% = 1 Then
         ARESAN$ = "CPROP_" + ANOX$
         ATRI$ = ATRI$ + " - Historicos"
       ElseIf TIVAIMP% = 2 Then
         ARESAN$ = "CPAJU_" + ANOX$
         ATRI$ = ATRI$ + " - Pre-Ajustados"
       Else
         GoTo 99
    End If
    '
    PERIO% = COMALTER%("Opciones de Actualización o Listado:\\1er. Semestre\2do. Semestre\Listado Anual")
    '
    If PERIO% = 1 Then
        Call TRALOCAL(ARESAN$, "COSPROPR")
        ATRI$ = "/TITUPAN=Costos de Reposición - " + ATRI$
        VTB% = VENTABU%("COSPROP1" + ATRI$)
        If VTB% >= 0 Then
          Call TRACENTRAL(ARESAN$, "COSPROPR")
          Call FILESORT(ARESAN$, "", 1, 1, "ASC")
        End If
      ElseIf PERIO% = 2 Then
        Call TRALOCAL(ARESAN$, "COSPROPR")
        ATRI$ = "/TITUPAN=Costos de Reposición - " + ATRI$
        VTB% = VENTABU%("COSPROP2" + ATRI$)
        If VTB% >= 0 Then
          Call TRACENTRAL(ARESAN$, "COSPROPR")
          Call FILESORT(ARESAN$, "", 1, 1, "ASC")
        End If
      ElseIf PERIO% = 3 Then
        FICO% = COPYFILE%(LUDAT$ + ARESAN$ + ".DAT", LUDAT$ + "COSPROPR.DAT")
        If FICO% = 1 Then
           Call HEADERS("COSPROPR", "")
           Call HEADERS("COSPROPR", "Corresponde a: Costos de Reposición - " + ATRI$)
           Call FINAL("*COSPROPL")
        End If
      Else
        GoTo 99
    End If
    '
    GoTo 10
    '
99  Command4.Enabled = True
    '
End Sub

Sub IMPORCOS()
   '
10 ARIMPO$ = InputBox$("Archivo Fuente", "Importar Costos Mensuales")
   If ARIMPO$ = "" Then Exit Sub
   '
   If EXISTE%(ARIMPO$) < 1 Then
     Call MENSERR(24, "Imposible Abrir Archivo")
     GoTo 10
   End If
   '
20 MESIMPO$ = InputBox$("Periodo a Importar (0 ... 12)", "Importar Costos Mensuales")
   If MESIMPO$ = "" Then Exit Sub
   PERIX& = Val(MESIMPO$)
   If PERIX& < 0 Or PERIX& > 12 Then
     Call MENSERR(24, "Periodo no Válido")
     GoTo 20
   End If
   '
30 TIVAIMP% = COMALTER%("Tipo de Valores a Importar:\\Históricos\Pre-Ajustados")
   If TIVAIMP% < 1 Or TIVAIMP% > 2 Then Exit Sub
   '
   ARESAN$ = "COSPROPR"
   If TIVAIMP% = 2 Then
     ARESAN$ = "COSPRAJU"
     Call COPYSTRU("COSPROPR", "COSPRAJU")
     If EXISTE%(LUDAT$ + "COSPRAJU.DAT") < 1 Then
       Screen.MousePointer = 11
       FIN& = ULTREG&("COSPROPR")
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XX$ = Left$(REGLEIDO$("COSPROPR", U&), 8) + String$(128, 0)
         Call GRAREG("COSPRAJU", XX$, U&)
       Next U&
       Call CIERRARCH("COSPRAJU")
       Screen.MousePointer = 1
     End If
   End If
   '
   Screen.MousePointer = 11
   N1% = FILEOPEN%(ARIMPO$, 1, 128)
   JJ& = 0
   While Not EOF(N1%)
     Line Input #N1%, AAX$
     PPXX% = InStr(AAX$, ";")
     If PPXX% > 0 Then
       BBX$ = Mid$(AAX$, PPXX% + 1)
       PPXX% = InStr(BBX$, ";")
       If PPXX% > 0 Then
         CODUZ$ = TRIM$(Left$(BBX$, PPXX% - 1))
         CCX$ = Mid$(BBX$, PPXX% + 1)
         DDX$ = REPLACAR$(CCX$, ",", ".")
         COSUX# = Val(DDX$)
         '
         CODUX$ = CODUZ$: GoSub 90
         CODUX$ = CODUZ$ + "L": GoSub 90
         CODUX$ = CODUZ$ + "X": GoSub 90
         CODUX$ = CODUZ$ + "L": GoSub 90
         '
       End If
     End If
   Wend
   Close #N1%
   '
   If JJ& > 0 Then
     Call SETULTREG("VUNINVAN", JJ&)
   End If
   Call CIERRARCH("VUNINVAN")
   '
   Screen.MousePointer = 1
   Exit Sub
   '
90 CII% = CODINT%(CODUX$)
   If CII% > 0 Then
     If CII% <= NUMAS% Then
       RFF$ = RECFILT$(ARESAN$, 1, MKI$(CII%))
       If Len(RFF$) >= 4 Then
         RECOL& = CVS(Left$(RFF$, 4))
         If RECOL& > 0 Then
           XX$ = REGLEIDO$(ARESAN$, RECOL&)
           For PRIKK& = PERIX& To 12
             POINSE% = 9 + 8 * PRIKK&
             '
             ' lo siguiente para que no sobreescriba
             ' valores ya cargados de meses siguientes
             If PRIKK& > PERIX& Then
               If CVD(Mid$(XX$, POINSE%, 8)) >= 0.00005 Then
                 GoTo 95
               End If
             End If
             '
             Call REPLA(XX$, MKD$(COSUX#), POINSE%, 8)
           Next PRIKK&
95         Call GRAREG(ARESAN$, XX$, RECOL&)
         End If
       End If
       '
       If TIVAIMP% = 1 Then
         If PERIX& = 0 Then
           If COSUX# >= 0.00005 Then
             YY$ = MKI$(CII%) + MKD$(COSUX#)
             JJ& = JJ& + 1
             Call GRAREG("VUNINVAN", YY$, JJ&)
           End If
         End If
       End If
       '
     End If
   End If
Return
   '
End Sub
'
Private Sub Command5_Click()
   '
   Command5.Enabled = False
   OPX% = COMALTER%("Opciones de Acceso a tabla de Obsoletos:\\Editar Tabla\Listado de Obsoletos\Cancelar")
   If OPX% = 1 Then
       Call TRALOCAL("TAOBSOLE", "")
       VTB% = VENTABU%("TAOBSOLE")
       If VTB% >= 0 Then
         Call TRACENTRAL("TAOBSOLE", "")
         Call FILESORT("TAOBSOLE", "", 1, 1, "ASC")
       End If
     ElseIf OPX% = 2 Then
       Call FINAL("*LIOBSO")
       'IMPOROBSOL.Show 1
   End If
   Command5.Enabled = True
   '
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   OPX% = COMALTER%("Opciones de Trabajo:\\Tabla Cotiza Dolar\Indice P.M.N.G.\Costo Rep.Invent.Anteriores")
   If OPX% = 1 Then
       OPXX% = ALTERNA%("Actualizar\Listar")
       If OPXX% = 1 Then
         Call FINAL("*CODIADOL")
       Else
         Call FINAL("*LCODIADO")
       End If
     ElseIf OPX% = 2 Then
       Call TRALOCAL("INDAJMEN", "")
       VTB% = VENTABU%("INDAJMEN")
       If VTB% >= 0 Then
         Call TRACENTRAL("INDAJMEN", "")
         OPXX% = COMALTER%("¿Desea Listar los Indices de Precios Mayoristas?\\No, Continuar\Si, Listar")
         If OPXX% = 2 Then
           Call FINAL("*LINDAJME")
         End If
       End If
     ElseIf OPX% = 3 Then
       Call LICOREPAN
   End If
   Command6.Enabled = True
End Sub

Sub LICOREPAN()
   '
10 ANXX$ = InputBox$("Costos de Reposición Actualizados de 20..", "Selección Año de Inventario")
   If ANXX$ = "" Then Exit Sub
   '
   AREPACT$ = "CORACT" + ANXX$
   If EXISTE%(LUDAT$ + AREPACT$ + ".DAT") < 1 Then
     Call MENSERR(24, "No Existe Archivo Solicitado")
     GoTo 10
   End If
   '
   ORIMAT% = COMALTER%("Clasificación por Origen:\\P.Prod.\Reventa\Todos Juntos")
   If ORIMAT% < 1 Or ORIMAT% > 3 Then Exit Sub
   '
   If ORIMAT% > 2 Then ORIMAT% = 0
   '
   JJ& = 0
   FIN& = ULTREG&(AREPACT$)
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$(AREPACT$, U&)
     GRUORI% = Asc(Mid$(XX$, 3, 1))
     If ORIMAT% = 0 Or GRUORI% = ORIMAT% Then
       JJ& = JJ& + 1
       Call GRAREG("COREPACT", XX$, JJ&)
     End If
   Next U&
   Call SETULTREG("COREPACT", JJ&)
   Call CIERRARCH("COREPACT")
   '
   Call HEADERS("COREPACT", "")
   Call HEADERS("COREPACT", "Corresponde a: 31 de Diciembre Año 20" + ANXX$)
   '
   Call FINAL("*LIREPACT")
   '
End Sub
'
Private Sub Command7_Click()
   GENINVAJU.Show 1
End Sub

Private Sub Command8_Click()
   Call COPYSTRU("TAOBSOLE", "TAOBSOLF")
   N1% = FILEOPEN%("F:\CLIENTES\SABO_WIN\DATOS\OBSOLEPP.PRN", 1, 128)
   JJ& = 0
   While Not EOF(N1%)
      Line Input #N1%, A$
      CODD$ = TRIM$(Mid$(A$, 1, 15))
      CIXI% = CODINT%(CODD$)
      CAXS$ = TRIM$(Mid$(A$, 61, 10))
      Call REPLACAR(CAXS$, " ", "")
      CANN# = Val(CAXS$)
      If CIXI% > 0 Then
        If CANN# > 0.05 Then
          BB$ = MKI$(CIXI%) + MKD$(CANN#) + "BAJA AL 04/03/03"
          JJ& = JJ& + 1
          Call GRAREG("TAOBSOLF", BB$, JJ&)
        End If
      End If
   Wend
   Close #N1%
   '
   N1% = FILEOPEN%("F:\CLIENTES\SABO_WIN\DATOS\OBSOLERV.PRN", 1, 128)
   While Not EOF(N1%)
      Line Input #N1%, A$
      CODD$ = TRIM$(Mid$(A$, 1, 10))
      CIXI% = CODINT%(CODD$)
      CAXS$ = TRIM$(Mid$(A$, 68, 10))
      Call REPLACAR(CAXS$, " ", "")
      CANN# = Val(CAXS$)
      If CIXI% > 0 Then
        If CANN# > 0.05 Then
          BB$ = MKI$(CIXI%) + MKD$(CANN#) + "BAJA AL 04/03/03"
          JJ& = JJ& + 1
          Call GRAREG("TAOBSOLF", BB$, JJ&)
        End If
      End If
   Wend
   Close #N1%
   '
   Call SETULTREG("TAOBSOLF", JJ&)
   Call CIERRARCH("TAOBSOLF")
   Close #N1%
   '
End Sub

Private Sub Command9_Click()
    FIN& = ULTREG&("FALTAVAL")
    QQ& = 0
    For U& = 1 To FIN&
      XX$ = REGLEIDO$("FALTAVAL", U&)
      CI% = CVI(Left$(XX$, 2))
      AA1$ = FILTERGETRECORD$("CPAJU_04", 1, MKI$(CI%))
      If CVI(Left$(AA1$, 2)) = CI% Then
        VALOR = CVD(Mid$(AA1$, 9, 8))
        If VALOR > 0 Then
          QQ& = QQ& + 1 'AAA = BBB
        End If
      End If
    Next U&
End Sub

Private Sub Form_Load()
    '
    VERANTER$ = "INVENTAR"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = EMPREAC$
    If EPAC$ <> "" Then
      Caption = Caption + " - " + EMPREAC$
    End If
    '
    FECHAINV$ = TRIM$(CONTROL$("INVENAJU", "FECHAINV"))
    If FECHANUM(FECHAINV$) < 1 Then
        FECHAINV$ = ""
    End If
    '
    Text1.TEXT = FECHAINV$
    Text2.TEXT = TRIM$(CONTROL$("INVENAJU", "FEGENINV"))
    Label3 = CONTROL$("INVENAJU", "TIVAPROP")
    '
    Show
    Refresh
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Text1_GotFocus()
    Command7.SetFocus
End Sub

Private Sub Text2_gotfocus()
    On Error Resume Next
    Command7.SetFocus
    On Error GoTo 0
End Sub

Sub MATVAMEN()
    If Option1.Value = True Then
       Call MENSERR(24, "Imposible Emitir Listado General.\  \Elija 'Reventa' o 'Propia Producción'")
       Exit Sub
    End If
    '
    If Option2.Value = True Then
        Screen.MousePointer = 11
        Call FILESORT("COSMEIMP", "", 1, 1, "ASC")
        Call HEADERS("COSMEIMP", "")
        Call HEADERS("COSMEIMP", "Corresponde a: Material de Reventa - Año 20" + Right$(TRIM$(Text1), 2))
        Screen.MousePointer = 1
        Call FINAL("*COSMEIMP")
      ElseIf Option3.Value = True Then
        Screen.MousePointer = 11
        STANALT$ = Right$(Text1, 2)
        ARESAN$ = "CPROP_" + STANALT$
        If Left$(UCase$(TRIM$(Label3)), 1) = "A" Then
            ARESAN$ = "CPAJU_" + STANALT$
            Call FILESORT(ARESAN$, "COSPRAJU", 1, 1, "ASC")
            Call HEADERS("COSPRAJU", "")
            Call HEADERS("COSPRAJU", "Corresponde a: Material de P.Prod - Costos Pre-Ajustados - Año 20" + Right$(TRIM$(Text1), 2))
            Screen.MousePointer = 1
            Call FINAL("*COSPRAJL")
          Else
            Call FILESORT(ARESAN$, "COSPROPR", 1, 1, "ASC")
            Call HEADERS("COSPROPR", "")
            Call HEADERS("COSPROPR", "Corresponde a: Material de P.Prod - Costos Historicos - Año 20" + Right$(TRIM$(Text1), 2))
            Screen.MousePointer = 1
            Call FINAL("*COSPROPL")
        End If
    End If
    '
End Sub
