VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CONFABPED 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuraciones  "
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Moldes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   4080
      TabIndex        =   23
      Top             =   1920
      Width           =   4035
      Begin VB.CommandButton Command10 
         Caption         =   "Asignación de dias de Secado por Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   42
         ToolTipText     =   "Asignación de dias de Secado por Código"
         Top             =   1560
         Width           =   3855
      End
      Begin VB.TextBox TxtDiasSecado 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3360
         TabIndex        =   35
         Text            =   " "
         Top             =   1920
         Width           =   585
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Tabla de  Listas de Precios Vs. Calculo de M2"
         BeginProperty Font 
            Name            =   "Arial"
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
         Top             =   1260
         Width           =   3855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Asignar Tipo de Moldes a Codigo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   3855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Configuración de Moldes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   3855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Asignar Cantidad de Moldes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   660
         Width           =   3855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Días de Secado X  Defecto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   36
         Top             =   2040
         Width           =   3075
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dias No Laborables -  Fijos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   210
      TabIndex        =   22
      Top             =   1920
      Width           =   3795
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sábado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   33
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Viernes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   32
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Jueves"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Miércoles"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Martes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lunes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Domingo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Listados de Control"
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
      Left            =   4830
      TabIndex        =   13
      Top             =   5670
      Width           =   4005
      Begin VB.CommandButton Command7 
         Caption         =   "Análisis A-B-C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2205
         TabIndex        =   19
         Top             =   1890
         Width           =   1590
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Estadísticas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   210
         TabIndex        =   18
         Top             =   1890
         Width           =   1590
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Críticos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2205
         TabIndex        =   17
         Top             =   1155
         Width           =   1590
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Consignación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   210
         TabIndex        =   16
         Top             =   1155
         Width           =   1590
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Negativos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2205
         TabIndex        =   15
         Top             =   420
         Width           =   1590
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Inventarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   210
         TabIndex        =   14
         Top             =   420
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00004000&
      Height          =   5000
      Left            =   8160
      ScaleHeight     =   4935
      ScaleWidth      =   1005
      TabIndex        =   10
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command19 
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
         Picture         =   "CONFABPED.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Terminar - Finalizar la Aplicación"
         Top             =   165
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
         Picture         =   "CONFABPED.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2040
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "CONFABPED.frx":0454
         Top             =   2760
         Width           =   1515
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Planilla de Carga"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   4080
      TabIndex        =   5
      Top             =   90
      Width           =   4005
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1545
         TabIndex        =   40
         Text            =   " "
         Top             =   1080
         Width           =   1380
      End
      Begin VB.CommandButton Command6 
         Caption         =   "."
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
         Left            =   2910
         TabIndex        =   39
         Top             =   1080
         Width           =   175
      End
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   3195
         Picture         =   "CONFABPED.frx":2376
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1080
         Width           =   550
      End
      Begin VB.TextBox Text3 
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
         Left            =   255
         TabIndex        =   37
         Text            =   " "
         Top             =   1440
         Width           =   2850
      End
      Begin VB.TextBox Text2 
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
         Left            =   260
         TabIndex        =   20
         Text            =   " "
         Top             =   675
         Width           =   2850
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
         Height          =   630
         Left            =   3200
         Picture         =   "CONFABPED.frx":2680
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   315
         Width           =   550
      End
      Begin VB.CommandButton Command21 
         Caption         =   "."
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
         Left            =   2910
         TabIndex        =   7
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1550
         TabIndex        =   6
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -840
         TabIndex        =   41
         Top             =   1170
         Width           =   2325
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -840
         TabIndex        =   9
         Top             =   405
         Width           =   2325
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   210
      TabIndex        =   0
      Top             =   90
      Width           =   3795
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "General de Planta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   210
         TabIndex        =   1
         Top             =   315
         Width           =   3375
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   240
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   2
            Top             =   360
            Width           =   2430
            Begin VB.CommandButton Command17 
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
               Picture         =   "CONFABPED.frx":298A
               Style           =   1  'Graphical
               TabIndex        =   3
               ToolTipText     =   " "
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Calendario de Dias Feriados"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   4
               Top             =   105
               Width           =   1590
            End
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7800
      OleObjectBlob   =   "CONFABPED.frx":2C94
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   4080
      OleObjectBlob   =   "CONFABPED.frx":2EC8
      TabIndex        =   21
      Top             =   420
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "CONFABPED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Command1_Click()
   Call TRALOCAL("MOLDES", "")
   VTB = VENTABU%("MOLDES")
   Call TRACENTRAL("MOLDES", "")
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    Screen.MousePointer = 11
20  Campos$ = "CODIGOINTERNO/A0;Código/A20;Descripción/A48;T.Secado/F14"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1, 0, ANTOT, , 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1, 0, ANTOT, , 1)
    GRILLAGRAL.Command3.Caption = "Asignar Tiempo de Secado"
    FIN& = CantRegistros&("MASTER")
    'CARGO LOS ITEMS CON LAS FAMILIAS
30  SQLX$ = " SELECT CODINT,CODIGO, DESCRIPCION, COLOR FROM MASTER WITH(NOLOCK)"
    SQLX$ = SQLX$ + " ORDER BY CODIGO ASC "
    Set rs1 = READSET(SQLX$)
    J& = 0
    With rs1
      Do While Not .EOF
        J& = J& + 1
        Call TRACE(20, J&, FIN&)

        CI1% = XVALO(!CODINT)
        CODIGOX$ = SAFETEXT$(!Codigo)
        DESCRI$ = SAFETEXT$(!Descripcion)
        TSECADO% = XVALO(!COLOR)
        '
        GRILLAGRAL.GRID.Rows = J& + 1
        GRILLAGRAL.GRID.TextMatrix(J&, 1) = CI1%
        GRILLAGRAL.GRID.TextMatrix(J&, 2) = CODIGOX$
        GRILLAGRAL.GRID.TextMatrix(J&, 3) = DESCRI$
        GRILLAGRAL.GRID.TextMatrix(J&, 4) = SAFETEXT$(TSECADO%)
        .MoveNext
      Loop
    End With
    rs1.Close
    Set rs1 = Nothing
    Screen.MousePointer = 1
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.mnuMenu2.Caption = "Asignar Dias de Secado a Items Seleccionados"
    GRILLAGRAL.mnuMenu2.Visible = True
    GRILLAGRAL.mnuMenu3.Caption = "Seleccionar Todo/Revertir Selección"
    GRILLAGRAL.mnuMenu3.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    'Call POSICIONAR_ULTIMAFILA(GRILLAGRAL.GRID)

    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    Set MSF = GRILLAGRAL.GRID
    If TRIM$(GRILLAGRAL.TXTBUSCAR) <> "" Then Set MSF = GRILLAGRAL.MSF_2

    If GRILLAGRAL.MENU1 > 0 Then
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "ASIGNASEC", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If
    HUBOSELECCION% = 0
    If GRILLAGRAL.MENU2 > 0 Then
50     ABUSQUE$ = UCase$(TRIM$(InputBox$("Ingrese la Cantidad de Días", , ARBUS$)))
       If ABUSQUE$ <> "" Then
          If XVALO(ABUSQUE$) < 1 Then
             Call COMUNI("Debe Ingresar un Valor Numérico")
             ARBUS$ = ABUSQUE$
             GoTo 50
          Else
            For i& = 1 To MSF.Rows - 1
              If TRIM$(MSF.TextMatrix(i&, 0)) = "*" Then
               CONDI$ = "CODINT = " & XVALO(MSF.TextMatrix(i&, 1))
               Call ACTUREGISTRO("MASTER", "COLOR", CONDI$, XVALO(ABUSQUE$), REGAF&, "NOMESS")
               HUBOSELECCION% = 1
              End If
            Next i&

            If HUBOSELECCION% > 0 Then Call COMUNI("Se ha Asingado la Familia a los Items Seleccionados")
            GRILLAGRAL.MSF_2.Clear
            GRILLAGRAL.GRID.Clear
            GRILLAGRAL.TXTBUSCAR = ""
            GRILLAGRAL.CMBFILTRO.Clear
            GoTo 20
          End If
       ElseIf GRILLAGRAL.MENU3 > 0 Then ' EDICIÓN PARA QUE PINTE LAS FILAS
           FIN& = MSF.Rows - 1
           MSF.Redraw = False
           For i& = 1 To FIN&
               Call TRACE(20, i&, FIN&)
               Call COLOREAR_GRILLA_SOLAFILA(MSF, vbYellow, i&)
           Next i&
           MSF.Redraw = True
       Else
         GoTo 99
       End If
    
    End If
    

    GoTo 99

    
99  Screen.MousePointer = 1
    Command10.Enabled = True
End Sub


Sub Command17_Click()
    Command17.Enabled = False
    Call FINAL("*CALFERI")
    Command17.Enabled = True
End Sub


Private Sub Command19_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub





Private Sub Command2_Click()
   VTB% = VENTABU%("MOLCANTI")
   Dim LEIDO%()
   'Verifica si el MOLDE esta repetido
   FIN& = ULTREG("!MOLCANTI")
   For i& = 1 To (FIN&)
      X$ = REGLEIDO$("!MOLCANTI", i&)
      MOL% = CVI(Mid$(X$, 1, 2))
      If MOL% > MAXX% Then
        ReDim Preserve LEIDO%(MOL%)
        MAXX = MOL%
      End If
      If LEIDO%(MOL%) > 0 Then
        Call COMUNI("ID " & MOL% & " Repetido")
        Erase LEIDO(): MAXX% = 0
        
        
      Else
        LEIDO%(MOL%) = 1
      End If
      
    Next i&

   Call TRACENTRAL("MOLCANTI", "")

End Sub

Private Sub Command21_Click()
    Command21.Enabled = False
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
    Command21.Enabled = True
End Sub
'
Private Sub Command22_Click()
    Command22.Enabled = False
    ACONEC$ = "FLFORMS"
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
    Command22.Enabled = True
End Sub

Private Sub Command23_Click()
    '
    For i% = 0 To Me.Check1.UBound
      If Check1(i%).Value = 1 Then
        NOLABORABLES$ = NOLABORABLES$ & i% & ";" '0;1
      End If
    Next i%
'    'LE QUITO LA ULTIMA COMA
'    If Right$(NOLABORABLES$, 1) = ";" Then
'       NOLABORABLES$ = Left$(NOLABORABLES$, Len(NOLABORABLES$) - 1)
'    End If
    '
    Call GRACONTROL("CONFABPE", "NOLABORA", NOLABORABLES$)
    Call GRACONTROL("CONFABPE", "PLANPROD", Text1)
    Call GRACONTROL("CONFABPE", "ETIQPALL", Text4)
    '
    VALO$ = TRIM$(Me.TxtDiasSecado)
    If XVALO(VALO$) <= 0 Then VALO$ = "1"
    Call GRACONTROL("CONFABPE", "DIASSECA", TRIM$(Me.TxtDiasSecado))
    '
    Call CIERRARCH("*.*")
    Call Command19_Click
    '
End Sub
Private Sub Command3_Click()
    Screen.MousePointer = 11
20  Campos$ = "CODIGOINTERNO/F0;Código/A20;Descripción/A48;C.Molde./f10;Descripción/A32"
    Call CARGA_ENCABEZADO(SELFAMILI.GRID, Campos$, SELFAMILI, 1, 0, ANTOT, , 1)
    Call CARGA_ENCABEZADO(SELFAMILI.MSF_2, Campos$, SELFAMILI, 1, 0, ANTOT, , 1)
    SELFAMILI.Command3.Caption = "Asignar Molde"
    FIN& = CantRegistros&("MASTER")
    'CARGO LOS ITEMS CON LAS FAMILIAS
30  SQLX$ = " SELECT CODINT,CODIGO, DESCRIPCION, MOLDE FROM MASTER WITH(NOLOCK)"
    SQLX$ = SQLX$ + " ORDER BY CODIGO ASC "
    Set rs1 = READSET(SQLX$)
    J& = 0
    With rs1
      Do While Not .EOF
        J& = J& + 1
        Call TRACE(20, J&, FIN&)

        CI1% = XVALO(!CODINT)
        CODIGOX$ = SAFETEXT$(!Codigo)
        DESCRI$ = SAFETEXT$(!Descripcion)
        Moldex% = XVALO(!MOLDE)
        DESCRIMOLDE$ = ""
        If Moldex% > 0 Then DESCRIMOLDE$ = ECOARCH$("MOLDES", MKI$(Moldex%))
        '
        SELFAMILI.GRID.Rows = J& + 1
        SELFAMILI.GRID.TextMatrix(J&, 1) = CI1%
        SELFAMILI.GRID.TextMatrix(J&, 2) = CODIGOX$
        SELFAMILI.GRID.TextMatrix(J&, 3) = DESCRI$
        SELFAMILI.GRID.TextMatrix(J&, 4) = SAFETEXT$(Moldex%)
        SELFAMILI.GRID.TextMatrix(J&, 5) = DESCRIMOLDE$
        .MoveNext
      Loop
    End With
    rs1.Close
    Set rs1 = Nothing
    If J& > 0 Then
       Screen.MousePointer = 1
       SELFAMILI.GRID.Visible = True
       SELFAMILI.GRID.ZOrder 0
       Screen.MousePointer = 1
50     SELFAMILI.Show 1
       If SELFAMILI.HaySeleccion > 0 Then ' SI HAY ITEMS SELECCIONADOS
            Call SELECHO("MOLDES")
            Moldex% = XVALO(VALACT1$("MOLDES")) ' OJO NO MODIFICAR SELFAMILI POR QUE SE USA EN OTRAS  APLICACIONES.)
            If Moldex% < 1 Then Screen.MousePointer = 1: GoTo 99 ' SI NO ELIGE MOLDES NO GRABA SALE
            For i& = 1 To SELFAMILI.MSFACTIVO.Rows - 1
              If TRIM$(SELFAMILI.MSFACTIVO.TextMatrix(i&, 0)) = "*" Then
               CONDI$ = "CODINT = " & XVALO(SELFAMILI.MSFACTIVO.TextMatrix(i&, 1))
               Call ACTUREGISTRO("MASTER", "Molde", CONDI$, Moldex%, REGAF&, "NOMESS")
              End If
            Next i&
            Call COMUNI("Se ha Asingado la Familia a los Items Seleccionados")
            SELFAMILI.HaySeleccion = 0
            Screen.MousePointer = 11
            SELFAMILI.MSF_2.Clear
            SELFAMILI.GRID.Clear
            SELFAMILI.TXTBUSCAR = ""
            SELFAMILI.CMBFILTRO.Clear
            GoTo 20
       End If
    Else
       Call COMUNI("No se Han Encontrado Items")
       Screen.MousePointer = 1
    End If
    
99  Screen.MousePointer = 1

End Sub


Private Sub Command4_Click()
   Command4.Enabled = False
   Call TRALOCAL("LISTASM2", "")
   VTB = VENTABU%("LISTASM2")
   Call TRACENTRAL("LISTASM2", "")
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    ACONEC$ = "FLFORMS"
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text4.TEXT), "Formularios de Impresion")
    Command5.Enabled = True

End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text4.TEXT = VDEVU$
        Text3.TEXT = VALACT2$("INDIFRM")
    End If
    Command6.Enabled = True

End Sub

Private Sub Form_Load()
    '
    Call APLICACIONSKINS(Me, Picture8)
    '
    Call ABREMASTER
    'DESMARCO TODAS LAS OPCIONES
    For i% = 0 To Me.Check1.UBound
      Check1(i%).Value = False
    Next i%
    '
    'MARCO LAS GRABADAS.
    XX$ = TRIM$(Control("CONFABPE", "NOLABORA"))
    If XX$ <> "" Then
      Call TEXTOLOTES$(XX$, ";")
      For i% = 1 To UBound(CADENATEX$)
       IND% = XVALO(CADENATEX$(i%))
       Check1(IND%).Value = 1
      Next i%
    End If
    'DIAS DE SECADO
    Me.TxtDiasSecado = TRIM$(Control("CONFABPE", "DIASSECA"))
    '
    Text1 = Control("CONFABPE", "PLANPROD")
    Text4 = Control("CONFABPE", "ETIQPALL")

End Sub

Private Sub Text1_Change()
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub



Private Sub Text4_Change()
   Text3.TEXT = ECOARCH$("INDIFRM", Text4.TEXT)
End Sub
