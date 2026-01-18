VERSION 5.00
Begin VB.Form OPREMSTA04 
   BackColor       =   &H00D0F0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remitos de Expedición - Opciones"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4995
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Repetir Comprobantes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Index           =   1
      Left            =   210
      TabIndex        =   35
      Top             =   5565
      Width           =   3690
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
         Height          =   600
         Left            =   760
         Picture         =   "OPREMSTA04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Consulta y Re-Impresión de Remitos"
         Top             =   350
         Width           =   560
      End
      Begin VB.PictureBox Picture12 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   720
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   36
         Top             =   315
         Width           =   2430
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Consulta y Re-Impresión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   860
            TabIndex        =   37
            Top             =   70
            Width           =   1725
         End
      End
   End
   Begin VB.CommandButton Command11 
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
      Height          =   630
      Left            =   210
      Picture         =   "OPREMSTA04.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Indice de Cumplimiento de Entregas"
      Top             =   6930
      Width           =   650
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00006060&
      Height          =   7260
      Left            =   4095
      ScaleHeight     =   7200
      ScaleWidth      =   4215
      TabIndex        =   21
      Top             =   -420
      Width           =   4275
      Begin VB.CommandButton Command15 
         Caption         =   "Depur"
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
         Left            =   105
         Picture         =   "OPREMSTA04.frx":0F34
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Depurar Remitos Por Rango de Fechas "
         Top             =   5160
         Width           =   650
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Check"
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
         Left            =   105
         Picture         =   "OPREMSTA04.frx":107E
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Control de Integridad de la B.Datos"
         Top             =   3580
         Width           =   650
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Import"
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
         Left            =   105
         Picture         =   "OPREMSTA04.frx":1388
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Importación de Remitos"
         Top             =   2890
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Setup"
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
         Left            =   105
         Picture         =   "OPREMSTA04.frx":14D2
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Configuracion de Funcionamiento"
         Top             =   4410
         Width           =   650
      End
      Begin VB.CommandButton Command7 
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
         Height          =   690
         Left            =   105
         Picture         =   "OPREMSTA04.frx":1914
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1390
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
         Height          =   690
         Left            =   105
         Picture         =   "OPREMSTA04.frx":1C1E
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   700
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   23
         Top             =   6405
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Equiv"
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
         Left            =   105
         Picture         =   "OPREMSTA04.frx":1D68
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Códigos Equivalentes de Clientes"
         Top             =   2200
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "OPREMSTA04.frx":1EB2
         Top             =   6720
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Resúmenes de  Entregas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Index           =   2
      Left            =   210
      TabIndex        =   4
      Top             =   2625
      Width           =   3690
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   15
         Top             =   1995
         Width           =   2430
         Begin VB.CommandButton Command8 
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
            Picture         =   "OPREMSTA04.frx":3DD4
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Consulta / Listado de Material en Consignación (Clientes)"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Entregas Pendien- tes de Facturar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   735
            TabIndex        =   17
            Top             =   105
            Width           =   1605
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   11
         Top             =   1470
         Width           =   2430
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
            Height          =   600
            Left            =   0
            Picture         =   "OPREMSTA04.frx":40DE
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Listado Secuencial por Número de Comprobante"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Secuencial por Nro. de Comprobante"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   630
            TabIndex        =   13
            Top             =   55
            Width           =   1750
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   5
         Top             =   945
         Width           =   2430
         Begin VB.CommandButton Command3 
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
            Left            =   -20
            Picture         =   "OPREMSTA04.frx":43E8
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Listado de Entregas por Cliente "
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Resumen de Entre- gas por Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   630
            TabIndex        =   7
            Top             =   55
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   8
         Top             =   420
         Width           =   2430
         Begin VB.CommandButton Command4 
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
            Picture         =   "OPREMSTA04.frx":482A
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Listado de Entregas por Artículo "
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Resumen de Entre- gas por Producto"
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
            Left            =   630
            TabIndex        =   10
            Top             =   55
            Width           =   1725
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Operaciones Complementarias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   3690
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   29
         Top             =   1470
         Width           =   2640
         Begin VB.CommandButton Command12 
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
            Left            =   2030
            Picture         =   "OPREMSTA04.frx":4B34
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Re-Asignación de Remitos a Pedidos"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command10 
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
            Picture         =   "OPREMSTA04.frx":4C7E
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Corrección Cantidades Despachadas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Corrección de Remitos"
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
            Left            =   800
            TabIndex        =   31
            Top             =   105
            Width           =   1005
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   18
         Top             =   945
         Width           =   2640
         Begin VB.CommandButton Command9 
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
            Picture         =   "OPREMSTA04.frx":4F88
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Devolución Material en Consignación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Devolución de Despa- chos de Expedición"
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
            Left            =   630
            TabIndex        =   19
            Top             =   45
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2595
         TabIndex        =   1
         Top             =   420
         Width           =   2655
         Begin VB.CommandButton Command6 
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
            Left            =   560
            Picture         =   "OPREMSTA04.frx":50D2
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Eliminación Solo de Lista de Pendientes de Factura"
            Top             =   0
            Width           =   570
         End
         Begin VB.CommandButton Command25 
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
            Picture         =   "OPREMSTA04.frx":53DC
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Anulación General de Remito / Orden de Despacho Emitidos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación de Comprobantes"
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
            Left            =   1260
            TabIndex        =   3
            ToolTipText     =   "Click Aqui Importa Archivo GETRAMO"
            Top             =   45
            Width           =   1740
         End
      End
   End
End
Attribute VB_Name = "OPREMSTA04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Command1_Click()
    '
    Command1.Enabled = False
    OPREMIT04.Show 1
    Command1.Enabled = True
    '
End Sub

Sub Command10_Click()
   Command10.Enabled = False
   If DERACCE%("CORRECRE", "Correccion Cantidades por Remito") >= 2 Then
     '
     CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
     If Len(CMO$) <> 2 Then
        CMO$ = "RT"
     End If
     '
     Call CARSELMESA
     '
10   Screen.MousePointer = 1
     Call SELECHO("SELMESA")
     MESASEL$ = VALACT1$("SELMESA")
     If TRIM$(MESASEL$) = "" Then
        GoTo 99
     End If
     CORRESQ$ = MESANO$(MESASEL$, DES%, HAS%)
     '
     Call FRESHECHO("SELREMI")
     Screen.MousePointer = 11
     '
     DES1# = CDbl(CVDAT(DES%))
     HAS1# = CDbl(CVDAT(HAS%))
     '
     SQLX$ = " SELECT DISTINCT DoPriCor, Nume_Clie, FechMov FROM MOVISTO "
     SQLX$ = SQLX$ + " WHERE INT(CDBL(FechMov)) >= " & DES1# & " "
     SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS1# & " "
     'PARA EVITAR COMPROBANTES ARRUINADOS
     SQLX$ = SQLX$ + " AND LEFT$(UCASE$(REFERCOR),4) <> 'ANUL' "
     SQLX$ = SQLX$ + " AND CodiMovi =  '" & CMO$ & "'  "
     '
     Call APERBASDAT("STOCKS")
     Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With MOVITMP
       On Error Resume Next
       .MoveFirst
       If Err Then
         Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
         On Error GoTo 0
         GoTo 10
       End If
       On Error GoTo 0
       '
       J& = 0
       Call FRESHECHO("!SELREMI")
       Do While (.EOF = False)
         CLI% = !Nume_Clie
         REMI$ = !DoPriCor
         FF% = CVINT(!FECHMOV)
         '\\\OT-06-0083-WAR-07/03/06///
         F1$ = FECHATEX$(FF%)
         REFE1$ = F1$ + " - " + RASOCLI$(CLI%)
         '\\\OT-06-0083-WAR-FIN///
         aa2 = !Nume_Clie
         z$ = Space$(40) + String$(8, 0)
         Call REPLA(z$, REMI$, 1, 10)
         Call REPLA(z$, REFE1$, 11, 30) 'RASOCLI$(CLI%), 11, 30)
         Call REPLA(z$, MKS$(U&), 41, 4)
         Call REPLA(z$, MKI$(FF%), 45, 2)
         Call REPLA(z$, MKI$(CLI%), 47, 2)
         J& = J& + 1
         Call GRAREG("!SELREMI", z$, J&)
       .MoveNext
       Loop
     End With
     '
     Call SETULTREG("!SELREMI", J&)
     Call CIERRARCH("!SELREMI")
     Screen.MousePointer = 1
     '
     If J& < 1 Then
        Call MENSERR(24, "No se Encontraron Remitos\Posibles de Corregir Cantidades\en el Período Elegido.")
        GoTo 10
     End If
     '
15   RERE$ = REGSEL$("!SELREMI")
     If Len(RERE$) <= 4 Then GoTo 10
     '
     REMISEL$ = TRIM$(Left$(RERE$, 10))
     CLIE$ = Mid$(RERE$, 11, 30)
     FEXI% = CVI(Mid$(RERE$, 45, 2))
     FEX = FEXI%
     FEMOVD = Int(CDbl(CVDAT(FEXI%)))
     CLISEL% = CVI(Mid$(RERE$, 47, 2))
     '
20   Call APERBASDAT("STOCKS")
     Screen.MousePointer = 11
     SQLX$ = "SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + "WHERE DoPriCor = '" & REMISEL$ & "' "
     SQLX$ = SQLX$ + "AND Nume_Clie = " & CLISEL% & " "
     '\\\OT-06-0083-WAR-07/03/06///
     SQLX$ = SQLX$ + " AND cdbl(FechMov) = " & FEMOVD & " "
     '\\\OT-06-0083-WAR-FIN///
     SQLX$ = SQLX$ + "ORDER BY NuOrItem ASC"
     Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
     '
     NORIMAX% = 0
     With MOVITMP
       On Error Resume Next
       .MoveLast
       If Err Then
         Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
         On Error GoTo 0
         GoTo 10
       End If
       On Error GoTo 0
       '
       NORIMAX% = XVALO(!NuOrItem)
       DOPRI$ = REMISEL$
       DOPRILA$ = SAFETEXT$(!DoPriLar)
       DOSEC$ = ""       ' QUEDA LIBRE PARA PONERLE EL NUMERO DE FACTURA
       DOSELA$ = ""      ' QUEDA LIBRE PARA PONERLE EL NUMERO DE FACTURA
       MONEVA% = XVALO(!MoneVal)
       DEPOMO% = XVALO(!DEPOMOVI)
       REFEX8$ = SAFETEXT$(!ReferCor)
       '
       HOII% = HOY(HOS$)
       '
       .MoveFirst
       J& = 0: NORI$ = ""
       Do While (.EOF = False)
         NORIT% = XVALO(!NuOrItem)
         If NORIT% < 1 Then
            NORIT% = NORIMAX% + 1
            .Edit
            !NuOrItem = NORIT%
            .Update
            NORIMAX% = NORIT%
            GoTo 29
         End If
         '
         For KKI% = 1 To Len(NORI$) Step 2
           If CVI(Mid$(NORI$, KKI%, 2)) = NORIT% Then
               NORIT% = NORIMAX% + 1
               .Edit
               !NuOrItem = NORIT%
               .Update
               NORIMAX% = NORIT%
               GoTo 29
           End If
         Next KKI%
         '
29       NORI$ = NORI$ + MKI$(NORIT%)
         '
         If Abs(XVALO(!CANTINGRE) - XVALO(!CANTSALID)) >= 0.005 Then
           XX$ = REGBLAN$("BOLREDET")
           Call REPLA(XX$, MKI$(XVALO(!NuOrItem)), 105, 2)
             If XVALO(!NuOrItem) > NORIMAX% Then NORIMAX% = XVALO(!NuOrItem)
           Call REPLA(XX$, MKI$(XVALO(!COD_INTE)), 83, 2)
           Call REPLA(XX$, UNIMED$(XVALO(!COD_INTE)), 85, 4)
           Call REPLA(XX$, MKD$(XVALO(!CANTSALID) - XVALO(!CANTINGRE)), 89, 8)
           Call REPLA(XX$, SAFETEXT$(!IdenLote), 119, 16)
             VUNIT = XVALO(!ValoUnit)
             NUPEDI& = XVALO(!NUPEDCLI) 'NUMERO DE PEDIDO
           Call REPLA(XX$, MKS$(NUPEDI&), 151, 4) 'Nº PEDIDO.
           Call REPLA(XX$, MKS$(VUNIT), 159, 4)
           Call REPLA(XX$, MKD$(0), 171, 8)
           J& = J& + 1
           Call GRAREG("!BOLREDET", XX$, J&)
         End If
       .MoveNext
       Loop
     End With
     Call SETULTREG("!BOLREDET", J&)
     Screen.MousePointer = 1
     '
30   FETEXI$ = REPLACAR$(FECHATEX$(FEX), "/", "-")
     VTB% = VENTABU%("CORREMIT/TITUPAN=Corrección " + REMISEL$ + " - " + TRIM$(CLIE$) + " - " + FETEXI$)
     If VTB% < 0 Then GoTo 15
     '
     Screen.MousePointer = 11
     Call APERBASDAT("STOCKS")
     Screen.MousePointer = 11
     J& = ULTREG&("!BOLREDET")
     For KKIL& = 1 To J&
       XX$ = REGLEIDO$("!BOLREDET", KKIL&)
       CIXI% = CVI(Mid$(XX$, 83, 2))
       LOTE$ = Mid$(XX$, 119, 16)
       NUORIT% = CVI(Mid$(XX$, 105, 2))
       CACORREC = CVD(Mid$(XX$, 89, 8))
       NUPEDID1& = CVS(Mid$(XX$, 151, 4))
       VUNIT = CVS(Mid$(XX$, 159, 4))
       DOSELA$ = "P/F - PC." + TRIM$(Str$(NUPEDID1&))
       DOSEC$ = "PC." + TRIM$(Str$(NUPEDID1&))
       '
       SQLX$ = " DoPriCor = '" & REMISEL$ & "' "
       SQLX$ = SQLX$ + " AND Nume_Clie = " & CLISEL% & " "
       SQLX$ = SQLX$ + " AND NuOrItem = " & NUORIT% & ""
       '
       With MOVITMP ' Movisto
         If NUORIT% < 1 Then   ' para tener en cuenta items agregados
           If Abs(CACORREC) >= 0.05 Then
             .AddNew
             !FECHMOV = CVDAT(FEXI%)
             !COD_INTE = CIXI% 'GUARDA EL CODIGO INTERNO EDITADO
             !COD_EXTE = CODEXT(CIXI%) 'GUARDA EL COD. EXT EDITADO
             !Descrip = Left$(DESCRIT(CIXI%), 64) 'GUARDA LA DESCRIPCION.
             !IdenLote = Left$(LOTE$, 16) 'GUARDA EL LOTE EDITADO
             !CodiMovi = CMO$
             !DoPriCor = Left$(DOPRI$, 12)
             !DoPriLar = Left$(DOPRILA$, 24)
             !DoSecCor = Left$(DOSEC$, 12)
             !DoSecLar = Left$(DOSELA$, 24)
             !NumRem_Cli = ""
               NORIMAX% = NORIMAX% + 1
             !NuOrItem = NORIMAX%
             !NUPEDCLI = NUPEDID1& 'GUARDA EL Nº PED. EDITADO
             !ReferCor = Left$("(Agr.Man) " + REFEX8$, 64)
             !ValoUnit = VUNIT 'GUARDA EL VALOR UNITARIO EDITADO
             !MoneVal = MONEVA%
             !DEPOMOVI = DEPOMO%
             !Nume_Clie = CLISEL%
             !Nume_Prov = 0
             !CANTSALID = 0: !CANTINGRE = 0
             If CACORREC > 0 Then
                 !CANTSALID = CACORREC 'GUARDA LA CANTIDAD EDITADA
               ElseIf CACORREC < 0 Then
                 !CANTINGRE = Abs(CACORREC)
             End If
             !CantSaldo = 0
             !FeReMovi = Now
             !NumUsuar = USNBR%
             .Update
           End If
           GoTo 39
         End If
         '
         .FindFirst SQLX$
         If .NoMatch = False Then
           .Edit
            If !CodiMovi = CMO$ Then
               If CACORREC >= 0 Then
                   !CANTSALID = CACORREC 'GUARDA LA CANTIDAD EDITADA
                   !CANTINGRE = 0
                   GoTo 35
                 Else
                   Call MENSERR(24, "Imposible Corregir Item Número " + TRIM$(Str$(NUORIT%)))
                   GoTo 39
               End If
            End If
            '
            If !CodiMovi = "DV" Then
               If CACORREC <= 0 Then
                   !CANTINGRE = Abs(CACORREC) 'GUARDA LA CANTIDAD EDITADA
                   !CANTSALID = 0
                   GoTo 35
                 Else
                   Call MENSERR(24, "Imposible Corregir Item Número " + TRIM$(Str$(NUORIT%)))
                   GoTo 39
               End If
            End If
            '
            If Abs(CACORREC) > 0.005 Then
               If Sgn(CACORREC) <> Sgn(!CANTSALID - !CANTINGRE) Then
                  Call MENSERR(24, "Imposible Corregir Item Número " + TRIM$(Str$(NUORIT%)))
                  GoTo 39
               End If
               If CACORREC > 0 Then
                   !CANTSALID = CACORREC 'GUARDA LA CANTIDAD EDITADA
                   !CANTINGRE = 0
                   GoTo 35
                 Else
                   !CANTINGRE = Abs(CACORREC)
                   !CANTSALID = 0
                   GoTo 35
               End If
            End If
            '
            If Abs(CACORREC) < 0.005 Then
               !CANTSALID = 0: !CANTINGRE = 0
            End If
            '
35          !ValoUnit = VUNIT 'GUARDA EL VALOR UNITARIO EDITADO
            !IdenLote = LOTE$ 'GUARDA EL LOTE EDITADO
            !COD_INTE = CIXI% 'GUARDA EL CODIGO INTERNO EDITADO
            !COD_EXTE = CODEXT(CIXI%) 'GUARDA EL COD. EXT EDITADO
            !NUPEDCLI = NUPEDID1& 'GUARDA EL Nº PED. EDITADO
            !Descrip = DESCRIT(CIXI%) 'GUARDA LA DESCRIPCION.
           .Update
         End If
       End With
       '
39   Next KKIL&
     'Movisto.Close
     Screen.MousePointer = 1
     '
     Call COMUNI("Corrección Procesada")
     GoTo 15
     '
   End If
   '
99 Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   Call CONECRUN("COCUMENT", "Control de Cumplimiento de Entregas")
   Command11.Enabled = True
End Sub

Sub Command12_Click()
'\\\OT-05-0437-WAR-FIN///
  If DERACCE%("CORREMIT", "Corrección de Remitos") < 2 Then
    Exit Sub
  End If
  '
  FEX = HOY(HO$)
  HAS% = FEX
  HASS$ = FECHATEX$(FEX)
  DESS$ = "01" + Mid$(HASS$, 3)
  DES% = FECHANUM(DESS$)
  '
  NCLIE% = Val(Text1)
  RECORD$ = Space(15)
  Call REPLA(RECORD$, MKI$(NCLIE%), 1, 2)
  Call REPLA(RECORD$, MKI(DES%), 3, 2)
  Call REPLA(RECORD$, MKI(HAS%), 5, 2)
  Call REPLA(RECORD$, MKS(0), 7, 4)
  Call REPLA(RECORD$, MKS(0), 11, 4)
  '
10 OBX$ = OBJPLA$("SELEREMPED", RECORD$)
  If OBX$ <> "" Then
    '
    RECORD$ = OBX$
    DES2% = CVI(Mid$(OBX$, 3, 2))
    HAS2% = CVI(Mid$(OBX$, 5, 2))
    DES3# = CDbl(CVDAT(DES2%))
    HAS3# = CDbl(CVDAT(HAS2%))
    NCLIE% = CVI(Mid$(OBX$, 1, 2))
    NPED = CVS(Mid$(OBX$, 7, 4))
    NPEDRE = CVS(Mid$(OBX$, 11, 4))
    '
    If NCLIE% < 1 Or REGICLI&(NCLIE%) < 1 Then
      Call MENSERR(24, "Cliente no Válido")
      GoTo 10
    End If
    '
    If NPED < 1 Or NPEDRE < 1 Then
      Call MENSERR(24, "Números de Pedido no Válidos")
      GoTo 10
    End If
    '
    Call APERBASDAT("PEDCLI")
    SQLX2$ = "NroPed = " & NPED & ""
    PedEnca.FindFirst SQLX2$
    If PedEnca.NoMatch = True Then
      Call MENSERR(24, "Pedido Origen no Encontrado")
      GoTo 10
    End If
    '
    SQLX2$ = "NroPed = " & NPEDRE & ""
    PedEnca.FindFirst SQLX2$
    If PedEnca.NoMatch = True Then
      Call MENSERR(24, "Pedido Reemplazo no Encontrado")
      GoTo 10
    End If
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
      CMO$ = "RT"
    End If
    Screen.MousePointer = 11
    Call APERBASDAT("STOCKS")
    '
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE Nume_Clie = " & NCLIE% & " "
    SQLX$ = SQLX$ + " AND (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & DES3# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS3# & " "
    SQLX$ = SQLX$ + " AND VAL(MID(DoSecCor,4))= " & NPED & " "
    '
    Set PEDTMP = Stocks.OpenRecordset(SQLX$, 2)
    With PEDTMP
      '
      On Error Resume Next
        .MoveFirst
      If Err Then
        On Error GoTo 0
        Call COMUNI("Ningun Registro Coincide con Condición SQL.")
        Exit Sub
      End If
      On Error GoTo 0
      '
      Do While .EOF = False
        CIXI% = !COD_INTE
        CANREM = !CANTSALID - !CANTINGRE
        '
        SQLX2$ = " CodiInt = " & CIXI% & " "
        SQLX2$ = SQLX2$ + " AND NroPed = " & NPED & ""
        Call APERBASDAT("PEDCLI")
        PedDeta.FindFirst SQLX2$
        If PedDeta.NoMatch = False Then
          TOREM = PedDeta!CantEnt - CANREM
          If TOREM < 0 Then TOREM = 0
          PedDeta.Edit
          PedDeta!CantEnt = TOREM
          PedDeta.Update
        End If
        '
        SQLX2$ = " CodiInt = " & CIXI% & " "
        SQLX2$ = SQLX2$ + " AND NroPed = " & NPEDRE & ""
        Call APERBASDAT("PEDCLI")
        PedDeta.FindFirst SQLX2$
        If PedDeta.NoMatch = False Then
          PUNI1 = PedDeta!PreUnid
          TOREM = PedDeta!CantEnt + CANREM
          PedDeta.Edit
          PedDeta!CantEnt = TOREM
          PedDeta.Update
        Else
          PUNI1 = !ValoUnit
        End If
        '
        .Edit 'RUTINA DE EDICIÓN
          !DoSecCor = "PC." & NPEDRE
          !DoSecLar = "P/F - PC." & NPEDRE
          !NUPEDCLI = NPEDRE
          !ValoUnit = PUNI1
        .Update
        '
        .MoveNext
      Loop
    End With
    Call COMUNI("Reemplazo Procesado")
  End If
  Screen.MousePointer = 1
  '
  '\\\OT-05-0556-WAR-FIN///

End Sub

Sub Command13_Click()
  '
  Command13.Enabled = False
  Call DESHASFECHA(DES%, HAS%, PERIO$)
  If TRIM$(PERIO$) = "" Then GoTo 99
  Call APERBASDAT("STOCKS")
  '
  CMO1$ = TRIM$(CONTROL$("", "CODREFAC"))
  If Len(CMO$) <> 2 Then
     CMO1$ = "RT"
  End If
  '
  DES1# = CDbl(CVDAT(DES%))
  HAS1# = CDbl(CVDAT(HAS%))
  '
  SQLX$ = " SELECT * FROM MOVISTO"
  SQLX$ = SQLX$ + " WHERE INT(CDBL(FechMov)) >= " & DES1# & " "
  SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS1# & " "
  SQLX$ = SQLX$ + " AND CodiMovi =  '" & CMO1$ & "' "
  SQLX$ = SQLX$ + " AND (CantSalid <> 0 OR CantIngre <> 0) "
  SQLX$ = SQLX$ + " ORDER BY DoPriCor ASC, Nuoritem ASC "
  Set CTRLTMP = Stocks.OpenRecordset(SQLX$, 4)
  With CTRLTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      Call COMUNI("No Existen Remitos para el Rango de Fechas Elegido")
      On Error GoTo 0
      GoTo 99
    End If
    On Error GoTo 0
    Do While .EOF = False
      If !DoPriCor = NPX1$ And !COD_INTE = CIX1% Then
        'If NPX2$ <> NPX1$ Then
          AA1 = CODEXT$(!COD_INTE)
          Call MENSERR(24, "Se ha Detectado Código '" + AA1 + "'\Duplicado en '" & NPX1$ & "' ")
        'end If
        'NPX2$ = NPX1$
      End If
      NPX1$ = !DoPriCor
      CIX1% = !COD_INTE
      .MoveNext
    Loop
  End With
  '
  Call MENSERR(24, "Fin del Control")
  '
99 Screen.MousePointer = 1
  Command13.Enabled = True
  '
End Sub

Sub Command14_Click()
    '\\\OT-08-0558-OD-29/07/08///
   Command14.Enabled = False
     '
     CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
     If Len(CMO$) <> 2 Then
        CMO$ = "RT"
     End If
     '
     Call CARSELMESA
     '
10   Screen.MousePointer = 1
     Call SELECHO("SELMESA")
     MESASEL$ = VALACT1$("SELMESA")
     If TRIM$(MESASEL$) = "" Then
        GoTo 99
     End If
     CORRESQ$ = MESANO$(MESASEL$, DES%, HAS%)
     '
     Call FRESHECHO("SELREMI")
     Screen.MousePointer = 11
     '
     DES1# = CDbl(CVDAT(DES%))
     HAS1# = CDbl(CVDAT(HAS%))
     '
     SQLX$ = " SELECT DISTINCT DoPriCor, Nume_Clie, FechMov FROM MOVISTO "
     SQLX$ = SQLX$ + " WHERE INT(CDBL(FechMov)) >= " & DES1# & " "
     SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS1# & " "
     'PARA EVITAR COMPROBANTES ARRUINADOS
     SQLX$ = SQLX$ + " AND LEFT$(UCASE$(REFERCOR),4) <> 'ANUL' "
     SQLX$ = SQLX$ + " AND CodiMovi =  '" & CMO$ & "'  "
     '
     Call APERBASDAT("STOCKS")
     Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With MOVITMP
       On Error Resume Next
       .MoveFirst
       If Err Then
         Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
         On Error GoTo 0
         GoTo 10
       End If
       On Error GoTo 0
       '
       J& = 0
       Call FRESHECHO("!SELREMI")
       Do While (.EOF = False)
         CLI% = !Nume_Clie
         REMI$ = !DoPriCor
         FF% = CVINT(!FECHMOV)
         '\\\OT-06-0083-WAR-07/03/06///
         F1$ = FECHATEX$(FF%)
         REFE1$ = F1$ + " - " + RASOCLI$(CLI%)
         '\\\OT-06-0083-WAR-FIN///
         aa2 = !Nume_Clie
         z$ = Space$(40) + String$(8, 0)
         Call REPLA(z$, REMI$, 1, 10)
         Call REPLA(z$, REFE1$, 11, 30) 'RASOCLI$(CLI%), 11, 30)
         Call REPLA(z$, MKS$(U&), 41, 4)
         Call REPLA(z$, MKI$(FF%), 45, 2)
         Call REPLA(z$, MKI$(CLI%), 47, 2)
         J& = J& + 1
         Call GRAREG("!SELREMI", z$, J&)
       .MoveNext
       Loop
     End With
     '
     Call SETULTREG("!SELREMI", J&)
     Call CIERRARCH("!SELREMI")
     Screen.MousePointer = 1
     '
     If J& < 1 Then
        Call MENSERR(24, "No se Encontraron Remitos en el Rango Seleccionado.")
        GoTo 10
     End If
     '
15   RERE$ = REGSEL$("!SELREMI")
     If Len(RERE$) <= 4 Then GoTo 10
     
      NPX& = XVALO(TRIM$(Mid$(RERE$, 6, 5)))
      '
      PVTA% = XVALO(TRIM$(Mid$(RERE$, 4, 1))) ' PUNTO DE VENTA
      PVTA1$ = TRIM$(Str$(PVTA%))
      While Len(PVTA1$) < 4: PVTA1$ = "0" + PVTA1$: Wend
      '
      TIDOCUM$ = "Remito Nro."
      FECHA1$ = Mid$(RERE$, 11, 8) ' ///OT-6-646-03/09/08
      NUDOCU$ = TRIM$(Str$(NPX&))
      While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
      NUDOCU$ = NUDOCU$ + " "
      Screen.MousePointer = 11
      '
      For UI& = ULTREG&("PDOCLST") To 1 Step -1
        XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
        '\\\OT-08-0646-OD-20/08/08///
         If InStr(XX$, TIDOCUM$) > 4 Then
            If InStr(XX$, NUDOCU$) > 4 Then
              If InStr(XX$, "." & PVTA1$) > 4 Then
                'If InStr(XX$, FECHA1$) > 4 Then ' ///OT-6-646-03/09/08
                'SE AGREGO LOS IF CON LA FECHA Y CON EL TIPO DE DOCUMENTO POR QUE SI _
                 NO TRAIA OTRO TIPO DE DOCUMENTO Y / O DE OTRA FECHA CUANDO SE REPETIA _
                 EL MISMO NUMERO DE COMPROBANTE PERO CON OTRO PUNTO DE VENTA
                 PPXX% = InStr(XX$, "Nro.")
                  If PPXX% > 0 Then
                   XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
                  End If
                  If InStr(XX$, NUDOCU$) > 4 Then
                   DOCUNU& = CVS(Left$(XX$, 4))
                   GoTo 50
                  End If
                'End If ' ///OT-6-646-03/09/08
              End If
              '\\\OT-08-0646-OD-FIN///
            End If
         End If
      Next UI&
      Screen.MousePointer = 1
40    Call COMUNI("No se Encontró el Documento Impreso")
      GoTo 15

     '
50   Call MUESTRADOC(DOCUNU&)
     GoTo 15
     
99   Command14.Enabled = True
     
    '\\\OT-08-0558-OD-FIN///
End Sub

Sub Command15_Click()
   Command15.Enabled = False
   
    Screen.MousePointer = 11
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    Call DESHASFECHA(DES%, HAS%, PERIO$)   ' SE MODIFICA PARA PODER
    '
    If PERIO$ = "" Then GoTo 99
    
    DES1# = Int(CDbl(CVDAT(DES%)))
    HAS1# = Int(CDbl(CVDAT(HAS%)))
    '
    
    Call APERBASDAT("STOCKS")
    SQLX$ = " UPDATE MOVISTO SET DoSecCor = 'S/Cargo',DoSecLar = 'Sin Cargo' "
    SQLX$ = SQLX$ + " WHERE INT(CDBL(FechMov)) >= " & DES1# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS1# & " "
    SQLX$ = SQLX$ + " AND (CodiMovi =  '" & CMO$ & "'  "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV')  "
    'PARA EVITAR COMPROBANTES ARRUINADOS
    SQLX$ = SQLX$ + " AND LEFT$(UCASE$(REFERCOR),4) <> 'ANUL' "
    SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
    '
    Stocks.Execute (SQLX$)
    '
    Screen.MousePointer = 1
    '
    Call COMUNI("Proceso Finalizado")
    '
99 Command15.Enabled = True
   
End Sub

Sub Command19_Click()
   Command19.Enabled = False
   Call FINAL("*CODORCL")
   Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Sub Command21_Click()
   '\\\OT-05-0831-WAR-17/11/05///
   Command21.Enabled = False
   '
   IMPOREM04.Show 1
   '
   '\\\OT-05-0831-WAR-FIN///
99 Command21.Enabled = True
End Sub

Sub Command25_Click()
    Command25.Enabled = False
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    Call CARSELMESA
    '
10  Screen.MousePointer = 1
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
       GoTo 99
    End If
    CORRESQ$ = MESANO$(MESASEL$, DES%, HAS%)
    '
    Call FRESHECHO("!SELREMI")
    Screen.MousePointer = 11
    '
    DES1# = CDbl(CVDAT(DES%))
    HAS1# = CDbl(CVDAT(HAS%))
    '
    'ESTO SE AGREGO PARA PODER ANULAR MOVIMIENTOS REGISTRADOS CON FECHA MENORES A 1980
    'OT-10-0186 DE ASEMBLI
    DEED = CVDAT(FECHANUM("01/01/1980"))        'VALOR DE FECHA 01/01/1980
    ANTE# = CDbl(DEED)                          'CORRESPONDE A 01/01/1980
    '
    SQLX$ = " SELECT DISTINCT DoPriCor, Nume_Clie, FechMov, ReferCor FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE (INT(CDBL(FechMov)) >= " & DES1# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS1# & " "
    SQLX$ = SQLX$ + " OR INT(CDBL(FechMov)) <= " & ANTE# & ") "     'AGREGA REMITOS CON FECHA MENORES A O IGUALES A 01/01/1980
    SQLX$ = SQLX$ + " AND (CodiMovi =  '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    'SQLX$ = SQLX$ + "ORDER BY val(MID$(DoPriCor,4)) asc "
    Call APERBASDAT("STOCKS")
    Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
        On Error GoTo 0
        GoTo 10
      End If
      On Error GoTo 0
      '
      J& = 0
      Do While (.EOF = False)
        If Left$(UCase$(!ReferCor), 4) <> "ANUL" Then   ' para evitar comprobantes arruinados
          CLI% = !Nume_Clie
          REMI$ = !DoPriCor
          If TRIM$(REMI$) <> "" Then
            FF% = CVINT(!FECHMOV)
            '
            z$ = Space$(40) + String$(8, 0)
            Call REPLA(z$, REMI$, 1, 10)
              RASOX$ = TRIM$(RASOCLI$(CLI%))
              If RASOX$ = "" Then RASOX$ = "CLIENTE Nro." + TRIM$(Str$(CLI%))
            Call REPLA(z$, RASOX$, 11, 30)
            Call REPLA(z$, MKS$(U&), 41, 4)
            Call REPLA(z$, MKI$(FF%), 45, 2)
            Call REPLA(z$, MKI$(CLI%), 47, 2)
            J& = J& + 1
            Call GRAREG("!SELREMI", z$, J&)
          End If
        End If
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("!SELREMI", J&)
    Call CIERRARCH("!SELREMI")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
       Call MENSERR(24, "No se Encontraron Remitos Posibles de Anular\en el Período Elegido.")
       GoTo 10
    End If
    '
    RERE$ = REGSEL$("!SELREMI")
    If Len(RERE$) <= 4 Then GoTo 10
    '
    REMISEL$ = TRIM$(Left$(RERE$, 10))
    CLIE$ = Mid$(RERE$, 11, 30)
    FEXI% = CVI(Mid$(RERE$, 45, 2))
    FEX = FEXI%
    FEMOVD = Int(CDbl(CVDAT(FEXI%)))
    CLISEL% = CVI(Mid$(RERE$, 47, 2))
    '
    TXT$ = "Remito Número " + TRIM$(REMISEL$) + "\Cliente: " + TRIM$(CLIE$) + "\Fecha: " + FECHATEX$(FEX) + "\   \\Abandonar\Confirma Anular"
    ALTE% = COMALTER%(TXT$)
    If ALTE% <> 2 Then GoTo 10
    '
    '\\\OT-5-0400-FG-14/06/05///
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE DoPriCor = '" & REMISEL$ & "' "
    SQLX$ = SQLX$ + "AND Nume_Clie = " & CLISEL% & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechMov)) = " & FEMOVD
    Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    With MoviTemp
      On Error Resume Next
      .MoveFirst
      If Err > 1 Then
        On Error GoTo 0
        Call MENSERR(24, "No se Encontró Detalle del Remito")
        GoTo 10
      End If
      '
      On Error GoTo 0
      ' VERIFICA SI EL REMITO ESTA FACTURADO
      Do While Not .EOF
        TIDOCSEC$ = ""
        On Error Resume Next
        TIDOCSEC$ = UCase$(Mid$(TRIM$(!DoSecCor), 1, 2))
        On Error GoTo 0
        If TIDOCSEC$ = "FC" Then
          NUMERFAC$ = TRIM$(!DoSecCor)
          Call APERBASDAT("CABAN")
          Screen.MousePointer = 11
          CajaBanc.FindFirst "TipoMovim = 'FVEN' AND CodiCom_Cor = '" & NUMERFAC$ & "'"
          If CajaBanc.NoMatch = False Then
            If (CajaBanc!IDEBECOMP + CajaBanc!IHABECOMP) < 0.005 Then GoTo 16
          End If
          Call COMUNI("Imposible Anular. Remito Ya Facturado\Mediante Comprobante " & NUMERFAC$ & "." _
                      & "\   \Debe Anular la Factura y Re-Intentar.")
          GoTo 10
        End If
      .MoveNext
      Loop
    End With
    '
16  'Call APERBASDAT"("STOCKS")
    Screen.MousePointer = 11
    SQLX$ = "DoPriCor = '" & REMISEL$ & "' "
    SQLX$ = SQLX$ + "AND Nume_Clie = " & CLISEL% & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechMov)) = " & FEMOVD
    '
20  MoviTemp.FindFirst SQLX$
21  If MoviTemp.NoMatch = False Then
      ' VERIFICA SI EL REMITO TIENE ASOCIADO UN PEDIDO
      NROPEDID& = MoviTemp!NUPEDCLI
      If NROPEDID& > 0 Then
        ' ACTUALIZA LOS PEDIDOS QUE FUERON ENTREGADOS CON EL REMITO
          CODINTE% = MoviTemp!COD_INTE
          CANTSALD# = MoviTemp!CANTSALID
          Call APERBASDAT("PEDCLI")
          SQLY$ = "NroPed = " & NROPEDID& & " "
          SQLY$ = SQLY$ + "AND CodiInt = " & CODINTE% & " "
          SQLY$ = SQLY$ + "AND NroClie = " & CLISEL% & " "
          SQLY$ = SQLY$ + "AND CantEnt >= " & FORMATNUM$((CANTSALD# - 0.05), "F15.2") & " "
          SQLY$ = SQLY$ & " AND Status < 9 "
          '
          PedDeta.FindFirst SQLY$
          If PedDeta.NoMatch = False Then
            PedDeta.Edit
            ENTREGA# = PedDeta!CantEnt - CANTSALD#
            SALDO# = PedDeta!CanSaldo + CANTSALD#
            PedDeta!CantEnt = ENTREGA#
            PedDeta!CanSaldo = SALDO#
            PedDeta!NroRemi = ""
            PedDeta.Update
          End If
      End If
      '
      ' ANULACION DEL REMITO
      MoviTemp.FindNext SQLX$
      GoTo 21
    End If
    '\\\OT-5-0400-FG-FIN///
'
' ESTE ES EL CODIGO DE ANULACION PROPIAMENTE DICHO
    MoviTemp.Close
    SQLX$ = " DELETE FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE DoPriCor = '" & REMISEL$ & "' "
    SQLX$ = SQLX$ + "AND Nume_Clie = " & CLISEL% & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechMov)) = " & FEMOVD
    '
    Stocks.Execute (SQLX$)
    '
    Screen.MousePointer = 1
    Call COMUNI("Remito " + REMISEL$ + " Anulado.")
    GoTo 10
    '
99  Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    Command25.Enabled = True
    '
End Sub

Sub Command3_Click()
    '
    Command3.Enabled = False
    '
    ' detalle de entregas por cliente
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    Dim PUNMO$(32767)
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    Call BLOQARCH("DETCLIEN")
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    Call HEADERS("DETCLIEN", "")
    Call HEADERS("DETCLIEN", "Periodo: " + PERIO$)
    '
    'nuevo
    DES1# = CDbl(CVDAT(DES%))
    HAS1# = CDbl(CVDAT(HAS%))
    '
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE INT(CDBL(FechMov)) >= " & DES1# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS1# & " "
     'PARA EVITAR COMPROBANTES ARRUINADOS
    SQLX$ = SQLX$ + " AND LEFT$(UCASE$(REFERCOR),4) <> 'ANUL' "
    SQLX$ = SQLX$ + " AND (CodiMovi =  '" & CMO$ & "'  "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV')  "
    Call APERBASDAT("STOCKS")
    Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
        On Error GoTo 0
        GoTo 10
      End If
      On Error GoTo 0
      '
      J& = 0
      REBLB$ = REGBLAN$("CONESREM")
      Do While (.EOF = False)
        CLI% = !Nume_Clie
        REMIFAC$ = !DoPriCor
        OBSERX$ = ""
        NUMEFAC$ = TRIM$(!DoSecCor)
          If Left$(NUMEFAC$, 2) <> "FC" Then
            NUMEFAC$ = ""
            If Left$(UCase$(!DoSecLar), 9) = "SIN CARGO" Then
              OBSERX$ = !DoSecLar
            End If
          End If
        NPEDX& = !NUPEDCLI
        CANTU# = !CANTSALID - !CANTINGRE
        CII% = !COD_INTE
        NOCOMX$ = NUOCOCLI$(NPEDX&, CII%)
        FF% = CVINT(!FECHMOV)
        If Left$(REMIFAC$, 4) = "DV.0" Then
            REMIFAC$ = !DoSecCor
        End If
        '
        z$ = REGBLAN$("DETCLIEN")
        Call REPLA(z$, MKI$(CLI%), 1, 2)
        Call REPLA(z$, MKI$(FF%), 3, 2)
        Call REPLA(z$, REMIFAC$, 5, 10)
        Call REPLA(z$, MKI$(CII%), 15, 2)
        Call REPLA(z$, MKD$(CANTU#), 17, 8)
        Call REPLA(z$, MKS$(!ValoUnit), 25, 4)
          TOTI# = !ValoUnit * (!CANTSALID - !CANTINGRE)
        Call REPLA(z$, MKD$(TOTI#), 33, 8)
        J& = J& + 1
        Call GRAREG("DETCLIEN", z$, J&)
        '
        z$ = REBLB$
        Call REPLA(z$, MKI$(CLI%), 1, 2)
        Call REPLA(z$, MKI$(FF%), 3, 2)
        Call REPLA(z$, REMIFAC$, 5, 12)
        Call REPLA(z$, NUMEFAC$, 17, 12)
        Call REPLA(z$, MKS$(NPEDX&), 29, 4)
        Call REPLA(z$, NOCOMX$, 35, 16)
        Call REPLA(z$, MKI$(CII%), 51, 2)
        Call REPLA(z$, MKD$(CANTU#), 53, 8)
        Call REPLA(z$, MKS$(!ValoUnit), 61, 4)
        Call REPLA(z$, MKD$(TOTI#), 69, 8)
        Call REPLA(z$, OBSERX$, 85, 25)
        Call GRAREG("CONESREM", z$, J&)
        '
        .MoveNext
      Loop
    End With
    '
    Call SETULTREG("DETCLIEN", J&)
    Call FILESORT("DETCLIEN", "", 1, 5, "ASC")
    Call CIERRARCH("DETCLIEN")
    Call SETULTREG("CONESREM", J&)
    Call FILESORT("CONESREM", "", 1, 5, "ASC")
    Call CIERRARCH("CONESREM")
    '
    OPX% = COMALTER%("Formato de Salida:\\Detalle por Cliente\Control de Estado")
    If OPX% = 1 Then
        Call FINAL("*DETACLI")
      ElseIf OPX% = 2 Then
        Call FINAL("*CONESREM")
    End If
    '
99  Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    Command3.Enabled = True
    '
End Sub
Sub Command4_Click()
    '
    Command4.Enabled = False
    '
    ' detalle de entregas por artículo
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    Call BLOQARCH("DETARTIC")
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    Call HEADERS("DETARTIC", "")
    Call HEADERS("DETARTIC", "Periodo: " + PERIO$)
    '
    'nuevo
    DES1# = CDbl(CVDAT(DES%))
    HAS1# = CDbl(CVDAT(HAS%))
    '
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE INT(CDBL(FechMov)) >= " & DES1# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS1# & " "
     'PARA EVITAR COMPROBANTES ARRUINADOS
    SQLX$ = SQLX$ + " AND LEFT$(UCASE$(REFERCOR),4) <> 'ANUL' "
    SQLX$ = SQLX$ + " AND (CodiMovi =  '" & CMO$ & "'  "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV')  "
    Call APERBASDAT("STOCKS")
    Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
        On Error GoTo 0
        GoTo 10
      End If
      On Error GoTo 0
      '
      J& = 0
      Do While (.EOF = False)
        CLI% = !Nume_Clie
        REMIFAC$ = !DoPriCor
        If !DoPriCor = "" Then
            REMIFAC$ = !DoSecCor
        End If
        
        CANTU# = !CANTSALID - !CANTINGRE
        CII% = !COD_INTE
        FF% = CVINT(!FECHMOV)
        If Left$(REMIFAC$, 4) = "DV.0" Then
            REMIFAC$ = !DoSecCor
        End If
        '
        z$ = REGBLAN$("DETARTIC")
        Call REPLA(z$, MKI$(CII%), 1, 2)
        Call REPLA(z$, MKI$(FF%), 3, 2)
        Call REPLA(z$, REMIFAC$, 5, 10)
        Call REPLA(z$, MKI$(CLI%), 15, 2)
        Call REPLA(z$, MKD$(CANTU#), 17, 8)
        J& = J& + 1
        Call GRAREG("DETARTIC", z$, J&)
        .MoveNext
      Loop
    End With
    '
    Call SETULTREG("DETARTIC", J&)
    Call FILESORT("DETARTIC", "", 1, 5, "ASC")
    Call CIERRARCH("DETARTIC")
    Call FINAL("*DETARTI")
    '
99  Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    Command4.Enabled = True
    '
End Sub

Sub Command5_Click()
    '
    Command5.Enabled = False
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    'nuevo
    DES1# = CDbl(CVDAT(DES%))
    HAS1# = CDbl(CVDAT(HAS%))
    '
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE INT(CDBL(FechMov)) >= " & DES1# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS1# & " "
    'PARA EVITAR COMPROBANTES ARRUINADOS
    SQLX$ = SQLX$ + " AND LEFT$(UCASE$(REFERCOR),4) <> 'ANUL' "
    SQLX$ = SQLX$ + " AND (CodiMovi =  '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV')  "
    SQLX$ = SQLX$ + "ORDER BY FechMov, DoPriCor ASC"
    Call APERBASDAT("STOCKS")
    Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
        On Error GoTo 0
        GoTo 10
      End If
      On Error GoTo 0
      '
      Screen.MousePointer = 11
      J& = 0
      Do While (.EOF = False)
        J& = J& + 1
        '
        DOPRI$ = !DoPriCor
        FRI% = CVINT(!FECHMOV)
        NC% = !Nume_Clie
        CIXI% = !COD_INTE
        CAN# = !CANTSALID - !CANTINGRE
        '
        z$ = REGBLAN$("LIREMIT")
        Call REPLA(z$, DOPRI$, 1, 10)
        Call REPLA(z$, MKI$(FRI%), 11, 2)
         RCCX$ = TRIM$(TEL2CLI$(NC%))
          If RCCX$ = "" Then RCCX$ = RASOCLI$(NC%)
        Call REPLA(z$, RCCX$, 13, 10)
        Call REPLA(z$, MKI$(CIXI%), 23, 2)
        Call REPLA(z$, MKD$(CAN#), 25, 8)
        Call GRAREG("LIREMIT", z$, J&)
       .MoveNext
      Loop
    End With
    '
    Call SETULTREG("LIREMIT", J&)
    Call CIERRARCH("LIREMIT")
    '
    Call HEADERS("LIREMIT", "")
    Call HEADERS("LIREMIT", "Periodo: " + PERIO$)
    '
    Screen.MousePointer = 1
    Call FILESORT("LIREMIT", "LIREMIT", 1, 4, "ASC")
    Call FINAL("*LIREMIT")
    '
99  Command5.Enabled = True
    '

'    Screen.MousePointer = 11
'    J& = 0
'    FIN& = ULTREG&("MATVEN")
'    '
'    For U& = 1 To FIN&
'        Call TRACE(20, U&, FIN&)
'        X$ = REGLEIDO$("MATVEN", U&)
'        FRI% = CVI(Mid$(X$, 7, 2))
'        If FRI% >= DES% Then
'            If FRI% <= HAS% Then
'                NC% = CVI(Left$(X$, 2))
'                DOPRI$ = "RT." + TRIM$(CVS(Mid$(X$, 23, 4)))
'                CIXI% = CVI(Mid$(X$, 9, 2))
'                CAN# = CVS(Mid$(X$, 19, 4))
'                '
'                J& = J& + 1
'                Z$ = REGBLAN$("LIREMIT")
'                Call REPLA(Z$, DOPRI$, 1, 10)
'                Call REPLA(Z$, MKI$(FRI%), 11, 2)
'                  RCCX$ = TRIM$(TEL2CLI$(NC%)) ' fantcli$(nc%)
'                  If RCCX$ = "" Then RCCX$ = RASOCLI$(NC%)
'                Call REPLA(Z$, RCCX$, 13, 10)
'                Call REPLA(Z$, MKI$(CIXI%), 23, 2)
'                Call REPLA(Z$, MKD$(CAN#), 25, 8)
'                Call GRAREG("LIREMIT", Z$, J&)
'            End If
'        End If
'    Next U&
'    '
'    FIN& = ULTREG&("MOVISTO")
'    For U& = 1 To FIN&
'        Call TRACE(20, U&, FIN&)
'        X$ = REGLEIDO$("MOVISTO", U&)
'        FRI% = CVI(Left$(X$, 2))
'        If FRI% >= DES% Then
'          If FRI% <= HAS% Then
'             If UCase$(Mid$(X$, 21, 2)) = CMO$ Then
'                NC% = CVI(Mid$(X$, 84, 2))
'                DOPRI$ = Mid$(X$, 23, 10)
'                CIXI% = CVI(Mid$(X$, 3, 2))
'                CAN# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
'                '
'                J& = J& + 1
'                Z$ = REGBLAN$("LIREMIT")
'                Call REPLA(Z$, DOPRI$, 1, 10)
'                Call REPLA(Z$, MKI$(FRI%), 11, 2)
'                  RCCX$ = TRIM$(TEL2CLI$(NC%))
'                  If RCCX$ = "" Then RCCX$ = RASOCLI$(NC%)
'                Call REPLA(Z$, RCCX$, 13, 10)
'                Call REPLA(Z$, MKI$(CIXI%), 23, 2)
'                Call REPLA(Z$, MKD$(CAN#), 25, 8)
'                Call GRAREG("LIREMIT", Z$, J&)
'             End If
'          End If
'        End If
'    Next U&
'    '
'    Call COPYSTRU("MOVISTO", "MOVIREPA")
'    FIN& = ULTREG&("MOVIREPA")
'    For U& = 1 To FIN&
'        Call TRACE(20, U&, FIN&)
'        X$ = REGLEIDO$("MOVIREPA", U&)
'        FRI% = CVI(Left$(X$, 2))
'        If FRI% >= DES% Then
'          If FRI% <= HAS% Then
'            NC% = CVI(Mid$(X$, 84, 2))
'            DOPRI$ = Mid$(X$, 23, 10)
'            CIXI% = CVI(Mid$(X$, 3, 2))
'            CAN# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
'            '
'            J& = J& + 1
'            Z$ = REGBLAN$("LIREMIT")
'            Call REPLA(Z$, DOPRI$, 1, 10)
'            Call REPLA(Z$, MKI$(FRI%), 11, 2)
'              RCCX$ = TRIM$(TEL2CLI$(NC%))
'              If RCCX$ = "" Then RCCX$ = RASOCLI$(NC%)
'            Call REPLA(Z$, RCCX$, 13, 10)
'            Call REPLA(Z$, MKI$(CIXI%), 23, 2)
'            Call REPLA(Z$, MKD$(CAN#), 25, 8)
'            Call GRAREG("LIREMIT", Z$, J&)
'          End If
'        End If
'    Next U&
'    '
'    Call SETULTREG("LIREMIT", J&)
'    Call CIERRARCH("LIREMIT")
'    '
'    Call HEADERS("LIREMIT", "")
'    Call HEADERS("LIREMIT", "Periodo: " + PERIO$)
'    '
'    Screen.MousePointer = 1
'    Call FINAL("*SOREMIT")
'    '
'99  Command5.Enabled = True
'    '
'End Sub
End Sub

Sub Command6_Click()
    '
    Command6.Enabled = False
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    Call CARSELMESA
    '
10  Screen.MousePointer = 1
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
       GoTo 99
    End If
    CORRESQ$ = MESANO$(MESASEL$, DES%, HAS%)
    '
    Call FRESHECHO("!SELREMI")
    Screen.MousePointer = 11
    '
    DES1# = CDbl(CVDAT(DES%))
    HAS1# = CDbl(CVDAT(HAS%))
    '
    SQLX$ = " SELECT DISTINCT DoPriCor, Nume_Clie, FechMov, DoSecCor FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE INT(CDBL(FechMov)) >= " & DES1# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS1# & " "
    SQLX$ = SQLX$ + " AND (CodiMovi =  '" & CMO$ & "'  "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV')  "
    'PARA EVITAR COMPROBANTES ARRUINADOS
    SQLX$ = SQLX$ + " AND LEFT$(UCASE$(REFERCOR),4) <> 'ANUL' "
    SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
    '
    Call APERBASDAT("STOCKS")
    Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
        On Error GoTo 0
        GoTo 10
      End If
      On Error GoTo 0
      '
      J& = 0
      Do While (.EOF = False)
        CLI% = !Nume_Clie
        REMI$ = !DoPriCor
        FF% = CVINT(!FECHMOV)
        '
        z$ = Space$(40) + String$(8, 0)
        Call REPLA(z$, REMI$, 1, 10)
        Call REPLA(z$, RASOCLI$(CLI%), 11, 30)
        Call REPLA(z$, MKS$(U&), 41, 4)
        Call REPLA(z$, MKI$(FF%), 45, 2)
        Call REPLA(z$, MKI$(CLI%), 47, 2)
        J& = J& + 1
        Call GRAREG("!SELREMI", z$, J&)
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("!SELREMI", J&)
    Call CIERRARCH("!SELREMI")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
       Call MENSERR(24, "No se Encontraron Remitos Posibles de Anular\en el Período Elegido.")
       GoTo 10
    End If
    '
    RERE$ = REGSEL$("!SELREMI")
    If Len(RERE$) <= 4 Then GoTo 10
    '
    REMISEL$ = TRIM$(Left$(RERE$, 10))
    CLIE$ = Mid$(RERE$, 11, 30)
    FEXI% = CVI(Mid$(RERE$, 45, 2))
    FEX = FEXI%
    FEMOVD = Int(CDbl(CVDAT(FEXI%)))
    CLISEL% = CVI(Mid$(RERE$, 47, 2))
    '
    TXT$ = "Remito Número " + TRIM$(REMISEL$) + "\Cliente: " + TRIM$(CLIE$) + "\Fecha: " + FECHATEX$(FEX) + "\   \\Abandonar\Marcar 'Sin Cargo'"
    ALTE% = COMALTER%(TXT$)
    If ALTE% <> 2 Then GoTo 10
    '
    Call APERBASDAT("STOCKS")
    'Se modifico el operador '=' por el like para que busque
    'el numero del Documento Seleccionado dentro del campo Dopricor
    'por que en el caso de la factura el numero aparecia truncado
    'Ejemplo el Documento seleccionado  aparecia de esta forma FC-123456-
    ' y en la base estaba guardado como FC-123456-12
    'y antes no lo encontraba por que lo buscaba por igual
    SQLX$ = "DoPriCor like '" & REMISEL$ & "*' "
    SQLX$ = SQLX$ + "AND Nume_Clie = " & CLISEL% & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechMov)) = " & FEMOVD
    With Movisto
      .FindFirst SQLX$
25    If .NoMatch = False Then
        .Edit
        !DoSecCor = "S/Cargo"
        !DoSecLar = "Sin Cargo"
        .Update
        .FindNext SQLX$
        GoTo 25
      End If
    End With
    '
    Call COMUNI("Movimiento " + REMISEL$ + " Marcado 'SinCargo'.")
    GoTo 10
    '
99  Command6.Enabled = True
    '
End Sub
'

Private Sub Command7_Click()
    Call HELPONLINE("REMITO")
End Sub

Sub Command8_Click()
    Dim RESP%
    Dim NUMEOCXX$
    '
    Command8.Enabled = False
    modo% = COMALTER%("Opciones de Salida\\Pantalla\Impresora\Cancelar")
    If modo% < 1 Or modo% > 2 Then GoTo 99
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
      CMO$ = "RT"
    End If
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    DES1# = CDbl(CVDAT(DES%))
    HAS1# = CDbl(CVDAT(HAS%))
    '
    If modo% = 1 Then GoTo 56
    '////////******NUEVO FILTRO CLIENTE*******///////
    NCMIN% = 32767: NCMAX% = 0
    FIN& = ULTREG&("DEUDOR1")
    For U& = 1 To FIN&
        NCX% = CVI(Left$(REGLEIDO$("DEUDOR1", U&), 2))
        If NCX% > 0 Then
            If NCX% > NCMAX% Then NCMAX% = NCX%
            If NCX% < NCMIN% Then NCMIN% = NCX%
        End If
    Next U&
    '
    With RACLIRPF
      .Text10 = TRIM$(Str$(NCMIN%))
      .Text9 = RASOCLI$(NCMIN%)
      .Text8 = TRIM$(Str$(NCMAX%))
      .Text7 = RASOCLI$(NCMAX%)
      .Show 1
    End With
    '
    'SI CANCELA EN EL FORMULARIO DE RANGO DE CTAS. DE CLIENTES
    If TRIM$(RACLIRPF.Text10.TEXT) = "" And TRIM$(RACLIRPF.Text8.TEXT) = "" Then
        Command8.Enabled = True
        Exit Sub
    End If
    '
    CLINFE = Val(RACLIRPF.Text10)
    CLISUP = Val(RACLIRPF.Text8)
    '//////********HASTA ACA NUEVO FILTRO CLIENTE****///////
56  Screen.MousePointer = 11
    Call APERBASDAT("STOCKS")
    '
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & DES1# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS1# & " "
    SQLX$ = SQLX$ + " AND LEFT$(UCASE$(REFERCOR),4) <> 'ANUL' "   'PARA EVITAR COMPROBANTES ARRUINADOS
    SQLX$ = SQLX$ + " AND (CantSalid >= 0.005 OR CantIngre >= 0.005)"
    If modo% <> 1 Then                                            'FILTRO CLIENTE
       SQLX$ = SQLX$ + " AND Nume_Clie >= " & CLINFE & ""
       SQLX$ = SQLX$ + "AND Nume_Clie<= " & CLISUP & ""
    End If
    '
    Set CONSITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    With CONSITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Remitos Pendientes de Facturación.")  'Generados para este Cliente")
        On Error GoTo 0
        Screen.MousePointer = 1
        Command8.Enabled = True
        Exit Sub
      End If
      On Error GoTo 0
      '
      RECONSI$ = REGBLAN("CONSICLI")
      RECONSJ$ = REGBLAN("REPENFAC")
      J& = 0
      Do While .EOF = False
         NCLIE% = !Nume_Clie
         Call REPLA(RECONSI$, MKI$(NCLIE%), 1, 2)
         Call REPLA(RECONSI$, MKI$(CVINT(!FECHMOV)), 3, 2)
         Call REPLA(RECONSI$, "", 5, 18)
         Call REPLA(RECONSI$, !DoPriCor, 23, 12)
         Call REPLA(RECONSI$, MKI$(!COD_INTE), 35, 2)
         Call REPLA(RECONSI$, MKD$(!CANTSALID), 37, 8)
         Call REPLA(RECONSI$, MKD$(!CANTINGRE), 45, 8)
         Call REPLA(RECONSI$, MKD$(!CantSaldo), 53, 8)
         Call REPLA(RECONSI$, !ReferCor, 61, 48)
         Call REPLA(RECONSI$, String$(12, 0), 109, 12)
         Call REPLA(RECONSI$, MKS$(Val(Mid(!DoPriCor, 5))), 109, 4)
         Call REPLA(RECONSI$, MKI$(CVINT(!FeReMovi)), 122, 2)
          HOR = Format(!FeReMovi, "HH:MM:SS")
         Call REPLA(RECONSI$, MKI$(CVINT(HOR)), 124, 2)
         Call REPLA(RECONSI$, MKI$(!NumUsuar), 127, 2)
         '
         Call REPLA(RECONSJ$, MKI$(NCLIE%), 1, 2)
         Call REPLA(RECONSJ$, MKI$(!COD_INTE), 3, 2)
         Call REPLA(RECONSJ$, !DoPriCor, 5, 12)
         Call REPLA(RECONSJ$, !DoPriLar, 17, 18)
         Call REPLA(RECONSJ$, MKI$(CVINT(!FECHMOV)), 35, 2)
         If IsNull(!NUPEDCLI) Then
            Call REPLA(RECONSJ$, MKS$(0), 37, 4)
         Else
            Call REPLA(RECONSJ$, MKS$(!NUPEDCLI), 37, 4)
         End If
         If IsNull(!NUPEDCLI) Then
             NUMEOCXX$ = ""
           Else
             NUMEOCXX$ = NUOCOCLI$(!NUPEDCLI, !COD_INTE)
         End If
         Call REPLA(RECONSJ$, NUMEOCXX$, 41, 24)
         Call REPLA(RECONSJ$, MKD$(!CANTSALID), 65, 8)
         Call REPLA(RECONSJ$, MKD$(!CANTINGRE), 73, 8)
         Call REPLA(RECONSJ$, MKD$(!CantSaldo), 81, 8)
         Call REPLA(RECONSJ$, MKD$(!ValoUnit), 89, 8)
           TOTI# = !ValoUnit * (!CANTSALID - !CANTINGRE)
         Call REPLA(RECONSJ$, MKD$(TOTI#), 97, 8)
         '
         J& = J& + 1
         Call GRAREG("CONSICLI", RECONSI$, J&)
         Call GRAREG("REPENFAC", RECONSJ$, J&)
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("CONSICLI", J&)
    Call FILESORT("CONSICLI", "", 1, 6, "ASC")
    Call CIERRARCH("CONSICLI")
    '
    Call SETULTREG("REPENFAC", J&)
    Call FILESORT("REPENFAC", "", 1, 6, "ASC")
    Call CIERRARCH("REPENFAC")
    '
    If modo% = 1 Then
        Call CONECRUN("*CREPENFA", "Remitos Sin Facturar (Clientes)")
      Else
        RESP% = COMALTER%("Opciones de Agrupamiento\\Comprobante\Articulo\Cancelar")
        If RESP% = 1 Then
            Call CONECRUN("*LIREPEFA", "Remitos Sin Facturar (Clientes)")
          ElseIf RESP% = 2 Then
            Call HEADERS("REPENFAC", "")
            If CLINFE = CLISUP Then
                NCXX% = CLINFE
                Call HEADERS("REPENFAC", "Cliente: " + RASOCLI$(NCXX%))
                Call CONECRUN("*LREPENFB", "Remitos Sin Facturar (Clientes)")
              Else
                Call CONECRUN("*LREPENFA", "Remitos Sin Facturar (Clientes)")
            End If
        End If
    End If
    '
99  Command8.Enabled = True
End Sub

Sub Command9_Click()
    Command9.Enabled = False
    '
    '\\\OT-05-0437-WAR-05/07/05///
        modo% = COMALTER%("Modalidad de Devolución\\Por Remito Propio\Por Remito de Cliente\Cancelar")
        If modo% < 1 Or modo% > 2 Then GoTo 99
        If modo% = 2 Then 'Devolución de Remito de Cliente.
          Call DEVRECLI
          GoTo 99
        End If
        '\\\OT-05-0437-WAR-FIN///
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If Len(CMO$) <> 2 Then
       CMO$ = "RT"
    End If
    '
    Call CARSELMESA
    '
10  Screen.MousePointer = 1
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
       GoTo 99
    End If
    CORRESQ$ = MESANO$(MESASEL$, DES%, HAS%)
    '
    Call FRESHECHO("SELREMI")
    Screen.MousePointer = 11
    '
    DES1# = CDbl(CVDAT(DES%))
    HAS1# = CDbl(CVDAT(HAS%))
    '
    SQLX$ = " SELECT DISTINCT DoPriCor, Nume_Clie, FechMov, DoSecCor FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE INT(CDBL(FechMov)) >= " & DES1# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS1# & " "
    'PARA EVITAR COMPROBANTES ARRUINADOS
    SQLX$ = SQLX$ + " AND LEFT$(UCASE$(REFERCOR),4) <> 'ANUL' "
    'SQLX$ = SQLX$ + " AND Len(DoSecCor) < 1 "  ' PARA QUE PUEDA DEVOLVERSE AUN FACTURADO
    SQLX$ = SQLX$ + " AND CodiMovi =  '" & CMO$ & "'  "
    '
    Call APERBASDAT("STOCKS")
    Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
        On Error GoTo 0
        GoTo 10
      End If
      On Error GoTo 0
      '
      J& = 0
      Do While (.EOF = False)
        CLI% = !Nume_Clie
        REMI$ = !DoPriCor
        FF% = CVINT(!FECHMOV)
        '\\\OT-06-0083-WAR-07/03/06///
        F1$ = FECHATEX$(FF%)
        REFE1$ = F1$ + " - " + RASOCLI$(CLI%)
        '\\\OT-06-0083-WAR-FIN///
        '
        z$ = Space$(40) + String$(8, 0)
        Call REPLA(z$, REMI$, 1, 10)
        Call REPLA(z$, REFE1$, 11, 30) ' RASOCLI$(CLI%), 11, 30)
        Call REPLA(z$, MKS$(U&), 41, 4)
        Call REPLA(z$, MKI$(FF%), 45, 2)
        Call REPLA(z$, MKI$(CLI%), 47, 2)
        J& = J& + 1
        Call GRAREG("!SELREMI", z$, J&)
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("!SELREMI", J&)
    Call CIERRARCH("!SELREMI")
    Call FRESHECHO("!SELREMI")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
       Call MENSERR(24, "No se Encontraron Remitos Posibles de Devolución\en el Período Elegido.")
       GoTo 10
    End If
    '
28  RERE$ = REGSEL$("!SELREMI")
    If Len(RERE$) <= 4 Then GoTo 10
    '
    REMISEL$ = TRIM$(Left$(RERE$, 10))
    CLIE$ = Mid$(RERE$, 11, 30)
    FEXI% = CVI(Mid$(RERE$, 45, 2))
    FEX = FEXI%
    FEMOVD = Int(CDbl(CVDAT(FEXI%)))
    CLISEL% = CVI(Mid$(RERE$, 47, 2))
    '
    'TXT$ = "Remito Número " + TRIM$(REMISEL$) + "\Cliente: " + TRIM$(CLIE$) + "\Fecha: " + FECHATEX$(FEX) + "\   \\Abandonar\Marcar 'Sin Cargo'"
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = "SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE DoPriCor = '" & REMISEL$ & "' "
    SQLX$ = SQLX$ + " AND CodiMovi =  '" & CMO$ & "'  "
    'SQLX$ = SQLX$ + "AND Len(DoSecCor) < 1 "   ' PARA QUE PUEDAN DEVOLVERSE AUN FACTURADOS
    SQLX$ = SQLX$ + "AND Nume_Clie = " & CLISEL% & " "
    'SQLX$ = SQLX$ + "AND INT(CDBL(FechMov)) = " & FEMOVD
    Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
        On Error GoTo 0
        GoTo 10
      End If
      On Error GoTo 0
      '
      DOPRI$ = REMISEL$
      DOPRILA$ = !DoPriLar
      DOSEC$ = ""       ' QUEDA LIBRE PARA PONERLE EL NUMERO DE FACTURA
      DOSELA$ = ""      ' QUEDA LIBRE PARA PONERLE EL NUMERO DE FACTURA
      MONEVA% = !MoneVal
      DEPOMO% = !DEPOMOVI
      '
      HOII% = HOY(HOS$)
      '
      J& = 0
      Do While (.EOF = False)
        XX$ = REGBLAN$("BOLREDET")
        Call REPLA(XX$, MKS$(!NuOrItem), 1, 4)
        Call REPLA(XX$, MKI$(!COD_INTE), 83, 2)
        Call REPLA(XX$, UNIMED$(!COD_INTE), 85, 4)
        Call REPLA(XX$, MKD$(!CANTSALID), 89, 8)
        Call REPLA(XX$, !IdenLote, 119, 16)
          VUNIT = !ValoUnit
        Call REPLA(XX$, MKS$(VUNIT), 159, 4)
        Call REPLA(XX$, MKD$(0), 171, 8)
        J& = J& + 1
        Call GRAREG("!BOLREDET", XX$, J&)
      .MoveNext
      Loop
    End With
    Call SETULTREG("!BOLREDET", J&)
    '
    '\\\OT-06-0084-WAR-07/03/06///
29  VDEF$ = Space(4)
    Call REPLA(VDEF$, MKI$(FEXI%), 1, 2)
    Call REPLA(VDEF$, MKI$(DEPOMO%), 3, 1)
    OBJ1$ = OBJPLA$("ENTRAFECHADEVU", VDEF$)
    If TRIM$(OBJ1$) = "" Then GoTo 28
    FEDEVU% = CVI(Mid$(OBJ1$, 1, 2))
    DEPOMO1% = Asc(Mid$(OBJ1$, 3, 1))
    If ECOARCH$("TADEPOSI", MKI$(DEPOMO1%)) = "" Then
      Call MENSERR(24, "Depósito no Válido.")
      GoTo 29
    End If
    '\\\OT-06-0084-WAR-FIN///
    'FEDEVU% = CVI(OBJPLA$("ENTRAFECHADEVU", MKI$(FEXI%)))
    If FEDEVU% < 1 Or FEDEVU% > HOY(HOS$) Then GoTo 10
    FEX = FEDEVU%
    '
30  VTB% = VENTABU%("DEVOLCLI/TITUPAN=Devolucion " + REMISEL$ + " - " + TRIM$(CLIE$) + " - " + REPLACAR$(FECHATEX$(FEX), "/", "-"))
    If VTB% < 0 Then GoTo 10
    '
    For KKIL& = 1 To J&
      XX$ = REGLEIDO$("!BOLREDET", KKIL&)
      If CVD(Mid$(XX$, 171, 8)) > CVD(Mid$(XX$, 89, 8)) + 0.05 Then
        Call MENSERR(24, "La Cantidad Devuelta no Puede Superar\la Despachada Originalmente.")
        GoTo 30
      End If
    Next KKIL&
    '
    For KKIL& = 1 To J&
      XX$ = REGLEIDO$("!BOLREDET", KKIL&)
      If CVD(Mid$(XX$, 171, 8)) > 0 Then
        CIXI% = CVI(Mid$(XX$, 83, 2))
        LOTE$ = Mid$(XX$, 119, 16)
        NUORIT% = CVS(Left$(XX$, 4)) - 1
        CADEVU = CVD(Mid$(XX$, 171, 8))
        VUNI# = CVS(Mid$(XX$, 159, 4))
        '
        Call MOVISTOK(FEDEVU%, CIXI%, LOTE$, "DV", DOPRILA$, DOPRI$, DOSELA$, DOSEC$, "Devol." + RASOCLI$(CLISEL%), NUORIT%, VUNI#, MONEVA%, CLISEL%, DEPOMO1%, CADEVU)
        '\\\OT-06-0105-WAR-27/03/06///
        If DESCOMPO(CIXI%) > 0 Then
          'SI HAY DEVULUCION DE KIT
          'COMPENSO CONJUNTO
          Call MOVISTOK(FEDEVU%, CIXI%, LOTE$, "CI", DOPRILA$, DOPRI$, DOSELA$, DOSEC$, "Devol." + RASOCLI$(CLISEL%), NUORIT%, VUNI#, MONEVA%, CLISEL%, DEPOMO1%, CADEVU * (-1))
          'COMPENSO COMPONENTES
          Call AREXPLO(CIXI%, 1, 1)
          For KKJ% = 1 To CAIT%
            CICO% = CIJ%(KKJ%)
            CACICO = USOI(KKJ%) * CADEVU
            Call MOVISTOK(FEDEVU%, CICO%, LOTE$, "CI", DOPRILA$, DOPRI$, DOSELA$, DOSEC$, "Devol." + RASOCLI$(CLISEL%), NUORIT%, VUNI#, MONEVA%, CLISEL%, DEPOMO1%, CACICO)
          Next KKJ%
        End If
        '\\\OT-06-0105-WAR-FIN///
      End If
    Next KKIL&
    Movisto.Close
    '
    Call COMUNI("Devolución Procesada")
    GoTo 10
    '
99  Command9.Enabled = True
    '
End Sub

Sub RECONCLI(NCX7%)
   '
   Dim SALCONSI(32767)
   Call SETULTREG("RECONCLI", 0)
   Call SETULTREG("RECONMAT", 0)
   '
   If NCX7% > 0 Then
     If ECOARCH$("DEUDOR1", MKI$(NCX7%)) <> "" Then
       Screen.MousePointer = 11
       FIN& = ULTREG&("CONSICLI")
       For KLK& = 1 To 32767: SALCONSI(KLK&) = 0: Next KLK&
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$("CONSICLI", U&)
         NCX8% = CVI(Left$(XX$, 2))
         If NCX8% = NCX7% Then
           CIXI% = CVI(Mid$(XX$, 35, 2))
           If CIXI% > 0 Then
             If CIXI% <= NUMAS% Then
               CACON = CVD(Mid$(XX$, 37, 8)) - CVD(Mid$(XX$, 45, 8))
               SALCONSI(CIXI%) = SALCONSI(CIXI%) + CACON
             End If
           End If
         End If
       Next U&
       '
       JJ& = 0
       FIN& = NUINV%
       For U& = 1 To NUINV%
         Call TRACE(20, U&, FIN&)
         II$ = REGLEIDO$("INVERT", U& + 1)
         CIXI% = CVI(Mid$(II$, 17, 2))
         If CIXI% > 0 Then
           If CIXI% <= NUMAS% Then
             If SALCONSI(CIXI%) > 0 Then
               XX$ = REGBLAN$("RECONCLI")
               YY$ = REGBLAN$("RECONMAT")
               Call REPLA(XX$, MKI$(NCX7%), 1, 2)
               Call REPLA(XX$, MKI$(CIXI%), 3, 2)
               Call REPLA(XX$, MKS$(SALCONSI(CIXI%)), 5, 4)
               Call REPLA(XX$, String$(16, "."), 9, 16)
               Call REPLA(YY$, MKI$(CIXI%), 1, 2)
               Call REPLA(YY$, MKS$(SALCONSI(CIXI%)), 3, 4)
               JJ& = JJ& + 1
               Call GRAREG("RECONCLI", XX$, JJ&)
               Call GRAREG("RECONMAT", YY$, JJ&)
             End If
           End If
         End If
       Next U&
       Call SETULTREG("RECONCLI", JJ&)
       Call CIERRARCH("RECONCLI")
       Call SETULTREG("RECONMAT", JJ&)
       Call CIERRARCH("RECONMAT")
       Screen.MousePointer = 1
       '
     End If
   End If

End Sub

Private Sub Form_Activate()
    '
    Command19.Enabled = False
    If CONTROL$("REMSTAN", "REPLAORI") = "SI" Then
      Command19.Enabled = True
    End If
    '
    If EXISTE%("COCUMENT.EXE") > 0 Then
      Command11.Enabled = True
    End If
    '
End Sub
'\\\OT-05-0437-WAR-05/05/07///

Sub DEVRECLI()
    '\\\OT-06-0084-WAR-08/03/06///
10  DP1% = Val(CONTROL$("", "DEPOEXPE"))
    y$ = Space(64)
    Call REPLA(y$, MKI$(0), 1, 2)
    Call REPLA(y$, MKI$(0), 3, 2)
    Call REPLA(y$, MKI$(0), 5, 2)
    Call REPLA(y$, MKI$(DP1%), 47, 1)
    '\\\OT-06-0084-WAR-FIN///
    '
    OBJE$ = "DEVOCLI"
    PEDY$ = OBJPLA$(OBJE$, y$)
    If PEDY$ = "" Then Exit Sub
    NC% = CVI(Mid$(PEDY$, 1, 2))
    CI% = CVI(Mid$(PEDY$, 5, 2))
    FEDEVU% = CVI(TRIM(Mid$(PEDY$, 3, 2)))
    FE2# = CDbl(CVDAT(FEDEVU%))
    NURE$ = TRIM(Mid$(PEDY$, 7, 16))
    '\\\OT-06-0084-WAR-08/03/06///
    DEPOMO% = Asc(Mid$(PEDY$, 47, 1))
    If ECOARCH$("TADEPOSI", MKI$(DEPOMO%)) = "" Then
      Call MENSERR(24, "Depósito no Válido.")
      GoTo 10
    End If
    '\\\OT-06-0084-WAR-FIN///
    'FALTA VALIDAR QUE NO EXISTA EL REMITO DEL CLIENTE EN UNA DEVOLUCION
    ' en !NumRem_cli CON NURE$
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If CMO$ = "" Then CMO$ = "RT"
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = "SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE CodiMovi =  '" & CMO$ & "'  "
    SQLX$ = SQLX$ + " AND LEFT(DOSECCOR,2) <> 'FC' " 'PARA QUE NO TRAIGA LOS FACTURADOS
    SQLX$ = SQLX$ + "AND Nume_Clie = " & NC% & " "
    SQLX$ = SQLX$ + " AND COD_INTE = " & CI% & ""
    SQLX$ = SQLX$ + " AND LEFT$(UCASE$(REFERCOR),4) <> 'ANUL' "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechMov)) <= " & FE2# & " "
    SQLX$ = SQLX$ + " ORDER BY DOPRILAR "
    '
    Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
        On Error GoTo 0
        GoTo 10
      End If
      On Error GoTo 0
      '
      HOII% = HOY(HOS$)
      '
      J& = 0
      Do While (.EOF = False)
         XX$ = REGBLAN$("BOLREDET")
         '
         DOPRILA$ = !DoPriLar
         DOSEC$ = ""       ' QUEDA LIBRE PARA PONERLE EL NUMERO DE FACTURA
         DOSELA$ = ""      ' QUEDA LIBRE PARA PONERLE EL NUMERO DE FACTURA
         MONEVA% = !MoneVal
         'DEPOMO% = !DepoMovi  'EL DEPOSITO LO ASGINA EL USURARIO'OT-06-0084-WAR-08/03/06///
         DOPRICO1$ = !DoPriCor 'HAY QUE GRABARLO EN EL ARCHI.
         '
        Call REPLA(XX$, MKS$(!NuOrItem), 1, 4)
          REMISEL$ = !DoPriLar
          DOPRI$ = REMISEL$
        Call REPLA(XX$, REMISEL$, 5, 45) 'ACA PONGO EL REMITO LARGO
        Call REPLA(XX$, DOPRICO1$, 50, 14) 'ACA GRABO EL REMITO CORTO
          FECHMO = !FECHMOV
        Call REPLA(XX$, MKI(CVINT(FECHMO)), 81, 2)
        Call REPLA(XX$, MKI$(!COD_INTE), 83, 2)
        Call REPLA(XX$, UNIMED$(!COD_INTE), 85, 4)
        Call REPLA(XX$, MKD$(!CANTSALID), 89, 8)
        Call REPLA(XX$, Chr(DEPOMO%), 97, 1)
        Call REPLA(XX$, !IdenLote, 119, 16)
          VUNIT = !ValoUnit
        Call REPLA(XX$, MKS$(VUNIT), 159, 4)
        Call REPLA(XX$, MKD$(0), 171, 8)
        J& = J& + 1
        Call GRAREG("!BOLREDET", XX$, J&)
      .MoveNext
      Loop
    End With
    Call SETULTREG("!BOLREDET", J&)
    '
40  VTB% = VENTABU%("DEVCLIRT/TITUPAN=Devolucion " + DESCRIT0(CI%) + " - " + TRIM$(RASOCLI(NC%)) + " - " + REPLACAR$(FECHATEX$(FEX), "/", "-"))
    If VTB% < 0 Then GoTo 10
    '
    For KKIL& = 1 To J&
      XX$ = REGLEIDO$("!BOLREDET", KKIL&)
      If CVD(Mid$(XX$, 163, 8)) > CVD(Mid$(XX$, 89, 8)) + 0.05 Then
        Call MENSERR(24, "La Cantidad Devuelta no Puede Superar\la Despachada Originalmente.")
        GoTo 40
      End If
    Next KKIL&
    '
    For KKIL& = 1 To J&
      XX$ = REGLEIDO$("!BOLREDET", KKIL&)
      If CVD(Mid$(XX$, 163, 8)) > 0 Then
        DOPRI$ = TRIM(Mid$(XX$, 50, 14))
        CIXI% = CVI(Mid$(XX$, 83, 2))
        LOTE$ = Mid$(XX$, 119, 16)
        NUORIT% = CVS(Left$(XX$, 4)) - 1
        CADEVU = CVD(Mid$(XX$, 163, 8))
        VUNI# = CVS(Mid$(XX$, 159, 4))
        DEPOMO% = Asc(Mid$(XX$, 97, 1))
        Call MOVISTOK(FEDEVU%, CIXI%, LOTE$, "DV", DOPRILA$, DOPRI$, DOSELA$, DOSEC$, "Devol." + RASOCLI$(NC%), NUORIT%, VUNI#, MONEVA%, NC%, DEPOMO%, CADEVU)
        If DESCOMPO(CIXI%) > 0 Then
          'SI HAY DEVULUCION DE KIT
          'COMPENSO EL CONJUNTO
            Call MOVISTOK(FEDEVU%, CIXI%, LOTE$, "CI", DOPRILA$, DOPRI$, DOSELA$, DOSEC$, "Devol." + RASOCLI$(NC%), NUORIT%, VUNI#, MONEVA%, NC%, DEPOMO%, CADEVU * (-1))
          'COMPENSO LOS COMPONENTES
          Call AREXPLO(CIXI%, 1, 1)
          For KKJ% = 1 To CAIT%
            CICO% = CIJ%(KKJ%)
            CACICO = USOI(KKJ%) * CADEVU
            Call MOVISTOK(FEDEVU%, CICO%, LOTE$, "CI", DOPRILA$, DOPRI$, DOSELA$, DOSEC$, "Devol." + RASOCLI$(NC%), NUORIT%, VUNI#, MONEVA%, NC%, DEPOMO%, CACICO)
          Next KKJ%
        End If
      End If
    Next KKIL&
    Movisto.Close
    '
    Call COMUNI("Devolución Procesada")
    '
999 End Sub
