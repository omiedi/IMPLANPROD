VERSION 5.00
Begin VB.Form OPCIPRE04 
   BackColor       =   &H00D0F0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Especificaciones de Venta - Opciones"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4995
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Picture         =   "OPCOPRE04.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Indice de Cumplimiento de Entregas"
      Top             =   6930
      Width           =   650
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00006060&
      Height          =   7155
      Left            =   4095
      ScaleHeight     =   7095
      ScaleWidth      =   4215
      TabIndex        =   18
      Top             =   -360
      Width           =   4275
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
         Left            =   120
         Picture         =   "OPCOPRE04.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Configuracion de Funcionamiento"
         Top             =   4200
         Visible         =   0   'False
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
         Picture         =   "OPCOPRE04.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1400
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
         Picture         =   "OPCOPRE04.frx":1016
         Style           =   1  'Graphical
         TabIndex        =   21
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
         TabIndex        =   19
         Top             =   2760
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "OPCOPRE04.frx":1160
         Top             =   3100
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Resúmenes de EV's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   3690
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   855
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   11
         Top             =   1590
         Visible         =   0   'False
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
            Picture         =   "OPCOPRE04.frx":3082
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Listado de Base de Datos Completa de E.V.'s"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Base de Datos de E.V.'s Completa"
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
            Top             =   0
            Width           =   1755
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   645
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   5
         Top             =   1065
         Visible         =   0   'False
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
            Picture         =   "OPCOPRE04.frx":338C
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Listado de E.V.'s por Cliente"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Resumen de E.V.'s por Cliente"
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
         Left            =   380
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   8
         Top             =   600
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
            Picture         =   "OPCOPRE04.frx":37CE
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Listado de E.V.'s por Fecha de Entrega Prometida"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "EV's por Fecha de Entrega"
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
      Height          =   1155
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   3690
      Begin VB.PictureBox Picture9 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   480
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   31
         Top             =   360
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
            Left            =   0
            Picture         =   "OPCOPRE04.frx":3AD8
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Devolución Material en Consignación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Materiales con Fecha de Necesidad Critica."
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
            TabIndex        =   33
            Top             =   45
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   480
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   28
         Top             =   1200
         Width           =   2640
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
            Picture         =   "OPCOPRE04.frx":3C22
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Devolución Material en Consignación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
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
            TabIndex        =   30
            Top             =   45
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   25
         Top             =   1470
         Width           =   2640
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
            Picture         =   "OPCOPRE04.frx":3D6C
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Devolución Material en Consignación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Corrección de Canti- dades Despachadas"
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
            TabIndex        =   27
            Top             =   105
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   15
         Top             =   1320
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
            Picture         =   "OPCOPRE04.frx":4076
            Style           =   1  'Graphical
            TabIndex        =   17
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
            TabIndex        =   16
            Top             =   45
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   360
         ScaleHeight     =   585
         ScaleWidth      =   2595
         TabIndex        =   1
         Top             =   1200
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
            Picture         =   "OPCOPRE04.frx":41C0
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
            Picture         =   "OPCOPRE04.frx":44CA
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
Attribute VB_Name = "OPCIPRE04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    '
    Command1.Enabled = False
    OPREMIT04.Show 1
    Command1.Enabled = True
    '
End Sub

Private Sub Command10_Click()
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
     Call ABRESTOCKS
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
         FF% = CVINT(!FechMov)
         '
         aa2 = !Nume_Clie
         Z$ = Space$(40) + String$(8, 0)
         Call REPLA(Z$, REMI$, 1, 10)
         Call REPLA(Z$, RASOCLI$(CLI%), 11, 30)
         Call REPLA(Z$, MKS$(U&), 41, 4)
         Call REPLA(Z$, MKI$(FF%), 45, 2)
         Call REPLA(Z$, MKI$(CLI%), 47, 2)
         J& = J& + 1
         Call GRAREG("!SELREMI", Z$, J&)
       .MoveNext
       Loop
     End With
     '
     Call SETULTREG("!SELREMI", J&)
     Call CIERRARCH("!SELREMI")
     Screen.MousePointer = 1
     '
     If J& < 1 Then
        Call MENSERR(24, "No se Encontraron Remitos Posibles de Devolución\en el Período Elegido.")
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
20   Call ABRESTOCKS
     Screen.MousePointer = 11
     SQLX$ = "SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + "WHERE DoPriCor = '" & REMISEL$ & "' "
     SQLX$ = SQLX$ + "AND Nume_Clie = " & CLISEL% & " "
     SQLX$ = SQLX$ + "ORDER BY NuOrItem ASC"
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
       DEPOMO% = !DepoMovi
       '
       HOII% = HOY(HOS$)
       '
       J& = 0: NORI$ = ""
       Do While (.EOF = False)
         NORIT% = !NuOrItem
         For KKI% = 1 To Len(NORI$) Step 2
           If CVI(Mid$(NORI$, KKI%, 2)) = NORIT% Then
             '
             ' RUTINA DE RE-NUMERACION DE ITEMS INCLUYENDO DEVOLUCIONES
             SCMA = Screen.MousePointer
             Screen.MousePointer = 11
             NORIT% = 0
             SQLX$ = " DoPriCor = '" & REMISEL$ & "' "
             SQLX$ = SQLX$ + " AND Nume_Clie = " & CLISEL% & " "
             With Movisto
               .FindFirst SQLX$
25             If .NoMatch = False Then
                 .Edit
                 NORIT% = NORIT% + 1
                 !NuOrItem = NORIT%
                 .Update
                 .FindNext SQLX$
                 GoTo 25
               End If
             End With
             Screen.MousePointer = SCMA
             GoTo 20
             '
           End If
         Next KKI%
         NORI$ = NORI$ + MKI$(NORIT%)
         '
         If Abs(!CantIngre - !CantSalid) >= 0.005 Then
           XX$ = REGBLAN$("BOLREDET")
           Call REPLA(XX$, MKS$(!NuOrItem), 1, 4)
           Call REPLA(XX$, MKI$(!COD_INTE), 83, 2)
           Call REPLA(XX$, UNIMED$(!COD_INTE), 85, 4)
           Call REPLA(XX$, MKD$(!CantSalid - !CantIngre), 89, 8)
           Call REPLA(XX$, !IdenLote, 119, 16)
             VUNIT = !ValoUnit
             NUPEDI& = !NUPEDCLI 'NUMERO DE PEDIDO
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
     Call ABRESTOCKS
     Screen.MousePointer = 11
     For KKIL& = 1 To J&
       XX$ = REGLEIDO$("!BOLREDET", KKIL&)
       CIXI% = CVI(Mid$(XX$, 83, 2))
       LOTE$ = Mid$(XX$, 119, 16)
       NUORIT% = CVS(Left$(XX$, 4))
       CACORREC = CVD(Mid$(XX$, 89, 8))
       NUPEDID1& = CVS(Mid$(XX$, 151, 4))
       VUNIT = CVS(Mid$(XX$, 159, 4))
       '
       SQLX$ = " DoPriCor = '" & REMISEL$ & "' "
       SQLX$ = SQLX$ + " AND Nume_Clie = " & CLISEL% & " "
       SQLX$ = SQLX$ + " AND NuOrItem = " & NUORIT% & ""
       '
       With Movisto
         .FindFirst SQLX$
         If .NoMatch = False Then
           .Edit
            !CantSalid = 0: !CantIngre = 0
            If CACORREC > 0 Then
                !CantSalid = CACORREC 'GUARDA LA CANTIDAD EDITADA
              ElseIf CACORREC < 0 Then
                !CantIngre = Abs(CACORREC)
            End If
            !ValoUnit = VUNIT 'GUARDA EL VALOR UNITARIO EDITADO
            !IdenLote = LOTE$ 'GUARDA EL LOTE EDITADO
            !COD_INTE = CIXI% 'GUARDA EL CODIGO INTERNO EDITADO
            !Cod_Exte = CODEXT(CIXI%) 'GUARDA EL COD. EXT EDITADO
            !NUPEDCLI = NUPEDID1& 'GUARDA EL Nº PED. EDITADO
            !Descrip = DESCRIT(CIXI%) 'GUARDA LA DESCRIPCION.
           .Update
         End If
       End With
       '
     Next KKIL&
     Movisto.Close
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

Private Sub Command19_Click()
   Command19.Enabled = False
   Call FINAL("*CODORCL")
   Command19.Enabled = True
End Sub

Private Sub Command12_Click()
    '
    Call ABREORTRA
    SQLX$ = " SELECT * FROM LISTAMAT "
    SQLX$ = SQLX$ + " WHERE (Nro_Ocom <> '' OR ISNULL(Nro_Ocom) = FALSE OR ISEMPTY(Nro_Ocom) = FALSE) "
    SQLX$ = SQLX$ + " AND (Nro_BolRe = '' or isnull(Nro_BolRe) = true OR IsEmpty(Nro_BolRe) = true)"
    SQLX$ = SQLX$ + " ORDER BY Nro_Otrab, Nro_Grup "
    Set ALARTMP = OTrabajo.OpenRecordset(SQLX$, 4)
    With ALARTMP
        '
        On Error Resume Next
        .MoveFirst
        If Err Then
            Call COMUNI("No Existen Materiales segun condicion SQL")
            On Error GoTo 0
            Exit Sub
        End If
        J& = 0
        XXY$ = REGBLAN("ALARMAT")
        Do While .EOF = False
            'HAGO EL CALCULO DE LA FEC ENTRE SOLI DEL MATERIAL
            ' calcula fecha de necesidad
            NOTRA$ = !Nro_Otrab
            NUGRU$ = !Nro_Grup
            CLAVX% = Val(!CLAVE) Mod 10 'CLAVE
            '
            FENECESI% = 0
            If CLAVX% = 5 Or CLAVX% = 7 Then
             GrupOtra.FindFirst "Nro_Otrab = '" & NOTRA$ & "'AND Nro_Grup = '" & NUGRU$ & "' "
             If GrupOtra.NoMatch = False Then
                  PLAX$ = "0000"
                  On Error Resume Next
                  PLAX$ = GrupOtra!PLA05_GRUP
                  If CLAVX% = 7 Then PLAX$ = GrupOtra!PLA07_GRUP
                  On Error GoTo 0
                  FENECESI% = FECHSEMAN(PLAX$)
              End If
            End If
            'obtiene fecha de entre solicitada de o/c
            If IsNull(!Nro_Ocom) Then
              NROC1$ = ""
            Else
                NROC1$ = !Nro_Ocom
            End If
            '
            NROC2& = Val(Mid$(NROC1$, 4, 9))
            CODO2$ = !CODEXTE
            MATE2$ = !DESCRI_COR
            'SI NRO DE O/C ES VACIO VA AL PROXIMO REGISTRO
            If TRIM$(NROC1$) = "" Then
                GoTo 36
            End If
            '
            FENSO1% = 0
            Call ABRECOMPRAS
            With OcomDeta
                SQLX$ = "Nume_Ocom = " & NROC2& & ""
                SQLX$ = SQLX$ + "AND Cod_Exte = '" & CODO2$ & "'"
                SQLX$ = SQLX$ + " OR DESCRITEM= '" & MATE2$ & "'"
                .FindFirst SQLX$
                If .NoMatch = False Then
                        FENSO1% = CVINT(!Fentre_Sol)
                End If
            End With
            'COMPARO DIFERENCIA DE FECHAS
            If FENECESI% < FENSO1% And FENECESI% <> 0 Then 'SI FECHA NECESIDAD ES MENOR A FECHA ENTREGA SOLI
                J& = J& + 1
                'ACA GRABA EL ARCHIVO PARA LISTAR
                Call REPLA(XXY$, NOTRA$, 1, 16)
                Call REPLA(XXY$, NUGRU$, 17, 10)
                Call REPLA(XXY$, MKS(!CANTID), 27, 4)
                If TRIM(CODO2$) = "" Then
                    Call REPLA(XXY$, MATE2$, 35, 24)
                 Else
                    Call REPLA(XXY$, CODO2$, 35, 24)
                End If
                '
                Call REPLA(XXY$, MKI(FENECESI%), 59, 2)
                Call REPLA(XXY$, NROC1$, 61, 16)
                Call REPLA(XXY$, MKI(FENSO1%), 77, 2)
                '
                Call GRAREG("ALARMAT", XXY$, J&)
            End If
            '
36          .MoveNext
        Loop
        '
        Call SETULTREG("ALARMAT", J&)
    End With
    '
    Call FINAL("*ALARMAT")
    '
End Sub

Private Sub Command2_Click()
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command25_Click()
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
    SQLX$ = " SELECT DISTINCT DoPriCor, Nume_Clie, FechMov FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE INT(CDBL(FechMov)) >= " & DES1# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS1# & " "
    SQLX$ = SQLX$ + " AND (CodiMovi =  '" & CMO$ & "'  "
    'PARA EVITAR COMPROBANTES ARRUINADOS
    SQLX$ = SQLX$ + " AND LEFT$(UCASE$(REFERCOR),4) <> 'ANUL' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    'SQLX$ = SQLX$ + "ORDER BY val(MID$(DoPriCor,4)) asc "
    Call ABRESTOCKS
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
        FF% = CVINT(!FechMov)
        '
        Z$ = Space$(40) + String$(8, 0)
        Call REPLA(Z$, REMI$, 1, 10)
        Call REPLA(Z$, RASOCLI$(CLI%), 11, 30)
        Call REPLA(Z$, MKS$(U&), 41, 4)
        Call REPLA(Z$, MKI$(FF%), 45, 2)
        Call REPLA(Z$, MKI$(CLI%), 47, 2)
        J& = J& + 1
        Call GRAREG("!SELREMI", Z$, J&)
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
    Call ABRESTOCKS
    SQLX$ = "DoPriCor = '" & REMISEL$ & "' "
    SQLX$ = SQLX$ + "AND Nume_Clie = " & CLISEL% & " "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechMov)) = " & FEMOVD
    With Movisto
25    .FindFirst SQLX$
      If .NoMatch = False Then
        .Delete
        GoTo 25
      End If
    End With
    '
    Call COMUNI("Movimiento " + REMISEL$ + " Borrado.")
    GoTo 10
    '
99  Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    Command25.Enabled = True
    '
End Sub

Private Sub Command3_Click()
    '
    Command3.Enabled = False
    Screen.MousePointer = 11
        Call genevecli
        Call FINAL("*LISEVCLI")
    Screen.MousePointer = 1
    Command3.Enabled = True
    '
End Sub
   Sub genevecli()
   Screen.MousePointer = 11
   Call ABREORTRA
   SQLX$ = "Select * from Ordetrab "
   SQLX$ = SQLX$ + " ORDER BY Nclie_Otrab "
   
   Set OtraCLITEMP = OTrabajo.OpenRecordset(SQLX$, dbOpenSnapshot)
   With OtraCLITEMP
   On Error Resume Next
       .MoveLast
       'Valida que la tabla no este vacia
       If Err Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         On Error GoTo 0
         GoTo 90
       End If
    On Error GoTo 0
   FIN& = .RecordCount
   .MoveFirst
  JJ& = 0
  Do While (.EOF = False)
     XXX$ = REGBLAN$("LISEVCLI")
         '
         Call REPLA(XXY$, MKI$(!Nclie_Otrab), 1, 2)
         Call REPLA(XXY$, !Nro_Otrab, 3, 16)
         Call REPLA(XXY$, !Refe_Otrab, 19, 64)
         Call REPLA(XXY$, MKI$(CVINT%(!Fentre_Otrab)), 83, 2)
         Call REPLA(XXY$, !NroOfer_Otrab, 85, 16)
         Call REPLA(XXY$, MKI(CVINT(!FeOfer_Otrab)), 101, 2)
         Call REPLA(XXY$, !NroOcom_Otrab, 103, 32)
         Call REPLA(XXY$, MKI(CVINT(!FeOcom_Otrab)), 135, 2)
         If IsNull(!NuRev_Otrab) Then
               NUREV$ = ""
         Else
               NUREV$ = !NuRev_Otrab
         End If
         Call REPLA(XXY$, NUREV$, 137, 6)
         Call REPLA(XXY$, MKI(CVINT(!FeRev_Otrab)), 143, 2)
         '
         JJ& = JJ& + 1
         Call GRAREG("LISEVCLI", XXY$, JJ&)
     .MoveNext
    Loop
   End With
   '
   Call SETULTREG("LISEVCLI", JJ&)
   Call CIERRARCH("LISEVCLI")
90 Screen.MousePointer = 1
   '
   End Sub
   Sub imprefec()
   Screen.MousePointer = 11
   Call ABREORTRA
   SQLX$ = "Select * from Ordetrab "
   SQLX$ = SQLX$ + " ORDER BY Fentre_Otrab "
   
   Set OtraTEMP = OTrabajo.OpenRecordset(SQLX$, dbOpenSnapshot)
   With OtraTEMP
   On Error Resume Next
       .MoveLast
       'Valida que la tabla no este vacia
       If Err Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         On Error GoTo 0
         GoTo 90
       End If
    On Error GoTo 0
   FIN& = .RecordCount
   .MoveFirst
  JJ& = 0
  Do While (.EOF = False)
     XXX$ = REGBLAN$("LISOTFEC")
         '
         Call REPLA(XXY$, !Nro_Otrab, 1, 16)
         Call REPLA(XXY$, !Refe_Otrab, 17, 64)
         Call REPLA(XXY$, MKI$(!Nclie_Otrab), 81, 2)
         Call REPLA(XXY$, MKI$(CVINT%(!Fentre_Otrab)), 83, 2)
         Call REPLA(XXY$, !NroOfer_Otrab, 85, 16)
         Call REPLA(XXY$, MKI(CVINT(!FeOfer_Otrab)), 101, 2)
         Call REPLA(XXY$, !NroOcom_Otrab, 103, 32)
         Call REPLA(XXY$, MKI(CVINT(!FeOcom_Otrab)), 135, 2)
         If IsNull(!NuRev_Otrab) Then
               NUREV$ = ""
         Else
               NUREV$ = !NuRev_Otrab
         End If
         Call REPLA(XXY$, NUREV$, 137, 6)
         Call REPLA(XXY$, MKI(CVINT(!FeRev_Otrab)), 143, 2)
         '
         JJ& = JJ& + 1
         Call GRAREG("LISOTFEC", XXY$, JJ&)
     .MoveNext
    Loop
   End With
   '
   Call SETULTREG("LISOTFEC", JJ&)
   Call CIERRARCH("LISOTFEC")
90 Screen.MousePointer = 1
   '
End Sub


Private Sub Command4_Click()
    Command4.Enabled = False
        Call LIENCPRE
        Call FINAL("*LIBDESVE")
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    '
    Command5.Enabled = False
    Call imprefec
    Call FINAL("*LIBDESVE")
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

Private Sub Command6_Click()
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
    Call ABRESTOCKS
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
        FF% = CVINT(!FechMov)
        '
        Z$ = Space$(40) + String$(8, 0)
        Call REPLA(Z$, REMI$, 1, 10)
        Call REPLA(Z$, RASOCLI$(CLI%), 11, 30)
        Call REPLA(Z$, MKS$(U&), 41, 4)
        Call REPLA(Z$, MKI$(FF%), 45, 2)
        Call REPLA(Z$, MKI$(CLI%), 47, 2)
        J& = J& + 1
        Call GRAREG("!SELREMI", Z$, J&)
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
    Call ABRESTOCKS
    SQLX$ = "DoPriCor = '" & REMISEL$ & "' "
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
    Call HELPONLINE("ORPASTA")
End Sub

Private Sub Command8_Click()
    Dim RESP%
    Dim NUMEOCXX$
    '
    Command8.Enabled = False
    MODO% = COMALTER%("Opciones de Salida\\Pantalla\Impresora\Cancelar")
    If MODO% < 1 Or MODO% > 2 Then GoTo 99
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
    If MODO% = 1 Then GoTo 56
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
    Call ABRESTOCKS
    '
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE (CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " OR CodiMovi =  'DV') "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    SQLX$ = SQLX$ + " AND ((ISNULL(DoSecLar) = true) or (DoSecLar = '') OR (DoSecLar LIKE 'P/F*')) "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) >= " & DES1# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechMov)) <= " & HAS1# & " "
    'PARA EVITAR COMPROBANTES ARRUINADOS
    'PARA EVITAR COMPROBANTES ARRUINADOS
    SQLX$ = SQLX$ + " AND LEFT$(UCASE$(REFERCOR),4) <> 'ANUL' "
    'FILTRO CLIENTE
    If MODO% <> 1 Then
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
         Call REPLA(RECONSI$, MKI$(CVINT(!FechMov)), 3, 2)
         Call REPLA(RECONSI$, "", 5, 18)
         Call REPLA(RECONSI$, !DoPriCor, 23, 12)
         Call REPLA(RECONSI$, MKI$(!COD_INTE), 35, 2)
         Call REPLA(RECONSI$, MKD$(!CantSalid), 37, 8)
         Call REPLA(RECONSI$, MKD$(!CantIngre), 45, 8)
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
         Call REPLA(RECONSJ$, MKI$(CVINT(!FechMov)), 35, 2)
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
         Call REPLA(RECONSJ$, MKD$(!CantSalid), 65, 8)
         Call REPLA(RECONSJ$, MKD$(!CantIngre), 73, 8)
         Call REPLA(RECONSJ$, MKD$(!CantSaldo), 81, 8)
         Call REPLA(RECONSJ$, MKD$(!ValoUnit), 89, 8)
           TOTI# = !ValoUnit * (!CantSalid - !CantIngre)
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
    If MODO% = 1 Then
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
'    Command19.Enabled = False
'    If CONTROL$("REMSTAN", "REPLAORI") = "SI" Then
'      Command19.Enabled = True
'    End If
'    '
'    If EXISTE%("COCUMENT.EXE") > 0 Then
'      Command11.Enabled = True
'    End If
'    '
End Sub


