VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form REGTRA07 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Trazabilidad de Componentes"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Reservas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   135
      TabIndex        =   26
      Top             =   3450
      Width           =   8400
      Begin VB.Frame Frame5 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   6600
         TabIndex        =   27
         Top             =   480
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   5565
         TabIndex        =   29
         Top             =   480
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "                "
         Height          =   1800
         Left            =   2000
         TabIndex        =   28
         Top             =   480
         Width           =   15
      End
      Begin VB.ListBox List2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         ItemData        =   "ASIGLOTE07.frx":0000
         Left            =   190
         List            =   "ASIGLOTE07.frx":0002
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   600
         Width           =   8025
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Código                    Descripción                                                  Cant.   Lote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   195
         TabIndex        =   31
         Top             =   360
         Width           =   8025
      End
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
      Left            =   8190
      TabIndex        =   0
      Top             =   180
      Width           =   175
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Actualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   135
      TabIndex        =   14
      Top             =   5925
      Width           =   8400
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
         Height          =   300
         Left            =   7080
         TabIndex        =   36
         Top             =   840
         Width           =   175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   7560
         Picture         =   "ASIGLOTE07.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   420
         Width           =   660
      End
      Begin VB.Label Label6 
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
         Left            =   2100
         TabIndex        =   38
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label4 
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
         Left            =   5565
         TabIndex        =   37
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Origen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -940
         TabIndex        =   32
         Top             =   840
         Width           =   1770
      End
      Begin VB.Label Label10 
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
         Left            =   945
         TabIndex        =   21
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   -940
         TabIndex        =   20
         Top             =   420
         Width           =   1770
      End
      Begin VB.Label Label8 
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
         Left            =   945
         TabIndex        =   19
         Top             =   420
         Width           =   1800
      End
      Begin VB.Label Label5 
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
         Left            =   3555
         TabIndex        =   18
         Top             =   855
         Width           =   1155
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   17
         Top             =   855
         Width           =   1770
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2625
         TabIndex        =   16
         Top             =   855
         Width           =   810
      End
      Begin VB.Label Label1 
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
         Left            =   2940
         TabIndex        =   3
         Top             =   420
         Width           =   4425
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consumos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   135
      TabIndex        =   8
      Top             =   945
      Width           =   8400
      Begin VB.Frame Frame1 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   6600
         TabIndex        =   25
         Top             =   480
         Width           =   15
      End
      Begin VB.Frame Frame14 
         Caption         =   "                "
         Height          =   1800
         Left            =   2000
         TabIndex        =   13
         Top             =   480
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   5565
         TabIndex        =   10
         Top             =   480
         Width           =   15
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         ItemData        =   "ASIGLOTE07.frx":030E
         Left            =   190
         List            =   "ASIGLOTE07.frx":0310
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   8025
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Código                    Descripción                                                  Cant.   Lote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   195
         TabIndex        =   9
         Top             =   360
         Width           =   8025
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   8400
      Left            =   8640
      ScaleHeight     =   8340
      ScaleWidth      =   1320
      TabIndex        =   4
      Top             =   -20
      Width           =   1380
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
         Height          =   690
         Left            =   120
         Picture         =   "ASIGLOTE07.frx":0312
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Importacion de Articulos Trazables"
         Top             =   5040
         Width           =   690
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Excel 5.0;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   0
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   5880
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
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
         Height          =   690
         Left            =   105
         Picture         =   "ASIGLOTE07.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Acceso a Maestro de Items de Stock"
         Top             =   2100
         Width           =   690
      End
      Begin VB.CommandButton Command16 
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
         Left            =   120
         Picture         =   "ASIGLOTE07.frx":0766
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Consulta de Trazabilidad de Componentes"
         Top             =   4305
         Width           =   690
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Kdx"
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
         Picture         =   "ASIGLOTE07.frx":0A70
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   3570
         Width           =   690
      End
      Begin VB.CommandButton Command11 
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
         Picture         =   "ASIGLOTE07.frx":0D7A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir"
         Top             =   225
         Width           =   690
      End
      Begin VB.CommandButton Command10 
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
         Left            =   120
         Picture         =   "ASIGLOTE07.frx":0EC4
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   975
         Width           =   690
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   6
         Top             =   6525
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -285
         Picture         =   "ASIGLOTE07.frx":11CE
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   6810
         Width           =   1515
      End
   End
   Begin VB.Label Label13 
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
      Left            =   2700
      TabIndex        =   24
      Top             =   585
      Width           =   5655
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
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
      Left            =   -705
      TabIndex        =   23
      Top             =   315
      Width           =   1545
   End
   Begin VB.Label Label11 
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
      Left            =   525
      TabIndex        =   22
      Top             =   585
      Width           =   1980
   End
   Begin VB.Label Label28 
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
      Left            =   6735
      TabIndex        =   12
      Top             =   180
      Width           =   1500
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Número de O/F"
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
      Left            =   4950
      TabIndex        =   11
      Top             =   225
      Width           =   1545
   End
End
Attribute VB_Name = "REGTRA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Command1.Enabled = False
3   Call SELECHO("ECORDEF/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
    NORFA$ = TRIM$(VALACT1$("ECORDEF"))
    If NORFA$ = "" Then GoTo 99
    '
    STATORF% = XVALO(VALOBADA("ORDEFABR", "StatuOrf", "IdSubOr = '" & NORFA$ & "'"))
    If STATORF% >= 8 Then
       Call MENSERR(24, "Orden '" + NORFA$ + "' Anulada")
       GoTo 99
    End If
    '
    Label28 = TRIM$(NORFA$)
    '
99  Command1.Enabled = True

End Sub

Private Sub Command11_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")

End Sub

Private Sub Command16_Click()
    '
    If DERACCE("CONTRAZA", "Consulta de Trazabilidad") > 1 Then
       CORETRA07.Show 1
    End If
    '
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   CODILLA$ = TRIM$(Label8)
   If CODILLA$ = "" Then CODILLA$ = TRIM$(Label11)
   FIKARDEX07.Text1 = CODILLA$
   FIKARDEX07.Show 1
   Command2.Enabled = True

End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    If DERACCE("REGTRAZA", "Registro de Trazabilidad") > 1 Then
    
        CI% = CODINT(Label8)
        If CI% < 1 Then Call MENSERR(24, "Falta seleccionar Item"): GoTo 99
    
        NUVAL = TRIM$(UCase(Label4))
        If NUVAL = "" Then Call MENSERR(24, "Falta Ingresar el Lote")
        NORFA$ = Label28
        '
        If Label10 = "Consumos" Then
          CONDI$ = "DopriCor = '" & NORFA$ & "' AND CodiMovi = 'CF' AND Cod_Inte = " & CI% & " AND OrdConsum = " & TRIM$(Label6)
          LOTEVIE$ = TRIM$(VALOBADA("MOVISTO", "IdenLote", CONDI$))
          Call ACTUREGISTRO("Movisto", "IdenLote", CONDI$, NUVAL, REGAF&, "NOMESS")
          If TRIM$(LOTEVIE$) <> "" Then Call ACSALOTE(CI%, LOTEVIE$, "NOMESS")
          NUVALX$ = NUVAL
          Call ACSALOTE(CI%, NUVALX$, "NOMESS")
          CONDIC$ = " DopriCor = '" & NORFA$ & "' AND CodiMovi = 'CF'"
          Call CARGALISTA(List1, "MOVISTO", "Cod_Exte/A16; Descrip/A32; CantSalid/F10.1;IdenLote/A16; OrdConsum/F12", CONDIC$)
10        ' ELIMINA LOS ITEMS NO TRAZABLES
          For I& = 1 To List1.ListCount
            CIXI% = CODINT%(TRIM$(Left$(List1.List(I& - 1), 16)))
            If ESTRAZABLE%(CIXI%) < 1 Then
               List1.RemoveItem (I& - 1)
               GoTo 10
            End If
          Next I&
          GoTo 90
        End If
        
        If Label10 = "Reservas" Then
          CONDI1$ = " IdSubOr = '" & NORFA$ & "' AND Cod_Inte = " & CI% & ""
          Call ACTUREGISTRO("Reserva", "IdenLote", CONDI1$, NUVAL, REGAF&, "NOMESS")
          CONDIC1$ = " IdSubOr = '" & NORFA$ & "'"
          Call CARGALISTA(List2, "RESERVA", "Cod_Exte/A16; Descrip/A32; (CantRes - CantCons)/F10.1;IdenLote/A16", CONDIC1$)
20        ' ELIMINA LOS ITEMS NO TRAZABLES
          For I& = 1 To List2.ListCount
            CIXI% = CODINT%(TRIM$(Left$(List2.List(I& - 1), 16)))
            If ESTRAZABLE%(CIXI%) < 1 Then
               List2.RemoveItem (I& - 1)
               GoTo 20
            End If
          Next I&
          GoTo 90
        End If
        '
90      Call ACSALOTE(CI%, TRIM$(UCase(Label4)), "NOMESS")
        Call Label8_Change
    End If
    '
99  Command3.Enabled = True

End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    If DERACCE%("AMASTO07", "Maestro de Items de Stock") > 1 Then
       AMASTOK07.Text1 = Label11
       AMASTOK07.Show 1
       Call Label28_Change
    End If
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    '
    Call SELECHO("!LILOTSE")
    If VALACT1$("!LILOTSE") <> "" Then
       Label4 = VALACT1$("!LILOTSE")
    End If
    '
End Sub

Private Sub Command6_Click()
    '
    Command6.Enabled = False
    '
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Archivo Origen de Datos"
    '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
    CommonDialog1.Filter = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0

    DoEvents
    Screen.MousePointer = 11
    Data1.DatabaseName = UCase$(CommonDialog1.filename)
    On Error Resume Next
    Data1.Refresh
    Data1.RecordSource = Data1.Database.TableDefs(0).Name
    Data1.Refresh
    On Error GoTo 0
    
    With Data1
     .Recordset.MoveFirst
     Do While Not .Recordset.EOF
        CODEXTT$ = ""
        CODEXTT$ = .Recordset.Fields(0)
        CI% = CODINT%(CODEXTT$)
        CONDI$ = "Codint = " & CI% & ""
        If CI% <= 0 Then
            Call MENSERR(24, "Código '" & CODEXTT$ & "' Inexistente o Dado de Baja.")
            GoTo 20
        End If
        TRAZABL% = XVALO(VALOBADA("MASTER", "Trazable", CONDI$))
        If TRAZABL% = 2 Then
            Call MENSERR(24, "Código " & CODEXTT$ & " Trazable por Número de Serie.")
            GoTo 20
        End If
        '
        ESTRAZA% = 1
        '
        Call ACTUREGISTRO("MASTER", "Trazable", CONDI$, ESTRAZA%, REGAF&)
20      Screen.MousePointer = 11
       .Recordset.MoveNext
     Loop
    End With
    Screen.MousePointer = 1
    Call MENSERR(24, "Proceso de Actualización Masiva Completo")
    '
99  Command6.Enabled = True
    '
End Sub

Private Sub Form_Load()
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    Call ABRECONEXION
    '
    If APLINVO$ = "CONSUTRA" Then
       CORETRA07.Show 1
       Call FINAL("")
    End If
End Sub

Private Sub Label28_Change()
   Call LIMPIAR
   NORFA$ = Label28
   CI% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", "IdSubOr = '" & NORFA$ & "'"))
   If CI% > 0 Then
     Label11 = CODEXT(CI%)
     Label13 = DESCRIT$(CI%)
     CONDI$ = " DopriCor = '" & NORFA$ & "' AND CodiMovi = 'CF'"
     Call ENUMERA("MOVISTO", "OrdConsum", CONDI$)
     Call CARGALISTA(List1, "MOVISTO", "Cod_Exte/A16; Descrip/A32; CantSalid/F10.1;IdenLote/A16; OrdConsum/F12", CONDI$)
     I& = 1
     ' ELIMINA LOS ITEMS NO TRAZABLES
     Do While I& <= List1.ListCount
       CIXI% = CODINT%(TRIM$(Left$(List1.List(I& - 1), 16)))
       If ESTRAZABLE%(CIXI%) > 0 Then
         I& = I& + 1
       Else
         List1.RemoveItem (I& - 1)
       End If
     Loop
     '
     CONDI1$ = " IdSubOr = '" & NORFA$ & "'"
     Call CARGALISTA(List2, "RESERVA", "Cod_Exte/A16; Descrip/A32; (CantRes - CantCons)/F10.1;IdenLote/A16", CONDI1$)
20   ' ELIMINA LOS ITEMS NO TRAZABLES
     For I& = 1 To List2.ListCount
        CIXI% = CODINT%(TRIM$(Left$(List2.List(I& - 1), 16)))
        If ESTRAZABLE%(CIXI%) < 1 Then
           List2.RemoveItem (I& - 1)
           GoTo 20
        End If
     Next I&
     '
     If List1.ListCount + List2.ListCount < 1 Then
        Call MENSERR(24, "Este Producto no Registra Componentes Trazables.")
        Call Command4_Click
     End If
     '
   End If

End Sub

Private Sub Label8_Change()
     CIXI% = CODINT%(Label8)
     If CIXI% > 0 Then
        Call CREALOTES(CIXI%)
        Call ARMALISLOTES(CIXI%, , , , 1)
     End If
     'Call AUTO_CREALOT(CIXI%)
End Sub

Private Sub List1_DblClick()
     Call LIMPIAR
     TTXX$ = List1.List(List1.ListIndex)
     Label8 = Mid$(TTXX$, 1, 16)
     Label10 = "Consumos"
     Label1 = TRIM$(Mid$(TTXX$, 17, 32))
     Label5 = Mid$(TTXX$, 50, 11)
     Label4 = Mid$(TTXX$, 62, 16)
     Label6 = TRIM$(Mid$(TTXX$, 79, 12))
End Sub

Private Sub List2_DblClick()
     Call LIMPIAR
     TTXX$ = List2.List(List2.ListIndex)
     Label8 = Mid$(TTXX$, 1, 16)
     Label10 = "Reservas"
     Label1 = TRIM$(Mid$(TTXX$, 17, 32))
     Label5 = Mid$(TTXX$, 50, 11)
     Label4 = Mid$(TTXX$, 62, 16)
End Sub

Sub LIMPIAR()
     Label8 = ""
     Label10 = ""
     Label1 = ""
     Label5 = ""
     Label4 = ""
End Sub

