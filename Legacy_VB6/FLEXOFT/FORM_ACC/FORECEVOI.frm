VERSION 5.00
Begin VB.Form FORECEVOI 
   BackColor       =   &H00E0E0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recepción de Materiales - Altas de Stock en Depósitos"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9720
   Icon            =   "FORECEVOI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0F0&
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
      Height          =   2910
      Left            =   4515
      TabIndex        =   2
      Top             =   210
      Width           =   4110
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0F0&
         Caption         =   "Consulta de Pendientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Left            =   210
         TabIndex        =   9
         Top             =   315
         Width           =   3690
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   16
            Top             =   1545
            Width           =   2220
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
               Left            =   -20
               Picture         =   "FORECEVOI.frx":030A
               Style           =   1  'Graphical
               TabIndex        =   17
               ToolTipText     =   "Material en Consignación"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Material en Consignación"
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
               Left            =   735
               TabIndex        =   18
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   13
            Top             =   1005
            Width           =   2220
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
               Left            =   -20
               Picture         =   "FORECEVOI.frx":0614
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Ordenes de Compra Pendientes de Recepción"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Pedidos y Orde- nes de Compra"
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
               Left            =   735
               TabIndex        =   15
               Top             =   50
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   10
            Top             =   510
            Width           =   2220
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
               Picture         =   "FORECEVOI.frx":091E
               Style           =   1  'Graphical
               TabIndex        =   11
               ToolTipText     =   "Detalle por Boleta de Recepción"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Boletas de Recepción"
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
               Left            =   735
               TabIndex        =   12
               Top             =   50
               Width           =   1590
            End
         End
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0F0&
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   4515
      TabIndex        =   33
      Top             =   3180
      Width           =   4110
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0F0&
         Caption         =   "Pantalla"
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
         Left            =   630
         TabIndex        =   35
         Top             =   320
         Width           =   1380
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0F0&
         Caption         =   "Impresión"
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
         Left            =   2310
         TabIndex        =   34
         Top             =   320
         Value           =   -1  'True
         Width           =   1170
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   1155
      ScaleHeight     =   585
      ScaleWidth      =   2265
      TabIndex        =   25
      Top             =   6720
      Width           =   2325
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
         Left            =   -20
         Picture         =   "FORECEVOI.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Devolución       de Materiales"
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
         TabIndex        =   27
         Top             =   105
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H00000080&
      ForeColor       =   &H00000000&
      Height          =   6210
      Left            =   8820
      ScaleHeight     =   6150
      ScaleWidth      =   1005
      TabIndex        =   20
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   37
         Top             =   3090
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command17 
         Caption         =   "EF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FORECEVOI.frx":0D72
         TabIndex        =   36
         ToolTipText     =   "Edición de Fecha de Recepción"
         Top             =   5145
         Visible         =   0   'False
         Width           =   650
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "FORECEVOI.frx":107C
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Ver y Re-Imprimir B.Recepción"
         Top             =   1050
         Width           =   650
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "FORECEVOI.frx":1386
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Acceso Directo a otras Aplicaciones"
         Top             =   2340
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
         Picture         =   "FORECEVOI.frx":1690
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Terminar - Salir de esta Aplicación"
         Top             =   315
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
         Picture         =   "FORECEVOI.frx":17DA
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Configuración de Comportamiento y Formularios"
         Top             =   1695
         Width           =   650
      End
      Begin VB.CommandButton Command10 
         Caption         =   "QC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FORECEVOI.frx":1C1C
         TabIndex        =   21
         ToolTipText     =   "Calificación de Proveedores"
         Top             =   4515
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -240
         Picture         =   "FORECEVOI.frx":1F26
         Stretch         =   -1  'True
         Top             =   3405
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0F0&
      Caption         =   "Ingresos a Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4110
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0F0&
         Caption         =   "Gestión de Compras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   210
         TabIndex        =   28
         Top             =   2200
         Width           =   3690
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   29
            Top             =   420
            Width           =   2430
            Begin VB.CommandButton Command13 
               Caption         =   "OC"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   0
               TabIndex        =   30
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Consulta de Or- denes de Compra"
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
               Left            =   735
               TabIndex        =   31
               Top             =   105
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0F0&
         Caption         =   "Recepciones Externas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1730
         Left            =   210
         TabIndex        =   1
         Top             =   315
         Width           =   3690
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   6
            Top             =   900
            Width           =   2325
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
               Height          =   600
               Left            =   0
               Picture         =   "FORECEVOI.frx":3E48
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Inspección y Aprobación"
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
               Left            =   735
               TabIndex        =   8
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   3
            Top             =   420
            Width           =   2325
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
               Picture         =   "FORECEVOI.frx":4152
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Boleta de Recep- ción (Provisoria)"
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
               Left            =   630
               TabIndex        =   5
               Top             =   0
               Width           =   1590
            End
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2940
            TabIndex        =   39
            Top             =   420
            Visible         =   0   'False
            Width           =   645
         End
      End
   End
   Begin VB.Label NORVENTA 
      Caption         =   "NORVENTA"
      Height          =   225
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
End
Attribute VB_Name = "FORECEVOI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HAYMATERIALIB%
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    FORECEVOI.Hide
End Sub

Private Sub Command10_Click()
Command10.Enabled = False
Call CONECRUN("CALPRO02", "Sistema Calificación de Proveedores")
Command10.Enabled = True
End Sub


  
Private Sub Command12_Click()
   Command12.Enabled = False
3  Call SELECHO("INDIBOL")
   NPX& = Val(VALACT1$("INDIBOL"))
   If NPX& > 0 Then
     Call SHOWBORE(NPX&)
     GoTo 3
   End If
   '
99 Command12.Enabled = True
End Sub

Private Sub Command13_Click()
 Command13.Enabled = False
     'agrego todas las O/C de ocomenca
      Call ABRECOMPRAS
      SQLX$ = " SELECT Nume_Ocom, Refe_Ocom FROM OCOMENCA"
      SQLX$ = SQLX$ + " ORDER BY Nume_Ocom  DESC "
      Set VEROCTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
      With VEROCTEMP
         On Error Resume Next ' si hay algun error msgbox err
         .MoveFirst
         If Err Then
           Call MENSERR(24, "No hay Ordenes de Compra Emitidas")
           On Error GoTo 0
           GoTo 99
         End If
         On Error GoTo 0
         '
         JJ& = 0
         Do While (.EOF = False)
             XXY$ = REGBLAN$("INDIOCOM")
             Call REPLA(XXY$, MKS(!NUME_OCOM), 1, 4)
             Call REPLA(XXY$, !Refe_OcoM, 5, 44)
             JJ& = JJ& + 1
             Call GRAREG("INDIOCOM", XXY$, JJ&)
             .MoveNext
         Loop
      End With
      Set VEROCTEMP = Nothing
      Call SETULTREG("INDIOCOM", JJ&)
      Call CIERRARCH("INDIOCOM")
      '
3     Call SELECHO("INDIOCOM/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
      
      NPX& = Val(VALACT1$("INDIOCOM"))
      If NPX& > 0 Then
        Call SHOWOCOM(NPX&)
        GoTo 3
      End If
      '
99  Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   '
   OPRECEP04.Show 1
   '
End Sub

Private Sub Command17_Click()
    Command17.Enabled = False
    Y& = Val(InputBox("Ingrese el numero de boleta de Recepcion"))
    If Y& > 0 Then
      NUFECHA = CDate(InputBox("Nueva Fecha"))
      Call ABRESTOCKS
      SQLX$ = "NumeBoRe = " & Y&
      With Bolrecep
        .FindFirst SQLX$
10        If .NoMatch = False Then
          .Edit
          !FechRecep = NUFECHA
          .Update
          .FindNext SQLX$
          GoTo 10
        End If
      End With
      Bolrecep.Close
      '
      SQLX$ = "DoPriCor = 'BR." & TRIM$(Str$(Y&)) & "'"
      With Movisto
        .FindFirst SQLX$
20        If .NoMatch = False Then
          .Edit
          !FechMov = NUFECHA
          .Update
          .FindNext SQLX$
          GoTo 20
        End If
      End With
      Movisto.Close
      '
    End If
    '
    Command17.Enabled = True
End Sub


Private Sub Command19_Click()
    '
    Call COPYSTRU("BOLREDET", "BOLRDETA")
    '
    '
1100 NROBOL& = 1
     Call ABRESTOCKS
     Set BolreTemp = Stocks.OpenRecordset("SELECT DISTINCT NumeBoRe from BOLRECEP ORDER BY NumeBoRe DESC")
     On Error Resume Next
     BolreTemp.MoveFirst
     If Err < 1 Then
       NROBOL& = 1 + BolreTemp!NumeBore
     End If
     On Error GoTo 0
     '
     If TIMATEI% = 0 Then
         X$ = REGLEIDO$("TATIREC", 1)
         TIMATEI% = Asc(Left$(X$, 1))
     End If
     '
     VDEF$ = MKI$(NC%) + Chr$(TIMATEI%) + CODICLI$((-1) * NC%)
     EBOL$ = OBJPLA$("SELPROVRECACC", VDEF$)
     If EBOL$ = "" Then
        Exit Sub
     End If
     '
     NC% = NUMEPRO%(TRIM$(Mid$(EBOL$, 4, 12)))
     If NC% < 1 Or STATCLI%((-1) * NC%) < 1 Then
       Call MENSERR(24, "Proveedor Inexistente")
       GoTo 1100
     End If
     '
     TIMATEI% = Asc(Mid$(EBOL$, 3, 1))
     If TIMATEI% < 1 Or ECOARCH("TATIREC", Chr$(TIMATEI%)) = "" Then
         Call MENSERR(24, "Falta Tipo de Material")
         GoTo 1100
     End If
     '
     REGT& = TIMATEI%
     X$ = REGLEIDO$("TATIREC", REGT&)
     FORMPLA$ = TRIM$(UCase$(Mid$(X$, 33, 16)))
     FORMTAB$ = TRIM$(UCase$(Mid$(X$, 49, 16)))
     FORIPRE$ = TRIM$(UCase$(Mid$(X$, 65, 16)))
     '
     NC1% = (-1) * NC%
     '
1105      Y$ = PEDY$
     If Y$ = "" Then
       Y$ = Space$(128)
       Call REPLA(Y$, Chr$(0), 121, 1)
     End If
     '
     Call REPLA(Y$, MKS$(NROBOL&), 1, 4)
     Call REPLA(Y$, EBOL$, 7, 2)
     Call REPLA(Y$, MKI$(HOY(HO$)), 5, 2)
     Call REPLA(Y$, RASOCLI$(NC1%), 9, 32)
     Call REPLA(Y$, DOMICLI$(NC1%), 41, 32)
     CLOCA$ = TRIM$(CPOSCLI$(NC1%)): If CLOCA$ <> "" Then CLOCA$ = CLOCA$ + " "
     Call REPLA(Y$, CLOCA$ + LOCACLI$(NC1%), 73, 32)
     Call REPLA(Y$, Chr$(0), 121, 1)
    '
     OBJE$ = "BOLREC-ENCB"
     If CONTROL$("", "NUMEBREC") = "MANUAL" Then
         OBJE$ = "BOLREC-ENCB-M"
     End If
     '
     TICOMPRO$ = TRIM$(CONTROL$("", "COMPBREC"))
     '
     PEDY$ = OBJPLA$(OBJE$, Y$)
     If PEDY$ = "" Then Exit Sub
     '
     FF% = CVI(Mid$(PEDY$, 5, 2))
     FEX = FF%: FECHAREC$ = FECHATEX$(FEX)
     If TRIM$(FECHAREC$) = "" Then
         Call MENSERR(24, "Fecha Incorrecta")
         GoTo 1105
     End If
     '
     REMI$ = TRIM$(Mid$(PEDY$, 105, 16))
     If REMI$ = "" Then
         Call MENSERR(24, "Falta Número de Remito del Proveedor")
         GoTo 1105
     End If
     '
     DP1% = Asc(Mid$(PEDY$, 121, 1))
     DEPOSAL$ = TRIM$(ECOARCH$("ECODEP", Chr$(DP1%)))
     If DP1% < 1 Or DEPOSAL$ = "" Then
         Call MENSERR(24, "Falta Depósito de Entrada")
         GoTo 1105
     End If
     '
     NROPRO& = NROBOL&
     NROBOL& = CVS(Left$(PEDY$, 4))
     If NROBOL& < 1 Then
         Call MENSERR(24, "Número Incorrecto de B.R.")
         NROBOL& = NROPRO&
         Call REPLA(PEDY$, MKS$(NROBOL&), 1, 4)
         GoTo 1105
     End If
     '
1112 Call CIERRARCH("*.*")
     JJ& = 0
     '
     DARX$ = TRIM$(CONTROL$("BOLRECEP", "DIANREC"))
     If DARX$ = "" Then
          DIANRE& = 10000
        Else
          DIANRE& = Val(DARX$)
     End If
     HOIX1 = HOY(HOS$)
     '
     anuca% = 0
     If CONTROL$("BOLRECEP", "ANUCANTI") = "SI" Then
       anuca% = 1
     End If
     '
     Screen.MousePointer = 11
     Call ABRECOMPRAS
     SQLX$ = "Select * from OCOMDETA"
     SQLX$ = SQLX$ + " Where Stat_Ocom < 8 "
     SQLX$ = SQLX$ + " AND  Stat_Ocom <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
     SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
     SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
     SQLX$ = SQLX$ + " AND Cod_Inte = 0 " '0/C's TEXTO LIBRE
     Set COMPTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With COMPTEMP
       '
       On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Screen.MousePointer = 1
         HAYOC% = 0
         'GoTo 1777
       End If
       On Error GoTo 0
       '
       JJ& = 0
        
       Do While (.EOF = False)
         CAPEN = !Cant_Ocom - !Cant_Rec
         COEFUNI = !Coef_Unids
         If COEFUNI < 0.0001 Then COEFUNI = 1
         If CAPEN > 0.005 Then
           NUOC& = !NUME_OCOM
           FENTRESO% = CVINT(!Fentre_Sol)
           FESO2 = FENTRESO%
           If DIENTRE%(HOIX1, FESO2) <= DIANRE& Then
             CI% = !COD_INTE
             LOTECO$ = !Doc_Orig
             Y$ = REGBLAN$("!BOLREDET")
             Call REPLA(Y$, MKI$(CI%), 83, 2)
             Call REPLA(Y$, UNIMED$(CI%), 85, 4)
             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 89, 8)
             If anuca% = 1 Then
               Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 163, 8)
                Call REPLA(Y$, MKD$(0), 89, 8)
             End If
             '
             Call REPLA(Y$, MKI$(!NuOrd_Item), 98, 2)
             '
             Call REPLA(Y$, LOTECO$, 107, 12)
             Call REPLA(Y$, MKS$(NUOC&), 151, 4)
             Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
             Call REPLA(Y$, MKS$(!N_OrServ), 189, 4)
             HAYOC% = 1
             JJ& = JJ& + 1
             Call GRAREG("!BOLREDET", Y$, JJ&)
           End If
         End If
        .MoveNext
       Loop
     End With
     '
     
End Sub

Private Sub Command2_Click()
    Call APROBOL
End Sub



Private Sub Command3_Click()
    If Option1.Value = True Then
        Call COBORECEP
        Call FINAL("*COBOLPE")
      ElseIf Option2.Value = True Then
        Call COBORECEP
        If HAYMATERIALIB% = 0 Then
            Call FINAL("*LIBOLPE")
        Else
            Call FINAL("*LIBOPEL")
        End If
    End If
End Sub

Private Sub Command4_Click()
    '
    Call Selcopen
3   Call SELECHO("INDIOCPE/Indice de Ordenes de Compra Pendientes")
    NPX& = Val(VALACT1$("INDIOCPE"))
    If NPX& > 0 Then
      Call SHOWOCOM(NPX&)
      GoTo 3
    End If
    '
End Sub

Private Sub Command5_Click()
    '
    Screen.MousePointer = 11
    Call ABRESTOCKS
    SQLX$ = "Select * from MACONSIG"
    SQLX$ = SQLX$ + " WHERE Cant_Rendid < Cant_Entreg - 0.001 "
    Set consigtmp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With consigtmp
      '
      JJ& = 0
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
          ZZ$ = REGBLAN$("MACONSIG")
          Call REPLA(ZZ$, MKI$(CVINT(!Femi_Ocom)), 1, 2)
            NC% = !Npro_Ocom
          Call REPLA(ZZ$, MKI$(NC%), 3, 2)
            NORFA$ = !Nro_Orfa
          Call REPLA(ZZ$, NOCOM$, 9, 12)
            NOCOM$ = !Nro_Ocom
          Call REPLA(ZZ$, NOCOM$, 21, 12)
            CI% = !COD_INTE
          Call REPLA(ZZ$, MKI$(CI%), 33, 2)
          Call REPLA(ZZ$, UNIMED$(CI%), 51, 4)
            CAN# = !Cant_Asig
          Call REPLA(ZZ$, MKD$(CAN#), 55, 8)
            CAN# = !Cant_Entreg
          Call REPLA(ZZ$, MKD$(CAN#), 63, 8)
            FF% = CVINT%(!Fecha_Entreg)
          Call REPLA(ZZ$, MKI$(FF%), 71, 2)
            CAN# = !Cant_Rendid
          Call REPLA(ZZ$, MKD$(CAN#), 111, 8)
            CAN# = !Saldo_Pend
          Call REPLA(ZZ$, MKD$(CAN#), 119, 8)
            '
          JJ& = JJ& + 1
          Call GRAREG("MACONSIG", ZZ$, JJ&)
        .MoveNext
        Loop
      End If
    End With
    Call SETULTREG("MACONSIG", JJ&)
    Screen.MousePointer = 1
    '
    If Option1.Value = True Then
        Call FINAL("*COMATCO")
      ElseIf Option2.Value = True Then
        Call FINAL("*LIMATCO")
    End If
    '
End Sub



Private Sub Command7_Click()
    Call BOLERECEP
End Sub

Private Sub Command8_Click()
    Call ACCDIR("RECEPMA")
End Sub

Private Sub Form_Load()
    '
    VERANTER$ = "RECEPMA"
    '
    'Call VERJOBID(OKEY%)
    'If OKEY% < 1 Then Call FINAL("")
    '
    If APLINVO$ = "CONRECEP" Then
      Frame2.Left = 210
      Picture11.Left = Picture11.Left - 4305
      Width = Width - 4305
      Call CENTRAFORM(Me)
      DoEvents
    End If
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    If EXISTE%("CALPRO02.EXE") > 0 Then
      Command10.Visible = True
    End If
    '
End Sub

Sub APROBOL() ' aprobacion de boleta de recepcion ************************************
        '
        Dim REGBO&(2048)
1100    Call GENIBOLPE
        '
1120    Call FRESHECHO("IBOLPEN")
        'Call CLEARVEN
        If ULTREG&("IBOLPEN") < 1 Then
          Exit Sub
        End If
        '
        X$ = "": Y$ = "": Z$ = "": YY$ = ""
        EBOL$ = "": PEDY$ = "": DESCRI$ = ""
        NBR$ = OBJPLA$("SELBOLPEN", "")
        If NBR$ = "" Then
            Exit Sub
        End If
        '
        NROBOL& = CVS(Left$(NBR$, 4))
        If NROBOL& < 1 Then
            Exit Sub
        End If
        '
        'Call BLIMESS("Seleccionando ...")
        FIN& = ULTREG&("BOLREPEN")
        Call BLANARCH("!BOLREDET")
        CABOPE% = 0
        For I& = 1 To FIN&
          X$ = REGLEIDO$("BOLREPEN", I&)
          If CVS(Left$(X$, 4)) = NROBOL& Then
            CANOM# = CVD(Mid$(X$, 89, 8))
            APRO# = CVD(Mid$(X$, 163, 8))
            RECHA# = CVD(Mid$(X$, 171, 8))
            FALTA# = CVD(Mid$(X$, 179, 8))
            '
            If CANOM# < 0.05 Or APRO# > 0.05 Or RECHA# > 0.05 Or Abs(FALTA#) > 0.05 Then GoTo 1121
            Call REPLA(X$, MKD$(CANOM#), 163, 8)
            '
            If CABOPE% = 0 Then
              NC% = CVI(Mid$(X$, 105, 2)): NC1% = (-1) * NC%
              FF% = CVI(Mid$(X$, 81, 2))
              FEX = FF%: FECHAREC$ = FECHATEX$(FEX)
              REMI$ = Mid$(X$, 119, 16)
              DP1% = Asc(Mid$(X$, 97, 1))
              If ECOARCH$("ECODEP", Chr$(DP1%)) = "" Then
                  DP1% = 0
              End If
              '
              Y$ = String$(128, 0)
              Call REPLA(Y$, MKS$(NROBOL&), 1, 4)
              Call REPLA(Y$, MKI$(FF%), 5, 2)
              Call REPLA(Y$, MKI$(NC%), 7, 2)
              Call REPLA(Y$, RASOCLI$(NC1%), 9, 32)
              Call REPLA(Y$, DOMICLI$(NC1%), 41, 32)
                  CLOCA$ = TRIM$(CPOSCLI$(NC1%)): If CLOCA$ <> "" Then CLOCA$ = CLOCA$ + " "
              Call REPLA(Y$, CLOCA$ + LOCACLI$(NC1%), 73, 32)
              Call REPLA(Y$, REMI$, 105, 16)
              Call REPLA(Y$, Chr$(DP1%), 121, 1)
              AINDI% = AINDI% + 1
            End If
            '
            CABOPE% = CABOPE% + 1: REGBO&(CABOPE%) = I&
            Call REGAPP("!BOLREDET", X$)
            '
          End If
          '
1121    Next I&
        '
        'Call BLIMESS("-")
        If CABOPE% < 1 Then
            Call COMUNI("No hay Recepciones Pendientes\para la Boleta Elegida")
            GoTo 1120
        End If
        '
11215   Z$ = OBJPLA$("BOLREC-MUESTRW", Y$)
        If Z$ = "" Then
            GoTo 1120
        End If
        DP1% = Asc(Mid$(Z$, 121, 1))
        If DP1% < 1 Or ECOARCH$("ECODEP", Chr$(DP1%)) = "" Then
          Call MENSERR(24, "Falta Indicar Depósito de Entrada")
          GoTo 11215
        End If
        '
        Call ABRESTOCKS
        ABUS$ = "NumeBoRe = " & NROBOL & " "
        With Bolrecep
          .FindFirst ABUS$
          If .NoMatch = True Then
            Call MENSERR(24, "Boleta de Recepción no Encontrada.")
            GoTo 1100
          End If
          '
          CIXBR% = !COD_INTE
          RESP5% = 1  ' Material Codificado
          If CIXBR% = 0 Then RESP5% = 2    ' mATERIAL lIBRE
        End With
        '
        If RESP5% = 2 Then
           Call APROCLI(NC%, FF%, NROBOL&, REMI$, DP1%)
           Screen.MousePointer = 1
           GoTo 1100
        End If
        'HASTA ACA NUEVO
        '
1122    VTB% = VENTABU%("APROBOL/NC/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(RASOCLI$(NC1%)) + " - " + FECHAREC$)
        ATX% = 0: HAYRECHA% = 0
        For U& = 1 To ULTREG&("!BOLREDET")
          X$ = REGLEIDO$("!BOLREDET", U&)
          CANOM# = CVD(Mid$(X$, 89, 8))
          APRO# = CVD(Mid$(X$, 163, 8))
          RECHA# = CVD(Mid$(X$, 171, 8))
            If RECHA# >= 0.05 Then HAYRECHA% = 1
          FALTA# = CVD(Mid$(X$, 179, 8))
          If Abs(APRO#) < 0.005 Then
            If Abs(RECHA#) < 0.005 Then
              If Abs(FALTA#) < 0.005 Then
                 GoTo 1123                        ' NO SE PROCESO MOVIMIENTO
              End If
            End If
          End If
          '
          FALTB# = CANOM# - APRO# - RECHA#
          If Abs(FALTB# - FALTA#) > 0.005 Then
            ATX% = ATX% + 1
            Call REPLA(X$, MKD$(FALTB#), 179, 8)
            Call GRAREG("!BOLREDET", X$, U&)
          End If
1123    Next U&
        '
        If ATX% > 0 Then
          Call COMUNI("Cantidades Incoherentes.\Revise y Corrija !")
          GoTo 1122
        End If
        '
1124    If VTB% <> 0 Then GoTo 1120
        '
        CODRECHA$ = "RD"
        REMIRECH$ = "RD." + TRIM$(Str$(NROBOL&))
        If HAYRECHA% > 0 Then
          If TRIM$(CONTROL$("", "REMIRECH")) <> "" Then
            CODRECHA$ = TRIM$(CONTROL$("", "CODRRECH"))
            If CODRECHA$ = "" Then CODRECHA$ = "RD"
          End If
        End If
        '
        QUEDAPEN% = 0
        HAYRECHA% = 0
        Call ABRESTOCKS
        FIN& = ULTREG&("!BOLREDET")
        '
        NOCLOSE% = 1
        Label12.Visible = True
        '
        For U& = 1 To FIN&
          Call TRACE(20, U&, FIN&)
          Label12.Caption = TRIM$(Str$(U&))
          DoEvents
          X$ = REGLEIDO$("!BOLREDET", U&)
          CI% = CVI(Mid$(X$, 83, 2))
          REFX$ = Mid$(X$, 5, 64)
          LOTE$ = Mid$(X$, 69, 12)
          CANOM# = CVD(Mid$(X$, 89, 8))
          APRO# = CVD(Mid$(X$, 163, 8))
          RECHA# = CVD(Mid$(X$, 171, 8))
          FALTA# = CVD(Mid$(X$, 179, 8))
          If Abs(APRO#) < 0.005 Then
            If Abs(RECHA#) < 0.005 Then
              If Abs(FALTA#) < 0.005 Then
                QUEDAPEN% = 1
                GoTo 1128                        ' NO SE PROCESO MOVIMIENTO
              End If
            End If
          End If
          '
          ABUS$ = "NumeBoRe = " & NROBOL & " "
          ABUS$ = ABUS$ + "AND IdenLote = '" & LOTE$ & "' "
          ABUS$ = ABUS$ + "AND Cod_Inte = " & CI%
          With Bolrecep
            .FindFirst ABUS$
            If .NoMatch = False Then
              .Edit
              !CanToApro = APRO#
              !CanToRech = RECHA#
              !CanToFalta = FALTA#
              !STATUS = 1
              !FechLibe = Date
              !USUAPRO = USERNAME   'NUEVO
              !FeHoApro = Now      'NUEVO
              !Observa = USERNAME 'nuevo
              .Update
              '
1127          DOPRI$ = "BR." + TRIM$(Str$(NROBOL&))
              DOPRILA$ = DOPRI$
              DOSELA$ = TRIM$(REMI$)
              DOSEC$ = TRIM$(REMI$)
              DEPX% = DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(CI%)
              If CANOM# >= 0.05 Then
                CANTI = CANOM#
                REFE$ = "Rto." + TRIM$(REMI$) + " - " + RASOCLI$(NC1%)
                PPX% = InStr(REFX$, "- BC D.")
                If PPX% > 0 Then
                    REFE$ = Mid$(REFX$, PPX% + 5)
                End If
                While Len(DOSEC$) > 10: DOSEC$ = Mid$(DOSEC$, 2): Wend
                If Len(DOSEC$) < 8 Then DOSEC$ = "Rt." + DOSEC$
                Call MOVISTOK(FF%, CI%, LOTE$, "BR", DOPRILA$, DOPRI$, DOSELA$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC1%, DEPX%, CANTI)
              End If
              '
              If RECHA# >= 0.05 Then
                CANTI = RECHA# * (-1)
                REFE$ = "RECHAZO - Rto." + TRIM$(REMI$)
                Call MOVISTOK(FF%, CI%, LOTE$, CODRECHA$, REMIRECH$, REMIRECH$, DOSELA$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC1%, DEPX%, CANTI)
                HAYRECHA% = 1
              End If
              '
              If Abs(FALTA#) >= 0.05 Then
                CANTI = FALTA# * (-1)
                REFE$ = "FALTANTE - Rto." + TRIM$(REMI$)
                If CANTI > 0 Then
                  REFE$ = "SOBRANTE - Rto." + TRIM$(REMI$)
                End If
                Call MOVISTOK(FF%, CI%, LOTE$, "BR", DOPRILA$, DOPRI$, DOSELA$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC1%, DEPX%, CANTI)
              End If
              '
              ' aqui re-activa la Orden de Compra por el rechazado mas el faltante
              RECHFAL = RECHA# + FALTA#
              If RECHFAL > 0.005 Then
                NOCORI& = CVS(Mid$(X$, 151, 4))
                If NOCORI& > 0 Then
                  Call ABRECOMPRAS
                  ABUS$ = "Nume_Ocom = " & NOCORI& & " "
                  ABUS$ = ABUS$ + "AND Cod_Inte = " & CI% & " "
                  'ABUS$ = ABUS$ + "AND (Cant_Rec+0.005) >= " & Abs(CANOM#)
                  With OcomDeta
                    .FindFirst ABUS$
1135                If .NoMatch = False Then
                      CFXUNI# = !Coef_Unids
                      If CFXUNI# < 0.005 Then CFXUNI# = 1
                      RECHAX# = RECHFAL / CFXUNI#
                      If !Cant_Rec >= RECHAX# - 0.0005 Then
                          .Edit
                          !Cant_Rec = !Cant_Rec - RECHAX#
                          !Stat_Ocom = 0
                          .Update
                          GoTo 1145
                        Else
                          .FindNext ABUS$
                          GoTo 1135
                      End If
                    End If
                    '
                    Call MENSERR(24, "No fue Posible Actualizar la Cantidad Pendiente\en la Orden de Compra Número " + TRIM$(Str$(NOCORI&)))
                    '
1145              End With
                End If
              End If
              ' ACA MARCA LA LISTAMAT CON EL Nº DE BR.
              ' ESTO ES PARA MARCAR EN LA LISTA DE MATERIALES LOS RECIBIDOS
              Call ABREORTRA
              CODEX$ = CODEXT$(CI%)
              CANAREP = APRO#
              If TRIM$(CODEX$) <> "" Then
                  SQLX$ = "CodExte = '" & CODEX$ & "' "
                  SQLX$ = SQLX$ + "AND (Nro_Ocom <> '' AND ISNULL(Nro_Ocom) = FALSE AND ISEMPTY(Nro_Ocom) = FALSE) "
                  SQLX$ = SQLX$ + "AND (Nro_BolRe = '' OR ISNULL(Nro_BolRe) = TRUE OR ISEMPTY(Nro_BolRe) = TRUE) "
                  With ListaMat
                    .FindFirst SQLX$
11271                  If .NoMatch = False Then
                       'FORMATEO EL NUMERO DE BR
                       N1$ = CStr(NROBOL)
                       Do While Len(N1$) < 6
                         N1$ = "0" + N1$
                       Loop
                       NBR2$ = "BR-" + N1$ + "-00"
                       '
                       If !CANTID <= CANAREP + 0.001 Then
                        .Edit
                        !nro_bolre = NBR2$
                        .Update
                        CANAREP = CANAREP - !CANTID
                      End If
                      .FindNext SQLX$
                      GoTo 11271
                    End If
                  End With
              End If
            End If
          End With
1128    Next U&
        'IMPRIMIR STICKER
         ' ESTO ES PARA IMPRIMIR LOS MATERIALES RECIBIDOS
   STIKAP$ = TRIM$(UCase$(CONTROL$("", "STIKAPRO")))
   If Left$(STIKAP$, 4) <> "STIK" Then
     If Left$(STIKAP$, 5) <> "STICK" Then
       GoTo 199
     End If
   End If
   '
                  Call ABREORTRA
                  SQLX$ = " SELECT * FROM LISTAMAT "
                  SQLX$ = SQLX$ + "WHERE Nro_BOLRE = '" & NBR2$ & "' "
                  Set MATRECTMP1 = OTrabajo.OpenRecordset(SQLX$, 4)
                  With MATRECTMP1
                    On Error Resume Next
                    .MoveFirst
                    If Err Then
                        On Error GoTo 0
                        Call MENSERR(24, "No se ha Podido Imprimir correctamente los STICKERS.")
                    End If
                    On Error GoTo 0
                    CRR1 = 1
                    Do While .EOF = False
                        '
                        NEV1$ = ""
                        GRU1$ = ""
                        CLI1$ = ""
                        DEV1$ = ""
                        DGR1$ = ""
                        DCR1$ = ""
                        PLA1$ = ""
                        MAT1$ = ""
                        LLA1$ = ""
                        '
                        NEV1$ = !Nro_Otrab
                        GRU1$ = !Nro_Grup
                        DCR1$ = !Descri_Cor
                        PLA1$ = !NPlano
                        MAT1$ = !Material
                        LLA1$ = !CLAVE
                        CAN1$ = FORMATNUM(!CANTID, "F8.1")
                        '
                        If NEV1$ <> "" Then
                             With OrdeTrab
                               SQLX$ = " Nro_Otrab = '" & NEV1$ & "' "
                               .FindFirst SQLX$
                               If .NoMatch = False Then
                                 DEV1$ = !Refe_Otrab
                                 NC% = !Nclie_Otrab
                                 CLI1$ = RASOCLI$(NC%)
                               End If
                             End With
                            '
                             If GRU1$ <> "" Then
                               With GrupOtra
                                 SQLX$ = " Nro_Otrab = '" & NEV1$ & "' "
                                 SQLX$ = SQLX$ + " AND Nro_Grup = '" & GRU1$ & "' "
                                 .FindFirst SQLX$
                                 If .NoMatch = False Then
                                   DGR1$ = !Deno_Grup
                                 End If
                               End With
                             End If
                            '
                            Call APROCLIB04.PRINSTIKER(NEV1$, GRU1$, CLI1$, DEV1$, DGR1$, DCR1$, PLA1$, MAT1$, LLA1$, CAN1$, CRR1)
                        End If
                        CRR1 = CRR1 + 1
                        .MoveNext
                    Loop
                  End With
        '
199     NOCLOSE% = 0
        Label12.Visible = False
        Movisto.Close
        'SalStock.Close
        Bolrecep.Close
        '
        Call CIERRARCH("*.*")
        Call BAJALDISCO
        '
        '
        LU$ = LUDAT$
        If EXISTE%(LU$ + "BRAP-ARC.FRM") > 0 Then
           HOII% = HOY(HOS$)
           FECHDOC$ = HOS$
           NUMEDOC$ = ""
           TIPODOC$ = "Vale Aprobación Recepción"
           DOCUORIG$ = "BR." + TRIM$(Str$(NROBOL&))
           '
           'CODPARAM$(1) ="TRANSPOR": DOCPARAM$(1) = TRIM$(FOREMIFAC.Text8.TEXT)
           CAPARDOC% = 0
           '
           CODTABU1$(1) = "COD-MAT"
           CODTABU1$(2) = "DES-MAT"
           CODTABU1$(3) = "CANTIDAD"
           CACOLTAB1% = 3
           '
           DESTIDOC% = 0
           CAITAB1% = 0
           '
           For YX& = 1 To ULTREG&("!BOLREDET")
              YY$ = REGLEIDO$("!BOLREDET", YX&)
              CI% = CVI(Mid$(YY$, 83, 2))
              CAN# = CVD(Mid$(YY$, 163, 8))
              If CI% > 0 Then
                If CI% <= NUMAS% Then
                  If CAN# > 0.005 Then
                    '
                    'UNID$ = UNIMED$(CI%)
                    'If UNID$ = "" Then UNID$ = UNIMED$(CI%)
                    CANS$ = CSTRING$(MKD$(CAN#), 3, 12, 1, 2)
                    '
                    CAITAB1% = CAITAB1% + 1
                    CODDX$ = CODEXT$(CI%)
                    TETABU1$(1, CAITAB1%) = CODDX$
                    TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
                    TETABU1$(3, CAITAB1%) = CANS$
                  End If
                End If
              End If
           Next YX&
           '
           Call PRINTDOC("@BRAP-ARC")   ' Remito de Facturacion
           '
        End If
        '
        Call GENIBOLPE
        '
        If HAYRECHA% = 1 Then
           Call REMIRECHA(NC1%, DEPX%, REMIRECH$)
        End If
        '
        Call CIERRARCH("*.*")
        Call BAJALDISCO
        If EXISTE%("CALPRO02.EXE") > 0 Then
          ICALREC.Label5 = TRIM$(Str$(NROBOL&))
          ICALREC.Show 1
        End If
        GoTo 1120
        '
End Sub

Sub APROCLI(NC5%, FF5%, NROBOL5&, REMI5$, DP15%)
    'APROBAR RECEPCIONES DE ORDEN DE COMPRA TEXTO LIBRE
    Call ABRESTOCKS
    With Bolrecep
         SQLX$ = " NumeBoRe = " & NROBOL5& & " "
         ' CON ESTE CRITERIO DE BUSQUEDA ALCANZA Y SOBRE
         ' YA QUE NO HAY RECEPCIONES PARCIALES.
         .FindFirst SQLX$
         If .NoMatch = False Then
               CANREN5# = !CantNom
               NOC3& = !NumOC
               DP2% = !DepoEnt
           Else
               Call COMUNI("Imposible Continuar. Número de B/R no Válido")
               Screen.MousePointer = 1
               Exit Sub
         End If
    End With
    '
    '
    Call ABRECOMPRAS
    With OcomDeta
            SQLX$ = " NUME_OCOM = " & NOC3& & ""
            'ALCANZA TAMBIEN POR Q LA O/C SOLO PUEDE
            'TENER UN ITEM.
            .FindFirst SQLX$
            If .NoMatch = False Then
                On Error Resume Next
                IMPCONCO% = !ICONTAB
                'ACA TOMO 'EV' y 'Grupo'
                NEV$ = !Nro_Otrab
                NGR$ = !Nro_Grup
                DESCRICOR$ = !Descritem
                DESCRILAR$ = !Delibre
                On Error GoTo 0
                '
            Else
               Call COMUNI("Imposible Continuar. Número de O/C no Válido")
               Screen.MousePointer = 1
               Exit Sub
            End If
        End With
    'MUESTRO LOS DATOS DE LA RECEPCION PROVISORIA
    With APROCLIB04
       .Label4 = NOC3&
       .Text1 = DESCRICOR$
       .Text2 = DESCRILAR$
       .Label5 = NC5%
       .Label11 = RASOCLI(NC5% * (-1))
       .Label7 = FORMATNUM$(CANREN5#, "F8.1")
       .Label2 = FECHATEX(FF5%)
       .Label9 = NROBOL5&
       .Label22 = REMI5$
       .Label17 = NEV$
       .Label19 = NGR$
       .Label13 = DP2%
       .Label14 = IMPCONCO%
       .Label16 = ECOARCH("TADEPOSI", Chr(DP2%))
       .Show 1
    End With
    '
End Sub

Sub BOLERECEP()
Dim IMPCONOC%
'
'falta generar ficha de recepcion fervi-air
     '
     If ULTREG&("TATIREC") < 1 Then
         X$ = REGBLAN$("TATIREC")
         Call REPLA(X$, Chr$(1), 1, 1)
         Call REPLA(X$, "Material General", 2, 31)
         Call REPLA(X$, "BOLREGEN", 49, 16)
         Call REPLA(X$, "BOLREGEN", 65, 16)
         Call REGAPP("TATIREC", X$)
         Call CIERRARCH("TATIREC")
     End If
     '
1100 Call CIERRARCH("*.*")
     '
     NROBOL& = 1
     Call ABRESTOCKS
     Set BolreTemp = Stocks.OpenRecordset("SELECT DISTINCT NumeBoRe from BOLRECEP ORDER BY NumeBoRe DESC")
     On Error Resume Next
     BolreTemp.MoveFirst
     If Err < 1 Then
       NROBOL& = 1 + BolreTemp!NumeBore
     End If
     On Error GoTo 0
     '
     If TIMATEI% = 0 Then
         X$ = REGLEIDO$("TATIREC", 1)
         TIMATEI% = Asc(Left$(X$, 1))
     End If
     '
     If TRIM$(Mid$(EBOLAN$, 4, 12)) <> "" Then
       NC% = 0
     End If
     '
     VDEF$ = MKI$(NC%) + Chr$(TIMATEI%) + CODICLI$((-1) * NC%)
     EBOL$ = OBJPLA$("SELPROVRECACC", VDEF$)
     If EBOL$ = "" Then
        Exit Sub
     End If
     EBOLAN$ = EBOL$
     '
     NC1% = CVI(Left$(EBOL$, 2))
     If NC1% > 0 Then
       NC% = NC1%
       GoTo 15
     End If
     '
     NC% = NUMEPRO%(TRIM$(Mid$(EBOL$, 4, 12)))
15   If NC% < 1 Or STATCLI%((-1) * NC%) < 1 Then
       Call MENSERR(24, "Proveedor Inexistente")
       GoTo 1100
     End If
     '
     TIMATEI% = Asc(Mid$(EBOL$, 3, 1))
     If TIMATEI% < 1 Or ECOARCH("TATIREC", Chr$(TIMATEI%)) = "" Then
         Call MENSERR(24, "Falta Tipo de Material")
         GoTo 1100
     End If
     '
     REGT& = TIMATEI%
     X$ = REGLEIDO$("TATIREC", REGT&)
     FORMPLA$ = TRIM$(UCase$(Mid$(X$, 33, 16)))
     FORMTAB$ = TRIM$(UCase$(Mid$(X$, 49, 16)))
     FORIPRE$ = TRIM$(UCase$(Mid$(X$, 65, 16)))
     '
     NC1% = (-1) * NC%
     '
1105 Y$ = PEDY$
     If Y$ = "" Then
       Y$ = Space$(128)
       Call REPLA(Y$, Chr$(0), 121, 1)
     End If
     '
     Call REPLA(Y$, MKS$(NROBOL&), 1, 4)
     Call REPLA(Y$, EBOL$, 7, 2)
     Call REPLA(Y$, MKI$(HOY(HO$)), 5, 2)
     Call REPLA(Y$, RASOCLI$(NC1%), 9, 32)
     Call REPLA(Y$, DOMICLI$(NC1%), 41, 32)
     CLOCA$ = TRIM$(CPOSCLI$(NC1%)): If CLOCA$ <> "" Then CLOCA$ = CLOCA$ + " "
     Call REPLA(Y$, CLOCA$ + LOCACLI$(NC1%), 73, 32)
     Call REPLA(Y$, Chr$(0), 121, 1)
    '
     OBJE$ = "BOLREC-ENCB"
     If CONTROL$("", "NUMEBREC") = "MANUAL" Then
         OBJE$ = "BOLREC-ENCB-M"
     End If
     '
     TICOMPRO$ = TRIM$(CONTROL$("", "COMPBREC"))
     '
     PEDY$ = OBJPLA$(OBJE$, Y$)
     If PEDY$ = "" Then Exit Sub
     '
     FF% = CVI(Mid$(PEDY$, 5, 2))
     FEX = FF%: FECHAREC$ = FECHATEX$(FEX)
     If TRIM$(FECHAREC$) = "" Then
         Call MENSERR(24, "Fecha Incorrecta")
         GoTo 1105
     End If
     '
     REMI$ = TRIM$(Mid$(PEDY$, 105, 16))
     If REMI$ = "" Then
         Call MENSERR(24, "Falta Número de Remito del Proveedor")
         GoTo 1105
     End If
     '
     DP1% = Asc(Mid$(PEDY$, 121, 1))
     DEPOSAL$ = TRIM$(ECOARCH$("ECODEP", Chr$(DP1%)))
     If DP1% < 1 Or DEPOSAL$ = "" Then
         Call MENSERR(24, "Falta Depósito de Entrada")
         GoTo 1105
     End If
     '
     NROPRO& = NROBOL&
     NROBOL& = CVS(Left$(PEDY$, 4))
     If NROBOL& < 1 Then
         Call MENSERR(24, "Número Incorrecto de B.R.")
         NROBOL& = NROPRO&
         Call REPLA(PEDY$, MKS$(NROBOL&), 1, 4)
         GoTo 1105
     End If
     '
1112 Call CIERRARCH("*.*")
     JJ& = 0
     '
     'ACA NUEVO
     RESP4% = COMALTER("Elija el Material Recibido:\\Codificado\Texto Libre")
     If RESP4% < 1 Or RESP4% = 3 Then
         Screen.MousePointer = 1
         GoTo 1100
     End If
     '
     If RESP4% = 1 Then
         GoTo 45
     Else
         Call RECEOCLI(NC%, FF%, NROBOL&, REMI$, DP1%)
         Screen.MousePointer = 1
         GoTo 1100
     End If
     'HASTA ACA NUEVO
     '
45  DARX$ = TRIM$(CONTROL$("BOLRECEP", "DIANREC"))
     If DARX$ = "" Then
          DIANRE& = 10000
        Else
          DIANRE& = Val(DARX$)
     End If
     HOIX1 = HOY(HOS$)
     '
     anuca% = 0
     If CONTROL$("BOLRECEP", "ANUCANTI") = "SI" Then
       anuca% = 1
     End If
     '
     Screen.MousePointer = 11
     Call ABRECOMPRAS
     SQLX$ = "Select * from OCOMDETA"
     SQLX$ = SQLX$ + " Where Stat_Ocom < 8 "
     SQLX$ = SQLX$ + " AND  Stat_Ocom <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
     SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
     SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
     SQLX$ = SQLX$ + " AND Cod_Inte > 0 " 'EXCLUYE LAS DE 0/C's TEXTO LIBRE
     Set COMPTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With COMPTEMP
       '
       On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Screen.MousePointer = 1
         HAYOC% = 0
         GoTo 1777
       End If
       On Error GoTo 0
       '
       JJ& = 0
        
       Do While (.EOF = False)
         CAPEN = !Cant_Ocom - !Cant_Rec
         COEFUNI = !Coef_Unids
         If COEFUNI < 0.0001 Then COEFUNI = 1
         If CAPEN > 0.005 Then
           NUOC& = !NUME_OCOM
           FENTRESO% = CVINT(!Fentre_Sol)
           FESO2 = FENTRESO%
           If DIENTRE%(HOIX1, FESO2) <= DIANRE& Then
             CI% = !COD_INTE
             LOTECO$ = !Doc_Orig
             Y$ = REGBLAN$("!BOLREDET")
             Call REPLA(Y$, MKI$(CI%), 83, 2)
             Call REPLA(Y$, UNIMED$(CI%), 85, 4)
             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 89, 8)
             If anuca% = 1 Then
               Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 163, 8)
                Call REPLA(Y$, MKD$(0), 89, 8)
             End If
             '
             Call REPLA(Y$, MKI$(!NuOrd_Item), 98, 2)
             '
             Call REPLA(Y$, LOTECO$, 107, 12)
             Call REPLA(Y$, MKS$(NUOC&), 151, 4)
             Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
             Call REPLA(Y$, MKS$(!N_OrServ), 189, 4)
             HAYOC% = 1
             JJ& = JJ& + 1
             Call GRAREG("!BOLREDET", Y$, JJ&)
           End If
         End If
        .MoveNext
       Loop
     End With
     '
     '/////****NUEVO****/////
     'TENGO QUE FIJARME SI TIENE ALGUN MATERIAL EL PROV
     'CONSIGNADO EN MACONSIG CON EL MISMO Nº DE PROV
1777 Call ABRESTOCKS
     SQLX$ = "Select * from MACONSIG"
     SQLX$ = SQLX$ + " WHERE Cant_Rendid < Cant_Entreg - 0.001 "
     SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
     Set consigtmp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     With consigtmp
       '
       On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Screen.MousePointer = 1
         GoTo 1129
       End If
       On Error GoTo 0
       '
       Y$ = REGBLAN$("!BOLREDET")
       Do While .EOF = False
         CI% = !COD_INTE
         NUOC& = Val(Mid$(!Nro_Ocom, 4, 6))
         CAPEN = !Cant_Entreg - !Cant_Rendid
         Call REPLA(Y$, MKI$(CI%), 83, 2)
         Call REPLA(Y$, UNIMED$(CI%), 85, 4)
         Call REPLA(Y$, MKD$(0), 89, 8)
         Call REPLA(Y$, "MP " + TRIM$(CSTRING$(MKS$(CAPEN), 3, 12, 1, 2)) + " " + UNIMED$(CI%), 107, 12)
         Call REPLA(Y$, MKS$(NUOC&), 151, 4)
         Call REPLA(Y$, MKS$((-1) * I&), 189, 4)
         JJ& = JJ& + 1
         Call GRAREG("!BOLREDET", Y$, JJ&)
       .MoveNext
       Loop
     End With
     '
1129 If HAYOC% = 0 Then
       Call COMUNI("No hay Ordenes de Compra Pendientes\para este Proveedor.")
     End If
     '
     Call SETULTREG("!BOLREDET", JJ&)
     Call CIERRARCH("!BOLREDET")
     '
1113 AVENTA$ = "BOLREDET"
     If anuca% = 1 Then AVENTA$ = "BOLRENUE"
     '
     VTB% = VENTABU%(AVENTA$ + "/NC/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(RASOCLI$(NC1%)) + " - " + FECHAREC$)
     If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
       X$ = "": Y$ = "": YY$ = ""
       EBOL$ = "": PEDY$ = "": DESCRI$ = ""
       GoTo 1100
     End If
     '
     For YX& = 1 To ULTREG&("!BOLREDET")
       YY$ = REGLEIDO$("!BOLREDET", YX&)
       CI% = CVI(Mid$(YY$, 83, 2))
       CAN# = CVD(Mid$(YY$, 89, 8))
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           If CAN# > 0.005 Then
             GoTo 1114
           End If
         End If
       End If
     Next YX&
     Call MENSERR(24, "Recepción Cancelada !!!\Faltan Cantidades o Códigos.")
     GoTo 1113
     '
1114 Call SETULTREG("!OBSERVOF", 0)
     Call CIERRARCH("!OBSERVOF")
     '
     VTB% = VENTABU%("OBSERVOF/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(RASOCLI$(NC1%)) + " - " + FECHAREC$)
     If VTB% < 0 Then
       If COMALTER%("Realmente Desea Cancelar ?\\Continuar\Cancelar") = 2 Then
         X$ = "": Y$ = "": YY$ = ""
         EBOL$ = "": PEDY$ = "": DESCRI$ = ""
         GoTo 1100
       End If
     End If
     '
1141 CONTA% = ULTREG&("!BOLREDET")
     '
1115 If CONTA% > 0 Then
       '
       For YX& = 1 To ULTREG&("!BOLREDET")
         YY$ = REGLEIDO$("!BOLREDET", YX&)
         CI% = CVI(Mid$(YY$, 83, 2))
         CAN# = CVD(Mid$(YY$, 89, 8))
         If CI% < 1 Then
           Call MENSERR(24, "Código no Válido.")
           GoTo 1113
         End If
         If anuca% < 1 Then
           If CAN# < 0.005 Then
             Call MENSERR(24, "Falta Cantidad.")
             GoTo 1113
           End If
         End If
       Next YX&
       '
1125   Call ABRESTOCKS
       Set BolreTemp = Stocks.OpenRecordset("SELECT DISTINCT NumeBoRe from BOLRECEP ORDER BY NumeBoRe DESC")
       On Error Resume Next
       BolreTemp.MoveFirst
       If Err < 1 Then
         If NROBOL& <= BolreTemp!NumeBore Then
           NROBOL& = 1 + BolreTemp!NumeBore
         End If
       End If
       On Error GoTo 0
       BolreTemp.Close
       '
       Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
       Call ABRESTOCKS
       Call ABRECOMPRAS
       '
       AINDI% = 0
       NORVENTA = ""
       For YX& = 1 To ULTREG&("!BOLREDET")
         YY$ = REGLEIDO$("!BOLREDET", YX&)
         CI% = CVI(Mid$(YY$, 83, 2))
         CAN# = CVD(Mid$(YY$, 89, 8))
         NUOC& = Abs(CVS(Mid$(YY$, 151, 4)))
         NUMOIT% = CVI(Mid$(YY$, 98, 2))   'EL Nº DE ORDEN DEL ITEM
         If CI% > 0 Then
           If CAN# > 0.005 Then
             CARECI = CAN#
             LOTECO$ = TRIM$(Mid$(YY$, 107, 12))
             '
             If CVD(Mid$(YY$, 163, 8)) >= 0.05 Then   ' solo si es comprado - no M.P. devuelto
               With OcomDeta
                 .FindFirst ("Nume_Ocom = " & NUOC& & " AND NuOrd_Item = " & NUMOIT%)
                 If .NoMatch = False Then
                   .Edit
                   If IsNull(!ICONTAB) Then
                        IMPCONOC% = 0
                      Else
                        IMPCONOC% = !ICONTAB 'ASIGNO LA IMPUTACION CONTABLE
                   End If
                     COEFUNI = !Coef_Unids
                     If COEFUNI < 0.0001 Then COEFUNI = 1
                   !Cant_Rec = !Cant_Rec + CARECI / COEFUNI
                     CAPENDIEN = !Cant_Ocom - !Cant_Rec
                     NUORSE& = !N_OrServ
                   If CAPENDIEN >= 0.005 Then
                       If COMALTER%("Esta O/C por Código " + TRIM$(CODEXT$(CI%)) + "\Queda con un Saldo Pendiente de " + TRIM$(FORMATNUM$(CAPENDIEN, "F10.1")) + " " + TRIM$(UNIMED$(CI%)) + ".\\Mantener Abierta\Cancelar Pendiente") = 2 Then
                         !Stat_Ocom = 8
                       End If
                     Else
                       !Stat_Ocom = 3
                   End If
                   .Update
                 End If
               End With
             End If
             '
             ' aca arma un string con las ordenes de compra recibidas
             NORVE& = Abs(CVS(Mid$(YY$, 151, 4)))
             If NORVE& > 0 Then
               If NORVE& <> NORVEA& Then
                 NORVEA& = NORVE&
                 If NORVENTA <> "" Then NORVENTA = NORVENTA + " - "
                 NORVENTA = NORVENTA + TRIM$(Str$(NORVE&))
               End If
             End If
             '
             DESCRI$ = TRIM$(CODEXT$(CI%)) + " - " + TRIM$(DESCRIT$(CI%))
             If AINDI% = 0 Then
                X$ = REGBLAN$("INDIBOL")
                Call REPLA(X$, MKS$(NROBOL&), 1, 4)
                FEX = FF%
                Call REPLA(X$, FECHATEX$(FEX), 5, 8)
                Call REPLA(X$, REMI$, 15, 16)
                Call REPLA(X$, ECOARCH$("PROVED1", MKI$(NC%)), 32, 29)
                Call REGAPP("INDIBOL", X$)
                If NUORSE& < 1 Then Call REGAPP("IBOLPEN", X$)
                AINDI% = 1
             End If
             '
             NBR$ = TRIM$(Str$(NROBOL&))
             NORDE% = NORDE% + 1
             NORDES$ = TRIM$(Str$(NORDE%))
             IDENLO$ = "BR-" + NBR$ + "-" + NORDES$
             '
             Call REPLA(YY$, MKS$(NROBOL&), 1, 4)
             If TIMATEI% <> 2 Then Call REPLA(YY$, DESCRI$, 5, 64)
             '
             If TRIM$(Mid$(YY$, 69, 12)) = "" Then
               Call REPLA(YY$, IDENLO$, 69, 12)
             End If
             Call REPLA(YY$, MKI$(FF%), 81, 2)
             DEPX% = DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(CI%)
             Call REPLA(YY$, Chr$(DEPX%), 97, 1)
             Call REPLA(YY$, MKI$(NC%), 105, 2)
             Call REPLA(YY$, REMI$, 119, 16)
             If CVD(Mid$(YY$, 163, 8)) >= 0.05 Then   ' solo si es comprado - no M.P. devuelto
               Call REPLA(YY$, MKS$(NUORSE&), 189, 4)
             End If
             Call GRAREG("!BOLREDET", YY$, YX&)
             '
''///**NUEVO ACA EDITO EL REGISTRO EN MACONSIG***///
'Call ABRESTOCKS
'With MaConsig
'  .FindFirst (" VAL(MID$(NRO_Ocom,4,6)) = " & NUOC& & " AND Cod_Inte = " & CI% & " ")
'  If .NoMatch = False Then
'    .Edit
'      !Cant_Rendid = !Cant_Rendid + CARECI / COEFUNI
'        CAPENDIEN = !Cant_Entreg - !Cant_Rendid
'      !Saldo_Pend = CAPENDIEN
'    .Update
'  End If
'End With
''////***HASTA ACA NUEVO****////
'
'
             With Bolrecep
              .AddNew
               !NumeBore = NROBOL&
               !Descrip = Left$(DESCRI$, 64)
               !IdenLote = Left$(IDENLO$, 12)
               !FechRecep = CVDAT(FF%)
               !COD_INTE = CI%
               !Cod_Exte = Left$(CODEXT$(CI%), 16)
               !DescripItem = Left$(DESCRIT0$(CI%), 64)
               !Unidad = Left$(UNIMED$(CI%), 4)
               !CantNom = CAN#
               !DepoEnt = DEPX%
               !NUMPROV = NC%
               !RaSo_Prov = Left$(RASOCLI$((-1) * NC%), 64)
                  LOTECO$ = TRIM$(LOTECO$)
                  While Len(LOTECO$) > 12
                    LOTECO$ = Mid$(LOTECO$, 2)
                  Wend
               !LoteCol_Prov = Left$(LOTECO$, 12)
               !NumRem_Prov = Left$(REMI$, 16)
               !NumFac_Prov = ""
               !NumOC = NUOC&
               !ICONTAB = IMPCONOC% 'Imputación contable de la OC.
               !MoneVaUni = ""
               !Valo_Unit = 0
               !CanToApro = 0
               !CanToRech = 0
               !CanToFalta = 0
               !FechLibe = CVDAT(0)
               !NumeOrdSer = NUORSE&
               !Observa = " " 'OBSE$
               !STATUS = 0
               !USUAPRO = " "
               !FeHoApro = CVDAT(0)
              .Update
             End With
             '
             If NUORSE& < 1 Then
               Call DESCACONSI(NC%, CI%, CAN#)
             End If
             '
             If NUORSE& > 0 Then
               If CAN# > 0.05 Then
                 GoSub BLANTEROP
               End If
             End If
             '
           End If
         End If
       Next YX&
       '
       Bolrecep.Close
       OcomEnca.Close
       OcomDeta.Close
'          For YX& = 1 To ULTREG&("!OBSERVOF")
'             YY$ = REGLEIDO$("!OBSERVOF", YX&)
'             YY$ = MKS$(NROBOL&) + YY$
'             Call REGAPP("OBSEBREC", YY$)
'          Next YX&
          '
'          Call CIERRARCH("BOLRECEP")
'          Call CIERRARCH("BOLREPEN")
'          Call CIERRARCH("INDIBOL")
'          Call CIERRARCH("IBOLPEN")
'          Call CIERRARCH("OBSEBREC")
'          '
'          Call FRESHECHO("IBOLPEN")
'          Screen.MousePointer = 1
'          '
       If TRIM$(FORIPRE$) <> "" Then
         '
         CANCELPRINT% = 0
         If CONTROL("BOLRECEP", "SUPRINT") = "SI" Then CANCELPRINT% = 1
         '
         TIPODOC$ = "Boleta de Recepción"
         DESTIDOC% = (-1) * Abs(NC1%)
         NUMEDOC$ = Str$(NROBOL&)
         FEX = FF%: FECHDOC$ = FECHATEX$(FEX)
         '
         CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = DEPOSAL$
         CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = REMI$
         CODPARAM$(3) = "N-AUTENT": DOCPARAM$(3) = NORVENTA
         CAPARDOC% = 3
         '
         CODTABU1$(1) = "COD-MAT"
         CODTABU1$(2) = "DES-MAT"
         CODTABU1$(3) = "COLOR"
         CODTABU1$(4) = "UNIMED"
         CODTABU1$(5) = "NRO-LOTE"
         CODTABU1$(6) = "CANTIDAD"
         CODTABU1$(7) = "N-OCOMPR"
         CODTABU1$(8) = "CAKILOS"
         CODTABU1$(9) = "F-PLANO"
         CODTABU1$(10) = "FECH-REC"
         CACOLTAB1% = 10
         '
         CAITAB1% = 0
         For YX& = 1 To ULTREG&("!BOLREDET")
           YY$ = REGLEIDO$("!BOLREDET", YX&)
           CI% = CVI(Mid$(YY$, 83, 2))
           CAN# = CVD(Mid$(YY$, 89, 8))
           If CI% > 0 Then
             If CI% <= NUMAS% Then
               If CAN# > 0.005 Then
                 '
                 UNID$ = TRIM$(Mid$(YY$, 85, 4))
                 If UNID$ = "" Then UNID$ = UNIMED$(CI%)
                 CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
                 KILAJE$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 10, 3, 2)
                 FEX = CVI(Mid$(YY$, 155, 2))
                 FESOLEN$ = FECHATEX$(FEX)
                 FEPLA$ = FEPLANO$(CI%)
                 NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
                 If NULOTE$ = "" Then
                    NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
                 End If
                 NUOCO$ = FORMATNUM$(Abs(CVS(Mid$(YY$, 151, 4))), "F7,0")
                 '
                 CAITAB1% = CAITAB1% + 1
                 TETABU1$(1, CAITAB1%) = TRIM$(CODEXT$(CI%))
                 '
                 If FORMPLA$ = "RECEP-ESPUMA" Then
                      DESMAT$ = Mid$(YY$, 5, 64)
                    Else
                      DESMAT$ = DESCRIT$(CI%)
                 End If
                 TETABU1$(2, CAITAB1%) = DESMAT$
                 TETABU1$(3, CAITAB1%) = COLOR$(CI%)
                 TETABU1$(4, CAITAB1%) = UNID$
                 TETABU1$(5, CAITAB1%) = NULOTE$
                 TETABU1$(6, CAITAB1%) = CANS$
                 TETABU1$(7, CAITAB1%) = NUOCO$
                 TETABU1$(8, CAITAB1%) = KILAJE$
                 TETABU1$(9, CAITAB1%) = FEPLA$
                 TETABU1$(10, CAITAB1%) = FESOLEN$
                 '
                 If TICOMPRO$ = "FICHA" Then
                   Call PRINTDOC("FORMBREC")
                 End If
                 '
               End If
             End If
           End If
         Next YX&
         '
         If TICOMPRO$ <> "FICHA" Then
           Call PRINTDOC("FORMBREC")
         End If
         '
       End If
          '
1199 End If
     '
11999 Call CIERRARCH("*.*")
     '
     X$ = "": Y$ = "": YY$ = ""
     EBOL$ = "": PEDY$ = "": DESCRI$ = ""
     GoTo 1100
     '
BLANTEROP:
   '
   FINTER& = ULTREG&("OPERTER")
   For U& = FINTER& To 1 Step -1
     X0$ = REGLEIDO$("OPERTER", U&)
     If CVS(Mid$(X0$, 75, 4)) >= NUORSE& Then
       REGITER& = U&
       GoTo 21
     End If
   Next U&
   GoTo 29
   '
21 ' Marca como "Cumplida" Operación en OPERTER.DAT
   SALRE# = CVD(Mid$(X0$, 79, 8)) - CAN#
   If SALRE# < 0 Then SALRE# = 0
   '
   FEBA% = HOY(HOS$)
   HOBA% = HORANUM%(Time$)
   USUA% = USNBR% Mod 100
   DBAJA$ = MKI$(FEBA%) + MKI$(HOBA%) + Chr$(USUA%) + Chr$(1) + "B.Rec." + TRIM$(Str$(NROBOL&))
   '
   NUOROP% = Asc(Mid$(X0$, 67, 1))
   Call REPLA(X0$, MKD$(CAN#), 87, 8)
   Call REPLA(X0$, MKD$(SALRE#), 95, 8)
   Call REPLA(X0$, "3ROS  ", 151, 6)
   Call REPLA(X0$, MKI$(FEBA%), 157, 2)
   Call REPLA(X0$, MKI$(HORANUM%(Time$)), 159, 2)
   Call REPLA(X0$, MKS$(0), 161, 4)
   Call REPLA(X0$, String$(8, 0), 165, 8)
   Call REPLA(X0$, MKS$(0), 173, 4)
   Call REPLA(X0$, DBAJA$, 177, 22)
   Call GRAREG("OPERTER", X0$, REGITER&)
   '
' Si se recibio un parcial, genera registro en OPERTER.DAT
' por el saldo pendiente
   If SALRE# > 0.01 Then
      NUOROP% = 1 + Asc(Mid$(X0$, 67, 1))
      CIIK% = CVI(Mid$(X0$, 13, 2))
      NOPE% = CVI(Mid$(X0$, 15, 2))
      DALTA$ = Chr$(NUOROP%) + Mid$(DBAJA$, 2, 5)
      INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
      PROGBLA$ = Space$(6) + String$(26, 0)
      MINVARI = TIESTAN(CIIK%, NOPE%, SALRE#, NUOROP%)
      '
      Call REPLA(X0$, DALTA$, 67, 6)
      Call REPLA(X0$, MKD$(SALRE#), 79, 8)
      Call REPLA(X0$, MKD$(0), 87, 8)
      Call REPLA(X0$, MKD$(SALRE#), 95, 8)
      Call REPLA(X0$, String$(32, 0), 103, 32)
      Call REPLA(X0$, MKS$(MINVARI), 111, 4)
      Call REPLA(X0$, MKS$(MINVARI), 131, 4)
      Call REPLA(X0$, String$(16, 0), 135, 16)
      Call REPLA(X0$, INFOBLA$, 151, 48)
      Call REPLA(X0$, PROGBLA$, 225, 32)
      Call REGAPP("OPERTER", X0$)
      AGREGA% = 1
   End If
   '
29 Return
   '
End Sub

Private Sub RECEOCLI(NC2%, FR%, NBR&, NREMI$, DP2%)
   'MATERIAL TEXTO LIBRE
   DARX$ = TRIM$(CONTROL$("BOLRECEP", "DIANREC"))
     If DARX$ = "" Then
          DIANRE& = 10000
        Else
          DIANRE& = Val(DARX$)
     End If
     HOIX1 = HOY(HOS$)
     '
     anuca% = 0
     If CONTROL$("BOLRECEP", "ANUCANTI") = "SI" Then
       anuca% = 1
     End If
     '
     Call COPYSTRU("INDIOCOM", "INDIRETE")
     '
     Screen.MousePointer = 11
     Call ABRECOMPRAS
     SQLX$ = "Select * from OCOMDETA"
     SQLX$ = SQLX$ + " Where Stat_Ocom < 8 "
     SQLX$ = SQLX$ + " AND  Stat_Ocom <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
     SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
     SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC2% & " "
     SQLX$ = SQLX$ + " AND Cod_Inte = 0 " ' 0/C's TEXTO LIBRE
     Set COMPTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With COMPTEMP
       '
       On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Screen.MousePointer = 1
         HAYOC% = 0
         GoTo 1117
       End If
       On Error GoTo 0
       '
       Y$ = REGBLAN("INDIRETE")
       JJ& = 0
       Do While (.EOF = False)
             Call REPLA(Y$, MKS(!NUME_OCOM), 1, 4)
             Call REPLA(Y$, !Descritem, 5, 44)
             HAYOC% = 1
             JJ& = JJ& + 1
             Call GRAREG("!INDIRETE", Y$, JJ&)
        .MoveNext
       Loop
     End With
     '
     Call SETULTREG("!INDIRETE", JJ&)
     Call FRESHECHO("!INDIRETE")
1117 If HAYOC% = 0 Then
       Call COMUNI("No hay Compras Libres Pendientes\para este Proveedor.")
       Screen.MousePointer = 1
       Exit Sub
     End If
     '
     Call SELECHO("!INDIRETE")
     VDEVU$ = VALACT1$("!INDIRETE")
     If VDEVU$ <> "" Then
        With OcomDeta
            NOC3& = Val(VDEVU$)
            SQLX$ = " NUME_OCOM = " & NOC3& & ""
            .FindFirst SQLX$
            If .NoMatch = False Then
                NROC& = !NUME_OCOM
                DECRICOR$ = TRIM$(!Descritem)
                DECRILAR$ = !Delibre
                NC% = !Npro_Ocom
                CANREN# = !Cant_Ocom - !Cant_Rec
                IMPCONCO% = !ICONTAB
                'ACA TOMO 'EV' y 'Grupo'
                NEV$ = !Nro_Otrab
                NGR$ = !Nro_Grup
                '
            Else
               Call COMUNI("Imposible Continuar. Número de O/C no Válido")
               Screen.MousePointer = 1
               Exit Sub
            End If
        End With
     Else
        Screen.MousePointer = 1
        Exit Sub
     End If
     '
1129 'LEVANTA LA VENTANA CON O/C
     anuca% = 0
     If CONTROL$("BOLRECEP", "ANUCANTI") = "SI" Then
       anuca% = 1
     End If
     '
     With PRESOCLIB04
       .Label4 = NROC&
       .Text1 = DECRICOR$
       .Text2 = DECRILAR$
       .Label5 = NC%
       .Label11 = RASOCLI(NC% * (-1))
       .Label7 = FORMATNUM$(CANREN#, "F8.1")
       .Text3 = .Label7: If anuca% > 0 Then Text3 = ""
       .Label2 = FECHATEX(FR%)
       .Label9 = NBR&
       .Label22 = NREMI$
       .Label17 = NEV$
       .Label19 = NGR$
       .Label13 = DP2%
       .Label14 = IMPCONCO%
       .Label16 = ECOARCH("TADEPOSI", Chr(DP2%))
       .Show 1
     End With
   '
   'HASTA ACA NUEVO MAT TEXTO LIBRE
   '
End Sub
'RUTINA ANTIGUA DE RECEPCION
'Sub BOLERECEP()
'Dim IMPCONOC%
''
''falta generar ficha de recepcion fervi-air
'     '
'     If ULTREG&("TATIREC") < 1 Then
'         X$ = REGBLAN$("TATIREC")
'         Call REPLA(X$, Chr$(1), 1, 1)
'         Call REPLA(X$, "Material General", 2, 31)
'         Call REPLA(X$, "BOLREGEN", 49, 16)
'         Call REPLA(X$, "BOLREGEN", 65, 16)
'         Call REGAPP("TATIREC", X$)
'         Call CIERRARCH("TATIREC")
'     End If
'     '
'1100 Call CIERRARCH("*.*")
'     '
'     NROBOL& = 1
'     Call ABRESTOCKS
'     Set BolreTemp = Stocks.OpenRecordset("SELECT DISTINCT NumeBoRe from BOLRECEP ORDER BY NumeBoRe DESC")
'     On Error Resume Next
'     BolreTemp.MoveFirst
'     If Err < 1 Then
'       NROBOL& = 1 + BolreTemp!NumeBore
'     End If
'     On Error GoTo 0
'     '
'     If TIMATEI% = 0 Then
'         X$ = REGLEIDO$("TATIREC", 1)
'         TIMATEI% = Asc(Left$(X$, 1))
'     End If
'     '
'     VDEF$ = MKI$(NC%) + Chr$(TIMATEI%) + CODICLI$((-1) * NC%)
'     EBOL$ = OBJPLA$("SELPROVRECACC", VDEF$)
'     If EBOL$ = "" Then
'        Exit Sub
'     End If
'     '
'     NC% = NUMEPRO%(TRIM$(Mid$(EBOL$, 4, 12)))
'     If NC% < 1 Or STATCLI%((-1) * NC%) < 1 Then
'       Call MENSERR(24, "Proveedor Inexistente")
'       GoTo 1100
'     End If
'     '
'     TIMATEI% = Asc(Mid$(EBOL$, 3, 1))
'     If TIMATEI% < 1 Or ECOARCH("TATIREC", Chr$(TIMATEI%)) = "" Then
'         Call MENSERR(24, "Falta Tipo de Material")
'         GoTo 1100
'     End If
'     '
'     REGT& = TIMATEI%
'     X$ = REGLEIDO$("TATIREC", REGT&)
'     FORMPLA$ = TRIM$(UCase$(Mid$(X$, 33, 16)))
'     FORMTAB$ = TRIM$(UCase$(Mid$(X$, 49, 16)))
'     FORIPRE$ = TRIM$(UCase$(Mid$(X$, 65, 16)))
'     '
'     NC1% = (-1) * NC%
'     '
'1105 Y$ = PEDY$
'     If Y$ = "" Then
'       Y$ = Space$(128)
'       Call REPLA(Y$, Chr$(0), 121, 1)
'     End If
'     '
'     Call REPLA(Y$, MKS$(NROBOL&), 1, 4)
'     Call REPLA(Y$, EBOL$, 7, 2)
'     Call REPLA(Y$, MKI$(HOY(HO$)), 5, 2)
'     Call REPLA(Y$, RASOCLI$(NC1%), 9, 32)
'     Call REPLA(Y$, DOMICLI$(NC1%), 41, 32)
'     CLOCA$ = TRIM$(CPOSCLI$(NC1%)): If CLOCA$ <> "" Then CLOCA$ = CLOCA$ + " "
'     Call REPLA(Y$, CLOCA$ + LOCACLI$(NC1%), 73, 32)
'     Call REPLA(Y$, Chr$(0), 121, 1)
'    '
'     OBJE$ = "BOLREC-ENCB"
'     If CONTROL$("", "NUMEBREC") = "MANUAL" Then
'         OBJE$ = "BOLREC-ENCB-M"
'     End If
'     '
'     TICOMPRO$ = TRIM$(CONTROL$("", "COMPBREC"))
'     '
'     PEDY$ = OBJPLA$(OBJE$, Y$)
'     If PEDY$ = "" Then Exit Sub
'     '
'     FF% = CVI(Mid$(PEDY$, 5, 2))
'     FEX = FF%: FECHAREC$ = FECHATEX$(FEX)
'     If TRIM$(FECHAREC$) = "" Then
'         Call MENSERR(24, "Fecha Incorrecta")
'         GoTo 1105
'     End If
'     '
'     REMI$ = TRIM$(Mid$(PEDY$, 105, 16))
'     If REMI$ = "" Then
'         Call MENSERR(24, "Falta Número de Remito del Proveedor")
'         GoTo 1105
'     End If
'     '
'     DP1% = Asc(Mid$(PEDY$, 121, 1))
'     DEPOSAL$ = TRIM$(ECOARCH$("ECODEP", Chr$(DP1%)))
'     If DP1% < 1 Or DEPOSAL$ = "" Then
'         Call MENSERR(24, "Falta Depósito de Entrada")
'         GoTo 1105
'     End If
'     '
'     NROPRO& = NROBOL&
'     NROBOL& = CVS(Left$(PEDY$, 4))
'     If NROBOL& < 1 Then
'         Call MENSERR(24, "Número Incorrecto de B.R.")
'         NROBOL& = NROPRO&
'         Call REPLA(PEDY$, MKS$(NROBOL&), 1, 4)
'         GoTo 1105
'     End If
'     '
'1112 Call CIERRARCH("*.*")
'     JJ& = 0
'     '
'     DARX$ = TRIM$(CONTROL$("BOLRECEP", "DIANREC"))
'     If DARX$ = "" Then
'          DIANRE& = 10000
'        Else
'          DIANRE& = Val(DARX$)
'     End If
'     HOIX1 = HOY(HOS$)
'     '
'     ANUCA% = 0
'     If CONTROL$("BOLRECEP", "ANUCANTI") = "SI" Then
'       ANUCA% = 1
'     End If
'     '
'     Screen.MousePointer = 11
'     Call ABRECOMPRAS
'     SQLX$ = "Select * from OCOMDETA"
'     SQLX$ = SQLX$ + " Where Stat_Ocom < 8 "
'     SQLX$ = SQLX$ + " AND  Stat_Ocom <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
'     SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
'     SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
'     Set COMPTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
'     '
'     With COMPTEMP
'       '
'       On Error Resume Next
'       .MoveFirst
'       If Err Then
'         On Error GoTo 0
'         Screen.MousePointer = 1
'         HAYOC% = 0
'         GoTo 11129
'       End If
'       On Error GoTo 0
'       '
'       JJ& = 0
'
'       Do While (.EOF = False)
'         CAPEN = !Cant_Ocom - !Cant_Rec
'         COEFUNI = !Coef_Unids
'         If COEFUNI < 0.0001 Then COEFUNI = 1
'         If CAPEN > 0.005 Then
'           NUOC& = !Nume_Ocom
'           FENTRESO% = CVINT(!Fentre_Sol)
'           FESO2 = FENTRESO%
'           If DIENTRE%(HOIX1, FESO2) <= DIANRE& Then
'             CI% = !Cod_Inte
'             LOTECO$ = !Doc_Orig
'             Y$ = REGBLAN$("!BOLREDET")
'             Call REPLA(Y$, MKI$(CI%), 83, 2)
'             Call REPLA(Y$, UNIMED$(CI%), 85, 4)
'             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 89, 8)
'             If ANUCA% = 1 Then
'               Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 163, 8)
'                Call REPLA(Y$, MKD$(0), 89, 8)
'             End If
'             '
'             Call REPLA(Y$, MKI$(!NuOrd_Item), 98, 2)
'             '
'             Call REPLA(Y$, LOTECO$, 107, 12)
'             Call REPLA(Y$, MKS$(NUOC&), 151, 4)
'             Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
'             Call REPLA(Y$, MKS$(REDET&), 189, 4)
'             HAYOC% = 1
'             JJ& = JJ& + 1
'             Call GRAREG("!BOLREDET", Y$, JJ&)
'           End If
'         End If
'        .MoveNext
'       Loop
'     End With
'     '/////****NUEVO****/////
'     'TENGO QUE FIJARME SI TIENE ALGUN MATERIAL EL PROV
'     'CONSIGNADO EN MACONSIG CON EL MISMO Nº DE PROV
'      Call ABRESTOCKS
'      SQLX$ = "Select * from MACONSIG"
'      SQLX$ = SQLX$ + " WHERE Cant_Rendid < Cant_Entreg - 0.001 "
'      SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
'      Set CONSIGTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
'      With CONSIGTMP
'        '
'        On Error Resume Next
'        .MoveFirst
'        If Err Then
'          On Error GoTo 0
'          Screen.MousePointer = 1
'          GoTo 11129
'        End If
'        On Error GoTo 0
'        '
'        Y$ = REGBLAN$("!BOLREDET")
'        Do While .EOF = False
'          CI% = !Cod_Inte
'          NUOC& = Val(Mid$(!Nro_Ocom, 4, 6))
'          CAPEN = !Cant_Entreg - !Cant_Rendid
'          Call REPLA(Y$, MKI$(CI%), 83, 2)
'          Call REPLA(Y$, UNIMED$(CI%), 85, 4)
'          Call REPLA(Y$, MKD$(0), 89, 8)
'          Call REPLA(Y$, "MP " + TRIM$(CSTRING$(MKS$(CAPEN), 3, 12, 1, 2)) + " " + UNIMED$(CI%), 107, 12)
'          Call REPLA(Y$, MKS$(NUOC&), 151, 4)
'          Call REPLA(Y$, MKS$((-1) * I&), 189, 4)
'          JJ& = JJ& + 1
'          Call GRAREG("!BOLREDET", Y$, JJ&)
'        .MoveNext
'        Loop
'      End With
'         'HASTA ACA LLEGA OK
'
''        UREOCO& = ULTREG&("OCOMDETA")
''        HAYOC% = 0
''        RFF$ = RECFILT$("OCOMDETA", 3, MKI$(NC%))
''        For UI& = 1 To Len(RFF$) Step 4
''          REDET& = CVS(Mid$(RFF$, UI&, 4))
''          If REDET& > 0 Then
''            If REDET& <= UREOCO& Then
''              X$ = REGLEIDO$("OCOMDETA", REDET&)
''              If CVI(Mid$(X$, 7, 2)) = NC% Then
''                If Asc(Mid$(X$, 124, 1)) < 8 Then ' abierta Y NO CUMPLIDA
''                  CAPEN = CVS(Mid$(X$, 55, 4)) - CVS(Mid$(X$, 99, 4))
''                  COEFUNI = CVS(Mid$(X$, 63, 4))
''                  If COEFUNI < 0.0001 Then COEFUNI = 1
''                  If CAPEN > 0 Then
''                    NUOC& = CVS(Left$(X$, 4))
''                    FENTRESO% = CVI(Mid$(X$, 95, 2))
''                    FESO2 = FENTRESO%
''                    If DIENTRE%(HOIX1, FESO2) <= DIANRE& Then
''                      CI% = CVI(Mid$(X$, 9, 2))
''                      LOTECO$ = Mid$(X$, 105, 12)
''                      Y$ = REGBLAN$("!BOLREDET")
''                      Call REPLA(Y$, MKI$(CI%), 83, 2)
''                      Call REPLA(Y$, UNIMED$(CI%), 85, 4)
''                      Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 89, 8)
''                      If ANUCA% = 1 Then
''                        Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 163, 8)
''                        Call REPLA(Y$, MKD$(0), 89, 8)
''                      End If
''                      Call REPLA(Y$, LOTECO$, 107, 12)
''                      Call REPLA(Y$, MKS$(NUOC&), 151, 4)
''                      Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
''                      Call REPLA(Y$, MKS$(REDET&), 189, 4)
''                      HAYOC% = 1
''                      JJ& = JJ& + 1
''                      Call GRAREG("!BOLREDET", Y$, JJ&)
''                    End If
''                  End If
''                End If
''              End If
''            End If
''          End If
''        Next UI&
'        '
'11129     If HAYOC% = 0 Then
'       Call COMUNI("No hay Ordenes de Compra Pendientes\para este Proveedor.")
'     End If
'     '
'' QUEDA PENDIENTE LA DEVOLUCION DE CONSIGNACIONES
''        For I& = 1 To ULTREG&("MACONSIG")
''          X$ = REGLEIDO$("MACONSIG", I&)
''          If CVI(Mid$(X$, 3, 2)) = NC% Then
''            CAPEN = CVD(Mid$(X$, 63, 8)) - CVD(Mid$(X$, 111, 8))
''            If CAPEN > 0 Then
''              NUOC& = Val(Mid$(X$, 24, 6))
''              CI% = CVI(Mid$(X$, 33, 2))
''              Y$ = REGBLAN$("!BOLREDET")
''              Call REPLA(Y$, MKI$(CI%), 83, 2)
''              Call REPLA(Y$, UNIMED$(CI%), 85, 4)
''              Call REPLA(Y$, MKD$(0), 89, 8)
''              Call REPLA(Y$, "MP " + TRIM$(CSTRING$(MKS$(CAPEN), 3, 12, 1, 2)) + " " + UNIMED$(CI%), 107, 12)
''              Call REPLA(Y$, MKS$(NUOC&), 151, 4)
''              Call REPLA(Y$, MKS$((-1) * I&), 189, 4)
''              For UI& = 1 To JJ&
''                If Left$(Y$, 154) = Left$(REGLEIDO$("!BOLREDET", UI&), 154) Then GoTo 11129 ' SUPRIME SI COINCIDE
''              Next UI&
''              JJ& = JJ& + 1
''              Call GRAREG("!BOLREDET", Y$, JJ&)
''            End If
''          End If
''11129   Next I&
'
'     Call SETULTREG("!BOLREDET", JJ&)
'     Call CIERRARCH("!BOLREDET")
'     '
'1113 AVENTA$ = "BOLREDET"
'     If ANUCA% = 1 Then AVENTA$ = "BOLRENUE"
'     '
'     VTB% = VENTABU%(AVENTA$ + "/NC/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(RASOCLI$(NC1%)) + " - " + FECHAREC$)
'     If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
'       X$ = "": Y$ = "": YY$ = ""
'       EBOL$ = "": PEDY$ = "": DESCRI$ = ""
'       GoTo 1100
'     End If
'     '
'     For YX& = 1 To ULTREG&("!BOLREDET")
'       YY$ = REGLEIDO$("!BOLREDET", YX&)
'       CI% = CVI(Mid$(YY$, 83, 2))
'       CAN# = CVD(Mid$(YY$, 89, 8))
'       If CI% > 0 Then
'         If CI% <= NUMAS% Then
'           If CAN# > 0.005 Then
'             GoTo 1114
'           End If
'         End If
'       End If
'     Next YX&
'     Call MENSERR(24, "Recepción Cancelada !!!\Faltan Cantidades o Códigos.")
'     GoTo 1113
'     '
'1114 Call SETULTREG("!OBSERVOF", 0)
'     Call CIERRARCH("!OBSERVOF")
'     '
'     VTB% = VENTABU%("OBSERVOF/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(RASOCLI$(NC1%)) + " - " + FECHAREC$)
'     If VTB% < 0 Then
'       If COMALTER%("Realmente Desea Cancelar ?\\Continuar\Cancelar") = 2 Then
'         X$ = "": Y$ = "": YY$ = ""
'         EBOL$ = "": PEDY$ = "": DESCRI$ = ""
'         GoTo 1100
'       End If
'     End If
'     '
'11141 CONTA% = ULTREG&("!BOLREDET")
'     '
'1115 If CONTA% > 0 Then
'       '
'       For YX& = 1 To ULTREG&("!BOLREDET")
'         YY$ = REGLEIDO$("!BOLREDET", YX&)
'         CI% = CVI(Mid$(YY$, 83, 2))
'         CAN# = CVD(Mid$(YY$, 89, 8))
'         If CI% < 1 Then
'           Call MENSERR(24, "Código no Válido.")
'           GoTo 1113
'         End If
'         If ANUCA% < 1 Then
'           If CAN# < 0.005 Then
'             Call MENSERR(24, "Falta Cantidad.")
'             GoTo 1113
'           End If
'         End If
'       Next YX&
'       '
'1125   Call ABRESTOCKS
'       Set BolreTemp = Stocks.OpenRecordset("SELECT DISTINCT NumeBoRe from BOLRECEP ORDER BY NumeBoRe DESC")
'       On Error Resume Next
'       BolreTemp.MoveFirst
'       If Err < 1 Then
'         If NROBOL& <= BolreTemp!NumeBore Then
'           NROBOL& = 1 + BolreTemp!NumeBore
'         End If
'       End If
'       On Error GoTo 0
'       BolreTemp.Close
'       '
'       Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
'       Call ABRESTOCKS
'       Call ABRECOMPRAS
'       '
'       AINDI% = 0
'       NORVENTA = ""
'       For YX& = 1 To ULTREG&("!BOLREDET")
'         YY$ = REGLEIDO$("!BOLREDET", YX&)
'         CI% = CVI(Mid$(YY$, 83, 2))
'         CAN# = CVD(Mid$(YY$, 89, 8))
'         NUOC& = CVS(Mid$(YY$, 151, 4))
'         NUMOIT% = CVI(Mid$(YY$, 98, 2))   'EL Nº DE ORDEN DEL ITEM
'         If CI% > 0 Then
'           If CAN# > 0.005 Then
'             CARECI = CAN#
'                '
'
'                  'If REGOC& > 0 Then
'                   ' If REGOC& <= ULTREG&("OCOMDETA") Then
'                     ' ZZ$ = REGLEIDO$("OCOMDETA", REGOC&)
'             COEFUNI = CVS(Mid$(ZZ$, 63, 4))
'             If COEFUNI < 0.0001 Then COEFUNI = 1
'             LOTECO$ = TRIM$(Mid$(YY$, 107, 12))
'             '
'             With OcomDeta
'               .FindFirst ("Nume_Ocom = " & NUOC& & " AND NuOrd_Item = " & NUMOIT%)
'               If .NoMatch = False Then
'                 .Edit
'                 If IsNull(!ICONTAB) Then
'                   IMPCONOC% = 0
'                 Else
'                    IMPCONOC% = !ICONTAB 'ASIGNO LA IMPUTACION CONTABLE
'                 End If
'                 !Cant_Rec = !Cant_Rec + CARECI / COEFUNI
'                 CAPENDIEN = !Cant_Ocom - !Cant_Rec
'                   NUORSE& = !N_OrServ
'                 If CAPENDIEN >= 0.005 Then
'                     If COMALTER%("Esta O/C por Código " + TRIM$(CODEXT$(CI%)) + "\Queda con un Saldo Pendiente de " + TRIM$(FORMATNUM$(CAPENDIEN, "F10.1")) + " " + TRIM$(UNIMED$(CI%)) + ".\\Mantener Abierta\Cancelar Pendiente") = 2 Then
'                       !Stat_Ocom = 8
'                     End If
'                   Else
'                     !Stat_Ocom = 3
'                 End If
'                 .Update
'               End If
'             End With
'             '///**NUEVO ACA EDITO EL REGISTRO EN MACONSIG***///
'             Call ABRESTOCKS
'             With MaConsig
'               .FindFirst (" VAL(MID$(NRO_Ocom,4,6)) = " & NUOC& & " AND Cod_Inte = " & CI% & " ")
'               If .NoMatch = False Then
'                 .Edit
'                   !Cant_Rendid = Cant_Rendid + CARECI / COEFUNI
'                     CAPENDIEN = !Cant_Entreg - !Cant_Rendid
'                   !Saldo_Pend = CAPENDIEN
'                 .Update
'               End If
'             End With
'            '////***HASTA ACA NUEVO****////
'
'             '
'             ' aca arma un string con las ordenes de compra recibidas
'             NORVE& = CVS(Mid$(YY$, 151, 4))
'             If NORVE& > 0 Then
'               If NORVE& <> NORVEA& Then
'                 NORVEA& = NORVE&
'                 If NORVENTA <> "" Then NORVENTA = NORVENTA + " - "
'                 NORVENTA = NORVENTA + TRIM$(Str$(NORVE&))
'               End If
'             End If
'             '
'             DESCRI$ = TRIM$(CODEXT$(CI%)) + " - " + TRIM$(DESCRIT$(CI%))
'                  If AINDI% = 0 Then
'                     X$ = REGBLAN$("INDIBOL")
'                     Call REPLA(X$, MKS$(NROBOL&), 1, 4)
'                     FEX = FF%
'                     Call REPLA(X$, FECHATEX$(FEX), 5, 8)
'                     Call REPLA(X$, REMI$, 15, 16)
'                     Call REPLA(X$, ECOARCH$("PROVED1", MKI$(NC%)), 32, 29)
'                     Call REGAPP("INDIBOL", X$)
'                     If NUORSE& < 1 Then Call REGAPP("IBOLPEN", X$)
'                     AINDI% = 1
'                  End If
'                  '
'             NBR$ = TRIM$(Str$(NROBOL&))
'             NORDE% = NORDE% + 1
'             NORDES$ = TRIM$(Str$(NORDE%))
'             IDENLO$ = "BR-" + NBR$ + "-" + NORDES$
'             '
'             Call REPLA(YY$, MKS$(NROBOL&), 1, 4)
'             If TIMATEI% <> 2 Then Call REPLA(YY$, DESCRI$, 5, 64)
'             '
'             If TRIM$(Mid$(YY$, 69, 12)) = "" Then
'               Call REPLA(YY$, IDENLO$, 69, 12)
'             End If
'             Call REPLA(YY$, MKI$(FF%), 81, 2)
'             DEPX% = DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(CI%)
'             Call REPLA(YY$, Chr$(DEPX%), 97, 1)
'             Call REPLA(YY$, MKI$(NC%), 105, 2)
'             Call REPLA(YY$, REMI$, 119, 16)
'             Call REPLA(YY$, MKS$(NUORSE&), 189, 4)
'             Call GRAREG("!BOLREDET", YY$, YX&)
'             '
'             With Bolrecep
'              .AddNew
'               !NumeBore = NROBOL&
'               !Descrip = Left$(DESCRI$, 64)
'               !IdenLote = Left$(IDENLO$, 12)
'               !FechRecep = CVDAT(FF%)
'               !Cod_Inte = CI%
'               !Cod_Exte = Left$(CODEXT$(CI%), 16)
'               !DescripItem = Left$(DESCRIT0$(CI%), 64)
'               !Unidad = Left$(UNIMED$(CI%), 4)
'               !CantNom = CAN#
'               !DepoEnt = DEPX%
'               !NUMPROV = NC%
'               !RaSo_Prov = Left$(RASOCLI$((-1) * NC%), 64)
'                  LOTECO$ = TRIM$(LOTECO$)
'                  While Len(LOTECO$) > 12
'                    LOTECO$ = Mid$(LOTECO$, 2)
'                  Wend
'               !LoteCol_Prov = Left$(LOTECO$, 12)
'               !NumRem_Prov = Left$(REMI$, 16)
'               !NumFac_Prov = ""
'               !NumOC = NUOC&
'               !ICONTAB = IMPCONOC% 'Imputación contable de la OC.
'               !MoneVaUni = ""
'               !Valo_Unit = 0
'               !CanToApro = 0
'               !CanToRech = 0
'               !CanToFalta = 0
'               !FechLibe = CVDAT(0)
'               !NumeOrdSer = NUORSE&
'               !Observa = " " 'OBSE$
'               !STATUS = 0
'               !USUAPRO = " "
'               !FeHoApro = CVDAT(0)
'              .Update
'             End With
'             '
'             If NUORSE& < 1 Then
'               If REGOC& > 0 Then
'                   Call DESCACONSI(NC%, CI%, CAN#)
'                 ElseIf REGOC& < 0 Then
'                   Call DEVOLCONSI(Abs(REGOC&), NC%, CI%, CAN#)
'               End If
'             End If
'             '
'             If NUORSE& > 0 Then
'               If CAN# > 0.05 Then
'                 GoSub BLANTEROP
'               End If
'             End If
'             '
'           End If
'         End If
'       Next YX&
'       '
'       Bolrecep.Close
'       OcomEnca.Close
'       OcomDeta.Close
''          For YX& = 1 To ULTREG&("!OBSERVOF")
''             YY$ = REGLEIDO$("!OBSERVOF", YX&)
''             YY$ = MKS$(NROBOL&) + YY$
''             Call REGAPP("OBSEBREC", YY$)
''          Next YX&
'          '
''          Call CIERRARCH("BOLRECEP")
''          Call CIERRARCH("BOLREPEN")
''          Call CIERRARCH("INDIBOL")
''          Call CIERRARCH("IBOLPEN")
''          Call CIERRARCH("OBSEBREC")
''          '
''          Call FRESHECHO("IBOLPEN")
''          Screen.MousePointer = 1
''          '
'       If TRIM$(FORIPRE$) <> "" Then
'         '
'         CANCELPRINT% = 0
'         If CONTROL("BOLRECEP", "SUPRINT") = "SI" Then CANCELPRINT% = 1
'         '
'         TIPODOC$ = "Boleta de Recepción"
'         DESTIDOC% = NC1%
'         NUMEDOC$ = Str$(NROBOL&)
'         FEX = FF%: FECHDOC$ = FECHATEX$(FEX)
'         '
'         CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = DEPOSAL$
'         CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = REMI$
'         CODPARAM$(3) = "N-AUTENT": DOCPARAM$(3) = NORVENTA
'         CAPARDOC% = 3
'         '
'         CODTABU1$(1) = "COD-MAT"
'         CODTABU1$(2) = "DES-MAT"
'         CODTABU1$(3) = "COLOR"
'         CODTABU1$(4) = "UNIMED"
'         CODTABU1$(5) = "NRO-LOTE"
'         CODTABU1$(6) = "CANTIDAD"
'         CODTABU1$(7) = "N-OCOMPR"
'         CODTABU1$(8) = "CAKILOS"
'         CODTABU1$(9) = "F-PLANO"
'         CODTABU1$(10) = "FECH-REC"
'         CACOLTAB1% = 10
'         '
'         CAITAB1% = 0
'         For YX& = 1 To ULTREG&("!BOLREDET")
'           YY$ = REGLEIDO$("!BOLREDET", YX&)
'           CI% = CVI(Mid$(YY$, 83, 2))
'           CAN# = CVD(Mid$(YY$, 89, 8))
'           If CI% > 0 Then
'             If CI% <= NUMAS% Then
'               If CAN# > 0.005 Then
'                 '
'                 UNID$ = TRIM$(Mid$(YY$, 85, 4))
'                 If UNID$ = "" Then UNID$ = UNIMED$(CI%)
'                 CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
'                 KILAJE$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 10, 3, 2)
'                 FEX = CVI(Mid$(YY$, 155, 2))
'                 FESOLEN$ = FECHATEX$(FEX)
'                 FEPLA$ = FEPLANO$(CI%)
'                 NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
'                 If NULOTE$ = "" Then
'                    NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
'                 End If
'                 NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
'                 '
'                 CAITAB1% = CAITAB1% + 1
'                 TETABU1$(1, CAITAB1%) = TRIM$(CODEXT$(CI%))
'                 '
'                 If FORMPLA$ = "RECEP-ESPUMA" Then
'                      DESMAT$ = Mid$(YY$, 5, 64)
'                    Else
'                      DESMAT$ = DESCRIT$(CI%)
'                 End If
'                 TETABU1$(2, CAITAB1%) = DESMAT$
'                 TETABU1$(3, CAITAB1%) = COLOR$(CI%)
'                 TETABU1$(4, CAITAB1%) = UNID$
'                 TETABU1$(5, CAITAB1%) = NULOTE$
'                 TETABU1$(6, CAITAB1%) = CANS$
'                 TETABU1$(7, CAITAB1%) = NUOCO$
'                 TETABU1$(8, CAITAB1%) = KILAJE$
'                 TETABU1$(9, CAITAB1%) = FEPLA$
'                 TETABU1$(10, CAITAB1%) = FESOLEN$
'                 '
'                 If TICOMPRO$ = "FICHA" Then
'                   Call PRINTDOC("FORMBREC")
'                 End If
'                 '
'               End If
'             End If
'           End If
'         Next YX&
'         '
'         If TICOMPRO$ <> "FICHA" Then
'           Call PRINTDOC("FORMBREC")
'         End If
'         '
'       End If
'          '
'1199 End If
'     '
'11999 Call CIERRARCH("*.*")
'     '
'     X$ = "": Y$ = "": YY$ = ""
'     EBOL$ = "": PEDY$ = "": DESCRI$ = ""
'     GoTo 1100
'     '
'BLANTEROP:
'   '
'   FINTER& = ULTREG&("OPERTER")
'   For U& = FINTER& To 1 Step -1
'     X0$ = REGLEIDO$("OPERTER", U&)
'     If CVS(Mid$(X0$, 75, 4)) >= NUORSE& Then
'       REGITER& = U&
'       GoTo 21
'     End If
'   Next U&
'   GoTo 29
'   '
'21 ' Marca como "Cumplida" Operación en OPERTER.DAT
'   SALRE# = CVD(Mid$(X0$, 79, 8)) - CAN#
'   If SALRE# < 0 Then SALRE# = 0
'   '
'   FEBA% = HOY(HOS$)
'   HOBA% = HORANUM%(Time$)
'   USUA% = USNBR% Mod 100
'   DBAJA$ = MKI$(FEBA%) + MKI$(HOBA%) + Chr$(USUA%) + Chr$(1) + "B.Rec." + TRIM$(Str$(NROBOL&))
'   '
'   NUOROP% = Asc(Mid$(X0$, 67, 1))
'   Call REPLA(X0$, MKD$(CAN#), 87, 8)
'   Call REPLA(X0$, MKD$(SALRE#), 95, 8)
'   Call REPLA(X0$, "3ROS  ", 151, 6)
'   Call REPLA(X0$, MKI$(FEBA%), 157, 2)
'   Call REPLA(X0$, MKI$(HORANUM%(Time$)), 159, 2)
'   Call REPLA(X0$, MKS$(0), 161, 4)
'   Call REPLA(X0$, String$(8, 0), 165, 8)
'   Call REPLA(X0$, MKS$(0), 173, 4)
'   Call REPLA(X0$, DBAJA$, 177, 22)
'   Call GRAREG("OPERTER", X0$, REGITER&)
'   '
'' Si se recibio un parcial, genera registro en OPERTER.DAT
'' por el saldo pendiente
'   If SALRE# > 0.01 Then
'      NUOROP% = 1 + Asc(Mid$(X0$, 67, 1))
'      CIIK% = CVI(Mid$(X0$, 13, 2))
'      NOPE% = CVI(Mid$(X0$, 15, 2))
'      DALTA$ = Chr$(NUOROP%) + Mid$(DBAJA$, 2, 5)
'      INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
'      PROGBLA$ = Space$(6) + String$(26, 0)
'      MINVARI = TIESTAN(CIIK%, NOPE%, SALRE#, NUOROP%)
'      '
'      Call REPLA(X0$, DALTA$, 67, 6)
'      Call REPLA(X0$, MKD$(SALRE#), 79, 8)
'      Call REPLA(X0$, MKD$(0), 87, 8)
'      Call REPLA(X0$, MKD$(SALRE#), 95, 8)
'      Call REPLA(X0$, String$(32, 0), 103, 32)
'      Call REPLA(X0$, MKS$(MINVARI), 111, 4)
'      Call REPLA(X0$, MKS$(MINVARI), 131, 4)
'      Call REPLA(X0$, String$(16, 0), 135, 16)
'      Call REPLA(X0$, INFOBLA$, 151, 48)
'      Call REPLA(X0$, PROGBLA$, 225, 32)
'      Call REGAPP("OPERTER", X0$)
'      AGREGA% = 1
'   End If
'   '
'29 Return
'   '
'   '
'End Sub
''

Sub GENIBOLPE()
    '
    Screen.MousePointer = 11
    Call ABRESTOCKS
    '
    SQLX$ = " SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + " WHERE Status = 0 "
    SQLX$ = SQLX$ + " ORDER BY NumeBoRe "
    '
    Set COBRECEPTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With COBRECEPTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Boletas de Recepcion Pendientes.")
        On Error GoTo 0
        JJ& = 0: K& = 0
        GoTo 90
      End If
      On Error GoTo 0
      JJ& = 0
      K& = 0
      Do While (.EOF = False)
         Y$ = REGBLAN("BOLREPEN")
         Call REPLA(Y$, MKS$(!NumeBore), 1, 4)
         Call REPLA(Y$, !IdenLote, 69, 12)
         Call REPLA(Y$, MKI$(CVINT(!FechRecep)), 81, 2)
         Call REPLA(Y$, MKI$(!COD_INTE), 83, 2)
         Call REPLA(Y$, !Unidad, 85, 4)
         Call REPLA(Y$, MKD$(!CantNom), 89, 8)
         Call REPLA(Y$, Chr$(!DepoEnt), 97, 1)
         Call REPLA(Y$, MKI$(!NUMPROV), 105, 2)
         Call REPLA(Y$, !LoteCol_Prov, 107, 12)
         Call REPLA(Y$, !NumRem_Prov, 119, 16)
         Call REPLA(Y$, !NumFac_Prov, 135, 16)
         Call REPLA(Y$, MKS$(!NumOC), 151, 4)
         Call REPLA(Y$, !MoneVaUni, 155, 4)
         Call REPLA(Y$, MKS$(!Valo_Unit), 159, 4)
         Call REPLA(Y$, MKD$(!CanToApro), 163, 8)
         Call REPLA(Y$, MKD$(!CanToRech), 171, 8)
         Call REPLA(Y$, MKD$(!CanToFalta), 179, 8)
         Call REPLA(Y$, MKI$(CVINT(!FechLibe)), 187, 2)
         Call REPLA(Y$, !Observa, 193, 64)
         JJ& = JJ& + 1
         Call GRAREG("BOLREPEN", Y$, JJ&)
         '
         If !NumeBore <> ANTNUM Then
           XX$ = REGBLAN$("IBOLPEN")
         '
           Call REPLA(XX$, MKS$(!NumeBore), 1, 4)
           FEX$ = Format(!FechRecep, "DD/MM/YY")
           Call REPLA(XX$, FEX$, 5, 8)
           Call REPLA(XX$, !NumRem_Prov, 15, 16)
           Call REPLA(XX$, ECOARCH$("PROVED1", MKI$(!NUMPROV)), 32, 29)
           K& = K& + 1
           Call GRAREG("IBOLPEN", XX$, K&)
           ANTNUM = !NumeBore
        End If
        '
      .MoveNext
      Loop
    End With
    '
90  Set COBRECEPTMP = Nothing
    Call SETULTREG("BOLREPEN", JJ&)
    Call CIERRARCH("BOLREPEN")
    Call SETULTREG("IBOLPEN", K&)
    Call CIERRARCH("IBOLPEN")
    Call FRESHECHO("IBOLPEN")
    '
    Screen.MousePointer = 1
    '
End Sub

Sub REPIBOL()
        '
        Dim REGBO&(2048)
        '
        TICOMPRO$ = TRIM$(CONTROL$("", "COMPBREC"))
        
1120    'Call CLEARVEN
        X$ = "": Y$ = "": Z$ = "": YY$ = ""
        EBOL$ = "": PEDY$ = "": DESCRI$ = ""
        NBR$ = OBJPLA$("SELBOLET", "")
        If NBR$ = "" Then
            Exit Sub
        End If
        '
        NROBOL& = CVS(Left$(NBR$, 4))
        If NROBOL& < 1 Then
            Exit Sub
        End If
        '
        Screen.MousePointer = 11
        FIN& = ULTREG&("BOLRECEP")
        Call SETULTREG("!BOLREDET", 0)
        Call CIERRARCH("!BOLREDET")
        Call SETULTREG("!OBSERVOF", 0)
        Call CIERRARCH("!OBSERVOF")
        '
        For U& = 1 To ULTREG&("BOLRECEP")
            X$ = REGLEIDO$("BOLRECEP", U&)
            If CVS(Left$(X$, 4)) = NROBOL& Then
                NC% = CVI(Mid$(X$, 7, 2))
                Call REGAPP("!BOLREDET", X$)
            End If
        Next U&
        '
        For U& = 1 To ULTREG&("OBSEBREC")
            X$ = REGLEIDO$("OBSEBREC", U&)
            If CVS(Left$(X$, 4)) = NROBOL& Then
                Call REGAPP("!OBSERVOF", Mid$(X$, 5, 72))
            End If
        Next U&
        '
        Screen.MousePointer = 1
        '
        Call CIERRARCH("!BOLREDET")
        Call CIERRARCH("!OBSERVOF")
        '
        If ULTREG&("!BOLREDET") < 1 Then
            Call COMUNI("Boleta de Recepcion Nro." + Str$(NROBOL&) + "\no Figura en la Base de Datos.")
            GoTo 1120
        End If
        '
        X$ = REGLEIDO$("TATIREC", 1)
        FORMPLA$ = TRIM$(UCase$(Mid$(X$, 33, 16)))
        FORMTAB$ = TRIM$(UCase$(Mid$(X$, 49, 16)))
        FORIPRE$ = TRIM$(UCase$(Mid$(X$, 65, 16)))
        '
        If TRIM$(FORIPRE$) = "" Then
            Call COMUNI("No se Definió Formulario de Impresión de B.R.")
            GoTo 1120
        End If
        '
        YY$ = REGLEIDO$("!BOLREDET", 1)
        NC% = CVI(Mid$(YY$, 105, 2))
        NC1% = (-1) * NC%
        FF% = CVI(Mid$(YY$, 81, 2))
        DP1% = Asc(Mid$(YY$, 97, 1))
        DEPOSAL$ = TRIM$(ECOARCH$("ECODEP", Chr$(DP1%)))
        REMI$ = Mid$(YY$, 119, 16)
        '
        If TRIM$(FORIPRE$) <> "" Then
               '
               TIPODOC$ = "Boleta de Recepción"
               DESTIDOC% = NC1%
               NUMEDOC$ = Str$(NROBOL&)
               FEX = FF%: FECHDOC$ = FECHATEX$(FEX)
               '
               CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = DEPOSAL$
               CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = REMI$
               CAPARDOC% = 2
               '
               CODTABU1$(1) = "COD-MAT"
               CODTABU1$(2) = "DES-MAT"
               CODTABU1$(3) = "COLOR"
               CODTABU1$(4) = "UNIMED"
               CODTABU1$(5) = "NRO-LOTE"
               CODTABU1$(6) = "CANTIDAD"
               CODTABU1$(7) = "N-OCOMPR"
               CODTABU1$(8) = "CAKILOS"
               CODTABU1$(9) = "F-PLANO"
               CACOLTAB1% = 9
               '
               CAITAB1% = 0
               For YX& = 1 To ULTREG&("!BOLREDET")
                  YY$ = REGLEIDO$("!BOLREDET", YX&)
                  CI% = CVI(Mid$(YY$, 83, 2))
                  CAN# = CVD(Mid$(YY$, 89, 8))
                  If CI% > 0 Then
                     If CI% <= NUMAS% Then
                        If CAN# > 0.005 Then
                           '
                           UNID$ = TRIM$(Mid$(YY$, 85, 4))
                           If UNID$ = "" Then UNID$ = UNIMED$(CI%)
                           CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
                           KILAJE$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 10, 3, 2)
                           FEPLA$ = FEPLANO$(CI%)
                           NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
                           If NULOTE$ = "" Then
                              NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
                           End If
                           NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
                           '
                           CAITAB1% = CAITAB1% + 1
                           TETABU1$(1, CAITAB1%) = TRIM$(CODEXT$(CI%))
                           '
                           If FORMPLA$ = "RECEP-ESPUMA" Then
                                DESMAT$ = Mid$(YY$, 5, 64)
                              Else
                                DESMAT$ = DESCRIT$(CI%)
                           End If
                           TETABU1$(2, CAITAB1%) = DESMAT$
                           TETABU1$(3, CAITAB1%) = COLOR$(CI%)
                           TETABU1$(4, CAITAB1%) = UNID$
                           TETABU1$(5, CAITAB1%) = NULOTE$
                           TETABU1$(6, CAITAB1%) = CANS$
                           TETABU1$(7, CAITAB1%) = NUOCO$
                           TETABU1$(8, CAITAB1%) = KILAJE$
                           TETABU1$(9, CAITAB1%) = FEPLA$
                           '
                           If TICOMPRO$ = "FICHA" Then
                               Call PRINTDOC("FORMBREC")
                           End If
                           '
                        End If
                     End If
                  End If
               Next YX&
               '
               If TICOMPRO$ <> "FICHA" Then
                   Call PRINTDOC("FORMBREC")
               End If
               '
1199    End If
        '
        Call CIERRARCH("*.*")
        '
        X$ = "": Y$ = "": YY$ = ""
        EBOL$ = "": PEDY$ = "": DESCRI$ = ""
        GoTo 1120
        '
End Sub
'

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
   FORECEVOI.Hide
End Sub
'
Sub SHOWOCOM(NPX&)
    '
    TIDOCUM$ = "Nota de Pedido"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU$ = " " + NUDOCU$ + " "
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    Exit Sub
    '
5  For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    Exit Sub
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    FIN& = Len(TPSP$)
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7070
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7070
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    UI& = UI& + 1
    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
    If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.HScroll1.Min = 0
    EDITFORM.HScroll1.Value = 0
    EDITFORM.HScroll1.Max = 1
    If ANCHPAGINA > EDITFORM.Frame1.Width Then
      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command5.Enabled = False
    EDITFORM.Show 1
    '
End Sub
'
Sub SHOWBORE(NPX&)
    '
    TIDOCUM$ = "Boleta de Recepción"
    NUDOCU$ = TRIM$(Str$(NPX&))
    NUDOCU$ = " " + NUDOCU$ + " -"
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        If InStr(XX$, NUDOCU$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    Exit Sub
    '
5  For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    Exit Sub
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    FIN& = Len(TPSP$)
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7070
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7070
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    UI& = UI& + 1
    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
    If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.HScroll1.Min = 0
    EDITFORM.HScroll1.Value = 0
    EDITFORM.HScroll1.Max = 1
    If ANCHPAGINA > EDITFORM.Frame1.Width Then
      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command5.Enabled = False
    EDITFORM.Show 1
    '
End Sub
'
Sub COBORECEP()
    '
    'MARCA PARA SABER SI HAY MATERIALES DE TEXTO LIBRE
    HAYMATERIALIB% = 0
    '
    Call ABRESTOCKS
    '
    SQLX$ = " SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + " WHERE Status = 0 "
    SQLX$ = SQLX$ + " ORDER BY NumeBoRe "
    '
    Set COBRECEPTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With COBRECEPTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Boletas de Recepcion Pendientes.")
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      JJ& = 0
      K& = 0
      TC& = 0
      Do While (.EOF = False)
            Y$ = REGBLAN("BOLREPEN")
            '
            Call REPLA(Y$, MKS$(!NumeBore), 1, 4)
            Call REPLA(Y$, !IdenLote, 69, 12)
            Call REPLA(Y$, MKI$(CVINT(!FechRecep)), 81, 2)
            Call REPLA(Y$, MKI$(!COD_INTE), 83, 2)
            Call REPLA(Y$, !Unidad, 85, 4)
            Call REPLA(Y$, MKD$(!CantNom), 89, 8)
            Call REPLA(Y$, Chr$(!DepoEnt), 97, 1)
            Call REPLA(Y$, MKI$(!NUMPROV), 105, 2)
            Call REPLA(Y$, !LoteCol_Prov, 107, 12)
            Call REPLA(Y$, !NumRem_Prov, 119, 16)
            Call REPLA(Y$, !NumFac_Prov, 135, 16)
            Call REPLA(Y$, MKS$(!NumOC), 151, 4)
            Call REPLA(Y$, !MoneVaUni, 155, 4)
            Call REPLA(Y$, MKS$(!Valo_Unit), 159, 4)
            Call REPLA(Y$, MKD$(!CanToApro), 163, 8)
            Call REPLA(Y$, MKD$(!CanToRech), 171, 8)
            Call REPLA(Y$, MKD$(!CanToFalta), 179, 8)
            Call REPLA(Y$, MKI$(CVINT(!FechLibe)), 187, 2)
            Call REPLA(Y$, !Observa, 193, 64)
            '
            If !COD_INTE <= 0 Then HAYMATERIALIB% = 1 'HAY MATERIAL TEXTO LIBRE
            '
            JJ& = JJ& + 1
            If !COD_INTE > 0 Then
                 TC& = TC& + 1
                Call GRAREG("BOLREPEN", Y$, TC&)
            End If
            '
            XXZ$ = REGBLAN("BOREPELI")
            XXZ$ = Left$(Y$, 82) + Space$(174) + Mid$(Y$, 85)
            If !COD_INTE <= 0 Then
                'Call REPLA(XXZ$, MKI$(!COD_INTE), 9, 2)
                Call REPLA(XXZ$, !Cod_Exte, 83, 16)
                Call REPLA(XXZ$, !DescripItem, 99, 64)
              Else
               CIXI% = !COD_INTE
               Call REPLA(XXZ$, CODEXT$(CIXI%), 83, 16)
               Call REPLA(XXZ$, DESCRIT0$(CIXI%), 99, 64)
            End If
            '
            NUOCO1& = !NumOC
            Call ABRECOMPRAS
            With OcomDeta 'PARA OBTENER EL DESTINO Y GRUPO
                '
                SQLX$ = " NUME_OCOM = " & NUOCO1& & ""
                .FindFirst SQLX$
                If .NoMatch = False Then
                  '
                  If IsNull(!Nro_Otrab) Then
                     NOTR$ = ""
                    Else
                         NOTR$ = !Nro_Otrab
                  End If
                  If IsNull(!Nro_Grup) Then
                         NGR$ = ""
                    Else
                         NGR$ = !Nro_Grup
                  End If
                  DEST1$ = NOTR$
                  If DEST1$ <> "" Then
                    If NGR$ <> "" Then
                      DEST1$ = DEST1$ + " / "
                    End If
                  End If
                  DEST1$ = DEST1$ + NGR$
                  Call REPLA(XXZ$, DEST1$, 163, 24)
                  '
                End If
                '
            End With
            '
            Call GRAREG("BOREPELI", XXZ$, JJ&)
            '
            '
            If !NumeBore <> ANTNUM Then
              XX$ = REGBLAN$("IBOLPEN")
            '
              Call REPLA(XX$, MKS$(!NumeBore), 1, 4)
              FEX$ = Format(!FechRecep, "DD/MM/YY")
              Call REPLA(XX$, FEX$, 5, 8)
              Call REPLA(XX$, !NumRem_Prov, 15, 16)
              Call REPLA(XX$, ECOARCH$("PROVED1", MKI$(!NUMPROV)), 32, 29)
              K& = K& + 1
              Call GRAREG("IBOLPEN", XX$, K&)
              ANTNUM = !NumeBore
           End If
           '
      .MoveNext
      Loop
    End With
    Set COBRECEPTMP = Nothing
    Call SETULTREG("BOLREPEN", JJ&)
    Call CIERRARCH("BOLREPEN")
    Call SETULTREG("IBOLPEN", K&)
    Call CIERRARCH("IBOLPEN")
    
    '
End Sub

Sub Selcopen()
    '
    Screen.MousePointer = 11
    Call COPYSTRU("INDIOCOM", "INDIOCPE")
    Call ABRECOMPRAS
    SQLX$ = " SELECT  DISTINCT OCOMENCA.Nume_Ocom, OCOMENCA.Refe_Ocom "
    SQLX$ = SQLX$ + " FROM OCOMENCA "
    SQLX$ = SQLX$ + " INNER JOIN  OCOMDETA "
    SQLX$ = SQLX$ + " ON (OCOMENCA.Nume_Ocom = OCOMDETA.Nume_Ocom) "
    SQLX$ = SQLX$ + " WHERE OCOMDETA.Stat_Ocom < 8 "   ' para que no este anulada ni cumplida
    SQLX$ = SQLX$ + " AND OCOMDETA.Cant_Rec < (OCOMDETA.Cant_Ocom - 0.005) "
    SQLX$ = SQLX$ + " AND OCOMENCA.Stat_Ocom < 1 "
    SQLX$ = SQLX$ + " AND OCOMDETA.Stat_Ocom >= 0 " 'YA APROBADAS
    SQLX$ = SQLX$ + " AND OCOMENCA.Stat_Ocom >= 0 " 'YA APROBADAS
    SQLX$ = SQLX$ + " ORDER BY OCOMENCA.Nume_Ocom  ASC "
    Set OCMODIFTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    With OCMODIFTEMP
       On Error Resume Next
        .MoveFirst
        If Err Then
          Call MENSERR(24, "No Hay Ordenes de Compra\Pendientes.")
          Screen.MousePointer = 1
          On Error GoTo 0
          Exit Sub
        End If
        On Error GoTo 0
        '
        JJ& = 0
        Do While (.EOF = False)
            XXY$ = REGBLAN$("INDIOCPE")
            Call REPLA(XXY$, MKS(!NUME_OCOM), 1, 4) 'Nº O/C
            Call REPLA(XXY$, !Refe_OcoM, 5, 44) ' REFERENCIA
            JJ& = JJ& + 1
            Call GRAREG("INDIOCPE", XXY$, JJ&)
            .MoveNext
        Loop
    End With
    Call SETULTREG("INDIOCPE", JJ&)
    Call CIERRARCH("INDIOCPE")
    '
End Sub
