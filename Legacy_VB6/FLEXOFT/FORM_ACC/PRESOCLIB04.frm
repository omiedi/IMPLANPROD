VERSION 5.00
Begin VB.Form PRESOCLIB04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recepción de Material de Texto Libre"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture15 
      Height          =   6670
      Left            =   7665
      ScaleHeight     =   6615
      ScaleWidth      =   1035
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   6700
         Left            =   0
         Picture         =   "PRESOCLIB04.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1080
      Left            =   4900
      TabIndex        =   31
      Top             =   4545
      Width           =   2535
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C.Vigil:"
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
         Left            =   -1200
         TabIndex        =   39
         Top             =   705
         Width           =   1920
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
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
         Left            =   860
         TabIndex        =   38
         Top             =   630
         Width           =   1485
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
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
         Left            =   860
         TabIndex        =   33
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   330
         Left            =   -1200
         TabIndex        =   32
         Top             =   310
         Width           =   1920
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1080
      Left            =   210
      TabIndex        =   26
      Top             =   4545
      Width           =   4575
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
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
         Left            =   2310
         TabIndex        =   40
         Top             =   630
         Width           =   2010
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Certificado de Calidad:"
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
         Index           =   2
         Left            =   105
         TabIndex        =   41
         Top             =   690
         Width           =   2130
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo:"
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
         Left            =   1440
         TabIndex        =   30
         Top             =   310
         Width           =   1920
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
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
         Left            =   3450
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "E.V.:"
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
         Left            =   -1290
         TabIndex        =   28
         Top             =   310
         Width           =   1920
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
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
         Left            =   720
         TabIndex        =   27
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   6945
      Left            =   7665
      ScaleHeight     =   6885
      ScaleWidth      =   1005
      TabIndex        =   16
      Top             =   -120
      Width           =   1065
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
         Height          =   650
         Left            =   105
         Picture         =   "PRESOCLIB04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   315
         Width           =   650
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
         Height          =   650
         Left            =   120
         Picture         =   "PRESOCLIB04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Aprueba la Selección de Items a Comprar"
         Top             =   5040
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   17
         Top             =   5880
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   18
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -240
         Picture         =   "PRESOCLIB04.frx":29AC
         Stretch         =   -1  'True
         Top             =   6195
         Width           =   1410
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   4080
      TabIndex        =   13
      Top             =   5700
      Width           =   3360
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
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
         Left            =   1005
         TabIndex        =   25
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Deposito Destino:"
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
         Left            =   -420
         TabIndex        =   24
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   1395
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   210
      TabIndex        =   12
      Top             =   5700
      Width           =   3750
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
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
         Left            =   2625
         TabIndex        =   35
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rec:"
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
         Index           =   1
         Left            =   630
         TabIndex        =   34
         Top             =   420
         Width           =   1920
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pend:"
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
         Index           =   0
         Left            =   -1155
         TabIndex        =   15
         Top             =   420
         Width           =   1920
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
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
         Left            =   840
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   210
      TabIndex        =   4
      Top             =   1260
      Width           =   7215
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
         Height          =   2415
         Left            =   1125
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   630
         Width           =   5910
      End
      Begin VB.TextBox Text1 
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
         Height          =   285
         Left            =   1140
         TabIndex        =   5
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Material:"
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
         Left            =   105
         TabIndex        =   11
         Top             =   320
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1110
      Left            =   210
      TabIndex        =   0
      Top             =   105
      Width           =   7215
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Item:"
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
         Left            =   2205
         TabIndex        =   37
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
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
         Left            =   3225
         TabIndex        =   36
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   4320
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   3480
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   2760
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Left            =   1140
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº O/C:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Left            =   5955
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   330
         Left            =   4920
         TabIndex        =   7
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
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
         Left            =   1140
         TabIndex        =   3
         Top             =   615
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
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
         Left            =   120
         TabIndex        =   2
         Top             =   690
         Width           =   960
      End
      Begin VB.Label Label11 
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
         Left            =   2295
         TabIndex        =   1
         Top             =   615
         Width           =   4755
      End
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu aprobar 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "PRESOCLIB04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aprobar_Click()
   Call Command2_Click
End Sub

Private Sub BOTEXIT_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    '
    'ALTA EN BOLRECEP
    If Val(Text3) < 0.05 Then
      Call MENSERR(24, "Falta Cantidad Recibida.")
      GoTo 99
    End If
    '
    CERCAPRO$ = TRIM$(Text4)
    If Control$("BOLRECEP", "CERCAPRO") = "SI" Then
       If CERCAPRO$ = "" Then
          If COMALTER%("Falta Certificado de Calidad del Proveedor.\\Corregir\Continuar") <> 2 Then
             Command2.Enabled = True
             Exit Sub
          End If
       End If
    End If
    '
    Call ABRESTOCKS
    With Bolrecep
      '
      NC% = Val(Label5)
      NUOC& = Val(Label4)
      CARECI# = Val(Text3)
      NBoRE& = Val(Label9)
      REMI$ = TRIM$(Label22)
      DEPX% = Val(Label13)
      NORIT& = Val(Label12)
      NUCONVI$ = Left$(TRIM$(Label24), 16)
      
      .AddNew
       !NUMEBORE = NBoRE&
       !Descrip = Left$(Text1, 64)
       !IdenLote = " "
       !FechRecep = CVDAT(CVINT(Label2))
       !Cod_Inte = 0
       !Cod_Exte = " "
       !DescripItem = Left$(Text1, 64) 'Left$(DESCRIT0$(0), 64)
       !Unidad = "PZA"
       !CantNom = CARECI#
       !DepoENT = DEPX%
       !NUMPROV = NC%
       !RaSo_Prov = Left$(RASOCLI$((-1) * NC%), 64)
          LOTECO$ = " "
          While Len(LOTECO$) > 12
            LOTECO$ = Mid$(LOTECO$, 2)
          Wend
       !LoteCol_Prov = Left$(LOTECO$, 12)
       !NUMREM_PROV = Left$(REMI$, 16)
       !CerCal_Prov = CERCAPRO$
       !NumFac_Prov = ""
       !NumConVig = NUCONVI$
       !NumOC = NUOC&
       !ICONTAB = Val(Label14) 'Imputación contable de la OC.
       !MoneVaUni = ""
       !Valo_Unit = 0
       !CanToApro = CARECI#
       !CanToRech = 0
       !CanToFalta = 0
       !FechLibe = CVDAT(CVINT(Label2))
       !NumeOrdSer = 0
       !Observa = USERNAME
       !Status = 1
       !UsuApro = USERNAME
       !FeHoApro = Now
      .Update
      '
    End With
    '
    'MARCO EN OCOMDETA LA CANT. RECIBIDA
    Call ABRECOMPRAS
    With OcomDeta
      SQLX$ = " NUME_OCOM = " & NUOC&
      SQLX$ = SQLX$ & " AND NuOrd_Item = " & NORIT&
      SQLX$ = SQLX$ + " AND Stat_Ocom < 8 "         ' NO ANULADAS
      SQLX$ = SQLX$ + " AND  Stat_Ocom <> - 1 "     ' EXCLUYE NO APROB.
      '
      .FindFirst SQLX$
      If .NoMatch = False Then
           .Edit
           !Cant_rec = !Cant_rec + CARECI#
           CAPENDIEN = !Cant_Ocom - !Cant_rec
           If CAPENDIEN < 0.05 Then
             !Stat_Ocom = 3
           End If
           .Update
         Else
           Call MENSERR(24, "No Fue Posible Actualizar\el Saldo Pendiente de la O/C")
           GoTo 90
      End If
    End With
    '
    Call COMUNI("Recepcion Procesada con Éxito.")
90  Unload Me
    '
99  Command2.Enabled = True
End Sub

Private Sub Form_Load()
    If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture15.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      MARCOBARRA.Top = 5800
      Set MARCOBARRA.Container = Me.Picture15
    End If

End Sub

Private Sub salir_Click()
    Call BOTEXIT_Click
End Sub

Private Sub Text2_gotfocus()
    '
    BOTEXIT.SetFocus
    '
End Sub

Private Sub Text1_GotFocus()
    '
    BOTEXIT.SetFocus
    '
End Sub

