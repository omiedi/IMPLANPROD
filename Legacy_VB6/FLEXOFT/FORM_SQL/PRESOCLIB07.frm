VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form PRESOCLIB07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recepción de Material de Texto Libre"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   40
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00CDDADA&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   120
      TabIndex        =   33
      Top             =   4560
      Width           =   4575
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   2310
         TabIndex        =   34
         Top             =   630
         Width           =   2010
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   720
         TabIndex        =   39
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "E.V."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -1290
         TabIndex        =   38
         Top             =   310
         Width           =   1920
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3450
         TabIndex        =   37
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   36
         Top             =   310
         Width           =   1920
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Certificado de Calidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   105
         TabIndex        =   35
         Top             =   690
         Width           =   2130
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00CDDADA&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   4800
      TabIndex        =   28
      Top             =   4560
      Width           =   2535
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   860
         TabIndex        =   31
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   860
         TabIndex        =   30
         Top             =   630
         Width           =   1485
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Control Vigilancia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   -350
         TabIndex        =   29
         Top             =   520
         Width           =   1080
      End
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   6945
      Left            =   7440
      ScaleHeight     =   6915
      ScaleWidth      =   1035
      TabIndex        =   16
      Top             =   -120
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   42
         Top             =   6240
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   12000
         End
      End
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
         Picture         =   "PRESOCLIB07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Left            =   105
         Picture         =   "PRESOCLIB07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Aprueba la Selección de Items a Comprar"
         Top             =   4725
         Width           =   650
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   13
      Top             =   5760
      Width           =   3360
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1005
         TabIndex        =   23
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito Destino"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -240
         TabIndex        =   22
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   1155
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   5760
      Width           =   3750
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   2625
         TabIndex        =   25
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Recibido"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1830
         TabIndex        =   24
         Top             =   420
         Width           =   720
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pendiente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   840
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   7215
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
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
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   1140
         TabIndex        =   5
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Material"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
      BackColor       =   &H00CDDADA&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   120
      TabIndex        =   0
      Top             =   0
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
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   4320
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   3480
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   2760
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1140
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº O/C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5955
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1140
         TabIndex        =   3
         Top             =   615
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2295
         TabIndex        =   1
         Top             =   615
         Width           =   4755
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "PRESOCLIB07.frx":0614
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "PRESOCLIB07.frx":0848
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "PRESOCLIB07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    If CONTROL$("BOLRECEP", "CERCAPRO") = "SI" Then
       If CERCAPRO$ = "" Then
          If COMALTER%("Falta Certificado de Calidad del Proveedor.\\Corregir\Continuar") <> 2 Then
             Command2.Enabled = True
             Exit Sub
          End If
       End If
    End If
    '
    Call ABRESTOCKS
    'With Bolrecep
      '
      NC% = Val(LABEL5)
      NUOC& = Val(Label4)
      CARECI# = Val(Text3)
      NBORE& = Val(Label9)
      REMI$ = TRIM$(Label22)
      DEPX% = Val(Label13)
      NORIT& = Val(Label12)
      NUCONVI$ = Left$(TRIM$(Label24), 16)
      
      SQLX$ = "SELECT * FROM BOLRECEP"
      SQLX$ = SQLX$ + " WHERE NumeBoRe = " & NBORE& & ""
      Dim BOLRRRECEP As ADODB.Recordset
      Set BOLRRRECEP = New ADODB.Recordset
      BOLRRRECEP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
      With BOLRRRECEP
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CODIEMPR = CodiEmp%
      On Error GoTo 0
       !NUMEBORE = NBORE&
       !Descrip = Left$(Text1, 64)
       !idenlote = " "
       !fechrecep = CVDAT(CVINT(Label2))
       !FEYHORecep = Now
       !COD_INTE = 0
       !Cod_Exte = " "
       !descripitem = Left$(Text1, 64) 'Left$(DESCRIT0$(0), 64)
       !Unidad = "PZA"
       !cantnom = CARECI#
       !DEPOENT = DEPX%
       !NumProv = NC%
       !Raso_Prov = Left$(rasocli$((-1) * NC%), 64)
          LOTECO$ = " "
          While Len(LOTECO$) > 12
            LOTECO$ = Mid$(LOTECO$, 2)
          Wend
       !LOTECOL_PROV = Left$(LOTECO$, 12)
       !numrem_prov = Left$(REMI$, 16)
       !CerCal_Prov = CERCAPRO$
       !NumFac_Prov = ""
       !NumConVig = NUCONVI$
       !NUMOC = NUOC&
       !IContab = Val(Label14) 'Imputación contable de la OC.
       !MoneVaUni = ""
       !Valo_Unit = 0
       !CanToApro = CARECI#
       !CanToRech = 0
       !CanToFalta = 0
       !FechLibe = CVDAT(CVINT(Label2))
       !NUMEORDSER = 0
       !Observa = USERNAME
       !Status = 1
       !USUAPRO = USERNAME
       !FeHoApro = Now
      .Update
      '
    End With
    BOLRRRECEP.Close
    Set BOLRRRECEP = Nothing
    '
    'MARCO EN OCOMDETA LA CANT. RECIBIDA
    Call ABRECOMPRAS
    'With OcomDeta
      SQLX$ = "SELECT * FROM OCOMDETA"
      SQLX$ = SQLX$ + " WHERE NUME_OCOM = " & NUOC& & ""
      SQLX$ = SQLX$ + " AND NuOrd_Item = " & NORIT& & ""
      SQLX$ = SQLX$ + " AND Stat_Ocom < 8 "         ' NO ANULADAS
      SQLX$ = SQLX$ + " AND  Stat_Ocom <> - 1 "     ' EXCLUYE NO APROB.
      Dim OKKKOMDETA As ADODB.Recordset
      Set OKKKOMDETA = New ADODB.Recordset
      OKKKOMDETA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
      With OKKKOMDETA
      '.FindFirst SQLX$
      'If .NoMatch = False Then
           '.Edit
      If Not (.EOF And .BOF) Then
           !Cant_Rec = !Cant_Rec + CARECI#
           CAPENDIEN = !cant_ocom - !Cant_Rec
           If CAPENDIEN < 0.05 Then
             !Stat_Ocom = 3
           End If
           .Update
          Else
           Call MENSERR(24, "No Fue Posible Actualizar\el Saldo Pendiente de la O/C")
           GoTo 90
      End If
    End With
    OKKKOMDETA.Close
    Set OKKKOMDETA = Nothing
    '
    Call COMUNI("Recepcion Procesada con Éxito.")
90  Unload Me
    '
99  Command2.Enabled = True
End Sub

Private Sub Form_Load()
  '
  UTILIZA_SKIN% = 1
  Call APLICACIONSKINS(Me)
  '
End Sub

Private Sub Text2_gotfocus()
    '
    BOTEXIT.SetFocus
    '
End Sub
'
Private Sub Text1_GotFocus()
    '
    BOTEXIT.SetFocus
    '
End Sub

Private Sub Text5_gotfocus()
    '
    BOTEXIT.SetFocus
    '
End Sub



