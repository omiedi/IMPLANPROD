VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form TRANEQUI07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transferencia de Equipos"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10365
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   6330
      Left            =   9480
      ScaleHeight     =   6300
      ScaleWidth      =   1785
      TabIndex        =   26
      Top             =   0
      Width           =   1815
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   32
         Top             =   5640
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   12000
         End
      End
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
         Height          =   730
         Left            =   105
         Picture         =   "TRANEQUI07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   4320
         Width           =   650
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "TRANEQUI07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   180
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "TRANEQUI07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Listado de Equipos por Depósito"
         Top             =   960
         Width           =   650
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Kdx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "TRANEQUI07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1760
         Width           =   650
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "TRANEQUI07.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3360
         Width           =   650
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   3960
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "FILTROS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   1440
      Width           =   9255
      Begin VB.TextBox Text4 
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
         Left            =   6960
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   315
         Width           =   2175
      End
      Begin VB.CommandButton Command7 
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
         Left            =   2880
         TabIndex        =   2
         Top             =   320
         Width           =   175
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "N.Serie"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6320
         TabIndex        =   23
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label10 
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
         Left            =   1080
         TabIndex        =   22
         Top             =   315
         Width           =   1860
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -1960
         TabIndex        =   21
         Top             =   360
         Width           =   2925
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   3120
         TabIndex        =   20
         Top             =   315
         Width           =   3060
      End
   End
   Begin VB.TextBox Text3 
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
      Height          =   300
      Left            =   1200
      TabIndex        =   4
      Top             =   5640
      Width           =   8065
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   9285
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         ItemData        =   "TRANEQUI07.frx":0EAA
         Left            =   120
         List            =   "TRANEQUI07.frx":0EAC
         Style           =   1  'Checkbox
         TabIndex        =   13
         Top             =   240
         Width           =   9045
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "    Código                           Descripción                                           Nº Serie                    F.Alta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   9060
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00CDDADA&
      Caption         =   "SELECCIÓN DE DEPÓSITOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   9240
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
         Left            =   8760
         TabIndex        =   15
         Top             =   720
         Width           =   175
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7680
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   720
         Width           =   1065
      End
      Begin VB.CommandButton Command2 
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
         Left            =   2760
         TabIndex        =   6
         Top             =   735
         Width           =   175
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
         Left            =   2760
         TabIndex        =   7
         Top             =   360
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         TabIndex        =   0
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         TabIndex        =   1
         Top             =   735
         Width           =   630
      End
      Begin VB.Label Label4 
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
         Height          =   225
         Left            =   7680
         TabIndex        =   16
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3105
         TabIndex        =   11
         Top             =   735
         Width           =   4020
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3105
         TabIndex        =   10
         Top             =   360
         Width           =   4020
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde Depósito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -1740
         TabIndex        =   9
         Top             =   430
         Width           =   3765
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hacia Depósito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -920
         TabIndex        =   8
         Top             =   820
         Width           =   2925
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "TRANEQUI07.frx":0EAE
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "TRANEQUI07.frx":10E2
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Referencia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   -2640
      TabIndex        =   18
      Top             =   5700
      Width           =   3765
   End
End
Attribute VB_Name = "TRANEQUI07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DEPOCONSU%
Dim TXTRAZA_TRANEQUI$

Sub BLANFORMU()
    Text1 = ""
    Text2 = ""
    Label10 = ""
    Text3 = ""
    Text4 = ""
    List2.Clear
End Sub

Private Sub Command1_Click()
    Call SELECHO("TADEPOSI")
    VDEV$ = VALACT1$("TADEPOSI")
    If VDEV$ <> "" Then
       Text1.TEXT = VDEV$
    End If
    Call Text4_Change ' ACTIVA LECTURA DE CODIGO DE BARRA
End Sub

Private Sub Command10_Click()
   Call CIERRARCH("*.*")
   Unload Me
End Sub

Private Sub Command2_Click()
    Call SELECHO("TADEPOSI")
    VDEV$ = VALACT1$("TADEPOSI")
    If VDEV$ <> "" Then
       Text2.TEXT = VDEV$
    End If
End Sub

Private Sub Command26_Click()
   Call CONECRUN("FSETUP04/OPTRAJUST", "Configuración de Funcionamiento")
   Call FINAL("")
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    HOOI% = HOY(HO$)
    TXTRAZA_TRANEQUI$ = ""
    If FECHANUM(Text10) < HOOI% - 30 Then
       Call MENSERR(24, "Fecha no Válida")
       GoTo 99
    End If
    '
    FORIPRE$ = TRIM$(Control$("", "FORTRAST"))
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Vale de Transferencia:\  \'" + RMCC$ + "'.")
            Exit Sub
        End If
    End If
    '
100 CMOV$ = TRIM$(Control$("", "CODTRAST"))
    If CMOV$ = "" Then CMOV$ = "TR"
    '
'    Call ABRESTOCKS ' COMENTADO POR QUE DA ERROR DE SINTAXIS ERROR QUE TOMA 12 CARACTERES
                     ' A PARTIR DE LA POS. 4 EN LA FUNCION UTILIZADA TOMA 6
     '
'    SQLX$ = " SELECT max(CAST(substring(DoPriCor, 4,12) AS INTEGER)) as Maximo FROM MOVISTO"
'    SQLX$ = SQLX$ + " WHERE CodiMovi = '" & CMOV$ & "'"
'    SQLX$ = SQLX$ + " AND SUBSTRING(DOPRICOR,1,2) <> 'RT'"
'    '
'    UNUM& = 0
'    Dim MIRS As ADODB.Recordset
'    Set MIRS = FLEXCONN.Execute(SQLX$)
'    If Not IsNull(MIRS!maximo) Then
'        UNUM& = XVALO(MIRS!maximo)
'    End If
'    NRO& = 1 + UNUM&
'    MIRS.Close
'    Set MIRS = Nothing
'    '
    TIPODOC$ = "Vale Transferencia de Stock"
    NRO& = TRIM$(ProNroComprobante(CMOV$)) ' SE TOMA EL PROXIMO NUMERO UTILIZANDO LA FUNCION
    NUMEDOC$ = TRIM$(Str$(NRO&))
    '
110 Screen.MousePointer = 1
    DepOrig% = XVALO(Text1): DEPOSI$ = ECOARCH$("ECODEP", Chr$(DepOrig%))
       If DepOrig% < 1 Then DEPOSI$ = "No Definido"
    DepoDes% = XVALO(Text2): DEPODE$ = ECOARCH$("ECODEP", Chr$(DepoDes%))
    If DEPODE$ = "" Then
       Call MENSERR(24, "Falta Deposito Destino")
       GoTo 99
    End If
    '
    FEMOVI% = FECHANUM(Text10)
    REFMO$ = TRIM$(Text3)
    If REFMO$ = "" Then
      Call MENSERR(24, "Referencia o Motivo es Mandatoria")
      GoTo 99
    End If
    '
    CAITMO% = 0
    For YX& = 1 To List2.ListCount
       If List2.Selected(YX& - 1) = True Then
          CI% = CODINT%(Left$(List2.List(YX& - 1), 16))
          CAN# = 1
          If CI% > 0 Then
            If CI% <= NUMAS% Then
              If CAN# > 0.005 Then
                CAITMO% = CAITMO% + 1
                CIMOVIM%(CAITMO%) = CI%
                CAMOVIM(CAITMO) = (-1) * CAN#
              End If
            End If
          End If
       End If
    Next YX&
    If PUEDEMOVI%(DepOrig%) < 0 Then GoTo 99 ' CANCELA SALIDA
    '
    HUBOSELOT% = 0
    NUORIT% = 0
    For YX& = 1 To List2.ListCount
       If List2.Selected(YX& - 1) = True Then
          CI% = CODINT%(Left$(List2.List(YX& - 1), 16))
          CAN# = 1
          DEPX% = DepOrig%
          LOTE$ = TRIM$(Mid$(List2.List(YX& - 1), 56, 16))
          REMI$ = "TR." + TRIM$(Str$(NRO&))
          If CI% > 0 Then
            If CI% <= NUMAS% Then
              If CAN# > 0.005 Then
                DOSEC$ = ""
1120                CANTI = (-1) * CAN#
                NUORIT% = NUORIT% + 1
                TXTRAZA_TRANEQUI$ = TXTRAZA_TRANEQUI$ & "PUNTO " & NUORIT% & "  = " & Now & vbCrLf
                Call MOVISTOK(FEMOVI%, CI%, LOTE$, "TR", REMI$, REMI$, DOSEC$, DOSEC$, REFMO$, NUORIT%, 0, 0, 0, DepOrig%, CANTI)
                NUORIT% = NUORIT% + 1
                Call MOVISTOK(FEMOVI%, CI%, LOTE$, "TR", REMI$, REMI$, DOSEC$, DOSEC$, REFMO$, NUORIT%, 0, 0, 0, DepoDes%, (-1) * CANTI)
                TXTRAZA_TRANEQUI$ = TXTRAZA_TRANEQUI$ & "PUNTO " & NUORIT% & "  = " & Now & vbCrLf
              End If
            End If
          End If
          '
          TXTRAZA_TRANEQUI$ = TXTRAZA_TRANEQUI$ & "PUNTO " & NUORIT% & "  = " & Now & vbCrLf
          Call ACTUREGISTRO("TANUMSE", "DEPONUMSE", "Cod_Inte = " & CI% & " AND NumeroSerie ='" & LOTE$ & "'", DepoDes%, REGAFECTA&)
          TXTRAZA_TRANEQUI$ = TXTRAZA_TRANEQUI$ & "PUNTO " & NUORIT% & "  = " & Now & vbCrLf
          '
          'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
          CUERPO$ = "Transferencia de Depósito: " & DepOrig%
          CUERPO$ = CUERPO$ + " A Depósito: " & DepoDes%
          If TRIM$(Text3) <> "" Then
            CUERPO$ = CUERPO$ & vbCrLf                  ' En el caso que tenga referencia
            CUERPO$ = CUERPO$ + "Ref.: " & TRIM$(Text3) 'genera un renglon mas con la misma
          End If
          Call ANOTANUSE(LOTE$, CUERPO$)
          TXTRAZA_TRANEQUI$ = TXTRAZA_TRANEQUI$ & "PUNTO " & NUORIT% & "  = " & Now & vbCrLf
          End If
    Next YX&
    '
    'IMPRESION DEL VALE DE TRANSFERENCIA MULTIPLE
    If FORIPRE$ <> "" Then
      CODPARAM$(1) = "NUME-COM": DOCPARAM$(1) = CStr(NRO&)
      CODPARAM$(2) = "FECH-EMI": DOCPARAM$(2) = FECHATEX(FEMOVI%)
      CODPARAM$(3) = "DEPOSITO": DOCPARAM$(3) = DEPOSI$
      CODPARAM$(4) = "DESTINO": DOCPARAM$(4) = DEPODE$
      CODPARAM$(5) = "REF-MAT1": DOCPARAM$(5) = REFMO$
      CAPARDOC% = 5
      '
      CODTABU1$(1) = "COD-MPRI"
      CODTABU1$(2) = "DES-MPRI"
      CODTABU1$(3) = "UNI-MPRI"
      CODTABU1$(4) = "CAN-MPRI"
      CODTABU1$(5) = "REF-MAT2"
      CACOLTAB1% = 5
      '
      CAITAB1% = 0
      '
      For YX& = 1 To List2.ListCount
         If List2.Selected(YX& - 1) = True Then
            CI% = CODINT%(Left$(List2.List(YX& - 1), 16))
            CAN# = 1
            '
            CAITAB1% = CAITAB1% + 1
              CODMAT$ = TRIM$(CODEXT(CI%))
            TETABU1$(1, CAITAB1%) = CODMAT$
              DESMAT$ = DESCRIT$(CI%)
            TETABU1$(2, CAITAB1%) = DESMAT$
              UNIMAT$ = Unimed$(CI%)
            TETABU1$(3, CAITAB1%) = UNIMAT$
              CANTMAT$ = FORMATNUM(CAN#, "F12.2")
            TETABU1$(4, CAITAB1%) = CANTMAT$
            TETABU1$(5, CAITAB1%) = TRIM$(Mid$(List2.List(YX& - 1), 56, 16))
         End If
      Next YX&
      '
      Call PRINTDOC("FORTRSTM")
      '
    End If
    Call COMUNI("Transferencia Procesada")
    'Unload Me
    Call BLANFORMU
    'GRABO EL ARCHIVO DE TRAZABILIDAD
For I1% = 1 To 32767
     ARCH$ = LUDAT$ & "TXTRAZA_TRANEQUI-" & TRIM$(Str$(I1%)) & ".TXT"
     If EXISTE%(ARCH$) < 1 Then
       NARCHI2% = FILEOPEN%(ARCH$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
       Print #NARCHI2%, TXTRAZA_TRANEQUI$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
       Close #NARCHI2%
       Exit For
     End If
Next I1%

    '
99  Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Call FINAL("?LINUMSEDEP") ' LISTADO DE EQUIPOS POR NUMERO DE SERIE
End Sub

Private Sub Command5_Click()
    '
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text10.TEXT = VDEV$
    End If

End Sub

Private Sub Command6_Click()
    Command4.Enabled = False
    Call SITSTOK07.Command3_Click
    Command4.Enabled = True
End Sub

Private Sub Command7_Click()
    '
    Label10 = ""
    Call SELECHO("MASTER")
    COD$ = VALACT1$("MASTER")
    Label10 = COD$
    '
End Sub

Private Sub Form_Load()
    Call APLICACIONSKINS(Me, Picture1)

    Text10 = FECHATEX$(HOY(HO$))
    '
    Call CREATABLA_EQUIPOS
    '
    SQLX$ = "UPDATE TANUMSE SET NumRemito = '' WHERE NumRemito IS NULL"
    On Error Resume Next
    FLEXCONN.Execute SQLX$
    On Error GoTo 0
    SQLX$ = "UPDATE TANUMSE SET NumFactura = '' WHERE NumFactura IS NULL"
    On Error Resume Next
    FLEXCONN.Execute SQLX$
    On Error GoTo 0
    Text4.MaxLength = 12 ' TEXT DE  NUMERO DE SERIE
End Sub

Private Sub Label10_Change()
    '
    Label8 = ""
    CIXI% = CODINT%(Label10)
    If CIXI% > 0 Then
       Label8 = DESCRIT0$(CIXI%)
    End If
    '
    Call CARGALALISTA
    '
    Call Text4_Change ' ACTIVA LECTURA DE CODIGO DE BARRA

End Sub

Sub CARGALALISTA()
    '
   CONDI$ = "DEPONUMSE = " & DEPOCONSU% & " AND (NumRemito = '' OR NumRemito IS NULL) AND (NumFactura = '' OR NumFactura IS NULL) "
    If Label10 <> "" Then
       CODI% = CODINT(TRIM$(Label10))
       CONDI$ = CONDI$ + " AND Cod_Inte = " & CODI%
    End If
    CONDI$ = CONDI$ + " ORDER BY COD_EXTE, NUMEROSERIE"
    Call CARGALISTA(List2, "TANUMSE", "COD_INTE/C20+MASTER/A33; NUMEROSERIE/A17; FECHALTA/D8 ", CONDI$)
    '
End Sub

Private Sub Text1_Change()
  '
  DEPOCONSU% = XVALO(Text1)
  Label1 = ECOARCH("TADEPOSI", Chr$(DEPOCONSU%))
  '
  Call CARGALALISTA
  '
  Label10 = ""
  Label8 = ""
  '
End Sub

Private Sub Text2_Change()
  DEPO% = XVALO(Text2)
  Label2 = ECOARCH("TADEPOSI", Chr$(DEPO%))
End Sub

Private Sub Text4_Change()
   
   If List2.ListCount > 0 Then ' CUANDO INGRESA EL NUMERO DE SERIE SI EXISTE EN EL LIST SE POSICIONA
     For i& = 1 To List2.ListCount
        If TRIM$(Text4) = TRIM$(Mid$(List2.List(i& - 1), 56, 16)) Then
          List2.ListIndex = i& - 1
          Exit For
        Else
          List2.ListIndex = 0
        End If
     Next i&
   End If
End Sub


