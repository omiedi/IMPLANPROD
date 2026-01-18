VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form REINEQUI07 
   BackColor       =   &H00CDDADA&
   Caption         =   "Reingreso de Equipos"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   31
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "."
      Height          =   300
      Left            =   7080
      TabIndex        =   27
      Top             =   120
      Width           =   175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "CONFIRMA REINGRESO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   120
      TabIndex        =   21
      Top             =   5450
      Width           =   7245
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
         Height          =   570
         Left            =   6240
         Picture         =   "REINEQUI07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Aprobar Reingreso"
         Top             =   240
         Width           =   810
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         Height          =   300
         Left            =   2115
         TabIndex        =   25
         Top             =   360
         Width           =   175
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
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
         Left            =   1560
         TabIndex        =   22
         Text            =   " "
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   820
         TabIndex        =   24
         Top             =   450
         Width           =   810
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2400
         TabIndex        =   23
         Top             =   360
         Width           =   3555
      End
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7650
      Left            =   7440
      ScaleHeight     =   7620
      ScaleWidth      =   1350
      TabIndex        =   6
      Top             =   0
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   32
         Top             =   6120
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
      Begin VB.CommandButton Command26 
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
         Left            =   120
         Picture         =   "REINEQUI07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Configurar Etiqueta de Reingreso de Equipos"
         Top             =   2520
         Width           =   690
      End
      Begin VB.CommandButton Command20 
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
         Left            =   120
         Picture         =   "REINEQUI07.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Búsqueda por Número de Serie"
         Top             =   1680
         Width           =   690
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Help"
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
         Left            =   120
         Picture         =   "REINEQUI07.frx":1AB2
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   870
         Width           =   690
      End
      Begin VB.CommandButton Command11 
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
         Left            =   120
         Picture         =   "REINEQUI07.frx":1DBC
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   690
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "EQUIPOS ENTREGADOS"
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
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   7245
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
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
         Height          =   1470
         ItemData        =   "REINEQUI07.frx":1F06
         Left            =   120
         List            =   "REINEQUI07.frx":1F08
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   6945
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha                        N.Serie                                        Remito"
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
         Height          =   330
         Left            =   120
         TabIndex        =   5
         Top             =   330
         Width           =   6945
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00CDDADA&
      Caption         =   "DATOS DE ENTREGA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   7245
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Factura"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3920
         TabIndex        =   20
         Top             =   1650
         Width           =   1050
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
         Height          =   300
         Left            =   4800
         TabIndex        =   19
         Top             =   1560
         Width           =   2220
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Remito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4000
         TabIndex        =   18
         Top             =   1290
         Width           =   1005
      End
      Begin VB.Label Label9 
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
         Height          =   300
         Left            =   4800
         TabIndex        =   17
         Top             =   1200
         Width           =   2220
      End
      Begin VB.Label Label8 
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
         Height          =   285
         Left            =   1000
         TabIndex        =   16
         Top             =   1290
         Width           =   690
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "         "
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
         Height          =   300
         Left            =   1560
         TabIndex        =   15
         Top             =   1200
         Width           =   1020
      End
      Begin VB.Label Label5 
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
         Height          =   300
         Left            =   1560
         TabIndex        =   14
         Top             =   750
         Width           =   5460
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         TabIndex        =   13
         Top             =   500
         Width           =   690
      End
      Begin VB.Label Label3 
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
         Height          =   300
         Left            =   6000
         TabIndex        =   12
         Top             =   405
         Width           =   1020
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Número Serie"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   2
         Top             =   500
         Width           =   1290
      End
      Begin VB.Label Label1 
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
         Height          =   300
         Left            =   1560
         TabIndex        =   1
         Top             =   405
         Width           =   2220
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "REINEQUI07.frx":1F0A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "REINEQUI07.frx":213E
      TabIndex        =   30
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label2 
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
      Height          =   300
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   7140
   End
   Begin VB.Label Label22 
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
      Height          =   300
      Left            =   3240
      TabIndex        =   10
      Top             =   240
      Width           =   1545
   End
   Begin VB.Label Label28 
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
      Height          =   300
      Left            =   4960
      TabIndex        =   9
      Top             =   120
      Width           =   2100
   End
End
Attribute VB_Name = "REINEQUI07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call SELECHO("TADEPOSI")
    VDEV$ = VALACT1$("TADEPOSI")
    If VDEV$ <> "" Then
       Text4.TEXT = VDEV$
    End If

End Sub

Private Sub Command11_Click()
     Call CIERRARCH("*.*")
     Unload Me

End Sub

Private Sub Command2_Click()
    Call SELECHO("MASTER")
    COD$ = VALACT1$("MASTER")
    Label28 = COD$
    CI% = CODINT%(COD$)
    Label2 = DESCRIT$(CI%)

End Sub


Private Sub Command20_Click()
   Static ARGU$
   Command20.Enabled = False ' BUSQUEDA POR NUMERO DE SERIE
   BAND% = 0 ' BANDERA PARA VERIFICAR SI ENCONTRO EL ARGUMENTO
10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda por Número de Serie", ARGU$)))
   If ARGU$ <> "" Then
    Screen.MousePointer = 11
    If List1.ListCount > 0 Then ' CUANDO INGRESA EL NUMERO DE SERIE SI EXISTE EN EL LIST SE POSICIONA
     For i& = 1 To List1.ListCount
        If TRIM$(ARGU$) = TRIM$(Mid$(List1.List(i& - 1), 15, 22)) Then
          List1.ListIndex = i& - 1
          BAND% = 1
          Exit For
        Else
        End If
     Next i&
    End If
   End If
   If BAND% < 1 Then Call MENSERR(24, "Argumento No Encontrado")
   Screen.MousePointer = 1
   Command20.Enabled = True
End Sub

Private Sub Command26_Click()
   '
   Command26.Enabled = False
   '
   If DERACCE%("CFGREQUI", "Configuración Re-Ingreso de Equipos") >= 2 Then
       FSTREINEQ.Show 1
   End If
   '
   Command26.Enabled = True
   '
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   If Label1 = "" Then
     Call MENSERR(24, "Falta Número de Serie")
     GoTo 99
   End If
   '
   DE2$ = ECOARCH("TADEPOSI", Chr$(XVALO(Text4)))
   If DE2$ = "" Then
     Call MENSERR(24, "Depósito No Valido")
     GoTo 99
   End If
   '
   Screen.MousePointer = 11
   fech% = HOY(HOS$)
   FECHAA$ = FECHATEX$(fech%)
   CMOV$ = "DV"
   NVM$ = TRIM$(ProNroComprobante("DV"))
   While Len(NVM$) < 6: NVM$ = "0" + NVM$: Wend
   NRO& = XVALO(NVM$)
   DOSEC$ = "AJ." + Mid$(FECHAA$, 7, 2) + Mid$(FECHAA$, 4, 2) + Left$(FECHAA$, 2)
   DOPRI$ = CMOV$ + "." + NVM$
   REFE$ = "Devol. S/N: " & TRIM$(Label1)
   DEPX% = XVALO(Text4)
   CANTI = 1
   CIXI% = CODINT%(Label28)
   LOTE$ = TRIM$(Label1)
   FECHENT = CVDAT(fech%)
   Call MOVISTOK(fech%, CIXI%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CANTI)
   Condi$ = "NUMEROSERIE = '" & TRIM$(Label1) & "'"
   Call ACTUREGISTRO("Tanumse", "NumRemito", Condi$, "", REG&, "NOMESS")
   Call ACTUREGISTRO("Tanumse", "NumFactura", Condi$, "", REG&, "NOMESS")
   Call ACTUREGISTRO("Tanumse", "FechEntrega", Condi$, "", REG&, "NOMESS")
   Call ACTUREGISTRO("Tanumse", "NumeCli", Condi$, 0, REG&, "NOMESS")
   Call ACTUREGISTRO("Tanumse", "DepoNumse", Condi$, DEPX%, REG&, "NOMESS")
   ' OMAR - Disparate 6/10/09: Call ACTUREGISTRO("Tanumse", "FeVenGar", "CONDI$", "", NUAFEC&, "NOMESS")
   Call ACTUREGISTRO("Tanumse", "FeVenGar", Condi$, "", NUAFEC&, "NOMESS")
   Call ACTUREGISTRO("PEDDETA", "NUMEROSERIE", Condi$, "", REG&, "NOMESS")
   '
   'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
   CUERPO$ = "Reingreso de Equipo"
   NSERIE$ = TRIM$(Label1)
   Call ANOTANUSE(NSERIE$, CUERPO$)
   '
   Screen.MousePointer = 1
   '
   Call PRINTREINEQUI
   '
   Call MENSERR(24, "Reingreso Realizado")
   '
   Call BLANFORMU
   
99 Command7.Enabled = True

End Sub

Private Sub Form_Load()

   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)

End Sub

Private Sub Label28_Change()
   LIMPIA_DatosEntrega
   CI% = CODINT%(Label28)
   List1.Clear
   If CI% > 0 Then
     Condi$ = "Cod_Inte = " & CI% & " AND NumRemito <> '' AND NumRemito IS NOT NULL"
     Call CARGALISTA(List1, "TANUMSE", "FechEntrega/D14;NumeroSerie/A22;NumRemito/A18", Condi$)
   End If
End Sub

Private Sub List1_DblClick()
   Call LIMPIA_DatosEntrega
   NUMSERX$ = TRIM$(Mid$(List1.TEXT, 15, 22))
   Label1 = VALOBADA("TANUMSE", "NUMEROSERIE", " NUMEROSERIE = '" & NUMSERX$ & "'")
   Label3 = VALOBADA("TANUMSE", "NUMECLI", " NUMEROSERIE = '" & NUMSERX$ & "'")
   Label5 = rasocli$(XVALO(Label3))
   Label6 = FECHATEX(CVINT(VALOBADA("TANUMSE", "FECHENTREGA", " NUMEROSERIE = '" & NUMSERX$ & "'")))
   Label9 = VALOBADA("TANUMSE", "NUMREMITO", " NUMEROSERIE = '" & NUMSERX$ & "'")
   Label11 = VALOBADA("TANUMSE", "NUMFACTURA", " NUMEROSERIE = '" & NUMSERX$ & "'")
End Sub

Sub LIMPIA_DatosEntrega()

   Label1 = ""
   Label3 = ""
   Label5 = ""
   Label6 = ""
   Label9 = ""
   Label11 = ""
End Sub

Private Sub Text4_Change()
  DEPOCONSU% = XVALO(Text4)
  Label13 = ECOARCH("TADEPOSI", Chr$(DEPOCONSU%))

End Sub
Sub BLANFORMU()
  
  Label28 = ""
  Label2 = ""
  Text4 = ""
  Label13 = ""
  Call LIMPIA_DatosEntrega
  List1.Clear
End Sub

 Sub PRINTREINEQUI()
    '
    FORIPRE$ = TRIM$(Control$("", "REINEQ"))
    If FORIPRE$ = "" Then Exit Sub
    '
    NUSERIE$ = Label1.Caption
    CODIIGG$ = Label28.Caption
    '
    NUMEDOC$ = NUSERIE$
    TIPODOC$ = "Reingreso de Equipo"
    '
    CODPARAM$(1) = "DOCUORIG": DOCPARAM$(1) = NUSERIE$
    CODPARAM$(2) = "COD-MAT": DOCPARAM$(2) = CODIIGG$
    CODPARAM$(3) = "CUEN-CLI": DOCPARAM$(3) = Label3.Caption
    CODPARAM$(4) = "REF-MAT1": DOCPARAM$(4) = Label9.Caption
    CODPARAM$(5) = "REF-MAT2": DOCPARAM$(5) = Label11.Caption
    CODPARAM$(6) = "DES-MAT": DOCPARAM$(6) = Label2.Caption
    CODPARAM$(7) = "RASO-CLI": DOCPARAM$(7) = Label5.Caption
    CODPARAM$(8) = "FECH-EMI": DOCPARAM$(8) = Label6.Caption
    
    CAPARDOC = 8
    '
    Call PRINTDOC("@" + FORIPRE$)
    '
 End Sub
