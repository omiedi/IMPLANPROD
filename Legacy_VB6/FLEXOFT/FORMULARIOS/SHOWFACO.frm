VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form SHOWFACO 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle del Movimiento"
   ClientHeight    =   6465
   ClientLeft      =   1050
   ClientTop       =   5835
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid GRID2 
      Height          =   765
      Left            =   2880
      TabIndex        =   34
      Top             =   240
      Visible         =   0   'False
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   1349
      _Version        =   393216
      Rows            =   23
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   8433840
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      Redraw          =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1560
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
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
      Height          =   405
      Left            =   165
      Picture         =   "SHOWFACO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Imprime Formulario de Carga - Vuelco de Pantalla"
      Top             =   6030
      Width           =   650
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   5460
      ScaleHeight     =   855
      ScaleWidth      =   540
      TabIndex        =   28
      Top             =   735
      Visible         =   0   'False
      Width           =   540
      Begin VB.Image Image1 
         Height          =   480
         Left            =   0
         Picture         =   "SHOWFACO.frx":066A
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         TabIndex        =   29
         Top             =   460
         Width           =   450
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Recupera Datos"
      Height          =   225
      Left            =   3000
      TabIndex        =   22
      Top             =   20
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4190
      Left            =   165
      ScaleHeight     =   4155
      ScaleWidth      =   9645
      TabIndex        =   0
      Top             =   1680
      Width           =   9670
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   4245
         Left            =   6825
         TabIndex        =   7
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   4245
         Left            =   9345
         TabIndex        =   6
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   4245
         Left            =   8085
         TabIndex        =   5
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   4245
         Left            =   2475
         TabIndex        =   4
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   4245
         Left            =   1350
         TabIndex        =   3
         Top             =   0
         Width           =   15
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Roboto Mono Medium"
            Size            =   9
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3870
         ItemData        =   "SHOWFACO.frx":0974
         Left            =   -30
         List            =   "SHOWFACO.frx":0976
         TabIndex        =   1
         Top             =   300
         Width           =   9690
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Roboto Mono Medium"
            Size            =   9
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   11355
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "SHOWFACO.frx":0978
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   1080
      OleObjectBlob   =   "SHOWFACO.frx":0BAC
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label22 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2280
      TabIndex        =   30
      Top             =   6090
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6000
      TabIndex        =   27
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label19 
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
      Height          =   255
      Left            =   6990
      TabIndex        =   26
      Top             =   120
      Width           =   2520
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   7770
      TabIndex        =   25
      Top             =   315
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   6510
      TabIndex        =   24
      Top             =   315
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   5355
      TabIndex        =   23
      Top             =   315
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label14 
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
      Height          =   540
      Left            =   315
      TabIndex        =   20
      Top             =   315
      Width           =   4875
   End
   Begin VB.Label Label12 
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
      Height          =   255
      Left            =   4080
      TabIndex        =   18
      Top             =   1230
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "Cotización:"
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
      Left            =   3885
      TabIndex        =   19
      Top             =   990
      Width           =   1170
   End
   Begin VB.Label Label11 
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
      Height          =   255
      Left            =   8310
      TabIndex        =   17
      Top             =   1260
      Width           =   1200
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Diferencia de Cambio ($):"
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
      Left            =   6525
      TabIndex        =   16
      Top             =   1155
      Width           =   1590
   End
   Begin VB.Label Label9 
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
      Height          =   255
      Left            =   8310
      TabIndex        =   15
      Top             =   630
      Width           =   1200
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor Actualizado en Pesos:"
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
      Left            =   6525
      TabIndex        =   14
      Top             =   525
      Width           =   1590
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
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   1230
      Width           =   1200
   End
   Begin VB.Label Label6 
      Caption         =   "Importe:"
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
      Left            =   2310
      TabIndex        =   12
      Top             =   990
      Width           =   1170
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
      Height          =   270
      Left            =   315
      TabIndex        =   11
      Top             =   1230
      Width           =   1905
   End
   Begin VB.Label Label4 
      Caption         =   "Moneda:"
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
      Left            =   105
      TabIndex        =   10
      Top             =   990
      Width           =   1170
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   8340
      TabIndex        =   9
      Top             =   6060
      Width           =   1200
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   7080
      TabIndex        =   8
      Top             =   6060
      Width           =   1200
   End
   Begin VB.Label Label15 
      Caption         =   "Referencia / Observación:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   21
      Top             =   75
      Width           =   2850
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo "
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "SHOWFACO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   '
   Static CAMP%, CUEMP%(2048), MEPA%(2048)
   '
   If CAMP% <= 0 Then
      For i& = 1 To ULTREG&("MEPAGO")
         F0$ = REGLEIDO$("MEPAGO", i&)
         CAMP% = i&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         CUE$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT%(CUE$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUE$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next i&
   End If
   '
   Screen.MousePointer = 11
   '
   NOPANT = XVALO(Label16): NPANT% = 0
   ATX% = 0
   FIN& = ULTREG&("CHEQUEMI")
   For IAX& = 1 To FIN&
      Call TRACE(20, IAX&, FIN&)
      CH0$ = REGLEIDO$("CHEQUEMI", IAX&)
      CH1$ = Mid$(CH0$, 1, 2)
      CH5$ = Mid$(CH0$, 33, 8)
      CH6$ = Mid$(CH0$, 41, 4)
      CH7$ = Mid$(CH0$, 45, 2)
      CH8$ = Mid$(CH0$, 47, 2)
      '
      If CVS(CH6$) = NOPANT Then
         IM# = CVD(CH5$)
         If Abs(IM#) >= 0.005 Then
            NOPA = CVS(CH6$)
            '
            If NOPA < 1 Then GoTo 76 ' INCLUIDO POR DOSIVAC - SCARINCI DEL 26/12/04
            NP% = CVI(CH1$)         ' NRO DE PROVEEDOR
            TIPOVA% = CVI(CH7$)
            QUE% = CVI(CH7$)
            '
            'If NOPA <> NOPANT Or NP% <> NPANT% Then
               ITOT# = 0
               For J& = IAX& To FIN&
                  CH0$ = REGLEIDO$("CHEQUEMI", J&)
                  CH1$ = Mid$(CH0$, 1, 2)
                  CH5$ = Mid$(CH0$, 33, 8)
                  CH6$ = Mid$(CH0$, 41, 4)
                  If CVI(CH1$) = NP% Then
                    If CVS(CH6$) = NOPA Then
                      If CVI(CH8$) = FF% Then
                        ITOT# = ITOT# + CVD(CH5$)
                        GoTo 14
                      End If
                    End If
                  End If
                  GoTo 15
14             Next J&
               '
15             If TICUEPRO((-1) * NP%) = 1 Then ' CUENTA DE GASTOS
                  For UDP& = 1 To ULTREG&("DETPAGAS")
                     DTP$ = REGLEIDO$("DETPAGAS", UDP&)
                     If CVS(Mid$(DTP$, 3, 4)) = NOPA Then
                       If CVI(Mid$(DTP$, 7, 2)) = FF% Then
                         CUEDP% = CVI(Mid$(DTP$, 17, 2))
                         If TRIM$(CODCUE$(CUEDP%)) <> "" Then
                           IDP# = CVD(Mid$(DTP$, 9, 8))
                           Z$ = Space$(128)
                           Call REPLA(Z$, CODCUE$(CUEDP%), 1, 12)
                           DETAX$ = "OP " + TRIM$(Str$(NOPA)) + " " + rasocli$((-1) * NP%)
                           Call REPLA(Z$, DETAX$, 14, 27)
                           If IDP# > 0 Then
                                Call REPLA(Z$, CSTRING$(MKD$(IDP#), 3, 12, 2, 2), 41, 12)
                              Else
                                Call REPLA(Z$, CSTRING$(MKD$(Abs(IDP#)), 3, 12, 2, 2), 53, 12)
                           End If
                           LIST1.AddItem Z$
                           ITOT# = ITOT# - IDP#
                         End If
                       End If
                     End If
                  Next UDP&
               End If
               '
               If Abs(ITOT#) >= 0.005 Then
                  Z$ = Space$(128)
                  Call REPLA(Z$, CUEMADRE$((-1) * NP%), 1, 12)
                  DETAX$ = "OP " + TRIM$(Str$(NOPA)) + " " + rasocli$((-1) * NP%)
                  Call REPLA(Z$, DETAX$, 14, 27)
                  If ITOT# > 0 Then
                       Call REPLA(Z$, CSTRING$(MKD$(ITOT#), 3, 12, 2, 2), 41, 12)
                     Else
                       Call REPLA(Z$, CSTRING$(MKD$(Abs(ITOT#)), 3, 12, 2, 2), 53, 12)
                  End If
                 'List1.AddItem Z$
               End If
               '
 '              NOPANT = NOPA
 '              NPANT% = NP%
            'End If
            '
            CH0$ = REGLEIDO$("CHEQUEMI", IAX&)
            CH1$ = Mid$(CH0$, 1, 2)
            CH2$ = Mid$(CH0$, 3, 12)
            CH3$ = Mid$(CH0$, 15, 16)
            CH5$ = Mid$(CH0$, 33, 8)
            CH6$ = Mid$(CH0$, 41, 4)
            CH7$ = Mid$(CH0$, 45, 2)
            CH8$ = Mid$(CH0$, 47, 2)
            If Abs(CVD(CH5$)) < 0.005 Then GoTo 76
            '
            CHE$ = ""
            For JI% = 1 To Len(CH2$)
               If Asc(Mid$(CH2$, JI%)) > 32 Then CHE$ = CHE$ + Mid$(CH2$, JI%, 1)
            Next JI%
            '
            NCHE$ = ""
            For JI% = 1 To Len(CH3$)
               If Asc(Mid$(CH3$, JI%)) > 32 Then NCHE$ = NCHE$ + Mid$(CH3$, JI%, 1)
            Next JI%
            '
            If QUE% < 1 Then
              EPAC$ = TRIM$(UCase$(EMPREAC$))
              If InStr(EPAC$, "SABO") > 0 Then
                If Left$(CH3$, 10) = "Ret.IB.Bs." Then
                  QUE% = 30
                End If
              End If
            End If
            '
            For UX% = 1 To CAMP%
               If MEPA%(UX%) = QUE% Then GoTo 18
            Next UX%
            Call MENSERR(24, "Medio de Pago" + Str$(QUE%) + " no Definido")
            Call FINAL("")
            '
18          CUEH% = CUEMP%(UX%)
            IPASE# = CVD(CH5$)
            DTX$ = TRIM$(CHE$)
            If DTX$ = "" Then DTX$ = TRIM$(NCHE$)
            If DTX$ <> "" Then DTX$ = DTX$ + " - "
            DTX$ = DTX$ + "OP" + Str$(NOPA) + " " + rasocli$((-1) * NP%)
            Z$ = Space$(128)
            Call REPLA(Z$, CODCUE$(CUEH%), 1, 12)
            Call REPLA(Z$, DTX$, 14, 27)
            Call REPLA(Z$, CSTRING$(MKD$(IPASE#), 3, 12, 2, 2), 53, 12)
            ' List1.AddItem Z$
            '
         End If
      End If
76 Next IAX&
   '
   On Error Resume Next
   CUEVIEX$ = MAYORPAN07.Ltext2
   If Err Then
      On Error GoTo 0
      Exit Sub
   End If
   On Error GoTo 0
   If TRIM$(CUEVIEX$) = "" Then Exit Sub
   CUENUE% = CUEH%
   CUEVIE% = CUEINT%(CUEVIEX$)
   If CUENUE% < 1 Or CUEVIE% < 1 Then Exit Sub
   '
   If COMALTER%("Reemplaza Cuenta '" + CUEVIEX$ + "' por '" + CODCUE$(CUENUE%) + "' ?\\Si, Reemplazar\No, Continuar") = 1 Then
      SQLX$ = " update CAJABANC SET CueContaI = " & CUENUE% & " , DescriMov = '" & Left$(DTX$, 48) & "' , RefeCom = '" & Left$(DTX$, 48) & "' "
      SQLX$ = SQLX$ + " WHERE FechConta = '" & Label17 & "' "
      SQLX$ = SQLX$ + " AND CodiCom_LAR= '" & Label18 & "' "
      SQLX$ = SQLX$ + " AND CueContaI = " & CUEVIE% & " "
       FLEXCONN.Execute (FILTSQL$(SQLX$))
      Unload Me
   End If
   '
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command5_Click()
    PRINTFORM
End Sub

Private Sub Form_Activate()
    If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
    End If

   'Call REMPLAGRID(Me, GRID2, Picture1, Label1, List1)
End Sub

Private Sub Form_Load()
    '
    If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
    End If

    Call CENTRAFORM(Me)
    T0$ = "Cuenta"
    T1$ = "Referencia"
    T2$ = "Debe"
    T3$ = "Haber"
    '
    Label1.Caption = " " + T0$ & Space(7) & "C.Costo   " & T1$ & Space(38) & T2$ & Space(7) & T3
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    If NIVCONFI% > 0 Then
       Label1.ToolTipText = "Doble Click para Clasificar"
    End If
    
    Call AplicaFuenteAnchoFijo(LIST1, , 9)
    Call AplicaFuenteAnchoFijo(Label2, , 9)
    Call AplicaFuenteAnchoFijo(Label3, , 9)
    
End Sub
'
Private Sub Label1_DblClick()
   If LIST1.ListCount > 0 Then
      NICON% = NIVCONFI%
      If NICON% > 0 Then
         Label21 = "<" + TRIM$(Str$(NICON%))
         Picture2.Visible = True
         '
         SQLZ$ = Label22
         If SQLZ$ <> "" Then
            SQLX$ = "UPDATE CAJABANC SET NIV_CLASI = " & NIVCONFI% & " " & SQLZ$
            SQLY$ = "UPDATE SACREPEN SET NIV_CLASI = " & NIVCONFI% & " " & SQLZ$
            On Error Resume Next
            FLEXCONN.Execute SQLX$
            FLEXCONN.Execute SQLY$
            On Error GoTo 0
         End If
      End If
   End If
End Sub

Private Sub Label21_DblClick()
   If DERACCE%("BAJNICLA/NOMESS", "Reducir Nivel de Clasificación") >= 2 Then
      NICLAC% = XVALO(Mid$(Label21, 2))
      If NICLAC% > 0 Then
         Label21 = "<" + TRIM$(Str$(NICLAC% - 1))
         Label21.Refresh
         SQLZ$ = Label22
         If SQLZ$ <> "" Then
            SQLX$ = "UPDATE CAJABANC SET NIV_CLASI = " & (NICLAC% - 1) & " " & SQLZ$
            SQLY$ = "UPDATE SACREPEN SET NIV_CLASI = " & (NICLAC% - 1) & " " & SQLZ$
            On Error Resume Next
            FLEXCONN.Execute SQLX$
            FLEXCONN.Execute SQLY$
            On Error GoTo 0
         End If
      End If
   End If
End Sub

Private Sub salir_Click()
   Unload Me
End Sub
