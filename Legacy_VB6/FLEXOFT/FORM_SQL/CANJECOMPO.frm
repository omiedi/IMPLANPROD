VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CANJECOMPO 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Canje de componentes por tipo"
   ClientHeight    =   7995
   ClientLeft      =   2970
   ClientTop       =   1395
   ClientWidth     =   8460
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   2280
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   21
      Top             =   840
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton command1 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   7630
      Picture         =   "CANJECOMPO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6980
      Width           =   750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Agregar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   105
      TabIndex        =   12
      Top             =   5040
      Width           =   7335
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
         Height          =   1860
         ItemData        =   "CANJECOMPO.frx":030A
         Left            =   105
         List            =   "CANJECOMPO.frx":030C
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   630
         Width           =   7110
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Código                    Descripción                                                     Cant.  "
         Height          =   270
         Left            =   105
         TabIndex        =   18
         Top             =   315
         Width           =   7110
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Canje de Componentes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   7335
      Begin VB.Label Label7 
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
         Height          =   285
         Left            =   945
         TabIndex        =   15
         Top             =   315
         Width           =   2010
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "N/Serie:"
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
         Left            =   105
         TabIndex        =   14
         Top             =   375
         Width           =   750
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
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
         Left            =   60
         TabIndex        =   11
         Top             =   1365
         Width           =   750
      End
      Begin VB.Label Label2 
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
         Height          =   285
         Left            =   1995
         TabIndex        =   10
         Top             =   1320
         Width           =   5205
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   945
         TabIndex        =   9
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Left            =   105
         TabIndex        =   6
         Top             =   855
         Width           =   750
      End
      Begin VB.Label Label11 
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
         Height          =   285
         Left            =   945
         TabIndex        =   2
         Top             =   795
         Width           =   2010
      End
      Begin VB.Label Label8 
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
         Height          =   285
         Left            =   3045
         TabIndex        =   1
         Top             =   795
         Width           =   4155
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00062A70&
      Height          =   6990
      Left            =   7560
      ScaleHeight     =   6930
      ScaleWidth      =   1320
      TabIndex        =   3
      Top             =   -105
      Width           =   1380
      Begin VB.CommandButton Command2 
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
         Height          =   640
         Left            =   120
         Picture         =   "CANJECOMPO.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   5355
         Width           =   640
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   120
         Picture         =   "CANJECOMPO.frx":0618
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   200
         Width           =   640
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   4
         Top             =   6195
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   5
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "CANJECOMPO.frx":0762
         Top             =   6510
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quitar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   105
      TabIndex        =   7
      Top             =   2100
      Width           =   7335
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
         Height          =   1860
         ItemData        =   "CANJECOMPO.frx":2684
         Left            =   105
         List            =   "CANJECOMPO.frx":2686
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   630
         Width           =   7110
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Código                    Descripción                                    Cant.  Lote"
         Height          =   270
         Left            =   105
         TabIndex        =   19
         Top             =   315
         Width           =   7110
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "CANJECOMPO.frx":2688
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CANJECOMPO.frx":28BC
      TabIndex        =   22
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
End
Attribute VB_Name = "CANJECOMPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
     Command1.Enabled = False
     '
     NUSE$ = TRIM$(Label7)
     CONDI$ = "NumeroSerie = '" & NUSE$ & "'"
     NORFA$ = TRIM$(VALOBADA("TANUMSE", "IDSUBOR", CONDI$, "NOMESS"))
     '
     'VERIFICAR QUE TODOS LOS COMPONENTES TRAZABLES DEL LIST1 (QUITAR)TENGAN LOTE
     'VERIFICAR QUE LA CANTIDAD A QUITAR SEA MENOR O IGUAL A LA CONSUMIDA
     For i& = 1 To List1.ListCount
         'CODIGO DE LIST DE TRAZCOMPO
         CI% = CODINT%(Left$(List1.List(i& - 1), 16))
         CANTI = XVALO(Mid$(List1.List(i& - 1), 44, 10))
         '
         'VALIDACION QUE SE HAYA CONSUMIDO
         CONDI$ = "DOPRICOR = '" & NORFA$ & "'  and codimovi = 'CF' AND COD_INTE = " & CI%
         ORFA$ = TRIM$(Mid$(List1.List(i& - 1), 100, 12))
         If ORFA$ <> "" Then CONDI$ = "DOPRICOR = '" & ORFA$ & "'  and codimovi = 'CF' AND COD_INTE = " & CI%
         CantCons = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID - CANTINGRE)", CONDI$, "NOMESS"))
         'If CantCons < CANTI Then
         '     Call COMUNI("Imposible Continuar!!\La Cantidad a Quitar del Código: " & CODEXT(CI%) & "\Es Superior a la Consumida.")
         '     GoTo 99
         'End If
         '
         'VALIDACION DE LOTE
         ESTRAZ% = ESTRAZABLE%(CI%)
         If ESTRAZ% = 1 Then
            LOTE$ = TRIM$(Mid$(List1.List(i& - 1), 54, 12))
            If LOTE$ = "" Then
              Call COMUNI("Imposible Continuar!!\Falta Lote de Consumo Para el Código: " & CODEXT(CI%))
              GoTo 99
            End If
            ORFA$ = TRIM$(Mid$(List1.List(i& - 1), 100, 12))
            If ORFA$ <> "" Then
                  CONDI$ = "CODICONJ = " & CODINT%(Label11) & " AND TIPOCANJE  = " & XVALO(Label1) & " AND AGREGAR > 0.05"
                  Call CARGALISEL("COMPQUITA", "CANJPREDEF", "CODIELEM/C16+MASTER/A36", CONDI$)
                  Codigo$ = TRIM$(Mid$(List1.List(i& - 1), 1, 16))
                  Call SELECHO("COMPQUITA/APLICACION COMPONENTE DE SUBCONJUNTO: " & Codigo$)
                  ASIGNACODI$ = TRIM$(VALACT1$("COMPQUITA"))
                  If ASIGNACODI$ <> "" Then
                     For II% = 0 To List2.ListCount - 1
                         Fila$ = List2.List(II%)
                         If TRIM$(Mid$(Fila$, 1, 16)) = ASIGNACODI$ Then
                            Call REPLA(Fila$, ORFA$, 100, 12)
                            List2.List(II%) = Fila$
                         End If
                     Next II%
                  End If
            End If
         End If
      Next i&

     'CARGO EL TRAZCOMPO
     For i& = 1 To List2.ListCount
         'CODIGO DE LIST DE TRAZCOMPO
         CI% = CODINT%(Left$(List2.List(i& - 1), 16))
         CANTI = XVALO(Mid$(List2.List(i& - 1), 54, 10))
         ESTRAZ% = ESTRAZABLE%(CI%)
         If ESTRAZ% = 1 Then
            NORDE& = i&
            TX$ = ""
            Call REPLA(TX$, CODEXT$(CI%), 1, 16)
            Call REPLA(TX$, DESCRIT$(CI%), 17, 32)
            Call REPLA(TX$, FORMATNUM(CANTI, "F10.1"), 50, 11)
            Call REPLA(TX$, FORMATNUM(NORDE&, "F6.0"), 80, 6)
            TRAZCOMPO.List2.AddItem TX$
         End If
     Next i&
     '
     'Presentar formulario para asignar lotes
     'LEO EL TRAZCOMPO
     If TRAZCOMPO.List2.ListCount > 0 Then
50    TRAZCOMPO.Show 1
      If TRAZCOMPO.List2.ListCount < 1 Then
           TRAZCOMPO.LIMPIAR
           GoTo 99
      Else
        'PRIMERO VALIDO QUE TODOS TENGAN LOTE
        For i& = 1 To TRAZCOMPO.List2.ListCount
           LOTE$ = TRIM$(Mid$(TRAZCOMPO.List2.List(i& - 1), 62, 16))
           CI% = CODINT%(Mid$(TRAZCOMPO.List2.List(i& - 1), 1, 16))
           If LOTE$ = "" Then
              Call COMUNI("Imposible Continuar!!\Falta Lote Para el Código: " & CODEXT(CI%))
              GoTo 50 ' PRESENTA NUEVAMENTE EL TRAZCOMPO
           End If
        Next i&
        'FIN VALIDACION DE LOTE
        '
        For i& = 1 To TRAZCOMPO.List2.ListCount
           'ACTUALIZO EL LIST  AGREGANDO EL LOTE
           LOTE$ = TRIM$(Mid$(TRAZCOMPO.List2.List(i& - 1), 62, 16))
           NUMEORD% = XVALO(Mid$(TRAZCOMPO.List2.List(i& - 1), 80, 6))
           '
           Fila$ = List2.List(NUMEORD% - 1)
           Call REPLA(Fila$, LOTE$, 80, 12)
           List2.List(NUMEORD% - 1) = Fila$
        Next i&
      End If
     End If
     '
     'VALIDAR QUE TODOS TENGAN STOCKS LOS ITEMS NO TRAZABLES CON LA RUTINA PUEDEMOVI
     DEPOCEN% = XVALO(CONTROL$("", "DEPOCEN"))
     CAITMO% = 0
     For i& = 1 To List2.ListCount
         CIXI% = CODINT%(Left$(List2.List(i& - 1), 16))
         CANTI = XVALO(Mid$(List2.List(i& - 1), 54, 10))
         If ESTRAZABLE%(CIXI%) < 1 Then
            If CANTI > 0.05 Then
               CAITMO% = CAITMO% + 1
               CIMOVIM%(CAITMO%) = CIXI%
               CAMOVIM(CAITMO%) = (-1) * CANTI
            End If
         End If
     Next i&
     If PUEDEMOVI%(DEPOCEN%) < 1 Then GoTo 99
     'FIN VALIDACION DE STOCKS DE ITEMS NO TRAZABLES
     '
     'GRABACION
     '*********
     '
     'GENERAR MOVIMIENTO DE STOCKS
     '****************************
     CMOV$ = "CF"
     HOII% = HOY(HOS$)
     FF% = HOII%
     '
     NUORIT% = 0
     UNI# = 0
     MONEII% = 0
     NC% = 0
     REFE$ = "Canje de Componentes"
     '
     VUELTA% = 0
     CUERPO$ = "Canje de Componentes" & vbCrLf + vbCrLf
     '
     For i& = 1 To List1.ListCount
         CIXI% = CODINT%(Left$(List1.List(i& - 1), 16))
         CANTI = XVALO(Mid$(List1.List(i& - 1), 44, 10))
         LOTE$ = TRIM$(Mid$(List1.List(i& - 1), 54, 12))
         ORFA$ = TRIM$(Mid$(List1.List(i& - 1), 100, 12))
         DOPRI1$ = NORFA$
         DOSEC1$ = "PF" + TRIM$(Mid$(NORFA$, 3))
         If ORFA$ <> "" Then
            DOPRI1$ = ORFA$
            DOSEC1$ = "PF" + TRIM$(Mid$(ORFA$, 3))
         End If
         
         'MOVIMIENTOS DE ENTRADA
         Call MOVISTOK(FF%, CIXI%, LOTE$, CMOV$, DOPRI1$, DOPRI1$, DOSEC1$, DOSEC1$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPOCEN%, CANTI)
         '
         'ESCRIBO EN HISTORIAL DE NUMEROS DE SERIE
         If VUELTA% < 1 Then
            CUERPO$ = CUERPO$ + "Se han Quitado los siguientes Componentes: " & vbCrLf + vbCrLf
            VUELTA% = 1
         End If
         TXX1$ = FORMATNUM$(CANTI, "F9.1") & " " & Unimed$(CIXI%)
         TEXTO1$ = ""
         Call REPLA(TEXTO1$, AJUSTI(TXX1$, 12), 1, 12)
         Call REPLA(TEXTO1$, CODEXT$(CIXI%), 14, 20)
         If LOTE$ <> "" Then
           Call REPLA(TEXTO1$, "(" & LOTE$ & ")", 36, 14)
         End If
         CUERPO$ = CUERPO$ + TEXTO1$ + vbCrLf
         '
     Next i&
     CUERPO$ = CUERPO$ & vbCrLf
     '
     CMOV$ = "TR"
     NRO& = ProNroComprobante(CMOV$)
     DOPRI$ = "TR." + TRIM$(str$(NRO&))
     DEPODESTI% = CONTROL("", "DEPODEST")
     VUELTA% = 0
     For i& = 1 To List2.ListCount
         CIXI% = CODINT%(Left$(List2.List(i& - 1), 16))
         CANTI = XVALO(Mid$(List2.List(i& - 1), 54, 10))
         LOTE$ = TRIM$(Mid$(List2.List(i& - 1), 80, 12))
         ORFA$ = TRIM$(Mid$(List2.List(i& - 1), 100, 12))
         DOSEC$ = NORFA$
         If ORFA$ <> "" Then DOSEC$ = ORFA$
         '
         'MOVIMIENTOS DE SALIDA
         Call MOVISTOK(FF%, CIXI%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPOCEN%, (-1) * CANTI)
         'MOVIMIENTOS DE ENTRADA
         Call MOVISTOK(FF%, CIXI%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPODESTI%, CANTI)
         'CONSUMO
         DOPRI1$ = DOSEC$
         DOSEC1$ = "PF" + TRIM$(Mid$(DOSEC$, 3))
         Call MOVISTOK(FF%, CIXI%, LOTE$, "CF", DOPRI1$, DOPRI1$, DOSEC1$, DOSEC1$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPODESTI%, (-1) * CANTI)
         '
         'ESCRIBO EN HISTORIAL DE NUMEROS DE SERIE
         If VUELTA% < 1 Then
            CUERPO$ = CUERPO$ + "Se han Agregado los siguientes Componentes: " & vbCrLf + vbCrLf
            VUELTA% = 1
         End If
         TXX1$ = FORMATNUM$(CANTI, "F9.1") & " " & Unimed$(CIXI%)
         TEXTO1$ = ""
         Call REPLA(TEXTO1$, AJUSTI(TXX1$, 12), 1, 12)
         Call REPLA(TEXTO1$, CODEXT$(CIXI%), 14, 20)
         If LOTE$ <> "" Then
           Call REPLA(TEXTO1$, "(" & LOTE$ & ")", 36, 14)
         End If
         CUERPO$ = CUERPO$ + TEXTO1$ + vbCrLf
         '
     Next i&
     Call ANOTANUSE(NUSE$, CUERPO$)
     '
     OPX% = COMALTER("¿Desea Cambiar la Tensión del Equipo?\\No, Continuar\Cambiar a 220\Cambiar a 110")
     If OPX% < 2 Or OPX% > 3 Then GoTo 98
     TENSION% = 220: If OPX% = 3 Then TENSION% = 110
     Call ACTUREGISTRO("TANUMSE", "TEN_ENSA", "NUMEROSERIE='" & TRIM$(Label7) & "'", TENSION%, REAF&, "NOMESS")
     Call CONECRUN("TRANSE07/IMPRIETI" + TRIM$(Label7), "Impresion de Etiquetas")
     '
98   Screen.MousePointer = 1
     Call COMUNI("Proceso Finalizado")
     '
     Call Command3_Click ' CIERRA EL FRM
99   Command1.Enabled = True
End Sub


Private Sub Command2_Click()
   Command2.Enabled = False
   FIKARDEX09.Show 1
   Command2.Enabled = True
End Sub

Private Sub Command3_Click()
    Call CIERRARCH("*.*")
    Unload Me
End Sub


Private Sub Form_Load()
  Call APLICACIONSKINS(Me, Picture1)
End Sub

Private Sub label1_Change()
   TIPOCANJ% = XVALO(Label1)
   Label2 = TRIM$(ECOARCH$("TATICANJ", MKI$(TIPOCANJ%)))
End Sub


Private Sub Label11_Change()
   CI1% = CODINT%(Label11)
   Label8 = DESCRIT0$(CI1%)

End Sub

Private Sub Label7_Change()
   
   NUSE$ = TRIM$(Label7)
   CONDI$ = "NumeroSerie = '" & NUSE$ & "'"
   CI% = XVALO(VALOBADA("TANUMSE", "COD_INTE", CONDI$, "NOMESS"))
   NORFA$ = TRIM$(VALOBADA("TANUMSE", "IDSUBOR", CONDI$, "NOMESS"))
   Label11 = CODEXT$(CI%)
   TIPOCANJ% = XVALO(Label1)
   
   SQLX$ = "SELECT CODIELEM, QUITAR, AGREGAR  FROM CANJPREDEF WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE CODICONJ = " & CI%
   SQLX$ = SQLX$ + " AND TIPOCANJE  = " & TIPOCANJ%
   Set CANJETMP = READSET(SQLX$)
   '
   With CANJETMP
      Do While Not .EOF
       CI% = XVALO(!CODIELEM)
       DESCRIPX$ = DESCRIT0$(CI%)
       QUIT = XVALO(!QUITAR)
       AGREG = XVALO(!AGREGAR)
       '
       LOTE$ = "": ORFA$ = ""
       If ESTRAZABLE%(CI%) = 1 Then LOTE$ = BUSCALOTE$(CI%, NORFA$, QUIT, ORFA$)
       '
       L1$ = Space(128)
       Call REPLA(L1$, CODEXT$(CI%), 1, 16)
       If QUIT > 0.05 Then
         Call REPLA(L1$, DESCRIPX$, 18, 24)
         Call REPLA(L1$, FORMATNUM$(QUIT, "F9.1"), 44, 10)
         Call REPLA(L1$, LOTE$, 54, 12)
         Call REPLA(L1$, ORFA$, 100, 12)
         List1.AddItem L1$
       ElseIf AGREG > 0.05 Then
         Call REPLA(L1$, DESCRIPX$, 18, 35)
         Call REPLA(L1$, FORMATNUM$(AGREG, "F9.1"), 54, 10)
         List2.AddItem L1$
       End If
      .MoveNext
      Loop
   End With
   CANJETMP.Close
   Set CANJETMP = Nothing
   
   '
End Sub

Function BUSCALOTE$(CI%, NORFA$, CANTI, ORFA$)
   '
   'BUSCO EL LOTE ASUMIENTO QUE EL QUE SE CAMBIA ES COMPONENTE INMEDIATO
   CONDI$ = "COD_INTE = " & CI% & " AND CODIMOVI = 'CF' AND DOPRILAR = '" & NORFA$ & "' AND (CANTSALID - CANTINGRE) >= " & CANTI
   LOTE$ = TRIM$(VALOBADA("MOVISTO", "IDENLOTE", CONDI$, "NOMESS"))
   '
   If TRIM$(LOTE$) = "" Then
      'BUSCO SI EL CAMBIO ES UN COMPONENTE DE ALGUN SUBCONJUNTO
      CIXI% = CODINT%(Label11)
      Call AREXPLO(CIXI%, 1, 1)
      ORFA$ = ""
      For k% = 1 To CAIT%
        CONDI$ = "CODICONJ = " & CIJ%(k%) & " AND CODIELEM=" & CI%
        'BUSCO A QUIEN PERTENECE
        If CantRegistros("FORMULAS", CONDI$, "NOMESS") > 0 Then
           'BUSCO EL LOTE DEL SUBCONJUNTO
           CONDI$ = "COD_INTE = " & CIJ%(k%) & " AND CODIMOVI = 'CF' AND DOPRILAR = '" & NORFA$ & "'"
           ORFA$ = TRIM$(VALOBADA("MOVISTO", "IDENLOTE", CONDI$, "NOMESS"))
           'BUSCO EL LOTE DEL COMPONENTE DEL SUBCONJUNTO
           CONDI$ = "COD_INTE = " & CI% & " AND CODIMOVI = 'CF' AND DOPRILAR = '" & ORFA$ & "' AND (CANTSALID - CANTINGRE) >= " & CANTI
           LOTE$ = TRIM$(VALOBADA("MOVISTO", "IDENLOTE", CONDI$, "NOMESS"))
           '
           GoTo 99
        End If
      Next k%
   End If
    
99 BUSCALOTE$ = LOTE$
   
End Function

