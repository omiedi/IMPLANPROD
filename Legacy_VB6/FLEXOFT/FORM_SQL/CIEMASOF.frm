VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form CIEMASOF 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cierre/Anulación Automático de Ordenes de Fabriación"
   ClientHeight    =   8745
   ClientLeft      =   75
   ClientTop       =   255
   ClientWidth     =   15315
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8745
   ScaleWidth      =   15315
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   9000
      TabIndex        =   25
      Top             =   7920
      Width           =   5415
      Begin VB.TextBox Text23 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   765
      End
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "Cant. Filas Selecc.:"
         ToolTipText     =   "Cantidad de Item Seleccionados"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "Cantidad FilasTotal:"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text19 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   16
      Top             =   7920
      Width           =   9015
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "Depósito Consumo"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Depósito Cierre"
         ToolTipText     =   "Cantidad de Item Seleccionados"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6435
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   390
      End
      Begin VB.CommandButton Command6 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         TabIndex        =   20
         Top             =   240
         Width           =   175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6240
         TabIndex        =   19
         Top             =   240
         Width           =   175
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6795
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   2085
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   2085
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7440
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   8880
      Left            =   14400
      ScaleHeight     =   8820
      ScaleWidth      =   870
      TabIndex        =   4
      Top             =   0
      Width           =   930
      Begin VB.CommandButton BTNMARCACIERREOF 
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
         Height          =   690
         Left            =   120
         Picture         =   "CIEMASOF.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Marca de Cierre de o.f. Seleccionadas"
         Top             =   6720
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   0
         ScaleHeight     =   75
         ScaleWidth      =   825
         TabIndex        =   12
         Top             =   7680
         Width           =   885
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Ref."
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
         Left            =   70
         TabIndex        =   11
         Top             =   5640
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha"
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
         Left            =   70
         Picture         =   "CIEMASOF.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Selección de Rango de Fechas"
         Top             =   2040
         Width           =   765
      End
      Begin VB.CommandButton Command43 
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
         Left            =   70
         Picture         =   "CIEMASOF.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Revertir Selección"
         Top             =   1080
         Width           =   765
      End
      Begin VB.CommandButton COMMAND3 
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
         Height          =   690
         Left            =   70
         Picture         =   "CIEMASOF.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Generar Cierre/Anulación de Ordenes de Fabricación Seleccionadas"
         Top             =   6960
         Width           =   765
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
         Left            =   70
         Picture         =   "CIEMASOF.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   120
         Width           =   765
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "CIEMASOF.frx":0BB2
         Top             =   7800
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   7935
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   14400
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   7560
         Width           =   5200
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   7560
         Width           =   8760
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   7365
         Left            =   0
         TabIndex        =   14
         ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar"
         Top             =   120
         Width           =   14355
         _ExtentX        =   25321
         _ExtentY        =   12991
         _Version        =   393216
         Cols            =   4
         BackColor       =   15724527
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   4920
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   9720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "CIEMASOF.frx":2AD4
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CIEMASOF.frx":2D08
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "Archivo"
      Visible         =   0   'False
      Begin VB.Menu mnucopiar 
         Caption         =   "Copiar Celda Activa"
         Visible         =   0   'False
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuarch 
      Caption         =   "Archivo"
      Begin VB.Menu grabacierre 
         Caption         =   "Generar Cierre/Anulación de O.Fabricación"
      End
      Begin VB.Menu exportaexcel 
         Caption         =   "Exportar a Excel"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "Edición"
      Begin VB.Menu mnuRevertirSeleccion 
         Caption         =   "Revertir Selección"
      End
   End
   Begin VB.Menu Ver 
      Caption         =   "Ver"
      Begin VB.Menu SelRangoFechas 
         Caption         =   "Ordenes de Fabricación Pendientes de Cierre"
      End
   End
End
Attribute VB_Name = "CIEMASOF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GLOBALFRM_strContenidoCelda$

Sub ANULAORDENESPENDIENTES(CI1%, NORFA$, CANTI, FECHACIERRE%, NROOFL)
     Call ABREORFAB
     Call ABRESERVA
     SQLX$ = "SELECT Statuorf, IdSubOr FROM OrdeFabr "
     SQLX$ = SQLX$ + "WHERE NumeOrFa = " & NROOFL & " "
     SQLX$ = SQLX$ + " and IDSUBOR= '" & NORFA$ & "'"
     Dim ESTADORDEN As ADODB.Recordset
     Set ESTADORDEN = New ADODB.Recordset
     ESTADORDEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     'With OrdeFabr
     With ESTADORDEN
          '.Edit
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
             Do While Not .EOF
              !StatuOrf = 9
              .Update
              IDSUB$ = SAFETEXT$(!idsubor)
              '
              SQLX$ = "SELECT * FROM Reserva "
              SQLX$ = SQLX$ + "WHERE IDSUBOR= '" & IDSUB$ & "' "
              Dim CANTIRES As ADODB.Recordset
              Set CANTIRES = New ADODB.Recordset
1420          On Error Resume Next
              CANTIRES.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
              '
              If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                On Error GoTo 0
                Call CapturaErrorOpen
                GoTo 1420
              End If
              On Error GoTo 0
              '
              With CANTIRES
                  On Error Resume Next
                  .MoveFirst
                  If Err < 1 Then
                    Do While Not .EOF
                      !CANTRES = XVALO(!CANTCONS)
                      !Saldo_Entre = XVALO(!CANTRES) - XVALO(!CANTCONS)
                      .Update
                      '
                      'ACTUALIZO EL SALDO DEL LOTE
                      CI% = XVALO(!cod_inte)
                      LOTE$ = SAFETEXT$(!idenlote)
                      Call ACSALOTE(CI%, LOTE$, "NOMESS")
                      '
                      .MoveNext
                    Loop
                  End If
                  On Error GoTo 0
              End With
              Set CANTIRES = Nothing
              'ANULACION DE OPERACIONES DE LA O-FABRICACION
              SQLX$ = "DELETE FROM OperFab "
              SQLX$ = SQLX$ + "WHERE IdSubOr = '" & IDSUB$ & "' "
              FLEXCONN.Execute (FILTSQL$(SQLX$))
              .MoveNext
             Loop
          End If
          On Error GoTo 0
     End With
     Set ESTADORDEN = Nothing
     
1500  If EXISTE%("TRAZAB07.EXE") > 0 Or EXISTE%("TRAZAB08.EXE") > 0 Then
        Call ANULARES(NORFA$)
     End If
     '
'     'BANDERA PARA BORRAR LOS REGISTROS EN -1 DE LOS EQUIPOS
'     If BORRARNUSE% > 0 Then
'        CONDI$ = "DEPONUMSE < 0 AND IDSUBOR = '" & NORFA$ & "' And Cod_Inte = " & CIXI%
'        Call BORRAREGISTROS("TANUMSE", CONDI$, REGA&, "NOMESS")
'     End If
     '
     'SI LA ORDEN DE COMPRA TIENE ASIGNADO EL NUMERO DE LA ORDEN A ANULAR LA ANULA TAMBIEN.
     CONDI$ = "Doc_Orig =  '" & NORFA$ & "' AND N_OrServ  > 0 "
     Call ACTUREGISTRO("OCOMDETA", "STATUS", CONDI$, 9, REGAF&, "NOMESS")
     '
     jj& = 0
     For U& = 1 To ULTREG&("!ECORANU")
       ZZ$ = REGLEIDO$("!ECORANU", U&)
       If TRIM$(Left$(ZZ$, 12)) <> TRIM$(NORFA$) Then
         jj& = jj& + 1
         Call GRAREG("!ECORANU", ZZ$, jj&)
       End If
     Next U&
     Call SETULTREG("!ECORANU", jj&)
     '
     Call CIERRARCH("*.*")
End Sub

Private Sub BTNMARCACIERREOF_Click()
   BTNMARCACIERREOF.Enabled = False
   Screen.MousePointer = 1
'   If TRIM$(Text4) = "" Then
'      Call COMUNI("Falta Depósito de Cierre (Ingreso)")
'      GoTo 99
'   End If
'   If TRIM$(Text3) = "" Then
'      Call COMUNI("Falta Depósito de Consumo")
'      GoTo 99
'   End If
'   '
'   OPX1% = COMALTER%("Confirme Transacción a Realizar\\Cancelar\Cerrar Ordenes de Fabricación Seleccionadas\Anular Ordenes de Fabricación Seleccionadas")
'   If OPX1% < 2 Then GoTo 99
'   If OPX1% = 3 Then GoTo 10
'   OPX% = COMALTER%("Confirmar El Cierre de La/s " & Text23 & " Ordenes de Fabricación Seleccionadas\\Cancelar\Realizar Cierre")
'   If OPX% < 2 Then GoTo 99
   
10 MSF1.Height = 7365
   '
   
   For i& = 1 To MSF1.Rows - 1
       If MSF1.TextMatrix(i&, 0) = "*" Then
         CI1% = XVALO(MSF1.TextMatrix(i&, 8))
         NORFA$ = TRIM$(MSF1.TextMatrix(i&, 1))
'         CANTI = XVALO(MSF1.TextMatrix(i&, 7))
         FECHACIERRE% = HOY(HO$)
         NumeOrFaX& = XVALO(MSF1.TextMatrix(i&, 9))
         Call COLOREAR_GRILLA_SOLAFILA(Me.MSF1, &HFF0000, i&)
         Call COLOREAR_TEXTO_SOLO_UNA_AFILA(Me.MSF1, vbWhite, i&)
         ' ESTO ES PARA POSICIONARSE AL FINAL DE LA LISTA
         MSF1.row = i&
         TPXI& = MSF1.row - 24: If TPXI& < 1 Then TPXI& = 1
         On Error Resume Next
         MSF1.TopRow = TPXI&
         On Error GoTo 0
         Call MarcaCierreOFBorraReserva(NORFA$)
          'REGISTRA EN EL ANOTADOR DE LA OF LA ASIGNACION DEL ARMADOR
         TTXYZ$ = Space$(76)
         HOII% = HOY(HOS$)
         Call REPLA(TTXYZ$, HOS$, 1, 8)
         Call REPLA(TTXYZ$, "Marca de Cierre Manual (Ciemasof)", 11, 56)
         Call REPLA(TTXYZ$, USERTER$, 69, 24)
         Call AGREGA_ANOTAORFAB(NORFA$, TTXYZ$)
       End If
   Next i&
   '
   Call COMUNI("Proceso Finalizado")
   '
   Unload Me
99 BTNMARCACIERREOF.Enabled = True
   Screen.MousePointer = 1

End Sub
Sub MarcaCierreOFBorraReserva(NORFAX$)
          EPAC$ = UCase$(TRIM$(EMPREAC$))

          If NORFAX$ <> "" Then
              CONDI$ = "IdSubOr = '" & NORFAX$ & "'"
              FECHAHOY = FETEXCOR1$(HOY(HO$))
              SQLX$ = "UPDATE ORDEFABR SET  StatuOrf = 3 ,  FechCierr = '" & FECHAHOY & "'"
              SQLX$ = SQLX$ + " WHERE " & CONDI$
              FLEXCONN.Execute (SQLX$)
              If InStr(EPAC$, "CONTEPLAST") < 1 Or InStr(EPAC$, "CAMBACERES") < 1 Then 'en conteplast no borro las operaciones
                 ' Cancela Operaciones Programadas   'por las dudas para que no queden huerfanas
                 Call BORRAREGISTROS("OperFab", CONDI$, RAFE&, "NOMESS")
              End If
              ' Cancela Reservas de Componentes
              Call BORRAREGISTROS("Reserva", CONDI$, RAFE&, "NOMESS")
              ' Cancela Orden en Indice de Ordenes Pendientes
              NORFAX$ = AJUSTI$(NORFAX$, 12)
              Call FILTERDELETE("ECORDEP", 1, NORFAX$)
              Call FRESHECHO("ECORDEP")
          End If
          Call CIERRARCH("*.*")
          Call LIBARCH("ECORDEP")
          Call CIERRARCH("*.*")
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   '
   If TRIM$(Control$("STOCKS", "CONEGATI")) = "SI" Then
     Call COMUNI("Atención!!! Imposible Realizar Proceso \Control de Negativos Activado.")
     GoTo 99
   End If
   
   If Control("VANEDECE", "DEPOCEN") = "SI" Then
     Call COMUNI("Atención!!! Imposible Realizar Proceso \Control de Negativos de O.Fabricación Activado ")
     GoTo 99
   End If

   'VER SI GENERANDO LA TABLA QUE GUARDE EL NPED Y EL CODIGO PARA OBTENER LA CANTIDAD YA CERRADA DE ESE PEDIDO
   '*******VER MEJOR DE TOMAR LA CANTIDAD APROBADA DE LAS OF  CON UN SUM(CANAPAROBADA) DE LA OF SEGUN CODIGO Y N.PED.
   'SI CUANDO VALIDA 'If SALDO_POR_CERRAR > 0 Then' SE CONSIDERA TAMBIEN LA CANTIDAD YA CERRADA PARA ESE ITEM Y N PED.
   'CAMPOS NPED, CODIGO, OF, CANTIDAD CERRADA.
   '
   'ESTA FALLANDO EN EL CASO QUE EN EL MISMO PEDIDO HAY MAS DE UN ITEM REPETIDO SI SE REALIZA UN CIERRE POR 1 ITEM DE LOS REPETIDOS
   'CUANDO SE VUELVE A LLAMAR AL FRM SIN FACTURAR NUEVAMENTE APARECE LA CANTIDAD FACTURADA EN EL SEGUNDO ITEM.

   'ESTO FUNCIONA SOLO EN LOS CASOS QUE TODOS LOS PEDIDOS GENEREN ORDEN DE FABRICACION.
   '************************************************************************************
   Call DESHASFECHA(Des%, Has%, PERIO$)
   If PERIO$ = "" Then Exit Sub
   '
   Desde = FETEXCOR1$(Des%) ' SE MODIFICO EL FORMATO FECHA DE NUMERICO A FECHA CORTA
   Hasta = FETEXCOR1$(Has%)
   '
   Call ABRECONEXION
   '
   MSF1.Rows = 1
   Screen.MousePointer = 11
   '
   Dim rs As New ADODB.Recordset
   
   SQLX$ = "SELECT * FROM ORDEFABR WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE STATUORF < 3 "
   SQLX$ = SQLX$ + " and FechLan >= '" & Desde & "'"
   SQLX$ = SQLX$ + " and FechLan <= '" & Hasta & "'"
   SQLX$ = SQLX$ + " ORDER BY  IDSUBOR "
   'MSF1.Redraw = False
   Set rs = FLEXCONN.Execute(SQLX$) ' SIN FILTSQL PARA QUE TRAIGA TODO
   i& = 0
   With rs
     Do While Not .EOF
        'NO LO HICE CON VALOBADA POR QUE ERA MAS LENTO
        CODIEXT$ = SAFETEXT$(!Cod_Exte)
        DESQRI$ = SAFETEXT$(!Descrip)
        FECHALANZ$ = FECHATEX$(CVINT(!FechLan))
        NORFA$ = SAFETEXT$(!idsubor)
        NROPED& = XVALO(!NROPED)
        CI1% = XVALO(!cod_inte)
        CANPRO = XVALO(!CANAPRO)
        CANPRY = XVALO(!CANPROY)
        NROOFL = XVALO(!NUMEORFA)
        SALDO_POR_CERRAR = XVALO(CANPRY) - XVALO(CANPRO)
        If SALDO_POR_CERRAR < 0 Then SALDO_POR_CERRAR = 0
        If SALDO_POR_CERRAR > 0 Then
            i& = i& + 1
            MSF1.Rows = i& + 1
            MSF1.TextMatrix(i&, 1) = NORFA$
            MSF1.TextMatrix(i&, 2) = CODIEXT$
            MSF1.TextMatrix(i&, 3) = DESQRI$
            MSF1.TextMatrix(i&, 4) = FECHALANZ$
            
            MSF1.TextMatrix(i&, 5) = FORMATNUM$(CANPRY, "F9.1")
            MSF1.TextMatrix(i&, 6) = FORMATNUM$(CANPRO, "F9.1")
            MSF1.TextMatrix(i&, 7) = FORMATNUM$(SALDO_POR_CERRAR, "F9.1")
            MSF1.TextMatrix(i&, 8) = CI1%
            MSF1.TextMatrix(i&, 9) = NROOFL
            MSF1.TextMatrix(i&, 10) = NROPED&
         End If
30     .MoveNext

     Loop
   End With
   MSF1.Redraw = True
   If i& = 0 Then Call COMUNI("No se Han Encontrado Ordenes de Fabricación Sin Cerrar en el Rango Seleccionado")
   Call ACTUTOTALES
99 Command1.Enabled = True
   Screen.MousePointer = 1
End Sub
Function YACONSIDERADO(NROP&, CI11%)
    For i& = 1 To MSF1.Rows - 1
      If NROP& = XVALO(MSF1.TextMatrix(i&, 5)) And CI11% = XVALO(MSF1.TextMatrix(i&, 10)) Then
          YACONSIDERADO = YACONSIDERADO + MSF1.TextMatrix(i&, 8)
      End If
    Next i&
End Function

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Screen.MousePointer = 1
   If TRIM$(Text4) = "" Then
      Call COMUNI("Falta Depósito de Cierre (Ingreso)")
      GoTo 99
   End If
   If TRIM$(Text3) = "" Then
      Call COMUNI("Falta Depósito de Consumo")
      GoTo 99
   End If
   '
   OPX1% = COMALTER%("Confirme Transacción a Realizar\\Cancelar\Cerrar Ordenes de Fabricación Seleccionadas\Anular Ordenes de Fabricación Seleccionadas")
   If OPX1% < 2 Then GoTo 99
'   If OPX1% = 3 Then GoTo 10
'   OPX% = COMALTER%("Confirmar El Cierre de La/s " & Text23 & " Ordenes de Fabricación Seleccionadas\\Cancelar\Realizar Cierre")
'   If OPX% < 2 Then GoTo 99
   
10 MSF1.Height = 7800
   '
   
   For i& = 1 To MSF1.Rows - 1
       If MSF1.TextMatrix(i&, 0) = "*" Then
         CI1% = XVALO(MSF1.TextMatrix(i&, 8))
         NORFA$ = TRIM$(MSF1.TextMatrix(i&, 1))
         CANTI = XVALO(MSF1.TextMatrix(i&, 7))
         FECHACIERRE% = HOY(HO$)
         NumeOrFaX& = XVALO(MSF1.TextMatrix(i&, 9))
         Call COLOREAR_GRILLA_SOLAFILA(Me.MSF1, &HFF0000, i&)
         Call COLOREAR_TEXTO_SOLO_UNA_AFILA(Me.MSF1, vbWhite, i&)
         ' ESTO ES PARA POSICIONARSE AL FINAL DE LA LISTA
         MSF1.row = i&
         TPXI& = MSF1.row - 24: If TPXI& < 1 Then TPXI& = 1
         On Error Resume Next
         MSF1.TopRow = TPXI&
         On Error GoTo 0
        
         If CANTI > 0 Then
           If OPX1% = 3 Then
             Call ANULAORDENESPENDIENTES(CI1%, NORFA$, CANTI, FECHACIERRE%, NumeOrFaX&)
           Else
             Call CIERRORFACTU(CI1%, NORFA$, CANTI, FECHACIERRE%)
           End If
         End If

       End If
   Next i&
   '
   Call COMUNI("Proceso Finalizado")
   '
   Unload Me
99 Command3.Enabled = True
   Screen.MousePointer = 1
End Sub
Sub CIERRORFACTU(CI1%, NORFA$, CANTI, FECHACIERRE%)
    'SE USA PARA JYM EN EL PROCESO DE CIERRE AUTOMATICO DE ITEMS CON O.F. PENDIENTES YA FACTURADOS Y EN AMBEST CIERRE MASIVO POR RANGO DE FECHAS
    Text5 = "Resolviendo Cierre de Orden de Fabricación de Item: " & CODEXT$(CI1%) & " - " & DESCRIT0$(CI1%)
    Call ENRAPRODUCTO_TERMINADO(CI1%, NORFA$, CANTI, FECHACIERRE%, XVALO(Text2), "Fabricación Propia - Cierre Masivo")
    '
    Text7 = "Verificando Componentes Reservados"
    Text7.Refresh
    CAN# = CANTI
    Call CARCOMPO(CODEXT$(CI1%), NORFA$, CAN#, 0, 1)
    Text7 = "Resolviendo Consumo de Componentes"
    Text7.Refresh

    Call CONSUCOMPONETESXOF(CI1%, NORFA$, CANTI, FECHACIERRE%, XVALO(Text1))
    '
    'ACTUALIZA LA RESERVA
    FIN& = ULTREG&("!MATEROF")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      ZZ$ = REGLEIDO$("!MATEROF", U&)
      CICOMPO% = CVI(Left$(ZZ$, 2))
      CANCONSUM = CVS(Mid$(ZZ$, 5, 4))
      If TRIM$(UCase$(Unimed$(CICOMPO%))) = "GK" Then CANCONSUM = CANCONSUM / 1000
      '
      FECHA% = FECHACIERRE%
      Text7 = "Actualizando Reservas "
      Text7.Refresh
      Call ACTURESERVA(CICOMPO%, NORFA$, CANCONSUM)
    Next U&
    '
    Text7 = "Escribiendo Anotador de Ordenes de Fabricación"
    Text7.Refresh
    Call ANOTAENORFA_FACTURADO(CI1%, NORFA$, CANTI, FECHACIERRE%, "Cierre de O/Fabricación (Proc.Másivo)")
    Call ACTUALIZAORFA(CI1%, NORFA$, CANTI, FECHACIERRE%)
    
End Sub

Sub ACUCONRESFAC(CIXI%, IDLO$, NORFAX$)
    '
    CMOV$ = "CF"
    DOPRI$ = NORFAX$
    DOSEC$ = "PF-" + TRIM$(Mid$(NORFAX$, 4))
    '
    ' BUSCAR SI SE PUEDE ACUMULAR A UN MOVIMIENTO YA REGISTRADO
    CONDI$ = "CODIMOVI = '" & CMOV$ & "' "
    CONDI$ = CONDI$ & " AND COD_INTE = " & CIXI%
    CONDI$ = CONDI$ & " AND IDENLOTE = '" & IDLO$ & "' "
    CONDI$ = CONDI$ & " AND DOPRICOR = '" & DOPRI$ & "'"
    CONDI$ = CONDI$ & " AND DOSECCOR = '" & DOSEC$ & "'"
    '
    ACMOV = VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS")
    '
    CONDIX$ = " IdSubOr = '" & NORFAX$ & "' AND Cod_Inte = " & CIXI% & " AND IDENLOTE = '" & IDLO$ & "'"
    SQLX$ = "SELECT CANTCONS, SALDO_ENTRE, CANTRES FROM RESERVA WHERE " & CONDIX$
    Dim ReserTemp As ADODB.Recordset
    Set ReserTemp = New ADODB.Recordset
    ReserTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    On Error Resume Next
    With ReserTemp
        .MoveFirst
        If Err < 1 Then
           On Error GoTo 0
           TORES = XVALO(!CANTRES)
           TCONSU = ACMOV: If TCONSU > TORES Then TCONSU = TORES
           SALENT = TORES - TCONSU
           !CANTCONS = TCONSU
           !Saldo_Entre = SALENT
           .Update
        End If
        On Error GoTo 0
    End With
    ReserTemp.Close
    Set ReserTemp = Nothing
    
    
    '
End Sub


Sub MOVICONSUFABFAC(FECONSUM%, CIXI%, NORFAX$, QCON, DEPX%)
    '
    CMOV$ = "CF"
    DOPRI$ = NORFAX$
    DOSEC$ = "PF-" + TRIM$(Mid$(NORFAX$, 4))
    CODICONJ$ = Label7
         REFE$ = "CONSUMO " + CODICONJ$
         Call MOVISTOK(FECONSUM%, CIXI%, IDLO$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, (-1) * QCON)
    '
End Sub






Private Sub Command4_Click()
If DERACCE%("CAMBIDEP", "CAMBIO DE DEPOSITOS") > 1 Then
    Call SELECHO("TADEPOSI")
    If VALACT1$("TADEPOSI") <> "" Then
      Text2 = TRIM$(VALACT1$("TADEPOSI"))
    End If
   End If
End Sub

Private Sub Command43_Click()
   Command43.Enabled = False
   Screen.MousePointer = 11
   FIN& = MSF1.Rows - 1
   MSF1.Redraw = False
   For i& = 1 To FIN&
       Call TRACE(20, i&, FIN&)
       Call COLOREAR_GRILLA_SOLAFILA(Me.MSF1, &HFFFF80, i&)
   Next i&
   Call ACTUTOTALES
   MSF1.Redraw = True
   Screen.MousePointer = 1
   Command43.Enabled = True
End Sub

Private Sub Command5_Click()
    '
    Exit Sub
    FF% = FECHANUM("01/01/14")
    Desde = FETEXCOR1$(FF%)
    Dim Rst As ADODB.Recordset
    Set Rst = New ADODB.Recordset
    SQLX$ = "select cantsalid, cantingre from movisto where codimovi='cf' and cantingre>0 and fechmov>'" & Desde & "'"
    Rst.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    With Rst
        Do While Not .EOF
            !CANTSALID = !CANTINGRE
            !CANTINGRE = 0
            .Update
            .MoveNext
        Loop
    End With
    '
    Rst.Close
    Set Rst = Nothing
    '
    Call COMUNI("Proceso Finalizado")
End Sub

Private Sub Command6_Click()
If DERACCE%("CAMBIDEP", "CAMBIO DE DEPOSITOS") > 1 Then
    Call SELECHO("TADEPOSI")
    If VALACT1$("TADEPOSI") <> "" Then
      Text1 = TRIM$(VALACT1$("TADEPOSI"))
    End If
   End If
End Sub

Private Sub exportaexcel_Click()
   
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.Filter = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   AREXPO$ = UCase$(CommonDialog1.FileName)
   If TRIM$(AREXPO$) <> "" Then
      Call GENEXCEL(AREXPO$, Me.MSF1)
   End If
   '
99
End Sub
Sub GENEXCEL(NOMBREDEARCHIVO$, GRID As msFlexGrid)
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim valor As Double
   Dim TEXTO As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   TipoDato% = 0: COLUM& = 0
   For i& = 1 To GRID.Rows
       RENG1& = i&: COLUM& = 0
       For COLU1& = 1 To GRID.Cols
          If GRID.ColWidth(COLU1& - 1) = 0 Then GoTo 10 ' OBVIA LAS COLUMNAS OCULTAS
          COLUM& = COLUM& + 1
          '
          VALORX = GRID.TextMatrix(i& - 1, COLU1 - 1)
          If VALORX = "OF-017626-1" Then
            A = Q
          End If
          If Len(VALORX) = 8 And Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
             FEXU$ = VALORX
             VALORX = FETEXCOR1$(Int(FECHANUM(FEXU$)))
             TipoDato% = 1
          End If
          ' ESTO ES PARA QUE EXPORTE COMO CERO Y NO ESPACIOS
          FFCC$ = UCase$(Left$(FORCAMPO$(COLU1&), 1))
          If FFCC$ = "I" Or FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
             If TRIM$(CStr(VALORX)) = "" Then
                VALORX = 0
             End If
             TipoDato% = 2
          End If
          If FFCC$ = "C" Or FFCC$ = "A" Then VALORX = "'" & VALORX
          '
          '
          If TipoDato% = 0 And IsDate(VALORX) Then 'SI ES FECHA
            If Not IsNumeric(VALORX) Then          'SI NO ES NUMERO
               VALORX = " " & VALORX & " "         'LE AGREGA UN ESPACIO DELANTE Y AL FINAL (SOLO AL FINAL NO FUNCIONA9
            End If
          End If
          XLSHEET.Cells(RENG1&, COLUM&) = VALORX
          If TipoDato% = 1 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yyyy"
          'If TIPODATO% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#,##0.00"
          If TipoDato% = 0 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
          TipoDato% = 0

          VALORX_ANT = VALORX
10     Next COLU1&
   Next i&
   '
   Close #N1%
   On Error GoTo 30
   XLSHEET.SaveAs NOMBREDEARCHIVO$
   On Error GoTo 0
   '
   XLAPP.QUIT      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   Screen.MousePointer = 1
   Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NOMBREDEARCHIVO$, "\", "/"))
   '
Exit Sub
   '
   '
20 Resume 21
21 On Error GoTo 0
   MsgBox "Utilitario EXCEL no Instalado"
   On Error GoTo 0
   Exit Sub
   '
30 Resume 31
31 On Error GoTo 0
   Screen.MousePointer = 1
   '
   ' Cierra el objeto Book, el cual es utilizado al crear
   ' una nueva hoja en el archivo Excel.
   ' El valor False es para que no pregunte si quiere
   ' guardar el archivo.
   XLBOOK.Close (False)
   XLAPP.QUIT ' Cierra Excel
   Set XLAPP = Nothing
   Set XLBOOK = Nothing
   Set XLSHEET = Nothing
   '
   Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
   '
End Sub



Private Sub Form_Load()
   Call APLICACIONSKINS(Me, Picture1)
   Call CARGA_ENCABEZADO(MSF1, "Ord.Fab./A12;Código/A16;Descripción/A32;F.Lanz./d8;Proyec./F8.0;Aprob./F8.0;Q.A Cerrar/F11.0;CODINT/F0.0;NUMEORFA/F0.0;Pedido N°/F8", Me)
End Sub

Sub ACTUCANTISELECCION()
   ACUMUCANTI = 0
   For i& = 1 To MSF1.Rows - 1
     If MSF1.TextMatrix(i&, 0) = "*" Then
       ACUMUCANTI = ACUMUCANTI + 1
     End If
   Next i&
   Text23 = ACUMUCANTI
End Sub

Sub ACTUTOTALES()
   'Call ACTUTOTALLISTA
   Text19 = TOTALFILAS&(Me.MSF1)
   Call ACTUCANTISELECCION

End Sub

Private Sub grabacierre_Click()
   Call Command3_Click
End Sub

Private Sub mnuRevertirSeleccion_Click()
   Call Command43_Click
End Sub

Private Sub MSF1_Click()
    Static MODO  As Boolean
    REG& = MSF1.MouseRow
    
    If (REG& = 0) Then
        ' Ordena en forma ascendente
        If MODO Then
            MSF1.COL = MSF1.MouseCol
            MSF1.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSF1.COL = MSF1.MouseCol
            MSF1.Sort = 1
            MODO = True
        End If
        Exit Sub
    End If

End Sub

Private Sub MSF1_DblClick()
    i& = MSF1.MouseRow
    J& = MSF1.MouseCol
    If i& = 0 Or i& > MSF1.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF1.row = i&
    MSF1.COL = J&
    On Error Resume Next

    REG& = i&
    If REG& < 1 Then
        Exit Sub
    End If
    Call COLOREAR_GRILLA_SOLAFILA(Me.MSF1, &HFFFF80, REG&)
    Call ACTUTOTALES

End Sub

Private Sub MSF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
     REG& = MSF1.MouseRow
     COL% = MSF1.MouseCol
     Call MENU_CELDA(MSF1, REG&, COL%)
  End If
End Sub
Sub MENU_CELDA(MSF As msFlexGrid, REG&, COL%)
        If MSF.Rows <= 1 Then Exit Sub
        
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        PopupMenu MNUARCHIVO
        Call INVICTRL
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA
End Sub
Private Sub mnucopiar_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub

Sub INVICTRL()
  '
  mnucopiar.Visible = Not (mnucopiar.Visible)
  '
End Sub


Private Sub SelRangoFechas_Click()
  Call Command1_Click
End Sub

Private Sub Text1_Change()
   Text4 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text1)))
End Sub

Private Sub Text2_Change()
  Text3 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text2)))
End Sub
