VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form PICKCON09 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Remitos de Consignación"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9570
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   28
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.TextBox Text2 
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
      Left            =   4250
      TabIndex        =   18
      Top             =   7665
      Width           =   4290
   End
   Begin VB.TextBox Text1 
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
      Left            =   4550
      TabIndex        =   16
      Top             =   105
      Width           =   3975
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Bultos Controlados"
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
      Height          =   3255
      Left            =   105
      TabIndex        =   10
      Top             =   4305
      Width           =   8450
      Begin VB.Frame Frame5 
         Caption         =   "                                                                                                                  "
         Height          =   2800
         Left            =   6615
         TabIndex        =   11
         Top             =   420
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "                "
         Height          =   2800
         Left            =   2205
         TabIndex        =   12
         Top             =   420
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
         Height          =   2400
         ItemData        =   "PICKCON09.frx":0000
         Left            =   105
         List            =   "PICKCON09.frx":0002
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   8235
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Código                         Descripción                                                                   Cantidad"
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
         Left            =   105
         TabIndex        =   14
         Top             =   360
         Width           =   8235
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Detalle del Remito"
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
      Height          =   2835
      Left            =   120
      TabIndex        =   5
      Top             =   1365
      Width           =   8450
      Begin VB.Frame Frame1 
         Caption         =   "                                                                                                                  "
         Height          =   2525
         Left            =   6615
         TabIndex        =   9
         Top             =   420
         Width           =   15
      End
      Begin VB.Frame Frame14 
         Caption         =   "                "
         Height          =   2525
         Left            =   2205
         TabIndex        =   7
         Top             =   420
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
         Height          =   2040
         ItemData        =   "PICKCON09.frx":0004
         Left            =   105
         List            =   "PICKCON09.frx":0006
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   600
         Width           =   8235
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Código                         Descripción                                                                   Cantidad"
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
         Left            =   105
         TabIndex        =   6
         Top             =   360
         Width           =   8235
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   8610
      Left            =   8610
      ScaleHeight     =   8550
      ScaleWidth      =   1320
      TabIndex        =   1
      Top             =   -20
      Width           =   1380
      Begin VB.CommandButton Command4 
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
         Picture         =   "PICKCON09.frx":0008
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Aprobar Control de Salida"
         Top             =   6930
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
         Picture         =   "PICKCON09.frx":0312
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   225
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
         Picture         =   "PICKCON09.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   975
         Width           =   690
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   3
         Top             =   7800
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -285
         Picture         =   "PICKCON09.frx":0766
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   8040
         Width           =   1515
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "PICKCON09.frx":2688
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "PICKCON09.frx":28BC
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label10 
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
      Left            =   3600
      TabIndex        =   27
      Top             =   8085
      Width           =   4950
   End
   Begin VB.Label Label9 
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
      Left            =   945
      TabIndex        =   26
      Top             =   8085
      Width           =   2535
   End
   Begin VB.Label Label8 
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   120
      TabIndex        =   25
      Top             =   8115
      Width           =   630
   End
   Begin VB.Label Label7 
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
      Height          =   285
      Left            =   5565
      TabIndex        =   24
      Top             =   945
      Width           =   2955
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Remito:"
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
      Height          =   225
      Left            =   4680
      TabIndex        =   23
      Top             =   990
      Width           =   810
   End
   Begin VB.Label Label6 
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
      Height          =   285
      Left            =   1890
      TabIndex        =   22
      Top             =   525
      Width           =   6630
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
      Left            =   735
      TabIndex        =   21
      Top             =   525
      Width           =   1065
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cta."
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
      Height          =   225
      Left            =   360
      TabIndex        =   20
      Top             =   540
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Barras (Artículo) >>>>>>>>:"
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
      Height          =   225
      Left            =   1080
      TabIndex        =   19
      Top             =   7710
      Width           =   3150
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Barras (Remito) >>>>>>>>:"
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
      Height          =   225
      Left            =   1200
      TabIndex        =   17
      Top             =   150
      Width           =   3255
   End
End
Attribute VB_Name = "PICKCON09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command11_Click()
    '
    Unload Me
    '
End Sub

Private Sub Command4_Click()
   '
   Command4.Enabled = False
   '
   REMI$ = TRIM$(Label7)
   NP% = XVALO(Label4)
   '
   'VALIDACIONES
   If TRIM$(REMI$) = "" And NP% < 1 Then
      Call MENSERR(24, "No Hay un Remito Seleccionado")
      GoTo 99
   End If
   '
   If List1.ListCount > 0 Then
      Call MENSERR(24, "Imposible Grabar Control Quedan Bultos sin Controlar")
      GoTo 99
   End If
   '
   If List2.ListCount < 1 Then
      Call MENSERR(24, "Imposible Grabar Control No Hay Bultos Controlados")
      GoTo 99
   End If
   '
   'GENNERO CONDICION Y GRABO
   CONDI$ = " (NUME_PROV = " & NP% & " OR NUME_PROV=0) "
   CONDI$ = CONDI$ + " AND (CodiMovi = 'RC' OR CodiMovi='TR') "
   CONDI$ = CONDI$ + " AND (DoPriLar =  '" & REMI$ & "' OR DoSecLar =  '" & REMI$ & "')"
   CONDI$ = CONDI$ + " AND COD_INTE > 0"
   '
   USERN% = USNBR% Mod 100
   '
   Call ACTUREGISTRO("MOVISTO", "VERIFICADOPOR", CONDI$, USERN%, REGAF&, "NOMESS")
   'ESTE ACTUREGISTRO SE LE HACE UN Left(Now, 19) PARA QUE NO GUARDE EL PREFIJO AM O PM POR QUE SI NO DA ERROR
   Call ACTUREGISTRO("MOVISTO", "FHor_Verif", CONDI$, AHORA$, REGAF&)
   '
   Call COMUNI("Control de Salida Realizado")
   Call BLANFORMU
   Text1 = "" ' caja de texto del codigo de barra del remito
   Text1.SetFocus

99 Command4.Enabled = True

End Sub

Private Sub Form_Activate()
   '
   Call APLICACIONSKINS(Me, Picture14)
   '
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0

End Sub

Private Sub Form_Load()
   '
   'ACTUALIZO LOS CAMPOS NULL
   Call PoneEnCeroNull("MOVISTO", "VERIFICADOPOR", REGAF&, "NOMESS")
   Call ACTUREGISTRO("MOVISTO", "FHOR_VERIF", "FHOR_VERIF Is Null", NUNCA$, REGAF&, "NOMESS")
   '
End Sub

Private Sub Text1_Change()

    Call BLANFORMU
    If Len(TRIM$(Text1)) = 14 Then
        If Control$("", "CODIBAR") = "COBAR128" Then
           Call Text1_KeyPress(13)
        End If
    End If
    '
    

End Sub

Private Sub Text1_GotFocus()
    'texto seleccionado por si pica con el texto en pantalla
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 13 Then
      Exit Sub
    End If
    '
    'REALIZAR EL CREACAMPO
    Call ABRESTOCKS
    '
    LECTURA$ = TRIM$(Text1.TEXT)
    '
    'ARMO EL NUMERO DE REMITO
    NP% = XVALO(Left$(LECTURA$, 6))
    NUREM$ = TRIM$(Str$(XVALO(Mid$(LECTURA$, 7))))
    REMI$ = "RC-" & NUREM$
    Label4 = Str$(NP%)
    Label6 = ECOARCH$("PROVED1", MKI$(NP%))
    Label7 = REMI$
    '
    'LO BUSCO EN MOVISTO
    Call APERBASDAT("STOCKS")
   '
   'SE AGREGA GROUP BY PARA QUE EN UN SOLO REGISTRO TRAIGA TODA LA CANTIDAD PARA DESPUES PODER
   'DIVIDIR EN LA CANTIDAD DE FILAS SEGUN EL LOTE INDIVISIBLE.
    SQLX$ = " SELECT SUM(cantsalid - CantIngre) AS CANTI, Cod_Inte FROM MOVISTO WITH(NOLOCK)"
    'sqlx$ = sqlx$ + "WHERE NUME_PROV = " & NP% & " "
    SQLX$ = SQLX$ + "WHERE (NUME_PROV = " & NP% & " OR NUME_PROV=0)"
    'sqlx$ = sqlx$ + " AND CodiMovi = 'RC' "
    SQLX$ = SQLX$ + " AND (CodiMovi = 'RC' OR CodiMovi='TR')"
    'sqlx$ = sqlx$ + " AND DoPriLar =  '" & REMI$ & "'"
    SQLX$ = SQLX$ + " AND (DoPriLar =  '" & REMI$ & "' OR DoSecLar='" & REMI$ & "')"
    SQLX$ = SQLX$ + " AND COD_INTE > 0"
    SQLX$ = SQLX$ + " AND (VERIFICADOPOR =  0 OR VERIFICADOPOR IS NULL)"
    SQLX$ = SQLX$ + " GROUP BY COD_INTE"
    '
    Dim MOVITMP As ADODB.Recordset
    Set MOVITMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With MOVITMP
       If .EOF Then
         Call COMUNI("Sin Pendientes Para el Remito: " & REMI$)
         Exit Sub
       End If
       '
       Do While Not .EOF
         CI% = XVALO(!cod_inte)
         CAN = XVALO(!CANTI)
         '
         If CAN > 0.05 Then ' CARGA EL DETALLE EN EL LIST SEGUN LOTE INDIVISIBLE
            Call CARGO_INFO_LIST_CONS(CI%, CAN)
         End If
       .MoveNext
       Loop
    End With
    MOVITMP.Close
    Set MOVITMP = Nothing

    On Error Resume Next
    Text2.SetFocus
    On Error GoTo 0
    '
End Sub

Private Sub Text2_Change()
      Label9 = ""
      Label10 = ""
      'CI% = CODINT(Mid$(Text2, 8, 5))
      CI% = CODINTBARRA%(TRIM$(Text2))
      '
      Label9 = CODEXT$(CI%)
      If CI% < 1 Then Exit Sub
      Label10 = DESCRIT0$(CI%)
      '
      If Len(TRIM$(Text2)) = 13 Then
        If Control$("", "CODIBAR") = "COBAR128" Then
           Call Text2_KeyPress(13)
        End If
      End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
      '
      If KeyAscii <> 13 Then Exit Sub
      '
      Dim ESTA As Boolean
      '
      ESTA = False
      '
      'CI% = CODINT(Mid$(Text2, 8, 5))
      CI% = CODINTBARRA%(TRIM$(Text2))
      If CI% < 1 Then
         Call MENSERR(24, "Código No Existente en Detalle o Ya Descargado")
         Text2 = ""
         On Error Resume Next
         Text2.SetFocus
         On Error GoTo 0

         Exit Sub
      End If
      '
      For KI% = 1 To List1.ListCount
            CILIST% = CODINT%(Mid$(List1.List(KI% - 1), 1, 20))
            If CI% = CILIST% Then
              ESTA = True
              List2.AddItem List1.List(KI% - 1)
              List1.RemoveItem (KI% - 1)
              Text2 = ""
              Exit For
            End If
      Next KI%
      '
      If ESTA = False Then
         Call MENSERR(24, "Código No Existente en Detalle o Ya Descargado")
         Text2 = ""
         On Error Resume Next
         Text2.SetFocus
         On Error GoTo 0
         Exit Sub
      End If
      '
      If List1.ListCount < 1 Then ' si la lista de arriba esta vacia dierectamente graba
         Call Command4_Click
      End If

End Sub


Sub BLANFORMU()

    List1.Clear
    List2.Clear
    Label4 = ""
    Label6 = ""
    Label7 = ""
    Label9 = ""
    Label10 = ""
    Text2 = ""
    
End Sub
Sub CARGO_INFO_LIST_CONS(CI%, QTIDAD)
    '
    CANTACUMUL = 0
    COD$ = CODEXT$(CI%)
    Call REPLA$(TX$, COD$, 1, 20)
    Call REPLA$(TX$, DESCRIT0$(CI%), 21, 40)
    '
    CAN# = QTIDAD
    CANT_FILAS% = CANBULTOS(CI%, CAN#)
    '
    For i% = 1 To CANT_FILAS% - 1 ' -1 PARA GENERAR EL ULTIMO POR EL RESTO
       CANTI = ROUND(LINDIVIS(CI%), 0)
       CANTACUMUL = CANTACUMUL + CANTI
       CANTIX$ = FORMATNUM$(CANTI, "F14." & CNTDC$(CIXI%))
       Call REPLA$(TX$, CANTIX$, 62, 14)
       List1.AddItem TX$
       '
    Next i%
    '
    'CREO EL ULTIMO POR EL RESTO
    CANTI = QTIDAD - CANTACUMUL
    CANTIX$ = FORMATNUM$(CANTI, "F14." & CNTDC$(CIXI%))
    Call REPLA$(TX$, CANTIX$, 62, 14)
    List1.AddItem TX$

End Sub

