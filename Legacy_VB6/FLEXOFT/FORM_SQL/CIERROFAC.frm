VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CIERROFAC 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cierre Automático de Ordenes de Fabriación, Facturadas"
   ClientHeight    =   8820
   ClientLeft      =   75
   ClientTop       =   285
   ClientWidth     =   15300
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8820
   ScaleWidth      =   15300
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8520
      Width           =   2085
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
      Left            =   6800
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   8520
      Width           =   2085
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
      Top             =   8520
      Width           =   175
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
      TabIndex        =   18
      Top             =   8520
      Width           =   175
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
      Left            =   6440
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   8520
      Width           =   390
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
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "Depósito Cierre"
      ToolTipText     =   "Cantidad de Item Seleccionados"
      Top             =   8520
      Width           =   1695
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
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "Depósito Consumo"
      Top             =   8520
      Width           =   1695
   End
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   8520
      Width           =   375
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
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   8520
      Width           =   765
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
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "Cantidad FilasTotal:"
      Top             =   8520
      Width           =   1815
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
      Left            =   11640
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "Cant. Filas Selecc.:"
      ToolTipText     =   "Cantidad de Item Seleccionados"
      Top             =   8520
      Width           =   1815
   End
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
      Left            =   13440
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   8520
      Width           =   765
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   8880
      Left            =   14400
      ScaleHeight     =   8820
      ScaleWidth      =   870
      TabIndex        =   5
      Top             =   0
      Width           =   930
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
         Left            =   120
         TabIndex        =   24
         Top             =   6720
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Height          =   690
         Left            =   120
         Picture         =   "CIERROFAC.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selección de Rango de Fechas"
         Top             =   2040
         Width           =   650
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
         Left            =   120
         Picture         =   "CIERROFAC.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Actualizar Indice de Selección de O.Fab."
         Top             =   1080
         Width           =   650
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
         Left            =   120
         Picture         =   "CIERROFAC.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Cargar Items"
         Top             =   7560
         Width           =   650
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
         Left            =   120
         Picture         =   "CIERROFAC.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "CIERROFAC.frx":08A8
         Top             =   8400
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   8415
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   14400
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Bindings        =   "CIERROFAC.frx":27CA
         Height          =   8175
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   14100
         _ExtentX        =   24871
         _ExtentY        =   14420
         _Version        =   393216
         BackColor       =   16777215
         FocusRect       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
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
         Left            =   9300
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   8040
         Width           =   4900
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
         Height          =   315
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   8040
         Width           =   7125
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
      OleObjectBlob   =   "CIERROFAC.frx":27DE
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CIERROFAC.frx":2A12
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Menu AGREGAR_RUTA 
      Caption         =   "Agregar Ruta"
      Visible         =   0   'False
      Begin VB.Menu mnuedi 
         Caption         =   "Editar"
      End
   End
End
Attribute VB_Name = "CIERROFAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
   Command1.Enabled = False
   '
   'VER SI GENERANDO LA TABLA QUE GUARDE EL NPED Y EL CODIGO PARA OBTENER LA CANTIDAD YA CERRADA DE ESE PEDIDO
   '*******VER MEJOR DE TOMAR LA CANTIDAD APROBADA DE LAS OF  CON UN SUM(CANAPAROBADA) DE LA OF SEGUN CODIGO Y N.PED.
   'SI CUANDO VALIDA 'If SALDO_POR_CERRAR > 0 Then' SE CONSIDERA TAMBIEN LA CANTIDAD YA CERRADA PARA ESE ITEM Y N PED.
   'CAMPOS NPED, CODIGO, OF, CANTIDAD CERRADA.
   '
   'ESTA FALLANDO EN EL CASO QUE EN EL MISMO PEDIDO HAY MAS DE UN ITEM REPETIDO SI SE REALIZA UN CIERRE POR 1 ITEM DE LOS REPETIDOS
   'CUANDO SE VUELVE A LLAMAR AL FRM SIN FACTURAR NUEVAMENTE APARECE LA CANTIDAD FACTURADA EN EL SEGUNDO ITEM.

   'ESTO FUNCIONA SOLO EN LOS CASOS QUE TODOS LOS PEDIDOS GENEREN ORDEN DE FABRICACION.
   '************************************************************************************
   
   
   MARCADECIERRE_ENVA% = 0
   EPAC$ = UCase$(TRIM$(EMPREAC$))
   If InStr(EPAC$, "CAMBACERES") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "ENVAPLAST") > 0 Then
     MARCADECIERRE_ENVA% = 1
   End If

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
   SQLX$ = SQLX$ + " WHERE NROPED > 0 And STATUORF < 3 "
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
        CANPRO = XVALO(!CanApro)
        CANPRY = XVALO(!CANPROY)
'        ReDim Preserve MATRIZ&(CI1%, NROPED&)= CANPRO
'        MATRIZ&(CI1%, NROPED&) = NROPED&
        SALDO_POR_CERRAR = XVALO(CANPRY) - XVALO(CANPRO)
        If SALDO_POR_CERRAR < 0 Then SALDO_POR_CERRAR = 0
        'SI ES CONTE, CAMBA O ENVA--PLAST
        If MARCADECIERRE_ENVA% > 0 Then GoTo 50 'SALTA Y COMPLETA

        If SALDO_POR_CERRAR > 0 Then
            'NO LO HICE CON VALOBADA POR QUE ERA MAS LENTO
            sqlx1$ = "SELECT DOSECCOR,DOPRICOR,DOPRILAR,CantSalid,FECHMOV  FROM MOVISTO WITH(NOLOCK)"
            sqlx1$ = sqlx1$ + " WHERE NuPedCli = " & NROPED&
            sqlx1$ = sqlx1$ + " AND COD_INTE = " & CI1%
            sqlx1$ = sqlx1$ + " AND CodiMovi = 'RT' "
            sqlx1$ = sqlx1$ + " ORDER BY FECHMOV"
            Set rs1 = READSET(sqlx1$)
            With rs1
                Do While Not .EOF
                   CANFACTU = XVALO(!CANTSALID)
                   FECHAFACT$ = FECHATEX$(CVINT(!FechMov)) 'ME QUEDO CON LA ULTIMA FECHA
'                   NORFAX$ = SAFETEXT$(!DOSECLAR)
                   ACU_CANFACTU = (ACU_CANFACTU + CANFACTU) ' - (YACONSIDERADO1)
                  .MoveNext
                Loop
            End With
            rs1.Close
            Set rs1 = Nothing
        End If
        If XVALO(ACU_CANFACTU) > CANPRO Then
       'comienzo de modificacion
       'esto se agrego por si hay algun item que se repite en el mismo pedido
       'sucedia que duplicaba la inforamcion por la misma cantidad a cerrar.
       QANTCERRAR = 0 ' solo ingresa si hay mismo pedido mismo articulo
       For k& = 1 To MSF1.Rows - 1
          If MSF1.TextMatrix(k&, 5) = NROPED& Then
            If MSF1.TextMatrix(k&, 10) = CI1% Then
               CANTIFACTU = XVALO(MSF1.TextMatrix(k&, 8))
               QANTCERRAR = QANTCERRAR + XVALO(MSF1.TextMatrix(k&, 9))
            End If
          End If
       Next k&
'       '
'       If QANTCERRAR > 0 Then
'        If QANTCERRAR <= SALDO_POR_CERRAR Then
'          GoTo 30
'        Else
'          SALDO_POR_CERRAR = SALDO_POR_CERRAR - QANTCERRAR
'        End If
'       End If
       'fin modificacion
50
            i& = i& + 1
            MSF1.Rows = i& + 1
            MSF1.TextMatrix(i&, 1) = NORFA$
            MSF1.TextMatrix(i&, 2) = CODIEXT$
            MSF1.TextMatrix(i&, 3) = DESQRI$
            MSF1.TextMatrix(i&, 4) = FECHALANZ$
            MSF1.TextMatrix(i&, 5) = NROPED&
            MSF1.TextMatrix(i&, 6) = FORMATNUM$(CANPRY, "F9.1")
            MSF1.TextMatrix(i&, 7) = FORMATNUM$(CANPRO, "F9.1")
            MSF1.TextMatrix(i&, 8) = ACU_CANFACTU
              CANT_A_CERRAR = ACU_CANFACTU - (CANPRO + QANTCERRAR): If CANT_A_CERRAR > SALDO_POR_CERRAR Then CANT_A_CERRAR = SALDO_POR_CERRAR
            MSF1.TextMatrix(i&, 9) = CANT_A_CERRAR
            MSF1.TextMatrix(i&, 10) = CI1%
            MSF1.TextMatrix(i&, 11) = FECHAFACT$
            'MSF1.TextMatrix(I&, 12) = SAFETEXT$(!DOPRILAR)
            
         End If
30     .MoveNext
       ACU_CANFACTU = 0

     Loop
   End With
   MSF1.Redraw = True
   If MARCADECIERRE_ENVA% > 0 Then
      If i& = 0 Then Call COMUNI("No se Han Encontrado Items Con Orden de Fabricación Sin Cerrar")
   Else
      If i& = 0 Then Call COMUNI("No se Han Encontrado Items Facturados Con Orden de Fabricación Sin Cerrar")
   End If
   Call ACTUTOTALES
99 Command1.Enabled = True
   Screen.MousePointer = 1
   
    
   
   
   
'''   SQLX$ = " SELECT M.ReferCor, M.FECHMOV, M.DOPRICOR, M.DESCRIP,M.COD_EXTE AS MCODEX, M.DOSECCOR, M.CANTSALID, O.IDSUBOR, O.NROPED, O.COD_INTE, O.CANAPRO, O.CANPROY, O.COD_EXTE"
'''   SQLX$ = SQLX$ + " FROM MOVISTO M WITH(NOLOCK)  INNER JOIN ORDEFABR O WITH(NOLOCK)  ON O.NROPED = M.NUPEDCLI AND O.COD_INTE = M.COD_INTE"
'''   SQLX$ = SQLX$ + " Where M.NUPEDCLI > 0 And O.STATUORF < 3 "
'''   SQLX$ = SQLX$ + " and M.FECHMOV >= '" & DESDE & "'"
'''   SQLX$ = SQLX$ + " and M.FECHMOV <= '" & HASTA & "'"
'''   SQLX$ = SQLX$ + " and M.ReferCor <> 'Fabricación Propia - Aut_x_Fact.'"
'''   SQLX$ = SQLX$ + " ORDER BY M.FECHMOV,  M.DOSECCOR, O.IDSUBOR "
'''   MSF1.Redraw = False
'''   Set RS = READSET(SQLX$)
'''   I& = 0
'''   With RS
'''     Do While Not .EOF
'''        CODIEXT$ = SAFETEXT$(!MCODEX)
'''        DESQRI$ = SAFETEXT$(!DESCRIP)
'''        FECHAM = FECHATEX$(CVINT(!FECHMOV))
'''        FCT$ = SAFETEXT$(!DOSECCOR)
'''        REMIX$ = SAFETEXT$(!DOPRICOR)
'''        NORFA$ = SAFETEXT$(!IDSUBOR)
'''        CANFACTU = XVALO(!CANTSALID)
'''        NROPED& = XVALO(!NROPED)
'''        CI1% = XVALO(!COD_INTE)
'''        CANPRO = XVALO(!CANAPRO)
'''        CANPRY = XVALO(!CANPROY)
'''        '
'''        I& = I& + 1
'''        MSF1.Rows = I& + 1
'''        MSF1.TextMatrix(I&, 1) = CODIEXT$
'''        MSF1.TextMatrix(I&, 2) = DESQRI$
'''        MSF1.TextMatrix(I&, 3) = FECHAM
'''        MSF1.TextMatrix(I&, 4) = FCT$ ' FORMATNUM$(CANFACTU, "F9.1")
'''
'''        MSF1.TextMatrix(I&, 5) = REMIX$
'''        MSF1.TextMatrix(I&, 6) = NORFA$
'''        MSF1.TextMatrix(I&, 7) = FORMATNUM$(CANFACTU, "F9.1")
'''        MSF1.TextMatrix(I&, 8) = NROPED&
'''        MSF1.TextMatrix(I&, 9) = FORMATNUM$(CANPRY, "F9.1")
'''        MSF1.TextMatrix(I&, 10) = FORMATNUM$(CANPRO, "F9.1")
'''         MSF1.TextMatrix(I&, 11) = CI1%
'''        .MoveNext
'''     Loop
'''   End With
'''   MSF1.Redraw = True
'''   If I& = 0 Then Call COMUNI("No se Han Encontrado Items Facturados Con Orden de Fabricación Sin Cerrar")
'''   Call ACTUTOTALES
'''99 Command1.Enabled = True
'''   Screen.MousePointer = 1
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
   
'   MarcarOFComoCerrada
   MARCO_COMOCERRADAS% = 0
   EPAC$ = UCase$(TRIM$(EMPREAC$))
   If InStr(EPAC$, "CAMBACERES") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "ENVAPLAST") > 0 Then
     MARCO_COMOCERRADAS% = 1
   End If
   
   If MARCO_COMOCERRADAS% = 1 Then
      For i& = 1 To MSF1.Rows - 1
         If MSF1.TextMatrix(i&, 0) = "*" Then
           NORFA$ = TRIM$(MSF1.TextMatrix(i&, 1))
           Call MarcarOFComoCerrada(NORFA$)
         End If
      Next i&
      GoTo 90
   End If
   
   If TRIM$(Text4) = "" Then
      Call COMUNI("Falta Depósito de Cierre (Ingreso)")
      GoTo 99
   End If
   If TRIM$(Text3) = "" Then
      Call COMUNI("Falta Depósito de Consumo")
      GoTo 99
   End If
   '
   MSF1.Height = 7800

   
   For i& = 1 To MSF1.Rows - 1
       If MSF1.TextMatrix(i&, 0) = "*" Then
         CI1% = XVALO(MSF1.TextMatrix(i&, 10))
         NORFA$ = TRIM$(MSF1.TextMatrix(i&, 1))
         CANTI = XVALO(MSF1.TextMatrix(i&, 9))
         FFACT% = CVINT(MSF1.TextMatrix(i&, 11))
         NROPED& = XVALO(MSF1.TextMatrix(i&, 5))
         CANTIPROYEC = XVALO(MSF1.TextMatrix(i&, 6))
         CANTIYA_APROBADA = XVALO(MSF1.TextMatrix(i&, 7))
         CANTIFACTU = XVALO(MSF1.TextMatrix(i&, 8))
         
         'SE VALIDA SI HAY REPETICION DE CODIGO Y PEDIDO LUEGO DE LA MODIFICACION DONDE EN UN  MISMO PEDIDO ESTA EL MISMO CODIGO
         'TAMBIEN CONSIDERANDO QUE SI SE FACTURO MAS  QUE LA ORDEN NO DEBERIA CERRAR POR MAS CANTIDAD.
         REPETIDO% = 0
         For k& = 1 To MSF1.Rows - 1
              If k& <> i& Then ' si no es el que esta recorriendo
               If MSF1.TextMatrix(k&, 5) = NROPED& Then
                 If MSF1.TextMatrix(k&, 10) = CI1% Then
                   REPETIDO% = 1
                   Exit For
                 End If
               End If
             End If
         Next k&
         '
         'SI NO ESTA REPETIDO MISMO CODIGO Y PEDIDO HACE EL AJUSTE DE LA CANTIDAD
         'CASO CONTRARIO NO HACE EL AJUSTE P Q SI NO SIEMPREA AJUSTA DE MAS POR QUE LA CANTIDAD FACTURADA APARECE DUPLICADA EN LA GRILLA.
         If REPETIDO% < 1 Then
            If CANTIFACTU > CANTI Then 'PARA EL CASO QUE SE HAYA FACTURADO MAS QUE LA CANTIDAD DE LA ORDEN DE LA ORDEN DE FABRICACION RELACIONADA.
               CANTI = CANTIFACTU - CANTIYA_APROBADA
            End If
         End If
         '
         If CANTI > 0 Then
           Call CIERRORFACTU(CI1%, NORFA$, CANTI, FFACT%)
         End If
       End If
   Next i&
   '
90 Call COMUNI("Proceso Finalizado")
   '
   Unload Me
99 Command3.Enabled = True
   Screen.MousePointer = 1
End Sub
Sub CIERRORFACTU(CI1%, NORFA$, CANTI, FFACT%)
    Text5 = "Resolviendo Cierre de Orden de Fabricación de Item: " & CODEXT$(CI1%) & " - " & DESCRIT0$(CI1%)
    Call ENRAPRODUCTO_TERMINADO(CI1%, NORFA$, CANTI, FFACT%, XVALO(Text2))
    '
    Text7 = "Verificando Componentes Reservados"
    Text7.Refresh
    CAN# = CANTI
    Call CARCOMPO(CODEXT$(CI1%), NORFA$, CAN#, 0, 1)
    Text7 = "Resolviendo Consumo de Componentes"
    Text7.Refresh

    Call CONSUCOMPONETESXOF(CI1%, NORFA$, CANTI, FFACT%, XVALO(Text1))
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
      FECHA% = FFACT%
      Text7 = "Actualizando Reservas "
      Text7.Refresh
      Call ACTURESERVA(CICOMPO%, NORFA$, CANCONSUM)
    Next U&
    '
    Text7 = "Escribiendo Anotador de Ordenes de Fabricación"
    Text7.Refresh
    Call ANOTAENORFA_FACTURADO(CI1%, NORFA$, CANTI, FFACT%)
    Call ACTUALIZAORFA(CI1%, NORFA$, CANTI, FFACT%)
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
   For i& = 1 To MSF1.Rows - 1
     
       Call COLOREAR_GRILLA_SOLAFILA(CIERROFAC.MSF1, &HFFFF80, i&)
     
   Next i&
   Call ACTUTOTALES
   Command43.Enabled = True
End Sub

Private Sub Command5_Click()
    '
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

Private Sub Form_Load()
   Call APLICACIONSKINS(Me, Picture1)
   Call CARGA_ENCABEZADO(MSF1, "Ord.Fab./A12;Código/A16;Descripción/A32;F.Lanz./d8;Nro.Ped./F8.0;Proyec./F8.0;Aprob./F8.0;Q.Fact./F8.0;Q.A Cerrar/F11.0;CODINT/F0.0;F.FACTURA/F10.0", CIERROFAC)
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
   Text19 = TOTALFILAS&(CIERROFAC.MSF1)
   Call ACTUCANTISELECCION

End Sub
Private Sub MSF1_DblClick()
    i& = MSF1.MouseRow
    J& = MSF1.MouseCol
    If i& = 0 Or i& > MSF1.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF1.Row = i&
    MSF1.COL = J&
    On Error Resume Next

    Reg& = i&
    If Reg& < 1 Then
        Exit Sub
    End If
    Call COLOREAR_GRILLA_SOLAFILA(CIERROFAC.MSF1, &HFFFF80, Reg&)
    Call ACTUTOTALES
End Sub

Private Sub Text1_Change()
   Text4 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text1)))
End Sub

Private Sub Text2_Change()
  Text3 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text2)))
End Sub
