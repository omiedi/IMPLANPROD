VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form GENOFMAT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generación de OF a partir del Stock de Material Controlado"
   ClientHeight    =   8880
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   12990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   12990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNroCliente 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      MaxLength       =   24
      TabIndex        =   38
      Top             =   9480
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   1680
      TabIndex        =   37
      Top             =   9480
      Width           =   255
   End
   Begin VB.TextBox txtRazonSocial 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   9480
      Width           =   3015
   End
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   3120
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      Caption         =   "PRODUCTO A FABRICAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   5
      Top             =   6240
      Width           =   12735
      Begin VB.TextBox txtPesoBolsa 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         MaxLength       =   32
         TabIndex        =   40
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtSubcodigo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         MaxLength       =   24
         TabIndex        =   34
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtPartida 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         MaxLength       =   32
         TabIndex        =   32
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtDescripcionProductoTerminado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox txtCodigoProductoTerminado 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   27
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Height          =   300
         Left            =   2640
         TabIndex        =   26
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton btnGenerarOF 
         Caption         =   "Generar O/F"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11280
         TabIndex        =   10
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtCantidadAFabricar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   840
         Width           =   4575
      End
      Begin VB.CommandButton btnListaCodigos 
         Height          =   300
         Left            =   2640
         TabIndex        =   7
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   2415
      End
      Begin MSFlexGridLib.MSFlexGrid mfgFormulaCodigoAFabricar 
         Height          =   1935
         Left            =   4920
         TabIndex        =   15
         Top             =   480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3413
         _Version        =   393216
         BackColorFixed  =   16501405
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label13 
         Caption         =   "Peso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   41
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Sub código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   35
         Top             =   2010
         Width           =   555
      End
      Begin VB.Label Label10 
         Caption         =   "Partida"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   33
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Código Producto Terminado"
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
         Left            =   240
         TabIndex        =   28
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Fórmula de producto"
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
         Left            =   4920
         TabIndex        =   16
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad"
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
         Left            =   3480
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Código a Fabricar"
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
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "COMPONENTES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12735
      Begin VB.TextBox txtOrdenFabricacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10680
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   240
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid mfgMaterialDisponible 
         Height          =   2055
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   3625
         _Version        =   393216
         BackColorFixed  =   16501405
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame3 
         Caption         =   "FILTROS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4800
         TabIndex        =   19
         Top             =   2880
         Width           =   7695
         Begin VB.TextBox txtBuscarTexto 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox txtCodigoFamilia 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2760
            TabIndex        =   21
            Top             =   600
            Width           =   2175
         End
         Begin VB.CommandButton Command1 
            Height          =   300
            Left            =   4920
            TabIndex        =   20
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label9 
            Caption         =   "Buscar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   200
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label8 
            Caption         =   "Familia de Componentes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   200
            Left            =   2760
            TabIndex        =   22
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.TextBox txtAnotaciones 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   3120
         Width           =   4335
      End
      Begin MSFlexGridLib.MSFlexGrid mfgMaterialAsignado 
         Height          =   1695
         Left            =   240
         TabIndex        =   2
         Top             =   4080
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   2990
         _Version        =   393216
         BackColorFixed  =   16501405
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblOrdenFabricacion 
         Caption         =   "O-Fabricación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9480
         TabIndex        =   31
         Top             =   360
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Anotaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   240
         TabIndex        =   18
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "MATERIAL ASIGNADO PARA FABRICACION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   240
         TabIndex        =   4
         Top             =   3840
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "MATERIAL DISPONIBLE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   2415
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "GENOFMAT.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "GENOFMAT.frx":0234
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label12 
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
      Height          =   195
      Left            =   120
      TabIndex        =   39
      Top             =   9585
      Width           =   795
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
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
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4800
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnu 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuMaestroArticulos 
         Caption         =   "Maestro de Artículos"
      End
      Begin VB.Menu lineAccesosDirectos1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTablaFamiliaComponentes 
         Caption         =   "Tabla Familia de Componentes"
      End
   End
End
Attribute VB_Name = "GENOFMAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ESTADO As String
Dim UltimaFilaEncontrada As Long
Dim ModificaCantidadMaterial As Boolean

Sub ActualizaListaDeMaterialDisponible(CodigoFamilia As String)

    Dim Rst As ADODB.Recordset
    
    ' CLASIFICACION DE RECEPCIONES
    ' ** Lista todas las entradas que se produjeron utilizando lotes FB de fraccionamiento. O sea, aque que se produjo
    ' ** al fraccionar/clasificar, en una transferencia de stock entre deposito o en un ajuste. Obtiene ademas el lote
    ' ** BR original a partir del codigo de corte CB.
    ConsultaSql10$ = "(select frct.docufrac, frct.codifrac, frct.codexfrac, frct.docucorte, frct.depomovi, frct.saldofrac, frct.loterecep, b.numprov from" _
    & " (select f.docufrac, f.codifrac, f.codexfrac, f.docucorte, f.ingre saldofrac, f.depomovi, m.idenlote loterecep from" _
    & " (select fr.cod_inte codifrac, fr.cod_exte codexfrac, fr.doprilar docufrac, fr.doseclar docucorte, m.depomovi, sum(cantingre) ingre from" _
    & " (select cod_inte, cod_exte, doprilar, doseclar from movisto where codimovi='fb') fr" _
    & " inner join movisto m on (fr.doprilar=m.idenlote)" _
    & " where cantingre > 0.05" _
    & " group by fr.cod_inte, fr.cod_exte, fr.doprilar, fr.doseclar, m.depomovi) f" _
    & " inner join movisto m on (f.docucorte=m.dopricor)" _
    & " where m.codimovi='cb') frct" _
    & " inner join bolrecep b with(nolock) on (frct.loterecep=b.idenlote))"
    If CodigoFamilia <> "" Then
        ConsultaSql11$ = "(select md.docufrac, md.codifrac, md.codexfrac, md.docucorte, md.depomovi, md.saldofrac, md.loterecep, md.numprov from " _
        & "(" & ConsultaSql10$ & ") md" _
        & " inner join datasql d on (md.codexfrac=d.campo3)" _
        & " where campo1='" & CodigoFamilia & "')"
        ConsultaSql10$ = ConsultaSql11$
    End If
    
    ' FABRICACION DE SEMIELABORADOS
    ConsultaSql20$ = "select fab.doprilar docufrac, fab.cod_inte codifrac, cod_exte codexfrac, '' docucorte, fab.depomovi, fab.ingre saldofrac, fab.doprilar loterecep, '0' numprov from" _
    & " (select ciefab.cod_inte, ciefab.cod_exte, ciefab.doprilar, m.depomovi, sum(cantingre) ingre from" _
    & " (select cod_inte, cod_exte, doprilar from movisto mo with(nolock)" _
    & " inner join master ma with(nolock) on (mo.cod_inte=ma.codint)" _
    & " where mo.codimovi='OF' and ma.timate<>5" _
    & " group by mo.cod_inte, mo.cod_exte, mo.doprilar) ciefab" _
    & " inner join movisto m on (ciefab.doprilar=m.idenlote)" _
    & " where m.cantingre > 0.05" _
    & " group by ciefab.cod_inte, ciefab.cod_exte, ciefab.doprilar, m.depomovi) fab"
    If CodigoFamilia <> "" Then
        ConsultaSql21$ = "(select md.docufrac, md.codifrac, md.codexfrac, md.docucorte, md.depomovi, md.saldofrac, md.loterecep, md.numprov from " _
        & "(" & ConsultaSql20$ & ") md" _
        & " inner join datasql  d on (md.codexfrac=d.campo3)" _
        & " where campo1='" & CodigoFamilia & "')"
        ConsultaSql20$ = ConsultaSql21$
    End If
    
    ConsultaSql$ = ConsultaSql10$ & " union all " & ConsultaSql20$
    ConsultaSql$ = ConsultaSql$ & " order by codifrac, docufrac, depomovi"
    
    Set Rst = READSET(ConsultaSql$)
    
    With Rst
        mfgMaterialDisponible.Rows = 1
        Fila& = 0
        Do While Not .EOF
            DEPOSITO% = SAFETEXT$(!depomovi)
                CondicionSql$ = "cod_inte=" & CStr(!codifrac) & " and idenlote='" & SAFETEXT$(!docufrac) & "' and deposito=" & CStr(DEPOSITO%)
                If txtOrdenFabricacion <> "" Then CondicionSql$ = CondicionSql$ & " and idsubor<>'" & txtOrdenFabricacion & "'"
            NetoReservado# = XVALO(VALOBADA("reserva", "sum(cantres-cantcons)", CondicionSql$))
                CondicionSql$ = "cod_inte=" & CStr(!codifrac) & " and idenlote='" & SAFETEXT$(!docufrac) & "' and depomovi=" & CStr(DEPOSITO%)
            ConsumoFabrica# = XVALO(VALOBADA("movisto", "sum(cantsalid)", CondicionSql$))
            If (!saldofrac - ConsumoFabrica#) - NetoReservado# > 0.05 Then
                Fila& = Fila& + 1
                mfgMaterialDisponible.Rows = mfgMaterialDisponible.Rows + 1
                
                mfgMaterialDisponible.TextMatrix(Fila&, 1) = CODEXT$(!codifrac)
                mfgMaterialDisponible.TextMatrix(Fila&, 2) = DESCRIT0$(!codifrac)
                    DocumentoPrimario$ = SAFETEXT$(!docufrac)
                mfgMaterialDisponible.TextMatrix(Fila&, 3) = DocumentoPrimario$
                
                    LoteRecepcion$ = SAFETEXT$(!loterecep)
                    If Mid$(DocumentoPrimario$, 1, 2) = "OF" Then
                        LoteRecepcion$ = LoteOriginalDeMateriaPrima(DocumentoPrimario$)
                        n = 0
                        For i& = 1 To Len(LoteRecepcion$)
                            If Mid$(LoteRecepcion$, i&, 1) = ";" Then n = n + 1
                        Next i&
                        LoteRecepcion$ = REPLACAR$(LoteRecepcion$, ";", " " + Chr$(13) + Chr$(10))
                        If n = 0 Then n = 1
                        mfgMaterialDisponible.RowHeight(Fila&) = mfgMaterialDisponible.RowHeight(Fila&) * n
                        
                    End If
                mfgMaterialDisponible.TextMatrix(Fila&, 4) = LoteRecepcion$
                    
                    NumeroProveedor% = !NumProv
                    RazonSocialProveedor$ = ""
                    If Mid$(DocumentoPrimario$, 1, 2) = "OF" Then
                        Call TEXTOLOTES(LoteRecepcion$, Chr$(13) + Chr$(10))
                        For i& = 1 To UBound(CADENATEX$)
                            NumeroProveedor% = XVALO(VALOBADA("bolrecep", "numprov", "idenlote='" & TRIM$(CADENATEX$(i&)) & "'"))
                            RazonSocialProveedor$ = RazonSocialProveedor$ & Left$(rasocli$((-1) * NumeroProveedor%), 26) + " " + Chr$(13) + Chr$(10)
                        Next i&
                    Else
                        RazonSocialProveedor$ = rasocli$((-1) * NumeroProveedor%)
                    End If
                mfgMaterialDisponible.TextMatrix(Fila&, 5) = RazonSocialProveedor$
                
                mfgMaterialDisponible.TextMatrix(Fila&, 6) = DEPOSITO%
                
                mfgMaterialDisponible.TextMatrix(Fila&, 7) = FORMATNUM$(!saldofrac - ConsumoFabrica# - NetoReservado#, "F12.1")
            End If
    
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
End Sub

Function AnotacionesEnClasificacion(IDENTIFICADOR As String) As String

    If Mid$(IDENTIFICADOR, 1, 2) = "OF" Then
        Fraccionamientos$ = LotesDeFraccionamiento(IDENTIFICADOR) & ";"
    Else
        Fraccionamientos$ = IDENTIFICADOR & ";"
    End If
    
    Call TEXTOLOTES(Fraccionamientos$, ";")
    For i& = 1 To UBound(CADENATEX$())
        AnotacionPorFraccion$ = SAFETEXT(VALOBADA("anotmovi", "anotaciones", "doprilar='" & TRIM$(CADENATEX$(i&)) & "'"))
        If AnotacionPorFraccion$ <> "" Then
            If AnotacionTodasFracciones$ = "" Then
                AnotacionTodasFracciones$ = AnotacionPorFraccion$
            Else
                AnotacionTodasFracciones$ = AnotacionTodasFracciones$ & ". " & AnotacionPorFraccion$
            End If
        End If
    Next i
    
    AnotacionesEnClasificacion = AnotacionTodasFracciones$

End Function

Function LoteOriginalDeMateriaPrima(OrdenDeFabricacion As String) As String
    
    LoteOriginalDeMateriaPrima = ""
   
    ConsultaSql$ = "select idenlote from movisto with(nolock) where doprilar='" & OrdenDeFabricacion$ & "' and codimovi='cf'"
    Set rstCF = READSET(ConsultaSql$)
    
    With rstCF
        Do While Not .EOF
            LoteFraccionamiento$ = !idenlote
            If Mid$(LoteFraccionamiento$, 1, 2) <> "OF" Then
                LoteCorte$ = TRIM$(VALOBADA("movisto", "doseccor", "doprilar='" & LoteFraccionamiento$ & "'"))
                If LoteCorte$ <> "" Then
                    LoteRecepcion$ = TRIM$(VALOBADA("movisto", "idenlote", "doprilar='" & LoteCorte$ & "'"))
                
                    If InStr(1, LoteOriginalDeMateriaPrima, LoteRecepcion$) < 1 Then
                        LoteOriginalDeMateriaPrima = LoteOriginalDeMateriaPrima & LoteRecepcion$ & ";"
                    End If
                End If
                
            Else
                LoteRecepcion$ = LoteOriginalDeMateriaPrima(LoteFraccionamiento)
                
                Call TEXTOLOTES(LoteRecepcion$, ";")
                For i& = 1 To UBound(CADENATEX$())
                    If InStr(1, LoteOriginalDeMateriaPrima, CADENATEX$(i&)) < 1 Then
                        LoteOriginalDeMateriaPrima = LoteOriginalDeMateriaPrima & CADENATEX$(i&) & ";"
                    End If
                Next i
                
            End If
            
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    rstCF.Close
    Set rstCF = Nothing
    On Error GoTo 0
    
End Function

Function LotesDeFraccionamiento(OrdenDeFabricacion As String) As String
    
    LotesDeFraccionamiento$ = ""
   
    ConsultaSql$ = "select distinct idenlote from" _
    & " ((select idenlote from movisto with(nolock) where doprilar='" & OrdenDeFabricacion$ & "' and codimovi='cf')" _
    & " union all " _
    & " (select idenlote from reserva with(nolock) where idsubor='" & OrdenDeFabricacion & "')) l"
    Set rstCF = READSET(ConsultaSql$)
    
    With rstCF
        Do While Not .EOF
            LoteFraccionamiento$ = !idenlote
            If Mid$(LoteFraccionamiento$, 1, 2) <> "OF" Then
                If LotesDeFraccionamiento$ = "" Then
                    LotesDeFraccionamiento$ = LoteFraccionamiento$
                Else
                    LotesDeFraccionamiento$ = LotesDeFraccionamiento$ & "; " & LoteFraccionamiento$
                End If
                
            Else
                LoteFraccionamiento$ = LotesDeFraccionamiento$(LoteFraccionamiento)
                
                Call TEXTOLOTES(LoteFraccionamiento$, ";")
                For i& = 1 To UBound(CADENATEX$())
                    If InStr(1, LotesDeFraccionamiento$, CADENATEX$(i&)) < 1 Then
                        LotesDeFraccionamiento$ = LotesDeFraccionamiento$ & CADENATEX$(i&) & ";"
                    End If
                Next i
                
            End If
            
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    rstCF.Close
    Set rstCF = Nothing
    On Error GoTo 0
    
End Function



Sub BuscadorDeTextoEnGrilla(Texto As String, Optional FilaComienzo As Long)

    For iFil = 0 To mfgMaterialDisponible.Rows - 1
        If TRIM$(mfgMaterialDisponible.TextMatrix(iFil, 0)) = "*" Then
            mfgMaterialDisponible.TextMatrix(iFil, 0) = ""
            Call COLOREAR_GRILLA_SOLAFILA(mfgMaterialDisponible, RGB(255, 255, 255), iFil)
        End If
    Next iFil
    For iFil = 0 To mfgMaterialDisponible.Rows - 1
        If TRIM$(mfgMaterialDisponible.TextMatrix(iFil, 0)) = "*" Then
            mfgMaterialDisponible.TextMatrix(iFil, 0) = ""
            Call COLOREAR_GRILLA_SOLAFILA(mfgMaterialDisponible, RGB(255, 255, 255), iFil)
        End If
    Next iFil
    
    Inicio = 0
    If FilaComienzo > 0 Then Inicio = FilaComienzo
    For iFil = Inicio To mfgMaterialDisponible.Rows - 1
        For iCol = 1 To mfgMaterialDisponible.Cols - 1
            If InStr(1, mfgMaterialDisponible.TextMatrix(iFil, iCol), Texto, vbTextCompare) > 0 Then
                Call COLOREAR_GRILLA_SOLAFILA(mfgMaterialDisponible, RGB(241, 194, 50), iFil)
                Encontro% = 1: UltimaFilaEncontrada = iFil
                
                Exit For
            End If
        Next iCol
        If Encontro% = 1 Then
            If mfgMaterialDisponible.RowIsVisible(UltimaFilaEncontrada) = False Then
            End If
            Exit For
        Else
            UltimaFilaEncontrada = 0
        End If
    Next iFil

End Sub

Private Sub btnGenerarOF_Click()

    If CODINT%(TXTCODIGO) = 0 Then
        Call COMUNI("Falta Código de Producto a Fabricar")
        Exit Sub
    End If
    
    If mfgMaterialAsignado.Rows < 2 Then
        Call COMUNI("No Existen Componentes para la Fabricación")
        Exit Sub
    End If
    
    If TRIM$(txtCodigoProductoTerminado) <> "" Then
        OFABRI07.CodigoInternoPTOfStockComponentes = CODINT%(txtCodigoProductoTerminado)
        If OFABRI07.CodigoInternoPTOfStockComponentes > 0 Then
            KgBolsa# = XVALO(txtPesoBolsa)
            If KgBolsa# < 0.05 Then
                Call COMUNI("Falta Configurar Peso por Bolsa en el Producto Terminado")
                Exit Sub
            End If
        Else
            Call COMUNI("Código de Producto Terminado Inválido")
            Exit Sub
        End If
    End If
    
    txtNroCliente = "0"
'    If IsNumeric(txtNroCliente) = True Then
'        If CantRegistros("deudor12", "nume_cli=" & txtNroCliente) = 0 Then
'            Call COMUNI("Número de Cliente Inválido")
'            Exit Sub
'        End If
'    Else
'        txtNroCliente = "0"
'    End If
    
    ' VALIDA QUE LOS MATERIALES ASIGNADOS FORMEN PARTE DE LA FORMULA DE PRODUCTO
    Dim EsComponente As Boolean
    For i = 1 To mfgMaterialAsignado.Rows - 1
        EsComponente = False
        CodigoAsignado$ = TRIM$(mfgMaterialAsignado.TextMatrix(i, 1))
        
        For J = 1 To mfgFormulaCodigoAFabricar.Rows - 1
            CodigoFormula$ = TRIM$(mfgFormulaCodigoAFabricar.TextMatrix(J, 1))
            If mfgFormulaCodigoAFabricar.TextMatrix(J, 4) = "SI" Then ' VERIFICA SI ES FAMILIA
                If CantRegistros("datasql", "campo1='" & CodigoFormula$ & "' and campo3='" & CodigoAsignado$ & "'") > 0 Then
                    EsComponente = True
                    Exit For
                End If
            Else
                If CodigoAsignado$ = CodigoFormula$ Then
                    EsComponente = True
                    Exit For
                End If
            End If
        Next J
        If EsComponente = False Then
            Call COMUNI("El Material '" & CodigoAsignado$ & "'\No Pertenece a la Fórmula de Producto.\ \Debe Agregarlo como Componente\o Formar Parte de una Familia.")
            Exit Sub
        End If
    Next i
    
    OFABRI07.GeneraOfAPartirDeStockDeComponentes = 1
    OFABRI07.CodigoInternoOfStockComponentes = CODINT%(TXTCODIGO)
    OFABRI07.CantidadOfStockComponentes = XVALO(txtCantidadAFabricar)
    OFABRI07.FechaEntregaOfStockComponentes = FECHATEX$(HOY(HO$))
    OFABRI07.DestinoOfStockComponentes = "STOCK"
    OFABRI07.NumeroOfStockComponentes = txtOrdenFabricacion
'    OFABRI07.PartidaPTOfStockComponentes = txtPartida
'    OFABRI07.SubcodigoPTOfStockComponentes = txtSubcodigo
'    OFABRI07.DestinoPTStockComponentes = CLng(txtNroCliente)
'    If OFABRI07.CodigoInternoPTOfStockComponentes > 0 Then
'        OFABRI07.CantidadPTOfStockComponentes = OFABRI07.CantidadOfStockComponentes
'        OFABRI07.KgBolsaPTOfStockComponentes = KgBolsa#
'    Else
'        OFABRI07.DestinoOfStockComponentes = OFABRI07.DestinoPTStockComponentes
'    End If

    Dim da As New DatosAuxiliares
    
    Call SETULTREG("!RESERVA", 0)
    For i = 1 To mfgMaterialAsignado.Rows - 1
        YY$ = REGBLAN$("RESERVA")
            CIIX% = CODINT%(mfgMaterialAsignado.TextMatrix(i, 1))
        Call REPLA(YY$, MKI$(CIIX%), 1, 2)
        Call REPLA(YY$, "OF-000000-1", 3, 12)
        Call REPLA(YY$, MKI$(CIIX%), 15, 2)
        Call REPLA(YY$, MKI$(HOY(HO$)), 17, 2)
            CANTX# = XVALO(mfgMaterialAsignado.TextMatrix(i, 7))
        Call REPLA(YY$, MKD$(CANTX#), 19, 8)
        Call REPLA(YY$, MKD$(CANTX#), 35, 8)
            LOTE$ = mfgMaterialAsignado.TextMatrix(i, 3)
        Call REPLA(YY$, LOTE$, 43, 16)
            DEPOSITO% = XVALO(mfgMaterialAsignado.TextMatrix(i, 6))
        Call REPLA(YY$, Chr$(DEPOSITO%), 59, 1)
            ESTRAZ% = 1
        Call REPLA(YY$, MKS$(ESTRAZ%), 61, 4)
            CodigoFamilia$ = TRIM$(da.EcoTablaAuxiliar("FAMCOMPO", CODEXT$(CIIX%), 3, 1))
            If CodigoFamilia$ <> "" Then
                CodigoBuscar$ = CodigoFamilia$
            Else
                CodigoBuscar$ = CODEXT$(CIIX%)
            End If
            For iFor = 1 To mfgFormulaCodigoAFabricar.Rows - 1
                CodigoFormula$ = mfgFormulaCodigoAFabricar.TextMatrix(iFor, 1)
                If CodigoBuscar$ = CodigoFormula$ Then
                    mfgFormulaCodigoAFabricar.TextMatrix(iFor, 0) = "*"
                    USOII = XVALO(mfgFormulaCodigoAFabricar.TextMatrix(iFor, 3))
                    Exit For
                End If
            Next iFor
        Call REPLA(YY$, MKS$(USOII), 75, 4)
        Call REGAPP("!RESERVA", YY$)
    Next i
    
    For iFor = 1 To mfgFormulaCodigoAFabricar.Rows - 1
        If mfgFormulaCodigoAFabricar.TextMatrix(iFor, 0) <> "*" Then
            USOII = XVALO(mfgFormulaCodigoAFabricar.TextMatrix(iFor, 3))
            If USOII > 0.00005 Then
                YY$ = REGBLAN$("RESERVA")
                CIIX% = CODINT%(mfgFormulaCodigoAFabricar.TextMatrix(iFor, 1))
                Call REPLA(YY$, MKI$(CIIX%), 1, 2)
                Call REPLA(YY$, "OF-000000-1", 3, 12)
                Call REPLA(YY$, MKI$(CIIX%), 15, 2)
                Call REPLA(YY$, MKI$(HOY(HO$)), 17, 2)
                    
                    CANTX# = XVALO(txtCantidadAFabricar) * USOII
                Call REPLA(YY$, MKD$(CANTX#), 19, 8)
                Call REPLA(YY$, MKD$(CANTX#), 35, 8)
                    LOTE$ = ""
                Call REPLA(YY$, LOTE$, 43, 16)
                    DEPOSITO% = 0
                Call REPLA(YY$, Chr$(DEPOSITO%), 59, 1)
                    ESTRAZ% = 1
                Call REPLA(YY$, MKS$(ESTRAZ%), 61, 4)
                Call REPLA(YY$, MKS$(USOII), 75, 4)
                Call REGAPP("!RESERVA", YY$)
            End If
        End If
    Next iFor
    
    ESTADO = "APROBADO"
    If txtOrdenFabricacion <> "" Then ESTADO = "REGENERA"
    Hide
    
End Sub

Function EsFamiliaDeComponentes(Codigo As String) As Boolean

    EsFamiliaDeComponentes = False
    
    If CantRegistros("datasql", "campo1='" & Codigo & "'") > 0 Then
        EsFamiliaDeComponentes = True
    End If
    
End Function

Private Sub btnListaCodigos_Click()

    TXTCODIGO = ListaSeleccionArticulos(Descripcion$, True)
    TXTDESCRIPCION = Descripcion$
        
End Sub



Private Sub Command1_Click()

    Dim da As New DatosAuxiliares
    
    Call da.ListaSeleccionTablaAuxiliar("FAMCOMPO", "", "campo1/A16;campo2/A32", True)
    
    txtCodigoFamilia = RESP_ZELE_VECT(1)
    
    Set da = Nothing
    

    
End Sub

Private Sub Command2_Click()

    txtCodigoProductoTerminado = ListaSeleccionArticulos(Descripcion$)
    txtDescripcionProductoTerminado = Descripcion$
    txtSubcodigo = SAFETEXT$(VALOBADA("master", "cod_interna", "codigo='" & txtCodigoProductoTerminado & "'"))

End Sub

Private Sub Command3_Click()

    txtNroCliente = ListaSeleccionClientes
    txtRazonSocial = rasocli$(CLng(txtNroCliente))
    
End Sub

Private Sub Form_Load()

    UTILIZA_SKIN% = 1
    ESTADO = ""
    Call APLICACIONSKINS(Me)
    Call T_Espera(1 / 2)

    Call CARGA_ENCABEZADO(mfgMaterialDisponible, "Codigo/A16;Descripcion/A24;Lote Clas./A10;Lote Orig./A10;Proveedor/A23;Dep/I3;Cantidad/F10.1", Me)
    
    Call CARGA_ENCABEZADO(mfgMaterialAsignado, "Codigo/A16;Descripcion/A24;Lote Clas./A10;Lote Orig/A10;Proveedor/A23;Dep/I3;Cantidad/F10.1;Saldo/F0", Me)
    
    Call CARGA_ENCABEZADO(mfgFormulaCodigoAFabricar, "Codigo/A16;Descripcion/A20;Uso/F10.1;Familia/A7", Me)

    mfgMaterialDisponible.WordWrap = True
    mfgMaterialAsignado.WordWrap = True
    
    Call ActualizaListaDeMaterialDisponible("")
    
    mfgMaterialDisponible.ColSel = 2: mfgMaterialDisponible.COL = 1
    mfgMaterialDisponible.Sort = 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If ESTADO <> "CERRAR" Then
        Cancel = 1
        ESTADO = "CANCELADO"
        Hide
    End If
    
End Sub


Private Sub mfgFormulaCodigoAFabricar_DblClick()

    iFila& = mfgFormulaCodigoAFabricar.MouseRow
    If mfgFormulaCodigoAFabricar.TextMatrix(iFila&, 4) <> "SI" And XVALO(txtCantidadAFabricar) Then
        CantidadComponente# = XVALO(InputBox("CANTIDAD TOTAL A CONSUMIR", mfgFormulaCodigoAFabricar.TextMatrix(iFila&, 2)))
        If CantidadComponente# > 0.05 Then
            CantidadComponente# = CantidadComponente# / XVALO(txtCantidadAFabricar)
            mfgFormulaCodigoAFabricar.TextMatrix(iFila&, 3) = FORMATNUM$(CantidadComponente#, "F10.4")
        End If
    End If

End Sub


Private Sub mfgMaterialAsignado_Click()

    If mfgMaterialAsignado.Rows > 1 And mfgMaterialAsignado.MouseRow > 0 Then
        
        FilaElegida& = mfgMaterialAsignado.MouseRow

        For i = 1 To mfgMaterialAsignado.Rows - 1
            If mfgMaterialAsignado.TextMatrix(i, 0) = "*" Then
                Call COLOREAR_GRILLA_SOLAFILA(mfgMaterialAsignado, RGB(255, 255, 255), i)
                mfgMaterialAsignado.TextMatrix(i, 0) = ""
                Exit For
            End If
        Next i
        Call COLOREAR_GRILLA_SOLAFILA(mfgMaterialAsignado, RGB(241, 194, 50), FilaElegida&)
        mfgMaterialAsignado.TextMatrix(FilaElegida&, 0) = "*"
        
        If ModificaCantidadMaterial = True Then
            Call ModificarCantidadComponente(FilaElegida&)
        End If
    End If
    
End Sub

Private Sub mfgMaterialAsignado_DblClick()

    FilaElegida& = mfgMaterialAsignado.MouseRow
    mfgMaterialDisponible.Rows = mfgMaterialDisponible.Rows + 1
    
    mfgMaterialDisponible.RowHeight(mfgMaterialDisponible.Rows - 1) = mfgMaterialAsignado.RowHeight(FilaElegida)
    
    mfgMaterialDisponible.TextMatrix(mfgMaterialDisponible.Rows - 1, 1) = mfgMaterialAsignado.TextMatrix(FilaElegida, 1)
    mfgMaterialDisponible.TextMatrix(mfgMaterialDisponible.Rows - 1, 2) = mfgMaterialAsignado.TextMatrix(FilaElegida, 2)
    mfgMaterialDisponible.TextMatrix(mfgMaterialDisponible.Rows - 1, 3) = mfgMaterialAsignado.TextMatrix(FilaElegida, 3)
    mfgMaterialDisponible.TextMatrix(mfgMaterialDisponible.Rows - 1, 4) = mfgMaterialAsignado.TextMatrix(FilaElegida, 4)
    mfgMaterialDisponible.TextMatrix(mfgMaterialDisponible.Rows - 1, 5) = mfgMaterialAsignado.TextMatrix(FilaElegida, 5)
    mfgMaterialDisponible.TextMatrix(mfgMaterialDisponible.Rows - 1, 6) = mfgMaterialAsignado.TextMatrix(FilaElegida, 6)
    mfgMaterialDisponible.TextMatrix(mfgMaterialDisponible.Rows - 1, 7) = mfgMaterialAsignado.TextMatrix(FilaElegida, 7)
    
    mfgMaterialDisponible.ColSel = 2: mfgMaterialDisponible.COL = 1
    mfgMaterialDisponible.Sort = 1
    
    If mfgMaterialAsignado.Rows > 2 Then
        mfgMaterialAsignado.RemoveItem FilaElegida
    Else
        mfgMaterialAsignado.Rows = 1
    End If
    
    Asignado# = 0
    For i = 1 To mfgMaterialAsignado.Rows - 1
        Asignado# = Asignado# + XVALO(mfgMaterialAsignado.TextMatrix(i, 6))
    Next i
    
    txtCantidadAFabricar = TRIM$(FORMATNUM$(Asignado#, "F10.1"))
    
End Sub

Private Sub mfgMaterialAsignado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ModificaCantidadMaterial = False
    If Button = 2 Then
        ModificaCantidadMaterial = True
    End If
    
End Sub

Sub ModificarCantidadComponente(Fila As Long)
    
    SaldoDelLote$ = TRIM$(mfgMaterialAsignado.TextMatrix(Fila, 6))
    CantidadModificada# = XVALO(InputBox("Cantidad", "Cantidad a Asignar", SaldoDelLote$))
    
    If CantidadModificada# > 0.05 And CantidadModificada# < XVALO(SaldoDelLote$) Then
        mfgMaterialAsignado.TextMatrix(Fila, 6) = TRIM$(FORMATNUM$(CantidadModificada#, "F10.1"))
        Asignado# = 0
        For i = 1 To mfgMaterialAsignado.Rows - 1
            Asignado# = Asignado# + XVALO(mfgMaterialAsignado.TextMatrix(i, 6))
        Next i
        
        txtCantidadAFabricar = TRIM$(FORMATNUM$(Asignado#, "F10.1"))
    Else
        Call COMUNI("Cantidad Incorrecta.\ \Debe ser Mayor a Cero y Menor al Saldo del Lote")
    End If

End Sub

Private Sub mfgMaterialDisponible_Click()

    If mfgMaterialDisponible.Rows > 1 And mfgMaterialDisponible.MouseRow > 0 Then
        
        FilaElegida& = mfgMaterialDisponible.MouseRow

        For i = 1 To mfgMaterialDisponible.Rows - 1
            If mfgMaterialDisponible.TextMatrix(i, 0) = "*" Then
                Call COLOREAR_GRILLA_SOLAFILA(mfgMaterialDisponible, RGB(255, 255, 255), i)
                mfgMaterialDisponible.TextMatrix(i, 0) = ""
                Exit For
            End If
        Next i
        Call COLOREAR_GRILLA_SOLAFILA(mfgMaterialDisponible, RGB(241, 194, 50), FilaElegida&)
        mfgMaterialDisponible.TextMatrix(FilaElegida&, 0) = "*"
        
        Documento$ = mfgMaterialDisponible.TextMatrix(FilaElegida, 3)
        txtAnotaciones = AnotacionesEnClasificacion(Documento$)
        If txtAnotaciones = "" And Mid$(Documento$, 1, 2) = "OF" Then
            txtAnotaciones = SAFETEXT$(VALOBADA("anotmovi", "anotaciones", "doprilar='" & Documento$ & "'"))
        End If
        
    End If

End Sub

Private Sub mfgMaterialDisponible_DblClick()

    FilaElegida = mfgMaterialDisponible.MouseRow
    mfgMaterialAsignado.Rows = mfgMaterialAsignado.Rows + 1

    mfgMaterialAsignado.RowHeight(mfgMaterialAsignado.Rows - 1) = mfgMaterialDisponible.RowHeight(FilaElegida)

    mfgMaterialAsignado.TextMatrix(mfgMaterialAsignado.Rows - 1, 1) = mfgMaterialDisponible.TextMatrix(FilaElegida, 1)
    mfgMaterialAsignado.TextMatrix(mfgMaterialAsignado.Rows - 1, 2) = mfgMaterialDisponible.TextMatrix(FilaElegida, 2)
    mfgMaterialAsignado.TextMatrix(mfgMaterialAsignado.Rows - 1, 3) = mfgMaterialDisponible.TextMatrix(FilaElegida, 3)
    mfgMaterialAsignado.TextMatrix(mfgMaterialAsignado.Rows - 1, 4) = mfgMaterialDisponible.TextMatrix(FilaElegida, 4)
    mfgMaterialAsignado.TextMatrix(mfgMaterialAsignado.Rows - 1, 5) = mfgMaterialDisponible.TextMatrix(FilaElegida, 5)
    mfgMaterialAsignado.TextMatrix(mfgMaterialAsignado.Rows - 1, 6) = mfgMaterialDisponible.TextMatrix(FilaElegida, 6)
    mfgMaterialAsignado.TextMatrix(mfgMaterialAsignado.Rows - 1, 7) = mfgMaterialDisponible.TextMatrix(FilaElegida, 7)
    mfgMaterialAsignado.TextMatrix(mfgMaterialAsignado.Rows - 1, 8) = mfgMaterialDisponible.TextMatrix(FilaElegida, 7)

    If mfgMaterialDisponible.Rows > 2 Then
        mfgMaterialDisponible.RemoveItem FilaElegida
    Else
        mfgMaterialDisponible.Rows = 1
    End If

    Asignado# = 0
    For i = 1 To mfgMaterialAsignado.Rows - 1
        Asignado# = Asignado# + XVALO(mfgMaterialAsignado.TextMatrix(i, 7))
    Next i

    txtCantidadAFabricar = TRIM$(FORMATNUM$(Asignado#, "F10.1"))

    mfgMaterialAsignado.ColSel = 2: mfgMaterialAsignado.COL = 1
    mfgMaterialAsignado.Sort = 1
    
End Sub

Private Sub mfgMaterialDisponible_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 114 Then
        Call BuscadorDeTextoEnGrilla(txtBuscarTexto, UltimaFilaEncontrada + 1)
    End If
    
End Sub

Private Sub mfgMaterialDisponible_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        txtBuscarTexto = ""
        txtCodigoFamilia = ""
    End If
    
End Sub

Private Sub mfgMaterialDisponibleBuscar_DblClick()

    filaElegidaBuscar = mfgMaterialDisponibleBuscar.MouseRow
    mfgMaterialAsignado.Rows = mfgMaterialAsignado.Rows + 1
    
    mfgMaterialAsignado.TextMatrix(mfgMaterialAsignado.Rows - 1, 1) = mfgMaterialDisponibleBuscar.TextMatrix(filaElegidaBuscar, 1)
    mfgMaterialAsignado.TextMatrix(mfgMaterialAsignado.Rows - 1, 2) = mfgMaterialDisponibleBuscar.TextMatrix(filaElegidaBuscar, 2)
    mfgMaterialAsignado.TextMatrix(mfgMaterialAsignado.Rows - 1, 3) = mfgMaterialDisponibleBuscar.TextMatrix(filaElegidaBuscar, 3)
    mfgMaterialAsignado.TextMatrix(mfgMaterialAsignado.Rows - 1, 4) = mfgMaterialDisponibleBuscar.TextMatrix(filaElegidaBuscar, 4)
    mfgMaterialAsignado.TextMatrix(mfgMaterialAsignado.Rows - 1, 5) = mfgMaterialDisponibleBuscar.TextMatrix(filaElegidaBuscar, 5)
    mfgMaterialAsignado.TextMatrix(mfgMaterialAsignado.Rows - 1, 6) = mfgMaterialDisponibleBuscar.TextMatrix(filaElegidaBuscar, 6)
    FilaElegida = XVALO(mfgMaterialDisponibleBuscar.TextMatrix(filaElegidaBuscar, 7))
    
    If mfgMaterialDisponibleBuscar.Rows > 2 Then
        mfgMaterialDisponibleBuscar.RemoveItem filaElegidaBuscar
    Else
        mfgMaterialDisponibleBuscar.Rows = 1
    End If
    
    If mfgMaterialDisponible.Rows > 2 Then
        mfgMaterialDisponible.RemoveItem FilaElegida
    Else
        mfgMaterialDisponible.Rows = 1
    End If
    
    Asignado# = 0
    For i = 1 To mfgMaterialAsignado.Rows - 1
        Asignado# = Asignado# + XVALO(mfgMaterialAsignado.TextMatrix(i, 6))
    Next i
    
    txtCantidadAFabricar = TRIM$(FORMATNUM$(Asignado#, "F10.1"))
    
    mfgMaterialAsignado.ColSel = 2: mfgMaterialAsignado.COL = 1
    mfgMaterialAsignado.Sort = 1

End Sub

Private Sub mfgMaterialDisponibleBuscar_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        mfgMaterialDisponible.Visible = True
        mfgMaterialDisponibleBuscar.Visible = False
        txtBuscarTexto = ""
        txtCodigoFamilia = ""
    End If
    
End Sub


Private Sub mnuMaestroArticulos_Click()

    Call CONECRUN("ARCHIG07/AMASTO", "Padron Maestro de Articulos")
    
End Sub

Private Sub mnuTablaFamiliaComponentes_Click()

    Dim da As New DatosAuxiliares
    
    Call da.MostrarTablaAuxilar("FAMCOMPO")
    
End Sub







Private Sub txtBuscarTexto_Change()

    UltimaFilaEncontrada = 0
    
End Sub

Private Sub txtBuscarTexto_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 114 Then
        Call BuscadorDeTextoEnGrilla(txtBuscarTexto, UltimaFilaEncontrada + 1)
    End If
    
End Sub


Private Sub txtBuscarTexto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        mfgMaterialDisponible.Visible = True
        txtBuscarTexto = ""
    ElseIf KeyAscii = 13 Then
        Call BuscadorDeTextoEnGrilla(txtBuscarTexto)
    End If
    
End Sub




Private Sub txtCodigo_Change()

    CodigoInterno% = CODINT%(TXTCODIGO)
    If CodigoInterno% = 0 Then
        TXTDESCRIPCION = ""
        mfgFormulaCodigoAFabricar.Rows = 1
        Exit Sub
    Else
        TXTDESCRIPCION = DESCRIT0$(CodigoInterno%)
    End If
    
    Dim Rst As ADODB.Recordset
    ConsultaSql$ = "select codielem, usobruto from formulas with(nolock) where codiconj=" & CStr(CODINT%(TXTCODIGO))
    Set Rst = READSET(ConsultaSql$)
    
    With Rst
        mfgFormulaCodigoAFabricar.Rows = 1
        Do While Not .EOF
            Codigo$ = CODEXT$(!codielem)
            Descripcion$ = DESCRIT0$(!codielem)
            USO# = XVALO(!usobruto)
        
            mfgFormulaCodigoAFabricar.Rows = mfgFormulaCodigoAFabricar.Rows + 1
            mfgFormulaCodigoAFabricar.TextMatrix(mfgFormulaCodigoAFabricar.Rows - 1, 1) = Codigo$
            mfgFormulaCodigoAFabricar.TextMatrix(mfgFormulaCodigoAFabricar.Rows - 1, 2) = Descripcion$
            mfgFormulaCodigoAFabricar.TextMatrix(mfgFormulaCodigoAFabricar.Rows - 1, 3) = FORMATNUM$(USO#, "F10.1")
            
            If CantRegistros("datasql", "lista='FAMCOMPO' and campo1='" & Codigo$ & "'") > 0 Then
                mfgFormulaCodigoAFabricar.TextMatrix(mfgFormulaCodigoAFabricar.Rows - 1, 4) = "SI"
            End If
            
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0

End Sub

Private Sub txtCodigoFamilia_Change()

    txtBuscarTexto = ""
    
    Call ActualizaListaDeMaterialDisponible(txtCodigoFamilia)

End Sub

Private Sub txtCodigoFamilia_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        txtCodigoFamilia = ""
    End If
    
End Sub


Private Sub txtCodigoProductoTerminado_Change()

    CodigoInterno% = CODINT%(txtCodigoProductoTerminado)
    If CodigoInterno% = 0 Then
        txtDescripcionProductoTerminado = ""
        txtSubcodigo = ""
    Else
        txtDescripcionProductoTerminado = DESCRIT0$(CodigoInterno%)
        txtSubcodigo = SAFETEXT$(VALOBADA("master", "cod_interna", "codint=" & CStr(CodigoInterno%)))
        txtPesoBolsa = SAFETEXT$(VALOBADA("master", "talle", "codint=" & CStr(CodigoInterno%)))
    End If
    
End Sub

Private Sub txtNroCliente_Change()

    txtRazonSocial = ""
    If CantRegistros("deudor12", "nume_cli=" & txtNroCliente) > 0 Then
        NumeroCliente& = XVALO(txtCodigoProductoTerminado)
        txtRazonSocial = rasocli$(NumeroCliente&)
    End If
    
End Sub

Private Sub txtOrdenFabricacion_Change()

    CodigoAFabricarInterno% = XVALO(VALOBADA("ordefabr", "cod_inte", "idsubor='" & txtOrdenFabricacion & "'"))
    TXTCODIGO = CODEXT$(CodigoAFabricarInterno%)
    TXTDESCRIPCION = DESCRIT0$(CodigoAFabricarInterno%)
    
End Sub





