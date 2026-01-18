VERSION 5.00
Begin VB.Form INIMAS07 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Items de Stock"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5130
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   2  'Horizontal Line
      Height          =   3105
      Left            =   0
      Picture         =   "INIMAS07.frx":0000
      ScaleHeight     =   3105
      ScaleMode       =   0  'User
      ScaleWidth      =   2725.888
      TabIndex        =   0
      Top             =   0
      Width           =   5370
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Maestro de Items de Stock "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   2835
         TabIndex        =   1
         Top             =   210
         Width           =   2205
      End
   End
End
Attribute VB_Name = "INIMAS07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   Call CONAPLI(APLINVO$, TERMINA%) ' Busca la aplicacion a mostrar
   If TERMINA% > 0 Then Call FINAL("")
   '
   '
   UTILIZA_SKIN% = 1
   AMASTOK07.Show
   Hide
   '
End Sub

Sub CONAPLI(APLINVO$, TERMINA%)
   TERMINA% = 1
   Select Case APLINVO$    'Realiza la llamada al objeto segun la opcion seleccionada
       Case "CONFIGURA"
           Call AMASTOK07.Command37_Click       'Configuración General Maestro Items Stock
       Case "TATIMAT"
           Call OPMAST07.Option1_DblClick(0)    'Tabla de Tipo de Material
       Case "TAGRUCOVEN"
           Call OPMAST07.Option1_DblClick(3)     'Tabla Grupos Contables Ventas
       Case "TAGRUVENIVA"
           Call OPMAST07.Option1_DblClick(4)    'Tabla Grupos Ventas IVA
       Case "TAUNIMED"
           Call OPMAST07.Option1_DblClick(9)    'Tabla Unidades de Medida
       Case "TADEPOSI"
           Call OPMAST07.Option1_DblClick(5)    'Tabla de Depositos y Almacenes
       Case "TASECENCO"
           Call OPMAST07.Option1_DblClick(6)    'Tabla de Secciones y Centros de Costos
       Case "TAMAQEQUI"
           Call OPMAST07.Option1_DblClick(7)    'Tabla de Maquinas y Equipos de Producción
       Case "TACATEMPL"
           Call OPMAST07.Option1_DblClick(8)    'Tabla de Categorias de Empleados Productivos
       Case "CONLISTA"
           Call AMASTOK07.Command38_Click       'Consultas, Listados y Complementos
       Case "LISATRIGE"
           Call OPLIMAS07.Command9_Click        'Listado de Atributos Generales
       Case "STMINLOREP"
           Call OPLIMAS07.Command7_Click        'Stock Minimo y lote de Reposición
       Case "LISCAPINV"
           Call OPLIMAS07.Command15_Click       'Listado de Captura Para Inventarios
       Case "COSREPROV"
           Call OPLIMAS07.Command5_Click        'Listado de Costo de Reposicion por Proveedor
       Case "COSREPART"
           Call OPLIMAS07.Command4_Click        'Listado de Costo de Reposición por Articulo
       Case "RESTITEM"
           Call OPLIMAS07.Command3_Click        'Consulta y Restauracion de Items de Baja
       Case "ITEMINAC"
           Call OPLIMAS07.Command2_Click        'Listado de Items Inactivos
       Case "MULTINIVEL"
           Call AMASTOK07.Command40_Click(4)    'Muestra estructura Multinivel
       Case "IMPLOSION"
           Call AMASTOK07.Command40_Click(6)    'Conjuntos a los que Pertenece Item Activo
       Case "LISCOMPRA"
           Call AMASTOK07.Command40_Click(2)    'Lista de Compras
       Case "LISMULTINI"
           Call AMASTOK07.Command40_Click(1)    'Lista de Materiales Multinivel
       Case "LISENLACE"
           Call AMASTOK07.Command40_Click(0)    'Lista de Estructuras y Enlaces
       Case "LISECOPER"
           Call AMASTOK07.Command42_Click       'Listado de Secuencias de Operaciones
       Case "ALTACODIGO"
           TERMINA% = 0
           Call AMASTOK07.Command21_Click       'ALTA DE NUEVO CODIGO
       Case Else
           TERMINA% = 0
   End Select
End Sub





