VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmMaster 
   Caption         =   "Master"
   ClientHeight    =   3645
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   5520
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5520
      TabIndex        =   11
      Top             =   3045
      Width           =   5520
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   300
         Left            =   4505
         TabIndex        =   16
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "Actuali&zar"
         Height          =   300
         Left            =   3409
         TabIndex        =   15
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdRenovar 
         Caption         =   "&Renovar"
         Height          =   300
         Left            =   2313
         TabIndex        =   14
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   300
         Left            =   1217
         TabIndex        =   13
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   300
         Left            =   121
         TabIndex        =   12
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Data datSecondaryRS 
      Connect         =   ""
      DatabaseName    =   "C:\VbAccess\BdAccess\ARTICULOS.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2190
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1695
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data datPrimaryRS 
      Align           =   2  'Align Bottom
      Caption         =   " "
      Connect         =   ""
      DatabaseName    =   "C:\VbAccess\BdAccess\ARTICULOS.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from [Master]"
      Top             =   3345
      Width           =   5520
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Provhabi"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   1340
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Stomax"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1020
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Umedida"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   2040
      MaxLength       =   4
      TabIndex        =   5
      Top             =   700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Descripcion"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2040
      MaxLength       =   64
      TabIndex        =   3
      Top             =   380
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Codigo"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   2040
      MaxLength       =   16
      TabIndex        =   1
      Top             =   60
      Width           =   3375
   End
   Begin MSDBGrid.DBGrid grdDataGrid 
      Bindings        =   "frmMaster.frx":0000
      Height          =   1300
      Left            =   0
      OleObjectBlob   =   "frmMaster.frx":0182
      TabIndex        =   10
      Top             =   1660
      Width           =   5540
   End
   Begin VB.Label lblLabels 
      Caption         =   "Provhabi:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Stomax:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Umedida:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripcion:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAgregar_Click()
  datPrimaryRS.Recordset.AddNew
End Sub

Private Sub cmdEliminar_Click()
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
End Sub

Private Sub cmdRenovar_Click()
  'Esto sólo es necesario en aplicaciones con múltiples usuarios
  datPrimaryRS.Refresh
End Sub

Private Sub cmdActualizar_Click()
  datPrimaryRS.UpdateRecord
  datPrimaryRS.Recordset.Bookmark = datPrimaryRS.Recordset.LastModified
End Sub

Private Sub cmdCerrar_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub datPrimaryRS_Error(DataErr As Integer, Response As Integer)
  'Aquí es dónde puede situar el código de tratamiento de error
  'Si desea ignorar los errores, quite el comentario de la siguiente línea
  'Si desea capturarlos, agregue código aquí para controlarlos
  MsgBox "Data error event hit err:" & Error$(DataErr)
  Response = 0  'Despreciar el error
End Sub

Private Sub datPrimaryRS_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  'Sincronizar la cuadrícula con el conjunto de registros principal
  datSecondaryRS.RecordSource = "select [CodXElem],[UsoBruto] from [Formulas] where [CodiConj]='" & datPrimaryRS.Recordset![Codigo] & "'"
  datSecondaryRS.Refresh
  'Mostrar la posición actual de registro para dynasets y snapshots
  datPrimaryRS.Caption = "Record: " & (datPrimaryRS.Recordset.AbsolutePosition + 1)
End Sub

Private Sub datPrimaryRS_Validate(Action As Integer, Save As Integer)
  'Aquí se sitúa el código de validación
  'Este evento se invoca cuando ocurre la siguiente acción
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
      Screen.MousePointer = vbDefault
  End Select
  Screen.MousePointer = vbHourglass
End Sub

Private Sub Form_Load()
  'Crea el conjunto de resultados de la cuadrícula
  datPrimaryRS.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'Esto redimensionará la cuadrícula al cambiar el tamaño del formulario
  grdDataGrid.Width = Me.ScaleWidth
  grdDataGrid.Height = Me.ScaleHeight - grdDataGrid.Top - datPrimaryRS.Height - picButtons.Height - 30
End Sub

