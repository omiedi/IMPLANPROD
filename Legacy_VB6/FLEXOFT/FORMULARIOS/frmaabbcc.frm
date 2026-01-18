VERSION 5.00
Begin VB.Form frmaabbcc 
   Caption         =   "'forecast 27 de aug to dec 2 (2)"
   ClientHeight    =   2310
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   5520
   LinkTopic       =   "Form3"
   ScaleHeight     =   2310
   ScaleWidth      =   5520
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4440
      TabIndex        =   14
      Top             =   1660
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Renovar"
      Height          =   300
      Left            =   3360
      TabIndex        =   13
      Top             =   1660
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Actuali&zar"
      Height          =   300
      Left            =   2280
      TabIndex        =   12
      Top             =   1660
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   1200
      TabIndex        =   11
      Top             =   1660
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   120
      TabIndex        =   10
      Top             =   1660
      Width           =   975
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Excel 5.0;"
      DatabaseName    =   "C:\CLIENTES\SABO_WIN\TIPIFICA\Programa Kaco  2004.xls"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "'forecast 27 de aug to dec 2 (2)'$"
      Top             =   1965
      Width           =   5520
   End
   Begin VB.TextBox txtFields 
      DataField       =   "F7"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "F5"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1000
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "F4"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   680
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "F3"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "F1"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   40
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      Caption         =   "F7:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "F5:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "F4:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "F3:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "F1:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmaabbcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
  Data1.Recordset.AddNew
End Sub

Private Sub cmdDelete_Click()
  'esto puede generar un error si elimina el último
  'registro o el único registro del recordset
  Data1.Recordset.Delete
  Data1.Recordset.MoveNext
End Sub

Private Sub cmdRefresh_Click()
  'esto sólo es necesario para una aplicación de múltiples usuarios
  Data1.Refresh
End Sub

Private Sub cmdUpdate_Click()
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Data1_Error(DataErr As Integer, Response As Integer)
  'Aquí es dónde se sitúa el código de tratamiento de errores
  'Si desea ignorar los errores, comente la línea siguiente
  'Si desea tratarlos, agregue código para controlarlos
  MsgBox "El evento de error de datos generó el error:" & Error$(DataErr)
  Response = 0  'desprecia el error
End Sub

Private Sub Data1_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  'Esto mostrará la posición actual del registro
  'para dynasets y snapshots
  Data1.Caption = "Registro: " & (Data1.Recordset.AbsolutePosition + 1)
  'para el objeto table debe establecer la propiedad index cuando
  'se cree el recordset y usar la siguiente línea
  'Data1.Caption = "Record: " & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
  'Aquí es donde se pondría el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
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
  End Select
  Screen.MousePointer = vbHourglass
End Sub

