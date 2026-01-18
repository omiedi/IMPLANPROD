VERSION 5.00
Begin VB.Form frmforcarcli 
   Caption         =   "clientes"
   ClientHeight    =   3495
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   5520
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2040
      TabIndex        =   11
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4440
      TabIndex        =   10
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Renovar"
      Height          =   300
      Left            =   3360
      TabIndex        =   9
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Actuali&zar"
      Height          =   300
      Left            =   2280
      TabIndex        =   8
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   1200
      TabIndex        =   7
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   975
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   "C:\CLIENTES\SABO_WIN\PRUEBA EN ACCESS.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "clientes"
      Top             =   3150
      Width           =   5520
   End
   Begin VB.TextBox txtFields 
      DataField       =   "numero de cuit"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   2040
      MaxLength       =   16
      TabIndex        =   5
      Top             =   680
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "razon social"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   2040
      MaxLength       =   30
      TabIndex        =   3
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "numero de cliente"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   40
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      Caption         =   "numero de cuit:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "razon social:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "numero de cliente:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmforcarcli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub AddNewX()

    Dim dbclientes As Database
    Dim rstEmpleados As Recordset
    Dim nucliente As Integer
    Dim rasocli As String
    Dim STRNOMBRE As String
    Dim STRAPELLIDOS As String

    Set dbclientes = OpenDatabase("C:\CLIENTES\SABO_WIN\PRUEBA EN ACCESS.mdb")
    Set rstEmpleados = _
        dbclientes.OpenRecordset("CLIENTES", dbOpenDynaset)

    ' Obtiene datos del usuario.
    STRNOMBRE = Trim(InputBox( _
        "Introduzca el nombre:"))
    STRAPELLIDOS = Trim(InputBox( _
        "Introduzca los apellidos:"))

    ' Sólo se ejecuta si el usuario escribe algo

' en los dos campos.
    If STRNOMBRE <> "" And STRAPELLIDOS <> "" Then

        ' Llama a la función que agrega el registro.
        AgregarNombre rstEmpleados, STRNOMBRE, STRAPELLIDOS

        ' Muestra los datos agregados más recientemente.
        With rstEmpleados
            Debug.Print "Registro nuevo: " & !NCLIENTE & _
                " " & !RASOCIAL
            ' Elimina el registro nuevo porque esto es un ejemplo.
            '.Delete
        End With

    Else
        Debug.Print _
            "¡Debe escribir una cadena para el nombre y los apellidos!"

End If

    rstEmpleados.Close
    dbclientes.Close

End Sub

Function AgregarNombre(rstTemp As Recordset, _
    strNomrbe As String, STRAPELLIDOS As String)

    ' Agrega un registro nuevo al Recordset utilizando
    ' datos transferidos del procedimiento que llama.
    ' El registro nuevo pasa a ser el registro actual.
    With rstTemp
        .AddNew
        !NCLIENTE = Val(strNomrbe)
        !RASOCIAL = STRAPELLIDOS
        .Update
        .Bookmark = .LastModified
    End With

End Function










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

Private Sub Command1_Click()
  Call AddNewX
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

Private Sub Form_Load()
MsgBox 1
End Sub
