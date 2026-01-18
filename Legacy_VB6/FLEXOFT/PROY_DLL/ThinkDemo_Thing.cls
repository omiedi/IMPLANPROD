VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Thing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Name As String
Dim AA1

' Para almacenar el valor de la propiedad DebugID.
Private mlngDebugID As Long

Public Property Get DebugID() As Long
    DebugID = mlngDebugID
End Property

Public Sub ReverseName()
    Dim intCt As Integer
    Dim strNew As String
    For intCt = 1 To Len(Name)
        strNew = Mid$(Name, intCt, 1) & strNew
    Next
    Name = strNew
End Sub

Private Sub Class_Initialize()
    ' Obtiene un identificador de depuración que
    '   puede devolver la propiedad de sólo lectura
    '   DebugID.
    mlngDebugID = GetDebugID
    Debug.Print "Initialize Thing " & DebugID _
        & ", Name=" & Name

End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Debug.Print "Terminate Thing " & DebugID _
        & ", Name=" & Name
End Sub

Public Function existaA()
   Randomize Timer
   AA1 = Int(10 * Rnd + 0.5)
   existaA = AA1
End Function
