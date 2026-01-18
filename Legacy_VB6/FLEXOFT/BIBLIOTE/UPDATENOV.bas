Attribute VB_Name = "UPDATENOV"
Private Const MAX_PATHLEN = &H104
Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
exeFilename(1 To MAX_PATHLEN) As Byte
End Type

Private Declare Function CloseHandle Lib "kernel32" (hObject As Long) As Boolean

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" _
(ByVal dwFlags As Long, ByVal dprocess As Long) As Long

Private Declare Function Process32First Lib "kernel32" _
(ByVal hSnapshot As Long, pProcessEntry As PROCESSENTRY32) As Long

Private Declare Function Process32Next Lib "kernel32" _
(ByVal hSnapshot As Long, pProcessEntry As PROCESSENTRY32) As Long

Private Const TH32CS_SNAPPROCESS = &H2&

Function VerificaInstalacionDelFramework(txtClave As String)
    ' Leer todas los valores y subclaves de la clave indicada       (22/Nov/00)
    '
    ' Mostrar en el TreeView las claves
    ' y en el ListView los nombres y valores
    '
    Dim tQR As cQueryReg
    Dim ma_scolKeys() As String
    
    Dim i As Long, j As Long, ii As Long
    Dim sID As String
    Dim sID2 As String
    Dim asSubClaves() As String
    Dim encontroClave As Integer
    Dim sKey As String
    '
    Set tQR = New cQueryReg
    ' Si tiene el separador del final, quitárselo                   (23/Nov/00)
    sKey = Trim(txtClave)
    sKey = Trim$(sKey)
    If Right$(sKey, 1) = "\" Then
        sKey = Left$(sKey, Len(sKey) - 1)
    End If
    '
    encontroClave = 0
    ' Enumerar los datos de la clave
    If tQR.EnumValues(ma_scolKeys(), sKey) Then
        sID = sKey
        'LeerValores sKey
        '
        ' Mostrar las subclaves de la clave indicada
        If tQR.EnumValues(ma_scolKeys(), sKey) Then
            j = UBound(ma_scolKeys)
            For i = 1 To j
                sID2 = sID & "\" & ma_scolKeys(i)
                ' Lee las subclaves
                If tQR.EnumKeys(asSubClaves(), sKey) Then
                    For ii = 1 To UBound(asSubClaves)
                        'Debug.Print asSubClaves(ii)
                        If InStr(asSubClaves(ii), "v2.") > 0 Then
                            encontroClave = 1
                            Exit For
                        End If
                    Next
                End If
                If encontroClave = 1 Then Exit For
            Next
        End If
    End If
    '
    VerificaInstalacionDelFramework = encontroClave
    DoEvents
    '
End Function

Function EstaelProceso(ByVal abuscar As String) As Boolean

Dim SnapHandle, f As Long
Dim Process As PROCESSENTRY32

EstaelProceso = False

SnapHandle = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)

If SnapHandle = -1 Then Exit Function

Process.dwSize = Len(Process)
f = Process32First(SnapHandle, Process)

Do While f <> 0
If StrComp(Trim(abuscar), strZtostr(Process.exeFilename), vbTextCompare) = 0 Then
EstaelProceso = True
Exit Do
End If

Process.dwSize = Len(Process)
f = Process32Next(SnapHandle, Process)
Loop

Call CloseHandle(SnapHandle)

End Function

Function strZtostr(Fuente() As Byte) As String

Dim i&, cb As Byte
strZtostr = ""
Do While i < MAX_PATHLEN
i = i + 1
cb = Fuente(i)
If cb = 0 Then Exit Function 'Return
strZtostr = strZtostr & Chr$(cb)
Loop

End Function

Function ExisteArchivo(nombreArchivo As String) As Boolean

ExisteArchivo = False
nombreArchivoB = Trim(nombreArchivo)
If nombreArchivoB = "" Then Exit Function

FileName = Dir(App.Path & "\" & nombreArchivoB)
If FileName <> "" Then ExisteArchivo = True

End Function
