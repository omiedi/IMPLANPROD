VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQueryReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' cQueryReg                                                         (13/Ago/98)
' Clase para obtener valores del Registro del Sistema
'
' Revisión 0.01 (18/Ago/98) Funciones de crear/borrar claves/valores
' Revisión 0.02 (12/Oct/98) Nueva función para enumeración de claves
' Revisión 0.03 (15/Oct/98) Importar/Exportar claves del registro
'                           (sólo exportar...)
' Revisión 0.04 (16/Dic/98) Modificadas las funciones de obtener los
'                           directorios del sistema,
'                           no se ha cambiado la forma de usarlas.
'
' Revisión 0.10 (12/Jun/99) Probado en Windows 2000 Professional Beta 3
'                           y parece que funciona bien...
'
' Revisión 0.11 (05/Jul/99) Nuevo método para des-registrar un servidor ActiveX
'                           Métodos para saber el Clsid y TypeLib de una clase
'
' Revisión 0.12 (04/Ago/99) Nuevos métodos para crear y comprobar claves
' Revisión 0.13 (22/Ago/99) Correcciones para Windows 2000
' Revisión 0.14 (22/Jun/00) En ShellFolders se tiene en cuenta el usuario actual
' Revisión 0.15 (22/Nov/00) Todas las asignaciones de cadenas con String$(n,0)
'                           se sustituyen por String$(n,Chr$(0)), no altera nada
'                           pero es más evidente su función.
'          0.16 (22/Nov/00) Los valores Binarios se formatean como Hexadecimales
'                           teniendo en cuenta las letras, no sólo los números.
'                           En lugar de usar Format$, uso Right$(...
'          0.17 (23/Nov/00) En ParseKey se comprueba que la clave no acabe en \
'          0.18 (23/Nov/00) En EnumValues se añade el valor de Default
'          0.19 (09/Feb/01) Nuevas comprobaciones al leer los datos.
'          0.20 (19/May/01) Corrección de un bug al usar "open" en AsociarExt
'          0.21 (19/May/01) Nuevo método: AsociarIcono
'          0.22 (26/May/01) Nueva propiedad de sólo lectura: UserName
'          0.23 (28/Dic/01) Nuevas propiedades y algunos cambios:
'                           GetRegType para saber el tipo de datos
'                           EnumValuesByType, enumera valores sólo del tipo indicado
'                           DeleteValue, borrar un valor, también se hace con DeleteKey
'                           Cambios en GetRegString y SetReg para usar el tipo REG_EXPAND_SZ
'
' La información para crear las funciones están tomadas de ejemplos
' y valores obtenidos en el código del Setup1.vbp
' y de artículos incluidos en los CDs del MSDN Library.
'
' ¡¡¡ De algún sitio tenía que sacar la información... !!!
'
' ©Guillermo 'guille' Som, 1998-2001 <guille@costasol.net>
'
'------------------------------------------------------------------------------
'   Método          Descripción
'   ------          -----------
'   AsociarExt      Asociar una extensión con un programa
'                   También sirve para añadir comandos a extensiones existentes
'   AsociarIcono    Asociar un icono a la extensión indicada
'   ClassCLSID      Devuelve el CLSID de la clase indicada
'   ClassTypeLib    Devuelve el TypeLib de la clase indicada
'   CloseKey        Cierra la clave abierta usando el handle pasado como parámetro
'   DeleteKey       Borra la clave o el valor especificado
'   DeleteKeyNT     Borra la clave especificada y sus subclaves y valores, para usar con Windows NT y Windows 98
'   DeleteKeyWin95  En Windows 95, borra la clave especificada y sus subclaves y valores. En Windows NT y Windows 98 no funcionará si la clave indicada tiene subclaves.
'   DesasociarExt   Desasociar la extensión (la borra del registro)
'   EnumKeys        Enumera todas las subclaves de la clave indicada y las devuelve en un array de tipo String que se pasa como parámetro.
'   EnumValues      Enumera todos los valores de la clave indicada y las devuelve en un array de tipo String que se pasa como parámetro.
'   GetFolder       Devuelve el path de la carpeta "especial" del sistema.
'                   El parámetro espera un nombre del tipo de carpeta a obtener,
'                   ver Nombres de directorios del sistema
'   GetReg          Obtener un valor, de cualquier tipo, de una entrada del registro
'   GetRegBinary    Obtener un valor binario de una entrada del registro
'   GetRegDWord     Obtener un valor DWORD de una entrada del registro
'   GetRegString    Obtener un valor cadena de una entrada del registro
'   QueryRegBase    Busca una entrada en el registro
'   RegSaveKey      Guarda en un fichero el contenido de una clave, las subclaves y datos.
'                   El formato no es ASCII, es un formato "propio" que se puede usar con la función RegLoadKey (aún no implementada)
'   RTrimZero       Devuelve una cadena hasta el primer Chr$(0)
'   SetReg          Asigna un valor de cualquier tipo a la clave indicada.
'   ShellFolders    Devuelve una colección con los Nombres de directorios del sistema
'                   o los paths a los que hacen referencia.
'                   Dependerá del parámetro pasado,
'                   por defecto False para devolver los paths
'   UnRegister      Des-registrar un servidor ActiveX.              (05/Jul/99)
'                   Borra el CLSID, el TypeLib y el nombre de la clase
'
' Estas funciones simplemente llaman a la función del API de windows, (el nombre de la función del API empieza con Reg, salvo que se indique lo contrario)
'   EnumKeyEx
'   EnumValue       RegEnumValue, usarla para tipos diferentes de String
'   EnumValueString RegEnumValue, usarla para tipos String
'   OpenKeyEx
'   OpenKeyQuery    RegOpenKeyEx, abre una clave para consultar información
'   QueryInfoKey
'   RegSetValue2    RegSetValueEx, pero para usar sólo con cadenas.
'                   Esta función asigna el valor por defecto de la clave indicada.
'                   Es decir, el "Predeterminado" o "Default".
'------------------------------------------------------------------------------
' La función / método GetReg devolverá el valor adecuado,
' si se ha encontrado la clave especificada en el registro.
'
' Las funciones GetRegXXX se usarán para asegurarnos que el valor
' devuelto es del tipo especificado.
' Por ejemplo:
'   si se usa GetRegString y el valor de la clave indicada
'   no es del tipo cadena, se devolverá una cadena vacía
'========================================================================
'=== NOTA: Las he dejado para poder ver cómo se usarían según el tipo ===
'========================================================================
'------------------------------------------------------------------------------
Option Explicit

Private colShellFolders As Collection
Private colShellFoldersKey As Collection

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long

' Registry manipulation API's (32-bit)

' Claves del Registro
Public Enum eHKEY
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    '
    HKEY_PERFORMANCE_DATA = &H80000004  ' Sólo para NT
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
    '
    HKEY_FIRST = HKEY_CLASSES_ROOT
    HKEY_LAST = HKEY_DYN_DATA
End Enum
'
'   HKEY_CLASSES_ROOT es un duplicado de HKEY_LOCAL_MACHINE\Software\Classes
'   HKEY_CURRENT_USER es un duplicado de HKEY_USERS\[Usuario]
'
'
Public Enum eHKEYError
    ERROR_SUCCESS = 0               'Todo correcto, sin error
    ERROR_NONE = 0                  '       "           "
                                    'The configuration registry...
    'ERROR_BADDB = 1                 'database is corrupt
                '1009&
    'ERROR_BADKEY = 2                'key is invalid
                '1010&
        'También declarada como:
    ERROR_FILE_NOT_FOUND = 2&       'este error ocurre cuando se abre
                                    'una clave y no existe
    
    'ERROR_CANTOPEN = 3              'key could not be opened
                '1011&
    'ERROR_CANTREAD = 4              'key could not be read
                '1012&
    'ERROR_CANTWRITE = 5             'key could not be written
                '1013&
        'También declarada como:
    ERROR_ACCESS_DENIED = 5&
    
    ERROR_OUTOFMEMORY = 6&          '
    ERROR_INVALID_PARAMETER = 7&    '
    'ERROR_ACCESS_DENIED = 8&        '
    ERROR_INVALID_PARAMETERS = 87&  '
    '
    ERROR_MORE_DATA = 234&          'More data is available
    ERROR_NO_MORE_ITEMS = 259&      'No more data is available
    
    ERROR_BADKEY = 1010&            'Se produce cuando se intenta acceder
                                    'a una clave que no está abierta
    'KEY_ALL_ACCESS = &H3F           '
    'REG_OPTION_NON_VOLATILE = 0
End Enum
'
' Los tipos de datos posibles, algunos sólo para Windows NT
Public Enum eHKEYDataType
    REG_NONE = 0&               'No value type
    REG_SZ = 1&                 'Unicode null terminated string
    REG_EXPAND_SZ = 2           'Unicode null terminated string
                                '(with environment variable references)
    REG_BINARY = 3              'Free form binary
    REG_DWORD = 4               '32-bit number
    REG_DWORD_LITTLE_ENDIAN = 4 '32-bit number (same as REG_DWORD)
    REG_DWORD_BIG_ENDIAN = 5    '32-bit number
    REG_LINK = 6                'Symbolic Link (unicode)
    REG_MULTI_SZ = 7            'Multiple Unicode strings
    REG_RESOURCE_LIST = 8       'Resource list in the resource map
    REG_FULL_RESOURCE_DESCRIPTOR = 9    'Resource list in the hardware description
    REG_RESOURCE_REQUIREMENTS_LIST = 10
End Enum

' Standard rights, used later below
Const SYNCHRONIZE = &H100000
Const READ_CONTROL = &H20000
Const STANDARD_RIGHTS_ALL = &H1F0000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)

' Security Access Mask
Public Enum eREGSAM
                                    'Permission to:
    KEY_QUERY_VALUE = &H1           '   query subkey data
    KEY_SET_VALUE = &H2             '   set subkey data
    KEY_CREATE_SUB_KEY = &H4        '   create subkeys
    KEY_ENUMERATE_SUB_KEYS = &H8    '   enumerate subkeys
    KEY_NOTIFY = &H10               '   for change notification
    KEY_CREATE_LINK = &H20          '   create a symbolic link

    'KEY_READ Combination of:
    '           KEY_QUERY_VALUE, KEY_ENUMERATE_SUB_KEYS, and
    '           KEY_NOTIFY access.
    KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
    'KEY_WRITE Combination of:
    '           KEY_SET_VALUE and KEY_CREATE_SUB_KEY access.
    KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
                                    
                                    'Permission for read access
    KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
    
    'KEY_ALL_ACCESS Combination of:
    '           KEY_QUERY_VALUE, KEY_SET_VALUE, KEY_CREATE_SUB_KEY,
    '           KEY_ENUMERATE_SUB_KEYS, KEY_NOTIFY and KEY_CREATE_LINK access.
    KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
    
    '#define DELETE                           (0x00010000L)
    'KEY_DELETE = &H10000
End Enum

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" _
    (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, _
    ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, _
    lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, _
    lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, _
    lpftLastWriteTime As FILETIME) As Long
'Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
    (ByVal hKey As Long, ByVal lpszSubKey As String, _
    phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, _
    phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal lpValueName As String, lpcbValueName As Long, _
    ByVal lpReserved As Long, lpType As Long, lpData As Any, _
    lpcbData As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
    (ByVal hKey As Long, ByVal lpszSubKey As String, _
    phkResult As Long) As Long
'
'Windows 95:
'   The RegDeleteKey function deletes a subkey and all its descendants.
'Windows NT:
'   The RegDeleteKey function deletes the specified subkey.
'   The subkey to be deleted must not have subkeys.
'
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
    (ByVal hKey As Long, ByVal lpszSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
    (ByVal hKey As Long, ByVal szValueName As String) As Long

Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" _
    (ByVal hKey As Long, ByVal iSubKey As Long, _
    ByVal lpszName As String, ByVal cchName As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal lpName As String, lpcbName As Long, _
    ByVal lpReserved As Long, ByVal lpClass As String, _
    lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal lpValue As String, lpcbValue As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpszValueName As String, _
    ByVal dwReserved As Long, lpdwType As Long, _
    lpbData As Any, cbData As Long) As Long

' The RegSetValue function sets the data for the default or unnamed
' value of a specified registry key. The data must be a text string.
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal dwType As Long, ByVal lpData As String, _
    ByVal cbData As Long) As Long

'
' The RegSetValueEx function sets the data and type of a
' specified value under a registry key.
'
'lpValueName:
' Pointer to a string containing the name of the value to set.
' If a value with this name is not already present in the key,
' the function adds it to the key.
' If lpValueName is NULL or an empty string, "", the function sets
' the type and data for the key's unnamed or default value.
'
'On Windows 95, the type of a key's default value is always REG_SZ,
' so the dwType parameter must specify REG_SZ for an unnamed value.
'On Windows 98, an unnamed value can be of any type.
'
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, ByVal lpszValueName As String, _
    ByVal dwReserved As Long, ByVal fdwType As Long, _
    lpbData As Any, ByVal cbData As Long) As Long

'
' Funciones del API para guardar y recuperar información del registro.
'
'Private Type SECURITY_ATTRIBUTES
'    nLength As Long
'    lpSecurityDescriptor As Long
'    bInheritHandle As Long
'End Type
'
' RegSaveKey:
' El nombre guardado en Windows 95 sólo permite nombres cortos,
' si no se especifica el path se guardará en el directorio del Windows.
' Además se guardará con los atributos Hidden, Read-Only y System
'
'Private Declare Function RegSaveKeyA Lib "advapi32.dll" _
    (ByVal hKey As Long, ByVal lpFile As String, _
    lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function RegSaveKeyA Lib "advapi32.dll" _
    (ByVal hKey As Long, ByVal lpFile As String, _
    lpSecurityAttributes As Long) As Long

'RegLoadKey:
' En Windows 95 el nombre del fichero no permite nombres largos
'
Private Declare Function RegLoadKeyA Lib "advapi32.dll" _
    (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal lpFile As String) As Long

'
' Tener en cuenta el usuario actual                                 (22/Jun/00)
' Gracias a Miquel Pop
'
' Funciones y vars para el trabajar con el usuario actual
Private sUser As String
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long


Public Function CreateKey(ByVal sKey As String) As eHKEYError
    ' Crear una clave sin datos adicionales                         (04/Ago/99)
    '
    ' Parámetros:
    '   sKey        Clave a la que se asignará el valor
    ' Devuelve:
    '   El valor de error devuelto por el API
    '
    Dim lRet As eHKEYError
    Dim hKey2 As Long
    Dim hKey As Long
    
    ' Convertimos la clave indicada en un valor correcto,
    ' para el caso que se indique la clave raiz en sKey
    hKey = ParseKey(sKey, hKey)
    
    ' Abrir la clave indicada
    lRet = RegOpenKeyEx(hKey, sKey, 0&, KEY_WRITE, hKey2)
    
    ' Si da error, es que no existe esa clave
    If lRet <> ERROR_SUCCESS Then
        ' Crear la clave
        lRet = RegCreateKey(hKey, sKey, hKey2)
    End If
    Call RegCloseKey(hKey2)
    
    CreateKey = lRet
End Function

Public Function ExistKey(ByVal sKey As String) As Boolean
    ' Comprobar si existe la clave indicada                         (04/Ago/99)
    ' Devolverá TRUE si la clave existe
    Dim ret As eHKEYError
    Dim hKey2 As Long
    Dim hKey As eHKEY
    
    hKey = HKEY_LOCAL_MACHINE
    
    hKey = ParseKey(sKey, hKey)
    
    ' Abrir la clave indicada
    ret = RegOpenKeyEx(hKey, sKey, 0&, KEY_READ, hKey2)
    
    ' Si todo va bien (se ha podido abrir la clave)
    If ret = ERROR_SUCCESS Then
        ExistKey = True
        ' Cerrar la clave abierta
        Call RegCloseKey(hKey2)
    Else
        ExistKey = False
    End If
End Function

Public Function ClassCLSID(ByVal sClass As String) As String
    ' Devuelve el Clsid de la clase indicada                        (05/Jul/99)
    ' El formato del parámetro debe ser Servidor.Clase
    ' Si no se ha encontrado la clase, devuelve una cadena vacía
    '
    Dim sClave As String
    Dim sClsid As String
    Const sRootKey As String = "HKEY_LOCAL_MACHINE\Software\Classes\"
    
    ' Obtener el Clsid
    sClave = sRootKey & sClass & "\clsid"
    sClsid = GetRegString(sClave)
    
    ClassCLSID = sClsid
End Function

Public Function ClassTypeLib(ByVal sClass As String) As String
    ' Devuelve el TypeLib de la clase indicada                      (05/Jul/99)
    ' El formato del parámetro debe ser Servidor.Clase
    ' Si no se ha encontrado la clase, devuelve una cadena vacía
    '
    Dim sClave As String
    Dim sClsid As String
    Const sRootKey As String = "HKEY_LOCAL_MACHINE\Software\Classes\"
    Dim sTypeLib As String
    
    ' Obtener el Clsid
    sClave = sRootKey & sClass & "\clsid"
    sClsid = GetRegString(sClave)
    
    If Len(sClsid) Then
        ' Obtener el TypeLib
        sClave = sRootKey & "CLSID\" & sClsid & "\TypeLib"
        sTypeLib = GetRegString(sClave)
    End If
    ClassTypeLib = sTypeLib
End Function

'------------------------------------------------------------------------------
' Este código está 'copiado' de un ejemplo de David Janson
' Sólo es necesario para Windows NT, ya que win95 permite borrar todas
' las subclaves
'
' También hay que usarla en windows 98
'
'
' this gets a bit tricky since you can't delete a key that has subkeys.
' We have to do this recursively.  This code ignores errors (such as security
' problems) when they occur.
'------------------------------------------------------------------------------
Private Function DeleteKeyNT(hParentKey As Long, szKey As String) As Long
    Dim hKey As Long
    Dim lRet As eHKEYError
    Dim cSubKeys As Long
    Dim cbMaxSubKeyLen As Long
    Dim cbSubKeyLen As Long
    Dim dwIndex As Long
    Dim ft As FILETIME
    
    Dim szTempSubKey As String
    Dim szSubKey As String
    
    ' open the key to look for subkeys
    lRet = RegOpenKeyEx(hParentKey, szKey, 0, KEY_ALL_ACCESS, hKey)
    If Not lRet = ERROR_SUCCESS Then
        ' ERROR_ACCESS_DENIED (5)
        DeleteKeyNT = lRet
        Exit Function
    End If
    'lRet = RegQueryInfoKey(hKey, ByVal 0&, ByVal 0&, 0, _
                           cSubKeys, cbMaxSubKeyLen, _
                           ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&, ft)
    
    lRet = RegQueryInfoKey(hKey, vbNullString, 0&, 0, _
                           cSubKeys, cbMaxSubKeyLen, _
                           0&, 0&, 0&, 0&, 0&, ft)
    If Not lRet = ERROR_SUCCESS Then
        ' ERROR_INVALID_PARAMETERS (87)
        DeleteKeyNT = lRet
        Call RegCloseKey(hKey)
        Exit Function
    End If
    
    ' if there are subkeys, then recursively delete them
    If cSubKeys > 0 Then
        dwIndex = cSubKeys - 1                       ' start at the end
        cbMaxSubKeyLen = cbMaxSubKeyLen + 1          ' +1 for the null terminator
        szTempSubKey = String(cbMaxSubKeyLen, "*")   ' buffer to get name back in
        Do
            '$(22/Ago/99)
            ' Multiplico por dos
            cbSubKeyLen = cbMaxSubKeyLen * 2
            
            'lRet = RegEnumKeyEx(hKey, dwIndex, szTempSubKey, cbSubKeyLen, 0, ByVal 0&, 0, ft)
            lRet = RegEnumKeyEx(hKey, dwIndex, szTempSubKey, cbSubKeyLen, 0&, vbNullString, 0&, ft)
            If lRet = ERROR_SUCCESS Then
                szSubKey = Left(szTempSubKey, cbSubKeyLen)
                Call DeleteKeyNT(hKey, szSubKey)
            End If
            dwIndex = dwIndex - 1                     ' enumerate backwards
        Loop While dwIndex >= 0
    End If
    
    ' done enumerating subkeys.  Close this key and delete it
    Call RegCloseKey(hKey)
    
    lRet = RegDeleteKey(hParentKey, szKey)
    'If Not lRet = ERROR_SUCCESS Then
    '   Exit Sub
    'End If
    DeleteKeyNT = lRet
End Function

Public Function GetRegDWord(ByVal sKey As String, Optional ByVal sValue As String = "", Optional ByVal hKey As eHKEY = HKEY_CURRENT_USER, Optional ByVal bAsString As Boolean = False) As Variant
Attribute GetRegDWord.VB_Description = "Obtener un valor DWORD de una entrada del registro"
Attribute GetRegDWord.VB_HelpID = 21090
    ' Obtener un valor DWORD de una entrada del registro
    '
    ' Parámetros de entrada:
    '   sKey        SubClave del registro
    '   sValue      Nombre de la entrada que queremos obtener
    '   hKey        Clave principal del registro
    '   bAsString   Mostrar en formato al estilo del RegEdit
    ' Devuelve:
    '   el contenido de esa clave o una valor cero
    '
    Dim ret As Long
    Dim hKey2 As Long
    Dim rDT As eHKEYDataType
    Dim lSize As Long
    Dim lDWord As Long
    
    hKey = ParseKey(sKey, hKey)
    
    ' Abrir la clave indicada
    ret = RegOpenKeyEx(hKey, sKey, 0&, KEY_READ, hKey2)
    
    ' Si todo va bien (se ha podido abrir la clave)
    If ret = ERROR_SUCCESS Then
        ' Leer esa entrada y obtener el tipo de dato, longitud, etc.
        ret = RegQueryValueEx(hKey2, sValue, 0&, rDT, 0&, lSize)
        ' Si es un valor DWORD
        If rDT = REG_DWORD Then
            ' Leer los datos DWORD
            ret = RegQueryValueEx(hKey2, sValue, 0&, rDT, lDWord, lSize)
        End If
        ' Cerrar la clave abierta
        RegCloseKey hKey2
    End If
    ' Devolver el valor leído
    If bAsString Then
        ' Al estilo de como se muestra con RegEdit
        'GetRegDWord = "0x" & Format$(Hex$(lDWord), "00000000") & " (" & lDWord & ")"
        GetRegDWord = "0x" & Right$("00000000" & Hex$(lDWord), 8) & " (" & lDWord & ")"
    Else
        GetRegDWord = lDWord
    End If
End Function

Public Function GetReg(ByVal sKey As String, Optional ByVal sValue As String = "", Optional ByVal hKey As eHKEY = HKEY_CURRENT_USER, Optional ByVal bAsString As Boolean = False) As Variant
    '--------------------------------------------------------------------------
    ' Obtener un valor de una entrada del registro
    '
    ' Parámetros de entrada:
    '   sKey        SubClave del registro
    '               Se puede especificar el nombre de la clave raiz
    '               que se convertirá al valor adecuado
    '   sValue      Nombre de la entrada que queremos obtener
    '   hKey        Clave principal del registro.
    '               Si en sKey se incluye, no es necesario especificarla
    '               Nota: este valor se obvia si se indica la raiz en sKey.
    '   bAsString   Mostrarlo como una cadena, al estilo de RegEdit
    ' Devuelve:
    '   el contenido de esa clave o un valor vacío
    '
    ' Revisado para usarlo con Windows NT (Win2000 Pro Beta 3)      (12/Jun/99)
    '--------------------------------------------------------------------------
    Dim lRet As Long
    Dim hKey2 As Long
    Dim rDT As eHKEYDataType
    Dim retDT As eHKEYDataType
    Dim lSize As Long
    Dim sData As String
    Dim aData() As Byte
    Dim lDWord As Long
    Dim i As Long
    Dim sTmp As String
    
    hKey = ParseKey(sKey, hKey)
    
    ' Valores por defecto
    ReDim aData(0)
    lDWord = 0
    sData = ""
    
    ' Abrir la clave indicada
    'lRet = RegOpenKeyEx(hKey, sKey, 0&, KEY_QUERY_VALUE, hKey2)
    lRet = RegOpenKeyEx(hKey, sKey, 0&, KEY_READ, hKey2)
    
    ' Si todo va bien (se ha podido abrir la clave)
    If lRet = ERROR_SUCCESS Then
        ' Leer esa entrada y obtener el tipo de dato, longitud, etc.
        lRet = RegQueryValueEx(hKey2, sValue, 0&, retDT, 0&, lSize)
        Select Case retDT
        Case REG_DWORD
            lRet = RegQueryValueEx(hKey2, sValue, 0&, rDT, lDWord, lSize)
        Case REG_EXPAND_SZ, REG_SZ, REG_MULTI_SZ
            If lSize Then
                sData = String$(lSize - 1, Chr$(0))
                ' Leer la cadena
                '(el ByVal es porque está declarada como Any)---v
                lRet = RegQueryValueEx(hKey2, sValue, 0&, rDT, ByVal sData, lSize)
            End If
        Case Else ' Tratarlos como REG_BINARY
            If lSize Then
                ReDim aData(lSize)
                'Leer los datos binarios
                lRet = RegQueryValueEx(hKey2, sValue, 0&, rDT, aData(0), lSize)
            End If
        End Select
        ' Cerrar la clave abierta
        RegCloseKey hKey2
    End If
    ' Devolver el valor leído
    Select Case retDT
    Case REG_DWORD
        If bAsString Then
            ' Al estilo de como se muestra con RegEdit
            'GetReg = "0x" & Format$(Hex$(lDWord), "00000000") & " (" & lDWord & ")"
            GetReg = "0x" & Right$("00000000" & Hex$(lDWord), 8) & " (" & lDWord & ")"
        Else
            GetReg = lDWord
        End If
    Case REG_EXPAND_SZ, REG_SZ
        GetReg = sData
    Case REG_MULTI_SZ
        ' Múltiples cadenas, separadas por Chr$(0)                  (12/Jun/99)
        ' La cadena termina en el último Chr$(0)
'        For i = Len(sData) To 1 Step -1
'            If Mid$(sData, i, 1) = Chr$(0) Then
'                sData = Left$(sData, i - 1)
'                Exit For
'            End If
'        Next
'        ' Sustituir los Chr$(0) por espacios
'        For i = 1 To Len(sData)
'            If Mid$(sData, i, 1) = Chr$(0) Then
'                Mid$(sData, i, 1) = " "
'            End If
'        Next
        '//////////////////////////////////////////////////////////////////////
        '$TODO: (22/Nov/00)
        ' Separar cada cadena con un punto y coma
        '//////////////////////////////////////////////////////////////////////
        GetReg = RTrimZero(sData, True)
    '--------------------------------------------------------------------------
    ' No poner Case Else, ya que al usar ahora KEY_READ,            (09/Feb/01)
    ' si la clave no existe, devolver un valor vacio
    '--------------------------------------------------------------------------
    'Case Else ' REG_BINARY
    Case REG_BINARY
        If bAsString Then
            ' Al estilo de como se muestra con RegEdit
            For i = 0 To UBound(aData) - 1
                'sTmp = sTmp & Hex$(aData(i)) & " "
                ' Los números formateados a dos cifras              (12/Oct/98)
                'sTmp = sTmp & Format$(Hex$(aData(i)), "00") & " "
                sTmp = sTmp & Right$("00" & Hex$(aData(i)), 2) & " "
            Next
            GetReg = sTmp
        Else
            GetReg = aData
        End If
    End Select
End Function

Public Function GetRegType(ByVal sKey As String, _
                           ByVal sValue As String, _
                           Optional ByVal hKey As eHKEY = HKEY_CURRENT_USER) As eHKEYDataType
    '--------------------------------------------------------------------------
    ' Devuelve el tipo de datos de una entrada del registro         (28/Dic/01)
    '
    ' Parámetros de entrada:
    '   sKey        SubClave del registro
    '               Se puede especificar el nombre de la clave raiz
    '               que se convertirá al valor adecuado
    '   sValue      Nombre de la entrada que queremos obtener
    '   hKey        Clave principal del registro.
    '               Si en sKey se incluye, no es necesario especificarla
    '               Nota: este valor se obvia si se indica la raiz en sKey.
    ' Devuelve:
    '   el tipo de datos del contenido de esa clave
    '
    ' Revisado para usarlo con Windows NT (Win2000 Pro Beta 3)      (12/Jun/99)
    '--------------------------------------------------------------------------
    Dim lRet As Long
    Dim hKey2 As Long
    Dim retDT As eHKEYDataType
    Dim lSize As Long
    '
    hKey = ParseKey(sKey, hKey)
    '
    retDT = REG_NONE
    '
    ' Abrir la clave indicada
    lRet = RegOpenKeyEx(hKey, sKey, 0&, KEY_READ, hKey2)
    '
    ' Si todo va bien (se ha podido abrir la clave)
    If lRet = ERROR_SUCCESS Then
        ' Leer esa entrada y obtener el tipo de dato, longitud, etc.
        lRet = RegQueryValueEx(hKey2, sValue, 0&, retDT, 0&, lSize)
        ' Cerrar la clave abierta
        RegCloseKey hKey2
        '
    End If
    GetRegType = retDT
End Function

Public Function GetRegBinary(ByVal sKey As String, Optional ByVal sValue As String = "", Optional ByVal hKey As eHKEY = HKEY_CURRENT_USER, Optional ByVal bAsString As Boolean = False) As Variant
Attribute GetRegBinary.VB_Description = "Obtener un valor BINARY de una entrada del registro"
Attribute GetRegBinary.VB_HelpID = 21090
    ' Obtener un valor binario de una entrada del registro
    '
    ' Parámetros de entrada:
    '   sKey        SubClave del registro
    '   sValue      Nombre de la entrada que queremos obtener
    '   hKey        Clave principal del registro
    '   bAsString   Mostrarlo como una cadena, al estilo de RegEdit
    ' Devuelve:
    '   el contenido de esa clave o una valor cero
    '
    Dim ret As Long
    Dim hKey2 As Long
    Dim rDT As eHKEYDataType
    Dim lSize As Long
    Dim aData() As Byte
    Dim i As Long
    Dim sTmp As String
    
    hKey = ParseKey(sKey, hKey)
    
    ReDim aData(0)
    
    ' Abrir la clave indicada
    'ret = RegOpenKeyEx(hKey, sKey, 0&, KEY_QUERY_VALUE, hKey2)
    ret = RegOpenKeyEx(hKey, sKey, 0&, KEY_READ, hKey2)
    
    ' Si todo va bien (se ha podido abrir la clave)
    If ret = ERROR_SUCCESS Then
        ' Leer esa entrada y obtener el tipo de dato, longitud, etc.
        ret = RegQueryValueEx(hKey2, sValue, 0&, rDT, 0&, lSize)
        ' Si es un valor binario
        If rDT = REG_BINARY Then
            If lSize Then
                ReDim aData(lSize)
                ' Leer los datos binarios
                ret = RegQueryValueEx(hKey2, sValue, 0&, rDT, aData(0), lSize)
            End If
        End If
        ' Cerrar la clave abierta
        RegCloseKey hKey2
    End If
    ' Devolver el valor leído
    If bAsString Then
        ' Al estilo de como se muestra con RegEdit
        For i = 0 To UBound(aData) - 1
            sTmp = sTmp & Hex$(aData(i)) & " "
        Next
        GetRegBinary = sTmp
    Else
        GetRegBinary = aData
    End If
End Function

Public Function GetRegString(ByVal sKey As String, Optional ByVal sValue As String = "", Optional ByVal hKey As eHKEY = HKEY_CURRENT_USER) As String
Attribute GetRegString.VB_Description = "Obtener un valor STRING de una entrada del registro"
Attribute GetRegString.VB_HelpID = 21090
    ' Obtener un valor cadena de una entrada del registro
    '
    ' Parámetros de entrada:
    '   sKey    Clave del registro
    '   sValue  Nombre de la entrada que queremos obtener
    '   hKey    Clave principal del registro
    ' Devuelve:
    '   el contenido de esa clave o una cadena vacía
    '
    Dim ret As Long
    Dim hKey2 As Long
    Dim rDT As eHKEYDataType
    Dim sData As String
    Dim lSize As Long
    
    hKey = ParseKey(sKey, hKey)
    
    ' Abrir la clave indicada
    'ret = RegOpenKeyEx(hKey, sKey, 0&, KEY_QUERY_VALUE, hKey2)
    ret = RegOpenKeyEx(hKey, sKey, 0&, KEY_READ, hKey2)
    
    ' Si todo va bien (se ha podido abrir la clave)
    If ret = ERROR_SUCCESS Then
        ' Leer esa entrada y obtener el tipo de dato, longitud, etc.
        ret = RegQueryValueEx(hKey2, sValue, 0&, rDT, 0&, lSize)
        ' Si es una cadena o REG_EXPAND_SZ                          (28/Dic/01)
        Select Case rDT
        Case REG_SZ, REG_EXPAND_SZ
        'If rDT = REG_SZ Then
            If lSize Then
                sData = String$(lSize - 1, Chr$(0))
                ' Leer la cadena
                ' (el ByVal es porque está declarada como Any)---v
                ret = RegQueryValueEx(hKey2, sValue, 0&, rDT, ByVal sData, lSize)
            End If
        End Select
        ' Cerrar la clave abierta
        RegCloseKey hKey2
    End If
    ' Devolver el valor leído
    GetRegString = sData
End Function

' Busca una entrada en el registro
Public Function QueryRegBase(ByVal sValue As String, _
                            Optional ByVal hKey As eHKEY = HKEY_CLASSES_ROOT _
                            ) As String
    ' Devuelve el valor de la entrada del registro
    ' Esta función se usará para los valores por defecto
    '
    Dim sBuf As String
    Dim buflen As Long
    
    ' Nos aseguramos que hKey tenga el valor correcto
    Select Case hKey
    'Case HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, HKEY_USERS
    Case HKEY_FIRST To HKEY_LAST
        ' nada que hacer, todo correcto
    Case Else
        ' Asignamos el valor por defecto
        hKey = HKEY_CLASSES_ROOT
    End Select
    
    'On Local Error Resume Next
    sBuf = String$(300, Chr$(0))
    buflen = Len(sBuf)
    ' Buscar la entrada especificada y devolver el valor asignado
    If RegQueryValue(hKey, sValue, sBuf, buflen) = ERROR_SUCCESS Then
        If buflen > 1 Then
            ' El formato devuelto es ASCIIZ, así que quitar el último caracter
            QueryRegBase = Left$(sBuf, buflen - 1)
        Else
            QueryRegBase = ""
        End If
    Else
        QueryRegBase = ""
    End If
    'On Local Error GoTo 0
End Function

Private Function ParseKey(sKey As String, _
                          Optional ByVal hKey As eHKEY = HKEY_CURRENT_USER _
                          ) As eHKEY
    '--------------------------------------------------------------------------
    ' Esta función se usa internamente (privada) para convertir una cadena
    ' en la correspondiente clave raiz.
    ' El segundo parámetro es para poder usarlo en caso que se pase como
    ' parámetro, pero normalmente será totalmente opcional.
    '
    ' En sKey se devolverá el valor de la clave una vez quitada la clave
    ' principal.
    '
    '--------------------------------------------------------------------------
    ' NOTA del 14/Oct/98
    '       En sKey se debe especificar el nombre de la clave raiz.
    '       La utilidad de esta función es que devuelve el valor de esa
    '       clave raiz y se usará en caso de que no sepamos que clave es.
    '       Si ya sabes el valor de la clave raiz, no es necesario que
    '       uses esta función.
    '----------------------------------------------------------------------
    Dim i As Long
    Dim sRootKey As String
    '
    ' Si tiene el separador del final, quitárselo                   (23/Nov/00)
    sKey = Trim$(sKey)
    If Right$(sKey, 1) = "\" Then
        sKey = Left$(sKey, Len(sKey) - 1)
    End If
    ' Comprobar si se indica la clave principal en sKey
    i = InStr(sKey, "HKEY_")
    If i Then
        i = InStr(sKey, "\")
        If i Then
            sRootKey = Left$(sKey, i - 1)
            sKey = Mid$(sKey, i + 1)
        Else
            sRootKey = sKey
            sKey = ""
        End If
    ' Por si se usan abreviaturas de las claves
    ElseIf Left$(sKey, 5) = "HKCR\" Then
        sRootKey = "HKEY_CLASSES_ROOT"
        sKey = Mid$(sKey, 6)
    ElseIf Left$(sKey, 5) = "HKCU\" Then
        sRootKey = "HKEY_CURRENT_USER"
        sKey = Mid$(sKey, 6)
    ElseIf Left$(sKey, 5) = "HKLM\" Then
        sRootKey = "HKEY_LOCAL_MACHINE"
        sKey = Mid$(sKey, 6)
    ElseIf Left$(sKey, 4) = "HKU\" Then
        sRootKey = "HKEY_USERS"
        sKey = Mid$(sKey, 5)
    ElseIf Left$(sKey, 5) = "HKCC\" Then
        sRootKey = "HKEY_CURRENT_CONFIG"
        sKey = Mid$(sKey, 6)
    ElseIf Left$(sKey, 5) = "HKDD\" Then
        sRootKey = "HKEY_DYN_DATA"
        sKey = Mid$(sKey, 6)
    ElseIf Left$(sKey, 5) = "HKPD\" Then
        sRootKey = "HKEY_PERFORMANCE_DATA"
        sKey = Mid$(sKey, 6)
    Else
        ' Nos aseguramos que kKey tenga el valor correcto
        Select Case hKey
        'Case HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, HKEY_USERS, HKEY_CURRENT_CONFIG, HKEY_DYN_DATA
        Case HKEY_FIRST To HKEY_LAST
            'nada que hacer, todo correcto
        Case Else
            ' Asignamos el valor por defecto
            hKey = HKEY_CLASSES_ROOT
        End Select
    End If
    ' Si se ha indicado el nombre de la clave raiz
    If Len(sRootKey) Then
        Select Case sRootKey
        Case "HKEY_CLASSES_ROOT"
            hKey = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_USER"
            hKey = HKEY_CURRENT_USER
        Case "HKEY_LOCAL_MACHINE"
            hKey = HKEY_LOCAL_MACHINE
        Case "HKEY_USERS"
            hKey = HKEY_USERS
        Case "HKEY_CURRENT_CONFIG"
            hKey = HKEY_CURRENT_CONFIG
        Case "HKEY_DYN_DATA"
            hKey = HKEY_DYN_DATA
        Case "HKEY_PERFORMANCE_DATA"
            hKey = HKEY_PERFORMANCE_DATA
        Case Else
            hKey = HKEY_CLASSES_ROOT
        End Select
    End If
    
    ParseKey = hKey
End Function

Public Function OpenKeyEx(ByVal hKey As Long, ByVal lpSubKey As String, _
                        ByVal ulOptions As Long, _
                        ByVal samDesired As eREGSAM, phkResult As Long) As Long
Attribute OpenKeyEx.VB_Description = "Abre una clave del registro, en phkResult devuelve el handle de la clave abierta y se usará para los siguientes accesos."
Attribute OpenKeyEx.VB_HelpID = 21090
    ' Abre una clave del registro, en phkResult devuelve el handle de
    ' la clave abierta y se usará para los siguientes accesos.
    '
    ' ulOptions es un valor reservado que debe ser 0&
    '
    ' Esta función simplemente llama a la original del API
    '
    OpenKeyEx = RegOpenKeyEx(hKey, lpSubKey, 0&, samDesired, phkResult)
End Function

Public Function OpenKeyQuery(ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As eREGSAM, phkResult As Long) As Long
Attribute OpenKeyQuery.VB_Description = "Llama a la función del API RegOpenKeyEx"
Attribute OpenKeyQuery.VB_HelpID = 21090
    ' Los parámetros: ulOptions (un valor reservado que debe ser 0&)
    '               y samDesired, no se tienen en cuenta
    ' pero se dejan por compatibilidad de parámetros de RegOpenKeyEx
    '
    ' Para usar otros valores de accesos, usar la función OpenKeyEx
    '
    ' Esta función simplemente llama a la original del API
    ' Con las "peculiaridades" indicadas
    '
    OpenKeyQuery = RegOpenKeyEx(hKey, lpSubKey, 0&, KEY_QUERY_VALUE, phkResult)
End Function

Public Function EnumValueString(ByVal hKey As Long, ByVal dwIndex As Long, _
    lpValueName As String, lpcbValueName As Long, _
    lpReserved As Long, lpType As Long, lpData As String, _
    lpcbData As Long) As Long
Attribute EnumValueString.VB_Description = "Llama a la función del API RegEnumValue"
Attribute EnumValueString.VB_HelpID = 21090
    '
    ' Esta función simplemente llama a la original del API
    ' Sólo para tipos String
    '
    EnumValueString = RegEnumValue(hKey, dwIndex, _
                            lpValueName, lpcbValueName, _
                            lpReserved, lpType, ByVal lpData, _
                            lpcbData)

End Function

Public Function EnumValue(ByVal hKey As Long, ByVal dwIndex As Long, _
    lpValueName As String, lpcbValueName As Long, _
    lpReserved As Long, lpType As Long, lpData As Byte, _
    lpcbData As Long) As Long
Attribute EnumValue.VB_Description = "Llama a la función del API RegEnumValue"
Attribute EnumValue.VB_HelpID = 21090
    '
    ' Esta función simplemente llama a la original del API
    ' Usarla para tipos diferentes de String
    '
    EnumValue = RegEnumValue(hKey, dwIndex, _
                            lpValueName, lpcbValueName, _
                            lpReserved, lpType, lpData, _
                            lpcbData)

End Function

Public Function CloseKey(ByVal hKey As Long) As Long
Attribute CloseKey.VB_Description = "Cierra la clave abierta anteriormente"
Attribute CloseKey.VB_HelpID = 21090
    ' Cierra la clave abierta usando el handle hKey
    '
    ' Esta función simplemente llama a la original del API
    '
    CloseKey = RegCloseKey(hKey)
End Function

Public Function QueryInfoKey(ByVal hKey As Long, lpcbMaxValueNameLen As Long) As Long
Attribute QueryInfoKey.VB_Description = "Llama a la función del API RegQueryInfoKey"
Attribute QueryInfoKey.VB_HelpID = 21090
    '
    ' Esta función simplemente llama a la original del API
    '
    Dim lpftLastWriteTime As FILETIME
    
    QueryInfoKey = RegQueryInfoKey(hKey, 0&, 0&, 0&, 0&, 0&, 0&, 0&, _
                    lpcbMaxValueNameLen, 0&, 0&, lpftLastWriteTime)
End Function

Public Function EnumKeyEx(ByVal hKey As Long, ByVal dwIndex As Long, lpName As String, lpcbName As Long) As Long
Attribute EnumKeyEx.VB_Description = "Llama a la función del API RegEnumKeyEx"
Attribute EnumKeyEx.VB_HelpID = 21090
    '
    ' Esta función simplemente llama a la original del API
    '
    Dim lpftLastWriteTime As FILETIME
    
    EnumKeyEx = RegEnumKeyEx(hKey, dwIndex, lpName, lpcbName, _
                             0&, 0&, 0&, lpftLastWriteTime)

End Function

Public Function ShellFolders(Optional bSoloClaves As Boolean = False, Optional Usuario As Boolean = True) As Variant
Attribute ShellFolders.VB_Description = "Devolverá las claves de la clave Shell Folders"
Attribute ShellFolders.VB_HelpID = 21090
    ' Devolverá las claves de la clave Shell Folders
    '
    ' El parámetro Usuario indica si se tendrá en cuenta el usuario actual
    '
    Dim sKey As String
    Dim buf As String
    Dim i As Long
    Dim sValue As String
    Dim iCount As Long
    '
    Dim colKeys() As String
    Dim colShellFoldersKey As Collection
    '
    ' Borrar el contenido de la colección
    Set colShellFolders = Nothing
    ' Esta colección tendrá los paths, el índice será la clave
    Set colShellFolders = New Collection
    ' En esta colección se guardarán las claves
    ' (sólo se usa por si se indica bSoloClaves=True)
    Set colShellFoldersKey = New Collection
    
    '==============================================================
    '
    '=== NOTA CACHONDA === por lo incomprensible...
    ' Es curioso, pero si utilizo estas intrucciones aquí
    ' el bucle For iCount=0 to 1 no acaba nunca
    '
    '==============================================================
    '
    'Para el directorio de windows
    'buf = "WindowsDir"
    'colShellFoldersKey.Add buf, buf
    'colShellFolders.Add "Windows", buf
    '
    'Para el directorio de System
    'buf = "SystemDir"
    'colShellFoldersKey.Add buf, buf
    'colShellFolders.Add "System", buf
    '
    '==============================================================
    
    For iCount = 0 To 1
        ' Enumerar el contenido de Shell Folders
        If iCount = 0 Then
            'sKey = "HKEY_USERS\.Default\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
            '
            ' Tener en cuenta el usuario actual                     (22/Jun/00)
            ' Gracias a Miquel Pop
            '
            If Usuario And sUser <> "" Then
                sKey = "HKEY_USERS\" & sUser & "\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
            Else
                sKey = "HKEY_USERS\.Default\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
            End If
        Else
            sKey = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion"
        End If
        
        ' Usar la función EnumValues
        If EnumValues(colKeys(), sKey) Then
            For i = 1 To UBound(colKeys) Step 2
                ' colKeys(i)        será el nombre de la clave
                ' colKeys(i + 1)    será el valor o dato almacenado
                If iCount = 0 Then
                    colShellFoldersKey.Add colKeys(i), colKeys(i)
                    colShellFolders.Add colKeys(i + 1), colKeys(i)
                Else
                    If InStr(colKeys(i + 1), ":\") Then
                        colShellFoldersKey.Add colKeys(i), colKeys(i)
                        colShellFolders.Add colKeys(i + 1), colKeys(i)
                    End If
                End If
            Next
        End If
    Next
    ' Obtener el directorio de windows
    buf = String$(300, Chr$(0))
    i = GetWindowsDirectory(buf, Len(buf))
    sValue = Left$(buf, i)
    buf = "WindowsDir"
    colShellFoldersKey.Add buf, buf
    colShellFolders.Add sValue, buf
    
    ' Obtener el directorio de System
    buf = String$(300, Chr$(0))
    i = GetSystemDirectory(buf, Len(buf))
    sValue = Left$(buf, i)
    buf = "SystemDir"
    colShellFoldersKey.Add buf, buf
    colShellFolders.Add sValue, buf
    
    If bSoloClaves Then
        Set ShellFolders = colShellFoldersKey
    Else
        Set ShellFolders = colShellFolders
    End If
        
    Set colShellFoldersKey = Nothing
End Function

Private Sub Class_Initialize()
    Set colShellFolders = New Collection
    'Set colShellFoldersKey = New Collection
    '
    ' Tener en cuenta el usuario actual                             (22/Jun/00)
    ' Gracias a Miquel Pop
    '
    Dim n As Long
    '
    n = 255
    sUser = Space(n)
    If GetUserName(sUser, n) <> 0 Then sUser = Left$(sUser, n - 1)
End Sub

Private Sub Class_Terminate()
    Set colShellFolders = Nothing
    'Set colShellFoldersKey = Nothing
End Sub

Public Function GetFolder(ByVal vIndex As Variant) As String
Attribute GetFolder.VB_Description = "Devuelve el directorio de la clave indicada"
Attribute GetFolder.VB_HelpID = 21090
    ' Devuelve el directorio de la clave indicada
    
    ' Si no está preparada la colección, prepararla
    If colShellFolders.Count = 0 Then
        Call ShellFolders
    End If
    
    On Local Error Resume Next
    ' Devolver el directorio de la clave indicada
    GetFolder = colShellFolders(vIndex)
    
    ' Si da error es que no existe la clave que se indica...
    If Err Then
        GetFolder = ""
    End If
    
    Err = 0
    On Local Error GoTo 0
End Function

Public Function RegSetValue2(ByVal hKey As Long, ByVal lpSubKey As String, _
                            ByVal dwType As eHKEYDataType, lpData As String, _
                            ByVal cbData As Long) As Long
Attribute RegSetValue2.VB_Description = "Llama a la función del API RegSetValueEx"
Attribute RegSetValue2.VB_HelpID = 21090
    '--------------------------------------------------------------------------
    ' Lo que dice la ayuda de Windows:
    '   The RegSetValue function sets the data for the default or unnamed
    '   value of a specified registry key. The data must be a text string.
    '
    ' Función para compatibilidad con versiones anteriores
    '--------------------------------------------------------------------------
    cbData = Len(lpData)
    ' Hay que usar ByVal porque está definida "As Any" -------v
    RegSetValue2 = RegSetValueEx(hKey, lpSubKey, 0&, REG_SZ, ByVal lpData, cbData)
End Function

Public Sub AsociarExt(ByVal sExt As String, _
                      Optional ByVal sExe As String = "", _
                      Optional ByVal sCommand As String = "open", _
                      Optional ByVal bDefault As Boolean = True, _
                      Optional ByVal sProgId As String = "", _
                      Optional ByVal sDescription As String = "")
    '--------------------------------------------------------------------------
    ' Asociar una extensión con un programa
    ' También sirve para añadir comandos a extensiones existentes
    '
    ' Revisada: 19/May/2001, arreglado bug cuando se especificaba "open"
    '
    ' Parámetros:
    '   sExt            Extensión a asociar
    '   sExe            Path completo del programa (y los parámetros)
    '   sProgId         Nombre de la clave asociada
    '   sDescription    Descripción de la extensión
    '   sCommand        Clave a crear, por defecto es Abrir (open)
    '   bDefault        Si la clave indicada es la que se usará por defecto
    '
    '--------------------------------------------------------------------------
    ' Para probar:
    ' tQR.AsociarExt ".cIt", "C:\Vb5_L\Cut-It\Cut-It.exe", "open", False, "gsCutIt", "Cut-It (trocear y unir archivos)"
    ' tQR.AsociarExt ".cIt", "C:\Windows\Notepad.exe", "&Editar", True, "gsCutIt", "Cut-It (trocear y unir archivos)"
    '
    ' Si se quieren especificar parámetros, hay que hacerlo en el nombre del ejecutable
    ' tQr.AsociarExt ".vbw", sAppName & " /S ", "AsignarAuto", False, sProgId, sDescription
    '
    ' Sólo se quitará el valor por defecto si se asigna a otra clave.
    ' tQR.AsociarExt ".cIt", "", "open", True, "gsCutIt", "Cut-It (trocear y unir archivos)"
    ' tQR.AsociarExt ".cIt", "", "&Editar", True, "gsCutIt", "Cut-It (trocear y unir archivos)"
    '--------------------------------------------------------------------------
    Dim sDef As String
    Dim hKey As Long
    Dim phkResult As Long
    Dim lRet As eHKEYError
    Dim sValue As String
    Dim sKey As String
    Dim sAccess As String
    
    ' Quitar los espacios
    sExt = Trim$(sExt)
    sExe = Trim$(sExe)
    sCommand = Trim$(sCommand)
    sProgId = Trim$(sProgId)
    sDescription = Trim$(sDescription)
    
    ' Si no se especifica el punto
    If InStr(sExt, ".") = 0 Then
        sExt = "." & sExt
    End If
    
    ' Comprobar el tipo de ejecutable, si no se especifica la extensión
    ' se añade .exe
    If Len(sExe) Then
        If InStr(sExe, ".") = 0 Then
            sExe = sExe & ".exe"
        End If
        sExe = sExe & " "
    End If
    
    ' Si no se especifica el ProgId
    If Len(sProgId) = 0 Then
        sProgId = "progID" & sExt
    End If
    
    ' Si no se especifica la descripción
    If Len(sDescription) = 0 Then
        sDescription = "Descripción de " & sProgId
    End If
    
    sAccess = sCommand
    ' Comprobar si tiene el símbolo & y quitarlo del commando
    lRet = InStr(sAccess, "&")
    If lRet Then
        sCommand = Left$(sAccess, lRet - 1) & Mid$(sAccess, lRet + 1)
    End If
    
'    On Local Error GoTo AsociarExtErr
    
    sValue = sProgId
    sProgId = QueryRegBase(sExt)
    If Len(sProgId) = 0 Then
        ' Registrar la extensión
        sKey = sExt
        sProgId = sValue
        lRet = RegSetValue(HKEY_CLASSES_ROOT, sKey, REG_SZ, sValue, Len(sValue))
        '
        sKey = sProgId
        sValue = sDescription
        lRet = RegSetValue(HKEY_CLASSES_ROOT, sKey, REG_SZ, sValue, Len(sValue))
    End If
    sProgId = QueryRegBase(sExt)
    If Len(sProgId) Then
        ' Nombre de la clave para esta extensión
        sDef = "Software\Classes\" & sProgId & "\shell"
        ' usar HKEY_LOCAL_MACHINE, ya que HKEY_CLASSES_ROOT es una copia de:
        ' HKEY_LOCAL_MACHINE\Software\Classes
        
        hKey = HKEY_LOCAL_MACHINE
        ' Crear la clave del registro, si ya existe, simplemente la abre.
        ' Nota: Esta función permite crear varios niveles
        lRet = RegCreateKey(hKey, sDef, phkResult)
        If lRet = ERROR_SUCCESS Then
            ' Si no hay error, la clave está creada y/o abierta
            '
            ' Si no es "open"
            If sCommand <> "open" Then
                sKey = sCommand
                sValue = sAccess
                lRet = RegSetValue(phkResult, sKey, REG_SZ, sValue, Len(sValue))
                '
                If Len(sExe) Then
                    sKey = sCommand & "\command"
                    sValue = sExe & Chr$(34) & "%1" & Chr$(34)
                    lRet = RegSetValue(phkResult, sKey, REG_SZ, sValue, Len(sValue))
                End If
            Else
                ' Abrir (open)
                If Len(sExe) Then
                    ' sKey era "\open\command" y no se registraba   (19/May/01)
                    sKey = "open\command"
                    sValue = sExe & Chr$(34) & "%1" & Chr$(34)
                    ' Si no se especifica sKey, se asigna a la clave abierta
                    lRet = RegSetValue(phkResult, sKey, REG_SZ, sValue, Len(sValue))
                End If
            End If
            If bDefault Then
                ' Poner este prograna por defecto (asignarlo a Shell)
                ' Si no se especifica sKey, se asigna a la clave abierta
                sKey = ""
                sValue = sCommand 'sProgId
                lRet = RegSetValue(phkResult, sKey, REG_SZ, sValue, Len(sValue))
            End If
            '
            ' Cerrar la clave abierta
            lRet = RegCloseKey(phkResult)
        End If
    End If
'    Exit Sub
'AsociarExtErr:
'    Debug.Print "AsociarExt, error # " & Err.Number & " " & Err.Description
'    Err = 0
End Sub

Public Sub AsociarIcono(ByVal sExt As String, _
                        Optional ByVal sExe As String = "", _
                        Optional ByVal Icono As Long = 0&)
    '--------------------------------------------------------------------------
    ' Asociar el icono de un programa a la extensión indicada       (19/May/01)
    '
    ' Parámetros:
    '   sExt            Extensión a asociar
    '   sExe            Path completo del ejecutable o librería con los iconos
    '                   Si no se especifica, quitar la clave del registro
    '   Icono           Número del icono a usar, por defecto cerá cero
    '--------------------------------------------------------------------------
    ' Para probar:
    '--------------------------------------------------------------------------
    Dim sDef As String
    Dim hKey As Long
    Dim phkResult As Long
    Dim lRet As eHKEYError
    Dim sValue As String
    Dim sKey As String
    Dim sProgId As String
    '
    ' Quitar los espacios
    sExt = Trim$(sExt)
    sExe = Trim$(sExe)
    '
    ' Si no se especifica el punto
    If InStr(sExt, ".") = 0 Then
        sExt = "." & sExt
    End If
    '
    ' Comprobar el tipo de ejecutable, si no se especifica la extensión
    ' se añade .exe
    If Len(sExe) Then
        If InStr(sExe, ".") = 0 Then
            sExe = sExe & ".exe"
        End If
    End If
    sProgId = QueryRegBase(sExt)
    ' Si la extensión está registrada...
    If Len(sProgId) Then
        ' Nombre de la clave para esta extensión
        sDef = "Software\Classes\" & sProgId & "\DefaultIcon"
        '
        hKey = HKEY_LOCAL_MACHINE
        '
        If Len(sExe) = 0 Then
            Call DeleteKey(sDef, "", HKEY_LOCAL_MACHINE)
        Else
            ' Crear la clave del registro, si ya existe, simplemente la abre.
            ' Nota: Esta función permite crear varios niveles
            lRet = RegCreateKey(hKey, sDef, phkResult)
            If lRet = ERROR_SUCCESS Then
                ' Si no hay error, la clave está creada y/o abierta
                '
                ' Poner este prograna por defecto (asignarlo a Shell)
                ' Si no se especifica sKey, se asigna a la clave abierta
                sKey = ""
                sValue = sExe & "," & CStr(Icono)
                lRet = RegSetValue(phkResult, sKey, REG_SZ, sValue, Len(sValue))
                '
                ' Cerrar la clave abierta
                lRet = RegCloseKey(phkResult)
            End If
        End If
    End If
End Sub

Public Function DeleteKeyWin95(ByVal hKey As Long, ByVal szKey As String) As Long
Attribute DeleteKeyWin95.VB_Description = "Borra claves del registro, sólo para Windows 95, no funciona en NT ni en Windows 98"
Attribute DeleteKeyWin95.VB_HelpID = 21090
    ' Esta no funciona en Windows NT y parece que tampoco en Win98
    ' Sólo en Windows 95
    
    DeleteKeyWin95 = RegDeleteKey(hKey, szKey)
    
    'Dim lRet As eHKEYError
    'Dim phkResult As Long
    
    'lRet = RegOpenKeyEx(hKey, szKey, 0&, KEY_ALL_ACCESS, phkResult)
    'If lRet = ERROR_SUCCESS Then
    '    lRet = RegDeleteKey(phkResult, szKey)
    '    Call RegCloseKey(phkResult)
    'End If
    
    'DeleteKeyWin95 = lRet
End Function

Public Sub DesasociarExt(ByVal sExt As String)
Attribute DesasociarExt.VB_Description = "Desasociar una extensión, borrar las entradas que haya en el registro"
Attribute DesasociarExt.VB_HelpID = 21090
    ' Para desasociar la extensión indicada
    '
    Dim sProgId As String
    Dim lRet As eHKEYError
    
    ' Si no se especifica el punto
    If InStr(sExt, ".") = 0 Then
        sExt = "." & sExt
    End If
    
    sProgId = QueryRegBase(sExt)
    ' Si la extensión está registrada...
    If Len(sProgId) Then
        ' Esto sólo funciona en Windows 95
        'lRet = DeleteKeyWin95(HKEY_CLASSES_ROOT, sExt)
        'If lRet = ERROR_SUCCESS Then
        '    Call DeleteKeyWin95(HKEY_CLASSES_ROOT, sProgId)
        'End If
        
        ' Esto funciona en Windows 98 y Windows NT,
        ' también en Win95, aunque algo más lento...
        Call DeleteKeyNT(HKEY_CLASSES_ROOT, sExt)
        Call DeleteKeyNT(HKEY_CLASSES_ROOT, sProgId)
    End If
End Sub

Public Function SetReg(ByVal sKey As String, ByVal sName As String, _
                       Optional ByVal vValue As Variant, _
                       Optional ByVal hKey As eHKEY = HKEY_CURRENT_USER, _
                       Optional ByVal RegDataType As eHKEYDataType = REG_SZ, _
                       Optional ByVal bCreateKey As Boolean = True) As eHKEYError
Attribute SetReg.VB_Description = "Asignar un valor de cualquier tipo"
Attribute SetReg.VB_HelpID = 21090
    ' Asignar un valor en el registro
    '
    ' Parámetros:
    '   sKey        Clave a la que se asignará el valor
    '   sName       Nombre de la entrada a asignar el valor
    '   vValue      Valor a asignar, el tipo se debe corresponder con el
    '               tipo indicado en el parámetro RegDataType
    '   hKey        Clave principal del registro.
    '               Si en sKey se incluye, no es necesario especificarla
    '   RegDataType Tipo de dato a asignar
    '   bCreateKey  Si no existe la clave, crearla
    '
    ' Devolverá un valor del tipo: eHKEYError
    '
    
    Dim lRet As Long
    Dim hKey2 As Long
    Dim cbData As Long
    Dim aData() As Byte
    Dim sData As String
    Dim lData As Long
    
    ' Convertimos la clave indicada en un valor correcto,
    ' para el caso que se indique la clave raiz en sKey
    hKey = ParseKey(sKey, hKey)
    
    ' Abrir la clave indicada
    lRet = RegOpenKeyEx(hKey, sKey, 0&, KEY_WRITE, hKey2)
    
    ' Si da error, comprobar si se crea la clave
    If lRet <> ERROR_SUCCESS Then
        If bCreateKey Then
            lRet = RegCreateKey(hKey, sKey, hKey2)
        End If
    End If
    ' Si se produce error, salir
    If lRet <> ERROR_SUCCESS Then
        SetReg = lRet
        Exit Function
    End If
    
    ' Asignar el valor
    '
    Select Case RegDataType
    Case REG_BINARY
        aData = vValue
        cbData = UBound(aData)
        lRet = RegSetValueEx(hKey2, sName, 0&, RegDataType, aData(0), cbData)
    Case REG_DWORD
        cbData = 4
        lData = CLng(vValue)
        lRet = RegSetValueEx(hKey2, sName, 0&, RegDataType, lData, cbData)
    Case REG_SZ, REG_EXPAND_SZ
        sData = CStr(vValue)
        If Len(sData) = 0 Then
            sData = ""
        End If
        cbData = Len(sData) + 1
        ' Hay que usar ByVal porque está declarado como Any---v
        lRet = RegSetValueEx(hKey2, sName, 0&, RegDataType, ByVal sData, cbData)
    Case Else
        ' No implementado...
    End Select
    lRet = RegCloseKey(hKey2)
    
    SetReg = lRet
End Function

Public Function DeleteKey(ByVal sKey As String, _
                          Optional ByVal sValue As String = "", _
                          Optional ByVal hKey As eHKEY = HKEY_CURRENT_USER _
                          ) As eHKEYError
    '--------------------------------------------------------------------------
    ' Borrar la clave especificada del registro
    ' o el valor especificado
    '
    ' Parámetros de entrada:
    '   sKey        SubClave del registro
    '               Se puede especificar el nombre de la clave raiz
    '               que se convertirá al valor adecuado
    '   sValue      Nombre de la entrada que queremos borrar.
    '               Si no se especifica, se borrará la clave.
    '   hKey        Clave principal del registro.
    '               Si en sKey se incluye, no es necesario especificarla
    ' Devuelve:
    '   el código devuelto por la operación realizada
    '--------------------------------------------------------------------------
    Dim lRet As eHKEYError
    Dim hKey2 As Long
    
    ' Nos aseguramos que hKey tenga el valor correcto
    Select Case hKey
    'Case HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, HKEY_USERS
    Case HKEY_FIRST To HKEY_LAST
        ' nada que hacer, todo correcto
    Case Else
        ' Asignamos el valor por defecto
        hKey = HKEY_CLASSES_ROOT
    End Select
    
    hKey = ParseKey(sKey)
    
    ' Si no se especifica sValue, se borra la clave
    If Len(sValue) = 0 Then
        DeleteKey = DeleteKeyNT(hKey, sKey)
        Exit Function
    End If
    ' Borrar el valor indicado
    lRet = RegOpenKeyEx(hKey, sKey, 0&, KEY_WRITE, hKey2)
    If lRet = ERROR_SUCCESS Then
        lRet = RegDeleteValue(hKey2, sValue)
        Call RegCloseKey(hKey2)
    End If
    
    DeleteKey = lRet
End Function
'
Public Function DeleteValue(ByVal sKey As String, _
                            ByVal sValue As String, _
                            Optional ByVal hKey As eHKEY = HKEY_CURRENT_USER _
                            ) As eHKEYError
    '--------------------------------------------------------------------------
    ' Borrar el valor indicado                                      (28/Dic/01)
    ' DeleteKey también lo puede hacer,
    ' pero lo pongo por separado para evitar confusiones
    '
    ' Parámetros de entrada:
    '   sKey        SubClave del registro
    '               Se puede especificar el nombre de la clave raiz
    '               que se convertirá al valor adecuado
    '   sValue      Nombre de la entrada que queremos borrar.
    '               No se admiten cadenas vacías
    '   hKey        Clave principal del registro.
    '               Si en sKey se incluye, no es necesario especificarla
    ' Devuelve:
    '   el código devuelto por la operación realizada
    '--------------------------------------------------------------------------
    Dim lRet As eHKEYError
    Dim hKey2 As Long
    
    ' Nos aseguramos que hKey tenga el valor correcto
    Select Case hKey
    'Case HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, HKEY_USERS
    Case HKEY_FIRST To HKEY_LAST
        ' nada que hacer, todo correcto
    Case Else
        ' Asignamos el valor por defecto
        hKey = HKEY_CLASSES_ROOT
    End Select
    
    hKey = ParseKey(sKey)
    lRet = ERROR_NONE
    '
    If Len(sValue) Then
        ' Borrar el valor indicado
        lRet = RegOpenKeyEx(hKey, sKey, 0&, KEY_WRITE, hKey2)
        If lRet = ERROR_SUCCESS Then
            lRet = RegDeleteValue(hKey2, sValue)
            Call RegCloseKey(hKey2)
        End If
    End If
    '
    DeleteValue = lRet
End Function

Public Function EnumKeys(ByRef colKeys() As String, ByVal sKey As String) As Boolean
Attribute EnumKeys.VB_Description = "Enumera todas las subclaves de la clave indicada"
Attribute EnumKeys.VB_HelpID = 21090
    '--------------------------------------------------------------------------
    ' Enumera todas las subclaves de la clave indicada en sKey      (12/Oct/98)
    '
    ' Parámetros:
    '   colKeys()   Array unidimensional que contendrá las claves halladas
    '               Los valores devueltos estarán comprendidos entre:
    '               1 y UBound(colKeys)
    '   sKey        Clave completa de la que se quiere la información
    '
    ' Devolverá True si todo va bien
    '
    ' Revisado para Array y buen funcionamiento (espero)            (14/Oct/98)
    ' Revisado para funcionar en Windows NT (Win2000 Prof Beta 3)   (12/Jun/99)
    '--------------------------------------------------------------------------
    Dim dwIndex             As Long
    Dim ret                 As Long
    Dim hKey2               As Long
    Dim hKey                As Long
    Dim lpName              As String
    Dim lpftLastWriteTime   As FILETIME
    Dim colItems            As Long
    Dim lSize               As Long
    Dim SubKeysNum          As Long
    Dim MaxSubKeyLen        As Long
    Dim numValues           As Long
    Dim MaxValueNameLen     As Long
    Dim MaxDataLen          As Long
    
    colItems = 0
    ReDim colKeys(0)
    
    ' Si se pasa una cadena en sKey, esta función la convierte
    ' en un valor válido para la clave principal
    hKey = ParseKey(sKey, hKey)
    
    ' Abrir la clave indicada
    '//////////////////////////////////////////////////////////////////////////
    '   Para que en Windows 2000 funcione,                          (12/Jun/99)
    '   he cambiado el tipo de acceso de KEY_ENUMERATE_SUB_KEYS a KEY_READ
    '//////////////////////////////////////////////////////////////////////////
    'ret = RegOpenKeyEx(hKey, sKey, 0&, KEY_ENUMERATE_SUB_KEYS, hKey2)
    ret = RegOpenKeyEx(hKey, sKey, 0&, KEY_READ, hKey2)
    
    EnumKeys = True
    ' Si todo va bien (se ha podido abrir la clave)
    If ret = ERROR_SUCCESS Then
        ' Obtener información de la clave y datos, devolverá:
        ' SubKeysNum        Número de subclaves
        ' MaxSubKeyLen      Tamaño máximo de nombre de clave
        ' numValues         Número de valores en esta clave
        ' MaxValueNameLen   Tamaño máximo del nombre del valor
        ' MaxDataLen        Tamaño máximo de los datos
        '
        ' Probado para Windows 2000 Professional                    (12/Jun/99)
        'ret = RegQueryInfoKey(hKey2, vbNullString, 0&, 0&, _
                            SubKeysNum, MaxSubKeyLen, _
                            0&, numValues, MaxValueNameLen, _
                            MaxDataLen, 0&, lpftLastWriteTime)
        ' Sólo nos interesa el número de subclaves y la longitud máxima
        ret = RegQueryInfoKey(hKey2, vbNullString, 0&, 0&, _
                            SubKeysNum, MaxSubKeyLen, _
                            0&, 0&, 0&, _
                            0&, 0&, lpftLastWriteTime)
        
        ' Se empieza desde cero
        For dwIndex = 0 To SubKeysNum
            ' He puesto el doble,                                   (22/Ago/99)
            ' ya que en Windows 2000 cortaba algunos nombres
            lSize = MaxSubKeyLen * 2
            lpName = String$(lSize + 1, Chr$(0))
            ' Sólo nos interesa los nombres de las subclaves
            ret = RegEnumKeyEx(hKey2, dwIndex, lpName, lSize, _
                            0&, vbNullString, 0&, lpftLastWriteTime)
            If ret = ERROR_MORE_DATA Or ret = ERROR_SUCCESS Then
                ' Redimensionar el array
                colItems = colItems + 1
                ReDim Preserve colKeys(colItems)
                ' lSize tiene el número de caracteres devuelto,
                ' sin incluir el CHR$(0) del final
                colKeys(colItems) = Left$(lpName, lSize)
            End If
        Next
    Else
        EnumKeys = False
    End If
    ' Cerrar la clave abierta
    Call CloseKey(hKey2)
End Function

Private Function RTrimZero(ByVal sString As String, _
                           Optional ByVal PorElFinal As Boolean = False) As String
    ' Devuelve una cadena hasta el primer Chr$(0)                   (12/Oct/98)
    ' Ampliada para poder devolver hasta el último                  (12/Jun/99)
    Dim i As Long
    
    ' Si se quitan los últimos Chr$(0)
    If PorElFinal Then
        ' La cadena termina en el último Chr$(0)
        For i = Len(sString) To 1 Step -1
            If Mid$(sString, i, 1) = Chr$(0) Then
                sString = Left$(sString, i - 1)
                Exit For
            End If
        Next
        ' Sustituir los Chr$(0) por espacios
        For i = 1 To Len(sString)
            If Mid$(sString, i, 1) = Chr$(0) Then
                Mid$(sString, i, 1) = " "
            End If
        Next
    
'        i = Len(sString)
'        Do While Mid$(sString, i, 1) = Chr$(0)
'            i = i - 1
'        Loop
'        ' i tendrá el primer caracter que no es un Chr$(0)
'        If i > 0 Then
'            sString = Left$(sString, i)
'        End If
    Else
        i = InStr(sString, Chr$(0))
        If i Then
            sString = Left$(sString, i - 1)
        End If
    End If
    RTrimZero = sString
End Function

Public Function EnumValues(ByRef colKeys() As String, ByVal sKey As String) As Boolean
    '--------------------------------------------------------------------------
    ' Enumera todos los valores de la clave indicada en sKey        (12/Oct/98)
    '
    ' Parámetros:
    '   colKeys()   Array unidimensional que contendrá las claves halladas
    '               En este array se almacena el nombre del valor y a
    '               continuación el valor en si, por tanto hay que tener
    '               esto en cuenta a la hora de recuperar la información.
    '               Los valores estarán comprendidos entre 1 y UBound(colKeys)
    '               colKeys(i)= nombre (valor), colKeys(i+1)= valor (dato)
    '               Ver el ejemplo de cómo usarla un poco más abajo.
    '   sKey        Clave completa de la que se quiere la información
    '
    ' Devolverá True si todo va bien
    '
    ' Revisado para Array y buen funcionamiento (espero)         (14/Oct/98)
    '--------------------------------------------------------------------------
    ' Para recuperar la información de colKeys(), hacer esto:
    '
    'If .EnumValues(colKeys(), sKey) Then
    '    For i = 1 To UBound(colKeys) Step 2
    '        'colKeys(i)        será el nombre
    '        'colKeys(i + 1)    será el valor o dato almacenado
    '    Next
    'End If
    '----------------------------------------------------------------------
    '
    Dim dwIndex             As Long
    Dim ret                 As Long
    Dim hKey2               As Long
    Dim hKey                As Long
    Dim lpName              As String
    Dim lpftLastWriteTime   As FILETIME
    Dim retDT               As eHKEYDataType
    Dim lSize               As Long
    Dim sData               As String
    Dim aData()             As Byte
    Dim lDWord              As Long
    Dim i                   As Long
    Dim colItems            As Long
    Dim SubKeysNum          As Long
    Dim MaxSubKeyLen        As Long
    Dim numValues           As Long
    Dim MaxValueNameLen     As Long
    Dim MaxDataLen          As Long
    
    ' Si se pasa una cadena en sKey, esta función la convierte
    ' en un valor válido para la clave principal
    hKey = ParseKey(sKey, hKey)
    
    ' Abrir la clave indicada
    ' En este caso da igual el tipo de acceso,                      (12/Jun/99)
    ' pero... más vale prevenir
    'ret = RegOpenKeyEx(hKey, sKey, 0&, KEY_QUERY_VALUE, hKey2)
    ret = RegOpenKeyEx(hKey, sKey, 0&, KEY_READ, hKey2)
    
    ' Valores por defecto
    EnumValues = True
    ReDim aData(0)
    lDWord = 0
    sData = ""
    
    ' Inicializar el array
    ' Añadir también el valor Default                               (23/Nov/00)
    colItems = 2
    ReDim colKeys(colItems)
    colKeys(1) = "(Default)"
    colKeys(2) = QueryRegBase(sKey, hKey)
    '
    ' Si todo va bien (se ha podido abrir la clave)
    If ret = ERROR_SUCCESS Then
        ' Obtener la información de esta clave, devolverá:
        ' SubKeysNum        Número de subclaves
        ' MaxSubKeyLen      Tamaño máximo de nombre de clave
        ' numValues         Número de valores en esta clave
        ' MaxValueNameLen   Tamaño máximo del nombre del valor
        ' MaxDataLen        Tamaño máximo de los datos
        'ret = RegQueryInfoKey(hKey2, 0&, 0&, 0&, SubKeysNum, MaxSubKeyLen, _
                            0&, numValues, MaxValueNameLen, _
                            MaxDataLen, 0&, lpftLastWriteTime)
        
        ' A ver si así funciona...                                  (12/Jun/99)
        ret = RegQueryInfoKey(hKey2, vbNullString, 0&, 0&, _
                            SubKeysNum, MaxSubKeyLen, _
                            0&, numValues, MaxValueNameLen, _
                            MaxDataLen, 0&, lpftLastWriteTime)
        
        ' Este es el error que me da el Windows 2000 Pro            (12/Jun/99)
        If ret = ERROR_INVALID_PARAMETERS Then
            Debug.Print "ERROR_INVALID_PARAMETERS"
            EnumValues = False
            GoTo SalirEnumValues
        End If
        
        'lpName = String$(MaxValueNameLen + 1, 0)
        ' Hacer un bucle para el número de valores posibles
        For dwIndex = 0 To numValues
            lpName = String$(MaxValueNameLen + 1, Chr$(0))
            ' Llamarlo primero para saber el tipo de datos,
            ' el cual estará en retDT
            '//////////////////////////////////////////////////////////////////
            ' De esta forma en Win2000 produce un error de protección
            'ret = RegEnumValue(hKey2, dwIndex, ByVal 0&, ByVal 0&, 0&, retDT, ByVal 0&, ByVal 0&)
            'ret = RegEnumValue(hKey2, dwIndex, vbNullString, ByVal 0&, 0&, retDT, ByVal 0&, ByVal 0&)
            '//////////////////////////////////////////////////////////////////
            ret = RegEnumValue(hKey2, dwIndex, 0&, 0&, 0&, retDT, 0&, 0&)
            'ret = RegEnumValue(hKey2, dwIndex, lpName, Len(lpName), 0&, retDT, ByVal sData, lSize)
            ' la primera vez, cuando dwIndex = cero, devuelve ERROR_SUCCESS,
            ' pero después devuelve ERROR_MORE_DATA mientras haya datos.
            If ret = ERROR_MORE_DATA Or ret = ERROR_SUCCESS Then
                lSize = MaxDataLen
                Select Case retDT
                Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
                    ' Datos de cadena
                    sData = String$(lSize, Chr$(0))
                    ret = RegEnumValue(hKey2, dwIndex, lpName, Len(lpName), 0&, retDT, ByVal sData, lSize)
                    If retDT = REG_MULTI_SZ Then
                        sData = RTrimZero(sData, True)
                    Else
                        sData = RTrimZero(sData)
                    End If
                    lpName = RTrimZero(lpName)
                    ReDim Preserve colKeys(colItems + 2)
                    colKeys(colItems + 1) = lpName
                    colKeys(colItems + 2) = sData
                    colItems = colItems + 2
                Case REG_DWORD
                    ' Datos numéricos (long)
                    ret = RegEnumValue(hKey2, dwIndex, lpName, Len(lpName), 0&, retDT, lDWord, lSize)
                    sData = CStr(lDWord)
                    lpName = RTrimZero(lpName)
                    ReDim Preserve colKeys(colItems + 2)
                    colKeys(colItems + 1) = lpName
                    colKeys(colItems + 2) = sData
                    colItems = colItems + 2
                'Case REG_BINARY
                '    'Datos binarios
                Case Else
                    ' Tratarlo como Binary
                    If lSize Then
                        ReDim aData(lSize)
                        ' Leer los datos binarios
                        ret = RegEnumValue(hKey2, dwIndex, lpName, Len(lpName), 0&, retDT, aData(0), lSize)
                        lpName = RTrimZero(lpName)
                        ' Al estilo de como se muestra con RegEdit
                        sData = ""
                        ' lSize ahora contendrá el número de bytes  (22/Nov/00)
                        ' almacenados en el búfer
                        For i = 0 To lSize - 1 'UBound(aData) - 1
                            sData = sData & Right$("00" & Hex$(aData(i)), 2) & " "
                        Next
                        ReDim Preserve colKeys(colItems + 2)
                        colKeys(colItems + 1) = lpName
                        colKeys(colItems + 2) = sData
                        colItems = colItems + 2
                    End If
                End Select
            End If
        Next
    Else
        EnumValues = False
    End If
SalirEnumValues:
    ' Cerrar la clave abierta
    ret = CloseKey(hKey2)
End Function
'
Public Function EnumValuesByType(ByRef colKeys() As String, ByVal sKey As String, _
                                 Optional ByVal tDT As eHKEYDataType = REG_SZ) As Boolean
    '--------------------------------------------------------------------------
    ' Enumera todos los valores de la clave indicada en sKey,       (28/Dic/01)
    ' pero sólo los que son del tipo de datos especificado
    '
    ' Parámetros:
    '   colKeys()   Array unidimensional que contendrá las claves halladas
    '               En este array se almacena el nombre del valor y a
    '               continuación el valor en si, por tanto hay que tener
    '               esto en cuenta a la hora de recuperar la información.
    '               Los valores estarán comprendidos entre 1 y UBound(colKeys)
    '               colKeys(i)= nombre (valor), colKeys(i+1)= valor (dato)
    '               Ver el ejemplo de cómo usarla un poco más abajo.
    '   sKey        Clave completa de la que se quiere la información
    '   tDT         El tipo de datos a enumerar
    '               Si se especifica REG_SZ, también se devolverán los otros:
    '               REG_EXPAND_SZ y REG_MULTI_SZ
    '
    ' Devolverá True si todo va bien
    '--------------------------------------------------------------------------
    ' Para recuperar la información de colKeys(), hacer esto:
    '
    'If .EnumValues(colKeys(), sKey) Then
    '    For i = 1 To UBound(colKeys) Step 2
    '        'colKeys(i)        será el nombre
    '        'colKeys(i + 1)    será el valor o dato almacenado
    '    Next
    'End If
    '----------------------------------------------------------------------
    '
    Dim dwIndex             As Long
    Dim ret                 As Long
    Dim hKey2               As Long
    Dim hKey                As Long
    Dim lpName              As String
    Dim lpftLastWriteTime   As FILETIME
    Dim retDT               As eHKEYDataType
    Dim lSize               As Long
    Dim sData               As String
    Dim aData()             As Byte
    Dim lDWord              As Long
    Dim i                   As Long
    Dim colItems            As Long
    Dim SubKeysNum          As Long
    Dim MaxSubKeyLen        As Long
    Dim numValues           As Long
    Dim MaxValueNameLen     As Long
    Dim MaxDataLen          As Long
    
    ' Si se pasa una cadena en sKey, esta función la convierte
    ' en un valor válido para la clave principal
    hKey = ParseKey(sKey, hKey)
    
    ' Abrir la clave indicada
    ' En este caso da igual el tipo de acceso,                      (12/Jun/99)
    ' pero... más vale prevenir
    'ret = RegOpenKeyEx(hKey, sKey, 0&, KEY_QUERY_VALUE, hKey2)
    ret = RegOpenKeyEx(hKey, sKey, 0&, KEY_READ, hKey2)
    
    ' Valores por defecto
    EnumValuesByType = True
    ReDim aData(0)
    lDWord = 0
    sData = ""
    
    ' Inicializar el array
    ' Añadir también el valor Default                               (23/Nov/00)
    colItems = 2
    ReDim colKeys(colItems)
    colKeys(1) = "(Default)"
    colKeys(2) = QueryRegBase(sKey, hKey)
    '
    ' Si todo va bien (se ha podido abrir la clave)
    If ret = ERROR_SUCCESS Then
        ' Obtener la información de esta clave, devolverá:
        ' SubKeysNum        Número de subclaves
        ' MaxSubKeyLen      Tamaño máximo de nombre de clave
        ' numValues         Número de valores en esta clave
        ' MaxValueNameLen   Tamaño máximo del nombre del valor
        ' MaxDataLen        Tamaño máximo de los datos
        'ret = RegQueryInfoKey(hKey2, 0&, 0&, 0&, SubKeysNum, MaxSubKeyLen, _
                            0&, numValues, MaxValueNameLen, _
                            MaxDataLen, 0&, lpftLastWriteTime)
        
        ' A ver si así funciona...                                  (12/Jun/99)
        ret = RegQueryInfoKey(hKey2, vbNullString, 0&, 0&, _
                            SubKeysNum, MaxSubKeyLen, _
                            0&, numValues, MaxValueNameLen, _
                            MaxDataLen, 0&, lpftLastWriteTime)
        
        ' Este es el error que me da el Windows 2000 Pro            (12/Jun/99)
        If ret = ERROR_INVALID_PARAMETERS Then
            Debug.Print "ERROR_INVALID_PARAMETERS"
            EnumValuesByType = False
            GoTo SalirEnumValues
        End If
        
        'lpName = String$(MaxValueNameLen + 1, 0)
        ' Hacer un bucle para el número de valores posibles
        For dwIndex = 0 To numValues
            lpName = String$(MaxValueNameLen + 1, Chr$(0))
            ' Llamarlo primero para saber el tipo de datos,
            ' el cual estará en retDT
            '//////////////////////////////////////////////////////////////////
            ' De esta forma en Win2000 produce un error de protección
            'ret = RegEnumValue(hKey2, dwIndex, ByVal 0&, ByVal 0&, 0&, retDT, ByVal 0&, ByVal 0&)
            'ret = RegEnumValue(hKey2, dwIndex, vbNullString, ByVal 0&, 0&, retDT, ByVal 0&, ByVal 0&)
            '//////////////////////////////////////////////////////////////////
            ret = RegEnumValue(hKey2, dwIndex, 0&, 0&, 0&, retDT, 0&, 0&)
            'ret = RegEnumValue(hKey2, dwIndex, lpName, Len(lpName), 0&, retDT, ByVal sData, lSize)
            ' la primera vez, cuando dwIndex = cero, devuelve ERROR_SUCCESS,
            ' pero después devuelve ERROR_MORE_DATA mientras haya datos.
            If ret = ERROR_MORE_DATA Or ret = ERROR_SUCCESS Then
                lSize = MaxDataLen
                ' Si es del tipo indicado o se busca REG_SZ
                ' y el tipo es cualquiera de los otros dos valores de cadena
                '--------------------------------------------------------------
                'TODO: Permitir indicar más de un tipo con OR
                '   Aunque puede ocurrir que se procesen tipos no deseados
                '   ya que, por ejemplo: REG_BINARY(3) y se procesaría si se indicara
                '   REG_SZ(1) y REG_EXPAND_SZ(2) por ejemplo.
                '--------------------------------------------------------------
                If (tDT = retDT) Or (tDT = REG_SZ And (retDT = REG_EXPAND_SZ Or retDT = REG_MULTI_SZ)) Then
                    Select Case retDT
                    Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
                        ' Datos de cadena
                        sData = String$(lSize, Chr$(0))
                        ret = RegEnumValue(hKey2, dwIndex, lpName, Len(lpName), 0&, retDT, ByVal sData, lSize)
                        If retDT = REG_MULTI_SZ Then
                            sData = RTrimZero(sData, True)
                        Else
                            sData = RTrimZero(sData)
                        End If
                        lpName = RTrimZero(lpName)
                        ReDim Preserve colKeys(colItems + 2)
                        colKeys(colItems + 1) = lpName
                        colKeys(colItems + 2) = sData
                        colItems = colItems + 2
                    Case REG_DWORD
                        ' Datos numéricos (long)
                        ret = RegEnumValue(hKey2, dwIndex, lpName, Len(lpName), 0&, retDT, lDWord, lSize)
                        sData = CStr(lDWord)
                        lpName = RTrimZero(lpName)
                        ReDim Preserve colKeys(colItems + 2)
                        colKeys(colItems + 1) = lpName
                        colKeys(colItems + 2) = sData
                        colItems = colItems + 2
                    'Case REG_BINARY
                    '    'Datos binarios
                    Case Else
                        ' Tratarlo como Binary
                        If lSize Then
                            ReDim aData(lSize)
                            ' Leer los datos binarios
                            ret = RegEnumValue(hKey2, dwIndex, lpName, Len(lpName), 0&, retDT, aData(0), lSize)
                            lpName = RTrimZero(lpName)
                            ' Al estilo de como se muestra con RegEdit
                            sData = ""
                            ' lSize ahora contendrá el número de bytes  (22/Nov/00)
                            ' almacenados en el búfer
                            For i = 0 To lSize - 1 'UBound(aData) - 1
                                sData = sData & Right$("00" & Hex$(aData(i)), 2) & " "
                            Next
                            ReDim Preserve colKeys(colItems + 2)
                            colKeys(colItems + 1) = lpName
                            colKeys(colItems + 2) = sData
                            colItems = colItems + 2
                        End If
                    End Select
                End If
            End If
        Next
    Else
        EnumValuesByType = False
    End If
SalirEnumValues:
    ' Cerrar la clave abierta
    ret = CloseKey(hKey2)
End Function

Public Function RegSaveKey(ByVal sKey As String, ByVal lpFile As String) As Long
Attribute RegSaveKey.VB_Description = "Guarda en un fichero el contenido de una clave con sus subclaves y datos, el formato no es el habitual de RegEdit"
Attribute RegSaveKey.VB_HelpID = 21090
    ' Guarda en un fichero el contenido de una clave, las subclaves y datos.
    ' Y funcionar, funciona, pero el fichero que da como resultado no es
    ' un fichero de texto...
    ' No he probado a asignar de nuevo el valor guardado, pero seguramente
    ' funcionará, lo que pasa es que no tiene un formato reconocido por
    ' RegEdit.exe (extensión .REG)
    '
    '$Por hacer: comprobar si esta función está bien...
    '
    ' Los atributos de seguridad se ignoran en Win95/98 (0&)
    ' si ese valor se usa en NT, se usarán los atributos por defecto...
    '
    ' Nombre a usar de forma temporal
    Const stmpFic As String = "\tmp.reg"
    Dim hKey As Long
    Dim hKey2 As Long
    Dim ret As eHKEYError
    
    ' Abrir la clave del registro
    hKey = ParseKey(sKey)
    
    ret = RegOpenKeyEx(hKey, sKey, 0&, 0&, hKey2)
    
    ' Guardarla en el fichero indicado
    ' como no se permiten nombres largos, se grabará en \tmp.reg
    ' y después se copiará en el nombre indicado.
    ' En Win95 se guarda con los atributos ReadOnly, Hide y System
    '
    ' La función falla si ya existe el fichero
    On Local Error Resume Next
    If Len(Dir$(stmpFic, vbHidden + vbReadOnly + vbSystem)) Then
        SetAttr stmpFic, vbNormal
        Kill stmpFic
    End If
    ret = RegSaveKeyA(hKey2, stmpFic, 0&)
    If ret = ERROR_SUCCESS Then
        ' Quitarle los atributos
        SetAttr stmpFic, vbNormal
        ' renombrar el fichero
        FileCopy stmpFic, lpFile
        ' borrar el temporal
        Kill stmpFic
    End If
    Err = 0
    RegCloseKey hKey2
End Function



Public Function UnRegister(ByVal sClass As String) As eHKEYError
    ' Des-Registrar un servidor ActiveX                             (05/Jul/99)
    ' Esta función quitará las entradas del registro de la clase indicada.
    ' El formato de la clave debe ser: Servidor.Clase
    '
    ' Devolverá 0 si todo fue bien (Cero es ERROR_NONE o ERROR_SUCCESS)
    ' sino, devolverá un código de error
    '
    ' Las claves del registro que se borrarán serán:
    ' HKEY_LOCAL_MACHINE\Software\Classes\Servidor.Clase
    '   En esta clave, bajo la clave Clsid, está el número a usar como {clsid}
    ' HKEY_LOCAL_MACHINE\Software\Classes\CLSID\{clsid}
    '   En esta clave, bajo la clave TypeLib, está el valor a usar como {TypeLib}
    ' HKEY_LOCAL_MACHINE\Software\Classes\Typelib\{TypeLib}
    '
    Dim sClave As String
    Dim sClsid As String
    Const sRootKey As String = "HKEY_LOCAL_MACHINE\Software\Classes\"
    Dim tKeyError As eHKEYError
    Dim sTypeLib As String
    
    UnRegister = ERROR_NONE ' También puede ser ERROR_SUCCESS
    
    ' Obtener el Clsid
    sClave = sRootKey & sClass & "\clsid"
    sClsid = GetRegString(sClave) ', "", HKEY_LOCAL_MACHINE)
    
    ' Avisará cuando no sea cierta
    'Debug.Assert Len(sClsid)
    
    If Len(sClsid) Then
        ' Borrar esta clave
        sClave = sRootKey & sClass
        tKeyError = DeleteKey(sClave)
        
        'Debug.Assert (tKeyError = ERROR_NONE)
        
        ' Sólo continuar si no da error
        If tKeyError = ERROR_NONE Then
            ' Eliminar las entradas de CLSID y Typelib
            ' Obtener el TypeLib
            sClave = sRootKey & "CLSID\" & sClsid & "\TypeLib"
            sTypeLib = GetRegString(sClave)
            
            'Debug.Assert Len(sTypeLib)
            
            If Len(sTypeLib) Then
                sClave = sRootKey & "CLSID\" & sClsid
                tKeyError = DeleteKey(sClave)
                
                'Debug.Assert (tKeyError = ERROR_NONE)
                
                ' Sólo continuar si no da error
                If tKeyError = ERROR_NONE Then
                    sClave = sRootKey & "TypeLib\" & sTypeLib
                    tKeyError = DeleteKey(sClave)
                    
                    'Debug.Assert (tKeyError = ERROR_NONE)
                End If
                UnRegister = tKeyError
            Else
                UnRegister = ERROR_FILE_NOT_FOUND 'ERROR_BADKEY
            End If
        End If
    Else
        UnRegister = ERROR_FILE_NOT_FOUND 'ERROR_BADKEY
    End If
End Function

Public Property Get UserName() As String
    UserName = sUser
End Property

