Attribute VB_Name = "ASPEC_FUN"
Public SUPRIMECLICK%

Private Const HWND_BROADCAST = &HFFFF&
Private Const WM_FONTCHANGE = &H1D
Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFilename As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Enum TipoFuente
    [0. Default] = 0
    [1. Normal] = 1
    [2. Mediana] = 2
    [3. Negrita] = 3
End Enum

Dim FuenteAnchoFijoNormal_Cargada As Integer
Dim FuenteAnchoFijoMedium_Cargada As Integer
Dim FuenteAnchoFijoNegrita_Cargada As Integer

Function FuenteAnchoFijoInstalada(FUENTE As String) As Boolean

    FuenteAnchoFijoInstalada = False
    
    For i_Font = 0 To Screen.FontCount - 1
        If InStr(1, FUENTE, Screen.Fonts(i_Font)) > 0 Then
            FuenteAnchoFijoInstalada = True: Exit For
        End If
    Next
    
End Function

Sub AplicaFuenteAnchoFijo(Objeto As Object, Optional TipoFuente As TipoFuente, Optional TamanioFuente As Integer)
   
    Dim resultadoNormal As Long, resultadoMedium As Long, resultadoBold As Long
    
    ' AGREGA UNA NUEVA FONT
    resultadoNormal = FuenteAnchoFijoNormal_Cargada
    If FuenteAnchoFijoNormal_Cargada = 0 Then
            resultadoNormal = AddFontResource(LUDAT$ & "RobotoMono-Regular.ttf")
            If resultadoNormal > 0 Then
                ' ACTUALIZA LISTA DE FUENTES EN WINDOWS
                'SendMessage HWND_BROADCAST, WM_FONTCHANGE, 0, 0 'al entrar aca se cuelga en algunos casos (no se por que)
                FuenteAnchoFijoNormal_Cargada = 1
            End If
    End If
        
    resultadoMedium = FuenteAnchoFijoMedium_Cargada
    If FuenteAnchoFijoMedium_Cargada = 0 Then
            resultadoMedium = AddFontResource(LUDAT$ & "RobotoMono-Medium.ttf")
            If resultadoMedium > 0 Then
                ' ACTUALIZA LISTA DE FUENTES EN WINDOWS
                'SendMessage HWND_BROADCAST, WM_FONTCHANGE, 0, 0
                FuenteAnchoFijoMedium_Cargada = 1
            End If
    End If
    
    resultadoBold = FuenteAnchoFijoNegrita_Cargada
    If FuenteAnchoFijoNegrita_Cargada = 0 Then
            resultadoBold = AddFontResource(LUDAT$ & "RobotoMono-Bold.ttf")
            If resultadoBold > 0 Then
                ' ACTUALIZA LISTA DE FUENTES EN WINDOWS
                'SendMessage HWND_BROADCAST, WM_FONTCHANGE, 0, 0
                FuenteAnchoFijoNegrita_Cargada = 1
            End If
    End If
    
    ' VALORES POR DEFAULT
    If TipoFuente = [0. Default] And resultadoMedium > 0 Then
        Objeto.Font.Bold = False
        Objeto.FontName = "Roboto Mono Medium"
        Objeto.Font.Size = 10
    End If
    
    ' TIPO DE FUENTE
    If TipoFuente = [1. Normal] And resultadoNormal > 0 Then
        Objeto.Font.Bold = False
        Objeto.FontName = "Roboto Mono"
    End If
    
    If TipoFuente = [2. Mediana] And resultadoMedium > 0 Then
        Objeto.Font.Bold = False
        Objeto.FontName = "Roboto Mono Medium"
    End If
    
    If TipoFuente = [3. Negrita] And resultadoBold > 0 Then
        Objeto.Font.Bold = False
        Objeto.FontName = "Roboto Mono Bold"
    End If
    
    ' TAMAÑO
    If TamanioFuente > 0 Then Objeto.Font.Size = TamanioFuente

End Sub


Sub InstalaFuenteAnchoFijo()
   
    Dim resultadoNormal As Long, resultadoMedium As Long, resultadoBold As Long
    
    ' AGREGA UNA NUEVA FONT
    resultadoNormal = AddFontResource(LUDAT$ & "RobotoMono-Regular.ttf")
    If resultadoNormal > 0 Then
        ' ACTUALIZA LISTA DE FUENTES EN WINDOWS
        'SendMessage HWND_BROADCAST, WM_FONTCHANGE, 0, 0
        FuenteAnchoFijoNormal_Cargada = 1
    End If
    
    resultadoMedium = AddFontResource(LUDAT$ & "RobotoMono-Medium.ttf")
    If resultadoMedium > 0 Then
        ' ACTUALIZA LISTA DE FUENTES EN WINDOWS
        'SendMessage HWND_BROADCAST, WM_FONTCHANGE, 0, 0
        FuenteAnchoFijoMedium_Cargada = 1
    End If
    
    resultadoBold = AddFontResource(LUDAT$ & "RobotoMono-Bold.ttf")
    If resultadoBold > 0 Then
        ' ACTUALIZA LISTA DE FUENTES EN WINDOWS
        'SendMessage HWND_BROADCAST, WM_FONTCHANGE, 0, 0
        FuenteAnchoFijoNegrita_Cargada = 1
    End If

End Sub


Sub REMPLAGRID(FORMUL As Form, NGRID As CONTROL, OBJET As CONTROL, ENCALAB As Label, DETALIS As ListBox)
     '
     Exit Sub  ' aun no esta liberado para su utilización
     '
End Sub
Function ControlMenu%(NAMEFRM$, FORMULARIO As Form)
   ControlMenu% = 1
   If CONTROL$("CONFMENU", "MENUNOVI") = "SI" Then
        'SI ESTA CONFIGURADO Y NO ESTA EL ARCHIVO NO DEJA CONTINUAR POR QUE EL ARCHIVO DEBE EXISTIR AUNQUE SOLO TENGA
        'INFORMACION DE ALGUN FORMULARIO, SI EN ALGUN CASO SE OTORGAN TODOS LOS MENU SE DEBERA CAMBIAR LA CONFIGURACION
        'DEL CONTROL PARA QUE NO REALICE LA VALIDACION Y ENTRE DIRECTAEMNTE
        If EXISTE%(LUDAT$ + "CONFIGMENU.TXT") < 1 Then
         ControlMenu% = 0
         Call COMUNI("Imposible Continuar !!!\ Falta Configuración de Menú ")
         Exit Function
        Else
         TITUFORM$ = UCase$(NAMEFRM$)
         J& = 0
         Dim NAMERUTTXT$()
         NARCHI1% = FILEOPEN%(LUDAT$ + "CONFIGMENU.TXT", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
         Do While Not EOF(NARCHI1%) ' RECORRE
           Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
           XX1$ = TRIM$(UCase$(XX$))
           If XX1$ = "|INICIO " & TITUFORM$ Then LEER1% = 1
           If XX1$ = "|FINAL " & TITUFORM$ Then LEER1% = 0
           If LEER1% > 0 And TRIM$(XX1$) <> "" Then
              J& = J& + 1
              ReDim Preserve NAMERUTTXT$(J&)
              NAMERUTTXT$(J&) = XX1$
           ElseIf J& > 0 Then 'SI J ES > 0 POR QUE YA ESTUVO ESCRIBIENDO EL VECTOR Y SI LLEGA A FINAL SALE
             Exit Do
           End If
         Loop
         Close #NARCHI1%
         '
         If J& < 1 Then ControlMenu% = 1: Exit Function
         '
         For Each CONTROLX In FORMULARIO.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
            If TypeOf CONTROLX Is menu Then
              Set MENUX = CONTROLX
              NAMECTRL$ = UCase(MENUX.Name)
              On Error Resume Next
              For L& = 1 To UBound(NAMERUTTXT$)
                   If NAMECTRL$ = NAMERUTTXT$(L&) Then
                      MENUX.Visible = False
                   End If
              Next L&
              On Error GoTo 0
            End If
          Next
        End If
   End If
         

End Function
