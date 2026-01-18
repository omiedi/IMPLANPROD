VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nPerifericos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub WebCam_Inicia(ByRef oPer As ePerifericos)

    Dim temp As Long
    
    hwdc = capCreateCaptureWindow("CapWindow", ws_child Or ws_visible, _
                                    0, 0, 640, 480, oPer.WEBCAM.ObjetoReproduce.hwnd, 0)
    If (hwdc <> 0) Then
        temp = SendMessage(hwdc, wm_cap_driver_connect, 0, 0)
        temp = SendMessage(hwdc, wm_cap_set_preview, 1, 0)
        temp = SendMessage(hwdc, WM_CAP_SET_PREVIEWRATE, 30, 0)
        temp = SendMessage(hwdc, WM_CAP_SET_SCALE, True, 0)
        'esto hace que la imagen recibida por el dispositivo se ajuste
        'al tamaño de la ventana de captura
        DoEvents
        startcap = True
        oPer.WEBCAM.hwdcReproduce = hwdc
    Else
        MsgBox "No hay Camara Web", 48, "Error"
  End If

End Sub

Sub WebCam_Desconecta(oPer As ePerifericos)

    temp = DestroyWindow(oPer.WEBCAM.hwdcReproduce)
    If startcap = True Then
        temp = SendMessage(oPer.WEBCAM.hwdcReproduce, WM_CAP_DRIVER_DISCONNECT, 0&, 0&)
        DoEvents
        startcap = False
    End If
    
End Sub

Sub WebCam_Formato(oPer As ePerifericos)

    If startcap = True Then
        temp = SendMessage(oPer.WEBCAM.hwdcReproduce, WM_CAP_DLG_VIDEOFORMAT, 0&, 0&)
        DoEvents
    End If
        
End Sub

Sub WebCam_Configuracion(oPer As ePerifericos)

    Dim temp As Long
    If startcap = True Then
        temp = SendMessage(oPer.WEBCAM.hwdcReproduce, WM_CAP_DLG_VIDEOCONFIG, 0&, 0&)
        DoEvents
    End If
    
End Sub

Sub WebCam_Resolucion(oPer As ePerifericos)

    Dim bmp As BITMAPINFO
    Dim lWidth As Long
    Dim lHeight As Long
    
    lWidth = 640
    lHeight = 480
    On Error GoTo Err

        With bmp.bmiHeader
            .biSize = Len(bmp.bmiHeader)
            .biWidth = lWidth
            .biHeight = lHeight
            .biPlanes = 1
            .biBitCount = 24
        End With
        
    On Error GoTo 0
    
    temp = SendMessageAsBitMap(oPer.WEBCAM.hwdcReproduce, WM_CAP_SET_VIDEOFORMAT, Len(bmp), bmp)
    
    Exit Sub

Err:
Exit Sub

End Sub

Sub WebCam_CapturaImagen(oPer As ePerifericos)

    ' borra el portapapeles
    Clipboard.Clear
    'copiar un clip al portapapeles
    temp = SendMessage(oPer.WEBCAM.hwdcReproduce, WM_CAP_EDIT_COPY, 1, 0)
    DoEvents
    ' Si el formato del clipboard es un bitmap
    If Clipboard.GetFormat(vbCFBitmap) Then
      oPer.WEBCAM.ObjetoCaptura.Picture = Clipboard.GetData(vbCFBitmap)
    Else
      Call COMUNI("Problemas al Capturar Imagen.\ \Cierre Esta Ventana y Vuelva a Intentarlo.")
    End If

End Sub

Sub WebCam_GuardaCaptura(oPer As ePerifericos)

    nombreArchivo$ = TRIM$(oPer.WEBCAM.NombreArchivoCaptura)
    If nombreArchivo$ = "" Then nombreArchivo$ = "imagen"
    RutaGuardaCaptura$ = TRIM$(oPer.WEBCAM.RutaGuardaCaptura)
    If RutaGuardaCaptura$ = "" Then RutaGuardaCaptura$ = LUDAT$
    
    ' GUARDA IMAGEN ORIGINAL EN FORMATO BMP
    SavePicture oPer.WEBCAM.ObjetoCaptura, RutaGuardaCaptura$ & "\" & nombreArchivo$ & ".bmp"
    
    fMain.Label2 = RutaGuardaCaptura$ + "\" + nombreArchivo$ & ".bmp"
    fMain.mnuOpen_Click
    Dim SaveForm As fSaveJPG
    Set SaveForm = New fSaveJPG
    SaveForm.SaveImage fMain.m_Image, RutaGuardaCaptura$ + "\" + nombreArchivo$ & ".jpg"
    SaveForm.cmdFinish_Click (1)
    Set SaveForm = Nothing
    
    Call FILEKILL(RutaGuardaCaptura$ & "\" & nombreArchivo$ & ".bmp")
    
End Sub
