VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWebCam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public RutaGuardaCaptura As String
Public NombreArchivoCaptura As String
Public ObjetoReproduce As PictureBox
Public hwdcReproduce As Long
Public ObjetoCaptura As PictureBox
Public hwdcCaptura As Long

