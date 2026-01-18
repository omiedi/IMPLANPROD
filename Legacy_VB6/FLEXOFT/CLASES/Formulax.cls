VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Formulax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"

Private Type tipogrilla
  FILA As Integer
  columnax As Integer
End Type
Dim Formulax As tipogrilla
Public Property Get filax() As Integer
   filax = Formulax.FILA
End Property
Public Property Let filax(VALOR As Integer)
   Formulax.FILA = VALOR
End Property

Public Property Get columnax() As Integer
   columnax = Formulax.columnax
End Property
Public Property Let columnax(VALOR As Integer)
   Formulax.columnax = VALOR
End Property



