VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Formulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type tipogrilla
  fila As Integer
  columnax As Integer
End Type
Public formulax As tipogrilla


Public Property Get filax() As Integer
   fila = formulax.fila
End Property
Public Property Let filax(valor As Integer)
   formulax.fila = valor
End Property

Public Property Get columnax() As Integer
   columnax = formulax.columnax
End Property
Public Property Let columnax(valor As Integer)
   formulax.columnax = valor
End Property

