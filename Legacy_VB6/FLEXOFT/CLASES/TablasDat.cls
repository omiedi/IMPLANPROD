VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TablasDat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type TDat
    EcoTadepos As String
    EcoVendedor As String
    EcoSecprod As String
    EcoConpag As String
    EcoGrucocli As String
    EcoCaretib As String
    EcoCaposiv As String
    EcoProvinc As String
    EcoListas As String
    RetornoOk As Integer
End Type
Private EcosGeneral As TDat 'este es para usar otras variables que no sean vectores.

'Declaramos el vector
Private EcosTadepos() As TDat
Private EcosVendedor() As TDat
Private EcosSecprod() As TDat
Private EcosConpag() As TDat
Private EcosGrucocli() As TDat
Private EcosCaretib() As TDat
Private EcosCaposiv() As TDat
Private EcosProvinc() As TDat
Private EcosListas() As TDat

Public Property Get EcoTadepos(i As Variant) As String
   EcoTadepos = EcosTadepos(i).EcoTadepos
End Property
Public Property Get EcoVendedor(i As Variant) As String
   EcoVendedor = EcosVendedor(i).EcoVendedor
End Property
Public Property Get EcoSecprod(i As Variant) As String
   EcoSecprod = EcosTadepos(i).EcoSecprod
End Property
Public Property Get EcoConpag(i As Variant) As String
   EcoConpag = EcosConpag(i).EcoConpag
End Property
Public Property Get EcoListas(i As Variant) As String
   EcoListas = EcosListas(i).EcoListas
End Property
Public Property Get EcoGrucocli(i As Variant) As String
   EcoGrucocli = EcosGrucocli(i).EcoGrucocli
End Property
Public Property Get EcoCaretib(i As Variant) As String
   EcoCaretib = EcosCaretib(i).EcoCaretib
End Property
Public Property Get EcoCaposiv(i As Variant) As String
   EcoCaposiv = EcosCaposiv(i).EcoCaposiv
End Property
Public Property Get EcoProvinc(i As Variant) As String
   EcoProvinc = EcosProvinc(i).EcoProvinc
End Property
Public Property Let RetornoOk(valor As Variant)
   EcosGeneral.RetornoOk = valor
End Property
Public Property Get RetornoOk() As Variant
   RetornoOk = EcosGeneral.RetornoOk
End Property
Public Sub CarVectorTadepos()
  For U& = 1 To ULTREG&("TADEPOSI")
      X$ = REGLEIDO$("TADEPOSI", U&)
      DepOrig% = Asc(Left$(X$, 1))
      DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DepOrig%))
      ReDim Preserve EcosTadepos(U&)
      EcosTadepos(DepOrig%).EcoTadepos = DEPOSI$
  Next U&
  Call CIERRARCH("TADEPOSI")
End Sub

Public Sub CarVectorTaVendedor()
  For U& = 1 To ULTREG&("VENDEDOR")
      X$ = REGLEIDO$("VENDEDOR", U&)
      CodVend% = Asc(Left$(X$, 1))
      On Error Resume Next
      If UBound(EcosVendedor) < CodVend% Then ReDim Preserve EcosVendedor(CodVend%)
      EcosVendedor(CodVend%).EcoVendedor = ECOARCH$("VENDEDOR", Chr$(CodVend%))
      On Error GoTo 0
  Next U&
  Call CIERRARCH("VENDEDOR")
End Sub
Public Sub CarVectorTaSecprod()
    For U& = 1 To ULTREG&("SECPROD")
        X$ = REGLEIDO$("SECPROD", U&)
        COSECC = CVS(Left$(X$, 4))
        ReDim Preserve EcosSecprod(COSECC)
        EcosSecprod(COSECC).EcoSecprod = ECOARCH$("SECPROD", MKS$(COSECC))
    Next U&
    Call CIERRARCH("SECPROD")
End Sub
Public Sub CarVectorTaConpag()
    For U& = 1 To ULTREG&("CONPAG")
      X$ = REGLEIDO$("CONPAG", U&)
      INDX% = Asc(Left$(X$, 1))
      ReDim Preserve EcosConpag(U&)
      EcosConpag(U&).EcoConpag = ECOARCH$("CONPAG", Chr$(INDX%))
    Next U&
    Call CIERRARCH("CONPAG")
End Sub
Public Sub CarVectorTaListas()
    For U& = 1 To ULTREG&("LISTAS")
      X$ = REGLEIDO$("LISTAS", U&)
      INDX% = Asc(Left$(X$, 1))
      ReDim Preserve EcosListas(INDX%)
      EcosListas(INDX%).EcoListas = ECOARCH$("LISTAS", Chr$(INDX%))
    Next U&
    Call CIERRARCH("LISTAS")
End Sub

Public Sub CarVectorTaGrucocli()
    For U& = 1 To ULTREG&("GRUCOCLI")
      X$ = REGLEIDO$("GRUCOCLI", U&)
      INDX = Asc(Left$(X$, 1))
      Descr$ = TRIM$(Mid(X$, 2))
      On Error Resume Next
      CantIndices = UBound(EcosGrucocli)
      On Error GoTo 0
      If INDX > CantIndices Then
        'ReDim Preserve EcosGrucocli(U&)
        ReDim Preserve EcosGrucocli(INDX)
      End If
      EcosGrucocli(INDX).EcoGrucocli = Descr$
    Next U&
    Call CIERRARCH("GRUCOCLI")
End Sub
Public Sub CarVectorTaCaposiv()
    For U& = 0 To ULTREG&("CAPOSIV")
      X$ = REGLEIDO$("CAPOSIV", U&)
      INDX = Asc(Left$(X$, 1))
      ReDim Preserve EcosCaposiv(INDX)
      EcosCaposiv(INDX).EcoCaposiv = ECOARCH$("CAPOSIV", Chr$(INDX))
    Next U&
    Call CIERRARCH("CAPOSIV")
End Sub
Public Sub CarVectorTaCaretib()
    ASS = ULTREG&("CARETIB")
    For U& = 1 To ULTREG&("CARETIB")
      X$ = REGLEIDO$("CARETIB", U&)
      INDX% = Left(X$, 1)
      ReDim Preserve EcosCaretib(INDX%)
      EcosCaretib(INDX%).EcoCaretib = Mid$(X$, 2)
    Next U&
    Call CIERRARCH("CARETIB")
End Sub
Public Sub CarVectorTaProvinc()
    For U& = 1 To ULTREG&("PROVINC")
      X$ = REGLEIDO$("PROVINC", U&)
      PR1$ = TRIM$(UCase$(Mid$(X$, 1, 1)))
      If Asc(PR1$) > 64 And Asc(PR1$) < 91 Then
        INDX% = Asc(PR1$) - 64 'PARA QUE COINCIDA EL INDICE CON EL ASCCI - 64
        If INDX% > 0 Then
          On Error Resume Next
          If UBound(EcosProvinc) < INDX% Then ReDim Preserve EcosProvinc(INDX%)
          EcosProvinc(INDX%).EcoProvinc = TRIM$(ECOARCH$("PROVINC", PR1$))
          On Error GoTo 0
        End If
      End If
    Next U&
    Call CIERRARCH("PROVINC")
End Sub
Public Sub CarVectGeneral(TablaDat$)
   TablaDat$ = UCase$(TRIM$(TablaDat$))
   Select Case TablaDat$
     Case "VENDEDOR"
       Call CarVectorTaVendedor
     Case "TADEPOSI"
       Call CarVectorTadepos
     Case "SECPROD"
       Call CarVectorTaSecprod
     Case "CONPAG"
       Call CarVectorTaConpag
     Case "GRUCOCLI"
       Call CarVectorTaGrucocli
     Case "CARETIB"
       Call CarVectorTaCaretib
     Case "CAPOSIV"
       Call CarVectorTaCaposiv
     Case "PROVINC"
       Call CarVectorTaProvinc
     Case "LISTAS"
       Call CarVectorTaListas
     Case "PROVINC"
       Call CarVectorTaProvinc
     Case Else
       RetornoOk = -1
   End Select
End Sub
Public Function ValorTblDat(TablaDat$, ArgBus$)
   TablaDat$ = UCase$(TRIM$(TablaDat$))
   RetornoOk = 1
   On Error Resume Next
   Select Case TablaDat$
     Case "VENDEDOR"
       CodVend% = XVALO(ArgBus$)
       ValorTblDat = EcosVendedor(CodVend%).EcoVendedor
     Case "TADEPOSI"
       DepOrig% = XVALO(ArgBus$)
       ValorTblDat = EcosTadepos(DepOrig%).EcoTadepos
     Case "SECPROD"
       INDX% = XVALO(ArgBus$)
       ValorTblDat = EcosSecprod(INDX%).EcoSecprod
     Case "CONPAG"
       COSECC = TRIM$(ArgBus$)
       ValorTblDat = EcosConpag(COSECC).EcoConpag
     Case "GRUCOCLI"
       COSECC = XVALO(ArgBus$)
       ValorTblDat = EcosGrucocli(COSECC).EcoGrucocli
     Case "CARETIB"
       COSECC = TRIM$(ArgBus$)
       ValorTblDat = EcosCaretib(COSECC).EcoCaretib
     Case "CAPOSIV"
       COSECC = TRIM$(ArgBus$)
       ValorTblDat = EcosCaposiv(COSECC).EcoCaposiv
     Case "LISTAS"
       COSECC = XVALO(ArgBus$)
       ValorTblDat = EcosListas(COSECC).EcoListas
     Case "PROVINC"
       PR1$ = TRIM$(ArgBus$)
       If Asc(PR1$) > 64 And Asc(PR1$) < 91 Then
        INDX% = Asc(PR1$) - 64 'PARA QUE COINCIDA EL INDICE CON EL ASCCI - 64
        If INDX% > 0 Then
          On Error Resume Next
          ValorTblDat = EcosProvinc(INDX%).EcoProvinc
          On Error GoTo 0
        End If
      End If

'       ValorTblDat = EcosProvinc(COSECC).EcoProvinc
     Case Else
       RetornoOk = -1
   End Select
   On Error GoTo 0
End Function



