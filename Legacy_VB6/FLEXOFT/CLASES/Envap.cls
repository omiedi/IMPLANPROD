VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Envap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Private Type Envaplax
        CantVecesPart As Integer
        CantTotParticipantes As Integer
        HoraMenor As Double
        HoraMayor As Double
        NumParticipante As Integer
        NameParticipante As String
        HORASHOMBRE As Double
        CantTotHoras_X_Particip As Double
    End Type
    Private ENVA  As Envaplax
    Private Type ENVX
        CantTotParticipantes As Integer
    End Type
    Private env  As Envaplax

   'Declaramos el vector
   Private ENVAX() As Envaplax

Sub ALTA_ARTICULO_WALDBOTT(CI1%)
        CODI$ = CODEXT$(CI1%)
        DESXRI$ = DESCRIT0$(CI1%)
        DESXRI$ = REPLACAR(DESXRI$, "'", " ") 'POR SI TIENE COMILLAS SIMPLES
        '
        SQLX$ = "INSERT INTO ARTICULO  "
        SQLX$ = SQLX$ + "(NUMERO ,DESCRIP,FECALTA,MARCA,PRESVENTA,PRESCOMPRA,DESCUENTO,IMPUESTOS,GRCOMISION,Tipo,ORIGEN,INVENTARIO,FACTURABLE)"
        SQLX$ = SQLX$ + " VALUES('" & CODI$ & "'"
        SQLX$ = SQLX$ + ",'" & Left$(DESXRI$, 50) & "'"
        SQLX$ = SQLX$ + ",'" & NOWSERVER & "'"
        'DE AQUI HACIA ABAJO DATOS FIJOS QUE HAY QUE PASAR ENVIADO POR GABRIEL MAIL 29/05/19
        SQLX$ = SQLX$ + ",'000'"
        SQLX$ = SQLX$ + ",'01'"
        SQLX$ = SQLX$ + ",'01'"
        SQLX$ = SQLX$ + ",0"
        SQLX$ = SQLX$ + ",'01'"
        SQLX$ = SQLX$ + ",'1'"
        SQLX$ = SQLX$ + ",'P'"
        SQLX$ = SQLX$ + ",'NA'"
        SQLX$ = SQLX$ + ",'N'"
        SQLX$ = SQLX$ + ",'S'"
        SQLX$ = SQLX$ + ")"
        
        Call abreBaDaWald
        BASEWALD.Execute (SQLX$)
End Sub

Sub ALTA_ARTICULO_CONTEPLAST_WALDBOTT(CODI$, DESXRI$)
'        CODI$ = CODEXT$(CI1%)
'        DESXRI$ = DESCRIT0$(CI1%)
        '
        Dim CTP As New Conteplast
        SQLX$ = "INSERT INTO ARTICULO  "
        SQLX$ = SQLX$ + "(NUMERO ,DESCRIP,FECALTA,MARCA,PRESVENTA,PRESCOMPRA,DESCUENTO,IMPUESTOS,GRCOMISION,Tipo,ORIGEN,INVENTARIO,FACTURABLE)"
        SQLX$ = SQLX$ + " VALUES('" & CODI$ & "'"
        SQLX$ = SQLX$ + ",'" & Left$(DESXRI$, 50) & "'"
        SQLX$ = SQLX$ + ",'" & NOWSERVER & "'"
        'DE AQUI HACIA ABAJO DATOS FIJOS QUE HAY QUE PASAR ENVIADO POR GABRIEL MAIL 29/05/19
        SQLX$ = SQLX$ + ",'000'"
        SQLX$ = SQLX$ + ",'01'"
        SQLX$ = SQLX$ + ",'01'"
        SQLX$ = SQLX$ + ",0"
        SQLX$ = SQLX$ + ",'01'"
        SQLX$ = SQLX$ + ",'1'"
        SQLX$ = SQLX$ + ",'P'"
        SQLX$ = SQLX$ + ",'NA'"
        SQLX$ = SQLX$ + ",'N'"
        SQLX$ = SQLX$ + ",'S'"
        SQLX$ = SQLX$ + ")"
        Call abreBaDaWald
        BASEWALD.Execute (SQLX$)
Set CTP = Nothing
End Sub
Sub MARCARREMITOTRANSFERIDO(REMIX$)
       ' AGREGA UNA FILA INDICANDO QUE ESTE REMITO YA ESTA TRANSFERIDO A WALDBOTT
'        SQLX$ = "UPDATE MOVTRANSFERIDOS"
'        SQLX$ = SQLX$ + " SET STAT_TRANSF = 1"
'
'        SQLX$ = SQLX$ + " WHERE DOSECLAR = '" & REMIX$ & "'"
'        BASEWALD.Execute (SQLX$)
'        On Error Resume Next
        '
        SQLX$ = "INSERT INTO MOVTRANSFERIDOS  "
        SQLX$ = SQLX$ + "(DOPRILAR ,STAT_TRANSF,FECHA_TRANF)"
        SQLX$ = SQLX$ + " VALUES('" & REMIX$ & "'"
        SQLX$ = SQLX$ + ",'1'"
        SQLX$ = SQLX$ + ",'" & NOWSERVER & "'"
        SQLX$ = SQLX$ + ")"
        
        Call abreBaMDBINTER_REMI_ENVAP
        BASEINTERREMI_ENVAP.Execute (SQLX$)
        If BASEINTERREMI_ENVAP.State = adStateOpen Then
           BASEINTERREMI_ENVAP.Close
        End If

End Sub

Sub SINCRONIZA_CondVenta_Waldbott()
        '
        Screen.MousePointer = 11
        Call CREATABLA_CONDIVENTA
        Call BORRAREGISTROS("LICONPAG", "LICONPAG_ID >=0", REGAF&, "NOMESS")
        
        SQLX$ = "SELECT * FROM LICONPAG WHERE LICONPAG_ID < 0"
        Dim RSF As ADODB.Recordset
        Set RSF = New ADODB.Recordset
27      On Error Resume Next
        RSF.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 27
        End If
        On Error GoTo 0
        KLX$ = "SI081"
        CONDI$ = "CLAVE LIKE '" & KLX & "%'"
        SQLX$ = "SELECT CLAVE, ALFA01 FROM TABLAS WHERE " & CONDI$
        SQLX$ = SQLX$ + " ORDER BY CLAVE"
        Call abreBaDaWald
        Dim RS As New ADODB.Recordset
        Set RS = BASEWALD.Execute(SQLX$)
        Do While Not RS.EOF
          CLAVEX$ = Right(TRIM$(RS!CLAVE), 2)
          REFE$ = SAFETEXT$(RS!ALFA01)
          With RSF
              .AddNew   ' YA CONVERTIDO
              On Error Resume Next
              !CODIEMPR = CodiEmp%
              On Error GoTo 0
              !CODICONPAG = XVALO(CLAVEX$)
              !REFERENCIA = REFE$
              .Update
          End With
          RS.MoveNext
        Loop
        RS.Close
        Set RS = Nothing
        RSF.Close
        Set RSF = Nothing
        Screen.MousePointer = 1
        Call COMUNI("Tabla de Condiciones de Pagos Sincronizada")
End Sub

Sub UPDATE_ARTICULO_WALDBOTT(CI1%)
        DESXRI$ = DESCRIT0$(CI1%)
        CODI$ = CODEXT$(CI1%)
        '
        SQLX$ = "UPDATE ARTICULO  "
        SQLX$ = SQLX$ + "SET DESCRIP = '" & Left(DESXRI$, 50) & "', "
        SQLX$ = SQLX$ + " MARCA = '000', "
        SQLX$ = SQLX$ + " PRESVENTA = '01', "
        SQLX$ = SQLX$ + " PRESCOMPRA = '01', "
        SQLX$ = SQLX$ + " DESCUENTO = 0, "
        SQLX$ = SQLX$ + " IMPUESTOS = '01', "
        SQLX$ = SQLX$ + " GRCOMISION = '1', "
        SQLX$ = SQLX$ + " Tipo = 'P', "
        SQLX$ = SQLX$ + " ORIGEN = 'NA', "
        SQLX$ = SQLX$ + " INVENTARIO = 'N', "
        SQLX$ = SQLX$ + " FACTURABLE = 'S' "
        
        SQLX$ = SQLX$ + " WHERE NUMERO = '" & CODI$ & "'"
        Call abreBaDaWald
        BASEWALD.Execute (SQLX$)
End Sub

Sub UPDATE_ARTICULO_CONTEPLAST_WALDBOTT(CODI$, DESXRI$)
Dim CTP As New Conteplast
'        DESXRI$ = DESCRIT0$(CI1%)
'        CODI$ = CODEXT$(CI1%)
        '
        DESXRI$ = REPLACAR(DESXRI$, "'", " ") 'POR SI TIENE COMILLAS SIMPLES

        SQLX$ = "UPDATE ARTICULO  "
        SQLX$ = SQLX$ + "SET DESCRIP = '" & Left(DESXRI$, 50) & "', "
        SQLX$ = SQLX$ + " MARCA = '000', "
        SQLX$ = SQLX$ + " PRESVENTA = '01', "
        SQLX$ = SQLX$ + " PRESCOMPRA = '01', "
        SQLX$ = SQLX$ + " DESCUENTO = 0, "
        SQLX$ = SQLX$ + " IMPUESTOS = '01', "
        SQLX$ = SQLX$ + " GRCOMISION = '1', "
        SQLX$ = SQLX$ + " Tipo = 'P', "
        SQLX$ = SQLX$ + " ORIGEN = 'NA', "
        SQLX$ = SQLX$ + " INVENTARIO = 'N', "
        SQLX$ = SQLX$ + " FACTURABLE = 'S' "
        
        SQLX$ = SQLX$ + " WHERE NUMERO = '" & CODI$ & "'"
        Call abreBaDaWald
        BASEWALD.Execute (SQLX$)
Set CTP = Nothing

End Sub
    Public Property Get NumParticipante(i As Variant) As Integer
       NumParticipante = ENVAX(i).NumParticipante
    End Property
    Public Property Let NumParticipante(i As Variant, ByVal valor As Integer)
       ENVAX(i).NumParticipante = valor
    End Property

    Public Property Get NameParticipante(i As Variant) As String
       NameParticipante = ENVAX(i).NameParticipante
    End Property
    Public Property Let NameParticipante(i As Variant, ByVal valor As String)
       ENVAX(i).NameParticipante = valor
    End Property
    Public Property Get CantVecesPart(i As Variant) As Integer
       CantVecesPart = ENVAX(i).CantVecesPart
    End Property
    Public Property Let CantVecesPart(i As Variant, ByVal valor As Integer)
       ENVAX(i).CantVecesPart = valor
    End Property
    Public Property Get CantTotParticipantes() As Integer
       CantTotParticipantes = env.CantTotParticipantes
    End Property
    Public Property Let CantTotParticipantes(valor As Integer)
       env.CantTotParticipantes = valor
    End Property
    Public Property Get HoraMenor() As Double
       HoraMenor = ENVA.HoraMenor
    End Property
    Public Property Let HoraMenor(valor As Double)
       ENVA.HoraMenor = valor
    End Property
    Public Property Get HoraMayor() As Double
       HoraMayor = ENVA.HoraMayor
    End Property
    Public Property Let HoraMayor(valor As Double)
       ENVA.HoraMayor = valor
    End Property
    Public Property Get CantTotHoras_X_Particip() As Double
       CantTotHoras_X_Particip = env.CantTotHoras_X_Particip
    End Property
    Public Property Let CantTotHoras_X_Particip(valor As Double)
       env.CantTotHoras_X_Particip = valor
     '---- DEVUELVE EL TOTAL DE HORAS X CANTIDAD DE PARITCIPANTES CONSIDERANDO REGISTRO A REGISTRO
     '---- EJ. SI DOS OPERIOS DECLARAN 2 DEVUELVE 4, SIGUIENTE REGISTRO 3 OPERAR. DECLARAN
     '---- 3 HORAS DEVULEVE 6 Y ACUMULA 6 + 4, SI FUERAN DOS REGISTROS DEVOLVERIA 10
    End Property
'    Public Property Get HORASHOMBRE() As Double
'       HORASHOMBRE = env.HORASHOMBRE
'    End Property
'    Public Property Let HORASHOMBRE(VALOR As Double)
'       env.HORASHOMBRE = VALOR
'    End Property

Public Sub CarVectParticipantes(NORFA As String, Optional INDISOLSUPX As Integer, Optional EXCLUYEMAQ%, Optional FE%, Optional MAQ$)
'---SI NO SE QUIERE CONSIDERAR LA SOLAPA PASAR COMO PARAMETRO (-1)

'EJEMPLO DE LLAMADA: Call env.CarVectParticipantes(NROOF$, GLOBALFRM_SOLASUP%)

'ESTA RUTINA CARGA LAS SIGUIENTES VARIABLES Y VECTORES QUE SE DETALLA ABAJO

'CONSIDERA LAS DOS TABLAS DE PRODUCCION NO LA DOTACION QUE SE CARGA APARTE.

'   A = ENVX.CantTotParticipantes ' -----DEVUELVE LA CANTIDAD TOTAL DE PARTICIPANTES CONSIDERANDO C/U SOLO UNA VEZ SIN INTERESAR SI SE REPITE

'   B = env.CantVecesPart(USER%) '------DEVUELVE LA CANTIDAD TOTAL DE VECES QUE SE REPITE  UN PARITICIPANTE INTERVINIDENDO EN LA OF Y SOLAPA

'   C = env.NameParticipante(USER%)'----DEVUELVE EL NOMBER DEL PARITICIPANTE SOLO ESCRIBE LOS QUE PARTICIPARON

'   D = ENVX.CantTotHoras_X_Particip ' ---- DEVUELVE EL TOTAL DE HORAS X CANTIDAD DE PARITCIPANTES CONSIDERANDO REGISTRO A REGISTRO
   'FILTRADO POR ORDEN DE FABRICACION  ---  EJ. SI DOS OPERIOS DECLARAN 2 DEVUELVE 4, SIGUIENTE REGISTRO 3 OPERAR. DECLARAN
                                    ' ---- 3 HORAS DEVULEVE 6 Y ACUMULA 6 + 4, SI FUERAN DOS REGISTROS DEVOLVERIA 10
                                  
                                  
'   Set env = Nothing----------------APLICAR LUEGO DE LA LLAMADA NOTHING PARA DESCARGAR LAS VARIABLES
On Error GoTo ERROR_CONTROL  'ESTABA DANDO UN ERROR AL EMITIR EN ENVA DEL 1 AL 07 DE ENERO 2025 SELECCIONADO SEL 1 SELLADORA 1 UN ERROR 9 INDICE FUERA DE RANGO
                             'NO DETECTE CUAL ES EL PROBLEMA LO QUE  HACE CON EL ON ERROR QUE ESE DATO SI ESTA MAL CARGADO NO LO VA A TRAER PERO NO DA ERROR NI DETIENE LA EJECUCION

        CantTotParticipantes = 0: TOTHORAX = 0
        Erase ENVAX()
        ReDim Preserve ENVAX(1024)
        '
        SQLX$ = "SELECT  P.H_DESDE,P.H_HASTA,P.NUMAYUD1,P.NUMAYUD2,P.NUMAYUD3,P.NUMAYUD4,P.NUMAYUD5,P.NUMAYUD6,P.NUMAYUD7,P.NUMAYUD8 "
        SQLX$ = SQLX$ + " ,P.FECHA,P.IDSUBOR,P.NULIDOP1,P.NULIDOP2,P.SOLAPASUP,P.PDIARPROD1_ID,P.INDISOLSUP,P.NUMUSUMAQ,P.H_TRABAJADAS "
        SQLX$ = SQLX$ + " ,P.NAMEUSUMAQ,P.NAMEAYUD1,P.NAMEAYUD2,P.NAMEAYUD3,P.NAMEAYUD4 "
        SQLX$ = SQLX$ + " ,P.NAMEAYUD5,P.NAMEAYUD6,P.NAMEAYUD7,P.NAMEAYUD8,P.NULNAMEOP1,P.NULNAMEOP2,P.DESC_P_DIAR,P.NUMMAQUI"
        SQLX$ = SQLX$ + " FROM (SELECT H_DESDE,H_HASTA,NUMAYUD1,NUMAYUD2 , NUMAYUD3, NUMAYUD4, NUMAYUD5"
        SQLX$ = SQLX$ + " , NUMAYUD6, NUMAYUD7, NUMAYUD8, FECHA, IDSUBOR, NULIDOP1, NULIDOP2 "
        SQLX$ = SQLX$ + " ,SOLAPASUP, INDISOLSUP,PDIARPROD1_ID,NUMUSUMAQ,H_TRABAJADAS,NAMEUSUMAQ,NAMEAYUD1,NAMEAYUD2, "
        SQLX$ = SQLX$ + " NAMEAYUD3,NAMEAYUD4,NAMEAYUD5,NAMEAYUD6,NAMEAYUD7,NAMEAYUD8,NULNAMEOP1,NULNAMEOP2,DESC_P_DIAR,NUMMAQUI "
        SQLX$ = SQLX$ + " FROM PDIARPROD1 "
        
        SQLX$ = SQLX$ + " union all "
        
        SQLX$ = SQLX$ + " SELECT H_DESDE,H_HASTA,NULNUMAYUD1,NULNUMAYUD2,NULNUMAYUD3,NULNUMAYUD4,NULNUMAYUD5,NULNUMAYUD6"
        SQLX$ = SQLX$ + " ,NULNUMAYUD7,NULNUMAYUD8,FECHA,IDSUBOR,IDOP1,IDOP2,SOLAPASUP"
        SQLX$ = SQLX$ + " ,INDISOLSUP,PDIARPROD2_ID,NULNUMUSUMAQ,H_TRABAJADAS,NULNAMEUSUMAQ,NULNAMEAYUD1"
        SQLX$ = SQLX$ + " ,NULNAMEAYUD2,NULNAMEAYUD3,NULNAMEAYUD4,NULNAMEAYUD5,NULNAMEAYUD6,NULNAMEAYUD7,NULNAMEAYUD8"
        SQLX$ = SQLX$ + " ,NAMEOP1 , NAMEOP2,DESCRIPTAREA,NUMMAQUI"
        SQLX$ = SQLX$ + " FROM PDIARPROD2)"
        SQLX$ = SQLX$ + " AS P  "
        SQLX$ = SQLX$ + " WHERE  P.IDSUBOR = '" & NORFA$ & "'"
        If INDISOLSUPX >= 0 Then SQLX$ = SQLX$ + " AND  P.INDISOLSUP = " & INDISOLSUPX%
        If FE% > 0 Then SQLX$ = SQLX$ + " AND FECHA = '" & FETEXCOR1$(FE%) & "'"
        If TRIM$(MAQ$) <> "" Then SQLX$ = SQLX$ + " AND NUMMAQUI = '" & MAQ$ & "'"
        SQLX$ = SQLX$ + " ORDER BY P.DESC_P_DIAR"
        Set RS = READSET(SQLX$, , 1)
        With RS
          Do While Not .EOF
              CANTIOPX% = 0
              CF = 0 '= CANTIDAD DE PARTICIPANTES POR FILA
               AY1% = XVALO(!NUMAYUD1): If AY1% > 0 Then CF = CF + 1: ENVAX(AY1%).NumParticipante = 1: ENVAX(AY1%).NameParticipante = SAFETEXT$(!NAMEAYUD1): ENVAX(AY1%).CantVecesPart = ENVAX(AY1%).CantVecesPart + 1
              AY2% = XVALO(!NUMAYUD2): If AY2% > 0 Then CF = CF + 1: ENVAX(AY2%).NumParticipante = 1: ENVAX(AY2%).NameParticipante = SAFETEXT$(!NAMEAYUD2): ENVAX(AY2%).CantVecesPart = ENVAX(AY2%).CantVecesPart + 1
              AY3% = XVALO(!NUMAYUD3): If AY3% > 0 Then CF = CF + 1: ENVAX(AY3%).NumParticipante = 1: ENVAX(AY3%).NameParticipante = SAFETEXT$(!NAMEAYUD3): ENVAX(AY3%).CantVecesPart = ENVAX(AY3%).CantVecesPart + 1
              AY4% = XVALO(!NUMAYUD4): If AY4% > 0 Then CF = CF + 1: ENVAX(AY4%).NumParticipante = 1: ENVAX(AY4%).NameParticipante = SAFETEXT$(!NAMEAYUD4): ENVAX(AY4%).CantVecesPart = ENVAX(AY4%).CantVecesPart + 1
              AY5% = XVALO(!NUMAYUD5): If AY5% > 0 Then CF = CF + 1: ENVAX(AY5%).NumParticipante = 1: ENVAX(AY5%).NameParticipante = SAFETEXT$(!NAMEAYUD5): ENVAX(AY5%).CantVecesPart = ENVAX(AY5%).CantVecesPart + 1
              AY6% = XVALO(!NUMAYUD6): If AY6% > 0 Then CF = CF + 1: ENVAX(AY6%).NumParticipante = 1: ENVAX(AY6%).NameParticipante = SAFETEXT$(!NAMEAYUD6): ENVAX(AY6%).CantVecesPart = ENVAX(AY6%).CantVecesPart + 1
              AY7% = XVALO(!NUMAYUD7): If AY7% > 0 Then CF = CF + 1: ENVAX(AY7%).NumParticipante = 1: ENVAX(AY7%).NameParticipante = SAFETEXT$(!NAMEAYUD7): ENVAX(AY7%).CantVecesPart = ENVAX(AY7%).CantVecesPart + 1
              AY8% = XVALO(!NUMAYUD8): If AY8% > 0 Then CF = CF + 1: ENVAX(AY8%).NumParticipante = 1: ENVAX(AY8%).NameParticipante = SAFETEXT$(!NAMEAYUD8): ENVAX(AY8%).CantVecesPart = ENVAX(AY8%).CantVecesPart + 1
              IDO1% = XVALO(!NULIDOP1): If IDO1% > 0 Then CF = CF + 1: ENVAX(IDO1%).NumParticipante = 1: ENVAX(IDO1%).NameParticipante = SAFETEXT$(!NULNAMEOP1): ENVAX(IDO1%).CantVecesPart = ENVAX(IDO1%).CantVecesPart + 1
              IDO2% = XVALO(!NULIDOP2): If IDO2% > 0 Then CF = CF + 1: ENVAX(IDO2%).NumParticipante = 1: ENVAX(IDO2%).NameParticipante = SAFETEXT$(!NULNAMEOP2): ENVAX(IDO2%).CantVecesPart = ENVAX(IDO2%).CantVecesPart + 1
              If EXCLUYEMAQ% < 1 Then
                 MAQX% = XVALO(!NUMUSUMAQ):
                 If MAQX% > 0 Then
                    CF = CF + 1
                    ENVAX(MAQX%).NumParticipante = 1: ENVAX(MAQX%).NameParticipante = SAFETEXT$(!NAMEUSUMAQ): ENVAX(MAQX%).CantVecesPart = ENVAX(MAQX%).CantVecesPart + 1
                 End If
'              CantTotHoras_X_Particip
              End If
             HORAX = XVALO(!H_TRABAJADAS) * CF 'CANTIDAD DE PARTICIPANTES X FILA LEIDA POR LAS HORAS TRABAJADAS
             TOTHORAX = TOTHORAX + HORAX

              .MoveNext
          Loop
        End With
        On Error Resume Next
        For i% = 1 To UBound(ENVAX)
          If Val(ENVAX(i%).NumParticipante) > 0 Then
            CantTotParticipantes = CantTotParticipantes + 1
          End If
        Next i%
        On Error GoTo 0
        env.CantTotHoras_X_Particip = TOTHORAX 'DEVUELVE EL TOTAL DE HORAS MULTIPLICADAS POR LA CANTIDAD DE PARTICIPANTES
                                                'FILA A FILA DE UNA ORDEN DE FABRICACION
        RS.Close: Set RS = Nothing
ERROR_CONTROL:
    If Err <> 0 Then
       On Error GoTo 0
    End If


End Sub
Public Sub HorasMenMaySectorXOfyFecha(NORFA As String, Optional INDISOLSUPX%, Optional FE%, Optional MAQ$)

        
        SQLX$ = "SELECT  P.H_DESDE,P.H_HASTA,P.FECHA,P.IDSUBOR,P.INDISOLSUP,P.NUMMAQUI  FROM "
        SQLX$ = SQLX$ + "(SELECT H_DESDE,H_HASTA, FECHA, IDSUBOR,INDISOLSUP,NUMMAQUI  FROM PDIARPROD1 "
        SQLX$ = SQLX$ + " union all "
        SQLX$ = SQLX$ + " SELECT H_DESDE,H_HASTA, FECHA, IDSUBOR,INDISOLSUP,NUMMAQUI  FROM PDIARPROD2) "
        SQLX$ = SQLX$ + " AS P  "
        SQLX$ = SQLX$ + " WHERE  P.IDSUBOR = '" & NORFA & "' "
        If INDISOLSUPX >= 0 Then SQLX$ = SQLX$ + " AND  P.INDISOLSUP = " & INDISOLSUPX%
        If FE% >= 0 Then SQLX$ = SQLX$ + " AND  P.FECHA = '" & FETEXCOR1$(FE%) & "'"
        If TRIM$(MAQ$) <> "" Then SQLX$ = SQLX$ + " AND NUMMAQUI = '" & MAQ$ & "'"
        SQLX$ = SQLX$ + " ORDER BY IDSUBOR"
        Set RS = READSET(SQLX$, , 1)
        i& = 0: HoraMenor1 = 0: HoraMayor1 = 0
        With RS
          Do While Not .EOF
            If XVALO(!H_DESDE) > 0 And XVALO(!H_HASTA) > 0 Then
               i& = i& + 1
               HORADESLEIDA = HORANUM(!H_DESDE)
               HORAHASLEIDA = HORANUM(!H_HASTA)
               If i& = 1 Then HoraMenor1 = HORADESLEIDA
               If i& = 1 Then HoraMayor1 = HORAHASLEIDA
               If HORADESLEIDA < HoraMenor1 Then HoraMenor1 = HORADESLEIDA
               If HORAHASLEIDA > HoraMayor1 Then HoraMayor1 = HORAHASLEIDA
            End If
            .MoveNext
          Loop
        End With
        RS.Close: Set RS = Nothing
        HoraMenor = HoraMenor1
        HoraMayor = HoraMayor1
End Sub
Public Function HORASHOMBRE(INDISOLSUX%, Optional FE%, Optional NORFA$, Optional MAQUINA$) As Double
   'DEVUELVE LAS HORAS HOMBRE QUE SON LAS QUE SE INFORMAN EN CADA SOLAPA EN BOTON DOTACION.
   'EN CASO QUE NO QUERER QUE FILTRE POR INDICE DE SOLAPA PASAR VALOR -1
   HORASHOMBRE = 0
   SQLX$ = "SELECT SUM(H_TRABAJADAS) AS CANTIHORAS FROM DOTACION WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE DOTACION_ID > 0 AND MARBORRA = 0 AND  H_TRABAJADAS > 0"
   If FE% > 0 Then SQLX$ = SQLX$ + " AND FECHA  = '" & FETEXCOR1$(FE%) & "' "
   If INDISOLSUX% >= 0 Then SQLX$ = SQLX$ + "  AND INDISOLSUP = " & INDISOLSUX%
   If TRIM$(NORFA$) <> "" Then SQLX$ = SQLX$ + " AND  IDSUBOR = '" & NORFA$ & "'"
   If TRIM$(MAQUINA$) <> "" Then SQLX$ = SQLX$ + " AND  NUMMAQUI = '" & MAQUINA$ & "'"
   Set RS = READSET(SQLX$)
   With RS
     If Not .EOF Then
       HORASHOMBRE = XVALO(!CANTIHORAS)
     End If
   End With
   RS.Close
   Set RS = Nothing
End Function
Public Function PARTICIP_SEGUN_DOTACION(INDISOLSUX%, Optional FE%, Optional NORFA$, Optional MAQUINA$) As Integer
   PARTICIP_SEGUN_DOTACION = 0 'DOTACION CARGADA EN BOTON DOTACION
   'DEVUELVE LA CANTIDAD DE PARTICIPANTES DE UN SECTOR SEGUN SOLAPA Y FECHA Y NORFA Y MAQUINA
   'EN CASO QUE NO QUERER QUE FILTRE POR INDICE DE SOLAPA PASAR VALOR -1
   SQLX$ = " SELECT COUNT(*) AS CANTI FROM (SELECT  COUNT(*) AS CANTI,NUSER  FROM DOTACION WITH(NOLOCK)"
   SQLX$ = SQLX$ + " Where DOTACION_ID > 0 And MARBORRA = 0 "
   If FE% > 0 Then SQLX$ = SQLX$ + " AND FECHA  = '" & FETEXCOR1$(FE%) & "' "
   If INDISOLSUX% >= 0 Then SQLX$ = SQLX$ + " AND  H_TRABAJADAS > 0 AND INDISOLSUP = " & INDISOLSUX%
   If TRIM$(NORFA$) <> "" Then SQLX$ = SQLX$ + " AND  IDSUBOR = '" & NORFA$ & "'"
   If TRIM$(MAQUINA$) <> "" Then SQLX$ = SQLX$ + " AND  NUMMAQUI  = '" & MAQUINA$ & "'"
   SQLX$ = SQLX$ + " GROUP BY NUSER) AS S"
   
   Set RS = READSET(SQLX$)
   With RS
     If Not .EOF Then
       PARTICIP_SEGUN_DOTACION = XVALO(!CANTI)
     End If
   End With
   RS.Close
   Set RS = Nothing
End Function
Public Function NumeUsSegunNombre%(NAME1$)
'   SELECT NUSER FROM USUARIOSSQL WHERE NOMBRE = 'SERGIO C.'
   CONDI$ = "NOMBRE = '" & NAME1$ & "'"
   NumeUsSegunNombre% = XVALO(VALOBADA("USUARIOSSQL", "NUSER", CONDI$, "NOMESS"))
End Function
Public Function EXISTEREMITO_EN_WALDBOTT%(REMILAR$)
   REMIX$ = FORMATO_RTO_WALDB(REMILAR$)
   CONDI$ = "TIPO = 'R' AND NROPED = '" & REMIX$ & "'"
   SQLX$ = "SELECT COUNT(*) AS CANTI FROM PDCABEZA WHERE " & CONDI$
   Call abreBaDaWald
   Dim RS As New ADODB.Recordset
   Set RS = BASEWALD.Execute(SQLX$)
   EXISTEREMITO_EN_WALDBOTT% = XVALO(RS!CANTI)
   
End Function
Public Function EXISTE_CODIGO_WALDBOTT%(CI1%)

   CONDI$ = "NUMERO = '" & CODEXT$(CI1%) & "'"
   SQLX$ = "SELECT COUNT(*) AS CANTI FROM ARTICULO WHERE " & CONDI$
   Call abreBaDaWald
   Dim RS As New ADODB.Recordset
   Set RS = BASEWALD.Execute(SQLX$)
   EXISTE_CODIGO_WALDBOTT% = XVALO(RS!CANTI)
End Function
Public Function EXISTE_CODIGO_EXTERNO_WALDBOTT%(CODI$)

   CONDI$ = "NUMERO = '" & CODI$ & "'"
   SQLX$ = "SELECT COUNT(*) AS CANTI FROM ARTICULO WHERE " & CONDI$
   Call abreBaDaWald
   Dim RS As New ADODB.Recordset
   Set RS = BASEWALD.Execute(SQLX$)
   EXISTE_CODIGO_EXTERNO_WALDBOTT% = XVALO(RS!CANTI)
End Function

Public Function REMITO_YA_FACTU_EN_WALDBOTT%(REMILAR$)
   REMIX$ = FORMATO_RTO_WALDB(REMILAR$)
   CONDI$ = "TIPO = 'R' AND NROPED = '" & REMIX$ & "' AND ESTADOSEG = 'F' AND FACTURA <> '' "
   SQLX$ = "SELECT COUNT(*) AS CANTI FROM PDCABEZA WHERE " & CONDI$
   Call abreBaDaWald
   Dim RS As New ADODB.Recordset
   Set RS = BASEWALD.Execute(SQLX$)
   REMITO_YA_FACTU_EN_WALDBOTT% = XVALO(RS!CANTI)
   
End Function
Public Function REMITO_EXISTE_EN_WALDBOTT%(REMILAR$)
   'DEBE VENIR CON EL FORMATO DE FLEXOFT RT-X-0001-00000020
   REMIX$ = FORMATO_RTO_WALDB(REMILAR$)
   CONDI$ = "TIPO = 'R' AND NROPED = '" & REMIX$ & "' "
   SQLX$ = "SELECT COUNT(*) AS CANTI FROM PDCABEZA WHERE " & CONDI$
   Call abreBaDaWald
   Dim RS As New ADODB.Recordset
   Set RS = BASEWALD.Execute(SQLX$)
   REMITO_EXISTE_EN_WALDBOTT% = XVALO(RS!CANTI)
   
End Function


Function VALIDAR_SOLAPAMIENTO_DOTACION%(NUSER%, FE%, DESDX%, HASTX%, MODO$, Optional ID&)
    'VALIDA SOLAPAMIENTO CONTRA LA BASE DE DATOS , TENIENDO EN CUENTA EL USUARIO Y QUE SEA DIFERENTE ID.
    VALIDAR_SOLAPAMIENTO_DOTACION% = 0
    SQLX$ = "SELECT  H_DESDE,H_HASTA,NUSER,NOMBRE,INDISOLSUP,SOLAPASUP,IDSUBOR,DOTACION_ID"
    SQLX$ = SQLX$ + " FROM DOTACION "
    SQLX$ = SQLX$ + " WHERE (NUSER =" & NUSER%
    SQLX$ = SQLX$ + " ) AND FECHA = '" & FETEXCOR1$(FE%) & "'"
    '
    Set RS = READSET(SQLX$, , 1)
    With RS
      Do While Not .EOF
         k& = k& + 1
         DESLEIDO% = HORANUM(SAFETEXT$(!H_DESDE))
         HASLEIDO% = HORANUM(SAFETEXT$(!H_HASTA))
         IDLEIDO& = XVALO(!DOTACION_ID)
         SECTOR$ = SAFETEXT$(!SOLAPASUP)
         If ID& <> IDLEIDO& Then
           If HAY_SOLAPAMIENTO_HORARIO%(DESDX%, HASTX%, DESLEIDO%, HASLEIDO%) > 0 Then    ' ESTA FUNCION VALIDA SI EXISTE SOLAPAMIENTO HORARIO
            If InStr(UCase$(MODO$), "NOMESS") < 1 Then
               Call COMUNI("Usuario N° " & NUSER% & "\Solapamiento Horario En Sector : " & SECTOR$ & "\Horario de: " & SAFETEXT$(!H_DESDE) & " a " & SAFETEXT$(!H_HASTA))
               VALIDAR_SOLAPAMIENTO_DOTACION% = 1
               Exit Function
            End If
           End If
         End If
         .MoveNext
       Loop
    End With
    RS.Close
    Set RS = Nothing
         
          
End Function

'Public Sub ALTA_CLIENTE_WALDBOT(NC1&, Optional FUERZA_ACTUALIZACION%)
'    'SI EXISTE LO ACTUALIZA , SI NO LO DA DE ALTA
''    Exit Sub
'    'VALIDO SI ESTA DE BAJA AVISO Y SE VA
'    BDX$ = TRIM$(Control("CONWALDB", "RUTAWALD"))
'    If BDX$ = "" Then Exit Sub
'    If CantRegistros("DEUDOR12", "Nume_Cli = " & NC1& & " AND STAT_CLI < 0") > 0 Then
'       Call COMUNI("Atención !!! Cliente Seleccionado Dado de Baja\Debe Restaurar el Mismo desde el Maestro de Clientes")
'       Exit Sub
'    End If
'    '
'    'VALIDO SI EXISTE, PARA GENERAR UNA BANDERA SI ACTUALIZA O AGREGA.
'    EXISTECLIENTE_EN_FLEXOFT% = 0
'    If CantRegistros("DEUDOR12", "Nume_Cli = " & NC1&) > 0 Then
'       EXISTECLIENTE_EN_FLEXOFT% = 1
'    End If
'    '
'
'    'ABRO LA TABLA PARA PARA TRAER LA INFO DE WALDBOT Y ACTUALIZAR O AGREGAR EL CLIENTE SELECCIONADO (NC1&)
'    Call abreBaDaWald
'    strSQL = "SELECT * FROM CLIENTES"
'    strSQL = strSQL + " WHERE NUMERO  = '" & FORMATOCON0$(NC1&, 7) & "'"
''    Set RS = BaDaWald.OpenRecordset(strSQL, dbOpenSnapshot)
'    Dim RS As New ADODB.Recordset
'    Set RS = BASEWALD.Execute(strSQL)
'
'    With RS
'     If Not .EOF Then
'        If EXISTECLIENTE_EN_FLEXOFT% > 0 And FUERZA_ACTUALIZACION% > 0 Then
'           SQLX$ = "UPDATE DEUDOR12 SET "
'           SQLX$ = SQLX$ + "Raso_Cli= '" & Left$(SAFETEXT$(!RAZONSOC), 64) & "'"
'           SQLX$ = SQLX$ + ", Domi_Cli= '" & Left$(SAFETEXT$(!DIRECCION), 64) & "'"
'           SQLX$ = SQLX$ + ", Pais_Cli= '" & Left$(SAFETEXT$(!PAIS), 48) & "'"
'           SQLX$ = SQLX$ + ", Mail_Cli= '" & Left$(SAFETEXT$(!Email), 64) & "'"
'           SQLX$ = SQLX$ + ", Cuit_Cli= '" & Left$(SAFETEXT$(!Cuit), 24) & "'"
'           SQLX$ = SQLX$ + ", Loca_Cli= '" & Left$(SAFETEXT$(!localidad), 48) & "'"
'           SQLX$ = SQLX$ + ", CPOS_CLI= '" & SAFETEXT$(!CodPostal) & "'"
'           SQLX$ = SQLX$ + ", Cpag_Cli=  " & XVALO(!CONDVTA)
'           SQLX$ = SQLX$ + ", Tele_Cli=  '" & Left$(SAFETEXT$(!TELEFONOS), 32) & "'"
'           SQLX$ = SQLX$ + ", VEND_Cli=  " & XVALO(!VENDEDOR)
'           SQLX$ = SQLX$ + ", CODI_Cli=  '" & SAFETEXT$(XVALO(NC1&)) & "'"
'             IVAWALB$ = SAFETEXT$(!IMPUESTO)
'             PIVXFLEX% = PosIvaFlexVSWaldbott%(IVAWALB$)
'           SQLX$ = SQLX$ + ", Piva_Cli=  " & XVALO(PIVXFLEX%)
'
'
'           SQLX$ = SQLX$ + " WHERE Nume_Cli = " & NC1&
'           'FALA LA POSICIION DE IVA
'        Else
'           SQLX$ = "INSERT INTO DEUDOR12  "
'           SQLX$ = SQLX$ + "(Nume_Cli ,Raso_Cli,Domi_Cli,Pais_Cli,Mail_Cli,Cuit_Cli,Loca_Cli,CPOS_CLI,Cpag_Cli,Tele_Cli,Stat_Cli,VEND_CLI,CODI_Cli,Piva_Cli)"
'           SQLX$ = SQLX$ + " VALUES(" & NC1&
'           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!RAZONSOC), 64) & "'"
'           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!DIRECCION), 64) & "'"
'           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!PAIS), 48) & "'"
'           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!Email), 64) & "'"
'           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!Cuit), 24) & "'"
'           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!localidad), 48) & "'"
'           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!CodPostal), 16) & "'"
'           SQLX$ = SQLX$ + "," & XVALO(!CONDVTA)
'           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!TELEFONOS), 32) & "'"
'           SQLX$ = SQLX$ + ",1" ' STAT_CLI
'           SQLX$ = SQLX$ + "," & SAFETEXT$(XVALO(!VENDEDOR))
'           SQLX$ = SQLX$ + ",'" & XVALO(NC1&) & "'"
'             IVAWALB$ = SAFETEXT$(!IMPUESTO)
'             PIVXFLEX% = PosIvaFlexVSWaldbott%(IVAWALB$)
'           SQLX$ = SQLX$ + "," & XVALO(PIVXFLEX%)
'
'           SQLX$ = SQLX$ + ")"
'        End If
'        FLEXCONN.Execute (SQLX$)
'     Else
'        Call COMUNI("No Se Ha Encontrado En Tabla de Clientes Waldbott el Nro. de Cliente: " & NC1&)
'     End If
'    End With
'    RS.Close
'    Set RS = Nothing
'End Sub
Public Sub ALTA_CLIENTE_WALDBOT(NC1&, Optional FUERZA_ACTUALIZACION%, Optional OK%)
    'SI EXISTE LO ACTUALIZA , SI NO LO DA DE ALTA
'    Exit Sub
    'VALIDO SI ESTA DE BAJA AVISO Y SE VA
    BDX$ = TRIM$(Control("CONWALDB", "RUTAWALD"))
    If BDX$ = "" Then Exit Sub
    If CantRegistros("DEUDOR12", "Nume_Cli = " & NC1& & " AND STAT_CLI < 0") > 0 Then
       Call COMUNI("Atención !!! Cliente Seleccionado Dado de Baja\Debe Restaurar el Mismo desde el Maestro de Clientes")
       Exit Sub
    End If
    '
    'VALIDO SI EXISTE, PARA GENERAR UNA BANDERA SI ACTUALIZA O AGREGA.
    EXISTECLIENTE_EN_FLEXOFT% = 0
    If CantRegistros("DEUDOR12", "Nume_Cli = " & NC1&) > 0 Then
       EXISTECLIENTE_EN_FLEXOFT% = 1
    End If
    '
    
    'ABRO LA TABLA PARA PARA TRAER LA INFO DE WALDBOT Y ACTUALIZAR O AGREGAR EL CLIENTE SELECCIONADO (NC1&)
    Call abreBaDaWald
    strSQL = "SELECT * FROM CLIENTES"
    strSQL = strSQL + " WHERE NUMERO  = '" & FORMATOCON0$(NC1&, 7) & "'"
'    Set RS = BaDaWald.OpenRecordset(strSQL, dbOpenSnapshot)
    Dim RS As New ADODB.Recordset
    Set RS = BASEWALD.Execute(strSQL)

    With RS
     If Not .EOF Then
        If EXISTECLIENTE_EN_FLEXOFT% > 0 And FUERZA_ACTUALIZACION% > 0 Then
           SQLX$ = "UPDATE DEUDOR12 SET "
           SQLX$ = SQLX$ + "Raso_Cli= '" & Left$(SAFETEXT$(!RAZONSOC), 64) & "'"
           SQLX$ = SQLX$ + ", Domi_Cli= '" & Left$(SAFETEXT$(!DIRECCION), 64) & "'"
           SQLX$ = SQLX$ + ", Pais_Cli= '" & Left$(SAFETEXT$(!PAIS), 48) & "'"
           SQLX$ = SQLX$ + ", Mail_Cli= '" & Left$(SAFETEXT$(!Email), 64) & "'"
           SQLX$ = SQLX$ + ", Cuit_Cli= '" & Left$(SAFETEXT$(!Cuit), 24) & "'"
           SQLX$ = SQLX$ + ", Loca_Cli= '" & Left$(SAFETEXT$(!localidad), 48) & "'"
           SQLX$ = SQLX$ + ", CPOS_CLI= '" & SAFETEXT$(!CodPostal) & "'"
           SQLX$ = SQLX$ + ", Cpag_Cli=  " & XVALO(!CONDVTA)
           SQLX$ = SQLX$ + ", Tele_Cli=  '" & Left$(SAFETEXT$(!TELEFONOS), 32) & "'"
           SQLX$ = SQLX$ + ", VEND_Cli=  " & XVALO(!VENDEDOR)
           SQLX$ = SQLX$ + ", CODI_Cli=  '" & SAFETEXT$(XVALO(NC1&)) & "'"

             IVAWALB$ = SAFETEXT$(!IMPUESTO)
             PIVXFLEX% = PosIvaFlexVSWaldbott%(IVAWALB$)
           SQLX$ = SQLX$ + ", Piva_Cli=  " & XVALO(PIVXFLEX%)

           
           SQLX$ = SQLX$ + " WHERE Nume_Cli = " & NC1&
           'FALA LA POSICIION DE IVA
        Else
           SQLX$ = "INSERT INTO DEUDOR12  "
           SQLX$ = SQLX$ + "(Nume_Cli ,Raso_Cli,Domi_Cli,Pais_Cli,Mail_Cli,Cuit_Cli,Loca_Cli,CPOS_CLI,Cpag_Cli,Tele_Cli,Stat_Cli,VEND_CLI,CODI_Cli,Piva_Cli)"
           SQLX$ = SQLX$ + " VALUES(" & NC1&
           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!RAZONSOC), 64) & "'"
           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!DIRECCION), 64) & "'"
           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!PAIS), 48) & "'"
           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!Email), 64) & "'"
           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!Cuit), 24) & "'"
           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!localidad), 48) & "'"
           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!CodPostal), 16) & "'"
           SQLX$ = SQLX$ + "," & XVALO(!CONDVTA)
           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!TELEFONOS), 32) & "'"
           SQLX$ = SQLX$ + ",1" ' STAT_CLI
           SQLX$ = SQLX$ + "," & XVALO(!VENDEDOR)
           SQLX$ = SQLX$ + ",'" & XVALO(NC1&) & "'"
             IVAWALB$ = SAFETEXT$(!IMPUESTO)
             PIVXFLEX% = PosIvaFlexVSWaldbott%(IVAWALB$)
           SQLX$ = SQLX$ + "," & XVALO(PIVXFLEX%)
           
           SQLX$ = SQLX$ + ")"
           'FALA LA POSICIION DE IVA
        End If
        FLEXCONN.Execute (SQLX$)
        OK% = 1
     Else
        Call COMUNI("No Se Ha Encontrado En Tabla de Clientes Waldbott el Nro. de Cliente: " & NC1&)
        OK% = -1
     End If
    End With
    RS.Close
    Set RS = Nothing
End Sub
Function PosIvaFlexVSWaldbott%(POSIVAWALDB$)
      PIVXFLEX% = 0
      PIVX$ = TRIM$(UCase$(POSIVAWALDB$))
      If PIVX$ <> "" Then
         Select Case PIVX$
            Case "01"
               PIVXFLEX% = 1
            Case "02"
               PIVXFLEX% = 2
            Case "03"
               PIVXFLEX% = 4
            Case "04"
               PIVXFLEX% = 2
            Case "05"
               PIVXFLEX% = 5
            Case "06"
               PIVXFLEX% = 0
            Case "07"
               PIVXFLEX% = 3
            Case "MP"
               PIVXFLEX% = 6
            Case "EP"
               PIVXFLEX% = 5
            Case "EX"
               PIVXFLEX% = 4
         End Select
      End If
      PosIvaFlexVSWaldbott% = PIVXFLEX%
End Function
Sub ARTICULO_REMITO_WALDBOTT()
    For YX& = 1 To ULTREG&("!BOLREDET")
      YY$ = REGLEIDO$("!BOLREDET", YX&)
      CI1% = CVI(Mid$(YY$, 83, 2))
      CAN# = CVD(Mid$(YY$, 89, 8))
      If CAN# > 0 Then
         If EXISTE_CODIGO_WALDBOTT%(CI1%) < 1 Then
            Call ALTA_ARTICULO_WALDBOTT(CI1%)
         Else
            Call UPDATE_ARTICULO_WALDBOTT(CI1%)
         End If
      End If
    Next YX&
End Sub
Function CODI_FORMATEADO$(CODIGOEXT$)
   If Len(CODIGOEXT$) <= 17 Then
      CODI_FORMATEADO$ = "CT-" & CODIGOEXT$
   End If
   
End Function
Function CANTPRODUCIDAENKILOSXBOLSA(CI1%)
    'CI1% ES EL CODIGO DE LA BOLSA
    'ESTA FUNCION SEGUN EL CODIGO CALCULA EL PESO DE LA BOLSA
    'LOS COMPONENTES QUE SE TOMAN EN CUENTA SON LOS PAPELES DE LAS FAMILIAS PAPEL, POLIETILENO Y LAMINAS POLIETILENO
    XORTE = CORTEESPECIFIC(CI1%)

    CANTPLIEGOS = 0
    COEFICIENTE = 0.00184
    'USO EXCLUSIVO PARA ENVAPLAST (HARCODEADO)

    SQLX$ = "SELECT CODIELEM"
    SQLX$ = SQLX$ + " FROM FORMULAS  WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE CODICONJ = " & CI1%

    Set RS = READSET(SQLX$)
    With RS
       Do While Not .EOF
         CODELM% = XVALO(!codielem)
         strArray = Split(AnchoColorGramaje$(CODELM%), ";")
         TPAPEL = XVALO(strArray(0))
         GRAMAJE = strArray(1): 'If GRAMAJE <= 0.0005 Then GRAMAJE = 1
         MICRONES = strArray(3): 'If MICRONES <= 0.0005 Then MICRONES = 1
         FAMIL = CODFAMIL$(CODELM%)
         CANTPLIEGOS = CANTPLIEGOS + 1
         Select Case FAMIL
            Case 10  'PAPEL
               PESOPLIEGOS = PESOPLIEGOS + (TPAPEL * XORTE * (GRAMAJE))
            Case 40 'POLIETILENO
               PESOLAMINA = PESOLAMINA + (TPAPEL * XORTE * (GRAMAJE) * COEFICIENTE)
            Case 90 'LAMINAS POLIETILENO
               PESOLAMINAPOLIENTILENO = PESOLAMINAPOLIENTILENO + (TPAPEL * XORTE * MICRONES * COEFICIENTE) / 2
         End Select
         .MoveNext
       Loop
    End With
    
    CANTPRODUCIDAENKILOSXBOLSA = PESOPLIEGOS + PESOLAMINA + PESOLAMINAPOLIENTILENO
    
End Function

Function CantOK_Embalaje(NORFA$, Optional Has%)
    'devuelve la cantidad ok en unidades  que paso ya por embalaje de una of considera opcionalmente una fecha hasta.
    Hasta = FETEXCOR1$(Has%)
    SQLX$ = "SELECT  SUM(P.CANTOK) AS CANTI,IDSUBOR,P.INDISOLSUP "
    SQLX$ = SQLX$ + " FROM (SELECT CANTOK,IDSUBOR,INDISOLSUP "
    SQLX$ = SQLX$ + " FROM PDIARPROD1 UNION ALL SELECT CANTIDAD ,IDSUBOR,INDISOLSUP FROM PDIARPROD2 WITH(NOLOCK)) AS P "
    SQLX$ = SQLX$ + " WHERE  IDSUBOR = '" & NORFA$ & "' AND P.INDISOLSUP = 11 "
    If Has% > 0 Then SQLX$ = SQLX$ + " AND FECHA <= '" & Hasta & "'"
    SQLX$ = SQLX$ + " GROUP BY IDSUBOR,P.INDISOLSUP  HAVING SUM(P.CANTOK) > 0 "
    Set RS = READSET(SQLX$)
    
    CantOK_Embalaje = XVALO(RS!CANTI)

End Function
Function NROPROV_ENVAPLAST_WALDOBTT(NC1&)
    'DEVUELVE EL NUMERO DE PROVEEDOR ENVAPLAST PARA EL CLIENTE QUE ESTA GUARDADO EN BASE WALDBOTT
    'ABRO LA TABLA PARA PARA TRAER LA INFO DE WALDBOT Y ACTUALIZAR O AGREGAR EL CLIENTE SELECCIONADO (NC1&)
    If NC1& < 1 Then Exit Function
    NROPROV_ENVAPLAST_WALDOBTT = ""
    Call abreBaDaWald
    strSQL = "SELECT * FROM CLIENTES"
    strSQL = strSQL + " WHERE NUMERO  = '" & FORMATOCON0$(NC1&, 7) & "'"
'    Set RS = BaDaWald.OpenRecordset(strSQL, dbOpenSnapshot)
    Dim RS As New ADODB.Recordset
    Set RS = BASEWALD.Execute(strSQL)

    With RS
     If Not .EOF Then
        NROPROV_ENVAPLAST_WALDOBTT = SAFETEXT$(!CONTACTO)
     End If
    End With
    RS.Close
    Set RS = Nothing


End Function
Public Sub ALTA_CLIENTE_WALDBOTCLIENTE(NC1&, Optional FUERZA_ACTUALIZACION%, Optional OK%)
    'SI EXISTE LO ACTUALIZA , SI NO LO DA DE ALTA
'    Exit Sub
    'VALIDO SI ESTA DE BAJA AVISO Y SE VA
'    BDX$ = TRIM$(Control("CONWALDB", "RUTAWALD"))
    BDX$ = TRIM$(Control("CONWALDB", "RTLECTCL"))

    If BDX$ = "" Then Exit Sub
    If CantRegistros("DEUDOR12", "Nume_Cli = " & NC1& & " AND STAT_CLI < 0") > 0 Then
       Call COMUNI("Atención !!! Cliente Seleccionado Dado de Baja\Debe Restaurar el Mismo desde el Maestro de Clientes")
       Exit Sub
    End If
    '
    'VALIDO SI EXISTE, PARA GENERAR UNA BANDERA SI ACTUALIZA O AGREGA.
    EXISTECLIENTE_EN_FLEXOFT% = 0
    If CantRegistros("DEUDOR12", "Nume_Cli = " & NC1&) > 0 Then
       EXISTECLIENTE_EN_FLEXOFT% = 1
    End If
    '
    
    'ABRO LA TABLA PARA PARA TRAER LA INFO DE WALDBOT Y ACTUALIZAR O AGREGAR EL CLIENTE SELECCIONADO (NC1&)
    Call abreBaDaWaldClientes
    strSQL = "SELECT * FROM CLIENTES"
    strSQL = strSQL + " WHERE NUMERO  = '" & FORMATOCON0$(NC1&, 7) & "'"
'    Set RS = BaDaWald.OpenRecordset(strSQL, dbOpenSnapshot)
    Dim RS As New ADODB.Recordset
    Set RS = BASEWALDCLIENTE.Execute(strSQL)

    With RS
     If Not .EOF Then
        If EXISTECLIENTE_EN_FLEXOFT% > 0 And FUERZA_ACTUALIZACION% > 0 Then
           SQLX$ = "UPDATE DEUDOR12 SET "
           SQLX$ = SQLX$ + "Raso_Cli= '" & Left$(SAFETEXT$(!RAZONSOC), 64) & "'"
           SQLX$ = SQLX$ + ", Domi_Cli= '" & Left$(SAFETEXT$(!DIRECCION), 64) & "'"
           SQLX$ = SQLX$ + ", Pais_Cli= '" & Left$(SAFETEXT$(!PAIS), 48) & "'"
           SQLX$ = SQLX$ + ", Mail_Cli= '" & Left$(SAFETEXT$(!Email), 64) & "'"
           SQLX$ = SQLX$ + ", Cuit_Cli= '" & Left$(SAFETEXT$(!Cuit), 24) & "'"
           SQLX$ = SQLX$ + ", Loca_Cli= '" & Left$(SAFETEXT$(!localidad), 48) & "'"
           SQLX$ = SQLX$ + ", CPOS_CLI= '" & SAFETEXT$(!CodPostal) & "'"
           SQLX$ = SQLX$ + ", Cpag_Cli=  " & XVALO(!CONDVTA)
           SQLX$ = SQLX$ + ", Tele_Cli=  '" & Left$(SAFETEXT$(!TELEFONOS), 32) & "'"
           SQLX$ = SQLX$ + ", VEND_Cli=  " & XVALO(!VENDEDOR)
           SQLX$ = SQLX$ + ", CODI_Cli=  '" & SAFETEXT$(XVALO(NC1&)) & "'"
             IVAWALB$ = SAFETEXT$(!IMPUESTO)
             PIVXFLEX% = PosIvaFlexVSWaldbott%(IVAWALB$)
           SQLX$ = SQLX$ + ", Piva_Cli=  " & XVALO(PIVXFLEX%)

           
           SQLX$ = SQLX$ + " WHERE Nume_Cli = " & NC1&
           'FALA LA POSICIION DE IVA
        Else
           SQLX$ = "INSERT INTO DEUDOR12  "
           SQLX$ = SQLX$ + "(Nume_Cli ,Raso_Cli,Domi_Cli,Pais_Cli,Mail_Cli,Cuit_Cli,Loca_Cli,CPOS_CLI,Cpag_Cli,Tele_Cli,Stat_Cli,VEND_CLI,CODI_Cli,Piva_Cli)"
           SQLX$ = SQLX$ + " VALUES(" & NC1&
           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!RAZONSOC), 64) & "'"
           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!DIRECCION), 64) & "'"
           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!PAIS), 48) & "'"
           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!Email), 64) & "'"
           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!Cuit), 24) & "'"
           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!localidad), 48) & "'"
           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!CodPostal), 16) & "'"
           SQLX$ = SQLX$ + "," & XVALO(!CONDVTA)
           SQLX$ = SQLX$ + ",'" & Left$(SAFETEXT$(!TELEFONOS), 32) & "'"
           SQLX$ = SQLX$ + ",1" ' STAT_CLI
           SQLX$ = SQLX$ + "," & XVALO(!VENDEDOR)
           SQLX$ = SQLX$ + ",'" & XVALO(NC1&) & "'"
             IVAWALB$ = SAFETEXT$(!IMPUESTO)
             PIVXFLEX% = PosIvaFlexVSWaldbott%(IVAWALB$)
           SQLX$ = SQLX$ + "," & XVALO(PIVXFLEX%)
           SQLX$ = SQLX$ + ")"
        End If
        FLEXCONN.Execute (SQLX$)
        OK% = 1
     Else
        Call COMUNI("No Se Ha Encontrado En Tabla de Clientes Waldbott el Nro. de Cliente: " & NC1&)
        OK% = -1
     End If
    End With
    RS.Close
    Set RS = Nothing
    BASEWALDCLIENTE.Close
End Sub

Function CantPliegosSegunConsumo%(NORFA$, INDISOLAPASUP%, INDISOLINF%)
    'DEVUELVE LA CANTIDAD DE REGISTROS QUE LA FAMILIA CORRESPONDA A 10,40,90 AGRUPADO POR CODIGO
    'EJEMPLO SI TIENE 5 REGISTROS 3 CON CODIGO A Y 2 CON CODIGO B VA A DEVOLVER 2
    
    SQLX$ = "SELECT COUNT(*) AS CANTI FROM ("
    SQLX$ = SQLX$ + " SELECT  COUNT(CODINT) CODINT FROM CONSUMO2 WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "' AND INDISOLSUP = " & INDISOLAPASUP%
    SQLX$ = SQLX$ + " AND INDISOLINF = " & INDISOLINF%
    SQLX$ = SQLX$ + " AND (MARBORRA < 1 OR MARBORRA IS NULL)"
    SQLX$ = SQLX$ + " group by codint) AS T "
    SQLX$ = SQLX$ + " INNER JOIN MASTER M ON M.Codint = T.CODINT"
    SQLX$ = SQLX$ + " WHERE M.FAMILIA IN ('10','40','90')"
    
    Set RS = READSET(SQLX$)
    CantPliegosSegunConsumo% = XVALO(RS!CANTI)
    RS.Close
    Set RS = Nothing

'   CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
'   CONDI$ = CONDI$ + " AND INDISOLSUP = " & INDISOLAPASUP%
'   CONDI$ = CONDI$ + " AND INDISOLINF = " & INDISOLINF%
'   CONDI$ = CONDI$ + " AND INDISOLINF = " & INDISOLINF%
'   CONDI$ = CONDI$ + " AND (MARBORRA < 1 OR MARBORRA IS NULL) group by codint"
End Function


