Attribute VB_Name = "CLIE_FUN"
Public ARCHIPRE$

' ACCDEUDO.QBS                         ' accesos varios a archivo de deudores
'
Public ARCHE$(16)
'Public NC As Integer
'Public MODO As Integer
'Public NROPRO&
'Public NCLILLAMA%
'
Public VERICUIT%
'
'Dim PUCLI(1) As String
'Dim PUBAC(1) As String
Public TRANACT%    ' NUMERO DE TRANSPORTE ACTIVO - PARA PODER CAMBIARLO EN EL PEDIDO
Dim ClienTemp As Recordset
'
Function NUMECLI%(COPRO$)
   '
   NPRX% = 0
   CPRX$ = UCase$(TRIM$(COPRO$))
   If CPRX$ <> "" Then
     Call APERBASDAT("CLIEN")
     Set ClienTemp = Badaclie.OpenRecordset("SELECT * FROM Deudor12 WHERE Codi_Cli = '" & CPRX$ & "'", dbOpenSnapshot)
     On Error Resume Next '
     ClienTemp.MoveFirst
     If Err Then GoTo 90          ' no existe
     NPRX% = Abs(ClienTemp!Nume_Cli)
   End If
   '
90 NUMECLI% = NPRX%
   '
End Function
'
Function NUMEPRO%(COPRO$)
   '
   NPRX% = 0
   CPRX$ = UCase$(TRIM$(COPRO$))
   If CPRX$ <> "" Then
     Call APERBASDAT("PROVE")
     Set ClienTemp = Badaprov.OpenRecordset("SELECT * FROM Proved12 WHERE Codi_Cli = '" & CPRX$ & "'", dbOpenSnapshot)
     On Error Resume Next '
     ClienTemp.MoveFirst
     If Err Then GoTo 90          ' no existe
     NPRX% = Abs(ClienTemp!Nume_Cli)
   End If
   '
90 NUMEPRO% = NPRX%
   '
End Function

Function STATCLI%(NCLIE)
   '
   Static NCLIEA&, STATUIA%
   '
   ' = 1 --> Existe y está activo
   ' =-1 --> Existe y esta de baja
   ' = 0 --> No existe el item
   '
   NCLX& = Int(NCLIE)
   If NCLX& = NCLIEA& Then
     If STATUIA% > 0 Then
       STATCLI% = STATUIA%
       Exit Function
     End If
   End If
   '
   STATCLI% = 0
   If NCLX& <> 0 Then
     If NCLX& > 0 Then
         Call APERBASDAT("CLIEN")
         Set ClienTemp = Badaclie.OpenRecordset("SELECT * FROM Deudor12 WHERE Nume_Cli = " & Abs(NCLX&), dbOpenSnapshot)
       Else
         Call APERBASDAT("PROVE")
         Set ClienTemp = Badaprov.OpenRecordset("SELECT * FROM Proved12 WHERE Nume_Cli = " & Abs(NCLX&), dbOpenSnapshot)
     End If
     On Error Resume Next '
     ClienTemp.MoveFirst
     If Err Then
        GoTo 90   ' no existe
     End If
     STATCLI = 1
     If ClienTemp!Stat_Cli < 1 Then
        STATCLI = (-1)
     End If
   End If
   '
90 NCLIEA& = NCLX&
   STATUIA% = STATCLI%
   '
End Function
'
Function REGICLI&(NC)
   RGCL& = 0
   If NC < 0 Then
       Call APERBASDAT("PROVE")
       Proved12.FindFirst "Nume_Cli = " & Abs(Int(NC))
       If Proved12.NoMatch = False Then
          RGCL& = 1
       End If
     ElseIf NC > 0 Then
       Call APERBASDAT("CLIEN")
       Deudor12.FindFirst "Nume_Cli = " & NC
       If Deudor12.NoMatch = False Then
          RGCL& = 1
       End If
   End If
   REGICLI& = RGCL&
End Function
'
Function CODICLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then CODICLI$ = TRIM$(ClienTemp!Codi_Cli)
End Function
'
Function RASOCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then RASOCLI$ = TRIM$(ClienTemp!Raso_Cli)
End Function
'
Function DOMICLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then DOMICLI$ = TRIM$(ClienTemp!Domi_Cli)
End Function
'
Function CPOSCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then CPOSCLI$ = TRIM$(ClienTemp!Cpos_Cli)
End Function
'
Function LOCACLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then LOCACLI$ = TRIM$(ClienTemp!Loca_Cli)
End Function
'
Function CPLOCLI$(NC)
   CPX$ = TRIM$(CPOSCLI$(NC))
   If CPX$ <> "" Then CPX$ = "(" + CPX$ + ") "
   CPLOCLI$ = TRIM$(CPX$ + LOCACLI$(NC))
End Function
'
Function PAISCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then PAISCLI$ = TRIM$(ClienTemp!Pais_Cli)
End Function
'
Function TEL1CLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then TEL1CLI$ = TRIM$(ClienTemp!Tele_Cli)
End Function
'
Function TEL2CLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then TEL2CLI$ = TRIM$(ClienTemp!Fant_Cli)
End Function
'
Function ABRECLI$(NC)
   ABRECLI$ = TEL2CLI$(NC)
End Function
'
Function FANTCLI$(NC)
   FANTCLI$ = TEL2CLI$(NC)
   If FANTCLI$ = "" Then
     FANTCLI$ = UCase$(RASOCLI$(NC))
     PPXX% = InStr(FANTCLI$, " ")
     If PPXX% > 1 Then
       FANTCLI$ = Left$(FANTCLI$, PPXX% - 1)
     End If
   End If
End Function
'
Function FAXCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then FAXCLI$ = TRIM$(ClienTemp!Faxi_Cli)
End Function
'
Function EMAILCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then EMAILCLI$ = TRIM$(ClienTemp!Mail_Cli)
End Function
'
Function PAGWEBCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then PAGWEBCLI$ = TRIM$(ClienTemp!PagWeb)
End Function

Function CONTACTCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then CONTACTCLI$ = TRIM$(ClienTemp!Ctac_Cli)
End Function
'
Function HORARCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then HORARCLI$ = TRIM$(ClienTemp!Hora_Cli)
End Function
'
Function CUITCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then CUITCLI$ = TRIM$(ClienTemp!Cuit_Cli)
End Function
'
Function NIBRCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then NIBRCLI$ = TRIM$(ClienTemp!Nibr_Cli)
End Function
'
Function NPROCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then NPROCLI$ = TRIM$(ClienTemp!Prvd_Cli)
End Function
'
Function PROVCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then PROVCLI$ = TRIM$(ClienTemp!Prov_Cli)
End Function
'
Function CPAGCLI%(NC%)
   On Error Resume Next
   If STATCLI%(NC%) <> 0 Then CPAGCLI% = ClienTemp!Cpag_Cli
   If CPAGCLI < 1 Then
     If NC% > 0 Then
         CPAGCLI% = XVALO(CONTROL$("AMACLI", "CONPAG")) ' PARA QUE TOME POR DEFAULT
       Else
         CPAGCLI% = XVALO(CONTROL$("AMAPRO", "CONPAG")) ' PARA QUE TOME POR DEFAULT
     End If
   End If
End Function
'
Function CONDICOM$(NC%)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then CONDICOM$ = ClienTemp!CCom_Cli
End Function
'
Function LIPRCLI%(NC%)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then LIPRCLI% = ClienTemp!Lpre_Cli
   If LIPRCLI% < 1 Then
     LIPRCLI% = XVALO(CONTROL$("AMACLI", "LISTAS")) ' PARA QUE TOME POR DEFAULT
   End If
End Function
'
Function TICUEPRO%(NC%)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then TICUEPRO% = ClienTemp!Tcpr_Cli
End Function
'
Function VENDCLI%(NC%)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then VENDCLI% = TRIM$(ClienTemp!Vend_Cli)
   If VENDCLI% < 0 Or VENDCLI% > 255 Then
      Call MENSERR(24, "Número de Vendedor no Válido\en Cliente Número " + TRIM$(Str$(NC%)) + ".")
      VENDCLI% = 0
   End If
End Function
'
Function PIVACLI%(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then PIVACLI% = TRIM$(ClienTemp!Piva_Cli)
End Function
'
Function CATIBRU$(NC%)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then CATIBRU$ = TRIM$(ClienTemp!Caib_Cli)
End Function
'
Function LIMICRE(NC%)
   LIMICRE = 0
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then LIMICRE = CDbl(ClienTemp!Lcre_Cli)
End Function
'
Function DESCUCLI$(NC%)
   '
   Static DSCX%
   If DSCX% = 0 Then
      DSCX% = 1
      If EXISTE%(LUDAT$ + "DESCOMBI.RFS") > 0 Then
        If ULTREG&("DESCOMBI") > 0 Then
          DSCX% = (-1)
        End If
      End If
   End If
   '
   DESCUCLI$ = ""
   On Error Resume Next
   If DSCX% > 0 Then
      If STATCLI%(NC%) <> 0 Then
         DESCUCLI$ = TRIM$(ClienTemp!Dsta_Cli)
      End If
   End If
   '
End Function
'
Function SALDCLI#(NCX7%)
   '
   SALDCLI# = 0
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT SUM(IdebeComp) AS SUMADEBE, SUM(IHabeComp) AS SUMAHABE FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE NuClien = " & NCX7% & " "
   SQLX$ = SQLX$ + "AND (TipoMovim = 'FVEN' "
   SQLX$ = SQLX$ + "OR TipoMovim = 'RCOB' "
   SQLX$ = SQLX$ + "OR TipoMovim = 'AJUD') "
   Set SUMADEBHAB = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   On Error Resume Next
   SUMADEBHAB.MoveFirst
   SALDCLI# = SUMADEBHAB!SUMADEBE - SUMADEBHAB!SUMAHABE
   On Error GoTo 0
   '
End Function
'
Function PORDESCLI(NC%)
    '
    PDXC = 0
    DESCX$ = REPLACAR$(DESCUCLI$(NC%), "%", " ")
    '
    If InStr(DESCX$, "+") < 1 Then
        PDXC = XVALO(DESCX$)
      Else
        XDC% = InStr(DESCX$, "+")
        DESC1 = XVALO(Left$(DESCX$, XDC% - 1))
        DESC2 = XVALO(Mid$(DESCX$, XDC% + 1))
        PDXC = 100 * (1 - (1 - DESC1 / 100) * (1 - DESC2 / 100))
    End If
    '
99  PORDESCLI = PDXC
    '
End Function
'
Function LUENTRECLI$(NC%, OPCI%)
   LUENTRECLI$ = ""
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then
     If OPCI% = 0 Then
         LUENTRECLI$ = TRIM$((ClienTemp!Lentre0_Cli))
         If LUENTRECLI$ = "" Then
           LTXX$ = FILTERGETRECORD$("LENTRECL", 1, MKI$(NC%))
           LUENTRECLI$ = TRIM$(Mid$(LTXX$, 3, 32))
         End If
         If LUENTRECLI$ = "" Then LUENTRECLI$ = DOMICLI$(NC%)
       ElseIf OPCI% = 1 Then
         LUENTRECLI$ = (ClienTemp!Lentre1_Cli)
         If LUENTRECLI$ = "" Then
           LTXX$ = FILTERGETRECORD$("LENTRECL", 1, MKI$(NC%))
           LUENTRECLI$ = TRIM$(Mid$(LTXX$, 33, 32))
         End If
         If LUENTRECLI$ = "" Then LUENTRECLI$ = CPOSCLI$(NC%) + " - " + LOCACLI$(NC%)
       ElseIf OPCI% = 2 Then
         NTRX% = NUTRANCLI%(NC%)
         If NTRX% < 1 Then
             LUENTRECLI$ = (ClienTemp!Lentre2_Cli)
             If LUENTRECLI$ = "" Then
               LTXX$ = FILTERGETRECORD$("LENTRECL", 1, MKI$(NC%))
               LUENTRECLI$ = TRIM$(Mid$(LTXX$, 67, 32))
             End If
           Else
             LUENTRECLI$ = ECOARCH$("TRANSPOR", MKI$(NTRX%))
         End If
       ElseIf OPCI% = 3 Then
         LUENTRECLI$ = HORAENT$(NC%)
     End If
   End If
End Function
'
Function DOMIENT$(NC%)
   DOMIENT$ = LUENTRECLI$(NC%, 0)
End Function
'
Function LOCAENT$(NC%)
   LOCAENT$ = LUENTRECLI$(NC%, 1)
End Function
'
Function TRANENT$(NC%)
   TRANENT$ = LUENTRECLI$(NC%, 2)
End Function
'
Function NUTRANCLI%(NC, Optional NumSucursal%)
   If TRANACT% > 0 Then
       NUTRANCLI% = TRANACT%
     Else
       On Error Resume Next
       If STATCLI%(NC) <> 0 Then NUTRANCLI% = TRIM$(ClienTemp!NTrans_Cli)
   End If
   '
   'Si tiene informacion del numero de  sucursal
   If NumSucursal% > 0 Then
      Call FILTERFILE("SUCENTRE", "", 5, MKI$(NC))
      URE1& = ULTREG&("!SUCENTRE")
      For KKU& = 1 To URE1&
       '\\\OT-08-0630-OD-12/08/08///
        TTXX$ = REGLEIDO$("!SUCENTRE", KKU&)
       '\\\OT-08-0630-OD-FIN///
        NuSucur% = CVI(Mid$(TTXX$, 1, 2))
        If NumSucursal% = NuSucur% Then
          NUTRANCLI% = CVI(Mid$(TTXX$, 195, 2))
          Exit Function
        End If
      Next KKU&
   End If
   '
End Function
'
Function DENOTRANS$(NTRANI%)
    XX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRANI%))
    DENOTRANS$ = TRIM$(Mid$(XX$, 3, 32))
End Function

Function RASOTRANS$(NTRANI%)
    XX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRANI%))
    RASOTRANS$ = TRIM$(Mid$(XX$, 35, 48))
End Function

Function DOMITRANS$(NTRANI%)
    XX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRANI%))
    DOMITRANS$ = TRIM$(Mid$(XX$, 83, 48))
End Function

Function LOCATRANS$(NTRANI%)
    XX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRANI%))
    LOCATRANS$ = TRIM$(Mid$(XX$, 131, 48))
End Function

Function TELETRANS$(NTRANI%)
    XX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRANI%))
    TELETRANS$ = TRIM$(Mid$(XX$, 195, 48))
End Function

Function NTRANCLI$(NC) ' NOMBRE DEL TRANSPORTISTA
    NTRANCLI$ = ""
    If NC > 0 Then
      NTR% = NUTRANCLI%(NC)
      If NTR% > 0 Then
        NTRANCLI$ = DENOTRANS$(NTR%)
      End If
    End If
End Function

Function RTRANCLI$(NC)           ' R.SOCIAL DEL TRANSPORTISTA
    RTRANCLI$ = ""
    If NC > 0 Then
      NTR% = NUTRANCLI%(NC)
      If NTR% > 0 Then
        RTRANCLI$ = RASOTRANS$(NTR%)
      End If
    End If
End Function

Function DTRANCLI$(NC)           ' DOMICILIO DEL TRANSPORTISTA
    DTRANCLI$ = ""
    If NC > 0 Then
      NTR% = NUTRANCLI%(NC)
      If NTR% > 0 Then
        DTRANCLI$ = DOMITRANS$(NTR%)
      End If
    End If
End Function

Function LTRANCLI$(NC)           ' LOCALIDAD DEL TRANSPORTISTA
    LTRANCLI$ = ""
    If NC > 0 Then
      NTR% = NUTRANCLI%(NC)
      If NTR% > 0 Then
        LTRANCLI$ = LOCATRANS$(NTR%)
      End If
    End If
End Function

'
Function HORAENT$(NC%)
   On Error Resume Next
   If STATCLI%(NC%) <> 0 Then HORAENT$ = TRIM$(ClienTemp!Hora_Ent)
End Function
'
Function ZONACLI$(NC%)
   On Error Resume Next
   If STATCLI%(NC%) <> 0 Then ZONACLI$ = TRIM$(ClienTemp!Zona_Cli)
End Function
'
Function OBSENTRE$(NC%)
   On Error Resume Next
   If STATCLI%(NC%) <> 0 Then OBSENTRE$ = ClienTemp!Obser_Entre
End Function
'
Function POREMBAL(NC%)
   On Error Resume Next
   If STATCLI%(NC%) <> 0 Then POREMBAL = TRIM$(ClienTemp!Recar_Embal)
End Function
'
Function CHEPEACRE#(NC%)
   '
   SUCHEPE# = 0
   FEXI% = DIANTE(DIANTE(HOY(HOS$)))
   AYER = CDbl(CVDAT(FEXI%))
15 Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
   SQLX$ = SQLX$ + "AND NuClien = " & NC% & " "
   SQLX$ = SQLX$ + "AND Status > 0 "
   SQLX$ = SQLX$ + "AND CDBL(FechAcred) >= " & AYER & " "
   SQLX$ = SQLX$ + "ORDER BY FechAcred ASC;"
   '
   Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         If TRIM$(!OpciMovim) <> "CH" Then GoTo 19 ' no es un cheque
         If IsNull(!FECHENTRE) = False Then
           If CVINT(!FECHENTRE) > 33 Then
             GoTo 19    ' salio de cartera
           End If
         End If
         '
         SUCHEPE# = SUCHEPE# + !ValAbsComp
         '
19     .MoveNext
       Loop
     End If
   End With
   '
   CHEPEACRE# = SUCHEPE#
   '
End Function
'
Function CUECOCLI%(NC%)
   CUECOCLI% = 0
   On Error Resume Next
   If STATCLI%(NC%) <> 0 Then CUECOCLI% = ClienTemp!CuMa_Cli
   '
   If CUECOCLI% < 1 Then
     CUMA$ = CONTROL$("AMACLI", "ECUECON")
     If NC% < 0 Then
       CUMA$ = CONTROL$("AMAPRO", "ECUECON")
     End If
     CUECOCLI% = CUEINT%(CUMA$)
   End If
End Function
'
Function CUEMADRE$(NC%)
   '
   CUEMADRE$ = ""
   On Error Resume Next
   If STATCLI%(NC%) <> 0 Then CUEMADRE$ = CODCUE$(ClienTemp!CuMa_Cli)
   '
   If TRIM$(CUEMADRE$) = "" Then
     CUEMADRE$ = CONTROL$("AMACLI", "ECUECON")
     If NC% < 0 Then
       CUEMADRE$ = CONTROL$("AMAPRO", "ECUECON")
     End If
   End If
   '
End Function
'
Function GRUCOCLI%(NC%)
   GRUCOCLI% = 0
   On Error Resume Next
   If STATCLI%(NC%) <> 0 Then GRUCOCLI% = ClienTemp!GrCo_Cli
End Function
'
Function GRUDIVCLI%(NC%)
   GRUDIVCLI% = NC%
   On Error Resume Next
   If STATCLI%(NC%) <> 0 Then
     GRDVC% = ClienTemp!Gru_Div
     If GRDVC% > 0 Then
       GRUDIVCLI% = GRDVC%
     End If
   End If
End Function
'
Function ORIGRUDI%(NC%)
   Static CTXYX1%, ORGD%(32767)
   If CTXYX1% < 1 Then
      For KKU& = 1 To 32767
         ORGD%(KKU&) = 0
      Next KKU&
      Call APERBASDAT("CLIEN")
      Set BCLIEN = Badaclie.OpenRecordset("SELECT * FROM Deudor12 WHERE Gru_Div > 0", dbOpenSnapshot)
      On Error Resume Next
      With BCLIEN
        .MoveFirst
        If Err < 1 Then
           On Error GoTo 0
           Do While Not .EOF
             GRD% = !Nume_Cli
             ORGD%(GRD%) = !Gru_Div
           .MoveNext
           Loop
        End If
        On Error GoTo 0
      End With
      CTXYX1% = 1
   End If
   '
   ORIGRUDI% = ORGD%(NC%)
   '
End Function
'
Function ORDECHEQ$(NC%)
   ORDECHEQ$ = RASOCLI$(NC%)
   If NC% < 0 Then
     On Error Resume Next
     If STATCLI%(NC%) <> 0 Then
       If IsNull(ClienTemp!Ord_Cheq) = False Then
         ORDECHEQ$ = ClienTemp!Ord_Cheq
       End If
     End If
   End If
End Function
'
Function ANOTACLI$(NC%)
   ANOTACLI$ = ""
   On Error Resume Next
   If STATCLI%(NC%) <> 0 Then ANOTACLI$ = ClienTemp!Nota_Cli
End Function
'
Sub NOTACLI(NC%, MODO%)
   '
   ' MODO% = 1    abre automaticamente si hay nota para el cliente
   ' MODO% = 2    se pidió apertura por botón de acceso
   '
   If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then Exit Sub
   '
   RNC$ = RECFILT$("NOTACLI", 1, MKI$(NC%))
   If Len(RNC$) > 0 Or MODO% = 2 Then
      Screen.MousePointer = 11
      JJ& = 0
      For U& = 1 To Len(RNC$) Step 4
         RENOTA& = CVS(Mid$(RNC$, U&, 4))
         If RENOTA& > 0 Then
            If RENOTA& <= ULTREG&("NOTACLI") Then
               X$ = REGLEIDO$("NOTACLI", RENOTA&)
               JJ& = JJ& + 1
               Call GRAREG("!NOTACLI", X$, JJ&)
            End If
         End If
      Next U&
      Call SETULTREG("!NOTACLI", JJ&)
      Call CIERRARCH("!NOTACLI")
      Screen.MousePointer = 1
      VTB% = VENTABU%("NOTACLI")
      If VTB% >= 0 Then
         Screen.MousePointer = 11
         URN& = ULTREG&("!NOTACLI")
         JJ& = ULTREG&("NOTACLI")
         While Len(RNC$) < 4 * URN&
            JJ& = JJ& + 1
            RNC$ = RNC$ + MKS$(JJ&)
         Wend
         KK& = (-3)
         For JJ& = 1 To ULTREG&("!NOTACLI")
            X$ = REGLEIDO$("!NOTACLI", JJ&)
            Call REPLA(X$, MKI$(NC%), 1, 2)
            KK& = KK& + 4
            RENOTA& = CVS(Mid$(RNC$, KK&, 4))
            Call GRAREG("NOTACLI", X$, RENOTA&)
         Next JJ&
         While KK& < Len(RNC$) - 3
            KK& = KK& + 4
            RENOTA& = CVS(Mid$(RNC$, KK&, 4))
            Call GRAREG("NOTACLI", MKI$(NC%) + Space$(78), RENOTA&)
         Wend
         Call CIERRARCH("NOTACLI")
         Screen.MousePointer = 1
      End If
   End If
End Sub
'
Function DEPROVI$(PR$)
   DEPROVI$ = ECOARCH$("PROVINC", PR$)
End Function
'
Function DELIPRE$(CL%)
   DX$ = "PRECIOS LIBRES"
   If CL% > 0 Then If CL% < 256 Then DX$ = ECOARCH$("LISTAS", Chr$(CL%))
   If DX$ = "" Then Beep: CL% = 0: DX$ = "PRECIOS LIBRES"
   DELIPRE$ = DX$
End Function
'
Function DECONPA$(CL%)
   DX$ = "CONTADO"
   If CL% > 0 Then If CL% < 256 Then DX$ = ECOARCH$("CONPAG", Chr$(CL%))
   If DX$ = "" Then Beep: CL% = 0: DX$ = "CONTADO"
   DECONPA$ = DX$
End Function
'
Function DEVENDE$(CL%)
   DX$ = "GERENCIA"
   If CL% > 0 Then DX$ = ECOARCH$("VENDEDOR", Chr$(CL%))
   If DX$ = "" Then Beep: CL% = 0: DX$ = "GERENCIA"
   DEVENDE$ = DX$
End Function
'
Function DEGRUCLI$(CL%)
   DX$ = "NO CLASIFICADO"
   If CL% > 0 Then DX$ = ECOARCH$("GRUCOCLI", MKI$(CL%))
   If DX$ = "" Then Beep: CL% = 0: DX$ = "NO CLASIFICADO"
   DEGRUCLI$ = DX$
End Function
'
Function POSIVAC$(NC%)
    '
    CAPOS% = 7
    Dim PIVA$(7)
    '
    PIVA$(0) = "CONS.FIN."
    PIVA$(1) = "RESP.INSC"
    PIVA$(2) = "R.NO INSC"
    PIVA$(3) = "NO RESP."
    PIVA$(4) = "EXENTO"
    PIVA$(5) = "COM.EXT. "
    PIVA$(6) = "MONOTR."
    PIVA$(7) = "EXEN-PRO"
    '
    APV% = PIVACLI%(NC%)
    PS$ = ""
    If APV% >= 0 Then
        If APV% <= CAPOS% Then
            PS$ = PIVA$(APV%)
        End If
    End If
    '
    POSIVAC$ = PS$
End Function
'
Function POSIVAL$(NC)
    '
    CAPOS% = 7
    Dim PIVA$(7)
    '
    PIVA$(0) = "A CONSUMIDOR FINAL"
    PIVA$(1) = "IVA RESPONSABLE INSCRIPTO"
    PIVA$(2) = "IVA RESPONSABLE NO INSCRIPTO"
    PIVA$(3) = "IVA NO RESPONSABLE"
    PIVA$(4) = "IVA EXENTO"
    PIVA$(5) = "OPs.COMERCIO EXTERIOR"
    PIVA$(6) = "CONTRIBUYENTE MONOTRIBUTO"
    PIVA$(7) = "EXENTO - ZONA PROMOVIDA"
    '
    If InStr(EMPREAC$, "SABO") > 0 Then
      PIVA$(5) = "OPERACIONES EXPORTACION"
    End If
    '
    APV% = PIVACLI%(NC)
    PS$ = ""
    If APV% >= 0 Then
        If APV% <= CAPOS% Then
            PS$ = PIVA$(APV%)
        End If
    End If
    '
POSIVAL$ = PS$
End Function
'
Function VALICUIT%(CUIT$)
    '
    CU$ = TRIM$(CUIT$)
    VC% = 0
    If CONTROL$("FORMCUIT", "SUPRIVAL") = "SI" Then GoTo 900
    '
    If Len(CU$) < 11 Then GoTo 999                                    ' no valida
    '
    VX% = Asc(Mid$(CU$, 3))
    If VX% >= 48 Then
        If VX% <= 57 Then
            CU$ = Left$(CU$, 2) + "-" + Mid$(CU$, 3)
        End If
    End If
    '
    If Len(CU$) < 12 Then GoTo 999
    VX% = Asc(Mid$(CU$, 12))
    If VX% >= 48 Then
        If VX% <= 57 Then
            CU$ = Left$(CU$, 11) + "-" + Mid$(CU$, 12)
        End If
    End If
    '
    If Len(CU$) <> 13 Then GoTo 999                                   ' no valida
    '
    For J% = 1 To 13
        If J% <> 3 Then
            If J% <> 12 Then
                VX% = Asc(Mid$(CU$, J%, 1))
                If VX% < 48 Or VX% > 57 Then GoTo 999                 ' no valida
            End If
        End If
    Next J%
    '
    Mid$(CU$, 3, 1) = "-"
    Mid$(CU$, 12, 1) = "-"
    '
    SUMA% = 0
    SUMA% = SUMA% + 5 * XVALO(Mid$(CU$, 1, 1))
    SUMA% = SUMA% + 4 * XVALO(Mid$(CU$, 2, 1))
    SUMA% = SUMA% + 3 * XVALO(Mid$(CU$, 4, 1))
    SUMA% = SUMA% + 2 * XVALO(Mid$(CU$, 5, 1))
    SUMA% = SUMA% + 7 * XVALO(Mid$(CU$, 6, 1))
    SUMA% = SUMA% + 6 * XVALO(Mid$(CU$, 7, 1))
    SUMA% = SUMA% + 5 * XVALO(Mid$(CU$, 8, 1))
    SUMA% = SUMA% + 4 * XVALO(Mid$(CU$, 9, 1))
    SUMA% = SUMA% + 3 * XVALO(Mid$(CU$, 10, 1))
    SUMA% = SUMA% + 2 * XVALO(Mid$(CU$, 11, 1))
    '
    SUMB% = SUMA% Mod 11
    DIVER% = 0
    If SUMB% Mod 11 <> 0 Then DIVER% = 11 - SUMB%
    If DIVER% <> XVALO(Mid$(CU$, 13, 1)) Then GoTo 999                  ' no valida
    '
900 VC% = 1
    CUIT$ = CU$

999 VALICUIT% = VC%
    '
End Function

Sub CARDATPRO(NC%)
    '
    Screen.MousePointer = 11
    '
    MODIPRO.Text2.TEXT = TEL2CLI$(NC%)
    MODIPRO.Text3.TEXT = RASOCLI$(NC%)
    MODIPRO.Text4.TEXT = DOMICLI$(NC%)
    MODIPRO.Text5.TEXT = CPOSCLI$(NC%)
    MODIPRO.Text6.TEXT = LOCACLI$(NC%)
    MODIPRO.Text10.TEXT = TEL1CLI$(NC%)
    MODIPRO.Text15.TEXT = CUITCLI$(NC%)
    MODIPRO.Text18.TEXT = NIBRCLI$(NC%)
    MODIPRO.Text9.TEXT = PAISCLI$(NC%)
    MODIPRO.Text11.TEXT = FAXCLI$(NC%)
    MODIPRO.Text12.TEXT = EMAILCLI$(NC%)
    '
    MODIPRO.COTEXT(0).TEXT = TRIM$(PROVCLI$(NC%))
    MODIPRO.COTEXT(1).TEXT = TRIM$(Str$(PIVACLI%(NC%)))
    MODIPRO.COTEXT(2).TEXT = TRIM$(CATIBRU$(NC%))
    MODIPRO.COTEXT(3).TEXT = TRIM$(Str$(LIPRCLI%(NC%)))
    MODIPRO.COTEXT(4).TEXT = TRIM$(Str$(CPAGCLI%(NC%)))
    MODIPRO.COTEXT(5).TEXT = TRIM$(CUEMADRE$(NC%))
    '
    For II = 0 To 5
       ECO$ = ""
       If TRIM$(MODIPRO.COTEXT(II).TEXT) <> "" Then
          TV11% = TIPFIELD(ARCHE$(II), 1)
          LF11% = LENFIELD(ARCHE$(II), 1)
          ARGU$ = CBIN$(TRIM$(MODIPRO.COTEXT(II).TEXT), TV11%, LF11%)
          ECO$ = TRIM$(ECOARCH$(ARCHE$(II), ARGU$))
       End If
       '
       If ECO$ <> "" Then
             MODIPRO.DETEXT(II).TEXT = ECO$
          Else
             MODIPRO.DETEXT(II).TEXT = "No Definido"
       End If
    Next II
    '
    Screen.MousePointer = 1
    '
End Sub
'

Sub CARTAVALI()
    '
    If ULTREG&("CARETIB") < 4 Then
      IQ& = 0
      IQ& = IQ& + 1: Call GRAREG("CARETIB", "0I.Brutos No Responsable", IQ&)
      IQ& = IQ& + 1: Call GRAREG("CARETIB", "1Contribuyente Puro", IQ&)
      IQ& = IQ& + 1: Call GRAREG("CARETIB", "2Convenio Multilateral", IQ&)
      IQ& = IQ& + 1: Call GRAREG("CARETIB", "3Conv.Multil. (Comerciales)", IQ&)
      Call SETULTREG("CARETIB", IQ&)
      Call CIERRARCH("CARETIB")
    End If
    '
    If ULTREG&("TICUEPRO") < 2 Then
      IQ& = 0
      IQ& = IQ& + 1: Call GRAREG("TICUEPRO", Chr$(0) + "Cuenta Corriente", IQ&)
      IQ& = IQ& + 1: Call GRAREG("TICUEPRO", Chr$(1) + "Cuenta de Gastos", IQ&)
      Call SETULTREG("TICUEPRO", IQ&)
      Call CIERRARCH("TICUEPRO")
    End If
    '
    If ULTREG&("CAPOSIV") < 8 Then
      IQ& = 0
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(0) + AJUSTI$("IVA Consumidor Final            ", 31) + AJUSTI$("Cons.Fin", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(1) + AJUSTI$("IVA Responsable Inscripto   ", 31) + AJUSTI$("Resp.Ins", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(2) + AJUSTI$("IVA Responsable no Inscripto", 31) + AJUSTI$("R.No Ins", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(3) + AJUSTI$("IVA No Responsable          ", 31) + AJUSTI$("No Resp.", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(4) + AJUSTI$("IVA Responsable Exento      ", 31) + AJUSTI$("Exento  ", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(5) + AJUSTI$("IVA Operac.Comercio Exterior", 31) + AJUSTI$("Com.Ext.", 8) + String$(24, 0), IQ&)
          If InStr(EMPREAC$, "SABO") > 0 Then Call GRAREG("CAPOSIV", Chr$(5) + AJUSTI$("IVA Operaciones Exportacion", 31) + AJUSTI$("Export.", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(6) + AJUSTI$("IVA Contrib.Monotributo     ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(7) + AJUSTI$("IVA Exento Zona Promovida   ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
      Call SETULTREG("CAPOSIV", IQ&)
      Call CIERRARCH("CAPOSIV")
    End If
    '
    For I& = ULTREG&("LISTAS") To 1 Step -1
       XX$ = REGLEIDO$("LISTAS", I&)
       If Asc(Left$(XX$, 1)) > 0 Then GoTo 59
       Call GRAREG("LISTAS", String$(128, 0), I&)
    Next I&
59  Call CIERRARCH("LISTAS")
    '
     For I& = ULTREG&("CONPAG") To 1 Step -1
       XX$ = REGLEIDO$("CONPAG", I&)
       If Asc(Left$(XX$, 1)) > 0 Then GoTo 69
       Call GRAREG("CONPAG", String$(128, 0), I&)
    Next I&
69  Call CIERRARCH("CONPAG")
    '
    For I& = ULTREG&("VENDEDOR") To 1 Step -1
       XX$ = REGLEIDO$("VENDEDOR", I&)
       If Asc(Left$(XX$, 1)) > 0 Then GoTo 79
       Call GRAREG("VENDEDOR", String$(128, 0), I&)
    Next I&
79  Call CIERRARCH("VENDEDOR")
    '
     For I& = ULTREG&("PROVINC") To 1 Step -1
       XX$ = REGLEIDO$("PROVINC", I&)
       If Asc(Left$(XX$, 1)) > 32 Then GoTo 89
       Call GRAREG("PROVINC", String$(128, 0), I&)
    Next I&
89  Call CIERRARCH("PROVINC")
    '
    ARCHE$(0) = "PROVINC"
    ARCHE$(1) = "CAPOSIV"
    ARCHE$(2) = "CARETIB"
    ARCHE$(3) = "LISTAS"
    ARCHE$(4) = "CONPAG"
    ARCHE$(5) = "VENDEDOR"
    ARCHE$(6) = "ECUECON"
    ARCHE$(7) = "GRUCOCLI"
    '
End Sub

Sub SORTCLI(ARDEU%)
    '
    Dim DEU1 As String * 32
    Dim DEU2 As String * 128
    Dim REGDEU$(8192)
    Dim OCUCLI%(32767)
    '
    PCLI$ = String$(32767, 0)
    '
    LU$ = LUDAT$
    ARX1$ = LU$ + "DEUDOR1.DAT"
    ASELE$ = "DEUDOR1"
    If ARDEU% = 2 Then
        ARX1$ = LU$ + "PROVED1.DAT"
        ASELE$ = "PROVED1"
    End If
    On Error GoTo 100
    N1% = FreeFile
    Open ARX1$ For Random Access Read Write Lock Read Write As #N1% Len = 32
    '
    ARX2$ = LU$ + "DEUDOR2.DAT"
    If ARDEU% = 2 Then
        ARX2$ = LU$ + "PROVED2.DAT"
    End If
    On Error GoTo 200
    N2% = FreeFile
    Open ARX2$ For Random Access Read Write Lock Read Write As #N2% Len = 128
    '
    Screen.MousePointer = 11
    AGRECLI% = 0
    On Error GoTo 0
    Lock #N1%
    Lock #N2%
    '
    FORSELEC.LISORT.Clear
    For U& = 1 To LOF(N1%) / 32
        Get #N1%, U&, DEU1$
        Get #N2%, U&, DEU2$
        If CVI(Left$(DEU1$, 2)) > 0 Then
            '
            NCXI% = CVI(Left$(DEU1$, 2))
            If NCXI% <= 32767 Then
              If OCUCLI%(NCXI%) > 0 Then
                GoTo 19       ' descarta repetidos
              End If
              OCUCLI%(NCXI%) = 1
            End If
            '
            CLI$ = Mid$(DEU1$, 3, 30)
            CUIT$ = TRIM$(Mid$(DEU2$, 75, 15))
            If VERICUIT% <> (-1) Then
              If CUIT$ <> "" Then
                If VALICUIT%(CUIT$) <> 1 Then
                  Screen.MousePointer = 1
                  TECLI$ = "Cliente": If ARDEU% = 2 Then TECLI$ = "Proveedor"
                  If COMALTER%(TECLI$ + ": " + TRIM$(CLI$) + "\C.U.I.T.: " + CUIT$ + "\  \Número de C.U.I.T. no Válido.\\Continuar Revisión\Cancelar") = 2 Then
                     VERICUIT% = (-1)
                  End If
                  Screen.MousePointer = 11
                End If
              End If
            End If
            REX$ = Space$(36)
            Call REPLA(REX$, CLI$, 1, 30)
            Call REPLA(REX$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 31, 6)
            FORSELEC.LISORT.AddItem REX$
            REGDEU$(U&) = DEU1$ + DEU2$
            '
          ElseIf CVI(Left$(DEU1$, 2)) < 0 Then
            '
            XX$ = DEU1$
            Call COPYSTRU("DEUDOR1", "DEUDOBAJ")
            Call REGAPP("DEUDOBAJ", XX$)
            Call CIERRARCH("DEUDOBAJ")
        End If
19  Next U&
    '
    URE& = 0
    For U& = 1 To FORSELEC.LISORT.ListCount
        FORSELEC.LISORT.ListIndex = U& - 1
        INDI& = XVALO(Mid$(FORSELEC.LISORT.TEXT, 31, 6))
        DEU1$ = Left$(REGDEU$(INDI&), 32)
        DEU2$ = Mid$(REGDEU$(INDI&), 33, 128)
        Put #N1%, U&, DEU1$
        Put #N2%, U&, DEU2$
        URE& = U&
    Next U&
    For U& = URE& + 1 To LOF(N1%) / 32
        DEU1$ = String$(32, 0)
        DEU2$ = String$(128, 0)
        Put #N1%, U&, DEU1$
        Put #N2%, U&, DEU2$
    Next U&
    '
    Unlock #N2%
    Unlock #N1%
    Close #N2%: N2% = 0
    Close #N1%: N1% = 0
    DEU1$ = REGLEIDO$(ASELE$, 1)
    Call GRAREG(ASELE$, DEU1$, 1)
    Call CIERRARCH(ASELE$)
    Screen.MousePointer = 1
    Exit Sub
    '
100 Resume 101
101 On Error GoTo 0
    Close #N1%: N1% = 0
    Call COMUNI("Imposible Ordenar - Archivo en Uso.")
    Exit Sub
    '
200 Resume 201
201 On Error GoTo 0
    Close #N1%: N1% = 0
    Close #N2%: N2% = 0
    Call COMUNI("Imposible Ordenar - Archivo en Uso.")
    Exit Sub
    '
End Sub
'

Function CUEIBRUPRO$(A$)
   '
   Static CTX%, COPROVI$, CUEPROVI$(64)
   If CTX% = 0 Then
     For I& = 1 To ULTREG&("PROVINB")
       X$ = REGLEIDO$("PROVINB", I&)
       If ECOARCH$("CUECON", Mid$(X$, 27, 2)) = "" Then
          Call REPLA(X$, MKI$(0), 27, 2)
       End If
       If TRIM$(Left$(X$, 1)) <> "" Then
          If Len(COPROVI$) < 64 Then
            COPROVI$ = COPROVI$ + Left$(X$, 1)
            CUEPROVI$(Len(COPROVI$)) = Mid$(ECOARCH$("CUECON", Mid$(X$, 27, 2)), 31, 12)
          End If
       End If
     Next I&
     Call CIERRARCH("PROVINB")
     CTX% = 1
   End If
   '
   A1$ = AJUSTI$(A$, 1)
   QQQ$ = Space$(12)
   PPXX% = InStr(COPROVI$, A1$)
   If PPXX% > 0 Then
     QQQ$ = CUEPROVI$(PPXX%)
   End If
   '
   CUEIBRUPRO$ = AJUSTI$(QQQ$, 12)
   '
End Function

Function CUEIPROVIN%(A$)
   '
   Static CTX%, COPROVI$, CUEPROVII%(64)
   If CTX% = 0 Then
     For I& = 1 To ULTREG&("PROVINB")
       X$ = REGLEIDO$("PROVINB", I&)
       If ECOARCH$("CUECON", Mid$(X$, 27, 2)) = "" Then
          Call REPLA(X$, MKI$(0), 27, 2)
       End If
       If TRIM$(Left$(X$, 1)) <> "" Then
          If Len(COPROVI$) < 64 Then
            COPROVI$ = COPROVI$ + Left$(X$, 1)
            CUEPROVII%(Len(COPROVI$)) = CVI(Mid$(X$, 27, 2))
          End If
       End If
     Next I&
     Call CIERRARCH("PROVINB")
     CTX% = 1
   End If
   '
   A1$ = AJUSTI$(A$, 1)
   QQI% = 0
   PPXX% = InStr(COPROVI$, A1$)
   If PPXX% > 0 Then
     QQI% = CUEPROVII%(PPXX%)
   End If
   '
   CUEIPROVIN% = QQI%
   '
End Function

Function CUEMECOB%(TIPOVAL%)
   '
   Static CMC0%
   '
   If CMC0% < 1 Then
     CMC0% = CUECOINT%(CONTROL$("", "CUECHECA"))
     If CMC0% < 1 Then
       Call MENSERR(24, "Error de Configuración.\  \Falta Definir Cuenta Contable\Asociada a la Cartera de Valores.")
       On Error Resume Next
       If DERACCE%("SETUPCHE", "Configuración Funcionamiento Gestion Valores") = 2 Then
          OPCHEQ04.Show 1
       End If
       On Error GoTo 0
       Call CIERRARCH("*.*")
       Call FINAL("")
     End If
   End If
   '
   CMC% = CMC0%
   RFF$ = RECFILT$("MECOBRA", 1, MKI$(TIPOVAL%))
   If Len(RFF$) >= 4 Then
     REME& = CVS(Mid$(RFF$, 1, 4))
     If REME& > 0 Then
       CUCU$ = Mid$(REGLEIDO$("MECOBRA", REME&), 45, 12)
       CMC% = CUECOINT%(CUCU$)
       If CMC% < 1 Then CMC% = CMC0%
     End If
   End If
   '
   CUEMECOB% = CMC%
   '
End Function

Function CUEMEPAG%(TIPOVAL%)
   '
   Static CMC0%
   '
   If CMC0% < 1 Then
     CMC0% = CUECOINT%(CONTROL$("", "CUECHECA"))
     If CMC0% < 1 Then
       Call MENSERR(24, "Error de Configuración.\  \Falta Definir Cuenta Contable\Asociada a la Cartera de Valores.")
       Call CIERRARCH("*.*")
       Call FINAL("")
     End If
   End If
   '
   CMC% = CMC0%
   RFF$ = RECFILT$("MEPAGO", 1, MKI$(TIPOVAL%))
   If Len(RFF$) >= 4 Then
     REME& = CVS(Mid$(RFF$, 1, 4))
     If REME& > 0 Then
       XX$ = REGLEIDO$("MEPAGO", REME&)
       CMC% = CVI(Mid$(XX$, 61, 2))
       If CMC% < 1 Then
         CUCU$ = Mid$(XX$, 45, 12)
         CMC% = CUECOINT%(CUCU$)
       End If
       If CMC% < 1 Then CMC% = CMC0%
     End If
   End If
   '
   CUEMEPAG% = CMC%
   '
End Function

Function CUECOINT%(CUEX$)
   '
   CUECOINT% = CUEINT%(CUEX$)
   '
End Function

Function CUECARCHE$()
   '
   CUECARCHE$ = CONTROL$("", "CUECHECA")
   CUEICHEQ% = CUECOINT%(CUECARCHE$)
   If TRIM$(CUECARCHE$) = "" Or CUEICHEQ% < 1 Then
     CUECARCHE$ = ""
     Exit Function
   End If
   '
End Function

Function CUEICARCH%()
   '
   CUECHECA$ = CONTROL$("", "CUECHECA")
   CUEICARCH% = CUECOINT%(CUECHECA$)
   If TRIM$(CUECHECA$) = "" Or CUEICARCH% < 1 Then
     CUEICARCH% = 0
   End If
   '
End Function

Function MONEMICO%(NPROX%)
   '
   MNC% = 1
   RFF$ = RECFILT$("CLAPROMO", 1, MKI$(Abs(NPROX%)))
   If Len(RFF$) >= 4 Then
     RECLA& = CVS(Left$(RFF$, 4))
     XX$ = REGLEIDO$("CLAPROMO", RECLA&)
     MNC% = Asc(Mid$(XX$, 31, 1))
   End If
   '
   MONEMICO% = MNC%
   '
End Function

Function MONEVACO%(NPROX%)
   '
   MNC% = 1
   RFF$ = RECFILT$("CLAPROMO", 1, MKI$(Abs(NPROX%)))
   If Len(RFF$) >= 4 Then
     RECLA& = CVS(Left$(RFF$, 4))
     XX$ = REGLEIDO$("CLAPROMO", RECLA&)
     MNC% = Asc(Mid$(XX$, 32, 1))
   End If
   '
   MONEVACO% = MNC%
   '
End Function

Sub SELRACLIEN(NCL1%, NCL2%, VENDSEL%)
   '
   NCMIN% = 32767: NCMAX% = 0
   IXXX$ = LOADFILE$(LUDAT$ + "DEUDOR1.X01")
   '
   For KKL& = 1 To Len(IXXX$) Step 2
     NCMX% = CVI(Mid$(IXXX$, KKL&, 2))
     If NCMX% > 0 Then
       If NCMX% < NCMIN% Then NCMIN% = NCMX%
       If NCMX% > NCMAX% Then NCMAX% = NCMX%
     End If
   Next KKL&
   '
10 VDEF$ = MKI$(NCMIN%) + MKI$(NCMAX%) + Chr$(0)
   OBX$ = OBJPLA$("DESHASCLIEVEN", VDEF$)
   If OBX$ = "" Then
     NCL1% = 0: NCL2% = 0
     Exit Sub
   End If
   '
   NCL1% = CVI(Left$(OBX$, 2))
   NCL2% = CVI(Mid$(OBX$, 3, 2))
   If NCL1% < 1 Or NCL2% < 1 Or NCL2% < NCL1% Then
     Call MENSERR(24, "Rango no Válido")
     GoTo 10
   End If
   VENDSEL% = Asc(Mid$(OBX$, 5, 1))
   '
End Sub
'
Function TELESUC$(NC%, ISUC%)
 '\\\OT-07-0010-RG-12/01/07///
  TELESUC$ = ""
  RFF$ = RECFILT$("SUCENTRE", 5, MKI$(NC%))
  For I1& = 1 To Len(RFF$) Step 4
    I& = CVS(Mid$(RFF$, I1&))
    X$ = REGLEIDO$("SUCENTRE", I&)
    SUCU1% = CVI(Mid$(X$, 1, 2))
    If SUCU1% = ISUC% Then
      TELESUC$ = TRIM$(Mid$(X$, 163, 32))
      Exit For
    End If
  Next I1&
  '\\\OT-07-0010-RG-FIN///
End Function
Function DEUCUECO#(NC%) 'devulve la deuda del cliente
    Call APERBASDAT("CABAN")
    Set DEUDA = CueCorri.OpenRecordset("SELECT * FROM SADEUPEN WHERE NuClien = " & NC%, dbOpenSnapshot)
    DEUTO# = 0
    With DEUDA
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
            Do While Not .EOF
                SALDEB# = !SaldDebe
                SALACR# = !SaldAcre
                DEUTO# = DEUTO# + SALDEB# - SALACR#
                .MoveNext
            Loop
      End If
      On Error GoTo 0
      DEUCUECO# = DEUTO#
    End With
End Function

Function VALOPEDPEN#(NC%) ' devuelve el total de los valores  de los pedidos pendientes segun el cliente
    Call APERBASDAT("PEDCLI")
    SQLX$ = "SELECT * FROM PEDDETA"
    SQLX$ = SQLX$ + " WHERE NroClie=" & NC%
    SQLX$ = SQLX$ + " AND Status < 8 "
    SQLX$ = SQLX$ + " AND STATUS <> -1 " ' EXCLUYE LAS NO APROBADAS
    SQLX$ = SQLX$ + " AND (CanPed - CantEnt) >= 0.05"
    Set DEUDA = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    DEUTO# = 0
    With DEUDA
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
            Do While Not .EOF
                CANTIDAD# = (!CANPED)
                CANTENTR# = (!CantEnt)
                PRECUNIT# = (!PreUnid)
                Moneda% = (!MoneEmis)
                
                TCambio# = 1
                If Moneda% > 1 Then TCambio = TICAMONE(Moneda%)
                DEUTO# = DEUTO# + (CANTIDAD# - CANTENTR#) * PRECUNIT# * TCambio
                .MoveNext
            Loop
      End If
      On Error GoTo 0
      VALOPEDPEN# = DEUTO#
    End With
End Function
 
Function NUMECAI$(NPRX%, FEVENCAI%)
     '
     Call APERBASDAT("CABAN")
     '
     NUCAI$ = "": FECHVENCAI% = 0
     SQLX$ = "SELECT * FROM CAJABANC "
     SQLX$ = SQLX$ + "WHERE NuProve = " & NPRX% & " "
     SQLX$ = SQLX$ + "AND TipoMovim = '" & "FCOM" & "' "
     SQLX$ = SQLX$ + "AND NroCai <> '' "
     SQLX$ = SQLX$ + "AND ISNULL(NroCai) = False "
     SQLX$ = SQLX$ + "AND FeVtoCai >= Now "
     SQLX$ = SQLX$ + "ORDER BY FechEmisi ASC"
     '
     Set FacCompras = CueCorri.OpenRecordset(SQLX$)
     With FacCompras
        On Error Resume Next
        .MoveLast
        If Err < 1 Then
          On Error GoTo 0
          NUCAI$ = !NROCAI
          FECHVENCAI% = CVINT(!FEVTOCAI)
        End If
     End With
     '
     FEVENCAI% = FECHVENCAI%
     NUMECAI$ = NUCAI$
     '
End Function

Function ALIRETBRU#(NPROVE%, MINOIMP#, MEPAGO%)
    '
    Static CTCALCU%, ALIBRU(8), MINIMPO#(8), MEPAGX%(8)
    'Static PROVANT%, ALIRETBRUANT#, MINOIMPANT#, MEPAGOANT%
    'Esta linea queda comentada ya que si se generan dos pago seguidos
    'al mismo proveedor, en el segundo traia la retencion de IIBB en cero.
    'Al comentar la linea fuerza a repetir la búsqueda.
    '
    If NPROVE% > 0 And NPROVE% = PROVANT% Then GoTo 99
    If Abs(NPROVE%) = 0 Then
        ALIRETBRUANT# = 0: MINOIMPANT# = 0: MEPAGOANT% = 0
        GoTo 99
    End If
    '
    If PIVACLI%(NPROVE% * (-1)) = 5 Then  ' Si el proveedor es importacion no se realiza calculo
        RETENC# = 0: MINOMPONIBLE# = 0: MEDIOPAGO% = 0
        GoTo 98
    End If
    '
    For U& = 1 To ULTREG&("TACATIBP")
       TCX$ = REGLEIDO$("TACATIBP", U&)
       TCY% = Asc(Mid$(TCX$, 1, 1))
       If TCY% > 0 Then
          If TCY% <= 8 Then
             ALIBRU(TCY%) = CVS(Mid$(TCX$, 33, 4))
             If ALIBRU(TCY%) > 0 Then CTCALCU% = 1
             MINIMPO#(TCY%) = CVS(Mid$(TCX$, 37, 4))
             MEPAGX%(TCY%) = CVI(Mid$(TCX$, 41, 2))
          End If
       End If
    Next U&
    '
    REBLA$ = REGBLAN$("DARIBPRO")
    RERET$ = FILTERGETRECORD$("DARIBPRO", 1, MKI$(Abs(NPROVE%)))
    CATIB1% = 1: PORCONVE = 100   ' VALORES POR DEFAULT
    If RERET$ <> REBLA$ Then
        CATIB1% = Asc(Mid$(RERET$, 3, 1))
        PORCONVE = CVS(Mid$(RERET$, 5, 4))
    End If
    RETENC# = ALIBRU(CATIB1%)
    MINOMPONIBLE# = MINIMPO#(CATIB1%)
    MEDIOPAGO% = MEPAGX%(CATIB1%)
    ' falta sacar el medio de pago
    '
    If EXISTE%(LUDAT$ + "PADIBRUR.TXT") > 0 Then
       CUITPROV$ = CUITCLI$((-1) * Abs(NPROVE%))
       XX$ = REPADIBRU$(CUITPROV$, 2)
       If TRIM$(XX$) <> "" Then
          ret$ = Mid$(XX$, 48, 4)
          RETENC# = XVALO(ret$)
       End If
    End If
    '
98  PROVANT% = NPROVE%
    ALIRETBRUANT# = RETENC#
    MINOIMPANT# = MINOMPONIBLE#
    MEPAGOANT% = MEDIOPAGO%
    '
99  ALIRETBRU# = ALIRETBRUANT#
    MINOIMP# = MINOIMPANT#
    MEPAGO% = MEPAGOANT%
    '
End Function

Function ALIPERBRU#(NCLIE%, MINIPER#, VARDO%)
    '
    'Static CLIEANT%, ALIPERBRUANT#, MINIPERANT#
    'Esta linea queda comentada ya que si se generan dos pedidos/facturas seguidos
    'al mismo cliente, en el segundo traia la percepcion de IIBB en cero.
    'Al comentar la linea fuerza a repetir la búsqueda.
    '
    Static CTTXX%, TASAA, TASAB, TASAEXP, MINIPERCE
    
    If NCLIE% > 0 And NCLIE% = CLIEANT% Then GoTo 99
    If Abs(NCLIE%) = 0 Then
        ALIPERBRUANT# = 0: MINIPERANT# = 0
        GoTo 99
    End If
    '
    If PIVACLI%(NCLIE) = 5 Then   ' Si el cliente es exportacion no se realiza calculo
        TASAX = 0: MINIPERCE = 0
        GoTo 98
    End If
    '
    If CTTXX% < 1 Then
       TASAA = XVALO(CONTROL$("PERIBRU", "TASA-A"))
       TASAB = XVALO(CONTROL$("PERIBRU", "TASA-B"))
       TASAEXP = XVALO(CONTROL$("PERIBRU", "TASA-EXP"))
       MINIPERCE = XVALO(CONTROL$("PERIBRU", "MINIPER"))
       CTTXX% = 1
    End If
    '
    TASAX = TASAA
    If VARDO% = 2 Then TASAX = TASAB
    If VARDO% = 3 Then TASAX = TASAEXP
    '
    REPERCI$ = FILTERGETRECORD$("DAPIBCLI", 1, MKI$(NCLIE%))
    CATIB1% = Asc(Mid$(REPERCI$, 3, 1))
    PORCONVE = CVS(Mid$(REPERCI$, 5, 4))
    If PORCONVE > 100 Then PORCONVE = 100
    '
    If UCase$(Mid$(REPERCI$, 9, 18)) = "EXENTO CERTIFICADO" Then
       TASAX = 0
       GoTo 98
    End If
    '
    If CATIB1% > 0 Then         ' NO ESTA EXENTO
      If PORCONVE >= 10.1 Then  ' PARTICIPACION CONVENIO MULTILATERAL
           TASAX = TASAX * PORCONVE / 100
         Else
           TASAX = 0
      End If
    End If
    If CATIB1% < 1 Then TASAX = 0
    '
    If EXISTE%(LUDAT$ + "PADIBRUP.TXT") > 0 Then
       CUITCLIEN$ = CUITCLI$(NCLIE%)
       XX$ = REPADIBRU$(CUITCLIEN$)
       If TRIM$(XX$) <> "" Then
          PER$ = Mid$(XX$, 48, 4)
          TASAX = XVALO(PER$)
       End If
    End If
    '
98  CLIEANT% = NCLIE%
    ALIPERBRUANT# = TASAX
    MINIPERANT# = MINIPERCE
    '
99  ALIPERBRU# = ALIPERBRUANT#
    MINIPER# = MINIPERANT#
    '
End Function
'
Function REPADIBRU$(CTTT$, Optional MODO%)    ' CTTT$ = CUIT DEL CLIENTE / PROVEEDOR
    '
    Static LREC%
    Static ARCHIIBB$
    Dim XX1 As String * 57
    
    If LREC% < 1 Then
       ARCHIIBB$ = "PADIBRUP.TXT"
       If MODO% = 2 Then ARCHIIBB$ = "PADIBRUR.TXT"
       NARCHI% = FILEOPEN%(LUDAT$ + ARCHIIBB$, 1, 128)
       Line Input #NARCHI%, XX$
       LREC% = 2 + Len(XX$)
       Close #NARCHI%
    End If
    '
    If LREC% <> 57 Then
       Call MENSERR(24, "Longitud de Registro Incorrecta (" + TRIM$(Str$(LREC%)) + ") en PADIBRU.TXT. Consulte.")
       Call FINAL("")
    End If
    '
    CUITPROV$ = CTTT$
    CuitProveedor$ = REPLACAR$(CUITPROV$, "-", "")
    '
    XX1 = Space$(57)
    RPDB$ = ""
    NARCHI% = FILEOPEN%(LUDAT$ + ARCHIIBB$, 0, LREC%)
    NREG& = LOF(NARCHI%) / 57
    '
    LI& = 0: LS& = NREG& + 1
11     E& = Int((LS& - LI&) / 2)
    If E& < 1 Then GoTo 30
'11  E& = Int((LS& - LI&) / 2) 'AGREGADO POR QUE PARA EL CUIT 30664029452 DEL CLIENTE 305 E& QUEDABA EN 0 Y NO DEVOLVIA EL STRING DEL  PADIBRU
'    If E& < 1 Then
'      If RPDB$ = "" Then 'PARA QUE EN ESTE CASO PASE A LA BUSQUEDA SECUENCIAL
'        GoTo 30
'      Else
'        GoTo 99
'      End If
'    End If

    L& = LI& + E&
    If L& < 1 Or L& > NREG& Then GoTo 30
    Get #NARCHI%, L&, XX1$
    CUITPADRON$ = Mid$(XX1$, 30, 11)
       If CUITPADRON$ < CuitProveedor$ Then LI& = L&: GoTo 11
       If CUITPADRON$ > CuitProveedor$ Then LS& = L&: GoTo 11
    RPDB$ = XX1$: GoTo 99          ' LO ENCONTRO EN BUSQUEDA BINARIA
    '
    ' NO LO ENCONTRO - HACE BUSQUEDA SECUENCIAL
30  Close #NARCHI%
    ' SE CAMBIA PORQUE SE ASUME QUE LOS CUITS ESTAN ORDENADOS,
    ' CON LO CUAL LA BUSQUEDA BINARIA SIEMPRE FUNCIONA. O SEA, QUE
    ' SI NO LO ENCUENTRA ES PORQUE NO EXISTE EN PADIBRU
    GoTo 99
    '
    NARCHI% = FILEOPEN%(LUDAT$ + ARCHIIBB$, 1, 128)
    Do While Not EOF(NARCHI%)
       Line Input #NARCHI%, XX$
       CUITPADRON$ = Mid$(XX$, 30, 11)
       If CUITPADRON$ = CuitProveedor$ Then
          RPDB$ = XX$: GoTo 99     ' LO ENCONTRO SECUENCIAL
       End If
    Loop
    '
99  Close #NARCHI%
    REPADIBRU$ = RPDB$
    '
End Function
'
Function VALOPEDPEN1#(NCLIE%)
   VALOPEDPEN1# = 0
   Call APERBASDAT("PEDCLI")
   SQLX$ = " SELECT SUM ((CANPED - CANTENT)* PREUNID)AS VAPEPEN FROM PEDDETA "
   SQLX$ = SQLX$ + " INNER JOIN PEDENCA ON PEDDETA.NroPed = PEDENCA.NroPed"
   SQLX$ = SQLX$ + " WHERE PEDENCA.STATUS < 9"
   SQLX$ = SQLX$ + " AND PEDENCA.NroClie = " & NCLIE% & ""
   SQLX$ = SQLX$ + " AND PEDDETA.STATUS < 8 "
   SQLX$ = SQLX$ + " AND PEDDETA.CANPED > PEDDETA.CANTENT "
   Set PEDCLITEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   With PEDCLITEMP
    If Not (.EOF And .BOF) Then
      VALOPEDPEN1 = XVALO(!VAPEPEN)
    End If
   End With
End Function
Function HAYFACVEN%(NCLIE%)
   HOII% = HOY(HOS$)
   DIAHOY = Int(CDbl(CVDAT(HOII%)))
   HFV% = 0
   '
   Call APERBASDAT("CABAN")
   If SALDCLI#(NCLIE%) > 0 Then
     SQLX$ = "SELECT SUM (SALDDEBE - SALDACRE) AS DEUDATO FROM SADEUPEN "
     SQLX$ = SQLX$ + " WHERE NUCLIEN =" & NCLIE%
     SQLX$ = SQLX$ + " AND SALDDEBE > SALDACRE "
     SQLX$ = SQLX$ + " AND INT(CDBL(FeVencim)) < " & DIAHOY & " "
     Set FACCCVEN = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With FACCCVEN
       If Not (.EOF And .BOF) Then
         If XVALO(!DEUDATO) > 0 Then HFV% = 1
       End If
     End With
     '
   End If
   HAYFACVEN% = HFV%
   
End Function

Function DESCOMBINA$(NCX%, cix%)
   '
   DCX$ = ""
   If NCX% > 0 Then
     If cix% > 0 Then
       GRX% = GRUCOCLI%(NCX%)
       LNX% = GRUCOVEI%(cix%)
       '
       DCX$ = ""
       For KU& = 1 To ULTREG&("DESCOMBI")
         RECORD$ = REGLEIDO$("DESCOMBI", KU&)
         NCY% = CVI(Mid$(RECORD$, 17, 2))
         If NCY% = 0 Or NCY% = NCX% Then
           GRY% = CVI(Mid$(RECORD$, 19, 2))
           If GRY% = 0 Or GRY% = GRX% Then
             CIY% = CVI(Mid$(RECORD$, 21, 2))
             If CIY% = 0 Or CIY% = cix% Then
                LNY% = CVI(Mid$(RECORD$, 23, 2))
                If LNY% = 0 Or LNY% = LNX% Then
                  DCY$ = TRIM$(Mid$(RECORD$, 49, 12))
                  If DCY$ <> "" Then
                     If DCX$ <> "" Then DCX$ = DCX$ + "+"
                     DCX$ = DCX$ + DCY$
                  End If
                End If
             End If
           End If
         End If
       Next KU&
     End If
   End If
   '
   DESCOMBINA$ = DCX$
End Function

Function PORCEDESCU(TDESCUEN$)
    '
    COEFI = 1
    DESCX$ = TRIM$(REPLACAR$(TDESCUEN$, "%", ""))
    '
    Do While Len(DESCX$) > 0
       PPXX% = InStr(DESCX$, "+")
       If PPXX% < 1 Then PPXX% = 1 + Len(DESCX$)
       AKX1$ = Left$(DESCX$, PPXX% - 1)
       DESCX$ = TRIM$(Mid$(DESCX$, PPXX% + 1))
       '
       PDC = XVALO(AKX1$)
       COEFI = COEFI * (1 - PDC / 100)
    Loop
    '
99  PORCEDESCU = 100 * (1 - COEFI)
    '
End Function
'
Function DESCOMBINA_VIE$(NCLIE%, CI%)
'\\\OT-08-0278-FL-22/05/08///
    DESCOMBI$ = ""
    GRUPOCLI% = GRUCOCLI%(NCLIE%)
    GRUPOLIN% = GRUCOVEI%(CI%)
    FIN& = ULTREG&("GRUFACLI")
    For U& = 1 To FIN&
       XXSS$ = REGLEIDO$("GRUFACLI", U&)
       GRUCLIEN% = CVI(Mid$(XXSS$, 1, 2))
       GRUVENTA% = CVI(Mid$(XXSS$, 3, 2))
       If GRUPOCLI% = GRUCLIEN% Then
         If GRUPOLIN% = GRUVENTA% Then
           DESCOMBI$ = TRIM(Mid$(XXSS$, 5, 28))
           Exit For
         End If
       End If
    Next U&
    DESCOMBINA_VIE$ = DESCOMBI$
'\\\OT-08-0278-FL-FIN///
End Function

Function DESCUGRULIN(NCLIE%, CI%)
    '
    'por intermedio de la funcion descombina trae el descuento que _
    tiene un grupo de Clientes con una linea de articulos el mismo _
    puede ser ejemplo 10% + 15% esta funcion lo que va a devolver es _
    el descuento combinado en un solo porcentaje
    
    PDXC = 0
    DESCX$ = REPLACAR$(DESCOMBINA$(NCLIE%, CI%), "%", " ")
    '
    If InStr(DESCX$, "+") < 1 Then
        PDXC = XVALO(DESCX$)
      Else
        XDC% = InStr(DESCX$, "+")
        DESC1 = XVALO(Left$(DESCX$, XDC% - 1))
        DESC2 = XVALO(Mid$(DESCX$, XDC% + 1))
        PDXC = 100 * (1 - (1 - DESC1 / 100) * (1 - DESC2 / 100))
    End If
    '
99  DESCUGRULIN = PDXC
    '
End Function

Function DESCULIS$(NUMERODELISTA%)
   '
   LPI% = NUMERODELISTA%
   TTXY$ = FILTERGETRECORD$("DESCLPRE", 1, MKI$(LPI%))
   DESCULIS$ = TRIM$(Mid$(TTXY$, 3))
   '
End Function

Function DESCUENLIS(NUMERODELISTA%)
    DESCUENLIS(NUMERODELISTA%) = PORCEDESCU(DESCULIS$(NUMERODELISTA%))
End Function

Function DESCULISPRE(NUMERODELISTA%)
    '
    PDXC = 0
    DESCX$ = REPLACAR$(DESCULIS$(NUMERODELISTA%), "%", " ")
    '
    If InStr(DESCX$, "+") < 1 Then
        PDXC = XVALO(DESCX$)
      Else
        XDC% = InStr(DESCX$, "+")
        DESC1 = XVALO(Left$(DESCX$, XDC% - 1))
        DESC2 = XVALO(Mid$(DESCX$, XDC% + 1))
        PDXC = 100 * (1 - (1 - DESC1 / 100) * (1 - DESC2 / 100))
    End If
    '
99  DESCULISPRE = PDXC
    '
End Function
Function ZONASUCENTRE$(NC%, NSUC%)
  ' DEVUELVE LA ZONA DE LA SUCURSAL SEGUN CLIENTE Y NUMERO DE SUCURSAL
  '
  ZONASUCENTRE$ = ""
  RFF$ = RECFILT$("SUCENTRE", 5, MKI$(NC%))
  For I1& = 1 To Len(RFF$) Step 4
    I& = CVS(Mid$(RFF$, I1&))
    X$ = REGLEIDO$("SUCENTRE", I&)
    SUCU1% = CVI(Mid$(X$, 1, 2))
    If SUCU1% = NSUC% Then
      ZONASUCENTRE$ = TRIM$(Mid$(X$, 195, 8))
      Exit For
    End If
  Next I1&
End Function

Function NUMESUC$(NPED%)
   'DEVUELVE EL NUMERO DE SUCURSAL DE LA TABLA PEDENCA PASANDOLE COMO ARGUMENTO _
   EL NUMERO DE PEDIDO DE LA TABLA PEDENCA
   NUMESUC$ = ""
   SQLX$ = "SELECT * FROM PEDENCA "
   SQLX$ = SQLX$ + " WHERE NroPed = " & NPED% & ""
   '
   Set SUCPEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
   With SUCPEDTEMP
     On Error Resume Next
     If Not .EOF Then
       On Error GoTo 0
       NUMESUKU = !NSucent
     End If
   End With
   If NUMESUKU <> "" Then
     NUMESUC$ = NUMESUKU
   End If
End Function

Function DESCUARTCLI$(CI%, NC%)
   DESCUARTCLI$ = ""
   FIN& = ULTREG&("DESCARCL")
   For U& = 1 To FIN&
    XXSS$ = REGLEIDO$("DESCARCL", U&)
    CODARTT% = CVI(Mid$(XXSS$, 1, 2))
    NUMCLIEN% = CVI(Mid$(XXSS$, 3, 2))
    If CI% = CODARTT% Then
     If NC% = NUMCLIEN% Then
      DESXT$ = TRIM(Mid$(XXSS$, 5, 16))
      Exit For
     End If
    End If
    Next U&
    DESCUARTCLI$ = DESXT$
End Function

Function DESCUCLIART(NC%, CI%)
    '
    'CALCULA PORCENTAJES COMBINADOS DE LA TABLA DESCARCL
    PDXC = 0
    DESCX$ = REPLACAR$(DESCUARTCLI$(CI%, NC%), "%", " ")
    '
    If InStr(DESCX$, "+") < 1 Then
        PDXC = XVALO(DESCX$)
      Else
        XDC% = InStr(DESCX$, "+")
        DESC1 = XVALO(Left$(DESCX$, XDC% - 1))
        DESC2 = XVALO(Mid$(DESCX$, XDC% + 1))
        PDXC = 100 * (1 - (1 - DESC1 / 100) * (1 - DESC2 / 100))
    End If
    '
99  DESCUCLIART = PDXC
    '
End Function

Sub MUESTRADOC(DOCUNU&, Optional MODOX$)
    '
    XXX$ = FILTERGETRECORD$("PDOCLST", 1, MKS$(DOCUNU&))
    EDITFORM.Label98 = TRIM$(Mid$(XXX$, 5, 60))
    '
5   For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    
    Dim TD0 As String * 512
    NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DAT", 0, 512)
    Get #NX0%, 1, TD0$
    If CVS(Left$(TD0$, 4)) > DOCUNU& Then
      Close #NX0%
      NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DT0", 0, 512)
    End If
    '
    URE& = Abs(LOF(NX0%) / 512)
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    Get #NX0%, L&, TD0$
    DCC& = CVS(Left$(TD0$, 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    GoTo 99
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      Get #NX0%, L&, TD0$
      DCC& = CVS(Left$(TD0$, 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      Get #NX0%, KKL&, TD0$
      TBUF$ = TD0$
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  Close #NX0%
    SPOOLSTRING$ = TPSP$
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7170
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7170
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    EDITFORM.Command3.Enabled = False
'    EDITFORM.Command1.Enabled = False
    '
    On Error Resume Next
    If UCase$(MODOX$) = "PRINT" Then
        Call EDITFORM.Command4_Click
      Else
        EDITFORM.Show 1
    End If
    On Error GoTo 0

99  End Sub
Sub SAVESPOOL(DESCRIDOC$, OKX%)
   '
   Static CACOMM%, PRICOMM$(128), COMMINDEX$
   OKX% = 0
   '
   If CACOMM% < 1 Then
      PRICOMM$(1) = "ORIENTATION"
      PRICOMM$(2) = "NEWPAGE"
      PRICOMM$(3) = "ENDDOC"
      PRICOMM$(4) = "HEIGHT"
      PRICOMM$(5) = "WIDTH"
      PRICOMM$(6) = "SCALEMODE"
      PRICOMM$(7) = "FORECOLOR"
      PRICOMM$(8) = "DRAWMODE"
      PRICOMM$(9) = "DRAWWIDTH"
      PRICOMM$(10) = "DRAWSTYLE"
      PRICOMM$(11) = "FILLSTYLE"
      PRICOMM$(12) = "LINE"
      PRICOMM$(13) = "FONTNAME"
      PRICOMM$(14) = "FONTITALIC"
      PRICOMM$(15) = "FONTSIZE"
      PRICOMM$(16) = "FONTBOLD"
      PRICOMM$(17) = "CURRENTX"
      PRICOMM$(18) = "CURRENTY"
      PRICOMM$(19) = "PRINT"
      PRICOMM$(20) = "FONTUNDERLINE"
      PRICOMM$(21) = "PAINTPICTURE"
      CACOMM% = 21
      COMMINDEX$ = ""
      For KKII% = 1 To CACOMM%
        COMMINDEX$ = COMMINDEX$ + AJUSTI$(PRICOMM$(KKII%), 16)
      Next KKII%
   End If
   '
   'If FORMALTER.PSPOOL.ListCount < 1 Then
'MsgBox "Entré en SAVESPOOL con CAREVEC = " & CAREVEC&
   If CAREVEC& < 1 Then
      Exit Sub
   End If
   '
'10 FIN& = FORMALTER.PSPOOL.ListCount
10 FIN& = CAREVEC&
   TPSP$ = String$(4096, 0): PINSE& = 1
   For KKIL& = 1 To FIN&
      Call TRACE(20, KKIL&, FIN&)
      'TTTT$ = TRIM$(FORMALTER.PSPOOL.List(KKIL& - 1))
      TTTT$ = TRIM$(VECPRICO$(KKIL&))
      PCOMM$ = TRIM$(Left$(TTTT$, 16))
      PATRI$ = Mid$(TTTT$, 17)
      KKKK% = InStr(COMMINDEX$, PCOMM$)
      If (KKKK% Mod 16) = 1 Then
        KKJJ% = (KKKK% + 15) / 16
        TPSQ$ = Chr$(KKJJ%) + PATRI$ + Chr$(255)
        If PINSE& + Len(TPSQ$) > Len(TPSP$) Then
          TPSP$ = TPSP$ + String$(4096, 0)
        End If
        Mid$(TPSP$, PINSE&, Len(TPSQ$)) = TPSQ$
        PINSE& = PINSE& + Len(TPSQ$)
        '+ Chr$(KKJJ%) + PATRI$ + Chr$(255)
        GoTo 19
      End If
      For KKJJ% = 1 To CACOMM%
        If PRICOMM$(KKJJ%) = PCOMM$ Then
           TPSQ$ = Chr$(KKJJ%) + PATRI$ + Chr$(255)
           If PINSE& + Len(TPSQ$) > Len(TPSP$) Then
             TPSP$ = TPSP$ + String$(4096, 0)
           End If
           Mid$(TPSP$, PINSE&, Len(TPSQ$)) = TPSQ$
           PINSE& = PINSE& + Len(TPSQ$)
           'TPSP$ = TPSP$ + Chr$(KKJJ%) + PATRI$ + Chr$(255)
           'If CANCELPRINT% < 1 Then Call PRINTCOMMAND(KKJJ%, PATRI$)
           GoTo 19
        End If
      Next KKJJ%
19 Next KKIL&
   TPSP$ = Left$(TPSP$, PINSE& - 1)
   '
'MsgBox "Armado el Vector de Transferencia."
20 On Error Resume Next
   Call BLOQARCH("PDOCLST")
   If Err Then GoTo 21
   Call BLOQARCH("PDOCSPL")
   If Err Then
21    On Error GoTo 0
      Call COMUNI("Error al Bloquear Tablas 'PDOCLST' y 'PDOCSPL'.\  \Espere un Instante y Pulse 'Aceptar'\para Reanudar Ejecución.")
      GoTo 20
   End If
   On Error GoTo 0
   '
   UDOCU& = 0
   UREGI& = ULTREG&("PDOCLST")
   If UREGI& > 0 Then
      UDOCU& = CVS(Left$(REGLEIDO$("PDOCLST", UREGI&), 4))
   End If
   PDLS$ = String$(256, 0)
   DOCUNU& = 1 + UDOCU&
   HOII% = HOY(HOS$)
   Call REPLA(PDLS$, MKS$(DOCUNU&), 1, 4)
   Call REPLA(PDLS$, DESCRIDOC$, 5, 60)
   Call REPLA(PDLS$, MKI$(HOII%), 65, 2)
   Call REPLA(PDLS$, MKI$(Int(HORANUM(Time$))), 67, 2)
   Call REPLA(PDLS$, USERNAME$, 69, 30)
   Call REPLA(PDLS$, NOMTER$, 99, 30)
     On Error Resume Next
   Call REPLA(PDLS$, Printer.DeviceName, 129, 128)
     On Error GoTo 0
   Call REGAPP("PDOCLST", PDLS$)
   Call CIERRARCH("PDOCLST")
   '
'MsgBox "Guardado Encabezado de Documento"
   KKJJ% = 0
   FIN1& = Len(TPSP$): QCOUN& = 0
   While Len(TPSP$) > 0
     QCOUN& = QCOUN& + 1
     Call TRACE(20, QCOUN&, FIN1&)
     TBUF$ = MKS$(DOCUNU&) + Left$(TPSP$ + String$(508, 0), 508)
     Call REGAPP("PDOCSPL", TBUF$)
     TPSP$ = Mid$(TPSP$, 509)
   Wend
   Call CIERRARCH("PDOCSPL")
   '
'MsgBox "Guardado Cuerpo del Documento"
   Call CIERRARCH("PIMGLST")
   Call CIERRARCH("PIMGSPL")
   '
30 On Error Resume Next
   Printer.PrintQuality = (-1)
   MATRIXPUN% = 0
   If InStr(Printer.DeviceName, "810") > 0 Then MATRIXPUN% = 1
   If InStr(Printer.DeviceName, "286") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "APEX") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "FACTURA") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "REMITO") > 0 Then MATRIXPUN% = 1
   'If InStr(Printer.DriverName, "EPSON9") > 0 Then MATRIXPUN% = 1
   On Error GoTo 0
   If MATRIXPUN% = 1 Then
      FORMALTER.List1.Clear
   End If
   '
   If CANCELPRINT% < 1 Then
     If HOJASUEL% = 1 Then
       Call COMUNI("Coloque Formulario Correspondiente a\" + TIPODOC$ + " " + NUMEDOC$ + "\y Pulse Botón 'Aceptar Mensaje'.")
     End If
   End If
   '
   If SENDBYMAIL$ <> "" Then
     IMPREDET$ = Printer.DeviceName
     Printer.TrackDefault = False
     For Each PRX1 In Printers
       If InStr(UCase$(PRX1.DeviceName), "PDFWRITER") > 0 Then
         Printer.EndDoc
         On Error Resume Next
         Set Printer = PRX1
         If Err Then
           On Error GoTo 0
           GoTo 32
         End If
         On Error GoTo 0
         ARANEX$ = DEDOCOR$
         If ARANEX$ = "" Then ARANEX$ = "PDF"
         On Error Resume Next
         Kill LUCOM$ + "" + ARANEX$ + ".PDF"
         On Error GoTo 0
         EDITFORM.Label1 = ARANEX$
         EDITFORM.Timer1.Enabled = True
         GoTo 35
       End If
     Next
     '
32   Call MENSERR(24, "Imposible Direccionar a PDFWriter")
   End If
   '
35 On Error Resume Next
'MsgBox "Comienza la Impresión"
   Printer.EndDoc   ' INICIALIZA LA IMPRESORA
   Printer.ScaleMode = 1
        'MsgBox "SCALEMODE: " & Printer.ScaleMode
        'MsgBox "SCALEWIDTH: " & Printer.ScaleWidth
        'MsgBox "SCALELEFT: " & Printer.ScaleLeft
        'MsgBox "SCALEHEIGHT: " & Printer.ScaleHeight
        'MsgBox "SCALETOP: " & Printer.ScaleTop
        'MsgBox "PAPERSIZE: " & Printer.PaperSize
   If Err Then
     Resume
     On Error GoTo 0
     If COMALTER%("Error de Comunicacion con la Impresora\\Resolver y Reintentar\Cancelar Impresion") <> 2 Then
       GoTo 35
     End If
     GoTo 40
   End If
   '
   FUERADEAREA% = 0
   FUENODISPO% = 0
   'For KKII% = 1 To FORMALTER.PSPOOL.ListCount
   For KKIIL& = 1 To CAREVEC&
      'TTTT$ = TRIM$(FORMALTER.PSPOOL.List(KKII% - 1))
      Call TRACE(20, KKIIL&, CAREVEC&)
      TTTT$ = TRIM$(VECPRICO$(KKIIL&))
      PCOMM$ = TRIM$(Left$(TTTT$, 16))
      PATRI$ = Mid$(TTTT$, 17)
      For KKJJ% = 1 To CACOMM%
        If PRICOMM$(KKJJ%) = PCOMM$ Then
           If CANCELPRINT% < 1 Then Call PRINTCOMMAND(KKJJ%, PATRI$)
               TTXX$ = TTTT$ + Chr$(10) + Chr$(13)
               TTXX$ = TTXX$ + "SCALEMODE: " & Printer.ScaleMode & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "SCALEWIDTH: " & Printer.ScaleWidth & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "SCALELEFT: " & Printer.ScaleLeft & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "SCALEHEIGHT: " & Printer.ScaleHeight & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "SCALETOP: " & Printer.ScaleTop & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "PAPERSIZE: " & Printer.PaperSize
               'MsgBox TTXX$
           GoTo 39
        End If
      Next KKJJ%
39 Next KKIIL&
   '
'MsgBox "Termina la Impresión"
   If FUERADEAREA% > 0 Then
      Call MENSERR(24, "MENSAJE CRÍTICO: Se Han Encontrado Leyendas\Fuera de Area de Impresión Horizontal o Vertical.\   \FLEXOFT ha Suprimido estas Leyendas.\  \Se Recomienda Revisar y Corregir el Tamaño de Papel\Configurado en Esta Estación de Trabajo\para la Impresora " + REPLACAR$(Printer.DeviceName, "\", "/") + ".")
   End If
   FUERADEAREA% = 0
   '
   If FUENODISPO% > 0 Then
      Call MENSERR(24, "MENSAJE CRÍTICO: El Formulario Electrónico Actual\Utiliza Fuentes que no Están Disponibles\en la Impresora " + REPLACAR$(Printer.DeviceName, "\", "/") + ".\   \Se Recomienda Revisar el Formulario Electrónico\o Utilizar Otra Impresora de la Red.")
   End If
   FUENODISPO% = 0
   '
'40 FORMALTER.PSPOOL.Clear
40 ReDim VECPRICO$(0)
   CAREVEC& = 0
   Call CLOSEFORMPRIN("")
   '
   If SENDBYMAIL$ <> "" Then
     TINI = Timer
     
50   If EXISTE%(LUCOM$ + "" + TRIM$(EDITFORM.Label1) + ".PDF") < 1 Then
       If Timer < TINI + 10 Then
         GoTo 50
       End If
       '
       Call MENSERR(24, "Imposible Enviar por Mail.\  \No se ha Generado Correctamente\el Archivo Adjunto en Formato PDF.")
       Exit Sub
     End If
     Call ENVIAPORMAIL
   End If
   OKX% = 1
   '
End Sub

Function CONTROL$(PROCE$, CLAVE$)
    '
    PROBUS$ = AJUSTI$(PROCE$, 8)
    If TRIM$(PROCE$) = "" Then PROBUS$ = ""
    CODBUS$ = AJUSTI$(CLAVE$, 8)
    '
    CONTROL$ = ""
    RFF$ = RECFILT$("FLEXCRL", 2, CODBUS$)
    For U& = 1 To Len(RFF$) Step 4
      REBU& = CVS(Mid$(RFF$, U&, 4))
      XI$ = REGLEIDO$("FLEXCRL", REBU&)
      If PROBUS$ = "" Or PROBUS$ = Left$(XI$, 8) Then
        CONTROL$ = TRIM$(Mid$(XI$, 17, 80))
        Exit For
      End If
    Next U&
    '
End Function
'
Sub GRACONTROL(PROCE$, CLAVE$, VALCON$)
    '
    PROBUS$ = TRIM$(UCase$(PROCE$))
    CODBUS$ = TRIM$(UCase$(CLAVE$))
    '
    For I& = 1 To ULTREG&("FLEXCRL")
       XI$ = REGLEIDO$("FLEXCRL", I&)
       If TRIM$(UCase$(Left$(XI$, 8))) = PROBUS$ Then
           If TRIM$(UCase$(Mid$(XI$, 9, 8))) = CODBUS$ Then
               Call REPLA(XI$, VALCON$, 17, 80)
               Call GRAREG("FLEXCRL", XI$, I&)
               Call CIERRARCH("FLEXCRL")
               Exit Sub
           End If
       End If
    Next I&
    '
    XI$ = REGBLAN$("FLEXCRL")
    Call REPLA(XI$, PROBUS$, 1, 8)
    Call REPLA(XI$, CODBUS$, 9, 8)
    Call REPLA(XI$, VALCON$, 17, 80)
    Call REGAPP("FLEXCRL", XI$)
    Call CIERRARCH("FLEXCRL")
    '
End Sub
'
Function PrinterIndex(ByVal printerName As String) As Integer
    Dim I As Integer
    
    For I = 0 To Printers.Count - 1
        If LCase(Printers(I).DeviceName) Like LCase(printerName) Then
            PrinterIndex = I
            Exit Function
        End If
    Next
    PrinterIndex = -1
End Function
