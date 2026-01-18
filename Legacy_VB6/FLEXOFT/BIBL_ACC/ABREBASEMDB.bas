Attribute VB_Name = "ABREBASEMDB"
Sub ABREBASE_MDB(RUTAX$)
   Static CTX%
   If CTX% < 1 Then
      CTX% = 1
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "FeReMovi", 8, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "NumUsuar", 3, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "IdenMovi", 10, 32, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "DeLaRemi", 12, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "Descri_Lar", 12, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "FechMov", 8, 0, 1, RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "Cod_Inte", 3, 0, 1, RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "Cod_Exte", 10, 24, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "Descrip", 10, 64, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "IdenLote", 10, 16, 1, RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "LotConsum", 12, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "PNL_Des", 10, 32, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "PLC_Has", 10, 32, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "CodiMovi", 10, 2, 1, RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "DoPriCor", 10, 12, 1, RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "DoPriLar", 10, 24, 1, RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "DoSecCor", 10, 12, 1, RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "DoSecLar", 10, 24, 1, RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "DocuOrig", 10, 24, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "NumRem_Cli", 10, 16, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "NuorItem", 3, 2, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "NuPedCli", 4, 4, 1, RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "NumeOrdPed", 3, 2, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "ReferCor", 10, 64, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "CostoRep", 7, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "ValoUnit", 7, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "ValoUnitIva", 7, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "ValTotIt", 7, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "ValTotItIva", 7, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "MoneVal", 3, 2, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "TipoCam", 7, 8, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "DepoMovi", 3, 0, 1, RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "Nume_Clie", 4, 0, 1, RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "Sucu_Clie", 3, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "Nume_Prov", 3, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "CantIngre", 7, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "CantSalid", 7, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "CantSaldo", 7, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "Verificado", 3, 1, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "Causa", 3, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "User_Tran", 10, 32, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "CodCarpImp", 10, 12, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "VerificadoPor", 3, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "FHor_Verif", 8, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "TipoTran", 3, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "OCOMPRA_RTOABIERTO", 10, 64, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "NroEnvio", 4, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "NUMPVTA", 3, 0, 1, RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "DOPRINUM", 4, 0, 1, RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "IdCalidad", 4, 0, , RUTAX$)
        Call CREACAMPO_MDB2("STOCKS", "MOVISTO", "YAMIGRADO", 3, 0, 1, RUTAX$)
   End If
   '
   On Error Resume Next
   RUTAFINAL$ = RUTAX$ + "BDACCESS"
   Set BaDaMoviExt = OpenDatabase(RUTAFINAL$ + "\STOCKS.MDB")
   On Error GoTo 0
'   SQLX$ = "SELECT * FROM MOVISTO "
'   BaDaMoviExt.OpenRecordset (SQLX$)
'   Set PARATEMP = BaDaMoviExt.OpenRecordset(SQLX$, dbOpenSnapshot)

End Sub


