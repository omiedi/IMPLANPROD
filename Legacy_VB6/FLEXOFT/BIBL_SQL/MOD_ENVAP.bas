Attribute VB_Name = "MOD_ENVAP"
Public env As New Envap

Function CARGAVENDEDORENVA%()
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "CODIGO/F8;REFERENCIA/A48", FRMZELECHO)
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "CODIGO/F8;REFERENCIA/A48", FRMZELECHO)
        FRMZELECHO.Caption = "CONDICIONES DE PAGO"

        Dim obj As New RECORXET
        Set RS1 = obj.RS_CONDIPAGO()
        Call Carga_MSF_Recordset(RS1, "CODIGO/F8;REFERENCIA/A48", FRMZELECHO.MSF2)
        Set obj = Nothing
        RS1.Close
        Set RS1 = Nothing
        FRMZELECHO.Width = 9500
        FRMZELECHO.Show 1
        CARGAVENDEDORENVA% = XVALO(RESP_ZELE_VECT(1))
End Function

