VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form PERCEIIBB 
   Caption         =   "Percepciones de Ingresos Brutos"
   ClientHeight    =   3105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   7890
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid mfgPercepciones 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4895
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono Medium"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "PERCEIIBB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cliente As Long
Public BaseCABA As Double
Public BaseBuenosAires As Double
Public BaseMisiones As Double
Public BaseTucuman As Double
Public Moneda As Integer

Public ImportePercepcionCABA As Double
Public ImportePercepcionBuenosAires As Double
Public ImportePercepcionMisiones As Double
Public ImportePercepcionTucuman As Double


Sub CalcularPercepciones()

     Dim AlicuotaCABA As Double
     Dim AlicuotaBuenosAires As Double
     Dim AlicuotaMisiones As Double
     Dim AlicuotaTucuman As Double
     
     'SE MODIFICA PARA PASARLE EL VALOR DE VARIDOC POR QUE LO ESTABA PASANDO SIN VALOR
     PIVADOC% = PIVACLI%(Cliente)
     '
     'SE ASIGNA VALOR DE VARIDOC
     VARIDOC% = 1
     If PIVADOC% = 0 Or PIVADOC% = 3 Or PIVADOC% = 4 Or PIVADOC% = 6 Then
         VARIDOC% = 2
       ElseIf PIVADOC% = 5 Or PIVADOC% = 7 Then
         VARIDOC% = 3
         If LETREXPO$ = "A" Then
           VARIDOC% = 1
         End If
     End If
     '
    ' *** CALCULO PERCEPCION IIBB PROVINCIA DE BUENOS AIRES
    ' *** SI ESTA CONFIGURADA LA CUENTA CONTABLE, ENTONCES SE HABILITA LA PERCEPCION DE IIBB
    CUEPERI$ = Control$("PERIBRU", "CUECONTA")
    CUENPER% = CUECOINT%(CUEPERI$)
    If TRIM$(CUEPERI$) <> "" And CUENPER% > 0 Then
        AlicuotaBuenosAires = ALIPERBRU#(Cliente, MINIPER#, VARIDOC%)
        '
        ImportePercepcionBuenosAires = 0
        If BaseBuenosAires * TICAMONE(Moneda) >= MINIPER# Then
          ImportePercepcionBuenosAires = ROUND#(BaseBuenosAires * (AlicuotaBuenosAires / 100), 2)
        End If
    End If
    
    ' *** CALCULO PERCEPCION IIBB PROVINCIA DE MISIONES
    ' *** SI ESTA CONFIGURADA LA CUENTA CONTABLE, ENTONCES SE HABILITA LA PERCEPCION DE IIBB
    CUEPERI$ = Control$("PERIBMI", "CUECONTA")
    CUENPER% = CUECOINT%(CUEPERI$)
    If TRIM$(CUEPERI$) <> "" And CUENPER% > 0 Then
        AlicuotaMisiones = ALIPERBRU_MISIONES#(Cliente, MINIPER#, VARIDOC%)
        '
        ImportePercepcionMisiones = 0
        If BaseMisiones * TICAMONE(MONEMI%) >= MINIPER# Then
            ImportePercepcionMisiones = ROUND#(BaseMisiones * (AlicuotaMisiones / 100), 2)
        End If
    End If
    
'    ' *** CALCULO PERCEPCION IIBB PROVINCIA DE TUCUMAN
'    ' *** SI ESTA CONFIGURADA LA CUENTA CONTABLE, ENTONCES SE HABILITA LA PERCEPCION DE IIBB
'    CUEPERI$ = Control$("PERIBTU", "CUECONTA")
'    CUENPER% = CUECOINT%(CUEPERI$)
'    If TRIM$(CUEPERI$) <> "" And CUENPER% > 0 Then
'        CodigoProvinciaTucuman$ = Control("IIBBTUCU", "CODPROVI")
'        If PROVCLI$(Cliente) = CodigoProvinciaTucuman$ Then
'            AlicuotaTucuman = ALIPERBRU_TUCUMAN#(Cliente, MINIPER#, VARIDOC%)
'            '
'            If BaseTucuman * TICAMONE(MONEMI%) >= MINIPER# Then
'                ImportePercepcionTucuman = ROUND#(BaseTucuman * (AlicuotaTucuman / 100), 2)
'            End If
'        End If
'    End If
'
'    '
'    ' *** CALCULO PERCEPCION IIBB CABA
'    If Label7.Caption <> "N.Débito:" Then    ' NO APLICA A LAS NOTAS DE DEBITO
'        If EXISTE%(LUDAT$ + "PAIBCABA.TXT") > 0 Or EXISTE%(LUDAT$ + "IIBBCABA.TXT") > 0 Then
'             TASAX# = ALIPERBRU_CABA#(NCLIE0, MINIPER#, VARIDOC%)
'             Text24 = ""
'             BIMPON# = XVALO(Total(2))
'             If BIMPON# * TICAMONE(MONEMI%) >= MINIPER Then
'                 PERSALTA# = BIMPON# * TASAX# / 100
'                 Text24 = TRIM$(FORMATNUM$(PERSALTA#, "F8.2"))
'             End If
'             If (Label7.Caption = "N.Crédito:" And AnulaFacturaComprobante$ = "") Then
'                Text24 = ""
'             End If
'        End If
'    End If
    
    With mfgPercepciones
        .Rows = 5
        
        .TextMatrix(1, 1) = "CABA"
        .TextMatrix(1, 2) = FORMATNUM$(BaseCABA, "F12.2")
        .TextMatrix(1, 3) = FORMATNUM$(AlicuotaCABA, "F5.2")
        .TextMatrix(1, 4) = FORMATNUM$(ImportePercepcionCABA, "F12.2")
        
        .TextMatrix(2, 1) = "Buenos Aires"
        .TextMatrix(2, 2) = FORMATNUM$(BaseBuenosAires, "F12.2")
        .TextMatrix(2, 3) = FORMATNUM$(AlicuotaBuenosAires, "F5.2")
        .TextMatrix(2, 4) = FORMATNUM$(ImportePercepcionBuenosAires, "F12.2")
        
        .TextMatrix(3, 1) = "Misiones"
        .TextMatrix(3, 2) = FORMATNUM$(BaseMisiones, "F12.2")
        .TextMatrix(3, 3) = FORMATNUM$(AlicuotaMisiones, "F5.2")
        .TextMatrix(3, 4) = FORMATNUM$(ImportePercepcionMisiones, "F12.2")
        
        .TextMatrix(4, 1) = "Tucuman"
        .TextMatrix(4, 2) = FORMATNUM$(BaseTucuman, "F12.2")
        .TextMatrix(4, 3) = FORMATNUM$(AlicuotaTucuman, "F5.2")
        .TextMatrix(4, 4) = FORMATNUM$(ImportePercepcionTucuman, "F12.2")
    End With
    
End Sub

Function Bases() As Double

    Bases = BaseCABA + BaseBuenosAires + BaseMisiones + BaseTucuman
    
End Function

Function TotalPercepciones() As Double

    TotalPercepciones = ImportePercepcionCABA + ImportePercepcionBuenosAires + ImportePercepcionMisiones + ImportePercepcionTucuman
    
End Function

Private Sub Form_Load()

    Call CARGA_ENCABEZADO(mfgPercepciones, "Jurisdicción/A22;Base/F14.2;Alícuota/F9.2;Importe/F12.2", FRMZELECHO)
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Hide
End Sub
