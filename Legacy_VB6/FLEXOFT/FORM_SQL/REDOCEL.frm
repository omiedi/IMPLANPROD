VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form REDOCEL 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Comprobantes Electrónicos Recuperados"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   8
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2835
      Top             =   1365
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   6825
      ScaleHeight     =   5970
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   0
      Width           =   1425
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   90
         Picture         =   "REDOCEL.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   1560
         Width           =   640
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   90
         Picture         =   "REDOCEL.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Resumen de Cuenta Corriente"
         Top             =   2220
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   90
         Picture         =   "REDOCEL.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   915
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   90
         Picture         =   "REDOCEL.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   255
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   3465
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -420
            TabIndex        =   3
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "REDOCEL.frx":0A68
         Top             =   3675
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   6600
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "REDOCEL.frx":298A
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "REDOCEL.frx":2BBE
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "REDOCEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MODORECU

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub
 
Private Sub Command4_Click()
      '
      'MOSTRAR CUENTA CORRIENTE DE LA FACTURA ACTIVA
      RECUEC07.Text1 = ""
      FILA$ = List1.List(List1.ListIndex)
      NCL1$ = TRIM$(Mid$(FILA$, 42, 5))
      RECUEC07.Text1 = NCL1$
      '
      RECUEC07.Show 1
      '
End Sub
Private Sub Form_Load()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture2)
    Call Actualizar_Lista
    '
End Sub
Sub Actualizar_Lista()
    'Call CARGALISTA(List1, "FARECAFIP", "FECHEMIS/D8;NFACTURA/A18;IMPORTE/F12.2;NUCLIEN/I5", "NUCLIEN > 0 AND STATUS < 2 ORDER BY FECHEMIS")
    
    Call CARGALISTA(List1, "FARECAFIP+SADEUPEN/FARECAFIP.NFACTURA = SADEUPEN.CODICOM_LAR", "FARECAFIP.FECHEMIS/D8;FARECAFIP.NFACTURA/A18;FARECAFIP.IMPORTE/F12.2;FARECAFIP.NUCLIEN/I5", "FARECAFIP.NUCLIEN > 0  ", "DISTINCT")


End Sub
 
Private Sub List1_DblClick()
    '
 
If List1.ListIndex < 0 Then Exit Sub
    '
    FILA$ = List1.List(List1.ListIndex)
    FACTURA$ = Mid$(FILA$, 10, 18)
    NC% = XVALO(Mid$(FILA$, 42, 5))
    '
    'CARGA DATOS EN FORFASQL
    '
    FORFASQL.Text1 = NC%
    Call FORFASQL.Text1_KeyPress(13)

    '
    'ESTABLECE TIPO DE COMPROBANTE Y PUNTO DE VENTA
    Tipo$ = Mid$(FACTURA$, 1, 2)
    PVTA$ = Mid$(FACTURA$, 4, 4)
    If Tipo$ = "FC" Or Tipo$ = "ND" Then
        Do While FORFASQL.Label7 <> "N.Crédito:"
           Call FORFASQL.Label7_Click
        Loop
    ElseIf Tipo$ = "NC" Then
        Do While FORFASQL.Label7 <> "Factura:"
           Call FORFASQL.Label7_Click
        Loop
    End If
    Do While FORFASQL.Label12 <> PVTA$
        Call FORFASQL.Label12_DblClick
    Loop
    If Control("FACTURA", "PUVENELE") = PVTA$ Then
        Do While FORFASQL.Label6 <> "(T.Libre)"
           Call FORFASQL.Label6_Click
        Loop
    Else
        Do While FORFASQL.Label6 <> "(Mat.Codif)"
           Call FORFASQL.Label6_Click
        Loop
    End If
    '
    'CARGA DETALLE PARA FACTURA ELECTRONICA COMUN
    If Control("FACTURA", "PUVENELE") = PVTA$ Or Control("FACTURA", "PUVEFAEX") = PVTA$ Then
        Campox$ = "IHABECOMP"
        If Tipo$ = "NC" Then Campox$ = "IDEBECOMP"
        IMPO# = XVALO(VALOBADA("CAJABANC", Campox$, "CODICOM_LAR='" & FACTURA$ & "' AND NUCLIEN>0 AND TIPOMOVIM='VENTA'"))
        XX$ = REGBLAN$("CARDEFAC")
        Call REPLA(XX$, MKS$(1), 1, 4)
        Call REPLA(XX$, "Cancela comprobante " & FACTURA$, 7, 64)
        Call REPLA(XX$, MKS$(IMPO#), 429, 4)
        Call REPLA(XX$, MKD$(IMPO#), 433, 8)
        Call GRAREG("!CARDEFAC", XX$, 1)
        Call SETULTREG("!CARDEFAC", 1)
        '
        FORFASQL.MARCONOTA.Visible = False
        Call FORFASQL.CARDETFAC
        FORFASQL.PORDESCU.TEXT = ""
        '
    ElseIf Control("FACTURA", "PUVENCAP") = PVTA$ Then
        'CARGA DETALLE DE FACTURA DE BIENES DE CAPTIAL
        CARPETOK$ = "C:\FLEXOFT\FACELEC\"
        Condi$ = "CODICOM_LAR = '" & FACTURA$ & "'"
        XXX = Shell(CARPETOK$ + "wsbfe_cliente.exe" + " RECU " & FACTURA$, 4)
        Timer1.Enabled = True
    End If
    '
    MODORECU = 1
    FORFASQL.Show 1
    '
End Sub
 
    '
Private Sub Timer1_Timer()
      '
      Static FACTU$, CONTA%
      If TRIM$(FACTU$) = "" Then
        FILA$ = List1.List(List1.ListIndex)
        FACTU$ = Mid$(FILA$, 10, 18)
      End If
      CONTA% = CONTA% + 1 ' SI SUPERA LOS 40 SEGUNDOS
      If CONTA% > 40 Then
        Timer1.Enabled = False
        CONTA% = 0
        'NO PUDE RECUPERAR EL COMPROBANTE Y LO GUARDO EN LA LISTA DE NO RECUPERADOS
        FACTU$ = ""
        Exit Sub
      End If
      '
      'VERIFICA SI SE PUDO RECUPERAR EL COMPROBANTE DESDE LA AFIP
      KAE$ = LeeDetalleArch$(FACTU$, MOTIVO$)   ' LEO EL ARCHIVO DEVUELTO.
      '
      If XVALO(KAE$) > 0 Then ' SI LO DEVUELTO TIENE CAE PUEDO RECUPERAR EL COMPROBANTE
         Timer1.Enabled = False
         CONTA% = 0
          'GRABAR EL CAMBIO DE STATUS EN LA GRABACION DE LA FACTURA
''         CONDI$ = "NFACTURA = '" & FACTU$ & "'"
''         Call ACTUREGISTRO("FARECAFIP", "STATUS", CONDI$, 2, REGAF&, "NOMESS")
         'EL SIGUIENTE CODIGO ES PARA CALCULAR EL DESCUENTO QUE SE APLICO A LA
         'FACTURA EN EL MOMENTO QUE SE PROCESO Y ASI TRASLADARLO A LA NOTA DE CREDITO
         NETO# = XVALO(VALOBADA("CAJABANC", "VALABSCOMP", "CODICOM_LAR='" & FACTU$ & "' AND TIPOMOVIM='VENTA'", "NOMESS"))
         IMPOITEM# = 0
         For i& = 1 To ULTREG&("!CARDEFAC")
            x$ = REGLEIDO$("!CARDEFAC", i&)
            IMPOITEM# = IMPOITEM# + CVD(Mid$(x$, 433, 8))
         Next i&
         If (IMPOITEM# - NETO#) > 0.005 Then
            DESC# = 0
            If IMPOITEM# > 0.005 Then DESC# = (1 - (NETO# / IMPOITEM#)) * 100
         End If
         FACTU$ = ""
         FORFASQL.PORDESCU = FORMATNUM$(DESC#, "F8.2")
      End If
      '
End Sub
 
Function LeeDetalleArch$(FACTU$, MOTIVO$)
    '
''    'TESTEO BORRAR
''    FACTU$ = "FC-0005-00000093-A"
    
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    LeeDetalleArch$ = ""
    Cae$ = ""
    MOTIVO$ = ""
    J& = 0
    DETALLE1% = 0
    '
    'VERIFICO SI EXISTE EL ARCHIVO CON LOS DATOS DEL COMPROBANTE A RECUPERAR
    If EXISTE%(CARPETOK$ + "CAE_" + FACTU$ + ".TXT") > 0 Then
        'VERIFICO SI EL CONTENIDO DEL ARCHIVO TIENE EL CAE. EN CASO DE QUE SI,
        'SE ASUME QUE EL COMPROBANTE SE HA PODIDO RECUPERAR
        NARCHIVO% = FILEOPEN%(CARPETOK$ + "CAE_" + FACTU$ + ".TXT", 1, 128)  'ABRE EL ARCHIVO DE LECTURA
        Do While Not EOF(NARCHIVO%) ' RECORRE
          Line Input #NARCHIVO%, XXXXX$ ' LEE CADA REGISTRO
          TX1$ = TX1$ + XXXXX$ + vbCrLf
          If InStr(UCase$(XXXXX$), UCase$("<cae>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<cae>")) + Len("<cae>")
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             Cae$ = Mid$(TXT$, 1, POS2% - 1)
          End If
 
          'VALIDA QUE EMPIECE EL BLOQUE DE LOS ITEMS
          If InStr(UCase$(XXXXX$), UCase$("<ITEMS>")) > 0 Then DETALLE1% = 1
          If InStr(UCase$(XXXXX$), UCase$("</ITEMS>")) > 0 Then Exit Do
          If DETALLE1% > 0 Then
              'CODIGO
               If InStr(UCase$(XXXXX$), UCase$("<PRO_CODIGO_SEC>")) > 0 Then
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<PRO_CODIGO_SEC>")) + Len("<PRO_CODIGO_SEC>")
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 CI% = CODINT%(Mid$(TXT$, 1, POS2% - 1))
               'CANTIDAD
               ElseIf InStr(UCase$(XXXXX$), UCase$("<PRO_QTY>")) > 0 Then
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<PRO_QTY>")) + Len("<PRO_QTY>")
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 CANTI = XVALO(Mid$(TXT$, 1, POS2% - 1))
               'PRECIO UNITARIO
               ElseIf InStr(UCase$(XXXXX$), UCase$("<PRO_PRECIO_UNI>")) > 0 Then
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<PRO_PRECIO_UNI>")) + Len("<PRO_PRECIO_UNI>")
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 IMPO# = XVALO(Mid$(TXT$, 1, POS2% - 1))
               'IMPORTE TOTAL
               ElseIf InStr(UCase$(XXXXX$), UCase$("<IMP_TOTAL>")) > 0 Then
                 POS1% = InStr(UCase$(XXXXX$), UCase$("<IMP_TOTAL>")) + Len("<IMP_TOTAL>")
                 TXT$ = Mid$(XXXXX$, POS1%)
                 POS2% = InStr(TXT$, "</")
                 IMPOTOTAL# = XVALO(Mid$(TXT$, 1, POS2% - 1))
               'SI TERMINO BLOQUE DE ITEM GRABA LA FILA
               ElseIf InStr(UCase$(XXXXX$), UCase$("</ITEM>")) > 0 Then
                    J& = J& + 1
                    XX$ = REGBLAN$("CARDEFAC")
                    Call REPLA(XX$, MKS$(CANTI), 1, 4)
                    Call REPLA(XX$, MKI$(CI%), 5, 2)
                    Call REPLA(XX$, DESCRIT$(CI%), 7, 64)
                    Call REPLA(XX$, "S/N", 139, 64)
                    Call REPLA(XX$, MKS$(IMPO#), 429, 4)
                    Call REPLA(XX$, MKD$(IMPOTOTAL#), 433, 8)
                    Call GRAREG("!CARDEFAC", XX$, J&)
               End If
          End If
          
        Loop
        '
        FORFASQL.MARCONOTA.Visible = False
        FORFASQL.PORDESCU.TEXT = ""

        Call SETULTREG("!CARDEFAC", J&)
        Call FORFASQL.CARDETFAC
        Close NARCHIVO%
        
        LeeDetalleArch$ = Cae$
 
        'SI EL ARCHIVO NO CONTIENE EL CAE, ENTONCES NO SE HA PODIDO RECUPERAR EL COMPROBANTE
        'SE BUSCA SI EXISTE ALGUN MOTIVO POR EL CUAL NO SE PUDO RECUPERAR EL COMPROBANTE
        If XVALO(Cae$) < 1 Then
          MOTIVO1$ = ""
          MOTIVO2$ = ""
          VUELTA% = 0 ' LECTURA DE MOTIVO DEL ARCHIVO CAE.TXT
          NARCHIVO% = FILEOPEN%(CARPETOK$ + "CAE_" + FACTU$ + ".TXT", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
          Do While Not EOF(NARCHIVO%) ' RECORRE
              Line Input #NARCHIVO%, XX$ ' LEE CADA REGISTRO
              If InStr(XX$, "<motivo>") > 0 Or InStr(XX$, "<Obs>") > 0 Then
                 POS1% = InStr(XX$, ">")
                 POS2% = InStr(XX$, "</")
                 HASTA% = POS2% - (POS1% + 1)
                 If VUELTA% < 1 Then ' SI ES LA PRIMERA VEZ ASIGNA MOTIVO1
                  MOTIVO1$ = Mid$(XX$, POS1% + 1, HASTA%)
                  VUELTA% = VUELTA% + 1
                 ElseIf VUELTA% > 0 Then ' LA SEGUNDA VEZ QUE PASA CARGA MOTIVO2
                  MOTIVO2$ = "-" & Mid$(XX$, POS1% + 1, HASTA%)
                 End If
              End If
          Loop
          '
          Close NARCHIVO%
          MOTIVO$ = MOTIVO1$ & MOTIVO2$ ' ASIGNA LOS DOS MOTIVOS
          LeeDetalleArch$ = MOTIVO$
          '
        End If
    End If
    '
End Function

