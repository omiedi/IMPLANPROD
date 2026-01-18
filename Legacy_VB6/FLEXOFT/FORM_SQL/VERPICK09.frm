VERSION 5.00
Begin VB.Form VERPICK09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verificación Lista de Picking"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5700
      TabIndex        =   41
      Top             =   780
      Width           =   405
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   920
      Left            =   8295
      ScaleHeight     =   915
      ScaleWidth      =   1065
      TabIndex        =   39
      Top             =   5950
      Width           =   1065
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   165
         Picture         =   "VERPICK09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Transferencia a Deposito Destino"
         Top             =   120
         Width           =   800
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6090
      TabIndex        =   38
      Top             =   765
      Width           =   175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2475
      MultiLine       =   -1  'True
      TabIndex        =   31
      Top             =   7320
      Width           =   5715
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   27
      Top             =   6900
      Width           =   3975
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Verificados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   120
      TabIndex        =   17
      Top             =   4065
      Width           =   8100
      Begin VB.Frame Frame14 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   1885
         TabIndex        =   23
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame13 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   3840
         TabIndex        =   22
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame12 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   3240
         TabIndex        =   21
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame11 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   5040
         TabIndex        =   20
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame10 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   6210
         TabIndex        =   19
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame9 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   7320
         TabIndex        =   18
         Top             =   360
         Width           =   15
      End
      Begin VB.ListBox List2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   210
         MousePointer    =   1  'Arrow
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   540
         Width           =   7680
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Codigo                         Id.Bulto                U-M            Asignado                 Neto         Kg Bruto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   25
         Top             =   315
         Width           =   7680
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Asigados Pendientes de Verificar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   120
      TabIndex        =   4
      Top             =   1185
      Width           =   8100
      Begin VB.Frame Frame7 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   7320
         TabIndex        =   16
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   6210
         TabIndex        =   15
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   5040
         TabIndex        =   14
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   3240
         TabIndex        =   10
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   3840
         TabIndex        =   9
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   1885
         TabIndex        =   8
         Top             =   360
         Width           =   15
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   210
         MousePointer    =   1  'Arrow
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   560
         Width           =   7680
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Codigo                         Id.Bulto                U-M            Asignado                 Neto         Kg Bruto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   7
         Top             =   315
         Width           =   7680
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   8200
      Left            =   8400
      ScaleHeight     =   8145
      ScaleWidth      =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command14 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "VERPICK09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1845
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Left            =   105
         Picture         =   "VERPICK09.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   850
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   2
         Top             =   7150
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
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
         Left            =   105
         Picture         =   "VERPICK09.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Salir"
         Top             =   200
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "VERPICK09.frx":0BA0
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   7480
         Width           =   1515
      End
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2415
      TabIndex        =   37
      Top             =   780
      Width           =   1845
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6360
      TabIndex        =   36
      Top             =   780
      Width           =   1845
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dep. Destino:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4440
      TabIndex        =   35
      Top             =   855
      Width           =   1245
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   34
      Top             =   780
      Width           =   405
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dep. Origen:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   33
      Top             =   810
      Width           =   1680
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción / Referencia:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   30
      Top             =   7350
      Width           =   2250
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   29
      Top             =   420
      Width           =   6285
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Destino:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   28
      Top             =   450
      Width           =   1680
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Verifique Código de Barras >>>>>>>>:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   26
      Top             =   6945
      Width           =   5160
   End
   Begin VB.Label Label24 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   13
      Top             =   60
      Width           =   5445
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Top             =   60
      Width           =   720
   End
   Begin VB.Label Label34 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Numero de Lista:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   11
      Top             =   150
      Width           =   1680
   End
End
Attribute VB_Name = "VERPICK09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command14_Click()
 If DERACCE%("CFGRECE", "Configuración de Recepción") > 1 Then
    Timer1.Enabled = False
    Command14.Enabled = False
    CFGRECEP.Show 1
    Command14.Enabled = True
 End If
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Timer1.Enabled = False
    '
    DEPOTRANS% = XVALO(CONTROL$("BOLRECEP", "DEPTRANS"))
    '
    If XVALO(Text3) = DEPOTRANS% Then
        COMUNI ("Imposible Transferir desde Depósito de Material en Transito")
       GoTo 99
    End If
    '
    If List1.ListCount > 0 Then
       COMUNI ("No se han Validado Todos los Bultos")
       GoTo 99
    End If
    '
    If List2.ListCount < 1 Then
       COMUNI ("Nada Para Transferir")
       GoTo 99
    End If
    '
    For Z% = 0 To List2.ListCount - 1
      Cantidad# = XVALO(Mid$(List2.List(Z%), 46, 11))
      If Cantidad# < 0.00005 Then
         COMUNI ("Todos los Items Deben Tener Cantidad")
         GoTo 99
      End If
    Next Z%
    '
    'Se obtienen datos para registrar en MOVISTO y LOTINGRE
    DepOrig% = XVALO(Label7)
    If DepOrig% < 1 Then
        Call MENSERR(24, "Depósito ORIGEN Indefinido - Consulte")
        GoTo 99
    End If
    '
    If TRIM(Text3) = "" Then
       COMUNI ("Falta Ingresar Depósito Destino")
       GoTo 99
    End If
    '
    If TRIM$(Text3) <> "0" And ECOARCH$("TADEPOSI", Chr$(XVALO(Text3))) = "" Then
       COMUNI ("Depósito Destino no Válido")
       GoTo 99
    End If
    '
    DEPDESTINO% = XVALO(Text3)                 ' Deposito DESTINO (central de pesadas)
    If DEPDESTINO% = DepOrig% Then
        Call MENSERR(24, "Imposible Transferir.\Coinciden Deposito ORIGEN y DESTINO.")
        GoTo 99
    End If
    '
    HOII% = HOY(HOS$)
    PROVE% = 0
    
    If DEPDESTINO% > 0 Then
      If ECOARCH$("TADEPOSI", Chr$(DEPDESTINO%)) = "" Then
         Call MENSERR(24, "Imposible Transferir.\Depósito DESTINO no Definido.")
         GoTo 99
      End If
      REGI$ = FILTERGETRECORD("TADEPOSI", 1, Chr$(DEPDESTINO%))
      PROVE% = CVI(Mid(REGI$, 61, 2))                   ' Nro Proveedor
      If PROVE% > 0 Then
        If TRIM$(RASOCLI$(-1 * PROVE%)) = "" Then
          Call MENSERR(24, "Error de Configuración.\Proveedor no Válido para Depósito " & DEPDESTINO%)
          GoTo 99
        End If
      End If
    End If
    '
    If ECOARCH$("TADEPOSI", Chr$(DEPOTRANS%)) = "" Then
        Call MENSERR(24, "Debe Definir Depósito de Material en Transito")
        GoTo 99
    End If
    '
    'ASIGNO MAIL DE ADMINISTRACION Y LO VALIDO
    EMAI$ = CONTROL$("BOLRECEP", "REMIMAIL")
    If EMAI$ = "" Or InStr(EMAI$, "@") < 1 Then
       Call MENSERR(24, "Imposible Transferir.\  \Dirección de Correo de Administracion es\Inexistente o no Válida.")
       GoTo 99
    End If
    '
    'VALIDO QUE SE ENCUENTRE INSTALADA IMPRESORA VIRTUAL PDF
    For Each PRX1 In Printers
       If InStr(UCase$(PRX1.DeviceName), "PDFWRITER") > 0 Then GoTo 25
    Next
    Call MENSERR(24, "Imposible Enviar por Mail.\  \Falta Impresora Virtual PDFWRITER.")
    SENDBYMAIL$ = ""
    GoTo 99
    
    'Inicializacion de variables
    '
25  Item% = 1
    Referencia$ = TRIM$(Label4)
    TIPO$ = "I"
    LOTE$ = ""
    RAZON$ = ""
    DOMI$ = ""
    CPOSTAL$ = ""
    LOCLI$ = ""
    CUIT$ = ""
    NUMCO$ = ""
    DEPO$ = ""                 'Deposito Origen, Pendiente, hay que determinar valor
    COMPROB$ = "TR"
    '
    DESTI$ = Mid(REGI$, 2, 20)                      'Deposito Destino descripcion
    RAZON$ = RASOCLI(PROVE% * (-1))
    PRONUM$ = TRIM$(Str$(ProNroComprobante("TR")))  'Proximo Nro Comprobante
    While Len(PRONUM$) < 6: PRONUM$ = "0" & PRONUM$: Wend
    PRONUM$ = COMPROB$ & "." & PRONUM$              'ProNum se usa para guardar en la base de datos
    NUMCO$ = PRONUM$                                'NUMCO es para mostrar en las impresiones
    '
    'TIPODOC = "Vale de Transferencia / Consumo"
    TIPODOC = "Vale de Transferencia - Consumo"     'se modifica / pòr - para poder generar archivo pdf
    If PROVE% > 0 Then
       'Si es deposito externo se Obtienen datos del Proveedor
       TIPO$ = "E"
       DOMI$ = DOMICLI$(PROVE% * (-1))
       CPOSTAL$ = CPOSCLI$(PROVE% * (-1))
       LOCLI$ = LOCACLI$(PROVE% * (-1))
       CUIT$ = CUITCLI$(PROVE% * (-1))
       COMPROB$ = "RC"
       NUMCO$ = InputBox("Ingrese Nro de Remito: ", "INGRESO DE REMITO")
       If TRIM$(NUMCO$) = "" Then
          COMUNI ("Falta Ingresar Nro De Remito")
          GoTo 99
       End If
       PRONUM$ = NUMCO$
       TIPODOC$ = "Remito de Traslado"
    End If
    '
    CODPARAM$(1) = "RASO-CLI": DOCPARAM$(1) = RAZON$
    CODPARAM$(2) = "DOMI-CLI": DOCPARAM$(2) = DOMI$
    CODPARAM$(3) = "CPOS-CLI": DOCPARAM$(3) = CPOSTAL$
    CODPARAM$(4) = "LOCA-CLI": DOCPARAM$(4) = LOCLI$    'localidad
    CODPARAM$(5) = "NUME-COM": DOCPARAM$(5) = NUMCO$
    CODPARAM$(6) = "REF-MAT1": DOCPARAM$(6) = Text2     'Descripcion
    CODPARAM$(7) = "CUIT-CLI": DOCPARAM$(7) = CUIT$
    CODPARAM$(8) = "DEPOSITO": DOCPARAM$(8) = Label11   'Deposito Origen
    CODPARAM$(9) = "DESTINO": DOCPARAM$(9) = Label10    'Deposito destino
    
    CAPARDOC% = 9
    FECHDOC$ = FECHATEX$(HOII%)
    NUMEDOC$ = TRIM$(NUMCO$)
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "NRO-LOTE"
    CODTABU1$(4) = "UNIMED"
    CODTABU1$(5) = "CANTIDAD"
    CACOLTAB1% = 5
    '
    DESTIDOC% = NC1%
    CAITAB1% = 0
    '
    'Se Extrae información de Cada Items de la 'Lista de Verificados'
    For AX% = 0 To List2.ListCount - 1
      '
      Codigo$ = TRIM$(Mid$(List2.List(AX%), 1, 16))
      LOTE$ = TRIM$(Mid$(List2.List(AX%), 17, 12))
      DOPRILAR$ = PRONUM$
      UNIMEDIDA$ = (Mid$(List2.List(AX%), 31, 2))
      DepOrigen% = VALOBADA("LOTINGRE", "DEPOUBIC", "idenlote = '" & LOTE$ & "'")
      DOSELAR$ = TRIM$(VALOBADA("BALAREQMO", "DESTINO", "NULISPICK = " & Label33))
      Cantidad# = XVALO(Mid$(List2.List(AX%), 46, 11))
      CIXI% = CODINT(Codigo$)
      '
      'Registro de Salida del Deposito Actual
      Call MOVISTOK(HOII%, CIXI%, LOTE$, COMPROB$, DOPRILAR$, DOPRILAR$, DOSELAR$, DOSELAR$, Referencia$, Item%, 0, 1, 0, DepOrigen%, Cantidad# * (-1))
      
      'Registro de Entrada en el Deposito de Central
      If DEPDESTINO% > 0 Then
        '
        If PROVE% > 0 Then           'PROVE% = proveedor asignado al deposito de destino
           DEPDESTINO% = DEPOTRANS%  'Si el deposito tiene proveedor asignado, el deposito de destino pasa a ser el deposito de transito
           REFEREN$ = "Material en Transito"
           Call MOVISTOK(HOII%, CIXI%, LOTE$, COMPROB$, DOPRILAR$, DOPRILAR$, DOSELAR$, DOSELAR$, REFEREN$, Item%, 0, 1, 0, DEPDESTINO%, Cantidad#)
        Else
           Call MOVISTOK(HOII%, CIXI%, LOTE$, COMPROB$, DOPRILAR$, DOPRILAR$, DOSELAR$, DOSELAR$, Referencia$, Item%, 0, 1, 0, DEPDESTINO%, Cantidad#)
           CONDI$ = "COD_INTE= " & CIXI% & " AND IDENLOTE='" & LOTE$ & "' AND DOPRILAR='" & DOPRILAR$ & "' AND CANTINGRE > 0"
           Call ACTUREGISTRO("MOVISTO", "VERIFICADO", CONDI$, -1, REG&)
        End If
        'Actualiza Deposito en Lotingre
        OFDEST$ = TRIM$(VALOBADA("BALAREQMO", "DESTINO", "NULISPICK = " & Label33))
        Call ACTUREGISTRO("LOTINGRE", "DEPOUBIC", "IDENLOTE='" & LOTE$ & "' and Cod_inte =" & CIXI%, DEPDESTINO%, REG&)
        Call ACTUREGISTRO("LOTINGRE", "UBICFIS", "IDENLOTE='" & LOTE$ & "' and Cod_inte =" & CIXI%, "", REG&)
        Call ACTUREGISTRO("LOTINGRE", "COD_DESTI", "IDENLOTE='" & LOTE$ & "' and Cod_inte =" & CIXI%, OFDEST$, REG&)
      Else
        ' marca el consumo en balareqmo
        CONDI$ = "Cod_Inte = " & CIXI% & " AND IdenLote = '" & LOTE$ & "' AND Kilo_Asig > 0 AND NULISPICK = " & Label33
        Call ACTUREGISTRO("BALAREQMO", "Kilo_Consu", CONDI$, Cantidad#, RAFE&, "NOMESS")
      End If
      Call ACSALOTE(CIXI%, LOTE$, "NOMESS")
      '
     
      Item% = Item% + 1
      'Detalle de Impresion
      CAITAB1% = CAITAB1% + 1
      TETABU1$(1, CAITAB1%) = Codigo$
      TETABU1$(2, CAITAB1%) = DESCRIT$(CIXI%)
      TETABU1$(3, CAITAB1%) = LOTE$
      TETABU1$(4, CAITAB1%) = UNIMEDIDA$
      TETABU1$(5, CAITAB1%) = Cantidad#
      '
    Next AX%
    '
    If TIPO$ = "E" Then
       'Imprime Remito
       FORMU$ = CONTROL$("", "FOREMTRA")
     ElseIf TIPO$ = "I" Then
       'Imprime Vale de Transferencia
       FORMU$ = CONTROL$("", "FORTRSTM")
    End If
    '
    CANCELPRINT% = 0
    '
    SENDBYMAIL$ = EMAI$
    Call PRINTDOC("@" & FORMU$)
    '
    If SENDBYMAIL$ <> "" Then
       SENDBYMAIL$ = ""
       Call PRINTDOC("@" & FORMU$)
    End If
    '
    SENDBYMAIL$ = ""
    '
    COMUNI ("Transferencia Registrada Correctamente")
    Screen.MousePointer = 1
    Unload Me
    Exit Sub
    '
99  Command2.Enabled = True
    Screen.MousePointer = 1
    '
End Sub



Private Sub Command4_Click()
  Command4.Enabled = False
  '
  Call SELECHO("TADEPOSI")
  DEPO$ = VALACT1$("TADEPOSI")
    
  If DEPO$ <> "" Then
      Text3 = DEPO$
  End If
  '
  Command4.Enabled = True
End Sub


Private Sub Form_Activate()
  Timer1.Enabled = True
End Sub

Private Sub Form_Deactivate()
  Timer1.Enabled = False
End Sub

Private Sub Form_Load()
    '
    '\\\OT-08-1071-FL-24/11/08/// Carga la lista
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
End Sub





Private Sub Label33_Change()
    List1.Clear: List2.Clear: Label4 = ""
    Call APERBASDAT("STOCKS")
    '
    If XVALO(Label33) < 1 Then Exit Sub
    '
    'Text3 = "0"      'CONTROL$("STOCKS", "DEPCENPE") 'DEPOSITO DESTINO
    'Label10 = ECOARCH("TADEPOSI", Chr$(XVALO(Text3)))
    'Label7.Caption = TRIM$(VALOBADA("BALAREQMO", "DEPOMOVI", "NULISPICK = " & Label33))
    Label11 = ECOARCH("TADEPOSI", Chr$(XVALO(Label7)))
    '
    DSTX$ = TRIM$(VALOBADA("BALAREQMO", "DESTINO", "NULISPICK = " & Label33))
      DSTX0$ = DSTX$
    DSTY$ = TRIM$(VALOBADA("ORDEFABR", "Cod_Exte", "IDSUBOR = '" & DSTX$ & "'"))
      If DSTY$ <> "" Then DSTX$ = DSTX$ + " - " + DSTY$
    DSTY$ = TRIM$(VALOBADA("ORDEFABR", "REFEREN", "IDSUBOR = '" & DSTX0$ & "'"))
      If DSTY$ <> "" Then DSTX$ = DSTX$ + " - " + DSTY$
    Label4 = DSTX$
    '
    SQLX$ = "SELECT Cod_Inte, SUM(Kilo_Asig) AS SxKilo_Asig, "
    SQLX$ = SQLX$ + " IdenLote, Destino FROM BALAREQMO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NuLisPick = " & Label33
    SQLX$ = SQLX$ + " AND Kilo_Asig > 0 "
    SQLX$ = SQLX$ + " GROUP BY Cod_Inte , IdenLote ,  Destino "
    SQLX$ = SQLX$ + " ORDER BY Cod_Inte ASC"
    '
    Dim BALARETMP As ADODB.Recordset
    Set BALARETMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    JJ& = 0
    With BALARETMP
       Do While Not .EOF
         LOTEE$ = SAFETEXT(!idenlote)
         CONDI$ = "IDENLOTE='" & LOTEE$ & "'"
         DEUBIC$ = VALOBADA("LOTINGRE", "DEPOUBIC", CONDI$)
         If DEUBIC$ = Label7 Then
           CODD% = XVALO(!COD_INTE)
           KILOS = XVALO(!SxKilo_Asig)
           UME$ = UNIMED$(CODD%)
           '
           TTXX$ = AJUSTI$(CODEXT$(CODD%), 16)
           TTXX$ = TTXX$ + AJUSTI$(LOTEE$, 14)
           TTXX$ = TTXX$ + AJUSTI$(UME$, 4)
           KILOS1$ = FORMATSTAN(KILOS, "F11", CODD%)
           TTXX$ = TTXX$ + KILOS1$ 'FORMATNUM$(KILOS1$, "F11." + CNTDC$)
           '
           SALNE = XVALO(VALOBADA("LOTINGRE", "SalNeto", "IdenLote = '" & LOTEE$ & "'"))
           SALNE1$ = FORMATSTAN(SALNE, "F11", CODD%)
           TTXX$ = TTXX$ + SALNE1$ 'FORMATNUM$(SALNE, "F11." + CNTDC$)
           KIBRU = XVALO(VALOBADA("LOTINGRE", "KilBruto", "IdenLote = '" & LOTEE$ & "'"))
           KIBRU1$ = FORMATSTAN(KIBRU, "F11", CODD%)
           TTXX$ = TTXX$ + KIBRU1$ 'FORMATNUM$(KIBRU, "F11." + CNTDC$)
           '
           If SALNE > 0 Then
            List1.AddItem TTXX$
           End If
         End If
         '
       .MoveNext
       Loop
    End With
    '
    Label11 = ECOARCH("TADEPOSI", Chr$(XVALO(Label7)))
    'If TRIM$(List1.List(0)) <> "" Then
       'LOTE$ = TRIM$(Mid$(List1.List(0), 17, 14))
       'Label7 = VALOBADA("LOTINGRE", "DEPOUBIC", "IDENLOTE='" & LOTE$ & "'")
    'End If
    '
End Sub

Private Sub List1_DblClick()
'   Text1 = Left$(List1.TEXT, 28)
    List2.AddItem List1.List(List1.ListIndex)
    List1.RemoveItem (List1.ListIndex)
End Sub

Private Sub List2_DblClick()
    'Text1 = Left$(List2.TEXT, 28)
End Sub

Private Sub Text1_Change()
    'If Len(Text1) < 17 Then
    '   Exit Sub
    'End If
    If InStr(Text1.TEXT, Chr$(10)) < 1 Then
      If InStr(Text1.TEXT, Chr$(13)) < 1 Then
        Exit Sub
      End If
    End If
    '
    TENTRA$ = TRIM$(Text1.TEXT)
    For U& = 1 To Len(TENTRA$)
      If Asc(Mid$(TENTRA$, U&, 1)) = 39 Then
        Mid$(TENTRA$, U&, 1) = Chr$(45)
      End If
    Next U&
    '
    For KKI% = 1 To List2.ListCount
       List2.Selected(KKI% - 1) = False
    Next KKI%
    '
    For KKI% = 1 To List1.ListCount
       TLIS$ = TRIM$(Left$(List1.List(KKI% - 1), 28))
'TTZZ$ = ""
'For KII% = 1 To Len(TLIS$): TTZZ$ = TTZZ$ + Str$(Asc(Mid$(TLIS$, KII%, 1))): Next KII%
'TTZZ$ = TTZZ$ + vbCrLf
'For KII% = 1 To Len(TENTRA$): TTZZ$ = TTZZ$ + Str$(Asc(Mid$(TENTRA$, KII%, 1))): Next KII%
'If TLIS$ = TENTRA$ Then TTZZ$ = TTZZ$ + vbCrLf + "IGUALES"
'MsgBox TTZZ$
       If TLIS$ = TENTRA$ Then
          TTXX$ = List1.List(KKI% - 1)
          List1.RemoveItem (KKI% - 1)
          List2.AddItem TTXX$
          List2.ListIndex = List2.ListCount - 1
          List1.ListIndex = (-1)
          GoTo 20
       End If
    Next KKI%
    '
20  For KKI% = 1 To List2.ListCount
       If TRIM$(Left$(List2.List(KKI% - 1), 28)) = TENTRA$ Then
          List2.ListIndex = KKI% - 1
          List2.Selected(KKI% - 1) = True
          List2.Refresh
          List1.ListIndex = (-1)
          Text1 = ""
          Text1.SetFocus
          Exit Sub
       End If
    Next KKI%
    '
    List1.ListIndex = -1
    List2.ListIndex = -1
    Call MENSERR(24, "Lote no Corresponde a Lista de Picking Activa")
    Text1 = ""
    Text1.SetFocus
    '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    '
    If KeyAscii > 13 Then
      Exit Sub
    End If
    '
    TENTRA$ = CONVERLECBAR(Text1)
    '
    For KKI% = 1 To List2.ListCount
       List2.Selected(KKI% - 1) = False
    Next KKI%
    '
    For KKI% = 1 To List1.ListCount
       TLIS$ = UCase(TRIM$(Left$(List1.List(KKI% - 1), 28)))
'TTZZ$ = "*"
'For KII% = 1 To Len(TLIS$): TTZZ$ = TTZZ$ + Str$(Asc(Mid$(TLIS$, KII%, 1))): Next KII%
'TTZZ$ = TTZZ$ + vbCrLf
'For KII% = 1 To Len(TENTRA$): TTZZ$ = TTZZ$ + Str$(Asc(Mid$(TENTRA$, KII%, 1))): Next KII%
'If TLIS$ = TENTRA$ Then TTZZ$ = TTZZ$ + vbCrLf + "IGUALES"
'MsgBox TTZZ$
       If TLIS$ = TENTRA$ Then
          TTXX$ = List1.List(KKI% - 1)
          List1.RemoveItem (KKI% - 1)
          List2.AddItem TTXX$
          List2.ListIndex = List2.ListCount - 1
          List1.ListIndex = (-1)
          GoTo 20
       End If
    Next KKI%
    '
20  For KKI% = 1 To List2.ListCount
       If TRIM$(Left$(List2.List(KKI% - 1), 28)) = TENTRA$ Then
          List2.ListIndex = KKI% - 1
          List2.Selected(KKI% - 1) = True
          List2.Refresh
          List1.ListIndex = (-1)
          Text1 = ""
          Text1.SetFocus
          Exit Sub
       End If
    Next KKI%
    '
    List1.ListIndex = -1
    List2.ListIndex = -1
    Call MENSERR(24, "Lote no Corresponde a Lista de Picking Activa")
    Text1 = ""
    Text1.SetFocus
    '
End Sub
Private Sub Text2_GotFocus()
 Timer1.Enabled = False
End Sub


Private Sub Text2_LostFocus()
  Timer1.Enabled = True
End Sub

Private Sub Text3_Change()
    DEPOX% = XVALO(Text3)
    DEPOT$ = ECOARCH$("TADEPOSI", Chr$(DEPOX%))
    If Text3 = "0" Then
        Label10 = "Consumo"
      Else
        Label10 = DEPOT$
    End If
End Sub

Private Sub Timer1_Timer()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub


