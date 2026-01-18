VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form RESEPEDI 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Situación de Stocks - Considera Tiempos de Secado."
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   12930
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox TXTCANTTOT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   8040
      Width           =   1335
   End
   Begin VB.TextBox TXTFECHCIERRE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox TXTSECADO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox TXTCANTPEND 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4050
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   4965
      Width           =   1215
   End
   Begin VB.TextBox TXTDISPO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   4965
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11880
      TabIndex        =   26
      Top             =   8040
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Agregar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11760
      TabIndex        =   25
      Top             =   5400
      Width           =   1035
   End
   Begin VB.TextBox TXTCANTIDAD 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10320
      TabIndex        =   23
      Top             =   4965
      Width           =   1215
   End
   Begin VB.TextBox TXTNORFA 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   4965
      Width           =   1215
   End
   Begin VB.TextBox TXTCODACTIVO 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   4965
      Width           =   1815
   End
   Begin VB.TextBox TXTDESCRIACTIVO 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   5400
      Width           =   6970
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12920
      Begin VB.TextBox LENTREGA 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1060
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   960
         Width           =   5760
      End
      Begin VB.TextBox LENTREGA 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   960
         Width           =   5760
      End
      Begin VB.TextBox TXTFECHA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11040
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1320
         Width           =   1680
      End
      Begin VB.TextBox TXTCLIENTE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1060
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox TXTRASOCLI 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2980
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   9735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2775
         TabIndex        =   8
         Top             =   245
         Width           =   195
      End
      Begin VB.TextBox TXTNROPED 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1060
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox TXTDESCRIPCION 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2980
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   9735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Entrega:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   825
      End
      Begin VB.Label Label5 
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label4 
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Pedido:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9600
         TabIndex        =   13
         Top             =   1320
         Width           =   1425
      End
      Begin VB.Label Label3 
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   9
         Top             =   285
         Width           =   825
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   2655
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Seleccione con Doble Click la Fila Con el Item a Seleccionar Reserva"
      Top             =   2040
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2175
      Left            =   0
      TabIndex        =   3
      Top             =   5760
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   12632256
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad Total Reservada:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   0
      Top             =   8070
      Width           =   2265
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Alta:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   8120
      TabIndex        =   35
      Top             =   5430
      Width           =   465
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seco:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   9840
      TabIndex        =   33
      Top             =   5430
      Width           =   585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cant.Pendiente:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   2640
      TabIndex        =   31
      Top             =   4995
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reservas:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   0
      TabIndex        =   29
      Top             =   4680
      Width           =   945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Disp.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   8040
      TabIndex        =   27
      Top             =   4995
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cant.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   9825
      TabIndex        =   24
      Top             =   4995
      Width           =   465
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "O.Fab.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   5880
      TabIndex        =   21
      Top             =   4995
      Width           =   585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   30
      TabIndex        =   20
      Top             =   4995
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle de Pedido:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   30
      TabIndex        =   4
      Top             =   1800
      Width           =   1545
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Visible         =   0   'False
      Begin VB.Menu mnuCopiar 
         Caption         =   "Copiar Celda Seleccionada"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuElimnarFila 
         Caption         =   "Eliminar Fila"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuPrintEtiq 
         Caption         =   "Imprimir Etiqueta"
         Visible         =   0   'False
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuArchivosegundo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuacc 
      Caption         =   "Acceso Directos"
      Begin VB.Menu mnuFikardex 
         Caption         =   "Ficha Kardex Código Activo"
      End
   End
End
Attribute VB_Name = "RESEPEDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GLOBALFRM_strContenidoCelda$
Dim GLOBALFRM_REG&
Dim GLOBALFRM_COL&
Sub PRINTETIQPALLET(NC, RAZOSOC$, CI1%, CODI$, TXTDESCRIACTIVO, CANTI, CANTIMPORM2, NORFA$, DOMI$, LOCALI$, SECADO$)
    FORIPRE$ = ""
    FORIPRE$ = TRIM$(CONTROL("CONFABPE", "ETIQPALL"))
    If FORIPRE$ = "" Then Exit Sub
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Etiqueta de Pallet.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    TIPODOC$ = "ETIQUETA DE RESERVA - PALLET"
    '
    CODPARAM$(1) = "CUEN-C": DOCPARAM$(1) = SAFETEXT$(NC)
    CODPARAM$(2) = "RASO-CLI": DOCPARAM$(2) = TRIM$(RAZOSOC$) 'DESCUENTO RETIRA CLIENTE GABAL
    CODPARAM$(3) = "COD-MAT": DOCPARAM$(3) = TRIM$(CODI$) 'IMPORTE DE DESCTO RETIRA CLIENTE GABAL
    CODPARAM$(4) = "CAN-CONJ": DOCPARAM$(4) = SAFETEXT$(CANTI) 'DESCUENTO POR CANTIDAD
    CODPARAM$(5) = "CAN-MPRI": DOCPARAM$(5) = SAFETEXT$(CANTIMPORM2) 'IMPORTE DE DESCTO POR CANTIDAD
    CODPARAM$(6) = "FECH-VEN": DOCPARAM$(6) = SECADO$
    CODPARAM$(7) = "DES-MAT": DOCPARAM$(7) = TRIM$(TXTDESCRIACTIVO) 'PESO ACUMULADO
    CODPARAM$(8) = "DOMI-EMP": DOCPARAM$(8) = TRIM$(DOMI$)
    CODPARAM$(9) = "LOCA-EMP": DOCPARAM$(9) = LOCALI$
    CODPARAM$(10) = "DESTINO": DOCPARAM$(10) = NORFA$
    '
    CAPARDOC% = 10


    Call PRINTDOC("@" + FORIPRE$) ' PEDIDO DE CLIENTE

   
End Sub

Private Sub MnuPrintEtiq_Click()

   'call imprimir
   textostring$ = TRIM$(STRINGFILA$(Me.MSF1, GLOBALFRM_REG&))
   Erase CADENATEX$
   If textostring$ <> "" Then
        Call TEXTOLOTES$(textostring$, ";")
        On Error Resume Next
        If UBound(CADENATEX$) >= 2 Then
            CANTI = XVALO(CADENATEX$(2))
            NORFA$ = CADENATEX$(3)
            SECADO$ = CADENATEX$(4)
            CANTIX$ = FORMATNUM$(CANTI, "F10.1")
            CODI$ = TXTCODACTIVO
            CI1% = CODINT%(CODI$)
            NC = XVALO(Me.TXTCLIENTE)
            LPREXX% = LIPRCLI%(NC)
            UNME$ = UNIMED$(CI1%)
            CANTIMPORM2 = 1 'ASIGNA POR DEFECTO
            If UNME$ = "M2" And LPREXX% > 0 Then 'SI TIENE LISTA DE PRECIOS Y ES M2
                CANTIMPORM2 = CANTIXM2(CI1%, LPREXX%)
                If CANTIMPORM2 < 0 Then 'SI NO DEVUELVE UN VALOR POR M2
                   CANTIMPORM2 = 1  'LE VUELVE A ASIGNAR LA CANTIDAD ORIGINAL
                End If
                CANTIMPORM2 = CANTIMPORM2 * CANTI
            End If
            Call PRINTETIQPALLET(NC, TXTRASOCLI, CI1%, CODI$, TXTDESCRIACTIVO, CANTI, CANTIMPORM2, NORFA$, LENTREGA(0), LENTREGA(1), SECADO$)
        End If
        On Error GoTo 0
   End If


End Sub

Private Sub Command1_Click()
     '
   Call CARGAPEDIDOS(0, 1)
   NPX& = XVALO(RESP_ZELE_VECT(1))
   TXTNROPED = TRIM$(Str$(NPX&))

'3  Call SELECHO("INDIPEDI/Indice General de Pedidos (Pendientes-Cumplidos-Anulados)")
'   NPX& = Val(VALACT1$("INDIPEDI"))
'   If NPX& < 1 Then GoTo 99
'   '
'   TXTNROPED = TRIM$(Str$(NPX&))
   '
99 Command1.Enabled = True
     '
End Sub

Private Sub Command2_Click()
   If XVALO(Me.TXTCANTIDAD) <= 0 Then
      Call COMUNI("Falta Ingresar Cantidad")
      Call PINTATEXT(Me.TXTCANTIDAD)
      Exit Sub
   End If
   '
   If XVALO(Me.TXTDISPO) <= 0 Then
      Call COMUNI("No Posee Cantidad Disponible")
      Exit Sub
   End If
   '
   If XVALO(Me.TXTCANTIDAD) > XVALO(Me.TXTDISPO) Then
      Call COMUNI("La Cantidad a Reservar es Mayor a la Cantidad Disponible")
      Exit Sub
   End If
   '
   CANTIRES_ENGRILLA = SUMA_CANT_SEGUNDATO(Me.MSF1, Me.TXTNORFA, 2, 1)
   
   If CANTIRES_ENGRILLA + XVALO(TXTDISPO) > XVALO(TXTDISPO) Then
      Call COMUNI("Se Ha Detectado Para la O.Fabricación Seleccionada\ Una Reserva de:  " & CANTIRES_ENGRILLA & " En Grilla \ Verifique")
      Exit Sub
   End If
   '
   If XVALO(SUMA_COLUMNA_SELECCION(Me.MSF1, 1)) + XVALO(Me.TXTCANTIDAD) > XVALO(Me.TXTCANTPEND) Then
      Call COMUNI("La Cantidad a Ingresar es Mayor a la Cantidad Pendiente del Pedido")
      Exit Sub
   End If
   
   MSF1.Rows = MSF1.Rows + 1
   MSF1.TextMatrix(MSF1.Rows - 1, 1) = Me.TXTCANTIDAD
   MSF1.TextMatrix(MSF1.Rows - 1, 2) = Me.TXTNORFA
   MSF1.TextMatrix(MSF1.Rows - 1, 3) = Me.TXTSECADO
   MSF1.TextMatrix(MSF1.Rows - 1, 4) = Me.TXTFECHCIERRE
   
   
   '
   Me.TXTCANTIDAD = ""
   Me.TXTDISPO = ""
   Me.TXTNORFA = ""
   Me.TXTSECADO = ""
   Me.TXTFECHCIERRE = ""
   Me.TXTCANTTOT = FORMATNUM$(XVALO(SUMA_COLUMNA_SELECCION(Me.MSF1, 1)), "F12.1")

   'BUSCAR EN LA GRILLA INFERIOR ESTA OF Y SI ESTA VERIFICAR QUE LA SUMA DE LA GRILLA + LA CANTIDAD DEL TEXT NO SUPERE EL TXTDISPO
   'SI NO AVISAR E INDICAR LA DIFERENCIA.
  ' Command2.Enabled = False
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   CI1% = CODINT%(Me.TXTCODACTIVO)
   CONDI$ = "Cod_Inte = " & CI1%
   CONDI$ = CONDI$ + " AND NroPed = " & XVALO(TXTNROPED)
   CONDI$ = CONDI$ + " AND NroClie = " & XVALO(TXTCLIENTE)
   '
   On Error GoTo ERROR_GUARDAR
   FLEXCONN.BeginTrans
   Call ACTUREGISTRO("RESERPED", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
   '
   sqlx$ = "SELECT * FROM RESERPED "
   sqlx$ = sqlx$ + " WHERE COD_INTE < 1 "
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset
   RS.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   For i& = 1 To MSF1.Rows - 1
      With RS
        .AddNew
        !COD_INTE = CI1%
        !COD_EXTE = TXTCODACTIVO
        !Descrip = TXTDESCRIACTIVO
        !FechRes = CVDAT(HOY(HO$))
        !CANTRES = XVALO(MSF1.TextMatrix(i&, 1))
        !CANTCONS = 0
        !IDENLOTE = MSF1.TextMatrix(i&, 2)
        !FechSecado = MSF1.TextMatrix(i&, 3)
        !FechCierre = MSF1.TextMatrix(i&, 4)
        !NROPED = XVALO(TXTNROPED)
        !NROCLIE = XVALO(TXTCLIENTE)
        !MARBORRA = 0
        .Update
      End With
   Next i&
   
99 Command3.Enabled = True
ERROR_GUARDAR:
    If Err <> 0 Then
       FLEXCONN.RollbackTrans
       On Error Resume Next
       RS.Close
       Set RS = Nothing
       On Error GoTo 0
       CONDI$ = CONDI$ + " AND MARBORRA = 1"
       Call ACTUREGISTRO("RESERPED", "MARBORRA", CONDI$, 0, REGAF&, "NOMESS")
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      RS.Close
      Set RS = Nothing
      CONDI$ = CONDI$ + " AND MARBORRA = 1"
      Call BORRAREGISTROS("RESERPED", CONDI$, REGAF&, "NOMESS")
      Call COMUNI("Reservas Generadas Correctamente")
    End If
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call TXTCODACTIVO_Change
   Command4.Enabled = True
End Sub

Private Sub Form_Load()
    Call CREAR_TABLA_RESERPED
    CAMPOS$ = "CODIGOINTERNO/F0;Código/A20;Descripción/A48;Cant.Pend/f11.1"
    Call CARGA_ENCABEZADO(RESEPEDI.MSF, CAMPOS$, RESEPEDI)
    'CAMPOS$ = "Fecha Reserva/d8;Cant.Reservada/F14.1;Cant.Cons./F14.1;Cant.Pend/f14.1;Ord.Fabricación/A15"
    CAMPOS$ = "Cant.Reservada/F14.1;Fabricación/A15;F.Secado/D8;Alta/A0"
    Call CARGA_ENCABEZADO(RESEPEDI.MSF1, CAMPOS$, RESEPEDI)
End Sub

Private Sub MSF_DblClick()
   REG& = MSF.MouseRow
   COL& = MSF.MouseCol
   If REG& < 1 Then Exit Sub
   If COL& > MSF.Cols - 1 Then Exit Sub
   '
   If Me.TXTCODACTIVO <> MSF.TextMatrix(REG&, 2) Then
     TXTCODACTIVO = ""
     TXTDESCRIACTIVO = ""
     '
     CI1% = XVALO(MSF.TextMatrix(REG&, 1))
     If CI1% > 0 Then
       TXTCODACTIVO = MSF.TextMatrix(REG&, 2)
       TXTDESCRIACTIVO = MSF.TextMatrix(REG&, 3)
       Me.TXTCANTPEND = MSF.TextMatrix(REG&, 4)
       
     End If
   End If
   '
   Call TXTCODACTIVO_DblClick
   Me.TXTCANTTOT = FORMATNUM$(XVALO(SUMA_COLUMNA_SELECCION(Me.MSF1, 1)), "F12.1")

End Sub


Private Sub MSF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
     REG& = MSF1.MouseRow
     COL% = MSF1.MouseCol
     Call MENU_CELDA(MSF1, REG&, COL%)
  End If
End Sub
Sub INVICTRL()
  '
  mnucopiar.Visible = Not (mnucopiar.Visible)
  mnuElimnarFila.Visible = Not (mnuElimnarFila.Visible)
  MnuPrintEtiq.Visible = Not (MnuPrintEtiq.Visible)
  '
End Sub
Private Sub mnucopiar_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub
Private Sub mnuElimnarFila_Click()
    Call cmdEliminaItem(RESEPEDI.MSF1, GLOBALFRM_REG&)
    Me.TXTCANTTOT = FORMATNUM$(XVALO(SUMA_COLUMNA_SELECCION(Me.MSF1, 1)), "F12.1")

End Sub


Sub MENU_CELDA(MSF As MSFlexGrid, REG&, COL%)
        If MSF.Rows <= 1 Then Exit Sub
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        GLOBALFRM_REG& = REG&
        PopupMenu mnuArchivo
        Call INVICTRL
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA
End Sub

Private Sub MSF1_RowColChange()
   Me.TXTCANTTOT = FORMATNUM$(XVALO(SUMA_COLUMNA_SELECCION(Me.MSF1, 1)), "F12.1")
End Sub

Private Sub TXTCANTIDAD_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      On Error Resume Next
      Command2.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub TXTCANTIDAD_LostFocus()
   TXTCANTIDAD = FORMATNUM$(TXTCANTIDAD, "f11.1")
End Sub


Private Sub TXTCODACTIVO_Change()
    MSF1.Rows = 1
    CI1% = CODINT%(TXTCODACTIVO)
    If CI1% > 0 Then
        sqlx$ = "SELECT * FROM RESERPED WITH(NOLOCK)"
        sqlx$ = sqlx$ + " WHERE NROPED = " & XVALO(TXTNROPED)
        sqlx$ = sqlx$ + " AND NROCLIE = " & XVALO(TXTCLIENTE)
        sqlx$ = sqlx$ + " AND COD_INTE = " & CI1%
        sqlx$ = sqlx$ + " AND CANTRES > CANTCONS"
        Set RS = READSET(sqlx$)
        With RS
            Do While Not .EOF
              J& = J& + 1
              MSF1.Rows = J& + 1
              MSF1.TextMatrix(J&, 1) = FORMATNUM$(XVALO(!CANTRES) - XVALO(!CANTCONS), "F10.1")
              MSF1.TextMatrix(J&, 2) = SAFETEXT$(!IDENLOTE)
              MSF1.TextMatrix(J&, 3) = FECHATEX$(CVINT(!FechSecado))
              MSF1.TextMatrix(J&, 4) = FECHATEX$(CVINT(!FechCierre))
              .MoveNext
            Loop
        End With
    End If
End Sub

Private Sub TXTCODACTIVO_DblClick()
   SITUSTO.SELECCCIONADO% = 1
   SITUSTO.TXTCODIGO = "" 'PARA QUE REFRESQUE
   SITUSTO.TXTCODIGO = Me.TXTCODACTIVO
   TXTNORFA = ""
   TXTDISPO = ""
   TXTCANTIDAD = ""
   TXTSECADO = ""
   SITUSTO.TXTOF1$ = "" ' INICIALIZO
   Me.TXTCANTTOT = FORMATNUM$(XVALO(SUMA_COLUMNA_SELECCION(Me.MSF1, 1)), "F12.1")

   SITUSTO.Show 1
   
   Erase CADENATEX$
   Call TEXTOLOTES$(SITUSTO.TXTOF1$, ";")
   On Error Resume Next
   If UBound(CADENATEX$) >= 2 Then
    NORFA$ = CADENATEX$(3)
    CANTDISPONIB = CADENATEX$(8)
    FESEC% = FECHANUM(CADENATEX$(9))
    FECIERRE% = FECHANUM(CADENATEX$(2))

    TXTNORFA = NORFA$
    TXTDISPO = FORMATNUM$(CANTDISPONIB, "F10.1")
    TXTSECADO = FECHATEX$(FESEC%)
    TXTFECHCIERRE = FECHATEX$(FECIERRE%)
   End If
   On Error GoTo 0
   SUGERIDO = XVALO(Me.TXTCANTPEND) - XVALO(TXTCANTTOT)
   If XVALO(SUGERIDO) > XVALO(Me.TXTDISPO) Then SUGERIDO = Me.TXTDISPO
   Me.TXTCANTIDAD = FORMATNUM$(SUGERIDO, "F12.1")
   Call PINTATEXT(Me.TXTCANTIDAD)
   SITUSTO.SELECCCIONADO% = 0
   
End Sub


Private Sub TXTNORFA_Change()
    Command2.Enabled = False
    If TRIM$(TXTNORFA) <> "" Then
       Command2.Enabled = True
    End If
End Sub

Private Sub TXTNROPED_Change()
   MSF.Rows = 1
   MSF1.Rows = 1
   Me.TXTCODACTIVO = ""
   Me.TXTCANTIDAD = ""
   Me.TXTCANTPEND = ""
   Me.TXTDESCRIACTIVO = ""
   Me.TXTDISPO = ""
   Me.TXTNORFA = ""
   '
   NPX& = XVALO(TXTNROPED)
   If NPX& < 1 Then Exit Sub
   NC = cliente_SEGUN_PEDIDO(NPX&)
   '
   'CARGO ENCABEZADO
20 Call APERBASDAT("PEDCLI")
   sqlx$ = "SELECT * FROM PEDENCA WITH(NOLOCK)"
   sqlx$ = sqlx$ + " WHERE NroPed = " & NPX& & " AND NROCLIE = " & NC
   Set RS = READSET(sqlx$)
   With RS
        Me.TXTCLIENTE = XVALO(!NROCLIE)
        TXTRASOCLI = SAFETEXT$(!RaSoClie)
        TXTFECHA = FECHATEX$(CVINT(!FECHEMIS))
        LENTREGA(0) = SAFETEXT$(!DOMIENT)
        LENTREGA(1) = SAFETEXT$(!LOCAENT)
        VUELTA% = 1
   End With
   RS.Close
   Set RS = Nothing
   '
   'CARGO PRIMERA GRILLA CON DETALLE DE PEDIDO ACUMULADO POR ARTICULO
   J& = 0
   sqlx$ = "SELECT  CODIINT,CODIEXT,Descrip, sum(CantEnt) AS CANTENT,sum(CANPED) AS CANTIPED ,NroClie FROM PEDDETA WITH(NOLOCK) "
   sqlx$ = sqlx$ + " WHERE NroPed = " & NPX& & " AND NROCLIE = " & NC
   sqlx$ = sqlx$ + " GROUP BY CODIINT,CODIEXT,DESCRIP, NroClie"
   Set RS = READSET(sqlx$)
   '
   VUELTA% = 0
   With RS
     Do While Not .EOF
      J& = J& + 1
      MSF.Rows = J& + 1
      MSF.TextMatrix(J&, 1) = XVALO(!CODIINT)
      MSF.TextMatrix(J&, 2) = SAFETEXT$(!CODIEXT)
      MSF.TextMatrix(J&, 3) = SAFETEXT$(!Descrip)
      PENDIENTEX = XVALO(!CANTIPED) - XVALO(!CANTENT)
      MSF.TextMatrix(J&, 4) = FORMATNUM$(PENDIENTEX, "F12.1")
      .MoveNext
     Loop
   End With
   '
   RS.Close
   Set RS = Nothing
   Me.TXTCANTTOT = FORMATNUM$(XVALO(SUMA_COLUMNA_SELECCION(Me.MSF1, 1)), "F12.1")

End Sub
