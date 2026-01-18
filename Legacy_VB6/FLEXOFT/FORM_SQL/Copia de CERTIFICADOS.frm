VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CERTIFICADOS 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Databook"
   ClientHeight    =   9045
   ClientLeft      =   75
   ClientTop       =   285
   ClientWidth     =   13980
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9045
   ScaleWidth      =   13980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12180
      Picture         =   "CERTIFICADOS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Imprime Databook Completo"
      Top             =   8410
      Width           =   1560
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   8415
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   13695
      Begin VB.CommandButton Command1 
         Caption         =   "x Familia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   11640
         Picture         =   "CERTIFICADOS.frx":066A
         TabIndex        =   20
         ToolTipText     =   "Selección Por Familia + N.Serie"
         Top             =   840
         Width           =   1005
      End
      Begin VB.CommandButton Command41 
         Height          =   435
         Left            =   11640
         Picture         =   "CERTIFICADOS.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Sube al Nivel Anterior"
         Top             =   360
         Width           =   1005
      End
      Begin VB.CommandButton COMMAND3 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   950
         Left            =   12840
         Picture         =   "CERTIFICADOS.frx":0C7E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Cargar Items"
         Top             =   330
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "O.F."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   950
         Left            =   10920
         TabIndex        =   11
         Top             =   330
         Width           =   525
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A.B.M. de Certificados x Orden de Fabricación "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   10740
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   10130
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   540
            Width           =   420
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   10130
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "Niv."
            Top             =   270
            Width           =   420
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "Número de Serie"
            Top             =   270
            Width           =   2100
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   8040
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   540
            Width           =   2100
         End
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "Descripción"
            Top             =   270
            Width           =   4100
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   2340
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   540
            Width           =   4100
         End
         Begin VB.TextBox Text10 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   540
            Width           =   1630
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   540
            Width           =   2235
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Text            =   "Código"
            Top             =   270
            Width           =   2235
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   6420
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Text            =   "O.Fabricación"
            Top             =   270
            Width           =   1630
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Bindings        =   "CERTIFICADOS.frx":0F88
         Height          =   6975
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   13500
         _ExtentX        =   23813
         _ExtentY        =   12303
         _Version        =   393216
         BackColor       =   16777215
         FocusRect       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   4920
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   9720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "CERTIFICADOS.frx":0F9C
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CERTIFICADOS.frx":11D0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Menu AGREGAR_RUTA 
      Caption         =   "Agregar Ruta"
      Visible         =   0   'False
      Begin VB.Menu mnuedi 
         Caption         =   "Editar"
      End
   End
End
Attribute VB_Name = "CERTIFICADOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" _
                                    Alias "ShellExecuteA" _
                                    (ByVal hwnd As Long, _
                                    ByVal lpOperation As String, _
                                    ByVal lpFile As String, _
                                    ByVal lpParameters As String, _
                                    ByVal lpDirectory As String, _
                                    ByVal nShowCmd As Long) _
                                    As Long
Public CAMBIO%

Sub IMPRIME_CARATULA()

   SQLX$ = " SELECT * FROM CERTI_ESTRUC WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE ORD_Fabric_Gral = '" & Text10 & "'"
   SQLX$ = SQLX$ + " AND N_Serie_Gral = '" & Text1 & "'"
   SQLX$ = SQLX$ + " order by  FAMILIA "
   
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF
        RUT$ = TRIM(SAFETEXT$(!RUT_ARCHIVO))
        COD$ = SAFETEXT$(!COD_EXTE)
        CI1% = CODINT%(COD$)
        LOTE$ = SAFETEXT$(!IDENLOTE)
        NUMSE$ = SAFETEXT$(!NUMEROSERIE)
        COFA$ = SAFETEXT$(!FAMILIA)
        FAMI$ = ECOARCH$("TAFAMIL", COFA$)
        TIT% = XVALO(!TIT_CERTIFICADO)
        HOJAX% = XVALO(!HOJAS)
        '
        If RUT$ = "" Then
           '
           CONDI2$ = "CODINT = " & CI1%
           CONDI2$ = CONDI2$ + " AND IDENLOTE = '" & LOTE$ & "'"
            If NUMSE$ <> "" Then
               CONDI2$ = CONDI2$ + " AND NUMEROSERIE = '" & NUMSE$ & "'"
            Else
               CONDI2$ = CONDI2$ + " AND (NUMEROSERIE = '' OR NUMEROSERIE IS NULL) "
            End If
            '
            INDI% = 0
            SQLX2$ = "SELECT * FROM CERTIFICADOS WITH(NOLOCK)" '**TABLA CERTIFICADOS***QUE SE GRABA EN ABMCERTIFICADOS
            SQLX2$ = SQLX2$ + " WHERE " & CONDI2$
            Set RST = READSET(SQLX2$)
            With RST
              Do While Not .EOF
                RUT$ = SAFETEXT$(!RUT_ARCHIVO)
                COFA$ = SAFETEXT$(!FAMILIA)
                FAMI$ = ECOARCH$("TAFAMIL", COFA$)
                TIT% = XVALO(!TIT_CERTIFICADO)
                TITU$ = ECOARCH$("TITUCERT", MKI$(TIT%))
                HOJAX% = XVALO(!HOJAS)
                
                With CARATULA
                   I% = I% + 1
                   
                   If FAMI_ANT$ <> FAMI$ Then
                     INDI% = INDI% + 1
                   End If
                   .MSF.Rows = I% + 1
                   .MSF.TextMatrix(I%, 1) = FAMI$
                   .MSF.TextMatrix(I%, 1) = TITU$
                      ACUMULA_HOJAX% = ACUMULA_HOJAX% + HOJAINI%
                      HOJAINI% = ACUMULA_HOJAX%
                      HOJAFIN% = HOJAINI% + HOJAINI%
                      INICIA$ = SAFETEXT$(INDI%) & "." & HOJAINI%
                      FINALIZA$ = SAFETEXT$(INDI%) & "." & HOJAFIN%
                      Call REPLA(TX$, INICIA$, 1, 5)
                      Call REPLA(TX$, "-", 6, 1)
                      Call REPLA(TX$, FINALIZA$, 7, 5)
                      
                   .MSF.TextMatrix(I%, 1) = TX$
                   FAMI_ANT$ = FAMI$
                  
                  
                .MoveNext
              Loop
            End With
          End If
       .MoveNext
     Loop
   End With

End Sub

Private Sub Command1_Click()
     Command1.Enabled = False
     '
     If CAMBIO% > 0 Then
      If COMALTER%("Se han Modificado Datos \\Registrar Cambios\Cancelar") = 1 Then
        Call GRABAR_CERTIFICADOS
      End If
     End If
     CAMBIO% = 0
     '
     Call SELECHO("TAFAMIL")
     CODFAMI$ = TRIM$(VALACT1$("TAFAMIL"))
     If CODFAMI$ = "" Then GoTo 99
     '
     ENCA$ = "/" & "O.F.                          N.Serie                              Código                                     Descripción"
     CONDI$ = "PACIENTE =  '" & CODFAMI$ & "' ORDER BY IDSUBOR ,COD_EXTE,NUMEROSERIE,NUMEORD"
     Call CARGALISEL("INDIFANU", "TANUMSE", "IDSUBOR/A12;NUMEROSERIE/A16;COD_EXTE/A18;DESCRIP/A24", CONDI$)
     Call SELECHO("INDIFANU" & ENCA$)
     NORFA$ = TRIM$(VALACT1$("INDIFANU"))
     If NORFA$ = "" Then GoTo 99
     '
     CI1% = CODIFAB%(NORFA$)
     NUSERIE$ = TRIM$(Left$(VALACT2$("INDIFANU"), 16))
     Text1 = NUSERIE$
     Text10 = NORFA$
     '
     Text11 = CODEXT(CI1%)
     '
99   Command1.Enabled = True

End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Command1.Enabled = False
    Command41.Enabled = False
    '
    If CAMBIO% > 0 Then
      If COMALTER%("Se han Modificado Datos \\Registrar Cambios\Cancelar") = 1 Then
        Call GRABAR_CERTIFICADOS
      End If
    End If
    CAMBIO% = 0
    '
    'CARGALISEL MUESTRA LAS ORDENES DE FABRICACION CON NUMEROS DE SERIE QUE NO ESTEN ANULADAS NI CUMPLIDAS
    CONDI$ = "TANUMSE.NUMEROSERIE <> '' AND ORDEFABR.StatuOrf < 8  ORDER BY TANUMSE.IDSUBOR,TANUMSE.COD_EXTE,TANUMSE.NUMEROSERIE,TANUMSE.NUMEORD"
    ENCA$ = "/" & "O.F.                          N.Serie                              Código                                     Descripción"
    Call CARGALISEL("INDIFANU", "TANUMSE+ORDEFABR/TANUMSE.IDSUBOR = ORDEFABR.IDSUBOR", "TANUMSE.IDSUBOR/A12;TANUMSE.NUMEROSERIE/A16;TANUMSE.COD_EXTE/A18;TANUMSE.DESCRIP/A24", CONDI$, "NOMESS")
    Call SELECHO("INDIFANU" & ENCA$)
    NORFA$ = TRIM$(VALACT1$("INDIFANU"))
    If NORFA$ = "" Then GoTo 99
    '
    Screen.MousePointer = 11
    '
    
    CI1% = CODIFAB%(NORFA$)
    If ESTRAZABLE%(CI1%) < 2 Then
       Call MENSERR(24, "Código: " & CODEXT(CI1%) & " No es Trazable por Número de Serie")
       GoTo 99
    End If
    '
    NUSERIE$ = TRIM$(Left$(VALACT2$("INDIFANU"), 16))
    Text1 = NUSERIE$
    Text10 = NORFA$
    Text6 = 0
    '
    Text11 = CODEXT(CI1%)
    '
99  Screen.MousePointer = 1
    Command1.Enabled = True
    Command41.Enabled = True

    Command2.Enabled = True
End Sub

Private Sub Command3_Click()
    COMMAND3.Enabled = False
    Screen.MousePointer = 11
    Call GRABAR_CERTIFICADOS
    CAMBIO% = 0
    Call Text11_Change ' REFERESCO LA GRILLA
    Screen.MousePointer = 1
    COMMAND3.Enabled = True
End Sub

Private Sub Command41_Click()
  Command41.Enabled = False
  
  If IUCOD% > 0 Then
    If CAMBIO% > 0 Then
      If COMALTER%("Se han Modificado Datos \\Registrar Cambios\Cancelar") <> 1 Then
        CAMBIO% = 0
        Call GRABAR_CERTIFICADOS
      End If
    End If

    IUCOD% = IUCOD% - 1
    Text6 = SAFETEXT$(XVALO(Text6) - 1)

    Text11.TEXT = UCOD$(IUCOD% + 1)
  End If
  Command41.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   '
   'IMPRESION DE CARATULA
   Call IMPRIME_CARATULA
   '
   SQLX$ = " SELECT * FROM CERTI_ESTRUC WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE ORD_Fabric_Gral = '" & Text10 & "'"
   SQLX$ = SQLX$ + " AND N_Serie_Gral = '" & Text1 & "'"
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF
        RUT$ = TRIM(SAFETEXT$(!RUT_ARCHIVO))
        COD$ = SAFETEXT$(!COD_EXTE)
        CI1% = CODINT%(COD$)
        LOTE$ = SAFETEXT$(!IDENLOTE)
        NUMSE$ = SAFETEXT$(!NUMEROSERIE)
        '
        If RUT$ <> "" Then
          If EXISTE%(RUT$) > 0 Then
           Call Imprimir(Me.hwnd, RUT$)
          Else
           Call COMUNI("Ruta de Archivo Inexistente\Para Código: " & COD$ & "\Lote-O.F.: " & LOTE$)
          End If
        Else
           '
           CONDI2$ = "CODINT = " & CI1%
           CONDI2$ = CONDI2$ + " AND IDENLOTE = '" & LOTE$ & "'"
            If NUMSE$ <> "" Then
               CONDI2$ = CONDI2$ + " AND NUMEROSERIE = '" & NUMSE$ & "'"
            Else
               CONDI2$ = CONDI2$ + " AND (NUMEROSERIE = '' OR NUMEROSERIE IS NULL) "
            End If
            '
            SQLX2$ = "SELECT * FROM CERTIFICADOS WITH(NOLOCK)" '**TABLA CERTIFICADOS***QUE SE GRABA EN ABMCERTIFICADOS
            SQLX2$ = SQLX2$ + " WHERE " & CONDI2$
            Set RST = READSET(SQLX2$)
            With RST
              Do While Not .EOF
                RUT$ = SAFETEXT$(!RUT_ARCHIVO)
                LOTE$ = SAFETEXT$(!IDENLOTE)
                If EXISTE%(RUT$) > 0 Then
                 Call Imprimir(Me.hwnd, RUT$)
                Else
                 Call COMUNI("Ruta de Archivo Inexistente\Para Código: " & COD$ & "\Lote-O.F.: " & LOTE$)
                End If
                .MoveNext
              Loop
            End With
          End If
       .MoveNext
     Loop
   End With
   
99 Command8.Enabled = True


End Sub

Private Sub Form_Load()

    Call CARGA_ENCABEZADO(MSF, "Código/A16;Descripción/A32;U.M./A3;Uso/F8.1;Lote-O.F./A12;N.Serie/A16;Ruta Archivo/A64;C.T./I5;Hojas/A5", CERTIFICADOS)
    MSF.MergeCells = flexMergeRestrictColumns
    MSF.MergeCol(1) = True

    Call ABRECONEXION
    
    Call CREACAMPO("", "CERTI_ESTRUC", "CODINT", 3, 0, 1)
    Call CREACAMPO("", "CERTI_ESTRUC", "COD_EXTE", 10, 20)
    Call CREACAMPO("", "CERTI_ESTRUC", "DESCRIPCION", 10, 64)
    Call CREACAMPO("", "CERTI_ESTRUC", "STATUS", 3, 0)
    Call CREACAMPO("", "CERTI_ESTRUC", "IDENLOTE", 10, 32, 1)
    Call CREACAMPO("", "CERTI_ESTRUC", "NUMEROSERIE", 10, 32, 1)
    Call CREACAMPO("", "CERTI_ESTRUC", "RUT_ARCHIVO", 12, 0)
    Call CREACAMPO("", "CERTI_ESTRUC", "MARBORRA", 3, 0, 1)
    Call CREACAMPO("", "CERTI_ESTRUC", "NIVEL_ESTRUCT", 3, 0, 1)
    Call CREACAMPO("", "CERTI_ESTRUC", "ORD_Fabric_Gral", 10, 12)
    Call CREACAMPO("", "CERTI_ESTRUC", "Codigo_Niv_Superior", 3, 0, 1)
    Call CREACAMPO("", "CERTI_ESTRUC", "N_Serie_Gral", 10, 32)
    Call CREACAMPO("", "CERTI_ESTRUC", "TIT_CERTIFICADO", 3, 10, 1)
    Call CREACAMPO("", "CERTI_ESTRUC", "Familia", 10, 4)
    Call CREACAMPO("", "CERTI_ESTRUC", "Hojas", 10, 4)
    
    CAMBIO% = 0
End Sub
Sub GRABAR_CERTIFICADOS()
    'CRITERIO DE GRABACION DE DATOS
    '*********************************
    '1.-RECORRO LA GRILLA
    '2.-GRABO EN LA TABLA CERTI_ESTRUC, LOTE Y N.SERIE
    '3.-SI TIENE LOTE Y/O NUMERO DE SERIE LA RUTA DEL ARCHIVO VA EN BLANCO
    '4.-SI NO TIENE LOTE Y/O NUMERO DE SERIE SE GRABA CON LA RUTA.
    '4.2.- SI NO TIENE LOTE NI RUTA NO SE GRABA, EN LA PRESENTACION LA GRILLA SE CARGA A PARTIR DE LA ESTRUCTURA.
    '5.-TENER EN CUENTA QUE SI EN LA GRILLA HAY MAS DE UN REGISTRO CON MISMO CODIGO,LOTE Y N.SERIE EL MISMO SE DEBE GRABAR UNA SOLA VEZ
       ' POR QUE SI NO CUANDO SE LEVANTE LA INFORMACION DE ESTE SE VA A MULTIPLICAR LAS FILAS PARA AQUELLOS QUE TENGAN UN LOTE CON MAS DE
       ' DE UN CERTICADO
       
        Screen.MousePointer = 11
    '
        CODI_SUP$ = Text11
        COSUPE% = CODINT%(CODI_SUP$)
        NORFA$ = TRIM$(Text10)
        NUMERSE$ = TRIM$(Text1)
        NIVEX% = XVALO(Text6)
        '
        'MARCO MARBORRA 1 LOS REGISTROS DE LA GRILLA CON LA INFORMACION
        'QUE ES IGUAL
        CONDI$ = " Codigo_Niv_Superior = " & COSUPE%
        CONDI$ = CONDI$ + " AND ORD_Fabric_Gral = '" & NORFA$ & "'"
        CONDI$ = CONDI$ + " AND N_Serie_Gral = '" & NUMERSE$ & "'"
        CONDI$ = CONDI$ + " AND NIVEL_ESTRUCT = " & NIVEX%
        Call ACTUREGISTRO("CERTI_ESTRUC", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
        '
        SQLX$ = "SELECT * FROM CERTI_ESTRUC "
        SQLX$ = SQLX$ + " WHERE CODINT < 1 "
        
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
25      On Error Resume Next
        RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText     ' SE MODIFICO EL adOpenDynamic POR EL adOpenKeyset
        '
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
        End If
        On Error GoTo 0
        '
        With RS
         For I& = 1 To MSF.Rows - 1
            
            LOTE$ = TRIM$(MSF.TextMatrix(I&, 5))
            RUTARCH$ = MSF.TextMatrix(I&, 7)
            CODI$ = TRIM$(MSF.TextMatrix(I&, 1))
            NUMSER1$ = TRIM$(MSF.TextMatrix(I&, 6))
            CANHOJA% = XVALO(MSF.TextMatrix(I&, 9))
            'PARA QUE NO GUARDE REPETIDOS
            If CODI$ = CODI_ANT$ And LOTE$ = LOTE_ANT$ And NUMSER1$ = NUMSER1_ACT$ And RUTARCH$ = RUTARCH_ANT$ Then GoTo 80
            If LOTE$ = "" And RUTARCH$ = "" And CANHOJA% < 1 Then GoTo 80
            .AddNew
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
            
            'DATOS FIJOS DE ENCABEZADO
            !ORD_Fabric_Gral = NORFA$
            !N_Serie_Gral = NUMERSE$
            'CODIGO SUPERIOR VARIANDO SEGUN POSICION DE ESTRUCTURA
            !Codigo_Niv_Superior = COSUPE%
            '
            'DATOS QUE RECORRE DE LA GRILLA.
            CI1% = CODINT%(CODI$)
            !CODINT = CI1%
            !COD_EXTE = CODI$
            !Descripcion = MSF.TextMatrix(I&, 2)
            !Status = 0
            !IDENLOTE = LOTE$
            !NUMEROSERIE = NUMSER1$
            
               RUTARCH$ = MSF.TextMatrix(I&, 7)
               'SI TIENE LOTE U OF (QUE ESTA EN EL MISMO CAMPO) LA RUTA ES VACIA POR QUE  LA TIENE QUE TRAER DE TABLA CERIFICADOS
               If TRIM$(MSF.TextMatrix(I&, 5)) <> "" Then RUTARCH$ = ""
            !RUT_ARCHIVO = RUTARCH$
            !TIT_CERTIFICADO = XVALO(MSF.TextMatrix(I&, 8))
            !FAMILIA = TRIM$(CODFAMIL$(CI1%))
            !HOJAS = CANHOJA%
            '
            'GUARDO EL DATO
            !NIVEL_ESTRUCT = NIVEX%
            CODI_ANT$ = CODI$
            LOTE_ANT$ = LOTE$
            NUMSER1_ANT$ = NUMSER1$
            RUTARCH_ANT$ = RUTARCH$
            
            .Update
80        Next I&
        End With
        '
        CONDI$ = CONDI$ + " AND MARBORRA = 1"
        Call BORRAREGISTROS("CERTI_ESTRUC", CONDI$, REGAF&, "NOMESS")
        '
        Screen.MousePointer = 1
End Sub





Private Sub MSF_DblClick()
    If CAMBIO% > 0 Then
      If COMALTER%("Se han Modificado Datos \\Registrar Cambios\Cancelar") = 1 Then
        CAMBIO% = 0
        Call GRABAR_CERTIFICADOS
      End If
    End If

   If IUCOD% < 128 Then
     If TRIM$(Text11) <> "" Then
       If Text11 <> UCOD$(IUCOD%) Then
         IUCOD% = IUCOD% + 1: UCOD$(IUCOD%) = Text11.TEXT
                              ICOMPO%(IUCOD%) = MSF.Row
       End If
     End If
   End If
   '
   'PONE EL CODIGO EN LA CAJA DE TEXTO  Y SE EJECUTA EL EVENTO CHANGE
   COD$ = TRIM$(MSF.TextMatrix(MSF.Row, 1))
   If COD$ <> "" Then
     Text11.TEXT = COD$
   End If
   '
   'SUBE UN NIVEL
   Text6 = SAFETEXT$(XVALO(Text6) + 1)
   '
   'SI NO TIENE ESTRUCTURA PRESENTAMENSAJE Y VUELVE PARA ATRAS
   If CAIT% < 1 Then
     Call COMUNI("Código Seleccionado No Posee Estructura")
     Call Command41_Click 'vuelve para atras
   End If
End Sub

Private Sub MSF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    '
    I& = MSF.MouseRow
    j& = MSF.MouseCol
    If I& = 0 Or I& > MSF.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF.Row = I&
    MSF.Col = j&
    On Error Resume Next

    COD$ = TRIM$(MSF.TextMatrix(I&, 1))
    If COD$ = "" Then
        Exit Sub
    Else
        ASIGLOTERUTA.Text11 = COD$
        ASIGLOTERUTA.Text4 = MSF.TextMatrix(I&, 2) 'DESC
        If Left$(MSF.TextMatrix(I&, 5), 2) = "BR" Then
           ASIGLOTERUTA.Text15 = MSF.TextMatrix(I&, 5)
           ASIGLOTERUTA.Text10 = ""

        Else
           ASIGLOTERUTA.Text10 = MSF.TextMatrix(I&, 5)
           ASIGLOTERUTA.Text15 = ""
        End If
        ASIGLOTERUTA.Text3 = MSF.TextMatrix(I&, 6)
        ASIGLOTERUTA.Text1 = MSF.TextMatrix(I&, 7)
        ASIGLOTERUTA.Text5 = ECOARCH$("TITUCERT", MKI$(XVALO(MSF.TextMatrix(I&, 8))))
        ASIGLOTERUTA.ORDEX% = XVALO(MSF.TextMatrix(I&, 8))
        ASIGLOTERUTA.Text6 = XVALO(MSF.TextMatrix(I&, 9))
        ASIGLOTERUTA.Show 1
    End If
    
  End If
End Sub

Sub INVICTRL()
  '
  AGREGAR_RUTA.Visible = Not (AGREGAR_RUTA.Visible)
  '
End Sub


Private Sub Text1_Change()
   
   Call Text11_Change

End Sub

Private Sub Text10_Change()
    Call Text11_Change
End Sub

Private Sub Text11_Change()
'
'CRITERIO DE PRESENTACION DE DATOS
'*********************************
'1.-TRAE LA INFORMACION DE LA ESTRUCTURA
'2.-RECORRE TODA LA TABLA Y ASIGNA LA INFO SI TIENE EN LA TABLA ESTRUC_CERTIFICADOS O SI TIENE LOTES DE LA TABLA CERTIFICADOS
'3.-SI TIENE MAS DE UN REGISTRO GENERA UNA FILA NUEVA.
      
     
      CODD$ = Text11.TEXT
      CI% = CODINT%(CODD$)
      NORFA$ = TRIM$(Text10)
      
      If CI% < 1 Or NORFA$ = "" Then Exit Sub
      '
      Text4 = DESCRIT0$(CI%)

      Screen.MousePointer = 11
      PORMILLAR% = 0
      'SE AGREGO PARAMETRO 'COMPLETO' A LEEEXPLO PARA QUE CUANDO ES 1 INCLUYE TAMBIEN A LOS ANULADOS
      Call LEEEXPLO(CODD$, 1, 1)
'      If CAIT% < 1 Then
'        Exit Sub
'      End If
      
      MSF.Rows = 1
      MSF.BackColor = vbWhite
      '
      CODI_SUP$ = CODD$
      COSUPE% = CI%
      NUMERSE$ = TRIM$(Text1)
      NIVEX% = XVALO(Text6)
      I% = 0
      
      'FOR QUE VA ESCRIBIENDO TODOS LOS COMPONENTES
      For j% = 1 To CAIT%
         'BUSCO RUTA EN TABLA CERTIFICADOS
         CODI$ = TRIM$(Mid$(EXPLOLIN$(j%), 1, 16))
         CI1% = CODINT%(CODI$)
         CONDI$ = " Codigo_Niv_Superior = " & COSUPE%
         CONDI$ = CONDI$ + " AND ORD_Fabric_Gral = '" & NORFA$ & "'"
         CONDI$ = CONDI$ + " AND N_Serie_Gral = '" & NUMERSE$ & "'"
         CONDI$ = CONDI$ + " AND NIVEL_ESTRUCT = " & NIVEX%
         CONDI$ = CONDI$ + " AND CODINT = " & CI1%
         '
         SQLX$ = " SELECT * FROM CERTI_ESTRUC WITH(NOLOCK)"
         SQLX$ = SQLX$ + " WHERE " & CONDI$
         '
         '1.-CONSULTO EN LA TABLA NIVEL_ESTRUCT, DONDE SE GUARDA EL CODIGO,LOTE U OF. Y NUMERO DE SERIE
         '    EN CASO QUE NO SE AGREGE LOTE , LA RUTA ESTA EN ESTA TABLA
         '2.-EN EL CASO QUE EL CODIGO TENGA LOTE TIENE QUE TRAER LA INFORMACION DE LA TABLA CERTIFICADOS
         '   ESTO ES POR QUE EN FRM ABMCERTI09 SE DEFINE PARA UN LOTE TODOS LOS CERTIFICADOS Y SI ESTO
         '   SE GRABARA EN TABLA NIVEL_ESTRUCT NO SE ACTUALIZARIA.
         
             
         Set RS = READSET(SQLX$)
         VUELTA% = 0:
         With RS
            Do While Not .EOF
              LOTE$ = SAFETEXT$(!IDENLOTE)
              NUMSE$ = TRIM$(SAFETEXT$(!NUMEROSERIE))
              CANTIHOJAS% = XVALO(!HOJAS)
              If LOTE$ = "" Then
                ENTRO% = 1
                I% = I% + 1
                MSF.Rows = I% + 1 'CARGA LAS PRIMERAS COLUMNAS,
                MSF.TextMatrix(I%, 1) = Mid$(EXPLOLIN$(j%), 1, 16)
                MSF.TextMatrix(I%, 2) = Mid$(EXPLOLIN$(j%), 17, 28)
                MSF.TextMatrix(I%, 3) = Mid$(EXPLOLIN$(j%), 47, 2)
                MSF.TextMatrix(I%, 4) = Mid$(EXPLOLIN$(j%), 49, 12)

                Ruta$ = SAFETEXT$(!RUT_ARCHIVO)
                MSF.TextMatrix(I%, 5) = ""
                MSF.TextMatrix(I%, 6) = ""
                MSF.TextMatrix(I%, 7) = Ruta$
                MSF.TextMatrix(I%, 8) = SAFETEXT$(!TIT_CERTIFICADO)
                MSF.TextMatrix(I%, 9) = CANTIHOJAS%
             Else
                'CASO DONDE EL CODIGO TIENE LOTE LO TIENE QUE IR A BUSCAR A LA TABLA DE CERTIFICADOS
                'UTILIZO CONSULTA POR QUE PUEDE HABER MAS DE UN REGISTRO.
                CONDI2$ = "CODINT = " & CI1%
                CONDI2$ = CONDI2$ + " AND IDENLOTE = '" & LOTE$ & "'"
                If NUMSE$ <> "" Then
                   CONDI2$ = CONDI2$ + " AND NUMEROSERIE = '" & NUMSE$ & "'"
                Else
                   CONDI2$ = CONDI2$ + " AND ( NUMEROSERIE = '' OR NUMEROSERIE IS NULL) "
                End If
                '
                SQLX2$ = "SELECT * FROM CERTIFICADOS WITH(NOLOCK)" '**TABLA CERTIFICADOS***QUE SE GRABA EN ABMCERTIFICADOS
                SQLX2$ = SQLX2$ + " WHERE " & CONDI2$
                Set RST = READSET(SQLX2$)
                With RST
            
                  Do While Not .EOF
                    ENTRO% = 1
                    Ruta$ = SAFETEXT$(!RUT_ARCHIVO)
                    LOTE$ = SAFETEXT$(!IDENLOTE)
                    NUMSE$ = TRIM$(SAFETEXT$(!NUMEROSERIE))
                    I% = I% + 1
                    MSF.Rows = I% + 1 'CARGA LAS PRIMERAS COLUMNAS,
                    MSF.TextMatrix(I%, 1) = Mid$(EXPLOLIN$(j%), 1, 16)
                    MSF.TextMatrix(I%, 2) = Mid$(EXPLOLIN$(j%), 17, 28)
                    MSF.TextMatrix(I%, 3) = Mid$(EXPLOLIN$(j%), 47, 2)
                    MSF.TextMatrix(I%, 4) = Mid$(EXPLOLIN$(j%), 49, 12)
                    MSF.TextMatrix(I%, 5) = LOTE$
                    MSF.TextMatrix(I%, 6) = NUMSE$
                    MSF.TextMatrix(I%, 7) = Ruta$
                    MSF.TextMatrix(I%, 8) = SAFETEXT$(!TIT_CERTIFICADO)
                    CANTIHO% = XVALO(!HOJAS)
                    If CANTIHO% < 1 Then CANTIHO% = CANTIHOJAS%
                    MSF.TextMatrix(I%, 9) = CANTIHO%
                    .MoveNext
                  Loop
                End With
               End If
80            .MoveNext

            Loop
        End With
            If ENTRO% < 1 Then
                I% = I% + 1
                MSF.Rows = I% + 1
                MSF.TextMatrix(I%, 1) = Mid$(EXPLOLIN$(j%), 1, 16)
                MSF.TextMatrix(I%, 2) = Mid$(EXPLOLIN$(j%), 17, 28)
                MSF.TextMatrix(I%, 3) = Mid$(EXPLOLIN$(j%), 47, 2)
                MSF.TextMatrix(I%, 4) = Mid$(EXPLOLIN$(j%), 49, 12)
            End If
            ENTRO% = 0
     Next j%
     On Error Resume Next
     RST.Close
     Set RST = Nothing
     RS.Close
     Set RS = Nothing
     On Error GoTo 0
     '
     Call ORDENAR_MSF(CERTIFICADOS.MSF, 1, 6)
     '
     INDEXI% = 0
     If IUCOD% >= 0 Then
        If ICOMPO%(IUCOD% + 1) >= 0 Then
          If ICOMPO%(IUCOD% + 1) < MSF.Rows Then
            INDEXI% = ICOMPO%(IUCOD% + 1)
          End If
        End If
     End If
     If INDEXI% > 0 Then
         MSF.Row = INDEXI%
         Call COLOREAR_GRILLA_SOLAFILA(CERTIFICADOS.MSF, &H8000000F, MSF.Row)
         If MSF.TopRow < MSF.Row - 8 Then
           MSF.TopRow = MSF.Row - 8
         End If
     End If
      
     Screen.MousePointer = 1
     Me.Caption = "Databook (Cantidad de Registros: " & I% & ")"
      '
End Sub

Private Sub Text6_Change()
   Call Text11_Change
End Sub
Private Sub Imprimir(ByVal HwndForm As Long, ByVal Ruta As String)
    Dim lRet    As Long
    lRet = ShellExecute(HwndForm, "Print", Ruta, ByVal 0&, 0&, 7)
End Sub

