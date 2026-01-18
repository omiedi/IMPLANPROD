VERSION 5.00
Begin VB.Form OPVARFA04 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Complementos de Facturación"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   4365
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Corregir Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   105
      TabIndex        =   20
      Top             =   1575
      Width           =   3270
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   24
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command9 
            Caption         =   "Inventarios"
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
            Left            =   210
            TabIndex        =   30
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Negativos"
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
            Left            =   2205
            TabIndex        =   29
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Consignación"
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
            Left            =   210
            TabIndex        =   28
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Críticos"
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
            Left            =   2205
            TabIndex        =   27
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Estadísticas"
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
            Left            =   210
            TabIndex        =   26
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Análisis A-B-C"
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
            Left            =   2205
            TabIndex        =   25
            Top             =   1890
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Index           =   0
         Left            =   210
         ScaleHeight     =   585
         ScaleWidth      =   2790
         TabIndex        =   21
         Top             =   420
         Width           =   2850
         Begin VB.CommandButton Command11 
            Height          =   600
            Left            =   2230
            Picture         =   "OPVARFA04.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Corregir Número de Documento"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "OPVARFA04.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Corregir Fecha de Emisión"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Corregir Fecha / Nro. Documento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   670
            TabIndex        =   23
            Top             =   105
            Width           =   1635
         End
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004040&
      Height          =   6000
      Left            =   3465
      ScaleHeight     =   5940
      ScaleWidth      =   2160
      TabIndex        =   15
      Top             =   0
      Width           =   2220
      Begin VB.PictureBox Picture16 
         Height          =   4330
         Left            =   0
         ScaleHeight     =   4275
         ScaleWidth      =   1020
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         Begin VB.Image Image3 
            Height          =   4335
            Left            =   -120
            Picture         =   "OPVARFA04.frx":0294
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command14 
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
         Left            =   105
         Picture         =   "OPVARFA04.frx":262C
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Acceso al Setup de Ventas"
         Top             =   2030
         Width           =   650
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Sol. CAE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   105
         TabIndex        =   34
         ToolTipText     =   "Solicitud de Cae Manual y Reimpresión de Comprobantes"
         Top             =   3255
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Token"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   105
         TabIndex        =   33
         ToolTipText     =   "Genera Archivo TOKEN.TXT para Factura Electrónica."
         Top             =   2730
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command10 
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
         Height          =   640
         Left            =   105
         Picture         =   "OPVARFA04.frx":2A6E
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Re-Imprime Factura"
         Top             =   1370
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3760
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -420
            TabIndex        =   19
            Top             =   0
            Width           =   12000
         End
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
         Left            =   105
         Picture         =   "OPVARFA04.frx":30D8
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   50
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
         Height          =   640
         Left            =   105
         Picture         =   "OPVARFA04.frx":3222
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   700
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "OPVARFA04.frx":352C
         Top             =   3950
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Diferencia de Cambio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   105
      TabIndex        =   8
      Top             =   2940
      Width           =   3270
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Index           =   1
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   12
         Top             =   420
         Width           =   2430
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
            Height          =   600
            Left            =   0
            Picture         =   "OPVARFA04.frx":544E
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Listado Aplicación Diferencias de Cambio"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de Control D.C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   14
            Top             =   105
            Width           =   1425
         End
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Anular"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   3270
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Index           =   2
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   9
         Top             =   420
         Width           =   2430
         Begin VB.CommandButton Command45 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "OPVARFA04.frx":5758
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Anulación de Comprobantes de Facturación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación de Comprobantes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   10
            Top             =   105
            Width           =   1425
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   1
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command16 
            Caption         =   "Análisis A-B-C"
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
            Left            =   2205
            TabIndex        =   7
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Estadísticas"
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
            Left            =   210
            TabIndex        =   6
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Críticos"
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
            Left            =   2205
            TabIndex        =   5
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Consignación"
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
            Left            =   210
            TabIndex        =   4
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Negativos"
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
            Left            =   2205
            TabIndex        =   3
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command23 
            Caption         =   "Inventarios"
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
            Left            =   210
            TabIndex        =   2
            Top             =   420
            Width           =   1590
         End
      End
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu solicitatoken 
         Caption         =   "Solicitar Token"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu caemanual 
         Caption         =   "Solicitud de Cae Manual"
      End
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Config 
      Caption         =   "Configuración"
      Begin VB.Menu confiventas 
         Caption         =   "Configuración de Ventas"
      End
   End
End
Attribute VB_Name = "OPVARFA04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub caemanual_Click()
   Call Command13_Click
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Sub Command10_Click()
    '
    If DERACCE%("REPIMFAC", "Repetir Impresión de Factura") < 2 Then
      Exit Sub
    End If
    '
    FECHALIM% = XVALO(CONTROL$("", "CIEMEVEN")) '\\\OT-06-0212-WAR-06/06/06///
    FECHALU = Int(CDbl(CVDAT(FECHALIM%)))
    '
10  Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then Exit Sub
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NC% & " "
    'SQLX$ = SQLX$ + "AND FechEmisi > " & FECHALU & " "  '\\\OT-06-0212-WAR-06/06/06///
    SQLX$ = SQLX$ + "Order by FechEmisi ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          NCLIE% = !NuProve
          IMPO# = !ValAbsComp
          'If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            If TRIM$(IMPOS$) = "" Then IMPOS$ = AJUSTD$("0.00", 14)
            NPX& = XVALO(Mid$(!CODICOM_LAR, 11, 8))
            If Left$(!CODICOM_LAR, 2) = "CC" Then
              NPX& = XVALO(Mid$(!CODICOM_LAR, 4))
            End If
            JJ& = JJ& + 1
            'Y$ = MKS$(NPX&) + AJUSTI$(!codicom_lar, 18) + IMPOS$
            Y$ = MKS$(JJ&) + AJUSTI$(!CODICOM_LAR, 18) + IMPOS$
            Call GRAREG("!SELCUERE", Y$, JJ&)
          'End If
        .MoveNext
        Loop
      End If
    End With
    Call SETULTREG("!SELCUERE", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUERE") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no tiene Facturas de Venta \ que Puedan Re-Imprimirse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
    Call SELECHO("!SELCUERE")
    If VALACT1$("!SELCUERE") = "" Then GoTo 10
    REGCUEC& = XVALO(VALACT1$("!SELCUERE"))
    If REGCUEC& < 1 Then GoTo 10
    '
    DOCULA$ = Left$(VALACT2$("!SELCUERE"), 18)
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6))
    IMPDOCU$ = TRIM$(Mid$(VALACT2$("!SELCUERE"), 19))
    If IMPDOCU$ = "" Then IMPDOCU$ = " 0.00"
    '
    If Abs(XVALO(IMPDOCU$)) < 0.005 Then
       Call MENSERR(24, "El Documento: " & DOCULA$ & " ,\Aparece como 'Anulado'")
       GoTo 10
    End If
    '
    OKX% = (-1)
    '
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    With CajaBanc
       SQLX$ = "TipoMovim = 'FVEN' "
       SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & DOCULA$ & "' "
       SQLX$ = SQLX$ + "AND NuCLIEN = " & NC% & " "
       SQLX$ = SQLX$ + "AND FechEmisi > " & FECHALU & " "
       .FindFirst SQLX$
       If .NoMatch = True Then
          Call MENSERR(24, "Re-Impresión. Consulte !")
          GoTo 10
       End If
       '
       YACARCLI% = 0
20     If .NoMatch = False Then
          ' cargar datos del cliente
          If YACARCLI% = 0 Then
             NC% = !NuClien
             REPRIFAC.Text1 = NC%
             YACARCLI% = 1
          End If
          ' cargar datos de la factura
          .FindNext SQLX$
          GoTo 20
      End If
    End With
    '
    REPRIFAC.Show 1
    'ACA BUSCO EL REGISTRO EN MOVISTO Y PONGO EN BLANCO EL NºFACT
    With Movisto
       .FindFirst "(DoSecLar = '" & DOCULA$ & "' OR  DoPriLar = '" & DOCULA$ & "') AND Nume_Clie = " & NC% & " "
24     If .NoMatch = False Then
        .Edit
            !DoSecLar = " "
            !DoSecCor = " "
            If !DOPRILAR = DOCULA$ Then     ' CASO DE FACTURACIÓN DIRECTA
               !CANTINGRE = 0
               !CANTSALID = 0
            End If
        .Update
        .FindNext "DoSecLar = '" & DOCULA$ & "' AND Nume_Clie= " & NC% & ""
        GoTo 24
      End If
    End With
    '
    SQLX$ = "SELECT * FROM APLICOBRA "
    SQLX$ = SQLX$ + "WHERE CoDocOr_Lar = '" & DOCULA$ & "' OR CoDocAp_Lar = '" & DOCULA$ & "' "
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NC% & " "
    Set APLICSEL = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    With APLICSEL
31    .FindFirst "NuClien = " & NC% & " "
      If .NoMatch = False Then
         .Delete
         GoTo 31
      End If
    End With
    '
    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
    For KUL& = Len(RFG$) - 3 To 1 Step -4
      RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
      XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
      If TRIM$(Mid$(XXC$, 45, 18)) = TRIM$(DOCULA$) Then
        Call GRAREG("CARBIMON", String$(256, 0), RECABIN&)
      End If
    Next KUL&
    '
    Screen.MousePointer = 1
    Call COMUNI("Anulación Comprobante " + TRIM$(DOCULA$) + " Completa")
    '
    RECUEC04.Text1.TEXT = ""
    RECUEC04.Text1.TEXT = TRIM$(Str$(NC%))
    On Error Resume Next
    Hide
    On Error GoTo 0
    Call FRESHALL
    PROLLAMA$ = Name
    RECUEC04.Show 1
    Call FRESHALL
    '
    '
999 Call CIERRARCH("*.*")
    Call LIBARCH("!SELCUERE")
    Screen.MousePointer = 1
    '
    GoTo 10
    '
End Sub

Sub Command11_Click()
    '//OT-07-0479//
    
    Command11.Enabled = False
    If DERACCE%("CORRFECH", "Corregir Número de Documento") < 2 Then
      Command11.Enabled = True
      Exit Sub
    End If
    '
    Call COMUNI("ATENCIÓN: No Podrá Modificarse el Número\de Comprobantes Correspondientes a Períodos Cerrados.\Para Estos Casos Deberá Re-Abrirse el Período Mensual.")
    '
    FECHALIM% = XVALO(CONTROL$("", "CIEMEVEN")) '\\\OT-06-0212-WAR-06/06/06///
    FECHALU = Int(CDbl(CVDAT(FECHALIM%)))
    '
10  Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then
       Command11.Enabled = True
       Exit Sub
    End If
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FVEN' OR TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NC% & " "
    SQLX$ = SQLX$ + "AND FechEmisi > " & FECHALU & " "  '\\\OT-06-0212-WAR-06/06/06///
    SQLX$ = SQLX$ + "Order by FECHEMISI ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
           NCLIE% = !NuClien
           IMPO# = !ValAbsComp
           FEMISI$ = FECHATEX$(CVINT(!FechEmisi))
           NPX& = XVALO(Mid$(!CODICOM_LAR, 9, 8))
           If Left$(!CODICOM_LAR, 2) = "CC" Then
              NPX& = XVALO(Mid$(!CODICOM_LAR, 4))
           End If
           Y$ = MKS$(NPX&) + AJUSTI$(!CODICOM_LAR, 18) + AJUSTD$(FEMISI$, 14)
           JJ& = JJ& + 1
           Call GRAREG("!SELCUERE", Y$, JJ&)
        .MoveNext
        Loop
      End If
    End With
    Call SETULTREG("!SELCUERE", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUERE") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no tiene Facturas de Venta\que Puedan Corregirse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
    Call SELECHO("!SELCUERE")
    If VALACT1$("!SELCUERE") = "" Then GoTo 10
    REGCUEC& = XVALO(VALACT1$("!SELCUERE"))
    If REGCUEC& < 1 Then GoTo 10
    '
    DOCULA$ = Left$(VALACT2$("!SELCUERE"), 18)
    NUOPANU& = XVALO(Mid$(DOCULA$, 9, 8))
    FEMISI$ = TRIM$(Mid$(VALACT2$("!SELCUERE"), 19))
    '
15  NuevoNumFact$ = InputBox$("Nuevo Número de Documento", "Comprobante " + DOCULA$ + " del " + FEMISI$, NUOPANU&)
    If TRIM$(NuevoNumFact$) = "" Then       ' transaccion cancelada
       Command11.Enabled = True
       Exit Sub
    End If
    '
    'VALIDACION DEL NUMERO DE BOLETA INGRESADO
    '
    NUMECOMPR& = XVALO(NuevoNumFact$)
    If IsNumeric(NuevoNumFact$) = False Then
       Call COMUNI("Número de Factura no Procesable.")
       GoTo 15
    End If
    If NUMECOMPR& = NUOPANU& Then
       Call COMUNI("Número de Factura no Modificado.")
       GoTo 15
    End If
    '
    'FORMATEO DEL NUMER0 INGRESADO
    NUFACT$ = NuevoNumFact$
    While Len(NUFACT$) < 8: NUFACT$ = "0" + NUFACT$: Wend
    NUMEFACT$ = DOCULA$
    Mid(DOCULA$, 9, 8) = NUFACT$
    NUMCOMP$ = DOCULA$
    NUMCOMPCOR$ = Left$(DOCULA$, 3) + Mid$(DOCULA$, 11)
    '
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + " WHERE CodiCom_Lar = '" & NUMCOMP$ & "' "
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, 4)
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         On Error GoTo 0
         Screen.MousePointer = 1
         Call MENSERR(24, "El Número Propuesto ya fue Utilizado.\Transacción Cancelada")
         GoTo 10
      End If
      On Error GoTo 0
    End With

    If COMALTER%("¿Realmente Quiere MODIFICAR el Número \del Documento: " & NUMEFACT$ & " ,\Por el Número de Documento '" + NUMCOMP$ + "'?\\No, Conservar " + NUMEFACT$ + "\Si, Cambiar al " + NUMCOMP$) <> 2 Then GoTo 10
    '
    Screen.MousePointer = 11
    '
    'GRABA EN SADEUPEN
    SQLX$ = "UPDATE SADEUPEN "
    SQLX$ = SQLX$ + "SET Codicom_Lar = '" & NUMCOMP$ & "', "
    SQLX$ = SQLX$ + " Codicom_Cor = '" & NUMCOMPCOR$ & "' "
    SQLX$ = SQLX$ + " WHERE Codicom_Lar = '" & NUMEFACT$ & "' "
    SQLX$ = SQLX$ + "AND NuClien = " & NC%
    CueCorri.Execute (SQLX$)
    '
    'GRABA EN MOVISTO
    '
    SQLX$ = "UPDATE MOVISTO "
    SQLX$ = SQLX$ + "SET DosecLar = '" & NUMCOMP$ & "', "
    SQLX$ = SQLX$ + " DosecCor = '" & NUMCOMPCOR$ & "' "
    SQLX$ = SQLX$ + " WHERE DosecLar = '" & NUMEFACT$ & "' ;"
    Stocks.Execute (SQLX$)
    '
    'GRABA EN APLICOBRA
    '
    SQLX$ = "UPDATE APLICOBRA "
    SQLX$ = SQLX$ + "SET CoDoCor_Lar = '" & NUMCOMP$ & "', "
    SQLX$ = SQLX$ + " CoDoCor_Cor = '" & NUMCOMPCOR$ & "' "
    SQLX$ = SQLX$ + " WHERE CoDoCor_Lar = '" & NUMEFACT$ & "' "
    SQLX$ = SQLX$ + "AND NuClien = " & NC%
    CueCorri.Execute (SQLX$)
    '
    'GRABA EN CAJABANC
    '
    SQLX$ = "UPDATE CAJABANC "
    SQLX$ = SQLX$ + "SET Codicom_Lar = '" & NUMCOMP$ & "', "
    SQLX$ = SQLX$ + " Codicom_Cor = '" & NUMCOMPCOR$ & "', "
    SQLX$ = SQLX$ + " NumeComp = " & NUMECOMPR& & " "
    SQLX$ = SQLX$ + " WHERE Codicom_Lar = '" & NUMEFACT$ & "' "
    SQLX$ = SQLX$ + "AND NuClien = " & NC%
    CueCorri.Execute (SQLX$)
    '
    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
    For KUL& = Len(RFG$) - 3 To 1 Step -4
       RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
       XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
       If TRIM$(Mid$(XXC$, 13, 18)) = TRIM$(NUMEFACT$) Then
         Call REPLA(XXC$, NUMCOMP$, 13, 18)
         Call GRAREG("CARBIMON", XXC$, RECABIN&)
       End If
       If TRIM$(Mid$(XXC$, 45, 18)) = TRIM$(NUMEFACT$) Then
         Call REPLA(XXC$, NUMCOMP$, 45, 18)
         Call GRAREG("CARBIMON", XXC$, RECABIN&)
       End If
    Next KUL&
    '
    Call CIERRARCH("*.*")
    Call LIBARCH("!SELCUERE")
    Screen.MousePointer = 1
    Call COMUNI("Comprobante '" + TRIM$(NUMEFACT$) + "'\Cambiado por Número de Documento '" + NUMCOMP$ + "'.")
    GoTo 10
    '
    '//OT-07-0479//
End Sub

Sub Command12_Click()

    Command12.Enabled = False
    If DERACCE%("GETTOKEN", "Obtener TOKEN Diario del Web Service") >= 2 Then
       CARPEACTUAL$ = App.Path + "\"
       CARPEDESTIN$ = "C:\FLEXOFT\FACELEC\"
       '
       ' CREAR CARPETA DE TRABAJO
       On Error Resume Next
       MkDir "C:\FLEXOFT"
       MkDir "C:\FLEXOFT\FACELEC"
       On Error GoTo 0
       If EXISTE%("C:\FLEXOFT\FACELEC\.") < 1 Then
          Call MENSERR(24, "Imposible crear carpeta 'C:/FLEXOFT/FACELEC'")
          GoTo 99
       End If
       '
       'Eliminar archivos
       If EXISTE%(CARPEDESTIN$ + "WSFE_CLIENTE.EXE") > 0 Then
            FILEKILL (CARPEDESTIN$ + "WSFE_CLIENTE.EXE")
       End If
       '
       If EXISTE%(CARPEDESTIN$ + "WSBFE_CLIENTE.EXE") > 0 Then
            FILEKILL (CARPEDESTIN$ + "WSBFE_CLIENTE.EXE")
       End If
       '
       If EXISTE%(CARPEDESTIN$ + "WSFEX_CLIENTE.EXE") > 0 Then
            FILEKILL (CARPEDESTIN$ + "WSFEX_CLIENTE.EXE")
       End If
       '
       'Generar archivos LINK Y CUIT
       DINTER$ = LOADFILE$(LUDAT$ + "LINK.TXT")
       If DINTER$ = "" Then DINTER$ = "https://wsaa.afip.gov.ar/ws/services/LoginCms?WSDL"
10     DINTER$ = TRIM$(InputBox$("URL", "Direccion de Internet - Web Service Autenticacion", DINTER$))
       If DINTER$ = "" Then GoTo 99
       If Left$(DINTER$, 8) <> "https://" Then
          Call MENSERR(24, "Direccion URL no Válida.")
          GoTo 10
       End If
       Call SAVEFILE(LUDAT$ + "LINK.TXT", DINTER$)
       '
       'Obtiene cuit de la empresa
       N1% = FILEOPEN(LUDAT$ + "SVD147.EXE", 0, 41)
       Dim X As String * 41
       Get #N1%, 6, X
       'VALIDACION
       CUITT$ = TRIM$(DESENCRI$(X))
       VALID% = VALICUIT%(CUITT$)
       If VALID% < 1 Then
          Call COMUNI("C.U.I.T. de Empresa No Válido")
          Call FINAL("")
       End If
       '
       CUITT$ = TRIM$(REPLACAR$(CUITT$, "-", ""))
       Close #N1%
       Call SAVEFILE(LUDAT$ + "CUIT.TXT", CUITT$)
       '
       'Copiar archivos a FACELEC
       Call FileCopy(LUDAT$ + "LINK.TXT", CARPEDESTIN$ + "LINK.TXT")
       Call FileCopy(LUDAT$ + "CUIT.TXT", CARPEDESTIN$ + "CUITEMP.TXT")
       '
       'Validar si existe archivo Copiarlo en FACELEC
       'ObtieneToken.exe
       If EXISTE%(CARPEACTUAL$ + "OBTIENETOKEN.EXE") < 1 Then
          Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPEACTUAL$, "\", "/") & "OBTIENETOKEN.EXE")
          GoTo 99
       End If
       On Error Resume Next
       Call FileCopy(CARPEACTUAL$ + "OBTIENETOKEN.EXE", CARPEDESTIN$ + "OBTIENETOKEN.EXE")
       If Err Then
          On Error GoTo 0
          Call MENSERR(24, "ObtieneToken ya en Ejecución")
          GoTo 99
       End If
       On Error GoTo 0
       '
       'ClienteLoginCms_VB
       If EXISTE%(CARPEACTUAL$ + "ClienteLoginCms_VB.exe") < 1 Then
          Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPEACTUAL$, "\", "/") & "ClienteLoginCms_VB.exe")
          GoTo 99
       End If
       Call FileCopy(CARPEACTUAL$ + "ClienteLoginCms_VB.EXE", CARPEDESTIN$ + "ClienteLoginCms_VB.EXE")
       '
       'micertificado
       If EXISTE%(LUDAT$ + "micertificado.pfx") < 1 Then
          Call COMUNI("Falta Ejecutable: " & REPLACAR$(LUDAT$, "\", "/") + "micertificado.pfx")
          GoTo 99
       End If
       Call FileCopy(LUDAT$ + "micertificado.pfx", CARPEDESTIN$ + "micertificado.pfx")
       '
       'WSFE_Cliente
       On Error Resume Next
       Kill CARPEDESTIN$ + "WSFE_Cliente.EXE"
       On Error GoTo 0
       If XVALO(CONTROL$("FACTURA", "PUVENELE")) < 1 Then GoTo 70
       If EXISTE%(CARPEACTUAL$ + "WSFE_Cliente.exe") < 1 Then
          Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPEACTUAL$, "\", "/") & "WSFE_Cliente.exe")
       End If
       On Error Resume Next
       Call FileCopy(CARPEACTUAL$ + "WSFE_Cliente.EXE", CARPEDESTIN$ + "WSFE_Cliente.EXE")
       On Error GoTo 0
       If EXISTE%(CARPEDESTIN$ + "WSFE_Cliente.EXE") < 1 Then
          COMUNI ("No se encuentra  el archivo WSFE_Cliente.exe")
       End If
       '
70     'WSBFE_Cliente
       On Error Resume Next
       Kill CARPEDESTIN$ + "WSBFE_Cliente.EXE"
       On Error GoTo 0
       If TRIM$(CONTROL$("FACTURA", "PUVENCAP")) = "" Then GoTo 75
       If EXISTE%(CARPEACTUAL$ + "WSBFE_Cliente.exe") < 1 Then
          Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPEACTUAL$, "\", "/") & "WSBFE_Cliente.exe")
       End If
       On Error Resume Next
       Call FileCopy(CARPEACTUAL$ + "WSBFE_Cliente.EXE", CARPEDESTIN$ + "WSBFE_Cliente.EXE")
       On Error GoTo 0
       If EXISTE%(CARPEDESTIN$ + "WSBFE_Cliente.EXE") < 1 Then
          COMUNI ("No se encuentra el archivo WSBFE_Cliente.exe")
       End If
       '
       
75     'WSFEX_Cliente
       On Error Resume Next
       Kill CARPEDESTIN$ + "WSFEX_Cliente.EXE"
       On Error GoTo 0
       If TRIM$(CONTROL$("FACTURA", "PUVEFAEX")) = "" Then GoTo 80
       If EXISTE%(CARPEACTUAL$ + "WSFEX_Cliente.exe") < 1 Then
          Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPEACTUAL$, "\", "/") & "WSFEX_Cliente.exe")
       End If
       On Error Resume Next
       Call FileCopy(CARPEACTUAL$ + "WSFEX_Cliente.EXE", CARPEDESTIN$ + "WSFEX_Cliente.EXE")
       On Error GoTo 0
       If EXISTE%(CARPEDESTIN$ + "WSFEX_Cliente.EXE") < 1 Then
          COMUNI ("No se encuentra el archivo WSFEX_Cliente.exe")
       End If

       
       
80     XXX = Shell(CARPEDESTIN$ + "OBTIENETOKEN.EXE", vbNormalFocus)
       
    End If
    '
99  Command12.Enabled = True
End Sub


Sub Command13_Click()
  TESTCAE04.MODOAUTO = "NO"
  TESTCAE04.Show 1

End Sub

Private Sub Command14_Click()
    Command14.Enabled = False
    OPFACTU00.Show 1
    'Call CONECRUN("FSETUP04/OPFACTUR", "Configuración de Funcionamiento")
    Call FINAL("")
    Command14.Enabled = True
End Sub

Sub Command2_Click()
    '
    Command2.Enabled = False
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ <> "BIMONETA" Then
       If MODOMONE$ <> "MULTIMON" Then
          Call MENSERR(24, "Solo Disponible en Sistemas Bi- O Multi-Monetarios")
          GoTo 99
       End If
    End If
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%) + MKI$(0)
    '
    OBX$ = OBJPLA$("DESHASFECHCLI", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    CLISEL% = CVI(Mid$(OBX$, 5, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    J& = 0
    '
    FIN& = ULTREG&("CARBIMON")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("CARBIMON", U&)
      FEMI% = CVI(Mid$(X$, 1, 2))
      If FEMI% >= DES% Then
        If FEMI% <= HAS% Then
          If CLISEL% < 1 Or CVI(Mid$(X$, 3, 2)) = CLISEL% Then
            '
            ' ENJUAGUE ASEMBLI-GRAMY
            MONEDAPLI% = CVI(Mid$(X$, 63, 2))
            If MONEDAPLI% > ULTREG&("TAMONED") Then
              MONEDAPLI% = MONEDAPLI% Mod 6
              Call REPLA(X$, MKI$(MONEDAPLI%), 63, 2)
              Call GRAREG("CARBIMON", X$, U&)
              If TRIM$(Mid$(X$, 45, 18)) = TRIM$(Mid$(X$, 13, 18)) Then
                Call REPLA(X$, MKI$(MONEDAPLI%), 31, 2)
                Call GRAREG("CARBIMON", X$, U&)
              End If
            End If
            ' FIN ENJUAGUE ASEMBLI-GRAMY
            '
            J& = J& + 1
            Call GRAREG("CARBIMOX", X$, J&)
          End If
        End If
      End If
    Next U&
    Call SETULTREG("CARBIMOX", J&)
    Call CIERRARCH("CARBIMOX")
    '
    Call FILESORT("CARBIMOX", "", 1, 6, "ASC")
    Call CONECRUN("*LISBIMON", "")
    '
99  Screen.MousePointer = 1
    Command2.Enabled = True
End Sub

Sub Command3_Click()
    '
    Command3.Enabled = False
    If DERACCE%("CORRFECH", "Corregir Fecha Emisión Facturas") < 2 Then
      Command3.Enabled = True
      Exit Sub
    End If
    '
    Call COMUNI("ATENCIÓN: No Podrá Modificarse la Fecha de Emisión\de Comprobantes Correspondientes a Períodos Cerrados.\   \Para Estos Casos Deberá Re-Abrirse el Período Mensual.")
    '
    FECHALIM% = XVALO(CONTROL$("", "CIEMEVEN")) '\\\OT-06-0212-WAR-06/06/06///
    FECHALU = Int(CDbl(CVDAT(FECHALIM%)))
    '
10  Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then
       Command3.Enabled = True
       Exit Sub
    End If
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FVEN' OR TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NC% & " "
    SQLX$ = SQLX$ + "AND FechEmisi > " & FECHALU & " "  '\\\OT-06-0212-WAR-06/06/06///
    SQLX$ = SQLX$ + "Order by FECHEMISI ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          NCLIE% = !NuClien
          IMPO# = !ValAbsComp
          'If IMPO# >= 0.005 Then
          '  IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
          '  If TRIM$(IMPOS$) = "" Then IMPOS$ = AJUSTD$("0.00", 14)
            FEMISI$ = FECHATEX$(CVINT(!FechEmisi))
            NPX& = XVALO(Mid$(!CODICOM_LAR, 11, 8))
            If Left$(!CODICOM_LAR, 2) = "CC" Then
              NPX& = XVALO(Mid$(!CODICOM_LAR, 4))
            End If
            Y$ = MKS$(NPX&) + AJUSTI$(!CODICOM_LAR, 18) + AJUSTD$(FEMISI$, 14)
            JJ& = JJ& + 1
            Call GRAREG("!SELCUERE", Y$, JJ&)
          'End If
        .MoveNext
        Loop
      End If
    End With
    Call SETULTREG("!SELCUERE", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUERE") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no tiene Facturas de Venta\que Puedan Corregirse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
    Call SELECHO("!SELCUERE")
    If VALACT1$("!SELCUERE") = "" Then GoTo 10
    REGCUEC& = XVALO(VALACT1$("!SELCUERE"))
    If REGCUEC& < 1 Then GoTo 10
    '
    DOCULA$ = Left$(VALACT2$("!SELCUERE"), 18)
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6))
    FEMISI$ = TRIM$(Mid$(VALACT2$("!SELCUERE"), 19))
    '
15  NUEFE$ = InputBox$("Nueva Fecha de Emisión", "Comprobante " + DOCULA$ + " del " + FEMISI$, FEMISI$)
    If NUEFE$ = "" Then GoTo 10
    If NUEFE$ = FEMISI$ Then
       Call COMUNI("Fecha no Modificada.")
       GoTo 10
    End If
    If FECHANUM(NUEFE$) < 1 Then
       Call MENSERR(24, "Formato de Fecha '" + NUEFE$ + "' no Válido")
       GoTo 15
    End If
    If FECHANUM(NUEFE$) < FECHALIM% Then
       Call MENSERR(24, "Fecha Ingresada no Válida.\Ese Período Mensual Está Cerrado.")
       GoTo 15
    End If
    '
    If COMALTER%("¿Realmente Quiere MODIFICAR la Fecha de Emisión\del Documento: " & DOCULA$ & " ,\del '" + FEMISI$ + "' al '" + NUEFE$ + "' ?\\No, Conservar " + FEMISI$ + "\Si, Cambiar al " + NUEFE$) <> 2 Then GoTo 10
    OKX% = (-1)
    '
    NUEVAFECHA = CVDAT(FECHANUM(NUEFE$))
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    With CajaBanc
       SQLX$ = "(TipoMovim = 'FVEN' OR TipoMovim = 'AJUD') "
       SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & DOCULA$ & "' "
       SQLX$ = SQLX$ + "AND NuCLIEN = " & NC% & " "
       SQLX$ = SQLX$ + "AND FechEmisi > " & FECHALU & " "
       .FindFirst SQLX$
       If .NoMatch = True Then
          Call MENSERR(24, "Transacción Imposible. Consulte !")
          GoTo 10
       End If
       REUTI% = 0
       '
       SQLX$ = "CodiCom_Lar = '" & DOCULA$ & "' "
       SQLX$ = SQLX$ + "AND (NuClien = " & NC% & " OR NuClien = 0) "
       SQLX$ = SQLX$ + "AND FechEmisi > " & FECHALU & " "
19     .FindFirst SQLX$
20     If .NoMatch = False Then
          .Edit
          !FechEmisi = NUEVAFECHA
          !FechConta = NUEVAFECHA
          .Update
          .FindNext SQLX$
          GoTo 20
       End If
    End With
    '
    With SaDeuPen
22    .FindFirst " CodiCom_Lar = '" & DOCULA$ & "' AND NuCLIEN = " & NC% & " "
23     If .NoMatch = False Then
        .Edit
        !FechEmisi = NUEVAFECHA
        .Update
        .FindNext SQLX$
        GoTo 23
      End If
    End With
    'ACA BUSCO EL REGISTRO EN MOVISTO Y PONGO EN BLANCO EL NºFACT
    With Movisto
       .FindFirst " DoSecLar = '" & DOCULA$ & "' AND Nume_Clie = " & NC% & " "
24     If .NoMatch = False Then
        .Edit
        !FechMov = NUEVAFECHA
        .Update
        .FindNext "DoSecLar = '" & DOCULA$ & "' AND Nume_Clie= " & NC% & ""
        GoTo 24
      End If
    End With
    '
    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
    For KUL& = Len(RFG$) - 3 To 1 Step -4
      RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
      XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
      If TRIM$(Mid$(XXC$, 13, 18)) = TRIM$(DOCULA$) Then
        Call REPLA(XXC$, MKI$(CVINT(NUEVAFECHA)), 1, 2)
        Call GRAREG("CARBIMON", XXC$, RECABIN&)
      End If
      If TRIM$(Mid$(XXC$, 45, 18)) = TRIM$(DOCULA$) Then
        Call REPLA(XXC$, MKI$(CVINT(NUEVAFECHA)), 33, 2)
        Call GRAREG("CARBIMON", XXC$, RECABIN&)
      End If
    Next KUL&
    '
    Screen.MousePointer = 1
    Call COMUNI("Comprobante '" + TRIM$(DOCULA$) + "'\Fecha Emisión Cambiada a '" + NUEFE$ + "'.")
    GoTo 10
    '
999 Call CIERRARCH("*.*")
    Call LIBARCH("!SELCUERE")
    Screen.MousePointer = 1
    '
    GoTo 10
    '
End Sub

Sub Command45_Click()
   '
   Command45.Enabled = False

   If DERACCE%("ANULAFAC", "Anulación de Comprobantes de Facturación") = 2 Then
     TTXX$ = "Pueden Anularse Comprobantes Emitidos\o Formularios Inutilizados"
     MODO% = COMALTER%(TTXX$ + "\\Comprobantes Emitidos\Formularios Arruinados\Cancelar")
     '
     If MODO% = 1 Then Call ANUFAC
     '
     If MODO% = 2 Then
         '
         NUMECORR% = 1
         QTRX$ = CONTROL$("NUMEFACT", "NUMECORR")
         If QTRX$ = "INDEPEN" Then
              NUMECORR% = (-1)
           ElseIf QTRX$ = "UNIFICA" Then
              NUMECORR% = (-2)
         End If
         'Si la numeracion no es correlativa no se pueden
         'anular comprobantes.
         If NUMECORR% <= 0 Then
             Call COMUNI("Solo Disponible para Numeración Correlativa.")
             Command20.Enabled = True
             Exit Sub
           Else
             Call ANUFORM
             Call COMUNI("Re-Inicie Ahora la Aplicacion")
             Call FINAL("")
         End If
        '
     End If
   End If
   Command45.Enabled = True
   '
End Sub

Sub ANUFAC()
'ANULACION DE FACTURAS YA EMITIDAS
    '
    If DERACCE%("ANUFAEM", "Anulación de Facturas Emitidas") < 2 Then
      Exit Sub
    End If
    '
    FECHALIM% = XVALO(CONTROL$("", "CIEMEVEN")) '\\\OT-06-0212-WAR-06/06/06///
    FECHALU = Int(CDbl(CVDAT(FECHALIM%)))
    '
10  Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then Exit Sub
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FVEN' OR TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NC% & " "
    SQLX$ = SQLX$ + "AND FechEmisi > " & FECHALU & " "  '\\\OT-06-0212-WAR-06/06/06///
    SQLX$ = SQLX$ + "Order by FechEmisi ASC "
    '
    Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          NCLIE% = !NuProve
          IMPO# = !ValAbsComp
          'If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            If TRIM$(IMPOS$) = "" Then IMPOS$ = AJUSTD$("0.00", 14)
            NPX& = XVALO(Mid$(!CODICOM_LAR, 11, 8))
            If Left$(!CODICOM_LAR, 2) = "CC" Then
              NPX& = XVALO(Mid$(!CODICOM_LAR, 4))
            End If
            JJ& = JJ& + 1
            'Y$ = MKS$(NPX&) + AJUSTI$(!codicom_lar, 18) + IMPOS$
            Y$ = MKS$(JJ&) + AJUSTI$(!CODICOM_LAR, 18) + IMPOS$
            Call GRAREG("!SELCUERE", Y$, JJ&)
          'End If
        .MoveNext
        Loop
      End If
    End With
    Call SETULTREG("!SELCUERE", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUERE") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no tiene Facturas de Venta \ que Puedan Anularse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
    Call SELECHO("!SELCUERE")
    If VALACT1$("!SELCUERE") = "" Then GoTo 10
    REGCUEC& = XVALO(VALACT1$("!SELCUERE"))
    If REGCUEC& < 1 Then GoTo 10
    '
    DOCULA$ = Left$(VALACT2$("!SELCUERE"), 18)
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6))
    IMPDOCU$ = TRIM$(Mid$(VALACT2$("!SELCUERE"), 19))
    If IMPDOCU$ = "" Then IMPDOCU$ = " 0.00"
    '
    If COMALTER%("¿Realmente Quiere ANULAR el Documento: " & DOCULA$ & " ,\con Importe: $" & IMPDOCU$ & "  de la Cuenta " & RASOCLI(NC%) & " ?\\No, Cancelar\Si, Continuar") <> 2 Then GoTo 10
    
    OKX% = (-1)
    '
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    With CajaBanc
       SQLX$ = "(TipoMovim = 'FVEN' OR TipoMovim = 'AJUD') "
       SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & DOCULA$ & "' "
       SQLX$ = SQLX$ + "AND NuCLIEN = " & NC% & " "
       SQLX$ = SQLX$ + "AND FechEmisi > " & FECHALU & " "
       .FindFirst SQLX$
       If .NoMatch = True Then
          Call MENSERR(24, "Anulación Imposible. Consulte !")
          GoTo 10
       End If
       
       REUTI% = 0
       If !OPCIMOVIM = "CI" Then REUTI% = 1
       'SI ES FACTURA CON -1 (SIN OBTENCION DE CAE) ELIMINA REGISTRO
       If XVALO(!NUMECOMP) < 0 Then REUTI% = 1
       '
       If REUTI% = 0 Then
            Call COMUNI("La Transacción RESERVARÁ el Número de Comprobante.")
          Else
            Call COMUNI("La Transacción PERMITIRÁ REUTILIZAR\el Número de Comprobante Anulado.")
       End If
       '
       SQLX$ = "CodiCom_Lar = '" & DOCULA$ & "' "
       SQLX$ = SQLX$ + "AND (NuClien = " & NC% & " OR NuClien = 0) "
       SQLX$ = SQLX$ + "AND FechEmisi > " & FECHALU & " "
19     .FindFirst SQLX$
20     If .NoMatch = False Then
          If REUTI% = 1 Then
              .Delete
              GoTo 19
            Else
             .Edit
             !ValAbsComp = 0
             !IDEBECOMP = 0
             !IHABECOMP = 0
             .Update
             .FindNext SQLX$
             GoTo 20
          End If
      End If
    End With
    '
    With SaDeuPen
22    .FindFirst " CodiCom_Lar = '" & DOCULA$ & "' AND NuCLIEN = " & NC% & " "
23     If .NoMatch = False Then
        .Delete
        GoTo 22
      End If
    End With
    'ACA BUSCO EL REGISTRO EN MOVISTO Y PONGO EN BLANCO EL NºFACT
    With Movisto
       .FindFirst "(DoSecLar = '" & DOCULA$ & "' OR  DoPriLar = '" & DOCULA$ & "') AND Nume_Clie = " & NC% & " "
24     If .NoMatch = False Then
          .Edit
          !DoSecLar = " "
          !DoSecCor = " "
          If !DOPRILAR = DOCULA$ Then     ' CASO DE FACTURACIÓN DIRECTA
             !CANTINGRE = 0
             !CANTSALID = 0
          End If
          .Update
          .FindNext "DoSecLar = '" & DOCULA$ & "' AND Nume_Clie= " & NC% & ""
          GoTo 24
       Else
         'BUSCA SOLO POR EL DOCUMENTO PARA PREVER CASOS DONDE HAY GRUPO DIVISIONAL.
         'LUEGO, SE VERIFICA SI EL CLIENTE DEL MOVISTO TIENE COMO GRUPO DIVISIONAL
         'A NC%, Y EN CASO AFIRMATIVO SE BLANQUEA EL MOVIMIENTO.
         'OT-8-1002-03/11/08
         .FindFirst "(DoSecLar = '" & DOCULA$ & "' OR  DoPriLar = '" & DOCULA$ & "')"
25       If .NoMatch = False Then
            NCLIEN% = XVALO(!Nume_Clie)
            If GRUDIVCLI(NCLIEN%) = NC% Then
               .Edit
               !DoSecLar = " "
               !DoSecCor = " "
               If !DOPRILAR = DOCULA$ Then     ' CASO DE FACTURACIÓN DIRECTA
                  !CANTINGRE = 0
                  !CANTSALID = 0
               End If
               .Update
            End If
            .FindNext "DoSecLar = '" & DOCULA$ & "' AND Nume_Clie= " & NCLIEN% & ""
            GoTo 25
         End If
       End If
       'OT-8-1002-03/11/08
    End With
    '
    SQLX$ = "SELECT * FROM APLICOBRA "
    SQLX$ = SQLX$ + "WHERE CoDocOr_Lar = '" & DOCULA$ & "' OR CoDocAp_Lar = '" & DOCULA$ & "' "
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NC% & " "
    Set APLICSEL = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    With APLICSEL
31    .FindFirst "NuClien = " & NC% & " "
      If .NoMatch = False Then
         .Delete
         GoTo 31
      End If
    End With
    '
    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
    For KUL& = Len(RFG$) - 3 To 1 Step -4
      RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
      XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
      If TRIM$(Mid$(XXC$, 45, 18)) = TRIM$(DOCULA$) Then
        Call GRAREG("CARBIMON", String$(256, 0), RECABIN&)
      End If
    Next KUL&
    '
    Screen.MousePointer = 1
    Call COMUNI("Anulación Comprobante " + TRIM$(DOCULA$) + " Completa")
    '
    RECUEC04.Text1.TEXT = ""
    RECUEC04.Text1.TEXT = TRIM$(Str$(NC%))
    On Error Resume Next
    Hide
    On Error GoTo 0
    Call FRESHALL
    PROLLAMA$ = Name
    RECUEC04.Show 1
    Call FRESHALL
    '
    '
999 Call CIERRARCH("*.*")
    Call LIBARCH("!SELCUERE")
    Screen.MousePointer = 1
    '
    GoTo 10
    '
End Sub

Sub ANUFORM()
'ANULACION DE FACTURA DE FORMULARIO ARRUINADO
    ANUFORMACC.Show 1
End Sub

Private Sub confiventas_Click()
   Call Command14_Click
End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame9.BackColor = &HFCD7B6
      Me.Frame11.BackColor = &HFCD7B6
      
      MARCOBARRA.Top = 3930
      
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
      ' &H8000000F
    End If

End Sub

Private Sub salir_Click()
    Call Command1_Click
End Sub

Private Sub solicitatoken_Click()
    Call Command12_Click
End Sub
