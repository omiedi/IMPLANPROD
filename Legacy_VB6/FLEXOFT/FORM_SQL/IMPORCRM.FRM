VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form IMPORCRM 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procesos de Importacion de Datos"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
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
      Height          =   700
      Left            =   11025
      Picture         =   "IMPORCRM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Des-Hacer Ultima Importacion"
      Top             =   1015
      Width           =   750
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Valores Predefinidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   240
      TabIndex        =   23
      Top             =   1200
      Width           =   4335
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1320
         TabIndex        =   28
         Text            =   " "
         Top             =   360
         Width           =   2835
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
         Height          =   280
         Left            =   1320
         TabIndex        =   25
         Text            =   " "
         Top             =   735
         Width           =   2835
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
         Height          =   280
         Left            =   1320
         TabIndex        =   24
         Text            =   " "
         Top             =   1080
         Width           =   2850
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo:"
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
         Left            =   -240
         TabIndex        =   29
         Top             =   405
         Width           =   1500
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ramo:"
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
         Left            =   -240
         TabIndex        =   27
         Top             =   765
         Width           =   1500
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Especialidad:"
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
         Left            =   -240
         TabIndex        =   26
         Top             =   1125
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Archivo a Importar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2610
      Left            =   120
      TabIndex        =   1
      Top             =   225
      Width           =   4575
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Sobre-Escribir Pre-Existentes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2880
         TabIndex        =   30
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Base de Datos de Clientes y Contactos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   22
         Top             =   360
         Width           =   2295
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
      Height          =   700
      Left            =   11025
      Picture         =   "IMPORCRM.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Cierra y Sale de la Aplicacion"
      Top             =   315
      Width           =   750
   End
   Begin VB.CommandButton Command4 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   11025
      Picture         =   "IMPORCRM.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Realiza la Importacion Final"
      Top             =   1715
      Width           =   750
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   210
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Planilla Origen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      Left            =   120
      TabIndex        =   11
      Top             =   3990
      Width           =   11625
      Begin MSFlexGridLib.MSFlexGrid DbGrid1 
         Height          =   3855
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   6800
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   1
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   11040
      ScaleHeight     =   75
      ScaleWidth      =   675
      TabIndex        =   7
      Top             =   2730
      Width           =   740
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   8
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Formato de Importacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2610
      Left            =   4800
      TabIndex        =   5
      Top             =   240
      Width           =   6100
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   5655
      End
      Begin VB.PictureBox Picture1 
         Height          =   300
         Left            =   240
         ScaleHeight     =   240
         ScaleWidth      =   5595
         TabIndex        =   9
         Top             =   350
         Width           =   5655
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Col. Dato / Variable            Tipo     Rango"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   50
            TabIndex        =   10
            Top             =   0
            Width           =   5655
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Origen de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   105
      TabIndex        =   0
      Top             =   2940
      Width           =   11655
      Begin VB.ComboBox Combo1 
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   8640
         TabIndex        =   20
         Text            =   " "
         Top             =   460
         Width           =   2295
      End
      Begin VB.CommandButton Command7 
         Height          =   615
         Left            =   11040
         Picture         =   "IMPORCRM.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Establece Conexion con Origen de Datos"
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   7035
         TabIndex        =   13
         Top             =   480
         Width           =   7095
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Height          =   285
            Left            =   15
            TabIndex        =   14
            Top             =   0
            Width           =   15000
         End
      End
      Begin VB.CommandButton Command22 
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
         Height          =   315
         Left            =   8040
         TabIndex        =   2
         Top             =   465
         Width           =   220
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Label9"
         Height          =   195
         Left            =   3240
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Planilla:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8400
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   540
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   " "
      Top             =   2200
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.ListBox List3 
      Height          =   1425
      Left            =   4620
      TabIndex        =   15
      Top             =   1050
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   11040
      TabIndex        =   32
      Top             =   2450
      Width           =   735
   End
   Begin VB.Label NARIM 
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "IMPORCRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CANCELIMP%
'
Private objexcel As Excel.Application
Private Libro As Excel.Workbook
Private Hoja As Excel.Worksheet

Sub IMPORASIE()
   '
   Dim LISTA$(256)
   Call ABREPLANCUEN
   '
20 On Error Resume Next
   Set EncaSiGra = Contable.OpenRecordset("ENCABEASIEN", dbOpenDynaset)
   If Err Then
     If COMALTER%("No se Puede Abrir 'ENCABEASIEN'\en Modo Exclusivo.\\Reintentar\Cancelar") <> 2 Then GoTo 20
     GoTo 90
   End If
   '
21 Set DetasiGra = Contable.OpenRecordset("DETAASIEN", dbOpenDynaset)
   If Err Then
     If COMALTER%("No se Puede Abrir 'DETAASIEN'\en Modo Exclusivo.\\Reintentar\Cancelar") <> 2 Then GoTo 21
     GoTo 90
   End If
   '
   JJ& = 0
   FIN& = List4.ListCount
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     NUMERO = XVALO(TRIM$(Mid$(TTXX$, 1, 12)))
     If NUMERO <> NUMEROANT Then
        If JJ& > 0 Then GoSub GRABATODO
        JJ& = 0
        NUMEROANT = NUMERO
        FECHA$ = TRIM$(Mid$(TTXX$, 16, 8))
        NUASI$ = NUPROA$(FECHA$)
        REFEX8$ = TRIM$(Mid$(TTXX$, 32, 48))
        SUMADEBE# = 0
     End If
     JJ& = JJ& + 1
     LISTA(JJ&) = TTXX$
     IDEBE# = XVALO(TRIM$(Mid$(LISTA$(JJ&), 146, 13)))
     SUMADEBE# = SUMADEBE# + ROUND(IDEBE#)
   Next U&
   GoSub GRABATODO
   '
   EncaSiGra.Close
   DetasiGra.Close
   Call COMUNI("Fin de la Importación Definitiva")
   Exit Sub
   '
GRABATODO:
   With EncaSiGra
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
      !NUIDASIEN = NUASI$
      !Refeasien = REFEX8$
          FECHAU% = FECHANUM(FECHA$)
      !FECHASIEN = CVDAT(FECHAU%)
          PPXX% = InStr(NUASI$, "-")
      !OrFeAsien = Val(Mid$(NUASI$, PPXX% + 1))
      !SucuAsien = 0
      !NuDfAsien = 0
      !TIPOASIEN = 0
         !StatAsien = 1  ' activo
      !EditAsien = 1
      !QPasAsien = JJ&
      !DebeAsien = SUMADEBE#
      !Observa = " "
      !Nume_Usuar = USNBR%
      .Update
   End With
   '
   ORPASE% = 0
   For K2% = 1 To JJ&
      COCUE$ = TRIM$(Mid$(LISTA$(K2%), 83, 16))
      DEPASE$ = TRIM$(Mid$(LISTA$(K2%), 100, 46))
      IDEBE# = XVALO(TRIM$(Mid$(LISTA$(K2%), 146, 13)))
         SUMADEBE# = SUMADEBE# + ROUND(IDEBE#)
      IHABE# = XVALO(TRIM$(Mid$(LISTA$(K2%), 162, 13)))
      MACON$ = " "
      CENCOS1$ = " "
      '
      With DetasiGra
         .AddNew   ' YA CONVERTIDO
         On Error Resume Next
         !CodiEmpr = CodiEmp%
         On Error GoTo 0
         !NUIDASIEN = NUASI$
           ORPASE% = ORPASE% + 1
         !OrNuPase = ORPASE%
         !REFEPASE = DEPASE$
           FECHAU% = FECHANUM(FECHA$)
         !FECHASIEN = CVDAT(FECHAU%)
           PPXX% = InStr(NUASI$, "-")
         !OrFeAsien = Val(Mid$(NUASI$, PPXX% + 1))
         !SucuAsien = 0
         !NuDfAsien = 0
         !StatPase = 1
         !CUECOEXT = COCUE$
         !CUECOINT = CUEINT%(COCUE$)
         !IDebePase = IDEBE#
         !IHabePase = IHABE#
         !MARCONSI = MACON$
         !CenCos = CENCOS1$ '\\\OT-06-0244-WAR-16/06/06///
         .Update
      End With
   Next K2%
   '
Return

90 End Sub

Function NUPROA$(FECHA$)
   '
   NUPROA$ = ""
   '
   FEASX = FECHANUM(FECHA$)
   FECHAS$ = FECHATEX$(FEASX)
   FECHAU% = FEASX
   FEASYYYY = CVDAT(FECHAU%)
   NUIDAS$ = Mid$(FECHAS$, 7, 2) + Mid$(FECHAS$, 4, 2) + Mid$(FECHAS$, 1, 2)
   If Left$(NUIDAS$, 1) = "0" Then NUIDAS$ = Mid$(NUIDAS$, 2)
   ORDAS% = 0
   '
   SUCUI% = 0
   Set SELECASIEN = Contable.OpenRecordset("SELECT * FROM EncaBEAsien WHERE (STR$(FECHASIEN) = '" & FEASYYYY & "' and SucuAsien = " & SUCUI% & ") ORDER BY  ORFEASIEN DESC;", dbOpenDynaset, dbReadOnly)
   On Error Resume Next
   SELECASIEN.MoveFirst
   If Err Then GoTo 20
   ORDAS% = SELECASIEN!OrFeAsien
   '
20 On Error GoTo 0
   ORDAS% = ORDAS% + 1
   ORDA$ = TRIM$(Str$(ORDAS%))
   While Len(ORDA$) < 3: ORDA$ = "0" + ORDA$: Wend
   NUPROA$ = NUIDAS$ + "-" + ORDA$
   '
End Function

Sub PREIMPORASIE()
   If DbGrid1.Cols < 7 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     Exit Sub
   End If
   Screen.MousePointer = 11
   DoEvents
   '
   URECORD& = DbGrid1.Rows 'Data1.Recordset.RecordCount
   '
  ' hrow = DbGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   'DbGrid1.RowHeight = 10
   '
   For U& = 1 To URECORD&
    Call TRACE(20, U&, URECORD&)
    TTXX$ = Space$(290)
    On Error Resume Next
    DbGrid1.Row = U& - 1
    If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Imposible Completar Pre-Importación.\Demasiados Registros en Tabla.")
      Exit Sub
    End If
    On Error GoTo 0
    IMPO1# = 0
    IMPO2# = 0
    
    DbGrid1.Col = 3
    'Si la cuenta no existe se va al next U&
    'Si existe sigue normalmente
    Cuentta$ = TRIM$(DbGrid1.TEXT)
    CuenExis% = CUEINT%(Cuentta$)
    If CuenExis% = 0 Then
      Call MENSERR(24, "Cuenta Inexistente")
      GoTo 19
    End If
    '
    DbGrid1.Col = 0                     'ID ASIENTO
    ASIENTTO$ = TRIM$(DbGrid1.TEXT)
    If Len(ASIENTTO$) < 1 Then
      MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Asiento"
      List1.AddItem MERROX$
      GoTo 19
    End If
    Call REPLA(TTXX$, ASIENTTO$, 1, 12)
    '
    DbGrid1.Col = 1
    FECHHA$ = FECHATEX$(CVINT(TRIM(DbGrid1.TEXT)))  ' FECHA
    Call REPLA(TTXX$, FECHHA$, 16, 14)
    '
    DbGrid1.Col = 2                     ' Referencia General
    RefGeneral$ = TRIM$(DbGrid1.TEXT)
    Call REPLA(TTXX$, RefGeneral$, 32, 48)
    '
    DbGrid1.Col = 3                    ' Cuenta
    Cuentta$ = TRIM$(DbGrid1.TEXT)
    If Len(Cuentta$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Cuenta"
       List1.AddItem MERROX$
       GoTo 19
    End If
    Call REPLA(TTXX$, Cuentta$, 83, 16)
    '
    DbGrid1.Col = 4                     ' Referencia Pase
    RefPase$ = TRIM$(DbGrid1.TEXT)
    Call REPLA(TTXX$, RefPase$, 100, 46)
    '
    DbGrid1.Col = 5                     'IMPORTE Debe
    If DbGrid1.TEXT <> "" Then
      IMPO1# = CDbl(TRIM$(DbGrid1.TEXT))
      If IMPO1# >= 0 Then
       IMPODEBE$ = FORMATNUM(IMPO1#, "F13.2")
       Call REPLA(TTXX$, IMPODEBE$, 146, 13)
      End If
    End If
    '
    DbGrid1.Col = 6                     'IMPORTE Haber
    If DbGrid1.TEXT <> "" Then
     If IMPO2# >= 0 Then
      IMPO2# = CDbl(TRIM$(DbGrid1.TEXT))
      IMPOHABER$ = FORMATNUM(IMPO2#, "F13.2")
      Call REPLA(TTXX$, IMPOHABER$, 162, 13)
     End If
    End If
    If (IMPO1# < 0 And IMPO2# < 0) Or (IMPO1# > 0 And IMPO2# > 0) Then
      Call MENSERR(24, "Debe o Haber  Valores Incompatibles")
      GoTo 19
    End If

    List4.AddItem TTXX$
    Label6 = TRIM$(Str$(List4.ListCount))
    If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
    End If
    '
19 Next U&
   '
   'DbGrid1.RowHeight = hrow
   '
End Sub


Sub CONCATENAR_LISCLI()

'  CONCATENA LOS REGISTROS DE UN MISMO CLIENTE EN LISCLI
   If EXISTE%("X:\DAKOT\FLEXOFT\TRAIDOS\PRECONVERTIDOS_TXT\LiscliOrden1_10-01-08.TXT") > 0 Then
      GoTo 20
     Else
      MsgBox "ARCHIVO DE ORIGEN INEXISTENTE"
      Exit Sub
   End If

20 N1% = FILEOPEN%("X:\DAKOT\FLEXOFT\TRAIDOS\PRECONVERTIDOS_TXT\LiscliOrden1_10-01-08.TXT", 1, 128)
   N2% = FILEOPEN%("X:\DAKOT\FLEXOFT\TRAIDOS\PRECONVERTIDOS_TXT\LiscliOrden2_10-01-08.TXT", 2, 128)
   CONT& = 0
   Do While Not EOF(N1%)
      Line Input #N1%, A$
      CONT& = CONT& + 1
      If CONT& = 1 Then
        Call REPLA(TXT$, A$, 0, 135)
        ElseIf CONT& = 2 Then
          Call REPLA(TXT$, A$, 140, 275)
          ElseIf CONT& = 3 Then
            Call REPLA(TXT$, A$, 285, 420)
            CONT& = 0
            Print #N2%, TXT$ 'IMPRIME
      End If
   Loop
End Sub
'
Sub CONCATENAR_LISPRO()

   'CONCATENA LOS REGISTROS DE UN MISMO PROVEEDOR EN LISPRO
   If EXISTE%("X:\DAKOT\FLEXOFT\TRAIDOS\PRECONVERTIDOS_TXT\LisprOOrden1_10-01-08.TXT") > 0 Then
      GoTo 20
    Else
      MsgBox "ARCHIVO DE ORIGEN INEXISTENTE"
      Exit Sub
   End If

20 N1% = FILEOPEN%("X:\DAKOT\FLEXOFT\TRAIDOS\PRECONVERTIDOS_TXT\LisprOOrden1_10-01-08.TXT", 1, 128)
   N2% = FILEOPEN%("X:\DAKOT\FLEXOFT\TRAIDOS\PRECONVERTIDOS_TXT\LisproOrden2_10-01-08.TXT", 2, 128)
   CONT& = 0
   Do While Not EOF(N1%)
      Line Input #N1%, A$
      CONT& = CONT& + 1
      If CONT& = 1 Then
        Call REPLA(TXT$, A$, 0, 100)
        ElseIf CONT& = 2 Then
          Call REPLA(TXT$, A$, 105, 205)
          ElseIf CONT& = 3 Then
            Call REPLA(TXT$, A$, 210, 310)
            CONT& = 0
            Print #N2%, TXT$
      End If
   Loop
End Sub

'
Private Sub Combo1_Click()
   '
   Set objexcel = New Excel.Application
   Set Libro = objexcel.Workbooks.Open(Label1.Caption)
   CANOJ% = Libro.Sheets.Count
   For KKI% = 1 To CANOJ%
      If Libro.Sheets(KKI%).Name = Combo1.TEXT Then
        Set Hoja = Libro.Sheets(KKI%)
        Exit Sub
      End If
   Next KKI%
   Combo1.TEXT = Libro.Sheets(1).Name
   Set Hoja = Libro.Sheets(1)
   DbGrid1.Clear
   DbGrid1.Rows = 1
   '
End Sub

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub
'
Private Sub Command22_Click()
   Command22.Enabled = False
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Archivo Origen de Datos"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   OPENFORMS = DoEvents
   Screen.MousePointer = 11
   Label1 = UCase$(CommonDialog1.filename)
   '
99 Screen.MousePointer = 1
   Command22.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   If List1.ListCount > 0 Then
     Screen.MousePointer = 11
     Printer.Print "Registros no Procesables"
     Printer.Print " "
     For U& = 1 To List1.ListCount
       Printer.Print List1.List(U& - 1)
     Next U&
     Printer.EndDoc
     Screen.MousePointer = 1
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   '
   Command4.Enabled = False
   If EXISTE%(Label1) < 1 Then
     Call MENSERR(24, "Establezca Archivo Origen de Importación")
     Frame5.Visible = False
     GoTo 99
   End If
   '
   TEPRUE = ""
   On Error Resume Next
   URECORD& = DbGrid1.Rows   '   Data1.Recordset.RecordCount
   If Err Then
     Call MENSERR(24, "Error de Acceso a Archivo\Origen de Importación")
     GoTo 99
   End If
   '
   DBHANT = DbGrid1.Height
   Call RIMPORCRM
   DbGrid1.Height = DBHANT
   '
99 Screen.MousePointer = 1
   Command4.Enabled = True
End Sub

Private Sub IMPORESPECIFI()
   If DbGrid1.Cols < 2 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     Exit Sub
   End If
   '
   Screen.MousePointer = 11
   URECORD& = DbGrid1.Rows 'Data1.Recordset.RecordCount
   ATX% = 0
   '
   Call ABRECONEXION
   ReemplazarSiempre% = 0
   '
   'With Data1
   '  .Recordset.MoveFirst
   '  U& = 0
   U& = 1
   '  Do While Not .Recordset.EOF
   Do While U& < URECORD&
        U& = U& + 1
        DbGrid1.Row = U& - 1
        Call TRACE(20, U&, URECORD&)
        TTXX$ = Space$(356)
        CODEX$ = ""
        On Error Resume Next
        DbGrid1.Col = 0
        CODEX$ = TRIM$(DbGrid1.TEXT) ' .Recordset.Fields(0)   ' TRIM$(DBGrid1.TEXT)
        DbGrid1.Col = 1
        ESPEFI$ = TRIM$(DbGrid1.TEXT) '.Recordset.Fields(1)   ' TRIM$(DBGrid1.TEXT)
        COALT1$ = ""
        On Error GoTo 0
        'Registra especificacion
        CONDI$ = "CODIGO='" & CODEX$ & "'"
        If CantRegistros("MASTER", CONDI$) < 1 Then GoTo 10
        If TRIM$(VALOBADA("MASTER", "ESPECIFICACION", CONDI$)) <> "" Then
           '
           If ReemplazarSiempre% <> 1 Then
              OPX% = COMALTER%("Codigo '" & CODEX$ & "'  ya tiene Especificaciones\¿Desea Sobreesribir?\\Conservar\Reemplazar\Reemplazar Siempre")
              If OPX% <= 1 Then GoTo 10                  ' Conservar
              If OPX% = 3 Then ReemplazarSiempre% = 1   ' Reemplazar Siempre
           End If
           '
           If OPX% <> 1 Then
              Call ACTUREGISTRO("MASTER", "ESPECIFICACION", "CODIGO='" & CODEX & "'", ESPEFI$, REG&)
           End If
          Else ' Si No tiene Especificacion se asigna esta
           Call ACTUREGISTRO("MASTER", "ESPECIFICACION", "CODIGO='" & CODEX & "'", ESPEFI$, REG&)
        End If
10 'Pasar al siguiente registro
        '.Recordset.MoveNext
     Loop
   'End With
   Call COMUNI("Importacion Compledata")
End Sub

Private Sub Command5_Click()
   Frame5.Visible = False
   DoEvents
End Sub


Sub PREIMPORMEDI()
   '
   Screen.MousePointer = 11
   URECORD& = DbGrid1.Rows 'Data1.Recordset.RecordCount
   ATX% = 0: CODEXA$ = ""
   '
   'hrow = DbGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   'DbGrid1.RowHeight = 10
   '
   HUBOAMAS% = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(1024)
     DbGrid1.Row = U& - 1
     DbGrid1.Col = 0
5    CODEX$ = TRIM$(DbGrid1.TEXT): If CODEX$ = "" Then CODEX$ = CODEXA$
     If CODINT%(CODEX$) < 1 Then
       CODEX$ = Left$(CODEX$, 15)
       If CODINT%(CODEX$) < 1 Then
         TTXX1$ = "Codigo de Conjunto '" + TRIM$(CODEX$) + "' Inexistente en B.Datos."
         TTXX1$ = TTXX1$ + "¿ Desea Operación de Alta Interactiva ?\\Si, Darlo de Alta\No, Continuar"
         If COMALTER%(TTXX1$) = 1 Then
           VDEF$ = String$(128, 0)
           Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
           Call REPLA(VDEF$, "U.", 109, 2)
           Call REPLA(VDEF$, Chr$(1), 93, 1)
           Call REPLA(VDEF$, CODEX$, 77, 16)
           DEDEX$ = ""
           Call REPLA(VDEF$, DEDEX$, 5, 64)
           Call AMAS_GES04.ALTACODIGO(VDEF$, OKX%)
           HUBOAMAS% = 1
           If OKX% = 1 Then GoTo 5
         End If
       End If
     End If
     '
     If CODEX$ = "" Or Len(CODEX$) < 1 Or Len(CODEX$) > 24 Or CODINT%(CODEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Conjunto '" + CODEX$ + "' Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
     If CODEX$ = CODEXA$ Then GoTo 10 ' EL MISMO CODIGO
     '
     CODEXA$ = CODEX$
     Call LEEEXPLO(CODEX$, 1)
     If CAIT% > 0 And ATX% = 2 Then GoTo 19
     '
     'ATX% = 3
     If CAIT% > 0 Then
         ATX% = COMALTER%("Código '" + CODEX$ + "' con Estructura en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
       Else
         GoTo 12
     End If
     '
10   If ATX% <= 2 And CAIT% > 0 Then GoTo 19 ' SALTEAR
     '
12   Call REPLA(TTXX$, CODEX$, 1, 24)
     '
     COEFI = 1
     ROWAC = DbGrid1.Row
     For JJ& = DbGrid1.Row - 1 To 0 Step -1
        DbGrid1.Row = JJ&
        DbGrid1.Col = 1
        If DbGrid1.TEXT = CODEX$ Then
          DbGrid1.Col = 2
          COEFI = XVALO(DbGrid1.TEXT)
          If COEFI < 1 Then COEFI = 1
          GoTo 125
        End If
     Next JJ&
     '
125  DbGrid1.Row = ROWAC
     '
     DbGrid1.Col = 1
13   COELE$ = Mid$(TRIM$(DbGrid1.TEXT), 1, 24)
     If CODINT%(COELE$) < 1 Then
       COELE$ = Left$(COELE$, 15)
       If CODINT%(COELE) < 1 Then
         TTXX1$ = "Codigo de Elemento '" + TRIM$(COELE$) + "' Inexistente en B.Datos."
         TTXX1$ = TTXX1$ + "¿ Desea Operación de Alta Interactiva ?\\Si, Darlo de Alta\No, Continuar"
         If COMALTER%(TTXX1$) = 1 Then
           VDEF$ = String$(128, 0)
           Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
           Call REPLA(VDEF$, "U.", 109, 2)
           Call REPLA(VDEF$, Chr$(1), 93, 1)
           Call REPLA(VDEF$, COELE$, 77, 16)
           DEDEX$ = ""
           Call REPLA(VDEF$, DEDEX$, 5, 64)
           Call AMAS_GES04.ALTACODIGO(VDEF$, OKX%)
           HUBOAMAS% = 1
           If OKX% = 1 Then GoTo 13
         End If
       End If
     End If
     '
     If COELE$ = "" Or Len(COELE$) < 1 Or Len(COELE$) > 24 Or COELE$ = CODEX$ Or CODINT%(COELE$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Elemento '" + COELE$ + "' Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, COELE$, 27, 24)
     Call REPLA(TTXX$, DESCRIT0$(CODINT%(COELE$)), 53, 32)
     '
     TTXX1$ = Left$(TTXX$, 52)
     For JJ1& = 1 To List4.ListCount
        If Left$(List4.List(JJ1& - 1), 52) = TTXX1$ Then GoTo 19 ' YA ESTABA
     Next JJ1&
     '
     DbGrid1.Col = 2                     ' USO UNITARIO
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DbGrid1.TEXT) / COEFI
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.4")
     Call REPLA(TTXX$, VALO$, 85, 12)
     '
     DbGrid1.Col = 3                     ' UNIDADES
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DbGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.4")
     Call REPLA(TTXX$, VALO$, 98, 12)
     '
     DbGrid1.Col = 4                     ' LARGO DE CORTE
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DbGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.3")
     Call REPLA(TTXX$, VALO$, 111, 12)
     '
     DbGrid1.Col = 5                     ' PLAZO DISPOSICION
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DbGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "I4")
     Call REPLA(TTXX$, VALO$, 124, 4)
     '
     DbGrid1.Col = 6                     ' REFERENCIA
     On Error Resume Next
     VALO$ = DbGrid1.TEXT
     On Error GoTo 0
     Call REPLA(TTXX$, VALO$, 131, 890)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     '
19 Next U&
   '
   'DbGrid1.RowHeight = hrow
   '
End Sub

Private Sub Command7_Click()
   Screen.MousePointer = 11
   CANFIL& = 10000   ' Hoja.Rows
   CANCOL& = List2.ListCount    ' Hoja.Columns
   JK& = 0
   For I& = 1 To CANFIL&
     YAGRELI% = 0
     For J& = 1 To CANCOL&
       On Error Resume Next
       VALOR = Hoja.Cells(I&, J&)
       If Err Then
          On Error GoTo 0
          Call MENSERR(24, "Error de Acceso a la Planilla")
          Screen.MousePointer = 1
          Exit Sub
       End If
       On Error GoTo 0
       '
       If TRIM$(CStr(VALOR)) <> "" Then
         If YAGRELI% < 1 Then
           YAGRELI% = 1
           JK& = JK& + 1
         End If
         If JK& > DbGrid1.Rows Then DbGrid1.Rows = DbGrid1.Rows + 1
         DbGrid1.Row = JK& - 1
       End If
         If J& > DbGrid1.Cols Then DbGrid1.Cols = DbGrid1.Cols + 1
       DbGrid1.Col = J& - 1
       DbGrid1.TEXT = VALOR
       '
       Label9.Font = DbGrid1.Font
       Label9.FontSize = DbGrid1.CellFontSize
       Label9.FontBold = DbGrid1.CellFontBold
       Label9.Caption = CStr(VALOR)
       If Label9.Width > DbGrid1.ColWidth(J& - 1) Then
         DbGrid1.ColWidth(J& - 1) = Label9.Width
       End If
     Next J&
     If I& > JK& + 256 Then
        GoTo 99
     End If
   Next I&
   '
99 Screen.MousePointer = 1
   Call COMUNI("Origen de Datos Establecido.")
   '
End Sub
'
Private Sub Form_Load()
   Call CENTRAFORM(Me)
   Call CARGAFORMATO
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
End Sub
'
Private Sub HScroll1_Change()
  On Error Resume Next
  List4.Left = (-1) * 10 * HScroll1.Value
  'Label8.Left = (-1) * 10 * HScroll1.Value + 40
  Picture7.Left = (-1) * 10 * HScroll1.Value ' + 40
  Picture2.Left = -20
  Picture2.Width = HScroll1.Width
  Ab1 = HScroll1.Value
  Ab1 = 10 * Ab1 + 40
  Ab2 = Abs(Picture7.Left)
  If Ab1 > Ab2 + 80 Then
    Picture2.Left = (-1) * (Ab1 - Ab2)
    Picture2.Width = HScroll1.Width - Picture2.Left
  End If
  On Error GoTo 0
End Sub

Private Sub HScroll1_Scroll()
  On Error Resume Next
  List4.Left = (-1) * 10 * HScroll1.Value
  'Label8.Left = (-1) * 10 * HScroll1.Value + 40
  Picture7.Left = (-1) * 10 * HScroll1.Value ' + 40
  Picture2.Left = -20
  Picture2.Width = HScroll1.Width
  Ab1 = HScroll1.Value
  Ab1 = 10 * Ab1 + 40
  Ab2 = Abs(Picture7.Left)
  If Ab1 > Ab2 + 80 Then
    Picture2.Left = (-1) * (Ab1 - Ab2)
    Picture2.Width = HScroll1.Width - Picture2.Left
  End If
  On Error GoTo 0
End Sub



Private Sub Label1_Change()
  DbGrid1.Clear
  DbGrid1.Rows = 1
  Combo1.Clear
  Combo1.TEXT = ""
  '
  If Label1.Caption <> "" Then
     Dim objexcel As Excel.Application
     Set objexcel = New Excel.Application
     Set Libro = objexcel.Workbooks.Open(Label1.Caption)
     '
     CANOJ% = Libro.Sheets.Count
     For KKI% = 1 To CANOJ%
        Combo1.AddItem Libro.Sheets(KKI%).Name
        If KKI% = 1 Then
           Combo1.TEXT = Libro.Sheets(KKI%).Name
        End If
     Next KKI%
     Call Combo1_Click
  End If
End Sub

Sub CARGAFORMATO()
   List2.Clear
   Label1 = ""
   Combo1.Clear
   'DbGrid1.ClearFields
   NARIM = TRIM$(Str$(Index))
     List2.AddItem "  0  Nombre de la Empresa       String   <= 62 car"
     List2.AddItem "  1  Domicilio                  String   <= 64 car"
     List2.AddItem "  2  Código Postal              String   <= 16 car"
     List2.AddItem "  3  Localidad                  String   <= 48 car"
     List2.AddItem "  4  Codigo de Provincia        String   <= 1 car"
     List2.AddItem "  5  Pais                       String   <= 64 car"
     List2.AddItem "  6  Teléfono                   String   <= 24 car"
     List2.AddItem "  7  Fax                        String   <= 24 car"
     List2.AddItem "  8  E-mail                     String   <= 36 car"
     List2.AddItem "  9  WebSite                    String   <= 128 car"
     List2.AddItem " 10  Número de Vendedor         Num      0 - 255"
     List2.AddItem " 11  Persona de Contacto 1      String   <= 48 car"
     List2.AddItem " 12  Cargo del Contacto a       String   <= 48 car"
     List2.AddItem " 13  Mail del Contacto 1        String   <= 48 car"
     List2.AddItem " 14  Telefono del Contacto 1    String   <= 24 car"
     List2.AddItem " 15  Persona de Contacto 2      String   <= 48 car"
     List2.AddItem " 16  Cargo del Contacto 2       String   <= 48 car"
     List2.AddItem " 17  Mail del Contacto 2        String   <= 48 car"
     List2.AddItem " 18  Telefono del Contacto 2    String   <= 24 car"
     List2.AddItem " 19  Observaciones              String   <=999 car"
     '
     ENCABE$ = " Cta. Razón Social                                    Domicilio                                       " & _
     "C.Postal    Localidad                       Teléfono                Fax                     Contacto                CUIT                     Gp.  Cod.Cta       V.  Pr  PI  CP     Lim.Cre  " & _
     "Descuento                 Lugar de Entrega                                Localidad                                       N-Tra   Transportista                            Gp.Div.  " & _
     "IB  e-mail                               R-Emb"

End Sub

Private Sub VScroll1_Change()
   Picture5.Top = (-10) * VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
   Picture5.Top = (-10) * VScroll1.Value
End Sub

Private Sub VScroll2_Change()
   On Error Resume Next
   List4.TopIndex = VScroll2.Value
   On Error GoTo 0
End Sub

Private Sub VScroll2_Scroll()
   On Error Resume Next
   List4.TopIndex = VScroll2.Value
   On Error GoTo 0
End Sub

Sub PREIMPORMAS()
   '
   If DbGrid1.Cols < 17 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = DbGrid1.Rows 'Data1.Recordset.RecordCount
   ATX% = 0
   '
   'With Data1
   '  .Recordset.MoveFirst
   '  U& = 0
   U& = 1
   '  Do While Not .Recordset.EOF
   Do While U& < URECORD&
        U& = U& + 1
        DbGrid1.Row = U& - 1
        Call TRACE(20, U&, URECORD&)
        TTXX$ = Space$(356)
        CODEX$ = ""
        On Error Resume Next
        DbGrid1.Col = 0
        CODEX$ = TRIM$(DbGrid1.TEXT) '.Recordset.Fields(0)   ' TRIM$(DBGrid1.TEXT)
        COALT1$ = ""
        On Error GoTo 0
        If CODEX$ = "" Or Len(CODEX$) < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Inexistente o No Valido"
          List1.AddItem MERROX$
          GoTo 19
        ElseIf Len(CODEX$) > 15 Then ' TRUNCA CODIGOS CON MAS DE 15 POSICIONES
          COALT1$ = Left$(CODEX$, 24)
          For KKI% = 1 To 9999
            CDXYX$ = TRIM$(Str$(KKI%))
            While Len(CDXYX$) < 4: CDXYX$ = "0" + CDXYX$: Wend
            CDXYX$ = Left$(CODEX$, 1) + "-" + CDXYX$
            If CODINT%(CDXYX$) < 1 Then
              For KKJ% = 1 To List4.ListCount
                If TRIM$(Left$(List4.List(KKJ% - 1), 15)) = CDXYX$ Then GoTo 10
              Next KKJ%
              CODEX$ = CDXYX$
              GoTo 11
            End If
10        Next KKI%
          MERROX$ = "Codigo " & "'" & codex1$ & "'" & " Excede 15 Posiciones"
          GoTo 19
        End If
        '
11      CDXI% = CODINT%(CODEX$)
        If CDXI% > 0 Then
          If ATX% = 2 Then GoTo 19  ' SALTEAR SIEMPRE
          If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
          ATX% = COMALTER%("Código / Campo Clave '" + CODEX$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
          If ATX% <= 2 Then GoTo 19 ' SALTEAR
        End If
        '
12      Call REPLA(TTXX$, CODEX$, 1, 16)
        VALO$ = AJUSTD$(Str$(CDXI%), 6)
        Call REPLA(TTXX$, VALO$, 17, 6)
        '
        '\\\OT-05-0499-WAR-13/07/05///
        COALT$ = ""
        On Error Resume Next
        DbGrid1.Col = 1
        COALT$ = TRIM$(DbGrid1.TEXT) '.Recordset.Fields(1))
        On Error GoTo 0
        If COALT$ = "" Then COALT$ = COALT1$
'        If Len(COALT$) < 1 Then
'          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Código Alternativo"
'          List1.AddItem MERROX$
'          GoTo 19
'        End If
        Call REPLA(TTXX$, COALT$, 25, 24)
        '\\\OT-05-0499-WAR-FIN///
        '
        DEDEX$ = ""
        On Error Resume Next
        DbGrid1.Col = 2
        DEDEX$ = TRIM$(DbGrid1.TEXT) '.Recordset.Fields(2))
        On Error GoTo 0
        If Len(DEDEX$) < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, DEDEX$, 50, 64)
        '
        '\\\OT-05-0499-WAR-13/07/05///
        DECOM$ = ""
        On Error Resume Next
        DbGrid1.Col = 3
        DECOM$ = TRIM$(DbGrid1.TEXT) '.Recordset.Fields(3))
        On Error GoTo 0
        'If Len(DECOM$) < 1 Then
        '  MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion Complementaria"
        '  List1.AddItem MERROX$
        '  GoTo 19
        'End If
        Call REPLA(TTXX$, DECOM$, 115, 64)
        '
        ATRA$ = ""
        On Error Resume Next
        DbGrid1.Col = 4
        ATRA$ = TRIM$(DbGrid1.TEXT) '.Recordset.Fields(4))
        On Error GoTo 0
'        If Len(ATRA$) < 1 Then
'          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Atributo A"
'          List1.AddItem MERROX$
'          GoTo 19
'        End If
        Call REPLA(TTXX$, ATRA$, 180, 32)
        '
        ATRB$ = ""
        On Error Resume Next
        DbGrid1.Col = 5
        ATRB$ = TRIM$(DbGrid1.TEXT) '.Recordset.Fields(5))
        On Error GoTo 0
        'If Len(ATRB$) < 1 Then
        '  MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Atributo B"
        '  List1.AddItem MERROX$
        '  GoTo 19
        'End If
        Call REPLA(TTXX$, ATRB$, 213, 32)
        '
        '\\\OT-05-0499-WAR-FIN///
        '
        VALO$ = ""
        On Error Resume Next
        DbGrid1.Col = 6
        VALO$ = AJUSTD$(TRIM$(DbGrid1.TEXT), 4) '.Recordset.Fields(6), 4)
        On Error GoTo 0
        If XVALO(VALO$) < 0 Or XVALO(VALO$) > 255 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Tipo de Material no Procesable"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, VALO$, 246, 4)
        '
        UME$ = "U."
        On Error Resume Next
        DbGrid1.Col = 7
        UME$ = TRIM$(DbGrid1.TEXT) '.Recordset.Fields(7))
        On Error GoTo 0
        Call REPLA(TTXX$, UME$, 251, 4)
        '
        VALOR = 0
        On Error Resume Next
        DbGrid1.Col = 8
        VALOR = CDbl(DbGrid1.TEXT) '.Recordset.Fields(8))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALOR, "F12.1")
        If XVALO(VALO$) < 0 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Stock Minimo no Procesable"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, VALO$, 256, 12)
        '
        VALOR = 0
        On Error Resume Next
        DbGrid1.Col = 9
        VALOR = CDbl(DbGrid1.TEXT) '.Recordset.Fields(9))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALOR, "F12.1")
        If XVALO(VALO$) < 0 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Lote de Reposicion no Procesable"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, VALO$, 269, 12)
        '
        '\\\OT-5-0633-FG-22/09/05///
        VALOR = 0                           ' KILOS POR UNIDAD
        On Error Resume Next
        DbGrid1.Col = 10
        VALOR = CDbl(DbGrid1.TEXT) '.Recordset.Fields(10))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALOR, "F8.3")
        Call REPLA(TTXX$, VALO$, 282, 8)
        '
        VALOR = 0                           ' KILOS POR METRO
        On Error Resume Next
        DbGrid1.Col = 11
        VALOR = CDbl(DbGrid1.TEXT) '.Recordset.Fields(11))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALOR, "F8.3")
        Call REPLA(TTXX$, VALO$, 292, 8)
        '\\\OT-5-0633-FG-FIN///
        '
        VALOR = 0
        On Error Resume Next
        DbGrid1.Col = 12
        VALOR = CDbl(DbGrid1.TEXT) '.Recordset.Fields(12))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALOR, "F12.4")
        Call REPLA(TTXX$, VALO$, 302, 12)
        '
        On Error Resume Next
        DbGrid1.Col = 13
        VALO$ = AJUSTD$(DbGrid1.TEXT, 4) '.Recordset.Fields(13), 4)
        On Error GoTo 0
         If XVALO(VALO$) < 0 Or XVALO(VALO$) > 255 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Moneda de Costeo no Procesable"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, VALO$, 315, 4)
        '
        VALO$ = ""
        On Error Resume Next
        DbGrid1.Col = 14
        VALO$ = Format$(DbGrid1.TEXT, "dd/mm/yy") '.Recordset.Fields(14), "dd/mm/yy")
        On Error GoTo 0
        Call REPLA(TTXX$, VALO$, 320, 8)
        '
        VALOR = 0
        On Error Resume Next
        DbGrid1.Col = 15
        VALOR = CDbl(DbGrid1.TEXT) '.Recordset.Fields(15))
        On Error GoTo 0
        If VALOR < 0 Or VALOR > 32766 Then VALOR = 0
        VALO$ = FORMATNUM$(VALOR, "F8.0")
        Call REPLA(TTXX$, VALO$, 329, 8)
        '
        VALOR = 0                           ' GRUPO DE VENTA
        On Error Resume Next
        DbGrid1.Col = 16
        VALOR = CDbl(DbGrid1.TEXT) '.Recordset.Fields(16))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALOR, "F6.0")
        Call REPLA(TTXX$, VALO$, 338, 6)
        '
        VALOR = 0
        On Error Resume Next
        DbGrid1.Col = 17
        VALOR = CDbl(DbGrid1.TEXT) '.Recordset.Fields(17))
        On Error GoTo 0
        If VALOR < (-9999999) Or VALOR > 9999999 Then VALOR = 0
        VALO$ = FORMATNUM$(VALOR, "F10.1")
        Call REPLA(TTXX$, VALO$, 346, 10)
        '
        VALO$ = ""
        On Error Resume Next
        DbGrid1.Col = 18
        VALO$ = SAFETEXT$(DbGrid1.TEXT) '.Recordset.Fields(18))
        On Error GoTo 0
        Call REPLA(TTXX$, VALO$, 360, 64)
        '
        List4.AddItem TTXX$
        Label6 = TRIM$(Str$(List4.ListCount))
        If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
          TEPRUE = RTrim(TTXX$ + "*")
        End If
        On Error Resume Next
        List4.TopIndex = List4.ListCount - 7
        On Error GoTo 0
        DoEvents
19  ' .Recordset.MoveNext
     Loop
   'End With
   '
20 Screen.MousePointer = 1
   '
End Sub

Sub PREIMPORSTRU()
   '
   If DbGrid1.Cols < 4 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = DbGrid1.Rows 'Data1.Recordset.RecordCount
   ATX% = 0: CODEXA$ = ""
   '
   'hrow = DBGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   'DBGrid1.RowHeight = 10
   '
   HUBOAMAS% = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(1024)
     DbGrid1.Row = U& - 1
     DbGrid1.Col = 0
5    CODEX$ = TRIM$(DbGrid1.TEXT): If CODEX$ = "" Then CODEX$ = CODEXA$
     If CODINT%(CODEX$) < 1 Then
       CODEX$ = Left$(CODEX$, 15)
       If CODINT%(CODEX$) < 1 Then
         TTXX1$ = "Codigo de Conjunto '" + TRIM$(CODEX$) + "' Inexistente en B.Datos."
         TTXX1$ = TTXX1$ + "¿ Desea Operación de Alta Interactiva ?\\Si, Darlo de Alta\No, Continuar"
         If COMALTER%(TTXX1$) = 1 Then
           VDEF$ = String$(128, 0)
           Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
           Call REPLA(VDEF$, "U.", 109, 2)
           Call REPLA(VDEF$, Chr$(1), 93, 1)
           Call REPLA(VDEF$, CODEX$, 77, 16)
           DEDEX$ = ""
           Call REPLA(VDEF$, DEDEX$, 5, 64)
           'Call AMAS_GES04.ALTACODIGO(VDEF$, OKX%)
           Call AMAS_GES07.ALTACODIGO(VDEF$, OKX%)
           HUBOAMAS% = 1
           If OKX% = 1 Then GoTo 5
         End If
       End If
     End If
     '
     If CODEX$ = "" Or Len(CODEX$) < 1 Or Len(CODEX$) > 24 Or CODINT%(CODEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Conjunto '" + CODEX$ + "' Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
     If CODEX$ = CODEXA$ Then GoTo 10 ' EL MISMO CODIGO
     '
     CODEXA$ = CODEX$
     Call LEEEXPLO(CODEX$, 1)
     If CAIT% > 0 And ATX% = 2 Then GoTo 19
     '
     'ATX% = 3
     If CAIT% > 0 Then
         ATX% = COMALTER%("Código '" + CODEX$ + "' con Estructura en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
       Else
         GoTo 12
     End If
     '
10   If ATX% <= 2 And CAIT% > 0 Then GoTo 19 ' SALTEAR
     '
12   Call REPLA(TTXX$, CODEX$, 1, 24)
     '
     DbGrid1.Col = 1
13   COELE$ = Mid$(TRIM$(DbGrid1.TEXT), 1, 24)
     If CODINT%(COELE$) < 1 Then
       COELE$ = Left$(COELE$, 15)
       If CODINT%(COELE) < 1 Then
         TTXX1$ = "Codigo de Elemento '" + TRIM$(COELE$) + "' Inexistente en B.Datos."
         TTXX1$ = TTXX1$ + "¿ Desea Operación de Alta Interactiva ?\\Si, Darlo de Alta\No, Continuar"
         If COMALTER%(TTXX1$) = 1 Then
           VDEF$ = String$(128, 0)
           Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
           Call REPLA(VDEF$, "U.", 109, 2)
           Call REPLA(VDEF$, Chr$(1), 93, 1)
           Call REPLA(VDEF$, COELE$, 77, 16)
           DEDEX$ = ""
           Call REPLA(VDEF$, DEDEX$, 5, 64)
           Call AMAS_GES07.ALTACODIGO(VDEF$, OKX%)
           HUBOAMAS% = 1
           If OKX% = 1 Then GoTo 13
         End If
       End If
     End If
     '
     If COELE$ = "" Or Len(COELE$) < 1 Or Len(COELE$) > 24 Or COELE$ = CODEX$ Or CODINT%(COELE$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Elemento '" + COELE$ + "' Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, COELE$, 27, 24)
     Call REPLA(TTXX$, DESCRIT0$(CODINT%(COELE$)), 53, 32)
     '
     DbGrid1.Col = 2                     ' USO UNITARIO
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DbGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.4")
     Call REPLA(TTXX$, VALO$, 85, 12)
     '
     DbGrid1.Col = 3                     ' UNIDADES
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DbGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.4")
     Call REPLA(TTXX$, VALO$, 98, 12)
     '
     DbGrid1.Col = 4                     ' LARGO DE CORTE
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DbGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.3")
     Call REPLA(TTXX$, VALO$, 111, 12)
     '
     DbGrid1.Col = 5                     ' PLAZO DISPOSICION
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DbGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "I4")
     Call REPLA(TTXX$, VALO$, 124, 4)
     '
     DbGrid1.Col = 6                     ' REFERENCIA
     On Error Resume Next
     VALO$ = DbGrid1.TEXT
     On Error GoTo 0
     Call REPLA(TTXX$, VALO$, 131, 890)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     '
19 Next U&
   '
   'DBGrid1.RowHeight = hrow
   '
   If HUBOAMAS% > 0 Then    ' HUBO CODIGOS AGREGADOS
     UPDATMASTER% = 1
     Call GENECOMAS
   End If
   '
   Screen.MousePointer = 1
   '
20 End Sub

Sub PREIMPORUTS()
   '
   If DbGrid1.Cols < 12 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = DbGrid1.Rows 'Data1.Recordset.RecordCount
   ATX% = 0: CODEXA$ = "": CODEYA$ = ""
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DbGrid1.Row = U& - 1
     DbGrid1.Col = 0
     CODEX$ = TRIM$(DbGrid1.TEXT)
     If CODEX$ = "" Or Len(CODEX$) > 15 Or CODINT%(CODEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Producto Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
     If CODEX$ = CODEXA$ Then GoTo 10 ' EL MISMO CODIGO
     '
     CODEXA$ = CODEX$
     RFF$ = RECFILT$("SECOPER", 1, MKI$(CODINT%(CODEX$)))
     If Len(RFF$) > 0 And ATX% = 2 Then GoTo 19
     '
     'ATX% = 3   ' SOBREESCRIBIR
     If Len(RFF$) > 0 Then
         ATX% = COMALTER%("Secuencia Operaciones de Código '" + CODEX$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
       Else
         GoTo 12
     End If
     '
10   If ATX% <= 2 And Len(RFF$) > 0 Then GoTo 19 ' SALTEAR
     '
12   Call REPLA(TTXX$, CODEX$, 1, 16)
     If CODEX$ <> CODEYA$ Then NOPE% = 0
     CODEYA$ = CODEX$
     '
     DbGrid1.Col = 1
     DEOPE$ = TRIM$(DbGrid1.TEXT)
     If DEOPE$ = "" Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion de la Operacion"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     NOPE% = NOPE% + 10
     VALO$ = FORMATNUM$(NOPE%, "I4")
     Call REPLA(TTXX$, VALO$, 17, 4)
     Call REPLA(TTXX$, DEOPE$, 23, 48)
     '
     DbGrid1.Col = 2                     ' MAQUINA
     CMAQUI$ = TRIM$(DbGrid1.TEXT)
     If CMAQUI$ <> "" Then
       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina '" + CMAQUI$ + "' Inexistente o No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CMAQUI$, 73, 6)
     '
     DbGrid1.Col = 3                     ' HORAS PREPARACION
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DbGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F10.3")
     Call REPLA(TTXX$, VALO$, 79, 10)
     '
     DbGrid1.Col = 4                     ' HORAS STANDARD FIJO
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DbGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F10.3")
     Call REPLA(TTXX$, VALO$, 89, 10)
     '
     DbGrid1.Col = 5                     ' MINUTOS STANDARD VARIABLE
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DbGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.5")
     Call REPLA(TTXX$, VALO$, 99, 12)
     '
     DbGrid1.Col = 6                     ' LOTE STANDARD UNIDADES
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DbGrid1.TEXT)
     On Error GoTo 0
     If VALOR < 1 Then VALOR = 1
     VALO$ = FORMATNUM$(VALOR, "F10.1")
     Call REPLA(TTXX$, VALO$, 111, 10)
     '
     DbGrid1.Col = 7                     ' % DEM.DISTRIBUIDA
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DbGrid1.TEXT)
     If VALOR < 0 Then VALOR = 0
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F8.1")
     Call REPLA(TTXX$, VALO$, 121, 8)
     '
     DbGrid1.Col = 8                     ' OPERARIOS
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DbGrid1.TEXT)
     On Error GoTo 0
     If VALOR < 0.01 Then VALOR = 1
     VALO$ = FORMATNUM$(VALOR, "F8.2")
     Call REPLA(TTXX$, VALO$, 129, 8)
     '
     DbGrid1.Col = 9                     ' MAQUINA AUX 1
     CMAQUI$ = TRIM$(DbGrid1.TEXT)
     If CMAQUI$ <> "" Then
       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina Aux.1'" + CMAQUI$ + "' Inexistente o No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CMAQUI$, 139, 6)
     '
     DbGrid1.Col = 10                     ' MAQUINA AUX 2
     CMAQUI$ = TRIM$(DbGrid1.TEXT)
     If CMAQUI$ <> "" Then
       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina Aux.2'" + CMAQUI$ + "' Inexistente o No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CMAQUI$, 147, 6)
     '
     DbGrid1.Col = 11                     ' MAQUINA AUX 3
     CMAQUI$ = TRIM$(DbGrid1.TEXT)
     If CMAQUI$ <> "" Then
       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina Aux.3'" + CMAQUI$ + "' Inexistente o No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CMAQUI$, 155, 6)
     '
     DbGrid1.Col = 12                     ' H.P.FABR.
     HOPROFA$ = TRIM$(DbGrid1.TEXT)
     Call REPLA(TTXX$, HOPROFA$, 163, 8)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     '
19 Next U&
   Screen.MousePointer = 1
   '
20 End Sub

Sub PREIMPORMAQ()
   '
   If DbGrid1.Cols < 7 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   ATX% = 0
   URECORD& = DbGrid1.Rows 'Data1.Recordset.RecordCount
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DbGrid1.Row = U& - 1
     DbGrid1.Col = 0
     CODEX$ = TRIM$(DbGrid1.TEXT)
     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 6 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Maquina Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     RFF$ = RECFILT$("PADMAQ", 1, CODEX$)
     If Len(RFF$) > 0 Then
       If ATX% = 2 Then GoTo 19  ' SALTEAR SIEMPRE
       If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
       ATX% = COMALTER%("Código / Campo Clave '" + CODEX$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
       If ATX% <= 2 Then GoTo 19 ' SALTEAR
     End If
     '
12   Call REPLA(TTXX$, CODEX$, 1, 6)
     '
     DbGrid1.Col = 1
     DEDEX$ = TRIM$(DbGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, DEDEX$, 9, 30)
     '
     DbGrid1.Col = 2                     ' Cuota Horaria
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DbGrid1.TEXT)
     On Error GoTo 0
     If VALOR < 0 Or VALOR > 999999.99 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Cuota Horaria fuera de Rango"
       List1.AddItem MERROX$
       GoTo 19
     End If
     VALO$ = FORMATNUM$(VALOR, "F12.4")
     Call REPLA(TTXX$, VALO$, 41, 12)
     '
     DbGrid1.Col = 3                    ' Seccion
     VTX$ = TRIM$(DbGrid1.TEXT)
     If VTX$ <> "" And XVALO(VTX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Seccion debe ser numerico 'nnn.nnn'"
       List1.AddItem MERROX$
       GoTo 19
     End If
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(VTX$)
     On Error GoTo 0
     If VALOR > 999.999 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Seccion no Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     VALO$ = FORMATNUM$(VALOR, "F10.3")
     Call REPLA(TTXX$, VALO$, 53, 10)
     '
     DbGrid1.Col = 4                    ' Maquina equivalente
     CODEQ$ = TRIM$(DbGrid1.TEXT)
     If CODEQ$ <> "" Then
       If Len(CODEQ$) < 3 Or Len(CODEQ$) > 6 Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Maquina Equivalente No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CODEQ$, 66, 6)
     '
     DbGrid1.Col = 5                     ' Coeficiente de equivalencia
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DbGrid1.TEXT)
     On Error GoTo 0
     VAMIN = 0: If CODEQ$ <> "" Then VAMIN = 10
     If CODEQ$ = "" Then VALOR = 0
     If VALOR < VAMIN Or VALOR > 999.9 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Coeficiente Equivalencia fuera de Rango"
       List1.AddItem MERROX$
       GoTo 19
     End If
     VALO$ = FORMATNUM$(VALOR, "F12.1")
     Call REPLA(TTXX$, VALO$, 72, 12)
     '
     DbGrid1.Col = 6                    ' Observaciones
     OBSER$ = TRIM$(DbGrid1.TEXT)
     Call REPLA(TTXX$, OBSER$, 87, 30)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     '
19 Next U&

20 End Sub

Sub PREIMPORCAT()
   '
   If DbGrid1.Cols < 3 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   ATX% = 0
   URECORD& = DbGrid1.Rows 'Data1.Recordset.RecordCount
   If URECORD& > 255 Then
     Call MENSERR(24, "Imposible Importar - Máximo 255 Registros.")
     GoTo 20
   End If
   '
   If ULTREG&("CATOPER") > 0 Then
     If COMALTER%("Tabla de Categorias Existente\\Conservar\Sobre-Escribir") <> 2 Then
       GoTo 20
     End If
   End If
   '
   JJ& = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     '
     DbGrid1.Row = U& - 1
     DbGrid1.Col = 0
     DECAT$ = TRIM$(DbGrid1.TEXT)
     If DECAT$ <> "" Then
       JJ& = JJ& + 1
       COCAT$ = AJUSTD$(Str$(JJ&), 4)
       Call REPLA(TTXX$, COCAT$, 1, 4)
       Call REPLA(TTXX$, DECAT$, 7, 30)
       '
       DbGrid1.Col = 1
       VALOR = 0
       On Error Resume Next
       VALOR = CDbl(DbGrid1.TEXT)
       On Error GoTo 0
       If VALOR < 0 Or VALOR > 999999.99 Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Cuota Horaria fuera de Rango"
         List1.AddItem MERROX$
         GoTo 19
       End If
       VALO$ = FORMATNUM$(VALOR, "F12.4")
       Call REPLA(TTXX$, VALO$, 37, 12)
       '
       DbGrid1.Col = 2
       OBSER$ = TRIM$(DbGrid1.TEXT)
       Call REPLA(TTXX$, OBSER$, 51, 32)
       '
       List4.AddItem TTXX$
       Label6 = TRIM$(Str$(List4.ListCount))
       If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
         TEPRUE = RTrim(TTXX$ + "*")
       End If
       On Error Resume Next
       List4.TopIndex = List4.ListCount - 7
       On Error GoTo 0
       DoEvents
       '
     End If
19 Next U&

20 End Sub

Sub RIMPORCRM()
   '
   Screen.MousePointer = 11
   Call ABRECONEXION
   '
   Call CREACAMPO("", "BSEGCOM", "MARULIMP", 3, 2, 1)
   Call ACTUREGISTRO("BSEGCOM", "MARULIMP", "MARULIMP > 0", 0, RAFE&)
   '
   SOBRESCRI% = Check1.Value           ' HABILITA SOBRE-ESCRIBE
   GRUPRE$ = TRIM$(Text3)
   If GRUPRE$ <> "" Then
      GRUMA% = 0
      For KU& = 1 To ULTREG&("TAGRUCON")
         XX$ = REGLEIDO$("TAGRUCON", KU&)
         GRUAC% = CVI(Left$(XX$, 2))
         If UCase$(TRIM$(Mid$(XX$, 3, 48))) = UCase$(GRUPRE$) Or UCase$(TRIM$(Mid$(XX$, 51, 24))) = UCase$(GRUPRE$) Then
            NGRUPRE% = GRUAC%
            GoTo 10
         End If
         If GRUAC% > GRUMA% Then GRUMA% = GRUAC%
      Next KU&
      NGRUPRE% = 1 + GRUMA%
      XX$ = MKI$(NGRUPRE%) + AJUSTI$(GRUPRE$, 48) + AJUSTI$(GRUPRE$, 24)
      Call REGAPP("TAGRUCON", XX$)
      Call CIERRARCH("TAGRUCON")
   End If
   '
10 URECORD& = DbGrid1.Rows 'Data1.Recordset.RecordCount
   ATX% = 0
   For U& = 2 To URECORD&                               ' empieza en 2 para no importar los encabezados
     Call TRACE(20, U&, URECORD&)
     Label7 = U&
     On Error Resume Next
     DbGrid1.Row = U& - 1
     If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Imposible Completar Pre-Importación.\Demasiados Registros en Tabla.")
       GoTo 20
     End If
     On Error GoTo 0
     '
     DbGrid1.Col = 0
     RACLI$ = TRIM$(DbGrid1.TEXT)
     '
     DbGrid1.Col = 1                     ' DOMICILIO
     DOMI$ = TRIM$(DbGrid1.TEXT)
     '
     DbGrid1.Col = 2                     ' CODIGO POSTAL
     CODPOS$ = TRIM$(DbGrid1.TEXT)
     '
     DbGrid1.Col = 3                     ' LOCALIDAD
     LOCA$ = TRIM$(DbGrid1.TEXT)
     '
     DbGrid1.Col = 4                     ' PROVINCIA
     PROVI$ = TRIM$(DbGrid1.TEXT)
     If PROVI$ = "" Then PROVI$ = "C"
     '
     DbGrid1.Col = 5                     ' PAIS
     PAIS$ = TRIM$(DbGrid1.TEXT)
     '
     DbGrid1.Col = 6                     ' TELEFONO
     TELEF$ = TRIM$(DbGrid1.TEXT)
     '
     DbGrid1.Col = 7                     ' FAX
     FAX$ = TRIM$(DbGrid1.TEXT)
     '
     DbGrid1.Col = 8                     ' EMAI$
     EMAI$ = TRIM$(DbGrid1.TEXT)
     '
     DbGrid1.Col = 9                     ' WEBSITE
     WEBS$ = TRIM$(DbGrid1.TEXT)
     '
     DbGrid1.Col = 10                    ' VENDEDOR
     VENDE% = XVALO(TRIM$(DbGrid1.TEXT))
     '
     DbGrid1.Col = 11                     ' PROVINCIA
     CONTAC1$ = TRIM$(DbGrid1.TEXT)
     '
     DbGrid1.Col = 12                     ' PAIS
     CARCON1$ = TRIM$(DbGrid1.TEXT)
     '
     DbGrid1.Col = 13                     ' PAIS
     MAICON1$ = TRIM$(DbGrid1.TEXT)

     DbGrid1.Col = 14                     ' TELEFONO
     TELCON1$ = TRIM$(DbGrid1.TEXT)
     '
     DbGrid1.Col = 15                     ' PROVINCIA
     CONTAC2$ = TRIM$(DbGrid1.TEXT)
     '
     DbGrid1.Col = 16                     ' PAIS
     CARCON2$ = TRIM$(DbGrid1.TEXT)
     '
     DbGrid1.Col = 17                     ' PAIS
     MAICON2$ = TRIM$(DbGrid1.TEXT)

     DbGrid1.Col = 18                     ' TELEFONO
     TELCON2$ = TRIM$(DbGrid1.TEXT)
     '
     DbGrid1.Col = 19                     ' OBSERVACIONES
     OBSERX$ = TRIM$(DbGrid1.TEXT)
     '
     If UCase$(CONTAC2$) = UCase$(CONTAC1$) Then
        CONTAC2$ = ""
        CARCON2$ = ""
        MAICON2$ = ""
        TELCON2$ = ""
     End If
     '
     GoSub IMPORDATOS
     '
19 Next U&

20 Screen.MousePointer = 1
   Call COMUNI("Fin de la Importacion")
   Unload Me
   '
Exit Sub
   'DbGrid1.RowHeight = hrow   ' RESTAURA ALTURA ANTERIOR
   '
IMPORDATOS:
   If RACLI$ = "" Then RACLI$ = CONTAC1$
   If RACLI$ = "" Then RACLI$ = CONTAC2$
   If RACLI$ = "" Then GoTo 99
   '
   QTR& = CantRegistros("BSEGCOM", "RASO_CLI = '" & RACLI$ & "'")
   If QTR& < 1 Then
        ' ASIGNA UN CODIGO DE CLIENTE Y CREA EL REGISTRO
        CODICLIX$ = FISECOM09.NUEVOPROESP
        Call CreaRegistro("BSEGCOM", "Codi_Cli", CODICLIX$)
        Call ACTUREGISTRO("BSEGCOM", "MARULIMP", "Codi_Cli = '" & CODICLIX$ & "'", 1, RAFE&)
      Else
        CODICLIX$ = VALOBADA("BSEGCOM", "Codi_Cli", "RASO_CLI = '" & RACLI$ & "'")
   End If
   '
   With FISECOM09
      .Label26 = CODICLIX$
      
      If TRIM$(.Text4) = "" Or SOBRESCRI% > 0 Then .Text4 = RACLI$
      If TRIM$(.Text20) = "" Or SOBRESCRI% > 0 Then .Text20 = DOMI$
      If TRIM$(.Text21) = "" Or SOBRESCRI% > 0 Then .Text21 = CODPOS$
      If TRIM$(.Text22) = "" Or SOBRESCRI% > 0 Then .Text22 = LOCA$
      If TRIM$(.Text27) = "" Or SOBRESCRI% > 0 Then .Text27 = PROVI$
      If TRIM$(.Text23) = "" Or SOBRESCRI% > 0 Then .Text23 = PAIS$
      If TRIM$(.Text26) = "" Or SOBRESCRI% > 0 Then .Text26 = TELEF$
      If TRIM$(.Text24) = "" Or SOBRESCRI% > 0 Then .Text24 = FAX$
      If TRIM$(.Text5) = "" Or SOBRESCRI% > 0 Then .Text5 = EMAI$
      If TRIM$(.Text13) = "" Or SOBRESCRI% > 0 Then .Text13 = WEBS$
      If TRIM$(.Text9) = "" Or SOBRESCRI% > 0 Then .Text9 = OBSERX$
      If TRIM$(.Label42) = "" Or SOBRESCRI% > 0 Then .Label42 = VENDE%
      If TRIM$(Text1) <> "" And (TRIM$(.Text1) = "" Or SOBRESCRI% > 0) Then .Text1 = Text1
      If TRIM$(Text2) <> "" And (TRIM$(.Text2) = "" Or SOBRESCRI% > 0) Then .Text2 = Text2
      '
      PUNGRU$ = SAFETEXT$(VALOBADA("BSEGCOM", "Grupos", "CODI_CLI = '" & CODICLIX$ & "'"))
      If InStr(PUNGRU$, Chr$(NGRUPRE%)) < 1 Then
         PUNGRU$ = Chr$(NGRUPRE%) + PUNGRU$
         UTACON% = ULTREG&("TAGRUCON")
         .List2.Clear
         For UKI% = 1 To Len(PUNGRU$)
           RETA& = Asc(Mid$(PUNGRU$, UKI%, 1))
           If RETA& > 0 Then
             If RETA& <= UTACON% Then
               XX$ = REGLEIDO$("TAGRUCON", RETA&)
               RAGRE% = Val(CVI(Left$(XX$, 2)))
               TAGRE$ = AJUSTI$(Mid$(XX$, 51, 24), 48) + AJUSTI$(Mid$(XX$, 3, 48), 48) + Str$(RAGRE%)
               .List2.AddItem TAGRE$
             End If
           End If
         Next UKI%
      End If
      '
      For KKI% = 1 To .Text6.Count
         If TRIM$(.Text6(KKI% - 1)) = "" Or UCase$(TRIM$(.Text6(KKI% - 1))) = UCase$(CONTAC1$) Then
            If TRIM$(.Text6(KKI% - 1)) = "" Or SOBRESCRI% > 0 Then .Text6(KKI% - 1) = CONTAC1$
            If TRIM$(.Text7(KKI% - 1)) = "" Or SOBRESCRI% > 0 Then .Text7(KKI% - 1) = CARCON1$
            If TRIM$(.Text8(KKI% - 1)) = "" Or SOBRESCRI% > 0 Then .Text8(KKI% - 1) = MAICON1$
            If TRIM$(.Text10(KKI% - 1)) = "" Or SOBRESCRI% > 0 Then .Text10(KKI% - 1) = TELCON1$
            If TRIM$(.Text30(KKI% - 1)) = "" Or SOBRESCRI% > 0 Then .Text30(KKI% - 1) = TELCEL1$
            GoTo 80
         End If
      Next KKI%
      '
80    For KKI% = 1 To .Text6.Count
         If TRIM$(.Text6(KKI% - 1)) = "" Or UCase$(TRIM$(.Text6(KKI% - 1))) = UCase$(CONTAC2$) Then
            If TRIM$(.Text6(KKI% - 1)) = "" Or SOBRESCRI% > 0 Then .Text6(KKI% - 1) = CONTAC2$
            If TRIM$(.Text7(KKI% - 1)) = "" Or SOBRESCRI% > 0 Then .Text7(KKI% - 1) = CARCON2$
            If TRIM$(.Text8(KKI% - 1)) = "" Or SOBRESCRI% > 0 Then .Text8(KKI% - 1) = MAICON2$
            If TRIM$(.Text10(KKI% - 1)) = "" Or SOBRESCRI% > 0 Then .Text10(KKI% - 1) = TELCON2$
            If TRIM$(.Text30(KKI% - 1)) = "" Or SOBRESCRI% > 0 Then .Text30(KKI% - 1) = TELCEL2$
            GoTo 90
         End If
      Next KKI%
      '
90 End With
   '
FISECOM09.Show 1
AAA = IMPORTANDO
   Call FISECOM09.Command1_Click
   Call FISECOM09.Command3_Click
   '
99 Return

End Sub

Sub PREIMPORPRO()
   '
   PUNUCLI$ = String$(32767, 0)
   Dim CUTTX$(10000)
   '
   If DbGrid1.Cols < 16 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = DbGrid1.Rows 'Data1.Recordset.RecordCount
   'MODIF ECHA POR UNA IMPORTACIÓN ECHA EN TELEBIT 16/06/05
   'hrow = DbGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   'DbGrid1.RowHeight = 10
   '
   Call APERBASDAT("PROVE")
   SQLX$ = "SELECT * FROM PROVED12"
   Dim PROVVVED As ADODB.Recordset
   Set PROVVVED = New ADODB.Recordset
   PROVVVED.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   '
   With PROVVVED
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       Do While Not .EOF
         NCX& = !Nume_Cli
         If NCX& > 0 Then
           If NCX& <= 32767 Then
             Mid$(PUNUCLI$, NCX&, 1) = Chr$(1)
           End If
         End If
       .MoveNext
       Loop
     End If
   End With
   PROVVVED.Close
   Set PROVVVED = Nothing
   
   Call ABREPROVE
   ATX% = 0: NUCU& = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(297)
     On Error Resume Next
     DbGrid1.Row = U& - 1
     If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Imposible Completar Pre-Importación.\Demasiados Registros en Tabla.")
       GoTo 20
     End If
     On Error GoTo 0
     '
     DbGrid1.Col = 8                     ' CUIT
     CUIT$ = TRIM$(DbGrid1.TEXT)
     If CUIT$ <> "" Then
        VLBD$ = VALOBADA("Proved12", "Cuit_Cli", "CUIT_CLI = '" & CUIT$ & "'")
        If VLBD$ <> "" Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - CUIT de Proveedor Existente"
          List1.AddItem MERROX$
          GoTo 19
        End If
     End If
     For KKL& = 1 To NUCU&
       If CUTTX$(KKL&) = CUIT$ Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - CUIT de Proveedor Repetido"
          List1.AddItem MERROX$
          GoTo 19
        End If
     Next KKL&
     '
     DbGrid1.Col = 0
     NUPROV& = XVALO(TRIM$(DbGrid1.TEXT))   ' NUMERO DE CUENTA
     If NUPROV& < 1 Or NUPROV& > 32000 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Numero de Proveedor no Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     ' Verifica si el numero ya ha sido utilizado
25   If Asc(Mid$(PUNUCLI$, NUPROV&, 1)) > 0 Then
       NUPROV& = NUPROV& + 10000
       If NUPROV& > 32000 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Numero de Proveedor No Procesable"
          List1.AddItem MERROX$
          GoTo 19
       End If
       GoTo 25
     End If
     Mid$(PUNUCLI$, NUPROV&, 1) = Chr$(1)
     '
     CODEX$ = AJUSTD$(Str$(Int(NUPROV + 0.5)), 5)
12   Call REPLA(TTXX$, CODEX$, 1, 5)
     '
     DbGrid1.Col = 1                     ' RAZON SOCIAL
     DEDEX$ = TRIM$(DbGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta R.Social del Proveedor"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, DEDEX$, 7, 48)
     '
     DbGrid1.Col = 2                     ' DOMICILIO
     DOMI$ = AJUSTI$(DbGrid1.TEXT, 64)
     Call REPLA(TTXX$, DOMI$, 55, 64)
     '
     DbGrid1.Col = 3                     ' CODIGO POSTAL
     CODPOS$ = TRIM$(DbGrid1.TEXT)
     Call REPLA(TTXX$, CODPOS$, 103, 16)
     '
     DbGrid1.Col = 4                     ' LOCALIDAD
     LOCA$ = TRIM$(DbGrid1.TEXT)
     Call REPLA(TTXX$, LOCA$, 115, 48)
     '
     DbGrid1.Col = 5                     ' TELEFONO
     TELEF$ = TRIM$(DbGrid1.TEXT)
     Call REPLA(TTXX$, TELEF$, 147, 24)
     '
     DbGrid1.Col = 6                     ' FAX
     FAX$ = TRIM$(DbGrid1.TEXT)
     Call REPLA(TTXX$, FAX$, 171, 24)
     '
     DbGrid1.Col = 7                     ' PERSONA DE CONTACTO
     CONTACTO$ = TRIM$(DbGrid1.TEXT)
     Call REPLA(TTXX$, CONTACTO$, 195, 24)
     '
     DbGrid1.Col = 8                     ' CUIT
     CUIT$ = TRIM$(DbGrid1.TEXT)
     Call REPLA(TTXX$, CUIT$, 219, 16)
     '
     DbGrid1.Col = 9                     ' TIPO DE CUENTA
     TIPOCTA$ = TRIM$(DbGrid1.TEXT)
     If TIPOCTA$ = "" Then TIPOCTA$ = "0"
     If TIPOCTA$ = "0" Or TIPOCTA$ = "1" Then
       Call REPLA(TTXX$, TIPOCTA$, 235, 4)
     Else
       MERROX$ = "Registro" + Str$(U& - 1) + " - Error en Tipo de Cuenta"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     DbGrid1.Col = 10                    ' CODIGO DE CUENTA
     CODICUEN$ = TRIM$(DbGrid1.TEXT) + "/B"
     Call REPLA(TTXX$, CODICUEN$, 240, 12)
     '
     '\\\OT-05-0202-WAR-18/04/05///
     DbGrid1.Col = 11                    ' CODIGO DE PROVINCIA
     CODIPROVI$ = TRIM$(DbGrid1.TEXT)
     Call REPLA(TTXX$, CODIPROVI$, 253, 2)
     '
     DbGrid1.Col = 12                    ' CODIGO DE POSICION IVA
     CODIVA% = XVALO(DbGrid1.TEXT)
     Call REPLA(TTXX$, FORMATNUM$(CODIVA%, "I4"), 255, 4)
     '\\\OT-05-0202-WAR-FIN///
     '
     DbGrid1.Col = 13                    ' CUENTA MADRE
     CTAMADRE$ = TRIM$(DbGrid1.TEXT)
     If Len(CTAMADRE$) > 12 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Error en Longitud de Cuenta Madre"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, CTAMADRE, 261, 12)
     '\\\OT-05-0202-WAR-FIN///
     '
     DbGrid1.Col = 14                    ' nombre de fantasia
     FANTASI$ = TRIM$(DbGrid1.TEXT)
     Call REPLA(TTXX$, FANTASI$, 275, 16)
     '
     DbGrid1.Col = 15                    ' RETENCION DE INGRESOS BRUTOS
     VALOR = 0
     On Error Resume Next
     VALOR = XVALO(DbGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM(VALOR, "F4.0")
     Call REPLA(TTXX$, VALO$, 293, 4)
     '
17   List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim$(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim$(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
19 Next U&
   '
20 'DbGrid1.RowHeight = hrow
   Screen.MousePointer = 1
   '
End Sub

Sub PREIMPORPLA()
   '
   If DbGrid1.Cols < 5 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   'hrow = DbGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   'DbGrid1.RowHeight = 10
   '
   ATX% = 0
   URECORD& = DbGrid1.Rows 'Data1.Recordset.RecordCount
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DbGrid1.Row = U& - 1
     DbGrid1.Col = 0
     '
     NIVEI% = XVALO(DbGrid1.TEXT)
     If NIVEI% < 0 Or NIVEI% > 5 Then NIVEI% = 0
     Call REPLA(TTXX$, AJUSTD$(Str$(NIVEI%), 4), 1, 4)
     '
     DbGrid1.Col = 1
     CODEX$ = TRIM$(DbGrid1.TEXT)
     If CODEX$ = "" Or Len(CODEX$) < 1 Or Len(CODEX$) > 12 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Cuenta Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, CODEX$, 7, 12)
     '
     For KKU% = 1 To List4.ListCount
       If TRIM$(Mid$(List4.List(KKU% - 1), 7, 12)) = CODEX$ Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Cuenta Duplicado o Pre-Existente"
         List1.AddItem MERROX$
         GoTo 19
       End If
     Next KKU%
     '
     DbGrid1.Col = 2
     DEDEX$ = TRIM$(DbGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, DEDEX$, 21, 48)
     '
     DbGrid1.Col = 3                    ' Seccion
     VTX$ = TRIM$(UCase$(DbGrid1.TEXT))
     If VTX$ <> "D" And VTX$ <> "H" And VTX$ <> "I" Then VTX$ = "G"
     Call REPLA(TTXX$, VTX$, 71, 1)
     '
     DbGrid1.Col = 4
     VTX$ = TRIM$(UCase$(DbGrid1.TEXT))
     If VTX$ = "0" Or VTX$ = "NO" Then VTX$ = "NO"
     If VTX$ <> "NO" Then VTX$ = "SI"
     Call REPLA(TTXX$, VTX$, 75, 2)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     '
19 Next U&
   '
   'DbGrid1.RowHeight = hrow   ' RESTAURA ALTURA ANTERIOR
   '
20 End Sub
 '\\\OT-05-0207-WAR-21/04/05///
Sub PREIMPORNORCLI()
   '
   If DbGrid1.Cols < 3 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = DbGrid1.Rows 'Data1.Recordset.RecordCount
   ATX% = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DbGrid1.Row = U& - 1
     DbGrid1.Col = 0            'CLIENTE CODIGO STRING
     CODEX$ = TRIM$(DbGrid1.TEXT)
     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 15 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     If U& = 250 Then
       XX = "AA"
     End If
     '
12   Call REPLA(TTXX$, CODEX$, 1, 12)
     '
     DbGrid1.Col = 1            ' ARTICULO
     DEDEX$ = TRIM$(DbGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Articulo"
       List1.AddItem MERROX$
       GoTo 19
     End If
     CDXI% = CODINT%(DEDEX$)
     If CDXI% < 0 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo no Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     Call REPLA(TTXX$, DEDEX$, 14, 16)
     Call REPLA(TTXX$, DESCRIT0(CDXI%), 32, 22)
     '
     DbGrid1.Col = 2                     ' NUMERO ORIGINAL
     VALO$ = TRIM$(DbGrid1.TEXT)
     If Len(VALO$) < 0 Or Len(VALO$) > 28 Or TRIM(VALO$) = "" Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Numero Original no Procesable"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, VALO$, 56, 28)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
19 Next U&

20 Screen.MousePointer = 1
   '
End Sub
'\\\OT-05-0207-WAR-FIN///
Sub IMPORMAS()
   '
   FIN& = List4.ListCount
   Call APERBASDAT("MASTER")
   '
   NMS% = NUMAS%
   'REBLA$ = REGBLAN$("MASTER")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     CODIT$ = TRIM$(Left$(TTXX$, 16))
     '\\\OT-05-0499-WAR-13/07/05///
       'ACA TOMAR NUEVOS CAMPOS
       CODALT$ = TRIM$(Mid$(TTXX$, 25, 24))
       DESCOMP$ = TRIM$(Mid$(TTXX$, 115, 64))
       ATRIA$ = TRIM$(Mid$(TTXX$, 180, 32))
       ATRIB$ = TRIM$(Mid$(TTXX$, 213, 32))
     '\\\OT-05-0499-WAR-FIN///
     DESIT$ = TRIM$(Mid$(TTXX$, 50, 64))
     TIMAT% = XVALO(Mid$(TTXX$, 246, 4)) Mod 256
     UNIME$ = TRIM$(Mid$(TTXX$, 251, 4))
     STOMI = XVALO(Mid$(TTXX$, 256, 12))
       If STOMI > 30000 Then STOMI = (-1) * STOMI / 100
       If STOMI < (-30000) Then STOMI = (-30000)
     LECON = XVALO(Mid$(TTXX$, 269, 12))
       If LECON > 30000 Then LECON = (-1) * LECON / 100
       If LECON < (-30000) Then LECON = (-30000)
     LOMAX = 1
       If LOMAX < Abs(LECON) Then LOMAX = LECON
     KILUNI = XVALO(Mid$(TTXX$, 282, 8))
     KILMET = XVALO(Mid$(TTXX$, 292, 8))
     COSUN = XVALO(Mid$(TTXX$, 302, 12))
     MONEC% = XVALO(Mid$(TTXX$, 315, 4)) Mod 256
     FECOU% = FECHANUM(Mid$(TTXX$, 320, 8))
       PRH = Abs(XVALO(Mid$(TTXX$, 329, 8)))
       While PRH > 30000: PRH = PRH - 10000: Wend
     PRHAB% = PRH
     '
     GRUCO = XVALO(Mid$(TTXX$, 338, 6))
     STOAC = XVALO(Mid$(TTXX$, 346, 10))
     CINTERNA$ = Mid(TTXX$, 360, 64)
     '
     SQLX$ = "SELECT * FROM Master "
     Dim ARTICULOS As ADODB.Recordset
     Set ARTICULOS = New ADODB.Recordset
     ARTICULOS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     'With ARTICULOS
     On Error GoTo 0
     If CODINT%(CODIT$) = 0 Then
       With ARTICULOS
         .AddNew   ' YA CONVERTIDO
         On Error Resume Next
         !CodiEmpr = CodiEmp%
         On Error GoTo 0
         '\\\OT-05-0499-WAR-13/07/05///
         !NPLANO = CODALT$
         !DESCRIDOS = DESCOMP$
         !TALLE = ATRIA$
         !COLOR = ATRIB$
         '\\\OT-05-0499-WAR-FIN///
         !Codigo = CODIT$
         !Descripcion = DESIT$
         !Status = 1
           NMS% = NMS% + 1
         !CODINT = NMS%
         !UMEDIDA = AJUSTI$(UNIME$, 4)
         !TIMATE = TIMAT%
         !CRITREP = 0
         !Monecos = MONEC%
         !FECOSTO = CVDAT(FECOU%)
         !COSUNI = COSUN
         !STOMIN = STOMI
         !LOREPOS = LECON
         !STOMAX = LOMAX
         !PROVHABI = PRHAB%
         '
         !Alias = " "
         !PORNACIO = 0
         !PESUNI = 0
         '!COLOR = " "
         '!TALLE = " "
         !DESCOMPO = 0
         !GRUPIVA = 0
         !Grucoven = GRUCO
         !CLinea = " "
         !Cfamilia = " "
         '!NPLANO = " "
         !Revision = " "
         !FEREVIS = CVDAT(0)
         !Observa = " "
         !Especificacion = " "
         !PESUNI = KILUNI
         !PESMETRO = KILMET
         !Cod_Interna = CINTERNA$
        .Update
       End With
       '
       If Abs(STOAC) >= 0.05 Then
         HOII% = HOY(HOS$)
         NORI% = NORI% + 1
         Call MOVISTOK(HOII%, NMS%, "", "AJ", "AJ-" + HOS$, "AJ-" + HOS$, "AJ-" + HOS$, "AJ-" + HOS$, "Stock Importado", NORI%, CDbl(COSUN), MONEC%, 0, 0, STOAC)
       End If
       '
     End If
     '
   Next U&
   '
   ARTICULOS.Close
   Set ARTICULOS = Nothing
   UPDATMASTER% = 1
   Call GENECOMAS
   Screen.MousePointer = 1
   '
End Sub
'

Sub IMPORSTRU()
   '
   Dim PUNSTRU$(32767)
   '
   FIN& = List4.ListCount
   NMS% = NUMAS%
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     COCON$ = TRIM$(Left$(TTXX$, 24))
     CDXI% = CODINT%(COCON$)
     '
     If CDXI% > 0 Then
       If CDXI% <= NMS% Then
         PUNSTRU$(CDXI%) = PUNSTRU$(CDXI%) + MKS$(U&)
       End If
     End If
   Next U&
   '
   ATENTIA% = ATENTI%: ATENTI% = 12345 ' PA QUE NO REORGANICE IMPLOSI
   FIN& = NMS%
   For CDXI% = 1 To NMS%
     KUU& = CDXI%
     Call TRACE(20, KUU&, FIN&)
     If Len(PUNSTRU$(CDXI%)) > 0 Then
       CAIT% = 0
'Call LEEEXPLO(CODEXT$(CDXI%), 1)
       For UKU& = 1 To Len(PUNSTRU$(CDXI%)) Step 4
         U& = CVS(Mid$(PUNSTRU$(CDXI%), UKU&, 4))
         TTXX$ = List4.List(U& - 1)
         TTYY$ = Space$(128)
         COELEI% = CODINT%(TRIM$(Mid$(TTXX$, 27, 24)))
         COELE$ = CODEXT$(COELEI%)
         CAUNIDS = XVALO(Mid$(TTXX$, 98, 12))
         PDX$ = Mid$(TTXX$, 124, 4)
         PESOMET = PESMETRO(COELEI%)
         LARCORTE = XVALO(Mid$(TTXX$, 111, 12))
         '
         CANTI1 = XVALO(Mid$(TTXX$, 85, 12))
         USOUNI = 1
         If PESOMET > 0 Then
           If LARCORTE > 0 Then
             USOUNI = LARCORTE * PESOMET / 1000
             CANTI1 = CAUNIDS * USOUNI
           End If
         End If
         '
         Call REPLA(TTYY$, COELE$, 1, 16)
         USOX$ = FORMATNUM$(CANTI1, "F12.4")
         Call REPLA(TTYY$, USOX$, 49, 12)
         Call REPLA(TTYY$, PDX$, 61, 4)
         Call REPLA(TTYY$, FORMATNUM(CAUNIDS, "F9.4"), 89, 9)
         Call REPLA(TTYY$, FORMATNUM(PESOMET, "F7.3"), 101, 7)
         Call REPLA(TTYY$, FORMATNUM(LARCORTE, "F6.1"), 111, 6)
         Call REPLA(TTYY$, FORMATNUM(USOUNI, "F8.4"), 119, 8)
         CAIT% = CAIT% + 1
         EXPLOLIN$(CAIT%) = TTYY$
         '
       Next UKU&
       Call EXPLOGR(CODEXT$(CDXI%), 1)
     End If
   Next CDXI%
   '
   ATENTI% = ATENTIA%
   '
End Sub
'
Sub IMPORUTS()
   '
   Dim PUNSTRU$(32767)
   '
   Screen.MousePointer = 11
   FIN& = List4.ListCount
   NMS% = NUMAS%
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     COCON$ = TRIM$(Left$(TTXX$, 16))
     CDXI% = CODINT%(COCON$)
     '
     If CDXI% > 0 Then
       If CDXI% <= NMS% Then
         PUNSTRU$(CDXI%) = PUNSTRU$(CDXI%) + MKS$(U&)
       End If
     End If
   Next U&
   '
   '
   Call ABRESECOP
   '
   REBLA$ = REGBLAN$("SECOPER")
   For CDXI% = 1 To NMS%
     If Len(PUNSTRU$(CDXI%)) > 0 Then
       With Secoper
63       .FindFirst "CODICONJ =" & CDXI%
         If .NoMatch = False Then
           .Delete
           .FindNext "CODICONJ =" & CDXI%
           GoTo 63
         End If
         '
         NOPEI% = 0
         For UKU& = 1 To Len(PUNSTRU$(CDXI%)) Step 4
           U& = CVS(Mid$(PUNSTRU$(CDXI%), UKU&, 4))
           TTXX$ = List4.List(U& - 1)
           '
           SQLX$ = "SELECT * FROM Secoper "
           Dim SECOPER1 As ADODB.Recordset
           Set SECOPER1 = New ADODB.Recordset
           SECOPER1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
           With SECOPER1
           '
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
            !CODICONJ = CDXI%
            !CODXCONJ = CODEXT$(CDXI%)
              NOPEI% = NOPEI% + 10
            !NumeOper = NOPEI%
            !DescOper = Mid$(TTXX$, 23, 48)
            !statoper = 1
            !CodMaq0 = Mid$(TTXX$, 73, 6)
            !CodMaq1 = Mid$(TTXX$, 139, 6)
            !CodMaq2 = Mid$(TTXX$, 147, 6)
            !CodMaq3 = Mid$(TTXX$, 155, 6)
            !HorsPrep = XVALO(Mid$(TTXX$, 79, 10))
            !HorsFijo = XVALO(Mid$(TTXX$, 89, 10))
            !MinuStan = XVALO(Mid$(TTXX$, 99, 12))
            !HorsLimp = 0
            !HorsMHer = 0
            !PieCiclo = 0
            !PorDemor = XVALO(Mid$(TTXX$, 111, 10))
            If !MinuStan > 0.005 Then
              !Cadencia = 60 / !MinuStan
            End If
            !CantiOps = XVALO(Mid$(TTXX$, 129, 8))
            !CategOps = 0
            !CoefSolp = 0
            !LOTESTAN = XVALO(Mid$(TTXX$, 121, 8))
            !ProgCNum = Mid$(TTXX$, 163, 6)
            !HoProFab = " "
            !ProTerOp = 0
            !PreUnid = 0
            !MonePrec = 0
            !FechPrec = CVDAT(0)
            '
            !OBSEOPER = " "
            !NUOPER_PREP = 1
            !CATOPER_PREP = 0
            .Update
           End With
           '
         Next UKU&
         
        ARTICULOS.Close
        Set ARTICULOS = Nothing
        
       End With
     End If
   Next CDXI%
   '
   Screen.MousePointer = 1
   '
End Sub
'
Sub IMPORMAQ()
   '
   FIN& = List4.ListCount
   REBLA$ = REGBLAN$("PADMAQ")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     COMAQ$ = TRIM$(Left$(TTXX$, 6))
     DEMAQ$ = TRIM$(Mid$(TTXX$, 9, 30))
     CUHOR = XVALO(Mid$(TTXX$, 41, 12))
     SECCI = XVALO(Mid$(TTXX$, 53, 10))
     MAQEQ$ = TRIM$(Mid$(TTXX$, 66, 6))
     COEFI% = XVALO(Mid$(TTXX$, 72, 12))
     OBSER$ = TRIM$(Mid$(TTXX$, 87, 30))
     '
     REMAQ$ = REBLA$
     Call REPLA(REMAQ$, COMAQ$, 1, 6)
     Call REPLA(REMAQ$, DEMAQ$, 7, 30)
     Call REPLA(REMAQ$, MKS$(SECCI), 37, 4)
     Call REPLA(REMAQ$, MAQEQ$, 43, 6)
     Call REPLA(REMAQ$, MKI$(COEFI), 49, 2)
     Call REPLA(REMAQ$, OBSER$, 51, 30)
     Call REPLA(REMAQ$, MKS$(CUHOR), 93, 4)
     Call FILTERSETRECORD("PADMAQ", 1, COMAQ$, REMAQ$) ' sobreescribe SI IGUAL CODIGO
     '
   Next U&
   Call CIERRARCH("PADMAQ")
   '
End Sub

Sub IMPORLISPRE()
   '
   Call SELECHO("LISTAS")
   LPRE% = XVALO(VALACT1$("LISTAS"))
   If LPRE% < 1 Then Exit Sub
   '
   ARCHIPRE$ = TRIM$(Str$(LPRE%))
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   '
   If ULTREG&(ARCHIPRE$) > 0 Then
     If COMALTER%("La Lista de Precios Elegida Contiene Información.\¿ Desea Sobre-Escribirla o Cancelar Importación ?\\Cancelar Importación\Sobre-Escribir Lista") <> 2 Then
       Exit Sub
     End If
   End If
   '
   Dim ListaPrecio As ADODB.Recordset
   Set ListaPrecio = New ADODB.Recordset
   
   ListaPrecio.Open FILTSQL$("SELECT * FROM PREVENTA"), FLEXCONN, adOpenKeyset, adLockPessimistic
   
   
   FIN& = List4.ListCount
   JJ& = 0
   'REBLA$ = REGBLAN$(ARCHIPRE$)
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     COLIS% = XVALO(Left$(TTXX$, 4))
     If COLIS% = LPRE% Then
       CODE$ = TRIM$(Mid$(TTXX$, 6, 16))
       CIXI% = CODINT%(CODE$)
       If CIXI% > 0 Then
         DEDE$ = TRIM$(Mid$(TTXX$, 24, 64))
         If DEDE$ = "" Then DEDE$ = DESCRIT0$(CIXI%)
         NORI$ = TRIM$(Mid$(TTXX$, 90, 16))
         PRUNI# = XVALO(Mid$(TTXX$, 108, 12))
         '
         'TTYY$ = REBLA$
         'Call REPLA(TTYY$, MKI$(CIXI%), 1, 2)
         'Call REPLA(TTYY$, DEDE$, 3, 42)
         'Call REPLA(TTYY$, NORI$, 46, 16)
        ' Call REPLA(TTYY$, MKD$(PRUNI#), 62, 8)
         'JJ& = JJ& + 1
         'Call GRAREG(ARCHIPRE$, TTYY$, JJ&)
         '
         With ListaPrecio
             .AddNew   ' YA CONVERTIDO
             On Error Resume Next
             !CodiEmpr = CodiEmp%
             On Error GoTo 0
             !Cod_Lista = LPRE%
             !Cod_Pieza = CIXI%
             !Num_Orden = U&
             !Desc_Pieza = DESCRIT0$(CIXI%)
             !Denom_Pieza = DESCRIT0$(CIXI%)
             !Num_Original = NORI$
             !Precio_Unitario = PRUNI#
             !Cant_Limite = 0
             !Porc_Limite = 0
             .Update
         End With
         
       End If
     End If
   Next U&
   '
   ListaPrecio.Close
   Set ListaPrecio = Nothing
   'Call SETULTREG(ARCHIPRE$, JJ&)
   'Call CIERRARCH(ARCHIPRE$)
   '
End Sub
'
Sub IMPORCAT()
   '
   FIN& = List4.ListCount
   REBLA$ = REGBLAN$("CATOPER")
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     ICAT% = XVALO(Left$(TTXX$, 4)) Mod 256
     DECAT$ = TRIM$(Mid$(TTXX$, 7, 30))
     CUHOR = XVALO(Mid$(TTXX$, 37, 12))
     OBSER$ = TRIM$(Mid$(TTXX$, 51, 32))
     '
     REMAQ$ = REBLA$
     Call REPLA(REMAQ$, Chr$(ICAT%), 1, 1)
     Call REPLA(REMAQ$, DECAT$, 2, 27)
     Call REPLA(REMAQ$, MKS$(CUHOR), 29, 4)
     Call REPLA(REMAQ$, OBSER$, 33, 32)
     JJ& = JJ& + 1
     Call GRAREG("CATOPER", REMAQ$, JJ&)
   Next U&
   '
   Call SETULTREG("CATOPER", JJ&)
   Call CIERRARCH("CATOPER")
   '
End Sub

Sub IMPORCLI()
   '
   LONGINT% = 0
   PUNUCLI$ = String$(32767, 0)
   XX0$ = REGBLAN$("DEUDOR1") + REGBLAN$("DEUDOR2")
   '
   ' Genera DEUDOR1 y DEUDOR2
   Call ACLI_GES07.GENECOCLI
   '
'   ' Establece en el "vector" PUNUCLI$, los numeros de clientes ya utilizados
'   For U& = 1 To ULTREG&("DEUDOR1")
'     XX$ = REGLEIDO$("DEUDOR1", U&)
'     NCI% = CVI(Left$(XX$, 2))
'     If NCI% > 0 Then
'       If NCI% <= 32767 Then
'         Mid$(PUNUCLI$, NCI%, 1) = Chr$(1)
'       End If
'     End If
'   Next U&
'   '
'   hrow = DbGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
'   DbGrid1.RowHeight = 10
'   DBHANT = DbGrid1.Height
'   URECORD& = Data1.Recordset.RecordCount
'   DbGrid1.Height = (URECORD& + 10) * DbGrid1.RowHeight
   '
   Call APERBASDAT("CLIEN")
   FIN& = List4.ListCount
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     '
     NCI& = XVALO(Mid$(TTXX$, 1, 6))
     If NCI& < 1 Then GoTo 19
     If NCI& > 32000 Then LONGINT% = 1
     ' Verifica si el numero ya ha sido utilizado
     While Len(PUNUCLI$) < NCI&: PUNUCLI$ = PUNUCLI$ + Chr$(0): Wend
     If Asc(Mid$(PUNUCLI$, NCI&, 1)) > 0 Then GoTo 19
     '
     RASO$ = Mid$(TTXX$, 8, 47)
     DOMI$ = Mid$(TTXX$, 55, 48)
     CPOSTX$ = Mid$(TTXX$, 103, 12)
     CPOST% = 0: If XVALO(CPOSTX$) < 32767 Then CPOST% = XVALO(CPOSTX$)
     LOCA$ = Mid$(TTXX$, 115, 32)
     TELE$ = Mid$(TTXX$, 147, 24)
     FAXX$ = Mid$(TTXX$, 171, 24)
     CONTACTX$ = Mid$(TTXX$, 195, 24)
     CUIT$ = TRIM$(Mid$(TTXX$, 219, 24))
       If CUIT$ <> "" Then VCC% = VALICUIT%(CUIT$)
     GruCli% = XVALO(Mid$(TTXX$, 244, 4))
     CODICUEN$ = TRIM$(Mid$(TTXX$, 249, 12))
        If CODICUEN$ = "" Then
           CODICUEN$ = "CLI-" + TRIM$(Str$(NCI&))
        End If
     '\\\OT-05-0192-WAR-18/04/05///
     CONDI% = XVALO(Mid$(TTXX$, 261, 4))
     PIVA% = XVALO(Mid$(TTXX$, 269, 4))
     PROVI$ = TRIM$(Mid$(TTXX$, 267, 2))
     LIPRE% = 0
     COVEN% = XVALO(Mid$(TTXX$, 273, 4))
     LICRE = XVALO(Mid$(TTXX$, 277, 12))
     CUEMA% = 0
     '\\\OT-5-632-FG-21/09/05///
     DESSTAN$ = Mid$(TTXX$, 291, 24)
     LUGENT$ = Mid$(TTXX$, 317, 48)
     LOCALI$ = Mid$(TTXX$, 365, 48)
     NTRAN& = XVALO(Mid$(TTXX$, 413, 5))
     TRANSPOR$ = Mid$(TTXX$, 421, 40)
     GRUDIV% = XVALO(Mid$(TTXX$, 461, 8))
     EMAIX$ = Mid$(TTXX$, 475, 36)
     POREMBA = XVALO(Mid$(TTXX$, 511, 6))
     '\\\OT-5-632-FG-FIN///
     'GoTo 19 'PARA TESTEAR
     '\\\OT-05-0192-WAR-FIN///
     INIX& = U& - 1
     PAISX$ = "": LPRX% = 0: WEBX$ = ""
     URECORD& = DbGrid1.Rows
12   For U2& = INIX& To URECORD& 'Data1.Recordset.RecordCount
        DbGrid1.Row = U2&
        DbGrid1.Col = 0
        NCLX& = XVALO(DbGrid1.TEXT) 'XVALO(Data1.Recordset.Fields(0).Value)
        If NCLX& = NCI& Then
           DbGrid1.Col = 5: PAISX$ = DbGrid1.TEXT
           DbGrid1.Col = 15: LPRX% = XVALO(DbGrid1.TEXT)
           DbGrid1.Col = 26: WEBX$ = LCase$(DbGrid1.TEXT)
           GoTo 15
        End If
     Next U2&
     If INIX& > 0 Then
        INIX& = 0
        GoTo 12
     End If
     '
15   On Error GoTo 0
'     With Deudor12
     SQLX$ = "SELECT * FROM DEUDOR12"
     Dim DEUDDDOR12 As ADODB.Recordset
     Set DEUDDDOR12 = New ADODB.Recordset
     DEUDDDOR12.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     With DEUDDDOR12
       .AddNew   ' YA CONVERTIDO
       On Error Resume Next
       !CodiEmpr = CodiEmp%
       On Error GoTo 0
       !Stat_Cli = 1                        'STATUS ACTIVO
       !Nume_Cli = NCI&                     'CODIGO ENTERO
       !CODI_CLI = CODICUEN$                'CODIGO TEXT
       !Tcpr_Cli = 0                        'TIPO CTA
       !GrCo_Cli = GruCli%                  'TIPO CLI  VA A GRUPO CLI
       !Raso_Cli = RASO$                    'RAZON SOCIAL
       !Fant_Cli = Left$(RASO$, 10)         'FANTASIA
       !Domi_Cli = DOMI$                    'DIRECCION
       !Loca_Cli = LOCA$                    'LOCALIDAD
       !Prov_Cli = PROVI$                  'COD PROVINCIA
       !Pais_Cli = PAISX$                 'PAIS VALIDADO
       !Cpos_Cli = CPOSTX$                  'CODIGO POSTAL
       !Tele_Cli = TELE$                    'TELEFONO
       '!Tele2_Cli = TELEF2$                '2º TELEFONO
       !Faxi_Cli = FAXX$                    'FAX
       !Cuit_Cli = CUIT$                    'CUIT SIN VALIDAR
       !Cpag_Cli = COVEN%                   'CONDICION DE PAGO
       !Lpre_Cli = LPRX%
       '!Nibr_Cli = INGBRU$                 'IMP BRUTOS
       !Piva_Cli = PIVA%                    'POS IVA
       '!Posivat_Cli = PIVATXT$             'POSICION DEL IVA EN TEXTO
       '!C_Ret = CRET$                      'RETENCION
       !Lcre_Cli = LICRE                    'LIMITE CREDITO
       !Vend_Cli = CONDI%                   'CODIGO VENDEDOR
       !Ctac_Cli = CONTACTX$                'CONTACTO
       !Dsta_Cli = DESSTAN$                 'DESCUENTO
       !Lentre0_Cli = LUGENT$               'LUGAR DE ENTREGA
       !Lentre1_Cli = LOCALI$               'LOCALIDAD
       !Lentre2_Cli = TRANSPOR$             'TRANSPORTE
       !NTrans_Cli = NTRAN&                 'NUMERO DE TRANSPORTISTA
       !Gru_Div = GRUDIV%                   'GRUPO DIVISIONAL
       !Mail_Cli = EMAIX$                   'DIRECCION DE MAIL
       !Recar_Embal = POREMBA               'PORCENTAJE RECARGO EMBALAJE
    .Update
     End With
     '
     If NCI& > 0 Then
        Mid$(PUNUCLI$, NCI&, 1) = Chr$(1)
     End If
     '
19 Next U&
   DEUDDDOR12.Close
   Set DEUDDDOR12 = Nothing
   '
   UPDATCLI% = 1
   If LONGINT% > 0 Then
      Call COPYSTRU("IBOLPEN", "DEUDORL1")
   End If
   Call ACLI_GES07.GENECOCLI(LONGINT%)
   '
   '\\\OT-5-632-FG-21/09/05///
   Call ACLI_GES07.GENDAPIBCLI
   FIN& = List4.ListCount
   For U& = 1 To FIN&
     TTXX$ = List4.List(U& - 1)
     NCI& = XVALO(Mid$(TTXX$, 1, 6))
     PERIB = XVALO(Mid$(TTXX$, 461, 4))
     '
'     X$ = FILTERGETRECORD$("DAPIBCLI", 1, MKI$(NCI&))
     Call REPLA(X$, Chr$(PERIB), 3, 1)
'     Call FILTERPUTRECORD("DAPIBCLI", 1, MKI$(NCI%), X$)
20 Next U&
   '
'   DbGrid1.RowHeight = hrow   ' RESTAURA ALTURA ANTERIOR
   '\\\OT-5-632-FG-FIN///
End Sub

Sub IMPORPRO()
   '
   XX0$ = REGBLAN$("PROVED1") + REGBLAN$("PROVED2")
   '
   Call APERBASDAT("PROVE")
   '
   FIN& = List4.ListCount
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     '
     NCI% = XVALO(Mid$(TTXX$, 1, 5))
     If NCI% < 1 Then GoTo 19
     CODICUEN$ = TRIM$(Mid$(TTXX$, 240, 12))
     If CODICUEN$ = "" Then CODICUEN$ = "PR-" + TRIM$(Str$(NCI%))
     If TRIM$(CODICUEN$) = "" Then GoTo 19
     '
     RASO$ = Mid$(TTXX$, 7, 48)
     DOMI$ = Mid$(TTXX$, 55, 48)
     CPOSTX$ = Mid$(TTXX$, 103, 12)
     CPOST% = 0: If XVALO(CPOSTX$) < 32767 Then CPOST% = XVALO(CPOSTX$)
     LOCA$ = Mid$(TTXX$, 115, 32)
     TELE$ = Mid$(TTXX$, 147, 24)
     FAXX$ = Mid$(TTXX$, 171, 24)
     CONTACTX$ = Mid$(TTXX$, 195, 24)
     CUIT$ = TRIM$(Mid$(TTXX$, 219, 16))
       If CUIT$ <> "" Then VCC% = VALICUIT%(CUIT$)
     TIPOCTA% = XVALO(Mid$(TTXX$, 235, 4))
     CODIPROVI1$ = Mid$(TTXX$, 253, 2)
     CODIVA1% = XVALO(Mid$(TTXX$, 255, 4))
     CTAMADRE$ = TRIM$(Mid$(TTXX$, 261, 12))
     FANTASI$ = TRIM$(Mid$(TTXX$, 275, 16))
     'PIVA% = 1
     IBRU$ = ""
     CONDI% = 0
     LIPRE% = 0
     COVEN% = 0
     LICRE = 0
     CUEMA% = 0
     '
     
'     With Proved12
'       .FindFirst "Codi_Cli = '" & CODICUEN$ & "'"
'       If .NoMatch = False Then
     SQLX$ = "SELECT * FROM Proved12 "
     SQLX$ = SQLX$ + " WHERE Codi_Cli = '" & CODICUEN$ & "'"
     Dim PROVVVED As ADODB.Recordset
     Set PROVVVED = New ADODB.Recordset
     PROVVVED.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     With PROVVVED
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
           On Error GoTo 0
           OPX% = COMALTER%("Cuenta '" + TRIM$(CODICUEN$) + "' Existente en B.Datos\\Actualizar\Conservar\Cancelar Proceso")
           If OPX% < 1 Or OPX% > 2 Then GoTo 99
           If OPX% = 2 Then GoTo 19
         Else
            On Error GoTo 0
           .AddNew   ' YA CONVERTIDO
           On Error Resume Next
           !CodiEmpr = CodiEmp%
           On Error GoTo 0
         !Nume_Cli = NCI%                     'CODIGO ENTERO
       End If
       !Stat_Cli = 1                        'STATUS ACTIVO
       !CODI_CLI = CODICUEN$                'CODIGO TEXT
       !Tcpr_Cli = TIPOCTA%                      'TIPO CTA
       !Raso_Cli = RASO$                    'RAZON SOCIAL
       !Fant_Cli = FANTASI$                 'FANTASIA
       !Domi_Cli = DOMI$                    'DIRECCION
       !Loca_Cli = LOCA$                    'LOCALIDAD
       !Prov_Cli = CODIPROVI1$              'COD PROVINCIA
       '!Pais_Cli = PAISOK$                 'PAIS VALIDADO
       !Cpos_Cli = CPOSTX$                  'CODIGO POSTAL
       !Tele_Cli = TELE$                    'TELEFONO
       '!Tele2_Cli = TELEF2$                '2º TELEFONO
       !Faxi_Cli = FAXX$                    'FAX
       !Cuit_Cli = CUIT$                    'CUIT SIN VALIDAR
       '!Cpag_Cli = 0                       'POR DEFAULT (XQ' NO SE JUSTIFICAN 60/900)
       '!Nibr_Cli = INGBRU$                 'IMP BRUTOS
       '!Mail_Cli = MAILPRO$                'MAIL
       !Piva_Cli = CODIVA1%                    'POS IVA
       '!Posivat_Cli = PIVATXT$             'POSICION DEL IVA EN TEXTO
       '!C_Ret = CRET$                      'RETENCION
       !Ctac_Cli = CONTACTX$
       CMINT% = CUECOINT%(CTAMADRE$)
       !CuMa_Cli = CMINT%                   'CUENTA MADRE
       .Update
     End With
     '
19 Next U&
   PROVVVED.Close
   Set PROVVVED = Nothing
   '
   '\\\OT-5-632-FG-21/09/05///
   UPDATPRO% = 1
   Call ACRE_GES07.GENECOPROV    ' PARA ACTUALIZAR LAS LISTAS DE SELECCIóN
   '
   Call ACRE_GES07.GENDARIBPRO
   FIN& = List4.ListCount
   For U& = 1 To FIN&
     TTXX$ = List4.List(U& - 1)
     NPI% = XVALO(Mid$(TTXX$, 1, 5))
     RETIB = XVALO(Mid$(TTXX$, 293, 4))
     '
     X$ = FILTERGETRECORD$("DARIBPRO", 1, MKI$(NPI%))
     Call REPLA(X$, Chr$(RETIB), 3, 1)
     Call FILTERPUTRECORD("DARIBPRO", 1, MKI$(NPI%), X$)
20 Next U&
   '
   '\\\OT-5-632-FG-FIN///
   '
99 On Error Resume Next
   'BadaproVE.Close
   On Error GoTo 0
   '
End Sub

