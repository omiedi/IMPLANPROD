VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form MARCH_GES07 
   BackColor       =   &H00B9DFEB&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A. Maestros"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1470
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ListBox SORTLIST 
      Height          =   255
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   24
      Top             =   3885
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00395F7D&
      Height          =   4335
      Left            =   4680
      ScaleHeight     =   4275
      ScaleWidth      =   915
      TabIndex        =   18
      Top             =   0
      Width           =   975
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H00624E28&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   30
         Top             =   3600
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00E2D3C0&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   31
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "QC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   105
         Picture         =   "MARCH_GES07.frx":0000
         TabIndex        =   22
         ToolTipText     =   "Sub-Sistema de Gestión de Calidad"
         Top             =   1680
         Width           =   650
      End
      Begin VB.CommandButton Command21 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   105
         Picture         =   "MARCH_GES07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Importación / Conversión de Datos"
         Top             =   2415
         Width           =   645
      End
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   100
         Picture         =   "MARCH_GES07.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   960
         Width           =   650
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   100
         Picture         =   "MARCH_GES07.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   315
         Width           =   650
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   105
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00B9DFEB&
      Caption         =   "ACTUALIZACIÓN BASE DE DATOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4425
      Begin VB.Frame Frame4 
         BackColor       =   &H00B9DFEB&
         Caption         =   "ARCHIVOS MAESTROS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   210
         TabIndex        =   9
         Top             =   360
         Width           =   4005
         Begin VB.CommandButton Command6 
            Caption         =   "Command6"
            Height          =   255
            Left            =   3480
            TabIndex        =   26
            Top             =   480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   255
            Left            =   3480
            TabIndex        =   25
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   840
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   10
            Top             =   420
            Width           =   2325
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
               Height          =   630
               Left            =   0
               Picture         =   "MARCH_GES07.frx":06E8
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Maestro de Artículos  de Stock"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   705
               Left            =   700
               TabIndex        =   29
               Top             =   100
               Width           =   1755
            End
         End
      End
      Begin VB.Frame Frame5 
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
         TabIndex        =   2
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command4 
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
            TabIndex        =   8
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command5 
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
            TabIndex        =   7
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command10 
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
            TabIndex        =   6
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command11 
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
         Begin VB.CommandButton Command12 
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
            TabIndex        =   4
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command13 
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
            TabIndex        =   3
            Top             =   1890
            Width           =   1590
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00B9DFEB&
         Caption         =   "CLIENTES / PROVEEDORES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   210
         TabIndex        =   1
         Top             =   1890
         Width           =   4005
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   15
            Top             =   840
            Width           =   2325
            Begin VB.CommandButton Command18 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   -20
               Picture         =   "MARCH_GES07.frx":09F2
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Clientes en Cuenta    Corriente"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   720
               TabIndex        =   17
               Top             =   100
               Width           =   1800
            End
         End
         Begin VB.PictureBox Picture10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   630
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   12
            Top             =   420
            Width           =   2325
            Begin VB.CommandButton Command19 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "MARCH_GES07.frx":0E34
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Proveedores y         Subcontratistas"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   705
               Left            =   700
               TabIndex        =   14
               Top             =   0
               Width           =   1755
            End
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "MARCH_GES07.frx":113E
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "MARCH_GES07.frx":1372
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu Archi 
      Caption         =   "Archivos Maestros"
      Begin VB.Menu Mastok 
         Caption         =   "Materiales, Fórmulas, Rutas"
      End
      Begin VB.Menu PLH1 
         Caption         =   "-"
      End
      Begin VB.Menu Amacli 
         Caption         =   "Clientes - Cuentas Corrientes"
      End
      Begin VB.Menu Proveed 
         Caption         =   "Proveedores / Contratistas"
      End
      Begin VB.Menu PLH2 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
End
Attribute VB_Name = "MARCH_GES07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CUEBALRECI%, CUEBALPAGO%, CUEBALAJUS%, CUEBALCART%
Dim INTERRUMPE%


Sub LIMPIAPDOCSPL()
    '
    'BARRIDA DE ARCHIVO PDOCLSPEL.DAT O DEL DT0 REALIZANDO LIMPIEZA DE LAS IMAGENES
    'QUE YA ESTAN GRABADAS EN LA TABLA
    '
    Call COMUNI("Conversion Primaria Finalizada.\   \Para Completar el Proceso\TODOS los Usuarios Deben\Salir del Sistema.")
    Call BLOQARCH("PDOCSPL")
    Call BLOQARCH("PDOCLST")
    '
10  If EXISTE%(LUDAT$ + "PDOCSPL.DT1") > 0 Then
       Call FILEKILL(LUDAT$ + "PDOCSPL.DT1")
       GoTo 10
    End If
    '
    '1.-Recorro el  el pdocclst
    Dim TD0 As String * 512
    NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DT0", 0, 512)
    LOFI = LOF(NX0%)
    URE& = LOFI / 512
    If LOFI < 0 Then URE& = 4194304
    '
    NX1% = FILEOPEN%(LUDAT$ + "PDOCSPL.DT1", 0, 512)
    JJ& = 0: DOCUAN& = 0
    '
    For L& = 1 To URE&
       Call TRACE(20, L&, URE&)
       Get #NX0%, L&, TD0$
       DOCUNU& = CVS(Left$(TD0$, 4))
       If DOCUNU& = DOCUAN& Then GoTo 15     ' REPITE LO ANTERIOR
       '
       'VALIDO SI EXISTE EL REGISTRO Y LA IMAGEN EN LA TABLA SQL --
       DOCUAN& = DOCUNU&
       Condi$ = "NUMEDOC =" & DOCUNU&
       IMGTABLA$ = VALOBADA("PDOC_TBL", "IMAGEN", Condi$, "NOMESS")
       '
15     If IMGTABLA$ = "" Then
          JJ& = JJ& + 1
          Put #NX1%, JJ&, TD0$
       End If
    Next L&
    '
    Close #NX0%
    Close #NX1%
    Call CIERRARCH("*.*")
    Close
    '
    Call FILEKILL(LUDAT$ + "PDOCSPL.DT0")
    If JJ& > 0 Then
       CPF% = COPYFILE(LUDAT$ + "PDOCSPL.DT1", LUDAT$ + "PDOCSPL.DT0")
    End If
    '
    '
20  If EXISTE%(LUDAT$ + "PDOCSPL.DT1") > 0 Then
       Call FILEKILL(LUDAT$ + "PDOCSPL.DT1")
       GoTo 20
    End If
    '
    NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DAT", 0, 512)
    LOFI = LOF(NX0%)
    URE& = LOFI / 512
    If LOFI < 0 Then URE& = 4194304
    '
    NX1% = FILEOPEN%(LUDAT$ + "PDOCSPL.DT1", 0, 512)
    JJ& = 0: DOCUAN& = 0
    '
    For L& = 1 To URE&
       Call TRACE(20, L&, URE&)
       Get #NX0%, L&, TD0$
       DOCUNU& = CVS(Left$(TD0$, 4))
       If DOCUNU& = DOCUAN& Then GoTo 25     ' REPITE LO ANTERIOR
       '
       'VALIDO SI EXISTE EL REGISTRO Y LA IMAGEN EN LA TABLA SQL --
       DOCUAN& = DOCUNU&
       Condi$ = "NUMEDOC =" & DOCUNU&
       IMGTABLA$ = VALOBADA("PDOC_TBL", "IMAGEN", Condi$, "NOMESS")
       '
25     If IMGTABLA$ = "" Then
          JJ& = JJ& + 1
          Put #NX1%, JJ&, TD0$
       End If
    Next L&
    '
    Close #NX0%
    Close #NX1%
    Call CIERRARCH("*.*")
    Close
    '
    Call FILEKILL(LUDAT$ + "PDOCSPL.DAT")
    If JJ& > 0 Then
       CPF% = COPYFILE(LUDAT$ + "PDOCSPL.DT1", LUDAT$ + "PDOCSPL.DAT")
    End If
    '
    Call CIERRARCH("PDOCLST")
    Call COMUNI("FIN DEL PROCESO")
    Call FINAL("")
    '
End Sub
'
Private Sub Amacli_Click()
   Call Command18_Click
End Sub

Private Sub BARRATRAZA_DblClick()
   INTERRUMPE% = 1
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   If EXISTE%("AMASTO07.EXE") > 0 Then
       Call CONECRUN("AMASTO07", "Maestro de Items de Stock")
     Else
       AMAS_GES07.Show 1
   End If
   Command1.Enabled = True
End Sub

Private Sub Command18_Click()
    Command18.Enabled = False
     If DERACCE%("AMACLI", "Maestro de Clientes") > 0 Then
'        If ACLI_GES07.Visible = True Then
'           ACLI_GES07.Hide
'        End If
'        ACLI_GES07.Show 1
        If EXISTE%("AMACLI09.EXE") > 0 Then
           Call CONECRUN("AMACLI09", "Archivo Maestro de Clientes"): GoTo 99
        End If
        If APLINVO$ <> "" Then
            ACLI_GES07.Show 1
        Else
            Me.Enabled = False
            ACLI_GES07.Show
            ACLI_GES07.SetFocus
        End If
     End If
99  Command18.Enabled = True
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    If DERACCE%("AMAPRO", "Maestro de Proveedores") > 0 Then
       'ACRE_GES07.Show 1
        If EXISTE%("AMAPRO09.EXE") > 0 Then Call CONECRUN("AMAPRO09", "Archivo Maestro de Proveedores"): GoTo 99
        If APLINVO$ <> "" Then
            ACRE_GES07.Show 1
        Else
            Me.Enabled = False
            ACRE_GES07.Show
            ACRE_GES07.SetFocus
        End If
    End If
    'Call CONECRUN("AMAPRO", "Actualización de Base de Datos de Proveedores")
99  Command19.Enabled = True
End Sub

Private Sub Command2_Click()
 Command2.Enabled = False
    If EXISTE%("ISO9K07.EXE") > 0 Then
       Call CONECRUN("ISO9K07", "SUB-SISTEMA DE GESTION DE CALIDAD")
    End If
    Command2.Enabled = True
End Sub

Private Sub Command21_Click()
    Command21.Enabled = False
10  RTA% = ALTERNA%("Importación desde Planillas Excel.\Conversión de 'DAT' a 'MDB'.\Conversion de Documentos Gráficos\Proceso en Tabla de Documentos")
    If RTA% = 1 Then
          IMPOR08.Show 1
      ElseIf RTA% = 2 Then
          Call CONVBADA 'CONVERSION DE .DAT A .MDB
      ElseIf RTA% = 3 Then
          Call DESCDOCGRA ' DESCARGA DE DOCUMENTOS GRAFICOS
     ElseIf RTA% = 4 Then
          Call PROCEDOCU ' DESCARGA DE DOCUMENTOS GRAFICOS
      Else
          GoTo 99
    End If
    GoTo 10
    '
99  Command21.Enabled = True
End Sub

Sub DESCDOCGRA()

    If CLACONTRA% < 1 Then                      ' iNGRESO DE CLAVE Y CONTRASEÑA
       Call MENSERR(24, "Santo Y Seña Inválido")
       Exit Sub
    End If
    
    Image2.ToolTipText = "Doble-Click para Interrumpir Proceso"
    '
    '1.-CREO LOS CAMPOS
    Call ABRECONEXION
    Call CREACAMPO("", "PDOC_TBL", "NUMEDOC", 4, 0, 1)
    Call CREACAMPO("", "PDOC_TBL", "Referencia", 10, 128)
    Call CREACAMPO("", "PDOC_TBL", "FechayHora", 8, 0)
    Call CREACAMPO("", "PDOC_TBL", "Usuario", 10, 64)
    Call CREACAMPO("", "PDOC_TBL", "Terminal", 10, 64)
    Call CREACAMPO("", "PDOC_TBL", "Impresora", 10, 64)
    Call CREACAMPO("", "PDOC_TBL", "Imagen", 11, 0)
    '
    ULTIDOC& = XVALO(VALOBADA("PDOC_TBL", "MAX(NUMEDOC)", "NUMEDOC > 0"))
    INIDOC& = ULTIDOC& - 128
    If INIDOC& < 1 Then INIDOC& = 1
    '
    If INIDOC& > 1 Then
       OPX% = ALTERNA%("Re-Comenzar\Continuar Desde Registro " & INIDOC&)
       If OPX% <> 2 Then INIDOC& = 1
    End If
    '
'' BORRA LOS REGISTROS ANTERIORES
'SQLX$ = "SELECT NUMEDOC FROM PUPDATE PDOC_TBL SET IMAGEN = '' WHERE NUMEDOC < 10000"
'FLEXCONN.Execute SQLX$, RAFE&
'MsgBox RAFE&

    INTERRUMPE% = 0
    Screen.MousePointer = 11
    '2.-Recorro el  el pdocclst
    FIN& = ULTREG&("PDOCLST")
    For i& = INIDOC& To FIN&
       If INTERRUMPE% = 1 Then
          If COMALTER%("¿ Realmente Quiere Interrumpir el Proceso ?\\No, Continuar\Si, Interrumpir") = 2 Then
             GoTo 90
          End If
       End If
       '
       Call TRACE(20, i&, FIN&)
       IMAG2$ = ""
       TX$ = REGLEIDO$("PDOCLST", i&)
       
       'LEO LOS VALORES DE PDOCLST
       NUMDOC1& = CVS(Left$(TX$, 4))
       '
       'VALIDO SI EXISTE EL REGISTRO Y LA IMAGEN EN LA TABLA SQL --
       Condi$ = "NUMEDOC =" & NUMDOC1&
       If CantRegistros("PDOC_TBL", Condi$) > 1 Then
           Call BORRAREGISTROS("PDOC_TBL", Condi$ + " AND (IMAGEN IS NULL OR IMAGEN LIKE '')", RAFE&, "NOMESS")
AAA = BBB
       End If
       '
       IMGTABLA$ = VALOBADA("PDOC_TBL", "IMAGEN", Condi$, "NOMESS")
       If IMGTABLA$ <> "" Then GoTo 50 '(NEXT)
       '
       'SI NO EXISTE SIGO LEYENDO
       REFER$ = TRIM$(Mid$(TX$, 5, 60))
         FE% = CVI(Mid(TX$, 65, 2))
         HORX% = CVI(Mid(TX$, 67, 2))
       FEYHO = FETEXCOR1$(FE%) & " " & HORATEX$(HORX%)
       USUA$ = TRIM$(Mid(TX$, 69, 30))
       TERMI$ = TRIM$(Mid(TX$, 99, 30))
       IMPRE$ = TRIM$(Mid(TX$, 129, 128))
       IMAG2$ = IMAGENPDOC$(NUMDOC1&) 'IMAGEN DE PDOCSPL
       '
       '3.GRABO EN SQL SI EL REGISTRO NO EXISTE
       On Error Resume Next
       If CantRegistros("PDOC_TBL", Condi$) > 1 Then
           Call BORRAREGISTROS("PDOC_TBL", Condi$, RAFE&, "NOMESS")
       End If
       '
       If CantRegistros("PDOC_TBL", Condi$) < 1 Then
           Call CreaRegistro("PDOC_TBL", "NUMEDOC", NUMDOC1&, "NOMESS/NOAUDIT")
       End If
       Call ACTUREGISTRO("PDOC_TBL", "REFERENCIA", Condi$, REFER$, REGAF&, "NOMESS/NOAUDIT")
       Call ACTUREGISTRO("PDOC_TBL", "FECHAYHORA", Condi$, FEYHO, REGAF&, "NOMESS/NOAUDIT")
       Call ACTUREGISTRO("PDOC_TBL", "USUARIO", Condi$, USUA$, REGAF&, "NOMESS/NOAUDIT")
       Call ACTUREGISTRO("PDOC_TBL", "TERMINAL", Condi$, TERMI$, REGAF&, "NOMESS/NOAUDIT")
       Call ACTUREGISTRO("PDOC_TBL", "IMPRESORA", Condi$, IMPRE$, REGAF&, "NOMESS/NOAUDIT")
       Call ACTUREGISTRO("PDOC_TBL", "IMAGEN", Condi$, IMAG2$, REGAF&, "NOMESS/NOAUDIT/RETRY")
       If Err Then
          On Error GoTo 0
          TTXX$ = "Error no Capturado al Procesar Comprobante " & NUMDOC1&
          If COMALTER%(TTXX$ + "\\Continuar Proceso\Interrumpir") = 2 Then GoTo 90
       End If
       On Error GoTo 0
50  Next i&
    '
    ' RECORRE EL ARCHIVO Y LIMPIA LAS IMAGENES QUE YA ESTAN GRABADAS EN LA TABLA
90  Call CIERRARCH("PDOCLST")
    Call LIMPIAPDOCSPL
    Screen.MousePointer = 1
    '
    Call COMUNI("Conversion Finalizada")

End Sub
'
Sub PROCEDOCU()

    'CREO LOS CAMPOS
    '***************
    Call ABRECONEXION
    Call CREACAMPO("", "PDOC_TBL", "TIPODOCU", 10, 32, 1)
    Call CREACAMPO("", "PDOC_TBL", "NUMECOMPLE", 10, 32, 1)
    Call CREACAMPO("", "PDOC_TBL", "NUMEBUSQUE", 10, 16, 1)
    '
    Call PoneEnCeroNull("PDOC_TBL", "TIPODOCU", REGAF&, "NOMESS")
    Call PoneEnCeroNull("PDOC_TBL", "NUMECOMPLE", REGAF&, "NOMESS")
    Call PoneEnCeroNull("PDOC_TBL", "NUMEBUSQUE", REGAF&, "NOMESS")
    '
    i& = 0
    'RECORRO EL PDOC_TBL
    Condi$ = "TIPODOCU = ''"
    SQLX$ = "SELECT * FROM PDOC_TBL WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & Condi$
    SQLX$ = SQLX$ + " ORDER BY NUMEDOC"
    '
    FIN& = CantRegistros&("PDOC_TBL", Condi$)
    '
    Set RSPDOC = READSET(SQLX$)
    With RSPDOC
      Do While Not .EOF
         i& = i& + 1
         Call TRACE(20, i&, FIN&)
 
         TIPOX$ = ""
         NUCOMP$ = ""
         REFE$ = TRIM$(UCase$(SAFETEXT$(!REFERENCIA)))
         numeroDocumento& = XVALO(!NUMEDOC)
         POS1% = 0
         POS2% = 0
         If Left(REFE$, Len("AJUSTE DE STOCKS")) = "AJUSTE DE STOCKS" Then
            'Ajuste de Stocks 10000
            TIPOX$ = "AJUSTE DE STOCKS"
            
         ElseIf Left$(REFE$, Len("AJUSTE SALDO CLIENTE")) = "AJUSTE SALDO CLIENTE" Then
            'Ajuste Saldo Cliente AS-101019-03 - RONALD FREDY IRIARTE - 19/10/10
            TIPOX$ = "AJUSTE SALDO CLIENTE"
         
         ElseIf Left$(REFE$, Len("ANALISIS DE COSTO")) = "ANALISIS DE COSTO" Then
            'ANALISIS DE COSTO 0014 - 17/01/11
            TIPOX$ = "ANALISIS DE COSTO"
         
         ElseIf Left$(REFE$, Len("BAJA DE NUMERO DE SERIE")) = "BAJA DE NUMERO DE SERIE" Then
            'BAJA DE NUMERO DE SERIE 889040742170
            TIPOX$ = "BAJA DE NUMERO DE SERIE"
            
         ElseIf Left$(REFE$, Len("BOLETA DE RECEPCIÓN")) = "BOLETA DE RECEPCIÓN" Then
            'BOLETA DE RECEPCIÓN  1 - PROVEEDOR DE PRUEBA - 02/10/08
            TIPOX$ = "BOLETA DE RECEPCIÓN"
           
         ElseIf Left$(REFE$, Len("COMPROBANTE DE DIARIO")) = "COMPROBANTE DE DIARIO" Then
           ' COMPROBANTE DE DIARIO 90107-001 - 07/01/09
            TIPOX$ = "COMPROBANTE DE DIARIO"
            
         ElseIf Left$(REFE$, Len("DEPOSITO BANCARIO")) = "DEPOSITO BANCARIO" Then
           'DEPOSITO BANCARIO DB-101126-00 - 26/11/10
            TIPOX$ = "DEPOSITO BANCARIO"
         
          ElseIf Left$(REFE$, Len("DESPACHO DE MATERIALES")) = "DESPACHO DE MATERIALES" Then
            'Despacho de Materiales DM.30765 - SIMTECH LTDA - 01/10/10
            TIPOX$ = "DESPACHO DE MATERIALES"
         
         ElseIf Left$(REFE$, Len("ETIQUETA DE EXPEDICIÓN")) = "ETIQUETA DE EXPEDICIÓN" Then
            'ETIQUETA DE EXPEDICIÓN 07441 - 04/11/10
            TIPOX$ = "ETIQUETA DE EXPEDICIÓN"
            
         ElseIf Left$(REFE$, Len("FACTURA NRO.")) = "FACTURA NRO." Then
            'FACTURA NRO.0001-00001250 - CLIENTE DE PRUEBA - 02/10/08
            TIPOX$ = "FACTURA NRO."
         ElseIf Left$(REFE$, Len("NOTA DE CREDITO")) = "NOTA DE CREDITO" Then
            'NOTA DE CREDITO 0001-00000156 - 3 WAY SOLUTION SA - 15/04/11
            TIPOX$ = "NOTA DE CREDITO"
            
         ElseIf Left$(REFE$, Len("NOTA DE PEDIDO")) = "NOTA DE PEDIDO" Then
            'NOTA DE PEDIDO 002136-1 - COELAR S.R.L. - 24 DE ENERO DE 2011
            TIPOX$ = "NOTA DE PEDIDO"
            
         ElseIf Left$(REFE$, Len("NOTA DE DEBITO")) = "NOTA DE DEBITO" Then
            'Nota de Debito 0001-00033662 - NUSA DUA SA - 15/09/10
            TIPOX$ = "NOTA DE DEBITO"
 
         ElseIf Left$(REFE$, Len("FUGA A TIERRA")) = "FUGA A TIERRA" Then
            'FUGA A TIERRA 2429123456-0
            TIPOX$ = "FUGA A TIERRA"
         ElseIf Left$(REFE$, Len("ORDEN DE FABRICACIÓN")) = "ORDEN DE FABRICACIÓN" Then
            'ORDEN DE FABRICACIÓN OF-022349-1
            TIPOX$ = "ORDEN DE FABRICACIÓN"
         
         ElseIf Left$(REFE$, Len("ORDEN DE PAGO NRO")) = "ORDEN DE PAGO NRO" Then
           'ORDEN DE PAGO NRO. 2 - WAGNER JAVIER - 03/10/08
            TIPOX$ = "ORDEN DE PAGO NRO."
            
         ElseIf Left$(REFE$, Len("ORDEN DE DESPACHO")) = "ORDEN DE DESPACHO" Then
            'ORDEN DE DESPACHO 002726 - PVR SA - 18/01/11
            TIPOX$ = "ORDEN DE DESPACHO"
            
         ElseIf Left$(REFE$, Len("ORDEN DE SERVICIO")) = "ORDEN DE SERVICIO" Then
            'ORDEN DE SERVICIO 460 - WAGNER JAVIER FERNANDO - 20/09/10
            TIPOX$ = "ORDEN DE SERVICIO"
            
         ElseIf Left$(REFE$, Len("PEDIDO INTERNO")) = "PEDIDO INTERNO" Then
            'PEDIDO INTERNO 000001 - CLIENTE DE PRUEBA - 02/10/08
            TIPOX$ = "PEDIDO INTERNO"
            
         ElseIf Left$(REFE$, Len("PEDIDO DE CLIENTE")) = "PEDIDO DE CLIENTE" Then
            'PEDIDO DE CLIENTE 002727 - PVR SA - 18/01/11
            TIPOX$ = "PEDIDO DE CLIENTE"
         
         ElseIf Left$(REFE$, Len("PLANILLA DE REPARACIÓN")) = "PLANILLA DE REPARACIÓN" Then
           'Planilla de Reparación 1 - ADAMOLI VALERIA-CLEMENTE INMACULA
            TIPOX$ = "PLANILLA DE REPARACIÓN"
   
         ElseIf Left$(REFE$, Len("PRESUP.REPARACIÓN EXTERNO")) = "PRESUP.REPARACIÓN EXTERNO" Then
            'Presup.Reparación Externo 19 - CENTURION CRISTALDO MARIA ESTHER - 17/03/11
            TIPOX$ = "PRESUP.REPARACIÓN EXTERNO"
 
         ElseIf Left$(REFE$, Len("PRESUPUESTO DE REPARACIÓN")) = "PRESUPUESTO DE REPARACIÓN" Then
            'Presupuesto de Reparación 27 - 4 DE FEBRERO S.R.L. - 04/04/11
            TIPOX$ = "PRESUPUESTO DE REPARACIÓN"
            
         ElseIf Left$(REFE$, Len("PRESUPUESTO")) = "PRESUPUESTO" Then
            'Presupuesto 74 - OBRA SOCIAL DEL PERSONAL DE TV - 14/09/09
            TIPOX$ = "PRESUPUESTO"
            
         ElseIf Left$(REFE$, Len("PUESTA A TIERRA")) = "PUESTA A TIERRA" Then
            'PUESTA A TIERRA 2429123456-0
            TIPOX$ = "PUESTA A TIERRA"
            
         ElseIf Left$(REFE$, Len("RECIBO NRO.")) = "RECIBO NRO." Then
           'Recibo Nro. 29038 - KAIDAR MARIEL FABIANA - 19/10/10
            TIPOX$ = "RECIBO NRO."
          
         ElseIf Left$(REFE$, Len("RECHAZO EN LÍNEA")) = "RECHAZO EN LÍNEA" Then
            'RECHAZO EN LÍNEA 002077 - 02/11/10
            TIPOX$ = "RECHAZO EN LÍNEA"
          
         ElseIf Left$(REFE$, Len("REGISTRO DE CONTROL FINAL")) = "REGISTRO DE CONTROL FINAL" Then
            'REGISTRO DE CONTROL FINAL 586312100014-0
            TIPOX$ = "REGISTRO DE CONTROL FINAL"
            
         ElseIf Left$(REFE$, Len("REGISTRO DE MEDICION")) = "REGISTRO DE MEDICION" Then
            'REGISTRO DE MEDICION 2429123456-0
            TIPOX$ = "REGISTRO DE MEDICION"
          
         ElseIf Left$(REFE$, Len("REMITO DE DEVOLUCIÓN")) = "REMITO DE DEVOLUCIÓN" Then
            'Remito de Devolución RD.1281 - ACTION TRAVEL - 13/10/10
            TIPOX$ = "REMITO DE DEVOLUCIÓN"
          
         ElseIf Left$(REFE$, Len("REMITO DE TRASLADO")) = "REMITO DE TRASLADO" Then
            'Remito de Traslado RC-751 - ROBERTO PERSIANI - 06/09/10
            TIPOX$ = "REMITO DE TRASLADO"
            
         ElseIf Left$(REFE$, Len("REMITO NRO.")) = "REMITO NRO." Then
            'Remito Nro.0001-00040261 - JUAN JOSE LOVERA E HIJOS - 14/09/10
            TIPOX$ = "REMITO NRO."
            
         ElseIf Left$(REFE$, Len("TRANSFERENCIA DE STOCK")) = "TRANSFERENCIA DE STOCK" Then
            'Transferencia de Stock 486
            TIPOX$ = "TRANSFERENCIA DE STOCK"
          
         ElseIf Left$(REFE$, Len("VALE DE CONSUMO DE MATERIALES")) = "VALE DE CONSUMO DE MATERIALES" Then
            'Vale de Consumo de Materiales VM.1144 - 20/09/10
            TIPOX$ = "VALE DE CONSUMO DE MATERIALES"
         
         ElseIf Left$(REFE$, Len("VALE DE MATERIALES")) = "VALE DE MATERIALES" Then
            'Vale de Materiales 023801 - 10/01/11
            TIPOX$ = "VALE DE MATERIALES"
         
         ElseIf Left$(REFE$, Len("COMPENSACIÓN DE CUENTAS")) = "COMPENSACIÓN DE CUENTAS" Then
            'Vale de Materiales 023801 - 10/01/11
            TIPOX$ = "COMPENSACIÓN DE CUENTAS"

         End If
         '
         POS1% = Len(TIPOX$)
         If POS1% > 0 Then
            POS1% = POS1% + 1
            TX$ = LTrim(Mid$(REFE$, POS1%))
            POS2% = InStr(TX$, " ")
            NUCOMP$ = TRIM$(Mid$(TX$, 1, POS2%))
            If POS2% < 1 Then NUCOMP$ = TRIM$(Mid$(REFE, POS1%))
            '
            NUBUSQ$ = NUBUSQUE$(NUCOMP$, TIPOX$)
            
            'PARA EL CASO DE FACTURA, REMITO, RECIBO QUE TERMINAN CON 'NRO.'
            If Right$(TIPOX$, 4) = "NRO." Then TIPOX$ = TRIM$(Mid$(TIPOX$, 1, Len(TIPOX$) - 4))
            '
            Condi$ = "NUMEDOC =" & numeroDocumento&
            Call ACTUREGISTRO("PDOC_TBL", "TIPODOCU", Condi$, TIPOX$, REGAF&, "NOMESS")
            Call ACTUREGISTRO("PDOC_TBL", "NUMECOMPLE", Condi$, NUCOMP$, REGAF&, "NOMESS")
            Call ACTUREGISTRO("PDOC_TBL", "NUMEBUSQUE", Condi$, NUBUSQ$, REGAF&, "NOMESS")
 
         End If
         '
         .MoveNext
      Loop
    End With
    '
    RSPDOC.Close
    Set RSPDOC = Nothing
    
    

End Sub
Function IMAGENPDOC$(DOCUNU&)

    Dim TD0 As String * 512
    TPSP$ = ""
    NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DAT", 0, 512)
    Get #NX0%, 1, TD0$
    If CVS(Left$(TD0$, 4)) > DOCUNU& Then
      Close #NX0%
      NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DT0", 0, 512)
    End If
    '
    LOFI = LOF(NX0%)
    URE& = LOFI / 512
    If LOFI < 0 Then URE& = 4194304
    '
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    Get #NX0%, L&, TD0$
    DCC& = CVS(Left$(TD0$, 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
'For KKII% = 1 To 24: Get #NX0%, KKII%, TD0$: TTXX$ = TTXX$ + Str$(CVS(Left$(TD0$, 4))): Next KKII%: MsgBox TTXX$
7   Close #NX0%
    IMAGENPDOC$ = ""
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
    IMAGENPDOC$ = TPSP$
    '
99  End Function

Private Sub Command24_Click()
    Call COPYSTRU("PROVED1", "SELPRBX") 'Se agrego por si no existe el archivo
    Call FRESHECHO("!SELPRBX")    'Se agrego para limpiar el archivo
    Call SETULTREG("!SELPRBX", 0) 'de busqueda por CUIT, Tel, Direccio, ETC
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call FINAL("")
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("ARCHIMA")
End Sub

Sub MIGRACARCHE()
   
   Dim IDECH$()
   HOII% = HOY(HOS$)
   CUECARTE% = CUECOINT%(Control$("", "CUECHECA"))
   '
   FIN& = ULTREG&("BADACHE")
   ReDim IDECH$(FIN&)
   UBACHE& = FIN&
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("BADACHE", U&)
      IDEVA$ = TRIM$(Left$(XX$, 40))
      IDECH$(U&) = IDEVA$
   Next U&
   
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI' "
   'Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   Dim CHETEMP1 As ADODB.Recordset
   Set CHETEMP1 = New ADODB.Recordset
   CHETEMP1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With CHETEMP1
      On Error Resume Next
      .MoveLast
      FIN& = .RecordCount
      .MoveFirst
      If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           U& = .AbsolutePosition
           Call TRACE(20, U&, FIN&)
           '.Edit
           !VariMovim = "1"
           .Update
         .MoveNext
         Loop
      End If
   End With
   On Error GoTo 0
   CHETEMP1.Close
   Set CHETEMP1 = Nothing
   '
   FIN& = ULTREG&("BADACHE")
   ReDim IDECH$(FIN&)
   UBACHE& = FIN&
   Label1.ToolTipText = "Hasta " & FIN&
   For U& = 1 To FIN&
      DIFECART# = 0
      Label1 = U&
      Label1.Visible = True
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("BADACHE", U&)
      IDEVA$ = TRIM$(Left$(XX$, 40))
         IDECH$(U&) = IDEVA$
      Import# = CVD(Mid$(XX$, 49, 8))
      FEMICHE% = CVI(Mid$(XX$, 57, 2))
      FEVECHE% = CVI(Mid$(XX$, 59, 2))
      FRECEP% = CVI(Mid$(XX$, 61, 2))
      DORECEP$ = TRIM$(Mid$(XX$, 63, 10))
      ORICHE$ = TRIM$(Mid$(XX$, 73, 30))
      DESTI$ = TRIM$(Mid$(XX$, 103, 30))
      FENTREGA% = CVI(Mid$(XX$, 133, 2))
      OB1$ = TRIM$(Mid$(XX$, 135, 40))
      OB2$ = TRIM$(Mid$(XX$, 175, 40))
      OB3$ = TRIM$(Mid$(XX$, 215, 40))
      '
      ' BUSCA POR IDEVA$
      If IDEVA$ = "" Then GoTo 19
15    SQLZ$ = "SELECT * FROM CAJABANC "
      SQLZ$ = SQLZ$ + "WHERE TipoMovim = 'CHRECI' "
      SQLZ$ = SQLZ$ + "AND DescriMov = '" & IDEVA$ & "' and VariMovim = '1' "
      'Set CHETEMP3 = CueCorri.OpenRecordset(SQLZ$, dbOpenDynaset)
      Dim CHETEMP3 As ADODB.Recordset
      Set CHETEMP3 = New ADODB.Recordset
      CHETEMP3.Open FILTSQL$(SQLZ$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      '
      With CHETEMP3
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
              On Error GoTo 0
              Call REPLA(XX$, "**", 255, 2)
              Call GRAREG("BADACHE", XX$, U&)
              '.Edit
              !VariMovim = "2"
              .Update
              GoTo 19    ' LO ENCONTRO
            Else
              On Error GoTo 0
              CAGRE% = CAGRE% + 1
              GoSub AGREGAENTRACHE
              GoTo 15
         End If
      End With
      CHETEMP3.Close
      Set CHETEMP3 = Nothing
19 Next U&
   '
   SQLY$ = "SELECT * FROM CAJABANC "
   SQLY$ = SQLY$ + "WHERE TipoMovim = 'CHENTR' ORDER BY DESCRIMOV ASC"
   'Set CHETEMP2 = CueCorri.OpenRecordset(SQLY$, dbOpenDynaset)
   Dim CHETEMP2 As ADODB.Recordset
   Set CHETEMP2 = New ADODB.Recordset
   CHETEMP2.Open FILTSQL$(SQLY$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With CHETEMP2
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
         On Error GoTo 0
         FINIX& = .RecordCount
         JKU& = 0
         .MoveFirst
         Do While Not .EOF
           JKU& = JKU& + 1
           Call TRACE(20, JKU&, FINIX&)
           '.Edit
           !VariMovim = "1"
           .Update
         .MoveNext
         Loop
      End If
   End With
   On Error GoTo 0
   CHETEMP2.Close
   Set CHETEMP2 = Nothing
   '
   Label1.ToolTipText = "Hasta " & FIN&
   For U& = 1 To FIN&
      DIFECART# = 0
      Label1 = U&
      Label1.Visible = True
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("BADACHE", U&)
      IDEVA$ = TRIM$(Left$(XX$, 40))
         IDECH$(U&) = IDEVA$
      Import# = CVD(Mid$(XX$, 49, 8))
      FEMICHE% = CVI(Mid$(XX$, 57, 2))
      FEVECHE% = CVI(Mid$(XX$, 59, 2))
      FRECEP% = CVI(Mid$(XX$, 61, 2))
      DORECEP$ = TRIM$(Mid$(XX$, 63, 10))
      ORICHE$ = TRIM$(Mid$(XX$, 73, 30))
      DESTI$ = TRIM$(Mid$(XX$, 103, 30))
      FENTREGA% = CVI(Mid$(XX$, 133, 2))
      OB1$ = TRIM$(Mid$(XX$, 135, 40))
      OB2$ = TRIM$(Mid$(XX$, 175, 40))
      OB3$ = TRIM$(Mid$(XX$, 215, 40))
      '
      If FENTREGA% > 0 Then
        ' BUSCA POR IDEVA$
        If IDEVA$ = "" Then GoTo 39
25      SQLZ$ = "SELECT * FROM CAJABANC "
        SQLZ$ = SQLZ$ + "WHERE TipoMovim = 'CHENTR' "
        SQLZ$ = SQLZ$ + "AND DescriMov = '" & IDEVA$ & "' and VariMovim = '1' "
        'Set CHETEMP3 = CueCorri.OpenRecordset(SQLZ$, dbOpenDynaset)
        'Dim CHETEMP3 As ADODB.Recordset
        Set CHETEMP3 = New ADODB.Recordset
        CHETEMP3.Open FILTSQL$(SQLZ$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        '
        With CHETEMP3
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
              On Error GoTo 0
              '.Edit
              !VariMovim = "2"
              .Update
              GoTo 29    ' LO ENCONTRO
            Else
              On Error GoTo 0
              GoSub AGREGASALIDA
              GoTo 25
          End If
        End With
        CHETEMP3.Close
        Set CHETEMP3 = Nothing
        '
29      SQLZ$ = "SELECT * FROM CAJABANC "
        SQLZ$ = SQLZ$ + "WHERE TipoMovim = 'CHRECI' "
        SQLZ$ = SQLZ$ + "AND DescriMov = '" & IDEVA$ & "' and VariMovim = '2' "
        'Set CHETEMP3 = CueCorri.OpenRecordset(SQLZ$, dbOpenDynaset)
        'Dim CHETEMP3 As ADODB.Recordset
        Set CHETEMP3 = New ADODB.Recordset
        CHETEMP3.Open FILTSQL$(SQLZ$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        '
        With CHETEMP3
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
              On Error GoTo 0
              '.Edit
              !VariMovim = "3"
              !Status = 2
              !FEDEPOSENT = CVDAT(FENTREGA%)
              !FECHENTRE = CVDAT(FENTREGA%)
              If TRIM$(" " & !Document) = "" Then !Document = "Mig.Datos"
              If TRIM$(" " & !DESTICHEQ) = "" Then !DESTICHEQ = DESTI$
              If TRIM$(" " & !REFESALCHE) = "" Then !REFESALCHE = OB1$
              If TRIM$(" " & !IDENSALCHE) = "" Then !IDENSALCHE = OB2$ ' AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
              If TRIM$(" " & !UsuaSalChe) = "" Then !UsuaSalChe = OB3$ ' AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
              .Update
          End If
        End With
        CHETEMP3.Close
        Set CHETEMP3 = Nothing
      End If
39 Next U&
   '
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
   SQLX$ = SQLX$ + "AND Status = 1 "
   'Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   'Dim CHETEMP1 As ADODB.Recordset
   Set CHETEMP1 = New ADODB.Recordset
   CHETEMP1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With CHETEMP1
      On Error Resume Next
      .MoveLast
      If Err Then
         On Error GoTo 0
         GoTo 59
      End If
      On Error GoTo 0
      '
      FIN& = .RecordCount
      .MoveFirst
      Do While Not .EOF
        U& = .AbsolutePosition
        Call TRACE(20, U&, FIN&)
        IDEVA$ = !DescriMov
        For JJ& = 1 To UBACHE&
          If IDECH$(JJ&) = IDEVA$ Then
            XX$ = REGLEIDO$("BADACHE", JJ&)
            Import# = CVD(Mid$(XX$, 49, 8))
            FEMICHE% = CVI(Mid$(XX$, 57, 2))
            FEVECHE% = CVI(Mid$(XX$, 59, 2))
            FRECEP% = CVI(Mid$(XX$, 61, 2))
            DORECEP$ = TRIM$(Mid$(XX$, 63, 10))
            ORICHE$ = TRIM$(Mid$(XX$, 73, 30))
            DESTI$ = TRIM$(Mid$(XX$, 103, 30))
            FENTREGA% = CVI(Mid$(XX$, 133, 2))
            OB1$ = TRIM$(Mid$(XX$, 135, 40))
            OB2$ = TRIM$(Mid$(XX$, 175, 40))
            OB3$ = TRIM$(Mid$(XX$, 215, 40))
            '
            If FENTREGA% > 1 Then
                GoTo 45
              Else
                GoTo 49
            End If
          End If
        Next JJ&
        '
        ' no lo encontró
        FENTREGA% = CVINT(!FECHCONTA)     ' SALE EN LA MISMA FECHA DE ENTRADA
        DESTI$ = "Baja por Consistencia"
        OB1$ = "No ingresado en BADACHE.DAT"
        OB2$ = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
        OB3$ = AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
        ' fin de no lo encontró
        '
45      '.Edit
        !Status = 2
        !VariMovim = "3"
        !FEDEPOSENT = CVDAT(FENTREGA%)
        !FECHENTRE = CVDAT(FENTREGA%)
        If TRIM$(" " & !Document) = "" Then !Document = "Mig.Datos"
        If TRIM$(" " & !DESTICHEQ) = "" Then !DESTICHEQ = DESTI$
        If TRIM$(" " & !REFESALCHE) = "" Then !REFESALCHE = OB1$
        If TRIM$(" " & !IDENSALCHE) = "" Then !IDENSALCHE = OB2$ ' AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
        If TRIM$(" " & !UsuaSalChe) = "" Then !UsuaSalChe = OB3$ ' AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
        .Update
49    .MoveNext
      Loop
59 End With
   CHETEMP1.Close
   Set CHETEMP1 = Nothing
Exit Sub
'
AGREGAENTRACHE:
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI' "
    Set CHETEMP1 = New ADODB.Recordset
    CHETEMP1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With CHETEMP1
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
      !CodiCom_Cor = DORECEP$
      !codicom_lar = DORECEP$
      !RefeCom = ORICHE$
      !FECHCONTA = CVDAT(FRECEP%)
      !FechAcred = CVDAT(FEVECHE%)
      !CUECONTAI = CUECARTE%
      !CenCosto = ""
      !ItemConc = ""
      !NUCLIEN = 0
      !NuProve = 0
      !TipoMovim = "CHRECI"
      !OpciMovim = "CH"
      !VariMovim = "1"
      !DescriMov = IDEVA$
      !FECHEMISI = CVDAT(FEMICHE%)
      !CodiEmpr = 0
      !CodiSucu = 0
      !NumeTalo = 0
      !NumeComp = 0
      !Status = 1
        On Error Resume Next
      !Mone_Emis = 1
      !Tipo_Cam = 1
        On Error GoTo 0
      !ValAbsComp = Abs(Import#)
      !Idebecomp = Import#
      !IHABECOMP = 0
      !Impo_Dola = Import#
      !BASEIMPO = Abs(Import#)
      !NROCAI = ""
        On Error Resume Next
      !IncSubIva = ""
        On Error GoTo 0
      !FEVTOCAI = CVDAT(0)
      !ObserGen = ""
      !IdenAlta = Left$(IDACTU$, 8)          ' FECHA - HORA - USUARIO - TERMINAL
      !IdenBaja = String$(6, 0)
      !FeHoReal = Now
      !NuAsiDiar = ""
     .Update
     '
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CodiEmpr = CodiEmp%
     On Error GoTo 0
      !CodiCom_Cor = DORECEP$
      !codicom_lar = DORECEP$
      !RefeCom = ORICHE$
      !FECHCONTA = CVDAT(FRECEP%)
      !FechAcred = CVDAT(FEVECHE%)
      !CUECONTAI = CUEBALCART%
      !CenCosto = ""
      !ItemConc = ""
      !NUCLIEN = 0
      !NuProve = 0
      !TipoMovim = "MIGRA"
      !OpciMovim = ""
      !VariMovim = ""
      !DescriMov = IDEVA$
      !FECHEMISI = CVDAT(FEMICHE%)
      !CodiEmpr = 0
      !CodiSucu = 0
      !NumeTalo = 0
      !NumeComp = 0
      !Status = 0
        On Error Resume Next
      !Mone_Emis = 1
      !Tipo_Cam = 1
        On Error GoTo 0
      !ValAbsComp = Import#
      !Idebecomp = 0
      !IHABECOMP = Import#
      !Impo_Dola = (-1) * Import#
      !BASEIMPO = Abs(Import#)
      !NROCAI = ""
        On Error Resume Next
      !IncSubIva = ""
        On Error GoTo 0
      !FEVTOCAI = CVDAT(0)
      !ObserGen = ""
      !IdenAlta = Left$(IDACTU$, 8)          ' FECHA - HORA - USUARIO - TERMINAL
      !IdenBaja = String$(6, 0)
      !FeHoReal = Now
      !NuAsiDiar = ""
      .Update
   End With
   CHETEMP1.Close
   Set CHETEMP1 = Nothing
Return
'
'
AGREGASALIDA:
    SQLY$ = "SELECT * FROM CAJABANC "
    SQLY$ = SQLY$ + "WHERE TipoMovim = 'CHENTR' ORDER BY DESCRIMOV ASC"
    Set CHETEMP2 = New ADODB.Recordset
    CHETEMP2.Open FILTSQL$(SQLY$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With CHETEMP2
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
      !CodiCom_Cor = "MIGRA"
      !codicom_lar = "MIGRA"
      !RefeCom = DESTI$
      !FECHCONTA = CVDAT(FENTREGA%)
      !FechAcred = CVDAT(FENTREGA%)
      !CUECONTAI = CUECARTE%
      !CenCosto = ""
      !ItemConc = ""
      !NUCLIEN = 0
      !NuProve = 0
      !TipoMovim = "CHENTR"
      !OpciMovim = "CH"
      !VariMovim = "1"
      !DescriMov = IDEVA$
      !FECHEMISI = CVDAT(FEMICHE%)
      !CodiEmpr = 0
      !CodiSucu = 0
      !NumeTalo = 0
      !NumeComp = 0
      !Status = 0
        On Error Resume Next
      !Mone_Emis = 1
      !Tipo_Cam = 1
        On Error GoTo 0
      !ValAbsComp = Abs(Import#)
      !Idebecomp = 0
      !IHABECOMP = Import#
      !Impo_Dola = (-1) * Import#
      !BASEIMPO = Abs(Import#)
      !NROCAI = ""
        On Error Resume Next
      !IncSubIva = ""
        On Error GoTo 0
      !FEVTOCAI = CVDAT(0)
      !ObserGen = ""
      !IdenAlta = Left$(IDACTU$, 8)          ' FECHA - HORA - USUARIO - TERMINAL
      !IdenBaja = String$(6, 0)
      !FeHoReal = Now
      !NuAsiDiar = ""
     .Update
     '
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CodiEmpr = CodiEmp%
     On Error GoTo 0
      !CodiCom_Cor = "MIGRA"
      !codicom_lar = "MIGRA"
      !RefeCom = DESTI$
      !FECHCONTA = CVDAT(FENTREGA%)
      !FechAcred = CVDAT(FENTREGA%)
      !CUECONTAI = CUEBALCART%
      !CenCosto = ""
      !ItemConc = ""
      !NUCLIEN = 0
      !NuProve = 0
      !TipoMovim = "MIGRA"
      !OpciMovim = ""
      !VariMovim = ""
      !DescriMov = IDEVA$
      !FECHEMISI = CVDAT(FEMICHE%)
      !CodiEmpr = 0
      !CodiSucu = 0
      !NumeTalo = 0
      !NumeComp = 0
      !Status = 0
        On Error Resume Next
      !Mone_Emis = 1
      !Tipo_Cam = 1
        On Error GoTo 0
      !ValAbsComp = Import#
      !Idebecomp = Import#
      !IHABECOMP = 0
      !Impo_Dola = Import#
      !BASEIMPO = Abs(Import#)
      !NROCAI = ""
        On Error Resume Next
      !IncSubIva = ""
        On Error GoTo 0
      !FEVTOCAI = CVDAT(0)
      !ObserGen = ""
      !IdenAlta = Left$(IDACTU$, 8)          ' FECHA - HORA - USUARIO - TERMINAL
      !IdenBaja = String$(6, 0)
      !FeHoReal = Now
      !NuAsiDiar = ""
      .Update
   End With
   CHETEMP2.Close
   Set CHETEMP2 = Nothing
Return
'
'
End Sub


Private Sub Command3_Click()
   Call ABRECONEXION
   Call CREACAMPO("", "PRUEBA", "NUMERO", 4, 0, 1)
   Call CREACAMPO("", "PRUEBA", "IMAGEN", 12, 0)
   TTXX$ = ""
   For i% = 0 To 255
      TTXX$ = TTXX$ + Chr$(i%)
   Next i%
   Call CreaRegistro("PRUEBA", "NUMERO", 1)
   Call ACTUREGISTRO("PRUEBA", "IMAGEN", "NUMERO = 1", TTXX$, RAFE&)
End Sub

Private Sub Command6_Click()
   Call ABRECONEXION
   FIN& = CantRegistros("PDOC_TBL")
   Dim RS As ADODB.Recordset
   Set NUDOC = New ADODB.Recordset
   Set NUDOC = FLEXCONN.Execute("SELECT * FROM PDOC_TBL ORDER BY NUMEDOC ASC")
   JJ& = 0
   With NUDOC
      Do While Not .EOF
      JJ& = JJ& + 1
      Call TRACE(20, JJ&, FIN&)
      On Error Resume Next
      SULON = SULON + Len(CStr(!IMAGEN))
      On Error GoTo 0
      .MoveNext
      Loop
   End With
   NUDOC.Close
   MsgBox SULON
End Sub

Private Sub exit_Click()
   Call Command24_Click
End Sub

Private Sub Form_Load()
   '
   Me.Visible = False
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   Call CONAPLI(APLINVO$, TERMINA%) ' Busca la aplicacion a mostrar
   If TERMINA% > 0 Then Call FINAL("")
   '
   EJER% = 1
   Call TRANSLABELS(Name)
   '
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me, Picture1)
   '
   Me.Visible = True
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   If EXISTE%("ISO9K07.EXE") = 0 Then 'Verifica si existe el modulo de calidad
     Command2.Visible = False
   End If
   '
   Call CREATABLAPROVINCIAS
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
End Sub

Private Sub Help_Click()
   Call Command29_Click
End Sub

Private Sub Image2_DblClick()
   INTERRUMPE% = 1
End Sub

Private Sub Label11_DblClick()
   If COMALTER%("Realmente Desea Convertir\los Números de Proveedores ?\\No, Continuar\Si, Convertir") = 2 Then
     OPX% = CLACONTRA%
     If OPX% = 1 Then
       Screen.MousePointer = 11
10     Call APERBASDAT("MASTER")            ' AQUI ABRIR BASE
       With Master                ' AQUI PONER TABLA
         On Error Resume Next
         .MoveFirst
         If Err Then
           On Error GoTo 0
           GoTo 20
         End If
         On Error GoTo 0
         Do While Not .EOF
           PRORI% = !PROVHABI     ' AQUI AJUSTAR VARIABLE
           PRNUE% = REPNU%(PRORI%)
           If PRNUE% <> PRORI% Then
             .Edit
             !PROVHABI = PRNUE%
             .Update
           End If
         .MoveNext
         Loop
       End With
       '
20     Call APERBASDAT("COMPRAS")            ' AQUI ABRIR BASE
       With OCOMDETA               ' AQUI PONER TABLA
         On Error Resume Next
         .MoveFirst
         If Err Then
           On Error GoTo 0
           GoTo 30
         End If
         On Error GoTo 0
         Do While Not .EOF
           PRORI% = !Npro_Ocom     ' AQUI AJUSTAR VARIABLE
           PRNUE% = REPNU%(PRORI%)
           If PRNUE% <> PRORI% Then
             .Edit
             !Npro_Ocom = PRNUE%   ' AQUI AJUSTAR VARIABLE
             .Update
           End If
         .MoveNext
         Loop
       End With
       '
30     Call APERBASDAT("COMPRAS")            ' AQUI ABRIR BASE
       With OCOMENCA               ' AQUI PONER TABLA
         On Error Resume Next
         .MoveFirst
         If Err Then
           On Error GoTo 0
           GoTo 40                 ' aqui al proximo bloque
         End If
         On Error GoTo 0
         Do While Not .EOF
           PRORI% = !Npro_Ocom     ' AQUI AJUSTAR VARIABLE
           PRNUE% = REPNU%(PRORI%)
           If PRNUE% <> PRORI% Then
             .Edit
             !Npro_Ocom = PRNUE%   ' AQUI AJUSTAR VARIABLE
             .Update
           End If
         .MoveNext
         Loop
       End With
       '
40     Call APERBASDAT("STOCKS")            ' AQUI ABRIR BASE
       With Bolrecep              ' AQUI PONER TABLA
         On Error Resume Next
         .MoveFirst
         If Err Then
           On Error GoTo 0
           GoTo 50                 ' aqui al proximo bloque
         End If
         On Error GoTo 0
         Do While Not .EOF
           PRORI% = !NumProv       ' AQUI AJUSTAR VARIABLE
           PRNUE% = REPNU%(PRORI%)
           If PRNUE% <> PRORI% Then
             .Edit
             !NumProv = PRNUE%     ' AQUI AJUSTAR VARIABLE
             .Update
           End If
         .MoveNext
         Loop
       End With
       '
50     Call APERBASDAT("STOCKS")            ' AQUI ABRIR BASE
       With Movisto               ' AQUI PONER TABLA
         On Error Resume Next
         .MoveFirst
         If Err Then
           On Error GoTo 0
           GoTo 60                 ' aqui al proximo bloque
         End If
         On Error GoTo 0
         Do While Not .EOF
           PRORI% = !Nume_Prov     ' AQUI AJUSTAR VARIABLE
           PRNUE% = REPNU%(PRORI%)
           If PRNUE% <> PRORI% Then
             .Edit
             !Nume_Prov = PRNUE%   ' AQUI AJUSTAR VARIABLE
             .Update
           End If
         .MoveNext
         Loop
       End With
       '
       Screen.MousePointer = 1
       '
     End If
   End If
   '
60 Call COMUNI("Conversion Completa")
   '
End Sub

Function REPNU%(NORIG%)
   '
   RPN% = NORIG%
   '
   If NORIG% = 1 Then RPN% = 10: GoTo 90
   If NORIG% = 2 Then RPN% = 34: GoTo 90
   If NORIG% = 3 Then RPN% = 61: GoTo 90
   If NORIG% = 4 Then RPN% = 64: GoTo 90
   If NORIG% = 5 Then RPN% = 66: GoTo 90
   If NORIG% = 6 Then RPN% = 68: GoTo 90
   If NORIG% = 7 Then RPN% = 96: GoTo 90
   If NORIG% = 8 Then RPN% = 98: GoTo 90
   If NORIG% = 9 Then RPN% = 295: GoTo 90
   If NORIG% = 10 Then RPN% = 304: GoTo 90
   If NORIG% = 11 Then RPN% = 308: GoTo 90
   If NORIG% = 12 Then RPN% = 333: GoTo 90
   If NORIG% = 13 Then RPN% = 425: GoTo 90
   If NORIG% = 14 Then RPN% = 665: GoTo 90
   If NORIG% = 15 Then RPN% = 858: GoTo 90
   If NORIG% = 16 Then RPN% = 1014: GoTo 90
   If NORIG% = 17 Then RPN% = 985: GoTo 90
   If NORIG% = 18 Then RPN% = 138: GoTo 90
   If NORIG% = 19 Then RPN% = 314: GoTo 90
   If NORIG% = 20 Then RPN% = 107: GoTo 90
   If NORIG% = 21 Then RPN% = 128: GoTo 90
   If NORIG% = 22 Then RPN% = 393: GoTo 90
   If NORIG% = 23 Then RPN% = 980: GoTo 90
   If NORIG% = 24 Then RPN% = 848: GoTo 90
   If NORIG% = 25 Then RPN% = 170: GoTo 90
   If NORIG% = 26 Then RPN% = 310: GoTo 90
   If NORIG% = 27 Then RPN% = 766: GoTo 90
   If NORIG% = 28 Then RPN% = 48: GoTo 90
   If NORIG% = 29 Then RPN% = 842: GoTo 90
   If NORIG% = 30 Then RPN% = 82: GoTo 90
   If NORIG% = 31 Then RPN% = 434: GoTo 90
   If NORIG% = 32 Then RPN% = 300: GoTo 90
   If NORIG% = 33 Then RPN% = 860: GoTo 90
   If NORIG% = 34 Then RPN% = 519: GoTo 90
   '
90 REPNU% = RPN%
   '
End Function

Private Sub MARCOBARRA_DblClick()
   INTERRUMPE% = 1
End Sub

'
Private Sub Mastok_Click()
   Call Command1_Click
End Sub

Private Sub Proveed_Click()
   Call Command19_Click
End Sub

'\\\OT-05-0656-WAR-03/10/05///
Sub CONVBADA() 'CONVERSIÓN DE .DAT A .MDB
   '
   If COMALTER%("La Conversión de Datos\Debe Correrse en el Servidor.\\Continuar\Cancelar") <> 1 Then
      Exit Sub
   End If
   '
   If COMALTER%("Antes de la Conversión de Datos\Conviene en el Sistema DAT\Pedir un Listado de Cuentas\por Cobrar y por Pagar\para Verificar la Composición de Deuda.\\Continuar\Cancelar") <> 1 Then
      Exit Sub
   End If
   '
   MIGRATODO% = 0
   If COMALTER%("Opciones de Trabajo:\\Por Cada Tabla\Barrida General") = 2 Then
      MIGRATODO% = 1
   End If
   '
   Label1.Visible = True
   If MIGRATODO% = 1 Then GoTo 21
   If COMALTER%("Plan de Cuentas Contables\\Migrar\Saltear") = 1 Then Call MIGRAPLANCUEN
   '
21 For IXX% = 0 To 6
      TXXX$ = TRIM$(Str$(IXX%))
      CUEM$ = Control$("", "CUEIVA-" + TXXX$)
      If CUEM$ = "" Or CUEINT%(CUEM$) < 1 Then
         Call MENSERR(24, "ADVERTENCIA: Falta Definir Cuentas Contables\Correspondientes a IVA Crédito Fiscal Tasa " + TTXX$ + ".")
         GoTo 22   ' Call FINAL("")
      End If
   Next IXX%
   '
22 FINI& = ULTREG&("PROVINB")
   If FINI& < 1 Then
      Call MENSERR(24, "ADVERTENCIA: Falta Tabla de Provincias")
      GoTo 23    '  Call FINAL("")
   End If
   '
   For i& = 1 To FINI&
      X$ = REGLEIDO$("PROVINB", i&)
      CUEII% = CVI(Mid$(X$, 27, 2))
      If CUEII% < 1 Or ECOARCH$("CUECON", MKI$(CUEII%)) = "" Then
         Call MENSERR(24, "Faltan Cuentas Contables de Percepcion Ingresos Brutos")
         GoTo 23    '   Call FINAL("")
      End If
   Next i&
   If CUEIPROVIN%("B") < 1 Then
      Call MENSERR(24, "Falta Cuenta Contable I.Brutos Provincia 'B'")
      GoTo 23 ' Call FINAL("")
   End If
   '
   CUEBAL04.Show 1
   Call FRESHALL
   CUEBALRECI% = CUEINT%(CUEBAL04.Label1)
   CUEBALPAGO% = CUEINT%(CUEBAL04.Label2)
   CUEBALAJUS% = CUEINT%(CUEBAL04.Label3)
   CUEBALCART% = CUEINT%(CUEBAL04.Label4)
   If CUEBALRECI% < 1 Or CUEBALPAGO% < 1 Or CUEBALAJUS% < 1 Or CUEBALCART% < 1 Then
      Call FINAL("")
   End If
   '
23 If MIGRATODO% = 1 Then
      Call MIGRACLIEN
      Call MIGRAPROVE
      Call MIGRAMASTER
      Call MIGRAFORMULA
      Call MIGRASECOPER
      Call MIGRAPEDIDOS
      Call MIGRACOMPRAS
      Call MIGRASTOCKS
      Call MIGRABOLREC
      Call MIGRACONSIGNA
      Call MIGRAORFAB
25    Call MIGRAANOTAORFAB
      Call MIGRARESERVA
      Call MIGRAPRESU
      Call MIGRAPRECIOS
      Call MIGRAREPARA
      Call MIGRAALMAMODE
      GoTo 42
   End If
   '
   If COMALTER%("Clientes\\Migrar\Saltear") = 1 Then Call MIGRACLIEN
   If COMALTER%("Proveedores\\Migrar\Saltear") = 1 Then Call MIGRAPROVE
   If COMALTER%("Maestro de Items de Stock\\Migrar\Saltear") = 1 Then Call MIGRAMASTER
   If COMALTER%("Formulas de Producto\\Migrar\Saltear") = 1 Then Call MIGRAFORMULA
   If COMALTER%("Secuencia de Operaciones\\Migrar\Saltear") = 1 Then Call MIGRASECOPER
30 If COMALTER%("Pedidos de Clientes\\Migrar\Saltear") = 1 Then Call MIGRAPEDIDOS
   If COMALTER%("Ordenes de Compras\\Migrar\Saltear") = 1 Then Call MIGRACOMPRAS
32 If COMALTER%("Movimientos de Stocks\\Migrar\Saltear") = 1 Then Call MIGRASTOCKS
   If COMALTER%("Boletas de Recepción\\Migrar\Saltear") = 1 Then Call MIGRABOLREC
   If COMALTER%("Material en Consignacion\\Migrar\Saltear") = 1 Then Call MIGRACONSIGNA
   If COMALTER%("Ordenes de Fabricacion\\Migrar\Saltear") = 1 Then Call MIGRAORFAB
   If COMALTER%("Reservas para O/F\\Migrar\Saltear") = 1 Then Call MIGRARESERVA
   If COMALTER%("Reparaciones\\Migrar\Saltear") = 1 Then Call MIGRAREPARA
   If COMALTER%("Almacen de Modelos\\Migrar\Saltear") = 1 Then Call MIGRAALMAMODE
   '
   If COMALTER%("Presupuestos Emitidos\\Migrar\Saltear") = 1 Then Call MIGRAPRESU
   '
35 If COMALTER%("Lista de Precios de Venta\\Migrar\Saltear") = 1 Then Call MIGRAPRECIOS
   '
40 If COMALTER%("Cuecorri + Cajabanc\\Migrar\Saltear") <> 1 Then GoTo 50
   '
42 'Call BORRACUECORR
   '
   MsgBox "Ventas"
   'Call CONVERFAC
   MsgBox "Cobranza"
   'Call CONVERCOB
   MsgBox "Deuda"
   'Call CONVERDEU
   MsgBox "Aplicacion Cobranza"
45 Call CONVAPLICOB
   '
   MsgBox "Acreedores"
46 Call GRAFACRE
    MsgBox "Cheques"
   Call MIGRADECHE
   '
48 PORCABA% = 0
   If ULTREG&("CAJABANC") > 0 Then
      PORCABA% = 1
   End If
   '
   If PORCABA = 0 Then
         Call MIGRACARCHE
      Else
         Call MARCACARCHE
   End If
   '
50 If MIGRATODO% = 1 Then GoTo 55
   If COMALTER%("Contabilidad General\\Migrar\Saltear") = 1 Then
      TTXX$ = "Para la Conversión los Siguientes Archivos\Correspondientes a los Ejercicios 2001 a 2007\Deben Colocarse en la Carpeta de Datos\'" + REPLACAR$(LUDAT$, "\", "/") + "'."
      TTXX$ = TTXX$ + "\   \La Extensión en cada Caso Debe Reemplazarse por '.nn'\en Correspondencia con el Año Considerado."
      TTXX$ = TTXX$ + "\   \Archivos ENCDIAR0, DIARIO0, DTDIAR0,IDIARIO0,\ENCAJA0, SUCAJA0,ISUCAJA0"
      If COMALTER%(TTXX$ + "\\Continuar\Cancelar") = 1 Then
55       Call APERBASDAT("PLANCUEN")
         On Error Resume Next
         SQLX$ = "DROP TABLE EncabeAsien "
             FLEXCONN.Execute (FILTSQL$(SQLX$))
         SQLX$ = "DROP TABLE DetaAsien "
             FLEXCONN.Execute (FILTSQL$(SQLX$))
         On Error GoTo 0
         Call GENERBASE("CONTAGEN")
         Call MIGRACONTA
         Call MIGRACONTB
         Call CIERRACONTA
         '
         REXTE$ = EXTE$
         For KANO% = 1 To 7
            EXTE$ = TRIM$(Str$(KANO%))
            While Len(EXTE$) < 2: EXTE$ = "0" + EXTE$: Wend
            Call COMUNI("Migrando Ejercicio 20" + EXTE$ + "  Pulse Aceptar ...")
            Call MIGRACONTA
            Call MIGRACONTB
            Call CIERRACONTA
         Next KANO%
      End If
   End If
   '
95 Screen.MousePointer = 1
  Label1.Visible = False
  Call COMUNI("Ha Finalizado con Éxito la Conversión.")
  '
End Sub

Sub MARCACARCHE()
   '
   RFF$ = RECFILT$("CAJABANC", 16, Chr$(1)) + RECFILT$("CAJABANC", 16, Chr$(11))
   JJ& = 0
   FIN& = Len(RFF$)
   '
   Call APERBASDAT("CABAN")
   
   For i& = 1 To FIN& Step 4
     Call TRACE(20, i&, FIN&)
     RECHE& = CVS(Mid$(RFF$, i&, 4))
     X$ = REGLEIDO$("CAJABANC", RECHE&)
     '
     If UCase$(Mid$(X$, 69, 2)) <> "CH" Then  ' no es un cheque
       GoTo 19
     End If
     '
     If CVI(Mid$(X$, 269, 2)) > 0 Then ' esta anulado
       Call REPLA(X$, Chr$(9), 121, 1)
       Call GRAREG("CAJABANC", X$, RECHE&)
       GoTo 19
     End If
     '
     If CVI(Mid$(X$, 275, 2)) > 0 Then ' salio de cartera
       Call REPLA(X$, Chr$(2), 121, 1)
       Call GRAREG("CAJABANC", X$, RECHE&)
       GoTo 19
     End If
     '
     IMCHE# = CVD(Mid$(X$, 153, 8)) - CVD(Mid$(Z$, 161, 8))
     IDEVA$ = Mid$(X$, 71, 40)
     IMCHEX$ = TRIM$(FORMATNUM$(Abs(IMCHE#), "F14.2"))
       If IMCHEX$ = "" Then IMCHEX$ = "0"
     
     SQLX$ = "SELECT * FROM CAJABANC "
     SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI'"
     SQLX$ = SQLX$ + "AND OPCIMOVIM = 'CH' "
     SQLX$ = SQLX$ + "AND DescriMov = '" & IDEVA$ & "' "
     SQLX$ = SQLX$ + "AND ABS(ValAbsComp - " & IMCHEX$ & ") < 0.005 "
     'Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
     Dim CHETEMP1 As ADODB.Recordset
     Set CHETEMP1 = New ADODB.Recordset
     CHETEMP1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     '
     On Error Resume Next
     CHETEMP1.MoveFirst
     If Err Then
          On Error GoTo 0
          MsgBox "No encontrado"
        Else
          On Error GoTo 0
          'CHETEMP1.Edit
          CHETEMP1!Status = 11
          CHETEMP1.Update
     End If
19 Next i&
   '
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI'"
   SQLX$ = SQLX$ + "AND OPCIMOVIM = 'CH' "
   SQLX$ = SQLX$ + "AND STATUS = 1"
   'Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   'Dim CHETEMP1 As ADODB.Recordset
   Set CHETEMP1 = New ADODB.Recordset
   CHETEMP1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   On Error Resume Next
   With CHETEMP1
   .MoveFirst
   If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
          IDEVA$ = !DescriMov
          Import# = !Idebecomp
          ABCX1$ = IDEVA$
          GoSub BUSCALASALIDA
          If ENCONTROLASALIDA% > 0 Then
             GoTo 29
          End If
          '
          FENTREGA% = CVINT%(!FECHCONTA) ' Si no en BADACHE, sale misma fecha que entró
          FEMICHE% = CVINT%(!FECHEMISI)
          DOCUEN$ = "Mig.Datos"
          DESTI$ = "Cuenta Balance Cartera"
          OB1$ = "Conversión Datos DAT-MDB "
          '.Edit
          !Status = 2
          !FEDEPOSENT = CVDAT(FENTREGA%)
          !FECHENTRE = CVDAT(FENTREGA%)
          !Document = DOCUEN$
          !DESTICHEQ = DESTI$
          !REFESALCHE = OB1$
          !IDENSALCHE = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
          !UsuaSalChe = OB3$ ' AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
          .Update
          '
          GoSub AGREGASALIDA
          GoTo 29
          '
29     .MoveNext
       Loop
   End If
   End With
   '
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI'"
   SQLX$ = SQLX$ + "AND OPCIMOVIM = 'CH' "
   SQLX$ = SQLX$ + "AND STATUS = 11"
   'Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   'Dim CHETEMP1 As ADODB.Recordset
   Set CHETEMP1 = New ADODB.Recordset
   CHETEMP1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   On Error Resume Next
   With CHETEMP1
      .MoveFirst
      If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
         '.Edit
         !Status = 1
         .Update
         .MoveNext
         Loop
      End If
   End With
   On Error GoTo 0
   CHETEMP1.Close
   Set CHETEMP1 = Nothing
Exit Sub
   '
BUSCALASALIDA:
    ENCONTROSALIDA% = 0
    With CajaBanc
     .FindFirst "DescriMov = '" & IDEVA$ & "' " ' and VariMovim = '1' "
37   If .NoMatch = False Then
        If Abs(Import# - IHABECOMP) < 0.005 Then
          .Edit
          !Status = 2
          .Update
          ENCONTROLASALIDA% = 1
          GoTo 85
        End If
        .FindNext "DescriMov = '" & IDEVA$ & "' "
        GoTo 37
     End If
   End With
85 Return
   '
   ' **********************************************************************
AGREGASALIDA:
    '
    With CajaBanc
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
      !CodiCom_Cor = "MIGRA"
      !codicom_lar = "MIGRA"
      !RefeCom = DESTI$
      !FECHCONTA = CVDAT(FENTREGA%)
      !FechAcred = CVDAT(FENTREGA%)
      !CUECONTAI = CUECARTE%
      !CenCosto = ""
      !ItemConc = ""
      !NUCLIEN = 0
      !NuProve = 0
      !TipoMovim = "CHENTR"
      !OpciMovim = "CH"
      !VariMovim = "1"
      !DescriMov = IDEVA$
      !FECHEMISI = CVDAT(FEMICHE%)
      !CodiEmpr = 0
      !CodiSucu = 0
      !NumeTalo = 0
      !NumeComp = 0
      !Status = 1
        On Error Resume Next
      !Mone_Emis = 1
      !Tipo_Cam = 1
        On Error GoTo 0
      !ValAbsComp = Abs(Import#)
      !Idebecomp = 0
      !IHABECOMP = Import#
      !Impo_Dola = (-1) * Import#
      !BASEIMPO = Abs(Import#)
      !NROCAI = ""
        On Error Resume Next
      !IncSubIva = ""
        On Error GoTo 0
      !FEVTOCAI = CVDAT(0)
      !ObserGen = ""
      !IdenAlta = Left$(IDACTU$, 8)          ' FECHA - HORA - USUARIO - TERMINAL
      !IdenBaja = String$(6, 0)
      !FeHoReal = Now
      !NuAsiDiar = ""
     .Update
     '
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CodiEmpr = CodiEmp%
     On Error GoTo 0
      !CodiCom_Cor = "MIGRA"
      !codicom_lar = "MIGRA"
      !RefeCom = DESTI$
      !FECHCONTA = CVDAT(FENTREGA%)
      !FechAcred = CVDAT(FENTREGA%)
      !CUECONTAI = CUEBALCART%
      !CenCosto = ""
      !ItemConc = ""
      !NUCLIEN = 0
      !NuProve = 0
      !TipoMovim = "MIGRA"
      !OpciMovim = ""
      !VariMovim = ""
      !DescriMov = IDEVA$
      !FECHEMISI = CVDAT(FEMICHE%)
      !CodiEmpr = 0
      !CodiSucu = 0
      !NumeTalo = 0
      !NumeComp = 0
      !Status = 0
        On Error Resume Next
      !Mone_Emis = 1
      !Tipo_Cam = 1
        On Error GoTo 0
      !ValAbsComp = Import#
      !Idebecomp = Import#
      !IHABECOMP = 0
      !Impo_Dola = Import#
      !BASEIMPO = Abs(Import#)
      !NROCAI = ""
        On Error Resume Next
      !IncSubIva = ""
        On Error GoTo 0
      !FEVTOCAI = CVDAT(0)
      !ObserGen = ""
      !IdenAlta = Left$(IDACTU$, 8)          ' FECHA - HORA - USUARIO - TERMINAL
      !IdenBaja = String$(6, 0)
      !FeHoReal = Now
      !NuAsiDiar = ""
      .Update
   End With
Return

'
'
End Sub

Sub CONVERDEU()
    '
   Picture1.ToolTipText = "Conversión de Saldos Deudores"
   '
65 Call APERBASDAT("CABAN")
    FLEXCONN.Execute (FILTSQL$("DELETE FROM SADEUPEN"))
   '
   FIN& = ULTREG&("SVD202")
   Label1.ToolTipText = "Hasta " & FIN&
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Label1 = U&
      Label1.Refresh
      X$ = REGLEIDO$("SVD202", U&)
      Call DATDEUPEN(X$, TIDOCU$, NFCX$, NFCY$, OPMOVIX$, VAX$, NC%, FFI%, FVENCI%, SIGX%, IORIG#, IPENDE#)

      ' AQUI GRABAR CAJABANC PARA FACTURA
      NOCLOSE% = 1
30    CoComCor$ = NFCX$
      CoComLar$ = NFCY$
'      CodiEmp% = 0
'      CodiSuc% = 0
      NumeClie = Int(GRUDIVCLI(NC))
      NumeProv% = 0
      FechEmi = CVDAT(FFI%)
      MONEMI% = 1
      IBruTot# = IORIG#
      IIvaTot# = 0
      FeVenc = CVDAT(FVENCI%)
      ImVenc = IPENDE#
      SIGNO% = 1
      '
      NuorVen% = 1
      Call GRADEUPEN
      '
   Next U&
   '
End Sub

Sub CONVAPLICOB()
   '
   Call APERBASDAT("CABAN")
   On Error Resume Next
   FLEXCONN.Execute (FILTSQL$("DELETE FROM APLICOBRA"))
   On Error GoTo 0
      '
   FIN& = ULTREG&("APLICOB")
   Label1.ToolTipText = "Hasta " & FIN&
   Label1.Visible = True
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     Label1 = U&
     Label1.Refresh
     XX$ = REGLEIDO$("APLICOB", U&)
     NumeClie = CVI(Mid$(XX$, 3, 2))
     '
     CoComLar$ = Mid$(XX$, 5, 18)
     NUMEFA1$ = CoComLar$
     NUMEFA2$ = Left$(CoComLar$, 3) + "X" + Mid$(CoComLar$, 5)
     NUMEFA3$ = Left$(NUMEFA2$, 5) + "0000" + Mid$(NUMEFA2$, 10)
     Condi$ = Condi$ + " NuClien = " & NumeClie
     Condi$ = Condi$ + " AND (CodiCom_Lar = '" & NUMEFA1$ & "' OR CodiCom_Lar = '" & NUMEFA2$ & "' OR CodiCom_Lar = '" & NUMEFA3$ & "') "
     CODICOM$ = SAFETEXT$(VALOBADA("CAJABANC", "CODICOM_LAR", Condi$, "NOMESS"))
     If CODICOM$ <> "" Then
        CoComLar = CODICOM$
        GoTo 15
     End If
     ' SEGUNDA BUSQUEDA SI NO LO ENCONTRO
     NUMEFA$ = Left$(CoComLar$, 3) + "_______" + Mid$(CoComLar$, 11)
     NUMEFB$ = Left$(CoComLar$, 3) + "_____" + Mid$(CoComLar$, 11, 8) + "__"
     NUMEFC$ = Left$(CoComLar$, 1) + "_______" + Mid$(CoComLar$, 9, 9) + "_"
     Condi$ = Condi$ + " NuClien = " & NumeClie
     Condi$ = Condi$ + " AND (CodiCom_Lar LIKE '" & NUMEFA$ & "' OR CodiCom_Lar LIKE '" & NUMEFB$ & "' OR CodiCom_Lar LIKE '" & NUMEFC$ & "') "
     CODICOM$ = SAFETEXT$(VALOBADA("CAJABANC", "CODICOM_LAR", Condi$, "NOMESS"))
     If CODICOM$ <> "" Then
        CoComLar = CODICOM$
        GoTo 15
     End If
     'MsgBox CoComLar$
     GoTo 39          ' SI NO LO ENCUENTRA NO LO MIGRA
     
15   NROX$ = CoComLar$
     CoComCor$ = Left$(NROX$, 3) + Mid$(NROX$, 9, 1) + "-" + Mid$(NROX$, 14)
     ImpApli# = CVD(Mid$(XX$, 23, 8))
     RefApli$ = Mid$(XX$, 31, 24)
     DocOrLar$ = Mid$(XX$, 55, 18)
     NUMEFA1$ = DocOrLar$
     NUMEFA2$ = Left$(DocOrLar$, 3) + "0001" + Mid$(DocOrLar$, 8)
     Condi$ = Condi$ + " NuClien = " & NumeClie
     Condi$ = Condi$ + " AND (CodiCom_Lar = '" & NUMEFA1$ & "' OR CodiCom_Lar = '" & NUMEFA2$ & "') "
     CODICOM$ = SAFETEXT$(VALOBADA("CAJABANC", "CODICOM_LAR", Condi$, "NOMESS"))
     If CODICOM$ <> "" Then
        GoTo 25
     End If
     
     'SEGUNDA VUELTA SI NO LO ENCONTRO
     NUMEFA$ = Left$(DocOrLar$, 3) + "_______" + Mid$(DocOrLar$, 11)
     NUMEFB$ = Left$(DocOrLar$, 3) + "_____" + Mid$(DocOrLar$, 9, 8) + "__"
     NUMEFC$ = Left$(DocOrLar$, 1) + "_______" + Mid$(DocOrLar$, 9, 9) + "_"
     Condi$ = Condi$ + " NuClien = " & NumeClie
     Condi$ = Condi$ + " AND (CodiCom_Lar LIKE '" & NUMEFA$ & "' OR CodiCom_Lar LIKE '" & NUMEFB$ & "' OR CodiCom_Lar LIKE '" & NUMEFC$ & "') "
     CODICOM$ = SAFETEXT$(VALOBADA("CAJABANC", "CODICOM_LAR", Condi$, "NOMESS"))
     If CODICOM$ <> "" Then
        GoTo 25
     End If
     'MsgBox DocOrLar$
     GoTo 39          ' SI NO LO ENCUENTRA NO LO MIGRA
     
25   NROX$ = DocOrLar$
     DocOrCor$ = Left$(NROX$, 3) + Mid$(NROX$, 9, 1) + "-" + Mid$(NROX$, 14)
     TotOrig# = CVD(Mid$(XX$, 73, 8))
     IvaOrig# = CVD(Mid$(XX$, 81, 8))
     Call GRAPLICOB
39 Next U&
   
'
'   Call APERBASDAT("CABAN")
'   On Error Resume Next
'   'FLEXCONN.Execute (FILTSQL$("DROP TABLE APLICOBRA"))
'   FLEXCONN.Execute (FILTSQL$("DELETE FROM APLICOBRA"))
'   On Error GoTo 0
'      '
'   FIN& = ULTREG&("APLICOB")
'   Label1.ToolTipText = "Hasta " & FIN&
'   Label1.Visible = True
'   For U& = 1 To FIN&
'     Call TRACE(20, U&, FIN&)
'     Label1 = U&
'     Label1.Refresh
'     XX$ = REGLEIDO$("APLICOB", U&)
'     NumeClie = CVI(Mid$(XX$, 3, 2))
'     '
'     SQLX$ = "SELECT CodiCom_Lar FROM CAJABANC "
'     SQLX$ = SQLX$ + "where NuClien = " & NumeClie
'     'Set MOFASEL = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
'     Dim MOFASEL As ADODB.Recordset
'     '
'     CoComLar$ = Mid$(XX$, 5, 18)
'     NUMEFA1$ = CoComLar$
'     NUMEFA2$ = Left$(CoComLar$, 3) + "X" + Mid$(CoComLar$, 5)
'     NUMEFA3$ = Left$(NUMEFA2$, 5) + "0000" + Mid$(NUMEFA2$, 10)
'     'NUMEFA$ = Left$(CoComLar$, 3) + "_______" + Mid$(CoComLar$, 11)
'     'NUMEFB$ = Left$(CoComLar$, 3) + "_____" + Mid$(CoComLar$, 9, 8) + "__"
'     SQLX$ = "SELECT CodiCom_Lar FROM CAJABANC "
'     SQLX$ = SQLX$ + " where NuClien = " & NumeClie
'     'SQLX$ = SQLX$ + " AND (CodiCom_Lar LIKE '" & NUMEFA$ & "' OR CodiCom_Lar LIKE '" & NUMEFB$ & "') "
'     SQLX$ = SQLX$ + " AND (CodiCom_Lar = '" & NUMEFA1$ & "' OR CodiCom_Lar = '" & NUMEFA2$ & "' OR CodiCom_Lar = '" & NUMEFA3$ & "') "
'     Set MOFASEL = New ADODB.Recordset
'     MOFASEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'     On Error Resume Next
'     MOFASEL.MoveFirst
'     If Err < 1 Then
'        On Error GoTo 0
'        CoComLar = MOFASEL!CodiCom_Lar
'        GoTo 15
'     End If
'     ' SEGUNDA BUSQUEDA SI NO LO ENCONTRO
'     NUMEFA$ = Left$(CoComLar$, 3) + "_______" + Mid$(CoComLar$, 11)
'     NUMEFB$ = Left$(CoComLar$, 3) + "_____" + Mid$(CoComLar$, 11, 8) + "__"
'     NUMEFC$ = Left$(CoComLar$, 1) + "_______" + Mid$(CoComLar$, 9, 9) + "_"
'     SQLX$ = "SELECT CodiCom_Lar FROM CAJABANC "
'     SQLX$ = SQLX$ + " where NuClien = " & NumeClie
'     SQLX$ = SQLX$ + " AND (CodiCom_Lar LIKE '" & NUMEFA$ & "' OR CodiCom_Lar LIKE '" & NUMEFB$ & "' OR CodiCom_Lar LIKE '" & NUMEFC$ & "') "
'     'SQLX$ = SQLX$ + " AND NUMECOMP = 412 "
'     Set MOFASEL = New ADODB.Recordset
'     MOFASEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'     On Error Resume Next
'     MOFASEL.MoveFirst
'     If Err < 1 Then
'        On Error GoTo 0
'        CoComLar = MOFASEL!CodiCom_Lar
'        GoTo 15
'     End If
'     On Error GoTo 0
'     GoTo 39          ' SI NO LO ENCUENTRA NO LO MIGRA
'15   NROX$ = CoComLar$
'     CoComCor$ = Left$(NROX$, 3) + Mid$(NROX$, 9, 1) + "-" + Mid$(NROX$, 14)
'     ImpApli# = CVD(Mid$(XX$, 23, 8))
'     RefApli$ = Mid$(XX$, 31, 24)
'     DocOrLar$ = Mid$(XX$, 55, 18)
'     NUMEFA1$ = DocOrLar$
'     NUMEFA2$ = Left$(DocOrLar$, 3) + "0001" + Mid$(DocOrLar$, 8)
'     SQLX$ = "SELECT CodiCom_Lar FROM CAJABANC "
'     SQLX$ = SQLX$ + " where NuClien = " & NumeClie
'     SQLX$ = SQLX$ + " AND (CodiCom_Lar = '" & NUMEFA1$ & "' OR CodiCom_Lar = '" & NUMEFA2$ & "') "
'     Set MOFASEL = New ADODB.Recordset
'     MOFASEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'     On Error Resume Next
'     MOFASEL.MoveFirst
'     If Err < 1 Then
'        On Error GoTo 0
'        'CoComLar = MOFASEL!CodiCom_Lar
'        GoTo 25
'     End If
'     'SEGUNDA VUELTA SI NO LO ENCONTRO
'     NUMEFA$ = Left$(DocOrLar$, 3) + "_______" + Mid$(DocOrLar$, 11)
'     NUMEFB$ = Left$(DocOrLar$, 3) + "_____" + Mid$(DocOrLar$, 9, 8) + "__"
'     NUMEFC$ = Left$(DocOrLar$, 1) + "_______" + Mid$(DocOrLar$, 9, 9) + "_"
'     SQLX$ = "SELECT CodiCom_Lar FROM CAJABANC "
'     SQLX$ = SQLX$ + " where NuClien = " & NumeClie
'     SQLX$ = SQLX$ + " AND (CodiCom_Lar LIKE '" & NUMEFA$ & "' OR CodiCom_Lar LIKE '" & NUMEFB$ & "' OR CodiCom_Lar LIKE '" & NUMEFC$ & "') "
'     'SQLX$ = SQLX$ + " AND NUMECOMP = 140419 "
'     Set MOFASEL = New ADODB.Recordset
'     MOFASEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'     On Error Resume Next
'     MOFASEL.MoveFirst
'     If Err < 1 Then
'        On Error GoTo 0
'        'CoComLar = MOFASEL!CodiCom_Lar
'        GoTo 25
'     End If
'     On Error GoTo 0
'     GoTo 39          ' SI NO LO ENCUENTRA NO LO MIGRA
''MsgBox "No Encontré " + DocOrLar$
'25   NROX$ = DocOrLar$
'     DocOrCor$ = Left$(NROX$, 3) + Mid$(NROX$, 9, 1) + "-" + Mid$(NROX$, 14)
'     TotOrig# = CVD(Mid$(XX$, 73, 8))
'     IvaOrig# = CVD(Mid$(XX$, 81, 8))
'     Call GRAPLICOB
'39 Next U&
   '
End Sub
'
Sub CONVERCOB()
   '
   Picture1.ToolTipText = "Conversion de Cobranzas"
   Call APERBASDAT("CABAN")
   NOCLOSE% = 1
   '
   CUECARTE% = CUECOINT%(Control$("", "CUECHECA"))
   FIN& = ULTREG&("COBRAN")
   DIFECOB# = 0
   '
   PORCABA% = 0
   If ULTREG&("CAJABANC") > 0 Then
      PORCABA% = 1
   End If
   '
   Label1.ToolTipText = "Hasta " & FIN&
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Label1 = U&
      Label1.Refresh
      X$ = REGLEIDO$("COBRAN", U&)
      NC% = CVI(Left$(X$, 2))
      NROO& = CVS(Mid$(X$, 3, 4))
      FC% = CVI(Mid$(X$, 7, 2))
      TOTAVAL# = CVD(Mid$(X$, 15, 8))
      If TOTAVAL# < 0.005 Then GoTo 89 ' ESTÁ ANULADO
      COBRACT% = CVI(Mid$(X$, 23, 2))
      '
      RECICOR$ = TRIM$(Str$(COBRACT%))
      While Len(RECICOR$) < 2: RECICOR$ = "0" + RECICOR$: Wend
      RECICOR$ = TRIM$(Str$(NROO&)) + "-" + RECICOR$
      While Len(RECICOR$) < 9: RECICOR$ = "0" + RECICOR$: Wend
      RECICOR$ = "RB-" + RECICOR$
      NRECIB$ = RECICOR$
      '
      RECILAR$ = TRIM$(Str$(NROO&))
      While Len(RECILAR$) < 8: RECILAR$ = "0" + RECILAR$: Wend
      RECILAR$ = TRIM$(Str$(COBRACT%)) + "-" + RECILAR$
      While Len(RECILAR$) < 13: RECILAR = "0" + RECILAR$: Wend
      RECILAR$ = "RB-X-" + RECILAR$
      '
      CoComCor$ = NRECIB$
      CoComLar$ = RECILAR$
      Referen$ = rasocli$(NC%)
      SUVALOR# = 0
      FeContab = CVDAT(FC%)
      FeAcredi = CVDAT(FC%)
      '
      '
      If PORCABA% = 1 Then
         RFF$ = RECFILT$("CAJABANC", 1, NRECIB$)
         For KKU& = 1 To Len(RFF$) Step 4
            REG& = CVS(Mid$(RFF$, KKU&, 4))
            Z$ = REGLEIDO$("CAJABANC", REG&)
            '
            IMCHE# = CVD(Mid$(Z$, 153, 8))
            If IMCHE# > 0 Then
              FC1% = CVI(Mid$(Z$, 61, 2))
              If FC1% = FC% Then
                 SUVALOR# = SUVALOR# + IMCHE#
'                 CodiEmp% = Asc(Mid$(Z$, 113, 1))
'                 CodiSuc% = Asc(Mid$(Z$, 114, 1))
                 NumeClie = NC%
                 NumeProv% = 0
                 '
                 FECHE% = CVI(Mid$(Z$, 111, 2))
                 FEVEN% = CVI(Mid$(Z$, 63, 2))
                 CUITLIB$ = ""                 ' NO HAY DATOS PARA MIGRAR
                 COBRACT% = CVI(Mid$(Z$, 115, 2))
                 SUCUI% = 0                    ' NO HAY DATOS PARA MIGRAR
                 '
                   FEX = FEVEN%: FEVECHE$ = FECHATEX$(FEX)
                 FeContab = CVDAT(FC%)
                   CUECO% = CVI(Mid$(Z$, 65, 2))
                 CuentaCo% = CUECO%
                 TiMovi$ = "CHRECI"
                 OpMovi$ = "CH"
                 VaMovi$ = ""
                   IDEVA$ = Mid$(Z$, 71, 40)
                 Descrip$ = IDEVA$
                   If FECHE% < 1 Then FECHE% = FC%
                 FechEmi = CVDAT(FECHE%)
                 FeAcredi = CVDAT(FEVEN%)
                 NumeTal = COBRACT%
                 NumeCom = CVD(Mid$(Z$, 117, 4))
                    STATUI% = Asc(Mid$(Z$, 121, 1))
                    If CUECO% = CUECARTE% Then STATUI% = 1
                 StatInit = STATUI%
                 ImpoDeb# = IMCHE#
                 Observa$ = ""
                   If CUITLIB$ <> "" Then
                     Observa$ = "CUIT Libr: " + CUITLIB$
                   End If
                 Call GRACABAN
                 '
              End If
            End If
         Next KKU&
         GoTo 850
      End If
      '
      '

840   ' ***** aqui grabar cheques en CAJABANC.DAT
      '
      RFF$ = RECFILT$("CHECAR", 11, MKS$(NROO&))
      SUVALOR# = 0
      For KKU& = 1 To Len(RFF$) Step 4
        REG& = CVS(Mid$(RFF$, KKU&, 4))
        Z$ = REGLEIDO$("CHECAR", REG&)
        '
        FC1% = CVI(Mid$(Z$, 45, 2))
        '
        If FC1% = FC% Then
           Tipo% = CVI(Left$(Z$, 2))
           BANCO$ = TRIM$(Mid$(Z$, 17, 16))
           If BANCO$ = "" Then
             BANCO$ = TRIM$(ECOARCH$("MECOBRA", MKI$(Tipo%)))
           End If
           If Left$(UCase$(BANCO$), 3) = "CH." Then
             If Len(BANCO$) > 3 Then
               BANCO$ = Mid$(BANCO$, 4)
             End If
           End If
           NUCHE$ = TRIM$(Mid$(Z$, 5, 12))
           IMCHE# = CVD(Mid$(Z$, 33, 8))
             SUVALOR# = SUVALOR# + IMCHE#
           FECHE% = CVI(Mid$(Z$, 43, 2))
           FEVEN% = CVI(Mid$(Z$, 41, 2))
           CUITLIB$ = ""                 ' NO HAY DATOS PARA MIGRAR
           COBRACT% = 0
           SUCUI% = 0                    ' NO HAY DATOS PARA MIGRAR
           '
           If Abs(IMCHE#) > 0 Then
             '
               FEX = FEVEN%: FEVECHE$ = FECHATEX$(FEX)
             FeContab = CVDAT(FC%)
               CUECO% = CUEMECOB%(Tipo%)
             CuentaCo% = CUECO%
             NumeClie = NC%
             NumeProv% = 0
             TiMovi$ = "CHRECI"
             OpMovi$ = "CH"
             VaMovi$ = ""
               IDEVA$ = AJUSTI$(BANCO$, 18)
               IDEVA$ = IDEVA$ + AJUSTI$(NUCHE$, 14)
               IDEVA$ = IDEVA$ + FEVECHE$
             Descrip$ = IDEVA$
               If FECHE% < 1 Then FECHE% = FC%
             FechEmi = CVDAT(FECHE%)
               If FEVEN% < FECHE% Then FEVEN% = FECHE%
               If FEVEN% < FC% Then FEVEN% = FC%
             FeAcredi = CVDAT(FEVEN%)
             NumeTal = COBRACT%
'             CodiEmp% = 0
'             CodiSuc% = SUCUI%
             NumeCom = NROO&
                STATUI% = 0
                If CUECO% = CUECARTE% Then STATUI% = 1
             StatInit = STATUI%
             ImpoDeb# = IMCHE#
             Observa$ = ""
               If CUITLIB$ <> "" Then
                 Observa$ = "CUIT Libr: " + CUITLIB$
               End If
             Call GRACABAN
             '
           End If
        End If
      Next KKU&
      '
850   If Abs(TOTAVAL#) >= 0.005 Then
         FeContab = CVDAT(FC%)
           CUECO% = CUECOCLI%(NC%)
         CuentaCo% = CUECO%
         NumeClie = NC%
         NumeProv% = 0
         TiMovi$ = "RCOB"
         OpMovi$ = "CCORR"
         VaMovi$ = ""
         Descrip$ = "Cobranza Cta.Corriente " + NRECIB$
         FechEmi = CVDAT(FC%)
         FeAcredi = CVDAT(FC%)
         NumeTal = COBRACT%
'         CodiEmp% = 0
'         CodiSuc% = SUCUI%
         NumeCom = NROO&
         StatInit = 0
         ImpoDeb# = (-1) * TOTAVAL#   ' SE FUERZA QUE CUADRE
         Call GRACABAN
      End If
      '
      If Abs(SUVALOR# - TOTAVAL#) > 0.005 Then
        DIFECOB# = ROUND(TOTAVAL# - SUVALOR#)
        Referen$ = "Dif.Importación " + NRECIB$
        CuentaCo% = CUEBALRECI%
        TiMovi$ = "AJUSTE"
        OpMovi$ = ""
        VaMovi$ = ""
        Descrip$ = "Dif.Importación " + NRECIB$
        ImpoDeb# = DIFECOB#
        Observa$ = ""
        Call GRACABAN
      End If
      '
89 Next U&
   '
End Sub

'\\\OT-05-0656-WAR-FIN///
Sub CONVERFAC()
   '
   Picture1.ToolTipText = "Conversión de Facturas"
   FIN& = ULTREG&("FACTUR")
   Label1.ToolTipText = "Hasta " & FIN&
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Label1 = U&
      Label1.Refresh
      X$ = REGLEIDO$("FACTUR", U&)
      Call DATFACTUR(X$, TIDOCU$, NFCX$, NFCY$, OPMOVIX$, VAX$, NC%, FFI%, SIGX%, TOTAGEN#, NEFA7#, IVAD1#, IVAD2#, RNI#, IPER)
      '
      ' AQUI GRABAR CAJABANC PARA FACTURA
      Call APERBASDAT("CABAN")
      NOCLOSE% = 1
30    CoComCor$ = NFCX$
      CoComLar$ = NFCY$
'      CodiEmp% = 0
'      CodiSuc% = 0
      NumeClie = Int(GRUDIVCLI(NC))
      NumeProv% = 0
      FechEmi = CVDAT(FFI%)
      Referen$ = rasocli$(NC%)
      FeContab = CVDAT(FFI%)
      FeAcredi = CVDAT(FFI%)
      CuentaCo% = CUECOCLI%(NC%)
      TiMovi$ = "FVEN"
      OpMovi$ = OPMOVIX$
      VaMovi$ = VAX$
      Descrip$ = TIDOCU$ + " " + VAX$ + " " + Mid$(NFCY$, 4, 13)
      NumeTal = XVALO(Mid$(NFCY$, 4, 4))
      NumeCom = XVALO(Mid$(NFCY$, 9, 8))
      StatInit = 0
      ImpoDeb = TOTAGEN# * SIGX%
      Observa$ = rasocli$(NC%)
      Call GRACABAN
      '
      'graba venta neta - Verificar que QQQ sea la IC de combinacion Cliente-Tipo de Comprobante
      TOTANETX# = NEFA7#
      If Abs(TOTANETX#) >= 0.005 Then
        PPVI% = PIVACLI%(NC%)
        If PPVI% < 0 Or PPVI% > 7 Then PPVI% = 1
        T16$ = Control$("SVD451S", "CUEVE-" + TRIM$(Str$(PPVI%)))
        QQ1% = CUECOINT%(T16$)
        'If QQ1% < 1 Then QQ1% = QQQ%
        CuentaCo% = QQ1%   ' TOMA LA QUE APARECE EN PANTALLA
        TiMovi$ = "VENTA"
        OpMovi$ = OPMOVIX$
        VaMovi$ = VAX
        Descrip$ = "Mercaderia Vendida " + NFCX$
        StatInit = 0
        ImpoDeb = TOTANETX# * SIGX% * (-1)
        Call GRACABAN
      End If
      '
      'graba iva inscripto
      TOTANETX# = IVAD1#
      If Abs(TOTANETX#) >= 0.005 Then
          If CUIVA0% < 1 Then
             CUEIVA$ = Control$("SVD451S", "CUEIVA")
             CUIVA0% = CUECOINT%(CUEIVA$)
          End If
          CuentaCo% = CUIVA0%
          TiMovi$ = "IVA"
          OpMovi$ = "DeFi"
          VaMovi$ = "21.0 %"
          Descrip$ = "Iva Debito Fiscal " + NFCX$
          StatInit = 0
          ImpoDeb = TOTANETX# * SIGX% * (-1)
          BasImpo# = ROUND#(TOTANETX# / 0.21, 2)
          If Abs(BasImpo# - NEFA7#) < 0.1 Then
             BasImpo# = NEFA7#
          End If
          Call GRACABAN
      End If
      BasImpo# = 0
      '
      'graba iva inscripto
      TOTANETX# = IVAD2#
      If Abs(TOTANETX#) >= 0.005 Then
          If CUIVA1% < 1 Then
             CUEIVA$ = Control$("SVD451S", "CUEIVA")
             CUIVA1% = CUECOINT%(CUEIVA$)
          End If
          CuentaCo% = CUIVA1%
          TiMovi$ = "IVA"
          OpMovi$ = "DeFi"
          VaMovi$ = "10.5 %"
          Descrip$ = "Iva Debito Fiscal " + NFCX$
          StatInit = 0
          ImpoDeb = TOTANETX# * SIGX% * (-1)
          BasImpo# = ROUND#(TOTANETX# / 0.105, 2)
          Call GRACABAN
      End If
      BasImpo# = 0
      '
      'graba iva NO inscripto
      TOTANETX# = RNI#
      If Abs(TOTANETX#) >= 0.005 Then
        If CUADU% < 1 Then
           CUEIVA$ = Control$("SVD451S", "CUEADU")
           CUADU% = CUECOINT%(CUEIVA$)
        End If
        CuentaCo% = CUADU%
        TiMovi$ = "IVA"
        OpMovi$ = "DeFi"
        VaMovi$ = "RNI"
        Descrip$ = "IVA Resp.No Inscripto " + NFCX$
        StatInit = 0
        ImpoDeb = TOTANETX# * SIGX% * (-1)
        BasImpo# = ROUND#(NEFA7#, 2)
        Call GRACABAN
      End If
      BasImpo# = 0
      '
      'graba percepcion I.Brutos
      TOTANETX# = IPER
      If Abs(TOTANETX#) >= 0.005 Then
        If CUIVA% < 1 Then
           CUEIVA$ = Control$("PERIBRU", "CUECONTA")
           CUIVA% = CUECOINT%(CUEIVA$)
        End If
        CuentaCo% = CUIVA%
        TiMovi$ = "IBRU"
        OpMovi$ = "IB-B"
        VaMovi$ = "PERC"
        Descrip$ = "Perc.IB.Bs.As. " + NFCX$
        StatInit = 0
        ImpoDeb = TOTANETX# * SIGX% * (-1)
        BasImpo# = ROUND#(NEFA7#, 2)
        Call GRACABAN
      End If
      BasImpo# = 0
      '
   Next U&
   '
End Sub

Sub MIGRADECHE()
   '
   Picture1.ToolTipText = "Conversion de obranzas"
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI'"
   SQLX$ = SQLX$ + "AND Status = 1 "
   SQLX$ = SQLX$ + "AND OPCIMOVIM = 'CH' "
   'Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   Dim CHETEMP1 As ADODB.Recordset
   '
   NOCLOSE% = 0
   Call APERBASDAT("CABAN")
   NOCLOSE% = 1
   '
   FIN& = ULTREG&("CAJABANC")
   For REG& = 1 To FIN&
      Call TRACE(20, REG&, FIN&)
      Z$ = REGLEIDO$("CAJABANC", REG&)
      '
      CoComCor$ = Left$(Z$, 12): NRECIB$ = Left$(Z$, 12)
      If Left$(CoComCor$, 2) <> "DB" Then GoTo 29
      '
      CoComLar$ = Left$(Z$, 12)
'      CodiEmp% = Asc(Mid$(Z$, 113, 1))
'      CodiSuc% = Asc(Mid$(Z$, 114, 1))
      NumeClie = CVI(Mid$(Z$, 67, 2))
      NumeProv% = 0
      FechEmi = CVDAT(CVI(Mid$(Z$, 111, 2)))
      Referen$ = Mid$(Z$, 13, 40): DESTIN$ = TRIM$(Mid$(Z$, 13, 48))

      IMCHE# = CVD(Mid$(Z$, 153, 8)) - CVD(Mid$(Z$, 161, 8))
      FC% = CVI(Mid$(Z$, 61, 2))
      FECHE% = CVI(Mid$(Z$, 111, 2))
      FEVEN% = CVI(Mid$(Z$, 63, 2))
      CUITLIB$ = ""                 ' NO HAY DATOS PARA MIGRAR
      COBRACT% = CVI(Mid$(Z$, 115, 2))
      SUCUI% = Asc(Mid$(Z$, 114, 1))
      '
      If Abs(IMCHE#) > 0 Then
      '
             FEX = FEVEN%: FEVECHE$ = FECHATEX$(FEX)
          FeContab = CVDAT(FC%)
             CUECO% = CVI(Mid$(Z$, 65, 2))
          CuentaCo% = CUECO%
             NumeClie = CVI(Mid$(Z$, 67, 2))
             NumeProv% = 0
             TiMovi$ = "DEPOBAN"
             If IMCHE# < 0 Then TiMovi$ = "CHENTR"
             OpMovi$ = Mid$(Z$, 69, 2)
             VaMovi$ = ""
                IDEVA$ = Mid$(Z$, 71, 40)
             Descrip$ = IDEVA$
             If FECHE% < 1 Then FECHE% = FC%
             FechEmi = CVDAT(FECHE%)
             FeAcredi = CVDAT(FEVEN%)
             NumeTal = COBRACT%
'             CodiEmp% = 0
'             CodiSuc% = SUCUI%
             NumeCom = CVD(Mid$(Z$, 117, 4))
                 STATUI% = Asc(Mid$(Z$, 121, 1))
             StatInit = STATUI%
             ImpoDeb# = IMCHE#
             Observa$ = ""
               If CUITLIB$ <> "" Then
                 Observa$ = "CUIT Libr: " + CUITLIB$
               End If
             Call GRACABAN
             '
             If IMCHE# < 0 Then
                 'BUSCAR SI ESTÁ INGRESADO
                 ' CASO AFIRMATIVO:
                 ' 1) MARCAR LA SALIDA DE CARTERA
                 ' 2) PONER CUENTA = CUENTA DE CARTERA DE VALORES
                 VUELTA% = 0: ENCUEN% = 0
                 IMCHEX$ = TRIM$(FORMATNUM$(Abs(IMCHE#), "F14.2"))
                    If IMCHEX$ = "" Then IMCHEX$ = "0"
                 SQLX$ = "SELECT * FROM CAJABANC "
                 SQLX$ = SQLX$ + " WHERE TipoMovim = 'CHRECI'"
                 SQLX$ = SQLX$ + " AND Status = 1 "
                 SQLX$ = SQLX$ + " AND OPCIMOVIM = 'CH' "
                 SQLX$ = SQLX$ + " AND DescriMov = '" & IDEVA$ & "' "
                 SQLX$ = SQLX$ + " AND ABS(ValAbsComp - " & IMCHEX$ & ") < 0.005 "
                 Set CHETEMP1 = New ADODB.Recordset
                 CHETEMP1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                 ' SQLX$ = SQLX$ + "AND Status = 1"
                 '
                 With CHETEMP1
                    On Error Resume Next
                    .MoveFirst
                    If Err < 1 Then
                       On Error GoTo 0
                       '.Edit
                       !FEDEPOSENT = CVDAT(FC%)
                       !Document = NRECIB$
                       !DESTICHEQ = Left$(DESTIN$, 48)
                         CUECONTABINT% = !CUECONTAI  ' PARA QUE LA SALIDA SEA DE LA MISMA CUENTA QUE ENTRO
                       !REFESALCHE = "Deposito Bancario " + NRECIB$
                       !IDENSALCHE = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
                       !UsuaSalChe = AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
                       !Status = 2                     ' Marca Salida de Cartera
                       !IdenBaja = IDACTU$             ' Marca Modificacion de registro - FECHA - HORA - USUARIO - TERMINAL
                       .Update
                       ENCUEN% = 1
                    End If
                    On Error GoTo 0
                 End With
                 '
             End If
      End If
      '
29  Next REG&
'      '
End Sub


Sub DATFACTUR(X$, TIDOCU$, NFCX$, NFCY$, OPMOVIX$, VAMOVIX$, NC%, FFI%, SIGX%, TOTAGEN#, NETO#, IVAD1#, IVAD2#, RNI#, IPER)
'**********
    '
    Static CTXX%, MODOSABO%, CUEPER$, CUEPIBRU%
    '
    If CTXX% = 0 Then
       EPAC$ = TRIM$(UCase$(EMPREAC$))
       MODOSABO% = 0: CUEPIBRU% = 0
       If InStr(EPAC$, "SABO") > 0 Or InStr(EPAC$, "DOSIVA") > 0 Then
         MODOSABO% = 1
         CUEPER$ = Control$("PERIBRU", "CUECONTA")
         CUEPIBRU% = CUECOINT%(CUEPER$)
       End If
    End If
    '
    NC% = CVI(Left$(X$, 2))
    FFI% = CVI(Mid$(X$, 7, 2))
       If FFI% = 0 Then FFI% = FFANT%
    FFANT% = FFI%
    '
    NRO& = CVS(Mid$(X$, 3, 4))
    NUSUCUVE% = CVI(Mid$(X$, 47, 2))
    NSUCX$ = TRIM$(Str$(NUSUCUVE%)): While Len(NSUCX$) < 4: NSUCX$ = "0" + NSUCX$: Wend
    '
    Tipo% = Asc(Mid$(X$, 13, 1)): If TI% < 0 Or TI% > 7 Then TI% = 2 ' default factura
    TICOMX$ = "FC": OPMOVIX$ = "FAC": TIDOCU$ = "Factura"
    If Tipo% = 3 Then TICOMX$ = "ND": OPMOVIX$ = "DEB": TIDOCU$ = "N.Debito"
    If Tipo% = 6 Then TICOMX$ = "NC": OPMOVIX$ = "CRE": TIDOCU$ = "N.Credito"
    If Tipo% = 0 Or Tipo% = 7 Then TICOMX$ = "AJ": OPMIVIX$ = "AJUD": TIDOCU$ = "Ajuste"
    '
    VA% = Asc(Mid$(X$, 14, 1)): If VA% < 1 Or VA% > 3 Then VA% = 1 ' default factura
    VAX$ = "A"
    If VA% = 2 Then VAX$ = "B"
    If VA% = 3 Then VAX$ = "E"
    '
    NFCZ$ = TRIM$(Str$(NRO&)): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
    NFCX$ = TICOMX$ + "-" + Mid$(NFCZ$, 3) + "-" + VAX$
    NFCY$ = TICOMX$ + "-" + NSUCX$ + "-" + NFCZ$ + "-" + VAX$
    '
    OPMOVIX$ = "FC"
    If Tipo% = 3 Then OPMOVIX$ = "ND"
    If Tipo% = 6 Then OPMOVIX$ = "NC"
    VAMOVIX$ = VAX$
    '
    NETO1# = (Int(100 * (CVD(Mid$(X$, 15, 8))) + 0.5)) / 100
    IVAD1# = (Int(100 * CVD(Mid$(X$, 23, 8)) + 0.5)) / 100
    RNI# = (Int(100 * CVD(Mid$(X$, 31, 8)) + 0.5)) / 100
    NETO2# = (Int(100 * CVD(Mid$(X$, 49, 8)) + 0.5)) / 100
    IVAD2# = (Int(100 * CVD(Mid$(X$, 57, 8)) + 0.5)) / 100
    IPER = 0
    '
    If MODOSABO% = 1 Then
       IPER = IVAD2#
       IVAD2# = 0
    End If
    '
    If Abs(NETO2#) > 0.1 Then
      If Abs(100 * IVAD2# / NETO2# - TAIVA(0)) < 0.2 Then
        NETO1# = NETO1# + NETO2#
        Call REPLA(X$, MKD$(NETO1#), 15, 8)
        NETO2# = 0
        Call REPLA(X$, MKD$(NETO2#), 49, 8)
        IVAD1# = IVAD1# + IVAD2#
        Call REPLA(X$, MKD$(IVAD1#), 23, 8)
        IVAD2# = 0
        Call REPLA(X$, MKD$(IVAD2#), 57, 8)
      End If
    End If
    '
    NETO# = NETO1# + NETO2#
    IVAD# = IVAD1# + IVAD2#
    TOTAGEN# = NETO# + IVAD# + RNI# + CDbl(IPER) + IRET#
    '
    SIGX% = 1
    If Tipo% > 3 Or TOTAGEN# < 0 Then
        NETO1# = Abs(NETO1#)
        NETO2# = Abs(NETO2#)
        NETO# = Abs(NETO#)
        IVAD1# = Abs(IVAD1#)
        IVAD2# = Abs(IVAD2#)
        IVAD# = Abs(IVAD#)
        RNI# = Abs(RNI#)
        IPER = Abs(IPER)
        TOTAGEN# = Abs(TOTAGEN#)
        SIGX% = (-1)
    End If
    '
End Sub

Sub DATDEUPEN(X$, TIDOCU$, NFCX$, NFCY$, OPMOVIX$, VAMOVIX$, NC%, FFI%, FVENCI%, SIGX%, IORIG#, IPENDE#)

'**********
    '
    NC% = CVI(Left$(X$, 2))
    FFI% = CVI(Mid$(X$, 19, 2))
       If FFI% = 0 Then FFI% = FFANT%
    FFANT% = FFI%
    '
    FVENCI% = CVI(Mid$(X$, 9, 2))
    If FVENCI% < FFI% Then FVENCI% = FFI%
    '
    NRO& = CVS(Mid$(X$, 5, 4))
    NUSUCUVE% = CVI(Mid$(X$, 29, 2))
    NSUCX$ = TRIM$(Str$(NUSUCUVE%)): While Len(NSUCX$) < 4: NSUCX$ = "0" + NSUCX$: Wend
    '
    Tipo% = Asc(Mid$(X$, 3, 1)): If TI% < 0 Or TI% > 7 Then TI% = 2 ' default factura
    TICOMX$ = "FC": OPMOVIX$ = "FAC": TIDOCU$ = "Factura"
    If Tipo% = 3 Then TICOMX$ = "ND": OPMOVIX$ = "DEB": TIDOCU$ = "N.Debito"
    If Tipo% = 4 Then TICOMX$ = "RB": OPMOVIX$ = "RCOB": TIDOCU$ = "R.Cobranza"
    If Tipo% = 6 Then TICOMX$ = "NC": OPMOVIX$ = "CRE": TIDOCU$ = "N.Credito"
    If Tipo% = 0 Or Tipo% = 7 Then TICOMX$ = "AJ": OPMIVIX$ = "AJUD": TIDOCU$ = "Ajuste"
    '
    VA% = Asc(Mid$(X$, 4, 1)): If VA% < 1 Or VA% > 3 Then VA% = 1 ' default factura
    VAX$ = "A"
    If VA% = 2 Then VAX$ = "B"
    If VA% = 3 Then VAX$ = "E"
    If Tipo% = 4 Then
      VA% = 0
      VAX$ = "X"
    End If
    '
    NFCZ$ = TRIM$(Str$(NRO&)): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
    NFCX$ = TICOMX$ + "-" + Mid$(NFCZ$, 3) + "-" + VAX$
    NFCY$ = TICOMX$ + "-" + NSUCX$ + "-" + NFCZ$ + "-" + VAX$
    '
    OPMOVIX$ = "FC"
    If Tipo% = 3 Then OPMOVIX$ = "ND"
    If Tipo% = 4 Then OPMOVIX$ = "RB"
    If Tipo% = 6 Then OPMOVIX$ = "NC"
    VAMOVIX$ = VAX$
    '
    IPENDE# = (Int(100 * (CVD(Mid$(X$, 11, 8))) + 0.5)) / 100
    IORIG# = (Int(100 * CVD(Mid$(X$, 21, 8)) + 0.5)) / 100
    '
    SIGX% = 1
    If Tipo% > 3 Or IORIG# < 0 Then
       SIGX% = (-1)
    End If
    '
    IPENDE# = SIGX% * Abs(IPENDE#)
    IORIG# = SIGX% * Abs(IORIG#)
    '
End Sub

Sub GRAFACRE(Optional TIMOFIL$)
   '
   Static YAINDEXCHE%
   '
   Dim FEVEN%(10)
   Dim VENCIPESOS#(10)
   '
   Dim CAMP%, CUEMP%(2048), MEPA%(2048)
   For i& = 1 To ULTREG&("MEPAGO")
       F0$ = REGLEIDO$("MEPAGO", i&)
       CAMP% = i&
       MEPA%(CAMP%) = CVI(Left$(F0$, 2))
       CUEX$ = Mid$(F0$, 45, 12)
       CUEMP%(CAMP%) = CUEINT%(CUEX$)
       If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUEX$) Then
          Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Pago '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
          Call FINAL("")
       End If
   Next i&
   '
   PORCABA% = 0
   If ULTREG&("CAJABANC") > 0 Then
      PORCABA% = 1
   End If
   '
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI'"
   SQLX$ = SQLX$ + "AND Status = 1 "
   SQLX$ = SQLX$ + "AND OPCIMOVIM = 'CH' "
   'Set CHETEMP1 = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   Dim CHETEMP1 As ADODB.Recordset
   Set CHETEMP1 = New ADODB.Recordset
   CHETEMP1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
              '
   NOCLOSE% = 0
   Call APERBASDAT("CABAN")
   NOCLOSE% = 1
   FIN& = ULTREG&("CUECORR")
   Label1.ToolTipText = "Hasta " & FIN&
   For U& = 1 To FIN&
      Label1 = U&
      Call TRACE(20, U&, FIN&)
      Label1 = U&
      Label1.Refresh
      XX$ = REGLEIDO$("CUECORR", U&)
      TS$ = REGLEIDO$("CUECOSS", U&)
      FF% = CVI(Left$(XX$, 2))
      NC% = CVI(Mid$(XX$, 5, 2))
      NROX$ = Mid$(XX$, 7, 18)
      If InStr(UCase$(NROX$), "ANUL") > 0 Then GoTo 15 ' suprime anulados
      NRECIB$ = Left$(NROX$, 3) + Mid$(NROX$, 13, 6) + "-" + Mid$(NROX$, 8, 2)
      TIDOCOM$ = Left$(NROX$, 2)
      '
      ' ESTO ES PARA RECUPERAR AJUSTES DE SALDOS
      If TIMOFIL$ <> "" Then
         If TIMOFIL$ = "AS" And TIDOCOM$ = "FF" Then GoTo 12
         If TIDOCOM$ <> TIMOFIL$ Then
            GoTo 15
         End If
      End If
      '
12    SUC% = 0
      NRO& = XVALO(Mid$(NROX$, 11, 8))
      Observa$ = Mid$(XX$, 25, 24)
      ICONTAX$ = Mid$(XX$, 199, 80)
      '
      If TIDOCOM$ = "OP" Or TIDOCOM$ = "FF" Then
          GoSub 30
        Else
          If TIDOCOM$ = "AS" Or TIDOCOM$ = "AJ" Then
             AAA = BBB
          End If
          GoSub 20
      End If
      '
15 Next U&
   GoSub GRAPLIP
Exit Sub
   '
   '
20 NC1% = (-1) * NC%
   '
   CoComCor$ = NRECIB$
   CoComLar$ = NROX$
'   CodiEmp% = 0
'   CodiSuc% = SUC%
   NumeClie = 0
   NumeProv% = NC%
   FechEmi = CVDAT(FF%)
   '
   Referen$ = rasocli$(NC1%)
   FeContab = CVDAT(FF%)
   FeAcredi = CVDAT(FF%)
   NumeTal = 0
   NumeCom = NRO&
      If Observa$ = "" Then Observa$ = rasocli$(NC1%)
   NumCai$ = ""
   FechCai = CVDAT(FECHANUM("01/01/80"))
   InclIva$ = ""
   TOFA0# = 0
   '
   ' graba detalle por imputación de compras
   For KKI% = 1 To Len(ICONTAX$) Step 10
     CuentaCo% = CVI(Mid$(ICONTAX$, KKI%, 2))
     CentroCos$ = ""
     ItConcep$ = "" 'P/ conceptos
     ImpoDeb# = Abs(CVD(Mid$(ICONTAX$, KKI% + 2, 8)))
     If TIDOCOM$ = "AS" Then
        If KKI% = 1 Then
            ImpoDeb# = CVD(Mid$(XX$, 123, 8)) - CVD(Mid$(XX$, 115, 8))
          Else
            ImpoDeb# = 0
        End If
        If CuentaCo% < 1 Then CuentaCo% = CUEBALAJUS%
     End If
     If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
     '
     If CuentaCo% > 0 Then
       If Abs(ImpoDeb#) >= 0.005 Then
         DEMOVX$ = DENOCUE$(CuentaCo%)
         TiMovi$ = "COMPRA"
         OpMovi$ = TIDOCOM$
         VaMovi$ = Mid$(NROX$, 4, 1)
         Descrip$ = TRIM$(DEMOVX$) + " - " + rasocli$(NC1%)
         TOFA0# = TOFA0# + ImpoDeb#
         Call GRACABAN
         SIG% = Sgn(ImpoDeb#)
       End If
     End If
     CentroCos$ = ""
   Next KKI%
   '
   ' grabar importes de IVA
   For KKI% = 1 To 8
     IGRA# = SIG% * Abs(CVD(Mid$(TS$, 300 + 24 * KKI% - 23, 8)))
     VIVA# = SIG% * Abs(CVD(Mid$(TS$, 300 + 24 * KKI% - 15, 8)))
     TASANO$ = Mid$(TS$, 300 + 24 * KKI% - 7, 8)
     COTASA% = KKI%
     '
     If Abs(IGRA#) >= 0.005 Or Abs(VIVA#) >= 0.005 Then
        TiMovi$ = "IVA"
        OpMovi$ = "CrFi"
        VaMovi$ = TASANO$
        ImpoDeb# = VIVA#
        BasImpo# = IGRA#
        Descrip$ = "IVA " + TASANO$ + " " + NRECIB$
        CuentaCo% = CUECOINT%(Control$("", "CUEIVA-0"))
        TOFA0# = TOFA0# + ImpoDeb#
        Call GRACABAN
        TAREA# = 0: If Abs(IGRA#) > 0.005 Then TAREA# = 100 * Abs(VIVA#) / Abs(IGRA#)
        If TAREA# > 99.9 Then TAREA# = 0
        '
        If TAREA# > 18.5 And TAREA# < 19.5 Then TASANO$ = "19.00 %"
        If TAREA# > 20.5 And TAREA# < 21.5 Then TASANO$ = "21.00 %"
        If TAREA# > 9 And TAREA# <= 10 Then TASANO$ = "9.50 %"
        If TAREA# > 10 And TAREA# <= 11 Then TASANO$ = "10.50 %"
        '
     End If
   Next KKI%
   '
   For KKI% = 91 To 114 Step 8
     TiMovi$ = "IVA"
     OpMovi$ = "CrFi"
     VaMovi$ = ""
     Descrip$ = "IVA Credito Fiscal " + NRECIB$
     ImpoDeb# = Abs(CVD(Mid$(XX$, KKI%, 8)))
     If TIDOCOM$ = "AS" Then
        ImpoDeb# = ImpoDeb# * Sgn(CVD(Mid$(XX$, 123, 8)) - CVD(Mid$(XX$, 115, 8)))
     End If
     If Abs(ImpoDeb#) < 0.005 Then GoTo 29
     '
     BasImpo# = TOFA7#
     If KKI% = 91 Then
         VaMovi$ = "Per-IVA"
         Descrip$ = "Percepcion I.V.A. " + NRECIB$
         CuentaCo% = CUECOINT%(Control$("", "CUEIVA-5"))
       ElseIf KKI% = 99 Then
         TiMovi$ = "GAN"
         VaMovi$ = "Per-Gan"
         Descrip$ = "Percepcion Ganancias " + NRECIB$
         CuentaCo% = CUECOINT%(Control$("", "CUEIVA-6"))
       ElseIf KKI% = 107 Then
         TiMovi$ = "I-BRU"
         VaMovi$ = "PERC"
         Descrip$ = "Percepcion I.Brutos " + NRECIB$
' ojo que esto está mal
         CuentaCo% = CUEIPROVIN%("B")
     End If
     '
     If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
     If Abs(ImpoDeb#) >= 0.005 Then
       TOFA0# = TOFA0# + ImpoDeb#
       Call GRACABAN
     End If
     BasImpo# = 0
29 Next KKI%
'''   '
   ' graba factura en c.corriente proveedor
   CuentaCo% = CUECOINT%(CUEMADRE$(-1 * NC%))
   ImpoDeb# = (-1) * TOFA0#
     'If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then ImpoDeb# = (-1) * ImpoDeb#
   TiMovi = "FCOM"
   OpMovi = Left$(Text11, 2)
   VaMovi = Label20
   Descrip$ = CoComLar$ + " - " + rasocli$(NC1%)
   FeAcredi = CVDAT(FF%)
   Call GRACABAN
   '
   ' Grabar Vencimientos de Facturas ******************************************
   SIGNO% = 1
   If TIDOCOM$ = "NC" Or TIDOCOM$ = "DP" Then SIGNO% = (-1)
     'IBruTot# = Abs(TOFAC#(0))
   IBruTot# = TOFA0#
   IIvaTot# = SUMIVA#
   '
   PRICONVEN% = -1
   SALBRUTO# = TOFA0# * SIGNO%
   TXVENCI$ = Mid$(XX$, 139, 60)
   For KKI% = 1 To 6
      FEVEN%(KKI%) = CVI(Mid$(TXVENCI$, 10 * KKI% - 9, 2))
      VENCIPESOS#(KKI%) = CVD(Mid$(TXVENCI$, 10 * KKI% - 7, 8))
   Next KKI%
   '
   For KKI% = 1 To 6
      If FEVEN(KKI%) <> 0 Then
        If Abs(VENCIPESOS#(KKI%)) >= 0.005 Then
           NuorVen% = KKI%
           FeVenc = CVDAT(FEVEN(KKI%))
           ImVenc# = VENCIPESOS#(KKI%)
           Call GRACREPEN
         End If
      End If
   Next KKI%
   '
Return
'
30 ' GRABA ORDENES DE PAGO
    IDEBEX# = CVD(Mid$(XX$, 115, 8))
    IHABEX# = CVD(Mid$(XX$, 123, 8))
    DIFEPAG# = 0
    'If Abs(IDEBEX#) + Abs(IHABEX#) < 0.005 Then
    '   GoTo 399    ' ESTÁ ANULADA - PERO SE GRABA IGUAL
    'End If
    '
    CoComCor$ = NRECIB$
    CoComLar$ = NROX$
    IMPOP# = ROUND(IDEBEX# - IHABEX#)
'    CodiEmp% = 0
'    CodiSuc% = SUC%
    NumeClie = 0
    NumeProv% = NC%
    FechEmi = CVDAT(FF%)
    Referen$ = rasocli$((-1) * NC%)
    FeContab = CVDAT(FF%)
    FeAcredi = CVDAT(FF%)
    CuentaCo% = CUECOCLI%(NC1%)
    TiMovi$ = "OPAG"
    OpMovi$ = "OP"
    VaMovi$ = "CCORR"
    Descrip$ = "Orden de Pago " + NRECIB$
    '
    If Left$(NRECIB$, 1) = "A" Then ' PARA AJUSTES DE SALDO
       TiMovi = "AJUP"              ' AJUSTE DE SALDO PROVEEDORES
       OpMovi = "AS"                ' AJUSTE DE SALDO
       VaMovi = "CCORR"             ' CUENTA CORRIENTE
       Descrip = "Aj.Saldo - " + rasocli$((-1) * NC%)
    End If
    '
    NumeTal = COBRACT%
    NumeCom = NRO&
    StatInit = 0
    ImpoDeb# = IMPOP#
    BasImpo# = IMPOP#    ' aqui va entonces el neto sin IVA
    Observa$ = ""
    If Observa$ = "" Then Observa$ = rasocli$(NC1%)
    '
    'DIFEPAG# = DIFEPAG# + ImpoDeb#
    If TICUEPRO%((-1) * NC%) = 0 Then
         DIFEPAG# = DIFEPAG# + ImpoDeb#
         Call GRACABAN
       Else
         RFG$ = RECFILT$("DETPAGAS", 3, MKS$(NRO&))
         If RFG$ = "" Then
            DIFEPAG# = DIFEPAG# + ImpoDeb#
            Call GRACABAN
            GoTo 32
         End If
         For K2& = 1 To Len(RFG$) Step 4
           K1& = CVS(Mid$(RFG$, K2&, 4))
           Z$ = REGLEIDO$("DETPAGAS", K1&)
           If CVI(Mid$(Z$, 7, 2)) = FF% Then
             CuentaCo% = CVI(Mid$(Z$, 17, 2))
             ImpoDeb# = CVD(Mid$(Z$, 9, 8))
             VaMovi$ = "CGAST"
             Referen$ = DENOCUE$(CuentaCo%)
             CentroCos$ = ""
             ItConcep$ = "" 'P/ conceptos
             DIFEPAG# = DIFEPAG# + ImpoDeb#
             Call GRACABAN
           End If
         Next K2&
    End If
    '
    'grabar cheques entregados
32  If YAINDEXCHE% = 0 Then
        XX1$ = REGLEIDO$("CHEQUEMI", 1)
        Call GRAREG("CHEQUEMI", XX1$, 1)
        Call CIERRARCH("CHEQUEMI")
        YAINDEXCHE% = 1
    End If
    '
    RFGH$ = RECFILT$("CHEQUEMI", 6, MKS$(NRO&))
    If PORCABA% = 1 Then
       RFGH$ = RECFILT$("CAJABANC", 1, NRECIB$)
    End If
    For KKUU& = 1 To Len(RFGH$) Step 4
       RECHE& = CVS(Mid$(RFGH$, KKUU&, 4))
       If PORCABA% = 0 Then
            Z$ = REGLEIDO$("CHEQUEMI", RECHE&)
         Else
            Z$ = REGLEIDO$("CAJABANC", RECHE&)
            IMCHE# = CVD(Mid$(Z$, 161, 8))
            FFX1% = CVI(Mid$(Z$, 61, 2))
       End If
       If (PORCABA% = 0 And CVI(Mid$(Z$, 1, 2)) = NC%) Or (PORCABA% = 1 And IMCHE# > 0 And FFX1% = FF%) Then ' COINCIDE PROVEEDOR
          BANCO$ = AJUSTI$(Mid$(Z$, 15, 16), 18)
          NUCHE$ = AJUSTI$(Mid$(Z$, 3, 12), 14)
            FEX = CVI(Mid$(Z$, 49, 2))
            FEVEN1% = FEX
          FEVECH$ = FECHATEX$(FEX)
          IDEVA$ = BANCO$ + NUCHE$ + FEVECH$
          IMCHE# = CVD(Mid$(Z$, 33, 8))
          FECHE% = CVI(Mid$(Z$, 31, 2))
          If PORCABA% = 1 Then
             FECHE% = CVI(Mid$(Z$, 111, 2))
             FEVEN1% = CVI(Mid$(Z$, 63, 2))
             IDEVA$ = Mid$(Z$, 71, 40)
             IMCHE# = CVD(Mid$(Z$, 161, 8))
          End If
          '
          'BUSCAR SI ESTÁ INGRESADO
          ' CASO AFIRMATIVO:
          ' 1) MARCAR LA SALIDA DE CARTERA
          ' 2) PONER CUENTA = CUENTA DE CARTERA DE VALORES
          VUELTA% = 0: ENCUEN% = 0
          IMCHEX$ = TRIM$(FORMATNUM$(IMCHE#, "F14.2"))
             If IMCHEX$ = "" Then IMCHEX$ = "0"
          SQLX$ = "DescriMov = '" & IDEVA$ & "' "
          SQLX1$ = SQLX$
          SQLX$ = SQLX$ + "AND ABS(ValAbsComp - " & IMCHEX$ & ") < 0.005 "
          'SQLX$ = SQLX$ + "AND Status = 1"
          '
          SQLZ$ = "SELECT * FROM CAJABANC "
          SQLZ$ = SQLZ$ + "WHERE TipoMovim = 'CHRECI'"
          SQLZ$ = SQLZ$ + "AND Status = 1 "
          SQLZ$ = SQLZ$ + "AND OPCIMOVIM = 'CH' "
          SQLZ$ = SQLZ$ + "AND DescriMov = '" & IDEVA$ & "' "
          SQLZ$ = SQLZ$ + "AND ABS(ValAbsComp - " & IMCHEX$ & ") < 0.005 "
          'Set CHETEMP11 = CueCorri.OpenRecordset(SQLZ$, dbOpenDynaset)
          Dim CHETEMP11 As ADODB.Recordset
          Set CHETEMP11 = New ADODB.Recordset
          CHETEMP11.Open FILTSQL$(SQLZ$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          
          With CHETEMP11
             On Error Resume Next
             .MoveFirst
             If Err < 1 Then
                On Error GoTo 0
                '.Edit
                !FEDEPOSENT = CVDAT(FF%)
                !Document = NRECIB$
                !DESTICHEQ = rasocli$((-1) * NC%)
                  CUECONTABINT% = !CUECONTAI  ' PARA QUE LA SALIDA SEA DE LA MISMA CUENTA QUE ENTRO
                !REFESALCHE = "Orden de Pago Nro." + Str$(NRO&)
                !IDENSALCHE = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
                !UsuaSalChe = AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
                !Status = 2                     ' Marca Salida de Cartera
                !IdenBaja = IDACTU$             ' Marca Modificacion de registro - FECHA - HORA - USUARIO - TERMINAL
                .Update
                ENCUEN% = 1
             End If
             On Error GoTo 0
          End With
          '
          ' ESTO SI NO ESTÁ INGRESADO
          If ENCUEN% < 1 Then
             QUE% = CVI(Mid$(Z$, 45, 2))
             For UX% = 1 To CAMP%
                 If MEPA%(UX%) = QUE% Then GoTo 38
             Next UX%
             ' PONER AQUI UNA OPCIÓN POR DEFAULT
             UX% = 3   ' ESTO ES EFECTIVO EN FERVI
             '
38           CUECONTABINT% = CUEMP%(UX%)
          End If
          '
          NC1% = (-1) * NC%
          IMPOP# = IMPOP# - IMCHE#
          CoComCor$ = NRECIB$
          CoComLar$ = NROX$
'          CodiEmp% = 0
'          CodiSuc% = SUC%
          NumeClie = 0
          NumeProv% = NC%
          FechEmi = CVDAT(FECHE%)
            FEX = FF%
          Referen$ = rasocli$(NC1%) + " - " + FECHATEX$(FEX)
          FeContab = CVDAT(FF%)
          '
          FeAcredi = CVDAT(FEVEN1%)
          CuentaCo% = CUECONTABINT%
          TiMovi$ = "CHENTR"
          OpMovi$ = "OP"
          VaMovi$ = ""
          StatInit = 0
          Descrip$ = IDEVA$
          NumeTal = COBRACT%
          NumeCom = NRO&
          ImpoDeb# = (-1) * Abs(IMCHE#)
          BasImpo# = 0               ' CERO PARA LA SALIDA DEL CHEQUE
          Observa$ = "O.Pago " + TRIM$(Str$(NRO&)) + " - " + rasocli$(NC1%)
          '
          DIFEPAG# = DIFEPAG# + ImpoDeb#
          Call GRACABAN
           '
       End If
    Next KKUU&
    '
    If Abs(DIFEPAG#) > 0.005 Then
       Referen$ = "Dif.Importación " + NRECIB$
       CuentaCo% = CUEBALPAGO%
       TiMovi$ = "AJUSTE"
       OpMovi$ = ""
       VaMovi$ = ""
       Descrip$ = "Dif.Importación " + NRECIB$
       ImpoDeb# = (-1) * DIFEPAG#
       Observa$ = ""
       Call GRACABAN
    End If
    '
   ' Grabar Vencimientos de Facturas ******************************************
39 SIGNO% = (-1)
   IBruTot# = Abs(ROUND(CVD(Mid$(XX$, 59, 8))))
   IIvaTot# = 0
   '
   PRICONVEN% = -1
   SALBRUTO# = Abs(IBruTot#) * SIGNO%
   TXVENCI$ = Mid$(XX$, 139, 60)
   For KKI% = 1 To 6
      FEVEN%(KKI%) = CVI(Mid$(TXVENCI$, 10 * KKI% - 9, 2))
      VENCIPESOS#(KKI%) = CVD(Mid$(TXVENCI$, 10 * KKI% - 7, 8))
   Next KKI%
   '
   For KKI% = 1 To 6
      If FEVEN(KKI%) <> 0 Then
        If Abs(VENCIPESOS#(KKI%)) >= 0.005 Then
           NuorVen% = KKI%
           FeVenc = CVDAT(FEVEN(KKI%))
           ImVenc# = VENCIPESOS#(KKI%)
           Call GRACREPEN
         End If
      End If
   Next KKI%
   '
399 Return

GRAPLIP:
  With Aplipago
    On Error Resume Next
1301   .MoveFirst
    If Err = 0 Then
      .Delete
      GoTo 1301
    End If
    On Error GoTo 0
  End With
  '
  FIN& = ULTREG&("APLIPAG")
  For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("APLIPAG", U&)
     NumeProv% = CVI(Mid$(XX$, 3, 2))
     CoComLar$ = Mid$(XX$, 5, 18)
        NROX$ = CoComLar$
     CoComCor$ = Left$(NROX$, 3) + Mid$(NROX$, 9, 1) + "-" + Mid$(NROX$, 14)
     ImpApli# = CVD(Mid$(XX$, 23, 8))
     RefApli$ = Mid$(XX$, 31, 24)
     DocOrLar$ = Mid$(XX$, 55, 18)
        NROX$ = DocOrLar$
     DocOrCor$ = Left$(NROX$, 3) + Mid$(NROX$, 9, 1) + "-" + Mid$(NROX$, 14)
     TotOrig# = CVD(Mid$(XX$, 73, 8))
     IvaOrig# = CVD(Mid$(XX$, 81, 8))
     Call GRAPLIPAG
   Next U&
Return

End Sub

Sub MIGRACLIEN()
  Picture1.ToolTipText = "Depuración - Clientes"
  '
  Call APERBASDAT("CLIEN")
'  On Error Resume Next
10 Call BORRAREGISTROS("DEUDOR12", "NUME_CLI <> 0 OR NUME_CLI = 0", REGAF&, "NOMESS")
'   FLEXCONN.Execute (FILTSQL$("DROP TABLE DEUDOR12"))
'  On Error GoTo 0
  '
  Call APERBASDAT("CLIEN")
  Call GENERBASE("BADACLIEN")
  REGSINBORRAR& = CantRegistros("DEUDOR12", "NUME_CLI > 0", "NOMESS")
  If REGSINBORRAR& > 0 Then
     TEXTO1$ = ("Se Han Detectado En la Tabla Deudor12\Contiene " & REGSINBORRAR& & " Registros")
     OPX% = COMALTER%(TEXTO1$ + "\\Volver a Intentar\Cancelar")
     If OPX% < 1 Or OPX% > 1 Then Exit Sub
     If OPX% = 1 Then GoTo 10
  End If
  '
  Dim DEUDORRR As ADODB.Recordset
  Set DEUDORRR = New ADODB.Recordset
  SQLX$ = "SELECT * FROM DEUDOR12"
  DEUDORRR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  '
  Picture1.ToolTipText = "Conversión de Clientes"
  '
  'LECTURA
  UR& = ULTREG&("DEUDOR1")
  For U& = 1 To ULTREG&("DEUDOR1")
    '
    X1$ = REGLEIDO$("DEUDOR1", U&)
      NCI% = CVI(Mid$(X1$, 1, 2))
    X2$ = REGLEIDO$("DEUDOR2", U&)
    X3$ = FILTERGETRECORD$("BADACLIE", 1, MKI$(NCI%))
    'LECTURA
    RASO$ = Mid$(X3$, 3, 62)
    FANTA$ = Mid$(X3$, 65, 32)
    DOMI$ = Mid$(X3$, 97, 64)
    CPOSTX$ = Mid$(X3$, 161, 16)
    CPOST% = 0: If XVALO(CPOSTX$) < 32767 Then CPOST% = XVALO(CPOSTX$)
    LOCA$ = TRIM$(Mid$(X3$, 177, 48))
    PAISOK$ = Mid$(X3$, 225, 32)
    TELE$ = Mid$(X3$, 257, 24)
    FAXX$ = TRIM$(Mid$(X3$, 281, 24))
    CONTACTX$ = TRIM$(Mid$(X3, 385, 48))
    MAILPRO$ = TRIM$(Mid$(X3$, 305, 48))
    HORARIX$ = TRIM$(Mid$(X3$, 449, 32))
    ZONCLI$ = Mid$(X3$, 481, 12)
    Cuit$ = TRIM$(Mid$(X2$, 75, 15))
      If Cuit$ <> "" Then VCC% = VALICUIT%(Cuit$)
    GRUCLI% = CVI(Mid$(X2$, 125, 2))
    '
    NCITXT$ = CStr(NCI%)
    Do Until Len(NCITXT$) = 5
      NCITXT$ = "0" + NCITXT$
    Loop
    CODICUEN$ = "CL-" & NCITXT$
    '
'      RASO$ = RASOCLI1$(NCI%)
'      FANTA$ = TEL2CLI1$(NCI%)
'      DOMI$ = DOMICLI1$(NCI%)
'      CPOSTX$ = CPOSCLI1$(NCI%)
'
'      LOCA$ = LOCACLI1$(NCI%)
'      PAISOK$ = PAISCLI1$(NCI%)
'      TELE$ = TEL1CLI1$(NCI%)
'      FAXX$ = FAXCLI1$(NCI%)
'      CONTACTX$ = CONTACTCLI1$(NCI%)
'      MAILPRO$ = EMAILCLI1$(NCI%)
'      HORARIX$ = HORAENT1$(NCI%)
'      ZONCLI$ = ZONACLI1$(NCI%)
'      CUIT$ = CUITCLI1$(NCI%)
'      If CUIT$ <> "" Then VCC% = VALICUIT%(CUIT$)
'      GruCli% = GRUCOCLI1%(NCI%)
'      '
'      NCITXT$ = CStr(NCI%)
'      Do Until Len(NCITXT$) = 5
'         NCITXT$ = "0" + NCITXT$
'      Loop
'      CODICUEN$ = "CL-" & NCITXT$
'      '
'      GRUDIV% = 0
    
    TELEF2$ = Mid$(X2$, 65, 10)
    PIVA% = Asc(Mid$(X2$, 104, 1))
    Condi% = Asc(Mid$(X2$, 118, 1))
    PROVI$ = TRIM$(Mid$(X2$, 117, 1))
    LIPRE% = Asc(Mid$(X2$, 119, 1))
    COVEN% = Asc(Mid$(X2$, 120, 1))
    LICRE = CVS(Mid$(X2$, 121, 4))
    ingbru$ = Mid$(X2$, 90, 13)
    CAIBRU$ = Mid$(X2$, 103, 1)
    CUEMA% = CVI(Mid$(X2$, 127, 2))
    GRUDIV% = 0
    '
    DOMIENT1$ = ""
    LOCAENT1$ = ""
    TRANSPOR$ = ""
    NTRANX% = 0
    HORARIENT$ = ""
    If NCI% > 0 Then
      XX$ = FILTERGETRECORD$("LENTRECL", 1, MKI$(NCI%))
      DOMIENT1$ = TRIM$(Mid$(XX$, 3, 32))
      LOCAENT1$ = TRIM$(Mid$(XX$, 35, 32))
      TRANSPOR$ = TRIM$(Mid$(XX$, 67, 32))
      NTRANX% = CVI(Mid$(XX$, 99, 2))
         If NTRANX% = 8224 Then NTRANX% = 0
      HORARIENT$ = TRIM$(Mid$(XX$, 101, 24))
    End If
    '
    DESSTAN$ = ""
    XXX$ = FILTERGETRECORD$("DESCUCLI", 1, MKI$(NCI%))
    DESSTAN$ = Mid$(XXX$, 3, 14)
    '
    ' anotaciones
    RNC$ = RECFILT$("NOTACLI", 1, MKI$(NCI%))
    URN& = ULTREG&("NOTACLI")
    '
    NOTAC$ = ""
    For UKL& = 1 To Len(RNC$) Step 4
       RENOTA1& = CVS(Mid$(RNC$, UKL&, 4))
       If RENOTA1& > 0 Then
          TTYY$ = TRIM$(Mid$(REGLEIDO$("NOTACLI", RENOTA1&), 3))
          If NOTAC$ <> "" Then
             NOTAC$ = NOTAC$ + Chr$(13) + Chr$(10)
          End If
          NOTAC$ = NOTAC$ + RTrim$(TTYY$)
       End If
    Next UKL&
    '
    'GRABACION
    On Error GoTo 0
    With DEUDORRR
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
      !Stat_Cli = 1                        'STATUS ACTIVO
      !nume_cli = NCI%                     'CODIGO ENTERO
      !Codi_Cli = CODICUEN$                'CODIGO TEXT
      !Tcpr_Cli = 0                        'TIPO CTA
      !GrCo_Cli = GRUCLI%                  'TIPO CLI  VA A GRUPO CLI
      !raso_cli = RASO$                    'RAZON SOCIAL
      !Fant_Cli = Left$(FANTA$, 10)        'FANTASIA
      !domi_cli = DOMI$                    'DIRECCION
      !loca_Cli = LOCA$                    'LOCALIDAD
      !Prov_Cli = PROVI$                   'COD PROVINCIA
      !Pais_Cli = PAISOK$                  'PAIS
      !Cpos_Cli = CPOSTX$                  'CODIGO POSTAL
      !Tele_Cli = TELE$                    'TELEFONO
'      !Tele2_Cli = TELEF2$                '2º TELEFONO
      !Faxi_Cli = FAXX$                    'FAX
      !Cuit_Cli = Cuit$                    'CUIT SIN VALIDAR
      !Cpag_Cli = Condi%                   'CONDICION DE PAGO
      !Nibr_Cli = ingbru$                  'IMP BRUTOS
      !Caib_Cli = CAIBRU$
      !Mail_Cli = MAILPRO$                 'MAIL
      !Piva_Cli = PIVA%                    'POS IVA
      '!Posivat_Cli = PIVATXT$             'POSICION DEL IVA EN TEXTO
      '!C_Ret = CRET$                      'RETENCION
      !Hora_Cli = HORARIX$
      !Lcre_Cli = LICRE                    'LIMITE CREDITO
      !Vend_Cli = COVEN%                   'CODIGO VENDEDOR
      !CTAC_CLI = CONTACTX$                'CONTACTO
      !Dsta_Cli = DESSTAN$                 'DESCUENTO
      !Lentre0_Cli = DOMIENT1$             'LUGAR DE ENTREGA
      !Lentre1_Cli = LOCAENT1$             'LOCALIDAD
      !Lentre2_Cli = TRANSPOR$             'TRANSPORTE
      !NTrans_Cli = NTRANX%
      !Hora_Ent = HORARIENT$
      !Zona_Cli = ZONCLI$
      !Obser_Entre = ""
      !Recar_Embal = 0
      !Gru_Div = GRUDIV%                   'GRUPO DIVISIONAL
      !CuMa_Cli = CUEMA%                   'CTA CONTABLE MADRE
      !Lpre_Cli = LIPRE%                   'LISTA DE PRECIOS
      !nota_Cli = NOTAC$                   'ANOTACION
      .Update
      Call TRACE(20, U&, UR&)
    End With
  Next U&
  '
  DEUDORRR.Close
  Set DEUDORRR = Nothing
  '
End Sub

Sub MIGRAPROVE()
   '
19 Picture1.ToolTipText = "Del - Proveedores"
  '
  Call APERBASDAT("PROVE")
'  On Error Resume Next
10 Call BORRAREGISTROS("PROVED12", "NUME_CLI <> 0 OR NUME_CLI = 0", REGAF&, "NOMESS")
'   FLEXCONN.Execute (FILTSQL$("DROP TABLE PROVED12"))
'  On Error GoTo 0
  '
  Call APERBASDAT("PROVE")
  Call GENERBASE("BADAPROVE")
  
  REGSINBORRAR& = CantRegistros("PROVED12", "NUME_CLI > 0", "NOMESS")
  If REGSINBORRAR& > 0 Then
     TEXTO1$ = ("Se Han Detectado En la Tabla PROVED12\Contiene " & REGSINBORRAR& & " Registros")
     OPX% = COMALTER%(TEXTO1$ + "\\Volver a Intentar\Cancelar")
     If OPX% < 1 Or OPX% > 1 Then Exit Sub
     If OPX% = 1 Then GoTo 10
  End If
  '
  Dim PROVEDRRR As ADODB.Recordset
  Set PROVEDRRR = New ADODB.Recordset
  SQLX$ = "SELECT * FROM PROVED12"
  PROVEDRRR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  '
  Picture1.ToolTipText = "Conversión de Proveedores"
  '
  'LECTURA
   UR& = ULTREG&("PROVED1")
  For U& = 1 To ULTREG&("PROVED1")
    '
    X1$ = REGLEIDO$("PROVED1", U&)
      NCI% = CVI(Mid$(X1$, 1, 2))
    X2$ = REGLEIDO$("PROVED2", U&)
    X3$ = FILTERGETRECORD$("BADAPROV", 1, MKI$(NCI%))
   'LECTURA
    NCITXT$ = CStr(NCI%)
    Do Until Len(NCITXT$) = 5
      NCITXT$ = "0" + NCITXT$
    Loop
    CODICUEN$ = "PR-" & NCITXT$
    '
    RASO$ = TRIM$(Mid$(X3$, 3, 62))
      If RASO$ = "" Then RASO$ = TRIM$(Mid$(X1$, 3, 30))
    FANTASI$ = TRIM$(Left(Mid$(X3$, 65, 32), 10))
      If FANTASI$ = "" Then FANTASI$ = TRIM$(Mid$(X2$, 65, 10))
    DOMI$ = TRIM$(Mid$(X3$, 97, 64))
      If DOMI$ = "" Then DOMI$ = TRIM$(Left$(X2$, 25))
    CPOSTX$ = TRIM$(Mid$(X3$, 161, 16))
      If CPOSTX$ = "" Then CPOSTX$ = Str$(CVI(Mid$(X2$, 28, 2)))
    CPOST% = 0: If XVALO(CPOSTX$) < 32767 Then CPOST% = XVALO(CPOSTX$)
    LOCA$ = TRIM$(Mid$(X3$, 177, 48))
      If LOCA$ = "" Then LOCA$ = TRIM$(Mid$(X2$, 30, 25))
    PAISOK$ = Mid$(X3$, 225, 32)
    TELE$ = TRIM$(Mid$(X3$, 257, 24))
      If TELE$ = "" Then TELE$ = TRIM$(Mid$(X2$, 55, 10))
    FAXX$ = Mid$(X3$, 281, 24)
    MAILPRO$ = Mid$(X3$, 305, 48)
    CONTACTX$ = Mid$(X3$, 385, 32)
    TIPOCTA% = Asc(Mid$(X2$, 119, 1))
    Cuit$ = TRIM$(Mid$(X2$, 75, 15))
      If Cuit$ <> "" Then VCC% = VALICUIT%(Cuit$)
    NUIBRU$ = TRIM$(Mid$(X2$, 90, 13))
    CODIPROVI1$ = Mid$(X2$, 117, 1)
    CODIVA1% = Asc(Mid$(X2$, 120, 1))
    IBRU$ = ""
    Condi% = Asc(Mid$(X2$, 118, 1))
    LIPRE% = 0
    COVEN% = 0
    LICRE = 0
    CUEMA% = CVI(Mid$(X2$, 127, 2))
    '
    TRT$ = FILTERGETRECORD$("TARETIB", 1, MKI$(NCI%))
    CATIBX$ = Mid$(TRT$, 55, 1)

    ' anotaciones
    RNC$ = RECFILT$("NOTAPRO", 1, MKI$(NCI%))
    URN& = ULTREG&("NOTACLI")
    '
    NOTAC$ = ""
    For UKL& = 1 To Len(RNC$) Step 4
       RENOTA1& = CVS(Mid$(RNC$, UKL&, 4))
       If RENOTA1& > 0 Then
          TTYY$ = TRIM$(Mid$(REGLEIDO$("NOTAPRO", RENOTA1&), 3))
          If NOTAC$ <> "" Then
             NOTAC$ = NOTAC$ + Chr$(13) + Chr$(10)
          End If
          NOTAC$ = NOTAC$ + RTrim$(TTYY$)
       End If
    Next UKL&
    '
   'GRABACION
   With PROVEDRRR
    .AddNew   ' YA CONVERTIDO
    On Error Resume Next
    !CodiEmpr = CodiEmp%
    On Error GoTo 0
    !nume_cli = NCI%                     'CODIGO ENTERO
    !Stat_Cli = 1                        'STATUS ACTIVO
    !Codi_Cli = CODICUEN$                'CODIGO TEXT
    !Tcpr_Cli = TIPOCTA%                 'TIPO CTA
    !raso_cli = RASO$                    'RAZON SOCIAL
    !Fant_Cli = FANTASI$                 'FANTASIA
    !domi_cli = DOMI$                    'DIRECCION
    !loca_Cli = LOCA$                    'LOCALIDAD
    !Prov_Cli = CODIPROVI1$              'COD PROVINCIA
    !Pais_Cli = PAISOK$                  'PAIS
    !Cpos_Cli = CPOSTX$                  'CODIGO POSTAL
    !Tele_Cli = TELE$                    'TELEFONO
    '!Tele2_Cli = TELEF2$                '2º TELEFONO
    !Faxi_Cli = FAXX$                    'FAX
    !Cuit_Cli = Cuit$                    'CUIT SIN VALIDAR
    !Cpag_Cli = Condi%                   'CONDICION DE PAGO
    !Nibr_Cli = NUIBRU$                  'IMG BRUTOS
    !Caib_Cli = CATIBX$
    !Mail_Cli = MAILPRO$                 'MAIL
    !Piva_Cli = CODIVA1%                 'POS IVA
    '!C_Ret = CRET$                      'RETENCION
    !CTAC_CLI = CONTACTX$
    !CuMa_Cli = CUEMA%                   'CUENTA MADRE
    !nota_Cli = NOTAC$                   'ANOTACION
    .Update
    Call TRACE(20, U&, UR&)
   End With
  Next U&
  '
  PROVEDRRR.Close
  Set PROVEDRRR = Nothing
End Sub

Sub MIGRAMASTER()
   '
29 Picture1.ToolTipText = "Del - Articulos"
  '
  Call APERBASDAT("MASTER")
'  On Error Resume Next
10 Call BORRAREGISTROS("MASTER", "CODINT <> 0 OR CODINT = 0", REGAF&, "NOMESS")
'  SQLX$ = "DROP TABLE MASTER"
'  FLEXCONN.Execute SQLX$
'  On Error GoTo 0
  '
  Call APERBASDAT("MASTER")
  Call GENERBASE("ARTICULOS")
  '
  REGSINBORRAR& = CantRegistros("MASTER", "CODINT > 0", "NOMESS")
  If REGSINBORRAR& > 0 Then
     TEXTO1$ = ("Se Han Detectado En la Tabla MASTER\Contiene " & REGSINBORRAR& & " Registros")
     OPX% = COMALTER%(TEXTO1$ + "\\Volver a Intentar\Cancelar")
     If OPX% < 1 Or OPX% > 1 Then Exit Sub
     If OPX% = 1 Then GoTo 10
  End If
  
  Dim MASTERRR As ADODB.Recordset
  Set MASTERRR = New ADODB.Recordset
  SQLX$ = "SELECT * FROM MASTER"
  MASTERRR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  '  Picture1.ToolTipText = "Conversión de Articulos"
  '
  'LECTURA
  FIN& = ULTREG&("MASTER")
  MRECU% = 0
  For U& = 2 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("MASTER", U&)
     CI% = CVI(Mid$(X$, 1, 2))
     GIVA% = Asc(Mid$(X$, 3, 1))
     CODIT$ = Mid$(X$, 77, 16)
     CODALT$ = ""
     DESCOMP$ = ""
     DESIT$ = Mid$(X$, 5, 64)
     STATUIII% = 1
     If Mid$(X$, 92, 1) = "`" Then
        CODIT$ = TRIM$(Left$(CODIT$, 15))
        STATUIII% = (-1)
     End If
     TIMAT% = Asc(Mid$(X$, 93, 1))
     CRITREP1% = Asc(Mid$(X$, 94, 1))
     UNIME$ = TRIM$(Mid$(X$, 109, 2))
     STOMI = CVI(Mid$(X$, 105, 2))
       If STOMI > 30000 Then STOMI = (-1) * STOMI / 100
       If STOMI < (-30000) Then STOMI = (-30000)
     LECON = CVI(Mid$(X$, 107, 2))
       If LECON > 30000 Then LECON = (-1) * LECON / 100
       If LECON < (-30000) Then LECON = (-30000)
     LOMAX = CVI(Mid$(X$, 103, 2))
     'LOMAX = 1
       If LOMAX < Abs(LECON) Then LOMAX = LECON
     KILUNI = CVS(Mid$(X$, 99, 4))
     KILMET = ""
     COSUN = CVS(Mid$(X$, 111, 4))
     MONEC% = Asc(Mid$(X$, 95, 1))
     FECOU% = CVI(Mid$(X$, 115, 2))
       PRH = CVI(Mid$(X$, 123, 2))
       While PRH > 30000: PRH = PRH - 10000: Wend
     PRHAB% = PRH
     '
     GRUCO% = CVI(Mid$(X$, 125, 2))
     '
     FCP = CVI(Mid$(X$, 127, 2)) / 100   ' % DE NACIONALIZACION
     '
     REV1$ = TRIM$(Mid$(X$, 117, 4))
     FERE1% = CVI(Mid$(X$, 121, 2))
     '
     URN& = ULTREG&("ESPECIFI")
     ESPEC$ = ESPECIFI$(CI%)
     '
     ATRIA$ = ""
     REG& = Asc(Mid$(X$, 97, 1))
     If REG& > 0 Then ATRIA$ = TRIM$(REGLEIDO$("TATALLES", REG&))
     '
     ATRIB$ = ""
     REG& = Asc(Mid$(X$, 96, 1))
     If REG& > 0 Then ATRIB$ = TRIM$(REGLEIDO$("TACOLOR", REG&))
     '
     REGISTROLA$ = FILTERGETRECORD$("DESCRILA", 1, MKI$(CI%))
     DESCRI2$ = TRIM$(Mid(REGISTROLA$, 3, 64))
     DESCOMP$ = REPLACAR(DESCRI2$, "'", "")
     On Error GoTo 0
     '
32   If CODINT%(CODIT$) > 0 Then
       MRECU% = MRECU% + 1
       CODIT1$ = "zrecu-" + TRIM$(Str$(MRECU%))
       Call COMUNI("Codigo '" + CODIT$ + "' Pre-existente en BD.\Reemplazado por '" + CODIT1$ + "'.")
       CODIT$ = CODIT1$
       GoTo 32 ' XQ SINO  ENTRA EN UN LOOP SIN FIN
     End If
     '
     If CODINT%(CODIT$) = 0 Then
       With MASTERRR
         .AddNew   ' YA CONVERTIDO
         On Error Resume Next
         !CodiEmpr = CodiEmp%
         On Error GoTo 0
         !CODINT = CI%
         !NPLANO = CODALT$
         !DESCRIDOS = DESCOMP$
         !TALLE = ATRIA$
         !COLOR = ATRIB$
         !Codigo = CODIT$
         !Descripcion = DESIT$
         !Status = STATUIII%
         !UMEDIDA = AJUSTI$(UNIME$, 4)
         !TIMATE = TIMAT%
         !CRITREP = CRITREP1%
         !Monecos = MONEC%
         !FECOSTO = CVDAT(FECOU%)
         !COSUNI = COSUN
         !STOMIN = STOMI
         !LOREPOS = LECON
         !STOMAX = LOMAX
         !PROVHABI = PRHAB%
         !Alias = " "
         !DESCOMPO = 0
         !GRUPIVA = GIVA%
         !Grucoven = GRUCO%
         !CLinea = " "
         !Cfamilia = " "
         !Revision = REV1$
         !FEREVIS = CVDAT(FERE1%)
         !Observa = " "
         !ESPECIFICACION = ESPEC$
         !PESUNI = KILUNI
         !PESMETRO = 0 'KILMET
         !PORNACIO = FCP
        .Update
       End With
       '
     End If
     '
     Call TRACE(20, U&, FIN&)
   Next U&
  '
  MASTERRR.Close
  Set MASTERRR = Nothing
  '
End Sub

Sub MIGRAFORMULA()
  '
  Picture1.ToolTipText = "Del - Formulas"
  '
  Call APERBASDAT("MASTER")
'  On Error Resume Next
10 Call BORRAREGISTROS("FORMULAS", "CODICONJ <> 0 OR CODICONJ = 0", REGAF&, "NOMESS")

'  SQLX$ = "DROP TABLE FORMULAS"
'  FLEXCONN.Execute SQLX$
'  On Error GoTo 0
  '
  Call ABRESTRUC
  Call GENERBASE("ARTICULOS")
  '
  REGSINBORRAR& = CantRegistros("FORMULAS", "CODICONJ <> 0", "NOMESS")
  If REGSINBORRAR& > 0 Then
     TEXTO1$ = ("Se Han Detectado En la Tabla FORMULAS\Contiene " & REGSINBORRAR& & " Registros")
     OPX% = COMALTER%(TEXTO1$ + "\\Volver a Intentar\Cancelar")
     If OPX% < 1 Or OPX% > 1 Then Exit Sub
     If OPX% = 1 Then GoTo 10
  End If
  
  Dim FORMULASS As ADODB.Recordset
  Set FORMULASS = New ADODB.Recordset
  SQLX$ = "SELECT * FROM FORMULAS"
  FORMULASS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  '  Picture1.ToolTipText = "Conversión de Articulos"
  '
  Picture1.ToolTipText = "Conversión de Formulas"
  '
  'LECTURA
  FIN& = ULTREG&("MASTER")
  For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     CIXI% = U&
     CAIT% = 0
     RFF$ = RECFILT$("ESTRUNUE", 1, MKI$(CIXI%))
     If RFF$ <> "" Then
       REG& = CVS(Left$(RFF$, 4))
       If REG& > 0 Then
         CODICON$ = CODEXT$(CIXI%)
         X$ = REGLEIDO$("ESTRUNUE", REG&)
         DPS$ = Mid$(X$, 3, 120)
         REVIFORMU = CVI(Mid$(X$, 123, 2))
         If ULTREG&("ESTRUDOS") >= REG& Then
           DPS$ = DPS$ + REGLEIDO$("ESTRUDOS", REG&)
         End If
         '
         RFF$ = RECFILT$("ESTRULAR", 1, MKI$(CIXI%))
         If Len(RFF$) >= 4 Then
           RELA& = CVS(RFF$)
           If RELA& > 0 Then
             If RELA& <= ULTREG&("ESTRULAR") Then
               DPS$ = DPS$ + Mid$(REGLEIDO$("ESTRULAR", RELA&), 3)
             End If
           End If
         End If
         '
         While Len(DPS$) Mod 8 <> 0
           DPS$ = DPS$ + Chr$(0)
         Wend
         '
         For k% = 1 To Len(DPS$) Step 8
           DPSX$ = Mid$(DPS$, k%, 8)
           CIXXII% = CVI(Left$(DPSX$, 2))
           If CIXXII% > 0 Then
             If Abs(CVS(Mid$(DPSX$, 3))) >= 0.000005 Then
               CODXX$ = CODEXT$(CIXXII%)
               If TRIM$(CODXX$) <> "" Then
                 CAIT% = CAIT% + 1
                 CIJ%(CAIT%) = CIXXII%
                 USOI(CAIT%) = CVS(Mid$(DPSX$, 3, 4))
                 PDI%(CAIT%) = Asc(Mid$(DPSX$, 7, 1))
                 '
                 DEXX$ = DESCRIT$(CIJ%(CAIT%))
                 UNXX$ = Unimed$(CIJ%(CAIT%))
                 USON = USOI(CAIT%)
                 USOX$ = CSTRING$(MKS$(USON), 4, 12, 4, 2)
                 PDX$ = CSTRING$(Mid$(DPSX$, 7, 1), 1, 4, 0, 2)
                 XXX$ = Space$(64)
                 Call REPLA(XXX$, CODXX$, 1, 16)
                 Call REPLA(XXX$, DEXX$, 17, 28)
                 Call REPLA(XXX$, UNXX$, 47, 2)
                 Call REPLA(XXX$, USOX$, 49, 12)
                 Call REPLA(XXX$, PDX$, 61, 4)
                 EXPLOLIN(CAIT%) = XXX$
               End If
             End If
           End If
         Next k%
         '
         Call EXPLOGR(CODICON$, 1)
         '
       End If
     End If
  Next U&
  '
End Sub

Sub MIGRASECOPER()
  '
  Picture1.ToolTipText = "Del - Sec_Oper"
  '
36 Call ABRESECOP
   Dim RESECO$(256), DESOPE$(256)
   Call APERBASDAT("MASTER")
'   On Error Resume Next
10 Call BORRAREGISTROS("SECOPER", "CODICONJ <> 0 OR CODICONJ = 0", REGAF&, "NOMESS")

'   SQLX$ = "DROP TABLE SECOPER"
'   FLEXCONN.Execute SQLX$
'   On Error GoTo 0
   '
   Call ABRESECOP
   Call GENERBASE("ARTICULOS")
   '
   REGSINBORRAR& = CantRegistros("SECOPER", "CODICONJ <> 0", "NOMESS")
   If REGSINBORRAR& > 0 Then
     TEXTO1$ = ("Se Han Detectado En la Tabla SECOPER\Contiene " & REGSINBORRAR& & " Registros")
     OPX% = COMALTER%(TEXTO1$ + "\\Volver a Intentar\Cancelar")
     If OPX% < 1 Or OPX% > 1 Then Exit Sub
     If OPX% = 1 Then GoTo 10
   End If
   
   Dim SECOPSS As ADODB.Recordset
   Set SECOPSS = New ADODB.Recordset
   SQLX$ = "SELECT * FROM SECOPER"
   SECOPSS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  '
  Picture1.ToolTipText = "Conversión de Secuencia de Operaciones"
  '
  'LECTURA
  FIN& = ULTREG&("MASTER")
  For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     CI% = U&
     RFFX$ = RECFILT$("SECOPER", 1, MKI$(CI%))
     CASEC% = 0
     For I1& = 1 To Len(RFFX$) Step 4
        i& = CVS(Mid$(RFFX$, I1&, 4))
        X$ = REGLEIDO$("SECOPER", i&)
        '
        NOPY% = CVI(Mid$(X$, 3, 2))
        If NOPY% > ULTOP% Then ULTOP% = NOPY%
        NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
        DOPE$ = Mid$(X$, 5, 48)
        REGO& = i&
        REGOX$ = AJUSTD$(Str$(REGO&), 6)
        SECO$ = NOPEX$ + "     " + DOPE$ + " " + REGOX$
        CASEC% = CASEC% + 1
        RESECO$(CASEC%) = Left$(X$ + String$(128, 0), 128) + Left$(NUHOPROF$(CI%, NOPY%) + Space$(16), 16)
        DESOPE$(CASEC%) = ""
     Next I1&
     '
     RNC$ = RECFILT$("OBSEOPER", 1, MKI$(CI%))
     URN& = ULTREG&("OBSEOPER")
     '
     For U3& = 1 To Len(RNC$) Step 4
        RENOTA& = CVS(Mid$(RNC$, U3&, 4))
        If RENOTA& > 0 Then
          If RENOTA& <= URN& Then
            TTZZ$ = REGLEIDO$("OBSEOPER", RENOTA&)
            NOPY% = CVI(Mid$(TTZZ$, 3, 2))
            For KKI% = 1 To CASEC%
              If CVI(Mid$(RESECO$(KKI%), 3, 2)) = NOPY% Then
                DESOPE$(KKI%) = DESOPE$(KKI%) + Mid$(TTZZ$, 5)
              End If
            Next KKI%
          End If
        End If
     Next U3&
     '
     For LL1% = 1 To CASEC%
        LL% = LL1% + 1
        With SECOPSS
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        !CODICONJ = CI%
        !CODXCONJ = CODEXT$(CI%)
        !NumeOper = CVI(Mid$(RESECO$(LL% - 1), 3, 2))
        !AltOper = Asc(Mid$(RESECO$(LL% - 1), 88, 1))
        !DescOper = TRIM$(Mid$(RESECO$(LL% - 1), 5, 48))
        !statoper = 1
        !CodMaq0 = TRIM$(Mid$(RESECO$(LL% - 1), 53, 6))
        !CodMaq1 = TRIM$(Mid$(RESECO$(LL% - 1), 103, 6))
        !CodMaq2 = TRIM$(Mid$(RESECO$(LL% - 1), 109, 6))
        !CodMaq3 = TRIM$(Mid$(RESECO$(LL% - 1), 115, 6))
        !HorsPrep = CVS(Mid$(RESECO$(LL% - 1), 59, 4))
        !HorsFijo = CVS(Mid$(RESECO$(LL% - 1), 63, 4))
        !MinuStan = CVS(Mid$(RESECO$(LL% - 1), 67, 4))
        !HorsLimp = CVS(Mid$(RESECO$(LL% - 1), 79, 4))
        !HorsMHer = CVS(Mid$(RESECO$(LL% - 1), 95, 4))
        !PieCiclo = CVS(Mid$(RESECO$(LL% - 1), 99, 4))
        !PorDemor = CVS(Mid$(RESECO$(LL% - 1), 75, 4))
        If !MinuStan > 0.005 Then
          !Cadencia = 60 / !MinuStan
        End If
        !CantiOps = CVS(Mid$(RESECO$(LL% - 1), 83, 4))
        !CategOps = Asc(Mid$(RESECO$(LL% - 1), 87, 1))
        !CoefSolp = CVI(Mid$(RESECO$(LL% - 1), 89, 2))
        !LOTESTAN = CVS(Mid$(RESECO$(LL% - 1), 91, 4))
        !ProgCNum = TRIM$(Mid$(RESECO$(LL% - 1), 121, 8))
        !HoProFab = TRIM$(Mid$(RESECO$(LL% - 1), 129, 16))
        !ProTerOp = CVI(Mid$(RESECO$(LL% - 1), 149, 2))
        !PreUnid = CVD(Mid$(RESECO$(LL% - 1), 151, 8))
        !MonePrec = CVI(Mid$(RESECO$(LL% - 1), 159, 2))
        !FechPrec = CVDAT(CVI(Mid$(RESECO$(LL% - 1), 161, 2)))
        '
        !OBSEOPER = DESOPE(LL% - 1)
        !NUOPER_PREP = CANPRE
        !CATOPER_PREP = CATPRE%
        .Update
        End With
     Next LL1%
     '
  Next U&
  '
  SECOPSS.Close
  Set SECOPSS = Nothing
  '
End Sub

Sub MIGRAPEDIDOS()
  Picture1.ToolTipText = "Del - Pedidos"
  '
  Call APERBASDAT("PEDCLI")
'  SQLX$ = "DROP TABLE PedEnca "
'     On Error Resume Next
'      FLEXCONN.Execute (FILTSQL$(SQLX$))
'     On Error GoTo 0
'  SQLX$ = "DROP TABLE PedDeta "
'     On Error Resume Next
'      FLEXCONN.Execute (FILTSQL$(SQLX$))
'     On Error GoTo 0
  '
10 Call BORRAREGISTROS("PEDENCA", "NROPED <> 0 OR NROPED = 0", REGAF&, "NOMESS")
15 Call BORRAREGISTROS("PedDeta", "NROPED <> 0 OR NROPED = 0", REGAF&, "NOMESS")
  '
  Call APERBASDAT("PEDCLI")
  Call GENERBASE("VENTAS")
  
  REGSINBORRAR& = CantRegistros("PEDENCA", "NROPED <> 0 OR NROPED = 0", "NOMESS")
  If REGSINBORRAR& > 0 Then
     TEXTO1$ = ("Se Han Detectado En la Tabla PEDENCA\Contiene " & REGSINBORRAR& & " Registros")
     OPX% = COMALTER%(TEXTO1$ + "\\Volver a Intentar\Cancelar")
     If OPX% < 1 Or OPX% > 1 Then Exit Sub
     If OPX% = 1 Then GoTo 10
  End If
  '
  REGSINBORRAR& = CantRegistros("PedDeta", "NROPED <> 0 OR NROPED = 0", "NOMESS")
  If REGSINBORRAR& > 0 Then
     TEXTO1$ = ("Se Han Detectado En la Tabla PedDeta\Contiene " & REGSINBORRAR& & " Registros")
     OPX% = COMALTER%(TEXTO1$ + "\\Volver a Intentar\Cancelar")
     If OPX% < 1 Or OPX% > 1 Then Exit Sub
     If OPX% = 1 Then GoTo 15
  End If

  Picture1.ToolTipText = "Conversión de Pedidos"
  '
  strsql = "Select * from Pedenca "
  Dim PedEnca As ADODB.Recordset
  Set PedEnca = New ADODB.Recordset
  PedEnca.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
  '
  FIN& = ULTREG&("PEDENCA")
  Label1.Visible = True
  Label1.ToolTipText = "Hasta " & FIN&
  For U& = 1 To FIN&
    X$ = REGLEIDO$("PEDENCA", U&)
    Call TRACE(20, U&, FIN&)
    Label1 = U&
    'LECTURA
    NUPE& = CVS(Mid$(X$, 1, 4)) 'NUMERO DE PEDIDO
    REFE$ = Mid$(X$, 5, 48) 'REFERENCIA
    VENRE$ = ""
    FEPE% = CVI(Mid$(X$, 53, 2)) 'FECHA DE PEDIDO
    NUCLI% = CVI(Mid$(X$, 55, 2)) 'NUMERO DE CLIENTE
    NUOC1$ = Mid$(X$, 57, 16) 'NUMERO DE ORDEN DE COMPRA
    DOMIENT1$ = Mid$(X$, 73, 32) 'DOMICILIO DE ENTREGA
    CPLOCA$ = Mid$(X$, 105, 32) 'CP-LOCALIDAD
    TRANSP1$ = Mid$(X$, 137, 32) 'TRANPORTE
    OBS1$ = Mid$(X$, 169, 32) 'OBSERVACION UNO
    OBS2$ = Mid$(X$, 201, 32) 'OBSERVACION DOS
    OBS3$ = Mid$(X$, 233, 32) 'OBSERVACION TRES
    FESOL1% = CVI(Mid$(X$, 265, 2)) 'FECHA DE ENTREGA SOLICITADA
    STAT1% = Asc(Mid$(X$, 267, 1)) 'STATUS (1-ABIER,3-CUMP,9-ANUL)
    LIPRE1% = Asc(Mid$(X$, 268, 1)) 'LISTA DE PRECIOS
    COPA1% = Asc(Mid$(X$, 269, 1)) 'CONDICION DE PAGO
    VEND1% = Asc(Mid$(X$, 270, 1)) 'VENDEDOR
    DTO1$ = TRIM$(Mid$(X$, 275, 14)) 'DESCUENTO APLICADO
    CAMPVER$ = Mid$(X$, 289, 16)
    '
    VENRE$ = "V"
    RFG$ = RECFILT$("PEDDETA", 1, MKS$(NUPE&))
    If RFG$ <> "" Then
       REDEL& = CVS(Left$(RFG$, 4))
       XDET$ = REGLEIDO$("PEDDETA", REDEL&)
       If UCase$(Mid$(XDET$, 51, 1)) = "R" Then
          VENRE$ = "R"
          AAA = BBB
       End If
       PORCOVE = (CVI(Mid$(XDET$, 31, 2))) / 10
    End If
    
    ANOTA$ = ""
    RFF$ = RECFILT$("ANOTAPED", 1, MKS$(NUPE&))
    For UK1& = 1 To Len(RFF$) Step 4
       RENOTA& = CVS(Mid$(RFF$, UK1&, 4))
       XX$ = REGLEIDO$("ANOTAPED", RENOTA&)
       ANOTA$ = ANOTA$ + Mid$(XX$, 5, 92) + Chr$(10) + Chr$(13)
    Next UK1&
    '
    'GRABACION
    With PedEnca
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
        !NROPED = NUPE&
        !REFERENCIA = REFE$
        !TipoPed = "" 'POR CONVERSIÓN
        !FechEmis = CVDAT(FEPE%)
        !NroClie = NUCLI%
        !RaSoClie = rasocli$(NUCLI%)
        !NROOCOM = NUOC1$
        !DOMIENT = DOMIENT1$
        !LOCAENT = CPLOCA$
        !TranspEnt = TRANSP1$
        !Observa = OBS1$ + " " + OBS2$ + " " + OBS3$
        !FeSolEnt = CVDAT(FESOL1%)
        !Status = STAT1%
        !FinVigen = CVDAT(0)
        !ListPre = LIPRE1%
        !CondPag = COPA1%
        !Vended = VEND1%
        !PorComVend = PORCOVE
        !TDescApli = DTO1$
        !AnotaPed = ""
        !FApro_Nope = CVDAT(0)
        !VenRep = VENRE$
        ' !Usua_Apro = 0 'POR CONVERSIÓN
        !AnotaPed = ANOTA$
      .Update
    End With
    Call TRACE(20, U&, FIN&)
  Next U&
  PedEnca.Close
  Set PedEnca = Nothing
  '
  'PEDDETA
  strsql = "Select * from PedDeta"
  Dim PedDeta As ADODB.Recordset
  Set PedDeta = New ADODB.Recordset
  PedDeta.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
  '
  Picture1.ToolTipText = "Conversión de Pedidos"
  'LECTURA
  FIN& = ULTREG&("PEDDETA")
  Label1.ToolTipText = "Hasta " & FIN&
  NUPEA& = 0: NORIPE% = 0
  For U& = 1 To FIN&
    Call TRACE(20, U&, FIN&)
    Label1 = U&
    X$ = REGLEIDO$("PEDDETA", U&)
    'LECTURA
    NUPE& = CVS(Mid$(X$, 1, 4)) 'NUMERO DE PEDIDO
      If NUPE& < 1 Then GoTo 29
      'If NUPE& <> NUPEA& Then NORIPE% = 0
      NORIPE% = XVALO(Mid$(X$, 120, 3))
    FEPE% = CVI(Mid$(X$, 5, 2)) 'FECHA DEL PEDIDO
    NUCLI% = CVI(Mid$(X$, 7, 2)) 'NUMERO DE CLIENTE
    NUOC1$ = Mid$(X$, 9, 16) 'NUMERO DE ORDEN DE COMPRA
    FESOL1% = CVI(Mid$(X$, 25, 2)) 'FECHA ENTREGA SOLICITADA
    STAT1% = Asc(Mid$(X$, 27, 1)) 'STATUS
    LisPre% = Asc(Mid$(X$, 28, 1)) 'LISTA DE PRECIO
    COPA1% = Asc(Mid$(X$, 29, 1)) 'CONDICION DE PAGO
    VEND1% = Asc(Mid$(X$, 30, 1)) 'VENDEDOR
    PORCOM = CVI(Mid$(X$, 31, 2)) 'PORCENTAJE DE COMISION
    CI% = CVI(Mid$(X$, 33, 2)) 'CODIGO DEL MATERIAL
    COLO$ = Mid$(X$, 35, 16) 'COLOR
    TALL$ = ""
    VENRE$ = "V"
       If UCase$(Mid$(X$, 51, 1)) = "R" Then VENRE$ = "R"
    MONEMI1% = Asc(Mid$(X$, 52, 1))
    TIPOCA1 = CVS(Mid$(X$, 53, 4))
    LAR1 = CVS(Mid$(X$, 59, 4)) 'LARGO
    ANC1 = CVS(Mid$(X$, 63, 4)) 'ANCHO
    ESP1 = CVS(Mid$(X$, 67, 4)) 'ESPESOR
    PRELI1 = CVS(Mid$(X$, 71, 4)) 'PRECIO DE LISTA
    CASI1 = CVS(Mid$(X$, 75, 4)) 'CANTIDAD ASIGNADA
    CAPE1 = CVS(Mid$(X$, 79, 4)) 'CANTIDAD PEDIDA
    PREUNI1 = CVS(Mid$(X$, 83, 4)) 'PRECIO UNITARIO
    IMPTOT1 = CVD(Mid$(X$, 87, 8)) 'IMPORTE TOTAL
    CAEN1 = CVS(Mid$(X$, 95, 4)) 'CANTIDAD ENTREGADA
    CASA1 = CVS(Mid$(X$, 99, 4)) 'CANTIDAD SALDO
    FEUE% = CVI(Mid$(X$, 103, 2)) 'FECHA DE ULTIMA ENTREGA
    NURE1& = CVS(Mid$(X$, 105, 4)) 'NUMERO DE REMITO
    'RT-X-0002-00000030
    NURE3$ = ""
    If NURE1& > 0 Then
      NURE2$ = CStr(NURE1&)
      Do Until Len(NURE2$) = 8
       NURE2$ = "0" + NURE2$
      Loop
      NURE3$ = "RT-X-0000-" + NURE2$
    End If
    '
    ' SITUACION DE DESPACHO
    On Error GoTo 0
    STAT2% = 9
    RFG$ = RECFILT$("DESPAPED", 1, MKS$(NUPE&))
    If Len(RFG$) >= 4 Then
      For KKLU& = 1 To Len(RFG$) Step 4
        REDEL& = CVS(Mid$(RFG$, KKLU&, 4))
        XDET$ = REGLEIDO$("DESPAPED", REDEL&)
        If CVI(Mid$(XDET$, 33, 2)) = CI% Then
          STAT2% = Asc(Mid$(XDET$, 27, 1)) Mod 10
          Exit For
        End If
      Next KKLU&
    End If
    '
    'GRABACION
    With PedDeta
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
        !NroRemi = NURE3$ 'NUMERO PEDIDO FORMATEADO
        !TPedItem = ""
        !NROPED = NUPE&
        !TipoPed = "" 'POR CONVERSIÓN
        !FechEmis = CVDAT(FEPE%)
        !NroClie = NUCLI%
        !RaSoClie = rasocli$(NUCLI%)
        !NROOCOM = NUOC1$
        !FeSolEnt = CVDAT(FESOL1%)
        !FinVigen = CVDAT(0)
        !Status = STAT1%
        !StaDespa = STAT2%
        !ListPre = LisPre%
        !CondPag = COPA1%
        !Vended = VEND1%
        !PorComVend = PORCOM
           'NORIPE% = NORIPE% + 1
        !NuOrItem = NORIPE%
        !CodiInt = CI%
        !CodiExt = CODEXT$(CI%)
        !Descrip = DESCRIT0$(CI%)
        !SubCod = ""
        !TColor = COLO$
        !TTalle = TALL$
        !TLargo = LAR1
        !TAncho = ANC1
        !TEspesor = ESP1
        !VenRep = VENRE$
        !MoneEmis = MONEMI1%
        !TIPOCAM = TIPOCA1
        !PreList = PRELI1
        !CanAsig = CASI1
        !CanPed = CAPE1
        !PreUnid = PREUNI1
        !ImpoTot = IMPTOT1
        !CantEnt = CAEN1
        !CanSaldo = CASA1
        !FeUltEnt = CVDAT(FEUE%)
        !FApro_Nope = CVDAT(FEPE%)
        !ImpoAnti = 0
      .Update
    End With
    NUPEA& = NUPE&
    Call TRACE(20, U&, FIN&)
29 Next U&
  PedDeta.Close
  Set PedDeta = Nothing
  '
End Sub
   '
Sub MIGRACOMPRAS()
   '
59 Picture1.ToolTipText = "Del - Ordenes de Compra"
  'OCOMENCA
  Call APERBASDAT("COMPRAS")
'  SQLX$ = "DROP TABLE OCOMEnca "
'     On Error Resume Next
'      FLEXCONN.Execute (FILTSQL$(SQLX$))
'     On Error GoTo 0
'  SQLX$ = "DROP TABLE OCOMDeta "
'     On Error Resume Next
'      FLEXCONN.Execute (FILTSQL$(SQLX$))
'     On Error GoTo 0
10 Call BORRAREGISTROS("OCOMENCA", "NUME_OCOM <> 0 OR NUME_OCOM = 0", REGAF&, "NOMESS")
15 Call BORRAREGISTROS("OCOMDETA", "NUME_OCOM <> 0 OR NUME_OCOM = 0", REGAF&, "NOMESS")
  '
  Call APERBASDAT("COMPRAS")
  Call GENERBASE("COMPRAS")
  REGSINBORRAR& = CantRegistros("OCOMDETA", "NUME_OCOM <> 0 OR NUME_OCOM = 0", "NOMESS")
  If REGSINBORRAR& > 0 Then
     TEXTO1$ = ("Se Han Detectado En la Tabla OCOMENCA\Contiene " & REGSINBORRAR& & " Registros")
     OPX% = COMALTER%(TEXTO1$ + "\\Volver a Intentar\Cancelar")
     If OPX% < 1 Or OPX% > 1 Then Exit Sub
     If OPX% = 1 Then GoTo 10
  End If
  '
  REGSINBORRAR& = CantRegistros("OCOMDETA", "NUME_OCOM <> 0 OR NUME_OCOM = 0", "NOMESS")
  If REGSINBORRAR& > 0 Then
     TEXTO1$ = ("Se Han Detectado En la Tabla OCOMDETA\Contiene " & REGSINBORRAR& & " Registros")
     OPX% = COMALTER%(TEXTO1$ + "\\Volver a Intentar\Cancelar")
     If OPX% < 1 Or OPX% > 1 Then Exit Sub
     If OPX% = 1 Then GoTo 15
  End If
  
  Picture1.ToolTipText = "Conversión de Ordenes de Compra"
  '
  Dim OCOMENCA As ADODB.Recordset
  Set OCOMENCA = New ADODB.Recordset
  strsql = "SELECT * FROM  OCOMENCA"
  OCOMENCA.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
  
  FIN& = ULTREG&("OCOMENCA")
  For U& = 1 To FIN&
    X$ = REGLEIDO$("OCOMENCA", U&)
    'LECTURA
    NUCO& = CVS(Mid$(X$, 1, 4)) 'NUMERO DE O/C
    REFE$ = Mid$(X$, 5, 44) 'REFERENCIA
    FEMICO% = CVI(Mid$(X$, 49, 2)) 'FECHA DE EMISION
    PROVE% = CVI(Mid$(X$, 51, 2)) 'PROVEEDOR
    TIPSUM% = Asc(Mid$(X$, 53, 1)) 'TIPO DE SUMINISTRO
    ESTOC% = Asc(Mid$(X$, 54, 1)) 'ESTADO DE LA O/C
    FEGEN% = CVI(Mid$(X$, 55, 2)) 'FECHA DE GENERACIÓN
    HORGEN% = CVI(Mid$(X$, 57, 2)) 'HORA DE GENERACIÓN
    USUGEN% = CVI(Mid$(X$, 59, 2)) 'USUARIO GENERACIÓN
    FEMOD% = CVI(Mid$(X$, 61, 2)) 'FECHA DE MODIFICACIÓN
    HOMOD% = CVI(Mid$(X$, 63, 2)) 'HORA REAL DE MODIFICACIÓN
    USUMOD% = CVI(Mid$(X$, 65, 2)) 'USUARIO MODIFICACION
    FEANU% = CVI(Mid$(X$, 67, 2)) 'FECHA REAL ANULACIÓN
    HOANU% = CVI(Mid$(X$, 69, 2)) 'HORA REAL ANULACIÓN
    USUANU% = CVI(Mid$(X$, 71, 2)) 'USUARIO ANULACIÓN
    MOTANUL$ = Mid$(X$, 73, 10) 'MOTIVO DE LA ANULACIÓN
    MOEVA% = Asc(Mid$(X$, 83, 1)) 'MONEDA DE VALORIZACIÓN
    MOEMI% = Asc(Mid$(X$, 84, 1)) 'MONEDA DE EMISIÓN
    TICA = CVS(Mid$(X$, 85, 4)) 'TIPO DE CAMBIO
    IBRU = CVD(Mid$(X$, 89, 8)) 'IMPORTE BRUTO SIN IVA
    PODE$ = Mid$(X$, 97, 16) 'PORCENTAJE DE DESCUENTO
    IMDE = CVD(Mid$(X$, 113, 8)) 'IMPORTE DEL DESCUENTO
    IMIV = CVD(Mid$(X$, 121, 8)) 'IMPORTE DEL IVA
    IMOT = CVD(Mid$(X$, 129, 8)) 'IMPORTE OTROS
    IMOT2 = CVD(Mid$(X$, 137, 8)) 'IMPORTE OTROS DOS
    TOTOC = CVD(Mid$(X$, 145, 8)) 'TOTAL ORDEN DE COMPRA
    COPA% = CVI(Mid$(X$, 153, 2)) 'CONDICIÓN DE PAGO
    ATESR$ = Mid$(X$, 163, 30) 'ATE. SEÑOR:
    OBS$ = Mid$(X$, 193, 64) 'OBSERVACIONES CERO
    OBS1$ = Mid$(X$, 257, 64) 'OBSERVACIONES UNO
    OBS2$ = Mid$(X$, 321, 64) 'OBSERVACIONES DOS
    OBS3$ = Mid$(X$, 385, 64) 'OBSERVACIONES TRES
    OBS4$ = Mid$(X$, 449, 64) 'OBSERVACIONES CUATRO
    'GRABACIÓN
    With OCOMENCA
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
        !NUME_OCOM = NUCO&
        !Refe_Ocom = REFE$
        !FEMI_OCOM = CVDAT(FEMICO%)
        !Npro_Ocom = PROVE%
        !RASO_PROV = Left$(rasocli$(PROVE% * (-1)), 48)
        !TSum_Ocom = TIPSUM%
        !Stat_Ocom = ESTOC%
        !FGen_Ocom = CVDAT(FEGEN%)
        !Usua_Gen = USUGEN%
        !FApro_Ocom = CVDAT(FEMICO%)
        !Usua_Apro = 0 'POR CONVERSIÓN
        On Error Resume Next
        !FMod_Ocom = CVDAT(FEMOD%)
        On Error GoTo 0
        !Usua_Mod = USUMOD%
        !FAnu_Ocom = CVDAT(FEANU%)
        !Usua_Anul = USUANU%
        !Moti_Anul = MOTANUL$
        !Mone_Cos = MOEVA%
        !Mone_Emis = MOEMI%
        !Tcam_Ocom = TICA
        !ToBru_Ocom = IBRU
        !Pdes_Ocom = PODE$
        !Ides_Ocom = IMDE
        !Iiva_Ocom = IMIV
        !Iotr1_Ocom = IMOT
        !Iotr2_Ocom = IMOT2
        !Total_Ocom = TOTOC
        !Cpag_Ocom = COPA%
        !Embala_Ocom = 0 'Por Conversión
        !AtSr_Ocom = ATESR$
        !Obse_Ocom = OBS$ + " " + OBS1$ + " " + OBS2$ + " " + OBS3$ + " " + OBS4$
        !Lentrega = "" 'Por Conversión
      .Update
    End With
    Call TRACE(20, U&, FIN&)
  Next U&
  OCOMENCA.Close
  Set OCOMENCA = Nothing
  'OCOMDETA
  '
  Dim OCOMDETA As ADODB.Recordset
  Set OCOMDETA = New ADODB.Recordset
  strsql = "SELECT * FROM  OCOMDETA"
  OCOMDETA.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
  
  Picture1.ToolTipText = "Conversión de Ordenes de Compra"
  '
  FIN& = ULTREG&("OCOMDETA")
  NUANT& = 0
  For U& = 1 To FIN&
    X$ = REGLEIDO$("OCOMDETA", U&)
    'LECTURA
    NUCO& = CVS(Mid$(X$, 1, 4)) 'NUMERO DE O/C
    FEMICO% = CVI(Mid$(X$, 5, 2)) 'FECHA DE EMISION
    PROVE% = CVI(Mid$(X$, 7, 2)) 'PROVEEDOR
    CI% = CVI(Mid$(X$, 9, 2)) 'CODIGO DEL MATERIAL
    UNME$ = Mid$(X$, 11, 4) 'UNIDAD DE MEDIDA SEGUN MAESTRO
    COLO$ = Mid$(X$, 15, 16) 'COLOR
    MEDI$ = Mid$(X$, 31, 24) 'MEDIDAS
    CAPE = CVS(Mid$(X$, 55, 4)) 'CANTIDAD PEDIDA
    UNOC$ = Mid$(X$, 59, 4) 'UNIDAD DE COMPRA
    COEUN = CVS(Mid$(X$, 63, 4)) 'COEFICIENTE DE UNIDAD
    PUNI = CVD(Mid$(X$, 67, 8)) 'PRECIO UNITARIO POR UNIDAD
    DEIT = CVS(Mid$(X$, 75, 4)) 'PORCENTAJE DE DESCUENTO DEL ITEM
    PUNINET = CVD(Mid$(X$, 79, 8)) 'PRECIO UNITARIO NETO
    IMPONET = CVD(Mid$(X$, 87, 8)) 'IMPORTE NETO DEL ITEM
    FENSOL% = CVI(Mid$(X$, 95, 2)) 'FECHA DE ENTREGA SOLICITADA
    COMP$ = Mid$(X$, 97, 2) 'CONSIGNA MATERIA PRIMA
    CARE = CVS(Mid$(X$, 99, 4)) 'CANTIDAD RECIBIDAD
    COPE% = CVI(Mid$(X$, 103, 2)) 'CODIGO DE OPERACIÓN
    DOCUORI$ = Mid$(X$, 105, 12) 'DOCUMENTO DE ORIGEN
    NUORSE& = CVS(Mid$(X$, 117, 4)) 'NUMERO DE ORDEN DE SERVICIO
    MONCO% = Asc(Mid$(X$, 121, 1)) 'MONEDA DE COSTEO
    MONE% = Asc(Mid$(X$, 122, 1)) 'MONEDA DE EMISION
    TIPSUM% = Asc(Mid$(X$, 123, 1)) 'TIPO DE SUMINISTRO
    STAOC% = Asc(Mid$(X$, 124, 1)) 'STATUS DE LA O/C
    'GRABACIÓN
    With OCOMDETA
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
        !NUME_OCOM = NUCO&
          If NUCO& <> NUANT& Then
             NORI% = 0
             NUANT& = NUCO&
          End If
          NORI% = NORI% + 1
        !NuOrd_Item = NORI%
        !FEMI_OCOM = CVDAT(FEMICO%)
        !Npro_Ocom = PROVE%
        !RASO_PROV = rasocli$(PROVE% * (-1))
        !Cod_Inte = CI%
        !COD_EXTE = CODEXT$(CI%)
        !Descritem = DESCRIT0$(CI%)
        !Delibre = "" 'POR CONVERSIÓN
        !IContab = 0 'POR CONVERSIÓN
        !Sect_Ocom = "" 'POR CONVERSIÓN
        !Uni_Mas = UNME$
        !Colo_Ocom = COLO$
        !Medi_Ocom = Left$(MEDI$, 20)
        !Valo_Unid = "" 'POR CONVERSIÓN
        !Unids_Ocom = 0 'POR CONVERSIÓN
        !Cant_Ocom = CAPE
        !Uni_Com = UNOC$
        !Coef_Unids = COEUN
        !PRE_UNIT = PUNI
        !Desc_Item = DEIT
        !prun_neto = PUNINET
        !INet_Item = IMPONET
        !Fentre_Sol = CVDAT(FENSOL%)
        !Consimat = COMP$
        !Cant_Rec = CARE
        !Cod_Oper = COPE%
        !Doc_Orig = DOCUORI$
        !N_OrServ = NUORSE&
        !Mone_Cos = MONCO%
        !Mone_Emis = MONE%
        !TSum_Ocom = TIPSUM%
        !Stat_Ocom = STAOC%
        !Nro_Otrab = "" 'POR CONVERSIÓN
        !Nro_Grup = "" 'POR CONVERSIÓN
      .Update
    End With
    Call TRACE(20, U&, FIN&)
  Next U&
  '
  OCOMDETA.Close
  Set OCOMDETA = Nothing
End Sub

Sub MIGRASTOCKS()
  '
  Call ABRECONEXION
  Call GENERBASE("STOCKS")
  Picture1.ToolTipText = "Del - Movimientos de Stock"
10 Call BORRAREGISTROS("MOVISTO", "COD_INTE <> 0 OR COD_INTE = 0", REGAF&, "NOMESS")
'  On Error Resume Next
'   FLEXCONN.Execute (FILTSQL$("DROP TABLE MOVISTO"))
'  On Error GoTo 0
  Call GENERBASE("STOCKS")
'  Call APERBASDAT("STOCKS")
  '
  REGSINBORRAR& = CantRegistros("MOVISTO", "COD_INTE <> 0 OR COD_INTE = 0", "NOMESS")
  If REGSINBORRAR& > 0 Then
     TEXTO1$ = ("Se Han Detectado En la Tabla MOVISTO\Contiene " & REGSINBORRAR& & " Registros")
     OPX% = COMALTER%(TEXTO1$ + "\\Volver a Intentar\Cancelar")
     If OPX% < 1 Or OPX% > 1 Then Exit Sub
     If OPX% = 1 Then GoTo 10
  End If
  '
  Picture1.ToolTipText = "Conversión de Movimientos de Stocks"
  '
  FIN& = ULTREG&("MOVISTO")
  DOCANT$ = ""
  PREFIREM$ = TRIM$(Control$("", "PREFIREM")): If PREFIREM$ = "" Then PREFIREM$ = "0001"
  '
  SQLX$ = "SELECT * FROM Movisto "
  Dim MOVISTMP As ADODB.Recordset
  Set MOVISTMP = New ADODB.Recordset
  MOVISTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  '
  For U& = 1 To FIN&
    X$ = REGLEIDO$("MOVISTO", U&)
    Call TRACE(20, U&, FIN&)
    Label1 = U&
    'LECTURA
      FEC% = CVI(Mid$(X$, 1, 2)) 'FECHA DE MOVIMIENTO
      CI% = CVI(Mid$(X$, 3, 2)) 'CODIGO DE PIEZA
      LOT$ = Mid$(X$, 5, 16) 'LOTE / PARTIDA
      COMO$ = Mid$(X$, 21, 2) 'CODIGO DE MOVIMIENTO
      DOC1$ = Mid$(X$, 23, 10) 'DOCUMENTO PRIMARIO ASOCIADO
      DOC1LAR$ = DOC1$
      If COMO$ = "RT" Then
         NUMEDO$ = TRIM$(Mid$(DOC1$, 4))
         While Len(NUMEDO$) < 8: NUMEDO$ = "0" + NUMEDO$: Wend
         DOC1LAR$ = Left$(DOC1$, 2) + "-X-" + PREFIREM$ + "-" + NUMEDO$
      End If
      DOC2$ = Mid$(X$, 33, 10) 'DOCUMENTO SECUNDARIO
      DOC2LAR$ = DOC2$
      If Left$(DOC2$, 2) = "FC" Then
         NUMEDO$ = TRIM$(Mid$(DOC2$, 4))
         While Len(NUMEDO$) < 8: NUMEDO$ = "0" + NUMEDO$: Wend
         DOC2LAR$ = Left$(DOC2$, 2) + "-0001-" + NUMEDO$ + "-A"
      End If
      REFE$ = Mid$(X$, 43, 30) 'REFERENCIA
      VAUN# = CVD(Mid$(X$, 73, 8)) 'VALOR UNITARIO
      MONE1$ = Mid$(X$, 81, 2) 'MONEDA
      DEP% = Asc(Mid$(X$, 83, 1)) 'DEPOSITO
      NUCLI% = CVI(Mid$(X$, 84, 2)) 'NUEMRO DE CLIENTE
      NUPRO% = CVI(Mid$(X$, 86, 2)) 'NUMERO DE PROVEEDOR
      CAEN = CVD(Mid$(X$, 88, 8)) 'CANTIDAD ENTRADA
      CASA = CVD(Mid$(X$, 96, 8)) 'CANTIDAD SALIDA
      CASALD = CVD(Mid$(X$, 104, 8)) 'CANTIDAD SALDO
      Fill$ = Mid$(X$, 112, 10) 'FILLER NO IDENTIFICADO
      FEMO% = CVI(Mid$(X$, 122, 2)) 'FECHA DE MOVIMIENTO
      USU% = Asc(Mid$(X$, 126, 1)) 'USUARIO
    'GRABACIÓN
    With MOVISTMP
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
        !fechMov = CVDAT(FEC%)
        !Cod_Inte = CI%
        !COD_EXTE = " " 'CODEXT$(CI%) - eliminado para ganar velocidad
        !Descrip = " " 'DESCRIT0(CI%) - eliminado para ganar velocidad
        !IDENLOTE = LOT$
        !CodiMovi = COMO$
        !DoPriCor = DOC1$
        !DoPriLar = DOC1LAR$
           If DOC1$ <> DOCANT$ Then
             NORI% = 0
             DOCANT$ = DOC1$
           End If
        !DoSecCor = DOC2$
        !DoSecLar = DOC2LAR$
        !NumRem_Cli = "" 'POR CONVERSIÓN
           NORI% = NORI% + 1
        !NuOrItem = NORI%
        !NUPEDCLI = CVS(Left$(Fill$, 4))      'POR CONVERSIÓN
          NOPXX% = CVI(Mid$(Fill$, 5, 2))
          If NOPXX% = 8224 Then NOPXX% = 0
        !NUMEORDPED = NOPXX%
        !ReferCor = REFE$
        !ValoUnit = VAUN
        !MONEVAL = XVALO(MONE1$)
        !DepoMovi = DEP%
        !Nume_Clie = NUCLI%
        !Nume_Prov = NUPRO%
        !CantIngre = CAEN
        !CANTSALID = CASA
        !CantSaldo = CASALD
        !FeReMovi = CVDAT(FEMO%)
      .Update
    End With
    '
    Call TRACE(20, U&, FIN&)
  Next U&
  MOVISTMP.Close
  Set MOVISTMP = Nothing
  '
  Call MIGRASTOCKSVIE
  '
End Sub

Sub MIGRASTOCKSVIE()
  '
  Call ABRECONEXION
  Call GENERBASE("STOCKS")
  Picture1.ToolTipText = "Del - Movimientos de Stock"
  Call BORRAREGISTROS("MOVIVIE", "COD_INTE <> 0 OR COD_INTE = 0", REGAF&, "NOMESS")
'  On Error Resume Next
'   FLEXCONN.Execute (FILTSQL$("DROP TABLE MOVISTO"))
'  On Error GoTo 0
  Call GENERBASE("STOCKS")
'  Call APERBASDAT("STOCKS")
  '
  Picture1.ToolTipText = "Conversión de Movimientos de Stocks"
  '
  FIN& = ULTREG&("MOVISVIE")
  DOCANT$ = ""
  PREFIREM$ = TRIM$(Control$("", "PREFIREM")): If PREFIREM$ = "" Then PREFIREM$ = "0001"
  '
  SQLX$ = "SELECT * FROM Movivie "
  Dim MOVISTMP As ADODB.Recordset
  Set MOVISTMP = New ADODB.Recordset
  MOVISTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  '
  For U& = 1 To FIN&
    X$ = REGLEIDO$("MOVISVIE", U&)
    Call TRACE(20, U&, FIN&)
    Label1 = U&
    'LECTURA
      FEC% = CVI(Mid$(X$, 1, 2)) 'FECHA DE MOVIMIENTO
      CI% = CVI(Mid$(X$, 3, 2)) 'CODIGO DE PIEZA
      LOT$ = Mid$(X$, 5, 16) 'LOTE / PARTIDA
      COMO$ = Mid$(X$, 21, 2) 'CODIGO DE MOVIMIENTO
      DOC1$ = Mid$(X$, 23, 10) 'DOCUMENTO PRIMARIO ASOCIADO
      DOC1LAR$ = DOC1$
      If COMO$ = "RT" Then
         NUMEDO$ = TRIM$(Mid$(DOC1$, 4))
         While Len(NUMEDO$) < 8: NUMEDO$ = "0" + NUMEDO$: Wend
         DOC1LAR$ = Left$(DOC1$, 2) + "-X-" + PREFIREM$ + "-" + NUMEDO$
      End If
      DOC2$ = Mid$(X$, 33, 10) 'DOCUMENTO SECUNDARIO
      DOC2LAR$ = DOC2$
      If Left$(DOC2$, 2) = "FC" Then
         NUMEDO$ = TRIM$(Mid$(DOC2$, 4))
         While Len(NUMEDO$) < 8: NUMEDO$ = "0" + NUMEDO$: Wend
         DOC2LAR$ = Left$(DOC2$, 2) + "-0001-" + NUMEDO$ + "-A"
      End If
      REFE$ = Mid$(X$, 43, 30) 'REFERENCIA
      VAUN# = CVD(Mid$(X$, 73, 8)) 'VALOR UNITARIO
      MONE1$ = Mid$(X$, 81, 2) 'MONEDA
      DEP% = Asc(Mid$(X$, 83, 1)) 'DEPOSITO
      NUCLI% = CVI(Mid$(X$, 84, 2)) 'NUEMRO DE CLIENTE
      NUPRO% = CVI(Mid$(X$, 86, 2)) 'NUMERO DE PROVEEDOR
      CAEN = CVD(Mid$(X$, 88, 8)) 'CANTIDAD ENTRADA
      CASA = CVD(Mid$(X$, 96, 8)) 'CANTIDAD SALIDA
      CASALD = CVD(Mid$(X$, 104, 8)) 'CANTIDAD SALDO
      Fill$ = Mid$(X$, 112, 10) 'FILLER NO IDENTIFICADO
      FEMO% = CVI(Mid$(X$, 122, 2)) 'FECHA DE MOVIMIENTO
      USU% = Asc(Mid$(X$, 126, 1)) 'USUARIO
    'GRABACIÓN
    With MOVISTMP
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
        !fechMov = CVDAT(FEC%)
        !Cod_Inte = CI%
        !COD_EXTE = " " 'CODEXT$(CI%) - eliminado para ganar velocidad
        !Descrip = " " 'DESCRIT0(CI%) - eliminado para ganar velocidad
        !IDENLOTE = LOT$
        !CodiMovi = COMO$
        !DoPriCor = DOC1$
        !DoPriLar = DOC1LAR$
           If DOC1$ <> DOCANT$ Then
             NORI% = 0
             DOCANT$ = DOC1$
           End If
        !DoSecCor = DOC2$
        !DoSecLar = DOC2LAR$
        !NumRem_Cli = "" 'POR CONVERSIÓN
           NORI% = NORI% + 1
        !NuOrItem = NORI%
        !NUPEDCLI = CVS(Left$(Fill$, 4))      'POR CONVERSIÓN
          NOPXX% = CVI(Mid$(Fill$, 5, 2))
          If NOPXX% = 8224 Then NOPXX% = 0
        !NUMEORDPED = NOPXX%
        !ReferCor = REFE$
        !ValoUnit = VAUN
        !MONEVAL = XVALO(MONE1$)
        !DepoMovi = DEP%
        !Nume_Clie = NUCLI%
        !Nume_Prov = NUPRO%
        !CantIngre = CAEN
        !CANTSALID = CASA
        !CantSaldo = CASALD
        !FeReMovi = CVDAT(FEMO%)
      .Update
    End With
    '
    Call TRACE(20, U&, FIN&)
  Next U&
  MOVISTMP.Close
  Set MOVISTMP = Nothing
  '
End Sub


Sub MIGRABOLREC()
  Picture1.ToolTipText = "Del - Boletas de Rececpción"
  Call APERBASDAT("STOCKS")
  '
10 Call BORRAREGISTROS("BOLRECEP", "NUMEBORE <> 0 OR NUMEBORE = 0", REGAF&, "NOMESS") '  On Error Resume Next
'   FLEXCONN.Execute (FILTSQL$("DROP TABLE BOLRECEP"))
'  On Error GoTo 0
  Call GENERBASE("STOCKS")
'  Call APERBASDAT("STOCKS")
  '
  REGSINBORRAR& = CantRegistros("BOLRECEP", "NUMEBORE <> 0 OR NUMEBORE = 0", "NOMESS")
  If REGSINBORRAR& > 0 Then
     TEXTO1$ = ("Se Han Detectado En la Tabla BOLRECEP\Contiene " & REGSINBORRAR& & " Registros")
     OPX% = COMALTER%(TEXTO1$ + "\\Volver a Intentar\Cancelar")
     If OPX% < 1 Or OPX% > 1 Then Exit Sub
     If OPX% = 1 Then GoTo 10
  End If
  '
  SQLX$ = "SELECT * FROM BOLRECEP "
  Dim BOLRECTMP As ADODB.Recordset
  Set BOLRECTMP = New ADODB.Recordset
  BOLRECTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  '
  Picture1.ToolTipText = "Conversión de Boletas de Recepción"
  '
  FIN& = ULTREG&("BOLRECEP")
  For U& = 1 To FIN&
    X$ = REGLEIDO$("BOLRECEP", U&)
    'LECTURA
    NUBORE& = CVS(Mid$(X$, 1, 4)) 'NUMERO DE BOLETA DE RECEPCIÓN
    DES$ = Mid$(X$, 5, 64) 'DESCRIPCION
    LOT$ = Mid$(X$, 69, 12) 'IDENTIFICADOR DE LOTE
    FERE% = CVI(Mid$(X$, 81, 2)) 'FECHA DE RECEPCION
    CI% = CVI(Mid$(X$, 83, 2)) 'CODIGO DEL MATERIAL
    UM$ = Mid$(X$, 85, 4) 'UNIDAD DE MEDIDA
    CANOM = CVD(Mid$(X$, 89, 8)) 'CANITDAD NOMINAL
    DEPEN% = Asc(Mid$(X$, 97, 1)) 'DEPOSITO DE ENTRADA
    PROVE% = CVI(Mid$(X$, 105, 2)) 'NUMERO DE PROVEEDOR
    LOTPROV$ = Mid$(X$, 107, 12) 'LOTE DEL PROVEEDOR
    REPROV$ = Mid$(X$, 119, 16) 'NUMERO DE REMITO DEL PROVEEDOR
    FAPROV$ = Mid$(X$, 135, 16) 'NUMERO DE FACTURA DEL PROVEEDOR
    NUOC& = CVS(Mid$(X$, 151, 4)) 'NUMERO DE ORDEN DE COMPRA
    MOVAUN$ = Mid$(X$, 155, 4) 'MONEDA VALOR UNITARIO
    VAUN = CVS(Mid$(X$, 159, 4)) 'VALOR UNITARIO
    CAPRO1 = CVD(Mid$(X$, 163, 8)) 'CANTIDAD APROBADA
    CARE = CVD(Mid$(X$, 171, 8)) 'CANTIDAD RECHAZADA
    CAFA = CVD(Mid$(X$, 179, 8)) 'CANTIDAD FALTANTE
    FELI% = CVI(Mid$(X$, 187, 2)) 'FECHA DE LIBERACIÓN
    NUORSE& = CVS(Mid$(X$, 189, 4)) 'NUMERO DE ORDEN DE SERVICIO
    'GRABACIÓN
    Call APERBASDAT("STOCKS")
    With BOLRECTMP
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
        !NUMEBORE = NUBORE&
        !Descrip = DES$
        !IDENLOTE = LOT$
        !FechRecep = CVDAT(FERE%)
        !Cod_Inte = CI%
        !COD_EXTE = " " ' CODEXT$(CI%)
        !DescripItem = " " 'DESCRIT0(CI%)
        !Unidad = UM$
        !CANTNOM = CANOM
        !DepoEnt = DEPEN%
        !NumProv = PROVE%
        !RASO_PROV = rasocli$(PROVE% * (-1))
        !LOTECOL_PROV = LOTPROV$
        !NumRem_Prov = REPROV$
        !NumFac_Prov = FAPROV$
        !NUMOC = NUOC&
        !IContab = 0 'POR CONVERSIÓN
        !MoneVaUni = MOVAUN$
        !Valo_Unit = VAUN
        !CANTOAPRO = CAPRO1
        !CanToRech = CARE
        !CanToFalta = CAFA
        !FechLibe = CVDAT(FELI%)
        !NumeOrdSer = NUORSE&
        !Observa = ""
        !Status = 1 'POR CONVERSIÓN
        !UsuApro = "" 'POR CONVERSIÓN
        !FeHoApro = CVDAT(FERE%) 'POR CONVERSIÓN
      .Update
    End With
    Call TRACE(20, U&, FIN&)
  Next U&
  '
  BOLRECTMP.Close
  Set BOLRECTMP = Nothing
  '
End Sub

Sub MIGRACONSIGNA()
  Picture1.ToolTipText = "Del - Material en Consignación"
  Call APERBASDAT("STOCKS")
  '
10 Call BORRAREGISTROS("MACONSIG", "NPRO_OCOM <> 0 OR NPRO_OCOM = 0", REGAF&, "NOMESS") '  On Error Resume Next
'   FLEXCONN.Execute (FILTSQL$("DROP TABLE MACONSIG"))
'  On Error GoTo 0
  Call GENERBASE("STOCKS")
  Call APERBASDAT("STOCKS")
  
  REGSINBORRAR& = CantRegistros("MACONSIG", "NPRO_OCOM <> 0 OR NPRO_OCOM = 0", "NOMESS")
  If REGSINBORRAR& > 0 Then
     TEXTO1$ = ("Se Han Detectado En la Tabla MACONSIG\Contiene " & REGSINBORRAR& & " Registros")
     OPX% = COMALTER%(TEXTO1$ + "\\Volver a Intentar\Cancelar")
     If OPX% < 1 Or OPX% > 1 Then Exit Sub
     If OPX% = 1 Then GoTo 10
  End If
  '
  SQLX$ = "SELECT * FROM MACONSIG "
  Dim MACONTMP As ADODB.Recordset
  Set MACONTMP = New ADODB.Recordset
  MACONTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  '
  Picture1.ToolTipText = "Conversión Material en Consignación"
  '
  FIN& = ULTREG&("MACONSIG")
  For U& = 1 To FIN&
    X$ = REGLEIDO$("MACONSIG", U&)
    'LECTURA
    FERE% = CVI(Mid$(X$, 1, 2)) 'FECHA DE RESERVA
    PROVE% = CVI(Mid$(X$, 3, 2)) 'NUMERO DE PROVEEDOR
    SECC% = CVS(Mid$(X$, 5, 4)) 'SECCION
    NUOF$ = Mid$(X$, 9, 12) 'NUMERO DE ORDEN DE FABRICACIÓN
    NUOC1$ = Mid$(X$, 21, 12) 'NUMERO DE ORDEN DE COMPRA
    CI% = CVI(Mid$(X$, 33, 2)) 'CODIGO INTERNO DEL MATERIAL
    LOT$ = Mid$(X$, 35, 16) 'LOTE - PARTIDA
    UM$ = Mid$(X$, 51, 4) 'UNIDAD DE MEDIDA
    CASI = CVD(Mid$(X$, 55, 8)) 'CANTIDAD ASIGNADA
    CAEN = CVD(Mid$(X$, 63, 8)) 'CANTIDAD ENTREGADA
    FENT% = CVI(Mid$(X$, 71, 2)) 'FECHA DE ENTREGA
    NURE$ = Mid$(X$, 73, 16) 'NUMERO DE REMITO
    VAUN = CVD(Mid$(X$, 93, 8)) 'VALOR UNITARIO
    MON% = CVI(Mid$(X$, 101, 2)) 'MONEDA
    VATO = CVD(Mid$(X$, 103, 8)) 'VALOR TOTAL
    CADE = CVD(Mid$(X$, 111, 8)) 'CANTIDAD RENDIDA (DEVUELTA)
    SAPE = CVD(Mid$(X$, 119, 8)) 'SALDO PENDIENTE
    OBS$ = Mid$(X$, 129, 112) 'OBSERVACIONES
    'GRABACIÓN
    Call APERBASDAT("STOCKS")
    With MACONTMP
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
        !FEMI_OCOM = CVDAT(FERE%)
        !Npro_Ocom = PROVE%
        !RASO_PROV = rasocli$(PROVE% * (-1))
        !Secc_Ejec = SECC%
        !Deno_Secc = "" 'POR CONVERSIÓN
        !Nro_Orfa = NUOF$
        !nro_ocom = NUOC1$
        !Cod_Inte = CI%
        !COD_EXTE = " " ' CODEXT$(CI%)
        !Descritem = " " ' DESCRIT0(CI%)
        !Lote_Ocom = LOT$
        !Unidad = UM$
        !Cant_Asig = CASI
        !Cant_Entreg = CAEN
        !Fecha_Entreg = CVDAT(FENT%)
        !NRemi_Consig = NURE$
        !Valo_Unit = VAUN
        !Moneda = MON%
        !Valo_Tot = VATO
        !Cant_Rendid = CADE
        !Saldo_Pend = SAPE
        !Observa = OBS$
      .Update
    End With
    Call TRACE(20, U&, FIN&)
  Next U&
  '
  MACONTMP.Close
  Set MACONTMP = Nothing
  '
End Sub

Sub MIGRAORFAB()
  Picture1.ToolTipText = "Del - Ordenes de Fabricación"
  Call APERBASDAT("ORFAB")
  '
10 Call BORRAREGISTROS("ORDEFABR", "NUMEORFA <> 0 OR NUMEORFA = 0", REGAF&, "NOMESS")
'  On Error Resume Next
'   FLEXCONN.Execute (FILTSQL$("DROP TABLE ORDEFABR"))
'  On Error GoTo 0
  Call GENERBASE("OFABRIC")
  Call APERBASDAT("ORFAB")
  '
  REGSINBORRAR& = CantRegistros("ORDEFABR", "NUMEORFA <> 0 OR NUMEORFA = 0", "NOMESS")
  If REGSINBORRAR& > 0 Then
     TEXTO1$ = ("Se Han Detectado En la Tabla ORDEFABR\Contiene " & REGSINBORRAR& & " Registros")
     OPX% = COMALTER%(TEXTO1$ + "\\Volver a Intentar\Cancelar")
     If OPX% < 1 Or OPX% > 1 Then Exit Sub
     If OPX% = 1 Then GoTo 10
  End If
  
  SQLX$ = "SELECT * FROM ORDEFABR "
  Dim ORFATMP As ADODB.Recordset
  Set ORFATMP = New ADODB.Recordset
  ORFATMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  '
  Picture1.ToolTipText = "Conversión Ordenes de Fabricación"
  '
  FIN& = ULTREG&("ORDEFABR")
  For U& = 1 To FIN&
    X$ = REGLEIDO$("ORDEFABR", U&)
    'LECTURA
    NUOF1& = CVS(Mid$(X$, 1, 4)) 'NUMERO DE ORDEN DE FABRICACION
    REFEOF$ = Mid$(X$, 5, 64) 'REFERENCIA
    LOTOF$ = Mid$(X$, 69, 12) 'IDENTIFICADOR DE LOTE
    FEGEOF% = CVI(Mid$(X$, 81, 2)) 'FECHA DE GENERACIÓN
    COPIOF% = CVI(Mid$(X$, 83, 2)) 'CODIGO DE PIEZA
    UMOF$ = Mid$(X$, 85, 4) 'UNIDAD DE MEDIDA
    CAPROYE = CVD(Mid$(X$, 89, 8)) 'CANTIDAD PROYECTADA
    DEOF$ = Mid$(X$, 97, 8) 'CLIENTE O DEPOSITO DESTINO
    FESOF% = CVI(Mid$(X$, 105, 2)) 'FECHA SOLICITADA DE ENTREGA
    PRIOF% = Asc(Mid$(X$, 107, 1)) 'PRIORIDAD DE LA OF
    STATOF% = Asc(Mid$(X$, 108, 1)) 'STATUS DE LA OF
    FELAOF% = CVI(Mid$(X$, 109, 2)) 'FECHA DE LANZAMIENTO DE LA OF
    COSEOF% = CVS(Mid$(X$, 113, 4)) 'CODIGO DEL SECTOR PRODUCTIVO
    FECIOF% = CVI(Mid$(X$, 117, 2)) 'FECHA DE CIERRE DE LA OF
    CAPROF = CVD(Mid$(X$, 119, 8)) 'CANTIDAD APROBADA OF
    CAREOF = CVD(Mid$(X$, 127, 8)) 'CANTIDAD RECHAZADA OF
    IDOROF$ = Mid$(X$, 135, 8) 'IDENTIFICADOR ORDEN RETRABAJO
    FENOF% = CVI(Mid$(X$, 147, 2)) 'FECHA PROYECTADA DE ENTREGA
    FETEOF% = CVI(Mid$(X$, 149, 2)) 'FECHA TEMPRANA DE ENTREGA
    FERECOF% = CVI(Mid$(X$, 151, 2)) 'FECHA DE RECALCULO
    OBSOF$ = Mid$(X$, 155, 4) 'OBSERVACIONES
    'GRABACIÓN
    Call APERBASDAT("ORFAB")
    With ORFATMP
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
        !NumeOrFa = NUOF1&
        !Referen = REFEOF$
        !IDSUBOR = LOTOF$
        !FECHGEN = CVDAT(FEGEOF%)
        !Cod_Inte = COPIOF%
        !COD_EXTE = " " ' CODEXT$(COPIOF%)
        !Descrip = " " ' DESCRIT0(COPIOF%)
        !Unidad = UMOF$
        !CanProy = CAPROYE
        !Destino = DEOF$
        !FechEnSol = CVDAT(FESOF%)
        !Priorid = PRIOF%
        !StatuOrf = STATOF%
        !FechLan = CVDAT(FELAOF%)
        !SecciPro = COSEOF%
        !FechCierr = CVDAT(FECIOF%)
        !CanApro = CAPROF
        !CanRech = CAREOF
        !IdenRetra = IDOROF$
        !FechProyEntre = CVDAT(FENOF%)
        !FechTempEntre = CVDAT(FETEOF%)
        !FechRecal = CVDAT(FERECOF%)
        !Observa = OBSOF$
        !AnotaOrf = " "    'Anotaciones de ordenes de fabricación
      .Update
    End With
    Call TRACE(20, U&, FIN&)
  Next U&
  '
  ORFATMP.Close
  Set ORFATMP = Nothing
End Sub

Sub MIGRAANOTAORFAB()
   '
'   Call APERBASDAT("ORFAB")
   SQLX$ = "SELECT * FROM ORDEFABR "
   Set ORFATMP = New ADODB.Recordset
   ORFATMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   '
   With ORFATMP
      On Error Resume Next
      .MoveFirst
      If Err Then
         On Error GoTo 0
         Exit Sub
      End If
      '
      Do While Not .EOF
         NUOF1& = XVALO(!NumeOrFa)
         ANOTA$ = ""
         RFF$ = RECFILT$("ANOTAORF", 1, MKS$(NUOF1&))
         For UY& = 1 To Len(RFF$) Step 4
           RENOTA& = CVS(Mid$(RFF$, UY&, 4))
           XX$ = REGLEIDO$("ANOTAORF", RENOTA&)
           XXYY$ = Mid$(XX$, 5, 92)
           ANOTA$ = ANOTA$ + XXYY$ + Chr$(10) + Chr$(13)
         Next UY&
         If ANOTA$ <> "" Then
            MsgBox ANOTA$
         End If
         !AnotaOrf = ANOTA$ 'Anotaciones de ordenes de fabricación
         .Update
         .MoveNext
      Loop
   End With
   '
End Sub

Sub MIGRARESERVA()
  Picture1.ToolTipText = "Del - Reservas"
'  Call APERBASDAT("RESERVA")
  '
10 Call BORRAREGISTROS("RESERVA", "COD_INTE <> 0 OR COD_INTE = 0", REGAF&, "NOMESS")
'  On Error Resume Next
'   FLEXCONN.Execute (FILTSQL$("DROP TABLE RESERVA"))
'  On Error GoTo 0
  Call GENERBASE("STOCKS")
  Call APERBASDAT("RESERVA")
  '
  REGSINBORRAR& = CantRegistros("RESERVA", "COD_INTE <> 0 OR COD_INTE = 0", "NOMESS")
  If REGSINBORRAR& > 0 Then
     TEXTO1$ = ("Se Han Detectado En la Tabla RESERVA\Contiene " & REGSINBORRAR& & " Registros")
     OPX% = COMALTER%(TEXTO1$ + "\\Volver a Intentar\Cancelar")
     If OPX% < 1 Or OPX% > 1 Then Exit Sub
     If OPX% = 1 Then GoTo 10
  End If
  '
  SQLX$ = "SELECT * FROM RESERVA "
  Dim RESERTMP As ADODB.Recordset
  Set RESERTMP = New ADODB.Recordset
  RESERTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  '
  Picture1.ToolTipText = "Conversión Reservas de Materiales"
  '
  FIN& = ULTREG&("RESERVA")
  For U& = 1 To FIN&
    Call TRACE(20, U&, FIN&)
    X$ = REGLEIDO$("RESERVA", U&)
    'LECTURA
    CI% = CVI(Mid$(X$, 1, 2)) 'CODIGO INTERNO DEL MATERIAL
    NOF$ = Mid$(X$, 3, 12) 'NUMERO DE ORDEN DE FABRICACION
    FERE% = CVI(Mid$(X$, 17, 2)) 'FECHA DE RESERVA
    CARE = CVD(Mid$(X$, 19, 8)) 'CANTIDAD RESERVADA
    CACO = CVD(Mid$(X$, 27, 8)) 'CANTIDAD CONSUMIDA
    SALEN = CVD(Mid$(X$, 35, 8)) 'SALDO A ENTREGAR
    LOT$ = Mid$(X$, 43, 16) 'LOTE O PARTIDA
    DEP% = Asc(Mid$(X$, 59, 1)) 'DEPOSITO
    USOUNI = CVS(Mid$(X$, 75, 4)) 'USO UNITARIO
    'GRABACIÓN
    With RESERTMP
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
        !Cod_Inte = CI%
        !COD_EXTE = " " ' CODEXT$(CI%)
        !Descrip = " " ' DESCRIT0(CI%)
        !IDSUBOR = NOF$
        !FechRes = CVDAT(FERE%)
        !CANTRES = CARE
        !CANTCONS = CACO
        !Saldo_Entre = SALEN
        !IDENLOTE = LOT$
        !DEPOSITO = DEP%
        !USOUNIT = USOUNI
      .Update
    End With
  Next U&
  '
  RESERTMP.Close
  Set RESERTMP = Nothing
  '
End Sub

Sub MIGRAPLANCUEN()

5  Call ABREPLANCUEN
  Picture1.ToolTipText = "Del - Plan de Cuentas"
'  On Error Resume Next
''  SQLX$ = "DROP TABLE PLANCUEN "
''      FLEXCONN.Execute (FILTSQL$(SQLX$))
''  On Error GoTo 0
  'Set Plancuen = Contable.OpenRecordset("PLANCUEN", dbOpenDynaset)
  '
  Picture1.ToolTipText = "Plan de Cuentas Contables"
  '
  Call ABREPLANCUEN
  Dim PLANCUEN1 As ADODB.Recordset
  Set PLANCUEN1 = New ADODB.Recordset
  SQLX$ = "SELECT * FROM PLANCUEN"
  PLANCUEN1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     
     Dim YASTA%(2048)
     FIN& = ULTREG&("ECUECOT")
     If FIN& > 0 Then
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         IVP$ = REGLEIDO$("ECUECOT", U&)
         '
         CUE% = CVI(Mid$(IVP$, 41, 2))
         If YASTA%(CUE%) = 1 Then GoTo 19
         YASTA%(CUE%) = 1
         NIVEI% = 2
         '
         If CUE% > 0 Then
           ZZ$ = Space$(256)
           COCUE$ = TRIM$(Left$(IVP$, 12))
           DECUE$ = TRIM$(Mid$(IVP$, 13, 28))
           '
           If COCUE$ <> "" Then
             Call REPLA(ZZ$, NVX$, 1, 6)
             Call REPLA(ZZ$, COCUE$, 9, 12)
             Call REPLA(ZZ$, DECUE$, 23, 44)
             Call REPLA(ZZ$, DECUE$, 129, 128)
             '
             RECUEN& = CUE%
             XX$ = REGLEIDO$("CUENTAS", RECUEN&)
             COCUE$ = TRIM$(Left$(XX$, 12))
             TICUEX$ = UCase$(Mid$(XX$, 13, 1))
             If TICUEX$ = "H" Then
                 TICUEX$ = "Haber"
               ElseIf TICUEX$ = "I" Then
                 TICUEX$ = "Ingresos"
               ElseIf TICUEX$ = "G" Then
                 TICUEX$ = "Gastos"
               Else
                 TICUEX$ = "Debe"
             End If
             YY$ = REGLEIDO$("DENOCUE", RECUEN&)
             '
             IMPUBLE$ = ""
             STATUX$ = ""
             '
             With PLANCUEN1
               .AddNew   ' YA CONVERTIDO
               On Error Resume Next
               !CodiEmpr = CodiEmp%
               On Error GoTo 0
               !CodExCue = COCUE$
               !DENOCUE = DECUE$
               !CODINCUE = CUE%
               !STATCUE = 1
               !NIVAGRU = NIVEI%
               !TipoCue = TICUEX$
               !NUORDIT = 0
               !POSIMPU = 1
               !Observa = ""
               !SALINI = 0
               !Sudebe = 0
               !Suhabe = 0
               !Result = 0
               !Salfin = 0
               !IndiAjus = 0
               !MoneExtr = 0
               .Update
             End With
             '
           End If
         End If
19     Next U&
     End If
     '
     ' PONE EN ORDEN SEGÚN PLAN GENERAL DE CUENTAS
     Dim PLANCUEN2 As ADODB.Recordset
     FIN& = ULTREG&("PLAGRAL")
     For U& = 1 To FIN&
       XX$ = REGLEIDO$("PLAGRAL", U&)
       CII% = CVI(XX$)
       NIVEI% = CVI(Mid$(XX$, 3, 2))
       '
       Set PLANCUEN2 = New ADODB.Recordset
       SQLX$ = "SELECT * FROM PLANCUEN WHERE CODINCUE = " & CII%
       PLANCUEN2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       With PLANCUEN2
          If Not (.EOF And .BOF) Then
'          .FindFirst "CODINCUE = " & CII%
'          If .NoMatch = False Then
'             .Edit
             !NIVAGRU = NIVEI%
             !NUORDIT = U&
'             MsgBox !CodExCue & "   " & !DENOCUE & U& & "   " & NIVEI%
             .Update
          End If
       End With
       PLANCUEN2.Close
       Set PLANCUEN2 = Nothing
       '
     Next U&
     '
     ' LOS QUE NO ESTÁN EN EL PLAN DE CUENTAS VAN AL FINAL Y SUSPENDIDAS
     U& = FIN&
     With PLANCUEN1
        .MoveFirst
        Do While Not .EOF
           If !NUORDIT = 0 Then
             U& = U& + 1
'             .Edit
             !NUORDIT = U&
             !STATCUE = 1
             .Update
           End If
        .MoveNext
        Loop
     End With
     '
     PLANCUEN1.Close
     'PLANCUEN2.Close
     Set PLANCUEN1 = Nothing
     'Set PLANCUEN2 = Nothing
End Sub

Sub BORRACUECORR()
   '
   Call ABRECONEXION
'   On Error Resume Next
'    FLEXCONN.Execute (FILTSQL$("DROP TABLE CAJABANC"))
'   On Error GoTo 0
   '
   Call ABRECABAN
   SQLX$ = "DELETE FROM CAJABANC"
   On Error Resume Next
   FLEXCONN.Execute SQLX$
   On Error GoTo 0
   '
'   On Error Resume Next
'    FLEXCONN.Execute (FILTSQL$("DROP TABLE SADEUPEN"))
'   On Error GoTo 0
   '
   SQLX$ = "DELETE FROM SADEUPEN"
   On Error Resume Next
   FLEXCONN.Execute SQLX$
   On Error GoTo 0
   '
'   On Error Resume Next
'    FLEXCONN.Execute (FILTSQL$("DROP TABLE SACREPEN"))
'   On Error GoTo 0
   '
   SQLX$ = "DELETE FROM SACREPEN"
   On Error Resume Next
   FLEXCONN.Execute SQLX$
   On Error GoTo 0
   '
'   On Error Resume Next
'    FLEXCONN.Execute (FILTSQL$("DROP TABLE APLICOBRA"))
'   On Error GoTo 0
   '
   SQLX$ = "DELETE FROM APLICOBRA"
   On Error Resume Next
   FLEXCONN.Execute SQLX$
   On Error GoTo 0
   '
'   On Error Resume Next
'    FLEXCONN.Execute (FILTSQL$("DROP TABLE APLIPAGO"))
'   On Error GoTo 0
   '
   SQLX$ = "DELETE FROM APLIPAGO"
   On Error Resume Next
   FLEXCONN.Execute SQLX$
   On Error GoTo 0
   '
   Call GENERBASE("CAJABANC")
   Screen.MousePointer = 1
   '
End Sub

Sub MIGRAPRESU()
  '\\\OT-07-0414-FL-05/11/07///
  Call ABREPRESVEN
  '
'  On Error Resume Next
'   FLEXCONN.Execute (FILTSQL$("DROP TABLE PRESUVEN"))
'  On Error GoTo 0
10 Call BORRAREGISTROS("PRESUVEN", "NROPRES <> 0 OR NROPRES = 0", REGAF&, "NOMESS")
  Call GENERBASE("VENTAS")
  Call ABREPRESVEN
  '
  REGSINBORRAR& = CantRegistros("PRESUVEN", "NROPRES <> 0 OR NROPRES = 0", "NOMESS")
  If REGSINBORRAR& > 0 Then
     TEXTO1$ = ("Se Han Detectado En la Tabla PRESUVEN\Contiene " & REGSINBORRAR& & " Registros")
     OPX% = COMALTER%(TEXTO1$ + "\\Volver a Intentar\Cancelar")
     If OPX% < 1 Or OPX% > 1 Then Exit Sub
     If OPX% = 1 Then GoTo 10
  End If
  
  SQLX$ = "SELECT * FROM PRESUVEN "
  Dim PRESUTMP As ADODB.Recordset
  Set PRESUTMP = New ADODB.Recordset
  PRESUTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  '
  'LECTURA
  ULTPRESU& = 0
  FIN& = ULTREG&("PRESUPU")
  For U& = 1 To FIN&
    Call TRACE(20, U&, FIN&)
    X3$ = REGLEIDO$("PRESUPU", U&)
    '
    'LECTURA
    NUMPRE& = CVS(Mid$(X3$, 1, 4))           'NUMERO DE PRESUPUESTO
    FECHE% = CVI(Mid$(X3$, 5, 2))            'FECHA DE EMISION
    NUMECLI1% = CVI(Mid$(X3$, 7, 2))         'NUMERO DE CLIENTE
    REFER$ = Mid$(X3$, 9, 16)                'REFERENCIA
    ATENSR$ = Mid$(X3$, 25, 40)              'ATENCION SR.
    LISTPREC% = Asc(Mid$(X3$, 65, 1))        'LISTA DE PRECIOS
    CONPAGO% = Asc(Mid$(X3$, 66, 1))         'CONDICION DE PAGO
    VENDEDO% = Asc(Mid$(X3$, 67, 1))         'VENDEDOR
    CODART% = CVI(Mid$(X3$, 69, 2))          'CODIGO DEL MATERIAL
    PRECLIST# = CVS(Mid$(X3$, 71, 4))        'PRECIO DE LISTA
    DESCUEN$ = Mid$(X3$, 75, 14)             'DESCUENTO
    DESCUE# = CVS(Mid$(X3$, 89, 4))          '% DESCUENTO
    IMPDESC# = CVS(Mid$(X3$, 93, 4))         'IMPORTE DESCUENTO
    PRECUNIT# = CVS(Mid$(X3$, 97, 4))        'PRECIO UNITARIO
    CANTIDAD# = CVS(Mid$(X3$, 101, 4))       'CANTIDAD
    IMPTOTNET# = CVD(Mid$(X3$, 105, 8))      'IMPORTE TOTAL NETO
    FECVENT% = CVI(Mid$(X3$, 113, 2))        'FECHA DE VENTA
    NUORDVEN& = CVS(Mid$(X3$, 115, 4))       'NUMERO DE RODEN DE VENTA
    NUMFACT& = CVS(Mid$(X3$, 119, 4))        'NUMERO DE FACTURA
    '
    'GRABACION
    On Error GoTo 0
    With PRESUTMP
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
        !NroPres = NUMPRE&
        !Cod_Rubro = " "
        !NumeCRubro = 0
        !Cod_Anio = " "
        FECHE1 = CVDAT(FECHE%)
        !FechEmis = FECHE1
        !NroClie = NUMECLI1%
        RASOCLIEN$ = ECOARCH$("DEUDOR1", MKI(NUMECLI1%))
        !RaSoClie = RASOCLIEN$
        !Referen = REFER$
        !AteSen = ATENSR$
        !LisPre = LISTPREC%
        !CONPAG = CONPAGO%
        !Vended = VENDEDO%
        !MonedEmi = 1
        If ULTPRESU& <> NUMPRE& Then
          NUORIT% = 1: ULTPRESU& = NUMPRE&
        Else
          NUORIT% = NUORIT% + 1
        End If
        !NUORDIT = NUORIT%
        !CodiInt = CODART%
        !CodiExt = CODEXT$(CODART%)
           Descrip$ = ECOARCH$("MASTER", MKI$(CODART%))
        !Descrip = Left$(Descrip$, 64)
        !Especif = Descrip$
        !PreList = PRECLIST#
        !TDESCUEN = DESCUEN$
        !PorcDesc = DESCUE#
        !ImpoDesc = IMPDESC#
        !PREUNIT = PRECUNIT#
        !Cantit = CANTIDAD#
        !ImToNe = IMPTOTNET#
        FECVENT1 = CVDAT(FECVENT%)
        !FechVent = FECVENT1
        !NroPedi = NUORDVEN&
        NroFac_Cor = NUMFACT&
      .Update
    End With
  Next U&
  '
  PRESUTMP.Close
  Set PRESUTMP = Nothing
  '\\\OT-07-0414-FIN///
End Sub

Sub MIGRAPRECIOS()
    '
    '\\\OT-07-0437-14-11-07///
   Call APERBASDAT("PEDCLI")
'   With ListaPrec
'     On Error Resume Next
'10   .MoveFirst
'     If Err < 1 Then
'        .Delete
'        GoTo 10
'     End If
'     On Error GoTo 0
'   End With

    'BORRA LA TABLA LISTAPRE COMPLETA
10  SQLX$ = " DELETE FROM ListaPre "
     FLEXCONN.Execute (FILTSQL$(SQLX$))
    
    'BORRA LA TABLA PREVENTA COMPLETA
15  SQLX$ = " DELETE FROM PREVENTA "
     FLEXCONN.Execute (FILTSQL$(SQLX$))

   '
    REGSINBORRAR& = CantRegistros("ListaPre", "Cod_Lista >= 0 OR Cod_Lista <= 0", "NOMESS")
    If REGSINBORRAR& > 0 Then
       TEXTO1$ = ("Se Han Detectado En la Tabla ListaPre\Contiene " & REGSINBORRAR& & " Registros")
       OPX% = COMALTER%(TEXTO1$ + "\\Volver a Intentar\Cancelar")
       If OPX% < 1 Or OPX% > 1 Then Exit Sub
       If OPX% = 1 Then GoTo 10
    End If
   '
   REGSINBORRAR& = CantRegistros("PREVENTA", "Cod_Lista >= 0 OR Cod_Lista <= 0", "NOMESS")
   If REGSINBORRAR& > 0 Then
       TEXTO1$ = ("Se Han Detectado En la Tabla PREVENTA\Contiene " & REGSINBORRAR& & " Registros")
       OPX% = COMALTER%(TEXTO1$ + "\\Volver a Intentar\Cancelar")
       If OPX% < 1 Or OPX% > 1 Then Exit Sub
       If OPX% = 1 Then GoTo 15
   End If
   '
   'Lee y graba cada registro del archivo LISTAS.DAT a la tabla LISTAPRE
   FIN& = ULTREG&("LISTAS")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X3$ = REGLEIDO$("LISTAS", U&)
      '
      'LECTURA
      CODLIST% = Asc(Mid$(X3$, 1, 1))          'CODIGO DE LISTA
      DENOMIN$ = Mid$(X3$, 2, 46)              'DENOMINACION DE LA LISTA
      Formato% = Asc(Mid$(X3$, 48, 1))         'FORMATO
      DECIMA% = Asc(Mid$(X3$, 49, 1))          'DECIMALES
      Moneda% = Asc(Mid$(X3$, 50, 1))          'MONEDA
      ABREVI$ = Mid$(X3$, 51, 12)              'ABREVIATURA
      VIGENC% = CVI(Mid$(X3$, 63, 2))          'VIGENCIA
      '
      'GRABACION
      SQLX$ = "SELECT * FROM LISTAPRE "
      Dim LISSSTAPRE As ADODB.Recordset
      Set LISSSTAPRE = New ADODB.Recordset
      LISSSTAPRE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText


      With LISSSTAPRE
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        !Cod_Lista = CODLIST%
        !DENOM_LISTA = DENOMIN$
        !VIGENCIA = CVDAT(VIGENC%)
        !Formato = Formato%
        !DECIMALES = DECIMA%
        !Moneda = Moneda%
        !DESC_ABREV = ABREVI$
        .Update
      End With
      LISSSTAPRE.Close
      Set LISSSTAPRE = Nothing
      '
      ARCHIPRE$ = TRIM$(Str$(CODLIST%))
      While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
      ARCHIPRE$ = "LIPRE" + ARCHIPRE$
      FINA& = ULTREG&(ARCHIPRE$)
      '
      SQLX$ = "SELECT * FROM PREVENTA "
      SQLX$ = SQLX$ + "WHERE Cod_Lista = " & CODLIST% & ""
      Dim PREVVVENTA As ADODB.Recordset
      Set PREVVVENTA = New ADODB.Recordset
      PREVVVENTA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
      ' Lee y graba cada dato del archivo LIPREnnn en la tabla PREVENTA
      For J& = 1 To FINA&
         'LECTURA
         X3$ = REGLEIDO$(ARCHIPRE$, J&)
         CODPIE% = CVI(Mid$(X3$, 1, 2))         'CODIGO DE PIEZA
         DESCRI$ = Mid$(X3$, 3, 42)             'DESCRIPCION
         CARACONT$ = Mid$(X3$, 45, 1)           'CARACTER DE CONTROL
         NUMORIGI$ = Mid$(X3$, 46, 16)          'NUMERO ORIGINAL
         PRECUNIT# = CVD(Mid$(X3$, 62, 8))      'PRECIO UNITARIO
         CANTLIMI% = CVI(Mid$(X3$, 70, 2))      'CANTIDAD LIMITE
         PORDESCU% = Asc(Mid$(X3$, 72, 1))      'PORCENTAJE DE DESCUENTO
         '
         'GRABACION
         DENOPI$ = DESCRIT$(CODPIE%)
         With PREVVVENTA
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
            !Cod_Lista = CODLIST%
            !Cod_Pieza = CODPIE%
            !Num_Orden = J&
            !Desc_Pieza = DESCRI$
            !Denom_Pieza = Left$(DENOPI$, 64)
            !CARAC_CONTROL = CARACONT$
            !Num_Original = NUMORIGI$
            !Precio_Unitario = PRECUNIT#
            !Cant_Limite = CANTLIMI%
            !Porc_Limite = PORDESCU%
            .Update
         End With
      Next J&
      PREVVVENTA.Close
      Set PREVVVENTA = Nothing
      '
   Next U&
   '\\\OT-07-0437-FIN///
End Sub

Sub MIGRACONTA()
   '
   AVARIA$ = LOADFILE$(LUDAT$ + "VARIAB.DAT") ', 0, 32)
   V$ = Mid$(AVARIA$, 33, 32)
   COMEJ1% = CVI(Mid$(V$, 29, 2))
   FINEJ1% = CVI(Mid$(V$, 31, 2))
   DES% = COMEJ1%: Has% = FINEJ1%
   '
   If Len(EXTE$) < 3 Then
      DES% = FECHANUM("01/01/" + EXTE$)
      Has% = FECHANUM("31/12/" + EXTE$)
   End If
   '
   MODOXX2% = 1 ' INCLUYE LOS ASIENTOS DE CAJA Y BANCOS
   MODOXX3% = 1 ' RENUMERA LOS ASIENTOS
   '
   Screen.MousePointer = 11
   Sudebe# = 0: Suhabe# = 0
   '
   SORTLIST.Clear
   FIN& = MAXDIARIO%
   For kk% = 1 To MAXDIARIO%
      UU& = kk%
      Call TRACE(20, UU&, FIN&)
      Call LEEPASDIA(kk%, 2)
      '
      If Asc(Left$(E8$, 1)) = 1 Then GoTo 29 ' SALTEA ASIENTOS RESUMEN DE VENTAS
      If Asc(Left$(E8$, 1)) = 2 Then GoTo 29 ' SALTEA ASIENTOS RESUMEN DE COMPRAS
      If Asc(Left$(E8$, 1)) = 3 Then GoTo 29 ' SALTEA ASIENTOS RESUMEN DE CAJA Y BANCOS
      '
      FEAS% = CVI(E15$)
      RESOR$ = Left$(TRIM$(Str$(10000 + FEAS%)) + Space$(10), 10)
      RESOR$ = RESOR$ + "DI" + TRIM$(Str$(100000 + kk%))
      SORTLIST.AddItem RESOR$
29 Next kk%
   '
' ACA VER COMO SE INCLUYEN LOS DE CAJA Y BANCOS
   If MODOXX2% = 1 Then
     FIN& = MAXDIA&
     For KKL& = 1 To MAXDIA&
       UU& = KKL&
       Call TRACE(20, UU&, FIN&)
       Call LEEPASCAJ(KKL&, 2)
       If Asc(Left$(D8$, 1)) = 1 Then GoTo 39 ' SALTEA ASIENTOS RESUMEN DE VENTAS
       If Asc(Left$(D8$, 1)) = 2 Then GoTo 39 ' SALTEA ASIENTOS RESUMEN DE COMPRAS
       FEAS% = CVI(D15$)
       RESOR$ = Left$(TRIM$(Str$(10000 + FEAS%)) + Space$(10), 10)
       RESOR$ = RESOR$ + "CA" + TRIM$(Str$(100000 + KKL&))
       SORTLIST.AddItem RESOR$
39   Next KKL&
   End If
   '
   ' CARGAR MOVIMIENTOS DE DIARIO
   FIN& = SORTLIST.ListCount    ' MAXDIARIO%
   YATRA% = 0: JJ& = 0: ASIANT = 0: ASIANTX = 0
   '
   NUASDEF& = 0
   N1% = FILEOPEN%(LUDAT$ + "ASIENTOS.TXT", 2, 256)

   For kk% = 1 To FIN&     'MAXDIARIO%
      UU& = kk%
      Call TRACE(20, UU&, FIN&)
      RESOR$ = SORTLIST.List(kk% - 1)
      REGII& = Val(Mid$(RESOR$, 13)) - 100000
      GRUXI$ = TRIM$(Mid$(RESOR$, 11, 2))
      SGX1% = 1
      FEASI% = Val(TRIM$(Left$(RESOR$, 10))) - 10000
      MODOLE% = 1
      If FEASI% >= DES% And FEASI% <= Has% Then MODOLE% = 2
      If GRUXI$ = "DI" Then
         REGIJ% = REGII&
         Call LEEPASDIA(REGIJ%, 2)
         FEAS% = CVI(E15$)
         IPAS# = CVD(E16$)
         CUE% = CVI(E11$)
         NUAS = CVS(E12$)
         REFEL$ = "- " + TRIM$(DETAPAS$)
         REFE$ = REPLACAR$(REFEL$, "@", "A")
         If TRIM$(REFE$) = "-" Then
           REFE$ = "- " + TRIM$(DENOCUE$(CUE%))
         End If
         DENOMI$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
        Else
         SGX1% = (-1)
         Call LEEPASCAJ(REGII&, 2)
         FEAS% = CVI(D15$)
         IPAS# = CVD(D16$)
         CUE% = CVI(D11$)
         NUAS = CVS(D12$)
         REFEL$ = "- " + TRIM$(DETAPAS$)
         REFE$ = REPLACAR$(REFEL$, "@", "A")
         If TRIM$(REFE$) = "-" Then
           REFE$ = "- " + TRIM$(DENOCUE$(CUE%))
         End If
         DENOMI$ = TRIM$(CODCUE$(CUE%)) + " - " + TRIM$(DENOCUE$(CUE%))
      End If
      '
      If NUAS * SGX1% <> ASIANT Then
        NUASDEF& = NUASDEF& + 1
        ASIANT = NUAS * SGX1%
      End If
      '
      If FEAS% < DES% Then
         If IPAS# > 0 Then
              Sudebe# = Sudebe# + IPAS#
            Else
              Suhabe# = Suhabe# - IPAS#
         End If
         GoTo 90
      End If
      '
      FECHASTA$ = FECHATEX$(FEAS%)
      If FEAS% <= Has% Then
        If NUAS * SGX1% <> ASIANTX Then
           ASIANTX = NUAS * SGX1%
           Z$ = Space$(256)
           Call REPLA(Z$, FECHASTA$, 16, 8)
           Call REPLA(Z$, FORMATNUM$(NUASDEF&, "F6.0"), 1, 6)
           JJ& = JJ& + 1
         End If
         '
         Call REPLA(Z$, FECHASTA$, 16, 8)
         Call REPLA(Z$, FORMATNUM$(NUASDEF&, "F6.0"), 1, 6)
         Call REPLA(Z$, REFASIEN$, 32, 48)
         Call REPLA(Z$, TRIM$(CODCUE$(CUE%)), 83, 16)
         Call REPLA(Z$, REFE$, 100, 46)
         Call REPLA(Z$, Space$(32), 146, 32)
         If IPAS# > 0 Then
              Call REPLA(Z$, FORMATNUM$(IPAS#, "F13.2"), 146, 13)
            Else
              Call REPLA(Z$, FORMATNUM$((-1) * IPAS#, "F13.2"), 162, 13)
         End If
         Print #N1%, Z$
      End If
90 Next kk%
   '
   Close #N1%
   Screen.MousePointer = 1
   '
99 End Sub

Sub MIGRACONTB()
   '
   Dim LISTA$(256)
   Call ABREPLANCUEN
   '
   Dim EncaSiGra As ADODB.Recordset
   Set EncaSiGra = New ADODB.Recordset
   strsql = "Select * from EncaBeAsien where NuIDAsien = '" & NUASI$ & "'"
   EncaSiGra.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   Dim DetasiGra As ADODB.Recordset
   Set DetasiGra = New ADODB.Recordset
   strsql = "Select * from DetaAsien where NuIDAsien = '" & NUASI$ & "'"
   DetasiGra.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   N1% = FILEOPEN%(LUDAT$ + "ASIENTOS.TXT", 1, 256)
   JJ& = 0
   FIN& = LOF(N1%) / 256
   Label1.ToolTipText = "Hasta " & FIN&
   Do While Not EOF(N1%)
     U& = U& + 1
     Call TRACE(20, U&, FIN&)
     Label1 = U&
     Line Input #N1%, TTXX$
     NUMERO = XVALO(TRIM$(Mid$(TTXX$, 1, 12)))
     If NUMERO <> NUMEROANT Then
        If JJ& > 0 Then GoSub GRABATODO
        JJ& = 0
        NUMEROANT = NUMERO
        FECHA$ = TRIM$(Mid$(TTXX$, 16, 8))
        FECHAU% = FECHANUM(FECHA$)
        NUASI$ = NUPROA$(FECHA$)
        REFEX8$ = TRIM$(Mid$(TTXX$, 32, 48))
        SUMADEBE# = 0
     End If
     JJ& = JJ& + 1
     LISTA(JJ&) = TTXX$
     IDEBE# = XVALO(TRIM$(Mid$(LISTA$(JJ&), 146, 13)))
     SUMADEBE# = SUMADEBE# + ROUND(IDEBE#)
   Loop
   GoSub GRABATODO
   '
   EncaSiGra.Close
   Set EncaSiGra = Nothing
   DetasiGra.Close
   Set DetasiGra = Nothing
   Close #N1%
   Exit Sub
   '
GRABATODO:
   If Left$(REFEX8$, 16) = "DEPOSITO CHEQUES" Then
      GoSub GRABAPORCHEQUES
      GoTo 49
   End If
   '
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
      If Left$(DEPASE$, 1) = "-" Then DEPASE$ = TRIM$(Mid$(DEPASE$, 2))
      IDEBE# = XVALO(TRIM$(Mid$(LISTA$(K2%), 146, 13)))
         SUMADEBE# = SUMADEBE# + ROUND(IDEBE#)
      IHABE# = XVALO(TRIM$(Mid$(LISTA$(K2%), 162, 13)))
      MACON$ = " "
      CENCOS1$ = " "
      '
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
49 Return

GRABAPORCHEQUES:
   '
   ' DETERMINA EL NÚMERO DE COMPROBANTE
   FEDEPO% = FECHANUM(FECHA$)
   FEX = FEDEPO%
   FETX$ = FECHATEX$(FEX)
   NRECIB$ = "DB-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
   NRECIBX$ = NRECIB$ + "__"
   ORMA% = (-1)
   '
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where CodiCom_Cor LIKE '" & NRECIBX$ & "' "
   Dim CajaBanc1 As ADODB.Recordset
   Set CajaBanc1 = New ADODB.Recordset
   CajaBanc1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   '
   With CajaBanc1
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          If Val(Mid$(!CodiCom_Cor, 11)) > ORMA% Then
             ORMA% = Val(Mid$(!CodiCom_Cor, 11))
          End If
        .MoveNext
        Loop
      End If
      On Error GoTo 0
   End With
   Set CajaBanc1 = Nothing
   ORMX$ = TRIM$(Str$(ORMA% + 1))
   While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
   NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
   '
   ' ********************************************************
   '
   SUMADEPO# = 0
   For K2% = 1 To JJ&
      DEPASE$ = TRIM$(Mid$(LISTA$(K2%), 100, 46))
      If Left$(DEPASE$, 1) = "-" Then DEPASE$ = TRIM$(Mid$(DEPASE$, 2))
      If Left$(DEPASE$, 8) = "DEPOSITO" Then
           CUEBANCO$ = TRIM$(Mid$(LISTA$(K2%), 83, 16))
           DENBANCO$ = "DEPOSITO " + DENOCUE$(CUEINT%(CUEBANCO$))
           IDEBE# = XVALO(TRIM$(Mid$(LISTA$(K2%), 146, 13)))
           IHABE# = XVALO(TRIM$(Mid$(LISTA$(K2%), 162, 13)))
           ' GRABA EL DEPOSITO
           CoComCor$ = NRECIB$
           CoComLar$ = NRECIB$
'           CodiEmp% = 0
'           CodiSuc% = SUC%
           NumeClie = 0
           NumeProv% = 0
           FechEmi = CVDAT(FEDEPO%)
           '
           Referen$ = REFEX8$
           FeContab = CVDAT(FEDEPO%)
              FEACR% = DIPOST(DIPOST(FEDEPO%))
           FeAcredi = CVDAT(FEACR%)
           NumeTal = Val(Mid$(NRECIB$, 11, 2))
           NumeCom = Val(Mid$(NRECIB$, 4, 6))
           Observa$ = "Depósito Cheques 48 Hs"
           '
           CuentaCo% = CUEINT%(CUEBANCO$)
           TiMovi = "DEPOBAN"            ' DEPOSITO BANCARIO
                   OpMovi = "48 Hs"             ' PLAZO DE ACREDITACION
           VaMovi = ""                   '
           Descrip = "Deposito Cheques 48 Hs"
           ImpoDeb# = IDEBE# - IHABE#
           StatInit = 0
           Call GRACABAN
           '
        Else
           IDEBE# = XVALO(TRIM$(Mid$(LISTA$(K2%), 146, 13)))
           IHABE# = XVALO(TRIM$(Mid$(LISTA$(K2%), 162, 13)))
           SUMADEPO# = SUMADEPO# + IDEBE# - IHABE#
           IPASX$ = TRIM$(FORMATNUM$(Abs(IDEBE# - IHABE#), "F15.2"))
           COCUE$ = TRIM$(Mid$(LISTA$(K2%), 83, 16))
           '
           Dim CABASEL1 As ADODB.Recordset
           Set CABASEL1 = New ADODB.Recordset
           DEPAS1$ = TRIM$(DEPASE$)
           BANUFE$ = DEPAS1$
           For PPXX% = Len(DEPAS1$) To 1 Step -1
             If Mid$(DEPAS1$, PPXX%, 1) = " " Then
                BANUFE$ = AJUSTI$(Left$(DEPAS1$, PPXX% - 1), 18) + Mid$(DEPAS1$, PPXX% + 1)
                GoTo 18
             End If
           Next PPXX%
           '
18    '    BORRA LOS MOVIMIENTOS
           SQLX$ = "DELETE FROM CAJABANC "
           SQLX$ = SQLX$ + "where TipoMovim <> 'CHRECI' "
           SQLX$ = SQLX$ + "AND DescriMov LIKE '" & BANUFE$ & "%' "
           SQLX$ = SQLX$ + "AND ABS(ValAbsComp - " & IPASX$ & ") < 0.005 "
           FLEXCONN.Execute SQLX$
           '
           SQLX$ = "SELECT * FROM CAJABANC "
           SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI' "
           'SQLX$ = SQLX$ + "WHERE Status > 0 "
           SQLX$ = SQLX$ + "AND DescriMov LIKE '" & BANUFE$ & "%' "
           SQLX$ = SQLX$ + "AND ABS(ValAbsComp - " & IPASX$ & ") < 0.005 "
           CABASEL1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
           '
           JJI& = 0
           With CABASEL1
             .MoveFirst
             Do While Not .EOF
               ' MARCA LA SALIDA
               !FEDEPOSENT = CVDAT(FEDEPO%)
               !Document = NRECIB$
               !DESTICHEQ = Left$(DENBANCO$, 64)       ' OJO
                  CUECONTABINT% = !CUECONTAI
               !REFESALCHE = Left$(REFEX8$, 48)
                  HOII% = HOY(HOS$)
               !IDENSALCHE = AJUSTI$(HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 48)
               !UsuaSalChe = AJUSTI$("Usuario: " + TRIM$(USERNAME$), 48)
               !Status = 2                     ' Marca Salida de Cartera
               !IdenBaja = IDACTU$             ' Marca Modificacion de registro - FECHA - HORA - USUARIO - TERMINAL
               .Update
               '
             .MoveNext
             Loop
           End With
           '
           CABASEL1.Close
           '
           ' GRABA LA SALIDA DEL CHEQUE
           CoComCor$ = NRECIB$
           CoComLar$ = NRECIB$
'           CodiEmp% = 0
'           CodiSuc% = SUC%
           NumeClie = 0
           NumeProv% = 0
           FechEmi = CVDAT(FEDEPO%)
           '
           Referen$ = REFEX8$
           FeContab = CVDAT(FEDEPO%)
           FeAcredi = CVDAT(FEDEPO%)
           NumeTal = Val(Mid$(NRECIB$, 11, 2))
           NumeCom = Val(Mid$(NRECIB$, 4, 6))
           '
           CuentaCo% = CUEINT%(COCUE$)
           TiMovi$ = "CHENTR"
           OpMovi$ = "DP"
           VaMovi$ = "48 Hs"
           Descrip = BANUFE$
           ImpoDeb# = IDEBE# - IHABE#
           StatInit = 0
           Call GRACABAN
           '
      End If
      
   Next K2%
   '
   ' verifica el balance
   On Error GoTo 0
   SQLZ$ = "SELECT SUM(IDEBECOMP-IHABECOMP) AS SUMA FROM CAJABANC"
   Dim CABASE As ADODB.Recordset
   Set CABASE = New ADODB.Recordset
   Set CABANSE = FLEXCONN.Execute(FILTSQL$(SQLZ$))
   '
   If Abs(CABANSE!SUMA) >= 0.005 Then
      ' GRABA LA DIFERENCIA
      CoComCor$ = NRECIB$
      CoComLar$ = NRECIB$
'      CodiEmp% = 0
'      CodiSuc% = SUC%
      NumeClie = 0
      NumeProv% = 0
      FechEmi = CVDAT(FEDEPO%)
      '
      Referen$ = REFEX8$
      FeContab = CVDAT(FEDEPO%)
      FeAcredi = CVDAT(FEDEPO%)
      NumeTal = Val(Mid$(NRECIB$, 11, 2))
      NumeCom = Val(Mid$(NRECIB$, 4, 6))
      '
      CuentaCo% = CUEBALCART%
      TiMovi$ = "MIGRA"
      OpMovi$ = ""
      VaMovi$ = ""
      Descrip = BANUFE$
      ImpoDeb# = (-1) * CABANSE!SUMA
      StatInit = 0
      Call GRACABAN
      '
   End If
   CABANSE.Close
   Set CABANSE = Nothing
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
   Dim MIRS As ADODB.Recordset
   Set MIRS = New ADODB.Recordset
   strsql = "SELECT TOP 1 OrFeAsien FROM EncaBEAsien WHERE FECHASIEN = '" & FEASYYYY & "'"
   strsql = strsql & "And SucuAsien = " & SUCUI% & " ORDER BY ORFEASIEN DESC"
   MIRS.Open FILTSQL$(strsql), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   On Error Resume Next
   MIRS.MoveFirst
   If Err Then GoTo 20
   ORDAS% = MIRS!OrFeAsien
   MIRS.Close
   Set MIRS = Nothing
   '
20 On Error GoTo 0
   ORDAS% = ORDAS% + 1
   ORDA$ = TRIM$(Str$(ORDAS%))
   While Len(ORDA$) < 3: ORDA$ = "0" + ORDA$: Wend
   NUPROA$ = NUIDAS$ + "-" + ORDA$
   '
End Function

Sub CONAPLI(APLINVO$, TERMINA%)
   TERMINA% = 1
   Select Case APLINVO$
           '
       Case "AMACLI"
         If EXISTE%("AMACLI09.EXE") > 0 Then
           Call CONECRUN("AMACLI09", "Archivo Maestro de Clientes"):
         Else
           Call Command18_Click  ' Actualización de Base de Datos de Clientes
         End If
         
       Case "AMASTO"
           Call Command1_Click   ' Actualización Maestro de Items de Stock
       Case "AMAPRO"
           If EXISTE%("AMAPRO09.EXE") > 0 Then
              Call CONECRUN("AMAPRO09", "Archivo Maestro de Proveedores"):
           Else
              Call Command19_Click  ' Actualización de Base de Datos de Proveedores
           End If
       Case "OPMASPRO"
           Call ACRE_GES07.Command14_Click  ' Valores por Omisión Base de Datos de Proveedores
       Case "LIMASTER"
           Call OPLIMAS07.LIMASTER(0)
       Case "TABLES"
           OPMAST07.Show 1
       Case "CONFIGURA"
           Call ACLI_GES07.Command14_Click      'Configuración - Tablas de Validación
       Case "TALISPRE"
           Call OPMASCL07.Option1_DblClick(7)   'Tabla de Lista de Precios
       Case "TACONPAG"
           Call OPMASCL07.Option1_DblClick(6)   'Tabla de Condiciones de Pago
       Case "TAPROV"
           Call OPMASCL07.Option1_DblClick(5)   'Tabla de Provincias y Zonas de Ventas
       Case "TAVENDE"
           Call OPMASCL07.Option1_DblClick(4)   'Tabla de Vendedores
       Case "TACUCONT"
           Call OPMASCL07.Option1_DblClick(8)   'Tabla de Cuentas Contables
       Case "TAGRUCLI"
           Call OPMASCL07.Option1_DblClick(9)   'Tabla de Grupos de Clientes
       Case "CONFIGIIBB"
           Call OPMASCL07.Command3_Click        'Configuracion Funcionamiento Percepciones IB Pcia de Bs.As
       Case "CONLISTA"
           Call ACLI_GES07.Command13_Click      'Emisión de Listados De Cleintes
       Case "LISGRAL"
           Call AMACLIS07.Command8_Click        'Listado general de clientes
       Case "LISCODPOS"
           Call AMACLIS07.Command4_Click        'Listado por codigos postales de clientes
       Case "LISDIRMAIL"
           Call AMACLIS07.Command2_Click        'Listado de direcciones de mail de clientes
       Case "LISDEULIM"
           Call AMACLIS07.Command3_Click        'Listado de deudas y limites de creditos de clientes
       
       'A partir de aca son comandos del ABMPROVEEDORES
       
       Case "CONFIGURAP"
           Call ACRE_GES07.Command14_Click      ' configuracion
       Case "CONLISTAP"
           Call ACRE_GES07.Command13_Click      'Consaultas y listados proveedores
       Case "TACONPAGP"
           Call OPMASPR07.Option1_DblClick(6)   'Listado de Condiciones de Pago Proveedores
       Case "TAPROVP"
           Call OPMASPR07.Option1_DblClick(5)   'Listado de Provincias y zonas de Ventas Proveedores
       Case "TACUCONTP"
           Call OPMASPR07.Option1_DblClick(8)   'Listado de Cuentas Contables Proveedores
       Case "TAGRUPRO"
           Call OPMASPR07.Option1_DblClick(9)   'Tabla de Grupo de Proveedores
       Case "LISGRALP"
           Call AMAPRLIS07.Command8_Click        'Listado general de Proveedores
       Case "LISCODPOSP"
           Call AMAPRLIS07.Command4_Click        'Listado por codigos postales de Proveedores
       Case "LISDIRMAILP"
           Call AMAPRLIS07.Command2_Click        'Listado de direcciones de mail de Proveedores
       Case "LIPROVINA"
           Call AMAPRLIS07.Command5_Click        'Listado de deudas y limites de creditos de Proveedores
       Case "IMPOMRP"
           IMPOR08.Option1(24).Value = 1
           For II% = 0 To 23
            IMPOR08.Option1(II%).Enabled = False
           Next II%
           IMPOR08.Option1(25).Enabled = False
           IMPOR08.Show 1
       Case Else
           TERMINA% = 0
   End Select
End Sub

Sub MIGRAALMAMODE()
Picture1.ToolTipText = "Depuración - Almacen de Modelos"
  '
  Call APERBASDAT("ORFAB")
  On Error Resume Next
10 Condi$ = "Cod_Inte >= 0"
  Call BORRAREGISTROS("BADAMODE", Condi$, REG&, "NOMESS")
'  FLEXCONN.Execute (FILTSQL$("DROP TABLE BADAMODE"))
  On Error GoTo 0
  '
'  Call APERBASDAT("MASTER")
  Call GENERBASE("ARTICULOS")
  '
  REGSINBORRAR& = CantRegistros("BADAMODE", Condi$, "NOMESS")
  If REGSINBORRAR& > 0 Then
     TEXTO1$ = ("Se Han Detectado En la Tabla BADAMODE\Contiene " & REGSINBORRAR& & " Registros")
     OPX% = COMALTER%(TEXTO1$ + "\\Volver a Intentar\Cancelar")
     If OPX% < 1 Or OPX% > 1 Then Exit Sub
     If OPX% = 1 Then GoTo 10
  End If
  
  Dim ALMODE As ADODB.Recordset
  Set ALMODE = New ADODB.Recordset
  SQLX$ = "SELECT * FROM BADAMODE"
  ALMODE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  '
  Picture1.ToolTipText = "Conversión de Almacen de Modelos"
  '
  'LECTURA
  UR& = ULTREG&("BADAMODE")
  For U& = 1 To ULTREG&("BADAMODE")
    '
    X1$ = REGLEIDO$("BADAMODE", U&)
    'LECTURA
    CODMOD$ = Mid$(X1$, 1, 16)
    DESCRI$ = Mid$(X1$, 17, 48)
    UBICA$ = Mid$(X1$, 65, 32)
    FRECUENC$ = Mid$(X1$, 97, 16)
    FECHCON% = CVI(Mid$(X1$, 113, 2))
    MATE% = Val(Mid$(X1$, 117, 1))
    DESCRIMAT$ = Mid$(X1$, 119, 16)
    CODART& = CVI(Mid$(X1$, 135, 2))
    CANPIE# = CVS(Mid$(X1$, 137, 4))
  
'    'GRABACION
    On Error GoTo 0
    With ALMODE
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
      !Cod_Mod = Left$(CODMOD$, 16)        'CODIGO DEL MODELO
      !Descripcion = Left$(DESCRI$, 64)    'DESCRIPCION DEL MODELO
      !Ubic = Left$(UBICA$, 32)            'UBICACION
      !Frecuen = Left$(FRECUENC$, 32)      'FRECUENCIA
      !FechUlt_Con = FETEXCOR1$(FECHCON%)  'FECHA ULTIMO CONTROL
       Select Case TRIM$(FRECUENC$)        'FECHA PROXIMO CONTROL
         Case "1 - Semanal": CANDIAS% = 7
         Case "2 - Mensual": CANDIAS% = 30
         Case "3 - Trimestral": CANDIAS% = 90
         Case "4 - Semestral": CANDIAS% = 180
         Case "5 - Anual": CANDIAS% = 360
       End Select
       FEBASE = FECHCON%
       If FEBASE < 1 Then
        GoTo 50
       End If
       FEPROCO% = FEBASE
       While DIENTRE%(FEBASE, FEPROCO%) < CANDIAS%
         FEPROCO% = DIPOST(FEPROCO%)
       Wend
      'On Error Resume Next
      !FechPro_Con = FETEXCOR1$(FEPROCO%)
      'On Error GoTo 0
50    !Material = MATE%                    'MATERIAL
      !Desc_Mat = Left$(DESCRIMAT$, 32)    'DESCRIPCION MATERIAL
      !Cod_Inte = CODART&                  'CODIGO MATERIAL
      !CantPiezas = CANPIE#                'CANTIDAD PIEZAS
       XX$ = ""
       RFF$ = RECFILT$("ANOTAMOD", 1, CODMOD$)
       For JJ& = 1 To Len(RFF$) Step 4
           RENOTA& = CVS(Mid$(RFF$, JJ&, 4))
           XX$ = XX$ + Mid$(REGLEIDO$("ANOTAMOD", RENOTA&), 17, 92) + Chr$(10) + Chr$(13)
       Next JJ&
      !AnotaMod = XX$                      'ANOTADOR
       XXY$ = ""
       RFFF$ = RECFILT$("OBMODELO", 1, CODMOD$)
       For JJJ& = 1 To Len(RFFF$) Step 4
           RENOTA& = CVS(Mid$(RFFF$, JJJ&, 4))
           XXY$ = XXY$ + Mid$(REGLEIDO$("OBMODELO", RENOTA&), 17, 34)
       Next JJJ&
       If TRIM$(CODMOD$) = "852011401" Then
       AA = BB
       End If
      !ObservaMod = TRIM$(XXY$)                    'OBSERVACIONES
       RNC$ = RECFILT$("PLANOMOD", 1, CODMOD$)
       For UU& = 1 To Len(RNC$) Step 4
          REPLANO& = CVS(Mid$(RNC$, UU&, 4))
          If REPLANO& > 0 Then
            If REPLANO& <= ULTREG&("PLANOMOD") Then
              RPLA$ = Mid$(REGLEIDO$("PLANOMOD", REPLANO&), 17)
              IPLA& = (UU& - 1) / 4
              Select Case IPLA&
                Case IPLA& = 0: !FOTO1 = RPLA$
                Case IPLA& = 1: !FOTO2 = RPLA$
                Case IPLA& = 2: !FOTO3 = RPLA$
                Case IPLA& = 3: !FOTO4 = RPLA$
              End Select
            End If
          End If
        Next UU&
      .Update
      Call TRACE(20, U&, UR&)
    End With
  Next U&
  ALMODE.Close
  Set ALMODE = Nothing
End Sub

'
Sub MIGRAREPARA()
  Picture1.ToolTipText = "Depuración - Reparaciones"
  '
  Call APERBASDAT("ORFAB")
  On Error Resume Next
  Condi$ = "NumRepa >= 0"
10 Call BORRAREGISTROS("REPARA", Condi$, REG&, "NOMESS")
  
  On Error GoTo 0
  '
'  Call APERBASDAT("ORFAB")
  Call GENERBASE("OFABRIC")
  '
  REGSINBORRAR& = CantRegistros("REPARA", "NumRepa >= 0", "NOMESS")
  If REGSINBORRAR& > 0 Then
     TEXTO1$ = ("Se Han Detectado En la Tabla REPARA\Contiene " & REGSINBORRAR& & " Registros")
     OPX% = COMALTER%(TEXTO1$ + "\\Volver a Intentar\Cancelar")
     If OPX% < 1 Or OPX% > 1 Then Exit Sub
     If OPX% = 1 Then GoTo 10
  End If
  '
  Dim REPABOMBAS As ADODB.Recordset
  Set REPABOMBAS = New ADODB.Recordset
  SQLX$ = "SELECT * FROM REPARA"
  REPABOMBAS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  '
  Picture1.ToolTipText = "Conversión de Reparaciones de Bombas"
  '
  'LECTURA
  UR& = ULTREG&("PRESREPA")
  For U& = 1 To ULTREG&("PRESREPA")
    '
    X1$ = REGLEIDO$("PRESREPA", U&)
    NP& = CVS(Mid$(X1$, 1, 4))
    'LECTURA
    REFE$ = Mid$(X1$, 5, 60)
    FECHPREPA% = CVI(Mid$(X1$, 65, 2))
    NCLIE% = CVI(Mid$(X1$, 67, 2))
    RASO$ = rasocli$(NCLIE%)
    CODMAT% = CVI(Mid$(X1$, 69, 2))
    DESCRI$ = Mid$(X1$, 71, 0)
    NUMSERIE$ = Mid$(X1$, 71, 18)
    REFE2$ = Mid$(X3$, 89, 12)
    HORASMO& = CVS(Mid$(X3, 101, 4))
    VALMO& = CVS(Mid$(X1$, 105, 4))
    IMPMO& = CVS(Mid$(X1$, 109, 4))
    IMPRE& = CVS(Mid$(X1$, 113, 4))
    IMPTN& = CVS(Mid$(X1$, 117, 4))
    FECHPRESU% = CVI(Mid$(X1$, 121, 2))
    FECHAPR% = CVI(Mid$(X1$, 123, 2))
    NUMPED& = CVS(Mid$(X1$, 125, 4))
    OBSER$ = Mid$(X1$, 129, 128)
    NROREP& = CVS(Mid$(X1$, 125, 4))
    YY$ = ""
    If NROREP& > 0 Then
      YY$ = FILTERGETRECORD("ORDEREPA", 1, MKS$(NROREP&))
    End If
    '
'    'GRABACION
    On Error GoTo 0
    With REPABOMBAS
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
      !NumRepa = NP&                       'NUMERO REPARACION
      !Referen = Left$(REFE$, 64)          'REFERENCIA
      !FechRecep = CVDAT(FECHPREPA%)       'FECHA PLANILLA DE RECEPCION
      !nume_cli = NCLIE%                   'NUMERO CLIENTE
      !raso_cli = RASO$                    'RAZON SOCIAL CLIENTE
      !Cod_Inte = CODMAT%                  'CODIGO DEL MATERIAL
      !Descripcion = Left$(DESCRI$, 64)    'DESCRIPCION DEL MATERIAL
      !NumeroSerie = Left$(NUMSERIE$, 32)  'NUMERO SERIE
      !Referen2 = Left$(REFE2$, 64)        'REFERENCIA NUMERO SERIE
      !HoraMObra = HORASMO&                'HORAS MANO DE OBRA
      !ValMObra = VALMO&                   'VALOR HORA MANO DE OBRA
      !Imp_Mobra = IMPMO&                  'IMPORTE MANO DE OBRA
      !Imp_Repu = IMPRE&                   'IMPORTE REPUESTOS
      !Imp_Tot = IMPTN&                    'IMPORTE TOTOL
      !FechPres = CVDAT(FECHPRESU%)        'FECHA PRESUPUESTO
      !FECHAPRO = CVDAT(FECHAPR%)          'FECHA APROBACION PRESUPUESTO
      !Observa = OBSER$                    'OBSERVACIONES
      !NumRemito = ""
       FF% = FECHANUM("01/01/80")
      !FechEnSol = CVDAT(FF%)
      !domi_cli = ""
      !loca_Cli = ""
      !Tele_Cli = ""
      !Pais_Cli = ""
      !Mail_Cli = ""
      !Accesorios = ""
      !Defec_Infor = ""
      !Trab_Efect = ""
      !PresuEsti = 0
      !Status = 0
      !Usuario_Repu = ""
      !Estima_Horas = 0
      !Or_Compra = ""
      !FOr_Compra = CVDAT(FF%)
      !FechEnt = CVDAT(FF%)
      !Nom_Tra = ""
      !Cuit_Cli = ""
      !Piva_Cli = ""
      !Conductor = ""
      !Patente = 0
      !Lentre_Cli = ""
      !Status_Gar = 0
      !CausaFalla = 0
      !DeCauFalla = ""
      !TipoFalla = 0
      !DeTipFalla = ""
      !NROPED = 0
       XX$ = ""
       RFF$ = RECFILT$("ANOTAREP", 1, MKS$(NP&))
       For JJ& = 1 To Len(RFF$) Step 4
           RENOTA& = CVS(Mid$(RFF$, JJ&, 4))
           XX$ = XX$ + Mid$(REGLEIDO$("ANOTAREP", RENOTA&), 5, 88) + Chr$(10) + Chr$(13)
       Next JJ&
      !AnotaRep = XX$
        ' obtener de tabla ORDEREPA
      If YY$ <> "" Then
        !NUORDREP = CVS(Mid$(YY$, 1, 4))
        !REFORDREP = Mid$(YY$, 5, 60)
        !NROPED = CVS(Mid$(YY$, 125, 4))
        !FECHORDREP = CVDAT(CVI(Mid$(YY$, 121, 2)))
        !Status = CVI(Mid$(YY$, 65, 2))
        !CODREPA = Mid$(YY$, 497, 16)
      Else
        !NUORDREP = 0
        !REFORDREP = ""
        !NROPED = 0
        !FECHORDREP = CVDAT(FECHANUM("01/01/80"))
        !Status = 0
        !CODREPA = ""
      End If
      ' ***
      .Update
      Call TRACE(20, U&, UR&)
    End With
  Next U&
  '
  REPABOMBAS.Close
  Set REPABOMBAS = Nothing
  '
    'TABLA DAT REPUESTOS DE REPARACIONES
  Picture1.ToolTipText = "Depuración - Repuestos de Reparaciones"
  '
  'Call APERBASDAT("ORFAB")
  On Error Resume Next
'  FLEXCONN.Execute (FILTSQL$("DROP TABLE REPUREPA"))
  Condi$ = "NumRepa >= 0"
  Call BORRAREGISTROS("REPUREPA", Condi$, REG&, "NOMESS")
  
  On Error GoTo 0
  '
  Call APERBASDAT("ORFAB")
  Call GENERBASE("OFABRIC")
  '
  Dim REPUBOMBAS As ADODB.Recordset
  Set REPUBOMBAS = New ADODB.Recordset
  SQLX$ = "SELECT * FROM REPUREPA"
  REPUBOMBAS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
  '
  Picture1.ToolTipText = "Conversión de Repuestos de Reparaciones"
  '
  'LECTURA
  URR& = ULTREG&("REPUREPA")
  For UU& = 1 To ULTREG&("REPUREPA")
    '
    X2$ = REGLEIDO$("REPUREPA", UU&)
    NPRE& = CVS(Mid$(X2$, 1, 4))
    'LECTURA
    NCPRE% = CVI(Mid$(X2$, 5, 2))
    CODI% = CVI(Mid$(X2$, 7, 2))
    CANTPRE% = CVS(Mid$(X2$, 9, 4))
    PREU% = CVS(Mid$(X2$, 13, 4))
    IMPO% = CVS(Mid$(X2$, 17, 4))
    CANTRE% = CVS(Mid$(X2$, 21, 4))
    CANTCON% = CVS(Mid$(X2$, 25, 4))
    SALDOC% = CVS(Mid$(X2$, 29, 4))
    NUMOREP$ = Mid$(X2$, 33, 12)

    'GRABACION
    On Error GoTo 0
    With REPUBOMBAS
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
      !NumRepa = NPRE&                     'NUMERO REPARACION
      !Cod_Inte = CODI%                    'CODIGO REPARACION
      !Cant_Sol = CANTPRE%                 'CANTIDAD PRESUPUESTADA
      !NuOrd_Item = XVALO(NUMOREP$)        'NUMERO ORDEN DE REPARACION
      !prun_neto = 0
      !MARBORRA = 0
      !IDENLOTE = ""
      !Impo_Neto = 0
      .Update
      Call TRACE(20, U&, UR&)
    End With
  Next UU&
  REPUBOMBAS.Close
  Set REPUBOMBAS = Nothing
End Sub




