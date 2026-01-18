VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form IMPSTAHP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procesos de Importacion de Datos"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   45
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0F0&
      Caption         =   "Origen de Datos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   34
      Top             =   2630
      Width           =   11655
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
         TabIndex        =   39
         Top             =   480
         Width           =   220
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   930
         ScaleHeight     =   255
         ScaleWidth      =   7035
         TabIndex        =   37
         Top             =   480
         Width           =   7095
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Height          =   285
            Left            =   15
            TabIndex        =   38
            Top             =   0
            Width           =   15000
         End
      End
      Begin VB.CommandButton Command7 
         Height          =   615
         Left            =   10970
         Picture         =   "IMPSTAHP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Establece Conexion con Origen de Datos"
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   8520
         TabIndex        =   35
         Text            =   " "
         Top             =   460
         Width           =   2295
      End
      Begin VB.Label Label12 
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
         TabIndex        =   41
         Top             =   480
         Width           =   735
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
         Left            =   8280
         TabIndex        =   40
         Top             =   240
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   210
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0F0&
      Caption         =   "Resultado de Pre-Importación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Importación Definitiva"
      Top             =   3675
      Visible         =   0   'False
      Width           =   11655
      Begin VB.PictureBox Picture8 
         BorderStyle     =   0  'None
         Height          =   1965
         Left            =   150
         ScaleHeight     =   1965
         ScaleWidth      =   10215
         TabIndex        =   22
         Top             =   420
         Width           =   10215
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   1965
            Left            =   -20
            ScaleHeight     =   1965
            ScaleWidth      =   10215
            TabIndex        =   24
            Top             =   0
            Width           =   10215
            Begin VB.PictureBox Picture3 
               Height          =   1950
               Left            =   -20
               ScaleHeight     =   1890
               ScaleWidth      =   0
               TabIndex        =   29
               Top             =   0
               Width           =   60
            End
            Begin VB.PictureBox Picture7 
               Height          =   270
               Left            =   0
               ScaleHeight     =   210
               ScaleWidth      =   1.19940e5
               TabIndex        =   26
               Top             =   0
               Width           =   1.20000e5
               Begin VB.Label Label83 
                  AutoSize        =   -1  'True
                  BackColor       =   &H008080FF&
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   150
                  TabIndex        =   30
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  BackColor       =   &H008080FF&
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   40
                  TabIndex        =   28
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label Label82 
                  AutoSize        =   -1  'True
                  BackColor       =   &H008080FF&
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   105
                  TabIndex        =   27
                  Top             =   0
                  Width           =   120
               End
            End
            Begin VB.ListBox List4 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1700
               IntegralHeight  =   0   'False
               ItemData        =   "IMPSTAHP.frx":030A
               Left            =   0
               List            =   "IMPSTAHP.frx":030C
               TabIndex        =   25
               Top             =   250
               Width           =   1.30000e5
            End
         End
         Begin VB.PictureBox Picture10 
            Height          =   1950
            Left            =   -20
            ScaleHeight     =   1890
            ScaleWidth      =   0
            TabIndex        =   23
            Top             =   0
            Width           =   60
         End
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         IntegralHeight  =   0   'False
         ItemData        =   "IMPSTAHP.frx":030E
         Left            =   120
         List            =   "IMPSTAHP.frx":0310
         TabIndex        =   19
         Top             =   2910
         Width           =   10400
      End
      Begin VB.CommandButton Command3 
         Caption         =   "List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   10650
         Picture         =   "IMPSTAHP.frx":0312
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Listado por Impresora"
         Top             =   3570
         Width           =   750
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Hide"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10680
         TabIndex        =   17
         Top             =   560
         Width           =   735
      End
      Begin VB.CommandButton Command2 
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
         Height          =   735
         Left            =   10710
         Picture         =   "IMPSTAHP.frx":061C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Inicia Importación Definitiva"
         Top             =   1785
         Width           =   750
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1725
         Left            =   10360
         TabIndex        =   11
         Top             =   670
         Width           =   225
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         Left            =   120
         Max             =   10000
         TabIndex        =   10
         Top             =   2340
         Width           =   10250
      End
      Begin VB.Label label9 
         AutoSize        =   -1  'True
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6000
         TabIndex        =   43
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Line Line6 
         X1              =   135
         X2              =   135
         Y1              =   420
         Y2              =   2520
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Registros no Procesables"
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
         TabIndex        =   20
         Top             =   2670
         Width           =   2775
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Left            =   10680
         TabIndex        =   15
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         Left            =   10620
         TabIndex        =   14
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label TEPRUE 
         AutoSize        =   -1  'True
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3720
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0F0&
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
      Height          =   4560
      Left            =   120
      TabIndex        =   8
      Top             =   3675
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid DbGrid1 
         Height          =   4100
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7223
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
      TabIndex        =   4
      Top             =   2310
      Width           =   740
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   5
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0F0&
      Caption         =   "Formato de Importacion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2190
      Left            =   4800
      TabIndex        =   2
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
         Height          =   1410
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   5655
      End
      Begin VB.PictureBox Picture1 
         Height          =   300
         Left            =   240
         ScaleHeight     =   240
         ScaleWidth      =   5595
         TabIndex        =   6
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
            TabIndex        =   7
            Top             =   0
            Width           =   5655
         End
      End
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
      Height          =   750
      Left            =   11025
      Picture         =   "IMPSTAHP.frx":0926
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Inicia Pre-Importación"
      Top             =   1260
      Width           =   750
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   1785
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1905
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
      Height          =   750
      Left            =   11025
      Picture         =   "IMPSTAHP.frx":0C30
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   360
      Width           =   750
   End
   Begin VB.ListBox List3 
      Height          =   1425
      Left            =   4620
      TabIndex        =   21
      Top             =   1050
      Visible         =   0   'False
      Width           =   1485
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "IMPSTAHP.frx":0D7A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "IMPSTAHP.frx":0FAE
      TabIndex        =   44
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "de Stocks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   1695
      Left            =   1680
      TabIndex        =   33
      Top             =   1680
      Width           =   4005
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "de Inventarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   645
      Left            =   1050
      TabIndex        =   32
      Top             =   1050
      Width           =   4005
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Importador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   645
      Left            =   315
      TabIndex        =   31
      Top             =   420
      Width           =   4005
   End
   Begin VB.Label NARIM 
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "IMPSTAHP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CANCELIMP%
'
Private objexcel As Excel.Application
Private Libro As Excel.Workbook
Private Hoja As Excel.Worksheet
Sub GENERAORDENDECOMPRA(CI%, CANTI, PROVE%, COSTOX#)
     Dim OrdCompra&(32767)
     NUEVACOMPRA% = 0
     If OrdCompra&(PROVE%) < 1 Then
        NUEVACOMPRA% = 1
        OrdCompra&(PROVE%) = XVALO(PRONUOCO&)
     End If
     If NUEVACOMPRA% > 0 Then
       Call CreaRegistro("OCOMENCA", "NUME_OCOM", NUOCONUE&, "NOMESS")
     End If
     
     HO% = HOY(HOS$)
     LU$ = LUDAT$
     '
     Screen.MousePointer = 11
     '
     'GENERO ORDEN DE COMPRA
     OK% = 1
     '
     'GRABO ORDEN DE COMPRA EN OCOMDETA
     SQLX$ = "SELECT * FROM OCOMDETA WHERE NUME_OCOM =  " & OrdCompra&(PROVE%)
     Dim OCOMTMP As ADODB.Recordset
     Set OCOMTMP = New ADODB.Recordset
25   On Error Resume Next
     OCOMTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     '
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
     End If
     On Error GoTo 0
     '
     With OCOMTMP
        If NUEVACOMPRA% > 0 Then
          .AddNew   ' YA CONVERTIDO
        End If
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        !Nume_Ocom = OrdCompra&(PROVE%)
        !Femi_Ocom = CVDAT(HO%)
        !Npro_Ocom = PROVE%
        !COD_INTE = CI%
        !Uni_Mas = UNIMED$(CI%)
        !cant_ocom = CANTI + XVALO(!cant_ocom)
        !Uni_Com = UNIMED$(CI%)
        !Coef_Unids = 1
            
        !Fentre_Sol = FECHATEX$(HO%)
        !CONSIMAT = 0
        !Cant_Rec = 0
        !Cod_Oper = 0
        !Doc_Orig = "IMPORT. AHP"
        !N_ORSERV = 0
        '
        !NuOrd_Item = 0
        !Raso_Prov = RASOCLI$((-1) * PROVE%)
        !Cod_Exte = CODEXT$(CI%)
        !Descritem = DESCRIT0$(CI%)
        !Delibre = ""
        !Sect_Ocom = ""
        !Medi_Ocom = ""
        !Valo_Unid = COSTOX#
           CANTIGRABADA = XVALO(!Unids_Ocom)
           
        !Unids_Ocom = CANTI + CANTIGRABADA
        !Desc_Item = 0
        !inet_item = (CANTIGRABADA + CANTI) * COSTOX#
        !Tsum_Ocom = 0
        !Stat_Ocom = 0
          MONEX% = 2
          'VER SI SE VAN A PASAR TODOS LOS PRECIOS EN DOLARES
        !MONE_EMIS = MONEX%
        !mone_cos = MONEX%
        !IContab = 0
        !Pre_Unit = COSTOX#
        !prun_neto = COSTOX#
        .Update
     End With
     OCOMTMP.Close
     Set OCOMTMP = Nothing
     '
     If NUEVACOMPRA% < 1 Then Exit Sub
     
     'GRABO ORDEN DE COMPRA EN OCOMENCA
     FechEmis% = HO%
     FEX = FechEmis%
     REFE$ = FECHATEX$(FEX) + " - " + "IMPORT. AHP" + " - " + RASOCLI$((-1) * proveed%)
     CPAG% = CPAGCLI%((-1) * proveed%)
     '
     'SE MODIFICA PORQUE MAS ARRIBA SE UTILIZA UN CREAREGISTRO PARA RESERVAR EL NUMERO EN OCOMENCA
     'SQLX$ = "SELECT * FROM OCOMENCA WHERE NUME_OCOM < 1" 'CONDICION INEXISTENTE
     SQLX$ = "SELECT * FROM OCOMENCA WHERE NUME_OCOM =" & NUOCONUE& 'CONDICION INEXISTENTE
     Dim OCOMENCTMP As ADODB.Recordset
     Set OCOMENCTMP = New ADODB.Recordset
26   On Error Resume Next
     OCOMENCTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     '
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 26
     End If
     On Error GoTo 0
     
     With OCOMENCTMP
        'SE COMENTA PORQUE AHORA SOLO ACTUALIZA
        '.AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        !Nume_Ocom = NUOCONUE&
        !Refe_Ocom = Left$(REFE$, 64)
        !Femi_Ocom = CVDAT(FechEmis%)
        !Npro_Ocom = PROVE%
        !Tsum_Ocom = 0 'MATERIALES
        !Stat_Ocom = 0 '
        !Cpag_Ocom = CPAG%
        !AtSr_Ocom = ""
        !Obse_Ocom = ""
        !Raso_Prov = Left$(RASOCLI$((-1) * proveed%), 48)
        !Fgen_Ocom = Now
        !Usua_Gen = USNBR%
        !tcam_ocom = TICAMONE(2)
        !TOBRU_OCOM = 0
        !IDes_Ocom = 0
        !Iiva_Ocom = 0
        !Total_Ocom = 0
        !Embala_Ocom = 0
        !Obse_Ocom = ""
        !Pdes_Ocom = ""
        !LENTREGA = ""
        !mone_cos = 2
        !MONE_EMIS = 2
        !Anex_Ocom = ""
        !Iotr1_Ocom = 0
        .Update
     End With
     OCOMENCTMP.Close
     Set OCOMENCTMP = Nothing
     'FIN ORDEN DE COMPRA
     
     
     
     
End Sub

'
Sub PREIMPINVSTOAHP()
   '
   ATX% = 0
   '
   U& = 1
   '  Do While Not .Recordset.EOF
      Do While U& < URECORD&
        U& = U& + 1
        DbGrid1.Row = U& - 1
        Call TRACE(20, U&, URECORD&)
        TTXX$ = Space$(254)
        CODEX$ = ""
        On Error Resume Next
        DbGrid1.COL = 0
        CODEX$ = TRIM$(DbGrid1.TEXT) ' .Recordset.Fields(0))   ' TRIM$(DBGrid1.TEXT)
     Loop
   '03/04/07 (OT-07-0141)
   
   If DbGrid1.Cols < 5 Then
      Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
      GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = DbGrid1.Rows 'Data1.Recordset.RecordCount
'   HROW = DbGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
'   DbGrid1.RowHeight = 10
   '
   ATX% = 0
'   HUBOAMAS% = 0
   '
'   With Data1
'     .Recordset.MoveFirst
     U& = 1
     Do While U& < URECORD
        U& = U& + 1
        DbGrid1.Row = U& - 1
        Call TRACE(20, U&, URECORD&)
        TTXX$ = Space$(120)
        '
        CODEX$ = ""                             ' CODIGO DE ARTICULO
        On Error Resume Next
        CODEX$ = DbGrid1.TextMatrix(U& - 1, 0) '.Recordset.Fields(0)
        On Error GoTo 0
        '
5       CDXI% = CODINT%(CODEX$)
        If CDXI% < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Código '" + TRIM$(CODEX$) + "' No Existente en el Maestro de Artículos"
          List1.AddItem MERROX$
          GoTo 19
        End If
        '
12      CODEX$ = CODEXT$(CDXI%)
        Call REPLA(TTXX$, CODEX$, 1, 16) 'CODIGO DEL ARTICULO
        '
        Call REPLA(TTXX$, DESCRIT0$(CDXI%), 19, 48) ' DESCRIPCION DEL ARTICULO
        '
        CANTI = XVALO(DbGrid1.TextMatrix(U& - 1, 2)) ' .Recordset.Fields(2))
        If CANTI <= 0 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Cantidad no valida"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, AJUSTI$(FORMATNUM$(CANTI, "F8"), 8), 67, 8)
        '

        On Error Resume Next
        PROVE% = XVALO(DbGrid1.TextMatrix(U& - 1, 3)) ' .Recordset.Fields(3))
        On Error GoTo 0
        If PROVE% < 1 Then
              MERROX$ = "Registro" + Str$(U& - 1) + " - PROVEEDOR NO VALIDO"
              List1.AddItem MERROX$
              GoTo 19
        End If
        Call REPLA(TTXX$, AJUSTI$(FORMATNUM$(PROVE%, "F16"), 16), 105, 16)
        
        
        On Error Resume Next
        COSTOX# = XVALO(DbGrid1.TextMatrix(U& - 1, 4)) ' .Recordset.Fields(3))
        On Error GoTo 0
        If COSTOX# < 0 Then
              MERROX$ = "Registro" + Str$(U& - 1) + " - COSTO NO VALIDO"
              List1.AddItem MERROX$
              GoTo 19
        End If
        Call REPLA(TTXX$, AJUSTI$(FORMATNUM$(COSTOX#, "f16.5"), 16), 121, 16)
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
   ' End With
   '
20 ' DbGrid1.RowHeight = HROW
   Screen.MousePointer = 1
   '
End Sub
'
Private Sub Combo1_Change()
'   '
'   Data1.RecordSource = Combo1.TEXT
'   Data1.Refresh
'   Data1.Recordset.MoveFirst
'   '
End Sub

Private Sub Combo1_Click()
   '
   'Set objexcel = New Excel.Application
   'Set Libro = objexcel.Workbooks.Open(Label1.Caption)
   CANOJ% = Libro.Sheets.Count
   For KKI% = 1 To CANOJ%
      If Combo1.LIST(KKI% - 1) = Libro.Sheets(KKI%).Name Then
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
    On Error Resume Next
    Set Hoja = Nothing
    Libro.Close: Set Libro = Nothing
    objexcel.Close: Set objexcel = Nothing
    On Error GoTo 0
    Unload Me
End Sub
'
Private Sub Command2_Click()
   Command2.Enabled = False
   Screen.MousePointer = 11
   CANCELIMP% = 0
   '
   Call IMPINVSTOAHP
   '
   Screen.MousePointer = 1
   If CANCELIMP% = 0 Then
     Call COMUNI("Se ha Completado la Importacion Definitiva")
   End If
   CANCELIMP% = 0
   Frame5.Visible = False
   DoEvents
   Command2.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Archivo Origen de Datos"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.Filter = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   Openforms = DoEvents
   Screen.MousePointer = 11
   Label1 = UCase$(CommonDialog1.FileName)
'   On Error Resume Next
'   If Err Then
'      'GoTo 98
'      ERNU$ = TRIM$(Str$(Err.Number))
'      On Error GoTo 0
'      If ERNU$ = "3125" Then GoTo 98    ' falta revisar - 25/01/2008
'      MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") al Abrir la Hoja de Cálculo Especificada. Imposible Completar la Petición. Consulte."
'      MsgBox MENSA$
'      GoTo 99
'   End If
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
       Printer.Print List1.LIST(U& - 1)
     Next U&
     Printer.EndDoc
     Screen.MousePointer = 1
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   '
   Command4.Enabled = False
   Frame5.Visible = True
   List1.Clear
   List4.Clear
   On Error Resume Next
   HScroll1.Value = 0
   VScroll2.Value = 0
   On Error GoTo 0
   '
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
   DBHANT = DbGrid1.Height
   '
   On Error GoTo 0
   Screen.MousePointer = 11
   Label6 = ""
   '
   Call PREIMPINVSTOAHP
   '
   Call COMUNI("Se Ha Completado la Pre-Importación.")
   '
20 DbGrid1.Height = DBHANT
   Command4.Enabled = True
   MXX% = List4.ListCount - 5
   If MXX% < 0 Then MXX% = 0
   VScroll2.Value = 0
   VScroll2.Max = MXX%
   '
   HScroll1.Value = 0
   HScroll1.Max = 1
   If TEPRUE.Width > HScroll1.Width Then
     HScroll1.Max = (TEPRUE.Width - HScroll1.Width) / 10
   End If
   '
99 Screen.MousePointer = 1
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Frame5.Visible = False
   DoEvents
End Sub

Private Sub Command7_Click()
   Screen.MousePointer = 11
   CANFIL& = 10000   ' Hoja.Rows
   CANCOL& = List2.ListCount    ' Hoja.Columns
   JK& = 0
   For I& = 1 To CANFIL&
     YAGRELI% = 0
     For j& = 1 To CANCOL&
       On Error Resume Next
       VALOR = Hoja.Cells(I&, j&)
       If Err Then
          On Error GoTo 0
          Call MENSERR(24, "Imposible Establecer Origen de Datos.")
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
         If j& > DbGrid1.Cols Then DbGrid1.Cols = DbGrid1.Cols + 1
       DbGrid1.COL = j& - 1
       If TRIM$(VALOR) <> "" Then
          DbGrid1.TextMatrix(JK& - 1, j& - 1) = VALOR
       End If
       '
       Label9.Font = DbGrid1.Font
       Label9.FontSize = DbGrid1.CellFontSize
       Label9.FontBold = DbGrid1.CellFontBold
       Label9.Caption = CStr(VALOR) + "**"
       If Label9.Width > DbGrid1.ColWidth(j& - 1) Then
         DbGrid1.ColWidth(j& - 1) = Label9.Width
       End If
     Next j&
     If I& > JK& + 256 Then
        GoTo 99
     End If
   Next I&
   '
99 Screen.MousePointer = 1
   Call COMUNI("Origen de Datos Establecido.")
   '
End Sub

Private Sub Form_Load()
   '03/04/07 (OT-07-0141)
'   Call APLICACIONSKINS(Me)

   Call CENTRAFORM(Me)
   '
   List2.AddItem "  1  Codigo de Articulo         String   <= 15 car"
   List2.AddItem "  2  Descripcion del Articulo   String   <= 64 car"
   List2.AddItem "  3  Fecha                      Date"
   List2.AddItem "  4  Depósito Número            Num      1 - 255"
   List2.AddItem "  5  Stock Ajustado             Num      +/- 999999"
   List2.AddItem "  6  PROVEEDOR                  Num      +/- 32765"
   List2.AddItem "  7  Costo                      Num      +/- 999999999"
   
   Label8 = "Código            Descripción                                     Fecha     Dep     S.Teo.     S.Ajus.  Proveedor       Costo  "
   '03/04/07 (FIN)
   '
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

Private Sub VScroll1_Change()
   Picture5.Top = (-10) * VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
   Picture5.Top = (-10) * VScroll1.Value
End Sub




Private Sub label1_Change()
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

Sub IMPINVSTOAHP()
   '
   Call APERBASDAT("STOCKS")
   SQLX$ = " SELECT DoPriCor FROM MOVISTO WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE CodiMovi = 'AJ' "
   'SQLX$ = SQLX$ + "ORDER BY VAL(SUBSTRING(DoPriCor,4,)) DESC"
   SQLX$ = SQLX$ + " ORDER BY CAST(SUBSTRING(DoPriCor,4,12) AS INTEGER) DESC" ' SE ACTUALIZO PARA SQL

   UNUM& = 0
   'Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
'   Dim MoviTemp As ADODB.Recordset
'   Set MoviTemp = New ADODB.Recordset
'   MoviTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   Set MoviTemp = READSET(SQLX$)
   On Error Resume Next
   MoviTemp.MoveFirst
   If Err < 1 Then
      UNUM& = XVALO(Mid$(MoviTemp!DoPriCor, 4))
   End If
   On Error GoTo 0
   NRO& = 1 + UNUM&
   MoviTemp.Close
   '
   Screen.MousePointer = 11
   FIN& = List4.ListCount
   NOCLOSE% = 1
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.LIST(U& - 1)
     CODIT$ = TRIM$(Left$(TTXX$, 16))
     CIXI% = CODINT%(CODIT$)
     CANTI = XVALO(Mid$(TTXX$, 67, 8))
     PROVE = XVALO(Mid$(TTXX$, 105, 16))
     COSTOX# = XVALO(Mid$(TTXX$, 121, 16))
     
     If Abs(CANTI) >= 0.005 Then
        Call GENERAORDENDECOMPRA(CIXI%, CANTI, PROVE, COSTOX#)
        NRO& = NRO& + 1
     End If
     
     '
   Next U&
   '
   'Movisto.Close
   NOCLOSE% = 0
   Screen.MousePointer = 1
   '
End Sub
