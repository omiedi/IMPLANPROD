VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPLIPREC07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listas de Precios - Complementos"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3840
   Icon            =   "OPLIPREC07.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   1200
      ScaleHeight     =   585
      ScaleWidth      =   2265
      TabIndex        =   11
      Top             =   1920
      Width           =   2325
      Begin VB.CommandButton Command7 
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
         Picture         =   "OPLIPREC07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Importador de Lista de Precios"
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Importar Lista de Precios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   840
         TabIndex        =   13
         Top             =   105
         Width           =   1380
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   9
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   915
      ScaleHeight     =   585
      ScaleWidth      =   2265
      TabIndex        =   6
      Top             =   1365
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
         Height          =   600
         Left            =   0
         Picture         =   "OPLIPREC07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Unir Listas de Precios "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   840
         TabIndex        =   8
         Top             =   105
         Width           =   1380
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   600
      ScaleHeight     =   585
      ScaleWidth      =   2265
      TabIndex        =   0
      Top             =   840
      Width           =   2325
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
         Left            =   -20
         Picture         =   "OPLIPREC07.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Datos Adicionales de Facturacion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   630
         TabIndex        =   2
         Top             =   50
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   285
      ScaleHeight     =   585
      ScaleWidth      =   2265
      TabIndex        =   3
      Top             =   315
      Width           =   2325
      Begin VB.CommandButton Command6 
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
         Picture         =   "OPLIPREC07.frx":09E0
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenes de Com- pra Terminales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   735
         TabIndex        =   5
         Top             =   50
         Width           =   1590
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPLIPREC07.frx":0CEA
      TabIndex        =   10
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "OPLIPREC07.frx":0D64
      Top             =   0
   End
End
Attribute VB_Name = "OPLIPREC07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Command1_Click()
   Command1.Enabled = False
   If DERACCE%("UNILIPRE", "Re-Unir Listas de Precios") = 2 Then
     Call UNIPRECIO
   End If
   Command1.Enabled = True
End Sub

Sub Command2_Click()
    Command2.Enabled = False
    Call FINAL("*DATADI")
    Command2.Enabled = True
End Sub

Sub Command6_Click()
    Command6.Enabled = False
    Call FINAL("*OCOTERM")
    Command6.Enabled = True
End Sub

Sub UNIPRECIO()
   '
   '\\OT-07-0434\\
   Dim LISTAX$(3), FINLI&(3)
   '
   VDEF$ = String$(3, 0)
10 OBX$ = OBJPLA$("UNILISPRE", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   VDEF$ = OBX$
   For KKI% = 1 To 3
     ILI% = Asc(Mid$(OBX$, KKI%, 1))
     If ILI% < 1 Then
       Call MENSERR(24, "Lista Inexistente")
       GoTo 10
     End If
     'VERIFICAMOS SI LAS LISTAS EXISTEN
      SQLX$ = "SELECT * FROM LISTAPRE "
      SQLX$ = SQLX$ + "WHERE Cod_Lista = " & ILI% & ";"
      Dim LISTASELEC As ADODB.Recordset
      Set LISTASELEC = FLEXCONN.Execute(FILTSQL$(SQLX$))

      On Error Resume Next
      With LISTASELEC
        .MoveFirst
        If Err > 0 Then
          Call MENSERR(24, "Lista Inexistente")
          GoTo 10
        End If
      End With
   Next KKI%
   Set LISTASELEC = Nothing
   '
   Screen.MousePointer = 11
   'KKI% identifica el contenido del ingreso
   'de las listas
   For KKI% = 1 To 3
     ILI% = Asc(Mid$(OBX$, KKI%, 1))
     If KKI% = 1 Then PrimeraLista& = ILI%
     If KKI% = 2 Then SEGUNDALISTA& = ILI%
     If KKI% = 3 Then LISTADESTINO& = ILI%
     'LISTAX$(KKI%) = TRIM$(Str$(ILI%))
   Next KKI%
   '
   'PREGUNTA SI LA LISTA DESTINO TIENE DATOS
   '
   SQLX$ = "SELECT * FROM PREVENTA "
   SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTADESTINO& & ";"
   Set LISTASELEC = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   On Error Resume Next
   With LISTASELEC
       Do While Not .EOF
         'CONSULTA SI SE SOBRE ESCRIBE ' MODIFICADA POR QUE SIEMPRE MOSTRABA MENSAJE
         If COMALTER%("Lista Destino con Información !!!\\Cancelar\Sobre-Escribir") <> 2 Then
           On Error GoTo 0
           Screen.MousePointer = 1
           Exit Sub
         End If
         Exit Do
       Loop
   End With
   Screen.MousePointer = 1
   '
   'BORRAMOS LOS DATOS DE LAS COINCIDENCIAS CON ESA LISTA
   SQLX$ = "DELETE  FROM Preventa "
   SQLX$ = SQLX$ + "WHERE Cod_Lista = " & LISTADESTINO& & " ;"
    FLEXCONN.Execute (FILTSQL$(SQLX$))

   'RECORRE LAS DOS LISTAS Y LAS GRABA UNA A CONTINUACION DE LA OTRA
   'SIN IMPORTAR REPETIDOS EN LA LISTA DE DESTINO
   FINI% = 2
   If LISTADESTINO& = SEGUNDALISTA& Then FINI% = 1
   For i& = 1 To FINI%
       If i& = 1 Then
         SQLX$ = "SELECT * FROM PREVENTA "
         SQLX$ = SQLX$ + "WHERE Cod_Lista = " & PrimeraLista& & ";"
        ElseIf i& = 2 And PrimeraLista& <> SEGUNDALISTA& Then
         SQLX$ = "SELECT * FROM PREVENTA "
         SQLX$ = SQLX$ + "WHERE Cod_Lista = " & SEGUNDALISTA& & ";"
        Else
        GoTo 99
      End If
      
      Dim ARTSELEC As ADODB.Recordset
      Set ARTSELEC = New ADODB.Recordset
      ARTSELEC.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
      '
      SQLX$ = "SELECT * FROM PREVENTA"
      Dim CARGAR As ADODB.Recordset
      Set CARGAR = New ADODB.Recordset
      CARGAR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
      '
      On Error Resume Next
      With ARTSELEC
         .MoveFirst
         If Err > 0 Then ' **SI NO TIENE INFORMACION EN **
            GoTo 20       ' **LA PRIMERA VA A LA SEGUNDA **
         End If          ' ** SI NO SALE **
         On Error GoTo 0
         '
         Do While Not .EOF
            Codigo = !Cod_Pieza
            NUMORDEN = !Num_Orden
            DESCPIEZA = !Desc_Pieza
            DENOPIEZA = !Denom_Pieza
            CARCONTROL = !Carac_Control
            NUMORIGINAL = !Num_Original
            PREUNITARIO = !Precio_Unitario
            CANLIMITE = !Cant_Limite
            PORLIMITE = !Porc_Limite
            '
            With CARGAR
              .AddNew   ' YA CONVERTIDO
              On Error Resume Next
              !CodiEmpr = CodiEmp%
              On Error GoTo 0
              !Cod_Lista = LISTADESTINO&
              !Cod_Pieza = Codigo
              !Num_Orden = NUMORDEN
              !Desc_Pieza = DESCPIEZA
              !Denom_Pieza = DENOPIEZA
              !Carac_Control = CARCONTROL
              !Num_Original = NUMORIGINAL
              !Precio_Unitario = PREUNITARIO
              !Cant_Limite = CANLIMITE
              !Porc_Limite = PORLIMITE
              .Update
            End With
         .MoveNext
         Loop
      End With
20 Next i&
   Set CARGAR = Nothing
   Set ARTSELEC = Nothing
   Screen.MousePointer = 1
   '
99 Call COMUNI("Fusión de Listas Completa.\Puede Hacer un Ordenamiento\de la Nueva Lista de Precios.")
   '
End Sub

Sub Command7_Click()
    '
    For i% = 0 To 25
        If i% <> 13 Then IMPOR08.Option1(i%).Enabled = False
    Next i% ' DESHABILITA EL RESTO
    IMPOR08.Option1(13).Value = True ' ACTIVA LA OPCION DE LISTA
    IMPOR08.Picture5.Top = -3450     ' SE UBICA EN LA POSICION DE LISTA DE PRECIOS
    IMPOR08.Show 1
    Unload IMPOR08
    '
End Sub

Private Sub Form_Load()
   Call APLICACIONSKINS(Me)
End Sub
