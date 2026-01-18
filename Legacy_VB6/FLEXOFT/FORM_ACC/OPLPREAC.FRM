VERSION 5.00
Begin VB.Form OPLPREAC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listas de Precios - Complementos"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4020
   Icon            =   "OPLPREAC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   1155
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
         Picture         =   "OPLPREAC.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Unir Listas de Precios "
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
         TabIndex        =   8
         Top             =   105
         Width           =   1380
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   840
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
         Picture         =   "OPLPREAC.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Datos Adicionales de Facturacion"
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
         Left            =   630
         TabIndex        =   2
         Top             =   50
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   525
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
         Picture         =   "OPLPREAC.frx":06D6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenes de Com- pra Terminales"
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
         Left            =   735
         TabIndex        =   5
         Top             =   50
         Width           =   1590
      End
   End
End
Attribute VB_Name = "OPLPREAC"
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

Private Sub Command2_Click()
    Command2.Enabled = False
    Call FINAL("*DATADI")
    Command2.Enabled = True
End Sub

Private Sub Command6_Click()
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
     'If ECOARCH$("LISTAS", Chr$(ILI%)) = "" Then
      SQLX$ = "SELECT * FROM LISTAPRE "
      SQLX$ = SQLX$ + "WHERE Cod_Lista = " & ILI% & ";"
      Set LISTASELEC = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
      On Error Resume Next
      With LISTASELEC
        .MoveFirst
        If Err > 0 Then
          Call MENSERR(24, "Lista Inexistente")
          GoTo 10
        End If
      End With
   Next KKI%
   '
   Screen.MousePointer = 11
   'KKI% identifica el contenido del ingreso
   'de las listas
   For KKI% = 1 To 3
     ILI% = Asc(Mid$(OBX$, KKI%, 1))
     If KKI% = 1 Then PrimeraLista& = ILI%
     If KKI% = 2 Then SegundaLista& = ILI%
     If KKI% = 3 Then ListaDestino& = ILI%
     'LISTAX$(KKI%) = TRIM$(Str$(ILI%))
   Next KKI%
   '
   'PREGUNTA SI LA LISTA DESTINO TIENE DATOS
   '
   'If ULTREG&(LISTAX$(3)) > 0 Then
      SQLX$ = "SELECT * FROM PREVENTA "
      SQLX$ = SQLX$ + "WHERE Cod_Lista = " & ListaDestino& & ";"
      Set LISTASELEC = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
      On Error Resume Next
      With LISTASELEC
        .MoveLast
        If Err = 0 Then
         'CONSULTA SI SE SOBRE ESCRIBE
          If COMALTER%("Lista Destino con Información !!!\\Cancelar\Sobre-Escribir") <> 2 Then
           Exit Sub
          End If
        End If
      End With
        Screen.MousePointer = 1
        
   
   '
   'JJ& = 0
   'BORRAMOS LOS DATOS DE LAS COINCIDENCIAS CON ESA LISTA
   '
   SQLX$ = "DELETE * FROM Preventa "
   SQLX$ = SQLX$ + "WHERE Cod_Lista = " & ListaDestino& & " ;"
   Ventas.Execute (SQLX$)

   'RECORRE LAS DOS LISTAS Y LAS GRABA UNA A CONTINUACION DE LA OTRA
   'SIN IMPORTAR REPETIDOS EN LA LISTA DE DESTINO
   For i& = 1 To 2
       If i& = 1 Then
         SQLX$ = "SELECT * FROM PREVENTA "
         SQLX$ = SQLX$ + "WHERE Cod_Lista = " & PrimeraLista& & ";"
       ElseIf i& = 2 Then
         SQLX$ = "SELECT * FROM PREVENTA "
         SQLX$ = SQLX$ + "WHERE Cod_Lista = " & SegundaLista& & ";"
      End If
         Set ARTSELEC = Ventas.OpenRecordset(SQLX$, 2)
         On Error Resume Next
    With ARTSELEC
           .MoveFirst
           If Err > 0 Then ' **SI NO TIENE INFORMACION EN **
             GoTo 20       ' **LA PRIMERA VA A LA SEGUNDA **
           End If          ' ** SI NO SALE **
           On Error GoTo 0
       Do While Not .EOF
         CODIGO = !Cod_Pieza
         NUMORDEN = !Num_Orden
         DESCPIEZA = !Desc_Pieza
         DENOPIEZA = !Denom_Pieza
         CARCONTROL = !Carac_Control
         NUMORIGINAL = !Num_Original
         PREUNITARIO = !Precio_Unitario
         CANLIMITE = !Cant_Limite
         PORLIMITE = !Porc_Limite
         .MoveNext
         
         Preventa.AddNew   ' YA CONVERTIDO
                 On Error Resume Next
         Preventa!CodiEmpr = CodiEmp%
                 On Error GoTo 0
         Preventa!Cod_Lista = ListaDestino&
         Preventa!Cod_Pieza = CODIGO
         Preventa!Num_Orden = NUMORDEN
         Preventa!Desc_Pieza = DESCPIEZA
         Preventa!Denom_Pieza = DENOPIEZA
         Preventa!Carac_Control = CARCONTROL
         Preventa!Num_Original = NUMORIGINAL
         Preventa!Precio_Unitario = PREUNITARIO
         Preventa!Cant_Limite = CANLIMITE
         Preventa!Porc_Limite = PORLIMITE
         Preventa.Update
       Loop
     End With
20 Next i&
   'For KKI% = 1 To 2
     'FIN& = FINLI&(KKI%)
     'For U& = 1 To FIN&
       'Call TRACE(20, U&, FIN&)
       'X$ = REGLEIDO$(LISTAX$(KKI%), U&) 'LEE EL REGISTRO
       'JJ& = JJ& + 1
       'Call GRAREG(LISTAX$(3), X$, JJ&) 'Y GRABA EN DESTINO
     'Next U&
     'Call CIERRARCH(LISTAX$(KKI%))
   'Next KKI%
   'Call SETULTREG(LISTAX$(3), JJ&)
   'Call CIERRARCH(LISTAX$(3))
   'Call BAJALDISCO
   Screen.MousePointer = 1
   '
   Call COMUNI("Fusión de Listas Completa.\Puede Hacer un Ordenamiento\de la Nueva Lista de Precios.")
   '
   '\\OT-07-0434\\
End Sub

