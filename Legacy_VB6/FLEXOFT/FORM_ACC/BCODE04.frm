VERSION 5.00
Begin VB.Form BCODE04 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lectura de Kanban"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6945
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   105
      TabIndex        =   7
      Top             =   185
      Width           =   5790
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   4725
      Width           =   5790
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   6945
      Left            =   6090
      ScaleHeight     =   6885
      ScaleWidth      =   1005
      TabIndex        =   2
      Top             =   -210
      Width           =   1065
      Begin VB.PictureBox Picture16 
         Height          =   5400
         Left            =   0
         ScaleHeight     =   5340
         ScaleWidth      =   1020
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         Begin VB.Image Image3 
            Height          =   5400
            Left            =   0
            Picture         =   "BCODE04.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00004080&
         Caption         =   "Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   630
         Left            =   90
         TabIndex        =   9
         Top             =   3360
         Width           =   630
         Begin VB.Label Label18 
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
            Height          =   270
            Left            =   95
            TabIndex        =   10
            Top             =   240
            Width           =   435
         End
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
         Height          =   690
         Left            =   105
         Picture         =   "BCODE04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   1005
         Width           =   650
      End
      Begin VB.CommandButton Command24 
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
         Picture         =   "BCODE04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Eliminar Producto de la Lista"
         Top             =   2580
         Width           =   650
      End
      Begin VB.CommandButton BOTNEXT 
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
         Height          =   690
         Left            =   105
         Picture         =   "BCODE04.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Continúa Carga de Remito."
         Top             =   4095
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Height          =   690
         Left            =   105
         Picture         =   "BCODE04.frx":2CB6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.CommandButton Command13 
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
         Picture         =   "BCODE04.frx":2E00
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Configuracion de Funcionamiento"
         Top             =   1890
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "BCODE04.frx":3242
         Top             =   4830
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   105
      TabIndex        =   1
      Top             =   450
      Width           =   5790
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu eliminarproductodelalista 
         Caption         =   "Eliminar Producto de la Lista"
      End
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu configurar 
      Caption         =   "Configurar"
      Begin VB.Menu configurarcodigodebarra 
         Caption         =   "Parámetros de Códigos de Barra"
      End
   End
End
Attribute VB_Name = "BCODE04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TOTITFO&

Private Sub BOTNEXT_Click()
  '\\\OT-06-01-0182-WAR-24-05-06///
  Label18 = ""
  Hide
  '\\\OT-06-01-0182-WAR-FIN///
End Sub

Sub Command13_Click()
    '
    Command13.Enabled = False
     STPKBN04.Show 1
    Command13.Enabled = True
    '
End Sub

Private Sub Command2_Click()
  '
  Command2.Enabled = False
  Call CIERRARCH("*.*")
  Screen.MousePointer = 1
  Command2.Enabled = True
  Hide 'Unload Me
  '
End Sub

Private Sub Command24_Click()
    '
    If List1.ListIndex >= 0 Then
      List1.RemoveItem (List1.ListIndex)
      Text1.SetFocus
    End If
    '
End Sub

Private Sub configurarcodigodebarra_Click()
    Call Command13_Click
End Sub

Private Sub eliminarproductodelalista_Click()
   Call Command24_Click
End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame17.BackColor = &HFCD7B6
      
      MARCOBARRA.Top = 5000
      
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
      Set Me.Frame17.Container = Me.Picture16
      Frame17.ZOrder 0

      ' &H8000000F
    End If
    '
    XXTITU$ = Space$(0)
    '
    XXCA$ = "Codigo"
    XXTC$ = " Cant."
    Call REPLA(XXTITU$, XXCA$, 1, 17)
    Call REPLA(XXTITU$, XXTC$, 39, 6)
    '
    List2.AddItem XXTITU$
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
    If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Remito.")
        OPREMIT04.Show 1
        Call FINAL("")
    End If
    TOTITFO& = XVALO(ATRIFORM$(FORIPRE$, "MAX-ITEM"))
    '
End Sub

Private Sub grabar_Click()
   Call BOTNEXT_Click
End Sub

Private Sub List1_Click()
    '
    Text1.SetFocus
    '
End Sub

Private Sub salir_Click()
    Call Command2_Click
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    '
    Static COD1%, COD2%, CANTI1%, CANTI2%, TARD1%, TARD2%, CTXX%
    '
    If CTXX% < 1 Then
       COD1% = CInt(CONTROL$("REMITO", "CODESDE"))
       COD2% = CInt(CONTROL$("REMITO", "CODHASTA"))
       CANTI1% = CInt(CONTROL$("REMITO", "CANDESDE"))
       CANTI2% = CInt(CONTROL$("REMITO", "CANHASTA"))
       TARD1% = CInt(CONTROL$("REMITO", "TARDESDE"))
       TARD2% = CInt(CONTROL$("REMITO", "TARHASTA"))
       CTXX% = 1
    End If
    '
    If KeyAscii = 13 Then
      If Text1 <> "" Then
        '
        XXARTIDES$ = Space$(60)
        '
        XXCODI$ = Left$(Mid(Text1, COD1%, (COD2% - COD1%) + 1), 16)
        XXCANT1$ = Mid(Text1, CANTI1%, (CANTI2% - CANTI1%) + 1)
        XXCANT$ = FORMATNUM$(XVALO(XXCANT1$), "F6.1")
        XXTARD1$ = TRIM$(Mid(Text1, TARD1%, (TARD2% - TARD1%) + 1))
        'VALIDACIONES
        If XVALO(XXCANT1$) < 0.005 Then
          Call MENSERR(24, "Cantidad no Válida")
          Text1.SelStart = 0
          Text1.SelLength = Len(Text1)
          GoTo 99
        End If
        If CODINT%(XXCODI$) <= 0 Then
          Call MENSERR(24, "Código no Válido")
          Text1.SelStart = 0
          Text1.SelLength = Len(Text1)
          GoTo 99
        End If
        '
        Call REPLA(XXARTIDES$, XXCODI$, 1, 16)
        Call REPLA(XXARTIDES$, XXCANT$, 39, 6)
        Call REPLA(XXARTIDES$, XXTARD1$, 61, 4)
        '
        '\\\OT-07-0391-FL-19-10-07///
        If XXTARD1$ <> "" Then 'SI EXISTE EL NUMERO DE TRAJETA, COMPARA QUE NO ESTE REPETIDO EN LA LISTA
             For U& = 1 To List1.ListCount
               If TRIM$(Left$(List1.List(U& - 1), 16)) = TRIM$(XXCODI$) Then ' COINCIDE EL CÓDIGO
                 SECUTARJE$ = Mid$(List1.List(U& - 1), 61)
                 For KU& = 1 To Len(SECUTARJE$) Step 4
                   If TRIM$(Mid$(SECUTARJE$, KU&, 4)) = TRIM$(XXTARD1$) Then
                     Call MENSERR(24, "Tarjeta Ya Ingresada")
                     GoTo 90
                   End If
                 Next KU&
                 '
                 ' ACUMULA EN EL REGISTRO
                 RELISTA$ = List1.List(U& - 1)
                 CANTOTA = XVALO(Mid$(RELISTA$, 39, 6)) + XVALO(XXCANT$)
                 XXCANT$ = FORMATNUM$(CANTOTA, "F6.1")
                 Call REPLA(RELISTA$, XXCANT$, 39, 6)
                 RELISTA$ = RELISTA$ + AJUSTI$(XXTARD1$, 4)
                 List1.RemoveItem U& - 1
                 List1.AddItem RELISTA$, U& - 1
                 GoTo 90
                 '
               End If
             Next U&
             '
             GoTo 80   ' AGREGA EL REGISTRO
             '
          Else ' SI NO EXISTE TARJETA, COMPARA COINCIDENCIA DE  CODIGO Y ACUMULA
             For U& = 1 To List1.ListCount
               If Left$(List1.List(U& - 1), 16) = Left$(XXARTIDES$, 16) Then
                 If TRIM$(Mid$(List1.List(U& - 1), 61, 4)) = "" Then ' el registro leido también sin nro de tarjeta
                   CANTOTA = XVALO(Mid$(List1.List(U& - 1), 39, 6)) + XVALO(XXCANT$)
                   XXCANT$ = FORMATNUM$(CANTOTA, "F6.1")
                   Call REPLA(XXARTIDES$, XXCANT$, 39, 6)
                   List1.RemoveItem U& - 1
                   List1.AddItem XXARTIDES$, U& - 1
                   GoTo 90
                 End If
               End If
             Next U&
        End If
        '\\\OT-07-0391-FIN///
        '
80      If List1.ListCount >= TOTITFO& Then
           Call MENSERR(24, "Imposible Agregar Esta Tarjeta.\Supera Máximo Permitido de " & TOTITFO& & " Renglones.")
           GoTo 90
        End If
        '
        List1.AddItem XXARTIDES$
        Label18 = List1.ListCount
        '
90      Text1 = ""
        Text1.SetFocus
        '
      End If
    End If
    '
99 End Sub
