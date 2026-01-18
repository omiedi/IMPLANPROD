VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form EDITINSTRUC 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar Instrucción"
   ClientHeight    =   2925
   ClientLeft      =   1905
   ClientTop       =   3300
   ClientWidth     =   13185
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11779.05
   ScaleMode       =   0  'User
   ScaleWidth      =   14201.53
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2925
      Left            =   21
      TabIndex        =   3
      Top             =   0
      Width           =   13185
      Begin VB.CommandButton cmdcancelar 
         Caption         =   "cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9600
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11350
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton cmdelinminar 
         Caption         =   "Quitar de la lista"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1680
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
         Height          =   285
         Left            =   10
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Instrucción"
         Top             =   270
         Width           =   13065
      End
      Begin VB.TextBox TXTDESCRIPCION 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   10
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   540
         Width           =   13020
      End
   End
   Begin VB.PictureBox Picture999 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   8400
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "EDITINSTRUC.frx":0000
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "EDITINSTRUC.frx":0234
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   855
   End
End
Attribute VB_Name = "EDITINSTRUC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ENTER EN CUALQUIER LUGAR DEL FORMULARIO Y REALIZA CLICK EN EL BOTON
'SE DEBE MARCAR EN TRUE LA PROPIEDAD DEFAULT DEL BOTON


Private Sub cmdcancelar_Click()
   
   Unload Me

End Sub

Private Sub cmdelinminar_Click()
    'QUITAR DE LA LISTA
    ACLI_GES07.CambioInstruccionEntrega% = 1
    Call borrafila_CON_ALTURA(ACLI_GES07.MSFINSTRUC, ACLI_GES07.MSFINSTRUC.Row)
99  Unload Me

End Sub


Private Sub Command1_Click()
   Command1.Enabled = False
   'ACTUALIZAR
  
    If InStr(TXTDESCRIPCION.TEXT, "'") > 0 Then
       Call COMUNI("Se Han Detectado Caracteres Reservados (Comilla Simple)\ El Mismo Serà Reemplazado Por Un Espacio Vacío")
       TXTDESCRIPCION.TEXT = REPLACAR$(TXTDESCRIPCION.TEXT, "'", " ")
    End If

    '
    'CARGA EN LA GRILLA
    With ACLI_GES07
        .MSFINSTRUC.TextMatrix(.MSFINSTRUC.Row, 2) = Me.TXTDESCRIPCION
        ACLI_GES07.CambioInstruccionEntrega% = 1
        Call .ACOMODARMSF
    End With
    '
    Unload Me
    
99  Command1.Enabled = True

End Sub
Sub ACTUALIZA_FLEX(MSF As MSFlexGrid, RENGLON%, Fila%, VALOR1$)
    '
    If MSF.Row < 1 Then Exit Sub
    '
    MSF.TextMatrix(RENGLON%, Fila%) = TRIM$(VALOR1$)
    
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call SELECHO("SECPROD")
    COD$ = VALACT1$("SECPROD")
    If TRIM$(COD$) <> "" Then
      TXTNOPER.TEXT = COD$
    End If
    Command2.Enabled = True

End Sub

Function borrafila_CON_ALTURA(MSF As MSFlexGrid, REG&)

        For i& = REG& To MSF.Rows - 2
          For j& = 1 To MSF.Cols - 1
            MSF.TextMatrix(i&, j) = MSF.TextMatrix(i& + 1, j)
            MSF.Row = i&
          Next
          'PARA QUE NO SE QUEDE CON LA ALTURA ANTERIOR QUE TENIA EL REGISTRO
          'POR QUE  SE VAN MOVIENDO Y LAS ALTURAS PUEDEN SER DIFERENTES
          ACLI_GES07.Label22 = MSF.TextMatrix(i&, 2)
          MSF.RowHeight(MSF.Row) = ACLI_GES07.Label22.Height + 100
          MSF.TextMatrix(MSF.Row, 2) = ACLI_GES07.Label22
        Next
        MSF.Rows = MSF.Rows - 1
        
End Function

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
   Call APLICACIONSKINS(Me)
End Sub

Private Sub TXTDESCRIPCION_Change()
  ACLI_GES07.Label22 = TXTDESCRIPCION
End Sub

Private Sub TXTNOPER_Change()
      CS1 = XVALO(TXTNOPER.TEXT)
      If CS1 > 0 Then
         TXTSECCION = TRIM$(ECOARCH$("SECPROD", MKS$(CS1)))
      Else
         TXTSECCION = ""
      End If

End Sub

