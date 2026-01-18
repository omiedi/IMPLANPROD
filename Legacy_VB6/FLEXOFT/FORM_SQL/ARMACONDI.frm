VERSION 5.00
Begin VB.Form ARMACONDI 
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Consulta / Destino SQL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6060
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
         Height          =   480
         Left            =   5280
         Picture         =   "ARMACONDI.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Termina - Cierra la Aplicación"
         Top             =   3240
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   3300
         Sorted          =   -1  'True
         TabIndex        =   11
         Text            =   "CMBCAMPOS"
         Top             =   525
         Width           =   2625
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   480
         Sorted          =   -1  'True
         TabIndex        =   10
         Text            =   "CMBCAMPOS"
         Top             =   525
         Width           =   2025
      End
      Begin VB.TextBox Text7 
         Height          =   810
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   1080
         Width           =   5445
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Construye Condicion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   120
         TabIndex        =   1
         Top             =   2040
         Width           =   5820
         Begin VB.ComboBox CMBCAMPOS 
            Height          =   315
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   5
            Text            =   "CMBCAMPOS"
            Top             =   600
            Width           =   2055
         End
         Begin VB.ComboBox CMBOPERADOR 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "ARMACONDI.frx":030A
            Left            =   2400
            List            =   "ARMACONDI.frx":030C
            TabIndex        =   4
            TabStop         =   0   'False
            Text            =   "CMBOPERADOR"
            Top             =   600
            Width           =   1080
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   3600
            TabIndex        =   3
            Text            =   "Combo1"
            Top             =   600
            Width           =   2115
         End
         Begin VB.CommandButton Command3 
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
            Height          =   255
            Left            =   5080
            TabIndex        =   2
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Campo:"
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
            Left            =   100
            TabIndex        =   8
            Top             =   390
            Width           =   1275
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Condicion:"
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
            Left            =   2280
            TabIndex        =   7
            Top             =   390
            Width           =   1275
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor:"
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
            Left            =   3480
            TabIndex        =   6
            Top             =   390
            Width           =   1275
         End
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tabla:"
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
         Left            =   105
         TabIndex        =   14
         Top             =   315
         Width           =   615
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Campo:"
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
         Left            =   2205
         TabIndex        =   13
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Condicion:"
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
         Left            =   120
         TabIndex        =   12
         Top             =   885
         Width           =   1275
      End
   End
End
Attribute VB_Name = "ARMACONDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public APROBO%
Private Sub Combo2_Click()
    If TRIM$(Combo2.TEXT) <> "" Then
       Combo3.Clear: CMBCAMPOS.Clear
       Dim AUX As ADODB.Recordset
       Set AUX = New ADODB.Recordset
       On Error Resume Next
       Set AUX = FLEXCONN.Execute("select * from " & Combo2.TEXT)
       If Err Then
          On Error GoTo 0
          Call MENSERR(24, "Referencia a Tabla SQL '" + TRIM$(Combo2.TEXT) + "' Inexistente.")
          Exit Sub
       End If
       On Error GoTo 0
       For i& = 1 To AUX.Fields.Count
           Combo3.AddItem AUX.Fields(i& - 1).Name
           CMBCAMPOS.AddItem AUX.Fields(i& - 1).Name
       Next i&
    End If
End Sub


Private Sub Command1_Click()
   Command1.Enabled = False
   APROBO% = 1
   Hide
99 Command1.Enabled = True
End Sub

Private Sub Command3_Click()
   '
   If TRIM$(CMBCAMPOS.TEXT) <> "" Then
     If TRIM$(CMBOPERADOR) <> "" Then
       If TRIM$(Combo1.TEXT) <> "" Then
         TTXX$ = CMBCAMPOS.TEXT + " " + CMBOPERADOR.TEXT + " " + Combo1.TEXT
         If Text7 <> "" Then Text7 = Text7 + " AND "
         Text7 = Text7 + TTXX$
       End If
     End If
   End If
   '
End Sub


Private Sub Form_Load()
    Combo1.Clear
    Call ABRECONEXION
    Call MUESTRATABLAS
    '
    CMBOPERADOR.Clear
    CMBOPERADOR.AddItem " =  "
    CMBOPERADOR.AddItem " <  "
    CMBOPERADOR.AddItem " <= "
    CMBOPERADOR.AddItem " >  "
    CMBOPERADOR.AddItem " >= "
    CMBOPERADOR.AddItem " <> "
    CMBOPERADOR.AddItem "LIKE"
    '
    'CARGO EL NOMBRE DE LOS CAMPOS
    With VENTADISE
        For i& = 1 To .MSF.Cols - 1
           Combo1.AddItem .MSF.TextMatrix(0, i&)
        Next i&
    End With

End Sub
Sub MUESTRATABLAS()
    Combo2.Clear: Combo3.Clear
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
22  sqlx$ = "Select name from sysobjects WHERE LEFT(NAME,3) <> 'dt_' and LEFT(NAME,3) <> 'pk_' and LEFT(NAME,3) <> 'sys'"
    Set AUX = FLEXCONN.Execute(sqlx$)
    Do While Not AUX.EOF
       If InStr(UCase$(AUX!Name), "DTPROPERT") > 0 Then GoTo 29
       Combo2.AddItem UCase$(AUX!Name)
29  AUX.MoveNext
    Loop
End Sub

