VERSION 5.00
Begin VB.Form CORRCLEAR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Limpieza de Correcciones"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Height          =   700
      Left            =   8160
      Picture         =   "CORRCLEAR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Anula la Observacion"
      Top             =   360
      Width           =   615
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   4920
      Width           =   8775
   End
   Begin VB.CommandButton Command3 
      Height          =   330
      Left            =   7440
      Picture         =   "CORRCLEAR.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Registro Anterior"
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Height          =   330
      Left            =   7440
      Picture         =   "CORRCLEAR.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Registro Siguiente"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "CORRCLEAR.frx":091E
      Top             =   1200
      Width           =   9375
   End
   Begin VB.CommandButton Command1 
      Height          =   700
      Left            =   8880
      Picture         =   "CORRCLEAR.frx":0924
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Guarda la Observacion"
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3630
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Tecnico:"
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
      Left            =   -2040
      TabIndex        =   5
      Top             =   480
      Width           =   3060
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      TabIndex        =   4
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Rutina:"
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
      Left            =   360
      TabIndex        =   3
      Top             =   825
      Width           =   3060
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3615
      TabIndex        =   2
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Módulo:"
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
      Left            =   375
      TabIndex        =   1
      Top             =   465
      Width           =   3060
   End
End
Attribute VB_Name = "CORRCLEAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CARGACORREC()
   CONDI$ = "TECRESPO = '" & Label5 & "' "
   CONDI$ = CONDI$ & " AND MODULO = '" & TRIM$(Left$(Label2, 64)) & "' "
   CONDI$ = CONDI$ & " AND RUTINA = '" & Label3 & "'"
   Text1 = VALOBADA("CORRPROG", "Deta_corr", CONDI$, "NOMESS")
End Sub

Private Sub Command1_Click()
   If TRIM$(Left$(Label2, 64)) <> "" Then
     If TRIM$(Label3) <> "" Then
       If TRIM$(Label5) <> "" Then
         CONDI$ = "TECRESPO = '" & Label5 & "' "
         CONDI$ = CONDI$ & " AND MODULO = '" & TRIM$(Left$(Label2, 64)) & "' "
         CONDI$ = CONDI$ & " AND RUTINA = '" & Label3 & "'"
         QTR& = CantRegistros("CORRPROG", CONDI$)
         If QTR& > 0 Then
            Call ACTUREGISTRO("CORRPROG", "Deta_Corr", CONDI$, Text1, RAFE&, "NOMESS")
         End If
       End If
     End If
   End If
End Sub

Private Sub Command2_Click()
   PXP& = List1.ListIndex
   If PXP& < List1.ListCount - 1 Then
      List1.ListIndex = PXP& + 1
      Call List1_Click
   End If
End Sub

Private Sub Command3_Click()
   PXP& = List1.ListIndex
   If PXP& > 0 Then
      List1.ListIndex = PXP& - 1
      Call List1_Click
   End If
End Sub

Private Sub Command4_Click()
   If TRIM$(Left$(Label2, 64)) <> "" Then
     If TRIM$(Label5) <> "" Then
       CONDI$ = "TECRESPO = '" & Label5 & "' "
       CONDI$ = CONDI$ & " AND MODULO = '" & TRIM$(Left$(Label2, 64)) & "' "
       If Label3 <> "" Then
            CONDI$ = CONDI$ & " AND RUTINA = '" & Label3 & "'"
          Else
            CONDI$ = CONDI$ & " AND (RUTINA = '' OR RUTINA IS NULL)"
       End If
       QTR& = CantRegistros("CORRPROG", CONDI$)
       If QTR& > 0 Then
          ILISTA& = List1.ListIndex
          Call BORRAREGISTROS("CORRPROG", CONDI$, RAFE&, "NOMESS")
          Call Label5_CHANGE
          On Error Resume Next
          List1.ListIndex = ILISTA&
          Call List1_Click
          On Error GoTo 0
       End If
     End If
   End If
End Sub

Private Sub Label5_CHANGE()
   '
   List1.Clear
   CONDI$ = "TECRESPO = '" & Label5 & "' "
   QTR& = CantRegistros("CORRPROG", CONDI$)
   If QTR& < 1 Then
      Call MENSERR(24, "Sin Pendientes para este Tecnico.")
      Unload Me
      Exit Sub
   End If
   '
   Dim CORRPRO As ADODB.Recordset
   Set CORRPRO = New ADODB.Recordset
   Set CORRPRO = FLEXCONN.Execute("SELECT * FROM CORRPROG WHERE " & CONDI$)
   With CORRPRO
      Do While Not .EOF
         TTXX$ = SAFETEXT$(!MODULO) & "; " & SAFETEXT$(!RUTINA)
         List1.AddItem TTXX$
      .MoveNext
      Loop
   End With
   '
   List1.ListIndex = 0
   Call List1_Click
   '
End Sub

Private Sub List1_Click()
   '
   PPXX% = InStr(List1.TEXT, ";")
   If PPXX% > 0 Then
      Label2 = TRIM$(Left$(List1.TEXT, PPXX% - 1))
      Label3 = TRIM$(Mid$(List1.TEXT, PPXX% + 1))
      CONDI$ = "TECRESPO = '" & Label5 & "' "
      CONDI$ = CONDI$ & " AND MODULO = '" & TRIM$(Left$(Label2, 64)) & "' "
      CONDI$ = CONDI$ & " AND RUTINA = '" & Label3 & "'"
      Text1 = VALOBADA("CORRPROG", "Deta_corr", CONDI$, "NOMESS")
   End If
   '
End Sub
