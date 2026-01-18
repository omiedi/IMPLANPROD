VERSION 5.00
Begin VB.Form CARDOC09 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conectar Documentos"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   BeginProperty Font 
      Name            =   "MS Sans Serif"
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
   ScaleHeight     =   4695
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List3 
      Height          =   3660
      Left            =   4950
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   405
      Width           =   3225
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aprobar"
      Height          =   420
      Left            =   4940
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Guarda los Cambios Realizados"
      Top             =   4080
      Width           =   3225
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   120
      TabIndex        =   7
      Top             =   405
      Width           =   4545
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Nuevo Documento"
      Height          =   420
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   4500
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   5040
      Left            =   8400
      ScaleHeight     =   4980
      ScaleWidth      =   1005
      TabIndex        =   0
      Top             =   -15
      Width           =   1065
      Begin VB.CommandButton Command12 
         BackColor       =   &H00E0E0E0&
         Height          =   585
         Left            =   120
         Picture         =   "CARDOC09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Control de Documentos"
         Top             =   2280
         Width           =   645
      End
      Begin VB.CommandButton AGenRep 
         Height          =   600
         Left            =   120
         Picture         =   "CARDOC09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   3000
         Width           =   650
      End
      Begin VB.CommandButton Command7 
         Height          =   570
         Left            =   105
         Picture         =   "CARDOC09.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir"
         Top             =   200
         Width           =   630
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   1
         Top             =   3840
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   2
            Top             =   -30
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -255
         Picture         =   "CARDOC09.frx":075E
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4080
         Width           =   1515
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Usuarios"
      Height          =   330
      Left            =   4950
      TabIndex        =   11
      Top             =   90
      Width           =   3225
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Documento"
      Height          =   330
      Left            =   120
      TabIndex        =   10
      Top             =   90
      Width           =   4530
   End
End
Attribute VB_Name = "CARDOC09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   '
   'BORRO TODO LOS REGISTRO CON EL DOCUMENTO SELECCIONADO
   COD$ = Mid$(List1.List(List1.ListIndex), 200, 10)  'codigo de documento seleccionado
  
   Call FILTERDELETE("VINCUDOC", 2, MKI$(XVALO(COD$)))
   '
   
   For I = 0 To List3.ListCount - 1
    If List3.Selected(I) = True Then
       REGISTRO$ = Space$(3)
       CODUSU$ = Mid(List3.List(I), 100, 2)
  
       Call REPLA(REGISTRO$, Chr$(TRIM$(CODUSU$)), 1, 1)
       Call REPLA(REGISTRO$, MKI$(XVALO(COD$)), 2, 2)
       Call REGAPP("VINCUDOC", REGISTRO$)
    End If
   '
   Next I
End Sub

Private Sub Command12_Click()
   Hide
   
End Sub

Private Sub Command2_Click()
  
  SELDOC09.Show 1
  


End Sub



Private Sub Command4_Click()

    Call SELECHO("FLEXUSR")
    If VALACT1$("FLEXUSR") = "" Then
       Exit Sub
    Else
       Text3 = VALACT1$("FLEXUSR")
       Label6 = VALACT2$("FLEXUSR")
    End If

End Sub

Private Sub Command7_Click()
  Unload Me
 
End Sub

Private Sub Command8_Click()
    CONSUDOC09.Show 1
End Sub

Private Sub Form_Load()
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    'cargar documentos de usuario
    URE& = ULTREG&("DOCUMEN")
    List1.Clear
    FILA$ = Space(256)
      
    'carga todos los documentos
    For U& = 1 To URE&
      FILA$ = ""
      XX$ = REGLEIDO$("DOCUMEN", U&)
      REG% = CVI(Mid(XX$, 1, 2))
      DENOM$ = Mid(XX$, 3, 64)
      UBICA$ = Mid(XX$, 68, 128)

      Call REPLA(FILA$, DENOM$, 1, 64) 'DENOMINACION
      Call REPLA(FILA$, SAFETEXT(REG%), 200, 10) 'CODDOCU
      'Call REPLA(FILA$, UBICA$, 42, 64) 'UBICACION

      List1.AddItem FILA$
     
    Next U&
    
    ' cargar lista de usuarios
    URE& = ULTREG("FLEXUSR")
    For I& = 1 To URE&
       XX$ = REGLEIDO("FLEXUSR", I&)
       FILA$ = (Space(150))
       CODIGO$ = Asc(Mid(XX$, 1, 2))
       NOMBRE$ = Mid(XX$, 3, 30)
       Call REPLA(FILA$, NOMBRE, 1, 20)
       Call REPLA(FILA$, CODIGO$, 100, 2)
       List3.AddItem FILA$
       'Label2 = Label2 + " " + CODIGO$
    Next I&
    
End Sub

Private Sub List1_Click()
   'SE SELECCIONAN LOS USUARIOS QUE TIENE ACCESO AL DOCUMENTO
   For AX = 0 To List3.ListCount - 1
      List3.Selected(AX) = False
   Next AX
   
   COD2% = XVALO(Mid$(List1.List(List1.ListIndex), 200, 10))  'codigo de documento seleccionado
   
   For I& = 1 To ULTREG("VINCUDOC")
     X$ = REGLEIDO$("VINCUDOC", I&)
     DOCUARC% = CVI(Mid(X$, 2, 2))

     If DOCUARC% = COD2% Then   'Si documento coincide
        CODUSUARC% = Asc(Mid$(X$, 1, 1))
        For J = 0 To List3.ListCount - 1
          USU% = Mid(List3.List(J), 100, 2)
          If CODUSUARC% = USU% Then
             List3.Selected(J) = True
          End If
       Next J
     End If
   Next I&
End Sub


'Private Sub Text3_Change()
'   Label6 = ECOARCH("FLEXUSR", MKI(XVALO((Text3))))
'   For AX = 0 To List2.ListCount - 1
'      List2.Selected(AX) = False
'   Next AX
'
'   For I& = 1 To ULTREG("VINCUDOC")
'     X$ = REGLEIDO$("VINCUDOC", I&)
'     COD1% = Asc(Mid$(X$, 1, 1))
'     DOCUARC% = CVI(Mid(X$, 2, 2))
'
'     If COD1% = XVALO(Text3) Then
'       For J = 0 To List2.ListCount - 1
'          COD2% = XVALO(Mid$(List2.List(J), 200, 10))
'          'If COD1% = COD2% Then List2.Selected(J) = True
'          If DOCUARC% = COD2% Then List2.Selected(J) = True
'       Next J
'     End If
'   Next I&
'End Sub

Private Sub Text3_DblClick()
    Call SELECHO("FLEXUSR")
    If VALACT1$("FLEXUSR") = "" Then
       Exit Sub
    Else
       Text3 = VALACT1$("FLEXUSR")
       Label6 = VALACT2$("FLEXUSR")
    End If
End Sub
