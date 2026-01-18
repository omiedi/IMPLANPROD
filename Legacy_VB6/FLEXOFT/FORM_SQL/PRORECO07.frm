VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form PRORECO07 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de Proveedor Recomendado"
   ClientHeight    =   2940
   ClientLeft      =   165
   ClientTop       =   5535
   ClientWidth     =   5640
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
   ScaleHeight     =   2940
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2730
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   15
      Top             =   630
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Asignar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1680
      TabIndex        =   13
      Top             =   2410
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Código de Artículo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   4500
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   220
         Width           =   1845
      End
      Begin VB.CommandButton Command9 
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
         Height          =   330
         Left            =   4095
         Picture         =   "PRORECO07.frx":0000
         TabIndex        =   10
         Top             =   210
         Width           =   175
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   420
         TabIndex        =   5
         Top             =   315
         Width           =   1830
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   1
         Top             =   630
         Width           =   4110
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00062A70&
      Height          =   5790
      Left            =   4725
      ScaleHeight     =   5730
      ScaleWidth      =   1320
      TabIndex        =   2
      Top             =   -20
      Width           =   1380
      Begin VB.CommandButton Command3 
         Height          =   640
         Left            =   105
         Picture         =   "PRORECO07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cierra el Formulario"
         Top             =   210
         Width           =   640
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   3
         Top             =   2040
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   4
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "PRORECO07.frx":0454
         Top             =   2400
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Proveedores Asignados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   105
      TabIndex        =   6
      Top             =   1230
      Width           =   4500
      Begin VB.CommandButton Command1 
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
         Height          =   330
         Left            =   4095
         Picture         =   "PRORECO07.frx":2376
         TabIndex        =   8
         Top             =   210
         Width           =   175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         TabIndex        =   14
         Top             =   225
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1200
         TabIndex        =   12
         Top             =   330
         Width           =   1830
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   7
         Top             =   630
         Width           =   4110
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "PRORECO07.frx":2680
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "PRORECO07.frx":28B4
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "PRORECO07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Command1.Enabled = False
    If ULTREG("!PROVED1") < 1 Then
       Call COMUNI("Código Sin Proveedor Asignado")
       GoTo 99
    End If
    Call SELECHO("!PROVED1")
    If VALACT1$("!PROVED1") <> "" Then
       Label1 = VALACT1$("!PROVED1")
    End If
99  Command1.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Screen.MousePointer = 11
   CI% = CODINT(Text2)
   NUPROV% = XVALO(Label1)
   '
   If CI% < 1 Then
     Call COMUNI("Falta Seleccionar Código")
     GoTo 99
   End If
   If NUPROV% < 1 Then
     Call COMUNI("Falta Seleccionar Proveedor")
     GoTo 99
   End If
   '
   For U& = 1 To ULTREG("PLACOMP")
      XX$ = REGLEIDO$("PLACOMP", U&)
      CIXI% = CVI(Left$(XX$, 2))
      If CIXI% = CI% Then
        Call REPLA(XX$, MKI$(NUPROV%), 33, 2)
        Call GRAREG("PLACOMP", XX$, U&)
      End If
    Next U&
    Call COMUNI("Cambio Realizado")
99 Command2.Enabled = True
   Screen.MousePointer = 1
End Sub

Private Sub Command3_Click()
   Call LIMPIAR
   Call CIERRARCH("*.*")
   Hide

End Sub

Private Sub Command9_Click()

    Call SELECHO("MASTER")
    If VALACT1$("MASTER") <> "" Then
       Text2 = VALACT1$("MASTER")
    End If

End Sub


Private Sub Form_Load()

  Call SETULTREG("!PROVED1", 0)
  Call APLICACIONSKINS(Me, Picture1)
End Sub

Private Sub Label1_DblClick()
  Call Command1_Click
End Sub

Private Sub Text2_Change()
    Label1 = ""
    CIXI% = CODINT%(Text2)
    Label8 = DESCRIT0$(CIXI%)
    '
    Call ABRECOMPRAS
    '
    Call SETULTREG("!PROVED1", 0)
    If CIXI% > 0 Then
       REBLA$ = REGBLAN$("PROVED1")
       SQLX$ = " SELECT * FROM PROVECOD "
       SQLX$ = SQLX$ + " WHERE COD_INTE = " & CIXI% & " "
       Dim PROVETMP As ADODB.Recordset
       Set PROVETMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
       With PROVETMP
          JJ& = 0
          Do While Not .EOF  ' CARGA EL PROVED1 CON
             XX$ = REBLA$    'LOS PROVEEDORES RELACIONADOS CON EL  CODIGO
             NUMPROV% = XVALO(!NUM_PROV)
             RAZONSOC$ = rasocli$(NUMPROV% * (-1))
             Call REPLA(XX$, MKI$(NUMPROV%), 1, 2)
             Call REPLA(XX$, RAZONSOC$, 3, 30)
             JJ& = JJ& + 1
             Call GRAREG("!PROVED1", XX$, JJ&)
            .MoveNext
          Loop
       End With
       Call SETULTREG("!PROVED1", JJ&)
       PROVETMP.Close
       Set PROVETMP = Nothing
       Call FRESHECHO("!PROVED1")
    End If
    '
    ' PRESENTAR EL PRIMER PROVEEDOR DEL PLACOMP
    ZZZ$ = FILTERGETRECORD$("PLACOMP", 1, MKI$(CIXI%))
    PROVHAB% = CVI(Mid$(ZZZ$, 33, 2)) 'PROVEEDOR
    Label1 = PROVHAB%
End Sub
Sub LIMPIAR()
   Text2 = ""
   Label1 = ""
End Sub

Private Sub Text2_DblClick()
   Call Command9_Click
End Sub

Private Sub label1_Change()
    NPRO% = XVALO(Label1)
    Label9 = rasocli$(NPRO% * (-1))
       
End Sub

