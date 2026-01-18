VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form MULITISEL09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccion de Cheques a Depositar"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8115
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1275
      Left            =   105
      TabIndex        =   11
      Top             =   3255
      Width           =   6945
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   300
         Left            =   5145
         TabIndex        =   17
         Top             =   900
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   300
         Left            =   5145
         TabIndex        =   16
         Top             =   525
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Diferencia:"
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
         Left            =   4095
         TabIndex        =   15
         Top             =   945
         Width           =   960
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe. del Sistema:"
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
         Left            =   3280
         TabIndex        =   14
         Top             =   600
         Width           =   1800
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccionados:"
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
         Left            =   3780
         TabIndex        =   13
         Top             =   225
         Width           =   1380
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5145
         TabIndex        =   12
         Top             =   165
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1470
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00008000&
      Height          =   4530
      Left            =   7140
      ScaleHeight     =   4470
      ScaleWidth      =   1845
      TabIndex        =   2
      Top             =   0
      Width           =   1905
      Begin VB.CommandButton Command2 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Picture         =   "MULITISEL09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Búsqueda por Número, Banco, Importe, etc."
         Top             =   2310
         Width           =   680
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Sort"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Picture         =   "MULITISEL09.frx":1366
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Conmuta Ordenamiento por Fecha / Importe"
         Top             =   1575
         Width           =   680
      End
      Begin VB.CommandButton BOTEXIT 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Picture         =   "MULITISEL09.frx":14B0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Cancelar la Aplicación"
         Top             =   105
         Width           =   680
      End
      Begin VB.CommandButton BOTREC 
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   100
         Picture         =   "MULITISEL09.frx":17BA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Grabar - Guardar Modificaciones"
         Top             =   840
         Width           =   680
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   130
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   3
         Top             =   3100
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -270
         Picture         =   "MULITISEL09.frx":1AC4
         Top             =   3990
         Width           =   1515
      End
   End
   Begin VB.ListBox LICHESEL 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3090
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   105
      Width           =   6945
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   3045
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2205
      Visible         =   0   'False
      Width           =   1065
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "MULITISEL09.frx":39E6
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "MULITISEL09.frx":3C1A
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "MULITISEL09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BOTEXIT_Click()
  Unload Me
End Sub

Private Sub BOTREC_Click()
   BOTREC.Enabled = False
   If XVALO(Label9) <> 0 Then
      Call COMUNI("No se Puede Conciliar Con Diferencia de Valores")
      Exit Sub
   End If
    For I& = 1 To CONCIMAN07.MSF3.Rows - 1
     If CONCIMAN07.MSF3.TextMatrix(I&, 0) = "*" Then
        REG3& = I&
        Exit For
     End If
    Next I&
    
    For I& = 1 To LICHESEL.ListCount
     If LICHESEL.Selected(I& - 1) = True Then
       Z$ = LICHESEL.List(I& - 1)
       NUORD2& = XVALO(Mid$(Z$, 136, 5))
       'SUBO LA INFO A LISTA 1
       With CONCIMAN07.MSF2
         Do While .Rows < J&
          CONCIMAN07.MSF1.Cols = 10
          CONCIMAN07.MSF2.Cols = 10
          CONCIMAN07.MSF1.Rows = J& + 1
          CONCIMAN07.MSF1.TextMatrix(J&, 1) = FECHATEX$(XVALO(Mid$(Z$, 1, 8)))
          CONCIMAN07.MSF1.TextMatrix(J&, 2) = CONCIMAN07.MSF3.TextMatrix(REG3&, 2) 'COMPROBANTE QUE LO TIENE EN LA LISTA 3
          CONCIMAN07.MSF1.TextMatrix(J&, 3) = CONCIMAN07.MSF3.TextMatrix(REG3&, 3) 'FECHA DE ACREDITACION
          CONCIMAN07.MSF1.TextMatrix(J&, 4) = TRIM$(Mid$(Z$, 60, 32))
          CONCIMAN07.MSF1.TextMatrix(J&, 5) = FORMATNUM$(XVALO(Mid$(Z$, 92, 12)), "F12.2")
          CONCIMAN07.MSF1.TextMatrix(J&, 6) = FORMATNUM$(XVALO(Mid$(Z$, 104, 12)), "F12.2")
          CONCIMAN07.MSF1.TextMatrix(J&, 7) = XVALO(Mid$(Z$, 136, 5))
          CONCIMAN07.MSF1.TextMatrix(J&, 8) = TRIM$(Mid$(Z$, 141, 85))
          CONCIMAN07.MSF1.TextMatrix(J&, 9) = XVALO(Mid$(Z$, 149, 6))
          '
          For K& = 1 To MSF2.Rows - 1
             NORD& = XVALO(CONCIMAN07.MSF2.TextMatrix(K&, 7))
             If NORD& = NUORD2& Then
                Call CONCIMAN07.borrafila(CONCIMAN07.MSF2, K&)
             End If
          Next K&
         '
        Loop
        End With
    
       Next I&
    
    
    
    BOTREC.Enabled = True
    End Sub

Private Sub Form_Load()
  With CONCIMAN07.MSF2
    .Redraw = False
    LICHESEL.Clear
    For I& = 1 To .Rows - 1
        FECHCO% = CVINT(.TextMatrix(I&, 1))
        COMPROB$ = TRIM$(.TextMatrix(I&, 2))
        FECHAC% = CVINT(.TextMatrix(I&, 3))
        REFE$ = TRIM$(.TextMatrix(I&, 4))
        DEBEX1# = XVALO(.TextMatrix(I&, 5))
        HABEX1# = XVALO(.TextMatrix(I&, 6))
        NUORD& = XVALO(.TextMatrix(I&, 7))
        TABLA$ = XVALO(.TextMatrix(I&, 8))
        CTACBLE% = XVALO(.TextMatrix(I&, 9))
        '
        Call REPLA(TX$, FECHATEX$(FECHCO%), 1, 8)
        Call REPLA(TX$, COMPROB$, 9, 32)
        Call REPLA(TX$, FECHATEX$(FECHAC%), 52, 8)
        Call REPLA(TX$, REFE$, 60, 32)
        Call REPLA(TX$, FORMATNUM$(DEBEX1#, "F12.2"), 92, 12)
        Call REPLA(TX$, FORMATNUM$(HABEX1#, "F12.2"), 104, 12)
        Call REPLA(TX$, TRIM$(Str$(NUORDE%)), 136, 5)
        Call REPLA(TX$, TABLA$, 141, 8)
        Call REPLA(TX$, SAFETEXT$(CTACBLE%), 149, 6)
        '
        LICHESEL.AddItem TX$
    Next I&
  End With

End Sub

Private Sub LICHESEL_Click()
    SUCHESE# = 0
    For I& = 1 To LICHESEL.ListCount
       If LICHESEL.Selected(I& - 1) = True Then
         SUCHESE# = SUCHESE# + XVALO(Mid$(LICHESEL.List(I& - 1), 41, 12))
       End If
    Next I&
    'ACUTALIZO LABEL
    Label2.Caption = FORMATNUM$(SUCHESE#, "F12.2")
    Label9.Caption = XVALO(Label2) - XVALO(Label8)
    Label2.Refresh
    Label9.Refresh
End Sub


