VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CONTROLFANT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Facturas Anticipo y sus Aplicaciones"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14685
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TXTBUSCAR 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   6000
      Width           =   4560
   End
   Begin MSFlexGridLib.MSFlexGrid MSF3 
      Height          =   1365
      Left            =   4680
      TabIndex        =   17
      Top             =   5760
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   2408
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      GridColor       =   16501405
      GridLines       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox CMBFILTRO 
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   9795
      Width           =   2055
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   8760
      Left            =   13680
      ScaleHeight     =   8700
      ScaleWidth      =   1320
      TabIndex        =   6
      Top             =   0
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   9
         Top             =   8040
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "CONTROLFANT.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir"
         Top             =   225
         Width           =   690
      End
      Begin VB.CommandButton Command11 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "CONTROLFANT.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Grabar Ajustes de Factura Anticipo"
         Top             =   7200
         Width           =   690
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -285
         Picture         =   "CONTROLFANT.frx":0454
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   8280
         Width           =   1515
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   4
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   2925
      Left            =   4680
      TabIndex        =   1
      Top             =   1800
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   5159
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      GridColor       =   16501405
      GridLines       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "CONTROLFANT.frx":2376
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CONTROLFANT.frx":25AA
      TabIndex        =   5
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin MSFlexGridLib.MSFlexGrid MSF6 
      Height          =   5685
      Left            =   0
      TabIndex        =   11
      Top             =   240
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   10028
      _Version        =   393216
      Cols            =   4
      BackColor       =   8421504
      ForeColor       =   16777152
      BackColorFixed  =   16501405
      BackColorSel    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF4 
      Height          =   1245
      Left            =   4680
      TabIndex        =   18
      Top             =   240
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   2196
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      GridColor       =   16501405
      GridLines       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF5 
      Height          =   1725
      Left            =   0
      TabIndex        =   25
      Top             =   7080
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   3043
      _Version        =   393216
      Cols            =   4
      BackColor       =   8421504
      ForeColor       =   16777152
      BackColorFixed  =   16501405
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   5685
      Left            =   0
      TabIndex        =   34
      Top             =   240
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   10028
      _Version        =   393216
      Cols            =   4
      BackColor       =   8421504
      ForeColor       =   0
      BackColorFixed  =   16501405
      BackColorSel    =   16777215
      BackColorBkg    =   14737632
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label17 
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   220
      Left            =   1320
      TabIndex        =   33
      Top             =   6510
      Width           =   1305
   End
   Begin VB.Label Label16 
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo Total:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   230
      Left            =   0
      TabIndex        =   32
      Top             =   6495
      Width           =   1305
   End
   Begin VB.Label Label15 
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Ajustar Importe Exento en :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   9660
      TabIndex        =   31
      Top             =   7880
      Width           =   2625
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12240
      TabIndex        =   30
      Top             =   7800
      Width           =   1305
   End
   Begin VB.Label Label14 
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Ajustar Importe Gravado al 10.50% en:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   8640
      TabIndex        =   29
      Top             =   7200
      Width           =   3825
   End
   Begin VB.Label Label34 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12240
      TabIndex        =   28
      Top             =   7485
      Width           =   1305
   End
   Begin VB.Label Label33 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12240
      TabIndex        =   27
      Top             =   7185
      Width           =   1305
   End
   Begin VB.Label Label12 
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Documentos Aplicados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   190
      Left            =   0
      TabIndex        =   26
      Top             =   6840
      Width           =   3225
   End
   Begin VB.Label Label11 
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones de Ajuste"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4680
      TabIndex        =   24
      Top             =   8040
      Width           =   4305
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4680
      TabIndex        =   23
      Top             =   8280
      Width           =   8895
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4680
      TabIndex        =   22
      Top             =   5040
      Width           =   8895
   End
   Begin VB.Label Label7 
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4680
      TabIndex        =   21
      Top             =   4800
      Width           =   4305
   End
   Begin VB.Label Label6 
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Ajustes de Saldos de Facturas Anticipo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4680
      TabIndex        =   20
      Top             =   5520
      Width           =   3945
   End
   Begin VB.Label Label3 
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle de Aplicaciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   4680
      TabIndex        =   19
      Top             =   1560
      Width           =   4305
   End
   Begin VB.Label Label9 
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Ajustar Importe Gravado al 21.00% en:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   8640
      TabIndex        =   16
      Top             =   7560
      Width           =   3705
   End
   Begin VB.Label Label5 
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Clientes con Factura Anticipo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   3225
   End
   Begin VB.Label Label4 
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Filtrar:"
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
      Left            =   50
      TabIndex        =   14
      Top             =   9600
      Width           =   705
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Campo:"
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
      Left            =   2520
      TabIndex        =   13
      Top             =   9600
      Width           =   705
   End
   Begin VB.Label Label1 
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
      Left            =   105
      TabIndex        =   3
      Top             =   105
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   5880
      TabIndex        =   2
      Top             =   9960
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "CONTROLFANT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Dim TASIVA$(2)
Dim NUFACTUSEL$
Sub BLANFORMU()
    msf2.Rows = 1
    MSF3.Rows = 1
    MSF4.Rows = 1
    MSF5.Rows = 1
    Label8 = ""
    Label10 = ""
    Label33 = ""
    Label34 = ""
    Label13 = ""
    NUFACTUSEL$ = ""
End Sub

Sub CARGA_APLICACIONES(NC%, DAPLI$)
      ' SELECT  * FROM CAJABANC WHERE NuClien = 431 AND TipoMovim = 'FVEN' AND OpciMovim <> 'NC'ORDER BY CodiCom_Lar ASC
      CONDI$ = "CODICOM_LAR = '" & DAPLI$ & "' AND TipoMovim = 'FVEN' AND OpciMovim <> 'NC' AND NUCLIEN  = " & NC%
      RESTO# = XVALO(VALOBADA("CAJABANC", "IDEBECOMP", CONDI$, "NOMESS"))
     '
      SQLX$ = " SELECT  * FROM APLICOBRA WITH(NOLOCK)"
      SQLX$ = SQLX$ + " WHERE NuClien = " & NC%
      SQLX$ = SQLX$ + " AND CoDocOR_LAR= '" & TRIM(DAPLI$) & "'"
      SQLX$ = SQLX$ + " ORDER BY CoDocOr_LAR ASC "
      Set APCOTMP = READSET(SQLX$)
      
      MSF5.Rows = 1
      i& = 0
      With APCOTMP
        Do While Not (.EOF)
          NROY$ = SAFETEXT$(!codocap_lar)
          REFY$ = Left(SAFETEXT$(!refeapli), 18)
          IMPO# = XVALO(!IMPOAPLI)
            If IMPO# > RESTO# Then IMPO# = RESTO#
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
            MSF5.Rows = MSF5.Rows + 1
            i& = i& + 1
            MSF5.TextMatrix(i&, 1) = NROY$
            MSF5.TextMatrix(i&, 2) = REFY$
            MSF5.TextMatrix(i&, 3) = IMPOS$
            RESTO# = RESTO# - IMPO#
          End If
          .MoveNext
          Loop
      End With
      APCOTMP.Close
      Set APCOTMP = Nothing
      
      '
      If RESTO# >= 0.005 Then
         MSF5.Rows = MSF5.Rows + 1
         NROY$ = "Pendiente de Aplicación"
         IMPOS$ = NUMSTRI$(RESTO#, 13, 2, 1)

         MSF5.TextMatrix(MSF5.Rows - 1, 2) = NROY$
         MSF5.TextMatrix(MSF5.Rows - 1, 3) = IMPOS$
      End If
      '
      '
    
End Sub

Function PROXIMO_AJUSTE_FACTUANTI&()

   PROXIMO_AJUSTE_FACTUANTI& = 1 + XVALO(VALOBADA("FACTUANTI", "MAX(AJUSTE_NUMERO)", "AJUSTE_NUMERO > 0", "NOMESS"))

End Function

Private Sub Command1_Click()
   Unload Me
End Sub


Private Sub Command11_VIE_Click()
        Command11.Enabled = False
        
        PROAJ& = PROXIMO_AJUSTE_FACTUANTI&
        i& = 1
        NCLIE% = MSF4.TextMatrix(1, 1)
        If NCLIE% < 1 Then
           Call COMUNI("Falta Seleccionar cliente")
           GoTo 99
        End If
        NFACT$ = TRIM$(MSF3.TextMatrix(i&, 2))
        SALDOX# = XVALO(MSF3.TextMatrix(i&, 3))
        Sign% = 1
        If SALDOX# < XVALO(MSF3.TextMatrix(i&, 4)) Then
           Sign% = (-1)
'        Else
'           SOLO POR LA DIFERENCIA
'           IMPOAPLICAR# = XVALO(MSF3.TextMatrix(I&, 3)) - XVALO(MSF3.TextMatrix(I&, 4))
        End If
        
        IMPOR105# = XVALO(MSF3.TextMatrix(i&, 5))
        GRAV105# = XVALO(MSF3.TextMatrix(i&, 6))
        IMPOR210# = XVALO(MSF3.TextMatrix(i&, 7))
        GRAV210# = XVALO(MSF3.TextMatrix(i&, 8))
        IMPOAPLICAR# = (GRAV210# + IMPOR210# + GRAV105# + IMPOR105#) * Sign%
        TiMovi1$ = "AJUS"
        BASEIMPODE# = 0
        BASEIMPOHA# = 0
        VaMovi1$ = Right$(NFACT$, 1)
880     J& = J& + 1
        If J& = 2 Then
                  TiMovi1$ = "VENTA"
                  IMPOAPLICAR# = GRAV105# + GRAV210# * Sign%
           ElseIf J& = 3 Then
                If IMPOR105# > 0.005 Then
                  TiMovi1$ = "IVA"
                  IMPOAPLICAR# = XVALO(MSF3.TextMatrix(i&, 5)) * Sign%
                  VaMovi1$ = "10.50 %"
                  BASEIMPOHA# = XVALO(MSF3.TextMatrix(i&, 6))
                Else
                  GoTo 888
                End If

           ElseIf J& = 4 Then
                If IMPOR210# > 0.005 Then
                  TiMovi1$ = "IVA"
                  IMPOAPLICAR# = XVALO(MSF3.TextMatrix(i&, 7)) * Sign%
                  VaMovi1$ = "21.00 %"
                  BASEIMPOHA# = XVALO(MSF3.TextMatrix(i&, 8))
                Else
                  GoTo 888
                End If
           End If
           '
           Set RS10 = New ADODB.Recordset
           SQLX$ = "SELECT * FROM FACTUANTI WHERE DOCU_ORIG = '' "
333        On Error Resume Next
           RS10.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 333
           End If
           On Error GoTo 0
           '
           With RS10
                  .AddNew
                  On Error Resume Next
                  !CodiEmpr = CodiEmp%
                  On Error GoTo 0

                  !DOCU_ORIG = NFACT$
                  !VARIMOVIM = VaMovi1$
                  !DOCU_APLI = "AJ-" & SAFETEXT$(PROAJ&)
                  !IMPODEBE = 0
                  If IMPOAPLICAR# > 0.005 Then
                    !IMPOHABE = Abs(IMPOAPLICAR#)
                    !IMPODEBE = 0
                    !BASEIMPOHABE = BASEIMPOHA#
                    !BASEIMPODEBE = 0
                  Else
                    !IMPODEBE = Abs(IMPOAPLICAR#)
                    !BASEIMPODEBE = BASEIMPOHA#
                    !IMPOHABE = 0
                    !BASEIMPOHABE = 0
                  End If
                  '
                  !NUCLIEN = Int(GRUDIVCLI(NCLIE%))
                  !Cli_Real = NCLIE%
                  !RefeCom = Left$(rasocli$(NCLIE%), 64)
                  !NumeCom = XVALO(Mid$(NFACT$, 9, 8))
                  !Fechemisi = CVDAT(HOY(HO$))
                  !Status = 0
                  !TIPOMOVIM = TiMovi1$
                  !COTIZA_DOLAR = COTI_DOLAR#
                  !BASEIMPODEBE = 0
                  !BASEIMPOHABE = BASEIMPOHA#
                  !USUARIO = USERNAME$
                  !OBSERVACIONES = Label10
                  !AJUSTE_NUMERO = PROAJ&
                  .Update
888
                  If J& <= 4 Then
                       GoTo 880
                  End If
            End With
        Call BLANFORMU
        
99      Command11.Enabled = True
End Sub

Private Sub Command11_MASVIEJO_Click()

        Command11.Enabled = False
        '
        If DERACCE%("Ajuste de Facturas Anticipo", "AJUFACAN") < 2 Then GoTo 99
          
        If Abs(XVALO(Label33)) + Abs(XVALO(Label34)) + Abs(XVALO(Label13)) <= 0 Then
           Call COMUNI("No Se Ha Realizado Cambios Para Registrar")
           GoTo 99
        End If
        '
        
        NCLIE% = MSF4.TextMatrix(1, 1)
        If NCLIE% < 1 Then
           Call COMUNI("Falta Seleccionar cliente")
           GoTo 99
        End If
        '
        PROAJ& = PROXIMO_AJUSTE_FACTUANTI&
        i& = 1
        NFACT$ = TRIM$(MSF3.TextMatrix(i&, 2))
        '
        DIFERENCIA105# = XVALO(Label33)
        DIFERENCIA210# = XVALO(Label34)
        DIFERENCIAEXENTO# = XVALO(Label13)
        SIGN105% = 1
        SIGN210% = 1
        SIGNEXENTO% = 1
        If Abs(DIFERENCIA105#) > 0.005 Then
           If DIFERENCIA105# < 0.005 Then SIGN105% = (-1)
        End If
        If Abs(DIFERENCIA210#) > 0.005 Then
           If DIFERENCIA210# < 0.005 Then SIGN210% = (-1)
        End If
        '
        If Abs(DIFERENCIAEXENTO#) > 0.005 Then
           If DIFERENCIAEXENTO# < 0.005 Then SIGNEXENTO% = (-1)
        End If
        J& = 0
        '
        IMPOR210# = XVALO(MSF3.TextMatrix(i&, 7))
        GRAV210# = XVALO(MSF3.TextMatrix(i&, 8))
        IMPOEXENTO# = XVALO(MSF3.TextMatrix(i&, 9))
        IMPOAPLICAR# = DIFERENCIA105# + (DIFERENCIA105# * XVALO(TASIVA$(1)) / 100) + (DIFERENCIA210# + (DIFERENCIA210# * XVALO(TASIVA$(0)) / 100) + (DIFERENCIAEXENTO#))
        TiMovi1$ = "AJUS"
        BASEIMPODE# = 0
        BASEIMPOHA# = 0
        VaMovi1$ = Right$(NFACT$, 1)
880     J& = J& + 1
        If J& = 2 Then

                  TiMovi1$ = "VENTA"
                  IMPOAPLICAR# = (DIFERENCIA105# + DIFERENCIA210# + DIFERENCIAEXENTO#) * SIGNEXENTO%
                  If Abs(DIFERENCIAEXENTO#) > 0.005 Then VaMovi1$ = "EXENTO"
                  BASEIMPOHA# = Abs(DIFERENCIAEXENTO#) 'LUEGO LEE EL QUE CORRESPONDA
                  BASEIMPODE# = Abs(DIFERENCIAEXENTO#)

           ElseIf J& = 3 Then
                If Abs(DIFERENCIA105#) > 0.005 Then
                  TiMovi1$ = "IVA"
                  IMPOAPLICAR# = (DIFERENCIA105# * XVALO(TASIVA$(1)) / 100) * SIGN105%
                  VaMovi1$ = "10.50 %"
                  BASEIMPOHA# = Abs(DIFERENCIA105#) 'LUEGO LEE EL QUE CORRESPONDA
                  BASEIMPODE# = Abs(DIFERENCIA105#)
                Else
                  GoTo 888
                End If

           ElseIf J& = 4 Then
                If Abs(DIFERENCIA210#) > 0.005 Then
                  TiMovi1$ = "IVA"
                  IMPOAPLICAR# = (DIFERENCIA210# * XVALO(TASIVA$(0)) / 100) * SIGN210%
                  VaMovi1$ = "21.00 %"
                  BASEIMPOHA# = Abs(DIFERENCIA210#) 'LUEGO LEE EL QUE CORRESPONDA
                  BASEIMPODE# = Abs(DIFERENCIA210#)
                Else
                  GoTo 888
                End If
           End If
           '
           Set RS10 = New ADODB.Recordset
           SQLX$ = "SELECT * FROM FACTUANTI WHERE DOCU_ORIG = '' "
333        On Error Resume Next
           RS10.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 333
           End If
           On Error GoTo 0
           '
           With RS10
                  .AddNew
                  On Error Resume Next
                  !CodiEmpr = CodiEmp%
                  On Error GoTo 0

                  !DOCU_ORIG = NFACT$
                  !VARIMOVIM = VaMovi1$
                  !DOCU_APLI = "AJ-" & SAFETEXT$(PROAJ&)
                  !IMPODEBE = 0
                  If IMPOAPLICAR# > 0.005 Then
                    !IMPOHABE = Abs(IMPOAPLICAR#)
                    !IMPODEBE = 0
                    !BASEIMPOHABE = Abs(BASEIMPOHA#)
                    !BASEIMPODEBE = 0
                  Else
                    !IMPODEBE = Abs(IMPOAPLICAR#)
                    !IMPOHABE = 0
                    !BASEIMPODEBE = Abs(BASEIMPODE#)
                    !BASEIMPOHABE = 0
                  End If
                  '
                  !NUCLIEN = Int(GRUDIVCLI(NCLIE%))
                  !Cli_Real = NCLIE%
                  !RefeCom = Left$(rasocli$(NCLIE%), 64)
                  !NumeCom = XVALO(Mid$(NFACT$, 9, 8))
                  !Fechemisi = CVDAT(HOY(HO$))
                  !Status = 0
                  !TIPOMOVIM = TiMovi1$
                  !COTIZA_DOLAR = COTI_DOLAR#
                  !USUARIO = USERNAME$
                  !OBSERVACIONES = Label10
                  !AJUSTE_NUMERO = PROAJ&
                  .Update
888
                  If J& < 4 Then
                       GoTo 880 'PARA QUE RECORRA 4 VECES Y VERIFIQUE SI HAY AJUSTE PARA CADA MOVMIEMIENTO
                  End If
            End With
            
        Call BLANFORMU
        'ACTUAILZAR LA LISTA DE CLIENTES QUE CONTIENE EN UNA COLUMMNA EL SALDO
        Call carga_clientes_factura_anticipo
99      Command11.Enabled = True
End Sub
Private Sub Command11_Click()

        Command11.Enabled = False
        '
        If DERACCE%("Ajuste de Facturas Anticipo", "AJUFACAN") < 2 Then GoTo 99
          
        If Abs(XVALO(Label33)) + Abs(XVALO(Label34)) + Abs(XVALO(Label13)) <= 0 Then
           Call COMUNI("No Se Ha Realizado Cambios Para Registrar")
           GoTo 99
        End If
        '
        
        NCLIE% = MSF4.TextMatrix(1, 1)
        If NCLIE% < 1 Then
           Call COMUNI("Falta Seleccionar cliente")
           GoTo 99
        End If
        '
        PROAJ& = PROXIMO_AJUSTE_FACTUANTI&
        i& = 1
        NFACT$ = TRIM$(MSF3.TextMatrix(i&, 2))
        '
        DIFERENCIA105# = XVALO(Label33)
        DIFERENCIA210# = XVALO(Label34)
        DIFERENCIAEXENTO# = XVALO(Label13)
'        SIGN105% = 1
'        SIGN210% = 1
'        SIGNEXENTO% = 1
'        If Abs(DIFERENCIA105#) > 0.005 Then
'           If DIFERENCIA105# < 0.005 Then SIGN105% = (-1)
'        End If
'        If Abs(DIFERENCIA210#) > 0.005 Then
'           If DIFERENCIA210# < 0.005 Then SIGN210% = (-1)
'        End If
'        '
'        If Abs(DIFERENCIAEXENTO#) > 0.005 Then
'           If DIFERENCIAEXENTO# < 0.005 Then SIGNEXENTO% = (-1)
'        End If
        J& = 0
        '
        IMPOR210# = XVALO(MSF3.TextMatrix(i&, 7))
        GRAV210# = XVALO(MSF3.TextMatrix(i&, 8))
        IMPOEXENTO# = XVALO(MSF3.TextMatrix(i&, 9))
        IMPOAPLICAR# = DIFERENCIA105# + (DIFERENCIA105# * XVALO(TASIVA$(1)) / 100) + (DIFERENCIA210# + (DIFERENCIA210# * XVALO(TASIVA$(0)) / 100) + (DIFERENCIAEXENTO#))
        TiMovi1$ = "AJUS"
        BASEIMPODE# = 0
        BASEIMPOHA# = 0
        VaMovi1$ = Right$(NFACT$, 1)
880     J& = J& + 1
        If J& = 2 Then
                  TiMovi1$ = "VENTA"
                  IMPOAPLICAR# = (DIFERENCIA105# + DIFERENCIA210# + DIFERENCIAEXENTO#)
                  If Abs(DIFERENCIAEXENTO#) > 0.005 Then VaMovi1$ = "EXENTO"
                  BASEIMPOHA# = Abs(DIFERENCIAEXENTO#) 'LUEGO LEE EL QUE CORRESPONDA
                  BASEIMPODE# = Abs(DIFERENCIAEXENTO#)

           ElseIf J& = 3 Then
                If Abs(DIFERENCIA105#) > 0.005 Then
                  TiMovi1$ = "IVA"
                  IMPOAPLICAR# = (DIFERENCIA105# * XVALO(TASIVA$(1)) / 100)
                  VaMovi1$ = "10.50 %"
                  BASEIMPOHA# = Abs(DIFERENCIA105#) 'LUEGO LEE EL QUE CORRESPONDA
                  BASEIMPODE# = Abs(DIFERENCIA105#)
                Else
                  GoTo 888
                End If

           ElseIf J& = 4 Then
                If Abs(DIFERENCIA210#) > 0.005 Then
                  TiMovi1$ = "IVA"
                  IMPOAPLICAR# = (DIFERENCIA210# * XVALO(TASIVA$(0)) / 100)
                  VaMovi1$ = "21.00 %"
                  BASEIMPOHA# = Abs(DIFERENCIA210#) 'LUEGO LEE EL QUE CORRESPONDA
                  BASEIMPODE# = Abs(DIFERENCIA210#)
                Else
                  GoTo 888
                End If
           End If
           '
           Set RS10 = New ADODB.Recordset
           SQLX$ = "SELECT * FROM FACTUANTI WHERE DOCU_ORIG = '' "
333        On Error Resume Next
           RS10.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 333
           End If
           On Error GoTo 0
           '
           With RS10
                  .AddNew
                  On Error Resume Next
                  !CodiEmpr = CodiEmp%
                  On Error GoTo 0

                  !DOCU_ORIG = NFACT$
                  !VARIMOVIM = VaMovi1$
                  !DOCU_APLI = "AJ-" & SAFETEXT$(PROAJ&)
                  !IMPODEBE = 0
                  If IMPOAPLICAR# > 0.005 Then
                    !IMPOHABE = Abs(IMPOAPLICAR#)
                    !IMPODEBE = 0
                    !BASEIMPOHABE = Abs(BASEIMPOHA#)
                    !BASEIMPODEBE = 0
                  Else
                    !IMPODEBE = Abs(IMPOAPLICAR#)
                    !IMPOHABE = 0
                    !BASEIMPODEBE = Abs(BASEIMPODE#)
                    !BASEIMPOHABE = 0
                  End If
                  '
                  !NUCLIEN = Int(GRUDIVCLI(NCLIE%))
                  !Cli_Real = NCLIE%
                  !RefeCom = Left$(rasocli$(NCLIE%), 64)
                  !NumeCom = XVALO(Mid$(NFACT$, 9, 8))
                  !Fechemisi = CVDAT(HOY(HO$))
                  !Status = 0
                  !TIPOMOVIM = TiMovi1$
                     If MONEMI% < 1 Then MONEMI% = 1
                    COTI_DOLAR# = TICAMONE(MONEMI%)
                  !COTIZA_DOLAR = COTI_DOLAR#
                  !USUARIO = USERNAME$
                  !OBSERVACIONES = Label10
                  !AJUSTE_NUMERO = PROAJ&
                  .Update
888
                  If J& < 4 Then
                       GoTo 880 'PARA QUE RECORRA 4 VECES Y VERIFIQUE SI HAY AJUSTE PARA CADA MOVMIEMIENTO
                  End If
            End With
            
        Call BLANFORMU
        'ACTUAILZAR LA LISTA DE CLIENTES QUE CONTIENE EN UNA COLUMMNA EL SALDO
        Call carga_clientes_factura_anticipo
99      Command11.Enabled = True
End Sub

Private Sub Command11_VIEJO_Click()

        Command11.Enabled = False
        '
        If DERACCE%("Ajuste de Facturas Anticipo", "AJUFACAN") < 2 Then GoTo 99
          
        If Abs(XVALO(Label33)) + Abs(XVALO(Label34)) + Abs(XVALO(Label13)) <= 0 Then
           Call COMUNI("No Se Ha Realizado Cambios Para Registrar")
           GoTo 99
        End If
        '
        
        NCLIE% = MSF4.TextMatrix(1, 1)
        If NCLIE% < 1 Then
           Call COMUNI("Falta Seleccionar cliente")
           GoTo 99
        End If
        '
        PROAJ& = PROXIMO_AJUSTE_FACTUANTI&
        i& = 1
        NFACT$ = TRIM$(MSF3.TextMatrix(i&, 2))
        '
        DIFERENCIA105# = XVALO(Label33)
        DIFERENCIA210# = XVALO(Label34)
        DIFERENCIAEXENTO# = XVALO(Label13)
        SIGN105% = 1
        SIGN210% = 1
        SIGNEXENTO% = 1
        If Abs(DIFERENCIA105#) > 0.005 Then
           If DIFERENCIA105# < 0.005 Then SIGN105% = (-1)
        End If
        If Abs(DIFERENCIA210#) > 0.005 Then
           If DIFERENCIA210# < 0.005 Then SIGN210% = (-1)
        End If
        '
        If Abs(DIFERENCIAEXENTO#) > 0.005 Then
           If DIFERENCIAEXENTO# < 0.005 Then SIGNEXENTO% = (-1)
        End If
        J& = 0
        '
        IMPOR210# = XVALO(MSF3.TextMatrix(i&, 7))
        GRAV210# = XVALO(MSF3.TextMatrix(i&, 8))
        IMPOEXENTO# = XVALO(MSF3.TextMatrix(i&, 9))
        IMPOAPLICAR# = DIFERENCIA105# + (DIFERENCIA105# * XVALO(TASIVA$(1)) / 100) + (DIFERENCIA210# + (DIFERENCIA210# * XVALO(TASIVA$(0)) / 100) + (DIFERENCIAEXENTO#))
        TiMovi1$ = "AJUS"
        BASEIMPODE# = 0
        BASEIMPOHA# = 0
        VaMovi1$ = Right$(NFACT$, 1)
880     J& = J& + 1
        If J& = 2 Then

                  TiMovi1$ = "VENTA"
                  IMPOAPLICAR# = (DIFERENCIA105# + DIFERENCIA210# + DIFERENCIAEXENTO#) * SIGNEXENTO%
                  If Abs(DIFERENCIAEXENTO#) > 0.005 Then VaMovi1$ = "EXENTO"
                  BASEIMPOHA# = Abs(DIFERENCIAEXENTO#) 'LUEGO LEE EL QUE CORRESPONDA
                  BASEIMPODE# = Abs(DIFERENCIAEXENTO#)

           ElseIf J& = 3 Then
                If Abs(DIFERENCIA105#) > 0.005 Then
                  TiMovi1$ = "IVA"
                  IMPOAPLICAR# = (DIFERENCIA105# * XVALO(TASIVA$(1)) / 100) * SIGN105%
                  VaMovi1$ = "10.50 %"
                  BASEIMPOHA# = Abs(DIFERENCIA105#) 'LUEGO LEE EL QUE CORRESPONDA
                  BASEIMPODE# = Abs(DIFERENCIA105#)
                Else
                  GoTo 888
                End If

           ElseIf J& = 4 Then
                If Abs(DIFERENCIA210#) > 0.005 Then
                  TiMovi1$ = "IVA"
                  IMPOAPLICAR# = (DIFERENCIA210# * XVALO(TASIVA$(0)) / 100) * SIGN210%
                  VaMovi1$ = "21.00 %"
                  BASEIMPOHA# = Abs(DIFERENCIA210#) 'LUEGO LEE EL QUE CORRESPONDA
                  BASEIMPODE# = Abs(DIFERENCIA210#)
                Else
                  GoTo 888
                End If
           End If
           '
           Set RS10 = New ADODB.Recordset
           SQLX$ = "SELECT * FROM FACTUANTI WHERE DOCU_ORIG = '' "
333        On Error Resume Next
           RS10.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 333
           End If
           On Error GoTo 0
           '
           With RS10
                  .AddNew
                  On Error Resume Next
                  !CodiEmpr = CodiEmp%
                  On Error GoTo 0

                  !DOCU_ORIG = NFACT$
                  !VARIMOVIM = VaMovi1$
                  !DOCU_APLI = "AJ-" & SAFETEXT$(PROAJ&)
                  !IMPODEBE = 0
                  If IMPOAPLICAR# > 0.005 Then
                    !IMPOHABE = Abs(IMPOAPLICAR#)
                    !IMPODEBE = 0
                    !BASEIMPOHABE = Abs(BASEIMPOHA#)
                    !BASEIMPODEBE = 0
                  Else
                    !IMPODEBE = Abs(IMPOAPLICAR#)
                    !IMPOHABE = 0
                    !BASEIMPODEBE = Abs(BASEIMPODE#)
                    !BASEIMPOHABE = 0
                  End If
                  '
                  !NUCLIEN = Int(GRUDIVCLI(NCLIE%))
                  !Cli_Real = NCLIE%
                  !RefeCom = Left$(rasocli$(NCLIE%), 64)
                  !NumeCom = XVALO(Mid$(NFACT$, 9, 8))
                  !Fechemisi = CVDAT(HOY(HO$))
                  !Status = 0
                  !TIPOMOVIM = TiMovi1$
                     If MONEMI% < 1 Then MONEMI% = 1
                     COTI_DOLAR# = TICAMONE(MONEMI%)
                  !COTIZA_DOLAR = COTI_DOLAR#
                  !USUARIO = USERNAME$
                  !OBSERVACIONES = Label10
                  !AJUSTE_NUMERO = PROAJ&
                  .Update
888
                  If J& < 4 Then
                       GoTo 880 'PARA QUE RECORRA 4 VECES Y VERIFIQUE SI HAY AJUSTE PARA CADA MOVMIEMIENTO
                  End If
            End With
            
        Call BLANFORMU
        'ACTUAILZAR LA LISTA DE CLIENTES QUE CONTIENE EN UNA COLUMMNA EL SALDO
        Call carga_clientes_factura_anticipo
99      Command11.Enabled = True
End Sub



Private Sub Command4_Click()
    Command4.Enabled = False
    RECUEC07.Text1.TEXT = Label111
    
    Command4.Enabled = True

End Sub

Private Sub Form_Load()
    Call APLICACIONSKINS(Me)
    '
    Call ABRECABAN
    '
    For U& = 1 To 2
      X$ = REGLEIDO$("GRUPIVA", U&)
      TASIVA$(U& - 1) = FORMATNUM$(CVS(Mid$(X$, 13, 4)), "F5.2")
    Next U&
    CAMPOS$ = "Cta./F6;R.Social/A18;Comprobante/A18;Imp.Original/F12.2;F.Emis./D8"
    Call CARGA_ENCABEZADO(CONTROLFANT.MSF4, CAMPOS$, CONTROLFANT)

    CAMPOS$ = "Cta/f5;Razon Social/A18;Saldo/F12.2"
    Call CARGA_ENCABEZADO(CONTROLFANT.MSF1, CAMPOS$, CONTROLFANT)
    Call CARGA_ENCABEZADO(CONTROLFANT.MSF6, CAMPOS$, CONTROLFANT)
    '
    CAMPOS$ = "Importe (+) /F12.2;Importe (-) /F12.2;Saldo/F12.2;Emisión/D8;Docu.Aplicado/A20;Fact.Anticipo/A20;N.Aj./F5.O"
    Call CARGA_ENCABEZADO(CONTROLFANT.msf2, CAMPOS$, CONTROLFANT)
    '
    CAMPOS$ = "F.Emis/D8;N.Cte./A24;Saldo/F12.2;Imp.Ajust/F12.2;Iva 10.50 %/F12.2;Grav. 10.5%/F12.2;Iva 21.00 %/F12.2;Grav. 21.0%/F12.2;Imp.Exento/F12.2"
    Call CARGA_ENCABEZADO(CONTROLFANT.MSF3, CAMPOS$, CONTROLFANT)
    '
    CAMPOS$ = "Documento/A17;Referencia/A20;Aplicación/F12.2"
    Call CARGA_ENCABEZADO(CONTROLFANT.MSF5, CAMPOS$, CONTROLFANT)
    
    Call carga_clientes_factura_anticipo
    Call PINTATEXT(TXTBUSCAR)
End Sub
Sub carga_clientes_factura_anticipo()

    MSF6.Rows = 1
    SQLX$ = " Select NUCLIEN,SUM(IMPODEBE) AS IDEBE, SUM(IMPOHABE) AS IHABE from factuanti with(nolock)"
    SQLX$ = SQLX$ + " where nuclien > 0 "
    SQLX$ = SQLX$ + " and (tipomovim = 'FVEN'"
    SQLX$ = SQLX$ + " OR tipomovim = 'AJUS')"
    SQLX$ = SQLX$ + " GROUP BY NUCLIEN"
    SQLX$ = SQLX$ + " ORDER BY NUCLIEN"
    '
    AcuAntiSaldo# = 0
    Set RS = READSET(SQLX$)
    J& = 0
    With RS
        Do While Not .EOF
           NC% = XVALO(!NUCLIEN)
           RASO$ = rasocli$(NC%)
           AntiSaldo# = XVALO(!IDEBE) - XVALO(!IHABE)
           '
           J& = J& + 1
           MSF6.Rows = J& + 1
           MSF6.TextMatrix(J&, 1) = SAFETEXT$(NC%)
           MSF6.TextMatrix(J&, 2) = RASO$
           MSF6.TextMatrix(J&, 3) = FORMATNUM$(AntiSaldo#, "F12.2")
           If AntiSaldo# > 0.005 Then Call COLOREAR_TEXTO_SOLO_UNA_AFILA(CONTROLFANT.MSF6, &H80FFFF, J&)
           AcuAntiSaldo# = AcuAntiSaldo# + AntiSaldo#
          .MoveNext
        Loop
    End With
    RS.Close
    Set RS = Nothing
    MSF6.COL = 2
    MSF6.Sort = 1
    Me.Label17 = TRIM$(FORMATNUM$(AcuAntiSaldo#, "F12.2"))
End Sub


   
Private Sub MSF1_Click()
    i& = MSF1.MouseRow
    J& = MSF1.MouseCol
    If i& = 0 Or i& > MSF1.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF1.Row = i&
    MSF1.COL = J&
    On Error Resume Next

    NC% = XVALO(MSF1.TextMatrix(i&, 1))
    If NC% = 0 Then
        Exit Sub
    End If
    '
    Call BLANFORMU
    
    Call CARGA_SOLO_FACTURA_ANTICIPO_X_CLIENTE(NC%)
    
End Sub
Sub CARGA_FACTURA_ANTICIPO_X_CLIENTE(NC%, NUFACTU$)
   
    If NC% < 1 Then Exit Sub
    NUFACTU$ = TRIM$(NUFACTU$)
    
    SQLX$ = " Select NUCLIEN,SUM(IMPODEBE) AS IDEBE, SUM(IMPOHABE)AS IHABE ,FECHEMISI, DOCU_APLI,DOCU_ORIG, AJUSTE_NUMERO  from factuanti with(nolock)"
    SQLX$ = SQLX$ + " where nuclien =" & NC%
    SQLX$ = SQLX$ + " AND  DOCU_ORIG ='" & NUFACTU$ & "'"
    SQLX$ = SQLX$ + " AND (TIPOMOVIM = 'FVEN' or TIPOMOVIM = 'AJUS')"
    SQLX$ = SQLX$ + " GROUP BY NUCLIEN,DOCU_ORIG,DOCU_APLI,FECHEMISI, AJUSTE_NUMERO"
    SQLX$ = SQLX$ + " ORDER BY FECHEMISI"
    Set RS = READSET(SQLX$)
    J& = 0
    With RS
        Do While Not .EOF
           NC% = XVALO(!NUCLIEN)
           RASO$ = rasocli$(NC%)
           AntiSaldo# = XVALO(!IDEBE) - XVALO(!IHABE)
           '
           J& = J& + 1
           msf2.Rows = J& + 1
           msf2.TextMatrix(J&, 1) = FORMATNUM$((XVALO(!IDEBE)), "F12.2")
           msf2.TextMatrix(J&, 2) = FORMATNUM$((XVALO(!IHABE)), "F12.2")
           SALDO# = SALDO# + XVALO(!IDEBE) - XVALO(!IHABE)
           msf2.TextMatrix(J&, 3) = FORMATNUM$(SALDO#, "F12.2")
           msf2.TextMatrix(J&, 4) = FECHATEX$(CVINT(!Fechemisi))
           msf2.TextMatrix(J&, 5) = SAFETEXT$(!DOCU_APLI)
           msf2.TextMatrix(J&, 6) = SAFETEXT$(!DOCU_ORIG)
           msf2.TextMatrix(J&, 7) = SAFETEXT$(!AJUSTE_NUMERO)
          .MoveNext
        Loop
    End With
    RS.Close
    Set RS = Nothing

End Sub

Sub CARGA_SOLO_FACTURA_ANTICIPO_X_CLIENTE(NC%)
   
    If NC% < 1 Then Exit Sub
    
    SQLX$ = " Select NUCLIEN,SUM(IMPODEBE) AS IDEBE, SUM(IMPOHABE)AS IHABE ,FECHEMISI, DOCU_APLI,DOCU_ORIG  from factuanti with(nolock)"
    SQLX$ = SQLX$ + " where nuclien =" & NC%
    SQLX$ = SQLX$ + " AND TIPOMOVIM = 'FVEN'"
    SQLX$ = SQLX$ + " AND IMPODEBE > 0.005"
    SQLX$ = SQLX$ + " GROUP BY NUCLIEN,DOCU_ORIG,DOCU_APLI,FECHEMISI"
    Set RS = READSET(SQLX$)
    J& = 0
    With RS
        Do While Not .EOF
           NC% = XVALO(!NUCLIEN)
           RASO$ = rasocli$(NC%)
           AntiSaldo# = XVALO(!IDEBE) - XVALO(!IHABE)
           '
           J& = J& + 1
           MSF4.Rows = J& + 1
           MSF4.TextMatrix(J&, 1) = SAFETEXT$(NC%)
           MSF4.TextMatrix(J&, 2) = RASO$
           MSF4.TextMatrix(J&, 3) = SAFETEXT$(!DOCU_ORIG)
           MSF4.TextMatrix(J&, 4) = FORMATNUM$((XVALO(!IDEBE)), "F12.2")
           MSF4.TextMatrix(J&, 5) = FECHATEX$(CVINT(!Fechemisi))
          .MoveNext
        Loop
    End With
    RS.Close
    Set RS = Nothing

End Sub

Private Sub MSF1_DblClick()
   i& = MSF1.MouseRow
   J& = MSF1.MouseCol
   '
   If i& < 1 Then
    Static MODO  As Boolean
    If (MSF1.MouseRow = 0) Then
        ' Ordena en forma ascendente
        If MODO Then
            MSF1.COL = MSF1.MouseCol
            MSF1.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSF1.COL = MSF1.MouseCol
            MSF1.Sort = 1
            MODO = True
        End If
     End If
     Exit Sub
   End If

'   If (MSF1.MouseRow = 0) Then Exit Sub
'   REG& = MSF1.Row
'
'
'   'If REG& < 1 Then Exit Sub
'   MSF2.Rows = MSF2.Rows + 1
'   MSF2.TextMatrix(MSF2.Rows - 1, 1) = MSF1.TextMatrix(REG&, 1)
'   MSF2.TextMatrix(MSF2.Rows - 1, 2) = MSF1.TextMatrix(REG&, 2)
'   MSF2.TextMatrix(MSF2.Rows - 1, 3) = MSF1.TextMatrix(REG&, 3)
'   MSF2.TextMatrix(MSF2.Rows - 1, 4) = MSF1.TextMatrix(REG&, 3)
'   '
'   NUFACTU$ = MSF1.TextMatrix(REG&, 8)
'   'NETO
'   CONDI$ = "TIPOMOVIM = 'VENTA'  AND DOCU_ORIG = '" & NUFACTU$ & "'"  'TOMAR EL IMPODEBE
'   IMPOTOTAL = FORMATNUM$(XVALO(VALOBADA("FACTUANTI", "SUM(IMPODEBE - IMPOHABE)", CONDI$, "NOMESS")), "F12.2")
'
'   'IVA1
'   CONDI$ = "TIPOMOVIM = 'IVA'  AND DOCU_ORIG = '" & NUFACTU$ & "' AND VARIMOVIM LIKE '" & PORIVA(1) & "%'"
'   IMPOIVA1 = XVALO(VALOBADA("FACTUANTI", "SUM(IMPODEBE - IMPOHABE)", CONDI$, "NOMESS"))
'   BASEIMPO1 = XVALO(VALOBADA("FACTUANTI", "SUM(BASEIMPODEBE - BASEIMPOHABE)", CONDI$, "NOMESS"))
'   BASIVA(3) = FORMATNUM$(XVALO(BASIVA(3)) + BASEIMPO1, "F12.2")
'   Total(0) = FORMATNUM$(XVALO(Total(0)) + IMPOIVA1, "F12.2")
'   'IVA2
'   CONDI$ = "TIPOMOVIM = 'IVA'  AND DOCU_ORIG = '" & NUFACTU$ & "' AND VARIMOVIM LIKE '" & PORIVA(0) & "%'"
'   IMPOIVA2 = XVALO(VALOBADA("FACTUANTI", "SUM(IMPODEBE - IMPOHABE)", CONDI$, "NOMESS"))
'   BASEIMPO2 = XVALO(VALOBADA("FACTUANTI", "SUM(BASEIMPODEBE - BASEIMPOHABE)", CONDI$, "NOMESS"))
'   BASIVA(2) = FORMATNUM$(XVALO(BASIVA(2)) + BASEIMPO2, "F12.2")
'   Total(1) = FORMATNUM$(XVALO(Total(1)) + IMPOIVA2, "F12.2")
'   '
'   MSF2.TextMatrix(MSF2.Rows - 1, 5) = FORMATNUM$(IMPOIVA1, "F12.2")
'   MSF2.TextMatrix(MSF2.Rows - 1, 6) = FORMATNUM$(BASEIMPO1, "F12.2")
'   MSF2.TextMatrix(MSF2.Rows - 1, 7) = FORMATNUM$(IMPOIVA2, "F12.2")
'   MSF2.TextMatrix(MSF2.Rows - 1, 8) = FORMATNUM$(BASEIMPO2, "F12.2")
'   '
'   Text1 = FORMATNUM$(XVALO(BASIVA(3)) + XVALO(BASIVA(2)) + XVALO(Total(0)) + XVALO(Total(1)), "F12.2")
'   Call borrafila(MSF1.Row, ANTSINAP09, MSF1)
''   Call ACTUTOTALANTI

End Sub

Private Sub msf2_Click()
    If (msf2.MouseRow = 0) Then Exit Sub
    REG& = msf2.Row
    Label8 = ""
    AJUSTEF& = XVALO(msf2.TextMatrix(REG&, 7))
    If AJUSTEF& < 1 Then Exit Sub
    
    NC% = MSF4.TextMatrix(1, 1)
    NUFACTU$ = msf2.TextMatrix(REG&, 6)
    '
    CONDI$ = "nuclien = " & NC%
    CONDI$ = CONDI$ + " AND  AJUSTE_NUMERO = " & AJUSTEF&
    Label8 = VALOBADA("FACTUANTI", "OBSERVACIONES", CONDI$, "NOMESS")
End Sub

Private Sub MSF3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    Static GRAV105#, GRAV210#, NUFACTUSEL$
    '
    i& = MSF3.MouseRow
    J& = MSF3.MouseCol
    If i& = 0 Or i& > MSF3.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF3.Row = i&
    MSF3.COL = J&
    On Error GoTo 0
    '
    'VALIDA VARIABLES ESTATICAS PARA QUE SI SE VULEVE A SELECCIONAR LUEGO DE HABER VARIADO SIN GRABAR
    'MANTENGA LOS VALORES ORIGINALES Y NO LEA DE LA GRILLA LOS VALORES AJUSTADOS QUE AUN NO SE GRABAROON
    
    If NUFACTUSEL$ <> TRIM$(MSF3.TextMatrix(i&, 2)) Then
      NUFACTUSEL$ = TRIM$(MSF3.TextMatrix(i&, 2))
      GRAV105# = 0
      GRAV210# = 0
    End If
    '
    MSF_GRAV105# = XVALO(MSF3.TextMatrix(i&, 6))
    MSF_GRAV210# = XVALO(MSF3.TextMatrix(i&, 8))
    MSF_EXENTO# = XVALO(MSF3.TextMatrix(i&, 9))

    GRAV105# = XVALO(MSF3.TextMatrix(i&, 6))
    GRAV210# = XVALO(MSF3.TextMatrix(i&, 8))
    GRAVIMPOEXENTO# = XVALO(MSF3.TextMatrix(i&, 9))
    '
    Call REPLA(VDEF$, NUFACTUSEL$, 1, 18)
    Call REPLA(VDEF$, MKD$(GRAV105#), 19, 8)
    Call REPLA(VDEF$, MKD$(MSF_GRAV105#), 27, 8)
    Call REPLA(VDEF$, MKD$(GRAV210#), 35, 8)
    Call REPLA(VDEF$, MKD$(MSF_GRAV210#), 43, 8)
    Call REPLA(VDEF$, MKD$(GRAVIMPOEXENTO#), 51, 8)
    Call REPLA(VDEF$, MKD$(MSF_EXENTO#), 59, 8)
    '
20  OBX$ = OBJPLA$("EDIFACAN", VDEF$)
    If OBX$ = "" Then
       Label10 = ""
       VDEF$ = OBX$
       Exit Sub
    End If
    
    '
    If TRIM$(Mid$(OBX$, 67, 64)) = "" Then
       Call COMUNI("Debe Completar La Observación del Ajuste")
       VDEF$ = OBX$
       GoTo 20
    End If
    '
    
    OBJGRAV105# = CVD(Mid$(OBX$, 27, 8))
    OBJGRAV210# = CVD(Mid$(OBX$, 43, 8))
    OBJEXENTO# = CVD(Mid$(OBX$, 59, 8))
    Label10 = Mid$(OBX$, 67, 64)
    '

    Label33 = FORMATNUM$(GRAV105# - OBJGRAV105#, "F12.2")
    Label34 = FORMATNUM$(GRAV210# - OBJGRAV210#, "F12.2")
    Label13 = FORMATNUM$(GRAVIMPOEXENTO# - OBJEXENTO#, "F12.2")
    '
    TASA = XVALO(TASIVA$(1))
    iva1 = OBJGRAV105# * (TASA / 100)
    
    MSF3.TextMatrix(i&, 5) = FORMATNUM$(iva1, "F12.2")
    MSF3.TextMatrix(i&, 6) = FORMATNUM$(OBJGRAV105#, "F12.2")
    TASA = XVALO(TASIVA$(0))
    iva2 = OBJGRAV210# * (TASA / 100)

    MSF3.TextMatrix(i&, 7) = FORMATNUM$(iva2, "F12.2")
    
    MSF3.TextMatrix(i&, 8) = FORMATNUM$(OBJGRAV210#, "F12.2")
    MSF3.TextMatrix(i&, 9) = FORMATNUM$(OBJEXENTO#, "F12.2")
    MSF3.TextMatrix(i&, 4) = FORMATNUM$(XVALO(MSF3.TextMatrix(i&, 5)) + XVALO(MSF3.TextMatrix(i&, 6)) + XVALO(MSF3.TextMatrix(i&, 7)) + XVALO(MSF3.TextMatrix(i&, 8)) + XVALO(MSF3.TextMatrix(i&, 9)), "F12.2")
    '
  End If
End Sub

Private Sub MSF2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'  If Button = 2 Then
'    '
'    I& = MSF2.MouseRow
'    j& = MSF2.MouseCol
'    If I& = 0 Or I& > MSF2.Rows Then Exit Sub
'    '
'    On Error Resume Next
'    MSF2.Row = I&
'    MSF2.Col = j&
'    On Error Resume Next
'
'    NUFACTU$ = TRIM$(MSF2.TextMatrix(I&, 2))
'    If NUFACTU$ = "" Then
'        Exit Sub
'    End If
'    '
'    GRAV105# = XVALO(MSF2.TextMatrix(I&, 6))
'    GRAV210# = XVALO(MSF2.TextMatrix(I&, 8))
'    '
'    Call REPLA(VDEF$, NUFACTU$, 1, 18)
'    Call REPLA(VDEF$, MKD$(GRAV105#), 19, 8)
'    Call REPLA(VDEF$, MKD$(GRAV210#), 27, 8)
'    '
'20  OBX$ = OBJPLA$("ANTIFACTU", VDEF$)
'    If OBX$ = "" Then Exit Sub
'    '
'    OBJGRAV105 = CVD(Mid$(OBX$, 19, 8))
'    OBJGRAV210 = CVD(Mid$(OBX$, 27, 8))
'    '
'    If OBJGRAV105 > GRAV105# Then
'       Call COMUNI("Imposible Ajustar Importe Gravado al 10.5% a Valor Mayor")
'       VDEF$ = OBX$
'       GoTo 20
'    End If
'    '
'    If OBJGRAV210 > GRAV210# Then
'       Call COMUNI("Imposible Ajustar Importe Gravado al 21.0% a Valor Mayor")
'       VDEF$ = OBX$
'       GoTo 20
'    End If
'    '
'    TASA = XVALO(PORIVA(3))
'    iva = OBJGRAV105 * (TASA / 100)
'    '
'    MSF2.TextMatrix(I&, 5) = FORMATNUM$(iva, "F12.2")
'    MSF2.TextMatrix(I&, 6) = FORMATNUM$(OBJGRAV105, "F12.2")
'    TASA = XVALO(PORIVA(2))
'    iva = OBJGRAV210 * (TASA / 100)
'    '
'    MSF2.TextMatrix(I&, 7) = FORMATNUM$(iva, "F12.2")
'    '
'    MSF2.TextMatrix(I&, 8) = FORMATNUM$(OBJGRAV210, "F12.2")
'    MSF2.TextMatrix(I&, 4) = FORMATNUM$(XVALO(MSF2.TextMatrix(I&, 5)) + XVALO(MSF2.TextMatrix(I&, 6)) + XVALO(MSF2.TextMatrix(I&, 7)) + XVALO(MSF2.TextMatrix(I&, 8)), "F12.2")
'    '
'    Call ACTU_MSF2
'    '
'  End If
End Sub



Private Sub MSF4_DblClick()
   If (MSF4.MouseRow = 0) Then Exit Sub
   NUFACTUSEL$ = ""
   REG& = MSF4.Row
   NUFACTU$ = MSF4.TextMatrix(REG&, 3)
   NC% = XVALO(MSF4.TextMatrix(REG&, 1))
   Label33 = "": Label34 = "": Label13 = ""
   Call CARGA_FACTURA_ANTICIPO_X_CLIENTE(NC%, NUFACTU$)
   Call CARGA_FACTURA_EDIT(REG&)
   Call CARGA_APLICACIONES(NC%, NUFACTU$)
End Sub
Sub CARGA_FACTURA_EDIT(REG&)

       If (MSF4.MouseRow = 0) Then Exit Sub
'        Campos$ = "F.Emis/D8;N.Cte./A24;Saldo/F12.2"
'    Call CARGA_ENCABEZADO(ANTSINAP09.MSF, Campos$, ANTSINAP09)
'    Campos$ = Campos$ & ";Imp.Ajust/F12.2;Iva 10.50 %/F12.2;Grav. 10.5%/F12.2;Iva 21.00 %/F12.2;Grav. 21.0%/F12.2"
'    Call CARGA_ENCABEZADO(ANTSINAP09.MSF2, Campos$, ANTSINAP09)

       REG& = MSF4.Row
       MSF3.Rows = 1
       MSF3.Rows = MSF3.Rows + 1
       MSF3.TextMatrix(MSF3.Rows - 1, 1) = MSF4.TextMatrix(REG&, 5)
       MSF3.TextMatrix(MSF3.Rows - 1, 2) = MSF4.TextMatrix(REG&, 3)
       MSF3.TextMatrix(MSF3.Rows - 1, 3) = msf2.TextMatrix(msf2.Rows - 1, 3)
       MSF3.TextMatrix(MSF3.Rows - 1, 4) = msf2.TextMatrix(msf2.Rows - 1, 3)
       '
       NUFACTU$ = MSF4.TextMatrix(REG&, 3)
       'NETO
'       CONDI$ = "TIPOMOVIM = 'VENTA'  AND DOCU_ORIG = '" & NUFACTU$ & "'"  'TOMAR EL IMPODEBE
'       IMPOTOTAL = FORMATNUM$(XVALO(VALOBADA("FACTUANTI", "SUM(IMPODEBE - IMPOHABE)", CONDI$, "NOMESS")), "F12.2")
'
       'IVA1
       CONDI$ = "TIPOMOVIM = 'IVA'  AND DOCU_ORIG = '" & NUFACTU$ & "' AND VARIMOVIM LIKE '" & TASIVA$(1) & "%'"
       IMPOIVA1 = XVALO(VALOBADA("FACTUANTI", "SUM(IMPODEBE - IMPOHABE)", CONDI$, "NOMESS"))
       BASEIMPO1 = XVALO(VALOBADA("FACTUANTI", "SUM(BASEIMPODEBE - BASEIMPOHABE)", CONDI$, "NOMESS"))
'       BASIVA(3) = FORMATNUM$(XVALO(BASIVA(3)) + BASEIMPO1, "F12.2")
'       Total(0) = FORMATNUM$(XVALO(Total(0)) + IMPOIVA1, "F12.2")
       'IVA2
       CONDI$ = "TIPOMOVIM = 'IVA'  AND DOCU_ORIG = '" & NUFACTU$ & "' AND VARIMOVIM LIKE '" & TASIVA$(0) & "%'"
       IMPOIVA2 = XVALO(VALOBADA("FACTUANTI", "SUM(IMPODEBE - IMPOHABE)", CONDI$, "NOMESS"))
       BASEIMPO2 = XVALO(VALOBADA("FACTUANTI", "SUM(BASEIMPODEBE - BASEIMPOHABE)", CONDI$, "NOMESS"))
'       BASIVA(2) = FORMATNUM$(XVALO(BASIVA(2)) + BASEIMPO2, "F12.2")
'       Total(1) = FORMATNUM$(XVALO(Total(1)) + IMPOIVA2, "F12.2")
       '
       'EXENTO
       CONDI$ = "TIPOMOVIM = 'VENTA'  AND DOCU_ORIG = '" & NUFACTU$ & "' AND VARIMOVIM = 'EXENTO'"
       IMPOEXENTO = XVALO(VALOBADA("FACTUANTI", "SUM(BASEIMPODEBE - BASEIMPOHABE)", CONDI$, "NOMESS"))

       MSF3.TextMatrix(MSF3.Rows - 1, 5) = FORMATNUM$(IMPOIVA1, "F12.2")
       MSF3.TextMatrix(MSF3.Rows - 1, 6) = FORMATNUM$(BASEIMPO1, "F12.2")
       MSF3.TextMatrix(MSF3.Rows - 1, 7) = FORMATNUM$(IMPOIVA2, "F12.2")
       MSF3.TextMatrix(MSF3.Rows - 1, 8) = FORMATNUM$(BASEIMPO2, "F12.2")
       MSF3.TextMatrix(MSF3.Rows - 1, 9) = FORMATNUM$(IMPOEXENTO, "F12.2")
       '
'       Text1 = FORMATNUM$(XVALO(BASIVA(3)) + XVALO(BASIVA(2)) + XVALO(Total(0)) + XVALO(Total(1)), "F12.2")
'       Call borrafila(ANTSINAP09.MSF4.Row, ANTSINAP09, ANTSINAP09.MSF4)
'       Call ACTUTOTALANTI
    
End Sub

Private Sub MSF6_Click()
    i& = MSF6.MouseRow
    J& = MSF6.MouseCol
    If i& = 0 Or i& > MSF6.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF6.Row = i&
    MSF6.COL = J&
    On Error Resume Next

    NC% = XVALO(MSF6.TextMatrix(i&, 1))
    If NC% = 0 Then
        Exit Sub
    End If
    '
    Call BLANFORMU
    
    Call CARGA_SOLO_FACTURA_ANTICIPO_X_CLIENTE(NC%)
End Sub

Private Sub TXTBUSCAR_Change()
   If TRIM$(TXTBUSCAR) = "" Then
      MSF1.Visible = False
      MSF6.Visible = True
      Exit Sub
    End If
    '
    MSF1.Rows = 1
    MSF1.Visible = True
    MSF1.ZOrder 0
    MSF1.Cols = MSF6.Cols
    MSF1.Redraw = False
    MSF6.Visible = False
    '
    ABUS$ = UCase(TXTBUSCAR.TEXT)
    ABUSEC$ = ""
    PPXYZ% = InStr(ABUS$, "+")
    If PPXYZ% > 1 Then
       ABUSEC$ = UCase$(TRIM$(Mid$(ABUS$, PPXYZ% + 1)))
       ABUS$ = UCase$(TRIM$(Left$(ABUS$, PPXYZ% - 1)))
    End If

    '''
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN MSF1  LOS QUE MACHEAN
    For i& = 1 To MSF6.Rows - 1
      For J& = 1 To MSF6.Cols - 1
        If ABUSEC$ <> "" Then
          If InStr(UCase$(MSF6.TextMatrix(i&, J&)), ABUS$) Or InStr(UCase$(MSF6.TextMatrix(i&, J&)), ABUSEC$) Then GoTo 50
        End If
        If InStr(UCase$(MSF6.TextMatrix(i&, J&)), ABUS$) Then
50        REG& = REG& + 1
          MSF1.Rows = MSF1.Rows + 1
          For H& = 1 To MSF6.Cols - 1
            MSF1.TextMatrix(REG&, H&) = MSF6.TextMatrix(i&, H&)
          Next H&
          Exit For
        End If
      Next J&
    Next i&
    '
    MSF1.Redraw = True
    '
End Sub

